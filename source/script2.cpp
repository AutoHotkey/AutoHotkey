/*
AutoHotkey

Copyright 2003-2009 Chris Mallett (support@autohotkey.com)

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
*/

#include "stdafx.h" // pre-compiled headers
#include <olectl.h> // for OleLoadPicture()
#include <winioctl.h> // For PREVENT_MEDIA_REMOVAL and CD lock/unlock.
#include "qmath.h" // Used by Transform() [math.h incurs 2k larger code size just for ceil() & floor()]
#include "mt19937ar-cok.h" // for sorting in random order
#include "script.h"
#include "window.h" // for IF_USE_FOREGROUND_WINDOW
#include "application.h" // for MsgSleep()
#include "resources/resource.h"  // For InputBox.
#include "TextIO.h"
#include <Psapi.h> // for GetModuleBaseName.

#undef _WIN32_WINNT // v1.1.10.01: Redefine this just for these APIs, to avoid breaking some other commands on Win XP (such as Process Close).
#define _WIN32_WINNT 0x0600 // Windows Vista
#include <mmdeviceapi.h> // for SoundSet/SoundGet.
#pragma warning(push)
#pragma warning(disable:4091) // Work around a bug in the SDK used by the v140_xp toolset.
#include <endpointvolume.h> // for SoundSet/SoundGet.
#pragma warning(pop)

#define PCRE_STATIC             // For RegEx. PCRE_STATIC tells PCRE to declare its functions for normal, static
#include "lib_pcre/pcre/pcre.h" // linkage rather than as functions inside an external DLL.

#include "script_func_impl.h"



////////////////////
// Window related //
////////////////////

ResultType Line::Splash(LPTSTR aOptions, LPTSTR aSubText, LPTSTR aMainText, LPTSTR aTitle, LPTSTR aFontName
	, LPTSTR aImageFile, bool aSplashImage)
{
	int window_index = 0;  // Set the default window to operate upon (the first).
	LPTSTR options, image_filename = aImageFile;  // Set default.
	bool turn_off = false;
	bool show_it_only = false;
	int bar_pos;
	bool bar_pos_has_been_set = false;
	bool options_consist_of_bar_pos_only = false;

	if (aSplashImage)
	{
		options = aOptions;
		if (*aImageFile)
		{
			LPTSTR colon_pos = _tcschr(aImageFile, ':');
			LPTSTR image_filename_omit_leading_whitespace = omit_leading_whitespace(image_filename); // Added in v1.0.38.04 per someone's suggestion.
			if (colon_pos)
			{
				TCHAR window_number_str[32];  // Allow extra room in case leading spaces or in hex format, e.g. 0x02
				size_t length_to_copy = colon_pos - aImageFile;
				if (length_to_copy < _countof(window_number_str))
				{
					tcslcpy(window_number_str, aImageFile, length_to_copy + 1);
					if (IsPureNumeric(window_number_str, false, false, true)) // Seems best to allow float at runtime.
					{
						// Note that filenames can start with spaces, so omit_leading_whitespace() is only
						// used if the string is entirely blank:
						image_filename = colon_pos + 1;
						image_filename_omit_leading_whitespace = omit_leading_whitespace(image_filename); // Update to reflect the change above.
						if (!*image_filename_omit_leading_whitespace)
							image_filename = image_filename_omit_leading_whitespace;
						window_index = ATOI(window_number_str) - 1;
						if (window_index < 0 || window_index >= MAX_SPLASHIMAGE_WINDOWS)
							return LineError(_T("Max window number is ") MAX_SPLASHIMAGE_WINDOWS_STR _T(".")
								, FAIL, aOptions);
					}
				}
			}
			if (!_tcsicmp(image_filename_omit_leading_whitespace, _T("Off"))) // v1.0.38.04: Ignores leading whitespace per someone's suggestion.
				turn_off = true;
			else if (!_tcsicmp(image_filename_omit_leading_whitespace, _T("Show"))) // v1.0.38.04: Ignores leading whitespace per someone's suggestion.
				show_it_only = true;
		}
	}
	else // Progress Window.
	{
		if (   !(options = _tcschr(aOptions, ':'))   )
			options = aOptions;
		else
		{
			window_index = ATOI(aOptions) - 1;
			if (window_index < 0 || window_index >= MAX_PROGRESS_WINDOWS)
				return LineError(_T("Max window number is ") MAX_PROGRESS_WINDOWS_STR _T("."), FAIL, aOptions);
			++options;
		}
		options = omit_leading_whitespace(options); // Added in v1.0.38.04 per someone's suggestion.
		if (!_tcsicmp(options, _T("Off")))
            turn_off = true;
		else if (!_tcsicmp(options, _T("Show")))
			show_it_only = true;
		else
		{
			// Allow floats at runtime for flexibility (i.e. in case aOptions was in a variable reference).
			// But still use ATOI for the conversion:
			if (IsPureNumeric(options, true, false, true)) // Negatives are allowed as of v1.0.25.
			{
				bar_pos = ATOI(options);
				bar_pos_has_been_set = true;
				options_consist_of_bar_pos_only = true;
			}
			//else leave it set to the default.
		}
	}

	SplashType &splash = aSplashImage ? g_SplashImage[window_index] : g_Progress[window_index];

	// In case it's possible for the window to get destroyed by other means (WinClose?).
	// Do this only after the above options were set so that the each window's settings
	// will be remembered until such time as "Command, Off" is used:
	if (splash.hwnd && !IsWindow(splash.hwnd))
		splash.hwnd = NULL;

	if (show_it_only)
	{
		if (splash.hwnd && !IsWindowVisible(splash.hwnd))
			ShowWindow(splash.hwnd,  SW_SHOWNOACTIVATE); // See bottom of this function for comments on SW_SHOWNOACTIVATE.
		//else for simplicity, do nothing.
		return OK;
	}

	if (!turn_off && splash.hwnd && !*image_filename && (options_consist_of_bar_pos_only || !*options)) // The "modify existing window" mode is in effect.
	{
		// If there is an existing window, just update its bar position and text.
		// If not, do nothing since we don't have the original text of the window to recreate it.
		// Since this is our thread's window, it shouldn't be necessary to use SendMessageTimeout()
		// since the window cannot be hung since by definition our thread isn't hung.  Also, setting
		// a text item from non-blank to blank is not supported so that elements can be omitted from an
		// update command without changing the text that's in the window.  The script can specify %a_space%
		// to explicitly make an element blank.
		if (!aSplashImage && bar_pos_has_been_set && splash.bar_pos != bar_pos) // Avoid unnecessary redrawing.
		{
			splash.bar_pos = bar_pos;
			if (splash.hwnd_bar)
				SendMessage(splash.hwnd_bar, PBM_SETPOS, (WPARAM)bar_pos, 0);
		}
		// SendMessage() vs. SetWindowText() is used for controls so that tabs are expanded.
		// For simplicity, the hwnd_text1 control is not expanded dynamically if it is currently of
		// height zero.  The user can recreate the window if a different height is needed.
		if (*aMainText && splash.hwnd_text1)
			SendMessage(splash.hwnd_text1, WM_SETTEXT, 0, (LPARAM)(aMainText));
		if (*aSubText)
			SendMessage(splash.hwnd_text2, WM_SETTEXT, 0, (LPARAM)(aSubText));
		if (*aTitle)
			SetWindowText(splash.hwnd, aTitle); // Use the simple method for parent window titles.
		return OK;
	}

	// Otherwise, destroy any existing window first:
	if (splash.hwnd)
		DestroyWindow(splash.hwnd);
	if (splash.hfont1) // Destroy font only after destroying the window that uses it.
		DeleteObject(splash.hfont1);
	if (splash.hfont2)
		DeleteObject(splash.hfont2);
	if (splash.hbrush)
		DeleteObject(splash.hbrush);
	if (splash.pic_bmp)
	{
		if (splash.pic_type == IMAGE_BITMAP)
			DeleteObject(splash.pic_bmp);
		else
			DestroyIcon(splash.pic_icon);
	}
	ZeroMemory(&splash, sizeof(splash)); // Set the above and all other fields to zero.

	if (turn_off)
		return OK;
	// Otherwise, the window needs to be created or recreated.

	if (!*aTitle) // Provide default title.
		aTitle = (g_script.mFileName && *g_script.mFileName) ? g_script.mFileName : _T("");

	// Since there is often just one progress/splash window, and it defaults to always-on-top,
	// it seems best to default owned to be true so that it doesn't get its own task bar button:
	bool owned = true;          // Whether this window is owned by the main window.
	bool centered_main = true;  // Whether the main text is centered.
	bool centered_sub = true;   // Whether the sub text is centered.
	bool initially_hidden = false;  // Whether the window should kept hidden (for later showing by the script).
	int style = WS_DISABLED|WS_POPUP|WS_CAPTION;  // WS_CAPTION implies WS_BORDER
	int exstyle = WS_EX_TOPMOST;
	int xpos = COORD_UNSPECIFIED;
	int ypos = COORD_UNSPECIFIED;
	int range_min = 0, range_max = 0;  // For progress bars.
	int font_size1 = 0; // 0 is the flag to "use default size".
	int font_size2 = 0;
	int font_weight1 = FW_DONTCARE;  // Flag later logic to use default.
	int font_weight2 = FW_DONTCARE;  // Flag later logic to use default.
	COLORREF bar_color = CLR_DEFAULT;
	splash.color_bk = CLR_DEFAULT;
	splash.color_text = CLR_DEFAULT;
	splash.height = COORD_UNSPECIFIED;
	if (aSplashImage)
	{
		#define SPLASH_DEFAULT_WIDTH DPIScale(300)
		splash.width = COORD_UNSPECIFIED;
		splash.object_height = COORD_UNSPECIFIED;
	}
	else // Progress window.
	{
		splash.width = SPLASH_DEFAULT_WIDTH;
		splash.object_height = 20;
	}
	splash.object_width = COORD_UNSPECIFIED;  // Currently only used for SplashImage, not Progress.
	if (*aMainText || *aSubText || !aSplashImage)
	{
		splash.margin_x = 10;
		splash.margin_y = 5;
	}
	else // Displaying only a naked image, so don't use borders.
		splash.margin_x = splash.margin_y = 0;

	for (LPTSTR cp2, cp = options; *cp; ++cp)
	{
		switch(ctoupper(*cp))
		{
		case 'A':  // Non-Always-on-top.  Synonymous with A0 in early versions.
			// Decided against this enforcement.  In the enhancement mentioned below is ever done (unlikely),
			// it seems that A1 can turn always-on-top on and A0 or A by itself can turn it off:
			//if (cp[1] == '0') // The zero is required to allow for future enhancement: modify attrib. of existing window.
			exstyle &= ~WS_EX_TOPMOST;
			break;
		case 'B': // Borderless and/or Titleless
			style &= ~WS_CAPTION;
			if (cp[1] == '1')
				style |= WS_BORDER;
			else if (cp[1] == '2')
				style |= WS_DLGFRAME;
			break;
		case 'C': // Colors
			if (!cp[1]) // Avoids out-of-bounds when the loop's own ++cp is done.
				break;
			++cp; // Always increment to omit the next char from consideration by the next loop iteration.
			switch(ctoupper(*cp))
			{
			case 'B': // Bar color.
			case 'T': // Text color.
			case 'W': // Window/Background color.
			{
				TCHAR color_str[32];
				tcslcpy(color_str, cp + 1, _countof(color_str));
				LPTSTR space_pos = StrChrAny(color_str, _T(" \t"));  // space or tab
				if (space_pos)
					*space_pos = '\0';
				//else a color name can still be present if it's at the end of the string.
				COLORREF color = ColorNameToBGR(color_str);
				if (color == CLR_NONE) // A matching color name was not found, so assume it's in hex format.
				{
					if (_tcslen(color_str) > 6)
						color_str[6] = '\0';  // Shorten to exactly 6 chars, which happens if no space/tab delimiter is present.
					color = rgb_to_bgr(_tcstol(color_str, NULL, 16));
					// if color_str does not contain something hex-numeric, black (0x00) will be assumed,
					// which seems okay given how rare such a problem would be.
				}
				switch (ctoupper(*cp))
				{
				case 'B':
					bar_color = color;
					break;
				case 'T':
					splash.color_text = color;
					break;
				case 'W':
					splash.color_bk = color;
					splash.hbrush = CreateSolidBrush(color); // Used for window & control backgrounds.
					break;
				}
				// Skip over the color string to avoid interpreting hex digits or color names as option letters:
				cp += _tcslen(color_str);
				break;
			}
			default:
				centered_sub = (*cp != '0');
				centered_main = (cp[1] != '0');
			}
		case 'F':
			if (!cp[1]) // Avoids out-of-bounds when the loop's own ++cp is done.
				break;
			++cp; // Always increment to omit the next char from consideration by the next loop iteration.
			switch(ctoupper(*cp))
			{
			case 'M':
				if ((font_size1 = _ttoi(cp + 1)) < 0)
					font_size1 = 0;
				break;
			case 'S':
				if ((font_size2 = _ttoi(cp + 1)) < 0)
					font_size2 = 0;
				break;
			}
			break;
		case 'M': // Movable and (optionally) resizable.
			style &= ~WS_DISABLED;
			if (cp[1] == '1')
				style |= WS_SIZEBOX;
			if (cp[1] == '2')
				style |= WS_SIZEBOX|WS_MINIMIZEBOX|WS_MAXIMIZEBOX|WS_SYSMENU;
			break;
		case 'P': // Starting position of progress bar [v1.0.25]
			bar_pos = _ttoi(cp + 1);
			bar_pos_has_been_set = true;
			break;
		case 'R': // Range of progress bar [v1.0.25]
			if (!cp[1]) // Ignore it because we don't want cp to ever point to the NULL terminator due to the loop's increment.
				break;
			range_min = ATOI(++cp); // Increment cp to point it to range_min.
			if (cp2 = _tcschr(cp + 1, '-'))  // +1 to omit the min's minus sign, if it has one.
			{
				cp = cp2;
				if (!cp[1]) // Ignore it because we don't want cp to ever point to the NULL terminator due to the loop's increment.
					break;
				range_max = ATOI(++cp); // Increment cp to point it to range_max, which can be negative as in this example: R-100--50
			}
			break;
		case 'T': // Give it a task bar button by making it a non-owned window.
			owned = false;
			break;
		// For options such as W, X and Y:
		// Use atoi() vs. ATOI() to avoid interpreting something like 0x01B as hex when in fact
		// the B was meant to be an option letter:
		case 'W':
			if (!cp[1]) // Avoids out-of-bounds when the loop's own ++cp is done.
				break;
			++cp; // Always increment to omit the next char from consideration by the next loop iteration.
			switch(ctoupper(*cp))
			{
			case 'M':
				if ((font_weight1 = _ttoi(cp + 1)) < 0)
					font_weight1 = 0;
				break;
			case 'S':
				if ((font_weight2 = _ttoi(cp + 1)) < 0)
					font_weight2 = 0;
				break;
			default:
				splash.width = _ttoi(cp);
				if (!aSplashImage)
					splash.width = DPIScale(splash.width);
			}
			break;
		case 'H':
			if (!_tcsnicmp(cp, _T("Hide"), 4)) // Hide vs. Hidden is debatable.
			{
				initially_hidden = true;
				cp += 3; // +3 vs. +4 due to the loop's own ++cp.
			}
			else // Allow any width/height to be specified so that the window can be "rolled up" to its title bar:
			{
				splash.height = _ttoi(cp + 1);
				if (!aSplashImage)
					splash.height = DPIScale(splash.height);
			}
			break;
		case 'X':
			xpos = _ttoi(cp + 1);
			break;
		case 'Y':
			ypos = _ttoi(cp + 1);
			break;
		case 'Z':
			if (!cp[1]) // Avoids out-of-bounds when the loop's own ++cp is done.
				break;
			++cp; // Always increment to omit the next char from consideration by the next loop iteration.
			switch(ctoupper(*cp))
			{
			case 'B':  // for backward compatibility with interim releases of v1.0.14
			case 'H':
				splash.object_height = _ttoi(cp + 1); // Allow it to be zero or negative to omit the object.
				break;
			case 'W':
				if (aSplashImage)
					splash.object_width = _ttoi(cp + 1); // Allow it to be zero or negative to omit the object.
				//else for Progress, don't allow width to be changed since a zero would omit the bar.
				break;
			case 'X':
				splash.margin_x = _ttoi(cp + 1);
				break;
			case 'Y':
				splash.margin_y = _ttoi(cp + 1);
				break;
			}
			break;
		// Otherwise: Ignore other characters, such as the digits that occur after the P/X/Y option letters.
		} // switch()
	} // for()

	HDC hdc = CreateDC(_T("DISPLAY"), NULL, NULL, NULL);
	int pixels_per_point_y = GetDeviceCaps(hdc, LOGPIXELSY);

	// Get name and size of default font.
	HFONT hfont_default = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
	HFONT hfont_old = (HFONT)SelectObject(hdc, hfont_default);
	TCHAR default_font_name[65];
	GetTextFace(hdc, _countof(default_font_name) - 1, default_font_name);
	TEXTMETRIC tm;
	GetTextMetrics(hdc, &tm);
	int default_gui_font_height = tm.tmHeight;

	// If both are zero or less, reset object height/width for maintainability and sizing later.
	// However, if one of them is -1, the caller is asking for that dimension to be auto-calc'd
	// to "keep aspect ratio" with the the other specified dimension:
	if (   splash.object_height < 1 && splash.object_height != COORD_UNSPECIFIED
		&& splash.object_width < 1 && splash.object_width != COORD_UNSPECIFIED
		|| !splash.object_height || !splash.object_width   )
		splash.object_height = splash.object_width = 0;

	// If there's an image, handle it first so that automatic-width can be applied (if no width was specified)
	// for later font calculations:
	if (aSplashImage && *image_filename && splash.object_height)
	{
		for (bool use_gdi_plus = false; ; use_gdi_plus = true)
		{
			splash.pic_bmp = LoadPicture(image_filename,
				splash.object_width == COORD_UNSPECIFIED ? 0 : splash.object_width,
				splash.object_height == COORD_UNSPECIFIED ? 0 : splash.object_height,
				splash.pic_type, 0, use_gdi_plus);
			if (splash.pic_bmp || use_gdi_plus)
				break;
			// Re-attempt with GDI+. The first attempt is made without it for backward compatibility.
			// In particular, GDI+ causes some issues with WinSet TransColor on Windows XP.
		}
		if (splash.pic_bmp && (splash.object_height < 0 || splash.object_width < 0))
		{
			HBITMAP hbmp_to_measure = NULL;
			ICONINFO iconinfo;
			if (splash.pic_type == IMAGE_BITMAP)
				hbmp_to_measure = splash.pic_bmp;
			else // IMAGE_ICON
				if (GetIconInfo(splash.pic_icon, &iconinfo))
					hbmp_to_measure = iconinfo.hbmColor;
			if (hbmp_to_measure)
			{
				BITMAP bmp;
				if (GetObject(hbmp_to_measure, sizeof(BITMAP), &bmp))
				{
					if (splash.object_height == -1 && splash.object_width > 0)
					{
						// Caller wants height calculated based on the specified width (keep aspect ratio).
						if (bmp.bmWidth) // Avoid any chance of divide-by-zero.
							splash.object_height = (int)(((double)bmp.bmHeight / bmp.bmWidth) * splash.object_width + .5); // Round.
					}
					else if (splash.object_width == -1 && splash.object_height > 0)
					{
						// Caller wants width calculated based on the specified height (keep aspect ratio).
						if (bmp.bmHeight) // Avoid any chance of divide-by-zero.
							splash.object_width = (int)(((double)bmp.bmWidth / bmp.bmHeight) * splash.object_height + .5); // Round.
					}
					else
					{
						// Use actual width/height where unspecified:
						if (splash.object_height == COORD_UNSPECIFIED)
							splash.object_height = bmp.bmHeight;
						if (splash.object_width == COORD_UNSPECIFIED)
							splash.object_width = bmp.bmWidth;
					}
					if (splash.width == COORD_UNSPECIFIED)
						splash.width = splash.object_width + (2 * splash.margin_x);
				}
				if (splash.pic_type == IMAGE_ICON)
				{
					// Delete the bitmaps created by GetIconInfo above:
					DeleteObject(iconinfo.hbmColor);
					DeleteObject(iconinfo.hbmMask);
				}
			}
		}
	}

	// If width is still unspecified -- which should only happen if it's a SplashImage window with
	// no image, or there was a problem getting the image above -- set it to be the default.
	if (splash.width == COORD_UNSPECIFIED)
		splash.width = SPLASH_DEFAULT_WIDTH;
	// Similarly, object_height is set to zero if the object is not present:
	if (splash.object_height == COORD_UNSPECIFIED)
		splash.object_height = 0;

	// Lay out client area.  If height is COORD_UNSPECIFIED, use a temp value for now until
	// it can be later determined.
	RECT client_rect, draw_rect;
	SetRect(&client_rect, 0, 0, splash.width, splash.height == COORD_UNSPECIFIED ? 500 : splash.height);

	// Create fonts based on specified point sizes.  A zero value for font_size1 & 2 are correctly handled
	// by CreateFont():
	if (*aMainText)
	{
		// If a zero size is specified, it should use the default size.  But the default brought about
		// by passing a zero here is not what the system uses as a default, so instead use a font size
		// that is 25% larger than the default size (since the default size itself is used for aSubtext).
		// On a typical system, the default GUI font's point size is 8, so this will make it 10 by default.
		// Also, it appears that changing the system's font size in Control Panel -> Display -> Appearance
		// does not affect the reported default font size.  Thus, the default is probably 8/9 for most/all
		// XP systems and probably other OSes as well.
		// By specifying PROOF_QUALITY the nearest matching font size should be chosen, which should avoid
		// any scaling artifacts that might be caused if default_gui_font_height is not 8.
		if (   !(splash.hfont1 = CreateFont(font_size1 ? -MulDiv(font_size1, pixels_per_point_y, 72) : (int)(1.25 * default_gui_font_height)
			, 0, 0, 0, font_weight1 ? font_weight1 : FW_SEMIBOLD, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_PRECIS
			, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, FF_DONTCARE, *aFontName ? aFontName : default_font_name))   )
			// Call it again with default font in case above failed due to non-existent aFontName.
			// Update: I don't think this actually does any good, at least on XP, because it appears
			// that CreateFont() does not fail merely due to a non-existent typeface.  But it is kept
			// in case it ever fails for other reasons:
			splash.hfont1 = CreateFont(font_size1 ? -MulDiv(font_size1, pixels_per_point_y, 72) : (int)(1.25 * default_gui_font_height)
				, 0, 0, 0, font_weight1 ? font_weight1 : FW_SEMIBOLD, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_PRECIS
				, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, FF_DONTCARE, default_font_name);
		// To avoid showing a runtime error, fall back to the default font if other font wasn't created:
		SelectObject(hdc, splash.hfont1 ? splash.hfont1 : hfont_default);
		// Calc height of text by taking into account font size, number of lines, and space between lines:
		draw_rect = client_rect;
		draw_rect.left += splash.margin_x;
		draw_rect.right -= splash.margin_x;
		splash.text1_height = DrawText(hdc, aMainText, -1, &draw_rect, DT_CALCRECT | DT_WORDBREAK | DT_EXPANDTABS);
	}
	// else leave the above fields set to the zero defaults.

	if (font_size2 || font_weight2 || aFontName)
		if (   !(splash.hfont2 = CreateFont(-MulDiv(font_size2, pixels_per_point_y, 72), 0, 0, 0
			, font_weight2, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS
			, PROOF_QUALITY, FF_DONTCARE, *aFontName ? aFontName : default_font_name))   )
			// Call it again with default font in case above failed due to non-existent aFontName.
			// Update: I don't think this actually does any good, at least on XP, because it appears
			// that CreateFont() does not fail merely due to a non-existent typeface.  But it is kept
			// in case it ever fails for other reasons:
			if (font_size2 || font_weight2)
				splash.hfont2 = CreateFont(-MulDiv(font_size2, pixels_per_point_y, 72), 0, 0, 0
					, font_weight2, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS
					, PROOF_QUALITY, FF_DONTCARE, default_font_name);
	//else leave it NULL so that hfont_default will be used.

	// The font(s) will be deleted the next time this window is destroyed or recreated,
	// or by the g_script destructor.

	SPLASH_CALC_YPOS  // Calculate the Y position of each control in the window.

	if (splash.height == COORD_UNSPECIFIED)
	{
		// Since the window height was not specified, determine what it should be based on the height
		// of all the controls in the window:
		int subtext_height;
		if (*aSubText)
		{
			SelectObject(hdc, splash.hfont2 ? splash.hfont2 : hfont_default);
			// Calc height of text by taking into account font size, number of lines, and space between lines:
			// Reset unconditionally because the previous DrawText() sometimes alters the rect:
			draw_rect = client_rect;
			draw_rect.left += splash.margin_x;
			draw_rect.right -= splash.margin_x;
			subtext_height = DrawText(hdc, aSubText, -1, &draw_rect, DT_CALCRECT | DT_WORDBREAK);
		}
		else
			subtext_height = 0;
		// For the below: sub_y was previously calc'd to be the top of the subtext control.
		// Also, splash.margin_y is added because the text looks a little better if the window
		// doesn't end immediately beneath it:
		splash.height = subtext_height + sub_y + splash.margin_y;
		client_rect.bottom = splash.height;
	}

	SelectObject(hdc, hfont_old); // Necessary to avoid memory leak.
	if (!DeleteDC(hdc))
		return FAIL;  // Force a failure to detect bugs such as hdc still having a created handle inside.

	// Based on the client area determined above, expand the main_rect to include title bar, borders, etc.
	// If the window has a border or caption this also changes top & left *slightly* from zero.
	RECT main_rect = client_rect;
	AdjustWindowRectEx(&main_rect, style, FALSE, exstyle);
	int main_width = main_rect.right - main_rect.left;  // main.left might be slightly less than zero.
	int main_height = main_rect.bottom - main_rect.top; // main.top might be slightly less than zero.

	RECT work_rect;
	SystemParametersInfo(SPI_GETWORKAREA, 0, &work_rect, 0);  // Get desktop rect excluding task bar.
	int work_width = work_rect.right - work_rect.left;  // Note that "left" won't be zero if task bar is on left!
	int work_height = work_rect.bottom - work_rect.top;  // Note that "top" won't be zero if task bar is on top!

	// Seems best (and easier) to unconditionally restrict window size to the size of the desktop,
	// since most users would probably want that.  This can be overridden by using WinMove afterward.
	if (main_width > work_width)
		main_width = work_width;
	if (main_height > work_height)
		main_height = work_height;

	// Centering doesn't currently handle multi-monitor systems explicitly, since those calculations
	// require API functions that don't exist in Win95/NT (and thus would have to be loaded
	// dynamically to allow the program to launch).  Therefore, windows will likely wind up
	// being centered across the total dimensions of all monitors, which usually results in
	// half being on one monitor and half in the other.  This doesn't seem too terrible and
	// might even be what the user wants in some cases (i.e. for really big windows).
	// See comments above for why work_rect.left and top are added in (they aren't always zero).
	if (xpos == COORD_UNSPECIFIED)
		xpos = work_rect.left + ((work_width - main_width) / 2);  // Don't use splash.width.
	if (ypos == COORD_UNSPECIFIED)
		ypos = work_rect.top + ((work_height - main_height) / 2);  // Don't use splash.width.

	// CREATE Main Splash Window
	// It seems best to make this an unowned window for two reasons:
	// 1) It will get its own task bar icon then, which is usually desirable for cases where
	//    there are several progress/splash windows or the window is monitoring something.
	// 2) The progress/splash window won't prevent the main window from being used (owned windows
	//    prevent their owners from ever becoming active).
	// However, it seems likely that some users would want the above to be configurable,
	// so now there is an option to change this behavior.
	HWND dialog_owner = THREAD_DIALOG_OWNER;  // Resolve macro only once to reduce code size.
	if (!(splash.hwnd = CreateWindowEx(exstyle, WINDOW_CLASS_SPLASH, aTitle, style, xpos, ypos
		, main_width, main_height, owned ? (dialog_owner ? dialog_owner : g_hWnd) : NULL // v1.0.35.01: For flexibility, allow these windows to be owned by GUIs via +OwnDialogs.
		, NULL, g_hInstance, NULL)))
		return FAIL;  // No error msg since so rare.

	if ((style & WS_SYSMENU) || !owned)
	{
		// Setting the small icon puts it in the upper left corner of the dialog window.
		// Setting the big icon makes the dialog show up correctly in the Alt-Tab menu (but big seems to
		// have no effect unless the window is unowned, i.e. it has a button on the task bar).
		
		// L17: Use separate big/small icons for best results.
		LPARAM big_icon, small_icon;
		if (g_script.mCustomIcon)
		{
			big_icon = (LPARAM)g_script.mCustomIcon;
			small_icon = (LPARAM)g_script.mCustomIconSmall; // Should always be non-NULL when mCustomIcon is non-NULL.
		}
		else
		{
			big_icon = (LPARAM)g_IconLarge;
			small_icon = (LPARAM)g_IconSmall;
		}

		if (style & WS_SYSMENU)
			SendMessage(splash.hwnd, WM_SETICON, ICON_SMALL, small_icon);
		if (!owned)
			SendMessage(splash.hwnd, WM_SETICON, ICON_BIG, big_icon);
	}

	// Update client rect in case it was resized due to being too large (above) or in case the OS
	// auto-resized it for some reason.  These updated values are also used by SPLASH_CALC_CTRL_WIDTH
	// to position the static text controls so that text will be centered properly:
	GetClientRect(splash.hwnd, &client_rect);
	splash.height = client_rect.bottom;
	splash.width = client_rect.right;
	int control_width = client_rect.right - (splash.margin_x * 2);

	// CREATE Main label
	if (*aMainText)
	{
		splash.hwnd_text1 = CreateWindowEx(0, _T("static"), aMainText
			, WS_CHILD|WS_VISIBLE|SS_NOPREFIX|(centered_main ? SS_CENTER : SS_LEFT)
			, PROGRESS_MAIN_POS, splash.hwnd, NULL, g_hInstance, NULL);
		SendMessage(splash.hwnd_text1, WM_SETFONT, (WPARAM)(splash.hfont1 ? splash.hfont1 : hfont_default), MAKELPARAM(TRUE, 0));
	}

	if (!aSplashImage && splash.object_height > 0) // Progress window
	{
		// CREATE Progress control (always starts off at its default position as determined by OS/common controls):
		if (splash.hwnd_bar = CreateWindowEx(WS_EX_CLIENTEDGE, PROGRESS_CLASS, NULL, WS_CHILD|WS_VISIBLE|PBS_SMOOTH
			, PROGRESS_BAR_POS, splash.hwnd, NULL, NULL, NULL))
		{
			if (range_min || range_max) // i.e. if both are zero, leave it at the default range, which is 0-100.
			{
				if (range_min > -1 && range_min < 0x10000 && range_max > -1 && range_max < 0x10000)
					// Since the values fall within the bounds for Win95/NT to support, use the old method
					// in case Win95/NT lacks MSIE 3.0:
					SendMessage(splash.hwnd_bar, PBM_SETRANGE, 0, MAKELPARAM(range_min, range_max));
				else
					SendMessage(splash.hwnd_bar, PBM_SETRANGE32, range_min, range_max);
			}


			if (bar_color != CLR_DEFAULT)
			{
				// Remove visual styles so that specified color will be obeyed:
				MySetWindowTheme(splash.hwnd_bar, L"", L"");
				SendMessage(splash.hwnd_bar, PBM_SETBARCOLOR, 0, bar_color); // Set color.
			}
			if (splash.color_bk != CLR_DEFAULT)
				SendMessage(splash.hwnd_bar, PBM_SETBKCOLOR, 0, splash.color_bk); // Set color.
			if (bar_pos_has_been_set) // Note that the window is not yet visible at this stage.
				// This happens when the window doesn't exist and a command such as the following is given:
				// Progress, 50 [, ...].  As of v1.0.25, it also happens via the new 'P' option letter:
				SendMessage(splash.hwnd_bar, PBM_SETPOS, (WPARAM)bar_pos, 0);
			else // Ask the control its starting/default position in case a custom range is in effect.
				bar_pos = (int)SendMessage(splash.hwnd_bar, PBM_GETPOS, 0, 0);
			splash.bar_pos = bar_pos; // Save the current position to avoid redraws when future positions are identical to current.
		}
	}

	// CREATE Sub label
	if (splash.hwnd_text2 = CreateWindowEx(0, _T("static"), aSubText
		, WS_CHILD|WS_VISIBLE|SS_NOPREFIX|(centered_sub ? SS_CENTER : SS_LEFT)
		, PROGRESS_SUB_POS, splash.hwnd, NULL, g_hInstance, NULL))
		SendMessage(splash.hwnd_text2, WM_SETFONT, (WPARAM)(splash.hfont2 ? splash.hfont2 : hfont_default), MAKELPARAM(TRUE, 0));

	// Show it without activating it.  Even with options that allow the window to be activated (such
	// as movable), it seems best to do this to prevent changing the current foreground window, which
	// is usually desirable for progress/splash windows since they should be seen but not be disruptive:
	if (!initially_hidden)
		ShowWindow(splash.hwnd,  SW_SHOWNOACTIVATE);
	return OK;
}



ResultType Line::ToolTip(LPTSTR aText, LPTSTR aX, LPTSTR aY, LPTSTR aID)
{
	int window_index = *aID ? ATOI(aID) - 1 : 0;
	if (window_index < 0 || window_index >= MAX_TOOLTIPS)
		return LineError(_T("Max window number is ") MAX_TOOLTIPS_STR _T("."), FAIL, aID);
	HWND tip_hwnd = g_hWndToolTip[window_index];

	// Destroy windows except the first (for performance) so that resources/mem are conserved.
	// The first window will be hidden by the TTM_UPDATETIPTEXT message if aText is blank.
	// UPDATE: For simplicity, destroy even the first in this way, because otherwise a script
	// that turns off a non-existent first tooltip window then later turns it on will cause
	// the window to appear in an incorrect position.  Example:
	// ToolTip
	// ToolTip, text, 388, 24
	// Sleep, 1000
	// ToolTip, text, 388, 24
	if (!*aText)
	{
		if (tip_hwnd && IsWindow(tip_hwnd))
			DestroyWindow(tip_hwnd);
		g_hWndToolTip[window_index] = NULL;
		return OK;
	}

	// Use virtual desktop so that tooltip can move onto non-primary monitor in a multi-monitor system:
	RECT dtw;
	GetVirtualDesktopRect(dtw);

	bool one_or_both_coords_unspecified = !*aX || !*aY;
	POINT pt, pt_cursor;
	if (one_or_both_coords_unspecified)
	{
		// Don't call GetCursorPos() unless absolutely needed because it seems to mess
		// up double-click timing, at least on XP.  UPDATE: Is isn't GetCursorPos() that's
		// interfering with double clicks, so it seems it must be the displaying of the ToolTip
		// window itself.
		GetCursorPos(&pt_cursor);
		pt.x = pt_cursor.x + 16;  // Set default spot to be near the mouse cursor.
		pt.y = pt_cursor.y + 16;  // Use 16 to prevent the tooltip from overlapping large cursors.
		// Update: Below is no longer needed due to a better fix further down that handles multi-line tooltips.
		// 20 seems to be about the right amount to prevent it from "warping" to the top of the screen,
		// at least on XP:
		//if (pt.y > dtw.bottom - 20)
		//	pt.y = dtw.bottom - 20;
	}

	POINT origin = {0};
	if (*aX || *aY) // Need the offsets.
		CoordToScreen(origin, COORD_MODE_TOOLTIP);

	// This will also convert from relative to screen coordinates if appropriate:
	if (*aX)
		pt.x = ATOI(aX) + origin.x;
	if (*aY)
		pt.y = ATOI(aY) + origin.y;

	TOOLINFO ti = {0};
	ti.cbSize = sizeof(ti) - sizeof(void *); // Fixed for v1.0.36.05: Tooltips fail to work on Win9x and probably NT4/2000 unless the size for the *lpReserved member in _WIN32_WINNT 0x0501 is omitted.
	ti.uFlags = TTF_TRACK;
	ti.lpszText = aText;
	// Note that the ToolTip won't work if ti.hwnd is assigned the HWND from GetDesktopWindow().
	// All of ti's other members are left at NULL/0, including the following:
	//ti.hinst = NULL;
	//ti.uId = 0;
	//ti.rect.left = ti.rect.top = ti.rect.right = ti.rect.bottom = 0;

	// My: This does more harm that good (it causes the cursor to warp from the right side to the left
	// if it gets to close to the right side), so for now, I did a different fix (above) instead:
	//ti.rect.bottom = dtw.bottom;
	//ti.rect.right = dtw.right;
	//ti.rect.top = dtw.top;
	//ti.rect.left = dtw.left;

	// No need to use SendMessageTimeout() since the ToolTip() is owned by our own thread, which
	// (since we're here) we know is not hung or heavily occupied.

	// v1.0.40.12: Added the IsWindow() check below to recreate the tooltip in cases where it was destroyed
	// by external means such as Alt-F4 or WinClose.
	if (!tip_hwnd || !IsWindow(tip_hwnd))
	{
		// This this window has no owner, it won't be automatically destroyed when its owner is.
		// Thus, it will be explicitly by the program's exit function.
		tip_hwnd = g_hWndToolTip[window_index] = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, NULL, TTS_NOPREFIX | TTS_ALWAYSTIP
			, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, NULL, NULL);
		SendMessage(tip_hwnd, TTM_ADDTOOL, 0, (LPARAM)(LPTOOLINFO)&ti);
		// v1.0.21: GetSystemMetrics(SM_CXSCREEN) is used for the maximum width because even on a
		// multi-monitor system, most users would not want a tip window to stretch across multiple monitors:
		SendMessage(tip_hwnd, TTM_SETMAXTIPWIDTH, 0, (LPARAM)GetSystemMetrics(SM_CXSCREEN));
		// Must do these next two when the window is first created, otherwise GetWindowRect() below will retrieve
		// a tooltip window size that is quite a bit taller than it winds up being:
		SendMessage(tip_hwnd, TTM_TRACKPOSITION, 0, (LPARAM)MAKELONG(pt.x, pt.y));
		SendMessage(tip_hwnd, TTM_TRACKACTIVATE, TRUE, (LPARAM)&ti);
	}
	// Bugfix for v1.0.21: The below is now called unconditionally, even if the above newly created the window.
	// If this is not done, the tip window will fail to appear the first time it is invoked, at least when
	// all of the following are true:
	// 1) Windows XP;
	// 2) Common controls v6 (via manifest);
	// 3) "Control Panel >> Display >> Effects >> Use transition >> Fade effect" setting is in effect.
	SendMessage(tip_hwnd, TTM_UPDATETIPTEXT, 0, (LPARAM)&ti);

	RECT ttw = {0};
	GetWindowRect(tip_hwnd, &ttw); // Must be called this late to ensure the tooltip has been created by above.
	int tt_width = ttw.right - ttw.left;
	int tt_height = ttw.bottom - ttw.top;

	// v1.0.21: Revised for multi-monitor support.  I read somewhere that dtw.left can be negative (perhaps
	// if the secondary monitor is to the left of the primary).  So it seems best to assume it is possible:
	if (pt.x + tt_width >= dtw.right)
		pt.x = dtw.right - tt_width - 1;
	if (pt.y + tt_height >= dtw.bottom)
		pt.y = dtw.bottom - tt_height - 1;
	// It seems best not to have each of the below paired with the above.  This is because it allows
	// the flexibility to explicitly move the tooltip above or to the left of the screen.  Such a feat
	// should only be possible if done via explicitly passed-in negative coordinates for aX and/or aY.
	// In other words, it should be impossible for a tooltip window to follow the mouse cursor somewhere
	// off the virtual screen because:
	// 1) The mouse cursor is normally trapped within the bounds of the virtual screen.
	// 2) The tooltip window defaults to appearing South-East of the cursor.  It can only appear
	//    in some other quadrant if jammed against the right or bottom edges of the screen, in which
	//    case it can't be partially above or to the left of the virtual screen unless it's really
	//    huge, which seems very unlikely given that it's limited to the maximum width of the
	//    primary display as set by TTM_SETMAXTIPWIDTH above.
	//else if (pt.x < dtw.left) // Should be impossible for this to happen due to mouse being off the screen.
	//	pt.x = dtw.left;      // But could happen if user explicitly passed in a coord that was too negative.
	//...
	//else if (pt.y < dtw.top)
	//	pt.y = dtw.top;

	if (one_or_both_coords_unspecified)
	{
		// Since Tooltip is being shown at the cursor's coordinates, try to ensure that the above
		// adjustment doesn't result in the cursor being inside the tooltip's window boundaries,
		// since that tends to cause problems such as blocking the tray area (which can make a
		// tooltip script impossible to terminate).  Normally, that can only happen in this case
		// (one_or_both_coords_unspecified == true) when the cursor is near the bottom-right
		// corner of the screen (unless the mouse is moving more quickly than the script's
		// ToolTip update-frequency can cope with, but that seems inconsequential since it
		// will adjust when the cursor slows down):
		ttw.left = pt.x;
		ttw.top = pt.y;
		ttw.right = ttw.left + tt_width;
		ttw.bottom = ttw.top + tt_height;
		if (pt_cursor.x >= ttw.left && pt_cursor.x <= ttw.right && pt_cursor.y >= ttw.top && pt_cursor.y <= ttw.bottom)
		{
			// Push the tool tip to the upper-left side, since normally the only way the cursor can
			// be inside its boundaries (when one_or_both_coords_unspecified == true) is when the
			// cursor is near the bottom right corner of the screen.
			pt.x = pt_cursor.x - tt_width - 3;    // Use a small offset since it can't overlap the cursor
			pt.y = pt_cursor.y - tt_height - 3;   // when pushed to the the upper-left side of it.
		}
	}

	SendMessage(tip_hwnd, TTM_TRACKPOSITION, 0, (LPARAM)MAKELONG(pt.x, pt.y));
	// And do a TTM_TRACKACTIVATE even if the tooltip window already existed upon entry to this function,
	// so that in case it was hidden or dismissed while its HWND still exists, it will be shown again:
	SendMessage(tip_hwnd, TTM_TRACKACTIVATE, TRUE, (LPARAM)&ti);
	return OK;
}



ResultType Line::TrayTip(LPTSTR aTitle, LPTSTR aText, LPTSTR aTimeout, LPTSTR aOptions)
{
	if (!g_os.IsWin2000orLater()) // Older OSes do not support it, so do nothing.
		return OK;
	NOTIFYICONDATA nic = {0};
	nic.cbSize = sizeof(nic);
	nic.uID = AHK_NOTIFYICON;  // This must match our tray icon's uID or Shell_NotifyIcon() will return failure.
	nic.hWnd = g_hWnd;
	nic.uFlags = NIF_INFO;
	nic.uTimeout = ATOI(aTimeout) * 1000;
	nic.dwInfoFlags = ATOI(aOptions);
	tcslcpy(nic.szInfoTitle, aTitle, _countof(nic.szInfoTitle)); // Empty title omits the title line entirely.
	tcslcpy(nic.szInfo, aText, _countof(nic.szInfo));	// Empty text removes the balloon.
	Shell_NotifyIcon(NIM_MODIFY, &nic);
	return OK; // i.e. never a critical error if it fails.
}



ResultType Line::Transform(LPTSTR aCmd, LPTSTR aValue1, LPTSTR aValue2)
{
	Var &output_var = *OUTPUT_VAR;
	TransformCmds trans_cmd = ConvertTransformCmd(aCmd);
	// Since command names are validated at load-time, this only happens if the command name
	// was contained in a variable reference.  Since that is very rare, output_var is simply
	// made blank to indicate the problem:
	if (trans_cmd == TRANS_CMD_INVALID)
		return output_var.Assign();

	TCHAR buf[32];
	int value32;
	INT64 value64;
	double value_double1, value_double2, multiplier;
	double result_double;
	SymbolType value1_is_pure_numeric, value2_is_pure_numeric;

	#undef DETERMINE_NUMERIC_TYPES
	#define DETERMINE_NUMERIC_TYPES \
		value1_is_pure_numeric = IsPureNumeric(aValue1, true, false, true);\
		value2_is_pure_numeric = IsPureNumeric(aValue2, true, false, true);

	#define EITHER_IS_FLOAT (value1_is_pure_numeric == PURE_FLOAT || value2_is_pure_numeric == PURE_FLOAT)

	// If neither input is float, the result is assigned as an integer (i.e. no decimal places):
	#define ASSIGN_BASED_ON_TYPE \
		DETERMINE_NUMERIC_TYPES \
		if (EITHER_IS_FLOAT) \
			return output_var.Assign(result_double);\
		else\
			return output_var.Assign((INT64)result_double);

	// Have a negative exponent always cause a floating point result:
	#define ASSIGN_BASED_ON_TYPE_POW \
		DETERMINE_NUMERIC_TYPES \
		if (EITHER_IS_FLOAT || value_double2 < 0) \
			return output_var.Assign(result_double);\
		else\
			return output_var.Assign((INT64)result_double);

	#define ASSIGN_BASED_ON_TYPE_SINGLE \
		if (IsPureNumeric(aValue1, true, false, true) == PURE_FLOAT)\
			return output_var.Assign(result_double);\
		else\
			return output_var.Assign((INT64)result_double);

	// If rounding to an integer, ensure the result is stored as an integer:
	#define ASSIGN_BASED_ON_TYPE_SINGLE_ROUND \
		if (IsPureNumeric(aValue1, true, false, true) == PURE_FLOAT && value32 > 0)\
			return output_var.Assign(result_double);\
		else\
			return output_var.Assign((INT64)result_double);

	switch(trans_cmd)
	{
	case TRANS_CMD_ASC:
		if (*aValue1)
			return output_var.Assign(TRANS_CHAR_TO_INT(*aValue1));
		else
			return output_var.Assign();

	case TRANS_CMD_CHR:
		value32 = ATOI(aValue1);
		if (value32 < 0 || value32 > TRANS_CHAR_MAX)
			return output_var.Assign();
		else
		{
			*buf = value32;  // Store value as a single-character string.
			*(buf + 1) = '\0';
			return output_var.Assign(buf);
		}

	case TRANS_CMD_DEREF:
		return Deref(&output_var, aValue1);

#ifndef UNICODE
	case TRANS_CMD_UNICODE:
		int char_count;
		if (output_var.Type() == VAR_CLIPBOARD)
		{
			// Since the output var is the clipboard, the mode is autodetected as the following:
			// Convert aValue1 from UTF-8 to Unicode and put the result onto the clipboard.
			// MSDN: "Windows 95: Under the Microsoft Layer for Unicode, MultiByteToWideChar also
			// supports CP_UTF7 and CP_UTF8."
			if (   !(char_count = UTF8ToWideChar(aValue1, NULL, 0))   ) // Get required buffer size in WCHARs (includes terminator).
				return output_var.Assign(); // Make output_var (the clipboard in this case) blank to indicate failure.
			LPVOID clip_buf;
			if (   !(clip_buf = g_clip.PrepareForWrite(char_count * sizeof(WCHAR)))   )
				return output_var.Assign(); // Make output_var (the clipboard in this case) blank to indicate failure.
			// Perform the conversion:
			if (!UTF8ToWideChar(aValue1, (LPWSTR)clip_buf, char_count))
			{
				g_clip.AbortWrite();
				return output_var.Assign(); // Make clipboard blank to indicate failure.
			}
			return g_clip.Commit(CF_UNICODETEXT); // Save as type Unicode. It will display any error that occurs.
		}
		// Otherwise, going in the reverse direction: convert the clipboard contents to UTF-8 and put
		// the result into a normal variable.
		if (!IsClipboardFormatAvailable(CF_UNICODETEXT) || !g_clip.Open()) // Relies on short-circuit boolean order.
			return output_var.Assign(); // Make the (non-clipboard) output_var blank to indicate failure.
		if (   !(g_clip.mClipMemNow = g_clip.GetClipboardDataTimeout(CF_UNICODETEXT)) // Relies on short-circuit boolean order.
			|| !(g_clip.mClipMemNowLocked = (LPTSTR)GlobalLock(g_clip.mClipMemNow))
			|| !(char_count = WideCharToUTF8((LPCWSTR)g_clip.mClipMemNowLocked, NULL, 0))   ) // char_count includes terminator.
		{
			// Above finds out how large the contents will be when converted to UTF-8.
			// In this case, it failed to determine the count, perhaps due to Win95 lacking Unicode layer, etc.
			g_clip.Close();
			return output_var.Assign(); // Make the (non-clipboard) output_var blank to indicate failure.
		}
		// Otherwise, it found the count.  Set up the output variable, enlarging it if needed:
		if (output_var.AssignString(NULL, char_count - 1) != OK) // Don't combine this with the above or below it can return FAIL.
		{
			g_clip.Close();
			return FAIL;  // It already displayed the error.
		}
		// Perform the conversion:
		char_count = WideCharToUTF8((LPCWSTR)g_clip.mClipMemNowLocked, output_var.Contents(), char_count);
		g_clip.Close(); // Close the clipboard and free the memory.
		output_var.Close(); // Length() was already set properly by Assign() above. Currently it can't be VAR_CLIPBOARD since that would auto-detect as the reverse direction.
		if (!char_count)
			return output_var.Assign(); // Make non-clipboard output_var blank to indicate failure.
		return OK;
#endif

	case TRANS_CMD_HTML:
	{
		// These are the encoding-neutral translations for ASC 128 through 255 as shown by Dreamweaver.
		// It's possible that using just the &#number convention (e.g. &#128 through &#255;) would be
		// more appropriate for some users, but that mode can be added in the future if it is ever
		// needed (by passing a mode setting for aValue2):
		// €‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿
		// ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ
		static const LPTSTR sHtml[] = { // v1.0.40.02: Removed leading '&' and trailing ';' to reduce code size.
#ifndef UNICODE
			  _T("euro"), _T("#129"), _T("sbquo"), _T("fnof"), _T("bdquo"), _T("hellip"), _T("dagger"), _T("Dagger")
			, _T("circ"), _T("permil"), _T("Scaron"), _T("lsaquo"), _T("OElig"), _T("#141"), _T("#381"), _T("#143")
			, _T("#144"), _T("lsquo"), _T("rsquo"), _T("ldquo"), _T("rdquo"), _T("bull"), _T("ndash"), _T("mdash")
			, _T("tilde"), _T("trade"), _T("scaron"), _T("rsaquo"), _T("oelig"), _T("#157"), _T("#382"), _T("Yuml")
			,
#endif
 			  _T("nbsp"), _T("iexcl"), _T("cent"), _T("pound"), _T("curren"), _T("yen"), _T("brvbar"), _T("sect")
			, _T("uml"), _T("copy"), _T("ordf"), _T("laquo"), _T("not"), _T("shy"), _T("reg"), _T("macr")
			, _T("deg"), _T("plusmn"), _T("sup2"), _T("sup3"), _T("acute"), _T("micro"), _T("para"), _T("middot")
			, _T("cedil"), _T("sup1"), _T("ordm"), _T("raquo"), _T("frac14"), _T("frac12"), _T("frac34"), _T("iquest")
			, _T("Agrave"), _T("Aacute"), _T("Acirc"), _T("Atilde"), _T("Auml"), _T("Aring"), _T("AElig"), _T("Ccedil")
			, _T("Egrave"), _T("Eacute"), _T("Ecirc"), _T("Euml"), _T("Igrave"), _T("Iacute"), _T("Icirc"), _T("Iuml")
			, _T("ETH"), _T("Ntilde"), _T("Ograve"), _T("Oacute"), _T("Ocirc"), _T("Otilde"), _T("Ouml"), _T("times")
			, _T("Oslash"), _T("Ugrave"), _T("Uacute"), _T("Ucirc"), _T("Uuml"), _T("Yacute"), _T("THORN"), _T("szlig")
			, _T("agrave"), _T("aacute"), _T("acirc"), _T("atilde"), _T("auml"), _T("aring"), _T("aelig"), _T("ccedil")
			, _T("egrave"), _T("eacute"), _T("ecirc"), _T("euml"), _T("igrave"), _T("iacute"), _T("icirc"), _T("iuml")
			, _T("eth"), _T("ntilde"), _T("ograve"), _T("oacute"), _T("ocirc"), _T("otilde"), _T("ouml"), _T("divide")
			, _T("oslash"), _T("ugrave"), _T("uacute"), _T("ucirc"), _T("uuml"), _T("yacute"), _T("thorn"), _T("yuml")
		};
#ifdef UNICODE
		#define TRANS_HTML_NAMED		0x00000001
		#define TRANS_HTML_NUMBERED		0x00000002
		DWORD aFlags = *aValue2 ? ATOI(aValue2) : TRANS_HTML_NAMED;
#endif

		// Determine how long the result string will be so that the output variable can be expanded
		// to handle it:
		VarSizeType length;
		TBYTE *ucp;
		for (length = 0, ucp = (TBYTE *)aValue1; *ucp; ++ucp)
		{
			switch(*ucp)
			{
			case '"':  // &quot;
				length += 6;
				break;
			case '&': // &amp;
			case '\n': // <br>\n
				length += 5;
				break; // v1.0.45: Added missing break.  This had caused incorrect lengths inside some variables, which led to problems in places that relied on the accuracy of the internal lengths.
			case '<': // &lt;
			case '>': // &gt;
				length += 4;
				break; // v1.0.45: Added missing break.
			default:
#ifdef UNICODE
				if (*ucp >= 0x80) {
					if (aFlags & TRANS_HTML_NAMED) {
						switch (*ucp) {
							case 0x0178: // &Yuml;
							case 0x0192: // &fnof;
							case 0x02C6: // &circ;
							case 0x2022: // &bull;
							case 0x20AC: // &euro;
								length += 6;
								goto end_get_length;
							case 0x0152: // &OElig;
							case 0x0153: // &oelig;
							case 0x02DC: // &tilde;
							case 0x2013: // &ndash;
							case 0x2014: // &mdash;
							case 0x2018: // &lsquo;
							case 0x2019: // &rsquo;
							case 0x201C: // &ldquo;
							case 0x201D: // &rdquo;
							case 0x201E: // &bdquo;
							case 0x2122: // &trade;
								length += 7;
								goto end_get_length;
							case 0x0160: // &Scaron;
							case 0x0161: // &scaron;
							case 0x2020: // &dagger;
							case 0x2021: // &Dagger;
							case 0x2026: // &hellip;
							case 0x2030: // &permil;
							case 0x2039: // &lsaquo;
							case 0x203A: // &rsaquo;
								length += 8;
								goto end_get_length;
							default:
								if (*ucp >= 0xA0 && *ucp <= 0xFF) {
									length += (VarSizeType)_tcslen(sHtml[*ucp - 0xA0]) + 2; // +2 for the leading '&' and the trailing ';'.
									goto end_get_length;
								}
								// else handled by the following
								break;
						}
					}
					if (aFlags & TRANS_HTML_NUMBERED)
						length += ((int) qmathLog10(*ucp)) + 4; // &#NNN;
					else
						++length;
end_get_length:
					; // prevents compilation error
				}
#else
				if (*ucp >= 0x80)
					length += (VarSizeType)_tcslen(sHtml[*ucp - 0x80]) + 2; // +2 for the leading '&' and the trailing ';'.
#endif
				else
					++length;
			}
		}

		// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
		// this call will set up the clipboard for writing:
		if (output_var.AssignString(NULL, length) != OK)
			return FAIL;  // It already displayed the error.
		LPTSTR contents = output_var.Contents();  // For performance and tracking.

		// Translate the text to HTML:
		for (ucp = (TBYTE *)aValue1; *ucp; ++ucp)
		{
			#define SET_HTML_ENTITY(s) { _tcscpy(contents, _T(s)); contents += (_countof(s) - 1); }
			switch(*ucp)
			{
			case '"':  // &quot;
				SET_HTML_ENTITY("&quot;");
				break;
			case '&': // &amp;
				SET_HTML_ENTITY("&amp;");
				break;
			case '\n': // <br>\n
				SET_HTML_ENTITY("<br>\n");
				break;
			case '<': // &lt;
				SET_HTML_ENTITY("&lt;");
				break;
			case '>': // &gt;
				SET_HTML_ENTITY("&gt;");
				break;
			default:
				if (*ucp >= 0x80)
				{
#ifdef UNICODE
					if (aFlags & TRANS_HTML_NAMED) {
						switch (*ucp) {
							case 0x0152:
								SET_HTML_ENTITY("&OElig;");
								goto end_set_entity;
							case 0x0153:
								SET_HTML_ENTITY("&oelig;");
								goto end_set_entity;
							case 0x0160:
								SET_HTML_ENTITY("&Scaron;");
								goto end_set_entity;
							case 0x0161:
								SET_HTML_ENTITY("&scaron;");
								goto end_set_entity;
							case 0x0178:
								SET_HTML_ENTITY("&Yuml;");
								goto end_set_entity;
							case 0x0192:
								SET_HTML_ENTITY("&fnof;");
								goto end_set_entity;
							case 0x02C6:
								SET_HTML_ENTITY("&circ;");
								goto end_set_entity;
							case 0x02DC:
								SET_HTML_ENTITY("&tilde;");
								goto end_set_entity;
							case 0x2013:
								SET_HTML_ENTITY("&ndash;");
								goto end_set_entity;
							case 0x2014:
								SET_HTML_ENTITY("&mdash;");
								goto end_set_entity;
							case 0x2018:
								SET_HTML_ENTITY("&lsquo;");
								goto end_set_entity;
							case 0x2019:
								SET_HTML_ENTITY("&rsquo;");
								goto end_set_entity;
							case 0x201A:
								SET_HTML_ENTITY("&sbquo;");
								goto end_set_entity;
							case 0x201C:
								SET_HTML_ENTITY("&ldquo;");
								goto end_set_entity;
							case 0x201D:
								SET_HTML_ENTITY("&rdquo;");
								goto end_set_entity;
							case 0x201E:
								SET_HTML_ENTITY("&bdquo;");
								goto end_set_entity;
							case 0x2020:
								SET_HTML_ENTITY("&dagger;");
								goto end_set_entity;
							case 0x2021:
								SET_HTML_ENTITY("&Dagger;");
								goto end_set_entity;
							case 0x2022:
								SET_HTML_ENTITY("&bull;");
								goto end_set_entity;
							case 0x2026:
								SET_HTML_ENTITY("&hellip;");
								goto end_set_entity;
							case 0x2030:
								SET_HTML_ENTITY("&permil;");
								goto end_set_entity;
							case 0x2039:
								SET_HTML_ENTITY("&lsaquo;");
								goto end_set_entity;
							case 0x203A:
								SET_HTML_ENTITY("&rsaquo;");
								goto end_set_entity;
							case 0x20AC:
								SET_HTML_ENTITY("&euro;");
								goto end_set_entity;
							case 0x2122:
								SET_HTML_ENTITY("&trade;");
								goto end_set_entity;
							default:
								if (*ucp >= 0xA0 && *ucp <= 0xFF)
								{
									*contents++ = '&';
									_tcscpy(contents, sHtml[*ucp - 0xA0]);
									contents += _tcslen(contents);
									*contents++ = ';';
									goto end_set_entity;
								}
								// else handled by the following
								break;
						} // switch (*ucp)
					} // if (aFlags & TRANS_HTML_NAMED)
					if (aFlags & TRANS_HTML_NUMBERED)
						contents += _stprintf(contents, _T("&#%d;"), (int) *ucp);
					else
						*contents++ = *ucp;
end_set_entity:
					; // prevents compilation error
#else
					*contents++ = '&'; // v1.0.40.02
					_tcscpy(contents, sHtml[*ucp - 0x80]);
					contents += _tcslen(contents); // Added as a fix in v1.0.41 (broken in v1.0.40.02).
					*contents++ = ';'; // v1.0.40.02
#endif
				}
				else
					*contents++ = *ucp;
			}
			#undef SET_HTML_ENTITY
		}
		*contents = '\0';  // Terminate the string.
		return output_var.Close(); // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_CLIPBOARD.
	}

	case TRANS_CMD_MOD:
		if (   !(value_double2 = ATOF(aValue2))   ) // Divide by zero, set it to be blank to indicate the problem.
			return output_var.Assign();
		// Otherwise:
		result_double = qmathFmod(ATOF(aValue1), value_double2);
		ASSIGN_BASED_ON_TYPE

	case TRANS_CMD_POW:
	{
		// v1.0.44.11: With Laszlo's help, negative bases are now supported as long as the exponent is not fractional.
		// See SYM_POWER in script_expression.cpp for similar code and more comments.
		value_double1 = ATOF(aValue1);
		value_double2 = ATOF(aValue2);
		bool value1_was_negative = (value_double1 < 0);
		if (value_double1 == 0.0 && value_double2 < 0  // In essence, this is divide-by-zero.
			|| value1_was_negative && qmathFmod(value_double2, 1.0) != 0.0) // Negative base but exponent isn't close enough to being an integer: unsupported (to simplify code).
			return output_var.Assign();  // Return a consistent result (blank) rather than something that varies.
		// Otherwise:
		if (value1_was_negative)
			value_double1 = -value_double1; // Force a positive due to the limitations of qmathPow().
		result_double = qmathPow(value_double1, value_double2);
		if (value1_was_negative && qmathFabs(qmathFmod(value_double2, 2.0)) == 1.0) // Negative base and exactly-odd exponent (otherwise, it can only be zero or even because if not it would have returned higher above).
			result_double = -result_double;
		ASSIGN_BASED_ON_TYPE_POW
	}

	case TRANS_CMD_EXP:
		return output_var.Assign(qmathExp(ATOF(aValue1)));

	case TRANS_CMD_SQRT:
		value_double1 = ATOF(aValue1);
		if (value_double1 < 0)
			return output_var.Assign();
		return output_var.Assign(qmathSqrt(value_double1));

	case TRANS_CMD_LOG:
		value_double1 = ATOF(aValue1);
		if (value_double1 < 0)
			return output_var.Assign();
		return output_var.Assign(qmathLog10(ATOF(aValue1)));

	case TRANS_CMD_LN:
		value_double1 = ATOF(aValue1);
		if (value_double1 < 0)
			return output_var.Assign();
		return output_var.Assign(qmathLog(ATOF(aValue1)));

	case TRANS_CMD_ROUND:
		// In the future, a string conversion algorithm might be better to avoid the loss
		// of 64-bit integer precision that it currently caused by the use of doubles in
		// the calculation:
		value32 = ATOI(aValue2);
		multiplier = *aValue2 ? qmathPow(10, value32) : 1;
		value_double1 = ATOF(aValue1);
		result_double = (value_double1 >= 0.0 ? qmathFloor(value_double1 * multiplier + 0.5)
			: qmathCeil(value_double1 * multiplier - 0.5)) / multiplier;
		ASSIGN_BASED_ON_TYPE_SINGLE_ROUND

	case TRANS_CMD_CEIL:
	case TRANS_CMD_FLOOR:
		// The code here is similar to that in BIF_FloorCeil(), so maintain them together.
		result_double = ATOF(aValue1);
		result_double = (trans_cmd == TRANS_CMD_FLOOR) ? qmathFloor(result_double) : qmathCeil(result_double);
		return output_var.Assign((__int64)(result_double + (result_double > 0 ? 0.2 : -0.2))); // Fixed for v1.0.40.05: See comments in BIF_FloorCeil() for details.

	case TRANS_CMD_ABS:
	{
		// Seems better to convert as string to avoid loss of 64-bit integer precision
		// that would be caused by conversion to double.  I think this will work even
		// for negative hex numbers that are close to the 64-bit limit since they too have
		// a minus sign when generated by the script (e.g. -0x1).
		//result_double = qmathFabs(ATOF(aValue1));
		//ASSIGN_BASED_ON_TYPE_SINGLE
		LPTSTR cp = omit_leading_whitespace(aValue1); // i.e. caller doesn't have to have ltrimmed it.
		if (*cp == '-')
			return output_var.Assign(cp + 1);  // Omit the first minus sign (simple conversion only).
		// Otherwise, no minus sign, so just omit the leading whitespace for consistency:
		return output_var.Assign(cp);
	}

	case TRANS_CMD_SIN:
		return output_var.Assign(qmathSin(ATOF(aValue1)));

	case TRANS_CMD_COS:
		return output_var.Assign(qmathCos(ATOF(aValue1)));

	case TRANS_CMD_TAN:
		return output_var.Assign(qmathTan(ATOF(aValue1)));

	case TRANS_CMD_ASIN:
		value_double1 = ATOF(aValue1);
		if (value_double1 > 1 || value_double1 < -1)
			return output_var.Assign(); // ASin and ACos aren't defined for other values.
		return output_var.Assign(qmathAsin(ATOF(aValue1)));

	case TRANS_CMD_ACOS:
		value_double1 = ATOF(aValue1);
		if (value_double1 > 1 || value_double1 < -1)
			return output_var.Assign(); // ASin and ACos aren't defined for other values.
		return output_var.Assign(qmathAcos(ATOF(aValue1)));

	case TRANS_CMD_ATAN:
		return output_var.Assign(qmathAtan(ATOF(aValue1)));

	// For all of the below bitwise operations:
	// Seems better to convert to signed rather than unsigned so that signed values can
	// be supported.  i.e. it seems better to trade one bit in capacity in order to support
	// negative numbers.  Another reason is that commands such as IfEquals use ATOI64 (signed),
	// so if we were to produce unsigned 64 bit values here, they would be somewhat incompatible
	// with other script operations.
	case TRANS_CMD_BITAND:
		return output_var.Assign(ATOI64(aValue1) & ATOI64(aValue2));

	case TRANS_CMD_BITOR:
		return output_var.Assign(ATOI64(aValue1) | ATOI64(aValue2));

	case TRANS_CMD_BITXOR:
		return output_var.Assign(ATOI64(aValue1) ^ ATOI64(aValue2));

	case TRANS_CMD_BITNOT:
		value64 = ATOI64(aValue1);
		if (value64 < 0 || value64 > UINT_MAX)
			// Treat it as a 64-bit signed value, since no other aspects of the program
			// (e.g. IfEqual) will recognize an unsigned 64 bit number.
			return output_var.Assign(~value64);
		else
			// Treat it as a 32-bit unsigned value when inverting and assigning.  This is
			// because assigning it as a signed value would "convert" it into a 64-bit
			// value, which in turn is caused by the fact that the script sees all negative
			// numbers as 64-bit values (e.g. -1 is 0xFFFFFFFFFFFFFFFF).
			return output_var.Assign(~(DWORD)value64);

	case TRANS_CMD_BITSHIFTLEFT:  // Equivalent to multiplying by 2^value2
		return output_var.Assign(ATOI64(aValue1) << ATOI(aValue2));

	case TRANS_CMD_BITSHIFTRIGHT:  // Equivalent to dividing (integer) by 2^value2
		return output_var.Assign(ATOI64(aValue1) >> ATOI(aValue2));
	}

	return FAIL;  // Never executed (increases maintainability and avoids compiler warning).
}



ResultType Line::Input()
// OVERVIEW:
// Although a script can have many concurrent quasi-threads, there can only be one input
// at a time.  Thus, if an input is ongoing and a new thread starts, and it begins its
// own input, that input should terminate the prior input prior to beginning the new one.
// In a "worst case" scenario, each interrupted quasi-thread could have its own
// input, which is in turn terminated by the thread that interrupts it.  Every time
// this function returns, it must be sure to set g_input.status to INPUT_OFF beforehand.
// This signals the quasi-threads beneath, when they finally return, that their input
// was terminated due to a new input that took precedence.
{
	if (g_os.IsWin9x()) // v1.0.44.14: For simplicity, do nothing on Win9x rather than try to see if it actually supports the hook (such as if its some kind of emulated/hybrid OS).
		return OK; // Could also set ErrorLevel to "Timeout" and output_var to be blank, but the benefits to backward compatibility seemed too dubious.

	// Since other script threads can interrupt this command while it's running, it's important that
	// this command not refer to sArgDeref[] and sArgVar[] anytime after an interruption becomes possible.
	// This is because an interrupting thread usually changes the values to something inappropriate for this thread.
	Var *output_var = OUTPUT_VAR; // See comment above.
	if (!output_var)
	{
		// No output variable, which due to load-time validation means there are no other args either.
		// This means that the user is specifically canceling the prior input (if any).  Thus, our
		// ErrorLevel here is set to 1 or 0, but the prior input's ErrorLevel will be set to "NewInput"
		// when its quasi-thread is resumed:
		bool prior_input_is_being_terminated = (g_input.status == INPUT_IN_PROGRESS);
		g_input.status = INPUT_OFF;
		return SetErrorLevelOrThrowBool(!prior_input_is_being_terminated);
		// Above: It's considered an "error" of sorts when there is no prior input to terminate.
	}

	// Below are done directly this way rather than passed in as args mainly to emphasize that
	// ArgLength() can safely be called in Line methods like this one (which is done further below).
	// It also may also slightly improve performance and reduce code size.
	LPTSTR aOptions = ARG2, aEndKeys = ARG3, aMatchList = ARG4;
	// The aEndKeys string must be modifiable (not constant), since for performance reasons,
	// it's allowed to be temporarily altered by this function.

	// Set default in case of early return (we want these to be in effect even if
	// FAIL is returned for our thread, since the underlying thread that had the
	// active input prior to us didn't fail and it it needs to know how its input
	// was terminated):
	g_input.status = INPUT_OFF;
	
	//////////////////////////////////////////
	// Set default options and parse aOptions:
	//////////////////////////////////////////
	g_input.BackspaceIsUndo = true;
	g_input.CaseSensitive = false;
	g_input.IgnoreAHKInput = false;
	g_input.TranscribeModifiedKeys = false;
	g_input.Visible = false;
	g_input.FindAnywhere = false;
	g_input.BufferLengthMax = INPUT_BUFFER_SIZE - 1;
	int timeout = 0;
	bool endchar_mode = false;
	for (LPTSTR cp = aOptions; *cp; ++cp)
	{
		switch(ctoupper(*cp))
		{
		case 'B':
			g_input.BackspaceIsUndo = false;
			break;
		case 'C':
			g_input.CaseSensitive = true;
			break;
		case 'I':
			g_input.IgnoreAHKInput = true;
			break;
		case 'M':
			g_input.TranscribeModifiedKeys = true;
			break;
		case 'L':
			// Use atoi() vs. ATOI() to avoid interpreting something like 0x01C as hex
			// when in fact the C was meant to be an option letter:
			g_input.BufferLengthMax = _ttoi(cp + 1);
			if (g_input.BufferLengthMax > INPUT_BUFFER_SIZE - 1)
				g_input.BufferLengthMax = INPUT_BUFFER_SIZE - 1;
			break;
		case 'T':
			// Although ATOF() supports hex, it's been documented in the help file that hex should
			// not be used (see comment above) so if someone does it anyway, some option letters
			// might be misinterpreted:
			timeout = (int)(ATOF(cp + 1) * 1000);
			break;
		case 'V':
			g_input.Visible = true;
			break;
		case '*':
			g_input.FindAnywhere = true;
			break;
		case 'E':
			// Interpret single-character keys as characters rather than converting them to VK codes.
			// This tends to work better when using multiple keyboard layouts, but changes behaviour:
			// for instance, an end char of "." cannot be triggered while holding Alt.
			endchar_mode = true;
			break;
		}
	}

	//////////////////////////////////////////////
	// Set up sparse arrays according to aEndKeys:
	//////////////////////////////////////////////
	UCHAR end_vk[VK_ARRAY_COUNT] = {0};  // A sparse array that indicates which VKs terminate the input.
	UCHAR end_sc[SC_ARRAY_COUNT] = {0};  // A sparse array that indicates which SCs terminate the input.

	vk_type vk;
	sc_type sc = 0;
	modLR_type modifiersLR;
	size_t key_text_length, single_char_count = 0;
	TCHAR *end_pos, single_char_string[2];
	single_char_string[1] = '\0'; // Init its second character once, since the loop only changes the first char.

	for (TCHAR *end_key = aEndKeys; *end_key; ++end_key) // This a modified version of the processing loop used in SendKeys().
	{
		vk = 0; // Set default.  Not strictly necessary but more maintainable.
		*single_char_string = '\0';  // Set default as "this key name is not a single-char string".

		switch (*end_key)
		{
		case '}': continue;  // Important that these be ignored.
		case '{':
		{
			if (   !(end_pos = _tcschr(end_key + 1, '}'))   )
				continue;  // Do nothing, just ignore the unclosed '{' and continue.
			if (   !(key_text_length = end_pos - end_key - 1)   )
			{
				if (end_pos[1] == '}') // The string "{}}" has been encountered, which is interpreted as a single "}".
				{
					++end_pos;
					key_text_length = 1;
				}
				else // Empty braces {} were encountered.
					continue;  // do nothing: let it proceed to the }, which will then be ignored.
			}
			if (key_text_length == 1) // A single-char key name, such as {.} or {{}.
			{
				if (endchar_mode) // Handle this single-char key name by char code, not by VK.
				{
					// Although it might be sometimes useful to treat "x" as a character and "{x}" as a key,
					// "{{}" and "{}}" can't be included without the extra braces.  {vkNN} can still be used
					// to handle the key by VK instead of by character.
					single_char_count++;
					continue; // It will be processed by another section.
				}
				*single_char_string = end_key[1]; // Only used when vk != 0.
			}

			*end_pos = '\0';  // temporarily terminate the string here.

			modifiersLR = 0;  // Init prior to below.
			if (  !(vk = TextToVK(end_key + 1, &modifiersLR, true))  )
				// No virtual key, so try to find a scan code.
				if (sc = TextToSC(end_key + 1))
					end_sc[sc] = END_KEY_ENABLED;

			*end_pos = '}';  // undo the temporary termination

			end_key = end_pos;  // In prep for ++end_key at the top of the loop.
			break; // Break out of the switch() and do the vk handling beneath it (if there is a vk).
		}

		default:
			if (endchar_mode)
			{
				single_char_count++;
				continue; // It will be processed by another section.
			}
			*single_char_string = *end_key;
			modifiersLR = 0;  // Init prior to below.
			vk = TextToVK(single_char_string, &modifiersLR, true);
		} // switch()

		if (vk) // A valid virtual key code was discovered above.
		{
			end_vk[vk] |= END_KEY_ENABLED; // Use of |= is essential for cases such as ";:".
			// Insist the shift key be down to form genuinely different symbols --
			// namely punctuation marks -- but not for alphabetic chars.
			if (*single_char_string && !IsCharAlpha(*single_char_string)) // v1.0.46.05: Added check for "*single_char_string" so that non-single-char strings like {F9} work as end keys even when the Shift key is being held down (this fixes the behavior to be like it was in pre-v1.0.45).
			{
				// Now we know it's not alphabetic, and it's not a key whose name
				// is longer than one char such as a function key or numpad number.
				// That leaves mostly just the number keys (top row) and all
				// punctuation chars, which are the ones that we want to be
				// distinguished between shifted and unshifted:
				if (modifiersLR & (MOD_LSHIFT | MOD_RSHIFT))
					end_vk[vk] |= END_KEY_WITH_SHIFT;
				else
					end_vk[vk] |= END_KEY_WITHOUT_SHIFT;
			}
		}
	} // for()

	g_input.EndChars = _T("");
	if (single_char_count)
	{
		// See single_char_count++ above for comments.
		g_input.EndChars = talloca(single_char_count + 1);
		TCHAR *dst, *src;
		for (dst = g_input.EndChars, src = aEndKeys; *src; ++src)
		{
			switch (*src)
			{
			case '{':
				if (end_pos = _tcschr(src + 1, '}'))
				{
					if (end_pos == src + 1 && end_pos[1] == '}') // {}}
						end_pos++;
					if (end_pos == src + 2)
						*dst++ = src[1]; // Copy the single character from between the braces.
					src = end_pos; // Skip '{key'.  Loop does ++src to skip the '}'.
				}
				// Otherwise, just ignore the '{'.
			case '}':
				continue;
			}
			*dst++ = *src;
		}
		*dst = '\0';
	}

	/////////////////////////////////////////////////
	// Parse aMatchList into an array of key phrases:
	/////////////////////////////////////////////////
	LPTSTR *realloc_temp;  // Needed since realloc returns NULL on failure but leaves original block allocated.
	g_input.MatchCount = 0;  // Set default.
	if (*aMatchList)
	{
		// If needed, create the array of pointers that points into MatchBuf to each match phrase:
		if (!g_input.match)
		{
			if (   !(g_input.match = (LPTSTR *)malloc(INPUT_ARRAY_BLOCK_SIZE * sizeof(LPTSTR)))   )
				return LineError(ERR_OUTOFMEM);  // Short msg. since so rare.
			g_input.MatchCountMax = INPUT_ARRAY_BLOCK_SIZE;
		}
		// If needed, create or enlarge the buffer that contains all the match phrases:
		size_t aMatchList_length = ArgLength(4); // Performs better than _tcslen(aMatchList);
		size_t space_needed = aMatchList_length + 1;  // +1 for the final zero terminator.
		if (space_needed > g_input.MatchBufSize)
		{
			g_input.MatchBufSize = (UINT)(space_needed > 4096 ? space_needed : 4096);
			if (g_input.MatchBuf) // free the old one since it's too small.
				free(g_input.MatchBuf);
			if (   !(g_input.MatchBuf = tmalloc(g_input.MatchBufSize))   )
			{
				g_input.MatchBufSize = 0;
				return LineError(ERR_OUTOFMEM);  // Short msg. since so rare.
			}
		}
		// Copy aMatchList into the match buffer:
		LPTSTR source, dest;
		for (source = aMatchList, dest = g_input.match[g_input.MatchCount] = g_input.MatchBuf
			; *source; ++source)
		{
			if (*source != ',') // Not a comma, so just copy it over.
			{
				*dest++ = *source;
				continue;
			}
			// Otherwise: it's a comma, which becomes the terminator of the previous key phrase unless
			// it's a double comma, in which case it's considered to be part of the previous phrase
			// rather than the next.
			if (*(source + 1) == ',') // double comma
			{
				*dest++ = *source;
				++source;  // Omit the second comma of the pair, i.e. each pair becomes a single literal comma.
				continue;
			}
			// Otherwise, this is a delimiting comma.
			*dest = '\0';
			// If the previous item is blank -- which I think can only happen now if the MatchList
			// begins with an orphaned comma (since two adjacent commas resolve to one literal comma)
			// -- don't add it to the match list:
			if (*g_input.match[g_input.MatchCount])
			{
				++g_input.MatchCount;
				g_input.match[g_input.MatchCount] = ++dest;
				*dest = '\0';  // Init to prevent crash on orphaned comma such as "btw,otoh,"
			}
			if (*(source + 1)) // There is a next element.
			{
				if (g_input.MatchCount >= g_input.MatchCountMax) // Rarely needed, so just realloc() to expand.
				{
					// Expand the array by one block:
					if (   !(realloc_temp = (LPTSTR *)realloc(g_input.match  // Must use a temp variable.
						, (g_input.MatchCountMax + INPUT_ARRAY_BLOCK_SIZE) * sizeof(LPTSTR)))   )
						return LineError(ERR_OUTOFMEM);  // Short msg. since so rare.
					g_input.match = realloc_temp;
					g_input.MatchCountMax += INPUT_ARRAY_BLOCK_SIZE;
				}
			}
		} // for()
		*dest = '\0';  // Terminate the last item.
		// This check is necessary for only a single isolated case: When the match list
		// consists of nothing except a single comma.  See above comment for details:
		if (*g_input.match[g_input.MatchCount]) // i.e. omit empty strings from the match list.
			++g_input.MatchCount;
	}

	// Notes about the below macro:
	// In case the Input timer has already put a WM_TIMER msg in our queue before we killed it,
	// clean out the queue now to avoid any chance that such a WM_TIMER message will take effect
	// later when it would be unexpected and might interfere with this input.  To avoid an
	// unnecessary call to PeekMessage(), which has been known to yield our timeslice to other
	// processes if the CPU is under load (which might be undesirable if this input is
	// time-critical, such as in a game), call GetQueueStatus() to see if there are any timer
	// messages in the queue.  I believe that GetQueueStatus(), unlike PeekMessage(), does not
	// have the nasty/undocumented side-effect of yielding our timeslice under certain hard-to-reproduce
	// circumstances, but Google and MSDN are completely devoid of any confirming info on this.
	#define KILL_AND_PURGE_INPUT_TIMER \
	if (g_InputTimerExists)\
	{\
		KILL_INPUT_TIMER \
		if (HIWORD(GetQueueStatus(QS_TIMER)) & QS_TIMER)\
			MsgSleep(-1);\
	}

	// Be sure to get rid of the timer if it exists due to a prior, ongoing input.
	// It seems best to do this only after signaling the hook to start the input
	// so that it's MsgSleep(-1), if it launches a new hotkey or timed subroutine,
	// will be less likely to interrupt us during our setup of the input, i.e.
	// it seems best that we put the input in progress prior to allowing any
	// interruption.  UPDATE: Must do this before changing to INPUT_IN_PROGRESS
	// because otherwise the purging of the timer message might call InputTimeout(),
	// which in turn would set the status immediately to INPUT_TIMED_OUT:
	KILL_AND_PURGE_INPUT_TIMER

	//////////////////////////////////////////////////////////////
	// Initialize buffers and state variables for use by the hook:
	//////////////////////////////////////////////////////////////
	TCHAR input_buf[INPUT_BUFFER_SIZE]; // Will contain the actual input from the user.
	*input_buf = '\0';
	g_input.buffer = input_buf;
	g_input.BufferLength = 0;
	// g_input.BufferLengthMax was set in the option parsing section.

	// Point the global addresses to our memory areas on the stack:
	g_input.EndVK = end_vk;
	g_input.EndSC = end_sc;
	g_input.status = INPUT_IN_PROGRESS; // Signal the hook to start the input.

	// Make script persistent.  This is mostly for backward compatibility because it is documented behavior.
	// even though as of v1.0.42.03, the keyboard hook does not become permanent (which allows a subsequent
	// use of the commands Suspend/Hotkey to deinstall it, which seems to add flexibility/benefit).
	g_persistent = true;
	Hotkey::InstallKeybdHook(); // Install the hook (if needed).

	// A timer is used rather than monitoring the elapsed time here directly because
	// this script's quasi-thread might be interrupted by a Timer or Hotkey subroutine,
	// which (if it takes a long time) would result in our Input not obeying its timeout.
	// By using an actual timer, the TimerProc() will run when the timer expires regardless
	// of which quasi-thread is active, and it will end our input on schedule:
	if (timeout > 0)
		SET_INPUT_TIMER(timeout < 10 ? 10 : timeout)

	//////////////////////////////////////////////////////////////////
	// Wait for one of the following to terminate our input:
	// 1) The hook (due a match in aEndKeys or aMatchList);
	// 2) A thread that interrupts us with a new Input of its own;
	// 3) The timer we put in effect for our timeout (if we have one).
	//////////////////////////////////////////////////////////////////
	for (;;)
	{
		// Rather than monitoring the timeout here, just wait for the incoming WM_TIMER message
		// to take effect as a TimerProc() call during the MsgSleep():
		MsgSleep();
		if (g_input.status != INPUT_IN_PROGRESS)
			break;
	}

	switch(g_input.status)
	{
	case INPUT_TIMED_OUT:
		g_ErrorLevel->Assign(_T("Timeout"));
		break;
	case INPUT_TERMINATED_BY_MATCH:
		g_ErrorLevel->Assign(_T("Match"));
		break;
	case INPUT_TERMINATED_BY_ENDKEY:
	{
		TCHAR key_name[128] = _T("EndKey:");
		if (g_input.EndingChar)
		{
			key_name[7] = g_input.EndingChar;
			key_name[8] = '\0';
		}
		else if (g_input.EndingRequiredShift)
		{
			// Since the only way a shift key can be required in our case is if it's a key whose name
			// is a single char (such as a shifted punctuation mark), use a diff. method to look up the
			// key name based on fact that the shift key was down to terminate the input.  We also know
			// that the key is an EndingVK because there's no way for the shift key to have been
			// required by a scan code based on the logic (above) that builds the end_key arrays.
			// MSDN: "Typically, ToAscii performs the translation based on the virtual-key code.
			// In some cases, however, bit 15 of the uScanCode parameter may be used to distinguish
			// between a key press and a key release. The scan code is used for translating ALT+
			// number key combinations.
			BYTE state[256] = {0};
			state[VK_SHIFT] |= 0x80; // Indicate that the neutral shift key is down for conversion purposes.
			Get_active_window_keybd_layout // Defines the variable active_window_keybd_layout for use below.
			int count = ToUnicodeOrAsciiEx(g_input.EndingVK, vk_to_sc(g_input.EndingVK), (PBYTE)&state // Nothing is done about ToAsciiEx's dead key side-effects here because it seems to rare to be worth it (assuming its even a problem).
				, key_name + 7, g_MenuIsVisible ? 1 : 0, active_window_keybd_layout); // v1.0.44.03: Changed to call ToAsciiEx() so that active window's layout can be specified (see hook.cpp for details).
			*(key_name + 7 + count) = '\0';  // Terminate the string.
		}
		else
		{
			g_input.EndedBySC ? SCtoKeyName(g_input.EndingSC, key_name + 7, _countof(key_name) - 7)
				: VKtoKeyName(g_input.EndingVK, key_name + 7, _countof(key_name) - 7);
			// For partial backward-compatibility, keys A-Z are upper-cased when handled by VK,
			// but only if they actually correspond to those characters.  If this wasn't done,
			// the character would always be lowercase since the shift state is not considered.
			if (key_name[7] >= 'a' && key_name[7] <= 'z' && !key_name[8])
				key_name[7] -= 32;
		}
		g_ErrorLevel->Assign(key_name);
		break;
	}
	case INPUT_LIMIT_REACHED:
		g_ErrorLevel->Assign(_T("Max"));
		break;
	default: // Our input was terminated due to a new input in a quasi-thread that interrupted ours.
		g_ErrorLevel->Assign(_T("NewInput"));
		break;
	}

	g_input.status = INPUT_OFF;  // See OVERVIEW above for why this must be set prior to returning.

	// In case it ended for reason other than a timeout, in which case the timer is still on:
	KILL_AND_PURGE_INPUT_TIMER

	// Seems ok to assign after the kill/purge above since input_buf is our own stack variable
	// and its contents shouldn't be affected even if KILL_AND_PURGE_INPUT_TIMER's MsgSleep()
	// results in a new thread being created that starts a new Input:
	return output_var->Assign(input_buf);
}



ResultType Line::PerformShowWindow(ActionTypeType aActionType, LPTSTR aTitle, LPTSTR aText
	, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	// By design, the WinShow command must always unhide a hidden window, even if the user has
	// specified that hidden windows should not be detected.  So set this now so that
	// DetermineTargetWindow() will make its calls in the right mode:
	bool need_restore = (aActionType == ACT_WINSHOW && !g->DetectHiddenWindows);
	if (need_restore)
		g->DetectHiddenWindows = true;
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (need_restore)
		g->DetectHiddenWindows = false;
	if (!target_window)
		return OK;

	// WinGroup's EnumParentActUponAll() is quite similar to the following, so the two should be
	// maintained together.

	int nCmdShow = SW_NONE; // Set default.

	switch (aActionType)
	{
	// SW_FORCEMINIMIZE: supported only in Windows 2000/XP and beyond: "Minimizes a window,
	// even if the thread that owns the window is hung. This flag should only be used when
	// minimizing windows from a different thread."
	// My: It seems best to use SW_FORCEMINIMIZE on OS's that support it because I have
	// observed ShowWindow() to hang (thus locking up our app's main thread) if the target
	// window is hung.
	// UPDATE: For now, not using "force" every time because it has undesirable side-effects such
	// as the window not being restored to its maximized state after it was minimized
	// this way.
	case ACT_WINMINIMIZE:
		if (IsWindowHung(target_window))
		{
			if (g_os.IsWin2000orLater())
				nCmdShow = SW_FORCEMINIMIZE;
			//else it's not Win2k or later.  I have confirmed that SW_MINIMIZE can
			// lock up our thread on WinXP, which is why we revert to SW_FORCEMINIMIZE above.
			// Older/obsolete comment for background: don't attempt to minimize hung windows because that
			// might hang our thread because the call to ShowWindow() would never return.
		}
		else
			nCmdShow = SW_MINIMIZE;
		break;
	case ACT_WINMAXIMIZE: if (!IsWindowHung(target_window)) nCmdShow = SW_MAXIMIZE; break;
	case ACT_WINRESTORE:  if (!IsWindowHung(target_window)) nCmdShow = SW_RESTORE;  break;
	// Seems safe to assume it's not hung in these cases, since I'm inclined to believe
	// (untested) that hiding and showing a hung window won't lock up our thread, and
	// there's a chance they may be effective even against hung windows, unlike the
	// others above (except ACT_WINMINIMIZE, which has a special FORCE method):
	case ACT_WINHIDE: nCmdShow = SW_HIDE; break;
	case ACT_WINSHOW: nCmdShow = SW_SHOW; break;
	}

	// UPDATE:  Trying ShowWindowAsync()
	// now, which should avoid the problems with hanging.  UPDATE #2: Went back to
	// not using Async() because sometimes the script lines that come after the one
	// that is doing this action here rely on this action having been completed
	// (e.g. a window being maximized prior to clicking somewhere inside it).
	if (nCmdShow != SW_NONE)
	{
		// I'm not certain that SW_FORCEMINIMIZE works with ShowWindowAsync(), but
		// it probably does since there's absolutely no mention to the contrary
		// anywhere on MS's site or on the web.  But clearly, if it does work, it
		// does so only because Async() doesn't really post the message to the thread's
		// queue, instead opting for more aggressive measures.  Thus, it seems best
		// to do it this way to have maximum confidence in it:
		//if (nCmdShow == SW_FORCEMINIMIZE) // Safer not to use ShowWindowAsync() in this case.
			ShowWindow(target_window, nCmdShow);
		//else
		//	ShowWindowAsync(target_window, nCmdShow);
//PostMessage(target_window, WM_SYSCOMMAND, SC_MINIMIZE, 0);
		DoWinDelay;
	}
	return OK;  // Return success for all the above cases.
}



ResultType Line::PerformWait()
// Since other script threads can interrupt these commands while they're running, it's important that
// these commands not refer to sArgDeref[] and sArgVar[] anytime after an interruption becomes possible.
// This is because an interrupting thread usually changes the values to something inappropriate for this thread.
// fincs: it seems best that this function not throw an exception if the wait timeouts.
{
	bool wait_indefinitely;
	int sleep_duration;
	DWORD start_time;

	vk_type vk; // For GetKeyState.
	HANDLE running_process; // For RUNWAIT
	DWORD exit_code; // For RUNWAIT

	// For ACT_KEYWAIT:
	bool wait_for_keydown;
	KeyStateTypes key_state_type;
	JoyControls joy;
	int joystick_id;
	ExprTokenType token;
	TCHAR buf[LINE_SIZE];

	if (mActionType == ACT_RUNWAIT)
	{
		bool use_el = tcscasestr(ARG3, _T("UseErrorLevel"));
		if (!g_script.ActionExec(ARG1, NULL, ARG2, !use_el, ARG3, &running_process, use_el, true, ARGVAR4)) // Load-time validation has ensured that the arg is a valid output variable (e.g. not a built-in var).
			return use_el ? g_ErrorLevel->Assign(_T("ERROR")) : FAIL;
		//else fall through to the waiting-phase of the operation.
		// Above: The special string ERROR is used, rather than a number like 1, because currently
		// RunWait might in the future be able to return any value, including 259 (STATUS_PENDING).
	}
	
	// Must NOT use ELSE-IF in line below due to ELSE further down needing to execute for RunWait.
	if (mActionType == ACT_KEYWAIT)
	{
		if (   !(vk = TextToVK(ARG1))   )
		{
			if (   !(joy = (JoyControls)ConvertJoy(ARG1, &joystick_id))   ) // Not a valid key name.
				// Indicate immediate timeout (if timeout was specified) or error.
				return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
			if (!IS_JOYSTICK_BUTTON(joy)) // Currently, only buttons are supported.
				return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		}
		// Set defaults:
		wait_for_keydown = false;  // The default is to wait for the key to be released.
		key_state_type = KEYSTATE_PHYSICAL;  // Since physical is more often used.
		wait_indefinitely = true;
		sleep_duration = 0;
		for (LPTSTR cp = ARG2; *cp; ++cp)
		{
			switch(ctoupper(*cp))
			{
			case 'D':
				wait_for_keydown = true;
				break;
			case 'L':
				key_state_type = KEYSTATE_LOGICAL;
				break;
			case 'T':
				// Although ATOF() supports hex, it's been documented in the help file that hex should
				// not be used (see comment above) so if someone does it anyway, some option letters
				// might be misinterpreted:
				wait_indefinitely = false;
				sleep_duration = (int)(ATOF(cp + 1) * 1000);
				break;
			}
		}
		// The following must be set for ScriptGetJoyState():
		token.symbol = SYM_STRING;
		token.marker = buf;
	}
	else if (   (mActionType != ACT_RUNWAIT && mActionType != ACT_CLIPWAIT && *ARG3)
		|| (mActionType == ACT_CLIPWAIT && *ARG1)   )
	{
		// Since the param containing the timeout value isn't blank, it must be numeric,
		// otherwise, the loading validation would have prevented the script from loading.
		wait_indefinitely = false;
		sleep_duration = (int)(ATOF(mActionType == ACT_CLIPWAIT ? ARG1 : ARG3) * 1000); // Can be zero.
		if (sleep_duration < 1)
			// Waiting 500ms in place of a "0" seems more useful than a true zero, which
			// doens't need to be supported because it's the same thing as something like
			// "IfWinExist".  A true zero for clipboard would be the same as
			// "IfEqual, clipboard, , xxx" (though admittedly it's higher overhead to
			// actually fetch the contents of the clipboard).
			sleep_duration = 500;
	}
	else
	{
		wait_indefinitely = true;
		sleep_duration = 0; // Just to catch any bugs.
	}

	if (mActionType != ACT_RUNWAIT)
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Set default ErrorLevel to be possibly overridden later on.

	bool any_clipboard_format = (mActionType == ACT_CLIPWAIT && ArgToInt(2) == 1);

	// Right before starting the wait-loop, make a copy of our args using the stack
	// space in our recursion layer.  This is done in case other hotkey subroutine(s)
	// are launched while we're waiting here, which might cause our args to be overwritten
	// if any of them happen to be in the Deref buffer:
	LPTSTR arg[MAX_ARGS], marker;
	int i, space_remaining;
	for (i = 0, space_remaining = LINE_SIZE, marker = buf; i < mArgc; ++i)
	{
		if (!space_remaining) // Realistically, should never happen.
			arg[i] = _T("");
		else
		{
			arg[i] = marker;  // Point it to its place in the buffer.
			tcslcpy(marker, sArgDeref[i], space_remaining); // Make the copy.
			marker += _tcslen(marker) + 1;  // +1 for the zero terminator of each arg.
			space_remaining = (int)(LINE_SIZE - (marker - buf));
		}
	}

	for (start_time = GetTickCount();;) // start_time is initialized unconditionally for use with v1.0.30.02's new logging feature further below.
	{ // Always do the first iteration so that at least one check is done.
		switch(mActionType)
		{
		case ACT_WINWAIT:
			#define SAVED_WIN_ARGS SAVED_ARG1, SAVED_ARG2, SAVED_ARG4, SAVED_ARG5
			if (WinExist(*g, SAVED_WIN_ARGS, false, true))
			{
				DoWinDelay;
				return OK;
			}
			break;
		case ACT_WINWAITCLOSE:
			if (!WinExist(*g, SAVED_WIN_ARGS))
			{
				DoWinDelay;
				return OK;
			}
			break;
		case ACT_WINWAITACTIVE:
			if (WinActive(*g, SAVED_WIN_ARGS, true))
			{
				DoWinDelay;
				return OK;
			}
			break;
		case ACT_WINWAITNOTACTIVE:
			if (!WinActive(*g, SAVED_WIN_ARGS, true))
			{
				DoWinDelay;
				return OK;
			}
			break;
		case ACT_CLIPWAIT:
			// Seems best to consider CF_HDROP to be a non-empty clipboard, since we
			// support the implicit conversion of that format to text:
			if (any_clipboard_format)
			{
				if (CountClipboardFormats())
					return OK;
			}
			else
				if (IsClipboardFormatAvailable(CF_NATIVETEXT) || IsClipboardFormatAvailable(CF_HDROP))
					return OK;
			break;
		case ACT_KEYWAIT:
			if (vk) // Waiting for key or mouse button, not joystick.
			{
				if (ScriptGetKeyState(vk, key_state_type) == wait_for_keydown)
					return OK;
			}
			else // Waiting for joystick button
			{
				if ((bool)ScriptGetJoyState(joy, joystick_id, token, false) == wait_for_keydown)
					return OK;
			}
			break;
		case ACT_RUNWAIT:
			// Pretty nasty, but for now, nothing is done to prevent an infinite loop.
			// In the future, maybe OpenProcess() can be used to detect if a process still
			// exists (is there any other way?):
			// MSDN: "Warning: If a process happens to return STILL_ACTIVE (259) as an error code,
			// applications that test for this value could end up in an infinite loop."
			if (running_process)
				GetExitCodeProcess(running_process, &exit_code);
			else // it can be NULL in the case of launching things like "find D:\" or "www.yahoo.com"
				exit_code = 0;
			if (exit_code != STATUS_PENDING) // STATUS_PENDING == STILL_ACTIVE
			{
				if (running_process)
					CloseHandle(running_process);
				// Use signed vs. unsigned, since that is more typical?  No, it seems better
				// to use unsigned now that script variables store 64-bit ints.  This is because
				// GetExitCodeProcess() yields a DWORD, implying that the value should be unsigned.
				// Unsigned also is more useful in cases where an app returns a (potentially large)
				// count of something as its result.  However, if this is done, it won't be easy
				// to check against a return value of -1, for example, which I suspect many apps
				// return.  AutoIt3 (and probably 2) use a signed int as well, so that is another
				// reason to keep it this way:
				return g_ErrorLevel->Assign((int)exit_code);
			}
			break;
		}

		// Must cast to int or any negative result will be lost due to DWORD type:
		if (wait_indefinitely || (int)(sleep_duration - (GetTickCount() - start_time)) > SLEEP_INTERVAL_HALF)
		{
			if (MsgSleep(INTERVAL_UNSPECIFIED)) // INTERVAL_UNSPECIFIED performs better.
			{
				// v1.0.30.02: Since MsgSleep() launched and returned from at least one new thread, put the
				// current waiting line into the line-log again to make it easy to see what the current
				// thread is doing.  This is especially useful for figuring out which subroutine is holding
				// another thread interrupted beneath it.  For example, if a timer gets interrupted by
				// a hotkey that has an indefinite WinWait, and that window never appears, this will allow
				// the user to find out the culprit thread by showing its line in the log (and usually
				// it will appear as the very last line, since usually the script is idle and thus the
				// currently active thread is the one that's still waiting for the window).
				if (g->ListLinesIsEnabled)
				{
					// ListLines is enabled in this thread, but if it was disabled in the interrupting thread,
					// the very last log entry will be ours.  In that case, we don't want to duplicate it.
					int previous_log_index = (sLogNext ? sLogNext : LINE_LOG_SIZE) - 1; // Wrap around if needed (the entry can be NULL in that case).
					if (sLog[previous_log_index] != this || sLogTick[previous_log_index] != start_time) // The previously logged line was not this one, or it was added by the interrupting thread (different start_time).
					{
						sLog[sLogNext] = this;
						sLogTick[sLogNext++] = start_time; // Store a special value so that Line::LogToText() can report that its "still waiting" from earlier.
						if (sLogNext >= LINE_LOG_SIZE)
							sLogNext = 0;
						// The lines above are the similar to those used in ExecUntil(), so the two should be
						// maintained together.
					}
				}
			}
		}
		else // Done waiting.
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Since it timed out, we override the default with this.
	} // for()
}



ResultType Line::WinMove(LPTSTR aTitle, LPTSTR aText, LPTSTR aX, LPTSTR aY
	, LPTSTR aWidth, LPTSTR aHeight, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	// So that compatibility is retained, don't set ErrorLevel for commands that are native to AutoIt2
	// but that AutoIt2 doesn't use ErrorLevel with (such as this one).
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		return OK;
	RECT rect;
	if (!GetWindowRect(target_window, &rect))
		return OK;  // Can't set errorlevel, see above.
	MoveWindow(target_window
		, *aX && _tcsicmp(aX, _T("default")) ? ATOI(aX) : rect.left  // X-position
		, *aY && _tcsicmp(aY, _T("default")) ? ATOI(aY) : rect.top   // Y-position
		, *aWidth && _tcsicmp(aWidth, _T("default")) ? ATOI(aWidth) : rect.right - rect.left
		, *aHeight && _tcsicmp(aHeight, _T("default")) ? ATOI(aHeight) : rect.bottom - rect.top
		, TRUE);  // Do repaint.
	DoWinDelay;
	return OK;
}



ResultType Line::ControlSend(LPTSTR aControl, LPTSTR aKeysToSend, LPTSTR aTitle, LPTSTR aText
	, LPTSTR aExcludeTitle, LPTSTR aExcludeText, SendRawModes aSendRaw)
{
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		goto error;
	HWND control_window = _tcsicmp(aControl, _T("ahk_parent"))
		? ControlExist(target_window, aControl) // This can return target_window itself for cases such as ahk_id %ControlHWND%.
		: target_window;
	if (!control_window)
		goto error;
	SendKeys(aKeysToSend, aSendRaw, SM_EVENT, control_window);
	// But don't do WinDelay because KeyDelay should have been in effect for the above.
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.

error:
	return SetErrorLevelOrThrow();
}



ResultType Line::ControlClick(vk_type aVK, int aClickCount, LPTSTR aOptions, LPTSTR aControl
	, LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		goto error;

	// Set the defaults that will be in effect unless overridden by options:
	KeyEventTypes event_type = KEYDOWNANDUP;
	bool position_mode = false;
	bool do_activate = true;
	// These default coords can be overridden either by aOptions or aControl's X/Y mode:
	{POINT click = {COORD_UNSPECIFIED, COORD_UNSPECIFIED};

	for (LPTSTR cp = aOptions; *cp; ++cp)
	{
		switch(ctoupper(*cp))
		{
		case 'D':
			event_type = KEYDOWN;
			break;
		case 'U':
			event_type = KEYUP;
			break;
		case 'N':
			// v1.0.45:
			// It was reported (and confirmed through testing) that this new NA mode (which avoids
			// AttachThreadInput() and SetActiveWindow()) improves the reliability of ControlClick when
			// the user is moving the mouse fairly quickly at the time the command tries to click a button.
			// In addition, the new mode avoids activating the window, which tends to happen otherwise.
			// HOWEVER, the new mode seems no more reliable than the old mode when the target window is
			// the active window.  In addition, there may be side-effects of the new mode (I caught it
			// causing Notepad's Save-As dialog to hang once, during the display of its "Overwrite?" dialog).
			// ALSO, SetControlDelay -1 seems to fix the unreliability issue as well (independently of NA),
			// though it might not work with some types of windows/controls (thus, for backward
			// compatibility, ControlClick still obeys SetControlDelay).
			if (ctoupper(cp[1]) == 'A')
			{
				cp += 1;  // Add 1 vs. 2 to skip over the rest of the letters in this option word.
				do_activate = false;
			}
			break;
		case 'P':
			if (!_tcsnicmp(cp, _T("Pos"), 3))
			{
				cp += 2;  // Add 2 vs. 3 to skip over the rest of the letters in this option word.
				position_mode = true;
			}
			break;
		// For the below:
		// Use atoi() vs. ATOI() to avoid interpreting something like 0x01D as hex
		// when in fact the D was meant to be an option letter:
		case 'X':
			click.x = _ttoi(cp + 1); // Will be overridden later below if it turns out that position_mode is in effect.
			break;
		case 'Y':
			click.y = _ttoi(cp + 1); // Will be overridden later below if it turns out that position_mode is in effect.
			break;
		}
	}

	// It's debatable, but might be best for flexibility (and backward compatibility) to allow target_window to itself
	// be a control (at least for the position_mode handler below).  For example, the script may have called SetParent
	// to make a top-level window the child of some other window, in which case this policy allows it to be seen like
	// a non-child.
	HWND control_window = position_mode ? NULL : ControlExist(target_window, aControl); // This can return target_window itself for cases such as ahk_id %ControlHWND%.
	if (!control_window) // Even if position_mode is false, the below is still attempted, as documented.
	{
		// New section for v1.0.24.  But only after the above fails to find a control do we consider
		// whether aControl contains X and Y coordinates.  That way, if a control class happens to be
		// named something like "X1 Y1", it will still be found by giving precedence to class names.
		point_and_hwnd_type pah = {0};
		pah.ignore_disabled_controls = true; // v1.1.20: Ignore disabled controls.
		// Parse the X an Y coordinates in a strict way to reduce ambiguity with control names and also
		// to keep the code simple.
		LPTSTR cp = omit_leading_whitespace(aControl);
		if (ctoupper(*cp) != 'X')
			goto error;
		++cp;
		if (!*cp)
			goto error;
		pah.pt.x = ATOI(cp);
		if (   !(cp = StrChrAny(cp, _T(" \t")))   ) // Find next space or tab (there must be one for it to be considered valid).
			goto error;
		cp = omit_leading_whitespace(cp + 1);
		if (!*cp || _totupper(*cp) != 'Y')
			goto error;
		++cp;
		if (!*cp)
			goto error;
		pah.pt.y = ATOI(cp);
		// The passed-in coordinates are always relative to target_window's upper left corner because offering
		// an option for absolute/screen coordinates doesn't seem useful.
		RECT rect;
		GetWindowRect(target_window, &rect);
		pah.pt.x += rect.left; // Convert to screen coordinates.
		pah.pt.y += rect.top;
		EnumChildWindows(target_window, EnumChildFindPoint, (LPARAM)&pah); // Find topmost control containing point.
		// If no control is at this point, try posting the mouse event message(s) directly to the
		// parent window to increase the flexibility of this feature:
		control_window = pah.hwnd_found ? pah.hwnd_found : target_window;
		// Convert click's target coordinates to be relative to the client area of the control or
		// parent window because that is the format required by messages such as WM_LBUTTONDOWN
		// used later below:
		click = pah.pt;
		ScreenToClient(control_window, &click);
	}

	// This is done this late because it seems better to set an ErrorLevel of 1 (above) whenever the
	// target window or control isn't found, or any other error condition occurs above:
	if (aClickCount < 1)
		// Allow this to simply "do nothing", because it increases flexibility
		// in the case where the number of clicks is a dereferenced script variable
		// that may sometimes (by intent) resolve to zero or negative:
		return g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	RECT rect;
	if (click.x == COORD_UNSPECIFIED || click.y == COORD_UNSPECIFIED)
	{
		// The following idea is from AutoIt3. It states: "Get the dimensions of the control so we can click
		// the centre of it" (maybe safer and more natural than 0,0).
		// My: In addition, this is probably better for some large controls (e.g. SysListView32) because
		// clicking at 0,0 might activate a part of the control that is not even visible:
		if (!GetWindowRect(control_window, &rect))
			goto error;
		if (click.x == COORD_UNSPECIFIED)
			click.x = (rect.right - rect.left) / 2;
		if (click.y == COORD_UNSPECIFIED)
			click.y = (rect.bottom - rect.top) / 2;
	}
	LPARAM lparam = MAKELPARAM(click.x, click.y);

	UINT msg_down, msg_up;
	WPARAM wparam, wparam_up = 0;
	bool vk_is_wheel = aVK == VK_WHEEL_UP || aVK == VK_WHEEL_DOWN;
	bool vk_is_hwheel = aVK == VK_WHEEL_LEFT || aVK == VK_WHEEL_RIGHT; // v1.0.48: Lexikos: Support horizontal scrolling in Windows Vista and later.

	if (vk_is_wheel)
	{
		wparam = (aClickCount * ((aVK == VK_WHEEL_UP) ? WHEEL_DELTA : -WHEEL_DELTA)) << 16;  // High order word contains the delta.
		msg_down = WM_MOUSEWHEEL;
		// Make the event more accurate by having the state of the keys reflected in the event.
		// The logical state (not physical state) of the modifier keys is used so that something
		// like this is supported:
		// Send, {ShiftDown}
		// MouseClick, WheelUp
		// Send, {ShiftUp}
		// In addition, if the mouse hook is installed, use its logical mouse button state so that
		// something like this is supported:
		// MouseClick, left, , , , , D  ; Hold down the left mouse button
		// MouseClick, WheelUp
		// MouseClick, left, , , , , U  ; Release the left mouse button.
		// UPDATE: Since the other ControlClick types (such as leftclick) do not reflect these
		// modifiers -- and we want to keep it that way, at least by default, for compatibility
		// reasons -- it seems best for consistency not to do them for WheelUp/Down either.
		// A script option can be added in the future to obey the state of the modifiers:
		//mod_type mod = GetModifierState();
		//if (mod & MOD_SHIFT)
		//	wparam |= MK_SHIFT;
		//if (mod & MOD_CONTROL)
		//	wparam |= MK_CONTROL;
        //if (g_MouseHook)
		//	wparam |= g_mouse_buttons_logical;
	}
	else if (vk_is_hwheel)	// Lexikos: Support horizontal scrolling in Windows Vista and later.
	{
		wparam = (aClickCount * ((aVK == VK_WHEEL_LEFT) ? -WHEEL_DELTA : WHEEL_DELTA)) << 16;
		msg_down = WM_MOUSEHWHEEL;
	}
	else
	{
		switch (aVK)
		{
			case VK_LBUTTON:  msg_down = WM_LBUTTONDOWN; msg_up = WM_LBUTTONUP; wparam = MK_LBUTTON; break;
			case VK_RBUTTON:  msg_down = WM_RBUTTONDOWN; msg_up = WM_RBUTTONUP; wparam = MK_RBUTTON; break;
			case VK_MBUTTON:  msg_down = WM_MBUTTONDOWN; msg_up = WM_MBUTTONUP; wparam = MK_MBUTTON; break;
			case VK_XBUTTON1: msg_down = WM_XBUTTONDOWN; msg_up = WM_XBUTTONUP; wparam_up = XBUTTON1<<16; wparam = MK_XBUTTON1|wparam_up; break;
			case VK_XBUTTON2: msg_down = WM_XBUTTONDOWN; msg_up = WM_XBUTTONUP; wparam_up = XBUTTON2<<16; wparam = MK_XBUTTON2|wparam_up; break;
			default: goto error; // Just do nothing since this should realistically never happen.
		}
	}

	// SetActiveWindow() requires ATTACH_THREAD_INPUT to succeed.  Even though the MSDN docs state
	// that SetActiveWindow() has no effect unless the parent window is foreground, Jon insists
	// that SetActiveWindow() resolved some problems for some users.  In any case, it seems best
	// to do this in case the window really is foreground, in which case MSDN indicates that
	// it will help for certain types of dialogs.
	ATTACH_THREAD_INPUT_AND_SETACTIVEWINDOW_IF_DO_ACTIVATE  // It's kept with a similar macro for maintainability.
	// v1.0.44.13: Notes for the above: Unlike some other Control commands, GetNonChildParent() is not
	// called here when target_window==control_window.  This is because the script may have called
	// SetParent to make target_window the child of some other window, in which case target_window
	// should still be used above (unclear).  Perhaps more importantly, it's allowed for control_window
	// to be the same as target_window, at least in position_mode, whose docs state, "If there is no
	// control, the target window itself will be sent the event (which might have no effect depending
	// on the nature of the window)."  In other words, it seems too complicated and rare to add explicit
	// handling for "ahk_id %ControlHWND%" (though the below rules should work).
	// The line "ControlClick,, ahk_id %HWND%" can have multiple meanings depending on the nature of HWND:
	// 1) If HWND is a top-level window, its topmost child will be clicked.
	// 2) If HWND is a top-level window that has become a child of another window via SetParent: same.
	// 3) If HWND is a control, its topmost child will be clicked (or itself if it has no children).
	//    For example, the following works (as documented in the first parameter):
	//    ControlGet, HWND, HWND,, OK, A  ; Get the HWND of the OK button.
	//    ControlClick,, ahk_id %HWND%

	if (vk_is_wheel || vk_is_hwheel) // v1.0.48: Lexikos: Support horizontal scrolling in Windows Vista and later.
	{
		PostMessage(control_window, msg_down, wparam, lparam);
		DoControlDelay;
	}
	else
	{
		for (int i = 0; i < aClickCount; ++i)
		{
			if (event_type != KEYUP) // It's either down-only or up-and-down so always to the down-event.
			{
				PostMessage(control_window, msg_down, wparam, lparam);
				// Seems best to do this one too, which is what AutoIt3 does also.  User can always reduce
				// ControlDelay to 0 or -1.  Update: Jon says this delay might be causing it to fail in
				// some cases.  Upon reflection, it seems best not to do this anyway because PostMessage()
				// should queue up the message for the app correctly even if it's busy.  Update: But I
				// think the timestamp is available on every posted message, so if some apps check for
				// inhumanly fast clicks (to weed out transients with partial clicks of the mouse, or
				// to detect artificial input), the click might not work.  So it might be better after
				// all to do the delay until it's proven to be problematic (Jon implies that he has
				// no proof yet).  IF THIS IS EVER DISABLED, be sure to do the ControlDelay anyway
				// if event_type == KEYDOWN:
				DoControlDelay;
			}
			if (event_type != KEYDOWN) // It's either up-only or up-and-down so always to the up-event.
			{
				PostMessage(control_window, msg_up, wparam_up, lparam);
				DoControlDelay;
			}
		}
	}

	DETACH_THREAD_INPUT  // Also takes into account do_activate, indirectly.

	return g_ErrorLevel->Assign(ERRORLEVEL_NONE);} // Indicate success.

error:
	return SetErrorLevelOrThrow();
}



ResultType Line::ControlMove(LPTSTR aControl, LPTSTR aX, LPTSTR aY, LPTSTR aWidth, LPTSTR aHeight
	, LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		goto error;
	HWND control_window = ControlExist(target_window, aControl); // This can return target_window itself for cases such as ahk_id %ControlHWND%.
	if (!control_window)
		goto error;

	POINT point;
	point.x = *aX ? ATOI(aX) : COORD_UNSPECIFIED;
	point.y = *aY ? ATOI(aY) : COORD_UNSPECIFIED;

	// First convert the user's given coordinates -- which by default are relative to the window's
	// upper left corner -- to screen coordinates:
	if (point.x != COORD_UNSPECIFIED || point.y != COORD_UNSPECIFIED)
	{
		RECT rect;
		// v1.0.44.13: Below was fixed to allow for the fact that target_window might be the control
		// itself (e.g. via ahk_id %ControlHWND%).  For consistency with ControlGetPos and other things,
		// it seems best to call GetNonChildParent rather than GetParent(); for example, a Tab control
		// that contains a child window that in turn contains the actual controls should probably report
		// the position of each control relative to the dialog itself rather than the tab control or its
		// master window.  The lost argument in favor of GetParent is that it seems more flexible, such
		// as cases where the script has called SetParent() to make a top-level window the child of some
		// other window, in which case the target control's immediate parent should be used, not its most
		// distant ancestor. This might also be desirable for controls that are children of other controls,
		// such as Combobox's Edit.
		if (!GetWindowRect(target_window == control_window ? GetNonChildParent(target_window) : target_window
			, &rect))
			goto error;
		if (point.x != COORD_UNSPECIFIED)
			point.x += rect.left;
		if (point.y != COORD_UNSPECIFIED)
			point.y += rect.top;
	}

	// If either coordinate is unspecified, put the control's current screen coordinate(s)
	// into point:
	RECT control_rect;
	if (!GetWindowRect(control_window, &control_rect))
		goto error;
	if (point.x == COORD_UNSPECIFIED)
		point.x = control_rect.left;
	if (point.y == COORD_UNSPECIFIED)
		point.y = control_rect.top;

	// Use the immediate parent since controls can themselves have child controls:
	HWND immediate_parent = GetParent(control_window);
	if (!immediate_parent)
		goto error;

	// Convert from absolute screen coordinates to coordinates used with MoveWindow(),
	// which are relative to control_window's parent's client area:
	if (!ScreenToClient(immediate_parent, &point))
		goto error;

	MoveWindow(control_window
		, point.x
		, point.y
		, *aWidth ? ATOI(aWidth) : control_rect.right - control_rect.left
		, *aHeight ? ATOI(aHeight) : control_rect.bottom - control_rect.top
		, TRUE);  // Do repaint.

	DoControlDelay
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.

error:
	return SetErrorLevelOrThrow();
}



ResultType Line::ControlGetPos(LPTSTR aControl, LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	Var *output_var_x = ARGVAR1;  // Ok if NULL. Load-time validation has ensured that these are valid output variables (e.g. not built-in vars).
	Var *output_var_y = ARGVAR2;  // Ok if NULL.
	Var *output_var_width = ARGVAR3;  // Ok if NULL.
	Var *output_var_height = ARGVAR4;  // Ok if NULL.

	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	HWND control_window = target_window ? ControlExist(target_window, aControl) : NULL; // This can return target_window itself for cases such as ahk_id %ControlHWND%.
	if (!control_window)
	{
		if (output_var_x)
			output_var_x->Assign();
		if (output_var_y)
			output_var_y->Assign();
		if (output_var_width)
			output_var_width->Assign();
		if (output_var_height)
			output_var_height->Assign();
		return OK;
	}

	RECT parent_rect, child_rect;
	// Realistically never fails since DetermineTargetWindow() and ControlExist() should always yield
	// valid window handles:
	GetWindowRect(target_window == control_window ? GetNonChildParent(target_window) : target_window
		, &parent_rect); // v1.0.44.13: Above was fixed to allow for the fact that target_window might be the control itself (e.g. via ahk_id %ControlHWND%).  See ControlMove for details.
	GetWindowRect(control_window, &child_rect);

	if (output_var_x && !output_var_x->Assign(child_rect.left - parent_rect.left))
		return FAIL;
	if (output_var_y && !output_var_y->Assign(child_rect.top - parent_rect.top))
		return FAIL;
	if (output_var_width && !output_var_width->Assign(child_rect.right - child_rect.left))
		return FAIL;
	if (output_var_height && !output_var_height->Assign(child_rect.bottom - child_rect.top))
		return FAIL;

	return OK;
}



ResultType Line::ControlGetFocus(LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	Var &output_var = *OUTPUT_VAR; // Must be resolved only once and prior to DetermineTargetWindow().  See Line::WinGetClass() for explanation.
	output_var.Assign();  // Set default: blank for the output variable.
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		goto error;

	GUITHREADINFO guithreadInfo;
	guithreadInfo.cbSize = sizeof(GUITHREADINFO);
	if (!GetGUIThreadInfo(GetWindowThreadProcessId(target_window, NULL), &guithreadInfo))
		goto error;

	class_and_hwnd_type cah;
	TCHAR class_name[WINDOW_CLASS_SIZE];
	cah.hwnd = guithreadInfo.hwndFocus;
	cah.class_name = class_name;
	if (!GetClassName(cah.hwnd, class_name, _countof(class_name) - 5)) // -5 to allow room for sequence number.
		goto error;
	
	cah.class_count = 0;  // Init for the below.
	cah.is_found = false; // Same.
	EnumChildWindows(target_window, EnumChildFindSeqNum, (LPARAM)&cah);
	if (!cah.is_found)
		goto error;
	// Append the class sequence number onto the class name set the output param to be that value:
	sntprintfcat(class_name, _countof(class_name), _T("%d"), cah.class_count);
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return output_var.Assign(class_name);

error:
	return SetErrorLevelOrThrow();
}



BOOL CALLBACK EnumChildFindSeqNum(HWND aWnd, LPARAM lParam)
{
	class_and_hwnd_type &cah = *(class_and_hwnd_type *)lParam;  // For performance and convenience.
	TCHAR class_name[WINDOW_CLASS_SIZE];
	if (!GetClassName(aWnd, class_name, _countof(class_name)))
		return TRUE;  // Continue the enumeration.
	if (!_tcscmp(class_name, cah.class_name)) // Class names match.
	{
		++cah.class_count;
		if (aWnd == cah.hwnd)  // The caller-specified window has been found.
		{
			cah.is_found = true;
			return FALSE;
		}
	}
	return TRUE; // Continue enumeration until a match is found or there aren't any windows remaining.
}



ResultType Line::ControlFocus(LPTSTR aControl, LPTSTR aTitle, LPTSTR aText
	, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		goto error;
	HWND control_window = ControlExist(target_window, aControl); // This can return target_window itself for cases such as ahk_id %ControlHWND%.
	if (!control_window)
		goto error;

	// Unlike many of the other Control commands, this one requires AttachThreadInput()
	// to have any realistic chance of success (though sometimes it may work by pure
	// chance even without it):
	ATTACH_THREAD_INPUT

	if (SetFocus(control_window))
	{
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		DoControlDelay;
	}

	// Very important to detach any threads whose inputs were attached above,
	// prior to returning, otherwise the next attempt to attach thread inputs
	// for these particular windows may result in a hung thread or other
	// undesirable effect:
	DETACH_THREAD_INPUT

	return OK;

error:
	return SetErrorLevelOrThrow();
}



ResultType Line::ControlSetText(LPTSTR aControl, LPTSTR aNewText, LPTSTR aTitle, LPTSTR aText
	, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		goto error;
	HWND control_window = ControlExist(target_window, aControl); // This can return target_window itself for cases such as ahk_id %ControlHWND%.
	if (!control_window)
		goto error;
	// SendMessage must be used, not PostMessage(), at least for some (probably most) apps.
	// Also: No need to call IsWindowHung() because SendMessageTimeout() should return
	// immediately if the OS already "knows" the window is hung:
	DWORD_PTR result;
	SendMessageTimeout(control_window, WM_SETTEXT, (WPARAM)0, (LPARAM)aNewText
		, SMTO_ABORTIFHUNG, 5000, &result);
	DoControlDelay;
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.

error:
	return SetErrorLevelOrThrow();
}



ResultType Line::ControlGetText(LPTSTR aControl, LPTSTR aTitle, LPTSTR aText
	, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	Var &output_var = *OUTPUT_VAR;
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	HWND control_window = target_window ? ControlExist(target_window, aControl) : NULL; // This can return target_window itself for cases such as ahk_id %ControlHWND%.
	// Even if control_window is NULL, we want to continue on so that the output
	// param is set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before.

	// Handle the output parameter.  This section is similar to that in
	// PerformAssign().  Note: Using GetWindowTextTimeout() vs. GetWindowText()
	// because it is able to get text from more types of controls (e.g. large edit controls):
	VarSizeType space_needed = control_window ? GetWindowTextTimeout(control_window) + 1 : 1; // 1 for terminator.
	if (space_needed > g_MaxVarCapacity) // Allow the command to succeed by truncating the text.
		space_needed = g_MaxVarCapacity;

	// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
	// this call will set up the clipboard for writing:
	if (output_var.AssignString(NULL, space_needed - 1) != OK)
		return FAIL;  // It already displayed the error.
	// Fetch the text directly into the var.  Also set the length explicitly
	// in case actual size written was off from the estimated size (since
	// GetWindowTextLength() can return more space that will actually be required
	// in certain circumstances, see MS docs):
	if (control_window)
	{
		if (   !(output_var.SetCharLength((VarSizeType)GetWindowTextTimeout(control_window
			, output_var.Contents(), space_needed)))   ) // There was no text to get or GetWindowTextTimeout() failed.
			*output_var.Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
	}
	else
	{
		*output_var.Contents() = '\0';
		output_var.SetCharLength(0);
		// And leave bSucceeded set to false to distinguish a non-existent control
		// from a one that does exist but returns no text.
	}
	// Consider the above to be always successful, even if the window wasn't found, except
	// when below returns an error:
	ResultType result = output_var.Close(); // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_CLIPBOARD.
	if (result != OK)
		return result;
	return SetErrorLevelOrThrowBool(!control_window);
}



ResultType Line::ControlGetListView(Var &aOutputVar, HWND aHwnd, LPTSTR aOptions)
// Called by ControlGet() below.  It has ensured that aHwnd is a valid handle to a ListView.
{
	aOutputVar.Assign(); // Init to blank in case of early return.

	// GET ROW COUNT
	LRESULT row_count;
	if (!SendMessageTimeout(aHwnd, LVM_GETITEMCOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&row_count)) // Timed out or failed.
		return SetErrorLevelOrThrow();

	// GET COLUMN COUNT
	// Through testing, could probably get to a level of 90% certainty that a ListView for which
	// InsertColumn() was never called (or was called only once) might lack a header control if the LV is
	// created in List/Icon view-mode and/or with LVS_NOCOLUMNHEADER. The problem is that 90% doesn't
	// seem to be enough to justify elimination of the code for "undetermined column count" mode.  If it
	// ever does become a certainty, the following could be changed:
	// 1) The extra code for "undetermined" mode rather than simply forcing col_count to be 1.
	// 2) Probably should be kept for compatibility: -1 being returned when undetermined "col count".
	//
	// The following approach might be the only simple yet reliable way to get the column count (sending
	// LVM_GETITEM until it returns false doesn't work because it apparently returns true even for
	// nonexistent subitems -- the same is reported to happen with LVM_GETCOLUMN and such, though I seem
	// to remember that LVM_SETCOLUMN fails on non-existent columns -- but calling that on a ListView
	// that isn't in Report view has been known to traumatize the control).
	// Fix for v1.0.37.01: It appears that the header doesn't always exist.  For example, when an
	// Explorer window opens and is *initially* in icon or list mode vs. details/tiles mode, testing
	// shows that there is no header control.  Testing also shows that there is exactly one column
	// in such cases but only for Explorer and other things that avoid creating the invisible columns.
	// For example, a script can create a ListView in Icon-mode and give it retrievable column data for
	// columns beyond the first.  Thus, having the undetermined-col-count mode preserves flexibility
	// by allowing individual columns beyond the first to be retrieved.  On a related note, testing shows
	// that attempts to explicitly retrieve columns (i.e. fields/subitems) other than the first in the
	// case of Explorer's Icon/List view modes behave the same as fetching the first column (i.e. Col3
	// would retrieve the same text as specifying Col1 or not having the Col option at all).
	// Obsolete because not always true: Testing shows that a ListView always has a header control
	// (at least on XP), even if you can't see it (such as when the view is Icon/Tile or when -Hdr has
	// been specified in the options).
	HWND header_control;
	LRESULT col_count = -1;  // Fix for v1.0.37.01: Use -1 to indicate "undetermined col count".
	if (SendMessageTimeout(aHwnd, LVM_GETHEADER, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&header_control)
		&& header_control) // Relies on short-circuit boolean order.
		SendMessageTimeout(header_control, HDM_GETITEMCOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&col_count);
		// Return value is not checked because if it fails, col_count is left at its default of -1 set above.
		// In fact, if any of the above conditions made it impossible to determine col_count, col_count stays
		// at -1 to indicate "undetermined".

	// PARSE OPTIONS (a simple vs. strict method is used to reduce code size)
	bool get_count = tcscasestr(aOptions, _T("Count"));
	bool include_selected_only = tcscasestr(aOptions, _T("Selected")); // Explicit "ed" to reserve "Select" for possible future use.
	bool include_focused_only = tcscasestr(aOptions, _T("Focused"));  // Same.
	LPTSTR col_option = tcscasestr(aOptions, _T("Col")); // Also used for mode "Count Col"
	int requested_col = col_option ? ATOI(col_option + 3) - 1 : -1;
	// If the above yields a negative col number for any reason, it's ok because below will just ignore it.
	if (col_count > -1 && requested_col > -1 && requested_col >= col_count) // Specified column does not exist.
		return SetErrorLevelOrThrow();

	// IF THE "COUNT" OPTION IS PRESENT, FULLY HANDLE THAT AND RETURN
	if (get_count)
	{
		int result; // Must be signed to support writing a col count of -1 to aOutputVar.
		if (include_focused_only) // Listed first so that it takes precedence over include_selected_only.
		{
			if (!SendMessageTimeout(aHwnd, LVM_GETNEXTITEM, -1, LVNI_FOCUSED, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&result)) // Timed out or failed.
				return SetErrorLevelOrThrow();
			++result; // i.e. Set it to 0 if not found, or the 1-based row-number otherwise.
		}
		else if (include_selected_only)
		{
			if (!SendMessageTimeout(aHwnd, LVM_GETSELECTEDCOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&result)) // Timed out or failed.
				return SetErrorLevelOrThrow();
		}
		else if (col_option) // "Count Col" returns the number of columns.
			result = (int)col_count;
		else // Total row count.
			result = (int)row_count;
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		return aOutputVar.Assign(result);
	}

	// FINAL CHECKS
	if (row_count < 1 || !col_count) // But don't return when col_count == -1 (i.e. always make the attempt when col count is undetermined).
		return g_ErrorLevel->Assign(ERRORLEVEL_NONE);  // No text in the control, so indicate success.
	
	// Notes about the following struct definitions:  The layout of LVITEM depends on
	// which platform THIS executable was compiled for, but we need it to match what
	// the TARGET process expects.  If the target process is 32-bit and we are 64-bit
	// or vice versa, LVITEM can't be used.  The following structs are copies of
	// LVITEM with UINT (32-bit) or UINT64 (64-bit) in place of the pointer fields.
	struct LVITEM32
	{
		UINT mask;
		int iItem;
		int iSubItem;
		UINT state;
		UINT stateMask;
		UINT pszText;
		int cchTextMax;
		int iImage;
		UINT lParam;
		int iIndent;
		int iGroupId;
		UINT cColumns;
		UINT puColumns;
		UINT piColFmt;
		int iGroup;
	};
	struct LVITEM64
	{
		UINT mask;
		int iItem;
		int iSubItem;
		UINT state;
		UINT stateMask;
		UINT64 pszText;
		int cchTextMax;
		int iImage;
		UINT64 lParam;
		int iIndent;
		int iGroupId;
		UINT cColumns;
		UINT64 puColumns;
		UINT64 piColFmt;
		int iGroup;
	};
	union
	{
		LVITEM32 i32;
		LVITEM64 i64;
	} local_lvi;

	// ALLOCATE INTERPROCESS MEMORY FOR TEXT RETRIEVAL
	HANDLE handle;
	LPVOID p_remote_lvi; // Not of type LPLVITEM to help catch bugs where p_remote_lvi->member is wrongly accessed here in our process.
	if (   !(p_remote_lvi = AllocInterProcMem(handle, sizeof(local_lvi) + _TSIZE(LV_REMOTE_BUF_SIZE), aHwnd, PROCESS_QUERY_INFORMATION))   ) // Allocate both the LVITEM struct and its internal string buffer in one go because VirtualAllocEx() is probably a high overhead call.
		return SetErrorLevelOrThrow();
	LPVOID p_remote_text = (LPVOID)((UINT_PTR)p_remote_lvi + sizeof(local_lvi)); // The next buffer is the memory area adjacent to, but after the struct.
	
	// PREPARE LVI STRUCT MEMBERS FOR TEXT RETRIEVAL
	if (IsProcess64Bit(handle))
	{
		// See the section below for comments.
		local_lvi.i64.cchTextMax = LV_REMOTE_BUF_SIZE - 1;
		local_lvi.i64.pszText = (UINT64)p_remote_text;
	}
	else
	{
		// Subtract 1 because of that nagging doubt about size vs. length. Some MSDN examples subtract one,
		// such as TabCtrl_GetItem()'s cchTextMax:
		local_lvi.i32.cchTextMax = LV_REMOTE_BUF_SIZE - 1; // Note that LVM_GETITEM doesn't update this member to reflect the new length.
		local_lvi.i32.pszText = (UINT)(UINT_PTR)p_remote_text; // Extra cast avoids a truncation warning (C4311).
	}

	LRESULT i, next, length, total_length;
	bool is_selective = include_focused_only || include_selected_only;
	bool single_col_mode = (requested_col > -1 || col_count == -1); // Get only one column in these cases.

	// ESTIMATE THE AMOUNT OF MEMORY NEEDED TO STORE ALL THE TEXT
	// It's important to note that a ListView might legitimately have a collection of rows whose
	// fields are all empty.  Since it is difficult to know whether the control is truly owner-drawn
	// (checking its style might not be enough?), there is no way to distinguish this condition
	// from one where the control's text can't be retrieved due to being owner-drawn.  In any case,
	// this all-empty-field behavior simplifies the code and will be documented in the help file.
	for (i = 0, next = -1, total_length = 0; i < row_count; ++i) // For each row:
	{
		if (is_selective)
		{
			// Fix for v1.0.37.01: Prevent an infinite loop that might occur if the target control no longer
			// exists (perhaps having been closed in the middle of the operation) or is permanently hung.
			// If GetLastError() were to return zero after the below, it would mean the function timed out.
			// However, rather than checking and retrying, it seems better to abort the operation because:
			// 1) Timeout should be quite rare.
			// 2) Reduces code size.
			// 3) Having a retry really should be accompanied by SLEEP_WITHOUT_INTERRUPTION because all this
			//    time our thread would not pumping messages (and worse, if the keyboard/mouse hooks are installed,
			//    mouse/key lag would occur).
			if (!SendMessageTimeout(aHwnd, LVM_GETNEXTITEM, next, include_focused_only ? LVNI_FOCUSED : LVNI_SELECTED
				, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&next) // Timed out or failed.
				|| next == -1) // No next item.  Relies on short-circuit boolean order.
				break; // End of estimation phase (if estimate is too small, the text retrieval below will truncate it).
		}
		else
			next = i;
		for (local_lvi.i32.iSubItem = (requested_col > -1) ? requested_col : 0 // iSubItem is which field to fetch. If it's zero, the item vs. subitem will be fetched.
			; col_count == -1 || local_lvi.i32.iSubItem < col_count // If column count is undetermined (-1), always make the attempt.
			; ++local_lvi.i32.iSubItem) // For each column:
		{
			if (WriteProcessMemory(handle, p_remote_lvi, &local_lvi, sizeof(local_lvi), NULL)
				&& SendMessageTimeout(aHwnd, LVM_GETITEMTEXT, next, (LPARAM)p_remote_lvi, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&length))
				total_length += length;
			//else timed out or failed, don't include the length in the estimate.  Instead, the
			// text-fetching routine below will ensure the text doesn't overflow the var capacity.
			if (single_col_mode)
				break;
		}
	}
	// Add to total_length enough room for one linefeed per row, and one tab after each column
	// except the last (formula verified correct, though it's inflated by 1 for safety). "i" contains the
	// actual number of rows that will be transcribed, which might be less than row_count if is_selective==true.
	total_length += i * (single_col_mode ? 1 : col_count);

	// SET UP THE OUTPUT VARIABLE, ENLARGING IT IF NECESSARY
	// If the aOutputVar is of type VAR_CLIPBOARD, this call will set up the clipboard for writing:
	aOutputVar.AssignString(NULL, (VarSizeType)total_length, true, false); // Since failure is extremely rare, continue onward using the available capacity.
	LPTSTR contents = aOutputVar.Contents();
	LRESULT capacity = (int)aOutputVar.CharCapacity(); // LRESULT avoids signed vs. unsigned compiler warnings.
	if (capacity > 0) // For maintainability, avoid going negative.
		--capacity; // Adjust to exclude the zero terminator, which simplifies things below.

	// RETRIEVE THE TEXT FROM THE REMOTE LISTVIEW
	// Start total_length at zero in case actual size is greater than estimate, in which case only a partial set of text along with its '\t' and '\n' chars will be written.
	for (i = 0, next = -1, total_length = 0; i < row_count; ++i) // For each row:
	{
		if (is_selective)
		{
			// Fix for v1.0.37.01: Prevent an infinite loop (for details, see comments in the estimation phase above).
			if (!SendMessageTimeout(aHwnd, LVM_GETNEXTITEM, next, include_focused_only ? LVNI_FOCUSED : LVNI_SELECTED
				, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&next) // Timed out or failed.
				|| next == -1) // No next item.
				break; // See comment above for why unconditional break vs. continue.
		}
		else // Retrieve every row, so the "next" row becomes the "i" index.
			next = i;
		// Insert a linefeed before each row except the first:
		if (i && total_length < capacity) // If we're at capacity, it will exit the loops when the next field is read.
		{
			*contents++ = '\n';
			++total_length;
		}

		// iSubItem is which field to fetch. If it's zero, the item vs. subitem will be fetched:
		for (local_lvi.i32.iSubItem = (requested_col > -1) ? requested_col : 0
			; col_count == -1 || local_lvi.i32.iSubItem < col_count // If column count is undetermined (-1), always make the attempt.
			; ++local_lvi.i32.iSubItem) // For each column:
		{
			// Insert a tab before each column except the first and except when in single-column mode:
			if (!single_col_mode && local_lvi.i32.iSubItem && total_length < capacity)  // If we're at capacity, it will exit the loops when the next field is read.
			{
				*contents++ = '\t';
				++total_length;
			}

			if (!WriteProcessMemory(handle, p_remote_lvi, &local_lvi, sizeof(local_lvi), NULL)
				|| !SendMessageTimeout(aHwnd, LVM_GETITEMTEXT, next, (LPARAM)p_remote_lvi, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&length))
				continue; // Timed out or failed. It seems more useful to continue getting text rather than aborting the operation.

			// Otherwise, the message was successfully sent.
			if (length > 0)
			{
				if (total_length + length > capacity)
					goto break_both; // "goto" for simplicity and code size reduction.
				// Otherwise:
				// READ THE TEXT FROM THE REMOTE PROCESS
				// Although MSDN has the following comment about LVM_GETITEM, it is not present for
				// LVM_GETITEMTEXT. Therefore, to improve performance (by avoiding a second call to
				// ReadProcessMemory) and to reduce code size, we'll take them at their word until
				// proven otherwise.  Here is the MSDN comment about LVM_GETITEM: "Applications
				// should not assume that the text will necessarily be placed in the specified
				// buffer. The control may instead change the pszText member of the structure
				// to point to the new text, rather than place it in the buffer."
				if (ReadProcessMemory(handle, p_remote_text, contents, length * sizeof(TCHAR), NULL))
				{
					contents += length; // Point it to the position where the next char will be written.
					total_length += length; // Recalculate length in case its different than the estimate (for any reason).
				}
				//else it failed; but even so, continue on to put in a tab (if called for).
			}
			//else length is zero; but even so, continue on to put in a tab (if called for).
			if (single_col_mode)
				break;
		} // for() each column
	} // for() each row

break_both:
	if (contents) // Might be NULL if Assign() failed and thus var has zero capacity.
		*contents = '\0'; // Final termination.  Above has reserved room for this one byte.

	// CLEAN UP
	FreeInterProcMem(handle, p_remote_lvi);
	aOutputVar.Close(); // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_CLIPBOARD.
	aOutputVar.SetCharLength((VarSizeType)total_length); // Update to actual vs. estimated length.
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE);  // Indicate success.
}



ResultType Line::StatusBarGetText(LPTSTR aPart, LPTSTR aTitle, LPTSTR aText
	, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	// Note: ErrorLevel is handled by StatusBarUtil(), below.
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	HWND control_window = target_window ? ControlExist(target_window, _T("msctls_statusbar321")) : NULL;
	// Call this even if control_window is NULL because in that case, it will set the output var to
	// be blank for us:
	return StatusBarUtil(OUTPUT_VAR, control_window, ATOI(aPart)); // It will handle any zero part# for us.
}



ResultType Line::StatusBarWait(LPTSTR aTextToWaitFor, LPTSTR aSeconds, LPTSTR aPart, LPTSTR aTitle, LPTSTR aText
	, LPTSTR aInterval, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
// Since other script threads can interrupt this command while it's running, it's important that
// this command not refer to sArgDeref[] and sArgVar[] anytime after an interruption becomes possible.
// This is because an interrupting thread usually changes the values to something inappropriate for this thread.
{
	// Note: ErrorLevel is handled by StatusBarUtil(), below.
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	// Make a copy of any memory areas that are volatile (due to Deref buf being overwritten
	// if a new hotkey subroutine is launched while we are waiting) but whose contents we
	// need to refer to while we are waiting:
	TCHAR text_to_wait_for[4096];
	tcslcpy(text_to_wait_for, aTextToWaitFor, _countof(text_to_wait_for));
	HWND control_window = target_window ? ControlExist(target_window, _T("msctls_statusbar321")) : NULL;
	return StatusBarUtil(NULL, control_window, ATOI(aPart) // It will handle a NULL control_window or zero part# for us.
		, text_to_wait_for, *aSeconds ? (int)(ATOF(aSeconds)*1000) : -1 // Blank->indefinite.  0 means 500ms.
		, ATOI(aInterval));
}



ResultType Line::ScriptPostSendMessage(bool aUseSend)
// Arg list:
// sArgDeref[0]: Msg number
// sArgDeref[1]: wParam
// sArgDeref[2]: lParam
// sArgDeref[3]: Control
// sArgDeref[4]: WinTitle
// sArgDeref[5]: WinText
// sArgDeref[6]: ExcludeTitle
// sArgDeref[7]: ExcludeText
// sArgDeref[8]: Timeout
{
	HWND target_window, control_window;
	if (   !(target_window = DetermineTargetWindow(sArgDeref[4], sArgDeref[5], sArgDeref[6], sArgDeref[7]))
		|| !(control_window = *sArgDeref[3] ? ControlExist(target_window, sArgDeref[3]) : target_window)   ) // Relies on short-circuit boolean order.
		return SetErrorLevelOrThrowStr(aUseSend ? _T("FAIL") : ERRORLEVEL_ERROR); // Need a special value to distinguish this from numeric reply-values.

	// v1.0.40.05: Support the passing of a literal (quoted) string by checking whether the
	// original/raw arg's first character is '"'.  The avoids the need to put the string into a
	// variable and then pass something like &MyVar.
	UINT msg = ArgToUInt(1);
	WPARAM wparam = (mArgc > 1 && mArg[1].text[0] == '"') ? (WPARAM)sArgDeref[1] : (WPARAM)ArgToInt64(2);
	LPARAM lparam = (mArgc > 2 && mArg[2].text[0] == '"') ? (LPARAM)sArgDeref[2] : (LPARAM)ArgToInt64(3);
	// Timeout increased from 2000 to 5000 in v1.0.27:
	// jackieku: specify timeout by the parameter.
	UINT timeout = mArgc > 8 ? ArgToUInt(9) : 5000;

	// Fixed for v1.0.48.04: Make copies of the wParam and lParam variables (if eligible for updating) prior
	// to sending the message in case the message triggers a callback or OnMessage function, which would be
	// likely to change the contents of the mArg array before we're doing using them after the Post/SendMsg.
	// Seems best to do the above EVEN for PostMessage in case it can ever trigger a SendMessage internally
	// (I seem to remember that the OS sometimes converts a PostMessage call into a SendMessage if the
	// origin and destination are the same thread.)
	// v1.0.43.06: If either wParam or lParam contained the address of a variable, update the mLength
	// member after sending the message in case the receiver of the message wrote something to the buffer.
	// This is similar to the way "Str" parameters work in DllCall.
	Var *var_to_update[2];
	int i;
	for (i = 1; i < 3; ++i) // Two iterations: wParam and lParam.
	{
		if (mArgc > i) // The arg exists.
		{
			ArgStruct &this_arg = mArg[i];
			var_to_update[i-1] = this_arg.text[0] == '&'  // Must start with '&', so things like 5+&MyVar aren't supported.
				&& this_arg.deref && !this_arg.deref->is_function
				&& this_arg.deref->var->Type() == VAR_NORMAL // Check VAR_NORMAL to be extra-certain it can't be the clipboard or a built-in variable (ExpandExpression() probably prevents taking the address of such a variable, but might not stop it from being in the deref array that way).
				? this_arg.deref->var
				: NULL;
		}
		else // L32: Bugfix - var_to_update must be initialised.
			var_to_update[i-1] = NULL;
	}

	if (aUseSend)
	{
		DWORD_PTR dwResult;
		if (!SendMessageTimeout(control_window, msg, wparam, lparam, SMTO_ABORTIFHUNG, timeout, &dwResult))
			return SetErrorLevelOrThrowStr(_T("FAIL")); // Need a special value to distinguish this from numeric reply-values.
		g_ErrorLevel->Assign(dwResult); // UINT seems best most of the time?
	}
	else // Post vs. Send
	{
		if (!PostMessage(control_window, msg, wparam, lparam))
			return SetErrorLevelOrThrow();
		g_ErrorLevel->Assign(ERRORLEVEL_NONE);
	}

	for (i = 0; i < 2; ++i) // Two iterations: wParam and lParam.
		if (var_to_update[i])
			var_to_update[i]->SetLengthFromContents();

	return OK;
}



ResultType Line::ScriptProcess(LPTSTR aCmd, LPTSTR aProcess, LPTSTR aParam3)
{
	ProcessCmds process_cmd = ConvertProcessCmd(aCmd);
	// Runtime error is rare since it is caught at load-time unless it's in a var. ref.
	if (process_cmd == PROCESS_CMD_INVALID)
		return LineError(ERR_PARAM1_INVALID, FAIL, aCmd);

	HANDLE hProcess;
	DWORD pid, priority;
	BOOL result;

	switch (process_cmd)
	{
	case PROCESS_CMD_EXIST:
		return g_ErrorLevel->Assign(*aProcess ? ProcessExist(aProcess) : GetCurrentProcessId()); // The discovered PID or zero if none.

	case PROCESS_CMD_CLOSE:
		if (pid = ProcessExist(aProcess))  // Assign
		{
			if (hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, pid))
			{
				result = TerminateProcess(hProcess, 0);
				CloseHandle(hProcess);
				return g_ErrorLevel->Assign(result ? pid : 0); // Indicate success or failure.
			}
		}
		// Since above didn't return, yield a PID of 0 to indicate failure.
		return g_ErrorLevel->Assign(_T("0"));

	case PROCESS_CMD_PRIORITY:
		switch (_totupper(*aParam3))
		{
		case 'L': priority = IDLE_PRIORITY_CLASS; break;
		case 'B': priority = BELOW_NORMAL_PRIORITY_CLASS; break;
		case 'N': priority = NORMAL_PRIORITY_CLASS; break;
		case 'A': priority = ABOVE_NORMAL_PRIORITY_CLASS; break;
		case 'H': priority = HIGH_PRIORITY_CLASS; break;
		case 'R': priority = REALTIME_PRIORITY_CLASS; break;
		default:
			return g_ErrorLevel->Assign(_T("0"));  // 0 indicates failure in this case (i.e. a PID of zero).
		}
		if (pid = *aProcess ? ProcessExist(aProcess) : GetCurrentProcessId())  // Assign
		{
			if (hProcess = OpenProcess(PROCESS_SET_INFORMATION, FALSE, pid)) // Assign
			{
				// If OS doesn't support "above/below normal", seems best to default to normal rather than high/low,
				// since "above/below normal" aren't that dramatically different from normal:
				if (!g_os.IsWin2000orLater() && (priority == BELOW_NORMAL_PRIORITY_CLASS || priority == ABOVE_NORMAL_PRIORITY_CLASS))
					priority = NORMAL_PRIORITY_CLASS;
				result = SetPriorityClass(hProcess, priority);
				CloseHandle(hProcess);
				return g_ErrorLevel->Assign(result ? pid : 0); // Indicate success or failure.
			}
		}
		// Otherwise, return a PID of 0 to indicate failure.
		return g_ErrorLevel->Assign(_T("0"));

	case PROCESS_CMD_WAIT:
	case PROCESS_CMD_WAITCLOSE:
	{
		// This section is similar to that used for WINWAIT and RUNWAIT:
		bool wait_indefinitely;
		int sleep_duration;
		DWORD start_time;
		if (*aParam3) // the param containing the timeout value isn't blank.
		{
			wait_indefinitely = false;
			sleep_duration = (int)(ATOF(aParam3) * 1000); // Can be zero.
			start_time = GetTickCount();
		}
		else
		{
			wait_indefinitely = true;
			sleep_duration = 0; // Just to catch any bugs.
		}
		for (;;)
		{ // Always do the first iteration so that at least one check is done.
			pid = ProcessExist(aProcess);
			if (process_cmd == PROCESS_CMD_WAIT)
			{
				if (pid)
					return g_ErrorLevel->Assign(pid);
			}
			else // PROCESS_CMD_WAITCLOSE
			{
				// Since PID cannot always be determined (i.e. if process never existed, there was
				// no need to wait for it to close), for consistency, return 0 on success.
				if (!pid)
					return g_ErrorLevel->Assign(_T("0"));
			}
			// Must cast to int or any negative result will be lost due to DWORD type:
			if (wait_indefinitely || (int)(sleep_duration - (GetTickCount() - start_time)) > SLEEP_INTERVAL_HALF)
				MsgSleep(100);  // For performance reasons, don't check as often as the WinWait family does.
			else // Done waiting.
				return g_ErrorLevel->Assign(pid);
				// Above assigns 0 if "Process Wait" times out; or the PID of the process that still exists
				// if "Process WaitClose" times out.
		} // for()
	} // case
	} // switch()

	return FAIL;  // Should never be executed; just here to catch bugs.
}



ResultType Line::WinSetRegion(HWND aWnd, LPTSTR aPoints)
{
	if (!*aPoints) // Attempt to restore the window's normal/correct region.
	{
		// Fix for v1.0.31.07: The old method used the following, but apparently it's not the correct
		// way to restore a window's proper/normal region because when such a window is later maximized,
		// it retains its incorrect/smaller region:
		//if (GetWindowRect(aWnd, &rect))
		//{
		//	// Adjust the rect to keep the same size but have its upper-left corner at 0,0:
		//	rect.right -= rect.left;
		//	rect.bottom -= rect.top;
		//	rect.left = 0;
		//	rect.top = 0;
		//	if (hrgn = CreateRectRgnIndirect(&rect)) // Assign
		//	{
		//		// Presumably, the system deletes the former region when upon a successful call to SetWindowRgn().
		//		if (SetWindowRgn(aWnd, hrgn, TRUE))
		//			return g_ErrorLevel->Assign(ERRORLEVEL_NONE);
		//		// Otherwise, get rid of it since it didn't take effect:
		//		DeleteObject(hrgn);
		//	}
		//}
		//// Since above didn't return:
		//return OK; // Let ErrorLevel tell the story.

		// It's undocumented by MSDN, but apparently setting the Window's region to NULL restores it
		// to proper working order:
		return SetErrorLevelOrThrowBool(!SetWindowRgn(aWnd, NULL, TRUE)); // Let ErrorLevel tell the story.
	}

	#define MAX_REGION_POINTS 2000  // 2000 requires 16 KB of stack space.
	POINT pt[MAX_REGION_POINTS];
	int pt_count;
	LPTSTR cp;

	// Set defaults prior to parsing options in case any options are absent:
	int width = COORD_UNSPECIFIED;
	int height = COORD_UNSPECIFIED;
	int rr_width = COORD_UNSPECIFIED; // These two are for the rounded-rectangle method.
	int rr_height = COORD_UNSPECIFIED;
	bool use_ellipse = false;

	int fill_mode = ALTERNATE;
	// Concerning polygon regions: ALTERNATE is used by default (somewhat arbitrarily, but it seems to be the
	// more typical default).
	// MSDN: "In general, the modes [ALTERNATE vs. WINDING] differ only in cases where a complex,
	// overlapping polygon must be filled (for example, a five-sided polygon that forms a five-pointed
	// star with a pentagon in the center). In such cases, ALTERNATE mode fills every other enclosed
	// region within the polygon (that is, the points of the star), but WINDING mode fills all regions
	// (that is, the points and the pentagon)."

	for (pt_count = 0, cp = aPoints; *(cp = omit_leading_whitespace(cp));)
	{
		// To allow the MAX to be increased in the future with less chance of breaking existing scripts, consider this an error.
		if (pt_count >= MAX_REGION_POINTS)
			goto error;

		if (isdigit(*cp) || *cp == '-' || *cp == '+') // v1.0.38.02: Recognize leading minus/plus sign so that the X-coord is just as tolerant as the Y.
		{
			// Assume it's a pair of X/Y coordinates.  It's done this way rather than using X and Y
			// as option letters because:
			// 1) The script is more readable when there are multiple coordinates (for polygon).
			// 2) It enforces the fact that each X must have a Y and that X must always come before Y
			//    (which simplifies and reduces the size of the code).
			pt[pt_count].x = ATOI(cp);
			// For the delimiter, dash is more readable than pipe, even though it overlaps with "minus sign".
			// "x" is not used to avoid detecting "x" inside hex numbers.
			#define REGION_DELIMITER '-'
			if (   !(cp = _tcschr(cp + 1, REGION_DELIMITER))   ) // v1.0.38.02: cp + 1 to omit any leading minus sign.
				goto error;
			pt[pt_count].y = ATOI(++cp);  // Increment cp by only 1 to support negative Y-coord.
			++pt_count; // Move on to the next element of the pt array.
		}
		else
		{
			++cp;
			switch(_totupper(cp[-1]))
			{
			case 'E':
				use_ellipse = true;
				break;
			case 'R':
				if (!*cp || *cp == ' ') // Use 30x30 default.
				{
					rr_width = 30;
					rr_height = 30;
				}
				else
				{
					rr_width = ATOI(cp);
					if (cp = _tcschr(cp, REGION_DELIMITER)) // Assign
						rr_height = ATOI(++cp);
					else // Avoid problems with going beyond the end of the string.
						goto error;
				}
				break;
			case 'W':
				if (!_tcsnicmp(cp, _T("ind"), 3)) // [W]ind.
					fill_mode = WINDING;
				else
					width = ATOI(cp);
				break;
			case 'H':
				height = ATOI(cp);
				break;
			default: // For simplicity and to reserve other letters for future use, unknown options result in failure.
				goto error;
			} // switch()
		} // else

		if (   !(cp = _tcschr(cp, ' '))   ) // No more items.
			break;
	}

	if (!pt_count)
		goto error;

	bool width_and_height_were_both_specified = !(width == COORD_UNSPECIFIED || height == COORD_UNSPECIFIED);
	if (width_and_height_were_both_specified)
	{
		width += pt[0].x;   // Make width become the right side of the rect.
		height += pt[0].y;  // Make height become the bottom.
	}

	HRGN hrgn;
	if (use_ellipse) // Ellipse.
		hrgn = width_and_height_were_both_specified ? CreateEllipticRgn(pt[0].x, pt[0].y, width, height) : NULL;
	else if (rr_width != COORD_UNSPECIFIED) // Rounded rectangle.
		hrgn = width_and_height_were_both_specified ? CreateRoundRectRgn(pt[0].x, pt[0].y, width, height, rr_width, rr_height) : NULL;
	else if (width_and_height_were_both_specified) // Rectangle.
		hrgn = CreateRectRgn(pt[0].x, pt[0].y, width, height);
	else // Polygon
		hrgn = CreatePolygonRgn(pt, pt_count, fill_mode);
	if (!hrgn)
		goto error;
	// Since above didn't return, hrgn is now a non-NULL region ready to be assigned to the window.

	// Presumably, the system deletes the window's former region upon a successful call to SetWindowRgn():
	if (!SetWindowRgn(aWnd, hrgn, TRUE))
	{
		DeleteObject(hrgn);
		goto error;
	}
	//else don't delete hrgn since the system has taken ownership of it.

	// Since above didn't return, indicate success.
	return SetErrorLevelOrThrowBool(false);

error:
	return SetErrorLevelOrThrow();
}


						
ResultType Line::WinSet(LPTSTR aAttrib, LPTSTR aValue, LPTSTR aTitle, LPTSTR aText
	, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	WinSetAttributes attrib = ConvertWinSetAttribute(aAttrib);
	if (attrib == WINSET_INVALID)
		return LineError(ERR_PARAM1_INVALID, FAIL, aAttrib);

	// Only the following sub-commands affect ErrorLevel:
	bool use_errorlevel = (attrib == WINSET_STYLE || attrib == WINSET_EXSTYLE || attrib == WINSET_REGION);

	// Since this is a macro, avoid repeating it for every case of the switch():
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		goto error;

	int value;
	DWORD exstyle;

	switch (attrib)
	{
	case WINSET_ALWAYSONTOP:
	{
		HWND topmost_or_not;
		switch(ConvertOnOffToggle(aValue))
		{
		case TOGGLED_ON: topmost_or_not = HWND_TOPMOST; break;
		case TOGGLED_OFF: topmost_or_not = HWND_NOTOPMOST; break;
		case NEUTRAL: // parameter was blank, so it defaults to TOGGLE.
		case TOGGLE:
			exstyle = GetWindowLong(target_window, GWL_EXSTYLE);
			topmost_or_not = (exstyle & WS_EX_TOPMOST) ? HWND_NOTOPMOST : HWND_TOPMOST;
			break;
		default: return OK;
		}
		// SetWindowLong() didn't seem to work, at least not on some windows.  But this does.
		// As of v1.0.25.14, SWP_NOACTIVATE is also specified, though its absence does not actually
		// seem to activate the window, at least on XP (perhaps due to anti-focus-stealing measure
		// in Win98/2000 and beyond).  Or perhaps its something to do with the presence of
		// topmost_or_not (HWND_TOPMOST/HWND_NOTOPMOST), which might always avoid activating the
		// window.
		SetWindowPos(target_window, topmost_or_not, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);
		break;
	}

	// Note that WINSET_TOP is not offered as an option since testing reveals it has no effect on
	// top level (parent) windows, perhaps due to the anti focus-stealing measures in the OS.
	case WINSET_BOTTOM:
		// Note: SWP_NOACTIVATE must be specified otherwise the target window often/always fails to go
		// to the bottom:
		SetWindowPos(target_window, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);
		break;
	case WINSET_TOP:
		// Note: SWP_NOACTIVATE must be specified otherwise the target window often/always fails to go
		// to the bottom:
		SetWindowPos(target_window, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);
		break;

	case WINSET_TRANSPARENT:
	case WINSET_TRANSCOLOR:
	{
		// IMPORTANT (when considering future enhancements to these commands): Unlike
		// SetLayeredWindowAttributes(), which works on Windows 2000, GetLayeredWindowAttributes()
		// is supported only on XP or later.

		// It appears that turning on WS_EX_LAYERED in an attempt to retrieve the window's
		// former transparency setting does not work.  The OS probably does not store the
		// former transparency level (i.e. it forgets it the moment the WS_EX_LAYERED exstyle
		// is turned off).  This is true even if the following are done after the SetWindowLong():
		//MySetLayeredWindowAttributes(target_window, 0, 0, 0)
		// or:
		//if (MyGetLayeredWindowAttributes(target_window, &color, &alpha, &flags))
		//	MySetLayeredWindowAttributes(target_window, color, alpha, flags);
		// The above is why there is currently no "on" or "toggle" sub-command, just "Off".

		// Since the color of an HBRUSH can't be easily determined (since it can be a pattern and
		// since there seem to be no easy API calls to discover the colors of pixels in an HBRUSH),
		// the following is not yet implemented: Use window's own class background color (via
		// GetClassLong) if aValue is entirely blank.

		exstyle = GetWindowLong(target_window, GWL_EXSTYLE);
		if (!_tcsicmp(aValue, _T("Off")))
			// One user reported that turning off the attribute helps window's scrolling performance.
			SetWindowLong(target_window, GWL_EXSTYLE, exstyle & ~WS_EX_LAYERED);
		else
		{
			if (attrib == WINSET_TRANSPARENT)
			{
				// Update to the below for v1.0.23: WS_EX_LAYERED can now be removed via the above:
				// NOTE: It seems best never to remove the WS_EX_LAYERED attribute, even if the value is 255
				// (non-transparent), since the window might have had that attribute previously and may need
				// it to function properly.  For example, an app may support making its own windows transparent
				// but might not expect to have to turn WS_EX_LAYERED back on if we turned it off.  One drawback
				// of this is a quote from somewhere that might or might not be accurate: "To make this window
				// completely opaque again, remove the WS_EX_LAYERED bit by calling SetWindowLong and then ask
				// the window to repaint. Removing the bit is desired to let the system know that it can free up
				// some memory associated with layering and redirection."
				value = ATOI(aValue);
				// A little debatable, but this behavior seems best, at least in some cases:
				if (value < 0)
					value = 0;
				else if (value > 255)
					value = 255;
				SetWindowLong(target_window, GWL_EXSTYLE, exstyle | WS_EX_LAYERED);
				SetLayeredWindowAttributes(target_window, 0, value, LWA_ALPHA);
			}
			else // attrib == WINSET_TRANSCOLOR
			{
				// The reason WINSET_TRANSCOLOR accepts both the color and an optional transparency settings
				// is that calling SetLayeredWindowAttributes() with only the LWA_COLORKEY flag causes the
				// window to lose its current transparency setting in favor of the transparent color.  This
				// is true even though the LWA_ALPHA flag was not specified, which seems odd and is a little
				// disappointing, but that's the way it is on XP at least.
				TCHAR aValue_copy[256];
				tcslcpy(aValue_copy, aValue, _countof(aValue_copy)); // Make a modifiable copy.
				LPTSTR space_pos = StrChrAny(aValue_copy, _T(" \t")); // Space or tab.
				if (space_pos)
				{
					*space_pos = '\0';
					++space_pos;  // Point it to the second substring.
				}
				COLORREF color = ColorNameToBGR(aValue_copy);
				if (color == CLR_NONE) // A matching color name was not found, so assume it's in hex format.
					// It seems _tcstol() automatically handles the optional leading "0x" if present:
					color = rgb_to_bgr(_tcstol(aValue_copy, NULL, 16));
				DWORD flags;
				if (   space_pos && *(space_pos = omit_leading_whitespace(space_pos))   ) // Relies on short-circuit boolean.
				{
					value = ATOI(space_pos);  // To keep it simple, don't bother with 0 to 255 range validation in this case.
					flags = LWA_COLORKEY|LWA_ALPHA;  // i.e. set both the trans-color and the transparency level.
				}
				else // No translucency value is present, only a trans-color.
				{
					value = 0;  // Init to avoid possible compiler warning.
					flags = LWA_COLORKEY;
				}
				SetWindowLong(target_window, GWL_EXSTYLE, exstyle | WS_EX_LAYERED);
				SetLayeredWindowAttributes(target_window, color, value, flags);
			}
		}
		break;
	}

	case WINSET_STYLE:
	case WINSET_EXSTYLE:
	{
		if (!*aValue)
			goto error; // Seems best not to treat an explicit blank as zero.
		int style_index = (attrib == WINSET_STYLE) ? GWL_STYLE : GWL_EXSTYLE;
		DWORD new_style, orig_style = GetWindowLong(target_window, style_index);
		if (!_tcschr(_T("+-^"), *aValue))
			new_style = ATOU(aValue); // No prefix, so this new style will entirely replace the current style.
		else
		{
			++aValue; // Won't work combined with next line, due to next line being a macro that uses the arg twice.
			DWORD style_change = ATOU(aValue);
			// +/-/^ are used instead of |&^ because the latter is confusing, namely that
			// "&" really means &=~style, etc.
			switch(aValue[-1])
			{
			case '+': new_style = orig_style | style_change; break;
			case '-': new_style = orig_style & ~style_change; break;
			case '^': new_style = orig_style ^ style_change; break;
			}
		}
		SetLastError(0); // Prior to SetWindowLong(), as recommended by MSDN.
		if (SetWindowLong(target_window, style_index, new_style) || !GetLastError()) // This is the precise way to detect success according to MSDN.
		{
			// Even if it indicated success, sometimes it failed anyway.  Find out for sure:
			if (GetWindowLong(target_window, style_index) != orig_style) // Even a partial change counts as a success.
			{
				// SetWindowPos is also necessary, otherwise the frame thickness entirely around the window
				// does not get updated (just parts of it):
				SetWindowPos(target_window, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
				// Since SetWindowPos() probably doesn't know that the style changed, below is probably necessary
				// too, at least in some cases:
				InvalidateRect(target_window, NULL, TRUE); // Quite a few styles require this to become visibly manifest.
				break; // Success. See below.
			}
		}
		goto error; // Since above didn't break, it's a failure.
	}

	case WINSET_ENABLE:
	case WINSET_DISABLE: // These are separate sub-commands from WINSET_STYLE because merely changing the WS_DISABLED style is usually not as effective as calling EnableWindow().
		EnableWindow(target_window, attrib == WINSET_ENABLE);
		return OK;

	case WINSET_REGION:
		return WinSetRegion(target_window, aValue);

	case WINSET_REDRAW:
		// Seems best to always have the last param be TRUE, for now, so that aValue can be
		// reserved for future use such as invalidating only part of a window, etc. Also, it
		// seems best not to call UpdateWindow(), which forces the window to immediately
		// process a WM_PAINT message, since that might not be desirable as a default (maybe
		// an option someday).  Other future options might include alternate methods of
		// getting a window to redraw, such as:
		// SendMessage(mHwnd, WM_NCPAINT, 1, 0);
		// RedrawWindow(mHwnd, NULL, NULL, RDW_INVALIDATE|RDW_FRAME|RDW_UPDATENOW);
		// SetWindowPos(mHwnd, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
		// GetClientRect(mControl[mDefaultButtonIndex].hwnd, &client_rect);
		// InvalidateRect(mControl[mDefaultButtonIndex].hwnd, &client_rect, TRUE);
		InvalidateRect(target_window, NULL, TRUE);
		break;

	} // switch()
	return use_errorlevel ? SetErrorLevelOrThrowBool(false) : OK;

error:
	// Only STYLE, EXSTYLE and REDRAW affect ErrorLevel for compatibility reasons,
	// but seems best to allow the other sub-commands to throw exceptions:
	return (use_errorlevel || g->InTryBlock()) ? SetErrorLevelOrThrow() : OK;
}



ResultType Line::WinSetTitle(LPTSTR aTitle, LPTSTR aText, LPTSTR aNewTitle, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
// Like AutoIt2, this function and others like it always return OK, even if the target window doesn't
// exist or there action doesn't actually succeed.
{
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		return OK;
	SetWindowText(target_window, aNewTitle);
	return OK;
}



ResultType Line::WinGetTitle(LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	Var &output_var = *OUTPUT_VAR; // Must be resolved only once and prior to DetermineTargetWindow().  See Line::WinGetClass() for explanation.
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	// Even if target_window is NULL, we want to continue on so that the output
	// param is set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before.

	// Handle the output parameter.  See the comments in ACT_CONTROLGETTEXT for details.
	int space_needed = target_window ? GetWindowTextLength(target_window) + 1 : 1; // 1 for terminator.
	if (output_var.AssignString(NULL, space_needed - 1) != OK)
		return FAIL;  // It already displayed the error.
	if (target_window)
	{
		// Update length using the actual length, rather than the estimate provided by GetWindowTextLength():
		output_var.SetCharLength((VarSizeType)GetWindowText(target_window, output_var.Contents(), space_needed));
		if (!output_var.Length())
			// There was no text to get or GetWindowTextTimeout() failed.
			*output_var.Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
	}
	else
	{
		*output_var.Contents() = '\0';
		output_var.SetCharLength(0);
	}
	return output_var.Close(); // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_CLIPBOARD.
}



ResultType Line::WinGetClass(LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	Var &output_var = *OUTPUT_VAR; // Fix for v1.0.48: Must be resolved only once and prior to DetermineTargetWindow() due to the following from Lexikos:
	// WinGetClass causes an access violation if one of the script's windows is sub-classed by the script [unless the above is done].
	// This occurs because WM_GETTEXT is sent to the GUI, triggering the window procedure. The script callback
	// then executes and invalidates sArgVar[0], which WinGetClass attempts to dereference. 
	// (Thanks to TodWulff for bringing this issue to my attention.) 
	// Solution: WinGetTitle resolves the OUTPUT_VAR (*sArgVar) macro once, before searching for the window.
	// I suggest the same be done for WinGetClass.
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		return output_var.Assign();
	TCHAR class_name[WINDOW_CLASS_SIZE];
	if (!GetClassName(target_window, class_name, _countof(class_name)))
		return output_var.Assign();
	return output_var.Assign(class_name);
}



ResultType WinGetList(Var &aOutputVar, WinGetCmds aCmd, LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
// Helper function for WinGet() to avoid having a WindowSearch object on its stack (since that object
// normally isn't needed).
{
	WindowSearch ws;
	ws.mFindLastMatch = true; // Must set mFindLastMatch to get all matches rather than just the first.
	ws.mArrayStart = (aCmd == WINGET_CMD_LIST) ? &aOutputVar : NULL; // Provide the position in the var list of where the array-element vars will be.
	// If aTitle is ahk_id nnnn, the Enum() below will be inefficient.  However, ahk_id is almost unheard of
	// in this context because it makes little sense, so no extra code is added to make that case efficient.
	if (ws.SetCriteria(*g, aTitle, aText, aExcludeTitle, aExcludeText)) // These criteria allow the possibility of a match.
		EnumWindows(EnumParentFind, (LPARAM)&ws);
	//else leave ws.mFoundCount set to zero (by the constructor).
	return aOutputVar.Assign(ws.mFoundCount);
}



ResultType Line::WinGet(LPTSTR aCmd, LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	Var &output_var = *OUTPUT_VAR;  // This is done even for WINGET_CMD_LIST.
	WinGetCmds cmd = ConvertWinGetCmd(aCmd);
	// Since command names are validated at load-time, this only happens if the command name
	// was contained in a variable reference.  But for simplicity of design here, return
	// failure in this case (unlike other functions similar to this one):
	if (cmd == WINGET_CMD_INVALID)
		return LineError(ERR_PARAM2_INVALID, FAIL, aCmd);

	bool target_window_determined = true;  // Set default.
	HWND target_window;
	IF_USE_FOREGROUND_WINDOW(g->DetectHiddenWindows, aTitle, aText, aExcludeTitle, aExcludeText)
	else if (!(*aTitle || *aText || *aExcludeTitle || *aExcludeText)
		&& !(cmd == WINGET_CMD_LIST || cmd == WINGET_CMD_COUNT)) // v1.0.30.02/v1.0.30.03: Have "list"/"count" get all windows on the system when there are no parameters.
		target_window = GetValidLastUsedWindow(*g);
	else
		target_window_determined = false;  // A different method is required.

	// Used with WINGET_CMD_LIST to create an array (if needed).  Make it longer than Max var name
	// so that FindOrAddVar() will be able to spot and report var names that are too long:
	TCHAR var_name[MAX_VAR_NAME_LENGTH + 20], buf[32];
	Var *array_item;

	switch(cmd)
	{
	case WINGET_CMD_ID:
	case WINGET_CMD_IDLAST:
		if (!target_window_determined)
			target_window = WinExist(*g, aTitle, aText, aExcludeTitle, aExcludeText, cmd == WINGET_CMD_IDLAST);
		if (target_window)
			return output_var.AssignHWND(target_window);
		else
			return output_var.Assign();

	case WINGET_CMD_PID:
	case WINGET_CMD_PROCESSNAME:
	case WINGET_CMD_PROCESSPATH:
		if (!target_window_determined)
			target_window = WinExist(*g, aTitle, aText, aExcludeTitle, aExcludeText);
		if (target_window)
		{
			DWORD pid;
			GetWindowThreadProcessId(target_window, &pid);
			if (cmd == WINGET_CMD_PID)
				return output_var.Assign(pid);
			// Otherwise, get the full path and name of the executable that owns this window.
			TCHAR process_name[MAX_PATH];
			GetProcessName(pid, process_name, _countof(process_name), cmd == WINGET_CMD_PROCESSNAME);
			return output_var.Assign(process_name);
		}
		// If above didn't return:
		return output_var.Assign();

	case WINGET_CMD_COUNT:
	case WINGET_CMD_LIST:
		// LIST retrieves a list of HWNDs for the windows that match the given criteria and stores them in
		// an array.  The number of items in the array is stored in the base array name (unlike
		// StringSplit, which stores them in array element #0).  This is done for performance reasons
		// (so that element #0 doesn't have to be looked up at runtime), but mostly because of the
		// complexity of resolving a parameter than can be either an output-var or an array name at
		// load-time -- namely that if param #1 were allowed to be an array name, there is ambiguity
		// about where the name of the array is actually stored depending on whether param#1 was literal
		// text or a deref.  So it's easier and performs better just to do it this way, even though it
		// breaks from the StringSplit tradition:
		if (target_window_determined)
		{
			if (!target_window)
				return output_var.Assign(_T("0")); // 0 windows found
			if (cmd == WINGET_CMD_LIST)
			{
				// Otherwise, since the target window has been determined, we know that it is
				// the only window to be put into the array:
				if (   !(array_item = g_script.FindOrAddVar(var_name
					, sntprintf(var_name, _countof(var_name), _T("%s1"), output_var.mName)
					, FINDVAR_FOR_PSEUDO_ARRAY(output_var)))   )  // Find or create element #1.

					return FAIL;  // It will have already displayed the error.
				if (!array_item->AssignHWND(target_window))
					return FAIL;
			}
			return output_var.Assign(_T("1"));  // 1 window found
		}
		// Otherwise, the target window(s) have not yet been determined and a special method
		// is required to gather them.
		return WinGetList(output_var, cmd, aTitle, aText, aExcludeTitle, aExcludeText); // Outsourced to avoid having a WindowSearch object on this function's stack.

	case WINGET_CMD_MINMAX:
		if (!target_window_determined)
			target_window = WinExist(*g, aTitle, aText, aExcludeTitle, aExcludeText);
		// Testing shows that it's not possible for a minimized window to also be maximized (under
		// the theory that upon restoration, it *would* be maximized.  This is unfortunate if there
		// is no other way to determine what the restoration size and maximized state will be for a
		// minimized window.
		if (target_window)
			return output_var.Assign(IsZoomed(target_window) ? 1 : (IsIconic(target_window) ? -1 : 0));
		else
			return output_var.Assign();

	case WINGET_CMD_CONTROLLIST:
	case WINGET_CMD_CONTROLLISTHWND:
		if (!target_window_determined)
			target_window = WinExist(*g, aTitle, aText, aExcludeTitle, aExcludeText);
		return target_window ? WinGetControlList(output_var, target_window, cmd == WINGET_CMD_CONTROLLISTHWND)
			: output_var.Assign();

	case WINGET_CMD_STYLE:
	case WINGET_CMD_EXSTYLE:
		if (!target_window_determined)
			target_window = WinExist(*g, aTitle, aText, aExcludeTitle, aExcludeText);
		if (!target_window)
			return output_var.Assign();
		_stprintf(buf, _T("0x%08X"), GetWindowLong(target_window, cmd == WINGET_CMD_STYLE ? GWL_STYLE : GWL_EXSTYLE));
		return output_var.Assign(buf);

	case WINGET_CMD_TRANSPARENT:
	case WINGET_CMD_TRANSCOLOR:
		if (!target_window_determined)
			target_window = WinExist(*g, aTitle, aText, aExcludeTitle, aExcludeText);
		if (!target_window)
			return output_var.Assign();
		typedef BOOL (WINAPI *MyGetLayeredWindowAttributesType)(HWND, COLORREF*, BYTE*, DWORD*);
		static MyGetLayeredWindowAttributesType MyGetLayeredWindowAttributes = (MyGetLayeredWindowAttributesType)
			GetProcAddress(GetModuleHandle(_T("user32")), "GetLayeredWindowAttributes");
		COLORREF color;
		BYTE alpha;
		DWORD flags;
		// IMPORTANT (when considering future enhancements to these commands): Unlike
		// SetLayeredWindowAttributes(), which works on Windows 2000, GetLayeredWindowAttributes()
		// is supported only on XP or later.
		if (!MyGetLayeredWindowAttributes || !(MyGetLayeredWindowAttributes(target_window, &color, &alpha, &flags)))
			return output_var.Assign();
		if (cmd == WINGET_CMD_TRANSPARENT)
			return (flags & LWA_ALPHA) ? output_var.Assign((DWORD)alpha) : output_var.Assign();
		else // WINGET_CMD_TRANSCOLOR
		{
			if (flags & LWA_COLORKEY)
			{
				// Store in hex format to aid in debugging scripts.  Also, the color is always
				// stored in RGB format, since that's what WinSet uses:
				_stprintf(buf, _T("0x%06X"), bgr_to_rgb(color));
				return output_var.Assign(buf);
			}
			else // This window does not have a transparent color (or it's not accessible to us, perhaps for reasons described at MSDN GetLayeredWindowAttributes()).
				return output_var.Assign();
		}
	}

	return FAIL;  // Never executed (increases maintainability and avoids compiler warning).
}



ResultType Line::WinGetControlList(Var &aOutputVar, HWND aTargetWindow, bool aFetchHWNDs)
// Caller must ensure that aTargetWindow is non-NULL and valid.
// Every control is fetched rather than just a list of distinct class names (possibly with a
// second script array containing the quantity of each class) because it's conceivable that the
// z-order of the controls will be useful information to some script authors.
// A delimited list is used rather than the array technique used by "WinGet, OutputVar, List" because:
// 1) It allows the flexibility of searching the list more easily with something like IfInString.
// 2) It seems rather rare that the count of items in the list would be useful info to a script author
//    (the count can be derived with a parsing loop if it's ever needed).
// 3) It saves memory since script arrays are permanent and since each array element would incur
//    the overhead of being a script variable, not to mention that each variable has a minimum
//    capacity (additional overhead) of 64 bytes.
{
	control_list_type cl; // A big struct containing room to store class names and counts for each.
	CL_INIT_CONTROL_LIST(cl)
	cl.fetch_hwnds = aFetchHWNDs;
	cl.target_buf = NULL;  // First pass: Signal it not not to write to the buf, but instead only calculate the length.
	EnumChildWindows(aTargetWindow, EnumChildGetControlList, (LPARAM)&cl);
	if (!cl.total_length) // No controls in the window.
		return aOutputVar.Assign();
	// This adjustment was added because someone reported that max variable capacity was being
	// exceeded in some cases (perhaps custom controls that retrieve large amounts of text
	// from the disk in response to the "get text" message):
	if (cl.total_length >= g_MaxVarCapacity) // Allow the command to succeed by truncating the text.
		cl.total_length = g_MaxVarCapacity - 1;
	// Set up the var, enlarging it if necessary.  If the aOutputVar is of type VAR_CLIPBOARD,
	// this call will set up the clipboard for writing:
	if (aOutputVar.AssignString(NULL, (VarSizeType)cl.total_length) != OK)
		return FAIL;  // It already displayed the error.
	// Fetch the text directly into the var.  Also set the length explicitly
	// in case actual size written was off from the estimated size (in case the list of
	// controls changed in the split second between pass #1 and pass #2):
	CL_INIT_CONTROL_LIST(cl)
	cl.target_buf = aOutputVar.Contents();  // Second pass: Write to the buffer.
	cl.capacity = aOutputVar.Capacity(); // Because granted capacity might be a little larger than we asked for.
	EnumChildWindows(aTargetWindow, EnumChildGetControlList, (LPARAM)&cl);
	aOutputVar.SetCharLength((VarSizeType)cl.total_length);  // In case it wound up being smaller than expected.
	if (!cl.total_length) // Something went wrong, so make sure its terminated just in case.
		*aOutputVar.Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
	return aOutputVar.Close(); // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_CLIPBOARD.
}



BOOL CALLBACK EnumChildGetControlList(HWND aWnd, LPARAM lParam)
{
	control_list_type &cl = *(control_list_type *)lParam;  // For performance and convenience.
	TCHAR line[WINDOW_CLASS_SIZE + 5];  // +5 to allow room for the sequence number to be appended later below.
	int line_length;

	// cl.fetch_hwnds==true is a new mode in v1.0.43.06+ to help performance of AHK Window Info and other
	// scripts that want to operate directly on the HWNDs.
	if (cl.fetch_hwnds)
	{
		line_length = (int)_tcslen(HwndToString(aWnd, line));
	}
	else // The mode that fetches ClassNN vs. HWND.
	{
		// Note: IsWindowVisible(aWnd) is not checked because although Window Spy does not reveal
		// hidden controls if the mouse happens to be hovering over one, it does include them in its
		// sequence numbering (which is a relieve, since results are probably much more consistent
		// then, esp. for apps that hide and unhide controls in response to actions on other controls).
		if (  !(line_length = GetClassName(aWnd, line, WINDOW_CLASS_SIZE))   ) // Don't include the +5 extra size since that is reserved for seq. number.
			return TRUE; // Probably very rare. Continue enumeration since Window Spy doesn't even check for failure.
		// It has been verified that GetClassName()'s returned length does not count the terminator.

		// Check if this class already exists in the class array:
		int class_index;
		for (class_index = 0; class_index < cl.total_classes; ++class_index)
			if (!_tcsicmp(cl.class_name[class_index], line)) // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales.
				break;
		if (class_index < cl.total_classes) // Match found.
		{
			++cl.class_count[class_index]; // Increment the number of controls of this class that have been found so far.
			if (cl.class_count[class_index] > 99999) // Sanity check; prevents buffer overflow or number truncation in "line".
				return TRUE;  // Continue the enumeration.
		}
		else // No match found, so create new entry if there's room.
		{
			if (cl.total_classes == CL_MAX_CLASSES // No pointers left.
				|| CL_CLASS_BUF_SIZE - (cl.buf_free_spot - cl.class_buf) - 1 < line_length) // Insuff. room in buf.
				return TRUE; // Very rare. Continue the enumeration so that class names already found can be collected.
			// Otherwise:
			cl.class_name[class_index] = cl.buf_free_spot;  // Set this pointer to its place in the buffer.
			_tcscpy(cl.class_name[class_index], line); // Copy the string into this place.
			cl.buf_free_spot += line_length + 1;  // +1 because every string in the buf needs its own terminator.
			cl.class_count[class_index] = 1;  // Indicate that the quantity of this class so far is 1.
			++cl.total_classes;
		}

		_itot(cl.class_count[class_index], line + line_length, 10); // Append the seq. number to line.
		line_length = (int)_tcslen(line);  // Update the length.
	}

	int extra_length;
	if (cl.is_first_iteration)
	{
		extra_length = 0; // All items except the first are preceded by a delimiting LF.
		cl.is_first_iteration = false;
	}
	else
		extra_length = 1;

	if (cl.target_buf)
	{
		if ((int)(cl.capacity - cl.total_length - extra_length - 1) < line_length)
			// No room in target_buf (i.e. don't write a partial item to the buffer).
			return TRUE;  // Rare: it should only happen if size in pass #2 differed from that calc'd in pass #1.
		if (extra_length)
		{
			cl.target_buf[cl.total_length] = '\n'; // Replace previous item's terminator with newline.
			cl.total_length += extra_length;
		}
		_tcscpy(cl.target_buf + cl.total_length, line); // Write hwnd or class name+seq. number.
		cl.total_length += line_length;
	}
	else // Caller only wanted the total length calculated.
		cl.total_length += line_length + extra_length;

	return TRUE; // Continue enumeration through all the windows.
}



ResultType Line::WinGetText(LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	Var &output_var = *OUTPUT_VAR;
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	// Even if target_window is NULL, we want to continue on so that the output
	// variables are set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before:
	if (!target_window)
	{
		if (output_var.Assign() != OK) // Tell it not to free the memory by omitting all params.
			return FAIL;
		return SetErrorLevelOrThrow();
	}

	length_and_buf_type sab;
	sab.buf = NULL; // Tell it just to calculate the length this time around.
	sab.total_length = 0; // Init
	sab.capacity = 0;     //
	EnumChildWindows(target_window, EnumChildGetText, (LPARAM)&sab);

	if (!sab.total_length) // No text in window.
	{
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		return output_var.Assign(); // Tell it not to free the memory by omitting all params.
	}
	// This adjustment was added because someone reported that max variable capacity was being
	// exceeded in some cases (perhaps custom controls that retrieve large amounts of text
	// from the disk in response to the "get text" message):
	if (sab.total_length >= g_MaxVarCapacity)    // Allow the command to succeed by truncating the text.
		sab.total_length = g_MaxVarCapacity - 1; // And this length will be used to limit the retrieval capacity below.

	// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
	// this call will set up the clipboard for writing:
	if (output_var.AssignString(NULL, (VarSizeType)sab.total_length) != OK)
		return FAIL;  // It already displayed the error.

	// Fetch the text directly into the var.  Also set the length explicitly
	// in case actual size written was different than the estimated size (since
	// GetWindowTextLength() can return more space that will actually be required
	// in certain circumstances, see MSDN):
	sab.buf = output_var.Contents();
	sab.total_length = 0; // Init
	// Note: The capacity member below exists because granted capacity might be a little larger than we asked for,
	// which allows the actual text fetched to be larger than the length estimate retrieved by the first pass
	// (which generally shouldn't happen since MSDN docs say that the actual length can be less, but never greater,
	// than the estimate length):
	sab.capacity = output_var.Capacity(); // Capacity includes the zero terminator, i.e. it's the size of the memory area.
	EnumChildWindows(target_window, EnumChildGetText, (LPARAM)&sab); // Get the text.

	// Length is set explicitly below in case it wound up being smaller than expected/estimated.
	// MSDN says that can happen generally, and also specifically because: "ANSI applications may have
	// the string in the buffer reduced in size (to a minimum of half that of the wParam value) due to
	// conversion from ANSI to Unicode."
	output_var.SetCharLength((VarSizeType)sab.total_length);
	if (!sab.total_length)
		// Something went wrong, so make sure we set to empty string.
		*sab.buf = '\0';  // Safe because Assign() gave us a non-constant memory area.
	if (output_var.Close() != OK) // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_CLIPBOARD.
		return FAIL;
	return SetErrorLevelOrThrowBool(!sab.total_length);
}



BOOL CALLBACK EnumChildGetText(HWND aWnd, LPARAM lParam)
{
	if (!g->DetectHiddenText && !IsWindowVisible(aWnd))
		return TRUE;  // This child/control is hidden and user doesn't want it considered, so skip it.
	length_and_buf_type &lab = *(length_and_buf_type *)lParam;  // For performance and convenience.
	int length;
	if (lab.buf)
		length = GetWindowTextTimeout(aWnd, lab.buf + lab.total_length
			, (int)(lab.capacity - lab.total_length)); // Not +1.  Verified correct because WM_GETTEXT accepts size of buffer, not its length.
	else
		length = GetWindowTextTimeout(aWnd);
	lab.total_length += length;
	if (length)
	{
		if (lab.buf)
		{
			if (lab.capacity - lab.total_length > 2) // Must be >2 due to zero terminator.
			{
				_tcscpy(lab.buf + lab.total_length, _T("\r\n")); // Something to delimit each control's text.
				lab.total_length += 2;
			}
			// else don't increment total_length
		}
		else
			lab.total_length += 2; // Since buf is NULL, accumulate the size that *would* be needed.
	}
	return TRUE; // Continue enumeration through all the child windows of this parent.
}



ResultType Line::WinGetPos(LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	Var *output_var_x = ARGVAR1;  // Ok if NULL. Load-time validation has ensured that these are valid output variables (e.g. not built-in vars).
	Var *output_var_y = ARGVAR2;  // Ok if NULL.
	Var *output_var_width = ARGVAR3;  // Ok if NULL.
	Var *output_var_height = ARGVAR4;  // Ok if NULL.

	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	// Even if target_window is NULL, we want to continue on so that the output
	// variables are set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before.
	RECT rect;
	if (target_window)
		GetWindowRect(target_window, &rect);
	else // ensure it's initialized for possible later use:
		rect.bottom = rect.left = rect.right = rect.top = 0;

	ResultType result = OK; // Set default;

	if (output_var_x)
		if (target_window)
		{
			if (!output_var_x->Assign(rect.left))  // X position
				result = FAIL;
		}
		else
			if (!output_var_x->Assign(_T("")))
				result = FAIL;
	if (output_var_y)
		if (target_window)
		{
			if (!output_var_y->Assign(rect.top))  // Y position
				result = FAIL;
		}
		else
			if (!output_var_y->Assign(_T("")))
				result = FAIL;
	if (output_var_width) // else user didn't want this value saved to an output param
		if (target_window)
		{
			if (!output_var_width->Assign(rect.right - rect.left))  // Width
				result = FAIL;
		}
		else
			if (!output_var_width->Assign(_T(""))) // Set it to be empty to signal the user that the window wasn't found.
				result = FAIL;
	if (output_var_height)
		if (target_window)
		{
			if (!output_var_height->Assign(rect.bottom - rect.top))  // Height
				result = FAIL;
		}
		else
			if (!output_var_height->Assign(_T("")))
				result = FAIL;

	return result;
}



ResultType Line::EnvGet(LPTSTR aEnvVarName)
{
	Var *output_var = OUTPUT_VAR;
	// Don't use a size greater than 32767 because that will cause it to fail on Win95 (tested by Robert Yalkin).
	// According to MSDN, 32767 is exactly large enough to handle the largest variable plus its zero terminator.
	// Update: In practice, at least on Windows 7, the limit only applies to the ANSI functions.
	TCHAR buf[32767];
	// GetEnvironmentVariable() could be called twice, the first time to get the actual size.  But that would
	// probably perform worse since GetEnvironmentVariable() is a very slow function, so it seems best to fetch
	// it into a large buffer then just copy it to dest-var.
	DWORD length = GetEnvironmentVariable(aEnvVarName, buf, _countof(buf));
	if (length >= _countof(buf))
	{
		// In this case, length indicates the required buffer size, and the contents of the buffer are undefined.
		// Since our buffer is 32767 characters, the var apparently exceeds the documented limit, as can happen
		// if the var was set with the Unicode API.
		if (!output_var->AssignString(NULL, length - 1, true))
			return FAIL;
		length = GetEnvironmentVariable(aEnvVarName, output_var->Contents(), length);
		if (!length)
			*output_var->Contents() = '\0'; // Ensure var is null-terminated.
		return output_var->Close();
	}
	return output_var->Assign(length ? buf : _T(""), length);
}



ResultType Line::SysGet(LPTSTR aCmd, LPTSTR aValue)
// Thanks to Gregory F. Hogg of Hogg's Software for providing sample code on which this function
// is based.
{
	// For simplicity and array look-up performance, this is done even for sub-commands that output to an array:
	Var &output_var = *OUTPUT_VAR;
	SysGetCmds cmd = ConvertSysGetCmd(aCmd);
	// Since command names are validated at load-time, this only happens if the command name
	// was contained in a variable reference.  But for simplicity of design here, return
	// failure in this case (unlike other functions similar to this one):
	if (cmd == SYSGET_CMD_INVALID)
		return LineError(ERR_PARAM2_INVALID, FAIL, aCmd);

	MonitorInfoPackage mip = {0};  // Improves maintainability to initialize unconditionally, here.
	mip.monitor_info_ex.cbSize = sizeof(MONITORINFOEX); // Also improves maintainability.

	// EnumDisplayMonitors() must be dynamically loaded; otherwise, the app won't launch at all on Win95/NT.
	typedef BOOL (WINAPI* EnumDisplayMonitorsType)(HDC, LPCRECT, MONITORENUMPROC, LPARAM);
	static EnumDisplayMonitorsType MyEnumDisplayMonitors = (EnumDisplayMonitorsType)
		GetProcAddress(GetModuleHandle(_T("user32")), "EnumDisplayMonitors");

	switch(cmd)
	{
	case SYSGET_CMD_METRICS: // In this case, aCmd is the value itself.
		return output_var.Assign(GetSystemMetrics(ATOI(aCmd)));  // Input and output are both signed integers.

	// For the next few cases, I'm not sure if it is possible to have zero monitors.  Obviously it's possible
	// to not have a monitor turned on or not connected at all.  But it seems likely that these various API
	// functions will provide a "default monitor" in the absence of a physical monitor connected to the
	// system.  To be safe, all of the below will assume that zero is possible, at least on some OSes or
	// under some conditions.  However, on Win95/NT, "1" is assumed since there is probably no way to tell
	// for sure if there are zero monitors except via GetSystemMetrics(SM_CMONITORS), which is a different
	// animal as described below.
	case SYSGET_CMD_MONITORCOUNT:
		// Don't use GetSystemMetrics(SM_CMONITORS) because of this:
		// MSDN: "GetSystemMetrics(SM_CMONITORS) counts only display monitors. This is different from
		// EnumDisplayMonitors, which enumerates display monitors and also non-display pseudo-monitors."
		if (!MyEnumDisplayMonitors) // Since system only supports 1 monitor, the first must be primary.
			return output_var.Assign(1); // Assign as 1 vs. "1" to use hexadecimal display if that is in effect.
		mip.monitor_number_to_find = COUNT_ALL_MONITORS;
		MyEnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
		return output_var.Assign(mip.count); // Will assign zero if the API ever returns a legitimate zero.

	// Even if the first monitor to be retrieved by the EnumProc is always the primary (which is doubtful
	// since there's no mention of this in the MSDN docs) it seems best to have this sub-cmd in case that
	// policy ever changes:
	case SYSGET_CMD_MONITORPRIMARY:
		if (!MyEnumDisplayMonitors) // Since system only supports 1 monitor, the first must be primary.
			return output_var.Assign(1); // Assign as 1 vs. "1" to use hexadecimal display if that is in effect.
		// The mip struct's values have already initialized correctly for the below:
		MyEnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
		return output_var.Assign(mip.count); // Will assign zero if the API ever returns a legitimate zero.

	case SYSGET_CMD_MONITORAREA:
	case SYSGET_CMD_MONITORWORKAREA:
		Var *output_var_left, *output_var_top, *output_var_right, *output_var_bottom;
		// Make it longer than max var name so that FindOrAddVar() will be able to spot and report
		// var names that are too long:
		TCHAR var_name[MAX_VAR_NAME_LENGTH + 20];
		// To help performance (in case the linked list of variables is huge), tell FindOrAddVar where
		// to start the search.  Use the base array name rather than the preceding element because,
		// for example, Array19 is alphabetically less than Array2, so we can't rely on the
		// numerical ordering:
		int always_use;
		always_use = FINDVAR_FOR_PSEUDO_ARRAY(output_var);
		if (   !(output_var_left = g_script.FindOrAddVar(var_name
			, sntprintf(var_name, _countof(var_name), _T("%sLeft"), output_var.mName)
			, always_use))   )
			return FAIL;  // It already reported the error.
		if (   !(output_var_top = g_script.FindOrAddVar(var_name
			, sntprintf(var_name, _countof(var_name), _T("%sTop"), output_var.mName)
			, always_use))   )
			return FAIL;
		if (   !(output_var_right = g_script.FindOrAddVar(var_name
			, sntprintf(var_name, _countof(var_name), _T("%sRight"), output_var.mName)
			, always_use))   )
			return FAIL;
		if (   !(output_var_bottom = g_script.FindOrAddVar(var_name
			, sntprintf(var_name, _countof(var_name), _T("%sBottom"), output_var.mName)
			, always_use))   )
			return FAIL;

		RECT monitor_rect;
		if (MyEnumDisplayMonitors)
		{
			mip.monitor_number_to_find = ATOI(aValue);  // If this returns 0, it will default to the primary monitor.
			MyEnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
			if (!mip.count || (mip.monitor_number_to_find && mip.monitor_number_to_find != mip.count))
			{
				// With the exception of the caller having specified a non-existent monitor number, all of
				// the ways the above can happen are probably impossible in practice.  Make all the variables
				// blank vs. zero to indicate the problem.
				output_var_left->Assign();
				output_var_top->Assign();
				output_var_right->Assign();
				output_var_bottom->Assign();
				return OK;
			}
			// Otherwise:
			monitor_rect = (cmd == SYSGET_CMD_MONITORAREA) ? mip.monitor_info_ex.rcMonitor : mip.monitor_info_ex.rcWork;
		}
		else // Win95/NT: Since system only supports 1 monitor, the first must be primary.
		{
			if (cmd == SYSGET_CMD_MONITORAREA)
			{
				monitor_rect.left = 0;
				monitor_rect.top = 0;
				monitor_rect.right = GetSystemMetrics(SM_CXSCREEN);
				monitor_rect.bottom = GetSystemMetrics(SM_CYSCREEN);
			}
			else // Work area
				SystemParametersInfo(SPI_GETWORKAREA, 0, &monitor_rect, 0);  // Get desktop rect excluding task bar.
		}
		output_var_left->Assign(monitor_rect.left);
		output_var_top->Assign(monitor_rect.top);
		output_var_right->Assign(monitor_rect.right);
		output_var_bottom->Assign(monitor_rect.bottom);
		return OK;

	case SYSGET_CMD_MONITORNAME:
		if (MyEnumDisplayMonitors)
		{
			mip.monitor_number_to_find = ATOI(aValue);  // If this returns 0, it will default to the primary monitor.
			MyEnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
			if (!mip.count || (mip.monitor_number_to_find && mip.monitor_number_to_find != mip.count))
				// With the exception of the caller having specified a non-existent monitor number, all of
				// the ways the above can happen are probably impossible in practice.  Make the variable
				// blank to indicate the problem:
				return output_var.Assign();
			else
				return output_var.Assign(mip.monitor_info_ex.szDevice);
		}
		else // Win95/NT: There is probably no way to find out the name of the monitor.
			return output_var.Assign();
	} // switch()

	return FAIL;  // Never executed (increases maintainability and avoids compiler warning).
}



BOOL CALLBACK EnumMonitorProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM lParam)
{
	MonitorInfoPackage &mip = *(MonitorInfoPackage *)lParam;  // For performance and convenience.
	if (mip.monitor_number_to_find == COUNT_ALL_MONITORS)
	{
		++mip.count;
		return TRUE;  // Enumerate all monitors so that they can be counted.
	}
	// GetMonitorInfo() must be dynamically loaded; otherwise, the app won't launch at all on Win95/NT.
	typedef BOOL (WINAPI* GetMonitorInfoType)(HMONITOR, LPMONITORINFO);
	static GetMonitorInfoType MyGetMonitorInfo = (GetMonitorInfoType)
		GetProcAddress(GetModuleHandle(_T("user32")), "GetMonitorInfo" WINAPI_SUFFIX);
	if (!MyGetMonitorInfo) // Shouldn't normally happen since caller wouldn't have called us if OS is Win95/NT. 
		return FALSE;
	if (!MyGetMonitorInfo(hMonitor, &mip.monitor_info_ex)) // Failed.  Probably very rare.
		return FALSE; // Due to the complexity of needing to stop at the correct monitor number, do not continue.
		// In the unlikely event that the above fails when the caller wanted us to find the primary
		// monitor, the caller will think the primary is the previously found monitor (if any).
		// This is just documented here as a known limitation since this combination of circumstances
		// is probably impossible.
	++mip.count; // So that caller can detect failure, increment only now that failure conditions have been checked.
	if (mip.monitor_number_to_find) // Caller gave a specific monitor number, so don't search for the primary monitor.
	{
		if (mip.count == mip.monitor_number_to_find) // Since the desired monitor has been found, must not continue.
			return FALSE;
	}
	else // Caller wants the primary monitor found.
		// MSDN docs are unclear that MONITORINFOF_PRIMARY is a bitwise value, but the name "dwFlags" implies so:
		if (mip.monitor_info_ex.dwFlags & MONITORINFOF_PRIMARY)
			return FALSE;  // Primary has been found and "count" contains its number. Must not continue the enumeration.
			// Above assumes that it is impossible to not have a primary monitor in a system that has at least
			// one monitor.  MSDN certainly implies this through multiple references to the primary monitor.
	// Otherwise, continue the enumeration:
	return TRUE;
}



LPCOLORREF getbits(HBITMAP ahImage, HDC hdc, LONG &aWidth, LONG &aHeight, bool &aIs16Bit, int aMinColorDepth = 8)
// Helper function used by PixelSearch below.
// Returns an array of pixels to the caller, which it must free when done.  Returns NULL on failure,
// in which case the contents of the output parameters is indeterminate.
{
	HDC tdc = CreateCompatibleDC(hdc);
	if (!tdc)
		return NULL;

	// From this point on, "goto end" will assume tdc is non-NULL, but that the below
	// might still be NULL.  Therefore, all of the following must be initialized so that the "end"
	// label can detect them:
	HGDIOBJ tdc_orig_select = NULL;
	LPCOLORREF image_pixel = NULL;
	bool success = false;

	// Confirmed:
	// Needs extra memory to prevent buffer overflow due to: "A bottom-up DIB is specified by setting
	// the height to a positive number, while a top-down DIB is specified by setting the height to a
	// negative number. THE BITMAP COLOR TABLE WILL BE APPENDED to the BITMAPINFO structure."
	// Maybe this applies only to negative height, in which case the second call to GetDIBits()
	// below uses one.
	struct BITMAPINFO3
	{
		BITMAPINFOHEADER    bmiHeader;
		RGBQUAD             bmiColors[260];  // v1.0.40.10: 260 vs. 3 to allow room for color table when color depth is 8-bit or less.
	} bmi;

	bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bmi.bmiHeader.biBitCount = 0; // i.e. "query bitmap attributes" only.
	if (!GetDIBits(tdc, ahImage, 0, 0, NULL, (LPBITMAPINFO)&bmi, DIB_RGB_COLORS)
		|| bmi.bmiHeader.biBitCount < aMinColorDepth) // Relies on short-circuit boolean order.
		goto end;

	// Set output parameters for caller:
	aIs16Bit = (bmi.bmiHeader.biBitCount == 16);
	aWidth = bmi.bmiHeader.biWidth;
	aHeight = bmi.bmiHeader.biHeight;

	int image_pixel_count = aWidth * aHeight;
	if (   !(image_pixel = (LPCOLORREF)malloc(image_pixel_count * sizeof(COLORREF)))   )
		goto end;

	// v1.0.40.10: To preserve compatibility with callers who check for transparency in icons, don't do any
	// of the extra color table handling for 1-bpp images.  Update: For code simplification, support only
	// 8-bpp images.  If ever support lower color depths, use something like "bmi.bmiHeader.biBitCount > 1
	// && bmi.bmiHeader.biBitCount < 9";
	bool is_8bit = (bmi.bmiHeader.biBitCount == 8);
	if (!is_8bit)
		bmi.bmiHeader.biBitCount = 32;
	bmi.bmiHeader.biHeight = -bmi.bmiHeader.biHeight; // Storing a negative inside the bmiHeader struct is a signal for GetDIBits().

	// Must be done only after GetDIBits() because: "The bitmap identified by the hbmp parameter
	// must not be selected into a device context when the application calls GetDIBits()."
	// (Although testing shows it works anyway, perhaps because GetDIBits() above is being
	// called in its informational mode only).
	// Note that this seems to return NULL sometimes even though everything still works.
	// Perhaps that is normal.
	tdc_orig_select = SelectObject(tdc, ahImage); // Returns NULL when we're called the second time?

	// Apparently there is no need to specify DIB_PAL_COLORS below when color depth is 8-bit because
	// DIB_RGB_COLORS also retrieves the color indices.
	if (   !(GetDIBits(tdc, ahImage, 0, aHeight, image_pixel, (LPBITMAPINFO)&bmi, DIB_RGB_COLORS))   )
		goto end;

	if (is_8bit) // This section added in v1.0.40.10.
	{
		// Convert the color indices to RGB colors by going through the array in reverse order.
		// Reverse order allows an in-place conversion of each 8-bit color index to its corresponding
		// 32-bit RGB color.
		LPDWORD palette = (LPDWORD)_alloca(256 * sizeof(PALETTEENTRY));
		GetSystemPaletteEntries(tdc, 0, 256, (LPPALETTEENTRY)palette); // Even if failure can realistically happen, consequences of using uninitialized palette seem acceptable.
		// Above: GetSystemPaletteEntries() is the only approach that provided the correct palette.
		// The following other approaches didn't give the right one:
		// GetDIBits(): The palette it stores in bmi.bmiColors seems completely wrong.
		// GetPaletteEntries()+GetCurrentObject(hdc, OBJ_PAL): Returned only 20 entries rather than the expected 256.
		// GetDIBColorTable(): I think same as above or maybe it returns 0.

		// The following section is necessary because apparently each new row in the region starts on
		// a DWORD boundary.  So if the number of pixels in each row isn't an exact multiple of 4, there
		// are between 1 and 3 zero-bytes at the end of each row.
		int remainder = aWidth % 4;
		int empty_bytes_at_end_of_each_row = remainder ? (4 - remainder) : 0;

		// Start at the last RGB slot and the last color index slot:
		BYTE *byte = (BYTE *)image_pixel + image_pixel_count - 1 + (aHeight * empty_bytes_at_end_of_each_row); // Pointer to 8-bit color indices.
		DWORD *pixel = image_pixel + image_pixel_count - 1; // Pointer to 32-bit RGB entries.

		int row, col;
		for (row = 0; row < aHeight; ++row) // For each row.
		{
			byte -= empty_bytes_at_end_of_each_row;
			for (col = 0; col < aWidth; ++col) // For each column.
				*pixel-- = rgb_to_bgr(palette[*byte--]); // Caller always wants RGB vs. BGR format.
		}
	}
	
	// Since above didn't "goto end", indicate success:
	success = true;

end:
	if (tdc_orig_select) // i.e. the original call to SelectObject() didn't fail.
		SelectObject(tdc, tdc_orig_select); // Probably necessary to prevent memory leak.
	DeleteDC(tdc);
	if (!success && image_pixel)
	{
		free(image_pixel);
		image_pixel = NULL;
	}
	return image_pixel;
}



ResultType Line::PixelSearch(int aLeft, int aTop, int aRight, int aBottom, COLORREF aColorBGR
	, int aVariation, LPTSTR aOptions, bool aIsPixelGetColor)
// Caller has ensured that aColor is in BGR format unless caller passed true for aUseRGB, in which case
// it's in RGB format.
// Author: The fast-mode PixelSearch was created by Aurelian Maga.
{
	// For maintainability, get options and RGB/BGR conversion out of the way early.
	bool fast_mode = aIsPixelGetColor || tcscasestr(aOptions, _T("Fast"));
	bool use_rgb = tcscasestr(aOptions, _T("RGB")) != NULL;
	COLORREF aColorRGB;
	if (use_rgb) // aColorBGR currently contains an RGB value.
	{
		aColorRGB = aColorBGR;
		aColorBGR = rgb_to_bgr(aColorBGR);
	}
	else
		aColorRGB = rgb_to_bgr(aColorBGR); // rgb_to_bgr() also converts in the reverse direction, i.e. bgr_to_rgb().

	// Many of the following sections are similar to those in ImageSearch(), so they should be
	// maintained together.

	Var *output_var_x = ARGVAR1;  // Ok if NULL. Load-time validation has ensured that these are valid output variables (e.g. not built-in vars).
	Var *output_var_y = aIsPixelGetColor ? NULL : ARGVAR2;  // Ok if NULL. ARGVARRAW2 wouldn't be safe because load-time validation requires a min of only zero parameters to allow the output variables to be left blank.

	if (output_var_x)
		output_var_x->Assign();  // Init to empty string regardless of whether we succeed here.
	if (output_var_y)
		output_var_y->Assign();  // Same.

	POINT origin = {0};
	CoordToScreen(origin, COORD_MODE_PIXEL);
	aLeft   += origin.x;
	aTop    += origin.y;
	aRight  += origin.x;
	aBottom += origin.y;

	if (aVariation < 0)
		aVariation = 0;
	if (aVariation > 255)
		aVariation = 255;

	// Allow colors to vary within the spectrum of intensity, rather than having them
	// wrap around (which doesn't seem to make much sense).  For example, if the user specified
	// a variation of 5, but the red component of aColorBGR is only 0x01, we don't want red_low to go
	// below zero, which would cause it to wrap around to a very intense red color:
	COLORREF pixel; // Used much further down.
	BYTE red, green, blue; // Used much further down.
	BYTE search_red, search_green, search_blue;
	BYTE red_low, green_low, blue_low, red_high, green_high, blue_high;
	if (aVariation > 0)
	{
		search_red = GetRValue(aColorBGR);
		search_green = GetGValue(aColorBGR);
		search_blue = GetBValue(aColorBGR);
	}
	//else leave uninitialized since they won't be used.

	HDC hdc = GetDC(NULL);
	if (!hdc)
		goto error;

	bool found = false; // Must init here for use by "goto fast_end" and for use by both fast and slow modes.

	if (fast_mode)
	{
		// From this point on, "goto fast_end" will assume hdc is non-NULL but that the below might still be NULL.
		// Therefore, all of the following must be initialized so that the "fast_end" label can detect them:
		HDC sdc = NULL;
		HBITMAP hbitmap_screen = NULL;
		LPCOLORREF screen_pixel = NULL;
		HGDIOBJ sdc_orig_select = NULL;

		// Some explanation for the method below is contained in this quote from the newsgroups:
		// "you shouldn't really be getting the current bitmap from the GetDC DC. This might
		// have weird effects like returning the entire screen or not working. Create yourself
		// a memory DC first of the correct size. Then BitBlt into it and then GetDIBits on
		// that instead. This way, the provider of the DC (the video driver) can make sure that
		// the correct pixels are copied across."

		// Create an empty bitmap to hold all the pixels currently visible on the screen (within the search area):
		int search_width = aRight - aLeft + 1;
		int search_height = aBottom - aTop + 1;
		if (   !(sdc = CreateCompatibleDC(hdc)) || !(hbitmap_screen = CreateCompatibleBitmap(hdc, search_width, search_height))   )
			goto fast_end;

		if (   !(sdc_orig_select = SelectObject(sdc, hbitmap_screen))   )
			goto fast_end;

		// Copy the pixels in the search-area of the screen into the DC to be searched:
		if (   !(BitBlt(sdc, 0, 0, search_width, search_height, hdc, aLeft, aTop, SRCCOPY))   )
			goto fast_end;

		LONG screen_width, screen_height;
		bool screen_is_16bit;
		if (   !(screen_pixel = getbits(hbitmap_screen, sdc, screen_width, screen_height, screen_is_16bit))   )
			goto fast_end;

		// Concerning 0xF8F8F8F8: "On 16bit and 15 bit color the first 5 bits in each byte are valid
		// (in 16bit there is an extra bit but i forgot for which color). And this will explain the
		// second problem [in the test script], since GetPixel even in 16bit will return some "valid"
		// data in the last 3bits of each byte."
		register int i;
		LONG screen_pixel_count = screen_width * screen_height;
		if (screen_is_16bit)
			for (i = 0; i < screen_pixel_count; ++i)
				screen_pixel[i] &= 0xF8F8F8F8;

		if (aIsPixelGetColor)
		{
			COLORREF color = screen_pixel[0] & 0x00FFFFFF; // See other 0x00FFFFFF below for explanation.
			TCHAR buf[32];
			_stprintf(buf, _T("0x%06X"), use_rgb ? color : rgb_to_bgr(color));
			output_var_x->Assign(buf); // Caller has ensured that first output_var (x) won't be NULL in this mode.
			found = true; // ErrorLevel will be set to 0 further below.
		}
		else if (aVariation < 1) // Caller wants an exact match on one particular color.
		{
			if (screen_is_16bit)
				aColorRGB &= 0xF8F8F8F8;
			for (i = 0; i < screen_pixel_count; ++i)
			{
				// Note that screen pixels sometimes have a non-zero high-order byte.  That's why
				// bit-and with 0x00FFFFFF is done.  Otherwise, reddish/orangish colors are not properly
				// found:
				if ((screen_pixel[i] & 0x00FFFFFF) == aColorRGB)
				{
					found = true;
					break;
				}
			}
		}
		else
		{
			// It seems more appropriate to do the 16-bit conversion prior to SET_COLOR_RANGE,
			// rather than applying 0xF8 to each of the high/low values individually.
			if (screen_is_16bit)
			{
				search_red &= 0xF8;
				search_green &= 0xF8;
				search_blue &= 0xF8;
			}

#define SET_COLOR_RANGE \
{\
	red_low = (aVariation > search_red) ? 0 : search_red - aVariation;\
	green_low = (aVariation > search_green) ? 0 : search_green - aVariation;\
	blue_low = (aVariation > search_blue) ? 0 : search_blue - aVariation;\
	red_high = (aVariation > 0xFF - search_red) ? 0xFF : search_red + aVariation;\
	green_high = (aVariation > 0xFF - search_green) ? 0xFF : search_green + aVariation;\
	blue_high = (aVariation > 0xFF - search_blue) ? 0xFF : search_blue + aVariation;\
}
			
			SET_COLOR_RANGE

			for (i = 0; i < screen_pixel_count; ++i)
			{
				// Note that screen pixels sometimes have a non-zero high-order byte.  But it doesn't
				// matter with the below approach, since that byte is not checked in the comparison.
				pixel = screen_pixel[i];
				red = GetBValue(pixel);   // Because pixel is in RGB vs. BGR format, red is retrieved with
				green = GetGValue(pixel); // GetBValue() and blue is retrieved with GetRValue().
				blue = GetRValue(pixel);
				if (red >= red_low && red <= red_high
					&& green >= green_low && green <= green_high
					&& blue >= blue_low && blue <= blue_high)
				{
					found = true;
					break;
				}
			}
		}

fast_end:
		// If found==false when execution reaches here, ErrorLevel is already set to the right value, so just
		// clean up then return.
		ReleaseDC(NULL, hdc);
		if (sdc)
		{
			if (sdc_orig_select) // i.e. the original call to SelectObject() didn't fail.
				SelectObject(sdc, sdc_orig_select); // Probably necessary to prevent memory leak.
			DeleteDC(sdc);
		}
		if (hbitmap_screen)
			DeleteObject(hbitmap_screen);
		if (screen_pixel)
			free(screen_pixel);
		else // One of the GDI calls failed and the search wasn't carried out.
			goto error;

		// Otherwise, success.  Calculate xpos and ypos of where the match was found and adjust
		// coords to make them relative to the position of the target window (rect will contain
		// zeroes if this doesn't need to be done):
		if (!aIsPixelGetColor && found)
		{
			if (output_var_x && !output_var_x->Assign((aLeft + i%screen_width) - origin.x))
				return FAIL;
			if (output_var_y && !output_var_y->Assign((aTop + i/screen_width) - origin.y))
				return FAIL;
		}

		return g_ErrorLevel->Assign(found ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR); // "0" indicates success; "1" indicates search completed okay, but didn't find it.
	}

	// Otherwise (since above didn't return): fast_mode==false
	// This old/slower method is kept  because fast mode will break older scripts that rely on
	// which match is found if there is more than one match (since fast mode searches the
	// pixels in a different order (horizontally rather than vertically, I believe).
	// In addition, there is doubt that the fast mode works in all the screen color depths, games,
	// and other circumstances that the slow mode is known to work in.

	// If the caller gives us inverted X or Y coordinates, conduct the search in reverse order.
	// This feature was requested; it was put into effect for v1.0.25.06.
	bool right_to_left = aLeft > aRight;
	bool bottom_to_top = aTop > aBottom;
	register int xpos, ypos;

	if (aVariation > 0)
		SET_COLOR_RANGE

	for (xpos = aLeft  // It starts at aLeft even if right_to_left is true.
		; (right_to_left ? (xpos >= aRight) : (xpos <= aRight)) // Verified correct.
		; xpos += right_to_left ? -1 : 1)
	{
		for (ypos = aTop  // It starts at aTop even if bottom_to_top is true.
			; bottom_to_top ? (ypos >= aBottom) : (ypos <= aBottom) // Verified correct.
			; ypos += bottom_to_top ? -1 : 1)
		{
			pixel = GetPixel(hdc, xpos, ypos); // Returns a BGR value, not RGB.
			if (aVariation < 1)  // User wanted an exact match.
			{
				if (pixel == aColorBGR)
				{
					found = true;
					break;
				}
			}
			else  // User specified that some variation in each of the RGB components is allowable.
			{
				red = GetRValue(pixel);
				green = GetGValue(pixel);
				blue = GetBValue(pixel);
				if (red >= red_low && red <= red_high && green >= green_low && green <= green_high
					&& blue >= blue_low && blue <= blue_high)
				{
					found = true;
					break;
				}
			}
		}
		// Check this here rather than in the outer loop's top line because otherwise the loop's
		// increment would make xpos too big by 1:
		if (found)
			break;
	}

	ReleaseDC(NULL, hdc);

	if (!found)
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // This value indicates "color not found".

	// Otherwise, this pixel matches one of the specified color(s).
	// Adjust coords to make them relative to the position of the target window
	// (rect will contain zeroes if this doesn't need to be done):
	if (output_var_x && !output_var_x->Assign(xpos - origin.x))
		return FAIL;
	if (output_var_y && !output_var_y->Assign(ypos - origin.y))
		return FAIL;
	// Since above didn't return:
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.

error:
	return SetErrorLevelOrThrowStr(aIsPixelGetColor ? ERRORLEVEL_ERROR : ERRORLEVEL_ERROR2); // 2 means error other than "color not found".
}



ResultType Line::ImageSearch(int aLeft, int aTop, int aRight, int aBottom, LPTSTR aImageFile)
// Author: ImageSearch was created by Aurelian Maga.
{
	// Many of the following sections are similar to those in PixelSearch(), so they should be
	// maintained together.
	Var *output_var_x = ARGVAR1;  // Ok if NULL. RAW wouldn't be safe because load-time validation actually
	Var *output_var_y = ARGVAR2;  // requires a minimum of zero parameters so that the output-vars can be optional. Also:
	// Load-time validation has ensured that these are valid output variables (e.g. not built-in vars).

	// Set default results (output variables), in case of early return:
	if (output_var_x)
		output_var_x->Assign();  // Init to empty string regardless of whether we succeed here.
	if (output_var_y)
		output_var_y->Assign(); // Same.

	POINT origin = {0};
	CoordToScreen(origin, COORD_MODE_PIXEL);
	aLeft   += origin.x;
	aTop    += origin.y;
	aRight  += origin.x;
	aBottom += origin.y;

	// Options are done as asterisk+option to permit future expansion.
	// Set defaults to be possibly overridden by any specified options:
	int aVariation = 0;  // This is named aVariation vs. variation for use with the SET_COLOR_RANGE macro.
	COLORREF trans_color = CLR_NONE; // The default must be a value that can't occur naturally in an image.
	int icon_number = 0; // Zero means "load icon or bitmap (doesn't matter)".
	int width = 0, height = 0;
	// For icons, override the default to be 16x16 because that is what is sought 99% of the time.
	// This new default can be overridden by explicitly specifying w0 h0:
	LPTSTR cp = _tcsrchr(aImageFile, '.');
	if (cp)
	{
		++cp;
		if (!(_tcsicmp(cp, _T("ico")) && _tcsicmp(cp, _T("exe")) && _tcsicmp(cp, _T("dll"))))
			width = GetSystemMetrics(SM_CXSMICON), height = GetSystemMetrics(SM_CYSMICON);
	}

	TCHAR color_name[32], *dp;
	cp = omit_leading_whitespace(aImageFile); // But don't alter aImageFile yet in case it contains literal whitespace we want to retain.
	while (*cp == '*')
	{
		++cp;
		switch (_totupper(*cp))
		{
		case 'W': width = ATOI(cp + 1); break;
		case 'H': height = ATOI(cp + 1); break;
		default:
			if (!_tcsnicmp(cp, _T("Icon"), 4))
			{
				cp += 4;  // Now it's the character after the word.
				icon_number = ATOI(cp); // LoadPicture() correctly handles any negative value.
			}
			else if (!_tcsnicmp(cp, _T("Trans"), 5))
			{
				cp += 5;  // Now it's the character after the word.
				// Isolate the color name/number for ColorNameToBGR():
				tcslcpy(color_name, cp, _countof(color_name));
				if (dp = StrChrAny(color_name, _T(" \t"))) // Find space or tab, if any.
					*dp = '\0';
				// Fix for v1.0.44.10: Treat trans_color as containing an RGB value (not BGR) so that it matches
				// the documented behavior.  In older versions, a specified color like "TransYellow" was wrong in
				// every way (inverted) and a specified numeric color like "Trans0xFFFFAA" was treated as BGR vs. RGB.
				trans_color = ColorNameToBGR(color_name);
				if (trans_color == CLR_NONE) // A matching color name was not found, so assume it's in hex format.
					// It seems _tcstol() automatically handles the optional leading "0x" if present:
					trans_color = _tcstol(color_name, NULL, 16);
					// if color_name did not contain something hex-numeric, black (0x00) will be assumed,
					// which seems okay given how rare such a problem would be.
				else
					trans_color = bgr_to_rgb(trans_color); // v1.0.44.10: See fix/comment above.

			}
			else // Assume it's a number since that's the only other asterisk-option.
			{
				aVariation = ATOI(cp); // Seems okay to support hex via ATOI because the space after the number is documented as being mandatory.
				if (aVariation < 0)
					aVariation = 0;
				if (aVariation > 255)
					aVariation = 255;
				// Note: because it's possible for filenames to start with a space (even though Explorer itself
				// won't let you create them that way), allow exactly one space between end of option and the
				// filename itself:
			}
		} // switch()
		if (   !(cp = StrChrAny(cp, _T(" \t")))   ) // Find the first space or tab after the option.
			goto error; // Bad option/format.
		// Now it's the space or tab (if there is one) after the option letter.  Advance by exactly one character
		// because only one space or tab is considered the delimiter.  Any others are considered to be part of the
		// filename (though some or all OSes might simply ignore them or tolerate them as first-try match criteria).
		aImageFile = ++cp; // This should now point to another asterisk or the filename itself.
		// Above also serves to reset the filename to omit the option string whenever at least one asterisk-option is present.
		cp = omit_leading_whitespace(cp); // This is done to make it more tolerant of having more than one space/tab between options.
	}

	// Update: Transparency is now supported in icons by using the icon's mask.  In addition, an attempt
	// is made to support transparency in GIF, PNG, and possibly TIF files via the *Trans option, which
	// assumes that one color in the image is transparent.  In GIFs not loaded via GDIPlus, the transparent
	// color might always been seen as pure white, but when GDIPlus is used, it's probably always black
	// like it is in PNG -- however, this will not relied upon, at least not until confirmed.
	// OLDER/OBSOLETE comment kept for background:
	// For now, images that can't be loaded as bitmaps (icons and cursors) are not supported because most
	// icons have a transparent background or color present, which the image search routine here is
	// probably not equipped to handle (since the transparent color, when shown, typically reveals the
	// color of whatever is behind it; thus screen pixel color won't match image's pixel color).
	// So currently, only BMP and GIF seem to work reliably, though some of the other GDIPlus-supported
	// formats might work too.
	int image_type;
	bool no_delete_bitmap;
	HBITMAP hbitmap_image = LoadPicture(aImageFile, width, height, image_type, icon_number, false, &no_delete_bitmap);
	// The comment marked OBSOLETE below is no longer true because the elimination of the high-byte via
	// 0x00FFFFFF seems to have fixed it.  But "true" is still not passed because that should increase
	// consistency when GIF/BMP/ICO files are used by a script on both Win9x and other OSs (since the
	// same loading method would be used via "false" for these formats across all OSes).
	// OBSOLETE: Must not pass "true" with the above because that causes bitmaps and gifs to be not found
	// by the search.  In other words, nothing works.  Obsolete comment: Pass "true" so that an attempt
	// will be made to load icons as bitmaps if GDIPlus is available.
	if (!hbitmap_image)
		goto error;

	HDC hdc = GetDC(NULL);
	if (!hdc)
	{
		if (!no_delete_bitmap)
		{
			if (image_type == IMAGE_ICON)
				DestroyIcon((HICON)hbitmap_image);
			else
				DeleteObject(hbitmap_image);
		}
		goto error;
	}

	// From this point on, "goto end" will assume hdc and hbitmap_image are non-NULL, but that the below
	// might still be NULL.  Therefore, all of the following must be initialized so that the "end"
	// label can detect them:
	HDC sdc = NULL;
	HBITMAP hbitmap_screen = NULL;
	LPCOLORREF image_pixel = NULL, screen_pixel = NULL, image_mask = NULL;
	HGDIOBJ sdc_orig_select = NULL;
	bool found = false; // Must init here for use by "goto end".
    
	bool image_is_16bit;
	LONG image_width, image_height;

	if (image_type == IMAGE_ICON)
	{
		// Must be done prior to IconToBitmap() since it deletes (HICON)hbitmap_image:
		ICONINFO ii;
		if (GetIconInfo((HICON)hbitmap_image, &ii))
		{
			// If the icon is monochrome (black and white), ii.hbmMask will contain twice as many pixels as
			// are actually in the icon.  But since the top half of the pixels are the AND-mask, it seems
			// okay to get all the pixels given the rarity of monochrome icons.  This scenario should be
			// handled properly because: 1) the variables image_height and image_width will be overridden
			// further below with the correct icon dimensions; 2) Only the first half of the pixels within
			// the image_mask array will actually be referenced by the transparency checker in the loops,
			// and that first half is the AND-mask, which is the transparency part that is needed.  The
			// second half, the XOR part, is not needed and thus ignored.  Also note that if width/height
			// required the icon to be scaled, LoadPicture() has already done that directly to the icon,
			// so ii.hbmMask should already be scaled to match the size of the bitmap created later below.
			image_mask = getbits(ii.hbmMask, hdc, image_width, image_height, image_is_16bit, 1);
			DeleteObject(ii.hbmColor); // DeleteObject() probably handles NULL okay since few MSDN/other examples ever check for NULL.
			DeleteObject(ii.hbmMask);
		}
		if (   !(hbitmap_image = IconToBitmap((HICON)hbitmap_image, true))   )
			goto error;
	}

	if (   !(image_pixel = getbits(hbitmap_image, hdc, image_width, image_height, image_is_16bit))   )
		goto end;

	// Create an empty bitmap to hold all the pixels currently visible on the screen that lie within the search area:
	int search_width = aRight - aLeft + 1;
	int search_height = aBottom - aTop + 1;
	if (   !(sdc = CreateCompatibleDC(hdc)) || !(hbitmap_screen = CreateCompatibleBitmap(hdc, search_width, search_height))   )
		goto end;

	if (   !(sdc_orig_select = SelectObject(sdc, hbitmap_screen))   )
		goto end;

	// Copy the pixels in the search-area of the screen into the DC to be searched:
	if (   !(BitBlt(sdc, 0, 0, search_width, search_height, hdc, aLeft, aTop, SRCCOPY))   )
		goto end;

	LONG screen_width, screen_height;
	bool screen_is_16bit;
	if (   !(screen_pixel = getbits(hbitmap_screen, sdc, screen_width, screen_height, screen_is_16bit))   )
		goto end;

	LONG image_pixel_count = image_width * image_height;
	LONG screen_pixel_count = screen_width * screen_height;
	int i, j, k, x, y; // Declaring as "register" makes no performance difference with current compiler, so let the compiler choose which should be registers.

	// If either is 16-bit, convert *both* to the 16-bit-compatible 32-bit format:
	if (image_is_16bit || screen_is_16bit)
	{
		if (trans_color != CLR_NONE)
			trans_color &= 0x00F8F8F8; // Convert indicated trans-color to be compatible with the conversion below.
		for (i = 0; i < screen_pixel_count; ++i)
			screen_pixel[i] &= 0x00F8F8F8; // Highest order byte must be masked to zero for consistency with use of 0x00FFFFFF below.
		for (i = 0; i < image_pixel_count; ++i)
			image_pixel[i] &= 0x00F8F8F8;  // Same.
	}

	// v1.0.44.03: The below is now done even for variation>0 mode so its results are consistent with those of
	// non-variation mode.  This is relied upon by variation=0 mode but now also by the following line in the
	// variation>0 section:
	//     || image_pixel[j] == trans_color
	// Without this change, there are cases where variation=0 would find a match but a higher variation
	// (for the same search) wouldn't. 
	for (i = 0; i < image_pixel_count; ++i)
		image_pixel[i] &= 0x00FFFFFF;

	// Search the specified region for the first occurrence of the image:
	if (aVariation < 1) // Caller wants an exact match.
	{
		// Concerning the following use of 0x00FFFFFF, the use of 0x00F8F8F8 above is related (both have high order byte 00).
		// The following needs to be done only when shades-of-variation mode isn't in effect because
		// shades-of-variation mode ignores the high-order byte due to its use of macros such as GetRValue().
		// This transformation incurs about a 15% performance decrease (percentage is fairly constant since
		// it is proportional to the search-region size, which tends to be much larger than the search-image and
		// is therefore the primary determination of how long the loops take). But it definitely helps find images
		// more successfully in some cases.  For example, if a PNG file is displayed in a GUI window, this
		// transformation allows certain bitmap search-images to be found via variation==0 when they otherwise
		// would require variation==1 (possibly the variation==1 success is just a side-effect of it
		// ignoring the high-order byte -- maybe a much higher variation would be needed if the high
		// order byte were also subject to the same shades-of-variation analysis as the other three bytes [RGB]).
		for (i = 0; i < screen_pixel_count; ++i)
			screen_pixel[i] &= 0x00FFFFFF;

		for (i = 0; i < screen_pixel_count; ++i)
		{
			// Unlike the variation-loop, the following one uses a first-pixel optimization to boost performance
			// by about 10% because it's only 3 extra comparisons and exact-match mode is probably used more often.
			// Before even checking whether the other adjacent pixels in the region match the image, ensure
			// the image does not extend past the right or bottom edges of the current part of the search region.
			// This is done for performance but more importantly to prevent partial matches at the edges of the
			// search region from being considered complete matches.
			// The following check is ordered for short-circuit performance.  In addition, image_mask, if
			// non-NULL, is used to determine which pixels are transparent within the image and thus should
			// match any color on the screen.
			if ((screen_pixel[i] == image_pixel[0] // A screen pixel has been found that matches the image's first pixel.
				|| image_mask && image_mask[0]     // Or: It's an icon's transparent pixel, which matches any color.
				|| image_pixel[0] == trans_color)  // This should be okay even if trans_color==CLR_NONE, since CLR_NONE should never occur naturally in the image.
				&& image_height <= screen_height - i/screen_width // Image is short enough to fit in the remaining rows of the search region.
				&& image_width <= screen_width - i%screen_width)  // Image is narrow enough not to exceed the right-side boundary of the search region.
			{
				// Check if this candidate region -- which is a subset of the search region whose height and width
				// matches that of the image -- is a pixel-for-pixel match of the image.
				for (found = true, x = 0, y = 0, j = 0, k = i; j < image_pixel_count; ++j)
				{
					if (!(found = (screen_pixel[k] == image_pixel[j] // At least one pixel doesn't match, so this candidate is discarded.
						|| image_mask && image_mask[j]      // Or: It's an icon's transparent pixel, which matches any color.
						|| image_pixel[j] == trans_color))) // This should be okay even if trans_color==CLR_NONE, since CLR none should never occur naturally in the image.
						break;
					if (++x < image_width) // We're still within the same row of the image, so just move on to the next screen pixel.
						++k;
					else // We're starting a new row of the image.
					{
						x = 0; // Return to the leftmost column of the image.
						++y;   // Move one row downward in the image.
						// Move to the next row within the current-candidate region (not the entire search region).
						// This is done by moving vertically downward from "i" (which is the upper-left pixel of the
						// current-candidate region) by "y" rows.
						k = i + y*screen_width; // Verified correct.
					}
				}
				if (found) // Complete match found.
					break;
			}
		}
	}
	else // Allow colors to vary by aVariation shades; i.e. approximate match is okay.
	{
		// The following section is part of the first-pixel-check optimization that improves performance by
		// 15% or more depending on where and whether a match is found.  This section and one the follows
		// later is commented out to reduce code size.
		// Set high/low range for the first pixel of the image since it is the pixel most often checked
		// (i.e. for performance).
		//BYTE search_red1 = GetBValue(image_pixel[0]);  // Because it's RGB vs. BGR, the B value is fetched, not R (though it doesn't matter as long as everything is internally consistent here).
		//BYTE search_green1 = GetGValue(image_pixel[0]);
		//BYTE search_blue1 = GetRValue(image_pixel[0]); // Same comment as above.
		//BYTE red_low1 = (aVariation > search_red1) ? 0 : search_red1 - aVariation;
		//BYTE green_low1 = (aVariation > search_green1) ? 0 : search_green1 - aVariation;
		//BYTE blue_low1 = (aVariation > search_blue1) ? 0 : search_blue1 - aVariation;
		//BYTE red_high1 = (aVariation > 0xFF - search_red1) ? 0xFF : search_red1 + aVariation;
		//BYTE green_high1 = (aVariation > 0xFF - search_green1) ? 0xFF : search_green1 + aVariation;
		//BYTE blue_high1 = (aVariation > 0xFF - search_blue1) ? 0xFF : search_blue1 + aVariation;
		// Above relies on the fact that the 16-bit conversion higher above was already done because like
		// in PixelSearch, it seems more appropriate to do the 16-bit conversion prior to setting the range
		// of high and low colors (vs. than applying 0xF8 to each of the high/low values individually).

		BYTE red, green, blue;
		BYTE search_red, search_green, search_blue;
		BYTE red_low, green_low, blue_low, red_high, green_high, blue_high;

		// The following loop is very similar to its counterpart above that finds an exact match, so maintain
		// them together and see above for more detailed comments about it.
		for (i = 0; i < screen_pixel_count; ++i)
		{
			// The following is commented out to trade code size reduction for performance (see comment above).
			//red = GetBValue(screen_pixel[i]);   // Because it's RGB vs. BGR, the B value is fetched, not R (though it doesn't matter as long as everything is internally consistent here).
			//green = GetGValue(screen_pixel[i]);
			//blue = GetRValue(screen_pixel[i]);
			//if ((red >= red_low1 && red <= red_high1
			//	&& green >= green_low1 && green <= green_high1
			//	&& blue >= blue_low1 && blue <= blue_high1 // All three color components are a match, so this screen pixel matches the image's first pixel.
			//		|| image_mask && image_mask[0]         // Or: It's an icon's transparent pixel, which matches any color.
			//		|| image_pixel[0] == trans_color)      // This should be okay even if trans_color==CLR_NONE, since CLR none should never occur naturally in the image.
			//	&& image_height <= screen_height - i/screen_width // Image is short enough to fit in the remaining rows of the search region.
			//	&& image_width <= screen_width - i%screen_width)  // Image is narrow enough not to exceed the right-side boundary of the search region.
			
			// Instead of the above, only this abbreviated check is done:
			if (image_height <= screen_height - i/screen_width    // Image is short enough to fit in the remaining rows of the search region.
				&& image_width <= screen_width - i%screen_width)  // Image is narrow enough not to exceed the right-side boundary of the search region.
			{
				// Since the first pixel is a match, check the other pixels.
				for (found = true, x = 0, y = 0, j = 0, k = i; j < image_pixel_count; ++j)
				{
   					search_red = GetBValue(image_pixel[j]);
	   				search_green = GetGValue(image_pixel[j]);
		   			search_blue = GetRValue(image_pixel[j]);
					SET_COLOR_RANGE
   					red = GetBValue(screen_pixel[k]);
	   				green = GetGValue(screen_pixel[k]);
		   			blue = GetRValue(screen_pixel[k]);

					if (!(found = red >= red_low && red <= red_high
						&& green >= green_low && green <= green_high
                        && blue >= blue_low && blue <= blue_high
							|| image_mask && image_mask[j]     // Or: It's an icon's transparent pixel, which matches any color.
							|| image_pixel[j] == trans_color)) // This should be okay even if trans_color==CLR_NONE, since CLR_NONE should never occur naturally in the image.
						break; // At least one pixel doesn't match, so this candidate is discarded.
					if (++x < image_width) // We're still within the same row of the image, so just move on to the next screen pixel.
						++k;
					else // We're starting a new row of the image.
					{
						x = 0; // Return to the leftmost column of the image.
						++y;   // Move one row downward in the image.
						k = i + y*screen_width; // Verified correct.
					}
				}
				if (found) // Complete match found.
					break;
			}
		}
	}

	if (!found) // Must override ErrorLevel to its new value prior to the label below.
		g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // "1" indicates search completed okay, but didn't find it.

end:
	// If found==false when execution reaches here, ErrorLevel is already set to the right value, so just
	// clean up then return.
	ReleaseDC(NULL, hdc);
	if (!no_delete_bitmap)
		DeleteObject(hbitmap_image);
	if (sdc)
	{
		if (sdc_orig_select) // i.e. the original call to SelectObject() didn't fail.
			SelectObject(sdc, sdc_orig_select); // Probably necessary to prevent memory leak.
		DeleteDC(sdc);
	}
	if (hbitmap_screen)
		DeleteObject(hbitmap_screen);
	if (image_pixel)
		free(image_pixel);
	if (image_mask)
		free(image_mask);
	if (screen_pixel)
		free(screen_pixel);
	else // One of the GDI calls failed.
		goto error;

	if (!found) // Let ErrorLevel, which is either "1" or "2" as set earlier, tell the story.
		return OK;

	// Otherwise, success.  Calculate xpos and ypos of where the match was found and adjust
	// coords to make them relative to the position of the target window (rect will contain
	// zeroes if this doesn't need to be done):
	if (output_var_x && !output_var_x->Assign((aLeft + i%screen_width) - origin.x))
		return FAIL;
	if (output_var_y && !output_var_y->Assign((aTop + i/screen_width) - origin.y))
		return FAIL;

	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.

error:
	return SetErrorLevelOrThrowStr(ERRORLEVEL_ERROR2);
}



/////////////////
// Main Window //
/////////////////

LRESULT CALLBACK MainWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
	int i;
	DWORD_PTR dwTemp;

	// Detect Explorer crashes so that tray icon can be recreated.  I think this only works on Win98
	// and beyond, since the feature was never properly implemented in Win95:
	static UINT WM_TASKBARCREATED = RegisterWindowMessage(_T("TaskbarCreated"));

	// See GuiWindowProc() for details about this first section:
	LRESULT msg_reply;
	if (g_MsgMonitor.Count() // Count is checked here to avoid function-call overhead.
		&& (!g->CalledByIsDialogMessageOrDispatch || g->CalledByIsDialogMessageOrDispatchMsg != iMsg) // v1.0.44.11: If called by IsDialog or Dispatch but they changed the message number, check if the script is monitoring that new number.
		&& MsgMonitor(hWnd, iMsg, wParam, lParam, NULL, msg_reply))
		return msg_reply; // MsgMonitor has returned "true", indicating that this message should be omitted from further processing.
	g->CalledByIsDialogMessageOrDispatch = false; // v1.0.40.01.

	TRANSLATE_AHK_MSG(iMsg, wParam)
	
	switch (iMsg)
	{
	case WM_COMMAND:
		if (HandleMenuItem(hWnd, LOWORD(wParam), NULL)) // It was handled fully. NULL flags it as a non-GUI menu item such as a tray menu or popup menu.
			return 0; // If an application processes this message, it should return zero.
		break; // Otherwise, let DefWindowProc() try to handle it (this actually seems to happen normally sometimes).

	case AHK_NOTIFYICON:  // Tray icon clicked on.
	{
        switch(lParam)
        {
// Don't allow the main window to be opened this way by a compiled EXE, since it will display
// the lines most recently executed, and many people who compile scripts don't want their users
// to see the contents of the script:
		case WM_LBUTTONDOWN:
			if (g_script.mTrayMenu->mClickCount != 1) // Activating tray menu's default item requires double-click.
				break; // Let default proc handle it (since that's what used to happen, it seems safest).
			//else fall through to the next case.
		case WM_LBUTTONDBLCLK:
			if (g_script.mTrayMenu->mDefault)
				POST_AHK_USER_MENU(hWnd, g_script.mTrayMenu->mDefault->mMenuID, NULL) // NULL flags it as a non-GUI menu item.
#ifdef AUTOHOTKEYSC
			else if (g_script.mTrayMenu->mIncludeStandardItems && g_AllowMainWindow)
				ShowMainWindow();
			// else do nothing.
#else
			else if (g_script.mTrayMenu->mIncludeStandardItems)
				ShowMainWindow();
			// else do nothing.
#endif
			return 0;
		case WM_RBUTTONUP:
			// v1.0.30.03:
			// Opening the menu upon UP vs. DOWN solves at least one set of problems: The fact that
			// when the right mouse button is remapped as shown in the example below, it prevents
			// the left button from being able to select a menu item from the tray menu.  It might
			// solve other problems also, and it seems fairly common for other apps to open the
			// menu upon UP rather than down.  Even Explorer's own context menus are like this.
			// The following example is trivial and serves only to illustrate the problem caused
			// by the old open-tray-on-mouse-down method:
			//MButton::Send {RButton down}
			//MButton up::Send {RButton up}
			g_script.mTrayMenu->Display(false);
			return 0;
		} // Inner switch()
		break;
	} // case AHK_NOTIFYICON

	case AHK_DIALOG:  // User defined msg sent from our functions MsgBox() or FileSelectFile().
	{
		// Always call this to close the clipboard if it was open (e.g. due to a script
		// line such as "MsgBox, %clipboard%" that got us here).  Seems better just to
		// do this rather than incurring the delay and overhead of a MsgSleep() call:
		CLOSE_CLIPBOARD_IF_OPEN;
		
		// Ensure that the app's top-most window (the modal dialog) is the system's
		// foreground window.  This doesn't use FindWindow() since it can hang in rare
		// cases.  And GetActiveWindow, GetTopWindow, GetWindow, etc. don't seem appropriate.
		// So EnumWindows is probably the way to do it:
		HWND top_box = FindOurTopDialog();
		if (top_box)
		{

			// v1.0.33: The following is probably reliable since the AHK_DIALOG should
			// be in front of any messages that would launch an interrupting thread.  In other
			// words, the "g" struct should still be the one that owns this MsgBox/dialog window.
			g->DialogHWND = top_box; // This is used to work around an AHK_TIMEOUT issue in which a MsgBox that has only an OK button fails to deliver the Timeout indicator to the script.

			SetForegroundWindowEx(top_box);

			// Setting the big icon makes AutoHotkey dialogs more distinct in the Alt-tab menu.
			// Unfortunately, it seems that setting the big icon also indirectly sets the small
			// icon, or more precisely, that the dialog simply scales the large icon whenever
			// a small one isn't available.  This results in the FileSelectFile dialog's title
			// being initially messed up (at least on WinXP) and also puts an unwanted icon in
			// the title bar of each MsgBox.  So for now it's disabled:
			//LPARAM main_icon = (LPARAM)LoadImage(g_hInstance, MAKEINTRESOURCE(IDI_MAIN), IMAGE_ICON, 0, 0, LR_SHARED);
			//SendMessage(top_box, WM_SETICON, ICON_BIG, main_icon);
			//SendMessage(top_box, WM_SETICON, ICON_SMALL, 0);  // Tried this to get rid of it, but it didn't help.
			// But don't set the small one, because it reduces the area available for title text
			// without adding any significant benefit:
			//SendMessage(top_box, WM_SETICON, ICON_SMALL, main_icon);

			UINT timeout = (UINT)lParam;  // Caller has ensured that this is non-negative.
			if (timeout)
				// Caller told us to establish a timeout for this modal dialog (currently always MessageBox).
				// In addition to any other reasons, the first param of the below must not be NULL because
				// that would cause the 2nd param to be ignored.  We want the 2nd param to be the actual
				// ID assigned to this timer.
				SetTimer(top_box, g_nMessageBoxes, (UINT)timeout, MsgBoxTimeout);
		}
		// else: if !top_box: no error reporting currently.
		return 0;
	}

	case AHK_USER_MENU:
		// Search for AHK_USER_MENU in GuiWindowProc() for comments about why this is done:
		PostMessage(hWnd, iMsg, wParam, lParam);
		MsgSleep(-1, RETURN_AFTER_MESSAGES_SPECIAL_FILTER);
		return 0;

	case WM_HOTKEY: // As a result of this app having previously called RegisterHotkey().
	case AHK_HOOK_HOTKEY:  // Sent from this app's keyboard or mouse hook.
	case AHK_HOTSTRING: // Added for v1.0.36.02 so that hotstrings work even while an InputBox or other non-standard msg pump is running.
	case AHK_CLIPBOARD_CHANGE: // Added for v1.0.44 so that clipboard notifications aren't lost while the script is displaying a MsgBox or other dialog.
		// If the following facts are ever confirmed, there would be no need to post the message in cases where
		// the MsgSleep() won't be done:
		// 1) The mere fact that any of the above messages has been received here in MainWindowProc means that a
		//    message pump other than our own main one is running (i.e. it is the closest pump on the call stack).
		//    This is because our main message pump would never have dispatched the types of messages above because
		//    it is designed to fully handle then discard them.
		// 2) All of these types of non-main message pumps would discard a message with a NULL hwnd.
		//
		// One source of confusion is that there are quite a few different types of message pumps that might
		// be running:
		// - InputBox/MsgBox, or other dialog
		// - Popup menu (tray menu, popup menu from Menu command, or context menu of an Edit/MonthCal, including
		//   our main window's edit control g_hWndEdit).
		// - Probably others, such as ListView marquee-drag, that should be listed here as they are
		//   remembered/discovered.
		//
		// Due to maintainability and the uncertainty over backward compatibility (see comments above), the
		// following message is posted even when INTERRUPTIBLE==false.
		// Post it with a NULL hwnd (update: also for backward compatibility) to avoid any chance that our
		// message pump will dispatch it back to us.  We want these events to always be handled there,
		// where almost all new quasi-threads get launched.  Update: Even if it were safe in terms of
		// backward compatibility to change NULL to gHwnd, testing shows it causes problems when a hotkey
		// is pressed while one of the script's menus is displayed (at least a menu bar).  For example:
		// *LCtrl::Send {Blind}{Ctrl up}{Alt down}
		// *LCtrl up::Send {Blind}{Alt up}
		PostMessage(NULL, iMsg, wParam, lParam);
		if (IsInterruptible())
			MsgSleep(-1, RETURN_AFTER_MESSAGES_SPECIAL_FILTER);
		//else let the other pump discard this hotkey event since in most cases it would do more harm than good
		// (see comments above for why the message is posted even when it is 90% certain it will be discarded
		// in all cases where MsgSleep isn't done).
		return 0;

	case WM_TIMER:
		// MSDN: "When you specify a TimerProc callback function, the default window procedure calls
		// the callback function when it processes WM_TIMER. Therefore, you need to dispatch messages
		// in the calling thread, even when you use TimerProc instead of processing WM_TIMER."
		// MSDN CONTRADICTION: "You can process the message by providing a WM_TIMER case in the window
		// procedure. Otherwise, DispatchMessage will call the TimerProc callback function specified in
		// the call to the SetTimer function used to install the timer."
		// In light of the above, it seems best to let the default proc handle this message if it
		// has a non-NULL lparam:
		if (lParam)
			break;
		// Otherwise, it's the main timer, which is the means by which joystick hotkeys and script timers
		// created via the script command "SetTimer" continue to execute even while a dialog's message pump
		// is running.  Even if the script is NOT INTERRUPTIBLE (which generally isn't possible, since
		// the mere fact that we're here means that a dialog's message pump dispatched a message to us
		// [since our msg pump would not dispatch this type of msg], which in turn means that the script
		// should be interruptible due to DIALOG_PREP), call MsgSleep() anyway so that joystick
		// hotkeys will be polled.  If any such hotkeys are "newly down" right now, those events queued
		// will be buffered/queued for later, when the script becomes interruptible again.  Also, don't
		// call CheckScriptTimers() or PollJoysticks() directly from here.  See comments at the top of
		// those functions for why.
		// This is an older comment, but I think it might still apply, which is why MsgSleep() is not
		// called when a popup menu or a window's main menu is visible.  We don't really want to run the
		// script's timed subroutines or monitor joystick hotkeys while a menu is displayed anyway:
		// Do not call MsgSleep() while a popup menu is visible because that causes long delays
		// sometime when the user is trying to select a menu (the user's click is ignored and the menu
		// stays visible).  I think this is because MsgSleep()'s PeekMessage() intercepts the user's
		// clicks and is unable to route them to TrackPopupMenuEx()'s message loop, which is the only
		// place they can be properly processed.  UPDATE: This also needs to be done when the MAIN MENU
		// is visible, because testing shows that that menu would otherwise become sluggish too, perhaps
		// more rarely, when timers are running.
		// Other background info:
		// Checking g_MenuIsVisible here prevents timed subroutines from running while the tray menu
		// or main menu is in use.  This is documented behavior, and is desirable most of the time
		// anyway.  But not to do this would produce strange effects because any timed subroutine
		// that took a long time to run might keep us away from the "menu loop", which would result
		// in the menu becoming temporarily unresponsive while the user is in it (and probably other
		// undesired effects).
		if (!g_MenuIsVisible)
			MsgSleep(-1, RETURN_AFTER_MESSAGES_SPECIAL_FILTER);
		return 0;

	case WM_SYSCOMMAND:
		if ((wParam == SC_CLOSE || wParam == SC_MINIMIZE) && hWnd == g_hWnd) // i.e. behave this way only for main window.
		{
			// The user has either clicked the window's "X" button, chosen "Close"
			// from the system (upper-left icon) menu, or pressed Alt-F4.  In all
			// these cases, we want to hide the window rather than actually closing
			// it.  If the user really wishes to exit the program, a File->Exit
			// menu option may be available, or use the Tray Icon, or launch another
			// instance which will close the previous, etc.  UPDATE: SC_MINIMIZE is
			// now handled this way also so that owned windows (such as Splash and
			// Progress) won't be hidden when the main window is hidden.
			ShowWindow(g_hWnd, SW_HIDE);
			return 0;
		}
		break;

	case WM_CLOSE:
		if (hWnd == g_hWnd) // i.e. not the SplashText window or anything other than the main.
		{
			// Receiving this msg is fairly unusual since SC_CLOSE is intercepted and redefined above.
			// However, it does happen if an external app is asking us to close, such as another
			// instance of this same script during the Reload command.  So treat it in a way similar
			// to the user having chosen Exit from the menu.
			//
			// Leave it up to ExitApp() to decide whether to terminate based upon whether
			// there is an OnExit subroutine, whether that subroutine is already running at
			// the time a new WM_CLOSE is received, etc.  It's also its responsibility to call
			// DestroyWindow() upon termination so that the WM_DESTROY message winds up being
			// received and process in this function (which is probably necessary for a clean
			// termination of the app and all its windows):
			g_script.ExitApp(EXIT_CLOSE);
			return 0;  // Verified correct.
		}
		// Otherwise, some window of ours other than our main window was destroyed
		// (perhaps the splash window).  Let DefWindowProc() handle it:
		break;

	case WM_ENDSESSION: // MSDN: "A window receives this message through its WindowProc function."
		if (wParam) // The session is being ended.
			g_script.ExitApp((lParam & ENDSESSION_LOGOFF) ? EXIT_LOGOFF : EXIT_SHUTDOWN);
		//else a prior WM_QUERYENDSESSION was aborted; i.e. the session really isn't ending.
		return 0;  // Verified correct.

	case AHK_EXIT_BY_RELOAD:
		g_script.ExitApp(EXIT_RELOAD);
		return 0; // Whether ExitApp() terminates depends on whether there's an OnExit subroutine and what it does.

	case AHK_EXIT_BY_SINGLEINSTANCE:
		g_script.ExitApp(EXIT_SINGLEINSTANCE);
		return 0; // Whether ExitApp() terminates depends on whether there's an OnExit subroutine and what it does.

	case WM_DESTROY:
		if (hWnd == g_hWnd) // i.e. not the SplashText window or anything other than the main.
		{
			if (!g_DestroyWindowCalled)
				// This is done because I believe it's possible for a WM_DESTROY message to be received
				// even though we didn't call DestroyWindow() ourselves (e.g. via DefWindowProc() receiving
				// and acting upon a WM_CLOSE or us calling DestroyWindow() directly) -- perhaps the window
				// is being forcibly closed or something else abnormal happened.  Make a best effort to run
				// the OnExit subroutine, if present, even without a main window (testing on an earlier
				// versions shows that most commands work fine without the window). For EXIT_DESTROY,
				// it always terminates after running the OnExit subroutine:
				g_script.ExitApp(EXIT_DESTROY);
			// Do not do PostQuitMessage() here because we don't know the proper exit code.
			// MSDN: "The exit value returned to the system must be the wParam parameter of
			// the WM_QUIT message."
			// If we're here, it means our thread called DestroyWindow() directly or indirectly
			// (currently, it's only called directly).  By returning, our thread should resume
			// execution at the statement after DestroyWindow() in whichever caller called that:
			return 0;  // "If an application processes this message, it should return zero."
		}
		// Otherwise, some window of ours other than our main window was destroyed
		// (perhaps the splash window).  Let DefWindowProc() handle it:
		break;

	case WM_CREATE:
		// MSDN: If an application processes this message, it should return zero to continue
		// creation of the window. If the application returns 1, the window is destroyed and
		// the CreateWindowEx or CreateWindow function returns a NULL handle.
		return 0;

	case WM_ERASEBKGND:
	case WM_CTLCOLORSTATIC:
	case WM_PAINT:
	case WM_SIZE:
	{
		if (iMsg == WM_SIZE)
		{
			if (hWnd == g_hWnd)
			{
				if (wParam == SIZE_MINIMIZED)
					// Minimizing the main window hides it.
					ShowWindow(g_hWnd, SW_HIDE);
				else
					MoveWindow(g_hWndEdit, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
				return 0; // The correct return value for this msg.
			}
			if (hWnd == g_hWndSplash || wParam == SIZE_MINIMIZED)
				break;  // Let DefWindowProc() handle it for splash window and Progress windows.
		}
		else
			if (hWnd == g_hWnd || hWnd == g_hWndSplash)
				break; // Let DWP handle it.

		for (i = 0; i < MAX_SPLASHIMAGE_WINDOWS; ++i)
			if (g_SplashImage[i].hwnd == hWnd)
				break;
		bool is_splashimage = (i < MAX_SPLASHIMAGE_WINDOWS);
		if (!is_splashimage)
		{
			for (i = 0; i < MAX_PROGRESS_WINDOWS; ++i)
				if (g_Progress[i].hwnd == hWnd)
					break;
			if (i == MAX_PROGRESS_WINDOWS) // It's not a progress window either.
				// Let DefWindowProc() handle it (should probably never happen since currently the only
				// other type of window is SplashText, which never receive this msg?)
				break;
		}

		SplashType &splash = is_splashimage ? g_SplashImage[i] : g_Progress[i];
		RECT client_rect;

		switch (iMsg)
		{
		case WM_SIZE:
		{
			// Allow any width/height to be specified so that the window can be "rolled up" to its title bar.
			int new_width = LOWORD(lParam);
			int new_height = HIWORD(lParam);
			if (new_width != splash.width || new_height != splash.height)
			{
				GetClientRect(splash.hwnd, &client_rect);
				int control_width = client_rect.right - (splash.margin_x * 2);
				SPLASH_CALC_YPOS
				// The Y offset for each control should match those used in Splash():
				if (new_width != splash.width)
				{
					if (splash.hwnd_text1) // This control doesn't exist if the main text was originally blank.
						MoveWindow(splash.hwnd_text1, PROGRESS_MAIN_POS, FALSE);
					if (splash.hwnd_bar)
						MoveWindow(splash.hwnd_bar, PROGRESS_BAR_POS, FALSE);
					splash.width = new_width;
				}
				// Move the window EVEN IF new_height == splash.height because otherwise the text won't
				// get re-centered when only the width of the window changes:
				MoveWindow(splash.hwnd_text2, PROGRESS_SUB_POS, FALSE); // Negative height seems handled okay.
				// Specifying true for "repaint" in MoveWindow() is not always enough refresh the text correctly,
				// so this is done instead:
				InvalidateRect(splash.hwnd, &client_rect, TRUE);
				// If the user resizes the window, have that size retained (remembered) until the script
				// explicitly changes it or the script destroys the window.
				splash.height = new_height;
			}
			return 0;  // i.e. completely handled here.
		}
		case WM_CTLCOLORSTATIC:
			if (!splash.hbrush && splash.color_text == CLR_DEFAULT) // Let DWP handle it.
				break;
			// Since we're processing this msg and not DWP, must set background color unconditionally,
			// otherwise plain white will likely be used:
			SetBkColor((HDC)wParam, splash.hbrush ? splash.color_bk : GetSysColor(COLOR_BTNFACE));
			if (splash.color_text != CLR_DEFAULT)
				SetTextColor((HDC)wParam, splash.color_text);
			// Always return a real HBRUSH so that Windows knows we altered the HDC for it to use:
			return (LRESULT)(splash.hbrush ? splash.hbrush : GetSysColorBrush(COLOR_BTNFACE));
		case WM_ERASEBKGND:
		{
			if (splash.pic_bmp) // And since there is a pic, its object_width/height should already be valid.
			{
				int ypos = splash.margin_y + (splash.text1_height ? (splash.text1_height + splash.margin_y) : 0);
				if (splash.pic_type == IMAGE_BITMAP)
				{
					HDC hdc = CreateCompatibleDC((HDC)wParam);
					HBITMAP hbmpOld = (HBITMAP)SelectObject(hdc, splash.pic_bmp);
					BitBlt((HDC)wParam, splash.margin_x, ypos, splash.object_width, splash.object_height, hdc, 0, 0, SRCCOPY);
					SelectObject(hdc, hbmpOld);
					DeleteDC(hdc);
				}
				else // IMAGE_ICON
					DrawIconEx((HDC)wParam, splash.margin_x, ypos, splash.pic_icon, splash.object_width, splash.object_height, 0, NULL, DI_NORMAL);
				// Prevent "flashing" by erasing only the part that hasn't already been drawn:
				ExcludeClipRect((HDC)wParam, splash.margin_x, ypos, splash.margin_x + splash.object_width
					, ypos + splash.object_height);
				HRGN hrgn = CreateRectRgn(0, 0, 1, 1);
				GetClipRgn((HDC)wParam, hrgn);
				FillRgn((HDC)wParam, hrgn, splash.hbrush ? splash.hbrush : GetSysColorBrush(COLOR_BTNFACE));
				DeleteObject(hrgn);
				return 1; // "An application should return nonzero if it erases the background."
			}
			// Otherwise, it's a Progress window (or a SplashImage window with no picture):
			if (!splash.hbrush) // Let DWP handle it.
				break;
			RECT clipbox;
			GetClipBox((HDC)wParam, &clipbox);
			FillRect((HDC)wParam, &clipbox, splash.hbrush);
			return 1; // "An application should return nonzero if it erases the background."
		}
		} // switch()
		break; // Let DWP handle it.
	}
		
	case WM_SETFOCUS:
		if (hWnd == g_hWnd)
		{
			SetFocus(g_hWndEdit);  // Always focus the edit window, since it's the only navigable control.
			return 0;
		}
		break;

	case WM_CLIPBOARDUPDATE: // For Vista and later.
	case WM_DRAWCLIPBOARD:
		if (g_script.mOnClipboardChangeLabel || g_script.mOnClipboardChange.Count()) // In case it's a bogus msg, it's our responsibility to avoid posting the msg if there's no label to launch.
			PostMessage(g_hWnd, AHK_CLIPBOARD_CHANGE, !g_script.mIsReadyToExecute, 0); // It's done this way to buffer it when the script is uninterruptible, etc.  v1.0.44: Post to g_hWnd vs. NULL so that notifications aren't lost when script is displaying a MsgBox or other dialog.
		if (g_script.mNextClipboardViewer) // Will be NULL if there are no other windows in the chain, or if we're on Vista or later and used AddClipboardFormatListener instead of SetClipboardViewer (in which case iMsg should be WM_CLIPBOARDUPDATE).
			SendMessageTimeout(g_script.mNextClipboardViewer, iMsg, wParam, lParam, SMTO_ABORTIFHUNG, 2000, &dwTemp);
		return 0;

	case WM_CHANGECBCHAIN:
		// MSDN: If the next window is closing, repair the chain. 
		if ((HWND)wParam == g_script.mNextClipboardViewer)
			g_script.mNextClipboardViewer = (HWND)lParam;
		// MSDN: Otherwise, pass the message to the next link. 
		else if (g_script.mNextClipboardViewer)
			SendMessageTimeout(g_script.mNextClipboardViewer, iMsg, wParam, lParam, SMTO_ABORTIFHUNG, 2000, &dwTemp);
		return 0;

	case AHK_GETWINDOWTEXT:
		// It's best to handle this msg here rather than in the main event loop in case a non-standard message
		// pump is running (such as MsgBox's), in which case this msg would be dispatched directly here.
		if (IsWindow((HWND)lParam)) // In case window has been destroyed since msg was posted.
			GetWindowText((HWND)lParam, (LPTSTR )wParam, KEY_HISTORY_WINDOW_TITLE_SIZE);
		// Probably best not to do the following because it could result in such "low priority" messages
		// getting out of step with each other, and could also introduce KeyHistory WinTitle "lag":
		// Could give low priority to AHK_GETWINDOWTEXT under the theory that sometimes the call takes a long
		// time to return: Upon receipt of such a message, repost it whenever Peek(specific_msg_range, PM_NOREMOVE)
		// detects a thread-starting event on the queue.  However, Peek might be a high overhead call in some cases,
		// such as when/if it yields our timeslice upon returning FALSE (uncertain/unlikely, but in any case
		// it might do more harm than good).
		return 0;

	case AHK_RETURN_PID:
		// This is obsolete in light of WinGet's support for fetching the PID of any window.
		// But since it's simple, it is retained for backward compatibility.
		// Rajat wanted this so that it's possible to discover the PID based on the title of each
		// script's main window (i.e. if there are multiple scripts running).  Also note that this
		// msg can be sent via TRANSLATE_AHK_MSG() to prevent it from ever being filtered out (and
		// thus delayed) while the script is uninterruptible.  For example:
		// SendMessage, 0x44, 1029,,, %A_ScriptFullPath% - AutoHotkey
		// SendMessage, 1029,,,, %A_ScriptFullPath% - AutoHotkey  ; Same as above but not sent via TRANSLATE.
		return GetCurrentProcessId(); // Don't use ReplyMessage because then our thread can't reply to itself with this answer.

	case AHK_HOT_IF_EVAL: // HotCriterionAllowsFiring uses this to ensure expressions are evaluated only on the main thread.
		// Ensure wParam is a valid criterion (might prevent shatter attacks):
		for (HotkeyCriterion *cp = g_FirstHotExpr; cp; cp = cp->NextCriterion)
			if ((WPARAM)cp == wParam)
				return cp->Eval((LPTSTR)lParam);
		return 0;

	case WM_MEASUREITEM: // L17: Measure menu icon. Not used on Windows Vista or later.
		if (hWnd == g_hWnd && wParam == 0 && !g_os.IsWinVistaOrLater())
			if (UserMenu::OwnerMeasureItem((LPMEASUREITEMSTRUCT)lParam))
				return TRUE;
		break;

	case WM_DRAWITEM: // L17: Draw menu icon. Not used on Windows Vista or later.
		if (hWnd == g_hWnd && wParam == 0 && !g_os.IsWinVistaOrLater())
			if (UserMenu::OwnerDrawItem((LPDRAWITEMSTRUCT)lParam))
				return TRUE;
		break;

	case WM_ENTERMENULOOP:
		CheckMenuItem(GetMenu(g_hWnd), ID_FILE_PAUSE, g->IsPaused ? MF_CHECKED : MF_UNCHECKED); // This is the menu bar in the main window; the tray menu's checkmark is updated only when the tray menu is actually displayed.
		if (!g_MenuIsVisible) // See comments in similar code in GuiWindowProc().
			g_MenuIsVisible = MENU_TYPE_BAR;
		break;
	case WM_EXITMENULOOP:
		g_MenuIsVisible = MENU_TYPE_NONE; // See comments in similar code in GuiWindowProc().
		break;

#ifdef CONFIG_DEBUGGER
	case AHK_CHECK_DEBUGGER:
		// This message is sent when data arrives on the debugger's socket.  It allows the
		// debugger to respond to commands which are sent while the script is sleeping or
		// waiting for messages.
		if (g_Debugger.IsConnected() && (g_Debugger.HasPendingCommand() || LOWORD(lParam) == FD_CLOSE))
			g_Debugger.ProcessCommands();
		break;
#endif

	default:
		// The following iMsg can't be in the switch() since it's not constant:
		if (iMsg == WM_TASKBARCREATED && !g_NoTrayIcon) // !g_NoTrayIcon --> the tray icon should be always visible.
		{
			g_script.CreateTrayIcon();
			g_script.UpdateTrayIcon(true);  // Force the icon into the correct pause, suspend, or mIconFrozen state.
			// And now pass this iMsg on to DefWindowProc() in case it does anything with it.
		}
		
#ifdef CONFIG_DEBUGGER
		static UINT sAttachDebuggerMessage = RegisterWindowMessage(_T("AHK_ATTACH_DEBUGGER"));
		if (iMsg == sAttachDebuggerMessage && !g_Debugger.IsConnected())
		{
			char dbg_host[16] = "localhost"; // IPv4 max string len
			char dbg_port[6] = "9000";

			if (wParam)
			{	// Convert 32-bit address to string for Debugger::Connect().
				in_addr addr;
				addr.S_un.S_addr = (ULONG)wParam;
				char *tmp = inet_ntoa(addr);
				if (tmp)
					strcpy(dbg_host, tmp);
			}
			if (lParam)
				// Convert 16-bit port number to string for Debugger::Connect().
				_itoa(LOWORD(lParam), dbg_port, 10);

			if (g_Debugger.Connect(dbg_host, dbg_port) == DEBUGGER_E_OK)
				g_Debugger.Break();
		}
#endif

	} // switch()

	return DefWindowProc(hWnd, iMsg, wParam, lParam);
}



bool FindAutoHotkeyUtilSub(LPTSTR aBuf, int aBufSize, LPTSTR aFile, LPTSTR aDir)
{
	int len = sntprintf(aBuf, aBufSize, _T("\"%s\\%s"), aDir, aFile);
	if (len + 1 > aBufSize // Too long. Should realistically never happen.
		|| GetFileAttributes(aBuf + 1) == INVALID_FILE_ATTRIBUTES) // File not found.
		return false;
	aBuf[len++] = '"';
	aBuf[len] = '\0';
	return true;
}

bool FindAutoHotkeyUtil(LPTSTR aBuf, int aBufSize, LPTSTR aFile, LPTSTR aInstallDirBuf, LPTSTR &aUtilDir)
{
	// Always try our directory first, in case it has different utils to the installed version.
	// ActionExec()'s CreateProcess() is currently done in a way that prefers enclosing double quotes:
	if (!FindAutoHotkeyUtilSub(aBuf, aBufSize, aFile, g_script.mOurEXEDir))
	{
		// Try GetAHKInstallDir() so that compiled scripts running on machines that happen
		// to have AHK installed will still be able to fetch the help file and Window Spy:
		if (   !GetAHKInstallDir(aInstallDirBuf)
			|| !FindAutoHotkeyUtilSub(aBuf, aBufSize, aFile, aInstallDirBuf)   )
			return false;
		aUtilDir = aInstallDirBuf;
	}
	else
		aUtilDir = g_script.mOurEXEDir;
	return true;
}

bool LaunchAutoHotkeyUtil(LPTSTR aFile, bool aIsScript)
{
	TCHAR buf_file[2048], buf_exe[2048], installdir[MAX_PATH];
	LPTSTR utildir, file = buf_file, args = _T(""); // Use "" vs. NULL to specify that there are no params at all.
	if (!FindAutoHotkeyUtil(buf_file, _countof(buf_file), aFile, installdir, utildir))
		return false;
	if (aIsScript)
	{
		// Always try AutoHotkey.exe in the same directory as the util first, if present,
		// since mOurEXE could be a different version of AutoHotkey (or a compiled script).
		if (FindAutoHotkeyUtilSub(buf_exe, _countof(buf_exe), _T("AutoHotkey.exe"), utildir))
			file = buf_exe, args = buf_file;
#ifndef AUTOHOTKEYSC
		else if (utildir == g_script.mOurEXEDir)
			// Use our EXE only if the util was found in our directory.
			file = g_script.mOurEXE, args = buf_file;
#endif
		//else: AutoHotkey appears to be installed but missing AutoHotkey.exe.
		// Try running the .ahk file directly in the off chance that it is registered
		// with some other EXE name.
	}
	// Attempt to run the file:
	return g_script.ActionExec(file, args, NULL, false) != FAIL;
}

void LaunchWindowSpy()
{
	if (   !LaunchAutoHotkeyUtil(_T("WindowSpy.ahk"), true)
		&& !LaunchAutoHotkeyUtil(_T("AU3_Spy.exe"), false)   )
		MsgBox(_T("Could not launch WindowSpy.ahk or AU3_Spy.exe"), MB_ICONERROR);
}

void LaunchAutoHotkeyHelp()
{
	if (   !LaunchAutoHotkeyUtil(AHK_HELP_FILE, false)   )
		MsgBox(_T("Could not launch ") AHK_HELP_FILE, MB_ICONERROR);
}

bool HandleMenuItem(HWND aHwnd, WORD aMenuItemID, HWND aGuiHwnd)
// See if an item was selected from the tray menu or main menu.  Note that it is possible
// for one of the standard menu items to be triggered from a GUI menu if the menu or one of
// its submenus was modified with the "menu, MenuName, Standard" command.
// Returns true if the message is fully handled here, false otherwise.
{
	switch (aMenuItemID)
	{
	case ID_TRAY_OPEN:
		ShowMainWindow();
		return true;
	case ID_TRAY_EDITSCRIPT:
	case ID_FILE_EDITSCRIPT:
		g_script.Edit();
		return true;
	case ID_TRAY_RELOADSCRIPT:
	case ID_FILE_RELOADSCRIPT:
		if (!g_script.Reload(false))
			MsgBox(_T("The script could not be reloaded."));
		return true;
	case ID_TRAY_WINDOWSPY:
	case ID_FILE_WINDOWSPY:
		LaunchWindowSpy();
		return true;
	case ID_TRAY_HELP:
	case ID_HELP_USERMANUAL:
		LaunchAutoHotkeyHelp();
		return true;
	case ID_TRAY_SUSPEND:
	case ID_FILE_SUSPEND:
		Line::ToggleSuspendState();
		return true;
	case ID_TRAY_PAUSE:
	case ID_FILE_PAUSE:
		if (g->IsPaused)
			--g_nPausedThreads;
		else
			++g_nPausedThreads; // For this purpose the idle thread is counted as a paused thread.
		g->IsPaused = !g->IsPaused;
		g_script.UpdateTrayIcon();
		return true;
	case ID_TRAY_EXIT:
	case ID_FILE_EXIT:
		g_script.ExitApp(EXIT_MENU);  // More reliable than PostQuitMessage(), which has been known to fail in rare cases.
		return true; // If there is an OnExit subroutine, the above might not actually exit.
	case ID_VIEW_LINES:
		ShowMainWindow(MAIN_MODE_LINES);
		return true;
	case ID_VIEW_VARIABLES:
		ShowMainWindow(MAIN_MODE_VARS);
		return true;
	case ID_VIEW_HOTKEYS:
		ShowMainWindow(MAIN_MODE_HOTKEYS);
		return true;
	case ID_VIEW_KEYHISTORY:
		ShowMainWindow(MAIN_MODE_KEYHISTORY);
		return true;
	case ID_VIEW_REFRESH:
		ShowMainWindow(MAIN_MODE_REFRESH);
		return true;
	case ID_HELP_WEBSITE:
		if (!g_script.ActionExec(_T(AHK_WEBSITE), _T(""), NULL, false))
			MsgBox(_T("Could not open URL ") _T(AHK_WEBSITE) _T(" in default browser."));
		return true;
	default:
		// See if this command ID is one of the user's custom menu items.  Due to the possibility
		// that some items have been deleted from the menu, can't rely on comparing
		// aMenuItemID to g_script.mMenuItemCount in any way.  Just look up the ID to make sure
		// there really is a menu item for it:
		if (!g_script.FindMenuItemByID(aMenuItemID)) // Do nothing, let caller try to handle it some other way.
			return false;
		// It seems best to treat the selection of a custom menu item in a way similar
		// to how hotkeys are handled by the hook. See comments near the definition of
		// POST_AHK_USER_MENU for more details.
		POST_AHK_USER_MENU(aHwnd, aMenuItemID, (WPARAM)aGuiHwnd) // Send the menu's cmd ID and the window index (index is safer than pointer, since pointer might get deleted).
		// Try to maintain a list here of all the ways the script can be uninterruptible
		// at this moment in time, and whether that uninterruptibility should be overridden here:
		// 1) YES: g_MenuIsVisible is true (which in turn means that the script is marked
		//    uninterruptible to prevent timed subroutines from running and possibly
		//    interfering with menu navigation): Seems impossible because apparently 
		//    the WM_RBUTTONDOWN must first be returned from before we're called directly
		//    with the WM_COMMAND message corresponding to the menu item chosen by the user.
		//    In other words, g_MenuIsVisible will be false and the script thus will
		//    not be uninterruptible, at least not solely for that reason.
		// 2) YES: A new hotkey or timed subroutine was just launched and it's still in its
		//    grace period.  In this case, ExecUntil()'s call of PeekMessage() every 10ms
		//    or so will catch the item we just posted.  But it seems okay to interrupt
		//    here directly in most such cases.  InitNewThread(): Newly launched
		//    timed subroutine or hotkey subroutine.
		// 3) YES: Script is engaged in an uninterruptible activity such as SendKeys().  In this
		//    case, since the user has managed to get the tray menu open, it's probably
		//    best to process the menu item with the same priority as if any other menu
		//    item had been selected, interrupting even a critical operation since that's
		//    probably what the user would want.  SLEEP_WITHOUT_INTERRUPTION: SendKeys,
		//    Mouse input, Clipboard open, SetForegroundWindowEx().
		// 4) YES: AutoExecSection(): Since its grace period is only 100ms, doesn't seem to be
		//    a problem.  In any case, the timer would fire and briefly interrupt the menu
		//    subroutine we're trying to launch here even if a menu item were somehow
		//    activated in the first 100ms.
		//
		// IN LIGHT OF THE ABOVE, it seems best not to do the below.  In addition, the msg
		// filtering done by MsgSleep when the script is uninterruptible now excludes the
		// AHK_USER_MENU message (i.e. that message is always retrieved and acted upon,
		// even when the script is uninterruptible):
		//if (!INTERRUPTIBLE)
		//	return true;  // Leave the message buffered until later.
		// Now call the main loop to handle the message we just posted (and any others):
		return true;
	} // switch()
	return false;  // Indicate that the message was NOT handled.
}



ResultType ShowMainWindow(MainWindowModes aMode, bool aRestricted)
// Always returns OK for caller convenience.
{
	// v1.0.30.05: Increased from 32 KB to 64 KB, which is the maximum size of an Edit
	// in Win9x:
	TCHAR buf_temp[65534];  // Formerly 32767.
	*buf_temp = '\0';
	bool jump_to_bottom = false;  // Set default behavior for edit control.
	static MainWindowModes current_mode = MAIN_MODE_NO_CHANGE;

#ifdef AUTOHOTKEYSC
	// If we were called from a restricted place, such as via the Tray Menu or the Main Menu,
	// don't allow potentially sensitive info such as script lines and variables to be shown.
	// This is done so that scripts can be compiled more securely, making it difficult for anyone
	// to use ListLines to see the author's source code.  Rather than make exceptions for things
	// like KeyHistory, it seems best to forbid all information reporting except in cases where
	// existing info in the main window -- which must have gotten their via an allowed command
	// such as ListLines encountered in the script -- is being refreshed.  This is because in
	// that case, the script author has given de facto permission for that loophole (and it's
	// a pretty small one, not easy to exploit):
	if (aRestricted && !g_AllowMainWindow && (current_mode == MAIN_MODE_NO_CHANGE || aMode != MAIN_MODE_REFRESH))
	{
		SendMessage(g_hWndEdit, WM_SETTEXT, 0, (LPARAM)
			_T("Script info will not be shown because the \"Menu, Tray, MainWindow\"\r\n")
			_T("command option was not enabled in the original script."));
		return OK;
	}
#endif

	// If the window is empty, caller wants us to default it to showing the most recently
	// executed script lines:
	if (current_mode == MAIN_MODE_NO_CHANGE && (aMode == MAIN_MODE_NO_CHANGE || aMode == MAIN_MODE_REFRESH))
		aMode = MAIN_MODE_LINES;

	switch (aMode)
	{
	// case MAIN_MODE_NO_CHANGE: do nothing
	case MAIN_MODE_LINES:
		Line::LogToText(buf_temp, _countof(buf_temp));
		jump_to_bottom = true;
		break;
	case MAIN_MODE_VARS:
		g_script.ListVars(buf_temp, _countof(buf_temp));
		break;
	case MAIN_MODE_HOTKEYS:
		Hotkey::ListHotkeys(buf_temp, _countof(buf_temp));
		break;
	case MAIN_MODE_KEYHISTORY:
		g_script.ListKeyHistory(buf_temp, _countof(buf_temp));
		break;
	case MAIN_MODE_REFRESH:
		// Rather than do a recursive call to self, which might stress the stack if the script is heavily recursed:
		switch (current_mode)
		{
		case MAIN_MODE_LINES:
			Line::LogToText(buf_temp, _countof(buf_temp));
			jump_to_bottom = true;
			break;
		case MAIN_MODE_VARS:
			g_script.ListVars(buf_temp, _countof(buf_temp));
			break;
		case MAIN_MODE_HOTKEYS:
			Hotkey::ListHotkeys(buf_temp, _countof(buf_temp));
			break;
		case MAIN_MODE_KEYHISTORY:
			g_script.ListKeyHistory(buf_temp, _countof(buf_temp));
			// Special mode for when user refreshes, so that new keys can be seen without having
			// to scroll down again:
			jump_to_bottom = true;
			break;
		}
		break;
	}

	if (aMode != MAIN_MODE_REFRESH && aMode != MAIN_MODE_NO_CHANGE)
		current_mode = aMode;

	// Update the text before displaying the window, since it might be a little less disruptive
	// and might also be quicker if the window is hidden or non-foreground.
	// Unlike SetWindowText(), this method seems to expand tab characters:
	if (aMode != MAIN_MODE_NO_CHANGE)
		SendMessage(g_hWndEdit, WM_SETTEXT, 0, (LPARAM)buf_temp);

	if (!IsWindowVisible(g_hWnd))
	{
		ShowWindow(g_hWnd, SW_SHOW);
		if (IsIconic(g_hWnd)) // This happens whenever the window was last hidden via the minimize button.
			ShowWindow(g_hWnd, SW_RESTORE);
	}
	if (g_hWnd != GetForegroundWindow())
		if (!SetForegroundWindow(g_hWnd))
			SetForegroundWindowEx(g_hWnd);  // Only as a last resort, since it uses AttachThreadInput()

	if (jump_to_bottom)
	{
		SendMessage(g_hWndEdit, EM_LINESCROLL , 0, 999999);
		//SendMessage(g_hWndEdit, EM_SETSEL, -1, -1);
		//SendMessage(g_hWndEdit, EM_SCROLLCARET, 0, 0);
	}
	return OK;
}



DWORD GetAHKInstallDir(LPTSTR aBuf)
// Caller must ensure that aBuf is large enough (either by having called this function a previous time
// to get the length, or by making it MAX_PATH in capacity).
// Returns the length of the string (0 if empty).
{
	TCHAR buf[MAX_PATH];
	DWORD length;
#ifdef _WIN64
	// First try 64-bit registry, then 32-bit registry.
	for (DWORD flag = 0; ; flag = KEY_WOW64_32KEY)
#else
	// First try 32-bit registry, then 64-bit registry.
	for (DWORD flag = 0; ; flag = KEY_WOW64_64KEY)
#endif
	{
		length = ReadRegString(HKEY_LOCAL_MACHINE, _T("SOFTWARE\\AutoHotkey"), _T("InstallDir"), buf, MAX_PATH, flag);
		if (length || flag)
			break;
	}
	if (aBuf)
		_tcscpy(aBuf, buf); // v1.0.47: Must be done as a separate copy because passing a size of MAX_PATH for aBuf can crash when aBuf is actually smaller than that (even though it's large enough to hold the string).
	return length;
}



//////////////
// InputBox //
//////////////

ResultType InputBox(Var *aOutputVar, LPTSTR aTitle, LPTSTR aText, bool aHideInput, int aWidth, int aHeight
	, int aX, int aY, double aTimeout, LPTSTR aDefault)
{
	if (g_nInputBoxes >= MAX_INPUTBOXES)
	{
		// Have a maximum to help prevent runaway hotkeys due to key-repeat feature, etc.
		MsgBox(_T("The maximum number of InputBoxes has been reached."));
		return FAIL;
	}
	if (!aOutputVar) return FAIL;
	if (!*aTitle)
		// If available, the script's filename seems a much better title in case the user has
		// more than one script running:
		aTitle = (g_script.mFileName && *g_script.mFileName) ? g_script.mFileName : T_AHK_NAME_VERSION;
	// Limit the size of what we were given to prevent unreasonably huge strings from
	// possibly causing a failure in CreateDialog().  This copying method is always done because:
	// Make a copy of all string parameters, using the stack, because they may reside in the deref buffer
	// and other commands (such as those in timed/hotkey subroutines) maybe overwrite the deref buffer.
	// This is not strictly necessary since InputBoxProc() is called immediately and makes instantaneous
	// and one-time use of these strings (not needing them after that), but it feels safer:
	TCHAR title[DIALOG_TITLE_SIZE];
	TCHAR text[4096];  // Size was increased in light of the fact that dialog can be made larger now.
	TCHAR default_string[4096];
	tcslcpy(title, aTitle, _countof(title));
	tcslcpy(text, aText, _countof(text));
	tcslcpy(default_string, aDefault, _countof(default_string));
	g_InputBox[g_nInputBoxes].title = title;
	g_InputBox[g_nInputBoxes].text = text;
	g_InputBox[g_nInputBoxes].default_string = default_string;

	if (aTimeout > 2147483) // This is approximately the max number of seconds that SetTimer() can handle.
		aTimeout = 2147483;
	if (aTimeout < 0) // But it can be equal to zero to indicate no timeout at all.
		aTimeout = 0.1;  // A value that might cue the user that something is wrong.
	g_InputBox[g_nInputBoxes].timeout = (DWORD)(aTimeout * 1000);  // Convert to ms

	// Allow 0 width or height (hides the window):
	g_InputBox[g_nInputBoxes].width = aWidth != INPUTBOX_DEFAULT ? (aWidth < 0 ? 0 : DPIScale(aWidth)) : INPUTBOX_DEFAULT;
	g_InputBox[g_nInputBoxes].height = aHeight != INPUTBOX_DEFAULT ? (aHeight < 0 ? 0 : DPIScale(aHeight)) : INPUTBOX_DEFAULT;
	g_InputBox[g_nInputBoxes].xpos = aX;  // But seems okay to allow these to be negative, even if absolute coords.
	g_InputBox[g_nInputBoxes].ypos = aY;
	g_InputBox[g_nInputBoxes].output_var = aOutputVar;
	g_InputBox[g_nInputBoxes].password_char = aHideInput ? '*' : '\0';

	// At this point, we know a dialog will be displayed.  See macro's comments for details:
	DIALOG_PREP

	// Specify NULL as the owner window since we want to be able to have the main window in the foreground even
	// if there are InputBox windows.  Update: A GUI window can now be the parent if thread has that setting.
	++g_nInputBoxes;
	INT_PTR result = DialogBox(g_hInstance, MAKEINTRESOURCE(IDD_INPUTBOX), THREAD_DIALOG_OWNER, InputBoxProc);
	--g_nInputBoxes;

	DIALOG_END

	// See the comments in InputBoxProc() for why ErrorLevel is set here rather than there.
	switch(result)
	{
	case AHK_TIMEOUT:
		// In this case the TimerProc already set the output variable to be what the user entered.
		return g_ErrorLevel->Assign(_T("2"));
	case IDOK:
	case IDCANCEL:
		// The output variable is set to whatever the user entered, even if the user pressed
		// the cancel button.  This allows the cancel button to specify that a different
		// operation should be performed on the entered text:
		return g_ErrorLevel->Assign(result == IDCANCEL ? ERRORLEVEL_ERROR : ERRORLEVEL_NONE);
	case -1:
		MsgBox(_T("The InputBox window could not be displayed."));
		// No need to set ErrorLevel since this is a runtime error that will kill the current quasi-thread.
		return FAIL;
	case FAIL:
		return FAIL;
	}

	return OK;
}



INT_PTR CALLBACK InputBoxProc(HWND hWndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
// MSDN:
// Typically, the dialog box procedure should return TRUE if it processed the message,
// and FALSE if it did not. If the dialog box procedure returns FALSE, the dialog
// manager performs the default dialog operation in response to the message.
{
	// See GuiWindowProc() for details about this first part:
	LRESULT msg_reply;
	if (g_MsgMonitor.Count() // Count is checked here to avoid function-call overhead.
		&& (!g->CalledByIsDialogMessageOrDispatch || g->CalledByIsDialogMessageOrDispatchMsg != uMsg) // v1.0.44.11: If called by IsDialog or Dispatch but they changed the message number, check if the script is monitoring that new number.
		&& MsgMonitor(hWndDlg, uMsg, wParam, lParam, NULL, msg_reply))
		return (BOOL)msg_reply; // MsgMonitor has returned "true", indicating that this message should be omitted from further processing.
	g->CalledByIsDialogMessageOrDispatch = false; // v1.0.40.01.

	HWND hControl;

	// Set default array index for g_InputBox[].  Caller has ensured that g_nInputBoxes > 0:
	int target_index = g_nInputBoxes - 1;
	#define CURR_INPUTBOX g_InputBox[target_index]

	switch(uMsg)
	{
	case WM_INITDIALOG:
	{
		// Clipboard may be open if its contents were used to build the text or title
		// of this dialog (e.g. "InputBox, out, %clipboard%").  It's best to do this before
		// anything that might take a relatively long time (e.g. SetForegroundWindowEx()):
		CLOSE_CLIPBOARD_IF_OPEN;

		CURR_INPUTBOX.hwnd = hWndDlg;

		if (CURR_INPUTBOX.password_char)
			SendDlgItemMessage(hWndDlg, IDC_INPUTEDIT, EM_SETPASSWORDCHAR, CURR_INPUTBOX.password_char, 0);

		SetWindowText(hWndDlg, CURR_INPUTBOX.title);
		if (hControl = GetDlgItem(hWndDlg, IDC_INPUTPROMPT))
			SetWindowText(hControl, CURR_INPUTBOX.text);

		// Don't do this check; instead allow the MoveWindow() to occur unconditionally so that
		// the new button positions and such will override those set in the dialog's resource
		// properties:
		//if (CURR_INPUTBOX.width != INPUTBOX_DEFAULT || CURR_INPUTBOX.height != INPUTBOX_DEFAULT
		//	|| CURR_INPUTBOX.xpos != INPUTBOX_DEFAULT || CURR_INPUTBOX.ypos != INPUTBOX_DEFAULT)
		RECT rect;
		GetWindowRect(hWndDlg, &rect);
		int new_width = (CURR_INPUTBOX.width == INPUTBOX_DEFAULT) ? rect.right - rect.left : CURR_INPUTBOX.width;
		int new_height = (CURR_INPUTBOX.height == INPUTBOX_DEFAULT) ? rect.bottom - rect.top : CURR_INPUTBOX.height;

		// If a non-default size was specified, the box will need to be recentered; thus, we can't rely on
		// the dialog's DS_CENTER style in its template.  The exception is when an explicit xpos or ypos is
		// specified, in which case centering is disabled for that dimension.
		int new_xpos, new_ypos;
		if (CURR_INPUTBOX.xpos != INPUTBOX_DEFAULT && CURR_INPUTBOX.ypos != INPUTBOX_DEFAULT)
		{
			new_xpos = CURR_INPUTBOX.xpos;
			new_ypos = CURR_INPUTBOX.ypos;
		}
		else
		{
			POINT pt = CenterWindow(new_width, new_height);
  			if (CURR_INPUTBOX.xpos == INPUTBOX_DEFAULT) // Center horizontally.
				new_xpos = pt.x;
			else
				new_xpos = CURR_INPUTBOX.xpos;
  			if (CURR_INPUTBOX.ypos == INPUTBOX_DEFAULT) // Center vertically.
				new_ypos = pt.y;
			else
				new_ypos = CURR_INPUTBOX.ypos;
		}

		MoveWindow(hWndDlg, new_xpos, new_ypos, new_width, new_height, TRUE);  // Do repaint.
		// This may also needed to make it redraw in some OSes or some conditions:
		GetClientRect(hWndDlg, &rect);  // Not to be confused with GetWindowRect().
		SendMessage(hWndDlg, WM_SIZE, SIZE_RESTORED, rect.right + (rect.bottom<<16));
		
		if (*CURR_INPUTBOX.default_string)
			SetDlgItemText(hWndDlg, IDC_INPUTEDIT, CURR_INPUTBOX.default_string);

		if (hWndDlg != GetForegroundWindow()) // Normally it will be foreground since the template has this property.
			SetForegroundWindowEx(hWndDlg);   // Try to force it to the foreground.

		// Setting the small icon puts it in the upper left corner of the dialog window.
		// Setting the big icon makes the dialog show up correctly in the Alt-Tab menu.
		
		// L17: Use separate big/small icons for best results.
		LPARAM big_icon, small_icon;
		if (g_script.mCustomIcon)
		{
			big_icon = (LPARAM)g_script.mCustomIcon;
			small_icon = (LPARAM)g_script.mCustomIconSmall; // Should always be non-NULL when mCustomIcon is non-NULL.
		}
		else
		{
			big_icon = (LPARAM)g_IconLarge;
			small_icon = (LPARAM)g_IconSmall;
		}

		SendMessage(hWndDlg, WM_SETICON, ICON_SMALL, small_icon);
		SendMessage(hWndDlg, WM_SETICON, ICON_BIG, big_icon);

		if(g_os.IsWinVistaOrLater())
		{
			// Use a more appealing font on Windows Vista and later (Segoe UI).
			HDC hdc = GetDC(hWndDlg);
			CURR_INPUTBOX.font = CreateFont(FONT_POINT(hdc, 10), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
					, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, _T("Segoe UI"));
			ReleaseDC(hWndDlg, hdc); // In theory it must be done.

			// Set the font.
			SendMessage(hControl, WM_SETFONT, (WPARAM)CURR_INPUTBOX.font, 0);
			SendMessage(GetDlgItem(hWndDlg, IDC_INPUTEDIT), WM_SETFONT, (WPARAM)CURR_INPUTBOX.font, 0);
			SendMessage(GetDlgItem(hWndDlg, IDOK), WM_SETFONT, (WPARAM)CURR_INPUTBOX.font, 0);
			SendMessage(GetDlgItem(hWndDlg, IDCANCEL), WM_SETFONT, (WPARAM)CURR_INPUTBOX.font, 0);
		}
		else
			CURR_INPUTBOX.font = NULL;

		// For the timeout, use a timer ID that doesn't conflict with MsgBox's IDs (which are the
		// integers 1 through the max allowed number of msgboxes).  Use +3 vs. +1 for a margin of safety
		// (e.g. in case a few extra MsgBoxes can be created directly by the program and not by
		// the script):
		#define INPUTBOX_TIMER_ID_OFFSET (MAX_MSGBOXES + 3)
		if (CURR_INPUTBOX.timeout)
			SetTimer(hWndDlg, INPUTBOX_TIMER_ID_OFFSET + target_index, CURR_INPUTBOX.timeout, InputBoxTimeout);

		return TRUE; // i.e. let the system set the keyboard focus to the first visible control.
	}

	case WM_DESTROY:
	{
		if (CURR_INPUTBOX.font)
			DeleteObject(CURR_INPUTBOX.font);

		return TRUE;
	}

	case WM_SIZE:
	{
		// Adapted from D.Nuttall's InputBox in the AutoIt3 source.

		// don't try moving controls if minimized
		if (wParam == SIZE_MINIMIZED)
			return TRUE;

		int dlg_new_width = LOWORD(lParam);
		int dlg_new_height = HIWORD(lParam);

		int last_ypos = 0, curr_width, curr_height;

		// Changing these might cause weird effects when user resizes the window since the default size and
		// margins is about 5 (as stored in the dialog's resource properties).  UPDATE: That's no longer
		// an issue since the dialog is resized when the dialog is first displayed to make sure everything
		// behaves consistently:
		const int XMargin = 5, YMargin = 5;

		RECT rTmp;

		// start at the bottom - OK button

		HWND hbtOk = GetDlgItem(hWndDlg, IDOK);
		if (hbtOk != NULL)
		{
			// how big is the control?
			GetWindowRect(hbtOk, &rTmp);
			if (rTmp.left > rTmp.right)
				swap(rTmp.left, rTmp.right);
			if (rTmp.top > rTmp.bottom)
				swap(rTmp.top, rTmp.bottom);
			curr_width = rTmp.right - rTmp.left;
			curr_height = rTmp.bottom - rTmp.top;
			last_ypos = dlg_new_height - YMargin - curr_height;
			// where to put the control?
			MoveWindow(hbtOk, dlg_new_width/4+(XMargin-curr_width)/2, last_ypos, curr_width, curr_height, FALSE);
		}

		// Cancel Button
		HWND hbtCancel = GetDlgItem(hWndDlg, IDCANCEL);
		if (hbtCancel != NULL)
		{
			// how big is the control?
			GetWindowRect(hbtCancel, &rTmp);
			if (rTmp.left > rTmp.right)
				swap(rTmp.left, rTmp.right);
			if (rTmp.top > rTmp.bottom)
				swap(rTmp.top, rTmp.bottom);
			curr_width = rTmp.right - rTmp.left;
			curr_height = rTmp.bottom - rTmp.top;
			// where to put the control?
			MoveWindow(hbtCancel, dlg_new_width*3/4-(XMargin+curr_width)/2, last_ypos, curr_width, curr_height, FALSE);
		}

		// Edit Box
		HWND hedText = GetDlgItem(hWndDlg, IDC_INPUTEDIT);
		if (hedText != NULL)
		{
			// how big is the control?
			GetWindowRect(hedText, &rTmp);
			if (rTmp.left > rTmp.right)
				swap(rTmp.left, rTmp.right);
			if (rTmp.top > rTmp.bottom)
				swap(rTmp.top, rTmp.bottom);
			curr_width = rTmp.right - rTmp.left;
			curr_height = rTmp.bottom - rTmp.top;
			last_ypos -= 5 + curr_height;  // Allows space between the buttons and the edit box.
			// where to put the control?
			MoveWindow(hedText, XMargin, last_ypos, dlg_new_width - XMargin*2
				, curr_height, FALSE);
		}

		// Static Box (Prompt)
		HWND hstPrompt = GetDlgItem(hWndDlg, IDC_INPUTPROMPT);
		if (hstPrompt != NULL)
		{
			last_ypos -= 10;  // Allows space between the edit box and the prompt (static text area).
			// where to put the control?
			MoveWindow(hstPrompt, XMargin, YMargin, dlg_new_width - XMargin*2
				, last_ypos, FALSE);
		}
		InvalidateRect(hWndDlg, NULL, TRUE);	// force window to be redrawn
		return TRUE;  // i.e. completely handled here.
	}

	case WM_COMMAND:
		// In this case, don't use (g_nInputBoxes - 1) as the index because it might
		// not correspond to the g_InputBox[] array element that belongs to hWndDlg.
		// This is because more than one input box can be on the screen at the same time.
		// If the user choses to work with on underneath instead of the most recent one,
		// we would be called with an hWndDlg whose index is less than the most recent
		// one's index (g_nInputBoxes - 1).  Instead, search the array for a match.
		// Work backward because the most recent one(s) are more likely to be a match:
		for (; target_index > -1; --target_index)
			if (g_InputBox[target_index].hwnd == hWndDlg)
				break;
		if (target_index < 0)  // Should never happen if things are designed right.
			return FALSE;
		switch (LOWORD(wParam))
		{
		case IDOK:
		case IDCANCEL:
		{
			WORD return_value = LOWORD(wParam);  // Set default, i.e. IDOK or IDCANCEL
			if (   !(hControl = GetDlgItem(hWndDlg, IDC_INPUTEDIT))   )
				return_value = (WORD)FAIL;
			else
			{
				// The output variable is set to whatever the user entered, even if the user pressed
				// the cancel button.  This allows the cancel button to specify that a different
				// operation should be performed on the entered text.
				// NOTE: ErrorLevel must not be set here because it's possible that the user has
				// dismissed a dialog that's underneath another, active dialog, or that's currently
				// suspended due to a timed/hotkey subroutine running on top of it.  In other words,
				// it's only safe to set ErrorLevel when the call to DialogProc() returns in InputBox().
				#undef INPUTBOX_VAR
				#define INPUTBOX_VAR (CURR_INPUTBOX.output_var)
				int space_needed = GetWindowTextLength(hControl) + 1;
				// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
				// this call will set up the clipboard for writing:
				if (INPUTBOX_VAR->AssignString(NULL, space_needed - 1) != OK)
					// It will have already displayed the error.  Displaying errors in a callback
					// function like this one isn't that good, since the callback won't return
					// to its caller in a timely fashion.  However, these type of errors are so
					// rare it's not a priority to change all the called functions (and the functions
					// they call) to skip the displaying of errors and just return FAIL instead.
					// In addition, this callback function has been tested with a MsgBox() call
					// inside and it doesn't seem to cause any crashes or undesirable behavior other
					// than the fact that the InputBox window is not dismissed until the MsgBox
					// window is dismissed:
					return_value = (WORD)FAIL;
				else
				{
					// Write to the variable:
					INPUTBOX_VAR->SetCharLength((VarSizeType)GetWindowText(hControl
						, INPUTBOX_VAR->Contents(), space_needed));
					if (!INPUTBOX_VAR->Length())
						// There was no text to get or GetWindowText() failed.
						*INPUTBOX_VAR->Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
					if (INPUTBOX_VAR->Close() != OK) // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_CLIPBOARD.
						return_value = (WORD)FAIL;
				}
			}
			// Since the user pressed a button to dismiss the dialog:
			// Kill its timer for performance reasons (might degrade perf. a little since OS has
			// to keep track of it as long as it exists).  InputBoxTimeout() already handles things
			// right even if we don't do this:
			if (CURR_INPUTBOX.timeout) // It has a timer.
				KillTimer(hWndDlg, INPUTBOX_TIMER_ID_OFFSET + target_index);
			EndDialog(hWndDlg, return_value);
			return TRUE;
		} // case
		} // Inner switch()
	} // Outer switch()
	// Otherwise, let the dialog handler do its default action:
	return FALSE;
}



VOID CALLBACK InputBoxTimeout(HWND hWnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	// First check if the window has already been destroyed.  There are quite a few ways this can
	// happen, and in all of them we want to make sure not to do things such as calling EndDialog()
	// again or updating the output variable.  Reasons:
	// 1) The user has already pressed the OK or Cancel button (the timer isn't killed there because
	//    it relies on us doing this check here).  In this case, EndDialog() has already been called
	//    (with the proper result value) and the script's output variable has already been set.
	// 2) Even if we were to kill the timer when the user presses a button to dismiss the dialog,
	//    this IsWindow() check would still be needed here because TimerProc()'s are called via
	//    WM_TIMER messages, some of which might still be in our msg queue even after the timer
	//    has been killed.  In other words, split second timing issues may cause this TimerProc()
	//    to fire even if the timer were killed when the user dismissed the dialog.
	// UPDATE: For performance reasons, the timer is now killed when the user presses a button,
	// so case #1 is obsolete (but kept here for background/insight).
	if (IsWindow(hWnd))
	{
		// This is the element in the array that corresponds to the InputBox for which
		// this function has just been called.
		INT_PTR target_index = idEvent - INPUTBOX_TIMER_ID_OFFSET;
		// Even though the dialog has timed out, we still want to write anything the user
		// had a chance to enter into the output var.  This is because it's conceivable that
		// someone might want a short timeout just to enter something quick and let the
		// timeout dismiss the dialog for them (i.e. so that they don't have to press enter
		// or a button:
		HWND hControl = GetDlgItem(hWnd, IDC_INPUTEDIT);
		if (hControl)
		{
			#undef INPUTBOX_VAR
			#define INPUTBOX_VAR (g_InputBox[target_index].output_var)
			int space_needed = GetWindowTextLength(hControl) + 1;
			// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
			// this call will set up the clipboard for writing:
			if (INPUTBOX_VAR->AssignString(NULL, space_needed - 1) == OK)
			{
				// Write to the variable:
				INPUTBOX_VAR->SetCharLength((VarSizeType)GetWindowText(hControl
					, INPUTBOX_VAR->Contents(), space_needed));
				if (!INPUTBOX_VAR->Length())
					// There was no text to get or GetWindowText() failed.
					*INPUTBOX_VAR->Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
				INPUTBOX_VAR->Close();
			}
		}
		EndDialog(hWnd, AHK_TIMEOUT);
	}
	KillTimer(hWnd, idEvent);
}



VOID CALLBACK DerefTimeout(HWND hWnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	Line::FreeDerefBufIfLarge(); // It will also kill the timer, if appropriate.
}



ResultType Line::MouseGetPos(DWORD aOptions)
// Returns OK or FAIL.
{
	// Caller should already have ensured that at least one of these will be non-NULL.
	// The only time this isn't true is for dynamically-built variable names.  In that
	// case, we don't worry about it if it's NULL, since the user will already have been
	// warned:
	Var *output_var_x = ARGVAR1;  // Ok if NULL. Load-time validation has ensured that these are valid output variables (e.g. not built-in vars).
	Var *output_var_y = ARGVAR2;  // Ok if NULL.
	Var *output_var_parent = ARGVAR3;  // Ok if NULL.
	Var *output_var_child = ARGVAR4;  // Ok if NULL.

	POINT point;
	GetCursorPos(&point);  // Realistically, can't fail?

	POINT origin = {0};
	CoordToScreen(origin, COORD_MODE_MOUSE);

	if (output_var_x) // else the user didn't want the X coordinate, just the Y.
		if (!output_var_x->Assign(point.x - origin.x))
			return FAIL;
	if (output_var_y) // else the user didn't want the Y coordinate, just the X.
		if (!output_var_y->Assign(point.y - origin.y))
			return FAIL;

	if (!output_var_parent && !output_var_child)
		return OK;

	// This is the child window.  Despite what MSDN says, WindowFromPoint() appears to fetch
	// a non-NULL value even when the mouse is hovering over a disabled control (at least on XP).
	HWND child_under_cursor = WindowFromPoint(point);
	if (!child_under_cursor)
	{
		if (output_var_parent)
			output_var_parent->Assign();
		if (output_var_child)
			output_var_child->Assign();
		return OK;
	}

	HWND parent_under_cursor = GetNonChildParent(child_under_cursor);  // Find the first ancestor that isn't a child.
	if (output_var_parent)
	{
		// Testing reveals that an invisible parent window never obscures another window beneath it as seen by
		// WindowFromPoint().  In other words, the below never happens, so there's no point in having it as a
		// documented feature:
		//if (!g->DetectHiddenWindows && !IsWindowVisible(parent_under_cursor))
		//	return output_var_parent->Assign();
		if (!output_var_parent->AssignHWND(parent_under_cursor))
			return FAIL;
	}

	if (!output_var_child)
		return OK;

	// Doing it this way overcomes the limitations of WindowFromPoint() and ChildWindowFromPoint()
	// and also better matches the control that Window Spy would think is under the cursor:
	if (!(aOptions & 0x01)) // Not in simple mode, so find the control the normal/complex way.
	{
		point_and_hwnd_type pah = {0};
		pah.pt = point;
		EnumChildWindows(parent_under_cursor, EnumChildFindPoint, (LPARAM)&pah); // Find topmost control containing point.
		if (pah.hwnd_found)
			child_under_cursor = pah.hwnd_found;
	}
	//else as of v1.0.25.10, leave child_under_cursor set the the value retrieved earlier from WindowFromPoint().
	// This allows MDI child windows to be reported correctly; i.e. that the window on top of the others
	// is reported rather than the one at the top of the z-order (the z-order of MDI child windows,
	// although probably constant, is not useful for determine which one is one top of the others).

	if (parent_under_cursor == child_under_cursor) // if there's no control per se, make it blank.
		return output_var_child->Assign();

	if (aOptions & 0x02) // v1.0.43.06: Bitwise flag that means "return control's HWND vs. ClassNN".
		return output_var_child->AssignHWND(child_under_cursor);

	class_and_hwnd_type cah;
	cah.hwnd = child_under_cursor;  // This is the specific control we need to find the sequence number of.
	TCHAR class_name[WINDOW_CLASS_SIZE];
	cah.class_name = class_name;
	if (!GetClassName(cah.hwnd, class_name, _countof(class_name) - 5))  // -5 to allow room for sequence number.
		return output_var_child->Assign();
	cah.class_count = 0;  // Init for the below.
	cah.is_found = false; // Same.
	EnumChildWindows(parent_under_cursor, EnumChildFindSeqNum, (LPARAM)&cah); // Find this control's seq. number.
	if (!cah.is_found)
		return output_var_child->Assign();  
	// Append the class sequence number onto the class name and set the output param to be that value:
	sntprintfcat(class_name, _countof(class_name), _T("%d"), cah.class_count);
	return output_var_child->Assign(class_name);
}



BOOL CALLBACK EnumChildFindPoint(HWND aWnd, LPARAM lParam)
// This is called by more than one caller.  It finds the most appropriate child window that contains
// the specified point (the point should be in screen coordinates).
{
	point_and_hwnd_type &pah = *(point_and_hwnd_type *)lParam;  // For performance and convenience.
	if (!IsWindowVisible(aWnd) // Omit hidden controls, like Window Spy does.
		|| (pah.ignore_disabled_controls && !IsWindowEnabled(aWnd))) // For ControlClick, also omit disabled controls, since testing shows that the OS doesn't post mouse messages to them.
		return TRUE;
	RECT rect;
	if (!GetWindowRect(aWnd, &rect))
		return TRUE;
	// The given point must be inside aWnd's bounds.  Then, if there is no hwnd found yet or if aWnd
	// is entirely contained within the previously found hwnd, update to a "better" found window like
	// Window Spy.  This overcomes the limitations of WindowFromPoint() and ChildWindowFromPoint():
	if (pah.pt.x >= rect.left && pah.pt.x <= rect.right && pah.pt.y >= rect.top && pah.pt.y <= rect.bottom)
	{
		// If the window's center is closer to the given point, break the tie and have it take
		// precedence.  This solves the problem where a particular control from a set of overlapping
		// controls is chosen arbitrarily (based on Z-order) rather than based on something the
		// user would find more intuitive (the control whose center is closest to the mouse):
		double center_x = rect.left + (double)(rect.right - rect.left) / 2;
		double center_y = rect.top + (double)(rect.bottom - rect.top) / 2;
		// Taking the absolute value first is not necessary because it seems that qmathHypot()
		// takes the square root of the sum of the squares, which handles negatives correctly:
		double distance = qmathHypot(pah.pt.x - center_x, pah.pt.y - center_y);
		//double distance = qmathSqrt(qmathPow(pah.pt.x - center_x, 2) + qmathPow(pah.pt.y - center_y, 2));
		bool update_it = !pah.hwnd_found;
		if (!update_it)
		{
			// If the new window's rect is entirely contained within the old found-window's rect, update
			// even if the distance is greater.  Conversely, if the new window's rect entirely encloses
			// the old window's rect, do not update even if the distance is less:
			if (rect.left >= pah.rect_found.left && rect.right <= pah.rect_found.right
				&& rect.top >= pah.rect_found.top && rect.bottom <= pah.rect_found.bottom)
				update_it = true; // New is entirely enclosed by old: update to the New.
			else if (   distance < pah.distance &&
				(pah.rect_found.left < rect.left || pah.rect_found.right > rect.right
					|| pah.rect_found.top < rect.top || pah.rect_found.bottom > rect.bottom)   )
				update_it = true; // New doesn't entirely enclose old and new's center is closer to the point.
		}
		if (update_it)
		{
			pah.hwnd_found = aWnd;
			pah.rect_found = rect; // And at least one caller uses this returned rect.
			pah.distance = distance;
		}
	}
	return TRUE; // Continue enumeration all the way through.
}



///////////////////////////////
// Related to other commands //
///////////////////////////////

ResultType Line::FormatTime(LPTSTR aYYYYMMDD, LPTSTR aFormat)
// The compressed code size of this function is about 1 KB (2 KB uncompressed), which compares
// favorably to using setlocale()+strftime(), which together are about 8 KB of compressed code
// (setlocale() seems to be needed to put the user's or system's locale into effect for strftime()).
// setlocale() weighs in at about 6.5 KB compressed (14 KB uncompressed).
{
	Var &output_var = *OUTPUT_VAR;

	#define FT_MAX_INPUT_CHARS 2000  // In preparation for future use of TCHARs, since GetDateFormat() uses char-count not size.
	// Input/format length is restricted since it must be translated and expanded into a new format
	// string that uses single quotes around non-alphanumeric characters such as punctuation:
	if (_tcslen(aFormat) > FT_MAX_INPUT_CHARS)
		return output_var.Assign();

	// Worst case expansion: .d.d.d.d. (9 chars) --> '.'d'.'d'.'d'.'d'.' (19 chars)
	// Buffer below is sized to a little more than twice as big as the largest allowed format,
	// which avoids having to constantly check for buffer overflow while translating aFormat
	// into format_buf:
	#define FT_MAX_OUTPUT_CHARS (2*FT_MAX_INPUT_CHARS + 10)
	TCHAR format_buf[FT_MAX_OUTPUT_CHARS + 1];
	TCHAR output_buf[FT_MAX_OUTPUT_CHARS + 1]; // The size of this is somewhat arbitrary, but buffer overflow is checked so it's safe.

	TCHAR yyyymmdd[256]; // Large enough to hold date/time and any options that follow it (note that D and T options can appear multiple times).
	*yyyymmdd = '\0';

	SYSTEMTIME st;
	LPTSTR options = NULL;

	if (!*aYYYYMMDD) // Use current local time by default.
		GetLocalTime(&st);
	else
	{
		tcslcpy(yyyymmdd, omit_leading_whitespace(aYYYYMMDD), _countof(yyyymmdd)); // Make a modifiable copy.
		if (*yyyymmdd < '0' || *yyyymmdd > '9') // First character isn't a digit, therefore...
		{
			// ... options are present without date (since yyyymmdd [if present] must come before options).
			options = yyyymmdd;
			GetLocalTime(&st);  // Use current local time by default.
		}
		else // Since the string starts with a digit, rules say it must be a YYYYMMDD string, possibly followed by options.
		{
			// Find first space or tab because YYYYMMDD portion might contain only the leading part of date/timestamp.
			if (options = StrChrAny(yyyymmdd, _T(" \t"))) // Find space or tab.
			{
				*options = '\0'; // Terminate yyyymmdd at the end of the YYYYMMDDHH24MISS string.
				options = omit_leading_whitespace(++options); // Point options to the right place (can be empty string).
			}
			//else leave options set to NULL to indicate that there are none.

			// Pass "false" for validation so that times can still be reported even if the year
			// is prior to 1601.  If the time and/or date is invalid, GetTimeFormat() and GetDateFormat()
			// will refuse to produce anything, which is documented behavior:
			YYYYMMDDToSystemTime(yyyymmdd, st, false);
		}
	}
	
	// Set defaults.  Some can be overridden by options (if there are any options).
	LCID lcid = LOCALE_USER_DEFAULT;
	DWORD date_flags = 0, time_flags = 0;
	bool date_flags_specified = false, time_flags_specified = false, reverse_date_time = false;
	#define FT_FORMAT_NONE 0
	#define FT_FORMAT_TIME 1
	#define FT_FORMAT_DATE 2
	int format_type1 = FT_FORMAT_NONE;
	LPTSTR format2_marker = NULL; // Will hold the location of the first char of the second format (if present).
	bool do_null_format2 = false;  // Will be changed to true if a default date *and* time should be used.

	if (options) // Parse options.
	{
		LPTSTR option_end;
		TCHAR orig_char;
		for (LPTSTR next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
		{
			// Find the end of this option item:
			if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
				option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.

			// Permanently terminate in between options to help eliminate ambiguity for words contained
			// inside other words, and increase confidence in decimal and hexadecimal conversion.
			orig_char = *option_end;
			*option_end = '\0';

			++next_option;
			switch (_totupper(next_option[-1]))
			{
			case 'D':
				date_flags_specified = true;
				date_flags |= ATOU(next_option); // ATOU() for unsigned.
				break;
			case 'T':
				time_flags_specified = true;
				time_flags |= ATOU(next_option); // ATOU() for unsigned.
				break;
			case 'R':
				reverse_date_time = true;
				break;
			case 'L':
				lcid = !_tcsicmp(next_option, _T("Sys")) ? LOCALE_SYSTEM_DEFAULT : (LCID)ATOU(next_option);
				break;
			// If not one of the above, such as zero terminator or a number, just ignore it.
			}

			*option_end = orig_char; // Undo the temporary termination so that loop's omit_leading() will work.
		} // for() each item in option list
	} // Parse options.

	if (!*aFormat)
	{
		aFormat = NULL; // Tell GetDateFormat() and GetTimeFormat() to use default for the specified locale.
		if (!date_flags_specified) // No preference was given, so use long (which seems generally more useful).
			date_flags |= DATE_LONGDATE;
		if (!time_flags_specified)
			time_flags |= TIME_NOSECONDS;  // Seems more desirable/typical to default to no seconds.
		// Put the time first by default, though this is debatable (Metapad does it and I like it).
		format_type1 = reverse_date_time ? FT_FORMAT_DATE : FT_FORMAT_TIME;
		do_null_format2 = true;
	}
	else // aFormat is non-blank.
	{
		// Omit whitespace only for consideration of special keywords.  Whitespace is later kept for
		// a normal format string such as %A_Space%MM/dd/yy:
		LPTSTR candidate = omit_leading_whitespace(aFormat);
		if (!_tcsicmp(candidate, _T("YWeek")))
		{
			GetISOWeekNumber(output_buf, st.wYear, GetYDay(st.wMonth, st.wDay, IS_LEAP_YEAR(st.wYear)), st.wDayOfWeek);
			return output_var.Assign(output_buf);
		}
		if (!_tcsicmp(candidate, _T("YDay")) || !_tcsicmp(candidate, _T("YDay0")))
		{
			int yday = GetYDay(st.wMonth, st.wDay, IS_LEAP_YEAR(st.wYear));
			if (!_tcsicmp(candidate, _T("YDay")))
				return output_var.Assign(yday); // Assign with no leading zeroes, also will be in hex format if that format is in effect.
			// Otherwise:
			_stprintf(output_buf, _T("%03d"), yday);
			return output_var.Assign(output_buf);
		}
		if (!_tcsicmp(candidate, _T("WDay")))
			return output_var.Assign(st.wDayOfWeek + 1);  // Convert to 1-based for compatibility with A_WDay.

		// Since above didn't return, check for those that require a call to GetTimeFormat/GetDateFormat
		// further below:
		if (!_tcsicmp(candidate, _T("ShortDate")))
		{
			aFormat = NULL;
			date_flags |= DATE_SHORTDATE;
			date_flags &= ~(DATE_LONGDATE | DATE_YEARMONTH); // If present, these would prevent it from working.
		}
		else if (!_tcsicmp(candidate, _T("LongDate")))
		{
			aFormat = NULL;
			date_flags |= DATE_LONGDATE;
			date_flags &= ~(DATE_SHORTDATE | DATE_YEARMONTH); // If present, these would prevent it from working.
		}
		else if (!_tcsicmp(candidate, _T("YearMonth")))
		{
			aFormat = NULL;
			date_flags |= DATE_YEARMONTH;
			date_flags &= ~(DATE_SHORTDATE | DATE_LONGDATE); // If present, these would prevent it from working.
		}
		else if (!_tcsicmp(candidate, _T("Time")))
		{
			format_type1 = FT_FORMAT_TIME;
			aFormat = NULL;
			if (!time_flags_specified)
				time_flags |= TIME_NOSECONDS;  // Seems more desirable/typical to default to no seconds.
		}
		else // Assume normal format string.
		{
			LPTSTR cp = aFormat, dp = format_buf;   // Initialize source and destination pointers.
			bool inside_their_quotes = false; // Whether we are inside a single-quoted string in the source.
			bool inside_our_quotes = false;   // Whether we are inside a single-quoted string of our own making in dest.
			for (; *cp; ++cp) // Transcribe aFormat into format_buf and also check for which comes first: date or time.
			{
				if (*cp == '\'') // Note that '''' (four consecutive quotes) is a single literal quote, which this logic handles okay.
				{
					if (inside_our_quotes)
					{
						// Upon encountering their quotes while we're still in ours, merge theirs with ours and
						// remark it as theirs.  This is done to avoid having two back-to-back quoted sections,
						// which would result in an unwanted literal single quote.  Example:
						// 'Some string'':' (the two quotes in the middle would be seen as a literal quote).
						inside_our_quotes = false;
						inside_their_quotes = true;
						continue;
					}
					if (inside_their_quotes)
					{
						// If next char needs to be quoted, don't close out this quote section because that
						// would introduce two consecutive quotes, which would be interpreted as a single
						// literal quote if its enclosed by two outer single quotes.  Instead convert this
						// quoted section over to "ours":
						if (cp[1] && !IsCharAlphaNumeric(cp[1]) && cp[1] != '\'') // Also consider single quotes to be theirs due to this example: dddd:''''y
							inside_our_quotes = true;
							// And don't do "*dp++ = *cp"
						else // there's no next-char or it's alpha-numeric, so it doesn't need to be inside quotes.
							*dp++ = *cp; // Close out their quoted section.
					}
					else // They're starting a new quoted section, so just transcribe this single quote as-is.
						*dp++ = *cp;
					inside_their_quotes = !inside_their_quotes; // Must be done after the above.
					continue;
				}
				// Otherwise, it's not a single quote.
				if (inside_their_quotes) // *cp is inside a single-quoted string, so it can be part of format/picture
					*dp++ = *cp; // Transcribe as-is.
				else
				{
					if (IsCharAlphaNumeric(*cp))
					{
						if (inside_our_quotes)
						{
							*dp++ = '\''; // Close out the previous quoted section, since this char should not be a part of it.
							inside_our_quotes = false;
						}
						if (_tcschr(_T("dMyg"), *cp)) // A format unique to Date is present.
						{
							if (!format_type1)
								format_type1 = FT_FORMAT_DATE;
							else if (format_type1 == FT_FORMAT_TIME && !format2_marker) // type2 should only be set if different than type1.
							{
								*dp++ = '\0';  // Terminate the first section and (below) indicate that there's a second.
								format2_marker = dp;  // Point it to the location in format_buf where the split should occur.
							}
						}
						else if (_tcschr(_T("hHmst"), *cp)) // A format unique to Time is present.
						{
							if (!format_type1)
								format_type1 = FT_FORMAT_TIME;
							else if (format_type1 == FT_FORMAT_DATE && !format2_marker) // type2 should only be set if different than type1.
							{
								*dp++ = '\0';  // Terminate the first section and (below) indicate that there's a second.
								format2_marker = dp;  // Point it to the location in format_buf where the split should occur.
							}
						}
						// For consistency, transcribe all AlphaNumeric chars not inside single quotes as-is
						// (numbers are transcribed in case they are ever used as part of pic/format).
						*dp++ = *cp;
					}
					else // Not alphanumeric, so enclose this and any other non-alphanumeric characters in single quotes.
					{
						if (!inside_our_quotes)
						{
							*dp++ = '\''; // Create a new quoted section of our own, since this char should be inside quotes to be understood.
							inside_our_quotes = true;
						}
						*dp++ = *cp;  // Add this character between the quotes, since it's of the right "type".
					}
				}
			} // for()
			if (inside_our_quotes)
				*dp++ = '\'';  // Close out our quotes.
			*dp = '\0'; // Final terminator.
			aFormat = format_buf; // Point it to the freshly translated format string, for use below.
		} // aFormat contains normal format/pic string.
	} // aFormat isn't blank.

	// If there are no date or time formats present, still do the transcription so that
	// any quoted strings and whatnot are resolved.  This increases runtime flexibility.
	// The below is also relied upon by "LongDate" and "ShortDate" above:
	if (!format_type1)
		format_type1 = FT_FORMAT_DATE;

	// MSDN: Time: "The function checks each of the time values to determine that it is within the
	// appropriate range of values. If any of the time values are outside the correct range, the
	// function fails, and sets the last-error to ERROR_INVALID_PARAMETER. 
	// Dates: "...year, month, day, and day of week. If the day of the week is incorrect, the
	// function uses the correct value, and returns no error. If any of the other date values
	// are outside the correct range, the function fails, and sets the last-error to ERROR_INVALID_PARAMETER.

	if (format_type1 == FT_FORMAT_DATE) // DATE comes first.
	{
		if (!GetDateFormat(lcid, date_flags, &st, aFormat, output_buf, FT_MAX_OUTPUT_CHARS))
			*output_buf = '\0';  // Ensure it's still the empty string, then try to continue to get the second half (if there is one).
	}
	else // TIME comes first.
		if (!GetTimeFormat(lcid, time_flags, &st, aFormat, output_buf, FT_MAX_OUTPUT_CHARS))
			*output_buf = '\0';  // Ensure it's still the empty string, then try to continue to get the second half (if there is one).

	if (format2_marker || do_null_format2) // There is also a second format present.
	{
		size_t output_buf_length = _tcslen(output_buf);
		LPTSTR output_buf_marker = output_buf + output_buf_length;
		LPTSTR format2;
		if (do_null_format2)
		{
			format2 = NULL;
			*output_buf_marker++ = ' '; // Provide a space between time and date.
			++output_buf_length;
		}
		else
			format2 = format2_marker;

		int buf_remaining_size = (int)(FT_MAX_OUTPUT_CHARS - output_buf_length);
		int result;

		if (format_type1 == FT_FORMAT_DATE) // DATE came first, so this one is TIME.
			result = GetTimeFormat(lcid, time_flags, &st, format2, output_buf_marker, buf_remaining_size);
		else
			result = GetDateFormat(lcid, date_flags, &st, format2, output_buf_marker, buf_remaining_size);
		if (!result)
			output_buf[output_buf_length] = '\0'; // Ensure the first part is still terminated and just return that rather than nothing.
	}

	return output_var.Assign(output_buf);
}



ResultType Line::PerformAssign()
// Returns OK or FAIL.  Caller has ensured that none of this line's derefs is a function-call.
{
	Var *p_output_var; // Can't use OUTPUT_VAR or sArgVar here because ExpandArgs() isn't called prior to PerformAssign().
	if (   !(p_output_var = ResolveVarOfArg(0))   ) // Fix for v1.0.46.07: Must do this check in case of illegal dynamically-build variable name.
		return FAIL;
	p_output_var = p_output_var->ResolveAlias(); // Resolve alias now to detect "source_is_being_appended_to_target" and perhaps other things.
	Var &output_var = *p_output_var; // For performance.
	// Now output_var.Type() must be clipboard or normal because otherwise load-time validation (or
	// ResolveVarOfArg() in GetExpandedArgSize, if it's dynamic) would have prevented us from getting this far.

	// Above must be checked prior to below since each uses "postfix" in a different way.
	if (mArgc > 1 && mArg[1].postfix) // There is a cached binary integer. is_expression is known to be false for ACT_ASSIGN, so no need to check it (since expression's use of postfix takes precedence over binary integer).
		return output_var.Assign(*(__int64 *)mArg[1].postfix);

	ArgStruct *arg2_with_at_least_one_deref;
	Var *arg_var[MAX_ARGS];

	if (mArgc < 2 || !mArg[1].deref || !mArg[1].deref[0].marker) // Relies on short-circuit boolean order. None of ACT_ASSIGN's args are ever ARG_TYPE_INPUT_VAR.
	{
		arg2_with_at_least_one_deref = NULL;
		arg_var[1] = NULL; // By contrast, arg_var[0] is the output_var, which historically is left uninitialized.
	}
	else // Arg #2 exists, has at least one deref, and (since function-calls aren't allowed in x=%y% statements) each such deref must be a variable.
	{
		arg2_with_at_least_one_deref = mArg + 1;
		// Can't use sArgVar here because ExecUntil() never calls ExpandArgs() for ACT_ASSIGN.
		// For simplicity, we don't check that it's the only deref, nor whether it has any literal text
		// around it, since those things aren't supported anyway.
		Var *source_var = arg2_with_at_least_one_deref->deref[0].var; // Caller has ensured none of this line's derefs is a function-call, so var should always be the proper member of the union to check.
		VarTypeType source_var_type = source_var->Type();
		if (source_var_type == VAR_CLIPBOARDALL) // The caller is performing the special mode "Var = %ClipboardAll%".
			return output_var.AssignClipboardAll(); // Outsourced to another function to help CPU cache hits/misses in this frequently-called function.
		if (source_var->IsBinaryClip()) // Caller wants a variable with binary contents assigned (copied) to another variable (usually VAR_CLIPBOARD).
			return output_var.AssignBinaryClip(*source_var); // Outsourced to another function to help CPU cache hits/misses in this frequently-called function.

		#define SINGLE_ISOLATED_DEREF (!arg2_with_at_least_one_deref->deref[1].marker\
			&& arg2_with_at_least_one_deref->deref[0].length == arg2_with_at_least_one_deref->length) // and the arg contains no literal text
		if (SINGLE_ISOLATED_DEREF) // The macro is used for maintainability because there are other places that use the same name for a macro of similar purposes.
		{
			if (   source_var_type == VAR_NORMAL // Not necessary to check output_var.Type()==VAR_NORMAL because VAR_CLIPBOARD is handled properly when AutoTrim is off (clipboard can never have HasUnflushedBinaryNumber()==true).
				&& (!g->AutoTrim || source_var->HasUnflushedBinaryNumber())   ) // When AutoTrim is off, or it's on but the source variable's mContents is out-of-date, output_var.Assign() is capable of handling the copy, and does so much faster.
				return output_var.Assign(*source_var); // In this case, it's okay if target_is_involved_in_source below would be true because this can handle copying a variable to itself.
				// Since modern scripts don't use Var=%Var2% very often and since a lot of complicated code
				// would be needed to change Assign(Var...) to accept a parameter such as aObeyAutoTrim, it
				// doesn't seem worth doing (all it would save is the copying one variable's mContents to another
				// when the original var is a number, in which case its mContents tends to be quite short anyway).
			//else continue on to later handling.
			arg_var[1] = source_var; // Set for use later on.
		}
		else
			arg_var[1] = NULL; // By contrast, arg_var[0] is the output_var, which historically is left uninitialized.
	}

	// Otherwise (since above didn't return):
	// Find out if output_var (the var being assigned to) is dereferenced (mentioned) in this line's
	// second arg, which is the value to be assigned.  If it isn't, things are much simpler.
	// Note: Since Arg#2 for this function is never an output or an input variable, it is not
	// necessary to check whether its the same variable as Arg#1 for this determination.
	bool target_is_involved_in_source = false;
	bool source_is_being_appended_to_target = false; // v1.0.25
	if (arg2_with_at_least_one_deref // There's at least one deref in arg #2, and...
		&& output_var.Type() != VAR_CLIPBOARD) // ...output_var isn't the clipboard. Checked because:
		// If type is VAR_CLIPBOARD, the checks below can be skipped because the clipboard can be used
		// in the source deref(s) while also being the target -- without having to use the deref buffer
		// -- because the clipboard has it's own temp buffer: the memory area to which the result is
		// written. The prior content of the clipboard remains available in its other memory area until
		// Commit() is called (i.e. long enough for this purpose).  For this reason,
		// source_is_being_appended_to_target also doesn't need to be determined for the clipboard.
	{
		// It has a second arg, which in this case is the value to be assigned to the var.
		// Examine any derefs that the second arg has to see if output_var is mentioned.
		// Also, calls to script functions aren't possible within these derefs because
		// our caller has ensured there are no expressions, and thus no function calls,
		// inside this line.
		for (DerefType *deref = arg2_with_at_least_one_deref->deref; deref->marker; ++deref)
		{
			if (source_is_being_appended_to_target)
			{
				// Check if target is mentioned more than once in source, e.g. Var = %Var%Some Text%Var%
				// would be disqualified for the "fast append" method because %Var% occurs more than once.
				if (deref->var->ResolveAlias() == p_output_var) // deref->is_function was checked above just in case.
				{
					source_is_being_appended_to_target = false;
					break;
				}
			}
			else
			{
				if (deref->var->ResolveAlias() == p_output_var) // deref->is_function was checked above just in case.
				{
					target_is_involved_in_source = true;
					// The below disqualifies both of the following cases from the simple-append mode:
					// Var = %OtherVar%%Var%   ; Var is not the first item as required.
					// Var = LiteralText%Var%  ; Same.
					if (deref->marker == arg2_with_at_least_one_deref->text)
						source_is_being_appended_to_target = true;
						// And continue the loop to ensure that Var is not referenced more than once,
						// e.g. Var = %Var%%Var% would be disqualified.
					else
						break;
				}
			}
		}
	}

	// Note: It might be possible to improve performance in the case where
	// the target variable is large enough to accommodate the new source data
	// by moving memory around inside it.  For example, Var1 = xxxxxVar1
	// could be handled by moving the memory in Var1 to make room to insert
	// the literal string.  In addition to being quicker than the ExpandArgs()
	// method, this approach would avoid the possibility of needing to expand the
	// deref buffer just to handle the operation.  However, if that is ever done,
	// be sure to check that output_var is mentioned only once in the list of derefs.
	// For example, something like this would probably be much easier to
	// implement by using ExpandArgs(): Var1 = xxxx %Var1% %Var2% %Var1% xxxx.
	// So the main thing to be possibly later improved here is the case where
	// output_var is mentioned only once in the deref list (which as of v1.0.25,
	// has been partially done via the concatenation improvement, e.g. Var = %Var%Text).
	VarSizeType space_needed;
	if (target_is_involved_in_source && !source_is_being_appended_to_target) // If true, output_var isn't the clipboard due to invariant: target_is_involved_in_source==false whenever output_var.Type()==VAR_CLIPBOARD.
	{
		if (ExpandArgs() != OK)
			return FAIL;
		// ARG2 now contains the dereferenced (literal) contents of the text we want to assign.
		// Therefore, calling ArgLength() is safe now too (i.e. ExpandArgs set things up for it).
		space_needed = (VarSizeType)ArgLength(2) + 1;  // +1 for the zero terminator.
	}
	else
	{
		// The following section is a simplified version of GetExpandedArgSize(), so maintain them together:
		if (mArgc < 2) // It's an assignment with nothing on the right side like "Var=".
			space_needed = 1;
		else if (arg_var[1]) // Arg #2 is a single isolated variable, discovered at an earlier stage.
			space_needed = arg_var[1]->Get() + 1;  // +1 for the zero terminator.
		else // This arg has more than one deref, or a single deref with some literal text around it.
		{
			space_needed = mArg[1].length + 1; // +1 for this arg's zero terminator in the buffer.
			if (arg2_with_at_least_one_deref)
			{
				for (DerefType *deref = arg2_with_at_least_one_deref->deref; deref->marker; ++deref)
				{
					// Replace the length of the deref's literal text with the length of its variable's contents.
					// All deref items for non-expressions like ACT_ASSIGN have been verified at loadtime to be
					// variables, not function-calls, so no need to check which type each deref is.
					space_needed -= deref->length;
					space_needed += deref->var->Get(); // If an environment var, Get() will yield its length.
				}
			}
		}
	}

	// Now above has ensured that space_needed is at least 1 (it should not be zero because even
	// the empty string uses up 1 char for its zero terminator).  The below relies upon this fact.

	if (space_needed < 2) // Variable is being assigned the empty string (or a deref that resolves to it).
		return output_var.Assign(_T(""));  // If the var is of large capacity, this will also free its memory.

	if (source_is_being_appended_to_target)
	{
		if (space_needed > output_var.Capacity())
		{
			// Since expanding the size of output_var while preserving its existing contents would
			// likely be a slow operation, revert to the normal method rather than the fast-append
			// mode.  Expand the args then continue on normally to the below.
			if (ExpandArgs(NULL, space_needed, arg_var) != OK) // In this case, both params were previously calculated by GetExpandedArgSize().
				return FAIL;
		}
		else // there's enough capacity in output_var to accept the text to be appended.
			target_is_involved_in_source = false;  // Tell the below not to consider expanding the args.
	}

	LPTSTR contents;
	if (target_is_involved_in_source) // output_var can't be clipboard due to invariant: target_is_involved_in_source==false whenever output_var.Type()==VAR_CLIPBOARD.
	{
		// It was already dereferenced above, so use ARG2, which points to the deref'ed contents of ARG2
		// (i.e. the data to be assigned).  I don't think there's any point to checking ARGVAR2!=NULL
		// and if so passing ARGVAR2->LengthIgnoreBinaryClip, because when we're here, ExpandArgs() will
		// have seen that it can't optimize it that way and thus it has fully expanded the variable into the buffer.
		if (!output_var.Assign(ARG2)) // Don't pass it space_needed-1 as the length because space_needed might be a conservative estimate larger than the actual length+terminator.
			return FAIL;
		if (g->AutoTrim)
		{
			contents = output_var.Contents();
			if (*contents)
			{
				output_var.SetCharLength((VarSizeType)trim(contents, output_var.Length())); // Passing length to trim() is known to greatly improve performance for long strings.
				output_var.Close(); // For maintainability (probably not currently necessary due to Assign() being called above).
			}
		}
		return OK;
	}

	// Otherwise: target isn't involved in source (output_var isn't itself involved in the source data/parameter).
	// First set everything up for the operation.  If output_var is the clipboard, this
	// will prepare the clipboard for writing.  Update: If source is being appended
	// to target using the simple method, we know output_var isn't the clipboard because the
	// logic at the top of this function ensures that.
	if (!source_is_being_appended_to_target)
		if (output_var.AssignString(NULL, space_needed - 1) != OK)
			return FAIL;
	// Expand Arg2 directly into the var.  Note: If output_var is the clipboard,
	// it's probably okay if the below actually writes less than the size of
	// the mem that has already been allocated for the new clipboard contents
	// That might happen due to a failure or size discrepancy between the
	// deref size-estimate and the actual deref itself:
	contents = output_var.Contents();
	// This knows not to copy the first var-ref onto itself (for when source_is_being_appended_to_target is true).
	// In addition, to reach this point, arg_var[0]'s value will already have been determined (possibly NULL)
	// by GetExpandedArgSize():
	LPTSTR one_beyond_contents_end = ExpandArg(contents, 1, arg_var[1]); // v1.0.45: Fixed arg_var[0] to be arg_var[1] (but this was only a performance issue).
	if (!one_beyond_contents_end)
		return FAIL;  // ExpandArg() will have already displayed the error.
	// Set the length explicitly rather than using space_needed because GetExpandedArgSize()
	// sometimes returns a larger size than is actually needed (e.g. for ScriptGetCursor()):
	size_t length = one_beyond_contents_end - contents - 1;
	// v1.0.25: Passing the precalculated length to trim() greatly improves performance,
	// especially for concat loops involving things like Var = %Var%String:
	output_var.SetCharLength((VarSizeType)(g->AutoTrim ? trim(contents, length) : length));
	return output_var.Close(); // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_CLIPBOARD.
}



ResultType Line::StringReplace()
// v1.0.45: Revised to improve average-case performance and reduce memory utilization.
{
	Var &output_var = *OUTPUT_VAR;
	LPTSTR source = ARG2;
	size_t length = ArgLength(2); // Going in, it's the haystack length. Later (coming out), it's the result length.

	bool alternate_errorlevel = tcscasestr(ARG5, _T("UseErrorLevel")); // This also implies replace-all.
	UINT replacement_limit = (alternate_errorlevel || StrChrAny(ARG5, _T("1aA"))) // This must be done in a way that recognizes "AllSlow" as meaning replace-all (even though the slow method itself is obsolete).
		? UINT_MAX : 1;

	// In case the strings involved are massive, free the output_var in advance of the operation to
	// reduce memory load and avoid swapping (but only if output_var isn't the same address as the input_var).
	if (output_var.Type() == VAR_NORMAL && source != output_var.Contents(FALSE)) // It's compared this way in case ByRef/aliases are involved.  This will detect even them.
		output_var.Free();
	//else source and dest are the same, so can't free the dest until after the operation.

	// Note: The current implementation of StrReplace() should be able to handle any conceivable inputs
	// without an empty string causing an infinite loop and without going infinite due to finding the
	// search string inside of newly-inserted replace strings (e.g. replacing all occurrences
	// of b with bcb would not keep finding b in the newly inserted bcd, infinitely).
	LPTSTR dest;
	UINT found_count = StrReplace(source, ARG3, ARG4, (StringCaseSenseType)g->StringCaseSense
		, replacement_limit, -1, &dest, &length); // Length of haystack is passed to improve performance because ArgLength() can often discover it instantaneously.

	if (!dest) // Failure due to out of memory.
		return LineError(ERR_OUTOFMEM);

	if (dest != source) // StrReplace() allocated new memory rather than returning "source" to us unaltered.
	{
		// v1.0.45: Take a shortcut for performance: Hang the newly allocated memory (populated by the callee)
		// directly onto the variable, which saves a memcpy() over the old method (and possible other savings).
		// AcceptNewMem() will shrink the memory for us, via _expand(), if there's a lot of extra/unused space in it.
		output_var.AcceptNewMem(dest, (VarSizeType)length); // Tells the variable to adopt this memory as its new memory. Callee has set "length" for us.
		// Above also handles the case where output_var is VAR_CLIPBOARD.
	}
	else // StrReplace gave us back "source" unaltered because no replacements were needed.
	{
		if (output_var.Type() == VAR_NORMAL)
		{
			// Technically the following check isn't necessary because Assign() also checks for it.
			// But since StringReplace is a frequently-used command, checking it here seems worthwhile
			// to avoid calling Assign().
			if (source != output_var.Contents(FALSE)) // It's compared this way in case ByRef/aliases are involved.  This will detect even them.
				output_var.Assign(source, (VarSizeType)length); // Callee has set "length" for us.
			//else the unaltered result and output_var same the same address.  Nothing needs to be done (for
			// simplicity, not even the binary-clipboard attribute is removed if it happens to be present).
		}
		else // output_var is of type VAR_CLIPBOARD.
			if (ARGVARRAW2->Type() != VAR_CLIPBOARD) // Arg index #1 (the second arg) is a normal var or some read-only var.
				output_var.Assign(source, (VarSizeType)length); // Callee has set "length" for us.
			//else the unaltered result and output_var are both the clipboard.  Nothing needs to be done.
	}

	if (alternate_errorlevel)
		g_ErrorLevel->Assign((DWORD)found_count);
	else // Use old ErrorLevel method for backward compatibility.
		g_ErrorLevel->Assign(found_count ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
	return OK;
}



ResultType Line::StringSplit(LPTSTR aArrayName, LPTSTR aInputString, LPTSTR aDelimiterList, LPTSTR aOmitList)
{
	// Make it longer than Max so that FindOrAddVar() will be able to spot and report var names
	// that are too long, either because the base-name is too long, or the name becomes too long
	// as a result of appending the array index number:
	TCHAR var_name[MAX_VAR_NAME_LENGTH + 21]; // Allow room for largest 64-bit integer, 20 chars: 18446744073709551616.
	tcslcpy(var_name, aArrayName, MAX_VAR_NAME_LENGTH+1); // This prefix is copied into it only once, for performance.
	LPTSTR var_name_suffix = var_name + _tcslen(var_name);

	Var *array0;
	if (mAttribute != ATTR_NONE) // 1.0.46.10: Fixed to rely on loadtime's determination of whether ArrayName0 is truly local or global (only loadtime currently has any awareness of declarations, so the determination must be made there unless "ArrayName" itself is a dynamic variable, which seems too rare to worry about).
		array0 = (Var *)mAttribute;
	else
	{
		var_name_suffix[0] = '0';
		var_name_suffix[1] = '\0';
		if (   !(array0 = g_script.FindOrAddVar(var_name))   )
			return FAIL;  // It will have already displayed the error.
	}
	int always_use = FINDVAR_FOR_PSEUDO_ARRAY(*array0);

	if (!*aInputString) // The input variable is blank, thus there will be zero elements.
		return array0->Assign(_T("0"));  // Store the count in the 0th element.

	DWORD next_element_number;
	Var *next_element;

	if (*aDelimiterList) // The user provided a list of delimiters, so process the input variable normally.
	{
		LPTSTR contents_of_next_element, delimiter, new_starting_pos;
		size_t element_length;
		for (contents_of_next_element = aInputString, next_element_number = 1; ; ++next_element_number)
		{
			_ultot(next_element_number, var_name_suffix, 10);
			// To help performance (in case the linked list of variables is huge), tell it where
			// to start the search.  Use element #0 rather than the preceding element because,
			// for example, Array19 is alphabetically less than Array2, so we can't rely on the
			// numerical ordering:
			if (   !(next_element = g_script.FindOrAddVar(var_name, 0, always_use))   )
				return FAIL;  // It will have already displayed the error.

			if (delimiter = StrChrAny(contents_of_next_element, aDelimiterList)) // A delimiter was found.
			{
				element_length = delimiter - contents_of_next_element;
				if (*aOmitList && element_length > 0)
				{
					contents_of_next_element = omit_leading_any(contents_of_next_element, aOmitList, element_length);
					element_length = delimiter - contents_of_next_element; // Update in case above changed it.
					if (element_length)
						element_length = omit_trailing_any(contents_of_next_element, aOmitList, delimiter - 1);
				}
				// If there are no chars to the left of the delim, or if they were all in the list of omitted
				// chars, the variable will be assigned the empty string:
				if (!next_element->Assign(contents_of_next_element, (VarSizeType)element_length))
					return FAIL;
				contents_of_next_element = delimiter + 1;  // Omit the delimiter since it's never included in contents.
			}
			else // the entire length of contents_of_next_element is what will be stored
			{
				element_length = _tcslen(contents_of_next_element);
				if (*aOmitList && element_length > 0)
				{
					new_starting_pos = omit_leading_any(contents_of_next_element, aOmitList, element_length);
					element_length -= (new_starting_pos - contents_of_next_element); // Update in case above changed it.
					contents_of_next_element = new_starting_pos;
					if (element_length)
						// If this is true, the string must contain at least one char that isn't in the list
						// of omitted chars, otherwise omit_leading_any() would have already omitted them:
						element_length = omit_trailing_any(contents_of_next_element, aOmitList
							, contents_of_next_element + element_length - 1);
				}
				// If there are no chars to the left of the delim, or if they were all in the list of omitted
				// chars, the variable will be assigned the empty string:
				if (!next_element->Assign(contents_of_next_element, (VarSizeType)element_length))
					return FAIL;
				// This is the only way out of the loop other than critical errors:
				return array0->Assign(next_element_number); // Store the count of how many items were stored in the array.
			}
		}
	}

	// Otherwise aDelimiterList is empty, so store each char of aInputString in its own array element.
	LPTSTR cp, dp;
	for (cp = aInputString, next_element_number = 1; *cp; ++cp)
	{
		for (dp = aOmitList; *dp; ++dp)
			if (*cp == *dp) // This char is a member of the omitted list, thus it is not included in the output array.
				break;
		if (*dp) // Omitted.
			continue;
		_ultot(next_element_number, var_name_suffix, 10);
		if (   !(next_element = g_script.FindOrAddVar(var_name, 0, always_use))   )
			return FAIL;  // It will have already displayed the error.
		if (!next_element->Assign(cp, 1))
			return FAIL;
		++next_element_number; // Only increment this if above didn't "continue".
	}
	return array0->Assign(next_element_number - 1); // Store the count of how many items were stored in the array.
}



BIF_DECL(BIF_StrSplit)
// Array := StrSplit(String [, Delimiters, OmitChars, MaxParts])
// This is the v2 version of Line::StringSplit(), and as such, is kept separate from Line::StringSplit().
// Unlike StringSplit, this function allows an array of Delimiters (vs a string of delimiter characters).
{
	LPTSTR aInputString = TokenToString(*aParam[0], aResultToken.buf);
	LPTSTR *aDelimiterList = NULL;
	int aDelimiterCount = 0;
	LPTSTR aOmitList = _T("");
	int splits_left = -2;

	if (aParamCount > 1)
	{
		if (Object *obj = dynamic_cast<Object *>(TokenToObject(*aParam[1])))
		{
			aDelimiterCount = obj->GetNumericItemCount();
			aDelimiterList = (LPTSTR *)_alloca(aDelimiterCount * sizeof(LPTSTR *));
			if (!obj->ArrayToStrings(aDelimiterList, aDelimiterCount, aDelimiterCount))
				// Array contains something other than a string.
				goto return_empty_string;
			for (int i = 0; i < aDelimiterCount; ++i)
				if (!*aDelimiterList[i])
					// Empty string in delimiter list. Although it could be treated similarly to the
					// "no delimiter" case, it's far more likely to be an error. If ever this check
					// is removed, the loop below must be changed to support "" as a delimiter.
					goto return_empty_string;
		}
		else
		{
			aDelimiterList = (LPTSTR *)_alloca(sizeof(LPTSTR *));
			*aDelimiterList = TokenToString(*aParam[1]);
			aDelimiterCount = **aDelimiterList != '\0'; // i.e. non-empty string.
		}
		if (aParamCount > 2)
		{
			aOmitList = TokenToString(*aParam[2]);
			if (aParamCount > 3)
				splits_left = (int)TokenToInt64(*aParam[3]) - 1;
		}
	}
	
	Object *output_array = Object::Create();
	if (!output_array)
		goto return_empty_string;
	aResultToken.symbol = SYM_OBJECT;	// Set default, overridden only for critical errors.
	aResultToken.object = output_array;	//

	if (!*aInputString // The input variable is blank, thus there will be zero elements.
		|| splits_left == -1) // The caller specified 0 parts.
		return;
	
	LPTSTR contents_of_next_element, delimiter, new_starting_pos;
	size_t element_length, delimiter_length;

	if (aDelimiterCount) // The user provided a list of delimiters, so process the input variable normally.
	{
		for (contents_of_next_element = aInputString; ; )
		{
			if (   !splits_left // Limit reached.
				|| !(delimiter = InStrAny(contents_of_next_element, aDelimiterList, aDelimiterCount, delimiter_length))   ) // No delimiter found.
				break; // This is the only way out of the loop other than critical errors.
			element_length = delimiter - contents_of_next_element;
			if (*aOmitList && element_length > 0)
			{
				contents_of_next_element = omit_leading_any(contents_of_next_element, aOmitList, element_length);
				element_length = delimiter - contents_of_next_element; // Update in case above changed it.
				if (element_length)
					element_length = omit_trailing_any(contents_of_next_element, aOmitList, delimiter - 1);
			}
			// If there are no chars to the left of the delim, or if they were all in the list of omitted
			// chars, the variable will be assigned the empty string:
			if (!output_array->Append(contents_of_next_element, element_length))
				goto outofmem;
			contents_of_next_element = delimiter + delimiter_length;  // Omit the delimiter since it's never included in contents.
			if (splits_left > 0)
				--splits_left;
		}
	}
	else
	{
		// Otherwise aDelimiterList is empty, so store each char of aInputString in its own array element.
		LPTSTR cp, dp;
		for (cp = aInputString; ; ++cp)
		{
			if (!*cp)
				return; // All done; result already set.
			for (dp = aOmitList; *dp; ++dp)
				if (*cp == *dp) // This char is a member of the omitted list, thus it is not included in the output array.
					break; // (inner loop)
			if (*dp) // Omitted.
				continue;
			if (!splits_left) // Limit reached (checked only after excluding omitted chars).
				break; // This is the only way out of the loop other than critical errors.
			if (splits_left > 0)
				--splits_left;
			if (!output_array->Append(cp, 1))
				goto outofmem;
		}
		contents_of_next_element = cp;
	}
	// Since above used break rather than goto or return, either the limit was reached or there are
	// no more delimiters, so store the remainder of the string minus any characters to be omitted.
	element_length = _tcslen(contents_of_next_element);
	if (*aOmitList && element_length > 0)
	{
		new_starting_pos = omit_leading_any(contents_of_next_element, aOmitList, element_length);
		element_length -= (new_starting_pos - contents_of_next_element); // Update in case above changed it.
		contents_of_next_element = new_starting_pos;
		if (element_length)
			// If this is true, the string must contain at least one char that isn't in the list
			// of omitted chars, otherwise omit_leading_any() would have already omitted them:
			element_length = omit_trailing_any(contents_of_next_element, aOmitList
				, contents_of_next_element + element_length - 1);
	}
	// If there are no chars to the left of the delim, or if they were all in the list of omitted
	// chars, the variable will be assigned the empty string:
	if (output_array->Append(contents_of_next_element, element_length))
		return; // All done; result already set.
	//else memory allocation failed, so fall through:
outofmem:
	// The fact that this section is executing means that a memory allocation failed and caused the
	// loop to break, so return a false value to let the caller detect the failure.  Empty string
	// is used vs 0 for consistency with Object() and Array().
	output_array->Release(); // Since we're not returning it.
return_empty_string:
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");
}



BIF_DECL(BIF_StrReplace)
{
	TCHAR old_buf[MAX_NUMBER_SIZE], new_buf[MAX_NUMBER_SIZE];
	// Must use aResultToken.buf for source in case StrReplace() performs no replacements:
	LPTSTR source = ParamIndexToString(0, aResultToken.buf);	// Parameter #1: Haystack
	size_t length = ParamIndexLength(0, source);
	LPTSTR oldstr = ParamIndexToString(1, old_buf);				// Parameter #2: SearchText
	LPTSTR newstr = ParamIndexToOptionalString(2, new_buf);		// Parameter #3: ReplaceText
	Var *output_var_count = ParamIndexToOptionalVar(3);			// Parameter #4: OutputVarCount
	UINT replacement_limit = (UINT)ParamIndexToOptionalInt64(4, UINT_MAX); // Parameter #5: Limit

	LPTSTR dest;
	UINT found_count = StrReplace(source, oldstr, newstr, (StringCaseSenseType)g->StringCaseSense
		, replacement_limit, -1, &dest, &length);

	if (!dest) // Failure due to out of memory.
	{
		aResult = g_script.ScriptError(ERR_OUTOFMEM);
		return;
	}

	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = dest;

	if (dest != source) // StrReplace() allocated new memory rather than returning "source" to us unaltered.
	{
		aResultToken.mem_to_free = dest; // Let caller know it needs to be freed.
		aResultToken.marker_length = length; // Must always be set if using mem_to_free.
	}

	if (output_var_count)
		output_var_count->Assign((DWORD)found_count);
}



ResultType Line::SplitPath(LPTSTR aFileSpec)
{
	Var *output_var_name = ARGVAR2;  // i.e. Param #2. Ok if NULL.
	Var *output_var_dir = ARGVAR3;  // Ok if NULL. Load-time validation has ensured that these are valid output variables (e.g. not built-in vars).
	Var *output_var_ext = ARGVAR4;  // Ok if NULL.
	Var *output_var_name_no_ext = ARGVAR5;  // Ok if NULL.
	Var *output_var_drive = ARGVAR6;  // Ok if NULL.

	// For URLs, "drive" is defined as the server name, e.g. http://somedomain.com
	LPTSTR name = _T(""), name_delimiter = NULL, drive_end = NULL; // Set defaults to improve maintainability.
	LPTSTR drive = omit_leading_whitespace(aFileSpec); // i.e. whitespace is considered for everything except the drive letter or server name, so that a pathless filename can have leading whitespace.
	LPTSTR colon_double_slash = _tcsstr(aFileSpec, _T("://"));

	if (colon_double_slash) // This is a URL such as ftp://... or http://...
	{
		if (   !(drive_end = _tcschr(colon_double_slash + 3, '/'))   )
		{
			if (   !(drive_end = _tcschr(colon_double_slash + 3, '\\'))   ) // Try backslash so that things like file://C:\Folder\File.txt are supported.
				drive_end = colon_double_slash + _tcslen(colon_double_slash); // Set it to the position of the zero terminator instead.
				// And because there is no filename, leave name and name_delimiter set to their defaults.
			//else there is a backslash, e.g. file://C:\Folder\File.txt, so treat that backslash as the end of the drive name.
		}
		name_delimiter = drive_end; // Set default, to be possibly overridden below.
		// Above has set drive_end to one of the following:
		// 1) The slash that occurs to the right of the doubleslash in a URL.
		// 2) The backslash that occurs to the right of the doubleslash in a URL.
		// 3) The zero terminator if there is no slash or backslash to the right of the doubleslash.
		if (*drive_end) // A slash or backslash exists to the right of the server name.
		{
			if (*(drive_end + 1))
			{
				// Find the rightmost slash.  At this stage, this is known to find the correct slash.
				// In the case of a file at the root of a domain such as http://domain.com/root_file.htm,
				// the directory consists of only the domain name, e.g. http://domain.com.  This is because
				// the directory always the "drive letter" by design, since that is more often what the
				// caller wants.  A script can use StringReplace to remove the drive/server portion from
				// the directory, if desired.
				name_delimiter = _tcsrchr(aFileSpec, '/');
				if (name_delimiter == colon_double_slash + 2) // To reach this point, it must have a backslash, something like file://c:\folder\file.txt
					name_delimiter = _tcsrchr(aFileSpec, '\\'); // Will always be found.
				name = name_delimiter + 1; // This will be the empty string for something like http://domain.com/dir/
			}
			//else something like http://domain.com/, so leave name and name_delimiter set to their defaults.
		}
		//else something like http://domain.com, so leave name and name_delimiter set to their defaults.
	}
	else // It's not a URL, just a file specification such as c:\my folder\my file.txt, or \\server01\folder\file.txt
	{
		// Differences between _splitpath() and the method used here:
		// _splitpath() doesn't include drive in output_var_dir, it includes a trailing
		// backslash, it includes the . in the extension, it considers ":" to be a filename.
		// _splitpath(pathname, drive, dir, file, ext);
		//char sdrive[16], sdir[MAX_PATH], sname[MAX_PATH], sext[MAX_PATH];
		//_splitpath(aFileSpec, sdrive, sdir, sname, sext);
		//if (output_var_name_no_ext)
		//	output_var_name_no_ext->Assign(sname);
		//strcat(sname, sext);
		//if (output_var_name)
		//	output_var_name->Assign(sname);
		//if (output_var_dir)
		//	output_var_dir->Assign(sdir);
		//if (output_var_ext)
		//	output_var_ext->Assign(sext);
		//if (output_var_drive)
		//	output_var_drive->Assign(sdrive);
		//return OK;

		// Don't use _splitpath() since it supposedly doesn't handle UNC paths correctly,
		// and anyway we need more info than it provides.  Also note that it is possible
		// for a file to begin with space(s) or a dot (if created programmatically), so
		// don't trim or omit leading space unless it's known to be an absolute path.

		// Note that "C:Some File.txt" is a valid filename in some contexts, which the below
		// tries to take into account.  However, there will be no way for this command to
		// return a path that differentiates between "C:Some File.txt" and "C:\Some File.txt"
		// since the first backslash is not included with the returned path, even if it's
		// the root directory (i.e. "C:" is returned in both cases).  The "C:Filename"
		// convention is pretty rare, and anyway this trait can be detected via something like
		// IfInString, Filespec, :, IfNotInString, Filespec, :\, MsgBox Drive with no absolute path.

		// UNCs are detected with this approach so that double sets of backslashes -- which sometimes
		// occur by accident in "built filespecs" and are tolerated by the OS -- are not falsely
		// detected as UNCs.
		if (drive[0] == '\\' && drive[1] == '\\') // Relies on short-circuit evaluation order.
		{
			if (   !(drive_end = _tcschr(drive + 2, '\\'))   )
				drive_end = drive + _tcslen(drive); // Set it to the position of the zero terminator instead.
		}
		else if (*(drive + 1) == ':') // It's an absolute path.
			// Assign letter and colon for consistency with server naming convention above.
			// i.e. so that server name and drive can be used without having to worry about
			// whether it needs a colon added or not.
			drive_end = drive + 2;
		else
		{
			// It's debatable, but it seems best to return a blank drive if a aFileSpec is a relative path.
			// rather than trying to use GetFullPathName() on a potentially non-existent file/dir.
			// _splitpath() doesn't fetch the drive letter of relative paths either.  This also reports
			// a blank drive for something like file://C:\My Folder\My File.txt, which seems too rarely
			// to justify a special mode.
			drive_end = _T("");
			drive = drive_end; // This is necessary to allow Assign() to work correctly later below, since it interprets a length of zero as "use string's entire length".
		}

		if (   !(name_delimiter = _tcsrchr(aFileSpec, '\\'))   ) // No backslash.
			if (   !(name_delimiter = _tcsrchr(aFileSpec, ':'))   ) // No colon.
				name_delimiter = NULL; // Indicate that there is no directory.

		name = name_delimiter ? name_delimiter + 1 : aFileSpec; // If no delimiter, name is the entire string.
	}

	// The above has now set the following variables:
	// name: As an empty string or the actual name of the file, including extension.
	// name_delimiter: As NULL if there is no directory, otherwise, the end of the directory's name.
	// drive: As the start of the drive/server name, e.g. C:, \\Workstation01, http://domain.com, etc.
	// drive_end: As the position after the drive's last character, either a zero terminator, slash, or backslash.

	if (output_var_name && !output_var_name->Assign(name))
		return FAIL;

	if (output_var_dir)
	{
		if (!name_delimiter)
			output_var_dir->Assign(); // Shouldn't fail.
		else if (*name_delimiter == '\\' || *name_delimiter == '/')
		{
			if (!output_var_dir->Assign(aFileSpec, (VarSizeType)(name_delimiter - aFileSpec)))
				return FAIL;
		}
		else // *name_delimiter == ':', e.g. "C:Some File.txt".  If aFileSpec starts with just ":",
			 // the dir returned here will also start with just ":" since that's rare & illegal anyway.
			if (!output_var_dir->Assign(aFileSpec, (VarSizeType)(name_delimiter - aFileSpec + 1)))
				return FAIL;
	}

	LPTSTR ext_dot = _tcsrchr(name, '.');
	if (output_var_ext)
	{
		// Note that the OS doesn't allow filenames to end in a period.
		if (!ext_dot)
			output_var_ext->Assign();
		else
			if (!output_var_ext->Assign(ext_dot + 1)) // Can be empty string if filename ends in just a dot.
				return FAIL;
	}

	if (output_var_name_no_ext && !output_var_name_no_ext->Assign(name, (VarSizeType)(ext_dot ? ext_dot - name : _tcslen(name))))
		return FAIL;

	if (output_var_drive && !output_var_drive->Assign(drive, (VarSizeType)(drive_end - drive)))
		return FAIL;

	return OK;
}



int SortWithOptions(const void *a1, const void *a2)
// Decided to just have one sort function since there are so many permutations.  The performance
// will be a little bit worse, but it seems simpler to implement and maintain.
// This function's input parameters are pointers to the elements of the array.  Since those elements
// are themselves pointers, the input parameters are therefore pointers to pointers (handles).
{
	LPTSTR sort_item1 = *(LPTSTR *)a1;
	LPTSTR sort_item2 = *(LPTSTR *)a2;
	if (g_SortColumnOffset > 0)
	{
		// Adjust each string (even for numerical sort) to be the right column position,
		// or the position of its zero terminator if the column offset goes beyond its length:
		size_t length = _tcslen(sort_item1);
		sort_item1 += (size_t)g_SortColumnOffset > length ? length : g_SortColumnOffset;
		length = _tcslen(sort_item2);
		sort_item2 += (size_t)g_SortColumnOffset > length ? length : g_SortColumnOffset;
	}
	if (g_SortNumeric) // Takes precedence over g_SortCaseSensitive
	{
		// For now, assume both are numbers.  If one of them isn't, it will be sorted as a zero.
		// Thus, all non-numeric items should wind up in a sequential, unsorted group.
		// Resolve only once since parts of the ATOF() macro are inline:
		double item1_minus_2 = ATOF(sort_item1) - ATOF(sort_item2);
		if (!item1_minus_2) // Exactly equal.
			return 0;
		// Otherwise, it's either greater or less than zero:
		int result = (item1_minus_2 > 0.0) ? 1 : -1;
		return g_SortReverse ? -result : result;
	}
	// Otherwise, it's a non-numeric sort.
	// v1.0.43.03: Added support the new locale-insensitive mode.
	int result = tcscmp2(sort_item1, sort_item2, g_SortCaseSensitive); // Resolve large macro only once for code size reduction.
	return g_SortReverse ? -result : result;
}



int SortByNakedFilename(const void *a1, const void *a2)
// See comments in prior function for details.
{
	LPTSTR sort_item1 = *(LPTSTR *)a1;
	LPTSTR sort_item2 = *(LPTSTR *)a2;
	LPTSTR cp;
	if (cp = _tcsrchr(sort_item1, '\\'))  // Assign
		sort_item1 = cp + 1;
	if (cp = _tcsrchr(sort_item2, '\\'))  // Assign
		sort_item2 = cp + 1;
	// v1.0.43.03: Added support the new locale-insensitive mode.
	int result = tcscmp2(sort_item1, sort_item2, g_SortCaseSensitive); // Resolve large macro only once for code size reduction.
	return g_SortReverse ? -result : result;
}



struct sort_rand_type
{
	LPTSTR cp; // This must be the first member of the struct, otherwise the array trickery in PerformSort will fail.
	union
	{
		// This must be the same size in bytes as the above, which is why it's maintained as a union with
		// a char* rather than a plain int.
		LPTSTR unused;
		int rand;
	};
};

int SortRandom(const void *a1, const void *a2)
// See comments in prior functions for details.
{
	return ((sort_rand_type *)a1)->rand - ((sort_rand_type *)a2)->rand;
}

int SortUDF(const void *a1, const void *a2)
// See comments in prior function for details.
{
	// Need to check if backup of function's variables is needed in case:
	// 1) The UDF is assigned to more than one callback, in which case the UDF could be running more than one
	//    simultaneously.
	// 2) The callback is intended to be reentrant (e.g. a subclass/WindowProc that doesn't Critical).
	// 3) Script explicitly calls the UDF in addition to using it as a callback.
	//
	// See ExpandExpression() for detailed comments about the following section.
	VarBkp *var_backup = NULL;  // If needed, it will hold an array of VarBkp objects.
	int var_backup_count; // The number of items in the above array.
	if (g_SortFunc->mInstances > 0) // Backup is needed.
		if (!Var::BackupFunctionVars(*g_SortFunc, var_backup, var_backup_count)) // Out of memory.
			return 0; // Since out-of-memory is so rare, it seems justifiable not to have any error reporting and instead just say "these items are equal".

	// The following isn't necessary because by definition, the current thread isn't paused because it's the
	// thing that called the sort in the first place.
	//g_script.UpdateTrayIcon();

	ExprTokenType result_token; // L31
	g_SortFunc->mParam[0].var->Assign(*(LPTSTR *)a1); // For simplicity and due to extreme rarity, parameters beyond
	g_SortFunc->mParam[1].var->Assign(*(LPTSTR *)a2); // the first 2 aren't populated even if they have default values.
	if (g_SortFunc->mParamCount > 2)
		g_SortFunc->mParam[2].var->Assign((__int64)(*(LPTSTR *)a2 - *(LPTSTR *)a1)); // __int64 to allow for a list greater than 2 GB, though that is currently impossible.
	g_SortFunc->Call(&result_token); // Call the UDF.

	// MUST handle return_value BEFORE calling FreeAndRestoreFunctionVars() because return_value might be
	// the contents of one of the function's local variables (which are about to be free'd).
	int returned_int;
	if (!TokenIsEmptyString(result_token)) // No need to check the following because they're implied for *return_value!=0: result != EARLY_EXIT && result != FAIL;
	{
		// Using float vs. int makes sort up to 46% slower, so decided to use int. Must use ATOI64 vs. ATOI
		// because otherwise a negative might overflow/wrap into a positive (at least with the MSVC++
		// implementation of ATOI).
		// ATOI64()'s implementation (and probably any/all others?) truncates any decimal portion;
		// e.g. 0.8 and 0.3 both yield 0.
		__int64 i64 = TokenToInt64(result_token);
		if (i64 > 0)  // Maybe there's a faster/better way to do these checks. Can't simply typecast to an int because some large positives wrap into negative, maybe vice versa.
			returned_int = 1;
		else if (i64 < 0)
			returned_int = -1;
		else
			returned_int = 0;

		if (result_token.symbol == SYM_OBJECT) // L31
			result_token.object->Release();
	}
	else
		returned_int = 0;

	Var::FreeAndRestoreFunctionVars(*g_SortFunc, var_backup, var_backup_count);
	return returned_int;
}



ResultType Line::PerformSort(LPTSTR aContents, LPTSTR aOptions)
// Caller must ensure that aContents is modifiable (ArgMustBeDereferenced() currently ensures this) because
// not only does this function modify it, it also needs to store its result back into output_var in a way
// that requires that output_var not be at the same address as the contents that were sorted.
// It seems best to treat ACT_SORT's var to be an input vs. output var because if
// it's an environment variable or the clipboard, the input variable handler will
// automatically resolve it to be ARG1 (i.e. put its contents into the deref buf).
// This is especially necessary if the clipboard contains files, in which case
// output_var->Get(), not Contents(),  must be used to resolve the filenames into text.
// And on average, using the deref buffer for this operation will not be wasteful in
// terms of expanding it unnecessarily, because usually the contents will have been
// already (or will soon be) in the deref buffer as a result of commands before
// or after the Sort command in the script.
{
	// Set defaults in case of early goto:
	LPTSTR mem_to_free = NULL;
	Func *sort_func_orig = g_SortFunc; // Because UDFs can be interrupted by other threads -- and because UDFs can themselves call Sort with some other UDF (unlikely to be sure) -- backup & restore original g_SortFunc so that the "collapsing in reverse order" behavior will automatically ensure proper operation.
	g_SortFunc = NULL; // Now that original has been saved above, reset to detect whether THIS sort uses a UDF.
	ResultType result_to_return = OK;
	DWORD ErrorLevel = -1; // Use -1 to mean "don't change/set ErrorLevel".

	// Resolve options.  First set defaults for options:
	TCHAR delimiter = '\n';
	g_SortCaseSensitive = SCS_INSENSITIVE;
	g_SortNumeric = false;
	g_SortReverse = false;
	g_SortColumnOffset = 0;
	bool trailing_delimiter_indicates_trailing_blank_item = false, terminate_last_item_with_delimiter = false
		, trailing_crlf_added_temporarily = false, sort_by_naked_filename = false, sort_random = false
		, omit_dupes = false;
	LPTSTR cp, cp_end;

	for (cp = aOptions; *cp; ++cp)
	{
		switch(_totupper(*cp))
		{
		case 'C':
			if (ctoupper(cp[1]) == 'L') // v1.0.43.03: Locale-insensitive mode, which probably performs considerably worse.
			{
				++cp;
				g_SortCaseSensitive = SCS_INSENSITIVE_LOCALE;
			}
			else
				g_SortCaseSensitive = SCS_SENSITIVE;
			break;
		case 'D':
			if (!cp[1]) // Avoids out-of-bounds when the loop's own ++cp is done.
				break;
			++cp;
			if (*cp)
				delimiter = *cp;
			break;
		case 'F': // v1.0.47: Support a callback function to extend flexibility.
			// Decided not to set ErrorLevel here because omit-dupes already uses it, and the code/docs
			// complexity of having one take precedence over the other didn't seem worth it given rarity
			// of errors and rarity of UDF use.
			cp = omit_leading_whitespace(cp + 1); // Point it to the function's name.
			if (   !(cp_end = StrChrAny(cp, _T(" \t")))   ) // Find space or tab, if any.
				cp_end = cp + _tcslen(cp); // Point it to the terminator instead.
			if (   !(g_SortFunc = g_script.FindFunc(cp, cp_end - cp))   )
				goto end; // For simplicity, just abort the sort.
			// To improve callback performance, ensure there are no ByRef parameters (for simplicity:
			// not even ones that have default values) among the first two parameters.  This avoids the
			// need to ensure formal parameters are non-aliases each time the callback is called.
			if (g_SortFunc->mIsBuiltIn || g_SortFunc->mParamCount < 2 // This validation is relied upon at a later stage.
				|| g_SortFunc->mParamCount > 3  // Reserve 4-or-more parameters for possible future use (to avoid breaking existing scripts if such features are ever added).
				|| g_SortFunc->mParam[0].is_byref || g_SortFunc->mParam[1].is_byref) // Relies on short-circuit boolean order.
				goto end; // For simplicity, just abort the sort.
			// Otherwise, the function meets the minimum constraints (though for simplicity, optional parameters
			// (default values), if any, aren't populated).
			// Fix for v1.0.47.05: The following line now subtracts 1 in case *cp_end=='\0'; otherwise the
			// loop's ++cp would go beyond the terminator when there are no more options.
			cp = cp_end - 1; // In the next iteration (which also does a ++cp), resume looking for options after the function's name.
			break;
		case 'N':
			g_SortNumeric = true;
			break;
		case 'P':
			// Use atoi() vs. ATOI() to avoid interpreting something like 0x01C as hex
			// when in fact the C was meant to be an option letter:
			g_SortColumnOffset = _ttoi(cp + 1);
			if (g_SortColumnOffset < 1)
				g_SortColumnOffset = 1;
			--g_SortColumnOffset;  // Convert to zero-based.
			break;
		case 'R':
			if (!_tcsnicmp(cp, _T("Random"), 6))
			{
				sort_random = true;
				cp += 5; // Point it to the last char so that the loop's ++cp will point to the character after it.
			}
			else
				g_SortReverse = true;
			break;
		case 'U':  // Unique.
			omit_dupes = true;
			ErrorLevel = 0; // Set default dupe-count to 0 in case of early return.
			break;
		case 'Z':
			// By setting this to true, the final item in the list, if it ends in a delimiter,
			// is considered to be followed by a blank item:
			trailing_delimiter_indicates_trailing_blank_item = true;
			break;
		case '\\':
			sort_by_naked_filename = true;
		}
	}

	// Check for early return only after parsing options in case an option that sets ErrorLevel is present:
	if (!*aContents) // Variable is empty, nothing to sort.
		goto end;
	Var &output_var = *OUTPUT_VAR; // The input var (ARG1) is also the output var in this case.
	// Do nothing for reserved variables, since most of them are read-only and besides, none
	// of them (realistically) should ever need sorting:
	if (VAR_IS_READONLY(output_var)) // output_var can be a reserved variable because it's not marked as an output-var by ArgIsVar() [since it has a dual purpose as an input-var].
		goto end;

	// size_t helps performance and should be plenty of capacity for many years of advancement.
	// In addition, things like realloc() can't accept anything larger than size_t anyway,
	// so there's no point making this 64-bit until size_t itself becomes 64-bit (it already is on some compilers?).
	size_t item_count;

	// Explicitly calculate the length in case it's the clipboard or an environment var.
	// (in which case Length() does not contain the current length).  While calculating
	// the length, also check how many delimiters are present:
	for (item_count = 1, cp = aContents; *cp; ++cp)  // Start at 1 since item_count is delimiter_count+1
		if (*cp == delimiter)
			++item_count;
	size_t aContents_length = cp - aContents;

	// If the last character in the unsorted list is a delimiter then technically the final item
	// in the list is a blank item.  However, if the options specify not to allow that, don't count that
	// blank item as an actual item (i.e. omit it from the list):
	if (!trailing_delimiter_indicates_trailing_blank_item && cp > aContents && cp[-1] == delimiter)
	{
		terminate_last_item_with_delimiter = true; // Have a later stage add the delimiter to the end of the sorted list so that the length and format of the sorted list is the same as the unsorted list.
		--item_count; // Indicate that the blank item at the end of the list isn't being counted as an item.
	}
	else // The final character isn't a delimiter or it is but there's a blank item to its right.  Either way, the following is necessary (verified correct).
	{
		if (delimiter == '\n')
		{
			LPTSTR first_delimiter = _tcschr(aContents, delimiter);
			if (first_delimiter && first_delimiter > aContents && first_delimiter[-1] == '\r')
			{
				// v1.0.47.05:
				// Here the delimiter is effectively CRLF vs. LF.  Therefore, signal a later section to append
				// an extra CRLF to simplify the code and fix bugs that existed prior to v1.0.47.05.
				// One such bug is that sorting "x`r`nx" in unique mode would previously produce two
				// x's rather than one because the first x is seen to have a `r to its right but the
				// second lacks it (due to the CRLF delimiter being simulated via LF-only).
				//
				// OLD/OBSOLETE comment from a section that was removed because it's now handled by this section:
				// Check if special handling is needed due to the following situation:
				// Delimiter is LF but the contents are lines delimited by CRLF, not just LF
				// and the original/unsorted list's last item was not terminated by an
				// "allowed delimiter".  The symptoms of this are that after the sort, the
				// last item will end in \r when it should end in no delimiter at all.
				// This happens pretty often, such as when the clipboard contains files.
				// In the future, an option letter can be added to turn off this workaround,
				// but it seems unlikely that anyone would ever want to.
				trailing_crlf_added_temporarily = true;
				terminate_last_item_with_delimiter = true; // This is done to "fool" later sections into thinking the list ends in a CRLF that doesn't have a blank item to the right, which in turn simplifies the logic and/or makes it more understandable.
			}
		}
	}

	if (item_count == 1) // 1 item is already sorted, and no dupes are possible.
	{
		// Put the exact contents back into the output_var, which is necessary in case
		// the variable was an environment variable or the clipboard-containing-files,
		// since in those cases we want the behavior to be consistent regardless of
		// whether there's only 1 item or more than one:
		// Clipboard-contains-files: The single file should be translated into its
		// text equivalent.  Var was an environment variable: the corresponding script
		// variable should be assigned the contents, so it will basically "take over"
		// for the environment variable.
		result_to_return = output_var.Assign(aContents, (VarSizeType)aContents_length);
		goto end;
	}

	// v1.0.47.05: It simplifies the code a lot to allocate and/or improves understandability to allocate
	// memory for trailing_crlf_added_temporarily even though technically it's done only to make room to
	// append the extra CRLF at the end.
	if (g_SortFunc || trailing_crlf_added_temporarily) // Do this here rather than earlier with the options parsing in case the function-option is present twice (unlikely, but it would be a memory leak due to strdup below).  Doing it here also avoids allocating if it isn't necessary.
	{
		// When g_SortFunc!=NULL, the copy of the string is needed because an earlier stage has ensured that
		// aContents is in the deref buffer, but that deref buffer is about to be overwritten by the
		// execution of the script's UDF body.
		if (   !(mem_to_free = tmalloc(aContents_length + 3))   ) // +1 for terminator and +2 in case of trailing_crlf_added_temporarily.
		{
			result_to_return = LineError(ERR_OUTOFMEM);  // Short msg. since so rare.
			goto end;
		}
		tmemcpy(mem_to_free, aContents, aContents_length + 1); // memcpy() usually benches a little faster than _tcscpy().
		aContents = mem_to_free;
		if (trailing_crlf_added_temporarily)
		{
			// Must be added early so that the next stage will terminate at the \n, leaving the \r as
			// the ending character in this item.
			_tcscpy(aContents + aContents_length, _T("\r\n"));
			aContents_length += 2;
		}
	}

	// Create the array of pointers that points into aContents to each delimited item.
	// Use item_count + 1 to allow space for the last (blank) item in case
	// trailing_delimiter_indicates_trailing_blank_item is false:
	int unit_size = sort_random ? 2 : 1;
	size_t item_size = unit_size * sizeof(LPTSTR);
	LPTSTR *item = (LPTSTR *)malloc((item_count + 1) * item_size);
	if (!item)
	{
		result_to_return = LineError(ERR_OUTOFMEM);  // Short msg. since so rare.
		goto end;
	}

	// If sort_random is in effect, the above has created an array twice the normal size.
	// This allows the random numbers to be interleaved inside the array as though it
	// were an array consisting of sort_rand_type (which it actually is when viewed that way).
	// Because of this, the array should be accessed through pointer addition rather than
	// indexing via [].

	// Scan aContents and do the following:
	// 1) Replace each delimiter with a terminator so that the individual items can be seen
	//    as real strings by the SortWithOptions() and when copying the sorted results back
	//    into output_vav.  It is safe to change aContents in this way because
	//    ArgMustBeDereferenced() has ensured that those contents are in the deref buffer.
	// 2) Store a marker/pointer to each item (string) in aContents so that we know where
	//    each item begins for sorting and recopying purposes.
	LPTSTR *item_curr = item; // i.e. Don't use [] indexing for the reason in the paragraph previous to above.
	for (item_count = 0, cp = *item_curr = aContents; *cp; ++cp)
	{
		if (*cp == delimiter)  // Each delimiter char becomes the terminator of the previous key phrase.
		{
			*cp = '\0';  // Terminate the item that appears before this delimiter.
			++item_count;
			if (sort_random)
				*(item_curr + 1) = (LPTSTR)(size_t)genrand_int31(); // i.e. the randoms are in the odd fields, the pointers in the even.
				// For the above:
				// I don't know the exact reasons, but using genrand_int31() is much more random than
				// using genrand_int32() in this case.  Perhaps it is some kind of statistical/cyclical
				// anomaly in the random number generator.  Or perhaps it's something to do with integer
				// underflow/overflow in SortRandom().  In any case, the problem can be proven via the
				// following script, which shows a sharply non-random distribution when genrand_int32()
				// is used:
				//count = 0
				//Loop 10000
				//{
				//	var = 1`n2`n3`n4`n5`n
				//	Sort, Var, Random
				//	StringLeft, Var1, Var, 1
				//	if Var1 = 5  ; Change this value to 1 to see the opposite problem.
				//		count += 1
				//}
				//Msgbox %count%
				//
				// I e-mailed the author about this sometime around/prior to 12/1/04 but never got a response.
			item_curr += unit_size; // i.e. Don't use [] indexing for the reason described above.
			*item_curr = cp + 1; // Make a pointer to the next item's place in aContents.
		}
	}
	// The above reset the count to 0 and recounted it.  So now re-add the last item to the count unless it was
	// disqualified earlier. Verified correct:
	if (!terminate_last_item_with_delimiter) // i.e. either trailing_delimiter_indicates_trailing_blank_item==true OR the final character isn't a delimiter. Either way the final item needs to be added.
	{
		++item_count;
		if (sort_random) // Provide a random number for the last item.
			*(item_curr + 1) = (LPTSTR)(size_t)genrand_int31(); // i.e. the randoms are in the odd fields, the pointers in the even.
	}
	else // Since the final item is not included in the count, point item_curr to the one before the last, for use below.
		item_curr -= unit_size;

	// Now aContents has been divided up based on delimiter.  Sort the array of pointers
	// so that they indicate the correct ordering to copy aContents into output_var:
	if (g_SortFunc) // Takes precedence other sorting methods.
		qsort((void *)item, item_count, item_size, SortUDF);
	else if (sort_random) // Takes precedence over all remaining options.
		qsort((void *)item, item_count, item_size, SortRandom);
	else
		qsort((void *)item, item_count, item_size, sort_by_naked_filename ? SortByNakedFilename : SortWithOptions);

	// Copy the sorted pointers back into output_var, which might not already be sized correctly
	// if it's the clipboard or it was an environment variable when it came in as the input.
	// If output_var is the clipboard, this call will set up the clipboard for writing:
	if (output_var.AssignString(NULL, aContents_length) != OK) // Might fail due to clipboard problem.
	{
		result_to_return = FAIL;
		goto end;
	}

	// Set default in case original last item is still the last item, or if last item was omitted due to being a dupe:
	size_t i, item_count_minus_1 = item_count - 1;
	DWORD omit_dupe_count = 0;
	bool keep_this_item;
	LPTSTR source, dest;
	LPTSTR item_prev = NULL;

	// Copy the sorted result back into output_var.  Do all except the last item, since the last
	// item gets special treatment depending on the options that were specified.  The call to
	// output_var->Contents() below should never fail due to the above having prepped it:
	item_curr = item; // i.e. Don't use [] indexing for the reason described higher above (same applies to item += unit_size below).
	for (dest = output_var.Contents(), i = 0; i < item_count; ++i, item_curr += unit_size)
	{
		keep_this_item = true;  // Set default.
		if (omit_dupes && item_prev)
		{
			// Update to the comment below: Exact dupes will still be removed when sort_by_naked_filename
			// or g_SortColumnOffset is in effect because duplicate lines would still be adjacent to
			// each other even in these modes.  There doesn't appear to be any exceptions, even if
			// some items in the list are sorted as blanks due to being shorter than the specified 
			// g_SortColumnOffset.
			// As documented, special dupe-checking modes are not offered when sort_by_naked_filename
			// is in effect, or g_SortColumnOffset is greater than 1.  That's because the need for such
			// a thing seems too rare (and the result too strange) to justify the extra code size.
			// However, adjacent dupes are still removed when any of the above modes are in effect,
			// or when the "random" mode is in effect.  This might have some usefulness; for example,
			// if a list of songs is sorted in random order, but it contains "favorite" songs listed twice,
			// the dupe-removal feature would remove duplicate songs if they happen to be sorted
			// to lie adjacent to each other, which would be useful to prevent the same song from
			// playing twice in a row.
			if (g_SortNumeric && !g_SortColumnOffset)
				// if g_SortColumnOffset is zero, fall back to the normal dupe checking in case its
				// ever useful to anyone.  This is done because numbers in an offset column are not supported
				// since the extra code size doensn't seem justified given the rarity of the need.
				keep_this_item = (ATOF(*item_curr) != ATOF(item_prev)); // ATOF() ignores any trailing \r in CRLF mode, so no extra logic is needed for that.
			else
				keep_this_item = tcscmp2(*item_curr, item_prev, g_SortCaseSensitive); // v1.0.43.03: Added support for locale-insensitive mode.
				// Permutations of sorting case sensitive vs. eliminating duplicates based on case sensitivity:
				// 1) Sort is not case sens, but dupes are: Won't work because sort didn't necessarily put
				//    same-case dupes adjacent to each other.
				// 2) Converse: probably not reliable because there could be unrelated items in between
				//    two strings that are duplicates but weren't sorted adjacently due to their case.
				// 3) Both are case sensitive: seems okay
				// 4) Both are not case sensitive: seems okay
				//
				// In light of the above, using the g_SortCaseSensitive flag to control the behavior of
				// both sorting and dupe-removal seems best.
		}
		if (keep_this_item)
		{
			for (source = *item_curr; *source;)
				*dest++ = *source++;
			// If we're at the last item and the original list's last item had a terminating delimiter
			// and the specified options said to treat it not as a delimiter but as a final char of sorts,
			// include it after the item that is now last so that the overall layout is the same:
			if (i < item_count_minus_1 || terminate_last_item_with_delimiter)
				*dest++ = delimiter;  // Put each item's delimiter back in so that format is the same as the original.
			item_prev = *item_curr; // Since the item just processed above isn't a dupe, save this item to compare against the next item.
		}
		else // This item is a duplicate of the previous item.
		{
			++omit_dupe_count; // But don't change the value of item_prev.
			// v1.0.47.05: The following section fixes the fact that the "unique" option would sometimes leave
			// a trailing delimiter at the end of the sorted list.  For example, sorting "x|x" in unique mode
			// would formerly produce "x|":
			if (i == item_count_minus_1 && !terminate_last_item_with_delimiter) // This is the last item and its being omitted, so remove the previous item's trailing delimiter (unless a trailing delimiter is mandatory).
				--dest; // Remove the previous item's trailing delimiter there's nothing for it to delimit due to omission of this duplicate.
		}
	} // for()
	free(item); // Free the index/pointer list used for the sort.

	// Terminate the variable's contents.
	if (trailing_crlf_added_temporarily) // Remove the CRLF only after its presence was used above to simplify the code by reducing the number of types/cases.
	{
		dest[-2] = '\0';
		output_var.ByteLength() -= 2 * sizeof(TCHAR);
	}
	else
		*dest = '\0';

	if (omit_dupes)
	{
		if (omit_dupe_count) // Update the length to actual whenever at least one dupe was omitted.
		{
			output_var.SetCharLength((VarSizeType)_tcslen(output_var.Contents()));
			ErrorLevel = omit_dupe_count; // Override the 0 set earlier.
		}
	}
	//else it is not necessary to set output_var.Length() here because its length hasn't changed
	// since it was originally set by the above call "output_var.AssignString(NULL..."

	result_to_return = output_var.Close(); // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_CLIPBOARD.

end:
	if (ErrorLevel != -1) // A change to ErrorLevel is desired.  Compare directly to -1 due to unsigned.
		g_ErrorLevel->Assign(ErrorLevel); // ErrorLevel is set only when dupe-mode is in effect.
	if (mem_to_free)
		free(mem_to_free);
	g_SortFunc = sort_func_orig;
	return result_to_return;
}



ResultType Line::GetKeyJoyState(LPTSTR aKeyName, LPTSTR aOption)
// Keep this in sync with FUNC_GETKEYSTATE.
{
	Var &output_var = *OUTPUT_VAR;
	JoyControls joy;
	int joystick_id;
	vk_type vk = TextToVK(aKeyName);
    if (!vk)
	{
		if (   !(joy = (JoyControls)ConvertJoy(aKeyName, &joystick_id))   )
			return output_var.Assign(_T(""));
		// Since the above didn't return, joy contains a valid joystick button/control ID.
		// Caller needs a token with a buffer of at least this size:
		TCHAR buf[MAX_NUMBER_SIZE];
		ExprTokenType token;
		token.symbol = SYM_STRING; // These must be set as defaults for ScriptGetJoyState().
		token.marker = buf;        //
		ScriptGetJoyState(joy, joystick_id, token, false);
		// Above: ScriptGetJoyState() returns FAIL and sets output_var to be blank if the result is
		// indeterminate or there was a problem reading the joystick.  But we don't want such a failure
		// to be considered a "critical failure" that will exit the current quasi-thread, so its result
		// is discarded.
		return output_var.Assign(token); // Write the result based on whether the token is a string or number.
	}
	// Otherwise: There is a virtual key (not a joystick control).
	KeyStateTypes key_state_type;
	switch (ctoupper(*aOption))
	{
	case 'T': key_state_type = KEYSTATE_TOGGLE; break; // Whether a toggleable key such as CapsLock is currently turned on.
	case 'P': key_state_type = KEYSTATE_PHYSICAL; break; // Physical state of key.
	default: key_state_type = KEYSTATE_LOGICAL;
	}
	return output_var.Assign(ScriptGetKeyState(vk, key_state_type) ? _T("D") : _T("U"));
}



ResultType Line::DriveSpace(LPTSTR aPath, bool aGetFreeSpace)
// Because of NTFS's ability to mount volumes into a directory, a path might not necessarily
// have the same amount of free space as its root drive.  However, I'm not sure if this
// method here actually takes that into account.
{
	OUTPUT_VAR->Assign(); // Init to empty string regardless of whether we succeed here.

	if (!aPath || !*aPath) goto error;  // Below relies on this check.

	TCHAR buf[MAX_PATH + 1];  // +1 to allow appending of backslash.
	tcslcpy(buf, aPath, _countof(buf));
	size_t length = _tcslen(buf);
	if (buf[length - 1] != '\\') // Trailing backslash is present, which some of the API calls below don't like.
	{
		if (length + 1 >= _countof(buf)) // No room to fix it.
			goto error;
		buf[length++] = '\\';
		buf[length] = '\0';
	}

	// The program won't launch at all on Win95a (original Win95) unless the function address is resolved
	// at runtime:
	typedef BOOL (WINAPI *GetDiskFreeSpaceExType)(LPCTSTR, PULARGE_INTEGER, PULARGE_INTEGER, PULARGE_INTEGER);
	static GetDiskFreeSpaceExType MyGetDiskFreeSpaceEx =
		(GetDiskFreeSpaceExType)GetProcAddress(GetModuleHandle(_T("kernel32")), "GetDiskFreeSpaceEx" WINAPI_SUFFIX);

	// MSDN: "The GetDiskFreeSpaceEx function returns correct values for all volumes, including those
	// that are greater than 2 gigabytes."
	__int64 free_space;
	if (MyGetDiskFreeSpaceEx)  // Function is available (unpatched Win95 and WinNT might not have it).
	{
		ULARGE_INTEGER total, free, used;
		if (!MyGetDiskFreeSpaceEx(buf, &free, &total, &used))
			goto error;
		// Casting this way allows sizes of up to 2,097,152 gigabytes:
		free_space = (__int64)((unsigned __int64)(aGetFreeSpace ? free.QuadPart : total.QuadPart)
			/ (1024*1024));
	}
	else // For unpatched versions of Win95/NT4, fall back to compatibility mode.
	{
		DWORD sectors_per_cluster, bytes_per_sector, free_clusters, total_clusters;
		if (!GetDiskFreeSpace(buf, &sectors_per_cluster, &bytes_per_sector, &free_clusters, &total_clusters))
			goto error;
		free_space = (__int64)((unsigned __int64)((aGetFreeSpace ? free_clusters : total_clusters)
			* sectors_per_cluster * bytes_per_sector) / (1024*1024));
	}

	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return OUTPUT_VAR->Assign(free_space);

error:
	return SetErrorLevelOrThrow();
}



ResultType Line::Drive(LPTSTR aCmd, LPTSTR aValue, LPTSTR aValue2) // aValue not aValue1, for use with a shared macro.
{
	DriveCmds drive_cmd = ConvertDriveCmd(aCmd);

	TCHAR path[MAX_PATH + 1];  // +1 to allow room for trailing backslash in case it needs to be added.
	size_t path_length;

	// Notes about the below macro:
	// - It adds a backslash to the contents of the path variable because certain API calls or OS versions
	//   might require it.
	// - It is used by both Drive() and DriveGet().
	// - Leave space for the backslash in case its needed.
	#define DRIVE_SET_PATH \
		tcslcpy(path, aValue, _countof(path) - 1);\
		path_length = _tcslen(path);\
		if (path_length && path[path_length - 1] != '\\')\
			path[path_length++] = '\\';

	switch(drive_cmd)
	{
	case DRIVE_CMD_INVALID:
		// Since command names are validated at load-time, this only happens if the command name
		// was contained in a variable reference.  Since that is very rare, just set ErrorLevel
		// and return:
		return SetErrorLevelOrThrow();

	case DRIVE_CMD_LOCK:
	case DRIVE_CMD_UNLOCK:
		return SetErrorLevelOrThrowBool(!DriveLock(*aValue, drive_cmd == DRIVE_CMD_LOCK));

	case DRIVE_CMD_EJECT:
		// Don't do DRIVE_SET_PATH in this case since trailing backslash might prevent it from
		// working on some OSes.
		// It seems best not to do the below check since:
		// 1) aValue usually lacks a trailing backslash so that it will work correctly with "open c: type cdaudio".
		//    That lack might prevent DriveGetType() from working on some OSes.
		// 2) It's conceivable that tray eject/retract might work on certain types of drives even though
		//    they aren't of type DRIVE_CDROM.
		// 3) One or both of the calls to mciSendString() will simply fail if the drive isn't of the right type.
		//if (GetDriveType(aValue) != DRIVE_CDROM) // Testing reveals that the below method does not work on Network CD/DVD drives.
		//	return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		TCHAR mci_string[256];
		MCIERROR error;
		// Note: The following comment is obsolete because research of MSDN indicates that there is no way
		// not to wait when the tray must be physically opened or closed, at least on Windows XP.  Omitting
		// the word "wait" from both "close cd wait" and "set cd door open/closed wait" does not help, nor
		// does replacing wait with the word notify in "set cdaudio door open/closed wait".
		// The word "wait" is always specified with these operations to ensure consistent behavior across
		// all OSes (on the off-chance that the absence of "wait" really avoids waiting on Win9x or future
		// OSes, or perhaps under certain conditions or for certain types of drives).  See above comment
		// for details.
		if (!*aValue) // When drive is omitted, operate upon default CD/DVD drive.
		{
			sntprintf(mci_string, _countof(mci_string), _T("set cdaudio door %s wait"), ATOI(aValue2) == 1 ? _T("closed") : _T("open"));
			error = mciSendString(mci_string, NULL, 0, NULL); // Open or close the tray.
			return SetErrorLevelOrThrowBool(error);
		}
		sntprintf(mci_string, _countof(mci_string), _T("open %s type cdaudio alias cd wait shareable"), aValue);
		if (mciSendString(mci_string, NULL, 0, NULL)) // Error.
			return SetErrorLevelOrThrow();
		sntprintf(mci_string, _countof(mci_string), _T("set cd door %s wait"), ATOI(aValue2) == 1 ? _T("closed") : _T("open"));
		error = mciSendString(mci_string, NULL, 0, NULL); // Open or close the tray.
		mciSendString(_T("close cd wait"), NULL, 0, NULL);
		return SetErrorLevelOrThrowBool(error);

	case DRIVE_CMD_LABEL: // Note that it is possible and allowed for the new label to be blank.
		DRIVE_SET_PATH
		return SetErrorLevelOrThrowBool(!SetVolumeLabel(path, aValue2));

	} // switch()

	return FAIL;  // Should never be executed.  Helps catch bugs.
}



ResultType Line::DriveLock(TCHAR aDriveLetter, bool aLockIt)
{
	HANDLE hdevice;
	DWORD unused;
	BOOL result;

	if (g_os.IsWin9x())
	{
		// blisteringhot@hotmail.com has confirmed that the code below works on Win98 with an IDE CD Drive:
		// System:  Win98 IDE CdRom (my ejector is CloseTray)
		// I get a blue screen when I try to eject after using the test script.
		// "eject request to drive in use"
		// It asks me to Ok or Esc, Ok is default.
		//	-probably a bit scary for a novice.
		// BUT its locked alright!"

		// Use the Windows 9x method.  The code below is based on an example posted by Microsoft.
		// Note: The presence of the code below does not add a detectible amount to the EXE size
		// (probably because it's mostly defines and data types).
		#pragma pack(push, 1)
		typedef struct _DIOC_REGISTERS
		{
			DWORD reg_EBX;
			DWORD reg_EDX;
			DWORD reg_ECX;
			DWORD reg_EAX;
			DWORD reg_EDI;
			DWORD reg_ESI;
			DWORD reg_Flags;
		} DIOC_REGISTERS, *PDIOC_REGISTERS;
		typedef struct _PARAMBLOCK
		{
			BYTE Operation;
			BYTE NumLocks;
		} PARAMBLOCK, *PPARAMBLOCK;
		#pragma pack(pop)

		// MS: Prepare for lock or unlock IOCTL call
		#define CARRY_FLAG 0x1
		#define VWIN32_DIOC_DOS_IOCTL 1
		#define LOCK_MEDIA   0
		#define UNLOCK_MEDIA 1
		#define STATUS_LOCK  2
		PARAMBLOCK pb = {0};
		pb.Operation = aLockIt ? LOCK_MEDIA : UNLOCK_MEDIA;
		
		DIOC_REGISTERS regs = {0};
		regs.reg_EAX = 0x440D;
		regs.reg_EBX = ctoupper(aDriveLetter) - 'A' + 1; // Convert to drive index. 0 = default, 1 = A, 2 = B, 3 = C
		regs.reg_ECX = 0x0848; // MS: Lock/unlock media
		regs.reg_EDX = (DWORD)(size_t)&pb;
		
		// MS: Open VWIN32
		hdevice = CreateFile(_T("\\\\.\\vwin32"), 0, 0, NULL, 0, FILE_FLAG_DELETE_ON_CLOSE, NULL);
		if (hdevice == INVALID_HANDLE_VALUE)
			return FAIL;
		
		// MS: Call VWIN32
		result = DeviceIoControl(hdevice, VWIN32_DIOC_DOS_IOCTL, &regs, sizeof(regs), &regs, sizeof(regs), &unused, 0);
		if (result)
			result = !(regs.reg_Flags & CARRY_FLAG);
	}
	else // NT4/2k/XP or later
	{
		// The calls below cannot work on Win9x (as documented by MSDN's PREVENT_MEDIA_REMOVAL).
		// Don't even attempt them on Win9x because they might blow up.
		TCHAR filename[64];
		_stprintf(filename, _T("\\\\.\\%c:"), aDriveLetter);
		// FILE_READ_ATTRIBUTES is not enough; it yields "Access Denied" error.  So apparently all or part
		// of the sub-attributes in GENERIC_READ are needed.  An MSDN example implies that GENERIC_WRITE is
		// only needed for GetDriveType() == DRIVE_REMOVABLE drives, and maybe not even those when all we
		// want to do is lock/unlock the drive (that example did quite a bit more).  In any case, research
		// indicates that all CD/DVD drives are ever considered DRIVE_CDROM, not DRIVE_REMOVABLE.
		// Due to this and the unlikelihood that GENERIC_WRITE is ever needed anyway, GetDriveType() is
		// not called for the purpose of conditionally adding the GENERIC_WRITE attribute.
		hdevice = CreateFile(filename, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
		if (hdevice == INVALID_HANDLE_VALUE)
			return FAIL;
		PREVENT_MEDIA_REMOVAL pmr;
		pmr.PreventMediaRemoval = aLockIt;
		result = DeviceIoControl(hdevice, IOCTL_STORAGE_MEDIA_REMOVAL, &pmr, sizeof(PREVENT_MEDIA_REMOVAL)
			, NULL, 0, &unused, NULL);
	}
	CloseHandle(hdevice);
	return result ? OK : FAIL;
}



ResultType Line::DriveGet(LPTSTR aCmd, LPTSTR aValue)
{
	DriveGetCmds drive_get_cmd = ConvertDriveGetCmd(aCmd);
	if (drive_get_cmd == DRIVEGET_CMD_CAPACITY)
		return DriveSpace(aValue, false);

	TCHAR path[MAX_PATH + 1];  // +1 to allow room for trailing backslash in case it needs to be added.
	size_t path_length;

	if (drive_get_cmd == DRIVEGET_CMD_SETLABEL) // The is retained for backward compatibility even though the Drive cmd is normally used.
	{
		DRIVE_SET_PATH
		LPTSTR new_label = omit_leading_whitespace(aCmd + 9);  // Example: SetLabel:MyLabel
		return g_ErrorLevel->Assign(SetVolumeLabel(path, new_label) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
	}

	Var &output_var = *OUTPUT_VAR;

	output_var.Assign(); // Init to empty string.

	switch(drive_get_cmd)
	{

	case DRIVEGET_CMD_INVALID:
		// Since command names are validated at load-time, this only happens if the command name
		// was contained in a variable reference.  Since that is very rare, just set ErrorLevel
		// and return:
		goto error;

	case DRIVEGET_CMD_LIST:
	{
		UINT drive_type;
		#define ALL_DRIVE_TYPES 256
		if (!*aValue) drive_type = ALL_DRIVE_TYPES;
		else if (!_tcsicmp(aValue, _T("CDRom"))) drive_type = DRIVE_CDROM;
		else if (!_tcsicmp(aValue, _T("Removable"))) drive_type = DRIVE_REMOVABLE;
		else if (!_tcsicmp(aValue, _T("Fixed"))) drive_type = DRIVE_FIXED;
		else if (!_tcsicmp(aValue, _T("Network"))) drive_type = DRIVE_REMOTE;
		else if (!_tcsicmp(aValue, _T("Ramdisk"))) drive_type = DRIVE_RAMDISK;
		else if (!_tcsicmp(aValue, _T("Unknown"))) drive_type = DRIVE_UNKNOWN;
		else
			goto error;

		TCHAR found_drives[32];  // Need room for all 26 possible drive letters.
		int found_drives_count;
		UCHAR letter;
		TCHAR buf[128], *buf_ptr;

		for (found_drives_count = 0, letter = 'A'; letter <= 'Z'; ++letter)
		{
			buf_ptr = buf;
			*buf_ptr++ = letter;
			*buf_ptr++ = ':';
			*buf_ptr++ = '\\';
			*buf_ptr = '\0';
			UINT this_type = GetDriveType(buf);
			if (this_type == drive_type || (drive_type == ALL_DRIVE_TYPES && this_type != DRIVE_NO_ROOT_DIR))
				found_drives[found_drives_count++] = letter;  // Store just the drive letters.
		}
		found_drives[found_drives_count] = '\0';  // Terminate the string of found drive letters.
		output_var.Assign(found_drives);
		if (!*found_drives)
			goto error;  // Seems best to flag zero drives in the system as ErrorLevel of "1".
		break;
	}

	case DRIVEGET_CMD_FILESYSTEM:
	case DRIVEGET_CMD_LABEL:
	case DRIVEGET_CMD_SERIAL:
	{
		TCHAR volume_name[256];
		TCHAR file_system[256];
		DRIVE_SET_PATH
		DWORD serial_number, max_component_length, file_system_flags;
		if (!GetVolumeInformation(path, volume_name, _countof(volume_name) - 1, &serial_number, &max_component_length
			, &file_system_flags, file_system, _countof(file_system) - 1))
			goto error;
		switch(drive_get_cmd)
		{
		case DRIVEGET_CMD_FILESYSTEM: output_var.Assign(file_system); break;
		case DRIVEGET_CMD_LABEL: output_var.Assign(volume_name); break;
		case DRIVEGET_CMD_SERIAL: output_var.Assign(serial_number); break;
		}
		break;
	}

	case DRIVEGET_CMD_TYPE:
	{
		DRIVE_SET_PATH
		switch (GetDriveType(path))
		{
		case DRIVE_UNKNOWN:   output_var.Assign(_T("Unknown")); break;
		case DRIVE_REMOVABLE: output_var.Assign(_T("Removable")); break;
		case DRIVE_FIXED:     output_var.Assign(_T("Fixed")); break;
		case DRIVE_REMOTE:    output_var.Assign(_T("Network")); break;
		case DRIVE_CDROM:     output_var.Assign(_T("CDROM")); break;
		case DRIVE_RAMDISK:   output_var.Assign(_T("RAMDisk")); break;
		default: // DRIVE_NO_ROOT_DIR
			goto error;
		}
		break;
	}

	case DRIVEGET_CMD_STATUS:
	{
		DRIVE_SET_PATH
		DWORD sectors_per_cluster, bytes_per_sector, free_clusters, total_clusters;
		switch (GetDiskFreeSpace(path, &sectors_per_cluster, &bytes_per_sector, &free_clusters, &total_clusters)
			? ERROR_SUCCESS : GetLastError())
		{
		case ERROR_SUCCESS:        output_var.Assign(_T("Ready")); break;
		case ERROR_PATH_NOT_FOUND: output_var.Assign(_T("Invalid")); break;
		case ERROR_NOT_READY:      output_var.Assign(_T("NotReady")); break;
		case ERROR_WRITE_PROTECT:  output_var.Assign(_T("ReadOnly")); break;
		default:                   output_var.Assign(_T("Unknown"));
		}
		break;
	}

	case DRIVEGET_CMD_STATUSCD:
		// Don't do DRIVE_SET_PATH in this case since trailing backslash might prevent it from
		// working on some OSes.
		// It seems best not to do the below check since:
		// 1) aValue usually lacks a trailing backslash so that it will work correctly with "open c: type cdaudio".
		//    That lack might prevent DriveGetType() from working on some OSes.
		// 2) It's conceivable that tray eject/retract might work on certain types of drives even though
		//    they aren't of type DRIVE_CDROM.
		// 3) One or both of the calls to mciSendString() will simply fail if the drive isn't of the right type.
		//if (GetDriveType(aValue) != DRIVE_CDROM) // Testing reveals that the below method does not work on Network CD/DVD drives.
		//	return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		TCHAR mci_string[256], status[128];
		// Note that there is apparently no way to determine via mciSendString() whether the tray is ejected
		// or not, since "open" is returned even when the tray is closed but there is no media.
		if (!*aValue) // When drive is omitted, operate upon default CD/DVD drive.
		{
			if (mciSendString(_T("status cdaudio mode"), status, _countof(status), NULL)) // Error.
				goto error;
		}
		else // Operate upon a specific drive letter.
		{
			sntprintf(mci_string, _countof(mci_string), _T("open %s type cdaudio alias cd wait shareable"), aValue);
			if (mciSendString(mci_string, NULL, 0, NULL)) // Error.
				goto error;
			MCIERROR error = mciSendString(_T("status cd mode"), status, _countof(status), NULL);
			mciSendString(_T("close cd wait"), NULL, 0, NULL);
			if (error)
				goto error;
		}
		// Otherwise, success:
		output_var.Assign(status);
		break;

	} // switch()

	// Note that ControlDelay is not done for the Get type commands, because it seems unnecessary.
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.

error:
	return SetErrorLevelOrThrow();
}



ResultType Line::SoundSetGet(LPTSTR aSetting, LPTSTR aComponentType, LPTSTR aControlType, LPTSTR aDevice)
// If the caller specifies NULL for aSetting, the mode will be "Get".  Otherwise, it will be "Set".
{
	int instance_number = 1;  // Set default.
	DWORD component_type = *ARG2 ? SoundConvertComponentType(ARG2, &instance_number) : MIXERLINE_COMPONENTTYPE_DST_SPEAKERS;
	DWORD control_type = *ARG3 ? SoundConvertControlType(ARG3) : MIXERCONTROL_CONTROLTYPE_VOLUME;

	#define SOUND_MODE_IS_SET aSetting // Boolean: i.e. if it's not NULL, the mode is "SET".
	if (!SOUND_MODE_IS_SET)
		OUTPUT_VAR->Assign(); // Init to empty string regardless of whether we succeed here.

	// Rare, since load-time validation would have caught problems unless the params were variable references.
	// Text values for ErrorLevels should be kept below 64 characters in length so that the variable doesn't
	// have to be expanded with a different memory allocation method:
	if (control_type == MIXERCONTROL_CONTROLTYPE_INVALID || aComponentType == MIXERLINE_COMPONENTTYPE_DST_UNDEFINED)
		return SetErrorLevelOrThrowStr(_T("Invalid Control Type or Component Type"));

	if (g_os.IsWinVistaOrLater())
		return SoundSetGetVista(aSetting, component_type, instance_number, control_type, aDevice);
	else
		return SoundSetGet2kXP(aSetting, component_type, instance_number, control_type, aDevice);
}


ResultType Line::SoundSetGet2kXP(LPTSTR aSetting, DWORD aComponentType, int aComponentInstance
	, DWORD aControlType, LPTSTR aDevice)
{
	int aMixerID = *aDevice ? ATOI(aDevice) - 1 : 0;
	if (aMixerID < 0)
		aMixerID = 0;

	double setting_percent;
	Var *output_var;
	if (SOUND_MODE_IS_SET)
	{
		output_var = NULL; // To help catch bugs.
		setting_percent = ATOF(aSetting);
		if (setting_percent < -100)
			setting_percent = -100;
		else if (setting_percent > 100)
			setting_percent = 100;
	}
	else // The mode is GET.
	{
		output_var = OUTPUT_VAR;
	}

	// Open the specified mixer ID:
	HMIXER hMixer;
    if (mixerOpen(&hMixer, aMixerID, 0, 0, 0) != MMSYSERR_NOERROR)
		return SetErrorLevelOrThrowStr(_T("Can't Open Specified Mixer"));

	// Find out how many destinations are available on this mixer (should always be at least one):
	int dest_count;
	MIXERCAPS mxcaps;
	if (mixerGetDevCaps((UINT_PTR)hMixer, &mxcaps, sizeof(mxcaps)) == MMSYSERR_NOERROR)
		dest_count = mxcaps.cDestinations;
	else
		dest_count = 1;  // Assume it has one so that we can try to proceed anyway.

	// Find specified line (aComponentType + aComponentInstance):
	MIXERLINE ml = {0};
    ml.cbStruct = sizeof(ml);
	if (aComponentInstance == 1)  // Just get the first line of this type, the easy way.
	{
		ml.dwComponentType = aComponentType;
		if (mixerGetLineInfo((HMIXEROBJ)hMixer, &ml, MIXER_GETLINEINFOF_COMPONENTTYPE) != MMSYSERR_NOERROR)
		{
			mixerClose(hMixer);
			return SetErrorLevelOrThrowStr(_T("Mixer Doesn't Support This Component Type"));
		}
	}
	else
	{
		// Search through each source of each destination, looking for the indicated instance
		// number for the indicated component type:
		int source_count;
		bool found = false;
		for (int d = 0, found_instance = 0; d < dest_count && !found; ++d) // For each destination of this mixer.
		{
			ml.dwDestination = d;
			if (mixerGetLineInfo((HMIXEROBJ)hMixer, &ml, MIXER_GETLINEINFOF_DESTINATION) != MMSYSERR_NOERROR)
				// Keep trying in case the others can be retrieved.
				continue;
			source_count = ml.cConnections;  // Make a copy of this value so that the struct can be reused.
			for (int s = 0; s < source_count && !found; ++s) // For each source of this destination.
			{
				ml.dwDestination = d; // Set it again in case it was changed.
				ml.dwSource = s;
				if (mixerGetLineInfo((HMIXEROBJ)hMixer, &ml, MIXER_GETLINEINFOF_SOURCE) != MMSYSERR_NOERROR)
					// Keep trying in case the others can be retrieved.
					continue;
				// This line can be used to show a soundcard's component types (match them against mmsystem.h):
				//MsgBox(ml.dwComponentType);
				if (ml.dwComponentType == aComponentType)
				{
					++found_instance;
					if (found_instance == aComponentInstance)
						found = true;
				}
			} // inner for()
		} // outer for()
		if (!found)
		{
			mixerClose(hMixer);
			return SetErrorLevelOrThrowStr(_T("Mixer Doesn't Have That Many of That Component Type"));
		}
	}

	// Find the mixer control (aControlType) for the above component:
    MIXERCONTROL mc; // MSDN: "No initialization of the buffer pointed to by [pamxctrl below] is required"
    MIXERLINECONTROLS mlc;
	mlc.cbStruct = sizeof(mlc);
	mlc.pamxctrl = &mc;
	mlc.cbmxctrl = sizeof(mc);
	mlc.dwLineID = ml.dwLineID;
	mlc.dwControlType = aControlType;
	mlc.cControls = 1;
	if (mixerGetLineControls((HMIXEROBJ)hMixer, &mlc, MIXER_GETLINECONTROLSF_ONEBYTYPE) != MMSYSERR_NOERROR)
	{
		mixerClose(hMixer);
		return SetErrorLevelOrThrowStr(_T("Component Doesn't Support This Control Type"));
	}

	// Does user want to adjust the current setting by a certain amount?
	// For v1.0.25, the first char of RAW_ARG is also checked in case this is an expression intended
	// to be a positive offset, such as +(var + 10)
	bool adjust_current_setting = aSetting && (*aSetting == '-' || *aSetting == '+' || *RAW_ARG1 == '+');

	// These are used in more than once place, so always initialize them here:
	MIXERCONTROLDETAILS mcd = {0};
    MIXERCONTROLDETAILS_UNSIGNED mcdMeter;
	mcd.cbStruct = sizeof(MIXERCONTROLDETAILS);
	mcd.dwControlID = mc.dwControlID;
	mcd.cChannels = 1; // MSDN: "when an application needs to get and set all channels as if they were uniform"
	mcd.paDetails = &mcdMeter;
	mcd.cbDetails = sizeof(mcdMeter);

	// Get the current setting of the control, if necessary:
	if (!SOUND_MODE_IS_SET || adjust_current_setting)
	{
		if (mixerGetControlDetails((HMIXEROBJ)hMixer, &mcd, MIXER_GETCONTROLDETAILSF_VALUE) != MMSYSERR_NOERROR)
		{
			mixerClose(hMixer);
			return SetErrorLevelOrThrowStr(_T("Can't Get Current Setting"));
		}
	}

	bool control_type_is_boolean;
	switch (aControlType)
	{
	case MIXERCONTROL_CONTROLTYPE_ONOFF:
	case MIXERCONTROL_CONTROLTYPE_MUTE:
	case MIXERCONTROL_CONTROLTYPE_MONO:
	case MIXERCONTROL_CONTROLTYPE_LOUDNESS:
	case MIXERCONTROL_CONTROLTYPE_STEREOENH:
	case MIXERCONTROL_CONTROLTYPE_BASS_BOOST:
		control_type_is_boolean = true;
		break;
	default: // For all others, assume the control can have more than just ON/OFF as its allowed states.
		control_type_is_boolean = false;
	}

	if (SOUND_MODE_IS_SET)
	{
		if (control_type_is_boolean)
		{
			if (adjust_current_setting) // The user wants this toggleable control to be toggled to its opposite state:
				mcdMeter.dwValue = (mcdMeter.dwValue > mc.Bounds.dwMinimum) ? mc.Bounds.dwMinimum : mc.Bounds.dwMaximum;
			else // Set the value according to whether the user gave us a setting that is greater than zero:
				mcdMeter.dwValue = (setting_percent > 0.0) ? mc.Bounds.dwMaximum : mc.Bounds.dwMinimum;
		}
		else // For all others, assume the control can have more than just ON/OFF as its allowed states.
		{
			// Make this an __int64 vs. DWORD to avoid underflow (so that a setting_percent of -100
			// is supported whenever the difference between Min and Max is large, such as MAXDWORD):
			__int64 specified_vol = (__int64)((mc.Bounds.dwMaximum - mc.Bounds.dwMinimum) * (setting_percent / 100.0));
			if (adjust_current_setting)
			{
				// Make it a big int so that overflow/underflow can be detected:
				__int64 vol_new = mcdMeter.dwValue + specified_vol;
				if (vol_new < mc.Bounds.dwMinimum)
					vol_new = mc.Bounds.dwMinimum;
				else if (vol_new > mc.Bounds.dwMaximum)
					vol_new = mc.Bounds.dwMaximum;
				mcdMeter.dwValue = (DWORD)vol_new;
			}
			else
				mcdMeter.dwValue = (DWORD)specified_vol; // Due to the above, it's known to be positive in this case.
		}

		MMRESULT result = mixerSetControlDetails((HMIXEROBJ)hMixer, &mcd, MIXER_GETCONTROLDETAILSF_VALUE);
		mixerClose(hMixer);
		return SetErrorLevelOrThrowStr(result == MMSYSERR_NOERROR ? ERRORLEVEL_NONE : _T("Can't Change Setting"));
	}

	// Otherwise, the mode is "Get":
	mixerClose(hMixer);
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.

	if (control_type_is_boolean)
		return output_var->Assign(mcdMeter.dwValue ? _T("On") : _T("Off"));
	else // For all others, assume the control can have more than just ON/OFF as its allowed states.
		// The MSDN docs imply that values fetched via the above method do not distinguish between
		// left and right volume levels, unlike waveOutGetVolume():
		return output_var->Assign(   ((double)100 * (mcdMeter.dwValue - (DWORD)mc.Bounds.dwMinimum))
			/ (mc.Bounds.dwMaximum - mc.Bounds.dwMinimum)   );
}


HRESULT SoundSetGet_GetDevice(LPTSTR aDeviceString, IMMDevice *&aDevice)
{
	IMMDeviceEnumerator *deviceEnum;
	IMMDeviceCollection *devices;
	HRESULT hr;

	aDevice = NULL;

	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&deviceEnum);
	if (SUCCEEDED(hr))
	{
		if (*aDeviceString)
		{
			int device_index = ATOI(aDeviceString) - 1;
			if (device_index < 0)
				device_index = 0; // For consistency with 2k/XP.

			// Enumerate devices; include unplugged devices so that indices don't change when a device is plugged in.
			hr = deviceEnum->EnumAudioEndpoints(eAll, DEVICE_STATE_ACTIVE | DEVICE_STATE_UNPLUGGED, &devices);
			if (SUCCEEDED(hr))
			{
				hr = devices->Item((UINT)device_index, &aDevice);
				devices->Release();
			}
		}
		else
		{
			// Get default playback device.
			hr = deviceEnum->GetDefaultAudioEndpoint(eRender, eConsole, &aDevice);
		}
		deviceEnum->Release();
	}
	return hr;
}


DWORD SoundSetGet_ComponentType(GUID &aKSNodeType)
{

	struct Pair
	{
		const GUID *ksnodetype; // KSNODETYPE_*
		DWORD mixertype; // MIXERLINE_COMPONENTTYPE_SRC_*
	};

	static Pair TypeMap[] =
	{
		// Reference: http://msdn.microsoft.com/en-us/library/windows/hardware/ff538578
		&KSNODETYPE_MICROPHONE,				MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE,
		&KSNODETYPE_DESKTOP_MICROPHONE,		MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE,
		&KSNODETYPE_LEGACY_AUDIO_CONNECTOR,	MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT,
		&KSCATEGORY_AUDIO,					MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT,
		&KSNODETYPE_SPEAKER,				MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT,
		&KSNODETYPE_CD_PLAYER,				MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC,
		&KSNODETYPE_SYNTHESIZER,			MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER,
		&KSNODETYPE_LINE_CONNECTOR,			MIXERLINE_COMPONENTTYPE_SRC_LINE,
		&KSNODETYPE_TELEPHONE,				MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE,
		&KSNODETYPE_PHONE_LINE,				MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE,
		&KSNODETYPE_DOWN_LINE_PHONE,		MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE,
		&KSNODETYPE_ANALOG_CONNECTOR,		MIXERLINE_COMPONENTTYPE_SRC_ANALOG,
		&KSNODETYPE_SPDIF_INTERFACE,		MIXERLINE_COMPONENTTYPE_SRC_DIGITAL
	};

	for (int i = 0; i < _countof(TypeMap); ++i)
		if (aKSNodeType == *TypeMap[i].ksnodetype)
			return TypeMap[i].mixertype;
	
	return MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED;
}


struct SoundComponentSearch
{
	// Parameters set by caller:
	DWORD target_type;
	int target_instance;
	LPCGUID target_iid;
	// Internal use/results:
	IUnknown *control;
	int count;
	// Internal use:
	DataFlow data_flow;
};


bool SoundSetGet_FindComponent(IPart *aRoot, SoundComponentSearch &aSearch)
{
	HRESULT hr;
	IPartsList *parts;
	IPart *part;
	UINT part_count;
	PartType part_type;
	GUID sub_type;

	if (aSearch.data_flow == In)
		hr = aRoot->EnumPartsIncoming(&parts);
	else
		hr = aRoot->EnumPartsOutgoing(&parts);
	
	if (FAILED(hr))
		return NULL;

	if (FAILED(parts->GetCount(&part_count)))
		part_count = 0;
	
	for (UINT i = 0; i < part_count; ++i)
	{
		if (FAILED(parts->GetPart(i, &part)))
			continue;

		if (SUCCEEDED(part->GetPartType(&part_type)))
		{
			if (part_type == Connector)
			{
				if (SUCCEEDED(part->GetSubType(&sub_type)))
				{
					if (SoundSetGet_ComponentType(sub_type) == aSearch.target_type)
					{
						if (++aSearch.count == aSearch.target_instance)
						{
							part->Release();
							parts->Release();
							return true;
						}
					}
				}
			}
			else // Subunit
			{
				// Recursively find the Connector nodes linked to this part.
				if (SoundSetGet_FindComponent(part, aSearch))
				{
					// A matching connector part has been found with this part as one of the nodes used
					// to reach it.  Therefore, if this part supports the requested control interface,
					// it can in theory be used to control the component.  An example path might be:
					//    Output < Master Mute < Master Volume < Sum < Mute < Volume < CD Audio
					// Parts are considered from right to left, as we return from recursion.
					if (aSearch.target_iid && !aSearch.control)
					{
						// Query this part for the requested interface and let caller check the result.
						part->Activate(CLSCTX_ALL, *aSearch.target_iid, (void **)&aSearch.control);
						
						// If this subunit has siblings, ignore any controls further up the line
						// as they're likely shared by other components (i.e. master controls).
						if (part_count > 1)
							aSearch.target_iid = NULL;
					}
					part->Release();
					parts->Release();
					return true;
				}
			}
		}

		part->Release();
	}

	parts->Release();
	return false;
}


bool SoundSetGet_FindComponent(IMMDevice *aDevice, SoundComponentSearch &aSearch)
{
	IDeviceTopology *topo;
	IConnector *conn, *conn_to;
	IPart *part;

	aSearch.count = 0;
	aSearch.control = NULL;
	
	if (SUCCEEDED(aDevice->Activate(__uuidof(IDeviceTopology), CLSCTX_ALL, NULL, (void**)&topo)))
	{
		if (SUCCEEDED(topo->GetConnector(0, &conn)))
		{
			if (SUCCEEDED(conn->GetDataFlow(&aSearch.data_flow))
			 && SUCCEEDED(conn->GetConnectedTo(&conn_to)))
			{
				if (SUCCEEDED(conn_to->QueryInterface(&part)))
				{
					// Search; the result is stored in the search struct.
					SoundSetGet_FindComponent(part, aSearch);
					part->Release();
				}
				conn_to->Release();
			}
			conn->Release();
		}
		topo->Release();
	}

	return aSearch.count == aSearch.target_instance;
}


ResultType Line::SoundSetGetVista(LPTSTR aSetting, DWORD aComponentType, int aComponentInstance
		, DWORD aControlType, LPTSTR aDeviceString)
{
	float setting_scalar;
	Var *output_var;
	if (SOUND_MODE_IS_SET)
	{
		output_var = NULL; // To help catch bugs.
		setting_scalar = (float)(ATOF(aSetting) / 100);
		if (setting_scalar < -1)
			setting_scalar = -1;
		else if (setting_scalar > 1)
			setting_scalar = 1;
	}
	else // The mode is GET.
	{
		output_var = OUTPUT_VAR;
	}

	// Does user want to adjust the current setting by a certain amount?
	bool adjust_current_setting = aSetting && (*aSetting == '-' || *aSetting == '+' || *RAW_ARG1 == '+');

	IMMDevice *device;
	HRESULT hr;

	hr = SoundSetGet_GetDevice(aDeviceString, device);
	if (FAILED(hr))
		return SetErrorLevelOrThrowStr(_T("Can't Open Specified Mixer"));

	LPCTSTR errorlevel = NULL;
	float result_float;
	BOOL result_bool;
	bool control_type_is_boolean;

	if (aComponentType == MIXERLINE_COMPONENTTYPE_DST_SPEAKERS) // ComponentType is Master/Speakers/omitted.
	{
		if (aComponentInstance != 1)
		{
			errorlevel = _T("Mixer Doesn't Have That Many of That Component Type");
		}
		else if (aControlType == MIXERCONTROL_CONTROLTYPE_MUTE
			|| aControlType == MIXERCONTROL_CONTROLTYPE_VOLUME)
		{
			// For Master/Speakers, use the IAudioEndpointVolume interface.  Some devices support master
			// volume control, but do not actually have a volume subunit (so the other method would fail).
			IAudioEndpointVolume *aev;
			hr = device->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, NULL, (void**)&aev);
			if (SUCCEEDED(hr))
			{
				if (aControlType == MIXERCONTROL_CONTROLTYPE_VOLUME)
				{
					if (!SOUND_MODE_IS_SET || adjust_current_setting)
					{
						hr = aev->GetMasterVolumeLevelScalar(&result_float);
					}

					if (SUCCEEDED(hr))
					{
						if (SOUND_MODE_IS_SET)
						{
							if (adjust_current_setting)
							{
								setting_scalar += result_float;
								if (setting_scalar > 1)
									setting_scalar = 1;
								else if (setting_scalar < 0)
									setting_scalar = 0;
							}
							hr = aev->SetMasterVolumeLevelScalar(setting_scalar, NULL);
						}
						else
						{
							result_float *= 100.0;
						} 
					}
					control_type_is_boolean = false;
				}
				else // Mute.
				{
					if (!SOUND_MODE_IS_SET || adjust_current_setting)
					{
						hr = aev->GetMute(&result_bool);
					}
					if (SOUND_MODE_IS_SET && SUCCEEDED(hr))
					{
						hr = aev->SetMute(adjust_current_setting ? !result_bool : setting_scalar > 0, NULL);
					}
					control_type_is_boolean = true;
				}
				aev->Release();
			}
		}
		else
			errorlevel = _T("Component Doesn't Support This Control Type");
	}
	else
	{
		SoundComponentSearch search;
		search.target_type = aComponentType;
		search.target_instance = aComponentInstance;
		
		switch (aControlType)
		{
		case MIXERCONTROL_CONTROLTYPE_VOLUME:		search.target_iid = &__uuidof(IAudioVolumeLevel); break;
		case MIXERCONTROL_CONTROLTYPE_MUTE:			search.target_iid = &__uuidof(IAudioMute); break;
		// Since specific code would need to be written for each control type, the types below are left
		// unimplemented.  Support for these controls is up to the audio drivers, and is completely absent
		// from the drivers for my Realtek HD onboard audio, Logitech G330 headset and ASUS Xonar DS card.
		//case MIXERCONTROL_CONTROLTYPE_ONOFF:
		//case MIXERCONTROL_CONTROLTYPE_MONO:
		//case MIXERCONTROL_CONTROLTYPE_LOUDNESS:
		//case MIXERCONTROL_CONTROLTYPE_STEREOENH:
		//case MIXERCONTROL_CONTROLTYPE_BASS_BOOST:
		//case MIXERCONTROL_CONTROLTYPE_PAN:
		//case MIXERCONTROL_CONTROLTYPE_QSOUNDPAN:
		//case MIXERCONTROL_CONTROLTYPE_BASS:
		//case MIXERCONTROL_CONTROLTYPE_TREBLE:
		//case MIXERCONTROL_CONTROLTYPE_EQUALIZER:
		default: search.target_iid = NULL;
		}
		if (!SoundSetGet_FindComponent(device, search))
		{
			if (search.count)
				errorlevel = _T("Mixer Doesn't Have That Many of That Component Type");
			else
				errorlevel = _T("Mixer Doesn't Support This Component Type");
		}
		else if (!search.control)
		{
			errorlevel = _T("Component Doesn't Support This Control Type");
		}
		else if (aControlType == MIXERCONTROL_CONTROLTYPE_VOLUME)
		{
			IAudioVolumeLevel *avl = (IAudioVolumeLevel *)search.control;
				
			UINT channel_count = 0;
			if (FAILED(avl->GetChannelCount(&channel_count)))
				goto control_fail;
				
			float *level = (float *)_alloca(sizeof(float) * 3 * channel_count);
			float *level_min = level + channel_count;
			float *level_range = level_min + channel_count;
			float f, db, min_db, max_db, max_level = 0;

			for (UINT i = 0; i < channel_count; ++i)
			{
				if (FAILED(avl->GetLevel(i, &db)) ||
					FAILED(avl->GetLevelRange(i, &min_db, &max_db, &f)))
					goto control_fail;
				// Convert dB to scalar.
				level_min[i] = (float)qmathExp10(min_db/20);
				level_range[i] = (float)qmathExp10(max_db/20) - level_min[i];
				// Compensate for differing level ranges. (No effect if range is -96..0 dB.)
				level[i] = ((float)qmathExp10(db/20) - level_min[i]) / level_range[i];
				// Windows reports the highest level as the overall volume.
				if (max_level < level[i])
					max_level = level[i];
			}

			if (SOUND_MODE_IS_SET)
			{
				if (adjust_current_setting)
				{
					setting_scalar += max_level;
					if (setting_scalar > 1)
						setting_scalar = 1;
					else if (setting_scalar < 0)
						setting_scalar = 0;
				}

				for (UINT i = 0; i < channel_count; ++i)
				{
					f = setting_scalar;
					if (max_level)
						f *= (level[i] / max_level); // Preserve balance.
					// Compensate for differing level ranges.
					f = level_min[i] + f * level_range[i];
					// Convert scalar to dB.
					level[i] = 20 * (float)qmathLog10(f);
				}

				hr = avl->SetLevelAllChannels(level, channel_count, NULL);
			}
			else
			{
				result_float = max_level * 100;
				control_type_is_boolean = false;
			}
		}
		else if (aControlType == MIXERCONTROL_CONTROLTYPE_MUTE)
		{
			IAudioMute *am = (IAudioMute *)search.control;

			if (!SOUND_MODE_IS_SET || adjust_current_setting)
			{
				hr = am->GetMute(&result_bool);
			}
			if (SOUND_MODE_IS_SET && SUCCEEDED(hr))
			{
				hr = am->SetMute(adjust_current_setting ? !result_bool : setting_scalar > 0, NULL);
			}
			control_type_is_boolean = true;
		}
control_fail:
		if (search.control)
			search.control->Release();
	}

	device->Release();
	
	if (FAILED(hr))
	{
		if (SOUND_MODE_IS_SET)
			errorlevel = _T("Can't Change Setting");
		else
			errorlevel = _T("Can't Get Current Setting");
	}
	if (errorlevel)
		return SetErrorLevelOrThrowStr(errorlevel);
	else
		g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	if (SOUND_MODE_IS_SET)
		return OK;

	if (control_type_is_boolean)
		return output_var->Assign(result_bool ? _T("On") : _T("Off"));
	else
		return output_var->Assign(result_float);
}



ResultType Line::SoundGetWaveVolume(HWAVEOUT aDeviceID)
{
	OUTPUT_VAR->Assign(); // Init to empty string regardless of whether we succeed here.

	if (g_os.IsWinVistaOrLater())
	{
		// For consistency and convenience when running scripts designed for XP,
		// make the following two lines equivalent:
		//  SoundGetWaveVolume v
		//  SoundGet v, Wave, Volume
		return SoundSetGetVista(NULL, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, 1, MIXERCONTROL_CONTROLTYPE_VOLUME, ARG2);
	}

	DWORD current_vol;
	if (waveOutGetVolume(aDeviceID, &current_vol) != MMSYSERR_NOERROR)
		return SetErrorLevelOrThrow();

	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.

	// Return only the left channel volume level (in case right is different, or device is mono vs. stereo).
	// MSDN: "If a device does not support both left and right volume control, the low-order word
	// of the specified location contains the mono volume level.
	return OUTPUT_VAR->Assign((double)(LOWORD(current_vol) * 100) / 0xFFFF);
}



ResultType Line::SoundSetWaveVolume(LPTSTR aVolume, HWAVEOUT aDeviceID)
{
	if (g_os.IsWinVistaOrLater())
	{
		// For consistency and convenience when running scripts designed for XP,
		// make the following two lines equivalent:
		//  SoundSetWaveVolume %v%
		//  SoundSet %v%, Wave, Volume
		return SoundSetGetVista(aVolume, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, 1, MIXERCONTROL_CONTROLTYPE_VOLUME, ARG2);
	}

	double volume = ATOF(aVolume);
	if (volume < -100)
		volume = -100;
	else if (volume > 100)
		volume = 100;

	// Make this an int vs. WORD so that negative values are supported (e.g. adjust volume by -10%).
	int specified_vol_per_channel = (int)(0xFFFF * (volume / 100));
	DWORD vol_new;

	// For v1.0.25, the first char of RAW_ARG is also checked in case this is an expression intended
	// to be a positive offset, such as +(var + 10)
	if (*aVolume == '-' || *aVolume == '+' || *RAW_ARG1 == '+') // User wants to adjust the current level by a certain amount.
	{
		DWORD current_vol;
		if (waveOutGetVolume(aDeviceID, &current_vol) != MMSYSERR_NOERROR)
			return SetErrorLevelOrThrow();
		// Adjust left & right independently so that we at least attempt to retain the user's
		// balance setting (if overflow or underflow occurs, the relative balance might be lost):
		int vol_left = LOWORD(current_vol); // Make them ints so that overflow/underflow can be detected.
		int vol_right = HIWORD(current_vol);
		vol_left += specified_vol_per_channel;
		vol_right += specified_vol_per_channel;
		// Handle underflow or overflow:
		if (vol_left < 0)
			vol_left = 0;
		else if (vol_left > 0xFFFF)
			vol_left = 0xFFFF;
		if (vol_right < 0)
			vol_right = 0;
		else if (vol_right > 0xFFFF)
			vol_right = 0xFFFF;
		vol_new = MAKELONG((WORD)vol_left, (WORD)vol_right);  // Left is low-order, right is high-order.
	}
	else // User wants the volume level set to an absolute level (i.e. ignore its current level).
		vol_new = MAKELONG((WORD)specified_vol_per_channel, (WORD)specified_vol_per_channel);

	if (waveOutSetVolume(aDeviceID, vol_new) == MMSYSERR_NOERROR)
		return g_ErrorLevel->Assign(ERRORLEVEL_NONE);
	else
		return SetErrorLevelOrThrow();
}



ResultType Line::SoundPlay(LPTSTR aFilespec, bool aSleepUntilDone)
{
	LPTSTR cp = omit_leading_whitespace(aFilespec);
	if (*cp == '*')
		return SetErrorLevelOrThrowBool(!MessageBeep(ATOU(cp + 1)));
		// ATOU() returns 0xFFFFFFFF for -1, which is relied upon to support the -1 sound.
	// See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/multimed/htm/_win32_play.asp
	// for some documentation mciSendString() and related.
	TCHAR buf[MAX_PATH * 2]; // Allow room for filename and commands.
	mciSendString(_T("status ") SOUNDPLAY_ALIAS _T(" mode"), buf, _countof(buf), NULL);
	if (*buf) // "playing" or "stopped" (so close it before trying to re-open with a new aFilespec).
		mciSendString(_T("close ") SOUNDPLAY_ALIAS, NULL, 0, NULL);
	sntprintf(buf, _countof(buf), _T("open \"%s\" alias ") SOUNDPLAY_ALIAS, aFilespec);
	if (mciSendString(buf, NULL, 0, NULL)) // Failure.
		return SetErrorLevelOrThrow();
	g_SoundWasPlayed = true;  // For use by Script's destructor.
	if (mciSendString(_T("play ") SOUNDPLAY_ALIAS, NULL, 0, NULL)) // Failure.
		return SetErrorLevelOrThrow();
	// Otherwise, the sound is now playing.
	g_ErrorLevel->Assign(ERRORLEVEL_NONE);
	if (!aSleepUntilDone)
		return OK;
	// Otherwise, caller wants us to wait until the file is done playing.  To allow our app to remain
	// responsive during this time, use a loop that checks our message queue:
	// Older method: "mciSendString("play " SOUNDPLAY_ALIAS " wait", NULL, 0, NULL)"
	for (;;)
	{
		mciSendString(_T("status ") SOUNDPLAY_ALIAS _T(" mode"), buf, _countof(buf), NULL);
		if (!*buf) // Probably can't happen given the state we're in.
			break;
		if (!_tcscmp(buf, _T("stopped"))) // The sound is done playing.
		{
			mciSendString(_T("close ") SOUNDPLAY_ALIAS, NULL, 0, NULL);
			break;
		}
		// Sleep a little longer than normal because I'm not sure how much overhead
		// and CPU utilization the above incurs:
		MsgSleep(20);
	}
	return OK;
}



void SetWorkingDir(LPTSTR aNewDir)
// Sets ErrorLevel to indicate success/failure, but only if the script has begun runtime execution (callers
// want that).
// This function was added in v1.0.45.01 for the reasons commented further below.
// It's similar to the one in the ahk2exe source, so maintain them together.
{
	if (!SetCurrentDirectory(aNewDir)) // Caused by nonexistent directory, permission denied, etc.
	{
		if (g_script.mIsReadyToExecute)
			g_script.SetErrorLevelOrThrow();
		return;
	}

	// Otherwise, the change to the working directory *apparently* succeeded (but is confirmed below for root drives
	// and also because we want the absolute path in cases where aNewDir is relative).
	TCHAR buf[_countof(g_WorkingDir)];
	LPTSTR actual_working_dir = g_script.mIsReadyToExecute ? g_WorkingDir : buf; // i.e. don't update g_WorkingDir when our caller is the #include directive.
	// Other than during program startup, this should be the only place where the official
	// working dir can change.  The exception is FileSelectFile(), which changes the working
	// dir as the user navigates from folder to folder.  However, the whole purpose of
	// maintaining g_WorkingDir is to workaround that very issue.

	// GetCurrentDirectory() is called explicitly, to confirm the change, in case aNewDir is a relative path.
	// We want to store the absolute path:
	if (!GetCurrentDirectory(_countof(buf), actual_working_dir)) // Might never fail in this case, but kept for backward compatibility.
	{
		tcslcpy(actual_working_dir, aNewDir, _countof(buf)); // Update the global to the best info available.
		// But ErrorLevel is set to "none" further below because the actual "set" did succeed; it's also for
		// backward compatibility.
	}
	else // GetCurrentDirectory() succeeded, so it's appropriate to compare what we asked for to what was received.
	{
		if (aNewDir[0] && aNewDir[1] == ':' && !aNewDir[2] // Root with missing backslash. Relies on short-circuit boolean order.
			&& _tcsicmp(aNewDir, actual_working_dir)) // The root directory we requested didn't actually get set. See below.
		{
			// There is some strange OS behavior here: If the current working directory is C:\anything\...
			// and SetCurrentDirectory() is called to switch to "C:", the function reports success but doesn't
			// actually change the directory.  However, if you first change to D: then back to C:, the change
			// works as expected.  Presumably this is for backward compatibility with DOS days; but it's 
			// inconvenient and seems desirable to override it in this case, especially because:
			// v1.0.45.01: Since A_ScriptDir omits the trailing backslash for roots of drives (such as C:),
			// and since that variable probably shouldn't be changed for backward compatibility, provide
			// the missing backslash to allow SetWorkingDir %A_ScriptDir% (and others) to work in the root
			// of a drive.
			TCHAR buf_temp[8];
			_stprintf(buf_temp, _T("%s\\"), aNewDir); // No danger of buffer overflow in this case.
			if (SetCurrentDirectory(buf_temp))
			{
				if (!GetCurrentDirectory(_countof(buf), actual_working_dir)) // Might never fail in this case, but kept for backward compatibility.
					tcslcpy(actual_working_dir, aNewDir, _countof(buf)); // But still report "no error" (below) because the Set() actually did succeed.
					// But treat this as a success like the similar one higher above.
			}
			//else Set() failed; but since the original Set() succeeded (and for simplicity) report ErrorLevel "none".
		}
	}

	// Since the above didn't return, it wants us to indicate success.
	if (g_script.mIsReadyToExecute) // Callers want ErrorLevel changed only during script runtime.
		g_ErrorLevel->Assign(ERRORLEVEL_NONE);
}



ResultType Line::FileSelectFile(LPTSTR aOptions, LPTSTR aWorkingDir, LPTSTR aGreeting, LPTSTR aFilter)
// Since other script threads can interrupt this command while it's running, it's important that
// this command not refer to sArgDeref[] and sArgVar[] anytime after an interruption becomes possible.
// This is because an interrupting thread usually changes the values to something inappropriate for this thread.
{
	Var &output_var = *OUTPUT_VAR; // Fix for v1.0.45.01: Must be resolved and saved early.  See comment above.
	if (g_nFileDialogs >= MAX_FILEDIALOGS)
	{
		// Have a maximum to help prevent runaway hotkeys due to key-repeat feature, etc.
		return LineError(_T("The maximum number of File Dialogs has been reached."));
	}
	
	// Large in case more than one file is allowed to be selected.
	// The call to GetOpenFileName() may fail if the first character of the buffer isn't NULL
	// because it then thinks the buffer contains the default filename, which if it's uninitialized
	// may be a string that's too long.
	TCHAR file_buf[65535];
	*file_buf = '\0'; // Set default.

	TCHAR working_dir[MAX_PATH];
	if (!aWorkingDir || !*aWorkingDir)
		*working_dir = '\0';
	else
	{
		tcslcpy(working_dir, aWorkingDir, _countof(working_dir));
		// v1.0.43.10: Support CLSIDs such as:
		//   My Computer  ::{20d04fe0-3aea-1069-a2d8-08002b30309d}
		//   My Documents ::{450d8fba-ad25-11d0-98a8-0800361b1103}
		// Also support optional subdirectory appended to the CLSID.
		// Neither SetCurrentDirectory() nor GetFileAttributes() directly supports CLSIDs, so rely on other means
		// to detect whether a CLSID ends in a directory vs. filename.
		bool is_directory, is_clsid;
		if (is_clsid = !_tcsncmp(working_dir, _T("::{"), 3))
		{
			LPTSTR end_brace;
			if (end_brace = _tcschr(working_dir, '}'))
				is_directory = !end_brace[1] // First '}' is also the last char in string, so it's naked CLSID (so assume directory).
					|| working_dir[_tcslen(working_dir) - 1] == '\\'; // Or path ends in backslash.
			else // Badly formatted clsid.
				is_directory = true; // Arbitrary default due to rarity.
		}
		else // Not a CLSID.
		{
			DWORD attr = GetFileAttributes(working_dir);
			is_directory = (attr != 0xFFFFFFFF) && (attr & FILE_ATTRIBUTE_DIRECTORY);
		}
		if (!is_directory)
		{
			// Above condition indicates it's either an existing file that's not a folder, or a nonexistent
			// folder/filename.  In either case, it seems best to assume it's a file because the user may want
			// to provide a default SAVE filename, and it would be normal for such a file not to already exist.
			LPTSTR last_backslash;
			if (last_backslash = _tcsrchr(working_dir, '\\'))
			{
				tcslcpy(file_buf, last_backslash + 1, _countof(file_buf)); // Set the default filename.
				*last_backslash = '\0'; // Make the working directory just the file's path.
			}
			else // The entire working_dir string is the default file (unless this is a clsid).
				if (!is_clsid)
				{
					tcslcpy(file_buf, working_dir, _countof(file_buf));
					*working_dir = '\0';  // This signals it to use the default directory.
				}
				//else leave working_dir set to the entire clsid string in case it's somehow valid.
		}
		// else it is a directory, so just leave working_dir set as it was initially.
	}

	TCHAR greeting[1024];
	if (aGreeting && *aGreeting)
		tcslcpy(greeting, aGreeting, _countof(greeting));
	else
		// Use a more specific title so that the dialogs of different scripts can be distinguished
		// from one another, which may help script automation in rare cases:
		sntprintf(greeting, _countof(greeting), _T("Select File - %s"), g_script.mFileName);

	// The filter must be terminated by two NULL characters.  One is explicit, the other automatic:
	TCHAR filter[1024], pattern[1024];
	*filter = '\0'; *pattern = '\0'; // Set default.
	if (*aFilter)
	{
		LPTSTR pattern_start = _tcschr(aFilter, '(');
		if (pattern_start)
		{
			// Make pattern a separate string because we want to remove any spaces from it.
			// For example, if the user specified Documents (*.txt; *.doc), the space after
			// the semicolon should be removed for the pattern string itself but not from
			// the displayed version of the pattern:
			tcslcpy(pattern, ++pattern_start, _countof(pattern));
			LPTSTR pattern_end = _tcsrchr(pattern, ')'); // strrchr() in case there are other literal parentheses.
			if (pattern_end)
				*pattern_end = '\0';  // If parentheses are empty, this will set pattern to be the empty string.
			else // no closing paren, so set to empty string as an indicator:
				*pattern = '\0';

		}
		else // No open-paren, so assume the entire string is the filter.
			tcslcpy(pattern, aFilter, _countof(pattern));
		if (*pattern)
		{
			// Remove any spaces present in the pattern, such as a space after every semicolon
			// that separates the allowed file extensions.  The API docs specify that there
			// should be no spaces in the pattern itself, even though it's okay if they exist
			// in the displayed name of the file-type:
			StrReplace(pattern, _T(" "), _T(""), SCS_SENSITIVE);
			// Also include the All Files (*.*) filter, since there doesn't seem to be much
			// point to making this an option.  This is because the user could always type
			// *.* and press ENTER in the filename field and achieve the same result:
			sntprintf(filter, _countof(filter), _T("%s%c%s%cAll Files (*.*)%c*.*%c")
				, aFilter, '\0', pattern, '\0', '\0', '\0'); // The final '\0' double-terminates by virtue of the fact that sntprintf() itself provides a final terminator.
		}
		else
			*filter = '\0';  // It will use a standard default below.
	}

	OPENFILENAME ofn = {0};
	// OPENFILENAME_SIZE_VERSION_400 must be used for 9x/NT otherwise the dialog will not appear!
	// MSDN: "In an application that is compiled with WINVER and _WIN32_WINNT >= 0x0500, use
	// OPENFILENAME_SIZE_VERSION_400 for this member.  Windows 2000/XP: Use sizeof(OPENFILENAME)
	// for this parameter."
	ofn.lStructSize = g_os.IsWin2000orLater() ? sizeof(OPENFILENAME) : OPENFILENAME_SIZE_VERSION_400;
	ofn.hwndOwner = THREAD_DIALOG_OWNER; // Can be NULL, which is used instead of main window since no need to have main window forced into the background for this.
	ofn.lpstrTitle = greeting;
	ofn.lpstrFilter = *filter ? filter : _T("All Files (*.*)\0*.*\0Text Documents (*.txt)\0*.txt\0");
	ofn.lpstrFile = file_buf;
	ofn.nMaxFile = _countof(file_buf) - 1; // -1 to be extra safe.
	// Specifying NULL will make it default to the last used directory (at least in Win2k):
	ofn.lpstrInitialDir = *working_dir ? working_dir : NULL;

	// Note that the OFN_NOCHANGEDIR flag is ineffective in some cases, so we'll use a custom
	// workaround instead.  MSDN: "Windows NT 4.0/2000/XP: This flag is ineffective for GetOpenFileName."
	// In addition, it does not prevent the CWD from changing while the user navigates from folder to
	// folder in the dialog, except perhaps on Win9x.

	// For v1.0.25.05, the new "M" letter is used for a new multi-select method since the old multi-select
	// is faulty in the following ways:
	// 1) If the user selects a single file in a multi-select dialog, the result is inconsistent: it
	//    contains the full path and name of that single file rather than the folder followed by the
	//    single file name as most users would expect.  To make matters worse, it includes a linefeed
	//    after that full path in name, which makes it difficult for a script to determine whether
	//    only a single file was selected.
	// 2) The last item in the list is terminated by a linefeed, which is not as easily used with a
	//    parsing loop as shown in example in the help file.
	bool always_use_save_dialog = false; // Set default.
	bool new_multi_select_method = false; // Set default.
	switch (ctoupper(*aOptions))
	{
	case 'M':  // Multi-select.
		++aOptions;
		new_multi_select_method = true;
		break;
	case 'S': // Have a "Save" button rather than an "Open" button.
		++aOptions;
		always_use_save_dialog = true;
		break;
	}

	int options = ATOI(aOptions);
	// v1.0.43.09: OFN_NODEREFERENCELINKS is now omitted by default because most people probably want a click
	// on a shortcut to navigate to the shortcut's target rather than select the shortcut and end the dialog.
	ofn.Flags = OFN_HIDEREADONLY | OFN_EXPLORER;  // OFN_HIDEREADONLY: Hides the Read Only check box.
	if (options & 0x20) // v1.0.43.09.
		ofn.Flags |= OFN_NODEREFERENCELINKS;
	if (options & 0x10)
		ofn.Flags |= OFN_OVERWRITEPROMPT;
	if (options & 0x08)
		ofn.Flags |= OFN_CREATEPROMPT;
	if (new_multi_select_method || (options & 0x04))
		ofn.Flags |= OFN_ALLOWMULTISELECT;
	if (options & 0x02)
		ofn.Flags |= OFN_PATHMUSTEXIST;
	if (options & 0x01)
		ofn.Flags |= OFN_FILEMUSTEXIST;

	// At this point, we know a dialog will be displayed.  See macro's comments for details:
	DIALOG_PREP
	POST_AHK_DIALOG(0) // Do this only after the above. Must pass 0 for timeout in this case.

	++g_nFileDialogs;
	// Below: OFN_CREATEPROMPT doesn't seem to work with GetSaveFileName(), so always
	// use GetOpenFileName() in that case:
	BOOL result = (always_use_save_dialog || ((ofn.Flags & OFN_OVERWRITEPROMPT) && !(ofn.Flags & OFN_CREATEPROMPT)))
		? GetSaveFileName(&ofn) : GetOpenFileName(&ofn);
	--g_nFileDialogs;

	DIALOG_END

	// Both GetOpenFileName() and GetSaveFileName() change the working directory as a side-effect
	// of their operation.  The below is not a 100% workaround for the problem because even while
	// a new quasi-thread is running (having interrupted this one while the dialog is still
	// displayed), the dialog is still functional, and as a result, the dialog changes the
	// working directory every time the user navigates to a new folder.
	// This is only needed when the user pressed OK, since the dialog auto-restores the
	// working directory if CANCEL is pressed or the window was closed by means other than OK.
	// UPDATE: No, it's needed for CANCEL too because GetSaveFileName/GetOpenFileName will restore
	// the working dir to the wrong dir if the user changed it (via SetWorkingDir) while the
	// dialog was displayed.
	// Restore the original working directory so that any threads suspended beneath this one,
	// and any newly launched ones if there aren't any suspended threads, will have the directory
	// that the user expects.  NOTE: It's possible for g_WorkingDir to have changed via the
	// SetWorkingDir command while the dialog was displayed (e.g. a newly launched quasi-thread):
	if (*g_WorkingDir)
		SetCurrentDirectory(g_WorkingDir);

	if (!result) // User pressed CANCEL vs. OK to dismiss the dialog or there was a problem displaying it.
	{
		// It seems best to clear the variable in these cases, since this is a scripting
		// language where performance is not the primary goal.  So do that and return OK,
		// but leave ErrorLevel set to ERRORLEVEL_ERROR.
		if (output_var.Assign() != OK)
			return FAIL;
		return CommDlgExtendedError() ? SetErrorLevelOrThrow() // An error occurred.
			: g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // User pressed CANCEL, so never throw an exception.
	}
	else
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate that the user pressed OK vs. CANCEL.

	if (ofn.Flags & OFN_ALLOWMULTISELECT)
	{
		LPTSTR cp;
		if (new_multi_select_method) // v1.0.25.05+ method.
		{
			// If the first terminator in file_buf is also the last, the user selected only
			// a single file:
			size_t length = _tcslen(file_buf);
			if (!file_buf[length + 1]) // The list contains only a single file (full path and name).
			{
				// v1.0.25.05: To make the result of selecting one file the same as selecting multiple files
				// -- and thus easier to work with in a script -- convert the result into the multi-file
				// format (folder as first item and naked filename as second):
				if (cp = _tcsrchr(file_buf, '\\'))
				{
					*cp = '\n';
					// If the folder is the root folder, add a backslash so that selecting a single
					// file yields the same reported folder as selecting multiple files.  One reason
					// for doing it this way is that SetCurrentDirectory() requires a backslash after
					// a root folder to succeed.  This allows a script to use SetWorkingDir to change
					// to the selected folder before operating on each of the selected/naked filenames.
					if (cp - file_buf == 2 && cp[-1] == ':') // e.g. "C:"
					{
						tmemmove(cp + 1, cp, _tcslen(cp) + 1); // Make room to insert backslash (since only one file was selected, the buf is large enough).
						*cp = '\\';
					}
				}
			}
			else // More than one file was selected.
			{
				// Use the same method as the old multi-select format except don't provide a
				// linefeed after the final item.  That final linefeed would make parsing via
				// a parsing loop more complex because a parsing loop would see a blank item
				// at the end of the list:
				for (cp = file_buf;;)
				{
					for (; *cp; ++cp); // Find the next terminator.
					if (!cp[1]) // This is the last file because it's double-terminated, so we're done.
						break;
					*cp = '\n'; // Replace zero-delimiter with a visible/printable delimiter, for the user.
				}
			}
		}
		else  // Old multi-select method is in effect (kept for backward compatibility).
		{
			// Replace all the zero terminators with a delimiter, except the one for the last file
			// (the last file should be followed by two sequential zero terminators).
			// Use a delimiter that can't be confused with a real character inside a filename, i.e.
			// not a comma.  We only have room for one without getting into the complexity of having
			// to expand the string, so \r\n is disqualified for now.
			for (cp = file_buf;;)
			{
				for (; *cp; ++cp); // Find the next terminator.
				*cp = '\n'; // Replace zero-delimiter with a visible/printable delimiter, for the user.
				if (!cp[1]) // This is the last file because it's double-terminated, so we're done.
					break;
			}
		}
	}
	return output_var.Assign(file_buf);
}



ResultType Line::FileCreateDir(LPTSTR aDirSpec)
{
	if (!aDirSpec || !*aDirSpec)
		return SetErrorsOrThrow(true, ERROR_INVALID_PARAMETER);

	DWORD attr = GetFileAttributes(aDirSpec);
	if (attr != 0xFFFFFFFF)  // aDirSpec already exists.
		return SetErrorsOrThrow(!(attr & FILE_ATTRIBUTE_DIRECTORY), ERROR_ALREADY_EXISTS); // Indicate success if it already exists as a dir.

	// If it has a backslash, make sure all its parent directories exist before we attempt
	// to create this directory:
	LPTSTR last_backslash = _tcsrchr(aDirSpec, '\\');
	if (last_backslash > aDirSpec) // v1.0.48.04: Changed "last_backslash" to "last_backslash > aDirSpec" so that an aDirSpec with a leading \ (but no other backslashes), such as \dir, is supported.
	{
		TCHAR parent_dir[MAX_PATH];
		if (_tcslen(aDirSpec) >= _countof(parent_dir)) // avoid overflow
			return SetErrorsOrThrow(true, ERROR_BUFFER_OVERFLOW);
		tcslcpy(parent_dir, aDirSpec, last_backslash - aDirSpec + 1); // Omits the last backslash.
		FileCreateDir(parent_dir); // Recursively create all needed ancestor directories.

		// v1.0.44: Fixed ErrorLevel being set to 1 when the specified directory ends in a backslash.  In such cases,
		// two calls were made to CreateDirectory for the same folder: the first without the backslash and then with
		// it.  Since the directory already existed on the second call, ErrorLevel was wrongly set to 1 even though
		// everything succeeded.  So now, when recursion finishes creating all the ancestors of this directory
		// our own layer here does not call CreateDirectory() when there's a trailing backslash because a previous
		// layer already did:
		if (!last_backslash[1] || *g_ErrorLevel->Contents() == *ERRORLEVEL_ERROR) // Compare first char of each string, which is valid because ErrorLevel is stored as a quoted/literal string rather than an integer.
			return OK; // Let the previously set ErrorLevel (whatever it is) tell the story.
	}

	// The above has recursively created all parent directories of aDirSpec if needed.
	// Now we can create aDirSpec.  Be sure to explicitly set g_ErrorLevel since it's value
	// is now indeterminate due to action above:
	return SetErrorsOrThrow(!CreateDirectory(aDirSpec, NULL));
}



ResultType Line::FileRead(LPTSTR aFilespec)
// Returns OK or FAIL.  Will almost always return OK because if an error occurs,
// the script's ErrorLevel variable will be set accordingly.  However, if some
// kind of unexpected and more serious error occurs, such as variable-out-of-memory,
// that will cause FAIL to be returned.
{
	Var &output_var = *OUTPUT_VAR;
	// Init output var to be blank as an additional indicator of failure (or empty file).
	// Caller must check ErrorLevel to distinguish between an empty file and an error.
	output_var.Assign();

	const DWORD DWORD_MAX = ~0;

	// Set default options:
	bool translate_crlf_to_lf = false;
	bool is_binary_clipboard = false;
	unsigned __int64 max_bytes_to_load = ULLONG_MAX; // By default, fail if the file is too large.  See comments near bytes_to_read below.
	UINT codepage = g->Encoding & CP_AHKCP;

	// It's done as asterisk+option letter to permit future expansion.  A plain asterisk such as used
	// by the FileAppend command would create ambiguity if there was ever an effort to add other asterisk-
	// prefixed options later.
	LPTSTR cp;
	for (;;)
	{
		cp = omit_leading_whitespace(aFilespec); // omit leading whitespace only temporarily in case aFilespec contains literal whitespace we want to retain.
		if (*cp != '*') // No more options.
			break; // Make no further changes to aFilespec.
		switch (ctoupper(*++cp)) // This could move cp to the terminator if string ends in an asterisk.
		{
		case 'C': // Clipboard (binary).
			is_binary_clipboard = true; // When this option is present, any others are parsed (to skip over them) but ignored as documented.
			break;
		case 'M': // Maximum number of bytes to load.
			max_bytes_to_load = ATOU64(cp + 1); // Relies upon the fact that it ceases conversion upon reaching a space or tab.
			// Skip over the digits of this option in case it's the last option.
			if (   !(cp = StrChrAny(cp, _T(" \t")))   ) // Find next space or tab (there should be one if options are properly formatted).
				return SetErrorsOrThrow(true, ERROR_INVALID_PARAMETER);
			--cp; // Standardize it to make it conform to the other options, for use below.
			break;
		case 'P':
			codepage = _ttoi(cp + 1);
			// Skip over the digits of this option in case it's the last option.
			if (   !(cp = StrChrAny(cp, _T(" \t")))   ) // Find next space or tab (there should be one if options are properly formatted).
				return SetErrorsOrThrow(true, ERROR_INVALID_PARAMETER);
			--cp; // Standardize it to make it conform to the other options, for use below.
			break;
		case 'T': // Text mode.
			translate_crlf_to_lf = true;
			break;
		}
		// Note: because it's possible for filenames to start with a space (even though Explorer itself
		// won't let you create them that way), allow exactly one space between end of option and the
		// filename itself:
		aFilespec = cp;  // aFilespec is now the option letter after the asterisk *or* empty string if there was none.
		if (*aFilespec)
		{
			++aFilespec;
			// Now it's the space or tab (if there is one) after the option letter.  It seems best for
			// future expansion to assume that this is a space or tab even if it's really the start of
			// the filename.  For example, in the future, multiletter options might be wanted, in which
			// case allowing the omission of the space or tab between *t and the start of the filename
			// would cause the following to be ambiguous:
			// FileRead, OutputVar, *delimC:\File.txt
			// (assuming *delim would accept an optional arg immediately following it).
			// Enforcing this format also simplifies parsing in the future, if there are ever multiple options.
			// It also conforms to the precedent/behavior of GuiControl when it accepts picture sizing options
			// such as *w/h and *x/y
			if (*aFilespec)
				++aFilespec; // And now it's the start of the filename or the asterisk of the next option.
							// This behavior is as documented in the help file.
		}
	} // for()

	// It seems more flexible to allow other processes to read and write the file while we're reading it.
	// For example, this allows the file to be appended to during the read operation, which could be
	// desirable, especially it's a very large log file that would take a long time to read.
	// MSDN: "To enable other processes to share the object while your process has it open, use a combination
	// of one or more of [FILE_SHARE_READ, FILE_SHARE_WRITE]."
	HANDLE hfile = CreateFile(aFilespec, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING
		, FILE_FLAG_SEQUENTIAL_SCAN, NULL); // MSDN says that FILE_FLAG_SEQUENTIAL_SCAN will often improve performance
	if (hfile == INVALID_HANDLE_VALUE)      // in cases like these (and it seems best even if max_bytes_to_load was specified).
		return SetErrorsOrThrow(true);

	unsigned __int64 bytes_to_read = GetFileSize64(hfile);
	if (bytes_to_read == ULLONG_MAX) // GetFileSize64() failed.
	{
		g->LastError = GetLastError();
		CloseHandle(hfile);
		return SetErrorLevelOrThrow();
	}
	// In addition to imposing the limit set by the *M option, the following check prevents an error
	// caused by 64 to 32-bit truncation -- that is, a file size of 0x100000001 would be truncated to
	// 0x1, allowing the command to complete even though it should fail.  UPDATE: This check was never
	// sufficient since max_bytes_to_load could exceed DWORD_MAX on x64 (prior to v1.1.16).  It's now
	// checked separately below to try to match the documented behaviour (truncating the data only to
	// the caller-specified limit).
	if (bytes_to_read > max_bytes_to_load) // This is the limit set by the caller.
		bytes_to_read = max_bytes_to_load;
	// Fixed for v1.1.16: Show an error message if the file is larger than DWORD_MAX, otherwise the
	// truncation issue described above could occur.  Reading more than DWORD_MAX could be supported
	// by calling ReadFile() in a loop, but it seems unlikely that a script will genuinely want to
	// do this AND actually be able to allocate a 4GB+ memory block (having 4GB of total free memory
	// is usually not sufficient, perhaps due to memory fragmentation).
#ifdef _WIN64
	if (bytes_to_read > DWORD_MAX)
#else
	// Reserve 2 bytes to avoid integer overflow below.  Although any amount larger than 2GB is almost
	// guaranteed to fail at the malloc stage, that might change if we ever become large address aware.
	if (bytes_to_read > DWORD_MAX - sizeof(wchar_t))
#endif
		return LineError(ERR_OUTOFMEM); // Using this instead of "File too large." to reduce code size, since this condition is very rare (and malloc succeeding would be even rarer).

	if (!bytes_to_read)
	{
		CloseHandle(hfile);
		return SetErrorsOrThrow(false, 0); // Indicate success (a zero-length file results in empty output_var).
	}

	LPBYTE output_buf;
	bool output_buf_is_var = is_binary_clipboard && output_var.Type() != VAR_CLIPBOARD;
	if (output_buf_is_var) 
	{
		// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
		// this call will set up the clipboard for writing:
		if (output_var.SetCapacity(VarSizeType(bytes_to_read) + (sizeof(wchar_t) - sizeof(TCHAR)), true, false) == OK) // SetCapacity() reserves 1 TCHAR for null-terminator.  Allow an extra byte on ANSI builds for wchar_t.
			output_buf = (LPBYTE) output_var.Contents();
		else
			output_buf = NULL; // Above already displayed the error message.
	}
	else
	{
		// Either we're reading text and need an intermediate buffer to allow text conversion,
		// or we're reading binary clipboard data into the Clipboard and need a temporary buffer
		// to read into before calling SetClipboardData().
		output_buf = (LPBYTE) malloc(size_t(bytes_to_read + sizeof(wchar_t)));
		if (!output_buf)
			LineError(ERR_OUTOFMEM);
	}
	if (!output_buf)
	{
		CloseHandle(hfile);
		// ErrorLevel doesn't matter now because the current quasi-thread will be aborted.
		return FAIL;
	}

	DWORD bytes_actually_read;
	BOOL result = ReadFile(hfile, output_buf, (DWORD)bytes_to_read, &bytes_actually_read, NULL);
	g->LastError = GetLastError();
	CloseHandle(hfile);

	// Upon result==success, bytes_actually_read is not checked against bytes_to_read because it
	// shouldn't be different (result should have set to failure if there was a read error).
	// If it ever is different, a partial read is considered a success since ReadFile() told us
	// that nothing bad happened.

	if (result)
	{
		if (!is_binary_clipboard) // text mode, do UTF-8 and UTF-16LE BOM checking
		{
			bool has_bom;
			if (bytes_actually_read >= 3 && output_buf[0] == 0xEF && output_buf[1] == 0xBB && output_buf[2] == 0xBF) // UTF-8 BOM
			{
				if (!output_var.AssignStringFromUTF8((LPCSTR)output_buf + 3, bytes_actually_read - 3))
					result = FALSE;
			}
			else if ( (has_bom = (bytes_actually_read >= 2 && output_buf[0] == 0xFF && output_buf[1] == 0xFE)) // UTF-16LE BOM
					|| codepage == CP_UTF16 ) // Covers FileEncoding UTF-16 and FileEncoding UTF-16-RAW.
			{
				LPCWSTR text_start = (LPCWSTR)output_buf;
				DWORD text_size = bytes_actually_read;
				if (has_bom) {
					text_start ++; // Skip BOM.
					text_size -= 2; // Exclude BOM from calculations below for consistency; include only the actual data.
				}
				if (!output_var.AssignStringW(text_start, text_size / sizeof(wchar_t)))
					result = FALSE;
			}
			else
			{
#ifndef UNICODE
				if (codepage == CP_ACP || codepage == GetACP())
				{
					// Avoid any unnecessary conversion or copying by using our malloc'd buffer directly.
					// This should be worth doing since the string must otherwise be converted to UTF-16 and back.
					output_buf[bytes_actually_read] = 0; // Ensure text is terminated where indicated.
					output_var.AcceptNewMem((LPTSTR)output_buf, bytes_actually_read);
					output_buf = NULL; // AcceptNewMem took charge of it.
				}
				else
#endif
				if (!output_var.AssignStringFromCodePage((LPCSTR)output_buf, bytes_actually_read, codepage))
					result = FALSE;
			}
			if (output_buf) // i.e. it wasn't "claimed" above.
				free(output_buf);
			output_buf = (LPBYTE) output_var.Contents();
			if (translate_crlf_to_lf)
			{
				// Since a larger string is being replaced with a smaller, there's a good chance the 2 GB
				// address limit will not be exceeded by StrReplace even if the file is close to the
				// 1 GB limit as described above:
				VarSizeType var_length = output_var.CharLength();
				StrReplace((LPTSTR)output_buf, _T("\r\n"), _T("\n"), SCS_SENSITIVE, UINT_MAX, -1, NULL, &var_length);
				output_var.SetCharLength(var_length);
			}
		}
		else // is_binary_clipboard == true
		{
			if (output_var.Type() == VAR_CLIPBOARD) // Reading binary clipboard data directly back onto the clipboard.
			{
				result = Var::SetClipboardAll(output_buf, bytes_actually_read) == OK;
				free(output_buf);
				if (!result)
					return FAIL;
				return SetErrorLevelOrThrowBool(false);
			}
			// Although binary clipboard data is always null-terminated, this might be some other kind
			// of binary data or actually text (but the caller passed *c to skip codepage conversion),
			// so might not be terminated.  Ensure the data is null-terminated:
			DWORD terminate_at = bytes_actually_read;
#ifdef UNICODE
			// Since the data might be interpreted as UTF-16, we need to ensure the null-terminator is
			// aligned correctly, like "xxxx 0000" and not "xx00 00??" (where xx is valid data and ??
			// is an uninitialized byte).
			if (terminate_at & 1) // Odd number of bytes.
				output_buf[terminate_at++] = 0; // Put an extra zero byte in and move the actual terminator right one byte.
#endif
			*LPTSTR(output_buf + terminate_at) = 0;
			// Update the output var's length.  In this case the script wants the actual data size rather
			// than the "usable" length.  v1.1.16: Although it might change the behaviour of some scripts,
			// it seems safer to use the "rounded up" size than an odd byte count, which would cause the
			// last byte to be truncated due to integer division in Var::CharLength().
			output_var.ByteLength() = terminate_at;
		}
	}
	else
	{
		// ReadFile() failed.  Since MSDN does not document what is in the buffer at this stage,
		// or whether what's in it is even null-terminated, or even whether bytes_to_read contains
		// a valid value, it seems best to abort the entire operation rather than try to put partial
		// file contents into output_var.  ErrorLevel will indicate the failure.
		// Since ReadFile() failed, to avoid complications or side-effects in functions such as Var::Close(),
		// avoid storing a potentially non-terminated string in the variable.
		*((LPTSTR)output_buf) = '\0'; // Assign() at this point would fail for the clipboard since it's already open for writing.
		output_var.ByteLength() = 0;
		if (!output_buf_is_var)
			free(output_buf);
	}

	if (!output_var.Close(is_binary_clipboard)) // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_CLIPBOARD.
		return FAIL;

	return SetErrorLevelOrThrowBool(!result);
}



ResultType Line::FileReadLine(LPTSTR aFilespec, LPTSTR aLineNumber)
// Returns OK or FAIL.  Will almost always return OK because if an error occurs,
// the script's ErrorLevel variable will be set accordingly.  However, if some
// kind of unexpected and more serious error occurs, such as variable-out-of-memory,
// that will cause FAIL to be returned.
{
	Var &output_var = *OUTPUT_VAR; // Fix for v1.0.45.01: Must be resolved and saved before MsgSleep() (LONG_OPERATION) because that allows some other thread to interrupt and overwrite sArgVar[].

	__int64 line_number = ATOI64(aLineNumber);
	if (line_number < 1)
		return SetErrorsOrThrow(true, ERROR_INVALID_PARAMETER);

	TextFile tfile;
	if (!tfile.Open(aFilespec, DEFAULT_READ_FLAGS, g->Encoding & CP_AHKCP))
		return SetErrorsOrThrow(true);

	// Remember that once the first call to MsgSleep() is done, a new hotkey subroutine
	// may fire and suspend what we're doing here.  Such a subroutine might also overwrite
	// the values our params, some of which may be in the deref buffer.  So be sure not
	// to refer to those strings once MsgSleep() has been done, below.  Alternatively,
	// a copy of such params can be made using our own stack space.

	LONG_OPERATION_INIT

	DWORD buf_length;
	TCHAR buf[READ_FILE_LINE_SIZE];
	for (__int64 i = 0; i < line_number; ++i)
	{
		if (  !(buf_length = tfile.ReadLine(buf, _countof(buf) - 1))  ) // end-of-file or error
		{
			g->LastError = GetLastError();
			tfile.Close();
			return SetErrorLevelOrThrow();
		}
		LONG_OPERATION_UPDATE
	}
	tfile.Close();

	if (buf_length && buf[buf_length - 1] == '\n') // Remove any trailing newline for the user.
		--buf_length;

	if (!buf_length)
	{
		if (!output_var.Assign()) // Explicitly call it this way so that it won't free the memory.
			return FAIL;
	}
	else
		if (!output_var.Assign(buf, (VarSizeType)buf_length))
			return FAIL;

	return SetErrorsOrThrow(false, 0); // Indicate success.
}



ResultType Line::FileAppend(LPTSTR aFilespec, LPTSTR aBuf, LoopReadFileStruct *aCurrentReadFile)
{
	// The below is avoided because want to allow "nothing" to be written to a file in case the
	// user is doing this to reset it's timestamp (or create an empty file).
	//if (!aBuf || !*aBuf)
	//	return g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	if (aCurrentReadFile) // It always takes precedence over aFilespec.
		aFilespec = aCurrentReadFile->mWriteFileName;
	if (!*aFilespec) // Nothing to write to (caller relies on this check).
		return SetErrorsOrThrow(true, ERROR_INVALID_PARAMETER);

	TextStream *ts = aCurrentReadFile ? aCurrentReadFile->mWriteFile : NULL;
	bool file_was_already_open = ts;
	BOOL result;

	bool open_as_binary = (*aFilespec == '*');
	if (open_as_binary)
	{
		if (aFilespec[1] && (aFilespec[1] != '*' || aFilespec[2])) // i.e. it's not just * (stdout) or ** (stderr).
		{
			// Do not do this because it's possible for filenames to start with a space
			// (even though Explorer itself won't let you create them that way):
			//write_filespec = omit_leading_whitespace(write_filespec + 1);
			// Instead just do this:
			++aFilespec;
		}
#ifdef CONFIG_DEBUGGER
		else if (!aFilespec[1] && g_Debugger.FileAppendStdOut(aBuf))
		{
			// StdOut has been redirected to the debugger, so return.
			return SetErrorsOrThrow(false, 0);
		}
#endif
	}
	else if (!file_was_already_open) // As of 1.0.25, auto-detect binary if that mode wasn't explicitly specified.
	{
		// sArgVar is used for two reasons:
		// 1) It properly resolves dynamic variables, such as "FileAppend, % %VarContainingTheStringClipboardAll%, File".
		// 2) It resolves them only once at a prior stage, rather than having to do them again here
		//    (which helps performance).
		if (ARGVAR1)
		{
			if (ARGVAR1->Type() == VAR_CLIPBOARDALL)
				return WriteClipboardToFile(aFilespec);
			else if (ARGVAR1->IsBinaryClip())
			{
				// Since there is at least one deref in Arg #1 and the first deref is binary clipboard,
				// assume this operation's only purpose is to write binary data from that deref to a file.
				// This is because that's the only purpose that seems useful and that's currently supported.
				// In addition, the file is always overwritten in this mode, since appending clipboard data
				// to an existing clipboard file would not work due to:
				// 1) Duplicate clipboard formats not making sense (i.e. two CF_TEXT formats would cause the
				//    first to be overwritten by the second when restoring to clipboard).
				// 2) There is a 4-byte zero terminator at the end of the file.
				return WriteClipboardToFile(aFilespec, ARGVAR1);
			}
		}
		// Auto-detection avoids the need to have to translate \r\n to \n when reading
		// a file via the FileRead command.  This seems extremely unlikely to break any
		// existing scripts because the intentional use of \r\r\n in a text file (two
		// consecutive carriage returns) -- which would happen if \r\n were written in
		// text mode -- is so rare as to be close to non-existent.  If this behavior
		// is ever specifically needed, the script can explicitly places some \r\r\n's
		// in the file and then write it as binary mode.
		open_as_binary = _tcsstr(aBuf, _T("\r\n")); // Performance: The following could be done instead, but seems likely to cause some scripts to write \r\r\n and even \r\r\r\n due to the text having both \n and \r\n in it: char *first_newline = _tcschr(aBuf, '\n')... open_as_binary = first_newline > aBuf && aBuf[-1] == '\r'
		// Due to "else if", the above will not turn off binary mode if binary was explicitly specified.
		// That is useful to write Unix style text files whose lines end in solitary linefeeds.
	}

	// Check if the file needs to be opened.  As of 1.0.25, this is done here rather than
	// at the time the loop first begins so that:
	// 1) Binary mode can be auto-detected if the first block of text appended to the file
	//    contains any \r\n's.
	// 2) To avoid opening the file if the file-reading loop has zero iterations (i.e. it's
	//    opened only upon first actual use to help performance and avoid changing the
	//    file-modification time when no actual text will be appended).
	if (!file_was_already_open)
	{
		DWORD flags = TextStream::APPEND | (open_as_binary ? 0 : TextStream::EOL_CRLF);
		
		UINT codepage = mArgc > 2 ? ConvertFileEncoding(ARG3) : g->Encoding;
		if (codepage == -1) // ARG3 was invalid.
			return SetErrorsOrThrow(true, ERROR_INVALID_PARAMETER);
		
		ASSERT( (~CP_AHKNOBOM) == CP_AHKCP );
		// codepage may include CP_AHKNOBOM, in which case below will not add BOM_UTFxx flag.
		if (codepage == CP_UTF8)
			flags |= TextStream::BOM_UTF8;
		else if (codepage == CP_UTF16)
			flags |= TextStream::BOM_UTF16;

		// Open the output file (if one was specified).  Unlike the input file, this is not
		// a critical error if it fails.  We want it to be non-critical so that FileAppend
		// commands in the body of the loop will set ErrorLevel to indicate the problem:
		if (  !(ts = new TextFile)  ) // ts was alredy verified NULL via !file_was_already_open.
			return LineError(ERR_OUTOFMEM);
		if ( !ts->Open(aFilespec, flags, codepage & CP_AHKCP) )
		{
			delete ts; // Must be deleted explicitly!
			return SetErrorsOrThrow(true);
		}
		if (aCurrentReadFile)
			aCurrentReadFile->mWriteFile = ts;
	}

	// Write to the file:
	DWORD length = (DWORD)_tcslen(aBuf);
	result = length && ts->Write(aBuf, length) == 0; // Relies on short-circuit boolean evaluation.  If buf is empty, we've already succeeded in creating the file and have nothing further to do.

	if (!aCurrentReadFile)
		delete ts;
	// else it's the caller's responsibility, or it's caller's, to close it.

	return SetErrorsOrThrow(result);
}



ResultType Line::WriteClipboardToFile(LPTSTR aFilespec, Var *aBinaryClipVar)
// Returns OK or FAIL.  If OK, it sets ErrorLevel to the appropriate result.
// If the clipboard is empty, a zero length file will be written, which seems best for its consistency.
{
	LPVOID data;
	size_t data_size;
	if (aBinaryClipVar)
	{
		// Get clipboard data from a variable.
		data = aBinaryClipVar->Contents();
		data_size = aBinaryClipVar->ByteLength();
	}
	else
	{
		// Get the clipboard's current contents.
		if (!Var::GetClipboardAll(NULL, &data, &data_size))
		{
			g->LastError = 0; // To avoid possible confusion, don't leave it at its previous value.
			return FAIL;
		}
	}

	BOOL success = FALSE; // Set default.
	DWORD bytes_to_write = (DWORD)data_size;

	HANDLE hfile = CreateFile(aFilespec, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, 0, NULL); // Overwrite. Unsharable (since reading the file while it is being written would probably produce bad data in this case).
	if (hfile != INVALID_HANDLE_VALUE)
	{
		DWORD bytes_written = 0;
		if (data) // Can be NULL if data_size == 0.
			WriteFile(hfile, data, bytes_to_write, &bytes_written, NULL);
		success = (bytes_written == bytes_to_write); // Even if both are zero.
	}
	g->LastError = GetLastError(); // Always done, for simplicity.  Must be called before CloseHandle().
	if (hfile != INVALID_HANDLE_VALUE)
		CloseHandle(hfile); // Close file.
	if (!aBinaryClipVar)
		free(data); // Free ClipboardAll data.  Can be NULL.

	return SetErrorLevelOrThrowBool(!success); // Set ErrorLevel based on result.
}



BOOL FileDeleteCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData)
{
	return DeleteFile(aFilename);
}

ResultType Line::FileDelete(LPTSTR aFilePattern)
{
	// The no-wildcard case could be handled via FilePatternApply(), but it is handled here
	// for backward-compatibility (just in case the use of FindFirstFile() affects something):
	if (!StrChrAny(aFilePattern, _T("?*")) // No wildcards; just a plain path/filename.
		&& *aFilePattern) // i.e. if it's empty, let FilePatternApply() report the error.
	{
		SetLastError(0); // For sanity: DeleteFile appears to set it only on failure.
		// ErrorLevel will indicate failure if DeleteFile fails.
		return SetErrorsOrThrow(!DeleteFile(aFilePattern));
	}

	// Otherwise aFilePattern contains wildcards, so we'll search for all matches and delete them.
	FilePatternApply(aFilePattern, FILE_LOOP_FILES_ONLY, false, FileDeleteCallback, NULL);
	return g->ThrownToken ? FAIL : OK;
}



ResultType Line::FileInstall(LPTSTR aSource, LPTSTR aDest, LPTSTR aFlag)
{
	bool success;
	bool allow_overwrite = (ATOI(aFlag) == 1);
#ifdef AUTOHOTKEYSC
	if (!allow_overwrite && Util_DoesFileExist(aDest))
		return SetErrorLevelOrThrow();

	// Open the file first since it's the most likely to fail:
	HANDLE hfile = CreateFile(aDest, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, 0, NULL);
	if (hfile == INVALID_HANDLE_VALUE)
		return SetErrorLevelOrThrow();

	// Create a temporary copy of aSource to ensure it is the correct case (upper-case).
	// Ahk2Exe converts it to upper-case before adding the resource. My testing showed that
	// using lower or mixed case in some instances prevented the resource from being found.
	// Since file paths are case-insensitive, it certainly doesn't seem harmful to do this:
	TCHAR source[MAX_PATH];
	size_t source_length = _tcslen(aSource);
	if (source_length >= _countof(source))
		// Probably can't happen; for simplicity, truncate it.
		source_length = _countof(source) - 1;
	tmemcpy(source, aSource, source_length + 1);
	_tcsupr(source);

	// Find and load the resource.
	HRSRC res;
	HGLOBAL res_load;
	LPVOID res_lock;
	if ( (res = FindResource(NULL, source, RT_RCDATA))
	  && (res_load = LoadResource(NULL, res))
	  && (res_lock = LockResource(res_load))  )
	{
		DWORD num_bytes_written;
		// Write the resource data to file.
		success = WriteFile(hfile, res_lock, SizeofResource(NULL, res), &num_bytes_written, NULL);
	}
	else
		success = false;
	CloseHandle(hfile);

#else // AUTOHOTKEYSC not defined:

	// v1.0.35.11: Must search in A_ScriptDir by default because that's where ahk2exe will search by default.
	// The old behavior was to search in A_WorkingDir, which seems pointless because ahk2exe would never
	// be able to use that value if the script changes it while running.
	TCHAR aDestPath[MAX_PATH];
	GetFullPathName(aDest, MAX_PATH, aDestPath, NULL);
	SetCurrentDirectory(g_script.mFileDir);
	success = CopyFile(aSource, aDestPath, !allow_overwrite);
	SetCurrentDirectory(g_WorkingDir); // Restore to proper value.

#endif

	return SetErrorLevelOrThrowBool(!success);
}



ResultType Line::FileGetAttrib(LPTSTR aFilespec)
{
	OUTPUT_VAR->Assign(); // Init to be blank, in case of failure.

	if (!aFilespec || !*aFilespec)
		return SetErrorsOrThrow(true, ERROR_INVALID_PARAMETER);

	DWORD attr = GetFileAttributes(aFilespec);
	if (attr == 0xFFFFFFFF)  // Failure, probably because file doesn't exist.
		return SetErrorsOrThrow(true);

	SetErrorsOrThrow(false, 0);
	TCHAR attr_string[128];
	return OUTPUT_VAR->Assign(FileAttribToStr(attr_string, attr));
}



BOOL FileSetAttribCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData);
struct FileSetAttribData
{
	DWORD and_mask, xor_mask;
};

int Line::FileSetAttrib(LPTSTR aAttributes, LPTSTR aFilePattern, FileLoopModeType aOperateOnFolders
	, bool aDoRecurse, bool aCalledRecursively)
// Returns the number of files and folders that could not be changed due to an error.
{
	// Convert the attribute string to three bit-masks: add, remove and toggle.
	FileSetAttribData attrib;
	DWORD mask;
	int op = 0;
	attrib.and_mask = 0xFFFFFFFF; // Set default: keep all bits.
	attrib.xor_mask = 0; // Set default: affect none.
	for (LPTSTR cp = aAttributes; *cp; ++cp)
	{
		switch (ctoupper(*cp))
		{
		case '+':
		case '-':
		case '^':
			op = *cp;
		default:
			continue;
		// Note that D (directory) and C (compressed) are currently not supported:
		case 'R': mask = FILE_ATTRIBUTE_READONLY; break;
		case 'A': mask = FILE_ATTRIBUTE_ARCHIVE; break;
		case 'S': mask = FILE_ATTRIBUTE_SYSTEM; break;
		case 'H': mask = FILE_ATTRIBUTE_HIDDEN; break;
		// N: Docs say it's valid only when used alone.  But let the API handle it if this is not so.
		case 'N': mask = FILE_ATTRIBUTE_NORMAL; break;
		case 'O': mask = FILE_ATTRIBUTE_OFFLINE; break;
		case 'T': mask = FILE_ATTRIBUTE_TEMPORARY; break;
		}
		switch (op)
		{
		case '+':
			attrib.and_mask &= ~mask; // Reset bit to 0.
			attrib.xor_mask |= mask; // Set bit to 1.
			break;
		case '-':
			attrib.and_mask &= ~mask; // Reset bit to 0.
			attrib.xor_mask &= ~mask; // Override any prior + or ^.
			break;
		case '^':
			attrib.xor_mask ^= mask; // Toggle bit.  ^= vs |= to invert any prior + or ^.
			// Leave and_mask as is, so any prior + or - will be inverted.
			break;
		}
	}
	return FilePatternApply(aFilePattern, aOperateOnFolders, aDoRecurse, FileSetAttribCallback, &attrib);
}

BOOL FileSetAttribCallback(LPTSTR file_path, WIN32_FIND_DATA &current_file, void *aCallbackData)
{
	FileSetAttribData &attrib = *(FileSetAttribData *)aCallbackData;
	DWORD file_attrib = ((current_file.dwFileAttributes & attrib.and_mask) ^ attrib.xor_mask);
	if (!SetFileAttributes(file_path, file_attrib))
	{
		g->LastError = GetLastError();
		return FALSE;
	}
	return TRUE;
}



int Line::FilePatternApply(LPTSTR aFilePattern, FileLoopModeType aOperateOnFolders
	, bool aDoRecurse, FilePatternCallback aCallback, void *aCallbackData
	, bool aCalledRecursively)
{
	if (!aCalledRecursively)  // i.e. Only need to do this if we're not called by ourself:
	{
		if (!*aFilePattern)
		{
			SetErrorsOrThrow(true, ERROR_INVALID_PARAMETER);
			return 0;
		}
		if (aOperateOnFolders == FILE_LOOP_INVALID) // In case runtime dereference of a var was an invalid value.
			aOperateOnFolders = FILE_LOOP_FILES_ONLY;  // Set default.
		g->LastError = 0; // Set default. Overridden only when a failure occurs.
	}

	if (_tcslen(aFilePattern) >= MAX_PATH) // Checked early to simplify other things below.
	{
		SetErrorsOrThrow(true, ERROR_BUFFER_OVERFLOW);
		return 0;
	}

	// Testing shows that the ANSI version of FindFirstFile() will not accept a path+pattern longer
	// than 256 or so, even if the pattern would match files whose names are short enough to be legal.
	// Therefore, as of v1.0.25, there is also a hard limit of MAX_PATH on all these variables.
	// MSDN confirms this in a vague way: "In the ANSI version of FindFirstFile(), [plpFileName] is
	// limited to MAX_PATH characters."
	TCHAR file_pattern[MAX_PATH], file_path[MAX_PATH]; // Giving +3 extra for "*.*" seems fairly pointless because any files that actually need that extra room would fail to be retrieved by FindFirst/Next due to their inability to support paths much over 256.
	_tcscpy(file_pattern, aFilePattern); // Make a copy in case of overwrite of deref buf during LONG_OPERATION/MsgSleep.
	_tcscpy(file_path, aFilePattern);    // An earlier check has ensured these won't overflow.

	size_t file_path_length; // The length of just the path portion of the filespec.
	LPTSTR last_backslash = _tcsrchr(file_path, '\\');
	if (last_backslash)
	{
		// Remove the filename and/or wildcard part.   But leave the trailing backslash on it for
		// consistency with below:
		*(last_backslash + 1) = '\0';
		file_path_length = _tcslen(file_path);
	}
	else // Use current working directory, e.g. if user specified only *.*
	{
		*file_path = '\0';
		file_path_length = 0;
	}
	LPTSTR append_pos = file_path + file_path_length; // For performance, copy in the unchanging part only once.  This is where the changing part gets appended.
	size_t space_remaining = _countof(file_path) - file_path_length - 1; // Space left in file_path for the changing part.

	// For use with aDoRecurse, get just the naked file name/pattern:
	LPTSTR naked_filename_or_pattern = _tcsrchr(file_pattern, '\\');
	if (naked_filename_or_pattern)
		++naked_filename_or_pattern;
	else
		naked_filename_or_pattern = file_pattern;

	if (!StrChrAny(naked_filename_or_pattern, _T("?*")))
		// Since no wildcards, always operate on this single item even if it's a folder.
		aOperateOnFolders = FILE_LOOP_FILES_AND_FOLDERS;

	LONG_OPERATION_INIT
	int failure_count = 0;
	WIN32_FIND_DATA current_file;
	HANDLE file_search = FindFirstFile(file_pattern, &current_file);

	if (file_search != INVALID_HANDLE_VALUE)
	{
		do
		{
			// Since other script threads can interrupt during LONG_OPERATION_UPDATE, it's important that
			// this command not refer to sArgDeref[] and sArgVar[] anytime after an interruption becomes
			// possible. This is because an interrupting thread usually changes the values to something
			// inappropriate for this thread.
			LONG_OPERATION_UPDATE

			if (current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
			{
				if (current_file.cFileName[0] == '.' && (!current_file.cFileName[1]    // Relies on short-circuit boolean order.
					|| current_file.cFileName[1] == '.' && !current_file.cFileName[2]) //
					// Regardless of whether this folder will be recursed into, this folder
					// will not be affected when the mode is files-only:
					|| aOperateOnFolders == FILE_LOOP_FILES_ONLY)
					continue; // Never operate upon or recurse into these.
			}
			else // It's a file, not a folder.
				if (aOperateOnFolders == FILE_LOOP_FOLDERS_ONLY)
					continue;

			if (_tcslen(current_file.cFileName) > space_remaining)
			{
				// v1.0.45.03: Don't even try to operate upon truncated filenames in case they accidentally
				// match the name of a real/existing file.
				g->LastError = ERROR_BUFFER_OVERFLOW;
				++failure_count;
				continue;
			}
			// Otherwise, make file_path be the filespec of the file to operate upon:
			_tcscpy(append_pos, current_file.cFileName); // Above has ensured this won't overflow.
			//
			// This is the part that actually does something to the file:
			if (!aCallback(file_path, current_file, aCallbackData))
				++failure_count;
			//
		} while (FindNextFile(file_search, &current_file));

		FindClose(file_search);
	} // if (file_search != INVALID_HANDLE_VALUE)

	if (aDoRecurse && space_remaining > 2) // The space_remaining check ensures there's enough room to append "*.*" (if not, just avoid recursing into it due to rarity).
	{
		// Testing shows that the ANSI version of FindFirstFile() will not accept a path+pattern longer
		// than 256 or so, even if the pattern would match files whose names are short enough to be legal.
		// Therefore, as of v1.0.25, there is also a hard limit of MAX_PATH on all these variables.
		// MSDN confirms this in a vague way: "In the ANSI version of FindFirstFile(), [plpFileName] is
		// limited to MAX_PATH characters."
		_tcscpy(append_pos, _T("*.*")); // Above has ensured this won't overflow.
		file_search = FindFirstFile(file_path, &current_file);

		if (file_search != INVALID_HANDLE_VALUE)
		{
			size_t pattern_length = _tcslen(naked_filename_or_pattern);
			do
			{
				LONG_OPERATION_UPDATE
				if (!(current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
					|| current_file.cFileName[0] == '.' && (!current_file.cFileName[1]     // Relies on short-circuit boolean order.
						|| current_file.cFileName[1] == '.' && !current_file.cFileName[2]) //
					// v1.0.45.03: Skip over folders whose full-path-names are too long to be supported by the ANSI
					// versions of FindFirst/FindNext.  Without this fix, it might be possible for infinite recursion
					// to occur (see PerformLoop() for more comments).
					|| pattern_length + _tcslen(current_file.cFileName) >= space_remaining) // >= vs. > to reserve 1 for the backslash to be added between cFileName and naked_filename_or_pattern.
					continue; // Never recurse into these.
				// This will build the string CurrentDir+SubDir+FilePatternOrName.
				// If FilePatternOrName doesn't contain a wildcard, the recursion
				// process will attempt to operate on the originally-specified
				// single filename or folder name if it occurs anywhere else in the
				// tree, e.g. recursing C:\Temp\temp.txt would affect all occurrences
				// of temp.txt both in C:\Temp and any subdirectories it might contain:
				_stprintf(append_pos, _T("%s\\%s") // Above has ensured this won't overflow.
					, current_file.cFileName, naked_filename_or_pattern);
				//
				// Apply the callback to files in this subdirectory:
				failure_count += FilePatternApply(file_path, aOperateOnFolders, aDoRecurse, aCallback, aCallbackData, true);
				//
			} while (FindNextFile(file_search, &current_file));
			FindClose(file_search);
		} // if (file_search != INVALID_HANDLE_VALUE)
	} // if (aDoRecurse)

	if (!aCalledRecursively) // i.e. Only need to do this if we're returning to top-level caller:
		SetErrorLevelOrThrowInt(failure_count); // i.e. indicate success if there were no failures.
	return failure_count;
}



ResultType Line::FileGetTime(LPTSTR aFilespec, TCHAR aWhichTime)
{
	OUTPUT_VAR->Assign(); // Init to be blank, in case of failure.

	if (!aFilespec || !*aFilespec)
		return SetErrorsOrThrow(true, ERROR_INVALID_PARAMETER);

	// Don't use CreateFile() & FileGetSize() size they will fail to work on a file that's in use.
	// Research indicates that this method has no disadvantages compared to the other method.
	WIN32_FIND_DATA found_file;
	HANDLE file_search = FindFirstFile(aFilespec, &found_file);
	if (file_search == INVALID_HANDLE_VALUE)
		return SetErrorsOrThrow(true);
	FindClose(file_search);

	FILETIME local_file_time;
	switch (ctoupper(aWhichTime))
	{
	case 'C': // File's creation time.
		FileTimeToLocalFileTime(&found_file.ftCreationTime, &local_file_time);
		break;
	case 'A': // File's last access time.
		FileTimeToLocalFileTime(&found_file.ftLastAccessTime, &local_file_time);
		break;
	default:  // 'M', unspecified, or some other value.  Use the file's modification time.
		FileTimeToLocalFileTime(&found_file.ftLastWriteTime, &local_file_time);
	}

	SetErrorsOrThrow(false, 0); // Indicate success.
	TCHAR local_file_time_string[128];
	return OUTPUT_VAR->Assign(FileTimeToYYYYMMDD(local_file_time_string, local_file_time));
}



BOOL FileSetTimeCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData);
struct FileSetTimeData
{
	FILETIME Time;
	TCHAR WhichTime;
};

int Line::FileSetTime(LPTSTR aYYYYMMDD, LPTSTR aFilePattern, TCHAR aWhichTime
	, FileLoopModeType aOperateOnFolders, bool aDoRecurse, bool aCalledRecursively)
// Returns the number of files and folders that could not be changed due to an error.
{
	// Related to the comment at the top: Since the script subroutine that resulted in the call to
	// this function can be interrupted during our MsgSleep(), make a copy of any params that might
	// currently point directly to the deref buffer.  This is done because their contents might
	// be overwritten by the interrupting subroutine:
	TCHAR yyyymmdd[64]; // Even do this one since its value is passed recursively in calls to self.
	tcslcpy(yyyymmdd, aYYYYMMDD, _countof(yyyymmdd));

	FileSetTimeData callbackData;
	callbackData.WhichTime = aWhichTime;
	FILETIME ft;
	if (*yyyymmdd)
	{
		// Convert the arg into the time struct as local (non-UTC) time:
		if (!YYYYMMDDToFileTime(yyyymmdd, ft))
		{
			SetErrorsOrThrow(true);
			return 0;
		}
		// Convert from local to UTC:
		if (!LocalFileTimeToFileTime(&ft, &callbackData.Time))
		{
			SetErrorsOrThrow(true);
			return 0;
		}
	}
	else // User wants to use the current time (i.e. now) as the new timestamp.
		GetSystemTimeAsFileTime(&callbackData.Time);

	return FilePatternApply(aFilePattern, aOperateOnFolders, aDoRecurse, FileSetTimeCallback, &callbackData);
}

BOOL FileSetTimeCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData)
{
	HANDLE hFile;
	// Open existing file.
	// FILE_FLAG_NO_BUFFERING might improve performance because all we're doing is
	// changing one of the file's attributes.  FILE_FLAG_BACKUP_SEMANTICS must be
	// used, otherwise changing the time of a directory under NT and beyond will
	// not succeed.  Win95 (not sure about Win98/Me) does not support this, but it
	// should be harmless to specify it even if the OS is Win95:
	hFile = CreateFile(aFilename, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE
		, (LPSECURITY_ATTRIBUTES)NULL, OPEN_EXISTING
		, FILE_FLAG_NO_BUFFERING | FILE_FLAG_BACKUP_SEMANTICS, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
	{
		g->LastError = GetLastError();
		return FALSE;
	}

	BOOL success;
	FileSetTimeData &a = *(FileSetTimeData *)aCallbackData;
	switch (ctoupper(a.WhichTime))
	{
	case 'C': // File's creation time.
		success = SetFileTime(hFile, &a.Time, NULL, NULL);
		break;
	case 'A': // File's last access time.
		success = SetFileTime(hFile, NULL, &a.Time, NULL);
		break;
	default:  // 'M', unspecified, or some other value.  Use the file's modification time.
		success = SetFileTime(hFile, NULL, NULL, &a.Time);
	}
	if (!success)
		g->LastError = GetLastError();
	CloseHandle(hFile);
	return success;
}



ResultType Line::FileGetSize(LPTSTR aFilespec, LPTSTR aGranularity)
{
	OUTPUT_VAR->Assign(); // Init to be blank, in case of failure.

	if (!aFilespec || !*aFilespec)
		return SetErrorsOrThrow(true, ERROR_INVALID_PARAMETER); // Let ErrorLevel indicate an error, since this is probably not what the user intended.
	
	BOOL got_file_size = false;
	__int64 size;

	// Try CreateFile() and GetFileSizeEx() first, since they can be more accurate. 
	// See "Why is the file size reported incorrectly for files that are still being written to?"
	// http://blogs.msdn.com/b/oldnewthing/archive/2011/12/26/10251026.aspx
	HANDLE hfile = CreateFile(aFilespec, FILE_READ_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE
		, NULL, OPEN_EXISTING, 0, NULL);
	if (hfile != INVALID_HANDLE_VALUE)
	{
		got_file_size = GetFileSizeEx(hfile, (PLARGE_INTEGER)&size);
		CloseHandle(hfile);
	}

	if (!got_file_size)
	{
		WIN32_FIND_DATA found_file;
		HANDLE file_search = FindFirstFile(aFilespec, &found_file);
		if (file_search == INVALID_HANDLE_VALUE)
			return SetErrorsOrThrow(true); // Let ErrorLevel tell the story.
		FindClose(file_search);
		size = ((__int64)found_file.nFileSizeHigh << 32) | found_file.nFileSizeLow;
	}

	switch(ctoupper(*aGranularity))
	{
	case 'K': // KB
		size /= 1024;
		break;
	case 'M': // MB
		size /= (1024 * 1024);
		break;
	// default: // i.e. either 'B' for bytes, or blank, or some other unknown value, so default to bytes.
		// do nothing
	}

	SetErrorsOrThrow(false, 0); // Indicate success.
	return OUTPUT_VAR->Assign(size);
	// The below comment is obsolete in light of the switch to 64-bit integers.  But it might
	// be good to keep for background:
	// Currently, the above is basically subject to a 2 gig limit, I believe, after which the
	// size will appear to be negative.  Beyond a 4 gig limit, the value will probably wrap around
	// to zero and start counting from there as file sizes grow beyond 4 gig (UPDATE: The size
	// is now set to -1 [the maximum DWORD when expressed as a signed int] whenever >4 gig).
	// There's not much sense in putting values larger than 2 gig into the var as a text string
	// containing a positive number because such a var could never be properly handled by anything
	// that compares to it (e.g. IfGreater) or does math on it (e.g. EnvAdd), since those operations
	// use ATOI() to convert the string.  So as a future enhancement (unless the whole program is
	// revamped to use 64bit ints or something) might add an optional param to the end to indicate
	// size should be returned in K(ilobyte) or M(egabyte).  However, this is sorta bad too since
	// adding a param can break existing scripts which use filenames containing commas (delimiters)
	// with this command.  Therefore, I think I'll just add the K/M param now.
	// Also, the above assigns an int because unsigned ints should never be stored in script
	// variables.  This is because an unsigned variable larger than INT_MAX would not be properly
	// converted by ATOI(), which is current standard method for variables to be auto-converted
	// from text back to a number whenever that is needed.
}



ResultType Line::SetToggleState(vk_type aVK, ToggleValueType &ForceLock, LPTSTR aToggleText)
// Caller must have already validated that the args are correct.
// Always returns OK, for use as caller's return value.
{
	ToggleValueType toggle = ConvertOnOffAlways(aToggleText, NEUTRAL);
	switch (toggle)
	{
	case TOGGLED_ON:
	case TOGGLED_OFF:
		// Turning it on or off overrides any prior AlwaysOn or AlwaysOff setting.
		// Probably need to change the setting BEFORE attempting to toggle the
		// key state, otherwise the hook may prevent the state from being changed
		// if it was set to be AlwaysOn or AlwaysOff:
		ForceLock = NEUTRAL;
		ToggleKeyState(aVK, toggle);
		break;
	case ALWAYS_ON:
	case ALWAYS_OFF:
		ForceLock = (toggle == ALWAYS_ON) ? TOGGLED_ON : TOGGLED_OFF; // Must do this first.
		ToggleKeyState(aVK, ForceLock);
		// The hook is currently needed to support keeping these keys AlwaysOn or AlwaysOff, though
		// there may be better ways to do it (such as registering them as a hotkey, but
		// that may introduce quite a bit of complexity):
		Hotkey::InstallKeybdHook();
		break;
	case NEUTRAL:
		// Note: No attempt is made to detect whether the keybd hook should be deinstalled
		// because it's no longer needed due to this change.  That would require some 
		// careful thought about the impact on the status variables in the Hotkey class, etc.,
		// so it can be left for a future enhancement:
		ForceLock = NEUTRAL;
		break;
	}
	return OK;
}



////////////////////////////////
// Misc lower level functions //
////////////////////////////////

HWND Line::DetermineTargetWindow(LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	HWND target_window; // A variable of this name is used by the macros below.
	IF_USE_FOREGROUND_WINDOW(g->DetectHiddenWindows, aTitle, aText, aExcludeTitle, aExcludeText)
	else if (*aTitle || *aText || *aExcludeTitle || *aExcludeText)
		target_window = WinExist(*g, aTitle, aText, aExcludeTitle, aExcludeText);
	else // Use the "last found" window.
		target_window = GetValidLastUsedWindow(*g);
	return target_window;
}



bool Line::FileIsFilteredOut(WIN32_FIND_DATA &aCurrentFile, FileLoopModeType aFileLoopMode
	, LPTSTR aFilePath, size_t aFilePathLength)
// Caller has ensured that aFilePath (if non-blank) has a trailing backslash.
{
	if (aCurrentFile.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) // It's a folder.
	{
		if (aFileLoopMode == FILE_LOOP_FILES_ONLY
			|| aCurrentFile.cFileName[0] == '.' && (!aCurrentFile.cFileName[1]      // Relies on short-circuit boolean order.
				|| aCurrentFile.cFileName[1] == '.' && !aCurrentFile.cFileName[2])) //
			return true; // Exclude this folder by returning true.
	}
	else // it's not a folder.
		if (aFileLoopMode == FILE_LOOP_FOLDERS_ONLY)
			return true; // Exclude this file by returning true.

	// Since file was found, also prepend the file's path to its name for the caller:
	if (*aFilePath)
	{
		// Seems best to check length in advance because it allows a faster move/copy method further below
		// (in lieu of sntprintf(), which is probably quite a bit slower than the method here).
		size_t name_length = _tcslen(aCurrentFile.cFileName);
		if (aFilePathLength + name_length >= MAX_PATH)
			// v1.0.45.03: Filter out filenames that would be truncated because it seems undesirable in 99% of
			// cases to include such "faulty" data in the loop.  Most scripts would want to skip them rather than
			// seeing the truncated names.  Furthermore, a truncated name might accidentally match the name
			// of a legitimate non-truncated filename, which could cause such a name to get retrieved twice by
			// the loop (or other undesirable side-effects).
			return true;
		//else no overflow is possible, so below can move things around inside the buffer without concern.
		tmemmove(aCurrentFile.cFileName + aFilePathLength, aCurrentFile.cFileName, name_length + 1); // memmove() because source & dest might overlap.  +1 to include the terminator.
		tmemcpy(aCurrentFile.cFileName, aFilePath, aFilePathLength); // Prepend in the area liberated by the above. Don't include the terminator since this is a concat operation.
	}
	return false; // Indicate that this file is not to be filtered out.
}



Label *Line::GetJumpTarget(bool aIsDereferenced)
{
	LPTSTR target_label = aIsDereferenced ? ARG1 : RAW_ARG1;
	Label *label = g_script.FindLabel(target_label);
	if (!label)
	{
		LineError(ERR_NO_LABEL, FAIL, target_label);
		return NULL;
	}
	if (!aIsDereferenced)
		mRelatedLine = (Line *)label; // The script loader has ensured that label->mJumpToLine isn't NULL.
	// else don't update it, because that would permanently resolve the jump target, and we want it to stay dynamic.
	// Seems best to do this even for GOSUBs even though it's a bit weird:
	return IsJumpValid(*label);
	// Any error msg was already displayed by the above call.
}



Label *Line::IsJumpValid(Label &aTargetLabel, bool aSilent)
// Returns aTargetLabel is the jump is valid, or NULL otherwise.
{
	// aTargetLabel can be NULL if this Goto's target is the physical end of the script.
	// And such a destination is always valid, regardless of where aOrigin is.
	// UPDATE: It's no longer possible for the destination of a Goto or Gosub to be
	// NULL because the script loader has ensured that the end of the script always has
	// an extra ACT_EXIT that serves as an anchor for any final labels in the script:
	//if (aTargetLabel == NULL)
	//	return OK;
	// The above check is also necessary to avoid dereferencing a NULL pointer below.

	Line *parent_line_of_label_line;
	if (   !(parent_line_of_label_line = aTargetLabel.mJumpToLine->mParentLine)   )
		// A Goto/Gosub can always jump to a point anywhere in the outermost layer
		// (i.e. outside all blocks) without restriction:
		return &aTargetLabel; // Indicate success.

	// So now we know this Goto/Gosub is attempting to jump into a block somewhere.  Is that
	// block a legal place to jump?:

	for (Line *ancestor = mParentLine; ancestor != NULL; ancestor = ancestor->mParentLine)
		if (parent_line_of_label_line == ancestor)
			// Since aTargetLabel is in the same block as the Goto line itself (or a block
			// that encloses that block), it's allowed:
			return &aTargetLabel; // Indicate success.
	// This can happen if the Goto's target is at a deeper level than it, or if the target
	// is at a more shallow level but is in some block totally unrelated to it!
	// Returns FAIL by default, which is what we want because that value is zero:
	if (!aSilent)
		LineError(_T("A Goto/Gosub must not jump into a block that doesn't enclose it.")); // Omit GroupActivate from the error msg since that is rare enough to justify the increase in common-case clarity.
	return NULL;
}


BOOL Line::IsOutsideAnyFunctionBody() // v1.0.48.02
{
	for (Line *ancestor = mParentLine; ancestor != NULL; ancestor = ancestor->mParentLine)
		if (ancestor->mAttribute == ATTR_TRUE && ancestor->mActionType == ACT_BLOCK_BEGIN) // Ordered for short-circuit performance.
			return FALSE; // ATTR_TRUE marks an open-brace as belonging to a function's body, so indicate this this line is inside a function.
	return TRUE; // Indicate that this line is not inside any function body.
}


BOOL Line::CheckValidFinallyJump(Line* jumpTarget) // v1.1.14
{
	Line* jumpParent = jumpTarget->mParentLine;
	for (Line *ancestor = mParentLine; ancestor != NULL; ancestor = ancestor->mParentLine)
	{
		if (ancestor == jumpParent)
			return TRUE; // We found the common ancestor.
		if (ancestor->mActionType == ACT_FINALLY)
		{
			LineError(ERR_BAD_JUMP_INSIDE_FINALLY);
			return FALSE; // The common ancestor is outside the FINALLY block and thus this jump is invalid.
		}
	}
	return TRUE; // The common ancestor is the root of the script.
}


////////////////////////
// BUILT-IN VARIABLES //
////////////////////////

VarSizeType BIV_True_False(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
	{
		*aBuf++ = aVarName[4] ? '0': '1';
		*aBuf = '\0';
	}
	return 1; // The length of the value.
}

VarSizeType BIV_MMM_DDD(LPTSTR aBuf, LPTSTR aVarName)
{
	LPTSTR format_str;
	switch(ctoupper(aVarName[2]))
	{
	// Use the case-sensitive formats required by GetDateFormat():
	case 'M': format_str = (aVarName[5] ? _T("MMMM") : _T("MMM")); break;
	case 'D': format_str = (aVarName[5] ? _T("dddd") : _T("ddd")); break;
	}
	// Confirmed: The below will automatically use the local time (not UTC) when 3rd param is NULL.
	return (VarSizeType)(GetDateFormat(LOCALE_USER_DEFAULT, 0, NULL, format_str, aBuf, aBuf ? 999 : 0) - 1);
}

VarSizeType BIV_DateTime(LPTSTR aBuf, LPTSTR aVarName)
{
	if (!aBuf)
		return 6; // Since only an estimate is needed in this mode, return the maximum length of any item.

	aVarName += 2; // Skip past the "A_".

	// The current time is refreshed only if it's been a certain number of milliseconds since
	// the last fetch of one of these built-in time variables.  This keeps the variables in
	// sync with one another when they are used consecutively such as this example:
	// Var = %A_Hour%:%A_Min%:%A_Sec%
	// Using GetTickCount() because it's very low overhead compared to the other time functions:
	static DWORD sLastUpdate = 0; // Static should be thread + recursion safe in this case.
	static SYSTEMTIME sST = {0}; // Init to detect when it's empty.
	BOOL is_msec = !_tcsicmp(aVarName, _T("MSec")); // Always refresh if it's milliseconds, for better accuracy.
	DWORD now_tick = GetTickCount();
	if (is_msec || now_tick - sLastUpdate > 50 || !sST.wYear) // See comments above.
	{
		GetLocalTime(&sST);
		sLastUpdate = now_tick;
	}

	if (is_msec)
		return _stprintf(aBuf, _T("%03d"), sST.wMilliseconds);

	TCHAR second_letter = ctoupper(aVarName[1]);
	switch(ctoupper(aVarName[0]))
	{
	case 'Y':
		switch(second_letter)
		{
		case 'D': // A_YDay
			return _stprintf(aBuf, _T("%d"), GetYDay(sST.wMonth, sST.wDay, IS_LEAP_YEAR(sST.wYear)));
		case 'W': // A_YWeek
			return GetISOWeekNumber(aBuf, sST.wYear
				, GetYDay(sST.wMonth, sST.wDay, IS_LEAP_YEAR(sST.wYear))
				, sST.wDayOfWeek);
		default:  // A_Year/A_YYYY
			return _stprintf(aBuf, _T("%d"), sST.wYear);
		}
		// No break because all cases above return:
		//break;
	case 'M':
		switch(second_letter)
		{
		case 'D': // A_MDay (synonymous with A_DD)
			return _stprintf(aBuf, _T("%02d"), sST.wDay);
		case 'I': // A_Min
			return _stprintf(aBuf, _T("%02d"), sST.wMinute);
		default: // A_MM and A_Mon (A_MSec was already completely handled higher above).
			return _stprintf(aBuf, _T("%02d"), sST.wMonth);
		}
		// No break because all cases above return:
		//break;
	case 'D': // A_DD (synonymous with A_MDay)
		return _stprintf(aBuf, _T("%02d"), sST.wDay);
	case 'W': // A_WDay
		return _stprintf(aBuf, _T("%d"), sST.wDayOfWeek + 1);
	case 'H': // A_Hour
		return _stprintf(aBuf, _T("%02d"), sST.wHour);
	case 'S': // A_Sec (A_MSec was already completely handled higher above).
		return _stprintf(aBuf, _T("%02d"), sST.wSecond);
	}
	return 0; // Never reached, but avoids compiler warning.
}

VarSizeType BIV_BatchLines(LPTSTR aBuf, LPTSTR aVarName)
{
	// The BatchLine value can be either a numerical string or a string that ends in "ms".
	TCHAR buf[256];
	LPTSTR target_buf = aBuf ? aBuf : buf;
	if (g->IntervalBeforeRest > -1) // Have this new method take precedence, if it's in use by the script.
		return _stprintf(target_buf, _T("%dms"), g->IntervalBeforeRest); // Not sntprintf().
	// Otherwise:
	ITOA64(g->LinesPerCycle, target_buf);
	return (VarSizeType)_tcslen(target_buf);
}

BIV_DECL_R(BIV_ListLines)
{
	if (aBuf)
	{
		*aBuf++ = g->ListLinesIsEnabled ? '1' : '0';
		*aBuf = '\0';
	}
	return 1;
}

VarSizeType BIV_TitleMatchMode(LPTSTR aBuf, LPTSTR aVarName)
{
	if (g->TitleMatchMode == FIND_REGEX) // v1.0.45.
	{
		if (aBuf)  // For backward compatibility (due to StringCaseSense), never change the case used here:
			_tcscpy(aBuf, _T("RegEx"));
		return 5; // The length.
	}
	// Otherwise, it's a numerical mode:
	// It's done this way in case it's ever allowed to go beyond a single-digit number.
	TCHAR buf[MAX_INTEGER_SIZE];
	LPTSTR target_buf = aBuf ? aBuf : buf;
	_itot(g->TitleMatchMode, target_buf, 10);  // Always output as decimal vs. hex in this case (so that scripts can use "If var in list" with confidence).
	return (VarSizeType)_tcslen(target_buf);
}

VarSizeType BIV_TitleMatchModeSpeed(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)  // For backward compatibility (due to StringCaseSense), never change the case used here:
		_tcscpy(aBuf, g->TitleFindFast ? _T("Fast") : _T("Slow"));
	return 4;  // Always length 4
}

VarSizeType BIV_DetectHiddenWindows(LPTSTR aBuf, LPTSTR aVarName)
{
	return aBuf
		? (VarSizeType)_tcslen(_tcscpy(aBuf, g->DetectHiddenWindows ? _T("On") : _T("Off"))) // For backward compatibility (due to StringCaseSense), never change the case used here.  Fixed in v1.0.42.01 to return exact length (required).
		: 3; // Room for either On or Off (in the estimation phase).
}

VarSizeType BIV_DetectHiddenText(LPTSTR aBuf, LPTSTR aVarName)
{
	return aBuf
		? (VarSizeType)_tcslen(_tcscpy(aBuf, g->DetectHiddenText ? _T("On") : _T("Off"))) // For backward compatibility (due to StringCaseSense), never change the case used here. Fixed in v1.0.42.01 to return exact length (required).
		: 3; // Room for either On or Off (in the estimation phase).
}

VarSizeType BIV_AutoTrim(LPTSTR aBuf, LPTSTR aVarName)
{
	return aBuf
		? (VarSizeType)_tcslen(_tcscpy(aBuf, g->AutoTrim ? _T("On") : _T("Off"))) // For backward compatibility (due to StringCaseSense), never change the case used here. Fixed in v1.0.42.01 to return exact length (required).
		: 3; // Room for either On or Off (in the estimation phase).
}

VarSizeType BIV_StringCaseSense(LPTSTR aBuf, LPTSTR aVarName)
{
	return aBuf
		? (VarSizeType)_tcslen(_tcscpy(aBuf, g->StringCaseSense == SCS_INSENSITIVE ? _T("Off") // For backward compatibility (due to StringCaseSense), never change the case used here.  Fixed in v1.0.42.01 to return exact length (required).
			: (g->StringCaseSense == SCS_SENSITIVE ? _T("On") : _T("Locale"))))
		: 6; // Room for On, Off, or Locale (in the estimation phase).
}

VarSizeType BIV_FormatInteger(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
	{
		*aBuf++ = g->FormatInt;
		*aBuf = '\0';
	}
	return 1;
}

VarSizeType BIV_FormatFloat(LPTSTR aBuf, LPTSTR aVarName)
{
	if (!aBuf)
		return (VarSizeType)_tcslen(g->FormatFloat);  // Include the extra chars since this is just an estimate.
	LPTSTR str_with_leading_percent_omitted = g->FormatFloat + 1;
	size_t length = _tcslen(str_with_leading_percent_omitted);
	tcslcpy(aBuf, str_with_leading_percent_omitted
		, length + !(length && str_with_leading_percent_omitted[length-1] == 'f')); // Omit the trailing character only if it's an 'f', not any other letter such as the 'e' in "%0.6e" (for backward compatibility).
	return (VarSizeType)_tcslen(aBuf); // Must return exact length when aBuf isn't NULL.
}

VarSizeType BIV_xDelay(LPTSTR aBuf, LPTSTR aVarName)
{
	TCHAR buf[MAX_INTEGER_SIZE];
	LPTSTR target_buf = aBuf ? aBuf : buf;
	global_struct &g = *::g; // Reduces code size.
	int result;
	switch (ctoupper(aVarName[2])) // a_X...
	{
	case 'K':
		if (ctolower(aVarName[6]) == 'e') // a_keydE...
		{
			if (aVarName[10]) // a_keydelayP...
				result = g.KeyDelayPlay;
			else
				result = g.KeyDelay;
		}
		else // a_keydU...
		{
			if (aVarName[13]) // a_keydurationP...
				result = g.PressDurationPlay;
			else
				result = g.PressDuration;
		}
		break;
	case 'M':
		if (aVarName[12]) // a_mousedelayP...
			result = g.MouseDelayPlay;
		else
			result = g.MouseDelay;
		break;
	case 'W':
		result = g.WinDelay;
		break;
	case 'C':
		result = g.ControlDelay;
		break;
	}
	_itot(result, target_buf, 10);  // Always output as decimal vs. hex in this case (so that scripts can use "If var in list" with confidence).
	return (VarSizeType)_tcslen(target_buf);
}

VarSizeType BIV_DefaultMouseSpeed(LPTSTR aBuf, LPTSTR aVarName)
{
	TCHAR buf[MAX_INTEGER_SIZE];
	LPTSTR target_buf = aBuf ? aBuf : buf;
	_itot(g->DefaultMouseSpeed, target_buf, 10);  // Always output as decimal vs. hex in this case (so that scripts can use "If var in list" with confidence).
	return (VarSizeType)_tcslen(target_buf);
}

VarSizeType BIV_CoordMode(LPTSTR aBuf, LPTSTR aVarName)
{
	static LPCTSTR sCoordModes[] = COORD_MODES;
	LPCTSTR result = sCoordModes[(g->CoordMode >> Line::ConvertCoordModeCmd(aVarName + 11)) & COORD_MODE_MASK];
	if (aBuf)
		_tcscpy(aBuf, result);
	return 6; // Currently all are 6 chars.
}

VarSizeType BIV_SendMode(LPTSTR aBuf, LPTSTR aVarName)
{
	static LPCTSTR sSendModes[] = SEND_MODES;
	LPCTSTR result = sSendModes[g->SendMode];
	if (aBuf)
		_tcscpy(aBuf, result);
	return (VarSizeType)_tcslen(result);
}

VarSizeType BIV_SendLevel(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
		return (VarSizeType)_tcslen(_itot(g->SendLevel, aBuf, 10));  // Always output as decimal vs. hex in this case (so that scripts can use "If var in list" with confidence).
	return 3; // Enough room for the maximum SendLevel (100).
}

VarSizeType BIV_StoreCapslockMode(LPTSTR aBuf, LPTSTR aVarName)
{
	return aBuf
		? (VarSizeType)_tcslen(_tcscpy(aBuf, g->StoreCapslockMode ? _T("On") : _T("Off"))) // For backward compatibility (due to StringCaseSense), never change the case used here.
		: 3; // Room for either On or Off (in the estimation phase).
}



VarSizeType BIV_IsPaused(LPTSTR aBuf, LPTSTR aVarName) // v1.0.48: Lexikos: Added BIV_IsPaused and BIV_IsCritical.
{
	// Although A_IsPaused could indicate how many threads are paused beneath the current thread,
	// that would be a problem because it would yield a non-zero value even when the underlying thread
	// isn't paused (i.e. other threads below it are paused), which would defeat the original purpose.
	// In addition, A_IsPaused probably won't be commonly used, so it seems best to keep it simple.
	// NAMING: A_IsPaused seems to be a better name than A_Pause or A_Paused due to:
	//    Better readability.
	//    Consistent with A_IsSuspended, which is strongly related to pause/unpause.
	//    The fact that it wouldn't be likely for a function to turn off pause then turn it back on
	//      (or vice versa), which was the main reason for storing "Off" and "On" in things like
	//      A_DetectHiddenWindows.
	if (aBuf)
	{
		// Checking g>g_array avoids any chance of underflow, which might otherwise happen if this is
		// called by the AutoExec section or a threadless callback running in thread #0.
		*aBuf++ = (g > g_array && g[-1].IsPaused) ? '1' : '0';
		*aBuf = '\0';
	}
	return 1;
}

VarSizeType BIV_IsCritical(LPTSTR aBuf, LPTSTR aVarName) // v1.0.48: Lexikos: Added BIV_IsPaused and BIV_IsCritical.
{
	if (!aBuf) // Return conservative estimate in case Critical status can ever change between the 1st and 2nd calls to this function.
		return MAX_INTEGER_LENGTH;
	// It seems more useful to return g->PeekFrequency than "On" or "Off" (ACT_CRITICAL ensures that
	// g->PeekFrequency!=0 whenever g->ThreadIsCritical==true).  Also, the word "Is" in "A_IsCritical"
	// implies a value that can be used as a boolean such as "if A_IsCritical".
	if (g->ThreadIsCritical)
		return (VarSizeType)_tcslen(UTOA(g->PeekFrequency, aBuf)); // ACT_CRITICAL ensures that g->PeekFrequency > 0 when critical is on.
	// Otherwise:
	*aBuf++ = '0';
	*aBuf = '\0';
	return 1; // Caller might rely on receiving actual length when aBuf!=NULL.
}

VarSizeType BIV_IsSuspended(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
	{
		*aBuf++ = g_IsSuspended ? '1' : '0';
		*aBuf = '\0';
	}
	return 1;
}



VarSizeType BIV_IsCompiled(LPTSTR aBuf, LPTSTR aVarName)
{
#ifdef AUTOHOTKEYSC
	if (aBuf)
	{
		*aBuf++ = '1';
		*aBuf = '\0';
	}
	return 1;
#else
	// v1.1.06: A_IsCompiled is defined so that it does not cause warnings with #Warn enabled,
	// but left empty for backward-compatibility.  Defining the variable (even though it is
	// left empty) has some side-effects:
	//
	//  1) Uncompiled scripts can no longer assign a value to A_IsCompiled.  Even if a script
	//     assigned A_IsCompiled:=0 to make certain calculations easier, it would have to be
	//     done dynamically to avoid a load-time error if the script is compiled.  So this
	//     seems unlikely to be a problem.
	//
	//  2) Address-of will return an empty string instead of the address of a global variable
	//     (or the address of Var::sEmptyString if the variable hasn't been given a value).
	//
	//  3) A_IsCompiled will never show up in ListVars, even if the script is uncompiled.
	//     
	if (aBuf)
		*aBuf = '\0';
	return 0;
#endif
}



VarSizeType BIV_IsUnicode(LPTSTR aBuf, LPTSTR aVarName)
{
#ifdef UNICODE
	if (aBuf)
	{
		*aBuf++ = '1';
		*aBuf = '\0';
	}
	return 1;
#else
	// v1.1.06: A_IsUnicode is defined so that it does not cause warnings with #Warn enabled,
	// but left empty to encourage compatibility with older versions and AutoHotkey Basic.
	// This prevents scripts from using expressions like A_IsUnicode+1, which would succeed
	// if A_IsUnicode is 0 or 1 but fail if it is "".  This change has side-effects similar
	// to those described for A_IsCompiled above.
	if (aBuf)
		*aBuf = '\0';
	return 0;
#endif
}



VarSizeType BIV_FileEncoding(LPTSTR aBuf, LPTSTR aVarName)
{
	// A similar section may be found under "case Encoding:" in FileObject::Invoke.  Maintain that with this:
	switch (g->Encoding)
	{
	case CP_ACP:
		if (aBuf)
			*aBuf = '\0';
		return 0;
#define FILEENCODING_CASE(n, s) \
	case n: \
		if (aBuf) \
			_tcscpy(aBuf, _T(s)); \
		return _countof(_T(s)) - 1;
	// Returning readable strings for these seems more useful than returning their numeric values, especially with CP_AHKNOBOM:
	FILEENCODING_CASE(CP_UTF8, "UTF-8")
	FILEENCODING_CASE(CP_UTF8 | CP_AHKNOBOM, "UTF-8-RAW")
	FILEENCODING_CASE(CP_UTF16, "UTF-16")
	FILEENCODING_CASE(CP_UTF16 | CP_AHKNOBOM, "UTF-16-RAW")
#undef FILEENCODING_CASE
	default:
	  {
		TCHAR buf[MAX_INTEGER_SIZE + 2]; // + 2 for "CP"
		LPTSTR target_buf = aBuf ? aBuf : buf;
		target_buf[0] = _T('C');
		target_buf[1] = _T('P');
		_itot(g->Encoding, target_buf + 2, 10);  // Always output as decimal since we aren't exactly returning a number.
		return (VarSizeType)_tcslen(target_buf);
	  }
	}
}



VarSizeType BIV_RegView(LPTSTR aBuf, LPTSTR aVarName)
{
	LPCTSTR value;
	switch (g->RegView)
	{
	case KEY_WOW64_32KEY: value = _T("32"); break;
	case KEY_WOW64_64KEY: value = _T("64"); break;
	default: value = _T("Default"); break;
	}
	if (aBuf)
		_tcscpy(aBuf, value);
	return (VarSizeType)_tcslen(value);
}



VarSizeType BIV_LastError(LPTSTR aBuf, LPTSTR aVarName)
{
	TCHAR buf[MAX_INTEGER_SIZE];
	LPTSTR target_buf = aBuf ? aBuf : buf;
	_itot(g->LastError, target_buf, 10);  // Always output as decimal vs. hex in this case (so that scripts can use "If var in list" with confidence).
	return (VarSizeType)_tcslen(target_buf);
}



VarSizeType BIV_PtrSize(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
	{
		// Return size in bytes of a pointer in the current build.
		*aBuf++ = '0' + sizeof(void *);
		*aBuf = '\0';
	}
	return 1;
}



VarSizeType BIV_ScreenDPI(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
		_itot(g_ScreenDPI, aBuf, 10);
	return aBuf ? (VarSizeType)_tcslen(aBuf) : MAX_INTEGER_SIZE;
}



VarSizeType BIV_IconHidden(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
	{
		*aBuf++ = g_NoTrayIcon ? '1' : '0';
		*aBuf = '\0';
	}
	return 1;  // Length is always 1.
}

VarSizeType BIV_IconTip(LPTSTR aBuf, LPTSTR aVarName)
{
	if (!aBuf)
		return g_script.mTrayIconTip ? (VarSizeType)_tcslen(g_script.mTrayIconTip) : 0;
	if (g_script.mTrayIconTip)
		return (VarSizeType)_tcslen(_tcscpy(aBuf, g_script.mTrayIconTip));
	else
	{
		*aBuf = '\0';
		return 0;
	}
}

VarSizeType BIV_IconFile(LPTSTR aBuf, LPTSTR aVarName)
{
	if (!aBuf)
		return g_script.mCustomIconFile ? (VarSizeType)_tcslen(g_script.mCustomIconFile) : 0;
	if (g_script.mCustomIconFile)
		return (VarSizeType)_tcslen(_tcscpy(aBuf, g_script.mCustomIconFile));
	else
	{
		*aBuf = '\0';
		return 0;
	}
}

VarSizeType BIV_IconNumber(LPTSTR aBuf, LPTSTR aVarName)
{
	TCHAR buf[MAX_INTEGER_SIZE];
	LPTSTR target_buf = aBuf ? aBuf : buf;
	if (!g_script.mCustomIconNumber) // Yield an empty string rather than the digit "0".
	{
		*target_buf = '\0';
		return 0;
	}
	return (VarSizeType)_tcslen(UTOA(g_script.mCustomIconNumber, target_buf));
}



VarSizeType BIV_PriorKey(LPTSTR aBuf, LPTSTR aVarName)
{
	const int bufSize = 32;
	if (!aBuf)
		return bufSize;

	*aBuf = '\0'; // Init for error & not-found cases

	int validEventCount = 0;
	// Start at the current event (offset 1)
	for (int iOffset = 1; iOffset <= g_MaxHistoryKeys; ++iOffset)
	{
		// Get index for circular buffer
		int i = (g_KeyHistoryNext + g_MaxHistoryKeys - iOffset) % g_MaxHistoryKeys;
		// Keep looking until we hit the second valid event
		if (g_KeyHistory[i].event_type != _T('i') && ++validEventCount > 1)
		{
			// Find the next most recent key-down
			if (!g_KeyHistory[i].key_up)
			{
				GetKeyName(g_KeyHistory[i].vk, g_KeyHistory[i].sc, aBuf, bufSize);
				break;
			}
		}
	}
	return (VarSizeType)_tcslen(aBuf);
}



LPTSTR GetExitReasonString(ExitReasons aExitReason)
{
	LPTSTR str;
	switch(g_script.mExitReason)
	{
	case EXIT_LOGOFF: str = _T("Logoff"); break;
	case EXIT_SHUTDOWN: str = _T("Shutdown"); break;
	// Since the below are all relatively rare, except WM_CLOSE perhaps, they are all included
	// as one word to cut down on the number of possible words (it's easier to write OnExit
	// routines to cover all possibilities if there are fewer of them).
	// Update: The redundant ExitReasons were merged to reduce code size.
	case EXIT_DESTROY:
	case EXIT_CLOSE: str = _T("Close"); break;
	case EXIT_ERROR: str = _T("Error"); break;
	case EXIT_MENU: str = _T("Menu"); break;  // Standard menu, not a user-defined menu.
	case EXIT_EXIT: str = _T("Exit"); break;  // ExitApp or Exit command.
	case EXIT_RELOAD: str = _T("Reload"); break;
	case EXIT_SINGLEINSTANCE: str = _T("Single"); break;
	default:  // EXIT_NONE or unknown value (unknown would be considered a bug if it ever happened).
		str = _T("");
	}
	return str;
}

VarSizeType BIV_ExitReason(LPTSTR aBuf, LPTSTR aVarName)
{
	LPTSTR str = GetExitReasonString(g_script.mExitReason);
	if (aBuf)
		_tcscpy(aBuf, str);
	return (VarSizeType)_tcslen(str);
}



VarSizeType BIV_Space_Tab(LPTSTR aBuf, LPTSTR aVarName)
{
	// Really old comment:
	// A_Space is a built-in variable rather than using an escape sequence such as `s, because the escape
	// sequence method doesn't work (probably because `s resolves to a space and is trimmed at some point
	// prior to when it can be used):
	if (aBuf)
	{
		*aBuf++ = aVarName[5] ? ' ' : '\t'; // A_Tab[]
		*aBuf = '\0';
	}
	return 1;
}

VarSizeType BIV_AhkVersion(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
		_tcscpy(aBuf, T_AHK_VERSION);
	return (VarSizeType)_tcslen(T_AHK_VERSION);
}

VarSizeType BIV_AhkPath(LPTSTR aBuf, LPTSTR aVarName) // v1.0.41.
{
#ifdef AUTOHOTKEYSC
	if (aBuf)
	{
		size_t length;
		if (length = GetAHKInstallDir(aBuf))
			// Name "AutoHotkey.exe" is assumed for code size reduction and because it's not stored in the registry:
			tcslcpy(aBuf + length, _T("\\AutoHotkey.exe"), MAX_PATH - length); // strlcpy() in case registry has a path that is too close to MAX_PATH to fit AutoHotkey.exe
		//else leave it blank as documented.
		return (VarSizeType)_tcslen(aBuf);
	}
	// Otherwise: Always return an estimate of MAX_PATH in case the registry entry changes between the
	// first call and the second.  This is also relied upon by strlcpy() above, which zero-fills the tail
	// of the destination up through the limit of its capacity (due to calling strncpy, which does this).
	return MAX_PATH;
#else
	TCHAR buf[MAX_PATH];
	VarSizeType length = (VarSizeType)GetModuleFileName(NULL, buf, MAX_PATH);
	if (aBuf)
		_tcscpy(aBuf, buf); // v1.0.47: Must be done as a separate copy because passing a size of MAX_PATH for aBuf can crash when aBuf is actually smaller than that (even though it's large enough to hold the string). This is true for ReadRegString()'s API call and may be true for other API calls like this one.
	return length;
#endif
}



VarSizeType BIV_TickCount(LPTSTR aBuf, LPTSTR aVarName)
{
	return aBuf
		? (VarSizeType)_tcslen(ITOA64(GetTickCount(), aBuf))
		: MAX_INTEGER_LENGTH; // IMPORTANT: Conservative estimate because tick might change between 1st & 2nd calls.
}



VarSizeType BIV_Now(LPTSTR aBuf, LPTSTR aVarName)
{
	if (!aBuf)
		return DATE_FORMAT_LENGTH;
	SYSTEMTIME st;
	if (aVarName[5]) // A_Now[U]TC
		GetSystemTime(&st);
	else
		GetLocalTime(&st);
	SystemTimeToYYYYMMDD(aBuf, st);
	return (VarSizeType)_tcslen(aBuf);
}

VarSizeType BIV_OSType(LPTSTR aBuf, LPTSTR aVarName)
{
	LPTSTR type = g_os.IsWinNT() ? _T("WIN32_NT") : _T("WIN32_WINDOWS");
	if (aBuf)
		_tcscpy(aBuf, type);
	return (VarSizeType)_tcslen(type); // Return length of type, not aBuf.
}

VarSizeType BIV_OSVersion(LPTSTR aBuf, LPTSTR aVarName)
{
	LPCTSTR version = g_os.Version();  // Init for new or unrecognized OSes.
	if (g_os.IsWinNT()) // "NT" includes all NT-kernel OSes: NT4/2000/XP/2003/Vista/7/8/etc.
	{
		if (g_os.IsWinXP())
			version = _T("WIN_XP");
		else if (g_os.IsWin7())
			version = _T("WIN_7");
		else if (g_os.IsWin8_1())
			version = _T("WIN_8.1");
		else if (g_os.IsWin8())
			version = _T("WIN_8");
		else if (g_os.IsWinVista())
			version = _T("WIN_VISTA");
		else if (g_os.IsWin2003())
			version = _T("WIN_2003");
		else
		{
			if (g_os.IsWin2000())
				version = _T("WIN_2000");
			else if (g_os.IsWinNT4())
				version = _T("WIN_NT4");
		}
	}
	else
	{
		if (g_os.IsWin95())
			version = _T("WIN_95");
		else
		{
			if (g_os.IsWin98())
				version = _T("WIN_98");
			else
				version = _T("WIN_ME");
		}
	}
	if (aBuf)
		_tcscpy(aBuf, version);
	return (VarSizeType)_tcslen(version); // Always return the length of version, not aBuf.
}

VarSizeType BIV_Is64bitOS(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
	{
		*aBuf++ = IsOS64Bit() ? '1' : '0';
		*aBuf = '\0';
	}
	return 1;
}

VarSizeType BIV_Language(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
		_stprintf(aBuf, _T("%04hX"), GetSystemDefaultUILanguage());
	return 4;
}

VarSizeType BIV_UserName_ComputerName(LPTSTR aBuf, LPTSTR aVarName)
{
	TCHAR buf[MAX_PATH];  // Doesn't use MAX_COMPUTERNAME_LENGTH + 1 in case longer names are allowed in the future.
	DWORD buf_size = MAX_PATH; // Below: A_Computer[N]ame (N is the 11th char, index 10, which if present at all distinguishes between the two).
	if (   !(aVarName[10] ? GetComputerName(buf, &buf_size) : GetUserName(buf, &buf_size))   )
		*buf = '\0';
	if (aBuf)
		_tcscpy(aBuf, buf); // v1.0.47: Must be done as a separate copy because passing a size of MAX_PATH for aBuf can crash when aBuf is actually smaller than that (even though it's large enough to hold the string). This is true for ReadRegString()'s API call and may be true for other API calls like the ones here.
	return (VarSizeType)_tcslen(buf); // I seem to remember that the lengths returned from the above API calls aren't consistent in these cases.
}

VarSizeType BIV_WorkingDir(LPTSTR aBuf, LPTSTR aVarName)
{
	// Use GetCurrentDirectory() vs. g_WorkingDir because any in-progress FileSelectFile()
	// dialog is able to keep functioning even when it's quasi-thread is suspended.  The
	// dialog can thus change the current directory as seen by the active quasi-thread even
	// though g_WorkingDir hasn't been updated.  It might also be possible for the working
	// directory to change in unusual circumstances such as a network drive being lost).
	//
	// Fix for v1.0.43.11: Changed size below from 9999 to MAX_PATH, otherwise it fails sometimes on Win9x.
	// Testing shows that the failure is not caused by GetCurrentDirectory() writing to the unused part of the
	// buffer, such as zeroing it (which is good because that would require this part to be redesigned to pass
	// the actual buffer size or use a temp buffer).  So there's something else going on to explain why the
	// problem only occurs in longer scripts on Win98se, not in trivial ones such as Var=%A_WorkingDir%.
	// Nor did the problem affect expression assignments such as Var:=A_WorkingDir.
	TCHAR buf[MAX_PATH];
	VarSizeType length = GetCurrentDirectory(MAX_PATH, buf);
	if (aBuf)
		_tcscpy(aBuf, buf); // v1.0.47: Must be done as a separate copy because passing a size of MAX_PATH for aBuf can crash when aBuf is actually smaller than that (even though it's large enough to hold the string). This is true for ReadRegString()'s API call and may be true for other API calls like the one here.
	return length;
	// Formerly the following, but I don't think it's as reliable/future-proof given the 1.0.47 comment above:
	//return aBuf
	//	? GetCurrentDirectory(MAX_PATH, aBuf)
	//	: GetCurrentDirectory(0, NULL); // MSDN says that this is a valid way to call it on all OSes, and testing shows that it works on WinXP and 98se.
		// Above avoids subtracting 1 to be conservative and to reduce code size (due to the need to otherwise check for zero and avoid subtracting 1 in that case).
}

VarSizeType BIV_WinDir(LPTSTR aBuf, LPTSTR aVarName)
{
	TCHAR buf[MAX_PATH];
	VarSizeType length = GetWindowsDirectory(buf, MAX_PATH);
	if (aBuf)
		_tcscpy(aBuf, buf); // v1.0.47: Must be done as a separate copy because passing a size of MAX_PATH for aBuf can crash when aBuf is actually smaller than that (even though it's large enough to hold the string). This is true for ReadRegString()'s API call and may be true for other API calls like the one here.
	return length;
	// Formerly the following, but I don't think it's as reliable/future-proof given the 1.0.47 comment above:
	//TCHAR buf_temp[1]; // Just a fake buffer to pass to some API functions in lieu of a NULL, to avoid any chance of misbehavior. Keep the size at 1 so that API functions will always fail to copy to buf.
	//// Sizes/lengths/-1/return-values/etc. have been verified correct.
	//return aBuf
	//	? GetWindowsDirectory(aBuf, MAX_PATH) // MAX_PATH is kept in case it's needed on Win9x for reasons similar to those in GetEnvironmentVarWin9x().
	//	: GetWindowsDirectory(buf_temp, 0);
		// Above avoids subtracting 1 to be conservative and to reduce code size (due to the need to otherwise check for zero and avoid subtracting 1 in that case).
}

VarSizeType BIV_Temp(LPTSTR aBuf, LPTSTR aVarName)
{
	TCHAR buf[MAX_PATH];
	VarSizeType length = GetTempPath(MAX_PATH, buf);
	if (aBuf)
	{
		_tcscpy(aBuf, buf); // v1.0.47: Must be done as a separate copy because passing a size of MAX_PATH for aBuf can crash when aBuf is actually smaller than that (even though it's large enough to hold the string). This is true for ReadRegString()'s API call and may be true for other API calls like the one here.
		if (length)
		{
			aBuf += length - 1;
			if (*aBuf == '\\') // For some reason, it typically yields a trailing backslash, so omit it to improve friendliness/consistency.
			{
				*aBuf = '\0';
				--length;
			}
		}
	}
	return length;
}

VarSizeType BIV_ComSpec(LPTSTR aBuf, LPTSTR aVarName)
{
	TCHAR buf_temp[1]; // Just a fake buffer to pass to some API functions in lieu of a NULL, to avoid any chance of misbehavior. Keep the size at 1 so that API functions will always fail to copy to buf.
	// Sizes/lengths/-1/return-values/etc. have been verified correct.
	return aBuf ? GetEnvVarReliable(_T("comspec"), aBuf) // v1.0.46.08: GetEnvVarReliable() fixes %Comspec% on Windows 9x.
		: GetEnvironmentVariable(_T("comspec"), buf_temp, 0); // Avoids subtracting 1 to be conservative and to reduce code size (due to the need to otherwise check for zero and avoid subtracting 1 in that case).
}

VarSizeType BIV_SpecialFolderPath(LPTSTR aBuf, LPTSTR aVarName)
{
	TCHAR buf[MAX_PATH]; // One caller relies on this being explicitly limited to MAX_PATH.
	int aFolder;
	switch (ctoupper(aVarName[2]))
	{
	case 'P': // A_[P]rogram...
	case 'O': // Pr[o]gramFiles
		if (ctoupper(aVarName[9]) == 'S') // A_Programs(Common)
			aFolder = aVarName[10] ? CSIDL_COMMON_PROGRAMS : CSIDL_PROGRAMS;
		else // A_Program[F]iles or ProgramFi[L]es
			aFolder = CSIDL_PROGRAM_FILES;
		break;
	case 'A': // A_AppData(Common)
		aFolder = aVarName[9] ? CSIDL_COMMON_APPDATA : CSIDL_APPDATA;
		break;
	case 'D': // A_Desktop(Common)
		aFolder = aVarName[9] ? CSIDL_COMMON_DESKTOPDIRECTORY : CSIDL_DESKTOPDIRECTORY;
		break;
	case 'S':
		if (ctoupper(aVarName[7]) == 'M') // A_Start[M]enu(Common)
			aFolder = aVarName[11] ? CSIDL_COMMON_STARTMENU : CSIDL_STARTMENU;
		else // A_Startup(Common)
			aFolder = aVarName[9] ? CSIDL_COMMON_STARTUP : CSIDL_STARTUP;
		break;
#ifdef _DEBUG
	default:
		MsgBox(_T("DEBUG: Unhandled SpecialFolderPath variable."));
#endif
	}
	if (SHGetFolderPath(NULL, aFolder, NULL, SHGFP_TYPE_CURRENT, buf) != S_OK)
		*buf = '\0';
	if (aBuf)
		_tcscpy(aBuf, buf); // Must be done as a separate copy because SHGetFolderPath requires a buffer of length MAX_PATH, and aBuf is usually smaller.
	return _tcslen(buf);
}

VarSizeType BIV_MyDocuments(LPTSTR aBuf, LPTSTR aVarName) // Called by multiple callers.
{
	TCHAR buf[MAX_PATH];
	if (SHGetFolderPath(NULL, CSIDL_MYDOCUMENTS, NULL, SHGFP_TYPE_CURRENT, buf) != S_OK)
		*buf = '\0';
	// Since it is common (such as in networked environments) to have My Documents on the root of a drive
	// (such as a mapped drive letter), remove the backslash from something like M:\ because M: is more
	// appropriate for most uses:
	VarSizeType length = (VarSizeType)strip_trailing_backslash(buf);
	if (aBuf)
		_tcscpy(aBuf, buf); // v1.0.47: Must be done as a separate copy because passing a size of MAX_PATH for aBuf can crash when aBuf is actually smaller than that (even though it's large enough to hold the string).
	return length;
}



VarSizeType BIV_Caret(LPTSTR aBuf, LPTSTR aVarName)
{
	if (!aBuf)
		return MAX_INTEGER_LENGTH; // Conservative, both for performance and in case the value changes between first and second call.

	// These static variables are used to keep the X and Y coordinates in sync with each other, as a snapshot
	// of where the caret was at one precise instant in time.  This is because the X and Y vars are resolved
	// separately by the script, and due to split second timing, they might otherwise not be accurate with
	// respect to each other.  This method also helps performance since it avoids unnecessary calls to
	// GetGUIThreadInfo().
	static HWND sForeWinPrev = NULL;
	static DWORD sTimestamp = GetTickCount();
	static POINT sPoint;
	static BOOL sResult;

	// I believe only the foreground window can have a caret position due to relationship with focused control.
	HWND target_window = GetForegroundWindow(); // Variable must be named target_window for ATTACH_THREAD_INPUT.
	if (!target_window) // No window is in the foreground, report blank coordinate.
	{
		*aBuf = '\0';
		return 0;
	}

	DWORD now_tick = GetTickCount();

	if (target_window != sForeWinPrev || now_tick - sTimestamp > 5) // Different window or too much time has passed.
	{
		// Otherwise:
		GUITHREADINFO info;
		info.cbSize = sizeof(GUITHREADINFO);
		sResult = GetGUIThreadInfo(GetWindowThreadProcessId(target_window, NULL), &info) // Got info okay...
			&& info.hwndCaret; // ...and there is a caret.
		if (!sResult)
		{
			*aBuf = '\0';
			return 0;
		}
		sPoint.x = info.rcCaret.left;
		sPoint.y = info.rcCaret.top;
		// Unconditionally convert to screen coordinates, for simplicity.
		ClientToScreen(info.hwndCaret, &sPoint);
		// Now convert back to whatever is expected for the current mode.
		POINT origin = {0};
		CoordToScreen(origin, COORD_MODE_CARET);
		sPoint.x -= origin.x;
		sPoint.y -= origin.y;
		// Now that all failure conditions have been checked, update static variables for the next caller:
		sForeWinPrev = target_window;
		sTimestamp = now_tick;
	}
	else // Same window and recent enough, but did prior call fail?  If so, provide a blank result like the prior.
	{
		if (!sResult)
		{
			*aBuf = '\0';
			return 0;
		}
	}
	// Now the above has ensured that sPoint contains valid coordinates that are up-to-date enough to be used.
	_itot(ctoupper(aVarName[7]) == 'X' ? sPoint.x : sPoint.y, aBuf, 10);  // Always output as decimal vs. hex in this case (so that scripts can use "If var in list" with confidence).
	return (VarSizeType)_tcslen(aBuf);
}



VarSizeType BIV_Cursor(LPTSTR aBuf, LPTSTR aVarName)
{
	if (!aBuf)
		return SMALL_STRING_LENGTH;  // We're returning the length of the var's contents, not the size.

	// Must fetch it at runtime, otherwise the program can't even be launched on Windows 95:
	typedef BOOL (WINAPI *MyGetCursorInfoType)(PCURSORINFO);
	static MyGetCursorInfoType MyGetCursorInfo = (MyGetCursorInfoType)GetProcAddress(GetModuleHandle(_T("user32")), "GetCursorInfo");

	HCURSOR current_cursor;
	if (MyGetCursorInfo) // v1.0.42.02: This method is used to avoid ATTACH_THREAD_INPUT, which interferes with double-clicking if called repeatedly at a high frequency.
	{
		CURSORINFO ci;
		ci.cbSize = sizeof(CURSORINFO);
		current_cursor = MyGetCursorInfo(&ci) ? ci.hCursor : NULL;
	}
	else // Windows 95 and old-service-pack versions of NT4 require the old method.
	{
		POINT point;
		GetCursorPos(&point);
		HWND target_window = WindowFromPoint(point);

		// MSDN docs imply that threads must be attached for GetCursor() to work.
		// A side-effect of attaching threads or of GetCursor() itself is that mouse double-clicks
		// are interfered with, at least if this function is called repeatedly at a high frequency.
		ATTACH_THREAD_INPUT
		current_cursor = GetCursor();
		DETACH_THREAD_INPUT
	}

	if (!current_cursor)
	{
		#define CURSOR_UNKNOWN _T("Unknown")
		tcslcpy(aBuf, CURSOR_UNKNOWN, SMALL_STRING_LENGTH + 1);
		return (VarSizeType)_tcslen(aBuf);
	}

	// Static so that it's initialized on first use (should help performance after the first time):
	static HCURSOR sCursor[] = {LoadCursor(NULL, IDC_APPSTARTING), LoadCursor(NULL, IDC_ARROW)
		, LoadCursor(NULL, IDC_CROSS), LoadCursor(NULL, IDC_HELP), LoadCursor(NULL, IDC_IBEAM)
		, LoadCursor(NULL, IDC_ICON), LoadCursor(NULL, IDC_NO), LoadCursor(NULL, IDC_SIZE)
		, LoadCursor(NULL, IDC_SIZEALL), LoadCursor(NULL, IDC_SIZENESW), LoadCursor(NULL, IDC_SIZENS)
		, LoadCursor(NULL, IDC_SIZENWSE), LoadCursor(NULL, IDC_SIZEWE), LoadCursor(NULL, IDC_UPARROW)
		, LoadCursor(NULL, IDC_WAIT)}; // If IDC_HAND were added, it would break existing scripts that rely on Unknown being synonymous with Hand.  If ever added, IDC_HAND should return NULL on Win95/NT.
	// The order in the below array must correspond to the order in the above array:
	static LPTSTR sCursorName[] = {_T("AppStarting"), _T("Arrow")
		, _T("Cross"), _T("Help"), _T("IBeam")
		, _T("Icon"), _T("No"), _T("Size")
		, _T("SizeAll"), _T("SizeNESW"), _T("SizeNS")  // NESW = NorthEast+SouthWest
		, _T("SizeNWSE"), _T("SizeWE"), _T("UpArrow")
		, _T("Wait"), CURSOR_UNKNOWN};  // The last item is used to mark end-of-array.
	static const size_t cursor_count = _countof(sCursor);

	int i;
	for (i = 0; i < cursor_count; ++i)
		if (sCursor[i] == current_cursor)
			break;

	tcslcpy(aBuf, sCursorName[i], SMALL_STRING_LENGTH + 1);  // If a is out-of-bounds, "Unknown" will be used.
	return (VarSizeType)_tcslen(aBuf);
}

VarSizeType BIV_ScreenWidth_Height(LPTSTR aBuf, LPTSTR aVarName)
{
	return aBuf
		? (VarSizeType)_tcslen(ITOA(GetSystemMetrics(aVarName[13] ? SM_CYSCREEN : SM_CXSCREEN), aBuf))
		: MAX_INTEGER_LENGTH;
}

VarSizeType BIV_ScriptName(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
		_tcscpy(aBuf, g_script.mFileName);
	return (VarSizeType)_tcslen(g_script.mFileName);
}

VarSizeType BIV_ScriptDir(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
		_tcscpy(aBuf, g_script.mFileDir);
	return _tcslen(g_script.mFileDir);
}

VarSizeType BIV_ScriptFullPath(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
		_tcscpy(aBuf, g_script.mFileSpec);
	return _tcslen(g_script.mFileSpec);
}

VarSizeType BIV_ScriptHwnd(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
		return (VarSizeType)_tcslen(HwndToString(g_hWnd, aBuf));
	return MAX_INTEGER_LENGTH;
}


LineNumberType Script::CurrentLine()
{
	return mCurrLine ? mCurrLine->mLineNumber : mCombinedLineNumber;
}

VarSizeType BIV_LineNumber(LPTSTR aBuf, LPTSTR aVarName)
// Caller has ensured that g_script.mCurrLine is not NULL.
{
	return aBuf
		? (VarSizeType)_tcslen(ITOA(g_script.CurrentLine(), aBuf))
		: MAX_INTEGER_LENGTH;
}


LPTSTR Script::CurrentFile()
{
	return Line::sSourceFile[mCurrLine ? mCurrLine->mFileIndex : mCurrFileIndex];
}

VarSizeType BIV_LineFile(LPTSTR aBuf, LPTSTR aVarName)
// Caller has ensured that g_script.mCurrLine is not NULL.
{
	if (aBuf)
		_tcscpy(aBuf, g_script.CurrentFile());
	return (VarSizeType)_tcslen(g_script.CurrentFile());
}



VarSizeType BIV_LoopFileName(LPTSTR aBuf, LPTSTR aVarName) // Called by multiple callers.
{
	LPTSTR naked_filename;
	if (g->mLoopFile)
	{
		// The loop handler already prepended the script's directory in here for us:
		if (naked_filename = _tcsrchr(g->mLoopFile->cFileName, '\\'))
			++naked_filename;
		else // No backslash, so just make it the entire file name.
			naked_filename = g->mLoopFile->cFileName;
	}
	else
		naked_filename = _T("");
	if (aBuf)
		_tcscpy(aBuf, naked_filename);
	return (VarSizeType)_tcslen(naked_filename);
}

VarSizeType BIV_LoopFileShortName(LPTSTR aBuf, LPTSTR aVarName)
{
	LPTSTR short_filename = _T("");  // Set default.
	if (g->mLoopFile)
	{
		if (   !*(short_filename = g->mLoopFile->cAlternateFileName)   )
			// Files whose long name is shorter than the 8.3 usually don't have value stored here,
			// so use the long name whenever a short name is unavailable for any reason (could
			// also happen if NTFS has short-name generation disabled?)
			return BIV_LoopFileName(aBuf, _T(""));
	}
	if (aBuf)
		_tcscpy(aBuf, short_filename);
	return (VarSizeType)_tcslen(short_filename);
}

VarSizeType BIV_LoopFileExt(LPTSTR aBuf, LPTSTR aVarName)
{
	LPTSTR file_ext = _T("");  // Set default.
	if (g->mLoopFile)
	{
		// The loop handler already prepended the script's directory in here for us:
		if (file_ext = _tcsrchr(g->mLoopFile->cFileName, '.'))
		{
			++file_ext;
			if (_tcschr(file_ext, '\\')) // v1.0.48.01: Disqualify periods found in the path instead of the filename; e.g. path.name\FileWithNoExtension.
				file_ext = _T("");
		}
		else // Reset to empty string vs. NULL.
			file_ext = _T("");
	}
	if (aBuf)
		_tcscpy(aBuf, file_ext);
	return (VarSizeType)_tcslen(file_ext);
}

VarSizeType BIV_LoopFileDir(LPTSTR aBuf, LPTSTR aVarName)
{
	LPTSTR file_dir = _T("");  // Set default.
	LPTSTR last_backslash = NULL;
	if (g->mLoopFile)
	{
		// The loop handler already prepended the script's directory in here for us.
		// But if the loop had a relative path in its FilePattern, there might be
		// only a relative directory here, or no directory at all if the current
		// file is in the origin/root dir of the search:
		if (last_backslash = _tcsrchr(g->mLoopFile->cFileName, '\\'))
		{
			*last_backslash = '\0'; // Temporarily terminate.
			file_dir = g->mLoopFile->cFileName;
		}
		else // No backslash, so there is no directory in this case.
			file_dir = _T("");
	}
	VarSizeType length = (VarSizeType)_tcslen(file_dir);
	if (!aBuf)
	{
		if (last_backslash)
			*last_backslash = '\\';  // Restore the original value.
		return length;
	}
	_tcscpy(aBuf, file_dir);
	if (last_backslash)
		*last_backslash = '\\';  // Restore the original value.
	return length;
}

VarSizeType BIV_LoopFileFullPath(LPTSTR aBuf, LPTSTR aVarName)
{
	// The loop handler already prepended the script's directory in cFileName for us:
	LPTSTR full_path = g->mLoopFile ? g->mLoopFile->cFileName : _T("");
	if (aBuf)
		_tcscpy(aBuf, full_path);
	return (VarSizeType)_tcslen(full_path);
}

VarSizeType BIV_LoopFileLongPath(LPTSTR aBuf, LPTSTR aVarName)
{
	TCHAR *unused, buf[MAX_PATH];
	*buf = '\0'; // Set default.
	if (g->mLoopFile)
	{
		// GetFullPathName() is done in addition to ConvertFilespecToCorrectCase() for the following reasons:
		// 1) It's currently the only easy way to get the full path of the directory in which a file resides.
		//    For example, if a script is passed a filename via command line parameter, that file could be
		//    either an absolute path or a relative path.  If relative, of course it's relative to A_WorkingDir.
		//    The problem is, the script would have to manually detect this, which would probably take several
		//    extra steps.
		// 2) A_LoopFileLongPath is mostly intended for the following cases, and in all of them it seems
		//    preferable to have the full/absolute path rather than the relative path:
		//    a) Files dragged onto a .ahk script when the drag-and-drop option has been enabled via the Installer.
		//    b) Files passed into the script via command line.
		// The below also serves to make a copy because changing the original would yield
		// unexpected/inconsistent results in a script that retrieves the A_LoopFileFullPath
		// but only conditionally retrieves A_LoopFileLongPath.
		if (!GetFullPathName(g->mLoopFile->cFileName, MAX_PATH, buf, &unused))
			*buf = '\0'; // It might fail if NtfsDisable8dot3NameCreation is turned on in the registry, and possibly for other reasons.
		else
			// The below is called in case the loop is being used to convert filename specs that were passed
			// in from the command line, which thus might not be the proper case (at least in the path
			// portion of the filespec), as shown in the file system:
			ConvertFilespecToCorrectCase(buf);
	}
	if (aBuf)
		_tcscpy(aBuf, buf); // v1.0.47: Must be done as a separate copy because passing a size of MAX_PATH for aBuf can crash when aBuf is actually smaller than that (even though it's large enough to hold the string). This is true for ReadRegString()'s API call and may be true for other API calls like the one here.
	return (VarSizeType)_tcslen(buf); // Must explicitly calculate the length rather than using the return value from GetFullPathName(), because ConvertFilespecToCorrectCase() expands 8.3 path components.
}

VarSizeType BIV_LoopFileShortPath(LPTSTR aBuf, LPTSTR aVarName)
// Unlike GetLoopFileShortName(), this function returns blank when there is no short path.
// This is done so that there's a way for the script to more easily tell the difference between
// an 8.3 name not being available (due to the being disabled in the registry) and the short
// name simply being the same as the long name.  For example, if short name creation is disabled
// in the registry, A_LoopFileShortName would contain the long name instead, as documented.
// But to detect if that short name is really a long name, A_LoopFileShortPath could be checked
// and if it's blank, there is no short name available.
{
	TCHAR buf[MAX_PATH];
	*buf = '\0'; // Set default.
	DWORD length = 0;        //
	if (g->mLoopFile)
		// The loop handler already prepended the script's directory in cFileName for us:
		if (   !(length = GetShortPathName(g->mLoopFile->cFileName, buf, MAX_PATH))   )
			*buf = '\0'; // It might fail if NtfsDisable8dot3NameCreation is turned on in the registry, and possibly for other reasons.
	if (aBuf)
		_tcscpy(aBuf, buf); // v1.0.47: Must be done as a separate copy because passing a size of MAX_PATH for aBuf can crash when aBuf is actually smaller than that (even though it's large enough to hold the string). This is true for ReadRegString()'s API call and may be true for other API calls like the one here.
	return (VarSizeType)length;
}

VarSizeType BIV_LoopFileTime(LPTSTR aBuf, LPTSTR aVarName)
{
	TCHAR buf[64];
	LPTSTR target_buf = aBuf ? aBuf : buf;
	*target_buf = '\0'; // Set default.
	if (g->mLoopFile)
	{
		FILETIME ft;
		switch(ctoupper(aVarName[14])) // A_LoopFileTime[A]ccessed
		{
		case 'M': ft = g->mLoopFile->ftLastWriteTime; break;
		case 'C': ft = g->mLoopFile->ftCreationTime; break;
		default: ft = g->mLoopFile->ftLastAccessTime;
		}
		FileTimeToYYYYMMDD(target_buf, ft, true);
	}
	return (VarSizeType)_tcslen(target_buf);
}

VarSizeType BIV_LoopFileAttrib(LPTSTR aBuf, LPTSTR aVarName)
{
	TCHAR buf[64];
	LPTSTR target_buf = aBuf ? aBuf : buf;
	*target_buf = '\0'; // Set default.
	if (g->mLoopFile)
		FileAttribToStr(target_buf, g->mLoopFile->dwFileAttributes);
	return (VarSizeType)_tcslen(target_buf);
}

VarSizeType BIV_LoopFileSize(LPTSTR aBuf, LPTSTR aVarName)
{
	// Don't use MAX_INTEGER_LENGTH in case user has selected a very long float format via SetFormat.
	TCHAR str[128];
	LPTSTR target_buf = aBuf ? aBuf : str;
	*target_buf = '\0';  // Set default.
	if (g->mLoopFile)
	{

		// UPDATE: 64-bit ints are now standard, so the following is obsolete:
		// It's a documented limitation that the size will show as negative if
		// greater than 2 gig, and will be wrong if greater than 4 gig.  For files
		// that large, scripts should use the KB version of this function instead.
		// If a file is over 4gig, set the value to be the maximum size (-1 when
		// expressed as a signed integer, since script variables are based entirely
		// on 32-bit signed integers due to the use of ATOI(), etc.).
		//sprintf(str, "%d%", g->mLoopFile->nFileSizeHigh ? -1 : (int)g->mLoopFile->nFileSizeLow);
		ULARGE_INTEGER ul;
		ul.HighPart = g->mLoopFile->nFileSizeHigh;
		ul.LowPart = g->mLoopFile->nFileSizeLow;
		int divider;
		switch (ctoupper(aVarName[14])) // A_LoopFileSize[K/M]B
		{
		case 'K': divider = 1024; break;
		case 'M': divider = 1024*1024; break;
		default:  divider = 0;
		}
		ITOA64((__int64)(divider ? ((unsigned __int64)ul.QuadPart / divider) : ul.QuadPart), target_buf);
	}
	return (VarSizeType)_tcslen(target_buf);
}

VarSizeType BIV_LoopRegType(LPTSTR aBuf, LPTSTR aVarName)
{
	LPTSTR value = g->mLoopRegItem ? Line::RegConvertValueType(g->mLoopRegItem->type) : _T("");
	if (aBuf)
		_tcscpy(aBuf, value);
	return (VarSizeType)_tcslen(value);
}

VarSizeType BIV_LoopRegKey(LPTSTR aBuf, LPTSTR aVarName)
{
	// Use root_key_type, not root_key (which might be a remote vs. local HKEY):
	LPTSTR value = g->mLoopRegItem ? Line::RegConvertRootKeyType(g->mLoopRegItem->root_key_type) : _T("");
	if (aBuf)
		_tcscpy(aBuf, value);
	return (VarSizeType)_tcslen(value);
}

VarSizeType BIV_LoopRegSubKey(LPTSTR aBuf, LPTSTR aVarName)
{
	LPTSTR str = g->mLoopRegItem ? g->mLoopRegItem->subkey : _T("");
	if (aBuf)
		_tcscpy(aBuf, str);
	return (VarSizeType)_tcslen(str);
}

VarSizeType BIV_LoopRegName(LPTSTR aBuf, LPTSTR aVarName)
{
	// This can be either the name of a subkey or the name of a value.
	LPTSTR str = g->mLoopRegItem ? g->mLoopRegItem->name : _T("");
	if (aBuf)
		_tcscpy(aBuf, str);
	return (VarSizeType)_tcslen(str);
}

VarSizeType BIV_LoopRegTimeModified(LPTSTR aBuf, LPTSTR aVarName)
{
	TCHAR buf[64];
	LPTSTR target_buf = aBuf ? aBuf : buf;
	*target_buf = '\0'; // Set default.
	// Only subkeys (not values) have a time.  In addition, Win9x doesn't support retrieval
	// of the time (nor does it store it), so make the var blank in that case:
	if (g->mLoopRegItem && g->mLoopRegItem->type == REG_SUBKEY && !g_os.IsWin9x())
		FileTimeToYYYYMMDD(target_buf, g->mLoopRegItem->ftLastWriteTime, true);
	return (VarSizeType)_tcslen(target_buf);
}

VarSizeType BIV_LoopReadLine(LPTSTR aBuf, LPTSTR aVarName)
{
	LPTSTR str = g->mLoopReadFile ? g->mLoopReadFile->mCurrentLine : _T("");
	if (aBuf)
		_tcscpy(aBuf, str);
	return (VarSizeType)_tcslen(str);
}

VarSizeType BIV_LoopField(LPTSTR aBuf, LPTSTR aVarName)
{
	LPTSTR str = g->mLoopField ? g->mLoopField : _T("");
	if (aBuf)
		_tcscpy(aBuf, str);
	return (VarSizeType)_tcslen(str);
}

VarSizeType BIV_LoopIndex(LPTSTR aBuf, LPTSTR aVarName)
{
	return aBuf
		? (VarSizeType)_tcslen(ITOA64(g->mLoopIteration, aBuf)) // Must return exact length when aBuf isn't NULL.
		: MAX_INTEGER_LENGTH; // Probably performs better to return a conservative estimate for the first pass than to call ITOA64 for both passes.
}



VarSizeType BIV_ThisFunc(LPTSTR aBuf, LPTSTR aVarName)
{
	LPTSTR name;
	if (g->CurrentFunc)
		name = g->CurrentFunc->mName;
	else if (g->CurrentFuncGosub) // v1.0.48.02: For flexibility and backward compatibility, support A_ThisFunc even when a function Gosubs an external subroutine.
		name = g->CurrentFuncGosub->mName;
	else
		name = _T("");
	if (aBuf)
		_tcscpy(aBuf, name);
	return (VarSizeType)_tcslen(name);
}

VarSizeType BIV_ThisLabel(LPTSTR aBuf, LPTSTR aVarName)
{
	LPTSTR name = g->CurrentLabel ? g->CurrentLabel->mName : _T("");
	if (aBuf)
		_tcscpy(aBuf, name);
	return (VarSizeType)_tcslen(name);
}

VarSizeType BIV_ThisMenuItem(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
		_tcscpy(aBuf, g_script.mThisMenuItemName);
	return (VarSizeType)_tcslen(g_script.mThisMenuItemName);
}

UINT Script::ThisMenuItemPos()
{
	UserMenu *menu = FindMenu(mThisMenuName);
	// The menu item's address was stored so we can distinguish between multiple items
	// which have the same text.  The volatility of the address is handled by clearing
	// it in UserMenu::DeleteItem and UserMenu::DeleteAllItems.  An ID would also be
	// volatile, since IDs can be re-used if the item is deleted.
	if (mThisMenuItem)
	{
		UINT pos = 0;
		for (UserMenuItem *mi = menu->mFirstMenuItem; mi; mi = mi->mNextMenuItem, ++pos)
			if (mi == mThisMenuItem)
				return pos;
	}
	// For backward-compatibility, fall back to the old method if the item/menu has been
	// deleted.  So by definition, this variable contains the CURRENT position of the most
	// recently selected menu item within its CURRENT menu:
	return menu ? menu->GetItemPos(mThisMenuItemName) : UINT_MAX;
}

VarSizeType BIV_ThisMenuItemPos(LPTSTR aBuf, LPTSTR aVarName)
{
	if (!aBuf) // To avoid doing possibly high-overhead calls twice, merely return a conservative estimate for the first pass.
		return MAX_INTEGER_LENGTH;
	UINT menu_item_pos = g_script.ThisMenuItemPos();
	if (menu_item_pos < UINT_MAX) // Success
		return (VarSizeType)_tcslen(UTOA(menu_item_pos + 1, aBuf)); // +1 to convert from zero-based to 1-based.
	// Otherwise:
	*aBuf = '\0';
	return 0;
}

VarSizeType BIV_ThisMenu(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
		_tcscpy(aBuf, g_script.mThisMenuName);
	return (VarSizeType)_tcslen(g_script.mThisMenuName);
}

VarSizeType BIV_ThisHotkey(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
		_tcscpy(aBuf, g_script.mThisHotkeyName);
	return (VarSizeType)_tcslen(g_script.mThisHotkeyName);
}

VarSizeType BIV_PriorHotkey(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
		_tcscpy(aBuf, g_script.mPriorHotkeyName);
	return (VarSizeType)_tcslen(g_script.mPriorHotkeyName);
}

VarSizeType BIV_TimeSinceThisHotkey(LPTSTR aBuf, LPTSTR aVarName)
{
	if (!aBuf) // IMPORTANT: Conservative estimate because the time might change between 1st & 2nd calls.
		return MAX_INTEGER_LENGTH;
	// It must be the type of hotkey that has a label because we want the TimeSinceThisHotkey
	// value to be "in sync" with the value of ThisHotkey itself (i.e. use the same method
	// to determine which hotkey is the "this" hotkey):
	if (*g_script.mThisHotkeyName)
		// Even if GetTickCount()'s TickCount has wrapped around to zero and the timestamp hasn't,
		// DWORD subtraction still gives the right answer as long as the number of days between
		// isn't greater than about 49.  See MyGetTickCount() for explanation of %d vs. %u.
		// Update: Using 64-bit ints now, so above is obsolete:
		//sntprintf(str, sizeof(str), "%d", (DWORD)(GetTickCount() - g_script.mThisHotkeyStartTime));
		ITOA64((__int64)(GetTickCount() - g_script.mThisHotkeyStartTime), aBuf);
	else
		_tcscpy(aBuf, _T("-1"));
	return (VarSizeType)_tcslen(aBuf);
}

VarSizeType BIV_TimeSincePriorHotkey(LPTSTR aBuf, LPTSTR aVarName)
{
	if (!aBuf) // IMPORTANT: Conservative estimate because the time might change between 1st & 2nd calls.
		return MAX_INTEGER_LENGTH;
	if (*g_script.mPriorHotkeyName)
		// See MyGetTickCount() for explanation for explanation:
		//sntprintf(str, sizeof(str), "%d", (DWORD)(GetTickCount() - g_script.mPriorHotkeyStartTime));
		ITOA64((__int64)(GetTickCount() - g_script.mPriorHotkeyStartTime), aBuf);
	else
		_tcscpy(aBuf, _T("-1"));
	return (VarSizeType)_tcslen(aBuf);
}

VarSizeType BIV_EndChar(LPTSTR aBuf, LPTSTR aVarName)
{
	if (aBuf)
	{
		if (g_script.mEndChar)
			*aBuf++ = g_script.mEndChar;
		//else we returned 0 previously, so MUST WRITE ONLY ONE NULL-TERMINATOR.
		*aBuf = '\0';
	}
	return g_script.mEndChar ? 1 : 0; // v1.0.48.04: Fixed to support a NULL char, which happens when the hotstring has the "no ending character required" option.
}



VarSizeType BIV_Gui(LPTSTR aBuf, LPTSTR aVarName)
// We're returning the length of the var's contents, not the size.
{
	TCHAR buf[MAX_INTEGER_SIZE];
	LPTSTR target_buf = aBuf ? aBuf : buf;
	GuiType* gui = g->GuiWindow; // For performance.

	if (!gui) // The current thread was not launched as a result of GUI action.
	{
		*target_buf = '\0';
		return 0;
	}

	switch (ctoupper(aVarName[5]))
	{
	case 'W':
		// g->GuiPoint.x was overloaded to contain the size, since there are currently never any cases when
		// A_GuiX/Y and A_GuiWidth/Height are both valid simultaneously.  It is documented that each of these
		// variables is defined only in proper types of subroutines.
		_itot(gui->Unscale(LOWORD(g->GuiPoint.x)), target_buf, 10);
		// Above is always stored as decimal vs. hex, regardless of script settings.
		break;
	case 'H':
		_itot(gui->Unscale(HIWORD(g->GuiPoint.x)), target_buf, 10); // See comments above.
		break;
	case 'X':
		_itot(gui->Unscale(g->GuiPoint.x), target_buf, 10);
		break;
	case 'Y':
		_itot(gui->Unscale(g->GuiPoint.y), target_buf, 10);
		break;
	case '\0': // A_Gui
		if (!*g->GuiWindow->mName) // v1.1.04: Anonymous GUI.
			return _stprintf(target_buf, _T("0x%Ix"), (UINT_PTR)g->GuiWindow->mHwnd);
		if (aBuf)
			_tcscpy(aBuf, g->GuiWindow->mName);
		return _tcslen(g->GuiWindow->mName);
	}

	return (VarSizeType)_tcslen(target_buf);
}



VarSizeType BIV_GuiControl(LPTSTR aBuf, LPTSTR aVarName)
{
	return GuiType::ControlGetName(g->GuiWindow, g->GuiControlIndex, aBuf);
}



VarSizeType BIV_GuiEvent(LPTSTR aBuf, LPTSTR aVarName)
// We're returning the length of the var's contents, not the size.
{
	global_struct &g = *::g; // Reduces code size and may improve performance.
	if (g.GuiEvent == GUI_EVENT_DROPFILES)
	{
		GuiType *pgui;
		UINT u, file_count;
		// GUI_EVENT_DROPFILES should mean that g.GuiWindow != NULL, but the below will double check that in
		// case g.GuiEvent can ever be set to that value as a result of receiving a bogus message in the queue.
		if (!(pgui = g.GuiWindow) // The current thread was not launched as a result of GUI action or this is a bogus msg.
			|| !pgui->mHwnd // Gui window no longer exists.  Relies on short-circuit boolean.
			|| !pgui->mHdrop // No HDROP (probably impossible unless g.GuiEvent was given a bogus value somehow).
			|| !(file_count = DragQueryFile(pgui->mHdrop, 0xFFFFFFFF, NULL, 0))) // No files in the drop (not sure if this is possible).
			// All of the above rely on short-circuit boolean order.
		{
			// Make the dropped-files list blank since there is no HDROP to query (or no files in it).
			if (aBuf)
				*aBuf = '\0';
			return 0;
		}
		// Above has ensured that file_count > 0
		if (aBuf)
		{
			TCHAR buf[MAX_PATH], *cp = aBuf;
			UINT length;
			for (u = 0; u < file_count; ++u)
			{
				length = DragQueryFile(pgui->mHdrop, u, buf, MAX_PATH); // MAX_PATH is arbitrary since aBuf is already known to be large enough.
				_tcscpy(cp, buf); // v1.0.47: Must be done as a separate copy because passing a size of MAX_PATH for something that isn't actually that large (though clearly large enough) due to previous size-estimation phase) can crash because the API may read/write data beyond what it actually needs.
				cp += length;
				if (u < file_count - 1) // i.e omit the LF after the last file to make parsing via "Loop, Parse" easier.
					*cp++ = '\n';
				// Although the transcription of files on the clipboard into their text filenames is done
				// with \r\n (so that they're in the right format to be pasted to other apps as a plain text
				// list), it seems best to use a plain linefeed for dropped files since they won't be going
				// onto the clipboard nearly as often, and `n is easier to parse.  Also, a script array isn't
				// used because large file lists would then consume a lot more of memory because arrays
				// are permanent once created, and also there would be wasted space due to the part of each
				// variable's capacity not used by the filename.
			}
			// No need for final termination of string because the last item lacks a newline.
			return (VarSizeType)(cp - aBuf); // This is the length of what's in the buffer.
		}
		else
		{
			VarSizeType total_length = 0;
			for (u = 0; u < file_count; ++u)
				total_length += DragQueryFile(pgui->mHdrop, u, NULL, 0);
				// Above: MSDN: "If the lpszFile buffer address is NULL, the return value is the required size,
				// in characters, of the buffer, not including the terminating null character."
			return total_length + file_count - 1; // Include space for a linefeed after each filename except the last.
		}
		// Don't call DragFinish() because this variable might be referred to again before this thread
		// is done.  DragFinish() is called by MsgSleep() when the current thread finishes.
	}

	// Otherwise, this event is not GUI_EVENT_DROPFILES, so use standard modes of operation.
	LPTSTR event_string = GuiType::ConvertEvent(g.GuiEvent);
	return (VarSizeType)_tcslen(aBuf ? _tcscpy(aBuf, event_string) : event_string);
}



VarSizeType BIV_DefaultGui(LPTSTR aBuf, LPTSTR aVarName)
// A_DefaultGui, A_DefaultListView, A_DefaultTreeView: Unlike BIV_Gui above, these
// correspond to "Gui, x: Default", not necessarily the Gui which launched the thread.
{
	global_struct &g = *::g; // Reduces code size and may improve performance.
	GuiType *gui = g.GuiDefaultWindowValid();
	GuiControlType *control = NULL;
	LPTSTR return_string = _T("");
	HWND return_hwnd = NULL;
	switch (ctoupper(aVarName[9]))
	{
	case 'G':
		if (!gui)
			gui = g.GuiDefaultWindow; // If non-NULL, it's a dummy struct containing just the Gui name.
		if (!gui) // Either no default has been set, or the default was an anonymous Gui which has been destroyed.
			return_string = _T("1");
		else if (*gui->mName) // Not an anonymous GUI.
			return_string = gui->mName;
		else // implies gui->mHwnd != NULL
			return_hwnd = gui->mHwnd;
		break;
	case 'L':
		if (gui)
			control = gui->mCurrentListView;
		break;
	case 'T':
		if (gui)
			control = gui->mCurrentTreeView;
		break;
	}
	if (control)
	{
		if (control->output_var) // Return associated var name (more useful for debugging).
			return_string = control->output_var->mName;
		else // Return HWND.
			return_hwnd = control->hwnd;
	}
	if (return_hwnd)
	{
		if (aBuf)
			return (VarSizeType)_tcslen(HwndToString(return_hwnd, aBuf));
		return MAX_INTEGER_LENGTH;
	}
	if (aBuf)
		_tcscpy(aBuf, return_string);
	return (VarSizeType)_tcslen(return_string);
}



VarSizeType BIV_EventInfo(LPTSTR aBuf, LPTSTR aVarName)
// We're returning the length of the var's contents, not the size.
{
	return aBuf
		? (VarSizeType)_tcslen(Exp32or64(UTOA,UTOA64)(g->EventInfo, aBuf)) // Must return exact length when aBuf isn't NULL.
		: MAX_INTEGER_LENGTH;
}



VarSizeType BIV_TimeIdle(LPTSTR aBuf, LPTSTR aVarName) // Called by multiple callers.
{
	if (!aBuf) // IMPORTANT: Conservative estimate because tick might change between 1st & 2nd calls.
		return MAX_INTEGER_LENGTH;
#ifdef CONFIG_WIN9X
	*aBuf = '\0';  // Set default.
	if (g_os.IsWin2000orLater()) // Checked in case the function is present in the OS but "not implemented".
	{
		// Must fetch it at runtime, otherwise the program can't even be launched on Win9x/NT:
		typedef BOOL (WINAPI *MyGetLastInputInfoType)(PLASTINPUTINFO);
		static MyGetLastInputInfoType MyGetLastInputInfo = (MyGetLastInputInfoType)
			GetProcAddress(GetModuleHandle(_T("user32")), "GetLastInputInfo");
		if (MyGetLastInputInfo)
		{
			LASTINPUTINFO lii;
			lii.cbSize = sizeof(lii);
			if (MyGetLastInputInfo(&lii))
				ITOA64(GetTickCount() - lii.dwTime, aBuf);
		}
	}
#else
	// Not Win9x: Calling it directly should (in theory) produce smaller code size.
	LASTINPUTINFO lii;
	lii.cbSize = sizeof(lii);
	if (GetLastInputInfo(&lii))
		ITOA64(GetTickCount() - lii.dwTime, aBuf);
	else
		*aBuf = '\0';
#endif
	return (VarSizeType)_tcslen(aBuf);
}



VarSizeType BIV_TimeIdlePhysical(LPTSTR aBuf, LPTSTR aVarName)
// This is here rather than in script.h with the others because it depends on
// hotkey.h and globaldata.h, which can't be easily included in script.h due to
// mutual dependency issues.
{
	DWORD time_last_input;
	switch (toupper(aVarName[10]))
	{
	case 'M': time_last_input = g_MouseHook ? g_TimeLastInputMouse : 0; break;
	case 'K': time_last_input = g_KeybdHook ? g_TimeLastInputKeyboard : 0; break;
	default: time_last_input = (g_KeybdHook || g_MouseHook) ? g_TimeLastInputPhysical : 0; break;
	}
	// If the relevant hook is not active, default this to the same as the regular idle time:
	if (!time_last_input)
		return BIV_TimeIdle(aBuf, _T(""));
	if (!aBuf)
		return MAX_INTEGER_LENGTH; // IMPORTANT: Conservative estimate because tick might change between 1st & 2nd calls.
	return (VarSizeType)_tcslen(ITOA64(GetTickCount() - time_last_input, aBuf)); // Switching keyboard layouts/languages sometimes sees to throw off the timestamps of the incoming events in the hook.
}


////////////////////////
// BUILT-IN FUNCTIONS //
////////////////////////

#ifdef ENABLE_DLLCALL

#ifdef WIN32_PLATFORM
// Interface for DynaCall():
#define  DC_MICROSOFT           0x0000      // Default
#define  DC_BORLAND             0x0001      // Borland compat
#define  DC_CALL_CDECL          0x0010      // __cdecl
#define  DC_CALL_STD            0x0020      // __stdcall
#define  DC_RETVAL_MATH4        0x0100      // Return value in ST
#define  DC_RETVAL_MATH8        0x0200      // Return value in ST

#define  DC_CALL_STD_BO         (DC_CALL_STD | DC_BORLAND)
#define  DC_CALL_STD_MS         (DC_CALL_STD | DC_MICROSOFT)
#define  DC_CALL_STD_M8         (DC_CALL_STD | DC_RETVAL_MATH8)
#endif

union DYNARESULT                // Various result types
{      
    int     Int;                // Generic four-byte type
    long    Long;               // Four-byte long
    void   *Pointer;            // 32-bit pointer
    float   Float;              // Four byte real
    double  Double;             // 8-byte real
    __int64 Int64;              // big int (64-bit)
	UINT_PTR UIntPtr;
};

struct DYNAPARM
{
    union
	{
		int value_int; // Args whose width is less than 32-bit are also put in here because they are right justified within a 32-bit block on the stack.
		float value_float;
		__int64 value_int64;
		UINT_PTR value_uintptr;
		double value_double;
		char *astr;
		wchar_t *wstr;
		void *ptr;
    };
	// Might help reduce struct size to keep other members last and adjacent to each other (due to
	// 8-byte alignment caused by the presence of double and __int64 members in the union above).
	DllArgTypes type;
	bool passed_by_address;
	bool is_unsigned; // Allows return value and output parameters to be interpreted as unsigned vs. signed.
};

#ifdef _WIN64
// This function was borrowed from http://dyncall.org/
extern "C" UINT_PTR PerformDynaCall(size_t stackArgsSize, DWORD_PTR* stackArgs, DWORD_PTR* regArgs, void* aFunction);

// Retrieve a float or double return value.  These don't actually do anything, since the value we
// want is already in the xmm0 register which is used to return float or double values.
// Many thanks to http://locklessinc.com/articles/c_abi_hacks/ for the original idea.
extern "C" float read_xmm0_float();
extern "C" double read_xmm0_double();

static inline UINT_PTR DynaParamToElement(DYNAPARM& parm)
{
	if(parm.passed_by_address)
		return (UINT_PTR) &parm.value_uintptr;
	else
		return parm.value_uintptr;
}
#endif

#ifdef WIN32_PLATFORM
DYNARESULT DynaCall(int aFlags, void *aFunction, DYNAPARM aParam[], int aParamCount, DWORD &aException
	, void *aRet, int aRetSize)
#elif defined(_WIN64)
DYNARESULT DynaCall(void *aFunction, DYNAPARM aParam[], int aParamCount, DWORD &aException)
#else
#error DllCall not supported on this platform
#endif
// Based on the code by Ton Plooy <tonp@xs4all.nl>.
// Call the specified function with the given parameters. Build a proper stack and take care of correct
// return value processing.
{
	aException = 0;  // Set default output parameter for caller.
	SetLastError(g->LastError); // v1.0.46.07: In case the function about to be called doesn't change last-error, this line serves to retain the script's previous last-error rather than some arbitrary one produced by AutoHotkey's own internal API calls.  This line has no measurable impact on performance.

    DYNARESULT Res = {0}; // This struct is to be returned to caller by value.

#ifdef WIN32_PLATFORM

	// Declaring all variables early should help minimize stack interference of C code with asm.
	DWORD *our_stack;
    int param_size;
	DWORD stack_dword, our_stack_size = 0; // Both might have to be DWORD for _asm.
	BYTE *cp;
    DWORD esp_start, esp_end, dwEAX, dwEDX;
	int i, esp_delta; // Declare this here rather than later to prevent C code from interfering with esp.

	// Reserve enough space on the stack to handle the worst case of our args (which is currently a
	// maximum of 8 bytes per arg). This avoids any chance that compiler-generated code will use
	// the stack in a way that disrupts our insertion of args onto the stack.
	DWORD reserved_stack_size = aParamCount * 8;
	_asm
	{
		mov our_stack, esp  // our_stack is the location where we will write our args (bypassing "push").
		sub esp, reserved_stack_size  // The stack grows downward, so this "allocates" space on the stack.
	}

	// "Push" args onto the portion of the stack reserved above. Every argument is aligned on a 4-byte boundary.
	// We start at the rightmost argument (i.e. reverse order).
	for (i = aParamCount - 1; i > -1; --i)
	{
		DYNAPARM &this_param = aParam[i]; // For performance and convenience.
		// Push the arg or its address onto the portion of the stack that was reserved for our use above.
		if (this_param.passed_by_address)
		{
			stack_dword = (DWORD)(size_t)&this_param.value_int; // Any union member would work.
			--our_stack;              // ESP = ESP - 4
			*our_stack = stack_dword; // SS:[ESP] = stack_dword
			our_stack_size += 4;      // Keep track of how many bytes are on our reserved portion of the stack.
		}
		else // this_param's value is contained directly inside the union.
		{
			param_size = (this_param.type == DLL_ARG_INT64 || this_param.type == DLL_ARG_DOUBLE) ? 8 : 4;
			our_stack_size += param_size; // Must be done before our_stack_size is decremented below.  Keep track of how many bytes are on our reserved portion of the stack.
			cp = (BYTE *)&this_param.value_int + param_size - 4; // Start at the right side of the arg and work leftward.
			while (param_size > 0)
			{
				stack_dword = *(DWORD *)cp;  // Get first four bytes
				cp -= 4;                     // Next part of argument
				--our_stack;                 // ESP = ESP - 4
				*our_stack = stack_dword;    // SS:[ESP] = stack_dword
				param_size -= 4;
			}
		}
    }

	if ((aRet != NULL) && ((aFlags & DC_BORLAND) || (aRetSize > 8)))
	{
		// Return value isn't passed through registers, memory copy
		// is performed instead. Pass the pointer as hidden arg.
		our_stack_size += 4;       // Add stack size
		--our_stack;               // ESP = ESP - 4
		*our_stack = (DWORD)(size_t)aRet;  // SS:[ESP] = pMem
	}

	// Call the function.
	__try // Each try/except section adds at most 240 bytes of uncompressed code, and typically doesn't measurably affect performance.
	{
		_asm
		{
			add esp, reserved_stack_size // Restore to original position
			mov esp_start, esp      // For detecting whether a DC_CALL_STD function was sent too many or too few args.
			sub esp, our_stack_size // Adjust ESP to indicate that the args have already been pushed onto the stack.
			call [aFunction]        // Stack is now properly built, we can call the function
		}
	}
	__except(EXCEPTION_EXECUTE_HANDLER)
	{
		aException = GetExceptionCode(); // aException is an output parameter for our caller.
	}

	// Even if an exception occurred (perhaps due to the callee having been passed a bad pointer),
	// attempt to restore the stack to prevent making things even worse.
	_asm
	{
		mov esp_end, esp        // See below.
		mov esp, esp_start      //
		// For DC_CALL_STD functions (since they pop their own arguments off the stack):
		// Since the stack grows downward in memory, if the value of esp after the call is less than
		// that before the call's args were pushed onto the stack, there are still items left over on
		// the stack, meaning that too many args (or an arg too large) were passed to the callee.
		// Conversely, if esp is now greater that it should be, too many args were popped off the
		// stack by the callee, meaning that too few args were provided to it.  In either case,
		// and even for CDECL, the following line restores esp to what it was before we pushed the
		// function's args onto the stack, which in the case of DC_CALL_STD helps prevent crashes
		// due to too many or to few args having been passed.
		mov dwEAX, eax          // Save eax/edx registers
		mov dwEDX, edx
	}

	// Possibly adjust stack and read return values.
	// The following is commented out because the stack (esp) is restored above, for both CDECL and STD.
	//if (aFlags & DC_CALL_CDECL)
	//	_asm add esp, our_stack_size    // CDECL requires us to restore the stack after the call.
	if (aFlags & DC_RETVAL_MATH4)
		_asm fstp dword ptr [Res]
	else if (aFlags & DC_RETVAL_MATH8)
		_asm fstp qword ptr [Res]
	else if (!aRet)
	{
		_asm
		{
			mov  eax, [dwEAX]
			mov  DWORD PTR [Res], eax
			mov  edx, [dwEDX]
			mov  DWORD PTR [Res + 4], edx
		}
	}
	else if (((aFlags & DC_BORLAND) == 0) && (aRetSize <= 8))
	{
		// Microsoft optimized less than 8-bytes structure passing
        _asm
		{
			mov ecx, DWORD PTR [aRet]
			mov eax, [dwEAX]
			mov DWORD PTR [ecx], eax
			mov edx, [dwEDX]
			mov DWORD PTR [ecx + 4], edx
		}
	}

#endif // WIN32_PLATFORM
#ifdef _WIN64

	int params_left = aParamCount;
	DWORD_PTR regArgs[4];
	DWORD_PTR* stackArgs = NULL;
	size_t stackArgsSize = 0;

	// The first four parameters are passed in x64 through registers... like ARM :D
	for(int i = 0; (i < 4) && params_left; i++, params_left--)
		regArgs[i] = DynaParamToElement(aParam[i]);

	// Copy the remaining parameters
	if(params_left)
	{
		stackArgsSize = params_left * 8;
		stackArgs = (DWORD_PTR*) _alloca(stackArgsSize);

		for(int i = 0; i < params_left; i ++)
			stackArgs[i] = DynaParamToElement(aParam[i+4]);
	}

	// Call the function.
	__try
	{
		Res.UIntPtr = PerformDynaCall(stackArgsSize, stackArgs, regArgs, aFunction);
	}
	__except(EXCEPTION_EXECUTE_HANDLER)
	{
		aException = GetExceptionCode(); // aException is an output parameter for our caller.
	}

#endif

	// v1.0.42.03: The following supports A_LastError. It's called even if an exception occurred because it
	// might add value in some such cases.  Benchmarks show that this has no measurable impact on performance.
	// A_LastError was implemented rather than trying to change things so that a script could use DllCall to
	// call GetLastError() because: Even if we could avoid calling any API function that resets LastError
	// (which seems unlikely) it would be difficult to maintain (and thus a source of bugs) as revisions are
	// made in the future.
	g->LastError = GetLastError();

	TCHAR buf[32];

#ifdef WIN32_PLATFORM
	esp_delta = esp_start - esp_end; // Positive number means too many args were passed, negative means too few.
	if (esp_delta && (aFlags & DC_CALL_STD))
	{
		*buf = 'A'; // The 'A' prefix indicates the call was made, but with too many or too few args.
		_itot(esp_delta, buf + 1, 10);
		g_script.SetErrorLevelOrThrowStr(buf, _T("DllCall")); // Assign buf not _itot()'s return value, which is the wrong location.
	}
	else
#endif
	// Too many or too few args takes precedence over reporting the exception because it's more informative.
	// In other words, any exception was likely caused by the fact that there were too many or too few.
	if (aException)
	{
		// It's a little easier to recognize the common error codes when they're in hex format.
		buf[0] = '0';
		buf[1] = 'x';
		_ultot(aException, buf + 2, 16);
		g_script.SetErrorLevelOrThrowStr(buf, _T("DllCall")); // Positive ErrorLevel numbers are reserved for exception codes.
	}
	else
		g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	return Res;
}



void ConvertDllArgType(LPTSTR aBuf[], DYNAPARM &aDynaParam)
// Helper function for DllCall().  Updates aDynaParam's type and other attributes.
// Caller has ensured that aBuf contains exactly two strings (though the second can be NULL).
{
	LPTSTR type_string;
	TCHAR buf[32];
	int i;

	// Up to two iterations are done to cover the following cases:
	// No second type because there was no SYM_VAR to get it from:
	//	blank means int
	//	invalid is err
	// (for the below, note that 2nd can't be blank because var name can't be blank, and the first case above would have caught it if 2nd is NULL)
	// 1Blank, 2Invalid: blank (but ensure is_unsigned and passed_by_address get reset)
	// 1Blank, 2Valid: 2
	// 1Valid, 2Invalid: 1 (second iteration would never have run, so no danger of it having erroneously reset is_unsigned/passed_by_address)
	// 1Valid, 2Valid: 1 (same comment)
	// 1Invalid, 2Invalid: invalid
	// 1Invalid, 2Valid: 2

	for (i = 0, type_string = aBuf[0]; i < 2 && type_string; type_string = aBuf[++i])
	{
		if (ctoupper(*type_string) == 'U') // Unsigned
		{
			aDynaParam.is_unsigned = true;
			++type_string; // Omit the 'U' prefix from further consideration.
		}
		else
			aDynaParam.is_unsigned = false;
		
		// Check for empty string before checking for pointer suffix, so that we can skip the first character.  This is needed to simplify "Ptr" type-name support.
		if (!*type_string)
		{
			// The following also serves to set the default in case this is the first iteration.
			// Set default but perform second iteration in case the second type string isn't NULL.
			// In other words, if the second type string is explicitly valid rather than blank,
			// it should override the following default:
			aDynaParam.type = DLL_ARG_INVALID;  // To assist with detection of errors like DllCall(...,flaot,n), treat empty string as an error; naked "CDecl" is now handled elsewhere.  OBSOLETE COMMENT: Assume int.  This is relied upon at least for having a return type such as a naked "CDecl".
			continue; // OK to do this regardless of whether this is the first or second iteration.
		}

		tcslcpy(buf, type_string, _countof(buf)); // Make a modifiable copy for easier parsing below.

		// v1.0.30.02: The addition of 'P' allows the quotes to be omitted around a pointer type.
		// However, the current detection below relies upon the fact that not of the types currently
		// contain the letter P anywhere in them, so it would have to be altered if that ever changes.
		LPTSTR cp = StrChrAny(buf + 1, _T("*pP")); // Asterisk or the letter P.  Relies on the check above to ensure type_string is not empty (and buf + 1 is valid).
		if (cp && !*omit_leading_whitespace(cp + 1)) // Additional validation: ensure nothing following the suffix.
		{
			aDynaParam.passed_by_address = true;
			// Remove trailing options so that stricmp() can be used below.
			// Allow optional space in front of asterisk (seems okay even for 'P').
			if (IS_SPACE_OR_TAB(cp[-1]))
			{
				cp = omit_trailing_whitespace(buf, cp - 1);
				cp[1] = '\0'; // Terminate at the leftmost whitespace to remove all whitespace and the suffix.
			}
			else
				*cp = '\0'; // Terminate at the suffix to remove it.
		}
		else
			aDynaParam.passed_by_address = false;

		if (false) {} // To simplify the macro below.  It should have no effect on the compiled code.
#define TEST_TYPE(t, n)  else if (!_tcsicmp(buf, _T(t)))  aDynaParam.type = (n);
		TEST_TYPE("Int",	DLL_ARG_INT) // The few most common types are kept up top for performance.
		TEST_TYPE("Str",	DLL_ARG_STR)
#ifdef _WIN64
		TEST_TYPE("Ptr",	DLL_ARG_INT64) // Ptr vs IntPtr to simplify recognition of the pointer suffix, to avoid any possible confusion with IntP, and because it is easier to type.
#else
		TEST_TYPE("Ptr",	DLL_ARG_INT)
#endif
		TEST_TYPE("Short",	DLL_ARG_SHORT)
		TEST_TYPE("Char",	DLL_ARG_CHAR)
		TEST_TYPE("Int64",	DLL_ARG_INT64)
		TEST_TYPE("Float",	DLL_ARG_FLOAT)
		TEST_TYPE("Double",	DLL_ARG_DOUBLE)
		TEST_TYPE("AStr",	DLL_ARG_ASTR)
		TEST_TYPE("WStr",	DLL_ARG_WSTR)
#undef TEST_TYPE
		else // It's non-blank but an unknown type.
		{
			if (i > 0) // Second iteration.
			{
				// Reset flags to go with any blank value (i.e. !*buf) we're falling back to from the first iteration
				// (in case our iteration changed the flags based on bogus contents of the second type_string):
				aDynaParam.passed_by_address = false;
				aDynaParam.is_unsigned = false;
				//aDynaParam.type: The first iteration already set it to DLL_ARG_INT or DLL_ARG_INVALID.
			}
			else // First iteration, so aDynaParam.type's value will be set by the second (however, the loop's own condition will skip the second iteration if the second type_string is NULL).
			{
				aDynaParam.type = DLL_ARG_INVALID; // Set in case of: 1) the second iteration is skipped by the loop's own condition (since the caller doesn't always initialize "type"); or 2) the second iteration can't find a valid type.
				continue;
			}
		}
		// Since above didn't "continue", the type is explicitly valid so "return" to ensure that
		// the second iteration doesn't run (in case this is the first iteration):
		return;
	}
}



bool IsDllArgTypeName(LPTSTR name)
// Test whether given name is a valid DllCall arg type (used by Script::MaybeWarnLocalSameAsGlobal).
{
	LPTSTR names[] = { name, NULL };
	DYNAPARM param;
	// An alternate method using an array of strings and tcslicmp in a loop benchmarked
	// slightly faster than this, but didn't seem worth the extra code size. This should
	// be more maintainable and is guaranteed to be consistent with what DllCall accepts.
	ConvertDllArgType(names, param);
	return param.type != DLL_ARG_INVALID;
}



void *GetDllProcAddress(LPCTSTR aDllFileFunc, HMODULE *hmodule_to_free) // L31: Contains code extracted from BIF_DllCall for reuse in ExpressionToPostfix.
{
	int i;
	void *function = NULL;
	TCHAR param1_buf[MAX_PATH*2], *_tfunction_name, *dll_name; // Must use MAX_PATH*2 because the function name is INSIDE the Dll file, and thus MAX_PATH can be exceeded.
#ifndef UNICODE
	char *function_name;
#endif

	// Define the standard libraries here. If they reside in %SYSTEMROOT%\system32 it is not
	// necessary to specify the full path (it wouldn't make sense anyway).
	static HMODULE sStdModule[] = {GetModuleHandle(_T("user32")), GetModuleHandle(_T("kernel32"))
		, GetModuleHandle(_T("comctl32")), GetModuleHandle(_T("gdi32"))}; // user32 is listed first for performance.
	static const int sStdModule_count = _countof(sStdModule);

	// Make a modifiable copy of param1 so that the DLL name and function name can be parsed out easily, and so that "A" or "W" can be appended if necessary (e.g. MessageBoxA):
	tcslcpy(param1_buf, aDllFileFunc, _countof(param1_buf) - 1); // -1 to reserve space for the "A" or "W" suffix later below.
	if (   !(_tfunction_name = _tcsrchr(param1_buf, '\\'))   ) // No DLL name specified, so a search among standard defaults will be done.
	{
		dll_name = NULL;
#ifdef UNICODE
		char function_name[MAX_PATH];
		WideCharToMultiByte(CP_ACP, 0, param1_buf, -1, function_name, _countof(function_name), NULL, NULL);
#else
		function_name = param1_buf;
#endif

		// Since no DLL was specified, search for the specified function among the standard modules.
		for (i = 0; i < sStdModule_count; ++i)
			if (   sStdModule[i] && (function = (void *)GetProcAddress(sStdModule[i], function_name))   )
				break;
		if (!function)
		{
			// Since the absence of the "A" suffix (e.g. MessageBoxA) is so common, try it that way
			// but only here with the standard libraries since the risk of ambiguity (calling the wrong
			// function) seems unacceptably high in a custom DLL.  For example, a custom DLL might have
			// function called "AA" but not one called "A".
			strcat(function_name, WINAPI_SUFFIX); // 1 byte of memory was already reserved above for the 'A'.
			for (i = 0; i < sStdModule_count; ++i)
				if (   sStdModule[i] && (function = (void *)GetProcAddress(sStdModule[i], function_name))   )
					break;
		}
	}
	else // DLL file name is explicitly present.
	{
		dll_name = param1_buf;
		*_tfunction_name = '\0';  // Terminate dll_name to split it off from function_name.
		++_tfunction_name; // Set it to the character after the last backslash.
#ifdef UNICODE
		char function_name[MAX_PATH];
		WideCharToMultiByte(CP_ACP, 0, _tfunction_name, -1, function_name, _countof(function_name), NULL, NULL);
#else
		function_name = _tfunction_name;
#endif

		// Get module handle. This will work when DLL is already loaded and might improve performance if
		// LoadLibrary is a high-overhead call even when the library already being loaded.  If
		// GetModuleHandle() fails, fall back to LoadLibrary().
		HMODULE hmodule;
		if (   !(hmodule = GetModuleHandle(dll_name))    )
			if (   !hmodule_to_free  ||  !(hmodule = *hmodule_to_free = LoadLibrary(dll_name))   )
			{
				if (hmodule_to_free) // L31: BIF_DllCall wants us to set ErrorLevel.  ExpressionToPostfix passes NULL.
					g_script.SetErrorLevelOrThrowStr(_T("-3"), _T("DllCall")); // Stage 3 error: DLL couldn't be loaded.
				return NULL;
			}
		if (   !(function = (void *)GetProcAddress(hmodule, function_name))   )
		{
			// v1.0.34: If it's one of the standard libraries, try the "A" suffix.
			// jackieku: Try it anyway, there are many other DLLs that use this naming scheme, and it doesn't seem expensive.
			// If an user really cares he or she can always work around it by editing the script.
			//for (i = 0; i < sStdModule_count; ++i)
			//	if (hmodule == sStdModule[i]) // Match found.
			//	{
					strcat(function_name, WINAPI_SUFFIX); // 1 byte of memory was already reserved above for the 'A'.
					function = (void *)GetProcAddress(hmodule, function_name);
			//		break;
			//	}
		}
	}

	if (!function && hmodule_to_free) // Caller wants us to set ErrorLevel.
	{
		// This must be done here since only we know for certain that the dll
		// was loaded okay (if GetModuleHandle succeeded, nothing is passed
		// back to the caller).
		g_script.SetErrorLevelOrThrowStr(_T("-4"), _T("DllCall")); // Stage 4 error: Function could not be found in the DLL(s).
	}

	return function;
}



BIF_DECL(BIF_DllCall)
// Stores a number or a SYM_STRING result in aResultToken.
// Sets ErrorLevel to the error code appropriate to any problem that occurred.
// Caller has set up aParam to be viewable as a left-to-right array of params rather than a stack.
// It has also ensured that the array has exactly aParamCount items in it.
// Author: Marcus Sonntag (Ultra)
{
	// Set default result in case of early return; a blank value:
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");
	HMODULE hmodule_to_free = NULL; // Set default in case of early goto; mostly for maintainability.
	void *function; // Will hold the address of the function to be called.

	// Check that the mandatory first parameter (DLL+Function) is valid.
	// (load-time validation has ensured at least one parameter is present).
	switch(aParam[0]->symbol)
	{
		case SYM_STRING: // By far the most common, so it's listed first for performance. Also for performance, don't even consider the possibility that a quoted literal string like "33" is a function-address.
			function = NULL; // Indicate that no function has been specified yet.
			break;
		case SYM_VAR:
			// v1.0.46.08: Allow script to specify the address of a function, which might be useful for
			// calling functions that the script discovers through unusual means such as C++ member functions.
			function = (aParam[0]->var->IsNonBlankIntegerOrFloat() == PURE_INTEGER)
				? (void *)aParam[0]->var->ToInt64(TRUE) // For simplicity and due to rarity, this doesn't check for zero or negative numbers.
				: NULL; // Not a pure integer, so fall back to normal method of considering it to be path+name.
			// A check like the following is not present due to rarity of need and because if the address
			// is zero or negative, the same result will occur as for any other invalid address:
			// an ErrorLevel of 0xc0000005.
			//if (temp64 <= 0)
			//{
			//	g_ErrorLevel->Assign(_T("-1")); // Stage 1 error: Invalid first param.
			//	return;
			//}
			//// Otherwise, assume it's a valid address:
			//	function = (void *)temp64;
			break;
		case SYM_INTEGER:
			function = (void *)aParam[0]->value_int64; // For simplicity and due to rarity, this doesn't check for zero or negative numbers.
			break;
		case SYM_FLOAT:
		case SYM_MISSING:
			g_script.SetErrorLevelOrThrowStr(_T("-1"), _T("DllCall")); // Stage 1 error: Invalid first param.
			return;
		default: // SYM_OPERAND (SYM_OPERAND is typically a numeric literal).
			function = (TokenIsPureNumeric(*aParam[0]) == PURE_INTEGER)
				? (void *)TokenToInt64(*aParam[0], TRUE) // For simplicity and due to rarity, this doesn't check for zero or negative numbers.
				: NULL; // Not a pure integer, so fall back to normal method of considering it to be path+name.
	}

	// Determine the type of return value.
	DYNAPARM return_attrib = {0}; // Init all to default in case ConvertDllArgType() isn't called below. This struct holds the type and other attributes of the function's return value.
#ifdef WIN32_PLATFORM
	int dll_call_mode = DC_CALL_STD; // Set default.  Can be overridden to DC_CALL_CDECL and flags can be OR'd into it.
#endif
	if (aParamCount % 2) // Odd number of parameters indicates the return type has been omitted, so assume BOOL/INT.
		return_attrib.type = DLL_ARG_INT;
	else
	{
		// Check validity of this arg's return type:
		ExprTokenType &token = *aParam[aParamCount - 1];
		LPTSTR return_type_string[2];
		switch (token.symbol)
		{
		case SYM_VAR: // SYM_VAR's Type() is always VAR_NORMAL (except lvalues in expressions).
			return_type_string[0] = token.var->Contents(TRUE, TRUE);	
			return_type_string[1] = token.var->mName; // v1.0.33.01: Improve convenience by falling back to the variable's name if the contents are not appropriate.
			break;
		case SYM_STRING:
		case SYM_OPERAND:
			return_type_string[0] = token.marker;
			return_type_string[1] = NULL; // Added in 1.0.48.
			break;
		default:
			return_type_string[0] = _T(""); // It will be detected as invalid below.
			return_type_string[1] = NULL;
			break;
		}

		// 64-bit note: The calling convention detection code is preserved here for script compatibility.

		if (!_tcsnicmp(return_type_string[0], _T("CDecl"), 5)) // Alternate calling convention.
		{
#ifdef WIN32_PLATFORM
			dll_call_mode = DC_CALL_CDECL;
#endif
			return_type_string[0] = omit_leading_whitespace(return_type_string[0] + 5);
			if (!*return_type_string[0])
			{	// Take a shortcut since we know this empty string will be used as "Int":
				return_attrib.type = DLL_ARG_INT;
				goto has_valid_return_type;
			}
		}
		// This next part is a little iffy because if a legitimate return type is contained in a variable
		// that happens to be named Cdecl, Cdecl will be put into effect regardless of what's in the variable.
		// But the convenience of being able to omit the quotes around Cdecl seems to outweigh the extreme
		// rarity of such a thing happening.
		else if (return_type_string[1] && !_tcsnicmp(return_type_string[1], _T("CDecl"), 5)) // Alternate calling convention.
		{
#ifdef WIN32_PLATFORM
			dll_call_mode = DC_CALL_CDECL;
#endif
			return_type_string[1] += 5; // Support return type immediately following CDecl (this was previously supported _with_ quotes, though not documented).  OBSOLETE COMMENT: Must be NULL since return_type_string[1] is the variable's name, by definition, so it can't have any spaces in it, and thus no space delimited items after "Cdecl".
			if (!*return_type_string[1])
				// Pass default return type.  Don't take shortcut approach used above as return_type_string[0] should take precedence if valid.
				return_type_string[1] = _T("Int");
		}

		ConvertDllArgType(return_type_string, return_attrib);
		if (return_attrib.type == DLL_ARG_INVALID)
		{
			g_script.SetErrorLevelOrThrowStr(_T("-2"), _T("DllCall")); // Stage 2 error: Invalid return type or arg type.
			return;
		}
has_valid_return_type:
		--aParamCount;  // Remove the last parameter from further consideration.
#ifdef WIN32_PLATFORM
		if (!return_attrib.passed_by_address) // i.e. the special return flags below are not needed when an address is being returned.
		{
			if (return_attrib.type == DLL_ARG_DOUBLE)
				dll_call_mode |= DC_RETVAL_MATH8;
			else if (return_attrib.type == DLL_ARG_FLOAT)
				dll_call_mode |= DC_RETVAL_MATH4;
		}
#endif
	}

	// Using stack memory, create an array of dll args large enough to hold the actual number of args present.
	int arg_count = aParamCount/2; // Might provide one extra due to first/last params, which is inconsequential.
	DYNAPARM *dyna_param = arg_count ? (DYNAPARM *)_alloca(arg_count * sizeof(DYNAPARM)) : NULL;
	// Above: _alloca() has been checked for code-bloat and it doesn't appear to be an issue.
	// Above: Fix for v1.0.36.07: According to MSDN, on failure, this implementation of _alloca() generates a
	// stack overflow exception rather than returning a NULL value.  Therefore, NULL is no longer checked,
	// nor is an exception block used since stack overflow in this case should be exceptionally rare (if it
	// does happen, it would probably mean the script or the program has a design flaw somewhere, such as
	// infinite recursion).

	LPTSTR arg_type_string[2];
	int i = arg_count * sizeof(void *);
	// for Unicode <-> ANSI charset conversion
#ifdef UNICODE
	CStringA **pStr = (CStringA **)
#else
	CStringW **pStr = (CStringW **)
#endif
		_alloca(i); // _alloca vs malloc can make a significant difference to performance in some cases.
	memset(pStr, 0, i);

	// Above has already ensured that after the first parameter, there are either zero additional parameters
	// or an even number of them.  In other words, each arg type will have an arg value to go with it.
	// It has also verified that the dyna_param array is large enough to hold all of the args.
	for (arg_count = 0, i = 1; i < aParamCount; ++arg_count, i += 2)  // Same loop as used later below, so maintain them together.
	{
		switch (aParam[i]->symbol)
		{
		case SYM_VAR: // SYM_VAR's Type() is always VAR_NORMAL (except lvalues in expressions).
			arg_type_string[0] = aParam[i]->var->Contents(TRUE, TRUE);
			arg_type_string[1] = aParam[i]->var->mName;
			// v1.0.33.01: arg_type_string[1] improves convenience by falling back to the variable's name
			// if the contents are not appropriate.  In other words, both Int and "Int" are treated the same.
			// It's done this way to allow the variable named "Int" to actually contain some other legitimate
			// type-name such as "Str" (in case anyone ever happens to do that).
			break;
		case SYM_STRING:
		case SYM_OPERAND:
			arg_type_string[0] = aParam[i]->marker;
			arg_type_string[1] = NULL; // Added in 1.0.48.
			break;
		default:
			arg_type_string[0] = _T(""); // It will be detected as invalid below.
			arg_type_string[1] = NULL;
			break;
		}

		ExprTokenType &this_param = *aParam[i + 1];         // Resolved for performance and convenience.
		DYNAPARM &this_dyna_param = dyna_param[arg_count];  //

		// Store the each arg into a dyna_param struct, using its arg type to determine how.
		ConvertDllArgType(arg_type_string, this_dyna_param);
		switch (this_dyna_param.type)
		{
		case DLL_ARG_STR:
			if (IS_NUMERIC(this_param.symbol))
			{
				// For now, string args must be real strings rather than floats or ints.  An alternative
				// to this would be to convert it to number using persistent memory from the caller (which
				// is necessary because our own stack memory should not be passed to any function since
				// that might cause it to return a pointer to stack memory, or update an output-parameter
				// to be stack memory, which would be invalid memory upon return to the caller).
				// The complexity of this doesn't seem worth the rarity of the need, so this will be
				// documented in the help file.
				g_script.SetErrorLevelOrThrowStr(_T("-2"), _T("DllCall")); // Stage 2 error: Invalid return type or arg type.
				return;
			}
			// Otherwise, it's a supported type of string.
			this_dyna_param.ptr = TokenToString(this_param); // SYM_VAR's Type() is always VAR_NORMAL (except lvalues in expressions).
			// NOTES ABOUT THE ABOVE:
			// UPDATE: The v1.0.44.14 item below doesn't work in release mode, only debug mode (turning off
			// "string pooling" doesn't help either).  So it's commented out until a way is found
			// to pass the address of a read-only empty string (if such a thing is possible in
			// release mode).  Such a string should have the following properties:
			// 1) The first byte at its address should be '\0' so that functions can read it
			//    and recognize it as a valid empty string.
			// 2) The memory address should be readable but not writable: it should throw an
			//    access violation if the function tries to write to it (like "" does in debug mode).
			// SO INSTEAD of the following, DllCall() now checks further below for whether sEmptyString
			// has been overwritten/trashed by the call, and if so displays a warning dialog.
			// See note above about this: v1.0.44.14: If a variable is being passed that has no capacity, pass a
			// read-only memory area instead of a writable empty string. There are two big benefits to this:
			// 1) It forces an immediate exception (catchable by DllCall's exception handler) so
			//    that the program doesn't crash from memory corruption later on.
			// 2) It avoids corrupting the program's static memory area (because sEmptyString
			//    resides there), which can save many hours of debugging for users when the program
			//    crashes on some seemingly unrelated line.
			// Of course, it's not a complete solution because it doesn't stop a script from
			// passing a variable whose capacity is non-zero yet too small to handle what the
			// function will write to it.  But it's a far cry better than nothing because it's
			// common for a script to forget to call VarSetCapacity before passing a buffer to some
			// function that writes a string to it.
			//if (this_dyna_param.str == Var::sEmptyString) // To improve performance, compare directly to Var::sEmptyString rather than calling Capacity().
			//	this_dyna_param.str = _T(""); // Make it read-only to force an exception.  See comments above.
			break;
		case DLL_ARG_xSTR:
			// See the section above for comments.
			if (IS_NUMERIC(this_param.symbol))
			{
				g_script.SetErrorLevelOrThrowStr(_T("-2"), _T("DllCall"));
				return;
			}
			// String needing translation: ASTR on Unicode build, WSTR on ANSI build.
			pStr[arg_count] = new UorA(CStringCharFromWChar,CStringWCharFromChar)(TokenToString(this_param));
			this_dyna_param.ptr = pStr[arg_count]->GetBuffer();
			break;

		case DLL_ARG_DOUBLE:
		case DLL_ARG_FLOAT:
			// This currently doesn't validate that this_dyna_param.is_unsigned==false, since it seems
			// too rare and mostly harmless to worry about something like "Ufloat" having been specified.
			this_dyna_param.value_double = TokenToDouble(this_param);
			if (this_dyna_param.type == DLL_ARG_FLOAT)
				this_dyna_param.value_float = (float)this_dyna_param.value_double;
			break;

		case DLL_ARG_INVALID:
			if (aParam[i]->symbol == SYM_VAR)
				aParam[i]->var->MaybeWarnUninitialized();
			g_script.SetErrorLevelOrThrowStr(_T("-2"), _T("DllCall")); // Stage 2 error: Invalid return type or arg type.
			return;

		default: // Namely:
		//case DLL_ARG_INT:
		//case DLL_ARG_SHORT:
		//case DLL_ARG_CHAR:
		//case DLL_ARG_INT64:
			if (this_dyna_param.is_unsigned && this_dyna_param.type == DLL_ARG_INT64 && !IS_NUMERIC(this_param.symbol))
				// The above and below also apply to BIF_NumPut(), so maintain them together.
				// !IS_NUMERIC() is checked because such tokens are already signed values, so should be
				// written out as signed so that whoever uses them can interpret negatives as large
				// unsigned values.
				// Support for unsigned values that are 32 bits wide or less is done via ATOI64() since
				// it should be able to handle both signed and unsigned values.  However, unsigned 64-bit
				// values probably require ATOU64(), which will prevent something like -1 from being seen
				// as the largest unsigned 64-bit int; but more importantly there are some other issues
				// with unsigned 64-bit numbers: The script internals use 64-bit signed values everywhere,
				// so unsigned values can only be partially supported for incoming parameters, but probably
				// not for outgoing parameters (values the function changed) or the return value.  Those
				// should probably be written back out to the script as negatives so that other parts of
				// the script, such as expressions, can see them as signed values.  In other words, if the
				// script somehow gets a 64-bit unsigned value into a variable, and that value is larger
				// that LLONG_MAX (i.e. too large for ATOI64 to handle), ATOU64() will be able to resolve
				// it, but any output parameter should be written back out as a negative if it exceeds
				// LLONG_MAX (return values can be written out as unsigned since the script can specify
				// signed to avoid this, since they don't need the incoming detection for ATOU()).
				this_dyna_param.value_int64 = (__int64)ATOU64(TokenToString(this_param)); // Cast should not prevent called function from seeing it as an undamaged unsigned number.
			else
				this_dyna_param.value_int64 = TokenToInt64(this_param);

			// Values less than or equal to 32-bits wide always get copied into a single 32-bit value
			// because they should be right justified within it for insertion onto the call stack.
			if (this_dyna_param.type != DLL_ARG_INT64) // Shift the 32-bit value into the high-order DWORD of the 64-bit value for later use by DynaCall().
				this_dyna_param.value_int = (int)this_dyna_param.value_int64; // Force a failure if compiler generates code for this that corrupts the union (since the same method is used for the more obscure float vs. double below).
		} // switch (this_dyna_param.type)
	} // for() each arg.
    
	if (!function) // The function's address hasn't yet been determined.
	{
		function = GetDllProcAddress(TokenToString(*aParam[0]), &hmodule_to_free);
		if (!function)
			goto end;
	}

	////////////////////////
	// Call the DLL function
	////////////////////////
	DWORD exception_occurred; // Must not be named "exception_code" to avoid interfering with MSVC macros.
	DYNARESULT return_value;  // Doing assignment (below) as separate step avoids compiler warning about "goto end" skipping it.
#ifdef WIN32_PLATFORM
	return_value = DynaCall(dll_call_mode, function, dyna_param, arg_count, exception_occurred, NULL, 0);
#endif
#ifdef _WIN64
	return_value = DynaCall(function, dyna_param, arg_count, exception_occurred);
#endif
	// The above has also set g_ErrorLevel appropriately.

	if (*Var::sEmptyString)
	{
		// v1.0.45.01 Above has detected that a variable of zero capacity was passed to the called function
		// and the function wrote to it (assuming sEmptyString wasn't already trashed some other way even
		// before the call).  So patch up the empty string to stabilize a little; but it's too late to
		// salvage this instance of the program because there's no knowing how much static data adjacent to
		// sEmptyString has been overwritten and corrupted.
		*Var::sEmptyString = '\0';
		// Don't bother with freeing hmodule_to_free since a critical error like this calls for minimal cleanup.
		// The OS almost certainly frees it upon termination anyway.
		// Call CriticalError() so that the user knows *which* DllCall is at fault:
		g_script.CriticalError(_T("This DllCall requires a prior VarSetCapacity."));
		// CriticalError always terminates the process.
	}

	// It seems best to have the above take precedence over "exception_occurred" below.
	if (exception_occurred)
	{
		// If the called function generated an exception, I think it's impossible for the return value
		// to be valid/meaningful since it the function never returned properly.  Confirmation of this
		// would be good, but in the meantime it seems best to make the return value an empty string as
		// an indicator that the call failed (in addition to ErrorLevel).
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
		// But continue on to write out any output parameters because the called function might have
		// had a chance to update them before aborting.
	}
	else // The call was successful.  Interpret and store the return value.
	{
		// If the return value is passed by address, dereference it here.
		if (return_attrib.passed_by_address)
		{
			return_attrib.passed_by_address = false; // Because the address is about to be dereferenced/resolved.

			switch(return_attrib.type)
			{
			case DLL_ARG_INT64:
			case DLL_ARG_DOUBLE:
#ifdef _WIN64 // fincs: pointers are 64-bit on x64.
			case DLL_ARG_WSTR:
			case DLL_ARG_ASTR:
#endif
				// Same as next section but for eight bytes:
				return_value.Int64 = *(__int64 *)return_value.Pointer;
				break;
			default: // Namely:
			//case DLL_ARG_STR:  // Even strings can be passed by address, which is equivalent to "char **".
			//case DLL_ARG_INT:
			//case DLL_ARG_SHORT:
			//case DLL_ARG_CHAR:
			//case DLL_ARG_FLOAT:
				// All the above are stored in four bytes, so a straight dereference will copy the value
				// over unchanged, even if it's a float.
				return_value.Int = *(int *)return_value.Pointer;
			}
		}
#ifdef _WIN64
		else
		{
			switch(return_attrib.type)
			{
			// Floating-point values are returned via the xmm0 register. Copy it for use in the next section:
			case DLL_ARG_FLOAT:
				return_value.Float = read_xmm0_float();
				break;
			case DLL_ARG_DOUBLE:
				return_value.Double = read_xmm0_double();
				break;
			}
		}
#endif

		switch(return_attrib.type)
		{
		case DLL_ARG_INT: // Listed first for performance. If the function has a void return value (formerly DLL_ARG_NONE), the value assigned here is undefined and inconsequential since the script should be designed to ignore it.
			aResultToken.symbol = SYM_INTEGER;
			if (return_attrib.is_unsigned)
				aResultToken.value_int64 = (UINT)return_value.Int; // Preserve unsigned nature upon promotion to signed 64-bit.
			else // Signed.
				aResultToken.value_int64 = return_value.Int;
			break;
		case DLL_ARG_STR:
			// The contents of the string returned from the function must not reside in our stack memory since
			// that will vanish when we return to our caller.  As long as every string that went into the
			// function isn't on our stack (which is the case), there should be no way for what comes out to be
			// on the stack either.
			//aResultToken.symbol = SYM_STRING; // This is the default.
			aResultToken.marker = (LPTSTR)(return_value.Pointer ? return_value.Pointer : _T(""));
			// Above: Fix for v1.0.33.01: Don't allow marker to be set to NULL, which prevents crash
			// with something like the following, which in this case probably happens because the inner
			// call produces a non-numeric string, which "int" then sees as zero, which CharLower() then
			// sees as NULL, which causes CharLower to return NULL rather than a real string:
			//result := DllCall("CharLower", "int", DllCall("CharUpper", "str", MyVar, "str"), "str")
			break;
		case DLL_ARG_xSTR:
			{	// String needing translation: ASTR on Unicode build, WSTR on ANSI build.
#ifdef UNICODE
				LPCSTR result = (LPCSTR)return_value.Pointer;
#else
				LPCWSTR result = (LPCWSTR)return_value.Pointer;
#endif
				if (result && *result)
				{
#ifdef UNICODE		// Perform the translation:
					CStringWCharFromChar result_buf(result);
#else
					CStringCharFromWChar result_buf(result);
#endif
					// Store the length of the translated string first since DetachBuffer() clears it.
					aResultToken.marker_length = result_buf.GetLength();
					// Now attempt to take ownership of the malloc'd memory, to return to our caller.
					if (aResultToken.mem_to_free = result_buf.DetachBuffer())
						aResultToken.marker = aResultToken.mem_to_free;
					//else mem_to_free is NULL, so marker_length should be ignored.  See next comment below.
				}
				//else leave aResultToken as it was set at the top of this function: an empty string.
			}
			break;
		case DLL_ARG_SHORT:
			aResultToken.symbol = SYM_INTEGER;
			if (return_attrib.is_unsigned)
				aResultToken.value_int64 = return_value.Int & 0x0000FFFF; // This also forces the value into the unsigned domain of a signed int.
			else // Signed.
				aResultToken.value_int64 = (SHORT)(WORD)return_value.Int; // These casts properly preserve negatives.
			break;
		case DLL_ARG_CHAR:
			aResultToken.symbol = SYM_INTEGER;
			if (return_attrib.is_unsigned)
				aResultToken.value_int64 = return_value.Int & 0x000000FF; // This also forces the value into the unsigned domain of a signed int.
			else // Signed.
				aResultToken.value_int64 = (char)(BYTE)return_value.Int; // These casts properly preserve negatives.
			break;
		case DLL_ARG_INT64:
			// Even for unsigned 64-bit values, it seems best both for simplicity and consistency to write
			// them back out to the script as signed values because script internals are not currently
			// equipped to handle unsigned 64-bit values.  This has been documented.
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = return_value.Int64;
			break;
		case DLL_ARG_FLOAT:
			aResultToken.symbol = SYM_FLOAT;
			aResultToken.value_double = return_value.Float;
			break;
		case DLL_ARG_DOUBLE:
			aResultToken.symbol = SYM_FLOAT; // There is no SYM_DOUBLE since all floats are stored as doubles.
			aResultToken.value_double = return_value.Double;
			break;
		//default: // Should never be reached unless there's a bug.
		//	aResultToken.symbol = SYM_STRING;
		//	aResultToken.marker = "";
		} // switch(return_attrib.type)
	} // Storing the return value when no exception occurred.

	// Store any output parameters back into the input variables.  This allows a function to change the
	// contents of a variable for the following arg types: String and Pointer to <various number types>.
	for (arg_count = 0, i = 1; i < aParamCount; ++arg_count, i += 2) // Same loop as used above, so maintain them together.
	{
		ExprTokenType &this_param = *aParam[i + 1];  // Resolved for performance and convenience.
		if (this_param.symbol != SYM_VAR) // Output parameters are copied back only if its counterpart parameter is a naked variable.
		{
			if (pStr[arg_count]) // We don't need to copy it back, so delete it.
				delete pStr[arg_count];
			continue;
		}
		DYNAPARM &this_dyna_param = dyna_param[arg_count]; // Resolved for performance and convenience.
		Var &output_var = *this_param.var;                 //
		if (this_dyna_param.type == DLL_ARG_STR) // Native string type for current build config.
		{
			LPTSTR contents = output_var.Contents(); // Contents() shouldn't update mContents in this case because Contents() was already called for each "str" parameter prior to calling the Dll function.
			VarSizeType capacity = output_var.Capacity();
			// Since the performance cost is low, ensure the string is terminated at the limit of its
			// capacity (helps prevent crashes if DLL function didn't do its job and terminate the string,
			// or when a function is called that deliberately doesn't terminate the string, such as
			// RtlMoveMemory()).
			if (capacity)
				contents[capacity - 1] = '\0';
			// The function might have altered Contents(), so update Length().
			output_var.SetCharLength((VarSizeType)_tcslen(contents));
			output_var.Close(); // Clear the attributes of the variable to reflect the fact that the contents may have changed.
			continue;
		}
		if (this_dyna_param.type == DLL_ARG_xSTR) // String needing translation: ASTR on Unicode build, WSTR on ANSI build.
		{
			pStr[arg_count]->ReleaseBuffer();
#ifdef UNICODE
			output_var.AssignStringFromCodePage(
#else
			output_var.AssignStringToCodePage(
#endif
				pStr[arg_count]->GetString() );
			delete pStr[arg_count];
			continue;
		}

		// Since above didn't "continue", this arg wasn't passed as a string.  Of the remaining types, only
		// those passed by address can possibly be output parameters, so skip the rest:
		if (!this_dyna_param.passed_by_address)
			continue;

		switch (this_dyna_param.type)
		{
		// case DLL_ARG_STR:  Already handled above.
		case DLL_ARG_INT:
			if (this_dyna_param.is_unsigned)
				output_var.Assign((DWORD)this_dyna_param.value_int);
			else // Signed.
				output_var.Assign(this_dyna_param.value_int);
			break;
		case DLL_ARG_SHORT:
			if (this_dyna_param.is_unsigned) // Force omission of the high-order word in case it is non-zero from a parameter that was originally and erroneously larger than a short.
				output_var.Assign(this_dyna_param.value_int & 0x0000FFFF); // This also forces the value into the unsigned domain of a signed int.
			else // Signed.
				output_var.Assign((int)(SHORT)(WORD)this_dyna_param.value_int); // These casts properly preserve negatives.
			break;
		case DLL_ARG_CHAR:
			if (this_dyna_param.is_unsigned) // Force omission of the high-order bits in case it is non-zero from a parameter that was originally and erroneously larger than a char.
				output_var.Assign(this_dyna_param.value_int & 0x000000FF); // This also forces the value into the unsigned domain of a signed int.
			else // Signed.
				output_var.Assign((int)(char)(BYTE)this_dyna_param.value_int); // These casts properly preserve negatives.
			break;
		case DLL_ARG_INT64: // Unsigned and signed are both written as signed for the reasons described elsewhere above.
			output_var.Assign(this_dyna_param.value_int64);
			break;
		case DLL_ARG_FLOAT:
			output_var.Assign(this_dyna_param.value_float);
			break;
		case DLL_ARG_DOUBLE:
			output_var.Assign(this_dyna_param.value_double);
			break;
		}
	}

end:
	if (hmodule_to_free)
		FreeLibrary(hmodule_to_free);
}

#endif


BIF_DECL(BIF_StrLen)
// Caller has ensured that SYM_VAR's Type() is VAR_NORMAL and that it's either not an environment
// variable or the caller wants environment variables treated as having zero length.
// Result is always an integer (caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need
// to set it here).
{
	// Loadtime validation has ensured that there's exactly one actual parameter.
	// Calling Length() is always valid for SYM_VAR because SYM_VAR's Type() is always VAR_NORMAL (except lvalues in expressions).
	aResultToken.value_int64 = (aParam[0]->symbol == SYM_VAR)
		? (aParam[0]->var->MaybeWarnUninitialized(), aParam[0]->var->Length())
		: _tcslen(TokenToString(*aParam[0], aResultToken.buf));  // Allow StrLen(numeric_expr) for flexibility.
}



BIF_DECL(BIF_SubStr) // Added in v1.0.46.
{
	// Set default return value in case of early return.
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");

	// Get the first arg, which is the string used as the source of the extraction. Call it "haystack" for clarity.
	TCHAR haystack_buf[MAX_NUMBER_SIZE]; // A separate buf because aResultToken.buf is sometimes used to store the result.
	LPTSTR haystack = ParamIndexToString(0, haystack_buf); // Remember that aResultToken.buf is part of a union, though in this case there's no danger of overwriting it since our result will always be of STRING type (not int or float).
	INT_PTR haystack_length = (INT_PTR)ParamIndexLength(0, haystack);

	// Load-time validation has ensured that at least the first two parameters are present:
	INT_PTR starting_offset = (INT_PTR)ParamIndexToInt64(1) - 1; // The one-based starting position in haystack (if any).  Convert it to zero-based.
	if (starting_offset > haystack_length)
		return; // Yield the empty string (a default set higher above).
	if (starting_offset < 0) // Same convention as RegExMatch/Replace(): Treat a StartingPos of 0 (offset -1) as "start at the string's last char".  Similarly, treat negatives as starting further to the left of the end of the string.
	{
		starting_offset += haystack_length;
		if (starting_offset < 0)
			starting_offset = 0;
	}

	INT_PTR remaining_length_available = haystack_length - starting_offset;
	INT_PTR extract_length;
	if (aParamCount < 3) // No length specified, so extract all the remaining length.
		extract_length = remaining_length_available;
	else
	{
		if (   !(extract_length = (INT_PTR)ParamIndexToInt64(2))   )  // It has asked to extract zero characters.
			return; // Yield the empty string (a default set higher above).
		if (extract_length < 0)
		{
			extract_length += remaining_length_available; // Result is the number of characters to be extracted (i.e. after omitting the number of chars specified in extract_length).
			if (extract_length < 1) // It has asked to omit all characters.
				return; // Yield the empty string (a default set higher above).
		}
		else // extract_length > 0
			if (extract_length > remaining_length_available)
				extract_length = remaining_length_available;
	}

	// Above has set extract_length to the exact number of characters that will actually be extracted.
	LPTSTR result = haystack + starting_offset; // This is the result except for the possible need to truncate it below.

	if (extract_length == remaining_length_available) // All of haystack is desired (starting at starting_offset).
	{
		aResultToken.marker = result; // No need for any copying or termination, just send back part of haystack.
		return;                       // Caller and Var:Assign() know that overlap is possible, so this seems safe.
	}
	
	// Otherwise, at least one character is being omitted from the end of haystack.
	if (!TokenSetResult(aResultToken, result, extract_length))
	{
		// Yield the empty string (a default set higher above).
	}
}



BIF_DECL(BIF_InStr)
{
	// Load-time validation has already ensured that at least two actual parameters are present.
	TCHAR needle_buf[MAX_NUMBER_SIZE];
	LPTSTR haystack = ParamIndexToString(0, aResultToken.buf);
	LPTSTR needle = ParamIndexToString(1, needle_buf);
	// Result type will always be an integer:
	// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.

	// v1.0.43.03: Rather than adding a third value to the CaseSensitive parameter, it seems better to
	// obey StringCaseSense because:
	// 1) It matches the behavior of the equal operator (=) in expressions.
	// 2) It's more friendly for typical international uses because it avoids having to specify that special/third value
	//    for every call of InStr.  It's nice to be able to omit the CaseSensitive parameter every time and know that
	//    the behavior of both InStr and its counterpart the equals operator are always consistent with each other.
	// 3) Avoids breaking existing scripts that may pass something other than true/false for the CaseSense parameter.
	StringCaseSenseType string_case_sense = (StringCaseSenseType)(!ParamIndexIsOmitted(2) && ParamIndexToInt64(2));
	// Above has assigned SCS_INSENSITIVE (0) or SCS_SENSITIVE (1).  If it's insensitive, resolve it to
	// be Locale-mode if the StringCaseSense mode is either case-sensitive or Locale-insensitive.
	if (g->StringCaseSense != SCS_INSENSITIVE && string_case_sense == SCS_INSENSITIVE) // Ordered for short-circuit performance.
		string_case_sense = SCS_INSENSITIVE_LOCALE;

	LPTSTR found_pos;
	INT_PTR offset = 0; // Set default.
	int occurrence_number = ParamIndexToOptionalInt(4, 1);

	if (!ParamIndexIsOmitted(3)) // There is a starting position present.
	{
		offset = ParamIndexToIntPtr(3); // i.e. the fourth arg.
		// For offset validation and reverse search we need to know the length of haystack:
		INT_PTR haystack_length = ParamIndexLength(0, haystack);
		if (offset <= 0) // Special mode to search from the right side.
		{
			haystack_length += offset; // i.e. reduce haystack_length by the absolute value of offset.
			found_pos = (haystack_length >= 0) ? tcsrstr(haystack, haystack_length, needle, string_case_sense, occurrence_number) : NULL;
			aResultToken.value_int64 = found_pos ? (found_pos - haystack + 1) : 0;  // +1 to convert to 1-based, since 0 indicates "not found".
			return;
		}
		--offset; // Convert from one-based to zero-based.
		if (offset > haystack_length || occurrence_number < 1)
		{
			aResultToken.value_int64 = 0; // Match never found when offset is beyond length of string.
			return;
		}
	}
	// Since above didn't return:
	size_t needle_length = (occurrence_number > 1) ? ParamIndexLength(1, needle) : 1; // Avoid unnecessary _tcslen() if occurrence_number == 1, which is the most common case.
	int i;
	for (i = 1, found_pos = haystack + offset; ; ++i, found_pos += needle_length)
		if (!(found_pos = tcsstr2(found_pos, needle, string_case_sense)) || i == occurrence_number)
			break;
	aResultToken.value_int64 = found_pos ? (found_pos - haystack + 1) : 0;
}


void RegExSetSubpatternVars(LPCTSTR haystack, pcret *re, pcret_extra *extra, TCHAR output_mode, Var &output_var, int *offset, int pattern_count, int captured_pattern_count, LPTSTR &mem_to_free)
{
	// OTHERWISE, CONTINUE ON TO STORE THE SUBSTRINGS THAT MATCHED THE SUBPATTERNS (EVEN IF PCRE_ERROR_NOMATCH).
	// For lookup performance, create a table of subpattern names indexed by subpattern number.
	LPCTSTR *subpat_name = NULL; // Set default as "no subpattern names present or available".
	bool allow_dupe_subpat_names = false; // Set default.
	LPCTSTR name_table;
	int name_count, name_entry_size;
	if (   !pcret_fullinfo(re, extra, PCRE_INFO_NAMECOUNT, &name_count) // Success. Fix for v1.0.45.01: Don't check captured_pattern_count>=0 because PCRE_ERROR_NOMATCH can still have named patterns!
		&& name_count // There's at least one named subpattern.  Relies on short-circuit boolean order.
		&& !pcret_fullinfo(re, extra, PCRE_INFO_NAMETABLE, &name_table) // Success.
		&& !pcret_fullinfo(re, extra, PCRE_INFO_NAMEENTRYSIZE, &name_entry_size)   ) // Success.
	{
		int pcre_options;
		if (!pcret_fullinfo(re, extra, PCRE_INFO_OPTIONS, &pcre_options)) // Success.
			allow_dupe_subpat_names = pcre_options & PCRE_DUPNAMES;
		// For indexing simplicity, also include an entry for the main/entire pattern at index 0 even though
		// it's never used because the entire pattern can't have a name without enclosing it in parentheses
		// (in which case it's not the entire pattern anymore, but in fact subpattern #1).
		size_t subpat_array_size = pattern_count * sizeof(LPCSTR);
		subpat_name = (LPCTSTR *)_alloca(subpat_array_size); // See other use of _alloca() above for reasons why it's used.
		ZeroMemory(subpat_name, subpat_array_size); // Set default for each index to be "no name corresponds to this subpattern number".
		for (int i = 0; i < name_count; ++i, name_table += name_entry_size)
		{
			// Below converts first two bytes of each name-table entry into the pattern number (it might be
			// possible to simplify this, but I'm not sure if big vs. little-endian will ever be a concern).
#ifdef UNICODE
			subpat_name[name_table[0]] = name_table + 1;
#else
			subpat_name[(name_table[0] << 8) + name_table[1]] = name_table + 2; // For indexing simplicity, subpat_name[0] is for the main/entire pattern though it is never actually used for that because it can't be named without being enclosed in parentheses (in which case it becomes a subpattern).
#endif
			// For simplicity and unlike PHP, IsPureNumeric() isn't called to forbid numeric subpattern names.
			// It seems the worst than could happen if it is numeric is that it would overlap/overwrite some of
			// the numerically-indexed elements in the output-array.  Seems pretty harmless given the rarity.
		}
	}
	//else one of the pcre_fullinfo() calls may have failed.  The PCRE docs indicate that this realistically never
	// happens unless bad inputs were given.  So due to rarity, just leave subpat_name==NULL; i.e. "no named subpatterns".

	if (output_mode == 'O')
	{
		LPTSTR mark = (extra->flags & PCRE_EXTRA_MARK) ? (LPTSTR)*extra->mark : NULL;
		IObject *m = RegExMatchObject::Create(haystack, offset, subpat_name, pattern_count, captured_pattern_count, mark);
		if (m)
			output_var.AssignSkipAddRef(m);
		else
			output_var.Assign();
		return;
	}

	// Make var_name longer than Max so that FindOrAddVar() will be able to spot and report var names
	// that are too long, either because the base-name is too long, or the name becomes too long
	// as a result of appending the array index number:
	TCHAR var_name[MAX_VAR_NAME_LENGTH + 68]; // Allow +3 extra for "Len" and "Pos" suffixes, +1 for terminator, and +64 for largest sub-pattern name (actually it's 32, but 64 allows room for future expansion).  64 is also enough room for the largest 64-bit integer, 20 chars: 18446744073709551616
	_tcscpy(var_name, output_var.mName); // This prefix is copied in only once, for performance.
	size_t suffix_length, prefix_length = _tcslen(var_name);
	LPTSTR var_name_suffix = var_name + prefix_length; // The position at which to copy the sequence number (index).
	int always_use = FINDVAR_FOR_PSEUDO_ARRAY(output_var);
	int n, p = 1, *this_offset = offset + 2; // Init for both loops below.
	Var *array_item;
	bool subpat_not_matched;
	int subpat_pos, subpat_len;

	if (output_mode == 'P')
	{
		for (; p < pattern_count; ++p, this_offset += 2) // Start at 1 because above already did pattern #0 (the full pattern).
		{
			subpat_not_matched = (p >= captured_pattern_count || this_offset[0] < 0); // See comments in similar section below about this.
			if (subpat_not_matched)
			{
				subpat_pos = 0;
				subpat_len = 0;
			}
			else
			{
				subpat_pos = this_offset[0];
				subpat_len = this_offset[1] - subpat_pos;
				++subpat_pos; // One-based (i.e. position zero means "not found").
			}

			if (subpat_name && subpat_name[p]) // This subpattern number has a name, so store it under that name.
			{
				if (*subpat_name[p]) // This check supports allow_dupe_subpat_names. See comments below.
				{
					const LPCTSTR &the_subpat_name = subpat_name[p];
					suffix_length = _stprintf(var_name_suffix, _T("Pos%s"), the_subpat_name); // Append the subpattern to the array's base name.
					if (array_item = g_script.FindOrAddVar(var_name, prefix_length + suffix_length, always_use))
						array_item->Assign(subpat_pos);
					suffix_length = _stprintf(var_name_suffix, _T("Len%s"), the_subpat_name); // Append the subpattern name to the array's base name.
					if (array_item = g_script.FindOrAddVar(var_name, prefix_length + suffix_length, always_use))
						array_item->Assign(subpat_len); // It seemed more convenient for scripts to store Length instead of an ending offset.
					// Fix for v1.0.45.01: Section below added.  See similar section further below for comments.
					if (!subpat_not_matched && allow_dupe_subpat_names) // Explicitly check subpat_not_matched not pos/len so that behavior is consistent with the default mode (non-position).
						for (n = p + 1; n < pattern_count; ++n) // Search to the right of this subpat to find others with the same name.
							if (subpat_name[n] && !_tcsicmp(subpat_name[n], subpat_name[p])) // Case-insensitive because unlike PCRE, named subpatterns conform to AHK convention of insensitive variable names.
								subpat_name[n] = _T(""); // Empty string signals subsequent iterations to skip it entirely.
				}
				//else an empty subpat name caused by "allow duplicate names".  Do nothing (see comments above).
			}
			else // This subpattern has no name, so write it out as its pattern number instead. For performance and memory utilization, it seems best to store only one or the other (named or number), not both.
			{
				// For comments about this section, see the similar for-loop later below.
				suffix_length = _stprintf(var_name_suffix, _T("Pos%d"), p); // Append the element number to the array's base name.
				if (array_item = g_script.FindOrAddVar(var_name, prefix_length + suffix_length, always_use))
					array_item->Assign(subpat_pos);
				//else var couldn't be created: no error reporting currently, since it basically should never happen.
				suffix_length = _stprintf(var_name_suffix, _T("Len%d"), p); // Append the element number to the array's base name.
				if (array_item = g_script.FindOrAddVar(var_name, prefix_length + suffix_length, always_use))
					array_item->Assign(subpat_len);
			}
		}
		//goto free_and_return;
		return;
	} // if (output_mode == 'P')

	// Otherwise, we're in get-substring mode (not offset mode), so store the substring that matches each subpattern.
	for (; p < pattern_count; ++p, this_offset += 2) // Start at 1 because above already did pattern #0 (the full pattern).
	{
		// If both items in this_offset are -1, that means the substring wasn't populated because it's
		// subpattern wasn't needed to find a match (or there was no match for *anything*).  For example:
		// "(xyz)|(abc)" (in which only one is subpattern will match).
		// NOTE: PCRE isn't clear on this, but it seems likely that captured_pattern_count
		// (returned from pcre_exec()) can be less than pattern_count (from pcre_fullinfo/
		// PCRE_INFO_CAPTURECOUNT).  So the below takes this into account by not trusting values
		// in offset[] that are beyond captured_pattern_count.  Further evidence of this is PCRE's
		// pcre_copy_substring() function, which consults captured_pattern_count to decide whether to
		// consult the offset array. The formula below works even if captured_pattern_count==PCRE_ERROR_NOMATCH.
		subpat_not_matched = (p >= captured_pattern_count || this_offset[0] < 0); // Relies on short-circuit boolean order.

		if (subpat_name && subpat_name[p]) // This subpattern number has a name, so store it under that name.
		{
			if (*subpat_name[p]) // This check supports allow_dupe_subpat_names. See comments below.
			{
				// This section is similar to the one in the "else" below, so see it for more comments.
				_tcscpy(var_name_suffix, subpat_name[p]); // Append the subpat name to the array's base name.  _tcscpy() seems safe because PCRE almost certainly enforces the 32-char limit on subpattern names.
				if (array_item = g_script.FindOrAddVar(var_name, 0, always_use))
				{
					if (p < pattern_count-1 // i.e. there's at least one more subpattern after this one (if there weren't, making a copy of haystack wouldn't be necessary because overlap can't harm this final assignment).
						&& haystack == array_item->Contents(FALSE)) // For more comments, see similar section in BIF_RegEx.
						if (mem_to_free = _tcsdup(haystack))
							haystack = mem_to_free;

					if (subpat_not_matched)
						array_item->Assign(); // Omit all parameters to make the var empty without freeing its memory (for performance, in case this RegEx is being used many times in a loop).
					else
					{
						subpat_pos = this_offset[0];
						subpat_len = this_offset[1] - subpat_pos;
						array_item->Assign(haystack + subpat_pos, subpat_len);
						// Fix for v1.0.45.01: When the J option (allow duplicate named subpatterns) is in effect,
						// PCRE returns entries for all the duplicates.  But we don't want an unmatched duplicate
						// to overwrite a previously matched duplicate.  To prevent this, when we're here (i.e.
						// this subpattern matched something), mark duplicate entries in the names array that lie
						// to the right of this item to indicate that they should be skipped by subsequent iterations.
						if (allow_dupe_subpat_names)
							for (n = p + 1; n < pattern_count; ++n) // Search to the right of this subpat to find others with the same name.
								if (subpat_name[n] && !_tcsicmp(subpat_name[n], subpat_name[p])) // Case-insensitive because unlike PCRE, named subpatterns conform to AHK convention of insensitive variable names.
									subpat_name[n] = _T(""); // Empty string signals subsequent iterations to skip it entirely.
					}
				}
				//else var couldn't be created: no error reporting currently, since it basically should never happen.
			}
			//else an empty subpat name caused by "allow duplicate names".  Do nothing (see comments above).
		}
		else // This subpattern has no name, so instead write it out as its actual pattern number. For performance and memory utilization, it seems best to store only one or the other (named or number), not both.
		{
			_itot(p, var_name_suffix, 10); // Append the element number to the array's base name.
			// To help performance (in case the linked list of variables is huge), tell it where
			// to start the search.  Use the base array name rather than the preceding element because,
			// for example, Array19 is alphabetically less than Array2, so we can't rely on the
			// numerical ordering:
			if (array_item = g_script.FindOrAddVar(var_name, 0, always_use))
			{
				if (p < pattern_count-1 // i.e. there's at least one more subpattern after this one (if there weren't, making a copy of haystack wouldn't be necessary because overlap can't harm this final assignment).
					&& haystack == array_item->Contents(FALSE)) // For more comments, see similar section in BIF_RegEx.
					if (mem_to_free = _tcsdup(haystack))
						haystack = mem_to_free;
				if (subpat_not_matched)
					array_item->Assign(); // Omit all parameters to make the var empty without freeing its memory (for performance, in case this RegEx is being used many times in a loop).
				else
				{
					subpat_pos = this_offset[0];
					subpat_len = this_offset[1] - subpat_pos;
					array_item->Assign(haystack + subpat_pos, subpat_len);
				}
			}
			//else var couldn't be created: no error reporting currently, since it basically should never happen.
		}
	} // for() each subpattern.
}


RegExMatchObject *RegExMatchObject::Create(LPCTSTR aHaystack, int *aOffset, LPCTSTR *aPatternName
	, int aPatternCount, int aCapturedPatternCount, LPCTSTR aMark)
{
	// If there was no match, seems best to not return an object:
	if (aCapturedPatternCount < 1)
		return NULL;

	RegExMatchObject *m = new RegExMatchObject();
	if (!m)
		return NULL;

	if (  aMark && !(m->mMark = _tcsdup(aMark))  )
	{
		m->Release();
		return NULL;
	}

	ASSERT(aCapturedPatternCount >= 1);
	ASSERT(aPatternCount >= aCapturedPatternCount);

	// Use aPatternCount vs aCapturedPatternCount since we want to be able to retrieve the
	// names of *all* subpatterns, even ones that weren't captured.  For instance, a loop
	// converting the object to an old-style pseudo-array would need to initialize even the
	// array items that weren't captured.
	m->mPatternCount = aPatternCount;
	
	// Allocate memory for a copy of the offset array.
	if (  !(m->mOffset = (int *)malloc(aPatternCount * 2 * sizeof(int *)))  )
	{
		m->Release();
		return NULL;
	}
	// memcpy currently benchmarks slightly faster on x64 than copying offsets in the loop below:
	memcpy(m->mOffset, aOffset, aPatternCount * 2 * sizeof(int));

	// Do some pre-processing:
	//  - Locate the smallest portion of haystack that contains all matches.
	//  - Convert end offsets to lengths.
	int p, min_offset = INT_MAX, max_offset = -1;
	for (p = 0; p < aCapturedPatternCount; ++p)
	{
		if (m->mOffset[p*2] > -1)
		{
			// Substring is non-empty, so ensure we copy this portion of haystack.
			if (min_offset > m->mOffset[p*2])
				min_offset = m->mOffset[p*2];
			if (max_offset < m->mOffset[p*2+1])
				max_offset = m->mOffset[p*2+1];
		}
		// Convert end offset to length.
		m->mOffset[p*2+1] -= m->mOffset[p*2];
	}
	// Initialize the remainder of the offset vector (patterns which were not captured),
	// which have indeterminate values if we're called by a regex callout and therefore
	// can't be handled the same way as in the loop above.
	for ( ; p < aPatternCount; ++p)
	{
		m->mOffset[p*2] = -1;
		m->mOffset[p*2+1] = 0;
	}
	
	// Copy only the portion of aHaystack which contains matches.  This can be much faster
	// than copying the whole string for larger haystacks.  For instance, searching for "GNU"
	// in the GPL v2 (18120 chars) and producing a match object is about 5 times faster with
	// this optimization than without if caller passes us the haystack length, and about 50
	// times faster than the old code which used _tcsdup().  However, the difference is much
	// smaller for small strings.
	if (min_offset < max_offset) // There are non-empty matches.
	{
		int our_haystack_size = (max_offset - min_offset) + 1;
		if (  !(m->mHaystack = tmalloc(our_haystack_size))  )
		{
			m->Release();
			return NULL;
		}
		tmemcpy(m->mHaystack, aHaystack + min_offset, our_haystack_size);
		m->mHaystackStart = min_offset;
	}

	// Copy subpattern names.
	if (aPatternName)
	{
		// Allocate array of pointers.
		if (  !(m->mPatternName = (LPTSTR *)malloc(aPatternCount * sizeof(LPTSTR *)))  )
		{
			m->Release();
			return NULL;
		}

		// Copy names and initialize array.
		m->mPatternName[0] = NULL;
		for (int p = 1; p < aPatternCount; ++p)
			if (aPatternName[p])
				// A failed allocation here seems rare and the consequences would be
				// negligible, so in that case just act as if the subpattern has no name.
				m->mPatternName[p] = _tcsdup(aPatternName[p]);
			else
				m->mPatternName[p] = NULL;
	}

	// Since above didn't return, the object has been set up successfully.
	return m;
}

ResultType STDMETHODCALLTYPE RegExMatchObject::Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount < 1 || aParamCount > 2 || IS_INVOKE_SET)
		return INVOKE_NOT_HANDLED;

	LPTSTR name;
	int p = -1;

	// Check for a subpattern offset/name first so that a subpattern named "Pos" takes
	// precedence over our "Pos" property when invoked like m.Pos (but not m.Pos()).
	if (aParamCount > 1 || !IS_INVOKE_CALL)
	{
		ExprTokenType &name_param = *aParam[aParamCount - 1];

		if (TokenIsPureNumeric(name_param))
		{
			p = (int)TokenToInt64(name_param, TRUE);
		}
		else if (mPatternName) // i.e. there is at least one named subpattern.
		{
			name = TokenToString(name_param);
			for (p = 0; p < mPatternCount; ++p)
				if (mPatternName[p] && !_tcsicmp(mPatternName[p], name))
				{
					if (mOffset[2*p] < 0)
						// This pattern wasn't matched, so check for one with a duplicate name.
						for (int i = p + 1; i < mPatternCount; ++i)
							if (mPatternName[i] && !_tcsicmp(mPatternName[i], name) // It has the same name.
								&& mOffset[2*i] >= 0) // It matched something.
							{
								// Prefer this pattern.
								p = i;
								break;
							}
					break;
				}
		}
	}

	bool pattern_found = p >= 0 && p < mPatternCount;
	
	// Checked for named properties:
	if (aParamCount > 1 || !pattern_found)
	{
		name = TokenToString(*aParam[0]);

		if (!pattern_found && aParamCount == 1)
		{
			p = 0; // For m.Pos, m.Len and m.Value, use the overall match.
			pattern_found = true; // Relies on below returning if the property name is invalid.
		}

		if (!_tcsicmp(name, _T("Pos")))
		{
			if (pattern_found)
			{
				aResultToken.symbol = SYM_INTEGER;
				aResultToken.value_int64 = mOffset[2*p] + 1;
			}
			return OK;
		}
		else if (!_tcsicmp(name, _T("Len")))
		{
			if (pattern_found)
			{
				aResultToken.symbol = SYM_INTEGER;
				aResultToken.value_int64 = mOffset[2*p + 1];
			}
			return OK;
		}
		else if (!_tcsicmp(name, _T("Count")))
		{
			if (aParamCount == 1)
			{
				aResultToken.symbol = SYM_INTEGER;
				aResultToken.value_int64 = mPatternCount - 1; // Return number of subpatterns (exclude overall match).
			}
			return OK;
		}
		else if (!_tcsicmp(name, _T("Name")))
		{
			if (pattern_found && mPatternName && mPatternName[p])
				TokenSetResult(aResultToken, mPatternName[p]);
			return OK;
		}
		else if (!_tcsicmp(name, _T("Mark")))
		{
			TokenSetResult(aResultToken, aParamCount == 1 && mMark ? mMark : _T(""));
			return OK;
		}
		else if (_tcsicmp(name, _T("Value"))) // i.e. NOT "Value".
		{
			// This is something like m[n] where n is not a valid subpattern or property name,
			// or m.Foo[n] where Foo is not a valid property name.  For the two-param case,
			// reset pattern_found so that if n is a valid subpattern it won't be returned:
			pattern_found = false;
		}
	}

	if (pattern_found)
	{
		// Gives the correct result even if there was no match (because length is 0):
		TokenSetResult(aResultToken, mHaystack - mHaystackStart + mOffset[p*2], mOffset[p*2+1]);
		return OK;
	}

	return INVOKE_NOT_HANDLED;
}

#ifdef CONFIG_DEBUGGER
void RegExMatchObject::DebugWriteProperty(IDebugProperties *aDebugger, int aPage, int aPageSize, int aMaxDepth)
{
	DebugCookie rootCookie, cookie;
	aDebugger->BeginProperty(NULL, "object", 5, rootCookie);
	if (aPage == 0)
	{
		aDebugger->WriteProperty("Count", ExprTokenType((__int64)mPatternCount));

		static LPSTR sNames[] = { "Value", "Pos", "Len", "Name" };
#ifdef UNICODE
		static LPWSTR sNamesT[] = { _T("Value"), _T("Pos"), _T("Len"), _T("Name") };
#else
		static LPSTR *sNamesT = sNames;
#endif
		TCHAR resultBuf[MAX_NUMBER_SIZE];
		ExprTokenType resultToken, thisTokenUnused, paramToken[2], *param[] = { &paramToken[0], &paramToken[1] };
		for (int i = 0; i < _countof(sNames); i++)
		{
			aDebugger->BeginProperty(sNames[i], "array", mPatternCount - (i == 3), cookie);
			paramToken[0].symbol = SYM_STRING;
			paramToken[0].marker = sNamesT[i];
			for (int p = (i == 3); p < mPatternCount; p++)
			{
				resultToken.symbol = SYM_STRING;
				resultToken.marker = _T("");
				resultToken.buf = resultBuf;
				resultToken.mem_to_free = NULL;
				paramToken[1].symbol = SYM_INTEGER;
				paramToken[1].value_int64 = p;
				Invoke(resultToken, thisTokenUnused, IT_GET, param, 2);
				aDebugger->WriteProperty(paramToken[1], resultToken);
				if (resultToken.mem_to_free)
					free(resultToken.mem_to_free);
			}
			aDebugger->EndProperty(cookie);
		}
	}
	aDebugger->EndProperty(rootCookie);
}
#endif


void *pcret_resolve_user_callout(LPCTSTR aCalloutParam, int aCalloutParamLength)
{
	// If no Func is found, pcre will handle the case where aCalloutParam is a pure integer.
	// In that case, the callout param becomes an integer between 0 and 255. No valid pointer
	// could be in this range, but we must take care to check (ptr>255) rather than (ptr!=NULL).
	Func *callout_func = g_script.FindFunc(aCalloutParam, aCalloutParamLength);
	if (!callout_func || callout_func->mIsBuiltIn)
		return NULL;
	return (void *)callout_func;
}

struct RegExCalloutData // L14: Used by BIF_RegEx to pass necessary info to RegExCallout.
{
	pcret *re;
	LPTSTR re_text; // original NeedleRegEx
	int options_length; // used to adjust cb->pattern_position
	int pattern_count; // to save calling pcre_fullinfo unnecessarily for each callout
	pcret_extra *extra;
	TCHAR output_mode;
};

int RegExCallout(pcret_callout_block *cb)
{
	// It should be documented that (?C) is ignored if encountered by the hook thread,
	// which could happen if SetTitleMatchMode,Regex and #IfWin are used. This would be a
	// problem if the callout should affect the outcome of the match or should be called
	// even if #IfWin will ultimately prevent the hotkey from firing. This is because:
	//	- The callout cannot be called from the hook thread, and therefore cannot affect
	//		the outcome of #IfWin when called by the hook thread.
	//	- If #IfWin does NOT prevent the hotkey from firing, it will be reevaluated from
	//		the main thread before the hotkey is actually fired. This will allow any
	//		callouts to occur on the main thread.
	//  - By contrast, if #IfWin DOES prevent the hotkey from firing, #IfWin will not be
	//		reevaluated from the main thread, so callouts cannot occur.
	if (GetCurrentThreadId() != g_MainThreadID)
		return 0;

	if (!cb->callout_data)
		return 0;

	Func *callout_func = (Func *)cb->user_callout;
	if (!callout_func)
	{
		Var *pcre_callout_var = g_script.FindVar(_T("pcre_callout"), 12); // This may be a local of the UDF which called RegExMatch/Replace().
		if (!pcre_callout_var)
			return 0; // Seems best to ignore the callout rather than aborting the match.

		callout_func = g_script.FindFunc(pcre_callout_var->Contents(), pcre_callout_var->Length());
		if (!callout_func || callout_func->mIsBuiltIn)
			return 0; // Could abort by returning PCRE_ERROR_CALLOUT, but ErrorLevel "-9" isn't very informative.
	}

	Func &func = *callout_func; // For simplicity and to keep the following section close to similar sections in OnMessage, RegisterCallbackCStub, etc.
	RegExCalloutData cd = *(RegExCalloutData *)cb->callout_data;

	// Adjust offset to account for options, which are excluded from the regex passed to PCRE.
	cb->pattern_position += cd.options_length;
	

	// See ExpandExpression() for detailed comments about the following section.
	VarBkp *var_backup = NULL;   // If needed, it will hold an array of VarBkp objects.
	int var_backup_count; // The number of items in the above array.
	if (func.mInstances > 0) // Backup is needed.
		if (!Var::BackupFunctionVars(func, var_backup, var_backup_count)) // Out of memory.
			return PCRE_ERROR_NOMEMORY;

	EventInfoType EventInfo_saved = g->EventInfo;

	g->EventInfo = (EventInfoType) cb;

	/*
	callout_number:		should be available since callout number can be specified within (?C...).
	subject:			useful when behaviour might depend on text surrounding a capture.
	start_match:		as above. equivalent to return value of RegExMatch, so should be available somehow.
	
	pattern_position:	useful to debug regexes when combined with auto-callouts. otherwise not useful.
	next_item_length:	as above. combined 'next_item' instead of these two would be less useful as it cannot distinguish between multiple identical items, and would sometimes be empty.
	
	capture_top:		not sure if useful? helps to distinguish between empty capture and non-capture. could maybe use callout number to determine this instead.
	capture_last:		as above.

	current_position:	can be derived from start_match and strlen(param1), or param1 itself if P option is used.
	offset_vector:		not very useful as same information available in local variables in more convenient form.
	callout_data:		not relevant, maybe use "user data" field of (RegExCalloutData*)callout_data if implemented.
	subject_length:		not useful, use strlen(subject).
	version:			not important.
	*/

	if (func.mParamCount > 0)
	{
		// UnquotedOutputVar
		Var &output_var = *func.mParam[0].var;

		Func *prev_func = g->CurrentFunc;
		// Set local vars of callout func where applicable.
		g->CurrentFunc = &func;

		int match_offset = cb->start_match;
		int match_length = cb->current_position - match_offset;

		// Overall match or its length.
		if (cd.output_mode == 'P')
			output_var.Assign(match_length);
		else if (cd.output_mode != 'O')
			output_var.AssignString(cb->subject + match_offset, match_length);

		LPTSTR mem_to_free = NULL;

		// Since matching is still in progress, these *should* be -1.
		// For maintainability and peace of mind, save them anyway:
		int offset[] = { cb->offset_vector[0], cb->offset_vector[1] };
		
		// Temporarily set these for use by the function below:
		cb->offset_vector[0] = cb->start_match;
		cb->offset_vector[1] = cb->current_position;
		if (cd.extra->flags & PCRE_EXTRA_MARK)
			*cd.extra->mark = UorA(wchar_t *, UCHAR *) cb->mark;
		
		// Set up local vars for capturing subpatterns.
		RegExSetSubpatternVars(cb->subject, cd.re, cd.extra, cd.output_mode, output_var, cb->offset_vector, cd.pattern_count, cb->capture_top, mem_to_free);

		// Restore to former offsets (probably -1):
		cb->offset_vector[0] = offset[0];
		cb->offset_vector[1] = offset[1];
		
		if (mem_to_free) // Should never happen since even if haystack were one of our local vars, BackupFunctionVars() would hide that from the above function. Check it anyway for maintainability.
			free(mem_to_free);

		// Restore g.CurrentFunc - func.Call() will also save, overwrite and restore it.
		g->CurrentFunc = prev_func;

		if (func.mParamCount > 1)
		{
			// Callout number
			func.mParam[1].var->Assign(cb->callout_number);

			if (func.mParamCount > 2)
			{
				// FoundPos
				func.mParam[2].var->Assign(match_offset + 1);

				if (func.mParamCount > 3)
				{
					// Haystack
					func.mParam[3].var->AssignString(cb->subject, cb->subject_length);
				
					if (func.mParamCount > 4)
					{
						// NeedleRegEx
						func.mParam[4].var->Assign(cd.re_text);
					}
				}
			}
		}
	}


	// Make all string positions one-based. UPDATE: offset_vector cannot be modified, so for consistency don't do this:
	//++cb->pattern_position;
	//++cb->start_match;
	//++cb->current_position;

	ExprTokenType result_token;
	ResultType result = func.Call(&result_token); // Call the UDF.

	// MUST handle return_value BEFORE calling FreeAndRestoreFunctionVars() because return_value
	// might be the contents of one of the function's local variables (which are about to be freed).
	int number_to_return = (int)TokenToInt64(result_token); // No need to check the following because they're implied for *return_value!=0: result != EARLY_EXIT && result != FAIL;
	if (result_token.symbol == SYM_OBJECT) // L31
		result_token.object->Release();

	Var::FreeAndRestoreFunctionVars(func, var_backup, var_backup_count);

	g->EventInfo = EventInfo_saved;

	if (result == FAIL)
	{
		// If a runtime error occurred, the user probably expects the thread to exit.  That can't
		// be easily done, so as a compromise, just abort matching.  If an exception was thrown,
		// the actual return value doesn't matter as long as matching is aborted, since execution
		// will be transferred to the CATCH block.
		number_to_return = PCRE_ERROR_NOMATCH;
	}

	// Behaviour of return values is defined by PCRE.
	return number_to_return;
}

pcret *get_compiled_regex(LPTSTR aRegEx, TCHAR &aOutputMode, pcret_extra *&aExtra
	, int *aOptionsLength, ExprTokenType *aResultToken)
// Returns the compiled RegEx, or NULL on failure.
// This function is called by things other than built-in functions so it should be kept general-purpose.
// Upon failure, if aResultToken!=NULL:
//   - ErrorLevel is set to a descriptive string other than "0".
//   - *aResultToken is set up to contain an empty string.
// Upon success, the following output parameters are set based on the options that were specified:
//    aGetPositionsNotSubstrings
//    aExtra
//    (but it doesn't change ErrorLevel on success, not even if aResultToken!=NULL)
// L14: aOptionsLength is used by callouts to adjust cb->pattern_position to be relative to beginning of actual user-specified NeedleRegEx instead of string seen by PCRE.
{	
	if (!pcret_callout)
	{	// Ensure this is initialized, even for ::RegExMatch() (to allow (?C) in window title regexes).
		pcret_callout = &RegExCallout;
	}

	// While reading from or writing to the cache, don't allow another thread entry.  This is because
	// that thread (or this one) might write to the cache while the other one is reading/writing, which
	// could cause loss of data integrity (the hook thread can enter here via #IfWin & SetTitleMatchMode RegEx).
	// Together, Enter/LeaveCriticalSection reduce performance by only 1.4% in the tightest possible script
	// loop that hits the first cache entry every time.  So that's the worst case except when there's an actual
	// collision, in which case performance suffers more because internally, EnterCriticalSection() does a
	// wait/semaphore operation, which is more costly.
	// Finally, the code size of all critical-section features together is less than 512 bytes (uncompressed),
	// so like performance, that's not a concern either.
	EnterCriticalSection(&g_CriticalRegExCache); // Request ownership of the critical section. If another thread already owns it, this thread will block until the other thread finishes.

	// SET UP THE CACHE.
	// This is a very crude cache for linear search. Of course, hashing would be better in the sense that it
	// would allow the cache to get much larger while still being fast (I believe PHP caches up to 4096 items).
	// Binary search might not be such a good idea in this case due to the time required to find the right spot
	// to insert a new cache item (however, items aren't inserted often, so it might perform quite well until
	// the cache contained thousands of RegEx's, which is unlikely to ever happen in most scripts).
	struct pcre_cache_entry
	{
		// For simplicity (and thus performance), the entire RegEx pattern including its options is cached
		// is stored in re_raw and that entire string becomes the RegEx's unique identifier for the purpose
		// of finding an entry in the cache.  Technically, this isn't optimal because some options like Study
		// and aGetPositionsNotSubstrings don't alter the nature of the compiled RegEx.  However, the CPU time
		// required to strip off some options prior to doing a cache search seems likely to offset much of the
		// cache's benefit.  So for this reason, as well as rarity and code size issues, this policy seems best.
		LPTSTR re_raw;      // The RegEx's literal string pattern such as "abc.*123".
		pcret *re_compiled; // The RegEx in compiled form.
		pcret_extra *extra; // NULL unless a study() was done (and NULL even then if study() didn't find anything).
		// int pcre_options; // Not currently needed in the cache since options are implicitly inside re_compiled.
		int options_length; // Lexikos: See aOptionsLength comment at beginning of this function.
		TCHAR output_mode;
	};

	#define PCRE_CACHE_SIZE 100 // Going too high would be counterproductive due to the slowness of linear search (and also the memory utilization of so many compiled RegEx's).
	static pcre_cache_entry sCache[PCRE_CACHE_SIZE] = {{0}};
	static int sLastInsert, sLastFound = -1; // -1 indicates "cache empty".
	int insert_pos; // v1.0.45.03: This is used to avoid updating sLastInsert until an insert actually occurs (it might not occur if a compile error occurs in the regex, or something else stops it early).

	// CHECK IF THIS REGEX IS ALREADY IN THE CACHE.
	if (sLastFound == -1) // Cache is empty, so insert this RegEx at the first position.
		insert_pos = 0;  // A section further below will change sLastFound to be 0.
	else
	{
		// Search the cache to see if it contains the caller-specified RegEx in compiled form.
		// First check if the last-found item is a match, since often it will be (such as cases
		// where a script-loop executes only one RegEx, and also for SetTitleMatchMode RegEx).
		if (!_tcscmp(aRegEx, sCache[sLastFound].re_raw)) // Match found (case sensitive).
			goto match_found; // And no need to update sLastFound because it's already set right.

		// Since above didn't find a match, search outward in both directions from the last-found match.
		// A bidirectional search is done because consecutively-called regex's tend to be adjacent to each other
		// in the array, so performance is improved on average (since most of the time when repeating a previously
		// executed regex, that regex will already be in the cache -- so optimizing the finding behavior is
		// more important than optimizing the never-found-because-not-cached behavior).
		bool go_right;
		int i, item_to_check, left, right;
		int last_populated_item = (sCache[PCRE_CACHE_SIZE-1].re_compiled) // When the array is full...
			? PCRE_CACHE_SIZE - 1  // ...all items must be checked except the one already done earlier.
			: sLastInsert;         // ...else only the items actually populated need to be checked.

		for (go_right = true, left = sLastFound, right = sLastFound, i = 0
			; i < last_populated_item  // This limits it to exactly the number of items remaining to be checked.
			; ++i, go_right = !go_right)
		{
			if (go_right) // Proceed rightward in the array.
			{
				right = (right == last_populated_item) ? 0 : right + 1; // Increment or wrap around back to the left side.
				item_to_check = right;
			}
			else // Proceed leftward.
			{
				left = (left == 0) ? last_populated_item : left - 1; // Decrement or wrap around back to the right side.
				item_to_check = left;
			}
			if (!_tcscmp(aRegEx, sCache[item_to_check].re_raw)) // Match found (case sensitive).
			{
				sLastFound = item_to_check;
				goto match_found;
			}
		}

		// Since above didn't goto, no match was found nor is one possible.  So just indicate the insert position
		// for where this RegEx will be put into the cache.
		// The following formula is for both cache-full and cache-partially-full.  When the cache is full,
		// it might not be the best possible formula; but it seems pretty good because it takes a round-robin
		// approach to overwriting/discarding old cache entries.  A discarded entry might have just been
		// used -- or even be sLastFound itself -- but on average, this approach seems pretty good because a
		// script loop that uses 50 unique RegEx's will quickly stabilize in the cache so that all 50 of them
		// stay compiled/cached until the loop ends.
		insert_pos = (sLastInsert == PCRE_CACHE_SIZE-1) ? 0 : sLastInsert + 1; // Formula works for both full and partially-full array.
	}
	// Since the above didn't goto:
	// - This RegEx isn't yet in the cache.  So compile it and put it in the cache, then return it to caller.
	// - Above is responsible for having set insert_pos to the cache position where the new RegEx will be stored.

	// The following macro is for maintainability, to enforce the definition of "default" in multiple places.
	// PCRE_NEWLINE_CRLF is the default in AutoHotkey rather than PCRE_NEWLINE_LF because *multiline* haystacks
	// that scripts will use are expected to come from:
	// 50%: FileRead: Uses `r`n by default, for performance)
	// 10%: Clipboard: Normally uses `r`n (includes files copied from Explorer, text data, etc.)
	// 20%: UrlDownloadToFile: Testing shows that it varies: e.g. microsoft.com uses `r`n, but `n is probably
	//      more common due to FTP programs automatically translating CRLF to LF when uploading to UNIX servers.
	// 20%: Other sources such as GUI edit controls: It's fairly unusual to want to use RegEx on multiline data
	//      from GUI controls, but in such case `n is much more common than `r`n.
#ifdef UNICODE
	#define AHK_PCRE_CHARSET_OPTIONS (PCRE_UTF8 | PCRE_NO_UTF8_CHECK)
#else
	#define AHK_PCRE_CHARSET_OPTIONS 0
#endif
	#define SET_DEFAULT_PCRE_OPTIONS \
	{\
		pcre_options = PCRE_NEWLINE_CRLF | AHK_PCRE_CHARSET_OPTIONS;\
		aOutputMode = '\0';\
		do_study = false;\
	}
	#define PCRE_NEWLINE_BITS (PCRE_NEWLINE_CRLF | PCRE_NEWLINE_ANY) // Covers all bits that are used for newline options.

	// SET DEFAULT OPTIONS:
	int pcre_options;
	long long do_study;
	SET_DEFAULT_PCRE_OPTIONS

	// PARSE THE OPTIONS (if any).
	LPCTSTR pat; // When options-parsing is done, pat will point to the start of the pattern itself.
	for (pat = aRegEx;; ++pat)
	{
		switch(*pat)
		{
		case 'i': pcre_options |= PCRE_CASELESS;  break;  // Perl-compatible options.
		case 'm': pcre_options |= PCRE_MULTILINE; break;  //
		case 's': pcre_options |= PCRE_DOTALL;    break;  //
		case 'x': pcre_options |= PCRE_EXTENDED;  break;  //
		case 'A': pcre_options |= PCRE_ANCHORED;  break;      // PCRE-specific options (uppercase used by convention, even internally by PCRE itself).
		case 'D': pcre_options |= PCRE_DOLLAR_ENDONLY; break; //
		case 'J': pcre_options |= PCRE_DUPNAMES;       break; //
		case 'U': pcre_options |= PCRE_UNGREEDY;       break; //
		case 'X': pcre_options |= PCRE_EXTRA;          break; //
		case 'C': pcre_options |= PCRE_AUTO_CALLOUT;   break; // L14: PCRE_AUTO_CALLOUT causes callouts to be created with callout_number == 255 before each item in the pattern.
		case '\a':pcre_options = (pcre_options & ~PCRE_NEWLINE_BITS) | PCRE_NEWLINE_ANY; break; // v1.0.46.06: alert/bell (i.e. `a) is used for PCRE_NEWLINE_ANY.
		case '\n':pcre_options = (pcre_options & ~PCRE_NEWLINE_BITS) | PCRE_NEWLINE_LF; break; // See below.
			// Above option: Could alternatively have called it "LF" rather than or in addition to "`n", but that
			// seems slightly less desirable due to potential overlap/conflict with future option letters,
			// plus the fact that `n should be pretty well known to AutoHotkey users, especially advanced ones
			// using RegEx.  Note: `n`r is NOT treated the same as `r`n because there's a slight chance PCRE
			// will someday support `n`r for some obscure usage (or just for symmetry/completeness).
			// The PCRE_NEWLINE_XXX options are valid for both compile() and exec(), but specifying it for exec()
			// would only serve to override the default stored inside the compiled pattern (seems rarely needed).
		case '\r':
			if (pat[1] == '\n') // Even though `r`n is the default, it's recognized as an option for flexibility and intuitiveness.
			{
				++pat; // Skip over the second character so that it's not recognized as a separate option by the next iteration.
				pcre_options = (pcre_options & ~PCRE_NEWLINE_BITS) | PCRE_NEWLINE_CRLF; // Set explicitly in case it was unset by an earlier option. Remember that PCRE_NEWLINE_CRLF is a bitwise combination of PCRE_NEWLINE_LF and CR.
			}
			else // For completeness, it's easy to support PCRE_NEWLINE_CR too, though nowadays I think it's quite rare (former Macintosh format).
				pcre_options = (pcre_options & ~PCRE_NEWLINE_BITS) | PCRE_NEWLINE_CR; // Do it this way because PCRE_NEWLINE_CRLF is a bitwise combination of PCRE_NEWLINE_CR and PCRE_NEWLINE_LF.
			break;

		// Other options (uppercase so that lowercase can be reserved for future/PERL options):
		case 'O':
		case 'P':
			aOutputMode = *pat;
			break;
		case 'S':
			do_study = true;
			break;

		case ' ':  // Allow only spaces and tabs as fillers so that everything else is protected/reserved for
		case '\t': // future use (such as future PERL options).
			break;

		case ')': // This character, when not escaped, marks the normal end of the options section.  We know it's not escaped because if it had been, the loop would have stopped at the backslash before getting here.
			++pat; // Set pat to be the start of the actual RegEx pattern, and leave options set to how they were by any prior iterations above.
			goto break_both;

		default: // Namely the following:
		//case '\0': No options are present, so ignore any letters that were accidentally recognized and treat entire string as the pattern.
		//case '(' : An open parenthesis must be considered an invalid option because otherwise it would be ambiguous with a subpattern.
		//case '\\': In addition to backslash being an invalid option, it also covers "\)" as being invalid (i.e. so that it's never necessary to check for an escaped close-parenthesis).
		//case all-other-chars: All others are invalid options; so like backslash above, ignore any letters that were accidentally recognized and treat entire string as the pattern.
			SET_DEFAULT_PCRE_OPTIONS // Revert to original options in case any early letters happened to match valid options.
			pat = aRegEx; // Indicate that the entire string is the pattern (no options).
			// To distinguish between a bad option and no options at all (for better error reporting), could check if
			// within the next few chars there's an unmatched close-parenthesis (non-escaped).  If so, the user
			// intended some options but one of them was invalid.  However, that would be somewhat difficult to do
			// because both \) and [)] are valid regex patterns that contain an unmatched close-parenthesis.
			// Since I don't know for sure if there are other cases (or whether future RegEx extensions might
			// introduce more such cases), it seems best not to attempt to distinguish.  Using more than two options
			// is extremely rare anyway, so syntax errors of this type do not happen often (and the only harm done
			// is a misleading error message from PCRE rather than something like "Bad option").  In addition,
			// omitting it simplifies the code and slightly improves performance.
			goto break_both;
		} // switch(*pat)
	} // for()

break_both:
	// Reaching here means that pat has been set to the beginning of the RegEx pattern itself and all options
	// are set properly.

	LPCSTR error_msg;
	TCHAR error_buf[ERRORLEVEL_SAVED_SIZE];
	int error_code, error_offset;
	pcret *re_compiled;

	// COMPILE THE REGEX.
	if (   !(re_compiled = pcret_compile2(pat, pcre_options, &error_code, &error_msg, &error_offset, NULL))   )
	{
		if (aResultToken) // Only when this is non-NULL does caller want ErrorLevel changed.
		{
			// Since both the error code and the offset are desirable outputs, it seems best to also
			// include descriptive error text (debatable).
			sntprintf(error_buf, _countof(error_buf), _T("Compile error %d at offset %d: %hs"), error_code
				, error_offset, error_msg);
			g_script.SetErrorLevelOrThrowStr(error_buf, aResultToken->marker);
		}
		goto error;
	}

	if (do_study)
	{
		// Enabling JIT compilation adds about 68 KB to the final executable size, which seems to outweigh
		// the speed-up that a minority of scripts would get.  Pass the option anyway, in case it is enabled:
		aExtra = pcret_study(re_compiled, PCRE_STUDY_JIT_COMPILE, &error_msg); // aExtra is an output parameter for caller.
		// Above returns NULL on failure or inability to find anything worthwhile in its study.  NULL is exactly
		// the right value to pass to exec() to indicate "no study info".
		// The following isn't done because:
		// 1) It seems best not to abort the caller's RegEx operation just due to a study error, since the only
		//    error likely to happen (from looking at PCRE's source code) is out-of-memory.
		// 2) ErrorLevel is traditionally used for error/abort conditions only, not warnings.  So it seems best
		//    not to pollute it with a warning message that indicates, "yes it worked, but here's a warning".
		//    ErrorLevel 0 (success) seems better and more desirable.
		// 3) Reduced code size.
		//if (error_msg)
		//{
			//if (aResultToken) // Only when this is non-NULL does caller want ErrorLevel changed.
			//{
			//	sntprintf(error_buf, sizeof(error_buf), "Study error: %s", error_msg);
			//	g_ErrorLevel->Assign(error_buf);
			//}
			//goto error;
		//}
	}
	else // No studying desired.
		aExtra = NULL; // aExtra is an output parameter for caller.

	// ADD THE NEWLY-COMPILED REGEX TO THE CACHE.
	// An earlier stage has set insert_pos to be the desired insert-position in the cache.
	pcre_cache_entry &this_entry = sCache[insert_pos]; // For performance and convenience.
	if (this_entry.re_compiled) // An existing cache item is being overwritten, so free it's attributes.
	{
		// Free the old cache entry's attributes in preparation for overwriting them with the new one's.
		free(this_entry.re_raw);           // Free the uncompiled pattern.
		pcret_free(this_entry.re_compiled); // Free the compiled pattern.
		if (this_entry.extra)
			pcret_free_study(this_entry.extra);
	}
	//else the insert-position is an empty slot, which is usually the case because most scripts contain fewer than
	// PCRE_CACHE_SIZE unique regex's.  Nothing extra needs to be done.
	this_entry.re_raw = _tcsdup(aRegEx); // _strdup() is very tiny and basically just calls _tcslen+malloc+_tcscpy.
	this_entry.re_compiled = re_compiled;
	this_entry.extra = aExtra;
	this_entry.output_mode = aOutputMode;
	// "this_entry.pcre_options" doesn't exist because it isn't currently needed in the cache.  This is
	// because the RE's options are implicitly stored inside re_compiled.

	// Lexikos: See aOptionsLength comment at beginning of this function.
	this_entry.options_length = (int)(pat - aRegEx);

	if (aOptionsLength) 
		*aOptionsLength = this_entry.options_length;

	sLastInsert = insert_pos; // v1.0.45.03: Must be done only *after* the insert succeeded because some things rely on sLastInsert being synonymous with the last populated item in the cache (when the cache isn't yet full).
	sLastFound = sLastInsert; // Relied upon in the case where sLastFound==-1. But it also sets things up to start the search at this item next time, because it's a bit more likely to be found here such as tight loops containing only one RegEx.
	// Remember that although sLastFound==sLastInsert in this case, it isn't always so -- namely when a previous
	// call found an existing match in the cache without having to compile and insert the item.

	LeaveCriticalSection(&g_CriticalRegExCache);
	return re_compiled; // Indicate success.

match_found: // RegEx was found in the cache at position sLastFound, so return the cached info back to the caller.
	aOutputMode = sCache[sLastFound].output_mode;
	aExtra = sCache[sLastFound].extra;
	if (aOptionsLength) // Lexikos: See aOptionsLength comment at beginning of this function.
		*aOptionsLength = sCache[sLastFound].options_length; 

	LeaveCriticalSection(&g_CriticalRegExCache);
	return sCache[sLastFound].re_compiled; // Indicate success.

error: // Since NULL is returned here, caller should ignore the contents of the output parameters.
	if (aResultToken)
	{
		aResultToken->symbol = SYM_STRING;
		aResultToken->marker = _T("");
	}

	LeaveCriticalSection(&g_CriticalRegExCache);
	return NULL; // Indicate failure.
}



LPTSTR RegExMatch(LPTSTR aHaystack, LPTSTR aNeedleRegEx)
// Returns NULL if no match.  Otherwise, returns the address where the pattern was found in aHaystack.
{
	TCHAR output_mode; // Currently ignored.
	pcret_extra *extra;
	pcret *re;

	// Compile the regex or get it from cache.
	if (   !(re = get_compiled_regex(aNeedleRegEx, output_mode, extra, NULL, NULL))   ) // Compiling problem.
		return NULL; // Our callers just want there to be "no match" in this case.

	// Set up the offset array, which consists of int-pairs containing the start/end offset of each match.
	// For simplicity, use a fixed size because even if it's too small (unlikely for our types of callers),
	// PCRE will still operate properly (though it returns 0 to indicate the too-small condition).
	#define RXM_INT_COUNT 30  // Should be a multiple of 3.
	int offset[RXM_INT_COUNT];

	// Execute the regex.
	int captured_pattern_count = pcret_exec(re, extra, aHaystack, (int)_tcslen(aHaystack), 0, 0, offset, RXM_INT_COUNT);
	if (captured_pattern_count < 0) // PCRE_ERROR_NOMATCH or some kind of error.
		return NULL;

	// Otherwise, captured_pattern_count>=0 (it's 0 when offset[] was too small; but that's harmless in this case).
	return aHaystack + offset[0]; // Return the position of the entire-pattern match.
}



void RegExReplace(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount
	, pcret *aRE, pcret_extra *aExtra, LPTSTR aHaystack, int aHaystackLength
	, int aStartingOffset, int aOffset[], int aNumberOfIntsInOffset)
{
	// Set default return value in case of early return.
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = aHaystack; // v1.0.46.06: aHaystack vs. "" is the new default because it seems a much safer and more convenient to return aHaystack when an unexpected PCRE-exec error occurs (such an error might otherwise cause loss of data in scripts that don't meticulously check ErrorLevel after each RegExReplace()).

	// If an output variable was provided for the count, resolve it early in case of early goto.
	// Fix for v1.0.47.05: In the unlikely event that output_var_count is the same script-variable as
	// as the haystack, needle, or replacement (i.e. the same memory), don't set output_var_count until
	// immediately prior to returning.  Otherwise, haystack, needle, or replacement would corrupted while
	// it's still being used here.
	Var *output_var_count = ParamIndexToOptionalVar(3); // SYM_VAR's Type() is always VAR_NORMAL (except lvalues in expressions).
	int replacement_count = 0; // This value will be stored in output_var_count, but only at the very end due to the reason above.

	// Get the replacement text (if any) from the incoming parameters.  If it was omitted, treat it as "".
	TCHAR repl_buf[MAX_NUMBER_SIZE];
	LPTSTR replacement = ParamIndexToOptionalString(2, repl_buf);

	// In PCRE, lengths and such are confined to ints, so there's little reason for using unsigned for anything.
	int captured_pattern_count, empty_string_is_not_a_match, match_length, ref_num
		, result_size, new_result_length, haystack_portion_length, second_iteration, substring_name_length
		, extra_offset, pcre_options;
	TCHAR *haystack_pos, *match_pos, *src, *src_orig, *closing_brace, *substring_name_pos;
	TCHAR *dest, char_after_dollar
		, substring_name[33] // In PCRE, "Names consist of up to 32 alphanumeric characters and underscores."
		, transform;

	// Caller has provided mem_to_free (initially NULL) as a means of passing back memory we allocate here.
	// So if we change "result" to be non-NULL, the caller will take over responsibility for freeing that memory.
	LPTSTR &result = aResultToken.mem_to_free; // Make an alias for convenience.
	size_t &result_length = aResultToken.marker_length; // MANDATORY FOR USERS OF CIRCUIT_TOKEN: set marker_length to the length of the string in marker.
	result_size = 0;   // And caller has already set "result" to be NULL.  The buffer is allocated only upon
	result_length = 0; // first use to avoid a potentially massive allocation that might be wasted and cause swapping (not to mention that we'll have better ability to estimate the correct total size after the first replacement is discovered).

	// Below uses a temp variable because realloc() returns NULL on failure but leaves original block allocated.
	// Note that if it's given a NULL pointer, realloc() does a malloc() instead.
	LPTSTR realloc_temp;
	#define REGEX_REALLOC(size) \
	{\
		result_size = size;\
		if (   !(realloc_temp = trealloc(result, result_size))   )\
			goto out_of_mem;\
		result = realloc_temp;\
	}

	// See if a replacement limit was specified.  If not, use the default (-1 means "replace all").
	int limit = ParamIndexToOptionalInt(4, -1);

	// aStartingOffset is altered further on in the loop; but for its initial value, the caller has ensured
	// that it lies within aHaystackLength.  Also, if there are no replacements yet, haystack_pos ignores
	// aStartingOffset because otherwise, when the first replacement occurs, any part of haystack that lies
	// to the left of a caller-specified aStartingOffset wouldn't get copied into the result.
	for (empty_string_is_not_a_match = 0, haystack_pos = aHaystack
		;; haystack_pos = aHaystack + aStartingOffset) // See comment above.
	{
		// Execute the expression to find the next match.
		captured_pattern_count = (limit == 0) ? PCRE_ERROR_NOMATCH // Only when limit is exactly 0 are we done replacing.  All negative values are "replace all".
			: pcret_exec(aRE, aExtra, aHaystack, aHaystackLength, aStartingOffset
				, empty_string_is_not_a_match, aOffset, aNumberOfIntsInOffset);

		if (captured_pattern_count == PCRE_ERROR_NOMATCH)
		{
			if (empty_string_is_not_a_match && aStartingOffset < aHaystackLength && limit != 0) // replacement_count>0 whenever empty_string_is_not_a_match!=0.
			{
				// This situation happens when a previous iteration found a match but it was the empty string.
				// That iteration told the pcre_exec that just occurred above to try to match something other than ""
				// at the same position.  But since we're here, it wasn't able to find such a match.  So just copy
				// the current character over literally then advance to the next character to resume normal searching.
				empty_string_is_not_a_match = 0; // Reset so that the next iteration starts off with the normal matching method.
#ifdef UNICODE
				// Need to avoid chopping a supplementary Unicode character in half.
				WCHAR c = haystack_pos[0];
				if (IS_SURROGATE_PAIR(c, haystack_pos[1])) // i.e. one supplementary character.
				{
					result[result_length++] = c;
					result[result_length++] = haystack_pos[1];
					aStartingOffset += 2; // Supplementary characters are in the range U+010000 to U+10FFFF,
					continue;
				}
#endif
				result[result_length++] = *haystack_pos; // This can't overflow because the size calculations in a previous iteration reserved 3 bytes: 1 for this character, 1 for the possible LF that follows CR, and 1 for the terminator.
				++aStartingOffset; // Advance to next candidate section of haystack.
				// v1.0.46.06: This following section was added to avoid finding a match between a CR and LF
				// when PCRE_NEWLINE_ANY mode is in effect.  The fact that this is the only change for
				// PCRE_NEWLINE_ANY relies on the belief that any pattern that matches the empty string in between
				// a CR and LF must also match the empty string that occurs right before the CRLF (even if that
				// pattern also matched a non-empty string right before the empty one in front of the CRLF).  If
				// this belief is correct, no logic similar to this is needed near the bottom of the main loop
				// because the empty string found immediately prior to this CRLF will put us into
				// empty_string_is_not_a_match mode, which will then execute this section of code (unless
				// empty_string_is_not_a_match mode actually found a match, in which case the logic here seems
				// superseded by that match?)  Even if this reasoning is not a complete solution, it might be
				// adequate if patterns that match empty strings are rare, which I believe they are.  In fact,
				// they might be so rare that arguably this could be documented as a known limitation rather than
				// having added the following section of code in the first place.
				// Examples that illustrate the effect:
				//    MsgBox % "<" . RegExReplace("`r`n", "`a).*", "xxx") . ">"
				//    MsgBox % "<" . RegExReplace("`r`n", "`am)^.*$", "xxx") . ">"
				if (*haystack_pos == '\r' && haystack_pos[1] == '\n')
				{
					// pcre_fullinfo() is a fast call, so it's called every time to simplify the code (I don't think
					// this whole "empty_string_is_not_a_match" section of code executes for most patterns anyway,
					// so performance seems less of a concern).
					if (!pcret_fullinfo(aRE, aExtra, PCRE_INFO_OPTIONS, &pcre_options) // Success.
						&& (pcre_options & PCRE_NEWLINE_ANY))
					{
						result[result_length++] = '\n'; // This can't overflow because the size calculations in a previous iteration reserved 3 bytes: 1 for this character, 1 for the possible LF that follows CR, and 1 for the terminator.
						++aStartingOffset; // Skip over this LF because it "belongs to" the CR that preceded it.
					}
				}
				continue; // i.e. we're not done yet because the "no match" above was a special one and there's still more haystack to check.
			}
			// Otherwise, there aren't any more matches.  So we're all done except for copying the last part of
			// haystack into the result (if applicable).
			if (replacement_count) // And by definition, result!=NULL due in this case to prior iterations.
			{
				if (haystack_portion_length = aHaystackLength - aStartingOffset) // This is the remaining part of haystack that needs to be copied over as-is.
				{
					new_result_length = (int)result_length + haystack_portion_length;
					if (new_result_length >= result_size)
						REGEX_REALLOC(new_result_length + 1); // This will end the loop if an alloc error occurs.
					tmemcpy(result + result_length, haystack_pos, haystack_portion_length); // memcpy() usually benches a little faster than _tcscpy().
					result_length = new_result_length; // Remember that result_length is actually an output for our caller, so even if for no other reason, it must be kept accurate for that.
				}
				result[result_length] = '\0'; // result!=NULL when replacement_count!=0.  Also, must terminate it unconditionally because other sections usually don't do it.
				// Set RegExMatch()'s return value to be "result":
				aResultToken.marker = result;  // Caller will take care of freeing result's memory.
			}
			// Section below is obsolete but is retained for its comments.
			//else // No replacements were actually done, so just return the original string to avoid malloc+memcpy
			// (in addition, returning the original might help the caller make other optimizations).
			//{
				// Already set as a default earlier, so commented out:
				//aResultToken.marker = aHaystack;
				// Not necessary to set output-var (length) for caller except when we allocated memory for the caller:
				//result_length = aHaystackLength; // result_length is an alias for an output parameter, so update for maintainability even though currently callers don't use it when no alloc of mem_to_free.
				//
				// There's no need to do the following because it should already be that way when replacement_count==0.
				//if (result)
				//	free(result);
				//result = NULL; // This tells the caller that we already freed it (i.e. from its POV, we never allocated anything).
			//}

			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // All done, indicate success via ErrorLevel.
			goto set_count_and_return;             //
		}

		// Otherwise:
		if (captured_pattern_count < 0) // An error other than "no match". These seem very rare, so it seems best to abort rather than yielding a partially-converted result.
		{
			g_script.SetErrorLevelOrThrowInt(captured_pattern_count, _T("RegExReplace")); // No error text is stored; just a negative integer (since these errors are pretty rare).
			goto set_count_and_return; // Goto vs. break to leave aResultToken.marker set to aHaystack and replacement_count set to 0, and let ErrorLevel tell the story.
		}

		// Otherwise (since above didn't return or break or continue), a match has been found (i.e.
		// captured_pattern_count > 0; it should never be 0 in this case because that only happens
		// when offset[] is too small, which it isn't).
		++replacement_count;
		--limit; // It's okay if it goes below -1 because all negatives are treated as "replace all".
		match_pos = aHaystack + aOffset[0]; // This is the location in aHaystack of the entire-pattern match.
		int match_end_offset = aOffset[1];
		haystack_portion_length = (int)(match_pos - haystack_pos); // The length of the haystack section between the end of the previous match and the start of the current one.

		// Handle this replacement by making two passes through the replacement-text: The first calculates the size
		// (which avoids having to constantly check for buffer overflow with potential realloc at multiple stages).
		// The second iteration copies the replacement (along with any literal text in haystack before it) into the
		// result buffer (which was expanded if necessary by the first iteration).
		for (second_iteration = 0; second_iteration < 2; ++second_iteration) // second_iteration is used as a boolean for readability.
		{
			if (second_iteration)
			{
				// Using the required length calculated by the first iteration, expand/realloc "result" if necessary.
				if (new_result_length + 3 > result_size) // Must use +3 not +1 in case of empty_string_is_not_a_match (which needs room for up to two extra characters).
				{
					// The first expression passed to REGEX_REALLOC is the average length of each replacement so far.
					// It's more typically more accurate to pass that than the following "length of current
					// replacement":
					//    new_result_length - haystack_portion_length - (aOffset[1] - aOffset[0])
					// Above is the length difference between the current replacement text and what it's
					// replacing (it's negative when replacement is smaller than what it replaces).
					REGEX_REALLOC((int)PredictReplacementSize((new_result_length - match_end_offset) / replacement_count // See above.
						, replacement_count, limit, aHaystackLength, new_result_length+2, match_end_offset)); // +2 in case of empty_string_is_not_a_match (which needs room for up to two extra characters).  The function will also do another +1 to convert length to size (for terminator).
					// The above will return if an alloc error occurs.
				}
				//else result_size is not only large enough, but also non-zero.  Other sections rely on it always
				// being non-zero when replacement_count>0.

				// Before doing the actual replacement and its backreferences, copy over the part of haystack that
				// appears before the match.
				if (haystack_portion_length)
				{
					tmemcpy(result + result_length, haystack_pos, haystack_portion_length);
					result_length += haystack_portion_length;
				}
				dest = result + result_length; // Init dest for use by the loops further below.
			}
			else // i.e. it's the first iteration, so begin calculating the size required.
				new_result_length = (int)result_length + haystack_portion_length; // Init length to the part of haystack before the match (it must be copied over as literal text).

			// DOLLAR SIGN ($) is the only method supported because it simplifies the code, improves performance,
			// and avoids the need to escape anything other than $ (which simplifies the syntax).
			for (src = replacement; ; ++src)  // For each '$' (increment to skip over the symbol just found by the inner for()).
			{
				// Find the next '$', if any.
				src_orig = src; // Init once for both loops below.
				if (second_iteration) // Mode: copy src-to-dest.
				{
					while (*src && *src != '$') // While looking for the next '$', copy over everything up until the '$'.
						*dest++ = *src++;
					result_length += (int)(src - src_orig);
				}
				else // This is the first iteration (mode: size-calculation).
				{
					for (; *src && *src != '$'; ++src); // Find the next '$', if any.
					new_result_length += (int)(src - src_orig); // '$' or '\0' was found: same expansion either way.
				}
				if (!*src)  // Reached the end of the replacement text.
					break;  // Nothing left to do, so if this is the first major iteration, begin the second.

				// Otherwise, a '$' has been found.  Check if it's a backreference and handle it.
				// But first process any special flags that are present.
				transform = '\0'; // Set default. Indicate "no transformation".
				extra_offset = 0; // Set default. Indicate that there's no need to hop over an extra character.
				if (char_after_dollar = src[1]) // This check avoids calling ctoupper on '\0', which directly or indirectly causes an assertion error in CRT.
				{
					switch(char_after_dollar = ctoupper(char_after_dollar))
					{
					case 'U':
					case 'L':
					case 'T':
						transform = char_after_dollar;
						extra_offset = 1;
						char_after_dollar = src[2]; // Ignore the transform character for the purposes of backreference recognition further below.
						break;
					//else leave things at their defaults.
					}
				}
				//else leave things at their defaults.

				ref_num = INT_MIN; // Set default to "no valid backreference".  Use INT_MIN to virtually guaranty that anything other than INT_MIN means that something like a backreference was found (even if it's invalid, such as ${-5}).
				switch (char_after_dollar)
				{
				case '{':  // Found a backreference: ${...
					substring_name_pos = src + 2 + extra_offset;
					if (closing_brace = _tcschr(substring_name_pos, '}'))
					{
						if (substring_name_length = (int)(closing_brace - substring_name_pos))
						{
							if (substring_name_length < _countof(substring_name))
							{
								tcslcpy(substring_name, substring_name_pos, substring_name_length + 1); // +1 to convert length to size, which truncates the new string at the desired position.
								if (IsPureNumeric(substring_name, true, false, true)) // Seems best to allow floating point such as 1.0 because it will then get truncated to an integer.  It seems to rare that anyone would want to use floats as names.
									ref_num = _ttoi(substring_name); // Uses _ttoi() vs. ATOI to avoid potential overlap with non-numeric names such as ${0x5}, which should probably be considered a name not a number?  In other words, seems best not to make some names that start with numbers "special" just because they happen to be hex numbers.
								else // For simplicity, no checking is done to ensure it consists of the "32 alphanumeric characters and underscores".  Let pcre_get_stringnumber() figure that out for us.
									ref_num = pcret_get_first_set(aRE, substring_name, aOffset); // Returns a negative on failure, which when stored in ref_num is relied upon as an indicator.
							}
							//else it's too long, so it seems best (debatable) to treat it as a unmatched/unfound name, i.e. "".
							src = closing_brace; // Set things up for the next iteration to resume at the char after "${..}"
						}
						//else it's ${}, so do nothing, which in effect will treat it all as literal text.
					}
					//else unclosed '{': for simplicity, do nothing, which in effect will treat it all as literal text.
					break;

				case '$':  // i.e. Two consecutive $ amounts to one literal $.
					++src; // Skip over the first '$', and the loop's increment will skip over the second. "extra_offset" is ignored due to rarity and silliness.  Just transcribe things like $U$ as U$ to indicate the problem.
					break; // This also sets up things properly to copy a single literal '$' into the result.

				case '\0': // i.e. a single $ was found at the end of the string.
					break; // Seems best to treat it as literal (strictly speaking the script should have escaped it).

				default:
					if (char_after_dollar >= '0' && char_after_dollar <= '9') // Treat it as a single-digit backreference. CONSEQUENTLY, $15 is really $1 followed by a literal '5'.
					{
						ref_num = char_after_dollar - '0'; // $0 is the whole pattern rather than a subpattern.
						src += 1 + extra_offset; // Set things up for the next iteration to resume at the char after $d. Consequently, $19 is seen as $1 followed by a literal 9.
					}
					//else not a digit: do nothing, which treats a $x as literal text (seems ok since like $19, $name will never be supported due to ambiguity; only ${name}).
				} // switch (char_after_dollar)

				if (ref_num == INT_MIN) // Nothing that looks like backreference is present (or the very unlikely ${-2147483648}).
				{
					if (second_iteration)
					{
						*dest++ = *src;  // src is incremented by the loop.  Copy only one character because the enclosing loop will take care of copying the rest.
						++result_length; // Update the actual length.
					}
					else
						++new_result_length; // Update the calculated length.
					// And now the enclosing loop will take care of the characters beyond src.
				}
				else // Something that looks like a backreference was found, even if it's invalid (e.g. ${-5}).
				{
					// It seems to improve convenience and flexibility to transcribe a nonexistent backreference
					// as a "" rather than literally (e.g. putting a ${1} literally into the new string).  Although
					// putting it in literally has the advantage of helping debugging, it doesn't seem to outweigh
					// the convenience of being able to specify nonexistent subpatterns. MORE IMPORANTLY a subpattern
					// might not exist per se if it hasn't been matched, such as an "or" like (abc)|(xyz), at least
					// when it's the last subpattern, in which case it should definitely be treated as "" and not
					// copied over literally.  So that would have to be checked for if this is changed.
					if (ref_num >= 0 && ref_num < captured_pattern_count) // Treat ref_num==0 as reference to the entire-pattern's match.
					{
						int ref_num0 = aOffset[ref_num*2];
						int ref_num1 = aOffset[ref_num*2 + 1];
						match_length = ref_num1 - ref_num0;
						if (match_length)
						{
							if (second_iteration)
							{
								tmemcpy(dest, aHaystack + ref_num0, match_length);
								if (transform)
								{
									dest[match_length] = '\0'; // Terminate for use below (shouldn't cause overflow because REALLOC reserved space for terminator; nor should there be any need to undo the termination afterward).
									switch(transform)
									{
									case 'U': CharUpper(dest); break;
									case 'L': CharLower(dest); break;
									case 'T': StrToTitleCase(dest); break;
									}
								}
								dest += match_length;
								result_length += match_length;
							}
							else // First iteration.
								new_result_length += match_length;
						}
					}
					//else subpattern doesn't exist (or it's invalid such as ${-5}, so treat it as blank because:
					// 1) It's boosts script flexibility and convenience (at the cost of making it hard to detect
					//    script bugs, which would be assisted by transcribing ${999} as literal text rather than "").
					// 2) It simplifies the code.
					// 3) A subpattern might not exist per se if it hasn't been matched, such as "(abc)|(xyz)"
					//    (in which case only one of them is matched).  If such a thing occurs at the end
					//    of the RegEx pattern, captured_pattern_count might not include it.  But it seems
					//    pretty clear that it should be treated as "" rather than some kind of error condition.
				}
			} // for() (for each '$')
		} // for() (a 2-iteration for-loop)

		// If we're here, a match was found.
		// Technique and comments from pcredemo.c:
		// If the previous match was NOT an empty string, we can just start the next match at the end
		// of the previous one.
		// If the previous match WAS an empty string, we can't do that, as it would lead to an
		// infinite loop. Instead, a special call of pcre_exec() is made with the PCRE_NOTEMPTY and
		// PCRE_ANCHORED flags set. The first of these tells PCRE that an empty string is not a valid match;
		// other possibilities must be tried. The second flag restricts PCRE to one match attempt at the
		// initial string position. If this match succeeds, that means there are two valid matches at the
		// SAME position: one for the empty string, and other for a non-empty string after it.  BOTH of
		// these matches are considered valid, and BOTH are eligible for replacement by RegExReplace().
		//
		// The following may be one example of this concept:
		// In the string "xy", replace the pattern "x?" by "z".  The traditional/proper answer (achieved by
		// the logic here) is "zzyz" because: 1) The first x is replaced by z; 2) The empty string before y
		// is replaced by z; 3) the logic here applies PCRE_NOTEMPTY to search again at the same position, but
		// that search doesn't find a match; so the logic higher above advances to the next character (y) and
		// continues the search; it finds the empty string at the end of haystack, which is replaced by z.
		// On the other hand, maybe there's a better example than the above that explains what would happen
		// if PCRE_NOTEMPTY actually finds a match, or what would happen if this PCRE_NOTEMPTY method weren't
		// used at all (i.e. infinite loop as mentioned in the previous paragraph).
		// 
		// If this match is "" (length 0), then by definition we just found a match in normal mode, not
		// PCRE_NOTEMPTY mode (since that mode isn't capable of finding "").  Thus, empty_string_is_not_a_match
		// is currently 0.  If "" was just found (and replaced), now try to find a second match at the same
		// position, but one that isn't "". This is done by switching to an alternate mode and doing another
		// iteration. Otherwise (the match found above wasn't "") advance to next candidate section of haystack
		// and resume searching.
		// v1.0.48.04: Fixed line below to reset to 0 (if appropriate) so that it doesn't get stuck in
		// PCRE_NOTEMPTY-mode for one extra iteration.  Otherwise there are too few replacements (4 vs. 5)
		// in examples like:
		//    RegExReplace("ABC", "Z*|A", "x")
		empty_string_is_not_a_match = (aOffset[0] == aOffset[1]) ? PCRE_NOTEMPTY|PCRE_ANCHORED : 0;
		aStartingOffset = match_end_offset; // In either case, set starting offset to the candidate for the next search.
	} // for()

	// All paths above should return (or goto some other label), so execution should never reach here except
	// through goto:
out_of_mem:
	// Due to extreme rarity and since this is a regex execution error of sorts, use PCRE's own error code.
	g_script.SetErrorLevelOrThrowInt(PCRE_ERROR_NOMEMORY, _T("RegExReplace"));
	if (result)
	{
		free(result);  // Since result is probably an non-terminated string (not to mention an incompletely created result), it seems best to free it here to remove it from any further consideration by the caller.
		result = NULL; // Tell caller that it was freed.
		// AND LEAVE aResultToken.marker (i.e. the final result) set to aHaystack, because the altered result is
		// indeterminate and thus discarded.
	}
	// Now fall through to below so that count is set even for out-of-memory error.
set_count_and_return:
	if (output_var_count)
		output_var_count->Assign(replacement_count); // v1.0.47.05: Must be done last in case output_var_count shares the same memory with haystack, needle, or replacement.
}



BIF_DECL(BIF_RegEx)
// This function is the initial entry point for both RegExMatch() and RegExReplace().
// Caller has set aResultToken.symbol to a default of SYM_INTEGER.
{
	bool mode_is_replace = ctoupper(aResultToken.marker[5]) == 'R'; // Union's marker initially contains the function name; e.g. RegEx[R]eplace.
	LPTSTR needle = ParamIndexToString(1, aResultToken.buf); // Load-time validation has already ensured that at least two actual parameters are present.

	TCHAR output_mode;
	pcret_extra *extra;
	pcret *re;
	int options_length;

	// COMPILE THE REGEX OR GET IT FROM CACHE.
	if (   !(re = get_compiled_regex(needle, output_mode, extra, &options_length, &aResultToken))   ) // Compiling problem.
		return; // It already set ErrorLevel and aResultToken for us. If caller provided an output var/array, it is not changed under these conditions because there's no way of knowing how many subpatterns are in the RegEx, and thus no way of knowing how far to init the array.

	// Since compiling succeeded, get info about other parameters.
	TCHAR haystack_buf[MAX_NUMBER_SIZE];
	LPTSTR haystack = ParamIndexToString(0, haystack_buf); // Load-time validation has already ensured that at least two actual parameters are present.
	int haystack_length = (int)ParamIndexLength(0, haystack);

	int param_index = mode_is_replace ? 5 : 3;
	int starting_offset;
	if (ParamIndexIsOmitted(param_index))
		starting_offset = 0; // The one-based starting position in haystack (if any).  Convert it to zero-based.
	else
	{
		starting_offset = ParamIndexToInt(param_index) - 1;
		if (starting_offset < 0) // Same convention as SubStr(): Treat a StartingPos of 0 (offset -1) as "start at the string's last char".  Similarly, treat negatives as starting further to the left of the end of the string.
		{
			starting_offset += haystack_length;
			if (starting_offset < 0)
				starting_offset = 0;
		}
		else if (starting_offset > haystack_length)
			// Although pcre_exec() seems to work properly even without this check, its absence would allow
			// the empty string to be found beyond the length of haystack, which could lead to problems and is
			// probably more trouble than its worth (assuming it has any worth -- perhaps for a pattern that
			// looks backward from itself; but that seems too rare to support and might create code that's
			// harder to maintain, especially in RegExReplace()).
			starting_offset = haystack_length; // Due to rarity of this condition, opt for simplicity: just point it to the terminator, which is in essence an empty string (which will cause result in "no match" except when searching for "").
	}

	// SET UP THE OFFSET ARRAY, which consists of int-pairs containing the start/end offset of each match.
	int pattern_count;
	pcret_fullinfo(re, extra, PCRE_INFO_CAPTURECOUNT, &pattern_count); // The number of capturing subpatterns (i.e. all except (?:xxx) I think). Failure is not checked because it seems too unlikely in this case.
	++pattern_count; // Increment to include room for the entire-pattern match.
	int number_of_ints_in_offset = pattern_count * 3; // PCRE uses 3 ints for each (sub)pattern: 2 for offsets and 1 for its internal use.
	int *offset = (int *)_alloca(number_of_ints_in_offset * sizeof(int)); // _alloca() boosts performance and seems safe because subpattern_count would usually have to be ridiculously high to cause a stack overflow.

	// The following section supports callouts (?C) and (*MARK:NAME).
	LPTSTR mark;
	RegExCalloutData callout_data;
	callout_data.re = re;
	callout_data.re_text = needle;
	callout_data.options_length = options_length;
	callout_data.pattern_count = pattern_count;
	callout_data.output_mode = output_mode;
	if (extra)
	{	// S (study) option was specified, use existing pcre_extra struct.
		extra->flags |= PCRE_EXTRA_CALLOUT_DATA | PCRE_EXTRA_MARK;	
	}
	else
	{	// Allocate a pcre_extra struct to pass callout_data.
		extra = (pcret_extra *)_alloca(sizeof(pcret_extra));
		extra->flags = PCRE_EXTRA_CALLOUT_DATA | PCRE_EXTRA_MARK;
	}
	// extra->callout_data is used to pass callout_data to PCRE.
	extra->callout_data = &callout_data;
	// callout_data.extra is used by RegExCallout, which only receives a pointer to callout_data.
	callout_data.extra = extra;
	// extra->mark is used by PCRE to return the NAME of a (*MARK:NAME), if encountered.
	extra->mark = UorA(wchar_t **, UCHAR **) &mark;

	if (mode_is_replace) // Handle RegExReplace() completely then return.
	{
		RegExReplace(aResultToken, aParam, aParamCount, re, extra, haystack, haystack_length
			, starting_offset, offset, number_of_ints_in_offset);
		return;
	}
	// OTHERWISE, THIS IS RegExMatch() not RegExReplace().

	// EXECUTE THE REGEX.
	int captured_pattern_count = pcret_exec(re, extra, haystack, haystack_length
		, starting_offset, 0, offset, number_of_ints_in_offset);

	int match_offset = 0; // Set default for no match/error cases below.

	// SET THE RETURN VALUE AND ERRORLEVEL BASED ON THE RESULTS OF EXECUTING THE EXPRESSION.
	if (captured_pattern_count == PCRE_ERROR_NOMATCH)
	{
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // i.e. "no match" isn't an error.
		aResultToken.value_int64 = 0;
		// BUT CONTINUE ON so that the output-array (if any) is fully reset (made blank), which improves
		// convenience for the script.
	}
	else if (captured_pattern_count < 0) // An error other than "no match".
	{
		g_script.SetErrorLevelOrThrowInt(captured_pattern_count, _T("RegExMatch")); // No error text is stored; just a negative integer (since these errors are pretty rare).
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
	}
	else // Match found, and captured_pattern_count >= 0 (but should never be 0 in this case because that only happens when offset[] is too small, which it isn't).
	{
		g_ErrorLevel->Assign(ERRORLEVEL_NONE);
		match_offset = offset[0];
		aResultToken.value_int64 = match_offset + 1; // i.e. the position of the entire-pattern match is the function's return value.
	}

	if (aParamCount < 3 || aParam[2]->symbol != SYM_VAR) // No output var, so nothing more to do.
		return;

	// OTHERWISE, THE CALLER PROVIDED AN OUTPUT VAR/ARRAY: Store the substrings that matched the patterns.
	Var &output_var = *aParam[2]->var; // SYM_VAR's Type() is always VAR_NORMAL (except lvalues in expressions).
	LPTSTR mem_to_free = NULL; // Set default.

	if (output_mode != 'O')
	{
		int match_length = (captured_pattern_count < 0) ? 0 // Seems better to store length of zero rather than something non-length like -1 (after all, the return value is blank in this case, which should be used as the error indicator).
			: offset[1] - offset[0];

		if (output_mode == 'P') // In this mode, it's done this way to avoid creating an array if there are no subpatterns; i.e. the return value is the starting position and the array name will contain the length of what was found.
			output_var.Assign(match_length);
		else
		{
			if (captured_pattern_count < 0) // Failed or no match.
				output_var.Assign(); // Make the full-pattern substring blank as a further indicator, and for convenience consistency in the script.
			else // Greater than 0 (it can't be equal to zero because offset[] was definitely large enough).
			{
				// Fix for v1.0.45.07: The following check allow haystack to be the same script-variable as the
				// output-var/array.  Unless a copy of haystack is made, any subpatterns to be populated after the
				// entire-pattern output-var below would be corrupted.  In other words, anything that refers to the
				// contents of haystack after the output-var has been assigned would otherwise refer to the wrong
				// string.  Note that the following isn't done for the 'P' or 'O' output modes higher above
				// because that mode never refers to haystack when populating its subpatterns.
				if (pattern_count > 1 && haystack == output_var.Contents(FALSE)) // i.e. there are subpatterns to be output afterward, and haystack is the same variable as the output-var that's about to be overwritten below.
					if (mem_to_free = _tcsdup(haystack)) // _strdup() is very tiny and basically just calls strlen+malloc+strcpy.
						haystack = mem_to_free;
					//else due to the extreme rarity of running out of memory AND SIMULTANEOUSLY having output-var match
					// haystack, continue on so that at least partial success is achieved (the only thing that will
					// be wrong in this case is the subpatterns, if any).
				output_var.Assign(haystack + match_offset, match_length); // It shouldn't be possible for the full-pattern match's offset to be -1, since if where here, a match on the full pattern was always found.
			}
		}
	}
	// Otherwise, output_mode == 'O', so the function call below does all the work.
	
	// L14: Moved this section into a function to allow it to be used by callouts.
	if (pattern_count > 1 || output_mode == 'O')
		RegExSetSubpatternVars(haystack, re, extra, output_mode, output_var, offset, pattern_count, captured_pattern_count, mem_to_free);

	if (mem_to_free)
		free(mem_to_free);
}



BIF_DECL(BIF_Ord)
{
	// Result will always be an integer (this simplifies scripts that work with binary zeros since an
	// empty string yields zero).
	// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
	LPTSTR cp = ParamIndexToString(0, aResultToken.buf);
#ifndef UNICODE
	// Always return a single byte in ANSI mode.
#else
	if (toupper(*aResultToken.marker) == 'O' && IS_SURROGATE_PAIR(cp[0], cp[1])) // Ord().
		aResultToken.value_int64 = ((cp[0] - 0xd800) << 10) + (cp[1] - 0xdc00) + 0x10000;
	else
#endif
		aResultToken.value_int64 = (TBYTE)*cp;
}



BIF_DECL(BIF_Chr)
{
	int param1 = ParamIndexToInt(0); // Convert to INT vs. UINT so that negatives can be detected.
	LPTSTR cp = aResultToken.buf; // If necessary, it will be moved to a persistent memory location by our caller.
	if (param1 < 0 || param1 > UorA(0x10FFFF, UCHAR_MAX))
		*cp = '\0'; // Empty string indicates both Chr(0) and an out-of-bounds param1.
#ifdef UNICODE
	else if (param1 >= 0x10000)
	{
		param1 -= 0x10000;
		cp[0] = 0xd800 + ((param1 >> 10) & 0x3ff);
		cp[1] = 0xdc00 + ( param1        & 0x3ff);
		cp[2] = '\0';
	}
#endif
	else
	{
		cp[0] = param1;
		cp[1] = '\0';
	}
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = cp;
}



BIF_DECL(BIF_NumGet)
{
	size_t right_side_bound, target; // Don't make target a pointer-type because the integer offset might not be a multiple of 4 (i.e. the below increments "target" directly by "offset" and we don't want that to use pointer math).
	ExprTokenType &target_token = *aParam[0];
	if (target_token.symbol == SYM_VAR) // SYM_VAR's Type() is always VAR_NORMAL (except lvalues in expressions).
	{
		target = (size_t)target_token.var->Contents(); // Although Contents(TRUE) will force an update of mContents if necessary, it very unlikely to be necessary here because we're about to fetch a binary number from inside mContents, not a normal/text number.
		right_side_bound = target + target_token.var->ByteCapacity(); // This is first illegal address to the right of target.
	}
	else
		target = (size_t)TokenToInt64(target_token);

	if (aParamCount > 1) // Parameter "offset" is present, so increment the address by that amount.  For flexibility, this is done even when the target isn't a variable.
	{
		if (aParamCount > 2 || TokenIsPureNumeric(*aParam[1])) // Checking aParamCount first avoids some unnecessary work in common cases where all parameters were specified.
			target += (ptrdiff_t)TokenToInt64(*aParam[1]); // Cast to ptrdiff_t vs. size_t to support negative offsets.
		else
			// Final param was omitted but this param is non-numeric, so use it as Type instead of Offset:
			++aParamCount, --aParam; // aParam[0] is no longer valid, but that's OK.
	}

	BOOL is_signed;
	size_t size = sizeof(DWORD_PTR); // Set default.

	if (aParamCount < 3) // The "type" parameter is absent (which is most often the case), so use defaults.
	{
#ifndef _WIN64 // is_signed is ignored on 64-bit builds due to lack of support for UInt64.  Explicitly disable this for maintainability.
		is_signed = FALSE;
#endif
		// And keep "size" at its default set earlier.
	}
	else // An explicit "type" is present.
	{
		LPTSTR type = TokenToString(*aParam[2], aResultToken.buf);
		if (ctoupper(*type) == 'U') // Unsigned.
		{
			++type; // Remove the first character from further consideration.
			is_signed = FALSE;
		}
		else
			is_signed = TRUE;

		switch(ctoupper(*type)) // Override "size" and aResultToken.symbol if type warrants it. Note that the above has omitted the leading "U", if present, leaving type as "Int" vs. "Uint", etc.
		{
		//case 'P': // Nothing extra needed in this case.
		case 'I':
			if (_tcschr(type, '6')) // Int64. It's checked this way for performance, and to avoid access violation if string is bogus and too short such as "i64".
				size = 8;
			else
				size = 4;
			break;
		case 'S': size = 2; break; // Short.
		case 'C': size = 1; break; // Char.

		case 'D': size = 8; aResultToken.symbol = SYM_FLOAT; break; // Double.
		case 'F': size = 4; aResultToken.symbol = SYM_FLOAT; break; // Float.

		// default: For any unrecognized values, keep "size" and aResultToken.symbol at their defaults set earlier
		// (for simplicity).
		}
	}

	// If the target is a variable, the following check ensures that the memory to be read lies within its capacity.
	// This seems superior to an exception handler because exception handlers only catch illegal addresses,
	// not ones that are technically legal but unintentionally bugs due to being beyond a variable's capacity.
	// Moreover, __try/except is larger in code size. Another possible alternative is IsBadReadPtr()/IsBadWritePtr(),
	// but those are discouraged by MSDN.
	// The following aren't covered by the check below:
	// - Due to rarity of negative offsets, only the right-side boundary is checked, not the left.
	// - Due to rarity and to simplify things, Float/Double (which "return" higher above) aren't checked.
	if (target < 65536 // Basic sanity check to catch incoming raw addresses that are zero or blank.
		|| target_token.symbol == SYM_VAR && target+size > right_side_bound) // i.e. it's ok if target+size==right_side_bound because the last byte to be read is actually at target+size-1. In other words, the position of the last possible terminator within the variable's capacity is considered an allowable address.
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
		return;
	}

	switch(size)
	{
	case 4: // Listed first for performance.
		if (aResultToken.symbol == SYM_FLOAT)
			aResultToken.value_double = *(float *)target;
		else if (is_signed)
			aResultToken.value_int64 = *(int *)target; // aResultToken.symbol was set to SYM_FLOAT or SYM_INTEGER higher above.
		else
			aResultToken.value_int64 = *(unsigned int *)target;
		break;
	case 8:
		// The below correctly copies both DOUBLE and INT64 into the union.
		// Unsigned 64-bit integers aren't supported because variables/expressions can't support them.
		aResultToken.value_int64 = *(__int64 *)target;
		break;
	case 2:
		if (is_signed) // Don't use ternary because that messes up type-casting.
			aResultToken.value_int64 = *(short *)target;
		else
			aResultToken.value_int64 = *(unsigned short *)target;
		break;
	default: // size 1
		if (is_signed) // Don't use ternary because that messes up type-casting.
			aResultToken.value_int64 = *(char *)target;
		else
			aResultToken.value_int64 = *(unsigned char *)target;
	}
}



BIF_DECL(BIF_Format)
{
	LPCTSTR fmt = ParamIndexToString(0), lit, cp, cp_end, cp_spec;
	LPTSTR target = NULL;
	int size = 0, spec_len;
	int param, last_param;
	TCHAR number_buf[MAX_NUMBER_SIZE];
	TCHAR spec[12+MAX_INTEGER_LENGTH*2];
	TCHAR custom_format;
	ExprTokenType value;
	*spec = '%';

	for (;;)
	{
		last_param = 0;

		for (lit = cp = fmt;; )
		{
			// Find next placeholder.
			for (cp_end = cp; *cp_end && *cp_end != '{'; ++cp_end);
			if (cp_end > lit)
			{
				// Handle literal text to the left of the placeholder.
				if (target)
					tmemcpy(target, lit, cp_end - lit), target += cp_end - lit;
				else
					size += int(cp_end - lit);
				lit = cp_end; // Mark this as the next literal character (to be overridden below if it's a valid placeholder).
			}
			cp = cp_end;
			if (!*cp)
				break;
			// else: Implies *cp == '{'.
			++cp;
			if ((*cp == '{' || *cp == '}') && cp[1] == '}') // {{} or {}}
			{
				if (target)
					*target++ = *cp;
				else
					++size;
				cp += 2;
				lit = cp; // Mark this as the next literal character.
				continue;
			}
			
			// Index.
			for (cp_end = cp; *cp_end >= '0' && *cp_end <= '9'; ++cp_end);
			if (cp_end > cp)
				param = ATOI(cp), cp = cp_end;
			else
				param = last_param + 1;
			if (param >= aParamCount) // Invalid parameter index.
				continue;

			custom_format = 0; // Set default.

			// Optional format specifier.
			if (*cp == ':')
			{
				cp_spec = ++cp;
				// Skip valid format specifier options.
				for (cp = cp_spec; *cp && _tcschr(_T("-+0 #"), *cp); ++cp); // flags
				for ( ; *cp >= '0' && *cp <= '9'; ++cp); // width
				if (*cp == '.') do ++cp; while (*cp >= '0' && *cp <= '9'); // .precision
				spec_len = int(cp - cp_spec);
				// For now, size specifiers (h | l | ll | w | I | I32 | I64) are not supported.
				
				if (spec_len + 4 >= _countof(spec)) // Format specifier too long (probably invalid).
					continue;
				// Copy options, if any (+1 to leave the leading %).
				tmemcpy(spec + 1, cp_spec, spec_len);
				++spec_len; // Include the leading %.

				if (_tcschr(_T("diouxX"), *cp))
				{
					spec[spec_len++] = 'I';
					spec[spec_len++] = '6';
					spec[spec_len++] = '4';
					// Integer value; apply I64 prefix to avoid truncation.
					value.value_int64 = ParamIndexToInt64(param);
					spec[spec_len++] = *cp++;
				}
				else if (_tcschr(_T("eEfgGaA"), *cp))
				{
					value.value_double = ParamIndexToDouble(param);
					spec[spec_len++] = *cp++;
				}
				else if (_tcschr(_T("cCp"), *cp))
				{
					// Input is an integer or pointer, but I64 prefix should not be applied.
					value.value_int64 = ParamIndexToInt64(param);
					spec[spec_len++] = *cp++;
				}
				else
				{
					spec[spec_len++] = 's'; // Default to string if not specified.
					if (_tcschr(_T("ULlTt"), *cp))
						custom_format = toupper(*cp++);
					if (*cp == 's')
						++cp;
				}
			}
			else
			{
				// spec[0] contains '%'.
				spec[1] = 's';
				spec_len = 2;
			}
			if (spec[spec_len - 1] == 's')
			{
				value.marker = ParamIndexToString(param, number_buf);
			}
			spec[spec_len] = '\0';
			
			if (*cp != '}') // Syntax error.
				continue;
			++cp;
			lit = cp; // Mark this as the next literal character.

			// Now that validation is complete, set last_param for use by the next {} or {:fmt}.
			last_param = param;
			
			if (target)
			{
				int len = _stprintf(target, spec, value.value_int64);
				switch (custom_format)
				{
				case 0: break; // Might help performance to list this first.
				case 'U': CharUpper(target); break;
				case 'L': CharLower(target); break;
				case 'T': StrToTitleCase(target); break;
				}
				target += len;
			}
			else
				size += _sctprintf(spec, value.value_int64);
		}
		if (target)
		{
			// Finished second pass.
			*target = '\0';
			return;
		}
		// Finished first pass (calculating required size).
		if (!TokenSetResult(aResultToken, NULL, size))
		{
			aResult = FAIL;
			return;
		}
		aResultToken.symbol = SYM_STRING;
		target = aResultToken.marker;
	}
}



BIF_DECL(BIF_NumPut)
{
	// Load-time validation has ensured that at least the first two parameters are present.
	ExprTokenType &token_to_write = *aParam[0];

	size_t right_side_bound, target; // Don't make target a pointer-type because the integer offset might not be a multiple of 4 (i.e. the below increments "target" directly by "offset" and we don't want that to use pointer math).
	ExprTokenType &target_token = *aParam[1];
	if (target_token.symbol == SYM_VAR) // SYM_VAR's Type() is always VAR_NORMAL (except lvalues in expressions).
	{
		target = (size_t)target_token.var->Contents(FALSE); // Pass FALSE for performance because contents is about to be overwritten, followed by a call to Close(). If something goes wrong and we return early, Contents() won't have been changed, so nothing about it needs updating.
		right_side_bound = target + target_token.var->ByteCapacity(); // This is the first illegal address to the right of target.
	}
	else
		target = (size_t)TokenToInt64(target_token);

	if (aParamCount > 2) // Parameter "offset" is present, so increment the address by that amount.  For flexibility, this is done even when the target isn't a variable.
	{
		if (aParamCount > 3 || TokenIsPureNumeric(*aParam[2])) // Checking aParamCount first avoids some unnecessary work in common cases where all parameters were specified.
			target += (ptrdiff_t)TokenToInt64(*aParam[2]); // Cast to ptrdiff_t vs. size_t to support negative offsets.
		else
			// Final param was omitted but this param is non-numeric, so use it as Type instead of Offset:
			++aParamCount, --aParam; // aParam[0] is no longer valid, but that's OK.
	}

	size_t size = sizeof(DWORD_PTR); // Set defaults.
	BOOL is_integer = TRUE;   //
	BOOL is_unsigned = (aParamCount > 3) ? FALSE : TRUE; // This one was added v1.0.48 to support unsigned __int64 the way DllCall does.

	if (aParamCount > 3) // The "type" parameter is present (which is somewhat unusual).
	{
		LPTSTR type = TokenToString(*aParam[3], aResultToken.buf);
		if (ctoupper(*type) == 'U') // Unsigned; but in the case of NumPut, it matters only for UInt64.
		{
			is_unsigned = TRUE;
			++type; // Remove the first character from further consideration.
		}

		switch(ctoupper(*type)) // Override "size" and is_integer if type warrants it. Note that the above has omitted the leading "U", if present, leaving type as "Int" vs. "Uint", etc.
		{
		case 'P': is_unsigned = TRUE; break; // Ptr.
		case 'I':
			if (_tcschr(type, '6')) // Int64. It's checked this way for performance, and to avoid access violation if string is bogus and too short such as "i64".
				size = 8;
			else
				size = 4;
			//else keep "size" at its default set earlier.
			break;
		case 'S': size = 2; break; // Short.
		case 'C': size = 1; break; // Char.

		case 'D': size = 8; is_integer = FALSE; break; // Double.
		case 'F': size = 4; is_integer = FALSE; break; // Float.

		// default: For any unrecognized values, keep "size" and is_integer at their defaults set earlier
		// (for simplicity).
		}
	}

	aResultToken.value_int64 = target + size; // This is used below and also as NumPut's return value. It's the address to the right of the item to be written.  aResultToken.symbol was set to SYM_INTEGER by our caller.

	// See comments in NumGet about the following section:
	if (target < 65536 // Basic sanity check to catch incoming raw addresses that are zero or blank.
		|| target_token.symbol == SYM_VAR && aResultToken.value_int64 > (INT_PTR)right_side_bound) // i.e. it's ok if target+size==right_side_bound because the last byte to be read is actually at target+size-1. In other words, the position of the last possible terminator within the variable's capacity is considered an allowable address.
	{
		if (target_token.symbol == SYM_VAR)
		{
			// Since target_token is a var, maybe the target is out of bounds because the var
			// hasn't been initialized (i.e. it has zero capacity).  Note that if a local var
			// has been given semi-permanent memory in a previous call to the function, the
			// check above might not catch it and we won't get an "uninitialized var" warning.
			// However, for that to happen the script must use VarSetCapacity that time but
			// not this time, which seems too rare to justify checking this every time.
			target_token.var->MaybeWarnUninitialized();
		}

		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
		return;
	}

	switch(size)
	{
	case 4: // Listed first for performance.
		if (is_integer)
			*(unsigned int *)target = (unsigned int)TokenToInt64(token_to_write);
		else // Float (32-bit).
			*(float *)target = (float)TokenToDouble(token_to_write);
		break;
	case 8:
		if (is_integer)
			// v1.0.48: Support unsigned 64-bit integers like DllCall does:
			*(__int64 *)target = (is_unsigned && !IS_NUMERIC(token_to_write.symbol)) // Must not be numeric because those are already signed values, so should be written out as signed so that whoever uses them can interpret negatives as large unsigned values.
				? (__int64)ATOU64(TokenToString(token_to_write)) // For comments, search for ATOU64 in BIF_DllCall().
				: TokenToInt64(token_to_write);
		else // Double (64-bit).
			*(double *)target = TokenToDouble(token_to_write);
		break;
	case 2:
		*(unsigned short *)target = (unsigned short)TokenToInt64(token_to_write);
		break;
	default: // size 1
		*(unsigned char *)target = (unsigned char)TokenToInt64(token_to_write);
	}
	if (target_token.symbol == SYM_VAR)
		target_token.var->Close(); // This updates various attributes of the variable.
	//else the target was an raw address.  If that address is inside some variable's contents, the above
	// attributes would already have been removed at the time the & operator was used on the variable.
}



BIF_DECL(BIF_StrGetPut)
{
	// To simplify flexible handling of parameters:
	ExprTokenType **aParam_end = aParam + aParamCount;

	LPCVOID source_string; // This may hold an intermediate UTF-16 string in ANSI builds.
	int source_length;
	if (ctoupper(aResultToken.marker[3]) == 'P')
	{
		// StrPut(String, Address[, Length][, Encoding])
		ExprTokenType &source_token = *aParam[0];
		source_string = (LPCVOID)TokenToString(source_token, aResultToken.buf); // Safe to use aResultToken.buf since StrPut won't use TokenSetResult.
		source_length = (int)((source_token.symbol == SYM_VAR) ? source_token.var->CharLength() : _tcslen((LPCTSTR)source_string));
		++aParam; // Remove the String param from further consideration.
	}
	else
	{
		// StrGet(Address[, Length][, Encoding])
		source_string = NULL;
		source_length = 0;
	}

	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T(""); // Set default in case of early return.

	LPVOID 	address;
	int 	length = -1; // actual length
	UINT 	encoding = UorA(CP_UTF16, CP_ACP); // native encoding

	// Parameters are interpreted according to the following rules (highest to lowest precedence):
	// Legend:  StrPut(String[, X, Y, Z])  or  StrGet(Address[, Y, Z])
	// - If X is non-numeric, it is Encoding.  Calculates required buffer size but does nothing else.  Y and Z must be omitted.
	// - If X is numeric, it is Address.  (For StrGet, non-numeric Address is treated as an error.)
	// - If Y is numeric, it is Length.  Otherwise "Actual length" is assumed.
	// - If a parameter remains, it is Encoding.
	// Encoding may therefore only be purely numeric if Address(X) and Length(Y) are specified.

	const LPVOID FIRST_VALID_ADDRESS = (LPVOID)65536;

	if (aParam < aParam_end && TokenIsPureNumeric(**aParam))
	{
		address = (LPVOID)TokenToInt64(**aParam);
		++aParam;
	}
	else
	{
		if (!source_string || aParamCount > 2)
			// This is StrGet or there are too many parameters; see below.
			return;
		// else this is the special measuring mode of StrPut, where Address and Length are omitted.
		// A length of 0 when passed to the Win API conversion functions (or the code below) means
		// "calculate the required buffer size, but don't do anything else."
		length = 0;
		address = FIRST_VALID_ADDRESS; // Skip validation below; address should never be used when length == 0.
	}

	if (aParam < aParam_end)
	{
		if (length == -1) // i.e. not StrPut(String, Encoding)
		{
			if (TokenIsPureNumeric(**aParam))
			{
				length = (int)TokenToInt64(**aParam);
				if (length < -1 || !length)
					return; // Invalid length; or caller of StrGet asked for 0 chars.
				++aParam; // Let encoding be the next param, if present.
			}
			else if ((*aParam)->symbol == SYM_MISSING)
			{
				// Length was "explicitly omitted", as in StrGet(Address,, Encoding),
				// which allows Encoding to be an integer without specifying Length.
				++aParam;
			}
			// aParam now points to aParam_end or the Encoding param.
		}
		if (aParam < aParam_end)
		{
			if (!TokenIsPureNumeric(**aParam))
			{
				encoding = Line::ConvertFileEncoding(TokenToString(**aParam));
				if (encoding == -1)
					return; // Invalid param.
			}
			else encoding = (UINT)TokenToInt64(**aParam);
		}
	}
	// Note: CP_AHKNOBOM is not supported; "-RAW" must be omitted.

	// Check for obvious errors to prevent an Access Violation.
	// Address can be zero for StrPut if length is also zero (see below).
	if ( address < FIRST_VALID_ADDRESS
		// Also check for overlap, in case memcpy is used instead of MultiByteToWideChar/WideCharToMultiByte.
		// (Behaviour for memcpy would be "undefined", whereas MBTWC/WCTBM would fail.)  Overlap in the
		// other direction (source_string beginning inside address..length) should not be possible.
		|| (address >= source_string && address <= ((LPTSTR)source_string + source_length)) )
		return;

	if (source_string) // StrPut
	{
		int char_count; // Either bytes or characters, depending on the target encoding.
		aResultToken.symbol = SYM_INTEGER; // All paths below return an integer.

		if (!source_length)
		{	// Take a shortcut when source_string is empty, since some paths below might not handle it correctly.
			if (length) {
				if (encoding == CP_UTF16)
					*(LPWSTR)address = '\0';
				else
					*(LPSTR)address = '\0';
			}
			aResultToken.value_int64 = 1;
			return;
		}

		if (encoding == UorA(CP_UTF16, CP_ACP))
		{
			// No conversion required: target encoding is the same as the native encoding of this build.
			char_count = source_length + 1; // + 1 because generally a null-terminator is wanted.
			if (length)
			{
				// Check for sufficient buffer space.  Cast to UINT and compare unsigned values: if length is
				// -1 it should be interpreted as a very large unsigned value, in effect bypassing this check.
				if ((UINT)source_length <= (UINT)length)
				{
					if (source_length == length)
						// Exceptional case: caller doesn't want a null-terminator (or passed this length in error).
						--char_count;
					// Copy the string, including null-terminator if requested.
					tmemcpy((LPTSTR)address, (LPCTSTR)source_string, char_count);
				}
				else
					// For consistency with the sections below, don't truncate the string.
					char_count = 0;
			}
			//else: Caller just wants the the required buffer size (char_count), which will be returned below.
			//	Note that although this seems equivalent to StrLen(), the caller might have explicitly
			//	passed an Encoding; in that case, the result of StrLen() might be different on the
			//	opposite build (ANSI vs Unicode) as the section below would be executed instead of this one.
		}
		else
		{
			// Conversion is required. For Unicode builds, this means encoding != CP_UTF16;
#ifndef UNICODE // therefore, this section is relevant only to ANSI builds:
			if (encoding == CP_UTF16)
			{
				// See similar section below for comments.
				if (length <= 0)
				{
					char_count = MultiByteToWideChar(CP_ACP, 0, (LPCSTR)source_string, source_length, NULL, 0) + 1;
					if (length == 0)
					{
						aResultToken.value_int64 = char_count;
						return;
					}
					length = char_count;
				}
				char_count = MultiByteToWideChar(CP_ACP, 0, (LPCSTR)source_string, source_length, (LPWSTR)address, length);
				if (char_count && char_count < length)
					((LPWSTR)address)[char_count++] = '\0';
			}
			else // encoding != CP_UTF16
			{
				// Convert native ANSI string to UTF-16 first.
				CStringWCharFromChar wide_buf((LPCSTR)source_string, source_length, CP_ACP);				
				source_string = wide_buf.GetString();
				source_length = wide_buf.GetLength();
#endif
				// UTF-8 does not support this flag.  Although the check further below would probably
				// compensate for this, UTF-8 is probably common enough to leave this exception here.
				DWORD flags = (encoding == CP_UTF8) ? 0 : WC_NO_BEST_FIT_CHARS;
				if (length <= 0) // -1 or 0
				{
					// Determine required buffer size.
					char_count = WideCharToMultiByte(encoding, flags, (LPCWSTR)source_string, source_length, NULL, 0, NULL, NULL);
					if (!char_count) // Above has ensured source is not empty, so this must be an error.
					{
						if (GetLastError() == ERROR_INVALID_FLAGS)
						{
							// Try again without flags.  MSDN lists a number of code pages for which flags must be 0, including UTF-7 and UTF-8 (but UTF-8 is handled above).
							flags = 0; // Must be set for this call and the call further below.
							char_count = WideCharToMultiByte(encoding, flags, (LPCWSTR)source_string, source_length, NULL, 0, NULL, NULL);
						}
						if (!char_count)
						{
							aResultToken.symbol = SYM_STRING;
							// aResultToken.marker is already set to "".
							return;
						}
					}
					++char_count; // + 1 for null-terminator (source_length causes it to be excluded from char_count).
					if (length == 0) // Caller just wants the required buffer size.
					{
						aResultToken.value_int64 = char_count;
						return;
					}
					// Assume there is sufficient buffer space and hope for the best:
					length = char_count;
				}
				// Convert to target encoding.
				char_count = WideCharToMultiByte(encoding, flags, (LPCWSTR)source_string, source_length, (LPSTR)address, length, NULL, NULL);
				// Since above did not null-terminate, check for buffer space and null-terminate if there's room.
				// It is tempting to always null-terminate (potentially replacing the last byte of data),
				// but that would exclude this function as a means to copy a string into a fixed-length array.
				if (char_count && char_count < length)
					((LPSTR)address)[char_count++] = '\0';
				// else no space to null-terminate; or conversion failed.
#ifndef UNICODE
			}
#endif
		}
		// Return the number of characters copied.
		aResultToken.value_int64 = char_count;
	}
	else // StrGet
	{
		if (encoding != UorA(CP_UTF16, CP_ACP))
		{
			// Conversion is required.
			int conv_length;
#ifdef UNICODE
			// Convert multi-byte encoded string to UTF-16.
			conv_length = MultiByteToWideChar(encoding, 0, (LPCSTR)address, length, NULL, 0);
			if (!TokenSetResult(aResultToken, NULL, conv_length)) // DO NOT SUBTRACT 1, conv_length might not include a null-terminator.
				return; // Out of memory.
			conv_length = MultiByteToWideChar(encoding, 0, (LPCSTR)address, length, aResultToken.marker, conv_length);
#else
			CStringW wide_buf;
			// If the target string is not UTF-16, convert it to that first.
			if (encoding != CP_UTF16)
			{
				StringCharToWChar((LPCSTR)address, wide_buf, length, encoding);
				address = (void *)wide_buf.GetString();
				length = wide_buf.GetLength();
			}

			// Now convert UTF-16 to ACP.
			conv_length = WideCharToMultiByte(CP_ACP, WC_NO_BEST_FIT_CHARS, (LPCWSTR)address, length, NULL, 0, NULL, NULL);
			if (!TokenSetResult(aResultToken, NULL, conv_length)) // DO NOT SUBTRACT 1, conv_length might not include a null-terminator.
				return; // Out of memory.
			conv_length = WideCharToMultiByte(CP_ACP, WC_NO_BEST_FIT_CHARS, (LPCWSTR)address, length, aResultToken.marker, conv_length, NULL, NULL);
#endif
			if (conv_length && !aResultToken.marker[conv_length - 1])
				--conv_length; // Exclude null-terminator.
			else
				aResultToken.marker[conv_length] = '\0';
			aResultToken.marker_length = conv_length; // Update this in case TokenSetResult used mem_to_free.
			return;
		}
		else if (length > -1)
		{
			// No conversion necessary, but we might not want the whole string.
			if (length == 0)
				return;	// Already set marker = "" above.
			// Copy and null-terminate at the specified length.
			TokenSetResult(aResultToken, (LPCTSTR)address, length);
			return;
		}

		// Return this null-terminated string, no conversion necessary.
		aResultToken.marker = (LPTSTR) address;
	}
}



BIF_DECL(BIF_IsLabel)
// For performance and code-size reasons, this function does not currently return what
// type of label it is (hotstring, hotkey, or generic).  To preserve the option to do
// this in the future, it has been documented that the function returns non-zero rather
// than "true".  However, if performance is an issue (since scripts that use IsLabel are
// often performance sensitive), it might be better to add a second parameter that tells
// IsLabel to look up the type of label, and return it as a number or letter.
{
	aResultToken.value_int64 = g_script.FindLabel(ParamIndexToString(0, aResultToken.buf)) ? 1 : 0; // "? 1 : 0" produces 15 bytes smaller OBJ size than "!= NULL" in this case (but apparently not in comparisons like x==y ? TRUE : FALSE).
}



BIF_DECL(BIF_IsFunc) // Lexikos: Added for use with dynamic function calls.
// Although it's tempting to return an integer like 0x8000000_min_max, where min/max are the function's
// minimum and maximum number of parameters stored in the low-order DWORD, it would be more friendly and
// readable to implement those outputs as optional ByRef parameters;
//     e.g. IsFunc(FunctionName, ByRef Minparameters, ByRef Maxparameters)
// It's also tempting to return something like 1+func.mInstances; but mInstances is tracked only due to
// the nature of the current implementation of function-recursion; it might not be something that would
// be tracked in future versions, and its value to the script is questionable.  Finally, a pointer to
// the Func struct itself could be returns so that the script could use NumGet() to retrieve function
// attributes.  However, that would expose implementation details that might be likely to change in the
// future, plus it would be cumbersome to use.  Therefore, something simple seems best; and since a
// dynamic function-call fails when too few parameters are passed (but not too many), it seems best to
// indicate to the caller not only that the function exists, but also how many parameters are required.
{
	Func *func = TokenToFunc(*aParam[0]);
	aResultToken.value_int64 = func ? (__int64)func->mMinParams+1 : 0;
}



BIF_DECL(BIF_Func)
// Returns a reference to an existing user-defined or built-in function, as an object.
{
	Func *func = g_script.FindFunc(ParamIndexToString(0, aResultToken.buf));
	if (func)
	{
		aResultToken.symbol = SYM_OBJECT;
		aResultToken.object = func;
	}
	else
		aResultToken.value_int64 = 0;
}


BIF_DECL(BIF_IsByRef)
{
	if (aParam[0]->symbol != SYM_VAR)
	{
		// Incorrect usage: return empty string to indicate the error.
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
	}
	else
	{
		// Return true if the var is an alias for another var.
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = (aParam[0]->var->ResolveAlias() != aParam[0]->var);
	}
}



BIF_DECL(BIF_GetKeyState)
{
	TCHAR key_name_buf[MAX_NUMBER_SIZE]; // Because aResultToken.buf is used for something else below.
	LPTSTR key_name = ParamIndexToString(0, key_name_buf);
	// Keep this in sync with GetKeyJoyState().
	// See GetKeyJoyState() for more comments about the following lines.
	JoyControls joy;
	int joystick_id;
	vk_type vk = TextToVK(key_name);
	if (!vk)
	{
		aResultToken.symbol = SYM_STRING; // ScriptGetJoyState() also requires that this be initialized.
		if (   !(joy = (JoyControls)ConvertJoy(key_name, &joystick_id))   )
			aResultToken.marker = _T("");
		else
		{
			// The following must be set for ScriptGetJoyState():
			aResultToken.marker = aResultToken.buf; // If necessary, it will be moved to a persistent memory location by our caller.
			ScriptGetJoyState(joy, joystick_id, aResultToken, true);
		}
		return;
	}
	// Since above didn't return: There is a virtual key (not a joystick control).
	TCHAR mode_buf[MAX_NUMBER_SIZE];
	LPTSTR mode = ParamIndexToOptionalString(1, mode_buf);
	KeyStateTypes key_state_type;
	switch (ctoupper(*mode)) // Second parameter.
	{
	case 'T': key_state_type = KEYSTATE_TOGGLE; break; // Whether a toggleable key such as CapsLock is currently turned on.
	case 'P': key_state_type = KEYSTATE_PHYSICAL; break; // Physical state of key.
	default: key_state_type = KEYSTATE_LOGICAL;
	}
	// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
	aResultToken.value_int64 = ScriptGetKeyState(vk, key_state_type); // 1 for down and 0 for up.
}



BIF_DECL(BIF_GetKeyName)
{
	// Get VK and/or SC from the first parameter, which may be a key name, scXXX or vkXX.
	// Key names are allowed even for GetKeyName() for simplicity and so that it can be
	// used to normalise a key name; e.g. GetKeyName("Esc") returns "Escape".
	LPTSTR key = ParamIndexToString(0, aResultToken.buf);
	vk_type vk;
	sc_type sc;
	TextToVKandSC(key, vk, sc);

	switch (ctoupper(aResultToken.marker[6]))
	{
	case 'V': // GetKey[V]K
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = vk ? vk : sc_to_vk(sc);
		break;
	case 'S': // GetKey[S]C
		aResultToken.symbol = SYM_INTEGER;
		aResultToken.value_int64 = sc ? sc : vk_to_sc(vk);
		break;
	default: // GetKey[N]ame
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = GetKeyName(vk, sc, aResultToken.buf, MAX_NUMBER_SIZE, _T(""));
	}
}



BIF_DECL(BIF_VarSetCapacity)
// Returns: The variable's new capacity.
// Parameters:
// 1: Target variable (unquoted).
// 2: Requested capacity.
// 3: Byte-value to fill the variable with (e.g. 0 to have the same effect as ZeroMemory).
{
	// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
	aResultToken.value_int64 = 0; // Set default. In spite of being ambiguous with the result of Free(), 0 seems a little better than -1 since it indicates "no capacity" and is also equal to "false" for easy use in expressions.
	if (aParam[0]->symbol == SYM_VAR) // SYM_VAR's Type() is always VAR_NORMAL (except lvalues in expressions).
	{
		Var &var = *aParam[0]->var; // For performance and convenience. SYM_VAR's Type() is always VAR_NORMAL (except lvalues in expressions).
		if (aParamCount > 1) // Second parameter is present.
		{
			// in bytes
			VarSizeType new_capacity = (VarSizeType)TokenToInt64(*aParam[1]);
			if (new_capacity == -1) // Adjust variable's internal length. Since new_capacity is unsigned, compare directly to -1 rather than doing <0.
			{
				// Seems more useful to report length vs. capacity in this special case. Scripts might be able
				// to use this to boost performance.
				aResultToken.value_int64 = var.ByteLength() = ((VarSizeType)_tcslen(var.Contents()) * sizeof(TCHAR)); // Performance: Length() and Contents() will update mContents if necessary, it's unlikely to be necessary under the circumstances of this call.  In any case, it seems appropriate to do it this way.
				var.Close(); // v1.0.44.14: Removes attributes like VAR_ATTRIB_BINARY_CLIP (if present) because it seems more flexible to convert binary-to-normal rather than checking IsBinaryClip() then doing nothing if it binary.
				return;
			}
			// Since above didn't return:
			if (!new_capacity && aParam[1]->symbol == SYM_MISSING)
			{
				// This case is likely to be rare, but allows VarSetCapacity(v,,b) to fill
				// v with byte b up to its current capacity, rather than freeing it.
				if (new_capacity = var.ByteCapacity())
					new_capacity -= sizeof(TCHAR);
			}
			if (new_capacity)
			{
				var.SetCapacity(new_capacity, true, false); // This also destroys the variables contents.
				// in characters
				VarSizeType capacity;
				if (aParamCount > 2 && (capacity = var.Capacity()) > 1) // Third parameter is present and var has enough capacity to make FillMemory() meaningful.
				{
					--capacity; // Convert to script-POV capacity. To avoid underflow, do this only now that Capacity() is known not to be zero.
					// The following uses capacity-1 because the last byte of a variable should always
					// be left as a binary zero to avoid crashes and problems due to unterminated strings.
					// In other words, a variable's usable capacity from the script's POV is always one
					// less than its actual capacity:
					BYTE fill_byte = (BYTE)TokenToInt64(*aParam[2]); // For simplicity, only numeric characters are supported, not something like "a" to mean the character 'a'.
					LPTSTR contents = var.Contents();
					FillMemory(contents, capacity * sizeof(TCHAR), fill_byte); // Last byte of variable is always left as a binary zero.
					contents[capacity] = '\0'; // Must terminate because nothing else is explicitly responsible for doing it.
					var.SetCharLength(fill_byte ? capacity : 0); // Length is same as capacity unless fill_byte is zero.
				}
				else
					// By design, Assign() has already set the length of the variable to reflect new_capacity.
					// This is not what is wanted in this case since it should be truly empty.
					var.ByteLength() = 0;
			}
			else // ALLOC_SIMPLE, due to its nature, will not actually be freed, which is documented.
				var.Free();
		} // if (aParamCount > 1)
		else
		{
			// RequestedCapacity was omitted, so the var is not altered; instead, the current capacity
			// is reported, which seems more intuitive/useful than having it do a Free(). In this case
			// it's an input var rather than an output var, so check if it has been initialized:
			// v1.1.11.01: Support VarSetCapacity(var) as a means for the script to check if it
			// has initialized a var.  In other words, don't show a warning even in that case.
			//var.MaybeWarnUninitialized();
		}

		if (aResultToken.value_int64 = var.ByteCapacity()) // Don't subtract 1 here in lieu doing it below (avoids underflow).
			aResultToken.value_int64 -= sizeof(TCHAR); // Omit the room for the zero terminator since script capacity is defined as length vs. size.
	} // (aParam[0]->symbol == SYM_VAR)
}



BIF_DECL(BIF_FileExist)
{
	TCHAR filename_buf[MAX_NUMBER_SIZE]; // Because aResultToken.buf is used for something else below.
	LPTSTR filename = ParamIndexToString(0, filename_buf);
	aResultToken.marker = aResultToken.buf; // If necessary, it will be moved to a persistent memory location by our caller.
	aResultToken.symbol = SYM_STRING;
	DWORD attr;
	if (DoesFilePatternExist(filename, &attr))
	{
		// Yield the attributes of the first matching file.  If not match, yield an empty string.
		// This relies upon the fact that a file's attributes are never legitimately zero, which
		// seems true but in case it ever isn't, this forces a non-empty string be used.
		// UPDATE for v1.0.44.03: Someone reported that an existing file (created by NTbackup.exe) can
		// apparently have undefined/invalid attributes (i.e. attributes that have no matching letter in
		// "RASHNDOCT").  Although this is unconfirmed, it's easy to handle that possibility here by
		// checking for a blank string.  This allows FileExist() to report boolean TRUE rather than FALSE
		// for such "mystery files":
		FileAttribToStr(aResultToken.marker, attr);
		if (!*aResultToken.marker) // See above.
		{
			// The attributes might be all 0, but more likely the file has some of the newer attributes
			// such as FILE_ATTRIBUTE_ENCRYPTED (or has undefined attributes).  So rather than storing attr as
			// a hex number (which could be zero and thus defeat FileExist's ability to detect the file), it
			// seems better to store some arbitrary letter (other than those in "RASHNDOCT") so that FileExist's
			// return value is seen as boolean "true".
			aResultToken.marker[0] = 'X';
			aResultToken.marker[1] = '\0';
		}
	}
	else // Empty string is the indicator of "not found" (seems more consistent than using an integer 0, since caller might rely on it being SYM_STRING).
		*aResultToken.marker = '\0';
}



BIF_DECL(BIF_WinExistActive)
{
	LPTSTR bif_name = aResultToken.marker;  // Save this early for maintainability (it is the name of the function, provided by the caller).
	aResultToken.symbol = SYM_STRING; // Returns a string to preserve hex format.

	TCHAR *param[4], param_buf[4][MAX_NUMBER_SIZE];
	for (int j = 0; j < 4; ++j) // For each formal parameter, including optional ones.
		param[j] = ParamIndexToOptionalString(j, param_buf[j]);

	// Should be called the same was as ACT_IFWINEXIST and ACT_IFWINACTIVE:
	HWND found_hwnd = (ctoupper(bif_name[3]) == 'E') // Win[E]xist.
		? WinExist(*g, param[0], param[1], param[2], param[3], false, true)
		: WinActive(*g, param[0], param[1], param[2], param[3], true);

	aResultToken.marker = HwndToString(found_hwnd, aResultToken.buf);
}



BIF_DECL(BIF_Round)
// For simplicity and backward compatibility, this always yields something numeric (or a string that's numeric).
// Even Round(empty_or_unintialized_var) is zero rather than "".
{
	LPTSTR buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).

	// See TRANS_CMD_ROUND for details.
	int param2;
	double multiplier;
	if (aParamCount > 1)
	{
		param2 = ParamIndexToInt(1);
		multiplier = qmathPow(10, param2);
	}
	else // Omitting the parameter is the same as explicitly specifying 0 for it.
	{
		param2 = 0;
		multiplier = 1;
	}
	double value = ParamIndexToDouble(0);
	aResultToken.value_double = (value >= 0.0 ? qmathFloor(value * multiplier + 0.5)
		: qmathCeil(value * multiplier - 0.5)) / multiplier;

	// If incoming value is an integer, it seems best for flexibility to convert it to a
	// floating point number whenever the second param is >0.  That way, it can be used
	// to "cast" integers into floats.  Conversely, it seems best to yield an integer
	// whenever the second param is <=0 or omitted.
	if (param2 > 0) // aResultToken.value_double already contains the result.
	{
		// v1.0.44.01: Since Round (in its param2>0 mode) is almost always used to facilitate some kind of
		// display or output of the number (hardly ever for intentional reducing the precision of a floating
		// point math operation), it seems best by default to omit only those trailing zeroes that are beyond
		// the specified number of decimal places.  This is done by converting the result into a string here,
		// which will cause the expression evaluation to write out the final result as this very string as long
		// as no further floating point math is done on it (such as Round(3.3333, 2)+0).  Also note that not
		// all trailing zeros are removed because it is often the intent that exactly the number of decimal
		// places specified should be *shown* (for column alignment, etc.).  For example, Round(3.5, 2) should
		// be 3.50 not 3.5.  Similarly, Round(1, 2) should be 1.00 not 1 (see above comment about "casting" for
		// why.
		// Performance: This method is about twice as slow as the old method (which did merely the line
		// "aResultToken.symbol = SYM_FLOAT" in place of the below).  However, that might be something
		// that can be further optimized in the caller (its calls to _tcslen, memcpy, etc. might be optimized
		// someday to omit certain calls when very simply situations allow it).  In addition, twice as slow is
		// not going to impact the vast majority of scripts since as mentioned above, Round (in its param2>0
		// mode) is almost always used for displaying data, not for intensive operations within a expressions.
		// AS DOCUMENTED: Round(..., positive_number) doesn't obey SetFormat (nor scientific notation).
		// The script can force Round(x, 2) to obey SetFormat by adding 0 to the result (if it wants).
		// Also, a new parameter an be added someday to trim excess trailing zeros from param2>0's result
		// (e.g. Round(3.50, 2, true) can be 3.5 rather than 3.50), but this seems less often desired due to
		// column alignment and other goals where consistency is important.
		_stprintf(buf, _T("%0.*f"), param2, aResultToken.value_double); // %f can handle doubles in MSVC++.
		aResultToken.marker = buf;
		aResultToken.symbol = SYM_STRING;
	}
	else
		// Fix for v1.0.47.04: See BIF_FloorCeil() for explanation of this fix.  Currently, the only known example
		// of when the fix is necessary is the following script in release mode (not debug mode):
		//   myNumber  := 1043.22  ; Bug also happens with -1043.22 (negative).
		//   myRounded1 := Round( myNumber, -1 )  ; Stores 1040 (correct).
		//   ChartModule := DllCall("LoadLibrary", "str", "rmchart.dll")
		//   myRounded2 := Round( myNumber, -1 )  ; Stores 1039 (wrong).
		aResultToken.value_int64 = (__int64)(aResultToken.value_double + (aResultToken.value_double > 0 ? 0.2 : -0.2));
		// Formerly above was: aResultToken.value_int64 = (__int64)aResultToken.value_double;
		// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
}



BIF_DECL(BIF_FloorCeil)
// Probably saves little code size to merge extremely short/fast functions, hence FloorCeil.
// Floor() rounds down to the nearest integer; that is, to the integer that lies to the left on the
// number line (this is not the same as truncation because Floor(-1.2) is -2, not -1).
// Ceil() rounds up to the nearest integer; that is, to the integer that lies to the right on the number line.
//
// For simplicity and backward compatibility, a numeric result is always returned (even if the input
// is non-numeric or an empty string).
{
	// The code here is similar to that in TRANS_CMD_FLOOR/CEIL, so maintain them together.
	// The qmath routines are used because Floor() and Ceil() are deceptively difficult to implement in a way
	// that gives the correct result in all permutations of the following:
	// 1) Negative vs. positive input.
	// 2) Whether or not the input is already an integer.
	// Therefore, do not change this without conduction a thorough test.
	double x = ParamIndexToDouble(0);
	x = (ctoupper(aResultToken.marker[0]) == 'F') ? qmathFloor(x) : qmathCeil(x);
	// Fix for v1.0.40.05: For some inputs, qmathCeil/Floor yield a number slightly to the left of the target
	// integer, while for others they yield one slightly to the right.  For example, Ceil(62/61) and Floor(-4/3)
	// yield a double that would give an incorrect answer if it were simply truncated to an integer via
	// type casting.  The below seems to fix this without breaking the answers for other inputs (which is
	// surprisingly harder than it seemed).  There is a similar fix in BIF_Round().
	aResultToken.value_int64 = (__int64)(x + (x > 0 ? 0.2 : -0.2));
	// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
}



BIF_DECL(BIF_Mod)
{
	// Load-time validation has already ensured there are exactly two parameters.
	// "Cast" each operand to Int64/Double depending on whether it has a decimal point.
	ExprTokenType param0, param1;
	if (!ParamIndexToNumber(0, param0) || !ParamIndexToNumber(1, param1)) // Non-operand or non-numeric string.
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
		return;
	}
	if (param0.symbol == SYM_INTEGER && param1.symbol == SYM_INTEGER)
	{
		if (!param1.value_int64) // Divide by zero.
		{
			aResultToken.symbol = SYM_STRING;
			aResultToken.marker = _T("");
		}
		else
			// For performance, % is used vs. qmath for integers.
			// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
			aResultToken.value_int64 = param0.value_int64 % param1.value_int64;
	}
	else // At least one is a floating point number.
	{
		double dividend = TokenToDouble(param0);
		double divisor = TokenToDouble(param1);
		if (divisor == 0.0) // Divide by zero.
		{
			aResultToken.symbol = SYM_STRING;
			aResultToken.marker = _T("");
		}
		else
		{
			aResultToken.symbol = SYM_FLOAT;
			aResultToken.value_double = qmathFmod(dividend, divisor);
		}
	}
}



BIF_DECL(BIF_MinMax)
{
	// Supports one or more parameters.
	// Load-time validation has already ensured there is at least one parameter.
	ExprTokenType param;
	int index, ib_index = 0, db_index = 0;
	bool isMin = (ctoupper(aResultToken.marker[1]) == 'I') ? TRUE : FALSE; // To save code size.
	__int64 ia, ib = 0; double da, db = 0;
	bool ib_empty = TRUE, db_empty = TRUE;
	for (int i = 0; i < aParamCount; ++i)
	{
		ParamIndexToNumber(i, param);
		switch (param.symbol)
		{
			case SYM_INTEGER: // Compare only integers.
				ia = param.value_int64;
				if ((ib_empty) || (isMin ? ia < ib : ia > ib))
				{
					ib_empty = FALSE;
					ib = ia;
					ib_index = i;
				}
				break;
			case SYM_FLOAT: // Compare only floats.
				da = param.value_double;
				if ((db_empty) || (isMin ? da < db : da > db))
				{
					db_empty = FALSE;
					db = da;
					db_index = i;
				}
				break;
			default: // Non-operand or non-numeric string.
				aResultToken.symbol = SYM_STRING;
				aResultToken.marker = _T("");
				return; // Return a blank value to indicate the problem.
		}
	}
	// Compare found integer with found float:
	index = (db_empty || !ib_empty && (isMin ? ib < db : ib > db)) ? ib_index : db_index;
	ParamIndexToNumber(index, param);
	aResultToken.symbol = param.symbol;
	aResultToken.value_int64 = param.value_int64;
}



BIF_DECL(BIF_Abs)
{
	// Unlike TRANS_CMD_ABS, which removes the minus sign from the string if it has one,
	// this is done in a more traditional way.  It's hard to imagine needing the minus
	// sign removal method here since a negative hex literal such as -0xFF seems too rare
	// to worry about.  One additional reason not to remove minus signs from strings is
	// that it might produce inconsistent results depending on whether the operand is
	// generic (SYM_OPERAND) and numeric.  In other words, abs() shouldn't treat a
	// sub-expression differently than a numeric literal.
	if (!TokenToDoubleOrInt64(*aParam[0], aResultToken)) // "Cast" token to Int64/Double depending on whether it has a decimal point.
		// Non-operand or non-numeric string. TokenToDoubleOrInt64() has already set the token to be an
		// empty string for us.
		return;
	if (aResultToken.symbol == SYM_INTEGER)
	{
		// The following method is used instead of __abs64() to allow linking against the multi-threaded
		// DLLs (vs. libs) if that option is ever used (such as for a minimum size AutoHotkeySC.bin file).
		// It might be somewhat faster than __abs64() anyway, unless __abs64() is a macro or inline or something.
		if (aResultToken.value_int64 < 0)
			aResultToken.value_int64 = -aResultToken.value_int64;
	}
	else // Must be SYM_FLOAT due to the conversion above.
		aResultToken.value_double = qmathFabs(aResultToken.value_double);
}



BIF_DECL(BIF_Sin)
// For simplicity and backward compatibility, a numeric result is always returned (even if the input
// is non-numeric or an empty string).
{
	aResultToken.symbol = SYM_FLOAT;
	aResultToken.value_double = qmathSin(ParamIndexToDouble(0));
}



BIF_DECL(BIF_Cos)
// For simplicity and backward compatibility, a numeric result is always returned (even if the input
// is non-numeric or an empty string).
{
	aResultToken.symbol = SYM_FLOAT;
	aResultToken.value_double = qmathCos(ParamIndexToDouble(0));
}



BIF_DECL(BIF_Tan)
// For simplicity and backward compatibility, a numeric result is always returned (even if the input
// is non-numeric or an empty string).
{
	aResultToken.symbol = SYM_FLOAT;
	aResultToken.value_double = qmathTan(ParamIndexToDouble(0));
}



BIF_DECL(BIF_ASinACos)
{
	double value = ParamIndexToDouble(0);
	if (value > 1 || value < -1) // ASin and ACos aren't defined for such values.
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
	}
	else
	{
		// For simplicity and backward compatibility, a numeric result is always returned in this case (even if
		// the input is non-numeric or an empty string).
		aResultToken.symbol = SYM_FLOAT;
		// Below: marker contains either "ASin" or "ACos"
		aResultToken.value_double = (ctoupper(aResultToken.marker[1]) == 'S') ? qmathAsin(value) : qmathAcos(value);
	}
}



BIF_DECL(BIF_ATan)
// For simplicity and backward compatibility, a numeric result is always returned (even if the input
// is non-numeric or an empty string).
{
	aResultToken.symbol = SYM_FLOAT;
	aResultToken.value_double = qmathAtan(ParamIndexToDouble(0));
}



BIF_DECL(BIF_Exp)
// For simplicity and backward compatibility, a numeric result is always returned (even if the input
// is non-numeric or an empty string).
{
	aResultToken.symbol = SYM_FLOAT;
	aResultToken.value_double = qmathExp(ParamIndexToDouble(0));
}



BIF_DECL(BIF_SqrtLogLn)
{
	double value = ParamIndexToDouble(0);
	if (value < 0) // Result is undefined in these cases, so make blank to indicate.
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = _T("");
	}
	else
	{
		// For simplicity and backward compatibility, a numeric result is always returned in this case (even if
		// the input is non-numeric or an empty string).
		aResultToken.symbol = SYM_FLOAT;
		switch (ctoupper(aResultToken.marker[1]))
		{
		case 'Q': // S[q]rt
			aResultToken.value_double = qmathSqrt(value);
			break;
		case 'O': // L[o]g
			aResultToken.value_double = qmathLog10(value);
			break;
		default: // L[n]
			aResultToken.value_double = qmathLog(value);
		}
	}
}



BIF_DECL(BIF_OnMessage)
// Returns: An empty string on failure or the name of a function (depends on mode) on success.
// Parameters:
// 1: Message number to monitor.
// 2: Name of the function that will monitor the message.
// 3: Maximum threads and "register first" flag.
{
	LPTSTR buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).
	// Set default result in case of early return; a blank value:
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");

	// Load-time validation has ensured there's at least one parameter for use here:
	UINT specified_msg = (UINT)ParamIndexToInt64(0); // Parameter #1

	// Set defaults:
	IObject *callback = NULL;
	bool mode_is_delete = false;
	bool legacy_mode = true;
	int max_instances = 1;
	bool call_it_last = true;

	if (!ParamIndexIsOmitted(2)) // Parameter #3 is present.
	{
		max_instances = (int)ParamIndexToInt64(2);
		// For backward-compatibility, values between MAX_INSTANCES+1 and SHORT_MAX must be supported.
		if (max_instances > MsgMonitorStruct::MAX_INSTANCES) // MAX_INSTANCES >= MAX_THREADS_LIMIT.
			max_instances = MsgMonitorStruct::MAX_INSTANCES;
		if (max_instances < 0) // MaxThreads < 0 is a signal to assign this monitor the lowest priority.
		{
			call_it_last = false; // Call it after any older monitors.  No effect if already registered.
			max_instances = -max_instances; // Convert to positive.
		}
		else if (max_instances == 0) // It would never be called, so this is used as a signal to delete the item.
			mode_is_delete = true;
	}

	if (!ParamIndexIsOmitted(1)) // Parameter #2 is present.
	{
		Func *func; // Func for validation of parameters, where possible.
		if (TokenIsEmptyString(*aParam[1])) // Explicitly blank function name ("") means delete this item.  By contrast, an omitted second parameter means "give me current function of this message".
		{
			mode_is_delete = true;
			func = NULL;
		}
		else if (callback = TokenToObject(*aParam[1]))
		{
			func = dynamic_cast<Func *>(callback);
			legacy_mode = false; // Since the caller passed a reference, use the new mode.
		}
		else
		{
			callback = func = g_script.FindFunc(TokenToString(*aParam[1]));
		}
		// Notes about func validation: ByRef and optional parameters are allowed for flexibility.
		// For example, a function may be called directly by the script to set static vars which
		// are used when a message arrives.  Raising an error might help catch bugs, but only in
		// very rare cases where a valid but wrong function name is given *and* that function has
		// ByRef or optional parameters.
		// If the parameter was not an empty string, an object or a valid function...
		if (!mode_is_delete && (!callback || func && (func->mIsBuiltIn || func->mMinParams > 4)))
		{
			if (!legacy_mode)
				aResult = g_script.ScriptError(ERR_PARAM2_INVALID);
			return; // Yield the default return value set earlier.
		}
	}

	// Check if this message already exists in the array:
	MsgMonitorStruct *pmonitor = g_MsgMonitor.Find(specified_msg, callback, legacy_mode);
	bool item_already_exists = (pmonitor != NULL);
	if (!item_already_exists)
	{
		if (!callback || mode_is_delete) // Delete or report function-name of a non-existent item.
			return; // Yield the default return value set earlier (an empty string).
		// From this point on, it is certain that an item will be added to the array.
		if (  !(pmonitor = g_MsgMonitor.Add(specified_msg, callback, legacy_mode, call_it_last))  )
		{
			if (!legacy_mode)
				aResult = g_script.ScriptError(ERR_OUTOFMEM);
			// Otherwise, indicate failure by yielding the default return value set earlier.
			return;
		}
	}

	MsgMonitorStruct &monitor = *pmonitor;

	if (item_already_exists)
	{
		if (legacy_mode) // Implies monitor.is_legacy_monitor, which means a Func was registered by name.
			// In all cases, yield the OLD function's name as the return value:
			aResultToken.marker = ((Func *)monitor.func)->mName;
		if (mode_is_delete)
		{
			// The msg-monitor is deleted from the array for two reasons:
			// 1) It improves performance because every incoming message for the app now needs to be compared
			//    to one less filter. If the count will now be zero, performance is improved even more because
			//    the overhead of the call to MsgMonitor() is completely avoided for every incoming message.
			// 2) It conserves space in the array in a situation where the script creates hundreds of
			//    msg-monitors and then later deletes them, then later creates hundreds of filters for entirely
			//    different message numbers.
			// The main disadvantage to deleting message filters from the array is that the deletion might
			// occur while the monitor is currently running, which requires more complex handling within
			// MsgMonitor() (see its comments for details).
			g_MsgMonitor.Remove(pmonitor);
			return;
		}
		if (aParamCount < 2) // Single-parameter mode: Report existing item's function name.
			return;
		// Otherwise, an existing item is being assigned a new function or MaxThreads limit.
		// Continue on to update this item's attributes.
	}
	else // This message was newly added to the array.
	{
		// The above already verified that callback is not NULL and there is room in the array.
		if (legacy_mode)
			// For backward-compatibility, return the function's name on success:
			aResultToken.marker = ((Func *)callback)->mName;
		monitor.instance_count = 0; // Reset instance_count only for new items since existing items might currently be running.
		// Continue on to the update-or-create logic below.
	}

	// Since above didn't return, above has ensured that msg_index is the index of the existing or new
	// MsgMonitorStruct in the array.  In addition, it has set the proper return value for us.
	// Update those struct attributes that get the same treatment regardless of whether this is an update or creation.
	if (callback && callback != monitor.func) // Callback is being registered or changed.
	{
		callback->AddRef(); // Keep the object alive while it's in g_MsgMonitor.
		if (monitor.func)
			monitor.func->Release();
		monitor.func = callback;
	}
	if (!item_already_exists || !ParamIndexIsOmitted(2))
		monitor.max_instances = max_instances;
	// Otherwise, the parameter was omitted so leave max_instances at its current value.
}


MsgMonitorStruct *MsgMonitorList::Find(UINT aMsg, IObject *aCallback, bool aIsLegacyMode)
{
	for (int i = 0; i < mCount; ++i)
		if (mMonitor[i].msg == aMsg
			&& (aIsLegacyMode ? mMonitor[i].is_legacy_monitor : mMonitor[i].func == aCallback))
			return mMonitor + i;
	return NULL;
}


MsgMonitorStruct *MsgMonitorList::Add(UINT aMsg, IObject *aCallback, bool aIsLegacyMode, bool aAppend)
{
	if (mCount == mCountMax)
	{
		int new_count = mCountMax ? mCountMax * mCountMax : 16;
		void *new_array = realloc(mMonitor, new_count * sizeof(MsgMonitorStruct));
		if (!new_array)
			return NULL;
		mMonitor = (MsgMonitorStruct *)new_array;
		mCountMax = new_count;
	}
	MsgMonitorStruct *new_mon;
	if (!aAppend)
	{
		for (MsgMonitorInstance *inst = mTop; inst; inst = inst->previous)
		{
			inst->index++; // Correct the index of each running monitor.
			inst->count++; // Iterate the same set of items which existed before.
			// By contrast, count isn't adjusted when adding at the end because we do not
			// want new items to be called by messages received before they were registered.
		}
		// Shift existing items to make room.
		memmove(mMonitor + 1, mMonitor, mCount * sizeof(MsgMonitorStruct));
		new_mon = mMonitor;
	}
	else
		new_mon = mMonitor + mCount;

	++mCount;
	aCallback->AddRef();
	new_mon->func = aCallback;
	new_mon->msg = aMsg;
	//new_mon->instance_count = 0;
	//new_mon->max_instances = 1;
	new_mon->is_legacy_monitor = aIsLegacyMode;
	return new_mon;
}


void MsgMonitorList::Remove(MsgMonitorStruct *aMonitor)
{
	ASSERT(aMonitor >= mMonitor && aMonitor < mMonitor + mCount);

	int mon_index = int(aMonitor - mMonitor);
	// Adjust the index of any active message monitors affected by this deletion.  This allows a
	// message monitor to delete older message monitors while still allowing any remaining monitors
	// of that message to be called (when there are multiple).
	for (MsgMonitorInstance *inst = mTop; inst; inst = inst->previous)
	{
		if (inst->index >= mon_index && inst->index >= 0)
			inst->index--; // So index+1 is the next item.
		inst->count--;
	}
	// Remove the item from the array.
	--mCount;  // Must be done prior to the below.
	IObject *release_me = aMonitor->func;
	if (mon_index < mCount) // An element other than the last is being removed. Shift the array to cover/delete it.
		memmove(aMonitor, aMonitor + 1, (mCount - mon_index) * sizeof(MsgMonitorStruct));
	release_me->Release(); // Must be called after the above in case it calls a __delete() meta-function.
}


BIF_DECL(BIF_On)
{
	enum OnEventType {
		OnExit,
		OnError,
		OnClipboardChange
	} event_type;
	MsgMonitorList *phandlers;
	switch (tolower(aResultToken.marker[3]))
	{
	case 'r':
		event_type = OnError;
		phandlers = &g_script.mOnError;
		break;
	case 'l':
		event_type = OnClipboardChange;
		phandlers = &g_script.mOnClipboardChange;
		break;
	default:
		event_type = OnExit;
		phandlers = &g_script.mOnExit;
		break;
	}
	MsgMonitorList &handlers = *phandlers;

	aResultToken.SetValue(_T("")); // In all cases there is no return value.

	IObject *callback;
	if (callback = TokenToFunc(*aParam[0]))
	{
		// Ensure this function is a valid one.
		if (((Func *)callback)->mMinParams > (event_type == OnExit ? 2 : 1))
			callback = NULL;
	}
	else
		callback = TokenToObject(*aParam[0]);
	if (!callback)
	{
		aResult = g_script.ScriptError(ERR_PARAM1_INVALID);
		return;
	}
	
	int mode = 1; // Default.
	if (!ParamIndexIsOmitted(1))
		mode = ParamIndexToInt(1);

	MsgMonitorStruct *existing = handlers.Find(0, callback, false);

	switch (mode)
	{
	case  1:
	case -1:
		if (existing)
			return;
		if (event_type == OnClipboardChange)
		{
			// Do this before adding the handler so that it won't be called as a result of the
			// SetClipboardViewer() call on Windows XP.  This won't cause existing handlers to
			// be called because in that case the clipboard listener is already enabled.
			g_script.EnableClipboardListener(true);
		}
		if (!handlers.Add(0, callback, false, mode == 1))
			aResult = g_script.ScriptError(ERR_OUTOFMEM);
		break;
	case  0:
		if (existing)
			handlers.Remove(existing);
		break;
	default:
		aResult = g_script.ScriptError(ERR_PARAM2_INVALID);
		break;
	}
	// In case the above enabled the clipboard listener but failed to add the handler,
	// do this even if mode != 0:
	if (event_type == OnClipboardChange && !g_script.mOnClipboardChangeLabel && !handlers.Count())
		g_script.EnableClipboardListener(false);
}


#ifdef ENABLE_REGISTERCALLBACK
struct RCCallbackFunc // Used by BIF_RegisterCallback() and related.
{
#ifdef WIN32_PLATFORM
	ULONG data1;	//E8 00 00 00
	ULONG data2;	//00 8D 44 24
	ULONG data3;	//08 50 FF 15
	UINT_PTR (CALLBACK **callfuncptr)(UINT_PTR*, char*);
	ULONG data4;	//59 84 C4 nn
	USHORT data5;	//FF E1
#endif
#ifdef _WIN64
	UINT64 data1; // 0xfffffffff9058d48
	UINT64 data2; // 0x9090900000000325
	void (*stub)();
	UINT_PTR (CALLBACK *callfuncptr)(UINT_PTR*, char*);
#endif
	//code ends
	UCHAR actual_param_count; // This is the actual (not formal) number of parameters passed from the caller to the callback. Kept adjacent to the USHORT above to conserve memory due to 4-byte struct alignment.
	bool create_new_thread; // Kept adjacent to above to conserve memory due to 4-byte struct alignment.
	EventInfoType event_info; // A_EventInfo
	Func *func; // The UDF to be called whenever the callback's caller calls callfuncptr.
};

#ifdef _WIN64
extern "C" void RegisterCallbackAsmStub();
#endif


UINT_PTR CALLBACK RegisterCallbackCStub(UINT_PTR *params, char *address) // Used by BIF_RegisterCallback().
// JGR: On Win32 parameters are always 4 bytes wide. The exceptions are functions which work on the FPU stack
// (not many of those). Win32 is quite picky about the stack always being 4 byte-aligned, (I have seen only one
// application which defied that and it was a patched ported DOS mixed mode application). The Win32 calling
// convention assumes that the parameter size equals the pointer size. 64 integers on Win32 are passed on
// pointers, or as two 32 bit halves for some functions...
{
	#define DEFAULT_CB_RETURN_VALUE 0  // The value returned to the callback's caller if script doesn't provide one.

#ifdef WIN32_PLATFORM
	RCCallbackFunc &cb = *((RCCallbackFunc*)(address-5)); //second instruction is 5 bytes after start (return address pushed by call)
#else
	RCCallbackFunc &cb = *((RCCallbackFunc*) address);
#endif
	Func &func = *cb.func; // For performance and convenience.

	TCHAR ErrorLevel_saved[ERRORLEVEL_SAVED_SIZE];
	EventInfoType EventInfo_saved;
	BOOL pause_after_execute;

	// NOTES ABOUT INTERRUPTIONS / CRITICAL:
	// An incoming call to a callback is considered an "emergency" for the purpose of determining whether
	// critical/high-priority threads should be interrupted because there's no way easy way to buffer or
	// postpone the call.  Therefore, NO check of the following is done here:
	// - Current thread's priority (that's something of a deprecated feature anyway).
	// - Current thread's status of Critical (however, Critical would prevent us from ever being called in
	//   cases where the callback is triggered indirectly via message/dispatch due to message filtering
	//   and/or Critical's ability to pump messes less often).
	// - INTERRUPTIBLE_IN_EMERGENCY (which includes g_MenuIsVisible and g_AllowInterruption), which primarily
	//   affects SLEEP_WITHOUT_INTERRUPTION): It's debatable, but to maximize flexibility it seems best to allow
	//   callbacks during the display of a menu and during SLEEP_WITHOUT_INTERRUPTION.  For most callers of
	//   SLEEP_WITHOUT_INTERRUPTION, interruptions seem harmless.  For some it could be a problem, but when you
	//   consider how rare such callbacks are (mostly just subclassing of windows/controls) and what those
	//   callbacks tend to do, conflicts seem very rare.
	// Of course, a callback can also be triggered through explicit script action such as a DllCall of
	// EnumWindows, in which case the script would want to be interrupted unconditionally to make the call.
	// However, in those cases it's hard to imagine that INTERRUPTIBLE_IN_EMERGENCY wouldn't be true anyway.
	if (cb.create_new_thread)
	{
		if (g_nThreads >= g_MaxThreadsTotal) // Since this is a callback, it seems too rare to make an exemption for functions whose first line is ExitApp. In any case, to avoid array overflow, g_MaxThreadsTotal must not be exceeded except where otherwise documented.
			return DEFAULT_CB_RETURN_VALUE;
		// See MsgSleep() for comments about the following section.
		tcslcpy(ErrorLevel_saved, g_ErrorLevel->Contents(), _countof(ErrorLevel_saved));
		InitNewThread(0, false, true, func.mJumpToLine->mActionType);
		DEBUGGER_STACK_PUSH(_T("Callback"))
	}
	else // Backup/restore only A_EventInfo. This avoids callbacks changing A_EventInfo for the current thread/context (that would be counterintuitive and a source of script bugs).
	{
		EventInfo_saved = g->EventInfo;
		if (pause_after_execute = g->IsPaused) // Assign.
		{
			// v1.0.48: If the current thread is paused, this threadless callback would get stuck in
			// ExecUntil()'s pause loop (keep in mind that this situation happens only when a fast-mode
			// callback has been created without a script thread to control it, which goes against the
			// advice in the documentation). To avoid that, it seems best to temporarily unpause the
			// thread until the callback finishes.  But for performance, tray icon color isn't updated.
			g->IsPaused = false;
			--g_nPausedThreads; // See below.
			// If g_nPausedThreads isn't adjusted here, g_nPausedThreads could become corrupted if the
			// callback (or some thread that interrupts it) uses the Pause command/menu-item because
			// those aren't designed to deal with g->IsPaused being out-of-sync with g_nPausedThreads.
			// However, if --g_nPausedThreads reduces g_nPausedThreads to 0, timers would allowed to
			// run during the callback.  But that seems like the lesser evil, especially given that
			// this whole situation is very rare, and the documentation advises against doing it.
		}
		//else the current thread wasn't paused, which is usually the case.
		// TRAY ICON: g_script.UpdateTrayIcon() is not called because it's already in the right state
		// except when pause_after_execute==true, in which case it seems best not to change the icon
		// because it's likely to hurt any callback that's performance-sensitive.
	}

	g->EventInfo = cb.event_info; // This is the means to identify which caller called the callback (if the script assigned more than one caller to this callback).

	// For performance and to preserve stack space, the indirect method of calling a function via the new
	// Func::Call overload is not used here.  Using it would only be necessary to support variadic functions,
	// which have very limited use as callbacks; instead, we pass such functions a pointer to surplus params.

	// Need to check if backup of function's variables is needed in case:
	// 1) The UDF is assigned to more than one callback, in which case the UDF could be running more than once
	//    simultaneously.
	// 2) The callback is intended to be reentrant (e.g. a subclass/WindowProc that doesn't use Critical).
	// 3) Script explicitly calls the UDF in addition to using it as a callback.
	//
	// See ExpandExpression() for detailed comments about the following section.
	VarBkp *var_backup = NULL;  // If needed, it will hold an array of VarBkp objects.
	int var_backup_count; // The number of items in the above array.
	if (func.mInstances > 0) // Backup is needed (see above for explanation).
		if (!Var::BackupFunctionVars(func, var_backup, var_backup_count)) // Out of memory.
			return DEFAULT_CB_RETURN_VALUE; // Since out-of-memory is so rare, it seems justifiable not to have any error reporting and instead just avoid calling the function.

	// The following section is similar to the one in ExpandExpression().  See it for detailed comments.
	int i, j = cb.actual_param_count < func.mParamCount ? cb.actual_param_count : func.mParamCount;
	for (i = 0; i < j; ++i)  // For each formal parameter that has a matching actual.
		func.mParam[i].var->Assign((UINT_PTR)params[i]); // All parameters are passed "by value" because an earlier stage ensured there are no ByRef parameters.
	if (func.mIsVariadic)
		// See the performance note further above.  Rather than having the "variadic" param remain empty,
		// pass it a pointer to the first actual parameter which wasn't assigned to a formal parameter:
		func.mParam[func.mParamCount].var->Assign((UINT_PTR)(params + i));
	for (; i < func.mParamCount; ++i) // For each remaining formal (i.e. those that lack actuals), apply a default value (an earlier stage verified that all such parameters have a default-value available).
	{
		FuncParam &this_formal_param = func.mParam[i]; // For performance and convenience.
		// The following isn't necessary because an earlier stage has already ensured that there
		// are no ByRef parameters in a callback:
		//if (this_formal_param.is_byref)
		//	this_formal_param.var->ConvertToNonAliasIfNecessary();
		switch(this_formal_param.default_type)
		{
		case PARAM_DEFAULT_STR:   this_formal_param.var->Assign(this_formal_param.default_str);    break;
		case PARAM_DEFAULT_INT:   this_formal_param.var->Assign(this_formal_param.default_int64);  break;
		case PARAM_DEFAULT_FLOAT: this_formal_param.var->Assign(this_formal_param.default_double); break;
		//case PARAM_DEFAULT_NONE: Not possible due to validation at an earlier stage.
		}
	}

	g_script.mLastScriptRest = g_script.mLastPeekTime = GetTickCount(); // Somewhat debatable, but might help minimize interruptions when the callback is called via message (e.g. subclassing a control; overriding a WindowProc).

	ExprTokenType result_token; // L31
	func.Call(&result_token); // Call the UDF.  Call()'s own return value (e.g. EARLY_EXIT or FAIL) is ignored because it wouldn't affect the handling below.

	UINT_PTR number_to_return = (UINT_PTR)TokenToInt64(result_token); // L31: For simplicity, DEFAULT_CB_RETURN_VALUE is not used - DEFAULT_CB_RETURN_VALUE is 0, which TokenToInt64 will return if the token is empty.
	if (result_token.symbol == SYM_OBJECT) // L31
		result_token.object->Release();

	Var::FreeAndRestoreFunctionVars(func, var_backup, var_backup_count); // ABOVE must be done BEFORE this because return_value might be the contents of one of the function's local variables (which are about to be free'd).

	if (cb.create_new_thread)
	{
		DEBUGGER_STACK_POP()
		ResumeUnderlyingThread(ErrorLevel_saved);
	}
	else
	{
		g->EventInfo = EventInfo_saved;
		if (g == g_array && !g_script.mAutoExecSectionIsRunning)
			// If the function just called used thread #0 and the AutoExec section isn't running, that means
			// the AutoExec section definitely didn't launch or control the callback (even if it is running,
			// it's not 100% certain it launched the callback). This can happen when a fast-mode callback has
			// been invoked via message, though the documentation advises against the fast mode when there is
			// no script thread controlling the callback.
			global_maximize_interruptibility(*g); // In case the script function called above used commands like Critical or "Thread Interrupt", ensure the idle thread is interruptible.  This avoids having to treat the idle thread as special in other places.
		//else never alter the interruptibility of AutoExec while it's running because it has its own method to do that.
		if (pause_after_execute) // See comments where it's defined.
		{
			g->IsPaused = true;
			++g_nPausedThreads;
		}
	}

	return number_to_return; //return integer value to callback stub
}



BIF_DECL(BIF_RegisterCallback)
// Returns: Address of callback procedure, or empty string on failure.
// Parameters:
// 1: Name of the function to be called when the callback routine is executed.
// 2: Options.
// 3: Number of parameters of callback.
// 4: EventInfo: a DWORD set for use by UDF to identify the caller (in case more than one caller).
//
// Author: RegisterCallback() was created by Jonathan Rennison (JGR).
{
	// Set default result in case of early return; a blank value:
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");

	// Loadtime validation has ensured that at least 1 parameter is present.
	Func *func;
	if (  !(func = TokenToFunc(*aParam[0])) || func->mIsBuiltIn  )  // Not a valid user-defined function.
		return; // Indicate failure by yielding the default result set earlier.

	LPTSTR options = ParamIndexToOptionalString(1);
	int actual_param_count;
	if (!ParamIndexIsOmittedOrEmpty(2)) // A parameter count was specified.
	{
		actual_param_count = ParamIndexToInt(2);
		if (   actual_param_count > func->mParamCount    // The function doesn't have enough formals to cover the specified number of actuals.
				&& !func->mIsVariadic					 // ...and the function isn't designed to accept parameters via an array (or in this case, a pointer).
			|| actual_param_count < func->mMinParams   ) // ...or the function has too many mandatory formals (caller specified insufficient actuals to cover them all).
			return; // Indicate failure by yielding the default result set earlier.
	}
	else // Default to the number of mandatory formal parameters in the function's definition.
		actual_param_count = func->mMinParams;

#ifdef WIN32_PLATFORM
	bool use_cdecl = StrChrAny(options, _T("Cc")); // Recognize "C" as the "CDecl" option.
	if (!use_cdecl && actual_param_count > 31) // The ASM instruction currently used limits parameters to 31 (which should be plenty for any realistic use).
		return; // Indicate failure by yielding the default result set earlier.
#endif

	// To improve callback performance, ensure there are no ByRef parameters (for simplicity: not even ones that
	// have default values).  This avoids the need to ensure formal parameters are non-aliases each time the
	// callback is called.
	for (int i = 0; i < func->mParamCount; ++i)
		if (func->mParam[i].is_byref)
			return; // Yield the default return value set earlier.

	// GlobalAlloc() and dynamically-built code is the means by which a script can have an unlimited number of
	// distinct callbacks. On Win32, GlobalAlloc is the same function as LocalAlloc: they both point to
	// RtlAllocateHeap on the process heap. For large chunks of code you would reserve a 64K section with
	// VirtualAlloc and fill that, but for the 32 bytes we use here that would be overkill; GlobalAlloc is
	// much more efficient. MSDN says about GlobalAlloc: "All memory is created with execute access; no
	// special function is required to execute dynamically generated code. Memory allocated with this function
	// is guaranteed to be aligned on an 8-byte boundary." 
	// ABOVE IS OBSOLETE/INACCURATE: Systems with DEP enabled (and some without) require a VirtualProtect call
	// to allow the callback to execute.  MSDN currently says only this about the topic in the documentation
	// for GlobalAlloc:  "To execute dynamically generated code, use the VirtualAlloc function to allocate
	//						memory and the VirtualProtect function to grant PAGE_EXECUTE access."
	RCCallbackFunc *callbackfunc=(RCCallbackFunc*) GlobalAlloc(GMEM_FIXED,sizeof(RCCallbackFunc));	//allocate structure off process heap, automatically RWE and fixed.
	if(!callbackfunc) return;
	RCCallbackFunc &cb = *callbackfunc; // For convenience and possible code-size reduction.

#ifdef WIN32_PLATFORM
	cb.data1=0xE8;       // call +0 -- E8 00 00 00 00 ;get eip, stays on stack as parameter 2 for C function (char *address).
	cb.data2=0x24448D00; // lea eax, [esp+8] -- 8D 44 24 08 ;eax points to params
	cb.data3=0x15FF5008; // push eax -- 50 ;eax pushed on stack as parameter 1 for C stub (UINT *params)
                         // call [xxxx] (in the lines below) -- FF 15 xx xx xx xx ;call C stub __stdcall, so stack cleaned up for us.

	// Comments about the static variable below: The reason for using the address of a pointer to a function,
	// is that the address is passed as a fixed address, whereas a direct call is passed as a 32-bit offset
	// relative to the beginning of the next instruction, which is more fiddly than it's worth to calculate
	// for dynamic code, as a relative call is designed to allow position independent calls to within the
	// same memory block without requiring dynamic fixups, or other such inconveniences.  In essence:
	//    call xxx ; is relative
	//    call [ptr_xxx] ; is position independent
	// Typically the latter is used when calling imported functions, etc., as only the pointers (import table),
	// need to be adjusted, not the calls themselves...

	static UINT_PTR (CALLBACK *funcaddrptr)(UINT_PTR*, char*) = RegisterCallbackCStub; // Use fixed absolute address of pointer to function, instead of varying relative offset to function.
	cb.callfuncptr = &funcaddrptr; // xxxx: Address of C stub.

	cb.data4=0xC48359 // pop ecx -- 59 ;return address... add esp, xx -- 83 C4 xx ;stack correct (add argument to add esp, nn for stack correction).
		+ (use_cdecl ? 0 : actual_param_count<<26);

	cb.data5=0xE1FF; // jmp ecx -- FF E1 ;return
#endif

#ifdef _WIN64
	/* Adapted from http://www.dyncall.org/
		lea rax, (rip)  # copy RIP (=p?) to RAX and use address in
		jmp [rax+16]    # 'entry' (stored at RIP+16) for jump
		nop
		nop
		nop
	*/
	cb.data1 = 0xfffffffff9058d48ULL;
	cb.data2 = 0x9090900000000325ULL;
	cb.stub = RegisterCallbackAsmStub;
	cb.callfuncptr = RegisterCallbackCStub;
#endif

	cb.event_info = (EventInfoType)ParamIndexToOptionalInt64(3, (size_t)callbackfunc);
	cb.func = func;
	cb.actual_param_count = actual_param_count;
	cb.create_new_thread = !StrChrAny(options, _T("Ff")); // Recognize "F" as the "fast" mode that avoids creating a new thread.

	// If DEP is enabled (and sometimes when DEP is apparently "disabled"), we must change the
	// protection of the page of memory in which the callback resides to allow it to execute:
	DWORD dwOldProtect;
	VirtualProtect(callbackfunc, sizeof(RCCallbackFunc), PAGE_EXECUTE_READWRITE, &dwOldProtect);

	aResultToken.symbol = SYM_INTEGER; // Override the default set earlier.
	aResultToken.value_int64 = (__int64)callbackfunc; // Yield the callable address as the result.
}

#endif



BIF_DECL(BIF_MenuGet)
{
	if (ctoupper(aResultToken.marker[7]) == 'H') // MenuGetHandle
	{
		UserMenu *menu = g_script.FindMenu(ParamIndexToString(0, aResultToken.buf));
		if (menu && !menu->mMenu)
			menu->Create(); // On failure (rare), we just return 0.
		aResultToken.value_int64 = menu ? (__int64)(UINT_PTR)menu->mMenu : 0;
	}
	else // MenuGetName
	{
		UserMenu *menu = g_script.FindMenu((HMENU)ParamIndexToInt64(0));
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = menu ? menu->mName : _T("");
	}
}



BIF_DECL(BIF_StatusBar)
{
	TCHAR mode = ctoupper(aResultToken.marker[6]); // Union's marker initially contains the function name. SB_Set[T]ext.
	LPTSTR buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).
	aResultToken.value_int64 = 0; // Set default return value. Must be done only after consulting marker above.
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no StatusBar in window).

	if (!g->GuiDefaultWindowValid()) // Always operate on thread's default window to simplify the syntax.
		return;
	GuiType& gui = *g->GuiDefaultWindow; // For performance.
	HWND control_hwnd;
	if (   !(control_hwnd = gui.mStatusBarHwnd)   )
		return;

	HICON hicon;
	switch(mode)
	{
	case 'T': // SB_SetText()
		aResultToken.value_int64 = SendMessage(control_hwnd, SB_SETTEXT
			, (WPARAM)((ParamIndexIsOmitted(1) ? 0 : ParamIndexToInt64(1) - 1) // The Part# param is present.
				     | (ParamIndexIsOmitted(2) ? 0 : ParamIndexToInt64(2) << 8)) // The uType parameter is present.
			, (LPARAM)ParamIndexToString(0, buf)); // Load-time validation has ensured that there's at least one param in this mode.
		break;

	case 'P': // SB_SetParts()
		LRESULT old_part_count, new_part_count;
		int edge, part[256]; // Load-time validation has ensured aParamCount is under 255, so it shouldn't overflow.
		for (edge = 0, new_part_count = 0; new_part_count < aParamCount; ++new_part_count)
		{
			edge += gui.Scale(ParamIndexToInt(new_part_count)); // For code simplicity, no check for negative (seems fairly harmless since the bar will simply show up with the wrong number of parts to indicate the problem).
			part[new_part_count] = edge;
		}
		// For code simplicity, there is currently no means to have the last part of the bar use less than
		// all of the bar's remaining width.  The desire to do so seems rare, especially since the script can
		// add an extra/unused part at the end to achieve nearly (or perhaps exactly) the same effect.
		part[new_part_count++] = -1; // Make the last part use the remaining width of the bar.

		old_part_count = SendMessage(control_hwnd, SB_GETPARTS, 0, NULL); // MSDN: "This message always returns the number of parts in the status bar [regardless of how it is called]".
		if (old_part_count > new_part_count) // Some parts are being deleted, so destroy their icons.  See other notes in GuiType::Destroy() for explanation.
			for (LRESULT i = new_part_count; i < old_part_count; ++i) // Verified correct.
				if (hicon = (HICON)SendMessage(control_hwnd, SB_GETICON, i, 0))
					DestroyIcon(hicon);

		aResultToken.value_int64 = SendMessage(control_hwnd, SB_SETPARTS, new_part_count, (LPARAM)part)
			? (__int64)control_hwnd : 0; // Return HWND to provide an easy means for the script to get the bar's HWND.
		break;

	case 'I': // SB_SetIcon()
		int unused, icon_number;
		icon_number = ParamIndexToOptionalInt(1, 1);
		if (icon_number == 0) // Must be != 0 to tell LoadPicture that "icon must be loaded, never a bitmap".
			icon_number = 1;
		if (hicon = (HICON)LoadPicture(ParamIndexToString(0, buf) // Load-time validation has ensured there is at least one parameter.
			, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON) // Apparently the bar won't scale them for us.
			, unused, icon_number, false)) // Defaulting to "false" for "use GDIplus" provides more consistent appearance across multiple OSes.
		{
			WPARAM part_index = ParamIndexIsOmitted(2) ? 0 : (WPARAM)ParamIndexToInt64(2) - 1;
			HICON hicon_old = (HICON)SendMessage(control_hwnd, SB_GETICON, part_index, 0); // Get the old one before setting the new one.
			// For code simplicity, the script is responsible for destroying the hicon later, if it ever destroys
			// the window.  Though in practice, most people probably won't do this, which is usually okay (if the
			// script doesn't load too many) since they're all destroyed by the system upon program termination.
			if (SendMessage(control_hwnd, SB_SETICON, part_index, (LPARAM)hicon))
			{
				aResultToken.value_int64 = (__int64)hicon; // Override the 0 default. This allows the script to destroy the HICON later when it doesn't need it (see comments above too).
				if (hicon_old)
					// Although the old icon is automatically destroyed here, the script can call SendMessage(SB_SETICON)
					// itself if it wants to work with HICONs directly (for performance reasons, etc.)
					DestroyIcon(hicon_old);
			}
			else
				DestroyIcon(hicon);
				//And leave aResultToken.value_int64 at its default value.
		}
		//else can't load icon, so leave aResultToken.value_int64 at its default value.
		break;
	// SB_SetTipText() not implemented (though can be done via SendMessage in the script) because the conditions
	// under which tooltips are displayed don't seem like something a script would want very often:
	// This ToolTip text is displayed in two situations: 
	// When the corresponding pane in the status bar contains only an icon. 
	// When the corresponding pane in the status bar contains text that is truncated due to the size of the pane.
	// In spite of the above, SB_SETTIPTEXT doesn't actually seem to do anything, even when the text is too long
	// to fit in a narrowed part, tooltip text has been set, and the user hovers the cursor over the bar.  Maybe
	// I'm not doing it right or maybe this feature is somehow disabled under certain service packs or conditions.
	//case 'T': // SB_SetTipText()
	//	break;
	} // switch(mode)
}



BIF_DECL(BIF_LV_GetNextOrCount)
// LV_GetNext:
// Returns: The index of the found item, or 0 on failure.
// Parameters:
// 1: Starting index (one-based when it comes in).  If absent, search starts at the top.
// 2: Options string.
// 3: (FUTURE): Possible for use with LV_FindItem (though I think it can only search item text, not subitem text).
{
	bool mode_is_count = ctoupper(aResultToken.marker[6]) == 'C'; // Union's marker initially contains the function name. LV_Get[C]ount.  Bug-fixed for v1.0.43.09.
	LPTSTR buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).
	aResultToken.value_int64 = 0; // Set default return value. Must be done only after consulting marker above.
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no ListView in window).
	// Item not found in ListView.

	if (!g->GuiDefaultWindowValid())
		return;
	GuiType &gui = *g->GuiDefaultWindow; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentListView)
		return;
	HWND control_hwnd = gui.mCurrentListView->hwnd;

	LPTSTR options;
	if (mode_is_count)
	{
		options = (aParamCount > 0) ? omit_leading_whitespace(ParamIndexToString(0, buf)) : _T("");
		if (*options)
		{
			if (ctoupper(*options) == 'S')
				aResultToken.value_int64 = SendMessage(control_hwnd, LVM_GETSELECTEDCOUNT, 0, 0);
			else if (!_tcsnicmp(options, _T("Col"), 3)) // "Col" or "Column". Don't allow "C" by itself, so that "Checked" can be added in the future.
				aResultToken.value_int64 = gui.mCurrentListView->union_lv_attrib->col_count;
			//else some unsupported value, leave aResultToken.value_int64 set to zero to indicate failure.
		}
		else
			aResultToken.value_int64 = SendMessage(control_hwnd, LVM_GETITEMCOUNT, 0, 0);
		return;
	}
	// Since above didn't return, this is GetNext() mode.

	int index = ParamIndexToOptionalInt(0, 0) - 1; // -1 to convert to zero-based.
	// For flexibility, allow index to be less than -1 to avoid first-iteration complications in script loops
	// (such as when deleting rows, which shifts the row index upward, require the search to resume at
	// the previously found index rather than the row after it).  However, reset it to -1 to ensure
	// proper return values from the API in the "find checked item" mode used below.
	if (index < -1)
		index = -1;  // Signal it to start at the top.

	// For performance, decided to always find next selected item when the "C" option hasn't been specified,
	// even when the checkboxes style is in effect.  Otherwise, would have to fetch and check checkbox style
	// bit for each call, which would slow down this heavily-called function.

	options = ParamIndexToOptionalString(1, buf);
	TCHAR first_char = ctoupper(*omit_leading_whitespace(options));
	// To retain compatibility in the future, also allow "Check(ed)" and "Focus(ed)" since any word that
	// starts with C or F is already supported.

	switch(first_char)
	{
	case '\0': // Listed first for performance.
	case 'F':
		aResultToken.value_int64 = ListView_GetNextItem(control_hwnd, index
			, first_char ? LVNI_FOCUSED : LVNI_SELECTED) + 1; // +1 to convert to 1-based.
		break;
	case 'C': // Checkbox: Find checked items. For performance assume that the control really has checkboxes.
		int item_count = ListView_GetItemCount(control_hwnd);
		for (int i = index + 1; i < item_count; ++i) // Start at index+1 to omit the first item from the search (for consistency with the other mode above).
			if (ListView_GetCheckState(control_hwnd, i)) // Item's box is checked.
			{
				aResultToken.value_int64 = i + 1; // +1 to convert from zero-based to one-based.
				return;
			}
		// Since above didn't return, no match found.  The 0/false value previously set as the default is retained.
		break;
	}
}



BIF_DECL(BIF_LV_GetText)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: Output variable (doing it this way allows success/fail return value to more closely mirror the API and
//    simplifies the code since there is currently no easy means of passing back large strings to our caller).
// 2: Row index (one-based when it comes in).
// 3: Column index (one-based when it comes in).
{
	aResultToken.value_int64 = 0; // Set default return value.
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no ListView in window).
	// Item not found in ListView.
	// And others.

	if (!g->GuiDefaultWindowValid())
		return;
	GuiType &gui = *g->GuiDefaultWindow; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentListView)
		return;
	// Caller has ensured there is at least two parameters:
	if (aParam[0]->symbol != SYM_VAR) // No output variable.  Supporting a NULL for the purpose of checking for the existence of a cell seems too rarely needed.
		return;

	// Caller has ensured there is at least two parameters.
	int row_index = ParamIndexToInt(1) - 1; // -1 to convert to zero-based.
	// If parameter 3 is omitted, default to the first column (index 0):
	int col_index = ParamIndexIsOmitted(2) ? 0 : ParamIndexToInt(2) - 1; // -1 to convert to zero-based.
	if (row_index < -1 || col_index < 0) // row_index==-1 is reserved to mean "get column heading's text".
		return;

	Var &output_var = *aParam[0]->var; // It was already ensured higher above that symbol==SYM_VAR.
	TCHAR buf[LV_TEXT_BUF_SIZE];

	if (row_index == -1) // Special mode to get column's text.
	{
		LVCOLUMN lvc;
		lvc.cchTextMax = LV_TEXT_BUF_SIZE - 1;  // See notes below about -1.
		lvc.pszText = buf;
		lvc.mask = LVCF_TEXT;
		if (aResultToken.value_int64 = SendMessage(gui.mCurrentListView->hwnd, LVM_GETCOLUMN, col_index, (LPARAM)&lvc)) // Assign.
			output_var.Assign(lvc.pszText); // See notes below about why pszText is used instead of buf (might apply to this too).
		else // On failure, it seems best to also clear the output var for better consistency and in case the script doesn't check the return value.
			output_var.Assign();
	}
	else // Get row's indicated item or subitem text.
	{
		LVITEM lvi;
		// Subtract 1 because of that nagging doubt about size vs. length. Some MSDN examples subtract one, such as
		// TabCtrl_GetItem()'s cchTextMax:
		lvi.iItem = row_index;
		lvi.iSubItem = col_index; // Which field to fetch.  If it's zero, the item vs. subitem will be fetched.
		lvi.mask = LVIF_TEXT;
		lvi.pszText = buf;
		lvi.cchTextMax = LV_TEXT_BUF_SIZE - 1; // Note that LVM_GETITEM doesn't update this member to reflect the new length.
		// Unlike LVM_GETITEMTEXT, LVM_GETITEM indicates success or failure, which seems more useful/preferable
		// as a return value since a text length of zero would be ambiguous: could be an empty field or a failure.
		if (aResultToken.value_int64 = SendMessage(gui.mCurrentListView->hwnd, LVM_GETITEM, 0, (LPARAM)&lvi)) // Assign
			// Must use lvi.pszText vs. buf because MSDN says: "Applications should not assume that the text will
			// necessarily be placed in the specified buffer. The control may instead change the pszText member
			// of the structure to point to the new text rather than place it in the buffer."
			output_var.Assign(lvi.pszText);
		else // On failure, it seems best to also clear the output var for better consistency and in case the script doesn't check the return value.
			output_var.Assign();
	}
}



BIF_DECL(BIF_LV_AddInsertModify)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: For Add(), this is the options.  For Insert/Modify, it's the row index (one-based when it comes in).
// 2: For Add(), this is the first field's text.  For Insert/Modify, it's the options.
// 3 and beyond: Additional field text.
// In Add/Insert mode, if there are no text fields present, a blank for is appended/inserted.
{
	TCHAR mode = ctoupper(aResultToken.marker[3]); // Union's marker initially contains the function name. e.g. LV_[I]nsert.
	LPTSTR buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).
	aResultToken.value_int64 = 0; // Set default return value. Must be done only after consulting marker above.
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no ListView in window).
	// And others as shown below.

	int index;
	if (mode == 'A') // For Add mode (A), use INT_MAX as a signal to append the item rather than inserting it.
	{
		index = INT_MAX;
		mode = 'I'; // Add has now been set up to be the same as insert, so change the mode to simplify other things.
	}
	else // Insert or Modify: the target row-index is their first parameter, which load-time has ensured is present.
	{
		index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
		if (index < -1 || (mode != 'M' && index < 0)) // Allow -1 to mean "all rows" when in modify mode.
			return;
		++aParam;  // Remove the first parameter from further consideration to make Insert/Modify symmetric with Add.
		--aParamCount;
	}

	if (!g->GuiDefaultWindowValid())
		return;
	GuiType &gui = *g->GuiDefaultWindow; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentListView)
		return;
	GuiControlType &control = *gui.mCurrentListView;

	LPTSTR options = ParamIndexToOptionalString(0, buf);
	bool ensure_visible = false, is_checked = false;  // Checkmark.
	int col_start_index = 0;
	LVITEM lvi;
	lvi.mask = LVIF_STATE; // LVIF_STATE: state member is valid, but only to the extent that corresponding bits are set in stateMask (the rest will be ignored).
	lvi.stateMask = 0;
	lvi.state = 0;

	// Parse list of space-delimited options:
	TCHAR *next_option, *option_end, orig_char;
	bool adding; // Whether this option is being added (+) or removed (-).

	for (next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
	{
		if (*next_option == '-')
		{
			adding = false;
			// omit_leading_whitespace() is not called, which enforces the fact that the option word must
			// immediately follow the +/- sign.  This is done to allow the flexibility to have options
			// omit the plus/minus sign, and also to reserve more flexibility for future option formats.
			++next_option;  // Point it to the option word itself.
		}
		else
		{
			// Assume option is being added in the absence of either sign.  However, when we were
			// called by GuiControl(), the first option in the list must begin with +/- otherwise the cmd
			// would never have been properly detected as GUICONTROL_CMD_OPTIONS in the first place.
			adding = true;
			if (*next_option == '+')
				++next_option;  // Point it to the option word itself.
			//else do not increment, under the assumption that the plus has been omitted from a valid
			// option word and is thus an implicit plus.
		}

		if (!*next_option) // In case the entire option string ends in a naked + or -.
			break;
		// Find the end of this option item:
		if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
			option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.
		if (option_end == next_option)
			continue; // i.e. the string contains a + or - with a space or tab after it, which is intentionally ignored.

		// Temporarily terminate to help eliminate ambiguity for words contained inside other words,
		// such as "Checked" inside of "CheckedGray":
		orig_char = *option_end;
		*option_end = '\0';

		if (!_tcsnicmp(next_option, _T("Select"), 6)) // Could further allow "ed" suffix by checking for that inside, but "Selected" is getting long so it doesn't seem something many would want to use.
		{
			next_option += 6;
			// If it's Select0, invert the mode to become "no select". This allows a boolean variable
			// to be more easily applied, such as this expression: "Select" . VarContainingState
			if (*next_option && !ATOI(next_option))
				adding = !adding;
			// Another reason for not having "Select" imply "Focus" by default is that it would probably
			// reduce performance when selecting all or a large number of rows.
			// Because a row might or might not have focus, the script may wish to retain its current
			// focused state.  For this reason, "select" does not imply "focus", which allows the
			// LVIS_FOCUSED bit to be omitted from the stateMask, which in turn retains the current
			// focus-state of the row rather than disrupting it.
			lvi.stateMask |= LVIS_SELECTED;
			if (adding)
				lvi.state |= LVIS_SELECTED;
			//else removing, so the presence of LVIS_SELECTED in the stateMask above will cause it to be de-selected.
		}
		else if (!_tcsnicmp(next_option, _T("Focus"), 5))
		{
			next_option += 5;
			if (*next_option && !ATOI(next_option)) // If it's Focus0, invert the mode to become "no focus".
				adding = !adding;
			lvi.stateMask |= LVIS_FOCUSED;
			if (adding)
				lvi.state |= LVIS_FOCUSED;
			//else removing, so the presence of LVIS_FOCUSED in the stateMask above will cause it to be de-focused.
		}
		else if (!_tcsnicmp(next_option, _T("Check"), 5))
		{
			// The rationale for not checking for an optional "ed" suffix here and incrementing next_option by 2
			// is that: 1) It would be inconsistent with the lack of support for "selected" (see reason above);
			// 2) Checkboxes in a ListView are fairly rarely used, so code size reduction might be more important.
			next_option += 5;
			if (*next_option && !ATOI(next_option)) // If it's Check0, invert the mode to become "unchecked".
				adding = !adding;
			if (mode == 'M') // v1.0.46.10: Do this section only for Modify, not Add/Insert, to avoid generating an extra "unchecked" notification when a row is added/inserted with an initial state of "checked".  In other words, the script now receives only a "checked" notification, not an "unchecked+checked". Search on is_checked for more comments.
			{
				lvi.stateMask |= LVIS_STATEIMAGEMASK;
				lvi.state |= adding ? 0x2000 : 0x1000; // The #1 image is "unchecked" and the #2 is "checked".
			}
			is_checked = adding;
		}
		else if (!_tcsnicmp(next_option, _T("Col"), 3))
		{
			if (adding)
			{
				col_start_index = ATOI(next_option + 3) - 1; // The ability to start at a column other than 1 (i.e. subitem vs. item).
				if (col_start_index < 0)
					col_start_index = 0;
			}
		}
		else if (!_tcsnicmp(next_option, _T("Icon"), 4))
		{
			// Testing shows that there is no way to avoid having an item icon in report view if the
			// ListView has an associated small-icon ImageList (well, perhaps you could have it show
			// a blank square by specifying an invalid icon index, but that doesn't seem useful).
			// If LVIF_IMAGE is entirely omitted when adding and item/row, the item will take on the
			// first icon in the list.  This is probably by design because the control wants to make
			// each item look consistent by indenting its first field by a certain amount for the icon.
			if (adding)
			{
				lvi.mask |= LVIF_IMAGE;
				lvi.iImage = ATOI(next_option + 4) - 1;  // -1 to convert to zero-based.
			}
			//else removal of icon currently not supported (see comment above), so do nothing in order
			// to reserve "-Icon" in case a future way can be found to do it.
		}
		else if (!_tcsicmp(next_option, _T("Vis"))) // v1.0.44
			// Since this option much more typically used with LV_Modify than LV_Add/Insert, the technique of
			// Vis%VarContainingOneOrZero% isn't supported, to reduce code size.
			ensure_visible = adding; // Ignored by modes other than LV_Modify(), since it's not really appropriate when adding a row (plus would add code complexity).

		// If the item was not handled by the above, ignore it because it is unknown.
		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
	}

	// More maintainable and performs better to have a separate struct for subitems vs. items.
	LVITEM lvi_sub;
	// Ensure mask is pure to avoid giving it any excuse to fail due to the fact that
	// "You cannot set the state or lParam members for subitems."
	lvi_sub.mask = LVIF_TEXT;

	int i, j, rows_to_change;
	if (index == -1) // Modify all rows (above has ensured that this is only happens in modify-mode).
	{
		rows_to_change = ListView_GetItemCount(control.hwnd);
		lvi.iItem = 0;
		ensure_visible = false; // Not applicable when operating on all rows.
	}
	else // Modify or insert a single row.  Set it up for the loop to perform exactly one iteration.
	{
		rows_to_change = 1;
		lvi.iItem = index; // Which row to operate upon.  This can be a huge number such as 999999 if the caller wanted to append vs. insert.
	}
	lvi.iSubItem = 0;  // Always zero to operate upon the item vs. sub-item (subitems have their own LVITEM struct).
	aResultToken.value_int64 = 1; // Set default from this point forward to be true/success. It will be overridden in insert mode to be the index of the new row.

	for (j = 0; j < rows_to_change; ++j, ++lvi.iItem) // ++lvi.iItem because if the loop has more than one iteration, by definition it is modifying all rows starting at 0.
	{
		if (!ParamIndexIsOmitted(1) && col_start_index == 0) // 2nd parameter: item's text (first field) is present, so include that when setting the item.
		{
			lvi.pszText = ParamIndexToString(1, buf); // Fairly low-overhead, so called every iteration for simplicity (so that buf can be used for both items and subitems).
			lvi.mask |= LVIF_TEXT;
		}
		if (mode == 'I') // Insert or Add.
		{
			// Note that ListView_InsertItem() will append vs. insert if the index is too large, in which case
			// it returns the items new index (which will be the last item in the list unless the control has
			// auto-sort style).
			// Below uses +1 to convert from zero-based to 1-based.  This also converts a failure result of -1 to 0.
			if (   !(aResultToken.value_int64 = ListView_InsertItem(control.hwnd, &lvi) + 1)   )
				return; // Since item can't be inserted, no reason to try attaching any subitems to it.
			// Update iItem with the actual index assigned to the item, which might be different than the
			// specified index if the control has an auto-sort style in effect.  This new iItem value
			// is used for ListView_SetCheckState() and for the attaching of any subitems to this item.
			lvi_sub.iItem = (int)aResultToken.value_int64 - 1;  // -1 to convert back to zero-based.
			// For add/insert (but not modify), testing shows that checkmark must be added only after
			// the item has been inserted rather than provided in the lvi.state/stateMask fields.
			// MSDN confirms this by saying "When an item is added with [LVS_EX_CHECKBOXES],
			// it will always be set to the unchecked state [ignoring any value placed in bits
			// 12 through 15 of the state member]."
			if (is_checked)
				ListView_SetCheckState(control.hwnd, lvi_sub.iItem, TRUE); // TRUE = Check the row's checkbox.
				// Note that 95/NT4 systems that lack comctl32.dll 4.70+ distributed with MSIE 3.x
				// do not support LVS_EX_CHECKBOXES, so the above will have no effect for them.
		}
		else // Modify.
		{
			// Rather than trying to detect if anything was actually changed, this is called
			// unconditionally to simplify the code (ListView_SetItem() is probably very fast if it
			// discovers that lvi.mask==LVIF_STATE and lvi.stateMask==0).
			// By design (to help catch script bugs), a failure here does not revert to append mode.
			if (!ListView_SetItem(control.hwnd, &lvi)) // Returns TRUE/FALSE.
				aResultToken.value_int64 = 0; // Indicate partial failure, but attempt to continue in case it failed for reason other than "row doesn't exist".
			lvi_sub.iItem = lvi.iItem; // In preparation for modifying any subitems that need it.
			if (ensure_visible) // Seems best to do this one prior to "select" below.
				SendMessage(control.hwnd, LVM_ENSUREVISIBLE, lvi.iItem, FALSE); // PartialOK==FALSE is somewhat arbitrary.
		}

		// For each remaining parameter, assign its text to a subitem.
		// Testing shows that if the control has too few columns for all of the fields/parameters
		// present, the ones at the end are automatically ignored: they do not consume memory nor
		// do they significantly impact performance (at least on Windows XP).  For this reason, there
		// is no code above the for-loop above to reduce aParamCount if it's "too large" because
		// it might reduce flexibility (in case future/past OSes allow non-existent columns to be
		// populated, or in case current OSes allow the contents of recently removed columns to be modified).
		for (lvi_sub.iSubItem = (col_start_index > 1) ? col_start_index : 1 // Start at the first subitem unless we were told to start at or after the third column.
			// "i" starts at 2 (the third parameter) unless col_start_index is greater than 0, in which case
			// it starts at 1 (the second parameter) because that parameter has not yet been assigned to anything:
			, i = 2 - (col_start_index > 0)
			; i < aParamCount
			; ++i, ++lvi_sub.iSubItem)
		{
			if (aParam[i]->symbol == SYM_MISSING) // Omitted, such as LV_Modify(1,Opt,"One",,"Three").
				continue;
			lvi_sub.pszText = ParamIndexToString(i, buf); // Done every time through the outer loop since it's not high-overhead, and for code simplicity.
			if (!ListView_SetItem(control.hwnd, &lvi_sub) && mode != 'I') // Relies on short-circuit. Seems best to avoid loss of item's index in insert mode, since failure here should be rare.
				aResultToken.value_int64 = 0; // Indicate partial failure, but attempt to continue in case it failed for reason other than "row doesn't exist".
		}
	} // outer for()

	// When the control has no rows, work around the fact that LVM_SETITEMCOUNT delivers less than 20%
	// of its full benefit unless done after the first row is added (at least on XP SP1).  A non-zero
	// row_count_hint tells us that this message should be sent after the row has been inserted/appended:
	if (control.union_lv_attrib->row_count_hint > 0 && mode == 'I')
	{
		SendMessage(control.hwnd, LVM_SETITEMCOUNT, control.union_lv_attrib->row_count_hint, 0); // Last parameter should be 0 for LVS_OWNERDATA (verified if you look at the definition of ListView_SetItemCount macro).
		control.union_lv_attrib->row_count_hint = 0; // Reset so that it only gets set once per request.
	}
}



BIF_DECL(BIF_LV_Delete)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: Row index (one-based when it comes in).
{
	aResultToken.value_int64 = 0; // Set default return value.
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no ListView in window).
	// And others as shown below.

	if (!g->GuiDefaultWindowValid())
		return;
	GuiType &gui = *g->GuiDefaultWindow; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentListView)
		return;
	HWND control_hwnd = gui.mCurrentListView->hwnd;

	if (ParamIndexIsOmitted(0))
	{
		aResultToken.value_int64 = SendMessage(control_hwnd, LVM_DELETEALLITEMS, 0, 0); // Returns TRUE/FALSE.
		return;
	}

	// Since above didn't return, there is a first parameter present.
	int index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
	if (index > -1)
		aResultToken.value_int64 = SendMessage(control_hwnd, LVM_DELETEITEM, index, 0); // Returns TRUE/FALSE.
	//else even if index==0, for safety, it seems not to do a delete-all.
}



BIF_DECL(BIF_LV_InsertModifyDeleteCol)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: Column index (one-based when it comes in).
// 2: String of options
// 3: New text of column
// There are also some special modes when only zero or one parameter is present, see below.
{
	TCHAR mode = ctoupper(aResultToken.marker[3]); // Union's marker initially contains the function name. LV_[I]nsertCol.
	LPTSTR buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).
	aResultToken.value_int64 = 0; // Set default return value. Must be done only after consulting marker above.
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no ListView in window).
	// Column not found in ListView.

	if (!g->GuiDefaultWindowValid())
		return;
	GuiType &gui = *g->GuiDefaultWindow; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentListView)
		return;
	GuiControlType &control = *gui.mCurrentListView;
	lv_attrib_type &lv_attrib = *control.union_lv_attrib;
	DWORD view_mode = mode != 'D' ? GuiType::ControlGetListViewMode(control.hwnd) : 0;

	int index;
	if (!ParamIndexIsOmitted(0))
		index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
	else // Zero parameters.  Load-time validation has ensured that the 'D' (delete) mode cannot have zero params.
	{
		if (mode == 'M')
		{
			if (view_mode != LVS_REPORT)
				return; // And leave aResultToken.value_int64 at 0 to indicate failure.
			// Otherwise:
			aResultToken.value_int64 = 1; // Always successful (for consistency), regardless of what happens below.
			// v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
			// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items):
			for (int i = 0; ; ++i) // Don't limit it to lv_attrib.col_count in case script added extra columns via direct API calls.
				if (!ListView_SetColumnWidth(control.hwnd, i, LVSCW_AUTOSIZE)) // Failure means last column has already been processed.
					break; // Break vs. return in case the loop has zero iterations due to zero columns (not currently possible, but helps maintainability).
			return;
		}
		// Since above didn't return, mode must be 'I' (insert).
		index = lv_attrib.col_count; // When no insertion index was specified, append to the end of the list.
	}

	// Do this prior to checking if index is in bounds so that it can support columns beyond LV_MAX_COLUMNS:
	if (mode == 'D') // Delete a column.  In this mode, index parameter was made mandatory via load-time validation.
	{
		if (aResultToken.value_int64 = ListView_DeleteColumn(control.hwnd, index))  // Returns TRUE/FALSE.
		{
			// It's important to note that when the user slides columns around via drag and drop, the
			// column index as seen by the script is not changed.  This is fortunate because otherwise,
			// the lv_attrib.col array would get out of sync with the column indices.  Testing shows that
			// all of the following operations respect the original column index, regardless of where the
			// user may have moved the column physically: InsertCol, DeleteCol, ModifyCol.  Insert and Delete
			// shifts the indices of those columns that *originally* lay to the right of the affected column.
			if (lv_attrib.col_count > 0) // Avoid going negative, which would otherwise happen if script previously added columns by calling the API directly.
				--lv_attrib.col_count; // Must be done prior to the below.
			if (index < lv_attrib.col_count) // When a column other than the last was removed, adjust the array so that it stays in sync with actual columns.
				MoveMemory(lv_attrib.col+index, lv_attrib.col+index+1, sizeof(lv_col_type)*(lv_attrib.col_count-index));
		}
		return;
	}
	// Do this prior to checking if index is in bounds so that it can support columns beyond LV_MAX_COLUMNS:
	if (mode == 'M' && aParamCount < 2) // A single parameter is a special modify-mode to auto-size that column.
	{
		// v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
		// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items):
		if (view_mode == LVS_REPORT)
			aResultToken.value_int64 = ListView_SetColumnWidth(control.hwnd, index, LVSCW_AUTOSIZE);
		//else leave aResultToken.value_int64 set to 0.
		return;
	}
	if (mode == 'I')
	{
		if (lv_attrib.col_count >= LV_MAX_COLUMNS) // No room to insert or append.
			return;
		if (index >= lv_attrib.col_count) // For convenience, fall back to "append" when index too large.
			index = lv_attrib.col_count;
	}
	//else do nothing so that modification and deletion of columns that were added via script's
	// direct calls to the API can sort-of work (it's documented in the help file that it's not supported,
	// since col-attrib array can get out of sync with actual columns that way).

	if (index < 0 || index >= LV_MAX_COLUMNS) // For simplicity, do nothing else if index out of bounds.
		return; // Avoid array under/overflow below.

	// In addition to other reasons, must convert any numeric value to a string so that an isolated width is
	// recognized, e.g. LV_SetCol(1, old_width + 10):
	LPTSTR options = ParamIndexToOptionalString(1, buf);

	// It's done the following way so that when in insert-mode, if the column fails to be inserted, don't
	// have to remove the inserted array element from the lv_attrib.col array:
	lv_col_type temp_col = {0}; // Init unconditionally even though only needed for mode=='I'.
	lv_col_type &col = (mode == 'I') ? temp_col : lv_attrib.col[index]; // Done only after index has been confirmed in-bounds.

	LVCOLUMN lvc;
	lvc.mask = LVCF_FMT;
	if (mode == 'M') // Fetch the current format so that it's possible to leave parts of it unaltered.
		ListView_GetColumn(control.hwnd, index, &lvc);
	else // Mode is "insert".
		lvc.fmt = 0;

	// Init defaults prior to parsing options:
	bool sort_now = false;
	int do_auto_size = (mode == 'I') ? LVSCW_AUTOSIZE_USEHEADER : 0;  // Default to auto-size for new columns.
	TCHAR sort_now_direction = 'A'; // Ascending.
	int new_justify = lvc.fmt & LVCFMT_JUSTIFYMASK; // Simplifies the handling of the justification bitfield.
	//lvc.iSubItem = 0; // Not necessary if the LVCF_SUBITEM mask-bit is absent.

	// Parse list of space-delimited options:
	TCHAR *next_option, *option_end, orig_char;
	bool adding; // Whether this option is being added (+) or removed (-).

	for (next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
	{
		if (*next_option == '-')
		{
			adding = false;
			// omit_leading_whitespace() is not called, which enforces the fact that the option word must
			// immediately follow the +/- sign.  This is done to allow the flexibility to have options
			// omit the plus/minus sign, and also to reserve more flexibility for future option formats.
			++next_option;  // Point it to the option word itself.
		}
		else
		{
			// Assume option is being added in the absence of either sign.  However, when we were
			// called by GuiControl(), the first option in the list must begin with +/- otherwise the cmd
			// would never have been properly detected as GUICONTROL_CMD_OPTIONS in the first place.
			adding = true;
			if (*next_option == '+')
				++next_option;  // Point it to the option word itself.
			//else do not increment, under the assumption that the plus has been omitted from a valid
			// option word and is thus an implicit plus.
		}

		if (!*next_option) // In case the entire option string ends in a naked + or -.
			break;
		// Find the end of this option item:
		if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
			option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.
		if (option_end == next_option)
			continue; // i.e. the string contains a + or - with a space or tab after it, which is intentionally ignored.

		// Temporarily terminate to help eliminate ambiguity for words contained inside other words,
		// such as "Checked" inside of "CheckedGray":
		orig_char = *option_end;
		*option_end = '\0';

		// For simplicity, the value of "adding" is ignored for this and the other number/alignment options.
		if (!_tcsicmp(next_option, _T("Integer")))
		{
			// For simplicity, changing the col.type dynamically (since it's so rarely needed)
			// does not try to set up col.is_now_sorted_ascending so that the next click on the column
			// puts it into default starting order (which is ascending unless the Desc flag was originally
			// present).
			col.type = LV_COL_INTEGER;
			new_justify = LVCFMT_RIGHT;
		}
		else if (!_tcsicmp(next_option, _T("Float")))
		{
			col.type = LV_COL_FLOAT;
			new_justify = LVCFMT_RIGHT;
		}
		else if (!_tcsicmp(next_option, _T("Text"))) // Seems more approp. name than "Str" or "String"
			// Since "Text" is so general, it seems to leave existing alignment (Center/Right) as it is.
			col.type = LV_COL_TEXT;

		// The following can exist by themselves or in conjunction with the above.  They can also occur
		// *after* one of the above words so that alignment can be used to override the default for the type;
		// e.g. "Integer Left" to have left-aligned integers.
		else if (!_tcsicmp(next_option, _T("Right")))
			new_justify = adding ? LVCFMT_RIGHT : LVCFMT_LEFT;
		else if (!_tcsicmp(next_option, _T("Center")))
			new_justify = adding ? LVCFMT_CENTER : LVCFMT_LEFT;
		else if (!_tcsicmp(next_option, _T("Left"))) // Supported so that existing right/center column can be changed back to left.
			new_justify = LVCFMT_LEFT; // The value of "adding" seems inconsequential so is ignored.

		else if (!_tcsicmp(next_option, _T("Uni"))) // Unidirectional sort (clicking the column will not invert to the opposite direction).
			col.unidirectional = adding;
		else if (!_tcsicmp(next_option, _T("Desc"))) // Make descending order the default order (applies to uni and first click of col for non-uni).
			col.prefer_descending = adding; // So that the next click will toggle to the opposite direction.
		else if (!_tcsnicmp(next_option, _T("Case"), 4))
		{
			if (adding)
				col.case_sensitive = !_tcsicmp(next_option + 4, _T("Locale")) ? SCS_INSENSITIVE_LOCALE : SCS_SENSITIVE;
			else
				col.case_sensitive = SCS_INSENSITIVE;
		}
		else if (!_tcsicmp(next_option, _T("Logical"))) // v1.0.44.12: Supports StrCmpLogicalW() method of sorting.
			col.case_sensitive = SCS_INSENSITIVE_LOGICAL;

		else if (!_tcsnicmp(next_option, _T("Sort"), 4)) // This is done as an option vs. LV_SortCol/LV_Sort so that the column's options can be changed simultaneously with a "sort now" to refresh.
		{
			// Defer the sort until after all options have been parsed and applied.
			sort_now = true;
			if (!_tcsicmp(next_option + 4, _T("Desc")))
				sort_now_direction = 'D'; // Descending.
		}
		else if (!_tcsicmp(next_option, _T("NoSort"))) // Called "NoSort" so that there's a way to enable and disable the setting via +/-.
			col.sort_disabled = adding;

		else if (!_tcsnicmp(next_option, _T("Auto"), 4)) // No separate failure result is reported for this item.
			// In case the mode is "insert", defer auto-width of column until col exists.
			do_auto_size = _tcsicmp(next_option + 4, _T("Hdr")) ? LVSCW_AUTOSIZE : LVSCW_AUTOSIZE_USEHEADER;

		else if (!_tcsnicmp(next_option, _T("Icon"), 4))
		{
			next_option += 4;
			if (!_tcsicmp(next_option, _T("Right")))
			{
				if (adding)
					lvc.fmt |= LVCFMT_BITMAP_ON_RIGHT;
				else
					lvc.fmt &= ~LVCFMT_BITMAP_ON_RIGHT;
			}
			else // Assume its an icon number or the removal of the icon via -Icon.
			{
				if (adding)
				{
					lvc.mask |= LVCF_IMAGE;
					lvc.fmt |= LVCFMT_IMAGE; // Flag this column as displaying an image.
					lvc.iImage = ATOI(next_option) - 1; // -1 to convert to zero based.
				}
				else
					lvc.fmt &= ~LVCFMT_IMAGE; // Flag this column as NOT displaying an image.
			}
		}

		else // Handle things that are more general than the above, such as single letter options and pure numbers.
		{
			// Width does not have a W prefix to permit a naked expression to be used as the entirely of
			// options.  For example: LV_SetCol(1, old_width + 10)
			// v1.0.37: Fixed to allow floating point (although ATOI below will convert it to integer).
			if (IsPureNumeric(next_option, true, false, true)) // Above has already verified that *next_option can't be whitespace.
			{
				lvc.mask |= LVCF_WIDTH;
				int width = gui.Scale(ATOI(next_option));
				// Specifying a width when the column is initially added prevents the scrollbar from
				// updating on Windows 7 and 10 (but not XP).  As a workaround, initialise the width
				// to 0 and then resize it afterward.  do_auto_size is overloaded for this purpose
				// since it's already passed to ListView_SetColumnWidth().
				if (mode == 'I' && view_mode == LVS_REPORT)
				{
					lvc.cx = 0; // Must be zero; if width is zero, ListView_SetColumnWidth() won't be called.
					do_auto_size = width; // If non-zero, this is passed to ListView_SetColumnWidth().
				}
				else
				{
					lvc.cx = width;
					do_auto_size = 0; // Turn off any auto-sizing that may have been put into effect (explicitly or by default).
				}
			}
		}

		// If the item was not handled by the above, ignore it because it is unknown.
		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
	}

	// Apply any changed justification/alignment to the fmt bit field:
	lvc.fmt = (lvc.fmt & ~LVCFMT_JUSTIFYMASK) | new_justify;

	if (!ParamIndexIsOmitted(2)) // Parameter #3 (text) is present.
	{
		lvc.pszText = ParamIndexToString(2, buf);
		lvc.mask |= LVCF_TEXT;
	}

	if (mode == 'M') // Modify vs. Insert (Delete was already returned from, higher above).
		// For code simplicity, this is called unconditionally even if nothing internal the control's column
		// needs updating.  This seems justified given how rarely columns are modified.
		aResultToken.value_int64 = ListView_SetColumn(control.hwnd, index, &lvc); // Returns TRUE/FALSE.
	else // Insert
	{
		// It's important to note that when the user slides columns around via drag and drop, the
		// column index as seen by the script is not changed.  This is fortunate because otherwise,
		// the lv_attrib.col array would get out of sync with the column indices.  Testing shows that
		// all of the following operations respect the original column index, regardless of where the
		// user may have moved the column physically: InsertCol, DeleteCol, ModifyCol.  Insert and Delete
		// shifts the indices of those columns that *originally* lay to the right of the affected column.
		// Doesn't seem to do anything -- not even with respect to inserting a new first column with it's
		// unusual behavior of inheriting the previously column's contents -- so it's disabled for now.
		// Testing shows that it also does not seem to cause a new column to inherit the indicated subitem's
		// text, even when iSubItem is set to index + 1 vs. index:
		//lvc.mask |= LVCF_SUBITEM;
		//lvc.iSubItem = index;
		// Testing shows that the following serve to set the column's physical/display position in the
		// heading to iOrder without affecting the specified index.  This concept is very similar to
		// when the user drags and drops a column heading to a new position: it's index doesn't change,
		// only it's displayed position:
		//lvc.mask |= LVCF_ORDER;
		//lvc.iOrder = index + 1;
		if (   !(aResultToken.value_int64 = ListView_InsertColumn(control.hwnd, index, &lvc) + 1)   ) // +1 to convert the new index to 1-based.
			return; // Since column could not be inserted, return so that below, sort-now, etc. are not done.
		index = (int)aResultToken.value_int64 - 1; // Update in case some other index was assigned. -1 to convert back to zero-based.
		if (index < lv_attrib.col_count) // Since col is not being appended to the end, make room in the array to insert this column.
			MoveMemory(lv_attrib.col+index+1, lv_attrib.col+index, sizeof(lv_col_type)*(lv_attrib.col_count-index));
			// Above: Shift columns to the right by one.
		lv_attrib.col[index] = col; // Copy temp struct's members to the correct element in the array.
		// The above is done even when index==0 because "col" may contain attributes set via the Options
		// parameter.  Therefore, for code simplicity and rarity of real-world need, no attempt is made
		// to make the following idea work:
		// When index==0, retain the existing attributes due to the unique behavior of inserting a new first
		// column: The new first column inherit's the old column's values (fields), so it seems best to also have it
		// inherit the old column's attributes.
		++lv_attrib.col_count; // New column successfully added.  Must be done only after the MoveMemory() above.
	}

	// Auto-size is done only at this late a stage, in case column was just created above.
	// Note that ListView_SetColumn() apparently does not support LVSCW_AUTOSIZE_USEHEADER for it's "cx" member.
	// do_auto_size contains the actual column width if mode == 'I' and a width was passed by the caller.
	if (do_auto_size && view_mode == LVS_REPORT)
		ListView_SetColumnWidth(control.hwnd, index, do_auto_size); // aResultToken.value_int64 was previous set to the more important result above.
	//else v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
	// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items).

	if (sort_now)
		GuiType::LV_Sort(control, index, false, sort_now_direction);
}



BIF_DECL(BIF_LV_SetImageList)
// Returns (MSDN): "handle to the image list previously associated with the control if successful; NULL otherwise."
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
// 2: Optional: Type of list.
{
	aResultToken.value_int64 = 0; // Set default return value.
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no ListView in window).
	// Column not found in ListView.

	if (!g->GuiDefaultWindowValid())
		return;
	GuiType &gui = *g->GuiDefaultWindow; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentListView)
		return;
	// Caller has ensured that there is at least one incoming parameter:
	HIMAGELIST himl = (HIMAGELIST)ParamIndexToInt64(0);
	int list_type;
	if (!ParamIndexIsOmitted(1))
		list_type = ParamIndexToInt(1);
	else // Auto-detect large vs. small icons based on the actual icon size in the image list.
	{
		int cx, cy;
		ImageList_GetIconSize(himl, &cx, &cy);
		list_type = (cx > GetSystemMetrics(SM_CXSMICON)) ? LVSIL_NORMAL : LVSIL_SMALL;
	}
	aResultToken.value_int64 = (__int64)ListView_SetImageList(gui.mCurrentListView->hwnd, himl, list_type);
}



BIF_DECL(BIF_TV_AddModifyDelete)
// TV_Add():
// Returns the HTREEITEM of the item on success, zero on failure.
// Parameters:
//    1: Text/name of item.
//    2: Parent of item.
//    3: Options.
// TV_Modify():
// Returns the HTREEITEM of the item on success (to allow nested calls in script, zero on failure or partial failure.
// Parameters:
//    1: ID of item to modify.
//    2: Options.
//    3: New name.
// Parameters for TV_Delete():
//    1: ID of item to delete (if omitted, all items are deleted).
{
	TCHAR mode = ctoupper(aResultToken.marker[3]); // Union's marker initially contains the function name. e.g. TV_[A]dd.
	LPTSTR buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).
	aResultToken.value_int64 = 0; // Set default return value. Must be done only after consulting marker above.
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no TreeView in window).
	// And others as shown below.

	if (!g->GuiDefaultWindowValid())
		return;
	GuiType &gui = *g->GuiDefaultWindow; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentTreeView)
		return;
	GuiControlType &control = *gui.mCurrentTreeView;

	if (mode == 'D') // TV_Delete
	{
		// If param #1 is present but is zero, for safety it seems best not to do a delete-all (in case a
		// script bug is so rare that it is never caught until the script is distributed).  Another reason
		// is that a script might do something like TV_Delete(TV_GetSelection()), which would be desired
		// to fail not delete-all if there's ever any way for there to be no selection.
		aResultToken.value_int64 = SendMessage(control.hwnd, TVM_DELETEITEM, 0
			, ParamIndexIsOmitted(0) ? NULL : (LPARAM)ParamIndexToInt64(0));
		return;
	}

	// Since above didn't return, this is TV_Add() or TV_Modify().
	TVINSERTSTRUCT tvi; // It contains a TVITEMEX, which is okay even if MSIE pre-4.0 on Win95/NT because those OSes will simply never access the new/bottommost item in the struct.
	bool add_mode = (mode == 'A'); // For readability & maint.

	LPTSTR options;
	if (add_mode) // TV_Add()
	{
		tvi.hParent = ParamIndexIsOmitted(1) ? NULL : (HTREEITEM)ParamIndexToInt64(1);
		tvi.hInsertAfter = TVI_LAST; // i.e. default is to insert the new item underneath the bottommost sibling.
		options = ParamIndexToOptionalString(2, buf);
	}
	else // TV_Modify()
	{
		// NOTE: Must allow hitem==0 for TV_Modify, at least for the Sort option, because otherwise there would
		// be no way to sort the root-level items.
		tvi.item.hItem = (HTREEITEM)ParamIndexToInt64(0); // Load-time validation has ensured there is a first parameter for TV_Modify().
		// For modify-mode, set default return value to be "success" from this point forward.  Note that
		// in the case of sorting the root-level items, this will set it to zero, but since that almost
		// always succeeds and the script rarely cares whether it succeeds or not, adding code size for that
		// doesn't seem worth it:
		aResultToken.value_int64 = (size_t)tvi.item.hItem;
		if (aParamCount < 2) // In one-parameter mode, simply select the item.
		{
			if (!TreeView_SelectItem(control.hwnd, tvi.item.hItem))
				aResultToken.value_int64 = 0; // Override the HTREEITEM default value set above.
			return;
		}
		// Otherwise, there's a second parameter (even if it's 0 or "").
		options = ParamIndexToString(1, buf);
	}

	// Set defaults prior to options-parsing, to cover all omitted defaults:
	tvi.item.mask = TVIF_STATE; // TVIF_STATE: The state and stateMask members are valid (all other members are ignored).
	tvi.item.stateMask = 0; // All bits in "state" below are ignored unless the corresponding bit is present here in the mask.
	tvi.item.state = 0;
	// It seems tvi.item.cChildren is typically maintained by the control, though one exception is I_CHILDRENCALLBACK
	// and TVN_GETDISPINFO as mentioned at MSDN.

	DWORD select_flag = 0;
	bool ensure_visible = false, ensure_visible_first = false;

	// Parse list of space-delimited options:
	TCHAR *next_option, *option_end, orig_char;
	bool adding; // Whether this option is being added (+) or removed (-).

	for (next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
	{
		if (*next_option == '-')
		{
			adding = false;
			// omit_leading_whitespace() is not called, which enforces the fact that the option word must
			// immediately follow the +/- sign.  This is done to allow the flexibility to have options
			// omit the plus/minus sign, and also to reserve more flexibility for future option formats.
			++next_option;  // Point it to the option word itself.
		}
		else
		{
			// Assume option is being added in the absence of either sign.  However, when we were
			// called by GuiControl(), the first option in the list must begin with +/- otherwise the cmd
			// would never have been properly detected as GUICONTROL_CMD_OPTIONS in the first place.
			adding = true;
			if (*next_option == '+')
				++next_option;  // Point it to the option word itself.
			//else do not increment, under the assumption that the plus has been omitted from a valid
			// option word and is thus an implicit plus.
		}

		if (!*next_option) // In case the entire option string ends in a naked + or -.
			break;
		// Find the end of this option item:
		if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
			option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.
		if (option_end == next_option)
			continue; // i.e. the string contains a + or - with a space or tab after it, which is intentionally ignored.

		// Temporarily terminate to help eliminate ambiguity for words contained inside other words,
		// such as "Checked" inside of "CheckedGray":
		orig_char = *option_end;
		*option_end = '\0';

		if (!_tcsicmp(next_option, _T("Select"))) // Could further allow "ed" suffix by checking for that inside, but "Selected" is getting long so it doesn't seem something many would want to use.
		{
			// Selection of an item apparently needs to be done via message for the control to update itself
			// properly.  Otherwise, single-select isn't enforced via de-selecting previous item and the newly
			// selected item isn't revealed/shown.  There may be other side-effects.
			if (adding)
				select_flag = TVGN_CARET;
			//else since "de-select" is not a supported action, no need to support "-Select".
			// Furthermore, since a TreeView is by its nature has only one item selected at a time, it seems
			// unnecessary to support Select%VarContainingOneOrZero%.  This is because it seems easier for a
			// script to simply load the Tree then select the desired item afterward.
		}
		else if (!_tcsnicmp(next_option, _T("Vis"), 3))
		{
			// Since this option much more typically used with TV_Modify than TV_Add, the technique of
			// Vis%VarContainingOneOrZero% isn't supported, to reduce code size.
			next_option += 3;
			if (!_tcsicmp(next_option, _T("First"))) // VisFirst
				ensure_visible_first = adding;
			else if (!*next_option)
				ensure_visible = adding;
		}
		else if (!_tcsnicmp(next_option, _T("Bold"), 4))
		{
			next_option += 4;
			if (*next_option && !ATOI(next_option)) // If it's Bold0, invert the mode.
				adding = !adding;
			tvi.item.stateMask |= TVIS_BOLD;
			if (adding)
				tvi.item.state |= TVIS_BOLD;
			//else removing, so the fact that this TVIS flag has just been added to the stateMask above
			// but is absent from item.state should remove this attribute from the item.
		}
		else if (!_tcsnicmp(next_option, _T("Expand"), 6))
		{
			next_option += 6;
			if (*next_option && !ATOI(next_option)) // If it's Expand0, invert the mode to become "collapse".
				adding = !adding;
			if (add_mode)
			{
				if (adding)
				{
					// Don't expand via msg because it won't work: since the item is being newly added
					// now, by definition it doesn't have any children, and testing shows that sending
					// the expand message has no effect, but setting the state bit does:
					tvi.item.stateMask |= TVIS_EXPANDED;
					tvi.item.state |= TVIS_EXPANDED;
					// Since the script is deliberately expanding the item, it seems best not to send the
					// TVN_ITEMEXPANDING/-ED messages because:
					// 1) Sending TVN_ITEMEXPANDED without first sending a TVN_ITEMEXPANDING message might
					//    decrease maintainability, and possibly even produce unwanted side-effects.
					// 2) Code size and performance (avoids generating extra message traffic).
				}
				//else removing, so nothing needs to be done because "collapsed" is the default state
				// of a TV item upon creation.
			}
			else // TV_Modify(): Expand and collapse both require a message to work properly on an existing item.
				// Strangely, this generates a notification sometimes (such as the first time) but not for subsequent
				// expands/collapses of that same item.  Also, TVE_TOGGLE is not currently supported because it seems
				// like it would be too rarely used.
				if (!TreeView_Expand(control.hwnd, tvi.item.hItem, adding ? TVE_EXPAND : TVE_COLLAPSE))
					aResultToken.value_int64 = 0; // Indicate partial failure by overriding the HTREEITEM return value set earlier.
					// It seems that despite what MSDN says, failure is returned when collapsing and item that is
					// already collapsed, but not when expanding an item that is already expanded.  For performance
					// reasons and rarity of script caring, it seems best not to try to adjust/fix this.
		}
		else if (!_tcsnicmp(next_option, _T("Check"), 5))
		{
			// The rationale for not checking for an optional "ed" suffix here and incrementing next_option by 2
			// is that: 1) It would be inconsistent with the lack of support for "selected" (see reason above);
			// 2) Checkboxes in a ListView are fairly rarely used, so code size reduction might be more important.
			next_option += 5;
			if (*next_option && !ATOI(next_option)) // If it's Check0, invert the mode to become "unchecked".
				adding = !adding;
			//else removing, so the fact that this TVIS flag has just been added to the stateMask above
			// but is absent from item.state should remove this attribute from the item.
			tvi.item.stateMask |= TVIS_STATEIMAGEMASK;  // Unlike ListViews, Tree checkmarks can be applied in the same step as creating a Tree item.
			tvi.item.state |= adding ? 0x2000 : 0x1000; // The #1 image is "unchecked" and the #2 is "checked".
		}
		else if (!_tcsnicmp(next_option, _T("Icon"), 4))
		{
			if (adding)
			{
				// To me, having a different icon for when the item is selected seems rarely used.  After all,
				// its obvious the item is selected because it's highlighted (unless it lacks a name?)  So this
				// policy makes things easier for scripts that don't want to distinguish.  If ever it is needed,
				// new options such as IconSel and IconUnsel can be added.
				tvi.item.mask |= TVIF_IMAGE|TVIF_SELECTEDIMAGE;
				tvi.item.iSelectedImage = tvi.item.iImage = ATOI(next_option + 4) - 1;  // -1 to convert to zero-based.
			}
			//else removal of icon currently not supported (see comment above), so do nothing in order
			// to reserve "-Icon" in case a future way can be found to do it.
		}
		else if (!_tcsicmp(next_option, _T("Sort")))
		{
			if (add_mode)
				tvi.hInsertAfter = TVI_SORT; // For simplicity, the value of "adding" is ignored.
			else
				// Somewhat debatable, but it seems best to report failure via the return value even though
				// failure probably only occurs when the item has no children, and the script probably
				// doesn't often care about such failures.  It does result in the loss of the HTREEITEM return
				// value, but even if that call is nested in another, the zero should produce no effect in most cases.
				if (!TreeView_SortChildren(control.hwnd, tvi.item.hItem, FALSE)) // Best default seems no-recurse, since typically this is used after a user edits merely a single item.
					aResultToken.value_int64 = 0; // Indicate partial failure by overriding the HTREEITEM return value set earlier.
		}
		else if (add_mode) // MUST BE LISTED LAST DUE TO "ELSE IF": Options valid only for TV_Add().
		{
			if (!_tcsicmp(next_option, _T("First")))
				tvi.hInsertAfter = TVI_FIRST; // For simplicity, the value of "adding" is ignored.
			else if (IsPureNumeric(next_option, false, false, false))
				tvi.hInsertAfter = (HTREEITEM)ATOI64(next_option); // ATOI64 vs. ATOU avoids need for extra casting to avoid compiler warning.
		}

		// If the item was not handled by the above, ignore it because it is unknown.
		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
	}

	if (add_mode) // TV_Add()
	{
		tvi.item.pszText = ParamIndexToString(0, buf);
		tvi.item.mask |= TVIF_TEXT;
		tvi.item.hItem = TreeView_InsertItem(control.hwnd, &tvi); // Update tvi.item.hItem for convenience/maint. I'ts for use in later sections because aResultToken.value_int64 is overridden to be zero for partial failure in modify-mode.
		aResultToken.value_int64 = (__int64)tvi.item.hItem; // Set return value.
	}
	else // TV_Modify()
	{
		if (!ParamIndexIsOmitted(2)) // An explicit empty string is allowed, which sets it to a blank value.  By contrast, if the param is omitted, the name is left changed.
		{
			tvi.item.pszText = ParamIndexToString(2, buf); // Reuse buf now that options (above) is done with it.
			tvi.item.mask |= TVIF_TEXT;
		}
		//else name/text parameter has been omitted, so don't change the item's name.
		if (tvi.item.mask != LVIF_STATE || tvi.item.stateMask) // An item's property or one of the state bits needs changing.
			if (!TreeView_SetItem(control.hwnd, &tvi.itemex))
				aResultToken.value_int64 = 0; // Indicate partial failure by overriding the HTREEITEM return value set earlier.
	}

	if (ensure_visible) // Seems best to do this one prior to "select" below.
		SendMessage(control.hwnd, TVM_ENSUREVISIBLE, 0, (LPARAM)tvi.item.hItem); // Return value is ignored in this case, since its definition seems a little weird.
	if (ensure_visible_first) // Seems best to do this one prior to "select" below.
		TreeView_Select(control.hwnd, tvi.item.hItem, TVGN_FIRSTVISIBLE); // Return value is also ignored due to rarity, code size, and because most people wouldn't care about a failure even if for some reason it failed.
	if (select_flag)
		if (!TreeView_Select(control.hwnd, tvi.item.hItem, select_flag) && !add_mode) // Relies on short-circuit boolean order.
			aResultToken.value_int64 = 0; // When not in add-mode, indicate partial failure by overriding the return value set earlier (add-mode should always return the new item's ID).
}



HTREEITEM GetNextTreeItem(HWND aTreeHwnd, HTREEITEM aItem)
// Helper function for others below.
// If aItem is NULL, caller wants topmost ROOT item returned.
// Otherwise, the next child, sibling, or parent's sibling is returned in a manner that allows the caller
// to traverse every item in the tree easily.
{
	if (!aItem)
		return TreeView_GetRoot(aTreeHwnd);
	// Otherwise, do depth-first recursion.  Must be done in the following order to allow full traversal:
	// Children first.
	// Then siblings.
	// Then parent's sibling(s).
	HTREEITEM hitem;
	if (hitem = TreeView_GetChild(aTreeHwnd, aItem))
		return hitem;
	if (hitem = TreeView_GetNextSibling(aTreeHwnd, aItem))
		return hitem;
	// The last stage is trickier than the above: parent's next sibling, or if none, its parent's parent's sibling, etc.
	for (HTREEITEM hparent = aItem;;)
	{
		if (   !(hparent = TreeView_GetParent(aTreeHwnd, hparent))   ) // No parent, so this is a root-level item.
			return NULL; // There is no next item.
		// Now it's known there is a parent.  It's not necessary to check that parent's children because that
		// would have been done by a prior iteration in the script.
		if (hitem = TreeView_GetNextSibling(aTreeHwnd, hparent))
			return hitem;
		// Otherwise, parent has no sibling, but does its parent (and so on)? Continue looping to find out.
	}
}



BIF_DECL(BIF_TV_GetRelatedItem)
// TV_GetParent/Child/Selection/Next/Prev(hitem):
// The above all return the HTREEITEM (or 0 on failure).
// When TV_GetNext's second parameter is present, the search scope expands to include not just siblings,
// but also children and parents, which allows a tree to be traversed from top to bottom without the script
// having to do something fancy.
{
	LPTSTR fn_name = aResultToken.marker; // Save early for maintainability: Union's marker initially contains the function name. TV_Get[S]election.
	LPTSTR buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).
	aResultToken.value_int64 = 0; // Set default return value. Must be done only after saving marker above.
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no TreeView in window).
	// Item not found in TreeView.

	if (!g->GuiDefaultWindowValid())
		return;
	GuiType &gui = *g->GuiDefaultWindow; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentTreeView)
		return;
	HWND control_hwnd = gui.mCurrentTreeView->hwnd;

	HTREEITEM hitem = (HTREEITEM)ParamIndexToOptionalIntPtr(0, NULL);

	if (ParamIndexIsOmitted(1))
	{
		WPARAM flag;
		TCHAR char7 = ctoupper(fn_name[7]);
		switch(ctoupper(fn_name[6]))
		{
		case 'S': flag = TVGN_CARET; break; // TV_GetSelection(). TVGN_CARET is focused item.
		case 'P': flag = (char7 == 'A') ? TVGN_PARENT : TVGN_PREVIOUS; break; // TV_GetParent/Prev.
		case 'N': flag = (aParamCount < 1 || !hitem) ? TVGN_ROOT : TVGN_NEXT; break; // TV_GetNext(no-parameters) yields very first item in Tree (TVGN_ROOT).
		// Above: It seems best to treat hitem==0 as "get root", even though it sacrifices some error detection,
		// because not doing so would be inconsistent with the fact that TV_GetNext(0, "Full") does get the root
		// (which needs to be retained to make script loops to traverse entire tree easier).
		case 'C':
			if (char7 == 'O') // i.e. the "CO" in TV_GetCount().
			{
				// There's a known bug mentioned at MSDN that a TreeView might report a negative count when there
				// are more than 32767 items in it (though of course HTREEITEM values are never negative since they're
				// defined as unsigned pseudo-addresses)).  But apparently, that bug only applies to Visual Basic and/or
				// older OSes, because testing shows that SendMessage(TVM_GETCOUNT) returns 32800+ when there are more
				// than 32767 items in the tree, even without casting to unsigned.  So I'm not sure exactly what the
				// story is with this, so for now just casting to UINT rather than something less future-proof like WORD:
				// Older note, apparently unneeded at least on XP SP2: Cast to WORD to convert -1 through -32768 to the
				// positive counterparts.
				aResultToken.value_int64 = (UINT)SendMessage(control_hwnd, TVM_GETCOUNT, 0, 0);
				return;
			}
			// Since above didn't return, it's TV_GetChild():
			flag = TVGN_CHILD;
			break;
		}
		// Apparently there's no direct call to get the topmost ancestor of an item, presumably because it's rarely
		// needed.  Therefore, no such mode is provide here yet (the syntax TV_GetParent(hitem, true) could be supported
		// if it's ever needed).
		aResultToken.value_int64 = SendMessage(control_hwnd, TVM_GETNEXTITEM, flag, (LPARAM)hitem);
		return;
	}

	// Since above didn't return, this TV_GetNext's 2-parameter mode, which has an expanded scope that includes
	// not just siblings, but also children and parents.  This allows a tree to be traversed from top to bottom
	// without the script having to do something fancy.
	TCHAR first_char_upper = ctoupper(*omit_leading_whitespace(ParamIndexToString(1, buf))); // Resolve parameter #2.
	bool search_checkmark;
	if (first_char_upper == 'C')
		search_checkmark = true;
	else if (first_char_upper == 'F')
		search_checkmark = false;
	else // Reserve other option letters/words for future use by being somewhat strict.
		return; // Retain the default value of 0 set for aResultToken.value_int64 higher above.

	// When an actual item was specified, search begins at the item *after* it.  Otherwise (when NULL):
	// It's a special mode that always considers the root node first.  Otherwise, there would be no way
	// to start the search at the very first item in the tree to find out whether it's checked or not.
	hitem = GetNextTreeItem(control_hwnd, hitem); // Handles the comment above.
	if (!search_checkmark) // Simple tree traversal, so just return the next item (if any).
	{
		aResultToken.value_int64 = (__int64)hitem; // OK if NULL.
		return;
	}

	// Otherwise, search for the next item having a checkmark. For performance, it seems best to assume that
	// the control has the checkbox style (the script would realistically never call it otherwise, so the
	// control's style isn't checked.
	for (; hitem; hitem = GetNextTreeItem(control_hwnd, hitem))
		if (TreeView_GetCheckState(control_hwnd, hitem) == 1) // 0 means unchecked, -1 means "no checkbox image".
		{
			aResultToken.value_int64 = (__int64)hitem;
			return;
		}
	// Since above didn't return, the entire tree starting at the specified item has been searched,
	// with no match found.  Retain the default value of 0 set for aResultToken.value_int64 higher above.
}



BIF_DECL(BIF_TV_Get)
// LV_Get()
// Returns: Varies depending on param #2.
// Parameters:
//    1: HTREEITEM.
//    2: Name of attribute to get.
// LV_GetText()
// Returns: 1 on success and 0 on failure.
// Parameters:
//    1: Output variable (doing it this way allows success/fail return value to more closely mirror the API and
//       simplifies the code since there is currently no easy means of passing back large strings to our caller).
//    2: HTREEITEM.
{
	bool get_text = (ctoupper(aResultToken.marker[6]) == 'T'); // Union's marker initially contains the function name. e.g. TV_Get[T]ext.
	LPTSTR buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).
	aResultToken.value_int64 = 0; // Set default return value. Must be done only after consulting marker above.
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no TreeView in window).
	// Item not found in TreeView.
	// And others.

	if (!g->GuiDefaultWindowValid())
		return;
	GuiType &gui = *g->GuiDefaultWindow; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentTreeView)
		return;
	HWND control_hwnd = gui.mCurrentTreeView->hwnd;

	if (!get_text)
	{
		// Loadtime validation has ensured that param #1 and #2 are present for all these cases.
		HTREEITEM hitem = (HTREEITEM)ParamIndexToInt64(0);
		UINT state_mask;
		switch (ctoupper(*omit_leading_whitespace(ParamIndexToString(1, buf))))
		{
		case 'E': state_mask = TVIS_EXPANDED; break; // Expanded
		case 'C': state_mask = TVIS_STATEIMAGEMASK; break; // Checked
		case 'B': state_mask = TVIS_BOLD; break; // Bold
		//case 'S' for "Selected" is not provided because TV_GetSelection() seems to cover that well enough.
		//case 'P' for "is item a parent?" is not provided because TV_GetChild() seems to cover that well enough.
		// (though it's possible that retrieving TVITEM's cChildren would perform a little better).
		}
		// Below seems to be need a bit-AND with state_mask to work properly, at least on XP SP2.  Otherwise,
		// extra bits are present such as 0x2002 for "expanded" when it's supposed to be either 0x00 or 0x20.
		UINT result = state_mask & (UINT)SendMessage(control_hwnd, TVM_GETITEMSTATE, (WPARAM)hitem, state_mask);
		if (state_mask == TVIS_STATEIMAGEMASK)
		{
			if (result == 0x2000) // It has a checkmark state image.
				aResultToken.value_int64 = (size_t)hitem; // Override 0 set earlier. More useful than returning 1 since it allows function-call nesting.
		}
		else // For all others, anything non-zero means the flag is present.
            if (result)
				aResultToken.value_int64 = (size_t)hitem; // Override 0 set earlier. More useful than returning 1 since it allows function-call nesting.
		return;
	}

	// Since above didn't return, this is LV_GetText().
	// Loadtime validation has ensured that param #1 and #2 are present.
	if (aParam[0]->symbol != SYM_VAR) // No output variable. Supporting a NULL for the purpose of checking for the existence of an item seems too rarely needed.
		return;
	Var &output_var = *aParam[0]->var;

	TCHAR text_buf[LV_TEXT_BUF_SIZE]; // i.e. uses same size as ListView.
	TVITEM tvi;
	tvi.hItem = (HTREEITEM)ParamIndexToInt64(1);
	tvi.mask = TVIF_TEXT;
	tvi.pszText = text_buf;
	tvi.cchTextMax = LV_TEXT_BUF_SIZE - 1; // -1 because of nagging doubt about size vs. length. Some MSDN examples subtract one), such as TabCtrl_GetItem()'s cchTextMax.

	if (SendMessage(control_hwnd, TVM_GETITEM, 0, (LPARAM)&tvi))
	{
		// Must use tvi.pszText vs. text_buf because MSDN says: "Applications should not assume that the text will
		// necessarily be placed in the specified buffer. The control may instead change the pszText member
		// of the structure to point to the new text rather than place it in the buffer."
		output_var.Assign(tvi.pszText);
		aResultToken.value_int64 = (size_t)tvi.hItem; // More useful than returning 1 since it allows function-call nesting.
	}
	else // On failure, it seems best to also clear the output var for better consistency and in case the script doesn't check the return value.
		output_var.Assign();
		// And leave aResultToken.value_int64 set to its default of 0.
}



BIF_DECL(BIF_TV_SetImageList)
// Returns (MSDN): "handle to the image list previously associated with the control if successful; NULL otherwise."
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
// 2: Optional: Type of list.
{
	aResultToken.value_int64 = 0; // Set default return value.
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no TreeView in window).

	if (!g->GuiDefaultWindowValid())
		return;
	GuiType &gui = *g->GuiDefaultWindow; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentTreeView)
		return;
	// Caller has ensured that there is at least one incoming parameter:
	HIMAGELIST himl = (HIMAGELIST)ParamIndexToInt64(0);
	int list_type;
	list_type = ParamIndexToOptionalInt(1, TVSIL_NORMAL);
	aResultToken.value_int64 = (__int64)TreeView_SetImageList(gui.mCurrentTreeView->hwnd, himl, list_type);
}



BIF_DECL(BIF_IL_Create)
// Returns: Handle to the new image list, or 0 on failure.
// Parameters:
// 1: Initial image count (ImageList_Create() ignores values <=0, so no need for error checking).
// 2: Grow count (testing shows it can grow multiple times, even when this is set <=0, so it's apparently only a performance aid)
// 3: Width of each image (overloaded to mean small icon size when omitted or false, large icon size otherwise).
// 4: Future: Height of each image [if this param is present and >0, it would mean param 3 is not being used in its TRUE/FALSE mode)
// 5: Future: Flags/Color depth
{
	// The following old comment makes no sense because large icons are only used if param3 is NON-ZERO,
	// and there was never a distinction between passing zero and omitting the param:
	// So that param3 can be reserved as a future "specified width" param, to go along with "specified height"
	// after it, only when the parameter is both present and numerically zero are large icons used.  Otherwise,
	// small icons are used.
	int param3 = ParamIndexToOptionalInt(2, 0);
	aResultToken.value_int64 = (__int64)ImageList_Create(GetSystemMetrics(param3 ? SM_CXICON : SM_CXSMICON)
		, GetSystemMetrics(param3 ? SM_CYICON : SM_CYSMICON)
		, ILC_MASK | ILC_COLOR32  // ILC_COLOR32 or at least something higher than ILC_COLOR is necessary to support true-color icons.
		, ParamIndexToOptionalInt(0, 2)    // cInitial. 2 seems a better default than one, since it might be common to have only two icons in the list.
		, ParamIndexToOptionalInt(1, 5));  // cGrow.  Somewhat arbitrary default.
}



BIF_DECL(BIF_IL_Destroy)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
{
	// Load-time validation has ensured there is at least one parameter.
	// Returns nonzero if successful, or zero otherwise, so force it to conform to TRUE/FALSE for
	// better consistency with other functions:
	aResultToken.value_int64 = ImageList_Destroy((HIMAGELIST)ParamIndexToInt64(0)) ? 1 : 0;
}



BIF_DECL(BIF_IL_Add)
// Returns: the one-based index of the newly added icon, or zero on failure.
// Parameters:
// 1: HIMAGELIST: Handle of an existing ImageList.
// 2: Filename from which to load the icon or bitmap.
// 3: Icon number within the filename (or mask color for non-icon images).
// 4: The mere presence of this parameter indicates that param #3 is mask RGB-color vs. icon number.
//    This param's value should be "true" to resize the image to fit the image-list's size or false
//    to divide up the image into a series of separate images based on its width.
//    (this parameter could be overloaded to be the filename containing the mask image, or perhaps an HBITMAP
//    provided directly by the script)
// 5: Future: can be the scaling height to go along with an overload of #4 as the width.  However,
//    since all images in an image list are of the same size, the use of this would be limited to
//    only those times when the imagelist would be scaled prior to dividing it into separate images.
// The parameters above (at least #4) can be overloaded in the future calling ImageList_GetImageInfo() to determine
// whether the imagelist has a mask.
{
	LPTSTR buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).
	aResultToken.value_int64 = 0; // Set default in case of early return.
	HIMAGELIST himl = (HIMAGELIST)ParamIndexToInt64(0); // Load-time validation has ensured there is a first parameter.
	if (!himl)
		return;

	int param3 = ParamIndexToOptionalInt(2, 0);
	int icon_number, width = 0, height = 0; // Zero width/height causes image to be loaded at its actual width/height.
	if (!ParamIndexIsOmitted(3)) // Presence of fourth parameter switches mode to be "load a non-icon image".
	{
		icon_number = 0; // Zero means "load icon or bitmap (doesn't matter)".
		if (ParamIndexToInt64(3)) // A value of True indicates that the image should be scaled to fit the imagelist's image size.
			ImageList_GetIconSize(himl, &width, &height); // Determine the width/height to which it should be scaled.
		//else retain defaults of zero for width/height, which loads the image at actual size, which in turn
		// lets ImageList_AddMasked() divide it up into separate images based on its width.
	}
	else
	{
		icon_number = param3; // LoadPicture() properly handles any wrong/negative value that might be here.
		ImageList_GetIconSize(himl, &width, &height); // L19: Determine the width/height of images in the image list to ensure icons are loaded at the correct size.
	}

	int image_type;
	HBITMAP hbitmap = LoadPicture(ParamIndexToString(1, buf) // Load-time validation has ensured there are at least two parameters.
		, width, height, image_type, icon_number, false); // Defaulting to "false" for "use GDIplus" provides more consistent appearance across multiple OSes.
	if (!hbitmap)
		return;

	if (image_type == IMAGE_BITMAP) // In this mode, param3 is always assumed to be an RGB color.
	{
		// Return the index of the new image or 0 on failure.
		aResultToken.value_int64 = ImageList_AddMasked(himl, hbitmap, rgb_to_bgr((int)param3)) + 1; // +1 to convert to one-based.
		DeleteObject(hbitmap);
	}
	else // ICON or CURSOR.
	{
		// Return the index of the new image or 0 on failure.
		aResultToken.value_int64 = ImageList_AddIcon(himl, (HICON)hbitmap) + 1; // +1 to convert to one-based.
		DestroyIcon((HICON)hbitmap); // Works on cursors too.  See notes in LoadPicture().
	}
}



BIF_DECL(BIF_LoadPicture)
{
	// h := LoadPicture(filename [, options, ByRef image_type])
	LPTSTR filename = ParamIndexToString(0, aResultToken.buf);
	LPTSTR options = ParamIndexToOptionalString(1);
	Var *image_type_var = ParamIndexToOptionalVar(2);

	int width = -1;
	int height = -1;
	int icon_number = 0;
	bool use_gdi_plus = false;

	for (LPTSTR cp = options; cp; cp = StrChrAny(cp, _T(" \t")))
	{
		cp = omit_leading_whitespace(cp);
		if (tolower(*cp) == 'w')
			width = ATOI(cp + 1);
		else if (tolower(*cp) == 'h')
			height = ATOI(cp + 1);
		else if (!_tcsnicmp(cp, _T("Icon"), 4))
			icon_number = ATOI(cp + 4);
		else if (!_tcsnicmp(cp, _T("GDI+"), 4))
			// GDI+ or GDI+1 to enable, GDI+0 to disable.
			use_gdi_plus = cp[4] != '0';
	}

	if (width == -1 && height == -1)
		width = 0;

	int image_type;
	HBITMAP hbm = LoadPicture(filename, width, height, image_type, icon_number, use_gdi_plus);
	if (image_type_var)
		image_type_var->Assign(image_type);
	else if (image_type != IMAGE_BITMAP && hbm)
		// Always return a bitmap when the ImageType output var is omitted.
		hbm = IconToBitmap32((HICON)hbm, true); // Also works for cursors.
	aResultToken.value_int64 = (__int64)(UINT_PTR)hbm;
}



BIF_DECL(BIF_Trim) // L31
{
	TCHAR trim_type = ctoupper(*aResultToken.marker); // aResultToken.marker points to the name of the Func which was called.

	// All return values are SYM_STRING, so set that now:
	aResultToken.symbol = SYM_STRING;

	if (IS_NUMERIC(aParam[0]->symbol))
	{
		// Take a shortcut for SYM_INTEGER/SYM_FLOAT.
		aResultToken.marker = TokenToString(*aParam[0], aResultToken.buf);
		return;
	}

	LPTSTR str = ParamIndexToString(0);
	LPTSTR result = str; // Prior validation has ensured at least 1 param.
	INT_PTR extract_length = ParamIndexLength(0, str);

	TCHAR omit_list_buf[MAX_NUMBER_SIZE]; // Support SYM_INTEGER/SYM_FLOAT even though it doesn't seem likely to happen.
	LPTSTR omit_list = ParamIndexIsOmitted(1) ? _T(" \t") : ParamIndexToString(1, omit_list_buf); // Default: space and tab.

	if (trim_type != 'R') // i.e. it's Trim() or LTrim()
	{
		result = omit_leading_any(result, omit_list, extract_length);
		extract_length -= (result - str); // Adjust for omitted characters.
	}
	if (extract_length && trim_type != 'L') // i.e. it's Trim() or RTrim();  THE LINE BELOW REQUIRES extract_length >= 1.
		extract_length = omit_trailing_any(result, omit_list, result + extract_length - 1);

	if (!TokenSetResult(aResultToken, result, extract_length))
		// Out of memory.
		aResultToken.marker = _T("");
}



BIF_DECL(BIF_Exception)
{
	LPTSTR message = ParamIndexToString(0, aResultToken.buf);
	TCHAR what_buf[MAX_NUMBER_SIZE], extra_buf[MAX_NUMBER_SIZE];
	LPTSTR what = NULL;
	Line *line = NULL;

	if (ParamIndexIsOmitted(1)) // "What"
	{
		line = g_script.mCurrLine;
		if (g->CurrentFunc)
			what = g->CurrentFunc->mName;
		else if (g->CurrentLabel)
			what = g->CurrentLabel->mName;
		else
			what = _T(""); // Probably the auto-execute section?
	}
	else
	{
#ifdef CONFIG_DEBUGGER
		int offset = TokenIsPureNumeric(*aParam[1]) ? ParamIndexToInt(1) : 0;
		if (offset < 0 && offset >= (g_Debugger.mStack.mBottom - g_Debugger.mStack.mTop)) // (mBottom - mTop) is safe against overflow, unlike (se >= mBottom). 
		{
			DbgStack::Entry *se = g_Debugger.mStack.mTop + offset;
			// Self-contained loop to ensure the entry belongs to the current thread
			// (below also relies on this loop to verify se[1].type != SE_Thread):
			while (++offset <= 0 && g_Debugger.mStack.mTop[offset].type != DbgStack::SE_Thread); // Relies on short-circuit evaluation.
			if (offset == 1)
			{
				line = se->line;
				// se->line contains the line at the given offset from the top of the stack.
				// Rather than returning the name of the function or sub which contains that
				// line, return the name of the function or sub which that line called.
				// In other words, an offset of -1 gives the name of the current function and
				// the file and number of the line which it was called from.
				what = se[1].type == DbgStack::SE_UDF ? se[1].udf->func->mName : se[1].sub->mName;
			}
		}
#endif
		if (!what)
		{
			line = g_script.mCurrLine;
			what = ParamIndexToString(1, what_buf);
		}
	}

	LPTSTR extra = ParamIndexToOptionalString(2, extra_buf);

	if (aResultToken.object = line->CreateRuntimeException(message, what, extra))
	{
		aResultToken.symbol = SYM_OBJECT;
	}
	else
	{
		// Out of memory. Seems best to alert the user.
		MsgBox(ERR_OUTOFMEM);
		aResultToken.value_int64 = 0;
	}
}



////////////////////////////////////////////////////////
// HELPER FUNCTIONS FOR TOKENS AND BUILT-IN FUNCTIONS //
////////////////////////////////////////////////////////

BOOL LegacyResultToBOOL(LPTSTR aResult)
// The logic here is similar to LegacyVarToBOOL(), so maintain them together.
// This is called "Legacy" because the following method of casting expression results to BOOL is inconsistent
// with the method used internally by expressions for intermediate results. This inconsistency is retained
// for backward compatibility.  On the other hand, some users say they would prefer using this method exclusively
// rather than TokenToBOOL() so that expressions internally treat numerically-zero strings as zero just like
// they do for variables (which will probably remain true as long as variables stay generic/untyped).
{
	UINT c1 = (UINT)*aResult; // UINT vs. UCHAR might squeeze a little more performance out of it.
	if (c1 > 48)     // Any UCHAR greater than '0' can't be a space(32), tab(9), '+'(43), or '-'(45), or '.'(46)...
		return TRUE; // ...so any string starting with c1>48 can't be anything that's false (e.g. " 0", "+0", "-0", ".0", "0").
	if (!c1 // Besides performance, must be checked anyway because otherwise IsPureNumeric() would consider "" to be non-numeric and thus TRUE.
		|| c1 == 48 && !aResult[1]) // The string "0" seems common enough to detect explicitly, for performance.
		return FALSE;
	// IsPureNumeric() is called below because there are many variants of a false string:
	// e.g. "0", "0.0", "0x0", ".0", "+0", "-0", and " 0" (leading spaces/tabs).
	switch (IsPureNumeric(aResult, true, false, true)) // It's purely numeric and not all whitespace (and due to earlier check, it's not blank).
	{
	case PURE_INTEGER: return ATOI64(aResult) != 0; // Could call ATOF() for both integers and floats; but ATOI64() probably performs better, and a float comparison to 0.0 might be a slower than an integer comparison to 0.
	case PURE_FLOAT:   return _tstof(aResult) != 0.0; // atof() vs. ATOF() because PURE_FLOAT is never hexadecimal.
	default: // PURE_NOT_NUMERIC.
		// Even a string containing all whitespace would be considered non-numeric since it's a non-blank string
		// that isn't equal to 0.
		return TRUE;
	}
}



BOOL LegacyVarToBOOL(Var &aVar)
// For comments see LegacyResultToBOOL().
{
	if (!aVar.HasContents()) // Must be checked first because otherwise IsPureNumeric() would consider "" to be non-numeric and thus TRUE.  For performance, it also exploits the binary number cache.
	{
		aVar.MaybeWarnUninitialized();
		return FALSE;
	}
	switch (aVar.IsNonBlankIntegerOrFloat()) // See comments in LegacyResultToBOOL().
	{
	case PURE_INTEGER: return aVar.ToInt64(TRUE) != 0;
	case PURE_FLOAT:   return aVar.ToDouble(TRUE) != 0.0;
	default: // PURE_NOT_NUMERIC.
		return TRUE; // See comments in LegacyResultToBOOL().
	}
}



BOOL TokenToBOOL(ExprTokenType &aToken, SymbolType aTokenIsNumber)
{
	if (aTokenIsNumber == SYM_INVALID) // Omitted.
		aTokenIsNumber = TokenIsPureNumeric(aToken);

	switch (aTokenIsNumber)
	{
	case PURE_INTEGER: // Probably the most common; e.g. both sides of "if (x>3 and x<6)" are the number 1/0.
		return TokenToInt64(aToken, TRUE) != 0; // Force it to be purely 1 or 0.
	case PURE_FLOAT: // Convert to float, not int, so that a number between 0.0001 and 0.9999 is considered "true".
		return TokenToDouble(aToken, FALSE, TRUE) != 0.0; // Pass FALSE for aCheckForHex since PURE_FLOAT is never hex.
	default:
		if (aToken.symbol == SYM_OBJECT || (aToken.symbol == SYM_VAR && aToken.var->HasObject())) // L31: Treat objects as "true".
			return TRUE;
		// This generally includes all SYM_STRINGs (even ones that are all digits such as "123") and all
		// non-numeric SYM_OPERANDS.
		// The only tokens considered FALSE are numeric 0 or 0.0, or an empty string.  All non-blank
		// SYM_STRINGs are considered TRUE.  This includes literal strings like "0", and subexpressions that
		// evaluate to pure SYM_STRING that isn't blank.
		return *TokenToString(aToken) != '\0'; // Force it to be purely 1 or 0.
	}
}



SymbolType TokenIsPureNumeric(ExprTokenType &aToken)
{
	switch(aToken.symbol)
	{
	case SYM_INTEGER:
	case SYM_FLOAT:
		return aToken.symbol;
	case SYM_VAR:     return aToken.var->IsNonBlankIntegerOrFloat(); // Supports VAR_NORMAL and VAR_CLIPBOARD.
	case SYM_OPERAND: return aToken.buf ? PURE_INTEGER // The "buf" of a SYM_OPERAND is non-NULL if it's a pure integer.
			: IsPureNumeric(TokenToString(aToken), true, false, true);
	//case SYM_STRING:
	//case SYM_OBJECT: // L31: Objects are currently treated as empty strings in most cases.
	//case SYM_MISSING:
	default:
		return PURE_NOT_NUMERIC; // Explicitly-marked strings are not numeric, which allows numeric strings to be compared as strings rather than as numbers.
	}
}


BOOL TokenIsEmptyString(ExprTokenType &aToken) // L31
{
	switch (aToken.symbol)
	{
	case SYM_OPERAND:
	case SYM_STRING:
		return !*aToken.marker;
	case SYM_VAR:
		return !aToken.var->HasContents();
	//case SYM_MISSING: // This case is omitted because it currently should be
		// impossible for all callers except for ParamIndexIsOmittedOrEmpty(),
		// which checks for it explicitly.
		//return TRUE;
	default:
		return FALSE;
	}
}


BOOL TokenIsEmptyString(ExprTokenType &aToken, BOOL aWarnUninitializedVar)
{
	if (aWarnUninitializedVar && aToken.symbol == SYM_VAR)
		aToken.var->MaybeWarnUninitialized();

	return TokenIsEmptyString(aToken);
}


__int64 TokenToInt64(ExprTokenType &aToken, BOOL aIsPureInteger)
// Caller has ensured that any SYM_VAR's Type() is VAR_NORMAL or VAR_CLIPBOARD.
// Converts the contents of aToken to a 64-bit int.
// Caller should pass FALSE for aIsPureInteger if aToken.var is either:
// 1) Not a pure number (i.e. 123abc).
// 2) It isn't known whether it's a pure number.
// 3) It's pure but it's a floating point number, not an integer.
{
	// Some callers, such as those that cast our return value to UINT, rely on the use of 64-bit to preserve
	// unsigned values and also wrap any signed values into the unsigned domain.
	switch (aToken.symbol)
	{
		case SYM_INTEGER: return aToken.value_int64; // Fixed in v1.0.45 not to cast to int.
		case SYM_OPERAND: // Listed near the top for performance.
			if (aToken.buf) // The "buf" of a SYM_OPERAND is non-NULL if points to a pure integer.
				return *(__int64 *)aToken.buf;
			//else don't return; continue on to the bottom.
			break;
		case SYM_FLOAT: return (__int64)aToken.value_double; // 1.0.48: fixed to cast to __int64 vs. int.
		case SYM_VAR: return aToken.var->ToInt64(aIsPureInteger);
		case SYM_OBJECT: return 0; // L31: Returning the IObject* doesn't seem appropriate since it would rarely be useful, and other uses treat objects as an empty string.
	}
	// Since above didn't return, it's SYM_STRING, or a SYM_OPERAND that lacks a binary-integer counterpart.
	return ATOI64(aToken.marker); // Fixed in v1.0.45 to use ATOI64 vs. ATOI().
}



double TokenToDouble(ExprTokenType &aToken, BOOL aCheckForHex, BOOL aIsPureFloat)
// Caller has ensured that any SYM_VAR's Type() is VAR_NORMAL or VAR_CLIPBOARD.
// Converts the contents of aToken to a double. Caller should pass FALSE for aIsPureFloat if aToken.var
// is either:
// 1) Not a pure number (i.e. 123abc).
// 2) It isn't known whether it's a pure number.
// 3) It's pure but it's a an integer, not a floating point number.
{
	switch (aToken.symbol)
	{
		case SYM_INTEGER: return (double)aToken.value_int64;
		case SYM_FLOAT: return aToken.value_double;
		case SYM_VAR: return aToken.var->ToDouble(aIsPureFloat);
		case SYM_OPERAND:
			if (aToken.buf) // The "buf" of a SYM_OPERAND is non-NULL if it's a pure integer.
				return (double)*(__int64 *)aToken.buf;
			//else continue on to the bottom.
			break;
		case SYM_OBJECT: return 0; // L31: Treat like empty string.
	}
	// Since above didn't return, it's SYM_STRING or a SYM_OPERAND that lacks a binary-integer counterpart.
	return aCheckForHex ? ATOF(aToken.marker) : _tstof(aToken.marker); // atof() omits the check for hexadecimal.
}



LPTSTR TokenToString(ExprTokenType &aToken, LPTSTR aBuf)
// Supports Type() VAR_NORMAL and VAR-CLIPBOARD.
// Returns "" on failure to simplify logic in callers.  Otherwise, it returns either aBuf (if aBuf was needed
// for the conversion) or the token's own string.  aBuf may be NULL, in which case the caller presumably knows
// that this token is SYM_STRING or SYM_OPERAND (or caller wants "" back for anything other than those).
// If aBuf is not NULL, caller has ensured that aBuf is at least MAX_NUMBER_SIZE in size.
{
	switch (aToken.symbol)
	{
	case SYM_VAR: // Caller has ensured that any SYM_VAR's Type() is VAR_NORMAL.
		return aToken.var->Contents(); // Contents() vs. mContents to support VAR_CLIPBOARD, and in case mContents needs to be updated by Contents().
	case SYM_STRING:
	case SYM_OPERAND:
		return aToken.marker;
	case SYM_INTEGER:
		if (aBuf)
			return ITOA64(aToken.value_int64, aBuf);
		//else continue on to return the default at the bottom.
		break;
	case SYM_FLOAT:
		if (aBuf)
		{
			sntprintf(aBuf, MAX_NUMBER_SIZE, g->FormatFloat, aToken.value_double);
			return aBuf;
		}
		//else continue on to return the default at the bottom.
		break;
	//case SYM_OBJECT: // L31: Treat objects as empty strings (or TRUE where appropriate).
	//default: // Not an operand: continue on to return the default at the bottom.
	}
	return _T("");
}



ResultType TokenToDoubleOrInt64(const ExprTokenType &aInput, ExprTokenType &aOutput)
// Converts aToken's contents to a numeric value, either int64 or double (whichever is more appropriate).
// Returns FAIL when aToken isn't an operand or is but contains a string that isn't purely numeric.
{
	LPTSTR str;
	switch (aInput.symbol)
	{
		case SYM_INTEGER:
		case SYM_FLOAT:
			aOutput.symbol = aInput.symbol;
			aOutput.value_int64 = aInput.value_int64;
			return OK;
		case SYM_VAR:
			return aInput.var->ToDoubleOrInt64(aOutput);
		case SYM_STRING:   // v1.0.40.06: Fixed to be listed explicitly so that "default" case can return failure.
			str = aInput.marker;
			break;
		case SYM_OPERAND:
			if (aInput.buf) // The "buf" of a SYM_OPERAND is non-NULL if it's a pure integer.
			{
				aOutput.symbol = SYM_INTEGER;
				aOutput.value_int64 = *(__int64 *)aInput.buf;
				return OK;
			}
			// Otherwise:
			str = aInput.marker;
			break;
		//case SYM_OBJECT: // L31: Treat objects as empty strings (or TRUE where appropriate).
		//case SYM_MISSING:
		default:
			aOutput.symbol = SYM_STRING;
			aOutput.marker = _T(""); // For completeness.  Some callers such as BIF_Abs() rely on this being done.
			// FALL THROUGH TO BELOW
			return FAIL;
	}
	// Since above didn't return, interpret "str" as a number.
	switch (aOutput.symbol = IsPureNumeric(str, true, false, true))
	{
	case PURE_INTEGER:
		aOutput.value_int64 = ATOI64(str);
		break;
	case PURE_FLOAT:
		aOutput.value_double = ATOF(str);
		break;
	default: // Not a pure number.
		aOutput.marker = _T(""); // For completeness.  Some callers such as BIF_Abs() rely on this being done.
		return FAIL;
	}
	return OK; // Since above didn't return, indicate success.
}



IObject *TokenToObject(ExprTokenType &aToken)
// L31: Returns IObject* from SYM_OBJECT or SYM_VAR (where var->HasObject()), NULL for other tokens.
// Caller is responsible for calling AddRef() if that is appropriate.
{
	if (aToken.symbol == SYM_OBJECT)
		return aToken.object;
	
	if (aToken.symbol == SYM_VAR)
	{
		if (aToken.var->HasObject())
			return aToken.var->Object();

		aToken.var->MaybeWarnUninitialized();
	}

	return NULL;
}



Func *TokenToFunc(ExprTokenType &aToken)
{
	// No need for buf since function names can't be pure numeric:
	//TCHAR buf[MAX_NUMBER_SIZE];
	Func *func;
	if (  !(func = dynamic_cast<Func *>(TokenToObject(aToken)))  )
	{
		LPTSTR func_name = TokenToString(aToken);
		// Dynamic function calls rely on the following check to avoid a lengthy and
		// futile search through all function and action names when aToken is an object
		// emulating a function.  The check works because TokenToString() returns ""
		// when aToken is an object (or a pure number, since no buffer was passed).
		if (*func_name)
			func = g_script.FindFunc(func_name);
	}
	return func;
}



ResultType TokenSetResult(ExprTokenType &aResultToken, LPCTSTR aResult, size_t aResultLength)
// Utility function for handling memory allocation and return to callers of built-in functions; based on BIF_SubStr.
// Returns FAIL if malloc failed, in which case our caller is responsible for returning a sensible default value.
{
	if (aResultLength == -1)
		aResultLength = _tcslen(aResult); // Caller must not pass NULL for aResult in this case.
	if (aResultLength <= MAX_NUMBER_LENGTH) // v1.0.46.01: Avoid malloc() for small strings.  However, this improves speed by only 10% in a test where random 25-byte strings were extracted from a 700 KB string (probably because VC++'s malloc()/free() are very fast for small allocations).
		aResultToken.marker = aResultToken.buf; // Store the address of the result for the caller.
	else
	{
		// Caller has provided a mem_to_free (initially NULL) as a means of passing back memory we allocate here.
		// So if we change "result" to be non-NULL, the caller will take over responsibility for freeing that memory.
		if (   !(aResultToken.mem_to_free = tmalloc(aResultLength + 1))   ) // Out of memory. Due to rarity, don't display an error dialog (there's currently no way for a built-in function to abort the current thread anyway?)
			return FAIL;
		aResultToken.marker = aResultToken.mem_to_free; // Store the address of the result for the caller.
		aResultToken.marker_length = aResultLength; // MANDATORY FOR USERS OF CIRCUIT_TOKEN: set marker_length to the length of the string in marker.
	}
	if (aResult) // Caller may pass NULL to retrieve a buffer of sufficient size.
		tmemcpy(aResultToken.marker, aResult, aResultLength);
	aResultToken.marker[aResultLength] = '\0'; // Must be done separately from the memcpy() because the memcpy() might just be taking a substring (i.e. long before result's terminator).
	return OK;
}



int ConvertJoy(LPTSTR aBuf, int *aJoystickID, bool aAllowOnlyButtons)
// The caller TextToKey() currently relies on the fact that when aAllowOnlyButtons==true, a value
// that can fit in a sc_type (USHORT) is returned, which is true since the joystick buttons
// are very small numbers (JOYCTRL_1==12).
{
	if (aJoystickID)
		*aJoystickID = 0;  // Set default output value for the caller.
	if (!aBuf || !*aBuf) return JOYCTRL_INVALID;
	LPTSTR aBuf_orig = aBuf;
	for (; *aBuf >= '0' && *aBuf <= '9'; ++aBuf); // self-contained loop to find the first non-digit.
	if (aBuf > aBuf_orig) // The string starts with a number.
	{
		int joystick_id = ATOI(aBuf_orig) - 1;
		if (joystick_id < 0 || joystick_id >= MAX_JOYSTICKS)
			return JOYCTRL_INVALID;
		if (aJoystickID)
			*aJoystickID = joystick_id;  // Use ATOI vs. atoi even though hex isn't supported yet.
	}

	if (!_tcsnicmp(aBuf, _T("Joy"), 3))
	{
		if (IsPureNumeric(aBuf + 3, false, false))
		{
			int offset = ATOI(aBuf + 3);
			if (offset < 1 || offset > MAX_JOY_BUTTONS)
				return JOYCTRL_INVALID;
			return JOYCTRL_1 + offset - 1;
		}
	}
	if (aAllowOnlyButtons)
		return JOYCTRL_INVALID;

	// Otherwise:
	if (!_tcsicmp(aBuf, _T("JoyX"))) return JOYCTRL_XPOS;
	if (!_tcsicmp(aBuf, _T("JoyY"))) return JOYCTRL_YPOS;
	if (!_tcsicmp(aBuf, _T("JoyZ"))) return JOYCTRL_ZPOS;
	if (!_tcsicmp(aBuf, _T("JoyR"))) return JOYCTRL_RPOS;
	if (!_tcsicmp(aBuf, _T("JoyU"))) return JOYCTRL_UPOS;
	if (!_tcsicmp(aBuf, _T("JoyV"))) return JOYCTRL_VPOS;
	if (!_tcsicmp(aBuf, _T("JoyPOV"))) return JOYCTRL_POV;
	if (!_tcsicmp(aBuf, _T("JoyName"))) return JOYCTRL_NAME;
	if (!_tcsicmp(aBuf, _T("JoyButtons"))) return JOYCTRL_BUTTONS;
	if (!_tcsicmp(aBuf, _T("JoyAxes"))) return JOYCTRL_AXES;
	if (!_tcsicmp(aBuf, _T("JoyInfo"))) return JOYCTRL_INFO;
	return JOYCTRL_INVALID;
}



bool ScriptGetKeyState(vk_type aVK, KeyStateTypes aKeyStateType)
// Returns true if "down", false if "up".
{
    if (!aVK) // Assume "up" if indeterminate.
		return false;

	switch (aKeyStateType)
	{
	case KEYSTATE_TOGGLE: // Whether a toggleable key such as CapsLock is currently turned on.
		// Under Win9x, at least certain versions and for certain hardware, this
		// doesn't seem to be always accurate, especially when the key has just
		// been toggled and the user hasn't pressed any other key since then.
		// I tried using GetKeyboardState() instead, but it produces the same
		// result.  Therefore, I've documented this as a limitation in the help file.
		// In addition, this was attempted but it didn't seem to help:
		//if (g_os.IsWin9x())
		//{
		//	DWORD fore_thread = GetWindowThreadProcessId(GetForegroundWindow(), NULL);
		//	bool is_attached_my_to_fore = false;
		//	if (fore_thread && fore_thread != g_MainThreadID)
		//		is_attached_my_to_fore = AttachThreadInput(g_MainThreadID, fore_thread, TRUE) != 0;
		//	output_var->Assign(IsKeyToggledOn(aVK) ? "D" : "U");
		//	if (is_attached_my_to_fore)
		//		AttachThreadInput(g_MainThreadID, fore_thread, FALSE);
		//	return OK;
		//}
		//else
		return IsKeyToggledOn(aVK); // This also works for the INSERT key, but only on XP (and possibly Win2k).
	case KEYSTATE_PHYSICAL: // Physical state of key.
		if (IsMouseVK(aVK)) // mouse button
		{
			if (g_MouseHook) // mouse hook is installed, so use it's tracking of physical state.
				return g_PhysicalKeyState[aVK] & STATE_DOWN;
			else // Even for Win9x/NT, it seems slightly better to call this rather than IsKeyDown9xNT():
				return IsKeyDownAsync(aVK);
		}
		else // keyboard
		{
			if (g_KeybdHook)
			{
				// Since the hook is installed, use its value rather than that from
				// GetAsyncKeyState(), which doesn't seem to return the physical state.
				// But first, correct the hook modifier state if it needs it.  See comments
				// in GetModifierLRState() for why this is needed:
				if (KeyToModifiersLR(aVK))    // It's a modifier.
					GetModifierLRState(true); // Correct hook's physical state if needed.
				return g_PhysicalKeyState[aVK] & STATE_DOWN;
			}
			else // Even for Win9x/NT, it seems slightly better to call this rather than IsKeyDown9xNT():
				return IsKeyDownAsync(aVK);
		}
	} // switch()

	// Otherwise, use the default state-type: KEYSTATE_LOGICAL
	if (g_os.IsWin9x() || g_os.IsWinNT4())
		return IsKeyDown9xNT(aVK); // See its comments for why it's more reliable on these OSes.
	else
		// On XP/2K at least, a key can be physically down even if it isn't logically down,
		// which is why the below specifically calls IsKeyDown2kXP() rather than some more
		// comprehensive method such as consulting the physical key state as tracked by the hook:
		// v1.0.42.01: For backward compatibility, the following hasn't been changed to IsKeyDownAsync().
		// For example, a script might rely on being able to detect whether Control was down at the
		// time the current Gui thread was launched rather than whether than whether it's down right now.
		// Another example is the journal playback hook: when a window owned by the script receives
		// such a keystroke, only GetKeyState() can detect the changed state of the key, not GetAsyncKeyState().
		// A new mode can be added to KeyWait & GetKeyState if Async is ever explicitly needed.
		return IsKeyDown2kXP(aVK);
		// Known limitation: For some reason, both the above and IsKeyDown9xNT() will indicate
		// that the CONTROL key is up whenever RButton is down, at least if the mouse hook is
		// installed without the keyboard hook.  No known explanation.
}



double ScriptGetJoyState(JoyControls aJoy, int aJoystickID, ExprTokenType &aToken, bool aUseBoolForUpDown)
// Caller must ensure that aToken.marker is a buffer large enough to handle the longest thing put into
// it here, which is currently jc.szPname (size=32). Caller has set aToken.symbol to be SYM_STRING.
// For buttons: Returns 0 if "up", non-zero if down.
// For axes and other controls: Returns a number indicating that controls position or status.
// If there was a problem determining the position/state, aToken is made blank and zero is returned.
// Also returns zero in cases where a non-numerical result is requested, such as the joystick name.
// In those cases, caller should use aToken.marker as the result.
{
	// Set default in case of early return.
	*aToken.marker = '\0'; // Blank vs. string "0" serves as an indication of failure.

	if (!aJoy) // Currently never called this way.
		return 0; // And leave aToken set to blank.

	bool aJoy_is_button = IS_JOYSTICK_BUTTON(aJoy);

	JOYCAPS jc;
	if (!aJoy_is_button && aJoy != JOYCTRL_POV)
	{
		// Get the joystick's range of motion so that we can report position as a percentage.
		if (joyGetDevCaps(aJoystickID, &jc, sizeof(JOYCAPS)) != JOYERR_NOERROR)
			ZeroMemory(&jc, sizeof(jc));  // Zero it on failure, for use of the zeroes later below.
	}

	// Fetch this struct's info only if needed:
	JOYINFOEX jie;
	if (aJoy != JOYCTRL_NAME && aJoy != JOYCTRL_BUTTONS && aJoy != JOYCTRL_AXES && aJoy != JOYCTRL_INFO)
	{
		jie.dwSize = sizeof(JOYINFOEX);
		jie.dwFlags = JOY_RETURNALL;
		if (joyGetPosEx(aJoystickID, &jie) != JOYERR_NOERROR)
			return 0; // And leave aToken set to blank.
		if (aJoy_is_button)
		{
			bool is_down = ((jie.dwButtons >> (aJoy - JOYCTRL_1)) & (DWORD)0x01);
			if (aUseBoolForUpDown) // i.e. Down==true and Up==false
			{
				aToken.symbol = SYM_INTEGER; // Override default type.
				aToken.value_int64 = is_down; // Forced to be 1 or 0 above, since it's "bool".
			}
			else
			{
				aToken.marker[0] = is_down ? 'D' : 'U';
				aToken.marker[1] = '\0';
			}
			return is_down;
		}
	}

	// Otherwise:
	UINT range;
	LPTSTR buf_ptr;
	double result_double;  // Not initialized to help catch bugs.

	switch(aJoy)
	{
	case JOYCTRL_XPOS:
		range = (jc.wXmax > jc.wXmin) ? jc.wXmax - jc.wXmin : 0;
		result_double = range ? 100 * (double)jie.dwXpos / range : jie.dwXpos;
		break;
	case JOYCTRL_YPOS:
		range = (jc.wYmax > jc.wYmin) ? jc.wYmax - jc.wYmin : 0;
		result_double = range ? 100 * (double)jie.dwYpos / range : jie.dwYpos;
		break;
	case JOYCTRL_ZPOS:
		range = (jc.wZmax > jc.wZmin) ? jc.wZmax - jc.wZmin : 0;
		result_double = range ? 100 * (double)jie.dwZpos / range : jie.dwZpos;
		break;
	case JOYCTRL_RPOS:  // Rudder or 4th axis.
		range = (jc.wRmax > jc.wRmin) ? jc.wRmax - jc.wRmin : 0;
		result_double = range ? 100 * (double)jie.dwRpos / range : jie.dwRpos;
		break;
	case JOYCTRL_UPOS:  // 5th axis.
		range = (jc.wUmax > jc.wUmin) ? jc.wUmax - jc.wUmin : 0;
		result_double = range ? 100 * (double)jie.dwUpos / range : jie.dwUpos;
		break;
	case JOYCTRL_VPOS:  // 6th axis.
		range = (jc.wVmax > jc.wVmin) ? jc.wVmax - jc.wVmin : 0;
		result_double = range ? 100 * (double)jie.dwVpos / range : jie.dwVpos;
		break;

	case JOYCTRL_POV:  // Need to explicitly compare against JOY_POVCENTERED because it's a WORD not a DWORD.
		if (jie.dwPOV == JOY_POVCENTERED)
		{
			// Retain default SYM_STRING type.
			_tcscpy(aToken.marker, _T("-1")); // Assign as string to ensure its written exactly as "-1". Documented behavior.
			return -1;
		}
		else
		{
			aToken.symbol = SYM_INTEGER; // Override default type.
			aToken.value_int64 = jie.dwPOV;
			return jie.dwPOV;
		}
		// No break since above always returns.

	case JOYCTRL_NAME:
		_tcscpy(aToken.marker, jc.szPname);
		return 0;  // Returns zero in cases where a non-numerical result is obtained.

	case JOYCTRL_BUTTONS:
		aToken.symbol = SYM_INTEGER; // Override default type.
		aToken.value_int64 = jc.wNumButtons;
		return jc.wNumButtons;  // wMaxButtons is the *driver's* max supported buttons.

	case JOYCTRL_AXES:
		aToken.symbol = SYM_INTEGER; // Override default type.
		aToken.value_int64 = jc.wNumAxes; // wMaxAxes is the *driver's* max supported axes.
		return jc.wNumAxes;

	case JOYCTRL_INFO:
		buf_ptr = aToken.marker;
		if (jc.wCaps & JOYCAPS_HASZ)
			*buf_ptr++ = 'Z';
		if (jc.wCaps & JOYCAPS_HASR)
			*buf_ptr++ = 'R';
		if (jc.wCaps & JOYCAPS_HASU)
			*buf_ptr++ = 'U';
		if (jc.wCaps & JOYCAPS_HASV)
			*buf_ptr++ = 'V';
		if (jc.wCaps & JOYCAPS_HASPOV)
		{
			*buf_ptr++ = 'P';
			if (jc.wCaps & JOYCAPS_POV4DIR)
				*buf_ptr++ = 'D';
			if (jc.wCaps & JOYCAPS_POVCTS)
				*buf_ptr++ = 'C';
		}
		*buf_ptr = '\0'; // Final termination.
		return 0;  // Returns zero in cases where a non-numerical result is obtained.
	} // switch()

	// If above didn't return, the result should now be in result_double.
	aToken.symbol = SYM_FLOAT; // Override default type.
	aToken.value_double = result_double;
	return result_double;
}



DWORD GetProcessName(DWORD aProcessID, LPTSTR aBuf, DWORD aBufSize, bool aGetNameOnly)
{
	*aBuf = '\0'; // Set default.
	HANDLE hproc;
	if (  !(hproc = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, aProcessID))  )
		// OpenProcess failed, so try fallback access; this will probably cause the
		// first method below to fail and fall back to GetProcessImageFileName.
		if (  !(hproc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, aProcessID))  )
			return 0;

	// Attempt these first, since they return exactly what we want and are available on Win2k:
	DWORD buf_length = aGetNameOnly
		? GetModuleBaseName(hproc, NULL, aBuf, aBufSize)
		: GetModuleFileNameEx(hproc, NULL, aBuf, aBufSize);

	typedef DWORD (WINAPI *MyGetName)(HANDLE, LPTSTR, DWORD);
#ifdef CONFIG_WIN2K
	// This must be loaded dynamically or the program will probably not launch at all on Win2k:
	static MyGetName lpfnGetName = (MyGetName)GetProcAddress(GetModuleHandle(_T("psapi")), "GetProcessImageFileName" WINAPI_SUFFIX);;
#else
	static MyGetName lpfnGetName = &GetProcessImageFileName;
#endif

	if (!buf_length && lpfnGetName)
	{
		// Above failed, possibly for one of the following reasons:
		//	- Our process is 32-bit, but that one is 64-bit.
		//	- That process is running at a higher integrity level (UAC is interfering).
		//	- We didn't have permission to use PROCESS_VM_READ access?
		//
		// So fall back to GetProcessImageFileName (XP or later required):
		buf_length = lpfnGetName(hproc, aBuf, aBufSize);
		if (buf_length)
		{
			LPTSTR cp;
			if (aGetNameOnly)
			{
				// Convert full path to just name.
				cp = _tcsrchr(aBuf, '\\');
				if (cp)
					tmemmove(aBuf, cp + 1, _tcslen(cp)); // Includes the null terminator.
			}
			else
			{
				// Convert device path to logical path.
				TCHAR device_path[MAX_PATH];
				TCHAR letter[3];
				letter[1] = ':';
				letter[2] = '\0';
				// For simplicity and because GetLogicalDriveStrings does not exist on Win2k, it is not used.
				for (*letter = 'A'; *letter <= 'Z'; ++(*letter))
				{
					DWORD device_path_length = QueryDosDevice(letter, device_path, _countof(device_path));
					if (device_path_length > 2) // Includes two null terminators.
					{
						device_path_length -= 2;
						if (!_tcsncmp(device_path, aBuf, device_path_length)
							&& aBuf[device_path_length] == '\\') // Relies on short-circuit evaluation.
						{
							// Copy drive letter:
							aBuf[0] = letter[0];
							aBuf[1] = letter[1];
							// Contract path to remove remainder of device name.
							tmemmove(aBuf + 2, aBuf + device_path_length, buf_length - device_path_length + 1);
							buf_length -= device_path_length - 2;
							break;
						}
					}
				}
			}
		}
	}

	CloseHandle(hproc);
	return buf_length;
}
