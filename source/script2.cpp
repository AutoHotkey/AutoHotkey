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
#include <winioctl.h> // For PREVENT_MEDIA_REMOVAL and CD lock/unlock.
#include "qmath.h" // Used by Transform() [math.h incurs 2k larger code size just for ceil() & floor()]
#include "script.h"
#include "window.h" // for IF_USE_FOREGROUND_WINDOW
#include "application.h" // for MsgSleep()
#include "resources/resource.h"  // For InputBox.
#include "TextIO.h"
#include <Psapi.h> // for GetModuleBaseName.
#include "shlwapi.h" // StrCmpLogicalW

#include <mmdeviceapi.h> // for SoundSet/SoundGet.
#include <endpointvolume.h> // for SoundSet/SoundGet.
#include <functiondiscoverykeys.h>

#define PCRE_STATIC             // For RegEx. PCRE_STATIC tells PCRE to declare its functions for normal, static
#include "lib_pcre/pcre/pcre.h" // linkage rather than as functions inside an external DLL.

#include "script_func_impl.h"



////////////////////
// Window related //
////////////////////



ResultType Line::ToolTip(LPTSTR aText, LPTSTR aX, LPTSTR aY, LPTSTR aID)
{
	int window_index = *aID ? ATOI(aID) - 1 : 0;
	if (window_index < 0 || window_index >= MAX_TOOLTIPS)
		return LineError(_T("Max window number is ") MAX_TOOLTIPS_STR _T("."), FAIL_OR_OK, aID);
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
	ti.cbSize = sizeof(ti);
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



ResultType TrayTipParseOptions(LPTSTR aOptions, NOTIFYICONDATA &nic)
{
	LPTSTR next_option, option_end;
	TCHAR option[1+MAX_NUMBER_SIZE];
	for (next_option = omit_leading_whitespace(aOptions); ; next_option = omit_leading_whitespace(option_end))
	{
		if (!*next_option)
			return OK;

		// Find the end of this option item:
		if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
			option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.
		size_t option_length = option_end - next_option;

		// Make a terminated copy for simplicity and to reduce ambiguity:
		if (option_length + 1 > _countof(option))
			goto invalid_option;
		tmemcpy(option, next_option, option_length);
		option[option_length] = '\0';

		if (option_length <= 5 && !_tcsnicmp(option, _T("Icon"), 4))
		{
			nic.dwInfoFlags &= ~NIIF_ICON_MASK;
			switch (option[4])
			{
			case 'x': case 'X': nic.dwInfoFlags |= NIIF_ERROR; break;
			case '!': nic.dwInfoFlags |= NIIF_WARNING; break;
			case 'i': case 'I': nic.dwInfoFlags |= NIIF_INFO; break;
			case '\0': break;
			default:
				goto invalid_option;
			}
		}
		else if (!_tcsicmp(option, _T("Mute")))
		{
			nic.dwInfoFlags |= NIIF_NOSOUND;
		}
		else if (IsNumeric(option, FALSE, FALSE, FALSE))
		{
			nic.dwInfoFlags |= ATOI(option);
		}
		else
		{
			goto invalid_option;
		}
	}
invalid_option:
	return ValueError(ERR_INVALID_OPTION, next_option, FAIL_OR_OK);
}


ResultType Line::TrayTip(LPTSTR aText, LPTSTR aTitle, LPTSTR aOptions)
{
	NOTIFYICONDATA nic = {0};
	nic.cbSize = sizeof(nic);
	nic.uID = AHK_NOTIFYICON;  // This must match our tray icon's uID or Shell_NotifyIcon() will return failure.
	nic.hWnd = g_hWnd;
	nic.uFlags = NIF_INFO;
	// nic.uTimeout is no longer used because it is valid only on Windows 2000 and Windows XP.
	if (!TrayTipParseOptions(aOptions, nic))
		return FAIL;
	if (*aTitle && !*aText)
		// As passing an empty string hides the TrayTip (or does nothing on Windows 10),
		// pass a space to ensure the TrayTip is shown.  Testing showed that Windows 10
		// will size the notification to fit only the title, as if there was no text.
		aText = _T(" ");
	if (nic.dwInfoFlags & NIIF_USER)
	{
		// Windows 10 toast notifications display the small tray icon stretched to the
		// large size if NIIF_USER is passed but without NIIF_LARGE_ICON or hBalloonIcon.
		// If a large icon is passed without the flag, the notification does not show at all.
		// But since this could change, let the script pass 0x24 to use the large icon.
		//if (g_os.IsWin10OrLater())
		//	nic.dwInfoFlags |= NIIF_LARGE_ICON;
		if (nic.dwInfoFlags & NIIF_LARGE_ICON)
			nic.hBalloonIcon = g_script.mCustomIcon ? g_script.mCustomIcon : g_IconLarge;
		else
			nic.hBalloonIcon = g_script.mCustomIconSmall ? g_script.mCustomIconSmall : g_IconSmall;
	}
	tcslcpy(nic.szInfoTitle, aTitle, _countof(nic.szInfoTitle)); // Empty title omits the title line entirely.
	tcslcpy(nic.szInfo, aText, _countof(nic.szInfo));	// Empty text removes the balloon.
	if (!Shell_NotifyIcon(NIM_MODIFY, &nic) && (nic.dwInfoFlags & NIIF_USER))
	{
		// Passing NIIF_USER without NIIF_LARGE_ICON on Windows 10.0.19018 caused failure,
		// even though a small icon is displayed by default (without NIIF_USER), the docs
		// indicate it should work, and it works on Vista.  There's a good chance that
		// removing the flag and trying again will produce the desired result, or at least
		// show a TrayTip without the icon, which is preferable to complete failure.
		nic.dwInfoFlags &= ~NIIF_USER;
		Shell_NotifyIcon(NIM_MODIFY, &nic);
	}
	return OK; // i.e. never a critical error if it fails.
}



BIF_DECL(BIF_TraySetIcon)
{
	if (!g_script.SetTrayIcon(
		ParamIndexToOptionalString(0, _f_number_buf) // buf is provided for error-reporting purposes.
		, ParamIndexToOptionalInt(1, 1)
		, ParamIndexIsOmitted(2) ? NEUTRAL : ParamIndexToBOOL(2) ? TOGGLED_ON : TOGGLED_OFF))
		_f_return_FAIL;
	_f_return_empty;
}



ResultType input_type::Setup(LPTSTR aOptions, LPTSTR aEndKeys, LPTSTR aMatchList, size_t aMatchList_length)
{
	ParseOptions(aOptions);
	if (!SetKeyFlags(aEndKeys))
		return FAIL;
	if (!SetMatchList(aMatchList, aMatchList_length))
		return FAIL;
	
	// For maintainability/simplicity/code size, it's allocated even if BufferLengthMax == 0.
	if (  !(Buffer = tmalloc(BufferLengthMax + 1))  )
		return MemoryError();
	*Buffer = '\0';

	return OK;
}


void InputStart(input_type &input)
{
	ASSERT(!input.InProgress());

	// Keep the object alive while it is active, even if the script discards it.
	// The corresponding Release() is done when g_input is reset by InputRelease().
	if (input.ScriptObject)
		input.ScriptObject->AddRef();
	
	// Set or update the timeout timer if needed.  The timer proc takes care to end
	// only those inputs which are due, and will reset or kill the timer as needed.
	if (input.Timeout > 0)
		input.SetTimeoutTimer();

	input.Prev = g_input;
	input.Start();
	g_input = &input; // Signal the hook to start the input.

	Hotkey::InstallKeybdHook(); // Install the hook (if needed).
}


void input_type::ParseOptions(LPTSTR aOptions)
{
	for (LPTSTR cp = aOptions; *cp; ++cp)
	{
		switch(ctoupper(*cp))
		{
		case 'B':
			BackspaceIsUndo = false;
			break;
		case 'C':
			CaseSensitive = true;
			break;
		case 'I':
			MinSendLevel = (cp[1] <= '9' && cp[1] >= '0') ? (SendLevelType)_ttoi(cp + 1) : 1;
			break;
		case 'M':
			TranscribeModifiedKeys = true;
			break;
		case 'L':
			// Use atoi() vs. ATOI() to avoid interpreting something like 0x01C as hex
			// when in fact the C was meant to be an option letter:
			BufferLengthMax = _ttoi(cp + 1);
			if (BufferLengthMax < 0)
				BufferLengthMax = 0;
			break;
		case 'T':
			// Although ATOF() supports hex, it's been documented in the help file that hex should
			// not be used (see comment above) so if someone does it anyway, some option letters
			// might be misinterpreted:
			Timeout = (int)(ATOF(cp + 1) * 1000);
			break;
		case 'V':
			VisibleText = true;
			VisibleNonText = true;
			break;
		case '*':
			FindAnywhere = true;
			break;
		case 'E':
			// Interpret single-character keys as characters rather than converting them to VK codes.
			// This tends to work better when using multiple keyboard layouts, but changes behaviour:
			// for instance, an end char of "." cannot be triggered while holding Alt.
			EndCharMode = true;
			break;
		}
	}
}


void input_type::SetTimeoutTimer()
{
	DWORD now = GetTickCount();
	TimeoutAt = now + Timeout;
	if (!g_InputTimerExists || Timeout < int(g_InputTimeoutAt - now))
		SET_INPUT_TIMER(Timeout, TimeoutAt)
}


ResultType input_type::SetKeyFlags(LPTSTR aKeys, bool aEndKeyMode, UCHAR aFlagsRemove, UCHAR aFlagsAdd)
{
	bool vk_by_number, sc_by_number;
	vk_type vk;
	sc_type sc = 0;
	modLR_type modifiersLR;
	size_t key_text_length;
	UINT single_char_count = 0;
	TCHAR *end_pos, single_char_string[2];
	single_char_string[1] = '\0'; // Init its second character once, since the loop only changes the first char.
	
	const bool endchar_mode = aEndKeyMode && EndCharMode;
	UCHAR * const end_vk = KeyVK;
	UCHAR * const end_sc = KeySC;

	for (TCHAR *end_key = aKeys; *end_key; ++end_key) // This a modified version of the processing loop used in SendKeys().
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

			sc_by_number = false; // Set default.
			modifiersLR = 0;  // Init prior to below.
			// Handle the key by VK if it was given by number, such as {vk26}.
			// Otherwise, for any key name which has a VK shared by two possible SCs
			// (such as Up and NumpadUp), handle it by SC so it's identified correctly.
			if (vk = TextToVK(end_key + 1, &modifiersLR, true))
			{
				vk_by_number = ctoupper(end_key[1]) == 'V' && ctoupper(end_key[2]) == 'K';
				if (!vk_by_number && (sc = vk_to_sc(vk, true)))
				{
					sc ^= 0x100; // Convert sc to the primary scan code, which is the one named by end_key.
					vk = 0; // Handle it only by SC.
				}
			}
			else
				// No virtual key, so try to find a scan code.
				sc = TextToSC(end_key + 1, &sc_by_number);

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
			vk_by_number = false;
		} // switch()

		if (vk) // A valid virtual key code was discovered above.
		{
			// Insist the shift key be down to form genuinely different symbols --
			// namely punctuation marks -- but not for alphabetic chars.
			if (*single_char_string && aEndKeyMode && !IsCharAlpha(*single_char_string)) // v1.0.46.05: Added check for "*single_char_string" so that non-single-char strings like {F9} work as end keys even when the Shift key is being held down (this fixes the behavior to be like it was in pre-v1.0.45).
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
			else
			{
				end_vk[vk] = (end_vk[vk] & ~aFlagsRemove) | aFlagsAdd;
				// Apply flag removal to this key's SC as well.  This is primarily
				// to support combinations like {All} +E, {LCtrl}{RCtrl} -E.
				sc_type temp_sc;
				if (aFlagsRemove && !vk_by_number && (temp_sc = vk_to_sc(vk)))
				{
					end_sc[temp_sc] &= ~aFlagsRemove; // But apply aFlagsAdd only by VK.
					// Since aFlagsRemove implies ScriptObject != NULL and !vk_by_number
					// was also checked, that implies vk_to_sc(vk, true) was already called
					// and did not find a secondary SC.
				}
			}
		}
		if (sc || sc_by_number) // Fixed for v1.1.33.02: Allow sc000 for setting/unsetting flags for any events that lack a scan code.
		{
			end_sc[sc] = (end_sc[sc] & ~aFlagsRemove) | aFlagsAdd;
		}
	} // for()

	if (single_char_count)  // See single_char_count++ above for comments.
	{
		if (single_char_count > EndCharsMax)
		{
			// Allocate a bigger buffer.
			if (EndCharsMax) // If zero, EndChars may point to static memory.
				free(EndChars);
			if (  !(EndChars = tmalloc(single_char_count + 1))  )
				return MemoryError();
			EndCharsMax = single_char_count;
		}
		TCHAR *dst, *src;
		for (dst = EndChars, src = aKeys; *src; ++src)
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
		ASSERT(dst > EndChars);
		*dst = '\0';
	}
	else if (aEndKeyMode) // single_char_count is false
	{
		if (EndCharsMax)
			*EndChars = '\0';
		else
			EndChars = _T("");
	}
	return OK;
}


ResultType input_type::SetMatchList(LPTSTR aMatchList, size_t aMatchList_length)
{
	LPTSTR *realloc_temp;  // Needed since realloc returns NULL on failure but leaves original block allocated.
	MatchCount = 0;  // Set default.
	if (*aMatchList)
	{
		// If needed, create the array of pointers that points into MatchBuf to each match phrase:
		if (!match)
		{
			if (   !(match = (LPTSTR *)malloc(INPUT_ARRAY_BLOCK_SIZE * sizeof(LPTSTR)))   )
				return MemoryError();  // Short msg. since so rare.
			MatchCountMax = INPUT_ARRAY_BLOCK_SIZE;
		}
		// If needed, create or enlarge the buffer that contains all the match phrases:
		size_t space_needed = aMatchList_length + 1;  // +1 for the final zero terminator.
		if (space_needed > MatchBufSize)
		{
			MatchBufSize = (UINT)(space_needed > 4096 ? space_needed : 4096);
			if (MatchBuf) // free the old one since it's too small.
				free(MatchBuf);
			if (   !(MatchBuf = tmalloc(MatchBufSize))   )
			{
				MatchBufSize = 0;
				return MemoryError();  // Short msg. since so rare.
			}
		}
		// Copy aMatchList into the match buffer:
		LPTSTR source, dest;
		for (source = aMatchList, dest = match[MatchCount] = MatchBuf
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
			if (*match[MatchCount])
			{
				++MatchCount;
				match[MatchCount] = ++dest;
				*dest = '\0';  // Init to prevent crash on orphaned comma such as "btw,otoh,"
			}
			if (*(source + 1)) // There is a next element.
			{
				if (MatchCount >= MatchCountMax - 1) // Rarely needed, so just realloc() to expand.
				{
					// Expand the array by one block:
					if (   !(realloc_temp = (LPTSTR *)realloc(match  // Must use a temp variable.
						, (MatchCountMax + INPUT_ARRAY_BLOCK_SIZE) * sizeof(LPTSTR)))   )
						return MemoryError();  // Short msg. since so rare.
					match = realloc_temp;
					MatchCountMax += INPUT_ARRAY_BLOCK_SIZE;
				}
			}
		} // for()
		*dest = '\0';  // Terminate the last item.
		// This check is necessary for only a single isolated case: When the match list
		// consists of nothing except a single comma.  See above comment for details:
		if (*match[MatchCount]) // i.e. omit empty strings from the match list.
			++MatchCount;
	}
	return OK;
}


LPTSTR input_type::GetEndReason(LPTSTR aKeyBuf, int aKeyBufSize)
{
	switch (Status)
	{
	case INPUT_TIMED_OUT:
		return _T("Timeout");
	case INPUT_TERMINATED_BY_MATCH:
		return _T("Match");
	case INPUT_TERMINATED_BY_ENDKEY:
	{
		LPTSTR key_name = aKeyBuf;
		if (!key_name)
			return _T("EndKey");
		if (EndingChar)
		{
			key_name[0] = EndingChar;
			key_name[1] = '\0';
		}
		else if (EndingRequiredShift)
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
			BYTE state[256] = { 0 };
			state[VK_SHIFT] |= 0x80; // Indicate that the neutral shift key is down for conversion purposes.
			Get_active_window_keybd_layout // Defines the variable active_window_keybd_layout for use below.
			int count = ToUnicodeOrAsciiEx(EndingVK, vk_to_sc(EndingVK), (PBYTE)&state // Nothing is done about ToAsciiEx's dead key side-effects here because it seems to rare to be worth it (assuming its even a problem).
				, key_name, g_MenuIsVisible ? 1 : 0, active_window_keybd_layout); // v1.0.44.03: Changed to call ToAsciiEx() so that active window's layout can be specified (see hook.cpp for details).
			key_name[count] = '\0';  // Terminate the string.
		}
		else
		{
			*key_name = '\0';
			if (EndingBySC)
				SCtoKeyName(EndingSC, key_name, aKeyBufSize, false);
			if (!*key_name)
				VKtoKeyName(EndingVK, key_name, aKeyBufSize, !EndingBySC);
			if (!*key_name)
				sntprintf(key_name, aKeyBufSize, _T("sc%03X"), EndingSC);
		}
		return _T("EndKey");
	}
	case INPUT_LIMIT_REACHED:
		return _T("Max");
	case INPUT_OFF:
		return _T("Stopped");
	default: // In progress.
		return _T("");
	}
}


void input_type::Start()
{
	ASSERT(!InProgress());
	Status = INPUT_IN_PROGRESS;
}

void input_type::EndByMatch(UINT aMatchIndex)
{
	ASSERT(InProgress());
	EndingMatchIndex = aMatchIndex;
	EndByReason(INPUT_TERMINATED_BY_MATCH);
}

void input_type::EndByKey(vk_type aVK, sc_type aSC, bool aBySC, bool aRequiredShift)
{
	ASSERT(InProgress());
	EndingVK = aVK;
	EndingSC = aSC;
	EndingBySC = aBySC;
	EndingRequiredShift = aRequiredShift;
	EndingChar = 0; // Must be zero if the above are to be used.
	EndByReason(INPUT_TERMINATED_BY_ENDKEY);
}

void input_type::EndByChar(TCHAR aChar)
{
	ASSERT(aChar && InProgress());
	EndingChar = aChar;
	// The other EndKey related fields are ignored when Char is non-zero.
	EndByReason(INPUT_TERMINATED_BY_ENDKEY);
}

void input_type::EndByReason(InputStatusType aReason)
{
	ASSERT(InProgress());
	EndingMods = g_modifiersLR_logical; // Not relevant to all end reasons, but might be useful anyway.
	Status = aReason;

	// It's done this way rather than calling InputRelease() directly...
	// ...so that we can rely on MsgSleep() to create a new thread for the OnEnd event.
	// ...because InputRelease() can't be called by the hook thread.
	// ...because some callers rely on the list not being broken by this call.
	PostMessage(g_hWnd, AHK_INPUT_END, (WPARAM)this, 0);
}


input_type *InputRelease(input_type *aInput)
{
	if (!aInput)
		return NULL;
	// Input should already have ended prior to this function being called.
	// Otherwise, removal of aInput from the chain will end input collection.
	if (g_input == aInput)
		g_input = aInput->Prev;
	else
		for (auto *input = g_input; ; input = input->Prev)
		{
			if (!input)
				return NULL; // aInput is not valid (faked AHK_INPUT_END message?) or not active.
			if (input->Prev == aInput)
			{
				input->Prev = aInput->Prev;
				break;
			}
		}

	// Ensure any pending use of aInput by the hook is finished.
	WaitHookIdle();
	
	aInput->Prev = NULL;
	if (aInput->ScriptObject)
	{
		Hotkey::MaybeUninstallHook();
		if (aInput->ScriptObject->onEnd)
			return aInput; // Return for caller to call OnEnd and Release.
		aInput->ScriptObject->Release();
		// The following is not done because this Release() is only to counteract an AddRef() in
		// InputStart().  ScriptObject != NULL indicates this input_type is actually embedded in
		// the InputObject and as such the link should never be broken until both are deleted.
		//aInput->ScriptObject = NULL;
		g_script.ExitIfNotPersistent(EXIT_EXIT); // In case this InputHook was the only thing keeping the script running.
	}
	return NULL;
}


input_type *InputFind(InputObject *object)
{
	for (auto *input = g_input; input; input = input->Prev)
		if (input->ScriptObject == object)
			return input;
	return NULL;
}



BIF_DECL(BIF_WinShow)
{
	auto action = _f_callee_id;

	_f_set_retval_p(_T(""), 0);

	_f_param_string_opt(aTitle, 0);
	_f_param_string_opt(aText, 1);
	// The remaining parameters depend on which function this is.

	// Set initial guess for is_ahk_group (further refined later).  For ahk_group, WinText,
	// ExcludeTitle, and ExcludeText must be blank so that they are reserved for future use
	// (i.e. they're currently not supported since the group's own criteria take precedence):
	bool is_ahk_group = !_tcsnicmp(aTitle, _T("ahk_group"), 9)
		&& ParamIndexIsOmittedOrEmpty(1) && ParamIndexIsOmittedOrEmpty(3);
	// The following is not quite accurate since is_ahk_group is only a guess at this stage, but
	// given the extreme rarity of the guess being wrong, this shortcut seems justified to reduce
	// the code size/complexity.  A wait_time of zero seems best for group closing because it's
	// currently implemented to do the wait after every window in the group.  In addition,
	// this makes "WinClose ahk_group GroupName" behave identically to "GroupClose GroupName",
	// which seems best, for consistency:
	int wait_time = is_ahk_group ? 0 : DEFAULT_WINCLOSE_WAIT;
	if (action == FID_WinClose || action == FID_WinKill) // aParam[2] contains the wait time.
	{
		if (!ParamIndexIsOmittedOrEmpty(2))
			wait_time = (int)(1000 * ParamIndexToDouble(2));
		if (!ParamIndexIsOmittedOrEmpty(4))
			is_ahk_group = false;  // Override the default.
	}
	else
		if (!ParamIndexIsOmittedOrEmpty(2))
			is_ahk_group = false;  // Override the default.
	if (is_ahk_group)
		if (WinGroup *group = g_script.FindGroup(omit_leading_whitespace(aTitle + 9)))
		{
			group->ActUponAll(action, wait_time); // It will do DoWinDelay if appropriate.
			_f_return_retval;
		}
	// Since above didn't return, either the group doesn't exist or it's paired with other
	// criteria, such as "ahk_group G ahk_class C", so do the normal single-window behavior.

	HWND target_window = NULL;
	if (aParamCount > 0)
	{
		switch (DetermineTargetHwnd(target_window, aResultToken, *aParam[0]))
		{
		case FAIL: return;
		case OK:
			if (!target_window) // Specified a HWND of 0, or IsWindow() returned false.
				_f_throw(ERR_NO_WINDOW, ErrorPrototype::Target);
		}
	}

	if (action == FID_WinClose || action == FID_WinKill)
	{
		if (target_window)
		{
			WinClose(target_window, wait_time, action == FID_WinKill);
			DoWinDelay;
			_f_return_retval;
		}
		_f_param_string_opt(aExcludeTitle, 3);
		_f_param_string_opt(aExcludeText, 4);
		if (!WinClose(*g, aTitle, aText, wait_time, aExcludeTitle, aExcludeText, action == FID_WinKill))
			// Currently WinClose returns NULL only for this case; it doesn't confirm the window closed.
			_f_throw(ERR_NO_WINDOW, ErrorPrototype::Target);
		DoWinDelay;
		_f_return_retval;
	}

	if (!target_window)
	{
		_f_param_string_opt(aExcludeTitle, 2);
		_f_param_string_opt(aExcludeText, 3);
		// By design, the WinShow command must always unhide a hidden window, even if the user has
		// specified that hidden windows should not be detected.  So set this now so that
		// DetermineTargetWindow() will make its calls in the right mode:
		bool need_restore = (_f_callee_id == FID_WinShow && !g->DetectHiddenWindows);
		if (need_restore)
			g->DetectHiddenWindows = true;
		target_window = Line::DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
		if (need_restore)
			g->DetectHiddenWindows = false;
		if (!target_window)
			_f_throw(ERR_NO_WINDOW, ErrorPrototype::Target);
	}

	// WinGroup's EnumParentActUponAll() is quite similar to the following, so the two should be
	// maintained together.

	int nCmdShow = SW_NONE; // Set default.

	switch (action)
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
	case FID_WinMinimize:
		if (IsWindowHung(target_window))
		{
			nCmdShow = SW_FORCEMINIMIZE;
			// SW_MINIMIZE can lock up our thread on WinXP, which is why we revert to SW_FORCEMINIMIZE above.
			// Older/obsolete comment for background: don't attempt to minimize hung windows because that
			// might hang our thread because the call to ShowWindow() would never return.
		}
		else
			nCmdShow = SW_MINIMIZE;
		break;
	case FID_WinMaximize: if (!IsWindowHung(target_window)) nCmdShow = SW_MAXIMIZE; break;
	case FID_WinRestore:  if (!IsWindowHung(target_window)) nCmdShow = SW_RESTORE;  break;
	// Seems safe to assume it's not hung in these cases, since I'm inclined to believe
	// (untested) that hiding and showing a hung window won't lock up our thread, and
	// there's a chance they may be effective even against hung windows, unlike the
	// others above (except ACT_WINMINIMIZE, which has a special FORCE method):
	case FID_WinHide: nCmdShow = SW_HIDE; break;
	case FID_WinShow: nCmdShow = SW_SHOW; break;
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
		DoWinDelay;
	}
	_f_return_retval;
}



BIF_DECL(BIF_WinActivate)
{
	_f_set_retval_p(_T(""), 0);

	if (aParamCount > 0)
	{
		HWND target_hwnd;
		switch (DetermineTargetHwnd(target_hwnd, aResultToken, *aParam[0]))
		{
		case FAIL: return;
		case OK:
			if (!target_hwnd)
				_f_throw(ERR_NO_WINDOW, ErrorPrototype::Target);
			SetForegroundWindowEx(target_hwnd);
			DoWinDelay;
			_f_return_retval;
		}
	}

	_f_param_string_opt(aTitle, 0);
	_f_param_string_opt(aText, 1);
	_f_param_string_opt(aExcludeTitle, 2);
	_f_param_string_opt(aExcludeText, 3);

	if (!WinActivate(*g, aTitle, aText, aExcludeTitle, aExcludeText, _f_callee_id == FID_WinActivateBottom, true))
		_f_throw(ERR_NO_WINDOW, ErrorPrototype::Target);

	// It seems best to do these sleeps here rather than in the windowing
	// functions themselves because that way, the program can use the
	// windowing functions without being subject to the script's delay
	// setting (i.e. there are probably cases when we don't need to wait,
	// such as bringing a message box to the foreground, since no other
	// actions will be dependent on it actually having happened):
	DoWinDelay;
	_f_return_retval;
}



BIF_DECL(BIF_GroupActivate)
{
	LPTSTR aGroup = ParamIndexToOptionalString(0, _f_number_buf);
	WinGroup *group;
	if (   !(group = g_script.FindGroup(aGroup, true))   ) // Last parameter -> create-if-not-found.
		_f_return_FAIL;  // It already displayed the error for us.
	
	LPTSTR aMode = ParamIndexToOptionalString(1, _f_number_buf);
	bool reverse = false;
	if (!_tcsicmp(aMode, _T("R")))
		reverse = true;
	else if (*aMode)
		_f_throw_param(1);

	HWND activated;
	if (!group->Activate(reverse, activated))
		_f_return_FAIL;
	_f_return_i((UINT_PTR)activated);
}



BIF_DECL(BIF_Wait)
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

	// For FID_KeyWait:
	bool wait_for_keydown;
	KeyStateTypes key_state_type;
	JoyControls joy;
	int joystick_id;
	ExprTokenType token;

	Line *waiting_line = g_script.mCurrLine;

	_f_set_retval_i(TRUE); // Set default return value to be possibly overridden later on.

	_f_param_string_opt(arg1, 0);
	_f_param_string_opt(arg2, 1);
	_f_param_string_opt(arg3, 2);

	HWND target_hwnd = NULL;

	switch (_f_callee_id)
	{
	case FID_RunWait:
		if (!g_script.ActionExec(arg1, NULL, arg2, true, arg3, &running_process, true, true
			, ParamIndexToOutputVar(4-1)))
			_f_return_FAIL;
		//else fall through to the waiting-phase of the operation.
		break;
	case FID_WinWait:
	case FID_WinWaitClose:
	case FID_WinWaitActive:
	case FID_WinWaitNotActive:
		if (ParamIndexIsOmitted(0))
			break;
		// The following supports pure HWND or {Hwnd:HWND} as in other WinTitle parameters,
		// in lieu of saving aParam[] and evaluating them vs. arg1,2,4,5 on each iteration.
		switch (DetermineTargetHwnd(target_hwnd, aResultToken, *aParam[0]))
		{
		case FAIL: return;
		case OK:
			if (!target_hwnd) // Caller passed 0 or a value that failed an IsWindow() check.
			{
				// If it's 0 due to failing IsWindow(), there's no way to determine whether
				// it was ever valid, so assume it was but that the window has been closed.
				DoWinDelay;
				if (_f_callee_id == FID_WinWaitClose || _f_callee_id == FID_WinWaitNotActive)
					_f_return_retval;
				// Otherwise, this window will never exist or be active again, so abort early.
				// It might be more correct to just wait the timeout (or stall indefinitely),
				// but that's probably not the user's intention.
				_f_return_i(FALSE);
			}
		}
		break;
	}

	// These are declared/accessed after "case FID_RunWait:" to avoid a UseUnset warning.
	_f_param_string_opt(arg4, 3);
	_f_param_string_opt(arg5, 4);
	
	// Must NOT use ELSE-IF in line below due to ELSE further down needing to execute for RunWait.
	if (_f_callee_id == FID_KeyWait)
	{
		if (   !(vk = TextToVK(arg1))   )
		{
			joy = (JoyControls)ConvertJoy(arg1, &joystick_id);
			if (!IS_JOYSTICK_BUTTON(joy)) // Currently, only buttons are supported.
				// It's either an invalid key name or an unsupported Joy-something.
				_f_throw_param(0);
		}
		// Set defaults:
		wait_for_keydown = false;  // The default is to wait for the key to be released.
		key_state_type = KEYSTATE_PHYSICAL;  // Since physical is more often used.
		wait_indefinitely = true;
		sleep_duration = 0;
		for (LPTSTR cp = arg2; *cp; ++cp)
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
	}
	else if (   (_f_callee_id != FID_RunWait && _f_callee_id != FID_ClipWait && *arg3)
		|| (_f_callee_id == FID_ClipWait && *arg1)   )
	{
		// Since the param containing the timeout value isn't blank, it must be numeric,
		// otherwise, the loading validation would have prevented the script from loading.
		wait_indefinitely = false;
		sleep_duration = (int)(ATOF(_f_callee_id == FID_ClipWait ? arg1 : arg3) * 1000); // Can be zero.
	}
	else
	{
		wait_indefinitely = true;
		sleep_duration = 0; // Just to catch any bugs.
	}

	bool any_clipboard_format = (_f_callee_id == FID_ClipWait && ATOI(arg2) == 1);

	for (start_time = GetTickCount();;) // start_time is initialized unconditionally for use with v1.0.30.02's new logging feature further below.
	{ // Always do the first iteration so that at least one check is done.
		if (target_hwnd) // Caller passed a pure HWND or {Hwnd:HWND}.
		{
			// Change the behaviour a little since we know that once the HWND is destroyed,
			// it is not meaningful to wait for another window with that same HWND.
			if (!IsWindow(target_hwnd))
			{
				DoWinDelay;
				if (_f_callee_id == FID_WinWaitClose || _f_callee_id == FID_WinWaitNotActive)
					_f_return_retval; // Condition met.
				// Otherwise, it would not be meaningful to wait for another window to be
				// created with the same HWND.  It seems more useful to abort immediately
				// but report timeout/failure than to wait for the timeout to elapse.
				_f_return_i(FALSE);
			}
			if (_f_callee_id == FID_WinWait || _f_callee_id == FID_WinWaitClose)
			{
				// Wait for the window to become visible/hidden.  Most functions ignore
				// DetectHiddenWindows when given a pure HWND/object (because it's more
				// useful that way), but in this case it seems more useful and intuitive
				// to respect DetectHiddenWindows.
				if (g->DetectWindow(target_hwnd) == (_f_callee_id == FID_WinWait))
				{
					DoWinDelay;
					if (_f_callee_id == FID_WinWaitClose)
						_f_return_retval;
					_f_return_i((size_t)target_hwnd);
				}
			}
			else
			{
				if ((GetForegroundWindow() == target_hwnd) == (_f_callee_id == FID_WinWaitActive))
				{
					DoWinDelay;
					if (_f_callee_id == FID_WinWaitNotActive)
						_f_return_retval;
					_f_return_i((size_t)target_hwnd);
				}
			}
		}
		else switch (_f_callee_id)
		{
		case FID_WinWait:
			#define SAVED_WIN_ARGS arg1, arg2, arg4, arg5
			if (HWND found = WinExist(*g, SAVED_WIN_ARGS, false, true))
			{
				DoWinDelay;
				_f_return_i((size_t)found);
			}
			break;
		case FID_WinWaitClose:
			if (!WinExist(*g, SAVED_WIN_ARGS, false, true))
			{
				DoWinDelay;
				_f_return_retval;
			}
			break;
		case FID_WinWaitActive:
			if (HWND found = WinActive(*g, SAVED_WIN_ARGS, true))
			{
				DoWinDelay;
				_f_return_i((size_t)found);
			}
			break;
		case FID_WinWaitNotActive:
			if (!WinActive(*g, SAVED_WIN_ARGS, true))
			{
				DoWinDelay;
				_f_return_retval;
			}
			break;
		case FID_ClipWait:
			// Seems best to consider CF_HDROP to be a non-empty clipboard, since we
			// support the implicit conversion of that format to text:
			if (any_clipboard_format)
			{
				if (CountClipboardFormats())
					_f_return_retval;
			}
			else
				if (IsClipboardFormatAvailable(CF_NATIVETEXT) || IsClipboardFormatAvailable(CF_HDROP))
					_f_return_retval;
			break;
		case FID_KeyWait:
			if (vk) // Waiting for key or mouse button, not joystick.
			{
				if (ScriptGetKeyState(vk, key_state_type) == wait_for_keydown)
					_f_return_retval;
			}
			else // Waiting for joystick button
			{
				TCHAR unused[32];
				if (ScriptGetJoyState(joy, joystick_id, token, unused) == wait_for_keydown)
					_f_return_retval;
			}
			break;
		case FID_RunWait:
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
				_f_return_i((int)exit_code);
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
					int previous_log_index = (Line::sLogNext ? Line::sLogNext : LINE_LOG_SIZE) - 1; // Wrap around if needed (the entry can be NULL in that case).
					if (Line::sLog[previous_log_index] != waiting_line || Line::sLogTick[previous_log_index] != start_time) // The previously logged line was not this one, or it was added by the interrupting thread (different start_time).
					{
						Line::sLog[Line::sLogNext] = waiting_line;
						Line::sLogTick[Line::sLogNext++] = start_time; // Store a special value so that Line::LogToText() can report that its "still waiting" from earlier.
						if (Line::sLogNext >= LINE_LOG_SIZE)
							Line::sLogNext = 0;
						// The lines above are the similar to those used in ExecUntil(), so the two should be
						// maintained together.
					}
				}
			}
		}
		else // Done waiting.
			_f_return_i(FALSE); // Since it timed out, we override the default with this.
	} // for()
}



BIF_DECL(BIF_WinMove)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam + 4, aParamCount - 4))
		return;
	RECT rect;
	if (!GetWindowRect(target_window, &rect)
		|| !MoveWindow(target_window
			, ParamIndexToOptionalInt(0, rect.left) // X-position
			, ParamIndexToOptionalInt(1, rect.top)  // Y-position
			, ParamIndexToOptionalInt(2, rect.right - rect.left)
			, ParamIndexToOptionalInt(3, rect.bottom - rect.top)
			, TRUE)) // Do repaint.
		_f_throw_win32();
	DoWinDelay;
	_f_return_empty;
}



BIF_DECL(BIF_ControlSend) // ControlSend and ControlSendText.
{
	DETERMINE_TARGET_CONTROL(1);

	_f_param_string(aKeysToSend, 0);
	SendKeys(aKeysToSend, (SendRawModes)_f_callee_id, SM_EVENT, control_window);
	// But don't do WinDelay because KeyDelay should have been in effect for the above.
	_f_return_empty;
}



BIF_DECL(BIF_ControlClick)
{
	_f_param_string_opt(aControl, 0);
	_f_param_string_opt(aWhichButton, 3);
	int aVK = Line::ConvertMouseButton(aWhichButton);
	if (!aVK)
		_f_throw_param(3);
	int aClickCount = ParamIndexToOptionalInt(4, 1);
	_f_param_string_opt(aOptions, 5);

	// Set the defaults that will be in effect unless overridden by options:
	KeyEventTypes event_type = KEYDOWNANDUP;
	bool position_mode = false;
	bool do_activate = true;
	// These default coords can be overridden either by aOptions or aControl's X/Y mode:
	POINT click = {COORD_UNSPECIFIED, COORD_UNSPECIFIED};

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
	
	HWND target_window, control_window;
	if (position_mode)
	{
		// Determine target window only.  Control will be found by position below.
		if (!DetermineTargetWindow(target_window, aResultToken, aParam + 1, aParamCount - 1, 3))
			return; // aResultToken.SetExitResult() or Error() was already called.
		control_window = NULL;
	}
	else
	{
		// Determine target window and control.
		if (!DetermineTargetControl(control_window, target_window, aResultToken, aParam, aParamCount, 3, false))
			return; // aResultToken.SetExitResult() or Error() was already called.
	}
	ASSERT(target_window != NULL);

	// It's debatable, but might be best for flexibility (and backward compatibility) to allow target_window to itself
	// be a control (at least for the position_mode handler below).  For example, the script may have called SetParent
	// to make a top-level window the child of some other window, in which case this policy allows it to be seen like
	// a non-child.
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
			goto control_error;
		++cp;
		if (!*cp)
			goto control_error;
		pah.pt.x = ATOI(cp);
		if (   !(cp = StrChrAny(cp, _T(" \t")))   ) // Find next space or tab (there must be one for it to be considered valid).
			goto control_error;
		cp = omit_leading_whitespace(cp + 1);
		if (!*cp || _totupper(*cp) != 'Y')
			goto control_error;
		++cp;
		if (!*cp)
			goto control_error;
		pah.pt.y = ATOI(cp);
		// The passed-in coordinates are always relative to target_window's client area because offering
		// an option for absolute/screen coordinates doesn't seem useful.
		ClientToScreen(target_window, &pah.pt); // Convert to screen coordinates.
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

	// This is done this late because it seems better to throw an exception whenever the
	// target window or control isn't found, or any other error condition occurs above:
	if (aClickCount < 1)
	{
		// Allow this to simply "do nothing", because it increases flexibility
		// in the case where the number of clicks is a dereferenced script variable
		// that may sometimes (by intent) resolve to zero or negative:
		_f_return_empty;
	}

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

	UINT msg_down, msg_up;
	WPARAM wparam, wparam_up = 0;
	bool vk_is_wheel = aVK == VK_WHEEL_UP || aVK == VK_WHEEL_DOWN;
	bool vk_is_hwheel = aVK == VK_WHEEL_LEFT || aVK == VK_WHEEL_RIGHT; // v1.0.48: Lexikos: Support horizontal scrolling in Windows Vista and later.

	if (vk_is_wheel)
	{
		ClientToScreen(control_window, &click); // Wheel messages use screen coordinates.
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
			default: // Just do nothing since this should realistically never happen.
				ASSERT(!"aVK value not handled");
		}
	}

	LPARAM lparam = MAKELPARAM(click.x, click.y);

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

	_f_return_empty;

error:
	_f_throw_win32();

control_error:
	_f_throw(ERR_NO_CONTROL, aControl, ErrorPrototype::Target);
}



BIF_DECL(BIF_ControlMove)
{
	DETERMINE_TARGET_CONTROL(4);
	
	// The following macro is used to keep ControlMove and ControlGetPos in sync:
	#define CONTROL_COORD_PARENT(target, control) \
		(target == control ? GetNonChildParent(target) : target)

	// Determine which window the supplied coordinates are relative to:
	HWND coord_parent = CONTROL_COORD_PARENT(target_window, control_window);

	// Determine the controls current coordinates relative to coord_parent in case one
	// or more parameters were omitted.
	RECT control_rect;
	if (!GetWindowRect(control_window, &control_rect)
		|| !MapWindowPoints(NULL, coord_parent, (LPPOINT)&control_rect, 2))
		_f_throw_win32();
	
	POINT point;
	point.x = ParamIndexToOptionalInt(0, control_rect.left);
	point.y = ParamIndexToOptionalInt(1, control_rect.top);

	// MoveWindow accepts coordinates relative to the control's immediate parent, which might
	// be different to coord_parent since controls can themselves have child controls.  So if
	// necessary, map the caller-supplied coordinates to the control's immediate parent:
	HWND immediate_parent = GetParent(control_window);
	if (immediate_parent != coord_parent)
		MapWindowPoints(coord_parent, immediate_parent, &point, 1);

	MoveWindow(control_window
		, point.x
		, point.y
		, ParamIndexToOptionalInt(2, control_rect.right - control_rect.left)
		, ParamIndexToOptionalInt(3, control_rect.bottom - control_rect.top)
		, TRUE);  // Do repaint.

	DoControlDelay
	_f_return_empty;
}



BIF_DECL(BIF_ControlGetPos)
{
	Var *output_var_x = ParamIndexToOutputVar(0);
	Var *output_var_y = ParamIndexToOutputVar(1);
	Var *output_var_width = ParamIndexToOutputVar(2);
	Var *output_var_height = ParamIndexToOutputVar(3);

	DETERMINE_TARGET_CONTROL(4);

	// Determine which window the returned coordinates should be relative to:
	HWND coord_parent = CONTROL_COORD_PARENT(target_window, control_window);

	RECT child_rect;
	// Realistically never fails since DetermineTargetWindow() and ControlExist() should always yield
	// valid window handles:
	GetWindowRect(control_window, &child_rect);
	// Map the screen coordinates returned by GetWindowRect to the client area of coord_parent.
	MapWindowPoints(NULL, coord_parent, (LPPOINT)&child_rect, 2);

	output_var_x && output_var_x->Assign(child_rect.left);
	output_var_y && output_var_y->Assign(child_rect.top);
	output_var_width && output_var_width->Assign(child_rect.right - child_rect.left);
	output_var_height && output_var_height->Assign(child_rect.bottom - child_rect.top);
	_f_return_empty;
}



BIF_DECL(BIF_ControlGetFocus)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam, aParamCount))
		return;

	GUITHREADINFO guithreadInfo;
	guithreadInfo.cbSize = sizeof(GUITHREADINFO);
	if (!GetGUIThreadInfo(GetWindowThreadProcessId(target_window, NULL), &guithreadInfo))
		_f_throw_win32();

	// Use IsChild() to ensure the focused control actually belongs to this window.
	// Otherwise, a HWND will be returned if any window in the same thread has focus,
	// including the target window itself (typically when it has no controls).
	if (!IsChild(target_window, guithreadInfo.hwndFocus))
		_f_return_i(0); // As documented, if "none of the target window's controls has focus, the return value is 0".
	_f_return_i((UINT_PTR)guithreadInfo.hwndFocus);
}



BIF_DECL(BIF_ControlGetClassNN)
{
	DETERMINE_TARGET_CONTROL(0);

	if (target_window == control_window)
		target_window = GetNonChildParent(control_window);

	class_and_hwnd_type cah;
	TCHAR class_name[WINDOW_CLASS_SIZE];
	cah.hwnd = control_window;
	cah.class_name = class_name;
	if (!GetClassName(cah.hwnd, class_name, _countof(class_name) - 5)) // -5 to allow room for sequence number.
		_f_throw_win32();
	
	cah.class_count = 0;  // Init for the below.
	cah.is_found = false; // Same.
	EnumChildWindows(target_window, EnumChildFindSeqNum, (LPARAM)&cah);
	if (!cah.is_found)
		_f_throw(ERR_FAILED);
	// Append the class sequence number onto the class name:
	sntprintfcat(class_name, _countof(class_name), _T("%d"), cah.class_count);
	_f_return(class_name);
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



BIF_DECL(BIF_ControlFocus)
{
	DETERMINE_TARGET_CONTROL(0);

	// Unlike many of the other Control commands, this one requires AttachThreadInput()
	// to have any realistic chance of success (though sometimes it may work by pure
	// chance even without it):
	ATTACH_THREAD_INPUT

	SetFocus(control_window);
	DoControlDelay; // Done unconditionally for simplicity, and in case SetFocus() had some effect despite indicating failure.
	// GetFocus() isn't called and failure to focus isn't treated as an error because
	// a successful change in focus doesn't guarantee that the focus will still be as
	// expected when the next line of code runs.

	// Very important to detach any threads whose inputs were attached above,
	// prior to returning, otherwise the next attempt to attach thread inputs
	// for these particular windows may result in a hung thread or other
	// undesirable effect:
	DETACH_THREAD_INPUT

	_f_return_empty;
}



BIF_DECL(BIF_ControlSetText)
{
	DETERMINE_TARGET_CONTROL(1);

	_f_param_string(aNewText, 0);
	// SendMessage must be used, not PostMessage(), at least for some (probably most) apps.
	// Also: No need to call IsWindowHung() because SendMessageTimeout() should return
	// immediately if the OS already "knows" the window is hung:
	DWORD_PTR result;
	SendMessageTimeout(control_window, WM_SETTEXT, (WPARAM)0, (LPARAM)aNewText
		, SMTO_ABORTIFHUNG, 5000, &result);
	DoControlDelay;
	_f_return_empty;
}



BIF_DECL(BIF_ControlGetText)
{
	DETERMINE_TARGET_CONTROL(0);

	// Even if control_window is NULL, we want to continue on so that the output
	// param is set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before.

	// Handle the output parameter.  Note: Using GetWindowTextTimeout() vs. GetWindowText()
	// because it is able to get text from more types of controls (e.g. large edit controls):
	VarSizeType space_needed = GetWindowTextTimeout(control_window) + 1; // 1 for terminator.

	// Allocate memory for the return value.
	if (!TokenSetResult(aResultToken, NULL, space_needed - 1))
		return;  // It already displayed the error.
	aResultToken.symbol = SYM_STRING;
	// Fetch the text directly into the buffer.  Also set the length explicitly
	// in case actual size written was off from the estimated size (since
	// GetWindowTextLength() can return more space that will actually be required
	// in certain circumstances, see MS docs):
	aResultToken.marker_length = GetWindowTextTimeout(control_window, aResultToken.marker, space_needed);
	if (!aResultToken.marker_length) // There was no text to get or GetWindowTextTimeout() failed.
		*aResultToken.marker = '\0';
}



void ControlGetListView(ResultToken &aResultToken, HWND aHwnd, LPTSTR aOptions)
// Called by ControlGet() below.  It has ensured that aHwnd is a valid handle to a ListView.
{
	// GET ROW COUNT
	LRESULT row_count;
	if (!SendMessageTimeout(aHwnd, LVM_GETITEMCOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&row_count)) // Timed out or failed.
		goto error;

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
	if (col_count > -1 && col_option && (requested_col < 0 || requested_col >= col_count)) // Specified column does not exist.
		_f_throw_value(ERR_PARAM1_INVALID, col_option);

	// IF THE "COUNT" OPTION IS PRESENT, FULLY HANDLE THAT AND RETURN
	if (get_count)
	{
		int result; // Must be signed to support writing a col count of -1 to aOutputVar.
		if (include_focused_only) // Listed first so that it takes precedence over include_selected_only.
		{
			if (!SendMessageTimeout(aHwnd, LVM_GETNEXTITEM, -1, LVNI_FOCUSED, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&result)) // Timed out or failed.
				goto error;
			++result; // i.e. Set it to 0 if not found, or the 1-based row-number otherwise.
		}
		else if (include_selected_only)
		{
			if (!SendMessageTimeout(aHwnd, LVM_GETSELECTEDCOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&result)) // Timed out or failed.
				goto error;
		}
		else if (col_option) // "Count Col" returns the number of columns.
			result = (int)col_count;
		else // Total row count.
			result = (int)row_count;
		_f_return(result);
	}

	// FINAL CHECKS
	if (row_count < 1 || !col_count) // But don't return when col_count == -1 (i.e. always make the attempt when col count is undetermined).
		_f_return_empty;  // No text in the control, so indicate success.
	
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
		goto error;
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

	// SET UP THE OUTPUT BUFFER
	if (!TokenSetResult(aResultToken, NULL, (size_t)total_length))
		goto cleanup_and_return; // Error() was already called.
	aResultToken.symbol = SYM_STRING;
	
	LPTSTR contents = aResultToken.marker;
	LRESULT capacity = total_length; // LRESULT avoids signed vs. unsigned compiler warnings.
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
	*contents = '\0'; // Final termination.  Above has reserved room for this one byte.
	aResultToken.marker_length = (size_t)total_length; // Update to actual vs. estimated length.

	// CLEAN UP
cleanup_and_return: // This is "called" if a memory allocation failed above
	FreeInterProcMem(handle, p_remote_lvi);
	return;

error:
	_f_throw_win32();
}



bool ControlSetTab(ResultToken &aResultToken, HWND aHwnd, DWORD aTabIndex)
{
	DWORD_PTR dwResult;
	// MSDN: "If the tab control does not have the TCS_BUTTONS style, changing the focus also changes
	// the selected tab. In this case, the tab control sends the TCN_SELCHANGING and TCN_SELCHANGE
	// notification codes to its parent window."
	if (!SendMessageTimeout(aHwnd, TCM_SETCURFOCUS, aTabIndex, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
		return false;
	// Tab controls with the TCS_BUTTONS style need additional work:
	if (GetWindowLong(aHwnd, GWL_STYLE) & TCS_BUTTONS)
	{
		// Problem:
		//  TCM_SETCURFOCUS does not change the selected tab if TCS_BUTTONS is set.
		//
		// False solution #1 (which used to be recommended in the docs):
		//  Send a TCM_SETCURSEL method afterward.  TCM_SETCURSEL changes the selected tab,
		//  but doesn't notify the control's parent, so it doesn't update the tab's contents.
		//
		// False solution #2:
		//  Send a WM_NOTIFY message to the parent window to notify it.  Can't be done.
		//  MSDN says: "For Windows 2000 and later systems, the WM_NOTIFY message cannot
		//  be sent between processes."
		//
		// Solution #1:
		//  Send VK_LEFT/VK_RIGHT as many times as needed.
		//
		// Solution #2:
		//  Set the focus to an adjacent tab and then send VK_LEFT/VK_RIGHT.
		//   - Must choose an appropriate tab index and vk depending on which tab is being
		//     selected, since VK_LEFT/VK_RIGHT don't wrap around.
		//   - Ends up tempting optimisations which increase code size, such as to avoid
		//     TCM_SETCURFOCUS if an adjacent tab is already focused.
		//   - Still needs VK_SPACE afterward to actually select the tab.
		//
		// Solution #3 (the one below):
		//  Set the focus to the appropriate tab and then send VK_SPACE.
		//   - Since we've already set the focus, all we need to do is send VK_SPACE.
		//   - If the tab index is invalid and the user has focused but not selected
		//     another tab, that tab will be selected.  This seems harmless enough.
		//
		PostMessage(aHwnd, WM_KEYDOWN, VK_SPACE, 0x00000001);
		PostMessage(aHwnd, WM_KEYUP, VK_SPACE, 0xC0000001);
	}
	return true;
}



BIF_DECL(BIF_StatusBarGetText)//(LPTSTR aPart, LPTSTR aTitle, LPTSTR aText
	//, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	int part = ParamIndexToOptionalInt(0, 1);
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam + 1, aParamCount - 1))
		return;
	HWND control_window = ControlExist(target_window, _T("msctls_statusbar321"));
	// StatusBarUtil will handle any NULL control_window or zero part# for us.
	StatusBarUtil(aResultToken, control_window, part);
}



BIF_DECL(BIF_StatusBarWait)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam + 3, aParamCount - 3, 1))
		return;

	LPTSTR aTextToWaitFor = ParamIndexToOptionalString(0, _f_number_buf);
	int aSeconds = ParamIndexIsOmittedOrEmpty(1) ? -1 : int(ParamIndexToDouble(1) * 1000);
	int aPart = ParamIndexToOptionalInt(2, 0);
	int aInterval = ParamIndexToOptionalInt(5, 50);

	// Make a copy of any memory areas that are volatile (due to caller passing a variable,
	// which could be reassigned by a new hotkey subroutine launched while we are waiting)
	// but whose contents we need to refer to while we are waiting:
	TCHAR text_to_wait_for[4096];
	tcslcpy(text_to_wait_for, aTextToWaitFor, _countof(text_to_wait_for));
	HWND control_window = ControlExist(target_window, _T("msctls_statusbar321"));
	// StatusBarUtil will handle any NULL control_window or zero part# for us.
	StatusBarUtil(aResultToken, control_window, aPart, text_to_wait_for, aSeconds, aInterval);
}



BIF_DECL(BIF_PostSendMessage)
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
	bool aUseSend = _f_callee_id == FID_SendMessage;
	bool successful = false;

	DETERMINE_TARGET_CONTROL(3);

	UINT msg = ParamIndexToInt(0);
	// Timeout increased from 2000 to 5000 in v1.0.27:
	// jackieku: specify timeout by the parameter.
	UINT timeout = ParamIndexToOptionalInt(8, 5000);

	// Fixed for v1.0.48.04: Make copies of the wParam and lParam variables (if eligible for updating) prior
	// to sending the message in case the message triggers a callback or OnMessage function, which would be
	// likely to change the contents of the mArg array before we're doing using them after the Post/SendMsg.
	// Seems best to do the above EVEN for PostMessage in case it can ever trigger a SendMessage internally
	// (I seem to remember that the OS sometimes converts a PostMessage call into a SendMessage if the
	// origin and destination are the same thread.)
	// v1.0.43.06: If either wParam or lParam contained the address of a variable, update the mLength
	// member after sending the message in case the receiver of the message wrote something to the buffer.
	// This is similar to the way "Str" parameters work in DllCall.
	INT_PTR param[2] = { 0, 0 };
	int i;
	for (i = 1; i < 3; ++i) // Two iterations: wParam and lParam.
	{
		if (ParamIndexIsOmitted(i))
			continue;
		ExprTokenType &this_param = *aParam[i];
		if (this_param.symbol == SYM_VAR)
			this_param.var->ToTokenSkipAddRef(this_param);
		switch (this_param.symbol)
		{
		case SYM_INTEGER:
			param[i-1] = (INT_PTR)this_param.value_int64;
			break;
		case SYM_OBJECT: // Support Buffer-like objects, i.e, objects with a "Ptr" property.
			size_t ptr_property_value;
			GetBufferObjectPtr(aResultToken, this_param.object, ptr_property_value);
			if (aResultToken.Exited())
				return;
			param[i - 1] = ptr_property_value;
			break;
		case SYM_STRING:
			LPTSTR error_marker;
			param[i - 1] = (INT_PTR)istrtoi64(this_param.marker, &error_marker);
			if (!*error_marker) // Valid number or empty string.
				break;
			//else: It's a non-numeric string; maybe the caller forgot the &address-of operator.
			// Note that an empty string would satisfy the check above.
			// Fall through:
		default:
			// SYM_FLOAT: Seems best to treat it as an error rather than truncating the value.
			_f_throw(i == 1 ? ERR_PARAM2_INVALID : ERR_PARAM3_INVALID, ErrorPrototype::Type);
		}
	}

	DWORD_PTR dwResult;
	if (aUseSend)
		successful = SendMessageTimeout(control_window, msg, (WPARAM)param[0], (LPARAM)param[1], SMTO_ABORTIFHUNG, timeout, &dwResult);
	else
		successful = PostMessage(control_window, msg, (WPARAM)param[0], (LPARAM)param[1]);

	if (!successful)
	{
		auto last_error = GetLastError();
		if (aUseSend && last_error == ERROR_TIMEOUT)
			_f_throw(ERR_TIMEOUT, ErrorPrototype::Timeout);
		_f_throw_win32(last_error); // Passing last_error reduces code size due to the implied additional GetLastError() call when omitting this parameter.
	}
	if (aUseSend)
		_f_return_i((__int64)dwResult);
	else
		_f_return_empty;
}



BIF_DECL(BIF_Process)
{
	_f_param_string_opt(aProcess, 0);

	HANDLE hProcess;
	DWORD pid;
	
	BuiltInFunctionID process_cmd = _f_callee_id;
	switch (process_cmd)
	{
	case FID_ProcessExist:
		// Return the discovered PID or zero if none.
		_f_return_i(*aProcess ? ProcessExist(aProcess) : GetCurrentProcessId());

	case FID_ProcessClose:
		_f_set_retval_i(0); // Set default.
		if (pid = ProcessExist(aProcess))  // Assign
		{
			if (hProcess = OpenProcess(PROCESS_TERMINATE, FALSE, pid))
			{
				if (TerminateProcess(hProcess, 0))
					_f_set_retval_i(pid); // Indicate success.
				CloseHandle(hProcess);
			}
		}
		// Since above didn't return, yield a PID of 0 to indicate failure.
		_f_return_retval;

	case FID_ProcessWait:
	case FID_ProcessWaitClose:
	{
		// This section is similar to that used for WINWAIT and RUNWAIT:
		bool wait_indefinitely;
		int sleep_duration;
		DWORD start_time;
		if (aParamCount > 1) // The param containing the timeout value was specified.
		{
			if (!ParamIndexIsNumeric(1))
				_f_throw_param(1, _T("Number"));
			wait_indefinitely = false;
			sleep_duration = (int)(TokenToDouble(*aParam[1]) * 1000); // Can be zero.
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
			if ((process_cmd == FID_ProcessWait) == (pid != 0)) // i.e. condition of this cmd is satisfied.
			{
				// For WaitClose: Since PID cannot always be determined (i.e. if process never existed,
				// there was no need to wait for it to close), for consistency, return 0 on success.
				_f_return_i(pid);
			}
			// Must cast to int or any negative result will be lost due to DWORD type:
			if (wait_indefinitely || (int)(sleep_duration - (GetTickCount() - start_time)) > SLEEP_INTERVAL_HALF)
			{
				MsgSleep(100);  // For performance reasons, don't check as often as the WinWait family does.
			}
			else // Done waiting.
			{
				// Return 0 if "Process Wait" times out; or the PID of the process that still exists
				// if "Process WaitClose" times out.
				_f_return_i(pid);
			}
		} // for()
	} // case
	} // switch()
}



BIF_DECL(BIF_ProcessSetPriority)
{
	_f_param_string_opt(aPriority, 0);
	_f_param_string_opt(aProcess, 1);

	DWORD pid, priority;
	HANDLE hProcess;

	switch (_totupper(*aPriority))
	{
	case 'L': priority = IDLE_PRIORITY_CLASS; break;
	case 'B': priority = BELOW_NORMAL_PRIORITY_CLASS; break;
	case 'N': priority = NORMAL_PRIORITY_CLASS; break;
	case 'A': priority = ABOVE_NORMAL_PRIORITY_CLASS; break;
	case 'H': priority = HIGH_PRIORITY_CLASS; break;
	case 'R': priority = REALTIME_PRIORITY_CLASS; break;
	default:
		// Since above didn't break, aPriority was invalid.
		_f_throw_param(0);
	}

	if (pid = *aProcess ? ProcessExist(aProcess) : GetCurrentProcessId())  // Assign
	{
		if (hProcess = OpenProcess(PROCESS_SET_INFORMATION, FALSE, pid)) // Assign
		{
			if (!SetPriorityClass(hProcess, priority))
				pid = 0; // Indicate failure.
			CloseHandle(hProcess);
		}
	}
	// Otherwise, return a PID of 0 to indicate failure.
	_f_return_i(pid);
}



void WinSetRegion(HWND aWnd, LPTSTR aPoints, ResultToken &aResultToken)
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
		//			_f_return_empty;
		//		// Otherwise, get rid of it since it didn't take effect:
		//		DeleteObject(hrgn);
		//	}
		//}
		//// Since above didn't return:
		//return OK;

		// It's undocumented by MSDN, but apparently setting the Window's region to NULL restores it
		// to proper working order:
		if (!SetWindowRgn(aWnd, NULL, TRUE))
			_f_throw_win32();
		_f_return_empty;
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
			goto arg_error;

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
				goto arg_error;
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
						goto arg_error;
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
				goto arg_error;
			} // switch()
		} // else

		if (   !(cp = _tcschr(cp, ' '))   ) // No more items.
			break;
	}

	if (!pt_count)
		goto arg_error;

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

	// Since above didn't return, it's a success.
	_f_return_empty;

arg_error:
	_f_throw_value(ERR_PARAM1_INVALID);

error:
	_f_throw_win32();
}



BIF_DECL(BIF_WinSet)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam + 1, aParamCount - 1))
		return;

	int value;
	BOOL success = FALSE;
	DWORD exstyle;
	BuiltInFunctionID cmd = _f_callee_id;

	LPTSTR aValue;
	ToggleValueType toggle = TOGGLE_INVALID;
	if (cmd == FID_WinSetAlwaysOnTop || cmd == FID_WinSetEnabled)
	{
		toggle = ParamIndexToToggleValue(0);
		if (toggle == TOGGLE_INVALID)
			_f_throw_param(0);
	}
	else
		aValue = ParamIndexToString(0, _f_number_buf);

	switch (cmd)
	{
	case FID_WinSetAlwaysOnTop:
	{
		HWND topmost_or_not;
		switch (toggle)
		{
		case TOGGLED_ON: topmost_or_not = HWND_TOPMOST; break;
		case TOGGLED_OFF: topmost_or_not = HWND_NOTOPMOST; break;
		case TOGGLE:
			exstyle = GetWindowLong(target_window, GWL_EXSTYLE);
			topmost_or_not = (exstyle & WS_EX_TOPMOST) ? HWND_NOTOPMOST : HWND_TOPMOST;
			break;
		}
		// SetWindowLong() didn't seem to work, at least not on some windows.  But this does.
		// As of v1.0.25.14, SWP_NOACTIVATE is also specified, though its absence does not actually
		// seem to activate the window, at least on XP (perhaps due to anti-focus-stealing measure
		// in Win98/2000 and beyond).  Or perhaps its something to do with the presence of
		// topmost_or_not (HWND_TOPMOST/HWND_NOTOPMOST), which might always avoid activating the
		// window.
		success = SetWindowPos(target_window, topmost_or_not, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
		break;
	}

	case FID_WinSetTransparent:
	case FID_WinSetTransColor:
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
		if (!_tcsicmp(aValue, _T("Off")) || !*aValue)
			// One user reported that turning off the attribute helps window's scrolling performance.
			success = SetWindowLong(target_window, GWL_EXSTYLE, exstyle & ~WS_EX_LAYERED);
		else
		{
			if (cmd == FID_WinSetTransparent)
			{
				if (!IsNumeric(aValue, FALSE, FALSE))
					_f_throw_param(0);

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
				success = SetLayeredWindowAttributes(target_window, 0, value, LWA_ALPHA);
			}
			else // cmd == FID_WinSetTransColor
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
				success = SetLayeredWindowAttributes(target_window, color, value, flags);
			}
		}
		break;
	}

	case FID_WinSetStyle:
	case FID_WinSetExStyle:
	{
		if (!*aValue)
		{
			// Seems best not to treat an explicit blank as zero.
			_f_throw_value(ERR_PARAM1_MUST_NOT_BE_BLANK);
		}
		int style_index = (cmd == FID_WinSetStyle) ? GWL_STYLE : GWL_EXSTYLE;
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
			if (new_style == orig_style || GetWindowLong(target_window, style_index) != orig_style) // Even a partial change counts as a success.
			{
				// SetWindowPos is also necessary, otherwise the frame thickness entirely around the window
				// does not get updated (just parts of it):
				SetWindowPos(target_window, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
				// Since SetWindowPos() probably doesn't know that the style changed, below is probably necessary
				// too, at least in some cases:
				InvalidateRect(target_window, NULL, TRUE); // Quite a few styles require this to become visibly manifest.
				success = TRUE;
			}
		}
		break;
	}

	case FID_WinSetEnabled:
		if (toggle == TOGGLE)
			toggle = IsWindowEnabled(target_window) ? TOGGLED_OFF : TOGGLED_ON;
		EnableWindow(target_window, toggle == TOGGLED_ON); // Return value is based on previous state, not success/failure.
		success = bool(IsWindowEnabled(target_window)) == (toggle == TOGGLED_ON);
		break;

	case FID_WinSetRegion:
		WinSetRegion(target_window, aValue, aResultToken);
		return;

	} // switch()
	if (!success)
		_f_throw_win32();
	_f_return_empty;
}



BIF_DECL(BIF_WinRedraw)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam, aParamCount))
		return;
	// Seems best to always have the last param be TRUE. Also, it seems best not to call
	// UpdateWindow(), which forces the window to immediately process a WM_PAINT message,
	// since that might not be desirable.  Some other methods of getting a window to redraw:
	// SendMessage(mHwnd, WM_NCPAINT, 1, 0);
	// RedrawWindow(mHwnd, NULL, NULL, RDW_INVALIDATE|RDW_FRAME|RDW_UPDATENOW);
	// SetWindowPos(mHwnd, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
	// GetClientRect(mHwnd, &client_rect); InvalidateRect(mHwnd, &client_rect, TRUE);
	InvalidateRect(target_window, NULL, TRUE);
	_f_return_empty;
}



BIF_DECL(BIF_WinMoveTopBottom)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam, aParamCount))
		return;
	HWND mode = _f_callee_id == FID_WinMoveBottom ? HWND_BOTTOM : HWND_TOP;
	// Note: SWP_NOACTIVATE must be specified otherwise the target window often fails to move:
	if (!SetWindowPos(target_window, mode, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE))
		_f_throw_win32();
	_f_return_empty;
}



BIF_DECL(BIF_WinSetTitle)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam + 1, aParamCount - 1))
		return;
	if (!SetWindowText(target_window, ParamIndexToString(0, _f_number_buf)))
		_f_throw_win32();
	_f_return_empty;
}



BIF_DECL(BIF_WinGetTitle)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam, aParamCount))
		return;

	int space_needed = target_window ? GetWindowTextLength(target_window) + 1 : 1; // 1 for terminator.
	if (!TokenSetResult(aResultToken, NULL, space_needed - 1))
		return;  // It already displayed the error.
	aResultToken.symbol = SYM_STRING;
	// Update length using the actual length, rather than the estimate provided by GetWindowTextLength():
	aResultToken.marker_length = GetWindowText(target_window, aResultToken.marker, space_needed);
	if (!aResultToken.marker_length)
		// There was no text to get or GetWindowText() failed.
		*aResultToken.marker = '\0';
}



BIF_DECL(BIF_WinGetClass)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam, aParamCount))
		return;
	TCHAR class_name[WINDOW_CLASS_SIZE];
	if (!GetClassName(target_window, class_name, _countof(class_name)))
		_f_throw_win32();
	_f_return(class_name);
}



void WinGetList(ResultToken &aResultToken, BuiltInFunctionID aCmd, LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
// Helper function for WinGet() to avoid having a WindowSearch object on its stack (since that object
// normally isn't needed).
{
	WindowSearch ws;
	ws.mFindLastMatch = true; // Must set mFindLastMatch to get all matches rather than just the first.
	if (aCmd == FID_WinGetList)
		if (  !(ws.mArray = Array::Create())  )
			_f_throw_oom;
	// Check if we were asked to "list" or count the active window:
	if (USE_FOREGROUND_WINDOW(aTitle, aText, aExcludeTitle, aExcludeText))
	{
		HWND target_window; // Set by the macro below.
		SET_TARGET_TO_ALLOWABLE_FOREGROUND(g->DetectHiddenWindows)
		if (aCmd == FID_WinGetCount)
		{
			_f_return_i(target_window != NULL); // i.e. 1 if a window was found.
		}
		// Otherwise, it's FID_WinGetList:
		if (target_window)
			// Since the target window has been determined, we know that it is the only window
			// to be put into the array:
			ws.mArray->Append((__int64)(size_t)target_window);
		// Otherwise, return an empty array.
		_f_return(ws.mArray);
	}
	// Enumerate all windows which match the criteria:
	// If aTitle is ahk_id nnnn, the Enum() below will be inefficient.  However, ahk_id is almost unheard of
	// in this context because it makes little sense, so no extra code is added to make that case efficient.
	if (ws.SetCriteria(*g, aTitle, aText, aExcludeTitle, aExcludeText)) // These criteria allow the possibility of a match.
		EnumWindows(EnumParentFind, (LPARAM)&ws);
	//else leave ws.mFoundCount set to zero (by the constructor).
	if (aCmd == FID_WinGetList)
		// Return the array even if it is empty:
		_f_return(ws.mArray);
	else // FID_WinGetCount
		_f_return_i(ws.mFoundCount);
}

void WinGetControlList(ResultToken &aResultToken, HWND aTargetWindow, bool aFetchHWNDs); // Forward declaration.



BIF_DECL(BIF_WinGet)
{
	_f_param_string_opt(aTitle, 0);
	_f_param_string_opt(aText, 1);
	_f_param_string_opt(aExcludeTitle, 2);
	_f_param_string_opt(aExcludeText, 3);

	BuiltInFunctionID cmd = _f_callee_id;

	if (cmd == FID_WinGetList || cmd == FID_WinGetCount)
	{
		WinGetList(aResultToken, cmd, aTitle, aText, aExcludeTitle, aExcludeText);
		return;
	}
	// Not List or Count, so it's a function that operates on a single window.
	HWND target_window = NULL;
	if (aParamCount > 0)
		switch (DetermineTargetHwnd(target_window, aResultToken, *aParam[0]))
		{
		case FAIL: return;
		case OK:
			if (!target_window)
				_f_throw(ERR_NO_WINDOW, ErrorPrototype::Target);
		}
	if (!target_window)
		target_window = WinExist(*g, aTitle, aText, aExcludeTitle, aExcludeText, cmd == FID_WinGetIDLast);
	if (!target_window)
		_f_throw(ERR_NO_WINDOW, ErrorPrototype::Target);

	switch(cmd)
	{
	case FID_WinGetID:
	case FID_WinGetIDLast:
		_f_return((size_t)target_window);

	case FID_WinGetPID:
	case FID_WinGetProcessName:
	case FID_WinGetProcessPath:
		DWORD pid;
		GetWindowThreadProcessId(target_window, &pid);
		if (cmd == FID_WinGetPID)
		{
			_f_return_i(pid);
		}
		// Otherwise, get the full path and name of the executable that owns this window.
		TCHAR process_name[MAX_PATH];
		GetProcessName(pid, process_name, _countof(process_name), cmd == FID_WinGetProcessName);
		_f_return(process_name);

	case FID_WinGetMinMax:
		// Testing shows that it's not possible for a minimized window to also be maximized (under
		// the theory that upon restoration, it *would* be maximized.  This is unfortunate if there
		// is no other way to determine what the restoration size and maximized state will be for a
		// minimized window.
		_f_return_i(IsZoomed(target_window) ? 1 : (IsIconic(target_window) ? -1 : 0));

	case FID_WinGetControls:
	case FID_WinGetControlsHwnd:
		WinGetControlList(aResultToken, target_window, cmd == FID_WinGetControlsHwnd);
		return;

	case FID_WinGetStyle:
	case FID_WinGetExStyle:
		_f_return_i(GetWindowLong(target_window, cmd == FID_WinGetStyle ? GWL_STYLE : GWL_EXSTYLE));

	case FID_WinGetTransparent:
	case FID_WinGetTransColor:
		COLORREF color;
		BYTE alpha;
		DWORD flags;
		if (!(GetLayeredWindowAttributes(target_window, &color, &alpha, &flags)))
			break;
		if (cmd == FID_WinGetTransparent)
		{
			if (flags & LWA_ALPHA)
			{
				_f_return_i(alpha);
			}
		}
		else // FID_WinGetTransColor
		{
			if (flags & LWA_COLORKEY)
			{
				// Store in hex format to aid in debugging scripts.  Also, the color is always
				// stored in RGB format, since that's what WinSet uses:
				LPTSTR result = _f_retval_buf;
				_stprintf(result, _T("0x%06X"), bgr_to_rgb(color));
				_f_return_p(result);
			}
			// Otherwise, this window does not have a transparent color (or it's not accessible to us,
			// perhaps for reasons described at MSDN GetLayeredWindowAttributes()).
		}
	}
	_f_return_empty;
}



void WinGetControlList(ResultToken &aResultToken, HWND aTargetWindow, bool aFetchHWNDs)
// Caller must ensure that aTargetWindow is non-NULL and valid.
// Every control is fetched rather than just a list of distinct class names (possibly with a
// second script array containing the quantity of each class) because it's conceivable that the
// z-order of the controls will be useful information to some script authors.
{
	control_list_type cl; // A big struct containing room to store class names and counts for each.
	if (  !(cl.target_array = Array::Create())  )
		_f_throw_oom;
	CL_INIT_CONTROL_LIST(cl)
	cl.fetch_hwnds = aFetchHWNDs;
	EnumChildWindows(aTargetWindow, EnumChildGetControlList, (LPARAM)&cl);
	_f_return(cl.target_array);
}



BOOL CALLBACK EnumChildGetControlList(HWND aWnd, LPARAM lParam)
{
	control_list_type &cl = *(control_list_type *)lParam;  // For performance and convenience.

	// cl.fetch_hwnds==true is a new mode in v1.0.43.06+ to help performance of AHK Window Info and other
	// scripts that want to operate directly on the HWNDs.
	if (cl.fetch_hwnds)
	{
		cl.target_array->Append((__int64)(size_t)aWnd);
	}
	else // The mode that fetches ClassNN vs. HWND.
	{
		TCHAR line[WINDOW_CLASS_SIZE + 5];  // +5 to allow room for the sequence number to be appended later below.
		int line_length;

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
		
		cl.target_array->Append(line); // Append class name+seq. number to the array.
	}

	return TRUE; // Continue enumeration through all the windows.
}



BIF_DECL(BIF_WinGetText)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam, aParamCount))
		return;

	length_and_buf_type sab;
	sab.buf = NULL; // Tell it just to calculate the length this time around.
	sab.total_length = 0; // Init
	sab.capacity = 0;     //
	EnumChildWindows(target_window, EnumChildGetText, (LPARAM)&sab);

	if (!sab.total_length) // No text in window.
		_f_return_empty;

	if (!TokenSetResult(aResultToken, NULL, sab.total_length))
		return;  // It already displayed the error.
	aResultToken.symbol = SYM_STRING;

	// Fetch the text directly into the var.  Also set the length explicitly
	// in case actual size written was different than the estimated size (since
	// GetWindowTextLength() can return more space that will actually be required
	// in certain circumstances, see MSDN):
	sab.buf = aResultToken.marker;
	// Note: The capacity member below exists because granted capacity might be a little larger than we asked for,
	// which allows the actual text fetched to be larger than the length estimate retrieved by the first pass
	// (which generally shouldn't happen since MSDN docs say that the actual length can be less, but never greater,
	// than the estimate length):
	sab.capacity = sab.total_length + 1; // Capacity includes the zero terminator, i.e. it's the size of the memory area.
	sab.total_length = 0; // Reinitialize.
	EnumChildWindows(target_window, EnumChildGetText, (LPARAM)&sab); // Get the text.

	// Length is set explicitly below in case it wound up being smaller than expected/estimated.
	// MSDN says that can happen generally, and also specifically because: "ANSI applications may have
	// the string in the buffer reduced in size (to a minimum of half that of the wParam value) due to
	// conversion from ANSI to Unicode."
	aResultToken.marker_length = sab.total_length;
	if (!sab.total_length)
	{
		// Something went wrong, so make sure we set to empty string.
		*sab.buf = '\0';
		_f_throw(ERR_FAILED);
	}
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



BIF_DECL(BIF_WinGetPos)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam + 4, aParamCount - 4))
		return;
	RECT rect;
	if (_f_callee_id == FID_WinGetPos)
	{
		GetWindowRect(target_window, &rect);
		rect.right -= rect.left; // Convert right to width.
		rect.bottom -= rect.top; // Convert bottom to height.
	}
	else // FID_WinGetClientPos
	{
		GetClientRect(target_window, &rect); // Get client pos relative to client (position is always 0,0).
		// Since the position is always 0,0, right,bottom are already equivalent to width,height.
		MapWindowPoints(target_window, NULL, (LPPOINT)&rect, 1); // Convert position to screen coordinates.
	}

	for (int i = 0; i < 4; ++i)
	{
		Var *var = ParamIndexToOutputVar(i);
		if (!var)
			continue;
		var->Assign(((int *)(&rect))[i]); // Always succeeds.
	}
	_f_return_empty;
}



BIF_DECL(BIF_Run)
{
	_f_param_string(target, 0);
	_f_param_string_opt(working_dir, 1);
	_f_param_string_opt(options, 2);
	Var *output_var_pid = ParamIndexToOutputVar(3);
	if (!g_script.ActionExec(target, nullptr, working_dir, true, options, nullptr, true, true, output_var_pid))
		_f_return_FAIL;
	_f_return_empty;
}



BIF_DECL(BIF_Env)
// Value := EnvGet(EnvVarName)
// EnvSet(EnvVarName, Value)
{
	LPTSTR aEnvVarName = ParamIndexToString(0, _f_number_buf);
	// Don't use a size greater than 32767 because that will cause it to fail on Win95 (tested by Robert Yalkin).
	// According to MSDN, 32767 is exactly large enough to handle the largest variable plus its zero terminator.
	// Update: In practice, at least on Windows 7, the limit only applies to the ANSI functions.
	TCHAR buf[32767];

	if (_f_callee_id == FID_EnvSet)
	{
		// MSDN: "If [the 2nd] parameter is NULL, the variable is deleted from the current process's environment."
		// No checking is currently done to ensure that aValue isn't longer than 32K, since testing shows that
		// larger values are supported by the Unicode functions, at least in some OSes.  If not supported,
		// SetEnvironmentVariable() should return 0 (fail) anyway.
		// Note: It seems that env variable names can contain spaces and other symbols, so it's best not to
		// validate aEnvVarName the same way we validate script variables (i.e. just let the return value
		// determine whether there's an error).
		LPTSTR aValue = ParamIndexIsOmitted(1) ? NULL : ParamIndexToString(1, buf);
		if (!SetEnvironmentVariable(aEnvVarName, aValue))
			_f_throw_win32();
		_f_return_empty;
	}

	// GetEnvironmentVariable() could be called twice, the first time to get the actual size.  But that would
	// probably perform worse since GetEnvironmentVariable() is a very slow function, so it seems best to fetch
	// it into a large buffer then just copy it to dest-var.
	DWORD length = GetEnvironmentVariable(aEnvVarName, buf, _countof(buf));
	if (length >= _countof(buf))
	{
		// In this case, length indicates the required buffer size, and the contents of the buffer are undefined.
		// Since our buffer is 32767 characters, the var apparently exceeds the documented limit, as can happen
		// if the var was set with the Unicode API.
		if (!TokenSetResult(aResultToken, NULL, length - 1))
			return;
		length = GetEnvironmentVariable(aEnvVarName, aResultToken.marker, length);
		if (!length)
			*aResultToken.marker = '\0'; // Ensure var is null-terminated.
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker_length = length; // Update length in case it changed.
		return;
	}
	_f_return(length ? buf : _T(""), length);
}



BIF_DECL(BIF_SysGet)
{
	_f_return(GetSystemMetrics(ParamIndexToInt(0)));
}



BIF_DECL(BIF_MonitorGet)
{
	MonitorInfoPackage mip = {0};  // Improves maintainability to initialize unconditionally, here.
	mip.monitor_info_ex.cbSize = sizeof(MONITORINFOEX); // Also improves maintainability.

	BuiltInFunctionID cmd = _f_callee_id;
	switch (cmd)
	{
	// For the next few cases, I'm not sure if it is possible to have zero monitors.  Obviously it's possible
	// to not have a monitor turned on or not connected at all.  But it seems likely that these various API
	// functions will provide a "default monitor" in the absence of a physical monitor connected to the
	// system.  To be safe, all of the below will assume that zero is possible, at least on some OSes or
	// under some conditions.  However, on Win95/NT, "1" is assumed since there is probably no way to tell
	// for sure if there are zero monitors except via GetSystemMetrics(SM_CMONITORS), which is a different
	// animal as described below.
	case FID_MonitorGetCount:
		// Don't use GetSystemMetrics(SM_CMONITORS) because of this:
		// MSDN: "GetSystemMetrics(SM_CMONITORS) counts only display monitors. This is different from
		// EnumDisplayMonitors, which enumerates display monitors and also non-display pseudo-monitors."
		mip.monitor_number_to_find = COUNT_ALL_MONITORS;
		EnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
		_f_return_i(mip.count); // Will return zero if the API ever returns a legitimate zero.

	// Even if the first monitor to be retrieved by the EnumProc is always the primary (which is doubtful
	// since there's no mention of this in the MSDN docs) it seems best to have this sub-cmd in case that
	// policy ever changes:
	case FID_MonitorGetPrimary:
		// The mip struct's values have already initialized correctly for the below:
		EnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
		_f_return_i(mip.count); // Will return zero if the API ever returns a legitimate zero.

	case FID_MonitorGet:
	case FID_MonitorGetWorkArea:
	// Params: N, Left, Top, Right, Bottom
	{
		mip.monitor_number_to_find = ParamIndexToOptionalInt(0, 0);  // If this returns 0, it will default to the primary monitor.
		EnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
		if (!mip.count || (mip.monitor_number_to_find && mip.monitor_number_to_find != mip.count))
			break;
		// Otherwise:
		LONG *monitor_rect = (LONG *)((cmd == FID_MonitorGetWorkArea) ? &mip.monitor_info_ex.rcWork : &mip.monitor_info_ex.rcMonitor);
		for (int i = 1; i <= 4; ++i) // Params: N (0), Left (1), Top, Right, Bottom.
			if (Var *var = ParamIndexToOutputVar(i))
				var->Assign(monitor_rect[i-1]);
		_f_return_i(mip.count); // Return the monitor number.
	}

	case FID_MonitorGetName: // Param: N
		mip.monitor_number_to_find = ParamIndexToOptionalInt(0, 0);  // If this returns 0, it will default to the primary monitor.
		EnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
		if (!mip.count || (mip.monitor_number_to_find && mip.monitor_number_to_find != mip.count))
			break;
		_f_return(mip.monitor_info_ex.szDevice);
	} // switch()
	
	// Since above didn't return, an error was detected.
	if (!mip.count) // Might be virtually impossible.
		_f_throw_win32();
	_f_throw_param(0);
}



BOOL CALLBACK EnumMonitorProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM lParam)
{
	MonitorInfoPackage &mip = *(MonitorInfoPackage *)lParam;  // For performance and convenience.
	if (mip.monitor_number_to_find == COUNT_ALL_MONITORS)
	{
		++mip.count;
		return TRUE;  // Enumerate all monitors so that they can be counted.
	}
	if (!GetMonitorInfo(hMonitor, &mip.monitor_info_ex)) // Failed.  Probably very rare.
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



void PixelSearch(Var *output_var_x, Var *output_var_y
	, int aLeft, int aTop, int aRight, int aBottom, COLORREF aColorRGB
	, int aVariation, bool aIsPixelGetColor, ResultToken &aResultToken)
// Author: The fast-mode PixelSearch was created by Aurelian Maga.
{
	// For maintainability, get RGB/BGR conversion out of the way early.
	COLORREF aColorBGR = rgb_to_bgr(aColorRGB);

	// Many of the following sections are similar to those in ImageSearch(), so they should be
	// maintained together.

	if (!aIsPixelGetColor)
	{
		output_var_x->Assign();  // Init to empty string regardless of whether we succeed here.
		output_var_y->Assign();  // Same.
	}

	POINT origin = {0};
	CoordToScreen(origin, COORD_MODE_PIXEL);
	aLeft   += origin.x;
	aTop    += origin.y;
	aRight  += origin.x;
	aBottom += origin.y;

	bool right_to_left = aLeft > aRight;
	bool bottom_to_top = aTop > aBottom;

	int left = aLeft;
	int top = aTop;
	int right = aRight;
	int bottom = aBottom;
	if (right_to_left)
	{
		left = aRight;
		right = aLeft;
	}
	if (bottom_to_top)
	{
		top = aBottom;
		bottom = aTop;
	}

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

	DWORD first_error = 0;
	HDC hdc = GetDC(NULL);
	if (!hdc)
	{
		first_error = GetLastError();
		goto error;
	}

	bool found = false; // Must init here for use by "goto fast_end".

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
	int search_width = right - left + 1;
	int search_height = bottom - top + 1;
	if (   !(sdc = CreateCompatibleDC(hdc)) || !(hbitmap_screen = CreateCompatibleBitmap(hdc, search_width, search_height))   )
		goto fast_end;

	if (   !(sdc_orig_select = SelectObject(sdc, hbitmap_screen))   )
		goto fast_end;

	// Copy the pixels in the search-area of the screen into the DC to be searched:
	if (   !(BitBlt(sdc, 0, 0, search_width, search_height, hdc, left, top, SRCCOPY))   )
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
		aResultToken.marker_length = _stprintf(aResultToken.marker, _T("0x%06X"), color);
	}
	else if (aVariation < 1) // Caller wants an exact match on one particular color.
	{
		if (screen_is_16bit)
			aColorRGB &= 0xF8F8F8F8;

		for (int j = 0; j < screen_pixel_count; ++j)
		{
			if (right_to_left && bottom_to_top)
				i = screen_pixel_count - j - 1;
			else if (right_to_left)
				i = j / screen_width * screen_width + screen_width - j % screen_width - 1;
			else if (bottom_to_top)
				i = (screen_pixel_count - j - 1) / screen_width * screen_width + j % screen_width;
			else
				i = j;

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

		for (int j = 0; j < screen_pixel_count; ++j)
		{
			if (right_to_left && bottom_to_top)
				i = screen_pixel_count - j - 1;
			else if (right_to_left)
				i = j / screen_width * screen_width + screen_width - j % screen_width - 1;
			else if (bottom_to_top)
				i = (screen_pixel_count - j - 1) / screen_width * screen_width + j % screen_width;
			else
				i = j;

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
	first_error = GetLastError(); // Might help to ensure the correct error number is thrown on failure.
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

	if (aIsPixelGetColor)
		return; // Return value was already set.

	// Otherwise, success.  Calculate xpos and ypos of where the match was found and adjust
	// coords to make them relative to the position of the target window (rect will contain
	// zeroes if this doesn't need to be done):
	if (found)
	{
		output_var_x->Assign((left + i%screen_width) - origin.x);
		output_var_y->Assign((top + i/screen_width) - origin.y);
	}
	_f_return_b(found);

error:
	_f_throw_win32(first_error);
}



BIF_DECL(BIF_PixelSearch)
{
	Var *output_var_x = ParamIndexToOutputVar(0);
	Var *output_var_y = ParamIndexToOutputVar(1);
	// Redundant due to prior validation of OutputVars:
	//if (!output_var_x)
	//	_f_throw_param(0, _T("variable reference"));
	//if (!output_var_y)
	//	_f_throw_param(1, _T("variable reference"));
	ASSERT(output_var_x && output_var_y);
	PixelSearch(output_var_x, output_var_y
		, ParamIndexToInt(2), ParamIndexToInt(3), ParamIndexToInt(4), ParamIndexToInt(5)
		, ParamIndexToInt(6), ParamIndexToOptionalInt(7, 0), false, aResultToken);
}



//ResultType Line::ImageSearch(int aLeft, int aTop, int aRight, int aBottom, LPTSTR aImageFile)
BIF_DECL(BIF_ImageSearch)
// Author: ImageSearch was created by Aurelian Maga.
{
	// Many of the following sections are similar to those in PixelSearch(), so they should be
	// maintained together.

	Var *output_var_x = ParamIndexToOutputVar(0);
	Var *output_var_y = ParamIndexToOutputVar(1);
	// Redundant due to prior validation of OutputVars:
	//if (!output_var_x)
	//	_f_throw_param(0, _T("variable reference"));
	//if (!output_var_y)
	//	_f_throw_param(1, _T("variable reference"));
	ASSERT(output_var_x && output_var_y);

	int aLeft = ParamIndexToInt(2);
	int aTop = ParamIndexToInt(3);
	int aRight = ParamIndexToInt(4);
	int aBottom = ParamIndexToInt(5);

	_f_param_string(aImageFile, 6);

	// Set default results (output variables), in case of early return:
	output_var_x->Assign();  // Init to empty string regardless of whether we succeed here.
	output_var_y->Assign();  // Same.

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
			goto arg7_error; // Bad option/format.
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
		goto arg7_error;

	DWORD first_error = 0;
	HDC hdc = GetDC(NULL);
	if (!hdc)
	{
		first_error = GetLastError();
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
			goto end;
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

end:
	first_error = GetLastError();
	ReleaseDC(NULL, hdc);
	if (!no_delete_bitmap && hbitmap_image)
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

	if (found)
	{
		// Calculate xpos and ypos of where the match was found and adjust coords to
		// make them relative to the position of the target window (rect will contain
		// zeroes if this doesn't need to be done):
		output_var_x->Assign((aLeft + i%screen_width) - origin.x);
		output_var_y->Assign((aTop + i/screen_width) - origin.y);
	}
	_f_return_b(found);

arg7_error:
	_f_throw_param(6);

error:
	_f_throw_win32(first_error);
}



/////////////////
// Main Window //
/////////////////

LRESULT CALLBACK MainWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
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
				PostMessage(hWnd, WM_COMMAND, g_script.mTrayMenu->mDefault->mMenuID, 0); // WM_COMMAND vs POST_AHK_USER_MENU to support the Standard menu items.
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

	case AHK_DIALOG:  // User defined msg sent from our functions MsgBox() or FileSelect().
	{
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
			// a small one isn't available.  This results in the FileSelect dialog's title
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
		if (IsInterruptible())
		{
			PostMessage(hWnd, iMsg, wParam, lParam);
			MsgSleep(-1, RETURN_AFTER_MESSAGES_SPECIAL_FILTER);
		}
		return 0;

	case WM_HOTKEY: // As a result of this app having previously called RegisterHotkey().
	case AHK_HOOK_HOTKEY:  // Sent from this app's keyboard or mouse hook.
	case AHK_HOTSTRING: // Added for v1.0.36.02 so that hotstrings work even while an InputBox or other non-standard msg pump is running.
	case AHK_CLIPBOARD_CHANGE: // Added for v1.0.44 so that clipboard notifications aren't lost while the script is displaying a MsgBox or other dialog.
	case AHK_INPUT_END:
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
			// now handled this way also so that owned windows won't be hidden when
			// the main window is hidden.
			ShowWindow(g_hWnd, SW_HIDE);
			return 0;
		}
		break;

	case WM_CLOSE:
		if (hWnd == g_hWnd) // i.e. not anything other than the main window.
		{
			// Receiving this msg is fairly unusual since SC_CLOSE is intercepted and redefined above.
			// However, it does happen if an external app is asking us to close, such as another
			// instance of this same script during the Reload command.  So treat it in a way similar
			// to the user having chosen Exit from the menu.
			//
			// Leave it up to ExitApp() to decide whether to terminate based upon whether
			// there is an OnExit function, whether that function is already running at
			// the time a new WM_CLOSE is received, etc.  It's also its responsibility to call
			// DestroyWindow() upon termination so that the WM_DESTROY message winds up being
			// received and process in this function (which is probably necessary for a clean
			// termination of the app and all its windows):
			g_script.ExitApp(EXIT_CLOSE);
			return 0;  // Verified correct.
		}
		// Otherwise, some window of ours other than our main window was destroyed.
		// Let DefWindowProc() handle it:
		break;

	case WM_ENDSESSION: // MSDN: "A window receives this message through its WindowProc function."
		if (wParam) // The session is being ended.
			g_script.ExitApp((lParam & ENDSESSION_LOGOFF) ? EXIT_LOGOFF : EXIT_SHUTDOWN);
		//else a prior WM_QUERYENDSESSION was aborted; i.e. the session really isn't ending.
		return 0;  // Verified correct.

	case AHK_EXIT_BY_RELOAD:
		g_script.ExitApp(EXIT_RELOAD);
		return 0; // Whether ExitApp() terminates depends on whether there's an OnExit function and what it does.

	case AHK_EXIT_BY_SINGLEINSTANCE:
		g_script.ExitApp(EXIT_SINGLEINSTANCE);
		return 0; // Whether ExitApp() terminates depends on whether there's an OnExit function and what it does.

	case WM_DESTROY:
		if (hWnd == g_hWnd) // i.e. not anything other than the main window.
		{
			if (!g_DestroyWindowCalled)
				// This is done because I believe it's possible for a WM_DESTROY message to be received
				// even though we didn't call DestroyWindow() ourselves (e.g. via DefWindowProc() receiving
				// and acting upon a WM_CLOSE or us calling DestroyWindow() directly) -- perhaps the window
				// is being forcibly closed or something else abnormal happened.  Make a best effort to run
				// the OnExit function, if present, even without a main window (testing on an earlier
				// versions shows that most commands work fine without the window). For EXIT_DESTROY,
				// it always terminates after running the OnExit callback:
				g_script.ExitApp(EXIT_DESTROY);
			// Do not do PostQuitMessage() here because we don't know the proper exit code.
			// MSDN: "The exit value returned to the system must be the wParam parameter of
			// the WM_QUIT message."
			// If we're here, it means our thread called DestroyWindow() directly or indirectly
			// (currently, it's only called directly).  By returning, our thread should resume
			// execution at the statement after DestroyWindow() in whichever caller called that:
			return 0;  // "If an application processes this message, it should return zero."
		}
		// Otherwise, some window of ours other than our main window was destroyed.
		// Let DefWindowProc() handle it:
		break;

	case WM_CREATE:
		// MSDN: If an application processes this message, it should return zero to continue
		// creation of the window. If the application returns 1, the window is destroyed and
		// the CreateWindowEx or CreateWindow function returns a NULL handle.
		return 0;

	case WM_WINDOWPOSCHANGED:
		if (hWnd == g_hWnd && (LPWINDOWPOS(lParam)->flags & SWP_HIDEWINDOW) && g_script.mIsReadyToExecute)
		{
			g_script.ExitIfNotPersistent(EXIT_CLOSE);
			return 0;
		}
		break; // Let DWP handle it.

	case WM_SIZE:
		if (hWnd == g_hWnd)
		{
			if (wParam == SIZE_MINIMIZED)
				// Minimizing the main window hides it.  This message generally doesn't arrive as a
				// result of user interaction, since WM_SYSCOMMAND, SC_MINIMIZE is handled as well.
				// However, this is necessary to keep the main window hidden when CreateWindows()
				// minimizes it during startup.
				ShowWindow(g_hWnd, SW_HIDE);
			else
				MoveWindow(g_hWndEdit, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
			return 0; // The correct return value for this msg.
		}
		break; // Let DWP handle it.
		
	case WM_SETFOCUS:
		if (hWnd == g_hWnd)
		{
			SetFocus(g_hWndEdit);  // Always focus the edit window, since it's the only navigable control.
			return 0;
		}
		break;

	case WM_CLIPBOARDUPDATE:
		if (g_script.mOnClipboardChange.Count()) // In case it's a bogus msg, it's our responsibility to avoid posting the msg if there's no function to call.
			PostMessage(g_hWnd, AHK_CLIPBOARD_CHANGE, 0, 0); // It's done this way to buffer it when the script is uninterruptible, etc.  v1.0.44: Post to g_hWnd vs. NULL so that notifications aren't lost when script is displaying a MsgBox or other dialog.
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

	case AHK_HOT_IF_EVAL: // HotCriterionAllowsFiring uses this to ensure expressions are evaluated only on the main thread.
		// Ensure wParam is a valid criterion (might prevent shatter attacks):
		for (HotkeyCriterion *cp = g_FirstHotExpr; cp; cp = cp->NextExpr)
			if ((WPARAM)cp == wParam)
				return cp->Eval((LPTSTR)lParam);
		return 0;

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
			g_script.RestoreTrayIcon();
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



bool FindAutoHotkeyUtilSub(LPTSTR aFile, LPTSTR aDir)
{
	SetCurrentDirectory(aDir);
	return GetFileAttributes(aFile) != INVALID_FILE_ATTRIBUTES;
}

bool FindAutoHotkeyUtil(LPTSTR aFile, bool &aFoundOurs)
{
	// Always try our directory first, in case it has different utils to the installed version.
	if (  !(aFoundOurs = FindAutoHotkeyUtilSub(aFile, g_script.mOurEXEDir))  )
	{
		// Try GetAHKInstallDir() so that compiled scripts running on machines that happen
		// to have AHK installed will still be able to fetch the help file and Window Spy:
		TCHAR installdir[MAX_PATH];
		if (   !GetAHKInstallDir(installdir)
			|| !FindAutoHotkeyUtilSub(aFile, installdir)   )
			return false;
	}
	return true;
}

void LaunchAutoHotkeyUtil(LPTSTR aFile, bool aIsScript)
{
	LPTSTR file = aFile, args = _T("");
	bool our_file, result = false;
	if (!FindAutoHotkeyUtil(aFile, our_file))
		return;
#ifndef AUTOHOTKEYSC
	// If it's a script in our directory, use our EXE to run it.
	TCHAR buf[64]; // More than enough for "/script WindowSpy.ahk".
	if (aIsScript && our_file)
	{
		sntprintf(buf, _countof(buf), _T("/script %s"), aFile);
		file = g_script.mOurEXE;
		args = buf;
	}
	//else it's not a script or it's the installed copy of WindowSpy.ahk, so just run it.
#endif
	if (!g_script.ActionExec(file, args, NULL, false))
	{
		TCHAR buf_file[64];
		sntprintf(buf_file, _countof(buf_file), _T("Could not launch %s"), aFile);
		MsgBox(buf_file, MB_ICONERROR);
	}
	SetCurrentDirectory(g_WorkingDir); // Restore the proper working directory.
}

void LaunchWindowSpy()
{
	LaunchAutoHotkeyUtil(_T("WindowSpy.ahk"), true);
}

void LaunchAutoHotkeyHelp()
{
	LaunchAutoHotkeyUtil(AHK_HELP_FILE, false);
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
		return true; // If there is an OnExit function, the above might not actually exit.
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

	// If we were called from a restricted place, such as via the Tray Menu or the Main Menu,
	// don't allow potentially sensitive info such as script lines and variables to be shown.
	// This is done so that scripts can be compiled more securely, making it difficult for anyone
	// to use ListLines to see the author's source code.  Rather than make exceptions for things
	// like KeyHistory, it seems best to forbid all information reporting except in cases where
	// existing info in the main window -- which must have gotten there via an allowed function
	// such as ListLines encountered in the script -- is being refreshed.  This is because in
	// that case, the script author has given de facto permission for that loophole (and it's
	// a pretty small one, not easy to exploit):
	if (aRestricted && !g_AllowMainWindow && (current_mode == MAIN_MODE_NO_CHANGE || aMode != MAIN_MODE_REFRESH))
	{
		// This used to set g_hWndEdit's text to an explanation for why the information will not
		// be shown, but it would never be seen unless the window was shown by some other means.
		// Since the menu items are disabled or removed, execution probably reached here as a
		// result of direct PostMessage to the script, so whoever did it can probably deal with
		// the lack of explanation (if the window was even visible).  The explanation contained
		// obsolete syntax (Menu, Tray, MainWindow) and was removed rather than updated to reduce
		// code size.  It seems unnecessary to even clear g_hWndEdit: either it's already empty,
		// or content was placed there deliberately and was already accessible.
		//SendMessage(g_hWndEdit, WM_SETTEXT, 0, (LPARAM)_T("Disabled"));
		return OK;
	}

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



LPTSTR Script::DefaultDialogTitle()
{
	// If the script has set A_ScriptName, use that:
	if (mScriptName)
		return mScriptName;
	// If available, the script's filename seems a much better title than the program name
	// in case the user has more than one script running:
	return (mFileName && *mFileName) ? mFileName : T_AHK_NAME_VERSION;
}

UserFunc* Script::CreateHotFunc()
{
	// Should only be called during load time.
	// Creates a new function for hotkeys and hotstrings.
	// Caller should abort loading if this function returns nullptr.
	
	if (mUnusedHotFunc)
	{
		auto tmp = mLastHotFunc = g->CurrentFunc = mUnusedHotFunc;
		mUnusedHotFunc = nullptr;
		mHotFuncs.mCount++;			// DefineFunc "removed" this func previously.
		ASSERT(mHotFuncs.mItem[mHotFuncs.mCount - 1] == tmp);
		return tmp;
	}
	
	static LPCTSTR sName = _T("<Hotkey>");
	auto func = new UserFunc(sName);
	
	g->CurrentFunc = func; // Must do this before calling AddVar

	// Add one parameter to hold the name of the hotkey/hotstring when triggered:
	func->mParam = SimpleHeap::Alloc<FuncParam>();
	if ( !(func->mParam[0].var = AddVar(_T("ThisHotkey"), 10, &func->mVars, 0, VAR_DECLARE_LOCAL | VAR_LOCAL_FUNCPARAM)) )
		return nullptr;

	func->mParam[0].default_type = PARAM_DEFAULT_NONE;
	func->mParam[0].is_byref = false;
	func->mParamCount = 1;
	func->mMinParams = 1;
	func->mIsFuncExpression = false;
	
	mLastHotFunc = func;
	mHotFuncs.Insert(func, mHotFuncs.mCount);
	return func;
}

ResultType MsgBoxParseOptions(LPTSTR aOptions, int &aType, double &aTimeout, HWND &aOwner)
{
	aType = 0;
	aTimeout = 0;

	//int button_option = 0;
	//int icon_option = 0;

	LPTSTR next_option, option_end;
	TCHAR option[1+MAX_NUMBER_SIZE];
	for (next_option = omit_leading_whitespace(aOptions); ; next_option = omit_leading_whitespace(option_end))
	{
		if (!*next_option)
			return OK;

		// Find the end of this option item:
		if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
			option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.
		size_t option_length = option_end - next_option;

		// Make a terminated copy for simplicity and to reduce ambiguity:
		if (option_length + 1 > _countof(option))
			goto invalid_option;
		tmemcpy(option, next_option, option_length);
		option[option_length] = '\0';

		if (option_length <= 5 && !_tcsnicmp(option, _T("Icon"), 4))
		{
			aType &= ~MB_ICONMASK;
			switch (option[4])
			{
			case 'x': case 'X': aType |= MB_ICONERROR; break;
			case '?': aType |= MB_ICONQUESTION; break;
			case '!': aType |= MB_ICONEXCLAMATION; break;
			case 'i': case 'I': aType |= MB_ICONINFORMATION; break;
			case '\0': break;
			default:
				goto invalid_option;
			}
		}
		else if (!_tcsnicmp(option, _T("Default"), 7) && IsNumeric(option + 7, FALSE, FALSE, FALSE))
		{
			int default_button = ATOI(option + 7);
			if (default_button < 1 || default_button > 0xF) // Currently MsgBox can only have 4 buttons, but MB_DEFMASK may allow for up to this many in future.
				goto invalid_option;
			aType = (aType & ~MB_DEFMASK) | ((default_button - 1) << 8); // 1=0, 2=0x100, 3=0x200, 4=0x300
		}
		else if (toupper(*option) == 'T' && IsNumeric(option + 1, FALSE, FALSE, TRUE))
		{
			aTimeout = ATOF(option + 1);
		}
		else if (!_tcsnicmp(option, _T("Owner"), 5) && IsNumeric(option + 5, TRUE, TRUE, FALSE))
		{
			aOwner = (HWND)ATOI64(option + 5); // This should be consistent with the Gui +Owner option.
		}
		else if (IsNumeric(option, FALSE, FALSE, FALSE))
		{
			int other_option = ATOI(option);
			// Clear any conflicting options which were previously set.
			if (other_option & MB_TYPEMASK) aType &= ~MB_TYPEMASK;
			if (other_option & MB_ICONMASK) aType &= ~MB_ICONMASK;
			if (other_option & MB_DEFMASK)  aType &= ~MB_DEFMASK;
			if (other_option & MB_MODEMASK) aType &= ~MB_MODEMASK;
			// All remaining options are bit flags (or not conflicting).
			aType |= other_option;
		}
		else
		{
			static LPCTSTR sButtonString[] = {
				_T("OK"), _T("OKCancel"), _T("AbortRetryIgnore"), _T("YesNoCancel"), _T("YesNo"), _T("RetryCancel"), _T("CancelTryAgainContinue"),
				_T("O"), _T("O/C"), _T("A/R/I"), _T("Y/N/C"), _T("Y/N"), _T("R/C"), _T("C/T/C"),
				_T("O"), _T("OC"), _T("ARI"), _T("YNC"), _T("YN"), _T("RC"), _T("CTC")
			};

			for (int i = 0; ; ++i)
			{
				if (i == _countof(sButtonString))
					goto invalid_option;

				if (!_tcsicmp(option, sButtonString[i]))
				{
					aType = (aType & ~MB_TYPEMASK) | (i % 7);
					break;
				}
			}
		}
	}
invalid_option:
	return ValueError(ERR_INVALID_OPTION, next_option, FAIL_OR_OK);
}


LPTSTR MsgBoxResultString(int aResult)
{
	switch (aResult)
	{
	case IDYES:			return _T("Yes");
	case IDNO:			return _T("No");
	case IDOK:			return _T("OK");
	case IDCANCEL:		return _T("Cancel");
	case IDABORT:		return _T("Abort");
	case IDIGNORE:		return _T("Ignore");
	case IDRETRY:		return _T("Retry");
	case IDCONTINUE:	return _T("Continue");
	case IDTRYAGAIN:	return _T("TryAgain");
	case AHK_TIMEOUT:	return _T("Timeout");
	default:			return NULL;
	}
}


BIF_DECL(BIF_MsgBox)
{
	int result;
	HWND dialog_owner = THREAD_DIALOG_OWNER; // Resolve macro only once to reduce code size.
	// dialog_owner is passed via parameter to avoid internally-displayed MsgBoxes from being
	// affected by script-thread's owner setting.
	if (!aParamCount) // When called explicitly with zero params, it displays this default msg.
	{
		result = MsgBox(_T("Press OK to continue."), MSGBOX_NORMAL, NULL, 0, dialog_owner);
	}
	else
	{
		_f_param_string_opt(aText, 0);
		_f_param_string_opt_def(aTitle, 1, NULL);
		_f_param_string_opt(aOptions, 2);
		int type;
		double timeout;
		if (!MsgBoxParseOptions(aOptions, type, timeout, dialog_owner))
		{
			aResultToken.SetExitResult(FAIL);
			return;
		}
		result = MsgBox(aText, type, aTitle, timeout, dialog_owner);
	}
	// If the MsgBox window can't be displayed for any reason, always return FAIL to
	// the caller because it would be unsafe to proceed with the execution of the
	// current script subroutine.  For example, if the script contains an IfMsgBox after,
	// this line, it's result would be unpredictable and might cause the subroutine to perform
	// the opposite action from what was intended (e.g. Delete vs. don't delete a file).
	// v1.0.40.01: Rather than displaying another MsgBox in response to a failed attempt to display
	// a MsgBox, it seems better (less likely to cause trouble) just to abort the thread.  This also
	// solves a double-msgbox issue when the maximum number of MsgBoxes is reached.  In addition, the
	// max-msgbox limit is the most common reason for failure, in which case a warning dialog has
	// already been displayed, so there is no need to display another:
	//if (!result)
	//	// It will fail if the text is too large (say, over 150K or so on XP), but that
	//	// has since been fixed by limiting how much it tries to display.
	//	// If there were too many message boxes displayed, it will already have notified
	//	// the user of this via a final MessageBox dialog, so our call here will
	//	// not have any effect.  The below only takes effect if MsgBox()'s call to
	//	// MessageBox() failed in some unexpected way:
	//	_f_throw("The MsgBox could not be displayed.");
	// v1.1.09.02: If the MsgBox failed due to invalid options, it seems better to display
	// an error dialog than to silently exit the thread:
	if (!result && GetLastError() == ERROR_INVALID_MSGBOX_STYLE)
		_f_throw_param(2);
	// Return a string such as "OK", "Yes" or "No" if possible, or fall back to the integer value.
	if (LPTSTR result_string = MsgBoxResultString(result))
		_f_return_p(result_string);
	else
		_f_return_i(result);
}



//////////////
// InputBox //
//////////////

ResultType InputBoxParseOptions(LPTSTR aOptions, InputBoxType &aInputBox)
{
	LPTSTR next_option, option_end;
	for (next_option = aOptions; *next_option; next_option = omit_leading_whitespace(option_end))
	{
		// Find the end of this option item:
		if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
			option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.

		// Temporarily terminate for simplicity and to reduce ambiguity:
		TCHAR orig_char = *option_end;
		*option_end = '\0';

		// The legacy InputBox command used "Hide", but "Password" seems clearer
		// and better for consistency with the equivalent Edit control option:
		if (!_tcsnicmp(next_option, _T("Password"), 8) && _tcslen(next_option) <= 9)
			aInputBox.password_char = next_option[8] ? next_option[8] : UorA(L'\x25CF', '*');
		else
		{
			// All of the remaining options are single-letter followed by a number:
			TCHAR option_char = ctoupper(*next_option);
			if (!_tcschr(_T("XYWHT"), option_char) // Not a valid option char.
				|| !IsNumeric(next_option + 1 // Or not a valid number.
					, option_char == 'X' || option_char == 'Y' // Only X and Y allow negative numbers.
					, FALSE, option_char == 'T')) // Only Timeout allows floating-point.
			{
				*option_end = orig_char; // Undo the temporary termination.
				return ValueError(ERR_INVALID_OPTION, next_option, FAIL_OR_OK);
			}

			switch (ctoupper(*next_option))
			{
			case 'W': aInputBox.width = DPIScale(ATOI(next_option + 1)); break;
			case 'H': aInputBox.height = DPIScale(ATOI(next_option + 1)); break;
			case 'X': aInputBox.xpos = ATOI(next_option + 1); break;
			case 'Y': aInputBox.ypos = ATOI(next_option + 1); break;
			case 'T': aInputBox.timeout = (DWORD)(ATOF(next_option + 1) * 1000); break;
			}
		}
		
		*option_end = orig_char; // Undo the temporary termination.
	}
	return OK;
}

BIF_DECL(BIF_InputBox)
{
	_f_param_string_opt(aText, 0);
	_f_param_string_opt_def(aTitle, 1, g_script.DefaultDialogTitle());
	_f_param_string_opt(aOptions, 2);
	_f_param_string_opt(aDefault, 3);

	InputBoxType inputbox;
	inputbox.title = aTitle;
	inputbox.text = aText;
	inputbox.default_string = aDefault;
	inputbox.return_string = nullptr;
	// Set defaults:
	inputbox.width = INPUTBOX_DEFAULT;
	inputbox.height = INPUTBOX_DEFAULT;
	inputbox.xpos = INPUTBOX_DEFAULT;
	inputbox.ypos = INPUTBOX_DEFAULT;
	inputbox.password_char = '\0';
	inputbox.timeout = 0;
	// Parse options and override defaults:
	if (!InputBoxParseOptions(aOptions, inputbox))
		_f_return_FAIL; // It already displayed the error.

	// At this point, we know a dialog will be displayed.  See macro's comments for details:
	DIALOG_PREP

	// Specify NULL as the owner window since we want to be able to have the main window in the foreground even
	// if there are InputBox windows.  Update: A GUI window can now be the parent if thread has that setting.
	INT_PTR result = DialogBoxParam(g_hInstance, MAKEINTRESOURCE(IDD_INPUTBOX), THREAD_DIALOG_OWNER
		, InputBoxProc, (LPARAM)&inputbox);

	DIALOG_END

	LPTSTR value = inputbox.return_string;
	LPTSTR reason;
	
	switch (result)
	{
	case AHK_TIMEOUT:	reason = _T("Timeout");	break;
	case IDOK:			reason = _T("OK");		break;
	case IDCANCEL:		reason = _T("Cancel");	break;
	default:			reason = nullptr;		break;
	}
	// result can be -1 or FAIL in case of failure, but since failure of any
	// kind is rare, all kinds are handled the same way, below.

	if (reason && value)
	{
		ExprTokenType argt[] = { _T("Result"), reason, _T("Value"), value };
		ExprTokenType *args[_countof(argt)] = { argt, argt+1, argt+2, argt+3 };
		if (Object *obj = Object::Create(args, _countof(args)))
		{
			free(value);
			_f_return(obj);
		}
	}

	free(value);
	// Since above didn't return, result is -1 (DialogBox somehow failed),
	// result is FAIL (something failed in InputBoxProc), or value is null.
	_f_throw(ERR_INTERNAL_CALL);
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

	switch(uMsg)
	{
	case WM_INITDIALOG:
	{
		SetWindowLongPtr(hWndDlg, DWLP_USER, lParam); // Store it for later use.
		auto &CURR_INPUTBOX = *(InputBoxType *)lParam;
		CURR_INPUTBOX.hwnd = hWndDlg;

		if (CURR_INPUTBOX.password_char)
			SendDlgItemMessage(hWndDlg, IDC_INPUTEDIT, EM_SETPASSWORDCHAR, CURR_INPUTBOX.password_char, 0);

		SetWindowText(hWndDlg, CURR_INPUTBOX.title);
		SetDlgItemText(hWndDlg, IDC_INPUTPROMPT, CURR_INPUTBOX.text);

		// Use the system's current language for the button names:
		typedef LPCWSTR(WINAPI*pfnUser)(int);
		HMODULE hMod = GetModuleHandle(_T("user32.dll"));
		pfnUser mbString = (pfnUser)GetProcAddress(hMod, "MB_GetString");
		if (mbString)
		{
			SetDlgItemTextW(hWndDlg, IDOK, mbString(0));
			SetDlgItemTextW(hWndDlg, IDCANCEL, mbString(1));
		}

		// Don't do this check; instead allow the MoveWindow() to occur unconditionally so that
		// the new button positions and such will override those set in the dialog's resource
		// properties:
		//if (CURR_INPUTBOX.width != INPUTBOX_DEFAULT || CURR_INPUTBOX.height != INPUTBOX_DEFAULT
		//	|| CURR_INPUTBOX.xpos != INPUTBOX_DEFAULT || CURR_INPUTBOX.ypos != INPUTBOX_DEFAULT)
		RECT rect;
		GetClientRect(hWndDlg, &rect);
		if (CURR_INPUTBOX.width != INPUTBOX_DEFAULT) rect.right = CURR_INPUTBOX.width;
		if (CURR_INPUTBOX.height != INPUTBOX_DEFAULT) rect.bottom = CURR_INPUTBOX.height;
		AdjustWindowRect(&rect, GetWindowLong(hWndDlg, GWL_STYLE), FALSE);
		int new_width = rect.right - rect.left;
		int new_height = rect.bottom - rect.top;

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

		// Regarding the timer ID: https://devblogs.microsoft.com/oldnewthing/20150924-00/?p=91521
		// Basically, timer IDs need only be non-zero and unique to the given HWND.
		if (CURR_INPUTBOX.timeout)
			SetTimer(hWndDlg, (UINT_PTR)&CURR_INPUTBOX, CURR_INPUTBOX.timeout, InputBoxTimeout);

		return TRUE; // i.e. let the system set the keyboard focus to the first visible control.
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

	case WM_GETMINMAXINFO:
	{
		// Increase the minimum width to prevent the buttons from overlapping:
		RECT rTmp;
		int min_width = 0;
		LPMINMAXINFO lpMMI = (LPMINMAXINFO)lParam;
		GetWindowRect(GetDlgItem(hWndDlg, IDOK), &rTmp);
		min_width += rTmp.right - rTmp.left;
		GetWindowRect(GetDlgItem(hWndDlg, IDCANCEL), &rTmp);
		min_width += rTmp.right - rTmp.left;
		lpMMI->ptMinTrackSize.x = min_width + 30;
	}

	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDOK:
		case IDCANCEL:
		{
			auto &CURR_INPUTBOX = *(InputBoxType *)GetWindowLongPtr(hWndDlg, DWLP_USER);
			// The entered text is used even if the user pressed the cancel button.  This allows the
			// cancel button to specify that a different operation should be performed on the text.
			WORD return_value = (WORD)FAIL;
			HWND hControl = GetDlgItem(hWndDlg, IDC_INPUTEDIT);
			if (hControl && CURR_INPUTBOX.UpdateResult(hControl))
				return_value = LOWORD(wParam); // IDOK or IDCANCEL
			// Since the user pressed a button to dismiss the dialog:
			// Timers belonging to a window are destroyed automatically when the window is destroyed,
			// but it seems prudent to clean up; also, EndDialog may fail, perhaps as a result of the
			// script interfering via OnMessage.
			if (CURR_INPUTBOX.timeout) // It has a timer.
				KillTimer(hWndDlg, (UINT_PTR)&CURR_INPUTBOX);
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
		auto &CURR_INPUTBOX = *(InputBoxType *)idEvent;
		// Even though the dialog has timed out, we still want to write anything the user
		// had a chance to enter into the output var.  This is because it's conceivable that
		// someone might want a short timeout just to enter something quick and let the
		// timeout dismiss the dialog for them (i.e. so that they don't have to press enter
		// or a button:
		INT_PTR result = FAIL;
		HWND hControl = GetDlgItem(hWnd, IDC_INPUTEDIT);
		if (hControl && CURR_INPUTBOX.UpdateResult(hControl))
			result = AHK_TIMEOUT;
		EndDialog(hWnd, result);
	}
	KillTimer(hWnd, idEvent);
}



ResultType InputBoxType::UpdateResult(HWND hControl)
{
	int space_needed = GetWindowTextLength(hControl) + 1;
	// Set up the result buffer.
	if (  !(return_string = tmalloc(space_needed))  )
		return FAIL; // BIF_InputBox will display an error.
	// Write to the variable:
	size_t len = (size_t)GetWindowText(hControl, return_string, space_needed);
	if (!len)
		// There was no text to get or GetWindowText() failed.
		*return_string = '\0';
	return OK;
}



VOID CALLBACK DerefTimeout(HWND hWnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	Line::FreeDerefBufIfLarge(); // It will also kill the timer, if appropriate.
}



ResultType Script::SetCoordMode(LPTSTR aCommand, LPTSTR aMode)
{
	CoordModeType mode = Line::ConvertCoordMode(aMode);
	CoordModeType shift = Line::ConvertCoordModeCmd(aCommand);
	if (shift == COORD_MODE_INVALID || mode == COORD_MODE_INVALID)
		return ValueError(ERR_INVALID_VALUE, aMode, FAIL_OR_OK);
	g->CoordMode = (g->CoordMode & ~(COORD_MODE_MASK << shift)) | (mode << shift);
	return OK;
}

ResultType Script::SetSendMode(LPTSTR aValue)
{
	g->SendMode = Line::ConvertSendMode(aValue, g->SendMode); // Leave value unchanged if ARG1 is invalid.
	return OK;
}

ResultType Script::SetSendLevel(int aValue, LPTSTR aValueStr)
{
	int sendLevel = aValue;
	if (!SendLevelIsValid(sendLevel))
		return ValueError(ERR_INVALID_VALUE, aValueStr, FAIL_OR_OK);
	g->SendLevel = sendLevel;
	return OK;
}



BIF_DECL(BIF_MouseGetPos)
// Returns OK or FAIL.
{
	// Since SYM_VAR is always VAR_NORMAL, these always resolve to normal vars or nullptr:
	Var *output_var_x = ParamIndexToOutputVar(0);
	Var *output_var_y = ParamIndexToOutputVar(1);
	Var *output_var_parent = ParamIndexToOutputVar(2);
	Var *output_var_child = ParamIndexToOutputVar(3);
	int aOptions = ParamIndexToOptionalInt(4, 0);

	POINT point;
	GetCursorPos(&point);  // Realistically, can't fail?

	POINT origin = {0};
	CoordToScreen(origin, COORD_MODE_MOUSE);

	if (output_var_x) // else the user didn't want the X coordinate, just the Y.
		if (!output_var_x->Assign(point.x - origin.x))
			_f_return_FAIL;
	if (output_var_y) // else the user didn't want the Y coordinate, just the X.
		if (!output_var_y->Assign(point.y - origin.y))
			_f_return_FAIL;

	_f_set_retval_p(_T(""), 0); // Set default.

	if (!output_var_parent && !output_var_child)
		_f_return_retval;

	if (output_var_parent)
		output_var_parent->Assign(); // Set default: empty.
	if (output_var_child)
		output_var_child->Assign(); // Set default: empty.

	// This is the child window.  Despite what MSDN says, WindowFromPoint() appears to fetch
	// a non-NULL value even when the mouse is hovering over a disabled control (at least on XP).
	HWND child_under_cursor = WindowFromPoint(point);
	if (!child_under_cursor)
		_f_return_retval;

	HWND parent_under_cursor = GetNonChildParent(child_under_cursor);  // Find the first ancestor that isn't a child.
	if (output_var_parent)
	{
		// Testing reveals that an invisible parent window never obscures another window beneath it as seen by
		// WindowFromPoint().  In other words, the below never happens, so there's no point in having it as a
		// documented feature:
		//if (!g->DetectHiddenWindows && !IsWindowVisible(parent_under_cursor))
		//	return output_var_parent->Assign();
		output_var_parent->AssignHWND(parent_under_cursor);
	}

	if (!output_var_child)
		_f_return_retval;

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
		_f_return_retval;

	if (aOptions & 0x02) // v1.0.43.06: Bitwise flag that means "return control's HWND vs. ClassNN".
	{
		output_var_child->AssignHWND(child_under_cursor);
		_f_return_retval;
	}

	class_and_hwnd_type cah;
	cah.hwnd = child_under_cursor;  // This is the specific control we need to find the sequence number of.
	TCHAR class_name[WINDOW_CLASS_SIZE];
	cah.class_name = class_name;
	if (!GetClassName(cah.hwnd, class_name, _countof(class_name) - 5))  // -5 to allow room for sequence number.
		_f_return_retval;
	cah.class_count = 0;  // Init for the below.
	cah.is_found = false; // Same.
	EnumChildWindows(parent_under_cursor, EnumChildFindSeqNum, (LPARAM)&cah); // Find this control's seq. number.
	if (!cah.is_found)
		_f_return_retval;
	// Append the class sequence number onto the class name and set the output param to be that value:
	sntprintfcat(class_name, _countof(class_name), _T("%d"), cah.class_count);
	if (!output_var_child->Assign(class_name))
		_f_return_FAIL;
	_f_return_retval;
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
	// Window Spy.  This overcomes the limitations of WindowFromPoint() and ChildWindowFromPoint().
	// The pixel at (left, top) lies inside the control, whereas MSDN says "the pixel at (right, bottom)
	// lies immediately outside the rectangle" -- so use < instead of <= below:
	if (pah.pt.x >= rect.left && pah.pt.x < rect.right && pah.pt.y >= rect.top && pah.pt.y < rect.bottom)
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

BIF_DECL(BIF_FormatTime)
// The compressed code size of this function is about 1 KB (2 KB uncompressed), which compares
// favorably to using setlocale()+strftime(), which together are about 8 KB of compressed code
// (setlocale() seems to be needed to put the user's or system's locale into effect for strftime()).
// setlocale() weighs in at about 6.5 KB compressed (14 KB uncompressed).
{
	_f_param_string_opt(aYYYYMMDD, 0);
	_f_param_string_opt(aFormat, 1);

	#define FT_MAX_INPUT_CHARS 2000
	// Input/format length is restricted since it must be translated and expanded into a new format
	// string that uses single quotes around non-alphanumeric characters such as punctuation:
	if (_tcslen(aFormat) > FT_MAX_INPUT_CHARS)
		_f_throw_param(1);

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
			if (!YYYYMMDDToSystemTime(yyyymmdd, st, false))
				_f_throw_param(0);
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
		// a normal format string such as " MM/dd/yy":
		LPTSTR candidate = omit_leading_whitespace(aFormat);
		if (!_tcsicmp(candidate, _T("YWeek")))
		{
			GetISOWeekNumber(output_buf, st.wYear, GetYDay(st.wMonth, st.wDay, IS_LEAP_YEAR(st.wYear)), st.wDayOfWeek);
			_f_return(output_buf);
		}
		if (!_tcsicmp(candidate, _T("YDay")) || !_tcsicmp(candidate, _T("YDay0")))
		{
			int yday = GetYDay(st.wMonth, st.wDay, IS_LEAP_YEAR(st.wYear));
			// Format with leading zeroes if YDay0 was used.
			_stprintf(output_buf, candidate[4] == '0' ? _T("%03d") : _T("%d"), yday);
			_f_return(output_buf);
		}
		if (!_tcsicmp(candidate, _T("WDay")))
		{
			LPTSTR buf = _f_retval_buf;
			*buf = '1' + st.wDayOfWeek; // '1' vs '0' to convert to 1-based for compatibility with A_WDay.
			buf[1] = '\0';
			_f_return_p(buf, 1);
		}

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

	_f_return(output_buf);
}



BIF_DECL(BIF_StrCase)
{
	size_t length;
	LPTSTR contents = ParamIndexToString(0, _f_number_buf, &length);

	// Make a modifiable copy of the string to return:
	if (!TokenSetResult(aResultToken, contents, length))
		return;
	aResultToken.symbol = SYM_STRING;
	contents = aResultToken.marker;

	if (_f_callee_id == FID_StrLower)
		CharLower(contents);
	else if (_f_callee_id == FID_StrUpper)
		CharUpper(contents);
	else // Convert to title case.
		StrToTitleCase(contents);
}



BIF_DECL(BIF_StrReplace)
{
	size_t length; // Going in to StrReplace(), it's the haystack length. Later (coming out), it's the result length. 
	_f_param_string(source, 0, &length);
	_f_param_string(oldstr, 1);
	_f_param_string_opt(newstr, 2);

	// Maintain this together with the equivalent section of BIF_InStr:
	StringCaseSenseType string_case_sense = ParamIndexToCaseSense(3);
	if (string_case_sense == SCS_INVALID 
		|| string_case_sense == SCS_INSENSITIVE_LOGICAL) // Not supported, seems more useful to throw rather than using SCS_INSENSITIVE.
		_f_throw_param(3);

	Var *output_var_count = ParamIndexToOutputVar(4); 
	UINT replacement_limit = (UINT)ParamIndexToOptionalInt64(5, UINT_MAX); 
	

	// Note: The current implementation of StrReplace() should be able to handle any conceivable inputs
	// without an empty string causing an infinite loop and without going infinite due to finding the
	// search string inside of newly-inserted replace strings (e.g. replacing all occurrences
	// of b with bcd would not keep finding b in the newly inserted bcd, infinitely).
	LPTSTR dest;
	UINT found_count = StrReplace(source, oldstr, newstr, string_case_sense, replacement_limit, -1, &dest, &length); // Length of haystack is passed to improve performance because TokenToString() can often discover it instantaneously.

	if (!dest) // Failure due to out of memory.
		_f_throw_oom; 

	if (dest != source) // StrReplace() allocated new memory rather than returning "source" to us unaltered.
	{
		// Return the newly allocated memory directly to our caller. 
		aResultToken.AcceptMem(dest, length);
	}
	else // StrReplace gave us back "source" unaltered because no replacements were needed.
	{
		_f_set_retval_p(dest, length); 
	}

	if (output_var_count)
		output_var_count->Assign((DWORD)found_count);
	_f_return_retval;
}



BIF_DECL(BIF_StrSplit)
// Array := StrSplit(String [, Delimiters, OmitChars, MaxParts])
{
	LPTSTR aInputString = ParamIndexToString(0, _f_number_buf);
	LPTSTR *aDelimiterList = NULL;
	int aDelimiterCount = 0;
	LPTSTR aOmitList = _T("");
	int splits_left = -2;

	if (aParamCount > 1)
	{
		if (auto arr = dynamic_cast<Array *>(TokenToObject(*aParam[1])))
		{
			aDelimiterCount = arr->Length();
			aDelimiterList = (LPTSTR *)_alloca(aDelimiterCount * sizeof(LPTSTR *));
			if (!arr->ToStrings(aDelimiterList, aDelimiterCount, aDelimiterCount))
				// Array contains something other than a string.
				goto throw_invalid_delimiter;
			for (int i = 0; i < aDelimiterCount; ++i)
				if (!*aDelimiterList[i])
					// Empty string in delimiter list. Although it could be treated similarly to the
					// "no delimiter" case, it's far more likely to be an error. If ever this check
					// is removed, the loop below must be changed to support "" as a delimiter.
					goto throw_invalid_delimiter;
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
	
	auto output_array = Array::Create();
	if (!output_array)
		goto throw_outofmem;

	if (!*aInputString // The input variable is blank, thus there will be zero elements.
		|| splits_left == -1) // The caller specified 0 parts.
		_f_return(output_array);
	
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
				_f_return(output_array); // All done.
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
	// chars, the item will be an empty string:
	if (output_array->Append(contents_of_next_element, element_length))
		_f_return(output_array); // All done.
	//else memory allocation failed, so fall through:
outofmem:
	// The fact that this section is executing means that a memory allocation failed and caused the
	// loop to break, so throw an exception.
	output_array->Release(); // Since we're not returning it.
	// Below: Using goto probably won't reduce code size (due to compiler optimizations), but keeping
	// rarely executed code at the end of the function might help performance due to cache utilization.
throw_outofmem:
	// Either goto was used or FELL THROUGH FROM ABOVE.
	_f_throw_oom;
throw_invalid_delimiter:
	_f_throw_param(1);
}



ResultType SplitPath(LPCTSTR aFileSpec, Var *output_var_name, Var *output_var_dir, Var *output_var_ext, Var *output_var_name_no_ext, Var *output_var_drive)
{
	// For URLs, "drive" is defined as the server name, e.g. http://somedomain.com
	LPCTSTR name = _T(""), name_delimiter = NULL, drive_end = NULL; // Set defaults to improve maintainability.
	LPCTSTR drive = omit_leading_whitespace(aFileSpec); // i.e. whitespace is considered for everything except the drive letter or server name, so that a pathless filename can have leading whitespace.
	LPCTSTR colon_double_slash = _tcsstr(aFileSpec, _T("://"));

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

	LPCTSTR ext_dot = _tcsrchr(name, '.');
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

BIF_DECL(BIF_SplitPath)
{
	LPTSTR mem_to_free = nullptr;
	_f_param_string(aFileSpec, 0);
	Var *vars[6];
	for (int i = 1; i < _countof(vars); ++i)
		vars[i] = ParamIndexToOutputVar(i);
	if (aParam[0]->symbol == SYM_VAR) // Check for overlap of input/output vars.
	{
		vars[0] = aParam[0]->var;
		// There are cases where this could be avoided, such as by careful ordering of the assignments
		// in SplitPath(), or when there's only one output var.  Also, real paths are generally short
		// enough that stack memory could be used.  However, perhaps simple is best in this case.
		for (int i = 1; i < _countof(vars); ++i)
		{
			if (vars[i] == vars[0])
			{
				aFileSpec = mem_to_free = _tcsdup(aFileSpec);
				if (!mem_to_free)
					_f_throw_oom;
				break;
			}
		}
	}
	aResultToken.SetValue(_T(""), 0);
	if (!SplitPath(aFileSpec, vars[1], vars[2], vars[3], vars[4], vars[5]))
		aResultToken.SetExitResult(FAIL);
	free(mem_to_free);
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
			return (sort_item1 > sort_item2) ? 1 : -1; // Stable sort.
		// Otherwise, it's either greater or less than zero:
		int result = (item1_minus_2 > 0.0) ? 1 : -1;
		return g_SortReverse ? -result : result;
	}
	// Otherwise, it's a non-numeric sort.
	// v1.0.43.03: Added support the new locale-insensitive mode.
	int result = (g_SortCaseSensitive != SCS_INSENSITIVE_LOGICAL)
		? tcscmp2(sort_item1, sort_item2, g_SortCaseSensitive) // Resolve large macro only once for code size reduction.
		: StrCmpLogicalW(sort_item1, sort_item2);
	if (!result)
		result = (sort_item1 > sort_item2) ? 1 : -1; // Stable sort.
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
	int result = (g_SortCaseSensitive != SCS_INSENSITIVE_LOGICAL)
		? tcscmp2(sort_item1, sort_item2, g_SortCaseSensitive) // Resolve large macro only once for code size reduction.
		: StrCmpLogicalW(sort_item1, sort_item2);
	if (!result)
		result = (sort_item1 > sort_item2) ? 1 : -1; // Stable sort.
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
	if (!g_SortFuncResult || g_SortFuncResult == EARLY_EXIT)
		return 0;

	// The following isn't necessary because by definition, the current thread isn't paused because it's the
	// thing that called the sort in the first place.
	//g_script.UpdateTrayIcon();

	LPTSTR aStr1 = *(LPTSTR *)a1;
	LPTSTR aStr2 = *(LPTSTR *)a2;
	ExprTokenType param[] = { aStr1, aStr2, __int64(aStr2 - aStr1) };
	__int64 i64;
	g_SortFuncResult = CallMethod(g_SortFunc, g_SortFunc, nullptr, param, _countof(param), &i64);
	// An alternative to g_SortFuncResult using 'throw' to abort qsort() produced slightly
	// smaller code, but in release builds the program crashed with code 0xC0000409 and the
	// catch() block never executed.

	int returned_int;
	if (i64)  // Maybe there's a faster/better way to do these checks. Can't simply typecast to an int because some large positives wrap into negative, maybe vice versa.
		returned_int = i64 < 0 ? -1 : 1;
	else  // The result was 0 or "".  Since returning 0 would make the results less predictable and doesn't seem to have any benefit,
		returned_int = aStr1 < aStr2 ? -1 : 1; // use the relative positions of the items as a tie-breaker.

	return returned_int;
}



BIF_DECL(BIF_Sort)
{
	// Set defaults in case of early goto:
	LPTSTR mem_to_free = NULL;
	LPTSTR *item = NULL; // The index/pointer list used for the sort.
	IObject *sort_func_orig = g_SortFunc; // Because UDFs can be interrupted by other threads -- and because UDFs can themselves call Sort with some other UDF (unlikely to be sure) -- backup & restore original g_SortFunc so that the "collapsing in reverse order" behavior will automatically ensure proper operation.
	ResultType sort_func_result_orig = g_SortFuncResult;
	g_SortFunc = NULL; // Now that original has been saved above, reset to detect whether THIS sort uses a UDF.
	g_SortFuncResult = OK;

	_f_param_string(aContents, 0);
	_f_param_string_opt(aOptions, 1);

	// Resolve options.  First set defaults for options:
	TCHAR delimiter = '\n';
	g_SortCaseSensitive = SCS_INSENSITIVE;
	g_SortNumeric = false;
	g_SortReverse = false;
	g_SortColumnOffset = 0;
	bool trailing_delimiter_indicates_trailing_blank_item = false, terminate_last_item_with_delimiter = false
		, trailing_crlf_added_temporarily = false, sort_by_naked_filename = false, sort_random = false
		, omit_dupes = false;
	LPTSTR cp;

	for (cp = aOptions; *cp; ++cp)
	{
		switch(_totupper(*cp))
		{
		case 'C':
			if (ctoupper(cp[1]) == 'L')
			{
				if (!_tcsnicmp(cp+2, _T("Logical") + 1, 6)) // CLogical.  Using "Logical" + 1 instead of "ogical" was confirmed to eliminate one string from the binary (due to string pooling).
				{
					cp += 7;
					g_SortCaseSensitive = SCS_INSENSITIVE_LOGICAL;
				}
				else
				{
					if (!_tcsnicmp(cp+2, _T("Locale") + 1, 5)) // CLocale
						cp += 6;
					else // CL
						++cp;
					g_SortCaseSensitive = SCS_INSENSITIVE_LOCALE;
				}
			}
			else if (!_tcsnicmp(cp+1, _T("Off"), 3)) // COff.  Using ctoupper() here significantly increased code size.
			{
				cp += 3;
				g_SortCaseSensitive = SCS_INSENSITIVE;
			}
			else if (cp[1] == '0') // C0
				g_SortCaseSensitive = SCS_INSENSITIVE;
			else // C  C1  COn
			{
				if (!_tcsnicmp(cp+1, _T("On"), 2)) // COn.  Using ctoupper() here significantly increased code size.
					cp += 2;
				g_SortCaseSensitive = SCS_SENSITIVE;
			}
			break;
		case 'D':
			if (!cp[1]) // Avoids out-of-bounds when the loop's own ++cp is done.
				break;
			++cp;
			if (*cp)
				delimiter = *cp;
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

	if (!ParamIndexIsOmitted(2))
	{
		if (  !(g_SortFunc = ParamIndexToObject(2))  )
		{
			aResultToken.ParamError(2, aParam[2]);
			goto end;
		}
		g_SortFunc->AddRef(); // Must be done in case the parameter was SYM_VAR and that var gets reassigned.
	}
	
	if (!*aContents) // Input is empty, nothing to sort, return empty string.
	{
		aResultToken.Return(_T(""), 0);
		goto end;
	}
	
	size_t item_count;

	// Check how many delimiters are present:
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
		aResultToken.Return(aContents, aContents_length);
		goto end;
	}

	// v1.0.47.05: It simplifies the code a lot to allocate and/or improves understandability to allocate
	// memory for trailing_crlf_added_temporarily even though technically it's done only to make room to
	// append the extra CRLF at the end.
	// v2.0: Never modify the caller's aContents, since it may be a quoted literal string or variable.
	//if (g_SortFunc || trailing_crlf_added_temporarily) // Do this here rather than earlier with the options parsing in case the function-option is present twice (unlikely, but it would be a memory leak due to strdup below).  Doing it here also avoids allocating if it isn't necessary.
	{
		// Comment is obsolete because if aContents is in a deref buffer, it has been privatized by ExpandArgs():
		// When g_SortFunc!=NULL, the copy of the string is needed because aContents may be in the deref buffer,
		// and that deref buffer is about to be overwritten by the execution of the script's UDF body.
		if (   !(mem_to_free = tmalloc(aContents_length + 3))   ) // +1 for terminator and +2 in case of trailing_crlf_added_temporarily.
		{
			aResultToken.MemoryError();
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
	item = (LPTSTR *)malloc((item_count + 1) * item_size);
	if (!item)
	{
		aResultToken.MemoryError();
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
	//    into the result.
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
			{
				auto &n = *(unsigned int *)(item_curr + 1); // i.e. the randoms are in the odd fields, the pointers in the even.
				GenRandom(&n, sizeof(n));
				n >>= 1;
				// For the above:
				// The random numbers given to SortRandom() must be within INT_MAX of each other, otherwise
				// integer overflow would cause the difference between two items to flip signs depending
				// on how far apart the numbers are, which would mean for some cases where A < B and B < C,
				// it's possible that C > A.  Without n >>= 1, the problem can be proven via the following
				// script, which shows a sharply biased distribution:
				//count := 0
				//Loop total := 10000
				//{
				//	var := Sort("1`n2`n3`n4`n5`n", "Random")
				//	if SubStr(var, 1, 1) = 5
				//		count += 1
				//}
				//MsgBox Round(count / total * 100, 2) "%"
			}
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
		{
			auto &n = *(unsigned int *)(item_curr + 1); // i.e. the randoms are in the odd fields, the pointers in the even.
			GenRandom(&n, sizeof(n));
			n >>= 1;
		}
	}
	else // Since the final item is not included in the count, point item_curr to the one before the last, for use below.
		item_curr -= unit_size;

	// Now aContents has been divided up based on delimiter.  Sort the array of pointers
	// so that they indicate the correct ordering to copy aContents into output_var:
	if (g_SortFunc) // Takes precedence other sorting methods.
	{
		qsort((void *)item, item_count, item_size, SortUDF);
		if (!g_SortFuncResult || g_SortFuncResult == EARLY_EXIT)
		{
			aResultToken.SetExitResult(g_SortFuncResult);
			goto end;
		}
	}
	else if (sort_random) // Takes precedence over all remaining options.
		qsort((void *)item, item_count, item_size, SortRandom);
	else
		qsort((void *)item, item_count, item_size, sort_by_naked_filename ? SortByNakedFilename : SortWithOptions);

	// Allocate space to store the result.
	if (!TokenSetResult(aResultToken, NULL, aContents_length))
		goto end;
	aResultToken.symbol = SYM_STRING;

	// Set default in case original last item is still the last item, or if last item was omitted due to being a dupe:
	size_t i, item_count_minus_1 = item_count - 1;
	DWORD omit_dupe_count = 0;
	bool keep_this_item;
	LPTSTR source, dest;
	LPTSTR item_prev = NULL;

	// Copy the sorted result back into output_var.  Do all except the last item, since the last
	// item gets special treatment depending on the options that were specified.
	item_curr = item; // i.e. Don't use [] indexing for the reason described higher above (same applies to item += unit_size below).
	for (dest = aResultToken.marker, i = 0; i < item_count; ++i, item_curr += unit_size)
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
			else if (g_SortCaseSensitive == SCS_INSENSITIVE_LOGICAL)
				keep_this_item = StrCmpLogicalW(*item_curr, item_prev);
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

	// Terminate the variable's contents.
	if (trailing_crlf_added_temporarily) // Remove the CRLF only after its presence was used above to simplify the code by reducing the number of types/cases.
	{
		dest[-2] = '\0';
		aResultToken.marker_length -= 2;
	}
	else
		*dest = '\0';

	if (omit_dupes)
	{
		if (omit_dupe_count) // Update the length to actual whenever at least one dupe was omitted.
		{
			aResultToken.marker_length = _tcslen(aResultToken.marker);
		}
	}
	//else it is not necessary to set the length here because its length hasn't
	// changed since it was originally set by the above call TokenSetResult.

end:
	if (!item)
		free(item);
	if (mem_to_free)
		free(mem_to_free);
	if (g_SortFunc)
		g_SortFunc->Release();
	g_SortFunc = sort_func_orig;
	g_SortFuncResult = sort_func_result_orig;
}



void DriveSpace(ResultToken &aResultToken, LPTSTR aPath, bool aGetFreeSpace)
// Because of NTFS's ability to mount volumes into a directory, a path might not necessarily
// have the same amount of free space as its root drive.  However, I'm not sure if this
// method here actually takes that into account.
{
	ASSERT(aPath && *aPath);

	TCHAR buf[MAX_PATH]; // MAX_PATH vs T_MAX_PATH because testing shows it doesn't support long paths even with \\?\.
	tcslcpy(buf, aPath, _countof(buf));
	size_t length = _tcslen(buf);
	if (buf[length - 1] != '\\' // Trailing backslash is absent,
		&& length + 1 < _countof(buf)) // and there's room to fix it.
	{
		buf[length++] = '\\';
		buf[length] = '\0';
	}
	//else it should still work unless this is a UNC path.

	// MSDN: "The GetDiskFreeSpaceEx function returns correct values for all volumes, including those
	// that are greater than 2 gigabytes."
	__int64 free_space;
	ULARGE_INTEGER total, free, used;
	if (!GetDiskFreeSpaceEx(buf, &free, &total, &used))
		_f_throw_win32();
	// Casting this way allows sizes of up to 2,097,152 gigabytes:
	free_space = (__int64)((unsigned __int64)(aGetFreeSpace ? free.QuadPart : total.QuadPart)
		/ (1024*1024));

	_f_return(free_space);
}



ResultType DriveLock(TCHAR aDriveLetter, bool aLockIt); // Forward declaration for BIF_Drive.

BIF_DECL(BIF_Drive)
{
	BuiltInFunctionID drive_cmd = _f_callee_id;

	LPTSTR aValue = ParamIndexToOptionalString(0, _f_number_buf);

	bool successful = false;
	TCHAR path[MAX_PATH]; // MAX_PATH vs. T_MAX_PATH because SetVolumeLabel() can't seem to make use of long paths.
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
	case FID_DriveLock:
	case FID_DriveUnlock:
		successful = DriveLock(*aValue, drive_cmd == FID_DriveLock);
		break;

	case FID_DriveEject:
	case FID_DriveRetract:
	{
		// Don't do DRIVE_SET_PATH in this case since trailing backslash is not wanted.
		// It seems best not to do the below check since:
		// 1) aValue usually lacks a trailing backslash, which might prevent DriveGetType() from working on some OSes.
		// 2) Eject (and more rarely, retract) works on some other drive types.
		// 3) CreateFile or DeviceIoControl will simply fail or have no effect if the drive isn't of the right type.
		//if (GetDriveType(aValue) != DRIVE_CDROM)
		//	_f_throw(ERR_FAILED);
		BOOL retract = (drive_cmd == FID_DriveRetract);
		TCHAR path[] { '\\', '\\', '.', '\\', 0, ':', '\0', '\0' };
		if (*aValue)
		{
			// Testing showed that a Volume GUID of the form \\?\Volume{...} will work even when
			// the drive isn't mapped to a drive letter.
			if (cisalpha(aValue[0]) && (!aValue[1] || aValue[1] == ':' && (!aValue[2] || aValue[2] == '\\' && !aValue[3])))
			{
				path[4] = aValue[0];
				aValue = path;
			}
		}
		else // When drive is omitted, operate upon the first CD/DVD drive.
		{
			path[6] = '\\'; // GetDriveType technically requires a slash, although it may work without.
			// Testing with mciSendString() while changing or removing drive letters showed that
			// its "default" drive is really just the first drive found in alphabetical order,
			// which is also the most obvious/intuitive choice.
			for (TCHAR c = 'A'; ; ++c)
			{
				path[4] = c;
				if (GetDriveType(path) == DRIVE_CDROM)
					break;
				if (c == 'Z')
					_f_throw(ERR_FAILED); // No CD/DVD drive found with a drive letter.  
			}
			path[6] = '\0'; // Remove the trailing slash for CreateFile to open the volume.
			aValue = path;
		}
		// Testing indicates neither this method nor the MCI method work with mapped drives or UNC paths.
		// That makes sense when one considers that the following opens the *volume*, whereas a network
		// share would correspond to a directory; i.e. this needs "\\.\D:" and not "\\.\D:\".
		HANDLE hVol = CreateFile(aValue, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
		if (hVol == INVALID_HANDLE_VALUE)
			break;
		DWORD unused;
		successful = DeviceIoControl(hVol, retract ? IOCTL_STORAGE_LOAD_MEDIA : IOCTL_STORAGE_EJECT_MEDIA, NULL, 0, NULL, 0, &unused, NULL);
		CloseHandle(hVol);
		break;
	}

	case FID_DriveSetLabel: // Note that it is possible and allowed for the new label to be blank.
		DRIVE_SET_PATH
		successful = SetVolumeLabel(path, ParamIndexToOptionalString(1, _f_number_buf)); // _f_number_buf can potentially be used by both parameters, since aValue is copied into path above.
		break;

	} // switch()

	if (!successful)
		_f_throw_win32();
	_f_return_empty;
}



ResultType DriveLock(TCHAR aDriveLetter, bool aLockIt)
{
	HANDLE hdevice;
	DWORD unused;
	BOOL result;
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
	CloseHandle(hdevice);
	return result ? OK : FAIL;
}



BIF_DECL(BIF_DriveGet)
{
	BuiltInFunctionID drive_get_cmd = _f_callee_id;

	// A parameter is mandatory for some, but optional for others:
	LPTSTR aValue = ParamIndexToOptionalString(0, _f_number_buf);
	if (!*aValue && drive_get_cmd != FID_DriveGetList && drive_get_cmd != FID_DriveGetStatusCD)
		_f_throw_value(ERR_PARAM1_MUST_NOT_BE_BLANK);
	
	if (drive_get_cmd == FID_DriveGetCapacity || drive_get_cmd == FID_DriveGetSpaceFree)
	{
		DriveSpace(aResultToken, aValue, drive_get_cmd == FID_DriveGetSpaceFree);
		return;
	}
	
	TCHAR path[T_MAX_PATH]; // T_MAX_PATH vs. MAX_PATH, though testing indicates only GetDriveType() supports long paths.
	size_t path_length;

	switch(drive_get_cmd)
	{

	case FID_DriveGetList:
	{
		UINT drive_type;
		#define ALL_DRIVE_TYPES 256
		if (!*aValue) drive_type = ALL_DRIVE_TYPES;
		else if (!_tcsicmp(aValue, _T("CDROM"))) drive_type = DRIVE_CDROM;
		else if (!_tcsicmp(aValue, _T("Removable"))) drive_type = DRIVE_REMOVABLE;
		else if (!_tcsicmp(aValue, _T("Fixed"))) drive_type = DRIVE_FIXED;
		else if (!_tcsicmp(aValue, _T("Network"))) drive_type = DRIVE_REMOTE;
		else if (!_tcsicmp(aValue, _T("RAMDisk"))) drive_type = DRIVE_RAMDISK;
		else if (!_tcsicmp(aValue, _T("Unknown"))) drive_type = DRIVE_UNKNOWN;
		else
			goto invalid_parameter;

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
		// An empty list should not be flagged as failure, even for FIXED drive_type.
		// For example, when booting Windows PE from a REMOVABLE drive, it mounts a RAMDISK
		// drive but there may be no FIXED drives present.
		//if (!*found_drives)
		//	goto error;
		_f_return(found_drives);
	}

	case FID_DriveGetFilesystem:
	case FID_DriveGetLabel:
	case FID_DriveGetSerial:
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
		case FID_DriveGetFilesystem: _f_return(file_system);
		case FID_DriveGetLabel:      _f_return(volume_name);
		case FID_DriveGetSerial:     _f_return(serial_number);
		}
		break;
	}

	case FID_DriveGetType:
	{
		DRIVE_SET_PATH
		switch (GetDriveType(path))
		{
		case DRIVE_UNKNOWN:   _f_return_p(_T("Unknown"));
		case DRIVE_REMOVABLE: _f_return_p(_T("Removable"));
		case DRIVE_FIXED:     _f_return_p(_T("Fixed"));
		case DRIVE_REMOTE:    _f_return_p(_T("Network"));
		case DRIVE_CDROM:     _f_return_p(_T("CDROM"));
		case DRIVE_RAMDISK:   _f_return_p(_T("RAMDisk"));
		default: // DRIVE_NO_ROOT_DIR
			_f_return_empty;
		}
		break;
	}

	case FID_DriveGetStatus:
	{
		DRIVE_SET_PATH
		DWORD sectors_per_cluster, bytes_per_sector, free_clusters, total_clusters;
		switch (GetDiskFreeSpace(path, &sectors_per_cluster, &bytes_per_sector, &free_clusters, &total_clusters)
			? ERROR_SUCCESS : GetLastError())
		{
		case ERROR_SUCCESS:        _f_return_p(_T("Ready"));
		case ERROR_PATH_NOT_FOUND: _f_return_p(_T("Invalid"));
		case ERROR_NOT_READY:      _f_return_p(_T("NotReady"));
		case ERROR_WRITE_PROTECT:  _f_return_p(_T("ReadOnly"));
		default:                   _f_return_p(_T("Unknown"));
		}
		break;
	}

	case FID_DriveGetStatusCD:
		// Don't do DRIVE_SET_PATH in this case since trailing backslash might prevent it from
		// working on some OSes.
		// It seems best not to do the below check since:
		// 1) aValue usually lacks a trailing backslash so that it will work correctly with "open c: type cdaudio".
		//    That lack might prevent DriveGetType() from working on some OSes.
		// 2) It's conceivable that tray eject/retract might work on certain types of drives even though
		//    they aren't of type DRIVE_CDROM.
		// 3) One or both of the calls to mciSendString() will simply fail if the drive isn't of the right type.
		//if (GetDriveType(aValue) != DRIVE_CDROM) // Testing reveals that the below method does not work on Network CD/DVD drives.
		//	_f_throw(ERR_FAILED);
		TCHAR mci_string[256], status[128];
		// Note that there is apparently no way to determine via mciSendString() whether the tray is ejected
		// or not, since "open" is returned even when the tray is closed but there is no media.
		if (!*aValue) // When drive is omitted, operate upon default CD/DVD drive.
		{
			if (mciSendString(_T("status cdaudio mode"), status, _countof(status), NULL)) // Error.
				goto failed;
		}
		else // Operate upon a specific drive letter.
		{
			sntprintf(mci_string, _countof(mci_string), _T("open %s type cdaudio alias cd wait shareable"), aValue);
			if (mciSendString(mci_string, NULL, 0, NULL)) // Error.
				goto failed;
			MCIERROR error = mciSendString(_T("status cd mode"), status, _countof(status), NULL);
			mciSendString(_T("close cd wait"), NULL, 0, NULL);
			if (error)
				goto failed;
		}
		// Otherwise, success:
		_f_return(status);

	} // switch()

failed: // Failure where GetLastError() isn't relevant.
	_f_throw(ERR_FAILED);

error: // Win32 error
	_f_throw_win32();

invalid_parameter:
	_f_throw_param(0);
}



#pragma region Sound support functions - Vista and later

enum class SoundControlType
{
	Volume,
	Mute,
	Name,
	IID
};

LPWSTR SoundDeviceGetName(IMMDevice *aDev)
{
	IPropertyStore *store;
	PROPVARIANT prop;
	if (SUCCEEDED(aDev->OpenPropertyStore(STGM_READ, &store)))
	{
		if (FAILED(store->GetValue(PKEY_Device_FriendlyName, &prop)))
			prop.pwszVal = nullptr;
		store->Release();
		return prop.pwszVal;
	}
	return nullptr;
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
		if (!*aDeviceString)
		{
			// Get default playback device.
			hr = deviceEnum->GetDefaultAudioEndpoint(eRender, eConsole, &aDevice);
		}
		else
		{
			// Parse Name:Index, Name or Index.
			int target_index = 0;
			LPWSTR target_name = L"";
			size_t target_name_length = 0;
			LPTSTR delim = _tcsrchr(aDeviceString, ':');
			if (delim)
			{
				target_index = ATOI(delim + 1) - 1;
				target_name = aDeviceString;
				target_name_length = delim - aDeviceString;
			}
			else if (IsNumeric(aDeviceString, FALSE, FALSE))
			{
				target_index = ATOI(aDeviceString) - 1;
			}
			else
			{
				target_name = aDeviceString;
				target_name_length = _tcslen(aDeviceString);
			}

			// Enumerate devices; include unplugged devices so that indices don't change when a device is plugged in.
			hr = deviceEnum->EnumAudioEndpoints(eAll, DEVICE_STATE_ACTIVE | DEVICE_STATE_UNPLUGGED, &devices);
			if (SUCCEEDED(hr))
			{
				if (target_name_length)
				{
					IMMDevice *dev;
					for (UINT u = 0; SUCCEEDED(hr = devices->Item(u, &dev)); ++u)
					{
						if (LPWSTR dev_name = SoundDeviceGetName(dev))
						{
							if (!_wcsnicmp(dev_name, target_name, target_name_length)
								&& target_index-- == 0)
							{
								CoTaskMemFree(dev_name);
								aDevice = dev;
								break;
							}
							CoTaskMemFree(dev_name);
						}
						dev->Release();
					}
				}
				else
					hr = devices->Item((UINT)target_index, &aDevice);
				devices->Release();
			}
		}
		deviceEnum->Release();
	}
	return hr;
}


struct SoundComponentSearch
{
	// Parameters of search:
	GUID target_iid;
	int target_instance;
	SoundControlType target_control;
	TCHAR target_name[128]; // Arbitrary; probably larger than any real component name.
	// Internal use/results:
	IUnknown *control;
	LPWSTR name; // Valid only when target_control == SoundControlType::Name.
	int count;
	// Internal use:
	DataFlow data_flow;
	bool ignore_remaining_subunits;
};


void SoundConvertComponent(LPTSTR aBuf, SoundComponentSearch &aSearch)
{
	if (IsNumeric(aBuf) == PURE_INTEGER)
	{
		*aSearch.target_name = '\0';
		aSearch.target_instance = ATOI(aBuf);
	}
	else
	{
		tcslcpy(aSearch.target_name, aBuf, _countof(aSearch.target_name));
		LPTSTR colon_pos = _tcsrchr(aSearch.target_name, ':');
		aSearch.target_instance = colon_pos ? ATOI(colon_pos + 1) : 1;
		if (colon_pos)
			*colon_pos = '\0';
	}
}


bool SoundSetGet_FindComponent(IPart *aRoot, SoundComponentSearch &aSearch)
{
	HRESULT hr;
	IPartsList *parts;
	IPart *part;
	UINT part_count;
	PartType part_type;
	LPWSTR part_name;
	
	bool check_name = *aSearch.target_name;

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
				if (   part_count == 1 // Ignore Connectors with no Subunits of their own.
					&& (!check_name || (SUCCEEDED(part->GetName(&part_name)) && !_wcsicmp(part_name, aSearch.target_name)))   )
				{
					if (++aSearch.count == aSearch.target_instance)
					{
						switch (aSearch.target_control)
						{
						case SoundControlType::IID:
							// Permit retrieving the IPart or IConnector itself.  Since there may be
							// multiple connected Subunits (and they can be enumerated or retrieved
							// via the Connector IPart), this is only done for the Connector.
							part->QueryInterface(aSearch.target_iid, (void **)&aSearch.control);
							break;
						case SoundControlType::Name:
							// check_name would typically be false in this case, since a value of true
							// would mean the caller already knows the component's name.
							part->GetName(&aSearch.name);
							break;
						}
						part->Release();
						parts->Release();
						return true;
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
					if (!aSearch.control && !aSearch.ignore_remaining_subunits)
					{
						// Query this part for the requested interface and let caller check the result.
						part->Activate(CLSCTX_ALL, aSearch.target_iid, (void **)&aSearch.control);

						// If this subunit has siblings, ignore any controls further up the line
						// as they're likely shared by other components (i.e. master controls).
						if (part_count > 1)
							aSearch.ignore_remaining_subunits = true;
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
	aSearch.control = nullptr;
	aSearch.name = nullptr;
	aSearch.ignore_remaining_subunits = false;
	
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

#pragma endregion


BIF_DECL(BIF_Sound)
{
	LPTSTR aSetting = nullptr;
	SoundComponentSearch search;
	if (_f_callee_id >= FID_SoundSetVolume)
	{
		search.target_control = SoundControlType(_f_callee_id - FID_SoundSetVolume);
		aSetting = ParamIndexToString(0, _f_number_buf);
		++aParam;
		--aParamCount;
	}
	else
		search.target_control = SoundControlType(_f_callee_id);

	switch (search.target_control)
	{
	case SoundControlType::Volume:
		search.target_iid = __uuidof(IAudioVolumeLevel);
		break;
	case SoundControlType::Mute:
		search.target_iid = __uuidof(IAudioMute);
		break;
	case SoundControlType::IID:
		LPTSTR iid; iid = ParamIndexToString(0);
		if (*iid != '{' || FAILED(CLSIDFromString(iid, &search.target_iid)))
			_f_throw_param(0);
		++aParam;
		--aParamCount;
		break;
	}

	_f_param_string_opt(aComponent, 0);
	_f_param_string_opt(aDevice, 1);

#define SOUND_MODE_IS_SET aSetting // Boolean: i.e. if it's not NULL, the mode is "SET".
	_f_set_retval_p(_T(""), 0); // Set default.

	float setting_scalar;
	if (SOUND_MODE_IS_SET)
	{
		setting_scalar = (float)(ATOF(aSetting) / 100);
		if (setting_scalar < -1)
			setting_scalar = -1;
		else if (setting_scalar > 1)
			setting_scalar = 1;
	}

	// Does user want to adjust the current setting by a certain amount?
	bool adjust_current_setting = aSetting && (*aSetting == '-' || *aSetting == '+');

	IMMDevice *device;
	HRESULT hr;

	hr = SoundSetGet_GetDevice(aDevice, device);
	if (FAILED(hr))
		_f_throw(ERR_SOUND_DEVICE, ErrorPrototype::Target);

	LPCTSTR errorlevel = NULL;
	float result_float;
	BOOL result_bool;

	if (!*aComponent) // Component is Master (omitted).
	{
		if (search.target_control == SoundControlType::IID)
		{
			void *result_ptr = nullptr;
			if (FAILED(device->QueryInterface(search.target_iid, (void **)&result_ptr)))
				device->Activate(search.target_iid, CLSCTX_ALL, NULL, &result_ptr);
			// For consistency with ComObjQuery, the result is returned even on failure.
			aResultToken.SetValue((UINT_PTR)result_ptr);
		}
		else if (search.target_control == SoundControlType::Name)
		{
			if (auto name = SoundDeviceGetName(device))
			{
				aResultToken.Return(name);
				CoTaskMemFree(name);
			}
		}
		else
		{
			// For Master/Speakers, use the IAudioEndpointVolume interface.  Some devices support master
			// volume control, but do not actually have a volume subunit (so the other method would fail).
			IAudioEndpointVolume *aev;
			hr = device->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, NULL, (void**)&aev);
			if (SUCCEEDED(hr))
			{
				if (search.target_control == SoundControlType::Volume)
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
				}
				aev->Release();
			}
		}
	}
	else
	{
		SoundConvertComponent(aComponent, search);
		
		if (!SoundSetGet_FindComponent(device, search))
		{
			errorlevel = ERR_SOUND_COMPONENT;
		}
		else if (search.target_control == SoundControlType::IID)
		{
			aResultToken.SetValue((UINT_PTR)search.control);
			search.control = nullptr; // Don't release it.
		}
		else if (search.target_control == SoundControlType::Name)
		{
			aResultToken.Return(search.name);
		}
		else if (!search.control)
		{
			errorlevel = ERR_SOUND_CONTROLTYPE;
		}
		else if (search.target_control == SoundControlType::Volume)
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
			}
		}
		else if (search.target_control == SoundControlType::Mute)
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
		}
control_fail:
		if (search.control)
			search.control->Release();
		if (search.name)
			CoTaskMemFree(search.name);
	}

	device->Release();
	
	if (errorlevel)
		_f_throw(errorlevel, ErrorPrototype::Target);
	if (FAILED(hr)) // This would be rare and unexpected.
		_f_throw_win32(hr);

	if (SOUND_MODE_IS_SET)
		return;

	switch (search.target_control)
	{
	case SoundControlType::Volume: aResultToken.Return(result_float); break;
	case SoundControlType::Mute: aResultToken.Return(result_bool); break;
	}
}



ResultType Line::SoundPlay(LPTSTR aFilespec, bool aSleepUntilDone)
{
	LPTSTR cp = omit_leading_whitespace(aFilespec);
	if (*cp == '*')
	{
		// ATOU() returns 0xFFFFFFFF for -1, which is relied upon to support the -1 sound.
		// Even if there are no enabled sound devices, MessageBeep() indicates success
		// (tested on Windows 10).  It's hard to imagine how the script would handle a
		// failure anyway, so just omit error checking.
		MessageBeep(ATOU(cp + 1));
		return OK;
	}
	// See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/multimed/htm/_win32_play.asp
	// for some documentation mciSendString() and related.
	// MAX_PATH note: There's no chance this API supports long paths even on Windows 10.
	// Research indicates it limits paths to 127 chars (not even MAX_PATH), but since there's
	// no apparent benefit to reducing it, we'll keep this size to ensure backward-compatibility.
	// Note that using a relative path does not help, but using short (8.3) names does.
	TCHAR buf[MAX_PATH * 2]; // Allow room for filename and commands.
	mciSendString(_T("status ") SOUNDPLAY_ALIAS _T(" mode"), buf, _countof(buf), NULL);
	if (*buf) // "playing" or "stopped" (so close it before trying to re-open with a new aFilespec).
		mciSendString(_T("close ") SOUNDPLAY_ALIAS, NULL, 0, NULL);
	sntprintf(buf, _countof(buf), _T("open \"%s\" alias ") SOUNDPLAY_ALIAS, aFilespec);
	if (mciSendString(buf, NULL, 0, NULL)) // Failure.
		goto error;
	g_SoundWasPlayed = true;  // For use by Script's destructor.
	if (mciSendString(_T("play ") SOUNDPLAY_ALIAS, NULL, 0, NULL)) // Failure.
		goto error;
	// Otherwise, the sound is now playing.
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

error:
	return LineError(ERR_FAILED, FAIL_OR_OK);
}



ResultType SetWorkingDir(LPTSTR aNewDir)
// Throws a script runtime exception on failure, but only if the script has begun runtime execution.
// This function was added in v1.0.45.01 for the reason described below.
{
	// v1.0.45.01: Since A_ScriptDir omits the trailing backslash for roots of drives (such as C:),
	// and since that variable probably shouldn't be changed for backward compatibility, provide
	// the missing backslash to allow SetWorkingDir %A_ScriptDir% (and others) to work as expected
	// in the root of a drive.
	// Update in 2018: The reason it wouldn't by default is that "C:" is actually a reference to the
	// the current directory if it's on C: drive, otherwise a reference to the path contained by the
	// env var "=C:".  Similarly, "C:x" is a reference to "x" inside that directory.
	// For details, see https://blogs.msdn.microsoft.com/oldnewthing/20100506-00/?p=14133
	// Although the override here creates inconsistency between SetWorkingDir and everything else
	// that can accept "C:", it is most likely what the user wants, and now there's also backward-
	// compatibility to consider since this workaround has been in place since 2006.
	// v1.1.31.00: Add the slash up-front instead of attempting SetCurrentDirectory(_T("C:"))
	// and comparing the result, since the comparison would always yield "not equal" due to either
	// a trailing slash or the directory being incorrect.
	TCHAR drive_buf[4];
	if (aNewDir[0] && aNewDir[1] == ':' && !aNewDir[2])
	{
		drive_buf[0] = aNewDir[0];
		drive_buf[1] = aNewDir[1];
		drive_buf[2] = '\\';
		drive_buf[3] = '\0';
		aNewDir = drive_buf;
	}

	if (!SetCurrentDirectory(aNewDir)) // Caused by nonexistent directory, permission denied, etc.
		return FAIL;
	// Otherwise, the change to the working directory succeeded.

	// Other than during program startup, this should be the only place where the official
	// working dir can change.  The exception is FileSelect(), which changes the working
	// dir as the user navigates from folder to folder.  However, the whole purpose of
	// maintaining g_WorkingDir is to workaround that very issue.
	if (g_script.mIsReadyToExecute) // Callers want this done only during script runtime.
		UpdateWorkingDir(aNewDir);
	return OK;
}



void UpdateWorkingDir(LPTSTR aNewDir)
// aNewDir is NULL or a path which was just passed to SetCurrentDirectory().
{
	TCHAR buf[T_MAX_PATH]; // Windows 10 long path awareness enables working dir to exceed MAX_PATH.
	// GetCurrentDirectory() is called explicitly, in case aNewDir is a relative path.
	// We want to store the absolute path:
	if (GetCurrentDirectory(_countof(buf), buf)) // Might never fail in this case, but kept for backward compatibility.
		aNewDir = buf;
	if (aNewDir)
		g_WorkingDir.SetString(aNewDir);
}



LPTSTR GetWorkingDir()
// Allocate a copy of the working directory from the heap.  This is used to support long
// paths without adding 64KB of stack usage per recursive #include <> on Unicode builds.
{
	TCHAR buf[T_MAX_PATH];
	if (GetCurrentDirectory(_countof(buf), buf))
		return _tcsdup(buf);
	return NULL;
}



BIF_DECL(BIF_FileSelect)
// Since other script threads can interrupt this command while it's running, it's important that
// this command not refer to sArgDeref[] and sArgVar[] anytime after an interruption becomes possible.
// This is because an interrupting thread usually changes the values to something inappropriate for this thread.
{
	if (g_nFileDialogs >= MAX_FILEDIALOGS)
	{
		// Have a maximum to help prevent runaway hotkeys due to key-repeat feature, etc.
		_f_throw(_T("The maximum number of File Dialogs has been reached."));
	}
	_f_param_string_opt(aOptions, 0);
	_f_param_string_opt(aWorkingDir, 1);
	_f_param_string_opt(aGreeting, 2);
	_f_param_string_opt(aFilter, 3);
	
	LPCTSTR default_file_name = _T("");

	TCHAR working_dir[MAX_PATH]; // Using T_MAX_PATH vs. MAX_PATH did not help on Windows 10.0.16299 (see below).
	if (!aWorkingDir || !*aWorkingDir)
		*working_dir = '\0';
	else
	{
		// Compress the path if possible to support longer paths.  Without this, any path longer
		// than MAX_PATH would be ignored, presumably because the dialog, as part of the shell,
		// does not support long paths.  Surprisingly, although Windows 10 long path awareness
		// does not allow us to pass a long path for working_dir, it does affect whether the long
		// path is used in the address bar and returned filenames.
		if (_tcslen(aWorkingDir) >= MAX_PATH)
			GetShortPathName(aWorkingDir, working_dir, _countof(working_dir));
		else
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
				default_file_name = last_backslash + 1;
				*last_backslash = '\0'; // Make the working directory just the file's path.
			}
			else // The entire working_dir string is the default file (unless this is a clsid).
				if (!is_clsid)
					default_file_name = working_dir; // This signals it to use the default directory.
				//else leave working_dir set to the entire clsid string in case it's somehow valid.
		}
		// else it is a directory, so just leave working_dir set as it was initially.
	}

	TCHAR pattern[1024];
	*pattern = '\0'; // Set default.
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
		}
		else // No open-paren, so assume the entire string is the filter.
			tcslcpy(pattern, aFilter, _countof(pattern));
	}
	UINT filter_count = 0;
	COMDLG_FILTERSPEC filters[2];
	if (*pattern)
	{
		// Remove any spaces present in the pattern, such as a space after every semicolon
		// that separates the allowed file extensions.  The API docs specify that there
		// should be no spaces in the pattern itself, even though it's okay if they exist
		// in the displayed name of the file-type:
		// Update by Lexikos: Don't remove spaces, since that gives incorrect behaviour for more
		// complex patterns like "prefix *.ext" (where the space should be considered part of the
		// pattern).  Although the docs for OPENFILENAMEW say "Do not include spaces", it may be
		// just because spaces are considered part of the pattern.  On the other hand, the docs
		// relating to IFileDialog::SetFileTypes() say nothing about spaces; and in fact, using a
		// pattern like "*.cpp; *.h" will work correctly (possibly due to how leading spaces work
		// with the file system).
		//StrReplace(pattern, _T(" "), _T(""), SCS_SENSITIVE);
		filters[0].pszName = aFilter;
		filters[0].pszSpec = pattern;
		++filter_count;
	}
	// Always include the All Files (*.*) filter, since there doesn't seem to be much
	// point to making this an option.  This is because the user could always type
	// *.* (or *) and press ENTER in the filename field and achieve the same result:
	filters[filter_count].pszName = _T("All Files (*.*)");
	filters[filter_count].pszSpec = _T("*.*");
	++filter_count;

	// v1.0.43.09: OFN_NODEREFERENCELINKS is now omitted by default because most people probably want a click
	// on a shortcut to navigate to the shortcut's target rather than select the shortcut and end the dialog.
	// v2: Testing on Windows 7 and 10 indicated IFileDialog doesn't change the working directory while the
	// user navigates, unlike GetOpenFileName/GetSaveFileName, and doesn't appear to affect the CWD at all.
	DWORD flags = FOS_NOCHANGEDIR; // FOS_NOCHANGEDIR according to MS: "Don't change the current working directory."

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
	switch (ctoupper(*aOptions))
	{
	case 'D':
		++aOptions;
		flags |= FOS_PICKFOLDERS;
		if (*aFilter)
			_f_throw_value(ERR_PARAM4_MUST_BE_BLANK);
		filter_count = 0;
		break;
	case 'M':  // Multi-select.
		++aOptions;
		flags |= FOS_ALLOWMULTISELECT;
		break;
	case 'S': // Have a "Save" button rather than an "Open" button.
		++aOptions;
		always_use_save_dialog = true;
		break;
	}

	TCHAR greeting[1024];
	if (aGreeting && *aGreeting)
		tcslcpy(greeting, aGreeting, _countof(greeting));
	else
		// Use a more specific title so that the dialogs of different scripts can be distinguished
		// from one another, which may help script automation in rare cases:
		sntprintf(greeting, _countof(greeting), _T("Select %s - %s")
			, (flags & FOS_PICKFOLDERS) ? _T("Folder") : _T("File"), g_script.DefaultDialogTitle());

	int options = ATOI(aOptions);
	if (options & 0x20)
		flags |= FOS_NODEREFERENCELINKS;
	if (options & 0x10)
		flags |= FOS_OVERWRITEPROMPT;
	if (options & 0x08)
		flags |= FOS_CREATEPROMPT;
	if (options & 0x02)
		flags |= FOS_PATHMUSTEXIST;
	if (options & 0x01)
		flags |= FOS_FILEMUSTEXIST;

	// Despite old documentation indicating it was due to an "OS quirk", previous versions were specifically
	// designed to enable the Save button when OFN_OVERWRITEPROMPT is present but not OFN_CREATEPROMPT, since
	// the former requires the Save dialog while the latter requires the Open dialog.  If both options are
	// present, the caller must specify or omit 'S' to choose the dialog type, and one option has no effect.
	if ((flags & FOS_OVERWRITEPROMPT) && !(flags & (FOS_CREATEPROMPT | FOS_PICKFOLDERS)))
		always_use_save_dialog = true;

	IFileDialog *pfd = NULL;
	HRESULT hr = CoCreateInstance(always_use_save_dialog ? CLSID_FileSaveDialog : CLSID_FileOpenDialog,
		NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pfd));
	if (FAILED(hr))
		_f_throw_win32(hr);

	pfd->SetOptions(flags);
	pfd->SetTitle(greeting);
	if (filter_count)
		pfd->SetFileTypes(filter_count, filters);
	pfd->SetFileName(default_file_name);

	if (*working_dir && default_file_name != working_dir)
	{
		IShellItem *psi;
		if (SUCCEEDED(SHCreateItemFromParsingName(working_dir, nullptr, IID_PPV_ARGS(&psi))))
		{
			pfd->SetFolder(psi);
			psi->Release();
		}
	}

	// At this point, we know a dialog will be displayed.  See macro's comments for details:
	DIALOG_PREP
	POST_AHK_DIALOG(0) // Do this only after the above. Must pass 0 for timeout in this case.
	++g_nFileDialogs;
	auto result = pfd->Show(THREAD_DIALOG_OWNER);
	--g_nFileDialogs;
	DIALOG_END

	if (flags & FOS_ALLOWMULTISELECT)
	{
		auto *files = Array::Create();
		IFileOpenDialog *pfod;
		if (SUCCEEDED(result) && SUCCEEDED(pfd->QueryInterface(&pfod)))
		{
			IShellItemArray *penum;
			if (SUCCEEDED(pfod->GetResults(&penum)))
			{
				DWORD count = 0;
				penum->GetCount(&count);
				for (DWORD i = 0; i < count; ++i)
				{
					IShellItem *psi;
					if (SUCCEEDED(penum->GetItemAt(i, &psi)))
					{
						LPWSTR filename;
						if (SUCCEEDED(psi->GetDisplayName(SIGDN_FILESYSPATH, &filename)))
						{
							files->Append(filename);
							CoTaskMemFree(filename);
						}
						psi->Release();
					}
				}
				penum->Release();
			}
			pfod->Release();
		}
		pfd->Release();
		_f_return(files);
	}
	aResultToken.SetValue(_T(""), 0); // Set default.
	IShellItem *psi;
	if (SUCCEEDED(result) && SUCCEEDED(pfd->GetResult(&psi)))
	{
		LPWSTR filename;
		if (SUCCEEDED(psi->GetDisplayName(SIGDN_FILESYSPATH, &filename)))
		{
			aResultToken.Return(filename);
			CoTaskMemFree(filename);
		}
		psi->Release();
	}
	//else: User pressed CANCEL vs. OK to dismiss the dialog or there was a problem displaying it.
		// Currently assuming the user canceled, otherwise this would tell us whether an error
		// occurred vs. the user canceling: if (result != HRESULT_FROM_WIN32(ERROR_CANCELLED))
	pfd->Release();
}



// As of 2019-09-29, noinline reduces code size by over 20KB on VC++ 2019.
// Prior to merging Util_CreateDir with this, it wasn't inlined.
DECLSPEC_NOINLINE
bool Line::FileCreateDir(LPTSTR aDirSpec, LPTSTR aCanModifyDirSpec)
{
	if (!aDirSpec || !*aDirSpec)
	{
		SetLastError(ERROR_INVALID_PARAMETER);
		return false;
	}

	DWORD attr = GetFileAttributes(aDirSpec);
	if (attr != 0xFFFFFFFF)  // aDirSpec already exists.
	{
		SetLastError(ERROR_ALREADY_EXISTS);
		return (attr & FILE_ATTRIBUTE_DIRECTORY) != 0; // Indicate success if it already exists as a dir.
	}

	// If it has a backslash, make sure all its parent directories exist before we attempt
	// to create this directory:
	LPTSTR last_backslash = _tcsrchr(aDirSpec, '\\');
	if (last_backslash > aDirSpec // v1.0.48.04: Changed "last_backslash" to "last_backslash > aDirSpec" so that an aDirSpec with a leading \ (but no other backslashes), such as \dir, is supported.
		&& last_backslash[-1] != ':') // v1.1.31.00: Don't attempt FileCreateDir("C:") since that's equivalent to either "C:\" or the working directory (which already exists), or FileCreateDir("\\?\C:") since it always fails.
	{
		LPTSTR parent_dir;
		if (aCanModifyDirSpec)
		{
			parent_dir = aDirSpec; // Caller provided a modifiable aDirSpec.
			*last_backslash = '\0'; // Temporarily terminate for parent directory.
		}
		else
		{
			// v1.1.31.00: Allocate a modifiable buffer to be used by all calls (supports long paths).
			parent_dir = (LPTSTR)_alloca((last_backslash - aDirSpec + 1) * sizeof(TCHAR));
			tcslcpy(parent_dir, aDirSpec, last_backslash - aDirSpec + 1); // Omits the last backslash.
		}
		bool exists = FileCreateDir(parent_dir, parent_dir); // Recursively create all needed ancestor directories.
		if (aCanModifyDirSpec)
			*last_backslash = '\\'; // Undo temporary termination.

		// v1.0.44: Fixed ErrorLevel being set to 1 when the specified directory ends in a backslash.  In such cases,
		// two calls were made to CreateDirectory for the same folder: the first without the backslash and then with
		// it.  Since the directory already existed on the second call, ErrorLevel was wrongly set to 1 even though
		// everything succeeded.  So now, when recursion finishes creating all the ancestors of this directory
		// our own layer here does not call CreateDirectory() when there's a trailing backslash because a previous
		// layer already did:
		if (!last_backslash[1] || !exists)
			return exists;
	}

	// The above has recursively created all parent directories of aDirSpec if needed.
	// Now we can create aDirSpec.
	return CreateDirectory(aDirSpec, NULL);
}



ResultType ConvertFileOptions(ResultToken &aResultToken, LPTSTR aOptions, UINT &codepage, bool &translate_crlf_to_lf, unsigned __int64 *pmax_bytes_to_load)
{
	for (LPTSTR next, cp = aOptions; cp && *(cp = omit_leading_whitespace(cp)); cp = next)
	{
		if (*cp == '\n')
		{
			translate_crlf_to_lf = true;
			// Rather than treating "`nxxx" as invalid or ignoring "xxx", let the delimiter be
			// optional for `n.  Treating "`nxxx" and "m1024`n" and "utf-8`n" as invalid would
			// require larger code, and would produce confusing error messages because the `n
			// isn't visible; e.g. "Invalid option. Specifically: utf-8"
			next = cp + 1; 
			continue;
		}
		// \n is included below to allow "m1024`n" and "utf-8`n" (see above).
		next = StrChrAny(cp, _T(" \t\n"));

		switch (ctoupper(*cp))
		{
		case 'M':
			if (pmax_bytes_to_load) // i.e. caller is FileRead.
			{
				*pmax_bytes_to_load = ATOU64(cp + 1); // Relies upon the fact that it ceases conversion upon reaching a space or tab.
				break;
			}
			// Otherwise, fall through to treat it as invalid:
		default:
			TCHAR name[12]; // Large enough for any valid encoding.
			if (next && (next - cp) < _countof(name))
			{
				// Create a temporary null-terminated copy.
				wmemcpy(name, cp, next - cp);
				name[next - cp] = '\0';
				cp = name;
			}
			if (!_tcsicmp(cp, _T("Raw")))
			{
				codepage = -1;
			}
			else
			{
				codepage = Line::ConvertFileEncoding(cp);
				if (codepage == -1 || cisdigit(*cp)) // Require "cp" prefix in FileRead/FileAppend options.
					return aResultToken.ValueError(ERR_INVALID_OPTION, cp);
			}
			break;
		} // switch()
	} // for()
	return OK;
}

BIF_DECL(BIF_FileRead)
{
	_f_param_string(aFilespec, 0);
	_f_param_string_opt(aOptions, 1);

	const DWORD DWORD_MAX = ~0;

	// Set default options:
	bool translate_crlf_to_lf = false;
	unsigned __int64 max_bytes_to_load = ULLONG_MAX; // By default, fail if the file is too large.  See comments near bytes_to_read below.
	UINT codepage = g->Encoding;

	if (!ConvertFileOptions(aResultToken, aOptions, codepage, translate_crlf_to_lf, &max_bytes_to_load))
		return; // It already displayed the error.

	_f_set_retval_p(_T(""), 0); // Set default.

	// It seems more flexible to allow other processes to read and write the file while we're reading it.
	// For example, this allows the file to be appended to during the read operation, which could be
	// desirable, especially it's a very large log file that would take a long time to read.
	// MSDN: "To enable other processes to share the object while your process has it open, use a combination
	// of one or more of [FILE_SHARE_READ, FILE_SHARE_WRITE]."
	HANDLE hfile = CreateFile(aFilespec, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING
		, FILE_FLAG_SEQUENTIAL_SCAN, NULL); // MSDN says that FILE_FLAG_SEQUENTIAL_SCAN will often improve performance
	if (hfile == INVALID_HANDLE_VALUE)      // in cases like these (and it seems best even if max_bytes_to_load was specified).
	{
		aResultToken.SetLastErrorMaybeThrow(true);
		return;
	}

	unsigned __int64 bytes_to_read = GetFileSize64(hfile);
	if (bytes_to_read == ULLONG_MAX) // GetFileSize64() failed.
	{
		aResultToken.SetLastErrorCloseAndMaybeThrow(hfile, true);
		return;
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
	{
		CloseHandle(hfile);
		_f_throw_oom; // Using this instead of "File too large." to reduce code size, since this condition is very rare (and malloc succeeding would be even rarer).
	}

	if (!bytes_to_read && codepage != -1) // In RAW mode, always return a zero-byte Buffer.
	{
		aResultToken.SetLastErrorCloseAndMaybeThrow(hfile, false, 0); // Indicate success (a zero-length file results in an empty string).
		return;
	}

	LPBYTE output_buf = (LPBYTE)malloc(size_t(bytes_to_read + (bytes_to_read & 1) + sizeof(wchar_t)));
	if (!output_buf)
	{
		CloseHandle(hfile);
		_f_throw_oom;
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
		if (codepage != -1) // Text mode, not "RAW" mode.
		{
			codepage &= CP_AHKCP; // Convert to plain Win32 codepage (remove CP_AHKNOBOM, which has no meaning here).
			bool has_bom;
			if ( (has_bom = (bytes_actually_read >= 2 && output_buf[0] == 0xFF && output_buf[1] == 0xFE)) // UTF-16LE BOM
					|| codepage == CP_UTF16 ) // Covers FileEncoding UTF-16 and FileEncoding UTF-16-RAW.
			{
				#ifndef UNICODE
				#error FileRead UTF-16 to ANSI string not implemented.
				#endif
				LPWSTR text = (LPWSTR)output_buf;
				DWORD length = bytes_actually_read / sizeof(WCHAR);
				if (has_bom)
				{
					// Move the data to eliminate the byte order mark.
					// Seems likely to perform better than allocating new memory and copying to it.
					--length;
					wmemmove(text, text + 1, length);
				}
				text[length] = '\0'; // Ensure text is terminated where indicated.  Two bytes were reserved for this purpose.
				aResultToken.AcceptMem(text, length);
				output_buf = NULL; // Don't free it; caller will take over.
			}
			else
			{
				LPCSTR text = (LPCSTR)output_buf;
				DWORD length = bytes_actually_read;
				if (length >= 3 && output_buf[0] == 0xEF && output_buf[1] == 0xBB && output_buf[2] == 0xBF) // UTF-8 BOM
				{
					codepage = CP_UTF8;
					length -= 3;
					text += 3;
				}
#ifndef UNICODE
				if (codepage == CP_ACP || codepage == GetACP())
				{
					// Avoid any unnecessary conversion or copying by using our malloc'd buffer directly.
					// This should be worth doing since the string must otherwise be converted to UTF-16 and back.
					output_buf[bytes_actually_read] = 0; // Ensure text is terminated where indicated.
					aResultToken.AcceptMem((LPSTR)output_buf, bytes_actually_read);
					output_buf = NULL; // Don't free it; caller will take over.
				}
				else
				#error FileRead non-ACP-ANSI to ANSI string not fully implemented.
#endif
				{
					int wlen = MultiByteToWideChar(codepage, 0, text, length, NULL, 0);
					if (wlen > 0)
					{
						if (!TokenSetResult(aResultToken, NULL, wlen))
						{
							free(output_buf);
							return;
						}
						wlen = MultiByteToWideChar(codepage, 0, text, length, aResultToken.marker, wlen);
						aResultToken.symbol = SYM_STRING;
						aResultToken.marker[wlen] = 0;
						aResultToken.marker_length = wlen;
						if (!wlen)
							result = FALSE;
					}
				}
			}
			if (output_buf) // i.e. it wasn't "claimed" above.
				free(output_buf);
			if (translate_crlf_to_lf && aResultToken.marker_length)
			{
				// Since a larger string is being replaced with a smaller, there's a good chance the 2 GB
				// address limit will not be exceeded by StrReplace even if the file is close to the
				// 1 GB limit as described above:
				StrReplace(aResultToken.marker, _T("\r\n"), _T("\n"), SCS_SENSITIVE, UINT_MAX, -1, NULL, &aResultToken.marker_length);
			}
		}
		else // codepage == -1 ("RAW" mode)
		{
			// Return the buffer to our caller.
			aResultToken.Return(BufferObject::Create(output_buf, bytes_actually_read));
		}
	}
	else
	{
		// ReadFile() failed.  Since MSDN does not document what is in the buffer at this stage, or
		// whether bytes_to_read contains a valid value, it seems best to abort the entire operation
		// rather than try to return partial file contents.  An exception will indicate the failure.
		free(output_buf);
	}

	aResultToken.SetLastErrorMaybeThrow(!result);
}



BIF_DECL(BIF_FileAppend)
{
	size_t aBuf_length;
	_f_param_string(aBuf, 0, &aBuf_length);
	_f_param_string_opt(aFilespec, 1);
	_f_param_string_opt(aOptions, 2);

	IObject *aBuf_obj = ParamIndexToObject(0); // Allow a Buffer-like object.
	if (aBuf_obj)
	{
		size_t ptr;
		GetBufferObjectPtr(aResultToken, aBuf_obj, ptr, aBuf_length);
		if (aResultToken.Exited())
			return;
		aBuf = (LPTSTR)ptr;
	}
	else
		aBuf_length *= sizeof(TCHAR); // Convert to byte count.

	_f_set_retval_p(_T(""), 0); // For all non-throw cases.

	// The below is avoided because want to allow "nothing" to be written to a file in case the
	// user is doing this to reset it's timestamp (or create an empty file).
	//if (!aBuf || !*aBuf)
	//	_f_return_retval;

	// Use the read-file loop's current item if filename was explicitly left blank (i.e. not just
	// a reference to a variable that's blank):
	LoopReadFileStruct *aCurrentReadFile = (aParamCount < 2) ? g->mLoopReadFile : NULL;
	if (aCurrentReadFile)
		aFilespec = aCurrentReadFile->mWriteFileName;
	if (!*aFilespec) // Nothing to write to.
		_f_throw_value(ERR_PARAM2_MUST_NOT_BE_BLANK);

	TextStream *ts = aCurrentReadFile ? aCurrentReadFile->mWriteFile : NULL;
	bool file_was_already_open = ts;

#ifdef CONFIG_DEBUGGER
	if (*aFilespec == '*' && !aFilespec[1] && !aBuf_obj && g_Debugger.OutputStdOut(aBuf))
	{
		// StdOut has been redirected to the debugger, and this "FileAppend" call has been
		// fully handled by the call above, so just return.
		g->LastError = 0;
		_f_return_retval;
	}
#endif

	UINT codepage;

	// Check if the file needs to be opened.  This is done here rather than at the time the
	// loop first begins so that:
	// 1) Any options/encoding specified in the first FileAppend call can take effect.
	// 2) To avoid opening the file if the file-reading loop has zero iterations (i.e. it's
	//    opened only upon first actual use to help performance and avoid changing the
	//    file-modification time when no actual text will be appended).
	if (!file_was_already_open)
	{
		codepage = aBuf_obj ? -1 : g->Encoding; // Never default to BOM if a Buffer object was passed.
		bool translate_crlf_to_lf = false;
		if (!ConvertFileOptions(aResultToken, aOptions, codepage, translate_crlf_to_lf, NULL))
			return;

		DWORD flags = TextStream::APPEND | (translate_crlf_to_lf ? TextStream::EOL_CRLF : 0);
		
		ASSERT( (~CP_AHKNOBOM) == CP_AHKCP );
		// codepage may include CP_AHKNOBOM, in which case below will not add BOM_UTFxx flag.
		if (codepage == CP_UTF8)
			flags |= TextStream::BOM_UTF8;
		else if (codepage == CP_UTF16)
			flags |= TextStream::BOM_UTF16;
		else if (codepage != -1)
			codepage &= CP_AHKCP;

		// Open the output file (if one was specified).  Unlike the input file, this is not
		// a critical error if it fails.  We want it to be non-critical so that FileAppend
		// commands in the body of the loop will throw to indicate the problem:
		ts = new TextFile; // ts was already verified NULL via !file_was_already_open.
		if ( !ts->Open(aFilespec, flags, codepage) )
		{
			aResultToken.SetLastErrorMaybeThrow(true);
			delete ts; // Must be deleted explicitly!
			return;
		}
		if (aCurrentReadFile)
			aCurrentReadFile->mWriteFile = ts;
	}
	else
		codepage = ts->GetCodePage();

	// Write to the file:
	DWORD result = 1;
	if (aBuf_length)
	{
		if (codepage == -1 || aBuf_obj) // "RAW" mode.
			result = ts->Write((LPCVOID)aBuf, (DWORD)aBuf_length);
		else
			result = ts->Write(aBuf, DWORD(aBuf_length / sizeof(TCHAR)));
	}
	//else: aBuf is empty; we've already succeeded in creating the file and have nothing further to do.
	aResultToken.SetLastErrorMaybeThrow(result == 0);

	if (!aCurrentReadFile)
		delete ts;
	// else it's the caller's responsibility, or it's caller's, to close it.
}



BOOL FileDeleteCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData)
{
	return DeleteFile(aFilename);
}

ResultType Line::FileDelete(LPTSTR aFilePattern)
{
	if (!*aFilePattern)
		return LineError(ERR_PARAM1_INVALID, FAIL_OR_OK);

	// The no-wildcard case could be handled via FilePatternApply(), but handling it this
	// way ensures deleting a non-existent path without wildcards is considered a failure:
	if (!StrChrAny(aFilePattern, _T("?*"))) // No wildcards; just a plain path/filename.
	{
		SetLastError(0); // For sanity: DeleteFile appears to set it only on failure.
		return SetLastErrorMaybeThrow(!DeleteFile(aFilePattern));
	}

	// Otherwise aFilePattern contains wildcards, so we'll search for all matches and delete them.
	return FilePatternApply(aFilePattern, FILE_LOOP_FILES_ONLY, false, FileDeleteCallback, NULL);
}



ResultType Line::FileInstall(LPTSTR aSource, LPTSTR aDest, LPTSTR aFlag)
{
	bool success;
	bool allow_overwrite = (ATOI(aFlag) == 1);
#ifndef AUTOHOTKEYSC
	if (g_script.mKind != Script::ScriptKindResource)
		success = FileInstallCopy(aSource, aDest, allow_overwrite);
	else
#endif
		success = FileInstallExtract(aSource, aDest, allow_overwrite);
	return ThrowIfTrue(!success);
}

bool Line::FileInstallExtract(LPTSTR aSource, LPTSTR aDest, bool aOverwrite)
{
	// Open the file first since it's the most likely to fail:
	HANDLE hfile = CreateFile(aDest, GENERIC_WRITE, 0, NULL, aOverwrite ? CREATE_ALWAYS : CREATE_NEW, 0, NULL);
	if (hfile == INVALID_HANDLE_VALUE)
		return false;

	// Create a temporary copy of aSource to ensure it is the correct case (upper-case).
	// Ahk2Exe converts it to upper-case before adding the resource. My testing showed that
	// using lower or mixed case in some instances prevented the resource from being found.
	// Since file paths are case-insensitive, it certainly doesn't seem harmful to do this:
	TCHAR source[T_MAX_PATH];
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
	bool success = false;
	if ( (res = FindResource(NULL, source, RT_RCDATA))
	  && (res_load = LoadResource(NULL, res))
	  && (res_lock = LockResource(res_load))  )
	{
		DWORD num_bytes_written;
		// Write the resource data to file.
		success = WriteFile(hfile, res_lock, SizeofResource(NULL, res), &num_bytes_written, NULL);
	}
	CloseHandle(hfile);
	return success;
}

#ifndef AUTOHOTKEYSC
bool Line::FileInstallCopy(LPTSTR aSource, LPTSTR aDest, bool aOverwrite)
{
	// v1.0.35.11: Must search in A_ScriptDir by default because that's where ahk2exe will search by default.
	// The old behavior was to search in A_WorkingDir, which seems pointless because ahk2exe would never
	// be able to use that value if the script changes it while running.
	TCHAR source_path[T_MAX_PATH], dest_path[T_MAX_PATH];
	GetFullPathName(aDest, _countof(dest_path), dest_path, NULL);
	// Avoid attempting the copy if both paths are the same (since it would fail with ERROR_SHARING_VIOLATION),
	// but resolve both to full paths in case mFileDir != g_WorkingDir.  There is a more thorough way to detect
	// when two *different* paths refer to the same file, but it doesn't work with different network shares, and
	// the additional complexity wouldn't be warranted.  Also, the limitations of this method are clearer.
	SetCurrentDirectory(g_script.mFileDir);
	GetFullPathName(aSource, _countof(source_path), source_path, NULL);
	SetCurrentDirectory(g_WorkingDir); // Restore to proper value.
	if (!lstrcmpi(source_path, dest_path) // Full paths are equal.
		&& !(GetFileAttributes(source_path) & FILE_ATTRIBUTE_DIRECTORY)) // Source file exists and is not a directory (otherwise, an error should be thrown).
		return true;

	return CopyFile(source_path, dest_path, !aOverwrite);
}
#endif



BIF_DECL(BIF_FileGetAttrib)
{
	_f_param_string_opt_def(aFilespec, 0, (g->mLoopFile ? g->mLoopFile->cFileName : _T("")));

	if (!*aFilespec)
		_f_throw_value(ERR_PARAM2_MUST_NOT_BE_BLANK);

	DWORD attr = GetFileAttributes(aFilespec);
	if (attr == 0xFFFFFFFF)  // Failure, probably because file doesn't exist.
	{
		aResultToken.SetLastErrorMaybeThrow(true);
		return;
	}

	g->LastError = 0;
	_f_return_p(FileAttribToStr(_f_retval_buf, attr));
}



BOOL FileSetAttribCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData);
struct FileSetAttribData
{
	DWORD and_mask, xor_mask;
};

ResultType Line::FileSetAttrib(LPTSTR aAttributes, LPTSTR aFilePattern
	, FileLoopModeType aOperateOnFolders, bool aDoRecurse)
// Returns the number of files and folders that could not be changed due to an error.
{
	if (!*aFilePattern)
		return LineError(ERR_PARAM2_INVALID, FAIL_OR_OK);

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
		case ' ':
		case '\t':
			continue;
		default:
			return LineError(ERR_PARAM1_INVALID, FAIL_OR_OK, cp);
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
		default: // No +/-/^ specified, so overwrite attributes (equal and opposite to FileGetAttrib).
			attrib.and_mask = 0; // Reset all bits to 0.
			attrib.xor_mask |= mask; // Set bit to 1.  |= to accumulate if multiple attributes are present.
			break;
		}
	}
	FilePatternApply(aFilePattern, aOperateOnFolders, aDoRecurse, FileSetAttribCallback, &attrib);
	return OK;
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



ResultType Line::FilePatternApply(LPTSTR aFilePattern, FileLoopModeType aOperateOnFolders
	, bool aDoRecurse, FilePatternCallback aCallback, void *aCallbackData)
{
	if (!*aFilePattern)
		// Caller should handle this case before calling us if an exception is to be thrown.
		return SetLastErrorMaybeThrow(true, ERROR_INVALID_PARAMETER);

	if (aOperateOnFolders == FILE_LOOP_INVALID) // In case runtime dereference of a var was an invalid value.
		aOperateOnFolders = FILE_LOOP_FILES_ONLY;  // Set default.
	g->LastError = 0; // Set default. Overridden only when a failure occurs.

	FilePatternStruct fps;

	LPTSTR last_backslash = _tcsrchr(aFilePattern, '\\');
	if (last_backslash)
		fps.dir_length = last_backslash - aFilePattern + 1; // Include the slash.
	else // Use current working directory, e.g. if user specified only *.*
		fps.dir_length = 0;
	fps.pattern_length = _tcslen(aFilePattern + fps.dir_length);
	
	// Testing shows that the ANSI version of FindFirstFile() will not accept a path+pattern longer
	// than 259, even if the pattern would match files whose names are short enough to be legal.
	if (fps.dir_length + fps.pattern_length >= _countof(fps.path)
		|| fps.pattern_length >= _countof(fps.pattern))
		return SetLastErrorMaybeThrow(true, ERROR_BUFFER_OVERFLOW);

	// Make copies in case of overwrite of deref buf during LONG_OPERATION/MsgSleep,
	// and to allow modification:
	_tcscpy(fps.path, aFilePattern); // Include the pattern initially.
	_tcscpy(fps.pattern, aFilePattern + fps.dir_length); // Just the naked filename or pattern, for use with aDoRecurse.

	if (!StrChrAny(fps.pattern, _T("?*")))
		// Since no wildcards, always operate on this single item even if it's a folder.
		aOperateOnFolders = FILE_LOOP_FILES_AND_FOLDERS;

	// Passing the parameters this way reduces code size:
	fps.aCallback = aCallback;
	fps.aCallbackData = aCallbackData;
	fps.aDoRecurse = aDoRecurse;
	fps.aOperateOnFolders = aOperateOnFolders;

	fps.failure_count = 0;

	FilePatternApply(fps);
	return ThrowIntIfNonzero(fps.failure_count); // i.e. indicate success if there were no failures.
}



void Line::FilePatternApply(FilePatternStruct &fps)
{
	size_t dir_length = fps.dir_length; // Length of this directory (saved before recursion).
	LPTSTR append_pos = fps.path + dir_length; // This is where the changing part gets appended.
	size_t space_remaining = _countof(fps.path) - dir_length - 1; // Space left in file_path for the changing part.

	LONG_OPERATION_INIT
	int failure_count = 0;
	WIN32_FIND_DATA current_file;
	HANDLE file_search = FindFirstFile(fps.path, &current_file);

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
					|| fps.aOperateOnFolders == FILE_LOOP_FILES_ONLY)
					continue; // Never operate upon or recurse into these.
			}
			else // It's a file, not a folder.
				if (fps.aOperateOnFolders == FILE_LOOP_FOLDERS_ONLY)
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
			if (!fps.aCallback(fps.path, current_file, fps.aCallbackData))
				++failure_count;
			//
		} while (FindNextFile(file_search, &current_file));

		FindClose(file_search);
	} // if (file_search != INVALID_HANDLE_VALUE)

	if (fps.aDoRecurse && space_remaining > 1) // The space_remaining check ensures there's enough room to append "*", though if false, that would imply lfs.pattern is empty.
	{
		_tcscpy(append_pos, _T("*")); // Above has ensured this won't overflow.
		file_search = FindFirstFile(fps.path, &current_file);

		if (file_search != INVALID_HANDLE_VALUE)
		{
			do
			{
				LONG_OPERATION_UPDATE
				if (!(current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
					|| current_file.cFileName[0] == '.' && (!current_file.cFileName[1]      // Relies on short-circuit boolean order.
						|| current_file.cFileName[1] == '.' && !current_file.cFileName[2])) //
					continue;
				size_t filename_length = _tcslen(current_file.cFileName);
				// v1.0.45.03: Skip over folders whose paths are too long to be supported by FindFirst.
				if (fps.pattern_length + filename_length >= space_remaining) // >= vs. > to reserve 1 for the backslash to be added between cFileName and naked_filename_or_pattern.
					continue; // Never recurse into these.
				// This will build the string CurrentDir+SubDir+FilePatternOrName.
				// If FilePatternOrName doesn't contain a wildcard, the recursion
				// process will attempt to operate on the originally-specified
				// single filename or folder name if it occurs anywhere else in the
				// tree, e.g. recursing C:\Temp\temp.txt would affect all occurrences
				// of temp.txt both in C:\Temp and any subdirectories it might contain:
				_stprintf(append_pos, _T("%s\\%s") // Above has ensured this won't overflow.
					, current_file.cFileName, fps.pattern);
				fps.dir_length = dir_length + filename_length + 1; // Include the slash.
				//
				// Apply the callback to files in this subdirectory:
				FilePatternApply(fps);
				//
			} while (FindNextFile(file_search, &current_file));
			FindClose(file_search);
		} // if (file_search != INVALID_HANDLE_VALUE)
	} // if (aDoRecurse)

	fps.failure_count += failure_count; // Update failure count (produces smaller code than doing ++fps.failure_count directly).
}



BIF_DECL(BIF_FileGetTime)
{
	_f_param_string_opt_def(aFilespec, 0, (g->mLoopFile ? g->mLoopFile->cFileName : _T("")));
	_f_param_string_opt(aWhichTime, 1);

	if (!*aFilespec)
		_f_throw_value(ERR_PARAM2_MUST_NOT_BE_BLANK);

	// Don't use CreateFile() & FileGetSize() size they will fail to work on a file that's in use.
	// Research indicates that this method has no disadvantages compared to the other method.
	WIN32_FIND_DATA found_file;
	HANDLE file_search = FindFirstFile(aFilespec, &found_file);
	if (file_search == INVALID_HANDLE_VALUE)
	{
		aResultToken.SetLastErrorMaybeThrow(true);
		return;
	}
	FindClose(file_search);

	FILETIME local_file_time;
	switch (ctoupper(*aWhichTime))
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

	g->LastError = 0;
	_f_return_p(FileTimeToYYYYMMDD(_f_retval_buf, local_file_time));
}



BOOL FileSetTimeCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData);
struct FileSetTimeData
{
	FILETIME Time;
	TCHAR WhichTime;
};

ResultType Line::FileSetTime(LPTSTR aYYYYMMDD, LPTSTR aFilePattern, TCHAR aWhichTime
	, FileLoopModeType aOperateOnFolders, bool aDoRecurse)
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
		if (   !YYYYMMDDToFileTime(yyyymmdd, ft)  // Convert the arg into the time struct as local (non-UTC) time.
			|| !LocalFileTimeToFileTime(&ft, &callbackData.Time)   )  // Convert from local to UTC.
		{
			// Invalid parameters are the only likely cause of this condition.
			return LineError(ERR_PARAM1_INVALID, FAIL_OR_OK, aYYYYMMDD);
		}
	}
	else // User wants to use the current time (i.e. now) as the new timestamp.
		GetSystemTimeAsFileTime(&callbackData.Time);

	if (!*aFilePattern)
		return LineError(ERR_PARAM2_INVALID, FAIL_OR_OK);

	FilePatternApply(aFilePattern, aOperateOnFolders, aDoRecurse, FileSetTimeCallback, &callbackData);
	return OK;
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



BIF_DECL(BIF_FileGetSize)
{
	_f_param_string_opt_def(aFilespec, 0, (g->mLoopFile ? g->mLoopFile->cFileName : _T("")));
	_f_param_string_opt(aGranularity, 1);

	if (!*aFilespec)
		_f_throw_value(ERR_PARAM2_MUST_NOT_BE_BLANK); // Throw an error, since this is probably not what the user intended.
	
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
		{
			aResultToken.SetLastErrorMaybeThrow(true);
			_f_return_empty;
		}
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

	g->LastError = 0;
	_f_return(size);
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

DECLSPEC_NOINLINE // Lexikos: noinline saves ~300 bytes.  Originally the duplicated code prevented inlining.
HWND Line::DetermineTargetWindow(LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	// Lexikos: Not sure why these checks were duplicated here and in WinExist(),
	// so they're left here for reference:
	//HWND target_window; // A variable of this name is used by the macros below.
	//IF_USE_FOREGROUND_WINDOW(g->DetectHiddenWindows, aTitle, aText, aExcludeTitle, aExcludeText)
	//else if (*aTitle || *aText || *aExcludeTitle || *aExcludeText)
	//	target_window = WinExist(*g, aTitle, aText, aExcludeTitle, aExcludeText);
	//else // Use the "last found" window.
	//	target_window = GetValidLastUsedWindow(*g);
	return WinExist(*g, aTitle, aText, aExcludeTitle, aExcludeText);
}


ResultType GetObjectIntProperty(IObject *aObject, LPTSTR aPropName, __int64 &aValue, ResultToken &aResultToken, bool aOptional)
{
	FuncResult result_token;
	ExprTokenType this_token = aObject;

	auto result = aObject->Invoke(result_token, IT_GET, aPropName, this_token, nullptr, 0);

	if (result_token.symbol != SYM_INTEGER)
	{
		result_token.Free();
		if (result == FAIL || result == EARLY_EXIT)
		{
			aResultToken.SetExitResult(result);
			return FAIL;
		}
		if (result != INVOKE_NOT_HANDLED) // Property exists but is not an integer.
			return aResultToken.Error(ERR_TYPE_MISMATCH, aPropName, ErrorPrototype::Type);
		//aValue = 0; // Caller should set default value for these cases.
		if (!aOptional)
			return aResultToken.UnknownMemberError(ExprTokenType(aObject), IT_GET, aPropName);
		return result; // Let caller know it wasn't found.
	}

	aValue = result_token.value_int64;
	return OK;
}

ResultType SetObjectIntProperty(IObject *aObject, LPTSTR aPropName, __int64 aValue, ResultToken &aResultToken)
{
	FuncResult result_token;
	ExprTokenType this_token = aObject, value_token = aValue, *param = &value_token;

	auto result = aObject->Invoke(result_token, IT_SET, aPropName, this_token, &param, 1);

	result_token.Free();
	if (result == FAIL || result == EARLY_EXIT)
		return aResultToken.SetExitResult(result);
	if (result == INVOKE_NOT_HANDLED)
		return aResultToken.UnknownMemberError(ExprTokenType(aObject), IT_GET, aPropName);
	return OK;
}

ResultType GetObjectPtrProperty(IObject *aObject, LPTSTR aPropName, UINT_PTR &aPtr, ResultToken &aResultToken, bool aOptional)
{
	__int64 value = NULL;
	auto result = GetObjectIntProperty(aObject, aPropName, value, aResultToken, aOptional);
	aPtr = (UINT_PTR)value;
	return result;
}


ResultType DetermineTargetHwnd(HWND &aWindow, ResultToken &aResultToken, ExprTokenType &aToken)
{
	__int64 n = NULL;
	if (IObject *obj = TokenToObject(aToken))
	{
		if (!GetObjectIntProperty(obj, _T("Hwnd"), n, aResultToken))
			return FAIL;
	}
	else if (TokenIsPureNumeric(aToken) == PURE_INTEGER)
		n = TokenToInt64(aToken);
	else
		return CONDITION_FALSE;
	aWindow = (HWND)(UINT_PTR)n;
	// Callers expect the return value to be either a valid HWND or 0:
	if (!IsWindow(aWindow))
		aWindow = 0;
	return OK;
}


ResultType DetermineTargetWindow(HWND &aWindow, ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, int aNonWinParamCount)
{
	if (aParamCount > 0)
	{
		auto result = DetermineTargetHwnd(aWindow, aResultToken, *aParam[0]);
		if (result != CONDITION_FALSE)
		{
			if (result == OK && !aWindow)
				return aResultToken.Error(ERR_NO_WINDOW, ErrorPrototype::Target);
			return result;
		}
	}
	TCHAR number_buf[4][MAX_NUMBER_SIZE];
	LPTSTR param[4];
	for (int i = 0, j = 0; i < 4; ++i, ++j)
	{
		if (i == 2) j += aNonWinParamCount;
		param[i] = j < aParamCount ? TokenToString(*aParam[j], number_buf[i]) : _T("");
	}
	aWindow = Line::DetermineTargetWindow(param[0], param[1], param[2], param[3]);
	if (aWindow)
		return OK;
	return aResultToken.Error(ERR_NO_WINDOW, param[0], ErrorPrototype::Target);
}


ResultType DetermineTargetControl(HWND &aControl, HWND &aWindow, ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount
	, int aNonWinParamCount, bool aThrowIfNotFound)
{
	aWindow = aControl = nullptr;
	// Only functions which can operate on top-level windows allow Control to be
	// omitted (and a select few other functions with more optional parameters).
	// This replaces the old "ahk_parent" string used with ControlSend, but is
	// also used by SendMessage.
	LPTSTR control_spec = nullptr;
	if (!ParamIndexIsOmitted(0))
	{
		switch (DetermineTargetHwnd(aWindow, aResultToken, *aParam[0]))
		{
		case OK:
			aControl = aWindow;
			if (!aControl)
				return aResultToken.Error(ERR_NO_CONTROL, ErrorPrototype::Target);
			return OK;
		case FAIL:
			return FAIL;
		}
		// Since above didn't return, it wasn't a pure Integer or object {Hwnd}.
		control_spec = ParamIndexToString(0, _f_number_buf);
	}
	if (!DetermineTargetWindow(aWindow, aResultToken, aParam + 1, aParamCount - 1, aNonWinParamCount))
		return FAIL;
	aControl = control_spec ? ControlExist(aWindow, control_spec) : aWindow;
	if (!aControl && aThrowIfNotFound)
		return aResultToken.Error(ERR_NO_CONTROL, ErrorPrototype::Target);
	return OK;
}



bool Line::FileIsFilteredOut(LoopFilesStruct &aCurrentFile, FileLoopModeType aFileLoopMode)
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

	// Since file was found, also append the file's name to its directory for the caller:
	// Seems best to check length in advance because it allows a faster move/copy method further below
	// (in lieu of sntprintf(), which is probably quite a bit slower than the method here).
	size_t name_length = _tcslen(aCurrentFile.cFileName);
	if (aCurrentFile.dir_length + name_length >= _countof(aCurrentFile.file_path)) // Should be impossible with current buffer sizes.
		return true; // Exclude this file/folder.
	tmemcpy(aCurrentFile.file_path + aCurrentFile.dir_length, aCurrentFile.cFileName, name_length + 1); // +1 to include the terminator.
	aCurrentFile.file_path_length = aCurrentFile.dir_length + name_length;
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
	// If g->CurrentFunc, label is never outside the function since it would not
	// have been found by FindLabel().  So there's no need to check for that here.
	if (!aIsDereferenced)
		mRelatedLine = (Line *)label; // The script loader has ensured that label->mJumpToLine isn't NULL.
	// else don't update it, because that would permanently resolve the jump target, and we want it to stay dynamic.
	return IsJumpValid(*label);
	// Any error msg was already displayed by the above call.
}



Label *Line::IsJumpValid(Label &aTargetLabel, bool aSilent)
// Returns aTargetLabel is the jump is valid, or NULL otherwise.
{
	// aTargetLabel can be NULL if this Goto's target is the physical end of the script.
	// And such a destination is always valid, regardless of where aOrigin is.
	// UPDATE: It's no longer possible for the destination of a Goto to be
	// NULL because the script loader has ensured that the end of the script always has
	// an extra ACT_EXIT that serves as an anchor for any final labels in the script:
	//if (aTargetLabel == NULL)
	//	return OK;
	// The above check is also necessary to avoid dereferencing a NULL pointer below.

	if (!CheckValidFinallyJump(aTargetLabel.mJumpToLine, aSilent))
		return NULL;

	Line *parent_line_of_label_line;
	if (   !(parent_line_of_label_line = aTargetLabel.mJumpToLine->mParentLine)   )
		// A Goto can always jump to a point anywhere in the outermost layer
		// (i.e. outside all blocks) without restriction (except from inside a
		// function to outside, but in that case the label would not be found):
		return &aTargetLabel; // Indicate success.

	// So now we know this Goto is attempting to jump into a block somewhere.  Is that
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
		LineError(_T("A Goto must not jump into a block that doesn't enclose it."));
	return NULL;
}


BOOL Line::CheckValidFinallyJump(Line* jumpTarget, bool aSilent)
{
	Line* jumpParent = jumpTarget->mParentLine;
	for (Line *ancestor = mParentLine; ancestor != NULL; ancestor = ancestor->mParentLine)
	{
		if (ancestor == jumpParent)
			return TRUE; // We found the common ancestor.
		if (ancestor->mActionType == ACT_FINALLY)
		{
			if (!aSilent)
				LineError(ERR_BAD_JUMP_INSIDE_FINALLY);
			return FALSE; // The common ancestor is outside the FINALLY block and thus this jump is invalid.
		}
	}
	return TRUE; // The common ancestor is the root of the script.
}


////////////////////////
// BUILT-IN VARIABLES //
////////////////////////


BIV_DECL_R(BIV_Clipboard)
{
	auto length = g_clip.Get();
	if (TokenSetResult(aResultToken, nullptr, length))
	{
		aResultToken.marker_length = g_clip.Get(aResultToken.marker);
		if (aResultToken.marker_length == CLIPBOARD_FAILURE)
			aResultToken.SetExitResult(FAIL);
		aResultToken.symbol = SYM_STRING;
	}
	g_clip.Close();
}

BIV_DECL_W(BIV_Clipboard_Set)
{
	if (auto *obj = BivRValueToObject())
	{
		if (ClipboardAll *cba = dynamic_cast<ClipboardAll *>(obj))
		{
			if (!Var::SetClipboardAll(cba->Data(), cba->Size()))
				_f_return_FAIL;
			return;
		}
		_f_throw_type(_T("ClipboardAll"), obj->Type());
	}
	size_t aLength;
	LPTSTR aBuf = BivRValueToString(&aLength);
	if (!g_clip.Set(aBuf, aLength))
		_f_return_FAIL;
}


BIV_DECL_R(BIV_MMM_DDD)
{
	LPTSTR format_str;
	switch(ctoupper(aVarName[2]))
	{
	// Use the case-sensitive formats required by GetDateFormat():
	case 'M': format_str = (aVarName[5] ? _T("MMMM") : _T("MMM")); break;
	case 'D': format_str = (aVarName[5] ? _T("dddd") : _T("ddd")); break;
	}
	// Confirmed: The below will automatically use the local time (not UTC) when 3rd param is NULL.
	int len = GetDateFormatEx(LOCALE_NAME_USER_DEFAULT, 0, nullptr, format_str, _f_retval_buf, _f_retval_buf_size, nullptr);
	if (len && !_f_retval_buf[len - 1])
		--len; // "Returns the number of characters written" apparently includes the terminator.
	_f_return_p(_f_retval_buf, len);
}

VarSizeType GetDateTimeBIV(LPTSTR aBuf, LPTSTR aVarName)
{
	if (!aBuf)
		return 6; // Since only an estimate is needed in this mode, return the maximum length of any item.

	aVarName += 2; // Skip past the "A_".

	// The current time is refreshed only if it's been a certain number of milliseconds since
	// the last fetch of one of these built-in time variables.  This keeps the variables in
	// sync with one another when they are used consecutively such as this example:
	// Var := A_Hour ':' A_Min ':' A_Sec
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

BIV_DECL_R(BIV_DateTime)
{
	GetDateTimeBIV(_f_retval_buf, aVarName);
	_f_return_p(_f_retval_buf);
}

BIV_DECL_R(BIV_ListLines)
{
	_f_return_b(g->ListLinesIsEnabled);
}

BIV_DECL_W(BIV_ListLines_Set)
{
	g->ListLinesIsEnabled = BivRValueToBOOL();
}

BIV_DECL_R(BIV_TitleMatchMode)
{
	if (g->TitleMatchMode == FIND_REGEX) // v1.0.45.
		// For backward compatibility (due to possible use of case-sensitive comparison), never change the case used here:
		_f_return_p(_T("RegEx"));
	// Otherwise, it's a numerical mode:
	_f_return_i(g->TitleMatchMode);
}

BIV_DECL_W(BIV_TitleMatchMode_Set)
{
	LPTSTR aBuf = BivRValueToString();
	switch (Line::ConvertTitleMatchMode(aBuf))
	{
	case FIND_IN_LEADING_PART: g->TitleMatchMode = FIND_IN_LEADING_PART; break;
	case FIND_ANYWHERE: g->TitleMatchMode = FIND_ANYWHERE; break;
	case FIND_REGEX: g->TitleMatchMode = FIND_REGEX; break;
	case FIND_EXACT: g->TitleMatchMode = FIND_EXACT; break;
	// For simplicity, this function handles both variables.
	case FIND_FAST: g->TitleFindFast = true; break;
	case FIND_SLOW: g->TitleFindFast = false; break;
	default:
		_f_throw_value(ERR_INVALID_VALUE, aBuf);
	}
}

BIV_DECL_R(BIV_TitleMatchModeSpeed)
{
	// For backward compatibility (due to possible use of case-sensitive comparison), never change the case used here:
	_f_return_p(g->TitleFindFast ? _T("Fast") : _T("Slow"));
}

BIV_DECL_R(BIV_DetectHiddenWindows)
{
	_f_return_b(g->DetectHiddenWindows);
}

BIV_DECL_W(BIV_DetectHiddenWindows_Set)
{
	g->DetectHiddenWindows = BivRValueToBOOL();
}

BIV_DECL_R(BIV_DetectHiddenText)
{
	_f_return_b(g->DetectHiddenText);
}

BIV_DECL_W(BIV_DetectHiddenText_Set)
{
	g->DetectHiddenText = BivRValueToBOOL();
}

int& BIV_xDelay(LPTSTR aVarName)
{
	global_struct &g = *::g; // Reduces code size.
	switch (ctoupper(aVarName[2])) // a_X...
	{
	case 'K':
		if (ctolower(aVarName[6]) == 'e') // a_keydE...
		{
			if (aVarName[10]) // a_keydelayP...
				return g.KeyDelayPlay;
			else
				return g.KeyDelay;
		}
		else // a_keydU...
		{
			if (aVarName[13]) // a_keydurationP...
				return g.PressDurationPlay;
			else
				return g.PressDuration;
		}
	case 'M':
		if (aVarName[12]) // a_mousedelayP...
			return g.MouseDelayPlay;
		else
			return g.MouseDelay;
	case 'W':
		return g.WinDelay;
	//case 'C':
	default:
		return g.ControlDelay;
	}
}

BIV_DECL_R(BIV_xDelay)
{
	_f_return_i(BIV_xDelay(aVarName));
}

BIV_DECL_W(BIV_xDelay_Set)
{
	BIV_xDelay(aVarName) = (int)BivRValueToInt64();
}

BIV_DECL_R(BIV_DefaultMouseSpeed)
{
	_f_return_i(g->DefaultMouseSpeed);
}

BIV_DECL_W(BIV_DefaultMouseSpeed_Set)
{
	g->DefaultMouseSpeed = (int)BivRValueToInt64();
}

BIV_DECL_R(BIV_CoordMode)
{
	static LPTSTR sCoordModes[] = COORD_MODES;
	_f_return_p(sCoordModes[(g->CoordMode >> Line::ConvertCoordModeCmd(aVarName + 11)) & COORD_MODE_MASK]);
}

BIV_DECL_W(BIV_CoordMode_Set)
{
	if (!Script::SetCoordMode(aVarName + 11, BivRValueToString())) // A_CoordMode is 11 chars.
		_f_return_FAIL;
}

BIV_DECL_R(BIV_SendMode)
{
	static LPTSTR sSendModes[] = SEND_MODES;
	_f_return_p(sSendModes[g->SendMode]);
}

BIV_DECL_W(BIV_SendMode_Set)
{
	if (!Script::SetSendMode(BivRValueToString()))
		_f_return_FAIL;
}

BIV_DECL_R(BIV_SendLevel)
{
	_f_return_i(g->SendLevel);
}

BIV_DECL_W(BIV_SendLevel_Set)
{
	if (!Script::SetSendLevel((int)BivRValueToInt64(), BivRValueToString()))
		_f_return_FAIL;
}

BIV_DECL_R(BIV_StoreCapsLockMode)
{
	_f_return_b(g->StoreCapslockMode);
}

BIV_DECL_W(BIV_StoreCapsLockMode_Set)
{
	g->StoreCapslockMode = BivRValueToBOOL();
}

BIV_DECL_R(BIV_Hotkey)
{
	if (aVarName[2] == 'M') // A_MaxHotkeysPerInterval
		_f_return_i(g_MaxHotkeysPerInterval);
	if (aVarName[8] == 'I') // A_HotkeyInterval
		_f_return_i(g_HotkeyThrottleInterval);
	// A_HotkeyModifierTimeout
	_f_return_i(g_HotkeyModifierTimeout);
}

BIV_DECL_W(BIV_Hotkey_Set)
{
	Throw_if_RValue_NaN();
	int value = (int)BivRValueToInt64();
	if (aVarName[2] == 'M') // A_MaxHotkeysPerInterval
	{
		if (value < 1)
			_f_throw_value(ERR_INVALID_VALUE);
		g_MaxHotkeysPerInterval = value;
	}
	else if (aVarName[8] == 'I') // A_HotkeyInterval
	{
		if (value < 0)
			_f_throw_value(ERR_INVALID_VALUE);
		g_HotkeyThrottleInterval = value;
	}
	else // A_HotkeyModifierTimeout
		g_HotkeyModifierTimeout = value;
}

BIV_DECL_R(BIV_MenuMaskKey)
{
	if (!g_MenuMaskKeyVK && !g_MenuMaskKeySC)
		_f_return_empty; // Return a "false" value to indicate there is no masking.
	// For uniformity, simplicity and to avoid any loss of information, always return vkNNscNNN.
	auto len = sntprintf(_f_retval_buf, _f_retval_buf_size, _T("vk%02Xsc%03X"), g_MenuMaskKeyVK, g_MenuMaskKeySC);
	_f_return_p(_f_retval_buf, len);
}

BIV_DECL_W(BIV_MenuMaskKey_Set)
{
	auto keyname = BivRValueToString();
	if (!*keyname) // Allow "" to mean "no masking".
	{
		g_MenuMaskKeyVK = 0;
		g_MenuMaskKeySC = 0;
		return;
	}
	vk_type vk;
	sc_type sc;
	// Testing shows that sending an event with zero VK but non-zero SC fails to suppress
	// the Start menu (although it does suppress the window menu).  However, allowing all
	// valid key strings seems more correct than requiring g_MenuMaskKeyVK != 0, and adds
	// flexibility at very little cost.  Note that this use of TextToVKandSC()'s return
	// value (vs. checking VK|SC) allows vk00sc000 to turn off masking altogether.
	if (!TextToVKandSC(keyname, vk, sc))
		_f_throw_value(ERR_INVALID_VALUE);
	g_MenuMaskKeyVK = vk;
	g_MenuMaskKeySC = sc;
}

BIV_DECL_R(BIV_IsPaused) // v1.0.48: Lexikos: Added BIV_IsPaused and BIV_IsCritical.
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
	// Checking g>g_array avoids any chance of underflow, which might otherwise happen if this is
	// called by the AutoExec section or a threadless callback running in thread #0.
	_f_return_b(g > g_array && g[-1].IsPaused);
}

BIV_DECL_R(BIV_IsCritical) // v1.0.48: Lexikos: Added BIV_IsPaused and BIV_IsCritical.
{
	// It seems more useful to return g->PeekFrequency than "On" or "Off" (ACT_CRITICAL ensures that
	// g->PeekFrequency!=0 whenever g->ThreadIsCritical==true).  Also, the word "Is" in "A_IsCritical"
	// implies a value that can be used as a boolean such as "if A_IsCritical".
	if (g->ThreadIsCritical)
		_f_return_i(g->PeekFrequency); // ACT_CRITICAL ensures that g->PeekFrequency > 0 when critical is on.
	// Otherwise:
	_f_return_i(0);
}

BIV_DECL_R(BIV_IsSuspended)
{
	_f_return_b(g_IsSuspended);
}



BIV_DECL_R(BIV_IsCompiled)
{
#ifdef AUTOHOTKEYSC
	_f_return_b(true);
#else
	_f_return_b(g_script.mKind == Script::ScriptKindResource);
#endif
}



BIV_DECL_R(BIV_FileEncoding)
{
	// A similar section may be found under "case Encoding:" in FileObject::Invoke.  Maintain that with this:
	LPTSTR enc;
	switch (g->Encoding)
	{
	// Returning readable strings for these seems more useful than returning their numeric values, especially with CP_AHKNOBOM:
	case CP_UTF8:                enc = _T("UTF-8");      break;
	case CP_UTF8 | CP_AHKNOBOM:  enc = _T("UTF-8-RAW");  break;
	case CP_UTF16:               enc = _T("UTF-16");     break;
	case CP_UTF16 | CP_AHKNOBOM: enc = _T("UTF-16-RAW"); break;
	default:
	  {
		enc = _f_retval_buf;
		enc[0] = _T('C');
		enc[1] = _T('P');
		_itot(g->Encoding, enc + 2, 10);
	  }
	}
	_f_return_p(enc);
}

BIV_DECL_W(BIV_FileEncoding_Set)
{
	LPTSTR aBuf = BivRValueToString();
	UINT new_encoding = Line::ConvertFileEncoding(aBuf);
	if (new_encoding == -1)
		_f_throw_value(ERR_INVALID_VALUE, aBuf);
	g->Encoding = new_encoding;
}



BIV_DECL_R(BIV_RegView)
{
	LPTSTR value;
	switch (g->RegView)
	{
	case KEY_WOW64_32KEY: value = _T("32"); break;
	case KEY_WOW64_64KEY: value = _T("64"); break;
	default: value = _T("Default"); break;
	}
	_f_return_p(value);
}

BIV_DECL_W(BIV_RegView_Set)
{
	LPTSTR aBuf = BivRValueToString();
	DWORD reg_view = Line::RegConvertView(aBuf);
	// Validate the parameter even if it's not going to be used.
	if (reg_view == -1)
		_f_throw_value(ERR_INVALID_VALUE, aBuf);
	// Since these flags cause the registry functions to fail on Win2k and have no effect on
	// any later 32-bit OS, ignore this command when the OS is 32-bit.  Leave A_RegView = "Default".
	if (IsOS64Bit())
		g->RegView = reg_view;
}



BIV_DECL_R(BIV_LastError)
{
	_f_return_i(g->LastError);
}

BIV_DECL_W(BIV_LastError_Set)
{
	SetLastError(g->LastError = (DWORD)BivRValueToInt64());
}



BIV_DECL_R(BIV_PtrSize)
{
	_f_return_i(sizeof(void *));
}



BIV_DECL_R(BIV_ScreenDPI)
{
	_f_return_i(g_ScreenDPI);
}



BIV_DECL_R(BIV_TrayMenu)
{
	// Currently ExpandExpression() does not automatically release objects returned
	// by BIVs since this is the only one, and we're retaining a reference to it.
	//g_script.mTrayMenu->AddRef();
	_f_return(g_script.mTrayMenu);
}



BIV_DECL_R(BIV_AllowMainWindow)
{
	_f_return_b(g_AllowMainWindow);
}

BIV_DECL_W(BIV_AllowMainWindow_Set)
{
	g_script.AllowMainWindow(BivRValueToBOOL());
}



BIV_DECL_R(BIV_IconHidden)
{
	_f_return_b(g_NoTrayIcon);
}

BIV_DECL_W(BIV_IconHidden_Set)
{
	g_script.ShowTrayIcon(!BivRValueToBOOL());
}

void Script::ShowTrayIcon(bool aShow)
{
	if (g_NoTrayIcon = !aShow) // Assign.
	{
		if (mNIC.hWnd) // Since it exists, destroy it.
		{
			Shell_NotifyIcon(NIM_DELETE, &mNIC); // Remove it.
			mNIC.hWnd = NULL;  // Set this as an indicator that tray icon is not installed.
			// but don't do DestroyMenu() on mTrayMenu->mMenu (if non-NULL) since it may have been
			// changed by the user to have the custom items on top of the standard items,
			// for example, and we don't want to lose that ordering in case the script turns
			// the icon back on at some future time during this session.  Also, the script
			// may provide some other means of displaying the tray menu.
		}
	}
	else
	{
		if (!mNIC.hWnd) // The icon doesn't exist, so create it.
		{
			CreateTrayIcon();
			UpdateTrayIcon(true);  // Force the icon into the correct pause/suspend state.
		}
	}
}

BIV_DECL_R(BIV_IconTip)
{
	// Return the custom tip if any, otherwise the default tip.
	_f_return_p(g_script.mTrayIconTip ? g_script.mTrayIconTip : g_script.mFileName);
}

BIV_DECL_W(BIV_IconTip_Set)
{
	g_script.SetTrayTip(BivRValueToString());
}

void Script::SetTrayTip(LPTSTR aText)
{
	// Allocate mTrayIconTip on first use even if aText is empty, so that
	// it will override the use of mFileName as the tray tip text.
	// This allows the script to completely disable the tray tooltip.
	if (!mTrayIconTip)
		mTrayIconTip = SimpleHeap::Alloc<TCHAR>(_countof(mNIC.szTip)); // SimpleHeap improves avg. case mem load.
	if (mTrayIconTip)
		tcslcpy(mTrayIconTip, aText, _countof(mNIC.szTip));
	if (mNIC.hWnd) // i.e. only update the tip if the tray icon exists (can't work otherwise).
	{
		UPDATE_TIP_FIELD
		Shell_NotifyIcon(NIM_MODIFY, &mNIC);  // Currently not checking its result (e.g. in case a shell other than Explorer is running).
	}
}

BIV_DECL_R(BIV_IconFile)
{
	_f_return_p(g_script.mCustomIconFile ? g_script.mCustomIconFile : _T(""));
}

BIV_DECL_R(BIV_IconNumber)
{
	_f_return_i(g_script.mCustomIconNumber);
}



BIV_DECL_R(BIV_PriorKey)
{
	int validEventCount = 0;
	// Start at the current event (offset 1)
	for (int iOffset = 1; iOffset <= g_MaxHistoryKeys; ++iOffset)
	{
		// Get index for circular buffer
		int i = (g_KeyHistoryNext + g_MaxHistoryKeys - iOffset) % g_MaxHistoryKeys;
		// Keep looking until we hit the second valid event
		if (g_KeyHistory[i].event_type != _T('i') // Not an ignored event.
			&& g_KeyHistory[i].event_type != _T('U') // Not a Unicode packet (SendInput/VK_PACKET).
			&& ++validEventCount > 1)
		{
			// Find the next most recent key-down
			if (!g_KeyHistory[i].key_up)
			{
				GetKeyName(g_KeyHistory[i].vk, g_KeyHistory[i].sc, _f_retval_buf, _f_retval_buf_size);
				_f_return_p(_f_retval_buf);
			}
		}
	}
	_f_return_empty;
}



LPTSTR GetExitReasonString(ExitReasons aExitReason)
{
	LPTSTR str;
	switch(aExitReason)
	{
	case EXIT_LOGOFF: str = _T("Logoff"); break;
	case EXIT_SHUTDOWN: str = _T("Shutdown"); break;
	// Since the below are all relatively rare, except WM_CLOSE perhaps, they are all included
	// as one word to cut down on the number of possible words (it's easier to write OnExit
	// functions to cover all possibilities if there are fewer of them).
	case EXIT_CRITICAL:
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



BIV_DECL_R(BIV_Space_Tab)
{
	_f_return_p(aVarName[5] ? _T(" ") : _T("\t"));
}

BIV_DECL_R(BIV_AhkVersion)
{
	_f_return_p(T_AHK_VERSION);
}

BIV_DECL_R(BIV_AhkPath)
{
#ifdef AUTOHOTKEYSC
	TCHAR buf[MAX_PATH];
	size_t length;
	if (length = GetAHKInstallDir(buf))
		// Name "AutoHotkey.exe" is assumed for code size reduction and because it's not stored in the registry:
		tcslcpy(buf + length, _T("\\AutoHotkey.exe"), MAX_PATH - length); // tcslcpy() in case registry has a path that is too close to MAX_PATH to fit AutoHotkey.exe
	//else leave it blank as documented.
	_f_return(buf, length);
#else
	_f_return_p(g_script.mOurEXE);
#endif
}



BIV_DECL_R(BIV_TickCount)
{
	_f_return(GetTickCount64());
}



BIV_DECL_R(BIV_Now)
{
	SYSTEMTIME st;
	if (aVarName[5]) // A_Now[U]TC
		GetSystemTime(&st);
	else
		GetLocalTime(&st);
	SystemTimeToYYYYMMDD(_f_retval_buf, st);
	_f_return_p(_f_retval_buf);
}

BIV_DECL_R(BIV_OSVersion)
{
	_f_return_p(const_cast<LPTSTR>(g_os.Version()));
}

BIV_DECL_R(BIV_Is64bitOS)
{
	_f_return_b(IsOS64Bit());
}

BIV_DECL_R(BIV_Language)
{
	LPTSTR aBuf = _f_retval_buf;
	_stprintf(aBuf, _T("%04hX"), GetSystemDefaultUILanguage());
	_f_return_p(aBuf);
}

BIV_DECL_R(BIV_UserName_ComputerName)
{
	TCHAR buf[MAX_PATH];  // Doesn't use MAX_COMPUTERNAME_LENGTH + 1 in case longer names are allowed in the future.
	DWORD buf_size = MAX_PATH; // Below: A_Computer[N]ame (N is the 11th char, index 10, which if present at all distinguishes between the two).
	if (   !(aVarName[10] ? GetComputerName(buf, &buf_size) : GetUserName(buf, &buf_size))   )
		*buf = '\0';
	_f_return(buf);
}

BIV_DECL_R(BIV_WorkingDir)
{
	// Use GetCurrentDirectory() vs. g_WorkingDir because any in-progress FileSelect()
	// dialog is able to keep functioning even when it's quasi-thread is suspended.  The
	// dialog can thus change the current directory as seen by the active quasi-thread even
	// though g_WorkingDir hasn't been updated.  It might also be possible for the working
	// directory to change in unusual circumstances such as a network drive being lost).
	// Update: FileSelectFile changing the current directory might be OS-specific;
	// I could not reproduce it on Windows 10.
	TCHAR buf[T_MAX_PATH]; // T_MAX_PATH vs. MAX_PATH only has an effect with Windows 10 long path awareness.
	DWORD length = GetCurrentDirectory(_countof(buf), buf);
	_f_return(buf, length);
}

BIV_DECL_W(BIV_WorkingDir_Set)
{
	if (!SetWorkingDir(BivRValueToString()))
		_f_throw_win32();
}

BIV_DECL_R(BIV_InitialWorkingDir)
{
	_f_return_p(g_WorkingDirOrig);
}

BIV_DECL_R(BIV_WinDir)
{
	TCHAR buf[MAX_PATH]; // MSDN (2018): The uSize parameter "should be set to MAX_PATH."
	VarSizeType length = GetSystemWindowsDirectory(buf, _countof(buf));
	_f_return(buf, length);
}

BIV_DECL_R(BIV_Temp)
{
	TCHAR buf[MAX_PATH+1]; // MSDN (2018): "The maximum possible return value is MAX_PATH+1 (261)."
	VarSizeType length = GetTempPath(_countof(buf), buf);
	if (length)
		if (buf[length - 1] == '\\') // Should always be true. MSDN: "The returned string ends with a backslash"
			buf[--length] = '\0'; // Omit the trailing backslash to improve friendliness/consistency.
	_f_return(buf, length);
}

BIV_DECL_R(BIV_ComSpec)
{
	TCHAR buf_temp[1]; // Just a fake buffer to pass to some API functions in lieu of a NULL, to avoid any chance of misbehavior. Keep the size at 1 so that API functions will always fail to copy to buf.
	auto size_required = GetEnvironmentVariable(_T("ComSpec"), buf_temp, 0);
	if (!TokenSetResult(aResultToken, nullptr, size_required)) // Avoids subtracting 1 to be conservative and to reduce code size (due to the need to otherwise check for zero and avoid subtracting 1 in that case).
		return;
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker_length = GetEnvVarReliable(_T("ComSpec"), aResultToken.marker);
}

BIV_DECL_R(BIV_SpecialFolderPath)
{
	TCHAR buf[MAX_PATH]; // One caller relies on this being explicitly limited to MAX_PATH.
	// SHGetFolderPath requires a buffer size of MAX_PATH, but the function was superseded
	// by SHGetKnownFolderPath in Windows Vista, and that function returns COM-allocated
	// memory of unknown length.  However, it seems the shell still does not support long
	// paths as of 2018.
	int aFolder;
	switch (ctoupper(aVarName[2]))
	{
	case 'P': // A_[P]rogram...
		if (ctoupper(aVarName[9]) == 'S') // A_Programs(Common)
			aFolder = aVarName[10] ? CSIDL_COMMON_PROGRAMS : CSIDL_PROGRAMS;
		else // A_Program[F]iles
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
	_f_return(buf);
}

BIV_DECL_R(BIV_MyDocuments) // Called by multiple callers.
{
	TCHAR buf[MAX_PATH]; // SHGetFolderPath requires a buffer size of MAX_PATH.  At least one caller relies on this.
	if (SHGetFolderPath(NULL, CSIDL_MYDOCUMENTS, NULL, SHGFP_TYPE_CURRENT, buf) != S_OK)
		*buf = '\0';
	// Since it is common (such as in networked environments) to have My Documents on the root of a drive
	// (such as a mapped drive letter), remove the backslash from something like M:\ because M: is more
	// appropriate for most uses:
	VarSizeType length = (VarSizeType)strip_trailing_backslash(buf);
	_f_return(buf, length);
}



BIF_DECL(BIF_CaretGetPos)
{
	Var *varX = ParamIndexToOutputVar(0);
	Var *varY = ParamIndexToOutputVar(1);
	
	// I believe only the foreground window can have a caret position due to relationship with focused control.
	HWND target_window = GetForegroundWindow(); // Variable must be named target_window for ATTACH_THREAD_INPUT.
	if (!target_window) // No window is in the foreground, report blank coordinate.
	{
		if (varX)
			varX->Assign();
		if (varY)
			varY->Assign();
		_f_return_i(FALSE);
	}

	DWORD now_tick = GetTickCount();

	GUITHREADINFO info;
	info.cbSize = sizeof(GUITHREADINFO);
	BOOL result = GetGUIThreadInfo(GetWindowThreadProcessId(target_window, NULL), &info) // Got info okay...
		&& info.hwndCaret; // ...and there is a caret.
	if (!result)
	{
		if (varX)
			varX->Assign();
		if (varY)
			varY->Assign();
		_f_return_i(FALSE);
	}
	POINT pt;
	pt.x = info.rcCaret.left;
	pt.y = info.rcCaret.top;
	// Unconditionally convert to screen coordinates, for simplicity.
	ClientToScreen(info.hwndCaret, &pt);
	// Now convert back to whatever is expected for the current mode.
	POINT origin = {0};
	CoordToScreen(origin, COORD_MODE_CARET);
	pt.x -= origin.x;
	pt.y -= origin.y;
	
	if (varX)
		varX->Assign(pt.x);
	if (varY)
		varY->Assign(pt.y);
	_f_return_i(TRUE);
}



BIV_DECL_R(BIV_Cursor)
{
	HCURSOR current_cursor;
	CURSORINFO ci;
	ci.cbSize = sizeof(CURSORINFO);
	current_cursor = GetCursorInfo(&ci) ? ci.hCursor : NULL;
	if (!current_cursor)
	{
		#define CURSOR_UNKNOWN _T("Unknown")
		_f_return_p(CURSOR_UNKNOWN);
	}

	// Static so that it's initialized on first use (should help performance after the first time):
	static HCURSOR sCursor[] = {LoadCursor(NULL, IDC_APPSTARTING), LoadCursor(NULL, IDC_ARROW)
		, LoadCursor(NULL, IDC_CROSS), LoadCursor(NULL, IDC_HELP), LoadCursor(NULL, IDC_IBEAM)
		, LoadCursor(NULL, IDC_ICON), LoadCursor(NULL, IDC_NO), LoadCursor(NULL, IDC_SIZE)
		, LoadCursor(NULL, IDC_SIZEALL), LoadCursor(NULL, IDC_SIZENESW), LoadCursor(NULL, IDC_SIZENS)
		, LoadCursor(NULL, IDC_SIZENWSE), LoadCursor(NULL, IDC_SIZEWE), LoadCursor(NULL, IDC_UPARROW)
		, LoadCursor(NULL, IDC_WAIT)}; // If IDC_HAND were added, it would break existing scripts that rely on Unknown being synonymous with Hand.  If ever added, IDC_HAND should return NULL on Win95/NT.
	static const size_t cursor_count = _countof(sCursor);
	// The order in the below array must correspond to the order in the above array:
	static LPTSTR sCursorName[cursor_count + 1] = {_T("AppStarting"), _T("Arrow")
		, _T("Cross"), _T("Help"), _T("IBeam")
		, _T("Icon"), _T("No"), _T("Size")
		, _T("SizeAll"), _T("SizeNESW"), _T("SizeNS")  // NESW = NorthEast+SouthWest
		, _T("SizeNWSE"), _T("SizeWE"), _T("UpArrow")
		, _T("Wait"), CURSOR_UNKNOWN}; // The default item must be last.

	int i;
	for (i = 0; i < cursor_count; ++i)
		if (sCursor[i] == current_cursor)
			break;

	_f_return_p(sCursorName[i]); // If i is out-of-bounds, "Unknown" will be used.
}

BIV_DECL_R(BIV_ScreenWidth_Height)
{
	_f_return_i(GetSystemMetrics(aVarName[13] ? SM_CYSCREEN : SM_CXSCREEN));
}

BIV_DECL_R(BIV_ScriptName)
{
	_f_return_p(g_script.mScriptName ? g_script.mScriptName : g_script.mFileName);
}

BIV_DECL_W(BIV_ScriptName_Set)
{
	// For simplicity, a new buffer is allocated each time.  It is not expected to be set frequently.
	LPTSTR script_name = _tcsdup(BivRValueToString());
	if (!script_name)
		_f_throw_oom;
	free(g_script.mScriptName);
	g_script.mScriptName = script_name;
}

BIV_DECL_R(BIV_ScriptDir)
{
	_f_return_p(g_script.mFileDir);
}

BIV_DECL_R(BIV_ScriptFullPath)
{
	_f_return_p(g_script.mFileSpec);
}

BIV_DECL_R(BIV_ScriptHwnd)
{
	_f_return((UINT_PTR)g_hWnd);
}


LineNumberType Script::CurrentLine()
{
	return mCurrLine ? mCurrLine->mLineNumber : mCombinedLineNumber;
}

BIV_DECL_R(BIV_LineNumber)
{
	_f_return_i(g_script.CurrentLine());
}


LPTSTR Script::CurrentFile()
{
	return Line::sSourceFile[mCurrLine ? mCurrLine->mFileIndex : mCurrFileIndex];
}

BIV_DECL_R(BIV_LineFile)
{
	_f_return_p(g_script.CurrentFile());
}


BIV_DECL_R(BIV_LoopFileName) // Called by multiple callers.
{
	LPTSTR filename = _T("");  // Set default.
	if (g->mLoopFile)
	{
		// cAlternateFileName can be blank if the file lacks a short name, but it can also be blank
		// if the file's proper name complies with all 8.3 requirements (not just length), so use
		// cFileName whenever cAlternateFileName is empty.  GetShortPathName() also behaves this way.
		if (   ctoupper(aVarName[10]) != 'S' // It's not A_LoopFileShortName or ...
			|| !*(filename = g->mLoopFile->cAlternateFileName)   ) // ... there's no alternate name (see above).
			filename = g->mLoopFile->cFileName;
	}
	_f_return_p(filename);
}

BIV_DECL_R(BIV_LoopFileExt)
{
	LPTSTR file_ext = _T("");  // Set default.
	if (g->mLoopFile)
	{
		if (file_ext = _tcsrchr(g->mLoopFile->cFileName, '.'))
			++file_ext;
		else // Reset to empty string vs. NULL.
			file_ext = _T("");
	}
	_f_return_p(file_ext);
}

BIV_DECL_R(BIV_LoopFileDir)
{
	if (!g->mLoopFile)
		_f_return_empty;
	LoopFilesStruct &lfs = *g->mLoopFile;
	LPTSTR dir_end = lfs.file_path + lfs.dir_length; // Start of the filename.
	size_t suffix_length = dir_end - lfs.file_path_suffix; // Directory names\ added since the loop started.
	size_t total_length = lfs.orig_dir_length + suffix_length;
	if (total_length)
		--total_length; // Omit the trailing slash.
	if (!TokenSetResult(aResultToken, nullptr, total_length))
		return;
	aResultToken.symbol = SYM_STRING;
	LPTSTR buf = aResultToken.marker;
	tmemcpy(buf, lfs.orig_dir, lfs.orig_dir_length);
	tmemcpy(buf + lfs.orig_dir_length, lfs.file_path_suffix, suffix_length);
	buf[total_length] = '\0'; // This replaces the final character copied above, if any.
}

void FixLoopFilePath(LPTSTR aBuf, LPTSTR aPattern)
// Fixes aBuf to account for "." and ".." as file patterns.  These match the directory itself
// or parent directory, so for example "x\y\.." returns a directory named "x" which appears to
// be inside "y".  Without the handling here, the invalid path "x\y\x" would be returned.
// A small amount of temporary buffer space might be wasted compared to handling this in the BIV,
// but this way minimizes code size (and these cases are rare anyway).
{
	int count = 0;
	if (*aPattern == '.')
	{
		if (!aPattern[1])
			count = 1; // aBuf "x\y\y" should be "x\y" for "x\y\.".
		else if (aPattern[1] == '.' && !aPattern[2])
			count = 2; // aBuf "x\y\x" should be "x" for "x\y\..".
	}
	for ( ; count > 0; --count)
	{
		LPTSTR end = _tcsrchr(aBuf, '\\');
		if (end)
			*end = '\0';
		else if (*aBuf && aBuf[1] == ':') // aBuf "C:x" should be "C:" for "C:" or "C:.".
			aBuf[2] = '\0';
	}
}

void ReturnLoopFilePath(ResultToken &aResultToken, LPTSTR aPattern, LPTSTR aPrefix, size_t aPrefixLen, LPTSTR aSuffix, size_t aSuffixLen)
{
	if (!TokenSetResult(aResultToken, nullptr, aPrefixLen + aSuffixLen))
		return;
	aResultToken.symbol = SYM_STRING;
	LPTSTR buf = aResultToken.marker;
	tmemcpy(buf, aPrefix, aPrefixLen);
	tmemcpy(buf + aPrefixLen, aSuffix, aSuffixLen + 1); // +1 for \0.
	FixLoopFilePath(buf, aPattern);
	aResultToken.marker_length = -1; // Let caller determine actual length.
}

BIV_DECL_R(BIV_LoopFilePath)
{
	if (!g->mLoopFile)
		_f_return_empty;
	LoopFilesStruct &lfs = *g->mLoopFile;
	// Combine the original directory specified by the script with the dynamic part of file_path
	// (i.e. the sub-directory and file names appended to it since the loop started):
	ReturnLoopFilePath(aResultToken, lfs.pattern
		, lfs.orig_dir, lfs.orig_dir_length
		, lfs.file_path_suffix, lfs.file_path_length - (lfs.file_path_suffix - lfs.file_path));
}

BIV_DECL_R(BIV_LoopFileFullPath)
{
	if (!g->mLoopFile)
		_f_return_empty;
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
	// Currently both are done by PerformLoopFilePattern(), for performance and in case the working
	// directory changes after the Loop begins.
	LoopFilesStruct &lfs = *g->mLoopFile;
	// Combine long_dir with the dynamic part of file_path:
	ReturnLoopFilePath(aResultToken, lfs.pattern
		, lfs.long_dir, lfs.long_dir_length
		, lfs.file_path_suffix, lfs.file_path_length - (lfs.file_path_suffix - lfs.file_path));
}

BIV_DECL_R(BIV_LoopFileShortPath)
// Unlike GetLoopFileShortName(), this function returns blank when there is no short path.
// This is done so that there's a way for the script to more easily tell the difference between
// an 8.3 name not being available (due to the being disabled in the registry) and the short
// name simply being the same as the long name.  For example, if short name creation is disabled
// in the registry, A_LoopFileShortName would contain the long name instead, as documented.
// But to detect if that short name is really a long name, A_LoopFileShortPath could be checked
// and if it's blank, there is no short name available.
{
	if (!g->mLoopFile)
		_f_return_empty;
	LoopFilesStruct &lfs = *g->mLoopFile;
	// MSDN says cAlternateFileName is empty if the file does not have a long name.
	// Testing and research shows that GetShortPathName() uses the long name for a directory
	// or file if no short name exists, so there's no check for the filename's length here.
	LPTSTR name = *lfs.cAlternateFileName ? lfs.cAlternateFileName : lfs.cFileName;
	ReturnLoopFilePath(aResultToken, lfs.pattern
		, lfs.short_path, lfs.short_path_length
		, name, _tcslen(name));
}

BIV_DECL_R(BIV_LoopFileTime)
{
	LPTSTR target_buf = _f_retval_buf;
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
	_f_return_p(target_buf);
}

BIV_DECL_R(BIV_LoopFileAttrib)
{
	LPTSTR target_buf = _f_retval_buf;
	*target_buf = '\0'; // Set default.
	if (g->mLoopFile)
		FileAttribToStr(target_buf, g->mLoopFile->dwFileAttributes);
	_f_return_p(target_buf);
}

BIV_DECL_R(BIV_LoopFileSize)
{
	if (g->mLoopFile)
	{
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
		_f_return_i(divider ? ((unsigned __int64)ul.QuadPart / divider) : ul.QuadPart);
	}
	_f_return_empty;
}

BIV_DECL_R(BIV_LoopRegType)
{
	_f_return_p(g->mLoopRegItem ? Line::RegConvertValueType(g->mLoopRegItem->type) : _T(""));
}

BIV_DECL_R(BIV_LoopRegKey)
{
	LPTSTR rootkey = _T("");
	LPTSTR subkey = _T("");
	if (g->mLoopRegItem)
	{
		// Use root_key_type, not root_key (which might be a remote vs. local HKEY):
		rootkey = Line::RegConvertRootKeyType(g->mLoopRegItem->root_key_type);
		subkey = g->mLoopRegItem->subkey;
	}
	if (!TokenSetResult(aResultToken, nullptr, _tcslen(rootkey) + (*subkey != 0) + _tcslen(subkey)))
		return;
	_stprintf(aResultToken.marker, _T("%s%s%s"), rootkey, *subkey ? _T("\\") : _T(""), subkey);
	aResultToken.symbol = SYM_STRING;
}

BIV_DECL_R(BIV_LoopRegName)
{
	// This can be either the name of a subkey or the name of a value.
	_f_return_p(g->mLoopRegItem ? g->mLoopRegItem->name : _T(""));
}

BIV_DECL_R(BIV_LoopRegTimeModified)
{
	LPTSTR target_buf = _f_retval_buf;
	*target_buf = '\0'; // Set default.
	// Only subkeys (not values) have a time.
	if (g->mLoopRegItem && g->mLoopRegItem->type == REG_SUBKEY)
		FileTimeToYYYYMMDD(target_buf, g->mLoopRegItem->ftLastWriteTime, true);
	_f_return_p(target_buf);
}

BIV_DECL_R(BIV_LoopReadLine)
{
	_f_return_p(g->mLoopReadFile ? g->mLoopReadFile->mCurrentLine : _T(""));
}

BIV_DECL_R(BIV_LoopField)
{
	_f_return_p(g->mLoopField ? g->mLoopField : _T(""));
}

BIV_DECL_R(BIV_LoopIndex)
{
	_f_return_i(g->mLoopIteration);
}

BIV_DECL_W(BIV_LoopIndex_Set)
{
	g->mLoopIteration = BivRValueToInt64();
}



BIV_DECL_R(BIV_ThisFunc)
{
	LPCTSTR name;
	if (g->CurrentFunc)
		name = g->CurrentFunc->mName;
	else
		name = _T("");
	_f_return_p(const_cast<LPTSTR>(name));
}

BIV_DECL_R(BIV_ThisHotkey)
{
	_f_return_p(g_script.mThisHotkeyName);
}

BIV_DECL_R(BIV_PriorHotkey)
{
	_f_return_p(g_script.mPriorHotkeyName);
}

BIV_DECL_R(BIV_TimeSinceThisHotkey)
{
	// It must be the type of hotkey that has a label because we want the TimeSinceThisHotkey
	// value to be "in sync" with the value of ThisHotkey itself (i.e. use the same method
	// to determine which hotkey is the "this" hotkey):
	if (*g_script.mThisHotkeyName)
		// Even if GetTickCount()'s TickCount has wrapped around to zero and the timestamp hasn't,
		// DWORD subtraction still gives the right answer as long as the number of days between
		// isn't greater than about 49.  See MyGetTickCount() for explanation of %d vs. %u.
		// Update: Using 64-bit ints now, so above is obsolete:
		//sntprintf(str, sizeof(str), "%d", (DWORD)(GetTickCount() - g_script.mThisHotkeyStartTime));
		_f_return_i((__int64)(GetTickCount() - g_script.mThisHotkeyStartTime));
	else
		_f_return_empty; // Cause any attempt at math to throw.
}

BIV_DECL_R(BIV_TimeSincePriorHotkey)
{
	if (*g_script.mPriorHotkeyName)
		// See MyGetTickCount() for explanation for explanation:
		//sntprintf(str, sizeof(str), "%d", (DWORD)(GetTickCount() - g_script.mPriorHotkeyStartTime));
		_f_return_i((__int64)(GetTickCount() - g_script.mPriorHotkeyStartTime));
	else
		_f_return_empty; // Cause any attempt at math to throw.
}

BIV_DECL_R(BIV_EndChar)
{
	_f_retval_buf[0] = g_script.mEndChar;
	_f_retval_buf[1] = '\0';
	_f_return_p(_f_retval_buf);
}



BIV_DECL_R(BIV_EventInfo)
// We're returning the length of the var's contents, not the size.
{
	_f_return_i(g->EventInfo);
}

BIV_DECL_W(BIV_EventInfo_Set)
{
	g->EventInfo = (EventInfoType)BivRValueToInt64();
}



BIV_DECL_R(BIV_TimeIdle)
// This is here rather than in script.h with the others because it depends on
// hotkey.h and globaldata.h, which can't be easily included in script.h due to
// mutual dependency issues.
{
	DWORD time_last_input = 0;
	switch (toupper(aVarName[10]))
	{
	case 'M': time_last_input = g_MouseHook ? g_TimeLastInputMouse : 0; break;
	case 'K': time_last_input = g_KeybdHook ? g_TimeLastInputKeyboard : 0; break;
	case 'P': time_last_input = (g_KeybdHook || g_MouseHook) ? g_TimeLastInputPhysical : 0; break;
	}
	// If the relevant hook is not active, default this to the same as the regular idle time:
	if (!time_last_input)
	{
		LASTINPUTINFO lii;
		lii.cbSize = sizeof(lii);
		if (GetLastInputInfo(&lii))
			time_last_input = lii.dwTime;
		else // This is rare; the possibility isn't even documented as of 2020-06-11.
			_f_return_empty; // Cause any attempt at math to throw.
	}
	_f_return_i(GetTickCount() - time_last_input);
}



BIF_DECL(BIF_SetBIV)
{
	static VirtualVar::Setter sBiv[] = { &BIV_DetectHiddenText_Set, &BIV_DetectHiddenWindows_Set, &BIV_FileEncoding_Set, &BIV_RegView_Set, &BIV_StoreCapsLockMode_Set, &BIV_TitleMatchMode_Set };
	auto biv = sBiv[_f_callee_id];
	_f_set_retval_p(_T(""), 0);
	biv(aResultToken, nullptr, *aParam[0]);
}



BIF_DECL(BIF_Persistent)
{
	// Need to set a return value explicitly, since the default is effectively StrPtr(""), not "".
	// Returning the old value might have some use, but if the caller doesn't want its value to change,
	// something awkward like "Persistent(isPersistent := Persistent())" is needed.  Rather than just
	// returning the current status, Persistent() makes the script persistent because that's likely to
	// be its most common use by far, and it's what users familiar with the old #Persistent may expect.
	_f_set_retval_i(g_persistent);
	g_persistent = ParamIndexToOptionalBOOL(0, true);
}



BIF_DECL(BIF_InstallHook)
{
	bool installing = ParamIndexToOptionalBOOL(0, true);
	bool use_force = ParamIndexToOptionalBOOL(1, false);
	auto which_hook = (HookType)_f_callee_id;
	// When the second parameter is true, unconditionally remove the hook.  If the first parameter is
	// also true, the hook will be reinstalled fresh.  Otherwise the hook will be left uninstalled,
	// until something happens to reinstall it, such as Hotkey::ManifestAllHotkeysHotstringsHooks().
	if (use_force)
		AddRemoveHooks(GetActiveHooks() & ~which_hook);
	Hotkey::RequireHook(which_hook, installing);
	if (!use_force || installing)
		Hotkey::ManifestAllHotkeysHotstringsHooks();
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
	bool is_hresult; // Only used for the return value.
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
		_itot(esp_delta, buf, 10);
		if (esp_delta > 0)
			g_script.ThrowRuntimeException(_T("Parameter list too large, or call requires CDecl."), buf);
		else
			g_script.ThrowRuntimeException(_T("Parameter list too small."), buf);
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
		g_script.ThrowRuntimeException(ERR_EXCEPTION, buf);
	}

	return Res;
}



void ConvertDllArgType(LPTSTR aBuf, DYNAPARM &aDynaParam)
// Helper function for DllCall().  Updates aDynaParam's type and other attributes.
{
	LPTSTR type_string = aBuf;
	TCHAR buf[32];
	
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
		aDynaParam.type = DLL_ARG_INVALID; 
		return; 
	}

	tcslcpy(buf, type_string, _countof(buf)); // Make a modifiable copy for easier parsing below.

	// v1.0.30.02: The addition of 'P'.
	// However, the current detection below relies upon the fact that none of the types currently
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
		aDynaParam.type = DLL_ARG_INVALID; 
		return;
	}
	return; // Since above didn't "return", the type is explicitly valid
	
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
				if (hmodule_to_free) // L31: BIF_DllCall wants us to set throw.  ExpressionToPostfix passes NULL.
					g_script.ThrowRuntimeException(_T("Failed to load DLL."), dll_name);
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

	if (!function && hmodule_to_free) // Caller wants us to throw.
	{
		// This must be done here since only we know for certain that the dll
		// was loaded okay (if GetModuleHandle succeeded, nothing is passed
		// back to the caller).
		g_script.ThrowRuntimeException(ERR_NONEXISTENT_FUNCTION, _tfunction_name);
	}

	return function;
}



BIF_DECL(BIF_DllCall)
// Stores a number or a SYM_STRING result in aResultToken.
// Caller has set up aParam to be viewable as a left-to-right array of params rather than a stack.
// It has also ensured that the array has exactly aParamCount items in it.
// Author: Marcus Sonntag (Ultra)
{
	HMODULE hmodule_to_free = NULL; // Set default in case of early goto; mostly for maintainability.
	LPTSTR function_name = NULL;
	void *function = NULL; // Will hold the address of the function to be called.
	int vf_index = -1; // Set default: not ComCall.

	if (_f_callee_id == FID_ComCall)
	{
		function = NULL;
		if (!ParamIndexIsNumeric(0))
			_f_throw_param(0, _T("Integer"));
		vf_index = (int)ParamIndexToInt64(0);
		if (vf_index < 0) // But positive values aren't checked since there's no known upper bound.
			_f_throw_param(0);
		// Cheat a bit to make the second arg both the source of the virtual function
		// and the first parameter value (always an interface pointer):
		static ExprTokenType t_this_arg_type = _T("Ptr");
		aParam[0] = &t_this_arg_type;
	}
	else
	{
		// Check that the mandatory first parameter (DLL+Function) is valid.
		// (load-time validation has ensured at least one parameter is present).
		switch (TypeOfToken(*aParam[0]))
		{
		case SYM_INTEGER: // Might be the most common case, due to FinalizeExpression resolving function names at load time.
			// v1.0.46.08: Allow script to specify the address of a function, which might be useful for
			// calling functions that the script discovers through unusual means such as C++ member functions.
			function = (void *)ParamIndexToInt64(0);
			// A check like the following is not present due to rarity of need and because if the address
			// is zero or negative, the same result will occur as for any other invalid address:
			// an exception code of 0xc0000005.
			//if ((UINT64)temp64 < 0x10000 || (UINT64)temp64 > UINTPTR_MAX)
			//	_f_throw_param(0); // Stage 1 error: Invalid first param.
			//// Otherwise, assume it's a valid address:
			//	function = (void *)temp64;
			break;
		case SYM_STRING: // For performance, don't even consider the possibility that a string like "33" is a function-address.
			//function = NULL; // Already set: indicates that no function has been specified yet.
			break;
		case SYM_OBJECT:
			// Permit an object with Ptr property.  This enables DllCall or DllCall.Bind() to be used directly
			// as a method of an object, such as one used for wrapping a dll function.  It could also have other
			// uses, such as resolving and memoizing function addresses on first use.
			__int64 n;
			if (!GetObjectIntProperty(ParamIndexToObject(0), _T("Ptr"), n, aResultToken))
				return;
			function = (void *)n;
			break;
		default: // SYM_FLOAT, SYM_MISSING or (should be impossible) something else.
			_f_throw(ERR_PARAM1_INVALID, ErrorPrototype::Type);
		}
		if (!function)
			function_name = TokenToString(*aParam[0]);
		++aParam; // Normalize aParam to simplify ComCall vs. DllCall.
		--aParamCount;
	}

	// Determine the type of return value.
	DYNAPARM return_attrib = {0}; // Init all to default in case ConvertDllArgType() isn't called below. This struct holds the type and other attributes of the function's return value.
#ifdef WIN32_PLATFORM
	int dll_call_mode = DC_CALL_STD; // Set default.  Can be overridden to DC_CALL_CDECL and flags can be OR'd into it.
#endif
	if ( !(aParamCount % 2) ) // An even number of parameters indicates the return type has been omitted. aParamCount excludes DllCall's first parameter at this point.
	{
		return_attrib.type = DLL_ARG_INT;
		if (vf_index >= 0) // Default to HRESULT for ComCall.
			return_attrib.is_hresult = true;
		// Otherwise, assume normal INT (also covers BOOL).
	}
	else
	{
		// Check validity of this arg's return type:
		ExprTokenType &token = *aParam[aParamCount - 1];
		LPTSTR return_type_string = TokenToString(token); // If non-numeric it will return "", which is detected as invalid below.

		// 64-bit note: The calling convention detection code is preserved here for script compatibility.

		if (!_tcsnicmp(return_type_string, _T("CDecl"), 5)) // Alternate calling convention.
		{
#ifdef WIN32_PLATFORM
			dll_call_mode = DC_CALL_CDECL;
#endif
			return_type_string = omit_leading_whitespace(return_type_string + 5);
			if (!*return_type_string)
			{	// Take a shortcut since we know this empty string will be used as "Int":
				return_attrib.type = DLL_ARG_INT;
				goto has_valid_return_type;
			}
		}
		if (!_tcsicmp(return_type_string, _T("HRESULT")))
		{
			return_attrib.type = DLL_ARG_INT;
			return_attrib.is_hresult = true;
			//return_attrib.is_unsigned = true; // Not relevant since an exception is thrown for any negative value.
		}
		else
			ConvertDllArgType(return_type_string, return_attrib);
		if (return_attrib.type == DLL_ARG_INVALID)
			_f_throw_value(ERR_INVALID_RETURN_TYPE);
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
	int arg_count = aParamCount/2;
	DYNAPARM *dyna_param = arg_count ? (DYNAPARM *)_alloca(arg_count * sizeof(DYNAPARM)) : NULL;
	// Above: _alloca() has been checked for code-bloat and it doesn't appear to be an issue.
	// Above: Fix for v1.0.36.07: According to MSDN, on failure, this implementation of _alloca() generates a
	// stack overflow exception rather than returning a NULL value.  Therefore, NULL is no longer checked,
	// nor is an exception block used since stack overflow in this case should be exceptionally rare (if it
	// does happen, it would probably mean the script or the program has a design flaw somewhere, such as
	// infinite recursion).

	LPTSTR arg_type_string;
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
	for (arg_count = 0, i = 0; i < aParamCount; ++arg_count, i += 2)  // Same loop as used later below, so maintain them together.
	{
		// Store each arg into a dyna_param struct, using its arg type to determine how.
		DYNAPARM &this_dyna_param = dyna_param[arg_count];

		arg_type_string = TokenToString(*aParam[i]); // aBuf not needed since numbers and "" are equally invalid.
		ConvertDllArgType(arg_type_string, this_dyna_param);
		if (this_dyna_param.type == DLL_ARG_INVALID)
			_f_throw_value(ERR_INVALID_ARG_TYPE);

		IObject *this_param_obj = TokenToObject(*aParam[i + 1]);
		if (this_param_obj)
		{
			if ((this_dyna_param.passed_by_address || this_dyna_param.type == DLL_ARG_STR)
				&& dynamic_cast<VarRef*>(this_param_obj))
			{
				aParam[i + 1] = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
				aParam[i + 1]->SetVarRef(static_cast<VarRef*>(this_param_obj));
				this_param_obj = nullptr;
			}
			else if (ctoupper(*arg_type_string) == 'P')
			{
				// Support Buffer.Ptr, but only for "Ptr" type.  All other types are reserved for possible
				// future use, which might be general like obj.ToValue(), or might be specific to DllCall
				// or the particular type of this arg (Int, Float, etc.).
				GetBufferObjectPtr(aResultToken, this_param_obj, this_dyna_param.value_uintptr);
				if (aResultToken.Exited())
					return;
				continue;
			}
		}
		ExprTokenType &this_param = *aParam[i + 1];
		if (this_param.symbol == SYM_MISSING)
			_f_throw(ERR_PARAM_REQUIRED);

		switch (this_dyna_param.type)
		{
		case DLL_ARG_STR:
			if (IS_NUMERIC(this_param.symbol) || this_param_obj)
			{
				// For now, string args must be real strings rather than floats or ints.  An alternative
				// to this would be to convert it to number using persistent memory from the caller (which
				// is necessary because our own stack memory should not be passed to any function since
				// that might cause it to return a pointer to stack memory, or update an output-parameter
				// to be stack memory, which would be invalid memory upon return to the caller).
				// The complexity of this doesn't seem worth the rarity of the need, so this will be
				// documented in the help file.
				_f_throw_type(_T("String"), this_param);
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
			// common for a script to (unintentionally) pass an empty/uninitialized variable to
			// some function that writes a string to it.
			//if (this_dyna_param.str == Var::sEmptyString) // To improve performance, compare directly to Var::sEmptyString rather than calling Capacity().
			//	this_dyna_param.str = _T(""); // Make it read-only to force an exception.  See comments above.
			break;
		case DLL_ARG_xSTR:
			// See the section above for comments.
			if (IS_NUMERIC(this_param.symbol) || this_param_obj)
				_f_throw_type(_T("String"), this_param);
			// String needing translation: ASTR on Unicode build, WSTR on ANSI build.
			pStr[arg_count] = new UorA(CStringCharFromWChar,CStringWCharFromChar)(TokenToString(this_param));
			this_dyna_param.ptr = pStr[arg_count]->GetBuffer();
			break;

		case DLL_ARG_DOUBLE:
		case DLL_ARG_FLOAT:
			// This currently doesn't validate that this_dyna_param.is_unsigned==false, since it seems
			// too rare and mostly harmless to worry about something like "Ufloat" having been specified.
			if (!TokenIsNumeric(this_param))
				_f_throw_type(_T("Number"), this_param);
			this_dyna_param.value_double = TokenToDouble(this_param);
			if (this_dyna_param.type == DLL_ARG_FLOAT)
				this_dyna_param.value_float = (float)this_dyna_param.value_double;
			break;

		default: // Namely:
		//case DLL_ARG_INT:
		//case DLL_ARG_SHORT:
		//case DLL_ARG_CHAR:
		//case DLL_ARG_INT64:
			if (!TokenIsNumeric(this_param))
				_f_throw_type(_T("Number"), this_param);
			// Note that since v2.0-a083-97803aeb, TokenToInt64 supports conversion of large unsigned 64-bit
			// numbers from strings (producing a negative value, but with the right bit representation).
			// This allows large unsigned literals and numeric strings to be passed to DllCall (regardless
			// of whether Int64 or UInt64 is used), but the script itself will interpret the value as signed
			// if greater than _I64_MAX.  Any UInt64 values returned by DllCall can be safely passed back
			// without loss, and can be operated upon by the bitwise operators, although arithmetic and
			// string conversion will treat the value as Int64.
			this_dyna_param.value_int64 = TokenToInt64(this_param);
		} // switch (this_dyna_param.type)
	} // for() each arg.
    
	if (vf_index >= 0) // ComCall
	{
		if ((UINT_PTR)dyna_param[0].ptr < 65536) // Basic sanity check to catch null pointers and small numbers.  On Win32, the first 64KB of address space is always invalid.
			_f_throw_param(1);
		LPVOID *vftbl = *(LPVOID **)dyna_param[0].ptr;
		function = vftbl[vf_index];
	}
	else if (!function) // The function's address hasn't yet been determined.
	{
		function = GetDllProcAddress(function_name, &hmodule_to_free);
		if (!function)
		{
			// GetDllProcAddress has thrown the appropriate exception.
			aResultToken.SetExitResult(FAIL);
			goto end;
		}
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
		g_script.CriticalError(_T("An invalid write to an empty variable was detected."));
		// CriticalError always terminates the process.
	}

	if (g->ThrownToken || return_attrib.is_hresult && FAILED((HRESULT)return_value.Int))
	{
		if (!g->ThrownToken)
			// "Error values (as defined by the FAILED macro) are never returned"; so FAIL, not FAIL_OR_OK.
			g_script.Win32Error((DWORD)return_value.Int, FAIL);
		// If a script exception was thrown by DynaCall(), it was either because the called function threw
		// a SEH exception or because the stdcall parameter list was the wrong size.  In any of these cases,
		// set FAIL result to ensure control is transferred as expected (exiting the thread or TRY block).
		aResultToken.SetExitResult(FAIL);
		// But continue on to write out any output parameters because the called function might have
		// had a chance to update them before aborting.  They might be of some use in debugging the
		// issue, though the script would have to catch the exception to be able to inspect them.
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
			ASSERT(aResultToken.symbol == SYM_INTEGER);
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
			aResultToken.symbol = SYM_STRING;
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
					else
						aResultToken.marker = _T("");
				}
				else
					aResultToken.marker = _T("");
				aResultToken.symbol = SYM_STRING;
			}
			break;
		case DLL_ARG_SHORT:
			ASSERT(aResultToken.symbol == SYM_INTEGER);
			if (return_attrib.is_unsigned)
				aResultToken.value_int64 = return_value.Int & 0x0000FFFF; // This also forces the value into the unsigned domain of a signed int.
			else // Signed.
				aResultToken.value_int64 = (SHORT)(WORD)return_value.Int; // These casts properly preserve negatives.
			break;
		case DLL_ARG_CHAR:
			ASSERT(aResultToken.symbol == SYM_INTEGER);
			if (return_attrib.is_unsigned)
				aResultToken.value_int64 = return_value.Int & 0x000000FF; // This also forces the value into the unsigned domain of a signed int.
			else // Signed.
				aResultToken.value_int64 = (char)(BYTE)return_value.Int; // These casts properly preserve negatives.
			break;
		case DLL_ARG_INT64:
			// Even for unsigned 64-bit values, it seems best both for simplicity and consistency to write
			// them back out to the script as signed values because script internals are not currently
			// equipped to handle unsigned 64-bit values.  This has been documented.
			ASSERT(aResultToken.symbol == SYM_INTEGER);
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
	for (arg_count = 0, i = 0; i < aParamCount; ++arg_count, i += 2) // Same loop as used above, so maintain them together.
	{
		ExprTokenType &this_param = *aParam[i + 1];  // Resolved for performance and convenience.
		DYNAPARM &this_dyna_param = dyna_param[arg_count];

		if (IObject * obj = TokenToObject(this_param)) // Implies the type is "Ptr" or "Ptr*".
		{
			if (this_dyna_param.passed_by_address)
				SetObjectIntProperty(obj, _T("Ptr"), this_dyna_param.value_int64, aResultToken);
			continue;
		}

		if (this_param.symbol != SYM_VAR)
			continue;
		Var &output_var = *this_param.var;

		if (!this_dyna_param.passed_by_address)
		{
			if (this_dyna_param.type == DLL_ARG_STR) // Native string type for current build config.
			{
				// Update the variable's length and check for null termination.  This could be skipped
				// when a naked variable (not VarRef) is passed since that's supposed to be input-only,
				// but seems better to do this unconditionally since the function can in fact modify
				// the variable's contents, and detecting buffer overrun errors seems more important
				// than any performance gain from skipping this.
				output_var.SetLengthFromContents();
				output_var.Close(); // Clear the attributes of the variable to reflect the fact that the contents may have changed.
			}
			// Nothing is done for xSTR since 1) we didn't pass the variable's contents to the function
			// so its length doesn't need updating, and 2) the buffer that was passed was only as large
			// as the input string, so has very little practical use for output.
			// No other types can be output parameters when !passed_by_address.
			continue;
		}
		if (this_param.var_usage == Script::VARREF_READ)
			continue; // Output parameters are copied back only if provided with a VarRef (&variable).

		switch (this_dyna_param.type)
		{
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
		case DLL_ARG_STR: // Str*
			// The use of LPWSTR* vs. LPWSTR typically means the function will pass back the
			// address of a string, not modify the string itself.  This is also consistent with
			// passed_by_address for all other types.  However, it must be used carefully since
			// there's no way for Str* to know how or whether the function requires the string
			// to be freed (e.g. by calling CoTaskMemFree()).
			if (this_dyna_param.ptr != output_var.Contents(FALSE)
				&& !output_var.AssignString((LPTSTR)this_dyna_param.ptr))
				aResultToken.SetExitResult(FAIL);
			break;
		case DLL_ARG_xSTR: // AStr* on Unicode builds and WStr* on ANSI builds.
			if (this_dyna_param.ptr != output_var.Contents(FALSE)
				&& !output_var.AssignStringFromCodePage(UorA(LPSTR,LPWSTR)this_dyna_param.ptr))
				aResultToken.SetExitResult(FAIL);
		}
	}

end:
	for (arg_count = (aParamCount / 2) - 1; arg_count >= 0; --arg_count)
		if (pStr[arg_count])
			delete pStr[arg_count];
	if (hmodule_to_free)
		FreeLibrary(hmodule_to_free);
}

#endif


void ObjectToString(ResultToken &aResultToken, ExprTokenType &aThisToken, IObject *aObject)
{
	// Something like this should be done for every TokenToString() call or
	// equivalent, but major changes are needed before that will be feasible.
	// For now, String(anytype) provides a limited workaround.
	switch (aObject->Invoke(aResultToken, IT_CALL, _T("ToString"), aThisToken, nullptr, 0))
	{
	case INVOKE_NOT_HANDLED:
		aResultToken.UnknownMemberError(aThisToken, IT_CALL, _T("ToString"));
		break;
	case FAIL:
		aResultToken.SetExitResult(FAIL);
		break;
	}
}

BIF_DECL(BIF_StrCompare)
{
	// In script:
	// Result := StrCompare(str1, str2 [, CaseSensitive := false])
	// str1, str2, the strings to compare, can be be pure numbers, such numbers are converted to strings before the comparison.
	// CaseSensitive, the case sensitive setting to use.
	// Result, the result of the comparison:
	//	< 0,	if str1 is less than str2.
	//	0,		if str1 is identical to str2.
	//	> 0,	if str1 is greater than str2.
	// Param 1 and 2, str1 and str2
	TCHAR str1_buf[MAX_NUMBER_SIZE];	// numeric input is converted to string.
	TCHAR str2_buf[MAX_NUMBER_SIZE];
	LPTSTR str1 = ParamIndexToString(0, str1_buf);
	LPTSTR str2 = ParamIndexToString(1, str2_buf);
	// Could return 0 here if str1 == str2, but it is probably rare to call StrCompare(str_var, str_var)
	// so for most cases that would just be an unnecessary cost, albeit low.
	// Param 3 - CaseSensitive
	StringCaseSenseType string_case_sense = ParamIndexToCaseSense(2);
	int result;
	switch (string_case_sense)		// Compare the strings according to the string_case_sense setting.
	{
	case SCS_SENSITIVE:				result = _tcscmp(str1, str2); break;
	case SCS_INSENSITIVE:			result = _tcsicmp(str1, str2); break;
	case SCS_INSENSITIVE_LOCALE:	result = lstrcmpi(str1, str2); break;
	case SCS_INSENSITIVE_LOGICAL:	result = StrCmpLogicalW(str1, str2); break;
	case SCS_INVALID:
		_f_throw_param(2);
	}
	_f_return_i(result);			// Return 
}

BIF_DECL(BIF_String)
{
	++aParam; // Skip `this`
	aResultToken.symbol = SYM_STRING;
	switch (aParam[0]->symbol)
	{
	case SYM_STRING:
		aResultToken.marker = aParam[0]->marker;
		aResultToken.marker_length = aParam[0]->marker_length;
		break;
	case SYM_VAR:
		if (aParam[0]->var->HasObject())
		{
			ObjectToString(aResultToken, *aParam[0], aParam[0]->var->Object());
			break;
		}
		aResultToken.marker = aParam[0]->var->Contents();
		aResultToken.marker_length = aParam[0]->var->CharLength();
		break;
	case SYM_INTEGER:
		aResultToken.marker = ITOA64(aParam[0]->value_int64, _f_retval_buf);
		break;
	case SYM_FLOAT:
		aResultToken.marker = _f_retval_buf;
		aResultToken.marker_length = FTOA(aParam[0]->value_double, aResultToken.marker, _f_retval_buf_size);
		break;
	case SYM_OBJECT:
		ObjectToString(aResultToken, *aParam[0], aParam[0]->object);
		break;
	// Impossible due to parameter count validation:
	//case SYM_MISSING:
	//	_f_throw_value(ERR_PARAM1_REQUIRED);
#ifdef _DEBUG
	default:
		MsgBox(_T("DEBUG: type not handled"));
		_f_return_FAIL;
#endif
	}
}



BIF_DECL(BIF_StrLen)
{
	// Caller has ensured that there's exactly one actual parameter.
	size_t length;
	ParamIndexToString(0, _f_number_buf, &length); // Allow StrLen(numeric_expr) for flexibility.
	_f_return_i(length);
}



BIF_DECL(BIF_SubStr) // Added in v1.0.46.
{
	// Set default return value in case of early return.
	_f_set_retval_p(_T(""), 0);

	// Get the first arg, which is the string used as the source of the extraction. Call it "haystack" for clarity.
	TCHAR haystack_buf[MAX_NUMBER_SIZE]; // A separate buf because _f_number_buf is sometimes used to store the result.
	INT_PTR haystack_length;
	LPTSTR haystack = ParamIndexToString(0, haystack_buf, (size_t *)&haystack_length);

	// Load-time validation has ensured that at least the first two parameters are present:
	INT_PTR starting_offset = (INT_PTR)ParamIndexToInt64(1); // The one-based starting position in haystack (if any).
	if (starting_offset > haystack_length || starting_offset == 0)
		_f_return_retval; // Yield the empty string (a default set higher above).
	if (starting_offset < 0) // Same convention as RegExMatch/Replace(): Treat negative StartingPos as a position relative to the end of the string.
	{
		starting_offset += haystack_length;
		if (starting_offset < 0)
			starting_offset = 0;
	}
	else
		--starting_offset; // Convert to zero-based.

	INT_PTR remaining_length_available = haystack_length - starting_offset;
	INT_PTR extract_length;
	if (ParamIndexIsOmitted(2)) // No length specified, so extract all the remaining length.
		extract_length = remaining_length_available;
	else
	{
		if (   !(extract_length = (INT_PTR)ParamIndexToInt64(2))   )  // It has asked to extract zero characters.
			_f_return_retval; // Yield the empty string (a default set higher above).
		if (extract_length < 0)
		{
			extract_length += remaining_length_available; // Result is the number of characters to be extracted (i.e. after omitting the number of chars specified in extract_length).
			if (extract_length < 1) // It has asked to omit all characters.
				_f_return_retval; // Yield the empty string (a default set higher above).
		}
		else // extract_length > 0
			if (extract_length > remaining_length_available)
				extract_length = remaining_length_available;
	}

	// Above has set extract_length to the exact number of characters that will actually be extracted.
	LPTSTR result = haystack + starting_offset; // This is the result except for the possible need to truncate it below.

	// No need for any copying or termination, just send back part of haystack.
	// Caller and Var::Assign() know that overlap is possible, so this seems safe.
	_f_return_p(result, extract_length);
}



BIF_DECL(BIF_InStr)
{
	// Caller has already ensured that at least two actual parameters are present.
	TCHAR needle_buf[MAX_NUMBER_SIZE];
	INT_PTR haystack_length;
	LPTSTR haystack = ParamIndexToString(0, _f_number_buf, (size_t *)&haystack_length);
	size_t needle_length;
	LPTSTR needle = ParamIndexToString(1, needle_buf, &needle_length);
	if (!needle_length) // Although arguably legitimate, this is more likely an error, so throw.
		_f_throw_param(1);

	StringCaseSenseType string_case_sense = ParamIndexToCaseSense(2);
	if (string_case_sense == SCS_INVALID
		|| string_case_sense == SCS_INSENSITIVE_LOGICAL) // Not supported, seems more useful to throw rather than using SCS_INSENSITIVE.
		_f_throw_param(2);
	// BIF_StrReplace sets string_case_sense similarly, maintain together.

	INT_PTR offset = ParamIndexToOptionalIntPtr(3, 1);
	if (!offset)
		_f_throw_param(3);
	
	int occurrence_number = ParamIndexToOptionalInt(4, 1);
	if (!occurrence_number)
		_f_throw_param(4);

	if (offset < 0)
	{
		if (ParamIndexIsOmitted(4))
			occurrence_number = -1; // Default to RTL.
		offset += haystack_length + 1; // Convert end-relative position to start-relative.
	}
	if (occurrence_number > 0)
		--offset; // Convert 1-based to 0-based.
	if (offset < 0) // To the left of the first character.
		offset = 0; // LTR: start at offset 0.  RTL: search 0 characters.
	else if (offset > haystack_length) // To the right of the last character.
		offset = haystack_length; // LTR: start at null-terminator (omit all).  RTL: search whole string.

	LPTSTR found_pos;
	if (occurrence_number < 0) // Special mode to search from the right side (RTL).
	{
		if (!ParamIndexIsOmitted(3)) // For this mode, offset is only relevant if specified by the caller.
			haystack_length = offset; // i.e. omit any characters to the right of the starting position.
		found_pos = tcsrstr(haystack, haystack_length, needle, string_case_sense, -occurrence_number);
		_f_return_i(found_pos ? (found_pos - haystack + 1) : 0);  // +1 to convert to 1-based, since 0 indicates "not found".
	}
	// Since above didn't return:
	int i;
	for (i = 1, found_pos = haystack + offset; ; ++i, found_pos += needle_length)
		if (!(found_pos = tcsstr2(found_pos, needle, string_case_sense)) || i == occurrence_number)
			break;
	_f_return_i(found_pos ? (found_pos - haystack + 1) : 0);
}


ResultType RegExCreateMatchArray(LPCTSTR haystack, pcret *re, pcret_extra *extra, int *offset, int pattern_count, int captured_pattern_count, IObject *&match_object)
{
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
			// For simplicity and unlike PHP, IsNumeric() isn't called to forbid numeric subpattern names.
			// It seems the worst than could happen if it is numeric is that it would overlap/overwrite some of
			// the numerically-indexed elements in the output-array.  Seems pretty harmless given the rarity.
		}
	}
	//else one of the pcre_fullinfo() calls may have failed.  The PCRE docs indicate that this realistically never
	// happens unless bad inputs were given.  So due to rarity, just leave subpat_name==NULL; i.e. "no named subpatterns".

	LPTSTR mark = (extra->flags & PCRE_EXTRA_MARK) ? (LPTSTR)*extra->mark : NULL;
	return RegExMatchObject::Create(haystack, offset, subpat_name, pattern_count, captured_pattern_count, mark, match_object);
}


ResultType RegExMatchObject::Create(LPCTSTR aHaystack, int *aOffset, LPCTSTR *aPatternName
	, int aPatternCount, int aCapturedPatternCount, LPCTSTR aMark, IObject *&aNewObject)
{
	aNewObject = NULL;

	// If there was no match, seems best to not return an object:
	if (aCapturedPatternCount < 1)
		return OK;

	RegExMatchObject *m = new RegExMatchObject();
	if (!m)
		return FAIL;
	m->SetBase(sPrototype);

	if (  aMark && !(m->mMark = _tcsdup(aMark))  )
	{
		m->Release();
		return FAIL;
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
		return FAIL;
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
			return FAIL;
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
			m->Release(); // Also frees other things allocated above.
			return FAIL;
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
	aNewObject = m;
	return OK;
}


void RegExMatchObject::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	switch (aID)
	{
	case M_Count: _o_return(mPatternCount - 1);
	case M_Mark: _o_return(mMark ? mMark : _T(""));
	case M___Enum: _o_return(new IndexEnumerator(this, ParamIndexToOptionalInt(0, 1)
		, static_cast<IndexEnumerator::Callback>(&RegExMatchObject::GetEnumItem)));
	}

	int p;
	if (ParamIndexIsOmitted(0))
		p = 0;
	else if (ParamIndexIsNumeric(0))
		p = ParamIndexToInt(0);
	else if (!mPatternName) // There are no named subpatterns, so param 0 is invalid.
		p = -1;
	else
	{
		auto name = ParamIndexToString(0);
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
	if (p < 0 || p >= mPatternCount)
		_o_throw_param(0); // p != 0 implies the parameter was not omitted.

	switch (aID)
	{
	case M___Get:
		if (auto arr = dynamic_cast<Array*>(ParamIndexToObject(1)))
			if (arr->Length())
				_o_throw(ERR_INVALID_USAGE);
			// Otherwise, fall through:
	// Gives the correct result even if there was no match (because length is 0):
	case M_Value: _o_return(mHaystack - mHaystackStart + mOffset[p*2], mOffset[p*2+1]);
	case M_Pos: _o_return(mOffset[2*p] + 1);
	case M_Len: _o_return(mOffset[2*p + 1]);
	case M_Name: _o_return((mPatternName && mPatternName[p]) ? mPatternName[p] : _T(""));
	}
}


void *pcret_resolve_user_callout(LPCTSTR aCalloutParam, int aCalloutParamLength)
{
	// If no Func is found, pcre will handle the case where aCalloutParam is a pure integer.
	// In that case, the callout param becomes an integer between 0 and 255. No valid pointer
	// could be in this range, but we must take care to check (ptr>255) rather than (ptr!=NULL).
	auto callout_var = g_script.FindVar(aCalloutParam, aCalloutParamLength);
	return callout_var ? callout_var->ToObject() : nullptr;
}

struct RegExCalloutData // L14: Used by BIF_RegEx to pass necessary info to RegExCallout.
{
	pcret *re;
	LPTSTR re_text; // original NeedleRegEx
	int options_length; // used to adjust cb->pattern_position
	int pattern_count; // to save calling pcre_fullinfo unnecessarily for each callout
	pcret_extra *extra;
	ResultToken *result_token;
};

int RegExCallout(pcret_callout_block *cb)
{
	// It should be documented that (?C) is ignored if encountered by the hook thread,
	// which could happen if SetTitleMatchMode,Regex and "#HotIf Winactive/Exist" are used.
	// This would be a problem if the callout should affect the outcome of the match or 
	// should be called even if #HotIf WinA/E. will ultimately prevent the hotkey from firing. 
	// This is because:
	//	- The callout cannot be called from the hook thread, and therefore cannot affect
	//		the outcome of #HotIf WinA/E. when called by the hook thread.
	//	- If #HotIf WinA/E does NOT prevent the hotkey from firing, it will be reevaluated from
	//		the main thread before the hotkey is actually fired. This will allow any
	//		callouts to occur on the main thread.
	//  - By contrast, if #HotIf WinA/E DOES prevent the hotkey from firing, 
	//		#HotIf WinA/E will not be reevaluated from the main thread,
	//		so callouts cannot occur.
	if (GetCurrentThreadId() != g_MainThreadID)
		return 0;

	if (!cb->callout_data)
		return 0;
	RegExCalloutData &cd = *(RegExCalloutData *)cb->callout_data;

	auto callout_func = (IObject *)cb->user_callout;
	if (!callout_func)
	{
		Var *pcre_callout_var = g_script.FindVar(_T("pcre_callout"), 12, FINDVAR_FOR_READ); // This may be a local of the UDF which called RegExMatch/Replace().
		if (!pcre_callout_var)
			return 0; // Seems best to ignore the callout rather than aborting the match.
		callout_func = pcre_callout_var->ToObject();
		if (!callout_func)
		{
			if (!pcre_callout_var->HasContents()) // Var exists but is empty.
				return 0;
			// Could be an invalid name, or the name of a built-in function.
			cd.result_token->Error(_T("Invalid pcre_callout"));
			return PCRE_ERROR_CALLOUT;
		}
	}

	// Adjust offset to account for options, which are excluded from the regex passed to PCRE.
	cb->pattern_position += cd.options_length;
	
	// Save EventInfo to be restored when we return.
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

	// Since matching is still in progress, these *should* be -1.
	// For maintainability and peace of mind, save them anyway:
	int offset[] = { cb->offset_vector[0], cb->offset_vector[1] };
		
	// Temporarily set these for use by the function below:
	cb->offset_vector[0] = cb->start_match;
	cb->offset_vector[1] = cb->current_position;
	if (cd.extra->flags & PCRE_EXTRA_MARK)
		*cd.extra->mark = UorA(wchar_t *, UCHAR *) cb->mark;
	
	IObject *match_object;
	if (!RegExCreateMatchArray(cb->subject, cd.re, cd.extra, cb->offset_vector, cd.pattern_count, cb->capture_top, match_object))
	{
		cd.result_token->MemoryError();
		return PCRE_ERROR_CALLOUT; // Abort.
	}

	// Restore to former offsets (probably -1):
	cb->offset_vector[0] = offset[0];
	cb->offset_vector[1] = offset[1];
	
	// Make all string positions one-based. UPDATE: offset_vector cannot be modified, so for consistency don't do this:
	//++cb->pattern_position;
	//++cb->start_match;
	//++cb->current_position;
	
	ExprTokenType param[] =
	{
		match_object,
		cb->callout_number,
		cb->start_match + 1, // FoundPos (distinct from Match.Pos, which hasn't been set yet)
		const_cast<LPTSTR>(cb->subject), // Haystack
		cd.re_text // NeedleRegEx
	};
	__int64 number_to_return;
	auto result = CallMethod(callout_func, callout_func, nullptr, param, _countof(param), &number_to_return);
	if (result == FAIL || result == EARLY_EXIT)
	{
		number_to_return = PCRE_ERROR_CALLOUT;
		cd.result_token->SetExitResult(result);
	}
	
	g->EventInfo = EventInfo_saved;

	// Behaviour of return values is defined by PCRE.
	return (int)number_to_return;
}

pcret *get_compiled_regex(LPTSTR aRegEx, pcret_extra *&aExtra, int *aOptionsLength, ResultToken *aResultToken)
// Returns the compiled RegEx, or NULL on failure.
// This function is called by things other than built-in functions so it should be kept general-purpose.
// Upon failure, if aResultToken!=NULL:
//   - An exception is thrown with a descriptive message on failure.
//   - *aResultToken is set up to contain an empty string.
// Upon success, the following output parameters are set based on the options that were specified:
//    aGetPositionsNotSubstrings
//    aExtra
// L14: aOptionsLength is used by callouts to adjust cb->pattern_position to be relative to beginning of actual user-specified NeedleRegEx instead of string seen by PCRE.
{	
	if (!pcret_callout)
	{	// Ensure this is initialized, even for ::RegExMatch() (to allow (?C) in window title regexes).
		pcret_callout = &RegExCallout;
	}

	// While reading from or writing to the cache, don't allow another thread entry.  This is because
	// that thread (or this one) might write to the cache while the other one is reading/writing, which
	// could cause loss of data integrity (the hook thread can enter here via #HotIf WinActive/Exist & SetTitleMatchMode RegEx).
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
	// 20%: Download: Testing shows that it varies: e.g. microsoft.com uses `r`n, but `n is probably
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
		pcre_options = AHK_PCRE_CHARSET_OPTIONS;\
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
		case '\a':
			// Enable matching of any kind of newline, including Unicode newline characters.
			// v2: \R doesn't match Unicode newlines by default, so `a also enables that.
			pcre_options = (pcre_options & ~PCRE_NEWLINE_BITS) | PCRE_NEWLINE_ANY | PCRE_BSR_UNICODE;
			break; 
		case '\n':pcre_options = (pcre_options & ~PCRE_NEWLINE_BITS) | PCRE_NEWLINE_LF; break; // See below.
			// Above option: Could alternatively have called it "LF" rather than or in addition to "`n", but that
			// seems slightly less desirable due to potential overlap/conflict with future option letters,
			// plus the fact that `n should be pretty well known to AutoHotkey users, especially advanced ones
			// using RegEx.  Note: `n`r is NOT treated the same as `r`n because there's a slight chance PCRE
			// will someday support `n`r for some obscure usage (or just for symmetry/completeness).
			// The PCRE_NEWLINE_XXX options are valid for both compile() and exec(), but specifying it for exec()
			// would only serve to override the default stored inside the compiled pattern (seems rarely needed).
		case '\r':
			if (pat[1] == '\n')
			{
				++pat; // Skip over the second character so that it's not recognized as a separate option by the next iteration.
				pcre_options = (pcre_options & ~PCRE_NEWLINE_BITS) | PCRE_NEWLINE_CRLF; // Set explicitly in case it was unset by an earlier option. Remember that PCRE_NEWLINE_CRLF is a bitwise combination of PCRE_NEWLINE_LF and CR.
			}
			else // For completeness, it's easy to support PCRE_NEWLINE_CR too, though nowadays I think it's quite rare (former Macintosh format).
				pcre_options = (pcre_options & ~PCRE_NEWLINE_BITS) | PCRE_NEWLINE_CR; // Do it this way because PCRE_NEWLINE_CRLF is a bitwise combination of PCRE_NEWLINE_CR and PCRE_NEWLINE_LF.
			break;

		// Other options (uppercase so that lowercase can be reserved for future/PERL options):
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
	TCHAR error_buf[128];
	int error_code, error_offset;
	pcret *re_compiled;

	// COMPILE THE REGEX.
	if (   !(re_compiled = pcret_compile2(pat, pcre_options, &error_code, &error_msg, &error_offset, NULL))   )
	{
		if (aResultToken) // A non-NULL value indicates our caller is RegExMatch() or RegExReplace() in a script.
		{
			sntprintf(error_buf, _countof(error_buf), _T("Compile error %d at offset %d: %hs"), error_code
				, error_offset, error_msg);
			// Seems best to bring the error to the user's attention rather than letting it potentially
			// escape their notice.  This sort of error should be corrected immediately, not handled
			// within the script (such as by try-catch).
			aResultToken->Error(error_buf);
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
		// 2) Reduced code size.
		//if (error_msg)
		//{
			//if (aResultToken)
			//{
			//	sntprintf(error_buf, sizeof(error_buf), "Study error: %s", error_msg);
			//	aResultToken->Error(error_buf);
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
	aExtra = sCache[sLastFound].extra;
	if (aOptionsLength) // Lexikos: See aOptionsLength comment at beginning of this function.
		*aOptionsLength = sCache[sLastFound].options_length; 

	LeaveCriticalSection(&g_CriticalRegExCache);
	return sCache[sLastFound].re_compiled; // Indicate success.

error: // Since NULL is returned here, caller should ignore the contents of the output parameters.
	LeaveCriticalSection(&g_CriticalRegExCache);
	return NULL; // Indicate failure.
}



LPTSTR RegExMatch(LPTSTR aHaystack, LPTSTR aNeedleRegEx)
// Returns NULL if no match.  Otherwise, returns the address where the pattern was found in aHaystack.
{
	pcret_extra *extra;
	pcret *re;

	// Compile the regex or get it from cache.
	if (   !(re = get_compiled_regex(aNeedleRegEx, extra, NULL, NULL))   ) // Compiling problem.
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



void RegExReplace(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount
	, pcret *aRE, pcret_extra *aExtra, LPTSTR aHaystack, int aHaystackLength
	, int aStartingOffset, int aOffset[], int aNumberOfIntsInOffset)
{
	// If an output variable was provided for the count, resolve it early in case of early goto.
	// Fix for v1.0.47.05: In the unlikely event that output_var_count is the same script-variable as
	// as the haystack, needle, or replacement (i.e. the same memory), don't set output_var_count until
	// immediately prior to returning.  Otherwise, haystack, needle, or replacement would corrupted while
	// it's still being used here.
	Var *output_var_count = ParamIndexToOutputVar(3);
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
	size_t &result_length = aResultToken.marker_length; // MANDATORY FOR USERS OF MEM_TO_FREE: set marker_length to the length of the string.
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
			else // No replacements were actually done, so just return the original string to avoid malloc+memcpy
				 // (in addition, returning the original might help the caller make other optimizations).
			{
				aResultToken.marker = aHaystack;
				aResultToken.marker_length = aHaystackLength;
				
				// There's no need to do the following because it should already be that way when replacement_count==0.
				//if (result)
				//	free(result);
				//result = NULL; // This tells the caller that we already freed it (i.e. from its POV, we never allocated anything).
			}
			aResultToken.symbol = SYM_STRING;
			goto set_count_and_return; // All done.
		}

		// Otherwise:
		if (captured_pattern_count < 0) // An error other than "no match". These seem very rare, so it seems best to abort rather than yielding a partially-converted result.
		{
			if (!aResultToken.Exited()) // Checked in case a callout already exited/raised an error.
			{
				ITOA(captured_pattern_count, repl_buf);
				aResultToken.Error(ERR_PCRE_EXEC, repl_buf);
			}
			goto abort; // Goto vs. break to leave replacement_count set to 0.
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
								if (IsNumeric(substring_name, true, false, true)) // Seems best to allow floating point such as 1.0 because it will then get truncated to an integer.  It seems to rare that anyone would want to use floats as names.
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
	aResultToken.MemoryError();
abort:
	if (result)
	{
		free(result);  // Since result is probably an non-terminated string (not to mention an incompletely created result), it seems best to free it here to remove it from any further consideration by the caller.
		result = NULL; // Tell caller that it was freed.
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
	bool mode_is_replace = _f_callee_id == FID_RegExReplace;
	LPTSTR needle = ParamIndexToString(1, _f_number_buf); // Caller has already ensured that at least two actual parameters are present.

	pcret_extra *extra;
	pcret *re;
	int options_length;

	// COMPILE THE REGEX OR GET IT FROM CACHE.
	if (   !(re = get_compiled_regex(needle, extra, &options_length, &aResultToken))   ) // Compiling problem.
		return; // It already set aResultToken for us.

	// Since compiling succeeded, get info about other parameters.
	TCHAR haystack_buf[MAX_NUMBER_SIZE];
	size_t temp_length;
	LPTSTR haystack = ParamIndexToString(0, haystack_buf, &temp_length); // Caller has already ensured that at least two actual parameters are present.
	int haystack_length = (int)temp_length;

	int param_index = mode_is_replace ? 5 : 3;
	int starting_offset;
	if (ParamIndexIsOmitted(param_index))
		starting_offset = 0; // The one-based starting position in haystack (if any).  Convert it to zero-based.
	else
	{
		starting_offset = ParamIndexToInt(param_index);
		if (starting_offset < 0) // Same convention as SubStr(): Treat negative StartingPos as a position relative to the end of the string.
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
			starting_offset = haystack_length; // Due to rarity of this condition, opt for simplicity: just point it to the terminator, which is in essence an empty string (which will cause result in "no match" except when searcing for "").
		else
			--starting_offset; // Convert to zero-based.
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
	callout_data.result_token = &aResultToken;
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

	// SET THE RETURN VALUE BASED ON THE RESULTS OF EXECUTING THE EXPRESSION.
	if (captured_pattern_count == PCRE_ERROR_NOMATCH)
	{
		aResultToken.value_int64 = 0;
		// BUT CONTINUE ON so that the output variable (if any) is fully reset (made blank).
	}
	else if (captured_pattern_count < 0) // An error other than "no match".
	{
		if (aResultToken.Exited()) // A callout exited/raised an error.
			return;
		TCHAR err_info[MAX_INTEGER_SIZE];
		ITOA(captured_pattern_count, err_info);
		aResultToken.Error(ERR_PCRE_EXEC, err_info);
	}
	else // Match found, and captured_pattern_count >= 0 (but should never be 0 in this case because that only happens when offset[] is too small, which it isn't).
	{
		match_offset = offset[0];
		aResultToken.value_int64 = match_offset + 1; // i.e. the position of the entire-pattern match is the function's return value.
	}

	Var *output_var = ParamIndexToOutputVar(2);
	if (!output_var)
		return;

	IObject *match_object;
	if (!RegExCreateMatchArray(haystack, re, extra, offset, pattern_count, captured_pattern_count, match_object))
		aResultToken.MemoryError();
	if (match_object)
		output_var->AssignSkipAddRef(match_object);
	else // Out-of-memory or there were no captured patterns.
		output_var->Assign();
}



BIF_DECL(BIF_Ord)
{
	// Result will always be an integer (this simplifies scripts that work with binary zeros since an
	// empty string yields zero).
	LPTSTR cp = ParamIndexToString(0, _f_number_buf);
#ifndef UNICODE
	// This section could convert a single- or multi-byte character to a Unicode code point
	// for consistency, but since the ANSI build won't be officially supported that doesn't
	// seem necessary.
#else
	if (IS_SURROGATE_PAIR(cp[0], cp[1]))
		_f_return_i(((cp[0] - 0xd800) << 10) + (cp[1] - 0xdc00) + 0x10000);
	else
#endif
		_f_return_i((TBYTE)*cp);
}



BIF_DECL(BIF_Chr)
{
	Throw_if_Param_NaN(0);
	int param1 = ParamIndexToInt(0); // Convert to INT vs. UINT so that negatives can be detected.
	LPTSTR cp = _f_retval_buf; // If necessary, it will be moved to a persistent memory location by our caller.
	int len;
	if (param1 < 0 || param1 > UorA(0x10FFFF, UCHAR_MAX))
		_f_throw_param(0);
#ifdef UNICODE
	else if (param1 >= 0x10000)
	{
		param1 -= 0x10000;
		cp[0] = 0xd800 + ((param1 >> 10) & 0x3ff);
		cp[1] = 0xdc00 + ( param1        & 0x3ff);
		cp[2] = '\0';
		len = 2;
	}
#else
	// See #ifndef in BIF_Ord for related comment.
#endif
	else
	{
		cp[0] = param1;
		cp[1] = '\0';
		len = 1; // Even if param1 == 0.
	}
	_f_return_p(cp, len);
}



struct NumGetParams
{
	size_t target, right_side_bound;
	size_t num_size = sizeof(DWORD_PTR);
	BOOL is_integer = TRUE, is_signed = FALSE;
};

void ConvertNumGetType(ExprTokenType &aToken, NumGetParams &op)
{
	LPTSTR type = TokenToString(aToken); // No need to pass aBuf since any numeric value would not be recognized anyway.
	if (ctoupper(*type) == 'U') // Unsigned.
	{
		++type; // Remove the first character from further consideration.
		op.is_signed = FALSE;
	}
	else
		op.is_signed = TRUE;

	switch(ctoupper(*type)) // Override "size" and aResultToken.symbol if type warrants it. Note that the above has omitted the leading "U", if present, leaving type as "Int" vs. "Uint", etc.
	{
	case 'P': // Nothing extra needed in this case.
		op.num_size = sizeof(void *), op.is_integer = TRUE;
		break;
	case 'I':
		if (_tcschr(type, '6')) // Int64. It's checked this way for performance, and to avoid access violation if string is bogus and too short such as "i64".
			op.num_size = 8, op.is_integer = TRUE;
		else
			op.num_size = 4, op.is_integer = TRUE;
		break;
	case 'S': op.num_size = 2, op.is_integer = TRUE; break; // Short.
	case 'C': op.num_size = 1, op.is_integer = TRUE; break; // Char.

	case 'D': op.num_size = 8, op.is_integer = FALSE; break; // Double.
	case 'F': op.num_size = 4, op.is_integer = FALSE; break; // Float.

	default: op.num_size = 0; break;
	}
}

void *BufferObject::sVTable = getVTable(); // Placing this here vs. in script_object.cpp improves some simple benchmarks by as much as 7%.

void GetBufferObjectPtr(ResultToken &aResultToken, IObject *obj, size_t &aPtr, size_t &aSize)
{
	if (BufferObject::IsInstanceExact(obj))
	{
		// Some primitive benchmarks showed that this was about as fast as passing
		// a pointer directly, whereas invoking the properties (below) doubled the
		// overall time taken by NumGet/NumPut.
		aPtr = (size_t)((BufferObject *)obj)->Data();
		aSize = ((BufferObject *)obj)->Size();
	}
	else
	{
		if (GetObjectPtrProperty(obj, _T("Ptr"), aPtr, aResultToken))
			GetObjectPtrProperty(obj, _T("Size"), aSize, aResultToken);
	}
}

void GetBufferObjectPtr(ResultToken &aResultToken, IObject *obj, size_t &aPtr)
// See above for comments.
{
	if (BufferObject::IsInstanceExact(obj))
		aPtr = (size_t)((BufferObject *)obj)->Data();
	else
		GetObjectPtrProperty(obj, _T("Ptr"), aPtr, aResultToken);
}

void ConvertNumGetTarget(ResultToken &aResultToken, ExprTokenType &target_token, NumGetParams &op)
{
	if (IObject *obj = TokenToObject(target_token))
	{
		GetBufferObjectPtr(aResultToken, obj, op.target, op.right_side_bound);
		if (aResultToken.Exited())
			return;
		op.right_side_bound += op.target;
	}
	else
	{
		op.target = (size_t)TokenToInt64(target_token);
		op.right_side_bound = SIZE_MAX;
	}
}


BIF_DECL(BIF_NumGet)
{
	NumGetParams op;
	ConvertNumGetTarget(aResultToken, *aParam[0], op);
	if (aResultToken.Exited())
		return;
	if (aParamCount > 2) // Offset was specified.
	{
		op.target += (ptrdiff_t)TokenToInt64(*aParam[1]);
		aParam++;
	}
	// MinParams ensures there is always one more parameter.
	ConvertNumGetType(*aParam[1], op);

	// If the target is a variable, the following check ensures that the memory to be read lies within its capacity.
	// This seems superior to an exception handler because exception handlers only catch illegal addresses,
	// not ones that are technically legal but unintentionally bugs due to being beyond a variable's capacity.
	// Moreover, __try/except is larger in code size. Another possible alternative is IsBadReadPtr()/IsBadWritePtr(),
	// but those are discouraged by MSDN.
	// The following aren't covered by the check below:
	// - Due to rarity of negative offsets, only the right-side boundary is checked, not the left.
	// - Due to rarity and to simplify things, Float/Double aren't checked.
	if (!op.num_size
		|| op.target < 65536 // Basic sanity check to catch incoming raw addresses that are zero or blank.  On Win32, the first 64KB of address space is always invalid.
		|| op.target+op.num_size > op.right_side_bound) // i.e. it's ok if target+size==right_side_bound because the last byte to be read is actually at target+size-1. In other words, the position of the last possible terminator within the variable's capacity is considered an allowable address.
	{
		_f_throw_value(ERR_PARAM_INVALID);
	}

	switch (op.num_size)
	{
	case 4: // Listed first for performance.
		if (!op.is_integer)
			aResultToken.value_double = *(float *)op.target;
		else if (op.is_signed)
			aResultToken.value_int64 = *(int *)op.target; // aResultToken.symbol defaults to SYM_INTEGER.
		else
			aResultToken.value_int64 = *(unsigned int *)op.target;
		break;
	case 8:
		if (op.is_integer)
			// Unsigned 64-bit integers aren't supported because variables/expressions can't support them.
			aResultToken.value_int64 = *(__int64 *)op.target;
		else
			aResultToken.value_double = *(double *)op.target;
		break;
	case 2:
		if (op.is_signed) // Don't use ternary because that messes up type-casting.
			aResultToken.value_int64 = *(short *)op.target;
		else
			aResultToken.value_int64 = *(unsigned short *)op.target;
		break;
	case 1:
		if (op.is_signed) // Don't use ternary because that messes up type-casting.
			aResultToken.value_int64 = *(char *)op.target;
		else
			aResultToken.value_int64 = *(unsigned char *)op.target;
		break;
	}
	if (!op.is_integer)
		aResultToken.symbol = SYM_FLOAT;
}



BIF_DECL(BIF_Format)
{
	if (TokenIsPureNumeric(*aParam[0]))
		_f_return_p(ParamIndexToString(0, _f_retval_buf));

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
			if (param >= aParamCount || param < 1) // Invalid parameter index.
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
			return;
		aResultToken.symbol = SYM_STRING;
		target = aResultToken.marker;
	}
}



BIF_DECL(BIF_NumPut)
{
	// Params can be any non-zero number of type-number pairs, followed by target[, offset].
	// Prior validation has ensured that there are at least three parameters.
	//   NumPut(t1, n1, t2, n2, p, o)
	//   NumPut(t1, n1, t2, n2, p)
	//   NumPut(t1, n1, p, o)
	//   NumPut(t1, n1, p)
	
	// Split target[,offset] from aParam.
	bool offset_was_specified = !(aParamCount & 1);
	aParamCount -= 1 + int(offset_was_specified);
	ExprTokenType &target_token = *aParam[aParamCount];
	
	NumGetParams op;
	ConvertNumGetTarget(aResultToken, target_token, op);
	if (aResultToken.Exited())
		return;
	if (offset_was_specified)
		op.target += (ptrdiff_t)TokenToInt64(*aParam[aParamCount + 1]);

	size_t num_end;
	for (int n_param = 1; n_param < aParamCount; n_param += 2, op.target = num_end)
	{
		ConvertNumGetType(*aParam[n_param - 1], op); // Type name.
		ExprTokenType &token_to_write = *aParam[n_param]; // Numeric value.

		num_end = op.target + op.num_size; // This is used below and also as NumPut's return value. It's the address to the right of the item to be written.

		// See comments in NumGet about the following section:
		if (!op.num_size
			|| !TokenIsNumeric(token_to_write)
			|| op.target < 65536 // Basic sanity check to catch incoming raw addresses that are zero or blank.  On Win32, the first 64KB of address space is always invalid.
			|| num_end > op.right_side_bound) // i.e. it's ok if target+size==right_side_bound because the last byte to be read is actually at target+size-1. In other words, the position of the last possible terminator within the variable's capacity is considered an allowable address.
		{
			_f_throw_value(ERR_PARAM_INVALID);
		}

		union
		{
			__int64 num_i64;
			double num_f64;
			float num_f32;
		};

		// Note that since v2.0-a083-97803aeb, TokenToInt64 supports conversion of large unsigned 64-bit
		// numbers from strings (producing a negative value, but with the right bit representation).
		if (op.is_integer)
			num_i64 = TokenToInt64(token_to_write);
		else
		{
			num_f64 = TokenToDouble(token_to_write);
			if (op.num_size == 4)
				num_f32 = (float)num_f64;
		}

		// This method benchmarked marginally faster than memcpy for the multi-param mode.
		switch (op.num_size)
		{
		case 8: *(UINT64 *)op.target = (UINT64)num_i64; break;
		case 4: *(UINT32 *)op.target = (UINT32)num_i64; break;
		case 2: *(UINT16 *)op.target = (UINT16)num_i64; break;
		case 1: *(UINT8 *)op.target = (UINT8)num_i64; break;
		}
	}
	if (target_token.symbol == SYM_VAR && !target_token.var->IsPureNumeric())
		target_token.var->Close(); // This updates various attributes of the variable.
	//else the target was an raw address.  If that address is inside some variable's contents, the above
	// attributes would already have been removed at the time the & operator was used on the variable.
	aResultToken.value_int64 = num_end; // aResultToken.symbol was set to SYM_INTEGER by our caller.
}



BIF_DECL(BIF_StrGetPut) // BIF_DECL(BIF_StrGet), BIF_DECL(BIF_StrPut)
{
	// To simplify flexible handling of parameters:
	ExprTokenType **aParam_end = aParam + aParamCount, **next_param = aParam;

	LPCVOID source_string; // This may hold an intermediate UTF-16 string in ANSI builds.
	int source_length;
	if (_f_callee_id == FID_StrPut)
	{
		// StrPut(String, Address[, Length][, Encoding])
		ExprTokenType &source_token = *aParam[0];
		source_string = (LPCVOID)TokenToString(source_token, _f_number_buf); // Safe to use _f_number_buf since StrPut won't use it for the result.
		source_length = (int)((source_token.symbol == SYM_VAR) ? source_token.var->CharLength() : _tcslen((LPCTSTR)source_string));
		++next_param; // Remove the String param from further consideration.
	}
	else
	{
		// StrGet(Address[, Length][, Encoding])
		source_string = NULL;
		source_length = 0;
	}

	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T(""); // Set default in case of early return.

	IObject *buffer_obj;
	LPVOID 	address;
	size_t  max_bytes = SIZE_MAX;
	int 	length = -1; // actual length
	bool	length_is_max_size = false;
	UINT 	encoding = UorA(CP_UTF16, CP_ACP); // native encoding

	// Parameters are interpreted according to the following rules (highest to lowest precedence):
	// Legend:  StrPut(String[, X, Y, Z])  or  StrGet(Address[, Y, Z])
	// - If X is non-numeric, it is Encoding.  Calculates required buffer size but does nothing else.  Y and Z must be omitted.
	// - If X is numeric, it is Address.  (For StrGet, non-numeric Address is treated as an error.)
	// - If Y is numeric, it is Length.  Otherwise "Actual length" is assumed.
	// - If a parameter remains, it is Encoding.
	// Encoding may therefore only be purely numeric if Address(X) and Length(Y) are specified.

	const LPVOID FIRST_VALID_ADDRESS = (LPVOID)65536;

	if (next_param < aParam_end && TokenIsNumeric(**next_param))
	{
		address = (LPVOID)TokenToInt64(**next_param);
		++next_param;
	}
	else if (next_param < aParam_end && (buffer_obj = TokenToObject(**next_param)))
	{
		size_t ptr;
		GetBufferObjectPtr(aResultToken, buffer_obj, ptr, max_bytes);
		if (aResultToken.Exited())
			return;
		address = (LPVOID)ptr;
		++next_param;
	}
	else
	{
		if (!source_string || aParamCount > 2)
		{
			// See the "Legend" above.  Either this is StrGet and Address was invalid (it can't be omitted due
			// to prior min-param checks), or it is StrPut and there are too many parameters.
			_f_throw_value(source_string ? ERR_PARAM_INVALID : ERR_PARAM1_INVALID);  // StrPut : StrGet
		}
		// else this is the special measuring mode of StrPut, where Address and Length are omitted.
		// A length of 0 when passed to the Win API conversion functions (or the code below) means
		// "calculate the required buffer size, but don't do anything else."
		length = 0;
		address = FIRST_VALID_ADDRESS; // Skip validation below; address should never be used when length == 0.
	}

	if (next_param < aParam_end)
	{
		if (length == -1) // i.e. not StrPut(String, Encoding)
		{
			if (TokenIsNumeric(**next_param)) // Length parameter
			{
				length = (int)TokenToInt64(**next_param);
				if (!source_string) // StrGet
				{
					if (length == 0)
						return; // Get 0 chars.
					if (length < 0)
						length = -length; // Retrieve exactly this many chars, even if there are null chars.
					else
						length_is_max_size = true; // Limit to this, but stop at the first null char.
				}
				else if (length <= 0)
					_f_throw_value(ERR_INVALID_LENGTH);
				++next_param; // Let encoding be the next param, if present.
			}
			else if ((*next_param)->symbol == SYM_MISSING)
			{
				// Length was "explicitly omitted", as in StrGet(Address,, Encoding),
				// which allows Encoding to be an integer without specifying Length.
				++next_param;
			}
			// aParam now points to aParam_end or the Encoding param.
		}
		if (next_param < aParam_end)
		{
			encoding = Line::ConvertFileEncoding(**next_param);
			if (encoding == -1)
				_f_throw_value(ERR_INVALID_ENCODING);
		}
	}
	// Note: CP_AHKNOBOM is not supported; "-RAW" must be omitted.

	// Check for obvious errors to prevent an Access Violation.
	// Address can be zero for StrPut if length is also zero (see below).
	if ( address < FIRST_VALID_ADDRESS
		// Also check for overlap, in case memcpy is used instead of MultiByteToWideChar/WideCharToMultiByte.
		// (Behaviour for memcpy would be "undefined", whereas MBTWC/WCTBM would fail.)  Overlap in the
		// other direction (source_string beginning inside address..length) should not be possible.
		|| (address >= source_string && address <= ((LPTSTR)source_string + source_length))
		// The following catches StrPut(X, &Y) where Y is uninitialized or has zero capacity.
		|| (address == Var::sEmptyString && source_length) )
	{
		_f_throw_param(source_string ? 1 : 0);
	}

	if (max_bytes != SIZE_MAX)
	{
		// Target is a Buffer object with known size, so limit length accordingly.
		int max_chars = int(max_bytes >> int(encoding == CP_UTF16));
		if (length > max_chars)
			_f_throw_value(ERR_INVALID_LENGTH);
		if (source_length > max_chars)
			_f_throw_param(1);
		if (length == -1)
		{
			length = max_chars;
			length_is_max_size = true;
		}
	}

	if (source_string) // StrPut
	{
		int char_count; // Either bytes or characters, depending on the target encoding.
		aResultToken.symbol = SYM_INTEGER; // Most paths below return an integer.

		if (!source_length)
		{
			// Take a shortcut when source_string is empty, since some paths below might not handle it correctly.
			if (length) // true except when in measuring mode.
			{
				if (encoding == CP_UTF16)
					*(LPWSTR)address = '\0';
				else
					*(LPSTR)address = '\0';
			}
			aResultToken.value_int64 = encoding == CP_UTF16 ? sizeof(WCHAR) : sizeof(CHAR);
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
					_f_throw_value(ERR_INVALID_LENGTH);
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
						aResultToken.value_int64 = char_count * (1 + (encoding == CP_UTF16));
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
							_f_throw_win32();
					}
					++char_count; // + 1 for null-terminator (source_length causes it to be excluded from char_count).
					if (length == 0) // Caller just wants the required buffer size.
					{
						aResultToken.value_int64 = char_count * (1 + (encoding == CP_UTF16));
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
			if (!char_count)
				_f_throw_win32();
		}
		// Return the number of bytes written.
		aResultToken.value_int64 = char_count * (1 + (encoding == CP_UTF16));
	}
	else // StrGet
	{
		if (length_is_max_size) // Implies length != -1.
		{
			// Caller specified the maximum character count, not the exact length.
			// If the length includes null characters, the conversion functions below
			// would convert more than necessary and we'd still have to recalculate the
			// length.  So find the exact length up front:
			if (encoding == CP_UTF16)
				length = (int)wcsnlen((LPWSTR)address, length);
			else
				length = (int)strnlen((LPSTR)address, length);
		}
		if (encoding != UorA(CP_UTF16, CP_ACP))
		{
			// Conversion is required.
			int conv_length;
			// MS docs: "Note that, if cbMultiByte is 0, the function fails."
			if (!length)
				_f_return_empty;
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
			{
				aResult = FAIL;
				return; // Out of memory.
			}
			conv_length = WideCharToMultiByte(CP_ACP, WC_NO_BEST_FIT_CHARS, (LPCWSTR)address, length, aResultToken.marker, conv_length, NULL, NULL);
#endif
			if (!conv_length) // This can only be failure, since ... (see below)
				_f_throw_win32();
			if (length == -1) // conv_length includes a null-terminator in this case.
				--conv_length;
			else
				aResultToken.marker[conv_length] = '\0'; // It wasn't terminated above.
			aResultToken.marker_length = conv_length; // Update it.
		}
		else if (length == -1)
		{
			// Return this null-terminated string, no conversion necessary.
			aResultToken.marker = (LPTSTR) address;
			aResultToken.marker_length = _tcslen(aResultToken.marker);
		}
		else
		{
			// No conversion necessary, but we might not want the whole string.
			// Copy and null-terminate the string; some callers might require it.
			TokenSetResult(aResultToken, (LPCTSTR)address, length);
		}
	}
}



BIF_DECL(BIF_StrPtr)
{
	switch (aParam[0]->symbol)
	{
	case SYM_STRING:
		_f_return((UINT_PTR)aParam[0]->marker);
	case SYM_VAR:
		if (!aParam[0]->var->IsPureNumericOrObject())
			_f_return((UINT_PTR)aParam[0]->var->Contents());
	default:
		_f_throw_type(_T("String"), *aParam[0]);
	}
}



BIF_DECL(BIF_IsLabel)
{
	_f_return_b(g_script.FindLabel(ParamIndexToString(0, _f_number_buf)) ? 1 : 0);
}


IObject *UserFunc::CloseIfNeeded()
{
	FreeVars *fv = (mUpVarCount && mOuterFunc && sFreeVars) ? sFreeVars->ForFunc(mOuterFunc) : NULL;
	if (!fv)
	{
		AddRef();
		return this;
	}
	return new Closure(this, fv, false);
}



BIF_DECL(BIF_IsTypeish)
{
	auto variable_type = (VariableTypeType)_f_callee_id;
	bool if_condition;
	TCHAR *cp;

	// The first set of checks are for isNumber(), isInteger() and isFloat(), which permit pure numeric values.
	switch (TypeOfToken(*aParam[0]))
	{
	case SYM_INTEGER:
		switch (variable_type)
		{
		case VAR_TYPE_NUMBER:
		case VAR_TYPE_INTEGER:
			_f_return_b(true);
		case VAR_TYPE_FLOAT:
			_f_return_b(false);
		default:
			// Do not permit pure numbers for the other functions, since the results would not be intuitive.
			// For instance, isAlnum() would return false for negative values due to '-'; isXDigit() would
			// return true for positive integers even though they are always in decimal.
			goto type_mismatch;
		}
	case SYM_FLOAT:
		switch (variable_type)
		{
		case VAR_TYPE_NUMBER:
		case VAR_TYPE_FLOAT:
			_f_return_b(true);
		case VAR_TYPE_INTEGER:
			// Given that legacy "if var is float" required a decimal point in var and isFloat() is false for
			// integers which can be represented as float, it seems inappropriate for isInteger(1.0) to be true.
			// A function like isWholeNumber() could be added if that was needed.
			_f_return_b(false);
		default:
			goto type_mismatch;
		}
	case SYM_OBJECT:
		switch (variable_type)
		{
		case VAR_TYPE_NUMBER:
		case VAR_TYPE_INTEGER:
		case VAR_TYPE_FLOAT:
			_f_return_b(false);
		default:
			goto type_mismatch;
		}
	}
	// Since above did not return or goto, the value is a string.
	LPTSTR aValueStr = ParamIndexToString(0);
	auto string_case_sense = ParamIndexToCaseSense(1); // For IsAlpha, IsAlnum, IsUpper, IsLower.
	switch (string_case_sense)
	{
	case SCS_INSENSITIVE: // This case also executes when the parameter is omitted, such as for functions which don't have this parameter.
	case SCS_SENSITIVE: // Insensitive vs. sensitive doesn't mean anything for these functions, but seems fair to allow either, rather than requiring 0 or deviating from the CaseSense convention by requiring "".
	case SCS_INSENSITIVE_LOCALE: // 'Locale'
		break;
	default:
		_f_throw_param(1);
	}

	// The remainder of this function is based on the original code for ACT_IFIS, which was removed
	// in commit 3382e6e2.
	switch (variable_type)
	{
	case VAR_TYPE_NUMBER:
		if_condition = IsNumeric(aValueStr, true, false, true);
		break;
	case VAR_TYPE_INTEGER:
		if_condition = IsNumeric(aValueStr, true, false, false);  // Passes false for aAllowFloat.
		break;
	case VAR_TYPE_FLOAT:
		if_condition = (IsNumeric(aValueStr, true, false, true) == PURE_FLOAT);
		break;
	case VAR_TYPE_TIME:
	{
		SYSTEMTIME st;
		// Also insist on numeric, because even though YYYYMMDDToFileTime() will properly convert a
		// non-conformant string such as "2004.4", for future compatibility, we don't want to
		// report that such strings are valid times:
		if_condition = IsNumeric(aValueStr, false, false, false) && YYYYMMDDToSystemTime(aValueStr, st, true);
		break;
	}
	case VAR_TYPE_DIGIT:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			if (*cp < '0' || *cp > '9') // Avoid iswdigit; as documented, only ASCII digits 0 .. 9 are permitted.
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_XDIGIT:
		cp = aValueStr;
		if (!_tcsnicmp(cp, _T("0x"), 2)) // Allow 0x prefix.
			cp += 2;
		if_condition = true;
		for (; *cp; ++cp)
			if (!cisxdigit(*cp)) // Avoid iswxdigit; as documented, only ASCII xdigits are permitted.
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_ALNUM:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			if (string_case_sense == SCS_INSENSITIVE_LOCALE ? !IsCharAlphaNumeric(*cp) : !cisalnum(*cp))
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_ALPHA:
		// Like AutoIt3, the empty string is considered to be alphabetic, which is only slightly debatable.
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			if (string_case_sense == SCS_INSENSITIVE_LOCALE ? !IsCharAlpha(*cp) : !cisalpha(*cp))
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_UPPER:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			if (string_case_sense == SCS_INSENSITIVE_LOCALE ? !IsCharUpper(*cp) : !cisupper(*cp))
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_LOWER:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			if (string_case_sense == SCS_INSENSITIVE_LOCALE ? !IsCharLower(*cp) : !cislower(*cp))
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_SPACE:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			if (!_istspace(*cp))
			{
				if_condition = false;
				break;
			}
		break;
#ifdef DEBUG
	default:
		MsgBox(_T("DEBUG: Unhandled IsXStr mode."));
#endif
	}
	_f_return_b(if_condition);

type_mismatch:
	_f_throw_type(_T("String"), *aParam[0]);
}



BIF_DECL(BIF_IsSet)
{
	Var *var = ParamIndexToOutputVar(0);
	if (!var)
		_f_throw_param(0, _T("variable reference"));
	_f_return_b(!var->IsUninitializedNormalVar());
}



BIF_DECL(BIF_GetKeyState)
{
	TCHAR key_name_buf[MAX_NUMBER_SIZE]; // Because _f_retval_buf is used for something else below.
	LPTSTR key_name = ParamIndexToString(0, key_name_buf);
	// Keep this in sync with GetKeyJoyState().
	// See GetKeyJoyState() for more comments about the following lines.
	JoyControls joy;
	int joystick_id;
	vk_type vk = TextToVK(key_name);
	if (!vk)
	{
		if (   !(joy = (JoyControls)ConvertJoy(key_name, &joystick_id))   )
		{
			// It is neither a key name nor a joystick button/axis.
			_f_throw_param(0);
		}
		ScriptGetJoyState(joy, joystick_id, aResultToken, _f_retval_buf);
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
	LPTSTR key = ParamIndexToString(0, _f_number_buf);
	vk_type vk;
	sc_type sc;
	TextToVKandSC(key, vk, sc);

	switch (_f_callee_id)
	{
	case FID_GetKeyVK:
		_f_return_i(vk ? vk : sc_to_vk(sc));
	case FID_GetKeySC:
		_f_return_i(sc ? sc : vk_to_sc(vk));
	//case FID_GetKeyName:
	default:
		_f_return_p(GetKeyName(vk, sc, _f_retval_buf, _f_retval_buf_size, _T("")));
	}
}



BIF_DECL(BIF_VarSetStrCapacity)
// Returns: The variable's new capacity.
// Parameters:
// 1: Target variable (unquoted).
// 2: Requested capacity.
{
	Var *target_var = ParamIndexToOutputVar(0);
	// Redundant due to prior validation of OutputVars:
	//if (!target_var)
	//	_f_throw_param(0, _T("variable reference"));
	Var &var = *target_var;
	ASSERT(var.Type() == VAR_NORMAL); // Should always be true.

	if (aParamCount > 1) // Second parameter is present.
	{
		__int64 param1 = TokenToInt64(*aParam[1]);
		// Check for invalid values, in particular small negatives which end up large when converted
		// to unsigned.  Var::AssignString used to have a similar check, but integer overflow caused
		// by "* sizeof(TCHAR)" allowed some errors to go undetected.  For this same reason, we can't
		// simply rely on SetCapacity() calling malloc() and then detecting failure.
		if ((unsigned __int64)param1 > (MAXINT_PTR / sizeof(TCHAR)))
		{
			if (param1 == -1) // Adjust variable's internal length.
			{
				var.SetLengthFromContents();
				// Seems more useful to report length vs. capacity in this special case. Scripts might be able
				// to use this to boost performance.
				aResultToken.value_int64 = var.CharLength();
				return;
			}
			// x64: it's negative but not -1.
			// x86: it's either >2GB or negative and not -1.
			_f_throw_param(1);
		}
		// Since above didn't return:
		size_t new_capacity = (size_t)param1 * sizeof(TCHAR); // Chars to bytes.
		if (new_capacity)
		{
			if (!var.SetCapacity(new_capacity, true)) // This also destroys the variables contents.
			{
				aResultToken.SetExitResult(FAIL); // ScriptError() was already called.
				return;
			}
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
		// it's an input var rather than an output var, so check if it contains a string.
		if (var.IsPureNumericOrObject())
			_f_throw_type(_T("String"), *aParam[0]);
	}

	// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
	if (aResultToken.value_int64 = var.CharCapacity()) // Don't subtract 1 here in lieu doing it below (avoids underflow).
		aResultToken.value_int64 -= 1; // Omit the room for the zero terminator since script capacity is defined as length vs. size.
}



BIF_DECL(BIF_FileExist)
{
	TCHAR filename_buf[MAX_NUMBER_SIZE]; // Because _f_number_buf is used for something else below.
	LPTSTR filename = ParamIndexToString(0, filename_buf);
	LPTSTR buf = _f_retval_buf; // If necessary, it will be moved to a persistent memory location by our caller.
	DWORD attr;
	if (DoesFilePatternExist(filename, &attr, _f_callee_id == FID_DirExist ? FILE_ATTRIBUTE_DIRECTORY : 0))
	{
		// Yield the attributes of the first matching file.  If not match, yield an empty string.
		// This relies upon the fact that a file's attributes are never legitimately zero, which
		// seems true but in case it ever isn't, this forces a non-empty string be used.
		// UPDATE for v1.0.44.03: Someone reported that an existing file (created by NTbackup.exe) can
		// apparently have undefined/invalid attributes (i.e. attributes that have no matching letter in
		// "RASHNDOCT").  Although this is unconfirmed, it's easy to handle that possibility here by
		// checking for a blank string.  This allows FileExist() to report boolean TRUE rather than FALSE
		// for such "mystery files":
		FileAttribToStr(buf, attr);
		if (!*buf) // See above.
		{
			// The attributes might be all 0, but more likely the file has some of the newer attributes
			// such as FILE_ATTRIBUTE_ENCRYPTED (or has undefined attributes).  So rather than storing attr as
			// a hex number (which could be zero and thus defeat FileExist's ability to detect the file), it
			// seems better to store some arbitrary letter (other than those in "RASHNDOCT") so that FileExist's
			// return value is seen as boolean "true".
			buf[0] = 'X';
			buf[1] = '\0';
		}
	}
	else // Empty string is the indicator of "not found" (seems more consistent than using an integer 0, since caller might rely on it being SYM_STRING).
		*buf = '\0';
	_f_return_p(buf);
}



BIF_DECL(BIF_WinExistActive)
{
	bool hwnd_specified = false;
	HWND hwnd;
	if (!ParamIndexIsOmitted(0))
	{
		switch (DetermineTargetHwnd(hwnd, aResultToken, *aParam[0]))
		{
		case FAIL:
			return;
		case OK:
			hwnd_specified = true;
			// DetermineTargetHwnd() already called IsWindow() to verify hwnd.
			// g->DetectHiddenWindows is intentionally ignored for these cases.
			if (_f_callee_id == FID_WinActive && GetForegroundWindow() != hwnd)
				hwnd = 0;
			if (hwnd)
				g->hWndLastUsed = hwnd;
			break;
		}
	}

	if (!hwnd_specified) // Do not call WinExist()/WinActive() even if the specified hwnd was 0.
	{
		TCHAR *param[4], param_buf[4][MAX_NUMBER_SIZE];
		for (int j = 0; j < 4; ++j) // For each formal parameter, including optional ones.
			param[j] = ParamIndexToOptionalString(j, param_buf[j]);

		hwnd = _f_callee_id == FID_WinExist
			? WinExist(*g, param[0], param[1], param[2], param[3], false, true)
			: WinActive(*g, param[0], param[1], param[2], param[3], true);
	}

	_f_return_i((size_t)hwnd);
}



BIF_DECL(BIF_Round)
// For simplicity, this always yields something numeric (or a string that's numeric).
// Even Round(empty_or_unintialized_var) is zero rather than "" or "NaN".
{
	// In the future, a string conversion algorithm might be better to avoid the loss
	// of 64-bit integer precision that is currently caused by the use of doubles in
	// the calculation:
	int param2;
	double multiplier;
	if (aParamCount > 1)
	{
		Throw_if_Param_NaN(1);
		param2 = ParamIndexToInt(1);
		multiplier = qmathPow(10, param2);
	}
	else // Omitting the parameter is the same as explicitly specifying 0 for it.
	{
		param2 = 0;
		multiplier = 1;
	}
	Throw_if_Param_NaN(0);
	double value = ParamIndexToDouble(0);
	value = (value >= 0.0 ? qmathFloor(value * multiplier + 0.5)
		: qmathCeil(value * multiplier - 0.5)) / multiplier;

	// If incoming value is an integer, it seems best for flexibility to convert it to a
	// floating point number whenever the second param is >0.  That way, it can be used
	// to "cast" integers into floats.  Conversely, it seems best to yield an integer
	// whenever the second param is <=0 or omitted.
	if (param2 > 0)
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
		// AS DOCUMENTED: Round(..., positive_number) displays exactly positive_number decimal places, which
		// might be inconsistent with normal float->string conversion.  If it wants, the script can force the
		// result to be formatted the normal way (omitting trailing 0s) by adding 0 to the result.
		// Also, a new parameter an be added someday to trim excess trailing zeros from param2>0's result
		// (e.g. Round(3.50, 2, true) can be 3.5 rather than 3.50), but this seems less often desired due to
		// column alignment and other goals where consistency is important.
		LPTSTR buf = _f_retval_buf;
		int len = _stprintf(buf, _T("%0.*f"), param2, value); // %f can handle doubles in MSVC++.
		_f_return_p(buf, len);
	}
	else
		// Fix for v1.0.47.04: See BIF_FloorCeil() for explanation of this fix.  Currently, the only known example
		// of when the fix is necessary is the following script in release mode (not debug mode):
		//   myNumber  := 1043.22  ; Bug also happens with -1043.22 (negative).
		//   myRounded1 := Round( myNumber, -1 )  ; Stores 1040 (correct).
		//   ChartModule := DllCall("LoadLibrary", "str", "rmchart.dll")
		//   myRounded2 := Round( myNumber, -1 )  ; Stores 1039 (wrong).
		_f_return_i((__int64)(value + (value > 0 ? 0.2 : -0.2)));
		// Formerly above simply returned (__int64)value.
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
	Throw_if_Param_NaN(0);
	// The qmath routines are used because Floor() and Ceil() are deceptively difficult to implement in a way
	// that gives the correct result in all permutations of the following:
	// 1) Negative vs. positive input.
	// 2) Whether or not the input is already an integer.
	// Therefore, do not change this without conducting a thorough test.
	double x = ParamIndexToDouble(0);
	x = (_f_callee_id == FID_Floor) ? qmathFloor(x) : qmathCeil(x);
	// Fix for v1.0.40.05: For some inputs, qmathCeil/Floor yield a number slightly to the left of the target
	// integer, while for others they yield one slightly to the right.  For example, Ceil(62/61) and Floor(-4/3)
	// yield a double that would give an incorrect answer if it were simply truncated to an integer via
	// type casting.  The below seems to fix this without breaking the answers for other inputs (which is
	// surprisingly harder than it seemed).  There is a similar fix in BIF_Round().
	_f_return_i((__int64)(x + (x > 0 ? 0.2 : -0.2)));
}



BIF_DECL(BIF_Integer)
{
	++aParam; // Skip `this`
	Throw_if_Param_NaN(0);
	_f_return_i(ParamIndexToInt64(0));
}



BIF_DECL(BIF_Float)
{
	++aParam; // Skip `this`
	Throw_if_Param_NaN(0);
	_f_return(ParamIndexToDouble(0));
}



BIF_DECL(BIF_Number)
{
	++aParam; // Skip `this`
	if (!ParamIndexToNumber(0, aResultToken))
		_f_throw_param(0, _T("Number"));
}



BIF_DECL(BIF_Mod)
{
	// Load-time validation has already ensured there are exactly two parameters.
	// "Cast" each operand to Int64/Double depending on whether it has a decimal point.
	ExprTokenType param0, param1;
	if (ParamIndexToNumber(0, param0) && ParamIndexToNumber(1, param1)) // Both are numeric.
	{
		if (param0.symbol == SYM_INTEGER && param1.symbol == SYM_INTEGER) // Both are integers.
		{
			if (param1.value_int64 == 0)
				_f_throw(ERR_DIVIDEBYZERO, ErrorPrototype::ZeroDivision);
			// For performance, % is used vs. qmath for integers.
			_f_return_i(param0.value_int64 % param1.value_int64);
		}
		else // At least one is a floating point number.
		{
			double dividend = TokenToDouble(param0);
			double divisor = TokenToDouble(param1);
			if (divisor == 0.0)
				_f_throw(ERR_DIVIDEBYZERO, ErrorPrototype::ZeroDivision);
			_f_return(qmathFmod(dividend, divisor));
		}
	}
	// Since above didn't return, one or both parameters were invalid.
	_f_throw(ERR_PARAM_INVALID, ErrorPrototype::Type);
}



BIF_DECL(BIF_MinMax)
{
	// Supports one or more parameters.
	// Load-time validation has already ensured there is at least one parameter.
	ExprTokenType param;
	int index, ib_index = 0, db_index = 0;
	bool isMin = _f_callee_id == FID_Min;
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
				_f_throw_param(i, _T("Number"));
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
	if (!TokenToDoubleOrInt64(*aParam[0], aResultToken)) // "Cast" token to Int64/Double depending on whether it has a decimal point.
		_f_throw_param(0, _T("Number")); // Non-operand or non-numeric string.
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
	Throw_if_Param_NaN(0);
	_f_return(qmathSin(ParamIndexToDouble(0)));
}



BIF_DECL(BIF_Cos)
// For simplicity and backward compatibility, a numeric result is always returned (even if the input
// is non-numeric or an empty string).
{
	Throw_if_Param_NaN(0);
	_f_return(qmathCos(ParamIndexToDouble(0)));
}



BIF_DECL(BIF_Tan)
// For simplicity and backward compatibility, a numeric result is always returned (even if the input
// is non-numeric or an empty string).
{
	Throw_if_Param_NaN(0);
	_f_return(qmathTan(ParamIndexToDouble(0)));
}



BIF_DECL(BIF_ASinACos)
{
	Throw_if_Param_NaN(0);
	double value = ParamIndexToDouble(0);
	if (value > 1 || value < -1) // ASin and ACos aren't defined for such values.
	{
		_f_throw_param(0);
	}
	else
	{
		// For simplicity and backward compatibility, a numeric result is always returned in this case (even if
		// the input is non-numeric or an empty string).
		_f_return((_f_callee_id == FID_ASin) ? qmathAsin(value) : qmathAcos(value));
	}
}



BIF_DECL(BIF_ATan)
// For simplicity and backward compatibility, a numeric result is always returned (even if the input
// is non-numeric or an empty string).
{
	Throw_if_Param_NaN(0);
	_f_return(qmathAtan(ParamIndexToDouble(0)));
}



BIF_DECL(BIF_Exp)
{
	Throw_if_Param_NaN(0);
	_f_return(qmathExp(ParamIndexToDouble(0)));
}



BIF_DECL(BIF_SqrtLogLn)
{
	Throw_if_Param_NaN(0);
	double value = ParamIndexToDouble(0);
	if (value < 0) // Result is undefined in these cases.
	{
		_f_throw_param(0);
	}
	else
	{
		// For simplicity and backward compatibility, a numeric result is always returned in this case (even if
		// the input is non-numeric or an empty string).
		switch (_f_callee_id)
		{
		case FID_Sqrt:	_f_return(qmathSqrt(value));
		case FID_Log:	_f_return(qmathLog10(value));
		//case FID_Ln:
		default:		_f_return(qmathLog(value));
		}
	}
}



BIF_DECL(BIF_Random)
{
	UINT64 rand = 0;
	if (!GenRandom(&rand, sizeof(rand)))
		_f_throw(ERR_INTERNAL_CALL);

	SymbolType arg1type = ParamIndexIsOmitted(0) ? SYM_MISSING : ParamIndexIsNumeric(0);
	SymbolType arg2type = ParamIndexIsOmitted(1) ? SYM_MISSING : ParamIndexIsNumeric(1);
	if (arg1type == PURE_NOT_NUMERIC)
		_f_throw_param(0, _T("Number"));
	if (arg2type == PURE_NOT_NUMERIC)
		_f_throw_param(1, _T("Number"));

	bool use_float = arg1type == PURE_FLOAT || arg2type == PURE_FLOAT || !aParamCount; // Let Random() be Random(0.0, 1.0).
	if (use_float)
	{
		double target_min = arg1type != SYM_MISSING ? ParamIndexToDouble(0) : 0.0;
		double target_max = arg2type != SYM_MISSING ? ParamIndexToDouble(1) : arg1type == SYM_MISSING ? 1.0 : 0.0;
		// Be permissive about the order of parameters, and convert Random(n) to Random(0, n).
		if (target_min > target_max)
			swap(target_min, target_max);
		// The first part below produces a 53-bit integer, and from that a value between
		// 0.0 (inclusive) and 1.0 (exclusive) with the maximum precision for a double.
		_f_return((((rand >> 11) / 9007199254740992.0) * (target_max - target_min)) + target_min);
	}
	else
	{
		INT64 target_min = arg1type != SYM_MISSING ? ParamIndexToInt64(0) : 0;
		INT64 target_max = arg2type != SYM_MISSING ? ParamIndexToInt64(1) : 0;
		// Be permissive about the order of parameters, and convert Random(n) to Random(0, n).
		if (target_min > target_max)
			swap(target_min, target_max);
		// Do NOT use floating-point to generate random integers because of cases like
		// min=0 and max=1: we want an even distribution of 1's and 0's in that case, not
		// something skewed that might result due to rounding/truncation issues caused by
		// the float method used above.
		// Furthermore, the simple modulo approach is biased when the target range does not
		// divide cleanly into rand.  Suppose that rand ranges from 0..7, and the target range
		// is 0..2.  By using modulo, rand is effectively divided into sets {0..2, 3..5, 6..7}
		// and each value in the set is mapped to the target range 0..2.  Because the last set
		// maps to 0..1, 0 and 1 have greater chance of appearing than 2.  However, since rand
		// is 64-bit, this isn't actually a problem for small ranges such as 1..100 due to the
		// vanishingly small chance of rand falling within the defective set.
		UINT64 u_max = (UINT64)(target_max - target_min);
		if (u_max < UINT64_MAX)
		{
			// What we actually want is (UINT64_MAX + 1) % (u_max + 1), but without overflow.
			UINT64 error_margin = UINT64_MAX % (u_max + 1);
			if (error_margin != u_max) // i.e. ((error_margin + 1) % (u_max + 1)) != 0.
			{
				++error_margin;
				// error_margin is now the remainder after dividing (UINT64_MAX+1) by (u_max+1).
				// This is also the size of the incomplete number set which must be excluded to
				// ensure even distribution of random numbers.  For simplicity we just take the
				// number set starting at 0, since use of modulo below will auto-correct.
				// For example, with a target range of 1..100, the error_margin should be 16,
				// which gives a mere 16 in 2**64 chance that a second iteration will be needed.
				while (rand < error_margin)
					// To ensure even distribution, keep trying until outside the error margin.
					GenRandom(&rand, sizeof(rand));
			}
			rand %= (u_max + 1);
		}
		_f_return((INT64)(rand + (UINT64)target_min));
	}
}



BIF_DECL(BIF_DateAdd)
{
	FILETIME ft;
	if (!YYYYMMDDToFileTime(ParamIndexToString(0, _f_number_buf), ft))
		_f_throw_value(ERR_PARAM1_INVALID);

	// Use double to support a floating point value for days, hours, minutes, etc:
	double nUnits;
	nUnits = TokenToDouble(*aParam[1]);

	// Convert to 10ths of a microsecond (the units of the FILETIME struct):
	switch (ctoupper(*TokenToString(*aParam[2])))
	{
	case 'S': // Seconds
		nUnits *= (double)10000000;
		break;
	case 'M': // Minutes
		nUnits *= ((double)10000000 * 60);
		break;
	case 'H': // Hours
		nUnits *= ((double)10000000 * 60 * 60);
		break;
	case 'D': // Days
		nUnits *= ((double)10000000 * 60 * 60 * 24);
		break;
	default: // Invalid
		_f_throw_param(2);
	}
	// Convert ft struct to a 64-bit variable (maybe there's some way to avoid these conversions):
	ULARGE_INTEGER ul;
	ul.LowPart = ft.dwLowDateTime;
	ul.HighPart = ft.dwHighDateTime;
	// Add the specified amount of time to the result value:
	ul.QuadPart += (__int64)nUnits;  // Seems ok to cast/truncate in light of the *=10000000 above.
	// Convert back into ft struct:
	ft.dwLowDateTime = ul.LowPart;
	ft.dwHighDateTime = ul.HighPart;
	_f_return_p(FileTimeToYYYYMMDD(_f_retval_buf, ft, false));
}



BIF_DECL(BIF_DateDiff)
{
	LPTSTR error_message;
	// If either parameter is blank, it will default to the current time:
	__int64 time_until; // Declaring separate from initializing avoids compiler warning when not inside a block.
	TCHAR number_buf[MAX_NUMBER_SIZE]; // Additional buf used in case both parameters are pure numbers.
	time_until = YYYYMMDDSecondsUntil(
		ParamIndexToString(1, _f_number_buf),
		ParamIndexToString(0, number_buf),
		error_message);
	if (error_message) // Usually caused by an invalid component in the date-time string.
		_f_throw(error_message);
	switch (ctoupper(*ParamIndexToString(2)))
	{
	case 'S': break;
	case 'M': time_until /= 60; break; // Minutes
	case 'H': time_until /= 60 * 60; break; // Hours
	case 'D': time_until /= 60 * 60 * 24; break; // Days
	default: // Invalid
		_f_throw_param(2);
	}
	_f_return_i(time_until);
}



BIF_DECL(BIF_Hotkey)
{
	_f_param_string_opt(aParam0, 0);
	_f_param_string_opt(aParam1, 1);
	_f_param_string_opt(aParam2, 2);
	
	ResultType result = OK;
	IObject *functor = nullptr;

	switch (_f_callee_id) 
	{
	case FID_Hotkey:
	{
		HookActionType hook_action = 0;
		if (!ParamIndexIsOmitted(1))
		{
			if (  !(functor = ParamIndexToObject(1)) && *aParam1
				&& !(hook_action = Hotkey::ConvertAltTab(aParam1, true))  )
			{
				// Search for a match in the hotkey variants' "original callbacks".
				// I.e., find the function implicitly defined by "x::action".
				for (int i = 0; i < Hotkey::sHotkeyCount; ++i)
				{
					if (_tcscmp(Hotkey::shk[i]->mName, aParam1))
						continue;
					
					for (HotkeyVariant* v = Hotkey::shk[i]->mFirstVariant; v; v = v->mNextVariant)
						if (v->mHotCriterion == g->HotCriterion)
						{
							functor = v->mOriginalCallback.ToFunc();
							goto break_twice;
						}
				}
			break_twice:;
				if (!functor)
					_f_throw_param(1);
			}
			if (!functor)
				hook_action = Hotkey::ConvertAltTab(aParam1, true);
		}
		result = Hotkey::Dynamic(aParam0, aParam2, functor, hook_action, aResultToken);
		break;
	}
	case FID_HotIf:
		functor = ParamIndexToOptionalObject(0);
		result = Hotkey::IfExpr(aParam0, functor, aResultToken);
		break;
	
	default: // HotIfWinXXX
		result = SetHotkeyCriterion(_f_callee_id, aParam0, aParam1); // Currently, it only fails upon out-of-memory.
	
	}
	
	if (!result)
		_f_return_FAIL;
	_f_return_empty;
}



BIF_DECL(BIF_SetTimer)
{
	IObject *callback;
	// Note that only one timer per callback is allowed because the callback is the
	// unique identifier that allows us to update or delete an existing timer.
	if (ParamIndexIsOmitted(0)) // Fully omitted, not an empty string.
	{
		if (g->CurrentTimer)
			// Default to the timer which launched the current thread.
			callback = g->CurrentTimer->mCallback.ToObject();
		else
			callback = NULL;
		if (!callback)
			// Either the thread was not launched by a timer or the timer has been deleted.
			_f_throw_value(ERR_PARAM1_MUST_NOT_BE_BLANK);
	}
	else
	{
		callback = ParamIndexToObject(0);
		if (!callback)
			_f_throw_param(0, _T("object"));
		if (!ValidateFunctor(callback, 0, aResultToken))
			return;
	}
	__int64 period = DEFAULT_TIMER_PERIOD;
	int priority = 0;
	bool update_period = false, update_priority = false;
	if (!ParamIndexIsOmitted(1))
	{
		Throw_if_Param_NaN(1);
		period = ParamIndexToInt64(1);
		if (!period)
		{
			g_script.DeleteTimer(callback);
			_f_return_empty;
		}
		update_period = true;
	}
	if (!ParamIndexIsOmitted(2))
	{
		priority = ParamIndexToInt(2);
		update_priority = true;
	}
	g_script.UpdateOrCreateTimer(callback, update_period, period, update_priority, priority);
	_f_return_empty;
}



BIF_DECL(BIF_OnMessage)
// Returns: An empty string.
// Parameters:
// 1: Message number to monitor.
// 2: Name of the function that will monitor the message.
// 3: Maximum threads and "register first" flag.
{
	// Currently OnMessage (in v2) has no return value.
	_f_set_retval_p(_T(""), 0);

	// Prior validation has ensured there's at least two parameters:
	UINT specified_msg = (UINT)ParamIndexToInt64(0); // Parameter #1

	// Set defaults:
	bool mode_is_delete = false;
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

	// Parameter #2: The callback to add or remove.  Must be an object.
	IObject *callback = TokenToObject(*aParam[1]);
	if (!callback)
		_f_throw_param(1, _T("object"));

	// Check if this message already exists in the array:
	MsgMonitorStruct *pmonitor = g_MsgMonitor.Find(specified_msg, callback);
	bool item_already_exists = (pmonitor != NULL);
	if (!item_already_exists)
	{
		if (mode_is_delete) // Delete a non-existent item.
			_f_return_retval; // Yield the default return value set earlier (an empty string).
		if (!ValidateFunctor(callback, 4, aResultToken))
			return;
		// From this point on, it is certain that an item will be added to the array.
		pmonitor = g_MsgMonitor.Add(specified_msg, callback, call_it_last);
		if (!pmonitor)
			_f_throw_oom;
	}

	MsgMonitorStruct &monitor = *pmonitor;

	if (item_already_exists)
	{
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
			g_MsgMonitor.Delete(pmonitor);
			_f_return_retval;
		}
		if (aParamCount < 2) // Single-parameter mode: Report existing item's function name.
			_f_return_retval; // Everything was already set up above to yield the proper return value.
		// Otherwise, an existing item is being assigned a new function or MaxThreads limit.
		// Continue on to update this item's attributes.
	}
	else // This message was newly added to the array.
	{
		// The above already verified that callback is not NULL and there is room in the array.
		monitor.instance_count = 0; // Reset instance_count only for new items since existing items might currently be running.
		// Continue on to the update-or-create logic below.
	}

	// Update those struct attributes that get the same treatment regardless of whether this is an update or creation.
	if (!item_already_exists || !ParamIndexIsOmitted(2))
		monitor.max_instances = max_instances;
	// Otherwise, the parameter was omitted so leave max_instances at its current value.
	_f_return_retval;
}


MsgMonitorStruct *MsgMonitorList::Find(UINT aMsg, IObject *aCallback, UCHAR aMsgType)
{
	for (int i = 0; i < mCount; ++i)
		if (mMonitor[i].msg == aMsg
			&& mMonitor[i].func == aCallback // No need to check is_method, since it's impossible for an object and string to exist at the same address.
			&& mMonitor[i].msg_type == aMsgType) // Checked last because it's nearly always true.
			return mMonitor + i;
	return NULL;
}


MsgMonitorStruct *MsgMonitorList::Find(UINT aMsg, LPTSTR aMethodName, UCHAR aMsgType)
{
	for (int i = 0; i < mCount; ++i)
		if (mMonitor[i].msg == aMsg
			&& mMonitor[i].is_method && !_tcsicmp(aMethodName, mMonitor[i].method_name)
			&& mMonitor[i].msg_type == aMsgType) // Checked last because it's nearly always true.
			return mMonitor + i;
	return NULL;
}


MsgMonitorStruct *MsgMonitorList::AddInternal(UINT aMsg, bool aAppend)
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
	new_mon->msg = aMsg;
	new_mon->msg_type = 0; // Must be initialised to 0 for all callers except GUI.
	// These are initialised by OnMessage, since OnExit and OnClipboardChange don't use them:
	//new_mon->instance_count = 0;
	//new_mon->max_instances = 1;
	return new_mon;
}


MsgMonitorStruct *MsgMonitorList::Add(UINT aMsg, IObject *aCallback, bool aAppend)
{
	MsgMonitorStruct *new_mon = AddInternal(aMsg, aAppend);
	if (new_mon)
	{
		aCallback->AddRef();
		new_mon->func = aCallback;
		new_mon->is_method = false;
	}
	return new_mon;
}


MsgMonitorStruct *MsgMonitorList::Add(UINT aMsg, LPTSTR aMethodName, bool aAppend)
{
	if (  !(aMethodName = _tcsdup(aMethodName))  )
		return NULL;
	MsgMonitorStruct *new_mon = AddInternal(aMsg, aAppend);
	if (new_mon)
	{
		new_mon->method_name = aMethodName;
		new_mon->is_method = true;
	}
	else
		free(aMethodName);
	return new_mon;
}


void MsgMonitorList::Delete(MsgMonitorStruct *aMonitor)
{
	ASSERT(aMonitor >= mMonitor && aMonitor < mMonitor + mCount);

	int mon_index = int(aMonitor - mMonitor);
	// Adjust the index of any active message monitors affected by this deletion.  This allows a
	// message monitor to delete older message monitors while still allowing any remaining monitors
	// of that message to be called (when there are multiple).
	for (MsgMonitorInstance *inst = mTop; inst; inst = inst->previous)
	{
		inst->Delete(mon_index);
	}
	// Remove the item from the array.
	--mCount;  // Must be done prior to the below.
	LPVOID release_me = aMonitor->union_value;
	bool is_method = aMonitor->is_method;
	if (mon_index < mCount) // An element other than the last is being removed. Shift the array to cover/delete it.
		memmove(aMonitor, aMonitor + 1, (mCount - mon_index) * sizeof(MsgMonitorStruct));
	if (is_method)
		free(release_me);
	else
		reinterpret_cast<IObject *>(release_me)->Release(); // Must be called last in case it calls a __delete() meta-function.
}


BOOL MsgMonitorList::IsMonitoring(UINT aMsg, UCHAR aMsgType)
{
	for (int i = 0; i < mCount; ++i)
		if (mMonitor[i].msg == aMsg && mMonitor[i].msg_type == aMsgType)
			return TRUE;
	return FALSE;
}


BOOL MsgMonitorList::IsRunning(UINT aMsg, UCHAR aMsgType)
// Returns true if there are any monitors for a message currently executing.
{
	for (MsgMonitorInstance *inst = mTop; inst; inst = inst->previous)
		if (!inst->deleted && mMonitor[inst->index].msg == aMsg && mMonitor[inst->index].msg_type == aMsgType)
			return TRUE;
	//if (!mTop)
	//	return FALSE;
	//for (int i = 0; i < mCount; ++i)
	//	if (mMonitor[i].msg == aMsg && mMonitor[i].instance_count)
	//		return TRUE;
	return FALSE;
}


void MsgMonitorList::Dispose()
{
	// Although other action taken by GuiType::Destroy() ensures the event list isn't
	// reachable from script once destruction begins, we take the careful approach and
	// decrement mCount at each iteration to ensure that if Release() executes external
	// code, this list is always in a valid state.
	while (mCount)
	{
		--mCount;
		if (mMonitor[mCount].is_method)
			free(mMonitor[mCount].method_name);
		else
			mMonitor[mCount].func->Release();
	}
	free(mMonitor);
	mMonitor = nullptr;
	mCountMax = 0;
	// Dispose all iterator instances to ensure Call() does not continue iterating:
	for (MsgMonitorInstance *inst = mTop; inst; inst = inst->previous)
		inst->Dispose();
}


BIF_DECL(BIF_On)
{
	_f_set_retval_p(_T("")); // In all cases there is no return value.
	auto event_type = _f_callee_id;
	MsgMonitorList *phandlers;
	switch (event_type)
	{
	case FID_OnError: phandlers = &g_script.mOnError; break;
	case FID_OnClipboardChange: phandlers = &g_script.mOnClipboardChange; break;
	default: phandlers = &g_script.mOnExit; break;
	}
	MsgMonitorList &handlers = *phandlers;


	IObject *callback = ParamIndexToObject(0);
	if (!callback)
		_f_throw_param(0, _T("object"));
	if (!ValidateFunctor(callback, event_type == FID_OnClipboardChange ? 1 : 2, aResultToken))
		return;
	
	int mode = 1; // Default.
	if (!ParamIndexIsOmitted(1))
		mode = ParamIndexToInt(1);

	MsgMonitorStruct *existing = handlers.Find(0, callback);

	switch (mode)
	{
	case  1:
	case -1:
		if (existing)
			return;
		if (event_type == FID_OnClipboardChange)
		{
			// Do this before adding the handler so that it won't be called as a result of the
			// SetClipboardViewer() call on Windows XP.  This won't cause existing handlers to
			// be called because in that case the clipboard listener is already enabled.
			g_script.EnableClipboardListener(true);
		}
		if (!handlers.Add(0, callback, mode == 1))
			_f_throw_oom;
		break;
	case  0:
		if (existing)
			handlers.Delete(existing);
		break;
	default:
		_f_throw_param(1);
	}
	// In case the above enabled the clipboard listener but failed to add the handler,
	// do this even if mode != 0:
	if (event_type == FID_OnClipboardChange && !handlers.Count())
		g_script.EnableClipboardListener(false);
}


#ifdef ENABLE_REGISTERCALLBACK
struct RCCallbackFunc // Used by BIF_CallbackCreate() and related.
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
#define CBF_CREATE_NEW_THREAD	1
#define CBF_PASS_PARAMS_POINTER	2
	UCHAR flags; // Kept adjacent to above to conserve memory due to 4-byte struct alignment in 32-bit builds.
	IObject *func; // The function object to be called whenever the callback's caller calls callfuncptr.
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
#ifdef WIN32_PLATFORM
	RCCallbackFunc &cb = *((RCCallbackFunc*)(address-5)); //second instruction is 5 bytes after start (return address pushed by call)
#else
	RCCallbackFunc &cb = *((RCCallbackFunc*) address);
#endif

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
	if (cb.flags & CBF_CREATE_NEW_THREAD)
	{
		if (g_nThreads >= g_MaxThreadsTotal) // To avoid array overflow, g_MaxThreadsTotal must not be exceeded except where otherwise documented.
			return 0;
		// See MsgSleep() for comments about the following section.
		InitNewThread(0, false, true);
		DEBUGGER_STACK_PUSH(_T("Callback"))
	}
	else
	{
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

	g_script.mLastPeekTime = GetTickCount(); // Somewhat debatable, but might help minimize interruptions when the callback is called via message (e.g. subclassing a control; overriding a WindowProc).

	__int64 number_to_return;
	FuncResult result_token;
	ExprTokenType *param, one_param;
	int param_count;

	if (cb.flags & CBF_PASS_PARAMS_POINTER)
	{
		param_count = 1;
		param = &one_param;
		one_param.SetValue((UINT_PTR)params);
	}
	else
	{
		param_count = cb.actual_param_count;
		param = (ExprTokenType *)_alloca(param_count * sizeof(ExprTokenType));
		for (int i = 0; i < param_count; ++i)
			param[i].SetValue((UINT_PTR)params[i]);
	}
	
	CallMethod(cb.func, cb.func, nullptr, param, param_count, &number_to_return);
	// CallMethod()'s own return value is ignored because it wouldn't affect the handling below.
	
	if (cb.flags & CBF_CREATE_NEW_THREAD)
	{
		DEBUGGER_STACK_POP()
		ResumeUnderlyingThread();
	}
	else
	{
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

	return (INT_PTR)number_to_return;
}



BIF_DECL(BIF_CallbackCreate)
// Returns: Address of callback procedure, or empty string on failure.
// Parameters:
// 1: Name of the function to be called when the callback routine is executed.
// 2: Options.
// 3: Number of parameters of callback.
//
// Author: Original x86 RegisterCallback() was created by Jonathan Rennison (JGR).
//   x64 support by fincs.  Various changes by Lexikos.
{
	IObject *func = ParamIndexToObject(0);
	if (!func)
		_f_throw_param(0, _T("object"));

	LPTSTR options = ParamIndexToOptionalString(1);
	bool pass_params_pointer = _tcschr(options, '&'); // Callback wants the address of the parameter list instead of their values.
#ifdef WIN32_PLATFORM
	bool use_cdecl = StrChrAny(options, _T("Cc")); // Recognize "C" as the "CDecl" option.
	bool require_param_count = !use_cdecl; // Param count must be specified for x86 stdcall.
#else
	bool require_param_count = false;
#endif

	bool params_specified = !ParamIndexIsOmittedOrEmpty(2);
	if (pass_params_pointer && require_param_count && !params_specified)
		_f_throw_value(ERR_PARAM3_MUST_NOT_BE_BLANK);

	int actual_param_count = params_specified ? ParamIndexToInt(2) : 0;
	if (!ValidateFunctor(func
		, pass_params_pointer ? 1 : actual_param_count // Count of script parameters being passed.
		, aResultToken
		// Use MinParams as actual_param_count if unspecified and no & option.
		, params_specified || pass_params_pointer ? nullptr : &actual_param_count))
	{
		return;
	}
	
#ifdef WIN32_PLATFORM
	if (!use_cdecl && actual_param_count > 31) // The ASM instruction currently used limits parameters to 31 (which should be plenty for any realistic use).
	{
		func->Release();
		_f_throw_param(2);
	}
#endif

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
	if (!callbackfunc)
		_f_throw_oom;
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

	func->AddRef();
	cb.func = func;
	cb.actual_param_count = actual_param_count;
	cb.flags = 0;
	if (!StrChrAny(options, _T("Ff"))) // Recognize "F" as the "fast" mode that avoids creating a new thread.
		cb.flags |= CBF_CREATE_NEW_THREAD;
	if (pass_params_pointer)
		cb.flags |= CBF_PASS_PARAMS_POINTER;

	// If DEP is enabled (and sometimes when DEP is apparently "disabled"), we must change the
	// protection of the page of memory in which the callback resides to allow it to execute:
	DWORD dwOldProtect;
	VirtualProtect(callbackfunc, sizeof(RCCallbackFunc), PAGE_EXECUTE_READWRITE, &dwOldProtect);

	_f_return_i((__int64)callbackfunc); // Yield the callable address as the result.
}

BIF_DECL(BIF_CallbackFree)
{
	INT_PTR address = ParamIndexToIntPtr(0);
	if (address < 65536 && address >= 0) // Basic sanity check to catch incoming raw addresses that are zero or blank.  On Win32, the first 64KB of address space is always invalid.
		_f_throw_param(0);
	RCCallbackFunc *callbackfunc = (RCCallbackFunc *)address;
	callbackfunc->func->Release();
	callbackfunc->func = NULL; // To help detect bugs.
	GlobalFree(callbackfunc);
	_f_return_empty;
}

#endif



BIF_DECL(BIF_MenuFromHandle)
{
	auto *menu = g_script.FindMenu((HMENU)ParamIndexToInt64(0));
	if (menu)
	{
		menu->AddRef();
		_f_return(menu);
	}
	_f_return_empty;
}



void GuiControlType::StatusBar(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	GuiType& gui = *this->gui;
	HWND control_hwnd = this->hwnd;
	LPTSTR buf = _f_number_buf;

	HICON hicon;
	switch (aID)
	{
	case FID_SB_SetText:
		_o_return(SendMessage(control_hwnd, SB_SETTEXT
			, (WPARAM)((ParamIndexIsOmitted(1) ? 0 : ParamIndexToInt64(1) - 1) // The Part# param is present.
				     | (ParamIndexIsOmitted(2) ? 0 : ParamIndexToInt64(2) << 8)) // The uType parameter is present.
			, (LPARAM)ParamIndexToString(0, buf))); // Caller has ensured that there's at least one param in this mode.

	case FID_SB_SetParts:
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

		_o_return(SendMessage(control_hwnd, SB_SETPARTS, new_part_count, (LPARAM)part)
			? (__int64)control_hwnd : 0); // Return HWND to provide an easy means for the script to get the bar's HWND.

	//case FID_SB_SetIcon:
	default:
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
				if (hicon_old)
					// Although the old icon is automatically destroyed here, the script can call SendMessage(SB_SETICON)
					// itself if it wants to work with HICONs directly (for performance reasons, etc.)
					DestroyIcon(hicon_old);
			}
			else
			{
				DestroyIcon(hicon);
				hicon = NULL;
			}
		}
		//else can't load icon, so return 0.
		_o_return((size_t)hicon); // This allows the script to destroy the HICON later when it doesn't need it (see comments above too).
	// SB_SetTipText() not implemented (though can be done via SendMessage in the script) because the conditions
	// under which tooltips are displayed don't seem like something a script would want very often:
	// This ToolTip text is displayed in two situations: 
	// When the corresponding pane in the status bar contains only an icon. 
	// When the corresponding pane in the status bar contains text that is truncated due to the size of the pane.
	// In spite of the above, SB_SETTIPTEXT doesn't actually seem to do anything, even when the text is too long
	// to fit in a narrowed part, tooltip text has been set, and the user hovers the cursor over the bar.  Maybe
	// I'm not doing it right or maybe this feature is somehow disabled under certain service packs or conditions.
	//case 'T': // SB.SetTipText()
	//	break;
	} // switch(mode)
}


void GuiControlType::LV_GetNextOrCount(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// LV.GetNext:
// Returns: The index of the found item, or 0 on failure.
// Parameters:
// 1: Starting index (one-based when it comes in).  If absent, search starts at the top.
// 2: Options string.
// 3: (FUTURE): Possible for use with LV.FindItem (though I think it can only search item text, not subitem text).
{
	HWND control_hwnd = hwnd;

	LPTSTR options;
	if (aID == FID_LV_GetCount)
	{
		options = (aParamCount > 0) ? omit_leading_whitespace(ParamIndexToString(0, _f_number_buf)) : _T("");
		if (*options)
		{
			if (ctoupper(*options) == 'S')
				_o_return(SendMessage(control_hwnd, LVM_GETSELECTEDCOUNT, 0, 0));
			else if (!_tcsnicmp(options, _T("Col"), 3)) // "Col" or "Column". Don't allow "C" by itself, so that "Checked" can be added in the future.
				_o_return(union_lv_attrib->col_count);
			else
				_o_throw_param(0);
		}
		_o_return(SendMessage(control_hwnd, LVM_GETITEMCOUNT, 0, 0));
	}
	// Since above didn't return, this is GetNext() mode.

	int index = -1;
	if (!ParamIndexIsOmitted(0))
	{
		if (!ParamIndexIsNumeric(0))
			_o_throw_param(0, _T("Number"));
		index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
		// For flexibility, allow index to be less than -1 to avoid first-iteration complications in script loops
		// (such as when deleting rows, which shifts the row index upward, require the search to resume at
		// the previously found index rather than the row after it).  However, reset it to -1 to ensure
		// proper return values from the API in the "find checked item" mode used below.
		if (index < -1)
			index = -1;  // Signal it to start at the top.
	}

	// For performance, decided to always find next selected item when the "C" option hasn't been specified,
	// even when the checkboxes style is in effect.  Otherwise, would have to fetch and check checkbox style
	// bit for each call, which would slow down this heavily-called function.

	options = ParamIndexToOptionalString(1, _f_number_buf);
	TCHAR first_char = ctoupper(*omit_leading_whitespace(options));
	// To retain compatibility in the future, also allow "Check(ed)" and "Focus(ed)" since any word that
	// starts with C or F is already supported.

	switch(first_char)
	{
	case '\0': // Listed first for performance.
	case 'F':
		_o_return(ListView_GetNextItem(control_hwnd, index
			, first_char ? LVNI_FOCUSED : LVNI_SELECTED) + 1); // +1 to convert to 1-based.
	case 'C': // Checkbox: Find checked items. For performance assume that the control really has checkboxes.
	  {
		int item_count = ListView_GetItemCount(control_hwnd);
		for (int i = index + 1; i < item_count; ++i) // Start at index+1 to omit the first item from the search (for consistency with the other mode above).
			if (ListView_GetCheckState(control_hwnd, i)) // Item's box is checked.
				_o_return(i + 1); // +1 to convert from zero-based to one-based.
		// Since above didn't return, no match found.
		_o_return(0);
	  }
	default:
		_o_throw_param(1);
	}
}



void GuiControlType::LV_GetText(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns: Text on success.
// Throws on failure.
// Parameters:
// 1: Row index (one-based when it comes in).
// 2: Column index (one-based when it comes in).
{
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Item not found in ListView.
	// And others.

	int row_index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
	if (row_index < -1) // row_index==-1 is reserved to mean "get column heading's text".
		_o_throw_param(0);
	// If parameter 2 is omitted, default to the first column (index 0):
	int col_index = ParamIndexIsOmitted(1) ? 0 : ParamIndexToInt(1) - 1; // -1 to convert to zero-based.
	if (col_index < 0)
		_o_throw_param(1);

	TCHAR buf[LV_TEXT_BUF_SIZE];

	if (row_index == -1) // Special mode to get column's text.
	{
		LVCOLUMN lvc;
		lvc.cchTextMax = LV_TEXT_BUF_SIZE - 1;  // See notes below about -1.
		lvc.pszText = buf;
		lvc.mask = LVCF_TEXT;
		if (SendMessage(hwnd, LVM_GETCOLUMN, col_index, (LPARAM)&lvc)) // Assign.
			_o_return(lvc.pszText); // See notes below about why pszText is used instead of buf (might apply to this too).
		else // On failure, it seems best to throw.
			_o_throw(ERR_FAILED);
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
		if (SendMessage(hwnd, LVM_GETITEM, 0, (LPARAM)&lvi)) // Assign
			// Must use lvi.pszText vs. buf because MSDN says: "Applications should not assume that the text will
			// necessarily be placed in the specified buffer. The control may instead change the pszText member
			// of the structure to point to the new text rather than place it in the buffer."
			_o_return(lvi.pszText);
		else // On failure, it seems best to throw.
			_o_throw(ERR_FAILED);
	}
}



void GuiControlType::LV_AddInsertModify(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: For Add(), this is the options.  For Insert/Modify, it's the row index (one-based when it comes in).
// 2: For Add(), this is the first field's text.  For Insert/Modify, it's the options.
// 3 and beyond: Additional field text.
// In Add/Insert mode, if there are no text fields present, a blank for is appended/inserted.
{
	GuiControlType &control = *this;
	auto mode = BuiltInFunctionID(aID);
	LPTSTR buf = _f_number_buf; // Resolve macro early for maintainability.

	int index;
	if (mode == FID_LV_Add) // For Add mode, use INT_MAX as a signal to append the item rather than inserting it.
	{
		index = INT_MAX;
		mode = FID_LV_Insert; // Add has now been set up to be the same as insert, so change the mode to simplify other things.
	}
	else // Insert or Modify: the target row-index is their first parameter, which load-time has ensured is present.
	{
		index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
		if (index < -1 || (mode != FID_LV_Modify && index < 0)) // Allow -1 to mean "all rows" when in modify mode.
			_o_throw_param(0);
		++aParam;  // Remove the first parameter from further consideration to make Insert/Modify symmetric with Add.
		--aParamCount;
	}

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
			if (mode == FID_LV_Modify) // v1.0.46.10: Do this section only for Modify, not Add/Insert, to avoid generating an extra "unchecked" notification when a row is added/inserted with an initial state of "checked".  In other words, the script now receives only a "checked" notification, not an "unchecked+checked". Search on is_checked for more comments.
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
		{
			// Since this option much more typically used with LV.Modify than LV.Add/Insert, the technique of
			// Vis%VarContainingOneOrZero% isn't supported, to reduce code size.
			ensure_visible = adding; // Ignored by modes other than LV.Modify(), since it's not really appropriate when adding a row (plus would add code complexity).
		}
		else
		{
			aResultToken.ValueError(ERR_INVALID_OPTION, next_option);
			*option_end = orig_char; // See comment below.
			return;
		}

		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
	}

	// Suppress any events raised by the changes made below:
	control.attrib |= GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS;

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
	int result = 1; // Set default from this point forward to be true/success. It will be overridden in insert mode to be the index of the new row.

	for (j = 0; j < rows_to_change; ++j, ++lvi.iItem) // ++lvi.iItem because if the loop has more than one iteration, by definition it is modifying all rows starting at 0.
	{
		if (!ParamIndexIsOmitted(1) && col_start_index == 0) // 2nd parameter: item's text (first field) is present, so include that when setting the item.
		{
			lvi.pszText = ParamIndexToString(1, buf); // Fairly low-overhead, so called every iteration for simplicity (so that buf can be used for both items and subitems).
			lvi.mask |= LVIF_TEXT;
		}
		if (mode == FID_LV_Insert) // Insert or Add.
		{
			// Note that ListView_InsertItem() will append vs. insert if the index is too large, in which case
			// it returns the items new index (which will be the last item in the list unless the control has
			// auto-sort style).
			if (   -1 == (lvi_sub.iItem = ListView_InsertItem(control.hwnd, &lvi))   )
			{
				control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
				_o_return(0); // Since item can't be inserted, no reason to try attaching any subitems to it.
			}
			// Update iItem with the actual index assigned to the item, which might be different than the
			// specified index if the control has an auto-sort style in effect.  This new iItem value
			// is used for ListView_SetCheckState() and for the attaching of any subitems to this item.
			result = lvi_sub.iItem + 1; // Convert from zero-based to 1-based.
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
				result = 0; // Indicate partial failure, but attempt to continue in case it failed for reason other than "row doesn't exist".
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
			if (aParam[i]->symbol == SYM_MISSING) // Omitted, such as LV.Modify(1,Opt,"One",,"Three").
				continue;
			lvi_sub.pszText = ParamIndexToString(i, buf); // Done every time through the outer loop since it's not high-overhead, and for code simplicity.
			if (!ListView_SetItem(control.hwnd, &lvi_sub) && mode != FID_LV_Insert) // Relies on short-circuit. Seems best to avoid loss of item's index in insert mode, since failure here should be rare.
				result = 0; // Indicate partial failure, but attempt to continue in case it failed for reason other than "row doesn't exist".
		}
	} // outer for()

	// When the control has no rows, work around the fact that LVM_SETITEMCOUNT delivers less than 20%
	// of its full benefit unless done after the first row is added (at least on XP SP1).  A non-zero
	// row_count_hint tells us that this message should be sent after the row has been inserted/appended:
	if (control.union_lv_attrib->row_count_hint > 0 && mode == FID_LV_Insert)
	{
		SendMessage(control.hwnd, LVM_SETITEMCOUNT, control.union_lv_attrib->row_count_hint, 0); // Last parameter should be 0 for LVS_OWNERDATA (verified if you look at the definition of ListView_SetItemCount macro).
		control.union_lv_attrib->row_count_hint = 0; // Reset so that it only gets set once per request.
	}

	control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
	_o_return(result);
}



void GuiControlType::LV_Delete(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: Row index (one-based when it comes in).
{
	if (ParamIndexIsOmitted(0))
		_o_return(SendMessage(hwnd, LVM_DELETEALLITEMS, 0, 0)); // Returns TRUE/FALSE.

	// Since above didn't return, there is a first parameter present.
	int index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
	if (index > -1)
		_o_return(SendMessage(hwnd, LVM_DELETEITEM, index, 0)); // Returns TRUE/FALSE.
	else
		// Even if index==0, for safety, it seems best not to do a delete-all.
		_o_throw_param(0);
}



void GuiControlType::LV_InsertModifyDeleteCol(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: Column index (one-based when it comes in).
// 2: String of options
// 3: New text of column
// There are also some special modes when only zero or one parameter is present, see below.
{
	auto mode = BuiltInFunctionID(aID);
	LPTSTR buf = _f_number_buf; // Resolve macro early for maintainability.
	aResultToken.SetValue(0); // Set default return value.

	GuiControlType &control = *this;
	GuiType &gui = *control.gui;
	lv_attrib_type &lv_attrib = *control.union_lv_attrib;
	DWORD view_mode = mode != 'D' ? ListView_GetView(control.hwnd) : 0;

	int index;
	if (!ParamIndexIsOmitted(0))
		index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
	else // Zero parameters.  Load-time validation has ensured that the 'D' (delete) mode cannot have zero params.
	{
		if (mode == FID_LV_ModifyCol)
		{
			if (view_mode != LV_VIEW_DETAILS)
				_o_return_retval; // Return 0 to indicate failure.
			// Otherwise:
			// v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
			// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items):
			for (int i = 0; ; ++i) // Don't limit it to lv_attrib.col_count in case script added extra columns via direct API calls.
				if (!ListView_SetColumnWidth(control.hwnd, i, LVSCW_AUTOSIZE)) // Failure means last column has already been processed.
					break;
			_o_return(1); // Always successful, regardless of what happened in the loop above.
		}
		// Since above didn't return, mode must be 'I' (insert).
		index = lv_attrib.col_count; // When no insertion index was specified, append to the end of the list.
	}

	// Do this prior to checking if index is in bounds so that it can support columns beyond LV_MAX_COLUMNS:
	if (mode == FID_LV_DeleteCol) // Delete a column.  In this mode, index parameter was made mandatory via load-time validation.
	{
		if (ListView_DeleteColumn(control.hwnd, index))  // Returns TRUE/FALSE.
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
			_f_set_retval_i(1);
		}
		_o_return_retval;
	}
	// Do this prior to checking if index is in bounds so that it can support columns beyond LV_MAX_COLUMNS:
	if (mode == FID_LV_ModifyCol && aParamCount < 2) // A single parameter is a special modify-mode to auto-size that column.
	{
		// v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
		// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items):
		if (view_mode == LV_VIEW_DETAILS)
			_f_set_retval_i(ListView_SetColumnWidth(control.hwnd, index, LVSCW_AUTOSIZE));
		//else leave retval set to 0.
		_o_return_retval;
	}
	if (mode == FID_LV_InsertCol)
	{
		if (lv_attrib.col_count >= LV_MAX_COLUMNS) // No room to insert or append.
			_o_return_retval;
		if (index >= lv_attrib.col_count) // For convenience, fall back to "append" when index too large.
			index = lv_attrib.col_count;
	}
	//else do nothing so that modification and deletion of columns that were added via script's
	// direct calls to the API can sort-of work (it's documented in the help file that it's not supported,
	// since col-attrib array can get out of sync with actual columns that way).

	if (index < 0 || index >= LV_MAX_COLUMNS) // For simplicity, do nothing else if index out of bounds.
		_o_return_retval; // Avoid array under/overflow below.

	// In addition to other reasons, must convert any numeric value to a string so that an isolated width is
	// recognized, e.g. LV.SetCol(1, old_width + 10):
	LPTSTR options = ParamIndexToOptionalString(1, buf);

	// It's done the following way so that when in insert-mode, if the column fails to be inserted, don't
	// have to remove the inserted array element from the lv_attrib.col array:
	lv_col_type temp_col = {0}; // Init unconditionally even though only needed for mode=='I'.
	lv_col_type &col = (mode == FID_LV_InsertCol) ? temp_col : lv_attrib.col[index]; // Done only after index has been confirmed in-bounds.

	LVCOLUMN lvc;
	lvc.mask = LVCF_FMT;
	if (mode == FID_LV_ModifyCol) // Fetch the current format so that it's possible to leave parts of it unaltered.
		ListView_GetColumn(control.hwnd, index, &lvc);
	else // Mode is "insert".
		lvc.fmt = 0;

	// Init defaults prior to parsing options:
	bool sort_now = false;
	int do_auto_size = (mode == FID_LV_InsertCol) ? LVSCW_AUTOSIZE_USEHEADER : 0;  // Default to auto-size for new columns.
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

		else if (!_tcsnicmp(next_option, _T("Sort"), 4)) // This is done as an option vs. LV.SortCol/LV.Sort so that the column's options can be changed simultaneously with a "sort now" to refresh.
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
			// options.  For example: LV.SetCol(1, old_width + 10)
			// v1.0.37: Fixed to allow floating point (although ATOI below will convert it to integer).
			if (IsNumeric(next_option, true, false, true)) // Above has already verified that *next_option can't be whitespace.
			{
				lvc.mask |= LVCF_WIDTH;
				int width = gui.Scale(ATOI(next_option));
				// Specifying a width when the column is initially added prevents the scrollbar from
				// updating on Windows 7 and 10 (but not XP).  As a workaround, initialise the width
				// to 0 and then resize it afterward.  do_auto_size is overloaded for this purpose
				// since it's already passed to ListView_SetColumnWidth().
				if (mode == 'I' && view_mode == LV_VIEW_DETAILS)
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
			else
			{
				aResultToken.ValueError(ERR_INVALID_OPTION, next_option);
				*option_end = orig_char; // See comment below.
				return;
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

	if (mode == FID_LV_ModifyCol) // Modify vs. Insert (Delete was already returned from, higher above).
		// For code simplicity, this is called unconditionally even if nothing internal the control's column
		// needs updating.  This seems justified given how rarely columns are modified.
		_f_set_retval_i(ListView_SetColumn(control.hwnd, index, &lvc)); // Returns TRUE/FALSE.
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
		if (   -1 == (index = ListView_InsertColumn(control.hwnd, index, &lvc))   )
			_o_return_retval; // Since column could not be inserted, return so that below, sort-now, etc. are not done.
		_f_set_retval_i(index + 1); // +1 to convert the new index to 1-based.
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
	if (do_auto_size && view_mode == LV_VIEW_DETAILS)
		ListView_SetColumnWidth(control.hwnd, index, do_auto_size); // retval was previously set to the more important result above.
	//else v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
	// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items).

	if (sort_now)
		GuiType::LV_Sort(control, index, false, sort_now_direction);

	_o_return_retval;
}



void GuiControlType::LV_SetImageList(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns (MSDN): "handle to the image list previously associated with the control if successful; NULL otherwise."
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
// 2: Optional: Type of list.
{
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
	_o_return((size_t)ListView_SetImageList(hwnd, himl, list_type));
}



void GuiControlType::TV_AddModifyDelete(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// TV.Add():
// Returns the HTREEITEM of the item on success, zero on failure.
// Parameters:
//    1: Text/name of item.
//    2: Parent of item.
//    3: Options.
// TV.Modify():
// Returns the HTREEITEM of the item on success (to allow nested calls in script, zero on failure or partial failure.
// Parameters:
//    1: ID of item to modify.
//    2: Options.
//    3: New name.
// Parameters for TV.Delete():
//    1: ID of item to delete (if omitted, all items are deleted).
{
	GuiControlType &control = *this;
	auto mode = BuiltInFunctionID(aID);
	LPTSTR buf = _f_number_buf; // Resolve macro early for maintainability.

	if (mode == FID_TV_Delete)
	{
		// If param #1 is present but is zero, for safety it seems best not to do a delete-all (in case a
		// script bug is so rare that it is never caught until the script is distributed).  Another reason
		// is that a script might do something like TV.Delete(TV.GetSelection()), which would be desired
		// to fail not delete-all if there's ever any way for there to be no selection.
		_o_return(SendMessage(control.hwnd, TVM_DELETEITEM, 0
			, ParamIndexIsOmitted(0) ? NULL : (LPARAM)ParamIndexToInt64(0)));
	}

	// Since above didn't return, this is TV.Add() or TV.Modify().
	TVINSERTSTRUCT tvi; // It contains a TVITEMEX, which is okay even if MSIE pre-4.0 on Win95/NT because those OSes will simply never access the new/bottommost item in the struct.
	bool add_mode = (mode == FID_TV_Add); // For readability & maint.
	HTREEITEM retval;
	
	// Suppress any events raised by the changes made below:
	control.attrib |= GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS;

	LPTSTR options;
	if (add_mode) // TV.Add()
	{
		tvi.hParent = ParamIndexIsOmitted(1) ? NULL : (HTREEITEM)ParamIndexToInt64(1);
		tvi.hInsertAfter = TVI_LAST; // i.e. default is to insert the new item underneath the bottommost sibling.
		options = ParamIndexToOptionalString(2, buf);
		retval = 0; // Set default return value.
	}
	else // TV.Modify()
	{
		// NOTE: Must allow hitem==0 for TV.Modify, at least for the Sort option, because otherwise there would
		// be no way to sort the root-level items.
		tvi.item.hItem = (HTREEITEM)ParamIndexToInt64(0); // Load-time validation has ensured there is a first parameter for TV.Modify().
		// For modify-mode, set default return value to be "success" from this point forward.  Note that
		// in the case of sorting the root-level items, this will set it to zero, but since that almost
		// always succeeds and the script rarely cares whether it succeeds or not, adding code size for that
		// doesn't seem worth it:
		retval = tvi.item.hItem; // Set default return value.
		if (aParamCount < 2) // In one-parameter mode, simply select the item.
		{
			if (!TreeView_SelectItem(control.hwnd, tvi.item.hItem))
				retval = 0; // Override the HTREEITEM default value set above.
			control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
			_o_return((size_t)retval);
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
			// Since this option much more typically used with TV.Modify than TV.Add, the technique of
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
			else // TV.Modify(): Expand and collapse both require a message to work properly on an existing item.
				// Strangely, this generates a notification sometimes (such as the first time) but not for subsequent
				// expands/collapses of that same item.  Also, TVE_TOGGLE is not currently supported because it seems
				// like it would be too rarely used.
				if (!TreeView_Expand(control.hwnd, tvi.item.hItem, adding ? TVE_EXPAND : TVE_COLLAPSE))
					retval = 0; // Indicate partial failure by overriding the HTREEITEM return value set earlier.
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
					retval = 0; // Indicate partial failure by overriding the HTREEITEM return value set earlier.
		}
		// MUST BE LISTED LAST DUE TO "ELSE IF": Options valid only for TV.Add().
		else if (add_mode && !_tcsicmp(next_option, _T("First")))
		{
			tvi.hInsertAfter = TVI_FIRST; // For simplicity, the value of "adding" is ignored.
		}
		else if (add_mode && IsNumeric(next_option, false, false, false))
		{
			tvi.hInsertAfter = (HTREEITEM)ATOI64(next_option); // ATOI64 vs. ATOU avoids need for extra casting to avoid compiler warning.
		}
		else
		{
			control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
			aResultToken.ValueError(ERR_INVALID_OPTION, next_option);
			*option_end = orig_char; // See comment below.
			return;
		}

		// If the item was not handled by the above, ignore it because it is unknown.
		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
	}

	if (add_mode) // TV.Add()
	{
		tvi.item.pszText = ParamIndexToString(0, buf);
		tvi.item.mask |= TVIF_TEXT;
		tvi.item.hItem = TreeView_InsertItem(control.hwnd, &tvi); // Update tvi.item.hItem for convenience/maint. It's for use in later sections because retval is overridden to be zero for partial failure in modify-mode.
		retval = tvi.item.hItem; // Set return value.
	}
	else // TV.Modify()
	{
		if (!ParamIndexIsOmitted(2)) // An explicit empty string is allowed, which sets it to a blank value.  By contrast, if the param is omitted, the name is left changed.
		{
			tvi.item.pszText = ParamIndexToString(2, buf); // Reuse buf now that options (above) is done with it.
			tvi.item.mask |= TVIF_TEXT;
		}
		//else name/text parameter has been omitted, so don't change the item's name.
		if (tvi.item.mask != LVIF_STATE || tvi.item.stateMask) // An item's property or one of the state bits needs changing.
			if (!TreeView_SetItem(control.hwnd, &tvi.itemex))
				retval = 0; // Indicate partial failure by overriding the HTREEITEM return value set earlier.
	}

	if (ensure_visible) // Seems best to do this one prior to "select" below.
		SendMessage(control.hwnd, TVM_ENSUREVISIBLE, 0, (LPARAM)tvi.item.hItem); // Return value is ignored in this case, since its definition seems a little weird.
	if (ensure_visible_first) // Seems best to do this one prior to "select" below.
		TreeView_Select(control.hwnd, tvi.item.hItem, TVGN_FIRSTVISIBLE); // Return value is also ignored due to rarity, code size, and because most people wouldn't care about a failure even if for some reason it failed.
	if (select_flag)
		if (!TreeView_Select(control.hwnd, tvi.item.hItem, select_flag) && !add_mode) // Relies on short-circuit boolean order.
			retval = 0; // When not in add-mode, indicate partial failure by overriding the return value set earlier (add-mode should always return the new item's ID).

	control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
	_o_return((size_t)retval);
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



void GuiControlType::TV_GetRelatedItem(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// TV.GetParent/Child/Selection/Next/Prev(hitem):
// The above all return the HTREEITEM (or 0 on failure).
// When TV.GetNext's second parameter is present, the search scope expands to include not just siblings,
// but also children and parents, which allows a tree to be traversed from top to bottom without the script
// having to do something fancy.
{
	GuiControlType &control = *this;
	HWND control_hwnd = control.hwnd;

	HTREEITEM hitem = (HTREEITEM)ParamIndexToOptionalIntPtr(0, NULL);

	if (ParamIndexIsOmitted(1))
	{
		WPARAM flag;
		switch (aID)
		{
		case FID_TV_GetSelection: flag = TVGN_CARET; break; // TVGN_CARET is the focused item.
		case FID_TV_GetParent: flag = TVGN_PARENT; break;
		case FID_TV_GetChild: flag = TVGN_CHILD; break;
		case FID_TV_GetPrev: flag = TVGN_PREVIOUS; break;
		case FID_TV_GetNext: flag = !hitem ? TVGN_ROOT : TVGN_NEXT; break; // TV_GetNext(no-parameters) yields very first item in Tree (TVGN_ROOT).
		// Above: It seems best to treat hitem==0 as "get root", even though it sacrifices some error detection,
		// because not doing so would be inconsistent with the fact that TV.GetNext(0, "Full") does get the root
		// (which needs to be retained to make script loops to traverse entire tree easier).
		//case FID_TV_GetCount:
		default:
			// There's a known bug mentioned at MSDN that a TreeView might report a negative count when there
			// are more than 32767 items in it (though of course HTREEITEM values are never negative since they're
			// defined as unsigned pseudo-addresses).  But apparently, that bug only applies to Visual Basic and/or
			// older OSes, because testing shows that SendMessage(TVM_GETCOUNT) returns 32800+ when there are more
			// than 32767 items in the tree, even without casting to unsigned.  So I'm not sure exactly what the
			// story is with this, so for now just casting to UINT rather than something less future-proof like WORD:
			// Older note, apparently unneeded at least on XP SP2: Cast to WORD to convert -1 through -32768 to the
			// positive counterparts.
			_o_return((UINT)SendMessage(control_hwnd, TVM_GETCOUNT, 0, 0));
		}
		// Apparently there's no direct call to get the topmost ancestor of an item, presumably because it's rarely
		// needed.  Therefore, no such mode is provide here yet (the syntax TV.GetParent(hitem, true) could be supported
		// if it's ever needed).
		_o_return(SendMessage(control_hwnd, TVM_GETNEXTITEM, flag, (LPARAM)hitem));
	}

	// Since above didn't return, this TV.GetNext's 2-parameter mode, which has an expanded scope that includes
	// not just siblings, but also children and parents.  This allows a tree to be traversed from top to bottom
	// without the script having to do something fancy.
	TCHAR first_char_upper = ctoupper(*omit_leading_whitespace(ParamIndexToString(1, _f_number_buf))); // Resolve parameter #2.
	bool search_checkmark;
	if (first_char_upper == 'C')
		search_checkmark = true;
	else if (first_char_upper == 'F')
		search_checkmark = false;
	else // Reserve other option letters/words for future use by being somewhat strict.
		_o_throw_param(1);

	// When an actual item was specified, search begins at the item *after* it.  Otherwise (when NULL):
	// It's a special mode that always considers the root node first.  Otherwise, there would be no way
	// to start the search at the very first item in the tree to find out whether it's checked or not.
	hitem = GetNextTreeItem(control_hwnd, hitem); // Handles the comment above.
	if (!search_checkmark) // Simple tree traversal, so just return the next item (if any).
		_o_return((size_t)hitem); // OK if NULL.

	// Otherwise, search for the next item having a checkmark. For performance, it seems best to assume that
	// the control has the checkbox style (the script would realistically never call it otherwise, so the
	// control's style isn't checked.
	for (; hitem; hitem = GetNextTreeItem(control_hwnd, hitem))
		if (TreeView_GetCheckState(control_hwnd, hitem) == 1) // 0 means unchecked, -1 means "no checkbox image".
			_o_return((size_t)hitem); // OK if NULL.
	// Since above didn't return, the entire tree starting at the specified item has been searched,
	// with no match found.
	_o_return(0);
}



void GuiControlType::TV_Get(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// TV.Get()
// Returns: Varies depending on param #2.
// Parameters:
//    1: HTREEITEM.
//    2: Name of attribute to get.
// TV.GetText()
// Returns: Text on success.
// Throws on failure.
// Parameters:
//    1: HTREEITEM.
{
	bool get_text = aID == FID_TV_GetText;

	HWND control_hwnd = hwnd;

	if (!get_text)
	{
		// Loadtime validation has ensured that param #1 and #2 are present for all these cases.
		HTREEITEM hitem = (HTREEITEM)ParamIndexToInt64(0);
		UINT state_mask;
		switch (ctoupper(*omit_leading_whitespace(ParamIndexToString(1, _f_number_buf))))
		{
		case 'E': state_mask = TVIS_EXPANDED; break; // Expanded
		case 'C': state_mask = TVIS_STATEIMAGEMASK; break; // Checked
		case 'B': state_mask = TVIS_BOLD; break; // Bold
		//case 'S' for "Selected" is not provided because TV.GetSelection() seems to cover that well enough.
		//case 'P' for "is item a parent?" is not provided because TV.GetChild() seems to cover that well enough.
		// (though it's possible that retrieving TVITEM's cChildren would perform a little better).
		}
		// Below seems to be need a bit-AND with state_mask to work properly, at least on XP SP2.  Otherwise,
		// extra bits are present such as 0x2002 for "expanded" when it's supposed to be either 0x00 or 0x20.
		UINT result = state_mask & (UINT)SendMessage(control_hwnd, TVM_GETITEMSTATE, (WPARAM)hitem, state_mask);
		if (state_mask == TVIS_STATEIMAGEMASK)
		{
			if (result != 0x2000) // It doesn't have a checkmark state image.
				hitem = 0;
		}
		else // For all others, anything non-zero means the flag is present.
            if (!result) // Flag not present.
				hitem = 0;
		_o_return((size_t)hitem);
	}

	TCHAR text_buf[LV_TEXT_BUF_SIZE]; // i.e. uses same size as ListView.
	TVITEM tvi;
	tvi.hItem = (HTREEITEM)ParamIndexToInt64(0);
	tvi.mask = TVIF_TEXT;
	tvi.pszText = text_buf;
	tvi.cchTextMax = LV_TEXT_BUF_SIZE - 1; // -1 because of nagging doubt about size vs. length. Some MSDN examples subtract one), such as TabCtrl_GetItem()'s cchTextMax.

	if (SendMessage(control_hwnd, TVM_GETITEM, 0, (LPARAM)&tvi))
	{
		// Must use tvi.pszText vs. text_buf because MSDN says: "Applications should not assume that the text will
		// necessarily be placed in the specified buffer. The control may instead change the pszText member
		// of the structure to point to the new text rather than place it in the buffer."
		_o_return(tvi.pszText);
	}
	else
	{
		// On failure, it seems best to throw an exception.
		_o_throw(ERR_FAILED);
	}
}



void GuiControlType::TV_SetImageList(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns (MSDN): "handle to the image list previously associated with the control if successful; NULL otherwise."
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
// 2: Optional: Type of list.
{
	// Caller has ensured that there is at least one incoming parameter:
	HIMAGELIST himl = (HIMAGELIST)ParamIndexToInt64(0);
	int list_type;
	list_type = ParamIndexToOptionalInt(1, TVSIL_NORMAL);
	_o_return((size_t)TreeView_SetImageList(hwnd, himl, list_type));
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
	_f_return_i((size_t)ImageList_Create(GetSystemMetrics(param3 ? SM_CXICON : SM_CXSMICON)
		, GetSystemMetrics(param3 ? SM_CYICON : SM_CYSMICON)
		, ILC_MASK | ILC_COLOR32  // ILC_COLOR32 or at least something higher than ILC_COLOR is necessary to support true-color icons.
		, ParamIndexToOptionalInt(0, 2)    // cInitial. 2 seems a better default than one, since it might be common to have only two icons in the list.
		, ParamIndexToOptionalInt(1, 5)));  // cGrow.  Somewhat arbitrary default.
}



BIF_DECL(BIF_IL_Destroy)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
{
	// Load-time validation has ensured there is at least one parameter.
	// Returns nonzero if successful, or zero otherwise, so force it to conform to TRUE/FALSE for
	// better consistency with other functions:
	_f_return_i(ImageList_Destroy((HIMAGELIST)ParamIndexToInt64(0)) ? 1 : 0);
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
	HIMAGELIST himl = (HIMAGELIST)ParamIndexToInt64(0); // Load-time validation has ensured there is a first parameter.
	if (!himl)
		_f_throw_param(0);

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
	HBITMAP hbitmap = LoadPicture(ParamIndexToString(1, _f_number_buf) // Caller has ensured there are at least two parameters.
		, width, height, image_type, icon_number, false); // Defaulting to "false" for "use GDIplus" provides more consistent appearance across multiple OSes.
	if (!hbitmap)
		_f_return_i(0);

	int index;
	if (image_type == IMAGE_BITMAP) // In this mode, param3 is always assumed to be an RGB color.
	{
		// Return the index of the new image or 0 on failure.
		index = ImageList_AddMasked(himl, hbitmap, rgb_to_bgr((int)param3)) + 1; // +1 to convert to one-based.
		DeleteObject(hbitmap);
	}
	else // ICON or CURSOR.
	{
		// Return the index of the new image or 0 on failure.
		index = ImageList_AddIcon(himl, (HICON)hbitmap) + 1; // +1 to convert to one-based.
		DestroyIcon((HICON)hbitmap); // Works on cursors too.  See notes in LoadPicture().
	}
	_f_return_i(index);
}



BIF_DECL(BIF_LoadPicture)
{
	// h := LoadPicture(filename [, options, ByRef image_type])
	LPTSTR filename = ParamIndexToString(0, aResultToken.buf);
	LPTSTR options = ParamIndexToOptionalString(1);
	Var *image_type_var = ParamIndexToOutputVar(2);

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
	BuiltInFunctionID trim_type = _f_callee_id;

	LPTSTR buf = _f_retval_buf;
	size_t extract_length;
	LPTSTR str = ParamIndexToString(0, buf, &extract_length);

	LPTSTR result = str; // Prior validation has ensured at least 1 param.

	TCHAR omit_list_buf[MAX_NUMBER_SIZE]; // Support SYM_INTEGER/SYM_FLOAT even though it doesn't seem likely to happen.
	LPTSTR omit_list = ParamIndexIsOmitted(1) ? _T(" \t") : ParamIndexToString(1, omit_list_buf); // Default: space and tab.

	if (trim_type != FID_RTrim) // i.e. it's Trim() or LTrim()
	{
		result = omit_leading_any(result, omit_list, extract_length);
		extract_length -= (result - str); // Adjust for omitted characters.
	}
	if (extract_length && trim_type != FID_LTrim) // i.e. it's Trim() or RTrim();  THE LINE BELOW REQUIRES extract_length >= 1.
		extract_length = omit_trailing_any(result, omit_list, result + extract_length - 1);

	_f_return_p(result, extract_length);
}



BIF_DECL(BIF_VerCompare)
{
	_f_param_string(a, 0);
	_f_param_string(b, 1);
	_f_return(CompareVersion(a, b));
}



BIF_DECL(BIF_Type)
{
	_f_return_p(TokenTypeString(*aParam[0]));
}

// Returns the type name of the given value.
LPTSTR TokenTypeString(ExprTokenType &aToken)
{
	switch (TypeOfToken(aToken))
	{
	case SYM_STRING: return STRING_TYPE_STRING;
	case SYM_INTEGER: return INTEGER_TYPE_STRING;
	case SYM_FLOAT: return FLOAT_TYPE_STRING;
	case SYM_OBJECT: return TokenToObject(aToken)->Type();
	default: return _T(""); // For maintainability.
	}
}



void Object::Error__New(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	LPTSTR message;
	TCHAR what_buf[MAX_NUMBER_SIZE], extra_buf[MAX_NUMBER_SIZE];
	LPCTSTR what = ParamIndexToOptionalString(1, what_buf);
	Line *line = g_script.mCurrLine;

	if (aID == M_OSError__New && (ParamIndexIsOmitted(0) || ParamIndexIsNumeric(0)))
	{
		DWORD error = ParamIndexIsOmitted(0) ? g->LastError : (DWORD)ParamIndexToInt64(0);
		SetOwnProp(_T("Number"), error);
		
		// Determine message based on error number.
		DWORD message_buf_size = _f_retval_buf_size;
		message = _f_retval_buf;
		DWORD size = (DWORD)_sntprintf(message, message_buf_size, (int)error < 0 ? _T("(0x%X) ") : _T("(%i) "), error);
		if (error) // Never show "Error: (0) The operation completed successfully."
			size += FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, NULL, error, 0, message + size, message_buf_size - size, NULL);
		if (size)
		{
			if (message[size - 1] == '\n')
				message[--size] = '\0';
			if (message[size - 1] == '\r')
				message[--size] = '\0';
		}
	}
	else
		message = ParamIndexIsOmitted(0) ? Type() : ParamIndexToString(0, _f_number_buf);

#ifndef CONFIG_DEBUGGER
	if (ParamIndexIsOmitted(1) && g->CurrentFunc)
		what = g->CurrentFunc->mName;
	SetOwnProp(_T("Stack"), _T("")); // Avoid "unknown property" in compiled scripts.
#else
	DbgStack::Entry *stack_top = g_Debugger.mStack.mTop - 1;
	if (stack_top->type == DbgStack::SE_BIF && _tcsicmp(what, stack_top->func->mName))
		--stack_top;

	if (ParamIndexIsOmitted(1)) // "What"
	{
		if (g->CurrentFunc)
			what = g->CurrentFunc->mName;
	}
	else
	{
		int offset = ParamIndexIsNumeric(1) ? ParamIndexToInt(1) : 0;
		for (auto se = stack_top; se >= g_Debugger.mStack.mBottom; --se)
		{
			if (++offset == 0 || *what && !_tcsicmp(se->Name(), what))
			{
				line = se > g_Debugger.mStack.mBottom ? se[-1].line : se->line;
				// se->line contains the line at the given offset from the top of the stack.
				// Rather than returning the name of the function or sub which contains that
				// line, return the name of the function or sub which that line called.
				// In other words, an offset of -1 gives the name of the current function and
				// the file and number of the line which it was called from.
				what = se->Name();
				stack_top = se;
				break;
			}
			if (se->type == DbgStack::SE_Thread)
				break; // Look only within the current thread.
		}
	}

	TCHAR stack_buf[2048];
	GetScriptStack(stack_buf, _countof(stack_buf), stack_top);
	SetOwnProp(_T("Stack"), stack_buf);
#endif

	LPTSTR extra = ParamIndexToOptionalString(2, extra_buf);

	SetOwnProp(_T("Message"), message);
	SetOwnProp(_T("What"), const_cast<LPTSTR>(what));
	SetOwnProp(_T("File"), Line::sSourceFile[line->mFileIndex]);
	SetOwnProp(_T("Line"), line->mLineNumber);
	SetOwnProp(_T("Extra"), extra);
}



////////////////////////////////////////////////////////
// HELPER FUNCTIONS FOR TOKENS AND BUILT-IN FUNCTIONS //
////////////////////////////////////////////////////////

BOOL ResultToBOOL(LPTSTR aResult)
{
	UINT c1 = (UINT)*aResult; // UINT vs. UCHAR might squeeze a little more performance out of it.
	if (c1 > 48)     // Any UCHAR greater than '0' can't be a space(32), tab(9), '+'(43), or '-'(45), or '.'(46)...
		return TRUE; // ...so any string starting with c1>48 can't be anything that's false (e.g. " 0", "+0", "-0", ".0", "0").
	if (!c1 // Besides performance, must be checked anyway because otherwise IsNumeric() would consider "" to be non-numeric and thus TRUE.
		|| c1 == 48 && !aResult[1]) // The string "0" seems common enough to detect explicitly, for performance.
		return FALSE;
	// IsNumeric() is called below because there are many variants of a false string:
	// e.g. "0", "0.0", "0x0", ".0", "+0", "-0", and " 0" (leading spaces/tabs).
	switch (IsNumeric(aResult, true, false, true)) // It's purely numeric and not all whitespace (and due to earlier check, it's not blank).
	{
	case PURE_INTEGER: return ATOI64(aResult) != 0; // Could call ATOF() for both integers and floats; but ATOI64() probably performs better, and a float comparison to 0.0 might be a slower than an integer comparison to 0.
	case PURE_FLOAT:   return _tstof(aResult) != 0.0; // atof() vs. ATOF() because PURE_FLOAT is never hexadecimal.
	default: // PURE_NOT_NUMERIC.
		// Even a string containing all whitespace would be considered non-numeric since it's a non-blank string
		// that isn't equal to 0.
		return TRUE;
	}
}



BOOL VarToBOOL(Var &aVar)
{
	if (!aVar.HasContents()) // Must be checked first because otherwise IsNumeric() would consider "" to be non-numeric and thus TRUE.  For performance, it also exploits the binary number cache.
		return FALSE;
	switch (aVar.IsNumeric())
	{
	case PURE_INTEGER:
		return aVar.ToInt64() != 0;
	case PURE_FLOAT:
		return aVar.ToDouble() != 0.0;
	default:
		// Even a string containing all whitespace would be considered non-numeric since it's a non-blank string
		// that isn't equal to 0.
		return TRUE;
	}
}



BOOL TokenToBOOL(ExprTokenType &aToken)
{
	switch (aToken.symbol)
	{
	case SYM_INTEGER: // Probably the most common; e.g. both sides of "if (x>3 and x<6)" are the number 1/0.
		return aToken.value_int64 != 0; // Force it to be purely 1 or 0.
	case SYM_VAR:
		return VarToBOOL(*aToken.var);
	case SYM_FLOAT:
		return aToken.value_double != 0.0;
	case SYM_STRING:
		return ResultToBOOL(aToken.marker);
	default:
		// The only remaining valid symbol is SYM_OBJECT, which is always TRUE.
		return TRUE;
	}
}



ToggleValueType TokenToToggleValue(ExprTokenType &aToken)
{
	if (TokenIsNumeric(aToken))
	switch (TokenToInt64(aToken)) // Reserve other values for potential future use by requiring exact match.
	{
	case 1: return TOGGLED_ON;
	case 0: return TOGGLED_OFF;
	case -1: return TOGGLE;
	}
	return TOGGLE_INVALID;
}



SymbolType TokenIsNumeric(ExprTokenType &aToken)
{
	switch(aToken.symbol)
	{
	case SYM_INTEGER:
	case SYM_FLOAT:
		return aToken.symbol;
	case SYM_VAR: 
		return aToken.var->IsNumeric();
	default: // SYM_STRING: Callers of this function expect a "numeric" result for numeric strings.
		return IsNumeric(aToken.marker, true, false, true);
	}
}


SymbolType TokenIsPureNumeric(ExprTokenType &aToken)
{
	switch (aToken.symbol)
	{
	case SYM_INTEGER:
	case SYM_FLOAT:
		return aToken.symbol;
	case SYM_VAR:
		if (!aToken.var->IsUninitializedNormalVar()) // Caller doesn't want a warning, so avoid calling Contents().
			return aToken.var->IsPureNumeric();
		//else fall through:
	default:
		return PURE_NOT_NUMERIC;
	}
}


SymbolType TokenIsPureNumeric(ExprTokenType &aToken, SymbolType &aNumType)
// This function is called very frequently by ExpandExpression(), which needs to distinguish
// between numeric strings and pure numbers, but still needs to know if the string is numeric.
{
	switch(aToken.symbol)
	{
	case SYM_INTEGER:
	case SYM_FLOAT:
		return aNumType = aToken.symbol;
	case SYM_VAR:
		if (aToken.var->IsUninitializedNormalVar()) // Caller doesn't want a warning, so avoid calling Contents().
			return aNumType = PURE_NOT_NUMERIC; // i.e. empty string is non-numeric.
		if (aNumType = aToken.var->IsPureNumeric())
			return aNumType; // This var contains a pure binary number.
		// Otherwise, it might be a numeric string (i.e. impure).
		aNumType = aToken.var->IsNumeric(); // This also caches the PURE_NOT_NUMERIC result if applicable.
		return PURE_NOT_NUMERIC;
	case SYM_STRING:
		aNumType = IsNumeric(aToken.marker, true, false, true);
		return PURE_NOT_NUMERIC;
	default:
		// Only SYM_OBJECT and SYM_MISSING should be possible.
		return aNumType = PURE_NOT_NUMERIC;
	}
}


BOOL TokenIsEmptyString(ExprTokenType &aToken)
{
	switch (aToken.symbol)
	{
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


__int64 TokenToInt64(ExprTokenType &aToken)
// Caller has ensured that any SYM_VAR's Type() is VAR_NORMAL.
// Converts the contents of aToken to a 64-bit int.
{
	// Some callers, such as those that cast our return value to UINT, rely on the use of 64-bit
	// to preserve unsigned values and also wrap any signed values into the unsigned domain.
	switch (aToken.symbol)
	{
		case SYM_INTEGER:	return aToken.value_int64;
		case SYM_FLOAT:		return (__int64)aToken.value_double;
		case SYM_VAR:		return aToken.var->ToInt64();
		case SYM_STRING:	return ATOI64(aToken.marker);
	}
	// Since above didn't return, it can only be SYM_OBJECT or not an operand.
	return 0;
}



double TokenToDouble(ExprTokenType &aToken, BOOL aCheckForHex)
// Caller has ensured that any SYM_VAR's Type() is VAR_NORMAL.
// Converts the contents of aToken to a double.
{
	switch (aToken.symbol)
	{
		case SYM_FLOAT:		return aToken.value_double;
		case SYM_INTEGER:	return (double)aToken.value_int64;
		case SYM_VAR:		return aToken.var->ToDouble();
		case SYM_STRING:	return aCheckForHex ? ATOF(aToken.marker) : _tstof(aToken.marker); // atof() omits the check for hexadecimal.
	}
	// Since above didn't return, it can only be SYM_OBJECT or not an operand.
	return 0;
}



LPTSTR TokenToString(ExprTokenType &aToken, LPTSTR aBuf, size_t *aLength)
// Returns "" on failure to simplify logic in callers.  Otherwise, it returns either aBuf (if aBuf was needed
// for the conversion) or the token's own string.  aBuf may be NULL, in which case the caller presumably knows
// that this token is SYM_STRING or SYM_VAR (or the caller wants "" back for anything other than those).
// If aBuf is not NULL, caller has ensured that aBuf is at least MAX_NUMBER_SIZE in size.
{
	LPTSTR result;
	switch (aToken.symbol)
	{
	case SYM_VAR: // Caller has ensured that any SYM_VAR's Type() is VAR_NORMAL.
		result = aToken.var->Contents(); // Contents() vs. mCharContents in case mCharContents needs to be updated by Contents().
		if (aLength)
			*aLength = aToken.var->Length();
		return result;
	case SYM_STRING:
		result = aToken.marker;
		if (aLength)
		{
			if (aToken.marker_length == -1)
				break; // Call _tcslen(result) below.
			*aLength = aToken.marker_length;
		}
		return result;
	case SYM_INTEGER:
		result = aBuf ? ITOA64(aToken.value_int64, aBuf) : _T("");
		break;
	case SYM_FLOAT:
		if (aBuf)
		{
			int length = FTOA(aToken.value_double, aBuf, MAX_NUMBER_SIZE);
			if (aLength)
				*aLength = length;
			return aBuf;
		}
		//else continue on to return the default at the bottom.
	//case SYM_OBJECT: // Treat objects as empty strings (or TRUE where appropriate).
	default:
		result = _T("");
	}
	if (aLength) // Caller wants to know the string's length as well.
		*aLength = _tcslen(result);
	return result;
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
		//case SYM_OBJECT: // L31: Treat objects as empty strings (or TRUE where appropriate).
		//case SYM_MISSING:
		default:
			return FAIL;
	}
	// Since above didn't return, interpret "str" as a number.
	switch (aOutput.symbol = IsNumeric(str, true, false, true))
	{
	case PURE_INTEGER:
		aOutput.value_int64 = ATOI64(str);
		break;
	case PURE_FLOAT:
		aOutput.value_double = ATOF(str);
		break;
	default: // Not a pure number.
		return FAIL;
	}
	return OK; // Since above didn't return, indicate success.
}



StringCaseSenseType TokenToStringCase(ExprTokenType& aToken)
{
	// Pure integers 1 and 0 corresponds to SCS_SENSITIVE and SCS_INSENSITIVE, respectively.
	// Pure floats returns SCS_INVALID.
	// For strings see Line::ConvertStringCaseSense.
	LPTSTR str = NULL;
	__int64 int_val = 0;
	switch (aToken.symbol)
	{
	case SYM_VAR:
		
		switch (aToken.var->IsPureNumeric())
		{
		case PURE_INTEGER: int_val = aToken.var->ToInt64(); break;
		case PURE_NOT_NUMERIC: str = aToken.var->Contents(); break;
		case PURE_FLOAT: 
		default:	
			return SCS_INVALID;
		}
		break;

	case SYM_INTEGER: int_val = TokenToInt64(aToken); break;
	case SYM_FLOAT: return SCS_INVALID;
	default: str = TokenToString(aToken); break;
	}
	if (str)
		return !_tcsicmp(str, _T("Logical"))	? SCS_INSENSITIVE_LOGICAL
												: Line::ConvertStringCaseSense(str);
	return int_val == 1 ? SCS_SENSITIVE						// 1	- Sensitive
						: (int_val == 0 ? SCS_INSENSITIVE	// 0	- Insensitive
										: SCS_INVALID);		// else - invalid.
}



Var *TokenToOutputVar(ExprTokenType &aToken)
{
	if (aToken.symbol == SYM_VAR && aToken.var_usage != Script::VARREF_READ)
		return aToken.var;
	return dynamic_cast<VarRef *>(TokenToObject(aToken));
}



IObject *TokenToObject(ExprTokenType &aToken)
// L31: Returns IObject* from SYM_OBJECT or SYM_VAR (where var->HasObject()), NULL for other tokens.
// Caller is responsible for calling AddRef() if that is appropriate.
{
	if (aToken.symbol == SYM_OBJECT)
		return aToken.object;
	
	if (aToken.symbol == SYM_VAR)
		return aToken.var->ToObject();

	return NULL;
}



ResultType ValidateFunctor(IObject *aFunc, int aParamCount, ResultToken &aResultToken, int *aUseMinParams, bool aShowError)
{
	ASSERT(aFunc);
	__int64 min_params = 0, max_params = INT_MAX;
	auto min_result = aParamCount == -1 ? INVOKE_NOT_HANDLED
		: GetObjectIntProperty(aFunc, _T("MinParams"), min_params, aResultToken, true);
	if (!min_result)
		return FAIL;
	bool has_minparams = min_result != INVOKE_NOT_HANDLED; // For readability.
	if (aUseMinParams) // CallbackCreate's signal to default to MinParams.
	{
		if (!has_minparams)
			return aShowError ? aResultToken.UnknownMemberError(ExprTokenType(aFunc), IT_GET, _T("MinParams")) : CONDITION_FALSE;
		*aUseMinParams = aParamCount = (int)min_params;
	}
	else if (has_minparams && aParamCount < (int)min_params)
		return aShowError ? aResultToken.ValueError(ERR_INVALID_FUNCTOR) : CONDITION_FALSE;
	auto max_result = (aParamCount <= 0 || has_minparams && min_params == aParamCount)
		? INVOKE_NOT_HANDLED // No need to check MaxParams in the above cases.
		: GetObjectIntProperty(aFunc, _T("MaxParams"), max_params, aResultToken, true);
	if (!max_result)
		return FAIL;
	if (max_result != INVOKE_NOT_HANDLED && aParamCount > (int)max_params)
	{
		__int64 is_variadic = 0;
		auto result = GetObjectIntProperty(aFunc, _T("IsVariadic"), is_variadic, aResultToken, true);
		if (!result)
			return FAIL;
		if (!is_variadic) // or not defined.
			return aShowError ? aResultToken.ValueError(ERR_INVALID_FUNCTOR) : CONDITION_FALSE;
	}
	// If either MinParams or MaxParams was confirmed to exist, this is likely a valid
	// function object, so skip the following check for performance.  Otherwise, catch
	// likely errors by checking that the object is callable.
	if (min_result == INVOKE_NOT_HANDLED && max_result == INVOKE_NOT_HANDLED)
		if (Object *obj = dynamic_cast<Object *>(aFunc))
			if (!obj->HasMethod(_T("Call")))
				return aShowError ? aResultToken.UnknownMemberError(ExprTokenType(aFunc), IT_CALL, _T("Call")) : CONDITION_FALSE;
		// Otherwise: COM objects can be callable via DISPID_VALUE.  There's probably
		// no way to determine whether the object supports that without invoking it.
	return OK;
}



ResultType TokenSetResult(ResultToken &aResultToken, LPCTSTR aValue, size_t aLength)
// Utility function for handling memory allocation and return to callers of built-in functions; based on BIF_SubStr.
// Returns FAIL if malloc failed, in which case our caller is responsible for returning a sensible default value.
{
	if (aLength == -1)
		aLength = _tcslen(aValue); // Caller must not pass NULL for aResult in this case.
	if (aLength <= MAX_NUMBER_LENGTH) // v1.0.46.01: Avoid malloc() for small strings.  However, this improves speed by only 10% in a test where random 25-byte strings were extracted from a 700 KB string (probably because VC++'s malloc()/free() are very fast for small allocations).
		aResultToken.marker = aResultToken.buf; // Store the address of the result for the caller.
	else
	{
		// Caller has provided a mem_to_free (initially NULL) as a means of passing back memory we allocate here.
		// So if we change "result" to be non-NULL, the caller will take over responsibility for freeing that memory.
		if (   !(aResultToken.mem_to_free = tmalloc(aLength + 1))   ) // Out of memory.
			return aResultToken.MemoryError();
		aResultToken.marker = aResultToken.mem_to_free; // Store the address of the result for the caller.
	}
	if (aValue) // Caller may pass NULL to retrieve a buffer of sufficient size.
		tmemcpy(aResultToken.marker, aValue, aLength);
	aResultToken.marker[aLength] = '\0'; // Must be done separately from the memcpy() because the memcpy() might just be taking a substring (i.e. long before result's terminator).
	aResultToken.marker_length = aLength;
	return OK;
}



// TypeOfToken: Similar result to TokenIsPureNumeric, but may return SYM_OBJECT.
SymbolType TypeOfToken(ExprTokenType &aToken)
{
	switch (aToken.symbol)
	{
	case SYM_VAR:
		switch (aToken.var->IsPureNumericOrObject())
		{
		case VAR_ATTRIB_IS_INT64: return SYM_INTEGER;
		case VAR_ATTRIB_IS_DOUBLE: return SYM_FLOAT;
		case VAR_ATTRIB_IS_OBJECT: return SYM_OBJECT;
		}
		return SYM_STRING;
	default: // Providing a default case produces smaller code on release builds as the compiler can omit the other symbol checks.
#ifdef _DEBUG
		MsgBox(_T("DEBUG: Unhandled symbol type."));
#endif
	case SYM_STRING:
	case SYM_INTEGER:
	case SYM_FLOAT:
	case SYM_OBJECT:
	case SYM_MISSING:
		return aToken.symbol;
	}
}



// TypeOfToken: Similar result to TokenIsPureNumeric, but may return SYM_OBJECT.
SymbolType TypeOfToken(ExprTokenType &aToken, SymbolType &aIsNum)
{
	switch (aToken.symbol)
	{
	case SYM_VAR:
		switch (aToken.var->IsPureNumericOrObject())
		{
		case VAR_ATTRIB_IS_INT64:
			aIsNum = PURE_INTEGER;
			return SYM_INTEGER;
		case VAR_ATTRIB_IS_DOUBLE:
			aIsNum = PURE_FLOAT;
			return SYM_FLOAT;
		case VAR_ATTRIB_IS_OBJECT:
			aIsNum = PURE_NOT_NUMERIC;
			return SYM_OBJECT;
		default:
			aIsNum = aToken.var->IsNumeric();
			return SYM_STRING;
		}
	case SYM_STRING:
	case SYM_MISSING:
		aIsNum = IsNumeric(aToken.marker, true, false, true);
		return SYM_STRING;
	case SYM_INTEGER:
	case SYM_FLOAT:
		aIsNum = aToken.symbol;
		return aToken.symbol;
	default:
#ifdef _DEBUG
		MsgBox(_T("DEBUG: Unhandled symbol type."));
#endif
	case SYM_OBJECT:
		aIsNum = PURE_NOT_NUMERIC;
		return aToken.symbol;
	}
}



ResultType ResultToken::Return(LPTSTR aValue, size_t aLength)
// Copy and return a string.
{
	ASSERT(aValue);
	symbol = SYM_STRING;
	return TokenSetResult(*this, aValue, aLength);
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
		if (IsNumeric(aBuf + 3, false, false))
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
		return IsKeyToggledOn(aVK); // This also works for non-"lock" keys, but in that case the toggle state can be out of sync with other processes/threads.
	case KEYSTATE_PHYSICAL: // Physical state of key.
		if (IsMouseVK(aVK)) // mouse button
		{
			if (g_MouseHook) // mouse hook is installed, so use it's tracking of physical state.
				return g_PhysicalKeyState[aVK] & STATE_DOWN;
			else
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
			else
				return IsKeyDownAsync(aVK);
		}
	} // switch()

	// Otherwise, use the default state-type: KEYSTATE_LOGICAL
	// On XP/2K at least, a key can be physically down even if it isn't logically down,
	// which is why the below specifically calls IsKeyDown() rather than some more
	// comprehensive method such as consulting the physical key state as tracked by the hook:
	// v1.0.42.01: For backward compatibility, the following hasn't been changed to IsKeyDownAsync().
	// One example is the journal playback hook: when a window owned by the script receives
	// such a keystroke, only GetKeyState() can detect the changed state of the key, not GetAsyncKeyState().
	// A new mode can be added to KeyWait & GetKeyState if Async is ever explicitly needed.
	return IsKeyDown(aVK);
}



bool ScriptGetJoyState(JoyControls aJoy, int aJoystickID, ExprTokenType &aToken, LPTSTR aBuf)
// Caller must ensure that aToken.marker is a buffer large enough to handle the longest thing put into
// it here, which is currently jc.szPname (size=32). Caller has set aToken.symbol to be SYM_STRING.
// aToken is used for the value being returned by GetKeyState() to the script, while this function's
// bool return value is used only by KeyWait, so is false for "up" and true for "down".
// If there was a problem determining the position/state, aToken is made blank and false is returned.
{
	// Set default in case of early return.
	aToken.symbol = SYM_STRING;
	aToken.marker = aBuf;
	*aBuf = '\0'; // Blank vs. string "0" serves as an indication of failure.

	if (!aJoy) // Currently never called this way.
		return false; // And leave aToken set to blank.

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
			return false; // And leave aToken set to blank.
		if (aJoy_is_button)
		{
			bool is_down = ((jie.dwButtons >> (aJoy - JOYCTRL_1)) & (DWORD)0x01);
			aToken.symbol = SYM_INTEGER; // Override default type.
			aToken.value_int64 = is_down; // Always 1 or 0, since it's "bool" (and if not for that, bitwise-and).
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

	case JOYCTRL_POV:
		aToken.symbol = SYM_INTEGER; // Override default type.
		if (jie.dwPOV == JOY_POVCENTERED) // Need to explicitly compare against JOY_POVCENTERED because it's a WORD not a DWORD.
		{
			aToken.value_int64 = -1;
			return false;
		}
		else
		{
			aToken.value_int64 = jie.dwPOV;
			return true;
		}
		// No break since above always returns.

	case JOYCTRL_NAME:
		_tcscpy(aToken.marker, jc.szPname);
		return false; // Return value not used.

	case JOYCTRL_BUTTONS:
		aToken.symbol = SYM_INTEGER; // Override default type.
		aToken.value_int64 = jc.wNumButtons; // wMaxButtons is the *driver's* max supported buttons.
		return false; // Return value not used.

	case JOYCTRL_AXES:
		aToken.symbol = SYM_INTEGER; // Override default type.
		aToken.value_int64 = jc.wNumAxes; // wMaxAxes is the *driver's* max supported axes.
		return false; // Return value not used.

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
		return false; // Return value not used.
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
	// Windows XP/2003 would require PROCESS_QUERY_INFORMATION, but those OSes are not supported.
	if (  !(hproc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, aProcessID))  )
		return 0;

	// Benchmarks showed that attempting GetModuleBaseName/GetModuleFileNameEx
	// first did not help performance.  Also, QueryFullProcessImageName appeared
	// to be slower than the following.
	DWORD buf_length = GetProcessImageFileName(hproc, aBuf, aBufSize);
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

	CloseHandle(hproc);
	return buf_length;
}

__int64 pow_ll(__int64 base, __int64 exp)
{
	/*
	Caller must ensure exp >= 0
	Below uses 'a^b' to denote 'raising a to the power of b'.
	Computes and returns base^exp. If the mathematical result doesn't fit in __int64, the result is undefined.
	By convention, x^0 returns 1, even when x == 0,	caller should ensure base is non-zero when exp is zero to handle 0^0.
	*/
	if (exp == 0)
		return 1ll;

	// based on: https://en.wikipedia.org/wiki/Exponentiation_by_squaring (2018-11-03)
	__int64 result = 1;
	while (exp > 1)
	{
		if (exp % 2) // exp is odd
			result *= base;
		base *= base;
		exp /= 2;
	}
	return result * base;
}
