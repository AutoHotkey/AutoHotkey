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
#include "script.h"
#include "script_func_impl.h"



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
