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

#include "stdafx.h"
#include "script.h"



FResult GuiControlType::SB_SetIcon(StrArg aFilename, optl<int> aIconNumber, optl<UINT> aPartNumber, UINT_PTR &aRetVal)
{
	GuiType& gui = *this->gui;
    
    UINT part_index = aPartNumber.value_or(1) - 1;
    if (part_index > 255) // Also covers wrapped negative values.
        return FR_E_ARG(2);
    
    int unused, icon_number = aIconNumber.value_or(1);
    if (icon_number == 0) // Must be != 0 to tell LoadPicture that "icon must be loaded, never a bitmap".
        icon_number = 1;
    auto hicon = (HICON)LoadPicture(aFilename // Load-time validation has ensured there is at least one parameter.
        , GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON) // Apparently the bar won't scale them for us.
        , unused, icon_number, false); // Defaulting to "false" for "use GDIplus" provides more consistent appearance across multiple OSes.
    if (!hicon)
        return FError(ERR_LOAD_ICON, aFilename);
    
    HICON hicon_old = (HICON)SendMessage(hwnd, SB_GETICON, part_index, 0); // Get the old one before setting the new one.
    // There's no need to keep track of the icon handle since it will be destroyed automatically if the
    // window is destroyed or the program exits.
    if (!SendMessage(hwnd, SB_SETICON, part_index, (LPARAM)hicon))
    {
        DestroyIcon(hicon);
        return FR_E_FAILED;
    }
    if (hicon_old)
        // Although the old icon is automatically destroyed here, the script can call SendMessage(SB_SETICON)
        // itself if it wants to work with HICONs directly (for performance reasons, etc.)
        DestroyIcon(hicon_old);
    // Return the icon handle mainly for historical reasons.  Old comments indicated the return value was
    // to allow the script to destroy the HICON later when it doesn't need it, but that doesn't make sense
    // since destroying the Gui or setting a new icon would implicitly cause the old icon to be destroyed.
    aRetVal = (UINT_PTR)hicon;
    return OK;
}



FResult GuiControlType::SB_SetParts(VariantParams &aParam, UINT& aRetVal)
{
	GuiType& gui = *this->gui;
    
	int old_part_count, new_part_count;
    int edge, part[256]; // Load-time validation has ensured aParamCount is under 255, so it shouldn't overflow.
    for (edge = 0, new_part_count = 0; new_part_count < aParam.count; ++new_part_count)
    {
        auto &value = *aParam.value[new_part_count];
        if (!TokenIsNumeric(value))
            return FParamError(new_part_count, &value, _T("Number"));
        int width = (int)TokenToInt64(value);
        if (width < 0)
            return FR_E_ARG(new_part_count);
        edge += gui.Scale(width);
        part[new_part_count] = edge;
    }
    // For code simplicity, there is currently no means to have the last part of the bar use less than
    // all of the bar's remaining width.  The desire to do so seems rare, especially since the script can
    // add an extra/unused part at the end to achieve nearly (or perhaps exactly) the same effect.
    part[new_part_count++] = -1; // Make the last part use the remaining width of the bar.

    old_part_count = (int)SendMessage(hwnd, SB_GETPARTS, 0, NULL); // MSDN: "This message always returns the number of parts in the status bar [regardless of how it is called]".
    if (old_part_count > new_part_count) // Some parts are being deleted, so destroy their icons.  See other notes in GuiType::Destroy() for explanation.
        for (int i = new_part_count; i < old_part_count; ++i) // Verified correct.
            if (auto hicon = (HICON)SendMessage(hwnd, SB_GETICON, i, 0))
                DestroyIcon(hicon);

    if (!SendMessage(hwnd, SB_SETPARTS, new_part_count, (LPARAM)part))
        return FR_E_FAILED;
    
    // hwnd is returned for historical reasons; there are more direct and ubiquitous ways to get it.
    aRetVal = (UINT)(size_t)hwnd;
    return OK;
}



FResult GuiControlType::SB_SetText(StrArg aNewText, optl<UINT> aPartNumber, optl<UINT> aStyle)
{
    UINT part_index = aPartNumber.value_or(1) - 1;
    if (part_index > 255) // Also covers wrapped negative values.
        return FR_E_ARG(2);
    UINT style = aStyle.value_or(0);
    if (style > 0xFF)
        return FR_E_ARG(3);
	return SendMessage(hwnd, SB_SETTEXT, (part_index | (style << 8)), (LPARAM)aNewText) ? OK : FR_E_FAILED;
}


