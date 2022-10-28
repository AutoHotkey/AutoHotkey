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
#include "script_func_impl.h"



FResult GuiControlType::TV_AddModify(bool add_mode, UINT_PTR aItemID, UINT_PTR aParentItemID, optl<StrArg> aOptions, optl<StrArg> aName, UINT_PTR &aRetVal)
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
{
	GuiControlType &control = *this;

	// Since above didn't return, this is TV.Add() or TV.Modify().
	TVINSERTSTRUCT tvi; // It contains a TVITEMEX, which is okay even if MSIE pre-4.0 on Win95/NT because those OSes will simply never access the new/bottommost item in the struct.
	
	// Suppress any events raised by the changes made below:
	control.attrib |= GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS;

	if (add_mode) // TV.Add()
	{
		tvi.hParent = (HTREEITEM)aParentItemID;
		tvi.hInsertAfter = TVI_LAST; // i.e. default is to insert the new item underneath the bottommost sibling.
	}
	else // TV.Modify()
	{
		// NOTE: Must allow hitem==0 for TV.Modify, at least for the Sort option, because otherwise there would
		// be no way to sort the root-level items.
		tvi.item.hItem = (HTREEITEM)aItemID;
		aRetVal = (UINT_PTR)tvi.item.hItem; // In all cases, return the item ID.
		if (!aOptions.has_value() && !aName.has_value()) // In one-parameter mode, simply select the item.
		{
			TreeView_SelectItem(control.hwnd, tvi.item.hItem);
			control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
			return OK;
		}
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
	auto options = aOptions.value_or_empty();
	TCHAR option_word[16]; // Enough for any single option word or number, with room to avoid false positives due to truncation.
	LPCTSTR next_option, option_end;
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
		for (option_end = next_option; *option_end && !IS_SPACE_OR_TAB(*option_end); ++option_end);
		if (option_end == next_option)
			continue; // i.e. the string contains a + or - with a space or tab after it, which is intentionally ignored.
		
		// if ((option_end[-1] == '1' || option_end[-1] == '0') && )

		// Make a copy to simplify comparisons below.
		tcslcpy(option_word, next_option, min((option_end - next_option) + 1, _countof(option_word)));

		if (!_tcsicmp(option_word, _T("Select"))) // Could further allow "ed" suffix by checking for that inside, but "Selected" is getting long so it doesn't seem something many would want to use.
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
		else if (!_tcsicmp(option_word, _T("Vis")))
		{
			// Since this option much more typically used with TV.Modify than TV.Add, the technique of
			// Vis%VarContainingOneOrZero% isn't supported, to reduce code size.
			ensure_visible = adding;
		}
		else if (!_tcsicmp(option_word, _T("VisFirst")))
		{
			ensure_visible_first = adding;
		}
		else if (!_tcsnicmp(option_word, _T("Bold"), 4))
		{
			next_option += 4;
			if (next_option < option_end && !ATOI(next_option)) // If it's Bold0, invert the mode.
				adding = !adding;
			tvi.item.stateMask |= TVIS_BOLD;
			if (adding)
				tvi.item.state |= TVIS_BOLD;
			//else removing, so the fact that this TVIS flag has just been added to the stateMask above
			// but is absent from item.state should remove this attribute from the item.
		}
		else if (!_tcsnicmp(option_word, _T("Expand"), 6))
		{
			next_option += 6;
			if (next_option < option_end && !ATOI(next_option)) // If it's Expand0, invert the mode to become "collapse".
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
				TreeView_Expand(control.hwnd, tvi.item.hItem, adding ? TVE_EXPAND : TVE_COLLAPSE);
		}
		else if (!_tcsnicmp(option_word, _T("Check"), 5))
		{
			// The rationale for not checking for an optional "ed" suffix here and incrementing next_option by 2
			// is that: 1) It would be inconsistent with the lack of support for "selected" (see reason above);
			// 2) Checkboxes in a ListView are fairly rarely used, so code size reduction might be more important.
			next_option += 5;
			if (next_option < option_end && !ATOI(next_option)) // If it's Check0, invert the mode to become "unchecked".
				adding = !adding;
			//else removing, so the fact that this TVIS flag has just been added to the stateMask above
			// but is absent from item.state should remove this attribute from the item.
			tvi.item.stateMask |= TVIS_STATEIMAGEMASK;  // Unlike ListViews, Tree checkmarks can be applied in the same step as creating a Tree item.
			tvi.item.state |= adding ? 0x2000 : 0x1000; // The #1 image is "unchecked" and the #2 is "checked".
		}
		else if (!_tcsnicmp(option_word, _T("Icon"), 4))
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
		else if (!_tcsicmp(option_word, _T("Sort")))
		{
			if (add_mode)
				tvi.hInsertAfter = TVI_SORT; // For simplicity, the value of "adding" is ignored.
			else
				// Somewhat debatable, but it seems best to report failure via the return value even though
				// failure probably only occurs when the item has no children, and the script probably
				// doesn't often care about such failures.  It does result in the loss of the HTREEITEM return
				// value, but even if that call is nested in another, the zero should produce no effect in most cases.
				TreeView_SortChildren(control.hwnd, tvi.item.hItem, FALSE); // Best default seems no-recurse, since typically this is used after a user edits merely a single item.
		}
		// MUST BE LISTED LAST DUE TO "ELSE IF": Options valid only for TV.Add().
		else if (add_mode && !_tcsicmp(option_word, _T("First")))
		{
			tvi.hInsertAfter = TVI_FIRST; // For simplicity, the value of "adding" is ignored.
		}
		else if (add_mode && IsNumeric(option_word, false, false, false))
		{
			tvi.hInsertAfter = (HTREEITEM)ATOI64(next_option);
		}
		else
		{
			control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
			return FValueError(ERR_INVALID_OPTION, option_word);
		}
	}

	// For TV.Modify(), an explicit empty string is allowed for aName, which sets it to a blank value.
	// By contrast, if the param is omitted, the name is left unchanged.
	if (aName.has_value())
	{
		tvi.item.pszText = const_cast<LPTSTR>(aName.value_or_empty());
		tvi.item.mask |= TVIF_TEXT;
	}
	if (add_mode) // TV.Add()
	{
		tvi.item.hItem = TreeView_InsertItem(control.hwnd, &tvi); // Update tvi.item.hItem for convenience/maint. It's for use in later sections because retval is overridden to be zero for partial failure in modify-mode.
		aRetVal = (UINT_PTR)tvi.item.hItem; // Set return value.
	}
	else // TV.Modify()
	{
		if (tvi.item.mask != LVIF_STATE || tvi.item.stateMask) // An item's property or one of the state bits needs changing.
			TreeView_SetItem(control.hwnd, &tvi.itemex);
	}

	if (ensure_visible) // Seems best to do this one prior to "select" below.
		SendMessage(control.hwnd, TVM_ENSUREVISIBLE, 0, (LPARAM)tvi.item.hItem); // Return value is ignored in this case, since its definition seems a little weird.
	if (ensure_visible_first) // Seems best to do this one prior to "select" below.
		TreeView_Select(control.hwnd, tvi.item.hItem, TVGN_FIRSTVISIBLE); // Return value is also ignored due to rarity, code size, and because most people wouldn't care about a failure even if for some reason it failed.
	if (select_flag)
		TreeView_Select(control.hwnd, tvi.item.hItem, select_flag);

	control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
	return OK;
}



FResult GuiControlType::TV_Delete(optl<UINT_PTR> aItemID)
{
	// If param #1 is present but is zero, for safety it seems best not to do a delete-all (in case a
	// script bug is so rare that it is never caught until the script is distributed).  Another reason
	// is that a script might do something like TV.Delete(TV.GetSelection()), which would be desired
	// to fail not delete-all if there's ever any way for there to be no selection.
	if (aItemID.has_value() && aItemID.value() == 0)
		return FR_E_ARG(0);
	if (!SendMessage(hwnd, TVM_DELETEITEM, 0, aItemID.value_or(NULL)))
		return FR_E_FAILED; // Invalid parameter?
	return OK;
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



FResult GuiControlType::TV_GetChild(UINT_PTR aItemID, UINT_PTR &aRetVal)
{
	aRetVal = (UINT_PTR)TreeView_GetChild(hwnd, aItemID);
	return OK;
}

FResult GuiControlType::TV_GetCount(UINT &aRetVal)
{
	aRetVal = TreeView_GetCount(hwnd);
	return OK;
}

FResult GuiControlType::TV_GetParent(UINT_PTR aItemID, UINT_PTR &aRetVal)
{
	aRetVal = (UINT_PTR)TreeView_GetParent(hwnd, aItemID);
	return OK;
}

FResult GuiControlType::TV_GetPrev(UINT_PTR aItemID, UINT_PTR &aRetVal)
{
	aRetVal = (UINT_PTR)TreeView_GetPrevSibling(hwnd, aItemID);
	return OK;
}

FResult GuiControlType::TV_GetSelection(UINT_PTR &aRetVal)
{
	aRetVal = (UINT_PTR)TreeView_GetSelection(hwnd);
	return OK;
}



FResult GuiControlType::TV_GetNext(optl<UINT_PTR> aItemID, optl<StrArg> aItemType, UINT_PTR &aRetVal)
{
	HTREEITEM hitem = (HTREEITEM)aItemID.value_or(NULL);
	
	if (!aItemType.has_value())
	{
		aRetVal = (UINT_PTR)SendMessage(hwnd, TVM_GETNEXTITEM, hitem ? TVGN_NEXT : TVGN_ROOT, (LPARAM)hitem);
		// Above: It seems best to treat hitem==0 as "get root", even though it sacrifices some error detection,
		// because not doing so would be inconsistent with the fact that TV.GetNext(0, "Full") does get the root
		// (which needs to be retained to make script loops to traverse entire tree easier).
		return OK;
	}

	// Since above didn't return, this TV.GetNext's 2-parameter mode, which has an expanded scope that includes
	// not just siblings, but also children and parents.  This allows a tree to be traversed from top to bottom
	// without the script having to do something fancy.
	TCHAR first_char_upper = ctoupper(*omit_leading_whitespace(aItemType.value_or_empty()));
	bool search_checkmark;
	if (first_char_upper == 'C')
		search_checkmark = true;
	else if (first_char_upper == 'F')
		search_checkmark = false;
	else // Reserve other option letters/words for future use by being somewhat strict.
		return FR_E_ARG(1);

	// When an actual item was specified, search begins at the item *after* it.  Otherwise (when NULL):
	// It's a special mode that always considers the root node first.  Otherwise, there would be no way
	// to start the search at the very first item in the tree to find out whether it's checked or not.
	hitem = GetNextTreeItem(hwnd, hitem); // Handles the comment above.
	if (!search_checkmark) // Simple tree traversal, so just return the next item (if any).
	{
		aRetVal = (UINT_PTR)hitem; // OK if NULL.
		return OK;
	}

	// Otherwise, search for the next item having a checkmark. For performance, it seems best to assume that
	// the control has the checkbox style (the script would realistically never call it otherwise, so the
	// control's style isn't checked.
	for (; hitem; hitem = GetNextTreeItem(hwnd, hitem))
		if (TreeView_GetCheckState(hwnd, hitem) == 1) // 0 means unchecked, -1 means "no checkbox image".
			break;
	aRetVal = (UINT_PTR)hitem; // OK if NULL.
	return OK;
}



// TV.Get()
// Returns: Varies depending on param #2.
// Parameters:
//    1: HTREEITEM.
//    2: Name of attribute to get.
FResult GuiControlType::TV_Get(UINT_PTR aItemID, StrArg aAttribute, UINT_PTR &aRetVal)
{
	HTREEITEM hitem = (HTREEITEM)aItemID;
	UINT state_mask;
	switch (ctoupper(*omit_leading_whitespace(aAttribute)))
	{
	case 'E': state_mask = TVIS_EXPANDED; break; // Expanded
	case 'C': state_mask = TVIS_STATEIMAGEMASK; break; // Checked
	case 'B': state_mask = TVIS_BOLD; break; // Bold
	//case 'S' for "Selected" is not provided because TV.GetSelection() seems to cover that well enough.
	//case 'P' for "is item a parent?" is not provided because TV.GetChild() seems to cover that well enough.
	// (though it's possible that retrieving TVITEM's cChildren would perform a little better).
	}
	// Below seems to need a bit-AND with state_mask to work properly, at least on XP SP2.  Otherwise,
	// extra bits are present such as 0x2002 for "expanded" when it's supposed to be either 0x00 or 0x20.
	UINT result = state_mask & (UINT)SendMessage(hwnd, TVM_GETITEMSTATE, (WPARAM)hitem, state_mask);
	if (state_mask == TVIS_STATEIMAGEMASK)
	{
		if (result != 0x2000) // It doesn't have a checkmark state image.
			hitem = 0;
	}
	else // For all others, anything non-zero means the flag is present.
		if (!result) // Flag not present.
			hitem = 0;
	aRetVal = (UINT_PTR)hitem;
	return OK;
}



// TV.GetText()
// Returns: Text on success.
// Throws on failure.
// Parameters:
//    1: HTREEITEM.
FResult GuiControlType::TV_GetText(UINT_PTR aItemID, StrRet &aRetVal)
{
	HWND control_hwnd = hwnd;

	TCHAR text_buf[LV_TEXT_BUF_SIZE]; // i.e. uses same size as ListView.
	TVITEM tvi;
	tvi.hItem = (HTREEITEM)aItemID;
	tvi.mask = TVIF_TEXT;
	tvi.pszText = text_buf;
	tvi.cchTextMax = LV_TEXT_BUF_SIZE - 1; // -1 because of nagging doubt about size vs. length. Some MSDN examples subtract one), such as TabCtrl_GetItem()'s cchTextMax.

	if (SendMessage(control_hwnd, TVM_GETITEM, 0, (LPARAM)&tvi))
	{
		// Must use tvi.pszText vs. text_buf because MSDN says: "Applications should not assume that the text will
		// necessarily be placed in the specified buffer. The control may instead change the pszText member
		// of the structure to point to the new text rather than place it in the buffer."
		return aRetVal.Copy(tvi.pszText) ? OK : FR_E_OUTOFMEM;
	}
	else
	{
		// On failure, it seems best to throw an exception.
		return FR_E_FAILED;
	}
}



FResult GuiControlType::TV_SetImageList(UINT_PTR aImageListID, optl<int> aIconType, UINT_PTR &aRetVal)
// Returns (MSDN): "handle to the image list previously associated with the control if successful; NULL otherwise."
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
// 2: Optional: Type of list.
{
	HIMAGELIST himl = (HIMAGELIST)aImageListID;
	int list_type = aIconType.value_or(TVSIL_NORMAL);
	aRetVal = (UINT_PTR)TreeView_SetImageList(hwnd, himl, list_type);
	return OK;
}
