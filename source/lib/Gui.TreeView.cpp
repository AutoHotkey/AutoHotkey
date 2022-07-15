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

#include "script_func_impl.h"



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
	int param3 = ParamIndexToOptionalBOOL(2, FALSE);
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
		if (ParamIndexToBOOL(3)) // A value of True indicates that the image should be scaled to fit the imagelist's image size.
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



////////////////////
// Misc Functions //
////////////////////

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



////////////////////
// Core Functions //
////////////////////


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
	//case SYM_MISSING:
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
	if (aToken.symbol == SYM_VAR && !VARREF_IS_READ(aToken.var_usage)) // VARREF_ISSET is tolerated for use by IsSet().
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
