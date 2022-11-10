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
#include "globaldata.h" // for a lot of things
#include "application.h" // for MsgSleep()
#include "window.h" // for SetForegroundWindowEx()
#include "script_func_impl.h"



ObjectMemberMd UserMenu::sMembers[] =
{
	md_member(UserMenu, Add, CALL, (In_Opt, String, Name), (In_Opt, Object, FunctionOrSubmenu), (In_Opt, String, Options)),
	md_member(UserMenu, AddStandard, CALL, md_arg_none),
	md_member(UserMenu, Check, CALL, (In, String, Item)),
	md_property(UserMenu, ClickCount, Int32),
	md_property(UserMenu, Default, String),
	md_member(UserMenu, Delete, CALL, (In_Opt, String, Item)),
	md_member(UserMenu, Disable, CALL, (In, String, Item)),
	md_member(UserMenu, Enable, CALL, (In, String, Item)),
	md_property_get(UserMenu, Handle, UIntPtr),
	md_member(UserMenu, Insert, CALL, (In_Opt, String, Before), (In_Opt, String, Name), (In_Opt, Object, FunctionOrSubmenu), (In_Opt, String, Options)),
	md_member(UserMenu, Rename, CALL, (In, String, Item), (In_Opt, String, NewName)),
	md_member(UserMenu, SetColor, CALL, (In_Opt, Variant, Color), (In_Opt, Bool32, ApplyToSubmenus)),
	md_member(UserMenu, SetIcon, CALL, (In, String, Item), (In, String, File), (In_Opt, Int32, Number), (In_Opt, Int32, Width)),
	md_member(UserMenu, Show, CALL, (In_Opt, Int32, X), (In_Opt, Int32, Y)),
	md_member(UserMenu, ToggleCheck, CALL, (In, String, Item)),
	md_member(UserMenu, ToggleEnable, CALL, (In, String, Item)),
	md_member(UserMenu, Uncheck, CALL, (In, String, Item)),
};
int UserMenu::sMemberCount = _countof(sMembers);



FResult UserMenu::Add(optl<StrArg> aName, optl<IObject*> aFuncOrSubmenu, optl<StrArg> aOptions)
{
	return Add(aName, aFuncOrSubmenu, aOptions, nullptr);
}



FResult UserMenu::Insert(optl<StrArg> aBefore, optl<StrArg> aName, optl<IObject*> aFuncOrSubmenu, optl<StrArg> aOptions)
{
	UserMenuItem *prev_item = mLastMenuItem; // Default to append.
	if (aBefore.has_nonempty_value()) // i.e. caller specified where to insert.
	{
		bool search_by_pos;
		if (!FindItem(aBefore.value(), prev_item, search_by_pos))
		{
			// The item wasn't found.  Treat it as an error unless it is the position
			// immediately after the last item.
			if (  !(search_by_pos && ATOI(aBefore.value()) == (int)mMenuItemCount + 1)  )
				return ItemNotFoundError(aBefore.value());
		}
	}
	return Add(aName, aFuncOrSubmenu, aOptions
		// To simplify insertion, give AddItem() a pointer to the variable within the
		// linked-list which points to the item, rather than a pointer to the item itself:
		, prev_item ? &prev_item->mNextMenuItem : &mFirstMenuItem);
}



FResult UserMenu::Add(optl<StrArg> aName, optl<IObject*> aFuncOrSubmenu, optl<StrArg> aOptions, UserMenuItem **aInsertAt)
{
	auto options = aOptions.value_or_empty();
	if (aName.is_blank_or_omitted())
	{
		if (aFuncOrSubmenu.has_value() || *options)
			return FR_E_ARGS; // Invalid combination of parameters.
		return AddItem(_T(""), 0, nullptr, nullptr, _T(""), aInsertAt) ? OK : FR_FAIL;
	}
	
	bool search_by_pos = false;
	UserMenuItem *menu_item = nullptr, *menu_item_prev = nullptr;
	if (!aInsertAt) // Add(), not Insert().
	{
		menu_item = FindItem(aName.value(), menu_item_prev, search_by_pos);
		if (!menu_item && search_by_pos)
			return FR_E_ARG(0); // Invalid position.
	}
	
	// Whether an existing menu item's options should be updated without updating its submenu or callback:
	bool update_existing_item_options = (menu_item && !aFuncOrSubmenu.has_value() && *options);
	
	IObject *callback = nullptr;
	UserMenu *submenu = nullptr;
	if (!update_existing_item_options)
	{
		if (!aFuncOrSubmenu.has_value())
			return FR_E_ARG(aInsertAt ? 2 : 1);
		submenu = dynamic_cast<UserMenu *>(aFuncOrSubmenu.value());
		if (submenu) // Param #2 is a Menu object.
		{
			// Before going further: since a submenu has been specified, make sure that the parent
			// menu is not included anywhere in the nested hierarchy of that submenu's submenus.
			// The OS doesn't seem to like that, creating empty or strange menus if it's attempted.
			// Although adding a menu bar as a submenu appears to work, displaying it does strange
			// things to the actual menu bar on any GUI it is assigned to.
			if (submenu == this || submenu->ContainsMenu(this)
				|| submenu->mMenuType != MENU_TYPE_POPUP)
				return FR_E_ARG(aInsertAt ? 2 : 1);
		}
		else 
		{
			// Param #2 is not a submenu.
			callback = aFuncOrSubmenu.value();
			auto fr = ValidateFunctor(callback, 3);
			if (fr != OK)
				return fr;
		}
	}
	
	if (menu_item)
		return ModifyItem(menu_item, callback, submenu, options) ? OK : FR_FAIL;
	else
		return AddItem(aName.value(), 0, callback, submenu, options, aInsertAt) ? OK : FR_FAIL;
}



FResult UserMenu::AddStandard()
{
	return AppendStandardItems() ? OK : FR_FAIL;
}



FResult UserMenu::Delete(optl<StrArg> aItemName)
{
	if (!aItemName.has_value())
	{
		DeleteAllItems();
		return OK;
	}
	if (aItemName.is_blank())
		return FR_E_ARG(0);
	bool search_by_pos = false;
	UserMenuItem *menu_item_prev
		, *menu_item = FindItem(aItemName.value(), menu_item_prev, search_by_pos);
	if (!menu_item)
		return ItemNotFoundError(aItemName.value());
	DeleteItem(menu_item, menu_item_prev);
	return OK;
}



FResult UserMenu::Rename(StrArg aItemName, optl<StrArg> aNewName)
{
	UserMenuItem *item;
	auto fr = GetItem(aItemName, item);
	if (fr != OK)
		return fr;
	auto new_name = aNewName.value_or_empty();
	if (!RenameItem(item, new_name))
		return FError(_T("Rename failed (name too long?)."), new_name);
	return OK;
}



FResult UserMenu::SetColor(ExprTokenType *aColor, optl<BOOL> aApplyToSubmenus)
{
	BOOL submenus = aApplyToSubmenus.value_or(TRUE);
	COLORREF color = CLR_DEFAULT;
	if (aColor && !ColorToBGR(*aColor, color))
		return FR_E_ARG(0);
	SetColor(color, submenus);
	return OK;
}



FResult UserMenu::SetIcon(StrArg aItemName, StrArg aIconFile, optl<int> aIconNumber, optl<int> aIconWidth)
{
	UserMenuItem *item;
	auto fr = GetItem(aItemName, item);
	if (fr != OK)
		return fr;
	// Icon width defaults to system small icon size.  Original icon size will be used if "0" is specified.
	if (!SetItemIcon(item, aIconFile, aIconNumber.value_or(0)
		, aIconWidth.has_value() ? aIconWidth.value() : GetSystemMetrics(SM_CXSMICON)))
		return FError(ERR_LOAD_ICON, aIconFile);
	return OK;
}



FResult UserMenu::Show(optl<int> aX, optl<int> aY)
{
	return Display(true, aX.value_or(COORD_UNSPECIFIED), aY.value_or(COORD_UNSPECIFIED)) ? OK : FR_FAIL;
}



FResult UserMenu::Check(StrArg aItemName)
{
	return SetItemState(aItemName, MFS_CHECKED, MFS_CHECKED);
}

FResult UserMenu::ToggleCheck(StrArg aItemName)
{
	return SetItemState(aItemName, MFS_CHECKED, 0);
}

FResult UserMenu::Uncheck(StrArg aItemName)
{
	return SetItemState(aItemName, 0, MFS_CHECKED);
}

FResult UserMenu::Disable(StrArg aItemName)
{
	return SetItemState(aItemName, MFS_DISABLED, MFS_DISABLED);
}

FResult UserMenu::Enable(StrArg aItemName)
{
	return SetItemState(aItemName, 0, MFS_DISABLED);
}

FResult UserMenu::ToggleEnable(StrArg aItemName)
{
	return SetItemState(aItemName, MFS_DISABLED, 0);
}



FResult UserMenu::get_ClickCount(int &aRetVal)
{
	aRetVal = mClickCount;
	return OK;
}

FResult UserMenu::set_ClickCount(int aValue)
{
	if (aValue < 1 || aValue > 2)
		return FR_E_ARG(0);
	mClickCount = aValue;
	return OK;
}



FResult UserMenu::get_Default(StrRet &aRetVal)
{
	aRetVal.SetTemp(mDefault ? mDefault->mName : _T(""));
	return OK;
}

FResult UserMenu::set_Default(StrArg aItemName)
{
	UserMenuItem *item = nullptr;
	if (*aItemName)
	{
		auto fr = GetItem(aItemName, item);
		if (fr != OK)
			return fr;
	}
	SetDefault(item);
	return OK;
}



FResult UserMenu::get_Handle(UINT_PTR &aRetVal)
{
	if (!mMenu)
		if (!CreateHandle())
			return FR_E_WIN32;
	aRetVal = (UINT_PTR)mMenu;
	return OK;
}



UserMenu *Script::FindMenu(HMENU aMenuHandle)
{
	if (!aMenuHandle) return NULL;
	for (UserMenu *menu = mFirstMenu; menu != NULL; menu = menu->mNextMenu)
		if (menu->mMenu == aMenuHandle)
			return menu;
	return NULL; // No match found.
}



UserMenu::UserMenu(MenuTypeType aMenuType)
	: mMenuType(aMenuType)
{
	SetBase(sPrototype);
	g_script.AddMenu(this);
}

UserMenu *Script::AddMenu(UserMenu *aMenu)
// Returns the newly created UserMenu object.
{
	ASSERT(aMenu);
	if (!mFirstMenu)
		mFirstMenu = mLastMenu = aMenu;
	else
	{
		mLastMenu->mNextMenu = aMenu;
		// This must be done after the above:
		mLastMenu = aMenu;
	}
	++mMenuCount;  // Only after memory has been successfully allocated.
	return aMenu;
}



ResultType Script::ScriptDeleteMenu(UserMenu *aMenu)
// Deletes a UserMenu object and all the UserMenuItem objects that belong to it.
// Any UserMenuItem object that has a submenu attached to it does not result in
// that submenu being deleted, even if no other menus are using that submenu
// (i.e. the user must delete all menus individually).  Any menus which have
// aMenu as one of their submenus will have that menu item deleted from their
// menus to avoid any chance of problems due to non-existent or NULL submenus.
{
	// Delete any other menu's menu item that has aMenu as its attached submenu:
	// This is not done because reference counting ensures that submenus are not
	// deleted before the parent menu, except when the script is exiting.
	//UserMenuItem *mi, *mi_prev, *mi_to_delete;
	//for (UserMenu *m = mFirstMenu; m; m = m->mNextMenu)
	//	if (m != aMenu) // Don't bother with this menu even if it's submenu of itself, since it will be destroyed anyway.
	//		for (mi = m->mFirstMenuItem, mi_prev = NULL; mi;)
	//		{
	//			mi_to_delete = mi;
	//			mi = mi->mNextMenuItem;
	//			if (mi_to_delete->mSubmenu == aMenu)
	//				m->DeleteItem(mi_to_delete, mi_prev);
	//			else
	//				mi_prev = mi_to_delete;
	//		}
	// Remove aMenu from the linked list.
	UserMenu *aMenu_prev;
	if (aMenu == mFirstMenu) // Checked first since it's always true when called by ~Script().
	{
		mFirstMenu = aMenu->mNextMenu; // Can be NULL if the list will now be empty.
		aMenu_prev = NULL;
	}
	else // Find the item that occurs prior to aMenu in the list:
		for (aMenu_prev = mFirstMenu; aMenu_prev; aMenu_prev = aMenu_prev->mNextMenu)
			if (aMenu_prev->mNextMenu == aMenu)
			{
				aMenu_prev->mNextMenu = aMenu->mNextMenu; // Can be NULL if aMenu was the last one.
				break;
			}
	if (aMenu == mLastMenu)
		mLastMenu = aMenu_prev; // Can be NULL if the list will now be empty.
	--mMenuCount;
	// Do this last when its contents are no longer needed.  It will delete all
	// the items in the menu and destroy the OS menu itself:
	aMenu->Dispose();
	return OK;
}



void UserMenu::Dispose()
{
	DestroyHandle();
	DeleteAllItems();
	if (mBrush) // Free the brush used for the menu's background color.
		DeleteObject(mBrush);
}



UserMenu::~UserMenu()
{
	g_script.ScriptDeleteMenu(this);
}



UINT Script::GetFreeMenuItemID()
// Returns an unused menu item ID, or 0 if all IDs are used.
{
	// Need to find a menuID that isn't already in use by one of the other menu items.
	// But also need to conserve menu items since only a relatively small number of IDs is available.
	// Can't simply use ID_USER_FIRST + mMenuItemCount because: 1) There might be more than one
	// user defined menu; 2) a menu item in the middle of the list may have been deleted,
	// in which case that value would already be in use by the last item.
	// Update: Now using caching of last successfully found free-ID to greatly improve avg.
	// performance, especially for menus that contain thousands of items and submenus, such as
	// ones that are built to mirror an entire nested directory structure.  Caching should
	// improve performance even after all menu IDs within the available range have been
	// allocated once (via adding and deleting menus + menu items) since large blocks of free IDs
	// should be free, and on average, the caching will exploit these large free blocks.  However,
	// if large amounts of menus and menu items are continually deleted and re-added by a script,
	// the pool of free IDs will become fragmented over time, which will reduce performance.
	// Since that kind of script behavior seems very rare, no attempt is made to "defragment".
	// If more performance is needed in the future (seems unlikely for 99.9999% of scripts),
	// could maintain an field of ~64000 bits, each bit representing whether a menu item ID is
	// free.  Then, every time a menu or one or more of its IDs is deleted or added, the corresponding
	// ID could be marked as free/taken.  That would add quite a bit of complexity to the menu
	// delete code, however, and it would reduce the overall maintainability.  So it definitely
	// doesn't seem worth it, especially since Windows XP seems to have trouble even displaying
	// menus larger than around 15000-25000 items.
	static UINT sLastFreeID = ID_USER_FIRST - 1;
	// Increment by one for each new search, both due to the above line and because the
	// last-found free ID has a high likelihood of still being in use:
	++sLastFreeID;
	bool id_in_use;
	// Note that the i variable is used to force the loop to complete exactly one full
	// circuit through all available IDs, regardless of where the starting/cached value:
	for (int i = 0; i < (ID_USER_LAST - ID_USER_FIRST + 1); ++i, ++sLastFreeID) // FOR EACH ID
	{
		if (sLastFreeID > ID_USER_LAST)
			sLastFreeID = ID_USER_FIRST;  // Wrap around to the beginning so that one complete circuit is made.
		id_in_use = false;  // Reset the default each iteration (overridden if the below finds a match).
		for (UserMenu *m = mFirstMenu; m; m = m->mNextMenu) // FOR EACH MENU
		{
			for (UserMenuItem *mi = m->mFirstMenuItem; mi; mi = mi->mNextMenuItem) // FOR EACH MENU ITEM
			{
				if (mi->mMenuID == sLastFreeID)
				{
					id_in_use = true;
					break;
				}
			}
			if (id_in_use) // No point in searching the other menus, since it's now known to be in use.
				break;
		}
		if (!id_in_use) // Break before the loop increments sLastFreeID.
			break;
	}
	return id_in_use ? 0 : sLastFreeID;
}



FResult UserMenu::GetItem(LPCTSTR aNameOrPos, UserMenuItem *&aItem)
{
	bool bypos;
	UserMenuItem *prev;
	aItem = FindItem(aNameOrPos, prev, bypos);
	if (aItem)
		return OK;
	return ItemNotFoundError(aNameOrPos);
}



FResult UserMenu::ItemNotFoundError(LPCTSTR aItem)
{
	return FError(ERR_INVALID_MENU_ITEM, aItem, ErrorPrototype::Target);
}



UserMenuItem *UserMenu::FindItem(LPCTSTR aNameOrPos, UserMenuItem *&aPrevItem, bool &aByPos)
{
	int index_to_find = -1;
	size_t length = _tcslen(aNameOrPos);
	// Check if the caller identified the menu item by position/index rather than by name.
	// This should be reasonably backwards-compatible, as any scripts that want literally
	// "1&" as menu item text would have to actually write "1&&".
	if (length > 1
		&& aNameOrPos[length - 1] == '&' // Use the same convention as MenuSelect: 1&, 2&, 3&...
		&& aNameOrPos[length - 2] != '&') // Not &&, which means one literal &.
		index_to_find = ATOI(aNameOrPos) - 1; // Yields -1 if aParam3 doesn't start with a number.
	aByPos = index_to_find > -1;
	// Find the item.
	int current_index = 0;
	UserMenuItem *menu_item_prev = NULL, *menu_item;
	for (menu_item = mFirstMenuItem
		; menu_item
		; menu_item_prev = menu_item, menu_item = menu_item->mNextMenuItem, ++current_index)
		if (current_index == index_to_find // Found by index.
			|| !lstrcmpi(menu_item->mName, aNameOrPos)) // Found by case-insensitive text match.
			break;
	aPrevItem = menu_item_prev;
	return menu_item;
}



UserMenuItem *UserMenu::FindItemByID(UINT aID)
{
	for (UserMenuItem *mi = mFirstMenuItem; mi; mi = mi->mNextMenuItem)
		if (mi->mMenuID == aID)
			return mi;
	return NULL;
}



// Macros for use with the below methods (in previous versions, submenus were identified by position):
#define aMenuItem_ID		aMenuItem->mMenuID
#define aMenuItem_MF_BY		MF_BYCOMMAND
#define UPDATE_GUI_MENU_BARS(menu_type, hmenu) \
	if (menu_type == MENU_TYPE_BAR && g_firstGui)\
		GuiType::UpdateMenuBars(hmenu); // Above: If it's not a popup, it's probably a menu bar.



ResultType UserMenu::AddItem(LPCTSTR aName, UINT aMenuID, IObject *aCallback, UserMenu *aSubmenu, LPCTSTR aOptions
	, UserMenuItem **aInsertAt)
// Caller must have already ensured that aName does not yet exist as a user-defined menu item
// in this->mMenu.
{
	size_t length = _tcslen(aName);
	if (length > MAX_MENU_NAME_LENGTH)
		return g_script.RuntimeError(_T("Menu item name too long."), aName);
	// If caller didn't specify a (built-in) item ID, get the next free ID.
	// Even separators get an ID, so that they can be modified later using the position& notation.
	if (  !aMenuID && !(aMenuID = g_script.GetFreeMenuItemID())  )
		return g_script.RuntimeError(_T("Too many menu items.")); // All ~64000 IDs are in use!
	// After mem is allocated, the object takes charge of its later deletion:
	LPTSTR name_dynamic;
	if (length)
	{
		if (   !(name_dynamic = tmalloc(length + 1))   )  // +1 for terminator.
			return MemoryError();
		_tcscpy(name_dynamic, aName);
	}
	else
		name_dynamic = Var::sEmptyString; // So that it can be detected as a non-allocated empty string.
	UserMenuItem *menu_item = new UserMenuItem(name_dynamic, length + 1, aMenuID, aCallback, aSubmenu, this);
	if (*aOptions && !UpdateOptions(menu_item, aOptions))
	{
		// Invalid options; an error message was displayed.
		delete menu_item; 
		return FAIL;
	}
	if (mMenu)
	{
		InternalAppendMenu(menu_item, aInsertAt ? *aInsertAt : NULL);
		UPDATE_GUI_MENU_BARS(mMenuType, mMenu)
	}
	if (aInsertAt)
	{
		// Caller has passed a pointer to the variable in the linked list which should
		// hold this new item; either &mFirstMenuItem or &previous_item->mNextMenuItem.
		menu_item->mNextMenuItem = *aInsertAt;
		// This must be done after the above:
		*aInsertAt = menu_item;
	}
	else
	{
		// Append the item.
		if (!mFirstMenuItem)
			mFirstMenuItem = menu_item;
		else
			mLastMenuItem->mNextMenuItem = menu_item;
		// This must be done after the above:
		mLastMenuItem = menu_item;
	}
	++mMenuItemCount;  // Only after memory has been successfully allocated.
	if (_tcschr(aName, '\t')) // v1.1.04: The new item has a keyboard accelerator.
		UpdateAccelerators();
	return OK;
}



ResultType UserMenu::InternalAppendMenu(UserMenuItem *mi, UserMenuItem *aInsertBefore)
// Appends an item to mMenu and and ensures the new item's ID is set.
{
	MENUITEMINFO mii;
	mii.cbSize = sizeof(mii);
	mii.fMask = MIIM_ID | MIIM_FTYPE | MIIM_STRING | MIIM_STATE;
	mii.wID = mi->mMenuID;
	mii.fType = mi->mMenuType;
	mii.dwTypeData = mi->mName;
	mii.fState = mi->mMenuState;
	if (mi->mSubmenu)
	{
		// Ensure submenu is created so that its handle can be used below.
		if (!mi->mSubmenu->CreateHandle())
			return FAIL;
		mii.fMask |= MIIM_SUBMENU;
		mii.hSubMenu = mi->mSubmenu->mMenu;
	}
	if (mi->mBitmap)
	{
		mii.fMask |= MIIM_BITMAP;
		mii.hbmpItem = mi->mBitmap;
	}
	UINT insert_at;
	BOOL by_position;
	if (aInsertBefore)
		insert_at = aInsertBefore->mMenuID, by_position = FALSE;
	else
		insert_at = GetMenuItemCount(mMenu), by_position = TRUE;
	// Although AppendMenu() ignores the ID when adding a separator and provides no way to
	// specify the ID when adding a submenu, that is purely a limitation of AppendMenu().
	// Using InsertMenuItem() instead allows us to always set the ID, which simplifies
	// identifying separator and submenu items later on.
	return InsertMenuItem(mMenu, insert_at, by_position, &mii) ? OK : FAIL;
}



UserMenuItem::UserMenuItem(LPTSTR aName, size_t aNameCapacity, UINT aMenuID, IObject *aCallback, UserMenu *aSubmenu, UserMenu *aMenu)
// UserMenuItem Constructor.
	: mName(aName), mNameCapacity(aNameCapacity), mMenuID(aMenuID), mCallback(aCallback), mSubmenu(aSubmenu), mMenu(aMenu)
	, mPriority(0) // default priority = 0
	, mMenuState(MFS_ENABLED | MFS_UNCHECKED), mMenuType(*aName ? MFT_STRING : MFT_SEPARATOR)
	, mNextMenuItem(NULL)
	, mBitmap(NULL) // L17: Initialize mIcon/mBitmap union.
{
	if (aSubmenu)
		aSubmenu->AddRef();
}



void UserMenu::DeleteItem(UserMenuItem *aMenuItem, UserMenuItem *aMenuItemPrev, bool aUpdateGuiMenuBars)
{
	// Remove this menu item from the linked list:
	if (aMenuItem == mLastMenuItem)
		mLastMenuItem = aMenuItemPrev; // Can be NULL if the list will now be empty.
	if (aMenuItemPrev) // there is another item prior to aMenuItem in the linked list.
		aMenuItemPrev->mNextMenuItem = aMenuItem->mNextMenuItem; // Can be NULL if aMenuItem was the last one.
	else // aMenuItem was the first one in the list.
		mFirstMenuItem = aMenuItem->mNextMenuItem; // Can be NULL if the list will now be empty.
	if (mMenu) // Delete the item from the menu.
		RemoveMenu(mMenu, aMenuItem_ID, aMenuItem_MF_BY); // v1.0.48: Lexikos: DeleteMenu() destroys any sub-menu handle associated with the item, so use RemoveMenu. Otherwise the submenu handle stored somewhere else in memory would suddenly become invalid.
	RemoveItemIcon(aMenuItem); // L17: Free icon or bitmap.
	if (mDefault == aMenuItem)
		mDefault = NULL; // No need to SetDefaultMenuItem since the previous default item was removed.
	delete aMenuItem; // Do this last when its contents are no longer needed.
	--mMenuItemCount;
	if (aUpdateGuiMenuBars)
		UPDATE_GUI_MENU_BARS(mMenuType, mMenu)  // Verified as being necessary.
}



void UserMenu::DeleteAllItems()
// Remove all menu items from the linked list and from the menu.
{
	// Fixed for v1.1.27.03: Don't attempt to take a shortcut by calling DestroyHandle(), as it
	// will fail if this is a sub-menu of a menu bar.  Removing the items individually will
	// do exactly what the user expects.  The following old comment indicates one reason
	// DestroyHandle() was used; that reason is now obsolete since submenus are given IDs:
	// "In addition, this avoids the need to find any submenus by position:"
	if (!mFirstMenuItem)
		return;  // If there are no user-defined menu items, it's already in the correct state.
	UserMenuItem *menu_item_to_delete;
	for (UserMenuItem *mi = mFirstMenuItem; mi;)
	{
		if (mMenu)
			RemoveMenu(mMenu, mi->mMenuID, MF_BYCOMMAND);
		menu_item_to_delete = mi;
		mi = mi->mNextMenuItem;
		RemoveItemIcon(menu_item_to_delete); // L26: Free icon or bitmap!
		delete menu_item_to_delete;
	}
	mFirstMenuItem = mLastMenuItem = NULL;
	mMenuItemCount = 0;
	mDefault = NULL;  // i.e. there can't be a *user-defined* default item anymore, even if this is the tray.
	UPDATE_GUI_MENU_BARS(mMenuType, mMenu)  // Verified as being necessary.
}



UserMenuItem::~UserMenuItem()
{
	if (mName != Var::sEmptyString)
		free(mName);
	// mCallback is handled by ~IObjectRef().
	if (mSubmenu)
		mSubmenu->Release();
}



ResultType UserMenu::ModifyItem(UserMenuItem *aMenuItem, IObject *aCallback, UserMenu *aSubmenu, LPCTSTR aOptions)
// Modify the callback, submenu, or options of a menu item (exactly one of these should be NULL and the
// other not except when updating only the options).
// If a menu item becomes a submenu, we don't relinquish its ID in case it's ever made a normal item
// again (avoids the need to re-lookup a unique ID).
{
	if (*aOptions && UpdateOptions(aMenuItem, aOptions) != OK)
		return FAIL; // UpdateOptions displays an error message if aOptions contains invalid options.
	if (!aCallback && !aSubmenu) // We were called only to update this item's options.
		return OK;

	if (aMenuItem->mMenuID >= ID_TRAY_FIRST && aCallback)
	{
		// For a custom label to work for this item, it must have a different ID.
		MENUITEMINFO mii;
		mii.cbSize = sizeof(mii);
		mii.fMask = MIIM_ID;
		mii.wID = g_script.GetFreeMenuItemID();
		if (mMenu)
			SetMenuItemInfo(mMenu, aMenuItem->mMenuID, FALSE, &mii);
		aMenuItem->mMenuID = mii.wID;
	}

	aMenuItem->mCallback = aCallback;  // This will be NULL if this menu item is a separator or submenu.
	if (aMenuItem->mSubmenu == aSubmenu) // Below relies on this check.
		return OK;

	if (aSubmenu)
		aSubmenu->AddRef();
	if (aMenuItem->mSubmenu)
		aMenuItem->mSubmenu->Release();
	aMenuItem->mSubmenu = aSubmenu;

	if (!mMenu)
		return OK;
	// Otherwise, since the OS menu exists, one of these is to be done to aMenuItem in it:
	// 1) Change a submenu to point to a different menu.
	// 2) Change a submenu so that it becomes a normal menu item.
	// 3) Change a normal menu item into a submenu.
	// Replacing or removing a submenu is known to destroy the old submenu as a side-effect,
	// so instead of changing the submenu directly, remove the item and add it back with the
	// correct properties.
	RemoveMenu(mMenu, aMenuItem->mMenuID, MF_BYCOMMAND);
	InternalAppendMenu(aMenuItem, aMenuItem->mNextMenuItem);
	return OK;
}



ResultType UserMenu::UpdateOptions(UserMenuItem *aMenuItem, LPCTSTR aOptions)
{
	UINT new_type = aMenuItem->mMenuType; // Set default.

	TCHAR option_word[16]; // Enough for any single option word or number, with room to avoid false positives due to truncation.
	LPCTSTR next_option, option_end;
	bool adding;

	// See GuiType::ControlParseOptions() for comments about how the options are parsed.
	for (next_option = aOptions; *next_option; next_option = option_end)
	{
		next_option = omit_leading_whitespace(next_option);
		if (*next_option == '-')
		{
			adding = false;
			++next_option;
		}
		else
		{
			adding = true;
			if (*next_option == '+')
				++next_option;
		}

		if (!*next_option)
			break;
		for (option_end = next_option; *option_end && !IS_SPACE_OR_TAB(*option_end); ++option_end);
		if (option_end == next_option)
			continue;
		
		tcslcpy(option_word, next_option, min((option_end - next_option) + 1, _countof(option_word)));

		// End generic option-parsing code; begin menu options.
		if (!_tcsicmp(option_word, _T("Radio"))) if (adding) new_type |= MFT_RADIOCHECK; else new_type &= ~MFT_RADIOCHECK;
		else if (mMenuType == MENU_TYPE_BAR && !_tcsicmp(option_word, _T("Right"))) if (adding) new_type |= MFT_RIGHTJUSTIFY; else new_type &= ~MFT_RIGHTJUSTIFY;
		else if (!_tcsicmp(option_word, _T("Break"))) if (adding) new_type |= MFT_MENUBREAK; else new_type &= ~MFT_MENUBREAK;
		else if (!_tcsicmp(option_word, _T("BarBreak"))) if (adding) new_type |= MFT_MENUBARBREAK; else new_type &= ~MFT_MENUBARBREAK;
		else if (ctoupper(*option_word) == 'P')
			aMenuItem->mPriority = ATOI(option_word + 1);	// invalid priority options are not detected, due to rarity and for brevity. Hence, eg Pxyz, is equivalent to P0.
		else
		{
			if (!ValueError(ERR_INVALID_OPTION, option_word, FAIL_OR_OK)) // invalid option
				return FAIL;
			// Otherwise, user wants to continue.
		}
	}

	if (new_type != aMenuItem->mMenuType)
	{
		if (mMenu)
		{
			MENUITEMINFO mii;
			mii.cbSize = sizeof(mii);
			mii.fMask = MIIM_FTYPE;
			mii.fType = new_type;
			SetMenuItemInfo(mMenu, aMenuItem->mMenuID, FALSE, &mii);
		}
		aMenuItem->mMenuType = (WORD)new_type;
	}
	return OK;
}



ResultType UserMenu::RenameItem(UserMenuItem *aMenuItem, LPCTSTR aNewName)
// Caller should specify "" for aNewName to convert aMenuItem into a separator.
// Returns FAIL if the new name conflicts with an existing name.
{
	if (_tcslen(aNewName) > MAX_MENU_NAME_LENGTH)
		return FAIL; // Caller should display error if desired.

	// Preserve any additional type flags set by options, but exclude the main type bits.
	// Also clear MFT_OWNERDRAW (if set by the script), since it changes the meaning of dwTypeData.
	// MSDN: "The MFT_BITMAP, MFT_SEPARATOR, and MFT_STRING values cannot be combined with one another."
	UINT new_type = (aMenuItem->mMenuType & ~(MFT_BITMAP | MFT_SEPARATOR | MFT_STRING | MFT_OWNERDRAW))
		| (*aNewName ? MFT_STRING : MFT_SEPARATOR);

	if (!mMenu) // Just update the member variables for later use when the menu is created.
	{
		aMenuItem->mMenuType = (WORD)new_type;
		return UpdateName(aMenuItem, aNewName);
	}

	MENUITEMINFO mii;
	mii.cbSize = sizeof(mii);
	mii.fMask = 0;

	if (*aNewName)
	{
		if (aMenuItem->mMenuType & MFT_SEPARATOR)
		{
			// Since this item is currently a separator, the system will have disabled it.
			// Set the item's state to what it should be:
			mii.fMask |= MIIM_STATE;
			mii.fState = aMenuItem->mMenuState;
		}
	}
	else // converting into a separator
	{
		// Testing shows that if an item is converted into a separator and back into a
		// normal item, it retains its submenu.  So don't set the submenu to NULL, since
		// it's not necessary and would result in the OS destroying the submenu:
		//if (aMenuItem->mSubmenu)  // Converting submenu into a separator.
		//{
		//	mii.fMask |= MIIM_SUBMENU;
		//	mii.hSubMenu = NULL;
		//}
	}

	mii.fMask |= MIIM_TYPE;
	mii.fType = new_type;
	mii.dwTypeData = const_cast<LPTSTR>(aNewName);

	// v1.1.04: If the new and old names both have accelerators, call UpdateAccelerators() if they
	// are different. Otherwise call it if only one is NULL (i.e. accelerator was added or removed).
	LPCTSTR old_accel = _tcschr(aMenuItem->mName, '\t'), new_accel = _tcschr(aNewName, '\t');
	bool update_accel = old_accel && new_accel ? _tcsicmp(old_accel, new_accel) : old_accel != new_accel;

	// Failure is rare enough in the below that no attempt is made to undo the above:
	BOOL result = SetMenuItemInfo(mMenu, aMenuItem->mMenuID, FALSE, &mii);
	UPDATE_GUI_MENU_BARS(mMenuType, mMenu)  // Verified as being necessary.
	if (  !(result && UpdateName(aMenuItem, aNewName))  )
		return FAIL;
	aMenuItem->mMenuType = (WORD)mii.fType; // Update this in case the menu is destroyed/recreated.
	if (update_accel) // v1.1.04: Above determined this item's accelerator was changed.
		UpdateAccelerators(); // Must not be done until after mName is updated.
	if (*aNewName)
		ApplyItemIcon(aMenuItem); // If any.  Simpler to call this than combine it into the logic above.
	return OK;
}



ResultType UserMenu::UpdateName(UserMenuItem *aMenuItem, LPCTSTR aNewName)
// Caller should already have ensured that aMenuItem is not too long.
{
	size_t new_length = _tcslen(aNewName);
	if (new_length)
	{
		if (new_length >= aMenuItem->mNameCapacity) // Too small, so reallocate.
		{
			// Use a temp var. so that mName will never wind up being NULL (relied on by other things).
			// This also retains the original menu name if the allocation fails:
			LPTSTR temp = tmalloc(new_length + 1);  // +1 for terminator.
			if (!temp)
				return FAIL;
			// Otherwise:
			if (aMenuItem->mName != Var::sEmptyString) // Since it was previously allocated, free it.
				free(aMenuItem->mName);
			aMenuItem->mName = temp;
			aMenuItem->mNameCapacity = new_length + 1;
		}
		_tcscpy(aMenuItem->mName, aNewName);
	}
	else // It will become a separator.
	{
		*aMenuItem->mName = '\0'; // Safe because even if it's capacity is 1 byte, it's a writable byte.
	}
	return OK;
}



void UserMenu::SetItemState(UserMenuItem *aMenuItem, UINT aState, UINT aStateMask)
{
	if (mMenu)
	{
		MENUITEMINFO mii;
		mii.cbSize = sizeof(mii);
		mii.fMask = MIIM_STATE;
		// Retrieve the current state from the menu rather than using mMenuState,
		// in case the script has modified the state via DllCall.
		if (GetMenuItemInfo(mMenu, aMenuItem->mMenuID, FALSE, &mii))
		{
			mii.fState = (mii.fState & ~aStateMask) ^ aState;
			// Update our state in case the menu gets destroyed/recreated.
			aMenuItem->mMenuState = (WORD)mii.fState;
			// Set the new state.
			SetMenuItemInfo(mMenu, aMenuItem->mMenuID, FALSE, &mii);
			if (aStateMask & MFS_DISABLED) // i.e. enabling or disabling, which would affect a menu bar.
				UPDATE_GUI_MENU_BARS(mMenuType, mMenu)  // Verified as being necessary.
			return;
		}
	}
	aMenuItem->mMenuState = (WORD)((aMenuItem->mMenuState & ~aStateMask) ^ aState);
}



FResult UserMenu::SetItemState(StrArg aItemName, UINT aState, UINT aStateMask)
{
	UserMenuItem *item;
	auto fr = GetItem(aItemName, item);
	if (fr != OK)
		return fr;
	SetItemState(item, aState, aStateMask);
	return OK;
}



void UserMenu::CheckItem(UserMenuItem *aMenuItem)
{
	return SetItemState(aMenuItem, MFS_CHECKED, MFS_CHECKED);
}

void UserMenu::UncheckItem(UserMenuItem *aMenuItem)
{
	return SetItemState(aMenuItem, MFS_UNCHECKED, MFS_CHECKED);
}

void UserMenu::ToggleCheckItem(UserMenuItem *aMenuItem)
{
	return SetItemState(aMenuItem, (aMenuItem->mMenuState & MFS_CHECKED) ^ MFS_CHECKED, MFS_CHECKED);
}

void UserMenu::EnableItem(UserMenuItem *aMenuItem)
{
	return SetItemState(aMenuItem, MFS_ENABLED, MFS_DISABLED);
}

void UserMenu::DisableItem(UserMenuItem *aMenuItem)
{
	return SetItemState(aMenuItem, MFS_DISABLED, MFS_DISABLED);
}

void UserMenu::ToggleEnableItem(UserMenuItem *aMenuItem)
{
	return SetItemState(aMenuItem, (aMenuItem->mMenuState & MFS_DISABLED) ^ MFS_DISABLED, MFS_DISABLED);
}



void UserMenu::SetDefault(UserMenuItem *aMenuItem, bool aUpdateGuiMenuBars)
{
	if (mDefault == aMenuItem)
		return;
	mDefault = aMenuItem;
	if (!mMenu) // No further action required: the new setting will be in effect when the menu is created.
		return;
	if (aMenuItem) // A user-defined menu item is being made the default.
		SetMenuDefaultItem(mMenu, aMenuItem->mMenuID, FALSE); // This also ensures that only one is default at a time.
	else
		SetMenuDefaultItem(mMenu, -1, FALSE);
	if (aUpdateGuiMenuBars)
		UPDATE_GUI_MENU_BARS(mMenuType, mMenu)  // Testing shows that menu bars themselves can have default items, and that this is necessary.
}



ResultType UserMenu::CreateHandle()
// Menu bars require non-popup menus (CreateMenu vs. CreatePopupMenu).  Rather than maintain two
// different types of HMENUs on the rare chance that a script might try to use a menu both as
// a popup and a menu bar, it seems best to have only one type to keep the code simple and reduce
// resources used for the menu.  This is reflected in MenuCreate() and MenuBarCreate(), which
// create the UserMenu with the appropriate value for mMenuType.
{
	if (mMenu)
		return OK;
	if (   !(mMenu = (mMenuType == MENU_TYPE_BAR) ? CreateMenu() : CreatePopupMenu())   )
		// Failure is rare, so no error msg here (caller can, if it wants).
		return FAIL;

	// It seems best not to have a mandatory EXIT item added to the bottom of the tray menu
	// because omitting it allows the tray icon to be shown even when the user wants it to
	// have no menu at all (i.e. avoids the need for #NoTrayIcon just to disable the showing
	// of the menu).  If it was done, it should not be done here because calling Menu.Handle
	// prior to adding items would have the effect of inserting EXIT at the top.
	//if (!mMenuItemCount)
	//{
	//	AppendMenu(mTrayMenu->mMenu, MF_STRING, ID_TRAY_EXIT, "E&xit");
	//	return OK;
	//}

	// Now append all of the user defined items:
	UserMenuItem *mi;
	for (mi = mFirstMenuItem; mi; mi = mi->mNextMenuItem)
		InternalAppendMenu(mi);

	if (mDefault)
		// This also automatically ensures that only one is default at a time:
		SetMenuDefaultItem(mMenu, mDefault->mMenuID, FALSE);

	// Apply background color if this menu has a non-standard one.  If this menu has submenus,
	// they will be individually given their own background color when created via CreateHandle(),
	// which is why false is passed:
	ApplyColor(false);

	// L17: Apply default style to merge checkmark/icon columns in menu.
	MENUINFO menu_info;
	menu_info.cbSize = sizeof(MENUINFO);
	menu_info.fMask = MIM_STYLE;
	menu_info.dwStyle = MNS_CHECKORBMP;
	SetMenuInfo(mMenu, &menu_info);

	return OK;
}



void UserMenu::SetColor(COLORREF aColor, bool aApplyToSubmenus)
{
	// AssignColor() takes care of deleting old brush, etc.
	AssignColor(aColor, mColor, mBrush);
	// To avoid complications, such as a submenu being detached from its parent and then its parent
	// later being deleted (which causes the HBRUSH to get deleted too), give each submenu it's
	// own HBRUSH handle by calling SetColor() for each:
	if (aApplyToSubmenus)
		for (UserMenuItem *mi = mFirstMenuItem; mi; mi = mi->mNextMenuItem)
			if (mi->mSubmenu)
				mi->mSubmenu->SetColor(aColor, aApplyToSubmenus);
	if (mMenu)
	{
		ApplyColor(aApplyToSubmenus);
		UPDATE_GUI_MENU_BARS(mMenuType, mMenu)  // Verified as being necessary.
	}
}



void UserMenu::ApplyColor(bool aApplyToSubmenus)
// Caller has ensured that mMenu is not NULL.
// The below should be done even if the default color is being (re)applied because
// testing shows that the OS sets the color to white if the HBRUSH becomes invalid.
// The caller is also responsible for calling UPDATE_GUI_MENU_BARS if desired.
{
	MENUINFO mi = {0}; 
	mi.cbSize = sizeof(MENUINFO);
	mi.fMask = MIM_BACKGROUND|(aApplyToSubmenus ? MIM_APPLYTOSUBMENUS : 0);
	mi.hbrBack = mBrush;
	SetMenuInfo(mMenu, &mi);
}



ResultType UserMenu::AppendStandardItems()
// Caller must ensure that this->mMenu exists if it wants the items to be added immediately.
{
	struct StandardItem
	{
		LPTSTR name;
		UINT id;
	};
	static StandardItem sItems[] =
	{
		_T("&Open"), ID_TRAY_OPEN,
#ifndef AUTOHOTKEYSC
		_T("&Help"), ID_TRAY_HELP,
		_T(""), ID_TRAY_SEP1,
		_T("&Window Spy"), ID_TRAY_WINDOWSPY,
		_T("&Reload Script"), ID_TRAY_RELOADSCRIPT,
		_T("&Edit Script"), ID_TRAY_EDITSCRIPT,
		_T(""), ID_TRAY_SEP2,
#endif
		_T("&Suspend Hotkeys"), ID_TRAY_SUSPEND,
		_T("&Pause Script"), ID_TRAY_PAUSE,
		_T("E&xit"), ID_TRAY_EXIT
	};
	UserMenuItem *&first_new_item = mLastMenuItem ? mLastMenuItem->mNextMenuItem : mFirstMenuItem;
	int i = g_AllowMainWindow ? 0 : 1;
	for (; i < _countof(sItems); ++i)
	{
#ifndef AUTOHOTKEYSC
		if (i == 1 && g_script.mKind == Script::ScriptKindResource)
			i += 6;
		// To minimize code size, and to make it easier for a stdin script to imitate a standard script,
		// the Reload and Edit options are not disabled or removed here for stdin scripts.
		//else if (i == 4 && g_script.mKind == Script::ScriptKindStdIn)
		//	i += 3;
#endif
		if (!FindItemByID(sItems[i].id)) // Avoid duplicating items, but add any missing ones.
			if (!AddItem(sItems[i].name, sItems[i].id, NULL, NULL, _T(""), NULL))
				return FAIL;
	}
	if (this == g_script.mTrayMenu && !mDefault && first_new_item
		&& first_new_item->mMenuID == ID_TRAY_OPEN)
	{
		// No user-defined default menu item, so use the standard one.
		SetDefault(first_new_item, false);
	}
	UPDATE_GUI_MENU_BARS(mMenuType, mMenu)  // Verified as being necessary (though it would be rare anyone would want a menu bar to contain the std items).
	return OK;  // For caller convenience.
}



ResultType UserMenu::EnableStandardOpenItem(bool aEnable)
{
	for (UserMenuItem *mi_prev = NULL, *mi = mFirstMenuItem; mi; mi_prev = mi, mi = mi->mNextMenuItem)
	{
		if (mi->mMenuID >= ID_TRAY_FIRST)
		{
			bool is_enabled = mi->mMenuID == ID_TRAY_OPEN;
			if (is_enabled != aEnable)
			{
				if (aEnable)
				{
					UserMenuItem **p = mi_prev ? &mi_prev->mNextMenuItem : &mFirstMenuItem;
					if (!AddItem(_T("&Open"), ID_TRAY_OPEN, NULL, NULL, _T(""), p))
						return FAIL;
					if (this == g_script.mTrayMenu && !mDefault)
						SetDefault(*p);
					return OK;
				}
				DeleteItem(mi, mi_prev);
				return OK;
			}
			break;
		}
	}
	return OK;
}



void UserMenu::DestroyHandle()
// Destroys the Win32 menu or marks it NULL if it has already been destroyed externally.
// This should be called only when the UserMenu is being deleted (or the script is exiting),
// otherwise any parent menus would still refer to the old Win32 menu.  If the UserMenu is
// being deleted, that implies that its reference count is zero, meaning it isn't in use as
// a submenu or menu bar.
{
	if (!mMenu)  // For performance.
		return;
	// Testing on Windows 10 shows that DestroyMenu() is able to destroy a menu even while it is
	// being displayed, causing only temporary cosmetic issues.  Previous testing with menu bars
	// showed similar results.  There wouldn't be much point in checking the following because the
	// script is in an uninterruptible state while the menu is displayed, which in addition to
	// pausing the current thread (which happens anyway), no new threads can be launched:
	//if (g_MenuIsVisible)
	//	return FAIL;

	// MSDN: "DestroyMenu is recursive, that is, it will destroy the menu and all its submenus."
	// Therefore remove all submenus before destroying the menu, in case the script is using them.
	// If the script is not using them, they will be destroyed automatically as they are released
	// by ~UserMenuItem().
	for (UserMenuItem *mi = mFirstMenuItem; mi ; mi = mi->mNextMenuItem)
		if (mi->mSubmenu)
			RemoveMenu(mMenu, mi->mMenuID, MF_BYCOMMAND);

	DestroyMenu(mMenu);
	mMenu = NULL;
}



ResultType UserMenu::Display(bool aForceToForeground, int aX, int aY)
// aForceToForeground defaults to true because when a menu is displayed spontaneously rather than
// in response to the user right-clicking the tray icon, I believe that the OS will revert to its
// behavior of "resisting" a window that tries to "steal focus".  I believe this resistance does
// not occur when the user clicks the icon because that click causes the task bar to get focus,
// and it is likely that the OS allows other windows to steal focus from the task bar without
// resistance.  This is done because if the main window is *not* successfully activated prior to
// displaying the menu, it might be impossible to dismiss the menu by clicking outside of it.
{
	if (mMenuType != MENU_TYPE_POPUP)
		return g_script.RuntimeError(ERR_INVALID_MENU_TYPE);
	if (!mMenuItemCount)
		return OK;  // Consider the display of an empty menu to be a success.
	if (!CreateHandle()) // Create if needed.  No error msg since so rare.
		return FAIL;
	//if (!IsMenu(mMenu))
	//	mMenu = NULL;
	if (this == g_script.mTrayMenu)
	{
		// These are okay even if the menu items don't exist (perhaps because the user customized the menu):
		CheckMenuItem(mMenu, ID_TRAY_SUSPEND, g_IsSuspended ? MF_CHECKED : MF_UNCHECKED);
		CheckMenuItem(mMenu, ID_TRAY_PAUSE, g->IsPaused ? MF_CHECKED : MF_UNCHECKED);
	}

	POINT pt;
	if (aX == COORD_UNSPECIFIED || aY == COORD_UNSPECIFIED)
		GetCursorPos(&pt);
	if (!(aX == COORD_UNSPECIFIED && aY == COORD_UNSPECIFIED)) // At least one was specified.
	{
		// If one coordinate was omitted, pt contains the cursor position in SCREEN COORDINATES.
		// So don't do something like this, which would incorrectly offset the cursor position
		// by the window position if CoordMode != Screen:
		//CoordToScreen(pt, COORD_MODE_MENU);
		POINT origin = {0};
		CoordToScreen(origin, COORD_MODE_MENU);
		if (aX != COORD_UNSPECIFIED)
			pt.x = aX + origin.x;
		if (aY != COORD_UNSPECIFIED)
			pt.y = aY + origin.y;
	}

	// UPDATE: For v1.0.35.14, must ensure one of the script's windows is active before showing the menu
	// because otherwise the menu cannot be dismissed via the escape key or by clicking outside the menu.
	// Testing shows that ensuring any of our thread's windows is active allows both the tray menu and
	// any popup or context menus to work correctly.
	// UPDATE: For v1.0.35.12, the script's main window (g_hWnd) is activated only for the tray menu because:
	// 1) Doing so for GUI context menus seems to prevent mouse clicks in the menu or elsewhere in the window.
	// 2) It would probably have other side effects for other uses of popup menus.
	HWND fore_win = GetForegroundWindow();
	bool change_fore;
	if (change_fore = (!fore_win || GetWindowThreadProcessId(fore_win, NULL) != g_MainThreadID))
	{
		// Always bring main window to foreground right before TrackPopupMenu(), even if window is hidden.
		// UPDATE: This is a problem because SetForegroundWindowEx() will restore the window if it's hidden,
		// but restoring also shows the window if it's hidden.  Could re-hide it... but the question here
		// is can a minimized window be the foreground window?  If not, how to explain why
		// SetForegroundWindow() always seems to work for the purpose of the tray menu?
		//if (aForceToForeground)
		//{
		//	// Seems best to avoid using the script's current setting of #WinActivateForce.  Instead, always
		//	// try the gentle approach first since it is unlikely that displaying a menu will cause the
		//	// "flashing task bar button" problem?
		//	bool original_setting = g_WinActivateForce;
		//	g_WinActivateForce = false;
		//	SetForegroundWindowEx(g_hWnd);
		//	g_WinActivateForce = original_setting;
		//}
		//else
		if (!SetForegroundWindow(g_hWnd))
		{
			// The below fixes the problem where the menu cannot be canceled by clicking outside of
			// it (due to the main window not being active).  That usually happens the first time the
			// menu is displayed after the script launches.  0 is not enough sleep time, but 10 is:
			SLEEP_WITHOUT_INTERRUPTION(10);
			SetForegroundWindow(g_hWnd);  // 2nd time always seems to work for this particular window.
			// OLDER NOTES:
			// Always bring main window to foreground right before TrackPopupMenu(), even if window is hidden.
			// UPDATE: This is a problem because SetForegroundWindowEx() will restore the window if it's hidden,
			// but restoring also shows the window if it's hidden.  Could re-hide it... but the question here
			// is can a minimized window be the foreground window?  If not, how to explain why
			// SetForegroundWindow() always seems to work for the purpose of displaying the tray menu?
			//if (aForceToForeground)
			//{
			//	// Seems best to avoid using the script's current setting of #WinActivateForce.  Instead, always
			//	// try the gentle approach first since it is unlikely that displaying a menu will cause the
			//	// "flashing task bar button" problem?
			//	bool original_setting = g_WinActivateForce;
			//	g_WinActivateForce = false;
			//	SetForegroundWindowEx(g_hWnd);
			//	g_WinActivateForce = original_setting;
			//}
			//else
			//...
		}
	}
	// Apparently, the HWND parameter of TrackPopupMenuEx() can be g_hWnd even if one of the script's
	// other (non-main) windows is foreground. The menu still seems to operate correctly.
	g_MenuIsVisible = MENU_TYPE_POPUP; // It seems this is also set by WM_ENTERMENULOOP because apparently, TrackPopupMenuEx generates WM_ENTERMENULOOP. So it's done here just for added safety in case WM_ENTERMENULOOP isn't ALWAYS generated.
	TrackPopupMenuEx(mMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON, pt.x, pt.y, g_hWnd, NULL);
	g_MenuIsVisible = MENU_TYPE_NONE;
	// MSDN recommends this to prevent menu from closing on 2nd click.  MSDN also says that it's only
	// necessary to do this "for a notification icon". So to to avoid unnecessary launches of MsgSleep(),
	// its done only for the tray menu in v1.0.35.12:
	if (this == g_script.mTrayMenu)
		PostMessage(g_hWnd, WM_NULL, 0, 0);
	else // Seems best to avoid the following for the tray menu since it doesn't seem work and might produce side-effects in some cases.
	{
		if (change_fore && fore_win && GetForegroundWindow() == g_hWnd)
		{
			// The last of the conditions above is checked in case the user clicked the taskbar or some
			// other window to dismiss the menu.  In that case, the following isn't done because it typically
			// steals focus from the user's intended window, and this attempt usually fails due to the OS's
			// anti-focus-stealing measure, which in turn would cause fore_win's taskbar button to flash annoyingly.
			SetForegroundWindow(fore_win); // See comments above for why SetForegroundWindowEx() isn't used.
			// The following resolves the issue where the window would not have enough time to become active
			// before we continued using our timeslice to return to our caller and launch our new thread.
			// In other words, the menu thread would launch before SetForegroundWindow() actually had a chance
			// to take effect:
			// 0 is exactly the amount of time (-1 is not enough because it doesn't yield) needed for that
			// other process to actually ack/perform the activation of its window and clean out its queue using
			// one timeslice.  This has been tested even when the CPU is maxed from some third-party process.
			// For performance and code simplicity, it seems best not to do a GetForegroundWindow() loop that
			// waits for it to become active (unless others report that this method is significantly unreliable):
			SLEEP_WITHOUT_INTERRUPTION(0);
		}
	}
	// Fix for v1.0.38.05: If the current thread is interruptible (which it should be since a menu was just
	// displayed, which almost certainly timed out the default Thread Interrupt setting), the following
	// MsgSleep() will launch the selected menu item's subroutine.  This fix is needed because of a change
	// in v1.0.38.04, namely the line "g_script.mLastPeekTime = tick_now;" in IsCycleComplete().
	// The root problem here is that it would not be intuitive to allow the command after
	// "Menu, MyMenu, Show" should to run before the menu item's subroutine launches as a new thread.
	// 
	// You could argue that selecting a menu item should immediately execute the selected menu item's
	// callback rather than queuing it up as a new thread.  However, even if that is a better method,
	// it would break existing scripts that rely on new-thread behavior (such as fresh default for
	// SetKeyDelay).
	//
	// Without this fix, a script such as the following (and many other things similar) would
	// counterintuitively fail to launch the selected item's subroutine:
	// Menu, MyMenu, Add, NOTEPAD
	// Menu, MyMenu, Show
	// ; Sleep 0  ; Uncommenting this line was necessary in v1.0.38.04 but not any other versions.
	// ExitApp
	MsgSleep(-1);
	return OK;
}



UINT UserMenuItem::Pos()
{
	UINT pos = 0;
	for (UserMenuItem *mi = mMenu->mFirstMenuItem; mi; mi = mi->mNextMenuItem, ++pos)
		if (mi == this)
			return pos;
	return UINT_MAX;
}



bool UserMenu::ContainsMenu(UserMenu *aMenu)
{
	if (!aMenu)
		return false;
	// For each submenu in mMenu: Check if it or any of its submenus equals aMenu.
	for (UserMenuItem *mi = mFirstMenuItem; mi; mi = mi->mNextMenuItem)
		if (mi->mSubmenu)
			if (mi->mSubmenu == aMenu || mi->mSubmenu->ContainsMenu(aMenu)) // recursive
				return true;
			//else keep searching
	return false;
}



void UserMenu::UpdateAccelerators()
{
	if (!mMenu)
		// Menu doesn't exist yet, so can't be attached (directly or indirectly) to any GUIs.
		return;

	if (mMenuType == MENU_TYPE_BAR)
	{
		for (GuiType* gui = g_firstGui; gui; gui = gui->mNextGui)
			if (GetMenu(gui->mHwnd) == mMenu)
			{
				gui->UpdateAccelerators(*this);
				// Continue in case there are other GUIs using this menu.
				//break;
			}
	}
	else
	{
		// This menu isn't a menu bar, but perhaps it is contained by one.
		for (UserMenu *menu = g_script.mFirstMenu; menu; menu = menu->mNextMenu)
			if (menu->mMenuType == MENU_TYPE_BAR && menu->ContainsMenu(this))
			{
				menu->UpdateAccelerators();
				// Continue in case there are other menus which contain this submenu.
				//break;
			}
		return;
	}
}



//
// L17: Menu-item icon functions.
//


ResultType UserMenu::SetItemIcon(UserMenuItem *aMenuItem, LPCTSTR aFilename, int aIconNumber, int aWidth)
{
	if (!*aFilename || (*aFilename == '*' && !aFilename[1]))
	{
		RemoveItemIcon(aMenuItem);
		return OK;
	}

	int image_type;
	HBITMAP new_icon, new_copy;
	// Currently height is always -1 and cannot be overridden. -1 means maintain aspect ratio, usually 1:1 for icons.
	if ( !(new_icon = LoadPicture(aFilename, aWidth, -1, image_type, aIconNumber, false)) )
		return FAIL;

	if (image_type != IMAGE_BITMAP) // Convert to 32-bit bitmap:
	{
		new_copy = IconToBitmap32((HICON)new_icon, true);
		// Even if conversion failed, we have no further use for the icon:
		DestroyIcon((HICON)new_icon);
		if (!new_copy)
			return FAIL;
		new_icon = new_copy;
	}

	if (aMenuItem->mBitmap) // Delete previous bitmap.
		DeleteObject(aMenuItem->mBitmap);
	
	aMenuItem->mBitmap = new_icon;

	if (mMenu)
		ApplyItemIcon(aMenuItem);

	return aMenuItem->mBitmap ? OK : FAIL;
}


// Caller has ensured mMenu is non-NULL.
void UserMenu::ApplyItemIcon(UserMenuItem *aMenuItem)
{
	if (aMenuItem->mBitmap)
	{
		MENUITEMINFO item_info;
		item_info.cbSize = sizeof(MENUITEMINFO);
		item_info.fMask = MIIM_BITMAP;
		// Set HBMMENU_CALLBACK or 32-bit bitmap as appropriate.
		item_info.hbmpItem = aMenuItem->mBitmap;
		SetMenuItemInfo(mMenu, aMenuItem_ID, aMenuItem_MF_BY, &item_info);
	}
}


void UserMenu::RemoveItemIcon(UserMenuItem *aMenuItem)
{
	if (aMenuItem->mBitmap)
	{
		if (mMenu)
		{
			MENUITEMINFO item_info;
			item_info.cbSize = sizeof(MENUITEMINFO);
			item_info.fMask = MIIM_BITMAP;
			item_info.hbmpItem = NULL;
			SetMenuItemInfo(mMenu, aMenuItem_ID, aMenuItem_MF_BY, &item_info);
		}
		DeleteObject(aMenuItem->mBitmap);
		aMenuItem->mBitmap = NULL;
	}
}

