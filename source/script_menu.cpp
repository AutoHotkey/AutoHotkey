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


ResultType Script::PerformMenu(char *aMenu, char *aCommand, char *aParam3, char *aParam4, char *aOptions, char *aOptions2)
{
	if (mMenuUseErrorLevel)
		g_ErrorLevel->Assign(ERRORLEVEL_NONE);  // Set default, which is "none" for the Menu command.

	#define RETURN_MENU_ERROR(msg, info) return mMenuUseErrorLevel ? g_ErrorLevel->Assign(ERRORLEVEL_ERROR) \
		: ScriptError(msg ERR_ABORT, info)
	#define RETURN_IF_NOT_TRAY if (!is_tray) RETURN_MENU_ERROR(ERR_MENUTRAY, aMenu)

	MenuCommands menu_command = Line::ConvertMenuCommand(aCommand);
	if (menu_command == MENU_CMD_INVALID)
		RETURN_MENU_ERROR(ERR_PARAM2_INVALID, aCommand);

	bool is_tray = !stricmp(aMenu, "tray");

	// Handle early on anything that doesn't require the menu to be found or created:
	switch(menu_command)
	{
	case MENU_CMD_USEERRORLEVEL:
		mMenuUseErrorLevel = (Line::ConvertOnOff(aParam3) != TOGGLED_OFF);
		// Even though the state may have changed by the above, it doesn't seem necessary
		// to adjust on the fly for the purpose of this particular return.  In other words,
		// the old mode will be in effect for this one return:
		return OK;

	case MENU_CMD_TIP:
		RETURN_IF_NOT_TRAY;
		if (*aParam3)
		{
			if (!mTrayIconTip)
				mTrayIconTip = SimpleHeap::Malloc(sizeof(mNIC.szTip)); // SimpleHeap improves avg. case mem load.
			if (mTrayIconTip)
				strlcpy(mTrayIconTip, aParam3, sizeof(mNIC.szTip));
		}
		else // Restore tip to default.
			if (mTrayIconTip)
				*mTrayIconTip = '\0';
		if (mNIC.hWnd) // i.e. only update the tip if the tray icon exists (can't work otherwise).
		{
			UPDATE_TIP_FIELD
			Shell_NotifyIcon(NIM_MODIFY, &mNIC);  // Currently not checking its result (e.g. in case a shell other than Explorer is running).
		}
		return OK;

	case MENU_CMD_ICON:
		// L17: If is_tray and aParam3 is omitted or aParam4 is an integer, set the tray icon. Otherwise set a menu item icon.
		if (is_tray)
		{
			bool mIconFrozen_prev = mIconFrozen;
			// Lexikos: aOptions still ambiguous with menu item icon number at this point.
			//if (*aOptions) // i.e. if it's blank, don't change the current setting of mIconFrozen.
			//	mIconFrozen = (ATOI(aOptions) == 1);

			if (!*aParam3)
			{
				// Lexikos: MenuItemName omitted, therefore no conflict. mIconFrozen may now be set.
				if (*aOptions) // i.e. if it's blank, don't change the current setting of mIconFrozen.
					mIconFrozen = (ATOI(aOptions) == 1);

				g_NoTrayIcon = false;
				if (!mNIC.hWnd) // The icon doesn't exist, so create it.
				{
					CreateTrayIcon();
					UpdateTrayIcon(true);  // Force the icon into the correct pause/suspend state.
				}
				else if (!mIconFrozen && mIconFrozen_prev) // To cause "Menu Tray, Icon,,, 0" to update the icon while the script is suspended.
					UpdateTrayIcon(true);
				return OK;
			}

			// Otherwise, user has specified a custom icon:
			if (*aParam3 == '*' && !*(aParam3 + 1)) // Restore the standard icon.
			{
				// Lexikos: For compatibility with older scripts, "Menu, Tray, Icon, *" must reset tray to default icon, even if an item "*" exists. mIconFrozen may now be set.
				if (*aOptions) // i.e. if it's blank, don't change the current setting of mIconFrozen.
					mIconFrozen = (ATOI(aOptions) == 1);

				if (mCustomIcon)
				{
					GuiType::DestroyIconsIfUnused(mCustomIcon, mCustomIconSmall); // v1.0.37.07: Solves reports of Gui windows losing their icons.
					// If the above doesn't destroy the icon, the GUI window(s) still using it are responsible for
					// destroying it later.
					mCustomIcon = NULL;  // To indicate that there is no custom icon.
					mCustomIconSmall = NULL;
					if (mCustomIconFile)
						*mCustomIconFile = '\0';
					mCustomIconNumber = 0;
					UpdateTrayIcon(true);  // Need to use true in this case too.
				}
				return OK;
			}

			if (IsPureNumeric(aParam4, true)) // pure integer or empty/whitespace
			{
				// Lexikos: We are unconditionally treating this as a request to set the tray icon, so mIconFrozen may now be set.
				if (*aOptions) // i.e. if it's blank, don't change the current setting of mIconFrozen.
					mIconFrozen = (ATOI(aOptions) == 1);

				// v1.0.43.03: Load via LoadPicture() vs. ExtractIcon() because ExtractIcon harms the quality
				// of 16x16 icons inside .ico files by first scaling them to 32x32 (which then has to be scaled
				// back to 16x16 for the tray and for the SysMenu icon). I've visually confirmed that the
				// distortion occurs at least when a 16x16 icon is loaded by ExtractIcon() then put into the
				// tray.  It might not be the scaling itself that distorts the icon: the pixels are all in the
				// right places, it's just that some are the wrong color/shade. This implies that some kind of
				// unwanted interpolation or color tweaking is being done by ExtractIcon (and probably LoadIcon),
				// but not by LoadImage.
				// Also, load the icon at actual size so that when/if this icon is used for a GUI window, its
				// appearance in the alt-tab menu won't be unexpectedly poor due to having been scaled from its
				// native size down to 16x16.
				int icon_number;
				if (*aParam4)
				{
					icon_number = ATOI(aParam4);
					if (icon_number == 0) // Must validate for use in two places below.
						icon_number = 1; // Must be != 0 to tell LoadPicture that "icon must be loaded, never a bitmap".
				}
				else
					icon_number = 1; // One vs. Zero tells LoadIcon: "must load icon, never a bitmap (e.g. no gif/jpg/png)".

				int image_type;
				// L17: For best results, load separate small and large icons.
				HICON new_icon_small;
				HICON new_icon = NULL; // Initialize to detect failure to load either icon.
				if ( new_icon_small = (HICON)LoadPicture(aParam3, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), image_type, icon_number, false) ) // Called with icon_number > 0, it guarantees return of an HICON/HCURSOR, never an HBITMAP.
					if ( !(new_icon = (HICON)LoadPicture(aParam3, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), image_type, icon_number, false)) )
						DestroyIcon(new_icon_small);
				if ( !new_icon )
					RETURN_MENU_ERROR("Can't load icon.", aParam3);

				GuiType::DestroyIconsIfUnused(mCustomIcon, mCustomIconSmall); // This destroys it if non-NULL and it's not used by an GUI windows.

				mCustomIcon = new_icon;
				mCustomIconSmall = new_icon_small;
				mCustomIconNumber = icon_number;
				// Allocate the full MAX_PATH in case the contents grow longer later.
				// SimpleHeap improves avg. case mem load:
				if (!mCustomIconFile)
					mCustomIconFile = SimpleHeap::Malloc(MAX_PATH);
				if (mCustomIconFile)
				{
					// Get the full path in case it's a relative path.  This is documented and it's done in case
					// the script ever changes its working directory:
					char full_path[MAX_PATH], *filename_marker;
					if (GetFullPathName(aParam3, sizeof(full_path) - 1, full_path, &filename_marker))
						strlcpy(mCustomIconFile, full_path, MAX_PATH);
					else
						strlcpy(mCustomIconFile, aParam3, MAX_PATH);
				}

				if (!g_NoTrayIcon)
					UpdateTrayIcon(true);  // Need to use true in this case too.
				return OK;
			}
		}
		break;

	case MENU_CMD_NOICON:
		if (is_tray && !*aParam3) // L17: "Menu, Tray, NoIcon, xxx" removes icon from tray menu item xxx.
		{
			g_NoTrayIcon = true;
			if (mNIC.hWnd) // Since it exists, destroy it.
			{
				Shell_NotifyIcon(NIM_DELETE, &mNIC); // Remove it.
				mNIC.hWnd = NULL;  // Set this as an indicator that tray icon is not installed.
				// but don't do DestroyMenu() on mTrayMenu->mMenu (if non-NULL) since it may have been
				// changed by the user to have the custom items on top of the standard items,
				// for example, and we don't want to lose that ordering in case the script turns
				// the icon back on at some future time during this session.
			}
			return OK;
		}
		// else: this request to remove a menu item's icon will be processed below.
		break;

	case MENU_CMD_CLICK:
		RETURN_IF_NOT_TRAY;
		mTrayMenu->mClickCount = ATOI(aParam3);
		if (mTrayMenu->mClickCount < 1)
			mTrayMenu->mClickCount = 1;  // Single-click to activate menu's default item.
		else if (mTrayMenu->mClickCount > 2)
			mTrayMenu->mClickCount = 2;  // Double-click.
		return OK;

	case MENU_CMD_MAINWINDOW:
		RETURN_IF_NOT_TRAY;
#ifdef AUTOHOTKEYSC
		if (!g_AllowMainWindow)
		{
			g_AllowMainWindow = true;
			EnableOrDisableViewMenuItems(GetMenu(g_hWnd), MF_ENABLED); // Added as a fix in v1.0.47.06.
			// Rather than using InsertMenu() to insert the item in the right position,
			// which makes the code rather unmaintainable, it seems best just to recreate
			// the entire menu.  This will result in the standard menu items going back
			// up to the top of the menu if the user previously had them at the bottom,
			// but it seems too rare to worry about, especially since it's easy to
			// work around that:
			if (mTrayMenu->mIncludeStandardItems)
				mTrayMenu->Destroy(); // It will be recreated automatically the next time the user displays it.
			// else there's no need.
		}
#endif
        return OK;

	case MENU_CMD_NOMAINWINDOW:
		RETURN_IF_NOT_TRAY;
#ifdef AUTOHOTKEYSC
		if (g_AllowMainWindow)
		{
			g_AllowMainWindow = false;
			EnableOrDisableViewMenuItems(GetMenu(g_hWnd), MF_DISABLED | MF_GRAYED); // Added as a fix in v1.0.47.06.
			// See comments in the prior case above for why it's done this way vs. using DeleteMenu():
			if (mTrayMenu->mIncludeStandardItems)
				mTrayMenu->Destroy(); // It will be recreated automatically the next time the user displays it.
			// else there's no need.
		}
#endif
		return OK;
	} // switch()


	// Now that most opportunities to return an error have passed, find or create the menu, since
	// all the commands that haven't already been fully handled above will need it:
	UserMenu *menu = FindMenu(aMenu);
	if (!menu)
	{
		// Menus can be created only in conjuction with the ADD command. Update: As of v1.0.25.12, they can
		// also be created with the "Menu, MyMenu, Standard" command.
		if (menu_command != MENU_CMD_ADD && menu_command != MENU_CMD_STANDARD)
			RETURN_MENU_ERROR(ERR_MENU, aMenu);
		if (   !(menu = AddMenu(aMenu))   )
			RETURN_MENU_ERROR("Menu name too long.", aMenu); // Could also be "out of mem" but that's too rare to display.
	}

	// The above has found or added the menu for use below.

	switch(menu_command)
	{
	case MENU_CMD_SHOW:
		return menu->Display(true, *aParam3 ? ATOI(aParam3) : COORD_UNSPECIFIED, *aParam4 ? ATOI(aParam4) : COORD_UNSPECIFIED);
	case MENU_CMD_ADD:
		if (*aParam3) // Since a menu item name was given, it's not a separator line.
			break;    // Let a later switch() handle it.
		if (!menu->AddItem("", 0, NULL, NULL, ""))
			RETURN_MENU_ERROR(ERR_OUTOFMEM, "");  // Out of mem should be the only possibility in this case.
		return OK;
	case MENU_CMD_DELETE:
		if (*aParam3) // Since a menu item name was given, an item is being deleted, not the whole menu.
			break;    // Let a later switch() handle it.
		if (menu == mTrayMenu)
			RETURN_MENU_ERROR("Tray menu must not be deleted.", "");
		if (!ScriptDeleteMenu(menu))
			RETURN_MENU_ERROR("Can't delete menu (in use?).", menu->mName); // Possibly in use as a menu bar.
		return OK;
	case MENU_CMD_DELETEALL:
		if (!menu->DeleteAllItems())
			RETURN_MENU_ERROR("Can't delete items (in use?).", menu->mName); // Possibly in use as a menu bar.
		return OK;
	case MENU_CMD_DEFAULT:
		if (*aParam3) // Since a menu item has been specified, let a later switch() handle it.
			break;
		//else no menu item, so it's the same as NoDefault: fall through to the next case.
	case MENU_CMD_NODEFAULT:
		return menu->SetDefault();
	case MENU_CMD_STANDARD:
		menu->IncludeStandardItems(); // Since failure is very rare, no check of its return value is done.
		return OK;
	case MENU_CMD_NOSTANDARD:
		menu->ExcludeStandardItems(); // Since failure is very rare, no check of its return value is done.
		return OK;
	case MENU_CMD_COLOR:
		menu->SetColor(aParam3, stricmp(aParam4, "Single"));
		return OK;
	}

	// All the remaining commands need a menu item to operate upon, or some other requirement met below.

	char *new_name = "";
	if (menu_command == MENU_CMD_RENAME) // aParam4 contains the menu item's new name in this case.
	{
		new_name = aParam4;
		aParam4 = "";
	}

	// The above has handled all cases that don't require a menu item to be found or added,
	// including the adding separator lines.  So at the point, it is necessary to either find
	// or create a menu item.  The latter only occurs for the ADD command.
	if (!*aParam3)
		RETURN_MENU_ERROR("Parameter #3 must not be blank in this case.", "");

	// Find the menu item name AND its previous item (needed for the DELETE command) in the linked list:
	UserMenuItem *mi, *menu_item = NULL, *menu_item_prev = NULL; // Set defaults.
	for (menu_item = menu->mFirstMenuItem
		; menu_item
		; menu_item_prev = menu_item, menu_item = menu_item->mNextMenuItem)
		if (!lstrcmpi(menu_item->mName, aParam3)) // Match found (case insensitive).
			break;

	// Whether an existing menu item's options should be updated without updating its submenu or label:
	bool update_exiting_item_options = (menu_command == MENU_CMD_ADD && menu_item && !*aParam4 && *aOptions);

	// Seems best to avoid performance enhancers such as (Label *)mAttribute here, since the "Menu"
	// command has so many modes of operation that would be difficult to parse at load-time:
	Label *target_label = NULL;  // Set default.
	UserMenu *submenu = NULL;    // Set default.
	if (menu_command == MENU_CMD_ADD && !update_exiting_item_options) // Labels and submenus are only used in conjuction with the ADD command.
	{
		if (!*aParam4) // Allow the label/submenu to default to the menu name.
			aParam4 = aParam3; // Note that aParam3 will be blank in the case of a separator line.
		if (*aParam4)
		{
			if (*aParam4 == ':') // It's a submenu.
			{
				++aParam4;
				if (   !(submenu = FindMenu(aParam4))   )
					RETURN_MENU_ERROR(ERR_SUBMENU, aParam4);
				// Before going further: since a submenu has been specified, make sure that the parent
				// menu is not included anywhere in the nested hierarchy of that submenu's submenus.
				// The OS doesn't seem to like that, creating empty or strange menus if it's attempted:
				if (   submenu && (submenu == menu || submenu->ContainsMenu(menu))   )
					RETURN_MENU_ERROR("Submenu must not contain its parent menu.", aParam4);
			}
			else // It's a label.
				if (   !(target_label = FindLabel(aParam4))   )
					RETURN_MENU_ERROR(ERR_NO_LABEL, aParam4);
		}
	}

	if (!menu_item)  // menu item doesn't exist, so create it (but only if the command is ADD).
	{
		if (menu_command != MENU_CMD_ADD)
			// Seems best not to create menu items on-demand like this because they might get put into
			// an incorrect position (i.e. it seems better than menu changes be kept separate from
			// menu additions):
			RETURN_MENU_ERROR("Nonexistent menu item.", aParam3);

		// Otherwise: Adding a new item that doesn't yet exist.
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
		// last-found free ID has a high likelyhood of still being in use:
		++sLastFreeID;
		bool id_in_use;
		UserMenu *m;
		// Note that the i variable is used to force the loop to complete exactly one full
		// circuit through all available IDs, regardless of where the starting/cached value:
		for (int i = 0; i < (ID_USER_LAST - ID_USER_FIRST + 1); ++i, ++sLastFreeID) // FOR EACH ID
		{
			if (sLastFreeID > ID_USER_LAST)
				sLastFreeID = ID_USER_FIRST;  // Wrap around to the beginning so that one complete circuit is made.
			id_in_use = false;  // Reset the default each iteration (overridden if the below finds a match).
			for (m = mFirstMenu; m; m = m->mNextMenu) // FOR EACH MENU
			{
				for (mi = m->mFirstMenuItem; mi; mi = mi->mNextMenuItem) // FOR EACH MENU ITEM
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
		if (id_in_use) // All ~64000 IDs are in use!
			RETURN_MENU_ERROR("Too many menu items.", aParam3); // Short msg since so rare.
		if (!menu->AddItem(aParam3, sLastFreeID, target_label, submenu, aOptions))
			RETURN_MENU_ERROR("Menu item name too long.", aParam3); // Can also happen due to out-of-mem, but that's too rare to display.
		return OK;  // Item has been successfully added with the correct properties.
	} // if (!menu_item)

	// Above has found the correct menu_item to operate upon (it already returned if
	// the item was just created).  Since the item was found, the UserMenu's popup
	// menu must already exist because a UserMenu object can't have menu items unless
	// its menu exists.

	switch (menu_command)
	{
	case MENU_CMD_ADD:
		// This is only reached if the ADD command is being used to update the label, submenu, or
		// options of an existing menu item (since it would have returned above if the item was
		// just newly created).
		return menu->ModifyItem(menu_item, target_label, submenu, aOptions);
	case MENU_CMD_RENAME:
		if (!menu->RenameItem(menu_item, new_name))
			RETURN_MENU_ERROR("Menu item name already in use (or too long).", new_name);
		return OK;
	case MENU_CMD_CHECK:
		return menu->CheckItem(menu_item);
	case MENU_CMD_UNCHECK:
		return menu->UncheckItem(menu_item);
	case MENU_CMD_TOGGLECHECK:
		return menu->ToggleCheckItem(menu_item);
	case MENU_CMD_ENABLE:
		return menu->EnableItem(menu_item);
	case MENU_CMD_DISABLE: // Disables and grays the item.
		return menu->DisableItem(menu_item);
	case MENU_CMD_TOGGLEENABLE:
		return menu->ToggleEnableItem(menu_item);
	case MENU_CMD_DEFAULT:
		return menu->SetDefault(menu_item);
	case MENU_CMD_DELETE:
		return menu->DeleteItem(menu_item, menu_item_prev);
	// L17: Set or remove a menu item's icon.
	case MENU_CMD_ICON:
		// aOptions2: Icon width if specified. Defaults to system small icon size; original icon size will be used if aOptions2 is "0".
		if (!menu->SetItemIcon(menu_item, aParam4, ATOI(aOptions), !*aOptions2 ? GetSystemMetrics(SM_CXSMICON) : ATOI(aOptions2)))
			RETURN_MENU_ERROR("Can't load icon.", aParam4);
		return OK;
	case MENU_CMD_NOICON:
		return menu->RemoveItemIcon(menu_item);
	} // switch()
	return FAIL;  // Should never be reached, but avoids compiler warning and improves bug detection.
}



UserMenu *Script::FindMenu(char *aMenuName)
// Returns the UserMenu whose name matches aMenuName, or NULL if not found.
{
	if (!aMenuName || !*aMenuName) return NULL;
	for (UserMenu *menu = mFirstMenu; menu != NULL; menu = menu->mNextMenu)
		if (!lstrcmpi(menu->mName, aMenuName)) // Match found.
			return menu;
	return NULL; // No match found.
}



UserMenu *Script::AddMenu(char *aMenuName)
// Caller must have already ensured aMenuName doesn't exist yet in the list.
// Returns the newly created UserMenu object.
{
	if (!aMenuName || !*aMenuName) return NULL;
	size_t length = strlen(aMenuName);
	if (length > MAX_MENU_NAME_LENGTH)
		return NULL;  // Caller should show error if desired.
	// After mem is allocated, the object takes charge of its later deletion:
	char *name_dynamic = new char[length + 1];  // +1 for terminator.
	if (!name_dynamic)
		return NULL;  // Caller should show error if desired.
	strcpy(name_dynamic, aMenuName);
	UserMenu *menu = new UserMenu(name_dynamic);
	if (!menu)
	{
		delete name_dynamic;
		return NULL;  // Caller should show error if desired.
	}
	if (!mFirstMenu)
		mFirstMenu = mLastMenu = menu;
	else
	{
		mLastMenu->mNextMenu = menu;
		// This must be done after the above:
		mLastMenu = menu;
	}
	++mMenuCount;  // Only after memory has been successfully allocated.
	return menu;
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
	UserMenuItem *mi, *mi_prev, *mi_to_delete;
	for (UserMenu *m = mFirstMenu; m; m = m->mNextMenu)
		if (m != aMenu) // Don't bother with this menu even if it's submenu of itself, since it will be destroyed anyway.
			for (mi = m->mFirstMenuItem, mi_prev = NULL; mi;)
			{
				mi_to_delete = mi;
				mi = mi->mNextMenuItem;
				if (mi_to_delete->mSubmenu == aMenu)
					m->DeleteItem(mi_to_delete, mi_prev);
				else
					mi_prev = mi_to_delete;
			}
	// Remove aMenu from the linked list.  First find the item that occurs prior the aMenu in the list:
	UserMenu *aMenu_prev;
	for (aMenu_prev = mFirstMenu; aMenu_prev; aMenu_prev = aMenu_prev->mNextMenu)
		if (aMenu_prev->mNextMenu == aMenu)
			break;
	if (aMenu == mLastMenu)
		mLastMenu = aMenu_prev; // Can be NULL if the list will now be empty.
	if (aMenu_prev) // there is another item prior to aMenu in the linked list.
		aMenu_prev->mNextMenu = aMenu->mNextMenu; // Can be NULL if aMenu was the last one.
	else // aMenu was the first one in the list.
		mFirstMenu = aMenu->mNextMenu; // Can be NULL if the list will now be empty.
	// Do this last when its contents are no longer needed.  Its destructor will delete all
	// the items in the menu and destroy the OS menu itself:
	aMenu->DeleteAllItems(); // This also calls Destroy() to free the menu's resources.
	if (aMenu->mBrush) // Free the brush used for the menu's background color.
		DeleteObject(aMenu->mBrush);
	delete aMenu->mName; // Since it was separately allocated.
	delete aMenu;
	--mMenuCount;
	return OK;
}



// Macros for use with the below methods:
#define aMenuItem_ID (aMenuItem->mSubmenu ? GetSubmenuPos(aMenuItem->mSubmenu->mMenu) : aMenuItem->mMenuID)
#define aMenuItem_MF_BY (aMenuItem->mSubmenu ? MF_BYPOSITION : MF_BYCOMMAND)
#define UPDATE_GUI_MENU_BARS(menu_type, hmenu) \
	if (menu_type == MENU_TYPE_BAR && GuiType::sGuiCount)\
		GuiType::UpdateMenuBars(hmenu); // Above: If it's not a popup, it's probably a menu bar.


#ifdef AUTOHOTKEYSC
#define CHANGE_DEFAULT_IF_NEEDED \
	if (mDefault == aMenuItem)\
	{\
		if (mMenu)\
		{\
			if (this == g_script.mTrayMenu)\
				SetMenuDefaultItem(mMenu, mIncludeStandardItems && g_AllowMainWindow ? ID_TRAY_OPEN : -1, FALSE);\
			else\
				SetMenuDefaultItem(mMenu, -1, FALSE);\
		}\
		mDefault = NULL;\
	}
#else
#define CHANGE_DEFAULT_IF_NEEDED \
	if (mDefault == aMenuItem)\
	{\
		if (mMenu)\
		{\
			if (this == g_script.mTrayMenu)\
				SetMenuDefaultItem(mMenu, mIncludeStandardItems ? ID_TRAY_OPEN : -1, FALSE);\
			else\
				SetMenuDefaultItem(mMenu, -1, FALSE);\
		}\
		mDefault = NULL;\
	}
#endif



ResultType UserMenu::AddItem(char *aName, UINT aMenuID, Label *aLabel, UserMenu *aSubmenu, char *aOptions)
// Caller must have already ensured that aName does not yet exist as a user-defined menu item
// in this->mMenu.
{
	size_t length = strlen(aName);
	if (length > MAX_MENU_NAME_LENGTH)
		return FAIL;  // Caller should show error if desired.
	// After mem is allocated, the object takes charge of its later deletion:
	char *name_dynamic;
	if (length)
	{
		if (   !(name_dynamic = new char[length + 1])   )  // +1 for terminator.
			return FAIL;  // Caller should show error if desired.
		strcpy(name_dynamic, aName);
	}
	else
		name_dynamic = Var::sEmptyString; // So that it can be detected as a non-allocated empty string.
	UserMenuItem *menu_item = new UserMenuItem(name_dynamic, length + 1, aMenuID, aLabel, aSubmenu, this);
	if (!menu_item) // Should also be very rare.
	{
		if (name_dynamic != Var::sEmptyString)
			delete name_dynamic;
		return FAIL;  // Caller should show error if desired.
	}
	if (!mFirstMenuItem)
		mFirstMenuItem = mLastMenuItem = menu_item;
	else
	{
		mLastMenuItem->mNextMenuItem = menu_item;
		// This must be done after the above:
		mLastMenuItem = menu_item;
	}
	++mMenuItemCount;  // Only after memory has been successfully allocated.
	if (*aOptions)
		UpdateOptions(menu_item, aOptions);
	return OK;
}



UserMenuItem::UserMenuItem(char *aName, size_t aNameCapacity, UINT aMenuID, Label *aLabel, UserMenu *aSubmenu, UserMenu *aMenu)
// UserMenuItem Constructor.
	: mName(aName), mNameCapacity(aNameCapacity), mMenuID(aMenuID), mLabel(aLabel), mSubmenu(aSubmenu), mMenu(aMenu)
	, mPriority(0) // default priority = 0
	, mEnabled(true), mChecked(false), mNextMenuItem(NULL)
	, mIcon(NULL) // L17: Initialize mIcon/mBitmap union.
{
	if (aMenu->mMenu)
	{
		if (aSubmenu) // Ensure the menu is created so that AppendMenu() will function properly.
			aSubmenu->Create();
		AppendMenu(aMenu->mMenu, (*aName ? MF_STRING : MF_SEPARATOR) | (aSubmenu ? MF_POPUP : 0)
			, aSubmenu ? (UINT_PTR)aSubmenu->mMenu : aMenuID, aName);
		UPDATE_GUI_MENU_BARS(aMenu->mMenuType, aMenu->mMenu)
	}
}



ResultType UserMenu::DeleteItem(UserMenuItem *aMenuItem, UserMenuItem *aMenuItemPrev)
{
	// Remove this menu item from the linked list:
	if (aMenuItem == mLastMenuItem)
		mLastMenuItem = aMenuItemPrev; // Can be NULL if the list will now be empty.
	if (aMenuItemPrev) // there is another item prior to aMenuItem in the linked list.
		aMenuItemPrev->mNextMenuItem = aMenuItem->mNextMenuItem; // Can be NULL if aMenuItem was the last one.
	else // aMenuItem was the first one in the list.
		mFirstMenuItem = aMenuItem->mNextMenuItem; // Can be NULL if the list will now be empty.
	CHANGE_DEFAULT_IF_NEEDED  // Should do this before freeing aMenuItem's memory.
	if (mMenu) // Delete the item from the menu.
		RemoveMenu(mMenu, aMenuItem_ID, aMenuItem_MF_BY); // v1.0.48: Lexikos: DeleteMenu() destroys any sub-menu handle associated with the item, so use RemoveMenu. Otherwise the submenu handle stored somewhere else in memory would suddenly become invalid.
	RemoveItemIcon(aMenuItem); // L17: Free icon or bitmap.
	if (aMenuItem->mName != Var::sEmptyString)
		delete aMenuItem->mName; // Since it was separately allocated.
	delete aMenuItem; // Do this last when its contents are no longer needed.
	--mMenuItemCount;
	UPDATE_GUI_MENU_BARS(mMenuType, mMenu)  // Verified as being necessary.
	return OK;
}



ResultType UserMenu::DeleteAllItems()
{
	if (!mFirstMenuItem)
		return OK;  // If there are no user-defined menu items, it's already in the correct state.
	// Remove all menu items from the linked list and from the menu.  First destroy the menu since
	// it's probably better to start off fresh than have the destructor individually remove each
	// menu item as the items in the linked list are deleted.  In addition, this avoids the need
	// to find any submenus by position:
	if (!Destroy())  // if mStandardMenuItems is true, the menu will be recreated later when needed.
		// If menu can't be destroyed, it's probably due to it being attached as a menu bar to an existing
		// GUI window.  In this case, when the window is destroyed, the menu bar will be too, so it's
		// probably best to do nothing.  If we were called as a result of "menu, MenuName, Delete", it
		// is documented that this will fail in this case.
		return FAIL;
	// The destructor relies on the fact that the above destroys the menu but does not recreate it.
	// This is because popup menus, not being affiliated with a window (unless they're attached to
	// a menu bar, but that isn't the case with our GUI windows, which detach such menus prior to
	// when the GUI window is destroyed in case the menu is in use by another window), must be
	// destroyed with DestroyMenu() to ensure a clean exit (resources freed).
	UserMenuItem *menu_item_to_delete;
	for (UserMenuItem *mi = mFirstMenuItem; mi;)
	{
		menu_item_to_delete = mi;
		mi = mi->mNextMenuItem;
		RemoveItemIcon(menu_item_to_delete); // L26: Free icon or bitmap!
		if (menu_item_to_delete->mName != Var::sEmptyString)
			delete menu_item_to_delete->mName; // Since it was separately allocated.
		delete menu_item_to_delete;
	}
	mFirstMenuItem = mLastMenuItem = NULL;
	mMenuItemCount = 0;
	mDefault = NULL;  // i.e. there can't be a *user-defined* default item anymore, even if this is the tray.
	return OK;
}



ResultType UserMenu::ModifyItem(UserMenuItem *aMenuItem, Label *aLabel, UserMenu *aSubmenu, char *aOptions)
// Modify the label, submenu, or options of a menu item (exactly one of these should be NULL and the
// other not except when updating only the options).
// If a menu item becomes a submenu, we don't relinquish its ID in case it's ever made a normal item
// again (avoids the need to re-lookup a unique ID).
{
	if (*aOptions)
		UpdateOptions(aMenuItem, aOptions);
	if (!aLabel && !aSubmenu) // We were called only to update this item's options.
		return OK;

	aMenuItem->mLabel = aLabel;  // This will be NULL if this menu item is a separator or submenu.
	if (aMenuItem->mSubmenu == aSubmenu) // Below relies on this check.
		return OK;
	if (!mMenu)
	{
		aMenuItem->mSubmenu = aSubmenu;  // Just set the indicator for when the menu is later created.
		return OK;
	}

	// Otherwise, since the OS menu exists, one of these is to be done to aMenuItem in it:
	// 1) Change a submenu to point to a different menu.
	// 2) Change a submenu so that it becomes a normal menu item.
	// 3) Change a normal menu item into a submenu.

	// Since Create() ensures that aSubmenu is non-null whenever this->mMenu is non-null, this is just
	// an extra safety check in case some other method destroyed aSubmenu since then:
	if (aSubmenu)
		if (!aSubmenu->Create()) // Create if needed.  No error msg since so rare.
			return FAIL;

	MENUITEMINFO mii;
	mii.cbSize = sizeof(mii);
	mii.fMask = MIIM_SUBMENU | MIIM_ID;
	mii.hSubMenu = aSubmenu ? aSubmenu->mMenu : NULL;
	// If this submenu is being back into a normal menu item, the ID must be re-specified
	// because an item that was formerly submenu will not have a real ID due to OS behavior:
	mii.wID = aMenuItem->mMenuID;
	if (SetMenuItemInfo(mMenu, aMenuItem_ID, aMenuItem->mSubmenu != NULL, &mii))
	{
		// Submenu was just made into a different submenu or converted into a normal menu item.
		// Since the OS (as an undocumented side effect) sometimes destroys the menu itself when
		// a submenu is changed in this way, update our state to indicate that the menu handle
		// is no longer valid:
		if (aMenuItem->mSubmenu && aMenuItem->mSubmenu->mMenu && !IsMenu(aMenuItem->mSubmenu->mMenu))
		{
			UserMenu *temp = aMenuItem->mSubmenu;
			aMenuItem->mSubmenu = aSubmenu; // Should be done before the below so that Destroy() sees the change.
			// The following shouldn't fail because submenus are popup menus, and popup menus can't be
			// menu bars. Update: Even if it does fail due to causing a cascade-destroy upward toward any
			// menu bar that happens to own it, it seems okay because the real purpose here is simply to
			// update that fact that "temp" was already destroyed indirectly by the OS, as evidenced by
			// the fact that IsMenu() returned FALSE above.
			temp->Destroy();
		}
		else
			aMenuItem->mSubmenu = aSubmenu;
	}
	// else no error msg and return OK so that the thread will continue.  This may help catch
	// bugs in the course of normal use of this feature.
	return OK;
}



void UserMenu::UpdateOptions(UserMenuItem *aMenuItem, char *aOptions)
{
	if (toupper(*aOptions) == 'P')
		aMenuItem->mPriority = atoi(aOptions + 1);
}



ResultType UserMenu::RenameItem(UserMenuItem *aMenuItem, char *aNewName)
// Caller should specify "" for aNewName to convert aMenuItem into a separator.
// Returns FAIL if the new name conflicts with an existing name.
{
	if (strlen(aNewName) > MAX_MENU_NAME_LENGTH)
		return FAIL; // Caller should diplay error if desired.

	if (!mMenu) // Just update the member variables for later use when the menu is created.
		return UpdateName(aMenuItem, aNewName);

	MENUITEMINFO mii;
	mii.cbSize = sizeof(mii);
	mii.fMask = MIIM_TYPE;
	mii.dwTypeData = aNewName;

	if (*aNewName)
	{
		// Names must be unique only within each menu:
		for (UserMenuItem *mi = mFirstMenuItem; mi; mi = mi->mNextMenuItem)
			if (!lstrcmpi(mi->mName, aNewName)) // Match found (case insensitive).
				return FAIL; // Caller should display an error message.
		mii.fType = MFT_STRING;
	}
	else // converting into a separator
	{
		// Notes about the below macro:
		// ID_TRAY_OPEN is not set to be the default for the self-contained version, since it lacks that menu item.
		CHANGE_DEFAULT_IF_NEEDED
		mii.fType = MFT_SEPARATOR;
		if (aMenuItem->mSubmenu)  // Converting submenu into a separator.
		{
			mii.fMask |= MIIM_SUBMENU;
			mii.hSubMenu = NULL;
		}
	}

	// Failure is rare enough in the below that no attempt is made to undo the above:
	BOOL result = SetMenuItemInfo(mMenu, aMenuItem_ID, aMenuItem->mSubmenu != NULL, &mii);
	UPDATE_GUI_MENU_BARS(mMenuType, mMenu)  // Verified as being necessary.
	return result ? UpdateName(aMenuItem, aNewName) : FAIL;
}



ResultType UserMenu::UpdateName(UserMenuItem *aMenuItem, char *aNewName)
// Caller should already have ensured that aMenuItem is not too long.
{
	size_t new_length = strlen(aNewName);
	if (new_length)
	{
		if (new_length >= aMenuItem->mNameCapacity) // Too small, so reallocate.
		{
			// Use a temp var. so that mName will never wind up being NULL (relied on by other things).
			// This also retains the original menu name if the allocation fails:
			char *temp = new char[new_length + 1];  // +1 for terminator.
			if (!temp)
				return FAIL;
			// Otherwise:
			if (aMenuItem->mName != Var::sEmptyString) // Since it was previously new'd, delete it.
				delete aMenuItem->mName;
			aMenuItem->mName = temp;
			aMenuItem->mNameCapacity = new_length + 1;
		}
		strcpy(aMenuItem->mName, aNewName);
	}
	else // It will become a separator.
	{
		*aMenuItem->mName = '\0'; // Safe because even if it's capacity is 1 byte, it's a writable byte.
		aMenuItem->mMenuID = 0; // Free up an ID since separators currently can't be converted back into items.
	}
	return OK;
}



ResultType UserMenu::CheckItem(UserMenuItem *aMenuItem)
// Note that items in a menu bar apparently cannot be checked, so it is not necessary to call
// UPDATE_GUI_MENU_BARS here or in the next few functions below.

{
	aMenuItem->mChecked = true;
	if (mMenu)
		CheckMenuItem(mMenu, aMenuItem_ID, aMenuItem_MF_BY | MF_CHECKED);
	return OK;
}



ResultType UserMenu::UncheckItem(UserMenuItem *aMenuItem)
{
	aMenuItem->mChecked = false;
	if (mMenu)
		CheckMenuItem(mMenu, aMenuItem_ID, aMenuItem_MF_BY | MF_UNCHECKED);
	return OK;
}



ResultType UserMenu::ToggleCheckItem(UserMenuItem *aMenuItem)
{
	aMenuItem->mChecked = !aMenuItem->mChecked;
	if (mMenu)
		CheckMenuItem(mMenu, aMenuItem_ID, aMenuItem_MF_BY | (aMenuItem->mChecked ? MF_CHECKED : MF_UNCHECKED));
	return OK;
}



ResultType UserMenu::EnableItem(UserMenuItem *aMenuItem)
{
	aMenuItem->mEnabled = true;
	if (mMenu)
	{
		EnableMenuItem(mMenu, aMenuItem_ID, aMenuItem_MF_BY | MF_ENABLED); // Automatically ungrays it too.
		UPDATE_GUI_MENU_BARS(mMenuType, mMenu)  // Verified as being necessary.
	}
	return OK;

}



ResultType UserMenu::DisableItem(UserMenuItem *aMenuItem)
{
	aMenuItem->mEnabled = false;
	if (mMenu)
	{
		EnableMenuItem(mMenu,aMenuItem_ID, aMenuItem_MF_BY | MF_DISABLED | MF_GRAYED);
		UPDATE_GUI_MENU_BARS(mMenuType, mMenu)  // Verified as being necessary.
	}
	return OK;
}



ResultType UserMenu::ToggleEnableItem(UserMenuItem *aMenuItem)
{
	aMenuItem->mEnabled = !aMenuItem->mEnabled;
	if (mMenu)
	{
		EnableMenuItem(mMenu, aMenuItem_ID, aMenuItem_MF_BY | (aMenuItem->mEnabled ? MF_ENABLED
			: (MF_DISABLED | MF_GRAYED)));
		UPDATE_GUI_MENU_BARS(mMenuType, mMenu)  // Verified as being necessary.
	}
	return OK;
}



ResultType UserMenu::SetDefault(UserMenuItem *aMenuItem)
{
	if (mDefault == aMenuItem)
		return OK;
	mDefault = aMenuItem;
	if (!mMenu) // No further action required: the new setting will be in effect when the menu is created.
		return OK;
	if (aMenuItem) // A user-defined menu item is being made the default.
		SetMenuDefaultItem(mMenu, aMenuItem_ID, aMenuItem->mSubmenu != NULL); // This also ensures that only one is default at a time.
	else
	{
		// Otherwise, a user-defined item that was previously the default is no longer the default.
		// Provide a new default if this is the tray menu, the standard items are present, and a default
		// action is called for:
		if (this == g_script.mTrayMenu) // Necessary for proper operation of the self-contained version:
#ifdef AUTOHOTKEYSC
			SetMenuDefaultItem(mMenu, g_AllowMainWindow && mIncludeStandardItems ? ID_TRAY_OPEN : -1, FALSE);
#else
			SetMenuDefaultItem(mMenu, mIncludeStandardItems ? ID_TRAY_OPEN : -1, FALSE);
#endif
		else
			SetMenuDefaultItem(mMenu, -1, FALSE);
	}
	UPDATE_GUI_MENU_BARS(mMenuType, mMenu)  // Testing shows that menu bars themselves can have default items, and that this is necessary.
	return OK;
}



ResultType UserMenu::IncludeStandardItems()
{
	if (mIncludeStandardItems)
		return OK;
	// In this case, immediately create the menu to support having the standard menu items on the
	// bottom or middle rather than at the top (which is the default). Older comment: Only do
	// this if it was false beforehand so that the standard menu items will be appended to whatever
	// the user has already added to the tray menu (increases flexibility).
	if (!Create()) // It may already exist, in which case this returns OK.
		return FAIL; // No error msg since so rare.
	return AppendStandardItems();
}



ResultType UserMenu::ExcludeStandardItems()
{
	if (!mIncludeStandardItems)
		return OK;
	mIncludeStandardItems = false;
	return Destroy(); // It will be recreated automatically the next time the user displays it.
}



ResultType UserMenu::Create(MenuTypeType aMenuType)
// Menu bars require non-popup menus (CreateMenu vs. CreatePopupMenu).  Rather than maintain two
// different types of HMENUs on the rare chance that a script might try to use a menu both as
// a popup and a menu bar, it seems best to have only one type to keep the code simple and reduce
// resources used for the menu.  This has been documented in the help file.
// Note that a menu bar's submenus can be (perhaps must be) of the popup type, so we only need
// to worry about the distinction for the menu bar itself.  The caller tells us which is which.
{
	if (mMenu)
	{
		// Since menu already exists, check if it's the right type.  If caller left the type unspecified,
		// assume it is the right type:
		if (aMenuType == MENU_TYPE_NONE || aMenuType == mMenuType)
			return OK;
		else // It exists but it's the wrong type.  Destroy and recreate it (but keep TRAY always as popup type).
			if (!stricmp(mName, "tray") || !Destroy()) // Could not be destroyed, perhaps because it is attached to a window as a menu bar.
				return FAIL;
	}
	if (aMenuType == MENU_TYPE_NONE) // Since caller didn't specify and it's about to be (re)created, assume popup.
		aMenuType = MENU_TYPE_POPUP;
	if (   !(mMenu = (aMenuType == MENU_TYPE_BAR) ? CreateMenu() : CreatePopupMenu())   )
		// Failure is rare, so no error msg here (caller can, if it wants).
		return FAIL;

	mMenuType = aMenuType;  // We have to track its type since I don't think there's any way to find out via API.

	// It seems best not to have a mandatory EXIT item added to the bottom of the tray menu
	// for these reasons:
	// 1) Allows the tray icon to be shown even at time when the user wants it to have no menu at all
	//    (i.e. avoids the need for #NoTrayIcon just to disable the showing of the menu).
	// 2) Avoids complexity because there would be a 3rd state: Standard, NoStandard, and
	//    NoStandardWithExit.  This might be inconsequential, but would require testing.
	//if (!mIncludeStandardItems && !mMenuItemCount)
	//{
	//	AppendMenu(mTrayMenu->mMenu, MF_STRING, ID_TRAY_EXIT, "E&xit");
	//	return OK;
	//}

	// By default, the standard menu items are added first, since the users would probably want
	// their own user defined menus at the bottom where they're easier to reach:
	if (mIncludeStandardItems)
		AppendStandardItems();

	// Now append all of the user defined items:
	UINT flags;
	UserMenuItem *mi;
	for (mi = mFirstMenuItem; mi; mi = mi->mNextMenuItem)
	{
		flags = 0;
		if (!*mi->mName)
			flags |= MF_SEPARATOR;  // MF_STRING is the default.
		if (!mi->mEnabled)
			flags |= (MF_DISABLED | MF_GRAYED);  // MF_ENABLED is the default.
		if (mi->mChecked)   // MF_UNCHECKED is the default.
			flags |= MF_CHECKED;
		if (mi->mSubmenu)
		{
			flags |= MF_POPUP;
			// Ensure submenu is created so that handle can be used below.
			if (!mi->mSubmenu->Create())
				return FAIL;
		}
		AppendMenu(mMenu, flags, mi->mSubmenu ? (UINT_PTR)mi->mSubmenu->mMenu : mi->mMenuID, mi->mName);
		
		if (mi->mIcon) // L17: Check mIcon/mBitmap union, apply icon if necessary.
			ApplyItemIcon(mi);
	}
	if (mDefault)
		// This also automatically ensures that only one is default at a time:
		SetMenuDefaultItem(mMenu, mDefault->mMenuID, FALSE);

	// Apply background color if this menu has a non-standard one.  If this menu has submenus,
	// they will be individually given their own background color when created via Create(),
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



void UserMenu::SetColor(char *aColorName, bool aApplyToSubmenus)
{
	// Avoid the overhead of creating HBRUSH's on OSes that don't support SetMenuInfo().
	// Perhaps there is some other way to change menu background color on Win95/NT?
	if (g_os.IsWin95() || g_os.IsWinNT4())
		return;
	AssignColor(aColorName, mColor, mBrush);  // Takes care of deleting old brush, etc.
	// To avoid complications, such as a submenu being detached from its parent and then its parent
	// later being being deleted (which causes the HBRUSH to get deleted too), give each submenu it's
	// own HBRUSH handle by calling AssignColor() for each:
	if (aApplyToSubmenus)
		for (UserMenuItem *mi = mFirstMenuItem; mi; mi = mi->mNextMenuItem)
			if (mi->mSubmenu)
				AssignColor(aColorName, mi->mSubmenu->mColor, mi->mSubmenu->mBrush);
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
	// Must fetch function address dynamically or program won't launch at all on Win95/NT:
	typedef BOOL (WINAPI *MySetMenuInfoType)(HMENU, LPCMENUINFO);
	static MySetMenuInfoType MySetMenuInfo = (MySetMenuInfoType)GetProcAddress(GetModuleHandle("user32"), "SetMenuInfo");
	if (!MySetMenuInfo)
		return;
	MENUINFO mi = {0}; 
	mi.cbSize = sizeof(MENUINFO);
	mi.fMask = MIM_BACKGROUND|(aApplyToSubmenus ? MIM_APPLYTOSUBMENUS : 0);
	mi.hbrBack = mBrush;
	MySetMenuInfo(mMenu, &mi);
}



ResultType UserMenu::AppendStandardItems()
// Caller must ensure that this->mMenu exists if it wants the items to be added immediately.
{
	mIncludeStandardItems = true; // even if the menu doesn't exist.
	if (!mMenu)
		return OK;
#ifdef AUTOHOTKEYSC
	if (g_AllowMainWindow)
	{
		AppendMenu(mMenu, MF_STRING, ID_TRAY_OPEN, "&Open");
		if (this == g_script.mTrayMenu && !mDefault) // No user-defined default menu item, so use the standard one.
			SetMenuDefaultItem(mMenu, ID_TRAY_OPEN, FALSE); // Seems to have no function other than appearance.
	}
#else
	AppendMenu(mMenu, MF_STRING, ID_TRAY_OPEN, "&Open");
	AppendMenu(mMenu, MF_STRING, ID_TRAY_HELP, "&Help");
	AppendMenu(mMenu, MF_SEPARATOR, 0, NULL);
	AppendMenu(mMenu, MF_STRING, ID_TRAY_WINDOWSPY, "&Window Spy");
	AppendMenu(mMenu, MF_STRING, ID_TRAY_RELOADSCRIPT, "&Reload This Script");
	AppendMenu(mMenu, MF_STRING, ID_TRAY_EDITSCRIPT, "&Edit This Script");
	AppendMenu(mMenu, MF_SEPARATOR, 0, NULL);
	if (this == g_script.mTrayMenu && !mDefault) // No user-defined default menu item, so use the standard one.
		SetMenuDefaultItem(mMenu, ID_TRAY_OPEN, FALSE); // Seems to have no function other than appearance.
#endif
	AppendMenu(mMenu, MF_STRING, ID_TRAY_SUSPEND, "&Suspend Hotkeys");
	AppendMenu(mMenu, MF_STRING, ID_TRAY_PAUSE, "&Pause Script");
	AppendMenu(mMenu, MF_STRING, ID_TRAY_EXIT, "E&xit");
	UPDATE_GUI_MENU_BARS(mMenuType, mMenu)  // Verified as being necessary (though it would be rare anyone would want the menu bar containing the std items).
	return OK;  // For caller convenience.
}



ResultType UserMenu::Destroy()
// Returns OK upon complete success or FAIL otherwise.  For example, even if this's menu
// is successfully destroyed, if the indirect destructions resulting from it don't succeed, this
// method returns FAIL.
{
	if (!mMenu)  // For performance.
		return OK;
	// I think DestroyMenu() can fail if an attempt is made to destroy the menu while it is being
	// displayed (but even if it doesn't fail, it seems very bad to try to destroy it then, which
	// is why g_MenuIsVisible is checked just to be sure).
	// But this all should be impossible in our case because the script is in an uninterruptible state
	// while the menu is displayed, which in addition to pausing the current thread (which happens
	// anyway), no new timed or hotkey subroutines can be launched.  Thus, this should rarely if
	// ever happen, which is why no error message is given here:
	//if (g_MenuIsVisible)
	//	return FAIL;

	// DestroyMenu fails (GetLastError() == ERROR_INVALID_MENU_HANDLE) if a parent menu that contained
	// mMenu as one of its submenus was destroyed above.  This seems to indicate that submenus are
	// destroyed whenever a parent menu is destroyed.  Therefore, don't check failure on the below,
	// just assume that afterward, the menu is gone.  IsMenu() is checked because the handle can be
	// invalid if the OS already destroyed it behind-the-scenes (this happens to a submenu whenever
	// its parent menu is destroyed, or whenever a submenu is converted back into a normal menu item):
	if (IsMenu(mMenu))
	{
		// As a precaution, don't allow a menu to be destroyed if a window is using it as its
		// menu bar. That might have bad side-effects on some OSes, especially older ones:
		if (mMenuType == MENU_TYPE_BAR && GuiType::sGuiCount)
		{
			int i, gui_count;
			for (i = 0, gui_count = 0; i < MAX_GUI_WINDOWS; ++i)
				if (g_gui[i])
				{
					if (g_gui[i]->mHwnd && GetMenu(g_gui[i]->mHwnd) == mMenu)
						return FAIL; // A GUI window is using this menu, so don't destroy the menu.
					if (GuiType::sGuiCount == ++gui_count) // No need to keep searching.
						break;
				}
		}
		if (!DestroyMenu(mMenu)) // v1.0.30.01: Doesn't seem to be a reason *not* to check the return value and return FAIL if it failed.
			return FAIL;
	}
	mMenu = NULL; // This must be done immediately after destroying the menu to prevent recursion problems below.

	// Bug-fix for v1.0.19: The below is now done OUTSIDE the above block because the moment a
	// parent menu is deleted all its submenus AND SUB-SUB-SUB...MENUS become invalid menu handles.
	// But even though the OS has done this, Destroy() must still be called recursively from here
	// so that the menu handles will be set to NULL.  This is because other functions -- such as
	// Display() -- do not do the IsMenu() check, relying instead on whether the handle is NULL to
	// determine whether the menu physically exists.
	// The moment the above is done, any submenus that were attached to mMenu are also destroyed
	// by the OS.  So mark them as destroyed in our bookkeeping also:
	UserMenuItem *mi;
	for (mi = mFirstMenuItem; mi ; mi = mi->mNextMenuItem)
		if (mi->mSubmenu && mi->mSubmenu->mMenu && !IsMenu(mi->mSubmenu->mMenu))
			mi->mSubmenu->Destroy(); // Its return value isn't checked since there doesn't seem to be anything that can/should be done if it fails.

	// Destroy any menu that contains this menu as a submenu.  This is done so that such
	// menus will be automatically recreated the next time they are used, which is necessary
	// because otherwise when such a menu is displayed the next time, the OS will show its
	// old contents even though the menu is gone.  Thus, those old menu items will be
	// selectable but will have no effect.  In addition, sometimes our caller plans to
	// recreate this->mMenu (or have it recreated automatically upon first use) and thus
	// we don't want to use DeleteMenu() because that would require having to detect whether
	// the menu needs updating (to reflect whether the submenu has been recreated) every
	// time we display it.  Another drawback to DeleteMenu() is that it would change the
	// order of the menu items to something other than what the user originally specified
	// unless InsertMenu() was woven in during the update:
	ResultType result = OK;
	for (UserMenu *m = g_script.mFirstMenu; m; m = m->mNextMenu)
		if (m->mMenu)
			for (mi = m->mFirstMenuItem; mi; mi = mi->mNextMenuItem)
				if (mi->mSubmenu == this)
					if (!m->Destroy())  // Attempt to destroy any menu that contains this menu as a submenu (will fail if m is a menu bar).
						result = FAIL; // Seems best to consider even one failure is considered a total failure.
	return result;
}



ResultType UserMenu::Display(bool aForceToForeground, int aX, int aY)
// aForceToForeground defaults to true because when a menu is displayed spontanesouly rather than
// in response to the user right-clicking the tray icon, I believe that the OS will revert to its
// behavior of "resisting" a window that tries to "steal focus".  I believe this resistance does
// not occur when the user clicks the icon because that click causes the task bar to get focus,
// and it is likely that the OS allows other windows to steal focus from the task bar without
// resistence.  This is done because if the main window is *not* successfully activated prior to
// displaying the menu, it might be impossible to dismiss the menu by clicking outside of it.
{
	if (!mMenuItemCount && !mIncludeStandardItems)
		return OK;  // Consider the display of an empty menu to be a success.
	//if (!IsMenu(mMenu))
	//	mMenu = NULL;
	if (!mMenu) // i.e. because this is the first time the user has opened the menu.
		if (!Create()) // no error msg since so rare
			return FAIL;
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
		if (aX != COORD_UNSPECIFIED)
			pt.x = aX;
		if (aY != COORD_UNSPECIFIED)
			pt.y = aY;
		if (!(g->CoordMode & COORD_MODE_MENU))  // Using coords relative to the active window (rather than screen).
			WindowToScreen((int &)pt.x, (int &)pt.y);
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
			// 0 is exactly the amount of time (-1 is not enough because it doesn't yeild) needed for that
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
	// You could argue that selecting a menu item should immediately Gosub the selected menu item's
	// subroutine rather than queuing it up as a new thread.  However, even if that is a better method,
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



UINT UserMenu::GetSubmenuPos(HMENU ahMenu)
// ahMenu will be searched for in this->mMenu.
// Returns UINT_MAX if this->mMenu is NULL, ahMenu is NULL, or if ahMenu can't be found in this->mMenu.
// Testing shows that neither ModifyMenu() nor SetMenuItemInfo() nor any other Menu modifying
// API call will accept aMenuItem->mMenuID or aMenuItem->mSubmenu as a way to uniquely
// indentify the menu item we want to change (even though GetMenuItemInfo indicates that
// aMenuItem->mSubmenu is the "ID" of sorts).  Thus, the menu item can only be modified
// by position.  Rather than having to maintain each submenu's position in every menu, thus
// making the code less maintainable since you always have to worry about standard menu items
// being added to the top vs. bottom of the menu, other menu items being inserted via
// InsertMenu() (if that is ever allowed), etc. -- just loop through the menu to find the
// right item, then return that to the caller so that it can modify the submenu based on position:
{
	if (!ahMenu || !mMenu)
		return UINT_MAX;
	int menu_item_count = GetMenuItemCount(mMenu);
	for (int i = 0; i < menu_item_count; ++i)
		if (GetSubMenu(mMenu, i) == ahMenu)
			return i;
	return UINT_MAX;
}



UINT UserMenu::GetItemPos(char *aMenuItemName)
// aMenuItemName will be searched for in this->mMenu.
// Returns UINT_MAX if this->mMenu is NULL or if aMenuItemName can't be found in this->mMenu.
{
	if (!mMenu)
		return UINT_MAX;
	int menu_item_count = GetMenuItemCount(mMenu);
	char buf[MAX_MENU_NAME_LENGTH + 2];  // +2 due to uncertainty over whether GetMenuString()'s nMaxCount includes room for terminator.
	for (int i = 0; i < menu_item_count; ++i)
		if (GetMenuString(mMenu, i, buf, sizeof(buf) - 1, MF_BYPOSITION))
			if (!lstrcmpi(buf, aMenuItemName))  // A case insensitive match was found.
				return i;
	return UINT_MAX;  // No match found.
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


// L17: Menu-item icon functions.

ResultType UserMenu::SetItemIcon(UserMenuItem *aMenuItem, char *aFilename, int aIconNumber, int aWidth)
{
	if (!*aFilename || (*aFilename == '*' && !aFilename[1]))
		return RemoveItemIcon(aMenuItem);

	// L29: The bitmap/icon returned by LoadPicture is converted to the appropriate format automatically,
	// so the following is no longer necessary:
	//if (aIconNumber == 0)
	//	aIconNumber = 1; // Must be != 0 to tell LoadPicture that "icon must be loaded, never a bitmap".

	int image_type;
	HICON new_icon;
	// Currently height is always -1 and cannot be overriden. -1 means maintain aspect ratio, usually 1:1 for icons.
	if ( !(new_icon = (HICON)LoadPicture(aFilename, aWidth, -1, image_type, aIconNumber, false)) )
		return FAIL;

	HBITMAP new_copy;

	if (g_os.IsWinVistaOrLater())
	{
		if (image_type == IMAGE_ICON) // Convert to 32-bit bitmap:
		{
			new_copy = IconToBitmap32(new_icon, true);
			// Even if conversion failed, we have no further use for the icon:
			DestroyIcon(new_icon);
			if (!new_copy)
				return FAIL;
			new_icon = (HICON)new_copy;
		}

		if (aMenuItem->mBitmap) // Delete previous bitmap.
			DeleteObject(aMenuItem->mBitmap);
	}
	else
	{
		if (image_type == IMAGE_BITMAP) // Convert to icon:
		{
			ICONINFO iconinfo;
			iconinfo.fIcon = TRUE;
			iconinfo.hbmMask = (HBITMAP)new_icon;
			iconinfo.hbmColor = (HBITMAP)new_icon;
			new_copy = (HBITMAP)CreateIconIndirect(&iconinfo);
			// Even if conversion failed, we have no further use for the bitmap:
			DeleteObject((HBITMAP)new_icon);
			if (!new_copy)
				return FAIL;
			new_icon = (HICON)new_copy;
		}

		if (aMenuItem->mIcon) // Delete previous icon.
			DestroyIcon(aMenuItem->mIcon);
	}
	// Also sets mBitmap via union:
	aMenuItem->mIcon = new_icon;

	if (mMenu)
		ApplyItemIcon(aMenuItem);

	return aMenuItem->mIcon ? OK : FAIL;
}

// Caller has ensured mMenu is non-NULL.
ResultType UserMenu::ApplyItemIcon(UserMenuItem *aMenuItem)
{
	if (aMenuItem->mIcon) // Check mIcon/mBitmap union.
	{
		MENUITEMINFO item_info;
		item_info.cbSize = sizeof(MENUITEMINFO);
		item_info.fMask = MIIM_BITMAP;
		// Set HBMMENU_CALLBACK or 32-bit bitmap as appropriate.
		item_info.hbmpItem = g_os.IsWinVistaOrLater() ? aMenuItem->mBitmap : HBMMENU_CALLBACK;
		SetMenuItemInfo(mMenu, aMenuItem_ID, aMenuItem_MF_BY, &item_info);
	}
	return OK;
}

ResultType UserMenu::RemoveItemIcon(UserMenuItem *aMenuItem)
{
	if (aMenuItem->mIcon) // Check mIcon/mBitmap union.
	{
		if (mMenu)
		{
			MENUITEMINFO item_info;
			item_info.cbSize = sizeof(MENUITEMINFO);
			item_info.fMask = MIIM_BITMAP;
			item_info.hbmpItem = NULL;
			// If g_os.IsWinVistaOrLater(), this removes the bitmap we set. Otherwise it removes HBMMENU_CALLBACK, therefore disabling owner-drawing.
			SetMenuItemInfo(mMenu, aMenuItem_ID, aMenuItem_MF_BY, &item_info);
		}
		if (g_os.IsWinVistaOrLater()) // Free the appropriate union member.
			DeleteObject(aMenuItem->mBitmap);
		else
			DestroyIcon(aMenuItem->mIcon);
		aMenuItem->mIcon = NULL; // Clear mIcon/mBitmap union.
	}
	return OK;
}
