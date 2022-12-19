///////////////////////////////////////////////////////////////////////////////
//
// AutoIt
//
// Copyright (C)1999-2006:
//		- Jonathan Bennett <jon@hiddensoft.com>
//		- Others listed at https://www.autoitscript.com/autoit3/docs/credits.htm
//      - Chris Mallett (support@autohotkey.com): various enhancements and
//        adaptation of this file's functions to interface with AutoHotkey.
//
// This program is free software; you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation; either version 2 of the License, or
// (at your option) any later version.
//
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h" // pre-compiled headers
//#include <winsock.h>  // for WSADATA.  This also requires wsock32.lib to be linked in.
#include <winsock2.h>
#include <tlhelp32.h> // For the ProcessExist routines.
#include <wininet.h> // For Download().
#include "script.h"
#include "globaldata.h"
#include "window.h" // For ControlExist().
#include "application.h" // For SLEEP_WITHOUT_INTERRUPTION and MsgSleep().
#include "script_func_impl.h"
#include "abi.h"


ResultType Script::DoRunAs(LPTSTR aCommandLine, LPCTSTR aWorkingDir, bool aDisplayErrors, WORD aShowWindow
	, PROCESS_INFORMATION &aPI, bool &aSuccess // Output parameters we set for caller, but caller must have initialized aSuccess to false.
	, HANDLE &aNewProcess, DWORD &aLastError)  // Same, but initialize to NULL.
{
	typedef BOOL (WINAPI *MyCreateProcessWithLogonW)(
		LPCWSTR lpUsername,                 // user's name
		LPCWSTR lpDomain,                   // user's domain
		LPCWSTR lpPassword,                 // user's password
		DWORD dwLogonFlags,                 // logon option
		LPCWSTR lpApplicationName,          // executable module name
		LPWSTR lpCommandLine,               // command-line string
		DWORD dwCreationFlags,              // creation flags
		LPVOID lpEnvironment,               // new environment block
		LPCWSTR lpCurrentDirectory,         // current directory name
		LPSTARTUPINFOW lpStartupInfo,       // startup information
		LPPROCESS_INFORMATION lpProcessInfo // process information
		);
	// Set up wide char version that we need for CreateProcessWithLogon
	// init structure for running programs (wide char version)
	STARTUPINFOW wsi = {0};
	wsi.cb			= sizeof(STARTUPINFOW);
	wsi.dwFlags		= STARTF_USESHOWWINDOW;
	wsi.wShowWindow = aShowWindow;
	// The following are left initialized to 0/NULL (initialized earlier above):
	//wsi.lpReserved = NULL;
	//wsi.lpDesktop	= NULL;
	//wsi.lpTitle = NULL;
	//wsi.cbReserved2 = 0;
	//wsi.lpReserved2 = NULL;

#ifndef UNICODE
	// Convert to wide character format:
	WCHAR command_line_wide[LINE_SIZE], working_dir_wide[MAX_PATH];
	ToWideChar(aCommandLine, command_line_wide, LINE_SIZE); // Dest. size is in wchars, not bytes.
	if (aWorkingDir && *aWorkingDir)
		ToWideChar(aWorkingDir, working_dir_wide, MAX_PATH); // Dest. size is in wchars, not bytes.
	else
		*working_dir_wide = 0;  // wide-char terminator.

	if (lpfnDLLProc(mRunAsUser, mRunAsDomain, mRunAsPass, LOGON_WITH_PROFILE, 0
		, command_line_wide, 0, 0, *working_dir_wide ? working_dir_wide : NULL, &wsi, &aPI))
#else
	if (CreateProcessWithLogonW(mRunAsUser, mRunAsDomain, mRunAsPass, LOGON_WITH_PROFILE, 0
		, aCommandLine, 0, 0, aWorkingDir && *aWorkingDir ? aWorkingDir : NULL, &wsi, &aPI))
#endif
	{
		aSuccess = true;
		if (aPI.hThread)
			CloseHandle(aPI.hThread); // Required to avoid memory leak.
		aNewProcess = aPI.hProcess;
	}
	else
		aLastError = GetLastError(); // Caller will use this to get an error message and set g->LastError if needed.
	return OK;
}



bif_impl void RunAs(optl<StrArg> aUser, optl<StrArg> aPass, optl<StrArg> aDomain)
{
	StringTCharToWChar(aUser.value_or_empty(), g_script.mRunAsUser);
	StringTCharToWChar(aPass.value_or_empty(), g_script.mRunAsPass);
	StringTCharToWChar(aDomain.value_or_empty(), g_script.mRunAsDomain);
}



bif_impl FResult SysGetIPAddresses(IObject *&aRetVal)
{
	// aaa.bbb.ccc.ddd = 15, but allow room for larger IP's in the future.
	#define IP_ADDRESS_SIZE 32 // The maximum size of any of the strings we return, including terminator.

	auto addresses = Array::Create();
	if (!addresses)
		return FR_E_OUTOFMEM;

	WSADATA wsadata;
	if (WSAStartup(MAKEWORD(1, 1), &wsadata)) // Failed (it returns 0 on success).
	{
		aRetVal = addresses;
		return OK;
	}

	char host_name[256];
	gethostname(host_name, _countof(host_name));
	HOSTENT *lpHost = gethostbyname(host_name);

	for (int i = 0; lpHost->h_addr_list[i]; ++i)
	{
		IN_ADDR inaddr;
		memcpy(&inaddr, lpHost->h_addr_list[i], 4);
#ifdef UNICODE
		CStringTCharFromChar addr_buf(inet_ntoa(inaddr));
		LPTSTR addr_str = const_cast<LPTSTR>(addr_buf.GetString());
#else
		LPTSTR addr_str = inet_ntoa(inaddr);
#endif
		if (!addresses->Append(addr_str))
		{
			addresses->Release();
			WSACleanup();
			return FR_E_OUTOFMEM;
		}
	}

	WSACleanup();
	aRetVal = addresses;
	return OK;
}



BIV_DECL_R(BIV_IsAdmin)
{
	bool result = false;  // Default.
	SC_HANDLE h = OpenSCManager(NULL, NULL, SC_MANAGER_LOCK);
	if (h)
	{
		SC_LOCK lock = LockServiceDatabase(h);
		if (lock)
		{
			UnlockServiceDatabase(lock);
			result = true; // Current user is admin.
		}
		else
		{
			DWORD lastErr = GetLastError();
			if (lastErr == ERROR_SERVICE_DATABASE_LOCKED)
				result = true; // Current user is admin.
		}
		CloseServiceHandle(h);
	}
	_f_return_b(result);
}



bif_impl FResult PixelGetColor(int aX, int aY, optl<StrArg> aMode, StrRet &aRetVal)
{
	LPTSTR buf = aRetVal.CallerBuf();
	aRetVal.SetTemp(buf, 8);
	
	auto aOptions = aMode.value_or_empty();

	if (tcscasestr(aOptions, _T("Slow"))) // New mode for v1.0.43.10.  Takes precedence over Alt mode.
	{
		return PixelSearch(nullptr, nullptr, nullptr, aX, aY, aX, aY, 0, 0, buf); // It takes care of setting the return value.
	}

	CoordToScreen(aX, aY, COORD_MODE_PIXEL);
	
	bool use_alt_mode = tcscasestr(aOptions, _T("Alt")) != NULL; // New mode for v1.0.43.10: Two users reported that CreateDC works better in certain windows such as SciTE, at least one some systems.
	HDC hdc = use_alt_mode ? CreateDC(_T("DISPLAY"), NULL, NULL, NULL) : GetDC(NULL);
	if (!hdc)
		return FR_E_WIN32;

	// Assign the value as an 32-bit int to match Window Spy reports color values.
	// Update for v1.0.21: Assigning in hex format seems much better, since it's easy to
	// look at a hex BGR value to get some idea of the hue.  In addition, the result
	// is zero padded to make it easier to convert to RGB and more consistent in
	// appearance:
	COLORREF color = GetPixel(hdc, aX, aY);
	if (use_alt_mode)
		DeleteDC(hdc);
	else
		ReleaseDC(NULL, hdc);

	_stprintf(buf, _T("0x%06X"), bgr_to_rgb(color));
	return OK;
}



bif_impl FResult MenuSelect(ExprTokenType *aWinTitle, optl<StrArg> aWinText, StrArg aMenu
	, optl<StrArg> a1, optl<StrArg> a2, optl<StrArg> a3, optl<StrArg> a4, optl<StrArg> a5, optl<StrArg> a6
	, optl<StrArg> aExcludeTitle, optl<StrArg> aExcludeText)
{
	HWND target_window;
	DETERMINE_TARGET_WINDOW;
	
	LPCTSTR names[] { aMenu, a1.value_or_null(), a2.value_or_null(), a3.value_or_null(), a4.value_or_null(), a5.value_or_null(), a6.value_or_null() };
	int name_index = 0, name_count = 0;
	// Count up to the last present name parameter.
	for (int i = 0; i < _countof(names); ++i)
		if (names[i])
			name_count = i + 1;
	// Any omitted parameters preceding name_count will be detected by the main loop as invalid.

	UINT message = WM_COMMAND;
	HMENU hMenu;
	if (!_tcsicmp(names[0], _T("0&")))
	{
		hMenu = GetSystemMenu(target_window, FALSE);
		name_index += 1;
		message = WM_SYSCOMMAND;
	}
	else
		hMenu = GetMenu(target_window);
	if (!hMenu) // Window has no menu bar.
		return FError(ERR_WINDOW_HAS_NO_MENU, nullptr, ErrorPrototype::Target);

	int menu_item_count = GetMenuItemCount(hMenu);
	//if (menu_item_count < 1) // Menu bar has no menus.
	//	goto error;
	
#define MENU_ITEM_IS_SUBMENU 0xFFFFFFFF
#define UPDATE_MENU_VARS(menu_pos) \
menu_id = GetMenuItemID(hMenu, menu_pos);\
if (menu_id == MENU_ITEM_IS_SUBMENU)\
	menu_item_count = GetMenuItemCount(hMenu = GetSubMenu(hMenu, menu_pos));\
else\
{\
	menu_item_count = 0;\
	hMenu = NULL;\
}

	UINT menu_id = MENU_ITEM_IS_SUBMENU;
	TCHAR menu_text[1024];
	bool match_found;
	size_t this_menu_param_length, menu_text_length;
	int pos, target_menu_pos;
	LPCTSTR this_menu_param = nullptr;

	for ( ; name_index < name_count; ++name_index)
	{
		if (!hMenu)  // The nesting of submenus ended prior to the end of the list of menu search terms.
			break;
		this_menu_param = names[name_index];
		if (!this_menu_param || !*this_menu_param)
			break;

		this_menu_param_length = _tcslen(this_menu_param);
		target_menu_pos = (this_menu_param[this_menu_param_length - 1] == '&') ? ATOI(this_menu_param) - 1 : -1;
		if (target_menu_pos > -1)
		{
			if (target_menu_pos >= menu_item_count)  // Invalid menu position (doesn't exist).
				break;
			UPDATE_MENU_VARS(target_menu_pos)
		}
		else // Searching by text rather than numerical position.
		{
			for (match_found = false, pos = 0; pos < menu_item_count; ++pos)
			{
				menu_text_length = GetMenuString(hMenu, pos, menu_text, _countof(menu_text) - 1, MF_BYPOSITION);
				// v1.0.43.03: It's debatable, but it seems best to support locale's case insensitivity for
				// menu items, since menu names tend to adapt to the user's locale.  By contrast, things
				// like process names (in the Process command) do not tend to change, so it seems best to
				// have them continue to use stricmp(): 1) avoids breaking existing scripts; 2) provides
				// consistent behavior across multiple locales; 3) performance.
				match_found = !lstrcmpni(menu_text  // This call is basically a strnicmp() that obeys locale.
					, menu_text_length > this_menu_param_length ? this_menu_param_length : menu_text_length
					, this_menu_param, this_menu_param_length);
				//match_found = strcasestr(menu_text, this_menu_param);
				TCHAR *cpamp;
				if (!match_found && (cpamp = _tcschr(menu_text, '&')))
				{
					// Try again to find a match, this time without the ampersands used to indicate a menu item's
					// shortcut key.  One might assume that only the first & needs to be removed, but the actual
					// behaviour confirmed on Windows 10.0.19041 was:
					//  - Only every second & in a series of consecutive &&s was kept.
					//  - Only the *first* &-letter was usable as a shortcut key.
					//  - Only the *last* & caused an underline.
					for (TCHAR *cplit = cpamp; ; ++cpamp, ++cplit)
					{
						if (*cpamp == '&')
							++cpamp; // Skip this '&' but copy the next character unconditionally, even if it is '&'.
						*cplit = *cpamp;
						if (!*cpamp)
							break; // Only after copying, so menu_text is null-terminated at the correct position (cpw).
					}
					menu_text_length = _tcslen(menu_text);
					match_found = !lstrcmpni(menu_text  // This call is basically a strnicmp() that obeys locale.
						, menu_text_length > this_menu_param_length ? this_menu_param_length : menu_text_length
						, this_menu_param, this_menu_param_length);
					//match_found = strcasestr(menu_text, this_menu_param);
				}
				if (match_found)
				{
					UPDATE_MENU_VARS(pos)
					break;
				}
			} // inner for()
			if (!match_found) // The search hierarchy (nested menus) specified in the params could not be found.
				break;
		} // else
	} // outer for()

	if (name_index < name_count)
		return FR_E_ARG(2 + name_index);
	// This would happen if the outer loop above had zero iterations due to aMenu1 being NULL or blank,
	// or if the caller specified a submenu as the target (which doesn't seem valid since an app would
	// never expect to ever receive a message for a submenu?):
	if (menu_id == MENU_ITEM_IS_SUBMENU)
		return FR_E_ARG(2 + name_count - 1);

	// Since the above didn't return, the specified search hierarchy was completely found.
	PostMessage(target_window, message, (WPARAM)menu_id, 0);
	return OK;
}



bif_impl FResult ControlSetChecked(int aValue, CONTROL_PARAMETERS_DECL)
{
	if (aValue < -1 || aValue > 1)
		return FR_E_ARG(0);
	DETERMINE_TARGET_CONTROL2;
	DWORD_PTR dwResult;
	if (aValue != -1)
	{
		if (!SendMessageTimeout(control_window, BM_GETCHECK, 0, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return FR_E_WIN32;
		if (dwResult == (aValue ? BST_CHECKED : BST_UNCHECKED)) // It's already in the right state, so don't press it.
			return OK;
	}
	// MSDN docs for BM_CLICK (and au3 author says it applies to this situation also):
	// "If the button is in a dialog box and the dialog box is not active, the BM_CLICK message
	// might fail. To ensure success in this situation, call the SetActiveWindow function to activate
	// the dialog box before sending the BM_CLICK message to the button."
	ATTACH_THREAD_INPUT
	SetActiveWindow(target_window == control_window ? GetNonChildParent(control_window) : target_window); // v1.0.44.13: Fixed to allow for the fact that target_window might be the control itself (e.g. via ahk_id %ControlHWND%).
	RECT rect;
	if (!GetWindowRect(control_window, &rect))	// au3: Code to primary click the centre of the control
		rect.bottom = rect.left = rect.right = rect.top = 0;
	auto lparam = MAKELPARAM((rect.right - rect.left) / 2, (rect.bottom - rect.top) / 2);
	FResult fr = (
			PostMessage(control_window, WM_LBUTTONDOWN, MK_LBUTTON, lparam) &&
			PostMessage(control_window, WM_LBUTTONUP, 0, lparam)
		) ? OK : FR_E_WIN32(GetLastError());
	DETACH_THREAD_INPUT
	DoControlDelay;
	return fr;
}



bif_impl FResult ControlGetChecked(CONTROL_PARAMETERS_DECL, BOOL &aRetVal)
{
	DETERMINE_TARGET_CONTROL2;
	DWORD_PTR dwResult;
	if (!SendMessageTimeout(control_window, BM_GETCHECK, 0, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
		return FR_E_WIN32;
	aRetVal = dwResult == BST_CHECKED;
	return OK;
}



FResult WinSetStyle(StrArg aValue, CONTROL_PARAMETERS_DECL_OPT, int style_index);

bif_impl FResult ControlSetStyle(StrArg aValue, CONTROL_PARAMETERS_DECL)
{
	return WinSetStyle(aValue, &CONTROL_PARAMETERS, GWL_STYLE);
}

bif_impl FResult ControlSetExStyle(StrArg aValue, CONTROL_PARAMETERS_DECL)
{
	return WinSetStyle(aValue, &CONTROL_PARAMETERS, GWL_EXSTYLE);
}



static FResult ControlGetLong(CONTROL_PARAMETERS_DECL, UINT &aRetVal, int nIndex)
{
	DETERMINE_TARGET_CONTROL2;
	aRetVal = GetWindowLong(control_window, nIndex);
	return OK;
}

bif_impl FResult ControlGetStyle(CONTROL_PARAMETERS_DECL, UINT &aRetVal)
{
	return ControlGetLong(CONTROL_PARAMETERS, aRetVal, GWL_STYLE);
}

bif_impl FResult ControlGetExStyle(CONTROL_PARAMETERS_DECL, UINT &aRetVal)
{
	return ControlGetLong(CONTROL_PARAMETERS, aRetVal, GWL_EXSTYLE);
}



static FResult ControlShowDropDown(CONTROL_PARAMETERS_DECL, BOOL aShow)
{
	DETERMINE_TARGET_CONTROL2;
	DWORD_PTR dwResult;
	// CB_SHOWDROPDOWN: Although the return value (dwResult) is always TRUE, SendMessageTimeout()
	// will return failure if it times out:
	if (!SendMessageTimeout(control_window, CB_SHOWDROPDOWN
		, (WPARAM)aShow, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
		return FR_E_WIN32;
	DoControlDelay;
	return OK;
}

bif_impl FResult ControlShowDropDown(CONTROL_PARAMETERS_DECL)
{
	return ControlShowDropDown(CONTROL_PARAMETERS, TRUE);
}

bif_impl FResult ControlHideDropDown(CONTROL_PARAMETERS_DECL)
{
	return ControlShowDropDown(CONTROL_PARAMETERS, FALSE);
}



enum IndexControlType {
	IC_INVALID = 0,
	IC_COMBO,
	IC_LIST,
	IC_TAB
};

static IndexControlType GetIndexControlType(HWND aControl, FResult &fr, bool aAllowTab = false)
{
	TCHAR classname[256 + 1];
	GetClassName(aControl, classname, _countof(classname));
	if (tcscasestr(classname, _T("Combo")))
		return IC_COMBO;
	else if (tcscasestr(classname, _T("List")))
		return IC_LIST;
	else if (aAllowTab && tcscasestr(classname, _T("Tab")))
		return IC_TAB;
	fr = FError(ERR_GUI_NOT_FOR_THIS_TYPE, classname, ErrorPrototype::Target);
	return IC_INVALID;
}



bif_impl FResult ControlAddItem(StrArg aItem, CONTROL_PARAMETERS_DECL, INT_PTR &aRetVal)
{
	DETERMINE_TARGET_CONTROL2;
	UINT msg;
	FResult fr = FAIL;
	switch (GetIndexControlType(control_window, fr))
	{
	case IC_COMBO: msg = CB_ADDSTRING; break;
	case IC_LIST: msg = LB_ADDSTRING; break;
	default: return fr;
	}
	DWORD_PTR dwResult;
	if (!SendMessageTimeout(control_window, msg, 0, (LPARAM)aItem, SMTO_ABORTIFHUNG, 2000, &dwResult))
		return FR_E_WIN32;
	if (dwResult == CB_ERR || dwResult == CB_ERRSPACE) // General error or insufficient space to store it.
		// CB_ERR == LB_ERR
		return FR_E_FAILED;
	DoControlDelay;
	aRetVal = dwResult + 1; // Return the one-based index of the new item.
	return OK;
}



bif_impl FResult ControlDeleteItem(INT_PTR aIndex, CONTROL_PARAMETERS_DECL)
{
	if (--aIndex < 0)
		return FR_E_ARG(0);
	DETERMINE_TARGET_CONTROL2;
	UINT msg;
	FResult fr = FAIL;
	switch (GetIndexControlType(control_window, fr))
	{
	case IC_COMBO: msg = CB_DELETESTRING; break;
	case IC_LIST: msg = LB_DELETESTRING; break;
	default: return fr;
	}
	DWORD_PTR dwResult;
	if (!SendMessageTimeout(control_window, msg, (WPARAM)aIndex, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
		return FR_E_WIN32;
	if (dwResult == CB_ERR)  // CB_ERR == LB_ERR
		return FR_E_FAILED;
	DoControlDelay;
	return OK;
}



static FResult ControlNotifyParent(HWND control_window, UINT x_msg, UINT y_msg)
{
	HWND immediate_parent;
	if (   !(immediate_parent = GetParent(control_window))   )
		return FR_E_WIN32;
	SetLastError(0); // Must be done to differentiate between success and failure when control has ID 0.
	auto control_id = GetDlgCtrlID(control_window);
	if (!control_id && GetLastError()) // Both conditions must be checked (see above).
		return FR_E_WIN32; // Avoid sending the notification in case some other control has ID 0.
	// Proceed even if control_id == 0, since some applications are known to
	// utilize the notification in that case (e.g. Notepad's Save As dialog).
	DWORD_PTR dwResult;
	if (!SendMessageTimeout(immediate_parent, WM_COMMAND, (WPARAM)MAKELONG(control_id, x_msg)
		, (LPARAM)control_window, SMTO_ABORTIFHUNG, 2000, &dwResult))
		return FR_E_WIN32;
	if (!SendMessageTimeout(immediate_parent, WM_COMMAND, (WPARAM)MAKELONG(control_id, y_msg)
		, (LPARAM)control_window, SMTO_ABORTIFHUNG, 2000, &dwResult))
		return FR_E_WIN32;
	DoControlDelay;
	return OK;
}



bif_impl FResult ControlChooseIndex(INT_PTR aIndex, CONTROL_PARAMETERS_DECL)
{
	if (--aIndex < -1)
		return FR_E_ARG(0);
	DETERMINE_TARGET_CONTROL2;
	UINT msg, x_msg, y_msg;
	FResult fr = FAIL;
	switch (GetIndexControlType(control_window, fr, true))
	{
	case IC_COMBO:
		msg = CB_SETCURSEL;
		x_msg = CBN_SELCHANGE;
		y_msg = CBN_SELENDOK;
		break;
	case IC_LIST:
		if (GetWindowLong(control_window, GWL_STYLE) & (LBS_EXTENDEDSEL|LBS_MULTIPLESEL))
			msg = LB_SETSEL;
		else // single-select listbox
			msg = LB_SETCURSEL;
		x_msg = LBN_SELCHANGE;
		y_msg = LBN_DBLCLK;
		break;
	case IC_TAB:
		if (aIndex < 0)
			return FR_E_ARG(0);
		return ControlSetTab(control_window, (DWORD)aIndex);
	default:
		return fr;
	}

	DWORD_PTR dwResult;
	if (msg == LB_SETSEL) // Multi-select, so use the cumulative method.
	{
		if (!SendMessageTimeout(control_window, msg, aIndex != -1, aIndex, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return FR_E_WIN32;
	}
	else // ComboBox or single-select ListBox.
		if (!SendMessageTimeout(control_window, msg, aIndex, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return FR_E_WIN32;
	if (dwResult == CB_ERR && aIndex != -1)  // CB_ERR == LB_ERR
		return FR_E_FAILED;

	return ControlNotifyParent(control_window, x_msg, y_msg);
}



bif_impl FResult ControlChooseString(StrArg aValue, CONTROL_PARAMETERS_DECL, INT_PTR &aRetVal)
{
	DETERMINE_TARGET_CONTROL2;
	UINT msg, x_msg, y_msg;
	FResult fr = FAIL;
	switch (GetIndexControlType(control_window, fr))
	{
	case IC_COMBO:
		msg = CB_SELECTSTRING;
		x_msg = CBN_SELCHANGE;
		y_msg = CBN_SELENDOK;
		break;
	case IC_LIST:
		if (GetWindowLong(control_window, GWL_STYLE) & (LBS_EXTENDEDSEL|LBS_MULTIPLESEL))
			msg = LB_FINDSTRING;
		else // single-select listbox
			msg = LB_SELECTSTRING;
		x_msg = LBN_SELCHANGE;
		y_msg = LBN_DBLCLK;
		break;
	default:
		return fr;
	}
	DWORD_PTR dwResult, item_index;
	if (msg == LB_FINDSTRING) // Multi-select ListBox (LB_SELECTSTRING is not supported by these).
	{
		if (!SendMessageTimeout(control_window, msg, -1, (LPARAM)aValue, SMTO_ABORTIFHUNG, 2000, &item_index))
			return FR_E_WIN32;
		if (item_index == LB_ERR)
			return FR_E_FAILED;
		if (!SendMessageTimeout(control_window, LB_SETSEL, TRUE, item_index, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return FR_E_WIN32;
		if (dwResult == LB_ERR)
			return FR_E_FAILED;
	}
	else // ComboBox or single-select ListBox.
	{
		if (!SendMessageTimeout(control_window, msg, -1, (LPARAM)aValue, SMTO_ABORTIFHUNG, 2000, &item_index))
			return FR_E_WIN32;
		if (item_index == CB_ERR) // CB_ERR == LB_ERR
			return FR_E_FAILED;
	}
	aRetVal = item_index + 1; // Return the index chosen.  Might have some use if the string was ambiguous.
	return ControlNotifyParent(control_window, x_msg, y_msg);
}



bif_impl FResult EditPaste(StrArg aValue, CONTROL_PARAMETERS_DECL)
{
	DETERMINE_TARGET_CONTROL2;
	DWORD_PTR dwResult;
	if (!SendMessageTimeout(control_window, EM_REPLACESEL, TRUE, (LPARAM)aValue, SMTO_ABORTIFHUNG, 2000, &dwResult))
		return FR_E_WIN32;
	DoControlDelay;
	return OK;
}



bif_impl FResult ControlFindItem(StrArg aString, CONTROL_PARAMETERS_DECL, INT_PTR &aRetVal)
{
	DETERMINE_TARGET_CONTROL2;
	UINT msg;
	FResult fr = FAIL;
	switch (GetIndexControlType(control_window, fr))
	{
	case IC_COMBO: msg = CB_FINDSTRINGEXACT; break;
	case IC_LIST: msg = LB_FINDSTRINGEXACT; break;
	default: return fr;
	}
	DWORD_PTR index;
	if (!SendMessageTimeout(control_window, msg, -1, (LPARAM)aString, SMTO_ABORTIFHUNG, 2000, &index))
		return FR_E_WIN32;
	if (index == CB_ERR) // CBERR == LB_ERR
		return FR_E_FAILED;
	aRetVal = index + 1;
	return OK;
}



bif_impl FResult ControlGetIndex(CONTROL_PARAMETERS_DECL, INT_PTR &aRetVal)
{
	DETERMINE_TARGET_CONTROL2;
	UINT msg;
	FResult fr = FAIL;
	switch (GetIndexControlType(control_window, fr, true))
	{
	case IC_COMBO: msg = CB_GETCURSEL; break;
	case IC_LIST: msg = LB_GETCURSEL; break;
	case IC_TAB: msg = TCM_GETCURSEL; break;
	default: return fr;
	}
	DWORD_PTR index;
	if (!SendMessageTimeout(control_window, msg, 0, 0, SMTO_ABORTIFHUNG, 2000, &index))
		return FR_E_WIN32;
	aRetVal = index + 1;
	return OK;
}



bif_impl FResult ControlGetChoice(CONTROL_PARAMETERS_DECL, StrRet &aRetVal)
{
	DETERMINE_TARGET_CONTROL2;
	UINT msg, x_msg, y_msg;
	FResult fr = FAIL;
	switch (GetIndexControlType(control_window, fr, true))
	{
	case IC_COMBO:
		msg = CB_GETCURSEL;
		x_msg = CB_GETLBTEXTLEN;
		y_msg = CB_GETLBTEXT;
		break;
	case IC_LIST:
		msg = LB_GETCURSEL;
		x_msg = LB_GETTEXTLEN;
		y_msg = LB_GETTEXT;
		break;
	default:
		return fr;
	}
	DWORD_PTR index, estimated_length, actual_length;
	if (!SendMessageTimeout(control_window, msg, 0, 0, SMTO_ABORTIFHUNG, 2000, &index))
		return FR_E_WIN32;
	if (index == CB_ERR) // CBERR == LB_ERR
		return FR_E_FAILED;
	if (!SendMessageTimeout(control_window, x_msg, (WPARAM)index, 0, SMTO_ABORTIFHUNG, 2000, &estimated_length))
		return FR_E_WIN32;
	if (estimated_length == CB_ERR)
		return FR_E_FAILED;
	// In unusual cases, MSDN says the indicated length might be longer than it actually winds up
	// being when the item's text is retrieved.
	auto buf = aRetVal.Alloc(estimated_length);
	if (!buf)
		return FR_E_OUTOFMEM;
	if (!SendMessageTimeout(control_window, y_msg, (WPARAM)index, (LPARAM)buf, SMTO_ABORTIFHUNG, 2000, &actual_length))
		return FR_E_WIN32;
	if (actual_length > estimated_length) // Guard against bogus return values. Also covers actual_length == CB_ERR due to unsigned type.
		return FR_E_FAILED;
	aRetVal.SetLength(actual_length); // Update to actual vs. estimated length.
	return OK;
}



static FResult ControlGetItems(Array *items, DWORD_PTR item_count, HWND control_window, UINT x_msg, UINT y_msg)
{
	DWORD_PTR length, u, item_length;
	// Calculate the required buffer size for the largest string.
	for (length = 0, u = 0; u < item_count; ++u)
	{
		if (!SendMessageTimeout(control_window, x_msg, u, 0, SMTO_ABORTIFHUNG, 5000, &item_length)
			|| item_length == LB_ERR) // Note that item_length is legitimately zero for a blank item in the list.
			return FR_E_FAILED;
		if (length < item_length)
			length = item_length;
	}
	// In unusual cases, MSDN says the indicated length might be longer than it actually winds up
	// being when the item's text is retrieved.
	++length; // To include the null-terminator.
	// Could use alloca when length is within a safe limit, but it seems unlikely to help
	// performance much since the bottleneck is probably in message passing between processes.
	// Although length is only the size of the largest item, not the total of all items,
	// I don't know how large an item can be, so it's safest to just use malloc().
	LPTSTR dyn_buf;
	if (  !(dyn_buf = tmalloc(length))  )
		return FR_E_OUTOFMEM;
	for (u = 0; u < item_count; ++u)
	{
		if (!SendMessageTimeout(control_window, y_msg, (WPARAM)u, (LPARAM)dyn_buf, SMTO_ABORTIFHUNG, 5000, &item_length)
			|| item_length > length) // Guard against bogus return values. Also covers item_length == CB_ERR due to unsigned type.
			break;
		dyn_buf[item_length] = '\0'; // Play it safe.
		if (!items->Append(dyn_buf))
			break;
	}
	free(dyn_buf);
	return u < item_count ? FR_E_FAILED : OK;
}

bif_impl FResult ControlGetItems(CONTROL_PARAMETERS_DECL, IObject *&aRetVal)
{
	DETERMINE_TARGET_CONTROL2;
	UINT msg, x_msg, y_msg;
	FResult fr = FAIL;
	switch (GetIndexControlType(control_window, fr, true))
	{
	case IC_COMBO:
		msg = CB_GETCOUNT;
		x_msg = CB_GETLBTEXTLEN;
		y_msg = CB_GETLBTEXT;
		break;
	case IC_LIST:
		msg = LB_GETCOUNT;
		x_msg = LB_GETTEXTLEN;
		y_msg = LB_GETTEXT;
		break;
	default:
		return fr;
	}
	DWORD_PTR item_count;
	if (!SendMessageTimeout(control_window, msg, 0, 0, SMTO_ABORTIFHUNG, 5000, &item_count))
		return FR_E_WIN32;
	if (item_count == LB_ERR) // There was a problem getting the count.
		return FR_E_FAILED;
	Array *items;
	if (  !(items = Array::Create())  )
		return FR_E_OUTOFMEM;
	fr = ControlGetItems(items, item_count, control_window, x_msg, y_msg);
	if (fr != OK)
	{
		items->Release();
		return fr;
	}
	aRetVal = items;
	return OK;
}



bif_impl FResult EditGetLineCount(CONTROL_PARAMETERS_DECL, UINT_PTR &aRetVal)
{
	DETERMINE_TARGET_CONTROL2;
	// MSDN: "If the control has no text, the return value is 1. The return value will never be less than 1."
	if (!SendMessageTimeout(control_window, EM_GETLINECOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&aRetVal))
		return FR_E_WIN32;
	return OK;
}



bif_impl FResult EditGetCurrentLine(CONTROL_PARAMETERS_DECL, UINT_PTR &aRetVal)
{
	DETERMINE_TARGET_CONTROL2;
	DWORD_PTR line_number;
	if (!SendMessageTimeout(control_window, EM_LINEFROMCHAR, -1, 0, SMTO_ABORTIFHUNG, 2000, &line_number))
		return FR_E_WIN32;
	aRetVal = line_number + 1;
	return OK;
}



bif_impl FResult EditGetCurrentCol(CONTROL_PARAMETERS_DECL, UINT &aRetVal)
{
	DETERMINE_TARGET_CONTROL2;
	DWORD_PTR dwResult, line_number;
	DWORD start = 0, end = 0;
	// The dwResult from the first msg below is not useful and is not checked.
	if (   !SendMessageTimeout(control_window, EM_GETSEL, (WPARAM)&start, (LPARAM)&end, SMTO_ABORTIFHUNG, 2000, &dwResult)
		|| !SendMessageTimeout(control_window, EM_LINEFROMCHAR, (WPARAM)start, 0, SMTO_ABORTIFHUNG, 2000, &line_number)   )
		return FR_E_WIN32;
	if (!line_number) // Since we're on line zero, the column number is simply start+1.
	{
		aRetVal = start + 1;  // +1 to convert from zero based.
		return OK;
	}
	// The original Au3 function repeatedly decremented the character index and called EM_LINEFROMCHAR
	// until the row changed. Don't know why; the EM_LINEINDEX method is MUCH faster for long lines and
	// probably has been available since the dawn of time, though I've only tested it on Windows 10.
	DWORD_PTR line_start;
	if (!SendMessageTimeout(control_window, EM_LINEINDEX, (WPARAM)line_number, 0, SMTO_ABORTIFHUNG, 2000, &line_start))
		return FR_E_WIN32;
	aRetVal = start - (UINT)line_start + 1;
	return OK;
}



bif_impl FResult EditGetLine(INT_PTR aIndex, CONTROL_PARAMETERS_DECL, StrRet &aRetVal)
{
	DETERMINE_TARGET_CONTROL2;
	auto control_index = aIndex - 1;
	if (control_index < 0)
		return FR_E_ARG(0);
	DWORD_PTR dwLineCount, dwResult;
	// Lexikos: Not sure if the following comment is relevant (does the OS multiply by sizeof(wchar_t)?).
	// jackieku: 32768 * sizeof(wchar_t) = 65536, which can not be stored in a unsigned 16bit integer.
	TCHAR line_buf[32767];
	*(LPWORD)line_buf = 32767; // EM_GETLINE requires first word of string to be set to its size.
	if (!SendMessageTimeout(control_window, EM_GETLINE, (WPARAM)control_index, (LPARAM)line_buf, SMTO_ABORTIFHUNG, 2000, &dwResult))
		return FR_E_WIN32;
	if (!dwResult) // Line is empty or line number is invalid.
	{
		if (!SendMessageTimeout(control_window, EM_GETLINECOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, &dwLineCount))
			return FR_E_WIN32;
		if ((DWORD_PTR)control_index >= dwLineCount)
			return FR_E_ARG(0);
	}
	return aRetVal.Copy(line_buf, dwResult) ? OK : FR_E_OUTOFMEM;
}



bif_impl FResult EditGetSelectedText(CONTROL_PARAMETERS_DECL, StrRet &aRetVal)
{
	DETERMINE_TARGET_CONTROL2;
	DWORD_PTR length, dwResult;
	DWORD start = 0, end = 0;
	// Note: The RichEdit controls of certain apps such as Metapad don't return the right selection
	// with this technique.  Au3 has the same problem with them, so for now it's just documented here
	// as a limitation.
	if (!SendMessageTimeout(control_window, EM_GETSEL, (WPARAM)&start, (LPARAM)&end, SMTO_ABORTIFHUNG, 2000, &dwResult))
		return FR_E_WIN32;
	// The above sets start to be the zero-based position of the start of the selection (similar for end).
	// If there is no selection, start and end will be equal, at least in the edit controls I tried it with.
	// The dwResult from the above is not useful and is not checked.
	if (start > end) // Later sections rely on this for safety with unsupported controls.
		return FR_E_FAILED; // The most likely cause is that this isn't an Edit control, but that isn't certain.
	if (start == end)
	{
		aRetVal.SetEmpty();
		return OK;
	}
	if (   !SendMessageTimeout(control_window, WM_GETTEXTLENGTH, 0, 0, SMTO_ABORTIFHUNG, 2000, &length)
		|| !length   ) // Since the above didn't return for start == end, this is an error because we have a selection of non-zero length, but no text to go with it!
		return FR_E_FAILED;
	// Dynamic memory is used because we must get all the control's text so that just the selected region
	// can be cropped out and returned:
	LPTSTR dyn_buf = tmalloc(length + 1);
	if (!dyn_buf)
		return FR_E_OUTOFMEM;
	if (   !SendMessageTimeout(control_window, WM_GETTEXT, (WPARAM)(length + 1), (LPARAM)dyn_buf, SMTO_ABORTIFHUNG, 2000, &length)
		|| !length || end > length   )
	{
		// The first check above implies failure since the length is unexpectedly zero
		// (above implied it shouldn't be).  The second check is also a problem because the
		// end of the selection should not be beyond the length of text that was retrieved.
		free(dyn_buf);
		return FR_E_FAILED;
	}
	dyn_buf[end] = '\0'; // Terminate the string at the end of the selection.
	auto fr = aRetVal.Copy(dyn_buf + start, end - start) ? OK : FR_E_OUTOFMEM;
	free(dyn_buf);
	return fr;
}



// The terminology "HWND" was chosen rather than "ID" to avoid confusion with a control's
// dialog ID (as retrieved by GetDlgCtrlID).  This also reserves the word ID for possible
// use with the control's Dialog ID in future versions.
bif_impl FResult ControlGetHwnd(CONTROL_PARAMETERS_DECL, UINT &aRetVal)
{
	DETERMINE_TARGET_CONTROL2;
	aRetVal = (UINT)(size_t)control_window;
	return OK;
}



bif_impl FResult Download(StrArg aURL, StrArg aFilespec)
{
	// v1.0.44.07: Set default to INTERNET_FLAG_RELOAD vs. 0 because the vast majority of usages would want
	// the file to be retrieved directly rather than from the cache.
	// v1.0.46.04: Added more no-cache flags because otherwise, it definitely falls back to the cache if
	// the remote server doesn't repond (and perhaps other errors), which defeats the ability to use the
	// Download command for uptime/server monitoring.  Also, in spite of what MSDN says, it seems nearly
	// certain based on other sources that more than one flag is supported.  Someone also mentioned that
	// INTERNET_FLAG_CACHE_IF_NET_FAIL is related to this, but there's no way to specify it in these
	// particular calls, and it's the opposite of the desired behavior anyway; so it seems impossible to
	// turn it off explicitly.
	DWORD flags_for_open_url = INTERNET_FLAG_RELOAD | INTERNET_FLAG_NO_CACHE_WRITE;
	aURL = omit_leading_whitespace(aURL);
	if (*aURL == '*') // v1.0.44.07: Provide an option to override flags_for_open_url.
	{
		flags_for_open_url = ATOU(++aURL);
		if (auto cp = StrChrAny(aURL, _T(" \t"))) // Find first space or tab.
			aURL = omit_leading_whitespace(cp);
	}

	// Open the internet session. v1.0.45.03: Provide a non-NULL user-agent because  some servers reject
	// requests that lack a user-agent.  Furthermore, it's more professional to have one, in which case it
	// should probably be kept as simple and unchanging as possible.  Using something like the script's name
	// as the user agent (even if documented) seems like a bad idea because it might contain personal/sensitive info.
	HINTERNET hInet = InternetOpen(_T("AutoHotkey"), INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY, NULL, NULL, 0);
	if (!hInet)
		return FR_E_WIN32;

	// Open the required URL
	HINTERNET hFile = InternetOpenUrl(hInet, aURL, NULL, 0, flags_for_open_url, 0);
	if (!hFile)
	{
		InternetCloseHandle(hInet);
		return FR_E_WIN32;
	}

	// Open our output file (overwrite if necessary)
	auto hOut = CreateFile(aFilespec, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, 0, NULL);
	if (hOut == INVALID_HANDLE_VALUE)
	{
		InternetCloseHandle(hFile);
		InternetCloseHandle(hInet);
		return FR_E_WIN32;
	}

	BYTE bufData[1024 * 1]; // v1.0.44.11: Reduced from 8 KB to alleviate GUI window lag during Download.  Testing shows this reduction doesn't affect performance on high-speed downloads (in fact, downloads are slightly faster; I tested two sites, one at 184 KB/s and the other at 380 KB/s).  It might affect slow downloads, but that seems less likely so wasn't tested.
	INTERNET_BUFFERSA buffers = {0};
	buffers.dwStructSize = sizeof(INTERNET_BUFFERSA);
	buffers.lpvBuffer = bufData;
	buffers.dwBufferLength = sizeof(bufData);

	LONG_OPERATION_INIT

	// Read the file.  I don't think synchronous transfers typically generate the pseudo-error
	// ERROR_IO_PENDING, so that is not checked here.  That's probably just for async transfers.
	// IRF_NO_WAIT is used to avoid requiring the call to block until the buffer is full.  By
	// having it return the moment there is any data in the buffer, the program is made more
	// responsive, especially when the download is very slow and/or one of the hooks is installed:
	BOOL result;
	DWORD bytes_written;
	if (*aURL == 'h' || *aURL == 'H')
	{
		while (result = InternetReadFileExA(hFile, &buffers, IRF_NO_WAIT, NULL)) // Assign
		{
			if (!buffers.dwBufferLength) // Transfer is complete.
				break;
			LONG_OPERATION_UPDATE  // Done in between the net-read and the file-write to improve avg. responsiveness.
			result = WriteFile(hOut, bufData, buffers.dwBufferLength, &bytes_written, nullptr);
			if (!result)
				break;
			buffers.dwBufferLength = sizeof(bufData);  // Reset buffer capacity for next iteration.
		}
	}
	else // v1.0.48.04: This section adds support for FTP and perhaps Gopher by using InternetReadFile() instead of InternetReadFileEx().
	{
		DWORD number_of_bytes_read;
		while (result = InternetReadFile(hFile, bufData, sizeof(bufData), &number_of_bytes_read))
		{
			if (!number_of_bytes_read)
				break;
			LONG_OPERATION_UPDATE
			result = WriteFile(hOut, bufData, number_of_bytes_read, &bytes_written, nullptr);
			if (!result)
				break;
		}
	}
	DWORD last_error = GetLastError();
	// Close internet session:
	InternetCloseHandle(hFile);
	InternetCloseHandle(hInet);
	// Close output file:
	CloseHandle(hOut);

	if (!result) // An error occurred during the transfer.
	{
		DeleteFile(aFilespec);  // Delete damaged/incomplete file.
		return FR_E_WIN32(last_error);
	}
	return OK;
}



int CALLBACK FileSelectFolderCallback(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData)
{
	if (uMsg == BFFM_INITIALIZED) // Caller has ensured that lpData isn't NULL by having set a valid lParam value.
		SendMessage(hwnd, BFFM_SETSELECTION, TRUE, lpData);
	// In spite of the quote below, the behavior does not seem to vary regardless of what value is returned
	// upon receipt of BFFM_VALIDATEFAILED, at least on XP.  But in case it matters on other OSes, preserve
	// compatibility with versions older than 1.0.36.03 by keeping the dialog displayed even if the user enters
	// an invalid folder:
	// MSDN: "Returns zero except in the case of BFFM_VALIDATEFAILED. For that flag, returns zero to dismiss
	// the dialog or nonzero to keep the dialog displayed."
	return uMsg == BFFM_VALIDATEFAILED; // i.e. zero should be returned in almost every case.
}



bif_impl FResult DirSelect(optl<StrArg> aRootDir, optl<int> aOptions, optl<StrArg> aGreeting, StrRet &aRetVal)
{
	if (g_nFolderDialogs >= MAX_FOLDERDIALOGS)
	{
		// Have a maximum to help prevent runaway hotkeys due to key-repeat feature, etc.
		return FError(_T("The maximum number of Folder Dialogs has been reached."));
	}

	LPMALLOC pMalloc;
	HRESULT hr = SHGetMalloc(&pMalloc);
    if (FAILED(hr))	// Initialize
		return hr;

	// v1.0.36.03: Support initial folder, which is different than the root folder because the root only
	// controls the origin point (above which the control cannot navigate).
	LPTSTR initial_folder;
	TCHAR root_dir[MAX_PATH*2 + 5];  // Up to two paths might be present inside, including an asterisk and spaces between them.
	if (aRootDir.has_value())
		tcslcpy(root_dir, aRootDir.value(), _countof(root_dir)); // Make a modifiable copy.
	else
		*root_dir = '\0';
	if (initial_folder = _tcschr(root_dir, '*'))
	{
		*initial_folder = '\0'; // Terminate so that root_dir becomes an isolated string.
		// Must eliminate the trailing whitespace or it won't work.  However, only up to one space or tab
		// so that path names that really do end in literal spaces can be used:
		if (initial_folder > root_dir && IS_SPACE_OR_TAB(initial_folder[-1]))
			initial_folder[-1] = '\0';
		// In case absolute paths can ever have literal leading whitespace, preserve that whitespace
		// by incrementing by only one and not calling omit_leading_whitespace().  This has been documented.
		++initial_folder;
	}
	else
		initial_folder = NULL;
	if (!*(omit_leading_whitespace(root_dir))) // Count all-whitespace as a blank string, but retain leading whitespace if there is also non-whitespace inside.
		*root_dir = '\0';

	BROWSEINFO bi;
	if (initial_folder)
	{
		bi.lpfn = FileSelectFolderCallback;
		bi.lParam = (LPARAM)initial_folder;  // Used by the callback above.
	}
	else
		bi.lpfn = NULL;  // It will ignore the value of bi.lParam when lpfn is NULL.

	if (*root_dir)
	{
		IShellFolder *pDF;
		if (SHGetDesktopFolder(&pDF) == NOERROR)
		{
			LPITEMIDLIST pIdl = NULL;
			ULONG        chEaten;
			ULONG        dwAttributes;
#ifdef UNICODE
			pDF->ParseDisplayName(NULL, NULL, root_dir, &chEaten, &pIdl, &dwAttributes);
#else
			OLECHAR olePath[MAX_PATH];			// wide-char version of path name
			ToWideChar(root_dir, olePath, MAX_PATH); // Dest. size is in wchars, not bytes.
			pDF->ParseDisplayName(NULL, NULL, olePath, &chEaten, &pIdl, &dwAttributes);
#endif
			pDF->Release();
			bi.pidlRoot = pIdl;
		}
	}
	else // No root directory.
		bi.pidlRoot = NULL;  // Make it use "My Computer" as the root dir.

	int iImage = 0;
	bi.iImage = iImage;
	bi.hwndOwner = THREAD_DIALOG_OWNER; // Can be NULL, which is used rather than main window since no need to have main window forced into the background by this.
	TCHAR greeting[1024];
	if (aGreeting.has_nonempty_value())
		tcslcpy(greeting, aGreeting.value(), _countof(greeting));
	else
		sntprintf(greeting, _countof(greeting), _T("Select Folder - %s"), g_script.DefaultDialogTitle());
	bi.lpszTitle = greeting;

	// Bitwise flags:
	#define FSF_ALLOW_CREATE 0x01
	#define FSF_EDITBOX      0x02
	#define FSF_NONEWDIALOG  0x04
	auto options = (DWORD)aOptions.value_or(FSF_ALLOW_CREATE);
	bi.ulFlags =
		  ((options & FSF_NONEWDIALOG)    ? 0           : BIF_NEWDIALOGSTYLE) // v1.0.48: Added to support BartPE/WinPE.
		| ((options & FSF_ALLOW_CREATE)   ? 0           : BIF_NONEWFOLDERBUTTON)
		| ((options & FSF_EDITBOX)        ? BIF_EDITBOX : 0);

	TCHAR Result[2048];
	bi.pszDisplayName = Result;  // This will hold the user's choice.

	// At this point, we know a dialog will be displayed.  See macro's comments for details:
	DIALOG_PREP
	POST_AHK_DIALOG(0) // Do this only after the above.  Must pass 0 for timeout in this case.

	++g_nFolderDialogs;
	LPITEMIDLIST lpItemIDList = SHBrowseForFolder(&bi);  // Spawn Dialog
	--g_nFolderDialogs;

	DIALOG_END
	if (!lpItemIDList)
	{
		// Due to rarity and because there doesn't seem to be any way to detect it,
		// no exception is thrown when the function fails.  Instead, we just assume
		// that the user pressed CANCEL (which should not be treated as an error):
		aRetVal.SetEmpty();
		return OK;
	}

	*Result = '\0';  // Reuse this var, this time to hold the result of the below:
	SHGetPathFromIDList(lpItemIDList, Result);
	pMalloc->Free(lpItemIDList);
	pMalloc->Release();

	return aRetVal.Copy(Result) ? OK : FR_E_OUTOFMEM;
}



bif_impl FResult FileGetShortcut(StrArg aShortcutFile, StrRet *aTarget, StrRet *aWorkingDir
	, StrRet *aArgs, StrRet *aDesc, StrRet *aIcon, ResultToken *aIconNum, int *aRunState)
// Credited to Holger <Holger.Kotsch at GMX de>.
{
	CoInitialize(NULL);
	IShellLink *psl;
	HRESULT hr;

	if (SUCCEEDED(hr = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID *)&psl)))
	{
		IPersistFile *ppf;
		if (SUCCEEDED(hr = psl->QueryInterface(IID_IPersistFile, (LPVOID *)&ppf)))
		{
#ifdef UNICODE
			if (SUCCEEDED(hr = ppf->Load(aShortcutFile, 0)))
#else
			WCHAR wsz[MAX_PATH+1]; // +1 hasn't been explained, but is retained in case it's needed.
			ToWideChar(aShortcutFile, wsz, MAX_PATH+1); // Dest. size is in wchars, not bytes.
			if (SUCCEEDED(hr = ppf->Load((const WCHAR*)wsz, 0)))
#endif
			{
				TCHAR buf[MAX_PATH+1];
				int icon_index, show_cmd;

				if (aTarget)
				{
					psl->GetPath(buf, MAX_PATH, NULL, SLGP_UNCPRIORITY);
					aTarget->Copy(buf);
				}
				if (aWorkingDir)
				{
					psl->GetWorkingDirectory(buf, MAX_PATH);
					aWorkingDir->Copy(buf);
				}
				if (aArgs)
				{
					psl->GetArguments(buf, MAX_PATH);
					aArgs->Copy(buf);
				}
				if (aDesc)
				{
					psl->GetDescription(buf, MAX_PATH); // Testing shows that the OS limits it to 260 characters.
					aDesc->Copy(buf);
				}
				if (aIcon || aIconNum)
				{
					psl->GetIconLocation(buf, MAX_PATH, &icon_index);
					if (aIcon)
						aIcon->Copy(buf);
					if (aIconNum)
						if (*buf)
							aIconNum->SetValue(icon_index + (icon_index >= 0 ? 1 : 0));  // Convert from 0-based to 1-based for consistency with the Menu command, etc. but leave negative resource IDs as-is.
						//else: Leave it blank to indicate that there is none.
				}
				if (aRunState)
				{
					psl->GetShowCmd(&show_cmd);
					*aRunState = show_cmd;
					// For the above, decided not to translate them to Max/Min/Normal since other
					// show-state numbers might be supported in the future (or are already).  In other
					// words, this allows the flexibility to specify some number other than 1/3/7 when
					// creating the shortcut in case it happens to work.  Of course, that applies only
					// to FileCreateShortcut, not here.  But it's done here so that this command is
					// compatible with that one.
				}
			}
			ppf->Release();
		}
		psl->Release();
	}
	CoUninitialize();
	return hr;
}



bif_impl FResult FileCreateShortcut(StrArg aTargetFile, StrArg aShortcutFile, optl<StrArg> aWorkingDir
	, optl<StrArg> aArgs, optl<StrArg> aDescription, optl<StrArg> aIconFile, optl<StrArg> aHotkey
	, optl<int> aIconNumber, optl<int> aRunState)
{
	CoInitialize(NULL);
	IShellLink *psl;
	HRESULT hr;

	if (SUCCEEDED(hr = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID *)&psl)))
	{
		psl->SetPath(aTargetFile);
		if (aWorkingDir.has_value())
			psl->SetWorkingDirectory(aWorkingDir.value());
		if (aArgs.has_value())
			psl->SetArguments(aArgs.value());
		if (aDescription.has_value())
			psl->SetDescription(aDescription.value());
		int icon_index = aIconNumber.value_or(0);
		if (aIconFile.has_value())
			psl->SetIconLocation(aIconFile.value(), icon_index - (icon_index > 0 ? 1 : 0)); // Convert 1-based index to 0-based, but leave negative resource IDs as-is.
		if (aHotkey.has_value())
		{
			// If badly formatted, it's not a critical error, just continue.
			// Currently, only shortcuts with a CTRL+ALT are supported.
			// AutoIt3 note: Make sure that CTRL+ALT is selected (otherwise invalid)
			vk_type vk = TextToVK(aHotkey.value());
			if (vk)
				// Vk in low 8 bits, mods in high 8:
				psl->SetHotkey(   (WORD)vk | ((WORD)(HOTKEYF_CONTROL | HOTKEYF_ALT) << 8)   );
		}
		if (aRunState.has_value())
			psl->SetShowCmd(*aRunState); // No validation is done since there's a chance other numbers might be valid now or in the future.

		IPersistFile *ppf;
		if (SUCCEEDED(hr = psl->QueryInterface(IID_IPersistFile,(LPVOID *)&ppf)))
		{
#ifndef UNICODE
			WCHAR wsz[MAX_PATH];
			ToWideChar(aShortcutFile, wsz, MAX_PATH); // Dest. size is in wchars, not bytes.
#else
			LPCWSTR wsz = aShortcutFile;
#endif
			// MSDN says to pass "The absolute path of the file".  Windows 10 requires it.
			WCHAR full_path[MAX_PATH];
			GetFullPathNameW(wsz, _countof(full_path), full_path, NULL);
			hr = ppf->Save(full_path, TRUE);
			ppf->Release();
		}
		psl->Release();
	}

	CoUninitialize();
	return hr;
}



bif_impl FResult FileRecycle(StrArg aFilePattern)
{
	if (!aFilePattern || !*aFilePattern)
		return FR_E_ARG(0);  // Since this is probably not what the user intended.

	SHFILEOPSTRUCT FileOp;
	TCHAR szFileTemp[_MAX_PATH+2];

	// au3: Get the fullpathname - required for UNDO to work
	Line::Util_GetFullPathName(aFilePattern, szFileTemp);

	// au3: We must also make it a double nulled string *sigh*
	szFileTemp[_tcslen(szFileTemp)+1] = '\0';

	// au3: set to known values - Corrects crash
	FileOp.hNameMappings = NULL;
	FileOp.lpszProgressTitle = NULL;
	FileOp.fAnyOperationsAborted = FALSE;
	FileOp.hwnd = NULL;
	FileOp.pTo = NULL;

	FileOp.pFrom = szFileTemp;
	FileOp.wFunc = FO_DELETE;
	FileOp.fFlags = FOF_SILENT | FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_WANTNUKEWARNING;

	// SHFileOperation() returns 0 on success:
	return SHFileOperation(&FileOp) ? FR_E_FAILED : OK;
}



bif_impl FResult FileRecycleEmpty(optl<StrArg> szPath)
{
	HRESULT hr = SHEmptyRecycleBin(NULL, szPath.value_or_null(), SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND);
	// Throw no error for E_UNEXPECTED, since that is returned in the common case where the
	// recycle bin is already empty.  Seems more useful and user-friendly to ignore it.
	return hr != E_UNEXPECTED ? hr : OK;
}



bif_impl FResult FileGetVersion(optl<StrArg> aPath, StrRet &aRetVal)
{
	g->LastError = 0; // Set default for successful return or non-Win32 errors.

	auto path = aPath.has_value() ? aPath.value() : g->mLoopFile ? g->mLoopFile->file_path : _T("");
	if (!*path)
		return FR_E_ARG(0);

	DWORD dwUnused, dwSize;
	if (   !(dwSize = GetFileVersionInfoSize(path, &dwUnused))   )  // No documented limit on how large it can be, so don't use _alloca().
		return FR_E_WIN32;

	BYTE *pInfo = (BYTE*)malloc(dwSize);  // Allocate the size retrieved by the above.
	VS_FIXEDFILEINFO *pFFI;
	UINT uSize;

	// Read the version resource
	if (!GetFileVersionInfo(path, 0, dwSize, (LPVOID)pInfo)
	// Locate the fixed information
		|| !VerQueryValue(pInfo, _T("\\"), (LPVOID *)&pFFI, &uSize))
	{
		free(pInfo);
		return FR_E_WIN32;
	}

	// extract the fields you want from pFFI
	UINT iFileMS = (UINT)pFFI->dwFileVersionMS;
	UINT iFileLS = (UINT)pFFI->dwFileVersionLS;
	sntprintf(aRetVal.CallerBuf(), aRetVal.CallerBufSize, _T("%u.%u.%u.%u")
		, (iFileMS >> 16), (iFileMS & 0xFFFF), (iFileLS >> 16), (iFileLS & 0xFFFF));
	aRetVal.SetTemp(aRetVal.CallerBuf());

	free(pInfo);
	return OK;
}



bool Line::Util_CopyDir(LPCTSTR szInputSource, LPCTSTR szInputDest, int OverwriteMode, bool bMove)
{
	if (!*szInputSource || !*szInputDest)
		return false;

	bool bOverwrite = OverwriteMode == 1 || OverwriteMode == 2; // Strict validation for safety.

	// Get the fullpathnames and strip trailing \s
	TCHAR szSource[_MAX_PATH+2];
	TCHAR szDest[_MAX_PATH+2];
	Util_GetFullPathName(szInputSource, szSource);
	Util_GetFullPathName(szInputDest, szDest);

	// Permit !IsDir when copying to support extracting zip files.  Other checks prevent
	// going in the other direction, which isn't supported by SHFileOperation anyway.
	// The "\*.*" suffix added below prevents any single file from being copied.
	if (bMove && Util_IsDir(szSource) == false)
		return false;							// Nope

	// Jon on the AutoIt forums says "Well SHFileOp is way too unpredictable under 9x. Grrr."
	// So the comments below about some OSes/old versions are probably just referring to 9x,
	// which we don't support anymore.  Testing on Windows 2000 and Windows 10 showed that
	// SHFileOperation can move a directory between two volumes.  However, performing copy
	// then remove ensures nothing is removed if the copy partially fails, so it's kept this
	// way for backward-compatibility, even though it may be inconsistent with local moves.
	// Obsolete comment from Util_MoveDir:
	// If the source and dest are on different volumes then we must copy rather than move
	// as move in this case only works on some OSes.  Copy and delete (poor man's move).
	if (bMove && (ctolower(szSource[0]) != ctolower(szDest[0]) || szSource[1] != ':'))
	{
		if (!Util_CopyDir(szSource, szDest, bOverwrite, false))
			return false;
		return Util_RemoveDir(szSource, true);
	}

	// Does the destination dir exist?
	DWORD attr = GetFileAttributes(szDest);
	if (attr != 0xFFFFFFFF) // Destination already exists as a file or directory.
	{
		if (attr & FILE_ATTRIBUTE_DIRECTORY) // Dest already exists as a directory.
		{
			if (!bOverwrite) // Overwrite Mode is "Never".
				return false;
		}
		else // Dest already exists as a file.
			return false; // Don't even attempt to overwrite a file with a dir, regardless of mode (I think SHFileOperation refuses to do it anyway).
	}
	else // Dest doesn't exist.
	{
		// We must create the top level directory
		// FOF_SILENT (which is included in FOF_NO_UI and means "Do not display a progress dialog box")
		// seems to be bugged on some older OSes (such as 2k and XP).  Specifically, it answers "No" to
		// the "confirmmkdir" dialog, which it isn't supposed to suppress, and ignores FOF_NOCONFIRMMKDIR.
		// Creating the directory first works around this.  Win 7 is okay without this; Vista wasn't tested.
		if (!bMove && !FileCreateDir(szDest))
			return false;
	}

	// The wildcard below is kept for backward-compatibility, although as indicated above, the
	// issues alluded to below are probably only on 9x, which is no longer supported.  Adding the
	// wildcard appears to permit copying a directory into itself (perhaps because the directory
	// itself isn't being copied), although we still document the result as "undefined".
	// This suffix also allows SHFileOperation to extract the contents of a zip file, and prevents
	// it from copying the source itself into the destination.
	if (!bMove)
		_tcscat(szSource, _T("\\*.*"));

	// We must also make source\dest double nulled strings for the SHFileOp API
	szSource[_tcslen(szSource)+1] = '\0';	
	szDest[_tcslen(szDest)+1] = '\0';	

	// Setup the struct
	SHFILEOPSTRUCT FileOp = {0};
	FileOp.pFrom = szSource;
	FileOp.pTo = szDest;
	FileOp.wFunc = bMove ? FO_MOVE : FO_COPY;
	FileOp.fFlags = FOF_NO_UI; // Set default.
	if (OverwriteMode == 2)
		FileOp.fFlags |= FOF_MULTIDESTFILES; // v1.0.46.07: Using the FOF_MULTIDESTFILES flag (as hinted by MSDN) overwrites/merges any existing target directory.  This logic supersedes and fixes old logic that didn't work properly when the source dir was being both renamed and moved to overwrite an existing directory.
	// All of the below left set to NULL/FALSE by the struct initializer higher above:
	//FileOp.hNameMappings			= NULL;
	//FileOp.lpszProgressTitle		= NULL;
	//FileOp.fAnyOperationsAborted	= FALSE;
	//FileOp.hwnd					= NULL;

	// If the source directory contains any saved webpages consisting of a SiteName.htm file and a
	// corresponding directory named SiteName_files, the following may indicate an error even when the
	// copy is successful. Under Windows XP at least, the return value is 7 under these conditions,
	// which according to WinError.h is "ERROR_ARENA_TRASHED: The storage control blocks were destroyed."
	// However, since this error might occur under a variety of circumstances, it probably wouldn't be
	// proper to consider it a non-error.
	// I also checked GetLastError() after calling SHFileOperation(), but it does not appear to be
	// valid/useful in this case (MSDN mentions this fact but isn't clear about it).
	// The issue appears to affect only FileCopyDir, not FileMoveDir or FileRemoveDir.  It also seems
	// unlikely to affect FileCopy/FileMove because they never copy directories.
	return !SHFileOperation(&FileOp);
}



bool Line::Util_RemoveDir(LPCTSTR szInputSource, bool bRecurse)
{
	if (!*szInputSource)
		return false;

	SHFILEOPSTRUCT	FileOp;
	TCHAR			szSource[_MAX_PATH+2];
	
	// If recursion not on just try a standard delete on the directory (the SHFile function WILL
	// delete a directory even if not empty no matter what flags you give it...)
	if (bRecurse == false)
	{
		// v1.1.31.00: Use the original source path in case its length exceeds _MAX_PATH.
		// Relative paths and trailing slashes are okay in this case, and Util_IsDir() is
		// not needed since the function only removes empty directories, not files.
		if (!RemoveDirectory(szInputSource))
			return false;
		else
			return true;
	}

	// Get the fullpathnames and strip trailing \s
	Util_GetFullPathName(szInputSource, szSource);

	// Ensure source is a directory
	if (Util_IsDir(szSource) == false)
		return false;							// Nope

	// We must also make double nulled strings for the SHFileOp API
	szSource[_tcslen(szSource)+1] = '\0';

	// Setup the struct
	FileOp.pFrom					= szSource;
	FileOp.pTo						= NULL;
	FileOp.hNameMappings			= NULL;
	FileOp.lpszProgressTitle		= NULL;
	FileOp.fAnyOperationsAborted	= FALSE;
	FileOp.hwnd						= NULL;

	FileOp.wFunc	= FO_DELETE;
	FileOp.fFlags	= FOF_SILENT | FOF_NOCONFIRMMKDIR | FOF_NOCONFIRMATION | FOF_NOERRORUI;
	
	return !SHFileOperation(&FileOp);
}



///////////////////////////////////////////////////////////////////////////////
// Util_CopyFile()
// (moves files too)
// Returns the number of files that could not be copied or moved due to error.
///////////////////////////////////////////////////////////////////////////////
int Line::Util_CopyFile(LPCTSTR szInputSource, LPCTSTR szInputDest, bool bOverwrite, bool bMove, DWORD &aLastError)
{
	TCHAR szSource[T_MAX_PATH];
	TCHAR szDest[T_MAX_PATH];
	TCHAR szDestPattern[MAX_PATH];

	// Get local version of our source/dest with full path names, strip trailing \s
	// and normalize the path separator (replace / with \).
	Util_GetFullPathName(szInputSource, szSource, _countof(szSource));
	Util_GetFullPathName(szInputDest, szDest, _countof(szDest));

	// If the source or dest is a directory then add *.* to the end
	if (Util_IsDir(szSource))
		_tcscat(szSource, _T("\\*.*"));
	if (Util_IsDir(szDest))
		_tcscat(szDest, _T("\\*.*"));

	WIN32_FIND_DATA	findData;
	HANDLE hSearch = FindFirstFile(szSource, &findData);
	if (hSearch == INVALID_HANDLE_VALUE)
	{
		aLastError = GetLastError(); // Set even in this case since FindFirstFile can fail due to actual errors, such as an invalid path.
		return StrChrAny(const_cast<LPTSTR>(szInputSource), _T("?*")) ? 0 : 1; // Indicate failure only if there were no wildcards.
	}
	aLastError = 0; // Set default. Overridden only when a failure occurs.

	// Otherwise, loop through all the matching files.

	// Locate the filename/pattern, which will be overwritten on each iteration.
	LPTSTR source_append_pos = _tcsrchr(szSource, '\\') + 1;
	LPTSTR dest_append_pos = _tcsrchr(szDest, '\\') + 1;
	size_t space_remaining = _countof(szSource) - (source_append_pos - szSource) - 1;

	// Copy destination filename or pattern, since it will be overwritten.
	tcslcpy(szDestPattern, dest_append_pos, _countof(szDestPattern));

	int failure_count = 0;
	LONG_OPERATION_INIT

	do
	{
		// Since other script threads can interrupt during LONG_OPERATION_UPDATE, it's important that
		// this function and those that call it not refer to sArgDeref[] and sArgVar[] anytime after an
		// interruption becomes possible. This is because an interrupting thread usually changes the
		// values to something inappropriate for this thread.
		LONG_OPERATION_UPDATE

		// Make sure the returned handle is a file and not a directory before we
		// try and do copy type things on it!
		if (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) // dwFileAttributes should never be invalid (0xFFFFFFFF) in this case.
			continue;

		if (_tcslen(findData.cFileName) > space_remaining) // v1.0.45.03: Basic check in case of files whose full spec is over 260 characters long.
		{
			aLastError = ERROR_BUFFER_OVERFLOW; // MSDN: "The file name is too long."
			++failure_count;
			continue;
		}
		_tcscpy(source_append_pos, findData.cFileName); // Indirectly populate szSource. Above has ensured this won't overflow.

		// Expand the destination based on this found file
		Util_ExpandFilenameWildcard(findData.cFileName, szDestPattern, dest_append_pos);

		// Fixed for v1.0.36.01: This section has been revised to avoid unnecessary calls; but more
		// importantly, it now avoids the deletion and complete loss of a file when it is copied or
		// moved onto itself.  That used to happen because any existing destination file used to be
		// deleted prior to attempting the move/copy.
		if (bMove)  // Move vs. copy mode.
		{
			// Note that MoveFile() is capable of moving a file to a different volume, regardless of
			// operating system version.  That's enough for what we need because this function never
			// moves directories, only files.

			// The following call will report success if source and dest are the same file, even if
			// source is something like "..\Folder\Filename.txt" and dest is something like
			// "C:\Folder\Filename.txt" (or if source is an 8.3 filename and dest is the long name
			// of the same file).  This is good because it avoids the need to devise code
			// to determine whether two different path names refer to the same physical file
			// (note that GetFullPathName() has shown itself to be inadequate for this purpose due
			// to problems with short vs. long names, UNC vs. mapped drive, and possibly NTFS hard
			// links (aliases) that might all cause two different filenames to point to the same
			// physical file on disk (hopefully MoveFile handles all of these correctly by indicating
			// success [below] when a file is moved onto itself, though it has only been tested for
			// basic cases of relative vs. absolute path).
			if (!MoveFile(szSource, szDest))
			{
				// If overwrite mode was not specified by the caller, or it was but the existing
				// destination file cannot be deleted (perhaps because it is a folder rather than
				// a file), or it can be deleted but the source cannot be moved, indicate a failure.
				// But by design, continue the operation.  The following relies heavily on
				// short-circuit boolean evaluation order:
				if (   !(bOverwrite && DeleteFile(szDest) && MoveFile(szSource, szDest))   )
				{
					aLastError = GetLastError();
					++failure_count; // At this stage, any of the above 3 being false is cause for failure.
				}
				//else everything succeeded, so nothing extra needs to be done.  In either case,
				// continue on to the next file.
			}
		}
		else // The mode is "Copy" vs. "Move"
			if (!CopyFile(szSource, szDest, !bOverwrite)) // Force it to fail if bOverwrite==false.
			{
				aLastError = GetLastError();
				++failure_count;
			}
	} while (FindNextFile(hSearch, &findData));

	FindClose(hSearch);
	return failure_count;
}



void Line::Util_ExpandFilenameWildcard(LPCTSTR szSource, LPCTSTR szDest, LPTSTR szExpandedDest)
{
	// copy one.two.three  *.txt     = one.two   .txt
	// copy one.two.three  *.*.txt   = one.two.  .txt  (extra asterisks are removed)
	// copy one.two		   test      = test

	TCHAR	szSrcFile[_MAX_PATH+1];
	TCHAR	szSrcExt[_MAX_PATH+1];

	TCHAR	szDestFile[_MAX_PATH+1];
	TCHAR	szDestExt[_MAX_PATH+1];

	// If the destination doesn't include a wildcard, send it back verbatim
	if (_tcschr(szDest, '*') == NULL)
	{
		_tcscpy(szExpandedDest, szDest);
		return;
	}

	// Split source and dest into file and extension
	_tsplitpath( szSource, NULL, NULL, szSrcFile, szSrcExt );
	_tsplitpath( szDest, NULL, NULL, szDestFile, szDestExt );

	// Source and Dest ext will either be ".nnnn" or "" or ".*", remove the period
	if (szSrcExt[0] == '.')
		_tcscpy(szSrcExt, &szSrcExt[1]);
	if (szDestExt[0] == '.')
		_tcscpy(szDestExt, &szDestExt[1]);

	// Replace first * in the destfile with the srcfile, remove any other *
	Util_ExpandFilenameWildcardPart(szSrcFile, szDestFile, szExpandedDest);
	
	if (*szSrcExt || *szDestExt)
	{
		LPTSTR ext = _tcschr(szExpandedDest, '\0');
		
		if (!szDestExt[0])
		{
			// Always include the source extension if destination extension was blank
			// (for backward-compatibility, this is done even if a '.' was present)
			szDestExt[0] = '*';
			szDestExt[1] = '\0';
		}

		// Replace first * in the destext with the srcext, remove any other *
		Util_ExpandFilenameWildcardPart(szSrcExt, szDestExt, ext + 1);

		// If there's a non-blank extension, replace the filename's null terminator with .
		if (ext[1])
			*ext = '.';
	}
}



void Line::Util_ExpandFilenameWildcardPart(LPCTSTR szSource, LPCTSTR szDest, LPTSTR szExpandedDest)
{
	LPTSTR lpTemp;
	int i, j, k;

	// Replace first * in the dest with the src, remove any other *
	i = 0; j = 0; k = 0;
	lpTemp = (LPTSTR)_tcschr(szDest, '*');
	if (lpTemp != NULL)
	{
		// Contains at least one *, copy up to this point
		while(szDest[i] != '*')
			szExpandedDest[j++] = szDest[i++];
		// Skip the * and replace in the dest with the srcext
		while(szSource[k] != '\0')
			szExpandedDest[j++] = szSource[k++];
		// Skip any other *
		i++;
		while(szDest[i] != '\0')
		{
			if (szDest[i] == '*')
				i++;
			else
				szExpandedDest[j++] = szDest[i++];
		}
		szExpandedDest[j] = '\0';
	}
	else
	{
		// No wildcard, straight copy of destext
		_tcscpy(szExpandedDest, szDest);
	}
}



bool Line::Util_IsDir(LPCTSTR szPath) // Returns true if the path is a directory
{
	DWORD dwTemp = GetFileAttributes(szPath);
	return dwTemp != 0xffffffff && (dwTemp & FILE_ATTRIBUTE_DIRECTORY);
}



void Line::Util_GetFullPathName(LPCTSTR szIn, LPTSTR szOut)
// Returns the full pathname and strips any trailing \s.  Assumes output is _MAX_PATH in size.
{
	ASSERT(*szIn && szOut);
	LPTSTR szFilePart;
	GetFullPathName(szIn, _MAX_PATH, szOut, &szFilePart);
	strip_trailing_backslash(szOut);
}



void Line::Util_GetFullPathName(LPCTSTR szIn, LPTSTR szOut, DWORD aBufSize)
{
	ASSERT(*szIn && szOut);
	GetFullPathName(szIn, aBufSize, szOut, NULL);
	strip_trailing_backslash(szOut);
}



FResult Shutdown(int nFlag)
// Shutdown or logoff the system.
// Returns OK if successful.
{
/* 
flags can be a combination of:
#define EWX_LOGOFF           0
#define EWX_SHUTDOWN         0x00000001
#define EWX_REBOOT           0x00000002
#define EWX_FORCE            0x00000004
#define EWX_POWEROFF         0x00000008 */

	HANDLE				hToken; 
	TOKEN_PRIVILEGES	tkp; 

	// Get a token for this process.
 	if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken)) 
		return FR_E_WIN32;
 
	// Get the LUID for the shutdown privilege.
 	LookupPrivilegeValue(NULL, SE_SHUTDOWN_NAME, &tkp.Privileges[0].Luid); 
 
	tkp.PrivilegeCount = 1;  /* one privilege to set */
	tkp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED; 
 
	// Get the shutdown privilege for this process.
 	AdjustTokenPrivileges(hToken, FALSE, &tkp, 0, (PTOKEN_PRIVILEGES)NULL, 0); 
 
	// Cannot test the return value of AdjustTokenPrivileges.
 	if (GetLastError() != ERROR_SUCCESS) 
		return FR_E_WIN32;

	// ExitWindows
	if (ExitWindowsEx(nFlag, 0))
		return OK;
	else
		return FR_E_WIN32;
}



void Util_WinKill(HWND hWnd)
{
	DWORD_PTR dwResult;
	// Use WM_CLOSE vs. SC_CLOSE in this case, since the target window is slightly more likely to
	// respond to that:
	if (!SendMessageTimeout(hWnd, WM_CLOSE, 0, 0, SMTO_ABORTIFHUNG, 500, &dwResult)) // Wait up to 500ms.
	{
		// Use more force - Mwuahaha
		DWORD pid = 0;
		GetWindowThreadProcessId(hWnd, &pid);
		HANDLE hProcess = pid ? OpenProcess(PROCESS_ALL_ACCESS, FALSE, pid) : NULL;
		if (hProcess)
		{
			TerminateProcess(hProcess, 0);
			CloseHandle(hProcess);
		}
	}
}



void DoIncrementalMouseMove(int aX1, int aY1, int aX2, int aY2, int aSpeed)
// aX1 and aY1 are the starting coordinates, and "2" are the destination coordinates.
// Caller has ensured that aSpeed is in the range 0 to 100, inclusive.
{
	// AutoIt3: So, it's a more gradual speed that is needed :)
	int delta;
	#define INCR_MOUSE_MIN_SPEED 32

	while (aX1 != aX2 || aY1 != aY2)
	{
		if (aX1 < aX2)
		{
			delta = (aX2 - aX1) / aSpeed;
			if (delta == 0 || delta < INCR_MOUSE_MIN_SPEED)
				delta = INCR_MOUSE_MIN_SPEED;
			if ((aX1 + delta) > aX2)
				aX1 = aX2;
			else
				aX1 += delta;
		} 
		else 
			if (aX1 > aX2)
			{
				delta = (aX1 - aX2) / aSpeed;
				if (delta == 0 || delta < INCR_MOUSE_MIN_SPEED)
					delta = INCR_MOUSE_MIN_SPEED;
				if ((aX1 - delta) < aX2)
					aX1 = aX2;
				else
					aX1 -= delta;
			}

		if (aY1 < aY2)
		{
			delta = (aY2 - aY1) / aSpeed;
			if (delta == 0 || delta < INCR_MOUSE_MIN_SPEED)
				delta = INCR_MOUSE_MIN_SPEED;
			if ((aY1 + delta) > aY2)
				aY1 = aY2;
			else
				aY1 += delta;
		} 
		else 
			if (aY1 > aY2)
			{
				delta = (aY1 - aY2) / aSpeed;
				if (delta == 0 || delta < INCR_MOUSE_MIN_SPEED)
					delta = INCR_MOUSE_MIN_SPEED;
				if ((aY1 - delta) < aY2)
					aY1 = aY2;
				else
					aY1 -= delta;
			}

		MouseEvent(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, 0, aX1, aY1);
		DoMouseDelay();
		// Above: A delay is required for backward compatibility and because it's just how the incremental-move
		// feature was originally designed in AutoIt v3.  It may in fact improve reliability in some cases,
		// especially with the mouse_event() method vs. SendInput/Play.
	} // while()
}



////////////////////
// PROCESS ROUTINES
////////////////////

DWORD ProcessExist(LPCTSTR aProcess, bool aGetParent)
{
	// Determine the PID if aProcess is a pure, non-negative integer (any negative number
	// is more likely to be the name of a process [with a leading dash], rather than the PID).
	DWORD specified_pid = IsNumeric(aProcess) ? ATOU(aProcess) : 0;
	if (specified_pid && !aGetParent)
	{
		// Most of the time while a PID is being used, the process still exists.  So try opening
		// the process directly, since doing so is much faster than enumerating all processes:
		if (HANDLE hproc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION | SYNCHRONIZE, FALSE, specified_pid))
		{
			DWORD wait_result = WAIT_FAILED;
			// OpenProcess can succeed for erroneous PID values; e.g. values of 1501 to 1503 can
			// open the process with ID 1500.  This is likely due to the undocumented fact that
			// PIDs are a multiple of four: https://devblogs.microsoft.com/oldnewthing/20080228-00/?p=23283
			DWORD actual_pid = GetProcessId(hproc); // Requires PROCESS_QUERY_LIMITED_INFORMATION access.
			if (actual_pid == specified_pid)
			{
				// OpenProcess can succeed for a process which has already exited if another process
				// still has an open handle to it.  So check whether it's still running:
				wait_result = WaitForSingleObject(hproc, 0); // Requires SYNCHRONIZE access.
			}
			CloseHandle(hproc);
			if (wait_result == WAIT_OBJECT_0)
				return 0; // Process has exited.
			if (wait_result == WAIT_TIMEOUT)
				return specified_pid; // Process still running.
			// Otherwise, fall back to the slow but more reliable method to get a more conclusive result.
		}
		// If OpenProcess failed, some likely causes are ERROR_ACCESS_DENIED and ERROR_INVALID_PARAMETER.
		// The latter probably indicates the PID is invalid, but we continue anyway, for the unlikely
		// case of a process with a name composed of digits and no extension (verified possible).
		// If ERROR_ACCESS_DENIED was returned, we can't rule out the false-positive cases described
		// above without doing a thorough enumeration of processes, so continue in that case as well.
	}

	PROCESSENTRY32 proc;
    proc.dwSize = sizeof(proc);
	HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
	Process32First(snapshot, &proc);

	while (Process32Next(snapshot, &proc))
	{
		if (specified_pid && specified_pid == proc.th32ProcessID)
		{
			CloseHandle(snapshot);
			return aGetParent ? proc.th32ParentProcessID : specified_pid;
		}
		// Otherwise, check for matching name even if aProcess is purely numeric (i.e. a number might
		// also be a valid name?):
		// It seems that proc.szExeFile never contains a path, just the executable name.
		// But in case it ever does, ensure consistency by removing the path:
		LPCTSTR proc_name = _tcsrchr(proc.szExeFile, '\\');
		proc_name = proc_name ? proc_name + 1 : proc.szExeFile;
		if (!_tcsicmp(proc_name, aProcess)) // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales; 3) performance.
		{
			CloseHandle(snapshot);
			return aGetParent ? proc.th32ParentProcessID : proc.th32ProcessID;
		}
	}
	CloseHandle(snapshot);
	return 0;  // Not found.
}
