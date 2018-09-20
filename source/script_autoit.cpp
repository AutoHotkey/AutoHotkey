///////////////////////////////////////////////////////////////////////////////
//
// AutoIt
//
// Copyright (C)1999-2006:
//		- Jonathan Bennett <jon@hiddensoft.com>
//		- Others listed at http://www.autoitscript.com/autoit3/docs/credits.htm
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
#include <wininet.h> // For URLDownloadToFile().
#include "script.h"
#include "globaldata.h" // for g_ErrorLevel and probably other globals.
#include "window.h" // For ControlExist().
#include "application.h" // For SLEEP_WITHOUT_INTERRUPTION and MsgSleep().


ResultType Script::DoRunAs(LPTSTR aCommandLine, LPTSTR aWorkingDir, bool aDisplayErrors, WORD aShowWindow
	, Var *aOutputVar, PROCESS_INFORMATION &aPI, bool &aSuccess // Output parameters we set for caller, but caller must have initialized aSuccess to false.
	, HANDLE &aNewProcess, DWORD &aLastError)                   // Same, but initialize to NULL.
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
	// Get a handle to the DLL module that contains CreateProcessWithLogonW
	HINSTANCE hinstLib = LoadLibrary(_T("advapi32"));
	if (!hinstLib)
	{
		if (aDisplayErrors)
			ScriptError(_T("RunAs: Missing advapi32.dll."));
		return FAIL;
	}
	MyCreateProcessWithLogonW lpfnDLLProc = (MyCreateProcessWithLogonW)GetProcAddress(hinstLib, "CreateProcessWithLogonW");
	if (!lpfnDLLProc)
	{
		FreeLibrary(hinstLib);
		if (aDisplayErrors)
			ScriptError(_T("CreateProcessWithLogonW.")); // Short msg since it probably never happens.
		return FAIL;
	}
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
	if (lpfnDLLProc(mRunAsUser, mRunAsDomain, mRunAsPass, LOGON_WITH_PROFILE, 0
		, aCommandLine, 0, 0, aWorkingDir && *aWorkingDir ? aWorkingDir : NULL, &wsi, &aPI))
#endif
	{
		aSuccess = true;
		if (aPI.hThread)
			CloseHandle(aPI.hThread); // Required to avoid memory leak.
		aNewProcess = aPI.hProcess;
		if (aOutputVar)
			aOutputVar->Assign(aPI.dwProcessId);
	}
	else
		aLastError = GetLastError(); // Caller will use this to get an error message and set g->LastError if needed.
	FreeLibrary(hinstLib);
	return OK;
}



VarSizeType BIV_IPAddress(LPTSTR aBuf, LPTSTR aVarName)
{
	// aaa.bbb.ccc.ddd = 15, but allow room for larger IP's in the future.
	#define IP_ADDRESS_SIZE 32 // The maximum size of any of the strings we return, including terminator.
	if (!aBuf)
		return IP_ADDRESS_SIZE - 1;  // -1 since we're returning the length of the var's contents, not the size.

	WSADATA wsadata;
	if (WSAStartup(MAKEWORD(1, 1), &wsadata)) // Failed (it returns 0 on success).
	{
		*aBuf = '\0';
		return 0;
	}

	char host_name[256];
	gethostname(host_name, _countof(host_name));
	HOSTENT *lpHost = gethostbyname(host_name);

	// au3: How many adapters have we?
	int adapter_count = 0;
	while (lpHost->h_addr_list[adapter_count])
		++adapter_count;

	int adapter_index = aVarName[11] - '1'; // A_IPAddress[1-4]
	if (adapter_index >= adapter_count)
		_tcscpy(aBuf, _T("0.0.0.0"));
	else
	{
		IN_ADDR inaddr;
		memcpy(&inaddr, lpHost->h_addr_list[adapter_index], 4);
		tcslcpy(aBuf, CStringTCharFromCharIfNeeded(inet_ntoa(inaddr)), IP_ADDRESS_SIZE);
	}

	WSACleanup();
	return (VarSizeType)_tcslen(aBuf);
}



VarSizeType BIV_IsAdmin(LPTSTR aBuf, LPTSTR aVarName)
{
	if (!aBuf)
		return 1;  // The length of the string "1" or "0".
	TCHAR result = '0';  // Default.
	if (g_os.IsWin9x())
		result = '1';
	else
	{
		SC_HANDLE h = OpenSCManager(NULL, NULL, SC_MANAGER_LOCK);
		if (h)
		{
			SC_LOCK lock = LockServiceDatabase(h);
			if (lock)
			{
				UnlockServiceDatabase(lock);
				result = '1'; // Current user is admin.
			}
			else
			{
				DWORD lastErr = GetLastError();
				if (lastErr == ERROR_SERVICE_DATABASE_LOCKED)
					result = '1'; // Current user is admin.
			}
			CloseServiceHandle(h);
		}
	}
	aBuf[0] = result;
	aBuf[1] = '\0';
	return 1; // Length of aBuf.
}



ResultType Line::PixelGetColor(int aX, int aY, LPTSTR aOptions)
{
	if (tcscasestr(aOptions, _T("Slow"))) // New mode for v1.0.43.10.  Takes precedence over Alt mode.
		return PixelSearch(aX, aY, aX, aY, 0, 0, aOptions, true); // It takes care of setting ErrorLevel and the output-var.

	Var &output_var = *OUTPUT_VAR;
	output_var.Assign(); // Init to empty string regardless of whether we succeed here.

	CoordToScreen(aX, aY, COORD_MODE_PIXEL);
	
	bool use_alt_mode = tcscasestr(aOptions, _T("Alt")) != NULL; // New mode for v1.0.43.10: Two users reported that CreateDC works better in certain windows such as SciTE, at least one some systems.
	HDC hdc = use_alt_mode ? CreateDC(_T("DISPLAY"), NULL, NULL, NULL) : GetDC(NULL);
	if (!hdc)
		return SetErrorLevelOrThrow();

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

	TCHAR buf[32];
	_stprintf(buf, _T("0x%06X"), tcscasestr(aOptions, _T("RGB")) ? bgr_to_rgb(color) : color);
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return output_var.Assign(buf);
}



ResultType Line::SplashTextOn(int aWidth, int aHeight, LPTSTR aTitle, LPTSTR aText)
{
	// Add some caption and frame size to window:
	aWidth += GetSystemMetrics(SM_CXFIXEDFRAME) * 2;
	int min_height = GetSystemMetrics(SM_CYCAPTION) + (GetSystemMetrics(SM_CXFIXEDFRAME) * 2);
	// This method seems more friendly than setting aHeight = min_height when aHeight < min_height.
	aHeight += min_height;

	POINT pt = CenterWindow(aWidth, aHeight); // Determine how to center the window in the region that excludes the task bar.

	// My: Probably not too much overhead to do this, though it probably would perform better to resize and
	// "re-text" the existing window rather than recreating it like this:
	DESTROY_SPLASH

	// Doesn't seem necessary to have it owned by the main window, but neither
	// does doing so seem to cause any harm.  Feels safer to have it be
	// an independent window.  Update: Must make it owned by the parent window
	// otherwise it will get its own task-bar icon, which is usually undesirable.
	// In addition, making it an owned window should automatically cause it to be
	// destroyed when it's parent window is destroyed:
	g_hWndSplash = CreateWindowEx(WS_EX_TOPMOST, WINDOW_CLASS_SPLASH, aTitle, WS_DISABLED|WS_POPUP|WS_CAPTION
		, pt.x, pt.y, aWidth, aHeight, g_hWnd, (HMENU)NULL, g_hInstance, NULL);

	RECT rect;
	GetClientRect(g_hWndSplash, &rect);	// get the client size

	// CREATE static label full size of client area.
	HWND static_win = CreateWindowEx(0, _T("static"), aText, WS_CHILD|WS_VISIBLE|SS_CENTER
		, 0, 0, rect.right - rect.left, rect.bottom - rect.top, g_hWndSplash, (HMENU)NULL, g_hInstance, NULL);

	if (!g_hFontSplash)
	{
		TCHAR default_font_name[65];
		int CyPixels, nSize = 12, nWeight = FW_NORMAL;
		HDC hdc = CreateDC(_T("DISPLAY"), NULL, NULL, NULL);
		if (FontExist(hdc, _T("Segoe UI"))) // Use a more appealing font under Windows Vista or later (Segoe UI).
		{
			nSize = 11;
			_tcscpy(default_font_name, _T("Segoe UI"));
		}
		else
		{
			SelectObject(hdc, (HFONT)GetStockObject(DEFAULT_GUI_FONT));		// Get Default Font Name
			GetTextFace(hdc, _countof(default_font_name) - 1, default_font_name); // -1 just in case, like AutoIt3.
		}
		CyPixels = GetDeviceCaps(hdc, LOGPIXELSY);			// For Some Font Size Math
		DeleteDC(hdc);
		//strcpy(default_font_name,vParams[7].szValue());	// Font Name
		//nSize = vParams[8].nValue();		// Font Size
		//if ( vParams[9].nValue() >= 0 && vParams[9].nValue() <= 1000 )
		//	nWeight = vParams[9].nValue();			// Font Weight
		g_hFontSplash = CreateFont(0-(nSize*CyPixels)/72,0,0,0,nWeight,0,0,0,DEFAULT_CHARSET,
			OUT_TT_PRECIS,CLIP_DEFAULT_PRECIS,PROOF_QUALITY,FF_DONTCARE,default_font_name);	// Create Font
		// The font is deleted when by g_script's destructor.
	}

	SendMessage(static_win, WM_SETFONT, (WPARAM)g_hFontSplash, MAKELPARAM(TRUE, 0));	// Do Font
	ShowWindow(g_hWndSplash, SW_SHOWNOACTIVATE);				// Show the Splash
	// Doesn't help with the brief delay in updating the window that happens when
	// something like URLDownloadToFile is used immediately after SplashTextOn:
	//InvalidateRect(g_hWndSplash, NULL, TRUE);
	// But this does, but for now it seems unnecessary since the user can always do
	// a manual sleep in the extremely rare cases this ever happens (even when it does
	// happen, the window updates eventually, after the download starts, at least on
	// my system.  Update: Might as well do it since it's a little nicer this way
	// (the text appears more quickly when the command after the splash is something
	// that might keep our thread tied up and unable to check messages).
	SLEEP_WITHOUT_INTERRUPTION(-1)
	// UpdateWindow() would probably achieve the same effect as the above, but it feels safer to do
	// the above because it ensures that our message queue is empty prior to returning to our caller.
	return OK;
}



ResultType Line::WinMenuSelectItem(LPTSTR aTitle, LPTSTR aText, LPTSTR aMenu1, LPTSTR aMenu2
	, LPTSTR aMenu3, LPTSTR aMenu4, LPTSTR aMenu5, LPTSTR aMenu6, LPTSTR aMenu7
	, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	// Set up a temporary array make it easier to traverse nested menus & submenus
	// in a loop.  Also add a NULL at the end to simplify the loop a little:
	LPTSTR menu_param[] = {aMenu1, aMenu2, aMenu3, aMenu4, aMenu5, aMenu6, aMenu7, NULL};

	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		goto error;

	int first_menu_param = 0;
	UINT message = WM_COMMAND;
	HMENU hMenu;
	if (!_tcsicmp(aMenu1, _T("0&")))
	{
		hMenu = GetSystemMenu(target_window, FALSE);
		first_menu_param = 1;
		message = WM_SYSCOMMAND;
	}
	else
		hMenu = GetMenu(target_window);
	if (!hMenu) // Window has no menu bar.
		goto error;

	int menu_item_count = GetMenuItemCount(hMenu);
	if (menu_item_count < 1) // Menu bar has no menus.
		goto error;
	
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
	LPTSTR this_menu_param;

	for (int i = first_menu_param; ; ++i)
	{
		this_menu_param = menu_param[i]; // For performance and convenience.
		if (!(this_menu_param && *this_menu_param))
			break;
		if (!hMenu)  // The nesting of submenus ended prior to the end of the list of menu search terms.
			goto error;

		this_menu_param_length = _tcslen(this_menu_param);
		target_menu_pos = (this_menu_param[this_menu_param_length - 1] == '&') ? ATOI(this_menu_param) - 1 : -1;
		if (target_menu_pos > -1)
		{
			if (target_menu_pos >= menu_item_count)  // Invalid menu position (doesn't exist).
				goto error;
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
				if (!match_found)
				{
					// Try again to find a match, this time without the ampersands used to indicate
					// a menu item's shortcut key:
					StrReplace(menu_text, _T("&"), _T(""), SCS_SENSITIVE);
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
				goto error;
		} // else
	} // outer for()

	// This would happen if the outer loop above had zero iterations due to aMenu1 being NULL or blank,
	// or if the caller specified a submenu as the target (which doesn't seem valid since an app would
	// next expect to ever receive a message for a submenu?):
	if (menu_id == MENU_ITEM_IS_SUBMENU)
		goto error;

	// Since the above didn't return, the specified search hierarchy was completely found.
	PostMessage(target_window, message, (WPARAM)menu_id, 0);
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.

error:
	return SetErrorLevelOrThrow();
}



ResultType Line::Control(LPTSTR aCmd, LPTSTR aValue, LPTSTR aControl, LPTSTR aTitle, LPTSTR aText
	, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
// ATTACH_THREAD_INPUT has been tested to see if they help any of these work with controls
// in MSIE (whose Internet Explorer_TridentCmboBx2 does not respond to "Control Choose" but
// does respond to "Control Focus").  But it didn't help.
{
	ControlCmds control_cmd = ConvertControlCmd(aCmd);
	// Since command names are validated at load-time, this only happens if the command name
	// was contained in a variable reference.  Since that is very rare, just set ErrorLevel
	// and return:
	if (control_cmd == CONTROL_CMD_INVALID)
		goto error;

	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		goto error;
	HWND control_window = ControlExist(target_window, aControl); // This can return target_window itself for cases such as ahk_id %ControlHWND%.
	if (!control_window)
		goto error;

	HWND immediate_parent;  // Possibly not the same as target_window since controls can themselves have children.
	int control_id, control_index;
	DWORD_PTR dwResult, new_button_state;
	UINT msg, x_msg, y_msg;
	RECT rect;
	LPARAM lparam;
	vk_type vk;
	int key_count;
	TCHAR temp_buf[32];

	switch(control_cmd)
	{
	case CONTROL_CMD_CHECK: // au3: Must be a Button
	case CONTROL_CMD_UNCHECK:
	{ // Need braces for ATTACH_THREAD_INPUT macro.
		new_button_state = (control_cmd == CONTROL_CMD_CHECK) ? BST_CHECKED : BST_UNCHECKED;
		if (!SendMessageTimeout(control_window, BM_GETCHECK, 0, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			goto error;
		if (dwResult == new_button_state) // It's already in the right state, so don't press it.
			break;
		// MSDN docs for BM_CLICK (and au3 author says it applies to this situation also):
		// "If the button is in a dialog box and the dialog box is not active, the BM_CLICK message
		// might fail. To ensure success in this situation, call the SetActiveWindow function to activate
		// the dialog box before sending the BM_CLICK message to the button."
		ATTACH_THREAD_INPUT
		SetActiveWindow(target_window == control_window ? GetNonChildParent(control_window) : target_window); // v1.0.44.13: Fixed to allow for the fact that target_window might be the control itself (e.g. via ahk_id %ControlHWND%).
		if (!GetWindowRect(control_window, &rect))	// au3: Code to primary click the centre of the control
			rect.bottom = rect.left = rect.right = rect.top = 0;
		lparam = MAKELPARAM((rect.right - rect.left) / 2, (rect.bottom - rect.top) / 2);
		PostMessage(control_window, WM_LBUTTONDOWN, MK_LBUTTON, lparam);
		PostMessage(control_window, WM_LBUTTONUP, 0, lparam);
		DETACH_THREAD_INPUT
		break;
	}

	case CONTROL_CMD_ENABLE:
		EnableWindow(control_window, TRUE);
		break;

	case CONTROL_CMD_DISABLE:
		EnableWindow(control_window, FALSE);
		break;

	case CONTROL_CMD_SHOW:
		ShowWindow(control_window, SW_SHOWNOACTIVATE); // SW_SHOWNOACTIVATE has been seen in some example code for this purpose.
		break;

	case CONTROL_CMD_HIDE:
		ShowWindow(control_window, SW_HIDE);
		break;

	case CONTROL_CMD_STYLE:
	case CONTROL_CMD_EXSTYLE:
	{
		if (!*aValue)
			return OK; // Seems best not to treat an explicit blank as zero.  Let ErrorLevel tell the story. 
		int style_index = (control_cmd == CONTROL_CMD_STYLE) ? GWL_STYLE : GWL_EXSTYLE;
		DWORD new_style, orig_style = GetWindowLong(control_window, style_index);
		// +/-/^ are used instead of |&^ because the latter is confusing, namely that & really means &=~style, etc.
		if (!_tcschr(_T("+-^"), *aValue))  // | and & are used instead of +/- to allow +/- to have their native function.
			new_style = ATOU(aValue); // No prefix, so this new style will entirely replace the current style.
		else
		{
			++aValue; // Won't work combined with next line, due to next line being a macro that uses the arg twice.
			DWORD style_change = ATOU(aValue);
			switch(aValue[-1])
			{
			case '+': new_style = orig_style | style_change; break;
			case '-': new_style = orig_style & ~style_change; break;
			case '^': new_style = orig_style ^ style_change; break;
			}
		}
		if (new_style == orig_style) // v1.0.45.04: Ask for an unnecessary change (i.e. one that is already in effect) should not be considered an error.
			return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		// Currently, BM_SETSTYLE is not done when GetClassName() says that the control is a button/checkbox/groupbox.
		// This is because the docs for BM_SETSTYLE don't contain much, if anything, that anyone would ever
		// want to change.
		SetLastError(0); // Prior to SetWindowLong(), as recommended by MSDN.
		if (SetWindowLong(control_window, style_index, new_style) || !GetLastError()) // This is the precise way to detect success according to MSDN.
		{
			// Even if it indicated success, sometimes it failed anyway.  Find out for sure:
			if (GetWindowLong(control_window, style_index) != orig_style) // Even a partial change counts as a success.
			{
				InvalidateRect(control_window, NULL, TRUE); // Quite a few styles require this to become visibly manifest.
				return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
			}
		}
		goto error; // As documented, DoControlDelay is not done for these.
	}

	case CONTROL_CMD_SHOWDROPDOWN:
	case CONTROL_CMD_HIDEDROPDOWN:
		// CB_SHOWDROPDOWN: Although the return value (dwResult) is always TRUE, SendMessageTimeout()
		// will return failure if it times out:
		if (!SendMessageTimeout(control_window, CB_SHOWDROPDOWN
			, (WPARAM)(control_cmd == CONTROL_CMD_SHOWDROPDOWN)
			, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			goto error;
		break;

	case CONTROL_CMD_TABLEFT:
	case CONTROL_CMD_TABRIGHT: // must be a Tab Control
		key_count = *aValue ? ATOI(aValue) : 1;
		vk = (control_cmd == CONTROL_CMD_TABLEFT) ? VK_LEFT : VK_RIGHT;
		lparam = (LPARAM)(vk_to_sc(vk) << 16);
		for (int i = 0; i < key_count; ++i)
		{
			// DoControlDelay isn't done for every iteration because it seems likely that
			// the Sleep(0) will take care of things.
			PostMessage(control_window, WM_KEYDOWN, vk, lparam | 0x00000001);
			SLEEP_WITHOUT_INTERRUPTION(0); // Au3 uses a Sleep(0).
			PostMessage(control_window, WM_KEYUP, vk, lparam | 0xC0000001);
		}
		break;

	case CONTROL_CMD_ADD:
		if (!*aControl) // Fix for v1.0.46.11: If aControl is blank, the control ID came in via a WinTitle of "ahk_id xxx".
		{
			GetClassName(control_window, temp_buf, _countof(temp_buf));
			aControl = temp_buf;
		}
		if (tcscasestr(aControl, _T("Combo"))) // v1.0.42: Changed to strcasestr vs. !strnicmp for TListBox/TComboBox.
			msg = CB_ADDSTRING;
		else if (tcscasestr(aControl, _T("List")))
			msg = LB_ADDSTRING;
		else
			goto error;  // Must be ComboBox or ListBox.
		if (!SendMessageTimeout(control_window, msg, 0, (LPARAM)aValue, SMTO_ABORTIFHUNG, 2000, &dwResult))
			goto error;
		if (dwResult == CB_ERR || dwResult == CB_ERRSPACE) // General error or insufficient space to store it.
			// CB_ERR == LB_ERR
			goto error;
		break;

	case CONTROL_CMD_DELETE:
		if (!*aValue)
			goto error;
		control_index = ATOI(aValue) - 1;
		if (control_index < 0)
			goto error;
		if (!*aControl) // Fix for v1.0.46.11: If aControl is blank, the control ID came in via a WinTitle of "ahk_id xxx".
		{
			GetClassName(control_window, temp_buf, _countof(temp_buf));
			aControl = temp_buf;
		}
		if (tcscasestr(aControl, _T("Combo"))) // v1.0.42: Changed to strcasestr vs. strnicmp for TListBox/TComboBox.
			msg = CB_DELETESTRING;
		else if (tcscasestr(aControl, _T("List")))
			msg = LB_DELETESTRING;
		else
			goto error;  // Must be ComboBox or ListBox.
		if (!SendMessageTimeout(control_window, msg, (WPARAM)control_index, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			goto error;
		if (dwResult == CB_ERR)  // CB_ERR == LB_ERR
			goto error;
		break;

	case CONTROL_CMD_CHOOSE:
		if (!*aValue)
			goto error;
		control_index = ATOI(aValue) - 1;
		if (control_index < 0)
			goto error;
		if (!*aControl) // Fix for v1.0.46.11: If aControl is blank, the control ID came in via a WinTitle of "ahk_id xxx".
		{
			GetClassName(control_window, temp_buf, _countof(temp_buf));
			aControl = temp_buf;
		}
		if (tcscasestr(aControl, _T("Combo"))) // v1.0.42: Changed to strcasestr vs. strnicmp for TListBox/TComboBox.
		{
			msg = CB_SETCURSEL;
			x_msg = CBN_SELCHANGE;
			y_msg = CBN_SELENDOK;
		}
		else if (tcscasestr(aControl, _T("List")))
		{
			if (GetWindowLong(control_window, GWL_STYLE) & (LBS_EXTENDEDSEL|LBS_MULTIPLESEL))
				msg = LB_SETSEL;
			else // single-select listbox
				msg = LB_SETCURSEL;
			x_msg = LBN_SELCHANGE;
			y_msg = LBN_DBLCLK;
		}
		else
			goto error;
		if (msg == LB_SETSEL) // Multi-select, so use the cumulative method.
		{
			if (!SendMessageTimeout(control_window, msg, TRUE, control_index, SMTO_ABORTIFHUNG, 2000, &dwResult))
				goto error;
		}
		else // ComboBox or single-select ListBox.
			if (!SendMessageTimeout(control_window, msg, control_index, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
				goto error;
		if (dwResult == CB_ERR)  // CB_ERR == LB_ERR
			goto error;
		goto notify_parent;

	case CONTROL_CMD_CHOOSESTRING:
		if (!*aControl) // Fix for v1.0.46.11: If aControl is blank, the control ID came in via a WinTitle of "ahk_id xxx".
		{
			GetClassName(control_window, temp_buf, _countof(temp_buf));
			aControl = temp_buf;
		}
		if (tcscasestr(aControl, _T("Combo"))) // v1.0.42: Changed to strcasestr vs. strnicmp for TListBox/TComboBox.
		{
			msg = CB_SELECTSTRING;
			x_msg = CBN_SELCHANGE;
			y_msg = CBN_SELENDOK;
		}
		else if (tcscasestr(aControl, _T("List")))
		{
			if (GetWindowLong(control_window, GWL_STYLE) & (LBS_EXTENDEDSEL|LBS_MULTIPLESEL))
				msg = LB_FINDSTRING;
			else // single-select listbox
				msg = LB_SELECTSTRING;
			x_msg = LBN_SELCHANGE;
			y_msg = LBN_DBLCLK;
		}
		else
			goto error;  // Must be ComboBox or ListBox.
		if (msg == LB_FINDSTRING) // Multi-select ListBox (LB_SELECTSTRING is not supported by these).
		{
			DWORD_PTR item_index;
			if (!SendMessageTimeout(control_window, msg, -1, (LPARAM)aValue, SMTO_ABORTIFHUNG, 2000, &item_index)
				|| item_index == LB_ERR
				|| !SendMessageTimeout(control_window, LB_SETSEL, TRUE, item_index, SMTO_ABORTIFHUNG, 2000, &dwResult)
				|| dwResult == LB_ERR) // Relies on short-circuit boolean.
				goto error;
		}
		else // ComboBox or single-select ListBox.
			if (!SendMessageTimeout(control_window, msg, -1, (LPARAM)aValue, SMTO_ABORTIFHUNG, 2000, &dwResult)
				|| dwResult == CB_ERR) // CB_ERR == LB_ERR
				goto error;
	notify_parent:
		if (   !(immediate_parent = GetParent(control_window))   )
			goto error;
		SetLastError(0); // Must be done to differentiate between success and failure when control has ID 0.
		control_id = GetDlgCtrlID(control_window);
		if (!control_id && GetLastError()) // Both conditions must be checked (see above).
			goto error; // Avoid sending the notification in case some other control has ID 0.
		// Proceed even if control_id == 0, since some applications are known to
		// utilize the notification in that case (e.g. Notepad's Save As dialog).
		if (!SendMessageTimeout(immediate_parent, WM_COMMAND, (WPARAM)MAKELONG(control_id, x_msg)
			, (LPARAM)control_window, SMTO_ABORTIFHUNG, 2000, &dwResult))
			goto error;
		if (!SendMessageTimeout(immediate_parent, WM_COMMAND, (WPARAM)MAKELONG(control_id, y_msg)
			, (LPARAM)control_window, SMTO_ABORTIFHUNG, 2000, &dwResult))
			goto error;
		// Otherwise break and do the end-function processing.
		break;

	case CONTROL_CMD_EDITPASTE:
		if (!SendMessageTimeout(control_window, EM_REPLACESEL, TRUE, (LPARAM)aValue, SMTO_ABORTIFHUNG, 2000, &dwResult))
			goto error;
		// Note: dwResult is not used by EM_REPLACESEL since it doesn't return a value.
		break;
	} // switch()

	DoControlDelay;  // Seems safest to do this for all of these commands.
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.

error:
	return SetErrorLevelOrThrow();
}



ResultType Line::ControlGet(LPTSTR aCmd, LPTSTR aValue, LPTSTR aControl, LPTSTR aTitle, LPTSTR aText
	, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	Var &output_var = *OUTPUT_VAR;
	ControlGetCmds control_cmd = ConvertControlGetCmd(aCmd);
	// Since command names are validated at load-time, this only happens if the command name
	// was contained in a variable reference.  Since that is very rare, just set ErrorLevel
	// and return:
	if (control_cmd == CONTROLGET_CMD_INVALID)
		goto error;

	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		goto error;
	HWND control_window = ControlExist(target_window, aControl); // This can return target_window itself for cases such as ahk_id %ControlHWND%.
	if (!control_window)
		goto error;

	DWORD_PTR dwResult, index, length, item_length, u, item_count;
	DWORD start, end;
	UINT msg, x_msg, y_msg;
	int control_index;
	TCHAR *cp, *dyn_buf, temp_buf[32];

	switch(control_cmd)
	{
	case CONTROLGET_CMD_CHECKED: //Must be a Button
		if (!SendMessageTimeout(control_window, BM_GETCHECK, 0, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			goto error;
		output_var.Assign(dwResult == BST_CHECKED ? _T("1") : _T("0"));
		break;

	case CONTROLGET_CMD_ENABLED:
		output_var.Assign(IsWindowEnabled(control_window) ? _T("1") : _T("0"));
		break;

	case CONTROLGET_CMD_VISIBLE:
		output_var.Assign(IsWindowVisible(control_window) ? _T("1") : _T("0"));
		break;

	case CONTROLGET_CMD_TAB: // must be a Tab Control
		if (!SendMessageTimeout(control_window, TCM_GETCURSEL, 0, 0, SMTO_ABORTIFHUNG, 2000, &index) || index == -1) // Relies on short-circuit boolean order.
			goto error;
		output_var.Assign(index + 1);
		break;

	case CONTROLGET_CMD_FINDSTRING:
		if (!*aControl) // Fix for v1.0.46.11: If aControl is blank, the control ID came in via a WinTitle of "ahk_id xxx".
		{
			GetClassName(control_window, temp_buf, _countof(temp_buf));
			aControl = temp_buf;
		}
		if (tcscasestr(aControl, _T("Combo"))) // v1.0.42: Changed to strcasestr vs. strnicmp for TListBox/TComboBox.
			msg = CB_FINDSTRINGEXACT;
		else if (tcscasestr(aControl, _T("List")))
			msg = LB_FINDSTRINGEXACT;
		else // Must be ComboBox or ListBox
			goto error;
		if (!SendMessageTimeout(control_window, msg, -1, (LPARAM)aValue, SMTO_ABORTIFHUNG, 2000, &index)
			|| index == CB_ERR) // CB_ERR == LB_ERR
			goto error;
		output_var.Assign(index + 1);
		break;

	case CONTROLGET_CMD_CHOICE:
		if (!*aControl) // Fix for v1.0.46.11: If aControl is blank, the control ID came in via a WinTitle of "ahk_id xxx".
		{
			GetClassName(control_window, temp_buf, _countof(temp_buf));
			aControl = temp_buf;
		}
		if (tcscasestr(aControl, _T("Combo"))) // v1.0.42: Changed to strcasestr vs. strnicmp for TListBox/TComboBox.
		{
			msg = CB_GETCURSEL;
			x_msg = CB_GETLBTEXTLEN;
			y_msg = CB_GETLBTEXT;
		}
		else if (tcscasestr(aControl, _T("List")))
		{
			msg = LB_GETCURSEL;
			x_msg = LB_GETTEXTLEN;
			y_msg = LB_GETTEXT;
		}
		else // Must be ComboBox or ListBox
			return output_var.Assign();  // Let ErrorLevel tell the story.
		if (!SendMessageTimeout(control_window, msg, 0, 0, SMTO_ABORTIFHUNG, 2000, &index)
			|| index == CB_ERR  // CB_ERR == LB_ERR.  There is no selection (or very rarely, some other type of problem).
			|| !SendMessageTimeout(control_window, x_msg, (WPARAM)index, 0, SMTO_ABORTIFHUNG, 2000, &length)
			|| length == CB_ERR)  // CB_ERR == LB_ERR
			goto error; // Above relies on short-circuit boolean order.
		// In unusual cases, MSDN says the indicated length might be longer than it actually winds up
		// being when the item's text is retrieved.  This should be harmless, since there are many
		// other precedents where a variable is sized to something larger than it winds up carrying.
		// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
		// this call will set up the clipboard for writing:
		if (output_var.AssignString(NULL, (VarSizeType)length) != OK) // It already displayed the error.
			return FAIL;
		if (!SendMessageTimeout(control_window, y_msg, (WPARAM)index, (LPARAM)output_var.Contents()
			, SMTO_ABORTIFHUNG, 2000, &length)
			|| length == CB_ERR) // Probably impossible given the way it was called above.  Also, CB_ERR == LB_ERR. Relies on short-circuit boolean order.
		{
			output_var.Close();
			goto error;
		}
		output_var.Close(); // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_CLIPBOARD.
		output_var.SetCharLength(length);  // Update to actual vs. estimated length.
		break;

	case CONTROLGET_CMD_LIST:
		if (!*aControl) // Fix for v1.0.46.11: If aControl is blank, the control ID came in via a WinTitle of "ahk_id xxx".
		{
			GetClassName(control_window, temp_buf, _countof(temp_buf));
			aControl = temp_buf;
		}
		//if (!_tcsnicmp(aControl, _T("SysListView32"), 13)) // Tried strcasestr(aControl, "ListView") to get it to work with IZArc's Delphi TListView1, but none of the modes or options worked.
		if (tcscasestr(aControl, _T("SysListView32"))) // Some users said this works with "WindowsForms10.SysListView32"
			return ControlGetListView(output_var, control_window, aValue); // It will also set ErrorLevel to "success" if successful.
		// This is done here as the special LIST sub-command rather than just being built into
		// ControlGetText because ControlGetText already has a function for ComboBoxes: it fetches
		// the current selection.  Although ListBox does not have such a function, it seem best
		// to consolidate both methods here.
		if (tcscasestr(aControl, _T("Combo"))) // v1.0.42: Changed to strcasestr vs. strnicmp for TListBox/TComboBox.
		{
			msg = CB_GETCOUNT;
			x_msg = CB_GETLBTEXTLEN;
			y_msg = CB_GETLBTEXT;
		}
		else if (tcscasestr(aControl, _T("List")))
		{
			msg = LB_GETCOUNT;
			x_msg = LB_GETTEXTLEN;
			y_msg = LB_GETTEXT;
		}
		else // Must be ComboBox or ListBox
			goto error;
		if (!(SendMessageTimeout(control_window, msg, 0, 0, SMTO_ABORTIFHUNG, 5000, &item_count))
			|| item_count < 1) // No items in ListBox/ComboBox or there was a problem getting the count.
			goto error;
		// Calculate the length of delimited list of items.  Length is initialized to provide enough
		// room for each item's delimiter (the last item does not have a delimiter).
		for (length = item_count - 1, u = 0; u < item_count; ++u)
		{
			if (!SendMessageTimeout(control_window, x_msg, u, 0, SMTO_ABORTIFHUNG, 5000, &item_length)
				|| item_length == LB_ERR) // Note that item_length is legitimately zero for a blank item in the list.
				return output_var.Assign();  // Let ErrorLevel tell the story.
			length += item_length;
		}
		// In unusual cases, MSDN says the indicated length might be longer than it actually winds up
		// being when the item's text is retrieved.  This should be harmless, since there are many
		// other precedents where a variable is sized to something larger than it winds up carrying.
		// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
		// this call will set up the clipboard for writing:
		if (output_var.AssignString(NULL, (VarSizeType)length, true, true) != OK)
			return FAIL;  // It already displayed the error.
		for (cp = output_var.Contents(), length = item_count - 1, u = 0; u < item_count; ++u)
		{
			if (SendMessageTimeout(control_window, y_msg, (WPARAM)u, (LPARAM)cp, SMTO_ABORTIFHUNG, 5000, &item_length)
				&& item_length != LB_ERR)
			{
				length += item_length; // Accumulate actual vs. estimated length.
				cp += item_length;  // Point it to the terminator in preparation for the next write.
			}
			//else do nothing, just consider this to be a blank item so that the process can continue.
			if (u < item_count - 1)
				*cp++ = '\n'; // Add delimiter after each item except the last (helps parsing loop).
			// Above: In this case, seems better to use \n rather than pipe as default delimiter in case
			// the listbox/combobox contains any real pipes.
		}
		output_var.Close(); // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_CLIPBOARD.
		output_var.SetCharLength(length);  // Update it to the actual length, which can vary from the estimate.
		break;

	case CONTROLGET_CMD_LINECOUNT:  //Must be an Edit
		// MSDN: "If the control has no text, the return value is 1. The return value will never be less than 1."
		if (!SendMessageTimeout(control_window, EM_GETLINECOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			goto error;
		output_var.Assign(dwResult);
		break;

	case CONTROLGET_CMD_CURRENTLINE:
		if (!SendMessageTimeout(control_window, EM_LINEFROMCHAR, -1, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			goto error;
		output_var.Assign(dwResult + 1);
		break;

	case CONTROLGET_CMD_CURRENTCOL:
	{
		DWORD_PTR line_number;
		// The dwResult from the first msg below is not useful and is not checked.
		if (   !SendMessageTimeout(control_window, EM_GETSEL, (WPARAM)&start, (LPARAM)&end, SMTO_ABORTIFHUNG, 2000, &dwResult)
			|| !SendMessageTimeout(control_window, EM_LINEFROMCHAR, (WPARAM)start, 0, SMTO_ABORTIFHUNG, 2000, &line_number)   )
			goto error;
		if (!line_number) // Since we're on line zero, the column number is simply start+1.
		{
			output_var.Assign(start + 1);  // +1 to convert from zero based.
			break; // Fall out of the switch so that ErrorLevel will be set to 0 (no error).
		}
		// Au3: Decrement the character index until the row changes.  Difference between this
		// char index and original is the column:
		DWORD_PTR start_orig = start;  // Au3: the character index
		for (;;)
		{
			if (!SendMessageTimeout(control_window, EM_LINEFROMCHAR, (WPARAM)start, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
				goto error;
			if (dwResult != line_number)
				break;
			--start;
		}
		output_var.Assign((int)(start_orig - start));
		break;
	}

	case CONTROLGET_CMD_LINE:
		if (!*aValue)
			goto error;
		control_index = ATOI(aValue) - 1;
		if (control_index < 0)
			goto error;
		// jackieku: 32768 * sizeof(wchar_t) = 65536, which can not be stored in a unsigned 16bit integer.
		dyn_buf = (LPTSTR)talloca(32767); // 32768 is the size Au3 uses for GETLINE and such.
		*(LPWORD)dyn_buf = 32767; // EM_GETLINE requires first word of string to be set to its size.
		if (   !SendMessageTimeout(control_window, EM_GETLINE, (WPARAM)control_index, (LPARAM)dyn_buf, SMTO_ABORTIFHUNG, 2000, &dwResult)
			|| !dwResult   ) // due to the specified line number being greater than the number of lines in the edit control.
			goto error;
		dyn_buf[dwResult] = '\0'; // Ensure terminated since the API might not do it in some cases.
		output_var.Assign(dyn_buf);
		break;

	case CONTROLGET_CMD_SELECTED: // Must be an Edit.
		// Note: The RichEdit controls of certain apps such as Metapad don't return the right selection
		// with this technique.  Au3 has the same problem with them, so for now it's just documented here
		// as a limitation.
		if (!SendMessageTimeout(control_window, EM_GETSEL, (WPARAM)&start, (LPARAM)&end, SMTO_ABORTIFHUNG, 2000, &dwResult))
			goto error;
		// The above sets start to be the zero-based position of the start of the selection (similar for end).
		// If there is no selection, start and end will be equal, at least in the edit controls I tried it with.
		// The dwResult from the above is not useful and is not checked.
		if (start == end) // Unlike Au3, it seems best to consider a blank selection to be a non-error.
		{
			output_var.Assign();
			break; // Fall out of the switch so that ErrorLevel will be set to 0 (no error).
		}
		// Dynamic memory is used because must get all the control's text so that just the selected region
		// can be cropped out and assigned to the output variable.  Otherwise, output_var might
		// have to be sized much larger than it would need to be:
		if (   !SendMessageTimeout(control_window, WM_GETTEXTLENGTH, 0, 0, SMTO_ABORTIFHUNG, 2000, &length)
			|| !length  // Since the above didn't return for start == end, this is an error because we have a selection of non-zero length, but no text to go with it!
			|| !(dyn_buf = tmalloc(length + 1))   ) // Relies on short-circuit boolean order.
			goto error;
		if (   !SendMessageTimeout(control_window, WM_GETTEXT, (WPARAM)(length + 1), (LPARAM)dyn_buf, SMTO_ABORTIFHUNG, 2000, &length)
			|| !length || end > length   )
		{
			// The first check above is reveals a problem (ErrorLevel = 1) since the length
			// is unexpectedly zero (above implied it shouldn't be).  The second check is also
			// a problem because the end of the selection should not be beyond length of text
			// that was retrieved.
			free(dyn_buf);
			goto error;
		}
		dyn_buf[end] = '\0'; // Terminate the string at the end of the selection.
		output_var.Assign(dyn_buf + start);
		free(dyn_buf);
		break;

	case CONTROLGET_CMD_STYLE:
		// Seems best to always format as hex, since it has more human-readable meaning then:
		_stprintf(temp_buf, _T("0x%08X"), GetWindowLong(control_window, GWL_STYLE));
		output_var.Assign(temp_buf);
		break;

	case CONTROLGET_CMD_EXSTYLE:
		// Seems best to always format as hex, since it has more human-readable meaning then:
		_stprintf(temp_buf, _T("0x%08X"), GetWindowLong(control_window, GWL_EXSTYLE));
		output_var.Assign(temp_buf);
		break;

	case CONTROLGET_CMD_HWND:
		// The terminology "HWND" was chosen rather than "ID" to avoid confusion with a control's
		// dialog ID (as retrieved by GetDlgCtrlID).  This also reserves the word ID for possible
		// use with the control's Dialog ID in future versions.
		output_var.AssignHWND(control_window);
		break;
	}

	// Note that ControlDelay is not done for the Get type commands, because it seems unnecessary.
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.

error:
	output_var.Assign();
	return SetErrorLevelOrThrow();
}



ResultType Line::URLDownloadToFile(LPTSTR aURL, LPTSTR aFilespec)
{
	// Check that we have IE3 and access to wininet.dll
	HINSTANCE hinstLib = LoadLibrary(_T("wininet"));
	if (!hinstLib)
		return SetErrorLevelOrThrow();

	typedef HINTERNET (WINAPI *MyInternetOpen)(LPCTSTR, DWORD, LPCTSTR, LPCTSTR, DWORD dwFlags);
	typedef HINTERNET (WINAPI *MyInternetOpenUrl)(HINTERNET hInternet, LPCTSTR, LPCTSTR, DWORD, DWORD, LPDWORD);
	typedef BOOL (WINAPI *MyInternetCloseHandle)(HINTERNET);
	typedef BOOL (WINAPI *MyInternetReadFileEx)(HINTERNET, LPINTERNET_BUFFERSA, DWORD, DWORD);
	typedef BOOL (WINAPI *MyInternetReadFile)(HINTERNET, LPVOID, DWORD, LPDWORD);

	#ifndef INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY
		#define INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY 4
	#endif

	// Get the address of all the functions we require.  It's done this way in case the system
	// lacks MSIE v3.0+, in which case the app would probably refuse to launch at all:
 	MyInternetOpen lpfnInternetOpen = (MyInternetOpen)GetProcAddress(hinstLib, "InternetOpen" WINAPI_SUFFIX);
	MyInternetOpenUrl lpfnInternetOpenUrl = (MyInternetOpenUrl)GetProcAddress(hinstLib, "InternetOpenUrl" WINAPI_SUFFIX);
	MyInternetCloseHandle lpfnInternetCloseHandle = (MyInternetCloseHandle)GetProcAddress(hinstLib, "InternetCloseHandle");
	MyInternetReadFileEx lpfnInternetReadFileEx = (MyInternetReadFileEx)GetProcAddress(hinstLib, "InternetReadFileExA"); // InternetReadFileExW() appears unimplemented prior to Windows 7, so always use InternetReadFileExA().
	MyInternetReadFile lpfnInternetReadFile = (MyInternetReadFile)GetProcAddress(hinstLib, "InternetReadFile"); // Called unconditionally to reduce code size and because the time required is likely insignificant compared to network latency.
	if (!(lpfnInternetOpen && lpfnInternetOpenUrl && lpfnInternetCloseHandle && lpfnInternetReadFileEx && lpfnInternetReadFile))
	{
		FreeLibrary(hinstLib);
		return SetErrorLevelOrThrow();
	}

	// v1.0.44.07: Set default to INTERNET_FLAG_RELOAD vs. 0 because the vast majority of usages would want
	// the file to be retrieved directly rather than from the cache.
	// v1.0.46.04: Added more no-cache flags because otherwise, it definitely falls back to the cache if
	// the remote server doesn't respond (and perhaps other errors), which defeats the ability to use
	// UrlDownloadToFile for uptime/server monitoring.  Also, in spite of what MSDN says, it seems nearly
	// certain based on other sources that more than one flag is supported.  Someone also mentioned that
	// INTERNET_FLAG_CACHE_IF_NET_FAIL is related to this, but there's no way to specify it in these
	// particular calls, and it's the opposite of the desired behavior anyway; so it seems impossible to
	// turn it off explicitly.
	DWORD flags_for_open_url = INTERNET_FLAG_RELOAD | INTERNET_FLAG_NO_CACHE_WRITE;
	aURL = omit_leading_whitespace(aURL);
	if (*aURL == '*') // v1.0.44.07: Provide an option to override flags_for_open_url.
	{
		flags_for_open_url = ATOU(++aURL);
		LPTSTR cp;
		if (cp = StrChrAny(aURL, _T(" \t"))) // Find first space or tab.
			aURL = omit_leading_whitespace(cp);
	}

	// Open the internet session. v1.0.45.03: Provide a non-NULL user-agent because  some servers reject
	// requests that lack a user-agent.  Furthermore, it's more professional to have one, in which case it
	// should probably be kept as simple and unchanging as possible.  Using something like the script's name
	// as the user agent (even if documented) seems like a bad idea because it might contain personal/sensitive info.
	HINTERNET hInet = lpfnInternetOpen(_T("AutoHotkey"), INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY, NULL, NULL, 0);
	if (!hInet)
	{
		FreeLibrary(hinstLib);
		return SetErrorLevelOrThrow();
	}

	// Open the required URL
	HINTERNET hFile = lpfnInternetOpenUrl(hInet, aURL, NULL, 0, flags_for_open_url, 0);
	if (!hFile)
	{
		lpfnInternetCloseHandle(hInet);
		FreeLibrary(hinstLib);
		return SetErrorLevelOrThrow();
	}

	// Open our output file
	FILE *fptr = _tfopen(aFilespec, _T("wb"));	// Open in binary write/destroy mode
	if (!fptr)
	{
		lpfnInternetCloseHandle(hFile);
		lpfnInternetCloseHandle(hInet);
		FreeLibrary(hinstLib);
		return SetErrorLevelOrThrow();
	}

	BYTE bufData[1024 * 1]; // v1.0.44.11: Reduced from 8 KB to alleviate GUI window lag during UrlDownloadtoFile.  Testing shows this reduction doesn't affect performance on high-speed downloads (in fact, downloads are slightly faster; I tested two sites, one at 184 KB/s and the other at 380 KB/s).  It might affect slow downloads, but that seems less likely so wasn't tested.
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
	if (*aURL == 'h' || *aURL == 'H')
	{
		while (result = lpfnInternetReadFileEx(hFile, &buffers, IRF_NO_WAIT, NULL)) // Assign
		{
			if (!buffers.dwBufferLength) // Transfer is complete.
				break;
			LONG_OPERATION_UPDATE  // Done in between the net-read and the file-write to improve avg. responsiveness.
			fwrite(bufData, buffers.dwBufferLength, 1, fptr);
			buffers.dwBufferLength = sizeof(bufData);  // Reset buffer capacity for next iteration.
		}
	}
	else // v1.0.48.04: This section adds support for FTP and perhaps Gopher by using InternetReadFile() instead of InternetReadFileEx().
	{
		DWORD number_of_bytes_read;
		while (result = lpfnInternetReadFile(hFile, bufData, sizeof(bufData), &number_of_bytes_read))
		{
			if (!number_of_bytes_read)
				break;
			LONG_OPERATION_UPDATE
			fwrite(bufData, number_of_bytes_read, 1, fptr);
		}
	}
	// Close internet session:
	lpfnInternetCloseHandle(hFile);
	lpfnInternetCloseHandle(hInet);
	FreeLibrary(hinstLib); // Only after the above.
	// Close output file:
	fclose(fptr);

	if (!result) // An error occurred during the transfer.
		DeleteFile(aFilespec);  // Delete damaged/incomplete file.
	return SetErrorLevelOrThrowBool(!result);
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



ResultType Line::FileSelectFolder(LPTSTR aRootDir, LPTSTR aOptions, LPTSTR aGreeting)
// Since other script threads can interrupt this command while it's running, it's important that
// the command not refer to sArgDeref[] and sArgVar[] anytime after an interruption becomes possible.
// This is because an interrupting thread usually changes the values to something inappropriate for this thread.
{
	Var &output_var = *OUTPUT_VAR; // Must be resolved early.  See comment above.
	if (!output_var.Assign())  // Initialize the output variable.
		return FAIL;

	if (g_nFolderDialogs >= MAX_FOLDERDIALOGS)
	{
		// Have a maximum to help prevent runaway hotkeys due to key-repeat feature, etc.
		return LineError(_T("The maximum number of Folder Dialogs has been reached."));
	}

	LPMALLOC pMalloc;
    if (SHGetMalloc(&pMalloc) != NOERROR)	// Initialize
		return SetErrorLevelOrThrow();

	// v1.0.36.03: Support initial folder, which is different than the root folder because the root only
	// controls the origin point (above which the control cannot navigate).
	LPTSTR initial_folder;
	TCHAR root_dir[MAX_PATH*2 + 5];  // Up to two paths might be present inside, including an asterisk and spaces between them.
	tcslcpy(root_dir, aRootDir, _countof(root_dir)); // Make a modifiable copy.
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
	if (aGreeting && *aGreeting)
		tcslcpy(greeting, aGreeting, _countof(greeting));
	else
		sntprintf(greeting, _countof(greeting), _T("Select Folder - %s"), g_script.mFileName);
	bi.lpszTitle = greeting;

	DWORD options = *aOptions ? ATOI(aOptions) : FSF_ALLOW_CREATE;
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
		// Due to rarity and because there doesn't seem to be any way to detect it,
		// no exception is thrown when the function fails.  Instead, we just assume
		// that the user pressed CANCEL (which should not be treated as an error):
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	*Result = '\0';  // Reuse this var, this time to old the result of the below:
	SHGetPathFromIDList(lpItemIDList, Result);
	pMalloc->Free(lpItemIDList);
	pMalloc->Release();

	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return output_var.Assign(Result);
}



ResultType Line::FileGetShortcut(LPTSTR aShortcutFile) // Credited to Holger <Holger.Kotsch at GMX de>.
{
	Var *output_var_target = ARGVAR2; // These might be omitted in the parameter list, so it's okay if 
	Var *output_var_dir = ARGVAR3;    // they resolve to NULL.  Also, load-time validation has ensured
	Var *output_var_arg = ARGVAR4;    // that these are valid output variables (e.g. not built-in vars).
	Var *output_var_desc = ARGVAR5;   // Load-time validation has ensured that these are valid output variables (e.g. not built-in vars).
	Var *output_var_icon = ARGVAR6;
	Var *output_var_icon_idx = ARGVAR7;
	Var *output_var_show_state = ARGVAR8;

	// For consistency with the behavior of other commands, the output variables are initialized to blank
	// so that there is another way to detect failure:
	if (output_var_target) output_var_target->Assign();
	if (output_var_dir) output_var_dir->Assign();
	if (output_var_arg) output_var_arg->Assign();
	if (output_var_desc) output_var_desc->Assign();
	if (output_var_icon) output_var_icon->Assign();
	if (output_var_icon_idx) output_var_icon_idx->Assign();
	if (output_var_show_state) output_var_show_state->Assign();

	bool bSucceeded = false;

	if (!Util_DoesFileExist(aShortcutFile))
		goto error;

	CoInitialize(NULL);
	IShellLink *psl;

	if (SUCCEEDED(CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID *)&psl)))
	{
		IPersistFile *ppf;
		if (SUCCEEDED(psl->QueryInterface(IID_IPersistFile, (LPVOID *)&ppf)))
		{
#ifdef UNICODE
			if (SUCCEEDED(ppf->Load(aShortcutFile, 0)))
#else
			WCHAR wsz[MAX_PATH+1]; // +1 hasn't been explained, but is retained in case it's needed.
			ToWideChar(aShortcutFile, wsz, MAX_PATH+1); // Dest. size is in wchars, not bytes.
			if (SUCCEEDED(ppf->Load((const WCHAR*)wsz, 0)))
#endif
			{
				TCHAR buf[MAX_PATH+1];
				int icon_index, show_cmd;

				if (output_var_target)
				{
					psl->GetPath(buf, MAX_PATH, NULL, SLGP_UNCPRIORITY);
					output_var_target->Assign(buf);
				}
				if (output_var_dir)
				{
					psl->GetWorkingDirectory(buf, MAX_PATH);
					output_var_dir->Assign(buf);
				}
				if (output_var_arg)
				{
					psl->GetArguments(buf, MAX_PATH);
					output_var_arg->Assign(buf);
				}
				if (output_var_desc)
				{
					psl->GetDescription(buf, MAX_PATH); // Testing shows that the OS limits it to 260 characters.
					output_var_desc->Assign(buf);
				}
				if (output_var_icon || output_var_icon_idx)
				{
					psl->GetIconLocation(buf, MAX_PATH, &icon_index);
					if (output_var_icon)
						output_var_icon->Assign(buf);
					if (output_var_icon_idx)
						if (*buf)
							output_var_icon_idx->Assign(icon_index + 1);  // Convert from 0-based to 1-based for consistency with the Menu command, etc.
						else
							output_var_icon_idx->Assign(); // Make it blank to indicate that there is none.
				}
				if (output_var_show_state)
				{
					psl->GetShowCmd(&show_cmd);
					output_var_show_state->Assign(show_cmd);
					// For the above, decided not to translate them to Max/Min/Normal since other
					// show-state numbers might be supported in the future (or are already).  In other
					// words, this allows the flexibility to specify some number other than 1/3/7 when
					// creating the shortcut in case it happens to work.  Of course, that applies only
					// to FileCreateShortcut, not here.  But it's done here so that this command is
					// compatible with that one.
				}
				g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
				bSucceeded = true;
			}
			ppf->Release();
		}
		psl->Release();
	}
	CoUninitialize();

	if (!bSucceeded)
		goto error;

	return OK;  // ErrorLevel might still indicate failure if one of the above calls failed.

error:
	return SetErrorLevelOrThrow();
}



ResultType Line::FileCreateShortcut(LPTSTR aTargetFile, LPTSTR aShortcutFile, LPTSTR aWorkingDir, LPTSTR aArgs
	, LPTSTR aDescription, LPTSTR aIconFile, LPTSTR aHotkey, LPTSTR aIconNumber, LPTSTR aRunState)
{
	bool bSucceeded = false;
	CoInitialize(NULL);
	IShellLink *psl;

	if (SUCCEEDED(CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID *)&psl)))
	{
		psl->SetPath(aTargetFile);
		if (*aWorkingDir)
			psl->SetWorkingDirectory(aWorkingDir);
		if (*aArgs)
			psl->SetArguments(aArgs);
		if (*aDescription)
			psl->SetDescription(aDescription);
		if (*aIconFile)
			psl->SetIconLocation(aIconFile, *aIconNumber ? ATOI(aIconNumber) - 1 : 0); // Doesn't seem necessary to validate aIconNumber as not being negative, etc.
		if (*aHotkey)
		{
			// If badly formatted, it's not a critical error, just continue.
			// Currently, only shortcuts with a CTRL+ALT are supported.
			// AutoIt3 note: Make sure that CTRL+ALT is selected (otherwise invalid)
			vk_type vk = TextToVK(aHotkey);
			if (vk)
				// Vk in low 8 bits, mods in high 8:
				psl->SetHotkey(   (WORD)vk | ((WORD)(HOTKEYF_CONTROL | HOTKEYF_ALT) << 8)   );
		}
		if (*aRunState)
			psl->SetShowCmd(ATOI(aRunState)); // No validation is done since there's a chance other numbers might be valid now or in the future.

		IPersistFile *ppf;
		if (SUCCEEDED(psl->QueryInterface(IID_IPersistFile,(LPVOID *)&ppf)))
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
			if (SUCCEEDED(ppf->Save(full_path, TRUE)))
			{
				g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
				bSucceeded = true;
			}
			ppf->Release();
		}
		psl->Release();
	}

	CoUninitialize();
	if (bSucceeded)
		return OK;
	else
		return SetErrorLevelOrThrow();
}



ResultType Line::FileRecycle(LPTSTR aFilePattern)
{
	if (!aFilePattern || !*aFilePattern)
		return SetErrorLevelOrThrow();  // Since this is probably not what the user intended.

	SHFILEOPSTRUCT FileOp;
	TCHAR szFileTemp[_MAX_PATH+2];

	// au3: Get the fullpathname - required for UNDO to work
	Util_GetFullPathName(aFilePattern, szFileTemp);

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
	return SetErrorLevelOrThrowBool(SHFileOperation(&FileOp));
}



ResultType Line::FileRecycleEmpty(LPTSTR aDriveLetter)
{
	LPCTSTR szPath = *aDriveLetter ? aDriveLetter : NULL;
	HRESULT hr = SHEmptyRecycleBin(NULL, szPath, SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND);
	return SetErrorLevelOrThrowBool(hr != S_OK);
}



ResultType Line::FileGetVersion(LPTSTR aFilespec)
{
	OUTPUT_VAR->Assign(); // Init to be blank, in case of failure.

	if (!aFilespec || !*aFilespec)
		return SetErrorsOrThrow(true, ERROR_INVALID_PARAMETER); // Error out, since this is probably not what the user intended.

	DWORD dwUnused, dwSize;
	if (   !(dwSize = GetFileVersionInfoSize(aFilespec, &dwUnused))   )  // No documented limit on how large it can be, so don't use _alloca().
		return SetErrorsOrThrow(true);

	BYTE *pInfo = (BYTE*)malloc(dwSize);  // Allocate the size retrieved by the above.
	VS_FIXEDFILEINFO *pFFI;
	UINT uSize;

	// Read the version resource
	if (!GetFileVersionInfo(aFilespec, 0, dwSize, (LPVOID)pInfo)
	// Locate the fixed information
		|| !VerQueryValue(pInfo, _T("\\"), (LPVOID *)&pFFI, &uSize))
	{
		g->LastError = GetLastError();
		free(pInfo);
		return SetErrorLevelOrThrow();
	}

	// extract the fields you want from pFFI
	UINT iFileMS = (UINT)pFFI->dwFileVersionMS;
	UINT iFileLS = (UINT)pFFI->dwFileVersionLS;
	TCHAR version_string[128];  // AutoIt3: 43+1 is the maximum size, but leave a little room to increase confidence.
	sntprintf(version_string, _countof(version_string), _T("%u.%u.%u.%u")
		, (iFileMS >> 16), (iFileMS & 0xFFFF), (iFileLS >> 16), (iFileLS & 0xFFFF));

	free(pInfo);

	SetErrorsOrThrow(false, 0); // Indicate success.
	return OUTPUT_VAR->Assign(version_string);
}



bool Line::Util_CopyDir(LPCTSTR szInputSource, LPCTSTR szInputDest, bool bOverwrite)
{
	// Get the fullpathnames and strip trailing \s
	TCHAR szSource[_MAX_PATH+2];
	TCHAR szDest[_MAX_PATH+2];
	Util_GetFullPathName(szInputSource, szSource);
	Util_GetFullPathName(szInputDest, szDest);

	// Ensure source is a directory
	if (Util_IsDir(szSource) == false)
		return false;							// Nope

	// Does the destination dir exist?
	if (Util_IsDir(szDest))
	{
		if (bOverwrite == false)
			return false;
	}
	else // Although dest doesn't exist as a dir, it might be a file, which is covered below too.
	{
		// We must create the top level directory
		if (!Util_CreateDir(szDest)) // Failure is expected to happen if szDest is an existing *file*, since a dir should never be allowed to overwrite a file (to avoid accidental loss of data).
			return false;
	}

	// To work under old versions AND new version of shell32.dll the source must be specified
	// as "dir\*.*" and the destination directory must already exist... Goddamn Microsoft and their APIs...
	_tcscat(szSource, _T("\\*.*"));

	// We must also make source\dest double nulled strings for the SHFileOp API
	szSource[_tcslen(szSource)+1] = '\0';	
	szDest[_tcslen(szDest)+1] = '\0';	

	// Setup the struct
	SHFILEOPSTRUCT FileOp = {0};
	FileOp.pFrom = szSource;
	FileOp.pTo = szDest;
	FileOp.wFunc = FO_COPY;
	FileOp.fFlags = FOF_SILENT | FOF_NOCONFIRMMKDIR | FOF_NOCONFIRMATION | FOF_NOERRORUI; // FOF_NO_UI ("perform the operation with no user input") is not present for in case it would break compatibility somehow, and because the other flags already present seem to make its behavior implicit.  Also, unlike FileMoveDir, FOF_MULTIDESTFILES never seems to be needed.
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



bool Line::Util_MoveDir(LPCTSTR szInputSource, LPCTSTR szInputDest, int OverwriteMode)
{
	// Get the fullpathnames and strip trailing \s
	TCHAR szSource[_MAX_PATH+2];
	TCHAR szDest[_MAX_PATH+2];
	Util_GetFullPathName(szInputSource, szSource);
	Util_GetFullPathName(szInputDest, szDest);

	// Ensure source is a directory
	if (Util_IsDir(szSource) == false)
		return false;							// Nope

	// Does the destination dir exist?
	DWORD attr = GetFileAttributes(szDest);
	if (attr != 0xFFFFFFFF) // Destination already exists as a file or directory.
	{
		if (attr & FILE_ATTRIBUTE_DIRECTORY) // Dest already exists as a directory.
		{
			if (OverwriteMode != 1 && OverwriteMode != 2) // Overwrite Mode is "Never".  Strict validation for safety.
				return false; // For consistency, mode1 actually should move the source-dir *into* the identically name dest dir.  But for backward compatibility, this change hasn't been made.
		}
		else // Dest already exists as a file.
			return false; // Don't even attempt to overwrite a file with a dir, regardless of mode (I think SHFileOperation refuses to do it anyway).
	}

	if (Util_IsDifferentVolumes(szSource, szDest))
	{
		// If the source and dest are on different volumes then we must copy rather than move
		// as move in this case only works on some OSes.  Copy and delete (poor man's move).
		if (!Util_CopyDir(szSource, szDest, true))
			return false;
		return Util_RemoveDir(szSource, true);
	}

	// Since above didn't return, source and dest are on same volume.
	// We must also make source\dest double nulled strings for the SHFileOp API
	szSource[_tcslen(szSource)+1] = '\0';
	szDest[_tcslen(szDest)+1] = '\0';

	// Setup the struct
	SHFILEOPSTRUCT FileOp = {0};
	FileOp.pFrom = szSource;
	FileOp.pTo = szDest;
	FileOp.wFunc = FO_MOVE;
	FileOp.fFlags = FOF_SILENT | FOF_NOCONFIRMMKDIR | FOF_NOCONFIRMATION | FOF_NOERRORUI; // Set default. FOF_NO_UI ("perform the operation with no user input") is not present for in case it would break compatibility somehow, and because the other flags already present seem to make its behavior implicit.
	if (OverwriteMode == 2) // v1.0.46.07: Using the FOF_MULTIDESTFILES flag (as hinted by MSDN) overwrites/merges any existing target directory.  This logic supersedes and fixes old logic that didn't work properly when the source dir was being both renamed and moved to overwrite an existing directory.
		FileOp.fFlags |= FOF_MULTIDESTFILES;
	// All of the below left set to NULL/FALSE by the struct initializer higher above:
	//FileOp.hNameMappings			= NULL;
	//FileOp.lpszProgressTitle		= NULL;
	//FileOp.fAnyOperationsAborted	= FALSE;
	//FileOp.hwnd					= NULL;

	return !SHFileOperation(&FileOp);
}



bool Line::Util_RemoveDir(LPCTSTR szInputSource, bool bRecurse)
{
	SHFILEOPSTRUCT	FileOp;
	TCHAR			szSource[_MAX_PATH+2];

	// Get the fullpathnames and strip trailing \s
	Util_GetFullPathName(szInputSource, szSource);

	// Ensure source is a directory
	if (Util_IsDir(szSource) == false)
		return false;							// Nope

	// If recursion not on just try a standard delete on the directory (the SHFile function WILL
	// delete a directory even if not empty no matter what flags you give it...)
	if (bRecurse == false)
	{
		if (!RemoveDirectory(szSource))
			return false;
		else
			return true;
	}

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
	TCHAR			szSource[_MAX_PATH+1];
	TCHAR			szDest[_MAX_PATH+1];
	TCHAR			szExpandedDest[MAX_PATH+1];
	TCHAR			szTempPath[_MAX_PATH+1];
	TCHAR			szDrive[_MAX_PATH+1];
	TCHAR			szDir[_MAX_PATH+1];
	TCHAR			szFile[_MAX_PATH+1];
	TCHAR			szExt[_MAX_PATH+1];

	// Get local version of our source/dest with full path names, strip trailing \s
	Util_GetFullPathName(szInputSource, szSource);
	Util_GetFullPathName(szInputDest, szDest);

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
		return 0; // Indicate no failures.
	}
	aLastError = 0; // Set default. Overridden only when a failure occurs.

	// Otherwise, loop through all the matching files.
	// Split source into file and extension (we need this info in the loop below to reconstruct the path)
	_tsplitpath(szSource, szDrive, szDir, szFile, szExt);
	// Note we now rely on the SOURCE being the contents of szDrive, szDir, szFile, etc.
	size_t szTempPath_length = sntprintf(szTempPath, _countof(szTempPath), _T("%s%s"), szDrive, szDir);
	LPTSTR append_pos = szTempPath + szTempPath_length;
	size_t space_remaining = _countof(szTempPath) - szTempPath_length - 1;

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
		_tcscpy(append_pos, findData.cFileName); // Indirectly populate szTempPath. Above has ensured this won't overflow.

		// Expand the destination based on this found file
		Util_ExpandFilenameWildcard(findData.cFileName, szDest, szExpandedDest);

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
			if (!MoveFile(szTempPath, szExpandedDest))
			{
				// If overwrite mode was not specified by the caller, or it was but the existing
				// destination file cannot be deleted (perhaps because it is a folder rather than
				// a file), or it can be deleted but the source cannot be moved, indicate a failure.
				// But by design, continue the operation.  The following relies heavily on
				// short-circuit boolean evaluation order:
				if (   !(bOverwrite && DeleteFile(szExpandedDest) && MoveFile(szTempPath, szExpandedDest))   )
				{
					aLastError = GetLastError();
					++failure_count; // At this stage, any of the above 3 being false is cause for failure.
				}
				//else everything succeeded, so nothing extra needs to be done.  In either case,
				// continue on to the next file.
			}
		}
		else // The mode is "Copy" vs. "Move"
			if (!CopyFile(szTempPath, szExpandedDest, !bOverwrite)) // Force it to fail if bOverwrite==false.
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
	// copy one.two.three  *.*.txt   = one.two   .three  .txt
	// copy one.two.three  *.*.*.txt = one.two   .three  ..txt
	// copy one.two		   test      = test

	TCHAR	szFileTemp[_MAX_PATH+1];
	TCHAR	szExtTemp[_MAX_PATH+1];

	TCHAR	szSrcFile[_MAX_PATH+1];
	TCHAR	szSrcExt[_MAX_PATH+1];

	TCHAR	szDestDrive[_MAX_PATH+1];
	TCHAR	szDestDir[_MAX_PATH+1];
	TCHAR	szDestFile[_MAX_PATH+1];
	TCHAR	szDestExt[_MAX_PATH+1];

	// If the destination doesn't include a wildcard, send it back verbatim
	if (_tcschr(szDest, '*') == NULL)
	{
		_tcscpy(szExpandedDest, szDest);
		return;
	}

	// Split source and dest into file and extension
	_tsplitpath( szSource, szDestDrive, szDestDir, szSrcFile, szSrcExt );
	_tsplitpath( szDest, szDestDrive, szDestDir, szDestFile, szDestExt );

	// Source and Dest ext will either be ".nnnn" or "" or ".*", remove the period
	if (szSrcExt[0] == '.')
		_tcscpy(szSrcExt, &szSrcExt[1]);
	if (szDestExt[0] == '.')
		_tcscpy(szDestExt, &szDestExt[1]);

	// Start of the destination with the drive and dir
	_tcscpy(szExpandedDest, szDestDrive);
	_tcscat(szExpandedDest, szDestDir);

	// Replace first * in the destext with the srcext, remove any other *
	Util_ExpandFilenameWildcardPart(szSrcExt, szDestExt, szExtTemp);

	// Replace first * in the destfile with the srcfile, remove any other *
	Util_ExpandFilenameWildcardPart(szSrcFile, szDestFile, szFileTemp);

	// Concat the filename and extension if req
	if (szExtTemp[0] != '\0')
	{
		_tcscat(szFileTemp, _T("."));
		_tcscat(szFileTemp, szExtTemp);	
	}
	else
	{
		// Dest extension was blank SOURCE MIGHT NOT HAVE BEEN!
		if (szSrcExt[0] != '\0')
		{
			_tcscat(szFileTemp, _T("."));
			_tcscat(szFileTemp, szSrcExt);	
		}
	}

	// Now add the drive and directory bit back onto the dest
	_tcscat(szExpandedDest, szFileTemp);

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



bool Line::Util_CreateDir(LPCTSTR szDirName) // Recursive directory creation function.
{
	DWORD	dwTemp;
	LPTSTR	szTemp = NULL;
	LPTSTR	psz_Loc = NULL;
	size_t  length;

	dwTemp = GetFileAttributes(szDirName);

	if (dwTemp == 0xffffffff) 
	{	// error getting attribute - what was the error?
		switch (GetLastError())
		{
		case ERROR_PATH_NOT_FOUND:
			// Create path
			length = _tcslen(szDirName);
			if (length > MAX_PATH) // Sanity check to reduce chance of stack overflow (since this function recursively calls self).
				return false;
			szTemp = (LPTSTR)talloca(length+1); // Faster, and also avoids need to delete it afterward.
			_tcscpy(szTemp, szDirName);
			psz_Loc = _tcsrchr(szTemp, '\\');	/* find last \ */
			if (psz_Loc == NULL)				// not found
				return false;
			else 
			{
				*psz_Loc = '\0';				// remove \ and everything after
				if (!Util_CreateDir(szTemp))
					return false;
				return CreateDirectory(szDirName, NULL) ? true : false;
			}
			// All paths above "return".
		case ERROR_FILE_NOT_FOUND:
			// Create directory
			return CreateDirectory(szDirName, NULL);
		// Otherwise, it's some unforeseen error, so fall through to the end, which reports failure.
		} // switch()
	}
	else // The specified name already exists as a file or directory.
		if (dwTemp & FILE_ATTRIBUTE_DIRECTORY) // Fixed for v1.0.36.01 (previously it used == vs &).
			return true;							// Directory exists, yay!
		//else it exists, but it's a file! Not allowed, so fall through and report failure.
			
	return false;

}



bool Line::Util_DoesFileExist(LPCTSTR szFilename)  // Returns true if file or directory exists.
{
	if ( _tcschr(szFilename,'*')||_tcschr(szFilename,'?') )
	{
		WIN32_FIND_DATA	wfd;
		HANDLE			hFile;

		hFile = FindFirstFile(szFilename, &wfd);

		if ( hFile == INVALID_HANDLE_VALUE )
			return false;

		FindClose(hFile);
		return true;
	}
    else
	{
		DWORD dwTemp;

		dwTemp = GetFileAttributes(szFilename);
		if ( dwTemp != 0xffffffff )
			return true;
		else
			return false;
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
	LPTSTR szFilePart;
	GetFullPathName(szIn, _MAX_PATH, szOut, &szFilePart);
	strip_trailing_backslash(szOut);
}



bool Line::Util_IsDifferentVolumes(LPCTSTR szPath1, LPCTSTR szPath2)
// Checks two paths to see if they are on the same volume.
{
	TCHAR			szP1Drive[_MAX_DRIVE+1];
	TCHAR			szP2Drive[_MAX_DRIVE+1];

	TCHAR			szDir[_MAX_DIR+1];
	TCHAR			szFile[_MAX_FNAME+1];
	TCHAR			szExt[_MAX_EXT+1];
	
	TCHAR			szP1[_MAX_PATH+1];	
	TCHAR			szP2[_MAX_PATH+1];

	// Get full pathnames
	Util_GetFullPathName(szPath1, szP1);
	Util_GetFullPathName(szPath2, szP2);

	// Split the target into bits
	_tsplitpath( szP1, szP1Drive, szDir, szFile, szExt );
	_tsplitpath( szP2, szP2Drive, szDir, szFile, szExt );

	if (szP1Drive[0] == '\0' || szP2Drive[0] == '\0')
		// One or both paths is a UNC - assume different volumes
		return true;
	else
		return _tcsicmp(szP1Drive, szP2Drive);
}



bool Util_Shutdown(int nFlag)
// Shutdown or logoff the system.
// Returns false if the function could not get the rights to shutdown.
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

	// If we are running NT/2k/XP, make sure we have rights to shutdown
	if (g_os.IsWinNT()) // NT/2k/XP/2003 and family
	{
		// Get a token for this process.
 		if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken)) 
			return false;						// Don't have the rights
 
		// Get the LUID for the shutdown privilege.
 		LookupPrivilegeValue(NULL, SE_SHUTDOWN_NAME, &tkp.Privileges[0].Luid); 
 
		tkp.PrivilegeCount = 1;  /* one privilege to set */
		tkp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED; 
 
		// Get the shutdown privilege for this process.
 		AdjustTokenPrivileges(hToken, FALSE, &tkp, 0, (PTOKEN_PRIVILEGES)NULL, 0); 
 
		// Cannot test the return value of AdjustTokenPrivileges.
 		if (GetLastError() != ERROR_SUCCESS) 
			return false;						// Don't have the rights
	}

	// if we are forcing the issue, AND this is 95/98 terminate all windows first
	if ( g_os.IsWin9x() && (nFlag & EWX_FORCE) ) 
	{
		nFlag ^= EWX_FORCE;	// remove this flag - not valid in 95
		EnumWindows((WNDENUMPROC) Util_ShutdownHandler, 0);
	}

	// ExitWindows
	if (ExitWindowsEx(nFlag, 0))
		return true;
	else
		return false;

}



BOOL Util_ShutdownHandler(HWND hwnd, DWORD lParam)
{
	// if the window is me, don't terminate!
	if (hwnd != g_hWnd && hwnd != g_hWndSplash)
		Util_WinKill(hwnd);

	// Continue the enumeration.
	return TRUE;

}



void Util_WinKill(HWND hWnd)
{
	DWORD_PTR dwResult;
	// Use WM_CLOSE vs. SC_CLOSE in this case, since the target window is slightly more likely to
	// respond to that:
	if (!SendMessageTimeout(hWnd, WM_CLOSE, 0, 0, SMTO_ABORTIFHUNG, 500, &dwResult)) // Wait up to 500ms.
	{
		// Use more force - Mwuahaha
		DWORD pid = GetWindowThreadProcessId(hWnd, NULL);
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

DWORD ProcessExist9x2000(LPTSTR aProcess)
{
	// We must dynamically load the function or program will probably not launch at all on NT4.
	typedef BOOL (WINAPI *PROCESSWALK)(HANDLE hSnapshot, LPPROCESSENTRY32 lppe);
	typedef HANDLE (WINAPI *CREATESNAPSHOT)(DWORD dwFlags, DWORD th32ProcessID);

	static CREATESNAPSHOT lpfnCreateToolhelp32Snapshot = (CREATESNAPSHOT)GetProcAddress(GetModuleHandle(_T("kernel32")), "CreateToolhelp32Snapshot");
    static PROCESSWALK lpfnProcess32First = (PROCESSWALK)GetProcAddress(GetModuleHandle(_T("kernel32")), "Process32First" PROCESS_API_SUFFIX);
    static PROCESSWALK lpfnProcess32Next = (PROCESSWALK)GetProcAddress(GetModuleHandle(_T("kernel32")), "Process32Next" PROCESS_API_SUFFIX);

	if (!lpfnCreateToolhelp32Snapshot || !lpfnProcess32First || !lpfnProcess32Next)
		return 0;

	PROCESSENTRY32 proc;
    proc.dwSize = sizeof(proc);
	HANDLE snapshot = lpfnCreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
	lpfnProcess32First(snapshot, &proc);

	// Determine the PID if aProcess is a pure, non-negative integer (any negative number
	// is more likely to be the name of a process [with a leading dash], rather than the PID).
	DWORD specified_pid = IsPureNumeric(aProcess) ? ATOU(aProcess) : 0;
	TCHAR szDrive[_MAX_PATH+1], szDir[_MAX_PATH+1], szFile[_MAX_PATH+1], szExt[_MAX_PATH+1];

	while (lpfnProcess32Next(snapshot, &proc))
	{
		if (specified_pid && specified_pid == proc.th32ProcessID)
		{
			CloseHandle(snapshot);
			return specified_pid;
		}
		// Otherwise, check for matching name even if aProcess is purely numeric (i.e. a number might
		// also be a valid name?):
		// It seems that proc.szExeFile never contains a path, just the executable name.
		// But in case it ever does, ensure consistency by removing the path:
		_tsplitpath(proc.szExeFile, szDrive, szDir, szFile, szExt);
		_tcscat(szFile, szExt);
		if (!_tcsicmp(szFile, aProcess)) // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales; 3) performance.
		{
			CloseHandle(snapshot);
			return proc.th32ProcessID;
		}
	}
	CloseHandle(snapshot);
	return 0;  // Not found.
}



#ifdef CONFIG_WINNT4
DWORD ProcessExistNT4(LPTSTR aProcess, LPTSTR aProcessName)
{
	if (aProcessName) // Init this output variable in case of early return.
		*aProcessName = '\0';
	//BOOL EnumProcesses(
	//  DWORD *lpidProcess,  // array of process identifiers
	//  DWORD cb,            // size of array
	//  DWORD *cbNeeded      // number of bytes returned
	//);
	typedef BOOL (WINAPI *MyEnumProcesses)(DWORD*, DWORD, DWORD*);

	//BOOL EnumProcessModules(
	//  HANDLE hProcess,      // handle to process
	//  HMODULE *lphModule,   // array of module handles
	//  DWORD cb,             // size of array
	//  LPDWORD lpcbNeeded    // number of bytes required
	//);
	typedef BOOL (WINAPI *MyEnumProcessModules)(HANDLE, HMODULE*, DWORD, LPDWORD);

	//DWORD GetModuleBaseName(
	//  HANDLE hProcess,    // handle to process
	//  HMODULE hModule,    // handle to module
	//  LPTSTR lpBaseName,  // base name buffer
	//  DWORD nSize         // maximum characters to retrieve
	//);
	typedef DWORD (WINAPI *MyGetModuleBaseName)(HANDLE, HMODULE, LPTSTR, DWORD);

	// We must dynamically load the function or program will probably not launch at all on Win95.
    // Get a handle to the DLL module that contains EnumProcesses
	HINSTANCE hinstLib = LoadLibrary(_T("psapi"));
	if (!hinstLib)
		return 0;

	// Not static in this case, since address can change with each new load of the library:
  	MyEnumProcesses lpfnEnumProcesses = (MyEnumProcesses)GetProcAddress(hinstLib, "EnumProcesses");
	MyEnumProcessModules lpfnEnumProcessModules = (MyEnumProcessModules)GetProcAddress(hinstLib, "EnumProcessModules");
	MyGetModuleBaseName lpfnGetModuleBaseName = (MyGetModuleBaseName)GetProcAddress(hinstLib, "GetModuleBaseName" WINAPI_SUFFIX);

	DWORD idProcessArray[512];		// 512 processes max
	DWORD cbNeeded;					// Bytes returned
	if (!lpfnEnumProcesses || !lpfnEnumProcessModules || !lpfnGetModuleBaseName
		|| !lpfnEnumProcesses(idProcessArray, sizeof(idProcessArray), &cbNeeded))
	{
		FreeLibrary(hinstLib);
		return 0;
	}

	// Get the count of PIDs in the array
	DWORD cProcesses = cbNeeded / sizeof(DWORD);
	// Determine the PID if aProcess is a pure, non-negative integer (any negative number
	// is more likely to be the name of a process [with a leading dash], rather than the PID).
	DWORD specified_pid = IsPureNumeric(aProcess) ? ATOU(aProcess) : 0;
	TCHAR szDrive[_MAX_PATH+1], szDir[_MAX_PATH+1], szFile[_MAX_PATH+1], szExt[_MAX_PATH+1];
	TCHAR szProcessName[_MAX_PATH+1];
	HMODULE hMod;
	HANDLE hProcess;

	for (UINT i = 0; i < cProcesses; ++i)
	{
		if (specified_pid && specified_pid == idProcessArray[i])
		{
			if (aProcessName) // Caller wanted process name also.
			{
				if (hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, idProcessArray[i])) // Assign
				{
					lpfnEnumProcessModules(hProcess, &hMod, sizeof(hMod), &cbNeeded);
					if (lpfnGetModuleBaseName(hProcess, hMod, szProcessName, _MAX_PATH))
					{
						// For consistency in results, use _splitpath() both here and below rather than
						// something that just checks for a rightmost backslash.
						_tsplitpath(szProcessName, szDrive, szDir, aProcessName, szExt);
						_tcscat(aProcessName, szExt);
					}
					CloseHandle(hProcess);
				}
			}
			FreeLibrary(hinstLib);
			return specified_pid;
		}
		// Otherwise, check for matching name even if aProcess is purely numeric (i.e. a number might
		// also be a valid name?):
		if (hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, idProcessArray[i])) // Assign
		{
			lpfnEnumProcessModules(hProcess, &hMod, sizeof(hMod), &cbNeeded);
			if (lpfnGetModuleBaseName(hProcess, hMod, szProcessName, _MAX_PATH))
			{
				_tsplitpath(szProcessName, szDrive, szDir, szFile, szExt);
				_tcscat(szFile, szExt);
				if (!_tcsicmp(szFile, aProcess)) // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales; 3) performance.
				{
					if (aProcessName) // Caller wanted process name also.
						_tcscpy(aProcessName, szProcessName);
					CloseHandle(hProcess);
					FreeLibrary(hinstLib);
					return idProcessArray[i];  // The PID.
				}
			}
			CloseHandle(hProcess);
		}
	}
	FreeLibrary(hinstLib);
	return 0;  // Not found.
}
#endif