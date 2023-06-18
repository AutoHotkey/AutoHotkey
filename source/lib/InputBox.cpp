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
#include "window.h"
#include "application.h"
#include "resources/resource.h"
#include "script_func_impl.h"



ResultType InputBoxParseOptions(LPCTSTR aOptions, InputBoxType &aInputBox)
{
	LPCTSTR next_option, option_end;
	for (next_option = aOptions; ; next_option = omit_leading_whitespace(option_end))
	{
		if (!*next_option)
			return OK; // All options parsed OK.

		// Find the end of this option item:
		for (option_end = next_option; *option_end && !IS_SPACE_OR_TAB(*option_end); ++option_end);

		// The legacy InputBox command used "Hide", but "Password" seems clearer
		// and better for consistency with the equivalent Edit control option:
		if (!_tcsnicmp(next_option, _T("Password"), 8))
		{
			if (next_option[8] && !IS_SPACE_OR_TAB(next_option[8]))
			{
				if (next_option[8] && (next_option[9] && !IS_SPACE_OR_TAB(next_option[9])))
					break; // More than one character - invalid.
				aInputBox.password_char = next_option[8];
			}
			else
				aInputBox.password_char = UorA(L'\x25CF', '*');
			continue;
		}

		// All of the remaining options are a single letter followed by a number.
		TCHAR option_char = ctoupper(*next_option);
		LPTSTR n_end;
		auto value = _tcstod(next_option + 1, &n_end); // Supports hex with VC++ 2015 and later.
		if (*n_end && !IS_SPACE_OR_TAB(*n_end) // Invalid option or non-numeric suffix.
			|| value < 0 && !(option_char == 'X' || option_char == 'Y'))
			break; 
		
		switch (ctoupper(*next_option))
		{
		case 'W': aInputBox.width = DPIScale((int)value); continue;
		case 'H': aInputBox.height = DPIScale((int)value); continue;
		case 'X': aInputBox.xpos = (int)value; continue;
		case 'Y': aInputBox.ypos = (int)value; continue;
		case 'T': aInputBox.timeout = (DWORD)(value * 1000); continue;
		}
		break;
	}
	return ValueError(ERR_INVALID_OPTION, next_option, FAIL_OR_OK);
}



bif_impl FResult InputBox(optl<StrArg> aPrompt, optl<StrArg> aTitle, optl<StrArg> aOptions, optl<StrArg> aDefault, IObject *&aRetVal)
{
	InputBoxType inputbox;
	inputbox.title = aTitle.has_value() ? aTitle.value() : g_script.DefaultDialogTitle();
	inputbox.text = aPrompt.value_or_empty();
	inputbox.default_string = aDefault.value_or_null();
	inputbox.return_string = nullptr;
	// Set defaults:
	inputbox.width = INPUTBOX_DEFAULT;
	inputbox.height = INPUTBOX_DEFAULT;
	inputbox.xpos = INPUTBOX_DEFAULT;
	inputbox.ypos = INPUTBOX_DEFAULT;
	inputbox.password_char = '\0';
	inputbox.timeout = 0;
	// Parse options and override defaults:
	if (aOptions.has_value() && !InputBoxParseOptions(aOptions.value(), inputbox))
		return FR_FAIL; // It already displayed the error.

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

	FResult fresult;
	if (reason && value)
	{
		ExprTokenType argt[] = { _T("Result"), reason, _T("Value"), value };
		ExprTokenType *args[_countof(argt)] = { argt, argt+1, argt+2, argt+3 };
		aRetVal = Object::Create(args, _countof(args));
		fresult = aRetVal ? OK : FR_E_OUTOFMEM;
	}
	else
		fresult = result == -1 ? FR_E_WIN32 : FR_E_FAILED;
	free(value);
	return fresult;
}



INT_PTR CALLBACK InputBoxProc(HWND hWndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
// MSDN:
// Typically, the dialog box procedure should return TRUE if it processed the message,
// and FALSE if it did not. If the dialog box procedure returns FALSE, the dialog
// manager performs the default dialog operation in response to the message.
{
	// See GuiWindowProc() for details about this first part:
	LRESULT msg_reply;
	if (g->CalledByIsDialogMessageOrDispatch && g->CalledByIsDialogMessageOrDispatchMsg == uMsg)
		g->CalledByIsDialogMessageOrDispatch = false;
	else if (g_MsgMonitor.Count() && MsgMonitor(hWndDlg, uMsg, wParam, lParam, NULL, msg_reply))
		return (BOOL)msg_reply; // MsgMonitor has returned "true", indicating that this message should be omitted from further processing.

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
		
		if (CURR_INPUTBOX.default_string)
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
