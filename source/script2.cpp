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
#include "qmath.h" // Used by Transform() [math.h incurs 2k larger code size just for ceil() & floor()]
#include "script.h"
#include "window.h" // for IF_USE_FOREGROUND_WINDOW
#include "application.h" // for MsgSleep()
#include "script_func_impl.h"
#include "abi.h"



///////////////////////
// GUI-related: Tray //
///////////////////////

ResultType TrayTipParseOptions(LPCTSTR aOptions, NOTIFYICONDATA &nic)
{
	if (!aOptions)
		return OK;
	DWORD flag;
	LPCTSTR next_option, option_end;
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
		else if (ParseInteger(option, flag))
		{
			nic.dwInfoFlags |= flag;
		}
		else
		{
			goto invalid_option;
		}
	}
invalid_option:
	return ValueError(ERR_INVALID_OPTION, next_option, FAIL_OR_OK);
}


bif_impl FResult TrayTip(optl<StrArg> aText, optl<StrArg> aTitle, optl<StrArg> aOptions)
{
	NOTIFYICONDATA nic = {0};
	nic.cbSize = sizeof(nic);
	nic.uID = AHK_NOTIFYICON;  // This must match our tray icon's uID or Shell_NotifyIcon() will return failure.
	nic.hWnd = g_hWnd;
	nic.uFlags = NIF_INFO;
	// nic.uTimeout is no longer used because it is valid only on Windows 2000 and Windows XP.
	if (!TrayTipParseOptions(aOptions.value_or_null(), nic))
		return FR_FAIL;
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
	if (aTitle.has_value())
		tcslcpy(nic.szInfoTitle, aTitle.value(), _countof(nic.szInfoTitle));
	else
		*nic.szInfoTitle = '\0'; // Empty title omits the title line entirely.
	if (aText.has_nonempty_value())
		tcslcpy(nic.szInfo, aText.value(), _countof(nic.szInfo));
	else if (aTitle.has_nonempty_value()) // Text was omitted or empty, but Title was specified.
		*nic.szInfo = ' ', nic.szInfo[1] = '\0'; // Pass a non-empty string to ensure the notification is shown.
	else // Both parameters were omitted or empty.
		*nic.szInfo = '\0'; // Empty text removes the notification (or tries to).
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
	if (g->CalledByIsDialogMessageOrDispatch && g->CalledByIsDialogMessageOrDispatchMsg == iMsg)
		g->CalledByIsDialogMessageOrDispatch = false; // Suppress this one message, not any other messages that could be sent due to recursion.
	else if (g_MsgMonitor.Count() && MsgMonitor(hWnd, iMsg, wParam, lParam, NULL, msg_reply))
		return msg_reply; // MsgMonitor has returned "true", indicating that this message should be omitted from further processing.

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
			g_script.mTrayMenu->Display();
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
		g_MenuIsVisible = true; // See comments in similar code in GuiWindowProc().
		break;
	case WM_EXITMENULOOP:
		g_MenuIsVisible = false; // See comments in similar code in GuiWindowProc().
		break;
	case WM_INITMENUPOPUP:
		InitMenuPopup((HMENU)wParam);
		break;
	case WM_UNINITMENUPOPUP:
		UninitMenuPopup((HMENU)wParam);
		break;
	case WM_ACTIVATEAPP:
		// Modeless menus correctly cancel when losing focus to another window within the same process,
		// but not when losing focus to another app, so we handle that here if we made the menu modeless.
		// Don't do it if the script made the menu modeless since in that case the script might want it
		// to remain open.  It appears that WM_ACTIVATEAPP is sent to all top-level windows, so there's
		// no need to handle this in GuiWindowProc().
		if (!wParam && g_MenuIsTempModeless)
			EndMenu();
		break;
	case WM_MENUSELECT:
		// The following is a workaround for left click failing to activate a menu item in a modeless
		// menu if the mouse was moved quickly from the main menu into a submenu (reproduced on Windows
		// 7 and 11).  Strangely, double-click still activates the item.  Moving the selection restores
		// normal behaviour, so the workaround just deselects the item and allows it to be reselected.
		// Care must be taken to avoid a loop, because deselecting the item causes the parent menu to
		// be reselected, which causes WM_MENUSELECT to be sent.
		if (g_MenuIsTempModeless)
		{
			static HMENU sLastSelectedMenu = NULL; // Limits unnecessarily application of the workaround.
			static bool sRecursiveCall = false; // Prevents looping due to MN_SELECTITEM causing WM_MENUSELECT.
			if (sLastSelectedMenu != (HMENU)lParam && !sRecursiveCall)
			{
				sLastSelectedMenu = (HMENU)lParam; // Before SendMessage() below.
				constexpr auto MF_WANTED = MF_MOUSESELECT | MF_HILITE; // Item selected by mouse.
				constexpr auto MF_UNWANTED = MF_GRAYED | MF_DISABLED | MF_POPUP; // Items not needing the workaround.
				HWND fore_win;
				if (   (HMENU)lParam != g_MenuIsTempModeless // Only submenus need the workaround.
					&& (HIWORD(wParam) & (MF_WANTED | MF_UNWANTED)) == MF_WANTED
					&& (fore_win = GetForegroundWindow())
					&& SendMessage(fore_win, MN_GETHMENU, 0, 0) == lParam   )
				{
					constexpr auto Mn_SELECTITEM = 0x01E5; // Undocumented message?
					sRecursiveCall = true;
					SendMessage(fore_win, Mn_SELECTITEM, -1, 0);
					sRecursiveCall = false;
				}
			}
		}
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
			// This message is sent by the system in two known cases:
			//  1) Explorer.exe has just started and the taskbar has been newly created.
			//     In this case, the taskbar icon doesn't exist yet, so NIM_MODIFY would fail.
			//  2) The screen DPI has just changed.  Our icon already exists, but has probably
			//     been resized by the system.  If we don't refresh it, it becomes blurry.
			g_script.UpdateTrayIcon(true);
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
		ToggleSuspendState();
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



bif_impl void ListLines(optl<int> aMode, int &aRetVal)
{
	aRetVal = g->ListLinesIsEnabled;
	if (!aMode.has_value())
	{
		ShowMainWindow(MAIN_MODE_LINES, false); // Pass "unrestricted" when the command is explicitly used in the script.
		return;
	}
	if (g->ListLinesIsEnabled)
	{
		// Since ExecUntil() just logged this ListLines On/Off in the line history, remove it to avoid
		// cluttering the line history with distracting lines that the user probably wouldn't want to see.
		// Might be especially useful in cases where a timer fires frequently (even if such a timer used
		// "ListLines 0" as its top line, that line itself would appear very frequently in the line history).
		if (Line::sLogNext > 0)
			--Line::sLogNext;
		else
			Line::sLogNext = LINE_LOG_SIZE - 1;
		Line::sLog[Line::sLogNext] = NULL; // Without this, one of the lines in the history would be invalid due to the circular nature of the line history array, which would also cause the line history to show the wrong chronological order in some cases.
	}
	// aMode is defined as Int32 rather than Bool32 to avoid confusion
	// (ListLines("") acting as ListLines(false) instead of ListLines())
	// and reserve string values for potential future use.
	g->ListLinesIsEnabled = *aMode != 0;
}



bif_impl void ListVars()
{
	ShowMainWindow(MAIN_MODE_VARS, false); // Pass "unrestricted" when the command is explicitly used in the script.
}



bif_impl void ListHotkeys()
{
	ShowMainWindow(MAIN_MODE_HOTKEYS, false); // Pass "unrestricted" when the command is explicitly used in the script.
}



bif_impl FResult KeyHistory(optl<int> aMaxEvents)
{
	if (!aMaxEvents.has_value())
	{
		ShowMainWindow(MAIN_MODE_KEYHISTORY, false); // Pass "unrestricted" when the command is explicitly used in the script.
		return OK;
	}
	int value = *aMaxEvents;
	if (value < 0 || value > 500)
		return FR_E_ARG(0);
	// GetHookStatus() only has a limited size buffer in which to transcribe the keystrokes.
	// 500 events is about what you would expect to fit in a 32 KB buffer (in the unlikely
	// event that the transcribed events create too much text, the text will be truncated,
	// so it's not dangerous anyway).
	if (g_KeybdHook || g_MouseHook)
		PostThreadMessage(g_HookThreadID, AHK_HOOK_SET_KEYHISTORY, value, 0);
	else
		SetKeyHistoryMax(value);
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



/////////////////////////
// GUI-related: MsgBox //
/////////////////////////


ResultType MsgBoxParseOptions(LPCTSTR aOptions, int &aType, double &aTimeout, HWND &aOwner)
{
	aType = 0;
	aTimeout = 0;

	if (!aOptions)
		return OK;

	//int button_option = 0;
	//int icon_option = 0;

	int option_int;
	LPCTSTR next_option, option_end;
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
		else if (!_tcsnicmp(option, _T("Default"), 7) && ParseInteger(option + 7, option_int))
		{
			if (option_int < 1 || option_int > 0xF) // Currently MsgBox can only have 4 buttons, but MB_DEFMASK may allow for up to this many in future.
				goto invalid_option;
			aType = (aType & ~MB_DEFMASK) | ((option_int - 1) << 8); // 1=0, 2=0x100, 3=0x200, 4=0x300
		}
		else if (toupper(*option) == 'T' && IsNumeric(option + 1, FALSE, FALSE, TRUE))
		{
			aTimeout = ATOF(option + 1);
		}
		else if (!_tcsnicmp(option, _T("Owner"), 5) && IsNumeric(option + 5, TRUE, TRUE, FALSE))
		{
			aOwner = (HWND)ATOI64(option + 5); // This should be consistent with the Gui +Owner option.
		}
		else if (ParseInteger(option, option_int))
		{
			// Clear any conflicting options which were previously set.
			if (option_int & MB_TYPEMASK) aType &= ~MB_TYPEMASK;
			if (option_int & MB_ICONMASK) aType &= ~MB_ICONMASK;
			if (option_int & MB_DEFMASK)  aType &= ~MB_DEFMASK;
			if (option_int & MB_MODEMASK) aType &= ~MB_MODEMASK;
			// All remaining options are bit flags (or not conflicting).
			aType |= option_int;
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


bif_impl FResult MsgBox(optl<StrArg> aText, optl<StrArg> aTitle, optl<StrArg> aOptions, StrRet &aRetVal)
{
	int result;
	HWND dialog_owner = THREAD_DIALOG_OWNER; // Resolve macro only once to reduce code size.
	// dialog_owner is passed via parameter to avoid internally-displayed MsgBoxes from being
	// affected by script-thread's owner setting.
	int type;
	double timeout;
	if (!MsgBoxParseOptions(aOptions.value_or_null(), type, timeout, dialog_owner))
		return FR_FAIL;
	SetLastError(0); // To differentiate MAX_MSGBOXES from invalid parameters.
	result = MsgBox(aText.value_or_null(), type, aTitle.value_or_null(), timeout, dialog_owner);
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

	// Return a button name such as "OK", "Yes" or "No" if applicable.
	if (LPTSTR result_string = MsgBoxResultString(result))
	{
		aRetVal.SetStatic(result_string);
		return OK;
	}
	if (result == AHK_TOO_MANY_MSGBOXES)
		// Raise an error, since the MsgBox() might be intended to request approval or input from
		// the user, and continuing without doing so could be risky:
		return FError(_T("The maximum number of MsgBoxes has been reached."));
	auto error = GetLastError();
	return error == ERROR_INVALID_MSGBOX_STYLE ? FR_E_ARG(2) : FR_E_WIN32(error);
}



//////////////////////////
// GUI-related: ToolTip //
//////////////////////////

bif_impl FResult ToolTip(optl<StrArg> aText, optl<int> aX, optl<int> aY, optl<int> aIndex, UINT &aRetVal)
{
	int window_index = aIndex.value_or(1) - 1;
	if (window_index < 0 || window_index >= MAX_TOOLTIPS)
		return FR_E_ARG(3);

	HWND tip_hwnd = g_hWndToolTip[window_index];

	// Destroy windows except the first (for performance) so that resources/mem are conserved.
	// The first window will be hidden by the TTM_UPDATETIPTEXT message if aText is blank.
	// UPDATE: For simplicity, destroy even the first in this way, because otherwise a script
	// that turns off a non-existent first tooltip window then later turns it on will cause
	// the window to appear in an incorrect position.  Example:
	// ToolTip
	// ToolTip text, 388, 24
	// Sleep 1000
	// ToolTip text, 388, 24
	auto tip_text = aText.value_or_empty();
	if (!*tip_text)
	{
		if (tip_hwnd && IsWindow(tip_hwnd))
			DestroyWindow(tip_hwnd);
		g_hWndToolTip[window_index] = NULL;
		aRetVal = 0;
		return OK;
	}
	
	bool one_or_both_coords_unspecified = !aX.has_value() || !aY.has_value();
		 
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

	POINT origin = { 0 };
	if (aX.has_value() || aY.has_value()) // Need the offsets.
		CoordToScreen(origin, COORD_MODE_TOOLTIP);

	// This will also convert from relative to screen coordinates if appropriate:
	if (aX.has_value())
		pt.x = aX.value() + origin.x;
	if (aY.has_value())
		pt.y = aY.value() + origin.y;
	HMONITOR hmon = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
	MONITORINFO mi;
	mi.cbSize = sizeof(mi);
	GetMonitorInfo(hmon, &mi);
	// v1.1.34: Use work area to avoid trying to overlap the taskbar on newer OSes, which otherwise
	// would cause the tooltip to appear at the top of the screen instead of the position we specify.
	// This was observed on Windows 10 and 11, and confirmed to not apply to Windows 7 or XP.
	RECT dtw = g_os.IsWin8orLater() ? mi.rcWork : mi.rcMonitor;

	TOOLINFO ti = { 0 };
	ti.cbSize = sizeof(ti);
	ti.uFlags = TTF_TRACK | TTF_ABSOLUTE;
	ti.lpszText = const_cast<LPTSTR>(tip_text);
	// ti.hwnd is the window to which notification messages are sent.  Set this to allow customization.
	ti.hwnd = g_hWnd;
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
	bool newly_created = !tip_hwnd || !IsWindow(tip_hwnd);
	if (newly_created)
	{
		// This this window has no owner, it won't be automatically destroyed when its owner is.
		// Thus, it will be explicitly by the program's exit function.
		tip_hwnd = g_hWndToolTip[window_index] = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, NULL, TTS_NOPREFIX | TTS_ALWAYSTIP
			, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, NULL, NULL);
		if (!tip_hwnd)
			return FR_E_WIN32;
		SendMessage(tip_hwnd, TTM_ADDTOOL, 0, (LPARAM)(LPTOOLINFO)&ti);
	}

	// v1.1.34: Fixed to use the appropriate monitor, in case it's sized differently to the primary.
	// Also fixed to account for incorrect DPI scaling done by the tooltip control; i.e. a value of
	// n ends up allowing tooltips n*g_ScreenDPI/96 pixels wide.  TTM_SETMAXTIPWIDTH seems to want
	// the max text width, not the max window width, so adjust for that.  Do this every time since
	// the tooltip might be moving between screens of different sizes.
	RECT text_rect = dtw;
	SendMessage(tip_hwnd, TTM_ADJUSTRECT, FALSE, (LPARAM)&text_rect);
	SendMessage(tip_hwnd, TTM_SETMAXTIPWIDTH, 0, (LPARAM)((text_rect.right - text_rect.left) * 96 / g_ScreenDPI));

	if (newly_created)
	{
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
	// v1.1.34: Avoid TTM_UPDATETIPTEXT if the text hasn't changed, to reduce flicker.  The behaviour described
	// above could not be replicated, EVEN ON WINDOWS XP.  Whether it was ever observed on other OSes is unknown.
	if (!newly_created && !ToolTipTextEquals(tip_hwnd, tip_text))
		SendMessage(tip_hwnd, TTM_UPDATETIPTEXT, 0, (LPARAM)&ti);

	RECT ttw = { 0 };
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

	// These messages seem to cause a complete update of the tooltip, which is slow and causes flickering.
	// It is tempting to use SetWindowPos() instead to speed things up, but if TTM_TRACKPOSITION isn't
	// sent each time, the next TTM_UPDATETIPTEXT message will move it back to whatever position was set
	// with TTM_TRACKPOSITION last.
	SendMessage(tip_hwnd, TTM_TRACKPOSITION, 0, (LPARAM)MAKELONG(pt.x, pt.y));
	// And do a TTM_TRACKACTIVATE even if the tooltip window already existed upon entry to this function,
	// so that in case it was hidden or dismissed while its HWND still exists, it will be shown again:
	SendMessage(tip_hwnd, TTM_TRACKACTIVATE, TRUE, (LPARAM)&ti);
	aRetVal = (UINT)(UINT_PTR)tip_hwnd;
	return OK;
}



///////////////////
// Misc internal //
///////////////////

VOID CALLBACK DerefTimeout(HWND hWnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	Line::FreeDerefBufIfLarge(); // It will also kill the timer, if appropriate.
}



/////////////////
// MouseGetPos //
/////////////////

bif_impl FResult MouseGetPos(int *aX, int *aY, ResultToken *aParent, ResultToken *aChild, optl<int> aFlag)
{
	POINT point;
	GetCursorPos(&point);  // Realistically, can't fail?

	POINT origin = {0};
	CoordToScreen(origin, COORD_MODE_MOUSE);

	if (aX) *aX = point.x - origin.x;
	if (aY) *aY = point.y - origin.y;

	if (!aParent && !aChild)
		return OK;

	int aOptions = aFlag.value_or(0);

	// This is the child window.  Despite what MSDN says, WindowFromPoint() appears to fetch
	// a non-NULL value even when the mouse is hovering over a disabled control (at least on XP).
	HWND child_under_cursor = WindowFromPoint(point);
	if (!child_under_cursor)
		return OK;

	HWND parent_under_cursor = GetNonChildParent(child_under_cursor);  // Find the first ancestor that isn't a child.
	if (aParent)
	{
		// Testing reveals that an invisible parent window never obscures another window beneath it as seen by
		// WindowFromPoint().  In other words, the below never happens, so there's no point in having it as a
		// documented feature:
		//if (!g->DetectHiddenWindows && !IsWindowVisible(parent_under_cursor))
		//	return output_var_parent->Assign();
		aParent->SetValue((UINT_PTR)parent_under_cursor);
	}

	if (!aChild)
		return OK;

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
		return OK;

	if (aOptions & 0x02) // v1.0.43.06: Bitwise flag that means "return control's HWND vs. ClassNN".
	{
		aChild->SetValue((UINT_PTR)child_under_cursor);
		return OK;
	}

	TCHAR class_nn[WINDOW_CLASS_NN_SIZE];
	auto fr = ControlGetClassNN(parent_under_cursor, child_under_cursor, class_nn, _countof(class_nn));
	if (fr != OK)
		return fr;
	if (!TokenSetResult(*aChild, class_nn))
		return aChild->Exited() ? FR_FAIL : FR_ABORTED;
	return OK;
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



///////////////////////
// Working Directory //
///////////////////////


FResult SetWorkingDir(LPCTSTR aNewDir)
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
		return FR_E_WIN32;
	// Otherwise, the change to the working directory succeeded.

	// Other than during program startup, this should be the only place where the official
	// working dir can change.  The exception is FileSelect(), which changes the working
	// dir as the user navigates from folder to folder.  However, the whole purpose of
	// maintaining g_WorkingDir is to workaround that very issue.
	if (g_script.mIsReadyToExecute) // Callers want this done only during script runtime.
		UpdateWorkingDir(aNewDir);
	return OK;
}



void UpdateWorkingDir(LPCTSTR aNewDir)
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



/////////////////////////////
// GUI-related: FileSelect //
/////////////////////////////

bif_impl FResult FileSelect(optl<StrArg> aOptions, optl<StrArg> aWorkingDir, optl<StrArg> aGreeting
	, optl<StrArg> aFilter, ResultToken &aResultToken)
// Since other script threads can interrupt this command while it's running, it's important that
// this command not refer to sArgDeref[] and sArgVar[] anytime after an interruption becomes possible.
// This is because an interrupting thread usually changes the values to something inappropriate for this thread.
{
	if (g_nFileDialogs >= MAX_FILEDIALOGS)
	{
		// Have a maximum to help prevent runaway hotkeys due to key-repeat feature, etc.
		return FError(_T("The maximum number of File Dialogs has been reached."));
	}
	
	LPCTSTR default_file_name = _T("");
	LPCTSTR initial_dir = nullptr;
	if (aWorkingDir.has_nonempty_value())
	{
		LPCTSTR dir_and_name = aWorkingDir.value();
		size_t initial_dir_length = 0;
		DWORD attr = GetFileAttributes(dir_and_name);
		if ((attr != 0xFFFFFFFF) && (attr & FILE_ATTRIBUTE_DIRECTORY))
		{
			// Use the entire dir_and_name as the initial directory.
			initial_dir = dir_and_name;
			initial_dir_length = _tcslen(dir_and_name);
		}
		else
		{
			// Above condition indicates it's either an existing file that's not a folder, or a nonexistent
			// folder/filename.  In either case, it seems best to assume it's a file because the user may want
			// to provide a default SAVE filename, and it would be normal for such a file not to already exist.
			LPCTSTR last_backslash = _tcsrchr(dir_and_name, '\\');
			if (!last_backslash)
			{
				// Use the entire dir_and_name as the default filename.
				default_file_name = dir_and_name;
			}
			else
			{
				default_file_name = last_backslash + 1;
				initial_dir = dir_and_name;
				initial_dir_length = last_backslash - dir_and_name;
			}
		}
		if (initial_dir && initial_dir[initial_dir_length]) // Null termination required.
		{
			if (initial_dir_length < T_MAX_PATH) // Avoid stack overflow in case of bad data; anything longer wouldn't work anyway.
			{
				LPTSTR buf = (LPTSTR)_alloca(sizeof(TCHAR) * (initial_dir_length + 1));
				initial_dir = tmemcpy(buf, initial_dir, initial_dir_length);
				buf[initial_dir_length] = '\0';
			}
			else
				initial_dir = NULL;
		}
	}

	TCHAR pattern[1024];
	*pattern = '\0'; // Set default.
	if (aFilter.has_nonempty_value())
	{
		auto pattern_start = _tcschr(aFilter.value(), '(');
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
			tcslcpy(pattern, aFilter.value(), _countof(pattern));
	}
	UINT filter_count = 0;
	COMDLG_FILTERSPEC filters[2];
	if (*pattern) // aFilter was not omitted or blank.
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
		filters[0].pszName = aFilter.value();
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
	auto options_str = aOptions.value_or_empty();
	switch (ctoupper(*options_str))
	{
	case 'D':
		++options_str;
		flags |= FOS_PICKFOLDERS;
		if (*pattern)
			return FR_E_ARG(3);
		filter_count = 0;
		break;
	case 'M':  // Multi-select.
		++options_str;
		flags |= FOS_ALLOWMULTISELECT;
		break;
	case 'S': // Have a "Save" button rather than an "Open" button.
		++options_str;
		always_use_save_dialog = true;
		break;
	}

	TCHAR greeting[1024];
	if (aGreeting.has_nonempty_value())
		tcslcpy(greeting, aGreeting.value(), _countof(greeting));
	else
		// Use a more specific title so that the dialogs of different scripts can be distinguished
		// from one another, which may help script automation in rare cases:
		sntprintf(greeting, _countof(greeting), _T("Select %s - %s")
			, (flags & FOS_PICKFOLDERS) ? _T("Folder") : _T("File"), g_script.DefaultDialogTitle());

	int options = ATOI(options_str);
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
		return hr;

	pfd->SetOptions(flags);
	pfd->SetTitle(greeting);
	if (filter_count)
		pfd->SetFileTypes(filter_count, filters);
	pfd->SetFileName(default_file_name);

	if (initial_dir)
	{
		IShellItem *psi;
		if (SUCCEEDED(SHCreateItemFromParsingName(initial_dir, nullptr, IID_PPV_ARGS(&psi))))
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
		aResultToken.Return(files);
		return OK;
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
	pfd->Release();
	return result == HRESULT_FROM_WIN32(ERROR_CANCELLED) ? OK : result;
}



////////////////////////
// Keyboard Functions //
////////////////////////

FResult SetToggleState(vk_type aVK, ToggleValueType &ForceLock, optl<StrArg> aToggleText)
{
	ToggleValueType toggle = Line::ConvertOnOffAlways(aToggleText.value_or_null(), NEUTRAL);
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
	default:
		return FR_E_ARG(0);
	}
	return OK;
}



////////////////////////////////
// Misc lower level functions //
////////////////////////////////


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


////////////////////
// Misc Functions //
////////////////////


bif_impl void Persistent(optl<BOOL> aNewValue, BOOL &aOldValue)
{
	// Returning the old value might have some use, but if the caller doesn't want its value to change,
	// something awkward like "Persistent(isPersistent := Persistent())" is needed.  Rather than just
	// returning the current status, Persistent() makes the script persistent because that's likely to
	// be its most common use by far, and it's what users familiar with the old #Persistent may expect.
	aOldValue = g_persistent;
	g_persistent = aNewValue.value_or(TRUE);
}



static void InstallHook(optl<BOOL> aInstalling, optl<BOOL> aUseForce, HookType which_hook)
{
	bool installing = aInstalling.value_or(true);
	bool use_force = aUseForce.value_or(false);
	// When the second parameter is true, unconditionally remove the hook.  If the first parameter is
	// also true, the hook will be reinstalled fresh.  Otherwise the hook will be left uninstalled,
	// until something happens to reinstall it, such as Hotkey::ManifestAllHotkeysHotstringsHooks().
	if (use_force)
		AddRemoveHooks(GetActiveHooks() & ~which_hook);
	Hotkey::RequireHook(which_hook, installing);
	if (!use_force || installing)
		Hotkey::ManifestAllHotkeysHotstringsHooks();
}

bif_impl void InstallKeybdHook(optl<BOOL> aInstalling, optl<BOOL> aUseForce)
{
	InstallHook(aInstalling, aUseForce, HOOK_KEYBD);
}

bif_impl void InstallMouseHook(optl<BOOL> aInstalling, optl<BOOL> aUseForce)
{
	InstallHook(aInstalling, aUseForce, HOOK_MOUSE);
}



////////////////////
// Core Functions //
////////////////////


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



////////////////////
// Core Functions //
////////////////////


bif_impl BOOL IsLabel(StrArg aName)
{
	return g_script.FindLabel(aName) ? 1 : 0;
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
		if_condition = YYYYMMDDToSystemTime(aValueStr, st, true);
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
	_f_throw_param(0, _T("String"));
}



BIF_DECL(BIF_IsSet)
{
	Var *var = ParamIndexToOutputVar(0);
	if (!var)
		_f_throw_param(0, _T("variable reference"));
	_f_return_b(!var->IsUninitializedNormalVar());
}



////////////////////////
// Keyboard Functions //
////////////////////////


bif_impl FResult GetKeyState(StrArg key_name, optl<StrArg> aMode, ResultToken &aResultToken)
{
	JoyControls joy;
	int joystick_id;
	vk_type vk = TextToVK(key_name);
	if (!vk)
	{
		if (   !(joy = (JoyControls)ConvertJoy(key_name, &joystick_id))   )
		{
			// It is neither a key name nor a joystick button/axis.
			return FR_E_ARG(0);
		}
		ScriptGetJoyState(joy, joystick_id, aResultToken, _f_retval_buf);
		return OK;
	}
	// Since above didn't return: There is a virtual key (not a joystick control).
	KeyStateTypes key_state_type;
	switch (aMode.has_value() ? ctoupper(*aMode.value()) : 'L') // Second parameter.
	{
	case 'T': key_state_type = KEYSTATE_TOGGLE; break; // Whether a toggleable key such as CapsLock is currently turned on.
	case 'P': key_state_type = KEYSTATE_PHYSICAL; break; // Physical state of key.
	case 'L': key_state_type = KEYSTATE_LOGICAL; break;
	default: return FR_E_ARG(1);
	}
	aResultToken.SetValue(ScriptGetKeyState(vk, key_state_type)); // 1 for down and 0 for up.
	return OK;
}



bif_impl int GetKeyVK(StrArg aKeyName)
{
	vk_type vk;
	sc_type sc;
	TextToVKandSC(aKeyName, vk, sc);
	return vk ? vk : sc_to_vk(sc);
}


bif_impl int GetKeySC(StrArg aKeyName)
{
	vk_type vk;
	sc_type sc;
	TextToVKandSC(aKeyName, vk, sc);
	return sc ? sc : vk_to_sc(vk);
}


bif_impl void GetKeyName(StrArg aKeyName, StrRet &aRetVal)
{
	vk_type vk;
	sc_type sc;
	TextToVKandSC(aKeyName, vk, sc);
	aRetVal.SetTemp(GetKeyName(vk, sc, aRetVal.CallerBuf(), aRetVal.CallerBufSize, _T("")));
}



////////////////////
// Core Functions //
////////////////////


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



////////////////////
// Misc Functions //
////////////////////


bif_impl FResult BIF_Hotkey(StrArg aName, ExprTokenType *aAction, optl<StrArg> aOptions)
{
	ResultType result = OK;
	IObject *functor = nullptr;
	HookActionType hook_action = 0;
	if (aAction)
	{
		auto action_string = TokenToString(*aAction);
		if (  !(functor = TokenToObject(*aAction)) && *action_string
			&& !(hook_action = Hotkey::ConvertAltTab(action_string, true))  )
		{
			// Search for a match in the hotkey variants' "original callbacks".
			// I.e., find the function implicitly defined by "x::action".
			for (int i = 0; i < Hotkey::sHotkeyCount; ++i)
			{
				if (_tcscmp(Hotkey::shk[i]->mName, action_string))
					continue;
				
				for (HotkeyVariant* v = Hotkey::shk[i]->mFirstVariant; v; v = v->mNextVariant)
					if (v->mHotCriterion == g->HotCriterion)
					{
						functor = v->mOriginalCallback.ToObject();
						goto break_twice;
					}
			}
		break_twice:;
			if (!functor)
				return FR_E_ARG(1);
		}
		if (!functor)
			hook_action = Hotkey::ConvertAltTab(action_string, true);
	}
	return Hotkey::Dynamic(aName, aOptions.value_or_empty(), functor, hook_action);
}



bif_impl FResult HotIf(ExprTokenType *aCriterion)
{
	TCHAR buf[MAX_NUMBER_SIZE];
	if (!aCriterion)
	{
		g->HotCriterion = nullptr;
		return OK;
	}
	else if (auto obj = TokenToObject(*aCriterion))
		return Hotkey::IfExpr(obj);
	else
		return Hotkey::IfExpr(TokenToString(*aCriterion, buf));
}



bif_impl FResult HotIfWinActive(optl<StrArg> aWinTitle, optl<StrArg> aWinText)
{
	return SetHotkeyCriterion(HOT_IF_ACTIVE, aWinTitle.value_or_empty(), aWinText.value_or_empty());
}

bif_impl FResult HotIfWinNotActive(optl<StrArg> aWinTitle, optl<StrArg> aWinText)
{
	return SetHotkeyCriterion(HOT_IF_NOT_ACTIVE, aWinTitle.value_or_empty(), aWinText.value_or_empty());
}

bif_impl FResult HotIfWinExist(optl<StrArg> aWinTitle, optl<StrArg> aWinText)
{
	return SetHotkeyCriterion(HOT_IF_EXIST, aWinTitle.value_or_empty(), aWinText.value_or_empty());
}

bif_impl FResult HotIfWinNotExist(optl<StrArg> aWinTitle, optl<StrArg> aWinText)
{
	return SetHotkeyCriterion(HOT_IF_NOT_EXIST, aWinTitle.value_or_empty(), aWinText.value_or_empty());
}



bif_impl FResult SetTimer(optl<IObject*> aFunction, optl<__int64> aPeriod, optl<int> aPriority)
{
	IObject *callback;
	// Note that only one timer per callback is allowed because the callback is the
	// unique identifier that allows us to update or delete an existing timer.
	if (!aFunction.has_value())
	{
		if (g->CurrentTimer)
			// Default to the timer which launched the current thread.
			callback = g->CurrentTimer->mCallback.ToObject();
		else
			callback = NULL;
		if (!callback)
			// Either the thread was not launched by a timer or the timer has been deleted.
			return FR_E_ARG(0);
	}
	else
	{
		callback = aFunction.value();
		auto fr = ValidateFunctor(callback, 0);
		if (fr != OK)
			return fr;
	}
	__int64 period = DEFAULT_TIMER_PERIOD;
	int priority = 0;
	bool update_period = false, update_priority = false;
	if (aPeriod.has_value())
	{
		period = aPeriod.value();
		if (!period)
		{
			g_script.DeleteTimer(callback);
			return OK;
		}
		update_period = true;
	}
	if (aPriority.has_value())
	{
		priority = aPriority.value();
		update_priority = true;
	}
	g_script.UpdateOrCreateTimer(callback, update_period, period, update_priority, priority);
	return OK;
}



bif_impl void ScriptSleep(int aDelay)
{
	MsgSleep(aDelay);
}



bif_impl void Critical(optl<StrArg> aSetting, int &aRetVal)
{
	aRetVal = g->ThreadIsCritical ? g->PeekFrequency : 0;
	// v1.0.46: When the current thread is critical, have the script check messages less often to
	// reduce situations where an OnMessage or GUI message must be discarded due to "thread already
	// running".  Using 16 rather than the default of 5 solves reliability problems in a custom-menu-draw
	// script and probably many similar scripts -- even when the system is under load (though 16 might not
	// be enough during an extreme load depending on the exact preemption/timeslice dynamics involved).
	// DON'T GO TOO HIGH because this setting reduces response time for ALL messages, even those that
	// don't launch script threads (especially painting/drawing and other screen-update events).
	// Some hardware has a tickcount granularity of 15 instead of 10, so this covers more variations.
	DWORD peek_frequency_when_critical_is_on = UNINTERRUPTIBLE_PEEK_FREQUENCY; // Set default.  See below.
	// v1.0.48: Below supports "Critical 0" as meaning "Off" to improve compatibility with A_IsCritical.
	// In fact, for performance, only the following are no recognized as turning on Critical:
	//     - "On"
	//     - ""
	//     - Integer other than 0.
	// Everything else is considered to be "Off", including "Off", any non-blank string that
	// doesn't start with a non-zero number, and zero itself.
	g->ThreadIsCritical = aSetting.is_blank_or_omitted() // i.e. omitted or blank is the same as "ON". See comments above.
		|| !_tcsicmp(aSetting.value(), _T("On"))
		|| (peek_frequency_when_critical_is_on = ATOU(aSetting.value())); // Non-zero integer also turns it on. Relies on short-circuit boolean order.
	if (g->ThreadIsCritical) // Critical has been turned on. (For simplicity even if it was already on, the following is done.)
	{
		g->PeekFrequency = peek_frequency_when_critical_is_on;
		g->AllowThreadToBeInterrupted = false;
		// Ensure uninterruptibility never times out.  IsInterruptible() relies on this to avoid the
		// need to check g->ThreadIsCritical, which in turn allows global_maximize_interruptibility()
		// and DialogPrep() to avoid resetting g->ThreadIsCritical, which allows it to reliably be
		// used as the default setting for new threads, even when the auto-execute thread itself
		// (or the idle thread) needs to be interruptible, such as while displaying a dialog->
		// In other words, g->ThreadIsCritical only represents the desired setting as set by the
		// script, and isn't the actual mechanism used to make the thread uninterruptible.
		g->UninterruptibleDuration = -1;
	}
	else // Critical has been turned off.
	{
		// Since Critical is being turned off, allow thread to be immediately interrupted regardless of
		// any "Thread Interrupt" settings.
		g->PeekFrequency = DEFAULT_PEEK_FREQUENCY;
		g->AllowThreadToBeInterrupted = true;
	}
	// The thread's interruptibility has been explicitly set; so the script is now in charge of
	// managing this thread's interruptibility.
}



bif_impl FResult Thread(StrArg aCommand, optl<int> aValue1, optl<int> aValue2)
{
	switch (Line::ConvertThreadCommand(aCommand))
	{
	case THREAD_CMD_PRIORITY:
		if (aValue1.has_value())
			g->Priority = *aValue1;
		return OK;
	case THREAD_CMD_INTERRUPT:
		// If either one is blank, leave that setting as it was before.
		if (aValue1.has_value())
			g_script.mUninterruptibleTime = *aValue1;  // 32-bit (for compatibility with DWORDs returned by GetTickCount).
		if (aValue2.has_value())
			g_script.mUninterruptedLineCountMax = *aValue2;  // 32-bit also, to help performance (since huge values seem unnecessary).
		return OK;
	case THREAD_CMD_NOTIMERS:
		g->AllowTimers = (aValue1.has_value() && *aValue1 == 0); // Double-negative NoTimers=false -> allow timers.
		return OK;
	default:
		return FR_E_ARG(0);
	}
}



bif_impl void OutputDebug(StrArg aText)
{
#ifdef CONFIG_DEBUGGER
	if (!g_Debugger.OutputStdErr(aText))
#endif
		OutputDebugString(aText);
}



//////////////////////////////
// Event Handling Functions //
//////////////////////////////

bif_impl FResult OnMessage(UINT aNumber, IObject *aFunction, optl<int> aMaxThreads)
{
	// Set defaults:
	bool mode_is_delete = false;
	int max_instances = 1;
	bool call_it_last = true;

	if (aMaxThreads.has_value())
	{
		max_instances = aMaxThreads.value();
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

	// Check if this message already exists in the array:
	MsgMonitorStruct *pmonitor = g_MsgMonitor.Find(aNumber, aFunction);
	bool item_already_exists = (pmonitor != NULL);
	if (!item_already_exists)
	{
		if (mode_is_delete) // Delete a non-existent item.
			return OK;
		auto fr = ValidateFunctor(aFunction, 4);
		if (fr != OK)
			return fr;
		// From this point on, it is certain that an item will be added to the array.
		pmonitor = g_MsgMonitor.Add(aNumber, aFunction, call_it_last);
		if (!pmonitor)
			return FR_E_OUTOFMEM;
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
			return OK;
		}
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
	if (!item_already_exists || aMaxThreads.has_value())
		monitor.max_instances = max_instances;
	// Otherwise, the parameter was omitted so leave max_instances at its current value.
	return OK;
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


static FResult OnScriptEvent(IObject *aFunction, optl<int> aAddRemove, MsgMonitorList &handlers, int aParamCount)
{
	auto fr = ValidateFunctor(aFunction, aParamCount);
	if (fr != OK)
		return fr;
	
	int mode = aAddRemove.value_or(1);

	MsgMonitorStruct *existing = handlers.Find(0, aFunction);

	switch (mode)
	{
	case  1:
	case -1:
		if (!existing && !handlers.Add(0, aFunction, mode == 1))
			return FR_E_OUTOFMEM;
		break;
	case  0:
		if (existing)
			handlers.Delete(existing);
		break;
	default:
		return FR_E_ARG(1);
	}
	return OK;
}

bif_impl FResult OnClipboardChange(IObject *aFunction, optl<int> aAddRemove)
{
	auto result = OnScriptEvent(aFunction, aAddRemove, g_script.mOnClipboardChange, 1);
	// Unlike SetClipboardViewer(), AddClipboardFormatListener() doesn't cause existing
	// listeners to be called, so we can simply enable or disable our listener based on
	// whether any handlers are registered.  This has no effect if already enabled:
	g_script.EnableClipboardListener(g_script.mOnClipboardChange.Count() > 0);
	return result;
}

bif_impl FResult OnError(IObject *aFunction, optl<int> aAddRemove)
{
	return OnScriptEvent(aFunction, aAddRemove, g_script.mOnError, 2);
}

bif_impl FResult OnExit(IObject *aFunction, optl<int> aAddRemove)
{
	return OnScriptEvent(aFunction, aAddRemove, g_script.mOnExit, 2);
}



///////////////////////
// Interop Functions //
///////////////////////


bif_impl void MenuFromHandle(UINT_PTR aHandle, IObject *&aRetVal)
{
	if (aRetVal = g_script.FindMenu((HMENU)aHandle))
		aRetVal->AddRef();
}



//////////////////////
// Gui & GuiControl //
//////////////////////


bif_impl UINT_PTR IL_Create(optl<int> aInitialCount, optl<int> aGrowCount, optl<BOOL> aLargeIcons)
// Returns: Handle to the new image list, or 0 on failure.
// Parameters:
// 1: Initial image count (ImageList_Create() ignores values <=0, so no need for error checking).
// 2: Grow count (testing shows it can grow multiple times, even when this is set <=0, so it's apparently only a performance aid)
// 3: Width of each image (currently meaning small icon size when omitted or false, large icon size otherwise).
// 4: Future: Height of each image [if this param is present and >0, it would mean param 3 is not being used in its TRUE/FALSE mode)
// 5: Future: Flags/Color depth
{
	BOOL large = aLargeIcons.value_or(FALSE);
	return (UINT_PTR)ImageList_Create(
		GetSystemMetrics(large ? SM_CXICON : SM_CXSMICON),
		GetSystemMetrics(large ? SM_CYICON : SM_CYSMICON),
		ILC_MASK | ILC_COLOR32,  // ILC_COLOR32 or at least something higher than ILC_COLOR is necessary to support true-color icons.
		aInitialCount.value_or(2),  // 2 seems a better default than one, since it might be common to have only two icons in the list.
		aGrowCount.value_or(5)  // Somewhat arbitrary default.
	);
}



bif_impl BOOL IL_Destroy(UINT_PTR aImageList)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
{
	// Returns nonzero if successful, or zero otherwise, so force it to conform to TRUE/FALSE for
	// better consistency with other functions:
	return ImageList_Destroy((HIMAGELIST)aImageList) ? 1 : 0;
}



bif_impl FResult IL_Add(UINT_PTR aImageList, StrArg aFilename, optl<int> aIconNumber, optl<BOOL> aResizeNonIcon, int &aIndex)
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
	HIMAGELIST himl = (HIMAGELIST)aImageList;
	if (!himl)
		return FR_E_ARG(0);

	int param3 = aIconNumber.value_or(0);
	int icon_number, width = 0, height = 0; // Zero width/height causes image to be loaded at its actual width/height.
	if (aResizeNonIcon.has_value()) // The presence of this parameter switches mode to be "load a non-icon image".
	{
		icon_number = 0; // Zero means "load icon or bitmap (doesn't matter)".
		if (aResizeNonIcon.value()) // A value of True indicates that the image should be scaled to fit the imagelist's image size.
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
	HBITMAP hbitmap = LoadPicture(aFilename, width, height, image_type, icon_number, false); // Defaulting to "false" for "use GDIplus" provides more consistent appearance across multiple OSes.
	if (!hbitmap)
	{
		aIndex = 0;
		return OK;
	}

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
	aIndex = index;
	return OK;
}



////////////////////
// Misc Functions //
////////////////////

bif_impl FResult LoadPicture(StrArg aFilename, optl<StrArg> aOptions, int *aImageType, UINT_PTR &aRetVal)
{
	int width = -1;
	int height = -1;
	int icon_number = 0;
	bool use_gdi_plus = false;

	for (auto cp = aOptions.value_or_null(); cp; cp = StrChrAny(cp, _T(" \t")))
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
	HBITMAP hbm = LoadPicture(aFilename, width, height, image_type, icon_number, use_gdi_plus);
	if (aImageType)
		*aImageType = image_type;
	else if (image_type != IMAGE_BITMAP && hbm)
		// Always return a bitmap when the ImageType output var is omitted.
		hbm = IconToBitmap32((HICON)hbm, true); // Also works for cursors.
	aRetVal = (UINT_PTR)hbm;
	return OK;
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
	// I think this was originally intended to omit Exception(); doesn't seem to be needed anymore?
	//if (stack_top->type == DbgStack::SE_BIF && _tcsicmp(what, stack_top->func->mName))
	//	--stack_top;
	if (g_script.mNewRuntimeException != this)
	{
		// This Error is likely being constructed by the script with a call like Error(),
		// but there may be additional calls to __new.  Find the stack entry corresponding
		// to the initial call, so we have a predictable starting point for -What values
		// and a good default Line.
		for (auto se = stack_top; se >= g_Debugger.mStack.mBottom; --se)
		{
			if (se->type == DbgStack::SE_BIF && se->func == Object::sObjectCall)
			{
				line = se->line;
				stack_top = se - 1;
				break;
			}
			// In order to avoid omitting any stack frames that aren't actually related to the
			// construction of this Error object, omit only frames where this == script's this.
			// Multiple instances of the same function on the stack will all report the same
			// value due to how UDFs work, but that's very unlikely to be an issue in practice.
			// This would bypass Object.Call: (Error.Prototype.__New)({base: Error.Prototype})
			if (se->type != DbgStack::SE_UDF || !se->udf->func->mClass || se->udf->func->mParam[0].var->ToObject() != this)
				break;
		}
	}

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

	TCHAR stack_buf[SCRIPT_STACK_BUF_SIZE];
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
	case PURE_FLOAT:   return _tstof(aResult) != 0.0; // _tstof() vs. ATOF() because PURE_FLOAT is never hexadecimal.
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
		return ResultToBOOL(aToken.marker)
			|| aToken.marker_length && !*aToken.marker; // Consider any non-empty string beginning with \0 to be TRUE.
	default:
		// The only remaining valid symbol is SYM_OBJECT, which is always TRUE.
		// Check symbol anyway, in case SYM_MISSING or something else sneaks in.
		return aToken.symbol == SYM_OBJECT;
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
	case SYM_STRING: // Callers of this function expect a "numeric" result for numeric strings.
		return IsNumeric(aToken.marker, true, false, true);
	default: // SYM_MISSING or SYM_OBJECT
		return PURE_NOT_NUMERIC;
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



ResultType TokenToStringParam(ResultToken &aResultToken, ExprTokenType *aParam[], int aIndex, LPTSTR aBuf, LPTSTR &aString, size_t *aLength, bool aPermitObject)
{
	ExprTokenType &token = *aParam[aIndex];
	LPTSTR result = nullptr;
	switch (token.symbol)
	{
	case SYM_VAR:
		if (!aPermitObject && token.var->HasObject())
			break;
		aString = token.var->Contents(); // Contents() vs. mCharContents in case mCharContents needs to be updated by Contents().
		if (aLength)
			*aLength = token.var->Length();
		return OK;
	case SYM_STRING:
		aString = token.marker;
		if (aLength)
			*aLength = token.marker_length != -1 ? token.marker_length : _tcslen(token.marker);
		return OK;
	case SYM_INTEGER:
		aString = ITOA64(token.value_int64, aBuf);
		if (aLength)
			*aLength = _tcslen(aBuf);
		return OK;
	case SYM_OBJECT:
		if (aPermitObject)
		{
			aString = _T("");
			if (aLength)
				*aLength = 0;
			return OK;
		}
		break;
	case SYM_FLOAT:
		int length = FTOA(token.value_double, aString = aBuf, MAX_NUMBER_SIZE);
		if (aLength)
			*aLength = length;
		return OK;
	//case SYM_MISSING: // Caller should handle this case
	}
	return aResultToken.ParamError(aIndex, aParam[aIndex], _T("String"));
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
		aOutput.value_double = _tstof(str); // _tstof() vs. ATOF() because PURE_FLOAT is never hexadecimal.
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



FResult ValidateFunctor(IObject *aFunc, int aParamCount, int *aMinParams, bool aShowError)
{
	ResultToken result_token;
	result_token.SetResult(OK);
	return ValidateFunctor(aFunc, aParamCount, result_token, aMinParams, aShowError)
		? OK : result_token.Exited() ? FR_FAIL : FR_ABORTED;
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



bool ColorToBGR(ExprTokenType &aColorNameOrRGB, COLORREF &aBGR)
{
	switch (TypeOfToken(aColorNameOrRGB))
	{
	case SYM_INTEGER:
		aBGR = rgb_to_bgr((COLORREF)TokenToInt64(aColorNameOrRGB));
		return true;
	case SYM_STRING:
		return ColorToBGR(TokenToString(aColorNameOrRGB), aBGR);
	default:
		aBGR = 0;
		return false;
	}
}



int ConvertJoy(LPCTSTR aBuf, int *aJoystickID, bool aAllowOnlyButtons)
// The caller TextToKey() currently relies on the fact that when aAllowOnlyButtons==true, a value
// that can fit in a sc_type (USHORT) is returned, which is true since the joystick buttons
// are very small numbers (JOYCTRL_1==12).
{
	if (aJoystickID)
		*aJoystickID = 0;  // Set default output value for the caller.
	if (!aBuf || !*aBuf) return JOYCTRL_INVALID;
	auto aBuf_orig = aBuf;
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
		int offset;
		if (ParseInteger(aBuf + 3, offset))
		{
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
