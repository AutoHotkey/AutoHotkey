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
#include "script_func_impl.h"
#include "abi.h"



BIF_DECL(BIF_WinShow)
{
	auto action = _f_callee_id;

	_f_set_retval_p(_T(""), 0);

	_f_param_string_opt(aTitle, 0);
	_f_param_string_opt(aText, 1);
	// The remaining parameters depend on which function this is.

	// Set initial guess for is_ahk_group (further refined later).  For ahk_group, WinText,
	// ExcludeTitle, and ExcludeText must be blank so that they are reserved for future use
	// (i.e. they're currently not supported since the group's own criteria take precedence):
	bool is_ahk_group = !_tcsnicmp(aTitle, _T("ahk_group"), 9)
		&& ParamIndexIsOmittedOrEmpty(1) && ParamIndexIsOmittedOrEmpty(3);
	// The following is not quite accurate since is_ahk_group is only a guess at this stage, but
	// given the extreme rarity of the guess being wrong, this shortcut seems justified to reduce
	// the code size/complexity.  A wait_time of zero seems best for group closing because it's
	// currently implemented to do the wait after every window in the group.  In addition,
	// this makes "WinClose ahk_group GroupName" behave identically to "GroupClose GroupName",
	// which seems best, for consistency:
	int wait_time = is_ahk_group ? 0 : DEFAULT_WINCLOSE_WAIT;
	if (action == FID_WinClose || action == FID_WinKill) // aParam[2] contains the wait time.
	{
		if (!ParamIndexIsOmittedOrEmpty(2))
			wait_time = (int)(1000 * ParamIndexToDouble(2));
		if (!ParamIndexIsOmittedOrEmpty(4))
			is_ahk_group = false;  // Override the default.
	}
	else
		if (!ParamIndexIsOmittedOrEmpty(2))
			is_ahk_group = false;  // Override the default.
	if (is_ahk_group)
		if (WinGroup *group = g_script.FindGroup(omit_leading_whitespace(aTitle + 9)))
		{
			group->ActUponAll(action, wait_time); // It will do DoWinDelay if appropriate.
			_f_return_retval;
		}
	// Since above didn't return, either the group doesn't exist or it's paired with other
	// criteria, such as "ahk_group G ahk_class C", so do the normal single-window behavior.

	HWND target_window = NULL;
	if (aParamCount > 0)
	{
		switch (DetermineTargetHwnd(target_window, aResultToken, *aParam[0]))
		{
		case FAIL: return;
		case OK:
			if (!target_window) // Specified a HWND of 0, or IsWindow() returned false.
				_f_throw(ERR_NO_WINDOW, ErrorPrototype::Target);
		}
	}

	if (action == FID_WinClose || action == FID_WinKill)
	{
		if (target_window)
		{
			WinClose(target_window, wait_time, action == FID_WinKill);
			DoWinDelay;
			_f_return_retval;
		}
		_f_param_string_opt(aExcludeTitle, 3);
		_f_param_string_opt(aExcludeText, 4);
		if (!WinClose(*g, aTitle, aText, wait_time, aExcludeTitle, aExcludeText, action == FID_WinKill))
			// Currently WinClose returns NULL only for this case; it doesn't confirm the window closed.
			_f_throw(ERR_NO_WINDOW, ErrorPrototype::Target);
		DoWinDelay;
		_f_return_retval;
	}

	if (!target_window)
	{
		_f_param_string_opt(aExcludeTitle, 2);
		_f_param_string_opt(aExcludeText, 3);
		// By design, the WinShow command must always unhide a hidden window, even if the user has
		// specified that hidden windows should not be detected.  So set this now so that
		// DetermineTargetWindow() will make its calls in the right mode:
		bool need_restore = (_f_callee_id == FID_WinShow && !g->DetectHiddenWindows);
		if (need_restore)
			g->DetectHiddenWindows = true;
		target_window = Line::DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
		if (need_restore)
			g->DetectHiddenWindows = false;
		if (!target_window)
			_f_throw(ERR_NO_WINDOW, ErrorPrototype::Target);
	}

	// WinGroup's EnumParentActUponAll() is quite similar to the following, so the two should be
	// maintained together.

	int nCmdShow = SW_NONE; // Set default.

	switch (action)
	{
	// SW_FORCEMINIMIZE: supported only in Windows 2000/XP and beyond: "Minimizes a window,
	// even if the thread that owns the window is hung. This flag should only be used when
	// minimizing windows from a different thread."
	// My: It seems best to use SW_FORCEMINIMIZE on OS's that support it because I have
	// observed ShowWindow() to hang (thus locking up our app's main thread) if the target
	// window is hung.
	// UPDATE: For now, not using "force" every time because it has undesirable side-effects such
	// as the window not being restored to its maximized state after it was minimized
	// this way.
	case FID_WinMinimize:
		if (IsWindowHung(target_window))
		{
			nCmdShow = SW_FORCEMINIMIZE;
			// SW_MINIMIZE can lock up our thread on WinXP, which is why we revert to SW_FORCEMINIMIZE above.
			// Older/obsolete comment for background: don't attempt to minimize hung windows because that
			// might hang our thread because the call to ShowWindow() would never return.
		}
		else
			nCmdShow = SW_MINIMIZE;
		break;
	case FID_WinMaximize: if (!IsWindowHung(target_window)) nCmdShow = SW_MAXIMIZE; break;
	case FID_WinRestore:  if (!IsWindowHung(target_window)) nCmdShow = SW_RESTORE;  break;
	// Seems safe to assume it's not hung in these cases, since I'm inclined to believe
	// (untested) that hiding and showing a hung window won't lock up our thread, and
	// there's a chance they may be effective even against hung windows, unlike the
	// others above (except ACT_WINMINIMIZE, which has a special FORCE method):
	case FID_WinHide: nCmdShow = SW_HIDE; break;
	case FID_WinShow: nCmdShow = SW_SHOW; break;
	}

	// UPDATE:  Trying ShowWindowAsync()
	// now, which should avoid the problems with hanging.  UPDATE #2: Went back to
	// not using Async() because sometimes the script lines that come after the one
	// that is doing this action here rely on this action having been completed
	// (e.g. a window being maximized prior to clicking somewhere inside it).
	if (nCmdShow != SW_NONE)
	{
		// I'm not certain that SW_FORCEMINIMIZE works with ShowWindowAsync(), but
		// it probably does since there's absolutely no mention to the contrary
		// anywhere on MS's site or on the web.  But clearly, if it does work, it
		// does so only because Async() doesn't really post the message to the thread's
		// queue, instead opting for more aggressive measures.  Thus, it seems best
		// to do it this way to have maximum confidence in it:
		//if (nCmdShow == SW_FORCEMINIMIZE) // Safer not to use ShowWindowAsync() in this case.
			ShowWindow(target_window, nCmdShow);
		//else
		//	ShowWindowAsync(target_window, nCmdShow);
		DoWinDelay;
	}
	_f_return_retval;
}



BIF_DECL(BIF_WinActivate)
{
	_f_set_retval_p(_T(""), 0);

	if (aParamCount > 0)
	{
		HWND target_hwnd;
		switch (DetermineTargetHwnd(target_hwnd, aResultToken, *aParam[0]))
		{
		case FAIL: return;
		case OK:
			if (!target_hwnd)
				_f_throw(ERR_NO_WINDOW, ErrorPrototype::Target);
			SetForegroundWindowEx(target_hwnd);
			DoWinDelay;
			_f_return_retval;
		}
	}

	_f_param_string_opt(aTitle, 0);
	_f_param_string_opt(aText, 1);
	_f_param_string_opt(aExcludeTitle, 2);
	_f_param_string_opt(aExcludeText, 3);

	if (!WinActivate(*g, aTitle, aText, aExcludeTitle, aExcludeText, _f_callee_id == FID_WinActivateBottom, true))
		_f_throw(ERR_NO_WINDOW, ErrorPrototype::Target);

	// It seems best to do these sleeps here rather than in the windowing
	// functions themselves because that way, the program can use the
	// windowing functions without being subject to the script's delay
	// setting (i.e. there are probably cases when we don't need to wait,
	// such as bringing a message box to the foreground, since no other
	// actions will be dependent on it actually having happened):
	DoWinDelay;
	_f_return_retval;
}



bif_impl FResult GroupAdd(StrArg aGroup, optl<StrArg> aTitle, optl<StrArg> aText, optl<StrArg> aExcludeTitle, optl<StrArg> aExcludeText)
{
	auto group = g_script.FindGroup(aGroup, true);
	if (!group)
		return FR_FAIL; // It already displayed the error for us.
	return group->AddWindow(aTitle.value_or_null(), aText.value_or_null(), aExcludeTitle.value_or_null(), aExcludeText.value_or_null()) ? OK : FR_FAIL;
}



bif_impl FResult GroupActivate(StrArg aGroup, optl<StrArg> aMode, __int64 *aRetVal)
{
	WinGroup *group;
	if (   !(group = g_script.FindGroup(aGroup, true))   ) // Last parameter -> create-if-not-found.
		return FR_FAIL; // It already displayed the error for us.
	
	TCHAR mode = 0;
	if (aMode.has_nonempty_value())
	{
		mode = *aMode.value();
		if (mode == 'r')
			mode = 'R';
		if (mode != 'R' || aMode.value()[1])
			return FR_E_ARG(1);
	}

	HWND activated;
	group->Activate(mode == 'R', activated);
	*aRetVal = (UINT_PTR)activated;
	return OK;
}



bif_impl FResult GroupDeactivate(StrArg aGroup, optl<StrArg> aMode)
{
	auto group = g_script.FindGroup(aGroup);
	if (!group)
		return FR_E_ARG(0);
	TCHAR mode = 0;
	if (aMode.has_nonempty_value())
	{
		mode = *aMode.value();
		if (mode == 'r')
			mode = 'R';
		if (mode != 'R' || aMode.value()[1])
			return FR_E_ARG(1);
	}
	group->Deactivate(mode == 'R');
	return OK;
}



bif_impl FResult GroupClose(StrArg aGroup, optl<StrArg> aMode)
{
	auto group = g_script.FindGroup(aGroup);
	if (!group)
		return FR_E_ARG(0);
	TCHAR mode = 0;
	if (aMode.has_nonempty_value())
	{
		mode = ctoupper(*aMode.value());
		if ((mode != 'R' && mode != 'A') || aMode.value()[1])
			return FR_E_ARG(1);
	}
	if (mode == 'A')
		group->ActUponAll(FID_WinClose, 0);
	else
		group->CloseAndGoToNext(mode == 'R');
	return OK;
}



BIF_DECL(BIF_WinMove)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam + 4, aParamCount - 4))
		return;
	RECT rect;
	if (!GetWindowRect(target_window, &rect)
		|| !MoveWindow(target_window
			, ParamIndexToOptionalInt(0, rect.left) // X-position
			, ParamIndexToOptionalInt(1, rect.top)  // Y-position
			, ParamIndexToOptionalInt(2, rect.right - rect.left)
			, ParamIndexToOptionalInt(3, rect.bottom - rect.top)
			, TRUE)) // Do repaint.
		_f_throw_win32();
	DoWinDelay;
	_f_return_empty;
}



BIF_DECL(BIF_ControlSend) // ControlSend and ControlSendText.
{
	DETERMINE_TARGET_CONTROL(1);

	_f_param_string(aKeysToSend, 0);
	SendKeys(aKeysToSend, (SendRawModes)_f_callee_id, SM_EVENT, control_window);
	// But don't do WinDelay because KeyDelay should have been in effect for the above.
	_f_return_empty;
}



BIF_DECL(BIF_ControlClick)
{
	_f_param_string_opt(aControl, 0);
	_f_param_string_opt(aWhichButton, 3);
	int aVK = Line::ConvertMouseButton(aWhichButton);
	if (!aVK)
		_f_throw_param(3);
	int aClickCount = ParamIndexToOptionalInt(4, 1);
	_f_param_string_opt(aOptions, 5);

	// Set the defaults that will be in effect unless overridden by options:
	KeyEventTypes event_type = KEYDOWNANDUP;
	bool position_mode = false;
	bool do_activate = true;
	// These default coords can be overridden either by aOptions or aControl's X/Y mode:
	POINT click = {COORD_UNSPECIFIED, COORD_UNSPECIFIED};

	for (LPTSTR cp = aOptions; *cp; ++cp)
	{
		switch(ctoupper(*cp))
		{
		case 'D':
			event_type = KEYDOWN;
			break;
		case 'U':
			event_type = KEYUP;
			break;
		case 'N':
			// v1.0.45:
			// It was reported (and confirmed through testing) that this new NA mode (which avoids
			// AttachThreadInput() and SetActiveWindow()) improves the reliability of ControlClick when
			// the user is moving the mouse fairly quickly at the time the command tries to click a button.
			// In addition, the new mode avoids activating the window, which tends to happen otherwise.
			// HOWEVER, the new mode seems no more reliable than the old mode when the target window is
			// the active window.  In addition, there may be side-effects of the new mode (I caught it
			// causing Notepad's Save-As dialog to hang once, during the display of its "Overwrite?" dialog).
			// ALSO, SetControlDelay -1 seems to fix the unreliability issue as well (independently of NA),
			// though it might not work with some types of windows/controls (thus, for backward
			// compatibility, ControlClick still obeys SetControlDelay).
			if (ctoupper(cp[1]) == 'A')
			{
				cp += 1;  // Add 1 vs. 2 to skip over the rest of the letters in this option word.
				do_activate = false;
			}
			break;
		case 'P':
			if (!_tcsnicmp(cp, _T("Pos"), 3))
			{
				cp += 2;  // Add 2 vs. 3 to skip over the rest of the letters in this option word.
				position_mode = true;
			}
			break;
		// For the below:
		// Use atoi() vs. ATOI() to avoid interpreting something like 0x01D as hex
		// when in fact the D was meant to be an option letter:
		case 'X':
			click.x = _ttoi(cp + 1); // Will be overridden later below if it turns out that position_mode is in effect.
			break;
		case 'Y':
			click.y = _ttoi(cp + 1); // Will be overridden later below if it turns out that position_mode is in effect.
			break;
		}
	}
	
	HWND target_window, control_window;
	if (position_mode)
	{
		// Determine target window only.  Control will be found by position below.
		if (!DetermineTargetWindow(target_window, aResultToken, aParam + 1, aParamCount - 1, 3))
			return; // aResultToken.SetExitResult() or Error() was already called.
		control_window = NULL;
	}
	else
	{
		// Determine target window and control.
		if (!DetermineTargetControl(control_window, target_window, aResultToken, aParam, aParamCount, 3, false))
			return; // aResultToken.SetExitResult() or Error() was already called.
	}
	ASSERT(target_window != NULL);

	// It's debatable, but might be best for flexibility (and backward compatibility) to allow target_window to itself
	// be a control (at least for the position_mode handler below).  For example, the script may have called SetParent
	// to make a top-level window the child of some other window, in which case this policy allows it to be seen like
	// a non-child.
	if (!control_window) // Even if position_mode is false, the below is still attempted, as documented.
	{
		// New section for v1.0.24.  But only after the above fails to find a control do we consider
		// whether aControl contains X and Y coordinates.  That way, if a control class happens to be
		// named something like "X1 Y1", it will still be found by giving precedence to class names.
		point_and_hwnd_type pah = {0};
		pah.ignore_disabled_controls = true; // v1.1.20: Ignore disabled controls.
		// Parse the X an Y coordinates in a strict way to reduce ambiguity with control names and also
		// to keep the code simple.
		LPTSTR cp = omit_leading_whitespace(aControl);
		if (ctoupper(*cp) != 'X')
			goto control_error;
		++cp;
		if (!*cp)
			goto control_error;
		pah.pt.x = ATOI(cp);
		if (   !(cp = StrChrAny(cp, _T(" \t")))   ) // Find next space or tab (there must be one for it to be considered valid).
			goto control_error;
		cp = omit_leading_whitespace(cp + 1);
		if (!*cp || _totupper(*cp) != 'Y')
			goto control_error;
		++cp;
		if (!*cp)
			goto control_error;
		pah.pt.y = ATOI(cp);
		// The passed-in coordinates are always relative to target_window's client area because offering
		// an option for absolute/screen coordinates doesn't seem useful.
		ClientToScreen(target_window, &pah.pt); // Convert to screen coordinates.
		EnumChildWindows(target_window, EnumChildFindPoint, (LPARAM)&pah); // Find topmost control containing point.
		// If no control is at this point, try posting the mouse event message(s) directly to the
		// parent window to increase the flexibility of this feature:
		control_window = pah.hwnd_found ? pah.hwnd_found : target_window;
		// Convert click's target coordinates to be relative to the client area of the control or
		// parent window because that is the format required by messages such as WM_LBUTTONDOWN
		// used later below:
		click = pah.pt;
		ScreenToClient(control_window, &click);
	}

	// This is done this late because it seems better to throw an exception whenever the
	// target window or control isn't found, or any other error condition occurs above:
	if (aClickCount < 1)
	{
		// Allow this to simply "do nothing", because it increases flexibility
		// in the case where the number of clicks is a dereferenced script variable
		// that may sometimes (by intent) resolve to zero or negative:
		_f_return_empty;
	}

	RECT rect;
	if (click.x == COORD_UNSPECIFIED || click.y == COORD_UNSPECIFIED)
	{
		// The following idea is from AutoIt3. It states: "Get the dimensions of the control so we can click
		// the centre of it" (maybe safer and more natural than 0,0).
		// My: In addition, this is probably better for some large controls (e.g. SysListView32) because
		// clicking at 0,0 might activate a part of the control that is not even visible:
		if (!GetWindowRect(control_window, &rect))
			goto error;
		if (click.x == COORD_UNSPECIFIED)
			click.x = (rect.right - rect.left) / 2;
		if (click.y == COORD_UNSPECIFIED)
			click.y = (rect.bottom - rect.top) / 2;
	}

	UINT msg_down, msg_up;
	WPARAM wparam, wparam_up = 0;
	bool vk_is_wheel = aVK == VK_WHEEL_UP || aVK == VK_WHEEL_DOWN;
	bool vk_is_hwheel = aVK == VK_WHEEL_LEFT || aVK == VK_WHEEL_RIGHT; // v1.0.48: Lexikos: Support horizontal scrolling in Windows Vista and later.

	if (vk_is_wheel)
	{
		ClientToScreen(control_window, &click); // Wheel messages use screen coordinates.
		wparam = (aClickCount * ((aVK == VK_WHEEL_UP) ? WHEEL_DELTA : -WHEEL_DELTA)) << 16;  // High order word contains the delta.
		msg_down = WM_MOUSEWHEEL;
		// Make the event more accurate by having the state of the keys reflected in the event.
		// The logical state (not physical state) of the modifier keys is used so that something
		// like this is supported:
		// Send, {ShiftDown}
		// MouseClick, WheelUp
		// Send, {ShiftUp}
		// In addition, if the mouse hook is installed, use its logical mouse button state so that
		// something like this is supported:
		// MouseClick, left, , , , , D  ; Hold down the left mouse button
		// MouseClick, WheelUp
		// MouseClick, left, , , , , U  ; Release the left mouse button.
		// UPDATE: Since the other ControlClick types (such as leftclick) do not reflect these
		// modifiers -- and we want to keep it that way, at least by default, for compatibility
		// reasons -- it seems best for consistency not to do them for WheelUp/Down either.
		// A script option can be added in the future to obey the state of the modifiers:
		//mod_type mod = GetModifierState();
		//if (mod & MOD_SHIFT)
		//	wparam |= MK_SHIFT;
		//if (mod & MOD_CONTROL)
		//	wparam |= MK_CONTROL;
        //if (g_MouseHook)
		//	wparam |= g_mouse_buttons_logical;
	}
	else if (vk_is_hwheel)	// Lexikos: Support horizontal scrolling in Windows Vista and later.
	{
		wparam = (aClickCount * ((aVK == VK_WHEEL_LEFT) ? -WHEEL_DELTA : WHEEL_DELTA)) << 16;
		msg_down = WM_MOUSEHWHEEL;
	}
	else
	{
		switch (aVK)
		{
			case VK_LBUTTON:  msg_down = WM_LBUTTONDOWN; msg_up = WM_LBUTTONUP; wparam = MK_LBUTTON; break;
			case VK_RBUTTON:  msg_down = WM_RBUTTONDOWN; msg_up = WM_RBUTTONUP; wparam = MK_RBUTTON; break;
			case VK_MBUTTON:  msg_down = WM_MBUTTONDOWN; msg_up = WM_MBUTTONUP; wparam = MK_MBUTTON; break;
			case VK_XBUTTON1: msg_down = WM_XBUTTONDOWN; msg_up = WM_XBUTTONUP; wparam_up = XBUTTON1<<16; wparam = MK_XBUTTON1|wparam_up; break;
			case VK_XBUTTON2: msg_down = WM_XBUTTONDOWN; msg_up = WM_XBUTTONUP; wparam_up = XBUTTON2<<16; wparam = MK_XBUTTON2|wparam_up; break;
			default: // Just do nothing since this should realistically never happen.
				ASSERT(!"aVK value not handled");
		}
	}

	LPARAM lparam = MAKELPARAM(click.x, click.y);

	// SetActiveWindow() requires ATTACH_THREAD_INPUT to succeed.  Even though the MSDN docs state
	// that SetActiveWindow() has no effect unless the parent window is foreground, Jon insists
	// that SetActiveWindow() resolved some problems for some users.  In any case, it seems best
	// to do this in case the window really is foreground, in which case MSDN indicates that
	// it will help for certain types of dialogs.
	ATTACH_THREAD_INPUT_AND_SETACTIVEWINDOW_IF_DO_ACTIVATE  // It's kept with a similar macro for maintainability.
	// v1.0.44.13: Notes for the above: Unlike some other Control commands, GetNonChildParent() is not
	// called here when target_window==control_window.  This is because the script may have called
	// SetParent to make target_window the child of some other window, in which case target_window
	// should still be used above (unclear).  Perhaps more importantly, it's allowed for control_window
	// to be the same as target_window, at least in position_mode, whose docs state, "If there is no
	// control, the target window itself will be sent the event (which might have no effect depending
	// on the nature of the window)."  In other words, it seems too complicated and rare to add explicit
	// handling for "ahk_id %ControlHWND%" (though the below rules should work).
	// The line "ControlClick,, ahk_id %HWND%" can have multiple meanings depending on the nature of HWND:
	// 1) If HWND is a top-level window, its topmost child will be clicked.
	// 2) If HWND is a top-level window that has become a child of another window via SetParent: same.
	// 3) If HWND is a control, its topmost child will be clicked (or itself if it has no children).
	//    For example, the following works (as documented in the first parameter):
	//    ControlGet, HWND, HWND,, OK, A  ; Get the HWND of the OK button.
	//    ControlClick,, ahk_id %HWND%

	if (vk_is_wheel || vk_is_hwheel) // v1.0.48: Lexikos: Support horizontal scrolling in Windows Vista and later.
	{
		PostMessage(control_window, msg_down, wparam, lparam);
		DoControlDelay;
	}
	else
	{
		for (int i = 0; i < aClickCount; ++i)
		{
			if (event_type != KEYUP) // It's either down-only or up-and-down so always to the down-event.
			{
				PostMessage(control_window, msg_down, wparam, lparam);
				// Seems best to do this one too, which is what AutoIt3 does also.  User can always reduce
				// ControlDelay to 0 or -1.  Update: Jon says this delay might be causing it to fail in
				// some cases.  Upon reflection, it seems best not to do this anyway because PostMessage()
				// should queue up the message for the app correctly even if it's busy.  Update: But I
				// think the timestamp is available on every posted message, so if some apps check for
				// inhumanly fast clicks (to weed out transients with partial clicks of the mouse, or
				// to detect artificial input), the click might not work.  So it might be better after
				// all to do the delay until it's proven to be problematic (Jon implies that he has
				// no proof yet).  IF THIS IS EVER DISABLED, be sure to do the ControlDelay anyway
				// if event_type == KEYDOWN:
				DoControlDelay;
			}
			if (event_type != KEYDOWN) // It's either up-only or up-and-down so always to the up-event.
			{
				PostMessage(control_window, msg_up, wparam_up, lparam);
				DoControlDelay;
			}
		}
	}

	DETACH_THREAD_INPUT  // Also takes into account do_activate, indirectly.

	_f_return_empty;

error:
	_f_throw_win32();

control_error:
	_f_throw(ERR_NO_CONTROL, aControl, ErrorPrototype::Target);
}



BIF_DECL(BIF_ControlMove)
{
	DETERMINE_TARGET_CONTROL(4);
	
	// The following macro is used to keep ControlMove and ControlGetPos in sync:
	#define CONTROL_COORD_PARENT(target, control) \
		(target == control ? GetNonChildParent(target) : target)

	// Determine which window the supplied coordinates are relative to:
	HWND coord_parent = CONTROL_COORD_PARENT(target_window, control_window);

	// Determine the controls current coordinates relative to coord_parent in case one
	// or more parameters were omitted.
	RECT control_rect;
	if (!GetWindowRect(control_window, &control_rect)
		|| !MapWindowPoints(NULL, coord_parent, (LPPOINT)&control_rect, 2))
		_f_throw_win32();
	
	POINT point;
	point.x = ParamIndexToOptionalInt(0, control_rect.left);
	point.y = ParamIndexToOptionalInt(1, control_rect.top);

	// MoveWindow accepts coordinates relative to the control's immediate parent, which might
	// be different to coord_parent since controls can themselves have child controls.  So if
	// necessary, map the caller-supplied coordinates to the control's immediate parent:
	HWND immediate_parent = GetParent(control_window);
	if (immediate_parent != coord_parent)
		MapWindowPoints(coord_parent, immediate_parent, &point, 1);

	MoveWindow(control_window
		, point.x
		, point.y
		, ParamIndexToOptionalInt(2, control_rect.right - control_rect.left)
		, ParamIndexToOptionalInt(3, control_rect.bottom - control_rect.top)
		, TRUE);  // Do repaint.

	DoControlDelay
	_f_return_empty;
}



BIF_DECL(BIF_ControlGetPos)
{
	Var *output_var_x = ParamIndexToOutputVar(0);
	Var *output_var_y = ParamIndexToOutputVar(1);
	Var *output_var_width = ParamIndexToOutputVar(2);
	Var *output_var_height = ParamIndexToOutputVar(3);

	DETERMINE_TARGET_CONTROL(4);

	// Determine which window the returned coordinates should be relative to:
	HWND coord_parent = CONTROL_COORD_PARENT(target_window, control_window);

	RECT child_rect;
	// Realistically never fails since DetermineTargetWindow() and ControlExist() should always yield
	// valid window handles:
	GetWindowRect(control_window, &child_rect);
	// Map the screen coordinates returned by GetWindowRect to the client area of coord_parent.
	MapWindowPoints(NULL, coord_parent, (LPPOINT)&child_rect, 2);

	output_var_x && output_var_x->Assign(child_rect.left);
	output_var_y && output_var_y->Assign(child_rect.top);
	output_var_width && output_var_width->Assign(child_rect.right - child_rect.left);
	output_var_height && output_var_height->Assign(child_rect.bottom - child_rect.top);
	_f_return_empty;
}



BIF_DECL(BIF_ControlGetFocus)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam, aParamCount))
		return;

	GUITHREADINFO guithreadInfo;
	guithreadInfo.cbSize = sizeof(GUITHREADINFO);
	if (!GetGUIThreadInfo(GetWindowThreadProcessId(target_window, NULL), &guithreadInfo))
		_f_throw_win32();

	// Use IsChild() to ensure the focused control actually belongs to this window.
	// Otherwise, a HWND will be returned if any window in the same thread has focus,
	// including the target window itself (typically when it has no controls).
	if (!IsChild(target_window, guithreadInfo.hwndFocus))
		_f_return_i(0); // As documented, if "none of the target window's controls has focus, the return value is 0".
	_f_return_i((UINT_PTR)guithreadInfo.hwndFocus);
}



BIF_DECL(BIF_ControlGetClassNN)
{
	DETERMINE_TARGET_CONTROL(0);

	if (target_window == control_window)
		target_window = GetNonChildParent(control_window);

	class_and_hwnd_type cah;
	TCHAR class_name[WINDOW_CLASS_SIZE];
	cah.hwnd = control_window;
	cah.class_name = class_name;
	if (!GetClassName(cah.hwnd, class_name, _countof(class_name) - 5)) // -5 to allow room for sequence number.
		_f_throw_win32();
	
	cah.class_count = 0;  // Init for the below.
	cah.is_found = false; // Same.
	EnumChildWindows(target_window, EnumChildFindSeqNum, (LPARAM)&cah);
	if (!cah.is_found)
		_f_throw(ERR_FAILED);
	// Append the class sequence number onto the class name:
	sntprintfcat(class_name, _countof(class_name), _T("%d"), cah.class_count);
	_f_return(class_name);
}



BOOL CALLBACK EnumChildFindSeqNum(HWND aWnd, LPARAM lParam)
{
	class_and_hwnd_type &cah = *(class_and_hwnd_type *)lParam;  // For performance and convenience.
	TCHAR class_name[WINDOW_CLASS_SIZE];
	if (!GetClassName(aWnd, class_name, _countof(class_name)))
		return TRUE;  // Continue the enumeration.
	if (!_tcscmp(class_name, cah.class_name)) // Class names match.
	{
		++cah.class_count;
		if (aWnd == cah.hwnd)  // The caller-specified window has been found.
		{
			cah.is_found = true;
			return FALSE;
		}
	}
	return TRUE; // Continue enumeration until a match is found or there aren't any windows remaining.
}



BIF_DECL(BIF_ControlFocus)
{
	DETERMINE_TARGET_CONTROL(0);

	// Unlike many of the other Control commands, this one requires AttachThreadInput()
	// to have any realistic chance of success (though sometimes it may work by pure
	// chance even without it):
	ATTACH_THREAD_INPUT

	SetFocus(control_window);
	DoControlDelay; // Done unconditionally for simplicity, and in case SetFocus() had some effect despite indicating failure.
	// GetFocus() isn't called and failure to focus isn't treated as an error because
	// a successful change in focus doesn't guarantee that the focus will still be as
	// expected when the next line of code runs.

	// Very important to detach any threads whose inputs were attached above,
	// prior to returning, otherwise the next attempt to attach thread inputs
	// for these particular windows may result in a hung thread or other
	// undesirable effect:
	DETACH_THREAD_INPUT

	_f_return_empty;
}



BIF_DECL(BIF_ControlSetText)
{
	DETERMINE_TARGET_CONTROL(1);

	_f_param_string(aNewText, 0);
	// SendMessage must be used, not PostMessage(), at least for some (probably most) apps.
	// Also: No need to call IsWindowHung() because SendMessageTimeout() should return
	// immediately if the OS already "knows" the window is hung:
	DWORD_PTR result;
	SendMessageTimeout(control_window, WM_SETTEXT, (WPARAM)0, (LPARAM)aNewText
		, SMTO_ABORTIFHUNG, 5000, &result);
	DoControlDelay;
	_f_return_empty;
}



BIF_DECL(BIF_ControlGetText)
{
	DETERMINE_TARGET_CONTROL(0);

	// Even if control_window is NULL, we want to continue on so that the output
	// param is set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before.

	// Handle the output parameter.  Note: Using GetWindowTextTimeout() vs. GetWindowText()
	// because it is able to get text from more types of controls (e.g. large edit controls):
	VarSizeType space_needed = GetWindowTextTimeout(control_window) + 1; // 1 for terminator.

	// Allocate memory for the return value.
	if (!TokenSetResult(aResultToken, NULL, space_needed - 1))
		return;  // It already displayed the error.
	aResultToken.symbol = SYM_STRING;
	// Fetch the text directly into the buffer.  Also set the length explicitly
	// in case actual size written was off from the estimated size (since
	// GetWindowTextLength() can return more space that will actually be required
	// in certain circumstances, see MS docs):
	aResultToken.marker_length = GetWindowTextTimeout(control_window, aResultToken.marker, space_needed);
	if (!aResultToken.marker_length) // There was no text to get or GetWindowTextTimeout() failed.
		*aResultToken.marker = '\0';
}



void ControlGetListView(ResultToken &aResultToken, HWND aHwnd, LPTSTR aOptions)
// Called by ControlGet() below.  It has ensured that aHwnd is a valid handle to a ListView.
{
	// GET ROW COUNT
	LRESULT row_count;
	if (!SendMessageTimeout(aHwnd, LVM_GETITEMCOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&row_count)) // Timed out or failed.
		goto error;

	// GET COLUMN COUNT
	// Through testing, could probably get to a level of 90% certainty that a ListView for which
	// InsertColumn() was never called (or was called only once) might lack a header control if the LV is
	// created in List/Icon view-mode and/or with LVS_NOCOLUMNHEADER. The problem is that 90% doesn't
	// seem to be enough to justify elimination of the code for "undetermined column count" mode.  If it
	// ever does become a certainty, the following could be changed:
	// 1) The extra code for "undetermined" mode rather than simply forcing col_count to be 1.
	// 2) Probably should be kept for compatibility: -1 being returned when undetermined "col count".
	//
	// The following approach might be the only simple yet reliable way to get the column count (sending
	// LVM_GETITEM until it returns false doesn't work because it apparently returns true even for
	// nonexistent subitems -- the same is reported to happen with LVM_GETCOLUMN and such, though I seem
	// to remember that LVM_SETCOLUMN fails on non-existent columns -- but calling that on a ListView
	// that isn't in Report view has been known to traumatize the control).
	// Fix for v1.0.37.01: It appears that the header doesn't always exist.  For example, when an
	// Explorer window opens and is *initially* in icon or list mode vs. details/tiles mode, testing
	// shows that there is no header control.  Testing also shows that there is exactly one column
	// in such cases but only for Explorer and other things that avoid creating the invisible columns.
	// For example, a script can create a ListView in Icon-mode and give it retrievable column data for
	// columns beyond the first.  Thus, having the undetermined-col-count mode preserves flexibility
	// by allowing individual columns beyond the first to be retrieved.  On a related note, testing shows
	// that attempts to explicitly retrieve columns (i.e. fields/subitems) other than the first in the
	// case of Explorer's Icon/List view modes behave the same as fetching the first column (i.e. Col3
	// would retrieve the same text as specifying Col1 or not having the Col option at all).
	// Obsolete because not always true: Testing shows that a ListView always has a header control
	// (at least on XP), even if you can't see it (such as when the view is Icon/Tile or when -Hdr has
	// been specified in the options).
	HWND header_control;
	LRESULT col_count = -1;  // Fix for v1.0.37.01: Use -1 to indicate "undetermined col count".
	if (SendMessageTimeout(aHwnd, LVM_GETHEADER, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&header_control)
		&& header_control) // Relies on short-circuit boolean order.
		SendMessageTimeout(header_control, HDM_GETITEMCOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&col_count);
		// Return value is not checked because if it fails, col_count is left at its default of -1 set above.
		// In fact, if any of the above conditions made it impossible to determine col_count, col_count stays
		// at -1 to indicate "undetermined".

	// PARSE OPTIONS (a simple vs. strict method is used to reduce code size)
	bool get_count = tcscasestr(aOptions, _T("Count"));
	bool include_selected_only = tcscasestr(aOptions, _T("Selected")); // Explicit "ed" to reserve "Select" for possible future use.
	bool include_focused_only = tcscasestr(aOptions, _T("Focused"));  // Same.
	LPTSTR col_option = tcscasestr(aOptions, _T("Col")); // Also used for mode "Count Col"
	int requested_col = col_option ? ATOI(col_option + 3) - 1 : -1;
	if (col_option && (get_count ? col_option[3] && !IS_SPACE_OR_TAB(col_option[3]) // "Col" has a suffix.
		: (requested_col < 0 || col_count > -1 && requested_col >= col_count))) // Specified column does not exist.
		_f_throw_value(ERR_PARAM1_INVALID, col_option);

	// IF THE "COUNT" OPTION IS PRESENT, FULLY HANDLE THAT AND RETURN
	if (get_count)
	{
		int result; // Must be signed to support writing a col count of -1 to aOutputVar.
		if (include_focused_only) // Listed first so that it takes precedence over include_selected_only.
		{
			if (!SendMessageTimeout(aHwnd, LVM_GETNEXTITEM, -1, LVNI_FOCUSED, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&result)) // Timed out or failed.
				goto error;
			++result; // i.e. Set it to 0 if not found, or the 1-based row-number otherwise.
		}
		else if (include_selected_only)
		{
			if (!SendMessageTimeout(aHwnd, LVM_GETSELECTEDCOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&result)) // Timed out or failed.
				goto error;
		}
		else if (col_option) // "Count Col" returns the number of columns.
			result = (int)col_count;
		else // Total row count.
			result = (int)row_count;
		_f_return(result);
	}

	// FINAL CHECKS
	if (row_count < 1 || !col_count) // But don't return when col_count == -1 (i.e. always make the attempt when col count is undetermined).
		_f_return_empty;  // No text in the control, so indicate success.
	
	// Notes about the following struct definitions:  The layout of LVITEM depends on
	// which platform THIS executable was compiled for, but we need it to match what
	// the TARGET process expects.  If the target process is 32-bit and we are 64-bit
	// or vice versa, LVITEM can't be used.  The following structs are copies of
	// LVITEM with UINT (32-bit) or UINT64 (64-bit) in place of the pointer fields.
	struct LVITEM32
	{
		UINT mask;
		int iItem;
		int iSubItem;
		UINT state;
		UINT stateMask;
		UINT pszText;
		int cchTextMax;
		int iImage;
		UINT lParam;
		int iIndent;
		int iGroupId;
		UINT cColumns;
		UINT puColumns;
		UINT piColFmt;
		int iGroup;
	};
	struct LVITEM64
	{
		UINT mask;
		int iItem;
		int iSubItem;
		UINT state;
		UINT stateMask;
		UINT64 pszText;
		int cchTextMax;
		int iImage;
		UINT64 lParam;
		int iIndent;
		int iGroupId;
		UINT cColumns;
		UINT64 puColumns;
		UINT64 piColFmt;
		int iGroup;
	};
	union
	{
		LVITEM32 i32;
		LVITEM64 i64;
	} local_lvi;

	// ALLOCATE INTERPROCESS MEMORY FOR TEXT RETRIEVAL
	HANDLE handle;
	LPVOID p_remote_lvi; // Not of type LPLVITEM to help catch bugs where p_remote_lvi->member is wrongly accessed here in our process.
	if (   !(p_remote_lvi = AllocInterProcMem(handle, sizeof(local_lvi) + _TSIZE(LV_REMOTE_BUF_SIZE), aHwnd, PROCESS_QUERY_INFORMATION))   ) // Allocate both the LVITEM struct and its internal string buffer in one go because VirtualAllocEx() is probably a high overhead call.
		goto error;
	LPVOID p_remote_text = (LPVOID)((UINT_PTR)p_remote_lvi + sizeof(local_lvi)); // The next buffer is the memory area adjacent to, but after the struct.
	
	// PREPARE LVI STRUCT MEMBERS FOR TEXT RETRIEVAL
	if (IsProcess64Bit(handle))
	{
		// See the section below for comments.
		local_lvi.i64.cchTextMax = LV_REMOTE_BUF_SIZE - 1;
		local_lvi.i64.pszText = (UINT64)p_remote_text;
	}
	else
	{
		// Subtract 1 because of that nagging doubt about size vs. length. Some MSDN examples subtract one,
		// such as TabCtrl_GetItem()'s cchTextMax:
		local_lvi.i32.cchTextMax = LV_REMOTE_BUF_SIZE - 1; // Note that LVM_GETITEM doesn't update this member to reflect the new length.
		local_lvi.i32.pszText = (UINT)(UINT_PTR)p_remote_text; // Extra cast avoids a truncation warning (C4311).
	}

	LRESULT i, next, length, total_length;
	bool is_selective = include_focused_only || include_selected_only;
	bool single_col_mode = (requested_col > -1 || col_count == -1); // Get only one column in these cases.

	// ESTIMATE THE AMOUNT OF MEMORY NEEDED TO STORE ALL THE TEXT
	// It's important to note that a ListView might legitimately have a collection of rows whose
	// fields are all empty.  Since it is difficult to know whether the control is truly owner-drawn
	// (checking its style might not be enough?), there is no way to distinguish this condition
	// from one where the control's text can't be retrieved due to being owner-drawn.  In any case,
	// this all-empty-field behavior simplifies the code and will be documented in the help file.
	for (i = 0, next = -1, total_length = 0; i < row_count; ++i) // For each row:
	{
		if (is_selective)
		{
			// Fix for v1.0.37.01: Prevent an infinite loop that might occur if the target control no longer
			// exists (perhaps having been closed in the middle of the operation) or is permanently hung.
			// If GetLastError() were to return zero after the below, it would mean the function timed out.
			// However, rather than checking and retrying, it seems better to abort the operation because:
			// 1) Timeout should be quite rare.
			// 2) Reduces code size.
			// 3) Having a retry really should be accompanied by SLEEP_WITHOUT_INTERRUPTION because all this
			//    time our thread would not pumping messages (and worse, if the keyboard/mouse hooks are installed,
			//    mouse/key lag would occur).
			if (!SendMessageTimeout(aHwnd, LVM_GETNEXTITEM, next, include_focused_only ? LVNI_FOCUSED : LVNI_SELECTED
				, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&next) // Timed out or failed.
				|| next == -1) // No next item.  Relies on short-circuit boolean order.
				break; // End of estimation phase (if estimate is too small, the text retrieval below will truncate it).
		}
		else
			next = i;
		for (local_lvi.i32.iSubItem = (requested_col > -1) ? requested_col : 0 // iSubItem is which field to fetch. If it's zero, the item vs. subitem will be fetched.
			; col_count == -1 || local_lvi.i32.iSubItem < col_count // If column count is undetermined (-1), always make the attempt.
			; ++local_lvi.i32.iSubItem) // For each column:
		{
			if (WriteProcessMemory(handle, p_remote_lvi, &local_lvi, sizeof(local_lvi), NULL)
				&& SendMessageTimeout(aHwnd, LVM_GETITEMTEXT, next, (LPARAM)p_remote_lvi, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&length))
				total_length += length;
			//else timed out or failed, don't include the length in the estimate.  Instead, the
			// text-fetching routine below will ensure the text doesn't overflow the var capacity.
			if (single_col_mode)
				break;
		}
	}
	// Add to total_length enough room for one linefeed per row, and one tab after each column
	// except the last (formula verified correct, though it's inflated by 1 for safety). "i" contains the
	// actual number of rows that will be transcribed, which might be less than row_count if is_selective==true.
	total_length += i * (single_col_mode ? 1 : col_count);

	// SET UP THE OUTPUT BUFFER
	if (!TokenSetResult(aResultToken, NULL, (size_t)total_length))
		goto cleanup_and_return; // Error() was already called.
	aResultToken.symbol = SYM_STRING;
	
	LPTSTR contents = aResultToken.marker;
	LRESULT capacity = total_length; // LRESULT avoids signed vs. unsigned compiler warnings.
	if (capacity > 0) // For maintainability, avoid going negative.
		--capacity; // Adjust to exclude the zero terminator, which simplifies things below.

	// RETRIEVE THE TEXT FROM THE REMOTE LISTVIEW
	// Start total_length at zero in case actual size is greater than estimate, in which case only a partial set of text along with its '\t' and '\n' chars will be written.
	for (i = 0, next = -1, total_length = 0; i < row_count; ++i) // For each row:
	{
		if (is_selective)
		{
			// Fix for v1.0.37.01: Prevent an infinite loop (for details, see comments in the estimation phase above).
			if (!SendMessageTimeout(aHwnd, LVM_GETNEXTITEM, next, include_focused_only ? LVNI_FOCUSED : LVNI_SELECTED
				, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&next) // Timed out or failed.
				|| next == -1) // No next item.
				break; // See comment above for why unconditional break vs. continue.
		}
		else // Retrieve every row, so the "next" row becomes the "i" index.
			next = i;
		// Insert a linefeed before each row except the first:
		if (i && total_length < capacity) // If we're at capacity, it will exit the loops when the next field is read.
		{
			*contents++ = '\n';
			++total_length;
		}

		// iSubItem is which field to fetch. If it's zero, the item vs. subitem will be fetched:
		for (local_lvi.i32.iSubItem = (requested_col > -1) ? requested_col : 0
			; col_count == -1 || local_lvi.i32.iSubItem < col_count // If column count is undetermined (-1), always make the attempt.
			; ++local_lvi.i32.iSubItem) // For each column:
		{
			// Insert a tab before each column except the first and except when in single-column mode:
			if (!single_col_mode && local_lvi.i32.iSubItem && total_length < capacity)  // If we're at capacity, it will exit the loops when the next field is read.
			{
				*contents++ = '\t';
				++total_length;
			}

			if (!WriteProcessMemory(handle, p_remote_lvi, &local_lvi, sizeof(local_lvi), NULL)
				|| !SendMessageTimeout(aHwnd, LVM_GETITEMTEXT, next, (LPARAM)p_remote_lvi, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&length))
				continue; // Timed out or failed. It seems more useful to continue getting text rather than aborting the operation.

			// Otherwise, the message was successfully sent.
			if (length > 0)
			{
				if (total_length + length > capacity)
					goto break_both; // "goto" for simplicity and code size reduction.
				// Otherwise:
				// READ THE TEXT FROM THE REMOTE PROCESS
				// Although MSDN has the following comment about LVM_GETITEM, it is not present for
				// LVM_GETITEMTEXT. Therefore, to improve performance (by avoiding a second call to
				// ReadProcessMemory) and to reduce code size, we'll take them at their word until
				// proven otherwise.  Here is the MSDN comment about LVM_GETITEM: "Applications
				// should not assume that the text will necessarily be placed in the specified
				// buffer. The control may instead change the pszText member of the structure
				// to point to the new text, rather than place it in the buffer."
				if (ReadProcessMemory(handle, p_remote_text, contents, length * sizeof(TCHAR), NULL))
				{
					contents += length; // Point it to the position where the next char will be written.
					total_length += length; // Recalculate length in case its different than the estimate (for any reason).
				}
				//else it failed; but even so, continue on to put in a tab (if called for).
			}
			//else length is zero; but even so, continue on to put in a tab (if called for).
			if (single_col_mode)
				break;
		} // for() each column
	} // for() each row

break_both:
	*contents = '\0'; // Final termination.  Above has reserved room for this one byte.
	aResultToken.marker_length = (size_t)total_length; // Update to actual vs. estimated length.

	// CLEAN UP
cleanup_and_return: // This is "called" if a memory allocation failed above
	FreeInterProcMem(handle, p_remote_lvi);
	return;

error:
	_f_throw_win32();
}



bool ControlSetTab(ResultToken &aResultToken, HWND aHwnd, DWORD aTabIndex)
{
	DWORD_PTR dwResult;
	// MSDN: "If the tab control does not have the TCS_BUTTONS style, changing the focus also changes
	// the selected tab. In this case, the tab control sends the TCN_SELCHANGING and TCN_SELCHANGE
	// notification codes to its parent window."
	if (!SendMessageTimeout(aHwnd, TCM_SETCURFOCUS, aTabIndex, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
		return false;
	// Tab controls with the TCS_BUTTONS style need additional work:
	if (GetWindowLong(aHwnd, GWL_STYLE) & TCS_BUTTONS)
	{
		// Problem:
		//  TCM_SETCURFOCUS does not change the selected tab if TCS_BUTTONS is set.
		//
		// False solution #1 (which used to be recommended in the docs):
		//  Send a TCM_SETCURSEL method afterward.  TCM_SETCURSEL changes the selected tab,
		//  but doesn't notify the control's parent, so it doesn't update the tab's contents.
		//
		// False solution #2:
		//  Send a WM_NOTIFY message to the parent window to notify it.  Can't be done.
		//  MSDN says: "For Windows 2000 and later systems, the WM_NOTIFY message cannot
		//  be sent between processes."
		//
		// Solution #1:
		//  Send VK_LEFT/VK_RIGHT as many times as needed.
		//
		// Solution #2:
		//  Set the focus to an adjacent tab and then send VK_LEFT/VK_RIGHT.
		//   - Must choose an appropriate tab index and vk depending on which tab is being
		//     selected, since VK_LEFT/VK_RIGHT don't wrap around.
		//   - Ends up tempting optimisations which increase code size, such as to avoid
		//     TCM_SETCURFOCUS if an adjacent tab is already focused.
		//   - Still needs VK_SPACE afterward to actually select the tab.
		//
		// Solution #3 (the one below):
		//  Set the focus to the appropriate tab and then send VK_SPACE.
		//   - Since we've already set the focus, all we need to do is send VK_SPACE.
		//   - If the tab index is invalid and the user has focused but not selected
		//     another tab, that tab will be selected.  This seems harmless enough.
		//
		PostMessage(aHwnd, WM_KEYDOWN, VK_SPACE, 0x00000001);
		PostMessage(aHwnd, WM_KEYUP, VK_SPACE, 0xC0000001);
	}
	return true;
}



BIF_DECL(BIF_StatusBarGetText)//(LPTSTR aPart, LPTSTR aTitle, LPTSTR aText
	//, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	int part = ParamIndexToOptionalInt(0, 1);
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam + 1, aParamCount - 1))
		return;
	HWND control_window = ControlExist(target_window, _T("msctls_statusbar321"));
	// StatusBarUtil will handle any NULL control_window or zero part# for us.
	StatusBarUtil(aResultToken, control_window, part);
}



BIF_DECL(BIF_StatusBarWait)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam + 3, aParamCount - 3, 1))
		return;

	LPTSTR aTextToWaitFor = ParamIndexToOptionalString(0, _f_number_buf);
	int aSeconds = ParamIndexIsOmittedOrEmpty(1) ? -1 : int(ParamIndexToDouble(1) * 1000);
	int aPart = ParamIndexToOptionalInt(2, 0);
	int aInterval = ParamIndexToOptionalInt(5, 50);

	// Make a copy of any memory areas that are volatile (due to caller passing a variable,
	// which could be reassigned by a new hotkey subroutine launched while we are waiting)
	// but whose contents we need to refer to while we are waiting:
	TCHAR text_to_wait_for[4096];
	tcslcpy(text_to_wait_for, aTextToWaitFor, _countof(text_to_wait_for));
	HWND control_window = ControlExist(target_window, _T("msctls_statusbar321"));
	// StatusBarUtil will handle any NULL control_window or zero part# for us.
	StatusBarUtil(aResultToken, control_window, aPart, text_to_wait_for, aSeconds, aInterval);
}



BIF_DECL(BIF_PostSendMessage)
// Arg list:
// sArgDeref[0]: Msg number
// sArgDeref[1]: wParam
// sArgDeref[2]: lParam
// sArgDeref[3]: Control
// sArgDeref[4]: WinTitle
// sArgDeref[5]: WinText
// sArgDeref[6]: ExcludeTitle
// sArgDeref[7]: ExcludeText
// sArgDeref[8]: Timeout
{
	bool aUseSend = _f_callee_id == FID_SendMessage;
	bool successful = false;

	DETERMINE_TARGET_CONTROL(3);

	UINT msg = ParamIndexToInt(0);
	// Timeout increased from 2000 to 5000 in v1.0.27:
	// jackieku: specify timeout by the parameter.
	UINT timeout = ParamIndexToOptionalInt(8, 5000);

	// Fixed for v1.0.48.04: Make copies of the wParam and lParam variables (if eligible for updating) prior
	// to sending the message in case the message triggers a callback or OnMessage function, which would be
	// likely to change the contents of the mArg array before we're doing using them after the Post/SendMsg.
	// Seems best to do the above EVEN for PostMessage in case it can ever trigger a SendMessage internally
	// (I seem to remember that the OS sometimes converts a PostMessage call into a SendMessage if the
	// origin and destination are the same thread.)
	// v1.0.43.06: If either wParam or lParam contained the address of a variable, update the mLength
	// member after sending the message in case the receiver of the message wrote something to the buffer.
	// This is similar to the way "Str" parameters work in DllCall.
	INT_PTR param[2] = { 0, 0 };
	int i;
	for (i = 1; i < 3; ++i) // Two iterations: wParam and lParam.
	{
		if (ParamIndexIsOmitted(i))
			continue;
		ExprTokenType &this_param = *aParam[i];
		if (this_param.symbol == SYM_VAR)
			this_param.var->ToTokenSkipAddRef(this_param);
		switch (this_param.symbol)
		{
		case SYM_INTEGER:
			param[i-1] = (INT_PTR)this_param.value_int64;
			break;
		case SYM_OBJECT: // Support Buffer-like objects, i.e, objects with a "Ptr" property.
			size_t ptr_property_value;
			GetBufferObjectPtr(aResultToken, this_param.object, ptr_property_value);
			if (aResultToken.Exited())
				return;
			param[i - 1] = ptr_property_value;
			break;
		case SYM_STRING:
			LPTSTR error_marker;
			param[i - 1] = (INT_PTR)istrtoi64(this_param.marker, &error_marker);
			if (!*error_marker) // Valid number or empty string.
				break;
			//else: It's a non-numeric string; maybe the caller forgot the &address-of operator.
			// Note that an empty string would satisfy the check above.
			// Fall through:
		default:
			// SYM_FLOAT: Seems best to treat it as an error rather than truncating the value.
			_f_throw(i == 1 ? ERR_PARAM2_INVALID : ERR_PARAM3_INVALID, ErrorPrototype::Type);
		}
	}

	DWORD_PTR dwResult;
	if (aUseSend)
		successful = SendMessageTimeout(control_window, msg, (WPARAM)param[0], (LPARAM)param[1], SMTO_ABORTIFHUNG, timeout, &dwResult);
	else
		successful = PostMessage(control_window, msg, (WPARAM)param[0], (LPARAM)param[1]);

	if (!successful)
	{
		auto last_error = GetLastError();
		if (aUseSend && last_error == ERROR_TIMEOUT)
			_f_throw(ERR_TIMEOUT, ErrorPrototype::Timeout);
		_f_throw_win32(last_error); // Passing last_error reduces code size due to the implied additional GetLastError() call when omitting this parameter.
	}
	if (aUseSend)
		_f_return_i((__int64)dwResult);
	else
		_f_return_empty;
}



void WinSetRegion(HWND aWnd, LPTSTR aPoints, ResultToken &aResultToken)
{
	if (!*aPoints) // Attempt to restore the window's normal/correct region.
	{
		// Fix for v1.0.31.07: The old method used the following, but apparently it's not the correct
		// way to restore a window's proper/normal region because when such a window is later maximized,
		// it retains its incorrect/smaller region:
		//if (GetWindowRect(aWnd, &rect))
		//{
		//	// Adjust the rect to keep the same size but have its upper-left corner at 0,0:
		//	rect.right -= rect.left;
		//	rect.bottom -= rect.top;
		//	rect.left = 0;
		//	rect.top = 0;
		//	if (hrgn = CreateRectRgnIndirect(&rect)) // Assign
		//	{
		//		// Presumably, the system deletes the former region when upon a successful call to SetWindowRgn().
		//		if (SetWindowRgn(aWnd, hrgn, TRUE))
		//			_f_return_empty;
		//		// Otherwise, get rid of it since it didn't take effect:
		//		DeleteObject(hrgn);
		//	}
		//}
		//// Since above didn't return:
		//return OK;

		// It's undocumented by MSDN, but apparently setting the Window's region to NULL restores it
		// to proper working order:
		if (!SetWindowRgn(aWnd, NULL, TRUE))
			_f_throw_win32();
		_f_return_empty;
	}

	#define MAX_REGION_POINTS 2000  // 2000 requires 16 KB of stack space.
	POINT pt[MAX_REGION_POINTS];
	int pt_count;
	LPTSTR cp;

	// Set defaults prior to parsing options in case any options are absent:
	int width = COORD_UNSPECIFIED;
	int height = COORD_UNSPECIFIED;
	int rr_width = COORD_UNSPECIFIED; // These two are for the rounded-rectangle method.
	int rr_height = COORD_UNSPECIFIED;
	bool use_ellipse = false;

	int fill_mode = ALTERNATE;
	// Concerning polygon regions: ALTERNATE is used by default (somewhat arbitrarily, but it seems to be the
	// more typical default).
	// MSDN: "In general, the modes [ALTERNATE vs. WINDING] differ only in cases where a complex,
	// overlapping polygon must be filled (for example, a five-sided polygon that forms a five-pointed
	// star with a pentagon in the center). In such cases, ALTERNATE mode fills every other enclosed
	// region within the polygon (that is, the points of the star), but WINDING mode fills all regions
	// (that is, the points and the pentagon)."

	for (pt_count = 0, cp = aPoints; *(cp = omit_leading_whitespace(cp));)
	{
		// To allow the MAX to be increased in the future with less chance of breaking existing scripts, consider this an error.
		if (pt_count >= MAX_REGION_POINTS)
			goto arg_error;

		if (isdigit(*cp) || *cp == '-' || *cp == '+') // v1.0.38.02: Recognize leading minus/plus sign so that the X-coord is just as tolerant as the Y.
		{
			// Assume it's a pair of X/Y coordinates.  It's done this way rather than using X and Y
			// as option letters because:
			// 1) The script is more readable when there are multiple coordinates (for polygon).
			// 2) It enforces the fact that each X must have a Y and that X must always come before Y
			//    (which simplifies and reduces the size of the code).
			pt[pt_count].x = ATOI(cp);
			// For the delimiter, dash is more readable than pipe, even though it overlaps with "minus sign".
			// "x" is not used to avoid detecting "x" inside hex numbers.
			#define REGION_DELIMITER '-'
			if (   !(cp = _tcschr(cp + 1, REGION_DELIMITER))   ) // v1.0.38.02: cp + 1 to omit any leading minus sign.
				goto arg_error;
			pt[pt_count].y = ATOI(++cp);  // Increment cp by only 1 to support negative Y-coord.
			++pt_count; // Move on to the next element of the pt array.
		}
		else
		{
			++cp;
			switch(_totupper(cp[-1]))
			{
			case 'E':
				use_ellipse = true;
				break;
			case 'R':
				if (!*cp || *cp == ' ') // Use 30x30 default.
				{
					rr_width = 30;
					rr_height = 30;
				}
				else
				{
					rr_width = ATOI(cp);
					if (cp = _tcschr(cp, REGION_DELIMITER)) // Assign
						rr_height = ATOI(++cp);
					else // Avoid problems with going beyond the end of the string.
						goto arg_error;
				}
				break;
			case 'W':
				if (!_tcsnicmp(cp, _T("ind"), 3)) // [W]ind.
					fill_mode = WINDING;
				else
					width = ATOI(cp);
				break;
			case 'H':
				height = ATOI(cp);
				break;
			default: // For simplicity and to reserve other letters for future use, unknown options result in failure.
				goto arg_error;
			} // switch()
		} // else

		if (   !(cp = _tcschr(cp, ' '))   ) // No more items.
			break;
	}

	if (!pt_count)
		goto arg_error;

	bool width_and_height_were_both_specified = !(width == COORD_UNSPECIFIED || height == COORD_UNSPECIFIED);
	if (width_and_height_were_both_specified)
	{
		width += pt[0].x;   // Make width become the right side of the rect.
		height += pt[0].y;  // Make height become the bottom.
	}

	HRGN hrgn;
	if (use_ellipse) // Ellipse.
		hrgn = width_and_height_were_both_specified ? CreateEllipticRgn(pt[0].x, pt[0].y, width, height) : NULL;
	else if (rr_width != COORD_UNSPECIFIED) // Rounded rectangle.
		hrgn = width_and_height_were_both_specified ? CreateRoundRectRgn(pt[0].x, pt[0].y, width, height, rr_width, rr_height) : NULL;
	else if (width_and_height_were_both_specified) // Rectangle.
		hrgn = CreateRectRgn(pt[0].x, pt[0].y, width, height);
	else // Polygon
		hrgn = CreatePolygonRgn(pt, pt_count, fill_mode);
	if (!hrgn)
		goto error;
	// Since above didn't return, hrgn is now a non-NULL region ready to be assigned to the window.

	// Presumably, the system deletes the window's former region upon a successful call to SetWindowRgn():
	if (!SetWindowRgn(aWnd, hrgn, TRUE))
	{
		DeleteObject(hrgn);
		goto error;
	}
	//else don't delete hrgn since the system has taken ownership of it.

	// Since above didn't return, it's a success.
	_f_return_empty;

arg_error:
	_f_throw_value(ERR_PARAM1_INVALID);

error:
	_f_throw_win32();
}



BIF_DECL(BIF_WinSet)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam + 1, aParamCount - 1))
		return;

	int value;
	BOOL success = FALSE;
	DWORD exstyle;
	BuiltInFunctionID cmd = _f_callee_id;

	LPTSTR aValue;
	ToggleValueType toggle = TOGGLE_INVALID;
	if (cmd == FID_WinSetAlwaysOnTop || cmd == FID_WinSetEnabled)
	{
		toggle = ParamIndexToToggleValue(0);
		if (toggle == TOGGLE_INVALID)
			_f_throw_param(0);
	}
	else
		aValue = ParamIndexToString(0, _f_number_buf);

	switch (cmd)
	{
	case FID_WinSetAlwaysOnTop:
	{
		HWND topmost_or_not;
		switch (toggle)
		{
		case TOGGLED_ON: topmost_or_not = HWND_TOPMOST; break;
		case TOGGLED_OFF: topmost_or_not = HWND_NOTOPMOST; break;
		case TOGGLE:
			exstyle = GetWindowLong(target_window, GWL_EXSTYLE);
			topmost_or_not = (exstyle & WS_EX_TOPMOST) ? HWND_NOTOPMOST : HWND_TOPMOST;
			break;
		}
		// SetWindowLong() didn't seem to work, at least not on some windows.  But this does.
		// As of v1.0.25.14, SWP_NOACTIVATE is also specified, though its absence does not actually
		// seem to activate the window, at least on XP (perhaps due to anti-focus-stealing measure
		// in Win98/2000 and beyond).  Or perhaps its something to do with the presence of
		// topmost_or_not (HWND_TOPMOST/HWND_NOTOPMOST), which might always avoid activating the
		// window.
		success = SetWindowPos(target_window, topmost_or_not, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
		break;
	}

	case FID_WinSetTransparent:
	case FID_WinSetTransColor:
	{
		// IMPORTANT (when considering future enhancements to these commands): Unlike
		// SetLayeredWindowAttributes(), which works on Windows 2000, GetLayeredWindowAttributes()
		// is supported only on XP or later.

		// It appears that turning on WS_EX_LAYERED in an attempt to retrieve the window's
		// former transparency setting does not work.  The OS probably does not store the
		// former transparency level (i.e. it forgets it the moment the WS_EX_LAYERED exstyle
		// is turned off).  This is true even if the following are done after the SetWindowLong():
		//MySetLayeredWindowAttributes(target_window, 0, 0, 0)
		// or:
		//if (MyGetLayeredWindowAttributes(target_window, &color, &alpha, &flags))
		//	MySetLayeredWindowAttributes(target_window, color, alpha, flags);
		// The above is why there is currently no "on" or "toggle" sub-command, just "Off".

		// Since the color of an HBRUSH can't be easily determined (since it can be a pattern and
		// since there seem to be no easy API calls to discover the colors of pixels in an HBRUSH),
		// the following is not yet implemented: Use window's own class background color (via
		// GetClassLong) if aValue is entirely blank.

		exstyle = GetWindowLong(target_window, GWL_EXSTYLE);
		if (!_tcsicmp(aValue, _T("Off")) || !*aValue)
			// One user reported that turning off the attribute helps window's scrolling performance.
			success = SetWindowLong(target_window, GWL_EXSTYLE, exstyle & ~WS_EX_LAYERED);
		else
		{
			if (cmd == FID_WinSetTransparent)
			{
				if (!IsNumeric(aValue, FALSE, FALSE))
					_f_throw_param(0);

				// Update to the below for v1.0.23: WS_EX_LAYERED can now be removed via the above:
				// NOTE: It seems best never to remove the WS_EX_LAYERED attribute, even if the value is 255
				// (non-transparent), since the window might have had that attribute previously and may need
				// it to function properly.  For example, an app may support making its own windows transparent
				// but might not expect to have to turn WS_EX_LAYERED back on if we turned it off.  One drawback
				// of this is a quote from somewhere that might or might not be accurate: "To make this window
				// completely opaque again, remove the WS_EX_LAYERED bit by calling SetWindowLong and then ask
				// the window to repaint. Removing the bit is desired to let the system know that it can free up
				// some memory associated with layering and redirection."
				value = ATOI(aValue);
				// A little debatable, but this behavior seems best, at least in some cases:
				if (value < 0)
					value = 0;
				else if (value > 255)
					value = 255;
				SetWindowLong(target_window, GWL_EXSTYLE, exstyle | WS_EX_LAYERED);
				success = SetLayeredWindowAttributes(target_window, 0, value, LWA_ALPHA);
			}
			else // cmd == FID_WinSetTransColor
			{
				// The reason WINSET_TRANSCOLOR accepts both the color and an optional transparency settings
				// is that calling SetLayeredWindowAttributes() with only the LWA_COLORKEY flag causes the
				// window to lose its current transparency setting in favor of the transparent color.  This
				// is true even though the LWA_ALPHA flag was not specified, which seems odd and is a little
				// disappointing, but that's the way it is on XP at least.
				TCHAR aValue_copy[256];
				tcslcpy(aValue_copy, aValue, _countof(aValue_copy)); // Make a modifiable copy.
				LPTSTR space_pos = StrChrAny(aValue_copy, _T(" \t")); // Space or tab.
				if (space_pos)
				{
					*space_pos = '\0';
					++space_pos;  // Point it to the second substring.
				}
				COLORREF color = ColorNameToBGR(aValue_copy);
				if (color == CLR_NONE) // A matching color name was not found, so assume it's in hex format.
					// It seems _tcstol() automatically handles the optional leading "0x" if present:
					color = rgb_to_bgr(_tcstol(aValue_copy, NULL, 16));
				DWORD flags;
				if (   space_pos && *(space_pos = omit_leading_whitespace(space_pos))   ) // Relies on short-circuit boolean.
				{
					value = ATOI(space_pos);  // To keep it simple, don't bother with 0 to 255 range validation in this case.
					flags = LWA_COLORKEY|LWA_ALPHA;  // i.e. set both the trans-color and the transparency level.
				}
				else // No translucency value is present, only a trans-color.
				{
					value = 0;  // Init to avoid possible compiler warning.
					flags = LWA_COLORKEY;
				}
				SetWindowLong(target_window, GWL_EXSTYLE, exstyle | WS_EX_LAYERED);
				success = SetLayeredWindowAttributes(target_window, color, value, flags);
			}
		}
		break;
	}

	case FID_WinSetStyle:
	case FID_WinSetExStyle:
	{
		if (!*aValue)
		{
			// Seems best not to treat an explicit blank as zero.
			_f_throw_value(ERR_PARAM1_MUST_NOT_BE_BLANK);
		}
		int style_index = (cmd == FID_WinSetStyle) ? GWL_STYLE : GWL_EXSTYLE;
		DWORD new_style, orig_style = GetWindowLong(target_window, style_index);
		if (!_tcschr(_T("+-^"), *aValue))
			new_style = ATOU(aValue); // No prefix, so this new style will entirely replace the current style.
		else
		{
			++aValue; // Won't work combined with next line, due to next line being a macro that uses the arg twice.
			DWORD style_change = ATOU(aValue);
			// +/-/^ are used instead of |&^ because the latter is confusing, namely that
			// "&" really means &=~style, etc.
			switch(aValue[-1])
			{
			case '+': new_style = orig_style | style_change; break;
			case '-': new_style = orig_style & ~style_change; break;
			case '^': new_style = orig_style ^ style_change; break;
			}
		}
		SetLastError(0); // Prior to SetWindowLong(), as recommended by MSDN.
		if (SetWindowLong(target_window, style_index, new_style) || !GetLastError()) // This is the precise way to detect success according to MSDN.
		{
			// Even if it indicated success, sometimes it failed anyway.  Find out for sure:
			if (new_style == orig_style || GetWindowLong(target_window, style_index) != orig_style) // Even a partial change counts as a success.
			{
				// SetWindowPos is also necessary, otherwise the frame thickness entirely around the window
				// does not get updated (just parts of it):
				SetWindowPos(target_window, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
				// Since SetWindowPos() probably doesn't know that the style changed, below is probably necessary
				// too, at least in some cases:
				InvalidateRect(target_window, NULL, TRUE); // Quite a few styles require this to become visibly manifest.
				success = TRUE;
			}
		}
		break;
	}

	case FID_WinSetEnabled:
		if (toggle == TOGGLE)
			toggle = IsWindowEnabled(target_window) ? TOGGLED_OFF : TOGGLED_ON;
		EnableWindow(target_window, toggle == TOGGLED_ON); // Return value is based on previous state, not success/failure.
		success = bool(IsWindowEnabled(target_window)) == (toggle == TOGGLED_ON);
		break;

	case FID_WinSetRegion:
		WinSetRegion(target_window, aValue, aResultToken);
		return;

	} // switch()
	if (!success)
		_f_throw_win32();
	_f_return_empty;
}



BIF_DECL(BIF_WinRedraw)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam, aParamCount))
		return;
	// Seems best to always have the last param be TRUE. Also, it seems best not to call
	// UpdateWindow(), which forces the window to immediately process a WM_PAINT message,
	// since that might not be desirable.  Some other methods of getting a window to redraw:
	// SendMessage(mHwnd, WM_NCPAINT, 1, 0);
	// RedrawWindow(mHwnd, NULL, NULL, RDW_INVALIDATE|RDW_FRAME|RDW_UPDATENOW);
	// SetWindowPos(mHwnd, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
	// GetClientRect(mHwnd, &client_rect); InvalidateRect(mHwnd, &client_rect, TRUE);
	InvalidateRect(target_window, NULL, TRUE);
	_f_return_empty;
}



BIF_DECL(BIF_WinMoveTopBottom)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam, aParamCount))
		return;
	HWND mode = _f_callee_id == FID_WinMoveBottom ? HWND_BOTTOM : HWND_TOP;
	// Note: SWP_NOACTIVATE must be specified otherwise the target window often fails to move:
	if (!SetWindowPos(target_window, mode, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE))
		_f_throw_win32();
	_f_return_empty;
}



BIF_DECL(BIF_WinSetTitle)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam + 1, aParamCount - 1))
		return;
	if (!SetWindowText(target_window, ParamIndexToString(0, _f_number_buf)))
		_f_throw_win32();
	_f_return_empty;
}



BIF_DECL(BIF_WinGetTitle)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam, aParamCount))
		return;

	int space_needed = target_window ? GetWindowTextLength(target_window) + 1 : 1; // 1 for terminator.
	if (!TokenSetResult(aResultToken, NULL, space_needed - 1))
		return;  // It already displayed the error.
	aResultToken.symbol = SYM_STRING;
	// Update length using the actual length, rather than the estimate provided by GetWindowTextLength():
	aResultToken.marker_length = GetWindowText(target_window, aResultToken.marker, space_needed);
	if (!aResultToken.marker_length)
		// There was no text to get or GetWindowText() failed.
		*aResultToken.marker = '\0';
}



BIF_DECL(BIF_WinGetClass)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam, aParamCount))
		return;
	TCHAR class_name[WINDOW_CLASS_SIZE];
	if (!GetClassName(target_window, class_name, _countof(class_name)))
		_f_throw_win32();
	_f_return(class_name);
}



void WinGetList(ResultToken &aResultToken, BuiltInFunctionID aCmd, LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
// Helper function for WinGet() to avoid having a WindowSearch object on its stack (since that object
// normally isn't needed).
{
	WindowSearch ws;
	ws.mFindLastMatch = true; // Must set mFindLastMatch to get all matches rather than just the first.
	if (aCmd == FID_WinGetList)
		if (  !(ws.mArray = Array::Create())  )
			_f_throw_oom;
	// Check if we were asked to "list" or count the active window:
	if (USE_FOREGROUND_WINDOW(aTitle, aText, aExcludeTitle, aExcludeText))
	{
		HWND target_window; // Set by the macro below.
		SET_TARGET_TO_ALLOWABLE_FOREGROUND(g->DetectHiddenWindows)
		if (aCmd == FID_WinGetCount)
		{
			_f_return_i(target_window != NULL); // i.e. 1 if a window was found.
		}
		// Otherwise, it's FID_WinGetList:
		if (target_window)
			// Since the target window has been determined, we know that it is the only window
			// to be put into the array:
			ws.mArray->Append((__int64)(size_t)target_window);
		// Otherwise, return an empty array.
		_f_return(ws.mArray);
	}
	// Enumerate all windows which match the criteria:
	// If aTitle is ahk_id nnnn, the Enum() below will be inefficient.  However, ahk_id is almost unheard of
	// in this context because it makes little sense, so no extra code is added to make that case efficient.
	if (ws.SetCriteria(*g, aTitle, aText, aExcludeTitle, aExcludeText)) // These criteria allow the possibility of a match.
		EnumWindows(EnumParentFind, (LPARAM)&ws);
	//else leave ws.mFoundCount set to zero (by the constructor).
	if (aCmd == FID_WinGetList)
		// Return the array even if it is empty:
		_f_return(ws.mArray);
	else // FID_WinGetCount
		_f_return_i(ws.mFoundCount);
}

void WinGetControlList(ResultToken &aResultToken, HWND aTargetWindow, bool aFetchHWNDs); // Forward declaration.



BIF_DECL(BIF_WinGet)
{
	_f_param_string_opt(aTitle, 0);
	_f_param_string_opt(aText, 1);
	_f_param_string_opt(aExcludeTitle, 2);
	_f_param_string_opt(aExcludeText, 3);

	BuiltInFunctionID cmd = _f_callee_id;

	if (cmd == FID_WinGetList || cmd == FID_WinGetCount)
	{
		WinGetList(aResultToken, cmd, aTitle, aText, aExcludeTitle, aExcludeText);
		return;
	}
	// Not List or Count, so it's a function that operates on a single window.
	HWND target_window = NULL;
	if (aParamCount > 0)
		switch (DetermineTargetHwnd(target_window, aResultToken, *aParam[0]))
		{
		case FAIL: return;
		case OK:
			if (!target_window)
				_f_throw(ERR_NO_WINDOW, ErrorPrototype::Target);
		}
	if (!target_window)
		target_window = WinExist(*g, aTitle, aText, aExcludeTitle, aExcludeText, cmd == FID_WinGetIDLast);
	if (!target_window)
		_f_throw(ERR_NO_WINDOW, ErrorPrototype::Target);

	switch(cmd)
	{
	case FID_WinGetID:
	case FID_WinGetIDLast:
		_f_return((size_t)target_window);

	case FID_WinGetPID:
	case FID_WinGetProcessName:
	case FID_WinGetProcessPath:
		DWORD pid;
		GetWindowThreadProcessId(target_window, &pid);
		if (cmd == FID_WinGetPID)
		{
			_f_return_i(pid);
		}
		// Otherwise, get the full path and name of the executable that owns this window.
		TCHAR process_name[MAX_PATH];
		GetProcessName(pid, process_name, _countof(process_name), cmd == FID_WinGetProcessName);
		_f_return(process_name);

	case FID_WinGetMinMax:
		// Testing shows that it's not possible for a minimized window to also be maximized (under
		// the theory that upon restoration, it *would* be maximized.  This is unfortunate if there
		// is no other way to determine what the restoration size and maximized state will be for a
		// minimized window.
		_f_return_i(IsZoomed(target_window) ? 1 : (IsIconic(target_window) ? -1 : 0));

	case FID_WinGetControls:
	case FID_WinGetControlsHwnd:
		WinGetControlList(aResultToken, target_window, cmd == FID_WinGetControlsHwnd);
		return;

	case FID_WinGetStyle:
	case FID_WinGetExStyle:
		_f_return_i(GetWindowLong(target_window, cmd == FID_WinGetStyle ? GWL_STYLE : GWL_EXSTYLE));

	case FID_WinGetTransparent:
	case FID_WinGetTransColor:
		COLORREF color;
		BYTE alpha;
		DWORD flags;
		if (!(GetLayeredWindowAttributes(target_window, &color, &alpha, &flags)))
			break;
		if (cmd == FID_WinGetTransparent)
		{
			if (flags & LWA_ALPHA)
			{
				_f_return_i(alpha);
			}
		}
		else // FID_WinGetTransColor
		{
			if (flags & LWA_COLORKEY)
			{
				// Store in hex format to aid in debugging scripts.  Also, the color is always
				// stored in RGB format, since that's what WinSet uses:
				LPTSTR result = _f_retval_buf;
				_stprintf(result, _T("0x%06X"), bgr_to_rgb(color));
				_f_return_p(result);
			}
			// Otherwise, this window does not have a transparent color (or it's not accessible to us,
			// perhaps for reasons described at MSDN GetLayeredWindowAttributes()).
		}
	}
	_f_return_empty;
}



void WinGetControlList(ResultToken &aResultToken, HWND aTargetWindow, bool aFetchHWNDs)
// Caller must ensure that aTargetWindow is non-NULL and valid.
// Every control is fetched rather than just a list of distinct class names (possibly with a
// second script array containing the quantity of each class) because it's conceivable that the
// z-order of the controls will be useful information to some script authors.
{
	control_list_type cl; // A big struct containing room to store class names and counts for each.
	if (  !(cl.target_array = Array::Create())  )
		_f_throw_oom;
	CL_INIT_CONTROL_LIST(cl)
	cl.fetch_hwnds = aFetchHWNDs;
	EnumChildWindows(aTargetWindow, EnumChildGetControlList, (LPARAM)&cl);
	_f_return(cl.target_array);
}



BOOL CALLBACK EnumChildGetControlList(HWND aWnd, LPARAM lParam)
{
	control_list_type &cl = *(control_list_type *)lParam;  // For performance and convenience.

	// cl.fetch_hwnds==true is a new mode in v1.0.43.06+ to help performance of AHK Window Info and other
	// scripts that want to operate directly on the HWNDs.
	if (cl.fetch_hwnds)
	{
		cl.target_array->Append((__int64)(size_t)aWnd);
	}
	else // The mode that fetches ClassNN vs. HWND.
	{
		TCHAR line[WINDOW_CLASS_SIZE + 5];  // +5 to allow room for the sequence number to be appended later below.
		int line_length;

		// Note: IsWindowVisible(aWnd) is not checked because although Window Spy does not reveal
		// hidden controls if the mouse happens to be hovering over one, it does include them in its
		// sequence numbering (which is a relieve, since results are probably much more consistent
		// then, esp. for apps that hide and unhide controls in response to actions on other controls).
		if (  !(line_length = GetClassName(aWnd, line, WINDOW_CLASS_SIZE))   ) // Don't include the +5 extra size since that is reserved for seq. number.
			return TRUE; // Probably very rare. Continue enumeration since Window Spy doesn't even check for failure.
		// It has been verified that GetClassName()'s returned length does not count the terminator.

		// Check if this class already exists in the class array:
		int class_index;
		for (class_index = 0; class_index < cl.total_classes; ++class_index)
			if (!_tcsicmp(cl.class_name[class_index], line)) // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales.
				break;
		if (class_index < cl.total_classes) // Match found.
		{
			++cl.class_count[class_index]; // Increment the number of controls of this class that have been found so far.
			if (cl.class_count[class_index] > 99999) // Sanity check; prevents buffer overflow or number truncation in "line".
				return TRUE;  // Continue the enumeration.
		}
		else // No match found, so create new entry if there's room.
		{
			if (cl.total_classes == CL_MAX_CLASSES // No pointers left.
				|| CL_CLASS_BUF_SIZE - (cl.buf_free_spot - cl.class_buf) - 1 < line_length) // Insuff. room in buf.
				return TRUE; // Very rare. Continue the enumeration so that class names already found can be collected.
			// Otherwise:
			cl.class_name[class_index] = cl.buf_free_spot;  // Set this pointer to its place in the buffer.
			_tcscpy(cl.class_name[class_index], line); // Copy the string into this place.
			cl.buf_free_spot += line_length + 1;  // +1 because every string in the buf needs its own terminator.
			cl.class_count[class_index] = 1;  // Indicate that the quantity of this class so far is 1.
			++cl.total_classes;
		}

		_itot(cl.class_count[class_index], line + line_length, 10); // Append the seq. number to line.
		
		cl.target_array->Append(line); // Append class name+seq. number to the array.
	}

	return TRUE; // Continue enumeration through all the windows.
}



BIF_DECL(BIF_WinGetText)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam, aParamCount))
		return;

	length_and_buf_type sab;
	sab.buf = NULL; // Tell it just to calculate the length this time around.
	sab.total_length = 0; // Init
	sab.capacity = 0;     //
	EnumChildWindows(target_window, EnumChildGetText, (LPARAM)&sab);

	if (!sab.total_length) // No text in window.
		_f_return_empty;

	if (!TokenSetResult(aResultToken, NULL, sab.total_length))
		return;  // It already displayed the error.
	aResultToken.symbol = SYM_STRING;

	// Fetch the text directly into the var.  Also set the length explicitly
	// in case actual size written was different than the estimated size (since
	// GetWindowTextLength() can return more space that will actually be required
	// in certain circumstances, see MSDN):
	sab.buf = aResultToken.marker;
	// Note: The capacity member below exists because granted capacity might be a little larger than we asked for,
	// which allows the actual text fetched to be larger than the length estimate retrieved by the first pass
	// (which generally shouldn't happen since MSDN docs say that the actual length can be less, but never greater,
	// than the estimate length):
	sab.capacity = sab.total_length + 1; // Capacity includes the zero terminator, i.e. it's the size of the memory area.
	sab.total_length = 0; // Reinitialize.
	EnumChildWindows(target_window, EnumChildGetText, (LPARAM)&sab); // Get the text.

	// Length is set explicitly below in case it wound up being smaller than expected/estimated.
	// MSDN says that can happen generally, and also specifically because: "ANSI applications may have
	// the string in the buffer reduced in size (to a minimum of half that of the wParam value) due to
	// conversion from ANSI to Unicode."
	aResultToken.marker_length = sab.total_length;
	if (!sab.total_length)
	{
		// Something went wrong, so make sure we set to empty string.
		*sab.buf = '\0';
		_f_throw(ERR_FAILED);
	}
}



BOOL CALLBACK EnumChildGetText(HWND aWnd, LPARAM lParam)
{
	if (!g->DetectHiddenText && !IsWindowVisible(aWnd))
		return TRUE;  // This child/control is hidden and user doesn't want it considered, so skip it.
	length_and_buf_type &lab = *(length_and_buf_type *)lParam;  // For performance and convenience.
	int length;
	if (lab.buf)
		length = GetWindowTextTimeout(aWnd, lab.buf + lab.total_length
			, (int)(lab.capacity - lab.total_length)); // Not +1.  Verified correct because WM_GETTEXT accepts size of buffer, not its length.
	else
		length = GetWindowTextTimeout(aWnd);
	lab.total_length += length;
	if (length)
	{
		if (lab.buf)
		{
			if (lab.capacity - lab.total_length > 2) // Must be >2 due to zero terminator.
			{
				_tcscpy(lab.buf + lab.total_length, _T("\r\n")); // Something to delimit each control's text.
				lab.total_length += 2;
			}
			// else don't increment total_length
		}
		else
			lab.total_length += 2; // Since buf is NULL, accumulate the size that *would* be needed.
	}
	return TRUE; // Continue enumeration through all the child windows of this parent.
}



BIF_DECL(BIF_WinGetPos)
{
	HWND target_window;
	if (!DetermineTargetWindow(target_window, aResultToken, aParam + 4, aParamCount - 4))
		return;
	RECT rect;
	if (_f_callee_id == FID_WinGetPos)
	{
		GetWindowRect(target_window, &rect);
		rect.right -= rect.left; // Convert right to width.
		rect.bottom -= rect.top; // Convert bottom to height.
	}
	else // FID_WinGetClientPos
	{
		GetClientRect(target_window, &rect); // Get client pos relative to client (position is always 0,0).
		// Since the position is always 0,0, right,bottom are already equivalent to width,height.
		MapWindowPoints(target_window, NULL, (LPPOINT)&rect, 1); // Convert position to screen coordinates.
	}

	for (int i = 0; i < 4; ++i)
	{
		Var *var = ParamIndexToOutputVar(i);
		if (!var)
			continue;
		var->Assign(((int *)(&rect))[i]); // Always succeeds.
	}
	_f_return_empty;
}



DECLSPEC_NOINLINE // Lexikos: noinline saves ~300 bytes.  Originally the duplicated code prevented inlining.
HWND Line::DetermineTargetWindow(LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText)
{
	// Lexikos: Not sure why these checks were duplicated here and in WinExist(),
	// so they're left here for reference:
	//HWND target_window; // A variable of this name is used by the macros below.
	//IF_USE_FOREGROUND_WINDOW(g->DetectHiddenWindows, aTitle, aText, aExcludeTitle, aExcludeText)
	//else if (*aTitle || *aText || *aExcludeTitle || *aExcludeText)
	//	target_window = WinExist(*g, aTitle, aText, aExcludeTitle, aExcludeText);
	//else // Use the "last found" window.
	//	target_window = GetValidLastUsedWindow(*g);
	return WinExist(*g, aTitle, aText, aExcludeTitle, aExcludeText);
}


ResultType DetermineTargetHwnd(HWND &aWindow, ResultToken &aResultToken, ExprTokenType &aToken)
{
	__int64 n = NULL;
	if (IObject *obj = TokenToObject(aToken))
	{
		if (!GetObjectIntProperty(obj, _T("Hwnd"), n, aResultToken))
			return FAIL;
	}
	else if (TokenIsPureNumeric(aToken) == PURE_INTEGER)
		n = TokenToInt64(aToken);
	else
		return CONDITION_FALSE;
	aWindow = (HWND)(UINT_PTR)n;
	// Callers expect the return value to be either a valid HWND or 0:
	if (!IsWindow(aWindow))
		aWindow = 0;
	return OK;
}


ResultType DetermineTargetWindow(HWND &aWindow, ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, int aNonWinParamCount)
{
	if (aParamCount > 0)
	{
		auto result = DetermineTargetHwnd(aWindow, aResultToken, *aParam[0]);
		if (result != CONDITION_FALSE)
		{
			if (result == OK && !aWindow)
				return aResultToken.Error(ERR_NO_WINDOW, ErrorPrototype::Target);
			return result;
		}
	}
	TCHAR number_buf[4][MAX_NUMBER_SIZE];
	LPTSTR param[4];
	for (int i = 0, j = 0; i < 4; ++i, ++j)
	{
		if (i == 2) j += aNonWinParamCount;
		param[i] = j < aParamCount ? TokenToString(*aParam[j], number_buf[i]) : _T("");
	}
	aWindow = Line::DetermineTargetWindow(param[0], param[1], param[2], param[3]);
	if (aWindow)
		return OK;
	return aResultToken.Error(ERR_NO_WINDOW, param[0], ErrorPrototype::Target);
}


ResultType DetermineTargetControl(HWND &aControl, HWND &aWindow, ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount
	, int aNonWinParamCount, bool aThrowIfNotFound)
{
	aWindow = aControl = nullptr;
	// Only functions which can operate on top-level windows allow Control to be
	// omitted (and a select few other functions with more optional parameters).
	// This replaces the old "ahk_parent" string used with ControlSend, but is
	// also used by SendMessage.
	LPTSTR control_spec = nullptr;
	if (!ParamIndexIsOmitted(0))
	{
		switch (DetermineTargetHwnd(aWindow, aResultToken, *aParam[0]))
		{
		case OK:
			aControl = aWindow;
			if (!aControl)
				return aResultToken.Error(ERR_NO_CONTROL, ErrorPrototype::Target);
			return OK;
		case FAIL:
			return FAIL;
		}
		// Since above didn't return, it wasn't a pure Integer or object {Hwnd}.
		control_spec = ParamIndexToString(0, _f_number_buf);
	}
	if (!DetermineTargetWindow(aWindow, aResultToken, aParam + 1, aParamCount - 1, aNonWinParamCount))
		return FAIL;
	aControl = control_spec ? ControlExist(aWindow, control_spec) : aWindow;
	if (!aControl && aThrowIfNotFound)
		return aResultToken.Error(ERR_NO_CONTROL, ErrorPrototype::Target);
	return OK;
}



BIF_DECL(BIF_WinExistActive)
{
	bool hwnd_specified = false;
	HWND hwnd;
	if (!ParamIndexIsOmitted(0))
	{
		switch (DetermineTargetHwnd(hwnd, aResultToken, *aParam[0]))
		{
		case FAIL:
			return;
		case OK:
			hwnd_specified = true;
			// DetermineTargetHwnd() already called IsWindow() to verify hwnd.
			// g->DetectHiddenWindows is intentionally ignored for these cases.
			if (_f_callee_id == FID_WinActive && GetForegroundWindow() != hwnd)
				hwnd = 0;
			if (hwnd)
				g->hWndLastUsed = hwnd;
			break;
		}
	}

	if (!hwnd_specified) // Do not call WinExist()/WinActive() even if the specified hwnd was 0.
	{
		TCHAR *param[4], param_buf[4][MAX_NUMBER_SIZE];
		for (int j = 0; j < 4; ++j) // For each formal parameter, including optional ones.
			param[j] = ParamIndexToOptionalString(j, param_buf[j]);

		hwnd = _f_callee_id == FID_WinExist
			? WinExist(*g, param[0], param[1], param[2], param[3], false, true)
			: WinActive(*g, param[0], param[1], param[2], param[3], true);
	}

	_f_return_i((size_t)hwnd);
}



bif_impl void WinMinimizeAll()
{
	PostMessage(FindWindow(_T("Shell_TrayWnd"), NULL), WM_COMMAND, 419, 0);
	DoWinDelay;
}



bif_impl void WinMinimizeAllUndo()
{
	PostMessage(FindWindow(_T("Shell_TrayWnd"), NULL), WM_COMMAND, 416, 0);
	DoWinDelay;
}
