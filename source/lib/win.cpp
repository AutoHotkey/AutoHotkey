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
#include "shlwapi.h" // StrCmpLogicalW

#include <mmdeviceapi.h> // for SoundSet/SoundGet.
#include <endpointvolume.h> // for SoundSet/SoundGet.
#include <functiondiscoverykeys.h>

#define PCRE_STATIC             // For RegEx. PCRE_STATIC tells PCRE to declare its functions for normal, static
#include "lib_pcre/pcre/pcre.h" // linkage rather than as functions inside an external DLL.

#include "script_func_impl.h"



//////////////////////
// Window Functions //
//////////////////////


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



BIF_DECL(BIF_GroupActivate)
{
	LPTSTR aGroup = ParamIndexToOptionalString(0, _f_number_buf);
	WinGroup *group;
	if (   !(group = g_script.FindGroup(aGroup, true))   ) // Last parameter -> create-if-not-found.
		_f_return_FAIL;  // It already displayed the error for us.
	
	LPTSTR aMode = ParamIndexToOptionalString(1, _f_number_buf);
	bool reverse = false;
	if (!_tcsicmp(aMode, _T("R")))
		reverse = true;
	else if (*aMode)
		_f_throw_param(1);

	HWND activated;
	if (!group->Activate(reverse, activated))
		_f_return_FAIL;
	_f_return_i((UINT_PTR)activated);
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
	if (col_count > -1 && col_option && (requested_col < 0 || requested_col >= col_count)) // Specified column does not exist.
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



//////////////////////
// Window Functions //
//////////////////////


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
	if (g_MsgMonitor.Count() // Count is checked here to avoid function-call overhead.
		&& (!g->CalledByIsDialogMessageOrDispatch || g->CalledByIsDialogMessageOrDispatchMsg != iMsg) // v1.0.44.11: If called by IsDialog or Dispatch but they changed the message number, check if the script is monitoring that new number.
		&& MsgMonitor(hWnd, iMsg, wParam, lParam, NULL, msg_reply))
		return msg_reply; // MsgMonitor has returned "true", indicating that this message should be omitted from further processing.
	g->CalledByIsDialogMessageOrDispatch = false; // v1.0.40.01.

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
			g_script.mTrayMenu->Display(false);
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
		if (!g_MenuIsVisible) // See comments in similar code in GuiWindowProc().
			g_MenuIsVisible = MENU_TYPE_BAR;
		break;
	case WM_EXITMENULOOP:
		g_MenuIsVisible = MENU_TYPE_NONE; // See comments in similar code in GuiWindowProc().
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
		Line::ToggleSuspendState();
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

UserFunc* Script::CreateHotFunc()
{
	// Should only be called during load time.
	// Creates a new function for hotkeys and hotstrings.
	// Caller should abort loading if this function returns nullptr.
	
	if (mUnusedHotFunc)
	{
		auto tmp = mLastHotFunc = g->CurrentFunc = mUnusedHotFunc;
		mUnusedHotFunc = nullptr;
		mHotFuncs.mCount++;			// DefineFunc "removed" this func previously.
		ASSERT(mHotFuncs.mItem[mHotFuncs.mCount - 1] == tmp);
		return tmp;
	}
	
	static LPCTSTR sName = _T("<Hotkey>");
	auto func = new UserFunc(sName);
	
	g->CurrentFunc = func; // Must do this before calling AddVar

	// Add one parameter to hold the name of the hotkey/hotstring when triggered:
	func->mParam = SimpleHeap::Alloc<FuncParam>();
	if ( !(func->mParam[0].var = AddVar(_T("ThisHotkey"), 10, &func->mVars, 0, VAR_DECLARE_LOCAL | VAR_LOCAL_FUNCPARAM)) )
		return nullptr;

	func->mParam[0].default_type = PARAM_DEFAULT_NONE;
	func->mParam[0].is_byref = false;
	func->mParamCount = 1;
	func->mMinParams = 1;
	func->mIsFuncExpression = false;
	
	mLastHotFunc = func;
	mHotFuncs.Insert(func, mHotFuncs.mCount);
	return func;
}



////////////
// MsgBox //
////////////


ResultType MsgBoxParseOptions(LPTSTR aOptions, int &aType, double &aTimeout, HWND &aOwner)
{
	aType = 0;
	aTimeout = 0;

	//int button_option = 0;
	//int icon_option = 0;

	LPTSTR next_option, option_end;
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
		else if (!_tcsnicmp(option, _T("Default"), 7) && IsNumeric(option + 7, FALSE, FALSE, FALSE))
		{
			int default_button = ATOI(option + 7);
			if (default_button < 1 || default_button > 0xF) // Currently MsgBox can only have 4 buttons, but MB_DEFMASK may allow for up to this many in future.
				goto invalid_option;
			aType = (aType & ~MB_DEFMASK) | ((default_button - 1) << 8); // 1=0, 2=0x100, 3=0x200, 4=0x300
		}
		else if (toupper(*option) == 'T' && IsNumeric(option + 1, FALSE, FALSE, TRUE))
		{
			aTimeout = ATOF(option + 1);
		}
		else if (!_tcsnicmp(option, _T("Owner"), 5) && IsNumeric(option + 5, TRUE, TRUE, FALSE))
		{
			aOwner = (HWND)ATOI64(option + 5); // This should be consistent with the Gui +Owner option.
		}
		else if (IsNumeric(option, FALSE, FALSE, FALSE))
		{
			int other_option = ATOI(option);
			// Clear any conflicting options which were previously set.
			if (other_option & MB_TYPEMASK) aType &= ~MB_TYPEMASK;
			if (other_option & MB_ICONMASK) aType &= ~MB_ICONMASK;
			if (other_option & MB_DEFMASK)  aType &= ~MB_DEFMASK;
			if (other_option & MB_MODEMASK) aType &= ~MB_MODEMASK;
			// All remaining options are bit flags (or not conflicting).
			aType |= other_option;
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


BIF_DECL(BIF_MsgBox)
{
	int result;
	HWND dialog_owner = THREAD_DIALOG_OWNER; // Resolve macro only once to reduce code size.
	// dialog_owner is passed via parameter to avoid internally-displayed MsgBoxes from being
	// affected by script-thread's owner setting.
	_f_param_string_opt_def(aText, 0, nullptr);
	_f_param_string_opt_def(aTitle, 1, nullptr);
	_f_param_string_opt(aOptions, 2);
	int type;
	double timeout;
	if (!MsgBoxParseOptions(aOptions, type, timeout, dialog_owner))
	{
		aResultToken.SetExitResult(FAIL);
		return;
	}
	result = MsgBox(aText, type, aTitle, timeout, dialog_owner);
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
	// v1.1.09.02: If the MsgBox failed due to invalid options, it seems better to display
	// an error dialog than to silently exit the thread:
	if (!result && GetLastError() == ERROR_INVALID_MSGBOX_STYLE)
		_f_throw_param(2);
	// Return a string such as "OK", "Yes" or "No" if possible, or fall back to the integer value.
	if (LPTSTR result_string = MsgBoxResultString(result))
		_f_return_p(result_string);
	else
		_f_return_i(result);
}



//////////////
// InputBox //
//////////////

ResultType InputBoxParseOptions(LPTSTR aOptions, InputBoxType &aInputBox)
{
	LPTSTR next_option, option_end;
	for (next_option = aOptions; *next_option; next_option = omit_leading_whitespace(option_end))
	{
		// Find the end of this option item:
		if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
			option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.

		// Temporarily terminate for simplicity and to reduce ambiguity:
		TCHAR orig_char = *option_end;
		*option_end = '\0';

		// The legacy InputBox command used "Hide", but "Password" seems clearer
		// and better for consistency with the equivalent Edit control option:
		if (!_tcsnicmp(next_option, _T("Password"), 8) && _tcslen(next_option) <= 9)
			aInputBox.password_char = next_option[8] ? next_option[8] : UorA(L'\x25CF', '*');
		else
		{
			// All of the remaining options are single-letter followed by a number:
			TCHAR option_char = ctoupper(*next_option);
			if (!_tcschr(_T("XYWHT"), option_char) // Not a valid option char.
				|| !IsNumeric(next_option + 1 // Or not a valid number.
					, option_char == 'X' || option_char == 'Y' // Only X and Y allow negative numbers.
					, FALSE, option_char == 'T')) // Only Timeout allows floating-point.
			{
				*option_end = orig_char; // Undo the temporary termination.
				return ValueError(ERR_INVALID_OPTION, next_option, FAIL_OR_OK);
			}

			switch (ctoupper(*next_option))
			{
			case 'W': aInputBox.width = DPIScale(ATOI(next_option + 1)); break;
			case 'H': aInputBox.height = DPIScale(ATOI(next_option + 1)); break;
			case 'X': aInputBox.xpos = ATOI(next_option + 1); break;
			case 'Y': aInputBox.ypos = ATOI(next_option + 1); break;
			case 'T': aInputBox.timeout = (DWORD)(ATOF(next_option + 1) * 1000); break;
			}
		}
		
		*option_end = orig_char; // Undo the temporary termination.
	}
	return OK;
}

BIF_DECL(BIF_InputBox)
{
	_f_param_string_opt(aText, 0);
	_f_param_string_opt_def(aTitle, 1, g_script.DefaultDialogTitle());
	_f_param_string_opt(aOptions, 2);
	_f_param_string_opt(aDefault, 3);

	InputBoxType inputbox;
	inputbox.title = aTitle;
	inputbox.text = aText;
	inputbox.default_string = aDefault;
	inputbox.return_string = nullptr;
	// Set defaults:
	inputbox.width = INPUTBOX_DEFAULT;
	inputbox.height = INPUTBOX_DEFAULT;
	inputbox.xpos = INPUTBOX_DEFAULT;
	inputbox.ypos = INPUTBOX_DEFAULT;
	inputbox.password_char = '\0';
	inputbox.timeout = 0;
	// Parse options and override defaults:
	if (!InputBoxParseOptions(aOptions, inputbox))
		_f_return_FAIL; // It already displayed the error.

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
	// result can be -1 or FAIL in case of failure, but since failure of any
	// kind is rare, all kinds are handled the same way, below.

	if (reason && value)
	{
		ExprTokenType argt[] = { _T("Result"), reason, _T("Value"), value };
		ExprTokenType *args[_countof(argt)] = { argt, argt+1, argt+2, argt+3 };
		if (Object *obj = Object::Create(args, _countof(args)))
		{
			free(value);
			_f_return(obj);
		}
	}

	free(value);
	// Since above didn't return, result is -1 (DialogBox somehow failed),
	// result is FAIL (something failed in InputBoxProc), or value is null.
	_f_throw(ERR_INTERNAL_CALL);
}

BIF_DECL(BIF_ToolTip)
{
	int window_index = 0; // Default if param index 3 omitted
	if (!ParamIndexIsOmitted(3))
	{
		Throw_if_Param_NaN(3);
		window_index = ParamIndexToInt(3) - 1;
		if (window_index < 0 || window_index >= MAX_TOOLTIPS)
			_f_throw_value(_T("Max window number is ") MAX_TOOLTIPS_STR _T("."));
	}

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
	TCHAR number_buf[MAX_NUMBER_SIZE];
	auto tip_text = ParamIndexToOptionalString(0, number_buf);
	if (!*tip_text)
	{
		if (tip_hwnd && IsWindow(tip_hwnd))
			DestroyWindow(tip_hwnd);
		g_hWndToolTip[window_index] = NULL;
		_f_return_empty;
	}
	
	bool param1_omitted = ParamIndexIsOmitted(1);
	bool param2_omitted = ParamIndexIsOmitted(2);
	bool one_or_both_coords_unspecified = param1_omitted || param2_omitted;
		 
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
	if (!param1_omitted || !param2_omitted) // Need the offsets.
		CoordToScreen(origin, COORD_MODE_TOOLTIP);

	// This will also convert from relative to screen coordinates if appropriate:
	if (!param1_omitted)
	{
		Throw_if_Param_NaN(1);
		pt.x = ParamIndexToInt(1) + origin.x;
	}
	if (!param2_omitted)
	{
		Throw_if_Param_NaN(2);
		pt.y = ParamIndexToInt(2) + origin.y;
	}
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
	ti.uFlags = TTF_TRACK;
	ti.lpszText = tip_text;
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
			_f_throw_win32();
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
	_f_return((size_t)tip_hwnd);
}

INT_PTR CALLBACK InputBoxProc(HWND hWndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
// MSDN:
// Typically, the dialog box procedure should return TRUE if it processed the message,
// and FALSE if it did not. If the dialog box procedure returns FALSE, the dialog
// manager performs the default dialog operation in response to the message.
{
	// See GuiWindowProc() for details about this first part:
	LRESULT msg_reply;
	if (g_MsgMonitor.Count() // Count is checked here to avoid function-call overhead.
		&& (!g->CalledByIsDialogMessageOrDispatch || g->CalledByIsDialogMessageOrDispatchMsg != uMsg) // v1.0.44.11: If called by IsDialog or Dispatch but they changed the message number, check if the script is monitoring that new number.
		&& MsgMonitor(hWndDlg, uMsg, wParam, lParam, NULL, msg_reply))
		return (BOOL)msg_reply; // MsgMonitor has returned "true", indicating that this message should be omitted from further processing.
	g->CalledByIsDialogMessageOrDispatch = false; // v1.0.40.01.

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
		
		if (*CURR_INPUTBOX.default_string)
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



///////////////////
// Misc internal //
///////////////////

VOID CALLBACK DerefTimeout(HWND hWnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	Line::FreeDerefBufIfLarge(); // It will also kill the timer, if appropriate.
}



////////////////////////
// Misc built-in vars //
////////////////////////

ResultType Script::SetCoordMode(LPTSTR aCommand, LPTSTR aMode)
{
	CoordModeType mode = Line::ConvertCoordMode(aMode);
	CoordModeType shift = Line::ConvertCoordModeCmd(aCommand);
	if (shift == COORD_MODE_INVALID || mode == COORD_MODE_INVALID)
		return ValueError(ERR_INVALID_VALUE, aMode, FAIL_OR_OK);
	g->CoordMode = (g->CoordMode & ~(COORD_MODE_MASK << shift)) | (mode << shift);
	return OK;
}

ResultType Script::SetSendMode(LPTSTR aValue)
{
	g->SendMode = Line::ConvertSendMode(aValue, g->SendMode); // Leave value unchanged if ARG1 is invalid.
	return OK;
}

ResultType Script::SetSendLevel(int aValue, LPTSTR aValueStr)
{
	int sendLevel = aValue;
	if (!SendLevelIsValid(sendLevel))
		return ValueError(ERR_INVALID_VALUE, aValueStr, FAIL_OR_OK);
	g->SendLevel = sendLevel;
	return OK;
}



/////////////////
// MouseGetPos //
/////////////////

BIF_DECL(BIF_MouseGetPos)
// Returns OK or FAIL.
{
	// Since SYM_VAR is always VAR_NORMAL, these always resolve to normal vars or nullptr:
	Var *output_var_x = ParamIndexToOutputVar(0);
	Var *output_var_y = ParamIndexToOutputVar(1);
	Var *output_var_parent = ParamIndexToOutputVar(2);
	Var *output_var_child = ParamIndexToOutputVar(3);
	int aOptions = ParamIndexToOptionalInt(4, 0);

	POINT point;
	GetCursorPos(&point);  // Realistically, can't fail?

	POINT origin = {0};
	CoordToScreen(origin, COORD_MODE_MOUSE);

	if (output_var_x) // else the user didn't want the X coordinate, just the Y.
		if (!output_var_x->Assign(point.x - origin.x))
			_f_return_FAIL;
	if (output_var_y) // else the user didn't want the Y coordinate, just the X.
		if (!output_var_y->Assign(point.y - origin.y))
			_f_return_FAIL;

	_f_set_retval_p(_T(""), 0); // Set default.

	if (!output_var_parent && !output_var_child)
		_f_return_retval;

	if (output_var_parent)
		output_var_parent->Assign(); // Set default: empty.
	if (output_var_child)
		output_var_child->Assign(); // Set default: empty.

	// This is the child window.  Despite what MSDN says, WindowFromPoint() appears to fetch
	// a non-NULL value even when the mouse is hovering over a disabled control (at least on XP).
	HWND child_under_cursor = WindowFromPoint(point);
	if (!child_under_cursor)
		_f_return_retval;

	HWND parent_under_cursor = GetNonChildParent(child_under_cursor);  // Find the first ancestor that isn't a child.
	if (output_var_parent)
	{
		// Testing reveals that an invisible parent window never obscures another window beneath it as seen by
		// WindowFromPoint().  In other words, the below never happens, so there's no point in having it as a
		// documented feature:
		//if (!g->DetectHiddenWindows && !IsWindowVisible(parent_under_cursor))
		//	return output_var_parent->Assign();
		output_var_parent->AssignHWND(parent_under_cursor);
	}

	if (!output_var_child)
		_f_return_retval;

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
		_f_return_retval;

	if (aOptions & 0x02) // v1.0.43.06: Bitwise flag that means "return control's HWND vs. ClassNN".
	{
		output_var_child->AssignHWND(child_under_cursor);
		_f_return_retval;
	}

	class_and_hwnd_type cah;
	cah.hwnd = child_under_cursor;  // This is the specific control we need to find the sequence number of.
	TCHAR class_name[WINDOW_CLASS_SIZE];
	cah.class_name = class_name;
	if (!GetClassName(cah.hwnd, class_name, _countof(class_name) - 5))  // -5 to allow room for sequence number.
		_f_return_retval;
	cah.class_count = 0;  // Init for the below.
	cah.is_found = false; // Same.
	EnumChildWindows(parent_under_cursor, EnumChildFindSeqNum, (LPARAM)&cah); // Find this control's seq. number.
	if (!cah.is_found)
		_f_return_retval;
	// Append the class sequence number onto the class name and set the output param to be that value:
	sntprintfcat(class_name, _countof(class_name), _T("%d"), cah.class_count);
	if (!output_var_child->Assign(class_name))
		_f_return_FAIL;
	_f_return_retval;
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



////////////////
// FormatTime //
////////////////

BIF_DECL(BIF_FormatTime)
// The compressed code size of this function is about 1 KB (2 KB uncompressed), which compares
// favorably to using setlocale()+strftime(), which together are about 8 KB of compressed code
// (setlocale() seems to be needed to put the user's or system's locale into effect for strftime()).
// setlocale() weighs in at about 6.5 KB compressed (14 KB uncompressed).
{
	_f_param_string_opt(aYYYYMMDD, 0);
	_f_param_string_opt(aFormat, 1);

	#define FT_MAX_INPUT_CHARS 2000
	// Input/format length is restricted since it must be translated and expanded into a new format
	// string that uses single quotes around non-alphanumeric characters such as punctuation:
	if (_tcslen(aFormat) > FT_MAX_INPUT_CHARS)
		_f_throw_param(1);

	// Worst case expansion: .d.d.d.d. (9 chars) --> '.'d'.'d'.'d'.'d'.' (19 chars)
	// Buffer below is sized to a little more than twice as big as the largest allowed format,
	// which avoids having to constantly check for buffer overflow while translating aFormat
	// into format_buf:
	#define FT_MAX_OUTPUT_CHARS (2*FT_MAX_INPUT_CHARS + 10)
	TCHAR format_buf[FT_MAX_OUTPUT_CHARS + 1];
	TCHAR output_buf[FT_MAX_OUTPUT_CHARS + 1]; // The size of this is somewhat arbitrary, but buffer overflow is checked so it's safe.

	TCHAR yyyymmdd[256]; // Large enough to hold date/time and any options that follow it (note that D and T options can appear multiple times).
	*yyyymmdd = '\0';

	SYSTEMTIME st;
	LPTSTR options = NULL;

	if (!*aYYYYMMDD) // Use current local time by default.
		GetLocalTime(&st);
	else
	{
		tcslcpy(yyyymmdd, omit_leading_whitespace(aYYYYMMDD), _countof(yyyymmdd)); // Make a modifiable copy.
		if (*yyyymmdd < '0' || *yyyymmdd > '9') // First character isn't a digit, therefore...
		{
			// ... options are present without date (since yyyymmdd [if present] must come before options).
			options = yyyymmdd;
			GetLocalTime(&st);  // Use current local time by default.
		}
		else // Since the string starts with a digit, rules say it must be a YYYYMMDD string, possibly followed by options.
		{
			// Find first space or tab because YYYYMMDD portion might contain only the leading part of date/timestamp.
			if (options = StrChrAny(yyyymmdd, _T(" \t"))) // Find space or tab.
			{
				*options = '\0'; // Terminate yyyymmdd at the end of the YYYYMMDDHH24MISS string.
				options = omit_leading_whitespace(++options); // Point options to the right place (can be empty string).
			}
			//else leave options set to NULL to indicate that there are none.

			// Pass "false" for validation so that times can still be reported even if the year
			// is prior to 1601.  If the time and/or date is invalid, GetTimeFormat() and GetDateFormat()
			// will refuse to produce anything, which is documented behavior:
			if (!YYYYMMDDToSystemTime(yyyymmdd, st, false))
				_f_throw_param(0);
		}
	}
	
	// Set defaults.  Some can be overridden by options (if there are any options).
	LCID lcid = LOCALE_USER_DEFAULT;
	DWORD date_flags = 0, time_flags = 0;
	bool date_flags_specified = false, time_flags_specified = false, reverse_date_time = false;
	#define FT_FORMAT_NONE 0
	#define FT_FORMAT_TIME 1
	#define FT_FORMAT_DATE 2
	int format_type1 = FT_FORMAT_NONE;
	LPTSTR format2_marker = NULL; // Will hold the location of the first char of the second format (if present).
	bool do_null_format2 = false;  // Will be changed to true if a default date *and* time should be used.

	if (options) // Parse options.
	{
		LPTSTR option_end;
		TCHAR orig_char;
		for (LPTSTR next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
		{
			// Find the end of this option item:
			if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
				option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.

			// Permanently terminate in between options to help eliminate ambiguity for words contained
			// inside other words, and increase confidence in decimal and hexadecimal conversion.
			orig_char = *option_end;
			*option_end = '\0';

			++next_option;
			switch (_totupper(next_option[-1]))
			{
			case 'D':
				date_flags_specified = true;
				date_flags |= ATOU(next_option); // ATOU() for unsigned.
				break;
			case 'T':
				time_flags_specified = true;
				time_flags |= ATOU(next_option); // ATOU() for unsigned.
				break;
			case 'R':
				reverse_date_time = true;
				break;
			case 'L':
				lcid = !_tcsicmp(next_option, _T("Sys")) ? LOCALE_SYSTEM_DEFAULT : (LCID)ATOU(next_option);
				break;
			// If not one of the above, such as zero terminator or a number, just ignore it.
			}

			*option_end = orig_char; // Undo the temporary termination so that loop's omit_leading() will work.
		} // for() each item in option list
	} // Parse options.

	if (!*aFormat)
	{
		aFormat = NULL; // Tell GetDateFormat() and GetTimeFormat() to use default for the specified locale.
		if (!date_flags_specified) // No preference was given, so use long (which seems generally more useful).
			date_flags |= DATE_LONGDATE;
		if (!time_flags_specified)
			time_flags |= TIME_NOSECONDS;  // Seems more desirable/typical to default to no seconds.
		// Put the time first by default, though this is debatable (Metapad does it and I like it).
		format_type1 = reverse_date_time ? FT_FORMAT_DATE : FT_FORMAT_TIME;
		do_null_format2 = true;
	}
	else // aFormat is non-blank.
	{
		// Omit whitespace only for consideration of special keywords.  Whitespace is later kept for
		// a normal format string such as " MM/dd/yy":
		LPTSTR candidate = omit_leading_whitespace(aFormat);
		if (!_tcsicmp(candidate, _T("YWeek")))
		{
			GetISOWeekNumber(output_buf, st.wYear, GetYDay(st.wMonth, st.wDay, IS_LEAP_YEAR(st.wYear)), st.wDayOfWeek);
			_f_return(output_buf);
		}
		if (!_tcsicmp(candidate, _T("YDay")) || !_tcsicmp(candidate, _T("YDay0")))
		{
			int yday = GetYDay(st.wMonth, st.wDay, IS_LEAP_YEAR(st.wYear));
			// Format with leading zeroes if YDay0 was used.
			_stprintf(output_buf, candidate[4] == '0' ? _T("%03d") : _T("%d"), yday);
			_f_return(output_buf);
		}
		if (!_tcsicmp(candidate, _T("WDay")))
		{
			LPTSTR buf = _f_retval_buf;
			*buf = '1' + st.wDayOfWeek; // '1' vs '0' to convert to 1-based for compatibility with A_WDay.
			buf[1] = '\0';
			_f_return_p(buf, 1);
		}

		// Since above didn't return, check for those that require a call to GetTimeFormat/GetDateFormat
		// further below:
		if (!_tcsicmp(candidate, _T("ShortDate")))
		{
			aFormat = NULL;
			date_flags |= DATE_SHORTDATE;
			date_flags &= ~(DATE_LONGDATE | DATE_YEARMONTH); // If present, these would prevent it from working.
		}
		else if (!_tcsicmp(candidate, _T("LongDate")))
		{
			aFormat = NULL;
			date_flags |= DATE_LONGDATE;
			date_flags &= ~(DATE_SHORTDATE | DATE_YEARMONTH); // If present, these would prevent it from working.
		}
		else if (!_tcsicmp(candidate, _T("YearMonth")))
		{
			aFormat = NULL;
			date_flags |= DATE_YEARMONTH;
			date_flags &= ~(DATE_SHORTDATE | DATE_LONGDATE); // If present, these would prevent it from working.
		}
		else if (!_tcsicmp(candidate, _T("Time")))
		{
			format_type1 = FT_FORMAT_TIME;
			aFormat = NULL;
			if (!time_flags_specified)
				time_flags |= TIME_NOSECONDS;  // Seems more desirable/typical to default to no seconds.
		}
		else // Assume normal format string.
		{
			LPTSTR cp = aFormat, dp = format_buf;   // Initialize source and destination pointers.
			bool inside_their_quotes = false; // Whether we are inside a single-quoted string in the source.
			bool inside_our_quotes = false;   // Whether we are inside a single-quoted string of our own making in dest.
			for (; *cp; ++cp) // Transcribe aFormat into format_buf and also check for which comes first: date or time.
			{
				if (*cp == '\'') // Note that '''' (four consecutive quotes) is a single literal quote, which this logic handles okay.
				{
					if (inside_our_quotes)
					{
						// Upon encountering their quotes while we're still in ours, merge theirs with ours and
						// remark it as theirs.  This is done to avoid having two back-to-back quoted sections,
						// which would result in an unwanted literal single quote.  Example:
						// 'Some string'':' (the two quotes in the middle would be seen as a literal quote).
						inside_our_quotes = false;
						inside_their_quotes = true;
						continue;
					}
					if (inside_their_quotes)
					{
						// If next char needs to be quoted, don't close out this quote section because that
						// would introduce two consecutive quotes, which would be interpreted as a single
						// literal quote if its enclosed by two outer single quotes.  Instead convert this
						// quoted section over to "ours":
						if (cp[1] && !IsCharAlphaNumeric(cp[1]) && cp[1] != '\'') // Also consider single quotes to be theirs due to this example: dddd:''''y
							inside_our_quotes = true;
							// And don't do "*dp++ = *cp"
						else // there's no next-char or it's alpha-numeric, so it doesn't need to be inside quotes.
							*dp++ = *cp; // Close out their quoted section.
					}
					else // They're starting a new quoted section, so just transcribe this single quote as-is.
						*dp++ = *cp;
					inside_their_quotes = !inside_their_quotes; // Must be done after the above.
					continue;
				}
				// Otherwise, it's not a single quote.
				if (inside_their_quotes) // *cp is inside a single-quoted string, so it can be part of format/picture
					*dp++ = *cp; // Transcribe as-is.
				else
				{
					if (IsCharAlphaNumeric(*cp))
					{
						if (inside_our_quotes)
						{
							*dp++ = '\''; // Close out the previous quoted section, since this char should not be a part of it.
							inside_our_quotes = false;
						}
						if (_tcschr(_T("dMyg"), *cp)) // A format unique to Date is present.
						{
							if (!format_type1)
								format_type1 = FT_FORMAT_DATE;
							else if (format_type1 == FT_FORMAT_TIME && !format2_marker) // type2 should only be set if different than type1.
							{
								*dp++ = '\0';  // Terminate the first section and (below) indicate that there's a second.
								format2_marker = dp;  // Point it to the location in format_buf where the split should occur.
							}
						}
						else if (_tcschr(_T("hHmst"), *cp)) // A format unique to Time is present.
						{
							if (!format_type1)
								format_type1 = FT_FORMAT_TIME;
							else if (format_type1 == FT_FORMAT_DATE && !format2_marker) // type2 should only be set if different than type1.
							{
								*dp++ = '\0';  // Terminate the first section and (below) indicate that there's a second.
								format2_marker = dp;  // Point it to the location in format_buf where the split should occur.
							}
						}
						// For consistency, transcribe all AlphaNumeric chars not inside single quotes as-is
						// (numbers are transcribed in case they are ever used as part of pic/format).
						*dp++ = *cp;
					}
					else // Not alphanumeric, so enclose this and any other non-alphanumeric characters in single quotes.
					{
						if (!inside_our_quotes)
						{
							*dp++ = '\''; // Create a new quoted section of our own, since this char should be inside quotes to be understood.
							inside_our_quotes = true;
						}
						*dp++ = *cp;  // Add this character between the quotes, since it's of the right "type".
					}
				}
			} // for()
			if (inside_our_quotes)
				*dp++ = '\'';  // Close out our quotes.
			*dp = '\0'; // Final terminator.
			aFormat = format_buf; // Point it to the freshly translated format string, for use below.
		} // aFormat contains normal format/pic string.
	} // aFormat isn't blank.

	// If there are no date or time formats present, still do the transcription so that
	// any quoted strings and whatnot are resolved.  This increases runtime flexibility.
	// The below is also relied upon by "LongDate" and "ShortDate" above:
	if (!format_type1)
		format_type1 = FT_FORMAT_DATE;

	// MSDN: Time: "The function checks each of the time values to determine that it is within the
	// appropriate range of values. If any of the time values are outside the correct range, the
	// function fails, and sets the last-error to ERROR_INVALID_PARAMETER. 
	// Dates: "...year, month, day, and day of week. If the day of the week is incorrect, the
	// function uses the correct value, and returns no error. If any of the other date values
	// are outside the correct range, the function fails, and sets the last-error to ERROR_INVALID_PARAMETER.

	if (format_type1 == FT_FORMAT_DATE) // DATE comes first.
	{
		if (!GetDateFormat(lcid, date_flags, &st, aFormat, output_buf, FT_MAX_OUTPUT_CHARS))
			*output_buf = '\0';  // Ensure it's still the empty string, then try to continue to get the second half (if there is one).
	}
	else // TIME comes first.
		if (!GetTimeFormat(lcid, time_flags, &st, aFormat, output_buf, FT_MAX_OUTPUT_CHARS))
			*output_buf = '\0';  // Ensure it's still the empty string, then try to continue to get the second half (if there is one).

	if (format2_marker || do_null_format2) // There is also a second format present.
	{
		size_t output_buf_length = _tcslen(output_buf);
		LPTSTR output_buf_marker = output_buf + output_buf_length;
		LPTSTR format2;
		if (do_null_format2)
		{
			format2 = NULL;
			*output_buf_marker++ = ' '; // Provide a space between time and date.
			++output_buf_length;
		}
		else
			format2 = format2_marker;

		int buf_remaining_size = (int)(FT_MAX_OUTPUT_CHARS - output_buf_length);
		int result;

		if (format_type1 == FT_FORMAT_DATE) // DATE came first, so this one is TIME.
			result = GetTimeFormat(lcid, time_flags, &st, format2, output_buf_marker, buf_remaining_size);
		else
			result = GetDateFormat(lcid, date_flags, &st, format2, output_buf_marker, buf_remaining_size);
		if (!result)
			output_buf[output_buf_length] = '\0'; // Ensure the first part is still terminated and just return that rather than nothing.
	}

	_f_return(output_buf);
}



//////////////////////
// String Functions //
//////////////////////


BIF_DECL(BIF_StrCase)
{
	size_t length;
	LPTSTR contents = ParamIndexToString(0, _f_number_buf, &length);

	// Make a modifiable copy of the string to return:
	if (!TokenSetResult(aResultToken, contents, length))
		return;
	aResultToken.symbol = SYM_STRING;
	contents = aResultToken.marker;

	if (_f_callee_id == FID_StrLower)
		CharLower(contents);
	else if (_f_callee_id == FID_StrUpper)
		CharUpper(contents);
	else // Convert to title case.
		StrToTitleCase(contents);
}



BIF_DECL(BIF_StrReplace)
{
	size_t length; // Going in to StrReplace(), it's the haystack length. Later (coming out), it's the result length. 
	_f_param_string(source, 0, &length);
	_f_param_string(oldstr, 1);
	_f_param_string_opt(newstr, 2);

	// Maintain this together with the equivalent section of BIF_InStr:
	StringCaseSenseType string_case_sense = ParamIndexToCaseSense(3);
	if (string_case_sense == SCS_INVALID 
		|| string_case_sense == SCS_INSENSITIVE_LOGICAL) // Not supported, seems more useful to throw rather than using SCS_INSENSITIVE.
		_f_throw_param(3);

	Var *output_var_count = ParamIndexToOutputVar(4); 
	UINT replacement_limit = (UINT)ParamIndexToOptionalInt64(5, UINT_MAX); 
	

	// Note: The current implementation of StrReplace() should be able to handle any conceivable inputs
	// without an empty string causing an infinite loop and without going infinite due to finding the
	// search string inside of newly-inserted replace strings (e.g. replacing all occurrences
	// of b with bcd would not keep finding b in the newly inserted bcd, infinitely).
	LPTSTR dest;
	UINT found_count = StrReplace(source, oldstr, newstr, string_case_sense, replacement_limit, -1, &dest, &length); // Length of haystack is passed to improve performance because TokenToString() can often discover it instantaneously.

	if (!dest) // Failure due to out of memory.
		_f_throw_oom; 

	if (dest != source) // StrReplace() allocated new memory rather than returning "source" to us unaltered.
	{
		// Return the newly allocated memory directly to our caller. 
		aResultToken.AcceptMem(dest, length);
	}
	else // StrReplace gave us back "source" unaltered because no replacements were needed.
	{
		_f_set_retval_p(dest, length); 
	}

	if (output_var_count)
		output_var_count->Assign((DWORD)found_count);
	_f_return_retval;
}



BIF_DECL(BIF_StrSplit)
// Array := StrSplit(String [, Delimiters, OmitChars, MaxParts])
{
	LPTSTR aInputString = ParamIndexToString(0, _f_number_buf);
	LPTSTR *aDelimiterList = NULL;
	int aDelimiterCount = 0;
	LPTSTR aOmitList = _T("");
	int splits_left = -2;

	if (aParamCount > 1)
	{
		if (auto arr = dynamic_cast<Array *>(TokenToObject(*aParam[1])))
		{
			aDelimiterCount = arr->Length();
			aDelimiterList = (LPTSTR *)_alloca(aDelimiterCount * sizeof(LPTSTR *));
			if (!arr->ToStrings(aDelimiterList, aDelimiterCount, aDelimiterCount))
				// Array contains something other than a string.
				goto throw_invalid_delimiter;
			for (int i = 0; i < aDelimiterCount; ++i)
				if (!*aDelimiterList[i])
					// Empty string in delimiter list. Although it could be treated similarly to the
					// "no delimiter" case, it's far more likely to be an error. If ever this check
					// is removed, the loop below must be changed to support "" as a delimiter.
					goto throw_invalid_delimiter;
		}
		else
		{
			aDelimiterList = (LPTSTR *)_alloca(sizeof(LPTSTR *));
			*aDelimiterList = TokenToString(*aParam[1]);
			aDelimiterCount = **aDelimiterList != '\0'; // i.e. non-empty string.
		}
		if (aParamCount > 2)
		{
			aOmitList = TokenToString(*aParam[2]);
			if (aParamCount > 3)
				splits_left = (int)TokenToInt64(*aParam[3]) - 1;
		}
	}
	
	auto output_array = Array::Create();
	if (!output_array)
		goto throw_outofmem;

	if (!*aInputString // The input variable is blank, thus there will be zero elements.
		|| splits_left == -1) // The caller specified 0 parts.
		_f_return(output_array);
	
	LPTSTR contents_of_next_element, delimiter, new_starting_pos;
	size_t element_length, delimiter_length;

	if (aDelimiterCount) // The user provided a list of delimiters, so process the input variable normally.
	{
		for (contents_of_next_element = aInputString; ; )
		{
			if (   !splits_left // Limit reached.
				|| !(delimiter = InStrAny(contents_of_next_element, aDelimiterList, aDelimiterCount, delimiter_length))   ) // No delimiter found.
				break; // This is the only way out of the loop other than critical errors.
			element_length = delimiter - contents_of_next_element;
			if (*aOmitList && element_length > 0)
			{
				contents_of_next_element = omit_leading_any(contents_of_next_element, aOmitList, element_length);
				element_length = delimiter - contents_of_next_element; // Update in case above changed it.
				if (element_length)
					element_length = omit_trailing_any(contents_of_next_element, aOmitList, delimiter - 1);
			}
			// If there are no chars to the left of the delim, or if they were all in the list of omitted
			// chars, the variable will be assigned the empty string:
			if (!output_array->Append(contents_of_next_element, element_length))
				goto outofmem;
			contents_of_next_element = delimiter + delimiter_length;  // Omit the delimiter since it's never included in contents.
			if (splits_left > 0)
				--splits_left;
		}
	}
	else
	{
		// Otherwise aDelimiterList is empty, so store each char of aInputString in its own array element.
		LPTSTR cp, dp;
		for (cp = aInputString; ; ++cp)
		{
			if (!*cp)
				_f_return(output_array); // All done.
			for (dp = aOmitList; *dp; ++dp)
				if (*cp == *dp) // This char is a member of the omitted list, thus it is not included in the output array.
					break; // (inner loop)
			if (*dp) // Omitted.
				continue;
			if (!splits_left) // Limit reached (checked only after excluding omitted chars).
				break; // This is the only way out of the loop other than critical errors.
			if (splits_left > 0)
				--splits_left;
			if (!output_array->Append(cp, 1))
				goto outofmem;
		}
		contents_of_next_element = cp;
	}
	// Since above used break rather than goto or return, either the limit was reached or there are
	// no more delimiters, so store the remainder of the string minus any characters to be omitted.
	element_length = _tcslen(contents_of_next_element);
	if (*aOmitList && element_length > 0)
	{
		new_starting_pos = omit_leading_any(contents_of_next_element, aOmitList, element_length);
		element_length -= (new_starting_pos - contents_of_next_element); // Update in case above changed it.
		contents_of_next_element = new_starting_pos;
		if (element_length)
			// If this is true, the string must contain at least one char that isn't in the list
			// of omitted chars, otherwise omit_leading_any() would have already omitted them:
			element_length = omit_trailing_any(contents_of_next_element, aOmitList
				, contents_of_next_element + element_length - 1);
	}
	// If there are no chars to the left of the delim, or if they were all in the list of omitted
	// chars, the item will be an empty string:
	if (output_array->Append(contents_of_next_element, element_length))
		_f_return(output_array); // All done.
	//else memory allocation failed, so fall through:
outofmem:
	// The fact that this section is executing means that a memory allocation failed and caused the
	// loop to break, so throw an exception.
	output_array->Release(); // Since we're not returning it.
	// Below: Using goto probably won't reduce code size (due to compiler optimizations), but keeping
	// rarely executed code at the end of the function might help performance due to cache utilization.
throw_outofmem:
	// Either goto was used or FELL THROUGH FROM ABOVE.
	_f_throw_oom;
throw_invalid_delimiter:
	_f_throw_param(1);
}



ResultType SplitPath(LPCTSTR aFileSpec, Var *output_var_name, Var *output_var_dir, Var *output_var_ext, Var *output_var_name_no_ext, Var *output_var_drive)
{
	// For URLs, "drive" is defined as the server name, e.g. http://somedomain.com
	LPCTSTR name = _T(""), name_delimiter = NULL, drive_end = NULL; // Set defaults to improve maintainability.
	LPCTSTR drive = omit_leading_whitespace(aFileSpec); // i.e. whitespace is considered for everything except the drive letter or server name, so that a pathless filename can have leading whitespace.
	LPCTSTR colon_double_slash = _tcsstr(aFileSpec, _T("://"));

	if (colon_double_slash) // This is a URL such as ftp://... or http://...
	{
		if (   !(drive_end = _tcschr(colon_double_slash + 3, '/'))   )
		{
			if (   !(drive_end = _tcschr(colon_double_slash + 3, '\\'))   ) // Try backslash so that things like file://C:\Folder\File.txt are supported.
				drive_end = colon_double_slash + _tcslen(colon_double_slash); // Set it to the position of the zero terminator instead.
				// And because there is no filename, leave name and name_delimiter set to their defaults.
			//else there is a backslash, e.g. file://C:\Folder\File.txt, so treat that backslash as the end of the drive name.
		}
		name_delimiter = drive_end; // Set default, to be possibly overridden below.
		// Above has set drive_end to one of the following:
		// 1) The slash that occurs to the right of the doubleslash in a URL.
		// 2) The backslash that occurs to the right of the doubleslash in a URL.
		// 3) The zero terminator if there is no slash or backslash to the right of the doubleslash.
		if (*drive_end) // A slash or backslash exists to the right of the server name.
		{
			if (*(drive_end + 1))
			{
				// Find the rightmost slash.  At this stage, this is known to find the correct slash.
				// In the case of a file at the root of a domain such as http://domain.com/root_file.htm,
				// the directory consists of only the domain name, e.g. http://domain.com.  This is because
				// the directory always the "drive letter" by design, since that is more often what the
				// caller wants.  A script can use StringReplace to remove the drive/server portion from
				// the directory, if desired.
				name_delimiter = _tcsrchr(aFileSpec, '/');
				if (name_delimiter == colon_double_slash + 2) // To reach this point, it must have a backslash, something like file://c:\folder\file.txt
					name_delimiter = _tcsrchr(aFileSpec, '\\'); // Will always be found.
				name = name_delimiter + 1; // This will be the empty string for something like http://domain.com/dir/
			}
			//else something like http://domain.com/, so leave name and name_delimiter set to their defaults.
		}
		//else something like http://domain.com, so leave name and name_delimiter set to their defaults.
	}
	else // It's not a URL, just a file specification such as c:\my folder\my file.txt, or \\server01\folder\file.txt
	{
		// Differences between _splitpath() and the method used here:
		// _splitpath() doesn't include drive in output_var_dir, it includes a trailing
		// backslash, it includes the . in the extension, it considers ":" to be a filename.
		// _splitpath(pathname, drive, dir, file, ext);
		//char sdrive[16], sdir[MAX_PATH], sname[MAX_PATH], sext[MAX_PATH];
		//_splitpath(aFileSpec, sdrive, sdir, sname, sext);
		//if (output_var_name_no_ext)
		//	output_var_name_no_ext->Assign(sname);
		//strcat(sname, sext);
		//if (output_var_name)
		//	output_var_name->Assign(sname);
		//if (output_var_dir)
		//	output_var_dir->Assign(sdir);
		//if (output_var_ext)
		//	output_var_ext->Assign(sext);
		//if (output_var_drive)
		//	output_var_drive->Assign(sdrive);
		//return OK;

		// Don't use _splitpath() since it supposedly doesn't handle UNC paths correctly,
		// and anyway we need more info than it provides.  Also note that it is possible
		// for a file to begin with space(s) or a dot (if created programmatically), so
		// don't trim or omit leading space unless it's known to be an absolute path.

		// Note that "C:Some File.txt" is a valid filename in some contexts, which the below
		// tries to take into account.  However, there will be no way for this command to
		// return a path that differentiates between "C:Some File.txt" and "C:\Some File.txt"
		// since the first backslash is not included with the returned path, even if it's
		// the root directory (i.e. "C:" is returned in both cases).  The "C:Filename"
		// convention is pretty rare, and anyway this trait can be detected via something like
		// IfInString, Filespec, :, IfNotInString, Filespec, :\, MsgBox Drive with no absolute path.

		// UNCs are detected with this approach so that double sets of backslashes -- which sometimes
		// occur by accident in "built filespecs" and are tolerated by the OS -- are not falsely
		// detected as UNCs.
		if (drive[0] == '\\' && drive[1] == '\\') // Relies on short-circuit evaluation order.
		{
			if (   !(drive_end = _tcschr(drive + 2, '\\'))   )
				drive_end = drive + _tcslen(drive); // Set it to the position of the zero terminator instead.
		}
		else if (*(drive + 1) == ':') // It's an absolute path.
			// Assign letter and colon for consistency with server naming convention above.
			// i.e. so that server name and drive can be used without having to worry about
			// whether it needs a colon added or not.
			drive_end = drive + 2;
		else
		{
			// It's debatable, but it seems best to return a blank drive if a aFileSpec is a relative path.
			// rather than trying to use GetFullPathName() on a potentially non-existent file/dir.
			// _splitpath() doesn't fetch the drive letter of relative paths either.  This also reports
			// a blank drive for something like file://C:\My Folder\My File.txt, which seems too rarely
			// to justify a special mode.
			drive_end = _T("");
			drive = drive_end; // This is necessary to allow Assign() to work correctly later below, since it interprets a length of zero as "use string's entire length".
		}

		if (   !(name_delimiter = _tcsrchr(aFileSpec, '\\'))   ) // No backslash.
			if (   !(name_delimiter = _tcsrchr(aFileSpec, ':'))   ) // No colon.
				name_delimiter = NULL; // Indicate that there is no directory.

		name = name_delimiter ? name_delimiter + 1 : aFileSpec; // If no delimiter, name is the entire string.
	}

	// The above has now set the following variables:
	// name: As an empty string or the actual name of the file, including extension.
	// name_delimiter: As NULL if there is no directory, otherwise, the end of the directory's name.
	// drive: As the start of the drive/server name, e.g. C:, \\Workstation01, http://domain.com, etc.
	// drive_end: As the position after the drive's last character, either a zero terminator, slash, or backslash.

	if (output_var_name && !output_var_name->Assign(name))
		return FAIL;

	if (output_var_dir)
	{
		if (!name_delimiter)
			output_var_dir->Assign(); // Shouldn't fail.
		else if (*name_delimiter == '\\' || *name_delimiter == '/')
		{
			if (!output_var_dir->Assign(aFileSpec, (VarSizeType)(name_delimiter - aFileSpec)))
				return FAIL;
		}
		else // *name_delimiter == ':', e.g. "C:Some File.txt".  If aFileSpec starts with just ":",
			 // the dir returned here will also start with just ":" since that's rare & illegal anyway.
			if (!output_var_dir->Assign(aFileSpec, (VarSizeType)(name_delimiter - aFileSpec + 1)))
				return FAIL;
	}

	LPCTSTR ext_dot = _tcsrchr(name, '.');
	if (output_var_ext)
	{
		// Note that the OS doesn't allow filenames to end in a period.
		if (!ext_dot)
			output_var_ext->Assign();
		else
			if (!output_var_ext->Assign(ext_dot + 1)) // Can be empty string if filename ends in just a dot.
				return FAIL;
	}

	if (output_var_name_no_ext && !output_var_name_no_ext->Assign(name, (VarSizeType)(ext_dot ? ext_dot - name : _tcslen(name))))
		return FAIL;

	if (output_var_drive && !output_var_drive->Assign(drive, (VarSizeType)(drive_end - drive)))
		return FAIL;

	return OK;
}

BIF_DECL(BIF_SplitPath)
{
	LPTSTR mem_to_free = nullptr;
	_f_param_string(aFileSpec, 0);
	Var *vars[6];
	for (int i = 1; i < _countof(vars); ++i)
		vars[i] = ParamIndexToOutputVar(i);
	if (aParam[0]->symbol == SYM_VAR) // Check for overlap of input/output vars.
	{
		vars[0] = aParam[0]->var;
		// There are cases where this could be avoided, such as by careful ordering of the assignments
		// in SplitPath(), or when there's only one output var.  Also, real paths are generally short
		// enough that stack memory could be used.  However, perhaps simple is best in this case.
		for (int i = 1; i < _countof(vars); ++i)
		{
			if (vars[i] == vars[0])
			{
				aFileSpec = mem_to_free = _tcsdup(aFileSpec);
				if (!mem_to_free)
					_f_throw_oom;
				break;
			}
		}
	}
	aResultToken.SetValue(_T(""), 0);
	if (!SplitPath(aFileSpec, vars[1], vars[2], vars[3], vars[4], vars[5]))
		aResultToken.SetExitResult(FAIL);
	free(mem_to_free);
}



int SortWithOptions(const void *a1, const void *a2)
// Decided to just have one sort function since there are so many permutations.  The performance
// will be a little bit worse, but it seems simpler to implement and maintain.
// This function's input parameters are pointers to the elements of the array.  Since those elements
// are themselves pointers, the input parameters are therefore pointers to pointers (handles).
{
	LPTSTR sort_item1 = *(LPTSTR *)a1;
	LPTSTR sort_item2 = *(LPTSTR *)a2;
	if (g_SortColumnOffset > 0)
	{
		// Adjust each string (even for numerical sort) to be the right column position,
		// or the position of its zero terminator if the column offset goes beyond its length:
		size_t length = _tcslen(sort_item1);
		sort_item1 += (size_t)g_SortColumnOffset > length ? length : g_SortColumnOffset;
		length = _tcslen(sort_item2);
		sort_item2 += (size_t)g_SortColumnOffset > length ? length : g_SortColumnOffset;
	}
	if (g_SortNumeric) // Takes precedence over g_SortCaseSensitive
	{
		// For now, assume both are numbers.  If one of them isn't, it will be sorted as a zero.
		// Thus, all non-numeric items should wind up in a sequential, unsorted group.
		// Resolve only once since parts of the ATOF() macro are inline:
		double item1_minus_2 = ATOF(sort_item1) - ATOF(sort_item2);
		if (!item1_minus_2) // Exactly equal.
			return (sort_item1 > sort_item2) ? 1 : -1; // Stable sort.
		// Otherwise, it's either greater or less than zero:
		int result = (item1_minus_2 > 0.0) ? 1 : -1;
		return g_SortReverse ? -result : result;
	}
	// Otherwise, it's a non-numeric sort.
	// v1.0.43.03: Added support the new locale-insensitive mode.
	int result = (g_SortCaseSensitive != SCS_INSENSITIVE_LOGICAL)
		? tcscmp2(sort_item1, sort_item2, g_SortCaseSensitive) // Resolve large macro only once for code size reduction.
		: StrCmpLogicalW(sort_item1, sort_item2);
	if (!result)
		result = (sort_item1 > sort_item2) ? 1 : -1; // Stable sort.
	return g_SortReverse ? -result : result;
}



int SortByNakedFilename(const void *a1, const void *a2)
// See comments in prior function for details.
{
	LPTSTR sort_item1 = *(LPTSTR *)a1;
	LPTSTR sort_item2 = *(LPTSTR *)a2;
	LPTSTR cp;
	if (cp = _tcsrchr(sort_item1, '\\'))  // Assign
		sort_item1 = cp + 1;
	if (cp = _tcsrchr(sort_item2, '\\'))  // Assign
		sort_item2 = cp + 1;
	// v1.0.43.03: Added support the new locale-insensitive mode.
	int result = (g_SortCaseSensitive != SCS_INSENSITIVE_LOGICAL)
		? tcscmp2(sort_item1, sort_item2, g_SortCaseSensitive) // Resolve large macro only once for code size reduction.
		: StrCmpLogicalW(sort_item1, sort_item2);
	if (!result)
		result = (sort_item1 > sort_item2) ? 1 : -1; // Stable sort.
	return g_SortReverse ? -result : result;
}



struct sort_rand_type
{
	LPTSTR cp; // This must be the first member of the struct, otherwise the array trickery in PerformSort will fail.
	union
	{
		// This must be the same size in bytes as the above, which is why it's maintained as a union with
		// a char* rather than a plain int.
		LPTSTR unused;
		int rand;
	};
};

int SortRandom(const void *a1, const void *a2)
// See comments in prior functions for details.
{
	return ((sort_rand_type *)a1)->rand - ((sort_rand_type *)a2)->rand;
}

int SortUDF(const void *a1, const void *a2)
// See comments in prior function for details.
{
	if (!g_SortFuncResult || g_SortFuncResult == EARLY_EXIT)
		return 0;

	// The following isn't necessary because by definition, the current thread isn't paused because it's the
	// thing that called the sort in the first place.
	//g_script.UpdateTrayIcon();

	LPTSTR aStr1 = *(LPTSTR *)a1;
	LPTSTR aStr2 = *(LPTSTR *)a2;
	ExprTokenType param[] = { aStr1, aStr2, __int64(aStr2 - aStr1) };
	__int64 i64;
	g_SortFuncResult = CallMethod(g_SortFunc, g_SortFunc, nullptr, param, _countof(param), &i64);
	// An alternative to g_SortFuncResult using 'throw' to abort qsort() produced slightly
	// smaller code, but in release builds the program crashed with code 0xC0000409 and the
	// catch() block never executed.

	int returned_int;
	if (i64)  // Maybe there's a faster/better way to do these checks. Can't simply typecast to an int because some large positives wrap into negative, maybe vice versa.
		returned_int = i64 < 0 ? -1 : 1;
	else  // The result was 0 or "".  Since returning 0 would make the results less predictable and doesn't seem to have any benefit,
		returned_int = aStr1 < aStr2 ? -1 : 1; // use the relative positions of the items as a tie-breaker.

	return returned_int;
}



BIF_DECL(BIF_Sort)
{
	// Set defaults in case of early goto:
	LPTSTR mem_to_free = NULL;
	LPTSTR *item = NULL; // The index/pointer list used for the sort.
	IObject *sort_func_orig = g_SortFunc; // Because UDFs can be interrupted by other threads -- and because UDFs can themselves call Sort with some other UDF (unlikely to be sure) -- backup & restore original g_SortFunc so that the "collapsing in reverse order" behavior will automatically ensure proper operation.
	ResultType sort_func_result_orig = g_SortFuncResult;
	g_SortFunc = NULL; // Now that original has been saved above, reset to detect whether THIS sort uses a UDF.
	g_SortFuncResult = OK;

	_f_param_string(aContents, 0);
	_f_param_string_opt(aOptions, 1);

	// Resolve options.  First set defaults for options:
	TCHAR delimiter = '\n';
	g_SortCaseSensitive = SCS_INSENSITIVE;
	g_SortNumeric = false;
	g_SortReverse = false;
	g_SortColumnOffset = 0;
	bool trailing_delimiter_indicates_trailing_blank_item = false, terminate_last_item_with_delimiter = false
		, trailing_crlf_added_temporarily = false, sort_by_naked_filename = false, sort_random = false
		, omit_dupes = false;
	LPTSTR cp;

	for (cp = aOptions; *cp; ++cp)
	{
		switch(_totupper(*cp))
		{
		case 'C':
			if (ctoupper(cp[1]) == 'L')
			{
				if (!_tcsnicmp(cp+2, _T("Logical") + 1, 6)) // CLogical.  Using "Logical" + 1 instead of "ogical" was confirmed to eliminate one string from the binary (due to string pooling).
				{
					cp += 7;
					g_SortCaseSensitive = SCS_INSENSITIVE_LOGICAL;
				}
				else
				{
					if (!_tcsnicmp(cp+2, _T("Locale") + 1, 5)) // CLocale
						cp += 6;
					else // CL
						++cp;
					g_SortCaseSensitive = SCS_INSENSITIVE_LOCALE;
				}
			}
			else if (!_tcsnicmp(cp+1, _T("Off"), 3)) // COff.  Using ctoupper() here significantly increased code size.
			{
				cp += 3;
				g_SortCaseSensitive = SCS_INSENSITIVE;
			}
			else if (cp[1] == '0') // C0
				g_SortCaseSensitive = SCS_INSENSITIVE;
			else // C  C1  COn
			{
				if (!_tcsnicmp(cp+1, _T("On"), 2)) // COn.  Using ctoupper() here significantly increased code size.
					cp += 2;
				g_SortCaseSensitive = SCS_SENSITIVE;
			}
			break;
		case 'D':
			if (!cp[1]) // Avoids out-of-bounds when the loop's own ++cp is done.
				break;
			++cp;
			if (*cp)
				delimiter = *cp;
			break;
		case 'N':
			g_SortNumeric = true;
			break;
		case 'P':
			// Use atoi() vs. ATOI() to avoid interpreting something like 0x01C as hex
			// when in fact the C was meant to be an option letter:
			g_SortColumnOffset = _ttoi(cp + 1);
			if (g_SortColumnOffset < 1)
				g_SortColumnOffset = 1;
			--g_SortColumnOffset;  // Convert to zero-based.
			break;
		case 'R':
			if (!_tcsnicmp(cp, _T("Random"), 6))
			{
				sort_random = true;
				cp += 5; // Point it to the last char so that the loop's ++cp will point to the character after it.
			}
			else
				g_SortReverse = true;
			break;
		case 'U':  // Unique.
			omit_dupes = true;
			break;
		case 'Z':
			// By setting this to true, the final item in the list, if it ends in a delimiter,
			// is considered to be followed by a blank item:
			trailing_delimiter_indicates_trailing_blank_item = true;
			break;
		case '\\':
			sort_by_naked_filename = true;
		}
	}

	if (!ParamIndexIsOmitted(2))
	{
		if (  !(g_SortFunc = ParamIndexToObject(2))  )
		{
			aResultToken.ParamError(2, aParam[2]);
			goto end;
		}
		g_SortFunc->AddRef(); // Must be done in case the parameter was SYM_VAR and that var gets reassigned.
	}
	
	if (!*aContents) // Input is empty, nothing to sort, return empty string.
	{
		aResultToken.Return(_T(""), 0);
		goto end;
	}
	
	size_t item_count;

	// Check how many delimiters are present:
	for (item_count = 1, cp = aContents; *cp; ++cp)  // Start at 1 since item_count is delimiter_count+1
		if (*cp == delimiter)
			++item_count;
	size_t aContents_length = cp - aContents;

	// If the last character in the unsorted list is a delimiter then technically the final item
	// in the list is a blank item.  However, if the options specify not to allow that, don't count that
	// blank item as an actual item (i.e. omit it from the list):
	if (!trailing_delimiter_indicates_trailing_blank_item && cp > aContents && cp[-1] == delimiter)
	{
		terminate_last_item_with_delimiter = true; // Have a later stage add the delimiter to the end of the sorted list so that the length and format of the sorted list is the same as the unsorted list.
		--item_count; // Indicate that the blank item at the end of the list isn't being counted as an item.
	}
	else // The final character isn't a delimiter or it is but there's a blank item to its right.  Either way, the following is necessary (verified correct).
	{
		if (delimiter == '\n')
		{
			LPTSTR first_delimiter = _tcschr(aContents, delimiter);
			if (first_delimiter && first_delimiter > aContents && first_delimiter[-1] == '\r')
			{
				// v1.0.47.05:
				// Here the delimiter is effectively CRLF vs. LF.  Therefore, signal a later section to append
				// an extra CRLF to simplify the code and fix bugs that existed prior to v1.0.47.05.
				// One such bug is that sorting "x`r`nx" in unique mode would previously produce two
				// x's rather than one because the first x is seen to have a `r to its right but the
				// second lacks it (due to the CRLF delimiter being simulated via LF-only).
				//
				// OLD/OBSOLETE comment from a section that was removed because it's now handled by this section:
				// Check if special handling is needed due to the following situation:
				// Delimiter is LF but the contents are lines delimited by CRLF, not just LF
				// and the original/unsorted list's last item was not terminated by an
				// "allowed delimiter".  The symptoms of this are that after the sort, the
				// last item will end in \r when it should end in no delimiter at all.
				// This happens pretty often, such as when the clipboard contains files.
				// In the future, an option letter can be added to turn off this workaround,
				// but it seems unlikely that anyone would ever want to.
				trailing_crlf_added_temporarily = true;
				terminate_last_item_with_delimiter = true; // This is done to "fool" later sections into thinking the list ends in a CRLF that doesn't have a blank item to the right, which in turn simplifies the logic and/or makes it more understandable.
			}
		}
	}

	if (item_count == 1) // 1 item is already sorted, and no dupes are possible.
	{
		aResultToken.Return(aContents, aContents_length);
		goto end;
	}

	// v1.0.47.05: It simplifies the code a lot to allocate and/or improves understandability to allocate
	// memory for trailing_crlf_added_temporarily even though technically it's done only to make room to
	// append the extra CRLF at the end.
	// v2.0: Never modify the caller's aContents, since it may be a quoted literal string or variable.
	//if (g_SortFunc || trailing_crlf_added_temporarily) // Do this here rather than earlier with the options parsing in case the function-option is present twice (unlikely, but it would be a memory leak due to strdup below).  Doing it here also avoids allocating if it isn't necessary.
	{
		// Comment is obsolete because if aContents is in a deref buffer, it has been privatized by ExpandArgs():
		// When g_SortFunc!=NULL, the copy of the string is needed because aContents may be in the deref buffer,
		// and that deref buffer is about to be overwritten by the execution of the script's UDF body.
		if (   !(mem_to_free = tmalloc(aContents_length + 3))   ) // +1 for terminator and +2 in case of trailing_crlf_added_temporarily.
		{
			aResultToken.MemoryError();
			goto end;
		}
		tmemcpy(mem_to_free, aContents, aContents_length + 1); // memcpy() usually benches a little faster than _tcscpy().
		aContents = mem_to_free;
		if (trailing_crlf_added_temporarily)
		{
			// Must be added early so that the next stage will terminate at the \n, leaving the \r as
			// the ending character in this item.
			_tcscpy(aContents + aContents_length, _T("\r\n"));
			aContents_length += 2;
		}
	}

	// Create the array of pointers that points into aContents to each delimited item.
	// Use item_count + 1 to allow space for the last (blank) item in case
	// trailing_delimiter_indicates_trailing_blank_item is false:
	int unit_size = sort_random ? 2 : 1;
	size_t item_size = unit_size * sizeof(LPTSTR);
	item = (LPTSTR *)malloc((item_count + 1) * item_size);
	if (!item)
	{
		aResultToken.MemoryError();
		goto end;
	}

	// If sort_random is in effect, the above has created an array twice the normal size.
	// This allows the random numbers to be interleaved inside the array as though it
	// were an array consisting of sort_rand_type (which it actually is when viewed that way).
	// Because of this, the array should be accessed through pointer addition rather than
	// indexing via [].

	// Scan aContents and do the following:
	// 1) Replace each delimiter with a terminator so that the individual items can be seen
	//    as real strings by the SortWithOptions() and when copying the sorted results back
	//    into the result.
	// 2) Store a marker/pointer to each item (string) in aContents so that we know where
	//    each item begins for sorting and recopying purposes.
	LPTSTR *item_curr = item; // i.e. Don't use [] indexing for the reason in the paragraph previous to above.
	for (item_count = 0, cp = *item_curr = aContents; *cp; ++cp)
	{
		if (*cp == delimiter)  // Each delimiter char becomes the terminator of the previous key phrase.
		{
			*cp = '\0';  // Terminate the item that appears before this delimiter.
			++item_count;
			if (sort_random)
			{
				auto &n = *(unsigned int *)(item_curr + 1); // i.e. the randoms are in the odd fields, the pointers in the even.
				GenRandom(&n, sizeof(n));
				n >>= 1;
				// For the above:
				// The random numbers given to SortRandom() must be within INT_MAX of each other, otherwise
				// integer overflow would cause the difference between two items to flip signs depending
				// on how far apart the numbers are, which would mean for some cases where A < B and B < C,
				// it's possible that C > A.  Without n >>= 1, the problem can be proven via the following
				// script, which shows a sharply biased distribution:
				//count := 0
				//Loop total := 10000
				//{
				//	var := Sort("1`n2`n3`n4`n5`n", "Random")
				//	if SubStr(var, 1, 1) = 5
				//		count += 1
				//}
				//MsgBox Round(count / total * 100, 2) "%"
			}
			item_curr += unit_size; // i.e. Don't use [] indexing for the reason described above.
			*item_curr = cp + 1; // Make a pointer to the next item's place in aContents.
		}
	}
	// The above reset the count to 0 and recounted it.  So now re-add the last item to the count unless it was
	// disqualified earlier. Verified correct:
	if (!terminate_last_item_with_delimiter) // i.e. either trailing_delimiter_indicates_trailing_blank_item==true OR the final character isn't a delimiter. Either way the final item needs to be added.
	{
		++item_count;
		if (sort_random) // Provide a random number for the last item.
		{
			auto &n = *(unsigned int *)(item_curr + 1); // i.e. the randoms are in the odd fields, the pointers in the even.
			GenRandom(&n, sizeof(n));
			n >>= 1;
		}
	}
	else // Since the final item is not included in the count, point item_curr to the one before the last, for use below.
		item_curr -= unit_size;

	// Now aContents has been divided up based on delimiter.  Sort the array of pointers
	// so that they indicate the correct ordering to copy aContents into output_var:
	if (g_SortFunc) // Takes precedence other sorting methods.
	{
		qsort((void *)item, item_count, item_size, SortUDF);
		if (!g_SortFuncResult || g_SortFuncResult == EARLY_EXIT)
		{
			aResultToken.SetExitResult(g_SortFuncResult);
			goto end;
		}
	}
	else if (sort_random) // Takes precedence over all remaining options.
		qsort((void *)item, item_count, item_size, SortRandom);
	else
		qsort((void *)item, item_count, item_size, sort_by_naked_filename ? SortByNakedFilename : SortWithOptions);

	// Allocate space to store the result.
	if (!TokenSetResult(aResultToken, NULL, aContents_length))
		goto end;
	aResultToken.symbol = SYM_STRING;

	// Set default in case original last item is still the last item, or if last item was omitted due to being a dupe:
	size_t i, item_count_minus_1 = item_count - 1;
	DWORD omit_dupe_count = 0;
	bool keep_this_item;
	LPTSTR source, dest;
	LPTSTR item_prev = NULL;

	// Copy the sorted result back into output_var.  Do all except the last item, since the last
	// item gets special treatment depending on the options that were specified.
	item_curr = item; // i.e. Don't use [] indexing for the reason described higher above (same applies to item += unit_size below).
	for (dest = aResultToken.marker, i = 0; i < item_count; ++i, item_curr += unit_size)
	{
		keep_this_item = true;  // Set default.
		if (omit_dupes && item_prev)
		{
			// Update to the comment below: Exact dupes will still be removed when sort_by_naked_filename
			// or g_SortColumnOffset is in effect because duplicate lines would still be adjacent to
			// each other even in these modes.  There doesn't appear to be any exceptions, even if
			// some items in the list are sorted as blanks due to being shorter than the specified 
			// g_SortColumnOffset.
			// As documented, special dupe-checking modes are not offered when sort_by_naked_filename
			// is in effect, or g_SortColumnOffset is greater than 1.  That's because the need for such
			// a thing seems too rare (and the result too strange) to justify the extra code size.
			// However, adjacent dupes are still removed when any of the above modes are in effect,
			// or when the "random" mode is in effect.  This might have some usefulness; for example,
			// if a list of songs is sorted in random order, but it contains "favorite" songs listed twice,
			// the dupe-removal feature would remove duplicate songs if they happen to be sorted
			// to lie adjacent to each other, which would be useful to prevent the same song from
			// playing twice in a row.
			if (g_SortNumeric && !g_SortColumnOffset)
				// if g_SortColumnOffset is zero, fall back to the normal dupe checking in case its
				// ever useful to anyone.  This is done because numbers in an offset column are not supported
				// since the extra code size doensn't seem justified given the rarity of the need.
				keep_this_item = (ATOF(*item_curr) != ATOF(item_prev)); // ATOF() ignores any trailing \r in CRLF mode, so no extra logic is needed for that.
			else if (g_SortCaseSensitive == SCS_INSENSITIVE_LOGICAL)
				keep_this_item = StrCmpLogicalW(*item_curr, item_prev);
			else
				keep_this_item = tcscmp2(*item_curr, item_prev, g_SortCaseSensitive); // v1.0.43.03: Added support for locale-insensitive mode.
				// Permutations of sorting case sensitive vs. eliminating duplicates based on case sensitivity:
				// 1) Sort is not case sens, but dupes are: Won't work because sort didn't necessarily put
				//    same-case dupes adjacent to each other.
				// 2) Converse: probably not reliable because there could be unrelated items in between
				//    two strings that are duplicates but weren't sorted adjacently due to their case.
				// 3) Both are case sensitive: seems okay
				// 4) Both are not case sensitive: seems okay
				//
				// In light of the above, using the g_SortCaseSensitive flag to control the behavior of
				// both sorting and dupe-removal seems best.
		}
		if (keep_this_item)
		{
			for (source = *item_curr; *source;)
				*dest++ = *source++;
			// If we're at the last item and the original list's last item had a terminating delimiter
			// and the specified options said to treat it not as a delimiter but as a final char of sorts,
			// include it after the item that is now last so that the overall layout is the same:
			if (i < item_count_minus_1 || terminate_last_item_with_delimiter)
				*dest++ = delimiter;  // Put each item's delimiter back in so that format is the same as the original.
			item_prev = *item_curr; // Since the item just processed above isn't a dupe, save this item to compare against the next item.
		}
		else // This item is a duplicate of the previous item.
		{
			++omit_dupe_count; // But don't change the value of item_prev.
			// v1.0.47.05: The following section fixes the fact that the "unique" option would sometimes leave
			// a trailing delimiter at the end of the sorted list.  For example, sorting "x|x" in unique mode
			// would formerly produce "x|":
			if (i == item_count_minus_1 && !terminate_last_item_with_delimiter) // This is the last item and its being omitted, so remove the previous item's trailing delimiter (unless a trailing delimiter is mandatory).
				--dest; // Remove the previous item's trailing delimiter there's nothing for it to delimit due to omission of this duplicate.
		}
	} // for()

	// Terminate the variable's contents.
	if (trailing_crlf_added_temporarily) // Remove the CRLF only after its presence was used above to simplify the code by reducing the number of types/cases.
	{
		dest[-2] = '\0';
		aResultToken.marker_length -= 2;
	}
	else
		*dest = '\0';

	if (omit_dupes)
	{
		if (omit_dupe_count) // Update the length to actual whenever at least one dupe was omitted.
		{
			aResultToken.marker_length = _tcslen(aResultToken.marker);
		}
	}
	//else it is not necessary to set the length here because its length hasn't
	// changed since it was originally set by the above call TokenSetResult.

end:
	if (!item)
		free(item);
	if (mem_to_free)
		free(mem_to_free);
	if (g_SortFunc)
		g_SortFunc->Release();
	g_SortFunc = sort_func_orig;
	g_SortFuncResult = sort_func_result_orig;
}



/////////////////////
// Drive Functions //
/////////////////////


void DriveSpace(ResultToken &aResultToken, LPTSTR aPath, bool aGetFreeSpace)
// Because of NTFS's ability to mount volumes into a directory, a path might not necessarily
// have the same amount of free space as its root drive.  However, I'm not sure if this
// method here actually takes that into account.
{
	ASSERT(aPath && *aPath);

	TCHAR buf[MAX_PATH]; // MAX_PATH vs T_MAX_PATH because testing shows it doesn't support long paths even with \\?\.
	tcslcpy(buf, aPath, _countof(buf));
	size_t length = _tcslen(buf);
	if (buf[length - 1] != '\\' // Trailing backslash is absent,
		&& length + 1 < _countof(buf)) // and there's room to fix it.
	{
		buf[length++] = '\\';
		buf[length] = '\0';
	}
	//else it should still work unless this is a UNC path.

	// MSDN: "The GetDiskFreeSpaceEx function returns correct values for all volumes, including those
	// that are greater than 2 gigabytes."
	__int64 free_space;
	ULARGE_INTEGER total, free, used;
	if (!GetDiskFreeSpaceEx(buf, &free, &total, &used))
		_f_throw_win32();
	// Casting this way allows sizes of up to 2,097,152 gigabytes:
	free_space = (__int64)((unsigned __int64)(aGetFreeSpace ? free.QuadPart : total.QuadPart)
		/ (1024*1024));

	_f_return(free_space);
}



ResultType DriveLock(TCHAR aDriveLetter, bool aLockIt); // Forward declaration for BIF_Drive.

BIF_DECL(BIF_Drive)
{
	BuiltInFunctionID drive_cmd = _f_callee_id;

	LPTSTR aValue = ParamIndexToOptionalString(0, _f_number_buf);

	bool successful = false;
	TCHAR path[MAX_PATH]; // MAX_PATH vs. T_MAX_PATH because SetVolumeLabel() can't seem to make use of long paths.
	size_t path_length;

	// Notes about the below macro:
	// - It adds a backslash to the contents of the path variable because certain API calls or OS versions
	//   might require it.
	// - It is used by both Drive() and DriveGet().
	// - Leave space for the backslash in case its needed.
	#define DRIVE_SET_PATH \
		tcslcpy(path, aValue, _countof(path) - 1);\
		path_length = _tcslen(path);\
		if (path_length && path[path_length - 1] != '\\')\
			path[path_length++] = '\\';

	switch(drive_cmd)
	{
	case FID_DriveLock:
	case FID_DriveUnlock:
		successful = DriveLock(*aValue, drive_cmd == FID_DriveLock);
		break;

	case FID_DriveEject:
	case FID_DriveRetract:
	{
		// Don't do DRIVE_SET_PATH in this case since trailing backslash is not wanted.
		// It seems best not to do the below check since:
		// 1) aValue usually lacks a trailing backslash, which might prevent DriveGetType() from working on some OSes.
		// 2) Eject (and more rarely, retract) works on some other drive types.
		// 3) CreateFile or DeviceIoControl will simply fail or have no effect if the drive isn't of the right type.
		//if (GetDriveType(aValue) != DRIVE_CDROM)
		//	_f_throw(ERR_FAILED);
		BOOL retract = (drive_cmd == FID_DriveRetract);
		TCHAR path[] { '\\', '\\', '.', '\\', 0, ':', '\0', '\0' };
		if (*aValue)
		{
			// Testing showed that a Volume GUID of the form \\?\Volume{...} will work even when
			// the drive isn't mapped to a drive letter.
			if (cisalpha(aValue[0]) && (!aValue[1] || aValue[1] == ':' && (!aValue[2] || aValue[2] == '\\' && !aValue[3])))
			{
				path[4] = aValue[0];
				aValue = path;
			}
		}
		else // When drive is omitted, operate upon the first CD/DVD drive.
		{
			path[6] = '\\'; // GetDriveType technically requires a slash, although it may work without.
			// Testing with mciSendString() while changing or removing drive letters showed that
			// its "default" drive is really just the first drive found in alphabetical order,
			// which is also the most obvious/intuitive choice.
			for (TCHAR c = 'A'; ; ++c)
			{
				path[4] = c;
				if (GetDriveType(path) == DRIVE_CDROM)
					break;
				if (c == 'Z')
					_f_throw(ERR_FAILED); // No CD/DVD drive found with a drive letter.  
			}
			path[6] = '\0'; // Remove the trailing slash for CreateFile to open the volume.
			aValue = path;
		}
		// Testing indicates neither this method nor the MCI method work with mapped drives or UNC paths.
		// That makes sense when one considers that the following opens the *volume*, whereas a network
		// share would correspond to a directory; i.e. this needs "\\.\D:" and not "\\.\D:\".
		HANDLE hVol = CreateFile(aValue, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
		if (hVol == INVALID_HANDLE_VALUE)
			break;
		DWORD unused;
		successful = DeviceIoControl(hVol, retract ? IOCTL_STORAGE_LOAD_MEDIA : IOCTL_STORAGE_EJECT_MEDIA, NULL, 0, NULL, 0, &unused, NULL);
		CloseHandle(hVol);
		break;
	}

	case FID_DriveSetLabel: // Note that it is possible and allowed for the new label to be blank.
		DRIVE_SET_PATH
		successful = SetVolumeLabel(path, ParamIndexToOptionalString(1, _f_number_buf)); // _f_number_buf can potentially be used by both parameters, since aValue is copied into path above.
		break;

	} // switch()

	if (!successful)
		_f_throw_win32();
	_f_return_empty;
}



ResultType DriveLock(TCHAR aDriveLetter, bool aLockIt)
{
	HANDLE hdevice;
	DWORD unused;
	BOOL result;
	TCHAR filename[64];
	_stprintf(filename, _T("\\\\.\\%c:"), aDriveLetter);
	// FILE_READ_ATTRIBUTES is not enough; it yields "Access Denied" error.  So apparently all or part
	// of the sub-attributes in GENERIC_READ are needed.  An MSDN example implies that GENERIC_WRITE is
	// only needed for GetDriveType() == DRIVE_REMOVABLE drives, and maybe not even those when all we
	// want to do is lock/unlock the drive (that example did quite a bit more).  In any case, research
	// indicates that all CD/DVD drives are ever considered DRIVE_CDROM, not DRIVE_REMOVABLE.
	// Due to this and the unlikelihood that GENERIC_WRITE is ever needed anyway, GetDriveType() is
	// not called for the purpose of conditionally adding the GENERIC_WRITE attribute.
	hdevice = CreateFile(filename, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
	if (hdevice == INVALID_HANDLE_VALUE)
		return FAIL;
	PREVENT_MEDIA_REMOVAL pmr;
	pmr.PreventMediaRemoval = aLockIt;
	result = DeviceIoControl(hdevice, IOCTL_STORAGE_MEDIA_REMOVAL, &pmr, sizeof(PREVENT_MEDIA_REMOVAL)
		, NULL, 0, &unused, NULL);
	CloseHandle(hdevice);
	return result ? OK : FAIL;
}



BIF_DECL(BIF_DriveGet)
{
	BuiltInFunctionID drive_get_cmd = _f_callee_id;

	// A parameter is mandatory for some, but optional for others:
	LPTSTR aValue = ParamIndexToOptionalString(0, _f_number_buf);
	if (!*aValue && drive_get_cmd != FID_DriveGetList && drive_get_cmd != FID_DriveGetStatusCD)
		_f_throw_value(ERR_PARAM1_MUST_NOT_BE_BLANK);
	
	if (drive_get_cmd == FID_DriveGetCapacity || drive_get_cmd == FID_DriveGetSpaceFree)
	{
		DriveSpace(aResultToken, aValue, drive_get_cmd == FID_DriveGetSpaceFree);
		return;
	}
	
	TCHAR path[T_MAX_PATH]; // T_MAX_PATH vs. MAX_PATH, though testing indicates only GetDriveType() supports long paths.
	size_t path_length;

	switch(drive_get_cmd)
	{

	case FID_DriveGetList:
	{
		UINT drive_type;
		#define ALL_DRIVE_TYPES 256
		if (!*aValue) drive_type = ALL_DRIVE_TYPES;
		else if (!_tcsicmp(aValue, _T("CDROM"))) drive_type = DRIVE_CDROM;
		else if (!_tcsicmp(aValue, _T("Removable"))) drive_type = DRIVE_REMOVABLE;
		else if (!_tcsicmp(aValue, _T("Fixed"))) drive_type = DRIVE_FIXED;
		else if (!_tcsicmp(aValue, _T("Network"))) drive_type = DRIVE_REMOTE;
		else if (!_tcsicmp(aValue, _T("RAMDisk"))) drive_type = DRIVE_RAMDISK;
		else if (!_tcsicmp(aValue, _T("Unknown"))) drive_type = DRIVE_UNKNOWN;
		else
			goto invalid_parameter;

		TCHAR found_drives[32];  // Need room for all 26 possible drive letters.
		int found_drives_count;
		UCHAR letter;
		TCHAR buf[128], *buf_ptr;

		for (found_drives_count = 0, letter = 'A'; letter <= 'Z'; ++letter)
		{
			buf_ptr = buf;
			*buf_ptr++ = letter;
			*buf_ptr++ = ':';
			*buf_ptr++ = '\\';
			*buf_ptr = '\0';
			UINT this_type = GetDriveType(buf);
			if (this_type == drive_type || (drive_type == ALL_DRIVE_TYPES && this_type != DRIVE_NO_ROOT_DIR))
				found_drives[found_drives_count++] = letter;  // Store just the drive letters.
		}
		found_drives[found_drives_count] = '\0';  // Terminate the string of found drive letters.
		// An empty list should not be flagged as failure, even for FIXED drive_type.
		// For example, when booting Windows PE from a REMOVABLE drive, it mounts a RAMDISK
		// drive but there may be no FIXED drives present.
		//if (!*found_drives)
		//	goto error;
		_f_return(found_drives);
	}

	case FID_DriveGetFilesystem:
	case FID_DriveGetLabel:
	case FID_DriveGetSerial:
	{
		TCHAR volume_name[256];
		TCHAR file_system[256];
		DRIVE_SET_PATH
		DWORD serial_number, max_component_length, file_system_flags;
		if (!GetVolumeInformation(path, volume_name, _countof(volume_name) - 1, &serial_number, &max_component_length
			, &file_system_flags, file_system, _countof(file_system) - 1))
			goto error;
		switch(drive_get_cmd)
		{
		case FID_DriveGetFilesystem: _f_return(file_system);
		case FID_DriveGetLabel:      _f_return(volume_name);
		case FID_DriveGetSerial:     _f_return(serial_number);
		}
		break;
	}

	case FID_DriveGetType:
	{
		DRIVE_SET_PATH
		switch (GetDriveType(path))
		{
		case DRIVE_UNKNOWN:   _f_return_p(_T("Unknown"));
		case DRIVE_REMOVABLE: _f_return_p(_T("Removable"));
		case DRIVE_FIXED:     _f_return_p(_T("Fixed"));
		case DRIVE_REMOTE:    _f_return_p(_T("Network"));
		case DRIVE_CDROM:     _f_return_p(_T("CDROM"));
		case DRIVE_RAMDISK:   _f_return_p(_T("RAMDisk"));
		default: // DRIVE_NO_ROOT_DIR
			_f_return_empty;
		}
		break;
	}

	case FID_DriveGetStatus:
	{
		DRIVE_SET_PATH
		DWORD sectors_per_cluster, bytes_per_sector, free_clusters, total_clusters;
		switch (GetDiskFreeSpace(path, &sectors_per_cluster, &bytes_per_sector, &free_clusters, &total_clusters)
			? ERROR_SUCCESS : GetLastError())
		{
		case ERROR_SUCCESS:        _f_return_p(_T("Ready"));
		case ERROR_PATH_NOT_FOUND: _f_return_p(_T("Invalid"));
		case ERROR_NOT_READY:      _f_return_p(_T("NotReady"));
		case ERROR_WRITE_PROTECT:  _f_return_p(_T("ReadOnly"));
		default:                   _f_return_p(_T("Unknown"));
		}
		break;
	}

	case FID_DriveGetStatusCD:
		// Don't do DRIVE_SET_PATH in this case since trailing backslash might prevent it from
		// working on some OSes.
		// It seems best not to do the below check since:
		// 1) aValue usually lacks a trailing backslash so that it will work correctly with "open c: type cdaudio".
		//    That lack might prevent DriveGetType() from working on some OSes.
		// 2) It's conceivable that tray eject/retract might work on certain types of drives even though
		//    they aren't of type DRIVE_CDROM.
		// 3) One or both of the calls to mciSendString() will simply fail if the drive isn't of the right type.
		//if (GetDriveType(aValue) != DRIVE_CDROM) // Testing reveals that the below method does not work on Network CD/DVD drives.
		//	_f_throw(ERR_FAILED);
		TCHAR mci_string[256], status[128];
		// Note that there is apparently no way to determine via mciSendString() whether the tray is ejected
		// or not, since "open" is returned even when the tray is closed but there is no media.
		if (!*aValue) // When drive is omitted, operate upon default CD/DVD drive.
		{
			if (mciSendString(_T("status cdaudio mode"), status, _countof(status), NULL)) // Error.
				goto failed;
		}
		else // Operate upon a specific drive letter.
		{
			sntprintf(mci_string, _countof(mci_string), _T("open %s type cdaudio alias cd wait shareable"), aValue);
			if (mciSendString(mci_string, NULL, 0, NULL)) // Error.
				goto failed;
			MCIERROR error = mciSendString(_T("status cd mode"), status, _countof(status), NULL);
			mciSendString(_T("close cd wait"), NULL, 0, NULL);
			if (error)
				goto failed;
		}
		// Otherwise, success:
		_f_return(status);

	} // switch()

failed: // Failure where GetLastError() isn't relevant.
	_f_throw(ERR_FAILED);

error: // Win32 error
	_f_throw_win32();

invalid_parameter:
	_f_throw_param(0);
}



/////////////////////
// Sound Functions //
/////////////////////


#pragma region Sound support functions - Vista and later

enum class SoundControlType
{
	Volume,
	Mute,
	Name,
	IID
};

LPWSTR SoundDeviceGetName(IMMDevice *aDev)
{
	IPropertyStore *store;
	PROPVARIANT prop;
	if (SUCCEEDED(aDev->OpenPropertyStore(STGM_READ, &store)))
	{
		if (FAILED(store->GetValue(PKEY_Device_FriendlyName, &prop)))
			prop.pwszVal = nullptr;
		store->Release();
		return prop.pwszVal;
	}
	return nullptr;
}


HRESULT SoundSetGet_GetDevice(LPTSTR aDeviceString, IMMDevice *&aDevice)
{
	IMMDeviceEnumerator *deviceEnum;
	IMMDeviceCollection *devices;
	HRESULT hr;

	aDevice = NULL;

	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&deviceEnum);
	if (SUCCEEDED(hr))
	{
		if (!*aDeviceString)
		{
			// Get default playback device.
			hr = deviceEnum->GetDefaultAudioEndpoint(eRender, eConsole, &aDevice);
		}
		else
		{
			// Parse Name:Index, Name or Index.
			int target_index = 0;
			LPWSTR target_name = L"";
			size_t target_name_length = 0;
			LPTSTR delim = _tcsrchr(aDeviceString, ':');
			if (delim)
			{
				target_index = ATOI(delim + 1) - 1;
				target_name = aDeviceString;
				target_name_length = delim - aDeviceString;
			}
			else if (IsNumeric(aDeviceString, FALSE, FALSE))
			{
				target_index = ATOI(aDeviceString) - 1;
			}
			else
			{
				target_name = aDeviceString;
				target_name_length = _tcslen(aDeviceString);
			}

			// Enumerate devices; include unplugged devices so that indices don't change when a device is plugged in.
			hr = deviceEnum->EnumAudioEndpoints(eAll, DEVICE_STATE_ACTIVE | DEVICE_STATE_UNPLUGGED, &devices);
			if (SUCCEEDED(hr))
			{
				if (target_name_length)
				{
					IMMDevice *dev;
					for (UINT u = 0; SUCCEEDED(hr = devices->Item(u, &dev)); ++u)
					{
						if (LPWSTR dev_name = SoundDeviceGetName(dev))
						{
							if (!_wcsnicmp(dev_name, target_name, target_name_length)
								&& target_index-- == 0)
							{
								CoTaskMemFree(dev_name);
								aDevice = dev;
								break;
							}
							CoTaskMemFree(dev_name);
						}
						dev->Release();
					}
				}
				else
					hr = devices->Item((UINT)target_index, &aDevice);
				devices->Release();
			}
		}
		deviceEnum->Release();
	}
	return hr;
}


struct SoundComponentSearch
{
	// Parameters of search:
	GUID target_iid;
	int target_instance;
	SoundControlType target_control;
	TCHAR target_name[128]; // Arbitrary; probably larger than any real component name.
	// Internal use/results:
	IUnknown *control;
	LPWSTR name; // Valid only when target_control == SoundControlType::Name.
	int count;
	// Internal use:
	DataFlow data_flow;
	bool ignore_remaining_subunits;
};


void SoundConvertComponent(LPTSTR aBuf, SoundComponentSearch &aSearch)
{
	if (IsNumeric(aBuf) == PURE_INTEGER)
	{
		*aSearch.target_name = '\0';
		aSearch.target_instance = ATOI(aBuf);
	}
	else
	{
		tcslcpy(aSearch.target_name, aBuf, _countof(aSearch.target_name));
		LPTSTR colon_pos = _tcsrchr(aSearch.target_name, ':');
		aSearch.target_instance = colon_pos ? ATOI(colon_pos + 1) : 1;
		if (colon_pos)
			*colon_pos = '\0';
	}
}


bool SoundSetGet_FindComponent(IPart *aRoot, SoundComponentSearch &aSearch)
{
	HRESULT hr;
	IPartsList *parts;
	IPart *part;
	UINT part_count;
	PartType part_type;
	LPWSTR part_name;
	
	bool check_name = *aSearch.target_name;

	if (aSearch.data_flow == In)
		hr = aRoot->EnumPartsIncoming(&parts);
	else
		hr = aRoot->EnumPartsOutgoing(&parts);
	
	if (FAILED(hr))
		return NULL;

	if (FAILED(parts->GetCount(&part_count)))
		part_count = 0;
	
	for (UINT i = 0; i < part_count; ++i)
	{
		if (FAILED(parts->GetPart(i, &part)))
			continue;

		if (SUCCEEDED(part->GetPartType(&part_type)))
		{
			if (part_type == Connector)
			{
				if (   part_count == 1 // Ignore Connectors with no Subunits of their own.
					&& (!check_name || (SUCCEEDED(part->GetName(&part_name)) && !_wcsicmp(part_name, aSearch.target_name)))   )
				{
					if (++aSearch.count == aSearch.target_instance)
					{
						switch (aSearch.target_control)
						{
						case SoundControlType::IID:
							// Permit retrieving the IPart or IConnector itself.  Since there may be
							// multiple connected Subunits (and they can be enumerated or retrieved
							// via the Connector IPart), this is only done for the Connector.
							part->QueryInterface(aSearch.target_iid, (void **)&aSearch.control);
							break;
						case SoundControlType::Name:
							// check_name would typically be false in this case, since a value of true
							// would mean the caller already knows the component's name.
							part->GetName(&aSearch.name);
							break;
						}
						part->Release();
						parts->Release();
						return true;
					}
				}
			}
			else // Subunit
			{
				// Recursively find the Connector nodes linked to this part.
				if (SoundSetGet_FindComponent(part, aSearch))
				{
					// A matching connector part has been found with this part as one of the nodes used
					// to reach it.  Therefore, if this part supports the requested control interface,
					// it can in theory be used to control the component.  An example path might be:
					//    Output < Master Mute < Master Volume < Sum < Mute < Volume < CD Audio
					// Parts are considered from right to left, as we return from recursion.
					if (!aSearch.control && !aSearch.ignore_remaining_subunits)
					{
						// Query this part for the requested interface and let caller check the result.
						part->Activate(CLSCTX_ALL, aSearch.target_iid, (void **)&aSearch.control);

						// If this subunit has siblings, ignore any controls further up the line
						// as they're likely shared by other components (i.e. master controls).
						if (part_count > 1)
							aSearch.ignore_remaining_subunits = true;
					}
					part->Release();
					parts->Release();
					return true;
				}
			}
		}

		part->Release();
	}

	parts->Release();
	return false;
}


bool SoundSetGet_FindComponent(IMMDevice *aDevice, SoundComponentSearch &aSearch)
{
	IDeviceTopology *topo;
	IConnector *conn, *conn_to;
	IPart *part;

	aSearch.count = 0;
	aSearch.control = nullptr;
	aSearch.name = nullptr;
	aSearch.ignore_remaining_subunits = false;
	
	if (SUCCEEDED(aDevice->Activate(__uuidof(IDeviceTopology), CLSCTX_ALL, NULL, (void**)&topo)))
	{
		if (SUCCEEDED(topo->GetConnector(0, &conn)))
		{
			if (SUCCEEDED(conn->GetDataFlow(&aSearch.data_flow))
			 && SUCCEEDED(conn->GetConnectedTo(&conn_to)))
			{
				if (SUCCEEDED(conn_to->QueryInterface(&part)))
				{
					// Search; the result is stored in the search struct.
					SoundSetGet_FindComponent(part, aSearch);
					part->Release();
				}
				conn_to->Release();
			}
			conn->Release();
		}
		topo->Release();
	}

	return aSearch.count == aSearch.target_instance;
}

#pragma endregion


BIF_DECL(BIF_Sound)
{
	LPTSTR aSetting = nullptr;
	SoundComponentSearch search;
	if (_f_callee_id >= FID_SoundSetVolume)
	{
		search.target_control = SoundControlType(_f_callee_id - FID_SoundSetVolume);
		aSetting = ParamIndexToString(0, _f_number_buf);
		++aParam;
		--aParamCount;
	}
	else
		search.target_control = SoundControlType(_f_callee_id);

	switch (search.target_control)
	{
	case SoundControlType::Volume:
		search.target_iid = __uuidof(IAudioVolumeLevel);
		break;
	case SoundControlType::Mute:
		search.target_iid = __uuidof(IAudioMute);
		break;
	case SoundControlType::IID:
		LPTSTR iid; iid = ParamIndexToString(0);
		if (*iid != '{' || FAILED(CLSIDFromString(iid, &search.target_iid)))
			_f_throw_param(0);
		++aParam;
		--aParamCount;
		break;
	}

	_f_param_string_opt(aComponent, 0);
	_f_param_string_opt(aDevice, 1);

#define SOUND_MODE_IS_SET aSetting // Boolean: i.e. if it's not NULL, the mode is "SET".
	_f_set_retval_p(_T(""), 0); // Set default.

	float setting_scalar;
	if (SOUND_MODE_IS_SET)
	{
		setting_scalar = (float)(ATOF(aSetting) / 100);
		if (setting_scalar < -1)
			setting_scalar = -1;
		else if (setting_scalar > 1)
			setting_scalar = 1;
	}

	// Does user want to adjust the current setting by a certain amount?
	bool adjust_current_setting = aSetting && (*aSetting == '-' || *aSetting == '+');

	IMMDevice *device;
	HRESULT hr;

	hr = SoundSetGet_GetDevice(aDevice, device);
	if (FAILED(hr))
		_f_throw(ERR_SOUND_DEVICE, ErrorPrototype::Target);

	LPCTSTR errorlevel = NULL;
	float result_float;
	BOOL result_bool;

	if (!*aComponent) // Component is Master (omitted).
	{
		if (search.target_control == SoundControlType::IID)
		{
			void *result_ptr = nullptr;
			if (FAILED(device->QueryInterface(search.target_iid, (void **)&result_ptr)))
				device->Activate(search.target_iid, CLSCTX_ALL, NULL, &result_ptr);
			// For consistency with ComObjQuery, the result is returned even on failure.
			aResultToken.SetValue((UINT_PTR)result_ptr);
		}
		else if (search.target_control == SoundControlType::Name)
		{
			if (auto name = SoundDeviceGetName(device))
			{
				aResultToken.Return(name);
				CoTaskMemFree(name);
			}
		}
		else
		{
			// For Master/Speakers, use the IAudioEndpointVolume interface.  Some devices support master
			// volume control, but do not actually have a volume subunit (so the other method would fail).
			IAudioEndpointVolume *aev;
			hr = device->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, NULL, (void**)&aev);
			if (SUCCEEDED(hr))
			{
				if (search.target_control == SoundControlType::Volume)
				{
					if (!SOUND_MODE_IS_SET || adjust_current_setting)
					{
						hr = aev->GetMasterVolumeLevelScalar(&result_float);
					}

					if (SUCCEEDED(hr))
					{
						if (SOUND_MODE_IS_SET)
						{
							if (adjust_current_setting)
							{
								setting_scalar += result_float;
								if (setting_scalar > 1)
									setting_scalar = 1;
								else if (setting_scalar < 0)
									setting_scalar = 0;
							}
							hr = aev->SetMasterVolumeLevelScalar(setting_scalar, NULL);
						}
						else
						{
							result_float *= 100.0;
						} 
					}
				}
				else // Mute.
				{
					if (!SOUND_MODE_IS_SET || adjust_current_setting)
					{
						hr = aev->GetMute(&result_bool);
					}
					if (SOUND_MODE_IS_SET && SUCCEEDED(hr))
					{
						hr = aev->SetMute(adjust_current_setting ? !result_bool : setting_scalar > 0, NULL);
					}
				}
				aev->Release();
			}
		}
	}
	else
	{
		SoundConvertComponent(aComponent, search);
		
		if (!SoundSetGet_FindComponent(device, search))
		{
			errorlevel = ERR_SOUND_COMPONENT;
		}
		else if (search.target_control == SoundControlType::IID)
		{
			aResultToken.SetValue((UINT_PTR)search.control);
			search.control = nullptr; // Don't release it.
		}
		else if (search.target_control == SoundControlType::Name)
		{
			aResultToken.Return(search.name);
		}
		else if (!search.control)
		{
			errorlevel = ERR_SOUND_CONTROLTYPE;
		}
		else if (search.target_control == SoundControlType::Volume)
		{
			IAudioVolumeLevel *avl = (IAudioVolumeLevel *)search.control;
				
			UINT channel_count = 0;
			if (FAILED(avl->GetChannelCount(&channel_count)))
				goto control_fail;
				
			float *level = (float *)_alloca(sizeof(float) * 3 * channel_count);
			float *level_min = level + channel_count;
			float *level_range = level_min + channel_count;
			float f, db, min_db, max_db, max_level = 0;

			for (UINT i = 0; i < channel_count; ++i)
			{
				if (FAILED(avl->GetLevel(i, &db)) ||
					FAILED(avl->GetLevelRange(i, &min_db, &max_db, &f)))
					goto control_fail;
				// Convert dB to scalar.
				level_min[i] = (float)qmathExp10(min_db/20);
				level_range[i] = (float)qmathExp10(max_db/20) - level_min[i];
				// Compensate for differing level ranges. (No effect if range is -96..0 dB.)
				level[i] = ((float)qmathExp10(db/20) - level_min[i]) / level_range[i];
				// Windows reports the highest level as the overall volume.
				if (max_level < level[i])
					max_level = level[i];
			}

			if (SOUND_MODE_IS_SET)
			{
				if (adjust_current_setting)
				{
					setting_scalar += max_level;
					if (setting_scalar > 1)
						setting_scalar = 1;
					else if (setting_scalar < 0)
						setting_scalar = 0;
				}

				for (UINT i = 0; i < channel_count; ++i)
				{
					f = setting_scalar;
					if (max_level)
						f *= (level[i] / max_level); // Preserve balance.
					// Compensate for differing level ranges.
					f = level_min[i] + f * level_range[i];
					// Convert scalar to dB.
					level[i] = 20 * (float)qmathLog10(f);
				}

				hr = avl->SetLevelAllChannels(level, channel_count, NULL);
			}
			else
			{
				result_float = max_level * 100;
			}
		}
		else if (search.target_control == SoundControlType::Mute)
		{
			IAudioMute *am = (IAudioMute *)search.control;

			if (!SOUND_MODE_IS_SET || adjust_current_setting)
			{
				hr = am->GetMute(&result_bool);
			}
			if (SOUND_MODE_IS_SET && SUCCEEDED(hr))
			{
				hr = am->SetMute(adjust_current_setting ? !result_bool : setting_scalar > 0, NULL);
			}
		}
control_fail:
		if (search.control)
			search.control->Release();
		if (search.name)
			CoTaskMemFree(search.name);
	}

	device->Release();
	
	if (errorlevel)
		_f_throw(errorlevel, ErrorPrototype::Target);
	if (FAILED(hr)) // This would be rare and unexpected.
		_f_throw_win32(hr);

	if (SOUND_MODE_IS_SET)
		return;

	switch (search.target_control)
	{
	case SoundControlType::Volume: aResultToken.Return(result_float); break;
	case SoundControlType::Mute: aResultToken.Return(result_bool); break;
	}
}



ResultType Line::SoundPlay(LPTSTR aFilespec, bool aSleepUntilDone)
{
	LPTSTR cp = omit_leading_whitespace(aFilespec);
	if (*cp == '*')
	{
		// ATOU() returns 0xFFFFFFFF for -1, which is relied upon to support the -1 sound.
		// Even if there are no enabled sound devices, MessageBeep() indicates success
		// (tested on Windows 10).  It's hard to imagine how the script would handle a
		// failure anyway, so just omit error checking.
		MessageBeep(ATOU(cp + 1));
		return OK;
	}
	// See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/multimed/htm/_win32_play.asp
	// for some documentation mciSendString() and related.
	// MAX_PATH note: There's no chance this API supports long paths even on Windows 10.
	// Research indicates it limits paths to 127 chars (not even MAX_PATH), but since there's
	// no apparent benefit to reducing it, we'll keep this size to ensure backward-compatibility.
	// Note that using a relative path does not help, but using short (8.3) names does.
	TCHAR buf[MAX_PATH * 2]; // Allow room for filename and commands.
	mciSendString(_T("status ") SOUNDPLAY_ALIAS _T(" mode"), buf, _countof(buf), NULL);
	if (*buf) // "playing" or "stopped" (so close it before trying to re-open with a new aFilespec).
		mciSendString(_T("close ") SOUNDPLAY_ALIAS, NULL, 0, NULL);
	sntprintf(buf, _countof(buf), _T("open \"%s\" alias ") SOUNDPLAY_ALIAS, aFilespec);
	if (mciSendString(buf, NULL, 0, NULL)) // Failure.
		goto error;
	g_SoundWasPlayed = true;  // For use by Script's destructor.
	if (mciSendString(_T("play ") SOUNDPLAY_ALIAS, NULL, 0, NULL)) // Failure.
		goto error;
	// Otherwise, the sound is now playing.
	if (!aSleepUntilDone)
		return OK;
	// Otherwise, caller wants us to wait until the file is done playing.  To allow our app to remain
	// responsive during this time, use a loop that checks our message queue:
	// Older method: "mciSendString("play " SOUNDPLAY_ALIAS " wait", NULL, 0, NULL)"
	for (;;)
	{
		mciSendString(_T("status ") SOUNDPLAY_ALIAS _T(" mode"), buf, _countof(buf), NULL);
		if (!*buf) // Probably can't happen given the state we're in.
			break;
		if (!_tcscmp(buf, _T("stopped"))) // The sound is done playing.
		{
			mciSendString(_T("close ") SOUNDPLAY_ALIAS, NULL, 0, NULL);
			break;
		}
		// Sleep a little longer than normal because I'm not sure how much overhead
		// and CPU utilization the above incurs:
		MsgSleep(20);
	}
	return OK;

error:
	return LineError(ERR_FAILED, FAIL_OR_OK);
}



///////////////////////
// Working Directory //
///////////////////////


ResultType SetWorkingDir(LPTSTR aNewDir)
// Throws a script runtime exception on failure, but only if the script has begun runtime execution.
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
		return FAIL;
	// Otherwise, the change to the working directory succeeded.

	// Other than during program startup, this should be the only place where the official
	// working dir can change.  The exception is FileSelect(), which changes the working
	// dir as the user navigates from folder to folder.  However, the whole purpose of
	// maintaining g_WorkingDir is to workaround that very issue.
	if (g_script.mIsReadyToExecute) // Callers want this done only during script runtime.
		UpdateWorkingDir(aNewDir);
	return OK;
}



void UpdateWorkingDir(LPTSTR aNewDir)
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



////////////////////
// File Functions //
////////////////////


BIF_DECL(BIF_FileSelect)
// Since other script threads can interrupt this command while it's running, it's important that
// this command not refer to sArgDeref[] and sArgVar[] anytime after an interruption becomes possible.
// This is because an interrupting thread usually changes the values to something inappropriate for this thread.
{
	if (g_nFileDialogs >= MAX_FILEDIALOGS)
	{
		// Have a maximum to help prevent runaway hotkeys due to key-repeat feature, etc.
		_f_throw(_T("The maximum number of File Dialogs has been reached."));
	}
	_f_param_string_opt(aOptions, 0);
	_f_param_string_opt(aWorkingDir, 1);
	_f_param_string_opt(aGreeting, 2);
	_f_param_string_opt(aFilter, 3);
	
	LPCTSTR default_file_name = _T("");

	TCHAR working_dir[MAX_PATH]; // Using T_MAX_PATH vs. MAX_PATH did not help on Windows 10.0.16299 (see below).
	if (!aWorkingDir || !*aWorkingDir)
		*working_dir = '\0';
	else
	{
		// Compress the path if possible to support longer paths.  Without this, any path longer
		// than MAX_PATH would be ignored, presumably because the dialog, as part of the shell,
		// does not support long paths.  Surprisingly, although Windows 10 long path awareness
		// does not allow us to pass a long path for working_dir, it does affect whether the long
		// path is used in the address bar and returned filenames.
		if (_tcslen(aWorkingDir) >= MAX_PATH)
			GetShortPathName(aWorkingDir, working_dir, _countof(working_dir));
		else
			tcslcpy(working_dir, aWorkingDir, _countof(working_dir));
		// v1.0.43.10: Support CLSIDs such as:
		//   My Computer  ::{20d04fe0-3aea-1069-a2d8-08002b30309d}
		//   My Documents ::{450d8fba-ad25-11d0-98a8-0800361b1103}
		// Also support optional subdirectory appended to the CLSID.
		// Neither SetCurrentDirectory() nor GetFileAttributes() directly supports CLSIDs, so rely on other means
		// to detect whether a CLSID ends in a directory vs. filename.
		bool is_directory, is_clsid;
		if (is_clsid = !_tcsncmp(working_dir, _T("::{"), 3))
		{
			LPTSTR end_brace;
			if (end_brace = _tcschr(working_dir, '}'))
				is_directory = !end_brace[1] // First '}' is also the last char in string, so it's naked CLSID (so assume directory).
					|| working_dir[_tcslen(working_dir) - 1] == '\\'; // Or path ends in backslash.
			else // Badly formatted clsid.
				is_directory = true; // Arbitrary default due to rarity.
		}
		else // Not a CLSID.
		{
			DWORD attr = GetFileAttributes(working_dir);
			is_directory = (attr != 0xFFFFFFFF) && (attr & FILE_ATTRIBUTE_DIRECTORY);
		}
		if (!is_directory)
		{
			// Above condition indicates it's either an existing file that's not a folder, or a nonexistent
			// folder/filename.  In either case, it seems best to assume it's a file because the user may want
			// to provide a default SAVE filename, and it would be normal for such a file not to already exist.
			LPTSTR last_backslash;
			if (last_backslash = _tcsrchr(working_dir, '\\'))
			{
				default_file_name = last_backslash + 1;
				*last_backslash = '\0'; // Make the working directory just the file's path.
			}
			else // The entire working_dir string is the default file (unless this is a clsid).
				if (!is_clsid)
					default_file_name = working_dir; // This signals it to use the default directory.
				//else leave working_dir set to the entire clsid string in case it's somehow valid.
		}
		// else it is a directory, so just leave working_dir set as it was initially.
	}

	TCHAR pattern[1024];
	*pattern = '\0'; // Set default.
	if (*aFilter)
	{
		LPTSTR pattern_start = _tcschr(aFilter, '(');
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
			tcslcpy(pattern, aFilter, _countof(pattern));
	}
	UINT filter_count = 0;
	COMDLG_FILTERSPEC filters[2];
	if (*pattern)
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
		filters[0].pszName = aFilter;
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
	switch (ctoupper(*aOptions))
	{
	case 'D':
		++aOptions;
		flags |= FOS_PICKFOLDERS;
		if (*aFilter)
			_f_throw_value(ERR_PARAM4_MUST_BE_BLANK);
		filter_count = 0;
		break;
	case 'M':  // Multi-select.
		++aOptions;
		flags |= FOS_ALLOWMULTISELECT;
		break;
	case 'S': // Have a "Save" button rather than an "Open" button.
		++aOptions;
		always_use_save_dialog = true;
		break;
	}

	TCHAR greeting[1024];
	if (aGreeting && *aGreeting)
		tcslcpy(greeting, aGreeting, _countof(greeting));
	else
		// Use a more specific title so that the dialogs of different scripts can be distinguished
		// from one another, which may help script automation in rare cases:
		sntprintf(greeting, _countof(greeting), _T("Select %s - %s")
			, (flags & FOS_PICKFOLDERS) ? _T("Folder") : _T("File"), g_script.DefaultDialogTitle());

	int options = ATOI(aOptions);
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
		_f_throw_win32(hr);

	pfd->SetOptions(flags);
	pfd->SetTitle(greeting);
	if (filter_count)
		pfd->SetFileTypes(filter_count, filters);
	pfd->SetFileName(default_file_name);

	if (*working_dir && default_file_name != working_dir)
	{
		IShellItem *psi;
		if (SUCCEEDED(SHCreateItemFromParsingName(working_dir, nullptr, IID_PPV_ARGS(&psi))))
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
		_f_return(files);
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
		// Currently assuming the user canceled, otherwise this would tell us whether an error
		// occurred vs. the user canceling: if (result != HRESULT_FROM_WIN32(ERROR_CANCELLED))
	pfd->Release();
}



// As of 2019-09-29, noinline reduces code size by over 20KB on VC++ 2019.
// Prior to merging Util_CreateDir with this, it wasn't inlined.
DECLSPEC_NOINLINE
bool Line::FileCreateDir(LPTSTR aDirSpec, LPTSTR aCanModifyDirSpec)
{
	if (!aDirSpec || !*aDirSpec)
	{
		SetLastError(ERROR_INVALID_PARAMETER);
		return false;
	}

	DWORD attr = GetFileAttributes(aDirSpec);
	if (attr != 0xFFFFFFFF)  // aDirSpec already exists.
	{
		SetLastError(ERROR_ALREADY_EXISTS);
		return (attr & FILE_ATTRIBUTE_DIRECTORY) != 0; // Indicate success if it already exists as a dir.
	}

	// If it has a backslash, make sure all its parent directories exist before we attempt
	// to create this directory:
	LPTSTR last_backslash = _tcsrchr(aDirSpec, '\\');
	if (last_backslash > aDirSpec // v1.0.48.04: Changed "last_backslash" to "last_backslash > aDirSpec" so that an aDirSpec with a leading \ (but no other backslashes), such as \dir, is supported.
		&& last_backslash[-1] != ':') // v1.1.31.00: Don't attempt FileCreateDir("C:") since that's equivalent to either "C:\" or the working directory (which already exists), or FileCreateDir("\\?\C:") since it always fails.
	{
		LPTSTR parent_dir;
		if (aCanModifyDirSpec)
		{
			parent_dir = aDirSpec; // Caller provided a modifiable aDirSpec.
			*last_backslash = '\0'; // Temporarily terminate for parent directory.
		}
		else
		{
			// v1.1.31.00: Allocate a modifiable buffer to be used by all calls (supports long paths).
			parent_dir = (LPTSTR)_alloca((last_backslash - aDirSpec + 1) * sizeof(TCHAR));
			tcslcpy(parent_dir, aDirSpec, last_backslash - aDirSpec + 1); // Omits the last backslash.
		}
		bool exists = FileCreateDir(parent_dir, parent_dir); // Recursively create all needed ancestor directories.
		if (aCanModifyDirSpec)
			*last_backslash = '\\'; // Undo temporary termination.

		// v1.0.44: Fixed ErrorLevel being set to 1 when the specified directory ends in a backslash.  In such cases,
		// two calls were made to CreateDirectory for the same folder: the first without the backslash and then with
		// it.  Since the directory already existed on the second call, ErrorLevel was wrongly set to 1 even though
		// everything succeeded.  So now, when recursion finishes creating all the ancestors of this directory
		// our own layer here does not call CreateDirectory() when there's a trailing backslash because a previous
		// layer already did:
		if (!last_backslash[1] || !exists)
			return exists;
	}

	// The above has recursively created all parent directories of aDirSpec if needed.
	// Now we can create aDirSpec.
	return CreateDirectory(aDirSpec, NULL);
}



ResultType ConvertFileOptions(ResultToken &aResultToken, LPTSTR aOptions, UINT &codepage, bool &translate_crlf_to_lf, unsigned __int64 *pmax_bytes_to_load)
{
	for (LPTSTR next, cp = aOptions; cp && *(cp = omit_leading_whitespace(cp)); cp = next)
	{
		if (*cp == '\n')
		{
			translate_crlf_to_lf = true;
			// Rather than treating "`nxxx" as invalid or ignoring "xxx", let the delimiter be
			// optional for `n.  Treating "`nxxx" and "m1024`n" and "utf-8`n" as invalid would
			// require larger code, and would produce confusing error messages because the `n
			// isn't visible; e.g. "Invalid option. Specifically: utf-8"
			next = cp + 1; 
			continue;
		}
		// \n is included below to allow "m1024`n" and "utf-8`n" (see above).
		next = StrChrAny(cp, _T(" \t\n"));

		switch (ctoupper(*cp))
		{
		case 'M':
			if (pmax_bytes_to_load) // i.e. caller is FileRead.
			{
				*pmax_bytes_to_load = ATOU64(cp + 1); // Relies upon the fact that it ceases conversion upon reaching a space or tab.
				break;
			}
			// Otherwise, fall through to treat it as invalid:
		default:
			TCHAR name[12]; // Large enough for any valid encoding.
			if (next && (next - cp) < _countof(name))
			{
				// Create a temporary null-terminated copy.
				wmemcpy(name, cp, next - cp);
				name[next - cp] = '\0';
				cp = name;
			}
			if (!_tcsicmp(cp, _T("Raw")))
			{
				codepage = -1;
			}
			else
			{
				codepage = Line::ConvertFileEncoding(cp);
				if (codepage == -1 || cisdigit(*cp)) // Require "cp" prefix in FileRead/FileAppend options.
					return aResultToken.ValueError(ERR_INVALID_OPTION, cp);
			}
			break;
		} // switch()
	} // for()
	return OK;
}

BIF_DECL(BIF_FileRead)
{
	_f_param_string(aFilespec, 0);
	_f_param_string_opt(aOptions, 1);

	const DWORD DWORD_MAX = ~0;

	// Set default options:
	bool translate_crlf_to_lf = false;
	unsigned __int64 max_bytes_to_load = ULLONG_MAX; // By default, fail if the file is too large.  See comments near bytes_to_read below.
	UINT codepage = g->Encoding;

	if (!ConvertFileOptions(aResultToken, aOptions, codepage, translate_crlf_to_lf, &max_bytes_to_load))
		return; // It already displayed the error.

	_f_set_retval_p(_T(""), 0); // Set default.

	// It seems more flexible to allow other processes to read and write the file while we're reading it.
	// For example, this allows the file to be appended to during the read operation, which could be
	// desirable, especially it's a very large log file that would take a long time to read.
	// MSDN: "To enable other processes to share the object while your process has it open, use a combination
	// of one or more of [FILE_SHARE_READ, FILE_SHARE_WRITE]."
	HANDLE hfile = CreateFile(aFilespec, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING
		, FILE_FLAG_SEQUENTIAL_SCAN, NULL); // MSDN says that FILE_FLAG_SEQUENTIAL_SCAN will often improve performance
	if (hfile == INVALID_HANDLE_VALUE)      // in cases like these (and it seems best even if max_bytes_to_load was specified).
	{
		aResultToken.SetLastErrorMaybeThrow(true);
		return;
	}

	unsigned __int64 bytes_to_read = GetFileSize64(hfile);
	if (bytes_to_read == ULLONG_MAX) // GetFileSize64() failed.
	{
		aResultToken.SetLastErrorCloseAndMaybeThrow(hfile, true);
		return;
	}
	// In addition to imposing the limit set by the *M option, the following check prevents an error
	// caused by 64 to 32-bit truncation -- that is, a file size of 0x100000001 would be truncated to
	// 0x1, allowing the command to complete even though it should fail.  UPDATE: This check was never
	// sufficient since max_bytes_to_load could exceed DWORD_MAX on x64 (prior to v1.1.16).  It's now
	// checked separately below to try to match the documented behaviour (truncating the data only to
	// the caller-specified limit).
	if (bytes_to_read > max_bytes_to_load) // This is the limit set by the caller.
		bytes_to_read = max_bytes_to_load;
	// Fixed for v1.1.16: Show an error message if the file is larger than DWORD_MAX, otherwise the
	// truncation issue described above could occur.  Reading more than DWORD_MAX could be supported
	// by calling ReadFile() in a loop, but it seems unlikely that a script will genuinely want to
	// do this AND actually be able to allocate a 4GB+ memory block (having 4GB of total free memory
	// is usually not sufficient, perhaps due to memory fragmentation).
#ifdef _WIN64
	if (bytes_to_read > DWORD_MAX)
#else
	// Reserve 2 bytes to avoid integer overflow below.  Although any amount larger than 2GB is almost
	// guaranteed to fail at the malloc stage, that might change if we ever become large address aware.
	if (bytes_to_read > DWORD_MAX - sizeof(wchar_t))
#endif
	{
		CloseHandle(hfile);
		_f_throw_oom; // Using this instead of "File too large." to reduce code size, since this condition is very rare (and malloc succeeding would be even rarer).
	}

	if (!bytes_to_read && codepage != -1) // In RAW mode, always return a zero-byte Buffer.
	{
		aResultToken.SetLastErrorCloseAndMaybeThrow(hfile, false, 0); // Indicate success (a zero-length file results in an empty string).
		return;
	}

	LPBYTE output_buf = (LPBYTE)malloc(size_t(bytes_to_read + (bytes_to_read & 1) + sizeof(wchar_t)));
	if (!output_buf)
	{
		CloseHandle(hfile);
		_f_throw_oom;
	}

	DWORD bytes_actually_read;
	BOOL result = ReadFile(hfile, output_buf, (DWORD)bytes_to_read, &bytes_actually_read, NULL);
	g->LastError = GetLastError();
	CloseHandle(hfile);

	// Upon result==success, bytes_actually_read is not checked against bytes_to_read because it
	// shouldn't be different (result should have set to failure if there was a read error).
	// If it ever is different, a partial read is considered a success since ReadFile() told us
	// that nothing bad happened.

	if (result)
	{
		if (codepage != -1) // Text mode, not "RAW" mode.
		{
			codepage &= CP_AHKCP; // Convert to plain Win32 codepage (remove CP_AHKNOBOM, which has no meaning here).
			bool has_bom;
			if ( (has_bom = (bytes_actually_read >= 2 && output_buf[0] == 0xFF && output_buf[1] == 0xFE)) // UTF-16LE BOM
					|| codepage == CP_UTF16 ) // Covers FileEncoding UTF-16 and FileEncoding UTF-16-RAW.
			{
				#ifndef UNICODE
				#error FileRead UTF-16 to ANSI string not implemented.
				#endif
				LPWSTR text = (LPWSTR)output_buf;
				DWORD length = bytes_actually_read / sizeof(WCHAR);
				if (has_bom)
				{
					// Move the data to eliminate the byte order mark.
					// Seems likely to perform better than allocating new memory and copying to it.
					--length;
					wmemmove(text, text + 1, length);
				}
				text[length] = '\0'; // Ensure text is terminated where indicated.  Two bytes were reserved for this purpose.
				aResultToken.AcceptMem(text, length);
				output_buf = NULL; // Don't free it; caller will take over.
			}
			else
			{
				LPCSTR text = (LPCSTR)output_buf;
				DWORD length = bytes_actually_read;
				if (length >= 3 && output_buf[0] == 0xEF && output_buf[1] == 0xBB && output_buf[2] == 0xBF) // UTF-8 BOM
				{
					codepage = CP_UTF8;
					length -= 3;
					text += 3;
				}
#ifndef UNICODE
				if (codepage == CP_ACP || codepage == GetACP())
				{
					// Avoid any unnecessary conversion or copying by using our malloc'd buffer directly.
					// This should be worth doing since the string must otherwise be converted to UTF-16 and back.
					output_buf[bytes_actually_read] = 0; // Ensure text is terminated where indicated.
					aResultToken.AcceptMem((LPSTR)output_buf, bytes_actually_read);
					output_buf = NULL; // Don't free it; caller will take over.
				}
				else
				#error FileRead non-ACP-ANSI to ANSI string not fully implemented.
#endif
				{
					int wlen = MultiByteToWideChar(codepage, 0, text, length, NULL, 0);
					if (wlen > 0)
					{
						if (!TokenSetResult(aResultToken, NULL, wlen))
						{
							free(output_buf);
							return;
						}
						wlen = MultiByteToWideChar(codepage, 0, text, length, aResultToken.marker, wlen);
						aResultToken.symbol = SYM_STRING;
						aResultToken.marker[wlen] = 0;
						aResultToken.marker_length = wlen;
						if (!wlen)
							result = FALSE;
					}
				}
			}
			if (output_buf) // i.e. it wasn't "claimed" above.
				free(output_buf);
			if (translate_crlf_to_lf && aResultToken.marker_length)
			{
				// Since a larger string is being replaced with a smaller, there's a good chance the 2 GB
				// address limit will not be exceeded by StrReplace even if the file is close to the
				// 1 GB limit as described above:
				StrReplace(aResultToken.marker, _T("\r\n"), _T("\n"), SCS_SENSITIVE, UINT_MAX, -1, NULL, &aResultToken.marker_length);
			}
		}
		else // codepage == -1 ("RAW" mode)
		{
			// Return the buffer to our caller.
			aResultToken.Return(BufferObject::Create(output_buf, bytes_actually_read));
		}
	}
	else
	{
		// ReadFile() failed.  Since MSDN does not document what is in the buffer at this stage, or
		// whether bytes_to_read contains a valid value, it seems best to abort the entire operation
		// rather than try to return partial file contents.  An exception will indicate the failure.
		free(output_buf);
	}

	aResultToken.SetLastErrorMaybeThrow(!result);
}



BIF_DECL(BIF_FileAppend)
{
	size_t aBuf_length;
	_f_param_string(aBuf, 0, &aBuf_length);
	_f_param_string_opt(aFilespec, 1);
	_f_param_string_opt(aOptions, 2);

	IObject *aBuf_obj = ParamIndexToObject(0); // Allow a Buffer-like object.
	if (aBuf_obj)
	{
		size_t ptr;
		GetBufferObjectPtr(aResultToken, aBuf_obj, ptr, aBuf_length);
		if (aResultToken.Exited())
			return;
		aBuf = (LPTSTR)ptr;
	}
	else
		aBuf_length *= sizeof(TCHAR); // Convert to byte count.

	_f_set_retval_p(_T(""), 0); // For all non-throw cases.

	// The below is avoided because want to allow "nothing" to be written to a file in case the
	// user is doing this to reset it's timestamp (or create an empty file).
	//if (!aBuf || !*aBuf)
	//	_f_return_retval;

	// Use the read-file loop's current item if filename was explicitly left blank (i.e. not just
	// a reference to a variable that's blank):
	LoopReadFileStruct *aCurrentReadFile = (aParamCount < 2) ? g->mLoopReadFile : NULL;
	if (aCurrentReadFile)
		aFilespec = aCurrentReadFile->mWriteFileName;
	if (!*aFilespec) // Nothing to write to.
		_f_throw_value(ERR_PARAM2_MUST_NOT_BE_BLANK);

	TextStream *ts = aCurrentReadFile ? aCurrentReadFile->mWriteFile : NULL;
	bool file_was_already_open = ts;

#ifdef CONFIG_DEBUGGER
	if (*aFilespec == '*' && !aFilespec[1] && !aBuf_obj && g_Debugger.OutputStdOut(aBuf))
	{
		// StdOut has been redirected to the debugger, and this "FileAppend" call has been
		// fully handled by the call above, so just return.
		g->LastError = 0;
		_f_return_retval;
	}
#endif

	UINT codepage;

	// Check if the file needs to be opened.  This is done here rather than at the time the
	// loop first begins so that:
	// 1) Any options/encoding specified in the first FileAppend call can take effect.
	// 2) To avoid opening the file if the file-reading loop has zero iterations (i.e. it's
	//    opened only upon first actual use to help performance and avoid changing the
	//    file-modification time when no actual text will be appended).
	if (!file_was_already_open)
	{
		codepage = aBuf_obj ? -1 : g->Encoding; // Never default to BOM if a Buffer object was passed.
		bool translate_crlf_to_lf = false;
		if (!ConvertFileOptions(aResultToken, aOptions, codepage, translate_crlf_to_lf, NULL))
			return;

		DWORD flags = TextStream::APPEND | (translate_crlf_to_lf ? TextStream::EOL_CRLF : 0);
		
		ASSERT( (~CP_AHKNOBOM) == CP_AHKCP );
		// codepage may include CP_AHKNOBOM, in which case below will not add BOM_UTFxx flag.
		if (codepage == CP_UTF8)
			flags |= TextStream::BOM_UTF8;
		else if (codepage == CP_UTF16)
			flags |= TextStream::BOM_UTF16;
		else if (codepage != -1)
			codepage &= CP_AHKCP;

		// Open the output file (if one was specified).  Unlike the input file, this is not
		// a critical error if it fails.  We want it to be non-critical so that FileAppend
		// commands in the body of the loop will throw to indicate the problem:
		ts = new TextFile; // ts was already verified NULL via !file_was_already_open.
		if ( !ts->Open(aFilespec, flags, codepage) )
		{
			aResultToken.SetLastErrorMaybeThrow(true);
			delete ts; // Must be deleted explicitly!
			return;
		}
		if (aCurrentReadFile)
			aCurrentReadFile->mWriteFile = ts;
	}
	else
		codepage = ts->GetCodePage();

	// Write to the file:
	DWORD result = 1;
	if (aBuf_length)
	{
		if (codepage == -1 || aBuf_obj) // "RAW" mode.
			result = ts->Write((LPCVOID)aBuf, (DWORD)aBuf_length);
		else
			result = ts->Write(aBuf, DWORD(aBuf_length / sizeof(TCHAR)));
	}
	//else: aBuf is empty; we've already succeeded in creating the file and have nothing further to do.
	aResultToken.SetLastErrorMaybeThrow(result == 0);

	if (!aCurrentReadFile)
		delete ts;
	// else it's the caller's responsibility, or it's caller's, to close it.
}



BOOL FileDeleteCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData)
{
	return DeleteFile(aFilename);
}

ResultType Line::FileDelete(LPTSTR aFilePattern)
{
	if (!*aFilePattern)
		return LineError(ERR_PARAM1_INVALID, FAIL_OR_OK);

	// The no-wildcard case could be handled via FilePatternApply(), but handling it this
	// way ensures deleting a non-existent path without wildcards is considered a failure:
	if (!StrChrAny(aFilePattern, _T("?*"))) // No wildcards; just a plain path/filename.
	{
		SetLastError(0); // For sanity: DeleteFile appears to set it only on failure.
		return SetLastErrorMaybeThrow(!DeleteFile(aFilePattern));
	}

	// Otherwise aFilePattern contains wildcards, so we'll search for all matches and delete them.
	return FilePatternApply(aFilePattern, FILE_LOOP_FILES_ONLY, false, FileDeleteCallback, NULL);
}



ResultType Line::FileInstall(LPTSTR aSource, LPTSTR aDest, LPTSTR aFlag)
{
	bool success;
	bool allow_overwrite = (ATOI(aFlag) == 1);
#ifndef AUTOHOTKEYSC
	if (g_script.mKind != Script::ScriptKindResource)
		success = FileInstallCopy(aSource, aDest, allow_overwrite);
	else
#endif
		success = FileInstallExtract(aSource, aDest, allow_overwrite);
	return ThrowIfTrue(!success);
}

bool Line::FileInstallExtract(LPTSTR aSource, LPTSTR aDest, bool aOverwrite)
{
	// Open the file first since it's the most likely to fail:
	HANDLE hfile = CreateFile(aDest, GENERIC_WRITE, 0, NULL, aOverwrite ? CREATE_ALWAYS : CREATE_NEW, 0, NULL);
	if (hfile == INVALID_HANDLE_VALUE)
		return false;

	// Create a temporary copy of aSource to ensure it is the correct case (upper-case).
	// Ahk2Exe converts it to upper-case before adding the resource. My testing showed that
	// using lower or mixed case in some instances prevented the resource from being found.
	// Since file paths are case-insensitive, it certainly doesn't seem harmful to do this:
	TCHAR source[T_MAX_PATH];
	size_t source_length = _tcslen(aSource);
	if (source_length >= _countof(source))
		// Probably can't happen; for simplicity, truncate it.
		source_length = _countof(source) - 1;
	tmemcpy(source, aSource, source_length + 1);
	_tcsupr(source);

	// Find and load the resource.
	HRSRC res;
	HGLOBAL res_load;
	LPVOID res_lock;
	bool success = false;
	if ( (res = FindResource(NULL, source, RT_RCDATA))
	  && (res_load = LoadResource(NULL, res))
	  && (res_lock = LockResource(res_load))  )
	{
		DWORD num_bytes_written;
		// Write the resource data to file.
		success = WriteFile(hfile, res_lock, SizeofResource(NULL, res), &num_bytes_written, NULL);
	}
	CloseHandle(hfile);
	return success;
}

#ifndef AUTOHOTKEYSC
bool Line::FileInstallCopy(LPTSTR aSource, LPTSTR aDest, bool aOverwrite)
{
	// v1.0.35.11: Must search in A_ScriptDir by default because that's where ahk2exe will search by default.
	// The old behavior was to search in A_WorkingDir, which seems pointless because ahk2exe would never
	// be able to use that value if the script changes it while running.
	TCHAR source_path[T_MAX_PATH], dest_path[T_MAX_PATH];
	GetFullPathName(aDest, _countof(dest_path), dest_path, NULL);
	// Avoid attempting the copy if both paths are the same (since it would fail with ERROR_SHARING_VIOLATION),
	// but resolve both to full paths in case mFileDir != g_WorkingDir.  There is a more thorough way to detect
	// when two *different* paths refer to the same file, but it doesn't work with different network shares, and
	// the additional complexity wouldn't be warranted.  Also, the limitations of this method are clearer.
	SetCurrentDirectory(g_script.mFileDir);
	GetFullPathName(aSource, _countof(source_path), source_path, NULL);
	SetCurrentDirectory(g_WorkingDir); // Restore to proper value.
	if (!lstrcmpi(source_path, dest_path) // Full paths are equal.
		&& !(GetFileAttributes(source_path) & FILE_ATTRIBUTE_DIRECTORY)) // Source file exists and is not a directory (otherwise, an error should be thrown).
		return true;

	return CopyFile(source_path, dest_path, !aOverwrite);
}
#endif



ResultType Line::FileCopyOrMove(LPTSTR aSource, LPTSTR aDest, bool aOverwrite)
{
	if (!*aDest) // Fix for v1.1.34.03: Previous behaviour was a Critical Error.
		return LineError(ERR_PARAM2_MUST_NOT_BE_BLANK);
	int error_count = 0;
	if (*aSource) // For backward-compatibility, empty Source is treated as "no files found".
		error_count = Util_CopyFile(aSource, aDest, aOverwrite, mActionType == ACT_FILEMOVE, g->LastError);
	return ThrowIntIfNonzero(error_count);
}



BIF_DECL(BIF_FileGetAttrib)
{
	_f_param_string_opt_def(aFilespec, 0, (g->mLoopFile ? g->mLoopFile->cFileName : _T("")));

	if (!*aFilespec)
		_f_throw_value(ERR_PARAM2_MUST_NOT_BE_BLANK);

	DWORD attr = GetFileAttributes(aFilespec);
	if (attr == 0xFFFFFFFF)  // Failure, probably because file doesn't exist.
	{
		aResultToken.SetLastErrorMaybeThrow(true);
		return;
	}

	g->LastError = 0;
	_f_return_p(FileAttribToStr(_f_retval_buf, attr));
}



BOOL FileSetAttribCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData);
struct FileSetAttribData
{
	DWORD and_mask, xor_mask;
};

ResultType Line::FileSetAttrib(LPTSTR aAttributes, LPTSTR aFilePattern
	, FileLoopModeType aOperateOnFolders, bool aDoRecurse)
// Returns the number of files and folders that could not be changed due to an error.
{
	if (!*aFilePattern)
		return LineError(ERR_PARAM2_INVALID, FAIL_OR_OK);

	// Convert the attribute string to three bit-masks: add, remove and toggle.
	FileSetAttribData attrib;
	DWORD mask;
	int op = 0;
	attrib.and_mask = 0xFFFFFFFF; // Set default: keep all bits.
	attrib.xor_mask = 0; // Set default: affect none.
	for (LPTSTR cp = aAttributes; *cp; ++cp)
	{
		switch (ctoupper(*cp))
		{
		case '+':
		case '-':
		case '^':
			op = *cp;
		case ' ':
		case '\t':
			continue;
		default:
			return LineError(ERR_PARAM1_INVALID, FAIL_OR_OK, cp);
		// Note that D (directory) and C (compressed) are currently not supported:
		case 'R': mask = FILE_ATTRIBUTE_READONLY; break;
		case 'A': mask = FILE_ATTRIBUTE_ARCHIVE; break;
		case 'S': mask = FILE_ATTRIBUTE_SYSTEM; break;
		case 'H': mask = FILE_ATTRIBUTE_HIDDEN; break;
		// N: Docs say it's valid only when used alone.  But let the API handle it if this is not so.
		case 'N': mask = FILE_ATTRIBUTE_NORMAL; break;
		case 'O': mask = FILE_ATTRIBUTE_OFFLINE; break;
		case 'T': mask = FILE_ATTRIBUTE_TEMPORARY; break;
		}
		switch (op)
		{
		case '+':
			attrib.and_mask &= ~mask; // Reset bit to 0.
			attrib.xor_mask |= mask; // Set bit to 1.
			break;
		case '-':
			attrib.and_mask &= ~mask; // Reset bit to 0.
			attrib.xor_mask &= ~mask; // Override any prior + or ^.
			break;
		case '^':
			attrib.xor_mask ^= mask; // Toggle bit.  ^= vs |= to invert any prior + or ^.
			// Leave and_mask as is, so any prior + or - will be inverted.
			break;
		default: // No +/-/^ specified, so overwrite attributes (equal and opposite to FileGetAttrib).
			attrib.and_mask = 0; // Reset all bits to 0.
			attrib.xor_mask |= mask; // Set bit to 1.  |= to accumulate if multiple attributes are present.
			break;
		}
	}
	FilePatternApply(aFilePattern, aOperateOnFolders, aDoRecurse, FileSetAttribCallback, &attrib);
	return OK;
}

BOOL FileSetAttribCallback(LPTSTR file_path, WIN32_FIND_DATA &current_file, void *aCallbackData)
{
	FileSetAttribData &attrib = *(FileSetAttribData *)aCallbackData;
	DWORD file_attrib = ((current_file.dwFileAttributes & attrib.and_mask) ^ attrib.xor_mask);
	if (!SetFileAttributes(file_path, file_attrib))
	{
		g->LastError = GetLastError();
		return FALSE;
	}
	return TRUE;
}



ResultType Line::FilePatternApply(LPTSTR aFilePattern, FileLoopModeType aOperateOnFolders
	, bool aDoRecurse, FilePatternCallback aCallback, void *aCallbackData)
{
	if (!*aFilePattern)
		// Caller should handle this case before calling us if an exception is to be thrown.
		return SetLastErrorMaybeThrow(true, ERROR_INVALID_PARAMETER);

	if (aOperateOnFolders == FILE_LOOP_INVALID) // In case runtime dereference of a var was an invalid value.
		aOperateOnFolders = FILE_LOOP_FILES_ONLY;  // Set default.
	g->LastError = 0; // Set default. Overridden only when a failure occurs.

	FilePatternStruct fps;

	LPTSTR last_backslash = _tcsrchr(aFilePattern, '\\');
	if (last_backslash)
		fps.dir_length = last_backslash - aFilePattern + 1; // Include the slash.
	else // Use current working directory, e.g. if user specified only *.*
		fps.dir_length = 0;
	fps.pattern_length = _tcslen(aFilePattern + fps.dir_length);
	
	// Testing shows that the ANSI version of FindFirstFile() will not accept a path+pattern longer
	// than 259, even if the pattern would match files whose names are short enough to be legal.
	if (fps.dir_length + fps.pattern_length >= _countof(fps.path)
		|| fps.pattern_length >= _countof(fps.pattern))
		return SetLastErrorMaybeThrow(true, ERROR_BUFFER_OVERFLOW);

	// Make copies in case of overwrite of deref buf during LONG_OPERATION/MsgSleep,
	// and to allow modification:
	_tcscpy(fps.path, aFilePattern); // Include the pattern initially.
	_tcscpy(fps.pattern, aFilePattern + fps.dir_length); // Just the naked filename or pattern, for use with aDoRecurse.

	if (!StrChrAny(fps.pattern, _T("?*")))
		// Since no wildcards, always operate on this single item even if it's a folder.
		aOperateOnFolders = FILE_LOOP_FILES_AND_FOLDERS;

	// Passing the parameters this way reduces code size:
	fps.aCallback = aCallback;
	fps.aCallbackData = aCallbackData;
	fps.aDoRecurse = aDoRecurse;
	fps.aOperateOnFolders = aOperateOnFolders;

	fps.failure_count = 0;

	FilePatternApply(fps);
	return ThrowIntIfNonzero(fps.failure_count); // i.e. indicate success if there were no failures.
}



void Line::FilePatternApply(FilePatternStruct &fps)
{
	size_t dir_length = fps.dir_length; // Length of this directory (saved before recursion).
	LPTSTR append_pos = fps.path + dir_length; // This is where the changing part gets appended.
	size_t space_remaining = _countof(fps.path) - dir_length - 1; // Space left in file_path for the changing part.

	LONG_OPERATION_INIT
	int failure_count = 0;
	WIN32_FIND_DATA current_file;
	HANDLE file_search = FindFirstFile(fps.path, &current_file);

	if (file_search != INVALID_HANDLE_VALUE)
	{
		do
		{
			// Since other script threads can interrupt during LONG_OPERATION_UPDATE, it's important that
			// this command not refer to sArgDeref[] and sArgVar[] anytime after an interruption becomes
			// possible. This is because an interrupting thread usually changes the values to something
			// inappropriate for this thread.
			LONG_OPERATION_UPDATE

			if (current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
			{
				if (current_file.cFileName[0] == '.' && (!current_file.cFileName[1]    // Relies on short-circuit boolean order.
					|| current_file.cFileName[1] == '.' && !current_file.cFileName[2]) //
					// Regardless of whether this folder will be recursed into, this folder
					// will not be affected when the mode is files-only:
					|| fps.aOperateOnFolders == FILE_LOOP_FILES_ONLY)
					continue; // Never operate upon or recurse into these.
			}
			else // It's a file, not a folder.
				if (fps.aOperateOnFolders == FILE_LOOP_FOLDERS_ONLY)
					continue;

			if (_tcslen(current_file.cFileName) > space_remaining)
			{
				// v1.0.45.03: Don't even try to operate upon truncated filenames in case they accidentally
				// match the name of a real/existing file.
				g->LastError = ERROR_BUFFER_OVERFLOW;
				++failure_count;
				continue;
			}
			// Otherwise, make file_path be the filespec of the file to operate upon:
			_tcscpy(append_pos, current_file.cFileName); // Above has ensured this won't overflow.
			//
			// This is the part that actually does something to the file:
			if (!fps.aCallback(fps.path, current_file, fps.aCallbackData))
				++failure_count;
			//
		} while (FindNextFile(file_search, &current_file));

		FindClose(file_search);
	} // if (file_search != INVALID_HANDLE_VALUE)

	if (fps.aDoRecurse && space_remaining > 1) // The space_remaining check ensures there's enough room to append "*", though if false, that would imply lfs.pattern is empty.
	{
		_tcscpy(append_pos, _T("*")); // Above has ensured this won't overflow.
		file_search = FindFirstFile(fps.path, &current_file);

		if (file_search != INVALID_HANDLE_VALUE)
		{
			do
			{
				LONG_OPERATION_UPDATE
				if (!(current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
					|| current_file.cFileName[0] == '.' && (!current_file.cFileName[1]      // Relies on short-circuit boolean order.
						|| current_file.cFileName[1] == '.' && !current_file.cFileName[2])) //
					continue;
				size_t filename_length = _tcslen(current_file.cFileName);
				// v1.0.45.03: Skip over folders whose paths are too long to be supported by FindFirst.
				if (fps.pattern_length + filename_length >= space_remaining) // >= vs. > to reserve 1 for the backslash to be added between cFileName and naked_filename_or_pattern.
					continue; // Never recurse into these.
				// This will build the string CurrentDir+SubDir+FilePatternOrName.
				// If FilePatternOrName doesn't contain a wildcard, the recursion
				// process will attempt to operate on the originally-specified
				// single filename or folder name if it occurs anywhere else in the
				// tree, e.g. recursing C:\Temp\temp.txt would affect all occurrences
				// of temp.txt both in C:\Temp and any subdirectories it might contain:
				_stprintf(append_pos, _T("%s\\%s") // Above has ensured this won't overflow.
					, current_file.cFileName, fps.pattern);
				fps.dir_length = dir_length + filename_length + 1; // Include the slash.
				//
				// Apply the callback to files in this subdirectory:
				FilePatternApply(fps);
				//
			} while (FindNextFile(file_search, &current_file));
			FindClose(file_search);
		} // if (file_search != INVALID_HANDLE_VALUE)
	} // if (aDoRecurse)

	fps.failure_count += failure_count; // Update failure count (produces smaller code than doing ++fps.failure_count directly).
}



BIF_DECL(BIF_FileGetTime)
{
	_f_param_string_opt_def(aFilespec, 0, (g->mLoopFile ? g->mLoopFile->cFileName : _T("")));
	_f_param_string_opt(aWhichTime, 1);

	if (!*aFilespec)
		_f_throw_value(ERR_PARAM2_MUST_NOT_BE_BLANK);

	// Don't use CreateFile() & FileGetSize() size they will fail to work on a file that's in use.
	// Research indicates that this method has no disadvantages compared to the other method.
	WIN32_FIND_DATA found_file;
	HANDLE file_search = FindFirstFile(aFilespec, &found_file);
	if (file_search == INVALID_HANDLE_VALUE)
	{
		aResultToken.SetLastErrorMaybeThrow(true);
		return;
	}
	FindClose(file_search);

	FILETIME local_file_time;
	switch (ctoupper(*aWhichTime))
	{
	case 'C': // File's creation time.
		FileTimeToLocalFileTime(&found_file.ftCreationTime, &local_file_time);
		break;
	case 'A': // File's last access time.
		FileTimeToLocalFileTime(&found_file.ftLastAccessTime, &local_file_time);
		break;
	default:  // 'M', unspecified, or some other value.  Use the file's modification time.
		FileTimeToLocalFileTime(&found_file.ftLastWriteTime, &local_file_time);
	}

	g->LastError = 0;
	_f_return_p(FileTimeToYYYYMMDD(_f_retval_buf, local_file_time));
}



BOOL FileSetTimeCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData);
struct FileSetTimeData
{
	FILETIME Time;
	TCHAR WhichTime;
};

ResultType Line::FileSetTime(LPTSTR aYYYYMMDD, LPTSTR aFilePattern, TCHAR aWhichTime
	, FileLoopModeType aOperateOnFolders, bool aDoRecurse)
// Returns the number of files and folders that could not be changed due to an error.
{
	// Related to the comment at the top: Since the script subroutine that resulted in the call to
	// this function can be interrupted during our MsgSleep(), make a copy of any params that might
	// currently point directly to the deref buffer.  This is done because their contents might
	// be overwritten by the interrupting subroutine:
	TCHAR yyyymmdd[64]; // Even do this one since its value is passed recursively in calls to self.
	tcslcpy(yyyymmdd, aYYYYMMDD, _countof(yyyymmdd));

	FileSetTimeData callbackData;
	callbackData.WhichTime = aWhichTime;
	FILETIME ft;
	if (*yyyymmdd)
	{
		if (   !YYYYMMDDToFileTime(yyyymmdd, ft)  // Convert the arg into the time struct as local (non-UTC) time.
			|| !LocalFileTimeToFileTime(&ft, &callbackData.Time)   )  // Convert from local to UTC.
		{
			// Invalid parameters are the only likely cause of this condition.
			return LineError(ERR_PARAM1_INVALID, FAIL_OR_OK, aYYYYMMDD);
		}
	}
	else // User wants to use the current time (i.e. now) as the new timestamp.
		GetSystemTimeAsFileTime(&callbackData.Time);

	if (!*aFilePattern)
		return LineError(ERR_PARAM2_INVALID, FAIL_OR_OK);

	FilePatternApply(aFilePattern, aOperateOnFolders, aDoRecurse, FileSetTimeCallback, &callbackData);
	return OK;
}

BOOL FileSetTimeCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData)
{
	HANDLE hFile;
	// Open existing file.
	// FILE_FLAG_NO_BUFFERING might improve performance because all we're doing is
	// changing one of the file's attributes.  FILE_FLAG_BACKUP_SEMANTICS must be
	// used, otherwise changing the time of a directory under NT and beyond will
	// not succeed.  Win95 (not sure about Win98/Me) does not support this, but it
	// should be harmless to specify it even if the OS is Win95:
	hFile = CreateFile(aFilename, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE
		, (LPSECURITY_ATTRIBUTES)NULL, OPEN_EXISTING
		, FILE_FLAG_NO_BUFFERING | FILE_FLAG_BACKUP_SEMANTICS, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
	{
		g->LastError = GetLastError();
		return FALSE;
	}

	BOOL success;
	FileSetTimeData &a = *(FileSetTimeData *)aCallbackData;
	switch (ctoupper(a.WhichTime))
	{
	case 'C': // File's creation time.
		success = SetFileTime(hFile, &a.Time, NULL, NULL);
		break;
	case 'A': // File's last access time.
		success = SetFileTime(hFile, NULL, &a.Time, NULL);
		break;
	default:  // 'M', unspecified, or some other value.  Use the file's modification time.
		success = SetFileTime(hFile, NULL, NULL, &a.Time);
	}
	if (!success)
		g->LastError = GetLastError();
	CloseHandle(hFile);
	return success;
}



BIF_DECL(BIF_FileGetSize)
{
	_f_param_string_opt_def(aFilespec, 0, (g->mLoopFile ? g->mLoopFile->cFileName : _T("")));
	_f_param_string_opt(aGranularity, 1);

	if (!*aFilespec)
		_f_throw_value(ERR_PARAM2_MUST_NOT_BE_BLANK); // Throw an error, since this is probably not what the user intended.
	
	BOOL got_file_size = false;
	__int64 size;

	// Try CreateFile() and GetFileSizeEx() first, since they can be more accurate. 
	// See "Why is the file size reported incorrectly for files that are still being written to?"
	// http://blogs.msdn.com/b/oldnewthing/archive/2011/12/26/10251026.aspx
	HANDLE hfile = CreateFile(aFilespec, FILE_READ_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE
		, NULL, OPEN_EXISTING, 0, NULL);
	if (hfile != INVALID_HANDLE_VALUE)
	{
		got_file_size = GetFileSizeEx(hfile, (PLARGE_INTEGER)&size);
		CloseHandle(hfile);
	}

	if (!got_file_size)
	{
		WIN32_FIND_DATA found_file;
		HANDLE file_search = FindFirstFile(aFilespec, &found_file);
		if (file_search == INVALID_HANDLE_VALUE)
		{
			aResultToken.SetLastErrorMaybeThrow(true);
			_f_return_empty;
		}
		FindClose(file_search);
		size = ((__int64)found_file.nFileSizeHigh << 32) | found_file.nFileSizeLow;
	}

	switch(ctoupper(*aGranularity))
	{
	case 'K': // KB
		size /= 1024;
		break;
	case 'M': // MB
		size /= (1024 * 1024);
		break;
	// default: // i.e. either 'B' for bytes, or blank, or some other unknown value, so default to bytes.
		// do nothing
	}

	g->LastError = 0;
	_f_return(size);
	// The below comment is obsolete in light of the switch to 64-bit integers.  But it might
	// be good to keep for background:
	// Currently, the above is basically subject to a 2 gig limit, I believe, after which the
	// size will appear to be negative.  Beyond a 4 gig limit, the value will probably wrap around
	// to zero and start counting from there as file sizes grow beyond 4 gig (UPDATE: The size
	// is now set to -1 [the maximum DWORD when expressed as a signed int] whenever >4 gig).
	// There's not much sense in putting values larger than 2 gig into the var as a text string
	// containing a positive number because such a var could never be properly handled by anything
	// that compares to it (e.g. IfGreater) or does math on it (e.g. EnvAdd), since those operations
	// use ATOI() to convert the string.  So as a future enhancement (unless the whole program is
	// revamped to use 64bit ints or something) might add an optional param to the end to indicate
	// size should be returned in K(ilobyte) or M(egabyte).  However, this is sorta bad too since
	// adding a param can break existing scripts which use filenames containing commas (delimiters)
	// with this command.  Therefore, I think I'll just add the K/M param now.
	// Also, the above assigns an int because unsigned ints should never be stored in script
	// variables.  This is because an unsigned variable larger than INT_MAX would not be properly
	// converted by ATOI(), which is current standard method for variables to be auto-converted
	// from text back to a number whenever that is needed.
}



/////////////////////
// SetXXXLockState //
/////////////////////

ResultType Line::SetToggleState(vk_type aVK, ToggleValueType &ForceLock, LPTSTR aToggleText)
// Caller must have already validated that the args are correct.
// Always returns OK, for use as caller's return value.
{
	ToggleValueType toggle = ConvertOnOffAlways(aToggleText, NEUTRAL);
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
	}
	return OK;
}



////////////////////////////////
// Misc lower level functions //
////////////////////////////////

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



//////////////////////
// Window Functions //
//////////////////////

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
