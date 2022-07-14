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



BIF_DECL(BIF_Wait)
// Since other script threads can interrupt these commands while they're running, it's important that
// these commands not refer to sArgDeref[] and sArgVar[] anytime after an interruption becomes possible.
// This is because an interrupting thread usually changes the values to something inappropriate for this thread.
// fincs: it seems best that this function not throw an exception if the wait timeouts.
{
	bool wait_indefinitely;
	int sleep_duration;
	DWORD start_time;

	vk_type vk; // For GetKeyState.
	HANDLE running_process; // For RUNWAIT
	DWORD exit_code; // For RUNWAIT

	// For FID_KeyWait:
	bool wait_for_keydown;
	KeyStateTypes key_state_type;
	JoyControls joy;
	int joystick_id;
	ExprTokenType token;

	Line *waiting_line = g_script.mCurrLine;

	_f_set_retval_i(TRUE); // Set default return value to be possibly overridden later on.

	_f_param_string_opt(arg1, 0);
	_f_param_string_opt(arg2, 1);
	_f_param_string_opt(arg3, 2);

	HWND target_hwnd = NULL;

	switch (_f_callee_id)
	{
	case FID_RunWait:
		if (!g_script.ActionExec(arg1, NULL, arg2, true, arg3, &running_process, true, true
			, ParamIndexToOutputVar(4-1)))
			_f_return_FAIL;
		//else fall through to the waiting-phase of the operation.
		break;
	case FID_WinWait:
	case FID_WinWaitClose:
	case FID_WinWaitActive:
	case FID_WinWaitNotActive:
		if (ParamIndexIsOmitted(0))
			break;
		// The following supports pure HWND or {Hwnd:HWND} as in other WinTitle parameters,
		// in lieu of saving aParam[] and evaluating them vs. arg1,2,4,5 on each iteration.
		switch (DetermineTargetHwnd(target_hwnd, aResultToken, *aParam[0]))
		{
		case FAIL: return;
		case OK:
			if (!target_hwnd) // Caller passed 0 or a value that failed an IsWindow() check.
			{
				// If it's 0 due to failing IsWindow(), there's no way to determine whether
				// it was ever valid, so assume it was but that the window has been closed.
				DoWinDelay;
				if (_f_callee_id == FID_WinWaitClose || _f_callee_id == FID_WinWaitNotActive)
					_f_return_retval;
				// Otherwise, this window will never exist or be active again, so abort early.
				// It might be more correct to just wait the timeout (or stall indefinitely),
				// but that's probably not the user's intention.
				_f_return_i(FALSE);
			}
		}
		break;
	}

	// These are declared/accessed after "case FID_RunWait:" to avoid a UseUnset warning.
	_f_param_string_opt(arg4, 3);
	_f_param_string_opt(arg5, 4);
	
	// Must NOT use ELSE-IF in line below due to ELSE further down needing to execute for RunWait.
	if (_f_callee_id == FID_KeyWait)
	{
		if (   !(vk = TextToVK(arg1))   )
		{
			joy = (JoyControls)ConvertJoy(arg1, &joystick_id);
			if (!IS_JOYSTICK_BUTTON(joy)) // Currently, only buttons are supported.
				// It's either an invalid key name or an unsupported Joy-something.
				_f_throw_param(0);
		}
		// Set defaults:
		wait_for_keydown = false;  // The default is to wait for the key to be released.
		key_state_type = KEYSTATE_PHYSICAL;  // Since physical is more often used.
		wait_indefinitely = true;
		sleep_duration = 0;
		for (LPTSTR cp = arg2; *cp; ++cp)
		{
			switch(ctoupper(*cp))
			{
			case 'D':
				wait_for_keydown = true;
				break;
			case 'L':
				key_state_type = KEYSTATE_LOGICAL;
				break;
			case 'T':
				// Although ATOF() supports hex, it's been documented in the help file that hex should
				// not be used (see comment above) so if someone does it anyway, some option letters
				// might be misinterpreted:
				wait_indefinitely = false;
				sleep_duration = (int)(ATOF(cp + 1) * 1000);
				break;
			}
		}
	}
	else if (   (_f_callee_id != FID_RunWait && _f_callee_id != FID_ClipWait && *arg3)
		|| (_f_callee_id == FID_ClipWait && *arg1)   )
	{
		// Since the param containing the timeout value isn't blank, it must be numeric,
		// otherwise, the loading validation would have prevented the script from loading.
		wait_indefinitely = false;
		sleep_duration = (int)(ATOF(_f_callee_id == FID_ClipWait ? arg1 : arg3) * 1000); // Can be zero.
	}
	else
	{
		wait_indefinitely = true;
		sleep_duration = 0; // Just to catch any bugs.
	}

	bool any_clipboard_format = (_f_callee_id == FID_ClipWait && ATOI(arg2) == 1);

	for (start_time = GetTickCount();;) // start_time is initialized unconditionally for use with v1.0.30.02's new logging feature further below.
	{ // Always do the first iteration so that at least one check is done.
		if (target_hwnd) // Caller passed a pure HWND or {Hwnd:HWND}.
		{
			// Change the behaviour a little since we know that once the HWND is destroyed,
			// it is not meaningful to wait for another window with that same HWND.
			if (!IsWindow(target_hwnd))
			{
				DoWinDelay;
				if (_f_callee_id == FID_WinWaitClose || _f_callee_id == FID_WinWaitNotActive)
					_f_return_retval; // Condition met.
				// Otherwise, it would not be meaningful to wait for another window to be
				// created with the same HWND.  It seems more useful to abort immediately
				// but report timeout/failure than to wait for the timeout to elapse.
				_f_return_i(FALSE);
			}
			if (_f_callee_id == FID_WinWait || _f_callee_id == FID_WinWaitClose)
			{
				// Wait for the window to become visible/hidden.  Most functions ignore
				// DetectHiddenWindows when given a pure HWND/object (because it's more
				// useful that way), but in this case it seems more useful and intuitive
				// to respect DetectHiddenWindows.
				if (g->DetectWindow(target_hwnd) == (_f_callee_id == FID_WinWait))
				{
					DoWinDelay;
					if (_f_callee_id == FID_WinWaitClose)
						_f_return_retval;
					_f_return_i((size_t)target_hwnd);
				}
			}
			else
			{
				if ((GetForegroundWindow() == target_hwnd) == (_f_callee_id == FID_WinWaitActive))
				{
					DoWinDelay;
					if (_f_callee_id == FID_WinWaitNotActive)
						_f_return_retval;
					_f_return_i((size_t)target_hwnd);
				}
			}
		}
		else switch (_f_callee_id)
		{
		case FID_WinWait:
			#define SAVED_WIN_ARGS arg1, arg2, arg4, arg5
			if (HWND found = WinExist(*g, SAVED_WIN_ARGS, false, true))
			{
				DoWinDelay;
				_f_return_i((size_t)found);
			}
			break;
		case FID_WinWaitClose:
			if (!WinExist(*g, SAVED_WIN_ARGS, false, true))
			{
				DoWinDelay;
				_f_return_retval;
			}
			break;
		case FID_WinWaitActive:
			if (HWND found = WinActive(*g, SAVED_WIN_ARGS, true))
			{
				DoWinDelay;
				_f_return_i((size_t)found);
			}
			break;
		case FID_WinWaitNotActive:
			if (!WinActive(*g, SAVED_WIN_ARGS, true))
			{
				DoWinDelay;
				_f_return_retval;
			}
			break;
		case FID_ClipWait:
			// Seems best to consider CF_HDROP to be a non-empty clipboard, since we
			// support the implicit conversion of that format to text:
			if (any_clipboard_format)
			{
				if (CountClipboardFormats())
					_f_return_retval;
			}
			else
				if (IsClipboardFormatAvailable(CF_NATIVETEXT) || IsClipboardFormatAvailable(CF_HDROP))
					_f_return_retval;
			break;
		case FID_KeyWait:
			if (vk) // Waiting for key or mouse button, not joystick.
			{
				if (ScriptGetKeyState(vk, key_state_type) == wait_for_keydown)
					_f_return_retval;
			}
			else // Waiting for joystick button
			{
				TCHAR unused[32];
				if (ScriptGetJoyState(joy, joystick_id, token, unused) == wait_for_keydown)
					_f_return_retval;
			}
			break;
		case FID_RunWait:
			// Pretty nasty, but for now, nothing is done to prevent an infinite loop.
			// In the future, maybe OpenProcess() can be used to detect if a process still
			// exists (is there any other way?):
			// MSDN: "Warning: If a process happens to return STILL_ACTIVE (259) as an error code,
			// applications that test for this value could end up in an infinite loop."
			if (running_process)
				GetExitCodeProcess(running_process, &exit_code);
			else // it can be NULL in the case of launching things like "find D:\" or "www.yahoo.com"
				exit_code = 0;
			if (exit_code != STATUS_PENDING) // STATUS_PENDING == STILL_ACTIVE
			{
				if (running_process)
					CloseHandle(running_process);
				// Use signed vs. unsigned, since that is more typical?  No, it seems better
				// to use unsigned now that script variables store 64-bit ints.  This is because
				// GetExitCodeProcess() yields a DWORD, implying that the value should be unsigned.
				// Unsigned also is more useful in cases where an app returns a (potentially large)
				// count of something as its result.  However, if this is done, it won't be easy
				// to check against a return value of -1, for example, which I suspect many apps
				// return.  AutoIt3 (and probably 2) use a signed int as well, so that is another
				// reason to keep it this way:
				_f_return_i((int)exit_code);
			}
			break;
		}

		// Must cast to int or any negative result will be lost due to DWORD type:
		if (wait_indefinitely || (int)(sleep_duration - (GetTickCount() - start_time)) > SLEEP_INTERVAL_HALF)
		{
			if (MsgSleep(INTERVAL_UNSPECIFIED)) // INTERVAL_UNSPECIFIED performs better.
			{
				// v1.0.30.02: Since MsgSleep() launched and returned from at least one new thread, put the
				// current waiting line into the line-log again to make it easy to see what the current
				// thread is doing.  This is especially useful for figuring out which subroutine is holding
				// another thread interrupted beneath it.  For example, if a timer gets interrupted by
				// a hotkey that has an indefinite WinWait, and that window never appears, this will allow
				// the user to find out the culprit thread by showing its line in the log (and usually
				// it will appear as the very last line, since usually the script is idle and thus the
				// currently active thread is the one that's still waiting for the window).
				if (g->ListLinesIsEnabled)
				{
					// ListLines is enabled in this thread, but if it was disabled in the interrupting thread,
					// the very last log entry will be ours.  In that case, we don't want to duplicate it.
					int previous_log_index = (Line::sLogNext ? Line::sLogNext : LINE_LOG_SIZE) - 1; // Wrap around if needed (the entry can be NULL in that case).
					if (Line::sLog[previous_log_index] != waiting_line || Line::sLogTick[previous_log_index] != start_time) // The previously logged line was not this one, or it was added by the interrupting thread (different start_time).
					{
						Line::sLog[Line::sLogNext] = waiting_line;
						Line::sLogTick[Line::sLogNext++] = start_time; // Store a special value so that Line::LogToText() can report that its "still waiting" from earlier.
						if (Line::sLogNext >= LINE_LOG_SIZE)
							Line::sLogNext = 0;
						// The lines above are the similar to those used in ExecUntil(), so the two should be
						// maintained together.
					}
				}
			}
		}
		else // Done waiting.
			_f_return_i(FALSE); // Since it timed out, we override the default with this.
	} // for()
}
