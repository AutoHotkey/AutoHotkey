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



void MsgSleepWithListLines(int aSleepDuration, Line *waiting_line, DWORD start_time)
{
	if (MsgSleep(aSleepDuration))
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



typedef bool (*WaitCompletedPredicate)(void *);

static bool Wait(int aTimeout, void *aParam, WaitCompletedPredicate aWaitCompleted)
{
	Line *waiting_line = g_script.mCurrLine;

	for (DWORD start_time = GetTickCount();;) // start_time is initialized unconditionally for use with ListLines.
	{
		if (aWaitCompleted(aParam))
			return true;

		// Must cast to int or any negative result will be lost due to DWORD type:
		if (aTimeout < 0 || (aTimeout - (int)(GetTickCount() - start_time)) > SLEEP_INTERVAL_HALF)
			MsgSleepWithListLines(INTERVAL_UNSPECIFIED, waiting_line, start_time);
		else // Done waiting (timed out).
			return false;
	}
}



bif_impl FResult ClipWait(optl<double> aTimeout, optl<int> aWaitForAnyData, BOOL &aRetVal)
{
	int timeout = -1;
	if (aTimeout.has_value())
	{
		timeout = (int)(aTimeout.value() * 1000);
		if (timeout < 0)
			return FR_E_ARG(0);
	}
	WaitCompletedPredicate predicate;
	switch (aWaitForAnyData.value_or(0))
	{
	case 0: // Wait for text
		predicate = [](void*) -> bool {
			// Seems best to consider CF_HDROP to be a non-empty clipboard, since we
			// support the implicit conversion of that format to text:
			return (IsClipboardFormatAvailable(CF_NATIVETEXT) || IsClipboardFormatAvailable(CF_HDROP));
		};
		break;
	case 1: // Wait for any clipboard data
		predicate = [](void*) -> bool {
			return CountClipboardFormats();
		};
		break;
	default: // Reserved
		return FR_E_ARG(1);
	}
	aRetVal = Wait(timeout, nullptr, predicate);
	return OK;
}



bif_impl FResult KeyWait(StrArg aKeyName, optl<StrArg> aOptions, BOOL &aRetVal)
{
	// Set defaults:
	int timeout = -1;
	struct Params
	{
		vk_type vk;
		JoyControls joy;
		int joystick_id;
		bool wait_for_keydown = false;  // The default is to wait for the key to be released.
		KeyStateTypes key_state_type = KEYSTATE_PHYSICAL;  // Since physical is more often used.
	} params;

	if (   !(params.vk = TextToVK(aKeyName))   )
	{
		params.joy = (JoyControls)ConvertJoy(aKeyName, &params.joystick_id);
		if (!IS_JOYSTICK_BUTTON(params.joy)) // Currently, only buttons are supported.
			// It's either an invalid key name or an unsupported Joy-something.
			return FR_E_ARG(0);
	}
	
	for (auto cp = aOptions.value_or_empty(); *cp; ++cp)
	{
		switch (ctoupper(*cp))
		{
		case 'D':
			params.wait_for_keydown = true;
			break;
		case 'L':
			params.key_state_type = KEYSTATE_LOGICAL;
			break;
		case 'T':
			timeout = (int)(ATOF(cp + 1) * 1000);
			if (timeout < 0)
				return FR_E_ARG(1);
			break;
		}
	}

	WaitCompletedPredicate predicate;
	if (params.vk) predicate = [](void *pp) -> bool { // Waiting for key or mouse button
		auto &p = *((Params*)pp);
		return ScriptGetKeyState(p.vk, p.key_state_type) == p.wait_for_keydown;
	};
	else predicate = [](void *pp) -> bool { // Waiting for joystick button
		auto &p = *((Params*)pp);
		TCHAR unused[32];
		ExprTokenType token;
		return ScriptGetJoyState(p.joy, p.joystick_id, token, unused) == p.wait_for_keydown;
	};
	aRetVal = Wait(timeout, &params, predicate);
	return OK;
}



static FResult WinWait(ExprTokenType *aWinTitle, optl<StrArg> aWinText, optl<double> aTimeout, optl<StrArg> aExcludeTitle, optl<StrArg> aExcludeText
	, UINT &aRetVal, BuiltInFunctionID aCondition)
{
	int timeout = -1;
	if (aTimeout.has_value())
	{
		timeout = (int)(aTimeout.value() * 1000);
		if (timeout < 0)
			return FR_E_ARG(2);
	}
	
	TCHAR title_buf[MAX_NUMBER_SIZE];
	WaitCompletedPredicate predicate;
	struct Params {
		BuiltInFunctionID condition;
		bool hwnd_specified;
		HWND hwnd;
		LPCTSTR title, text, exclude_title, exclude_text;
	} p;
	
	p.condition = aCondition;
	p.hwnd_specified = false;
	p.hwnd = NULL;
	p.title = _T("");
	p.text = aWinText.value_or_empty();
	p.exclude_title = aExcludeTitle.value_or_empty();
	p.exclude_text = aExcludeText.value_or_empty();

	if (aWinTitle)
	{
		auto fr = DetermineTargetHwnd(p.hwnd, p.hwnd_specified, *aWinTitle);
		if (fr != OK)
			return fr;
		if (p.hwnd_specified && !p.hwnd)
		{
			// Caller may have specified a NULL HWND or it may be NULL due to an IsWindow() check.
			// In either case, it isn't meaningful to wait (see !IsWindow(p.hwnd) comments below).
			aRetVal = (aCondition == FID_WinWaitClose || aCondition == FID_WinWaitNotActive);
			if (aRetVal)
				DoWinDelay;
			return OK;
		}
		if (!p.hwnd_specified)
			p.title = TokenToString(*aWinTitle, title_buf);
	}

	if (p.hwnd_specified)
	{
		predicate = [](void *pp)
		{
			auto &p = *(Params*)pp;
			if (!IsWindow(p.hwnd))
			{
				// It's not meaningful to wait for another window to be created with this
				// same HWND, so return now regardless of whether the wait condition was met.
				// Let the return value be 1 for WinWaitClose/WinWaitNotActive to indicate the
				// condition was met, and 0 for WinWait/WinWaitActive to indicate that it wasn't.
				p.hwnd = NULL;
				return true;
			}
			bool is_match, want_match;
			if (p.condition == FID_WinWait || p.condition == FID_WinWaitClose)
			{
				// Wait for the window to become visible/hidden.  Most functions ignore
				// DetectHiddenWindows when given a pure HWND/object (because it's more
				// useful that way), but in this case it seems more useful and intuitive
				// to respect DetectHiddenWindows.
				is_match = g->DetectWindow(p.hwnd);
				want_match = (p.condition == FID_WinWait);
			}
			else
			{
				is_match = (GetForegroundWindow() == p.hwnd);
				want_match = (p.condition == FID_WinWaitActive);
			}
			if (is_match)
				g->hWndLastUsed = p.hwnd;
			return is_match == want_match;
		};
	}
	else // hwnd_specified == false
	switch (aCondition)
	{
	default:
#ifdef _DEBUG
		ASSERT(!"Unhandled case");
		break;
#endif
	case FID_WinWait:
	case FID_WinWaitClose:
		predicate = [](void *pp)
		{
			auto &p = *(Params*)pp;
			HWND found = WinExist(*g, p.title, p.text, p.exclude_title, p.exclude_text, false, true);
			if ((found != NULL) == (p.condition == FID_WinWait))
			{
				p.hwnd = found;
				return true;
			}
			return false;
		};
		break;
	case FID_WinWaitActive:
	case FID_WinWaitNotActive:
		predicate = [](void *pp)
		{
			auto &p = *(Params*)pp;
			HWND found = WinActive(*g, p.title, p.text, p.exclude_title, p.exclude_text, true);
			if ((found != NULL) == (p.condition == FID_WinWaitActive))
			{
				p.hwnd = found;
				return true;
			}
			return false;
		};
		break;
	}

	if (Wait(timeout, &p, predicate))
	{
		DoWinDelay;
		if (aCondition == FID_WinWaitClose || aCondition == FID_WinWaitNotActive)
			aRetVal = 1;
		else
			aRetVal = (UINT)(size_t)p.hwnd;
	}
	else
		aRetVal = 0;
	return OK;
}

bif_impl FResult WinWait(ExprTokenType *aWinTitle, optl<StrArg> aWinText, optl<double> aTimeout, optl<StrArg> aExcludeTitle, optl<StrArg> aExcludeText, UINT &aRetVal)
{
	return WinWait(aWinTitle, aWinText, aTimeout, aExcludeTitle, aExcludeText, aRetVal, FID_WinWait);
}

bif_impl FResult WinWaitClose(ExprTokenType *aWinTitle, optl<StrArg> aWinText, optl<double> aTimeout, optl<StrArg> aExcludeTitle, optl<StrArg> aExcludeText, UINT &aRetVal)
{
	return WinWait(aWinTitle, aWinText, aTimeout, aExcludeTitle, aExcludeText, aRetVal, FID_WinWaitClose);
}

bif_impl FResult WinWaitActive(ExprTokenType *aWinTitle, optl<StrArg> aWinText, optl<double> aTimeout, optl<StrArg> aExcludeTitle, optl<StrArg> aExcludeText, UINT &aRetVal)
{
	return WinWait(aWinTitle, aWinText, aTimeout, aExcludeTitle, aExcludeText, aRetVal, FID_WinWaitActive);
}

bif_impl FResult WinWaitNotActive(ExprTokenType *aWinTitle, optl<StrArg> aWinText, optl<double> aTimeout, optl<StrArg> aExcludeTitle, optl<StrArg> aExcludeText, UINT &aRetVal)
{
	return WinWait(aWinTitle, aWinText, aTimeout, aExcludeTitle, aExcludeText, aRetVal, FID_WinWaitNotActive);
}



// RunWait is still a "BIF" for now because the OutputVarPID parameter can only be used
// by another thread, while the function is running, and there's currently no means for
// a function implemented via MdFunc to assign to the output var directly.
BIF_DECL(BIF_RunWait)
{
	Line *waiting_line = g_script.mCurrLine;

	_f_param_string_opt(arg1, 0);
	_f_param_string_opt(arg2, 1);
	_f_param_string_opt(arg3, 2);
	Var *output_var = ParamIndexToOutputVar(3);
	if (output_var)
		output_var->Assign();

	HANDLE running_process;
	if (!g_script.ActionExec(arg1, NULL, arg2, true, arg3, &running_process, true, true))
		_f_return_FAIL;
	
	// For the output var to be useful, it must be assigned before we wait:
	if (output_var && running_process)
		output_var->Assign(GetProcessId(running_process));

	if (!running_process) // Nothing to wait for (rare?).
		_f_return_i(0);

	Wait(-1, (void*)running_process, [](void *p)
		{
			// Using WaitForSingleObject() rather than GetExitCodeProcess() avoids an
			// infinite loop if a process returns 259 (STILL_ACTIVE) as its exit code.
			return WaitForSingleObject((HANDLE)p, 0) != WAIT_TIMEOUT;
		});

	DWORD exit_code = 0;
	GetExitCodeProcess(running_process, &exit_code);
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
