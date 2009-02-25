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

#ifndef application_h
#define application_h

#include "defines.h"

// Callers should note that using INTERVAL_UNSPECIFIED might not rest the CPU at all if there is
// already at least one msg waiting in our thread's msg queue:
// Use some negative value unlikely to ever be passed explicitly:
#define INTERVAL_UNSPECIFIED (INT_MIN + 303)
#define NO_SLEEP -1
enum MessageMode {WAIT_FOR_MESSAGES, RETURN_AFTER_MESSAGES, RETURN_AFTER_MESSAGES_SPECIAL_FILTER};
bool MsgSleep(int aSleepDuration = INTERVAL_UNSPECIFIED, MessageMode aMode = RETURN_AFTER_MESSAGES);

// This macro is used to Sleep without the possibility of a new hotkey subroutine being launched.
// Timed subroutines will also be prevented from running while it is enabled.
// It should be used when an operation needs to sleep, but doesn't want to be interrupted (suspended)
// by any hotkeys the user might press during that time.  Reasons why the caller wouldn't want to
// be suspended:
// 1) If it's doing something with a window -- such as sending keys or clicking the mouse or trying
//    to activate it -- that might get messed up if a new hotkey fires in the middle of the operation.
// 2) If its a command that's still using some of its parameters that might reside in the deref buffer.
//    In this case, the launching of a new hotkey would likely overwrite those values, causing
//    unpredictable behavior.
#define SLEEP_WITHOUT_INTERRUPTION(aSleepTime) \
{\
	g_AllowInterruption = FALSE;\
	MsgSleep(aSleepTime);\
	g_AllowInterruption = TRUE;\
}

// Have this be dynamically resolved each time.  For example, when MsgSleep() uses this
// while in mode WAIT_FOR_MESSSAGES, its msg loop should use this macro in case the
// value of g_AllowInterruption changes from one iteration to the next.  Thankfully,
// MS made WM_HOTKEY have a very high value, so filtering in this way should not exclude
// any other important types of messages:
#define MSG_FILTER_MAX (IsInterruptible() ? 0 : WM_HOTKEY - 1)
#define INTERRUPTIBLE_IN_EMERGENCY (g_AllowInterruption && !g_MenuIsVisible)

// Do a true Sleep() for short sleeps on Win9x because it is much more accurate than the MsgSleep()
// method on that OS, at least for when short sleeps are done on Win98SE:
#define DoWinDelay \
	if (::g->WinDelay > -1)\
	{\
		if (::g->WinDelay < 25 && g_os.IsWin9x())\
			Sleep(::g->WinDelay);\
		else\
			MsgSleep(::g->WinDelay);\
	}

#define DoControlDelay \
	if (g->ControlDelay > -1)\
	{\
		if (g->ControlDelay < 25 && g_os.IsWin9x())\
			Sleep(g->ControlDelay);\
		else\
			MsgSleep(g->ControlDelay);\
	}

ResultType IsCycleComplete(int aSleepDuration, DWORD aStartTime, bool aAllowEarlyReturn);

// These should only be called from MsgSleep() (or something called by MsgSleep()) because
// we don't want to be in the situation where a thread launched by CheckScriptTimers() returns
// first to a dialog's message pump rather than MsgSleep's pump.  That's because our thread
// might then have queued messages that would be stuck in the queue (due to the possible absence
// of the main timer) until the dialog's msg pump ended.
bool CheckScriptTimers();
#define CHECK_SCRIPT_TIMERS_IF_NEEDED if (g_script.mTimerEnabledCount && CheckScriptTimers()) return_value = true; // Change the existing value only if it returned true.

void PollJoysticks();
#define POLL_JOYSTICK_IF_NEEDED if (Hotkey::sJoyHotkeyCount) PollJoysticks();

bool MsgMonitor(HWND aWnd, UINT aMsg, WPARAM awParam, LPARAM alParam, MSG *apMsg, LRESULT &aMsgReply);

void InitNewThread(int aPriority, bool aSkipUninterruptible, bool aIncrementThreadCountAndUpdateTrayIcon
	, ActionTypeType aTypeOfFirstLine);
void ResumeUnderlyingThread(char *aSavedErrorLevel);
BOOL IsInterruptible();

VOID CALLBACK MsgBoxTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);
VOID CALLBACK AutoExecSectionTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);
VOID CALLBACK UninterruptibleTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);
VOID CALLBACK InputTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);
VOID CALLBACK RefreshInterruptibility(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);

#endif
