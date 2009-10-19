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
#include "window.h"
#include "util.h" // for strlcpy()
#include "application.h" // for MsgSleep()


HWND WinActivate(global_struct &aSettings, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText
	, bool aFindLastMatch, HWND aAlreadyVisited[], int aAlreadyVisitedCount)
{
	// If window is already active, be sure to leave it that way rather than activating some
	// other window that may match title & text also.  NOTE: An explicit check is done
	// for this rather than just relying on EnumWindows() to obey the z-order because
	// EnumWindows() is *not* guaranteed to enumerate windows in z-order, thus the currently
	// active window, even if it's an exact match, might become overlapped by another matching
	// window.  Also, use the USE_FOREGROUND_WINDOW vs. IF_USE_FOREGROUND_WINDOW macro for
	// this because the active window can sometimes be NULL (i.e. if it's a hidden window
	// and DetectHiddenWindows is off):
	HWND target_window;
	if (USE_FOREGROUND_WINDOW(aTitle, aText, aExcludeTitle, aExcludeText))
	{
		// User asked us to activate the "active" window, which by definition already is.
		// However, if the active (foreground) window is hidden and DetectHiddenWindows is
		// off, the below will set target_window to be NULL, which seems like the most
		// consistent result to use:
		SET_TARGET_TO_ALLOWABLE_FOREGROUND(aSettings.DetectHiddenWindows)
		return target_window;
	}

	if (!aFindLastMatch && !*aTitle && !*aText && !*aExcludeTitle && !*aExcludeText)
	{
		// User passed no params, so use the window most recently found by WinExist():
		if (   !(target_window = GetValidLastUsedWindow(aSettings))   )
			return NULL;
	}
	else
	{
		/*
		// Might not help avg. perfomance any?
		if (!aFindLastMatch) // Else even if the windows is already active, we want the bottomost one.
			if (hwnd = WinActive(aTitle, aText, aExcludeTitle, aExcludeText)) // Already active.
				return target_window;
		*/
		// Don't activate in this case, because the top-most window might be an
		// always-on-top but not-meant-to-be-activated window such as AutoIt's
		// splash text:
		if (   !(target_window = WinExist(aSettings, aTitle, aText, aExcludeTitle, aExcludeText, aFindLastMatch
			, false, aAlreadyVisited, aAlreadyVisitedCount))   )
			return NULL;
	}
	// Above has ensured that target_window is non-NULL, that it is a valid window, and that
	// it is eligible due to g->DetectHiddenWindows being true or the window not being hidden
	// (or being one of the script's GUI windows).
	return SetForegroundWindowEx(target_window);
}



#ifdef _DEBUG_WINACTIVATE
#define LOGF "c:\\AutoHotkey SetForegroundWindowEx.txt"
HWND AttemptSetForeground(HWND aTargetWindow, HWND aForeWindow, char *aTargetTitle)
#else
HWND AttemptSetForeground(HWND aTargetWindow, HWND aForeWindow)
#endif
// Returns NULL if aTargetWindow or its owned-window couldn't be brought to the foreground.
// Otherwise, on success, it returns either aTargetWindow or an HWND owned by aTargetWindow.
{
	// Probably best not to trust its return value.  It's been shown to be unreliable at times.
	// Example: I've confirmed that SetForegroundWindow() sometimes (perhaps about 10% of the time)
	// indicates failure even though it succeeds.  So we specifically check to see if it worked,
	// which helps to avoid using the keystroke (2-alts) method, because that may disrupt the
	// desired state of the keys or disturb any menus that the user may have displayed.
	// Also: I think the 2-alts last-resort may fire when the system is lagging a bit
	// (i.e. a drive spinning up) and the window hasn't actually become active yet,
	// even though it will soon become active on its own.  Also, SetForegroundWindow() sometimes
	// indicates failure even though it succeeded, usually because the window didn't become
	// active immediately -- perhaps because the system was under load -- but did soon become
	// active on its own (after, say, 50ms or so).  UPDATE: If SetForegroundWindow() is called
	// on a hung window, at least when AttachThreadInput is in effect and that window has
	// a modal dialog (such as MSIE's find dialog), this call might never return, locking up
	// our thread.  So now we do this fast-check for whether the window is hung first (and
	// this call is indeed very fast: its worst case is at least 30x faster than the worst-case
	// performance of the ABORT-IF-HUNG method used with SendMessageTimeout.
	// UPDATE for v1.0.42.03: To avoid a very rare crashing issue, IsWindowHung() is no longer called
	// here, but instead by our caller.  Search on "v1.0.42.03" for more comments.
	BOOL result = SetForegroundWindow(aTargetWindow);
	// Note: Increasing the sleep time below did not help with occurrences of "indicated success
	// even though it failed", at least with metapad.exe being activated while command prompt
	// and/or AutoIt2's InputBox were active or present on the screen:
	SLEEP_WITHOUT_INTERRUPTION(SLEEP_INTERVAL); // Specify param so that it will try to specifically sleep that long.
	HWND new_fore_window = GetForegroundWindow();
	if (new_fore_window == aTargetWindow)
	{
#ifdef _DEBUG_WINACTIVATE
		if (!result)
		{
			FileAppend(LOGF, "SetForegroundWindow() indicated failure even though it succeeded: ", false);
			FileAppend(LOGF, aTargetTitle);
		}
#endif
		return aTargetWindow;
	}
	if (new_fore_window != aForeWindow && aTargetWindow == GetWindow(new_fore_window, GW_OWNER))
		// The window we're trying to get to the foreground is the owner of the new foreground window.
		// This is considered to be a success because a window that owns other windows can never be
		// made the foreground window, at least if the windows it owns are visible.
		return new_fore_window;
	// Otherwise, failure:
#ifdef _DEBUG_WINACTIVATE
	if (result)
	{
		FileAppend(LOGF, "SetForegroundWindow() indicated success even though it failed: ", false);
		FileAppend(LOGF, aTargetTitle);
	}
#endif
	return NULL;
}



HWND SetForegroundWindowEx(HWND aTargetWindow)
// Caller must have ensured that aTargetWindow is a valid window or NULL, since we
// don't call IsWindow() here.
{
	if (!aTargetWindow)
		return NULL;  // When called this way (as it is sometimes), do nothing.

	// v1.0.42.03: Calling IsWindowHung() once here rather than potentially more than once in AttemptSetForeground()
	// solves a crash that is not fully understood, nor is it easily reproduced (it occurs only in release mode,
	// not debug mode).  It's likely a bug in the API's IsHungAppWindow(), but that is far from confirmed.
	DWORD target_thread = GetWindowThreadProcessId(aTargetWindow, NULL);
	if (target_thread != g_MainThreadID && IsWindowHung(aTargetWindow)) // Calls to IsWindowHung should probably be avoided if the window belongs to our thread.  Relies upon short-circuit boolean order.
		return NULL;

#ifdef _DEBUG_WINACTIVATE
	char win_name[64];
	GetWindowText(aTargetWindow, win_name, sizeof(win_name));
#endif

	HWND orig_foreground_wnd = GetForegroundWindow();
	// AutoIt3: If there is not any foreground window, then input focus is on the TaskBar.
	// MY: It is definitely possible for GetForegroundWindow() to return NULL, even on XP.
	if (!orig_foreground_wnd)
		orig_foreground_wnd = FindWindow("Shell_TrayWnd", NULL);

	if (aTargetWindow == orig_foreground_wnd) // It's already the active window.
		return aTargetWindow;

	if (IsIconic(aTargetWindow))
		// This might never return if aTargetWindow is a hung window.  But it seems better
		// to do it this way than to use the PostMessage() method, which might not work
		// reliably with apps that don't handle such messages in a standard way.
		// A minimized window must be restored or else SetForegroundWindow() always(?)
		// won't work on it.  UPDATE: ShowWindowAsync() would prevent a hang, but
		// probably shouldn't use it because we rely on the fact that the message
		// has been acted on prior to trying to activate the window (and all Async()
		// does is post a message to its queue):
		ShowWindow(aTargetWindow, SW_RESTORE);

	// This causes more trouble than it's worth.  In fact, the AutoIt author said that
	// he didn't think it even helped with the IE 5.5 related issue it was originally
	// intended for, so it seems a good idea to NOT to this, especially since I'm 80%
	// sure it messes up the Z-order in certain circumstances, causing an unexpected
	// window to pop to the foreground immediately after a modal dialog is dismissed:
	//BringWindowToTop(aTargetWindow); // AutoIt3: IE 5.5 related hack.

	HWND new_foreground_wnd;

	if (!g_WinActivateForce)
	// if (g_os.IsWin95() || (!g_os.IsWin9x() && !g_os.IsWin2000orLater())))  // Win95 or NT
		// Try a simple approach first for these two OS's, since they don't have
		// any restrictions on focus stealing:
#ifdef _DEBUG_WINACTIVATE
#define IF_ATTEMPT_SET_FORE if (new_foreground_wnd = AttemptSetForeground(aTargetWindow, orig_foreground_wnd, win_name))
#else
#define IF_ATTEMPT_SET_FORE if (new_foreground_wnd = AttemptSetForeground(aTargetWindow, orig_foreground_wnd))
#endif
		IF_ATTEMPT_SET_FORE
			return new_foreground_wnd;
		// Otherwise continue with the more drastic methods below.

	// MY: The AttachThreadInput method, when used by itself, seems to always
	// work the first time on my XP system, seemingly regardless of whether the
	// "allow focus steal" change has been made via SystemParametersInfo()
	// (but it seems a good idea to keep the SystemParametersInfo() in effect
	// in case Win2k or Win98 needs it, or in case it really does help in rare cases).
	// In many cases, this avoids the two SetForegroundWindow() attempts that
	// would otherwise be needed; and those two attempts cause some windows
	// to flash in the taskbar, such as Metapad and Excel (less frequently) whenever
	// you quickly activate another window after activating it first (e.g. via hotkeys).
	// So for now, it seems best just to use this method by itself.  The
	// "two-alts" case never seems to fire on my system?  Maybe it will
	// on Win98 sometimes.
	// Note: In addition to the "taskbar button flashing" annoyance mentioned above
	// any SetForegroundWindow() attempt made prior to the one below will,
	// as a side-effect, sometimes trigger the need for the "two-alts" case
	// below.  So that's another reason to just keep it simple and do it this way
	// only.

#ifdef _DEBUG_WINACTIVATE
	char buf[1024];
#endif

	bool is_attached_my_to_fore = false, is_attached_fore_to_target = false;
	DWORD fore_thread;
	if (orig_foreground_wnd) // Might be NULL from above.
	{
		// Based on MSDN docs, these calls should always succeed due to the other
		// checks done above (e.g. that none of the HWND's are NULL):
		fore_thread = GetWindowThreadProcessId(orig_foreground_wnd, NULL);

		// MY: Normally, it's suggested that you only need to attach the thread of the
		// foreground window to our thread.  However, I've confirmed that doing all three
		// attaches below makes the attempt much more likely to succeed.  In fact, it
		// almost always succeeds whereas the one-attach method hardly ever succeeds the first
		// time (resulting in a flashing taskbar button due to having to invoke a second attempt)
		// when one window is quickly activated after another was just activated.
		// AutoIt3: Attach all our input threads, will cause SetForeground to work under 98/Me.
		// MSDN docs: The AttachThreadInput function fails if either of the specified threads
		// does not have a message queue (My: ok here, since any window's thread MUST have a
		// message queue).  [It] also fails if a journal record hook is installed.  ... Note
		// that key state, which can be ascertained by calls to the GetKeyState or
		// GetKeyboardState function, is reset after a call to AttachThreadInput.  You cannot
		// attach a thread to a thread in another desktop.  A thread cannot attach to itself.
		// Therefore, idAttachTo cannot equal idAttach.  Update: It appears that of the three,
		// this first call does not offer any additional benefit, at least on XP, so not
		// using it for now:
		//if (g_MainThreadID != target_thread) // Don't attempt the call otherwise.
		//	AttachThreadInput(g_MainThreadID, target_thread, TRUE);
		if (fore_thread && g_MainThreadID != fore_thread && !IsWindowHung(orig_foreground_wnd))
			is_attached_my_to_fore = AttachThreadInput(g_MainThreadID, fore_thread, TRUE) != 0;
		if (fore_thread && target_thread && fore_thread != target_thread) // IsWindowHung(aTargetWindow) was called earlier.
			is_attached_fore_to_target = AttachThreadInput(fore_thread, target_thread, TRUE) != 0;
	}

	// The log showed that it never seemed to need more than two tries.  But there's
	// not much harm in trying a few extra times.  The number of tries needed might
	// vary depending on how fast the CPU is:
	for (int i = 0; i < 5; ++i)
	{
		IF_ATTEMPT_SET_FORE
		{
#ifdef _DEBUG_WINACTIVATE
			if (i > 0) // More than one attempt was needed.
			{
				snprintf(buf, sizeof(buf), "AttachThreadInput attempt #%d indicated success: %s"
					, i + 1, win_name);
				FileAppend(LOGF, buf);
			}
#endif
			break;
		}
	}

	// I decided to avoid the quick minimize + restore method of activation.  It's
	// not that much more effective (if at all), and there are some significant
	// disadvantages:
	// - This call will often hang our thread if aTargetWindow is a hung window: ShowWindow(aTargetWindow, SW_MINIMIZE)
	// - Using SW_FORCEMINIMIZE instead of SW_MINIMIZE has at least one (and probably more)
	// side effect: When the window is restored, at least via SW_RESTORE, it is no longer
	// maximized even if it was before the minmize.  So don't use it.
	if (!new_foreground_wnd) // Not successful yet.
	{
		// Some apps may be intentionally blocking us by having called the API function
		// LockSetForegroundWindow(), for which MSDN says "The system automatically enables
		// calls to SetForegroundWindow if the user presses the ALT key or takes some action
		// that causes the system itself to change the foreground window (for example,
		// clicking a background window)."  Also, it's probably best to avoid doing
		// the 2-alts method except as a last resort, because I think it may mess up
		// the state of menus the user had displayed.  And of course if the foreground
		// app has special handling for alt-key events, it might get confused.
		// My original note: "The 2-alts case seems to mess up on rare occasions,
		// perhaps due to menu weirdness triggered by the alt key."
		// AutoIt3: OK, this is not funny - bring out the extreme measures (usually for 2000/XP).
		// Simulate two single ALT keystrokes.  UPDATE: This hardly ever succeeds.  Usually when
		// it fails, the foreground window is NULL (none).  I'm going to try an Win-tab instead,
		// which selects a task bar button.  This seems less invasive than doing an alt-tab
		// because not only doesn't it activate some other window first, it also doesn't appear
		// to change the Z-order, which is good because we don't want the alt-tab order
		// that the user sees to be affected by this.  UPDATE: Win-tab isn't doing it, so try
		// Alt-tab.  Alt-tab doesn't do it either.  The window itself (metapad.exe is the only
		// culprit window I've found so far) seems to resist being brought to the foreground,
		// but later, after the hotkey is released, it can be.  So perhaps this is being
		// caused by the fact that the user has keys held down (logically or physically?)
		// Releasing those keys with a key-up event might help, so try that sometime:
		KeyEvent(KEYDOWNANDUP, VK_MENU);
		KeyEvent(KEYDOWNANDUP, VK_MENU);
		//KeyEvent(KEYDOWN, VK_LWIN);
		//KeyEvent(KEYDOWN, VK_TAB);
		//KeyEvent(KEYUP, VK_TAB);
		//KeyEvent(KEYUP, VK_LWIN);
		//KeyEvent(KEYDOWN, VK_MENU);
		//KeyEvent(KEYDOWN, VK_TAB);
		//KeyEvent(KEYUP, VK_TAB);
		//KeyEvent(KEYUP, VK_MENU);
		// Also replacing "2-alts" with "alt-tab" below, for now:

#ifndef _DEBUG_WINACTIVATE
		new_foreground_wnd = AttemptSetForeground(aTargetWindow, orig_foreground_wnd);
#else // debug mode
		IF_ATTEMPT_SET_FORE
			FileAppend(LOGF, "2-alts ok: ", false);
		else
		{
			FileAppend(LOGF, "2-alts (which is the last resort) failed.  ", false);
			HWND h = GetForegroundWindow();
			if (h)
			{
				char fore_name[64];
				GetWindowText(h, fore_name, sizeof(fore_name));
				FileAppend(LOGF, "Foreground: ", false);
				FileAppend(LOGF, fore_name, false);
			}
			FileAppend(LOGF, ".  Was trying to activate: ", false);
		}
		FileAppend(LOGF, win_name);
#endif
	} // if()

	// Very important to detach any threads whose inputs were attached above,
	// prior to returning, otherwise the next attempt to attach thread inputs
	// for these particular windows may result in a hung thread or other
	// undesirable effect:
	if (is_attached_my_to_fore)
		AttachThreadInput(g_MainThreadID, fore_thread, FALSE);
	if (is_attached_fore_to_target)
		AttachThreadInput(fore_thread, target_thread, FALSE);

	// Finally.  This one works, solving the problem of the MessageBox window
	// having the input focus and being the foreground window, but not actually
	// being visible (even though IsVisible() and IsIconic() say it is)!  It may
	// help with other conditions under which this function would otherwise fail.
	// Here's the way the repeat the failure to test how the absence of this line
	// affects things, at least on my XP SP1 system:
	// y::MsgBox, test
	// #e::(some hotkey that activates Windows Explorer)
	// Now: Activate explorer with the hotkey, then invoke the MsgBox.  It will
	// usually be activated but invisible.  Also: Whenever this invisible problem
	// is about to occur, with or without this fix, it appears that the OS's z-order
	// is a bit messed up, because when you dismiss the MessageBox, an unexpected
	// window (probably the one two levels down) becomes active rather than the
	// window that's only 1 level down in the z-order:
	if (new_foreground_wnd) // success.
	{
		// Even though this is already done for the IE 5.5 "hack" above, must at
		// a minimum do it here: The above one may be optional, not sure (safest
		// to leave it unless someone can test with IE 5.5).
		// Note: I suspect the two lines below achieve the same thing.  They may
		// even be functionally identical.  UPDATE: This may no longer be needed
		// now that the first BringWindowToTop(), above, has been disabled due to
		// its causing more trouble than it's worth.  But seems safer to leave
		// this one enabled in case it does resolve IE 5.5 related issues and
		// possible other issues:
		BringWindowToTop(aTargetWindow);
		//SetWindowPos(aTargetWindow, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
		return new_foreground_wnd; // Return this rather than aTargetWindow because it's more appropriate.
	}
	else
		return NULL;
}



HWND WinClose(global_struct &aSettings, char *aTitle, char *aText, int aTimeToWaitForClose
	, char *aExcludeTitle, char *aExcludeText, bool aKillIfHung)
// Return the HWND of any found-window to the caller so that it has the option of waiting
// for it to become an invalid (closed) window.
{
	HWND target_window;
	IF_USE_FOREGROUND_WINDOW(aSettings.DetectHiddenWindows, aTitle, aText, aExcludeTitle, aExcludeText)
		// Close topmost (better than !F4 since that uses the alt key, effectively resetting
		// its status to UP if it was down before.  Use WM_CLOSE rather than WM_EXIT because
		// I think that's what Alt-F4 sends (and otherwise, app may quit without offering
		// a chance to save).
		// DON'T DISPLAY a MsgBox (e.g. debugging) before trying to close foreground window.
		// Otherwise, it may close the owner of the dialog window (this app), perhaps due to
		// split-second timing issues.
	else if (*aTitle || *aText || *aExcludeTitle || *aExcludeText)
	{
		// Since EnumWindows() is *not* guaranteed to start proceed in z-order from topmost to
		// bottomost (though it almost certainly does), do it this way to ensure that the
		// topmost window is closed in preference to any other windows with the same <aTitle>
		// and <aText>:
		if (   !(target_window = WinActive(aSettings, aTitle, aText, aExcludeTitle, aExcludeText))   )
			if (   !(target_window = WinExist(aSettings, aTitle, aText, aExcludeTitle, aExcludeText))   )
				return NULL;
	}
	else
		target_window = GetValidLastUsedWindow(aSettings);

	return target_window ? WinClose(target_window, aTimeToWaitForClose, aKillIfHung) : NULL;
}



HWND WinClose(HWND aWnd, int aTimeToWaitForClose, bool aKillIfHung)
{
	if (aKillIfHung) // This part is based on the AutoIt3 source.
		// Update: Another reason not to wait a long time with the below is that WinKill
		// is normally only used when the target window is suspected of being hung.  It
		// seems bad to wait something like 2 seconds in such a case, when the caller
		// probably already knows it's hung.
		// Obsolete in light of dedicated hook thread: Because this app is much more sensitive to being
		// in a "not-pumping-messages" state, due to the keyboard & mouse hooks, it seems better to wait
		// for only 200 ms (e.g. in case the user is gaming and there's a script
		// running in the background that uses WinKill, we don't want key and mouse events
		// to freeze for a long time).
		Util_WinKill(aWnd);
	else // Don't kill.
		// SC_CLOSE is the same as clicking a window's "X"(close) button or using Alt-F4.
		// Although it's a more friendly way to close windows than WM_CLOSE (and thus
		// avoids incompatibilities with apps such as MS Visual C++), apps that
		// have disabled Alt-F4 processing will not be successfully closed.  It seems
		// best not to send both SC_CLOSE and WM_CLOSE because some apps with an 
		// "Unsaved.  Are you sure?" type dialog might close down completely rather than
		// waiting for the user to confirm.  Anyway, it's extrememly rare for a window
		// not to respond to Alt-F4 (though it is possible that it handles Alt-F4 in a
		// non-standard way, i.e. that sending SC_CLOSE equivalent to Alt-F4
		// for windows that handle Alt-F4 manually?)  But on the upside, this is nicer
		// for apps that upon receiving Alt-F4 do some behavior other than closing, such
		// as minimizing to the tray.  Such apps might shut down entirely if they received
		// a true WM_CLOSE, which is probably not what the user would want.
		// Update: Swithced back to using WM_CLOSE so that instances of AutoHotkey
		// can be terminated via another instances use of the WinClose command:
		//PostMessage(aWnd, WM_SYSCOMMAND, SC_CLOSE, 0);
		PostMessage(aWnd, WM_CLOSE, 0, 0);

	if (aTimeToWaitForClose < 0)
		aTimeToWaitForClose = 0;
	if (!aTimeToWaitForClose)
		return aWnd; // v1.0.30: EnumParentActUponAll() relies on the ability to do no delay at all.

	// Slight delay.  Might help avoid user having to modify script to use WinWaitClose()
	// in many cases.  UPDATE: But this does a Sleep(0), which won't yield the remainder
	// of our thread's timeslice unless there's another app trying to use 100% of the CPU?
	// So, in reality it really doesn't accomplish anything because the window we just
	// closed won't get any CPU time (unless, perhaps, it receives the close message in
	// time to ask the OS for us to yield the timeslice).  Perhaps some finer tuning
	// of this can be done in the future.  UPDATE: Testing of WinActivate, which also
	// uses this to do a Sleep(0), reveals that it may in fact help even when the CPU
	// isn't under load.  Perhaps this is because upon Sleep(0), the OS runs the
	// WindowProc's of windows that have messages waiting for them so that appropriate
	// action can be taken (which may often be nearly instantaneous, perhaps under
	// 1ms for a Window to be logically destroyed even if it hasn't physically been
	// removed from the screen?) prior to returning the CPU to our thread:
	DWORD start_time = GetTickCount(); // Before doing any MsgSleeps, set this.
    //MsgSleep(0); // Always do one small one, see above comments.
	// UPDATE: It seems better just to always do one unspecified-interval sleep
	// rather than MsgSleep(0), which often returns immediately, probably having
	// no effect.

	// Remember that once the first call to MsgSleep() is done, a new hotkey subroutine
	// may fire and suspend what we're doing here.  Such a subroutine might also overwrite
	// the values our params, some of which may be in the deref buffer.  So be sure not
	// to refer to those strings once MsgSleep() has been done, below:

	// This is the same basic code used for ACT_WINWAITCLOSE and such:
	for (;;)
	{
		// Seems best to always do the first one regardless of the value 
		// of aTimeToWaitForClose:
		MsgSleep(INTERVAL_UNSPECIFIED);
		if (!IsWindow(aWnd)) // It's gone, so we're done.
			return aWnd;
		// Must cast to int or any negative result will be lost due to DWORD type:
		if ((int)(aTimeToWaitForClose - (GetTickCount() - start_time)) <= SLEEP_INTERVAL_HALF)
			break;
			// Last param 0 because we don't want it to restore the
			// current active window after the time expires (in case
			// it's suspended).  INTERVAL_UNSPECIFIED performs better.
	}
	return aWnd;  // Done waiting.
}


	
HWND WinActive(global_struct &aSettings, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText
	, bool aUpdateLastUsed)
// This function must be kept thread-safe because it may be called (indirectly) by hook thread too.
// In addition, it must not change the value of anything in aSettings except when aUpdateLastUsed==true.
{
	HWND target_window;
	if (USE_FOREGROUND_WINDOW(aTitle, aText, aExcludeTitle, aExcludeText))
	{
		// User asked us if the "active" window is active, which is true if it's not a
		// hidden window, or if DetectHiddenWindows is ON:
		SET_TARGET_TO_ALLOWABLE_FOREGROUND(aSettings.DetectHiddenWindows)
		#define UPDATE_AND_RETURN_LAST_USED_WINDOW(hwnd) \
		{\
			if (aUpdateLastUsed && hwnd)\
				aSettings.hWndLastUsed = hwnd;\
			return hwnd;\
		}
		UPDATE_AND_RETURN_LAST_USED_WINDOW(target_window)
	}

	HWND fore_win = GetForegroundWindow();
	if (!fore_win)
		return NULL;

	if (!(*aTitle || *aText || *aExcludeTitle || *aExcludeText)) // Use the "last found" window.
		return (fore_win == GetValidLastUsedWindow(aSettings)) ? fore_win : NULL;

	// Only after the above check should the below be done.  This is because "IfWinActive" (with no params)
	// should be "true" if one of the script's GUI windows is active:
	if (!(aSettings.DetectHiddenWindows || IsWindowVisible(fore_win))) // In this case, the caller's window can't be active.
		return NULL;

	WindowSearch ws;
	ws.SetCandidate(fore_win);

	if (ws.SetCriteria(aSettings, aTitle, aText, aExcludeTitle, aExcludeText) && ws.IsMatch()) // aSettings.DetectHiddenWindows was already checked above.
		UPDATE_AND_RETURN_LAST_USED_WINDOW(fore_win) // This also does a "return".
	else // If the line above didn't return, indicate that the specified window is not active.
		return NULL;
}



HWND WinExist(global_struct &aSettings, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText
	, bool aFindLastMatch, bool aUpdateLastUsed, HWND aAlreadyVisited[], int aAlreadyVisitedCount)
// This function must be kept thread-safe because it may be called (indirectly) by hook thread too.
// In addition, it must not change the value of anything in aSettings except when aUpdateLastUsed==true.
{
	HWND target_window;
	if (USE_FOREGROUND_WINDOW(aTitle, aText, aExcludeTitle, aExcludeText))
	{
		// User asked us if the "active" window exists, which is true if it's not a
		// hidden window or DetectHiddenWindows is ON:
		SET_TARGET_TO_ALLOWABLE_FOREGROUND(aSettings.DetectHiddenWindows)
		// Updating LastUsed to be hwnd even if it's NULL seems best for consistency?
		// UPDATE: No, it's more flexible not to never set it to NULL, because there
		// will be times when the old value is still useful:
		UPDATE_AND_RETURN_LAST_USED_WINDOW(target_window);
	}

	if (!(*aTitle || *aText || *aExcludeTitle || *aExcludeText))
		// User passed no params, so use the window most recently found by WinExist().
		// It's correct to do this even in this function because it's called by
		// WINWAITCLOSE and IFWINEXIST specifically to discover if the Last-Used
		// window still exists.
		return GetValidLastUsedWindow(aSettings);

	WindowSearch ws;
	if (!ws.SetCriteria(aSettings, aTitle, aText, aExcludeTitle, aExcludeText)) // No match is possible with these criteria.
		return NULL;

	ws.mFindLastMatch = aFindLastMatch;
	ws.mAlreadyVisited = aAlreadyVisited;
	ws.mAlreadyVisitedCount = aAlreadyVisitedCount;

	if (ws.mCriteria & CRITERION_ID) // "ahk_id" will be satisified if that HWND still exists and is valid.
	{
		// Explicitly allow HWND_BROADCAST for all commands that use WinExist (which is just about all
		// window commands), even though it's only valid with ScriptPostSendMessage().
		// This is because HWND_BROADCAST is probably never used as the HWND for a real window, so there
		// should be no danger of any reasonable script ever passing that value in as a real target window,
		// which should thus minimize the chance of a crash due to calling various API functions
		// with invalid window handles.
		if (   ws.mCriterionHwnd != HWND_BROADCAST // It's not exempt from the other checks on the two lines below.
			&& (!IsWindow(ws.mCriterionHwnd)    // And it's either not a valid window...
				// ...or the window is not detectible (in v1.0.40.05, child windows are detectible even if hidden):
				|| !(aSettings.DetectHiddenWindows || IsWindowVisible(ws.mCriterionHwnd)
					|| (GetWindowLong(ws.mCriterionHwnd, GWL_STYLE) & WS_CHILD)))   )
			return NULL;

		// Otherwise, the window is valid and detectible.
		ws.SetCandidate(ws.mCriterionHwnd);
		if (!ws.IsMatch()) // Checks if it matches any other criteria: WinTitle, WinText, ExcludeTitle, and anything in the aAlreadyVisited list.
			return NULL;
		//else fall through to the section below, since ws.mFoundCount and ws.mFoundParent were set by ws.IsMatch().
	}
	else // aWinTitle doesn't start with "ahk_id".  Try to find a matching window.
		EnumWindows(EnumParentFind, (LPARAM)&ws);

	UPDATE_AND_RETURN_LAST_USED_WINDOW(ws.mFoundParent) // This also does a "return".
}



HWND GetValidLastUsedWindow(global_struct &aSettings)
// If the last found window is one of the script's own GUI windows, it is considered valid even if
// DetectHiddenWindows is ON.  Note that this exemption does not apply to things like "IfWinExist,
// My Gui Title", "WinActivate, ahk_id <gui id>", etc.
// A GUI window can become the last found window while DetectHiddenWindows is ON in two ways:
// Gui +LastFound
// The launch of a GUI thread that explicitly set the last found window to be that GUI window.
{
	if (!aSettings.hWndLastUsed || !IsWindow(aSettings.hWndLastUsed))
		return NULL;
	if (   aSettings.DetectHiddenWindows || IsWindowVisible(aSettings.hWndLastUsed)
		|| (GetWindowLong(aSettings.hWndLastUsed, GWL_STYLE) & WS_CHILD)   ) // v1.0.40.05: Child windows (via ahk_id) are always detectible.
		return aSettings.hWndLastUsed;
	// Otherwise, DetectHiddenWindows is OFF and the window is not visible.  Return NULL
	// unless this is a GUI window belonging to this particular script, in which case
	// the setting of DetectHiddenWindows is ignored (as of v1.0.25.13).
	return GuiType::FindGui(aSettings.hWndLastUsed) ? aSettings.hWndLastUsed : NULL;
}



BOOL CALLBACK EnumParentFind(HWND aWnd, LPARAM lParam)
// This function must be kept thread-safe because it may be called (indirectly) by hook thread too.
// To continue enumeration, the function must return TRUE; to stop enumeration, it must return FALSE. 
// It's a little strange, but I think EnumWindows() returns FALSE when the callback stopped
// the enumeration prematurely by returning false to its caller.  Otherwise (the enumeration went
// through every window), it returns TRUE:
{
	WindowSearch &ws = *(WindowSearch *)lParam;  // For performance and convenience.
	// According to MSDN, GetWindowText() will hang only if it's done against
	// one of your own app's windows and that window is hung.  I suspect
	// this might not be true in Win95, and possibly not even Win98, but
	// it's not really an issue because GetWindowText() has to be called
	// eventually, either here or in an EnumWindowsProc.  The only way
	// to prevent hangs (if indeed it does hang on Win9x) would be to
	// call something like IsWindowHung() before every call to
	// GetWindowText(), which might result in a noticeable delay whenever
	// we search for a window via its title (or even worse: by the title
	// of one of its controls or child windows).  UPDATE: Trying GetWindowTextTimeout()
	// now, which might be the best compromise.  UPDATE: It's annoyingly slow,
	// so went back to using the old method.
	if (!(ws.mSettings->DetectHiddenWindows || IsWindowVisible(aWnd))) // Skip windows the script isn't supposed to detect.
		return TRUE;
	ws.SetCandidate(aWnd);
	// If this window doesn't match, continue searching for more windows (via TRUE).  Likewise, if
	// mFindLastMatch is true, continue searching even if this window is a match.  Otherwise, this
	// first match is the one that's desired so stop here:
	return ws.IsMatch() ? ws.mFindLastMatch : TRUE;
}



BOOL CALLBACK EnumChildFind(HWND aWnd, LPARAM lParam)
// This function must be kept thread-safe because it may be called (indirectly) by hook thread too.
// Although this function could be rolled into a generalized version of the EnumWindowsProc(),
// it will perform better this way because there's less checking required and no mode/flag indicator
// is needed inside lParam to indicate which struct element should be searched for.  In addition,
// it's more comprehensible this way.  lParam is a pointer to the struct rather than just a
// string because we want to give back the HWND of any matching window.
{
	// Since WinText and ExcludeText are seldom used in typical scripts, the following buffer
	// is put on the stack here rather than on our callers (inside the WindowSearch object),
	// which should help conserve stack space on average.  Can't use the ws.mCandidateTitle
	// buffer because ws.mFindLastMatch might be true, in which case the original title must
	// be preserved.
	char win_text[WINDOW_TEXT_SIZE];
	WindowSearch &ws = *(WindowSearch *)lParam;  // For performance and convenience.

	if (!(ws.mSettings->DetectHiddenText || IsWindowVisible(aWnd))) // This text element should not be detectible by the script.
		return TRUE;  // Skip this child and keep enumerating to try to find a match among the other children.

	// The below was formerly outsourced to the following function, but since it is only called from here,
	// it has been moved inline:
	// int GetWindowTextByTitleMatchMode(HWND aWnd, char *aBuf = NULL, int aBufSize = 0)
	int text_length = ws.mSettings->TitleFindFast ? GetWindowText(aWnd, win_text, sizeof(win_text))
		: GetWindowTextTimeout(aWnd, win_text, sizeof(win_text));  // The slower method that is able to get text from more types of controls (e.g. large edit controls).
	// Older idea that for the above that was not adopted:
	// Only if GetWindowText() gets 0 length would we try the other method (and of course, don't bother
	// using GetWindowTextTimeout() at all if "fast" mode is in effect).  The problem with this is that
	// many controls always return 0 length regardless of which method is used, so this would slow things
	// down a little (but not too badly since GetWindowText() is so much faster than GetWindowTextTimeout()).
	// Another potential problem is that some controls may return less text, or different text, when used
	// with the fast mode vs. the slow mode (unverified).  So it seems best NOT to do this and stick with
	// the simple approach above.
	if (!text_length) // It has no text (or failure to fetch it).
		return TRUE;  // Skip this child and keep enumerating to try to find a match among the other children.

	// For compatibility with AutoIt v2, strstr() is always used for control/child text elements.

	// EXCLUDE-TEXT: The following check takes precedence over the next, so it's done first:
	if (*ws.mCriterionExcludeText) // For performance, avoid doing the checks below when blank.
	{
		if (ws.mSettings->TitleMatchMode == FIND_REGEX)
		{
			if (RegExMatch(win_text, ws.mCriterionExcludeText))
				return FALSE; // Parent can't be a match, so stop searching its children.
		}
		else // For backward compatibility, all modes other than RegEx behave as follows.
			if (strstr(win_text, ws.mCriterionExcludeText))
				// Since this child window contains the specified ExcludeText anywhere inside its text,
				// the parent window is always a non-match.
				return FALSE; // Parent can't be a match, so stop searching its children.
	}

	// WIN-TEXT:
	if (!*ws.mCriterionText) // Match always found in this case. This check is for performance: it avoids doing the checks below when not needed, especially RegEx. Note: It's possible for mCriterionText to be blank, at least when mCriterionExcludeText isn't blank.
	{
		ws.mFoundChild = aWnd;
		return FALSE; // Match found, so stop searching.
	}
	if (ws.mSettings->TitleMatchMode == FIND_REGEX)
	{
		if (RegExMatch(win_text, ws.mCriterionText)) // Match found.
		{
			ws.mFoundChild = aWnd;
			return FALSE; // Match found, so stop searching.
		}
	}
	else // For backward compatibility, all modes other than RegEx behave as follows.
		if (strstr(win_text, ws.mCriterionText)) // Match found.
		{
			ws.mFoundChild = aWnd;
			return FALSE; // Match found, so stop searching.
		}

	// UPDATE to the below: The MSDN docs state that EnumChildWindows() already handles the
	// recursion for us: "If a child window has created child windows of its own,
	// EnumChildWindows() enumerates those windows as well."
	// Mostly obsolete comments: Since this child doesn't match, make sure none of its
	// children (recursive) match prior to continuing the original enumeration.  We don't
	// discard the return value from EnumChildWindows() because it's FALSE in two cases:
	// 1) The given HWND has no children.
	// 2) The given EnumChildProc() stopped prematurely rather than enumerating all the windows.
	// and there's no way to distinguish between the two cases without using the
	// struct's hwnd because GetLastError() seems to return ERROR_SUCCESS in both
	// cases.
	//EnumChildWindows(aWnd, EnumChildFind, lParam);
	// If matching HWND still hasn't been found, return TRUE to keep searching:
	//return ws.mFoundChild == NULL;

	return TRUE; // Keep searching.
}



ResultType StatusBarUtil(Var *aOutputVar, HWND aBarHwnd, int aPartNumber, char *aTextToWaitFor
	, int aWaitTime, int aCheckInterval)
// aOutputVar is allowed to be NULL if aTextToWaitFor isn't NULL or blank. aBarHwnd is allowed
// to be NULL because in that case, the caller wants us to set ErrorLevel appropriately and also
// make aOutputVar empty.
{
	if (aOutputVar)
		aOutputVar->Assign(); // Init to blank in case of early return.
	// Set default ErrorLevel, which is a special value (2 vs. 1) in the case of StatusBarWait:
	g_ErrorLevel->Assign(aOutputVar ? ERRORLEVEL_ERROR : ERRORLEVEL_ERROR2);

	// Legacy: Waiting 500ms in place of a "0" seems more useful than a true zero, which doens't need
	// to be supported because it's the same thing as something like "IfWinExist":
	if (!aWaitTime)
		aWaitTime = 500;
	if (aCheckInterval < 1)
		aCheckInterval = SB_DEFAULT_CHECK_INTERVAL; // Caller relies on us doing this.
	if (aPartNumber < 1)
		aPartNumber = 1;  // Caller relies on us to set default in this case.

	// Must have at least one of these.  UPDATE: We want to allow this so that the command can be
	// used to wait for the status bar text to become blank:
	//if (!aOutputVar && !*aTextToWaitFor) return OK;

	// Whenever using SendMessageTimeout(), our app will be unresponsive until
	// the call returns, since our message loop isn't running.  In addition,
	// if the keyboard or mouse hook is installed, the input events will lag during
	// this call.  So keep the timeout value fairly short.  Update for v1.0.24:
	// There have been at least two reports of the StatusBarWait command ending
	// prematurely with an ErrorLevel of 2.  The most likely culprit is the below,
	// which has now been increased from 100 to 2000:
	#define SB_TIMEOUT 2000

	HANDLE handle;
	LPVOID remote_buf;
	LRESULT part_count; // The number of parts this status bar has.
	if (!aBarHwnd  // These conditions rely heavily on short-circuit boolean order.
		|| !SendMessageTimeout(aBarHwnd, SB_GETPARTS, 0, 0, SMTO_ABORTIFHUNG, SB_TIMEOUT, (PDWORD_PTR)&part_count) // It failed or timed out.
		|| aPartNumber > part_count
		|| !(remote_buf = AllocInterProcMem(handle, WINDOW_TEXT_SIZE + 1, aBarHwnd))) // Alloc mem last.
		return OK; // Let ErrorLevel tell the story.

	char buf_for_nt[WINDOW_TEXT_SIZE + 1]; // Needed only for NT/2k/XP: the local counterpart to the buf allocated remotely above.
	bool is_win9x = g_os.IsWin9x();
	char *local_buf = is_win9x ? (char *)remote_buf : buf_for_nt; // Local is the same as remote for Win9x.

	DWORD result, start_time;
	--aPartNumber; // Convert to zero-based for use below.

	// Always do the first iteration so that at least one check is done.  Also,  start_time is initialized
	// unconditionally in the name of code size reduction (it's a low overhead call):
	for (*local_buf = '\0', start_time = GetTickCount();;)
	{
		// MSDN recommends always checking the length of the bar text.  It implies that the length is
		// unrestricted, so a crash due to buffer overflow could otherwise occur:
		if (SendMessageTimeout(aBarHwnd, SB_GETTEXTLENGTH, aPartNumber, 0, SMTO_ABORTIFHUNG, SB_TIMEOUT, &result))
		{
			// Testing confirms that LOWORD(result) [the length] does not include the zero terminator.
			if (LOWORD(result) > WINDOW_TEXT_SIZE) // Text would be too large (very unlikely but good to check for security).
				break; // Abort the operation and leave ErrorLevel set to its default to indicate the problem.
			// Retrieve the bar's text:
			if (SendMessageTimeout(aBarHwnd, SB_GETTEXT, aPartNumber, (LPARAM)remote_buf, SMTO_ABORTIFHUNG, SB_TIMEOUT, &result))
			{
				if (!is_win9x)
				{
					if (!ReadProcessMemory(handle, remote_buf, local_buf, LOWORD(result) + 1, NULL)) // +1 to include the terminator (verified: length doesn't include zero terminator).
					{
						// Fairly critical error (though rare) so seems best to abort.
						*local_buf = '\0';  // In case it changed the buf before failing.
						break;
					}
				}
				//else Win9x, in which case the local and remote buffers are the same (no copying is needed).

				// Check if the retrieved text matches the caller's criteria. In addition to
				// normal/intuitive matching, a match is also achieved if both are empty strings.
				// In fact, IsTextMatch() yields "true" whenever aTextToWaitFor is the empty string:
				if (IsTextMatch(local_buf, aTextToWaitFor))
				{
					g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate "match found".
					break;
				}
			}
			//else SB_GETTEXT msg timed out or failed.  Leave local_buf unaltered.  See comment below.
		}
		//else SB_GETTEXTLENGTH msg timed out or failed.  For v1.0.37, continue to wait (if other checks
		// say its okay below) rather than aborting the operation.  This should help prevent an abort
		// when the target window (or the entire system) is unresponsive for a long time, perhaps due
		// to a drive spinning up, etc.

		// Only when above didn't break are the following secondary conditions checked.  When aOutputVar
		// is non-NULL, the caller wanted a single check only (no waiting) [however, in most such cases,
		// the checking above would already have done a "break" because of aTextToWaitFor being blank when
		// passed to IsTextMatch()].  Also, don't continue to wait if the status bar no longer exists
		// (which is usually caused by the parent window having been destroyed):
		if (aOutputVar || !IsWindow(aBarHwnd))
			break; // Leave ErrorLevel at its default to indicate bar text retrieval problem in both cases.

		// Since above didn't break, we're in "wait" mode (more than one iteration).
		// In the following, must cast to int or any negative result will be lost due to DWORD type.
		// Note: A negative aWaitTime means we're waiting indefinitely for a match to appear.
		if (aWaitTime < 0 || (int)(aWaitTime - (GetTickCount() - start_time)) > SLEEP_INTERVAL_HALF)
			MsgSleep(aCheckInterval);
		else // Timed out.
		{
			g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Override default to indicate timeout vs. "other error".
			break;
		}
	} // for()

	// Consider this to be always successful, even if aBarHwnd == NULL
	// or the status bar didn't have the part number provided, unless the below fails.
	// Note we use a temp buf rather than writing directly to the var contents above, because
	// we don't know how long the text will be until after the above operation finishes.
	ResultType result_to_return = aOutputVar ? aOutputVar->Assign(local_buf) : OK;
	FreeInterProcMem(handle, remote_buf); // Don't free until after the above because above needs file mapping for Win9x.
	return result_to_return;
}



HWND ControlExist(HWND aParentWindow, char *aClassNameAndNum)
// This can return target_window itself for cases such as ahk_id %ControlHWND%.
{
	if (!aParentWindow)
		return NULL;
	if (!aClassNameAndNum || !*aClassNameAndNum)
		return (GetWindowLong(aParentWindow, GWL_STYLE) & WS_CHILD) ? aParentWindow : GetTopChild(aParentWindow);
		// Above: In v1.0.43.06, the parent window itself is returned if it's a child rather than its top child
		// because it seems more useful and intuitive.  This change allows ahk_id %ControlHwnd% to always operate
		// directly on the specified control's HWND rather than some sub-child.

	WindowSearch ws;
	bool is_class_name = isdigit(aClassNameAndNum[strlen(aClassNameAndNum) - 1]);

	if (is_class_name)
	{
		// Tell EnumControlFind() to search by Class+Num.  Don't call ws.SetCriteria() because
		// that has special handling for ahk_id, ahk_class, etc. in the first parameter.
		strlcpy(ws.mCriterionClass, aClassNameAndNum, sizeof(ws.mCriterionClass));
		ws.mCriterionText = "";
	}
	else // Tell EnumControlFind() to search the control's text.
	{
		*ws.mCriterionClass = '\0';
		ws.mCriterionText = aClassNameAndNum;
	}

	EnumChildWindows(aParentWindow, EnumControlFind, (LPARAM)&ws); // mFoundChild was initialized by the contructor.

	if (is_class_name && !ws.mFoundChild)
	{
		// To reduce problems with ambiguity (a class name and number of one control happens
		// to match the title/text of another control), search again only after the search
		// for the ClassNameAndNum didn't turn up anything.
		// Tell EnumControlFind() to search the control's text.
		*ws.mCriterionClass = '\0';
		ws.mCriterionText = aClassNameAndNum;
		EnumChildWindows(aParentWindow, EnumControlFind, (LPARAM)&ws); // ws.mFoundChild is already initialized to NULL due to the above check.
	}

	return ws.mFoundChild;
}



BOOL CALLBACK EnumControlFind(HWND aWnd, LPARAM lParam)
{
	WindowSearch &ws = *(WindowSearch *)lParam;  // For performance and convenience.
	if (*ws.mCriterionClass) // Caller told us to search by class name and number.
	{
		int length = GetClassName(aWnd, ws.mCandidateTitle, WINDOW_CLASS_SIZE); // Restrict the length to a small fraction of the buffer's size (also serves to leave room to append the sequence number).
		// Below: i.e. this control's title (e.g. List) in contained entirely
		// within the leading part of the user specified title (e.g. ListBox).
		// Even though this is incorrect, the appending of the sequence number
		// in the second comparison will weed out any false matches.
		// Note: since some controls end in a number (e.g. SysListView32),
		// it would not be easy to parse out the user's sequence number to
		// simplify/accelerate the search here.  So instead, use a method
		// more certain to work even though it's a little ugly.  It's also
		// necessary to do this in a way functionally identical to the below
		// so that Window Spy's sequence numbers match the ones generated here:
		// Concerning strnicmp(), see lstrcmpi note below for why a locale-insensitive match isn't done instead.
		if (length && !strnicmp(ws.mCriterionClass, ws.mCandidateTitle, length)) // Preliminary match of base class name.
		{
			// mAlreadyVisitedCount was initialized to zero by WindowSearch's constructor.  It is used
			// to accumulate how many quasi-matches on this class have been found so far.  Also,
			// comparing ws.mAlreadyVisitedCount to atoi(ws.mCriterionClass + length) would not be
			// the same as the below examples such as the following:
			// Say the ClassNN being searched for is List01 (where List0 is the class name and 1
			// is the sequence number). If a class called "List" exists in the parent window, it
			// would be found above as a preliminary match.  The below would copy "1" into the buffer,
			// which is correctly deemed not to match "01".  By contrast, the atoi() method would give
			// the wrong result because the two numbers are numerically equal.
			_itoa(++ws.mAlreadyVisitedCount, ws.mCandidateTitle, 10);  // Overwrite the buffer to contain only the count.
			// lstrcmpi() is not used: 1) avoids breaking exisitng scripts; 2) provides consistent behavior
			// across multiple locales:
			if (!stricmp(ws.mCandidateTitle, ws.mCriterionClass + length)) // The counts match too, so it's a full match.
			{
				ws.mFoundChild = aWnd; // Save this in here for return to the caller.
				return FALSE; // stop the enumeration.
			}
		}
	}
	else // Caller told us to search by the text of the control (e.g. the text printed on a button)
	{
		// Use GetWindowText() rather than GetWindowTextTimeout() because we don't want to find
		// the name accidentally in the vast amount of text present in some edit controls (e.g.
		// if the script's source code is open for editing in notepad, GetWindowText() would
		// likely find an unwanted match for just about anything).  In addition,
		// GetWindowText() is much faster.  Update: Yes, it seems better not to use
		// GetWindowTextByTitleMatchMode() in this case, since control names tend to be so
		// short (i.e. they would otherwise be very likely to be undesirably found in any large
		// edit controls the target window happens to own).  Update: Changed from strstr()
		// to strncmp() for greater selectivity.  Even with this degree of selectivity, it's
		// still possible to have ambiguous situations where a control can't be found due
		// to its title being entirely contained within that of another (e.g. a button
		// with title "Connect" would be found in the title of a button "Connect All").
		// The only way to address that would be to insist on an entire title match, but
		// that might be tedious if the title of the control is very long.  As alleviation,
		// the class name + seq. number method above can often be used instead in cases
		// of such ambiguity.  Update: Using IsTextMatch() now so that user-specified
		// TitleMatchMode will be in effect for this also.  Also, it's case sensitivity
		// helps increase selectivity, which is helpful due to how common short or ambiguous
		// control names tend to be:
		GetWindowText(aWnd, ws.mCandidateTitle, sizeof(ws.mCandidateTitle));
		if (IsTextMatch(ws.mCandidateTitle, ws.mCriterionText))
		{
			ws.mFoundChild = aWnd; // save this in here for return to the caller.
			return FALSE;
		}
	}
	// Note: The MSDN docs state that EnumChildWindows already handles the
	// recursion for us: "If a child window has created child windows of its own,
	// EnumChildWindows() enumerates those windows as well."
	return TRUE; // Keep searching.
}



int MsgBox(int aValue)
{
	char str[128];
	snprintf(str, sizeof(str), "Value = %d (0x%X)", aValue, aValue);
	return MsgBox(str);
}



int MsgBox(char *aText, UINT uType, char *aTitle, double aTimeout, HWND aOwner)
// Returns 0 if the attempt failed because of too many existing MessageBox windows,
// or if MessageBox() itself failed.
// MB_SETFOREGROUND or some similar setting appears to dismiss some types of screen savers (if active).
// However, it doesn't undo monitor low-power mode.
{
	// Set the below globals so that any WM_TIMER messages dispatched by this call to
	// MsgBox() (which may result in a recursive call back to us) know not to display
	// any more MsgBoxes:
	if (g_nMessageBoxes > MAX_MSGBOXES + 1)  // +1 for the final warning dialog.  Verified correct.
		return 0;

	if (g_nMessageBoxes == MAX_MSGBOXES)
	{
		// Do a recursive call to self so that it will be forced to the foreground.
		// But must increment this so that the recursive call allows the last MsgBox
		// to be displayed:
		++g_nMessageBoxes;
		MsgBox("The maximum number of MsgBoxes has been reached.");
		--g_nMessageBoxes;
		return 0;
	}

	// Set these in case the caller explicitly called it with a NULL, overriding the default:
	if (!aText)
		aText = "";
	if (!aTitle || !*aTitle)
		// If available, the script's filename seems a much better title in case the user has
		// more than one script running:
		aTitle = (g_script.mFileName && *g_script.mFileName) ? g_script.mFileName : NAME_PV;

	// It doesn't feel safe to modify the contents of the caller's aText and aTitle,
	// even if the caller were to tell us it is modifiable.  This is because the text
	// might be the actual contents of a variable, which we wouldn't want to truncate,
	// even temporarily, since other hotkeys can fire while this hotkey subroutine is
	// suspended, and those subroutines may refer to the contents of this (now-altered)
	// variable.  In addition, the text may reside in the clipboard's locked memory
	// area, and altering that might result in the clipboard's contents changing
	// when MsgSleep() closes the clipboard for us (after we display our dialog here).
	// Even though testing reveals that the contents aren't altered (somehow), it
	// seems best to have our own local, limited length versions here:
	// Note: 8000 chars is about the max you could ever fit on-screen at 1024x768 on some
	// XP systems, but it will hold much more before refusing to display at all (i.e.
	// MessageBox() returning failure), perhaps about 150K:
	char text[MSGBOX_TEXT_SIZE];
	char title[DIALOG_TITLE_SIZE];
	strlcpy(text, aText, sizeof(text));
	strlcpy(title, aTitle, sizeof(title));

	uType |= MB_SETFOREGROUND;  // Always do these so that caller doesn't have to specify.

	// In the below, make the MsgBox owned by the topmost window rather than our main
	// window, in case there's another modal dialog already displayed.  The forces the
	// user to deal with the modal dialogs starting with the most recent one, which
	// is what we want.  Otherwise, if a middle dialog was dismissed, it probably
	// won't be able to return which button was pressed to its original caller.
	// UPDATE: It looks like these modal dialogs can't own other modal dialogs,
	// so disabling this:
	/*
	HWND topmost = GetTopWindow(g_hWnd);
	if (!topmost) // It has no child windows.
		topmost = g_hWnd;
	*/

	// Unhide the main window, but have it minimized.  This creates a task
	// bar button so that it's easier the user to remember that a dialog
	// is open and waiting (there are probably better ways to handle
	// this whole thing).  UPDATE: This isn't done because it seems
	// best not to have the main window be inaccessible until the
	// dialogs are dismissed (in case ever want to use it to display
	// status info, etc).  It seems that MessageBoxes get their own
	// task bar button when they're not AppModal, which is one of the
	// main things I wanted, so that's good too):
//	if (!IsWindowVisible(g_hWnd) || !IsIconic(g_hWnd))
//		ShowWindowAsync(g_hWnd, SW_SHOWMINIMIZED);

	/*
	If the script contains a line such as "#y::MsgBox, test", and a hotkey is used
	to activate Windows Explorer and another hotkey is then used to invoke a MsgBox,
	that MsgBox will be psuedo-minimized or invisible, even though it does have the
	input focus.  This attempt to fix it didn't work, so something is probably checking
	the physical key state of LWIN/RWIN and seeing that they're down:
	modLR_type modLR_now = GetModifierLRState();
	modLR_type win_keys_down = modLR_now & (MOD_LWIN | MOD_RWIN);
	if (win_keys_down)
		SetModifierLRStateSpecific(win_keys_down, modLR_now, KEYUP);
	*/

	// Note: Even though when multiple messageboxes exist, they might be
	// destroyed via a direct call to their WindowProc from our message pump's
	// DispatchMessage, or that of another MessageBox's message pump, it
	// appears that MessageBox() is designed to be called recursively like
	// this, since it always returns the proper result for the button on the
	// actual MessageBox it originally invoked.  In other words, if a bunch
	// of Messageboxes are displayed, and this user dismisses an older
	// one prior to dealing with a newer one, all the MessageBox()
	// return values will still wind up being correct anyway, at least
	// on XP.  The only downside to the way this is designed is that
	// the keyboard can't be used to navigate the buttons on older
	// messageboxes (only the most recent one).  This is probably because
	// the message pump of MessageBox() isn't designed to properly dispatch
	// keyboard messages to other MessageBox window instances.  I tried
	// to fix that by making our main message pump handle all messages
	// for all dialogs, but that turns out to be pretty complicated, so
	// I abandoned it for now.

	// Note: It appears that MessageBox windows, and perhaps all modal dialogs in general,
	// cannot own other windows.  That's too bad because it would have allowed each new
	// MsgBox window to be owned by any previously existing one, so that the user would
	// be forced to close them in order if they were APPL_MODAL.  But it's not too big
	// an issue since the only disadvantage is that the keyboard can't be use to
	// to navigate in MessageBoxes other than the most recent.  And it's actually better
	// the way it is now in the sense that the user can dismiss the messageboxes out of
	// order, which might (in rare cases) be desirable.

	if (aTimeout > 2147483) // This is approximately the max number of seconds that SetTimer can handle.
		aTimeout = 2147483;
	if (aTimeout < 0) // But it can be equal to zero to indicate no timeout at all.
		aTimeout = 0.1;  // A value that might cue the user that something is wrong.
	// For the above:
	// MsgBox's smart comma handling will usually prevent negatives due to the fact that it considers
	// a negative to be part of the text param.  But if it does happen, timeout after a short time,
	// which may signal the user that the script passed a bad parameter.

	// v1.0.33: The following is a workaround for the fact that an MsgBox with only an OK button
	// doesn't obey EndDialog()'s parameter:
	g->DialogHWND = NULL;
	g->MsgBoxTimedOut = false;

	// At this point, we know a dialog will be displayed.  See macro's comments for details:
	DIALOG_PREP // Must be done prior to POST_AHK_DIALOG() below.
	POST_AHK_DIALOG((DWORD)(aTimeout * 1000))

	++g_nMessageBoxes;  // This value will also be used as the Timer ID if there's a timeout.
	g->MsgBoxResult = MessageBox(aOwner, text, title, uType);
	--g_nMessageBoxes;
	// Above's use of aOwner: MsgBox, FileSelectFile, and other dialogs seem to consider aOwner to be NULL
	// when aOwner is minimized or hidden.

	DIALOG_END

	// If there's a timer, kill it for performance reasons since it's no longer needed.
	// Actually, this isn't easy to do because we don't know what the HWND of the MsgBox
	// window was, so don't bother:
	//if (aTimeout != 0.0)
	//	KillTimer(...);

//	if (!g_nMessageBoxes)
//		ShowWindowAsync(g_hWnd, SW_HIDE);  // Hide the main window if it no longer has any child windows.
//	else

	// This is done so that the next message box of ours will be brought to the foreground,
	// to remind the user that they're still out there waiting, and for convenience.
	// Update: It seems bad to do this in cases where the user intentionally wants the older
	// messageboxes left in the background, to deal with them later.  So, in those cases,
	// we might be doing more harm than good because the user's foreground window would
	// be intrusively changed by this:
	//WinActivateOurTopDialog();

	// The following comment is apparently not always true -- sometimes the AHK_TIMEOUT from
	// EndDialog() is received correctly.  But I haven't discovered the circumstances of how
	// and why the behavior varies:
	// Unfortunately, it appears that MessageBox() will return zero rather
	// than AHK_TIMEOUT that was specified in EndDialog() at least under WinXP.
	if (g->MsgBoxTimedOut || (!g->MsgBoxResult && aTimeout > 0)) // v1.0.33: Added g->MsgBoxTimedOut, see comment higher above.
		// Assume it timed out rather than failed, since failure should be VERY rare.
		g->MsgBoxResult = AHK_TIMEOUT;
	// else let the caller handle the display of the error message because only it knows
	// whether to also tell the user something like "the script will not continue".
	return g->MsgBoxResult;
}



HWND FindOurTopDialog()
// Returns the HWND of our topmost MsgBox or FileOpen dialog (and perhaps other types of modal
// dialogs if they are of class #32770) even if it wasn't successfully brought to
// the foreground here.
// Using Enum() seems to be the only easy way to do this, since these modal MessageBoxes are
// *owned*, not children of the main window.  There doesn't appear to be any easier way to
// find out which windows another window owns.  GetTopWindow(), GetActiveWindow(), and GetWindow()
// do not work for this purpose.  And using FindWindow() discouraged because it can hang
// in certain circumtances (Enum is probably just as fast anyway).
{
	// The return value of EnumWindows() is probably a raw indicator of success or failure,
	// not whether the Enum found something or continued all the way through all windows.
	// So don't bother using it.
	pid_and_hwnd_type pid_and_hwnd;
	pid_and_hwnd.pid = GetCurrentProcessId();
	pid_and_hwnd.hwnd = NULL;  // Init.  Called function will set this for us if it finds a match.
	EnumWindows(EnumDialog, (LPARAM)&pid_and_hwnd);
	return pid_and_hwnd.hwnd;
}



BOOL CALLBACK EnumDialog(HWND aWnd, LPARAM lParam)
// lParam should be a pointer to a ProcessId (ProcessIds are always non-zero?)
// To continue enumeration, the function must return TRUE; to stop enumeration, it must return FALSE. 
{
	pid_and_hwnd_type &pah = *(pid_and_hwnd_type *)lParam;  // For performance and convenience.
	if (!lParam || !pah.pid) return FALSE;
	DWORD pid;
	GetWindowThreadProcessId(aWnd, &pid);
	if (pid == pah.pid)
	{
		char buf[32];
		GetClassName(aWnd, buf, sizeof(buf));
		// This is the class name for windows created via MessageBox(), GetOpenFileName(), and probably
		// other things that use modal dialogs:
		if (!strcmp(buf, "#32770"))
		{
			pah.hwnd = aWnd;  // An output value for the caller.
			return FALSE;  // We're done.
		}
	}
	return TRUE;  // Keep searching.
}



struct owning_struct {HWND owner_hwnd; HWND first_child;};
HWND WindowOwnsOthers(HWND aWnd)
// Only finds owned windows if they are visible, by design.
{
	owning_struct own = {aWnd, NULL};
	EnumWindows(EnumParentFindOwned, (LPARAM)&own);
	return own.first_child;
}



BOOL CALLBACK EnumParentFindOwned(HWND aWnd, LPARAM lParam)
{
	HWND owner_hwnd = GetWindow(aWnd, GW_OWNER);
	// Note: Many windows seem to own other invisible windows that have blank titles.
	// In our case, require that it be visible because we don't want to return an invisible
	// window to the caller because such windows aren't designed to be activated:
	if (owner_hwnd && owner_hwnd == ((owning_struct *)lParam)->owner_hwnd && IsWindowVisible(aWnd))
	{
		((owning_struct *)lParam)->first_child = aWnd;
		return FALSE; // Match found, we're done.
	}
	return TRUE;  // Continue enumerating.
}



HWND GetNonChildParent(HWND aWnd)
// Returns the first ancestor of aWnd that isn't itself a child.  aWnd itself is returned if
// it is not a child.  Returns NULL only if aWnd is NULL.  Also, it should always succeed
// based on the axiom that any window with the WS_CHILD style (aka WS_CHILDWINDOW) must have
// a non-child ancestor somewhere up the line.
// This function doesn't do anything special with owned vs. unowned windows.  Despite what MSDN
// says, GetParent() does not return the owner window, at least in some cases on Windows XP
// (e.g. BulletProof FTP Server). It returns NULL instead. In any case, it seems best not to
// worry about owner windows for this function's caller (MouseGetPos()), since it might be
// desirable for that command to return the owner window even though it can't actually be
// activated.  This is because attempts to activate an owner window should automatically cause
// the OS to activate the topmost owned window instead.  In addition, the owner window may
// contain the actual title or text that the user is interested in.  UPDATE: Due to the fact
// that this function retrieves the first parent that's not a child window, it's likely that
// that window isn't its owner anyway (since the owner problem usually applies to a parent
// window being owned by some controlling window behind it).
{
	if (!aWnd) return aWnd;
	HWND parent, parent_prev;
	for (parent_prev = aWnd; ; parent_prev = parent)
	{
		if (!(GetWindowLong(parent_prev, GWL_STYLE) & WS_CHILD))  // Found the first non-child parent, so return it.
			return parent_prev;
		// Because Windows 95 doesn't support GetAncestor(), we'll use GetParent() instead:
		if (   !(parent = GetParent(parent_prev))   )
			return parent_prev;  // This will return aWnd if aWnd has no parents.
	}
}



HWND GetTopChild(HWND aParent)
{
	if (!aParent) return aParent;
	HWND hwnd_top, next_top;
	// Get the topmost window of the topmost window of...
	// i.e. Since child windows can also have children, we keep going until
	// we reach the "last topmost" window:
	for (hwnd_top = GetTopWindow(aParent)
		; hwnd_top && (next_top = GetTopWindow(hwnd_top))
		; hwnd_top = next_top);

	//if (!hwnd_top)
	//{
	//	MsgBox("no top");
	//	return FAIL;
	//}
	//else
	//{
	//	//if (GetTopWindow(hwnd_top))
	//	//	hwnd_top = GetTopWindow(hwnd_top);
	//	char class_name[64];
	//	GetClassName(next_top, class_name, sizeof(class_name));
	//	MsgBox(class_name);
	//}

	return hwnd_top ? hwnd_top : aParent;  // Caller relies on us never returning NULL if aParent is non-NULL.
}



bool IsWindowHung(HWND aWnd)
{
	if (!aWnd) return false;

	// OLD, SLOWER METHOD:
	// Don't want to use a long delay because then our messages wouldn't get processed
	// in a timely fashion.  But I'm not entirely sure if the 10ms delay used below
	// is even used by the function in this case?  Also, the docs aren't clear on whether
	// the function returns success or failure if the window is hung (probably failure).
	// If failure, perhaps you have to call GetLastError() to determine whether it failed
	// due to being hung or some other reason?  Does the output param dwResult have any
	// useful info in this case?  I expect what will happen is that in most cases, the OS
	// will already know that the window is hung.  However, if the window just became hung
	// in the last 5 seconds, I think it may take the remainder of the 5 seconds for the OS
	// to notice it.  However, allowing it the option of sleeping up to 5 seconds seems
	// really bad, since keyboard and mouse input would probably be frozen (actually it
	// would just be really laggy because the OS would bypass the hook during that time).
	// So some compromise value seems in order.  500ms seems about right.  UPDATE: Some
	// windows might need longer than 500ms because their threads are engaged in
	// heavy operations.  Since this method is only used as a fallback method now,
	// it seems best to give them the full 5000ms default, which is what (all?) Windows
	// OSes use as a cutoff to determine whether a window is "not responding":
	DWORD dwResult;
	#define Slow_IsWindowHung !SendMessageTimeout(aWnd, WM_NULL, 0, 0, SMTO_ABORTIFHUNG, 5000, &dwResult)

	// NEW, FASTER METHOD:
	// This newer method's worst-case performance is at least 30x faster than the worst-case
	// performance of the old method that  uses SendMessageTimeout().
	// And an even worse case can be envisioned which makes the use of this method
	// even more compelling: If the OS considers a window NOT to be hung, but the
	// window's message pump is sluggish about responding to the SendMessageTimeout() (perhaps
	// taking 2000ms or more to respond due to heavy disk I/O or other activity), the old method
	// will take several seconds to return, causing mouse and keyboard lag if our hook(s)
	// are installed; not to mention making our app's windows, tray menu, and other GUI controls
	// unresponsive during that time).  But I believe in this case the new method will return
	// instantly, since the OS has been keeping track in the background, and can tell us
	// immediately that the window isn't hung.
	// Here are some seemingly contradictory statements uttered by MSDN.  Perhaps they're
	// not contradictory if the first sentence really means "by a different thread of the same
	// process":
	// "If the specified window was created by a different thread, the system switches to that
	// thread and calls the appropriate window procedure.  Messages sent between threads are
	// processed only when the receiving thread executes message retrieval code. The sending
	// thread is blocked until the receiving thread processes the message."
	if (g_os.IsWin9x())
	{
		typedef BOOL (WINAPI *MyIsHungThread)(DWORD);
		static MyIsHungThread IsHungThread = (MyIsHungThread)GetProcAddress(GetModuleHandle("user32")
			, "IsHungThread");
		// When function not available, fall back to the old method:
		return IsHungThread ? IsHungThread(GetWindowThreadProcessId(aWnd, NULL)) : Slow_IsWindowHung;
	}

	// Otherwise: NT/2k/XP/2003 or later, so try to use the newer method.
	// The use of IsHungAppWindow() (supported under Win2k+) is discouraged by MS,
	// but it's useful to prevent the script from getting hung when it tries to do something
	// to a hung window.
	typedef BOOL (WINAPI *MyIsHungAppWindow)(HWND);
	static MyIsHungAppWindow IsHungAppWindow = (MyIsHungAppWindow)GetProcAddress(GetModuleHandle("user32")
		, "IsHungAppWindow");
	return IsHungAppWindow ? IsHungAppWindow(aWnd) : Slow_IsWindowHung;
}



int GetWindowTextTimeout(HWND aWnd, char *aBuf, int aBufSize, UINT aTimeout)
// This function must be kept thread-safe because it may be called (indirectly) by hook thread too.
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Returns the length of what would be copied (not including the zero terminator).
// In addition, if aBuf is not NULL, the window text is copied into aBuf (not to exceed aBufSize).
// AutoIt3 author indicates that using WM_GETTEXT vs. GetWindowText() sometimes yields more text.
// Perhaps this is because GetWindowText() has built-in protection against hung windows and
// thus isn't actually sending WM_GETTEXT.  The method here is hopefully the best of both worlds
// (protection against hung windows causing our thread to hang, and getting more text).
// Another tidbit from MSDN about SendMessage() that might be of use here sometime:
// "However, the sending thread will process incoming nonqueued (those sent directly to a window
// procedure) messages while waiting for its message to be processed. To prevent this, use
// SendMessageTimeout with SMTO_BLOCK set."  Currently not using SMTO_BLOCK because it
// doesn't seem necessary.
// Update: GetWindowText() is so much faster than SendMessage() and SendMessageTimeout(), at
// least on XP, so GetWindowTextTimeout() should probably only be used when getting the max amount
// of text is important (e.g. this function can fetch the text in a RichEdit20A control and
// other edit controls, whereas GetWindowText() doesn't).  This function is used to implement
// things like WinGetText and ControlGetText, in which getting the maximum amount and types
// of text is more important than performance.
{
	if (!aWnd || (aBuf && aBufSize < 1)) // No HWND or no room left in buffer (some callers rely on this check).
		return 0; // v1.0.40.04: Fixed to return 0 rather than setting aBuf to NULL and continuing (callers don't want that).

	// Override for Win95 because AutoIt3 author says it might crash otherwise:
	if (aBufSize > WINDOW_TEXT_SIZE && g_os.IsWin95())
		aBufSize = WINDOW_TEXT_SIZE;

	LRESULT result, length;
	if (aBuf)
	{
		*aBuf = '\0';  // Init just to get it out of the way in case of early return/error.
		if (aBufSize == 1) // Room only for the terminator, so go no further (some callers rely on this check).
			return 0;

		// Below demonstrated that GetWindowText() is dramatically faster than either SendMessage()
		// or SendMessageTimeout() (noticeably faster when you have hotkeys that activate
		// windows, or toggle between two windows):
		//return GetWindowText(aWnd, aBuf, aBufSize);
		//return (int)SendMessage(aWnd, WM_GETTEXT, (WPARAM)aBufSize, (LPARAM)aBuf);

		// Don't bother calling IsWindowHung() because the below call will return
		// nearly instantly if the OS already "knows" that the target window has
		// be unresponsive for 5 seconds or so (i.e. it keeps track of such things
		// on an ongoing basis, at least XP seems to).
		result = SendMessageTimeout(aWnd, WM_GETTEXT, (WPARAM)aBufSize, (LPARAM)aBuf
			, SMTO_ABORTIFHUNG, aTimeout, (LPDWORD)&length);
		if (length >= aBufSize) // Happens sometimes (at least ==aBufSize) for apps that wrongly include the terminator in the reported length.
			length = aBufSize - 1; // Override.

		// v1.0.40.04: The following check was added because when the text is too large to to fit in the
		// buffer, the OS (or at least certain applications such as AIM) return a length that *includes*
		// the zero terminator, violating the documented behavior of WM_GETTEXT.  In case the returned
		// length is too long by 1 (or even more than 1), calculate the length explicitly by checking if
		// there's another terminator to the left of the indicated length.  The following loop
		// is used in lieu of strlen() for performance reasons (because sometimes the text is huge).
		// It assumes that there will be no more than one additional terminator to the left of the
		// indicated length, which so far seems to be true:
		for (char *cp = aBuf + length; cp >= aBuf; --cp)
		{
			if (!*cp)
			{
				// Keep going to the left until the last consecutive terminator is found.
				// Necessary for AIM when compiled in release mode (but not in debug mode
				// for some reason!):
				for (; cp > aBuf && !cp[-1]; --cp); // Self-contained loop.  Verified correct.
				length = cp - aBuf;
				break;
			}
		}
		// If the above loop didn't "break", a terminator wasn't found.
		// Terminate explicitly because MSDN docs aren't clear that it will always be terminated automatically.
		// Update: This also protects against misbehaving apps that might handle the WM_GETTEXT message
		// rather than passing it to DefWindowProc() but that don't terminate the buffer.  This has been
		// confirmed to be necessary at least for AIM when aBufSize==1 (although 1 is no longer possible due
		// to a check that has been added further above):
		aBuf[length] = '\0';
	}
	else
	{
		result = SendMessageTimeout(aWnd, WM_GETTEXTLENGTH, 0, 0, SMTO_ABORTIFHUNG, aTimeout, (LPDWORD)&length);
		// The following can be temporarily uncommented out to demonstrate how some apps such as AIM's
		// write-an-instant-message window have some controls that respond to WM_GETTEXTLENGTH with a
		// length that's completely different than the length with which they respond to WM_GETTEXT.
		// Here are some of the discrepancies:
		// WM_GETTEXTLENGTH vs. WM_GETTEXT:
		// 92 vs. 318 (bigger)
		// 50 vs. 159 (bigger)
		// 3 vs. 0 (smaller)
		// 24 vs. 88 (etc.)
		// 80 vs. 188
		// 24 vs. 88
		// 80 vs. 188
		//char buf[32000];
		//LRESULT length2;
		//result = SendMessageTimeout(aWnd, WM_GETTEXT, (WPARAM)sizeof(buf), (LPARAM)buf
		//	, SMTO_ABORTIFHUNG, aTimeout, (LPDWORD)&length2);
		//if (length2 != length)
		//{
		//	int x = 0;  // Put breakpoint here.
		//}
		// An attempt to fix the size estimate to be larger for misbehaving apps like AIM, but it's ineffective
		// so commented out:
		//if (!length)
		//	length = GetWindowTextLength(aWnd);
	}

	// "length" contains the length of what was (or would have been) copied, not including the terminator:
	return result ? (int)length : 0;  // "result" is zero upon failure or timeout.
}



///////////////////////////////////////////////////////////////////////////



ResultType WindowSearch::SetCriteria(global_struct &aSettings, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
// Returns FAIL if the new criteria can't possibly match a window (due to ahk_id being in invalid
// window or the specfied ahk_group not existing).  Otherwise, it returns OK.
// Callers must ensure that aText, aExcludeTitle, and aExcludeText point to buffers whose contents
// will be available for the entire duration of the search.  In other words, the caller should not
// call MsgSleep() in a way that would allow another thread to launch and overwrite the contents
// of the sDeref buffer (which might contain the contents).  Things like mFoundHWND and mFoundCount
// are not initialized here because sometimes the caller changes the criteria without wanting to
// reset the search.
// This function must be kept thread-safe because it may be called (indirectly) by hook thread too.
{
	// Set any criteria which are not context sensitive.  It doesn't seem necessary to make copies of
	// mCriterionText, mCriterionExcludeTitle, and mCriterionExcludeText because they're never altered
	// here, nor does there seem to be a risk that deref buffer's contents will get overwritten
	// while this set of criteria is in effect because our callers never allow interrupting script-threads
	// *during* the duration of any one set of criteria.
	bool exclude_title_became_non_blank = *aExcludeTitle && !*mCriterionExcludeTitle;
	mCriterionExcludeTitle = aExcludeTitle;
	mCriterionExcludeTitleLength = strlen(mCriterionExcludeTitle); // Pre-calculated for performance.
	mCriterionText = aText;
	mCriterionExcludeText = aExcludeText;
	mSettings = &aSettings;

	DWORD orig_criteria = mCriteria;
	char *ahk_flag, *cp, buf[MAX_VAR_NAME_LENGTH + 1];
	int criteria_count;
	size_t size;

	for (mCriteria = 0, ahk_flag = aTitle, criteria_count = 0;; ++criteria_count, ahk_flag += 4) // +4 only since an "ahk_" string that isn't qualified may have been found.
	{
		if (   !(ahk_flag = strcasestr(ahk_flag, "ahk_"))   ) // No other special strings are present.
		{
			if (!criteria_count) // Since no special "ahk_" criteria were present, it is CRITERION_TITLE by default.
			{
				mCriteria = CRITERION_TITLE; // In this case, there is only one criterion.
				strlcpy(mCriterionTitle, aTitle, sizeof(mCriterionTitle));
				mCriterionTitleLength = strlen(mCriterionTitle); // Pre-calculated for performance.
			}
			break;
		}
		// Since above didn't break, another instance of "ahk_" has been found. To reduce ambiguity,
		// the following requires that any "ahk_" criteria beyond the first be preceded by at least
		// one space or tab:
		if (criteria_count && !IS_SPACE_OR_TAB(ahk_flag[-1])) // Relies on short-circuit boolean order.
		{
			--criteria_count; // Decrement criteria_count to compensate for the loop's increment.
			continue;
		}
		// Since above didn't "continue", it meets the basic test.  But is it an exact match for one of the
		// special criteria strings?  If not, it's really part of the title criterion instead.
		cp = ahk_flag + 4;
		if (!strnicmp(cp, "id", 2))
		{
			cp += 2;
			mCriteria |= CRITERION_ID;
			mCriterionHwnd = (HWND)ATOU64(cp);
			// Note that this can validly be the HWND of a child window; i.e. ahk_id %ChildWindowHwnd% is supported.
			if (mCriterionHwnd != HWND_BROADCAST && !IsWindow(mCriterionHwnd)) // Checked here once rather than each call to IsMatch().
			{
				mCriterionHwnd = NULL;
				return FAIL; // Inform caller of invalid criteria.  No need to do anything else further below.
			}
		}
		else if (!strnicmp(cp, "pid", 3))
		{
			cp += 3;
			mCriteria |= CRITERION_PID;
			mCriterionPID = ATOU(cp);
		}
		else if (!strnicmp(cp, "class", 5))
		{
			cp += 5;
			mCriteria |= CRITERION_CLASS;
			// In the following line, it may have been preferable to skip only zero or one spaces rather than
			// calling omit_leading_whitespace().  But now this should probably be kept for backward compatibility.
			// Besides, even if it's possible for a class name to start with a space, a RegEx dot or other symbol
			// can be used to match it via SetTitleMatchMode RegEx.
			strlcpy(mCriterionClass, omit_leading_whitespace(cp), sizeof(mCriterionClass)); // Copy all of the remaining string to simplify the below.
			for (cp = mCriterionClass; cp = strcasestr(cp, "ahk_"); cp += 4)  // Fix for v1.0.47.06: strstr() changed to strcasestr() for consistency with the other sections.
			{
				// This loop truncates any other criteria from the class criteria.  It's not a complete
				// solution because it doesn't validate that what comes after the "ahk_" string is a
				// valid criterion name. But for it not to be and yet also be part of some valid class
				// name seems far too unlikely to worry about.  It would have to be a legitimate class name
				// such as "ahk_class SomeClassName ahk_wrong".
				if (cp == mCriterionClass) // This check prevents underflow in the next check.
				{
					*cp = '\0';
					break;
				}
				else
					if (IS_SPACE_OR_TAB(cp[-1]))
					{
						cp[-1] = '\0';
						break;
					}
					//else assume this "ahk_" string is part of the literal text, continue looping in case
					// there is a legitimate "ahk_" string after this one.
			} // for()
		}
		else if (!strnicmp(cp, "group", 5))
		{
			cp += 5;
			mCriteria |= CRITERION_GROUP;
			strlcpy(buf, omit_leading_whitespace(cp), sizeof(buf));
			if (cp = StrChrAny(buf, " \t")) // Group names can't contain spaces, so terminate at the first one to exclude any "ahk_" criteria that come afterward.
				*cp = '\0';
			if (   !(mCriterionGroup = g_script.FindGroup(buf))   )
				return FAIL; // No such group: Inform caller of invalid criteria.  No need to do anything else further below.
		}
		else // It doesn't qualify as a special criteria name even though it starts with "ahk_".
		{
			--criteria_count; // Decrement criteria_count to compensate for the loop's increment.
			continue;
		}
		// Since above didn't return or continue, a valid "ahk_" criterion has been discovered.
		// If this is the first such criterion, any text that lies to its left should be interpreted
		// as CRITERION_TITLE.  However, for backward compatibility it seems best to disqualify any title
		// consisting entirely of whitespace.  This is because some scripts might have a variable containing
		// whitespace followed by the string ahk_class, etc. (however, any such whitespace is included as a
		// literal part of the title criterion for flexibilty and backward compatibility).
		if (!criteria_count && ahk_flag > omit_leading_whitespace(aTitle))
		{
			mCriteria |= CRITERION_TITLE;
			// Omit exactly one space or tab from the title criterion. That space or tab is the one
			// required to delimit the special "ahk_" string.  Any other spaces or tabs to the left of
			// that one are considered literal (for flexibility):
			size = ahk_flag - aTitle; // This will always be greater than one due to other checks above, which will result in at least one non-whitespace character in the title criterion.
			if (size > sizeof(mCriterionTitle)) // Prevent overflow.
				size = sizeof(mCriterionTitle);
			strlcpy(mCriterionTitle, aTitle, size); // Copy only the eligible substring as the criteria.
			mCriterionTitleLength = strlen(mCriterionTitle); // Pre-calculated for performance.
		}
	}

	// Since this function doesn't change mCandidateParent, there is no need to update the candidate's
	// attributes unless the type of criterion has changed or if mExcludeTitle became non-blank as
	// a result of our action above:
	if (mCriteria != orig_criteria || exclude_title_became_non_blank)
		UpdateCandidateAttributes(); // In case mCandidateParent isn't NULL, fetch different attributes based on what was set above.
	//else for performance reasons, avoid unnecessary updates.
	return OK;
}



void WindowSearch::UpdateCandidateAttributes()
// This function must be kept thread-safe because it may be called (indirectly) by hook thread too.
{
	// Nothing to do until SetCandidate() is called with a non-NULL candidate and SetCriteria()
	// has been called for the first time (otherwise, mCriterionExcludeTitle and other things
	// are not yet initialized:
	if (!mCandidateParent || !mCriteria)
		return;
	if ((mCriteria & CRITERION_TITLE) || *mCriterionExcludeTitle) // Need the window's title in both these cases.
		if (!GetWindowText(mCandidateParent, mCandidateTitle, sizeof(mCandidateTitle)))
			*mCandidateTitle = '\0'; // Failure or blank title is okay.
	if (mCriteria & CRITERION_PID) // In which case mCriterionPID should already be filled in, though it might be an explicitly specified zero.
		GetWindowThreadProcessId(mCandidateParent, &mCandidatePID);
	if (mCriteria & CRITERION_CLASS)
		GetClassName(mCandidateParent, mCandidateClass, sizeof(mCandidateClass)); // Limit to WINDOW_CLASS_SIZE in this case since that's the maximum that can be searched.
	// Nothing to do for these:
	//CRITERION_GROUP:    Can't be pre-processed at this stage.
	//CRITERION_ID:       It is mCandidateParent, which has already been set by SetCandidate().
}



HWND WindowSearch::IsMatch(bool aInvert)
// Caller must have called SetCriteria prior to calling this method, at least for the purpose of setting
// mSettings to a valid address (and possibly other reasons).
// This method returns the HWND of mCandidateParent if it matches the previously specified criteria
// (title/pid/id/class/group) or NULL otherwise.  Upon NULL, it doesn't reset mFoundParent or mFoundCount
// in case previous match(es) were found when mFindLastMatch is in effect.
// Thread-safety: With the following exception, this function must be kept thread-safe because it may be
// called (indirectly) by hook thread too: The hook thread must never call here directly or indirectly with
// mArrayStart!=NULL because the corresponding section below is probably not thread-safe.
{
	if (!mCandidateParent || !mCriteria) // Nothing to check, so no match.
		return NULL;

	if ((mCriteria & CRITERION_TITLE) && *mCriterionTitle) // For performance, avoid the calls below (especially RegEx) when mCriterionTitle is blank (assuming it's even possible for it to be blank under these conditions).
	{
		switch(mSettings->TitleMatchMode)
		{
		case FIND_ANYWHERE:
			if (!strstr(mCandidateTitle, mCriterionTitle)) // Suitable even if mCriterionTitle is blank, though that's already ruled out above.
				return NULL;
			break;
		case FIND_IN_LEADING_PART:
			if (strncmp(mCandidateTitle, mCriterionTitle, mCriterionTitleLength)) // Suitable even if mCriterionTitle is blank, though that's already ruled out above. If it were possible, mCriterionTitleLength would be 0 and thus strncmp would yield 0 to indicate "strings are equal".
				return NULL;
			break;
		case FIND_REGEX:
			if (!RegExMatch(mCandidateTitle, mCriterionTitle))
				return NULL;
			break;
		default: // Exact match.
			if (strcmp(mCandidateTitle, mCriterionTitle))
				return NULL;
		}
		// If above didn't return, it's a match so far so continue onward to the other checks.
	}

	if (mCriteria & CRITERION_CLASS) // mCriterionClass is probably always non-blank when CRITERION_CLASS is present (harmless even if it isn't), so *mCriterionClass isn't checked.
	{
		if (mSettings->TitleMatchMode == FIND_REGEX)
		{
			if (!RegExMatch(mCandidateClass, mCriterionClass))
				return NULL;
		}
		else // For backward compatibility, all other modes use exact-match for Class.
			if (strcmp(mCandidateClass, mCriterionClass)) // Doesn't match the required class name.
				return NULL;
		// If nothing above returned, it's a match so far so continue onward to the other checks.
	}

	// For the following, mCriterionPID would already be filled in, though it might be an explicitly specified zero.
	if ((mCriteria & CRITERION_PID) && mCandidatePID != mCriterionPID) // Doesn't match required PID.
		return NULL;
	//else it's a match so far, but continue onward in case there are other criteria.

	// The following also handles the fact that mCriterionGroup might be NULL if the specified group
	// does not exist or was never successfully created:
	if ((mCriteria & CRITERION_GROUP) && (!mCriterionGroup || !mCriterionGroup->IsMember(mCandidateParent, *mSettings)))
		return NULL; // Isn't a member of specified group.
	//else it's a match so far, but continue onward in case there are other criteria (a little strange in this case, but might be useful).

	// CRITERION_ID is listed last since in terms of actual calling frequency, this part is hardly ever
	// executed: It's only ever called this way from WinActive(), and possibly indirectly by an ahk_group
	// that contains an ahk_id specification.  It's also called by WinGetList()'s EnumWindows(), though
	// extremely rarely. It's also called this way from other places to determine whether an ahk_id window
	// matches the other criteria such as WinText, ExcludeTitle, and mAlreadyVisited.
	// mCriterionHwnd should already be filled in, though it might be an explicitly specified zero.
	// Note: IsWindow(mCriterionHwnd) was already called by SetCriteria().
	if ((mCriteria & CRITERION_ID) && mCandidateParent != mCriterionHwnd) // Doesn't match the required HWND.
		return NULL;
	//else it's a match so far, but continue onward in case there are other criteria.

	// The above would have returned if the candidate window isn't a match for what was specified by
	// the script's WinTitle parameter.  So now check that the ExcludeTitle criterion is satisfied.
	// This is done prior to checking WinText/ExcludeText for performance reasons:

	if (*mCriterionExcludeTitle)
	{
		switch(mSettings->TitleMatchMode)
		{
		case FIND_ANYWHERE:
			if (strstr(mCandidateTitle, mCriterionExcludeTitle))
				return NULL;
			break;
		case FIND_IN_LEADING_PART:
			if (!strncmp(mCandidateTitle, mCriterionExcludeTitle, mCriterionExcludeTitleLength))
				return NULL;
			break;
		case FIND_REGEX:
			if (RegExMatch(mCandidateTitle, mCriterionExcludeTitle))
				return NULL;
			break;
		default: // Exact match.
			if (!strcmp(mCandidateTitle, mCriterionExcludeTitle))
				return NULL;
		}
		// If above didn't return, WinTitle and ExcludeTitle are both satisified.  So continue
		// on below in case there is some WinText or ExcludeText to search.
	}

	if (!aInvert) // If caller specified aInvert==true, it will do the below instead of us.
		for (int i = 0; i < mAlreadyVisitedCount; ++i)
			if (mCandidateParent == mAlreadyVisited[i])
				return NULL;

	if (*mCriterionText || *mCriterionExcludeText) // It's not quite a match yet since there are more criteria.
	{
		// Check the child windows for the specified criteria.
		// EnumChildWindows() will return FALSE (failure) in at least two common conditions:
		// 1) It's EnumChildProc callback returned false (i.e. it ended the enumeration prematurely).
		// 2) The specified parent has no children.
		// Since in both these cases GetLastError() returns ERROR_SUCCESS, we discard the return
		// value and just check mFoundChild to determine whether a match has been found:
		mFoundChild = NULL;  // Init prior to each call, in case mFindLastMatch is true.
		EnumChildWindows(mCandidateParent, EnumChildFind, (LPARAM)this);
		if (!mFoundChild) // This parent has no matching child, or no children at all.
			return NULL;
	}

	// Since the above didn't return or none of the checks above were needed, it's a complete match.
	// If mFindLastMatch is true, this new value for mFoundParent will stay in effect until
	// overridden by another matching window later:
	if (!aInvert)
	{
		mFoundParent = mCandidateParent;
		++mFoundCount; // This must be done prior to the mArrayStart section below.
	}
	//else aInvert==true, which means caller doesn't want the above set.

	if (mArrayStart) // Probably not thread-safe due to FindOrAddVar(), so hook thread must call only with NULL mArrayStart.
	{
		// Make it longer than Max var name so that FindOrAddVar() will be able to spot and report
		// var names that are too long:
		char var_name[MAX_VAR_NAME_LENGTH + 20];
		// To help performance (in case the linked list of variables is huge), tell it where
		// to start the search.  Use the base array name rather than the preceding element because,
		// for example, Array19 is alphabetially less than Array2, so we can't rely on the
		// numerical ordering:
		Var *array_item = g_script.FindOrAddVar(var_name
			, snprintf(var_name, sizeof(var_name), "%s%u", mArrayStart->mName, mFoundCount)
			, mArrayStart->IsLocal() ? ALWAYS_USE_LOCAL : ALWAYS_USE_GLOBAL);
		if (array_item)
			array_item->AssignHWND(mFoundParent);
		//else no error reporting currently, since should be very rare.
	}

	// Fix for v1.0.30.01: Don't return mFoundParent because its NULL when aInvert is true.
	// At this stage, the candidate is a known match, so return it:
	return mCandidateParent;
}



///////////////////////////////////////////////////////////////////



void SetForegroundLockTimeout()
{
	// Even though they may not help in all OSs and situations, this lends peace-of-mind.
	// (it doesn't appear to help on my XP?)
	if (g_os.IsWin98orLater() || g_os.IsWin2000orLater())
	{
		// Don't check for failure since this operation isn't critical, and don't want
		// users continually haunted by startup error if for some reason this doesn't
		// work on their system:
		if (SystemParametersInfo(SPI_GETFOREGROUNDLOCKTIMEOUT, 0, &g_OriginalTimeout, 0))
			if (g_OriginalTimeout) // Anti-focus stealing measure is in effect.
			{
				// Set it to zero instead, disabling the measure:
				SystemParametersInfo(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, (PVOID)0, SPIF_SENDCHANGE);
//				if (!SystemParametersInfo(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, (PVOID)0, SPIF_SENDCHANGE))
//					MsgBox("Enable focus-stealing: set-call to SystemParametersInfo() failed.");
			}
//			else
//				MsgBox("Enable focus-stealing: it was already enabled.");
//		else
//			MsgBox("Enable focus-stealing: get-call to SystemParametersInfo() failed.");
	}
//	else
//		MsgBox("Enable focus-stealing: neither needed nor supported under Win95 and WinNT.");
}



bool DialogPrep()
// Having it as a function vs. macro should reduce code size due to expansion of macros inside.
{
	bool thread_was_critical = g->ThreadIsCritical;
	g->ThreadIsCritical = false;
	g->AllowThreadToBeInterrupted = true;
	if (HIWORD(GetQueueStatus(QS_ALLEVENTS))) // See DIALOG_PREP for explanation.
		MsgSleep(-1);
	return thread_was_critical; // Caller is responsible for using this to later restore g->ThreadIsCritical.
}
