﻿/*
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
#include "WinGroup.h"
#include "window.h" // for several lower level window functions
#include "globaldata.h" // for DoWinDelay
#include "application.h" // for DoWinDelay's MsgSleep()

// Define static members data:
WinGroup *WinGroup::sGroupLastUsed = NULL;
HWND *WinGroup::sAlreadyVisited = NULL;
int WinGroup::sAlreadyVisitedCount = 0;


ResultType WinGroup::AddWindow(LPCTSTR aTitle, LPCTSTR aText, LPCTSTR aExcludeTitle, LPCTSTR aExcludeText)
// Caller should ensure that at least one param isn't NULL/blank.
// This function is not thread-safe because it adds an entry to the list of window specs.
// In addition, if this function is being called by one thread while another thread is calling IsMember(),
// the thread-safety notes in IsMember() apply.
{
	if (!aTitle) aTitle = _T("");
	if (!aText) aText = _T("");
	if (!aExcludeTitle) aExcludeTitle = _T("");
	if (!aExcludeText) aExcludeText = _T("");
	// v1.0.41: If a window group can ever be deleted (or its window specs), that might defeat the
	// thread-safety of WinExist/WinActive.
	// v1.0.36.05: If all four window parameters are blank, allow it to be added but provide
	// a non-blank ExcludeTitle so that the window-finding routines won't see it as the
	// "last found window".  99.99% of the time, it is undesirable to have Program Manager
	// in a window group of any kind, so that is used as the placeholder:
	if (!(*aTitle || *aText || *aExcludeTitle || *aExcludeText))
		aExcludeTitle = _T("Program Manager");

	// Though the documentation is clear on this, some users will still probably execute
	// each GroupAdd statement more than once.  Thus, to prevent more and more memory
	// from being allocated for duplicates, do not add the window specification if it
	// already exists in the group:
	if (mFirstWindow) // Traverse the circular linked-list to look for a match.
		for (WindowSpec *win = mFirstWindow
			; win != NULL; win = (win->mNextWindow == mFirstWindow) ? NULL : win->mNextWindow)
			if (!_tcscmp(win->mTitle, aTitle) && !_tcscmp(win->mText, aText) // All are case sensitive.
				&& !_tcscmp(win->mExcludeTitle, aExcludeTitle) && !_tcscmp(win->mExcludeText, aExcludeText))
				return OK;

	// SimpleHeap::Malloc() will set these new vars to the constant empty string if their
	// corresponding params are blank:
	LPTSTR new_title, new_text, new_exclude_title, new_exclude_text;
	if (!(new_title = SimpleHeap::Malloc(aTitle))) return FAIL; // It already displayed the error for us.
	if (!(new_text = SimpleHeap::Malloc(aText)))return FAIL;
	if (!(new_exclude_title = SimpleHeap::Malloc(aExcludeTitle))) return FAIL;
	if (!(new_exclude_text = SimpleHeap::Malloc(aExcludeText)))   return FAIL;

	// The precise method by which the follows steps are done should be thread-safe even if
	// some other thread calls IsMember() in the middle of the operation.  But any changes
	// must be carefully reviewed:
	WindowSpec *the_new_win = new WindowSpec(new_title, new_text, new_exclude_title, new_exclude_text);
	if (mFirstWindow == NULL)
		mFirstWindow = the_new_win;
	else
		mLastWindow->mNextWindow = the_new_win; // Formerly it pointed to First, so nothing is lost here.
	// This must be done after the above:
	mLastWindow = the_new_win;
	// Make it circular: Last always points to First.  It's okay if it points to itself:
	mLastWindow->mNextWindow = mFirstWindow;
	++mWindowCount;
	return OK;
}



void WinGroup::ActUponAll(BuiltInFunctionID aActionType, int aTimeToWaitForClose)
{
	// A group should never be empty, since it is created by "adding" a set of criteria.
	//if (IsEmpty())
	//	return;
	// Don't need to call Update() in this case.
	WindowSearch ws;
	ws.mFirstWinSpec = mFirstWindow; // Act upon all windows that match any WindowSpec in the group.
	ws.mActionType = aActionType;    // Set the type of action to be performed on each window.
	ws.mTimeToWaitForClose = aTimeToWaitForClose;  // Only relevant for WinClose and WinKill.
	EnumWindows(EnumParentActUponAll, (LPARAM)&ws);
	if (ws.mFoundParent) // It acted upon least one window.
		DoWinDelay;
}



void WinGroup::CloseAndGoToNext(bool aStartWithMostRecent)
// If the foreground window is a member of this group, close it and activate
// the next member.
{
	if (IsEmpty())
		return;
	// Otherwise:
	// Don't call Update(), let (De)Activate() do that.

	HWND fore_win = GetForegroundWindow();
	// Even if it's NULL, don't return since the legacy behavior is to continue on to the final part below.

	WindowSpec *win_spec = IsMember(fore_win, *g);
	if (   (mIsModeActivate && win_spec) || (!mIsModeActivate && !win_spec)   )
	{
		// If the user is using a GroupActivate hotkey, we don't want to close
		// the foreground window if it's not a member of the group.  Conversely,
		// if the user is using GroupDeactivate, we don't want to close a
		// member of the group.  This precaution helps prevent accidental closing
		// of windows that suddenly pop up to the foreground just as you've
		// realized (too late) that you pressed the "close" hotkey.
		// MS Visual Studio/C++ gets messed up when it is directly sent a WM_CLOSE,
		// probably because the wrong window (it has two mains) is being sent the close.
		// But since that's the only app I've ever found that doesn't work right,
		// it seems best not to change our close method just for it because sending
		// keys is a fairly high overhead operation, and not without some risk due to
		// not knowing exactly what keys the user may have physically held down.
		// Also, we'd have to make this module dependent on the keyboard module,
		// which would be another drawback.
		// Try to wait for it to close, otherwise the same window may be activated
		// again before it has been destroyed, defeating the purpose of the
		// "ActivateNext" part of this function's job:
		// SendKeys("!{F4}");
		if (fore_win)
		{
			WinClose(fore_win, 500);
			DoWinDelay;
		}
		if (mIsModeActivate)
		{
			HWND new_fore_win = GetForegroundWindow();
			if (new_fore_win != fore_win && IsMember(new_fore_win, *g))
			{
				// new_fore_win appears to have been activated as a result of closing fore_win.
				// Since new_fore_win belongs to the group, avoid activating some other window
				// even if !aStartWithMostRecent, since it would be a bit jarring.  Historical
				// behaviour was to leave it active only if it belonged to the same *spec*,
				// which doesn't make a lot of sense.
				return;
			}
		}
	}
	//else do the activation below anyway, even though no close was done.
	if (mIsModeActivate)
		Activate(aStartWithMostRecent);
	else
		Deactivate(aStartWithMostRecent);
}



HWND WinGroup::Activate(bool aStartWithMostRecent)
{
	if (IsEmpty())
		return NULL;
	Update(true); // Update our private member vars.
	HWND active_window = GetForegroundWindow();
	bool group_is_active = IsMember(active_window, *g);
	if (!group_is_active) // Foreground window is not a member of this group.
	{
		// Reset the list of visited windows:
		sAlreadyVisitedCount = 0;
	}

	// Find the first/last unvisited window that matches any spec in the group.
	WindowSearch ws;
	// This next line is whether to find last or first match.  We always find the oldest
	// (bottommost) match except when the user has specifically asked to start with the
	// most recent AND the group isn't currently active (i.e. we're starting fresh).
	// This allows the order to be the same each time the sequence repeats, instead of
	// reversing due to the nature of the window z-order.
	ws.mFindLastMatch = !aStartWithMostRecent || group_is_active;
	ws.mAlreadyVisited = sAlreadyVisited;
	ws.mAlreadyVisitedCount = sAlreadyVisitedCount;
	ws.SetCriteria(*g, *this);

	for (bool retry_needed = sAlreadyVisitedCount > 0; ; retry_needed = false)
	{
		if (group_is_active) // Foreground window is a member of this group.
		{
			// Ensure the active window is excluded from both attempts to find a match.
			MarkAsVisited(active_window);
		}
		ws.mAlreadyVisitedCount = sAlreadyVisitedCount;
		EnumWindows(EnumParentFind, (LPARAM)&ws);
		if (ws.mFoundParent)
		{
			// We found a window to activate, so we're done.
			SetForegroundWindowEx(ws.mFoundParent);
			MarkAsVisited(ws.mFoundParent);
			DoWinDelay;
			return ws.mFoundParent;
		}
		if (!retry_needed)
			break;
		// Try again without excluding previously visited windows.
		sAlreadyVisitedCount = 0;
	}
	// No window was found to activate.
	return NULL;
}



void WinGroup::Deactivate(bool aStartWithMostRecent)
{
	Update(false); // Update our private member vars.

	HWND active_window = GetForegroundWindow();
	// Since EnumParentFindAnyExcept does not evaluate owned windows but one may have
	// been activated by a previous call, evaluate the owner of active_window, if any.
	while (HWND owner = GetWindow(active_window, GW_OWNER))
		active_window = owner;
	if (IsMember(active_window, *g))
		sAlreadyVisitedCount = 0;

	// Activate the next unvisited non-member:
	WindowSearch ws;
	ws.mFindLastMatch = !aStartWithMostRecent || sAlreadyVisitedCount;
	ws.mAlreadyVisited = sAlreadyVisited;
	ws.mAlreadyVisitedCount = sAlreadyVisitedCount;
	ws.mFirstWinSpec = mFirstWindow;

	EnumWindows(EnumParentFindAnyExcept, (LPARAM)&ws);

	if (ws.mFoundParent)
	{
		// Contrary to comments previously found here, a window that owns other windows can be
		// (and often is) activated, unless the window is also disabled.  However, when the
		// window is activated via the Taskbar, Alt-Tab or due to closing or minimizing some
		// other window, the OS activates the *last active owned window* if it was active more
		// recently than the owner.  Doing it this way will activate either the owner or one
		// of its owned windows, instead of the old way, activating *every one* of the owned
		// windows in turn and never the owner.
		SetForegroundWindowEx(GetLastActivePopup(ws.mFoundParent));
		// Probably best to do this before WinDelay in case another hotkey fires during the delay:
		MarkAsVisited(ws.mFoundParent); // Mark the found HWND, not necessarily the one that was activated.
		DoWinDelay;
	}
	else // No window was found to activate (they have all been visited).
	{
		if (sAlreadyVisitedCount)
		{
			bool wrap_around = (sAlreadyVisitedCount > 1);
			sAlreadyVisitedCount = 0;
			if (wrap_around)
			{
				// The user pressed a hotkey to do something, yet nothing has happened yet.
				// We want something to happen every time if there's a qualifying
				// "something" that we can do.  And in this case there is: we can start
				// over again through the list, excluding the foreground window (which
				// the user has already had a chance to review):
				MarkAsVisited(active_window);
				// Make a recursive call to self.  This can't result in an infinite
				// recursion (stack fault) because the called layer will only
				// recurse a second time if sAlreadyVisitedCount > 1, which is
				// impossible with the current logic:
				Deactivate(false); // Seems best to ignore aStartWithMostRecent in this case?
			}
		}
		else
		{
			// Since there are apparently no eligible windows, deactivate the window by
			// activating the taskbar, which is usually what happens when you minimize
			// the last non-minimized top-level window.
			SetForegroundWindowEx(FindWindow(_T("Shell_TrayWnd"), nullptr));
		}
	}
}



inline void WinGroup::Update(bool aIsModeActivate)
{
	mIsModeActivate = aIsModeActivate;
	if (sGroupLastUsed != this)
	{
		sGroupLastUsed = this;
		sAlreadyVisitedCount = 0; // Since it's a new group, reset the array to start fresh.
	}
	if (!sAlreadyVisited) // Allocate the array on first use.
		// Getting it from SimpleHeap reduces overhead for the avg. case (i.e. the first
		// block of SimpleHeap is usually never fully used, and this array won't even
		// be allocated for short scripts that don't even use window groups.
		sAlreadyVisited = SimpleHeap::Alloc<HWND>(MAX_ALREADY_VISITED);
}



WindowSpec *WinGroup::IsMember(HWND aWnd, ScriptThreadSettings &aSettings)
// Thread-safety: This function is thread-safe even when the main thread happens to be calling AddWindow()
// and changing the linked list while it's being traversed here by the hook thread.  However, any subsequent
// changes to this function or AddWindow() must be carefully reviewed.
// Although our caller may be a WindowSearch method, and thus we might make
// a recursive call back to that same method, things have been reviewed to ensure that
// thread-safety is maintained, even if the calling thread is the hook.
{
	if (!aWnd)
		return NULL;  // Some callers on this.
	WindowSearch ws;
	ws.SetCandidate(aWnd);
	for (WindowSpec *win = mFirstWindow; win != NULL;)  // v1.0.41: "win != NULL" was added for thread-safety.
	{
		if (ws.SetCriteria(aSettings, win->mTitle, win->mText, win->mExcludeTitle, win->mExcludeText) && ws.IsMatch())
			return win;
		// Otherwise, no match, so go onto the next one:
		win = win->mNextWindow;
		if (!win || win == mFirstWindow) // v1.0.41: The check of !win was added for thread-safety.
			// We've made one full circuit of the circular linked list,
			// discovering that the foreground window isn't a member
			// of the group:
			break;
	}
	return NULL;  // Because it would have returned already if a match was found.
}


/////////////////////////////////////////////////////////////////////////


BOOL CALLBACK EnumParentFindAnyExcept(HWND aWnd, LPARAM lParam)
// Find the first parent window that doesn't match any of the WindowSpecs in
// the linked list, and that hasn't already been visited.
{
	// Since the following two sections apply only to GroupDeactivate (since that's our only caller),
	// they both seem okay even in light of the ahk_group method.

	DWORD style = GetWindowLong(aWnd, GWL_STYLE);
	if (!(style & WS_VISIBLE) || (style & WS_DISABLED) && GetLastActivePopup(aWnd) == aWnd)
		// Skip hidden windows because we always want them to stay invisible, regardless
		// of the setting for g->DetectHiddenWindows.  v2.0: Also skip disabled windows,
		// except those which own a popup window (such as a modal dialog).
		// Since these are top-level windows, IsWindowVisible() shouldn't be necessary.
		return TRUE;

	// UPDATE: Because the window of class Shell_TrayWnd (the taskbar) is also always-on-top,
	// the below prevents it from ever being activated too, which is almost always desirable.
	// Skip always-on-top windows, such as SplashText, because probably shouldn't
	// be activated (especially in this mode, which is often used to visit the user's
	// "non-favorite" windows).  In addition, they're already visible so the user already
	// knows about them, so there's no need to have them presented for review.
	// v2.0:
	//  - Windows with WS_EX_NOACTIVATE are excluded because they probably aren't intended
	//    to be activated.  Past testing indicated they were omitted from Alt-Tab and/or
	//    the taskbar on some OSes, but there must be some other factor as I currently can't
	//    reproduce that.  However, they are skipped by the Alt+Esc/Alt+Shift+Esc hotkeys.
	//  - Windows with WS_EX_TOOLWINDOW but without WS_EX_APPWINDOW are excluded from the
	//    taskbar and Alt-Tab.  They are also usually owned anyway, except for the WorkerW
	//    window which contains some of the Desktop content on Windows 10 (and 7, but it's
	//    hidden there).  Excluding this style avoids the need to check for WorkerW, which
	//    also means we don't need to retrieve the window class.
	DWORD ex_style = GetWindowLong(aWnd, GWL_EXSTYLE);
	if ((ex_style & (WS_EX_TOPMOST | WS_EX_NOACTIVATE))
		|| (ex_style & (WS_EX_TOOLWINDOW | WS_EX_APPWINDOW)) == WS_EX_TOOLWINDOW)
		return TRUE;

	// Exclude owned windows.  The owner should be considered by a subsequent call.
	// If eligible, either the owner or its last active owned window will be activated.
	if (GetWindow(aWnd, GW_OWNER))
		return TRUE;

	if (IsWindowCloaked(aWnd)) // This is probably higher cost than the style checks, so done after.
		return TRUE;
	
	// Skip "Program Manager" (the Desktop) because the user probably doesn't want it
	// activated (although that would serve the purpose of deactivating any other window).
	if (aWnd == GetShellWindow())
		return TRUE;

	WindowSearch &ws = *(WindowSearch *)lParam;  // For performance and convenience.
	ws.SetCandidate(aWnd);

	// Check this window's attributes against each set of criteria present in the group.  If
	// it's a match for any set of criteria, it's a member of the group and thus should be
	// excluded since we want only NON-members:
	for (WindowSpec *win = ws.mFirstWinSpec;;)
	{
		// For each window in the linked list, check if aWnd is a match for it:
		if (ws.SetCriteria(*g, win->mTitle, win->mText, win->mExcludeTitle, win->mExcludeText) && ws.IsMatch(true))
			// Match found, so aWnd is a member of the group. But we want to find non-members only,
			// so keep searching:
			return TRUE;
		// Otherwise, no match, but keep checking until aWnd has been compared against
		// all the WindowSpecs in the group:
		win = win->mNextWindow;
		if (win == ws.mFirstWinSpec)
		{
			// We've made one full circuit of the circular linked list without
			// finding a match.  So aWnd is the one we're looking for unless
			// it's in the list of exceptions:
			for (int i = 0; i < ws.mAlreadyVisitedCount; ++i)
				if (aWnd == ws.mAlreadyVisited[i])
					return TRUE; // It's an exception, so keep searching.
			// Otherwise, this window meets the criteria, so return it to the caller and
			// stop the enumeration.  UPDATE: Rather than stopping the enumeration,
			// continue on through all windows so that the last match is found.
			// That makes things nicer because when the sequence wraps around to the
			// beginning, the windows will occur in the same order that they did
			// the first time, rather than going backwards through the sequence
			// (which is counterintuitive for the user):
			ws.mFoundParent = aWnd; // No need to increment ws.mFoundCount in this case.
			return ws.mFindLastMatch;
		}
	} // The loop above is infinite unless a "return" is encountered inside.
}



BOOL CALLBACK EnumParentActUponAll(HWND aWnd, LPARAM lParam)
// Caller must have ensured that lParam isn't NULL and that it contains a non-NULL mFirstWinSpec.
{
	WindowSearch &ws = *(WindowSearch *)lParam;  // For performance and convenience.

	// Skip windows the command isn't supposed to detect.  ACT_WINSHOW is exempt because
	// hidden windows are always detected by the WinShow command:
	if (!(ws.mActionType == FID_WinShow || g->DetectWindow(aWnd)))
		return TRUE;

	int nCmdShow;
	ws.SetCandidate(aWnd);

	for (WindowSpec *win = ws.mFirstWinSpec;;)
	{
		// For each window in the linked list, check if aWnd is a match for it:
		if (ws.SetCriteria(*g, win->mTitle, win->mText, win->mExcludeTitle, win->mExcludeText) && ws.IsMatch())
		{
			// Match found, so aWnd is a member of the group.  In addition, IsMatch() has set
			// the value of ws.mFoundParent to tell our caller that at least one window was acted upon.
			// See BIF_WinShow for comments about the following section.
			nCmdShow = SW_NONE; // Set default each time.
			switch (ws.mActionType)
			{
			case FID_WinClose:
			case FID_WinKill:
				// mTimeToWaitForClose is not done in a very efficient way here: to keep code size
				// in check, it closes each window individually rather than sending WM_CLOSE to all
				// the windows simultaneously and then waiting until all have vanished:
				WinClose(aWnd, ws.mTimeToWaitForClose, ws.mActionType == FID_WinKill); // DoWinDelay is done by our caller.
				return TRUE; // All done with the current window, so fetch the next one.

			case FID_WinMinimize:
				if (IsWindowHung(aWnd))
					nCmdShow = SW_FORCEMINIMIZE;
				else
					nCmdShow = SW_MINIMIZE;
				break;

			case FID_WinMaximize: if (!IsWindowHung(aWnd)) nCmdShow = SW_MAXIMIZE; break;
			case FID_WinRestore:  if (!IsWindowHung(aWnd)) nCmdShow = SW_RESTORE;  break;
			case FID_WinHide: nCmdShow = SW_HIDE; break;
			case FID_WinShow: nCmdShow = SW_SHOW; break;
			}

			if (nCmdShow != SW_NONE)
				ShowWindow(aWnd, nCmdShow);
				// DoWinDelay is not done here because our caller will do it once only, which
				// seems best when there are a lot of windows being acted upon here.

			// Now that this matching window has been acted upon (or avoided due to being hung),
			// continue the enumeration to get the next candidate window:
			return TRUE;
		}
		// Otherwise, no match, keep checking until aWnd has been compared against all the WindowSpecs in the group:
		win = win->mNextWindow;
		if (win == ws.mFirstWinSpec)
			// We've made one full circuit of the circular linked list without
			// finding a match, so aWnd is not a member of the group and
			// should not be closed.
			return TRUE; // Continue the enumeration.
	}
}
