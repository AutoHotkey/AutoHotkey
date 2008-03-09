/*
AutoHotkey

Copyright 2003-2008 Chris Mallett (support@autohotkey.com)

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
*/

#ifndef globaldata_h
#define globaldata_h

#include "hook.h" // For KeyHistoryItem and probably other things.
#include "clipboard.h"  // For the global clipboard object
#include "script.h" // For the global script object and g_ErrorLevel
#include "os_version.h" // For the global OS_Version object

extern HINSTANCE g_hInstance;
extern DWORD g_MainThreadID;
extern DWORD g_HookThreadID;
extern CRITICAL_SECTION g_CriticalRegExCache;

extern bool g_DestroyWindowCalled;
extern HWND g_hWnd;  // The main window
extern HWND g_hWndEdit;  // The edit window, child of main.
extern HWND g_hWndSplash;  // The SplashText window.
extern HFONT g_hFontSplash;
extern HACCEL g_hAccelTable; // Accelerator table for main menu shortcut keys.

typedef int (WINAPI *StrCmpLogicalW_type)(LPCWSTR, LPCWSTR);
extern StrCmpLogicalW_type g_StrCmpLogicalW;
extern WNDPROC g_TabClassProc;

extern modLR_type g_modifiersLR_logical;   // Tracked by hook (if hook is active).
extern modLR_type g_modifiersLR_logical_non_ignored;
extern modLR_type g_modifiersLR_physical;  // Same as above except it's which modifiers are PHYSICALLY down.

#ifdef FUTURE_USE_MOUSE_BUTTONS_LOGICAL
extern WORD g_mouse_buttons_logical; // A bitwise combination of MK_LBUTTON, etc.
#endif

#define STATE_DOWN 0x80
#define STATE_ON 0x01
extern BYTE g_PhysicalKeyState[VK_ARRAY_COUNT];
extern bool g_BlockWinKeys;
extern DWORD g_HookReceiptOfLControlMeansAltGr;
extern DWORD g_IgnoreNextLControlDown;
extern DWORD g_IgnoreNextLControlUp;

// If a SendKeys() operation takes longer than this, hotkey's modifiers won't be pressed back down:
extern int g_HotkeyModifierTimeout;
extern int g_ClipboardTimeout;

extern HHOOK g_KeybdHook;
extern HHOOK g_MouseHook;
extern HHOOK g_PlaybackHook;
extern bool g_ForceLaunch;
extern bool g_WinActivateForce;
extern SingleInstanceType g_AllowOnlyOneInstance;
extern bool g_persistent;
extern bool g_NoEnv;
extern bool g_NoTrayIcon;
#ifdef AUTOHOTKEYSC
	extern bool g_AllowMainWindow;
#endif
extern bool g_AllowSameLineComments;
extern bool g_AllowInterruption;
extern bool g_DeferMessagesForUnderlyingPump;
extern bool g_MainTimerExists;
extern bool g_UninterruptibleTimerExists;
extern bool g_AutoExecTimerExists;
extern bool g_InputTimerExists;
extern bool g_DerefTimerExists;
extern bool g_SoundWasPlayed;
extern bool g_IsSuspended;
extern int g_nLayersNeedingTimer;
extern int g_nThreads;
extern int g_nPausedThreads;
extern bool g_IdleIsPaused;
extern int g_MaxHistoryKeys;

extern VarSizeType g_MaxVarCapacity;
// This value is the absolute limit:
#define MAX_THREADS_LIMIT 20
#define MAX_THREADS_DEFAULT 10
extern UCHAR g_MaxThreadsPerHotkey;
extern int g_MaxThreadsTotal;
extern int g_MaxHotkeysPerInterval;
extern int g_HotkeyThrottleInterval;
extern bool g_MaxThreadsBuffer;
extern HotCriterionType g_HotCriterion;
extern char *g_HotWinTitle;
extern char *g_HotWinText;
extern HotkeyCriterion *g_FirstHotCriterion, *g_LastHotCriterion;

extern MenuTypeType g_MenuIsVisible;
extern int g_nMessageBoxes;
extern int g_nInputBoxes;
extern int g_nFileDialogs;
extern int g_nFolderDialogs;
extern InputBoxType g_InputBox[MAX_INPUTBOXES];
extern SplashType g_Progress[MAX_PROGRESS_WINDOWS];
extern SplashType g_SplashImage[MAX_SPLASHIMAGE_WINDOWS];
EXTERN_GUI;
extern HWND g_hWndToolTip[MAX_TOOLTIPS];
extern MsgMonitorStruct *g_MsgMonitor; // An array to be allocated upon first use (if any).
extern int g_MsgMonitorCount;

extern UCHAR g_SortCaseSensitive;
extern bool g_SortNumeric;
extern bool g_SortReverse;
extern int g_SortColumnOffset;
extern Func *g_SortFunc;

extern char g_delimiter;
extern char g_DerefChar;
extern char g_EscapeChar;

// Hot-string vars:
extern char g_HSBuf[HS_BUF_SIZE];
extern int g_HSBufLength;
extern HWND g_HShwnd;

// Hot-string global settings:
extern int g_HSPriority;
extern int g_HSKeyDelay;
extern SendModes g_HSSendMode;
extern bool g_HSCaseSensitive;
extern bool g_HSConformToCase;
extern bool g_HSDoBackspace;
extern bool g_HSOmitEndChar;
extern bool g_HSSendRaw;
extern bool g_HSEndCharRequired;
extern bool g_HSDetectWhenInsideWord;
extern bool g_HSDoReset;
extern bool g_HSResetUponMouseClick;
extern char g_EndChars[HS_MAX_END_CHARS + 1];

// Global objects:
extern Var *g_ErrorLevel;
extern input_type g_input;
EXTERN_SCRIPT;
EXTERN_CLIPBOARD;
EXTERN_OSVER;

extern int g_IconTray;
extern int g_IconTraySuspend;

extern DWORD g_OriginalTimeout;

EXTERN_G;
extern global_struct g_default;

extern char g_WorkingDir[MAX_PATH];  // Explicit size needed here in .h file for use with sizeof().
extern char *g_WorkingDirOrig;

extern bool g_ContinuationLTrim;
extern bool g_ForceKeybdHook;
extern ToggleValueType g_ForceNumLock;
extern ToggleValueType g_ForceCapsLock;
extern ToggleValueType g_ForceScrollLock;

extern ToggleValueType g_BlockInputMode;
extern bool g_BlockInput;  // Whether input blocking is currently enabled.
extern bool g_BlockMouseMove; // Whether physical mouse movement is currently blocked via the mouse hook.

extern Action g_act[];
extern int g_ActionCount;
extern Action g_old_act[];
extern int g_OldActionCount;

extern key_to_vk_type g_key_to_vk[];
extern key_to_sc_type g_key_to_sc[];
extern int g_key_to_vk_count;
extern int g_key_to_sc_count;

extern KeyHistoryItem *g_KeyHistory;
extern int g_KeyHistoryNext;
extern DWORD g_HistoryTickNow;
extern DWORD g_HistoryTickPrev;
extern HWND g_HistoryHwndPrev;
extern DWORD g_TimeLastInputPhysical;

#ifdef ENABLE_KEY_HISTORY_FILE
extern bool g_KeyHistoryToFile;
#endif


// 9 might be better than 10 because if the granularity/timer is a little
// off on certain systems, a Sleep(10) might really result in a Sleep(20),
// whereas a Sleep(9) is almost certainly a Sleep(10) on OS's such as
// NT/2k/XP.  UPDATE: Roundoff issues with scripts having
// even multiples of 10 in them, such as "Sleep,300", shouldn't be hurt
// by this because they use GetTickCount() to verify how long the
// sleep duration actually was.  UPDATE again: Decided to go back to 10
// because I'm pretty confident that that always sleeps 10 on NT/2k/XP
// unless the system is under load, in which case any Sleep between 0
// and 20 inclusive seems to sleep for exactly(?) one timeslice.
// A timeslice appears to be 20ms in duration.  Anyway, using 10
// allows "SetKeyDelay, 10" to be really 10 rather than getting
// rounded up to 20 due to doing first a Sleep(10) and then a Sleep(1).
// For now, I'm avoiding using timeBeginPeriod to improve the resolution
// of Sleep() because of possible incompatibilities on some systems,
// and also because it may degrade overall system performance.
// UPDATE: Will get rounded up to 10 anyway by SetTimer().  However,
// future OSs might support timer intervals of less than 10.
#define SLEEP_INTERVAL 10
#define SLEEP_INTERVAL_HALF (int)(SLEEP_INTERVAL / 2)

// The first timers in the series are used by the MessageBoxes.  Start at +2 to give
// an extra margin of safety:
enum OurTimers {TIMER_ID_MAIN = MAX_MSGBOXES + 2, TIMER_ID_UNINTERRUPTIBLE, TIMER_ID_AUTOEXEC
	, TIMER_ID_INPUT, TIMER_ID_DEREF};

// MUST MAKE main timer and uninterruptible timers associated with our main window so that
// MainWindowProc() will be able to process them when it is called by the DispatchMessage()
// of a non-standard message pump such as MessageBox().  In other words, don't let the fact
// that the script is displaying a dialog interfere with the timely receipt and processing
// of the WM_TIMER messages, including those "hidden messages" which cause DefWindowProc()
// (I think) to call the TimerProc() of timers that use that method.
// Realistically, SetTimer() called this way should never fail?  But the event loop can't
// function properly without it, at least when there are suspended subroutines.
// MSDN docs for SetTimer(): "Windows 2000/XP: If uElapse is less than 10,
// the timeout is set to 10."
#define SET_MAIN_TIMER \
if (!g_MainTimerExists)\
	g_MainTimerExists = SetTimer(g_hWnd, TIMER_ID_MAIN, SLEEP_INTERVAL, (TIMERPROC)NULL);
// v1.0.39 for above: Apparently, one of the few times SetTimer fails is after the thread has done
// PostQuitMessage. That particular failure was causing an unwanted recursive call to ExitApp(),
// which is why the above no longer calls ExitApp on failure.  Here's the sequence:
// Someone called ExitApp (such as the max-hotkeys-per-interval warning dialog).
// ExitApp() removes the hooks.
// The hook-removal function calls MsgSleep() while waiting for the hook-thread to finish.
// MsgSleep attempts to set the main timer so that it can judge how long to wait.
// The timer fails and calls ExitApp even though a previous call to ExitApp is currently underway.

// When someone calls SET_UNINTERRUPTIBLE_TIMER, by definition the current script subroutine is
// becoming non-interruptible.  Therefore, their should never be a need to have more than one
// of these timer going simultanously since they're only created for the launch of a new
// quasi-thread, which is forbidden when the current thread is uninterruptible.
// Remember than the 2nd param of SetTimer() is ignored when the 1st param is NULL.
// For this one, the timer is not recreated if it already exists because I think SetTimer(), when
// called with NULL as a first parameter, may wind up creating more than one Timer and we only
// want one of these to exist at a time.
// The caller should ensure that aTimeoutValue is <= INT_MAX because otherwise SetTimer()'s behavior
// will vary depending on OS type & version.
// Also have this one abort on unexpected error, since failure to set the timer might result in the
// script becoming permanently uninterruptible (which prevents new hotkeys from being activated
// even though the program is still responsive).
#define SET_UNINTERRUPTIBLE_TIMER \
if (!g_UninterruptibleTimerExists)\
	g_UninterruptibleTimerExists = SetTimer(g_hWnd, TIMER_ID_UNINTERRUPTIBLE \
		, g_script.mUninterruptibleTime < 10 ? 10 : g_script.mUninterruptibleTime, UninterruptibleTimeout);
// v1.0.39 for above: Removed the call to ExitApp() upon failure.  See SET_MAIN_TIMER for details.

#define KILL_UNINTERRUPTIBLE_TIMER \
if (g_UninterruptibleTimerExists && KillTimer(g_hWnd, TIMER_ID_UNINTERRUPTIBLE))\
	g_UninterruptibleTimerExists = false;

// See AutoExecSectionTimeout() for why g.AllowThreadToBeInterrupted is used rather than the other var.
// Also, from MSDN: "When you specify a TimerProc callback function, the default window procedure calls the
// callback function when it processes WM_TIMER. Therefore, you need to dispatch messages in the calling thread,
// even when you use TimerProc instead of processing WM_TIMER."  My: This is why all TimerProc type timers
// should probably have a window rather than passing NULL as first param of SetTimer().
#define SET_AUTOEXEC_TIMER(aTimeoutValue) \
{\
	g.AllowThreadToBeInterrupted = false;\
	if (!g_AutoExecTimerExists)\
		g_AutoExecTimerExists = SetTimer(g_hWnd, TIMER_ID_AUTOEXEC, aTimeoutValue, AutoExecSectionTimeout);\
} // v1.0.39 for above: Removed the call to ExitApp() upon failure.  See SET_MAIN_TIMER for details.

#define SET_INPUT_TIMER(aTimeoutValue) \
if (!g_InputTimerExists)\
	g_InputTimerExists = SetTimer(g_hWnd, TIMER_ID_INPUT, aTimeoutValue, InputTimeout);

// For this one, SetTimer() is called unconditionally because our caller wants the timer reset
// (as though it were killed and recreated) uncondtionally.  MSDN's comments are a little vague
// about this, but testing shows that calling SetTimer() against an existing timer does completely
// reset it as though it were killed and recreated.  Note also that g_hWnd is used vs. NULL so that
// the timer will fire even when a msg pump other than our own is running, such as that of a MsgBox.
#define LARGE_DEREF_BUF_SIZE (4*1024*1024)

#define KILL_MAIN_TIMER \
if (g_MainTimerExists && KillTimer(g_hWnd, TIMER_ID_MAIN))\
	g_MainTimerExists = false;

// Although the caller doesn't always need g.AllowThreadToBeInterrupted reset to true,
// it's much more maintainable and nicer to do it unconditionally due to the complexity of
// managing quasi-threads.  At the very least, it's needed for when the "idle thread"
// is "resumed" (see MsgSleep for explanation).
#define MAKE_THREAD_INTERRUPTIBLE \
{\
	g.AllowThreadToBeInterrupted = true;\
	KILL_UNINTERRUPTIBLE_TIMER \
}

// See above comment about g.AllowThreadToBeInterrupted.
// Also, must restore to true in this case since auto-exec section isn't run as a new thread
// (i.e. there's nothing to resume).
#define KILL_AUTOEXEC_TIMER \
{\
	g.AllowThreadToBeInterrupted = !g.ThreadIsCritical;\
	if (g_AutoExecTimerExists && KillTimer(g_hWnd, TIMER_ID_AUTOEXEC))\
		g_AutoExecTimerExists = false;\
}

#define KILL_INPUT_TIMER \
if (g_InputTimerExists && KillTimer(g_hWnd, TIMER_ID_INPUT))\
	g_InputTimerExists = false;

#define KILL_DEREF_TIMER \
if (g_DerefTimerExists && KillTimer(g_hWnd, TIMER_ID_DEREF))\
	g_DerefTimerExists = false;

#endif
