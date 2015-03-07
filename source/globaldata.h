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

#ifndef globaldata_h
#define globaldata_h

#include "hook.h" // For KeyHistoryItem and probably other things.
#include "clipboard.h"  // For the global clipboard object
#include "script.h" // For the global script object and g_ErrorLevel
#include "os_version.h" // For the global OS_Version object

#include "Debugger.h"

extern HINSTANCE g_hInstance;
extern DWORD g_MainThreadID;
extern DWORD g_HookThreadID;
extern CRITICAL_SECTION g_CriticalRegExCache;

extern UINT g_DefaultScriptCodepage;

extern bool g_DestroyWindowCalled;
extern HWND g_hWnd;  // The main window
extern HWND g_hWndEdit;  // The edit window, child of main.
extern HWND g_hWndSplash;  // The SplashText window.
extern HFONT g_hFontEdit;
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

extern BYTE g_MenuMaskKey; // L38: See #MenuMaskKey.

// If a SendKeys() operation takes longer than this, hotkey's modifiers won't be pressed back down:
extern int g_HotkeyModifierTimeout;
extern int g_ClipboardTimeout;

extern HHOOK g_KeybdHook;
extern HHOOK g_MouseHook;
extern HHOOK g_PlaybackHook;
extern bool g_ForceLaunch;
extern bool g_WinActivateForce;
extern bool g_RunStdIn;
extern WarnMode g_Warn_UseUnsetLocal;
extern WarnMode g_Warn_UseUnsetGlobal;
extern WarnMode g_Warn_UseEnv;
extern WarnMode g_Warn_LocalSameAsGlobal;
extern SingleInstanceType g_AllowOnlyOneInstance;
extern bool g_persistent;
extern bool g_NoTrayIcon;
#ifdef AUTOHOTKEYSC
	extern bool g_AllowMainWindow;
#endif
extern bool g_DeferMessagesForUnderlyingPump;
extern bool g_MainTimerExists;
extern bool g_AutoExecTimerExists;
extern bool g_InputTimerExists;
extern bool g_DerefTimerExists;
extern bool g_SoundWasPlayed;
extern bool g_IsSuspended;
extern BOOL g_WriteCacheDisabledInt64;
extern BOOL g_WriteCacheDisabledDouble;
extern BOOL g_NoEnv;
extern BOOL g_AllowInterruption;
extern int g_nLayersNeedingTimer;
extern int g_nThreads;
extern int g_nPausedThreads;
extern int g_MaxHistoryKeys;

extern VarSizeType g_MaxVarCapacity;
extern UCHAR g_MaxThreadsPerHotkey;
extern int g_MaxThreadsTotal;
extern int g_MaxHotkeysPerInterval;
extern int g_HotkeyThrottleInterval;
extern bool g_MaxThreadsBuffer;
extern SendLevelType g_InputLevel;
extern HotCriterionType g_HotCriterion;
extern LPTSTR g_HotWinTitle;
extern LPTSTR g_HotWinText;
extern HotkeyCriterion *g_FirstHotCriterion, *g_LastHotCriterion;

// Global variables for #if (expression). See globaldata.cpp for comments.
extern int g_HotExprIndex;
extern Line **g_HotExprLines;
extern int g_HotExprLineCount;
extern int g_HotExprLineCountMax;
extern UINT g_HotExprTimeout;
extern HWND g_HotExprLFW;

extern int g_ScreenDPI;
extern MenuTypeType g_MenuIsVisible;
extern int g_nMessageBoxes;
extern int g_nInputBoxes;
extern int g_nFileDialogs;
extern int g_nFolderDialogs;
extern InputBoxType g_InputBox[MAX_INPUTBOXES];
extern SplashType g_Progress[MAX_PROGRESS_WINDOWS];
extern SplashType g_SplashImage[MAX_SPLASHIMAGE_WINDOWS];
extern GuiType **g_gui;
extern int g_guiCount, g_guiCountMax;
extern HWND g_hWndToolTip[MAX_TOOLTIPS];
extern MsgMonitorList g_MsgMonitor;

extern UCHAR g_SortCaseSensitive;
extern bool g_SortNumeric;
extern bool g_SortReverse;
extern int g_SortColumnOffset;
extern Func *g_SortFunc;

extern TCHAR g_delimiter;
extern TCHAR g_DerefChar;
extern TCHAR g_EscapeChar;

// Hot-string vars:
extern TCHAR g_HSBuf[HS_BUF_SIZE];
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
extern TCHAR g_EndChars[HS_MAX_END_CHARS + 1];

// Global objects:
extern Var *g_ErrorLevel;
extern input_type g_input;
EXTERN_SCRIPT;
EXTERN_CLIPBOARD;
EXTERN_OSVER;

extern HICON g_IconSmall;
extern HICON g_IconLarge;

extern DWORD g_OriginalTimeout;

EXTERN_G;
extern global_struct g_default, *g_array;

extern TCHAR g_WorkingDir[MAX_PATH];  // Explicit size needed here in .h file for use with sizeof().
extern LPTSTR g_WorkingDirOrig;

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

enum OurTimers {TIMER_ID_MAIN = MAX_MSGBOXES + 2 // The first timers in the series are used by the MessageBoxes.  Start at +2 to give an extra margin of safety.
	, TIMER_ID_UNINTERRUPTIBLE // Obsolete but kept as a a placeholder for backward compatibility, so that this and the other the timer-ID's stay the same, and so that obsolete IDs aren't reused for new things (in case anyone is interfacing these OnMessage() or with external applications).
	, TIMER_ID_AUTOEXEC, TIMER_ID_INPUT, TIMER_ID_DEREF, TIMER_ID_REFRESH_INTERRUPTIBILITY};

// MUST MAKE main timer and uninterruptible timers associated with our main window so that
// MainWindowProc() will be able to process them when it is called by the DispatchMessage()
// of a non-standard message pump such as MessageBox().  In other words, don't let the fact
// that the script is displaying a dialog interfere with the timely receipt and processing
// of the WM_TIMER messages, including those "hidden messages" which cause DefWindowProc()
// (I think) to call the TimerProc() of timers that use that method.
// Realistically, SetTimer() called this way should never fail?  But the event loop can't
// function properly without it, at least when there are suspended subroutines.
// MSDN docs for SetTimer(): "Windows 2000/XP: If uElapse is less than 10,
// the timeout is set to 10." TO GET CONSISTENT RESULTS across all operating systems,
// it may be necessary never to pass an uElapse parameter outside the range USER_TIMER_MINIMUM
// (0xA) to USER_TIMER_MAXIMUM (0x7FFFFFFF).
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

// See AutoExecSectionTimeout() for why g->AllowThreadToBeInterrupted is used rather than the other var.
// The below also sets g->ThreadStartTime and g->UninterruptibleDuration.  Notes about this:
// In case the AutoExecute section takes a long time (or never completes), allow interruptions
// such as hotkeys and timed subroutines after a short time. Use g->AllowThreadToBeInterrupted
// vs. g_AllowInterruption in case commands in the AutoExecute section need exclusive use of
// g_AllowInterruption (i.e. they might change its value to false and then back to true,
// which would interfere with our use of that var).
// From MSDN: "When you specify a TimerProc callback function, the default window procedure calls the
// callback function when it processes WM_TIMER. Therefore, you need to dispatch messages in the calling thread,
// even when you use TimerProc instead of processing WM_TIMER."  My: This is why all TimerProc type timers
// should probably have an associated window rather than passing NULL as first param of SetTimer().
//
// UPDATE v1.0.48: g->ThreadStartTime and g->UninterruptibleDuration were added so that IsInterruptible()
// won't make the AutoExec section interruptible prematurely.  In prior versions, KILL_AUTOEXEC_TIMER() did this,
// but with the new IsInterruptible() function, doing it in KILL_AUTOEXEC_TIMER() wouldn't be reliable because
// it might already have been done by IsInterruptible() [or vice versa], which might provide a window of
// opportunity in which any use of Critical by the AutoExec section would be undone by the second timeout.
// More info: Since AutoExecSection() never calls InitNewThread(), it never used to set the uninterruptible
// timer.  Instead, it had its own timer.  But now that IsInterruptible() checks for the timeout of
// "Thread Interrupt", AutoExec might become interruptible prematurely unless it uses the new method below.
#define SET_AUTOEXEC_TIMER(aTimeoutValue) \
{\
	g->AllowThreadToBeInterrupted = false;\
	g->ThreadStartTime = GetTickCount();\
	g->UninterruptibleDuration = aTimeoutValue;\
	if (!g_AutoExecTimerExists)\
		g_AutoExecTimerExists = SetTimer(g_hWnd, TIMER_ID_AUTOEXEC, aTimeoutValue, AutoExecSectionTimeout);\
} // v1.0.39 for above: Removed the call to ExitApp() upon failure.  See SET_MAIN_TIMER for details.

#define SET_INPUT_TIMER(aTimeoutValue) \
if (!g_InputTimerExists)\
	g_InputTimerExists = SetTimer(g_hWnd, TIMER_ID_INPUT, aTimeoutValue, InputTimeout);

// For this one, SetTimer() is called unconditionally because our caller wants the timer reset
// (as though it were killed and recreated) unconditionally.  MSDN's comments are a little vague
// about this, but testing shows that calling SetTimer() against an existing timer does completely
// reset it as though it were killed and recreated.  Note also that g_hWnd is used vs. NULL so that
// the timer will fire even when a msg pump other than our own is running, such as that of a MsgBox.
#define SET_DEREF_TIMER(aTimeoutValue) g_DerefTimerExists = SetTimer(g_hWnd, TIMER_ID_DEREF, aTimeoutValue, DerefTimeout);
#define LARGE_DEREF_BUF_SIZE (4*1024*1024)

#define KILL_MAIN_TIMER \
if (g_MainTimerExists && KillTimer(g_hWnd, TIMER_ID_MAIN))\
	g_MainTimerExists = false;

// See above comment about g->AllowThreadToBeInterrupted.
#define KILL_AUTOEXEC_TIMER \
{\
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
