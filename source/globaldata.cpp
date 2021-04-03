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
// These includes should probably a superset of those in globaldata.h:
#include "hook.h" // For KeyHistoryItem and probably other things.
#include "clipboard.h"  // For the global clipboard object
#include "script.h" // For the global script object
#include "os_version.h" // For the global OS_Version object

#include "Debugger.h"

// Since at least some of some of these (e.g. g_modifiersLR_logical) should not
// be kept in the struct since it's not correct to save and restore their
// state, don't keep anything in the global_struct except those things
// which are necessary to save and restore (even though it would clean
// up the code and might make maintaining it easier):
HINSTANCE g_hInstance = NULL; // Set by WinMain().
DWORD g_MainThreadID = GetCurrentThreadId();
DWORD g_HookThreadID; // Not initialized by design because 0 itself might be a valid thread ID.
CRITICAL_SECTION g_CriticalRegExCache;

UINT g_DefaultScriptCodepage = CP_UTF8;

bool g_DestroyWindowCalled = false;
HWND g_hWnd = NULL;
HWND g_hWndEdit = NULL;
HFONT g_hFontEdit = NULL;
HACCEL g_hAccelTable = NULL;

WNDPROC g_TabClassProc = NULL;

modLR_type g_modifiersLR_logical = 0;
modLR_type g_modifiersLR_logical_non_ignored = 0;
modLR_type g_modifiersLR_physical = 0;
modLR_type g_modifiersLR_numpad_mask = 0;
modLR_type g_modifiersLR_ctrlaltdel_mask = 0;

#ifdef FUTURE_USE_MOUSE_BUTTONS_LOGICAL
WORD g_mouse_buttons_logical = 0;
#endif

// Used by the hook to track physical state of all virtual keys, since GetAsyncKeyState() does
// not retrieve the physical state of a key.  Note that this array is sometimes used in a way that
// requires its format to be the same as that returned from GetKeyboardState():
BYTE g_PhysicalKeyState[VK_ARRAY_COUNT] = {0};
bool g_BlockWinKeys = false;
DWORD g_AltGrExtraInfo = 0;

BYTE g_MenuMaskKeyVK = VK_CONTROL; // For #MenuMaskKey.
USHORT g_MenuMaskKeySC = SC_LCONTROL;

int g_HotkeyModifierTimeout = 50;  // Reduced from 100, which was a little too large for fast typists.
int g_ClipboardTimeout = 1000; // v1.0.31

HHOOK g_KeybdHook = NULL;
HHOOK g_MouseHook = NULL;
HHOOK g_PlaybackHook = NULL;
bool g_ForceLaunch = false;
bool g_WinActivateForce = false;
bool g_RunStdIn = false;
WarnMode g_Warn_LocalSameAsGlobal = WARNMODE_OFF;
WarnMode g_Warn_Unreachable = WARNMODE_MSGBOX;
WarnMode g_Warn_VarUnset = WARNMODE_MSGBOX;
SingleInstanceType g_AllowOnlyOneInstance = SINGLE_INSTANCE_PROMPT;
bool g_persistent = false;  // Whether the script should stay running even after the auto-exec section finishes.
bool g_NoTrayIcon = false;
#ifdef AUTOHOTKEYSC
	bool g_AllowMainWindow = false;
#endif
bool g_MainTimerExists = false;
bool g_InputTimerExists = false;
bool g_DerefTimerExists = false;
bool g_SoundWasPlayed = false;
bool g_IsSuspended = false;  // Make this separate from g_AllowInterruption since that is frequently turned off & on.
bool g_DeferMessagesForUnderlyingPump = false;
bool g_OnExitIsRunning = false;
BOOL g_AllowInterruption = TRUE;  // BOOL vs. bool might improve performance a little for frequently-accessed variables.
int g_nLayersNeedingTimer = 0;
int g_nThreads = 0;
int g_nPausedThreads = 0;
int g_MaxHistoryKeys = 40;
DWORD g_InputTimeoutAt = 0;

UCHAR g_MaxThreadsPerHotkey = 1;
int g_MaxThreadsTotal = MAX_THREADS_DEFAULT;
// On my system, the repeat-rate (which is probably set to XP's default) is such that between 20
// and 25 keys are generated per second.  Therefore, 50 in 2000ms seems like it should allow the
// key auto-repeat feature to work on most systems without triggering the warning dialog.
// In any case, using auto-repeat with a hotkey is pretty rare for most people, so it's best
// to keep these values conservative:
UINT g_MaxHotkeysPerInterval = 70; // Increased to 70 because 60 was still causing the warning dialog for repeating keys sometimes.  Increased from 50 to 60 for v1.0.31.02 since 50 would be triggered by keyboard auto-repeat when it is set to its fastest.
UINT g_HotkeyThrottleInterval = 2000; // Milliseconds.
bool g_MaxThreadsBuffer = false;  // This feature usually does more harm than good, so it defaults to OFF.
bool g_SuspendExempt = false; // #SuspendExempt, applies to hotkeys and hotstrings.
bool g_SuspendExemptHS = false; // This is just to prevent #Hotstring "S" from affecting hotkeys.
SendLevelType g_InputLevel = 0;
HotkeyCriterion *g_FirstHotCriterion = NULL, *g_LastHotCriterion = NULL;

// Global variables for #HotIf (expression).
UINT g_HotExprTimeout = 1000; // Timeout for #HotIf (expression) evaluation, in milliseconds.
HWND g_HotExprLFW = NULL; // Last Found Window of last #HotIf expression.
HotkeyCriterion *g_FirstHotExpr = NULL, *g_LastHotExpr = NULL;

static int GetScreenDPI()
{
	// The DPI setting can be different for each screen axis, but
	// apparently it is such a rare situation that it is not worth
	// supporting it. So we just retrieve the X axis DPI.

	HDC hdc = GetDC(NULL);
	int dpi = GetDeviceCaps(hdc, LOGPIXELSX);
	ReleaseDC(NULL, hdc);
	return dpi;
}

int g_ScreenDPI = GetScreenDPI();
MenuTypeType g_MenuIsVisible = MENU_TYPE_NONE;
int g_nMessageBoxes = 0;
int g_nFileDialogs = 0;
int g_nFolderDialogs = 0;
GuiType *g_firstGui = NULL, *g_lastGui = NULL;
HWND g_hWndToolTip[MAX_TOOLTIPS] = {NULL};
MsgMonitorList g_MsgMonitor;

// Init not needed for these:
UCHAR g_SortCaseSensitive;
bool g_SortNumeric;
bool g_SortReverse;
int g_SortColumnOffset;
IObject *g_SortFunc;
ResultType g_SortFuncResult;

// Hot-string vars (initialized when ResetHook() is first called):
TCHAR g_HSBuf[HS_BUF_SIZE];
int g_HSBufLength;
HWND g_HShwnd;

// Hot-string global settings:
int g_HSPriority = 0;  // default priority is always 0
int g_HSKeyDelay = 0;  // Fast sends are much nicer for auto-replace and auto-backspace.
SendModes g_HSSendMode = SM_INPUT; // v1.0.43: New default for more reliable hotstrings.
SendRawType g_HSSendRaw = SCM_NOT_RAW;
bool g_HSCaseSensitive = false;
bool g_HSConformToCase = true;
bool g_HSDoBackspace = true;
bool g_HSOmitEndChar = false;
bool g_HSEndCharRequired = true;
bool g_HSDetectWhenInsideWord = false;
bool g_HSDoReset = false;
bool g_HSResetUponMouseClick = true;
bool g_HSSameLineAction = false;
TCHAR g_EndChars[HS_MAX_END_CHARS + 1] = _T("-()[]{}:;'\"/\\,.?!\n \t");  // Hotstring default end chars, including a space.
// The following were considered but seemed too rare and/or too likely to result in undesirable replacements
// (such as while programming or scripting, or in usernames or passwords): <>*+=_%^&|@#$|
// Although dash/hyphen is used for multiple purposes, it seems to me that it is best (on average) to include it.
// Jay D. Novak suggested ([{/ for things such as fl/nj or fl(nj) which might resolve to USA state names.
// i.e. word(synonym) and/or word/synonym

// Global objects:
input_type *g_input = NULL;
Script g_script;
// This made global for performance reasons (determining size of clipboard data then
// copying contents in or out without having to close & reopen the clipboard in between):
Clipboard g_clip;
OS_Version g_os;  // OS version object, courtesy of AutoIt3.

HICON g_IconSmall;
HICON g_IconLarge;

global_struct g_startup, *g_array;
global_struct *g = &g_startup; // g_startup provides a non-NULL placeholder during script loading. Afterward it's replaced with an array.

// I considered maintaining this on a per-quasi-thread basis (i.e. in global_struct), but the overhead
// of having to check and restore the working directory when a suspended thread is resumed (especially
// when the script has many high-frequency timers), and possibly changing the working directory
// whenever a new thread is launched, doesn't seem worth it.  This is because the need to change
// the working directory is comparatively rare:
CString g_WorkingDir;
LPTSTR g_WorkingDirOrig = NULL;  // Assigned a value in WinMain().

bool g_ForceKeybdHook = false;
ToggleValueType g_ForceNumLock = NEUTRAL;
ToggleValueType g_ForceCapsLock = NEUTRAL;
ToggleValueType g_ForceScrollLock = NEUTRAL;

ToggleValueType g_BlockInputMode = TOGGLE_DEFAULT;
bool g_BlockInput = false;
bool g_BlockMouseMove = false;

// The order of initialization here must match the order in the enum contained in script.h
// It's in there rather than in globaldata.h so that the action-type constants can be referred
// to without having access to the global array itself (i.e. it avoids having to include
// globaldata.h in modules that only need access to the enum's constants, which in turn prevents
// many mutual dependency problems between modules).  Note: Action names must not contain any
// spaces or tabs because within a script, those characters can be used in lieu of a delimiter
// to separate the action-type-name from the first parameter.

// STEPS TO ADD A NEW COMMAND:
// 1) Add an entry to the command enum in script.h.
// 2) Add an entry to the below array (it's position here MUST exactly match that in the enum).
//    The first item is the command name, the second is the minimum number of parameters (e.g.
//    if you enter 3, the first 3 args are mandatory) and the third is the maximum number of
//    parameters (the user need not escape commas within the last parameter).
//    Note: If you use a value for MinParams than is greater than zero, remember than any params
//    beneath that threshold will also be required to be non-blank (i.e. user can't omit them even
//    if later, non-blank params are provided).
// 3) If the new command has any params that are output or input vars, change Line::ArgIsVar().
// 4) Implement the command in Line::ExecUntil().
//    If the command waits for anything (e.g. calls MsgSleep()), be sure to make a local
//    copy of any ARG values that are needed during the wait period, because if another hotkey
//    subroutine suspends the current one while its waiting, it could also overwrite the ARG
//    deref buffer with its own values.

Action g_act[] =
{
	{_T(""), 0, 0}  // ACT_INVALID.

	// ASSIGNEXPR: Give it a name for Line::ToText().
	// 1st param is the target, 2nd (optional) is the value:
	, {_T(":="), 2, 2} // Same, though param #2 is flagged as numeric so that expression detection is automatic.

	// ACT_EXPRESSION, which is a stand-alone expression outside of any IF or assignment-command;
	// e.g. fn1(123, fn2(y)) or x&=3
	// Its name should be "" so that Line::ToText() will properly display it.
	, {_T(""), 1, 1}

	, {_T("{"), 0, 0}
	, {_T("}"), 0, 0}

	, {_T("Static"), 1, 1} // ACT_STATIC (used only at load time).
	, {_T("#HotIf"), 0, 1}
	, {_T("Exit"), 0, 1} // ExitCode

	, {_T("If"), 1, 1}
	, {_T("Else"), 0, 0} // No args; it has special handling to support same-line ELSE-actions (e.g. "else if").
	, {_T("Loop"), 0, 1} // IterationCount
	, {_T("Loop Files"), 1, 2} // FilePattern [, Mode] -- Files vs File for clarity.
	, {_T("Loop Reg"), 1, 2} // Key [, Mode]
	, {_T("Loop Read"), 1, 2} // InputFile [, OutputFile]
	, {_T("Loop Parse"), 1, 3} // InputString [, Delimiters, OmitChars]
	, {_T("For"), 0, MAX_ARGS}  // For vars in expression
	, {_T("While"), 1, 1} // LoopCondition.  v1.0.48: Lexikos: Added g_act entry for ACT_WHILE.
	, {_T("Until"), 1, 1} // Until expression (follows a Loop)
	, {_T("Break"), 0, 1}
	, {_T("Continue"), 0, 1}
	, {_T("Goto"), 1, 1}
	, {_T("Return"), 0, 1}
	, {_T("Try"), 0, 0}
	, {_T("Catch"), 0, 1} // fincs: seems best to allow catch without a parameter
	, {_T("Finally"), 0, 0}
	, {_T("Throw"), 0, 1}
	, {_T("Switch"), 0, 1}
	, {_T("Case"), 1, MAX_ARGS}
};
// Below is the most maintainable way to determine the actual count?
// Due to C++ lang. restrictions, can't easily make this a const because constants
// automatically get static (internal) linkage, thus such a var could never be
// used outside this module:
int g_ActionCount = _countof(g_act);


key_to_vk_type g_key_to_vk[] =
{ {_T("Numpad0"), VK_NUMPAD0}
, {_T("Numpad1"), VK_NUMPAD1}
, {_T("Numpad2"), VK_NUMPAD2}
, {_T("Numpad3"), VK_NUMPAD3}
, {_T("Numpad4"), VK_NUMPAD4}
, {_T("Numpad5"), VK_NUMPAD5}
, {_T("Numpad6"), VK_NUMPAD6}
, {_T("Numpad7"), VK_NUMPAD7}
, {_T("Numpad8"), VK_NUMPAD8}
, {_T("Numpad9"), VK_NUMPAD9}
, {_T("NumpadMult"), VK_MULTIPLY}
, {_T("NumpadDiv"), VK_DIVIDE}
, {_T("NumpadAdd"), VK_ADD}
, {_T("NumpadSub"), VK_SUBTRACT}
// , {_T("NumpadEnter"), VK_RETURN}  // Must do this one via scan code, see below for explanation.
, {_T("NumpadDot"), VK_DECIMAL}
, {_T("Numlock"), VK_NUMLOCK}
, {_T("ScrollLock"), VK_SCROLL}
, {_T("CapsLock"), VK_CAPITAL}

, {_T("Escape"), VK_ESCAPE}  // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {_T("Esc"), VK_ESCAPE}
, {_T("Tab"), VK_TAB}
, {_T("Space"), VK_SPACE}
, {_T("Backspace"), VK_BACK} // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {_T("BS"), VK_BACK}

// These keys each have a counterpart on the number pad with the same VK.  Use the VK for these,
// since they are probably more likely to be assigned to hotkeys (thus minimizing the use of the
// keyboard hook, and use the scan code (SC) for their counterparts.  UPDATE: To support handling
// these keys with the hook (i.e. the sc_takes_precedence flag in the hook), do them by scan code
// instead.  This allows Numpad keys such as Numpad7 to be differentiated from NumpadHome, which
// would otherwise be impossible since both of them share the same scan code (i.e. if the
// sc_takes_precedence flag is set for the scan code of NumpadHome, that will effectively prevent
// the hook from telling the difference between it and Numpad7 since the hook is currently set
// to handle an incoming key by either vk or sc, but not both.

// For VKs with multiple SCs, such as VK_RETURN, the keyboard hook is made mandatory unless the
// user specifies the VK by number.  This ensures that Enter:: and NumpadEnter::, for example,
// only fire when the appropriate key is pressed.
, {_T("Enter"), VK_RETURN}  // So that VKtoKeyName() delivers consistent results, always have the preferred name first.

// See g_key_to_sc for why these Numpad keys are handled here:
, {_T("NumpadDel"), VK_DELETE}
, {_T("NumpadIns"), VK_INSERT}
, {_T("NumpadClear"), VK_CLEAR}  // same physical key as Numpad5 on most keyboards?
, {_T("NumpadUp"), VK_UP}
, {_T("NumpadDown"), VK_DOWN}
, {_T("NumpadLeft"), VK_LEFT}
, {_T("NumpadRight"), VK_RIGHT}
, {_T("NumpadHome"), VK_HOME}
, {_T("NumpadEnd"), VK_END}
, {_T("NumpadPgUp"), VK_PRIOR}
, {_T("NumpadPgDn"), VK_NEXT}

, {_T("PrintScreen"), VK_SNAPSHOT}
, {_T("CtrlBreak"), VK_CANCEL}  // Might want to verify this, and whether it has any peculiarities.
, {_T("Pause"), VK_PAUSE} // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {_T("Help"), VK_HELP}  // VK_HELP is probably not the extended HELP key.  Not sure what this one is.
, {_T("Sleep"), VK_SLEEP}

, {_T("AppsKey"), VK_APPS}

, {_T("LControl"), VK_LCONTROL} // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {_T("RControl"), VK_RCONTROL} // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {_T("LCtrl"), VK_LCONTROL} // Abbreviated versions of the above.
, {_T("RCtrl"), VK_RCONTROL} //
, {_T("LShift"), VK_LSHIFT}
, {_T("RShift"), VK_RSHIFT}
, {_T("LAlt"), VK_LMENU}
, {_T("RAlt"), VK_RMENU}
// These two are always left/right centric and I think their vk's are always supported by the various
// Windows API calls, unlike VK_RSHIFT, etc. (which are seldom supported):
, {_T("LWin"), VK_LWIN}
, {_T("RWin"), VK_RWIN}

, {_T("Control"), VK_CONTROL} // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {_T("Ctrl"), VK_CONTROL}  // An alternate for convenience.
, {_T("Alt"), VK_MENU}
, {_T("Shift"), VK_SHIFT}

, {_T("F1"), VK_F1}
, {_T("F2"), VK_F2}
, {_T("F3"), VK_F3}
, {_T("F4"), VK_F4}
, {_T("F5"), VK_F5}
, {_T("F6"), VK_F6}
, {_T("F7"), VK_F7}
, {_T("F8"), VK_F8}
, {_T("F9"), VK_F9}
, {_T("F10"), VK_F10}
, {_T("F11"), VK_F11}
, {_T("F12"), VK_F12}
, {_T("F13"), VK_F13}
, {_T("F14"), VK_F14}
, {_T("F15"), VK_F15}
, {_T("F16"), VK_F16}
, {_T("F17"), VK_F17}
, {_T("F18"), VK_F18}
, {_T("F19"), VK_F19}
, {_T("F20"), VK_F20}
, {_T("F21"), VK_F21}
, {_T("F22"), VK_F22}
, {_T("F23"), VK_F23}
, {_T("F24"), VK_F24}

// Mouse buttons:
, {_T("LButton"), VK_LBUTTON}
, {_T("RButton"), VK_RBUTTON}
, {_T("MButton"), VK_MBUTTON}
, {_T("XButton1"), VK_XBUTTON1}
, {_T("XButton2"), VK_XBUTTON2}
// Custom/fake VKs for use by the mouse hook:
, {_T("WheelDown"), VK_WHEEL_DOWN}
, {_T("WheelUp"), VK_WHEEL_UP}
, {_T("WheelLeft"), VK_WHEEL_LEFT}
, {_T("WheelRight"), VK_WHEEL_RIGHT}

, {_T("Browser_Back"), VK_BROWSER_BACK}
, {_T("Browser_Forward"), VK_BROWSER_FORWARD}
, {_T("Browser_Refresh"), VK_BROWSER_REFRESH}
, {_T("Browser_Stop"), VK_BROWSER_STOP}
, {_T("Browser_Search"), VK_BROWSER_SEARCH}
, {_T("Browser_Favorites"), VK_BROWSER_FAVORITES}
, {_T("Browser_Home"), VK_BROWSER_HOME}
, {_T("Volume_Mute"), VK_VOLUME_MUTE}
, {_T("Volume_Down"), VK_VOLUME_DOWN}
, {_T("Volume_Up"), VK_VOLUME_UP}
, {_T("Media_Next"), VK_MEDIA_NEXT_TRACK}
, {_T("Media_Prev"), VK_MEDIA_PREV_TRACK}
, {_T("Media_Stop"), VK_MEDIA_STOP}
, {_T("Media_Play_Pause"), VK_MEDIA_PLAY_PAUSE}
, {_T("Launch_Mail"), VK_LAUNCH_MAIL}
, {_T("Launch_Media"), VK_LAUNCH_MEDIA_SELECT}
, {_T("Launch_App1"), VK_LAUNCH_APP1}
, {_T("Launch_App2"), VK_LAUNCH_APP2}

// Probably safest to terminate it this way, with a flag value.  (plus this makes it a little easier
// to code some loops, maybe).  Can also calculate how many elements are in the array using sizeof(array)
// divided by sizeof(element).  UPDATE: Decided not to do this in case ever decide to sort this array; don't
// want to rely on the fact that this will wind up in the right position after the sort (even though it
// should):
//, {_T(""), 0}
};



// For VKs with multiple SCs, one is handled by VK and one by SC.  Since MapVirtualKey() prefers
// non-extended keys, we treat extended keys as secondary and handle them by SC.  This simplifies
// the logic in vk_to_sc().  Another reason not to handle the Numlock-off Numpad keys here is that
// extra logic would be required to prioritize the Numlock-on VKs.  By contrast, handling the
// non-Numpad keys here requires a little extra logic in GetKeyName() to have it return the generic
// names when a generic VK (such as VK_END) is given without SC.
key_to_sc_type g_key_to_sc[] =
{ {_T("NumpadEnter"), SC_NUMPADENTER}

, {_T("Delete"), SC_DELETE}
, {_T("Del"), SC_DELETE}
, {_T("Insert"), SC_INSERT}
, {_T("Ins"), SC_INSERT}
// , {_T("Clear"), SC_CLEAR}  // Seems unnecessary because there is no counterpart to the Numpad5 clear key?
, {_T("Up"), SC_UP}
, {_T("Down"), SC_DOWN}
, {_T("Left"), SC_LEFT}
, {_T("Right"), SC_RIGHT}
, {_T("Home"), SC_HOME}
, {_T("End"), SC_END}
, {_T("PgUp"), SC_PGUP}
, {_T("PgDn"), SC_PGDN}
};


// Can calc the counts only after the arrays are initialized above:
int g_key_to_vk_count = _countof(g_key_to_vk);
int g_key_to_sc_count = _countof(g_key_to_sc);

KeyHistoryItem *g_KeyHistory = NULL; // Array is allocated during startup.
int g_KeyHistoryNext = 0;

// These must be global also, since both the keyboard and mouse hook functions,
// in addition to KeyEvent() when it's logging keys with only the mouse hook installed,
// MUST refer to the same variables.  Otherwise, the elapsed time between keyboard and
// and mouse events will be wrong:
DWORD g_HistoryTickNow = 0;
DWORD g_HistoryTickPrev = GetTickCount();  // So that the first logged key doesn't have a huge elapsed time.
HWND g_HistoryHwndPrev = NULL;

// Also hook related:
DWORD g_TimeLastInputPhysical = GetTickCount();
DWORD g_TimeLastInputKeyboard = g_TimeLastInputPhysical;
DWORD g_TimeLastInputMouse = g_TimeLastInputPhysical;
