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
#include "script.h" // For the global script object and g_ErrorLevel
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

UINT g_DefaultScriptCodepage = CP_ACP;

bool g_DestroyWindowCalled = false;
HWND g_hWnd = NULL;
HWND g_hWndEdit = NULL;
HWND g_hWndSplash = NULL;
HFONT g_hFontEdit = NULL;
HFONT g_hFontSplash = NULL;  // So that font can be deleted on program close.
HACCEL g_hAccelTable = NULL;

typedef int (WINAPI *StrCmpLogicalW_type)(LPCWSTR, LPCWSTR);
StrCmpLogicalW_type g_StrCmpLogicalW = NULL;
WNDPROC g_TabClassProc = NULL;

modLR_type g_modifiersLR_logical = 0;
modLR_type g_modifiersLR_logical_non_ignored = 0;
modLR_type g_modifiersLR_physical = 0;

#ifdef FUTURE_USE_MOUSE_BUTTONS_LOGICAL
WORD g_mouse_buttons_logical = 0;
#endif

// Used by the hook to track physical state of all virtual keys, since GetAsyncKeyState() does
// not retrieve the physical state of a key.  Note that this array is sometimes used in a way that
// requires its format to be the same as that returned from GetKeyboardState():
BYTE g_PhysicalKeyState[VK_ARRAY_COUNT] = {0};
bool g_BlockWinKeys = false;
DWORD g_HookReceiptOfLControlMeansAltGr = 0; // In these cases, zero is used as a false value, any others are true.
DWORD g_IgnoreNextLControlDown = 0;          //
DWORD g_IgnoreNextLControlUp = 0;            //

BYTE g_MenuMaskKey = VK_CONTROL; // L38: See #MenuMaskKey.

int g_HotkeyModifierTimeout = 50;  // Reduced from 100, which was a little too large for fast typists.
int g_ClipboardTimeout = 1000; // v1.0.31

HHOOK g_KeybdHook = NULL;
HHOOK g_MouseHook = NULL;
HHOOK g_PlaybackHook = NULL;
bool g_ForceLaunch = false;
bool g_WinActivateForce = false;
bool g_RunStdIn = false;
WarnMode g_Warn_UseUnsetLocal = WARNMODE_OFF;		// Used by #Warn directive.
WarnMode g_Warn_UseUnsetGlobal = WARNMODE_OFF;		//
WarnMode g_Warn_UseEnv = WARNMODE_OFF;				//
WarnMode g_Warn_LocalSameAsGlobal = WARNMODE_OFF;	//
SingleInstanceType g_AllowOnlyOneInstance = ALLOW_MULTI_INSTANCE;
bool g_persistent = false;  // Whether the script should stay running even after the auto-exec section finishes.
bool g_NoTrayIcon = false;
#ifdef AUTOHOTKEYSC
	bool g_AllowMainWindow = false;
#endif
bool g_MainTimerExists = false;
bool g_AutoExecTimerExists = false;
bool g_InputTimerExists = false;
bool g_DerefTimerExists = false;
bool g_SoundWasPlayed = false;
bool g_IsSuspended = false;  // Make this separate from g_AllowInterruption since that is frequently turned off & on.
bool g_DeferMessagesForUnderlyingPump = false;
BOOL g_WriteCacheDisabledInt64 = FALSE;  // BOOL vs. bool might improve performance a little for
BOOL g_WriteCacheDisabledDouble = FALSE; // frequently-accessed variables (it has helped performance in
BOOL g_NoEnv = FALSE;                    // ExpandExpression(), but didn't seem to help performance in g_NoEnv.
BOOL g_AllowInterruption = TRUE;         //
int g_nLayersNeedingTimer = 0;
int g_nThreads = 0;
int g_nPausedThreads = 0;
int g_MaxHistoryKeys = 40;

// g_MaxVarCapacity is used to prevent a buggy script from consuming all available system RAM. It is defined
// as the maximum memory size of a variable, including the string's zero terminator.
// The chosen default seems big enough to be flexible, yet small enough to not be a problem on 99% of systems:
VarSizeType g_MaxVarCapacity = 64 * 1024 * 1024;
UCHAR g_MaxThreadsPerHotkey = 1;
int g_MaxThreadsTotal = MAX_THREADS_DEFAULT;
// On my system, the repeat-rate (which is probably set to XP's default) is such that between 20
// and 25 keys are generated per second.  Therefore, 50 in 2000ms seems like it should allow the
// key auto-repeat feature to work on most systems without triggering the warning dialog.
// In any case, using auto-repeat with a hotkey is pretty rare for most people, so it's best
// to keep these values conservative:
int g_MaxHotkeysPerInterval = 70; // Increased to 70 because 60 was still causing the warning dialog for repeating keys sometimes.  Increased from 50 to 60 for v1.0.31.02 since 50 would be triggered by keyboard auto-repeat when it is set to its fastest.
int g_HotkeyThrottleInterval = 2000; // Milliseconds.
bool g_MaxThreadsBuffer = false;  // This feature usually does more harm than good, so it defaults to OFF.
SendLevelType g_InputLevel = 0;
HotCriterionType g_HotCriterion = HOT_NO_CRITERION;
LPTSTR g_HotWinTitle = _T(""); // In spite of the above being the primary indicator,
LPTSTR g_HotWinText = _T("");  // these are initialized for maintainability.
HotkeyCriterion *g_FirstHotCriterion = NULL, *g_LastHotCriterion = NULL;

// Global variables for #if (expression).
int g_HotExprIndex = -1; // The index of the Line containing the expression defined by the most recent #if (expression) directive.
Line **g_HotExprLines = NULL; // Array of pointers to expression lines, allocated when needed.
int g_HotExprLineCount = 0; // Number of expression lines currently present.
int g_HotExprLineCountMax = 0; // Current capacity of g_HotExprLines.
UINT g_HotExprTimeout = 1000; // Timeout for #if (expression) evaluation, in milliseconds.
HWND g_HotExprLFW = NULL; // Last Found Window of last #if expression.

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
int g_nInputBoxes = 0;
int g_nFileDialogs = 0;
int g_nFolderDialogs = 0;
InputBoxType g_InputBox[MAX_INPUTBOXES];
SplashType g_Progress[MAX_PROGRESS_WINDOWS] = {{0}};
SplashType g_SplashImage[MAX_SPLASHIMAGE_WINDOWS] = {{0}};
GuiType **g_gui = NULL;
int g_guiCount = 0, g_guiCountMax = 0;
HWND g_hWndToolTip[MAX_TOOLTIPS] = {NULL};
MsgMonitorStruct *g_MsgMonitor = NULL; // An array to be allocated upon first use (if any).
int g_MsgMonitorCount = 0;

// Init not needed for these:
UCHAR g_SortCaseSensitive;
bool g_SortNumeric;
bool g_SortReverse;
int g_SortColumnOffset;
Func *g_SortFunc;

TCHAR g_delimiter = ',';
TCHAR g_DerefChar = '%';
TCHAR g_EscapeChar = '`';

// Hot-string vars (initialized when ResetHook() is first called):
TCHAR g_HSBuf[HS_BUF_SIZE];
int g_HSBufLength;
HWND g_HShwnd;

// Hot-string global settings:
int g_HSPriority = 0;  // default priority is always 0
int g_HSKeyDelay = 0;  // Fast sends are much nicer for auto-replace and auto-backspace.
SendModes g_HSSendMode = SM_INPUT; // v1.0.43: New default for more reliable hotstrings.
bool g_HSCaseSensitive = false;
bool g_HSConformToCase = true;
bool g_HSDoBackspace = true;
bool g_HSOmitEndChar = false;
bool g_HSSendRaw = false;
bool g_HSEndCharRequired = true;
bool g_HSDetectWhenInsideWord = false;
bool g_HSDoReset = false;
bool g_HSResetUponMouseClick = true;
TCHAR g_EndChars[HS_MAX_END_CHARS + 1] = _T("-()[]{}:;'\"/\\,.?!\n \t");  // Hotstring default end chars, including a space.
// The following were considered but seemed too rare and/or too likely to result in undesirable replacements
// (such as while programming or scripting, or in usernames or passwords): <>*+=_%^&|@#$|
// Although dash/hyphen is used for multiple purposes, it seems to me that it is best (on average) to include it.
// Jay D. Novak suggested ([{/ for things such as fl/nj or fl(nj) which might resolve to USA state names.
// i.e. word(synonym) and/or word/synonym

// Global objects:
Var *g_ErrorLevel = NULL; // Allows us (in addition to the user) to set this var to indicate success/failure.
input_type g_input;
Script g_script;
// This made global for performance reasons (determining size of clipboard data then
// copying contents in or out without having to close & reopen the clipboard in between):
Clipboard g_clip;
OS_Version g_os;  // OS version object, courtesy of AutoIt3.

HICON g_IconSmall;
HICON g_IconLarge;

DWORD g_OriginalTimeout;

global_struct g_default, g_startup, *g_array;
global_struct *g = &g_startup; // g_startup provides a non-NULL placeholder during script loading. Afterward it's replaced with an array.

// I considered maintaining this on a per-quasi-thread basis (i.e. in global_struct), but the overhead
// of having to check and restore the working directory when a suspended thread is resumed (especially
// when the script has many high-frequency timers), and possibly changing the working directory
// whenever a new thread is launched, doesn't seem worth it.  This is because the need to change
// the working directory is comparatively rare:
TCHAR g_WorkingDir[MAX_PATH] = _T("");
TCHAR *g_WorkingDirOrig = NULL;  // Assigned a value in WinMain().

bool g_ContinuationLTrim = false;
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
// Note about the sub-array: Since the parent array is global, it would be automatically
// zero-filled if we didn't provide specific initialization.  But since we do, I'm not sure
// what value the unused elements in the NumericParams subarray will have.  Therefore, it seems
// safest to always terminate these subarrays with an explicit zero, below.

// STEPS TO ADD A NEW COMMAND:
// 1) Add an entry to the command enum in script.h.
// 2) Add an entry to the below array (it's position here MUST exactly match that in the enum).
//    The first item is the command name, the second is the minimum number of parameters (e.g.
//    if you enter 3, the first 3 args are mandatory) and the third is the maximum number of
//    parameters (the user need not escape commas within the last parameter).
//    The subarray should indicate the param numbers that must be numeric (first param is numbered 1,
//    not zero).  That subarray should be terminated with an explicit zero to be safe and
//    so that the compiler will complain if the sub-array size needs to be increased to
//    accommodate all the elements in the new sub-array, including room for its 0 terminator.
//    Note: If you use a value for MinParams than is greater than zero, remember than any params
//    beneath that threshold will also be required to be non-blank (i.e. user can't omit them even
//    if later, non-blank params are provided).  UPDATE: For a parameter to recognize an expression
//    such as x+100, it must be listed in the sub-array as a pure numeric parameter.
// 3) If the new command has any params that are output or input vars, change Line::ArgIsVar().
// 4) Add any desired load-time validation in Script::AddLine() in an syntax-checking section.
// 5) Implement the command in Line::Perform() or Line::EvaluateCondition (if it's an IF).
//    If the command waits for anything (e.g. calls MsgSleep()), be sure to make a local
//    copy of any ARG values that are needed during the wait period, because if another hotkey
//    subroutine suspends the current one while its waiting, it could also overwrite the ARG
//    deref buffer with its own values.

// v1.0.45 The following macro sets the high-bit for those commands that require overlap-checking of their
// input/output variables during runtime (commands that don't have an output variable never need this byte
// set, and runtime performance is improved even for them).  Some of commands are given the high-bit even
// though they might not strictly require it because rarity/performance/maintainability say it's best to do
// so when in doubt.  Search on "MaxParamsAu2WithHighBit" for more details.
#define H |(char)0x80

Action g_act[] =
{
	{_T(""), 0, 0, 0, NULL}  // ACT_INVALID.

	// ACT_ASSIGN, ACT_ADD/SUB/MULT/DIV: Give them names for display purposes.
	// Note: Line::ToText() relies on the below names being the correct symbols for the operation:
	// 1st param is the target, 2nd (optional) is the value:
	, {_T("="), 1, 2, 2 H, NULL}  // Omitting the second param sets the var to be empty. "H" (high-bit) is probably needed for those cases when PerformAssign() must call ExpandArgs() or similar.
	, {_T(":="), 1, 2, 2, {2, 0}} // Same, though param #2 is flagged as numeric so that expression detection is automatic.  "H" (high-bit) doesn't appear to be needed even when ACT_ASSIGNEXPR calls AssignBinaryClip() because that AssignBinaryClip() checks for source==dest.

	// ACT_EXPRESSION, which is a stand-alone expression outside of any IF or assignment-command;
	// e.g. fn1(123, fn2(y)) or x&=3
	// Its name should be "" so that Line::ToText() will properly display it.
	, {_T(""), 1, 1, 1, {1, 0}}

	, {_T("+="), 2, 3, 3, {2, 0}}
	, {_T("-="), 1, 3, 3, {2, 0}} // Subtraction (but not addition) allows 2nd to be blank due to 3rd param.
	, {_T("*="), 2, 2, 2, {2, 0}}
	, {_T("/="), 2, 2, 2, {2, 0}}

	, {_T("Else"), 0, 0, 0, NULL}

	, {_T("in"), 2, 2, 2, NULL}, {_T("not in"), 2, 2, 2, NULL}
	, {_T("contains"), 2, 2, 2, NULL}, {_T("not contains"), 2, 2, 2, NULL}  // Very similar to "in" and "not in"
	, {_T("is"), 2, 2, 2, NULL}, {_T("is not"), 2, 2, 2, NULL}
	, {_T("between"), 1, 3, 3, NULL}, {_T("not between"), 1, 3, 3, NULL}  // Min 1 to allow #2 and #3 to be the empty string.
	, {_T(""), 1, 1, 1, {1, 0}} // ACT_IFEXPR's name should be "" so that Line::ToText() will properly display it.

	// Comparison operators take 1 param (if they're being compared to blank) or 2.
	// For example, it's okay (though probably useless) to compare a string to the empty
	// string this way: "If var1 >=".  Note: Line::ToText() relies on the below names:
	, {_T("="), 1, 2, 2, NULL}, {_T("<>"), 1, 2, 2, NULL}, {_T(">"), 1, 2, 2, NULL}
	, {_T(">="), 1, 2, 2, NULL}, {_T("<"), 1, 2, 2, NULL}, {_T("<="), 1, 2, 2, NULL}

	// For these, allow a minimum of zero, otherwise, the first param (WinTitle) would
	// be considered mandatory-non-blank by default.  It's easier to make all the params
	// optional and validate elsewhere that at least one of the four isn't blank.
	// Also, All the IFs must be physically adjacent to each other in this array
	// so that ACT_FIRST_IF and ACT_LAST_IF can be used to detect if a command is an IF:
	, {_T("IfWinExist"), 0, 4, 4, NULL}, {_T("IfWinNotExist"), 0, 4, 4, NULL}  // Title, text, exclude-title, exclude-text
	// Passing zero params results in activating the LastUsed window:
	, {_T("IfWinActive"), 0, 4, 4, NULL}, {_T("IfWinNotActive"), 0, 4, 4, NULL} // same
	, {_T("IfInString"), 2, 2, 2, NULL} // String var, search string
	, {_T("IfNotInString"), 2, 2, 2, NULL} // String var, search string
	, {_T("IfExist"), 1, 1, 1, NULL} // File or directory.
	, {_T("IfNotExist"), 1, 1, 1, NULL} // File or directory.
	// IfMsgBox must be physically adjacent to the other IFs in this array:
	, {_T("IfMsgBox"), 1, 1, 1, NULL} // MsgBox result (e.g. OK, YES, NO)
	, {_T("MsgBox"), 0, 4, 3, NULL} // Text (if only 1 param) or: Mode-flag, Title, Text, Timeout.
	, {_T("InputBox"), 1, 11, 11 H, {5, 6, 7, 8, 10, 0}} // Output var, title, prompt, hide-text (e.g. passwords), width, height, X, Y, Font (e.g. courier:8 maybe), Timeout, Default
	, {_T("SplashTextOn"), 0, 4, 4, {1, 2, 0}} // Width, height, title, text
	, {_T("SplashTextOff"), 0, 0, 0, NULL}
	, {_T("Progress"), 0, 6, 6, NULL}  // Off|Percent|Options, SubText, MainText, Title, Font, FutureUse
	, {_T("SplashImage"), 0, 7, 7, NULL}  // Off|ImageFile, |Options, SubText, MainText, Title, Font, FutureUse
	, {_T("ToolTip"), 0, 4, 4, {2, 3, 4, 0}}  // Text, X, Y, ID.  If Text is omitted, the Tooltip is turned off.
	, {_T("TrayTip"), 0, 4, 4, {3, 4, 0}}  // Title, Text, Timeout, Options

	, {_T("Input"), 0, 4, 4 H, NULL}  // OutputVar, Options, EndKeys, MatchList.

	, {_T("Transform"), 2, 4, 4 H, NULL}  // output var, operation, value1, value2

	, {_T("StringLeft"), 3, 3, 3, {3, 0}}  // output var, input var, number of chars to extract
	, {_T("StringRight"), 3, 3, 3, {3, 0}} // same
	, {_T("StringMid"), 3, 5, 5, {3, 4, 0}} // Output Variable, Input Variable, Start char, Number of chars to extract, L
	, {_T("StringTrimLeft"), 3, 3, 3, {3, 0}}  // output var, input var, number of chars to trim
	, {_T("StringTrimRight"), 3, 3, 3, {3, 0}} // same
	, {_T("StringLower"), 2, 3, 3, NULL} // output var, input var, T = Title Case
	, {_T("StringUpper"), 2, 3, 3, NULL} // output var, input var, T = Title Case
	, {_T("StringLen"), 2, 2, 2, NULL} // output var, input var
	, {_T("StringGetPos"), 3, 5, 3, {5, 0}}  // Output Variable, Input Variable, Search Text, R or Right (from right), Offset
	, {_T("StringReplace"), 3, 5, 4, NULL} // Output Variable, Input Variable, Search String, Replace String, do-all.
	, {_T("StringSplit"), 2, 5, 5, NULL} // Output Array, Input Variable, Delimiter List (optional), Omit List, Future Use
	, {_T("SplitPath"), 1, 6, 6 H, NULL} // InputFilespec, OutName, OutDir, OutExt, OutNameNoExt, OutDrive
	, {_T("Sort"), 1, 2, 2, NULL} // OutputVar (it's also the input var), Options

	, {_T("EnvGet"), 2, 2, 2 H, NULL} // OutputVar, EnvVar
	, {_T("EnvSet"), 1, 2, 2, NULL} // EnvVar, Value
	, {_T("EnvUpdate"), 0, 0, 0, NULL}

	, {_T("RunAs"), 0, 3, 3, NULL} // user, pass, domain (0 params can be passed to disable the feature)
	, {_T("Run"), 1, 4, 4 H, NULL}      // TargetFile, Working Dir, WinShow-Mode/UseErrorLevel, OutputVarPID
	, {_T("RunWait"), 1, 4, 4 H, NULL}  // TargetFile, Working Dir, WinShow-Mode/UseErrorLevel, OutputVarPID
	, {_T("URLDownloadToFile"), 2, 2, 2, NULL} // URL, save-as-filename

	, {_T("GetKeyState"), 2, 3, 3 H, NULL} // OutputVar, key name, mode (optional) P = Physical, T = Toggle
	, {_T("Send"), 1, 1, 1, NULL}         // But that first param can validly be a deref that resolves to a blank param.
	, {_T("SendRaw"), 1, 1, 1, NULL}      //
	, {_T("SendInput"), 1, 1, 1, NULL}    //
	, {_T("SendPlay"), 1, 1, 1, NULL}     //
	, {_T("SendEvent"), 1, 1, 1, NULL}    // (due to rarity, there is no raw counterpart for this one)

	// For these, the "control" param can be blank.  The window's first visible control will
	// be used.  For this first one, allow a minimum of zero, otherwise, the first param (control)
	// would be considered mandatory-non-blank by default.  It's easier to make all the params
	// optional and validate elsewhere that the 2nd one specifically isn't blank:
	, {_T("ControlSend"), 0, 6, 6, NULL} // Control, Chars-to-Send, std. 4 window params.
	, {_T("ControlSendRaw"), 0, 6, 6, NULL} // Control, Chars-to-Send, std. 4 window params.
	, {_T("ControlClick"), 0, 8, 8, {5, 0}} // Control, WinTitle, WinText, WhichButton, ClickCount, Hold/Release, ExcludeTitle, ExcludeText
	, {_T("ControlMove"), 0, 9, 9, {2, 3, 4, 5, 0}} // Control, x, y, w, h, WinTitle, WinText, ExcludeTitle, ExcludeText
	, {_T("ControlGetPos"), 0, 9, 9 H, NULL} // Four optional output vars: xpos, ypos, width, height, control, std. 4 window params.
	, {_T("ControlFocus"), 0, 5, 5, NULL}     // Control, std. 4 window params
	, {_T("ControlGetFocus"), 1, 5, 5 H, NULL}  // OutputVar, std. 4 window params
	, {_T("ControlSetText"), 0, 6, 6, NULL}   // Control, new text, std. 4 window params
	, {_T("ControlGetText"), 1, 6, 6 H, NULL}   // Output-var, Control, std. 4 window params
	, {_T("Control"), 1, 7, 7, NULL}   // Command, Value, Control, std. 4 window params
	, {_T("ControlGet"), 2, 8, 8 H, NULL}   // Output-var, Command, Value, Control, std. 4 window params

	, {_T("SendMode"), 1, 1, 1, NULL}
	, {_T("SendLevel"), 1, 1, 1, {1, 0}}
	, {_T("CoordMode"), 1, 2, 2, NULL} // Attribute, screen|relative
	, {_T("SetDefaultMouseSpeed"), 1, 1, 1, {1, 0}} // speed (numeric)
	, {_T("Click"), 0, 1, 1, NULL} // Flex-list of options.
	, {_T("MouseMove"), 2, 4, 4, {1, 2, 3, 0}} // x, y, speed, option
	, {_T("MouseClick"), 0, 7, 7, {2, 3, 4, 5, 0}} // which-button, x, y, ClickCount, speed, d=hold-down/u=release, Relative
	, {_T("MouseClickDrag"), 1, 7, 7, {2, 3, 4, 5, 6, 0}} // which-button, x1, y1, x2, y2, speed, Relative
	, {_T("MouseGetPos"), 0, 5, 5 H, {5, 0}} // 4 optional output vars: xpos, ypos, WindowID, ControlName. Finally: Mode. MinParams must be 0.

	, {_T("StatusBarGetText"), 1, 6, 6 H, {2, 0}} // Output-var, part# (numeric), std. 4 window params
	, {_T("StatusBarWait"), 0, 8, 8, {2, 3, 6, 0}} // Wait-text(blank ok),seconds,part#,title,text,interval,exclude-title,exclude-text
	, {_T("ClipWait"), 0, 2, 2, {1, 2, 0}} // Seconds-to-wait (0 = 500ms), 1|0: Wait for any format, not just text/files
	, {_T("KeyWait"), 1, 2, 2, NULL} // KeyName, Options

	, {_T("Sleep"), 1, 1, 1, {1, 0}} // Sleep time in ms (numeric)
	, {_T("Random"), 0, 3, 3, {2, 3, 0}} // Output var, Min, Max (Note: MinParams is 1 so that param2 can be blank).

	, {_T("Goto"), 1, 1, 1, NULL}
	, {_T("Gosub"), 1, 1, 1, NULL}   // Label (or dereference that resolves to a label).
	, {_T("OnExit"), 0, 2, 2, NULL}  // Optional label, future use (since labels are allowed to contain commas)
	, {_T("Hotkey"), 1, 3, 3, NULL}  // Mod+Keys, Label/Action (blank to avoid changing curr. label), Options
	, {_T("SetTimer"), 0, 3, 3, {3, 0}}  // Label (or dereference that resolves to a label), period (or ON/OFF), Priority
	, {_T("Critical"), 0, 1, 1, NULL}  // On|Off
	, {_T("Thread"), 1, 3, 3, {2, 3, 0}}  // Command, value1 (can be blank for interrupt), value2
	, {_T("Return"), 0, 1, 1, {1, 0}}
	, {_T("Exit"), 0, 1, 1, {1, 0}} // ExitCode
	, {_T("Loop"), 0, 4, 4, NULL} // Iteration Count or FilePattern or root key name [,subkey name], FileLoopMode, Recurse? (custom validation for these last two)
	, {_T("For"), 1, 3, 3, {3, 0}}  // For var [,var] in expression
	, {_T("While"), 1, 1, 1, {1, 0}} // LoopCondition.  v1.0.48: Lexikos: Added g_act entry for ACT_WHILE.
	, {_T("Until"), 1, 1, 1, {1, 0}} // Until expression (follows a Loop)
	, {_T("Break"), 0, 1, 1, NULL}, {_T("Continue"), 0, 1, 1, NULL}
	, {_T("Try"), 0, 0, 0, NULL}
	, {_T("Catch"), 0, 1, 0, NULL} // fincs: seems best to allow catch without a parameter
	, {_T("Throw"), 0, 1, 1, {1, 0}}
	, {_T("Finally"), 0, 0, 0, NULL}
	, {_T("{"), 0, 0, 0, NULL}, {_T("}"), 0, 0, 0, NULL}

	, {_T("WinActivate"), 0, 4, 2, NULL} // Passing zero params results in activating the LastUsed window.
	, {_T("WinActivateBottom"), 0, 4, 4, NULL} // Min. 0 so that 1st params can be blank and later ones not blank.

	// These all use Title, Text, Timeout (in seconds not ms), Exclude-title, Exclude-text.
	// See above for why zero is the minimum number of params for each:
	, {_T("WinWait"), 0, 5, 5, {3, 0}}, {_T("WinWaitClose"), 0, 5, 5, {3, 0}}
	, {_T("WinWaitActive"), 0, 5, 5, {3, 0}}, {_T("WinWaitNotActive"), 0, 5, 5, {3, 0}}

	, {_T("WinMinimize"), 0, 4, 2, NULL}, {_T("WinMaximize"), 0, 4, 2, NULL}, {_T("WinRestore"), 0, 4, 2, NULL} // std. 4 params
	, {_T("WinHide"), 0, 4, 2, NULL}, {_T("WinShow"), 0, 4, 2, NULL} // std. 4 params
	, {_T("WinMinimizeAll"), 0, 0, 0, NULL}, {_T("WinMinimizeAllUndo"), 0, 0, 0, NULL}
	, {_T("WinClose"), 0, 5, 2, {3, 0}} // title, text, time-to-wait-for-close (0 = 500ms), exclude title/text
	, {_T("WinKill"), 0, 5, 2, {3, 0}} // same as WinClose.
	, {_T("WinMove"), 0, 8, 8, {1, 2, 3, 4, 5, 6, 0}} // title, text, xpos, ypos, width, height, exclude-title, exclude_text
	// Note for WinMove: title/text are marked as numeric because in two-param mode, they are the X/Y params.
	// This helps speed up loading expression-detection.  Also, xpos/ypos/width/height can be the string "default",
	// but that is explicitly checked for, even though it is required it to be numeric in the definition here.
	, {_T("WinMenuSelectItem"), 0, 11, 11, NULL} // WinTitle, WinText, Menu name, 6 optional sub-menu names, ExcludeTitle/Text

	, {_T("Process"), 1, 3, 3, NULL}  // Sub-cmd, PID/name, Param3 (use minimum of 1 param so that 2nd can be blank)

	, {_T("WinSet"), 1, 6, 6, NULL} // attribute, setting, title, text, exclude-title, exclude-text
	// WinSetTitle: Allow a minimum of zero params so that title isn't forced to be non-blank.
	// Also, if the user passes only one param, the title of the "last used" window will be
	// set to the string in the first param:
	, {_T("WinSetTitle"), 0, 5, 3, NULL} // title, text, newtitle, exclude-title, exclude-text
	, {_T("WinGetTitle"), 1, 5, 3 H, NULL} // Output-var, std. 4 window params
	, {_T("WinGetClass"), 1, 5, 5 H, NULL} // Output-var, std. 4 window params
	, {_T("WinGet"), 1, 6, 6 H, NULL} // Output-var/array, cmd (if omitted, defaults to ID), std. 4 window params
	, {_T("WinGetPos"), 0, 8, 8 H, NULL} // Four optional output vars: xpos, ypos, width, height.  Std. 4 window params.
	, {_T("WinGetText"), 1, 5, 5 H, NULL} // Output var, std 4 window params.

	, {_T("SysGet"), 2, 4, 4 H, NULL} // Output-var/array, sub-cmd or sys-metrics-number, input-value1, future-use

	, {_T("PostMessage"), 1, 8, 8, {1, 2, 3, 0}}  // msg, wParam, lParam, Control, WinTitle, WinText, ExcludeTitle, ExcludeText
	, {_T("SendMessage"), 1, 9, 9, {1, 2, 3, 9, 0}}  // msg, wParam, lParam, Control, WinTitle, WinText, ExcludeTitle, ExcludeText, Timeout

	, {_T("PixelGetColor"), 3, 4, 4 H, {2, 3, 0}} // OutputVar, X-coord, Y-coord [, RGB]
	, {_T("PixelSearch"), 0, 9, 9 H, {3, 4, 5, 6, 7, 8, 0}} // OutputX, OutputY, left, top, right, bottom, Color, Variation [, RGB]
	, {_T("ImageSearch"), 0, 7, 7 H, {3, 4, 5, 6, 0}} // OutputX, OutputY, left, top, right, bottom, ImageFile
	// NOTE FOR THE ABOVE: 0 min args so that the output vars can be optional.

	// See above for why minimum is 1 vs. 2:
	, {_T("GroupAdd"), 1, 6, 6, NULL} // Group name, WinTitle, WinText, Label, exclude-title/text
	, {_T("GroupActivate"), 1, 2, 2, NULL}
	, {_T("GroupDeactivate"), 1, 2, 2, NULL}
	, {_T("GroupClose"), 1, 2, 2, NULL}

	, {_T("DriveSpaceFree"), 2, 2, 2 H, NULL} // Output-var, path (e.g. c:\)
	, {_T("Drive"), 1, 3, 3, NULL} // Sub-command, Value1 (can be blank for Eject), Value2
	, {_T("DriveGet"), 0, 3, 3 H, NULL} // Output-var (optional in at least one case), Command, Value

	, {_T("SoundGet"), 1, 4, 4 H, {4, 0}} // OutputVar, ComponentType (default=master), ControlType (default=vol), Mixer/Device Number
	, {_T("SoundSet"), 1, 4, 4, {1, 4, 0}} // Volume percent-level (0-100), ComponentType, ControlType (default=vol), Mixer/Device Number
	, {_T("SoundGetWaveVolume"), 1, 2, 2 H, {2, 0}} // OutputVar, Mixer/Device Number
	, {_T("SoundSetWaveVolume"), 1, 2, 2, {1, 2, 0}} // Volume percent-level (0-100), Device Number (1 is the first)
	, {_T("SoundBeep"), 0, 2, 2, {1, 2, 0}} // Frequency, Duration.
	, {_T("SoundPlay"), 1, 2, 2, NULL} // Filename [, wait]

	, {_T("FileAppend"), 0, 3, 3, NULL} // text, filename (which can be omitted in a read-file loop). Update: Text can be omitted too, to create an empty file or alter the timestamp of an existing file.
	, {_T("FileRead"), 2, 2, 2 H, NULL} // Output variable, filename
	, {_T("FileReadLine"), 3, 3, 3 H, {3, 0}} // Output variable, filename, line-number
	, {_T("FileDelete"), 1, 1, 1, NULL} // filename or pattern
	, {_T("FileRecycle"), 1, 1, 1, NULL} // filename or pattern
	, {_T("FileRecycleEmpty"), 0, 1, 1, NULL} // optional drive letter (all bins will be emptied if absent.
	, {_T("FileInstall"), 2, 3, 3, {3, 0}} // source, dest, flag (1/0, where 1=overwrite)
	, {_T("FileCopy"), 2, 3, 3, {3, 0}} // source, dest, flag
	, {_T("FileMove"), 2, 3, 3, {3, 0}} // source, dest, flag
	, {_T("FileCopyDir"), 2, 3, 3, {3, 0}} // source, dest, flag
	, {_T("FileMoveDir"), 2, 3, 3, NULL} // source, dest, flag (which can be non-numeric in this case)
	, {_T("FileCreateDir"), 1, 1, 1, NULL} // dir name
	, {_T("FileRemoveDir"), 1, 2, 1, {2, 0}} // dir name, flag

	, {_T("FileGetAttrib"), 1, 2, 2 H, NULL} // OutputVar, Filespec (if blank, uses loop's current file)
	, {_T("FileSetAttrib"), 1, 4, 4, {3, 4, 0}} // Attribute(s), FilePattern, OperateOnFolders?, Recurse? (custom validation for these last two)
	, {_T("FileGetTime"), 1, 3, 3 H, NULL} // OutputVar, Filespec, WhichTime (modified/created/accessed)
	, {_T("FileSetTime"), 0, 5, 5, {1, 4, 5, 0}} // datetime (YYYYMMDDHH24MISS), FilePattern, WhichTime, OperateOnFolders?, Recurse?
	, {_T("FileGetSize"), 1, 3, 3 H, NULL} // OutputVar, Filespec, B|K|M (bytes, kb, or mb)
	, {_T("FileGetVersion"), 1, 2, 2 H, NULL} // OutputVar, Filespec

	, {_T("SetWorkingDir"), 1, 1, 1, NULL} // New path
	, {_T("FileSelectFile"), 1, 5, 3 H, NULL} // output var, options, working dir, greeting, filter
	, {_T("FileSelectFolder"), 1, 4, 4 H, {3, 0}} // output var, root directory, options, greeting

	, {_T("FileGetShortcut"), 1, 8, 8 H, NULL} // Filespec, OutTarget, OutDir, OutArg, OutDescrip, OutIcon, OutIconIndex, OutShowState.
	, {_T("FileCreateShortcut"), 2, 9, 9, {8, 9, 0}} // file, lnk [, workdir, args, desc, icon, hotkey, icon_number, run_state]

	, {_T("IniRead"), 2, 5, 4 H, NULL}   // OutputVar, Filespec, Section, Key, Default (value to return if key not found)
	, {_T("IniWrite"), 3, 4, 4, NULL}  // Value, Filespec, Section, Key
	, {_T("IniDelete"), 2, 3, 3, NULL} // Filespec, Section, Key

	// These require so few parameters due to registry loops, which provide the missing parameter values
	// automatically.  In addition, RegRead can't require more than 1 param since the 2nd param is
	// an option/obsolete parameter:
	, {_T("RegRead"), 1, 5, 5 H, NULL} // output var, (ValueType [optional]), RegKey, RegSubkey, ValueName
	, {_T("RegWrite"), 0, 5, 5, NULL} // ValueType, RegKey, RegSubKey, ValueName, Value (set to blank if omitted?)
	, {_T("RegDelete"), 0, 3, 3, NULL} // RegKey, RegSubKey, ValueName
	, {_T("SetRegView"), 1, 1, 1, NULL}

	, {_T("OutputDebug"), 1, 1, 1, NULL}

	, {_T("SetKeyDelay"), 0, 3, 3, {1, 2, 0}} // Delay in ms (numeric, negative allowed), PressDuration [, Play]
	, {_T("SetMouseDelay"), 1, 2, 2, {1, 0}} // Delay in ms (numeric, negative allowed) [, Play]
	, {_T("SetWinDelay"), 1, 1, 1, {1, 0}} // Delay in ms (numeric, negative allowed)
	, {_T("SetControlDelay"), 1, 1, 1, {1, 0}} // Delay in ms (numeric, negative allowed)
	, {_T("SetBatchLines"), 1, 1, 1, NULL} // Can be non-numeric, such as 15ms, or a number (to indicate line count).
	, {_T("SetTitleMatchMode"), 1, 1, 1, NULL} // Allowed values: 1, 2, slow, fast
	, {_T("SetFormat"), 2, 2, 2, NULL} // Float|Integer, FormatString (for float) or H|D (for int)
	, {_T("FormatTime"), 1, 3, 3 H, NULL} // OutputVar, YYYYMMDDHH24MISS, Format (format is last to avoid having to escape commas in it).

	, {_T("Suspend"), 0, 1, 1, NULL} // On/Off/Toggle/Permit/Blank (blank is the same as toggle)
	, {_T("Pause"), 0, 2, 2, NULL} // On/Off/Toggle/Blank (blank is the same as toggle), AlwaysAffectUnderlying
	, {_T("AutoTrim"), 1, 1, 1, NULL} // On/Off
	, {_T("StringCaseSense"), 1, 1, 1, NULL} // On/Off/Locale
	, {_T("DetectHiddenWindows"), 1, 1, 1, NULL} // On/Off
	, {_T("DetectHiddenText"), 1, 1, 1, NULL} // On/Off
	, {_T("BlockInput"), 1, 1, 1, NULL} // On/Off

	, {_T("SetNumlockState"), 0, 1, 1, NULL} // On/Off/AlwaysOn/AlwaysOff or blank (unspecified) to return to normal.
	, {_T("SetScrollLockState"), 0, 1, 1, NULL} // same
	, {_T("SetCapslockState"), 0, 1, 1, NULL} // same
	, {_T("SetStoreCapslockMode"), 1, 1, 1, NULL} // On/Off

	, {_T("KeyHistory"), 0, 2, 2, NULL}, {_T("ListLines"), 0, 1, 1, NULL}
	, {_T("ListVars"), 0, 0, 0, NULL}, {_T("ListHotkeys"), 0, 0, 0, NULL}

	, {_T("Edit"), 0, 0, 0, NULL}
	, {_T("Reload"), 0, 0, 0, NULL}
	, {_T("Menu"), 2, 6, 6, NULL}  // tray, add, name, label, options, future use
	, {_T("Gui"), 1, 4, 4, NULL}  // Cmd/Add, ControlType, Options, Text
	, {_T("GuiControl"), 0, 3, 3 H, NULL} // Sub-cmd (defaults to "contents"), ControlName/ID, Text
	, {_T("GuiControlGet"), 1, 4, 4, NULL} // OutputVar, Sub-cmd (defaults to "contents"), ControlName/ID (defaults to control assoc. with OutputVar), Text/FutureUse

	, {_T("ExitApp"), 0, 1, 1, {1, 0}}  // Optional exit-code. v1.0.48.01: Allow an expression like ACT_EXIT does.
	, {_T("Shutdown"), 1, 1, 1, {1, 0}} // Seems best to make the first param (the flag/code) mandatory.

	, {_T("FileEncoding"), 0, 1, 1, NULL}
};
// Below is the most maintainable way to determine the actual count?
// Due to C++ lang. restrictions, can't easily make this a const because constants
// automatically get static (internal) linkage, thus such a var could never be
// used outside this module:
int g_ActionCount = _countof(g_act);



Action g_old_act[] =
{
	{_T(""), 0, 0, 0, NULL}  // OLD_INVALID.
	, {_T("SetEnv"), 1, 2, 2, NULL}
	, {_T("EnvAdd"), 2, 3, 3, {2, 0}}, {_T("EnvSub"), 1, 3, 3, {2, 0}} // EnvSub (but not Add) allow 2nd to be blank due to 3rd param.
	, {_T("EnvMult"), 2, 2, 2, {2, 0}}, {_T("EnvDiv"), 2, 2, 2, {2, 0}}
	, {_T("IfEqual"), 1, 2, 2, NULL}, {_T("IfNotEqual"), 1, 2, 2, NULL}
	, {_T("IfGreater"), 1, 2, 2, NULL}, {_T("IfGreaterOrEqual"), 1, 2, 2, NULL}
	, {_T("IfLess"), 1, 2, 2, NULL}, {_T("IfLessOrEqual"), 1, 2, 2, NULL}
	, {_T("WinGetActiveTitle"), 1, 1, 1, NULL} // <Title Var>
	, {_T("WinGetActiveStats"), 5, 5, 5, NULL} // <Title Var>, <Width Var>, <Height Var>, <Xpos Var>, <Ypos Var>
};
int g_OldActionCount = _countof(g_old_act);


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

// Even though ENTER is probably less likely to be assigned than NumpadEnter, must have ENTER as
// the primary vk because otherwise, if the user configures only naked-NumPadEnter to do something,
// RegisterHotkey() would register that vk and ENTER would also be configured to do the same thing.
, {_T("Enter"), VK_RETURN}  // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {_T("Return"), VK_RETURN}

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
, {_T("Break"), VK_PAUSE} // Not really meaningful, but kept for as a synonym of Pause for backward compatibility.  See CtrlBreak.
, {_T("Help"), VK_HELP}  // VK_HELP is probably not the extended HELP key.  Not sure what this one is.
, {_T("Sleep"), VK_SLEEP}

, {_T("AppsKey"), VK_APPS}

// UPDATE: For the NT/2k/XP version, now doing these by VK since it's likely to be
// more compatible with non-standard or non-English keyboards:
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

// The left/right versions of these are handled elsewhere since their virtual keys aren't fully API-supported:
, {_T("Control"), VK_CONTROL} // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {_T("Ctrl"), VK_CONTROL}  // An alternate for convenience.
, {_T("Alt"), VK_MENU}
, {_T("Shift"), VK_SHIFT}
/*
These were used to confirm the fact that you can't use RegisterHotkey() on VK_LSHIFT, even if the shift
modifier is specified along with it:
, {_T("LShift"), VK_LSHIFT}
, {_T("RShift"), VK_RSHIFT}
*/
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
// Supported in only in Win2k and beyond:
, {_T("XButton1"), VK_XBUTTON1}
, {_T("XButton2"), VK_XBUTTON2}
// Custom/fake VKs for use by the mouse hook (supported only in WinNT SP3 and beyond?):
, {_T("WheelDown"), VK_WHEEL_DOWN}
, {_T("WheelUp"), VK_WHEEL_UP}
// Lexikos: Added fake VKs for support for horizontal scrolling in Windows Vista and later.
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



key_to_sc_type g_key_to_sc[] =
// Even though ENTER is probably less likely to be assigned than NumpadEnter, must have ENTER as
// the primary vk because otherwise, if the user configures only naked-NumPadEnter to do something,
// RegisterHotkey() would register that vk and ENTER would also be configured to do the same thing.
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

// If user specified left or right, must use scan code to distinguish *both* halves of the pair since
// each half has same vk *and* since their generic counterparts (e.g. CONTROL vs. L/RCONTROL) are
// already handled by vk.  Note: RWIN and LWIN don't need to be handled here because they each have
// their own virtual keys.
// UPDATE: For the NT/2k/XP version, now doing these by VK since it's likely to be
// more compatible with non-standard or non-English keyboards:
/*
, {_T("LControl"), SC_LCONTROL}
, {_T("RControl"), SC_RCONTROL}
, {_T("LShift"), SC_LSHIFT}
, {_T("RShift"), SC_RSHIFT}
, {_T("LAlt"), SC_LALT}
, {_T("RAlt"), SC_RALT}
*/
};


// Can calc the counts only after the arrays are initialized above:
int g_key_to_vk_count = _countof(g_key_to_vk);
int g_key_to_sc_count = _countof(g_key_to_sc);

KeyHistoryItem *g_KeyHistory = NULL; // Array is allocated during startup.
int g_KeyHistoryNext = 0;

#ifdef ENABLE_KEY_HISTORY_FILE
bool g_KeyHistoryToFile = false;
#endif

// These must be global also, since both the keyboard and mouse hook functions,
// in addition to KeyEvent() when it's logging keys with only the mouse hook installed,
// MUST refer to the same variables.  Otherwise, the elapsed time between keyboard and
// and mouse events will be wrong:
DWORD g_HistoryTickNow = 0;
DWORD g_HistoryTickPrev = GetTickCount();  // So that the first logged key doesn't have a huge elapsed time.
HWND g_HistoryHwndPrev = NULL;

// Also hook related:
DWORD g_TimeLastInputPhysical = GetTickCount();
