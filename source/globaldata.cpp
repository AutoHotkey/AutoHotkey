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

bool g_DestroyWindowCalled = false;
HWND g_hWnd = NULL;
HWND g_hWndEdit = NULL;
HWND g_hWndSplash = NULL;
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

int g_HotkeyModifierTimeout = 50;  // Reduced from 100, which was a little too large for fast typists.
int g_ClipboardTimeout = 1000; // v1.0.31

HHOOK g_KeybdHook = NULL;
HHOOK g_MouseHook = NULL;
HHOOK g_PlaybackHook = NULL;
bool g_ForceLaunch = false;
bool g_WinActivateForce = false;
SingleInstanceType g_AllowOnlyOneInstance = ALLOW_MULTI_INSTANCE;
bool g_WriteCacheDisabledInt64 = false;
bool g_WriteCacheDisabledDouble = false;
bool g_persistent = false;  // Whether the script should stay running even after the auto-exec section finishes.
bool g_NoEnv = false; // BOOL vs. bool didn't help performance in spite of the frequent accesses to it.
bool g_NoTrayIcon = false;
#ifdef AUTOHOTKEYSC
	bool g_AllowMainWindow = false;
#endif
bool g_AllowSameLineComments = true;
bool g_MainTimerExists = false;
bool g_UninterruptibleTimerExists = false;
bool g_AutoExecTimerExists = false;
bool g_InputTimerExists = false;
bool g_DerefTimerExists = false;
bool g_SoundWasPlayed = false;
bool g_IsSuspended = false;  // Make this separate from g_AllowInterruption since that is frequently turned off & on.
bool g_AllowInterruption = true;
bool g_DeferMessagesForUnderlyingPump = false;
int g_nLayersNeedingTimer = 0;
int g_nThreads = 0;
int g_nPausedThreads = 0;
bool g_IdleIsPaused = false;
int g_MaxHistoryKeys = 40;

// g_MaxVarCapacity is used to prevent a buggy script from consuming all available system RAM. It is defined
// as the maximum memory size of a variable, including the string's zero terminator.
// The chosen default seems big enough to be flexible, yet small enough to not be a problem on 99% of systems:
VarSizeType g_MaxVarCapacity = 64 * 1024 * 1024;
UCHAR g_MaxThreadsPerHotkey = 1;
int g_MaxThreadsTotal = 10;
// On my system, the repeat-rate (which is probably set to XP's default) is such that between 20
// and 25 keys are generated per second.  Therefore, 50 in 2000ms seems like it should allow the
// key auto-repeat feature to work on most systems without triggering the warning dialog.
// In any case, using auto-repeat with a hotkey is pretty rare for most people, so it's best
// to keep these values conservative:
int g_MaxHotkeysPerInterval = 70; // Increased to 70 because 60 was still causing the warning dialog for repeating keys sometimes.  Increased from 50 to 60 for v1.0.31.02 since 50 would be triggered by keyboard auto-repeat when it is set to its fastest.
int g_HotkeyThrottleInterval = 2000; // Milliseconds.
bool g_MaxThreadsBuffer = false;  // This feature usually does more harm than good, so it defaults to OFF.
HotCriterionType g_HotCriterion = HOT_NO_CRITERION;
char *g_HotWinTitle = ""; // In spite of the above being the primary indicator,
char *g_HotWinText = "";  // these are initialized for maintainability.
HotkeyCriterion *g_FirstHotCriterion = NULL, *g_LastHotCriterion = NULL;

// Lexikos: Added global variables for #if (expression).
int g_HotExprIndex = -1; // The index of the Line containing the expression defined by the most recent #if (expression) directive.
Line **g_HotExprLines = NULL; // Array of pointers to expression lines, allocated when needed.
int g_HotExprLineCount = 0; // Number of expression lines currently present.
int g_HotExprLineCountMax = 0; // Current capacity of g_HotExprLines.
UINT g_HotExprTimeout = 1000; // Timeout for #if (expression) evaluation, in milliseconds.

MenuTypeType g_MenuIsVisible = MENU_TYPE_NONE;
int g_nMessageBoxes = 0;
int g_nInputBoxes = 0;
int g_nFileDialogs = 0;
int g_nFolderDialogs = 0;
InputBoxType g_InputBox[MAX_INPUTBOXES];
SplashType g_Progress[MAX_PROGRESS_WINDOWS] = {{0}};
SplashType g_SplashImage[MAX_SPLASHIMAGE_WINDOWS] = {{0}};
GuiType *g_gui[MAX_GUI_WINDOWS] = {NULL};
HWND g_hWndToolTip[MAX_TOOLTIPS] = {NULL};
MsgMonitorStruct *g_MsgMonitor = NULL; // An array to be allocated upon first use (if any).
int g_MsgMonitorCount = 0;

// Init not needed for these:
UCHAR g_SortCaseSensitive;
bool g_SortNumeric;
bool g_SortReverse;
int g_SortColumnOffset;
Func *g_SortFunc;

char g_delimiter = ',';
char g_DerefChar = '%';
char g_EscapeChar = '`';

// Hot-string vars (initialized when ResetHook() is first called):
char g_HSBuf[HS_BUF_SIZE];
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
char g_EndChars[HS_MAX_END_CHARS + 1] = "-()[]{}:;'\"/\\,.?!\n \t";  // Hotstring default end chars, including a space.
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

// THIS MUST BE DONE AFTER the g_os object is initialized above:
// These are conditional because on these OSes, only standard-palette 16-color icons are supported,
// which would cause the normal icons to look mostly gray when used with in the tray.  So we use
// special 16x16x16 icons, but only for the tray because these OSes display the nicer icons okay
// in places other than the tray.  Also note that the red icons look okay, at least on Win98,
// because they are "red enough" not to suffer the graying effect from the palette shifting done
// by the OS:
int g_IconTray = (g_os.IsWinXPorLater() || g_os.IsWinMeorLater()) ? IDI_MAIN : IDI_TRAY_WIN9X;
int g_IconTraySuspend = (g_IconTray == IDI_MAIN) ? IDI_SUSPEND : IDI_TRAY_WIN9X_SUSPEND;

DWORD g_OriginalTimeout;

global_struct g, g_default;

// I considered maintaining this on a per-quasi-thread basis (i.e. in global_struct), but the overhead
// of having to check and restore the working directory when a suspended thread is resumed (especially
// when the script has many high-frequency timers), and possibly changing the working directory
// whenever a new thread is launched, doesn't seem worth it.  This is because the need to change
// the working directory is comparatively rare:
char g_WorkingDir[MAX_PATH] = "";
char *g_WorkingDirOrig = NULL;  // Assigned a value in WinMain().

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
// Note about the sub-array: Since the parent array array is global, it would be automatically
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
//    accommodate all the elements in the new sub-array, including room for it's 0 terminator.
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
// though they might not strictly require it because rarity/performance/maintainability say its best to do
// so when in doubt.  Search on "MaxParamsAu2WithHighBit" for more details.
#define H |(char)0x80

Action g_act[] =
{
	{"", 0, 0, 0, NULL}  // ACT_INVALID.

	// ACT_ASSIGN, ACT_ADD/SUB/MULT/DIV: Give them names for display purposes.
	// Note: Line::ToText() relies on the below names being the correct symbols for the operation:
	// 1st param is the target, 2nd (optional) is the value:
	, {"=", 1, 2, 2 H, NULL}  // Omitting the second param sets the var to be empty. "H" (high-bit) is probably needed for those cases when PerformAssign() must call ExpandArgs() or similar.
	, {":=", 1, 2, 2, {2, 0}} // Same, though param #2 is flagged as numeric so that expression detection is automatic.  "H" (high-bit) doesn't appear to be needed even when ACT_ASSIGNEXPR calls AssignBinaryClip() because that AssignBinaryClip() checks for source==dest.

	// ACT_EXPRESSION, which is a stand-alone expression outside of any IF or assignment-command;
	// e.g. fn1(123, fn2(y)) or x&=3
	// Its name should be "" so that Line::ToText() will properly display it.
	, {"", 1, 1, 1, {1, 0}}

	, {"+=", 2, 3, 3, {2, 0}}
	, {"-=", 1, 3, 3, {2, 0}} // Subtraction (but not addition) allows 2nd to be blank due to 3rd param.
	, {"*=", 2, 2, 2, {2, 0}}
	, {"/=", 2, 2, 2, {2, 0}}

	// This command is never directly parsed, but we need to have it here as a translation
	// target for the old "repeat" command.  This is because that command treats a zero
	// first-param as an infinite loop.  Since that param can be a dereferenced variable,
	// there's no way to reliably translate each REPEAT command into a LOOP command at
	// load-time.  Thus, we support both types of loops as actual commands that are
	// handled separately at runtime.
	, {"Repeat", 0, 1, 1, {1, 0}}  // Iteration Count: was mandatory in AutoIt2 but doesn't seem necessary here.
	, {"Else", 0, 0, 0, NULL}

	, {"in", 2, 2, 2, NULL}, {"not in", 2, 2, 2, NULL}
	, {"contains", 2, 2, 2, NULL}, {"not contains", 2, 2, 2, NULL}  // Very similar to "in" and "not in"
	, {"is", 2, 2, 2, NULL}, {"is not", 2, 2, 2, NULL}
	, {"between", 1, 3, 3, NULL}, {"not between", 1, 3, 3, NULL}  // Min 1 to allow #2 and #3 to be the empty string.
	, {"", 1, 1, 1, {1, 0}} // ACT_IFEXPR's name should be "" so that Line::ToText() will properly display it.

	// Comparison operators take 1 param (if they're being compared to blank) or 2.
	// For example, it's okay (though probably useless) to compare a string to the empty
	// string this way: "If var1 >=".  Note: Line::ToText() relies on the below names:
	, {"=", 1, 2, 2, NULL}, {"<>", 1, 2, 2, NULL}, {">", 1, 2, 2, NULL}
	, {">=", 1, 2, 2, NULL}, {"<", 1, 2, 2, NULL}, {"<=", 1, 2, 2, NULL}

	// For these, allow a minimum of zero, otherwise, the first param (WinTitle) would
	// be considered mandatory-non-blank by default.  It's easier to make all the params
	// optional and validate elsewhere that at least one of the four isn't blank.
	// Also, All the IFs must be physically adjacent to each other in this array
	// so that ACT_FIRST_IF and ACT_LAST_IF can be used to detect if a command is an IF:
	, {"IfWinExist", 0, 4, 4, NULL}, {"IfWinNotExist", 0, 4, 4, NULL}  // Title, text, exclude-title, exclude-text
	// Passing zero params results in activating the LastUsed window:
	, {"IfWinActive", 0, 4, 4, NULL}, {"IfWinNotActive", 0, 4, 4, NULL} // same
	, {"IfInString", 2, 2, 2, NULL} // String var, search string
	, {"IfNotInString", 2, 2, 2, NULL} // String var, search string
	, {"IfExist", 1, 1, 1, NULL} // File or directory.
	, {"IfNotExist", 1, 1, 1, NULL} // File or directory.
	// IfMsgBox must be physically adjacent to the other IFs in this array:
	, {"IfMsgBox", 1, 1, 1, NULL} // MsgBox result (e.g. OK, YES, NO)
	, {"MsgBox", 0, 4, 3, {4, 0}} // Text (if only 1 param) or: Mode-flag, Title, Text, Timeout.
	, {"InputBox", 1, 11, 11 H, {5, 6, 7, 8, 10, 0}} // Output var, title, prompt, hide-text (e.g. passwords), width, height, X, Y, Font (e.g. courier:8 maybe), Timeout, Default
	, {"SplashTextOn", 0, 4, 4, {1, 2, 0}} // Width, height, title, text
	, {"SplashTextOff", 0, 0, 0, NULL}
	, {"Progress", 0, 6, 6, NULL}  // Off|Percent|Options, SubText, MainText, Title, Font, FutureUse
	, {"SplashImage", 0, 7, 7, NULL}  // Off|ImageFile, |Options, SubText, MainText, Title, Font, FutureUse
	, {"ToolTip", 0, 4, 4, {2, 3, 4, 0}}  // Text, X, Y, ID.  If Text is omitted, the Tooltip is turned off.
	, {"TrayTip", 0, 4, 4, {3, 4, 0}}  // Title, Text, Timeout, Options

	, {"Input", 0, 4, 4 H, NULL}  // OutputVar, Options, EndKeys, MatchList.

	, {"Transform", 2, 4, 4 H, NULL}  // output var, operation, value1, value2

	, {"StringLeft", 3, 3, 3, {3, 0}}  // output var, input var, number of chars to extract
	, {"StringRight", 3, 3, 3, {3, 0}} // same
	, {"StringMid", 3, 5, 5, {3, 4, 0}} // Output Variable, Input Variable, Start char, Number of chars to extract, L
	, {"StringTrimLeft", 3, 3, 3, {3, 0}}  // output var, input var, number of chars to trim
	, {"StringTrimRight", 3, 3, 3, {3, 0}} // same
	, {"StringLower", 2, 3, 3, NULL} // output var, input var, T = Title Case
	, {"StringUpper", 2, 3, 3, NULL} // output var, input var, T = Title Case
	, {"StringLen", 2, 2, 2, NULL} // output var, input var
	, {"StringGetPos", 3, 5, 3, {5, 0}}  // Output Variable, Input Variable, Search Text, R or Right (from right), Offset
	, {"StringReplace", 3, 5, 4, NULL} // Output Variable, Input Variable, Search String, Replace String, do-all.
	, {"StringSplit", 2, 5, 5, NULL} // Output Array, Input Variable, Delimiter List (optional), Omit List, Future Use
	, {"SplitPath", 1, 6, 6 H, NULL} // InputFilespec, OutName, OutDir, OutExt, OutNameNoExt, OutDrive
	, {"Sort", 1, 2, 2, NULL} // OutputVar (it's also the input var), Options

	, {"EnvGet", 2, 2, 2 H, NULL} // OutputVar, EnvVar
	, {"EnvSet", 1, 2, 2, NULL} // EnvVar, Value
	, {"EnvUpdate", 0, 0, 0, NULL}

	, {"RunAs", 0, 3, 3, NULL} // user, pass, domain (0 params can be passed to disable the feature)
	, {"Run", 1, 4, 4 H, NULL}      // TargetFile, Working Dir, WinShow-Mode/UseErrorLevel, OutputVarPID
	, {"RunWait", 1, 4, 4 H, NULL}  // TargetFile, Working Dir, WinShow-Mode/UseErrorLevel, OutputVarPID
	, {"URLDownloadToFile", 2, 2, 2, NULL} // URL, save-as-filename

	, {"GetKeyState", 2, 3, 3 H, NULL} // OutputVar, key name, mode (optional) P = Physical, T = Toggle
	, {"Send", 1, 1, 1, NULL}         // But that first param can validly be a deref that resolves to a blank param.
	, {"SendRaw", 1, 1, 1, NULL}      //
	, {"SendInput", 1, 1, 1, NULL}    //
	, {"SendPlay", 1, 1, 1, NULL}     //
	, {"SendEvent", 1, 1, 1, NULL}    // (due to rarity, there is no raw counterpart for this one)

	// For these, the "control" param can be blank.  The window's first visible control will
	// be used.  For this first one, allow a minimum of zero, otherwise, the first param (control)
	// would be considered mandatory-non-blank by default.  It's easier to make all the params
	// optional and validate elsewhere that the 2nd one specifically isn't blank:
	, {"ControlSend", 0, 6, 6, NULL} // Control, Chars-to-Send, std. 4 window params.
	, {"ControlSendRaw", 0, 6, 6, NULL} // Control, Chars-to-Send, std. 4 window params.
	, {"ControlClick", 0, 8, 8, {5, 0}} // Control, WinTitle, WinText, WhichButton, ClickCount, Hold/Release, ExcludeTitle, ExcludeText
	, {"ControlMove", 0, 9, 9, {2, 3, 4, 5, 0}} // Control, x, y, w, h, WinTitle, WinText, ExcludeTitle, ExcludeText
	, {"ControlGetPos", 0, 9, 9 H, NULL} // Four optional output vars: xpos, ypos, width, height, control, std. 4 window params.
	, {"ControlFocus", 0, 5, 5, NULL}     // Control, std. 4 window params
	, {"ControlGetFocus", 1, 5, 5 H, NULL}  // OutputVar, std. 4 window params
	, {"ControlSetText", 0, 6, 6, NULL}   // Control, new text, std. 4 window params
	, {"ControlGetText", 1, 6, 6 H, NULL}   // Output-var, Control, std. 4 window params
	, {"Control", 1, 7, 7, NULL}   // Command, Value, Control, std. 4 window params
	, {"ControlGet", 2, 8, 8 H, NULL}   // Output-var, Command, Value, Control, std. 4 window params

	, {"SendMode", 1, 1, 1, NULL}
	, {"CoordMode", 1, 2, 2, NULL} // Attribute, screen|relative
	, {"SetDefaultMouseSpeed", 1, 1, 1, {1, 0}} // speed (numeric)
	, {"Click", 0, 1, 1, NULL} // Flex-list of options.
	, {"MouseMove", 2, 4, 4, {1, 2, 3, 0}} // x, y, speed, option
	, {"MouseClick", 0, 7, 7, {2, 3, 4, 5, 0}} // which-button, x, y, ClickCount, speed, d=hold-down/u=release, Relative
	, {"MouseClickDrag", 1, 7, 7, {2, 3, 4, 5, 6, 0}} // which-button, x1, y1, x2, y2, speed, Relative
	, {"MouseGetPos", 0, 5, 5 H, {5, 0}} // 4 optional output vars: xpos, ypos, WindowID, ControlName. Finally: Mode. MinParams must be 0.

	, {"StatusBarGetText", 1, 6, 6 H, {2, 0}} // Output-var, part# (numeric), std. 4 window params
	, {"StatusBarWait", 0, 8, 8, {2, 3, 6, 0}} // Wait-text(blank ok),seconds,part#,title,text,interval,exclude-title,exclude-text
	, {"ClipWait", 0, 2, 2, {1, 2, 0}} // Seconds-to-wait (0 = 500ms), 1|0: Wait for any format, not just text/files
	, {"KeyWait", 1, 2, 2, NULL} // KeyName, Options

	, {"Sleep", 1, 1, 1, {1, 0}} // Sleep time in ms (numeric)
	, {"Random", 0, 3, 3, {2, 3, 0}} // Output var, Min, Max (Note: MinParams is 1 so that param2 can be blank).

	, {"Goto", 1, 1, 1, NULL}
	, {"Gosub", 1, 1, 1, NULL}   // Label (or dereference that resolves to a label).
	, {"OnExit", 0, 2, 2, NULL}  // Optional label, future use (since labels are allowed to contain commas)
	, {"Hotkey", 1, 3, 3, NULL}  // Mod+Keys, Label/Action (blank to avoid changing curr. label), Options
	, {"SetTimer", 1, 3, 3, {3, 0}}  // Label (or dereference that resolves to a label), period (or ON/OFF), Priority
	, {"Critical", 0, 1, 1, NULL}  // On|Off
	, {"Thread", 1, 3, 3, {2, 3, 0}}  // Command, value1 (can be blank for interrupt), value2
	, {"Return", 0, 1, 1, {1, 0}}
	, {"Exit", 0, 1, 1, {1, 0}} // ExitCode
	, {"Loop", 0, 4, 4, NULL} // Iteration Count or FilePattern or root key name [,subkey name], FileLoopMode, Recurse? (custom validation for these last two)
	, {"While", 1, 1, 1, {1, 0}} // LoopCondition	// Lexikos: Added g_act entry for ACT_WHILE.
	, {"Break", 0, 0, 0, NULL}, {"Continue", 0, 0, 0, NULL}
	, {"{", 0, 0, 0, NULL}, {"}", 0, 0, 0, NULL}

	, {"WinActivate", 0, 4, 2, NULL} // Passing zero params results in activating the LastUsed window.
	, {"WinActivateBottom", 0, 4, 4, NULL} // Min. 0 so that 1st params can be blank and later ones not blank.

	// These all use Title, Text, Timeout (in seconds not ms), Exclude-title, Exclude-text.
	// See above for why zero is the minimum number of params for each:
	, {"WinWait", 0, 5, 5, {3, 0}}, {"WinWaitClose", 0, 5, 5, {3, 0}}
	, {"WinWaitActive", 0, 5, 5, {3, 0}}, {"WinWaitNotActive", 0, 5, 5, {3, 0}}

	, {"WinMinimize", 0, 4, 2, NULL}, {"WinMaximize", 0, 4, 2, NULL}, {"WinRestore", 0, 4, 2, NULL} // std. 4 params
	, {"WinHide", 0, 4, 2, NULL}, {"WinShow", 0, 4, 2, NULL} // std. 4 params
	, {"WinMinimizeAll", 0, 0, 0, NULL}, {"WinMinimizeAllUndo", 0, 0, 0, NULL}
	, {"WinClose", 0, 5, 2, {3, 0}} // title, text, time-to-wait-for-close (0 = 500ms), exclude title/text
	, {"WinKill", 0, 5, 2, {3, 0}} // same as WinClose.
	, {"WinMove", 0, 8, 8, {1, 2, 3, 4, 5, 6, 0}} // title, text, xpos, ypos, width, height, exclude-title, exclude_text
	// Note for WinMove: title/text are marked as numeric because in two-param mode, they are the X/Y params.
	// This helps speed up loading expression-detection.  Also, xpos/ypos/width/height can be the string "default",
	// but that is explicitly checked for, even though it is required it to be numeric in the definition here.
	, {"WinMenuSelectItem", 0, 11, 11, NULL} // WinTitle, WinText, Menu name, 6 optional sub-menu names, ExcludeTitle/Text

	, {"Process", 1, 3, 3, NULL}  // Sub-cmd, PID/name, Param3 (use minimum of 1 param so that 2nd can be blank)

	, {"WinSet", 1, 6, 6, NULL} // attribute, setting, title, text, exclude-title, exclude-text
	// WinSetTitle: Allow a minimum of zero params so that title isn't forced to be non-blank.
	// Also, if the user passes only one param, the title of the "last used" window will be
	// set to the string in the first param:
	, {"WinSetTitle", 0, 5, 3, NULL} // title, text, newtitle, exclude-title, exclude-text
	, {"WinGetTitle", 1, 5, 3 H, NULL} // Output-var, std. 4 window params
	, {"WinGetClass", 1, 5, 5 H, NULL} // Output-var, std. 4 window params
	, {"WinGet", 1, 6, 6 H, NULL} // Output-var/array, cmd (if omitted, defaults to ID), std. 4 window params
	, {"WinGetPos", 0, 8, 8 H, NULL} // Four optional output vars: xpos, ypos, width, height.  Std. 4 window params.
	, {"WinGetText", 1, 5, 5 H, NULL} // Output var, std 4 window params.

	, {"SysGet", 2, 4, 4 H, NULL} // Output-var/array, sub-cmd or sys-metrics-number, input-value1, future-use

	, {"PostMessage", 1, 8, 8, {1, 2, 3, 0}}  // msg, wParam, lParam, Control, WinTitle, WinText, ExcludeTitle, ExcludeText
	, {"SendMessage", 1, 8, 8, {1, 2, 3, 0}}  // msg, wParam, lParam, Control, WinTitle, WinText, ExcludeTitle, ExcludeText

	, {"PixelGetColor", 3, 4, 4 H, {2, 3, 0}} // OutputVar, X-coord, Y-coord [, RGB]
	, {"PixelSearch", 0, 9, 9 H, {3, 4, 5, 6, 7, 8, 0}} // OutputX, OutputY, left, top, right, bottom, Color, Variation [, RGB]
	, {"ImageSearch", 0, 7, 7 H, {3, 4, 5, 6, 0}} // OutputX, OutputY, left, top, right, bottom, ImageFile
	// NOTE FOR THE ABOVE: 0 min args so that the output vars can be optional.

	// See above for why minimum is 1 vs. 2:
	, {"GroupAdd", 1, 6, 6, NULL} // Group name, WinTitle, WinText, Label, exclude-title/text
	, {"GroupActivate", 1, 2, 2, NULL}
	, {"GroupDeactivate", 1, 2, 2, NULL}
	, {"GroupClose", 1, 2, 2, NULL}

	, {"DriveSpaceFree", 2, 2, 2 H, NULL} // Output-var, path (e.g. c:\)
	, {"Drive", 1, 3, 3, NULL} // Sub-command, Value1 (can be blank for Eject), Value2
	, {"DriveGet", 0, 3, 3 H, NULL} // Output-var (optional in at least one case), Command, Value

	, {"SoundGet", 1, 4, 4 H, {4, 0}} // OutputVar, ComponentType (default=master), ControlType (default=vol), Mixer/Device Number
	, {"SoundSet", 1, 4, 4, {1, 4, 0}} // Volume percent-level (0-100), ComponentType, ControlType (default=vol), Mixer/Device Number
	, {"SoundGetWaveVolume", 1, 2, 2 H, {2, 0}} // OutputVar, Mixer/Device Number
	, {"SoundSetWaveVolume", 1, 2, 2, {1, 2, 0}} // Volume percent-level (0-100), Device Number (1 is the first)
	, {"SoundBeep", 0, 2, 2, {1, 2, 0}} // Frequency, Duration.
	, {"SoundPlay", 1, 2, 2, NULL} // Filename [, wait]

	, {"FileAppend", 0, 2, 2, NULL} // text, filename (which can be omitted in a read-file loop). Update: Text can be omitted too, to create an empty file or alter the timestamp of an existing file.
	, {"FileRead", 2, 2, 2 H, NULL} // Output variable, filename
	, {"FileReadLine", 3, 3, 3 H, {3, 0}} // Output variable, filename, line-number
	, {"FileDelete", 1, 1, 1, NULL} // filename or pattern
	, {"FileRecycle", 1, 1, 1, NULL} // filename or pattern
	, {"FileRecycleEmpty", 0, 1, 1, NULL} // optional drive letter (all bins will be emptied if absent.
	, {"FileInstall", 2, 3, 3, {3, 0}} // source, dest, flag (1/0, where 1=overwrite)
	, {"FileCopy", 2, 3, 3, {3, 0}} // source, dest, flag
	, {"FileMove", 2, 3, 3, {3, 0}} // source, dest, flag
	, {"FileCopyDir", 2, 3, 3, {3, 0}} // source, dest, flag
	, {"FileMoveDir", 2, 3, 3, NULL} // source, dest, flag (which can be non-numeric in this case)
	, {"FileCreateDir", 1, 1, 1, NULL} // dir name
	, {"FileRemoveDir", 1, 2, 1, {2, 0}} // dir name, flag

	, {"FileGetAttrib", 1, 2, 2 H, NULL} // OutputVar, Filespec (if blank, uses loop's current file)
	, {"FileSetAttrib", 1, 4, 4, {3, 4, 0}} // Attribute(s), FilePattern, OperateOnFolders?, Recurse? (custom validation for these last two)
	, {"FileGetTime", 1, 3, 3 H, NULL} // OutputVar, Filespec, WhichTime (modified/created/accessed)
	, {"FileSetTime", 0, 5, 5, {1, 4, 5, 0}} // datetime (YYYYMMDDHH24MISS), FilePattern, WhichTime, OperateOnFolders?, Recurse?
	, {"FileGetSize", 1, 3, 3 H, NULL} // OutputVar, Filespec, B|K|M (bytes, kb, or mb)
	, {"FileGetVersion", 1, 2, 2 H, NULL} // OutputVar, Filespec

	, {"SetWorkingDir", 1, 1, 1, NULL} // New path
	, {"FileSelectFile", 1, 5, 3 H, NULL} // output var, options, working dir, greeting, filter
	, {"FileSelectFolder", 1, 4, 4 H, {3, 0}} // output var, root directory, options, greeting

	, {"FileGetShortcut", 1, 8, 8 H, NULL} // Filespec, OutTarget, OutDir, OutArg, OutDescrip, OutIcon, OutIconIndex, OutShowState.
	, {"FileCreateShortcut", 2, 9, 9, {8, 9, 0}} // file, lnk [, workdir, args, desc, icon, hotkey, icon_number, run_state]

	, {"IniRead", 4, 5, 4 H, NULL}   // OutputVar, Filespec, Section, Key, Default (value to return if key not found)
	, {"IniWrite", 4, 4, 4, NULL}  // Value, Filespec, Section, Key
	, {"IniDelete", 2, 3, 3, NULL} // Filespec, Section, Key

	// These require so few parameters due to registry loops, which provide the missing parameter values
	// automatically.  In addition, RegRead can't require more than 1 param since the 2nd param is
	// an option/obsolete parameter:
	, {"RegRead", 1, 5, 5 H, NULL} // output var, (ValueType [optional]), RegKey, RegSubkey, ValueName
	, {"RegWrite", 0, 5, 5, NULL} // ValueType, RegKey, RegSubKey, ValueName, Value (set to blank if omitted?)
	, {"RegDelete", 0, 3, 3, NULL} // RegKey, RegSubKey, ValueName

	, {"OutputDebug", 1, 1, 1, NULL}

	, {"SetKeyDelay", 0, 3, 3, {1, 2, 0}} // Delay in ms (numeric, negative allowed), PressDuration [, Play]
	, {"SetMouseDelay", 1, 2, 2, {1, 0}} // Delay in ms (numeric, negative allowed) [, Play]
	, {"SetWinDelay", 1, 1, 1, {1, 0}} // Delay in ms (numeric, negative allowed)
	, {"SetControlDelay", 1, 1, 1, {1, 0}} // Delay in ms (numeric, negative allowed)
	, {"SetBatchLines", 1, 1, 1, NULL} // Can be non-numeric, such as 15ms, or a number (to indicate line count).
	, {"SetTitleMatchMode", 1, 1, 1, NULL} // Allowed values: 1, 2, slow, fast
	, {"SetFormat", 2, 2, 2, NULL} // Float|Integer, FormatString (for float) or H|D (for int)
	, {"FormatTime", 1, 3, 3 H, NULL} // OutputVar, YYYYMMDDHH24MISS, Format (format is last to avoid having to escape commas in it).

	, {"Suspend", 0, 1, 1, NULL} // On/Off/Toggle/Permit/Blank (blank is the same as toggle)
	, {"Pause", 0, 2, 2, NULL} // On/Off/Toggle/Blank (blank is the same as toggle), AlwaysAffectUnderlying
	, {"AutoTrim", 1, 1, 1, NULL} // On/Off
	, {"StringCaseSense", 1, 1, 1, NULL} // On/Off/Locale
	, {"DetectHiddenWindows", 1, 1, 1, NULL} // On/Off
	, {"DetectHiddenText", 1, 1, 1, NULL} // On/Off
	, {"BlockInput", 1, 1, 1, NULL} // On/Off

	, {"SetNumlockState", 0, 1, 1, NULL} // On/Off/AlwaysOn/AlwaysOff or blank (unspecified) to return to normal.
	, {"SetScrollLockState", 0, 1, 1, NULL} // same
	, {"SetCapslockState", 0, 1, 1, NULL} // same
	, {"SetStoreCapslockMode", 1, 1, 1, NULL} // On/Off

	, {"KeyHistory", 0, 2, 2, NULL}, {"ListLines", 0, 0, 0, NULL}
	, {"ListVars", 0, 0, 0, NULL}, {"ListHotkeys", 0, 0, 0, NULL}

	, {"Edit", 0, 0, 0, NULL}
	, {"Reload", 0, 0, 0, NULL}
	, {"Menu", 2, 6, 6, NULL}  // tray, add, name, label, options, future use
	, {"Gui", 1, 4, 4, NULL}  // Cmd/Add, ControlType, Options, Text
	, {"GuiControl", 0, 3, 3 H, NULL} // Sub-cmd (defaults to "contents"), ControlName/ID, Text
	, {"GuiControlGet", 1, 4, 4, NULL} // OutputVar, Sub-cmd (defaults to "contents"), ControlName/ID (defaults to control assoc. with OutputVar), Text/FutureUse

	, {"ExitApp", 0, 1, 1, NULL}  // Optional exit-code
	, {"Shutdown", 1, 1, 1, {1, 0}} // Seems best to make the first param (the flag/code) mandatory.
};
// Below is the most maintainable way to determine the actual count?
// Due to C++ lang. restrictions, can't easily make this a const because constants
// automatically get static (internal) linkage, thus such a var could never be
// used outside this module:
int g_ActionCount = sizeof(g_act) / sizeof(Action);



Action g_old_act[] =
{
	{"", 0, 0, 0, NULL}  // OLD_INVALID.
	, {"SetEnv", 1, 2, 2, NULL}
	, {"EnvAdd", 2, 3, 3, {2, 0}}, {"EnvSub", 1, 3, 3, {2, 0}} // EnvSub (but not Add) allow 2nd to be blank due to 3rd param.
	, {"EnvMult", 2, 2, 2, {2, 0}}, {"EnvDiv", 2, 2, 2, {2, 0}}
	, {"IfEqual", 1, 2, 2, NULL}, {"IfNotEqual", 1, 2, 2, NULL}
	, {"IfGreater", 1, 2, 2, NULL}, {"IfGreaterOrEqual", 1, 2, 2, NULL}
	, {"IfLess", 1, 2, 2, NULL}, {"IfLessOrEqual", 1, 2, 2, NULL}
	, {"LeftClick", 2, 2, 2, {1, 2, 0}}, {"RightClick", 2, 2, 2, {1, 2, 0}}
	, {"LeftClickDrag", 4, 4, 4, {1, 2, 3, 4, 0}}, {"RightClickDrag", 4, 4, 4, {1, 2, 3, 4, 0}}
	, {"HideAutoItWin", 1, 1, 1, NULL}
	  // Allow zero params, unlike AutoIt.  These params should match those for REPEAT in the above array:
	, {"Repeat", 0, 1, 1, {1, 0}}, {"EndRepeat", 0, 0, 0, NULL}
	, {"WinGetActiveTitle", 1, 1, 1, NULL} // <Title Var>
	, {"WinGetActiveStats", 5, 5, 5, NULL} // <Title Var>, <Width Var>, <Height Var>, <Xpos Var>, <Ypos Var>
};
int g_OldActionCount = sizeof(g_old_act) / sizeof(Action);


key_to_vk_type g_key_to_vk[] =
{ {"Numpad0", VK_NUMPAD0}
, {"Numpad1", VK_NUMPAD1}
, {"Numpad2", VK_NUMPAD2}
, {"Numpad3", VK_NUMPAD3}
, {"Numpad4", VK_NUMPAD4}
, {"Numpad5", VK_NUMPAD5}
, {"Numpad6", VK_NUMPAD6}
, {"Numpad7", VK_NUMPAD7}
, {"Numpad8", VK_NUMPAD8}
, {"Numpad9", VK_NUMPAD9}
, {"NumpadMult", VK_MULTIPLY}
, {"NumpadDiv", VK_DIVIDE}
, {"NumpadAdd", VK_ADD}
, {"NumpadSub", VK_SUBTRACT}
// , {"NumpadEnter", VK_RETURN}  // Must do this one via scan code, see below for explanation.
, {"NumpadDot", VK_DECIMAL}
, {"Numlock", VK_NUMLOCK}
, {"ScrollLock", VK_SCROLL}
, {"CapsLock", VK_CAPITAL}

, {"Escape", VK_ESCAPE}  // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {"Esc", VK_ESCAPE}
, {"Tab", VK_TAB}
, {"Space", VK_SPACE}
, {"Backspace", VK_BACK} // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {"BS", VK_BACK}

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
, {"Enter", VK_RETURN}  // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {"Return", VK_RETURN}

, {"NumpadDel", VK_DELETE}
, {"NumpadIns", VK_INSERT}
, {"NumpadClear", VK_CLEAR}  // same physical key as Numpad5 on most keyboards?
, {"NumpadUp", VK_UP}
, {"NumpadDown", VK_DOWN}
, {"NumpadLeft", VK_LEFT}
, {"NumpadRight", VK_RIGHT}
, {"NumpadHome", VK_HOME}
, {"NumpadEnd", VK_END}
, {"NumpadPgUp", VK_PRIOR}
, {"NumpadPgDn", VK_NEXT}

, {"PrintScreen", VK_SNAPSHOT}
, {"CtrlBreak", VK_CANCEL}  // Might want to verify this, and whether it has any peculiarities.
, {"Pause", VK_PAUSE} // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {"Break", VK_PAUSE} // Not really meaningful, but kept for as a synonym of Pause for backward compatibility.  See CtrlBreak.
, {"Help", VK_HELP}  // VK_HELP is probably not the extended HELP key.  Not sure what this one is.
, {"Sleep", VK_SLEEP}

, {"AppsKey", VK_APPS}

// UPDATE: For the NT/2k/XP version, now doing these by VK since it's likely to be
// more compatible with non-standard or non-English keyboards:
, {"LControl", VK_LCONTROL} // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {"RControl", VK_RCONTROL} // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {"LCtrl", VK_LCONTROL} // Abbreviated versions of the above.
, {"RCtrl", VK_RCONTROL} //
, {"LShift", VK_LSHIFT}
, {"RShift", VK_RSHIFT}
, {"LAlt", VK_LMENU}
, {"RAlt", VK_RMENU}
// These two are always left/right centric and I think their vk's are always supported by the various
// Windows API calls, unlike VK_RSHIFT, etc. (which are seldom supported):
, {"LWin", VK_LWIN}
, {"RWin", VK_RWIN}

// The left/right versions of these are handled elsewhere since their virtual keys aren't fully API-supported:
, {"Control", VK_CONTROL} // So that VKtoKeyName() delivers consistent results, always have the preferred name first.
, {"Ctrl", VK_CONTROL}  // An alternate for convenience.
, {"Alt", VK_MENU}
, {"Shift", VK_SHIFT}
/*
These were used to confirm the fact that you can't use RegisterHotkey() on VK_LSHIFT, even if the shift
modifier is specified along with it:
, {"LShift", VK_LSHIFT}
, {"RShift", VK_RSHIFT}
*/
, {"F1", VK_F1}
, {"F2", VK_F2}
, {"F3", VK_F3}
, {"F4", VK_F4}
, {"F5", VK_F5}
, {"F6", VK_F6}
, {"F7", VK_F7}
, {"F8", VK_F8}
, {"F9", VK_F9}
, {"F10", VK_F10}
, {"F11", VK_F11}
, {"F12", VK_F12}
, {"F13", VK_F13}
, {"F14", VK_F14}
, {"F15", VK_F15}
, {"F16", VK_F16}
, {"F17", VK_F17}
, {"F18", VK_F18}
, {"F19", VK_F19}
, {"F20", VK_F20}
, {"F21", VK_F21}
, {"F22", VK_F22}
, {"F23", VK_F23}
, {"F24", VK_F24}

// Mouse buttons:
, {"LButton", VK_LBUTTON}
, {"RButton", VK_RBUTTON}
, {"MButton", VK_MBUTTON}
// Supported in only in Win2k and beyond:
, {"XButton1", VK_XBUTTON1}
, {"XButton2", VK_XBUTTON2}
// Custom/fake VKs for use by the mouse hook (supported only in WinNT SP3 and beyond?):
, {"WheelDown", VK_WHEEL_DOWN}
, {"WheelUp", VK_WHEEL_UP}
// Lexikos: Added fake VKs for support for horizontal scrolling in Windows Vista and later.
, {"WheelLeft", VK_WHEEL_LEFT}
, {"WheelRight", VK_WHEEL_RIGHT}

, {"Browser_Back", VK_BROWSER_BACK}
, {"Browser_Forward", VK_BROWSER_FORWARD}
, {"Browser_Refresh", VK_BROWSER_REFRESH}
, {"Browser_Stop", VK_BROWSER_STOP}
, {"Browser_Search", VK_BROWSER_SEARCH}
, {"Browser_Favorites", VK_BROWSER_FAVORITES}
, {"Browser_Home", VK_BROWSER_HOME}
, {"Volume_Mute", VK_VOLUME_MUTE}
, {"Volume_Down", VK_VOLUME_DOWN}
, {"Volume_Up", VK_VOLUME_UP}
, {"Media_Next", VK_MEDIA_NEXT_TRACK}
, {"Media_Prev", VK_MEDIA_PREV_TRACK}
, {"Media_Stop", VK_MEDIA_STOP}
, {"Media_Play_Pause", VK_MEDIA_PLAY_PAUSE}
, {"Launch_Mail", VK_LAUNCH_MAIL}
, {"Launch_Media", VK_LAUNCH_MEDIA_SELECT}
, {"Launch_App1", VK_LAUNCH_APP1}
, {"Launch_App2", VK_LAUNCH_APP2}

// Probably safest to terminate it this way, with a flag value.  (plus this makes it a little easier
// to code some loops, maybe).  Can also calculate how many elements are in the array using sizeof(array)
// divided by sizeof(element).  UPDATE: Decided not to do this in case ever decide to sort this array; don't
// want to rely on the fact that this will wind up in the right position after the sort (even though it
// should):
//, {"", 0}
};



key_to_sc_type g_key_to_sc[] =
// Even though ENTER is probably less likely to be assigned than NumpadEnter, must have ENTER as
// the primary vk because otherwise, if the user configures only naked-NumPadEnter to do something,
// RegisterHotkey() would register that vk and ENTER would also be configured to do the same thing.
{ {"NumpadEnter", SC_NUMPADENTER}

, {"Delete", SC_DELETE}
, {"Del", SC_DELETE}
, {"Insert", SC_INSERT}
, {"Ins", SC_INSERT}
// , {"Clear", SC_CLEAR}  // Seems unnecessary because there is no counterpart to the Numpad5 clear key?
, {"Up", SC_UP}
, {"Down", SC_DOWN}
, {"Left", SC_LEFT}
, {"Right", SC_RIGHT}
, {"Home", SC_HOME}
, {"End", SC_END}
, {"PgUp", SC_PGUP}
, {"PgDn", SC_PGDN}

// If user specified left or right, must use scan code to distinguish *both* halves of the pair since
// each half has same vk *and* since their generic counterparts (e.g. CONTROL vs. L/RCONTROL) are
// already handled by vk.  Note: RWIN and LWIN don't need to be handled here because they each have
// their own virtual keys.
// UPDATE: For the NT/2k/XP version, now doing these by VK since it's likely to be
// more compatible with non-standard or non-English keyboards:
/*
, {"LControl", SC_LCONTROL}
, {"RControl", SC_RCONTROL}
, {"LShift", SC_LSHIFT}
, {"RShift", SC_RSHIFT}
, {"LAlt", SC_LALT}
, {"RAlt", SC_RALT}
*/
};


// Can calc the counts only after the arrays are initialized above:
int g_key_to_vk_count = sizeof(g_key_to_vk) / sizeof(key_to_vk_type);
int g_key_to_sc_count = sizeof(g_key_to_sc) / sizeof(key_to_sc_type);

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
