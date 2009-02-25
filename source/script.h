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

#ifndef script_h
#define script_h

#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "SimpleHeap.h" // for overloaded new/delete operators.
#include "keyboard_mouse.h" // for modLR_type
#include "var.h" // for a script's variables.
#include "WinGroup.h" // for a script's Window Groups.
#include "Util.h" // for FileTimeToYYYYMMDD(), strlcpy()
#include "resources\resource.h"  // For tray icon.
#ifdef AUTOHOTKEYSC
	#include "lib\exearc_read.h"
#endif

#include "os_version.h" // For the global OS_Version object
EXTERN_OSVER; // For the access to the g_os version object without having to include globaldata.h
EXTERN_G;

#define MAX_THREADS_LIMIT UCHAR_MAX // Uses UCHAR_MAX (255) because some variables that store a thread count are UCHARs.
#define MAX_THREADS_DEFAULT 10 // Must not be higher than above.
#define EMERGENCY_THREADS 2 // This is the number of extra threads available after g_MaxThreadsTotal has been reached for the following to launch: hotkeys/etc. whose first line is something important like ExitApp or Pause. (see #MaxThreads documentation).
#define MAX_THREADS_EMERGENCY (g_MaxThreadsTotal + EMERGENCY_THREADS)
#define TOTAL_ADDITIONAL_THREADS (EMERGENCY_THREADS + 2) // See below.
// Must allow two beyond EMERGENCY_THREADS: One for the AutoExec/idle thread and one so that ExitApp()
// can run even when MAX_THREADS_EMERGENCY has been reached.
// Explanation: If/when AutoExec() finishes, although it no longer needs g_array[0] (not even
// AutoExecSectionTimeout() needs it because it either won't be called or it will return early),
// at least the following might still use g_array[0]:
// 1) Threadless (fast-mode) callbacks that have no controlling script thread; see RegisterCallbackCStub().
// 2) g_array[0].IsPaused indicates whether the script is in a paused state while idle.
// In addition, it probably simplifies the code not to reclaim g_array[0]; e.g. ++g and --g can be done
// unconditionally when creating new threads.

enum ExecUntilMode {NORMAL_MODE, UNTIL_RETURN, UNTIL_BLOCK_END, ONLY_ONE_LINE};

// It's done this way so that mAttribute can store a pointer or one of these constants.
// If it is storing a pointer for a given Action Type, be sure never to compare it
// for equality against these constants because by coincidence, the pointer value
// might just match one of them:
#define ATTR_NONE (void *)0  // Some places migh rely on this being zero.
#define ATTR_TRUE (void *)1
#define ATTR_LOOP_UNKNOWN (void *)1 // Same value as the above.        // KEEP IN SYNC WITH BELOW.
#define ATTR_LOOP_IS_UNKNOWN_OR_NONE(attr) (attr <= ATTR_LOOP_UNKNOWN) // KEEP IN SYNC WITH ABOVE.
#define ATTR_LOOP_NORMAL (void *)2
#define ATTR_LOOP_FILEPATTERN (void *)3
#define ATTR_LOOP_REG (void *)4
#define ATTR_LOOP_READ_FILE (void *)5
#define ATTR_LOOP_PARSE (void *)6
#define ATTR_LOOP_WHILE (void *)7 // Lexikos: This is used to differentiate ACT_WHILE from ACT_LOOP, allowing code to be shared.
typedef void *AttributeType;

enum FileLoopModeType {FILE_LOOP_INVALID, FILE_LOOP_FILES_ONLY, FILE_LOOP_FILES_AND_FOLDERS, FILE_LOOP_FOLDERS_ONLY};
enum VariableTypeType {VAR_TYPE_INVALID, VAR_TYPE_NUMBER, VAR_TYPE_INTEGER, VAR_TYPE_FLOAT
	, VAR_TYPE_TIME	, VAR_TYPE_DIGIT, VAR_TYPE_XDIGIT, VAR_TYPE_ALNUM, VAR_TYPE_ALPHA
	, VAR_TYPE_UPPER, VAR_TYPE_LOWER, VAR_TYPE_SPACE};

#define ATTACH_THREAD_INPUT \
	bool threads_are_attached = false;\
	DWORD target_thread = GetWindowThreadProcessId(target_window, NULL);\
	if (target_thread && target_thread != g_MainThreadID && !IsWindowHung(target_window))\
		threads_are_attached = AttachThreadInput(g_MainThreadID, target_thread, TRUE) != 0;
// BELOW IS SAME AS ABOVE except it checks do_activate and also does a SetActiveWindow():
#define ATTACH_THREAD_INPUT_AND_SETACTIVEWINDOW_IF_DO_ACTIVATE \
	bool threads_are_attached = false;\
	DWORD target_thread;\
	if (do_activate)\
	{\
		target_thread = GetWindowThreadProcessId(target_window, NULL);\
		if (target_thread && target_thread != g_MainThreadID && !IsWindowHung(target_window))\
			threads_are_attached = AttachThreadInput(g_MainThreadID, target_thread, TRUE) != 0;\
		SetActiveWindow(target_window);\
	}

#define DETACH_THREAD_INPUT \
	if (threads_are_attached)\
		AttachThreadInput(g_MainThreadID, target_thread, FALSE);

#define RESEED_RANDOM_GENERATOR \
{\
	FILETIME ft;\
	GetSystemTimeAsFileTime(&ft);\
	init_genrand(ft.dwLowDateTime);\
}

#define IS_PERSISTENT (Hotkey::sHotkeyCount || Hotstring::sHotstringCount || g_KeybdHook || g_MouseHook || g_persistent)

// Since WM_COMMAND IDs must be shared among all menus and controls, they are carefully conserved,
// especially since there are only 65,535 possible IDs.  In addition, they are assigned to ranges
// to minimize the need that they will need to be changed in the future (changing the ID of a main
// menu item, tray menu item, or a user-defined menu item [by way of increasing MAX_CONTROLS_PER_GUI]
// is bad because some scripts might be using PostMessage/SendMessage to automate AutoHotkey itself).
// For this reason, the following ranges are reserved:
// 0: unused (possibly special in some contexts)
// 1: IDOK
// 2: IDCANCEL
// 3 to 1002: GUI window control IDs (these IDs must be unique only within their parent, not across all GUI windows)
// 1003 to 65299: User Defined Menu IDs
// 65300 to 65399: Standard tray menu items.
// 65400 to 65534: main menu items (might be best to leave 65535 unused in case it ever has special meaning)
enum CommandIDs {CONTROL_ID_FIRST = IDCANCEL + 1
	, ID_USER_FIRST = MAX_CONTROLS_PER_GUI + 3 // The first ID available for user defined menu items. Do not change this (see above for why).
	, ID_USER_LAST = 65299  // The last. Especially do not change this due to scripts using Post/SendMessage to automate AutoHotkey.
	, ID_TRAY_FIRST, ID_TRAY_OPEN = ID_TRAY_FIRST
	, ID_TRAY_HELP, ID_TRAY_WINDOWSPY, ID_TRAY_RELOADSCRIPT
	, ID_TRAY_EDITSCRIPT, ID_TRAY_SUSPEND, ID_TRAY_PAUSE, ID_TRAY_EXIT
	, ID_TRAY_LAST = ID_TRAY_EXIT // But this value should never hit the below. There is debug code to enforce.
	, ID_MAIN_FIRST = 65400, ID_MAIN_LAST = 65534}; // These should match the range used by resource.h

#define GUI_INDEX_TO_ID(index) (index + CONTROL_ID_FIRST)
#define GUI_ID_TO_INDEX(id) (id - CONTROL_ID_FIRST) // Returns a small negative if "id" is invalid, such as 0.
#define GUI_HWND_TO_INDEX(hwnd) GUI_ID_TO_INDEX(GetDlgCtrlID(hwnd)) // Returns a small negative on failure (e.g. HWND not found).
// Notes about above:
// 1) Callers should call GuiType::FindControl() instead of GUI_HWND_TO_INDEX() if the hwnd might be a combobox's
//    edit control.
// 2) Testing shows that GetDlgCtrlID() is much faster than looping through a GUI window's control array to find
//    a matching HWND.


#define ERR_ABORT_NO_SPACES "The current thread will exit."
#define ERR_ABORT "  " ERR_ABORT_NO_SPACES
#define WILL_EXIT "The program will exit."
#define OLD_STILL_IN_EFFECT "The script was not reloaded; the old version will remain in effect."
#define ERR_CONTINUATION_SECTION_TOO_LONG "Continuation section too long."
#define ERR_UNRECOGNIZED_ACTION "This line does not contain a recognized action."
#define ERR_NONEXISTENT_HOTKEY "Nonexistent hotkey."
#define ERR_NONEXISTENT_VARIANT "Nonexistent hotkey variant (IfWin)."
#define ERR_NONEXISTENT_FUNCTION "Call to nonexistent function."
#define ERR_EXE_CORRUPTED "EXE corrupted"
#define ERR_PARAM1_INVALID "Parameter #1 invalid"
#define ERR_PARAM2_INVALID "Parameter #2 invalid"
#define ERR_PARAM3_INVALID "Parameter #3 invalid"
#define ERR_PARAM4_INVALID "Parameter #4 invalid"
#define ERR_PARAM5_INVALID "Parameter #5 invalid"
#define ERR_PARAM6_INVALID "Parameter #6 invalid"
#define ERR_PARAM7_INVALID "Parameter #7 invalid"
#define ERR_PARAM8_INVALID "Parameter #8 invalid"
#define ERR_PARAM1_REQUIRED "Parameter #1 required"
#define ERR_PARAM2_REQUIRED "Parameter #2 required"
#define ERR_PARAM3_REQUIRED "Parameter #3 required"
#define ERR_PARAM4_OMIT "Parameter #4 should be omitted in this case."
#define ERR_PARAM2_MUST_BE_BLANK "Parameter #2 must be blank in this case."
#define ERR_PARAM3_MUST_BE_BLANK "Parameter #3 must be blank in this case."
#define ERR_PARAM4_MUST_BE_BLANK "Parameter #4 must be blank in this case."
#define ERR_INVALID_KEY_OR_BUTTON "Invalid key or button name"
#define ERR_MISSING_OUTPUT_VAR "Requires at least one of its output variables."
#define ERR_MISSING_OPEN_PAREN "Missing \"(\""
#define ERR_MISSING_OPEN_BRACE "Missing \"{\""
#define ERR_MISSING_CLOSE_PAREN "Missing \")\""
#define ERR_MISSING_CLOSE_BRACE "Missing \"}\""
#define ERR_MISSING_CLOSE_QUOTE "Missing close-quote" // No period after short phrases.
#define ERR_MISSING_COMMA "Missing comma"             //
#define ERR_BLANK_PARAM "Blank parameter"             //
#define ERR_BYREF "Caller must pass a variable to this ByRef parameter."
#define ERR_ELSE_WITH_NO_IF "ELSE with no matching IF"
#define ERR_OUTOFMEM "Out of memory."  // Used by RegEx too, so don't change it without also changing RegEx to keep the former string.
#define ERR_EXPR_TOO_LONG "Expression too long"
#define ERR_MEM_LIMIT_REACHED "Memory limit reached (see #MaxMem in the help file)." ERR_ABORT
#define ERR_NO_LABEL "Target label does not exist."
#define ERR_MENU "Menu does not exist."
#define ERR_SUBMENU "Submenu does not exist."
#define ERR_WINDOW_PARAM "Requires at least one of its window parameters."
#define ERR_ON_OFF "Requires ON/OFF/blank"
#define ERR_ON_OFF_LOCALE "Requires ON/OFF/LOCALE"
#define ERR_ON_OFF_TOGGLE "Requires ON/OFF/TOGGLE/blank"
#define ERR_ON_OFF_TOGGLE_PERMIT "Requires ON/OFF/TOGGLE/PERMIT/blank"
#define ERR_TITLEMATCHMODE "Requires 1/2/3/Slow/Fast"
#define ERR_MENUTRAY "Supported only for the tray menu"
#define ERR_REG_KEY "Invalid registry root key"
#define ERR_REG_VALUE_TYPE "Invalid registry value type"
#define ERR_INVALID_DATETIME "Invalid YYYYMMDDHHMISS value"
#define ERR_MOUSE_BUTTON "Invalid mouse button"
#define ERR_MOUSE_COORD "X & Y must be either both absent or both present."
#define ERR_DIVIDEBYZERO "Divide by zero"
#define ERR_PERCENT "Must be between -100 and 100."
#define ERR_MOUSE_SPEED "Mouse speed must be between 0 and " MAX_MOUSE_SPEED_STR "."
#define ERR_VAR_IS_READONLY "Not allowed as an output variable."

//----------------------------------------------------------------------------------

void DoIncrementalMouseMove(int aX1, int aY1, int aX2, int aY2, int aSpeed);
DWORD ProcessExist9x2000(char *aProcess, char *aProcessName);
DWORD ProcessExistNT4(char *aProcess, char *aProcessName);

inline DWORD ProcessExist(char *aProcess, char *aProcessName = NULL)
{
	return g_os.IsWinNT4() ? ProcessExistNT4(aProcess, aProcessName)
		: ProcessExist9x2000(aProcess, aProcessName);
}

bool Util_Shutdown(int nFlag);
BOOL Util_ShutdownHandler(HWND hwnd, DWORD lParam);
void Util_WinKill(HWND hWnd);

enum MainWindowModes {MAIN_MODE_NO_CHANGE, MAIN_MODE_LINES, MAIN_MODE_VARS
	, MAIN_MODE_HOTKEYS, MAIN_MODE_KEYHISTORY, MAIN_MODE_REFRESH};
ResultType ShowMainWindow(MainWindowModes aMode = MAIN_MODE_NO_CHANGE, bool aRestricted = true);
DWORD GetAHKInstallDir(char *aBuf);


struct InputBoxType
{
	char *title;
	char *text;
	int width;
	int height;
	int xpos;
	int ypos;
	Var *output_var;
	char password_char;
	char *default_string;
	DWORD timeout;
	HWND hwnd;
};

struct SplashType
{
	int width;
	int height;
	int bar_pos;  // The amount of progress of the bar (it's position).
	int margin_x; // left/right margin
	int margin_y; // top margin
	int text1_height; // Height of main text control.
	int object_width;   // Width of image.
	int object_height;  // Height of the progress bar or image.
	HWND hwnd;
	LPPICTURE pic; // For SplashImage.
	HWND hwnd_bar;
	HWND hwnd_text1;  // MainText
	HWND hwnd_text2;  // SubText
	HFONT hfont1; // Main
	HFONT hfont2; // Sub
	HBRUSH hbrush; // Window background color brush.
	COLORREF color_bk; // The background color itself.
	COLORREF color_text; // Color of the font.
};

// Use GetClientRect() to determine the available width so that control's can be centered.
#define SPLASH_CALC_YPOS \
	int bar_y = splash.margin_y + (splash.text1_height ? (splash.text1_height + splash.margin_y) : 0);\
	int sub_y = bar_y + splash.object_height + (splash.object_height ? splash.margin_y : 0); // i.e. don't include margin_y twice if there's no bar.
#define PROGRESS_MAIN_POS splash.margin_x, splash.margin_y, control_width, splash.text1_height
#define PROGRESS_BAR_POS  splash.margin_x, bar_y, control_width, splash.object_height
#define PROGRESS_SUB_POS  splash.margin_x, sub_y, control_width, (client_rect.bottom - client_rect.top) - sub_y

// From AutoIt3's InputBox.  This doesn't add a measurable amount of code size, so the compiler seems to implement
// it efficiently (somewhat like a macro).
template <class T>
inline void swap(T &v1, T &v2) {
	T tmp=v1;
	v1=v2;
	v2=tmp;
}

#define INPUTBOX_DEFAULT INT_MIN
ResultType InputBox(Var *aOutputVar, char *aTitle, char *aText, bool aHideInput
	, int aWidth, int aHeight, int aX, int aY, double aTimeout, char *aDefault);
BOOL CALLBACK InputBoxProc(HWND hWndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
VOID CALLBACK InputBoxTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);
VOID CALLBACK DerefTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);
BOOL CALLBACK EnumChildFindSeqNum(HWND aWnd, LPARAM lParam);
BOOL CALLBACK EnumChildFindPoint(HWND aWnd, LPARAM lParam);
BOOL CALLBACK EnumChildGetControlList(HWND aWnd, LPARAM lParam);
BOOL CALLBACK EnumMonitorProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM lParam);
BOOL CALLBACK EnumChildGetText(HWND aWnd, LPARAM lParam);
LRESULT CALLBACK MainWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam);
bool HandleMenuItem(HWND aHwnd, WORD aMenuItemID, WPARAM aGuiIndex);


typedef UINT LineNumberType;
typedef WORD FileIndexType; // Use WORD to conserve memory due to its use in the Line class (adjacency to other members and due to 4-byte struct alignment).
#define ABSOLUTE_MAX_SOURCE_FILES 0xFFFF // Keep this in sync with the capacity of the type above.  Actually it could hold 0xFFFF+1, but avoid the final item for maintainability (otherwise max-index won't be able to fit inside a variable of that type).

#define LOADING_FAILED UINT_MAX

// -2 for the beginning and ending g_DerefChars:
#define MAX_VAR_NAME_LENGTH (UCHAR_MAX - 2)
#define MAX_FUNCTION_PARAMS UCHAR_MAX // Also conserves stack space to support future attributes such as param default values.
#define MAX_DEREFS_PER_ARG 512

typedef WORD DerefLengthType; // WORD might perform better than UCHAR, but this can be changed to UCHAR if another field is ever needed in the struct.
typedef UCHAR DerefParamCountType;

class Func; // Forward declaration for use below.
struct DerefType
{
	char *marker;
	union
	{
		Var *var;
		Func *func;
	};
	// Keep any fields that aren't an even multiple of 4 adjacent to each other.  This conserves memory
	// due to byte-alignment:
	bool is_function; // This should be kept pure bool to allow easy determination of what's in the union, above.
	DerefParamCountType param_count; // The actual number of parameters present in this function *call*.  Left uninitialized except for functions.
	DerefLengthType length; // Listed only after byte-sized fields, due to it being a WORD.
};

typedef UCHAR ArgTypeType;  // UCHAR vs. an enum, to save memory.
#define ARG_TYPE_NORMAL     (UCHAR)0
#define ARG_TYPE_INPUT_VAR  (UCHAR)1
#define ARG_TYPE_OUTPUT_VAR (UCHAR)2

struct ArgStruct
{
	ArgTypeType type;
	bool is_expression; // Whether this ARG is known to contain an expression.
	// Above are kept adjacent to each other to conserve memory (any fields that aren't an even
	// multiple of 4, if adjacent to each other, consume less memory due to default byte alignment
	// setting [which helps performance]).
	WORD length; // Keep adjacent to above so that it uses no extra memory. This member was added in v1.0.44.14 to improve runtime performance.  It relies on the fact that an arg's literal text can't be longer than LINE_SIZE.
	char *text;
	DerefType *deref;  // Will hold a NULL-terminated array of var-deref locations within <text>.
	ExprTokenType *postfix;  // An array of tokens in postfix order. Also used for ACT_ADD and others to store pre-converted binary integers.
};


// Some of these lengths and such are based on the MSDN example at
// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/sysinfo/base/enumerating_registry_subkeys.asp:
// FIX FOR v1.0.48: 
// OLDER (v1.0.44.07): Someone reported that a stack overflow was possible, implying that it only happens
// during extremely deep nesting of subkey names (perhaps a hundred or more nested subkeys).  Upon review, it seems
// that the prior limit of 16383 for value-name-length is higher than needed; testing shows that a value name can't
// be longer than 259 (limit might even be 255 if API vs. RegEdit is used to create the name).  Testing also shows
// that the total path name of a registry item (including item/value name but excluding the name of the root key)
// obeys the same limit BUT ONLY within the RegEdit GUI.  RegEdit seems capable of importing subkeys whose names
// (even without any value name appended) are longer than 259 characters (see comments higher above).
#define MAX_REG_ITEM_SIZE 1024 // Needs to be greater than 260 (see comments above), but I couldn't find any documentation at MSDN or the web about the max length of a subkey name.  One example at MSDN RegEnumKeyEx() uses MAX_KEY_LENGTH=255 and MAX_VALUE_NAME=16383, but clearly MAX_KEY_LENGTH should be larger.
#define REG_SUBKEY -2 // Custom type, not standard in Windows.
struct RegItemStruct
{
	HKEY root_key_type, root_key;  // root_key_type is always a local HKEY, whereas root_key can be a remote handle.
	char subkey[MAX_REG_ITEM_SIZE];  // The branch of the registry where this subkey or value is located.
	char name[MAX_REG_ITEM_SIZE]; // The subkey or value name.
	DWORD type; // Value Type (e.g REG_DWORD).
	FILETIME ftLastWriteTime; // Non-initialized.
	void InitForValues() {ftLastWriteTime.dwHighDateTime = ftLastWriteTime.dwLowDateTime = 0;}
	void InitForSubkeys() {type = REG_SUBKEY;}  // To distinguish REG_DWORD and such from the subkeys themselves.
	RegItemStruct(HKEY aRootKeyType, HKEY aRootKey, char *aSubKey)
		: root_key_type(aRootKeyType), root_key(aRootKey), type(REG_NONE)
	{
		*name = '\0';
		// Make a local copy on the caller's stack so that if the current script subroutine is
		// interrupted to allow another to run, the contents of the deref buffer is saved here:
		strlcpy(subkey, aSubKey, sizeof(subkey));
		// Even though the call may work with a trailing backslash, it's best to remove it
		// so that consistent results are delivered to the user.  For example, if the script
		// is enumerating recursively into a subkey, subkeys deeper down will not include the
		// trailing backslash when they are reported.  So the user's own subkey should not
		// have one either so that when A_ScriptSubKey is referenced in the script, it will
		// always show up as the value without a trailing backslash:
		size_t length = strlen(subkey);
		if (length && subkey[length - 1] == '\\')
			subkey[length - 1] = '\0';
	}
};

struct LoopReadFileStruct
{
	FILE *mReadFile, *mWriteFile;
	char mWriteFileName[MAX_PATH];
	#define READ_FILE_LINE_SIZE (64 * 1024)  // This is also used by FileReadLine().
	char mCurrentLine[READ_FILE_LINE_SIZE];
	LoopReadFileStruct(FILE *aReadFile, char *aWriteFileName)
		: mReadFile(aReadFile), mWriteFile(NULL) // mWriteFile is opened by FileAppend() only upon first use.
	{
		// Use our own buffer because caller's is volatile due to possibly being in the deref buffer:
		strlcpy(mWriteFileName, aWriteFileName, sizeof(mWriteFileName));
		*mCurrentLine = '\0';
	}
};


typedef UCHAR ArgCountType;
#define MAX_ARGS 20   // Maximum number of args used by any command.


enum DllArgTypes {DLL_ARG_INVALID, DLL_ARG_STR, DLL_ARG_INT, DLL_ARG_SHORT, DLL_ARG_CHAR, DLL_ARG_INT64
	, DLL_ARG_FLOAT, DLL_ARG_DOUBLE};  // Some sections might rely on DLL_ARG_INVALID being 0.


// Note that currently this value must fit into a sc_type variable because that is how TextToKey()
// stores it in the hotkey class.  sc_type is currently a UINT, and will always be at least a
// WORD in size, so it shouldn't be much of an issue:
#define MAX_JOYSTICKS 16  // The maximum allowed by any Windows operating system.
#define MAX_JOY_BUTTONS 32 // Also the max that Windows supports.
enum JoyControls {JOYCTRL_INVALID, JOYCTRL_XPOS, JOYCTRL_YPOS, JOYCTRL_ZPOS
, JOYCTRL_RPOS, JOYCTRL_UPOS, JOYCTRL_VPOS, JOYCTRL_POV
, JOYCTRL_NAME, JOYCTRL_BUTTONS, JOYCTRL_AXES, JOYCTRL_INFO
, JOYCTRL_1, JOYCTRL_2, JOYCTRL_3, JOYCTRL_4, JOYCTRL_5, JOYCTRL_6, JOYCTRL_7, JOYCTRL_8  // Buttons.
, JOYCTRL_9, JOYCTRL_10, JOYCTRL_11, JOYCTRL_12, JOYCTRL_13, JOYCTRL_14, JOYCTRL_15, JOYCTRL_16
, JOYCTRL_17, JOYCTRL_18, JOYCTRL_19, JOYCTRL_20, JOYCTRL_21, JOYCTRL_22, JOYCTRL_23, JOYCTRL_24
, JOYCTRL_25, JOYCTRL_26, JOYCTRL_27, JOYCTRL_28, JOYCTRL_29, JOYCTRL_30, JOYCTRL_31, JOYCTRL_32
, JOYCTRL_BUTTON_MAX = JOYCTRL_32
};
#define IS_JOYSTICK_BUTTON(joy) (joy >= JOYCTRL_1 && joy <= JOYCTRL_BUTTON_MAX)

enum WinGetCmds {WINGET_CMD_INVALID, WINGET_CMD_ID, WINGET_CMD_IDLAST, WINGET_CMD_PID, WINGET_CMD_PROCESSNAME
	, WINGET_CMD_COUNT, WINGET_CMD_LIST, WINGET_CMD_MINMAX, WINGET_CMD_CONTROLLIST, WINGET_CMD_CONTROLLISTHWND
	, WINGET_CMD_STYLE, WINGET_CMD_EXSTYLE, WINGET_CMD_TRANSPARENT, WINGET_CMD_TRANSCOLOR
};

enum SysGetCmds {SYSGET_CMD_INVALID, SYSGET_CMD_METRICS, SYSGET_CMD_MONITORCOUNT, SYSGET_CMD_MONITORPRIMARY
	, SYSGET_CMD_MONITORAREA, SYSGET_CMD_MONITORWORKAREA, SYSGET_CMD_MONITORNAME
};

enum TransformCmds {TRANS_CMD_INVALID, TRANS_CMD_ASC, TRANS_CMD_CHR, TRANS_CMD_DEREF
	, TRANS_CMD_UNICODE, TRANS_CMD_HTML
	, TRANS_CMD_MOD, TRANS_CMD_POW, TRANS_CMD_EXP, TRANS_CMD_SQRT, TRANS_CMD_LOG, TRANS_CMD_LN
	, TRANS_CMD_ROUND, TRANS_CMD_CEIL, TRANS_CMD_FLOOR, TRANS_CMD_ABS
	, TRANS_CMD_SIN, TRANS_CMD_COS, TRANS_CMD_TAN, TRANS_CMD_ASIN, TRANS_CMD_ACOS, TRANS_CMD_ATAN
	, TRANS_CMD_BITAND, TRANS_CMD_BITOR, TRANS_CMD_BITXOR, TRANS_CMD_BITNOT
	, TRANS_CMD_BITSHIFTLEFT, TRANS_CMD_BITSHIFTRIGHT
};

enum MenuCommands {MENU_CMD_INVALID, MENU_CMD_SHOW, MENU_CMD_USEERRORLEVEL
	, MENU_CMD_ADD, MENU_CMD_RENAME, MENU_CMD_CHECK, MENU_CMD_UNCHECK, MENU_CMD_TOGGLECHECK
	, MENU_CMD_ENABLE, MENU_CMD_DISABLE, MENU_CMD_TOGGLEENABLE
	, MENU_CMD_STANDARD, MENU_CMD_NOSTANDARD, MENU_CMD_COLOR, MENU_CMD_DEFAULT, MENU_CMD_NODEFAULT
	, MENU_CMD_DELETE, MENU_CMD_DELETEALL, MENU_CMD_TIP, MENU_CMD_ICON, MENU_CMD_NOICON
	, MENU_CMD_CLICK, MENU_CMD_MAINWINDOW, MENU_CMD_NOMAINWINDOW
};

#define AHK_LV_SELECT       0x0100
#define AHK_LV_DESELECT     0x0200
#define AHK_LV_FOCUS        0x0400
#define AHK_LV_DEFOCUS      0x0800
#define AHK_LV_CHECK        0x1000
#define AHK_LV_UNCHECK      0x2000
#define AHK_LV_DROPHILITE   0x4000
#define AHK_LV_UNDROPHILITE 0x8000
// Although there's no room remaining in the BYTE for LVIS_CUT (AHK_LV_CUT) [assuming it's ever needed],
// it might be possible to squeeze more info into it as follows:
// Each pair of bits can represent three values (other than zero).  But since only two values are needed
// (since an item can't be both selected an deselected simultaneously), one value in each pair is available
// for future use such as LVIS_CUT.

enum GuiCommands {GUI_CMD_INVALID, GUI_CMD_OPTIONS, GUI_CMD_ADD, GUI_CMD_MARGIN, GUI_CMD_MENU
	, GUI_CMD_SHOW, GUI_CMD_SUBMIT, GUI_CMD_CANCEL, GUI_CMD_MINIMIZE, GUI_CMD_MAXIMIZE, GUI_CMD_RESTORE
	, GUI_CMD_DESTROY, GUI_CMD_FONT, GUI_CMD_TAB, GUI_CMD_LISTVIEW, GUI_CMD_TREEVIEW, GUI_CMD_DEFAULT
	, GUI_CMD_COLOR, GUI_CMD_FLASH
};

enum GuiControlCmds {GUICONTROL_CMD_INVALID, GUICONTROL_CMD_OPTIONS, GUICONTROL_CMD_CONTENTS, GUICONTROL_CMD_TEXT
	, GUICONTROL_CMD_MOVE, GUICONTROL_CMD_MOVEDRAW, GUICONTROL_CMD_FOCUS, GUICONTROL_CMD_ENABLE, GUICONTROL_CMD_DISABLE
	, GUICONTROL_CMD_SHOW, GUICONTROL_CMD_HIDE, GUICONTROL_CMD_CHOOSE, GUICONTROL_CMD_CHOOSESTRING
	, GUICONTROL_CMD_FONT
};

enum GuiControlGetCmds {GUICONTROLGET_CMD_INVALID, GUICONTROLGET_CMD_CONTENTS, GUICONTROLGET_CMD_POS
	, GUICONTROLGET_CMD_FOCUS, GUICONTROLGET_CMD_FOCUSV, GUICONTROLGET_CMD_ENABLED, GUICONTROLGET_CMD_VISIBLE
	, GUICONTROLGET_CMD_HWND
};

typedef UCHAR GuiControls;
enum GuiControlTypes {GUI_CONTROL_INVALID // GUI_CONTROL_INVALID must be zero due to things like ZeroMemory() on the struct.
	, GUI_CONTROL_TEXT, GUI_CONTROL_PIC, GUI_CONTROL_GROUPBOX
	, GUI_CONTROL_BUTTON, GUI_CONTROL_CHECKBOX, GUI_CONTROL_RADIO
	, GUI_CONTROL_DROPDOWNLIST, GUI_CONTROL_COMBOBOX
	, GUI_CONTROL_LISTBOX, GUI_CONTROL_LISTVIEW, GUI_CONTROL_TREEVIEW
	, GUI_CONTROL_EDIT, GUI_CONTROL_DATETIME, GUI_CONTROL_MONTHCAL, GUI_CONTROL_HOTKEY
	, GUI_CONTROL_UPDOWN, GUI_CONTROL_SLIDER, GUI_CONTROL_PROGRESS, GUI_CONTROL_TAB, GUI_CONTROL_TAB2
	, GUI_CONTROL_STATUSBAR}; // Kept last to reflect it being bottommost in switch()s (for perf), since not too often used.

enum ThreadCommands {THREAD_CMD_INVALID, THREAD_CMD_PRIORITY, THREAD_CMD_INTERRUPT, THREAD_CMD_NOTIMERS};

#define PROCESS_PRIORITY_LETTERS "LBNAHR"
enum ProcessCmds {PROCESS_CMD_INVALID, PROCESS_CMD_EXIST, PROCESS_CMD_CLOSE, PROCESS_CMD_PRIORITY
	, PROCESS_CMD_WAIT, PROCESS_CMD_WAITCLOSE};

enum ControlCmds {CONTROL_CMD_INVALID, CONTROL_CMD_CHECK, CONTROL_CMD_UNCHECK
	, CONTROL_CMD_ENABLE, CONTROL_CMD_DISABLE, CONTROL_CMD_SHOW, CONTROL_CMD_HIDE
	, CONTROL_CMD_STYLE, CONTROL_CMD_EXSTYLE
	, CONTROL_CMD_SHOWDROPDOWN, CONTROL_CMD_HIDEDROPDOWN
	, CONTROL_CMD_TABLEFT, CONTROL_CMD_TABRIGHT
	, CONTROL_CMD_ADD, CONTROL_CMD_DELETE, CONTROL_CMD_CHOOSE
	, CONTROL_CMD_CHOOSESTRING, CONTROL_CMD_EDITPASTE};

enum ControlGetCmds {CONTROLGET_CMD_INVALID, CONTROLGET_CMD_CHECKED, CONTROLGET_CMD_ENABLED
	, CONTROLGET_CMD_VISIBLE, CONTROLGET_CMD_TAB, CONTROLGET_CMD_FINDSTRING
	, CONTROLGET_CMD_CHOICE, CONTROLGET_CMD_LIST, CONTROLGET_CMD_LINECOUNT, CONTROLGET_CMD_CURRENTLINE
	, CONTROLGET_CMD_CURRENTCOL, CONTROLGET_CMD_LINE, CONTROLGET_CMD_SELECTED
	, CONTROLGET_CMD_STYLE, CONTROLGET_CMD_EXSTYLE, CONTROLGET_CMD_HWND};

enum DriveCmds {DRIVE_CMD_INVALID, DRIVE_CMD_EJECT, DRIVE_CMD_LOCK, DRIVE_CMD_UNLOCK, DRIVE_CMD_LABEL};

enum DriveGetCmds {DRIVEGET_CMD_INVALID, DRIVEGET_CMD_LIST, DRIVEGET_CMD_FILESYSTEM, DRIVEGET_CMD_LABEL
	, DRIVEGET_CMD_SETLABEL, DRIVEGET_CMD_SERIAL, DRIVEGET_CMD_TYPE, DRIVEGET_CMD_STATUS
	, DRIVEGET_CMD_STATUSCD, DRIVEGET_CMD_CAPACITY};

enum WinSetAttributes {WINSET_INVALID, WINSET_TRANSPARENT, WINSET_TRANSCOLOR, WINSET_ALWAYSONTOP
	, WINSET_BOTTOM, WINSET_TOP, WINSET_STYLE, WINSET_EXSTYLE, WINSET_REDRAW, WINSET_ENABLE, WINSET_DISABLE
	, WINSET_REGION};


class Label; // Forward declaration so that each can use the other.
class Line
{
private:
	#define SET_S_DEREF_BUF(ptr, size) sDerefBuf = ptr, sDerefBufSize = size
	#define NULLIFY_S_DEREF_BUF \
	{\
		SET_S_DEREF_BUF(NULL, 0);\
		if (sDerefBufSize > LARGE_DEREF_BUF_SIZE)\
			--sLargeDerefBufs;\
	}
	static char *sDerefBuf;  // Buffer to hold the values of any args that need to be dereferenced.
	static size_t sDerefBufSize;
	static int sLargeDerefBufs;

	// Static because only one line can be Expanded at a time (not to mention the fact that we
	// wouldn't want the size of each line to be expanded by this size):
	static char *sArgDeref[MAX_ARGS];
	static Var *sArgVar[MAX_ARGS];

	ResultType EvaluateCondition();
	ResultType Line::PerformLoop(char **apReturnValue, bool &aContinueMainLoop, Line *&aJumpToLine
		, __int64 aIterationLimit, bool aIsInfinite);
	ResultType Line::PerformLoopFilePattern(char **apReturnValue, bool &aContinueMainLoop, Line *&aJumpToLine
		, FileLoopModeType aFileLoopMode, bool aRecurseSubfolders, char *aFilePattern);
	ResultType PerformLoopReg(char **apReturnValue, bool &aContinueMainLoop, Line *&aJumpToLine
		, FileLoopModeType aFileLoopMode, bool aRecurseSubfolders, HKEY aRootKeyType, HKEY aRootKey, char *aRegSubkey);
	ResultType PerformLoopParse(char **apReturnValue, bool &aContinueMainLoop, Line *&aJumpToLine);
	ResultType Line::PerformLoopParseCSV(char **apReturnValue, bool &aContinueMainLoop, Line *&aJumpToLine);
	ResultType PerformLoopReadFile(char **apReturnValue, bool &aContinueMainLoop, Line *&aJumpToLine, FILE *aReadFile, char *aWriteFileName);
	ResultType PerformLoopWhile(char **apReturnValue, bool &aContinueMainLoop, Line *&aJumpToLine); // Lexikos: ACT_WHILE.
	ResultType Perform();

	ResultType MouseGetPos(DWORD aOptions);
	ResultType FormatTime(char *aYYYYMMDD, char *aFormat);
	ResultType PerformAssign();
	ResultType StringReplace();
	ResultType StringSplit(char *aArrayName, char *aInputString, char *aDelimiterList, char *aOmitList);
	ResultType SplitPath(char *aFileSpec);
	ResultType PerformSort(char *aContents, char *aOptions);
	ResultType GetKeyJoyState(char *aKeyName, char *aOption);
	ResultType DriveSpace(char *aPath, bool aGetFreeSpace);
	ResultType Drive(char *aCmd, char *aValue, char *aValue2);
	ResultType DriveLock(char aDriveLetter, bool aLockIt);
	ResultType DriveGet(char *aCmd, char *aValue);
	ResultType SoundSetGet(char *aSetting, DWORD aComponentType, int aComponentInstance
		, DWORD aControlType, UINT aMixerID);
	ResultType SoundGetWaveVolume(HWAVEOUT aDeviceID);
	ResultType SoundSetWaveVolume(char *aVolume, HWAVEOUT aDeviceID);
	ResultType SoundPlay(char *aFilespec, bool aSleepUntilDone);
	ResultType URLDownloadToFile(char *aURL, char *aFilespec);
	ResultType FileSelectFile(char *aOptions, char *aWorkingDir, char *aGreeting, char *aFilter);

	// Bitwise flags:
	#define FSF_ALLOW_CREATE 0x01
	#define FSF_EDITBOX      0x02
	#define FSF_NONEWDIALOG  0x04
	ResultType FileSelectFolder(char *aRootDir, char *aOptions, char *aGreeting);

	ResultType FileGetShortcut(char *aShortcutFile);
	ResultType FileCreateShortcut(char *aTargetFile, char *aShortcutFile, char *aWorkingDir, char *aArgs
		, char *aDescription, char *aIconFile, char *aHotkey, char *aIconNumber, char *aRunState);
	ResultType FileCreateDir(char *aDirSpec);
	ResultType FileRead(char *aFilespec);
	ResultType FileReadLine(char *aFilespec, char *aLineNumber);
	ResultType FileAppend(char *aFilespec, char *aBuf, LoopReadFileStruct *aCurrentReadFile);
	ResultType WriteClipboardToFile(char *aFilespec);
	ResultType ReadClipboardFromFile(HANDLE hfile);
	ResultType FileDelete();
	ResultType FileRecycle(char *aFilePattern);
	ResultType FileRecycleEmpty(char *aDriveLetter);
	ResultType FileInstall(char *aSource, char *aDest, char *aFlag);

	ResultType FileGetAttrib(char *aFilespec);
	int FileSetAttrib(char *aAttributes, char *aFilePattern, FileLoopModeType aOperateOnFolders
		, bool aDoRecurse, bool aCalledRecursively = false);
	ResultType FileGetTime(char *aFilespec, char aWhichTime);
	int FileSetTime(char *aYYYYMMDD, char *aFilePattern, char aWhichTime
		, FileLoopModeType aOperateOnFolders, bool aDoRecurse, bool aCalledRecursively = false);
	ResultType FileGetSize(char *aFilespec, char *aGranularity);
	ResultType FileGetVersion(char *aFilespec);

	ResultType IniRead(char *aFilespec, char *aSection, char *aKey, char *aDefault);
	ResultType IniWrite(char *aValue, char *aFilespec, char *aSection, char *aKey);
	ResultType IniDelete(char *aFilespec, char *aSection, char *aKey);
	ResultType RegRead(HKEY aRootKey, char *aRegSubkey, char *aValueName);
	ResultType RegWrite(DWORD aValueType, HKEY aRootKey, char *aRegSubkey, char *aValueName, char *aValue);
	ResultType RegDelete(HKEY aRootKey, char *aRegSubkey, char *aValueName);
	static bool RegRemoveSubkeys(HKEY hRegKey);

	#define DESTROY_SPLASH \
	{\
		if (g_hWndSplash && IsWindow(g_hWndSplash))\
			DestroyWindow(g_hWndSplash);\
		g_hWndSplash = NULL;\
	}
	ResultType SplashTextOn(int aWidth, int aHeight, char *aTitle, char *aText);
	ResultType Splash(char *aOptions, char *aSubText, char *aMainText, char *aTitle, char *aFontName
		, char *aImageFile, bool aSplashImage);

	ResultType ToolTip(char *aText, char *aX, char *aY, char *aID);
	ResultType TrayTip(char *aTitle, char *aText, char *aTimeout, char *aOptions);
	ResultType Transform(char *aCmd, char *aValue1, char *aValue2);
	ResultType Input(); // The Input command.

	#define SW_NONE -1
	ResultType PerformShowWindow(ActionTypeType aActionType, char *aTitle = "", char *aText = ""
		, char *aExcludeTitle = "", char *aExcludeText = "");
	ResultType PerformWait();

	ResultType WinMove(char *aTitle, char *aText, char *aX, char *aY
		, char *aWidth = "", char *aHeight = "", char *aExcludeTitle = "", char *aExcludeText = "");
	ResultType WinMenuSelectItem(char *aTitle, char *aText, char *aMenu1, char *aMenu2
		, char *aMenu3, char *aMenu4, char *aMenu5, char *aMenu6, char *aMenu7
		, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlSend(char *aControl, char *aKeysToSend, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText, bool aSendRaw);
	ResultType ControlClick(vk_type aVK, int aClickCount, char *aOptions, char *aControl
		, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlMove(char *aControl, char *aX, char *aY, char *aWidth, char *aHeight
		, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlGetPos(char *aControl, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlGetFocus(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlFocus(char *aControl, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlSetText(char *aControl, char *aNewText, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlGetText(char *aControl, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlGetListView(Var &aOutputVar, HWND aHwnd, char *aOptions);
	ResultType Control(char *aCmd, char *aValue, char *aControl, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlGet(char *aCommand, char *aValue, char *aControl, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType GuiControl(char *aCommand, char *aControlID, char *aParam3);
	ResultType GuiControlGet(char *aCommand, char *aControlID, char *aParam3);
	ResultType StatusBarGetText(char *aPart, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType StatusBarWait(char *aTextToWaitFor, char *aSeconds, char *aPart, char *aTitle, char *aText
		, char *aInterval, char *aExcludeTitle, char *aExcludeText);
	ResultType ScriptPostSendMessage(bool aUseSend);
	ResultType ScriptProcess(char *aCmd, char *aProcess, char *aParam3);
	ResultType WinSet(char *aAttrib, char *aValue, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType WinSetTitle(char *aTitle, char *aText, char *aNewTitle
		, char *aExcludeTitle = "", char *aExcludeText = "");
	ResultType WinGetTitle(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType WinGetClass(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType WinGet(char *aCmd, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType WinGetControlList(Var &aOutputVar, HWND aTargetWindow, bool aFetchHWNDs);
	ResultType WinGetText(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType WinGetPos(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType EnvGet(char *aEnvVarName);
	ResultType SysGet(char *aCmd, char *aValue);
	ResultType PixelSearch(int aLeft, int aTop, int aRight, int aBottom, COLORREF aColorBGR, int aVariation
		, char *aOptions, bool aIsPixelGetColor);
	ResultType ImageSearch(int aLeft, int aTop, int aRight, int aBottom, char *aImageFile);
	ResultType PixelGetColor(int aX, int aY, char *aOptions);

	static ResultType SetToggleState(vk_type aVK, ToggleValueType &ForceLock, char *aToggleText);

public:
	// Keep any fields that aren't an even multiple of 4 adjacent to each other.  This conserves memory
	// due to byte-alignment:
	ActionTypeType mActionType; // What type of line this is.
	ArgCountType mArgc; // How many arguments exist in mArg[].
	FileIndexType mFileIndex; // Which file the line came from.  0 is the first, and it's the main script file.

	ArgStruct *mArg; // Will be used to hold a dynamic array of dynamic Args.
	LineNumberType mLineNumber;  // The line number in the file from which the script was loaded, for debugging.
	AttributeType mAttribute;
	Line *mPrevLine, *mNextLine; // The prev & next lines adjacent to this one in the linked list; NULL if none.
	Line *mRelatedLine;  // e.g. the "else" that belongs to this "if"
	Line *mParentLine; // Indicates the parent (owner) of this line.
	// Probably best to always use ARG1 even if other things have supposedly verfied
	// that it exists, since it's count-check should make the dereference of a NULL
	// pointer (or accessing non-existent array elements) virtually impossible.
	// Empty-string is probably more universally useful than NULL, since some
	// API calls and other functions might not appreciate receiving NULLs.  In addition,
	// always remembering to have to check for NULL makes things harder to maintain
	// and more bug-prone.  The below macros rely upon the fact that the individual
	// elements of mArg cannot be NULL (because they're explicitly set to be blank
	// when the user has omitted an arg in between two non-blank args).  Later, might
	// want to review if any of the API calls used expect a string whose contents are
	// modifiable.
	#define RAW_ARG1 (mArgc > 0 ? mArg[0].text : "")
	#define RAW_ARG2 (mArgc > 1 ? mArg[1].text : "")
	#define RAW_ARG3 (mArgc > 2 ? mArg[2].text : "")
	#define RAW_ARG4 (mArgc > 3 ? mArg[3].text : "")
	#define RAW_ARG5 (mArgc > 4 ? mArg[4].text : "")
	#define RAW_ARG6 (mArgc > 5 ? mArg[5].text : "")
	#define RAW_ARG7 (mArgc > 6 ? mArg[6].text : "")
	#define RAW_ARG8 (mArgc > 7 ? mArg[7].text : "")

	#define LINE_RAW_ARG1 (line->mArgc > 0 ? line->mArg[0].text : "")
	#define LINE_RAW_ARG2 (line->mArgc > 1 ? line->mArg[1].text : "")
	#define LINE_RAW_ARG3 (line->mArgc > 2 ? line->mArg[2].text : "")
	#define LINE_RAW_ARG4 (line->mArgc > 3 ? line->mArg[3].text : "")
	#define LINE_RAW_ARG5 (line->mArgc > 4 ? line->mArg[4].text : "")
	#define LINE_RAW_ARG6 (line->mArgc > 5 ? line->mArg[5].text : "")
	#define LINE_RAW_ARG7 (line->mArgc > 6 ? line->mArg[6].text : "")
	#define LINE_RAW_ARG8 (line->mArgc > 7 ? line->mArg[7].text : "")
	#define LINE_RAW_ARG9 (line->mArgc > 8 ? line->mArg[8].text : "")
	
	#define NEW_RAW_ARG1 (aArgc > 0 ? new_arg[0].text : "") // Helps performance to use this vs. LINE_RAW_ARG where possible.
	#define NEW_RAW_ARG2 (aArgc > 1 ? new_arg[1].text : "")
	#define NEW_RAW_ARG3 (aArgc > 2 ? new_arg[2].text : "")
	#define NEW_RAW_ARG4 (aArgc > 3 ? new_arg[3].text : "")
	#define NEW_RAW_ARG5 (aArgc > 4 ? new_arg[4].text : "")
	#define NEW_RAW_ARG6 (aArgc > 5 ? new_arg[5].text : "")
	#define NEW_RAW_ARG7 (aArgc > 6 ? new_arg[6].text : "")
	#define NEW_RAW_ARG8 (aArgc > 7 ? new_arg[7].text : "")
	#define NEW_RAW_ARG9 (aArgc > 8 ? new_arg[8].text : "")
	
	#define SAVED_ARG1 (mArgc > 0 ? arg[0] : "")
	#define SAVED_ARG2 (mArgc > 1 ? arg[1] : "")
	#define SAVED_ARG3 (mArgc > 2 ? arg[2] : "")
	#define SAVED_ARG4 (mArgc > 3 ? arg[3] : "")
	#define SAVED_ARG5 (mArgc > 4 ? arg[4] : "")

	// For the below, it is the caller's responsibility to ensure that mArgc is
	// large enough (either via load-time validation or a runtime check of mArgc).
	// This is because for performance reasons, the sArgVar entry for omitted args isn't
	// initialized, so may have an old/obsolete value from some previous command.
	#define OUTPUT_VAR (*sArgVar) // ExpandArgs() has ensured this first ArgVar is always initialized, so there's never any need to check mArgc > 0.
	#define ARGVARRAW1 (*sArgVar)   // i.e. sArgVar[0], and same as OUTPUT_VAR (it's a duplicate to help readability).
	#define ARGVARRAW2 (sArgVar[1]) // It's called RAW because its user is responsible for ensuring the arg
	#define ARGVARRAW3 (sArgVar[2]) // exists by checking mArgc at loadtime or runtime.
	#define ARGVAR1 ARGVARRAW1 // This first one doesn't need the check below because ExpandArgs() has ensured it's initialized.
	#define ARGVAR2 (mArgc > 1 ? sArgVar[1] : NULL) // Caller relies on the check of mArgc because for performance,
	#define ARGVAR3 (mArgc > 2 ? sArgVar[2] : NULL) // sArgVar[] isn't initialied for parameters the script
	#define ARGVAR4 (mArgc > 3 ? sArgVar[3] : NULL) // omitted entirely from the end of the parameter list.
	#define ARGVAR5 (mArgc > 4 ? sArgVar[4] : NULL)
	#define ARGVAR6 (mArgc > 5 ? sArgVar[5] : NULL)
	#define ARGVAR7 (mArgc > 6 ? sArgVar[6] : NULL)
	#define ARGVAR8 (mArgc > 7 ? sArgVar[7] : NULL)

	#define ARG1 sArgDeref[0] // These are the expanded/resolved parameters for the currently-executing command.
	#define ARG2 sArgDeref[1] // They're populated by ExpandArgs().
	#define ARG3 sArgDeref[2]
	#define ARG4 sArgDeref[3]
	#define ARG5 sArgDeref[4]
	#define ARG6 sArgDeref[5]
	#define ARG7 sArgDeref[6]
	#define ARG8 sArgDeref[7]
	#define ARG9 sArgDeref[8]
	#define ARG10 sArgDeref[9]
	#define ARG11 sArgDeref[10]

	#define TWO_ARGS    ARG1, ARG2
	#define THREE_ARGS  ARG1, ARG2, ARG3
	#define FOUR_ARGS   ARG1, ARG2, ARG3, ARG4
	#define FIVE_ARGS   ARG1, ARG2, ARG3, ARG4, ARG5
	#define SIX_ARGS    ARG1, ARG2, ARG3, ARG4, ARG5, ARG6
	#define SEVEN_ARGS  ARG1, ARG2, ARG3, ARG4, ARG5, ARG6, ARG7
	#define EIGHT_ARGS  ARG1, ARG2, ARG3, ARG4, ARG5, ARG6, ARG7, ARG8
	#define NINE_ARGS   ARG1, ARG2, ARG3, ARG4, ARG5, ARG6, ARG7, ARG8, ARG9
	#define TEN_ARGS    ARG1, ARG2, ARG3, ARG4, ARG5, ARG6, ARG7, ARG8, ARG9, ARG10
	#define ELEVEN_ARGS ARG1, ARG2, ARG3, ARG4, ARG5, ARG6, ARG7, ARG8, ARG9, ARG10, ARG11

	// If the arg's text is non-blank, it means the variable is a dynamic name such as array%i%
	// that must be resolved at runtime rather than load-time.  Therefore, this macro must not
	// be used without first having checked that arg.text is blank:
	#define VAR(arg) ((Var *)arg.deref)
	// Uses arg number (i.e. the first arg is 1, not 0).  Caller must ensure that ArgNum >= 1 and that
	// the arg in question is an input or output variable (since that isn't checked there).
	//#define ARG_HAS_VAR(ArgNum) (mArgc >= ArgNum && (*mArg[ArgNum-1].text || mArg[ArgNum-1].deref))

	// Shouldn't go much higher than 400 since the main window's Edit control is currently limited
	// to 64KB to be compatible with the Win9x limit.  Avg. line length is probably under 100 for
	// the vast majority of scripts, so 400 seems unlikely to exceed the buffer size.  Even in the
	// worst case where the buffer size is exceeded, the text is simply truncated, so it's not too bad:
	#define LINE_LOG_SIZE 400  // See above.
	static Line *sLog[LINE_LOG_SIZE];
	static DWORD sLogTick[LINE_LOG_SIZE];
	static int sLogNext;

#ifdef AUTOHOTKEYSC  // Reduces code size to omit things that are unused, and helps catch bugs at compile-time.
	static char *sSourceFile[1]; // Only need to be able to hold the main script since compiled scripts don't support dynamic including.
#else
	static char **sSourceFile;   // Will hold an array of strings.
	static int sMaxSourceFiles;  // Maximum number of items it can currently hold.
#endif
	static int sSourceFileCount; // Number of items in the above array.

	static void FreeDerefBufIfLarge();

	ResultType ExecUntil(ExecUntilMode aMode, char **apReturnValue = NULL, Line **apJumpToLine = NULL);

	// The following are characters that can't legally occur after an AND or OR.  It excludes all unary operators
	// "!~*&-+" as well as the parentheses chars "()":
	#define EXPR_CORE "<>=/|^,:"
	// The characters common to both EXPR_TELLTALES and EXPR_OPERAND_TERMINATORS:
	#define EXPR_COMMON " \t" EXPR_CORE "*&~!()"  // Space and Tab are included at the beginning for performance.
	#define EXPR_COMMON_FORBIDDEN_BYREF "<>/|^,*&~!" // Omits space/tab because operators like := can have them. Omits colon because want to be able to pass a ternary byref. Omits = because colon is omitted (otherwise the logic is written in a way that wouldn't allow :=). Omits parentheses because a variable or assignment can be enclosed in them even though they're redundant.
	#define CONTINUATION_LINE_SYMBOLS EXPR_CORE ".+-*&!?~" // v1.0.46.
	// Characters whose presence in a mandatory-numeric param make it an expression for certain.
	// + and - are not included here because legacy numeric parameters can contain unary plus or minus,
	// e.g. WinMove, -%x%, -%y%:
	#define EXPR_TELLTALES EXPR_COMMON "\""
	// Characters that mark the end of an operand inside an expression.  Double-quote must not be included:
	#define EXPR_OPERAND_TERMINATORS EXPR_COMMON "+-"
	#define EXPR_ALL_SYMBOLS EXPR_OPERAND_TERMINATORS "\"" // Excludes '.' and '?' since they need special treatment due to the present/future allowance of them inside the names of variable and functions.
	#define EXPR_FORBIDDEN_BYREF EXPR_COMMON_FORBIDDEN_BYREF ".+-\"" // Dot is also included.
	#define EXPR_ILLEGAL_CHARS "'\\;`{}" // Characters illegal in an expression.
	// The following HOTSTRING option recognizer is kept somewhat forgiving/non-specific for backward compatibility
	// (e.g. scripts may have some invalid hotstring options, which are simply ignored).  This definition is here
	// because it's related to continuation line symbols. Also, avoid ever adding "&" to hotstring options because
	// it might introduce ambiguity in the differentiation of things like:
	//    : & x::hotkey action
	//    : & *::abbrev with leading colon::
	#define IS_HOTSTRING_OPTION(chr) (isalnum(chr) || strchr("?*- \t", chr))
	// The characters below are ordered with most-often used ones first, for performance:
	#define DEFINE_END_FLAGS \
		char end_flags[] = {' ', g_delimiter, '(', '\t', '<', '>', ':', '=', '+', '-', '*', '/', '!', '~', '&', '|', '^', '\0'}; // '\0' must be last.
		// '?' and '.' are omitted from the above because they require special handling due to being permitted
		// in the curruent or future names of variables and functions.
	static bool StartsWithAssignmentOp(char *aStr) // RELATED TO ABOVE, so kept adjacent to it.
	// Returns true if aStr begins with an assignment operator such as :=, >>=, ++, etc.
	// For simplicity, this doesn't check that what comes AFTER an operator is valid.  For example,
	// :== isn't valid, yet is reported as valid here because it starts with :=.
	// Caller is responsible for having omitted leading whitespace, if desired.
	{
		if (!(*aStr && aStr[1])) // Relies on short-circuit boolean order.
			return false;
		char cp0 = *aStr;
		switch(aStr[1])
		{
		// '=' is listed first for performance, since it's the most common.
		case '=': return strchr(":+-*.|&^/", cp0); // Covers :=, +=, -=, *=, .=, |=, &=, ^=, /= (9 operators).
		case '+': // Fall through to below. Covers ++.
		case '-': return cp0 == aStr[1]; // Covers --.
		case '/': // Fall through to below. Covers //=.
		case '>': // Fall through to below. covers >>=.
		case '<': return cp0 == aStr[1] && aStr[2] == '='; // Covers <<=.
		}
		// Otherwise:
		return false;
	}

	#define ArgLength(aArgNum) ArgIndexLength((aArgNum)-1)
	#define ArgToDouble(aArgNum) ArgIndexToDouble((aArgNum)-1)
	#define ArgToInt64(aArgNum) ArgIndexToInt64((aArgNum)-1)
	#define ArgToInt(aArgNum) (int)ArgToInt64(aArgNum) // Benchmarks show that having a "real" ArgToInt() that calls ATOI vs. ATOI64 (and ToInt() vs. ToInt64()) doesn't measurably improve performance.
	#define ArgToUInt(aArgNum) (UINT)ArgToInt64(aArgNum) // Similar to what ATOU() does.
	__int64 ArgIndexToInt64(int aArgIndex);
	double ArgIndexToDouble(int aArgIndex);
	size_t ArgIndexLength(int aArgIndex);

	Var *ResolveVarOfArg(int aArgIndex, bool aCreateIfNecessary = true);
	ResultType ExpandArgs(VarSizeType aSpaceNeeded = VARSIZE_ERROR, Var *aArgVar[] = NULL);
	VarSizeType GetExpandedArgSize(Var *aArgVar[]);
	char *ExpandArg(char *aBuf, int aArgIndex, Var *aArgVar = NULL);
	char *ExpandExpression(int aArgIndex, ResultType &aResult, char *&aTarget, char *&aDerefBuf
		, size_t &aDerefBufSize, char *aArgDeref[], size_t aExtraSize);
	ResultType ExpressionToPostfix(ArgStruct &aArg);

	ResultType Deref(Var *aOutputVar, char *aBuf);

	static bool FileIsFilteredOut(WIN32_FIND_DATA &aCurrentFile, FileLoopModeType aFileLoopMode
		, char *aFilePath, size_t aFilePathLength);

	Label *GetJumpTarget(bool aIsDereferenced);
	Label *IsJumpValid(Label &aTargetLabel);

	HWND DetermineTargetWindow(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);

#ifndef AUTOHOTKEYSC
	static int ConvertEscapeChar(char *aFilespec);
	static size_t ConvertEscapeCharGetLine(char *aBuf, int aMaxCharsToRead, FILE *fp);
#endif  // The functions above are not needed by the self-contained version.

	
	// This is in the .h file so that it's more likely the compiler's cost/benefit estimate will
	// make it inline (since it is called from only one place).  Inline would be good since it
	// is called frequently during script loading and is difficult to macro-ize in a way that
	// retains readability.
	static ArgTypeType ArgIsVar(ActionTypeType aActionType, int aArgIndex)
	{
		switch(aArgIndex)
		{
		case 0:  // Arg #1
			switch(aActionType)
			{
			case ACT_ASSIGN:
			case ACT_ASSIGNEXPR:
			case ACT_ADD:
			case ACT_SUB:
			case ACT_MULT:
			case ACT_DIV:
			case ACT_TRANSFORM:
			case ACT_STRINGLEFT:
			case ACT_STRINGRIGHT:
			case ACT_STRINGMID:
			case ACT_STRINGTRIMLEFT:
			case ACT_STRINGTRIMRIGHT:
			case ACT_STRINGLOWER:
			case ACT_STRINGUPPER:
			case ACT_STRINGLEN:
			case ACT_STRINGREPLACE:
			case ACT_STRINGGETPOS:
			case ACT_GETKEYSTATE:
			case ACT_CONTROLGETFOCUS:
			case ACT_CONTROLGETTEXT:
			case ACT_CONTROLGET:
			case ACT_GUICONTROLGET:
			case ACT_STATUSBARGETTEXT:
			case ACT_INPUTBOX:
			case ACT_RANDOM:
			case ACT_INIREAD:
			case ACT_REGREAD:
			case ACT_DRIVESPACEFREE:
			case ACT_DRIVEGET:
			case ACT_SOUNDGET:
			case ACT_SOUNDGETWAVEVOLUME:
			case ACT_FILEREAD:
			case ACT_FILEREADLINE:
			case ACT_FILEGETATTRIB:
			case ACT_FILEGETTIME:
			case ACT_FILEGETSIZE:
			case ACT_FILEGETVERSION:
			case ACT_FILESELECTFILE:
			case ACT_FILESELECTFOLDER:
			case ACT_MOUSEGETPOS:
			case ACT_WINGETTITLE:
			case ACT_WINGETCLASS:
			case ACT_WINGET:
			case ACT_WINGETTEXT:
			case ACT_WINGETPOS:
			case ACT_SYSGET:
			case ACT_ENVGET:
			case ACT_CONTROLGETPOS:
			case ACT_PIXELGETCOLOR:
			case ACT_PIXELSEARCH:
			case ACT_IMAGESEARCH:
			case ACT_INPUT:
			case ACT_FORMATTIME:
				return ARG_TYPE_OUTPUT_VAR;

			case ACT_SORT:
			case ACT_SPLITPATH:
			case ACT_IFINSTRING:
			case ACT_IFNOTINSTRING:
			case ACT_IFEQUAL:
			case ACT_IFNOTEQUAL:
			case ACT_IFGREATER:
			case ACT_IFGREATEROREQUAL:
			case ACT_IFLESS:
			case ACT_IFLESSOREQUAL:
			case ACT_IFBETWEEN:
			case ACT_IFNOTBETWEEN:
			case ACT_IFIN:
			case ACT_IFNOTIN:
			case ACT_IFCONTAINS:
			case ACT_IFNOTCONTAINS:
			case ACT_IFIS:
			case ACT_IFISNOT:
				return ARG_TYPE_INPUT_VAR;
			}
			break;

		case 1:  // Arg #2
			switch(aActionType)
			{
			case ACT_STRINGLEFT:
			case ACT_STRINGRIGHT:
			case ACT_STRINGMID:
			case ACT_STRINGTRIMLEFT:
			case ACT_STRINGTRIMRIGHT:
			case ACT_STRINGLOWER:
			case ACT_STRINGUPPER:
			case ACT_STRINGLEN:
			case ACT_STRINGREPLACE:
			case ACT_STRINGGETPOS:
			case ACT_STRINGSPLIT:
				return ARG_TYPE_INPUT_VAR;

			case ACT_MOUSEGETPOS:
			case ACT_WINGETPOS:
			case ACT_CONTROLGETPOS:
			case ACT_PIXELSEARCH:
			case ACT_IMAGESEARCH:
			case ACT_SPLITPATH:
			case ACT_FILEGETSHORTCUT:
				return ARG_TYPE_OUTPUT_VAR;
			}
			break;

		case 2:  // Arg #3
			switch(aActionType)
			{
			case ACT_WINGETPOS:
			case ACT_CONTROLGETPOS:
			case ACT_MOUSEGETPOS:
			case ACT_SPLITPATH:
			case ACT_FILEGETSHORTCUT:
				return ARG_TYPE_OUTPUT_VAR;
			}
			break;

		case 3:  // Arg #4
			switch(aActionType)
			{
			case ACT_WINGETPOS:
			case ACT_CONTROLGETPOS:
			case ACT_MOUSEGETPOS:
			case ACT_SPLITPATH:
			case ACT_FILEGETSHORTCUT:
			case ACT_RUN:
			case ACT_RUNWAIT:
				return ARG_TYPE_OUTPUT_VAR;
			}
			break;

		case 4:  // Arg #5
		case 5:  // Arg #6
			if (aActionType == ACT_SPLITPATH || aActionType == ACT_FILEGETSHORTCUT)
				return ARG_TYPE_OUTPUT_VAR;
			break;

		case 6:  // Arg #7
		case 7:  // Arg #8
			if (aActionType == ACT_FILEGETSHORTCUT)
				return ARG_TYPE_OUTPUT_VAR;
		}
		// Otherwise:
		return ARG_TYPE_NORMAL;
	}



	ResultType ArgMustBeDereferenced(Var *aVar, int aArgIndex, Var *aArgVar[]);

	#define ArgHasDeref(aArgNum) ArgIndexHasDeref((aArgNum)-1)
	bool ArgIndexHasDeref(int aArgIndex)
	// This function should always be called in lieu of doing something like "strchr(arg.text, g_DerefChar)"
	// because that method is unreliable due to the possible presence of literal (escaped) g_DerefChars
	// in the text.
	// Caller must ensure that aArgIndex is 0 or greater.
	{
#ifdef _DEBUG
		if (aArgIndex < 0)
		{
			LineError("DEBUG: BAD", WARN);
			aArgIndex = 0;  // But let it continue.
		}
#endif
		if (aArgIndex >= mArgc) // Arg doesn't exist.
			return false;
		ArgStruct &arg = mArg[aArgIndex]; // For performance.
		// Return false if it's not of a type caller wants deemed to have derefs.
		if (arg.type == ARG_TYPE_NORMAL)
			return arg.deref && arg.deref[0].marker; // Relies on short-circuit boolean evaluation order to prevent NULL-deref.
		else // Callers rely on input variables being seen as "true" because sometimes single isolated derefs are converted into ARG_TYPE_INPUT_VAR at load-time.
			return (arg.type == ARG_TYPE_INPUT_VAR);
	}

	static HKEY RegConvertRootKey(char *aBuf, bool *aIsRemoteRegistry = NULL)
	{
		// Even if the computer name is a single letter, it seems like using a colon as delimiter is ok
		// (e.g. a:HKEY_LOCAL_MACHINE), since we wouldn't expect the root key to be used as a filename
		// in that exact way, i.e. a drive letter should be followed by a backslash 99% of the time in
		// this context.
		// Research indicates that colon is an illegal char in a computer name (at least for NT,
		// and by extension probably all other OSes).  So it should be safe to use it as a delimiter
		// for the remote registry feature.  But just in case, get the right-most one,
		// e.g. Computer:01:HKEY_LOCAL_MACHINE  ; the first colon is probably illegal on all OSes.
		// Additional notes from the Internet:
		// "A Windows NT computer name can be up to 15 alphanumeric characters with no blank spaces
		// and must be unique on the network. It can contain the following special characters:
		// ! @ # $ % ^ & ( ) -   ' { } .
		// It may not contain:
		// \ * + = | : ; " ? ,
		// The following is a list of illegal characters in a computer name:
		// regEx.Pattern = "`|~|!|@|#|\$|\^|\&|\*|\(|\)|\=|\+|{|}|\\|;|:|'|<|>|/|\?|\||%"

		char *colon_pos = strrchr(aBuf, ':');
		char *key_name = colon_pos ? omit_leading_whitespace(colon_pos + 1) : aBuf;
		if (aIsRemoteRegistry) // Caller wanted the below put into the output parameter.
			*aIsRemoteRegistry = (colon_pos != NULL);
		HKEY root_key = NULL; // Set default.
		if (!stricmp(key_name, "HKLM") || !stricmp(key_name, "HKEY_LOCAL_MACHINE"))       root_key = HKEY_LOCAL_MACHINE;
		else if (!stricmp(key_name, "HKCR") || !stricmp(key_name, "HKEY_CLASSES_ROOT"))   root_key = HKEY_CLASSES_ROOT;
		else if (!stricmp(key_name, "HKCC") || !stricmp(key_name, "HKEY_CURRENT_CONFIG")) root_key = HKEY_CURRENT_CONFIG;
		else if (!stricmp(key_name, "HKCU") || !stricmp(key_name, "HKEY_CURRENT_USER"))   root_key = HKEY_CURRENT_USER;
		else if (!stricmp(key_name, "HKU") || !stricmp(key_name, "HKEY_USERS"))           root_key = HKEY_USERS;
		if (!root_key)  // Invalid or unsupported root key name.
			return NULL;
		if (!aIsRemoteRegistry || !colon_pos) // Either caller didn't want it opened, or it doesn't need to be.
			return root_key; // If it's a remote key, this value should only be used by the caller as an indicator.
		// Otherwise, it's a remote computer whose registry the caller wants us to open:
		// It seems best to require the two leading backslashes in case the computer name contains
		// spaces (just in case spaces are allowed on some OSes or perhaps for Unix interoperability, etc.).
		// Therefore, make no attempt to trim leading and trailing spaces from the computer name:
		char computer_name[128];
		strlcpy(computer_name, aBuf, sizeof(computer_name));
		computer_name[colon_pos - aBuf] = '\0';
		HKEY remote_key;
		return (RegConnectRegistry(computer_name, root_key, &remote_key) == ERROR_SUCCESS) ? remote_key : NULL;
	}

	static char *RegConvertRootKey(char *aBuf, size_t aBufSize, HKEY aRootKey)
	{
		// switch() doesn't directly support expression of type HKEY:
		if (aRootKey == HKEY_LOCAL_MACHINE)       strlcpy(aBuf, "HKEY_LOCAL_MACHINE", aBufSize);
		else if (aRootKey == HKEY_CLASSES_ROOT)   strlcpy(aBuf, "HKEY_CLASSES_ROOT", aBufSize);
		else if (aRootKey == HKEY_CURRENT_CONFIG) strlcpy(aBuf, "HKEY_CURRENT_CONFIG", aBufSize);
		else if (aRootKey == HKEY_CURRENT_USER)   strlcpy(aBuf, "HKEY_CURRENT_USER", aBufSize);
		else if (aRootKey == HKEY_USERS)          strlcpy(aBuf, "HKEY_USERS", aBufSize);
		else if (aBufSize)                        *aBuf = '\0'; // Make it be the empty string for anything else.
		// These are either unused or so rarely used (DYN_DATA on Win9x) that they aren't supported:
		// HKEY_PERFORMANCE_DATA, HKEY_PERFORMANCE_TEXT, HKEY_PERFORMANCE_NLSTEXT, HKEY_DYN_DATA
		return aBuf;
	}
	static int RegConvertValueType(char *aValueType)
	{
		if (!stricmp(aValueType, "REG_SZ")) return REG_SZ;
		if (!stricmp(aValueType, "REG_EXPAND_SZ")) return REG_EXPAND_SZ;
		if (!stricmp(aValueType, "REG_MULTI_SZ")) return REG_MULTI_SZ;
		if (!stricmp(aValueType, "REG_DWORD")) return REG_DWORD;
		if (!stricmp(aValueType, "REG_BINARY")) return REG_BINARY;
		return REG_NONE; // Unknown or unsupported type.
	}
	static char *RegConvertValueType(char *aBuf, size_t aBufSize, DWORD aValueType)
	{
		switch(aValueType)
		{
		case REG_SZ: strlcpy(aBuf, "REG_SZ", aBufSize); return aBuf;
		case REG_EXPAND_SZ: strlcpy(aBuf, "REG_EXPAND_SZ", aBufSize); return aBuf;
		case REG_BINARY: strlcpy(aBuf, "REG_BINARY", aBufSize); return aBuf;
		case REG_DWORD: strlcpy(aBuf, "REG_DWORD", aBufSize); return aBuf;
		case REG_DWORD_BIG_ENDIAN: strlcpy(aBuf, "REG_DWORD_BIG_ENDIAN", aBufSize); return aBuf;
		case REG_LINK: strlcpy(aBuf, "REG_LINK", aBufSize); return aBuf;
		case REG_MULTI_SZ: strlcpy(aBuf, "REG_MULTI_SZ", aBufSize); return aBuf;
		case REG_RESOURCE_LIST: strlcpy(aBuf, "REG_RESOURCE_LIST", aBufSize); return aBuf;
		case REG_FULL_RESOURCE_DESCRIPTOR: strlcpy(aBuf, "REG_FULL_RESOURCE_DESCRIPTOR", aBufSize); return aBuf;
		case REG_RESOURCE_REQUIREMENTS_LIST: strlcpy(aBuf, "REG_RESOURCE_REQUIREMENTS_LIST", aBufSize); return aBuf;
		case REG_QWORD: strlcpy(aBuf, "REG_QWORD", aBufSize); return aBuf;
		case REG_SUBKEY: strlcpy(aBuf, "KEY", aBufSize); return aBuf;  // Custom (non-standard) type.
		default: if (aBufSize) *aBuf = '\0'; return aBuf;  // Make it be the empty string for REG_NONE and anything else.
		}
	}

	static DWORD SoundConvertComponentType(char *aBuf, int *aInstanceNumber = NULL)
	{
		char *colon_pos = strchr(aBuf, ':');
		UINT length_to_check = (UINT)(colon_pos ? colon_pos - aBuf : strlen(aBuf));
		if (aInstanceNumber) // Caller wanted the below put into the output parameter.
		{
			if (colon_pos)
			{
				*aInstanceNumber = ATOI(colon_pos + 1);
				if (*aInstanceNumber < 1)
					*aInstanceNumber = 1;
			}
			else
				*aInstanceNumber = 1;
		}
		if (!strlicmp(aBuf, "Master", length_to_check)
			|| !strlicmp(aBuf, "Speakers", length_to_check))   return MIXERLINE_COMPONENTTYPE_DST_SPEAKERS;
		if (!strlicmp(aBuf, "Headphones", length_to_check))    return MIXERLINE_COMPONENTTYPE_DST_HEADPHONES;
		if (!strlicmp(aBuf, "Digital", length_to_check))       return MIXERLINE_COMPONENTTYPE_SRC_DIGITAL;
		if (!strlicmp(aBuf, "Line", length_to_check))          return MIXERLINE_COMPONENTTYPE_SRC_LINE;
		if (!strlicmp(aBuf, "Microphone", length_to_check))    return MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE;
		if (!strlicmp(aBuf, "Synth", length_to_check))         return MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER;
		if (!strlicmp(aBuf, "CD", length_to_check))            return MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC;
		if (!strlicmp(aBuf, "Telephone", length_to_check))     return MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE;
		if (!strlicmp(aBuf, "PCSpeaker", length_to_check))     return MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER;
		if (!strlicmp(aBuf, "Wave", length_to_check))          return MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT;
		if (!strlicmp(aBuf, "Aux", length_to_check))           return MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY;
		if (!strlicmp(aBuf, "Analog", length_to_check))        return MIXERLINE_COMPONENTTYPE_SRC_ANALOG;
		// v1.0.37.06: The following was added because it's legitimate on some sound cards such as
		// SB Audigy's recording (dest #2) Wave/Mp3 volume:
		if (!strlicmp(aBuf, "N/A", length_to_check))           return MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED; // 0x1000
		return MIXERLINE_COMPONENTTYPE_DST_UNDEFINED; // Zero.
	}
	static DWORD SoundConvertControlType(char *aBuf)
	{
		// v1.0.37.06: The following was added to allow unnamed control types (if any) to be accessed via number:
		if (IsPureNumeric(aBuf, false, false, true)) // Seems best to allowing floating point here, since .000 on the end might happen sometimes.
			return ATOU(aBuf);
		// The following are the types that seem to correspond to actual sound attributes.  Some of the
		// values are not included here, such as MIXERCONTROL_CONTROLTYPE_FADER, which seems to be a type
		// of sound control rather than a quality of the sound itself.  For performance, put the most
		// often used ones up top.
		if (!stricmp(aBuf, "Vol")
			|| !stricmp(aBuf, "Volume")) return MIXERCONTROL_CONTROLTYPE_VOLUME;
		if (!stricmp(aBuf, "OnOff"))     return MIXERCONTROL_CONTROLTYPE_ONOFF;
		if (!stricmp(aBuf, "Mute"))      return MIXERCONTROL_CONTROLTYPE_MUTE;
		if (!stricmp(aBuf, "Mono"))      return MIXERCONTROL_CONTROLTYPE_MONO;
		if (!stricmp(aBuf, "Loudness"))  return MIXERCONTROL_CONTROLTYPE_LOUDNESS;
		if (!stricmp(aBuf, "StereoEnh")) return MIXERCONTROL_CONTROLTYPE_STEREOENH;
		if (!stricmp(aBuf, "BassBoost")) return MIXERCONTROL_CONTROLTYPE_BASS_BOOST;
		if (!stricmp(aBuf, "Pan"))       return MIXERCONTROL_CONTROLTYPE_PAN;
		if (!stricmp(aBuf, "QSoundPan")) return MIXERCONTROL_CONTROLTYPE_QSOUNDPAN;
		if (!stricmp(aBuf, "Bass"))      return MIXERCONTROL_CONTROLTYPE_BASS;
		if (!stricmp(aBuf, "Treble"))    return MIXERCONTROL_CONTROLTYPE_TREBLE;
		if (!stricmp(aBuf, "Equalizer")) return MIXERCONTROL_CONTROLTYPE_EQUALIZER;
		#define MIXERCONTROL_CONTROLTYPE_INVALID 0xFFFFFFFF // 0 might be a valid type, so use something definitely undefined.
		return MIXERCONTROL_CONTROLTYPE_INVALID;
	}

	static TitleMatchModes ConvertTitleMatchMode(char *aBuf)
	{
		if (!aBuf || !*aBuf) return MATCHMODE_INVALID;
		if (*aBuf == '1' && !*(aBuf + 1)) return FIND_IN_LEADING_PART;
		if (*aBuf == '2' && !*(aBuf + 1)) return FIND_ANYWHERE;
		if (*aBuf == '3' && !*(aBuf + 1)) return FIND_EXACT;
		if (!stricmp(aBuf, "RegEx")) return FIND_REGEX; // Goes with the above, not fast/slow below.

		if (!stricmp(aBuf, "FAST")) return FIND_FAST;
		if (!stricmp(aBuf, "SLOW")) return FIND_SLOW;
		return MATCHMODE_INVALID;
	}

	static SysGetCmds ConvertSysGetCmd(char *aBuf)
	{
		if (!aBuf || !*aBuf) return SYSGET_CMD_INVALID;
		if (IsPureNumeric(aBuf)) return SYSGET_CMD_METRICS;
		if (!stricmp(aBuf, "MonitorCount")) return SYSGET_CMD_MONITORCOUNT;
		if (!stricmp(aBuf, "MonitorPrimary")) return SYSGET_CMD_MONITORPRIMARY;
		if (!stricmp(aBuf, "Monitor")) return SYSGET_CMD_MONITORAREA; // Called "Monitor" vs. "MonitorArea" to make it easier to remember.
		if (!stricmp(aBuf, "MonitorWorkArea")) return SYSGET_CMD_MONITORWORKAREA;
		if (!stricmp(aBuf, "MonitorName")) return SYSGET_CMD_MONITORNAME;
		return SYSGET_CMD_INVALID;
	}

	static TransformCmds ConvertTransformCmd(char *aBuf)
	{
		if (!aBuf || !*aBuf) return TRANS_CMD_INVALID;
		if (!stricmp(aBuf, "Asc")) return TRANS_CMD_ASC;
		if (!stricmp(aBuf, "Chr")) return TRANS_CMD_CHR;
		if (!stricmp(aBuf, "Deref")) return TRANS_CMD_DEREF;
		if (!stricmp(aBuf, "Unicode")) return TRANS_CMD_UNICODE;
		if (!stricmp(aBuf, "HTML")) return TRANS_CMD_HTML;
		if (!stricmp(aBuf, "Mod")) return TRANS_CMD_MOD;
		if (!stricmp(aBuf, "Pow")) return TRANS_CMD_POW;
		if (!stricmp(aBuf, "Exp")) return TRANS_CMD_EXP;
		if (!stricmp(aBuf, "Sqrt")) return TRANS_CMD_SQRT;
		if (!stricmp(aBuf, "Log")) return TRANS_CMD_LOG;
		if (!stricmp(aBuf, "Ln")) return TRANS_CMD_LN;  // Natural log.
		if (!stricmp(aBuf, "Round")) return TRANS_CMD_ROUND;
		if (!stricmp(aBuf, "Ceil")) return TRANS_CMD_CEIL;
		if (!stricmp(aBuf, "Floor")) return TRANS_CMD_FLOOR;
		if (!stricmp(aBuf, "Abs")) return TRANS_CMD_ABS;
		if (!stricmp(aBuf, "Sin")) return TRANS_CMD_SIN;
		if (!stricmp(aBuf, "Cos")) return TRANS_CMD_COS;
		if (!stricmp(aBuf, "Tan")) return TRANS_CMD_TAN;
		if (!stricmp(aBuf, "ASin")) return TRANS_CMD_ASIN;
		if (!stricmp(aBuf, "ACos")) return TRANS_CMD_ACOS;
		if (!stricmp(aBuf, "ATan")) return TRANS_CMD_ATAN;
		if (!stricmp(aBuf, "BitAnd")) return TRANS_CMD_BITAND;
		if (!stricmp(aBuf, "BitOr")) return TRANS_CMD_BITOR;
		if (!stricmp(aBuf, "BitXOr")) return TRANS_CMD_BITXOR;
		if (!stricmp(aBuf, "BitNot")) return TRANS_CMD_BITNOT;
		if (!stricmp(aBuf, "BitShiftLeft")) return TRANS_CMD_BITSHIFTLEFT;
		if (!stricmp(aBuf, "BitShiftRight")) return TRANS_CMD_BITSHIFTRIGHT;
		return TRANS_CMD_INVALID;
	}

	static MenuCommands ConvertMenuCommand(char *aBuf)
	{
		if (!aBuf || !*aBuf) return MENU_CMD_INVALID;
		if (!stricmp(aBuf, "Show")) return MENU_CMD_SHOW;
		if (!stricmp(aBuf, "UseErrorLevel")) return MENU_CMD_USEERRORLEVEL;
		if (!stricmp(aBuf, "Add")) return MENU_CMD_ADD;
		if (!stricmp(aBuf, "Rename")) return MENU_CMD_RENAME;
		if (!stricmp(aBuf, "Check")) return MENU_CMD_CHECK;
		if (!stricmp(aBuf, "Uncheck")) return MENU_CMD_UNCHECK;
		if (!stricmp(aBuf, "ToggleCheck")) return MENU_CMD_TOGGLECHECK;
		if (!stricmp(aBuf, "Enable")) return MENU_CMD_ENABLE;
		if (!stricmp(aBuf, "Disable")) return MENU_CMD_DISABLE;
		if (!stricmp(aBuf, "ToggleEnable")) return MENU_CMD_TOGGLEENABLE;
		if (!stricmp(aBuf, "Standard")) return MENU_CMD_STANDARD;
		if (!stricmp(aBuf, "NoStandard")) return MENU_CMD_NOSTANDARD;
		if (!stricmp(aBuf, "Color")) return MENU_CMD_COLOR;
		if (!stricmp(aBuf, "Default")) return MENU_CMD_DEFAULT;
		if (!stricmp(aBuf, "NoDefault")) return MENU_CMD_NODEFAULT;
		if (!stricmp(aBuf, "Delete")) return MENU_CMD_DELETE;
		if (!stricmp(aBuf, "DeleteAll")) return MENU_CMD_DELETEALL;
		if (!stricmp(aBuf, "Tip")) return MENU_CMD_TIP;
		if (!stricmp(aBuf, "Icon")) return MENU_CMD_ICON;
		if (!stricmp(aBuf, "NoIcon")) return MENU_CMD_NOICON;
		if (!stricmp(aBuf, "Click")) return MENU_CMD_CLICK;
		if (!stricmp(aBuf, "MainWindow")) return MENU_CMD_MAINWINDOW;
		if (!stricmp(aBuf, "NoMainWindow")) return MENU_CMD_NOMAINWINDOW;
		return MENU_CMD_INVALID;
	}

	static GuiCommands ConvertGuiCommand(char *aBuf, int *aWindowIndex = NULL, char **aOptions = NULL)
	{
		// Notes about the below macro:
		// "< 3" avoids ambiguity with a future use such as "gui +cmd:whatever" while still allowing
		// up to 99 windows, e.g. "gui 99:add"
		// omit_leading_whitespace(): Move the buf pointer to the location of the sub-command.
		#define DETERMINE_WINDOW_INDEX \
			char *colon_pos = strchr(aBuf, ':');\
			if (colon_pos && colon_pos - aBuf < 3)\
			{\
				if (aWindowIndex)\
					*aWindowIndex = ATOI(aBuf) - 1;\
				aBuf = omit_leading_whitespace(colon_pos + 1);\
			}
			//else leave it set to the default already put in it by the caller.
		DETERMINE_WINDOW_INDEX
		if (aOptions)
			*aOptions = aBuf; // Return position where options start to the caller.
		if (!*aBuf || *aBuf == '+' || *aBuf == '-') // Assume a var ref that resolves to blank is "options" (for runtime flexibility).
			return GUI_CMD_OPTIONS;
		if (!stricmp(aBuf, "Add")) return GUI_CMD_ADD;
		if (!stricmp(aBuf, "Show")) return GUI_CMD_SHOW;
		if (!stricmp(aBuf, "Submit")) return GUI_CMD_SUBMIT;
		if (!stricmp(aBuf, "Cancel") || !stricmp(aBuf, "Hide")) return GUI_CMD_CANCEL;
		if (!stricmp(aBuf, "Minimize")) return GUI_CMD_MINIMIZE;
		if (!stricmp(aBuf, "Maximize")) return GUI_CMD_MAXIMIZE;
		if (!stricmp(aBuf, "Restore")) return GUI_CMD_RESTORE;
		if (!stricmp(aBuf, "Destroy")) return GUI_CMD_DESTROY;
		if (!stricmp(aBuf, "Margin")) return GUI_CMD_MARGIN;
		if (!stricmp(aBuf, "Menu")) return GUI_CMD_MENU;
		if (!stricmp(aBuf, "Font")) return GUI_CMD_FONT;
		if (!stricmp(aBuf, "Tab")) return GUI_CMD_TAB;
		if (!stricmp(aBuf, "ListView")) return GUI_CMD_LISTVIEW;
		if (!stricmp(aBuf, "TreeView")) return GUI_CMD_TREEVIEW;
		if (!stricmp(aBuf, "Default")) return GUI_CMD_DEFAULT;
		if (!stricmp(aBuf, "Color")) return GUI_CMD_COLOR;
		if (!stricmp(aBuf, "Flash")) return GUI_CMD_FLASH;
		return GUI_CMD_INVALID;
	}

	GuiControlCmds ConvertGuiControlCmd(char *aBuf, int *aWindowIndex = NULL, char **aOptions = NULL)
	{
		DETERMINE_WINDOW_INDEX
		if (aOptions)
			*aOptions = aBuf; // Return position where options start to the caller.
		// If it's blank without a deref, that's CONTENTS.  Otherwise, assume it's OPTIONS for better
		// runtime flexibility (i.e. user can leave the variable blank to make the command do nothing).
		// Fix for v1.0.40.11: Since the above is counterintuitive and undocumented, it has been fixed
		// to behave the way most users would expect; that is, the contents of any deref in parameter 1
		// will behave the same as when such contents is present literally as parametter 1.  Another
		// reason for doing this is that otherwise, there is no way to specify the CONTENTS sub-command
		// in a variable.  For example, the following wouldn't work:
		// GuiControl, %WindowNumber%:, ...
		// GuiControl, %WindowNumberWithColon%, ...
		if (!*aBuf)
			return GUICONTROL_CMD_CONTENTS;
		if (*aBuf == '+' || *aBuf == '-') // Assume a var ref that resolves to blank is "options" (for runtime flexibility).
			return GUICONTROL_CMD_OPTIONS;
		if (!stricmp(aBuf, "Text")) return GUICONTROL_CMD_TEXT;
		if (!stricmp(aBuf, "Move")) return GUICONTROL_CMD_MOVE;
		if (!stricmp(aBuf, "MoveDraw")) return GUICONTROL_CMD_MOVEDRAW;
		if (!stricmp(aBuf, "Focus")) return GUICONTROL_CMD_FOCUS;
		if (!stricmp(aBuf, "Choose")) return GUICONTROL_CMD_CHOOSE;
		if (!stricmp(aBuf, "ChooseString")) return GUICONTROL_CMD_CHOOSESTRING;
		if (!stricmp(aBuf, "Font")) return GUICONTROL_CMD_FONT;

		// v1.0.38.02: Anything not already returned from above supports an optional boolean suffix.
		// The following example would hide the control: GuiControl, Show%VarContainingFalse%, MyControl
		// To support hex (due to the 'x' in it), search from the left rather than the right for the
		// first digit:
		char *suffix;
		for (suffix = aBuf; *suffix && !isdigit(*suffix); ++suffix);
		bool invert = (*suffix ? !ATOI(suffix) : false);
		if (!strnicmp(aBuf, "Enable", 6)) return invert ? GUICONTROL_CMD_DISABLE : GUICONTROL_CMD_ENABLE;
		if (!strnicmp(aBuf, "Disable", 7)) return invert ? GUICONTROL_CMD_ENABLE : GUICONTROL_CMD_DISABLE;
		if (!strnicmp(aBuf, "Show", 4)) return invert ? GUICONTROL_CMD_HIDE : GUICONTROL_CMD_SHOW;
		if (!strnicmp(aBuf, "Hide", 4)) return invert ? GUICONTROL_CMD_SHOW : GUICONTROL_CMD_HIDE;

		return GUICONTROL_CMD_INVALID;
	}

	static GuiControlGetCmds ConvertGuiControlGetCmd(char *aBuf, int *aWindowIndex = NULL)
	{
		DETERMINE_WINDOW_INDEX
		if (!*aBuf) return GUICONTROLGET_CMD_CONTENTS; // The implicit command when nothing was specified.
		if (!stricmp(aBuf, "Pos")) return GUICONTROLGET_CMD_POS;
		if (!stricmp(aBuf, "Focus")) return GUICONTROLGET_CMD_FOCUS;
		if (!stricmp(aBuf, "FocusV")) return GUICONTROLGET_CMD_FOCUSV; // Returns variable vs. ClassNN.
		if (!stricmp(aBuf, "Enabled")) return GUICONTROLGET_CMD_ENABLED;
		if (!stricmp(aBuf, "Visible")) return GUICONTROLGET_CMD_VISIBLE;
		if (!stricmp(aBuf, "Hwnd")) return GUICONTROLGET_CMD_HWND;
		return GUICONTROLGET_CMD_INVALID;
	}

	static GuiControls ConvertGuiControl(char *aBuf)
	{
		if (!aBuf || !*aBuf) return GUI_CONTROL_INVALID;
		if (!stricmp(aBuf, "Text")) return GUI_CONTROL_TEXT;
		if (!stricmp(aBuf, "Edit")) return GUI_CONTROL_EDIT;
		if (!stricmp(aBuf, "Button")) return GUI_CONTROL_BUTTON;
		if (!stricmp(aBuf, "Checkbox")) return GUI_CONTROL_CHECKBOX;
		if (!stricmp(aBuf, "Radio")) return GUI_CONTROL_RADIO;
		if (!stricmp(aBuf, "DDL") || !stricmp(aBuf, "DropDownList")) return GUI_CONTROL_DROPDOWNLIST;
		if (!stricmp(aBuf, "ComboBox")) return GUI_CONTROL_COMBOBOX;
		if (!stricmp(aBuf, "ListBox")) return GUI_CONTROL_LISTBOX;
		if (!stricmp(aBuf, "ListView")) return GUI_CONTROL_LISTVIEW;
		if (!stricmp(aBuf, "TreeView")) return GUI_CONTROL_TREEVIEW;
		// Keep those seldom used at the bottom for performance:
		if (!stricmp(aBuf, "UpDown")) return GUI_CONTROL_UPDOWN;
		if (!stricmp(aBuf, "Slider")) return GUI_CONTROL_SLIDER;
		if (!stricmp(aBuf, "Progress")) return GUI_CONTROL_PROGRESS;
		if (!stricmp(aBuf, "Tab")) return GUI_CONTROL_TAB;
		if (!stricmp(aBuf, "Tab2")) return GUI_CONTROL_TAB2; // v1.0.47.05: Used only temporarily: becomes TAB vs. TAB2 upon creation.
		if (!stricmp(aBuf, "GroupBox")) return GUI_CONTROL_GROUPBOX;
		if (!stricmp(aBuf, "Pic") || !stricmp(aBuf, "Picture")) return GUI_CONTROL_PIC;
		if (!stricmp(aBuf, "DateTime")) return GUI_CONTROL_DATETIME;
		if (!stricmp(aBuf, "MonthCal")) return GUI_CONTROL_MONTHCAL;
		if (!stricmp(aBuf, "Hotkey")) return GUI_CONTROL_HOTKEY;
		if (!stricmp(aBuf, "StatusBar")) return GUI_CONTROL_STATUSBAR;
		return GUI_CONTROL_INVALID;
	}

	static ThreadCommands ConvertThreadCommand(char *aBuf)
	{
		if (!aBuf || !*aBuf) return THREAD_CMD_INVALID;
		if (!stricmp(aBuf, "Priority")) return THREAD_CMD_PRIORITY;
		if (!stricmp(aBuf, "Interrupt")) return THREAD_CMD_INTERRUPT;
		if (!stricmp(aBuf, "NoTimers")) return THREAD_CMD_NOTIMERS;
		return THREAD_CMD_INVALID;
	}
	
	static ProcessCmds ConvertProcessCmd(char *aBuf)
	{
		if (!aBuf || !*aBuf) return PROCESS_CMD_INVALID;
		if (!stricmp(aBuf, "Exist")) return PROCESS_CMD_EXIST;
		if (!stricmp(aBuf, "Close")) return PROCESS_CMD_CLOSE;
		if (!stricmp(aBuf, "Priority")) return PROCESS_CMD_PRIORITY;
		if (!stricmp(aBuf, "Wait")) return PROCESS_CMD_WAIT;
		if (!stricmp(aBuf, "WaitClose")) return PROCESS_CMD_WAITCLOSE;
		return PROCESS_CMD_INVALID;
	}

	static ControlCmds ConvertControlCmd(char *aBuf)
	{
		if (!aBuf || !*aBuf) return CONTROL_CMD_INVALID;
		if (!stricmp(aBuf, "Check")) return CONTROL_CMD_CHECK;
		if (!stricmp(aBuf, "Uncheck")) return CONTROL_CMD_UNCHECK;
		if (!stricmp(aBuf, "Enable")) return CONTROL_CMD_ENABLE;
		if (!stricmp(aBuf, "Disable")) return CONTROL_CMD_DISABLE;
		if (!stricmp(aBuf, "Show")) return CONTROL_CMD_SHOW;
		if (!stricmp(aBuf, "Hide")) return CONTROL_CMD_HIDE;
		if (!stricmp(aBuf, "Style")) return CONTROL_CMD_STYLE;
		if (!stricmp(aBuf, "ExStyle")) return CONTROL_CMD_EXSTYLE;
		if (!stricmp(aBuf, "ShowDropDown")) return CONTROL_CMD_SHOWDROPDOWN;
		if (!stricmp(aBuf, "HideDropDown")) return CONTROL_CMD_HIDEDROPDOWN;
		if (!stricmp(aBuf, "TabLeft")) return CONTROL_CMD_TABLEFT;
		if (!stricmp(aBuf, "TabRight")) return CONTROL_CMD_TABRIGHT;
		if (!stricmp(aBuf, "Add")) return CONTROL_CMD_ADD;
		if (!stricmp(aBuf, "Delete")) return CONTROL_CMD_DELETE;
		if (!stricmp(aBuf, "Choose")) return CONTROL_CMD_CHOOSE;
		if (!stricmp(aBuf, "ChooseString")) return CONTROL_CMD_CHOOSESTRING;
		if (!stricmp(aBuf, "EditPaste")) return CONTROL_CMD_EDITPASTE;
		return CONTROL_CMD_INVALID;
	}

	static ControlGetCmds ConvertControlGetCmd(char *aBuf)
	{
		if (!aBuf || !*aBuf) return CONTROLGET_CMD_INVALID;
		if (!stricmp(aBuf, "Checked")) return CONTROLGET_CMD_CHECKED;
		if (!stricmp(aBuf, "Enabled")) return CONTROLGET_CMD_ENABLED;
		if (!stricmp(aBuf, "Visible")) return CONTROLGET_CMD_VISIBLE;
		if (!stricmp(aBuf, "Tab")) return CONTROLGET_CMD_TAB;
		if (!stricmp(aBuf, "FindString")) return CONTROLGET_CMD_FINDSTRING;
		if (!stricmp(aBuf, "Choice")) return CONTROLGET_CMD_CHOICE;
		if (!stricmp(aBuf, "List")) return CONTROLGET_CMD_LIST;
		if (!stricmp(aBuf, "LineCount")) return CONTROLGET_CMD_LINECOUNT;
		if (!stricmp(aBuf, "CurrentLine")) return CONTROLGET_CMD_CURRENTLINE;
		if (!stricmp(aBuf, "CurrentCol")) return CONTROLGET_CMD_CURRENTCOL;
		if (!stricmp(aBuf, "Line")) return CONTROLGET_CMD_LINE;
		if (!stricmp(aBuf, "Selected")) return CONTROLGET_CMD_SELECTED;
		if (!stricmp(aBuf, "Style")) return CONTROLGET_CMD_STYLE;
		if (!stricmp(aBuf, "ExStyle")) return CONTROLGET_CMD_EXSTYLE;
		if (!stricmp(aBuf, "Hwnd")) return CONTROLGET_CMD_HWND;
		return CONTROLGET_CMD_INVALID;
	}

	static DriveCmds ConvertDriveCmd(char *aBuf)
	{
		if (!aBuf || !*aBuf) return DRIVE_CMD_INVALID;
		if (!stricmp(aBuf, "Eject")) return DRIVE_CMD_EJECT;
		if (!stricmp(aBuf, "Lock")) return DRIVE_CMD_LOCK;
		if (!stricmp(aBuf, "Unlock")) return DRIVE_CMD_UNLOCK;
		if (!stricmp(aBuf, "Label")) return DRIVE_CMD_LABEL;
		return DRIVE_CMD_INVALID;
	}

	static DriveGetCmds ConvertDriveGetCmd(char *aBuf)
	{
		if (!aBuf || !*aBuf) return DRIVEGET_CMD_INVALID;
		if (!stricmp(aBuf, "List")) return DRIVEGET_CMD_LIST;
		if (!stricmp(aBuf, "FileSystem") || !stricmp(aBuf, "FS")) return DRIVEGET_CMD_FILESYSTEM;
		if (!stricmp(aBuf, "Label")) return DRIVEGET_CMD_LABEL;
		if (!strnicmp(aBuf, "SetLabel:", 9)) return DRIVEGET_CMD_SETLABEL;  // Uses strnicmp() vs. stricmp().
		if (!stricmp(aBuf, "Serial")) return DRIVEGET_CMD_SERIAL;
		if (!stricmp(aBuf, "Type")) return DRIVEGET_CMD_TYPE;
		if (!stricmp(aBuf, "Status")) return DRIVEGET_CMD_STATUS;
		if (!stricmp(aBuf, "StatusCD")) return DRIVEGET_CMD_STATUSCD;
		if (!stricmp(aBuf, "Capacity") || !stricmp(aBuf, "Cap")) return DRIVEGET_CMD_CAPACITY;
		return DRIVEGET_CMD_INVALID;
	}

	static WinSetAttributes ConvertWinSetAttribute(char *aBuf)
	{
		if (!aBuf || !*aBuf) return WINSET_INVALID;
		if (!stricmp(aBuf, "Trans") || !stricmp(aBuf, "Transparent")) return WINSET_TRANSPARENT;
		if (!stricmp(aBuf, "TransColor")) return WINSET_TRANSCOLOR;
		if (!stricmp(aBuf, "AlwaysOnTop") || !stricmp(aBuf, "Topmost")) return WINSET_ALWAYSONTOP;
		if (!stricmp(aBuf, "Bottom")) return WINSET_BOTTOM;
		if (!stricmp(aBuf, "Top")) return WINSET_TOP;
		if (!stricmp(aBuf, "Style")) return WINSET_STYLE;
		if (!stricmp(aBuf, "ExStyle")) return WINSET_EXSTYLE;
		if (!stricmp(aBuf, "Redraw")) return WINSET_REDRAW;
		if (!stricmp(aBuf, "Enable")) return WINSET_ENABLE;
		if (!stricmp(aBuf, "Disable")) return WINSET_DISABLE;
		if (!stricmp(aBuf, "Region")) return WINSET_REGION;
		return WINSET_INVALID;
	}


	static WinGetCmds ConvertWinGetCmd(char *aBuf)
	{
		if (!aBuf || !*aBuf) return WINGET_CMD_ID;  // If blank, return the default command.
		if (!stricmp(aBuf, "ID")) return WINGET_CMD_ID;
		if (!stricmp(aBuf, "IDLast")) return WINGET_CMD_IDLAST;
		if (!stricmp(aBuf, "PID")) return WINGET_CMD_PID;
		if (!stricmp(aBuf, "ProcessName")) return WINGET_CMD_PROCESSNAME;
		if (!stricmp(aBuf, "Count")) return WINGET_CMD_COUNT;
		if (!stricmp(aBuf, "List")) return WINGET_CMD_LIST;
		if (!stricmp(aBuf, "MinMax")) return WINGET_CMD_MINMAX;
		if (!stricmp(aBuf, "Style")) return WINGET_CMD_STYLE;
		if (!stricmp(aBuf, "ExStyle")) return WINGET_CMD_EXSTYLE;
		if (!stricmp(aBuf, "Transparent")) return WINGET_CMD_TRANSPARENT;
		if (!stricmp(aBuf, "TransColor")) return WINGET_CMD_TRANSCOLOR;
		if (!strnicmp(aBuf, "ControlList", 11))
		{
			aBuf += 11;
			if (!*aBuf)
				return WINGET_CMD_CONTROLLIST;
			if (!stricmp(aBuf, "Hwnd"))
				return WINGET_CMD_CONTROLLISTHWND;
			// Otherwise fall through to the below.
		}
		// Otherwise:
		return WINGET_CMD_INVALID;
	}

	static ToggleValueType ConvertOnOff(char *aBuf, ToggleValueType aDefault = TOGGLE_INVALID)
	// Returns aDefault if aBuf isn't either ON, OFF, or blank.
	{
		if (!aBuf || !*aBuf) return NEUTRAL;
		if (!stricmp(aBuf, "ON")) return TOGGLED_ON;
		if (!stricmp(aBuf, "OFF")) return TOGGLED_OFF;
		return aDefault;
	}

	static ToggleValueType ConvertOnOffAlways(char *aBuf, ToggleValueType aDefault = TOGGLE_INVALID)
	// Returns aDefault if aBuf isn't either ON, OFF, ALWAYSON, ALWAYSOFF, or blank.
	{
		if (!aBuf || !*aBuf) return NEUTRAL;
		if (!stricmp(aBuf, "On")) return TOGGLED_ON;
		if (!stricmp(aBuf, "Off")) return TOGGLED_OFF;
		if (!stricmp(aBuf, "AlwaysOn")) return ALWAYS_ON;
		if (!stricmp(aBuf, "AlwaysOff")) return ALWAYS_OFF;
		return aDefault;
	}

	static ToggleValueType ConvertOnOffToggle(char *aBuf, ToggleValueType aDefault = TOGGLE_INVALID)
	// Returns aDefault if aBuf isn't either ON, OFF, TOGGLE, or blank.
	{
		if (!aBuf || !*aBuf) return NEUTRAL;
		if (!stricmp(aBuf, "On")) return TOGGLED_ON;
		if (!stricmp(aBuf, "Off")) return TOGGLED_OFF;
		if (!stricmp(aBuf, "Toggle")) return TOGGLE;
		return aDefault;
	}

	static StringCaseSenseType ConvertStringCaseSense(char *aBuf)
	{
		if (!stricmp(aBuf, "On")) return SCS_SENSITIVE;
		if (!stricmp(aBuf, "Off")) return SCS_INSENSITIVE;
		if (!stricmp(aBuf, "Locale")) return SCS_INSENSITIVE_LOCALE;
		return SCS_INVALID;
	}

	static ToggleValueType ConvertOnOffTogglePermit(char *aBuf, ToggleValueType aDefault = TOGGLE_INVALID)
	// Returns aDefault if aBuf isn't either ON, OFF, TOGGLE, PERMIT, or blank.
	{
		if (!aBuf || !*aBuf) return NEUTRAL;
		if (!stricmp(aBuf, "On")) return TOGGLED_ON;
		if (!stricmp(aBuf, "Off")) return TOGGLED_OFF;
		if (!stricmp(aBuf, "Toggle")) return TOGGLE;
		if (!stricmp(aBuf, "Permit")) return TOGGLE_PERMIT;
		return aDefault;
	}

	static ToggleValueType ConvertBlockInput(char *aBuf)
	{
		if (!aBuf || !*aBuf) return NEUTRAL;  // For backward compatibility, blank is not considered INVALID.
		if (!stricmp(aBuf, "On")) return TOGGLED_ON;
		if (!stricmp(aBuf, "Off")) return TOGGLED_OFF;
		if (!stricmp(aBuf, "Send")) return TOGGLE_SEND;
		if (!stricmp(aBuf, "Mouse")) return TOGGLE_MOUSE;
		if (!stricmp(aBuf, "SendAndMouse")) return TOGGLE_SENDANDMOUSE;
		if (!stricmp(aBuf, "Default")) return TOGGLE_DEFAULT;
		if (!stricmp(aBuf, "MouseMove")) return TOGGLE_MOUSEMOVE;
		if (!stricmp(aBuf, "MouseMoveOff")) return TOGGLE_MOUSEMOVEOFF;
		return TOGGLE_INVALID;
	}

	static SendModes ConvertSendMode(char *aBuf, SendModes aValueToReturnIfInvalid)
	{
		if (!stricmp(aBuf, "Play")) return SM_PLAY;
		if (!stricmp(aBuf, "Event")) return SM_EVENT;
		if (!strnicmp(aBuf, "Input", 5)) // This IF must be listed last so that it can fall through to bottom line.
		{
			aBuf += 5;
			if (!*aBuf || !stricmp(aBuf, "ThenEvent")) // "ThenEvent" is supported for backward compatibiltity with 1.0.43.00.
				return SM_INPUT;
			if (!stricmp(aBuf, "ThenPlay"))
				return SM_INPUT_FALLBACK_TO_PLAY;
			//else fall through and return the indication of invalidity.
		}
		return aValueToReturnIfInvalid;
	}

	static FileLoopModeType ConvertLoopMode(char *aBuf)
	// Returns the file loop mode, or FILE_LOOP_INVALID if aBuf contains an invalid mode.
	{
		switch (ATOI(aBuf))
		{
		case 0: return FILE_LOOP_FILES_ONLY; // This is also the default mode if the param is blank.
		case 1: return FILE_LOOP_FILES_AND_FOLDERS;
		case 2: return FILE_LOOP_FOLDERS_ONLY;
		}
		// Otherwise:
		return FILE_LOOP_INVALID;
	}

	static int ConvertMsgBoxResult(char *aBuf)
	// Returns the matching ID, or zero if none.
	{
		if (!aBuf || !*aBuf) return 0;
		// Keeping the most oft-used ones up top helps perf. a little:
		if (!stricmp(aBuf, "YES")) return IDYES;
		if (!stricmp(aBuf, "NO")) return IDNO;
		if (!stricmp(aBuf, "OK")) return IDOK;
		if (!stricmp(aBuf, "CANCEL")) return IDCANCEL;
		if (!stricmp(aBuf, "ABORT")) return IDABORT;
		if (!stricmp(aBuf, "IGNORE")) return IDIGNORE;
		if (!stricmp(aBuf, "RETRY")) return IDRETRY;
		if (!stricmp(aBuf, "CONTINUE")) return IDCONTINUE; // v1.0.44.08: For use with 2000/XP's "Cancel/Try Again/Continue" MsgBox.
		if (!stricmp(aBuf, "TRYAGAIN")) return IDTRYAGAIN; //
		if (!stricmp(aBuf, "Timeout")) return AHK_TIMEOUT; // Our custom result value.
		return 0;
	}

	static int ConvertRunMode(char *aBuf)
	// Returns the matching WinShow mode, or SW_SHOWNORMAL if none.
	// These are also the modes that AutoIt3 uses.
	{
		// For v1.0.19, this was made more permissive (the use of strcasestr vs. stricmp) to support
		// the optional word ErrorLevel inside this parameter:
		if (!aBuf || !*aBuf) return SW_SHOWNORMAL;
		if (strcasestr(aBuf, "MIN")) return SW_MINIMIZE;
		if (strcasestr(aBuf, "MAX")) return SW_MAXIMIZE;
		if (strcasestr(aBuf, "HIDE")) return SW_HIDE;
		return SW_SHOWNORMAL;
	}

	static int ConvertMouseButton(char *aBuf, bool aAllowWheel = true, bool aUseLogicalButton = false)
	// Returns the matching VK, or zero if none.
	{
		if (!*aBuf || !stricmp(aBuf, "LEFT") || !stricmp(aBuf, "L"))
			return aUseLogicalButton ? VK_LBUTTON_LOGICAL : VK_LBUTTON; // Some callers rely on this default when !*aBuf.
		if (!stricmp(aBuf, "RIGHT") || !stricmp(aBuf, "R")) return aUseLogicalButton ? VK_RBUTTON_LOGICAL : VK_RBUTTON;
		if (!stricmp(aBuf, "MIDDLE") || !stricmp(aBuf, "M")) return VK_MBUTTON;
		if (!stricmp(aBuf, "X1")) return VK_XBUTTON1;
		if (!stricmp(aBuf, "X2")) return VK_XBUTTON2;
		if (aAllowWheel)
		{
			if (!stricmp(aBuf, "WheelUp") || !stricmp(aBuf, "WU")) return VK_WHEEL_UP;
			if (!stricmp(aBuf, "WheelDown") || !stricmp(aBuf, "WD")) return VK_WHEEL_DOWN;
			// Lexikos: Support horizontal scrolling in Windows Vista and later.
			if (!stricmp(aBuf, "WheelLeft") || !stricmp(aBuf, "WL")) return VK_WHEEL_LEFT;
			if (!stricmp(aBuf, "WheelRight") || !stricmp(aBuf, "WR")) return VK_WHEEL_RIGHT;
		}
		return 0;
	}

	static CoordModeAttribType ConvertCoordModeAttrib(char *aBuf)
	{
		if (!aBuf || !*aBuf) return 0;
		if (!stricmp(aBuf, "Pixel")) return COORD_MODE_PIXEL;
		if (!stricmp(aBuf, "Mouse")) return COORD_MODE_MOUSE;
		if (!stricmp(aBuf, "ToolTip")) return COORD_MODE_TOOLTIP;
		if (!stricmp(aBuf, "Caret")) return COORD_MODE_CARET;
		if (!stricmp(aBuf, "Menu")) return COORD_MODE_MENU;
		return 0;
	}

	static VariableTypeType ConvertVariableTypeName(char *aBuf)
	// Returns the matching type, or zero if none.
	{
		if (!aBuf || !*aBuf) return VAR_TYPE_INVALID;
		if (!stricmp(aBuf, "Integer")) return VAR_TYPE_INTEGER;
		if (!stricmp(aBuf, "Float")) return VAR_TYPE_FLOAT;
		if (!stricmp(aBuf, "Number")) return VAR_TYPE_NUMBER;
		if (!stricmp(aBuf, "Time")) return VAR_TYPE_TIME;
		if (!stricmp(aBuf, "Date")) return VAR_TYPE_TIME;  // "date" is just an alias for "time".
		if (!stricmp(aBuf, "Digit")) return VAR_TYPE_DIGIT;
		if (!stricmp(aBuf, "Xdigit")) return VAR_TYPE_XDIGIT;
		if (!stricmp(aBuf, "Alnum")) return VAR_TYPE_ALNUM;
		if (!stricmp(aBuf, "Alpha")) return VAR_TYPE_ALPHA;
		if (!stricmp(aBuf, "Upper")) return VAR_TYPE_UPPER;
		if (!stricmp(aBuf, "Lower")) return VAR_TYPE_LOWER;
		if (!stricmp(aBuf, "Space")) return VAR_TYPE_SPACE;
		return VAR_TYPE_INVALID;
	}

	static ResultType ValidateMouseCoords(char *aX, char *aY)
	{
		// OK: Both are absent, which is the signal to use the current position.
		// OK: Both are present (that they are numeric is validated elsewhere).
		// FAIL: One is absent but the other is present.
		return (!*aX && !*aY) || (*aX && *aY) ? OK : FAIL;
	}

	static char *LogToText(char *aBuf, int aBufSize);
	char *VicinityToText(char *aBuf, int aBufSize);
	char *ToText(char *aBuf, int aBufSize, bool aCRLF, DWORD aElapsed = 0, bool aLineWasResumed = false);

	static void ToggleSuspendState();
	static void PauseUnderlyingThread(bool aTrueForPauseFalseForUnpause);
	ResultType ChangePauseState(ToggleValueType aChangeTo, bool aAlwaysOperateOnUnderlyingThread);
	static ResultType ScriptBlockInput(bool aEnable);

	Line *PreparseError(char *aErrorText, char *aExtraInfo = "");
	// Call this LineError to avoid confusion with Script's error-displaying functions:
	ResultType LineError(char *aErrorText, ResultType aErrorType = FAIL, char *aExtraInfo = "");

	Line(FileIndexType aFileIndex, LineNumberType aFileLineNumber, ActionTypeType aActionType
		, ArgStruct aArg[], ArgCountType aArgc) // Constructor
		: mFileIndex(aFileIndex), mLineNumber(aFileLineNumber), mActionType(aActionType)
		, mAttribute(ATTR_NONE), mArgc(aArgc), mArg(aArg)
		, mPrevLine(NULL), mNextLine(NULL), mRelatedLine(NULL), mParentLine(NULL)
		{}
	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}  // Intentionally does nothing because we're using SimpleHeap for everything.
	void operator delete[](void *aPtr) {}

	// AutoIt3 functions:
	static bool Util_CopyDir(const char *szInputSource, const char *szInputDest, bool bOverwrite);
	static bool Util_MoveDir(const char *szInputSource, const char *szInputDest, int OverwriteMode);
	static bool Util_RemoveDir(const char *szInputSource, bool bRecurse);
	static int Util_CopyFile(const char *szInputSource, const char *szInputDest, bool bOverwrite, bool bMove);
	static void Util_ExpandFilenameWildcard(const char *szSource, const char *szDest, char *szExpandedDest);
	static void Util_ExpandFilenameWildcardPart(const char *szSource, const char *szDest, char *szExpandedDest);
	static bool Util_CreateDir(const char *szDirName);
	static bool Util_DoesFileExist(const char *szFilename);
	static bool Util_IsDir(const char *szPath);
	static void Util_GetFullPathName(const char *szIn, char *szOut);
	static bool Util_IsDifferentVolumes(const char *szPath1, const char *szPath2);
};



class Label
{
public:
	char *mName;
	Line *mJumpToLine;
	Label *mPrevLabel, *mNextLabel;  // Prev & Next items in linked list.

	bool IsExemptFromSuspend()
	{
		// Hotkey and Hotstring subroutines whose first line is the Suspend command are exempt from
		// being suspended themselves except when their first parameter is the literal
		// word "on":
		return mJumpToLine->mActionType == ACT_SUSPEND && (!mJumpToLine->mArgc || mJumpToLine->ArgHasDeref(1)
			|| stricmp(mJumpToLine->mArg[0].text, "On"));
	}

	ResultType Execute()
	// This function was added in v1.0.46.16 to support A_ThisLabel.
	{
		Label *prev_label =g->CurrentLabel; // This will be non-NULL when a subroutine is called from inside another subroutine.
		g->CurrentLabel = this;
		ResultType result = mJumpToLine->ExecUntil(UNTIL_RETURN); // The script loader has ensured that Label::mJumpToLine can't be NULL.
		g->CurrentLabel = prev_label;
		return result;
	}

	Label(char *aLabelName)
		: mName(aLabelName) // Caller gave us a pointer to dynamic memory for this (or an empty string in the case of mPlaceholderLabel).
		, mJumpToLine(NULL)
		, mPrevLabel(NULL), mNextLabel(NULL)
	{}
	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};



enum FuncParamDefaults {PARAM_DEFAULT_NONE, PARAM_DEFAULT_STR, PARAM_DEFAULT_INT, PARAM_DEFAULT_FLOAT};
struct FuncParam
{
	Var *var;
	WORD is_byref; // Boolean, but defined as WORD in case it helps data alignment and/or performance (BOOL vs. WORD didn't help benchmarks).
	WORD default_type;
	union {char *default_str; __int64 default_int64; double default_double;};
};

typedef void (* BuiltInFunctionType)(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);

class Func
{
public:
	char *mName;
	union {BuiltInFunctionType mBIF; Line *mJumpToLine;};
	FuncParam *mParam;  // Will hold an array of FuncParams.
	int mParamCount; // The number of items in the above array.  This is also the function's maximum number of params.
	int mMinParams;  // The number of mandatory parameters (populated for both UDFs and built-in's).
	Var **mVar, **mLazyVar; // Array of pointers-to-variable, allocated upon first use and later expanded as needed.
	int mVarCount, mVarCountMax, mLazyVarCount; // Count of items in the above array as well as the maximum capacity.
	int mInstances; // How many instances currently exist on the call stack (due to recursion or thread interruption).  Future use: Might be used to limit how deep recursion can go to help prevent stack overflow.
	Func *mNextFunc; // Next item in linked list.

	// Keep small members adjacent to each other to save space and improve perf. due to byte alignment:
	UCHAR mDefaultVarType;
	#define VAR_DECLARE_NONE   0
	#define VAR_DECLARE_GLOBAL 1
	#define VAR_DECLARE_LOCAL  2
	#define VAR_DECLARE_STATIC 3

	bool mIsBuiltIn; // Determines contents of union. Keep this member adjacent/contiguous with the above.
	// Note that it's possible for a built-in function such as WinExist() to become a normal/UDF via
	// override in the script.  So mIsBuiltIn should always be used to determine whether the function
	// is truly built-in, not its name.

	ResultType Call(char *&aReturnValue) // Making this a function vs. inline doesn't measurably impact performance.
	{
		aReturnValue = ""; // Init to default in case function doesn't return a value or it EXITs or fails.
		// Launch the function similar to Gosub (i.e. not as a new quasi-thread):
		// The performance gain of conditionally passing NULL in place of result (when this is the
		// outermost function call of a line consisting only of function calls, namely ACT_EXPRESSION)
		// would not be significant because the Return command's expression (arg1) must still be evaluated
		// in case it calls any functions that have side-effects, e.g. "return LogThisError()".
		Func *prev_func = g->CurrentFunc; // This will be non-NULL when a function is called from inside another function.
		g->CurrentFunc = this;
		// Although a GOTO that jumps to a position outside of the function's body could be supported,
		// it seems best not to for these reasons:
		// 1) The extreme rarity of a legitimate desire to intentionally do so.
		// 2) The fact that any return encountered after the Goto cannot provide a return value for
		//    the function because load-time validation checks for this (it's preferable not to
		//    give up this check, since it is an informative error message and might also help catch
		//    bugs in the script).  Gosub does not suffer from this because the return that brings it
		//    back into the function body belongs to the Gosub and not the function itself.
		// 3) More difficult to maintain because we have handle jump_to_line the same way ExecUntil() does,
		//    checking aResult the same way it does, then checking jump_to_line the same way it does, etc.
		// Fix for v1.0.31.05: g->mLoopFile and the other g_script members that follow it are
		// now passed to ExecUntil() for two reasons (update for v1.0.44.14: now they're implicitly "passed"
		// because they're done via parameter anymore):
		// 1) To fix the fact that any function call in one parameter of a command would reset
		// A_Index and related variables so that if those variables are referenced in another
		// parameter of the same command, they would be wrong.
		// 2) So that the caller's value of A_Index and such will always be valid even inside
		// of called functions (unless overridden/eclipsed by a loop in the body of the function),
		// which seems to add flexibility without giving up anything.  This fix is necessary at least
		// for a command that references A_Index in two of its args such as the following:
		// ToolTip, O, ((cos(A_Index) * 500) + 500), A_Index
		++mInstances;
		ResultType result = mJumpToLine->ExecUntil(UNTIL_BLOCK_END, &aReturnValue);
		--mInstances;
		// Restore the original value in case this function is called from inside another function.
		// Due to the synchronous nature of recursion and recursion-collapse, this should keep
		// g->CurrentFunc accurate, even amidst the asynchronous saving and restoring of "g" itself:
		g->CurrentFunc = prev_func;
		return result;
	}

	Func(char *aFuncName, bool aIsBuiltIn) // Constructor.
		: mName(aFuncName) // Caller gave us a pointer to dynamic memory for this.
		, mBIF(NULL)
		, mParam(NULL), mParamCount(0), mMinParams(0)
		, mVar(NULL), mVarCount(0), mVarCountMax(0), mLazyVar(NULL), mLazyVarCount(0)
		, mInstances(0), mNextFunc(NULL)
		, mDefaultVarType(VAR_DECLARE_NONE)
		, mIsBuiltIn(aIsBuiltIn)
	{}
	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};



class ScriptTimer
{
public:
	Label *mLabel;
	DWORD mPeriod; // v1.0.36.33: Changed from int to DWORD to double its capacity.
	DWORD mTimeLastRun;  // TickCount
	int mPriority;  // Thread priority relative to other threads, default 0.
	UCHAR mExistingThreads;  // Whether this timer is already running its subroutine.
	bool mEnabled;
	bool mRunOnlyOnce;
	ScriptTimer *mNextTimer;  // Next items in linked list
	void ScriptTimer::Disable();
	ScriptTimer(Label *aLabel)
		#define DEFAULT_TIMER_PERIOD 250
		: mLabel(aLabel), mPeriod(DEFAULT_TIMER_PERIOD), mPriority(0) // Default is always 0.
		, mExistingThreads(0), mTimeLastRun(0)
		, mEnabled(false), mRunOnlyOnce(false), mNextTimer(NULL)  // Note that mEnabled must default to false for the counts to be right.
	{}
	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};



struct MsgMonitorStruct
{
	UINT msg;
	Func *func;
	// Keep any members smaller than 4 bytes adjacent to save memory:
	short instance_count;  // Distinct from func.mInstances because the script might have called the function explicitly.
	short max_instances; // v1.0.47: Support more than one thread.
};



#define MAX_MENU_NAME_LENGTH MAX_PATH // For both menu and menu item names.
class UserMenuItem;  // Forward declaration since classes use each other (i.e. a menu *item* can have a submenu).
class UserMenu
{
public:
	char *mName;  // Dynamically allocated.
	UserMenuItem *mFirstMenuItem, *mLastMenuItem, *mDefault;
	// Keep any fields that aren't an even multiple of 4 adjacent to each other.  This conserves memory
	// due to byte-alignment:
	bool mIncludeStandardItems;
	int mClickCount; // How many clicks it takes to trigger the default menu item.  2 = double-click
	UINT mMenuItemCount;  // The count of user-defined menu items (doesn't include the standard items, if present).
	UserMenu *mNextMenu;  // Next item in linked list
	HMENU mMenu;
	MenuTypeType mMenuType; // MENU_TYPE_POPUP (via CreatePopupMenu) vs. MENU_TYPE_BAR (via CreateMenu).
	HBRUSH mBrush;   // Background color to apply to menu.
	COLORREF mColor; // The color that corresponds to the above.

	// Don't overload new and delete operators in this case since we want to use real dynamic memory
	// (since menus can be read in from a file, destroyed and recreated, over and over).

	UserMenu(char *aName) // Constructor
		: mName(aName), mFirstMenuItem(NULL), mLastMenuItem(NULL), mDefault(NULL)
		, mIncludeStandardItems(false), mClickCount(2), mMenuItemCount(0), mNextMenu(NULL), mMenu(NULL)
		, mMenuType(MENU_TYPE_POPUP) // The MENU_TYPE_NONE flag is not used in this context.  Default = POPUP.
		, mBrush(NULL), mColor(CLR_DEFAULT)
	{
	}

	ResultType AddItem(char *aName, UINT aMenuID, Label *aLabel, UserMenu *aSubmenu, char *aOptions);
	ResultType DeleteItem(UserMenuItem *aMenuItem, UserMenuItem *aMenuItemPrev);
	ResultType DeleteAllItems();
	ResultType ModifyItem(UserMenuItem *aMenuItem, Label *aLabel, UserMenu *aSubmenu, char *aOptions);
	void UpdateOptions(UserMenuItem *aMenuItem, char *aOptions);
	ResultType RenameItem(UserMenuItem *aMenuItem, char *aNewName);
	ResultType UpdateName(UserMenuItem *aMenuItem, char *aNewName);
	ResultType CheckItem(UserMenuItem *aMenuItem);
	ResultType UncheckItem(UserMenuItem *aMenuItem);
	ResultType ToggleCheckItem(UserMenuItem *aMenuItem);
	ResultType EnableItem(UserMenuItem *aMenuItem);
	ResultType DisableItem(UserMenuItem *aMenuItem);
	ResultType ToggleEnableItem(UserMenuItem *aMenuItem);
	ResultType SetDefault(UserMenuItem *aMenuItem = NULL);
	ResultType IncludeStandardItems();
	ResultType ExcludeStandardItems();
	ResultType Create(MenuTypeType aMenuType = MENU_TYPE_NONE); // NONE means UNSPECIFIED in this context.
	void SetColor(char *aColorName, bool aApplyToSubmenus);
	void ApplyColor(bool aApplyToSubmenus);
	ResultType AppendStandardItems();
	ResultType Destroy();
	ResultType Display(bool aForceToForeground = true, int aX = COORD_UNSPECIFIED, int aY = COORD_UNSPECIFIED);
	UINT GetSubmenuPos(HMENU ahMenu);
	UINT GetItemPos(char *aMenuItemName);
	bool ContainsMenu(UserMenu *aMenu);
};



class UserMenuItem
{
public:
	char *mName;  // Dynamically allocated.
	size_t mNameCapacity;
	UINT mMenuID;
	Label *mLabel;
	UserMenu *mSubmenu;
	UserMenu *mMenu;  // The menu to which this item belongs.  Needed to support script var A_ThisMenu.
	int mPriority;
	// Keep any fields that aren't an even multiple of 4 adjacent to each other.  This conserves memory
	// due to byte-alignment:
	bool mEnabled, mChecked;
	UserMenuItem *mNextMenuItem;  // Next item in linked list

	// Constructor:
	UserMenuItem(char *aName, size_t aNameCapacity, UINT aMenuID, Label *aLabel, UserMenu *aSubmenu, UserMenu *aMenu);

	// Don't overload new and delete operators in this case since we want to use real dynamic memory
	// (since menus can be read in from a file, destroyed and recreated, over and over).
};



struct FontType
{
	#define MAX_FONT_NAME_LENGTH 63  // Longest name I've seen is 29 chars, "Franklin Gothic Medium Italic". Anyway, there's protection against overflow.
	char name[MAX_FONT_NAME_LENGTH + 1];
	// Keep any fields that aren't an even multiple of 4 adjacent to each other.  This conserves memory
	// due to byte-alignment:
	bool italic;
	bool underline;
	bool strikeout;
	int point_size; // Decided to use int vs. float to simplify the code in many places. Fractional sizes seem rarely needed.
	int weight;
	HFONT hfont;
};

#define	LV_REMOTE_BUF_SIZE 1024  // 8192 (below) seems too large in hindsight, given that an LV can only display the first 260 chars in a field.
#define LV_TEXT_BUF_SIZE 8192  // Max amount of text in a ListView sub-item.  Somewhat arbitrary: not sure what the real limit is, if any.
enum LVColTypes {LV_COL_TEXT, LV_COL_INTEGER, LV_COL_FLOAT}; // LV_COL_TEXT must be zero so that it's the default with ZeroMemory.
struct lv_col_type
{
	UCHAR type;             // UCHAR vs. enum LVColTypes to save memory.
	bool sort_disabled;     // If true, clicking the column will have no automatic sorting effect.
	UCHAR case_sensitive;   // Ignored if type isn't LV_COL_TEXT.  SCS_INSENSITIVE is the default.
	bool unidirectional;    // Sorting cannot be reversed/toggled.
	bool prefer_descending; // Whether this column defaults to descending order (on first click or for unidirectional).
};

struct lv_attrib_type
{
	int sorted_by_col; // Index of column by which the control is currently sorted (-1 if none).
	bool is_now_sorted_ascending; // The direction in which the above column is currently sorted.
	bool no_auto_sort; // Automatic sorting disabled.
	#define LV_MAX_COLUMNS 200
	lv_col_type col[LV_MAX_COLUMNS];
	int col_count; // Number of columns currently in the above array.
	int row_count_hint;
};

typedef UCHAR TabControlIndexType;
typedef UCHAR TabIndexType;
// Keep the below in sync with the size of the types above:
#define MAX_TAB_CONTROLS 255  // i.e. the value 255 itself is reserved to mean "doesn't belong to a tab".
#define MAX_TABS_PER_CONTROL 256
struct GuiControlType
{
	HWND hwnd;
	// Keep any fields that are smaller than 4 bytes adjacent to each other.  This conserves memory
	// due to byte-alignment.  It has been verified to save 4 bytes per struct in this case:
	GuiControls type;
	#define GUI_CONTROL_ATTRIB_IMPLICIT_CANCEL     0x01
	#define GUI_CONTROL_ATTRIB_ALTSUBMIT           0x02
	#define GUI_CONTROL_ATTRIB_LABEL_IS_RUNNING    0x04
	#define GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN   0x08
	#define GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED 0x10
	#define GUI_CONTROL_ATTRIB_BACKGROUND_DEFAULT  0x20 // i.e. Don't conform to window/control background color; use default instead.
	#define GUI_CONTROL_ATTRIB_BACKGROUND_TRANS    0x40 // i.e. Leave this control's background transparent.
	#define GUI_CONTROL_ATTRIB_ALTBEHAVIOR         0x80 // For sliders: Reverse/Invert the value. Also for up-down controls (ALT means 32-bit vs. 16-bit). Also for ListView and Tab, and for Edit.
	UCHAR attrib; // A field of option flags/bits defined above.
	TabControlIndexType tab_control_index; // Which tab control this control belongs to, if any.
	TabIndexType tab_index; // For type==TAB, this stores the tab control's index.  For other types, it stores the page.
	Var *output_var;
	Label *jump_to_label;
	union
	{
		COLORREF union_color;  // Color of the control's text.
		HBITMAP union_hbitmap; // For PIC controls, stores the bitmap.
		// Note: Pic controls cannot obey the text color, but they can obey the window's background
		// color if the picture's background is transparent (at least in the case of icons on XP).
		lv_attrib_type *union_lv_attrib; // For ListView: Some attributes and an array of columns.
	};
	#define USES_FONT_AND_TEXT_COLOR(type) !(type == GUI_CONTROL_PIC || type == GUI_CONTROL_UPDOWN \
		|| type == GUI_CONTROL_SLIDER || type == GUI_CONTROL_PROGRESS)
};

struct GuiControlOptionsType
{
	DWORD style_add, style_remove, exstyle_add, exstyle_remove, listview_style;
	int listview_view; // Viewing mode, such as LVS_ICON, LVS_REPORT.  Int vs. DWORD to more easily use any negative value as "invalid".
	HIMAGELIST himagelist;
	Var *hwnd_output_var; // v1.0.46.01: Allows a script to retrieve the control's HWND upon creation of control.
	int x, y, width, height;  // Position info.
	float row_count;
	int choice;  // Which item of a DropDownList/ComboBox/ListBox to initially choose.
	int range_min, range_max;  // Allowable range, such as for a slider control.
	int tick_interval; // The interval at which to draw tickmarks for a slider control.
	int line_size, page_size; // Also for slider.
	int thickness;  // Thickness of slider's thumb.
	int tip_side; // Which side of the control to display the tip on (0 to use default side).
	GuiControlType *buddy1, *buddy2;
	COLORREF color_listview; // Used only for those controls that need control.union_color for something other than color.
	COLORREF color_bk; // Control's background color.
	int limit;   // The max number of characters to permit in an edit or combobox's edit (also used by ListView).
	int hscroll_pixels;  // The number of pixels for a listbox's horizontal scrollbar to be able to scroll.
	int checked; // When zeroed, struct contains default starting state of checkbox/radio, i.e. BST_UNCHECKED.
	int icon_number; // Which icon of a multi-icon file to use.  Zero means use-default, i.e. the first icon.
	#define GUI_MAX_TABSTOPS 50
	UINT tabstop[GUI_MAX_TABSTOPS]; // Array of tabstops for the interior of a multi-line edit control.
	UINT tabstop_count;  // The number of entries in the above array.
	SYSTEMTIME sys_time[2]; // Needs to support 2 elements for MONTHCAL's multi/range mode.
	SYSTEMTIME sys_time_range[2];
	DWORD gdtr, gdtr_range; // Used in connection with sys_time and sys_time_range.
	ResultType redraw;  // Whether the state of WM_REDRAW should be changed.
	char password_char; // When zeroed, indicates "use default password" for an edit control with the password style.
	bool range_changed;
	bool color_changed; // To discern when a control has been put back to the default color. [v1.0.26]
	bool start_new_section;
	bool use_theme; // v1.0.32: Provides the means for the window's current setting of mUseTheme to be overridden.
	bool listview_no_auto_sort; // v1.0.44: More maintainable and frees up GUI_CONTROL_ATTRIB_ALTBEHAVIOR for other uses.
};

LRESULT CALLBACK GuiWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK TabWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam);

class GuiType
{
public:
	#define GUI_STANDARD_WIDTH_MULTIPLIER 15 // This times font size = width, if all other means of determining it are exhausted.
	#define GUI_STANDARD_WIDTH (GUI_STANDARD_WIDTH_MULTIPLIER * sFont[mCurrentFontIndex].point_size)
	// Update for v1.0.21: Reduced it to 8 vs. 9 because 8 causes the height each edit (with the
	// default style) to exactly match that of a Combo or DropDownList.  This type of spacing seems
	// to be what other apps use too, and seems to make edits stand out a little nicer:
	#define GUI_CTL_VERTICAL_DEADSPACE 8
	#define PROGRESS_DEFAULT_THICKNESS (2 * sFont[mCurrentFontIndex].point_size)
	HWND mHwnd, mStatusBarHwnd;
	// Control IDs are higher than their index in the array by the below amount.  This offset is
	// necessary because windows that behave like dialogs automatically return IDOK and IDCANCEL in
	// response to certain types of standard actions:
	GuiIndexType mWindowIndex;
	GuiIndexType mControlCount;
	GuiIndexType mControlCapacity; // How many controls can fit into the current memory size of mControl.
	GuiControlType *mControl; // Will become an array of controls when the window is first created.
	GuiIndexType mDefaultButtonIndex; // Index vs. pointer is needed for some things.
	Label *mLabelForClose, *mLabelForEscape, *mLabelForSize, *mLabelForDropFiles, *mLabelForContextMenu;
	bool mLabelForCloseIsRunning, mLabelForEscapeIsRunning, mLabelForSizeIsRunning; // DropFiles doesn't need one of these.
	bool mLabelsHaveBeenSet;
	DWORD mStyle, mExStyle; // Style of window.
	bool mInRadioGroup; // Whether the control currently being created is inside a prior radio-group.
	bool mUseTheme;  // Whether XP theme and styles should be applied to the parent window and subsequently added controls.
	HWND mOwner;  // The window that owns this one, if any.  Note that Windows provides no way to change owners after window creation.
	char mDelimiter;  // The default field delimiter when adding items to ListBox, DropDownList, ListView, etc.
	int mCurrentFontIndex;
	GuiControlType *mCurrentListView, *mCurrentTreeView; // The ListView and TreeView upon which the LV/TV functions operate.
	TabControlIndexType mTabControlCount;
	TabControlIndexType mCurrentTabControlIndex; // Which tab control of the window.
	TabIndexType mCurrentTabIndex;// Which tab of a tab control is currently the default for newly added controls.
	COLORREF mCurrentColor;       // The default color of text in controls.
	COLORREF mBackgroundColorWin; // The window's background color itself.
	HBRUSH mBackgroundBrushWin;   // Brush corresponding to the above.
	COLORREF mBackgroundColorCtl; // Background color for controls.
	HBRUSH mBackgroundBrushCtl;   // Brush corresponding to the above.
	HDROP mHdrop;                 // Used for drag and drop operations.
	HICON mIconEligibleForDestruction; // The window's SysMenu icon, which can be destroyed when the window is destroyed if nothing else is using it.
	int mMarginX, mMarginY, mPrevX, mPrevY, mPrevWidth, mPrevHeight, mMaxExtentRight, mMaxExtentDown
		, mSectionX, mSectionY, mMaxExtentRightSection, mMaxExtentDownSection;
	LONG mMinWidth, mMinHeight, mMaxWidth, mMaxHeight;
	bool mGuiShowHasNeverBeenDone, mFirstActivation, mShowIsInProgress, mDestroyWindowHasBeenCalled;
	bool mControlWidthWasSetByContents; // Whether the most recently added control was auto-width'd to fit its contents.

	#define MAX_GUI_FONTS 200  // v1.0.44.14: Increased from 100 to 200 due to feedback that 100 wasn't enough.  But to alleviate memory usage, the array is now allocated upon first use.
	static FontType *sFont; // An array of structs, allocated upon first use.
	static int sFontCount;
	static int sGuiCount; // The number of non-NULL items in the g_gui array. Maintained only for performance reasons.
	static HWND sTreeWithEditInProgress; // Needed because TreeView's edit control for label-editing conflicts with IDOK (default button).

	// Don't overload new and delete operators in this case since we want to use real dynamic memory
	// (since GUIs can be destroyed and recreated, over and over).

	// Keep the default destructor to avoid entering the "Law of the Big Three": If your class requires a
	// copy constructor, copy assignment operator, or a destructor, then it very likely will require all three.

	GuiType(int aWindowIndex) // Constructor
		: mHwnd(NULL), mStatusBarHwnd(NULL), mWindowIndex(aWindowIndex), mControlCount(0), mControlCapacity(0)
		, mDefaultButtonIndex(-1), mLabelForClose(NULL), mLabelForEscape(NULL), mLabelForSize(NULL)
		, mLabelForDropFiles(NULL), mLabelForContextMenu(NULL)
		, mLabelForCloseIsRunning(false), mLabelForEscapeIsRunning(false), mLabelForSizeIsRunning(false)
		, mLabelsHaveBeenSet(false)
		// The styles DS_CENTER and DS_3DLOOK appear to be ineffectual in this case.
		// Also note that WS_CLIPSIBLINGS winds up on the window even if unspecified, which is a strong hint
		// that it should always be used for top level windows across all OSes.  Usenet posts confirm this.
		// Also, it seems safer to have WS_POPUP under a vague feeling that it seems to apply to dialog
		// style windows such as this one, and the fact that it also allows the window's caption to be
		// removed, which implies that POPUP windows are more flexible than OVERLAPPED windows.
		, mStyle(WS_POPUP|WS_CLIPSIBLINGS|WS_CAPTION|WS_SYSMENU|WS_MINIMIZEBOX) // WS_CLIPCHILDREN (doesn't seem helpful currently)
		, mExStyle(0) // This and the above should not be used once the window has been created since they might get out of date.
		, mInRadioGroup(false), mUseTheme(true), mOwner(NULL), mDelimiter('|')
		, mCurrentFontIndex(FindOrCreateFont()) // Must call this in constructor to ensure sFont array is never NULL while a GUI object exists.  Omit params to tell it to find or create DEFAULT_GUI_FONT.
		, mCurrentListView(NULL), mCurrentTreeView(NULL)
		, mTabControlCount(0), mCurrentTabControlIndex(MAX_TAB_CONTROLS), mCurrentTabIndex(0)
		, mCurrentColor(CLR_DEFAULT)
		, mBackgroundColorWin(CLR_DEFAULT), mBackgroundBrushWin(NULL)
		, mBackgroundColorCtl(CLR_DEFAULT), mBackgroundBrushCtl(NULL)
		, mHdrop(NULL), mIconEligibleForDestruction(NULL)
		, mMarginX(COORD_UNSPECIFIED), mMarginY(COORD_UNSPECIFIED) // These will be set when the first control is added.
		, mPrevX(0), mPrevY(0)
		, mPrevWidth(0), mPrevHeight(0) // Needs to be zero for first control to start off at right offset.
		, mMaxExtentRight(0), mMaxExtentDown(0)
		, mSectionX(COORD_UNSPECIFIED), mSectionY(COORD_UNSPECIFIED)
		, mMaxExtentRightSection(COORD_UNSPECIFIED), mMaxExtentDownSection(COORD_UNSPECIFIED)
		, mMinWidth(COORD_UNSPECIFIED), mMinHeight(COORD_UNSPECIFIED)
		, mMaxWidth(COORD_UNSPECIFIED), mMaxHeight(COORD_UNSPECIFIED)
		, mGuiShowHasNeverBeenDone(true), mFirstActivation(true), mShowIsInProgress(false)
		, mDestroyWindowHasBeenCalled(false), mControlWidthWasSetByContents(false)
	{
		// The array of controls is left unitialized to catch bugs.  Each control's attributes should be
		// fully populated when it is created.
		//ZeroMemory(mControl, sizeof(mControl));
	}

	static ResultType Destroy(GuiIndexType aWindowIndex);
	static void DestroyIconIfUnused(HICON ahIcon);
	ResultType Create();
	void SetLabels(char *aLabelPrefix);
	static void UpdateMenuBars(HMENU aMenu);
	ResultType AddControl(GuiControls aControlType, char *aOptions, char *aText);

	ResultType ParseOptions(char *aOptions, bool &aSetLastFoundWindow, ToggleValueType &aOwnDialogs);
	void GetNonClientArea(LONG &aWidth, LONG &aHeight);
	void GetTotalWidthAndHeight(LONG &aWidth, LONG &aHeight);

	ResultType ControlParseOptions(char *aOptions, GuiControlOptionsType &aOpt, GuiControlType &aControl
		, GuiIndexType aControlIndex = -1); // aControlIndex is not needed upon control creation.
	void ControlInitOptions(GuiControlOptionsType &aOpt, GuiControlType &aControl);
	void ControlAddContents(GuiControlType &aControl, char *aContent, int aChoice, GuiControlOptionsType *aOpt = NULL);
	ResultType Show(char *aOptions, char *aTitle);
	ResultType Clear();
	ResultType Cancel();
	ResultType Close(); // Due to SC_CLOSE, etc.
	ResultType Escape(); // Similar to close, except typically called when the user presses ESCAPE.
	ResultType Submit(bool aHideIt);
	ResultType ControlGetContents(Var &aOutputVar, GuiControlType &aControl, char *aMode = "");

	static VarSizeType ControlGetName(GuiIndexType aGuiWindowIndex, GuiIndexType aControlIndex, char *aBuf);
	static GuiType *FindGui(HWND aHwnd) // Find which GUI object owns the specified window.
	{
		#define EXTERN_GUI extern GuiType *g_gui[MAX_GUI_WINDOWS]
		EXTERN_GUI;
		if (!sGuiCount)
			return NULL;

		// The loop will usually find it on the first iteration since the #1 window is default
		// and thus most commonly used.
		int i, gui_count;
		for (i = 0, gui_count = 0; i < MAX_GUI_WINDOWS; ++i)
		{
			if (g_gui[i])
			{
				if (g_gui[i]->mHwnd == aHwnd)
					return g_gui[i];
				if (sGuiCount == ++gui_count) // No need to keep searching.
					break;
			}
		}
		return NULL;
	}


	GuiIndexType FindControl(char *aControlID);
	GuiControlType *FindControl(HWND aHwnd, bool aRetrieveIndexInstead = false)
	{
		GuiIndexType index = GUI_HWND_TO_INDEX(aHwnd); // Retrieves a small negative on failure, which will be out of bounds when converted to unsigned.
		if (index >= mControlCount) // Not found yet; try again with parent.
		{
			// Since ComboBoxes (and possibly other future control types) have children, try looking
			// up aHwnd's parent to see if its a known control of this dialog.  Some callers rely on us making
			// this extra effort:
			if (aHwnd = GetParent(aHwnd)) // Note that a ComboBox's drop-list (class ComboLBox) is apparently a direct child of the desktop, so this won't help us in that case.
				index = GUI_HWND_TO_INDEX(aHwnd); // Retrieves a small negative on failure, which will be out of bounds when converted to unsigned.
		}
		if (index < mControlCount) // A match was found.
			return aRetrieveIndexInstead ? (GuiControlType *)(size_t)index : mControl + index;
		else // No match, so indicate failure.
			return aRetrieveIndexInstead ? (GuiControlType *)NO_CONTROL_INDEX : NULL;
	}
	int FindGroup(GuiIndexType aControlIndex, GuiIndexType &aGroupStart, GuiIndexType &aGroupEnd);

	ResultType SetCurrentFont(char *aOptions, char *aFontName);
	static int FindOrCreateFont(char *aOptions = "", char *aFontName = "", FontType *aFoundationFont = NULL
		, COLORREF *aColor = NULL);
	static int FindFont(FontType &aFont);

	void Event(GuiIndexType aControlIndex, UINT aNotifyCode, USHORT aGuiEvent = GUI_EVENT_NONE, UINT aEventInfo = 0);

	static WORD TextToHotkey(char *aText);
	static char *HotkeyToText(WORD aHotkey, char *aBuf);
	void ControlCheckRadioButton(GuiControlType &aControl, GuiIndexType aControlIndex, WPARAM aCheckType);
	void ControlSetUpDownOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt);
	int ControlGetDefaultSliderThickness(DWORD aStyle, int aThumbThickness);
	void ControlSetSliderOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt);
	int ControlInvertSliderIfNeeded(GuiControlType &aControl, int aPosition);
	void ControlSetListViewOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt);
	void ControlSetTreeViewOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt);
	void ControlSetProgressOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt, DWORD aStyle);
	bool ControlOverrideBkColor(GuiControlType &aControl);

	void ControlUpdateCurrentTab(GuiControlType &aTabControl, bool aFocusFirstControl);
	GuiControlType *FindTabControl(TabControlIndexType aTabControlIndex);
	int FindTabIndexByName(GuiControlType &aTabControl, char *aName, bool aExactMatch = false);
	int GetControlCountOnTabPage(TabControlIndexType aTabControlIndex, TabIndexType aTabIndex);
	POINT GetPositionOfTabClientArea(GuiControlType &aTabControl);
	ResultType SelectAdjacentTab(GuiControlType &aTabControl, bool aMoveToRight, bool aFocusFirstControl
		, bool aWrapAround);
	void ControlGetPosOfFocusedItem(GuiControlType &aControl, POINT &aPoint);
	static void LV_Sort(GuiControlType &aControl, int aColumnIndex, bool aSortOnlyIfEnabled, char aForceDirection = '\0');
	static DWORD ControlGetListViewMode(HWND aWnd);
};



class Script
{
private:
	friend class Hotkey;
	Line *mFirstLine, *mLastLine;     // The first and last lines in the linked list.
	UINT mLineCount;                  // The number of lines.
	Label *mFirstLabel, *mLastLabel;  // The first and last labels in the linked list.
	Func *mFirstFunc, *mLastFunc;     // The first and last functions in the linked list.
	Var **mVar, **mLazyVar; // Array of pointers-to-variable, allocated upon first use and later expanded as needed.
	int mVarCount, mVarCountMax, mLazyVarCount; // Count of items in the above array as well as the maximum capacity.
	WinGroup *mFirstGroup, *mLastGroup;  // The first and last variables in the linked list.
	int mOpenBlockCount; // How many blocks are currently open.
	bool mNextLineIsFunctionBody; // Whether the very next line to be added will be the first one of the body.
	Var **mFuncExceptionVar;   // A list of variables declared explicitly local or global.
	int mFuncExceptionVarCount; // The number of items in the array.

	// These two track the file number and line number in that file of the line currently being loaded,
	// which simplifies calls to ScriptError() and LineError() (reduces the number of params that must be passed).
	// These are used ONLY while loading the script into memory.  After that (while the script is running),
	// only mCurrLine is kept up-to-date:
	int mCurrFileIndex;
	LineNumberType mCombinedLineNumber; // In the case of a continuation section/line(s), this is always the top line.

	bool mNoHotkeyLabels;
	bool mMenuUseErrorLevel;  // Whether runtime errors should be displayed by the Menu command, vs. ErrorLevel.

	#define UPDATE_TIP_FIELD strlcpy(mNIC.szTip, (mTrayIconTip && *mTrayIconTip) ? mTrayIconTip \
		: (mFileName ? mFileName : NAME_P), sizeof(mNIC.szTip));
	NOTIFYICONDATA mNIC; // For ease of adding and deleting our tray icon.

#ifdef AUTOHOTKEYSC
	ResultType CloseAndReturn(HS_EXEArc_Read *fp, UCHAR *aBuf, ResultType return_value);
	size_t GetLine(char *aBuf, int aMaxCharsToRead, int aInContinuationSection, UCHAR *&aMemFile);
#else
	ResultType CloseAndReturn(FILE *fp, UCHAR *aBuf, ResultType return_value);
	size_t GetLine(char *aBuf, int aMaxCharsToRead, int aInContinuationSection, FILE *fp);
#endif
	ResultType IsDirective(char *aBuf);

	ResultType ParseAndAddLine(char *aLineText, ActionTypeType aActionType = ACT_INVALID
		, ActionTypeType aOldActionType = OLD_INVALID, char *aActionName = NULL
		, char *aEndMarker = NULL, char *aLiteralMap = NULL, size_t aLiteralMapLength = 0);
	ResultType ParseDerefs(char *aArgText, char *aArgMap, DerefType *aDeref, int &aDerefCount);
	char *ParseActionType(char *aBufTarget, char *aBufSource, bool aDisplayErrors);
	static ActionTypeType ConvertActionType(char *aActionTypeString);
	static ActionTypeType ConvertOldActionType(char *aActionTypeString);
	ResultType AddLabel(char *aLabelName, bool aAllowDupe);
	ResultType AddLine(ActionTypeType aActionType, char *aArg[] = NULL, ArgCountType aArgc = 0, char *aArgMap[] = NULL);

	// These aren't in the Line class because I think they're easier to implement
	// if aStartingLine is allowed to be NULL (for recursive calls).  If they
	// were member functions of class Line, a check for NULL would have to
	// be done before dereferencing any line's mNextLine, for example:
	Line *PreparseBlocks(Line *aStartingLine, bool aFindBlockEnd = false, Line *aParentLine = NULL);
	Line *PreparseIfElse(Line *aStartingLine, ExecUntilMode aMode = NORMAL_MODE, AttributeType aLoopTypeFile = ATTR_NONE
		, AttributeType aLoopTypeReg = ATTR_NONE, AttributeType aLoopTypeRead = ATTR_NONE
		, AttributeType aLoopTypeParse = ATTR_NONE);

public:
	Line *mCurrLine;     // Seems better to make this public than make Line our friend.
	Label *mPlaceholderLabel; // Used in place of a NULL label to simplify code.
	char mThisMenuItemName[MAX_MENU_NAME_LENGTH + 1];
	char mThisMenuName[MAX_MENU_NAME_LENGTH + 1];
	char *mThisHotkeyName, *mPriorHotkeyName;
	HWND mNextClipboardViewer;
	bool mOnClipboardChangeIsRunning;
	Label *mOnClipboardChangeLabel, *mOnExitLabel;  // The label to run when the script terminates (NULL if none).
	ExitReasons mExitReason;

	ScriptTimer *mFirstTimer, *mLastTimer;  // The first and last script timers in the linked list.
	UINT mTimerCount, mTimerEnabledCount;

	UserMenu *mFirstMenu, *mLastMenu;
	UINT mMenuCount;

	DWORD mThisHotkeyStartTime, mPriorHotkeyStartTime;  // Tickcount timestamp of when its subroutine began.
	char mEndChar;  // The ending character pressed to trigger the most recent non-auto-replace hotstring.
	modLR_type mThisHotkeyModifiersLR;
	char *mFileSpec; // Will hold the full filespec, for convenience.
	char *mFileDir;  // Will hold the directory containing the script file.
	char *mFileName; // Will hold the script's naked file name.
	char *mOurEXE; // Will hold this app's module name (e.g. C:\Program Files\AutoHotkey\AutoHotkey.exe).
	char *mOurEXEDir;  // Same as above but just the containing diretory (for convenience).
	char *mMainWindowTitle; // Will hold our main window's title, for consistency & convenience.
	bool mIsReadyToExecute;
	bool mAutoExecSectionIsRunning;
	bool mIsRestart; // The app is restarting rather than starting from scratch.
	bool mIsAutoIt2; // Whether this script is considered to be an AutoIt2 script.
	bool mErrorStdOut; // true if load-time syntax errors should be sent to stdout vs. a MsgBox.
#ifdef AUTOHOTKEYSC
	bool mCompiledHasCustomIcon; // Whether the compiled script uses a custom icon.
#else
	FILE *mIncludeLibraryFunctionsThenExit;
#endif
	__int64 mLinesExecutedThisCycle; // Use 64-bit to match the type of g->LinesPerCycle
	int mUninterruptedLineCountMax; // 32-bit for performance (since huge values seem unnecessary here).
	int mUninterruptibleTime;
	DWORD mLastScriptRest, mLastPeekTime;

	#define RUNAS_SIZE_IN_WCHARS 257  // Includes the terminator.
	#define RUNAS_SIZE_IN_BYTES (RUNAS_SIZE_IN_WCHARS * sizeof(WCHAR))
	WCHAR *mRunAsUser, *mRunAsPass, *mRunAsDomain; // Memory is allocated at runtime, upon first use.

	HICON mCustomIcon;  // NULL unless the script has loaded a custom icon during its runtime.
	char *mCustomIconFile; // Filename of icon.  Allocated on first use.
	bool mIconFrozen; // If true, the icon does not change state when the state of pause or suspend changes.
	char *mTrayIconTip;  // Custom tip text for tray icon.  Allocated on first use.
	UINT mCustomIconNumber; // The number of the icon inside the above file.

	UserMenu *mTrayMenu; // Our tray menu, which should be destroyed upon exiting the program.
    
	ResultType Init(global_struct &g, char *aScriptFilename, bool aIsRestart);
	ResultType CreateWindows();
	void EnableOrDisableViewMenuItems(HMENU aMenu, UINT aFlags);
	void CreateTrayIcon();
	void UpdateTrayIcon(bool aForceUpdate = false);
	ResultType AutoExecSection();
	ResultType Edit();
	ResultType Reload(bool aDisplayErrors);
	ResultType ExitApp(ExitReasons aExitReason, char *aBuf = NULL, int ExitCode = 0);
	void TerminateApp(int aExitCode);
#ifdef AUTOHOTKEYSC
	LineNumberType LoadFromFile();
#else
	LineNumberType LoadFromFile(bool aScriptWasNotspecified);
#endif
	ResultType LoadIncludedFile(char *aFileSpec, bool aAllowDuplicateInclude, bool aIgnoreLoadFailure);
	ResultType UpdateOrCreateTimer(Label *aLabel, char *aPeriod, char *aPriority, bool aEnable
		, bool aUpdatePriorityOnly);

	ResultType DefineFunc(char *aBuf, Var *aFuncExceptionVar[]);
#ifndef AUTOHOTKEYSC
	Func *FindFuncInLibrary(char *aFuncName, size_t aFuncNameLength, bool &aErrorWasShown);
#endif
	Func *FindFunc(char *aFuncName, size_t aFuncNameLength = 0);
	Func *AddFunc(char *aFuncName, size_t aFuncNameLength, bool aIsBuiltIn);

	#define ALWAYS_USE_DEFAULT  0
	#define ALWAYS_USE_GLOBAL   1
	#define ALWAYS_USE_LOCAL    2
	#define ALWAYS_PREFER_LOCAL 3
	Var *FindOrAddVar(char *aVarName, size_t aVarNameLength = 0, int aAlwaysUse = ALWAYS_USE_DEFAULT
		, bool *apIsException = NULL);
	Var *FindVar(char *aVarName, size_t aVarNameLength = 0, int *apInsertPos = NULL
		, int aAlwaysUse = ALWAYS_USE_DEFAULT, bool *apIsException = NULL
		, bool *apIsLocal = NULL);
	Var *AddVar(char *aVarName, size_t aVarNameLength, int aInsertPos, int aIsLocal);
	static void *GetVarType(char *aVarName);

	WinGroup *FindGroup(char *aGroupName, bool aCreateIfNotFound = false);
	ResultType AddGroup(char *aGroupName);
	Label *FindLabel(char *aLabelName);

	ResultType DoRunAs(char *aCommandLine, char *aWorkingDir, bool aDisplayErrors, bool aUpdateLastError, WORD aShowWindow
		, Var *aOutputVar, PROCESS_INFORMATION &aPI, bool &aSuccess, HANDLE &aNewProcess, char *aSystemErrorText);
	ResultType ActionExec(char *aAction, char *aParams = NULL, char *aWorkingDir = NULL
		, bool aDisplayErrors = true, char *aRunShowMode = NULL, HANDLE *aProcess = NULL
		, bool aUpdateLastError = false, bool aUseRunAs = false, Var *aOutputVar = NULL);

	char *ListVars(char *aBuf, int aBufSize);
	char *ListKeyHistory(char *aBuf, int aBufSize);

	ResultType PerformMenu(char *aMenu, char *aCommand, char *aParam3, char *aParam4, char *aOptions);
	UserMenu *FindMenu(char *aMenuName);
	UserMenu *AddMenu(char *aMenuName);
	ResultType ScriptDeleteMenu(UserMenu *aMenu);
	UserMenuItem *FindMenuItemByID(UINT aID)
	{
		UserMenuItem *mi;
		for (UserMenu *m = mFirstMenu; m; m = m->mNextMenu)
			for (mi = m->mFirstMenuItem; mi; mi = mi->mNextMenuItem)
				if (mi->mMenuID == aID)
					return mi;
		return NULL;
	}

	ResultType PerformGui(char *aCommand, char *aControlType, char *aOptions, char *aParam4);

	// Call this SciptError to avoid confusion with Line's error-displaying functions:
	ResultType ScriptError(char *aErrorText, char *aExtraInfo = ""); // , ResultType aErrorType = FAIL);

	#define SOUNDPLAY_ALIAS "AHK_PlayMe"  // Used by destructor and SoundPlay().

	Script();
	~Script();
	// Note that the anchors to any linked lists will be lost when this
	// object goes away, so for now, be sure the destructor is only called
	// when the program is about to be exited, which will thereby reclaim
	// the memory used by the abandoned linked lists (otherwise, a memory
	// leak will result).
};



////////////////////////
// BUILT-IN VARIABLES //
////////////////////////
VarSizeType BIV_True_False(char *aBuf, char *aVarName);
VarSizeType BIV_MMM_DDD(char *aBuf, char *aVarName);
VarSizeType BIV_DateTime(char *aBuf, char *aVarName);
VarSizeType BIV_BatchLines(char *aBuf, char *aVarName);
VarSizeType BIV_TitleMatchMode(char *aBuf, char *aVarName);
VarSizeType BIV_TitleMatchModeSpeed(char *aBuf, char *aVarName);
VarSizeType BIV_DetectHiddenWindows(char *aBuf, char *aVarName);
VarSizeType BIV_DetectHiddenText(char *aBuf, char *aVarName);
VarSizeType BIV_AutoTrim(char *aBuf, char *aVarName);
VarSizeType BIV_StringCaseSense(char *aBuf, char *aVarName);
VarSizeType BIV_FormatInteger(char *aBuf, char *aVarName);
VarSizeType BIV_FormatFloat(char *aBuf, char *aVarName);
VarSizeType BIV_KeyDelay(char *aBuf, char *aVarName);
VarSizeType BIV_WinDelay(char *aBuf, char *aVarName);
VarSizeType BIV_ControlDelay(char *aBuf, char *aVarName);
VarSizeType BIV_MouseDelay(char *aBuf, char *aVarName);
VarSizeType BIV_DefaultMouseSpeed(char *aBuf, char *aVarName);
VarSizeType BIV_IsPaused(char *aBuf, char *aVarName);
VarSizeType BIV_IsCritical(char *aBuf, char *aVarName);
VarSizeType BIV_IsSuspended(char *aBuf, char *aVarName);
#ifdef AUTOHOTKEYSC  // A_IsCompiled is left blank/undefined in uncompiled scripts.
VarSizeType BIV_IsCompiled(char *aBuf, char *aVarName);
#endif
VarSizeType BIV_LastError(char *aBuf, char *aVarName);
VarSizeType BIV_IconHidden(char *aBuf, char *aVarName);
VarSizeType BIV_IconTip(char *aBuf, char *aVarName);
VarSizeType BIV_IconFile(char *aBuf, char *aVarName);
VarSizeType BIV_IconNumber(char *aBuf, char *aVarName);
VarSizeType BIV_ExitReason(char *aBuf, char *aVarName);
VarSizeType BIV_Space_Tab(char *aBuf, char *aVarName);
VarSizeType BIV_AhkVersion(char *aBuf, char *aVarName);
VarSizeType BIV_AhkPath(char *aBuf, char *aVarName);
VarSizeType BIV_TickCount(char *aBuf, char *aVarName);
VarSizeType BIV_Now(char *aBuf, char *aVarName);
VarSizeType BIV_OSType(char *aBuf, char *aVarName);
VarSizeType BIV_OSVersion(char *aBuf, char *aVarName);
VarSizeType BIV_Language(char *aBuf, char *aVarName);
VarSizeType BIV_UserName_ComputerName(char *aBuf, char *aVarName);
VarSizeType BIV_WorkingDir(char *aBuf, char *aVarName);
VarSizeType BIV_WinDir(char *aBuf, char *aVarName);
VarSizeType BIV_Temp(char *aBuf, char *aVarName);
VarSizeType BIV_ComSpec(char *aBuf, char *aVarName);
VarSizeType BIV_ProgramFiles(char *aBuf, char *aVarName);
VarSizeType BIV_AppData(char *aBuf, char *aVarName);
VarSizeType BIV_Desktop(char *aBuf, char *aVarName);
VarSizeType BIV_StartMenu(char *aBuf, char *aVarName);
VarSizeType BIV_Programs(char *aBuf, char *aVarName);
VarSizeType BIV_Startup(char *aBuf, char *aVarName);
VarSizeType BIV_MyDocuments(char *aBuf, char *aVarName);
VarSizeType BIV_Caret(char *aBuf, char *aVarName);
VarSizeType BIV_Cursor(char *aBuf, char *aVarName);
VarSizeType BIV_ScreenWidth_Height(char *aBuf, char *aVarName);
VarSizeType BIV_ScriptName(char *aBuf, char *aVarName);
VarSizeType BIV_ScriptDir(char *aBuf, char *aVarName);
VarSizeType BIV_ScriptFullPath(char *aBuf, char *aVarName);
VarSizeType BIV_LineNumber(char *aBuf, char *aVarName);
VarSizeType BIV_LineFile(char *aBuf, char *aVarName);
VarSizeType BIV_LoopFileName(char *aBuf, char *aVarName);
VarSizeType BIV_LoopFileShortName(char *aBuf, char *aVarName);
VarSizeType BIV_LoopFileExt(char *aBuf, char *aVarName);
VarSizeType BIV_LoopFileDir(char *aBuf, char *aVarName);
VarSizeType BIV_LoopFileFullPath(char *aBuf, char *aVarName);
VarSizeType BIV_LoopFileLongPath(char *aBuf, char *aVarName);
VarSizeType BIV_LoopFileShortPath(char *aBuf, char *aVarName);
VarSizeType BIV_LoopFileTime(char *aBuf, char *aVarName);
VarSizeType BIV_LoopFileAttrib(char *aBuf, char *aVarName);
VarSizeType BIV_LoopFileSize(char *aBuf, char *aVarName);
VarSizeType BIV_LoopRegType(char *aBuf, char *aVarName);
VarSizeType BIV_LoopRegKey(char *aBuf, char *aVarName);
VarSizeType BIV_LoopRegSubKey(char *aBuf, char *aVarName);
VarSizeType BIV_LoopRegName(char *aBuf, char *aVarName);
VarSizeType BIV_LoopRegTimeModified(char *aBuf, char *aVarName);
VarSizeType BIV_LoopReadLine(char *aBuf, char *aVarName);
VarSizeType BIV_LoopField(char *aBuf, char *aVarName);
VarSizeType BIV_LoopIndex(char *aBuf, char *aVarName);
VarSizeType BIV_ThisFunc(char *aBuf, char *aVarName);
VarSizeType BIV_ThisLabel(char *aBuf, char *aVarName);
VarSizeType BIV_ThisMenuItem(char *aBuf, char *aVarName);
VarSizeType BIV_ThisMenuItemPos(char *aBuf, char *aVarName);
VarSizeType BIV_ThisMenu(char *aBuf, char *aVarName);
VarSizeType BIV_ThisHotkey(char *aBuf, char *aVarName);
VarSizeType BIV_PriorHotkey(char *aBuf, char *aVarName);
VarSizeType BIV_TimeSinceThisHotkey(char *aBuf, char *aVarName);
VarSizeType BIV_TimeSincePriorHotkey(char *aBuf, char *aVarName);
VarSizeType BIV_EndChar(char *aBuf, char *aVarName);
VarSizeType BIV_Gui(char *aBuf, char *aVarName);
VarSizeType BIV_GuiControl(char *aBuf, char *aVarName);
VarSizeType BIV_GuiEvent(char *aBuf, char *aVarName);
VarSizeType BIV_EventInfo(char *aBuf, char *aVarName);
VarSizeType BIV_TimeIdle(char *aBuf, char *aVarName);
VarSizeType BIV_TimeIdlePhysical(char *aBuf, char *aVarName);
VarSizeType BIV_IPAddress(char *aBuf, char *aVarName);
VarSizeType BIV_IsAdmin(char *aBuf, char *aVarName);



////////////////////////
// BUILT-IN FUNCTIONS //
////////////////////////
// Caller has ensured that SYM_VAR's Type() is VAR_NORMAL and that it's either not an environment
// variable or the caller wants environment varibles treated as having zero length.
#define EXPR_TOKEN_LENGTH(token_raw, token_as_string) \
(token_raw->symbol == SYM_VAR && !token_raw->var->IsBinaryClip()) \
	? token_raw->var->Length()\
	: strlen(token_as_string)

void BIF_DllCall(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_StrLen(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_SubStr(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_InStr(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_RegEx(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_Asc(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_Chr(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_NumGet(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_NumPut(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_IsLabel(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_IsFunc(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_GetKeyState(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_VarSetCapacity(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_FileExist(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_WinExistActive(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_Round(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_FloorCeil(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_Mod(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_Abs(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_Sin(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_Cos(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_Tan(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_ASinACos(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_ATan(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_Exp(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_SqrtLogLn(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);

void BIF_OnMessage(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_RegisterCallback(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);

void BIF_StatusBar(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);

void BIF_LV_GetNextOrCount(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_LV_GetText(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_LV_AddInsertModify(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_LV_Delete(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_LV_InsertModifyDeleteCol(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_LV_SetImageList(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);

void BIF_TV_AddModifyDelete(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_TV_GetRelatedItem(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_TV_Get(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);

void BIF_IL_Create(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_IL_Destroy(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
void BIF_IL_Add(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);

BOOL LegacyResultToBOOL(char *aResult);
BOOL LegacyVarToBOOL(Var &aVar);
BOOL TokenToBOOL(ExprTokenType &aToken, SymbolType aTokenIsNumber);
SymbolType TokenIsPureNumeric(ExprTokenType &aToken);
__int64 TokenToInt64(ExprTokenType &aToken, BOOL aIsPureInteger = FALSE);
double TokenToDouble(ExprTokenType &aToken, BOOL aCheckForHex = TRUE, BOOL aIsPureFloat = FALSE);
char *TokenToString(ExprTokenType &aToken, char *aBuf = NULL);
ResultType TokenToDoubleOrInt64(ExprTokenType &aToken);

char *RegExMatch(char *aHaystack, char *aNeedleRegEx);
void SetWorkingDir(char *aNewDir);
int ConvertJoy(char *aBuf, int *aJoystickID = NULL, bool aAllowOnlyButtons = false);
bool ScriptGetKeyState(vk_type aVK, KeyStateTypes aKeyStateType);
double ScriptGetJoyState(JoyControls aJoy, int aJoystickID, ExprTokenType &aToken, bool aUseBoolForUpDown);

#endif
