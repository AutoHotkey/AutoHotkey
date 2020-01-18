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
#include "resources/resource.h"  // For tray icon.
#include "script_object.h"
#include "Debugger.h"

#include "os_version.h" // For the global OS_Version object
EXTERN_OSVER; // For the access to the g_os version object without having to include globaldata.h
EXTERN_G;

class ScriptModule;						// forward declaration
class ModuleList;
extern ScriptModule* g_CurrentModule;

#define MAX_THREADS_LIMIT UCHAR_MAX // Uses UCHAR_MAX (255) because some variables that store a thread count are UCHARs.
#define MAX_THREADS_DEFAULT 10 // Must not be higher than above.
#define TOTAL_ADDITIONAL_THREADS 2 // See below.
// Must allow two additional threads: One for the AutoExec/idle thread and one so that ExitApp()
// can run even when #MaxThreads has been reached.
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
#define ATTR_NONE (void *)0  // Some places might rely on this being zero.
#define ATTR_TRUE (void *)1
typedef void *AttributeType;

typedef int FileLoopModeType;
#define FILE_LOOP_INVALID		0
#define FILE_LOOP_FILES_ONLY	1
#define FILE_LOOP_FOLDERS_ONLY	2
#define FILE_LOOP_RECURSE		4
#define FILE_LOOP_FILES_AND_FOLDERS (FILE_LOOP_FILES_ONLY | FILE_LOOP_FOLDERS_ONLY)

enum VariableTypeType {VAR_TYPE_INVALID, VAR_TYPE_NUMBER, VAR_TYPE_INTEGER, VAR_TYPE_FLOAT
	, VAR_TYPE_TIME	, VAR_TYPE_DIGIT, VAR_TYPE_XDIGIT, VAR_TYPE_ALNUM, VAR_TYPE_ALPHA
	, VAR_TYPE_UPPER, VAR_TYPE_LOWER, VAR_TYPE_SPACE
	};

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
	, ID_TRAY_SEP1, ID_TRAY_SEP2, ID_TRAY_LAST = ID_TRAY_SEP2 // But this value should never hit the below. There is debug code to enforce.
	, ID_MAIN_FIRST = 65400, ID_MAIN_LAST = 65534}; // These should match the range used by resource.h

#define GUI_INDEX_TO_ID(index) (index + CONTROL_ID_FIRST)
#define GUI_ID_TO_INDEX(id) (id - CONTROL_ID_FIRST) // Returns a small negative if "id" is invalid, such as 0.
#define GUI_HWND_TO_INDEX(hwnd) GUI_ID_TO_INDEX(GetDlgCtrlID(hwnd)) // Returns a small negative on failure (e.g. HWND not found).
// Notes about above:
// 1) Callers should call GuiType::FindControl() instead of GUI_HWND_TO_INDEX() if the hwnd might be a combobox's
//    edit control.
// 2) Testing shows that GetDlgCtrlID() is much faster than looping through a GUI window's control array to find
//    a matching HWND.


#define ERR_ABORT_NO_SPACES _T("The current thread will exit.")
#define ERR_ABORT _T("  ") ERR_ABORT_NO_SPACES
#define WILL_EXIT _T("The program will exit.")
#define UNSTABLE_WILL_EXIT _T("The program is now unstable and will exit.")
#define OLD_STILL_IN_EFFECT _T("The script was not reloaded; the old version will remain in effect.")
#define ERR_CONTINUE_THREAD_Q _T("Try to continue anyway?")
#define ERR_SCRIPT_NOT_FOUND _T("Script file not found.")
#define ERR_ABORT_DELETE _T("__Delete will now return.")
#define ERR_UNRECOGNIZED_ACTION _T("This line does not contain a recognized action.")
#define ERR_NONEXISTENT_HOTKEY _T("Nonexistent hotkey.")
#define ERR_NONEXISTENT_VARIANT _T("Nonexistent hotkey variant (IfWin).")
#define ERR_INVALID_KEYNAME _T("Invalid key name.")
#define ERR_UNSUPPORTED_PREFIX _T("Unsupported prefix key.")
#define ERR_ALTTAB_MODLR _T("This AltTab hotkey must specify which key (L or R).")
#define ERR_ALTTAB_ONEMOD _T("This AltTab hotkey must have exactly one modifier/prefix.")
#define ERR_INVALID_SINGLELINE_HOT _T("Not valid for a single-line hotkey/hotstring.")
#define ERR_NONEXISTENT_FUNCTION _T("Call to nonexistent function.")
#define ERR_UNRECOGNIZED_DIRECTIVE _T("Unknown directive.")
#define ERR_EXE_CORRUPTED _T("EXE corrupted")
#define ERR_INVALID_INDEX _T("Invalid index.")
#define ERR_INVALID_VALUE _T("Invalid value.")
#define ERR_INVALID_FUNCTOR _T("Invalid callback function.")
#define ERR_PARAM_INVALID _T("Invalid parameter(s).")
#define ERR_PARAM_COUNT_INVALID _T("Invalid number of parameters.")
#define ERR_PARAM1_INVALID _T("Parameter #1 invalid.")
#define ERR_PARAM2_INVALID _T("Parameter #2 invalid.")
#define ERR_PARAM3_INVALID _T("Parameter #3 invalid.")
#define ERR_PARAM4_INVALID _T("Parameter #4 invalid.")
#define ERR_PARAM5_INVALID _T("Parameter #5 invalid.")
#define ERR_PARAM6_INVALID _T("Parameter #6 invalid.")
#define ERR_PARAM7_INVALID _T("Parameter #7 invalid.")
#define ERR_PARAM8_INVALID _T("Parameter #8 invalid.")
#define ERR_PARAM1_REQUIRED _T("Parameter #1 required")
#define ERR_PARAM2_REQUIRED _T("Parameter #2 required")
#define ERR_PARAM3_REQUIRED _T("Parameter #3 required")
#define ERR_PARAM2_MUST_BE_BLANK _T("Parameter #2 must be blank in this case.")
#define ERR_PARAM3_MUST_BE_BLANK _T("Parameter #3 must be blank in this case.")
#define ERR_PARAM4_MUST_BE_BLANK _T("Parameter #4 must be blank in this case.")
#define ERR_PARAM1_MUST_NOT_BE_BLANK _T("Parameter #1 must not be blank in this case.")
#define ERR_PARAM2_MUST_NOT_BE_BLANK _T("Parameter #2 must not be blank in this case.")
#define ERR_PARAM3_MUST_NOT_BE_BLANK _T("Parameter #3 must not be blank in this case.")
#define ERR_PARAM4_MUST_NOT_BE_BLANK _T("Parameter #4 must not be blank in this case.")
#define ERR_MISSING_OUTPUT_VAR _T("Requires at least one of its output variables.")
#define ERR_MISSING_OPEN_PAREN _T("Missing \"(\"")
#define ERR_MISSING_OPEN_BRACE _T("Missing \"{\"")
#define ERR_MISSING_CLOSE_PAREN _T("Missing \")\"")
#define ERR_MISSING_CLOSE_BRACE _T("Missing \"}\"")
#define ERR_MISSING_CLOSE_BRACKET _T("Missing \"]\"") // L31
#define ERR_UNEXPECTED_OPEN_BRACE _T("Unexpected \"{\"")
#define ERR_UNEXPECTED_CLOSE_PAREN _T("Unexpected \")\"")
#define ERR_UNEXPECTED_CLOSE_BRACKET _T("Unexpected \"]\"")
#define ERR_UNEXPECTED_CLOSE_BRACE _T("Unexpected \"}\"")
#define ERR_UNEXPECTED_COMMA _T("Unexpected comma")
#define ERR_BAD_AUTO_CONCAT _T("Missing space or operator before this.")
#define ERR_MISSING_CLOSE_QUOTE _T("Missing close-quote") // No period after short phrases.
#define ERR_MISSING_COMMA _T("Missing comma")             //
#define ERR_MISSING_COLON _T("Missing \":\"")             //
#define ERR_MISSING_PARAM_NAME _T("Missing parameter name.")
#define ERR_PARAM_REQUIRED _T("Missing a required parameter.")
#define ERR_TOO_MANY_PARAMS _T("Too many parameters passed to function.") // L31
#define ERR_TOO_FEW_PARAMS _T("Too few parameters passed to function.") // L31
#define ERR_BAD_OPTIONAL_PARAM _T("Expected \":=\"")
#define ERR_HOTKEY_FUNC_PARAMS _T("Only the first parameter of a hotkey function is permitted to be non-optional.")
#define ERR_HOTKEY_MISSING_BRACE _T("Hotkey or hotstring is missing its opening brace.")
#define ERR_ELSE_WITH_NO_IF _T("ELSE with no matching IF")
#define ERR_UNTIL_WITH_NO_LOOP _T("UNTIL with no matching LOOP")
#define ERR_CATCH_WITH_NO_TRY _T("CATCH with no matching TRY")
#define ERR_FINALLY_WITH_NO_PRECEDENT _T("FINALLY with no matching TRY or CATCH")
#define ERR_BAD_JUMP_INSIDE_FINALLY _T("Jumps cannot exit a FINALLY block.")
#define ERR_UNEXPECTED_CASE _T("Case/Default must be enclosed by a Switch.")
#define ERR_TOO_MANY_CASE_VALUES _T("Too many case values.")
#define ERR_EXPECTED_BLOCK_OR_ACTION _T("Expected \"{\" or single-line action.")
#define ERR_EXPECTED_ACTION _T("Expected single-line action.")
#define ERR_OUTOFMEM _T("Out of memory.")  // Used by RegEx too, so don't change it without also changing RegEx to keep the former string.
#define ERR_EXPR_TOO_LONG _T("Expression too complex")
#define ERR_TOO_MANY_REFS ERR_EXPR_TOO_LONG // No longer applies to just var/func refs. Old message: "Too many var/func refs."
#define ERR_NO_LABEL _T("Label not found in current scope.")
#define ERR_INVALID_MENU_TYPE _T("Invalid menu type.")
#define ERR_INVALID_SUBMENU _T("Invalid submenu.")
#define ERR_WINDOW_PARAM _T("Requires at least one of its window parameters.")
#define ERR_MOUSE_COORD _T("X & Y must be either both absent or both present.")
#define ERR_DIVIDEBYZERO _T("Divide by zero.")
#define ERR_EXP_ILLEGAL_CHAR _T("Illegal character in expression.")
#define ERR_UNQUOTED_NON_ALNUM _T("Unquoted literals may only consist of alphanumeric characters/underscore.")
#define ERR_DUPLICATE_DECLARATION _T("Duplicate declaration.")
#define ERR_INVALID_FUNCDECL _T("Invalid function declaration.")
#define ERR_INVALID_CLASS_VAR _T("Invalid class variable declaration.")
#define ERR_INVALID_LINE_IN_CLASS_DEF _T("Not a valid method, class or property definition.")
#define ERR_INVALID_LINE_IN_PROPERTY_DEF _T("Not a valid property getter/setter.")
#define ERR_INVALID_GUI_NAME _T("Invalid Gui name.")
#define ERR_INVALID_OPTION _T("Invalid option.") // Generic message used by the Gui system.
#define ERR_GUI_NO_WINDOW _T("Gui has no window.")
#define ERR_GUI_NOT_FOR_THIS_TYPE _T("Not supported for this control type.") // Used by GuiControl object and Control functions.
#define ERR_MUST_DECLARE _T("This variable must be declared.")
#define ERR_REMOVE_THE_PERCENT _T("If this variable was not intended to be dynamic, remove the % symbols from it.")
#define ERR_DYNAMIC_TOO_LONG _T("This dynamically built variable name is too long.  ") ERR_REMOVE_THE_PERCENT
#define ERR_DYNAMIC_BLANK _T("This dynamic variable is blank.  ") ERR_REMOVE_THE_PERCENT
#define ERR_HOTKEY_IF_EXPR _T("Parameter #1 must match an existing #HotIf expression.")
#define ERR_EXCEPTION _T("An exception was thrown.")
#define ERR_INVALID_ASSIGNMENT _T("Invalid assignment.")
#define ERR_EXPR_EVAL _T("Error evaluating expression.")
#define ERR_EXPR_SYNTAX _T("Syntax error.")
#define ERR_EXPR_MISSING_OPERAND _T("Missing operand.")
#define ERR_TYPE_MISMATCH _T("Type mismatch.")
#define ERR_NOT_ENUMERABLE _T("Value not enumerable.")
#define ERR_PROPERTY_READONLY _T("Property is read-only.")
#define ERR_NO_KEY _T("Key not found.")
#define ERR_NO_WINDOW _T("Target window not found.")
#define ERR_NO_CONTROL _T("Target control not found.")
#define ERR_NO_GUI _T("No default GUI.")
#define ERR_NO_STATUSBAR _T("No StatusBar.")
#define ERR_NO_LISTVIEW _T("No ListView.")
#define ERR_NO_TREEVIEW _T("No TreeView.")
#define ERR_WINDOW_HAS_NO_MENU _T("Non-existent or unsupported menu.")
#define ERR_PCRE_EXEC _T("PCRE execution error.")
#define ERR_INVALID_ARG_TYPE _T("Invalid arg type.")
#define ERR_INVALID_RETURN_TYPE _T("Invalid return type.")
#define ERR_INVALID_LENGTH _T("Invalid Length.")
#define ERR_INVALID_ENCODING _T("Invalid Encoding.")
#define ERR_INVALID_USAGE _T("Invalid usage.")
#define ERR_INVALID_BASE _T("Invalid base.")
#define ERR_INTERNAL_CALL _T("An internal function call failed.") // Win32 function failed.  Eventually an error message should be generated based on GetLastError().
#define ERR_FAILED _T("Failed") // A function failed to achieve its primary purpose for unspecified reason.  Equivalent to v1 throwing 1 (ErrorLevel).
#define ERR_STRING_NOT_TERMINATED _T("String not null-terminated.")
#define ERR_SOUND_DEVICE _T("Device not found")
#define ERR_SOUND_COMPONENT _T("Component not found")
#define ERR_SOUND_CONTROLTYPE _T("Component doesn't support this control type")
#define ERR_TIMEOUT _T("Timeout")
#define WARNING_USE_UNSET_VARIABLE _T("This variable has not been assigned a value.")
#define WARNING_LOCAL_SAME_AS_GLOBAL _T("This local variable has the same name as a global variable.")
#define WARNING_USE_ENV_VARIABLE _T("An environment variable is being accessed; see #NoEnv.")
#define WARNING_CLASS_OVERWRITE _T("Class may be overwritten.")

//----------------------------------------------------------------------------------

void DoIncrementalMouseMove(int aX1, int aY1, int aX2, int aY2, int aSpeed);

DWORD ProcessExist(LPTSTR aProcess);
DWORD GetProcessName(DWORD aProcessID, LPTSTR aBuf, DWORD aBufSize, bool aGetNameOnly);

bool Util_Shutdown(int nFlag);
BOOL Util_ShutdownHandler(HWND hwnd, DWORD lParam);
void Util_WinKill(HWND hWnd);

enum MainWindowModes {MAIN_MODE_NO_CHANGE, MAIN_MODE_LINES, MAIN_MODE_VARS
	, MAIN_MODE_HOTKEYS, MAIN_MODE_KEYHISTORY, MAIN_MODE_REFRESH};
ResultType ShowMainWindow(MainWindowModes aMode = MAIN_MODE_NO_CHANGE, bool aRestricted = true);
DWORD GetAHKInstallDir(LPTSTR aBuf);


struct InputBoxType
{
	LPTSTR title;
	LPTSTR text;
	LPTSTR default_string;
	LPTSTR return_string;
	int width;
	int height;
	int xpos;
	int ypos;
	TCHAR password_char;
	bool set_password_char;
	DWORD timeout;
	HWND hwnd;

	ResultType UpdateResult(HWND hControl);
};

// From AutoIt3's InputBox.  This doesn't add a measurable amount of code size, so the compiler seems to implement
// it efficiently (somewhat like a macro).
template <class T>
inline void swap(T &v1, T &v2) {
	T tmp=v1;
	v1=v2;
	v2=tmp;
}

// The following functions are used in GUI DPI scaling, so that
// GUIs designed for a 96 DPI setting (i.e. using absolute coords
// or explicit widths/sizes) can continue to run with mostly no issues.

static inline int DPIScale(int x)
{
	extern int g_ScreenDPI;
	return MulDiv(x, g_ScreenDPI, 96);
}

static inline int DPIUnscale(int x)
{
	extern int g_ScreenDPI;
	return MulDiv(x, 96, g_ScreenDPI);
}

#define INPUTBOX_DEFAULT INT_MIN
ResultType InputBoxParseOptions(LPTSTR aOptions, InputBoxType &aInputBox);
ResultType InputBox(Var *aOutputVar, LPTSTR aText, LPTSTR aTitle, LPTSTR aOptions, LPTSTR aDefault);
INT_PTR CALLBACK InputBoxProc(HWND hWndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
VOID CALLBACK InputBoxTimeout(HWND hWnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
VOID CALLBACK DerefTimeout(HWND hWnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
BOOL CALLBACK EnumChildFindSeqNum(HWND aWnd, LPARAM lParam);
BOOL CALLBACK EnumChildFindPoint(HWND aWnd, LPARAM lParam);
BOOL CALLBACK EnumChildGetControlList(HWND aWnd, LPARAM lParam);
BOOL CALLBACK EnumMonitorProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM lParam);
BOOL CALLBACK EnumChildGetText(HWND aWnd, LPARAM lParam);
LRESULT CALLBACK MainWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam);
bool HandleMenuItem(HWND aHwnd, WORD aMenuItemID, HWND aGuiHwnd);
INT_PTR CALLBACK TabDialogProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
#define TABDIALOG_ATTRIB_INDEX(a) (TabControlIndexType)(a & 0xFF)
#define TABDIALOG_ATTRIB_THEMED 0x100


typedef UINT LineNumberType;
typedef WORD FileIndexType; // Use WORD to conserve memory due to its use in the Line class (adjacency to other members and due to 4-byte struct alignment).
#define ABSOLUTE_MAX_SOURCE_FILES 0xFFFF // Keep this in sync with the capacity of the type above.  Actually it could hold 0xFFFF+1, but avoid the final item for maintainability (otherwise max-index won't be able to fit inside a variable of that type).

#define LOADING_FAILED UINT_MAX

// -2 for the beginning and ending g_DerefChars:
#define MAX_VAR_NAME_LENGTH (UCHAR_MAX - 2)
#define MAX_FUNCTION_PARAMS UCHAR_MAX // Also conserves stack space to support future attributes such as param default values.

typedef UINT DerefLengthType;
typedef int DerefParamCountType;

// Traditionally DerefType was used to hold var and func references, which are parsed at an
// early stage, but when the capability to nest expressions between percent signs was added,
// it became necessary to pre-parse more.  All non-numeric operands are represented in it.
enum DerefTypeType : BYTE
{
	DT_VAR,			// Variable reference, including built-ins.
	DT_DOUBLE,		// Marks the end of a double-deref.
	DT_STRING,		// Segment of text in a text arg (delimited by '%').
	DT_QSTRING,		// Segment of text in a quoted string (delimited by '%').
	DT_WORDOP,		// Word operator: and, or, not, new.
	DT_CONST_INT,	// Constant integer value (true, false).
	DT_DOTPERCENT,	// Dynamic member: .%name%
	DT_FUNCREF,		// Reference to function (for fat arrow functions).
	// DerefType::is_function() requires that these are last:
	DT_FUNC,		// Function call.
	DT_VARIADIC		// Variadic function call.
};

template<class T>
class SimpleList
{
	/*
		To keep a simple "auto expanding" and "searchable" list.
		Construct with aFreeItems = false to avoid calling the virtual FreeItem method in each item.
		Derived classes should implement AreEqual and / or FreeItem as desired. See HasItem and ~SimpleList()

		Methods:
		AddItem(T t)	// add an item - returns 0 on failure else the number of items in the list
		HasItem(T t)	// returns true if the item t is in the list, else false. AreEqual() determines if two items of type T are equal.

	*/
	bool mFreeItems;	// to indicate wheter to call free on each item in the list or not.
	T* mList;			// the list
	int mLastIndex;		// the number of items in the list.
public:

	SimpleList(bool aFreeItems = false) : mList(NULL), mLastIndex(0), mFreeItems(aFreeItems) {};

	~SimpleList()
	{
		// Calls FreeItem for each item in the list if appropriate, frees the list.
		if (!mList) return;
		if (mFreeItems)
			for (int i = 0; i < mLastIndex; ++i)
				FreeItem(mList[i]);
		free(mList);
	}

	int AddItem(T t)
	{
		// returns 0 on memory allocation failure.
		// Else it returns the number of elements in the list.
		T* new_list = (T*)realloc(mList, (mLastIndex + 1) * sizeof T);
		if (!new_list)
			return 0;
		mList = new_list;
		mList[mLastIndex] = t;
		return ++mLastIndex;
	}
	bool HasItem(T t)
	{
		// returns true if t is in the list, else false.
		for (int i = 0; i < mLastIndex; ++i)
			if (AreEqual(mList[i], t))	// Virtual method, derived classes should define if appropriate.
				return true;
		return false;
	}
	T GetItem(int aIndex, bool* apWasFound = NULL)
	{
		if (aIndex >= mLastIndex || aIndex < 0 || mLastIndex == 0) // bound check
		{
			if (apWasFound) *apWasFound = false;
			return (T)NULL;
		}
		if (apWasFound) *apWasFound = true;
		return mList[aIndex];
	}
	virtual bool AreEqual(T t1, T t2) { return t1 == t2; } // default comparison.
	virtual void FreeItem(T t) {}; // does nothing
};

class Func; // Forward declaration for use below.
struct DerefType
{
	LPTSTR marker;
	union
	{
		Var *var; // DT_VAR
		Func *func; // DT_FUNC
		DerefType *next; // DT_STRING
		SymbolType symbol; // DT_WORDOP
		int int_value; // DT_CONST_INT
	};
	DerefParamCountType param_count; // The actual number of parameters present in this function *call*.  Left uninitialized except for functions.
	// Keep any fields that aren't an even multiple of 4 adjacent to each other.  This conserves memory
	// due to byte-alignment:
	DerefTypeType type;
	bool is_function() { return type >= DT_FUNC; }
	DerefLengthType length; // Listed only after byte-sized fields, due to it being a WORD.
};

typedef UCHAR ArgTypeType;  // UCHAR vs. an enum, to save memory.
typedef UINT ArgLengthType;
#define ARG_TYPE_NORMAL     (UCHAR)0
#define ARG_TYPE_INPUT_VAR  (UCHAR)1
#define ARG_TYPE_OUTPUT_VAR (UCHAR)2
#define ARGMAP_END_MARKER ((ArgLengthType)~0) // ExpressionToPostfix() may rely on this being greater than any possible arg character offset.

struct ArgStruct
{
	ArgTypeType type;
	bool is_expression; // Whether this ARG is known to contain an expression.
	// Above are kept adjacent to each other to conserve memory (any fields that aren't an even
	// multiple of 4, if adjacent to each other, consume less memory due to default byte alignment
	// setting [which helps performance]).
	ArgLengthType length; // Keep adjacent to above so that it uses no extra memory. This member was added in v1.0.44.14 to improve runtime performance.
	LPTSTR text;
	DerefType *deref;  // Will hold a NULL-terminated array of operands/word-operators pre-parsed by ParseDerefs()/ParseOperands().
	ExprTokenType *postfix;  // An array of tokens in postfix order.
	int max_stack, max_alloc;
};

__int64 pow_ll(__int64 base, __int64 exp); // integer power function

#define BIF_DECL_PARAMS ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount

// The following macro is used for definitions and declarations of built-in functions:
#define BIF_DECL(name) void name(BIF_DECL_PARAMS)

#define _f__oneline(act)		do { act } while (0)		// Make the macro safe to use like a function, under if(), etc.
#define _f__ret(act)			_f__oneline( act; return; )	// BIFs have no return value.
#define _o__ret(act)			return (act)				// IObject::Invoke() returns ResultType.
// The following macros are used in built-in functions and objects to reduce code repetition
// and facilitate changes to the script "ABI" (i.e. the way in which parameters and return
// values are passed around).  For instance, the built-in functions might someday be exposed
// via COM IDispatch or coupled with different scripting languages.
#define _f_return(...)			_f__ret(aResultToken.Return(__VA_ARGS__))
#define _o_return(...)			_o__ret(aResultToken.Return(__VA_ARGS__))
#define _f_throw(...)			_f__ret(aResultToken.Error(__VA_ARGS__))
#define _f_throw_win32(...)		return ((void)aResultToken.Win32Error(__VA_ARGS__))
#define _f_throw_type(...)		return ((void)aResultToken.TypeError(__VA_ARGS__))
#define _o_throw(...)			_o__ret(aResultToken.Error(__VA_ARGS__))
#define _o_throw_win32(...)		return aResultToken.Win32Error(__VA_ARGS__)
#define _f_return_FAIL			_f__ret(aResultToken.SetExitResult(FAIL))
#define _o_return_FAIL			_o__ret(aResultToken.SetExitResult(FAIL))
// The _f_set_retval macros should be used with care because the integer macros assume symbol
// is set to its default value; i.e. don't set a string and then attempt to return an integer.
// It is also best for maintainability to avoid setting mem_to_free or an object without
// returning if there's any chance _f_throw() will be used, since in that case the caller
// may or may not Free() the result.  _f_set_retval_i() may also invalidate _f_retval_buf.
//#define _f_set_retval(...)		aResultToken.Return(__VA_ARGS__)  // Overrides the default return value but doesn't return.
#define _f_set_retval_i(n)		(aResultToken.value_int64 = static_cast<__int64>(n)) // Assumes symbol == SYM_INTEGER, the default for BIFs.
#define _f_set_retval_p(...)	aResultToken.ReturnPtr(__VA_ARGS__) // Overrides the default return value but doesn't return.  Arg must already be in persistent memory.
#define _f_return_i(n)			_f__ret(_f_set_retval_i(n)) // Return an integer.  Reduces code size vs _f_return() by assuming symbol == SYM_INTEGER, the default for BIFs.
#define _f_return_b(b)			_f_return_i((bool)(b)) // Boolean.  Currently just returns an int because we have no boolean type.
#define _f_return_p(...)		_f__ret(_f_set_retval_p(__VA_ARGS__)) // Return a string which is already in persistent memory.
#define _o_return_p(...)		_o__ret(_f_set_retval_p(__VA_ARGS__)) // Return a string which is already in persistent memory.
#define _f_return_retval		return  // Return the value set by _f_set_retval().
#define _o_return_retval		return OK
#define _f_return_empty			_f_return_p(_T(""), 0)
#define _o_return_empty			return OK  // Default return value for Invoke is "".
#define _o_return_or_throw(p)	if (p) _o_return(p); else _o_throw(ERR_OUTOFMEM);
#define _f_retval_buf			(aResultToken.buf)
#define _f_retval_buf_size		MAX_NUMBER_SIZE
#define _f_number_buf			_f_retval_buf  // An alias to show intended usage, and in case the buffer size is changed.
#define _f_callee_id			(aResultToken.func->mFID)


struct LoopFilesStruct : WIN32_FIND_DATA
{
	// Note that using fixed buffer sizes significantly reduces code size vs. using CString
	// or probably any other method of dynamically allocating/expanding the buffers.  It also
	// performs marginally better, but file system performance has a much bigger impact.
	// Unicode builds allow for the maximum path size supported by Win32 as of 2018, although
	// in some cases the script might need to use the \\?\ prefix to go beyond MAX_PATH.
	// On Windows 10 v1607+, MAX_PATH limits can be lifted by opting-in to long path support
	// via the application's manifest and LongPathsEnabled registry setting.  In any case,
	// ANSI APIs are still limited to MAX_PATH, but MAX_PATH*2 allows for the longest path
	// supported by FindFirstFile() concatenated with the longest filename it can return.
	// This preserves backward-compatibility under the following set of conditions:
	//  1) the absolute path and pattern fits within MAX_PATH;
	//  2) the relative path and filename fits within MAX_PATH; and
	//  3) the absolute path and filename exceeds MAX_PATH.
	static const size_t BUF_SIZE = UorA(MAX_WIDE_PATH, MAX_PATH*2);
	// file_path contains the full path of the directory being looped, with trailing slash.
	// Temporarily also contains the pattern for FindFirstFile(), which is either a copy of
	// 'pattern' or "*" for scanning sub-directories.
	// During execution of the loop body, it contains the full path of the file.
	TCHAR file_path[BUF_SIZE];
	TCHAR pattern[MAX_PATH]; // Naked filename or pattern.  Allows max NTFS filename length plus a few chars.
	TCHAR short_path[BUF_SIZE]; // Short name version of orig_dir.
	TCHAR *file_path_suffix; // The dynamic part of file_path (used by A_LoopFilePath).
	TCHAR *orig_dir; // Initial directory as specified by caller (used by A_LoopFilePath).
	TCHAR *long_dir; // Full/long path of initial directory (used by A_LoopFileLongPath).
	size_t file_path_length, pattern_length, short_path_length, orig_dir_length, long_dir_length
		, dir_length; // Portion of file_path which is the directory, used by BIVs.

	LoopFilesStruct() : orig_dir_length(0), long_dir(NULL) {}
	~LoopFilesStruct()
	{
		if (orig_dir_length)
			free(orig_dir);
		//else: orig_dir is the constant _T("").
		free(long_dir);
	}
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
	TCHAR subkey[MAX_REG_ITEM_SIZE];  // The branch of the registry where this subkey or value is located.
	TCHAR name[MAX_REG_ITEM_SIZE]; // The subkey or value name.
	DWORD type; // Value Type (e.g REG_DWORD).
	FILETIME ftLastWriteTime; // Non-initialized.
	void InitForValues() {ftLastWriteTime.dwHighDateTime = ftLastWriteTime.dwLowDateTime = 0;}
	void InitForSubkeys() {type = REG_SUBKEY;}  // To distinguish REG_DWORD and such from the subkeys themselves.
	RegItemStruct(HKEY aRootKeyType, HKEY aRootKey, LPTSTR aSubKey)
		: root_key_type(aRootKeyType), root_key(aRootKey), type(REG_NONE)
	{
		*name = '\0';
		// Make a local copy on the caller's stack so that if the current script subroutine is
		// interrupted to allow another to run, the contents of the deref buffer is saved here:
		tcslcpy(subkey, aSubKey, _countof(subkey));
		// Even though the call may work with a trailing backslash, it's best to remove it
		// so that consistent results are delivered to the user.  For example, if the script
		// is enumerating recursively into a subkey, subkeys deeper down will not include the
		// trailing backslash when they are reported.  So the user's own subkey should not
		// have one either so that when A_ScriptSubKey is referenced in the script, it will
		// always show up as the value without a trailing backslash:
		size_t length = _tcslen(subkey);
		if (length && subkey[length - 1] == '\\')
			subkey[length - 1] = '\0';
	}
};

class TextStream; // TextIO
struct LoopReadFileStruct
{
	TextStream *mReadFile, *mWriteFile;
	LPTSTR mWriteFileName;
	#define READ_FILE_LINE_SIZE (64 * 1024)
	TCHAR mCurrentLine[READ_FILE_LINE_SIZE];
	LoopReadFileStruct(TextStream *aReadFile, LPTSTR aWriteFileName)
		: mReadFile(aReadFile), mWriteFile(NULL) // mWriteFile is opened by FileAppend() only upon first use.
		, mWriteFileName(aWriteFileName) // Caller has passed the result of _tcsdup() for us to take over.
	{
		*mCurrentLine = '\0';
	}
	~LoopReadFileStruct()
	{
		free(mWriteFileName);
	}
};

// TextStream flags for LoadIncludedFile (script files) and file-reading loops.
// Do not lock read/write: older versions used fopen(), which is implicitly permissive.
#define DEFAULT_READ_FLAGS (TextStream::READ | TextStream::EOL_CRLF | TextStream::EOL_ORPHAN_CR | TextStream::SHARE_READ | TextStream::SHARE_WRITE)


typedef UCHAR ArgCountType;
#define MAX_ARGS 20   // Maximum number of args used by any command.


enum DllArgTypes {
	  DLL_ARG_INVALID
	, DLL_ARG_ASTR
	, DLL_ARG_INT
	, DLL_ARG_SHORT
	, DLL_ARG_CHAR
	, DLL_ARG_INT64
	, DLL_ARG_FLOAT
	, DLL_ARG_DOUBLE
	, DLL_ARG_WSTR
	, DLL_ARG_STR  = UorA(DLL_ARG_WSTR, DLL_ARG_ASTR)
	, DLL_ARG_xSTR = UorA(DLL_ARG_ASTR, DLL_ARG_WSTR) // To simplify some sections.
};  // Some sections might rely on DLL_ARG_INVALID being 0.


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


// Each line in the enumeration below corresponds to a group of built-in functions (defined
// in g_BIF) which are implemented using a single C++ function.  These IDs are passed to the
// C++ function to tell it which function is being called.  Each group starts at ID 0 in case
// it helps the compiler to reduce code size.
enum BuiltInFunctionID {
	FID_DllCall = 0, FID_ComCall,
	FID_Func = 0, FID_FuncClose,
	FID_LV_GetNext = 0, FID_LV_GetCount,
	FID_LV_Add = 0, FID_LV_Insert, FID_LV_Modify,
	FID_LV_InsertCol = 0, FID_LV_ModifyCol, FID_LV_DeleteCol,
	FID_TV_Add = 0, FID_TV_Modify, FID_TV_Delete,
	FID_TV_GetNext = 0, FID_TV_GetPrev, FID_TV_GetParent, FID_TV_GetChild, FID_TV_GetSelection, FID_TV_GetCount,
	FID_TV_Get = 0, FID_TV_GetText,
	FID_SB_SetText = 0, FID_SB_SetParts, FID_SB_SetIcon,
	FID_Trim = 0, FID_LTrim, FID_RTrim,
	FID_RegExMatch = 0, FID_RegExReplace,
	FID_Input = 0, FID_InputEnd,
	FID_GetKeyName = 0, FID_GetKeyVK = 1, FID_GetKeySC,
	FID_StrLower = 0, FID_StrUpper,
	FID_StrGet = 0, FID_StrPut,
	FID_FileExist = 0, FID_DirExist,
	FID_WinExist = 0, FID_WinActive,
	FID_Floor = 0, FID_Ceil,
	FID_ASin = 0, FID_ACos,
	FID_Sqrt = 0, FID_Log, FID_Ln,
	FID_Min = 0, FID_Max,
	FID_Random = 0, FID_RandomSeed,
	FID_ObjAddRef = 0, FID_ObjRelease,
	FID_ObjHasOwnProp = 0, FID_ObjOwnPropCount, FID_ObjGetCapacity, FID_ObjSetCapacity, FID_ObjOwnProps, FID_ObjOwnMethods,
	FID_ObjGetBase = 0, FID_ObjSetBase,
	FID_ObjPtr = 0, FID_ObjPtrAddRef, FID_ObjFromPtr, FID_ObjFromPtrAddRef,
	FID_WinGetID = 0, FID_WinGetIDLast, FID_WinGetPID, FID_WinGetProcessName, FID_WinGetProcessPath, FID_WinGetCount, FID_WinGetList, FID_WinGetMinMax, FID_WinGetControls, FID_WinGetControlsHwnd, FID_WinGetTransparent, FID_WinGetTransColor, FID_WinGetStyle, FID_WinGetExStyle,
	FID_WinGetPos = 0, FID_WinGetClientPos,
	FID_WinSetTransparent = 0, FID_WinSetTransColor, FID_WinSetAlwaysOnTop, FID_WinSetStyle, FID_WinSetExStyle, FID_WinSetEnabled, FID_WinSetRegion,
	FID_WinMoveBottom = 0, FID_WinMoveTop,
	FID_WinShow = 0, FID_WinHide, FID_WinMinimize, FID_WinMaximize, FID_WinRestore, FID_WinClose, FID_WinKill,
	FID_WinActivate = 0, FID_WinActivateBottom,
	FID_ProcessExist = 0, FID_ProcessClose, FID_ProcessWait, FID_ProcessWaitClose, 
	FID_MonitorGet = 0, FID_MonitorGetWorkArea, FID_MonitorGetCount, FID_MonitorGetPrimary, FID_MonitorGetName, 
	FID_OnExit = 0, FID_OnClipboardChange, FID_OnError,
	FID_ControlGetChecked = 0, FID_ControlGetEnabled, FID_ControlGetVisible, FID_ControlFindItem, FID_ControlGetIndex, FID_ControlGetChoice, FID_ControlGetItems, FID_ListViewGetContent, FID_EditGetLineCount, FID_EditGetCurrentLine, FID_EditGetCurrentCol, FID_EditGetLine, FID_EditGetSelectedText, FID_ControlGetStyle, FID_ControlGetExStyle, FID_ControlGetHwnd,
	FID_ControlSetChecked = 0, FID_ControlSetEnabled, FID_ControlShow, FID_ControlHide, FID_ControlSetStyle, FID_ControlSetExStyle, FID_ControlShowDropDown, FID_ControlHideDropDown, FID_ControlAddItem, FID_ControlDeleteItem, FID_ControlChooseIndex, FID_ControlChooseString, FID_EditPaste,
	FID_ControlSend = SCM_NOT_RAW, FID_ControlSendText = SCM_RAW_TEXT,
	FID_DriveEject = 0, FID_DriveLock, FID_DriveUnlock, FID_DriveSetLabel,
	FID_DriveGetList = 0, FID_DriveGetFilesystem, FID_DriveGetLabel, FID_DriveGetSerial, FID_DriveGetType, FID_DriveGetStatus, FID_DriveGetStatusCD, FID_DriveGetCapacity, FID_DriveGetSpaceFree,
	FID_EnvGet = 0, FID_EnvSet,
	FID_PostMessage = 0, FID_SendMessage,
	FID_RegRead = 0, FID_RegWrite, FID_RegDelete, FID_RegDeleteKey,
	FID_SoundGetVolume = 0, FID_SoundGetMute, FID_SoundGetName, FID_SoundGetInterface, FID_SoundSetVolume, FID_SoundSetMute,
	FID_RunWait = 0, FID_ClipWait, FID_KeyWait, FID_WinWait, FID_WinWaitClose, FID_WinWaitActive, FID_WinWaitNotActive,
	// For BIF_SetBIV (functions corresponding to built-in vars): keep the order of these in sync with the array in BIF_SetBIV.
	FID_DetectHiddenText = 0, FID_DetectHiddenWindows, FID_FileEncoding, FID_SetRegView, FID_SetStoreCapsLockMode, FID_SetTitleMatchMode, FID_StringCaseSense,
	// Hotkey/HotIf/...
	FID_HotIfWinActive = HOT_IF_ACTIVE, FID_HotIfWinNotActive = HOT_IF_NOT_ACTIVE,
		FID_HotIfWinExist = HOT_IF_EXIST, FID_HotIfWinNotExist = HOT_IF_NOT_EXIST,
		FID_Hotkey, FID_HotIf
};


typedef UCHAR GuiControls;
enum GuiControlTypes {GUI_CONTROL_INVALID // GUI_CONTROL_INVALID must be zero due to things like ZeroMemory() on the struct.
	, GUI_CONTROL_TEXT, GUI_CONTROL_PIC, GUI_CONTROL_GROUPBOX
	, GUI_CONTROL_BUTTON, GUI_CONTROL_CHECKBOX, GUI_CONTROL_RADIO
	, GUI_CONTROL_DROPDOWNLIST, GUI_CONTROL_COMBOBOX
	, GUI_CONTROL_LISTBOX, GUI_CONTROL_LISTVIEW, GUI_CONTROL_TREEVIEW
	, GUI_CONTROL_EDIT, GUI_CONTROL_DATETIME, GUI_CONTROL_MONTHCAL, GUI_CONTROL_HOTKEY
	, GUI_CONTROL_UPDOWN, GUI_CONTROL_SLIDER, GUI_CONTROL_PROGRESS, GUI_CONTROL_TAB, GUI_CONTROL_TAB2, GUI_CONTROL_TAB3
	, GUI_CONTROL_ACTIVEX, GUI_CONTROL_LINK, GUI_CONTROL_CUSTOM, GUI_CONTROL_STATUSBAR}; // Kept last to reflect it being bottommost in switch()s (for perf), since not too often used.

#define GUI_CONTROL_TYPE_NAMES  _T(""), \
	_T("Text"), _T("Pic"), _T("GroupBox"), \
	_T("Button"), _T("CheckBox"), _T("Radio"), \
	_T("DDL"), _T("ComboBox"), \
	_T("ListBox"), _T("ListView"), _T("TreeView"), \
	_T("Edit"), _T("DateTime"), _T("MonthCal"), _T("Hotkey"), \
	_T("UpDown"), _T("Slider"), _T("Progress"), _T("Tab"), _T("Tab2"), _T("Tab3"), \
	_T("ActiveX"), _T("Link"), _T("Custom"), _T("StatusBar")

enum ThreadCommands {THREAD_CMD_INVALID, THREAD_CMD_PRIORITY, THREAD_CMD_INTERRUPT, THREAD_CMD_NOTIMERS};


class Label; // Forward declaration so that each can use the other.
class Line
{
private:
	ResultType EvaluateCondition();
	bool EvaluateLoopUntil(ResultType &aResult);
	ResultType Line::PerformLoop(ResultToken *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil
		, __int64 aIterationLimit, bool aIsInfinite);
	ResultType Line::PerformLoopFilePattern(ResultToken *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil
		, FileLoopModeType aFileLoopMode, bool aRecurseSubfolders, LPTSTR aFilePattern);
	bool ParseLoopFilePattern(LPTSTR aFilePattern, LoopFilesStruct &lfs, ResultType &aResult);
	ResultType PerformLoopFilePattern(ResultToken *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil
		, FileLoopModeType aFileLoopMode, bool aRecurseSubfolders, LoopFilesStruct &lfs);
	ResultType PerformLoopReg(ResultToken *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil
		, FileLoopModeType aFileLoopMode, bool aRecurseSubfolders, HKEY aRootKeyType, HKEY aRootKey, LPTSTR aRegSubkey);
	ResultType PerformLoopParse(ResultToken *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil);
	ResultType Line::PerformLoopParseCSV(ResultToken *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil);
	ResultType PerformLoopReadFile(ResultToken *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil, TextStream *aReadFile, LPTSTR aWriteFileName);
	ResultType PerformLoopWhile(ResultToken *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine); // Lexikos: ACT_WHILE.
	ResultType PerformLoopFor(ResultToken *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil); // Lexikos: ACT_FOR.
	ResultType PerformAssign();
	ResultType Perform();
	friend BIF_DECL(BIF_PerformAction);

	ResultType SoundPlay(LPTSTR aFilespec, bool aSleepUntilDone);
	ResultType Download(LPTSTR aURL, LPTSTR aFilespec);

	ResultType FileCreateShortcut(LPTSTR aTargetFile, LPTSTR aShortcutFile, LPTSTR aWorkingDir, LPTSTR aArgs
		, LPTSTR aDescription, LPTSTR aIconFile, LPTSTR aHotkey, LPTSTR aIconNumber, LPTSTR aRunState);
	static bool FileCreateDir(LPTSTR aDirSpec, LPTSTR aCanModifyDirSpec = NULL);
	ResultType FileDelete(LPTSTR aFilePattern);
	ResultType FileRecycle(LPTSTR aFilePattern);
	ResultType FileRecycleEmpty(LPTSTR aDriveLetter);
	ResultType FileInstall(LPTSTR aSource, LPTSTR aDest, LPTSTR aFlag);

	typedef BOOL (* FilePatternCallback)(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData);
	struct FilePatternStruct
	{
		TCHAR path[T_MAX_PATH]; // Directory and naked filename or pattern.
		TCHAR pattern[MAX_PATH]; // Naked filename or pattern.
		size_t dir_length, pattern_length;
		FilePatternCallback aCallback;
		void *aCallbackData;
		FileLoopModeType aOperateOnFolders;
		bool aDoRecurse;
		int failure_count;
	};
	ResultType FilePatternApply(LPTSTR aFilePattern, FileLoopModeType aOperateOnFolders
		, bool aDoRecurse, FilePatternCallback aCallback, void *aCallbackData);
	void FilePatternApply(FilePatternStruct &);

	ResultType FileSetAttrib(LPTSTR aAttributes, LPTSTR aFilePattern
		, FileLoopModeType aOperateOnFolders, bool aDoRecurse);
	ResultType FileSetTime(LPTSTR aYYYYMMDD, LPTSTR aFilePattern, TCHAR aWhichTime
		, FileLoopModeType aOperateOnFolders, bool aDoRecurse);

	ResultType IniWrite(LPTSTR aValue, LPTSTR aFilespec, LPTSTR aSection, LPTSTR aKey);
	ResultType IniDelete(LPTSTR aFilespec, LPTSTR aSection, LPTSTR aKey);

	ResultType ToolTip(LPTSTR aText, LPTSTR aX, LPTSTR aY, LPTSTR aID);
	ResultType TrayTip(LPTSTR aText, LPTSTR aTitle, LPTSTR aOptions);

	static ResultType SetToggleState(vk_type aVK, ToggleValueType &ForceLock, LPTSTR aToggleText);

public:
	#define SET_S_DEREF_BUF(ptr, size) Line::sDerefBuf = ptr, Line::sDerefBufSize = size

	#define NULLIFY_S_DEREF_BUF \
	{\
		SET_S_DEREF_BUF(NULL, 0);\
		if (sDerefBufSize > LARGE_DEREF_BUF_SIZE)\
			--sLargeDerefBufs;\
	}

	#define PRIVATIZE_S_DEREF_BUF \
		LPTSTR our_deref_buf = Line::sDerefBuf;\
		size_t our_deref_buf_size = Line::sDerefBufSize;\
		SET_S_DEREF_BUF(NULL, 0) // For detecting whether ExpandExpression() caused a new buffer to be created.

	#define DEPRIVATIZE_S_DEREF_BUF \
		if (our_deref_buf)\
		{\
			if (Line::sDerefBuf)\
			{\
				free(Line::sDerefBuf);\
				if (Line::sDerefBufSize > LARGE_DEREF_BUF_SIZE)\
					--Line::sLargeDerefBufs;\
			}\
			SET_S_DEREF_BUF(our_deref_buf, our_deref_buf_size);\
		}
		//else the original buffer is NULL, so keep any new sDerefBuf that might have been created (should
		// help avg-case performance).

	static LPTSTR sDerefBuf;  // Buffer to hold the values of any args that need to be dereferenced.
	static size_t sDerefBufSize;
	static int sLargeDerefBufs;

	// Static because only one line can be Expanded at a time (not to mention the fact that we
	// wouldn't want the size of each line to be expanded by this size):
	static LPTSTR sArgDeref[MAX_ARGS];

	ScriptModule *mModule; // Every line belongs to one module.

	// Keep any fields that aren't an even multiple of 4 adjacent to each other.  This conserves memory
	// due to byte-alignment:
	ActionTypeType mActionType; // What type of line this is.
	ArgCountType mArgc; // How many arguments exist in mArg[].
	FileIndexType mFileIndex; // Which file the line came from.  0 is the first, and it's the main script file.
	LineNumberType mLineNumber;  // The line number in the file from which the script was loaded, for debugging.

	ArgStruct *mArg; // Will be used to hold a dynamic array of dynamic Args.
	AttributeType mAttribute;
	Line *mPrevLine, *mNextLine; // The prev & next lines adjacent to this one in the linked list; NULL if none.
	Line *mRelatedLine;  // e.g. the "else" that belongs to this "if"
	Line *mParentLine; // Indicates the parent (owner) of this line.

#ifdef CONFIG_DEBUGGER
	Breakpoint *mBreakpoint;
#endif

	// Probably best to always use ARG1 even if other things have supposedly verified
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
	#define RAW_ARG1 (mArgc > 0 ? mArg[0].text : _T(""))
	#define RAW_ARG2 (mArgc > 1 ? mArg[1].text : _T(""))
	#define RAW_ARG3 (mArgc > 2 ? mArg[2].text : _T(""))
	#define RAW_ARG4 (mArgc > 3 ? mArg[3].text : _T(""))
	#define RAW_ARG5 (mArgc > 4 ? mArg[4].text : _T(""))
	#define RAW_ARG6 (mArgc > 5 ? mArg[5].text : _T(""))
	#define RAW_ARG7 (mArgc > 6 ? mArg[6].text : _T(""))
	#define RAW_ARG8 (mArgc > 7 ? mArg[7].text : _T(""))

	#define LINE_RAW_ARG1 (line->mArgc > 0 ? line->mArg[0].text : _T(""))
	#define LINE_RAW_ARG2 (line->mArgc > 1 ? line->mArg[1].text : _T(""))
	#define LINE_RAW_ARG3 (line->mArgc > 2 ? line->mArg[2].text : _T(""))
	#define LINE_RAW_ARG4 (line->mArgc > 3 ? line->mArg[3].text : _T(""))
	#define LINE_RAW_ARG5 (line->mArgc > 4 ? line->mArg[4].text : _T(""))
	#define LINE_RAW_ARG6 (line->mArgc > 5 ? line->mArg[5].text : _T(""))
	#define LINE_RAW_ARG7 (line->mArgc > 6 ? line->mArg[6].text : _T(""))
	#define LINE_RAW_ARG8 (line->mArgc > 7 ? line->mArg[7].text : _T(""))
	#define LINE_RAW_ARG9 (line->mArgc > 8 ? line->mArg[8].text : _T(""))
	
	#define NEW_RAW_ARG1 (aArgc > 0 ? new_arg[0].text : _T("")) // Helps performance to use this vs. LINE_RAW_ARG where possible.
	#define NEW_RAW_ARG2 (aArgc > 1 ? new_arg[1].text : _T(""))
	#define NEW_RAW_ARG3 (aArgc > 2 ? new_arg[2].text : _T(""))
	#define NEW_RAW_ARG4 (aArgc > 3 ? new_arg[3].text : _T(""))
	#define NEW_RAW_ARG5 (aArgc > 4 ? new_arg[4].text : _T(""))
	#define NEW_RAW_ARG6 (aArgc > 5 ? new_arg[5].text : _T(""))
	#define NEW_RAW_ARG7 (aArgc > 6 ? new_arg[6].text : _T(""))
	#define NEW_RAW_ARG8 (aArgc > 7 ? new_arg[7].text : _T(""))
	#define NEW_RAW_ARG9 (aArgc > 8 ? new_arg[8].text : _T(""))
	
	#define SAVED_ARG1 (mArgc > 0 ? arg[0] : _T(""))
	#define SAVED_ARG2 (mArgc > 1 ? arg[1] : _T(""))
	#define SAVED_ARG3 (mArgc > 2 ? arg[2] : _T(""))
	#define SAVED_ARG4 (mArgc > 3 ? arg[3] : _T(""))
	#define SAVED_ARG5 (mArgc > 4 ? arg[4] : _T(""))

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
	static LPTSTR sSourceFile[1]; // Only need to be able to hold the main script since compiled scripts don't support dynamic including.
#else
	static LPTSTR *sSourceFile;   // Will hold an array of strings.
	static int sMaxSourceFiles;  // Maximum number of items it can currently hold.
#endif
	static int sSourceFileCount; // Number of items in the above array.

	static void FreeDerefBufIfLarge();

	ResultType ExecUntil(ExecUntilMode aMode, ResultToken *aResultToken = NULL, Line **apJumpToLine = NULL);

	// The following are characters that can't legally occur after a binary operator.  It excludes all unary operators
	// "!~*&-+" as well as the parentheses chars "()":
	#define EXPR_CORE _T("<>=/|^,?:")
	// The characters common to both EXPR_TELLTALES (obsolete) and EXPR_OPERAND_TERMINATORS:
	#define EXPR_COMMON _T(" \t") EXPR_CORE _T("*&~!()[]{}")  // Space and Tab are included at the beginning for performance.  L31: Added [] for array-like syntax.
	#define CONTINUATION_LINE_SYMBOLS EXPR_CORE _T(".+-*&!~") // v1.0.46.
	#define EXPR_OPERATOR_SYMBOLS CONTINUATION_LINE_SYMBOLS  // The set of operator symbols which can't appear at the end of a valid expression, plus '+' and '-' (which are valid for ++/--).
	// Characters that mark the end of an operand inside an expression.  Double-quote must not be included:
	#define EXPR_OPERAND_TERMINATORS_EX_DOT EXPR_COMMON _T("%+-\n") // L31: Used in a few places where '.' needs special treatment.
	#define EXPR_OPERAND_TERMINATORS EXPR_OPERAND_TERMINATORS_EX_DOT _T(".") // L31: Used in expressions where '.' is always an operator.
	#define EXPR_ALL_SYMBOLS EXPR_OPERAND_TERMINATORS _T("\"'")
	// The following HOTSTRING option recognizer is kept somewhat forgiving/non-specific for backward compatibility
	// (e.g. scripts may have some invalid hotstring options, which are simply ignored).  This definition is here
	// because it's related to continuation line symbols. Also, avoid ever adding "&" to hotstring options because
	// it might introduce ambiguity in the differentiation of things like:
	//    : & x::hotkey action
	//    : & *::abbrev with leading colon::
	#define IS_HOTSTRING_OPTION(chr) (cisalnum(chr) || _tcschr(_T("?*- \t"), chr))

	#define ArgLength(aArgNum) ArgIndexLength((aArgNum)-1)
	#define ArgToInt64(aArgNum) ArgIndexToInt64((aArgNum)-1)
	#define ArgToInt(aArgNum) (int)ArgToInt64(aArgNum) // Benchmarks show that having a "real" ArgToInt() that calls ATOI vs. ATOI64 (and ToInt() vs. ToInt64()) doesn't measurably improve performance.
	#define ArgToUInt(aArgNum) (UINT)ArgToInt64(aArgNum) // Similar to what ATOU() does.
	__int64 ArgIndexToInt64(int aArgIndex);
	size_t ArgIndexLength(int aArgIndex);

	ResultType ExpandArgs(ResultToken *aResultTokens = NULL);
	VarSizeType GetExpandedArgSize();
	LPTSTR ExpandExpression(int aArgIndex, ResultType &aResult, ResultToken *aResultToken
		, LPTSTR &aTarget, LPTSTR &aDerefBuf, size_t &aDerefBufSize, LPTSTR aArgDeref[], size_t aExtraSize);
	ResultType ExpandSingleArg(int aArgIndex, ResultToken &aResultToken, LPTSTR &aDerefBuf, size_t &aDerefBufSize);
	ResultType ExpressionToPostfix(ArgStruct &aArg);
	ResultType ExpressionToPostfix(ArgStruct &aArg, ExprTokenType *&aInfix);

	static bool FileIsFilteredOut(LoopFilesStruct &aCurrentFile, FileLoopModeType aFileLoopMode);

	Label *GetJumpTarget(bool aIsDereferenced);
	Label *IsJumpValid(Label &aTargetLabel, bool aSilent = false);
	BOOL CheckValidFinallyJump(Line* jumpTarget, bool aSilent = false);

	static HWND DetermineTargetWindow(LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText);

	
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
			case ACT_ASSIGNEXPR:
			case ACT_FOR:
			case ACT_CATCH:
				return ARG_TYPE_OUTPUT_VAR;
			}
			break;

		case 1:  // Arg #2
			switch(aActionType)
			{
			case ACT_FOR:
				return ARG_TYPE_OUTPUT_VAR;
			}
			break;
		}
		// Otherwise:
		return ARG_TYPE_NORMAL;
	}



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
			LineError(_T("DEBUG: BAD"), WARN);
			aArgIndex = 0;  // But let it continue.
		}
#endif
		if (aArgIndex >= mArgc) // Arg doesn't exist.
			return false;
		ArgStruct &arg = mArg[aArgIndex]; // For performance.
		// Return false if it's not of a type caller wants deemed to have derefs.
		if (arg.type == ARG_TYPE_NORMAL)
			return arg.deref && arg.deref[0].marker // Relies on short-circuit boolean evaluation order to prevent NULL-deref.
				|| arg.is_expression; // Return true for this case since most callers assume the arg is a simple literal string if we return false.
		else // Callers rely on input variables being seen as "true" because sometimes single isolated derefs are converted into ARG_TYPE_INPUT_VAR at load-time.
			return (arg.type == ARG_TYPE_INPUT_VAR);
	}

	static HKEY RegConvertKey(LPTSTR aBuf, LPTSTR *aSubkey = NULL, bool *aIsRemoteRegistry = NULL)
	{
		const size_t COMPUTER_NAME_BUF_SIZE = 128;

		LPTSTR key_name_pos = aBuf, computer_name_end = NULL;

		if (*aBuf == '\\' && aBuf[1] == '\\') // Something like \\ComputerName\HKLM.
		{
			if (  !(computer_name_end = _tcschr(aBuf + 2, '\\'))
				|| (computer_name_end - aBuf) >= COMPUTER_NAME_BUF_SIZE  )
				return NULL;
			key_name_pos = computer_name_end + 1;
		}

		// Copy root key name into temporary buffer for use by _tcsicmp().
		TCHAR key_name[20];
		int i;
		for (i = 0; key_name_pos[i] && key_name_pos[i] != '\\'; ++i)
		{
			if (i == 19)
				return NULL; // Too long to be valid.
			key_name[i] = key_name_pos[i];
		}
		key_name[i] = '\0';
		
		// Set output parameters for caller.
		if (aSubkey)
			*aSubkey = key_name_pos + i + (key_name_pos[i] == '\\');
		if (aIsRemoteRegistry)
			*aIsRemoteRegistry = (computer_name_end != NULL);

		HKEY root_key = RegConvertRootKeyType(key_name);
		if (!root_key) // Invalid or unsupported root key name.
			return NULL;

		if (!aIsRemoteRegistry || !computer_name_end) // Either caller didn't want it opened, or it doesn't need to be.
			return root_key; // If it's a remote key, this value should only be used by the caller as an indicator.
		// Otherwise, it's a remote computer whose registry the caller wants us to open:
		// It seems best to require the two leading backslashes in case the computer name contains
		// spaces (just in case spaces are allowed on some OSes or perhaps for Unix interoperability, etc.).
		// Therefore, make no attempt to trim leading and trailing spaces from the computer name:
		TCHAR computer_name[COMPUTER_NAME_BUF_SIZE];
		tcslcpy(computer_name, aBuf, _countof(computer_name));
		computer_name[computer_name_end - aBuf] = '\0';
		HKEY remote_key;
		return (RegConnectRegistry(computer_name, root_key, &remote_key) == ERROR_SUCCESS) ? remote_key : NULL;
	}

	static HKEY RegConvertRootKeyType(LPTSTR aName);
	static LPTSTR RegConvertRootKeyType(HKEY aKey);

	static int RegConvertValueType(LPTSTR aValueType)
	{
		if (!_tcsicmp(aValueType, _T("REG_SZ"))) return REG_SZ;
		if (!_tcsicmp(aValueType, _T("REG_EXPAND_SZ"))) return REG_EXPAND_SZ;
		if (!_tcsicmp(aValueType, _T("REG_MULTI_SZ"))) return REG_MULTI_SZ;
		if (!_tcsicmp(aValueType, _T("REG_DWORD"))) return REG_DWORD;
		if (!_tcsicmp(aValueType, _T("REG_BINARY"))) return REG_BINARY;
		return REG_NONE; // Unknown or unsupported type.
	}
	static LPTSTR RegConvertValueType(DWORD aValueType)
	{
		switch(aValueType)
		{
		case REG_SZ: return _T("REG_SZ");
		case REG_EXPAND_SZ: return _T("REG_EXPAND_SZ");
		case REG_BINARY: return _T("REG_BINARY");
		case REG_DWORD: return _T("REG_DWORD");
		case REG_DWORD_BIG_ENDIAN: return _T("REG_DWORD_BIG_ENDIAN");
		case REG_LINK: return _T("REG_LINK");
		case REG_MULTI_SZ: return _T("REG_MULTI_SZ");
		case REG_RESOURCE_LIST: return _T("REG_RESOURCE_LIST");
		case REG_FULL_RESOURCE_DESCRIPTOR: return _T("REG_FULL_RESOURCE_DESCRIPTOR");
		case REG_RESOURCE_REQUIREMENTS_LIST: return _T("REG_RESOURCE_REQUIREMENTS_LIST");
		case REG_QWORD: return _T("REG_QWORD");
		case REG_SUBKEY: return _T("KEY");  // Custom (non-standard) type.
		default: return _T("");  // Make it be the empty string for REG_NONE and anything else.
		}
	}
	static DWORD RegConvertView(LPTSTR aBuf)
	{
		if (!_tcsicmp(aBuf, _T("Default")))
			return 0;
		else if (!_tcscmp(aBuf, _T("32")))
			return KEY_WOW64_32KEY;
		else if (!_tcscmp(aBuf, _T("64")))
			return KEY_WOW64_64KEY;
		else
			return -1;
	}

	static DWORD SoundConvertControlType(LPTSTR aBuf)
	{
		// v1.0.37.06: The following was added to allow unnamed control types (if any) to be accessed via number:
		if (IsNumeric(aBuf, false, false, true)) // Seems best to allowing floating point here, since .000 on the end might happen sometimes.
			return ATOU(aBuf);
		// The following are the types that seem to correspond to actual sound attributes.  Some of the
		// values are not included here, such as MIXERCONTROL_CONTROLTYPE_FADER, which seems to be a type
		// of sound control rather than a quality of the sound itself.  For performance, put the most
		// often used ones up top.
		if (!_tcsicmp(aBuf, _T("Vol"))
			|| !_tcsicmp(aBuf, _T("Volume"))) return MIXERCONTROL_CONTROLTYPE_VOLUME;
		if (!_tcsicmp(aBuf, _T("OnOff")))     return MIXERCONTROL_CONTROLTYPE_ONOFF;
		if (!_tcsicmp(aBuf, _T("Mute")))      return MIXERCONTROL_CONTROLTYPE_MUTE;
		if (!_tcsicmp(aBuf, _T("Mono")))      return MIXERCONTROL_CONTROLTYPE_MONO;
		if (!_tcsicmp(aBuf, _T("Loudness")))  return MIXERCONTROL_CONTROLTYPE_LOUDNESS;
		if (!_tcsicmp(aBuf, _T("StereoEnh"))) return MIXERCONTROL_CONTROLTYPE_STEREOENH;
		if (!_tcsicmp(aBuf, _T("BassBoost"))) return MIXERCONTROL_CONTROLTYPE_BASS_BOOST;
		if (!_tcsicmp(aBuf, _T("Pan")))       return MIXERCONTROL_CONTROLTYPE_PAN;
		if (!_tcsicmp(aBuf, _T("QSoundPan"))) return MIXERCONTROL_CONTROLTYPE_QSOUNDPAN;
		if (!_tcsicmp(aBuf, _T("Bass")))      return MIXERCONTROL_CONTROLTYPE_BASS;
		if (!_tcsicmp(aBuf, _T("Treble")))    return MIXERCONTROL_CONTROLTYPE_TREBLE;
		if (!_tcsicmp(aBuf, _T("Equalizer"))) return MIXERCONTROL_CONTROLTYPE_EQUALIZER;
		#define MIXERCONTROL_CONTROLTYPE_INVALID 0xFFFFFFFF // 0 might be a valid type, so use something definitely undefined.
		return MIXERCONTROL_CONTROLTYPE_INVALID;
	}

	static TitleMatchModes ConvertTitleMatchMode(LPTSTR aBuf)
	{
		if (!aBuf || !*aBuf) return MATCHMODE_INVALID;
		if (*aBuf == '1' && !*(aBuf + 1)) return FIND_IN_LEADING_PART;
		if (*aBuf == '2' && !*(aBuf + 1)) return FIND_ANYWHERE;
		if (*aBuf == '3' && !*(aBuf + 1)) return FIND_EXACT;
		if (!_tcsicmp(aBuf, _T("RegEx"))) return FIND_REGEX; // Goes with the above, not fast/slow below.

		if (!_tcsicmp(aBuf, _T("FAST"))) return FIND_FAST;
		if (!_tcsicmp(aBuf, _T("SLOW"))) return FIND_SLOW;
		return MATCHMODE_INVALID;
	}

	static ThreadCommands ConvertThreadCommand(LPTSTR aBuf)
	{
		if (!aBuf || !*aBuf) return THREAD_CMD_INVALID;
		if (!_tcsicmp(aBuf, _T("Priority"))) return THREAD_CMD_PRIORITY;
		if (!_tcsicmp(aBuf, _T("Interrupt"))) return THREAD_CMD_INTERRUPT;
		if (!_tcsicmp(aBuf, _T("NoTimers"))) return THREAD_CMD_NOTIMERS;
		return THREAD_CMD_INVALID;
	}

	static ToggleValueType ConvertTrueFalse(LPTSTR aBuf, ToggleValueType aDefault = TOGGLE_INVALID)
	// Returns aDefault if aBuf isn't either ON, OFF, or blank.
	{
		if (!aBuf || !*aBuf) return NEUTRAL;
		if (!_tcsicmp(aBuf, _T("True")) || !_tcscmp(aBuf, _T("1"))) return TOGGLED_ON;
		if (!_tcsicmp(aBuf, _T("False")) || !_tcscmp(aBuf, _T("0"))) return TOGGLED_OFF;
		return aDefault;
	}

	static ToggleValueType ConvertOnOff(LPTSTR aBuf, ToggleValueType aDefault = TOGGLE_INVALID)
	// Returns aDefault if aBuf isn't either ON, OFF, or blank.
	{
		if (!aBuf || !*aBuf) return NEUTRAL;
		if (!_tcsicmp(aBuf, _T("On")) || !_tcscmp(aBuf, _T("1"))) return TOGGLED_ON;
		if (!_tcsicmp(aBuf, _T("Off")) || !_tcscmp(aBuf, _T("0"))) return TOGGLED_OFF;
		return aDefault;
	}

	static ToggleValueType ConvertOnOffAlways(LPTSTR aBuf, ToggleValueType aDefault = TOGGLE_INVALID)
	// Returns aDefault if aBuf isn't either ON, OFF, ALWAYSON, ALWAYSOFF, or blank.
	{
		if (ToggleValueType toggle = ConvertOnOff(aBuf))
			return toggle;
		if (!_tcsicmp(aBuf, _T("AlwaysOn"))) return ALWAYS_ON;
		if (!_tcsicmp(aBuf, _T("AlwaysOff"))) return ALWAYS_OFF;
		return aDefault;
	}

	static ToggleValueType Convert10Toggle(LPTSTR aBuf)
	{
		if (!aBuf || !*aBuf) return NEUTRAL;
		if (IsNumeric(aBuf, true))
			switch (ATOI(aBuf))
			{
			case 1: return TOGGLED_ON;
			case 0: return TOGGLED_OFF;
			case -1: return TOGGLE;
			}
		return TOGGLE_INVALID;
	}

	static ToggleValueType ConvertOnOffToggle(LPTSTR aBuf, ToggleValueType aDefault = TOGGLE_INVALID)
	// Returns aDefault if aBuf isn't either ON, OFF, TOGGLE, or blank.
	{
		if (ToggleValueType toggle = ConvertOnOff(aBuf))
			return toggle;
		if (!_tcsicmp(aBuf, _T("Toggle")) || !_tcscmp(aBuf, _T("-1"))) return TOGGLE;
		return aDefault;
	}

	static StringCaseSenseType ConvertStringCaseSense(LPTSTR aBuf)
	{
		if (!_tcsicmp(aBuf, _T("On")) || !_tcscmp(aBuf, _T("1"))) return SCS_SENSITIVE;
		if (!_tcsicmp(aBuf, _T("Off")) || !_tcscmp(aBuf, _T("0"))) return SCS_INSENSITIVE;
		if (!_tcsicmp(aBuf, _T("Locale"))) return SCS_INSENSITIVE_LOCALE;
		return SCS_INVALID;
	}

	static ToggleValueType ConvertBlockInput(LPTSTR aBuf)
	{
		if (ToggleValueType toggle = ConvertOnOff(aBuf))
			return toggle;
		if (!_tcsicmp(aBuf, _T("Send"))) return TOGGLE_SEND;
		if (!_tcsicmp(aBuf, _T("Mouse"))) return TOGGLE_MOUSE;
		if (!_tcsicmp(aBuf, _T("SendAndMouse"))) return TOGGLE_SENDANDMOUSE;
		if (!_tcsicmp(aBuf, _T("Default"))) return TOGGLE_DEFAULT;
		if (!_tcsicmp(aBuf, _T("MouseMove"))) return TOGGLE_MOUSEMOVE;
		if (!_tcsicmp(aBuf, _T("MouseMoveOff"))) return TOGGLE_MOUSEMOVEOFF;
		return TOGGLE_INVALID;
	}

	static SendModes ConvertSendMode(LPTSTR aBuf, SendModes aValueToReturnIfInvalid)
	{
		if (!_tcsicmp(aBuf, _T("Play"))) return SM_PLAY;
		if (!_tcsicmp(aBuf, _T("Event"))) return SM_EVENT;
		if (!_tcsnicmp(aBuf, _T("Input"), 5)) // This IF must be listed last so that it can fall through to bottom line.
		{
			aBuf += 5;
			if (!*aBuf || !_tcsicmp(aBuf, _T("ThenEvent"))) // "ThenEvent" is supported for backward compatibility with 1.0.43.00.
				return SM_INPUT;
			if (!_tcsicmp(aBuf, _T("ThenPlay")))
				return SM_INPUT_FALLBACK_TO_PLAY;
			//else fall through and return the indication of invalidity.
		}
		return aValueToReturnIfInvalid;
	}

	static FileLoopModeType ConvertLoopMode(LPTSTR aBuf)
	// Returns the file loop mode, or FILE_LOOP_INVALID if aBuf contains an invalid mode.
	{
		for (FileLoopModeType mode = FILE_LOOP_INVALID;;)
		{
			switch (ctoupper(*aBuf++))
			{
			// For simplicity, both are allowed with either kind of loop:
			case 'F': // Files
			case 'V': // Values
				mode |= FILE_LOOP_FILES_ONLY;
				break;
			case 'D': // Directories
			case 'K': // Keys
				mode |= FILE_LOOP_FOLDERS_ONLY;
				break;
			case 'R':
				mode |= FILE_LOOP_RECURSE;
				break;
			case ' ':  // Allow whitespace.
			case '\t': //
				break;
			case '\0':
				if ((mode & FILE_LOOP_FILES_AND_FOLDERS) == 0)
					mode |= FILE_LOOP_FILES_ONLY; // Set default.
				return mode;
			default: // Invalid character.
				return FILE_LOOP_INVALID;
			}
		}
	}

	static int ConvertRunMode(LPTSTR aBuf)
	// Returns the matching WinShow mode, or SW_SHOWNORMAL if none.
	// These are also the modes that AutoIt3 uses.
	{
		if (!aBuf || !*aBuf) return SW_SHOWNORMAL;
		if (!_tcsicmp(aBuf, _T("MIN"))) return SW_MINIMIZE;
		if (!_tcsicmp(aBuf, _T("MAX"))) return SW_MAXIMIZE;
		if (!_tcsicmp(aBuf, _T("HIDE"))) return SW_HIDE;
		return SW_SHOWNORMAL;
	}

	static int ConvertMouseButton(LPTSTR aBuf, bool aAllowWheel = true, bool aUseLogicalButton = true)
	// Returns the matching VK, or zero if none.
	{
		if (!*aBuf || !_tcsicmp(aBuf, _T("LEFT")) || !_tcsicmp(aBuf, _T("L")))
			return aUseLogicalButton ? VK_LBUTTON_LOGICAL : VK_LBUTTON; // Some callers rely on this default when !*aBuf.
		if (!_tcsicmp(aBuf, _T("RIGHT")) || !_tcsicmp(aBuf, _T("R"))) return aUseLogicalButton ? VK_RBUTTON_LOGICAL : VK_RBUTTON;
		if (!_tcsicmp(aBuf, _T("MIDDLE")) || !_tcsicmp(aBuf, _T("M"))) return VK_MBUTTON;
		if (!_tcsicmp(aBuf, _T("X1"))) return VK_XBUTTON1;
		if (!_tcsicmp(aBuf, _T("X2"))) return VK_XBUTTON2;
		if (aAllowWheel)
		{
			if (!_tcsicmp(aBuf, _T("WheelUp")) || !_tcsicmp(aBuf, _T("WU"))) return VK_WHEEL_UP;
			if (!_tcsicmp(aBuf, _T("WheelDown")) || !_tcsicmp(aBuf, _T("WD"))) return VK_WHEEL_DOWN;
			// Lexikos: Support horizontal scrolling in Windows Vista and later.
			if (!_tcsicmp(aBuf, _T("WheelLeft")) || !_tcsicmp(aBuf, _T("WL"))) return VK_WHEEL_LEFT;
			if (!_tcsicmp(aBuf, _T("WheelRight")) || !_tcsicmp(aBuf, _T("WR"))) return VK_WHEEL_RIGHT;
		}
		return 0;
	}

	static CoordModeType ConvertCoordMode(LPTSTR aBuf)
	{
		if (!_tcsicmp(aBuf, _T("Screen")))
			return COORD_MODE_SCREEN;
		else if (!_tcsicmp(aBuf, _T("Window")))
			return COORD_MODE_WINDOW;
		else if (!_tcsicmp(aBuf, _T("Client")))
			return COORD_MODE_CLIENT;
		return COORD_MODE_INVALID;
	}

	static CoordModeType ConvertCoordModeCmd(LPTSTR aBuf)
	{
		if (!_tcsicmp(aBuf, _T("Pixel"))) return COORD_MODE_PIXEL;
		if (!_tcsicmp(aBuf, _T("Mouse"))) return COORD_MODE_MOUSE;
		if (!_tcsicmp(aBuf, _T("ToolTip"))) return COORD_MODE_TOOLTIP;
		if (!_tcsicmp(aBuf, _T("Caret"))) return COORD_MODE_CARET;
		if (!_tcsicmp(aBuf, _T("Menu"))) return COORD_MODE_MENU;
		return COORD_MODE_INVALID;
	}

	static bool IsValidFileCodePage(UINT aCP)
	{
		return aCP == 0 || aCP == 1200 || IsValidCodePage(aCP);
	}

	static UINT ConvertFileEncoding(LPTSTR aBuf)
	// Returns the encoding with possible CP_AHKNOBOM flag, or (UINT)-1 if invalid.
	{
		if (!aBuf || !*aBuf)
			// Active codepage, equivalent to specifying CP0.
			return CP_ACP;
		if (!_tcsicmp(aBuf, _T("UTF-8")))		return CP_UTF8;
		if (!_tcsicmp(aBuf, _T("UTF-8-RAW")))	return CP_UTF8 | CP_AHKNOBOM;
		if (!_tcsicmp(aBuf, _T("UTF-16")))		return 1200;
		if (!_tcsicmp(aBuf, _T("UTF-16-RAW")))	return 1200 | CP_AHKNOBOM;
		if (!_tcsnicmp(aBuf, _T("CP"), 2))
			aBuf += 2;
		if (IsNumeric(aBuf, false, false))
		{
			// CPnnn
			UINT cp = ATOU(aBuf);
			// Catch invalid or (not installed) code pages early rather than
			// failing conversion later on.
			if (IsValidFileCodePage(cp))
				return cp;
		}
		return -1;
	}

	static UINT ConvertFileEncoding(ExprTokenType &aToken);

	static ResultType ValidateMouseCoords(LPTSTR aX, LPTSTR aY)
	{
		// OK: Both are absent, which is the signal to use the current position.
		// OK: Both are present (that they are numeric is validated elsewhere).
		// FAIL: One is absent but the other is present.
		return (!*aX && !*aY) || (*aX && *aY) ? OK : FAIL;
	}

	static LPTSTR LogToText(LPTSTR aBuf, int aBufSize);
	LPTSTR VicinityToText(LPTSTR aBuf, int aBufSize);
	LPTSTR ToText(LPTSTR aBuf, int aBufSize, bool aCRLF, DWORD aElapsed = 0, bool aLineWasResumed = false);

	static void ToggleSuspendState();
	static void PauseUnderlyingThread(bool aTrueForPauseFalseForUnpause);
	ResultType ChangePauseState(LPTSTR aChangeTo);
	ResultType PauseCurrentThread();
	static ResultType ScriptBlockInput(bool aEnable);

	Line *PreparseError(LPTSTR aErrorText, LPTSTR aExtraInfo = _T(""));
	// Call this LineError to avoid confusion with Script's error-displaying functions:
	ResultType LineError(LPCTSTR aErrorText, ResultType aErrorType = FAIL, LPCTSTR aExtraInfo = _T(""));
	IObject *CreateRuntimeException(LPCTSTR aErrorText, LPCTSTR aWhat = NULL, LPCTSTR aExtraInfo = _T(""));
	ResultType ThrowRuntimeException(LPCTSTR aErrorText, LPCTSTR aWhat = NULL, LPCTSTR aExtraInfo = _T(""));
	
	ResultType VarIsReadOnlyError(Var *aVar, int aErrorType);

	ResultType SetLastErrorMaybeThrow(bool aError, DWORD aLastError = GetLastError());
	ResultType Throw() { return ThrowIfTrue(true); }
	ResultType ThrowIfTrue(bool aError);
	ResultType ThrowIntIfNonzero(int aErrorValue);

	Line(FileIndexType aFileIndex, LineNumberType aFileLineNumber, ActionTypeType aActionType
		, ArgStruct aArg[], ArgCountType aArgc) // Constructor
		: mFileIndex(aFileIndex), mLineNumber(aFileLineNumber), mActionType(aActionType)
		, mAttribute(ATTR_NONE), mArgc(aArgc), mArg(aArg)
		, mPrevLine(NULL), mNextLine(NULL), mRelatedLine(NULL), mParentLine(NULL)
		, mModule(g_CurrentModule)
#ifdef CONFIG_DEBUGGER
		, mBreakpoint(NULL)
#endif
		{}
	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}  // Intentionally does nothing because we're using SimpleHeap for everything.
	void operator delete[](void *aPtr) {}

	// AutoIt3 functions:
	static bool Util_CopyDir(LPCTSTR szInputSource, LPCTSTR szInputDest, int OverwriteMode, bool bMove);
	static bool Util_RemoveDir(LPCTSTR szInputSource, bool bRecurse);
	static int Util_CopyFile(LPCTSTR szInputSource, LPCTSTR szInputDest, bool bOverwrite, bool bMove, DWORD &aLastError);
	static void Util_ExpandFilenameWildcard(LPCTSTR szSource, LPCTSTR szDest, LPTSTR szExpandedDest);
	static void Util_ExpandFilenameWildcardPart(LPCTSTR szSource, LPCTSTR szDest, LPTSTR szExpandedDest);
	static bool Util_DoesFileExist(LPCTSTR szFilename);
	static bool Util_IsDir(LPCTSTR szPath);
	static void Util_GetFullPathName(LPCTSTR szIn, LPTSTR szOut);
	static void Util_GetFullPathName(LPCTSTR szIn, LPTSTR szOut, DWORD aBufSize);
};



class Label
{
public:
	LPTSTR mName;
	Line *mJumpToLine;
	Label *mPrevLabel, *mNextLabel;  // Prev & Next items in linked list.

	Label(LPTSTR aLabelName)
		: mName(aLabelName) // Caller gave us a pointer to dynamic memory for this (or an empty string in the case of mPlaceholderLabel).
		, mJumpToLine(NULL)
		, mPrevLabel(NULL), mNextLabel(NULL)
	{}
	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};



// This class encapsulates a pointer to an object which can be called by a timer,
// hotkey, etc.  It provides common functionality that wouldn't be suitable for the
// base IObject interface, but is needed for detection of "Suspend" or "Critical"
// prior to calling the sub or function.
class LabelPtr
{
protected:
	IObject *mObject;
	
public:
	LabelPtr() : mObject(NULL) {}
	LabelPtr(IObject *object) : mObject(object) {}
	ResultType ExecuteInNewThread(TCHAR *aNewThreadDesc
		, ExprTokenType *aParamValue = NULL, int aParamCount = 0, __int64 *aRetVal = NULL) const;
	const LabelPtr* operator-> () { return this; } // Act like a pointer.
	operator void *() const { return mObject; } // For comparisons and boolean eval.

	// Caller beware: does not check for NULL.
	Func *ToFunc() const;
	IObject *ToObject() const { return mObject; }
	
	// Currently only used by ListLines, listing active timers.
	LPCTSTR Name() const;
};

// LabelPtr with automatic reference-counting, for storing an object safely,
// such as in a HotkeyVariant, UserMenuItem, etc.  Its specific purpose is to
// work with old code that wasn't concerned with reference counting.
class LabelRef : public LabelPtr
{
private:
	LabelRef(const LabelRef &); // Disable default copy constructor.
	LabelRef & operator = (const LabelRef &); // ...and copy assignment.

public:
	LabelRef() : LabelPtr() {}
	LabelRef(IObject *object) : LabelPtr(object)
	{
		if (object)
			object->AddRef();
	}
	LabelRef(const LabelPtr &other) : LabelPtr(other)
	{
		if (mObject)
			mObject->AddRef();
	}
	LabelRef & operator = (IObject *object)
	{
		if (object)
			object->AddRef();
		if (mObject)
			mObject->Release();
		mObject = object;
		return *this;
	}
	LabelRef & operator = (const LabelPtr &other)
	{
		return *this = other.ToObject();
	}
	~LabelRef()
	{
		if (mObject)
			mObject->Release();
	}
};



enum FuncParamDefaults {PARAM_DEFAULT_NONE, PARAM_DEFAULT_STR, PARAM_DEFAULT_INT, PARAM_DEFAULT_FLOAT, PARAM_DEFAULT_UNSET};
struct FuncParam
{
	Var *var;
	WORD is_byref; // Boolean, but defined as WORD in case it helps data alignment and/or performance (BOOL vs. WORD didn't help benchmarks).
	WORD default_type;
	union {LPTSTR default_str; __int64 default_int64; double default_double;};
};

struct FuncResult : public ResultToken
{
	TCHAR mRetValBuf[MAX_NUMBER_SIZE];

	FuncResult()
	{
		InitResult(mRetValBuf);
	}
};


class Func;
struct FuncList
{
	Func **mItem;
	int mCount, mCountMax;

	Func *Find(LPCTSTR aName, int *apInsertPos);
	ResultType Insert(Func *aFunc, int aInsertPos);
	ResultType Alloc(int aAllocCount);
	FuncList() : mItem(NULL), mCount(0), mCountMax(0) {}
};


struct FreeVars
{
	int mRefCount, mVarCount;
	Var *mVar;
	UserFunc *mFunc;
	FreeVars *mOuterVars;

	void AddRef()
	{
		++mRefCount;
	}

	void Release()
	{
		if (mRefCount == 1)
			delete this;
		else
			--mRefCount;
	}

	FreeVars *ForFunc(UserFunc *aFunc)
	{
		FreeVars *fv = this;
		do
		{
			if (fv->mFunc == aFunc)
				return fv;
		}
		while (fv = fv->mOuterVars);
		return NULL;
	}

	static FreeVars *Alloc(UserFunc &aFunc, int aVarCount, FreeVars *aOuterVars)
	{
		Var *v = aVarCount ? ::new Var[aVarCount] : NULL;
		return new FreeVars(v, aFunc, aVarCount, aOuterVars); // Must use :: to avoid SimpleHeap.
	}

private:
	FreeVars(Var *aVars, UserFunc &aFunc, int aVarCount, FreeVars *aOuterVars)
		: mVar(aVars), mVarCount(aVarCount), mRefCount(1)
		, mFunc(&aFunc), mOuterVars(aOuterVars)
	{
		if (aOuterVars)
			aOuterVars->AddRef();
	}

	~FreeVars()
	{
		if (mOuterVars)
			mOuterVars->Release();
		for (int i = 0; i < mVarCount; ++i)
			mVar[i].Free(VAR_ALWAYS_FREE, true); // Pass "true" to exclude aliases, since their targets should not be freed (they don't belong to this function).
		::delete[] mVar; // Must use :: to avoid SimpleHeap.
	}
};


struct UDFCallInfo
{
	UserFunc *func;
	VarBkp *backup = nullptr; // Backup of previous instance's local vars.  NULL if no previous instance or no vars.
	int backup_count = 0; // Number of previous instance's local vars.  0 if no previous instance or no vars.
	UDFCallInfo(UserFunc *f) : func(f) {}
};


typedef BIF_DECL((* BuiltInFunctionType));


class DECLSPEC_NOVTABLE Func : public Object
{
public:
	LPCTSTR mName;
	int mParamCount = 0; // The function's maximum number of parameters.  For UDFs, also the number of items in the mParam array.
	int mMinParams = 0;  // The number of mandatory parameters (populated for both UDFs and built-in's).
	bool mIsVariadic = false; // Whether to allow mParamCount to be exceeded.

	virtual bool IsBuiltIn() = 0; // FIXME: Should not need to rely on this.
	virtual bool ArgIsOutputVar(int aArg) = 0;

	// bool result indicates whether aResultToken contains a value (i.e. false for FAIL/EARLY_EXIT).
	bool Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, bool aIsVariadic = false);
	virtual bool Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, IObject *aParamObj);
	
	virtual IObject *CloseIfNeeded()
	{
		AddRef();
		return this;
	}

	enum MemberID
	{
		M_Call,
		M_Bind,
		M_IsOptional,
		M_IsByRef,

		P_Name,
		P_MinParams,
		P_MaxParams,
		P_IsBuiltIn,
		P_IsVariadic
	};
	static ObjectMember sMembers[];
	ResultType Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);

	static Object *sPrototype;
	ResultType Invoke(IObject_Invoke_PARAMS_DECL);

	Func(LPCTSTR aFuncName)
		: mName(aFuncName) // Caller gave us a pointer to dynamic memory for this.
	{
		SetBase(sPrototype);
	}
};


class UserFunc : public Func
{
public:
	int mInstances = 0; // How many instances currently exist on the call stack (due to recursion or thread interruption).  Future use: Might be used to limit how deep recursion can go to help prevent stack overflow.
	Line *mJumpToLine = nullptr;
	FuncParam *mParam = nullptr; // Holds an array of FuncParams (array length: mParamCount).
	Object *mClass = nullptr; // The class or prototype object which this user-defined method was defined for, or nullptr.
	Label *mFirstLabel = nullptr, *mLastLabel = nullptr; // Linked list of private labels.
	UserFunc *mOuterFunc = nullptr; // Func which contains this func (usually nullptr).
	FuncList mFuncs {}; // List of nested functions (usually empty).
	Var **mVar = nullptr, **mLazyVar = nullptr; // Array of pointers-to-variable, allocated upon first use and later expanded as needed.
	Var **mGlobalVar = nullptr; // Array of global declarations.
	Var **mDownVar = nullptr, **mUpVar = nullptr;
	int *mUpVarIndex = nullptr;
	static FreeVars *sFreeVars;
#define MAX_FUNC_UP_VARS 1000
	int mVarCount = 0, mVarCountMax = 0, mLazyVarCount = 0, mGlobalVarCount = 0; // Count of items in the above array as well as the maximum capacity.
	int mDownVarCount = 0, mUpVarCount = 0;

	// Keep small members adjacent to each other to save space and improve perf. due to byte alignment:
	bool mIsFuncExpression; // Whether this function was defined *within* an expression and is therefore allowed under a control flow statement.
#define VAR_DECLARE_GLOBAL (VAR_DECLARED | VAR_GLOBAL)
#define VAR_DECLARE_SUPER_GLOBAL (VAR_DECLARE_GLOBAL | VAR_SUPER_GLOBAL)
#define VAR_DECLARE_LOCAL  (VAR_DECLARED | VAR_LOCAL)
#define VAR_DECLARE_STATIC (VAR_DECLARED | VAR_LOCAL | VAR_LOCAL_STATIC)
	// The last two may be combined (bitwise-OR) with VAR_FORCE_LOCAL.
	UCHAR mDefaultVarType = VAR_DECLARE_LOCAL;

	UserFunc(LPCTSTR aName) : Func(aName) {}

	bool IsBuiltIn() override { return false; }

	bool ArgIsOutputVar(int aArg) override
	{
		return aArg < mParamCount && mParam[aArg].is_byref;
	}

	bool AllowSuperGlobals()
	{
		// A function allows super-globals unless it is force-local or contained by another
		// function which is force-local (i.e. a nested function should inherit the rules and
		// declarations of the scope which encloses it).
		if (mDefaultVarType & VAR_FORCE_LOCAL)
			return false;
		return mOuterFunc ? mOuterFunc->AllowSuperGlobals() : true;
	}

	IObject *CloseIfNeeded() override; // Returns this UserFunc or (if mUpVarCount != 0) a Closure.

	bool Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, IObject *aParamObj) override;
	bool Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, IObject *aParamObj, FreeVars *aUpVars);

	// Execute the body of the function.
	ResultType Execute(ResultToken *aResultToken)
	{
		// The performance gain of conditionally passing NULL in place of result (when this is the
		// outermost function call of a line consisting only of function calls, namely ACT_EXPRESSION)
		// would not be significant because the Return command's expression (arg1) must still be evaluated
		// in case it calls any functions that have side-effects, e.g. "return LogThisError()".
		auto prev_func = g->CurrentFunc; // This will be non-NULL when a function is called from inside another function.
		g->CurrentFunc = this;
		// Although a GOTO that jumps to a position outside of the function's body could be supported,
		// it seems best not to for these reasons:
		// 1) The extreme rarity of a legitimate desire to intentionally do so.
		// 2) The fact that any return encountered after the Goto cannot provide a return value for
		//    the function because load-time validation checks for this (it's preferable not to
		//    give up this check, since it is an informative error message and might also help catch
		//    bugs in the script).
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

		ResultType result;
		result = mJumpToLine->ExecUntil(UNTIL_BLOCK_END, aResultToken);
#ifdef CONFIG_DEBUGGER
		if (g_Debugger.IsConnected())
		{
			if (result == EARLY_RETURN)
			{
				// Find the end of this function.
				//Line *line;
				//for (line = mJumpToLine; line && (line->mActionType != ACT_BLOCK_END || !line->mAttribute); line = line->mNextLine);
				// Since mJumpToLine points at the first line *inside* the body, mJumpToLine->mParentLine
				// is the block-begin.  That line's mRelatedLine is the line *after* the block-end, so
				// use it's mPrevLine.  mRelatedLine is guaranteed to be non-NULL by load-time logic.
				Line *line = mJumpToLine->mParentLine->mRelatedLine->mPrevLine;
				// Give user the opportunity to inspect variables before returning.
				if (line)
					g_Debugger.PreExecLine(line);
			}
		}
#endif

		// Restore the original value in case this function is called from inside another function.
		// Due to the synchronous nature of recursion and recursion-collapse, this should keep
		// g->CurrentFunc accurate, even amidst the asynchronous saving and restoring of "g" itself:
		g->CurrentFunc = prev_func;
		return result;
	}

	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};


class Closure : public Func
{
	UserFunc *mFunc;
	FreeVars *mVars;

	static Object *sPrototype;

public:
	Closure(UserFunc *aFunc, FreeVars *aVars)
		: mFunc(aFunc), mVars(aVars), Func(aFunc->mName)
	{
		mMinParams = aFunc->mMinParams;
		mParamCount = aFunc->mParamCount;
		mIsVariadic = aFunc->mIsVariadic;
		SetBase(sPrototype);
	}
	~Closure();

	bool IsBuiltIn() override { return false; }
	bool ArgIsOutputVar(int aArg) override { return mFunc->ArgIsOutputVar(aArg); }
	bool Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, IObject *aParamObj) override;
};


class BoundFunc : public Func
{
	IObject *mFunc;
	LPTSTR mMember;
	Array *mParams;
	int mFlags;

	BoundFunc(IObject *aFunc, LPTSTR aMember, Array *aParams, int aFlags)
		: mFunc(aFunc), mMember(aMember), mParams(aParams), mFlags(aFlags)
		, Func(_T(""))
	{
		mIsVariadic = true;
		SetBase(sPrototype);
	}

	static Object *sPrototype;

public:
	static BoundFunc *Bind(IObject *aFunc, int aFlags, LPCTSTR aMember, ExprTokenType **aParam, int aParamCount);
	~BoundFunc();
	
	bool IsBuiltIn() override { return true; }
	bool ArgIsOutputVar(int aArg) override { return false; }
	bool Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, IObject *aParamObj) override;
};


class DECLSPEC_NOVTABLE NativeFunc : public Func
{
protected:
	NativeFunc(LPCTSTR aName) : Func(aName) {}

public:
	UCHAR *mOutputVars = nullptr; // String of indices indicating which params are output vars (for BIF_PerformAction).

	bool IsBuiltIn() override { return true; }

	bool Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, IObject *aParamObj) override;

	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};


struct FuncEntry;
class BuiltInFunc : public NativeFunc
{
public:
	BuiltInFunctionType mBIF;
	BuiltInFunctionID mFID; // For code sharing: this function's ID in the group of functions which share the same C++ function.
	
	BuiltInFunc(LPCTSTR aName) : NativeFunc(aName) {}
	BuiltInFunc(FuncEntry &, UCHAR *aOutputVars);

#define MAX_FUNC_OUTPUT_VAR 7
	bool ArgIsOutputVar(int aIndex) override
	{
		if (!mOutputVars)
			return false;
		++aIndex; // Convert to one-based.
		for (int i = 0; i < MAX_FUNC_OUTPUT_VAR && mOutputVars[i]; ++i)
			if (mOutputVars[i] == aIndex)
				return true;
		return false;
	}

	bool Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, IObject *aParamObj) override;
};


class BuiltInMethod : public NativeFunc
{
public:
	ObjectMethod mBIM;
	Object *mClass; // The class or prototype object which this method was defined for, and which `this` must derive from.
	UCHAR mMID;
	UCHAR mMIT;

	BuiltInMethod(LPTSTR aName) : NativeFunc(aName) {}

	bool ArgIsOutputVar(int aArg) override { return false; }

	bool Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, IObject *aParamObj) override;
};


class ExprOpFunc : public BuiltInFunc
{	// ExprOpFunc: Used in combination with SYM_FUNC to implement certain operations in expressions.
	// These are not inserted into the script's function list, so mName is used only by the debugger.
public:
	ExprOpFunc(BuiltInFunctionType aBIF, int aID) : BuiltInFunc(_T("<object>"))
	{
		mBIF = aBIF;
		mFID = (BuiltInFunctionID)aID;
		// Allow any number of parameters, since these functions aren't called directly by users
		// and might break the rules in some cases, such as Op_ObjGetInPlace() having 0 *visible*
		// parameters but actually reading 2 which are then also passed to the next function call.
		mParamCount = 255;
		mIsVariadic = true;
	}
};

template<BuiltInFunctionType bif, int flags>
struct ExprOpT
{
	static ExprOpFunc Func;
};

template<BuiltInFunctionType bif, int flags>
ExprOpFunc ExprOpT<bif, flags>::Func(bif, flags);

// ExprOp<bif, flags>() returns a Func* which calls bif with the given ID/flags.
// The call is optimized out and replaced with a reference to a static variable,
// which should be unique for that combination of bif and flags.
template<BuiltInFunctionType bif, int flags>
inline BuiltInFunc *ExprOp()
{
	// Using static ExprOpFunc directly increased code size considerably.
	return &ExprOpT<bif, flags>::Func;
}


struct FuncEntry
{
	LPCTSTR mName;
	BuiltInFunctionType mBIF;
	UCHAR mMinParams, mMaxParams;
	UCHAR mID;
	UCHAR mOutputVars[MAX_FUNC_OUTPUT_VAR];
};


class DECLSPEC_NOVTABLE EnumBase : public Func
{
public:
	static Object *sPrototype;
	EnumBase() : Func(_T(""))
	{
		mParamCount = 2;
		SetBase(sPrototype);
	}
	bool IsBuiltIn() override { return true; };
	bool ArgIsOutputVar(int aArg) override { return true; }
	bool Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, IObject *aParamObj) override;
	virtual ResultType Next(Var *, Var *) = 0;
};


class IndexEnumerator : public EnumBase
{
public:
	typedef ResultType (Object::* Callback)(UINT &aIndex, Var *aOutputVar1, Var *aOutputVar2);
private:
	Object *mObject;
	UINT mIndex = UINT_MAX;
	Callback mGetItem;
public:
	IndexEnumerator(Object *aObject, Callback aGetItem) : mObject(aObject), mGetItem(aGetItem)
	{
		mObject->AddRef();
		mParamCount = 2;
		SetBase(EnumBase::sPrototype);
	}
	~IndexEnumerator()
	{
		mObject->Release();
	}
	ResultType Next(Var *, Var *) override;
};



class ScriptTimer
{
public:
	LabelRef mCallback;
	DWORD mPeriod; // v1.0.36.33: Changed from int to DWORD to double its capacity.
	DWORD mTimeLastRun;  // TickCount
	int mPriority;  // Thread priority relative to other threads, default 0.
	UCHAR mExistingThreads;  // Whether this timer is already running its subroutine.
	bool mEnabled;
	bool mRunOnlyOnce;
	ScriptTimer *mNextTimer;  // Next items in linked list
	void ScriptTimer::Disable();
	ScriptTimer(IObject *aLabel)
		#define DEFAULT_TIMER_PERIOD 250
		: mCallback(aLabel), mPeriod(DEFAULT_TIMER_PERIOD), mPriority(0) // Default is always 0.
		, mExistingThreads(0), mTimeLastRun(0)
		, mEnabled(false), mRunOnlyOnce(false), mNextTimer(NULL)  // Note that mEnabled must default to false for the counts to be right.
	{}
};



struct MsgMonitorStruct
{
	union
	{
		IObject *func;
		LPTSTR method_name; // Used only by GUI.
		LPVOID union_value; // Internal use.
	};
	UINT msg;
	// Keep any members smaller than 4 bytes adjacent to save memory:
	static const UCHAR MAX_INSTANCES = MAX_THREADS_LIMIT; // For maintainability.  Causes a compiler warning if MAX_THREADS_LIMIT > MAX_UCHAR.
	UCHAR instance_count; // Distinct from func.mInstances because the script might have called the function explicitly.
	UCHAR max_instances; // v1.0.47: Support more than one thread.
	UCHAR msg_type; // Used only by GUI, so may be ignored by some methods.
	bool is_method; // Used only by GUI.
};


struct MsgMonitorInstance;
class MsgMonitorList
{
	MsgMonitorStruct *mMonitor;
	MsgMonitorInstance *mTop;
	int mCount, mCountMax;

	friend struct MsgMonitorInstance;

	MsgMonitorStruct *AddInternal(UINT aMsg, bool aAppend);

public:
	MsgMonitorStruct *Find(UINT aMsg, IObject *aCallback, UCHAR aMsgType = 0);
	MsgMonitorStruct *Find(UINT aMsg, LPTSTR aMethodName, UCHAR aMsgType = 0);
	MsgMonitorStruct *Add(UINT aMsg, IObject *aCallback, bool aAppend = TRUE);
	MsgMonitorStruct *Add(UINT aMsg, LPTSTR aMethodName, bool aAppend = TRUE);
	void Delete(MsgMonitorStruct *aMonitor);
	ResultType Call(ExprTokenType *aParamValue, int aParamCount, int aInitNewThreadIndex, __int64 *aRetVal = nullptr); // Used for OnExit and OnClipboardChange, but not OnMessage.
	ResultType Call(ExprTokenType *aParamValue, int aParamCount, UINT aMsg, UCHAR aMsgType, GuiType *aGui, INT_PTR *aRetVal = NULL); // Used by GUI.

	MsgMonitorStruct& operator[] (const int aIndex) { return mMonitor[aIndex]; }
	int Count() { return mCount; }
	BOOL IsMonitoring(UINT aMsg, UCHAR aMsgType = 0);
	BOOL IsRunning(UINT aMsg, UCHAR aMsgType = 0);

	MsgMonitorList() : mCount(0), mCountMax(0), mMonitor(NULL), mTop(NULL) {}
	void Dispose();
};


struct MsgMonitorInstance
{
	MsgMonitorList &list;
	MsgMonitorInstance *previous;
	int index;
	int count;
	bool deleted;

	MsgMonitorInstance(MsgMonitorList &aList)
		: list(aList), previous(aList.mTop), index(0), count(aList.mCount)
		, deleted(false)
	{
		aList.mTop = this;
	}

	~MsgMonitorInstance()
	{
		list.mTop = previous;
	}

	void Delete(int mon_index)
	{
		if (index >= mon_index && index >= 0)
		{
			if (index == mon_index)
				deleted = true; // Callers who care about this will reset it after each iteration.
			index--; // So index+1 is the next item.
		}
		count--;
	}
	
	void Dispose()
	{
		count = 0; // Prevent further iteration.
		deleted = true; // Mark the current item as deleted, so it won't be accessed.
	}
};



#define MAX_MENU_NAME_LENGTH MAX_PATH // For both menu and menu item names.
class UserMenuItem;  // Forward declaration since classes use each other (i.e. a menu *item* can have a submenu).
class UserMenu : public Object
{
	class Bar;

public:
	UserMenuItem *mFirstMenuItem = nullptr, *mLastMenuItem = nullptr, *mDefault = nullptr;
	UserMenu *mNextMenu = nullptr;  // Next item in linked list
	// Keep any fields that aren't an even multiple of 4 adjacent to each other.  This conserves memory
	// due to byte-alignment:
	int mClickCount = 2; // How many clicks it takes to trigger the default menu item.  2 = double-click
	UINT mMenuItemCount = 0; // The count of user-defined menu items (doesn't include the standard items, if present).
	MenuTypeType mMenuType; // MENU_TYPE_POPUP (via CreatePopupMenu) vs. MENU_TYPE_BAR (via CreateMenu).
	HMENU mMenu = NULL;
	HBRUSH mBrush = NULL; // Background color to apply to menu.
	COLORREF mColor = CLR_DEFAULT; // The color that corresponds to the above.

	static ObjectMember sMembers[];
	static int sMemberCount;
	static Object *sPrototype, *sBarPrototype, *sClass, *sBarClass;

	// Don't overload new and delete operators in this case since we want to use real dynamic memory
	// (since menus can be read in from a file, destroyed and recreated, over and over).

	UserMenu(MenuTypeType aMenuType);
	static UserMenu *Create() { return new UserMenu(MENU_TYPE_POPUP); }
	void Dispose();
	~UserMenu();

	enum MemberID
	{
		INVALID = 0,

		// Methods
		M_Add,
		M_AddStandard,
		M_Insert,
		M_Delete,
		M_Rename,
		M_Check,
		M_Uncheck,
		M_ToggleCheck,
		M_Enable,
		M_Disable,
		M_ToggleEnable,
		M_SetIcon,
		M_Show,
		M_SetColor,
		LastMethodPlusOne,

		// Properties
		P_Default,
		P_Handle,
		P_ClickCount,
	};
	
	ResultType Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);

	ResultType AddItem(LPTSTR aName, UINT aMenuID, IObject *aCallback, UserMenu *aSubmenu, LPTSTR aOptions, UserMenuItem **aInsertAt);
	ResultType InternalAppendMenu(UserMenuItem *aMenuItem, UserMenuItem *aInsertBefore = NULL);
	ResultType DeleteItem(UserMenuItem *aMenuItem, UserMenuItem *aMenuItemPrev, bool aUpdateGuiMenuBars = true);
	ResultType DeleteAllItems();
	ResultType ModifyItem(UserMenuItem *aMenuItem, IObject *aCallback, UserMenu *aSubmenu, LPTSTR aOptions);
	ResultType UpdateOptions(UserMenuItem *aMenuItem, LPTSTR aOptions);
	ResultType RenameItem(UserMenuItem *aMenuItem, LPTSTR aNewName);
	ResultType UpdateName(UserMenuItem *aMenuItem, LPTSTR aNewName);
	ResultType SetItemState(UserMenuItem *aMenuItem, UINT aState, UINT aStateMask);
	ResultType CheckItem(UserMenuItem *aMenuItem);
	ResultType UncheckItem(UserMenuItem *aMenuItem);
	ResultType ToggleCheckItem(UserMenuItem *aMenuItem);
	ResultType EnableItem(UserMenuItem *aMenuItem);
	ResultType DisableItem(UserMenuItem *aMenuItem);
	ResultType ToggleEnableItem(UserMenuItem *aMenuItem);
	ResultType SetDefault(UserMenuItem *aMenuItem = NULL, bool aUpdateGuiMenuBars = true);
	ResultType CreateHandle();
	void DestroyHandle();
	void SetColor(ExprTokenType &aColor, bool aApplyToSubmenus);
	void ApplyColor(bool aApplyToSubmenus);
	ResultType AppendStandardItems();
	ResultType EnableStandardOpenItem(bool aEnable);
	ResultType Display(bool aForceToForeground = true, int aX = COORD_UNSPECIFIED, int aY = COORD_UNSPECIFIED);
	UserMenuItem *FindItem(LPTSTR aNameOrPos, UserMenuItem *&aPrevItem, bool &aByPos);
	UserMenuItem *FindItemByID(UINT aID);
	bool ContainsMenu(UserMenu *aMenu);
	bool ContainsCustomItems();
	void UpdateAccelerators();
	// L17: Functions for menu icons.
	ResultType SetItemIcon(UserMenuItem *aMenuItem, LPTSTR aFilename, int aIconNumber, int aWidth);
	ResultType ApplyItemIcon(UserMenuItem *aMenuItem);
	ResultType RemoveItemIcon(UserMenuItem *aMenuItem);
};

class UserMenu::Bar : public UserMenu
{
	Bar(const Bar &) = delete; // Never instantiated.

public:
	static UserMenu *Create()
	{
		return new UserMenu(MENU_TYPE_BAR);
	}
};



class UserMenuItem
{
public:
	LPTSTR mName;  // Dynamically allocated.
	size_t mNameCapacity;
	LabelRef mCallback;
	UserMenu *mSubmenu;
	UserMenu *mMenu;  // The menu to which this item belongs.  Needed to support script var A_ThisMenu.
	UINT mMenuID;
	int mPriority;
	// Keep any fields that aren't an even multiple of 4 adjacent to each other.  This conserves memory
	// due to byte-alignment:
	WORD mMenuState;
	WORD mMenuType;
	UserMenuItem *mNextMenuItem;  // Next item in linked list
	
	// For menu item icons:
	// Windows Vista and later support alpha channels via 32-bit bitmaps. Since owner-drawing prevents
	// visual styles being applied to menus, we convert each icon to a 32-bit bitmap, calculating the
	// alpha channel as necessary. This is done only once, when the icon is initially set.
	HBITMAP mBitmap;

	UserMenuItem(LPTSTR aName, size_t aNameCapacity, UINT aMenuID, IObject *aCallback, UserMenu *aSubmenu, UserMenu *aMenu);
	~UserMenuItem();

	// Don't overload new and delete operators in this case since we want to use real dynamic memory
	// (since menus can be read in from a file, destroyed and recreated, over and over).

	UINT Pos();
};



struct FontType : public LOGFONT
{
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
struct GuiControlType : public Object
{
	GuiType* gui;
	HWND hwnd = NULL;
	LPTSTR name = nullptr;
	MsgMonitorList events;
	// Keep any fields that are smaller than 4 bytes adjacent to each other.  This conserves memory
	// due to byte-alignment.  It has been verified to save 4 bytes per struct in this case:
	GuiControls type = GUI_CONTROL_INVALID;
	// Unused: 0x01
	#define GUI_CONTROL_ATTRIB_ALTSUBMIT           0x02
	// Unused: 0x04
	#define GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN   0x08
	#define GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED 0x10
	#define GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS     0x20
	// Unused: 0x40
	#define GUI_CONTROL_ATTRIB_ALTBEHAVIOR         0x80 // Slider +Invert, ListView/TreeView +WantF2, Edit +WantTab
	UCHAR attrib = 0; // A field of option flags/bits defined above.
	TabControlIndexType tab_control_index = 0; // Which tab control this control belongs to, if any.
	TabIndexType tab_index = 0; // For type==TAB, this stores the tab control's index.  For other types, it stores the page.
	#define CLR_TRANSPARENT 0xFF000001L
	#define IS_AN_ACTUAL_COLOR(color) !((color) & ~0xffffff) // Produces smaller code than checking against CLR_DEFAULT || CLR_INVALID.
	COLORREF background_color = CLR_INVALID;
	HBRUSH background_brush = NULL;
	union
	{
		COLORREF union_color;  // Color of the control's text.
		HBITMAP union_hbitmap = NULL; // For PIC controls, stores the bitmap.
		// Note: Pic controls cannot obey the text color, but they can obey the window's background
		// color if the picture's background is transparent (at least in the case of icons on XP).
		lv_attrib_type *union_lv_attrib; // For ListView: Some attributes and an array of columns.
		IObject *union_object; // For ActiveX.
	};

	GuiControlType() = delete;
	GuiControlType(GuiType* owner) : gui(owner) {}

	static LPTSTR sTypeNames[];
	static GuiControls ConvertTypeName(LPTSTR aTypeName);
	LPTSTR GetTypeName();

	// An array of these attributes is used in place of several long switch() statements,
	// to reduce code size and possibly improve performance:
	enum TypeAttribType
	{
		TYPE_SUPPORTS_BGTRANS = 0x01, // Supports +BackgroundTrans.
		TYPE_SUPPORTS_BGCOLOR = 0x02, // Supports +Background followed by a color value.
		TYPE_REQUIRES_BGBRUSH = 0x04, // Requires a brush to be created to implement background color.
		TYPE_MSGBKCOLOR = TYPE_SUPPORTS_BGCOLOR | TYPE_REQUIRES_BGBRUSH, // Supports background color by responding to WM_CTLCOLOR, WM_ERASEBKGND or WM_DRAWITEM.
		TYPE_SETBKCOLOR = TYPE_SUPPORTS_BGCOLOR, // Supports setting a background color by sending it a message.
		TYPE_NO_SUBMIT = 0x08, // Doesn't accept user input, or is excluded from Submit() for some other reason.
		TYPE_HAS_NO_TEXT = 0x10, // Has no text and therefore doesn't use the font or text color.
		TYPE_RESERVE_UNION = 0x20, // Uses the union for some other purpose, so union_color must not be set.
		TYPE_USES_BGCOLOR = 0x40, // Uses Gui.BackColor.
		TYPE_HAS_ITEMS = 0x80, // Add() accepts an Array of items rather than text.
		TYPE_STATICBACK = TYPE_MSGBKCOLOR | TYPE_USES_BGCOLOR, // For brevity in the attrib array.
	};
	typedef UCHAR TypeAttribs;
	static TypeAttribs TypeHasAttrib(GuiControls aType, TypeAttribs aAttrib);
	TypeAttribs TypeHasAttrib(TypeAttribs aAttrib) { return TypeHasAttrib(type, aAttrib); }

	static UCHAR **sRaisesEvents;
	bool SupportsEvent(GuiEventType aEvent);

	bool SupportsBackgroundTrans()
	{
		return TypeHasAttrib(TYPE_SUPPORTS_BGTRANS);
		//switch (type)
		//{
		// Supported via WM_CTLCOLORSTATIC:
		//case GUI_CONTROL_TEXT:
		//case GUI_CONTROL_PIC:
		//case GUI_CONTROL_GROUPBOX:
		//case GUI_CONTROL_BUTTON:
		//	return true;
		//case GUI_CONTROL_CHECKBOX:     Checkbox and radios with trans background have problems with
		//case GUI_CONTROL_RADIO:        their focus rects being drawn incorrectly.
		//case GUI_CONTROL_LISTBOX:      These are also a problem, at least under some theme settings.
		//case GUI_CONTROL_EDIT:
		//case GUI_CONTROL_DROPDOWNLIST:
		//case GUI_CONTROL_SLIDER:       These are a problem under both classic and non-classic themes.
		//case GUI_CONTROL_COMBOBOX:
		//case GUI_CONTROL_LINK:         BackgroundTrans would have no effect.
		//case GUI_CONTROL_LISTVIEW:     Can't reach this point because WM_CTLCOLORxxx is never received for it.
		//case GUI_CONTROL_TREEVIEW:     Same (verified).
		//case GUI_CONTROL_PROGRESS:     Same (verified).
		//case GUI_CONTROL_UPDOWN:       Same (verified).
		//case GUI_CONTROL_DATETIME:     Same (verified).
		//case GUI_CONTROL_MONTHCAL:     Same (verified).
		//case GUI_CONTROL_HOTKEY:       Same (verified).
		//case GUI_CONTROL_TAB:          Same.
		//case GUI_CONTROL_STATUSBAR:    Its text fields (parts) are its children, not ours, so its window proc probably receives WM_CTLCOLORSTATIC, not ours.
		//default:
		//	return false; // Prohibit the TRANS setting for the above control types.
		//}
	}

	bool SupportsBackgroundColor()
	{
		return TypeHasAttrib(TYPE_SUPPORTS_BGCOLOR);
	}

	bool RequiresBackgroundBrush()
	{
		return TypeHasAttrib(TYPE_REQUIRES_BGBRUSH);
	}

	bool HasSubmittableValue()
	{
		return !TypeHasAttrib(TYPE_NO_SUBMIT);
	}

	bool UsesFontAndTextColor()
	{
		return !TypeHasAttrib(TYPE_HAS_NO_TEXT);
	}

	bool UsesUnionColor()
	{
		// It's easier to exclude those which require the union for some other purpose
		// than to whitelist all controls which could potentially cause a WM_CTLCOLOR
		// message (or WM_ERASEBKGND/WM_DRAWITEM in the case of Tab).
		return !TypeHasAttrib(TYPE_RESERVE_UNION);
	}

	bool UsesGuiBgColor()
	{
		return TypeHasAttrib(TYPE_USES_BGCOLOR);
	}

	enum MemberID
	{
		INVALID = 0,

		// Methods
		M_Opt, // a.k.a. Opt
		M_Focus,
		M_Move,
		M_GetPos,
		M_Choose,
		M_OnEvent,
		M_OnNotify,
		M_OnCommand,
		M_SetFont,
		M_Redraw,
		M_Tab_UseTab,
		M_List_Add,
		M_List_Delete,
		M_DateTime_SetFormat,

		// Properties
		P_Hwnd,
		P_Gui,
		P_Name,
		P_Type,
		P_ClassNN,
		P_Text,
		P_Value,
		P_Enabled,
		P_Visible,
		P_Focused,
	};

	static ObjectMember sMembers[];
	static ObjectMember sMembersList[]; // Tab, ListBox, ComboBox, DDL
	static ObjectMember sMembersTab[];
	static ObjectMember sMembersDate[];
	static ObjectMember sMembersLV[];
	static ObjectMember sMembersTV[];
	static ObjectMember sMembersSB[];

	static Object *GetPrototype(GuiControls aType);

	ResultType StatusBar(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType LV_GetNextOrCount(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType LV_GetText(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType LV_AddInsertModify(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType LV_Delete(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType LV_InsertModifyDeleteCol(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType LV_SetImageList(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType TV_AddModifyDelete(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType TV_GetRelatedItem(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType TV_Get(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType TV_SetImageList(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);

	ResultType Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType Invoke(IObject_Invoke_PARAMS_DECL);

	void Dispose(); // Called by GuiType::Dispose().
};

struct GuiControlOptionsType
{
	DWORD style_add, style_remove, exstyle_add, exstyle_remove, listview_style;
	int listview_view; // Viewing mode, such as LV_VIEW_ICON, LV_VIEW_DETAILS.  Int vs. DWORD to more easily use any negative value as "invalid".
	HIMAGELIST himagelist;
	int x, y, width, height;  // Position info.
	float row_count;
	int choice;  // Which item of a DropDownList/ComboBox/ListBox to initially choose.
	int range_min, range_max;  // Allowable range, such as for a slider control.
	int tick_interval; // The interval at which to draw tickmarks for a slider control.
	int line_size, page_size; // Also for slider.
	int thickness;  // Thickness of slider's thumb.
	int tip_side; // Which side of the control to display the tip on (0 to use default side).
	GuiControlType *buddy1, *buddy2;
	COLORREF color; // Control's text color.
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
	TCHAR password_char; // When zeroed, indicates "use default password" for an edit control with the password style.
	bool range_changed;
	bool tick_interval_changed, tick_interval_specified;
	bool start_new_section;
	bool use_theme; // v1.0.32: Provides the means for the window's current setting of mUseTheme to be overridden.
	bool listview_no_auto_sort; // v1.0.44: More maintainable and frees up GUI_CONTROL_ATTRIB_ALTBEHAVIOR for other uses.
	bool tab_control_uses_dialog;
	#define TAB3_AUTOWIDTH 1
	#define TAB3_AUTOHEIGHT 2
	CHAR tab_control_autosize;
	ATOM customClassAtom;
};

LRESULT CALLBACK GuiWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK TabWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam);

class GuiType : public Object
{
	// The use of 72 and 96 below comes from v1, using the font's point size in the
	// calculation.  It's really just 11.25 * font height in pixels.  Could use 12
	// or 11 * font height, but keeping the same defaults as v1 seems worthwhile.
	#define GUI_STANDARD_WIDTH_MULTIPLIER 15 // This times font size = width, if all other means of determining it are exhausted.
	#define GUI_STANDARD_WIDTH GUI_STANDARD_WIDTH_MULTIPLIER * (MulDiv(sFont[mCurrentFontIndex].lfHeight, -72, 96)) // 96 vs. g_ScreenDPI since lfHeight already accounts for DPI.  Don't combine GUI_STANDARD_WIDTH_MULTIPLIER with -72 as it changes the result (due to rounding).
	// Update for v1.0.21: Reduced it to 8 vs. 9 because 8 causes the height each edit (with the
	// default style) to exactly match that of a Combo or DropDownList.  This type of spacing seems
	// to be what other apps use too, and seems to make edits stand out a little nicer:
	#define GUI_CTL_VERTICAL_DEADSPACE DPIScale(8)
	#define PROGRESS_DEFAULT_THICKNESS MulDiv(sFont[mCurrentFontIndex].lfHeight, -2 * 72, 96) // 96 vs. g_ScreenDPI to preserve DPI scale.

public:
	// Ensure fields of the same size are grouped together to avoid padding between larger types
	// and smaller types.  On last check, this saved 8 bytes per GUI on x64 (where pointers are
	// of course 64-bit, so a sequence like HBRUSH, COLORREF, HBRUSH would cause padding).
	// POINTER-SIZED FIELDS:
	GuiType *mNextGui, *mPrevGui; // For global Gui linked list.
	HWND mHwnd;
	LPTSTR mName;
	HWND mStatusBarHwnd;
	HWND mOwner;  // The window that owns this one, if any.  Note that Windows provides no way to change owners after window creation.
	// Control IDs are higher than their index in the array by +CONTROL_ID_FIRST.  This offset is
	// necessary because windows that behave like dialogs automatically return IDOK and IDCANCEL in
	// response to certain types of standard actions:
	GuiControlType **mControl; // Will become an array of controls when the window is first created.
	HBRUSH mBackgroundBrushWin;   // Brush corresponding to mBackgroundColorWin.
	HDROP mHdrop;                 // Used for drag and drop operations.
	HICON mIconEligibleForDestruction; // The window's icon, which can be destroyed when the window is destroyed if nothing else is using it.
	HICON mIconEligibleForDestructionSmall; // L17: A window may have two icons: ICON_SMALL and ICON_BIG.
	HACCEL mAccel; // Keyboard accelerator table.
	UserMenu *mMenu;
	IObject* mEventSink;
	MsgMonitorList mEvents;
	// 32-BIT FIELDS:
	GuiIndexType mControlCount;
	GuiIndexType mControlCapacity; // How many controls can fit into the current memory size of mControl.
	GuiIndexType mDefaultButtonIndex; // Index vs. pointer is needed for some things.
	DWORD mStyle, mExStyle; // Style of window.
	int mCurrentFontIndex;
	COLORREF mCurrentColor;       // The default color of text in controls.
	COLORREF mBackgroundColorWin; // The window's background color itself.
	int mMarginX, mMarginY, mPrevX, mPrevY, mPrevWidth, mPrevHeight, mMaxExtentRight, mMaxExtentDown
		, mSectionX, mSectionY, mMaxExtentRightSection, mMaxExtentDownSection;
	LONG mMinWidth, mMinHeight, mMaxWidth, mMaxHeight;
	// 8-BIT FIELDS:
	TabControlIndexType mTabControlCount;
	TabControlIndexType mCurrentTabControlIndex; // Which tab control of the window.
	TabIndexType mCurrentTabIndex;// Which tab of a tab control is currently the default for newly added controls.
	bool mInRadioGroup; // Whether the control currently being created is inside a prior radio-group.
	bool mUseTheme;  // Whether XP theme and styles should be applied to the parent window and subsequently added controls.
	bool mGuiShowHasNeverBeenDone, mFirstActivation, mShowIsInProgress, mDestroyWindowHasBeenCalled;
	bool mControlWidthWasSetByContents; // Whether the most recently added control was auto-width'd to fit its contents.
	bool mUsesDPIScaling; // Whether the GUI uses DPI scaling.
	bool mDisposed; // Simplifies Dispose().
	bool mVisibleRefCounted; // Whether AddRef() has been done as a result of the window being shown.

	#define MAX_GUI_FONTS 200  // v1.0.44.14: Increased from 100 to 200 due to feedback that 100 wasn't enough.  But to alleviate memory usage, the array is now allocated upon first use.
	static FontType *sFont; // An array of structs, allocated upon first use.
	static int sFontCount;
	static HWND sTreeWithEditInProgress; // Needed because TreeView's edit control for label-editing conflicts with IDOK (default button).

	// Don't overload new and delete operators in this case since we want to use real dynamic memory
	// (since GUIs can be destroyed and recreated, over and over).

	enum MemberID
	{
		// Methods
		M_Destroy,
		//M_AddControl,
		M_Show,
		M_Hide,
		M_Move,
		M_GetPos,
		M_GetClientPos,
		M_SetFont,
		M_Opt,
		M_Minimize,
		M_Maximize,
		M_Restore,
		M_Flash,
		M_Submit,
		M___Enum,
		M_OnEvent,

		// Properties
		P_Hwnd,
		P_Title,
		P_Name,
		P___Item,
		P_FocusedCtrl,
		P_BackColor,
		P_MarginX,
		P_MarginY,
		P_MenuBar,
	};

	static ObjectMember sMembers[];
	static int sMemberCount;
	static Object *sPrototype, *sClass;

	GuiType() // Constructor
		: mHwnd(NULL), mOwner(NULL), mName(NULL)
		, mPrevGui(NULL), mNextGui(NULL)
		, mControl(NULL), mControlCount(0), mControlCapacity(0)
		, mStatusBarHwnd(NULL)
		, mDefaultButtonIndex(-1), mEventSink(NULL)
		, mMenu(NULL)
		// The styles DS_CENTER and DS_3DLOOK appear to be ineffectual in this case.
		// Also note that WS_CLIPSIBLINGS winds up on the window even if unspecified, which is a strong hint
		// that it should always be used for top level windows across all OSes.  Usenet posts confirm this.
		// Also, it seems safer to have WS_POPUP under a vague feeling that it seems to apply to dialog
		// style windows such as this one, and the fact that it also allows the window's caption to be
		// removed, which implies that POPUP windows are more flexible than OVERLAPPED windows.
		, mStyle(WS_POPUP|WS_CLIPSIBLINGS|WS_CAPTION|WS_SYSMENU|WS_MINIMIZEBOX) // WS_CLIPCHILDREN (doesn't seem helpful currently)
		, mExStyle(0) // This and the above should not be used once the window has been created since they might get out of date.
		, mInRadioGroup(false), mUseTheme(true)
		, mCurrentFontIndex(FindOrCreateFont()) // Must call this in constructor to ensure sFont array is never empty while a GUI object exists.  Omit params to tell it to find or create DEFAULT_GUI_FONT.
		, mTabControlCount(0), mCurrentTabControlIndex(MAX_TAB_CONTROLS), mCurrentTabIndex(0)
		, mCurrentColor(CLR_DEFAULT)
		, mBackgroundColorWin(CLR_DEFAULT), mBackgroundBrushWin(NULL)
		, mHdrop(NULL), mIconEligibleForDestruction(NULL), mIconEligibleForDestructionSmall(NULL)
		, mAccel(NULL)
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
		, mUsesDPIScaling(true)
		, mDisposed(false)
		, mVisibleRefCounted(false)
	{
		SetBase(sPrototype);
	}

	void Destroy();
	void Dispose();
	static void DestroyIconsIfUnused(HICON ahIcon, HICON ahIconSmall); // L17: Renamed function and added parameter to also handle the window's small icon.
	static GuiType *Create() { return new GuiType(); } // For Object::New<GuiType>().
	ResultType __New(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType AddControl(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType Create(LPTSTR aTitle);
	ResultType SetName(LPTSTR aName);
	ResultType NameToEventHandler(LPTSTR aName, IObject *&aObject);
	ResultType OnEvent(GuiControlType *aControl, UINT aEvent, UCHAR aEventKind, ExprTokenType *aParam[], int aParamCount, ResultToken &aResultToken);
	ResultType OnEvent(GuiControlType *aControl, UINT aEvent, UCHAR aEventKind, IObject *aFunc, LPTSTR aMethodName, int aMaxThreads, ResultToken &aResultToken);
	void ApplyEventStyles(GuiControlType *aControl, UINT aEvent, bool aAdded);
	static LPTSTR sEventNames[];
	static LPTSTR ConvertEvent(GuiEventType evt);
	static GuiEventType ConvertEvent(LPTSTR evt);
	// Currently this returns true for all events if we're using an event sink,
	// because checking for the presence of a method in the event sink could be
	// unreliable (but maybe placing some limitations would solve that?).
	BOOL IsMonitoring(GuiEventType aEvent) { return mEvents.IsMonitoring(aEvent); }

	ResultType GetEnumItem(UINT &aIndex, Var *, Var *);

	static IObject* CreateDropArray(HDROP hDrop);
	ResultType SetMenu(ExprTokenType &aParam);
	static void UpdateMenuBars(HMENU aMenu);
	ResultType AddControl(GuiControls aControlType, LPTSTR aOptions, LPTSTR aText, GuiControlType*& apControl, Array *aObj = NULL);
	ResultType MethodGetPos(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, RECT &aPos);

	ResultType ParseOptions(LPTSTR aOptions, bool &aSetLastFoundWindow, ToggleValueType &aOwnDialogs);
	void SetOwnDialogs(ToggleValueType state)
	{
		if (state == TOGGLE_INVALID)
			return;
		g->DialogOwner = state == TOGGLED_ON ? mHwnd : NULL;
	}
	void GetNonClientArea(LONG &aWidth, LONG &aHeight);
	void GetTotalWidthAndHeight(LONG &aWidth, LONG &aHeight);

	ResultType ControlParseOptions(LPTSTR aOptions, GuiControlOptionsType &aOpt, GuiControlType &aControl
		, GuiIndexType aControlIndex = -1); // aControlIndex is not needed upon control creation.
	void ControlInitOptions(GuiControlOptionsType &aOpt, GuiControlType &aControl);
	void ControlAddItems(GuiControlType &aControl, Array *aObj);
	void ControlSetChoice(GuiControlType &aControl, int aChoice);
	ResultType ControlLoadPicture(GuiControlType &aControl, LPTSTR aFilename, int aWidth, int aHeight, int aIconNumber);
	ResultType Show(LPTSTR aOptions);
	void Cancel();
	void Close(); // Due to SC_CLOSE, etc.
	void Escape(); // Similar to close, except typically called when the user presses ESCAPE.
	void VisibilityChanged();
	ResultType Submit(ResultToken &aResultToken, bool aHideIt);

	static GuiType *FindGui(HWND aHwnd);
	static GuiType *FindGuiParent(HWND aHwnd);

	GuiIndexType FindControl(LPTSTR aControlID);
	GuiIndexType FindControlIndex(HWND aHwnd)
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
		if (index < mControlCount && mControl[index]->hwnd == aHwnd) // A match was found.  Fix for v1.1.09.03: Confirm it is actually one of our controls.
			return index;
		else // No match, so indicate failure.
			return NO_CONTROL_INDEX;
	}
	GuiControlType *FindControl(HWND aHwnd)
	{
		GuiIndexType index = FindControlIndex(aHwnd);
		return index == NO_CONTROL_INDEX ? NULL : mControl[index];
	}

	int FindGroup(GuiIndexType aControlIndex, GuiIndexType &aGroupStart, GuiIndexType &aGroupEnd);

	ResultType SetCurrentFont(LPTSTR aOptions, LPTSTR aFontName);
	static int FindOrCreateFont(LPTSTR aOptions = _T(""), LPTSTR aFontName = _T(""), FontType *aFoundationFont = NULL
		, COLORREF *aColor = NULL);
	static int FindFont(FontType &aFont);
	static void FontGetAttributes(FontType &aFont);

	void Event(GuiIndexType aControlIndex, UINT aNotifyCode, USHORT aGuiEvent = GUI_EVENT_NONE, UINT_PTR aEventInfo = 0);
	bool ControlWmNotify(GuiControlType &aControl, LPNMHDR aNmHdr, INT_PTR &aRetVal);

	static WORD TextToHotkey(LPTSTR aText);
	static LPTSTR HotkeyToText(WORD aHotkey, LPTSTR aBuf);
	ResultType ControlSetName(GuiControlType &aControl, LPTSTR aName);
	void ControlSetEnabled(GuiControlType &aControl, bool aEnabled);
	void ControlSetVisible(GuiControlType &aControl, bool aVisible);
	ResultType ParseMoveParams(int aCoord[4], ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount);
	ResultType ControlMove(GuiControlType &aControl, int aX, int aY, int aWidth, int aHeight);
	ResultType ControlSetFont(GuiControlType &aControl, LPTSTR aOptions, LPTSTR aFontName);
	void ControlSetTextColor(GuiControlType &aControl, COLORREF aColor);
	void ControlSetMonthCalColor(GuiControlType &aControl, COLORREF aColor, UINT aMsg);
	ResultType ControlChoose(GuiControlType &aControl, ExprTokenType &aParam, BOOL aOneExact = FALSE);
	void ControlCheckRadioButton(GuiControlType &aControl, GuiIndexType aControlIndex, WPARAM aCheckType);
	void ControlSetUpDownOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt);
	int ControlGetDefaultSliderThickness(DWORD aStyle, int aThumbThickness);
	void ControlSetSliderOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt);
	int ControlInvertSliderIfNeeded(GuiControlType &aControl, int aPosition);
	void ControlSetListViewOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt);
	void ControlSetTreeViewOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt);
	void ControlSetProgressOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt, DWORD aStyle);
	GuiControlType *ControlOverrideBkColor(GuiControlType &aControl);
	void ControlGetBkColor(GuiControlType &aControl, bool aUseWindowColor, HBRUSH &aBrush, COLORREF &aColor);
	
	ResultType ControlSetContents(GuiControlType &aControl, LPTSTR aContents, ResultToken &aResultToken, bool aIsText);
	ResultType ControlSetPic(GuiControlType &aControl, LPTSTR aContents, ResultToken &aResultToken);
	ResultType ControlSetCheck(GuiControlType &aControl, LPTSTR aContents, ResultToken &aResultToken); // CheckBox, Radio
	ResultType ControlSetChoice(GuiControlType &aControl, LPTSTR aContents, ResultToken &aResultToken, bool aIsText); // DDL, ComboBox, ListBox, Tab
	ResultType ControlSetEdit(GuiControlType &aControl, LPTSTR aContents, ResultToken &aResultToken, bool aIsText);
	ResultType ControlSetDateTime(GuiControlType &aControl, LPTSTR aContents, ResultToken &aResultToken);
	ResultType ControlSetDateTimeFormat(GuiControlType &aControl, LPTSTR aFormat, ResultToken &aResultToken);
	ResultType ControlSetMonthCal(GuiControlType &aControl, LPTSTR aContents, ResultToken &aResultToken);
	ResultType ControlSetHotkey(GuiControlType &aControl, LPTSTR aContents, ResultToken &aResultToken);
	ResultType ControlSetUpDown(GuiControlType &aControl, LPTSTR aContents, ResultToken &aResultToken);
	ResultType ControlSetSlider(GuiControlType &aControl, LPTSTR aContents, ResultToken &aResultToken);
	ResultType ControlSetProgress(GuiControlType &aControl, LPTSTR aContents, ResultToken &aResultToken);

	enum ValueModeType { Value_Mode, Text_Mode, Submit_Mode };

	ResultType ControlGetContents(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode = Value_Mode);
	ResultType ControlGetCheck(ResultToken &aResultToken, GuiControlType &aControl); // CheckBox, Radio
	ResultType ControlGetDDL(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode);
	ResultType ControlGetComboBox(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode);
	ResultType ControlGetListBox(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode);
	ResultType ControlGetEdit(ResultToken &aResultToken, GuiControlType &aControl);
	ResultType ControlGetDateTime(ResultToken &aResultToken, GuiControlType &aControl);
	ResultType ControlGetMonthCal(ResultToken &aResultToken, GuiControlType &aControl);
	ResultType ControlGetHotkey(ResultToken &aResultToken, GuiControlType &aControl);
	ResultType ControlGetUpDown(ResultToken &aResultToken, GuiControlType &aControl);
	ResultType ControlGetSlider(ResultToken &aResultToken, GuiControlType &aControl);
	ResultType ControlGetProgress(ResultToken &aResultToken, GuiControlType &aControl);
	ResultType ControlGetTab(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode);
	
	ResultType ControlGetWindowText(ResultToken &aResultToken, GuiControlType &aControl);
	void ControlRedraw(GuiControlType &aControl, bool aOnlyWithinTab = false);

	void ControlUpdateCurrentTab(GuiControlType &aTabControl, bool aFocusFirstControl);
	GuiControlType *FindTabControl(TabControlIndexType aTabControlIndex);
	int FindTabIndexByName(GuiControlType &aTabControl, LPTSTR aName, bool aExactMatch = false);
	int GetControlCountOnTabPage(TabControlIndexType aTabControlIndex, TabIndexType aTabIndex);
	void GetTabDisplayAreaRect(HWND aTabControlHwnd, RECT &aRect);
	POINT GetPositionOfTabDisplayArea(GuiControlType &aTabControl);
	ResultType SelectAdjacentTab(GuiControlType &aTabControl, bool aMoveToRight, bool aFocusFirstControl
		, bool aWrapAround);
	void AutoSizeTabControl(GuiControlType &aTabControl);
	ResultType CreateTabDialog(GuiControlType &aTabControl, GuiControlOptionsType &aOpt);
	void UpdateTabDialog(HWND aTabControlHwnd);
	void ControlGetPosOfFocusedItem(GuiControlType &aControl, POINT &aPoint);
	static void LV_Sort(GuiControlType &aControl, int aColumnIndex, bool aSortOnlyIfEnabled, TCHAR aForceDirection = '\0');
	static IObject *ControlGetActiveX(HWND aWnd);
	
	void UpdateAccelerators(UserMenu &aMenu);
	void UpdateAccelerators(UserMenu &aMenu, LPACCEL aAccel, int &aAccelCount);
	void RemoveAccelerators();
	static bool ConvertAccelerator(LPTSTR aString, ACCEL &aAccel);

	// See DPIScale() and DPIUnscale() for more details.
	int Scale(int x) { return mUsesDPIScaling ? DPIScale(x) : x; }
	int Unscale(int x) { return mUsesDPIScaling ? DPIUnscale(x) : x; }
	// The following is a workaround for the "w-1" and "h-1" options:
	int ScaleSize(int x) { return mUsesDPIScaling && x != -1 ? DPIScale(x) : x; }

protected:
	bool Delete() override;
};

struct DerefList
{
	DerefType *items;
	int size, count;
	DerefList() : items(NULL), size(0), count(0) {}
	~DerefList() { free(items); }
	ResultType Push();
	DerefType *Last() { return items + count - 1; }
};

class Script
{
private:
	friend class Hotkey;
#ifdef CONFIG_DEBUGGER
	friend class Debugger;
#endif

	Line *mFirstLine, *mLastLine;     // The first and last lines in the linked list.
	Label *mFirstLabel, *mLastLabel;  // The first and last labels in the linked list.
	FuncList mFuncs;
	
	UserFunc *mLastHotFunc;		// For hotkey/hotstring functions
	UserFunc *mUnusedHotFunc;	// If defining a named function under a "trigger::" the implicit
								// function stored in mLastHotFunc will not be used, store it in this
								// variable for reuse.
	FuncList mHotFuncs;			// All implicit hotkey funcs are stored here for variable processing.
								// This list is not sorted, all insertions are done at the end.
								// In particular, note that DefineFunc and CreateHotFunc directly
								// changes mCount. This list's member mItem is freed after being
								// passed to PreprocessLocalVars. Do not use this list after that. 

	Var **mVar, **mLazyVar; // Array of pointers-to-variable, allocated upon first use and later expanded as needed.
	int mVarCount, mVarCountMax, mLazyVarCount; // Count of items in the above array as well as the maximum capacity.
	int mGlobalVarCountMax; // While loading the script, the maximum number of global declarations allowed for the current function.
	WinGroup *mFirstGroup, *mLastGroup;  // The first and last variables in the linked list.
	Line *mOpenBlock; // While loading the script, this is the beginning of a block which is currently open.

	ModuleList *mModules; // The script's modules.
	
	bool mNextLineIsFunctionBody; // Whether the very next line to be added will be the first one of the body.
	bool mNoUpdateLabels;

#define MAX_NESTED_CLASSES 5
#define MAX_CLASS_NAME_LENGTH UCHAR_MAX
	int mClassObjectCount;
	Object *mClassObject[MAX_NESTED_CLASSES]; // Class definition currently being parsed.
	TCHAR mClassName[MAX_CLASS_NAME_LENGTH + 1]; // Only used during load-time.
	Object *mUnresolvedClasses;
	Array *mClasses = nullptr;
	Property *mClassProperty;
	LPTSTR mClassPropertyDef;

	// These two track the file number and line number in that file of the line currently being loaded,
	// which simplifies calls to ScriptError() and LineError() (reduces the number of params that must be passed).
	// These are used ONLY while loading the script into memory.  After that (while the script is running),
	// only mCurrLine is kept up-to-date:
	int mCurrFileIndex;
	LineNumberType mCombinedLineNumber; // In the case of a continuation section/line(s), this is always the top line.

	bool mClassPropertyStatic;

	#define UPDATE_TIP_FIELD tcslcpy(mNIC.szTip, mTrayIconTip ? mTrayIconTip \
		: mFileName, _countof(mNIC.szTip));
	NOTIFYICONDATA mNIC; // For ease of adding and deleting our tray icon.

	struct LineBuffer
	{
		TCHAR *p = nullptr;
		size_t length = 0;
		size_t size = 0;
		const size_t EXPANSION_INTERVAL = 0x1000;
		const size_t RESERVED_SPACE = 3; // Allow for a null-terminator and appending "()" for call statements.
		size_t Capacity() { ASSERT(size); return size - RESERVED_SPACE; }
		ResultType Expand();
		ResultType EnsureCapacity(size_t aLength);
		ResultType Realloc(size_t aNewSize);
		~LineBuffer() { free(p); }
		operator LPTSTR() const { return p; }
	};
	size_t GetLine(LineBuffer &aBuf, int aInContinuationSection, bool aInBlockComment, TextStream *ts);
	ResultType GetLineContinuation(TextStream *ts, LineBuffer &aBuf, LineBuffer &aNextBuf
		, LineNumberType &aPhysLineNumber, bool &aHasContinuationSection, int aExprBalance = 0);
	ResultType GetLineContExpr(TextStream *ts, LineBuffer &aBuf, LineBuffer &aNextBuf
		, LineNumberType &aPhysLineNumber, bool &aHasContinuationSection);
	ResultType BalanceExprError(int aBalance, TCHAR aExpect[], LPTSTR aLineText);
	static bool IsFunctionDefinition(LPTSTR aBuf, LPTSTR aNextBuf);
	ResultType IsDirective(LPTSTR aBuf);
	ResultType ConvertDirectiveBool(LPTSTR aBuf, bool &aResult, bool aDefault);
	ResultType ParseAndAddLine(LPTSTR aLineText, ActionTypeType aActionType = ACT_INVALID
		, LPTSTR aLiteralMap = NULL, size_t aLiteralMapLength = 0);
	ResultType ParseOperands(LPTSTR aArgText, LPTSTR aArgMap, DerefList &aDeref, int *aPos = NULL, TCHAR aEndChar = 0);
	ResultType ParseDoubleDeref(LPTSTR aArgText, LPTSTR aArgMap, DerefList &aDeref, int *aPos);
	ResultType ParseFatArrow(LPTSTR aArgText, LPTSTR aArgMap, DerefList &aDeref
		, LPTSTR aPrmStart, LPTSTR aPrmEnd, LPTSTR aExpr, LPTSTR &aExprEnd);
	ResultType ParseFatArrow(DerefType &aDeref, LPTSTR aPrmStart, LPTSTR aPrmEnd, LPTSTR aExpr, LPTSTR aExprEnd, LPTSTR aExprMap);
	LPTSTR ParseActionType(LPTSTR aBufTarget, LPTSTR aBufSource, bool aDisplayErrors);
	static ActionTypeType ConvertActionType(LPTSTR aActionTypeString, int aFirstAction, int aLastActionPlus1);
	ResultType AddLabel(LPTSTR aLabelName, bool aAllowDupe);
	void RemoveLabel(Label *aLabel);
	ResultType AddLine(ActionTypeType aActionType, LPTSTR aArg[] = NULL, int aArgc = 0, LPTSTR aArgMap[] = NULL, bool aAllArgsAreExpressions = false);

	// These aren't in the Line class because I think they're easier to implement
	// if aStartingLine is allowed to be NULL (for recursive calls).  If they
	// were member functions of class Line, a check for NULL would have to
	// be done before dereferencing any line's mNextLine, for example:
	ResultType PreparseExpressions(Line *aStartingLine);
	void PreparseHotkeyIfExpr(Line *aLine);
	Line *PreparseBlocks(Line *aStartingLine, ExecUntilMode aMode = NORMAL_MODE, Line *aParentLine = NULL, const ActionTypeType aLoopType = ACT_INVALID);
	Line *PreparseBlocksStmtBody(Line *aStartingLine, Line *aParentLine = NULL, const ActionTypeType aLoopType = ACT_INVALID);
	Line *PreparseCommands(Line *aStartingLine);
	bool IsLabelTarget(Line *aLine);

public:
	Line *mCurrLine;     // Seems better to make this public than make Line our friend.
	Label *mPlaceholderLabel; // Used in place of a NULL label to simplify code.
	LPTSTR mThisHotkeyName, mPriorHotkeyName;
	MsgMonitorList mOnExit, mOnClipboardChange, mOnError; // Event handlers for OnExit(), OnClipboardChange() and OnError().
	bool mOnClipboardChangeIsRunning;
	ExitReasons mExitReason;

	ScriptTimer *mFirstTimer, *mLastTimer;  // The first and last script timers in the linked list.
	UINT mTimerCount, mTimerEnabledCount;

	UserMenu *mFirstMenu, *mLastMenu;
	UINT mMenuCount;

	DWORD mThisHotkeyStartTime, mPriorHotkeyStartTime;  // Tickcount timestamp of when its subroutine began.
	TCHAR mEndChar;  // The ending character pressed to trigger the most recent non-auto-replace hotstring.
	modLR_type mThisHotkeyModifiersLR;
	LPTSTR mFileSpec; // Will hold the full filespec, for convenience.
	LPTSTR mFileDir;  // Will hold the directory containing the script file.
	LPTSTR mFileName; // Will hold the script's naked file name.
	LPTSTR mScriptName; // Value of A_ScriptName; defaults to mFileName if NULL. See also DefaultDialogTitle().
	LPTSTR mOurEXE; // Will hold this app's module name (e.g. C:\Program Files\AutoHotkey\AutoHotkey.exe).
	LPTSTR mOurEXEDir;  // Same as above but just the containing directory (for convenience).
	LPTSTR mMainWindowTitle; // Will hold our main window's title, for consistency & convenience.
	bool mIsReadyToExecute;
	bool mAutoExecSectionIsRunning;
	bool mIsRestart; // The app is restarting rather than starting from scratch.
	bool mErrorStdOut; // true if load-time syntax errors should be sent to stdout vs. a MsgBox.
	UINT mErrorStdOutCP;
	void SetErrorStdOut(LPTSTR aParam);
	void PrintErrorStdOut(LPCTSTR aErrorText, int aLength = 0, LPCTSTR aFile = _T("*"));
	void PrintErrorStdOut(LPCTSTR aErrorText, LPCTSTR aExtraInfo, FileIndexType aFileIndex, LineNumberType aLineNumber);
#ifndef AUTOHOTKEYSC
	TextStream *mIncludeLibraryFunctionsThenExit;
#endif
	int mUninterruptedLineCountMax; // 32-bit for performance (since huge values seem unnecessary here).
	int mUninterruptibleTime;
	DWORD mLastPeekTime;

	CStringW mRunAsUser, mRunAsPass, mRunAsDomain;

	HICON mCustomIcon;  // NULL unless the script has loaded a custom icon during its runtime.
	HICON mCustomIconSmall; // L17: Use separate big/small icons for best results.
	LPTSTR mCustomIconFile; // Filename of icon.  Allocated on first use.
	bool mIconFrozen; // If true, the icon does not change state when the state of pause or suspend changes.
	LPTSTR mTrayIconTip;  // Custom tip text for tray icon.  Allocated on first use.
	UINT mCustomIconNumber; // The number of the icon inside the above file.

	UserMenu *mTrayMenu; // Our tray menu, which should be destroyed upon exiting the program.
    
	SimpleList<ScriptModule*> mModuleSimpleList; // A linear list off all modules, used for non-recursive access to any module.

	ResultType Init(global_struct &g, LPTSTR aScriptFilename, bool aIsRestart);
	ResultType CreateWindows();
	void EnableClipboardListener(bool aEnable);
#ifdef AUTOHOTKEYSC
	void AllowMainWindow(bool aAllow);
	void EnableOrDisableViewMenuItems(HMENU aMenu, UINT aFlags);
#endif
	void CreateTrayIcon();
	void UpdateTrayIcon(bool aForceUpdate = false);
	void ShowTrayIcon(bool aShow);
	ResultType SetTrayIcon(LPTSTR aIconFile, int aIconNumber, ToggleValueType aFreezeIcon);
	void SetTrayTip(LPTSTR aText);
	ResultType AutoExecSection();
	bool IsPersistent();
	void ExitIfNotPersistent(ExitReasons aExitReason);
	ResultType Edit();
	ResultType Reload(bool aDisplayErrors);
	ResultType ExitApp(ExitReasons aExitReason, int aExitCode = 0);
	void TerminateApp(ExitReasons aExitReason, int aExitCode); // L31: Added aExitReason. See script.cpp.
	LineNumberType LoadFromFile();
	ResultType LoadIncludedFile(LPTSTR aFileSpec, bool aAllowDuplicateInclude, bool aIgnoreLoadFailure);
	ResultType LoadIncludedFile(TextStream *fp);
	ResultType OpenIncludedFile(TextStream &ts, LPTSTR aFileSpec, bool aAllowDuplicateInclude, bool aIgnoreLoadFailure);
	LineNumberType CurrentLine();
	LPTSTR CurrentFile();

	enum SetTimerFlags
	{
		SET_TIMER_PERIOD = 1,
		SET_TIMER_STATE = 2,
		SET_TIMER_PRIORITY = 4,
		SET_TIMER_RESET = 8
	};
	ResultType UpdateOrCreateTimer(IObject *aCallback
		, bool aUpdatePeriod, __int64 aPeriod, bool aUpdatePriority, int aPriority);
	void DeleteTimer(IObject *aCallback);
	LPTSTR DefaultDialogTitle();
	UserFunc* CreateHotFunc(Var* aFuncGlobalVar[], int aGlobalVarCount);
	ResultType DefineFunc(LPTSTR aBuf, Var *aFuncGlobalVar[], bool aStatic = false, bool aIsInExpression = false);
#ifndef AUTOHOTKEYSC
	struct FuncLibrary
	{
		LPTSTR path;
		size_t length;
	};
	void InitFuncLibraries(FuncLibrary aLibs[]);
	void InitFuncLibrary(FuncLibrary &aLib, LPTSTR aPathBase, LPTSTR aPathSuffix);
	Func *FindFuncInLibrary(LPTSTR aFuncName, size_t aFuncNameLength, bool &aErrorWasShown, bool &aFileWasFound, bool aIsAutoInclude);
#endif
	Func *FindFunc(LPCTSTR aFuncName, size_t aFuncNameLength = -1, int *apInsertPos = NULL);
	FuncEntry *FindBuiltInFunc(LPTSTR aFuncName);
	UserFunc *AddFunc(LPCTSTR aFuncName, size_t aFuncNameLength, int aInsertPos, Object *aClassObject = NULL);

	ResultType DefineClass(LPTSTR aBuf);
	UserFunc *DefineClassInit(bool aStatic);
	ResultType DefineClassVars(LPTSTR aBuf, bool aStatic);
	ResultType DefineClassProperty(LPTSTR aBuf, bool aStatic, Var **aFuncGlobalVar, bool &aBufHasBraceOrNotNeeded);
	ResultType DefineClassPropertyXet(LPTSTR aBuf, LPTSTR aEnd, Var **aFuncGlobalVar);
	Object *FindClass(LPCTSTR aClassName, size_t aClassNameLength = 0);
	ResultType ResolveClasses();
	ResultType InitClasses();

	static SymbolType ConvertWordOperator(LPCTSTR aWord, size_t aLength);
	static bool EndsWithOperator(LPTSTR aBuf, LPTSTR aBuf_marker);

	#define FINDVAR_DEFAULT  (VAR_LOCAL | VAR_GLOBAL)
	#define FINDVAR_GLOBAL   VAR_GLOBAL
	#define FINDVAR_LOCAL    VAR_LOCAL
	Var *FindOrAddVar(LPTSTR aVarName, size_t aVarNameLength = 0, int aScope = FINDVAR_DEFAULT);
	Var *FindVar(LPTSTR aVarName, size_t aVarNameLength = 0, int *apInsertPos = NULL
		, int aScope = FINDVAR_DEFAULT
		, bool *apIsLocal = NULL);
	Var *AddVar(LPTSTR aVarName, size_t aVarNameLength, int aInsertPos, int aScope);
	static VarEntry *GetBuiltInVar(LPTSTR aVarName);

	ResultType DerefInclude(LPTSTR &aOutput, LPTSTR aBuf);

	WinGroup *FindGroup(LPTSTR aGroupName, bool aCreateIfNotFound = false);
	ResultType AddGroup(LPTSTR aGroupName);
	Label *FindLabel(LPTSTR aLabelName);

	ResultType DoRunAs(LPTSTR aCommandLine, LPTSTR aWorkingDir, bool aDisplayErrors, WORD aShowWindow
		, Var *aOutputVar, PROCESS_INFORMATION &aPI, bool &aSuccess, HANDLE &aNewProcess, DWORD &aLastError);
	ResultType ActionExec(LPTSTR aAction, LPTSTR aParams = NULL, LPTSTR aWorkingDir = NULL
		, bool aDisplayErrors = true, LPTSTR aRunShowMode = NULL, HANDLE *aProcess = NULL
		, bool aUpdateLastError = false, bool aUseRunAs = false, Var *aOutputVar = NULL);

	LPTSTR ListVars(LPTSTR aBuf, int aBufSize);
	LPTSTR ListKeyHistory(LPTSTR aBuf, int aBufSize);

	UINT GetFreeMenuItemID();
	UserMenu *FindMenu(HMENU aMenuHandle);
	UserMenu *AddMenu(UserMenu *aMenu);
	ResultType ScriptDeleteMenu(UserMenu *aMenu);
	UserMenuItem *FindMenuItemByID(UINT aID)
	{
		for (UserMenu *m = mFirstMenu; m; m = m->mNextMenu)
			if (UserMenuItem *mi = m->FindItemByID(aID))
				return mi;
		return NULL;
	}
	UserMenuItem *FindMenuItemBySubmenu(HMENU aSubmenu) // L26: Used by WM_MEASUREITEM/WM_DRAWITEM to find the menu item with an associated submenu. Fixes icons on such items when owner-drawn menus are in use.
	{
		UserMenuItem *mi;
		for (UserMenu *m = mFirstMenu; m; m = m->mNextMenu)
			for (mi = m->mFirstMenuItem; mi; mi = mi->mNextMenuItem)
				if (mi->mSubmenu && mi->mSubmenu->mMenu == aSubmenu)
					return mi;
		return NULL;
	}

	static ResultType SetCoordMode(LPTSTR aCommand, LPTSTR aMode);
	static ResultType SetSendMode(LPTSTR aValue);
	static ResultType SetSendLevel(int aValue, LPTSTR aValueStr);

	// Call this ScriptError to avoid confusion with Line's error-displaying functions.
	// Use it for load time errors and non-continuable runtime errors:
	ResultType ScriptError(LPCTSTR aErrorText, LPCTSTR aExtraInfo = _T("")); // , ResultType aErrorType = FAIL);
	// CriticalError forces the program to exit after displaying an error.
	// Bypasses try/catch but does allow OnError and OnExit callbacks.
	ResultType CriticalError(LPCTSTR aErrorText, LPCTSTR aExtraInfo = _T(""));
	// RuntimeError allows the user to choose to continue, in which case OK is returned instead of FAIL;
	// therefore, caller must not rely on a FAIL result to abort the overall operation.
	ResultType RuntimeError(LPCTSTR aErrorText, LPCTSTR aExtraInfo = _T(""), ResultType aErrorType = FAIL_OR_OK, Line *aLine = nullptr);

	ResultType ConflictingDeclarationError(LPCTSTR aDeclType, Var *aExisting);
	enum AssignmentErrorType { INVALID_ASSIGNMENT = 1, INVALID_OUTPUT_VAR };
	ResultType VarIsReadOnlyError(Var *aVar, int aErrorType = INVALID_ASSIGNMENT);

	ResultType ShowError(LPCTSTR aErrorText, ResultType aErrorType, LPCTSTR aExtraInfo, Line *aLine);
	int FormatError(LPTSTR aBuf, int aBufSize, ResultType aErrorType, LPCTSTR aErrorText, LPCTSTR aExtraInfo, Line *aLine);

	void ScriptWarning(WarnMode warnMode, LPCTSTR aWarningText, LPCTSTR aExtraInfo = _T(""), Line *line = NULL);
	void WarnUninitializedVar(Var *var);
	void MaybeWarnLocalSameAsGlobal(UserFunc &func, Var &var);

	ResultType PreprocessLocalVars(FuncList &aFuncs);
	ResultType PreprocessLocalVars(UserFunc &aFunc, Var **aVarList, int &aVarCount);
	ResultType PreprocessFindUpVar(LPTSTR aName, UserFunc &aOuter, UserFunc &aInner, Var *&aFound, Var *aLocal);
	void ConvertLocalToAlias(Var &aLocal, Var *aAliasFor, int aPos, Var **aVarList, int &aVarCount);
	ResultType RetroactivelyFixConstants();

	ResultType ThrowIfTrue(bool aError);
	ResultType ThrowIntIfNonzero(int aErrorValue, LPCTSTR aWhat = NULL);
	ResultType ThrowRuntimeException(LPCTSTR aErrorText, LPCTSTR aWhat, LPCTSTR aExtraInfo, Line *aLine, ResultType aErrorType);
	ResultType ThrowRuntimeException(LPCTSTR aErrorText, LPCTSTR aWhat = nullptr, LPCTSTR aExtraInfo = _T(""));
	ResultType Win32Error(DWORD aError = GetLastError());
	
	ResultType UnhandledException(Line* aLine, ResultType aErrorType = FAIL);
	static void FreeExceptionToken(ResultToken*& aToken);


	#define SOUNDPLAY_ALIAS _T("AHK_PlayMe")  // Used by destructor and SoundPlay().

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

BIV_DECL_RW(BIV_Clipboard);
BIV_DECL_R (BIV_MMM_DDD);
BIV_DECL_R (BIV_DateTime);
BIV_DECL_RW(BIV_ListLines);
BIV_DECL_RW(BIV_TitleMatchMode);
BIV_DECL_R (BIV_TitleMatchModeSpeed); // Write is handled by BIV_TitleMatchMode_Set.
BIV_DECL_RW(BIV_DetectHiddenWindows);
BIV_DECL_RW(BIV_DetectHiddenText);
BIV_DECL_RW(BIV_StringCaseSense);
BIV_DECL_RW(BIV_xDelay);
BIV_DECL_RW(BIV_DefaultMouseSpeed);
BIV_DECL_RW(BIV_CoordMode);
BIV_DECL_RW(BIV_SendMode);
BIV_DECL_RW(BIV_SendLevel);
BIV_DECL_RW(BIV_StoreCapsLockMode);
BIV_DECL_R (BIV_IsPaused);
BIV_DECL_R (BIV_IsCritical);
BIV_DECL_R (BIV_IsSuspended);
BIV_DECL_R (BIV_IsCompiled);
BIV_DECL_RW(BIV_FileEncoding);
BIV_DECL_RW(BIV_RegView);
BIV_DECL_RW(BIV_LastError);
BIV_DECL_RW(BIV_AllowMainWindow);
BIV_DECL_R (BIV_TrayMenu);
BIV_DECL_RW(BIV_IconHidden);
BIV_DECL_RW(BIV_IconTip);
BIV_DECL_R (BIV_IconFile);
BIV_DECL_R (BIV_IconNumber);
BIV_DECL_R (BIV_Space_Tab);
BIV_DECL_R (BIV_AhkVersion);
BIV_DECL_R (BIV_AhkPath);
BIV_DECL_R (BIV_TickCount);
BIV_DECL_R (BIV_Now);
BIV_DECL_R (BIV_OSVersion);
BIV_DECL_R (BIV_Is64bitOS);
BIV_DECL_R (BIV_Language);
BIV_DECL_R (BIV_UserName_ComputerName);
BIV_DECL_RW(BIV_WorkingDir);
BIV_DECL_R (BIV_InitialWorkingDir);
BIV_DECL_R (BIV_WinDir);
BIV_DECL_R (BIV_Temp);
BIV_DECL_R (BIV_ComSpec);
BIV_DECL_R (BIV_SpecialFolderPath); // Handles various variables.
BIV_DECL_R (BIV_MyDocuments);
BIV_DECL_R (BIV_Caret);
BIV_DECL_R (BIV_Cursor);
BIV_DECL_R (BIV_ScreenWidth_Height);
BIV_DECL_RW(BIV_ScriptName);
BIV_DECL_R (BIV_ScriptDir);
BIV_DECL_R (BIV_ScriptFullPath);
BIV_DECL_R (BIV_ScriptHwnd);
BIV_DECL_R (BIV_LineNumber);
BIV_DECL_R (BIV_LineFile);
BIV_DECL_R (BIV_LoopFileName);
BIV_DECL_R (BIV_LoopFileExt);
BIV_DECL_R (BIV_LoopFileDir);
BIV_DECL_R (BIV_LoopFilePath);
BIV_DECL_R (BIV_LoopFileFullPath);
BIV_DECL_R (BIV_LoopFileShortPath);
BIV_DECL_R (BIV_LoopFileTime);
BIV_DECL_R (BIV_LoopFileAttrib);
BIV_DECL_R (BIV_LoopFileSize);
BIV_DECL_R (BIV_LoopRegType);
BIV_DECL_R (BIV_LoopRegKey);
BIV_DECL_R (BIV_LoopRegName);
BIV_DECL_R (BIV_LoopRegTimeModified);
BIV_DECL_R (BIV_LoopReadLine);
BIV_DECL_R (BIV_LoopField);
BIV_DECL_RW(BIV_LoopIndex);
BIV_DECL_R (BIV_ThisFunc);
BIV_DECL_R (BIV_ThisHotkey);
BIV_DECL_R (BIV_PriorHotkey);
BIV_DECL_R (BIV_TimeSinceThisHotkey);
BIV_DECL_R (BIV_TimeSincePriorHotkey);
BIV_DECL_R (BIV_EndChar);
BIV_DECL_RW(BIV_EventInfo);
BIV_DECL_R (BIV_TimeIdle);
BIV_DECL_R (BIV_TimeIdlePhysical);
BIV_DECL_R (BIV_IPAddress);
BIV_DECL_R (BIV_IsAdmin);
BIV_DECL_R (BIV_PtrSize);
BIV_DECL_R (BIV_PriorKey);
BIV_DECL_R (BIV_ScreenDPI);


////////////////////////
// BUILT-IN FUNCTIONS //
////////////////////////

#ifdef ENABLE_DLLCALL
void *GetDllProcAddress(LPCTSTR aDllFileFunc, HMODULE *hmodule_to_free = NULL);
BIF_DECL(BIF_DllCall);
#endif

BIF_DECL(BIF_StrCompare);
BIF_DECL(BIF_String);
BIF_DECL(BIF_StrLen);
BIF_DECL(BIF_SubStr);
BIF_DECL(BIF_InStr);
BIF_DECL(BIF_StrCase);
BIF_DECL(BIF_StrSplit);
BIF_DECL(BIF_StrReplace);
BIF_DECL(BIF_Sort);
BIF_DECL(BIF_RegEx);
BIF_DECL(BIF_Ord);
BIF_DECL(BIF_Chr);
BIF_DECL(BIF_Format);
BIF_DECL(BIF_FormatTime);
BIF_DECL(BIF_NumGet);
BIF_DECL(BIF_NumPut);
BIF_DECL(BIF_StrGetPut);
BIF_DECL(BIF_StrPtr);
BIF_DECL(BIF_IsLabel);
BIF_DECL(BIF_IsFunc);
BIF_DECL(BIF_Func);
BIF_DECL(BIF_IsByRef);
BIF_DECL(BIF_IsTypeish);
BIF_DECL(BIF_IsSet);
BIF_DECL(BIF_GetKeyState);
BIF_DECL(BIF_GetKeyName);
BIF_DECL(BIF_VarSetStrCapacity);
BIF_DECL(BIF_FileExist);
BIF_DECL(BIF_WinExistActive);
BIF_DECL(BIF_Round);
BIF_DECL(BIF_FloorCeil);
BIF_DECL(BIF_Integer);
BIF_DECL(BIF_Float);
BIF_DECL(BIF_Mod);
BIF_DECL(BIF_Abs);
BIF_DECL(BIF_Sin);
BIF_DECL(BIF_Cos);
BIF_DECL(BIF_Tan);
BIF_DECL(BIF_ASinACos);
BIF_DECL(BIF_ATan);
BIF_DECL(BIF_Exp);
BIF_DECL(BIF_SqrtLogLn);
BIF_DECL(BIF_MinMax);
BIF_DECL(BIF_DateAdd);
BIF_DECL(BIF_DateDiff);
BIF_DECL(BIF_Env);
BIF_DECL(BIF_SysGet);
BIF_DECL(BIF_SysGetIPAddresses);
BIF_DECL(BIF_PostSendMessage);
BIF_DECL(BIF_Hotkey);
BIF_DECL(BIF_SetTimer);
BIF_DECL(BIF_OnMessage);
BIF_DECL(BIF_On);
BIF_DECL(BIF_BufferAlloc);
BIF_DECL(BIF_ClipboardAll);

#ifdef ENABLE_REGISTERCALLBACK
BIF_DECL(BIF_CallbackCreate);
BIF_DECL(BIF_CallbackFree);
#endif

BIF_DECL(BIF_MenuFromHandle);
BIF_DECL(BIF_TraySetIcon);

BIF_DECL(BIF_MsgBox);
BIF_DECL(BIF_InputBox);

// Gui
BIF_DECL(BIF_GuiFromHwnd);
BIF_DECL(BIF_GuiCtrlFromHwnd);

BIF_DECL(BIF_IL_Create);
BIF_DECL(BIF_IL_Destroy);
BIF_DECL(BIF_IL_Add);

BIF_DECL(BIF_LoadPicture);

BIF_DECL(BIF_Trim); // L31: Also handles LTrim and RTrim.

BIF_DECL(BIF_Hotstring);
BIF_DECL(BIF_InputHook);

BIF_DECL(BIF_Type);
BIF_DECL(BIF_IsObject);


BIF_DECL(BIF_Object);
BIF_DECL(BIF_Array);
BIF_DECL(BIF_Map);
BIF_DECL(BIF_ObjAddRefRelease);
BIF_DECL(BIF_ObjBindMethod);
BIF_DECL(BIF_ObjPtr);
// Built-ins also available as methods -- these are available as functions for use primarily by overridden methods (i.e. where using the built-in methods isn't possible as they're no longer accessible).
BIF_DECL(BIF_ObjXXX);

BIF_DECL(BIF_Base);
BIF_DECL(BIF_HasBase);
BIF_DECL(BIF_HasProp);
BIF_DECL(BIF_HasMethod);
BIF_DECL(BIF_GetMethod);

// Expression operators implemented via SYM_FUNC:
BIF_DECL(Op_ObjInvoke);
BIF_DECL(Op_ObjGetInPlace);
BIF_DECL(Op_ObjIncDec);


// Advanced file IO interfaces
BIF_DECL(BIF_FileOpen);


// COM interop
BIF_DECL(BIF_ComObject);
BIF_DECL(BIF_ComObjActive);
BIF_DECL(BIF_ComObjCreate);
BIF_DECL(BIF_ComObjGet);
BIF_DECL(BIF_ComObjConnect);
BIF_DECL(BIF_ComObjType);
BIF_DECL(BIF_ComObjValue);
BIF_DECL(BIF_ComObjFlags);
BIF_DECL(BIF_ComObjArray);
BIF_DECL(BIF_ComObjQuery);


BIF_DECL(BIF_Exception);


BIF_DECL(BIF_Control);
BIF_DECL(BIF_ControlClick);
BIF_DECL(BIF_ControlFocus);
BIF_DECL(BIF_ControlGet);
BIF_DECL(BIF_ControlGetClassNN);
BIF_DECL(BIF_ControlGetFocus);
BIF_DECL(BIF_ControlGetPos);
BIF_DECL(BIF_ControlGetText);
BIF_DECL(BIF_ControlMove);
BIF_DECL(BIF_ControlSend);
BIF_DECL(BIF_ControlSetText);
BIF_DECL(BIF_DirSelect);
BIF_DECL(BIF_Drive);
BIF_DECL(BIF_DriveGet);
BIF_DECL(BIF_FileAppend);
BIF_DECL(BIF_FileGetAttrib);
BIF_DECL(BIF_FileGetShortcut);
BIF_DECL(BIF_FileGetSize);
BIF_DECL(BIF_FileGetTime);
BIF_DECL(BIF_FileGetVersion);
BIF_DECL(BIF_FileRead);
BIF_DECL(BIF_FileSelect);
BIF_DECL(BIF_GroupActivate);
BIF_DECL(BIF_ImageSearch);
BIF_DECL(BIF_IniRead);
BIF_DECL(BIF_MouseGetPos);
BIF_DECL(BIF_PixelGetColor);
BIF_DECL(BIF_PixelSearch);
BIF_DECL(BIF_Reg);
BIF_DECL(BIF_Random);
BIF_DECL(BIF_Run);
BIF_DECL(BIF_Sound);
BIF_DECL(BIF_SplitPath);
BIF_DECL(BIF_StatusBarGetText);
BIF_DECL(BIF_StatusBarWait);
BIF_DECL(BIF_CaretGetPos);
BIF_DECL(BIF_WinGetClass);
BIF_DECL(BIF_WinGetText);
BIF_DECL(BIF_WinGetTitle);
BIF_DECL(BIF_WinGetPos);
BIF_DECL(BIF_WinGet);
BIF_DECL(BIF_WinSet);
BIF_DECL(BIF_WinSetTitle);
BIF_DECL(BIF_WinRedraw);
BIF_DECL(BIF_WinMove);
BIF_DECL(BIF_WinMoveTopBottom);
BIF_DECL(BIF_WinShow);
BIF_DECL(BIF_WinActivate);
BIF_DECL(BIF_MenuSelect);
BIF_DECL(BIF_Process);
BIF_DECL(BIF_ProcessSetPriority);
BIF_DECL(BIF_MonitorGet);
BIF_DECL(BIF_Wait);

BIF_DECL(BIF_PerformAction);
BIF_DECL(BIF_SetBIV);


BOOL ResultToBOOL(LPTSTR aResult);
BOOL VarToBOOL(Var &aVar);
BOOL TokenToBOOL(ExprTokenType &aToken);
ToggleValueType TokenToToggleValue(ExprTokenType &aToken);
SymbolType TokenIsNumeric(ExprTokenType &aToken);
SymbolType TokenIsPureNumeric(ExprTokenType &aToken);
SymbolType TokenIsPureNumeric(ExprTokenType &aToken, SymbolType &aIsImpureNumeric);
BOOL TokenIsEmptyString(ExprTokenType &aToken);
BOOL TokenIsEmptyString(ExprTokenType &aToken, BOOL aWarnUninitializedVar); // Same as TokenIsEmptyString but optionally warns if the token is an uninitialized var.
SymbolType TypeOfToken(ExprTokenType &aToken);
SymbolType TypeOfToken(ExprTokenType &aToken, SymbolType &aIsNum);
__int64 TokenToInt64(ExprTokenType &aToken);
double TokenToDouble(ExprTokenType &aToken, BOOL aCheckForHex = TRUE);
LPTSTR TokenToString(ExprTokenType &aToken, LPTSTR aBuf = NULL, size_t *aLength = NULL);
ResultType TokenToDoubleOrInt64(const ExprTokenType &aInput, ExprTokenType &aOutput);
IObject *TokenToObject(ExprTokenType &aToken); // L31
Func *TokenToFunc(ExprTokenType &aToken);
IObject *TokenToFunctor(ExprTokenType &aToken);
IObject *StringToFunctor(LPTSTR aStr);
ResultType ValidateFunctor(IObject *aFunc, int aParamCount, ResultToken &aResultToken, LPTSTR aNullErr = ERR_TYPE_MISMATCH, int *aMinParams = nullptr);
ResultType TokenSetResult(ResultToken &aResultToken, LPCTSTR aValue, size_t aLength = -1);
BOOL TokensAreEqual(ExprTokenType &left, ExprTokenType &right);
LPTSTR TokenTypeString(ExprTokenType &aToken);
#define STRING_TYPE_STRING _T("String")
#define INTEGER_TYPE_STRING _T("Integer")
#define FLOAT_TYPE_STRING _T("Float")

LPTSTR RegExMatch(LPTSTR aHaystack, LPTSTR aNeedleRegEx);
ResultType SetWorkingDir(LPTSTR aNewDir);
void UpdateWorkingDir(LPTSTR aNewDir = NULL);
LPTSTR GetWorkingDir();
int ConvertJoy(LPTSTR aBuf, int *aJoystickID = NULL, bool aAllowOnlyButtons = false);
bool ScriptGetKeyState(vk_type aVK, KeyStateTypes aKeyStateType);
bool ScriptGetJoyState(JoyControls aJoy, int aJoystickID, ExprTokenType &aToken, LPTSTR aBuf);

ResultType DetermineTargetHwnd(HWND &aWindow, ResultToken &aResultToken, ExprTokenType &aToken);
ResultType DetermineTargetWindow(HWND &aWindow, ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, int aNonWinParamCount = 0);
ResultType DetermineTargetControl(HWND &aControl, HWND &aWindow, ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, int aNonWinParamCount = 0, bool aThrowIfNotFound = true);
#define DETERMINE_TARGET_CONTROL(param_offset) \
	HWND target_window, control_window; \
	if (!DetermineTargetControl(control_window, target_window, aResultToken, aParam + param_offset, aParamCount - param_offset)) \
		return;

LPTSTR GetExitReasonString(ExitReasons aExitReason);

void ControlGetListView(ResultToken &aResultToken, HWND aHwnd, LPTSTR aOptions);
bool ControlSetTab(ResultToken &aResultToken, HWND aHwnd, DWORD aTabIndex);

void PixelSearch(Var *aOutputVarX, Var *aOutputVarY
	, int aLeft, int aTop, int aRight, int aBottom, COLORREF aColorRGB
	, int aVariation, bool aIsPixelGetColor, ResultToken &aResultToken);

ResultType GetObjectPtrProperty(IObject *aObject, LPTSTR aPropName, UINT_PTR &aPtr, ResultToken &aResultToken, bool aOptional = false);
ResultType GetObjectIntProperty(IObject *aObject, LPTSTR aPropName, __int64 &aValue, ResultToken &aResultToken, bool aOptional = false);
void GetBufferObjectPtr(ResultToken &aResultToken, IObject *obj, size_t &aPtr, size_t &aSize);
void GetBufferObjectPtr(ResultToken &aResultToken, IObject *obj, size_t &aPtr);


#endif

