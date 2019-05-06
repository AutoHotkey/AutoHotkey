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
#include "script.h"
#include "globaldata.h" // for a lot of things
#include "util.h" // for strlcpy() etc.
#include "mt19937ar-cok.h" // for random number generator
#include "window.h" // for a lot of things
#include "application.h" // for MsgSleep()
#include "TextIO.h"

// Globals that are for only this module:
#define MAX_COMMENT_FLAG_LENGTH 15
static TCHAR g_CommentFlag[MAX_COMMENT_FLAG_LENGTH + 1] = _T(";"); // Adjust the below for any changes.
static size_t g_CommentFlagLength = 1; // pre-calculated for performance
static ExprOpFunc g_ObjGet(BIF_ObjInvoke, IT_GET), g_ObjSet(BIF_ObjInvoke, IT_SET);
static ExprOpFunc g_ObjGetInPlace(BIF_ObjGetInPlace, IT_GET);
static ExprOpFunc g_ObjNew(BIF_ObjNew, IT_CALL);
static ExprOpFunc g_ObjPreInc(BIF_ObjIncDec, SYM_PRE_INCREMENT), g_ObjPreDec(BIF_ObjIncDec, SYM_PRE_DECREMENT)
				, g_ObjPostInc(BIF_ObjIncDec, SYM_POST_INCREMENT), g_ObjPostDec(BIF_ObjIncDec, SYM_POST_DECREMENT);
ExprOpFunc g_ObjCall(BIF_ObjInvoke, IT_CALL); // Also needed in script_expression.cpp.


#define VF(name, fn) { _T(#name), fn }
#define A_x(name, fn) { _T(#name), fn }
#define A_(name) A_x(name, BIV_##name)
// IMPORTANT: Both of the following arrays must be kept in alphabetical order
// for binary search to work.  See Script::GetBuiltInVar for further comments.
// g_BIV: All built-in vars not beginning with "A_".  Keeping these separate allows
// the search to be limited to just these few whenever the var name does not begin
// with "A_", as for most user-defined variables.  This helps average-case performance.
VarEntry g_BIV[] =
{
	VF(Clipboard, (BuiltInVarType)VAR_CLIPBOARD),
	VF(ClipboardAll, (BuiltInVarType)VAR_CLIPBOARDALL),
	VF(ComSpec, BIV_ComSpec), // Lacks an "A_" prefix for backward compatibility with pre-NoEnv scripts and also it's easier to type & remember.,
	VF(False, BIV_True_False),
	VF(ProgramFiles, BIV_SpecialFolderPath), // v1.0.43.08: Added to ease the transition to #NoEnv.,
	VF(True, BIV_True_False)
};
// g_BIV_A: All built-in vars beginning with "A_".  The prefix is omitted from each
// name to reduce code size and speed up the comparisons.
VarEntry g_BIV_A[] =
{
	A_(AhkPath),
	A_(AhkVersion),
	A_x(AppData, BIV_SpecialFolderPath),
	A_x(AppDataCommon, BIV_SpecialFolderPath),
	A_(AutoTrim),
	A_x(BatchLines, BIV_BatchLines),
	A_x(CaretX, BIV_Caret),
	A_x(CaretY, BIV_Caret),
	A_x(ComputerName, BIV_UserName_ComputerName),
	A_(ComSpec),
	A_x(ControlDelay, BIV_xDelay),
	A_x(CoordModeCaret, BIV_CoordMode),
	A_x(CoordModeMenu, BIV_CoordMode),
	A_x(CoordModeMouse, BIV_CoordMode),
	A_x(CoordModePixel, BIV_CoordMode),
	A_x(CoordModeToolTip, BIV_CoordMode),
	A_(Cursor),
	A_x(DD, BIV_DateTime),
	A_x(DDD, BIV_MMM_DDD),
	A_x(DDDD, BIV_MMM_DDD),
	A_x(DefaultGui, BIV_DefaultGui),
	A_x(DefaultListView, BIV_DefaultGui),
	A_(DefaultMouseSpeed),
	A_x(DefaultTreeView, BIV_DefaultGui),
	A_x(Desktop, BIV_SpecialFolderPath),
	A_x(DesktopCommon, BIV_SpecialFolderPath),
	A_(DetectHiddenText),
	A_(DetectHiddenWindows),
	A_(EndChar),
	A_(EventInfo), // It's called "EventInfo" vs. "GuiEventInfo" because it applies to non-Gui events such as OnClipboardChange.,
	A_(ExitReason),
	A_(FileEncoding),
	A_(FormatFloat),
	A_(FormatInteger),
	A_x(Gui, BIV_Gui),
	A_(GuiControl),
	A_x(GuiControlEvent, BIV_GuiEvent),
	A_x(GuiEvent, BIV_GuiEvent), // v1.0.36: A_GuiEvent was added as a synonym for A_GuiControlEvent because it seems unlikely that A_GuiEvent will ever be needed for anything:,
	A_x(GuiHeight, BIV_Gui),
	A_x(GuiWidth, BIV_Gui),
	A_x(GuiX, BIV_Gui), // Naming: Brevity seems more a benefit than would A_GuiEventX's improved clarity.,
	A_x(GuiY, BIV_Gui), // These can be overloaded if a GuiMove label or similar is ever needed.,
	A_x(Hour, BIV_DateTime),
	A_(IconFile),
	A_(IconHidden),
	A_(IconNumber),
	A_(IconTip),
	A_x(Index, BIV_LoopIndex),
	A_x(IPAddress1, BIV_IPAddress),
	A_x(IPAddress2, BIV_IPAddress),
	A_x(IPAddress3, BIV_IPAddress),
	A_x(IPAddress4, BIV_IPAddress),
	A_(Is64bitOS),
	A_(IsAdmin),
	A_(IsCompiled),
	A_(IsCritical),
	A_(IsPaused),
	A_(IsSuspended),
	A_(IsUnicode),
	A_x(KeyDelay, BIV_xDelay),
	A_x(KeyDelayPlay, BIV_xDelay),
	A_x(KeyDuration, BIV_xDelay),
	A_x(KeyDurationPlay, BIV_xDelay),
	A_(Language),
	A_(LastError),
	A_(LineFile),
	A_(LineNumber),
	A_(ListLines),
	A_(LoopField),
	A_(LoopFileAttrib),
	A_(LoopFileDir),
	A_(LoopFileExt),
	A_(LoopFileFullPath),
	A_(LoopFileLongPath),
	A_(LoopFileName),
	A_x(LoopFilePath, BIV_LoopFileFullPath),
	A_(LoopFileShortName),
	A_(LoopFileShortPath),
	A_x(LoopFileSize, BIV_LoopFileSize),
	A_x(LoopFileSizeKB, BIV_LoopFileSize),
	A_x(LoopFileSizeMB, BIV_LoopFileSize),
	A_x(LoopFileTimeAccessed, BIV_LoopFileTime),
	A_x(LoopFileTimeCreated, BIV_LoopFileTime),
	A_x(LoopFileTimeModified, BIV_LoopFileTime),
	A_(LoopReadLine),
	A_(LoopRegKey),
	A_(LoopRegName),
	A_(LoopRegSubKey),
	A_(LoopRegTimeModified),
	A_(LoopRegType),
	A_x(MDay, BIV_DateTime),
	A_x(Min, BIV_DateTime),
	A_x(MM, BIV_DateTime),
	A_x(MMM, BIV_MMM_DDD),
	A_x(MMMM, BIV_MMM_DDD),
	A_x(Mon, BIV_DateTime),
	A_x(MouseDelay, BIV_xDelay),
	A_x(MouseDelayPlay, BIV_xDelay),
	A_x(MSec, BIV_DateTime),
	A_(MyDocuments),
	A_x(Now, BIV_Now),
	A_x(NowUTC, BIV_Now),
	A_x(NumBatchLines, BIV_BatchLines),
	A_(OSType),
	A_(OSVersion),
	A_(PriorHotkey),
	A_(PriorKey),
	A_x(ProgramFiles, BIV_SpecialFolderPath),
	A_x(Programs, BIV_SpecialFolderPath),
	A_x(ProgramsCommon, BIV_SpecialFolderPath),
	A_(PtrSize),
	A_(RegView),
	A_(ScreenDPI),
	A_x(ScreenHeight, BIV_ScreenWidth_Height),
	A_x(ScreenWidth, BIV_ScreenWidth_Height),
	A_(ScriptDir),
	A_(ScriptFullPath),
	A_(ScriptHwnd),
	A_(ScriptName),
	A_x(Sec, BIV_DateTime),
	A_(SendLevel),
	A_(SendMode),
	A_x(Space, BIV_Space_Tab),
	A_x(StartMenu, BIV_SpecialFolderPath),
	A_x(StartMenuCommon, BIV_SpecialFolderPath),
	A_x(Startup, BIV_SpecialFolderPath),
	A_x(StartupCommon, BIV_SpecialFolderPath),
	A_(StoreCapslockMode),
	A_(StringCaseSense),
	A_x(Tab, BIV_Space_Tab),
	A_(Temp), // Debatably should be A_TempDir, but brevity seemed more popular with users, perhaps for heavy uses of the temp folder.,
	A_(ThisFunc),
	A_(ThisHotkey),
	A_(ThisLabel),
	A_(ThisMenu),
	A_(ThisMenuItem),
	A_(ThisMenuItemPos),
	A_(TickCount),
	A_(TimeIdle),
	A_x(TimeIdleKeyboard, BIV_TimeIdlePhysical),
	A_x(TimeIdleMouse, BIV_TimeIdlePhysical),
	A_(TimeIdlePhysical),
	A_(TimeSincePriorHotkey),
	A_(TimeSinceThisHotkey),
	A_(TitleMatchMode),
	A_(TitleMatchModeSpeed),
	A_x(UserName, BIV_UserName_ComputerName),
	A_x(WDay, BIV_DateTime),
	A_x(WinDelay, BIV_xDelay),
	A_(WinDir),
	A_(WorkingDir),
	A_x(YDay, BIV_DateTime),
	A_x(Year, BIV_DateTime),
	A_x(YWeek, BIV_DateTime),
	A_x(YYYY, BIV_DateTime)
};
#undef A_
#undef VF


// See Script::CreateWindows() for details about the following:
typedef BOOL (WINAPI* AddRemoveClipboardListenerType)(HWND);
static AddRemoveClipboardListenerType MyRemoveClipboardListener = (AddRemoveClipboardListenerType)
	GetProcAddress(GetModuleHandle(_T("user32")), "RemoveClipboardFormatListener");
static AddRemoveClipboardListenerType MyAddClipboardListener = (AddRemoveClipboardListenerType)
	GetProcAddress(GetModuleHandle(_T("user32")), "AddClipboardFormatListener");

// General note about the methods in here:
// Want to be able to support multiple simultaneous points of execution
// because more than one subroutine can be executing simultaneously
// (well, more precisely, there can be more than one script subroutine
// that's in a "currently running" state, even though all such subroutines,
// except for the most recent one, are suspended.  So keep this in mind when
// using things such as static data members or static local variables.


Script::Script()
	: mFirstLine(NULL), mLastLine(NULL), mCurrLine(NULL), mPlaceholderLabel(NULL), mFirstStaticLine(NULL), mLastStaticLine(NULL)
	, mThisHotkeyName(_T("")), mPriorHotkeyName(_T("")), mThisHotkeyStartTime(0), mPriorHotkeyStartTime(0)
	, mEndChar(0), mThisHotkeyModifiersLR(0)
	, mNextClipboardViewer(NULL), mOnClipboardChangeIsRunning(false), mOnClipboardChangeLabel(NULL)
	, mOnExitLabel(NULL), mExitReason(EXIT_NONE)
	, mFirstLabel(NULL), mLastLabel(NULL)
	, mFunc(NULL), mFuncCount(0), mFuncCountMax(0)
	, mFirstTimer(NULL), mLastTimer(NULL), mTimerEnabledCount(0), mTimerCount(0)
	, mFirstMenu(NULL), mLastMenu(NULL), mMenuCount(0), mThisMenuItem(NULL)
	, mVar(NULL), mVarCount(0), mVarCountMax(0), mLazyVar(NULL), mLazyVarCount(0)
	, mCurrentFuncOpenBlockCount(0), mNextLineIsFunctionBody(false), mNoUpdateLabels(false)
	, mClassObjectCount(0), mUnresolvedClasses(NULL), mClassProperty(NULL), mClassPropertyDef(NULL)
	, mCurrFileIndex(0), mCombinedLineNumber(0), mNoHotkeyLabels(true), mMenuUseErrorLevel(false)
	, mFileSpec(_T("")), mFileDir(_T("")), mFileName(_T("")), mOurEXE(_T("")), mOurEXEDir(_T("")), mMainWindowTitle(_T(""))
	, mIsReadyToExecute(false), mAutoExecSectionIsRunning(false)
	, mIsRestart(false), mErrorStdOut(false)
#ifndef AUTOHOTKEYSC
	, mIncludeLibraryFunctionsThenExit(NULL)
#endif
	, mLinesExecutedThisCycle(0), mUninterruptedLineCountMax(1000), mUninterruptibleTime(15)
	, mCustomIcon(NULL), mCustomIconSmall(NULL) // Normally NULL unless there's a custom tray icon loaded dynamically.
	, mCustomIconFile(NULL), mIconFrozen(false), mTrayIconTip(NULL) // Allocated on first use.
	, mCustomIconNumber(0)
{
	// v1.0.25: mLastScriptRest and mLastPeekTime are now initialized right before the auto-exec
	// section of the script is launched, which avoids an initial Sleep(10) in ExecUntil
	// that would otherwise occur.
	*mThisMenuItemName = *mThisMenuName = '\0';
	ZeroMemory(&mNIC, sizeof(mNIC));  // Constructor initializes this, to be safe.
	mNIC.hWnd = NULL;  // Set this as an indicator that it tray icon is not installed.

	// Lastly (after the above have been initialized), anything that can fail:
	if (   !(mTrayMenu = AddMenu(_T("Tray")))   ) // realistically never happens
	{
		ScriptError(_T("No tray mem"));
		ExitApp(EXIT_DESTROY);
	}
	else
		mTrayMenu->mIncludeStandardItems = true;

#ifdef _DEBUG
	if (ID_FILE_EXIT < ID_MAIN_FIRST) // Not a very thorough check.
		ScriptError(_T("DEBUG: ID_FILE_EXIT is too large (conflicts with IDs reserved via ID_USER_FIRST)."));
	if (MAX_CONTROLS_PER_GUI > ID_USER_FIRST - 3)
		ScriptError(_T("DEBUG: MAX_CONTROLS_PER_GUI is too large (conflicts with IDs reserved via ID_USER_FIRST)."));
	if (g_ActionCount != ACT_COUNT) // This enum value only exists in debug mode.
		ScriptError(_T("DEBUG: g_act and enum_act are out of sync."));
	int LargestMaxParams, i, j;
	ActionTypeType *np;
	// Find the Largest value of MaxParams used by any command and make sure it
	// isn't something larger than expected by the parsing routines:
	for (LargestMaxParams = i = 0; i < g_ActionCount; ++i)
	{
		if (g_act[i].MaxParams > LargestMaxParams)
			LargestMaxParams = g_act[i].MaxParams;
		// This next part has been tested and it does work, but only if one of the arrays
		// contains exactly MAX_NUMERIC_PARAMS number of elements and isn't zero terminated.
		// Relies on short-circuit boolean order:
		for (np = g_act[i].NumericParams, j = 0; j < MAX_NUMERIC_PARAMS && *np; ++j, ++np);
		if (j >= MAX_NUMERIC_PARAMS)
		{
			ScriptError(_T("DEBUG: At least one command has a NumericParams array that isn't zero-terminated.")
				_T("  This would result in reading beyond the bounds of the array."));
			return;
		}
	}
	if (LargestMaxParams > MAX_ARGS)
		ScriptError(_T("DEBUG: At least one command supports more arguments than allowed."));
	if (sizeof(ActionTypeType) == 1 && g_ActionCount > 256)
		ScriptError(_T("DEBUG: Since there are now more than 256 Action Types, the ActionTypeType")
			_T(" typedef must be changed."));
#endif
	OleInitialize(NULL);
}



Script::~Script() // Destructor.
{
	// MSDN: "Before terminating, an application must call the UnhookWindowsHookEx function to free
	// system resources associated with the hook."
	AddRemoveHooks(0); // Remove all hooks.
	if (mNIC.hWnd) // Tray icon is installed.
		Shell_NotifyIcon(NIM_DELETE, &mNIC); // Remove it.
	// Destroy any Progress/SplashImage windows that haven't already been destroyed.  This is necessary
	// because sometimes these windows aren't owned by the main window:
	int i;
	for (i = 0; i < MAX_PROGRESS_WINDOWS; ++i)
	{
		if (g_Progress[i].hwnd && IsWindow(g_Progress[i].hwnd))
			DestroyWindow(g_Progress[i].hwnd);
		if (g_Progress[i].hfont1) // Destroy font only after destroying the window that uses it.
			DeleteObject(g_Progress[i].hfont1);
		if (g_Progress[i].hfont2) // Destroy font only after destroying the window that uses it.
			DeleteObject(g_Progress[i].hfont2);
		if (g_Progress[i].hbrush)
			DeleteObject(g_Progress[i].hbrush);
	}
	for (i = 0; i < MAX_SPLASHIMAGE_WINDOWS; ++i)
	{
		if (g_SplashImage[i].pic_bmp)
		{
			if (g_SplashImage[i].pic_type == IMAGE_BITMAP)
				DeleteObject(g_SplashImage[i].pic_bmp);
			else
				DestroyIcon(g_SplashImage[i].pic_icon);
		}
		if (g_SplashImage[i].hwnd && IsWindow(g_SplashImage[i].hwnd))
			DestroyWindow(g_SplashImage[i].hwnd);
		if (g_SplashImage[i].hfont1) // Destroy font only after destroying the window that uses it.
			DeleteObject(g_SplashImage[i].hfont1);
		if (g_SplashImage[i].hfont2) // Destroy font only after destroying the window that uses it.
			DeleteObject(g_SplashImage[i].hfont2);
		if (g_SplashImage[i].hbrush)
			DeleteObject(g_SplashImage[i].hbrush);
	}

	// It is safer/easier to destroy the GUI windows prior to the menus (especially the menu bars).
	// This is because one GUI window might get destroyed and take with it a menu bar that is still
	// in use by an existing GUI window.  GuiType::Destroy() adheres to this philosophy by detaching
	// its menu bar prior to destroying its window:
	while (g_guiCount)
		GuiType::Destroy(*g_gui[g_guiCount-1]); // Static method to avoid problems with object destroying itself.
	for (i = 0; i < GuiType::sFontCount; ++i) // Now that GUI windows are gone, delete all GUI fonts.
		if (GuiType::sFont[i].hfont)
			DeleteObject(GuiType::sFont[i].hfont);
	// The above might attempt to delete an HFONT from GetStockObject(DEFAULT_GUI_FONT), etc.
	// But that should be harmless:
	// MSDN: "It is not necessary (but it is not harmful) to delete stock objects by calling DeleteObject."

	// Above: Probably best to have removed icon from tray and destroyed any Gui/Splash windows that were
	// using it prior to getting rid of the script's custom icon below:
	if (mCustomIcon)
	{
		DestroyIcon(mCustomIcon);
		DestroyIcon(mCustomIconSmall); // Should always be non-NULL if mCustomIcon is non-NULL.
	}

	// Since they're not associated with a window, we must free the resources for all popup menus.
	// Update: Even if a menu is being used as a GUI window's menu bar, see note above for why menu
	// destruction is done AFTER the GUI windows are destroyed:
	UserMenu *menu_to_delete;
	for (UserMenu *m = mFirstMenu; m;)
	{
		menu_to_delete = m;
		m = m->mNextMenu;
		ScriptDeleteMenu(menu_to_delete);
		// Above call should not return FAIL, since the only way FAIL can realistically happen is
		// when a GUI window is still using the menu as its menu bar.  But all GUI windows are gone now.
	}

	// Since tooltip windows are unowned, they should be destroyed to avoid resource leak:
	for (i = 0; i < MAX_TOOLTIPS; ++i)
		if (g_hWndToolTip[i] && IsWindow(g_hWndToolTip[i]))
			DestroyWindow(g_hWndToolTip[i]);

	if (g_hFontSplash) // The splash window itself should auto-destroyed, since it's owned by main.
		DeleteObject(g_hFontSplash);

	if (mOnClipboardChangeLabel || mOnClipboardChange.Count()) // Remove from viewer chain.
		EnableClipboardListener(false);

	// Close any open sound item to prevent hang-on-exit in certain operating systems or conditions.
	// If there's any chance that a sound was played and not closed out, or that it is still playing,
	// this check is done.  Otherwise, the check is avoided since it might be a high overhead call,
	// especially if the sound subsystem part of the OS is currently swapped out or something:
	if (g_SoundWasPlayed)
	{
		TCHAR buf[MAX_PATH * 2];
		mciSendString(_T("status ") SOUNDPLAY_ALIAS _T(" mode"), buf, _countof(buf), NULL);
		if (*buf) // "playing" or "stopped"
			mciSendString(_T("close ") SOUNDPLAY_ALIAS, NULL, 0, NULL);
	}

#ifdef ENABLE_KEY_HISTORY_FILE
	KeyHistoryToFile();  // Close the KeyHistory file if it's open.
#endif

	DeleteCriticalSection(&g_CriticalRegExCache); // g_CriticalRegExCache is used elsewhere for thread-safety.
	OleUninitialize();
}



ResultType Script::Init(global_struct &g, LPTSTR aScriptFilename, bool aIsRestart)
// Returns OK or FAIL.
// Caller has provided an empty string for aScriptFilename if this is a compiled script.
// Otherwise, aScriptFilename can be NULL if caller hasn't determined the filename of the script yet.
{
	mIsRestart = aIsRestart;
	TCHAR buf[2048]; // Just to make sure we have plenty of room to do things with.
#ifdef AUTOHOTKEYSC
	// Fix for v1.0.29: Override the caller's use of __argv[0] by using GetModuleFileName(),
	// so that when the script is started from the command line but the user didn't type the
	// extension, the extension will be included.  This necessary because otherwise
	// #SingleInstance wouldn't be able to detect duplicate versions in every case.
	// It also provides more consistency.
	GetModuleFileName(NULL, buf, _countof(buf));
#else
	TCHAR def_buf[MAX_PATH + 1], exe_buf[MAX_PATH + 20]; // For simplicity, allow at least space for +2 (see below) and "AutoHotkey.chm".
	if (!aScriptFilename) // v1.0.46.08: Change in policy: store the default script in the My Documents directory rather than in Program Files.  It's more correct and solves issues that occur due to Vista's file-protection scheme.
	{
		// Since no script-file was specified on the command line, use the default name.
		// For portability, first check if there's an <EXENAME>.ahk file in the current directory.
		LPTSTR suffix, dot;
		DWORD exe_len = GetModuleFileName(NULL, exe_buf, MAX_PATH + 2);
		// exe_len can be MAX_PATH+2 on Windows XP, in which case it is not null-terminated.
		// MAX_PATH+1 could mean it was truncated.  Any path longer than MAX_PATH would be rare.
		if (exe_len > MAX_PATH)
			return FAIL; // Seems the safest option for this unlikely case.
		if (  (suffix = _tcsrchr(exe_buf, '\\')) // Find name part of path.
			&& (dot = _tcsrchr(suffix, '.'))  ) // Find extension part of name.
			// Even if the extension is somehow zero characters, more than enough space was
			// reserved in exe_buf to add "ahk":
			//&& dot - exe_buf + 5 < _countof(exe_buf)  ) // Enough space in buffer?
		{
			_tcscpy(dot, EXT_AUTOHOTKEY);
		}
		else // Very unlikely.
			return FAIL;

		aScriptFilename = exe_buf; // Use the entire path, including the exe's directory.
		if (GetFileAttributes(aScriptFilename) == 0xFFFFFFFF) // File doesn't exist, so fall back to new method.
		{
			aScriptFilename = def_buf;
			VarSizeType filespec_length = BIV_MyDocuments(aScriptFilename, _T("")); // e.g. C:\Documents and Settings\Home\My Documents
			if (filespec_length + _tcslen(suffix) + 1 > _countof(def_buf))
				return FAIL; // Very rare, so for simplicity just abort.
			_tcscpy(aScriptFilename + filespec_length, suffix); // Append the filename: .ahk vs. .ini seems slightly better in terms of clarity and usefulness (e.g. the ability to double click the default script to launch it).
			if (GetFileAttributes(aScriptFilename) == 0xFFFFFFFF)
			{
				_tcscpy(suffix, _T("\\") AHK_HELP_FILE); // Replace the executable name.
				if (GetFileAttributes(exe_buf) != 0xFFFFFFFF) // Avoids hh.exe showing an error message if the file doesn't exist.
				{
					_sntprintf(buf, _countof(buf), _T("\"ms-its:%s::/docs/Welcome.htm\""), exe_buf);
					if (ActionExec(_T("hh.exe"), buf, exe_buf, false, _T("Max")))
						return FAIL;
				}
				// Since above didn't return, the help file is missing or failed to launch,
				// so continue on and let the missing script file be reported as an error.
			}
		}
		//else since the file exists, everything is now set up right. (The file might be a directory, but that isn't checked due to rarity.)
	}
	// In case the script is a relative filespec (relative to current working dir):
	if (!GetFullPathName(aScriptFilename, _countof(buf), buf, NULL)) // This is also relied upon by mIncludeLibraryFunctionsThenExit.  Succeeds even on nonexistent files.
		return FAIL; // Due to rarity, no error msg, just abort.
#endif
	if (g_RunStdIn = (*aScriptFilename == '*' && !aScriptFilename[1])) // v1.1.17: Read script from stdin.
	{
		// Seems best to disable #SingleInstance and enable #NoEnv for stdin scripts.
		g_AllowOnlyOneInstance = SINGLE_INSTANCE_OFF;
		g_NoEnv = true;
	}
	else // i.e. don't call the following function for stdin.
	// Using the correct case not only makes it look better in title bar & tray tool tip,
	// it also helps with the detection of "this script already running" since otherwise
	// it might not find the dupe if the same script name is launched with different
	// lowercase/uppercase letters:
	ConvertFilespecToCorrectCase(buf); // This might change the length, e.g. due to expansion of 8.3 filename.
	if (   !(mFileSpec = SimpleHeap::Malloc(buf))   )  // The full spec is stored for convenience, and it's relied upon by mIncludeLibraryFunctionsThenExit.
		return FAIL;  // It already displayed the error for us.
	LPTSTR filename_marker;
	if (filename_marker = _tcsrchr(buf, '\\'))
	{
		*filename_marker = '\0'; // Terminate buf in this position to divide the string.
		if (   !(mFileDir = SimpleHeap::Malloc(buf))   )
			return FAIL;  // It already displayed the error for us.
		++filename_marker;
	}
	else
	{
		// The only known cause of this condition is a path being too long for GetFullPathName
		// to expand it into buf, in which case buf and mFileSpec are now empty, and this will
		// cause LoadFromFile() to fail and the program to exit.
		//mFileDir = _T(""); // Already done by the constructor.
		filename_marker = buf;
	}
	if (   !(mFileName = SimpleHeap::Malloc(filename_marker))   )
		return FAIL;  // It already displayed the error for us.
#ifdef AUTOHOTKEYSC
	// Omit AutoHotkey from the window title, like AutoIt3 does for its compiled scripts.
	// One reason for this is to reduce backlash if evil-doers create viruses and such
	// with the program:
	sntprintf(buf, _countof(buf), _T("%s\\%s"), mFileDir, mFileName);
#else
	sntprintf(buf, _countof(buf), _T("%s\\%s - %s"), mFileDir, mFileName, T_AHK_NAME_VERSION);
#endif
	if (   !(mMainWindowTitle = SimpleHeap::Malloc(buf))   )
		return FAIL;  // It already displayed the error for us.

	// It may be better to get the module name this way rather than reading it from the registry
	// (though it might be more proper to parse it out of the command line args or something),
	// in case the user has moved it to a folder other than the install folder, hasn't installed it,
	// or has renamed the EXE file itself.  Also, enclose the full filespec of the module in double
	// quotes since that's how callers usually want it because ActionExec() currently needs it that way:
	*buf = '"';
	if (GetModuleFileName(NULL, buf + 1, _countof(buf) - 2)) // -2 to leave room for the enclosing double quotes.
	{
		size_t buf_length = _tcslen(buf);
		buf[buf_length++] = '"';
		buf[buf_length] = '\0';
		if (   !(mOurEXE = SimpleHeap::Malloc(buf))   )
			return FAIL;  // It already displayed the error for us.
		else
		{
			LPTSTR last_backslash = _tcsrchr(buf, '\\');
			if (!last_backslash) // probably can't happen due to the nature of GetModuleFileName().
				mOurEXEDir = _T("");
			*last_backslash = '\0';
			if (   !(mOurEXEDir = SimpleHeap::Malloc(buf + 1))   ) // +1 to omit the leading double-quote.
				return FAIL;  // It already displayed the error for us.
		}
	}
	return OK;
}

	

ResultType Script::CreateWindows()
// Returns OK or FAIL.
{
	if (!mMainWindowTitle || !*mMainWindowTitle) return FAIL;  // Init() must be called before this function.
	// Register a window class for the main window:
	WNDCLASSEX wc = {0};
	wc.cbSize = sizeof(wc);
	wc.lpszClassName = WINDOW_CLASS_MAIN;
	wc.hInstance = g_hInstance;
	wc.lpfnWndProc = MainWindowProc;
	// The following are left at the default of NULL/0 set higher above:
	//wc.style = 0;  // CS_HREDRAW | CS_VREDRAW
	//wc.cbClsExtra = 0;
	//wc.cbWndExtra = 0;
	// Load the main icon in the two sizes needed throughout the program:
	g_IconLarge = ExtractIconFromExecutable(NULL, -IDI_MAIN, 0, 0);
	g_IconSmall = ExtractIconFromExecutable(NULL, -IDI_MAIN, GetSystemMetrics(SM_CXSMICON), 0);
	wc.hIcon = g_IconLarge;
	wc.hIconSm = g_IconSmall;
	wc.hCursor = LoadCursor((HINSTANCE) NULL, IDC_ARROW);
	wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);  // Needed for ProgressBar. Old: (HBRUSH)GetStockObject(WHITE_BRUSH);
	wc.lpszMenuName = MAKEINTRESOURCE(IDR_MENU_MAIN); // NULL; // "MainMenu";
	if (!RegisterClassEx(&wc))
	{
		MsgBox(_T("RegClass")); // Short/generic msg since so rare.
		return FAIL;
	}

	// Register a second class for the splash window.  The only difference is that
	// it doesn't have the menu bar:
	wc.lpszClassName = WINDOW_CLASS_SPLASH;
	wc.lpszMenuName = NULL; // Override the non-NULL value set higher above.
	if (!RegisterClassEx(&wc))
	{
		MsgBox(_T("RegClass")); // Short/generic msg since so rare.
		return FAIL;
	}

	TCHAR class_name[64];
	HWND fore_win = GetForegroundWindow();
	bool do_minimize = !fore_win || (GetClassName(fore_win, class_name, _countof(class_name))
		&& !_tcsicmp(class_name, _T("Shell_TrayWnd"))); // Shell_TrayWnd is the taskbar's class on Win98/XP and probably the others too.

	// Note: the title below must be constructed the same was as is done by our
	// WinMain() (so that we can detect whether this script is already running)
	// which is why it's standardized in g_script.mMainWindowTitle.
	// Create the main window.  Prevent momentary disruption of Start Menu, which
	// some users understandably don't like, by omitting the taskbar button temporarily.
	// This is done because testing shows that minimizing the window further below, even
	// though the window is hidden, would otherwise briefly show the taskbar button (or
	// at least redraw the taskbar).  Sometimes this isn't noticeable, but other times
	// (such as when the system is under heavy load) a user reported that it is quite
	// noticeable. WS_EX_TOOLWINDOW is used instead of WS_EX_NOACTIVATE because
	// WS_EX_NOACTIVATE is available only on 2000/XP.
	if (   !(g_hWnd = CreateWindowEx(do_minimize ? WS_EX_TOOLWINDOW : 0
		, WINDOW_CLASS_MAIN
		, mMainWindowTitle
		, WS_OVERLAPPEDWINDOW // Style.  Alt: WS_POPUP or maybe 0.
		, CW_USEDEFAULT // xpos
		, CW_USEDEFAULT // ypos
		, CW_USEDEFAULT // width
		, CW_USEDEFAULT // height
		, NULL // parent window
		, NULL // Identifies a menu, or specifies a child-window identifier depending on the window style
		, g_hInstance // passed into WinMain
		, NULL))   ) // lpParam
	{
		MsgBox(_T("CreateWindow")); // Short msg since so rare.
		return FAIL;
	}
#ifdef AUTOHOTKEYSC
	HMENU menu = GetMenu(g_hWnd);
	// Disable the Edit menu item, since it does nothing for a compiled script:
	EnableMenuItem(menu, ID_FILE_EDITSCRIPT, MF_DISABLED | MF_GRAYED);
	EnableOrDisableViewMenuItems(menu, MF_DISABLED | MF_GRAYED); // Fix for v1.0.47.06: No point in checking g_AllowMainWindow because the script hasn't starting running yet, so it will always be false.
	// But leave the ID_VIEW_REFRESH menu item enabled because if the script contains a
	// command such as ListLines in it, Refresh can be validly used.
#endif

	if (    !(g_hWndEdit = CreateWindow(_T("edit"), NULL, WS_CHILD | WS_VISIBLE | WS_BORDER
		| ES_LEFT | ES_MULTILINE | ES_READONLY | WS_VSCROLL // | WS_HSCROLL (saves space)
		, 0, 0, 0, 0, g_hWnd, (HMENU)1, g_hInstance, NULL))   )
	{
		MsgBox(_T("CreateWindow")); // Short msg since so rare.
		return FAIL;
	}
	// FONTS: The font used by default, at least on XP, is GetStockObject(SYSTEM_FONT).
	// It seems preferable to smaller fonts such DEFAULT_GUI_FONT(DEFAULT_GUI_FONT).
	// For more info on pre-loaded fonts (not too many choices), see MSDN's GetStockObject().
	if(g_os.IsWinNT())
	{
		// Use a more appealing font on NT versions of Windows.

		// Windows NT to Windows XP -> Lucida Console
		HDC hdc = GetDC(g_hWndEdit);
		if(!g_os.IsWinVistaOrLater())
			g_hFontEdit = CreateFont(FONT_POINT(hdc, 10), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
				, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, _T("Lucida Console"));
		else // Windows Vista and later -> Consolas
			g_hFontEdit = CreateFont(FONT_POINT(hdc, 10), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
				, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, _T("Consolas"));
		ReleaseDC(g_hWndEdit, hdc); // In theory it must be done.
		SendMessage(g_hWndEdit, WM_SETFONT, (WPARAM)g_hFontEdit, 0);
	}

	// v1.0.30.05: Specifying a limit of zero opens the control to its maximum text capacity,
	// which removes the 32K size restriction.  Testing shows that this does not increase the actual
	// amount of memory used for controls containing small amounts of text.  All it does is allow
	// the control to allocate more memory as needed.  By specifying zero, a max
	// of 64K becomes available on Windows 9x, and perhaps as much as 4 GB on NT/2k/XP.
	SendMessage(g_hWndEdit, EM_LIMITTEXT, 0, 0);

	// Some of the MSDN docs mention that an app's very first call to ShowWindow() makes that
	// function operate in a special mode. Therefore, it seems best to get that first call out
	// of the way to avoid the possibility that the first-call behavior will cause problems with
	// our normal use of ShowWindow() below and other places.  Also, decided to ignore nCmdShow,
    // to avoid any momentary visual effects on startup.
	// Update: It's done a second time because the main window might now be visible if the process
	// that launched ours specified that.  It seems best to override the requested state because
	// some calling processes might specify "maximize" or "shownormal" as generic launch method.
	// The script can display it's own main window with ListLines, etc.
	// MSDN: "the nCmdShow value is ignored in the first call to ShowWindow if the program that
	// launched the application specifies startup information in the structure. In this case,
	// ShowWindow uses the information specified in the STARTUPINFO structure to show the window.
	// On subsequent calls, the application must call ShowWindow with nCmdShow set to SW_SHOWDEFAULT
	// to use the startup information provided by the program that launched the application."
	ShowWindow(g_hWnd, SW_HIDE);
	ShowWindow(g_hWnd, SW_HIDE);

	// Now that the first call to ShowWindow() is out of the way, minimize the main window so that
	// if the script is launched from the Start Menu (and perhaps other places such as the
	// Quick-launch toolbar), the window that was active before the Start Menu was displayed will
	// become active again.  But as of v1.0.25.09, this minimize is done more selectively to prevent
	// the launch of a script from knocking the user out of a full-screen game or other application
	// that would be disrupted by an SW_MINIMIZE:
	if (do_minimize)
	{
		ShowWindow(g_hWnd, SW_MINIMIZE);
		SetWindowLong(g_hWnd, GWL_EXSTYLE, 0); // Give the main window back its taskbar button.
	}
	// Note: When the window is not minimized, task manager reports that a simple script (such as
	// one consisting only of the single line "#Persistent") uses 2600 KB of memory vs. ~452 KB if
	// it were immediately minimized.  That is probably just due to the vagaries of how the OS
	// manages windows and memory and probably doesn't actually impact system performance to the
	// degree indicated.  In other words, it's hard to imagine that the failure to do
	// ShowWidnow(g_hWnd, SW_MINIMIZE) unconditionally upon startup (which causes the side effects
	// discussed further above) significantly increases the actual memory load on the system.

	g_hAccelTable = LoadAccelerators(g_hInstance, MAKEINTRESOURCE(IDR_ACCELERATOR1));

	if (g_NoTrayIcon)
		mNIC.hWnd = NULL;  // Set this as an indicator that tray icon is not installed.
	else
		// Even if the below fails, don't return FAIL in case the user is using a different shell
		// or something.  In other words, it is expected to fail under certain circumstances and
		// we want to tolerate that:
		CreateTrayIcon();

	if (mOnClipboardChangeLabel)
		EnableClipboardListener(true);

	return OK;
}



void Script::EnableClipboardListener(bool aEnable)
{
	static bool sEnabled = false;
	if (aEnable == sEnabled) // Simplifies BIF_On.
		return;
	if (aEnable)
	{
		if (MyAddClipboardListener && MyRemoveClipboardListener) // Should be impossible for only one of these to be NULL, but check both anyway to be safe.
		{
			// The old clipboard viewer chain method is prone to break when some other application uses
			// it incorrectly.  This newer method should be more reliable, but requires Vista or later:
			MyAddClipboardListener(g_hWnd);
			// But this method doesn't appear to send an initial WM_CLIPBOARDUPDATE message.
			// For consistency with the other method (below) and for backward compatibility,
			// run the OnClipboardChange label once when the script first starts:
			if (!mIsReadyToExecute)
			{
				// Pass 1 for wParam so that MsgSleep() will call only the legacy OnClipboardChange label,
				// not any functions which are registered between now and when the message is handled.
				PostMessage(g_hWnd, AHK_CLIPBOARD_CHANGE, 1, 0);
			}
		}
		else
		{
			mNextClipboardViewer = SetClipboardViewer(g_hWnd);
			// SetClipboardViewer() sends a WM_DRAWCLIPBOARD message and causes MainWindowProc()
			// to be called before returning.  MainWindowProc() posts an AHK_CLIPBOARD_CHANGE
			// message only if an OnClipboardChange label exists, since mOnClipboardChange.Count()
			// is always 0 at this point.  It also uses wParam for the reason described above.
		}
	}
	else
	{
		if (MyRemoveClipboardListener && MyAddClipboardListener)
			MyRemoveClipboardListener(g_hWnd); // MyAddClipboardListener was used.
		else
			ChangeClipboardChain(g_hWnd, mNextClipboardViewer); // SetClipboardViewer was used.
	}
	sEnabled = aEnable;
}



void Script::EnableOrDisableViewMenuItems(HMENU aMenu, UINT aFlags)
{
	EnableMenuItem(aMenu, ID_VIEW_KEYHISTORY, aFlags);
	EnableMenuItem(aMenu, ID_VIEW_LINES, aFlags);
	EnableMenuItem(aMenu, ID_VIEW_VARIABLES, aFlags);
	EnableMenuItem(aMenu, ID_VIEW_HOTKEYS, aFlags);
}



void Script::CreateTrayIcon()
// It is the caller's responsibility to ensure that the previous icon is first freed/destroyed
// before calling us to install a new one.  However, that is probably not needed if the Explorer
// crashed, since the memory used by the tray icon was probably destroyed along with it.
{
	ZeroMemory(&mNIC, sizeof(mNIC));  // To be safe.
	// Using NOTIFYICONDATA_V2_SIZE vs. sizeof(NOTIFYICONDATA) improves compatibility with Win9x maybe.
	// MSDN: "Using [NOTIFYICONDATA_V2_SIZE] for cbSize will allow your application to use NOTIFYICONDATA
	// with earlier Shell32.dll versions, although without the version 6.0 enhancements."
	// Update: Using V2 gives an compile error so trying V1.  Update: Trying sizeof(NOTIFYICONDATA)
	// for compatibility with VC++ 6.x.  This is also what AutoIt3 uses:
	mNIC.cbSize = sizeof(NOTIFYICONDATA);  // NOTIFYICONDATA_V1_SIZE
	mNIC.hWnd = g_hWnd;
	mNIC.uID = AHK_NOTIFYICON; // This is also used for the ID, see TRANSLATE_AHK_MSG for details.
	mNIC.uFlags = NIF_MESSAGE | NIF_TIP | NIF_ICON;
	mNIC.uCallbackMessage = AHK_NOTIFYICON;
	mNIC.hIcon = mCustomIconSmall ? mCustomIconSmall : g_IconSmall;
	UPDATE_TIP_FIELD
	// If we were called due to an Explorer crash, I don't think it's necessary to call
	// Shell_NotifyIcon() to remove the old tray icon because it was likely destroyed
	// along with Explorer.  So just add it unconditionally:
	if (!Shell_NotifyIcon(NIM_ADD, &mNIC))
		mNIC.hWnd = NULL;  // Set this as an indicator that tray icon is not installed.
}



void Script::UpdateTrayIcon(bool aForceUpdate)
{
	if (!mNIC.hWnd) // tray icon is not installed
		return;
	static bool icon_shows_paused = false;
	static bool icon_shows_suspended = false;
	if (!aForceUpdate && (mIconFrozen || (g->IsPaused == icon_shows_paused && g_IsSuspended == icon_shows_suspended)))
		return; // it's already in the right state
	int icon;
	if (g->IsPaused && g_IsSuspended)
		icon = IDI_PAUSE_SUSPEND;
	else if (g->IsPaused)
		icon = IDI_PAUSE;
	else if (g_IsSuspended)
		icon = IDI_SUSPEND;
	else
		icon = IDI_MAIN;
	// Use the custom tray icon if the icon is normal (non-paused & non-suspended):
	mNIC.hIcon = (mCustomIconSmall && (mIconFrozen || (!g->IsPaused && !g_IsSuspended))) ? mCustomIconSmall // L17: Always use small icon for tray.
		: (icon == IDI_MAIN) ? g_IconSmall // Use the pre-loaded small icon for best quality.
		: (HICON)LoadImage(g_hInstance, MAKEINTRESOURCE(icon), IMAGE_ICON, 0, 0, LR_SHARED); // Use LR_SHARED for simplicity and performance more than to conserve memory in this case.
	if (Shell_NotifyIcon(NIM_MODIFY, &mNIC))
	{
		icon_shows_paused = g->IsPaused;
		icon_shows_suspended = g_IsSuspended;
	}
	// else do nothing, just leave it in the same state.
}



ResultType Script::AutoExecSection()
// Returns FAIL if can't run due to critical error.  Otherwise returns OK.
{
	// Now that g_MaxThreadsTotal has been permanently set by the processing of script directives like
	// #MaxThreads, an appropriately sized array can be allocated:
	if (   !(g_array = (global_struct *)malloc((g_MaxThreadsTotal+TOTAL_ADDITIONAL_THREADS) * sizeof(global_struct)))   )
		return FAIL; // Due to rarity, just abort. It wouldn't be safe to run ExitApp() due to possibility of an OnExit routine.
	CopyMemory(g_array, g, sizeof(global_struct)); // Copy the temporary/startup "g" into array[0] to preserve historical behaviors that may rely on the idle thread starting with that "g".
	g = g_array; // Must be done after above.

	// v1.0.48: Due to switching from SET_UNINTERRUPTIBLE_TIMER to IsInterruptible():
	// In spite of the comments in IsInterruptible(), periodically have a timer call IsInterruptible() due to
	// the following scenario:
	// - Interrupt timeout is 60 seconds (or 60 milliseconds for that matter).
	// - For some reason IsInterrupt() isn't called for 24+ hours even though there is a current/active thread.
	// - RefreshInterruptibility() fires at 23 hours and marks the thread interruptible.
	// - Sometime after that, one of the following happens:
	//      Computer is suspended/hibernated and stays that way for 50+ days.
	//      IsInterrupt() is never called (except by RefreshInterruptibility()) for 50+ days.
	//      (above is currently unlikely because MSG_FILTER_MAX calls IsInterruptible())
	// In either case, RefreshInterruptibility() has prevented the uninterruptibility duration from being
	// wrongly extended by up to 100% of g_script.mUninterruptibleTime.  This isn't a big deal if
	// g_script.mUninterruptibleTime is low (like it almost always is); but if it's fairly large, say an hour,
	// this can prevent an unwanted extension of up to 1 hour.
	// Although any call frequency less than 49.7 days should work, currently calling once per 23 hours
	// in case any older operating systems have a SetTimer() limit of less than 0x7FFFFFFF (and also to make
	// it less likely that a long suspend/hibernate would cause the above issue).  The following was
	// actually tested on Windows XP and a message does indeed arrive 23 hours after the script starts.
	SetTimer(g_hWnd, TIMER_ID_REFRESH_INTERRUPTIBILITY, 23*60*60*1000, RefreshInterruptibility); // 3rd param must not exceed 0x7FFFFFFF (2147483647; 24.8 days).

	ResultType ExecUntil_result;

	if (!mFirstLine) // In case it's ever possible to be empty.
		ExecUntil_result = OK;
		// And continue on to do normal exit routine so that the right ExitCode is returned by the program.
	else
	{
		// Choose a timeout that's a reasonable compromise between the following competing priorities:
		// 1) That we want hotkeys to be responsive as soon as possible after the program launches
		//    in case the user launches by pressing ENTER on a script, for example, and then immediately
		//    tries to use a hotkey.  In addition, we want any timed subroutines to start running ASAP
		//    because in rare cases the user might rely upon that happening.
		// 2) To support the case when the auto-execute section never finishes (such as when it contains
		//    an infinite loop to do background processing), yet we still want to allow the script
		//    to put custom defaults into effect globally (for things such as KeyDelay).
		// Obviously, the above approach has its flaws; there are ways to construct a script that would
		// result in unexpected behavior.  However, the combination of this approach with the fact that
		// the global defaults are updated *again* when/if the auto-execute section finally completes
		// raises the expectation of proper behavior to a very high level.  In any case, I'm not sure there
		// is any better approach that wouldn't break existing scripts or require a redesign of some kind.
		// If this method proves unreliable due to disk activity slowing the program down to a crawl during
		// the critical milliseconds after launch, one thing that might fix that is to have ExecUntil()
		// be forced to run a minimum of, say, 100 lines (if there are that many) before allowing the
		// timer expiration to have its effect.  But that's getting complicated and I'd rather not do it
		// unless someone actually reports that such a thing ever happens.  Still, to reduce the chance
		// of such a thing ever happening, it seems best to boost the timeout from 50 up to 100:
		SET_AUTOEXEC_TIMER(100);
		mAutoExecSectionIsRunning = true;

		// v1.0.25: This is now done here, closer to the actual execution of the first line in the script,
		// to avoid an unnecessary Sleep(10) that would otherwise occur in ExecUntil:
		mLastScriptRest = mLastPeekTime = GetTickCount();

		++g_nThreads;
		DEBUGGER_STACK_PUSH(_T("Auto-execute"))
		ExecUntil_result = mFirstLine->ExecUntil(UNTIL_RETURN); // Might never return (e.g. infinite loop or ExitApp).
		DEBUGGER_STACK_POP()
		--g_nThreads;
		// Our caller will take care of setting g_default properly.

		KILL_AUTOEXEC_TIMER // See also: AutoExecSectionTimeout().
		mAutoExecSectionIsRunning = false;
	}
	// REMEMBER: The ExecUntil() call above will never return if the AutoExec section never finishes
	// (e.g. infinite loop) or it uses Exit/ExitApp.

	// Check if an exception has been thrown
	if (g->ThrownToken)
		g_script.FreeExceptionToken(g->ThrownToken);

	// The below is done even if AutoExecSectionTimeout() already set the values once.
	// This is because when the AutoExecute section finally does finish, by definition it's
	// supposed to store the global settings that are currently in effect as the default values.
	// In other words, the only purpose of AutoExecSectionTimeout() is to handle cases where
	// the AutoExecute section takes a long time to complete, or never completes (perhaps because
	// it is being used by the script as a "background thread" of sorts):
	// Save the values of KeyDelay, WinDelay etc. in case they were changed by the auto-execute part
	// of the script.  These new defaults will be put into effect whenever a new hotkey subroutine
	// is launched.  Each launched subroutine may then change the values for its own purposes without
	// affecting the settings for other subroutines:
	global_clear_state(*g);  // Start with a "clean slate" in both g_default and g (in case things like InitNewThread() check some of the values in g prior to launching a new thread).

	// Always want g_default.AllowInterruption==true so that InitNewThread()  doesn't have to
	// set it except when Critical or "Thread Interrupt" require it.  If the auto-execute section ended
	// without anyone needing to call IsInterruptible() on it, AllowInterruption could be false
	// even when Critical is off.
	// Even if the still-running AutoExec section has turned on Critical, the assignment below is still okay
	// because InitNewThread() adjusts AllowInterruption based on the value of ThreadIsCritical.
	// See similar code in AutoExecSectionTimeout().
	g->AllowThreadToBeInterrupted = true; // Mostly for the g_default line below. See comments above.
	CopyMemory(&g_default, g, sizeof(global_struct)); // g->IsPaused has been set to false higher above in case it's ever possible that it's true as a result of AutoExecSection().
	// After this point, the values in g_default should never be changed.
	global_maximize_interruptibility(*g); // See below.
	// Now that any changes made by the AutoExec section have been saved to g_default (including
	// the commands Critical and Thread), ensure that the very first g-item is always interruptible.
	// This avoids having to treat the first g-item as special in various places.

	// It seems best to set ErrorLevel to NONE after the auto-execute part of the script is done.
	// However, it isn't set to NONE right before launching each new thread (e.g. hotkey subroutine)
	// because it's more flexible that way (i.e. the user may want one hotkey subroutine to use the value
	// of ErrorLevel set by another).  This reset was also done by LoadFromFile(), but it is done again
	// here in case the auto-execute section changed it:
	g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	// BEFORE DOING THE BELOW, "g" and "g_default" should be set up properly in case there's an OnExit
	// routine (even non-persistent scripts can have one).
	// If no hotkeys are in effect, the user hasn't requested a hook to be activated, and the script
	// doesn't contain the #Persistent directive we're done unless there is an OnExit subroutine and it
	// doesn't do "ExitApp":
	if (!IS_PERSISTENT) // Resolve macro again in case any of its components changed since the last time.
		g_script.ExitApp(ExecUntil_result == FAIL ? EXIT_ERROR : EXIT_EXIT);

	return OK;
}



ResultType Script::Edit()
{
#ifdef AUTOHOTKEYSC
	return OK; // Do nothing.
#else
	// This is here in case a compiled script ever uses the Edit command.  Since the "Edit This
	// Script" menu item is not available for compiled scripts, it can't be called from there.
	TitleMatchModes old_mode = g->TitleMatchMode;
	g->TitleMatchMode = FIND_ANYWHERE;
	HWND hwnd = WinExist(*g, mFileName, _T(""), mMainWindowTitle, _T("")); // Exclude our own main window.
	g->TitleMatchMode = old_mode;
	if (hwnd)
	{
		TCHAR class_name[32];
		GetClassName(hwnd, class_name, _countof(class_name));
		if (!_tcscmp(class_name, _T("#32770")) || !_tcsnicmp(class_name, _T("AutoHotkey"), 10)) // MessageBox(), InputBox(), FileSelectFile(), or GUI/script-owned window.
			hwnd = NULL;  // Exclude it from consideration.
	}
	if (hwnd)  // File appears to already be open for editing, so use the current window.
		SetForegroundWindowEx(hwnd);
	else
	{
		TCHAR buf[MAX_PATH * 2];
		// Enclose in double quotes anything that might contain spaces since the CreateProcess()
		// method, which is attempted first, is more likely to succeed.  This is because it uses
		// the command line method of creating the process, with everything all lumped together:
		sntprintf(buf, _countof(buf), _T("\"%s\""), mFileSpec);
		if (!ActionExec(_T("edit"), buf, mFileDir, false))  // Since this didn't work, try notepad.
		{
			// v1.0.40.06: Try to open .ini files first with their associated editor rather than trying the
			// "edit" verb on them:
			LPTSTR file_ext;
			if (   !(file_ext = _tcsrchr(mFileName, '.')) || _tcsicmp(file_ext, _T(".ini"))
				|| !ActionExec(_T("open"), buf, mFileDir, false)   ) // Relies on short-circuit boolean order.
			{
				// Even though notepad properly handles filenames with spaces in them under WinXP,
				// even without double quotes around them, it seems safer and more correct to always
				// enclose the filename in double quotes for maximum compatibility with all OSes:
				if (!ActionExec(_T("notepad.exe"), buf, mFileDir, false))
					MsgBox(_T("Could not open script.")); // Short message since so rare.
			}
		}
	}
	return OK;
#endif
}



ResultType Script::Reload(bool aDisplayErrors)
{
	// The new instance we're about to start will tell our process to stop, or it will display
	// a syntax error or some other error, in which case our process will still be running:
#ifdef AUTOHOTKEYSC
	// This is here in case a compiled script ever uses the Reload command.  Since the "Reload This
	// Script" menu item is not available for compiled scripts, it can't be called from there.
	return g_script.ActionExec(mOurEXE, _T("/restart"), g_WorkingDirOrig, aDisplayErrors);
#else
	TCHAR arg_string[MAX_PATH + 512];
	sntprintf(arg_string, _countof(arg_string), _T("/restart \"%s\""), mFileSpec);
	return g_script.ActionExec(mOurEXE, arg_string, g_WorkingDirOrig, aDisplayErrors);
#endif
}



ResultType Script::ExitApp(ExitReasons aExitReason, int aExitCode)
// Normal exit (if aBuf is NULL), or a way to exit immediately on error (which is mostly
// for times when it would be unsafe to call MsgBox() due to the possibility that it would
// make the situation even worse).
{
	mExitReason = aExitReason;
	static bool sOnExitIsRunning = false, sExitAppShouldTerminate = true;
	static int sExitCode;
	if (sOnExitIsRunning || !mIsReadyToExecute)
	{
		// There is another instance of this function beneath us on the stack, executing an
		// OnExit subroutine or function.  If a legacy OnExit sub is running, we still need
		// to execute any other OnExit callbacks before exiting.  Otherwise ExitApp should
		// terminate the app.  Causes of script exit other than ExitApp are expected to
		// terminate the app immediately even if OnExit is running.
		if (sExitAppShouldTerminate || aExitReason != EXIT_EXIT)
			TerminateApp(aExitReason, aExitCode); // Exit early; don't run the OnExit callbacks (again).
		if (*Line::sArgDeref[0]) // ExitApp with a parameter -- relies on the aExitReason check above.
			sExitCode = aExitCode; // Override the previous exit code.
		sExitAppShouldTerminate = true; // Signal our other instance that ExitApp was called.
		return EARLY_EXIT; // Exit the thread (our other instance will call TerminateApp).
		// MUST NOT create a new thread when sOnExitIsRunning because g_array allows only one
		// extra thread for ExitApp() (which allows it to run even when MAX_THREADS_EMERGENCY
		// has been reached).  See TOTAL_ADDITIONAL_THREADS for details.
	}
	sExitCode = aExitCode;

	// Otherwise, the script contains the special RunOnExit label that we will run here instead
	// of exiting.  And since it does, we know that the script is in a ready-to-execute state
	// because that is the only way an OnExit label could have been defined in the first place.
	// Usually, the RunOnExit subroutine will contain an Exit or ExitApp statement
	// which results in a recursive call to this function, but this is not required (e.g. the
	// Exit subroutine could display an "Are you sure?" prompt, and if the user chooses "No",
	// the Exit sequence can be aborted by simply not calling ExitApp and letting the thread
	// we create below end normally).

	// Next, save the current state of the globals so that they can be restored just prior
	// to returning to our caller:
	TCHAR ErrorLevel_saved[ERRORLEVEL_SAVED_SIZE];
	tcslcpy(ErrorLevel_saved, g_ErrorLevel->Contents(), _countof(ErrorLevel_saved)); // Save caller's errorlevel.
	InitNewThread(0, true, true, ACT_INVALID); // Uninterruptibility is handled below. Since this special thread should always run, no checking of g_MaxThreadsTotal is done before calling this.

	// Turn on uninterruptibility to forbid any hotkeys, timers, or user defined menu items
	// to interrupt.  This is mainly done for peace-of-mind (since possible interactions due to
	// interruptions have not been studied) and the fact that this most users would not want this
	// subroutine to be interruptible (it usually runs quickly anyway).  Another reason to make
	// it non-interruptible is that some OnExit subroutines might destruct things used by the
	// script's hotkeys/timers/menu items, and activating these items during the deconstruction
	// would not be safe.  Finally, if a logoff or shutdown is occurring, it seems best to prevent
	// timed subroutines from running -- which might take too much time and prevent the exit from
	// occurring in a timely fashion.  An option can be added via the FutureUse param to make it
	// interruptible if there is ever a demand for that.
	// UPDATE: g_AllowInterruption is now used instead of g->AllowThreadToBeInterrupted for two reasons:
	// 1) It avoids the need to do "int mUninterruptedLineCountMax_prev = g_script.mUninterruptedLineCountMax;"
	//    (Disable this item so that ExecUntil() won't automatically make our new thread uninterruptible
	//    after it has executed a certain number of lines).
	// 2) Mostly obsolete: If the thread we're interrupting is uninterruptible, the uninterruptible timer
	//    might be currently pending.  When it fires, it would make the OnExit subroutine interruptible
	//    rather than the underlying subroutine.  The above fixes the first part of that problem.
	//    The 2nd part is fixed by reinstating the timer when the uninterruptible thread is resumed.
	//    This special handling is only necessary here -- not in other places where new threads are
	//    created -- because OnExit is the only type of thread that can interrupt an uninterruptible
	//    thread.
	BOOL g_AllowInterruption_prev = g_AllowInterruption;  // Save current setting.
	g_AllowInterruption = FALSE; // Mark the thread just created above as permanently uninterruptible (i.e. until it finishes and is destroyed).

	sOnExitIsRunning = true;
	DEBUGGER_STACK_PUSH(_T("OnExit"))

	bool terminate_afterward = true; // Set default - see below for comments.
	
	// When a legacy OnExit label is present, the default behaviour is to exit the script only if
	// it calls ExitApp.  Therefore to make OnExit() useful in a script which uses legacy OnExit,
	// we need to prevent ExitApp from actually terminating the app:
	sExitAppShouldTerminate = false;
	// If the subroutine encounters a failure condition such as a runtime error, exit afterward.
	// Otherwise, there will be no way to exit the script if the subroutine fails on each attempt.
	if (mOnExitLabel && mOnExitLabel->Execute() && !sExitAppShouldTerminate)
	{
		// The subroutine completed normally and did not call ExitApp, so don't exit.
		terminate_afterward = false;
	}
	sExitAppShouldTerminate = true;

	// If an OnExit label was called and didn't call ExitApp, terminate_afterward was set to false,
	// so the script isn't exiting.  Otherwise, call the chain of OnExit functions:
	if (terminate_afterward)
	{
		ExprTokenType param[] = { GetExitReasonString(aExitReason), (__int64)sExitCode };
		if (mOnExit.Call(param, _countof(param), mOnExitLabel ? 0 : 1) == CONDITION_TRUE)
			terminate_afterward = false; // A handler returned true to prevent exit.
	}
	
	DEBUGGER_STACK_POP()
	sOnExitIsRunning = false;  // In case the user wanted the thread to end normally (see above).

	if (terminate_afterward || aExitReason == EXIT_DESTROY)
		TerminateApp(aExitReason, aExitCode);

	// Otherwise:
	ResumeUnderlyingThread(ErrorLevel_saved);
	g_AllowInterruption = g_AllowInterruption_prev;  // Restore original setting.

	return EARLY_EXIT;
}



void Script::TerminateApp(ExitReasons aExitReason, int aExitCode)
// Note that g_script's destructor takes care of most other cleanup work, such as destroying
// tray icons, menus, and unowned windows such as ToolTip.
{
	// L31: Release objects stored in variables, where possible.
	if (aExitReason != CRITICAL_ERROR) // i.e. Avoid making matters worse if CRITICAL_ERROR.
	{
		// Ensure the current thread is not paused and can't be interrupted
		// in case one or more objects need to call a __delete meta-function.
		g_AllowInterruption = FALSE;
		g->IsPaused = false;

		int v, i;
		for (v = 0; v < mVarCount; ++v)
			if (mVar[v]->IsObject())
				mVar[v]->ReleaseObject();
		for (v = 0; v < mLazyVarCount; ++v)
			if (mLazyVar[v]->IsObject())
				mLazyVar[v]->ReleaseObject();
		for (i = 0; i < mFuncCount; ++i)
		{
			Func &f = *mFunc[i];
			if (f.mIsBuiltIn)
				continue;
			// Since it doesn't seem feasible to release all var backups created by recursive function
			// calls and all tokens in the 'stack' of each currently executing expression, currently
			// only static and global variables are released.  It seems best for consistency to also
			// avoid releasing top-level non-static local variables (i.e. which aren't in var backups).
			for (v = 0; v < f.mVarCount; ++v)
				if (f.mVar[v]->IsStatic() && f.mVar[v]->IsObject()) // For consistency, only free static vars (see above).
					f.mVar[v]->ReleaseObject();
			for (v = 0; v < f.mLazyVarCount; ++v)
				if (f.mLazyVar[v]->IsStatic() && f.mLazyVar[v]->IsObject())
					f.mLazyVar[v]->ReleaseObject();
		}
	}
#ifdef CONFIG_DEBUGGER // L34: Exit debugger *after* the above to allow debugging of any invoked __Delete handlers.
	g_Debugger.Exit(aExitReason);
#endif

	// We call DestroyWindow() because MainWindowProc() has left that up to us.
	// DestroyWindow() will cause MainWindowProc() to immediately receive and process the
	// WM_DESTROY msg, which should in turn result in any child windows being destroyed
	// and other cleanup being done:
	if (IsWindow(g_hWnd)) // Adds peace of mind in case WM_DESTROY was already received in some unusual way.
	{
		g_DestroyWindowCalled = true;
		DestroyWindow(g_hWnd);
	}
	Hotkey::AllDestructAndExit(aExitCode);
}



UINT Script::LoadFromFile()
// Returns the number of non-comment lines that were loaded, or LOADING_FAILED on error.
{
	mNoHotkeyLabels = true;  // Indicate that there are no hotkey labels, since we're (re)loading the entire file.
	mIsReadyToExecute = mAutoExecSectionIsRunning = false;
	if (!mFileSpec || !*mFileSpec) return LOADING_FAILED;

#ifndef AUTOHOTKEYSC  // When not in stand-alone mode, read an external script file.
	DWORD attr = g_RunStdIn ? 0 : GetFileAttributes(mFileSpec); // v1.1.17: Don't check if reading script from stdin.
	if (attr == MAXDWORD) // File does not exist or lacking the authorization to get its attributes.
	{
		TCHAR buf[MAX_PATH + 256];
		sntprintf(buf, _countof(buf), _T("Script file not found:\n%s"), mFileSpec);
		MsgBox(buf, MB_ICONHAND);
		return 0;
	}
#endif

	// v1.0.42: Placeholder to use in place of a NULL label to simplify code in some places.
	// This must be created before loading the script because it's relied upon when creating
	// hotkeys to provide an alternative to having a NULL label. It will be given a non-NULL
	// mJumpToLine further down.
	if (   !(mPlaceholderLabel = new Label(_T("")))   ) // Not added to linked list since it's never looked up.
		return LOADING_FAILED;

	// L4: Changed this next section to support lines added for #if (expression).
	// Each #if (expression) is pre-parsed *before* the main script in order for
	// function library auto-inclusions to be processed correctly.

	// Load the main script file.  This will also load any files it includes with #Include.
	if (   LoadIncludedFile(g_RunStdIn ? _T("*") : mFileSpec, false, false) != OK
		|| !AddLine(ACT_EXIT)) // Add an Exit to ensure lib auto-includes aren't auto-executed, for backward compatibility.
		return LOADING_FAILED;

	if (!PreparseExpressions(mFirstLine))
		return LOADING_FAILED; // Error was already displayed by the above call.
	// ABOVE: In v1.0.47, the above may have auto-included additional files from the userlib/stdlib.
	// That's why the above is done prior to adding the EXIT lines and other things below.

	// Preparse static initializers and #if expressions.
	PreparseStaticLines(mFirstLine);
	if (mFirstStaticLine)
	{
		// Prepend all Static initializers to the beginning of the auto-execute section.
		mLastStaticLine->mNextLine = mFirstLine;
		mFirstLine->mPrevLine = mLastStaticLine;
		mFirstLine = mFirstStaticLine;
	}
	
	// Scan for undeclared local variables which are named the same as a global variable.
	// This loop has two purposes (but it's all handled in PreprocessLocalVars()):
	//
	//  1) Allow super-global variables to be referenced above the point of declaration.
	//     This is a bit of a hack to work around the fact that variable references are
	//     resolved as they are encountered, before all declarations have been processed.
	//
	//  2) Warn the user (if appropriate) since they probably meant it to be global.
	//
	for (int i = 0; i < mFuncCount; ++i)
	{
		Func &func = *mFunc[i];
		if (!func.mIsBuiltIn && !(func.mDefaultVarType & VAR_FORCE_LOCAL))
		{
			PreprocessLocalVars(func, func.mVar, func.mVarCount);
			PreprocessLocalVars(func, func.mLazyVar, func.mLazyVarCount);
		}
	}

	// Resolve any unresolved base classes.
	if (mUnresolvedClasses)
	{
		if (!ResolveClasses())
			return LOADING_FAILED;
		mUnresolvedClasses->Release();
		mUnresolvedClasses = NULL;
	}

	// Check for classes potentially being overwritten.
	if (g_Warn_ClassOverwrite)
		CheckForClassOverwrite();

#ifndef AUTOHOTKEYSC
	if (mIncludeLibraryFunctionsThenExit)
	{
		delete mIncludeLibraryFunctionsThenExit;
		return 0; // Tell our caller to do a normal exit.
	}
#endif

	// v1.0.35.11: Restore original working directory so that changes made to it by the above (via
	// "#Include C:\Scripts" or "#Include %A_ScriptDir%" or even stdlib/userlib) do not affect the
	// script's runtime working directory.  This preserves the flexibility of having a startup-determined
	// working directory for the script's runtime (i.e. it seems best that the mere presence of
	// "#Include NewDir" should not entirely eliminate this flexibility).
	SetCurrentDirectory(g_WorkingDirOrig); // g_WorkingDirOrig previously set by WinMain().

	// Rather than do this, which seems kinda nasty if ever someday support same-line
	// else actions such as "else return", just add two EXITs to the end of every script.
	// That way, if the first EXIT added accidentally "corrects" an actionless ELSE
	// or IF, the second one will serve as the anchoring end-point (mRelatedLine) for that
	// IF or ELSE.  In other words, since we never want mRelatedLine to be NULL, this should
	// make absolutely sure of that:
	//if (mLastLine->mActionType == ACT_ELSE ||
	//	ACT_IS_IF(mLastLine->mActionType)
	//	...
	// Second ACT_EXIT: even if the last line of the script is already "exit", always add
	// another one in case the script ends in a label.  That way, every label will have
	// a non-NULL target, which simplifies other aspects of script execution.
	// Making sure that all scripts end with an EXIT ensures that if the script
	// file ends with ELSEless IF or an ELSE, that IF's or ELSE's mRelatedLine
	// will be non-NULL, which further simplifies script execution.
	// Not done since it's number doesn't much matter: ++mCombinedLineNumber;
	++mCombinedLineNumber;  // So that the EXITs will both show up in ListLines as the line # after the last physical one in the script.
	if (!(AddLine(ACT_EXIT) && AddLine(ACT_EXIT))) // Second exit guaranties non-NULL mRelatedLine(s).
		return LOADING_FAILED;
	mPlaceholderLabel->mJumpToLine = mLastLine; // To follow the rule "all labels should have a non-NULL line before the script starts running".

	if (   !PreparseBlocks(mFirstLine)
		|| !PreparseCommands(mFirstLine)   )
		return LOADING_FAILED; // Error was already displayed by the above calls.

	// Use FindOrAdd, not Add, because the user may already have added it simply by
	// referring to it in the script:
	if (   !(g_ErrorLevel = FindOrAddVar(_T("ErrorLevel")))   )
		return LOADING_FAILED; // Error.  Above already displayed it for us.
	// Initialize the var state to zero right before running anything in the script:
	g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	// Initialize the random number generator:
	// Note: On 32-bit hardware, the generator module uses only 2506 bytes of static
	// data, so it doesn't seem worthwhile to put it in a class (so that the mem is
	// only allocated on first use of the generator).  For v1.0.24, _ftime() is not
	// used since it could be as large as 0.5 KB of non-compressed code.  A simple call to
	// GetSystemTimeAsFileTime() seems just as good or better, since it produces
	// a FILETIME, which is "the number of 100-nanosecond intervals since January 1, 1601."
	// Use the low-order DWORD since the high-order one rarely changes.  If my calculations are correct,
	// the low-order 32-bits traverses its full 32-bit range every 7.2 minutes, which seems to make
	// using it as a seed superior to GetTickCount for most purposes.
	RESEED_RANDOM_GENERATOR;

	return TRUE; // Must be non-zero.
	// OBSOLETE: mLineCount was always non-zero at this point since above did AddLine().
	//return mLineCount; // The count of runnable lines that were loaded, which might be zero.
}



bool IsFunction(LPTSTR aBuf, bool *aPendingFunctionHasBrace = NULL)
// Helper function for LoadIncludedFile().
// Caller passes in an aBuf containing a candidate line such as "function(x, y)"
// Caller has ensured that aBuf is rtrim'd.
// Caller should pass NULL for aPendingFunctionHasBrace to indicate that function definitions (open-brace
// on same line as function) are not allowed.  When non-NULL *and* aBuf is a function call/def,
// *aPendingFunctionHasBrace is set to true if a brace is present at the end, or false otherwise.
// In addition, any open-brace is removed from aBuf in this mode.
{
	LPTSTR action_end = StrChrAny(aBuf, EXPR_ALL_SYMBOLS EXPR_ILLEGAL_CHARS);
	// Can't be a function definition or call without an open-parenthesis as first char found by the above.
	// In addition, if action_end isn't NULL, that confirms that the string in aBuf prior to action_end contains
	// no spaces, tabs, colons, or equal-signs.  As a result, it can't be:
	// 1) a hotstring, since they always start with at least one colon that would be caught immediately as 
	//    first-expr-char-is-not-open-parenthesis by the above.
	// 2) Any kind of math or assignment, such as var:=(x+y) or var+=(x+y).
	// The only things it could be other than a function call or function definition are:
	// Normal label that ends in single colon but contains an open-parenthesis prior to the colon, e.g. Label(x):
	// Single-line hotkey such as KeyName::MsgBox.  But since '(' isn't valid inside KeyName, this isn't a concern.
	// In addition, note that it isn't necessary to check for colons that lie outside of quoted strings because
	// we're only interested in the first "word" of aBuf: If this is indeed a function call or definition, what
	// lies to the left of its first open-parenthesis can't contain any colons anyway because the above would
	// have caught it as first-expr-char-is-not-open-parenthesis.  In other words, there's no way for a function's
	// opening parenthesis to occur after a legitimate/quoted colon or double-colon in its parameters.
	// v1.0.40.04: Added condition "action_end != aBuf" to allow a hotkey or remap or hotkey such as
	// such as "(::" to work even if it ends in a close-parenthesis such as "(::)" or "(::MsgBox )"
	if (   !(action_end && *action_end == '(' && action_end != aBuf
		&& tcslicmp(aBuf, _T("IF"), action_end - aBuf)
		&& tcslicmp(aBuf, _T("WHILE"), action_end - aBuf)) // v1.0.48.04: Recognize While() as loop rather than a function because many programmers are in the habit of writing while() and if().
		|| action_end[1] == ':'   ) // v1.0.44.07: This prevents "$(::fn_call()" from being seen as a function-call vs. hotkey-with-call.  For simplicity and due to rarity, omit_leading_whitespace() isn't called; i.e. assumes that the colon immediate follows the '('.
		return false;
	LPTSTR aBuf_last_char = action_end + _tcslen(action_end) - 1; // Above has already ensured that action_end is "(...".
	if (aPendingFunctionHasBrace) // Caller specified that an optional open-brace may be present at the end of aBuf.
	{
		if (*aPendingFunctionHasBrace = (*aBuf_last_char == '{')) // Caller has ensured that aBuf is rtrim'd.
		{
			*aBuf_last_char = '\0'; // For the caller, remove it from further consideration.
			aBuf_last_char = aBuf + rtrim(aBuf, aBuf_last_char - aBuf) - 1; // Omit trailing whitespace too.
		}
	}
	return *aBuf_last_char == ')'; // This last check avoids detecting a label such as "Label(x):" as a function.
	// Also, it seems best never to allow if(...) to be a function call, even if it's blank inside such as if().
	// In addition, it seems best not to allow if(...) to ever be a function definition since such a function
	// could never be called as ACT_EXPRESSION since it would be seen as an IF-stmt instead.
}



inline LPTSTR IsClassDefinition(LPTSTR aBuf, bool &aHasOTB)
{
	if (_tcsnicmp(aBuf, _T("Class"), 5) || !IS_SPACE_OR_TAB(aBuf[5])) // i.e. it's not "Class" followed by a space or tab.
		return NULL;
	LPTSTR class_name = omit_leading_whitespace(aBuf + 6);
	if (_tcschr(EXPR_ALL_SYMBOLS EXPR_ILLEGAL_CHARS, *class_name))
		// It's probably something like "Class := GetClass()".
		return NULL;
	// Check for opening brace on same line:
	LPTSTR aBuf_last_char = class_name + _tcslen(class_name) - 1;
	if (aHasOTB = (*aBuf_last_char == '{')) // Caller has ensured that aBuf is rtrim'd.
	{
		*aBuf_last_char = '\0'; // For the caller, remove it from further consideration.
		rtrim(aBuf, aBuf_last_char - aBuf); // Omit trailing whitespace too.
	}
	// Validation of the name is left up to the caller, for simplicity.
	return class_name;
}



ResultType Script::LoadIncludedFile(LPTSTR aFileSpec, bool aAllowDuplicateInclude, bool aIgnoreLoadFailure)
// Returns OK or FAIL.
// Below: Use double-colon as delimiter to set these apart from normal labels.
// The main reason for this is that otherwise the user would have to worry
// about a normal label being unintentionally valid as a hotkey, e.g.
// "Shift:" might be a legitimate label that the user forgot is also
// a valid hotkey:
#define HOTKEY_FLAG _T("::")
#define HOTKEY_FLAG_LENGTH 2
{
	if (!aFileSpec || !*aFileSpec) return FAIL;

#ifndef AUTOHOTKEYSC
	if (Line::sSourceFileCount >= Line::sMaxSourceFiles)
	{
		if (Line::sSourceFileCount >= ABSOLUTE_MAX_SOURCE_FILES)
			return ScriptError(_T("Too many includes.")); // Short msg since so rare.
		int new_max;
		if (Line::sMaxSourceFiles)
		{
			new_max = 2*Line::sMaxSourceFiles;
			if (new_max > ABSOLUTE_MAX_SOURCE_FILES)
				new_max = ABSOLUTE_MAX_SOURCE_FILES;
		}
		else
			new_max = 100;
		// For simplicity and due to rarity of every needing to, expand by reallocating the array.
		// Use a temp var. because realloc() returns NULL on failure but leaves original block allocated.
		LPTSTR *realloc_temp = (LPTSTR *)realloc(Line::sSourceFile, new_max * sizeof(LPTSTR)); // If passed NULL, realloc() will do a malloc().
		if (!realloc_temp)
			return ScriptError(ERR_OUTOFMEM); // Short msg since so rare.
		Line::sSourceFile = realloc_temp;
		Line::sMaxSourceFiles = new_max;
	}

	TCHAR full_path[MAX_PATH];
#endif

	// Keep this var on the stack due to recursion, which allows newly created lines to be given the
	// correct file number even when some #include's have been encountered in the middle of the script:
	int source_file_index = Line::sSourceFileCount;

	if (!source_file_index)
		// Since this is the first source file, it must be the main script file.  Just point it to the
		// location of the filespec already dynamically allocated:
		Line::sSourceFile[source_file_index] = mFileSpec;
#ifndef AUTOHOTKEYSC  // The "else" part below should never execute for compiled scripts since they never include anything (other than the main/combined script).
	else
	{
		// Get the full path in case aFileSpec has a relative path.  This is done so that duplicates
		// can be reliably detected (we only want to avoid including a given file more than once):
		LPTSTR filename_marker;
		GetFullPathName(aFileSpec, _countof(full_path), full_path, &filename_marker);
		// Check if this file was already included.  If so, it's not an error because we want
		// to support automatic "include once" behavior.  So just ignore repeats:
		if (!aAllowDuplicateInclude)
			for (int f = 0; f < source_file_index; ++f) // Here, source_file_index==Line::sSourceFileCount
				if (!lstrcmpi(Line::sSourceFile[f], full_path)) // Case insensitive like the file system (testing shows that "Ä" == "ä" in the NTFS, which is hopefully how lstrcmpi works regardless of locale).
					return OK;
		// The file is added to the list further below, after the file has been opened, in case the
		// opening fails and aIgnoreLoadFailure==true.
	}
#endif

	// <buf> should be no larger than LINE_SIZE because some later functions rely upon that:
	TCHAR msg_text[MAX_PATH + 256], buf1[LINE_SIZE], buf2[LINE_SIZE], suffix[16], pending_buf[LINE_SIZE];
	*pending_buf = '\0';
	LPTSTR buf = buf1, next_buf = buf2; // Oscillate between bufs to improve performance (avoids memcpy from buf2 to buf1).
	size_t buf_length, next_buf_length, suffix_length;
	bool pending_buf_has_brace;
	enum {
		Pending_Func,
		Pending_Class,
		Pending_Property
	} pending_buf_type;

#ifndef AUTOHOTKEYSC
	TextFile tfile, *fp = &tfile;
	if (!tfile.Open(aFileSpec, DEFAULT_READ_FLAGS, g_DefaultScriptCodepage))
	{
		if (aIgnoreLoadFailure)
			return OK;
		sntprintf(msg_text, _countof(msg_text), _T("%s file \"%s\" cannot be opened.")
			, Line::sSourceFileCount > 0 ? _T("#Include") : _T("Script"), aFileSpec);
		return ScriptError(msg_text);
	}

	// This is done only after the file has been successfully opened in case aIgnoreLoadFailure==true:
	if (source_file_index > 0)
		Line::sSourceFile[source_file_index] = SimpleHeap::Malloc(full_path);
	//else the first file was already taken care of by another means.

#else // Stand-alone mode (there are no include files in this mode since all of them were merged into the main script at the time of compiling).
	TextMem::Buffer textbuf(NULL, 0, false);

	HRSRC hRes;
	HGLOBAL hResData;

#ifdef _DEBUG
	if (hRes = FindResource(NULL, _T("AHK"), RT_RCDATA))
#else
	if (hRes = FindResource(NULL, _T(">AUTOHOTKEY SCRIPT<"), RT_RCDATA))
#endif
	{}
	else if (hRes = FindResource(NULL, _T(">AHK WITH ICON<"), RT_RCDATA))
	{}
	
	if ( !( hRes 
			&& (textbuf.mLength = SizeofResource(NULL, hRes))
			&& (hResData = LoadResource(NULL, hRes))
			&& (textbuf.mBuffer = LockResource(hResData)) ) )
	{
		MsgBox(_T("Could not extract script from EXE."), 0, aFileSpec);
		return FAIL;
	}

	TextMem tmem, *fp = &tmem;
	// NOTE: Ahk2Exe strips off the UTF-8 BOM.
	tmem.Open(textbuf, TextStream::READ | TextStream::EOL_CRLF | TextStream::EOL_ORPHAN_CR, CP_UTF8);
#endif

	++Line::sSourceFileCount;

	// File is now open, read lines from it.

	LPTSTR hotkey_flag, cp, cp1, hotstring_start, hotstring_options;
	Hotkey *hk;
	LineNumberType pending_buf_line_number, saved_line_number;
	HookActionType hook_action;
	bool is_label, suffix_has_tilde, hook_is_mandatory, in_comment_section, hotstring_options_all_valid, hotstring_execute;
	ResultType hotkey_validity;

	// For the remap mechanism, e.g. a::b
	int remap_stage;
	vk_type remap_source_vk, remap_dest_vk = 0; // Only dest is initialized to enforce the fact that it is the flag/signal to indicate whether remapping is in progress.
	TCHAR remap_source[32], remap_dest[32], remap_dest_modifiers[8]; // Must fit the longest key name (currently Browser_Favorites [17]), but buffer overflow is checked just in case.
	LPTSTR extra_event;
	bool remap_source_is_combo, remap_source_is_mouse, remap_dest_is_mouse, remap_keybd_to_mouse;

	// For the line continuation mechanism:
	bool do_ltrim, do_rtrim, literal_escapes, literal_derefs, literal_delimiters
		, has_continuation_section, is_continuation_line;
	#define CONTINUATION_SECTION_WITHOUT_COMMENTS 1 // MUST BE 1 because it's the default set by anything that's boolean-true.
	#define CONTINUATION_SECTION_WITH_COMMENTS    2 // Zero means "not in a continuation section".
	int in_continuation_section;

	LPTSTR next_option, option_end;
	TCHAR orig_char, one_char_string[2], two_char_string[3]; // Line continuation mechanism's option parsing.
	one_char_string[1] = '\0';  // Pre-terminate these to simplify code later below.
	two_char_string[2] = '\0';  //
	int continuation_line_count;

	#define MAX_FUNC_VAR_GLOBALS 2000
	Var *func_global_var[MAX_FUNC_VAR_GLOBALS];

	// Init both for main file and any included files loaded by this function:
	mCurrFileIndex = source_file_index;  // source_file_index is kept on the stack due to recursion (from #include).

#ifdef AUTOHOTKEYSC
	// -1 (MAX_UINT in this case) to compensate for the fact that there is a comment containing
	// the version number added to the top of each compiled script:
	LineNumberType phys_line_number = -1;
#else
	LineNumberType phys_line_number = 0;
#endif
	buf_length = GetLine(buf, LINE_SIZE - 1, 0, fp);

	if (in_comment_section = !_tcsncmp(buf, _T("/*"), 2))
	{
		// Fixed for v1.0.35.08. Must reset buffer to allow a script's first line to be "/*".
		*buf = '\0';
		buf_length = 0;
	}

	while (buf_length != -1)  // Compare directly to -1 since length is unsigned.
	{
		// For each whole line (a line with continuation section is counted as only a single line
		// for the purpose of this outer loop).

		// Keep track of this line's *physical* line number within its file for A_LineNumber and
		// error reporting purposes.  This must be done only in the outer loop so that it tracks
		// the topmost line of any set of lines merged due to continuation section/line(s)..
		mCombinedLineNumber = phys_line_number + 1;

		// This must be reset for each iteration because a prior iteration may have changed it, even
		// indirectly by calling something that changed it:
		mCurrLine = NULL;  // To signify that we're in transition, trying to load a new one.

		// v1.0.44.13: An additional call to IsDirective() is now made up here so that #CommentFlag affects
		// the line beneath it the same way as other lines (#EscapeChar et. al. didn't have this bug).
		// It's best not to process ALL directives up here because then they would no longer support a
		// continuation section beneath them (and possibly other drawbacks because it was never thoroughly
		// tested).
		if (!_tcsnicmp(buf, _T("#CommentFlag"), 12)) // Have IsDirective() process this now (it will also process it again later, which is harmless).
			if (IsDirective(buf) == FAIL) // IsDirective() already displayed the error.
				return FAIL;

		// Read in the next line (if that next line is the start of a continuation section, append
		// it to the line currently being processed:
		for (has_continuation_section = false, in_continuation_section = 0;;)
		{
			// This increment relies on the fact that this loop always has at least one iteration:
			++phys_line_number; // Tracks phys. line number in *this* file (independent of any recursion caused by #Include).
			next_buf_length = GetLine(next_buf, LINE_SIZE - 1, in_continuation_section, fp);
			if (next_buf_length && next_buf_length != -1 // Prevents infinite loop when file ends with an unclosed "/*" section.  Compare directly to -1 since length is unsigned.
				&& !in_continuation_section) // Multi-line comments can't be used in continuation sections. This line fixes '*/' being discarded in continuation sections (broken by L54).
			{
				if (!_tcsncmp(next_buf, _T("*/"), 2) // Check this even if !in_comment_section so it can be ignored (for convenience) and not treated as a line-continuation operator.
					&& (in_comment_section || next_buf[2] != ':' || next_buf[3] != ':')) // ...but support */:: as a hotkey.
				{
					in_comment_section = false;
					next_buf_length -= 2; // Adjust for removal of /* from the beginning of the string.
					tmemmove(next_buf, next_buf + 2, next_buf_length + 1);  // +1 to include the string terminator.
					next_buf_length = ltrim(next_buf, next_buf_length); // Get rid of any whitespace that was between the comment-end and remaining text.
					if (!*next_buf) // The rest of the line is empty, so it was just a naked comment-end.
						continue;
				}
				else if (in_comment_section)
					continue;

				if (!_tcsncmp(next_buf, _T("/*"), 2))
				{
					in_comment_section = true;
					continue; // It's now commented out, so the rest of this line is ignored.
				}
			}

			if (in_comment_section) // Above has incremented and read the next line, which is everything needed while inside /* .. */
			{
				if (next_buf_length == -1) // Compare directly to -1 since length is unsigned.
					break; // By design, it's not an error.  This allows "/*" to be used to comment out the bottommost portion of the script without needing a matching "*/".
				// Otherwise, continue reading lines so that they can be merged with the line above them
				// if they qualify as continuation lines.
				continue;
			}

			if (!in_continuation_section) // This is either the first iteration or the line after the end of a previous continuation section.
			{
				// v1.0.38.06: The following has been fixed to exclude "(:" and "(::".  These should be
				// labels/hotkeys, not the start of a continuation section.  In addition, a line that starts
				// with '(' but that ends with ':' should be treated as a label because labels such as
				// "(label):" are far more common than something obscure like a continuation section whose
				// join character is colon, namely "(Join:".
				if (   !(in_continuation_section = (next_buf_length != -1 && *next_buf == '(' // Compare directly to -1 since length is unsigned.
					&& next_buf[1] != ':' && next_buf[next_buf_length - 1] != ':'))   ) // Relies on short-circuit boolean order.
				{
					if (next_buf_length == -1)  // Compare directly to -1 since length is unsigned.
						break;
					if (!next_buf_length)
						// It is permitted to have blank lines and comment lines in between the line above
						// and any continuation section/line that might come after the end of the
						// comment/blank lines:
						continue;
					// SINCE ABOVE DIDN'T BREAK/CONTINUE, NEXT_BUF IS NON-BLANK.
					if (next_buf[next_buf_length - 1] == ':' && *next_buf != ',')
						// With the exception of lines starting with a comma, the last character of any
						// legitimate continuation line can't be a colon because expressions can't end
						// in a colon. The only exception is the ternary operator's colon, but that is
						// very rare because it requires the line after it also be a continuation line
						// or section, which is unusual to say the least -- so much so that it might be
						// too obscure to even document as a known limitation.  Anyway, by excluding lines
						// that end with a colon from consideration ambiguity with normal labels
						// and non-single-line hotkeys and hotstrings is eliminated.
						break;

					is_continuation_line = false; // Set default.
					switch(ctoupper(*next_buf)) // Above has ensured *next_buf != '\0' (toupper might have problems with '\0').
					{
					case 'A': // "AND".
						// See comments in the default section further below.
						if (!_tcsnicmp(next_buf, _T("and"), 3) && IS_SPACE_OR_TAB_OR_NBSP(next_buf[3])) // Relies on short-circuit boolean order.
						{
							cp = omit_leading_whitespace(next_buf + 3);
							// v1.0.38.06: The following was fixed to use EXPR_CORE vs. EXPR_OPERAND_TERMINATORS
							// to properly detect a continuation line whose first char after AND/OR is "!~*&-+()":
							if (!_tcschr(EXPR_CORE, *cp))
								// This check recognizes the following examples as NON-continuation lines by checking
								// that AND/OR aren't followed immediately by something that's obviously an operator:
								//    and := x, and = 2 (but not and += 2 since the an operand can have a unary plus/minus).
								// This is done for backward compatibility.  Also, it's documented that
								// AND/OR/NOT aren't supported as variable names inside expressions.
								is_continuation_line = true; // Override the default set earlier.
						}
						break;
					case 'O': // "OR".
						// See comments in the default section further below.
						if (ctoupper(next_buf[1]) == 'R' && IS_SPACE_OR_TAB_OR_NBSP(next_buf[2])) // Relies on short-circuit boolean order.
						{
							cp = omit_leading_whitespace(next_buf + 2);
							// v1.0.38.06: The following was fixed to use EXPR_CORE vs. EXPR_OPERAND_TERMINATORS
							// to properly detect a continuation line whose first char after AND/OR is "!~*&-+()":
							if (!_tcschr(EXPR_CORE, *cp)) // See comment in the "AND" case above.
								is_continuation_line = true; // Override the default set earlier.
						}
						break;
					default:
						// Desired line continuation operators:
						// Pretty much everything, namely:
						// +, -, *, /, //, **, <<, >>, &, |, ^, <, >, <=, >=, =, ==, <>, !=, :=, +=, -=, /=, *=, ?, :
						// And also the following remaining unaries (i.e. those that aren't also binaries): !, ~
						// The first line below checks for ::, ++, and --.  Those can't be continuation lines because:
						// "::" isn't a valid operator (this also helps performance if there are many hotstrings).
						// ++ and -- are ambiguous with an isolated line containing ++Var or --Var (and besides,
						// wanting to use ++ to continue an expression seems extremely rare, though if there's ever
						// demand for it, might be able to look at what lies to the right of the operator's operand
						// -- though that would produce inconsistent continuation behavior since ++Var itself still
						// could never be a continuation line due to ambiguity).
						//
						// The logic here isn't smart enough to differentiate between a leading ! or - that's
						// meant as a continuation character and one that isn't. Even if it were, it would
						// still be ambiguous in some cases because the author's intent isn't known; for example,
						// the leading minus sign on the second line below is ambiguous, so will probably remain
						// a continuation character in both v1 and v2:
						//    x := y 
						//    -z ? a:=1 : func() 
						if ((*next_buf == ':' || *next_buf == '+' || *next_buf == '-') && next_buf[1] == *next_buf // See above.
							// L31: '.' and '?' no longer require spaces; '.' without space is member-access (object) operator.
							//|| (*next_buf == '.' || *next_buf == '?') && !IS_SPACE_OR_TAB_OR_NBSP(next_buf[1]) // The "." and "?" operators require a space or tab after them to be legitimate.  For ".", this is done in case period is ever a legal character in var names, such as struct support.  For "?", it's done for backward compatibility since variable names can contain question marks (though "?" by itself is not considered a variable in v1.0.46).
								//&& next_buf[1] != '=' // But allow ".=" (and "?=" too for code simplicity), since ".=" is the concat-assign operator.
							|| !_tcschr(CONTINUATION_LINE_SYMBOLS, *next_buf)) // Line doesn't start with a continuation char.
							break; // Leave is_continuation_line set to its default of false.
						// Some of the above checks must be done before the next ones.
						if (   !(hotkey_flag = _tcsstr(next_buf, HOTKEY_FLAG))   ) // Without any "::", it can't be a hotkey or hotstring.
						{
							is_continuation_line = true; // Override the default set earlier.
							break;
						}
						if (*next_buf == ':') // First char is ':', so it's more likely a hotstring than a hotkey.
						{
							// Remember that hotstrings can contain what *appear* to be quoted literal strings,
							// so detecting whether a "::" is in a quoted/literal string in this case would
							// be more complicated.  That's one reason this other method is used.
							for (hotstring_options_all_valid = true, cp = next_buf + 1; *cp && *cp != ':'; ++cp)
								if (!IS_HOTSTRING_OPTION(*cp)) // Not a perfect test, but eliminates most of what little remaining ambiguity exists between ':' as a continuation character vs. ':' as the start of a hotstring.  It especially eliminates the ":=" operator.
								{
									hotstring_options_all_valid = false;
									break;
								}
							if (hotstring_options_all_valid && *cp == ':') // It's almost certainly a hotstring.
								break; // So don't treat it as a continuation line.
							//else it's not a hotstring but it might still be a hotkey such as ": & x::".
							// So continue checking below.
						}
						// Since above didn't "break", this line isn't a hotstring but it is probably a hotkey
						// because above already discovered that it contains "::" somewhere. So try to find out
						// if there's anything that disqualifies this from being a hotkey, such as some
						// expression line that contains a quoted/literal "::" (or a line starting with
						// a comma that contains an unquoted-but-literal "::" such as for FileAppend).
						if (*next_buf == ',')
						{
							cp = omit_leading_whitespace(next_buf + 1);
							// The above has set cp to the position of the non-whitespace item to the right of
							// this comma.  Normal (single-colon) labels can't contain commas, so only hotkey
							// labels are sources of ambiguity.  In addition, normal labels and hotstrings have
							// already been checked for, higher above.
							if (   _tcsncmp(cp, HOTKEY_FLAG, HOTKEY_FLAG_LENGTH) // It's not a hotkey such as ",::action".
								&& _tcsncmp(cp - 1, COMPOSITE_DELIMITER, COMPOSITE_DELIMITER_LENGTH)   ) // ...and it's not a hotkey such as ", & y::action".
								is_continuation_line = true; // Override the default set earlier.
						}
						else // First symbol in line isn't a comma but some other operator symbol.
						{
							// Check if the "::" found earlier appears to be inside a quoted/literal string.
							// This check is NOT done for a line beginning with a comma since such lines
							// can contain an unquoted-but-literal "::".  In addition, this check is done this
							// way to detect hotkeys such as the following:
							//   +keyname:: (and other hotkey modifier symbols such as ! and ^)
							//   +keyname1 & keyname2::
							//   +^:: (i.e. a modifier symbol followed by something that is a hotkey modifier and/or a hotkey suffix and/or an expression operator).
							//   <:: and &:: (i.e. hotkeys that are also expression-continuation symbols)
							// By contrast, expressions that qualify as continuation lines can look like:
							//   . "xxx::yyy"
							//   + x . "xxx::yyy"
							// In addition, hotkeys like the following should continue to be supported regardless
							// of how things are done here:
							//   ^"::
							//   . & "::
							// Finally, keep in mind that an expression-continuation line can start with two
							// consecutive unary operators like !! or !*. It can also start with a double-symbol
							// operator such as <=, <>, !=, &&, ||, //, **.
							for (cp = next_buf; cp < hotkey_flag && *cp != '"'; ++cp);
							if (cp == hotkey_flag) // No '"' found to left of "::", so this "::" appears to be a real hotkey flag rather than part of a literal string.
								break; // Treat this line as a normal line vs. continuation line.
							for (cp = hotkey_flag + HOTKEY_FLAG_LENGTH; *cp && *cp != '"'; ++cp);
							if (*cp)
							{
								// Closing quote was found so "::" is probably inside a literal string of an
								// expression (further checking seems unnecessary given the fairly extreme
								// rarity of using '"' as a key in a hotkey definition).
								is_continuation_line = true; // Override the default set earlier.
							}
							//else no closing '"' found, so this "::" probably belongs to something like +":: or
							// . & "::.  Treat this line as a normal line vs. continuation line.
						}
					} // switch(toupper(*next_buf))

					if (is_continuation_line)
					{
						if (buf_length + next_buf_length >= LINE_SIZE - 1) // -1 to account for the extra space added below.
							return ScriptError(ERR_CONTINUATION_SECTION_TOO_LONG, next_buf);
						if (*next_buf != ',') // Insert space before expression operators so that built/combined expression works correctly (some operators like 'and', 'or', '.', and '?' currently require spaces on either side) and also for readability of ListLines.
							buf[buf_length++] = ' ';
						tmemcpy(buf + buf_length, next_buf, next_buf_length + 1); // Append this line to prev. and include the zero terminator.
						buf_length += next_buf_length;
						continue; // Check for yet more continuation lines after this one.
					}
					// Since above didn't continue, there is no continuation line or section.  In addition,
					// since this line isn't blank, no further searching is needed.
					break;
				} // if (!in_continuation_section)

				// OTHERWISE in_continuation_section != 0, so the above has found the first line of a new
				// continuation section.
				continuation_line_count = 0; // Reset for this new section.
				// Otherwise, parse options.  First set the defaults, which can be individually overridden
				// by any options actually present.  RTrim defaults to ON for two reasons:
				// 1) Whitespace often winds up at the end of a lines in a text editor by accident.  In addition,
				//    whitespace at the end of any consolidated/merged line will be rtrim'd anyway, since that's
				//    how command parsing works.
				// 2) Copy & paste from the forum and perhaps other web sites leaves a space at the end of each
				//    line.  Although this behavior is probably site/browser-specific, it's a consideration.
				do_ltrim = g_ContinuationLTrim; // Start off at global default.
				do_rtrim = true; // Seems best to rtrim even if this line is a hotstring, since it is very rare that trailing spaces and tabs would ever be desirable.
				// For hotstrings (which could be detected via *buf==':'), it seems best not to default the
				// escape character (`) to be literal because the ability to have `t `r and `n inside the
				// hotstring continuation section seems more useful/common than the ability to use the
				// accent character by itself literally (which seems quite rare in most languages).
				literal_escapes = false;
				literal_derefs = false;
				literal_delimiters = true; // This is the default even for hotstrings because although using (*buf != ':') would improve loading performance, it's not a 100% reliable way to detect hotstrings.
				// The default is linefeed because:
				// 1) It's the best choice for hotstrings, for which the line continuation mechanism is well suited.
				// 2) It's good for FileAppend.
				// 3) Minor: Saves memory in large sections by being only one character instead of two.
				suffix[0] = '\n';
				suffix[1] = '\0';
				suffix_length = 1;
				for (next_option = omit_leading_whitespace(next_buf + 1); *next_option; next_option = omit_leading_whitespace(option_end))
				{
					// Find the end of this option item:
					if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
						option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.

					// Temporarily terminate to help eliminate ambiguity for words contained inside other words,
					// such as hypothetical "Checked" inside of "CheckedGray":
					orig_char = *option_end;
					*option_end = '\0';

					if (!_tcsnicmp(next_option, _T("Join"), 4))
					{
						next_option += 4;
						tcslcpy(suffix, next_option, _countof(suffix)); // The word "Join" by itself will product an empty string, as documented.
						// Passing true for the last parameter supports `s as the special escape character,
						// which allows space to be used by itself and also at the beginning or end of a string
						// containing other chars.
						ConvertEscapeSequences(suffix, g_EscapeChar, true);
						suffix_length = _tcslen(suffix);
					}
					else if (!_tcsnicmp(next_option, _T("LTrim"), 5))
						do_ltrim = (next_option[5] != '0');  // i.e. Only an explicit zero will turn it off.
					else if (!_tcsnicmp(next_option, _T("RTrim"), 5))
						do_rtrim = (next_option[5] != '0');
					else
					{
						// Fix for v1.0.36.01: Missing "else" above, because otherwise, the option Join`r`n
						// would be processed above but also be processed again below, this time seeing the
						// accent and thinking it's the signal to treat accents literally for the entire
						// continuation section rather than as escape characters.
						// Within this terminated option substring, allow the characters to be adjacent to
						// improve usability:
						for (; *next_option; ++next_option)
						{
							switch (*next_option)
							{
							case '`': // Although not using g_EscapeChar (reduces code size/complexity), #EscapeChar is still supported by continuation sections; it's just that enabling the option uses '`' rather than the custom escape-char (one reason is that that custom escape-char might be ambiguous with future/past options if it's something weird like an alphabetic character).
								literal_escapes = true;
								break;
							case '%': // Same comment as above.
								literal_derefs = true;
								break;
							case ',': // Same comment as above.
								literal_delimiters = false;
								break;
							case 'C': // v1.0.45.03: For simplicity, anything that begins with "C" is enough to
							case 'c': // identify it as the option to allow comments in the section.
								in_continuation_section = CONTINUATION_SECTION_WITH_COMMENTS; // Override the default, which is boolean true (i.e. 1).
								break;
							case ')':
								// Probably something like (x.y)[z](), which is not intended as the beginning of
								// a continuation section.  Doing this only when ")" is found should remove the
								// need to escape "(" in most real-world expressions while still allowing new
								// options to be added later with minimal risk of breaking scripts.
								in_continuation_section = 0;
								*option_end = orig_char; // Undo the temporary termination.
								goto process_completed_line;
							}
						}
					}

					// If the item was not handled by the above, ignore it because it is unknown.
					*option_end = orig_char; // Undo the temporary termination.
				} // for() each item in option list

				// "has_continuation_section" indicates whether the line we're about to construct is partially
				// composed of continuation lines beneath it.  It's separate from continuation_line_count
				// in case there is another continuation section immediately after/adjacent to the first one,
				// but the second one doesn't have any lines in it:
				has_continuation_section = true;

				continue; // Now that the open-parenthesis of this continuation section has been processed, proceed to the next line.
			} // if (!in_continuation_section)

			// Since above didn't "continue", we're in the continuation section and thus next_buf contains
			// either a line to be appended onto buf or the closing parenthesis of this continuation section.
			if (next_buf_length == -1) // Compare directly to -1 since length is unsigned.
				return ScriptError(ERR_MISSING_CLOSE_PAREN, buf);
			if (next_buf_length == -2) // v1.0.45.03: Special flag that means "this is a commented-out line to be
				continue;              // entirely omitted from the continuation section." Compare directly to -2 since length is unsigned.

			if (*next_buf == ')')
			{
				in_continuation_section = 0; // Facilitates back-to-back continuation sections and proper incrementing of phys_line_number.
				next_buf_length = rtrim(next_buf); // Done because GetLine() wouldn't have done it due to have told it we're in a continuation section.
				// Anything that lies to the right of the close-parenthesis gets appended verbatim, with
				// no trimming (for flexibility) and no options-driven translation:
				cp = next_buf + 1;  // Use temp var cp to avoid altering next_buf (for maintainability).
				--next_buf_length;  // This is now the length of cp, not next_buf.
			}
			else
			{
				cp = next_buf;
				// The following are done in this block only because anything that comes after the closing
				// parenthesis (i.e. the block above) is exempt from translations and custom trimming.
				// This means that commas are always delimiters and percent signs are always deref symbols
				// in the previous block.
				if (do_rtrim)
					next_buf_length = rtrim(next_buf, next_buf_length);
				if (do_ltrim)
					next_buf_length = ltrim(next_buf, next_buf_length);
				// Escape each comma and percent sign in the body of the continuation section so that
				// the later parsing stages will see them as literals.  Although, it's not always
				// necessary to do this (e.g. commas in the last parameter of a command don't need to
				// be escaped, nor do percent signs in hotstrings' auto-replace text), the settings
				// are applied unconditionally because:
				// 1) Determining when its safe to omit the translation would add a lot of code size and complexity.
				// 2) The translation doesn't affect the functionality of the script since escaped literals
				//    are always de-escaped at a later stage, at least for everything that's likely to matter
				//    or that's reasonable to put into a continuation section (e.g. a hotstring's replacement text).
				// UPDATE for v1.0.44.11: #EscapeChar, #DerefChar, #Delimiter are now supported by continuation
				// sections because there were some requests for that in forum.
				int replacement_count = 0;
				if (literal_escapes) // literal_escapes must be done FIRST because otherwise it would also replace any accents added for literal_delimiters or literal_derefs.
				{
					one_char_string[0] = g_EscapeChar; // These strings were terminated earlier, so no need to
					two_char_string[0] = g_EscapeChar; // do it here.  In addition, these strings must be set by
					two_char_string[1] = g_EscapeChar; // each iteration because the #EscapeChar (and similar directives) can occur multiple times, anywhere in the script.
					replacement_count += StrReplace(next_buf, one_char_string, two_char_string, SCS_SENSITIVE, UINT_MAX, LINE_SIZE);
				}
				if (literal_derefs)
				{
					one_char_string[0] = g_DerefChar;
					two_char_string[0] = g_EscapeChar;
					two_char_string[1] = g_DerefChar;
					replacement_count += StrReplace(next_buf, one_char_string, two_char_string, SCS_SENSITIVE, UINT_MAX, LINE_SIZE);
				}
				if (literal_delimiters)
				{
					one_char_string[0] = g_delimiter;
					two_char_string[0] = g_EscapeChar;
					two_char_string[1] = g_delimiter;
					replacement_count += StrReplace(next_buf, one_char_string, two_char_string, SCS_SENSITIVE, UINT_MAX, LINE_SIZE);
				}

				if (replacement_count) // Update the length if any actual replacements were done.
					next_buf_length = _tcslen(next_buf);
			} // Handling of a normal line within a continuation section.

			// Must check the combined length only after anything that might have expanded the string above.
			if (buf_length + next_buf_length + suffix_length >= LINE_SIZE)
				return ScriptError(ERR_CONTINUATION_SECTION_TOO_LONG, cp);

			++continuation_line_count;
			// Append this continuation line onto the primary line.
			// The suffix for the previous line gets written immediately prior writing this next line,
			// which allows the suffix to be omitted for the final line.  But if this is the first line,
			// No suffix is written because there is no previous line in the continuation section.
			// In addition, cp!=next_buf, this is the special line whose text occurs to the right of the
			// continuation section's closing parenthesis. In this case too, the previous line doesn't
			// get a suffix.
			if (continuation_line_count > 1 && suffix_length && cp == next_buf)
			{
				tmemcpy(buf + buf_length, suffix, suffix_length + 1); // Append and include the zero terminator.
				buf_length += suffix_length; // Must be done only after the old value of buf_length was used above.
			}
			if (next_buf_length)
			{
				tmemcpy(buf + buf_length, cp, next_buf_length + 1); // Append this line to prev. and include the zero terminator.
				buf_length += next_buf_length; // Must be done only after the old value of buf_length was used above.
			}
		} // for() each sub-line (continued line) that composes this line.

process_completed_line:
		// buf_length can't be -1 (though next_buf_length can) because outer loop's condition prevents it:
		if (!buf_length) // Done only after the line number increments above so that the physical line number is properly tracked.
			goto continue_main_loop; // In lieu of "continue", for performance.

		// Since neither of the above executed, or they did but didn't "continue",
		// buf now contains a non-commented line, either by itself or built from
		// any continuation sections/lines that might have been present.  Also note that
		// by design, phys_line_number will be greater than mCombinedLineNumber whenever
		// a continuation section/lines were used to build this combined line.

		// If there's a previous line waiting to be processed, its fate can now be determined based on the
		// nature of *this* line:
		if (*pending_buf)
		{
			// Somewhat messy to decrement then increment later, but it's probably easier than the
			// alternatives due to the use of "continue" in some places above.  NOTE: phys_line_number
			// would not need to be decremented+incremented even if the below resulted in a recursive
			// call to us (though it doesn't currently) because line_number's only purpose is to
			// remember where this layer left off when the recursion collapses back to us.
			// Fix for v1.0.31.05: It's not enough just to decrement mCombinedLineNumber because there
			// might be some blank lines or commented-out lines between this function call/definition
			// and the line that follows it, each of which will have previously incremented mCombinedLineNumber.
			saved_line_number = mCombinedLineNumber;
			mCombinedLineNumber = pending_buf_line_number;  // Done so that any syntax errors that occur during the calls below will report the correct line number.
			// Open brace means this is a function definition. NOTE: buf was already ltrimmed by GetLine().
			// Could use *g_act[ACT_BLOCK_BEGIN].Name instead of '{', but it seems too elaborate to be worth it.
			if (*buf == '{' || pending_buf_has_brace) // v1.0.41: Support one-true-brace, e.g. fn(...) {
			{
				switch (pending_buf_type)
				{
				case Pending_Class:
					if (!DefineClass(pending_buf))
						return FAIL;
					break;
				case Pending_Property:
					if (!DefineClassProperty(pending_buf))
						return FAIL;
					break;
				case Pending_Func:
					// Note that two consecutive function definitions aren't possible:
					// fn1()
					// fn2()
					// {
					//  ...
					// }
					// In the above, the first would automatically be deemed a function call by means of
					// the check higher above (by virtue of the fact that the line after it isn't an open-brace).
					if (g->CurrentFunc)
					{
						// Though it might be allowed in the future -- perhaps to have nested functions have
						// access to their parent functions' local variables, or perhaps just to improve
						// script readability and maintainability -- it's currently not allowed because of
						// the practice of maintaining the func_global_var list on our stack:
						return ScriptError(_T("Functions cannot contain functions."), pending_buf);
					}
					if (!DefineFunc(pending_buf, func_global_var))
						return FAIL;
					if (pending_buf_has_brace) // v1.0.41: Support one-true-brace for function def, e.g. fn() {
					{
						if (!AddLine(ACT_BLOCK_BEGIN))
							return FAIL;
						mCurrLine = NULL; // L30: Prevents showing misleading vicinity lines if the line after a OTB function def is a syntax error.
					}
					break;
#ifdef _DEBUG
				default:
					return ScriptError(_T("DEBUG: pending_buf_type has an unexpected value."));
#endif
				}
			}
			else // It's a function call on a line by itself, such as fn(x). It can't be if(..) because another section checked that.
			{
				if (pending_buf_type != Pending_Func) // Missing open-brace for class definition.
					return ScriptError(ERR_MISSING_OPEN_BRACE, pending_buf);
				if (mClassObjectCount && !g->CurrentFunc) // Unexpected function call in class definition.
					return ScriptError(mClassProperty ? ERR_MISSING_OPEN_BRACE : ERR_INVALID_LINE_IN_CLASS_DEF, pending_buf);
				if (!ParseAndAddLine(pending_buf, ACT_EXPRESSION))
					return FAIL;
				mCurrLine = NULL; // Prevents showing misleading vicinity lines if the line after a function call is a syntax error.
			}
			*pending_buf = '\0'; // Reset now that it's been fully handled, as an indicator for subsequent iterations.
			if (pending_buf_type != Pending_Func) // Class or property.
			{
				if (!pending_buf_has_brace)
				{
					// This is the open-brace of a class definition, so requires no further processing.
					if (  !*(cp = omit_leading_whitespace(buf + 1))  )
					{
						mCombinedLineNumber = saved_line_number;
						goto continue_main_loop;
					}
					// Otherwise, there's something following the "{", possibly "}" or a function definition.
					tmemmove(buf, cp, (buf_length = _tcslen(cp)) + 1);
				}
			}
			mCombinedLineNumber = saved_line_number;
			// Now fall through to the below so that *this* line (the one after it) will be processed.
			// Note that this line might be a pre-processor directive, label, etc. that won't actually
			// become a runtime line per se.
		} // if (*pending_function)

		if (*buf == '}' && mClassObjectCount && !g->CurrentFunc)
		{
			if (mClassProperty)
			{
				// Close this property definition.
				mClassProperty = NULL;
				if (mClassPropertyDef)
				{
					free(mClassPropertyDef);
					mClassPropertyDef = NULL;
				}
			}
			else
			{
				// End of class definition.
				--mClassObjectCount;
				mClassObject[mClassObjectCount]->EndClassDefinition(); // Remove instance variables from the class object.
				mClassObject[mClassObjectCount]->Release();
				// Revert to the name of the class this class is nested inside, or "" if none.
				if (cp1 = _tcsrchr(mClassName, '.'))
					*cp1 = '\0';
				else
					*mClassName = '\0';
			}
			// Allow multiple end-braces or other declarations to the right of "}":
			if (   *(buf = omit_leading_whitespace(buf + 1))   )
			{
				buf_length = _tcslen(buf); // Update.
				mCurrLine = NULL;  // To signify that we're in transition, trying to load a new line.
				goto process_completed_line; // Have the main loop process the contents of "buf" as though it came in from the script.
			}
			goto continue_main_loop; // It's just a naked "{" or "}", so no more processing needed for this line.
		}

		if (mClassProperty && !g->CurrentFunc) // This is checked before IsFunction() to prevent method definitions inside a property.
		{
			if (!_tcsnicmp(buf, _T("Get"), 3) || !_tcsnicmp(buf, _T("Set"), 3))
			{
				LPTSTR cp = omit_leading_whitespace(buf + 3);
				if ( !*cp || (*cp == '{' && !cp[1]) )
				{
					// Defer this line until the next line comes in to simplify handling of '{' and OTB.
					// For simplicity, pass the property definition to DefineFunc instead of the actual
					// line text, even though it makes some error messages a bit inaccurate. (That would
					// happen anyway when DefineFunc() finds a syntax error in the parameter list.)
					_tcscpy(pending_buf, mClassPropertyDef);
					LPTSTR dot = _tcschr(pending_buf, '.');
					dot[1] = *buf; // Replace the x in property.xet(params).
					pending_buf_line_number = mCombinedLineNumber;
					pending_buf_has_brace = *cp == '{';
					pending_buf_type = Pending_Func;
					goto continue_main_loop;
				}
			}
			return ScriptError(ERR_INVALID_LINE_IN_PROPERTY_DEF, buf);
		}

		// By doing the following section prior to checking for hotkey and hotstring labels, double colons do
		// not need to be escaped inside naked function calls and function definitions such as the following:
		// fn("::")      ; Function call.
		// fn(Str="::")  ; Function definition with default value for its parameter.
		if (IsFunction(buf, &pending_buf_has_brace)) // If true, it's either a function definition or a function call (to be distinguished later).
		{
			// Defer this line until the next line comes in, which helps determine whether this line is
			// a function call vs. definition:
			_tcscpy(pending_buf, buf);
			pending_buf_line_number = mCombinedLineNumber;
			pending_buf_type = Pending_Func;
			goto continue_main_loop; // In lieu of "continue", for performance.
		}
		
		if (!g->CurrentFunc)
		{
			if (LPTSTR class_name = IsClassDefinition(buf, pending_buf_has_brace))
			{
				// Defer this line until the next line comes in to simplify handling of '{' and OTB:
				_tcscpy(pending_buf, class_name);
				pending_buf_line_number = mCombinedLineNumber;
				pending_buf_type = Pending_Class;
				goto continue_main_loop; // In lieu of "continue", for performance.
			}

			if (mClassObjectCount)
			{
				// Check for assignment first, in case of something like "Static := 123".
				for (cp = buf; IS_IDENTIFIER_CHAR(*cp) || *cp == '.'; ++cp);
				if (cp > buf) // i.e. buf begins with an identifier.
				{
					cp = omit_leading_whitespace(cp);
					if (*cp == ':' && cp[1] == '=') // This is an assignment.
					{
						if (!DefineClassVars(buf, false)) // See above for comments.
							return FAIL;
						goto continue_main_loop;
					}
					if ( !*cp || *cp == '[' || (*cp == '{' && !cp[1]) ) // Property
					{
						size_t length = _tcslen(buf);
						if (pending_buf_has_brace = (buf[length - 1] == '{'))
						{
							// Omit '{' and trailing whitespace from further consideration.
							rtrim(buf, length - 1);
						}
						
						// Defer this line until the next line comes in to simplify handling of '{' and OTB:
						_tcscpy(pending_buf, buf);
						pending_buf_line_number = mCombinedLineNumber;
						pending_buf_type = Pending_Property;
						goto continue_main_loop; // In lieu of "continue", for performance.
					}
				}
				if (!_tcsnicmp(buf, _T("Static"), 6) && IS_SPACE_OR_TAB(buf[6]))
				{
					if (!DefineClassVars(buf + 7, true))
						return FAIL; // Above already displayed the error.
					goto continue_main_loop; // In lieu of "continue", for performance.
				}
				if (*buf == '#') // See the identical section further below for comments.
				{
					saved_line_number = mCombinedLineNumber;
					switch(IsDirective(buf))
					{
					case CONDITION_TRUE:
						mCurrFileIndex = source_file_index;
						mCombinedLineNumber = saved_line_number;
						goto continue_main_loop;
					case FAIL:
						return FAIL;
					}
				}
				// Anything not already handled above is not valid directly inside a class definition.
				return ScriptError(ERR_INVALID_LINE_IN_CLASS_DEF, buf);
			}
		}

		// The following "examine_line" label skips the following parts above:
		// 1) IsFunction() because that's only for a function call or definition alone on a line
		//    e.g. not "if fn()" or x := fn().  Those who goto this label don't need that processing.
		// 2) The "if (*pending_function)" block: Doesn't seem applicable for the callers of this label.
		// 3) The inner loop that handles continuation sections: Not needed by the callers of this label.
		// 4) Things like the following should be skipped because callers of this label don't want the
		//    physical line number changed (which would throw off the count of lines that lie beneath a remap):
		//    mCombinedLineNumber = phys_line_number + 1;
		//    ++phys_line_number;
		// 5) "mCurrLine = NULL": Probably not necessary since it's only for error reporting.  Worst thing
		//    that could happen is that syntax errors would be thrown off, which testing shows isn't the case.
examine_line:
		// "::" alone isn't a hotstring, it's a label whose name is colon.
		// Below relies on the fact that no valid hotkey can start with a colon, since
		// ": & somekey" is not valid (since colon is a shifted key) and colon itself
		// should instead be defined as "+;::".  It also relies on short-circuit boolean:
		hotstring_start = NULL;
		hotkey_flag = NULL;
		if (buf[0] == ':' && buf[1])
		{
			hotstring_options = buf + 1; // Point it to the hotstring's option letters, if any.
			if (buf[1] != ':')
			{
				// The following relies on the fact that options should never contain a literal colon.
				// ALSO, the following doesn't use IS_HOTSTRING_OPTION() for backward compatibility,
				// performance, and because it seems seldom if ever necessary at this late a stage.
				if (   !(hotstring_start = _tcschr(hotstring_options, ':'))   )
					hotstring_start = NULL; // Indicate that this isn't a hotstring after all.
				else
					++hotstring_start; // Points to the hotstring itself.
			}
			else // Double-colon, so it's a hotstring if there's more after this (but this means no options are present).
				if (buf[2])
					hotstring_start = buf + 2;
				//else it's just a naked "::", which is considered to be an ordinary label whose name is colon.
		}
		if (hotstring_start)
		{
			// Check for 'X' option early since escape sequence processing depends on it.
			hotstring_execute = g_HSSameLineAction;
			for (cp = hotstring_options; cp < hotstring_start; ++cp)
				if (ctoupper(*cp) == 'X')
				{
					hotstring_execute = cp[1] != '0';
					break;
				}
			// Find the hotstring's final double-colon by considering escape sequences from left to right.
			// This is necessary for to handles cases such as the following:
			// ::abc```::::Replacement String
			// The above hotstring translates literally into "abc`::".
			LPTSTR escaped_double_colon = NULL;
			for (cp = hotstring_start; ; ++cp)  // Increment to skip over the symbol just found by the inner for().
			{
				for (; *cp && *cp != g_EscapeChar && *cp != ':'; ++cp);  // Find the next escape char or colon.
				if (!*cp) // end of string.
					break;
				cp1 = cp + 1;
				if (*cp == ':')
				{
					if (*cp1 == ':') // Found a non-escaped double-colon, so this is the right one.
					{
						hotkey_flag = cp++;  // Increment to have loop skip over both colons.
						if (hotstring_execute)
							break; // Let ParseAndAddLine() properly handle any escape sequences.
						// else continue with the loop so that escape sequences in the replacement
						// text (if there is replacement text) are also translated.
					}
					// else just a single colon, or the second colon of an escaped pair (`::), so continue.
					continue;
				}
				switch (*cp1)
				{
					// Only lowercase is recognized for these:
					case 'a': *cp1 = '\a'; break;  // alert (bell) character
					case 'b': *cp1 = '\b'; break;  // backspace
					case 'f': *cp1 = '\f'; break;  // formfeed
					case 'n': *cp1 = '\n'; break;  // newline
					case 'r': *cp1 = '\r'; break;  // carriage return
					case 't': *cp1 = '\t'; break;  // horizontal tab
					case 'v': *cp1 = '\v'; break;  // vertical tab
					// Otherwise, if it's not one of the above, the escape-char is considered to
					// mark the next character as literal, regardless of what it is. Examples:
					// `` -> `
					// `:: -> :: (effectively)
					// `; -> ;
					// `c -> c (i.e. unknown escape sequences resolve to the char after the `)
				}
				// Below has a final +1 to include the terminator:
				tmemmove(cp, cp1, _tcslen(cp1) + 1);
				// Since single colons normally do not need to be escaped, this increments one extra
				// for double-colons to skip over the entire pair so that its second colon
				// is not seen as part of the hotstring's final double-colon.  Example:
				// ::ahc```::::Replacement String
				if (*cp == ':' && *cp1 == ':')
					++cp;
			} // for()
			if (!hotkey_flag)
				hotstring_start = NULL;  // Indicate that this isn't a hotstring after all.
		}
		if (!hotstring_start) // Not a hotstring (hotstring_start is checked *again* in case above block changed it; otherwise hotkeys like ": & x" aren't recognized).
		{
			// Note that there may be an action following the HOTKEY_FLAG (on the same line).
			if (hotkey_flag = _tcsstr(buf, HOTKEY_FLAG)) // Find the first one from the left, in case there's more than 1.
			{
				if (hotkey_flag == buf && hotkey_flag[2] == ':') // v1.0.46: Support ":::" to mean "colon is a hotkey".
					++hotkey_flag;
				// v1.0.40: It appears to be a hotkey, but validate it as such before committing to processing
				// it as a hotkey.  If it fails validation as a hotkey, treat it as a command that just happens
				// to contain a double-colon somewhere.  This avoids the need to escape double colons in scripts.
				// Note: Hotstrings can't suffer from this type of ambiguity because a leading colon or pair of
				// colons makes them easier to detect.
				cp = omit_trailing_whitespace(buf, hotkey_flag); // For maintainability.
				orig_char = *cp;
				*cp = '\0'; // Temporarily terminate.
				hotkey_validity = Hotkey::TextInterpret(omit_leading_whitespace(buf), NULL, false); // Passing NULL calls it in validate-only mode.
				switch (hotkey_validity)
				{
				case FAIL:
					hotkey_flag = NULL; // It's not a valid hotkey, so indicate that it's a command (i.e. one that contains a literal double-colon, which avoids the need to escape the double-colon).
					break;
				case CONDITION_FALSE:
					return FAIL; // It's an invalid hotkey and above already displayed the error message.
				//case CONDITION_TRUE:
					// It's a key that doesn't exist on the current keyboard layout.  Leave hotkey_flag set
					// so that the section below handles it as a hotkey.  This allows it to end the auto-exec
					// section and register the appropriate label even though it won't be an active hotkey.
				}
				*cp = orig_char; // Undo the temp. termination above.
			}
		}

		// Treat a naked "::" as a normal label whose label name is colon:
		if (is_label = (hotkey_flag && hotkey_flag > buf)) // It's a hotkey/hotstring label.
		{
			if (g->CurrentFunc)
			{
				// Even if it weren't for the reasons below, the first hotkey/hotstring label in a script
				// will end the auto-execute section with a "return".  Therefore, if this restriction here
				// is ever removed, be sure that that extra return doesn't get put inside the function.
				//
				// The reason for not allowing hotkeys and hotstrings inside a function's body is that
				// when the subroutine is launched, the hotstring/hotkey would be using the function's
				// local variables.  But that is not appropriate and it's likely to cause problems even
				// if it were.  It doesn't seem useful in any case.  By contrast, normal labels can
				// safely exist inside a function body and since the body is a block, other validation
				// ensures that a Gosub or Goto can't jump to it from outside the function.
				return ScriptError(_T("Hotkeys/hotstrings are not allowed inside functions."), buf);
			}
			if (mLastLine && mLastLine->mActionType == ACT_IFWINACTIVE)
			{
				mCurrLine = mLastLine; // To show vicinity lines.
				return ScriptError(_T("IfWin should be #IfWin."), buf);
			}
			*hotkey_flag = '\0'; // Terminate so that buf is now the label itself.
			hotkey_flag += HOTKEY_FLAG_LENGTH;  // Now hotkey_flag is the hotkey's action, if any.
			if (!hotstring_start)
			{
				ltrim(hotkey_flag); // Has already been rtrimmed by GetLine().
				// Not done because Hotkey::TextInterpret() does not allow trailing whitespace: 
				//rtrim(buf); // Trim the new substring inside of buf (due to temp termination). It has already been ltrimmed.
				cp = hotkey_flag; // Set default, conditionally overridden below (v1.0.44.07).
				// v1.0.40: Check if this is a remap rather than hotkey:
				if (   *hotkey_flag // This hotkey's action is on the same line as its label.
					&& (remap_dest_vk = hotkey_flag[1] ? TextToVK(cp = Hotkey::TextToModifiers(hotkey_flag, NULL)) : 0xFF)   ) // And the action appears to be a remap destination rather than a command.
					// For above:
					// Fix for v1.0.44.07: Set remap_dest_vk to 0xFF if hotkey_flag's length is only 1 because:
					// 1) It allows a destination key that doesn't exist in the keyboard layout (such as 6::ð in
					//    English).
					// 2) It improves performance a little by not calling TextToVK except when the destination key
					//    might be a mouse button or some longer key name whose actual/correct VK value is relied
					//    upon by other places below.
					// Fix for v1.0.40.01: Since remap_dest_vk is also used as the flag to indicate whether
					// this line qualifies as a remap, must do it last in the statement above.  Otherwise,
					// the statement might short-circuit and leave remap_dest_vk as non-zero even though
					// the line shouldn't be a remap.  For example, I think a hotkey such as "x & y::return"
					// would trigger such a bug.
				{
					// These will be ignored in other stages if it turns out not to be a remap later below:
					remap_source_vk = TextToVK(cp1 = Hotkey::TextToModifiers(buf, NULL)); // An earlier stage verified that it's a valid hotkey, though VK could be zero.
					remap_source_is_combo = _tcsstr(cp1, COMPOSITE_DELIMITER);
					remap_source_is_mouse = IsMouseVK(remap_source_vk);
					remap_dest_is_mouse = IsMouseVK(remap_dest_vk);
					remap_keybd_to_mouse = !remap_source_is_mouse && remap_dest_is_mouse;
					sntprintf(remap_source, _countof(remap_source), _T("%s%s%s")
						, remap_source_is_combo ? _T("") : _T("*") // v1.1.27.01: Omit * when the remap source is a custom combo.
						, _tcslen(cp1) == 1 && IsCharUpper(*cp1) ? _T("+") : _T("")  // Allow A::b to be different than a::b.
						, buf); // Include any modifiers too, e.g. ^b::c.
					tcslcpy(remap_dest, cp, _countof(remap_dest));      // But exclude modifiers here; they're wanted separately.
					tcslcpy(remap_dest_modifiers, hotkey_flag, _countof(remap_dest_modifiers));
					if (cp - hotkey_flag < _countof(remap_dest_modifiers)) // Avoid reading beyond the end.
						remap_dest_modifiers[cp - hotkey_flag] = '\0';   // Terminate at the proper end of the modifier string.
					remap_stage = 0; // Init for use in the next stage.
					// In the unlikely event that the dest key has the same name as a command, disqualify it
					// from being a remap (as documented).  v1.0.40.05: If the destination key has any modifiers,
					// it is unambiguously a key name rather than a command, so the switch() isn't necessary.
					if (*remap_dest_modifiers)
						goto continue_main_loop; // It will see that remap_dest_vk is non-zero and act accordingly.
					switch (remap_dest_vk)
					{
					case VK_CONTROL: // Checked in case it was specified as "Control" rather than "Ctrl".
					case VK_SLEEP:
						if (StrChrAny(hotkey_flag, _T(" \t,"))) // Not using g_delimiter (reduces code size/complexity).
							break; // Any space, tab, or enter means this is a command rather than a remap destination.
						goto continue_main_loop; // It will see that remap_dest_vk is non-zero and act accordingly.
					// "Return" and "Pause" as destination keys are always considered commands instead.
					// This is documented and is done to preserve backward compatibility.
					case VK_RETURN:
						// v1.0.40.05: Although "Return" can't be a destination, "Enter" can be.  Must compare
						// to "Return" not "Enter" so that things like "vk0d" (the VK of "Enter") can also be a
						// destination key:
						if (!_tcsicmp(remap_dest, _T("Return")))
							break;
						goto continue_main_loop; // It will see that remap_dest_vk is non-zero and act accordingly.
					case VK_PAUSE:  // Used for both "Pause" and "Break"
						break;
					default: // All other VKs are valid destinations and thus the remap is valid.
						goto continue_main_loop; // It will see that remap_dest_vk is non-zero and act accordingly.
					}
					// Since above didn't goto, indicate that this is not a remap after all:
					remap_dest_vk = 0;
				}
			}
			// else don't trim hotstrings since literal spaces in both substrings are significant.

			// If this is the first hotkey label encountered, Add a return before
			// adding the label, so that the auto-execute section is terminated.
			// Only do this if the label is a hotkey because, for example,
			// the user may want to fully execute a normal script that contains
			// no hotkeys but does contain normal labels to which the execution
			// should fall through, if specified, rather than returning.
			// But this might result in weirdness?  Example:
			//testlabel:
			// Sleep, 1
			// return
			// ^a::
			// return
			// It would put the hard return in between, which is wrong.  But in the case above,
			// the first sub shouldn't have a return unless there's a part up top that ends in Exit.
			// So if Exit is encountered before the first hotkey, don't add the return?
			// Even though wrong, the return is harmless because it's never executed?  Except when
			// falling through from above into a hotkey (which probably isn't very valid anyway)?
			// Update: Below must check if there are any true hotkey labels, not just regular labels.
			// Otherwise, a normal (non-hotkey) label in the autoexecute section would count and
			// thus the RETURN would never be added here, even though it should be:
			
			// Notes about the below macro:
			// Fix for v1.0.34: Don't point labels to this particular RETURN so that labels
			// can point to the very first hotkey or hotstring in a script.  For example:
			// Goto Test
			// Test:
			// ^!z::ToolTip Without the fix`, this is never displayed by "Goto Test".
			// UCHAR_MAX signals it not to point any pending labels to this RETURN.
			// mCurrLine = NULL -> signifies that we're in transition, trying to load a new one.
			#define CHECK_mNoHotkeyLabels \
			if (mNoHotkeyLabels)\
			{\
				mNoHotkeyLabels = false;\
				if (!AddLine(ACT_RETURN, NULL, UCHAR_MAX))\
					return FAIL;\
				mCurrLine = NULL;\
			}
			CHECK_mNoHotkeyLabels
			// For hotstrings, the below makes the label include leading colon(s) and the full option
			// string (if any) so that the uniqueness of labels is preserved.  For example, we want
			// the following two hotstring labels to be unique rather than considered duplicates:
			// ::abc::
			// :c:abc::
			if (!AddLabel(buf, true)) // Always add a label before adding the first line of its section.
				return FAIL;
			
			if (hotstring_start)
			{
				if (!*hotstring_start)
				{
					// The following error message won't indicate the correct line number because
					// the hotstring (as a label) does not actually exist as a line.  But it seems
					// best to report it this way in case the hotstring is inside a #Include file,
					// so that the correct file name and approximate line number are shown:
					return ScriptError(_T("This hotstring is missing its abbreviation."), buf); // Display buf vs. hotkey_flag in case the line is simply "::::".
				}
				// In the case of hotstrings, hotstring_start is the beginning of the hotstring itself,
				// i.e. the character after the second colon.  hotstring_options is the first character
				// in the options list (which is terminated by a colon).  hotkey_flag is blank or the
				// hotstring's auto-replace text or same-line action.
				// v1.0.42: Unlike hotkeys, duplicate hotstrings are not detected.  This is because
				// hotstrings are less commonly used and also because it requires more code to find
				// hotstring duplicates (and performs a lot worse if a script has thousands of
				// hotstrings) because of all the hotstring options.
				if (!Hotstring::AddHotstring(mLastLabel->mName, mLastLabel, hotstring_options
					, hotstring_start, hotstring_execute ? _T("") : hotkey_flag, has_continuation_section))
					return FAIL;
				// The following section is similar to the one for hotkeys below, but is done after
				// parsing the hotstring's options. An attempt at combining the two sections actually
				// increased the final code size, so they're left separate.
				if (*hotkey_flag)
				{
					if (hotstring_execute)
						if (!ParseAndAddLine(hotkey_flag))
							return FAIL;
					// This is done for hotstrings with same-line action via 'X' and also auto-replace
					// hotstrings in case gosub/goto is ever used to jump to their labels:
					if (!AddLine(ACT_RETURN))
						return FAIL;
				}
			}
			else // It's a hotkey vs. hotstring.
			{
				hook_action = 0; // Set default.
				if (*hotkey_flag) // This hotkey's action is on the same line as its label.
				{
					// Don't add the alt-tabs as a line, since it has no meaning as a script command.
					// But do put in the Return regardless, in case this label is ever jumped to
					// via Goto/Gosub:
					if (   !(hook_action = Hotkey::ConvertAltTab(hotkey_flag, false))   )
						if (!ParseAndAddLine(hotkey_flag))
							return FAIL;
					// Also add a Return that's implicit for a single-line hotkey.
					if (!AddLine(ACT_RETURN))
						return FAIL;
				}
				if (hk = Hotkey::FindHotkeyByTrueNature(buf, suffix_has_tilde, hook_is_mandatory)) // Parent hotkey found.  Add a child/variant hotkey for it.
				{
					if (hook_action) // suffix_has_tilde has always been ignored for these types (alt-tab hotkeys).
					{
						// Hotkey::Dynamic() contains logic and comments similar to this, so maintain them together.
						// An attempt to add an alt-tab variant to an existing hotkey.  This might have
						// merit if the intention is to make it alt-tab now but to later disable that alt-tab
						// aspect via the Hotkey cmd to let the context-sensitive variants shine through
						// (take effect).
						hk->mHookAction = hook_action;
					}
					else
					{
						// Detect duplicate hotkey variants to help spot bugs in scripts.
						if (hk->FindVariant()) // See if there's already a variant matching the current criteria (suffix_has_tilde does not make variants distinct form each other because it would require firing two hotkey IDs in response to pressing one hotkey, which currently isn't in the design).
						{
							mCurrLine = NULL;  // Prevents showing unhelpful vicinity lines.
							return ScriptError(_T("Duplicate hotkey."), buf);
						}
						if (!hk->AddVariant(mLastLabel, suffix_has_tilde))
							return ScriptError(ERR_OUTOFMEM, buf);
						if (hook_is_mandatory || (!g_os.IsWin9x() && g_ForceKeybdHook))
						{
							// Require the hook for all variants of this hotkey if any variant requires it.
							// This seems more intuitive than the old behaviour, which required $ or #UseHook
							// to be used on the *first* variant, even though it affected all variants.
#ifdef CONFIG_WIN9X
							if (g_os.IsWin9x())
								hk->mUnregisterDuringThread = true;
							else
#endif
								hk->mKeybdHookMandatory = true;
						}
					}
				}
				else // No parent hotkey yet, so create it.
					if (   !(hk = Hotkey::AddHotkey(mLastLabel, hook_action, mLastLabel->mName, suffix_has_tilde, false))   )
					{
						if (hotkey_validity != CONDITION_TRUE)
							return FAIL; // It already displayed the error.
						// This hotkey uses a single-character key name, which could be valid on some other
						// keyboard layout.  Allow the script to start, but warn the user about the problem.
						// Note that this hotkey's label is still valid even though the hotkey wasn't created.
#ifndef AUTOHOTKEYSC
						if (!mIncludeLibraryFunctionsThenExit) // Current keyboard layout is not relevant in /iLib mode.
#endif
						{
							sntprintf(msg_text, _countof(msg_text), _T("Note: The hotkey %s will not be active because it does not exist in the current keyboard layout."), buf);
							MsgBox(msg_text);
						}
					}
			}
			goto continue_main_loop; // In lieu of "continue", for performance.
		} // if (is_label = ...)

		// Otherwise, not a hotkey or hotstring.  Check if it's a generic, non-hotkey label:
		if (buf[buf_length - 1] == ':') // Labels must end in a colon (buf was previously rtrimmed).
		{
			if (buf_length == 1) // v1.0.41.01: Properly handle the fact that this line consists of only a colon.
				return ScriptError(ERR_UNRECOGNIZED_ACTION, buf);
			// Labels (except hotkeys) must contain no whitespace, delimiters, or escape-chars.
			// This is to avoid problems where a legitimate action-line ends in a colon,
			// such as "WinActivate SomeTitle" and "#Include c:".
			// We allow hotkeys to violate this since they may contain commas, and since a normal
			// script line (i.e. just a plain command) is unlikely to ever end in a double-colon:
			for (cp = buf, is_label = true; *cp; ++cp)
				if (IS_SPACE_OR_TAB(*cp) || *cp == g_delimiter || *cp == g_EscapeChar)
				{
					is_label = false;
					break;
				}
			if (is_label // It's a generic label, since valid hotkeys and hotstrings have already been handled.
				&& !(buf[buf_length - 2] == ':' && buf_length > 2)) // i.e. allow "::" as a normal label, but consider anything else with double-colon to be an error (reported at a later stage).
			{
				// v1.0.44.04: Fixed this check by moving it after the above loop.
				// Above has ensured buf_length>1, so it's safe to check for double-colon:
				// v1.0.44.03: Don't allow anything that ends in "::" (other than a line consisting only
				// of "::") to be a normal label.  Assume it's a command instead (if it actually isn't, a
				// later stage will report it as "invalid hotkey"). This change avoids the situation in
				// which a hotkey like ^!ä:: is seen as invalid because the current keyboard layout doesn't
				// have a "ä" key. Without this change, if such a hotkey appears at the top of the script,
				// its subroutine would execute immediately as a normal label, which would be especially
				// bad if the hotkey were something like the "Shutdown" command.
				// Update: Hotkeys with single-character names like ^!ä are now handled earlier, so that
				// anything else with double-colon can be detected as an error.  The checks above prevent
				// something like foo:: from being interpreted as a generic label, so when the line fails
				// to resolve to a command or expression, an error message will be shown.
				buf[--buf_length] = '\0';  // Remove the trailing colon.
				rtrim(buf, buf_length); // Has already been ltrimmed.
				if (!AddLabel(buf, false))
					return FAIL;
				goto continue_main_loop; // In lieu of "continue", for performance.
			}
		}
		// Since above didn't "goto", it's not a label.
		if (*buf == '#')
		{
			saved_line_number = mCombinedLineNumber; // Backup in case IsDirective() processes an include file, which would change mCombinedLineNumber's value.
			switch(IsDirective(buf)) // Note that it may alter the contents of buf, at least in the case of #IfWin.
			{
			case CONDITION_TRUE:
				// Since the directive may have been a #include which called us recursively,
				// restore the class's values for these two, which are maintained separately
				// like this to avoid having to specify them in various calls, especially the
				// hundreds of calls to ScriptError() and LineError():
				mCurrFileIndex = source_file_index;
				mCombinedLineNumber = saved_line_number;
				goto continue_main_loop; // In lieu of "continue", for performance.
			case FAIL: // IsDirective() already displayed the error.
				return FAIL;
			//case CONDITION_FALSE: Do nothing; let processing below handle it.
			}
		}
		// Otherwise, treat it as a normal script line.

		if (*buf == '{' || *buf == '}')
		{
			if (!AddLine(*buf == '{' ? ACT_BLOCK_BEGIN : ACT_BLOCK_END))
				return FAIL;
			// Allow any command/action, directive or label to the right of "{" or "}":
			if (   *(buf = omit_leading_whitespace(buf + 1))   )
			{
				buf_length = _tcslen(buf); // Update.
				mCurrLine = NULL;  // To signify that we're in transition, trying to load a new line.
				goto process_completed_line; // Have the main loop process the contents of "buf" as though it came in from the script.
			}
			goto continue_main_loop; // It's just a naked "{" or "}", so no more processing needed for this line.
		}

		// Parse the command, assignment or expression, including any same-line open brace or sub-action
		// for ELSE, TRY, CATCH or FINALLY.  Unlike braces at the start of a line (processed above), this
		// does not allow directives or labels to the right of the command.
		if (!ParseAndAddLine(buf))
			return FAIL;

continue_main_loop: // This method is used in lieu of "continue" for performance and code size reduction.
		if (remap_dest_vk)
		{
			// For remapping, decided to use a "macro expansion" approach because I think it's considerably
			// smaller in code size and complexity than other approaches would be.  I originally wanted to
			// do it with the hook by changing the incoming event prior to passing it back out again (for
			// example, a::b would transform an incoming 'a' keystroke into 'b' directly without having
			// to suppress the original keystroke and simulate a new one).  Unfortunately, the low-level
			// hooks apparently do not allow this.  Here is the test that confirmed it:
			// if (event.vkCode == 'A')
			// {
			//	event.vkCode = 'B';
			//	event.scanCode = 0x30; // Or use vk_to_sc(event.vkCode).
			//	return CallNextHookEx(g_KeybdHook, aCode, wParam, lParam);
			// }
			switch (++remap_stage)
			{
			case 1: // Stage 1: Add key-down hotkey label, e.g. *LButton::
				buf_length = _stprintf(buf, _T("%s::"), remap_source); // Should be no risk of buffer overflow due to prior validation.
				goto examine_line; // Have the main loop process the contents of "buf" as though it came in from the script.
			case 2: // Stage 2.
				// Copied into a writable buffer for maintainability: AddLine() might rely on this.
				// Also, it seems unnecessary to set press-duration to -1 even though the auto-exec section might
				// have set it to something higher than -1 because:
				// 1) Press-duration doesn't apply to normal remappings since they use down-only and up-only events.
				// 2) Although it does apply to remappings such as a::B and a::^b (due to press-duration being
				//    applied after a change to modifier state), those remappings are fairly rare and supporting
				//    a non-negative-one press-duration (almost always 0) probably adds a degree of flexibility
				//    that may be desirable to keep.
				// 3) SendInput may become the predominant SendMode, so press-duration won't often be in effect anyway.
				// 4) It has been documented that remappings use the auto-execute section's press-duration.
				_tcscpy(buf, _T("-1")); // Does NOT need to be "-1, -1" for SetKeyDelay (see above).
				// The primary reason for adding Key/MouseDelay -1 is to minimize the chance that a one of
				// these hotkey threads will get buried under some other thread such as a timer, which
				// would disrupt the remapping if #MaxThreadsPerHotkey is at its default of 1.
				AddLine(remap_dest_is_mouse ? ACT_SETMOUSEDELAY : ACT_SETKEYDELAY, &buf, 1, NULL); // PressDuration doesn't need to be specified because it doesn't affect down-only and up-only events.
				if (remap_keybd_to_mouse)
				{
					// Since source is keybd and dest is mouse, prevent keyboard auto-repeat from auto-repeating
					// the mouse button (since that would be undesirable 90% of the time).  This is done
					// by inserting a single extra IF-statement above the Send that produces the down-event:
					buf_length = _stprintf(buf, _T("if not GetKeyState(\"%s\")"), remap_dest); // Should be no risk of buffer overflow due to prior validation.
					remap_stage = 9; // Have it hit special stage 9+1 next time for code reduction purposes.
					goto examine_line; // Have the main loop process the contents of "buf" as though it came in from the script.
				}
				// Otherwise, remap_keybd_to_mouse==false, so fall through to next case.
			case 10:
				extra_event = _T(""); // Set default.
				switch (remap_dest_vk)
				{
				case VK_LMENU:
				case VK_RMENU:
				case VK_MENU:
					switch (remap_source_vk)
					{
					case VK_LCONTROL:
					case VK_CONTROL:
						extra_event = _T("{LCtrl up}"); // Somewhat surprisingly, this is enough to make "Ctrl::Alt" properly remap both right and left control.
						break;
					case VK_RCONTROL:
						extra_event = _T("{RCtrl up}");
						break;
					// Below is commented out because its only purpose was to allow a shift key remapped to alt
					// to be able to alt-tab.  But that wouldn't work correctly due to the need for the following
					// hotkey, which does more harm than good by impacting the normal Alt key's ability to alt-tab
					// (if the normal Alt key isn't remapped): *Tab::Send {Blind}{Tab}
					//case VK_LSHIFT:
					//case VK_SHIFT:
					//	extra_event = "{LShift up}";
					//	break;
					//case VK_RSHIFT:
					//	extra_event = "{RShift up}";
					//	break;
					}
					break;
				}
				mCurrLine = NULL; // v1.0.40.04: Prevents showing misleading vicinity lines for a syntax-error such as %::%
				_stprintf(buf, _T("{Blind}%s%s{%s DownR}"), extra_event, remap_dest_modifiers, remap_dest); // v1.0.44.05: DownTemp vs. Down. See Send's DownTemp handler for details.
				if (!AddLine(ACT_SEND, &buf, 1, NULL)) // v1.0.40.04: Check for failure due to bad remaps such as %::%.
					return FAIL;
				AddLine(ACT_RETURN);
				mCurrLine = NULL; // Prevents showing misleading vicinity lines for something like "RAlt up::AppsKey" -> "*RAlt up up::".
				// Add key-up hotkey label, e.g. *LButton up::
				buf_length = _stprintf(buf, _T("%s up::"), remap_source); // Should be no risk of buffer overflow due to prior validation.
				remap_stage = 2; // Adjust to hit stage 3 next time (in case this is stage 10).
				goto examine_line; // Have the main loop process the contents of "buf" as though it came in from the script.
			case 3: // Stage 3.
				_tcscpy(buf, _T("-1"));
				AddLine(remap_dest_is_mouse ? ACT_SETMOUSEDELAY : ACT_SETKEYDELAY, &buf, 1, NULL);
				_stprintf(buf, _T("{Blind}{%s Up}"), remap_dest); // Unlike the down-event above, remap_dest_modifiers is not included for the up-event; e.g. ^{b up} is inappropriate.
				AddLine(ACT_SEND, &buf, 1, NULL);
				AddLine(ACT_RETURN);
				remap_dest_vk = 0; // Reset to signal that the remapping expansion is now complete.
				break; // Fall through to the next section so that script loading can resume at the next line.
			}
		} // if (remap_dest_vk)
		// Since above didn't "continue", resume loading script line by line:
		buf = next_buf;
		buf_length = next_buf_length;
		next_buf = (buf == buf1) ? buf2 : buf1;
		// The line above alternates buffers (toggles next_buf to be the unused buffer), which helps
		// performance because it avoids memcpy from buf2 to buf1.
	} // for each whole/constructed line.

	if (*pending_buf) // Since this is the last non-comment line, the pending function must be a function call, not a function definition.
	{
		// Somewhat messy to decrement then increment later, but it's probably easier than the
		// alternatives due to the use of "continue" in some places above.
		saved_line_number = mCombinedLineNumber;
		mCombinedLineNumber = pending_buf_line_number; // Done so that any syntax errors that occur during the calls below will report the correct line number.
		if (pending_buf_type != Pending_Func)
			return ScriptError(pending_buf_has_brace ? ERR_MISSING_CLOSE_BRACE : ERR_MISSING_OPEN_BRACE, pending_buf);
		if (!ParseAndAddLine(pending_buf, ACT_EXPRESSION)) // Must be function call vs. definition since otherwise the above would have detected the opening brace beneath it and already cleared pending_function.
			return FAIL;
		mCombinedLineNumber = saved_line_number;
	}

	if (mClassObjectCount && !source_file_index) // or mClassProperty, which implies mClassObjectCount != 0.
	{
		// A class definition has not been closed with "}".  Previously this was detected by adding
		// the open and close braces as lines, but this way is simpler and has less overhead.
		// The downside is that the line number won't be shown; however, the class name will.
		// Seems okay not to show mClassProperty->mName since the class is missing "}" as well.
		return ScriptError(ERR_MISSING_CLOSE_BRACE, mClassName);
	}

	++mCombinedLineNumber; // L40: Put the implicit ACT_EXIT on the line after the last physical line (for the debugger).

	// This is not required, it is called by the destructor.
	// fp->Close();
	return OK;
}



size_t Script::GetLine(LPTSTR aBuf, int aMaxCharsToRead, int aInContinuationSection, TextStream *ts)
{
	size_t aBuf_length = 0;

	if (!aBuf || !ts) return -1;
	if (aMaxCharsToRead < 1) return 0;
	if (  !(aBuf_length = ts->ReadLine(aBuf, aMaxCharsToRead))  ) // end-of-file or error
	{
		*aBuf = '\0';  // Reset since on error, contents added by fgets() are indeterminate.
		return -1;
	}
	if (aBuf[aBuf_length-1] == '\n')
		--aBuf_length;
	aBuf[aBuf_length] = '\0';

	if (aInContinuationSection)
	{
		LPTSTR cp = omit_leading_whitespace(aBuf);
		if (aInContinuationSection == CONTINUATION_SECTION_WITHOUT_COMMENTS) // By default, continuation sections don't allow comments (lines beginning with a semicolon are treated as literal text).
		{
			// Caller relies on us to detect the end of the continuation section so that trimming
			// will be done on the final line of the section and so that a comment can immediately
			// follow the closing parenthesis (on the same line).  Example:
			// (
			//	Text
			// ) ; Same line comment.
			if (*cp != ')') // This isn't the last line of the continuation section, so leave the line untrimmed (caller will apply the ltrim setting on its own).
				return aBuf_length; // Earlier sections are responsible for keeping aBufLength up-to-date with any changes to aBuf.
			//else this line starts with ')', so continue on to later section that checks for a same-line comment on its right side.
		}
		else // aInContinuationSection == CONTINUATION_SECTION_WITH_COMMENTS (i.e. comments are allowed in this continuation section).
		{
			// Fix for v1.0.46.09+: The "com" option shouldn't put "ltrim" into effect.
			if (!_tcsncmp(cp, g_CommentFlag, g_CommentFlagLength)) // Case sensitive.
			{
				*aBuf = '\0'; // Since this line is a comment, have the caller ignore it.
				return -2; // Callers tolerate -2 only when in a continuation section.  -2 indicates, "don't include this line at all, not even as a blank line to which the JOIN string (default "\n") will apply.
			}
			if (*cp == ')') // This isn't the last line of the continuation section, so leave the line untrimmed (caller will apply the ltrim setting on its own).
			{
				ltrim(aBuf); // Ltrim this line unconditionally so that caller will see that it starts with ')' without having to do extra steps.
				aBuf_length = _tcslen(aBuf); // ltrim() doesn't always return an accurate length, so do it this way.
			}
		}
	}
	// Since above didn't return, either:
	// 1) We're not in a continuation section at all, so apply ltrim() to support semicolons after tabs or
	//    other whitespace.  Seems best to rtrim also.
	// 2) CONTINUATION_SECTION_WITHOUT_COMMENTS but this line is the final line of the section.  Apply
	//    trim() and other logic further below because caller might rely on it.
	// 3) CONTINUATION_SECTION_WITH_COMMENTS (i.e. comments allowed), but this line isn't a comment (though
	//    it may start with ')' and thus be the final line of this section). In either case, need to check
	//    for same-line comments further below.
	if (aInContinuationSection != CONTINUATION_SECTION_WITH_COMMENTS) // Case #1 & #2 above.
	{
		aBuf_length = trim(aBuf);
		if (!_tcsncmp(aBuf, g_CommentFlag, g_CommentFlagLength)) // Case sensitive.
		{
			// Due to other checks, aInContinuationSection==false whenever the above condition is true.
			*aBuf = '\0';
			return 0;
		}
	}
	//else CONTINUATION_SECTION_WITH_COMMENTS (case #3 above), which due to other checking also means that
	// this line isn't a comment (though it might have a comment on its right side, which is checked below).
	// CONTINUATION_SECTION_WITHOUT_COMMENTS would already have returned higher above if this line isn't
	// the last line of the continuation section.

	// Handle comment-flags that appear to the right of a valid line.
	LPTSTR cp, prevp;
	for (cp = _tcsstr(aBuf, g_CommentFlag); cp; cp = _tcsstr(cp + g_CommentFlagLength, g_CommentFlag))
	{
		// If no whitespace to its left, it's not a valid comment.
		// We insist on this so that a semi-colon (for example) immediately after
		// a word (as semi-colons are often used) will not be considered a comment.
		prevp = cp - 1;
		if (prevp < aBuf) // should never happen because we already checked above.
		{
			*aBuf = '\0';
			return 0;
		}
		if (IS_SPACE_OR_TAB_OR_NBSP(*prevp)) // consider it to be a valid comment flag
		{
			*prevp = '\0';
			aBuf_length = rtrim_with_nbsp(aBuf, prevp - aBuf); // Since it's our responsibility to return a fully trimmed string.
			break; // Once the first valid comment-flag is found, nothing after it can matter.
		}
		else // No whitespace to the left.
		{
			// The following is done here, at this early stage, to support escaping the comment flag in
			// hotkeys and directives (the latter is mostly for backward-compatibility).
			LPTSTR p;
			for (p = prevp; p > aBuf && *p == g_EscapeChar && p[-1] == g_EscapeChar; p -= 2);
			if (p >= aBuf && *p == g_EscapeChar) // Odd number of escape chars: remove the final escape char.
			{
				// The following isn't exactly correct because it prevents an include filename from ever
				// containing the literal string "`;".  This is because attempts to escape the accent via
				// "``;" are not supported.  This is documented here as a known limitation because fixing
				// it would probably break existing scripts that rely on the fact that accents do not need
				// to be escaped inside #Include.  Also, the likelihood of "`;" appearing literally in a
				// legitimate #Include file seems vanishingly small.
				tmemmove(prevp, prevp + 1, _tcslen(prevp + 1) + 1);  // +1 for the terminator.
				--aBuf_length;
				// Then continue looking for others.
			}
			// else there wasn't any whitespace to its left, so keep looking in case there's
			// another further on in the line.
		}
	} // for()

	return aBuf_length; // The above is responsible for keeping aBufLength up-to-date with any changes to aBuf.
}



inline ResultType Script::IsDirective(LPTSTR aBuf)
// aBuf must be a modifiable string since this function modifies it in the case of "#Include %A_ScriptDir%"
// changes it.  It must also be large enough to accept the replacement of %A_ScriptDir% with a larger string.
// Returns CONDITION_TRUE, CONDITION_FALSE, or FAIL.
// Note: Don't assume that every line in the script that starts with '#' is a directive
// because hotkeys can legitimately start with that as well.  i.e., the following line should
// not be unconditionally ignored, just because it starts with '#', since it is a valid hotkey:
// #y::run, notepad
{
	TCHAR end_flags[] = {' ', '\t', g_delimiter, '\0'}; // '\0' must be last.
	LPTSTR directive_end, parameter_raw;
	if (   !(directive_end = StrChrAny(aBuf, end_flags))   )
	{
		directive_end = aBuf + _tcslen(aBuf); // Point it to the zero terminator.
		parameter_raw = NULL;
	}
	else
		if (!*(parameter_raw = omit_leading_whitespace(directive_end)))
			parameter_raw = NULL;

	// The raw parameter retains any leading comma for those directives that need that (none currently).
	// But the following omits that comma:
	LPTSTR parameter;
	if (!parameter_raw)
		parameter = NULL;
	else // Since parameter_raw is non-NULL, it's also non-blank and non-whitespace due to the above checking.
		if (*parameter_raw != g_delimiter)
			parameter = parameter_raw;
		else // It's a delimiter, so "parameter" will be whatever non-whitespace character follows it, if any.
			if (!*(parameter = omit_leading_whitespace(parameter_raw + 1)))
				parameter = NULL;
			//else leave it set to the value returned by omit_leading_whitespace().

	int value; // Helps detect values that are too large, since some of the target globals are UCHAR.

	// Use _tcsnicmp() so that a match is found as long as aBuf starts with the string in question.
	// e.g. so that "#SingleInstance, on" will still work too, but
	// "#a::run, something, "#SingleInstance" (i.e. a hotkey) will not be falsely detected
	// due to using a more lenient function such as tcscasestr().
	// UPDATE: Using strlicmp() now so that overlapping names, such as #MaxThreads and #MaxThreadsPerHotkey,
	// won't get mixed up:
	#define IS_DIRECTIVE_MATCH(directive) (!tcslicmp(aBuf, directive, directive_name_length))
	size_t directive_name_length = directive_end - aBuf; // To avoid calculating it every time in the macro above.

	bool is_include_again = false; // Set default in case of short-circuit boolean.
	if (IS_DIRECTIVE_MATCH(_T("#Include")) || (is_include_again = IS_DIRECTIVE_MATCH(_T("#IncludeAgain"))))
	{
		// Standalone EXEs ignore this directive since the included files were already merged in
		// with the main file when the script was compiled.  These should have been removed
		// or commented out by Ahk2Exe, but just in case, it's safest to ignore them:
#ifdef AUTOHOTKEYSC
		return CONDITION_TRUE;
#else
		// If the below decision is ever changed, be sure to update ahk2exe with the same change:
		// "parameter" is checked rather than parameter_raw for backward compatibility with earlier versions,
		// in which a leading comma is not considered part of the filename.  Although this behavior is incorrect
		// because it prevents files whose names start with a comma from being included without the first
		// delim-comma being there too, it is kept because filenames that start with a comma seem
		// exceedingly rare.  As a workaround, the script can do #Include ,,FilenameWithLeadingComma.ahk
		if (!parameter)
			return ScriptError(ERR_PARAM1_REQUIRED, aBuf);
		// v1.0.32:
		bool ignore_load_failure = (parameter[0] == '*' && ctoupper(parameter[1]) == 'I'); // Relies on short-circuit boolean order.
		if (ignore_load_failure)
		{
			parameter += 2;
			if (IS_SPACE_OR_TAB(*parameter)) // Skip over at most one space or tab, since others might be a literal part of the filename.
				++parameter;
		}

		if (*parameter == '<') // Support explicitly-specified <standard_lib_name>.
		{
			LPTSTR parameter_end = _tcschr(parameter, '>');
			if (parameter_end && !parameter_end[1])
			{
				++parameter; // Remove '<'.
				*parameter_end = '\0'; // Remove '>'.
				bool error_was_shown, file_was_found;
				// Save the working directory.
				TCHAR buf[MAX_PATH];
				if (!GetCurrentDirectory(_countof(buf) - 1, buf))
					*buf = '\0';
				// Attempt to include a script file based on the same rules as func() auto-include:
				FindFuncInLibrary(parameter, parameter_end - parameter, error_was_shown, file_was_found, false);
				// Restore the working directory so that any ordinary #includes will work correctly.
				SetCurrentDirectory(buf);
				// If any file was included, consider it a success; i.e. allow #include <lib> and #include <lib_func>.
				if (!error_was_shown && (file_was_found || ignore_load_failure))
					return CONDITION_TRUE;
				*parameter_end = '>'; // Restore '>' for display to the user.
				return error_was_shown ? FAIL : ScriptError(_T("Function library not found."), aBuf);
			}
			//else invalid syntax; treat it as a regular #include which will almost certainly fail.
		}

		LPTSTR include_path;
		if (!DerefInclude(include_path, parameter))
			return FAIL;

		DWORD attr = GetFileAttributes(include_path);
		if (attr != 0xFFFFFFFF && (attr & FILE_ATTRIBUTE_DIRECTORY)) // File exists and its a directory (possibly A_ScriptDir or A_AppData set above).
		{
			// v1.0.35.11 allow changing of load-time directory to increase flexibility.  This feature has
			// been asked for directly or indirectly several times.
			// If a script ever wants to use a string like "%A_ScriptDir%" literally in an include's filename,
			// that would not work.  But that seems too rare to worry about.
			// v1.0.45.01: Call SetWorkingDir() vs. SetCurrentDirectory() so that it succeeds even for a root
			// drive like C: that lacks a backslash (see SetWorkingDir() for details).
			SetWorkingDir(include_path);
			free(include_path);
			return CONDITION_TRUE;
		}
		// Since above didn't return, it's a file (or non-existent file, in which case the below will display
		// the error).  This will also display any other errors that occur:
		ResultType result = LoadIncludedFile(include_path, is_include_again, ignore_load_failure) ? CONDITION_TRUE : FAIL;
		free(include_path);
		return result;
#endif
	}

	if (IS_DIRECTIVE_MATCH(_T("#NoEnv")))
	{
		g_NoEnv = TRUE;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#NoTrayIcon")))
	{
		g_NoTrayIcon = true;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#Persistent")))
	{
		g_persistent = true;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#SingleInstance")))
	{
		g_AllowOnlyOneInstance = SINGLE_INSTANCE_PROMPT; // Set default.
		if (parameter)
		{
			if (!_tcsicmp(parameter, _T("Force")))
				g_AllowOnlyOneInstance = SINGLE_INSTANCE_REPLACE;
			else if (!_tcsicmp(parameter, _T("Ignore")))
				g_AllowOnlyOneInstance = SINGLE_INSTANCE_IGNORE;
			else if (!_tcsicmp(parameter, _T("Off")))
				g_AllowOnlyOneInstance = SINGLE_INSTANCE_OFF;
		}
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#InstallKeybdHook")))
	{
		// It seems best not to report this warning because a user may want to use partial functionality
		// of a script on Win9x:
		//MsgBox("#InstallKeybdHook is not supported on Windows 95/98/Me.  This line will be ignored.");
		if (!g_os.IsWin9x())
			Hotkey::RequireHook(HOOK_KEYBD);
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#InstallMouseHook")))
	{
		// It seems best not to report this warning because a user may want to use partial functionality
		// of a script on Win9x:
		//MsgBox("#InstallMouseHook is not supported on Windows 95/98/Me.  This line will be ignored.");
		if (!g_os.IsWin9x())
			Hotkey::RequireHook(HOOK_MOUSE);
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#UseHook")))
	{
		g_ForceKeybdHook = !parameter || Line::ConvertOnOff(parameter) != TOGGLED_OFF;
		return CONDITION_TRUE;
	}

	// L4: Handle #if (expression) directive.
	if (IS_DIRECTIVE_MATCH(_T("#If")))
	{
		if (!parameter) // The omission of the parameter indicates that any existing criteria should be turned off.
		{
			g_HotCriterion = NULL; // Indicate that no criteria are in effect for subsequent hotkeys.
			return CONDITION_TRUE;
		}

		// Check for a duplicate #If expression;
		//  - Prevents duplicate hotkeys under separate copies of the same expression.
		//  - Hotkey,If would only be able to select the first expression with the given source code.
		//  - Conserves memory.
		if (g_HotCriterion = FindHotkeyIfExpr(parameter))
			return CONDITION_TRUE;
		
		Func *current_func = g->CurrentFunc;
		g->CurrentFunc = NULL; // Use global scope.
		mNoUpdateLabels = true; // Avoid pointing any pending labels at this line.
		
		if (!ParseAndAddLine(parameter, ACT_HOTKEY_IF, 0, _T(""))) // Pass non-NULL for aActionName.
			return FAIL;
		
		mNoUpdateLabels = false;
		g->CurrentFunc = current_func;
		
		Line *hot_expr_line = mLastLine;

		// Set the new criterion.
		if (  !(g_HotCriterion = AddHotkeyIfExpr())  )
			return FAIL;
		g_HotCriterion->Type = HOT_IF_EXPR;
		g_HotCriterion->ExprLine = hot_expr_line;
		g_HotCriterion->WinTitle = hot_expr_line->mArg[0].text;
		g_HotCriterion->WinText = _T("");
		return CONDITION_TRUE;
	}

	// L4: Allow #if timeout to be adjusted.
	if (IS_DIRECTIVE_MATCH(_T("#IfTimeout")))
	{
		if (parameter)
			g_HotExprTimeout = ATOU(parameter);
		return CONDITION_TRUE;
	}

	if (!_tcsnicmp(aBuf, _T("#IfWin"), 6))
	{
		HotCriterionType hot_criterion;
		bool invert = !_tcsnicmp(aBuf + 6, _T("Not"), 3);
		if (!_tcsnicmp(aBuf + (invert ? 9 : 6), _T("Active"), 6)) // It matches #IfWin[Not]Active.
			hot_criterion = invert ? HOT_IF_NOT_ACTIVE : HOT_IF_ACTIVE;
		else if (!_tcsnicmp(aBuf + (invert ? 9 : 6), _T("Exist"), 5))
			hot_criterion = invert ? HOT_IF_NOT_EXIST : HOT_IF_EXIST;
		else // It starts with #IfWin but isn't Active or Exist: Don't alter g_HotCriterion.
			return CONDITION_FALSE; // Indicate unknown directive since there are currently no other possibilities.
		if (!parameter) // The omission of the parameter indicates that any existing criteria should be turned off.
		{
			g_HotCriterion = NULL; // Indicate that no criteria are in effect for subsequent hotkeys.
			return CONDITION_TRUE;
		}
		LPTSTR hot_win_title = parameter, hot_win_text; // Set default for title; text is determined later.
		// Scan for the first non-escaped comma.  If there is one, it marks the second parameter: WinText.
		LPTSTR cp, first_non_escaped_comma;
		for (first_non_escaped_comma = NULL, cp = hot_win_title; ; ++cp)  // Increment to skip over the symbol just found by the inner for().
		{
			for (; *cp && !(*cp == g_EscapeChar || *cp == g_delimiter || *cp == g_DerefChar); ++cp);  // Find the next escape char, comma, or %.
			if (!*cp) // End of string was found.
				break;
#define ERR_ESCAPED_COMMA_PERCENT _T("Literal commas and percent signs must be escaped (e.g. `%)")
			if (*cp == g_DerefChar)
				return ScriptError(ERR_ESCAPED_COMMA_PERCENT, aBuf);
			if (*cp == g_delimiter) // non-escaped delimiter was found.
			{
				// Preserve the ability to add future-use parameters such as section of window
				// over which the mouse is hovering, e.g. #IfWinActive, Untitled - Notepad,, TitleBar
				if (first_non_escaped_comma) // A second non-escaped comma was found.
					return ScriptError(ERR_ESCAPED_COMMA_PERCENT, aBuf);
				// Otherwise:
				first_non_escaped_comma = cp;
				continue; // Check if there are any more non-escaped commas.
			}
			// Otherwise, an escape character was found, so skip over the next character (if any).
			if (!*(++cp)) // The string unexpectedly ends in an escape character, so avoid out-of-bounds.
				break;
			// Otherwise, the ++cp above has skipped over the escape-char itself, and the loop's ++cp will now
			// skip over the char-to-be-escaped, which is not the one we want (even if it is a comma).
		}
		if (first_non_escaped_comma) // Above found a non-escaped comma, so there is a second parameter (WinText).
		{
			// Omit whitespace to (seems best to conform to convention/expectations rather than give
			// strange whitespace flexibility that would likely cause unwanted bugs due to inadvertently
			// have two spaces instead of one).  The user may use `s and `t to put literal leading/trailing
			// spaces/tabs into these parameters.
			hot_win_text = omit_leading_whitespace(first_non_escaped_comma + 1);
			*first_non_escaped_comma = '\0'; // Terminate at the comma to split off hot_win_title on its own.
			rtrim(hot_win_title, first_non_escaped_comma - hot_win_title);  // Omit whitespace (see similar comment above).
			// The following must be done only after trimming and omitting whitespace above, so that
			// `s and `t can be used to insert leading/trailing spaces/tabs.  ConvertEscapeSequences()
			// also supports insertion of literal commas via escaped sequences.
			ConvertEscapeSequences(hot_win_text, g_EscapeChar, true);
		}
		else
			hot_win_text = _T(""); // And leave hot_win_title set to the entire string because there's only one parameter.
		// The following must be done only after trimming and omitting whitespace above (see similar comment above).
		ConvertEscapeSequences(hot_win_title, g_EscapeChar, true);
		// The following also handles the case where both title and text are blank, which could happen
		// due to something weird but legit like: #IfWinActive, ,
		if (!SetHotkeyCriterion(hot_criterion, hot_win_title, hot_win_text))
			return ScriptError(ERR_OUTOFMEM); // So rare that no second param is provided (since its contents may have been temp-terminated or altered above).
		return CONDITION_TRUE;
	} // Above completely handles all directives and non-directives that start with "#IfWin".

	if (IS_DIRECTIVE_MATCH(_T("#Hotstring")))
	{
		if (parameter)
		{
			LPTSTR suboption = tcscasestr(parameter, _T("EndChars"));
			if (suboption)
			{
				// Since it's not realistic to have only a couple, spaces and literal tabs
				// must be included in between other chars, e.g. `n `t has a space in between.
				// Also, EndChar  \t  will have a space and a tab since there are two spaces
				// after the word EndChar.
				if (    !(parameter = StrChrAny(suboption, _T("\t ")))   )
					return CONDITION_TRUE;
				tcslcpy(g_EndChars, ++parameter, _countof(g_EndChars));
				ConvertEscapeSequences(g_EndChars, g_EscapeChar, false);
				return CONDITION_TRUE;
			}
			if (!_tcsnicmp(parameter, _T("NoMouse"), 7)) // v1.0.42.03
			{
				g_HSResetUponMouseClick = false;
				return CONDITION_TRUE;
			}
			// Otherwise assume it's a list of options.  Note that for compatibility with its
			// other caller, it will stop at end-of-string or ':', whichever comes first.
			Hotstring::ParseOptions(parameter, g_HSPriority, g_HSKeyDelay, g_HSSendMode, g_HSCaseSensitive
				, g_HSConformToCase, g_HSDoBackspace, g_HSOmitEndChar, g_HSSendRaw, g_HSEndCharRequired
				, g_HSDetectWhenInsideWord, g_HSDoReset, g_HSSameLineAction);
		}
		return CONDITION_TRUE;
	}

	if (IS_DIRECTIVE_MATCH(_T("#HotkeyModifierTimeout")))
	{
		if (parameter)
			g_HotkeyModifierTimeout = ATOI(parameter);  // parameter was set to the right position by the above macro
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#HotkeyInterval")))
	{
		if (parameter)
		{
			g_HotkeyThrottleInterval = ATOI(parameter);  // parameter was set to the right position by the above macro
			if (g_HotkeyThrottleInterval < 10) // values under 10 wouldn't be useful due to timer granularity.
				g_HotkeyThrottleInterval = 10;
		}
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#MaxHotkeysPerInterval")))
	{
		if (parameter)
		{
			g_MaxHotkeysPerInterval = ATOI(parameter);  // parameter was set to the right position by the above macro
			if (g_MaxHotkeysPerInterval < 1) // sanity check
				g_MaxHotkeysPerInterval = 1;
		}
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#MaxThreadsPerHotkey")))
	{
		if (parameter)
		{
			// Use value as a temp holder since it's int vs. UCHAR and can thus detect very large or negative values:
			value = ATOI(parameter);  // parameter was set to the right position by the above macro
			if (value > MAX_THREADS_LIMIT) // For now, keep this limited to prevent stack overflow due to too many pseudo-threads.
				value = MAX_THREADS_LIMIT; // UPDATE: To avoid array overflow, this limit must by obeyed except where otherwise documented.
			else if (value < 1)
				value = 1;
			g_MaxThreadsPerHotkey = value; // Note: g_MaxThreadsPerHotkey is UCHAR.
		}
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#MaxThreadsBuffer")))
	{
		g_MaxThreadsBuffer = !parameter || Line::ConvertOnOff(parameter) != TOGGLED_OFF;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#MaxThreads")))
	{
		if (parameter)
		{
			value = ATOI(parameter);  // parameter was set to the right position by the above macro
			if (value > MAX_THREADS_LIMIT) // For now, keep this limited to prevent stack overflow due to too many pseudo-threads.
				value = MAX_THREADS_LIMIT; // UPDATE: To avoid array overflow, this limit must by obeyed except where otherwise documented.
			else if (value < 1)
				value = 1;
			g_MaxThreadsTotal = value;
		}
		return CONDITION_TRUE;
	}

	if (IS_DIRECTIVE_MATCH(_T("#ClipboardTimeout")))
	{
		if (parameter)
			g_ClipboardTimeout = ATOI(parameter);  // parameter was set to the right position by the above macro
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#LTrim")))
	{
		g_ContinuationLTrim = !parameter || Line::ConvertOnOff(parameter) != TOGGLED_OFF;
		return CONDITION_TRUE;
	}

	if (IS_DIRECTIVE_MATCH(_T("#WinActivateForce")))
	{
		g_WinActivateForce = true;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#ErrorStdOut")))
	{
		mErrorStdOut = true;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#MaxMem")))
	{
		if (parameter)
		{
			double valuef = ATOF(parameter);  // parameter was set to the right position by the above macro
			if (valuef > 4095)  // Don't exceed capacity of VarSizeType, which is currently a DWORD (4 gig).
				valuef = 4095;  // Don't use 4096 since that might be a special/reserved value for some functions.
			else if (valuef  < 1)
				valuef = 1;
			g_MaxVarCapacity = (VarSizeType)(valuef * 1024 * 1024);
		}
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#KeyHistory")))
	{
		if (parameter)
		{
			g_MaxHistoryKeys = ATOI(parameter);  // parameter was set to the right position by the above macro
			if (g_MaxHistoryKeys < 0)
				g_MaxHistoryKeys = 0;
			else if (g_MaxHistoryKeys > 500)
				g_MaxHistoryKeys = 500;
			// Above: There are two reasons for limiting the history file to 500 keystrokes:
			// 1) GetHookStatus() only has a limited size buffer in which to transcribe the keystrokes.
			//    500 events is about what you would expect to fit in a 32 KB buffer (it the unlikely event
			//    that the transcribed events create too much text, the text will be truncated, so it's
			//    not dangerous anyway).
			// 2) To reduce the impression that AutoHotkey designed for key logging (the key history file
			//    is in a very unfriendly format that type of key logging anyway).
		}
		return CONDITION_TRUE;
	}

	// For the below series, it seems okay to allow the comment flag to contain other reserved chars,
	// such as DerefChar, since comments are evaluated, and then taken out of the game at an earlier
	// stage than DerefChar and the other special chars.
	if (IS_DIRECTIVE_MATCH(_T("#CommentFlag")))
	{
		if (parameter)
		{
			if (!*(parameter + 1))  // i.e. the length is 1
			{
				// Don't allow '#' since it's the preprocessor directive symbol being used here.
				// Seems ok to allow "." to be the comment flag, since other constraints mandate
				// that at least one space or tab occur to its left for it to be considered a
				// comment marker.
				if (*parameter == '#' || *parameter == g_DerefChar || *parameter == g_EscapeChar || *parameter == g_delimiter)
					return ScriptError(ERR_PARAM1_INVALID, aBuf);
				// Exclude hotkey definition chars, such as ^ and !, because otherwise
				// the following example wouldn't work:
				// User defines ! as the comment flag.
				// The following hotkey would never be in effect since it's considered to
				// be commented out:
				// !^a::run,notepad
				if (*parameter == '!' || *parameter == '^' || *parameter == '+' || *parameter == '$' || *parameter == '~' || *parameter == '*'
					|| *parameter == '<' || *parameter == '>')
					// Note that '#' is already covered by the other stmt. above.
					return ScriptError(ERR_PARAM1_INVALID, aBuf);
			}
			tcslcpy(g_CommentFlag, parameter, MAX_COMMENT_FLAG_LENGTH + 1);
			g_CommentFlagLength = _tcslen(g_CommentFlag);  // Keep this in sync with above.
		}
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#EscapeChar")))
	{
		if (parameter)
		{
			// Don't allow '.' since that can be part of literal floating point numbers:
			if (   *parameter == '#' || *parameter == g_DerefChar || *parameter == g_delimiter || *parameter == '.'
				|| (g_CommentFlagLength == 1 && *parameter == *g_CommentFlag)   )
				return ScriptError(ERR_PARAM1_INVALID, aBuf);
			g_EscapeChar = *parameter;
		}
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#DerefChar")))
	{
		if (parameter)
		{
			if (   *parameter == g_EscapeChar || *parameter == g_delimiter || *parameter == '.'
				|| (g_CommentFlagLength == 1 && *parameter == *g_CommentFlag)   ) // Fix for v1.0.47.05: Allow deref char to be # as documented.
				return ScriptError(ERR_PARAM1_INVALID, aBuf);
			g_DerefChar = *parameter;
		}
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#Delimiter")))
	{
		// Attempts to change the delimiter to its starting default (comma) are ignored.
		// For example, "#Delimiter ," isn't meaningful if the delimiter already is a comma,
		// which is good because "parameter" has already assumed that the comma is accidental
		// (not a symbol) and omitted it.
		if (parameter)
		{
			if (   *parameter == '#' || *parameter == g_EscapeChar || *parameter == g_DerefChar || *parameter == '.'
				|| (g_CommentFlagLength == 1 && *parameter == *g_CommentFlag)   )
				return ScriptError(ERR_PARAM1_INVALID, aBuf);
			g_delimiter = *parameter;
		}
		return CONDITION_TRUE;
	}

	if (IS_DIRECTIVE_MATCH(_T("#MenuMaskKey")))
	{
		// L38: Allow scripts to specify an alternate "masking" key in place of VK_CONTROL.
		if (parameter)
		{
			// Testing shows that sending an event with zero VK but non-zero SC fails to suppress
			// the Start menu (although it does suppress the window menu).  However, checking the
			// validity of the key seems more correct than requiring g_MenuMaskKeyVK != 0, and
			// adds flexibility at very little cost.  Note that this use of TextToVKandSC()'s
			// return value (vs. checking VK|SC) allows vk00sc000 to turn off masking altogether.
			if (TextToVKandSC(parameter, g_MenuMaskKeyVK, g_MenuMaskKeySC))
				return CONDITION_TRUE;
			//else: It's okay that above modified our variables since we're about to exit.
		}
		return ScriptError(parameter ? ERR_PARAM1_INVALID : ERR_PARAM1_REQUIRED, aBuf);
	}
	if (IS_DIRECTIVE_MATCH(_T("#InputLevel")))
	{
		// All hotkeys declared after this directive are assigned the specified InputLevel.
		// Input generated at a given SendLevel can only trigger hotkeys that belong to a
		// lower InputLevel. Hotkeys at the lowest level (0) cannot be triggered by any
		// generated input (the same behavior as AHK versions before this feature).
		// The default level is 0.

		int group = parameter ? ATOI(parameter) : 0;
		if (!SendLevelIsValid(group))
			return ScriptError(ERR_PARAM1_INVALID, aBuf);

		g_InputLevel = group;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#Warn")))
	{
		if (!parameter)
			parameter = _T("");

		LPTSTR param1_end = _tcschr(parameter, g_delimiter);
		LPTSTR param2 = _T("");
		if (param1_end)
		{
			param2 = omit_leading_whitespace(param1_end + 1);
			param1_end = omit_trailing_whitespace(parameter, param1_end - 1);
			param1_end[1] = '\0';
		}

		int i;

		static LPTSTR sWarnTypes[] = { WARN_TYPE_STRINGS };
		WarnType warnType = WARN_ALL; // Set default.
		if (*parameter)
		{
			for (i = 0; ; ++i)
			{
				if (i == _countof(sWarnTypes))
					return ScriptError(ERR_PARAM1_INVALID, aBuf);
				if (!_tcsicmp(parameter, sWarnTypes[i]))
					break;
			}
			warnType = (WarnType)i;
		}

		static LPTSTR sWarnModes[] = { WARN_MODE_STRINGS };
		WarnMode warnMode = WARNMODE_MSGBOX; // Set default.
		if (*param2)
		{
			for (i = 0; ; ++i)
			{
				if (i == _countof(sWarnModes))
					return ScriptError(ERR_PARAM2_INVALID, param2);
				if (!_tcsicmp(param2, sWarnModes[i]))
					break;
			}
			warnMode = (WarnMode)i;
		}

		// The following series of "if" statements was confirmed to produce smaller code
		// than a switch() with a final case WARN_ALL that duplicates all of the assignments:

		if (warnType == WARN_USE_UNSET_LOCAL || warnType == WARN_ALL)
			g_Warn_UseUnsetLocal = warnMode;

		if (warnType == WARN_USE_UNSET_GLOBAL || warnType == WARN_ALL)
			g_Warn_UseUnsetGlobal = warnMode;

		if (warnType == WARN_USE_ENV || warnType == WARN_ALL)
			g_Warn_UseEnv = warnMode;

		if (warnType == WARN_LOCAL_SAME_AS_GLOBAL || warnType == WARN_ALL)
			g_Warn_LocalSameAsGlobal = warnMode;

		if (warnType == WARN_CLASS_OVERWRITE || warnType == WARN_ALL)
			g_Warn_ClassOverwrite = warnMode;

		return CONDITION_TRUE;
	}

	// Otherwise, report that this line isn't a directive:
	return CONDITION_FALSE;
}



void ScriptTimer::Disable()
{
	mEnabled = false;
	--g_script.mTimerEnabledCount;
	if (!g_script.mTimerEnabledCount && !g_nLayersNeedingTimer && !Hotkey::sJoyHotkeyCount)
		KILL_MAIN_TIMER
	// Above: If there are now no enabled timed subroutines, kill the main timer since there's no other
	// reason for it to exist if we're here.   This is because or direct or indirect caller is
	// currently always ExecUntil(), which doesn't need the timer while its running except to
	// support timed subroutines.  UPDATE: The above is faulty; Must also check g_nLayersNeedingTimer
	// because our caller can be one that still needs a timer as proven by this script that
	// hangs otherwise:
	//SetTimer, Test, on 
	//Sleep, 1000 
	//msgbox, done
	//return
	//Test: 
	//SetTimer, Test, off 
	//return
}



ResultType Script::UpdateOrCreateTimer(IObject *aLabel, LPTSTR aPeriod, LPTSTR aPriority, bool aEnable
	, bool aUpdatePriorityOnly)
// Caller should specific a blank aPeriod to prevent the timer's period from being changed
// (i.e. if caller just wants to turn on or off an existing timer).  But if it does this
// for a non-existent timer, that timer will be created with the default period as specified in
// the constructor.
{
	ScriptTimer *timer;
	for (timer = mFirstTimer; timer != NULL; timer = timer->mNextTimer)
		if (timer->mLabel == aLabel) // Match found.
			break;
	bool timer_existed = (timer != NULL);
	if (!timer_existed)  // Create it.
	{
		if (   !(timer = new ScriptTimer(aLabel))   )
			return ScriptError(ERR_OUTOFMEM);
		if (!mFirstTimer)
			mFirstTimer = mLastTimer = timer;
		else
		{
			mLastTimer->mNextTimer = timer;
			// This must be done after the above:
			mLastTimer = timer;
		}
		++mTimerCount;
	}
	// Update its members:
	if (aEnable && !timer->mEnabled) // Must check both or the mTimerEnabledCount below will be wrong.
	{
		// The exception is if the timer already existed but the caller only wanted its priority changed:
		if (!(timer_existed && aUpdatePriorityOnly))
		{
			timer->mEnabled = true;
			++mTimerEnabledCount;
			SET_MAIN_TIMER  // Ensure the API timer is always running when there is at least one enabled timed subroutine.
		}
		//else do nothing, leave it disabled.
	}
	else if (!aEnable && timer->mEnabled) // Must check both or the below count will be wrong.
		timer->Disable();

	aPeriod = omit_leading_whitespace(aPeriod); // This causes A_Space to be treated as "omitted" rather than zero, so may change the behaviour of some poorly-written scripts, but simplifies the check below which allows -0 to work.
	if (*aPeriod) // Caller wanted us to update this member.
	{
		__int64 period = ATOI64(aPeriod);
		if (*aPeriod == '-') // v1.0.46.16: Support negative periods to mean "run only once".
		{
			timer->mRunOnlyOnce = true;
			timer->mPeriod = (DWORD)-period;
		}
		else // Positive number.  v1.0.36.33: Changed from int to DWORD, and ATOI to ATOU, to double its capacity:
		{
			timer->mPeriod = (DWORD)period; // Always use this method & check to retain compatibility with existing scripts.
			timer->mRunOnlyOnce = false;
		}
	}

	if (*aPriority) // Caller wants this member to be changed from its current or default value.
		timer->mPriority = ATOI(aPriority); // Read any float in a runtime variable reference as an int.

	if (!(timer_existed && aUpdatePriorityOnly))
		// Caller relies on us updating mTimeLastRun in this case.  This is done because it's more
		// flexible, e.g. a user might want to create a timer that is triggered 5 seconds from now.
		// In such a case, we don't want the timer's first triggering to occur immediately.
		// Instead, we want it to occur only when the full 5 seconds have elapsed:
		timer->mTimeLastRun = GetTickCount();

    // Below is obsolete, see above for why:
	// We don't have to kill or set the main timer because the only way this function is called
	// is directly from the execution of a script line inside ExecUntil(), in which case:
	// 1) KILL_MAIN_TIMER is never needed because the timer shouldn't exist while in ExecUntil().
	// 2) SET_MAIN_TIMER is never needed because it will be set automatically the next time ExecUntil()
	//    calls MsgSleep().
	return OK;
}



void Script::DeleteTimer(IObject *aLabel)
{
	ScriptTimer *timer, *previous = NULL;
	for (timer = mFirstTimer; timer != NULL; previous = timer, timer = timer->mNextTimer)
		if (timer->mLabel == aLabel) // Match found.
		{
			// Disable it, even if it's not technically being deleted yet.
			if (timer->mEnabled)
				timer->Disable(); // Keeps track of mTimerEnabledCount and whether the main timer is needed.
			if (timer->mExistingThreads) // This condition differs from g->CurrentTimer == timer, which only detects the "top-most" timer.
			{
				if (!aLabel) // Caller requested we delete a previously marked timer which
					continue; // has now finished, but this one hasn't, so keep looking.
				// In this case we can't delete the timer yet, so mark it for later deletion.
				timer->mLabel = NULL;
				// Clearing mLabel:
				//  1) Marks the timer to be deleted after its last thread finishes.
				//  2) Ensures any subsequently created timer will get default settings.
				//  3) Allows the object to be freed before the timer subroutine returns
				//     if all other references to it are released.
				break;
			}
			// Remove it from the list.
			if (previous)
				previous->mNextTimer = timer->mNextTimer;
			else
				mFirstTimer = timer->mNextTimer;
			if (mLastTimer == timer)
				mLastTimer = previous;
			mTimerCount--;
			// Delete the timer, automatically releasing its reference to the object.
			delete timer;
			break;
		}
}



Label *Script::FindLabel(LPTSTR aLabelName)
// Returns the first label whose name matches aLabelName, or NULL if not found.
// v1.0.42: Since duplicates labels are now possible (to support #IfWin variants of a particular
// hotkey or hotstring), callers must be aware that only the first match is returned.
// This helps performance by requiring on average only half the labels to be searched before
// a match is found.
{
	if (!aLabelName || !*aLabelName) return NULL;
	for (Label *label = mFirstLabel; label != NULL; label = label->mNextLabel)
		if (!_tcsicmp(label->mName, aLabelName)) // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales; 3) performance.
			return label; // Match found.
	return NULL; // No match found.
}



IObject *Script::FindCallable(LPTSTR aLabelName, Var *aVar, int aParamCount)
{
	if (aVar && aVar->HasObject())
	{
		IObject *obj = aVar->Object();
		if (Func *func = LabelPtr(obj).ToFunc())
		{
			// It seems worth performing this additional check; without it, registration
			// of this function would appear to succeed but it would never be called.
			// In particular, this will alert the user that a *method* is not a valid
			// callable object (because it requires at least one parameter: "this").
			// For simplicity, the error message will indicate that no label was found.
			if (func->mMinParams > aParamCount)
				return NULL;
		}
		return obj;
	}
	if (*aLabelName)
	{
		if (Label *label = FindLabel(aLabelName))
			return label;
		if (Func *func = FindFunc(aLabelName))
			if (func->mMinParams <= aParamCount) // See comments above.
				return func;
	}
	return NULL;
}



ResultType Script::AddLabel(LPTSTR aLabelName, bool aAllowDupe)
// Returns OK or FAIL.
{
	if (!*aLabelName)
		return FAIL; // For now, silent failure because callers should check this beforehand.
	if (!aAllowDupe && FindLabel(aLabelName)) // Relies on short-circuit boolean order.
		// Don't attempt to dereference "duplicate_label->mJumpToLine because it might not
		// exist yet.  Example:
		// label1:
		// label1:  <-- This would be a dupe-error but it doesn't yet have an mJumpToLine.
		// return
		return ScriptError(_T("Duplicate label."), aLabelName);
	LPTSTR new_name = SimpleHeap::Malloc(aLabelName);
	if (!new_name)
		return FAIL;  // It already displayed the error for us.
	Label *the_new_label = new Label(new_name); // Pass it the dynamic memory area we created.
	if (the_new_label == NULL)
		return ScriptError(ERR_OUTOFMEM);
	the_new_label->mPrevLabel = mLastLabel;  // Whether NULL or not.
	if (mFirstLabel == NULL)
		mFirstLabel = the_new_label;
	else
		mLastLabel->mNextLabel = the_new_label;
	// This must be done after the above:
	mLastLabel = the_new_label;
	if (!_tcsicmp(new_name, _T("OnClipboardChange")))
		mOnClipboardChangeLabel = the_new_label;
	return OK;
}



void Script::RemoveLabel(Label *aLabel)
// Remove a label from the linked list.
// Used by DefineFunc to implement hotkey/hotstring functions.
{
	if (aLabel->mPrevLabel)
		aLabel->mPrevLabel->mNextLabel = aLabel->mNextLabel;
	else
		mFirstLabel = aLabel->mNextLabel;
	if (aLabel->mNextLabel)
		aLabel->mNextLabel->mPrevLabel = aLabel->mPrevLabel;
	else
		mLastLabel = aLabel->mPrevLabel;
}



ResultType Script::ParseAndAddLine(LPTSTR aLineText, ActionTypeType aActionType, ActionTypeType aOldActionType
	, LPTSTR aActionName, LPTSTR aEndMarker, LPTSTR aLiteralMap, size_t aLiteralMapLength)
// Returns OK or FAIL.
// aLineText needs to be a string whose contents are modifiable (though the string won't be made any
// longer than it is now, so it doesn't have to be of size LINE_SIZE). This helps performance by
// allowing the string to be split into sections without having to make temporary copies.
{
#ifdef _DEBUG
	if (!aLineText || !*aLineText)
		return ScriptError(_T("DEBUG: ParseAndAddLine() called incorrectly."));
#endif

	TCHAR action_name[MAX_VAR_NAME_LENGTH + 1], *end_marker;
	if (aActionName) // i.e. this function was called recursively with explicit values for the optional params.
	{
		_tcscpy(action_name, aActionName);
		end_marker = aEndMarker;
	}
	else if (aActionType == ACT_EXPRESSION)
	{
		*action_name = '\0';
		end_marker = NULL; // Indicate that there is no action to mark the end of.
	}
	else // We weren't called recursively from self, nor is it ACT_EXPRESSION, so set action_name and end_marker the normal way.
	{
		for (;;) // A loop with only one iteration so that "break" can be used instead of a lot of nested if's.
		{
			int declare_type;
			LPTSTR cp;
			if (!_tcsnicmp(aLineText, _T("Global"), 6)) // Checked first because it's more common than the others.
			{
				cp = aLineText + 6; // The character after the declaration word.
				declare_type = g->CurrentFunc ? VAR_DECLARE_GLOBAL : VAR_DECLARE_SUPER_GLOBAL;
			}
			else
			{
				if (!g->CurrentFunc) // Not inside a function body, so "Local"/"Static" get no special treatment.
					break;

				if (!_tcsnicmp(aLineText, _T("Local"), 5))
				{
					cp = aLineText + 5; // The character after the declaration word.
					declare_type = VAR_DECLARE_LOCAL;
				}
				else if (!_tcsnicmp(aLineText, _T("Static"), 6)) // Static also implies local (for functions that default to global).
				{
					cp = aLineText + 6; // The character after the declaration word.
					declare_type = VAR_DECLARE_STATIC;
				}
				else // It's not the word "global", "local", or static, so no further checking is done.
					break;
			}

			if (*cp && !IS_SPACE_OR_TAB(*cp)) // There is a character following the word local but it's not a space or tab.
				break; // It doesn't qualify as being the global or local keyword because it's something like global2.
			if (*cp && *(cp = omit_leading_whitespace(cp))) // Probably always a true stmt since caller rtrimmed it, but even if not it's handled correctly.
			{
				// Check whether the first character is an operator by seeing if it alone would be a
				// valid variable name.  If it's not valid, this doesn't qualify as the global or local
				// keyword because it's something like this instead:
				// local := xyz
				// local += 3
				TCHAR orig_char = cp[1];
				cp[1] = '\0'; // Temporarily terminate.
				ResultType result = Var::ValidateName(cp, DISPLAY_NO_ERROR);
				cp[1] = orig_char; // Undo the termination.
				if (!result) // It's probably operator, e.g. local = %var%
					break;
			}
			else // It's the word "global", "local", "static" by itself.
			{
				// All of the following must be checked to catch back-to-back conflicting declarations such
				// as these:
				//    global x
				//    global  ; Should be an error because global vars are implied/automatic.
				// v1.0.48: Lexikos: Added assume-static mode. For now, this requires "static" to be
				// placed above local or global variable declarations.
				if (mNextLineIsFunctionBody)
				{
					if (g->CurrentFunc->mDefaultVarType == VAR_DECLARE_NONE)
					{
						if (declare_type == VAR_DECLARE_LOCAL)
							declare_type |= VAR_FORCE_LOCAL; // v1.1.27: "local" by itself restricts globals to only those declared inside the function.
						g->CurrentFunc->mDefaultVarType = declare_type;
						// No further action is required.
						return OK;
					}
					// v1.1.27: Allow "local" and "static" to be combined, leaving the restrictions on globals in place.
					else if (g->CurrentFunc->mDefaultVarType == (VAR_DECLARE_LOCAL | VAR_FORCE_LOCAL) && declare_type == VAR_DECLARE_STATIC)
					{
						g->CurrentFunc->mDefaultVarType = (VAR_DECLARE_STATIC | VAR_FORCE_LOCAL);
						return OK;
					}
				}
				// Otherwise, it occurs too far down in the body.
				return ScriptError(ERR_UNRECOGNIZED_ACTION, aLineText); // Vague error since so rare.
			}
			if (mNextLineIsFunctionBody && g->CurrentFunc->mDefaultVarType == VAR_DECLARE_NONE)
			{
				// Both of the above must be checked to catch back-to-back conflicting declarations such
				// as these:
				// local x
				// global y  ; Should be an error because global vars are implied/automatic.
				// This line will become first non-directive, non-label line in the function's body.

				// If the first non-directive, non-label line in the function's body contains
				// the "local" keyword, everything inside this function will assume that variables
				// are global unless they are explicitly declared local (this is the opposite of
				// the default).  The converse is also true.
				if (declare_type != VAR_DECLARE_STATIC)
					g->CurrentFunc->mDefaultVarType = declare_type == VAR_DECLARE_LOCAL ? VAR_DECLARE_GLOBAL : VAR_DECLARE_LOCAL;
				// Otherwise, leave it as-is to allow the following:
				// static x
				// local y
			}
			else // Since this isn't the first line of the function's body, mDefaultVarType has already been set permanently.
			{
				if (g->CurrentFunc && declare_type == g->CurrentFunc->mDefaultVarType) // Can't be VAR_DECLARE_NONE at this point.
				{
					// Seems best to flag redundant/unnecessary declarations since they might be an indication
					// to the user that something is being done incorrectly in this function. This errors also
					// remind the user what mode the function is in:
					if (declare_type == VAR_DECLARE_GLOBAL)
						return ScriptError(_T("Global variables must not be declared in this function."), aLineText);
					if (declare_type == VAR_DECLARE_LOCAL)
						return ScriptError(_T("Local variables must not be declared in this function."), aLineText);
					// In assume-static mode, allow declarations in case they contain initializers.
					// Would otherwise lose the ability to "initialize only once upon startup".
					//if (declare_type == VAR_DECLARE_STATIC)
					//	return ScriptError("Static variables must not be declared in this function.", aLineText);
				}
			}
			// Since above didn't break or return, a variable is being declared as an exception to the
			// mode specified by mDefaultVarType.

			bool open_brace_was_added, belongs_to_line_above;
			size_t var_name_length;
			LPTSTR item;

			for (belongs_to_line_above = mLastLine && ACT_IS_LINE_PARENT(mLastLine->mActionType)
				, open_brace_was_added = false, item = cp
				; *item;) // FOR EACH COMMA-SEPARATED ITEM IN THE DECLARATION LIST.
			{
				LPTSTR item_end = StrChrAny(item, _T(", \t=:"));  // Comma, space or tab, equal-sign, colon.
				if (!item_end) // This is probably the last/only variable in the list; e.g. the "x" in "local x"
					item_end = item + _tcslen(item);
				var_name_length = (VarSizeType)(item_end - item);

				Var *var = NULL;
				int i;
				if (g->CurrentFunc)
				{
					for (i = 0; i < g->CurrentFunc->mParamCount; ++i) // Search by name to find both global and local declarations.
						if (!tcslicmp(item, g->CurrentFunc->mParam[i].var->mName, var_name_length))
							return ScriptError(_T("Parameters must not be declared."), item);
					// Detect conflicting declarations:
					var = FindVar(item, var_name_length, NULL, FINDVAR_LOCAL);
					if (var && (var->Scope() & ~VAR_DECLARED) == (declare_type & ~VAR_DECLARED) && declare_type != VAR_DECLARE_STATIC)
						var = NULL; // Allow redeclaration using same scope; e.g. "local x := 1 ... local x := 2" down two different code paths.
					if (!var && declare_type != VAR_DECLARE_GLOBAL)
						for (i = 0; i < g->CurrentFunc->mGlobalVarCount; ++i) // Explicitly search this array vs calling FindVar() in case func is assume-global.
							if (!tcslicmp(g->CurrentFunc->mGlobalVar[i]->mName, item, -1, var_name_length))
							{
								var = g->CurrentFunc->mGlobalVar[i];
								break;
							}
					if (var)
						return ScriptError(var->IsDeclared() ? ERR_DUPLICATE_DECLARATION : _T("Declaration conflicts with existing var."), item);
				}
				
				if (   !(var = FindOrAddVar(item, var_name_length, declare_type))   )
					return FAIL; // It already displayed the error.
				if (var->Type() != VAR_NORMAL || !tcslicmp(item, _T("ErrorLevel"), var_name_length)) // Shouldn't be declared either way (global or local).
					return ScriptError(_T("Built-in variables must not be declared."), item);
				if (declare_type == VAR_DECLARE_GLOBAL) // Can only be true if g->CurrentFunc is non-NULL.
				{
					if (g->CurrentFunc->mGlobalVarCount >= MAX_FUNC_VAR_GLOBALS)
						return ScriptError(_T("Too many declarations."), item); // Short message since it's so unlikely.
					g->CurrentFunc->mGlobalVar[g->CurrentFunc->mGlobalVarCount++] = var;
				}
				else if (declare_type == VAR_DECLARE_SUPER_GLOBAL)
				{
					// Ensure the "declared" and "super-global" flags are set, in case this
					// var was added to the list via a reference prior to the declaration.
					var->Scope() = declare_type;
				}

				item_end = omit_leading_whitespace(item_end); // Move up to the next comma, assignment-op, or '\0'.

				LPTSTR the_operator = item_end;
				bool convert_the_operator;
				switch(*item_end)
				{
				case ',':  // No initializer is present for this variable, so move on to the next one.
					item = omit_leading_whitespace(item_end + 1); // Set "item" for use by the next iteration.
					continue; // No further processing needed below.
				case '\0': // No initializer is present for this variable, so move on to the next one.
					item = item_end; // Set "item" for use by the next iteration.
					continue;
				case ':':
					if (item_end[1] != '=') // Colon with no following '='.
						return ScriptError(ERR_UNRECOGNIZED_ACTION, item); // Vague error since so rare.
					item_end += 2; // Point to the character after the ":=".
					convert_the_operator = false;
					break;
				case '=': // Here '=' is clearly an assignment not a comparison, so further below it will be converted to :=
					++item_end; // Point to the character after the "=".
					convert_the_operator = true;
					break;
				default:
					// L31: This can be reached by something not officially supported like "global var .= value_to_append".
					// In previous versions, convert_the_operator was left uninitialized and whether it would work or be
					// replaced with ".:=" and fail was up to chance.  (Testing showed it failed only in Debug mode.)
					convert_the_operator = false;
				}

				// Since above didn't "continue", this declared variable also has an initializer.
				// Add that initializer as a separate line to be executed at runtime. Separate lines
				// might actually perform better at runtime because most initializers tend to be simple
				// literals or variables that are simplified into non-expressions at runtime. In addition,
				// items without an initializer are omitted, further improving runtime performance.
				// However, the following must be done ONLY after having done the FindOrAddVar()
				// above, since that may have changed this variable to a non-default type (local or global).
				// But what about something like "global x, y=x"? Even that should work as long as x
				// appears in the list prior to initializers that use it.
				// Now, find the comma (or terminator) that marks the end of this sub-statement.
				// The search must exclude commas that are inside quoted/literal strings and those that
				// are inside parentheses (chiefly those of function-calls, but possibly others).

				item_end += FindNextDelimiter(item_end); // FIND THE NEXT "REAL" COMMA (or the end of the string).
				
				// Above has now found the final comma of this sub-statement (or the terminator if there is no comma).
				LPTSTR terminate_here = omit_trailing_whitespace(item, item_end-1) + 1; // v1.0.47.02: Fix the fact that "x=5 , y=6" would preserve the whitespace at the end of "5".  It also fixes wrongly showing a syntax error for things like: static d="xyz"  , e = 5
				TCHAR orig_char = *terminate_here;
				*terminate_here = '\0'; // Temporarily terminate (it might already be the terminator, but that's harmless).

				// PERFORMANCE: As of v1.0.48 (with cached binary numbers and pre-postfixed expressions),
				// assignments of literal integers to variables are up to 10% slower when done as a combined
				// (comma-separated) expression rather than each as a separate line.  However,  this slowness
				// eventually disappears and may even reverse as more and more such expressions are combined
				// into a single expression (e.g. the following is almost the same speed either way:
				// x:=1,y:=22,z:=333,a:=4444,b:=55555).  By contrast, assigning a literal string, another
				// variable, or a complex expression is the opposite: they are always faster when done via
				// commas, and they continue to get faster and faster as more expressions are merged into a
				// single comma-separated expression. In light of this, a future version could combine ONLY
				// those declarations that have initializers into a single comma-separately expression rather
				// than making a separate expression for each.  However, since it's not always faster to do
				// so (e.g. x:=0,y:=1 is faster as separate statements), and since it is somewhat rare to
				// have a long chain of initializers, and since these performance differences are documented,
				// it might not be worth changing.
				LPTSTR line_to_add;
				TCHAR new_buf[LINE_SIZE]; // Declared outside the braces below so that it stays in scope long enough. Using so much stack space here and in caller seems unlikely to affect performance, so _alloca seems unlikely to help.
				if (convert_the_operator) // Convert first '=' in item to be ":=".
				{
					// Prevent any chance of overflow by using new_buf (overflow might otherwise occur in cases
					// such as this sub-statement being the very last one in the declaration list, and being
					// at the limit of the buffer's capacity).
					StrReplace(_tcscpy(new_buf, item), _T("="), _T(":="), SCS_SENSITIVE, 1); // Can't overflow because there's only one replacement and we know item's length can't be that close to the capacity limit.
					line_to_add = new_buf;
				}
				else
					line_to_add = item;
				if (declare_type == VAR_DECLARE_STATIC)
				{
					// Avoid pointing labels or the function's mJumpToLine at a static declaration.
					mNoUpdateLabels = true;
				}
				else if (belongs_to_line_above && !open_brace_was_added) // v1.0.46.01: Put braces to allow initializers to work even directly under an IF/ELSE/LOOP.  Note that the braces aren't added or needed for static initializers.
				{
					if (!AddLine(ACT_BLOCK_BEGIN))
						return FAIL;
					open_brace_was_added = true;
				}
				// Call Parse() vs. AddLine() because it detects and optimizes simple assignments into
				// non-expressions for faster runtime execution.
				if (!ParseAndAddLine(line_to_add)) // For simplicity and maintainability, call self rather than trying to set things up properly to stay in self.
					return FAIL; // Above already displayed the error.
				if (declare_type == VAR_DECLARE_STATIC)
				{
					mNoUpdateLabels = false;
					mLastLine->mAttribute = (AttributeType)mLastLine->mActionType;
					mLastLine->mActionType = ACT_STATIC; // Mark this line for the preparser.
				}

				*terminate_here = orig_char; // Undo the temporary termination.
				// Set "item" for use by the next iteration:
				item = (*item_end == ',') // i.e. it's not the terminator and thus not the final item in the list.
					? omit_leading_whitespace(item_end + 1)
					: item_end; // It's the terminator, so let the loop detect that to finish.
			} // for() each item in the declaration list.
			if (open_brace_was_added)
				if (!AddLine(ACT_BLOCK_END))
					return FAIL;
			return OK;
		} // single-iteration for-loop

		// Since above didn't return, it's not a declaration such as "global MyVar".
		if (   !(end_marker = ParseActionType(action_name, aLineText, true))   )
			return FAIL; // It already displayed the error.
	}
	
	// Above has ensured that end_marker is the address of the last character of the action name,
	// or NULL if there is no action name.
	// Find the arguments (not to be confused with exec_params) of this action, if it has any:
	LPTSTR action_args;
	bool could_be_named_action;
	if (end_marker)
	{
		action_args = omit_leading_whitespace(end_marker + 1);
		// L34: Require that named commands and their args are delimited with a space, tab or comma.
		// Detects errors such as "MsgBox< foo" or "If!foo" and allows things like "control[x]:=y".
		TCHAR end_char = end_marker[1];
		could_be_named_action = (end_char == g_delimiter || !end_char || IS_SPACE_OR_TAB(end_char)
			// Allow If() and While() but something like MsgBox() should always be a function-call:
			|| (end_char == '(' && (!_tcsicmp(action_name, _T("IF")) || !_tcsicmp(action_name, _T("WHILE")))));
	}
	else
	{
		action_args = aLineText;
		// Since this entire line is the action's arg, it can't contain an action name.
		could_be_named_action = false;
	}
	// Now action_args is either the first delimiter or the first parameter (if the optional first
	// delimiter was omitted).
	bool add_openbrace_afterward = false; // v1.0.41: Set default for use in supporting brace in "if (expr) {" and "Loop {".

	if (*action_args == g_delimiter)
	{
		// Since there's a comma, don't change aActionType because if it's ACT_INVALID, it should stay that way
		// so that "something, += 4" is not a valid assignment or other operator, but should still be checked
		// against the list of commands to see if it's something like "MsgBox, += 4" (in this case, a script may
		// use the comma to avoid ambiguity).
		// Find the start of the next token (or its ending delimiter if the token is blank such as ", ,"):
		for (++action_args; IS_SPACE_OR_TAB(*action_args); ++action_args);
	}
	else if (!aActionType && !aOldActionType) // i.e. the caller hasn't yet determined this line's action type.
	{
		if (could_be_named_action && !_tcsicmp(action_name, _T("IF"))) // It's an IF-statement.
		{
			/////////////////////////////////////
			// Detect all types of IF-statements.
			/////////////////////////////////////
			LPTSTR operation, next_word;
			if (   *action_args == '(' // i.e. if (expression)
				|| *action_args == g_DerefChar && IS_SPACE_OR_TAB(action_args[1])   ) // v1.0.48.04: "if % expr" is always an expressions. This check was added to allow lines like "if % IniWinCount = b" to work rather than being misinterpreted as "if var in", "if var is", and possibly other things.  However, "if %var%..." is NOT always an expression because it might be something like: if %A_Index%Array <> unquoted_literal_string
			{
				// To support things like the following, the outermost enclosing parentheses are not removed:
				// if (x < 3) or (x > 6)
				// Also note that although the expression must normally start with an open-parenthesis to be
				// recognized as ACT_IFEXPR, it need not end in a close-paren; e.g. if (x = 1) or !done.
				// If these or any other parentheses are unbalanced, it will caught further below.
				aActionType = ACT_IFEXPR; // Fixed for v1.0.31.01.
			}
			else // Generic or indeterminate IF-statement, so find out what type it is.
			{
				// Skip over the variable name so that the "is" and "is not" operators are properly supported:
				DEFINE_END_FLAGS
				if (operation = StrChrAny(action_args, end_flags))
					operation = omit_leading_whitespace(operation);
				else
					operation = action_args + _tcslen(action_args); // Point it to the NULL terminator instead.

				// v1.0.42: Fix "If not Installed" not be seen as "If var-named-'not' in MatchList", being
				// careful not to break "If NotInstalled in MatchList".  The following are also fixed in
				// a similar way:
				// If not BetweenXXX
				// If not ContainsXXX
				bool first_word_is_not = !_tcsnicmp(action_args, _T("Not"), 3) && _tcschr(end_flags, action_args[3]);

				switch (*operation)
				{
				case '=': // But don't allow == to be "Equals" since the 2nd '=' might be literal.
					aActionType = ACT_IFEQUAL;
					break;
				case '<':
					switch(operation[1])
					{
					// Note: User can use whitespace to differentiate a literal symbol from
					// part of an operator, e.g. if var1 < =  <--- char is literal
					case '=':
						aActionType = ACT_IFLESSOREQUAL;
						operation[1] = ' ';
						break;
					case '>':
						aActionType = ACT_IFNOTEQUAL;
						operation[1] = ' ';
						break;
					default: // i.e. some other character follows '<'
						aActionType = ACT_IFLESS;
					}
					break;
				case '>': // Don't allow >< to be NotEqual since the '<' might be intended as a literal part of an arg.
					if (operation[1] == '=')
					{
						aActionType = ACT_IFGREATEROREQUAL;
						operation[1] = ' '; // Remove it from so that it won't be considered by later parsing.
					}
					else
						aActionType = ACT_IFGREATER;
					break;
				case '!':
					if (operation[1] == '=')
					{
						aActionType = ACT_IFNOTEQUAL;
						operation[1] = ' '; // Remove it from so that it won't be considered by later parsing.
					}
					else
						// To minimize the times where expressions must have an outer set of parentheses,
						// assume all unknown operators are expressions, e.g. "if !var"
						aActionType = ACT_IFEXPR;
					break;
				case 'b': // "Between"
				case 'B':
					// Must fall back to ACT_IFEXPR, otherwise "if not var_name_beginning_with_b" is a syntax error.
					if (first_word_is_not || _tcsnicmp(operation, _T("between"), 7))
						aActionType = ACT_IFEXPR;
					else
					{
						aActionType = ACT_IFBETWEEN;
						// Set things up to be parsed as args further down.  A delimiter is inserted later below:
						tmemset(operation, ' ', 7);
					}
					break;
				case 'c': // "Contains"
				case 'C':
					// Must fall back to ACT_IFEXPR, otherwise "if not var_name_beginning_with_c" is a syntax error.
					if (first_word_is_not || _tcsnicmp(operation, _T("contains"), 8))
						aActionType = ACT_IFEXPR;
					else
					{
						aActionType = ACT_IFCONTAINS;
						// Set things up to be parsed as args further down.  A delimiter is inserted later below:
						tmemset(operation, ' ', 8);
					}
					break;
				case 'i':  // "is" or "is not"
				case 'I':
					switch (ctoupper(operation[1]))
					{
					case 's':  // "IS"
					case 'S':
						if (first_word_is_not)        // v1.0.45: Had forgotten to fix this one with the others,
							aActionType = ACT_IFEXPR; // so now "if not is_something" and "if not is_something()" work.
						else
						{
							next_word = omit_leading_whitespace(operation + 2);
							if (_tcsnicmp(next_word, _T("not"), 3)) // No need to check for whitespace after the word "not" because things like "if var is notxxx" are never valid.
								aActionType = ACT_IFIS;
							else
							{
								aActionType = ACT_IFISNOT;
								// Remove the word "not" to set things up to be parsed as args further down.
								tmemset(next_word, ' ', 3);
							}
							operation[1] = ' '; // Remove the 'S' in "IS".  'I' is replaced with ',' later below.
						}
						break;
					case 'n':  // "IN"
					case 'N':
						if (first_word_is_not)
							aActionType = ACT_IFEXPR;
						else
						{
							aActionType = ACT_IFIN;
							operation[1] = ' '; // Remove the 'N' in "IN".  'I' is replaced with ',' later below.
						}
						break;
					default:
						// v1.0.35.01 It must fall back to ACT_IFEXPR, otherwise "if not var_name_beginning_with_i"
						// is a syntax error.
						aActionType = ACT_IFEXPR;
					} // switch()
					break;
				case 'n':  // It's either "not in", "not between", or "not contains"
				case 'N':
					// Must fall back to ACT_IFEXPR, otherwise "if not var_name_beginning_with_n" is a syntax error.
					if (_tcsnicmp(operation, _T("not"), 3) || !IS_SPACE_OR_TAB(operation[3])) // Fix for v1.0.48: Must also check for whitespace after the word "not" to avoid a syntax error for lines like "if not note".
						aActionType = ACT_IFEXPR;
					else
					{
						// Remove the "NOT" separately in case there is more than one space or tab between
						// it and the following word, e.g. "not   between":
						tmemset(operation, ' ', 3);
						next_word = omit_leading_whitespace(operation + 3);
						if (!_tcsnicmp(next_word, _T("in"), 2))
						{
							aActionType = ACT_IFNOTIN;
							tmemset(next_word, ' ', 2);
						}
						else if (!_tcsnicmp(next_word, _T("between"), 7))
						{
							aActionType = ACT_IFNOTBETWEEN;
							tmemset(next_word, ' ', 7);
						}
						else if (!_tcsnicmp(next_word, _T("contains"), 8))
						{
							aActionType = ACT_IFNOTCONTAINS;
							tmemset(next_word, ' ', 8);
						}
					}
					break;

				default: // To minimize the times where expressions must have an outer set of parentheses, assume all unknown operators are expressions.
					aActionType = ACT_IFEXPR;
				} // switch()
			} // Detection of type of IF-statement.

			if (aActionType == ACT_IFEXPR) // There are various ways above for aActionType to become ACT_IFEXPR.
			{
				// Since this is ACT_IFEXPR, action_args is known not to be an empty string, which is relied on below.
				LPTSTR action_args_last_char = action_args + _tcslen(action_args) - 1; // Shouldn't be a whitespace char since those should already have been removed at an earlier stage.
				if (*action_args_last_char == '{') // This is an if-expression statement with an open-brace on the same line.
				{
					*action_args_last_char = '\0';
					rtrim(action_args, action_args_last_char - action_args);  // Remove the '{' and all its whitespace from further consideration.
					add_openbrace_afterward = true;
				}
			}
			else // It's a IF-statement, but a traditional/non-expression one.
			{
				// Set things up to be parsed as args later on.
				*operation = g_delimiter;
				if (aActionType == ACT_IFBETWEEN || aActionType == ACT_IFNOTBETWEEN)
				{
					// I decided against the syntax "if var between 3,8" because the gain in simplicity
					// and the small avoidance of ambiguity didn't seem worth the cost in terms of readability.
					for (next_word = operation;;)
					{
						if (   !(next_word = tcscasestr(next_word, _T("and")))   )
							return ScriptError(_T("BETWEEN requires the word AND."), aLineText); // Seems too rare a thing to warrant falling back to ACT_IFEXPR for this.
						if (_tcschr(_T(" \t"), *(next_word - 1)) && _tcschr(_T(" \t"), *(next_word + 3)))
						{
							// Since there's a space or tab on both sides, we know this is the correct "and",
							// i.e. not one contained within one of the parameters.  Examples:
							// if var between band and cat  ; Don't falsely detect "band"
							// if var between Andy and David  ; Don't falsely detect "Andy".
							// Replace the word AND with a delimiter so that it will be parsed correctly later:
							*next_word = g_delimiter;
							*(next_word + 1) = ' ';
							*(next_word + 2) = ' ';
							break;
						}
						else
							next_word += 3;  // Skip over this false "and".
					} // for()
				} // ACT_IFBETWEEN
			} // aActionType != ACT_IFEXPR
		}
		else // It isn't an IF-statement, so check for assignments/operators that determine that this line isn't one that starts with a named command.
		{
			//////////////////////////////////////////////////////
			// Detect operators and assignments such as := and +=
			//////////////////////////////////////////////////////
			// This section is done before the section that checks whether action_name is a valid command
			// because it avoids ambiguity in a line such as the following:
			//    Input = test  ; Would otherwise be confused with the Input command.
			// But there may be times when a line like this is used:
			//    MsgBox =  ; i.e. the equals is intended to be the first parameter, not an operator.
			// In the above case, the user can provide the optional comma to avoid the ambiguity:
			//    MsgBox, =
			TCHAR action_args_2nd_char = action_args[1];
			bool convert_pre_inc_or_dec = false; // Set default.

			switch(*action_args)
			{
			case '=': // i.e. var=value (old-style assignment)
				aActionType = ACT_ASSIGN;
				break;
			case ':':
				// v1.0.40: Allow things like "MsgBox :: test" to be valid by insisting that '=' follows ':'.
				if (action_args_2nd_char == '=') // i.e. :=
					aActionType = ACT_ASSIGNEXPR;
				break;
			case '+':
				// Support for ++i (and in the next case, --i).  In these cases, action_name must be either
				// "+" or "-", and the first character of action_args must match it.
				if ((convert_pre_inc_or_dec = action_name[0] == '+' && !action_name[1]) // i.e. the pre-increment operator; e.g. ++index.
					|| action_args_2nd_char == '=') // i.e. x+=y (by contrast, post-increment is recognized only after we check for a command name to cut down on ambiguity).
					aActionType = ACT_ADD;
				break;
			case '-':
				// Do a complete validation/recognition of the operator to allow a line such as the following,
				// which omits the first optional comma, to still be recognized as a command rather than a
				// variable-with-operator:
				// SetBatchLines -1
				if ((convert_pre_inc_or_dec = action_name[0] == '-' && !action_name[1]) // i.e. the pre-decrement operator; e.g. --index.
					|| action_args_2nd_char == '=') // i.e. x-=y  (by contrast, post-decrement is recognized only after we check for a command name to cut down on ambiguity).
					aActionType = ACT_SUB;
				break;
			case '*':
				if (action_args_2nd_char == '=') // i.e. *=
					aActionType = ACT_MULT;
				break;
			case '/':
				if (action_args_2nd_char == '=') // i.e. /=
					aActionType = ACT_DIV;
				// ACT_DIV is different than //= and // because ACT_DIV supports floating point inputs by yielding
				// a floating point result (i.e. it doesn't Floor() the result when the inputs are floats).
				else if (action_args_2nd_char == '/' && action_args[2] == '=') // i.e. //=
					aActionType = ACT_EXPRESSION; // Mark this line as a stand-alone expression.
				break;
			case '|':
			case '&':
			case '^':
				if (action_args_2nd_char == '=') // i.e. .= and |= and &= and ^=
					aActionType = ACT_EXPRESSION; // Mark this line as a stand-alone expression.
				break;
			//case '?': Stand-alone ternary such as true ? fn1() : fn2().  These are rare so are
			// checked later, only after action_name has been checked to see if it's a valid command.
			case '>':
			case '<':
				if (action_args_2nd_char == *action_args && action_args[2] == '=') // i.e. >>= and <<=
					aActionType = ACT_EXPRESSION; // Mark this line as a stand-alone expression.
				break;
			case '.': // L34: Handle dot differently now that dot is considered an action end flag. Detects some more errors and allows some valid expressions which weren't previously allowed.
				if (action_args_2nd_char == '=')
				{	// Concat-assign .=
					aActionType = ACT_EXPRESSION;
				}
				else
				{
					LPTSTR id_begin = action_args + 1;
					LPTSTR cp;
					for (;;) // L35: Loop to fix x.y.z() and similar.
					{
						cp = find_identifier_end(id_begin);
						if (*cp == '(')
						{	// Allow function/method Call as standalone expression.
							aActionType = ACT_EXPRESSION;
							break;
						}
						if (cp == id_begin)
							// No valid identifier, doesn't look like a valid expression.
							break;
						cp = omit_leading_whitespace(cp);
						if (*cp == '[' || !*cp // x.y[z] or x.y
							|| cp[1] == '=' && _tcschr(_T(":+-*/|&^."), cp[0]) // Two-char assignment operator.
							|| cp[1] == cp[0]
								&& (   _tcschr(_T("/<>"), cp[0]) && cp[2] == '=' // //=, <<= or >>=
									|| *cp == '+' || *cp == '-'   )) // x.y++ or x.y--
						{	// Allow Set and bracketed Get as standalone expression.
							aActionType = ACT_EXPRESSION;
							break;
						}
						if (*cp != '.')
							// Must be something which is not allowed as a standalone expression.
							break;
						id_begin = cp + 1;
					}
				}
				break;
			//default: Leave aActionType set to ACT_INVALID. This also covers case '\0' in case that's possible.
			} // switch()

			if (aActionType) // An assignment or other type of action was discovered above.
			{
				if (convert_pre_inc_or_dec) // Set up pre-ops like ++index and --index to be parsed properly later.
				{
					// The following converts:
					// ++x -> EnvAdd x,1 (not really "EnvAdd" per se; but ACT_ADD).
					// Set action_args to be the word that occurs after the ++ or --:
					action_args = omit_leading_whitespace(++action_args); // Though there generally isn't any.
					if (StrChrAny(action_args, EXPR_ALL_SYMBOLS)) // Support things like ++Var ? f1() : f2() and ++Var /= 5. Don't need strstr(action_args, " ?") because the search already looks for ':'.
						aActionType = ACT_EXPRESSION; // Mark this line as a stand-alone expression.
					else
					{
						// Set up aLineText and action_args to be parsed later on as a list of two parameters:
						// The variable name followed by the amount to be added or subtracted (e.g. "ScriptVar, 1").
						// We're not changing the length of aLineText by doing this, so it should be large enough:
						size_t new_length = _tcslen(action_args);
						// Since action_args is just a pointer into the aLineText buffer (which caller has ensured
						// is modifiable), use memmove() so that overlapping source & dest are properly handled:
						tmemmove(aLineText, action_args, new_length + 1); // +1 to include the zero terminator.
						// Append the second param, which is just "1" since the ++ and -- only inc/dec by 1:
						aLineText[new_length++] = g_delimiter;
						aLineText[new_length++] = '1';
						aLineText[new_length] = '\0';
					}
				}
				else if (aActionType != ACT_EXPRESSION) // i.e. it's ACT_ASSIGN/ASSIGNEXPR/ADD/SUB/MULT/DIV
				{
					if (aActionType != ACT_ASSIGN) // i.e. it's ACT_ASSIGNEXPR/ADD/SUB/MULT/DIV
					{
						// Find the first non-function comma, which in the case of ACT_ADD/SUB can be
						// either a statement-separator comma (expression) or the time units arg.
						// Reasons for this:
						// 1) ACT_ADD/SUB: Need to distinguish compound statements from date/time math;
						//    e.g. "x+=1, y+=2" should be marked as a stand-alone expression, not date math.
						// 2) ACT_ASSIGNEXPR/MULT/DIV (and ACT_ADD/SUB for that matter): Need to make
						//    comma-separated sub-expressions into one big ACT_EXPRESSION so that the
						//    leftmost sub-expression will get evaluated prior to the others (for consistency
						//    and as documented).  However, this has some side-effects, such as making
						//    the leftmost /= operator into true division rather than ENV_DIV behavior,
						//    and treating blanks as errors in math expressions when otherwise ENV_MULT
						//    would treat them as zero.
						// ALSO: ACT_ASSIGNEXPR/ADD/SUB/MULT/DIV are made into ACT_EXPRESSION *only* when multi-
						// statement commas are present because the following legacy behaviors must be retained:
						// 1) Math treatment of blanks as zero in ACT_ADD/SUB/etc.
						// 2) EnvDiv's special behavior, which is different than both true divide and floor divide.
						// 3) Possibly add/sub's date/time math.
						// 4) Maybe obsolete: For performance, don't want trivial assignments to become ACT_EXPRESSION.
						LPTSTR cp = action_args + FindNextDelimiter(action_args, g_delimiter, 2);
						if (*cp) // Found a delimiting comma other than one in a sub-statement or function. Shouldn't need to worry about unquoted escaped commas since they don't make sense with += and -=.
						{
							if (aActionType == ACT_ADD || aActionType == ACT_SUB)
							{
								cp = omit_leading_whitespace(cp + 1);
								if (StrChrAny(cp, EXPR_ALL_SYMBOLS)) // Don't need strstr(cp, " ?") because the search already looks for ':'.
									aActionType = ACT_EXPRESSION; // It's clearly an expression not a word like Days or %VarContainingTheWordDays%.
								//else it's probably date/time math, so leave it as-is.
							}
							else // ACT_ASSIGNEXPR/MULT/DIV, for which any non-function comma qualifies it as multi-statement.
								aActionType = ACT_EXPRESSION;
						}
					}
					if (aActionType != ACT_EXPRESSION) // The above didn't make it a stand-alone expression.
					{
						// The following converts:
						// x+=2 -> ACT_ADD x, 2.
						// x:=2 -> ACT_ASSIGNEXPR, x, 2
						// etc.
						// But post-inc/dec are recognized only after we check for a command name to cut down on ambiguity
						*action_args = g_delimiter; // Replace the =,+,-,:,*,/ with a delimiter for later parsing.
						if (aActionType != ACT_ASSIGN) // i.e. it's not just a plain equal-sign (which has no 2nd char).
							action_args[1] = ' '; // Remove the "=" from consideration.
					}
				}
				//else it's already an isolated expression, so no changes are desired.
				action_args = aLineText; // Since this is an assignment and/or expression, use the line's full text for later parsing.
			} // if (aActionType)
		} // Handling of assignments and other operators.
	}
	//else aActionType was already determined by the caller.

	// Now the above has ensured that action_args is the first parameter itself, or empty-string if none.
	// If action_args now starts with a delimiter, it means that the first param is blank/empty.

	if (!aActionType && could_be_named_action && !aOldActionType) // Caller nor logic above has yet determined the action.
		if (   !(aActionType = ConvertActionType(action_name))   ) // Is this line a command?
			aOldActionType = ConvertOldActionType(action_name);    // If not, is it an old-command?

	if (!aActionType && !aOldActionType) // Didn't find any action or command in this line.
	{
		// v1.0.41: Support one-true brace style even if there's no space, but make it strict so that
		// things like "Loop{ string" are reported as errors (in case user intended a file-pattern loop).
		if (*action_args == '{')
		{
			switch (ActionTypeType otb_act = ConvertActionType(action_name))
			{
			case ACT_LOOP:
				add_openbrace_afterward = true;
				if (action_args[1]) // See above.
					break;
				//else fall through:
			case ACT_ELSE:
			case ACT_TRY:
			case ACT_CATCH:
			case ACT_FINALLY:
				aActionType = otb_act;
				break;
			}
		}
		if (!aActionType && _tcschr(EXPR_ALL_SYMBOLS, *action_args))
		{
			LPTSTR question_mark;
			if ((*action_args == '+' || *action_args == '-') && action_args[1] == *action_args) // Post-inc/dec. See comments further below.
			{
				if (action_args[2]) // i.e. if the ++ and -- isn't the last thing; e.g. x++ ? fn1() : fn2() ... Var++ //= 2
					aActionType = ACT_EXPRESSION; // Mark this line as a stand-alone expression.
				else
				{
					// The logic here allows things like IfWinActive-- to be seen as commands even without
					// a space before the -- or ++.  For backward compatibility and code simplicity, it seems
					// best to keep that behavior rather than distinguishing between Command-- and Command --.
					// In any case, "Command --" should continue to be seen as a command regardless of what
					// changes are ever made.  That's why this section occurs below the command-name lookup.
					// The following converts x++ to "ACT_ADD x,1".
					aActionType = (*action_args == '+') ? ACT_ADD : ACT_SUB;
					*action_args = g_delimiter;
					action_args[1] = '1';
				}
				action_args = aLineText; // Since this is an assignment and/or expression, use the line's full text for later parsing.
			}
			else if (*action_args == '?'  // L34: Below no longer requires spaces around '?'.
				|| (question_mark = _tcschr(action_args,'?')) && _tcschr(question_mark,':')) // Rough check (see comments below). Relies on short-circuit boolean order.
			{
				// To avoid hindering load-time error detection such as misspelled command names, allow stand-alone
				// expressions only for things that can produce a side-effect (currently only ternaries like
				// the ones mentioned later below need to be checked since the following other things were
				// previously recognized as ACT_EXPRESSION if appropriate: function-calls, post- and
				// pre-inc/dec (++/--), and assignment operators like := += *= (though these don't necessarily
				// need to be ACT_EXPRESSION to support multi-statement; they can be ACT_ASSIGNEXPR, ACT_ADD, etc.
				// and still support comma-separated statements.
				// Stand-alone ternaries are checked for here rather than earlier to allow a command name
				// (of present) to take precedence (since stand-alone ternaries seem much rarer than 
				// "Command ? something" such as "MsgBox ? something".  Could also check for a colon somewhere
				// to the right if further ambiguity-resolution is ever needed.  Also, a stand-alone ternary
				// should have at least one function-call and/or assignment; otherwise it would serve no purpose.
				// A line may contain a stand-alone ternary operator to call functions that have side-effects
				// or perform assignments.  For example:
				//    IsDone ? fn1() : fn2()
				//    3 > 2 ? x:=1 : y:=1
				//    (3 > 2) ... not supported due to overlap with continuation sections.
				aActionType = ACT_EXPRESSION; // Mark this line as a stand-alone expression.
				action_args = aLineText; // Since this is an assignment and/or expression, use the line's full text for later parsing.
			}
			//else leave it as an unknown action to avoid hindering load-time error detection.
			// In other words, don't be too permissive about what gets marked as a stand-alone expression.
		}
		if (!aActionType) // Above still didn't find a valid action (i.e. check aActionType again in case the above changed it).
		{
			if (*action_args == '(' || *action_args == '[' // v1.0.46.11: Recognize as multi-statements that start with a function, like "fn(), x:=4".  v1.0.47.03: Removed the following check to allow a close-brace to be followed by a comma-less function-call: strchr(action_args, g_delimiter).
				|| *aLineText == '(' // Probably an expression with parentheses to control order of evaluation.
				|| !_tcsnicmp(aLineText, _T("new"), 3) && IS_SPACE_OR_TAB(aLineText[3]))
			{
				aActionType = ACT_EXPRESSION; // Mark this line as a stand-alone expression.
				action_args = aLineText; // Since this is a function-call followed by a comma and some other expression, use the line's full text for later parsing.
			}
			else
				// v1.0.40: Give a more specific error message now that hotkeys can make it here due to
				// the change that avoids the need to escape double-colons:
				return ScriptError(_tcsstr(aLineText, HOTKEY_FLAG) ? _T("Invalid hotkey.") : ERR_UNRECOGNIZED_ACTION, aLineText);
		}
	}
	switch (aActionType)
	{
	case ACT_CATCH:
		if (*action_args != '{') // i.e. "Catch varname" must be handled a different way.
			break;
	case ACT_ELSE:
	case ACT_TRY:
	case ACT_FINALLY:
		if (!AddLine(aActionType))
			return FAIL;
		if (*action_args == '{')
		{
			if (!AddLine(ACT_BLOCK_BEGIN))
				return FAIL;
			action_args = omit_leading_whitespace(action_args + 1);
		}
		if (!*action_args)
			return OK;
		// Call self recursively to parse the sub-action.  Doing this before the args are
		// processed any further avoids some complexity, since literal_map would otherwise
		// have to be passed recursively, in which case the action name is also expected.
		//mCurrLine = NULL; // Seems more useful to leave this set to the line added above.
		return ParseAndAddLine(action_args);
	}

	Action &this_action = aActionType ? g_act[aActionType] : g_old_act[aOldActionType];

	//////////////////////////////////////////////////////////////////////////////////////////////
	// Handle escaped-sequences (escaped delimiters and all others except variable deref symbols).
	// This section must occur after all other changes to the pointer value action_args have
	// occurred above.
	//////////////////////////////////////////////////////////////////////////////////////////////
	// The size of this relies on the fact that caller made sure that aLineText isn't
	// longer than LINE_SIZE.  Also, it seems safer to use char rather than bool, even
	// though on most compilers they're the same size.  Char is always of size 1, but bool
	// can be bigger depending on platform/compiler:
	TCHAR literal_map[LINE_SIZE];
	ZeroMemory(literal_map, sizeof(literal_map));  // Must be fully zeroed for this purpose.
	if (aLiteralMap)
	{
		// Since literal map is NOT a string, just an array of char values, be sure to
		// use memcpy() vs. _tcscpy() on it.  Also, caller's aLiteralMap starts at aEndMarker,
		// so adjust it so that it starts at the newly found position of action_args instead:
		int map_offset = (int)(action_args - end_marker);  // end_marker is known not to be NULL when aLiteralMap is non-NULL.
		int map_length = (int)(aLiteralMapLength - map_offset);
		if (map_length > 0)
			tmemcpy(literal_map, aLiteralMap + map_offset, map_length);
	}
	else
	{
		// Resolve escaped sequences and make a map of which characters in the string should
		// be interpreted literally rather than as their native function.  In other words,
		// convert any escape sequences in order from left to right (this order is important,
		// e.g. ``% should evaluate to `g_DerefChar not `LITERAL_PERCENT.  This part must be
		// done *after* checking for comment-flags that appear to the right of a valid line, above.
		// How literal comment-flags (e.g. semicolons) work:
		//string1; string2 <-- not a problem since string2 won't be considered a comment by the above.
		//string1 ; string2  <-- this would be a user mistake if string2 wasn't supposed to be a comment.
		//string1 `; string 2  <-- since esc seq. is resolved *after* checking for comments, this behaves as intended.
		// Current limitation: a comment-flag longer than 1 can't be escaped, so if "//" were used,
		// as a comment flag, it could never have whitespace to the left of it if it were meant to be literal.
		// Note: This section resolves all escape sequences except those involving g_DerefChar, which
		// are handled by a later section.
		TCHAR c;
		int i;
		for (i = 0; ; ++i)  // Increment to skip over the symbol just found by the inner for().
		{
			for (; action_args[i] && action_args[i] != g_EscapeChar; ++i);  // Find the next escape char.
			if (!action_args[i]) // end of string.
				break;
			c = action_args[i + 1];
			switch (c)
			{
				// Only lowercase is recognized for these:
				case 'a': action_args[i + 1] = '\a'; break;  // alert (bell) character
				case 'b': action_args[i + 1] = '\b'; break;  // backspace
				case 'f': action_args[i + 1] = '\f'; break;  // formfeed
				case 'n': action_args[i + 1] = '\n'; break;  // newline
				case 'r': action_args[i + 1] = '\r'; break;  // carriage return
				case 't': action_args[i + 1] = '\t'; break;  // horizontal tab
				case 'v': action_args[i + 1] = '\v'; break;  // vertical tab
			}
			// Replace escape-sequence with its single-char value.  This is done event if the pair isn't
			// a recognizable escape sequence (e.g. `? becomes ?), which is the Microsoft approach
			// and might not be a bad way of handing things.  There are some exceptions, however.
			// The first of these exceptions (g_DerefChar) is mandatory because that char must be
			// handled at a later stage or escaped g_DerefChars won't work right.  The others are
			// questionable, and might be worth further consideration.  UPDATE: g_DerefChar is now
			// done here because otherwise, examples such as this fail:
			// - The escape char is backslash.
			// - any instances of \\%, such as c:\\%var% , will not work because the first escape
			// sequence (\\) is resolved to a single literal backslash.  But then when \% is encountered
			// by the section that resolves escape sequences for g_DerefChar, the backslash is seen
			// as an escape char rather than a literal backslash, which is not correct.  Thus, we
			// resolve all escapes sequences HERE in one go, from left to right.

			// So these are also done as well, and don't need an explicit check:
			// g_EscapeChar , g_delimiter , (when g_CommentFlagLength > 1 ??): *g_CommentFlag
			// Below has a final +1 to include the terminator:
			tmemcpy(action_args + i, action_args + i + 1, _tcslen(action_args + i + 1) + 1);
			literal_map[i] = 1;  // In the map, mark this char as literal.
		}
	}

	////////////////////////////////////////////////////////////////////////////////////////
	// Do some special preparsing of the MsgBox command, since it is so frequently used and
	// it is also the source of problem areas going from AutoIt2 to 3 and also due to the
	// new numeric parameter at the end.  Whenever possible, we want to avoid the need for
	// the user to have to escape commas that are intended to be literal.
	///////////////////////////////////////////////////////////////////////////////////////
	int mark, max_params_override = 0; // Set default.
	if (aActionType == ACT_MSGBOX)
	{
		for (int next, mark = 0, arg = 1; action_args[mark]; mark = next, ++arg)
		{
			if (arg > 1)
				mark++; // Skip the delimiter...
			while (IS_SPACE_OR_TAB(action_args[mark]))
				mark++; // ...and any leading whitespace.

			if (action_args[mark] == g_DerefChar && !literal_map[mark] && IS_SPACE_OR_TAB(action_args[mark+1]))
			{
				// Since this parameter is an expression, commas inside it can't be intended to be
				// literal/displayed by the user unless they're enclosed in quotes; but in that case,
				// the smartness below isn't needed because it's provided by the parameter-parsing
				// logic in a later section.
				if (arg >= 3) // Text or Timeout
					break;
				// Otherwise, just jump to the next parameter so we can check it too:
				next = FindNextDelimiter(action_args, g_delimiter, mark+2, literal_map);
				continue;
			}
			
			// Find the next non-literal delimiter:
			for (next = mark; action_args[next]; ++next)
				if (action_args[next] == g_delimiter && !literal_map[next])
					break;

			if (arg == 1) // Options (or Text in single-arg mode)
			{
				if (!action_args[next]) // Below relies on this check.
					break; // There's only one parameter, so no further checks are required.
				// It has more than one apparent param, but is the first param even numeric?
				action_args[next] = '\0'; // Temporarily terminate action_args at the first delimiter.
				// Note: If it's a number inside a variable reference, it's still considered 1-parameter
				// mode to avoid ambiguity (unlike the deref check for param #4 in the section below,
				// there seems to be too much ambiguity in this case to justify trying to figure out
				// if the first parameter is a pure deref, and thus that the command should use
				// 3-param or 4-param mode instead).
				if (!IsPureNumeric(action_args + mark)) // No floats allowed.
					max_params_override = 1;
				action_args[next] = g_delimiter; // Restore the string.
				if (max_params_override)
					break;
			}
			else if (arg == 4) // Timeout (or part of Text)
			{
				// If the 4th parameter isn't blank or pure numeric, assume the user didn't intend it
				// to be the MsgBox timeout (since that feature is rarely used), instead intending it
				// to be part of parameter #3.
				if (!IsPureNumeric(action_args + mark, false, true, true))
				{
					// Not blank and not a int or float.  Update for v1.0.20: Check if it's a single
					// deref.  If so, assume that deref contains the timeout and thus 4-param mode is
					// in effect.  This allows the timeout to be contained in a variable, which was
					// requested by one user.  Update for v1.1.06: Do some additional checking to
					// exclude things like "%x% marks the spot" but not "%Timeout%.500".
					LPTSTR deref_end;
					if (action_args[next] // There are too many delimiters (there appears to be another arg after this one).
						|| action_args[mark] != g_DerefChar || literal_map[mark] // Not a proper deref char.
						|| !(deref_end = _tcschr(action_args + mark + 1, g_DerefChar)) // No deref end char (this syntax error will be handled later).
						|| !IsPureNumeric(deref_end + 1, false, true, true)) // There is something non-numeric following the deref (things like %Timeout%.500 are allowed for flexibility and backward-compatibility).
						max_params_override = 3;
				}
				break;
			}
		}
	} // end of special handling for MsgBox.

	else if (aActionType == ACT_FOR)
	{
		// Validate "For" syntax and translate to conventional command syntax.
		// "For x,y in z" -> "For x,y, z"
		// "For x in y"   -> "For x,, y"
		LPTSTR in;
		for (in = action_args; *in; ++in)
			if (IS_SPACE_OR_TAB(*in)
				&& tolower(in[1]) == 'i'
				&& tolower(in[2]) == 'n'
				&& IS_SPACE_OR_TAB(in[3])) // Relies on short-circuit boolean evaluation.
				break;
		if (!*in)
			return ScriptError(_T("This \"For\" is missing its \"in\"."), aLineText);
		int vars = 1;
		for (mark = (int)(in - action_args); mark > 0; --mark)
			if (action_args[mark] == g_delimiter)
				++vars;
		in[1] = g_delimiter; // Replace "in" with a conventional delimiter.
		if (vars > 1)
		{	// Something like "For x,y in z".
			if (vars > 2)
				return ScriptError(_T("Syntax error or too many variables in \"For\" statement."), aLineText);
			in[2] = ' ';
		}
		else
			in[2] = g_delimiter; // Insert another delimiter so the expression is always arg 3.
	}

	/////////////////////////////////////////////////////////////
	// Parse the parameter string into a list of separate params.
	/////////////////////////////////////////////////////////////
	// MaxParams has already been verified as being <= MAX_ARGS.
	// Any g_delimiter-delimited items beyond MaxParams will be included in a lump inside the last param:
	int nArgs;
	LPTSTR arg[MAX_ARGS], arg_map[MAX_ARGS];
	ActionTypeType subaction_type = ACT_INVALID; // Must init these.
	ActionTypeType suboldaction_type = OLD_INVALID;
	TCHAR subaction_name[MAX_VAR_NAME_LENGTH + 1], *subaction_end_marker = NULL, *subaction_start = NULL;
	int max_params = max_params_override ? max_params_override : this_action.MaxParams;
	int max_params_minus_one = max_params - 1;
	bool is_expression;

	for (nArgs = mark = 0; action_args[mark] && nArgs < max_params; ++nArgs)
	{
		if (nArgs == 2) // i.e. the 3rd arg is about to be added.
		{
			switch (aActionType) // will be ACT_INVALID if this_action is an old-style command.
			{
			case ACT_IFWINEXIST:
			case ACT_IFWINNOTEXIST:
			case ACT_IFWINACTIVE:
			case ACT_IFWINNOTACTIVE:
				subaction_start = action_args + mark;
				if (subaction_end_marker = ParseActionType(subaction_name, subaction_start, false))
					if (   !(subaction_type = ConvertActionType(subaction_name))   )
						suboldaction_type = ConvertOldActionType(subaction_name);
				break;
			}
			if (subaction_type || suboldaction_type)
				// A valid command was found (i.e. AutoIt2-style) in place of this commands Exclude Title
				// parameter, so don't add this item as a param to the command.
				break;
		}
		arg[nArgs] = action_args + mark;
		arg_map[nArgs] = literal_map + mark;
		if (nArgs == max_params_minus_one)
		{
			// Don't terminate the last param, just put all the rest of the line
			// into it.  This avoids the need for the user to escape any commas
			// that may appear in the last param.  i.e. any commas beyond this
			// point can't be delimiters because we've already reached MaxArgs
			// for this command:
			++nArgs;
			break;
		}
		// The above does not need the in_quotes and in_parens checks because commas in the last arg
		// are always literal, so there's no problem even in expressions.

		// The following implements the "% " prefix as a means of forcing an expression:
		is_expression = *arg[nArgs] == g_DerefChar && !*arg_map[nArgs] // It's a non-literal deref character.
			&& IS_SPACE_OR_TAB(arg[nArgs][1]); // Followed by a space or tab.

		if (!is_expression)
		{
			// v1.0.43.07: Fixed below to use this_action instead of g_act[aActionType] so that the
			// numeric params of legacy commands like EnvAdd/Sub/LeftClick can be detected.  Without
			// this fix, the last comma in a line like "EnvSub, var, Add(2, 3)" is seen as a parameter
			// delimiter, which causes a loadtime syntax error.
			is_expression = ArgIsNumeric(aActionType, this_action.NumericParams, arg, nArgs);
		}

		// Find the end of the above arg:
		if (is_expression)
			// Find the next delimiter, taking into account quote marks, parentheses, etc.
			mark = FindNextDelimiter(action_args, g_delimiter, mark, literal_map);
		else
			// Find the next non-literal delimiter.
			for (; action_args[mark]; ++mark)
				if (action_args[mark] == g_delimiter && !literal_map[mark])
					break;

		if (action_args[mark])  // A non-literal delimiter was found.
		{
			action_args[mark] = '\0';  // Terminate the previous arg.
			// Trim any whitespace from the previous arg.  This operation will not alter the contents
			// of anything beyond action_args[i], so it should be safe.  In addition, even though it
			// changes the contents of the arg[nArgs] substring, we don't have to update literal_map
			// because the map is still accurate due to the nature of rtrim).  UPDATE: Note that this
			// version of rtrim() specifically avoids trimming newline characters, since the user may
			// have included literal newlines at the end of the string by using an escape sequence:
			rtrim_literal(arg[nArgs], arg_map[nArgs]);
			// Omit the leading whitespace from the next arg:
			for (++mark; IS_SPACE_OR_TAB(action_args[mark]) && !literal_map[mark]; ++mark);
			// Now <mark> marks the end of the string, the start of the next arg,
			// or a delimiter-char (if the next arg is blank).
		}
	}

	///////////////////////////////////////////////////////////////////////////////
	// Ensure there are sufficient parameters for this command.  Note: If MinParams
	// is greater than 0, the param numbers 1 through MinParams are required to be
	// non-blank.
	///////////////////////////////////////////////////////////////////////////////
	TCHAR error_msg[1024];
	if (nArgs < this_action.MinParams)
	{
		sntprintf(error_msg, _countof(error_msg), _T("\"%s\" requires at least %d parameter%s.")
			, this_action.Name, this_action.MinParams
			, this_action.MinParams > 1 ? _T("s") : _T(""));
		return ScriptError(error_msg, aLineText);
	}
	for (int i = 0; i < this_action.MinParams; ++i) // It's only safe to do this after the above.
		if (!*arg[i])
		{
			sntprintf(error_msg, _countof(error_msg), _T("\"%s\" requires that parameter #%u be non-blank.")
				, this_action.Name, i + 1);
			return ScriptError(error_msg, aLineText);
		}

	////////////////////////////////////////////////////////////////////////
	// Handle legacy commands that are supported for backward compatibility.
	////////////////////////////////////////////////////////////////////////
	if (aOldActionType)
	{
		switch(aOldActionType)
		{
		case OLD_WINGETACTIVETITLE:
			arg[nArgs] = _T("A");  arg_map[nArgs] = NULL; // "A" signifies the active window.
			++nArgs;
			return AddLine(ACT_WINGETTITLE, arg, nArgs, arg_map);
		case OLD_WINGETACTIVESTATS:
		{
			// Convert OLD_WINGETACTIVESTATS into *two* new commands:
			// Command #1: WinGetTitle, OutputVar, A
			LPTSTR width = arg[1];  // Temporary placeholder.
			arg[1] = _T("A");  arg_map[1] = NULL;  // Signifies the active window.
			if (!AddLine(ACT_WINGETTITLE, arg, 2, arg_map))
				return FAIL;
			// Command #2: WinGetPos, XPos, YPos, Width, Height, A
			// Reassign args in the new command's ordering.  These lines must occur
			// in this exact order for the copy to work properly:
			arg[0] = arg[3];  arg_map[0] = arg_map[3];  // xpos
			arg[3] = arg[2];  arg_map[3] = arg_map[2];  // height
			arg[2] = width;   arg_map[2] = arg_map[1];  // width
			arg[1] = arg[4];  arg_map[1] = arg_map[4];  // ypos
			arg[4] = _T("A");  arg_map[4] = NULL;  // "A" signifies the active window.
			return AddLine(ACT_WINGETPOS, arg, 5, arg_map);
		}

		case OLD_SETENV:
			return AddLine(ACT_ASSIGN, arg, nArgs, arg_map);
		case OLD_ENVADD:
			return AddLine(ACT_ADD, arg, nArgs, arg_map);
		case OLD_ENVSUB:
			return AddLine(ACT_SUB, arg, nArgs, arg_map);
		case OLD_ENVMULT:
			return AddLine(ACT_MULT, arg, nArgs, arg_map);
		case OLD_ENVDIV:
			return AddLine(ACT_DIV, arg, nArgs, arg_map);

		// For these, break rather than return so that further processing can be done:
		case OLD_IFEQUAL:
			aActionType = ACT_IFEQUAL;
			break;
		case OLD_IFNOTEQUAL:
			aActionType = ACT_IFNOTEQUAL;
			break;
		case OLD_IFGREATER:
			aActionType = ACT_IFGREATER;
			break;
		case OLD_IFGREATEROREQUAL:
			aActionType = ACT_IFGREATEROREQUAL;
			break;
		case OLD_IFLESS:
			aActionType = ACT_IFLESS;
			break;
		case OLD_IFLESSOREQUAL:
			aActionType = ACT_IFLESSOREQUAL;
			break;
#ifdef _DEBUG
		default:
			return ScriptError(_T("DEBUG: Unhandled Old-Command."), action_name);
#endif
		} // switch()
	}

	//////////////////////////////////////////////////////////////////////////////////////////////////
	// Handle AutoIt2-style IF-statements (i.e. the IF's action is on the same line as the condition).
	//////////////////////////////////////////////////////////////////////////////////////////////////
	// The check below: Don't bother if this IF (e.g. IfWinActive) has zero params or if the
	// subaction was already found above:
	if (nArgs && !subaction_type && !suboldaction_type && ACT_IS_IF_OLD(aActionType, aOldActionType))
	{
		LPTSTR delimiter;
		LPTSTR last_arg = arg[nArgs - 1];
		for (mark = (int)(last_arg - action_args); action_args[mark]; ++mark)
		{
			if (action_args[mark] == g_delimiter && !literal_map[mark])  // Match found: a non-literal delimiter.
			{
				delimiter = action_args + mark; // save the location of this delimiter
				// Omit the leading whitespace from the next arg:
				for (++mark; IS_SPACE_OR_TAB(action_args[mark]); ++mark);
				// Now <mark> marks the end of the string, the start of the next arg,
				// or a delimiter-char (if the next arg is blank).
				subaction_start = action_args + mark;
				if (subaction_end_marker = ParseActionType(subaction_name, subaction_start, false))
				{
					if (   !(subaction_type = ConvertActionType(subaction_name))   )
						suboldaction_type = ConvertOldActionType(subaction_name);
					if (subaction_type || suboldaction_type) // A valid sub-action (command) was found.
					{
						// Remove this subaction from its parent line; we want it separate:
						*delimiter = '\0';
						rtrim(last_arg);
					}
					// else leave it as-is, i.e. as part of the last param, because the delimiter
					// found above is probably being used as a literal char even though it isn't
					// escaped, e.g. "ifequal, var1, string with embedded, but non-escaped, commas"
				}
				// else, do nothing; reasoning perhaps similar to above comment.
				break;
			}
		}
	}

	// In v1.0.41, the following one-true-brace styles are also supported:
	// Loop {   ; Known limitation: Overlaps with file-pattern loop that retrieves single file of name "{".
	// Loop 5 { ; Also overlaps, this time with file-pattern loop that retrieves numeric filename ending in '{'.
	// Loop %Var% {  ; Similar, but like the above seems acceptable given extreme rarity of user intending a file pattern.
	if ((aActionType == ACT_LOOP || aActionType == ACT_WHILE) && nArgs == 1 && arg[0][0] // A loop with exactly one, non-blank arg.
		|| ((aActionType == ACT_FOR || aActionType == ACT_CATCH) && nArgs))
	{
		LPTSTR arg1 = arg[nArgs - 1]; // For readability and possibly performance.
		// A loop with the above criteria (exactly one arg) can only validly be a normal/counting loop or
		// a file-pattern loop if its parameter's last character is '{'.  For the following reasons, any
		// single-parameter loop that ends in '{' is considered to be one-true brace:
		// 1) Extremely rare that a file-pattern loop such as "Loop filename {" would ever be used,
		//    and even if it is, the syntax checker will report an unclosed block, making it apparent
		//    to the user that a workaround is needed, such as putting the filename into a variable first.
		// 2) Difficulty and code size of distinguishing all possible valid-one-true-braces from those
		//    that aren't.  For example, the following are ambiguous, so it seems best for consistency
		//    and code size reduction just to treat them as one-truce-brace, which will immediately alert
		//    the user if the brace isn't closed:
		//    a) Loop % (expression) {   ; Ambiguous because expression could resolve to a string, thus it would be seen as a file-pattern loop.
		//    b) Loop %Var% {            ; Similar as above, which means all three of these unintentionally support
		//    c) Loop filename{          ; OTB for some types of file loops because it's not worth the code size to "unsupport" them.
		//    d) Loop *.txt {            ; Like the above: Unintentionally supported, but not documented.
		//    e) (While-loops are also checked here now)
		// Insist that no characters follow the '{' in case the user intended it to be a file-pattern loop
		// such as "Loop {literal-filename".
		LPTSTR arg1_last_char = arg1 + _tcslen(arg1) - 1;
		if (*arg1_last_char == '{')
		{
			add_openbrace_afterward = true;
			*arg1_last_char = '\0';  // Since it will be fully handled here, remove the brace from further consideration.
			if (!rtrim(arg1)) // Trimmed down to nothing, so only a brace was present: remove the arg completely.
				if (aActionType == ACT_LOOP || aActionType == ACT_CATCH)
					nArgs = 0;    // This makes later stages recognize it as an infinite loop rather than a zero-iteration loop.
				else // ACT_WHILE or ACT_FOR
					return ScriptError(ERR_PARAM1_REQUIRED, aLineText);
		}
	}

	if (!AddLine(aActionType, arg, nArgs, arg_map))
		return FAIL;
	if (add_openbrace_afterward)
		if (!AddLine(ACT_BLOCK_BEGIN))
			return FAIL;
	if (!subaction_type && !suboldaction_type) // There is no subaction in this case.
		return OK;
	// Otherwise, recursively add the subaction, and any subactions it might have, beneath
	// the line just added.  The following example:
	// IfWinExist, x, y, IfWinNotExist, a, b, Gosub, Sub1
	// would break down into these lines:
	// IfWinExist, x, y
	//    IfWinNotExist, a, b
	//       Gosub, Sub1
	return ParseAndAddLine(subaction_start, subaction_type, suboldaction_type, subaction_name, subaction_end_marker
		, literal_map + (subaction_end_marker - action_args) // Pass only the relevant substring of literal_map.
		, _tcslen(subaction_end_marker));
}



inline LPTSTR Script::ParseActionType(LPTSTR aBufTarget, LPTSTR aBufSource, bool aDisplayErrors)
// inline since it's called so often.
// aBufTarget should be at least MAX_VAR_NAME_LENGTH + 1 in size.
// Returns NULL if a failure condition occurs; otherwise, the address of the last
// character of the action name in aBufSource.
{
	////////////////////////////////////////////////////////
	// Find the action name and the start of the param list.
	////////////////////////////////////////////////////////
	// Allows the delimiter between action-type-name and the first param to be optional by
	// relying on the fact that action-type-names can't contain spaces. Find first char in
	// aLineText that is a space, a delimiter, or a tab. Also search for operator symbols
	// so that assignments and IFs without whitespace are supported, e.g. var1=5,
	// if var2<%var3%.  Not static in case g_delimiter is allowed to vary:
	DEFINE_END_FLAGS
	LPTSTR end_marker = StrChrAny(aBufSource, end_flags);
	if (end_marker) // Found a delimiter.
	{
		if (end_marker > aBufSource) // The delimiter isn't very first char in aBufSource.
			--end_marker;
		// else we allow it to be the first char to support "++i" etc.
	}
	else // No delimiter found, so set end_marker to the location of the last char in string.
		end_marker = aBufSource + _tcslen(aBufSource) - 1;
	// Now end_marker is the character just prior to the first delimiter or whitespace,
	// or (in the case of ++ and --) the first delimiter itself.  Find the end of
	// the action-type name by omitting trailing whitespace:
	end_marker = omit_trailing_whitespace(aBufSource, end_marker);
	// If first char in aBufSource is a delimiter, action_name will consist of just that first char:
	size_t action_name_length = end_marker - aBufSource + 1;
	if (action_name_length > MAX_VAR_NAME_LENGTH)
	{
		if (aDisplayErrors)
			ScriptError(ERR_UNRECOGNIZED_ACTION, aBufSource); // Short/vague message since so rare.
		return NULL;
	}
	tcslcpy(aBufTarget, aBufSource, action_name_length + 1);
	return end_marker;
}



inline ActionTypeType Script::ConvertActionType(LPTSTR aActionTypeString)
// inline since it's called so often, but don't keep it in the .h due to #include issues.
{
	// For the loop's index:
	// Use an int rather than ActionTypeType since it's sure to be large enough to go beyond
	// 256 if there happen to be exactly 256 actions in the array:
 	for (int action_type = ACT_FIRST_COMMAND; action_type < g_ActionCount; ++action_type)
		if (!_tcsicmp(aActionTypeString, g_act[action_type].Name)) // Match found.
			return action_type;
	return ACT_INVALID;  // On failure to find a match.
}



inline ActionTypeType Script::ConvertOldActionType(LPTSTR aActionTypeString)
// inline since it's called so often, but don't keep it in the .h due to #include issues.
{
 	for (int action_type = OLD_INVALID + 1; action_type < g_OldActionCount; ++action_type)
		if (!_tcsicmp(aActionTypeString, g_old_act[action_type].Name)) // Match found.
			return action_type;
	return OLD_INVALID;  // On failure to find a match.
}



bool Script::ArgIsNumeric(ActionTypeType aActionType, ActionTypeType *np, LPTSTR arg[], int nArgs, int aArgCount)
{
	// As of v1.0.25, pure numeric parameters can optionally be numeric expressions, so check for that:
	int nArgs_plus_one = nArgs + 1;
	for (; *np; ++np)
		if (*np == nArgs_plus_one) // This arg is enforced to be purely numeric.
			break;
	if (*np) // Match found, so this is a purely numeric arg.
	{
		if (aActionType == ACT_WINMOVE)
		{
			if (nArgs > 1)
			{
				// i indicates this is Arg #3 or beyond, which is one of the args that is
				// either the word "default" or a number/expression.
				if (!_tcsicmp(arg[nArgs], _T("default"))) // It's not an expression.
					return false;
			}
			else // This is the first or second arg, which are title/text vs. X/Y when aArgCount > 2.
				if (aArgCount > 2) // Title/text are not numeric/expressions.
					return false;
		}
		return true;
	}
	if (aActionType == ACT_TRANSFORM && (nArgs == 2 || nArgs == 3)) // i.e. the 3rd or 4th arg is about to be added.
	{
		// Somewhat inefficient since it has to be called for both Arg#2 and Arg#3, but seems
		// worth the benefit in terms of code simplification.  Note that the following might
		// return TRANS_CMD_INVALID just because the sub-command is contained in a variable
		// reference.  That is why TRANS_CMD_INVALID does not produce an error at this stage,
		// but only later when the line has been constructed far enough to call ArgHasDeref():
		// i.e. Not the first param, only the third and fourth, which currently are either both numeric or both non-numeric for all cases.
		switch (Line::ConvertTransformCmd(arg[1])) // arg[1] is the second arg.
		{
		// See comment above for why TRANS_CMD_INVALID isn't yet reported as an error:
		case TRANS_CMD_INVALID:
		case TRANS_CMD_ASC:
#ifndef UNICODE
		case TRANS_CMD_UNICODE:
#endif
		case TRANS_CMD_DEREF:
		case TRANS_CMD_HTML:
			break; // Do nothing.  Return default of false.

		default:
			// For all other sub-commands, Arg #3 and #4 are expression-capable.
			return true;
		}
	}
	return false;
}



bool LegacyArgIsExpression(LPTSTR aArgText, LPTSTR aArgMap)
// Helper function for AddLine
{
	// The section below is here in light of rare legacy cases such as the below:
	// -%y%   ; i.e. make it negative.
	// +%y%   ; might happen with up/down adjustments on SoundSet, GuiControl progress/slider, etc?
	// Although the above are detected as non-expressions and thus non-double-derefs,
	// the following are not because they're too rare or would sacrifice too much flexibility:
	// 1%y%.0 ; i.e. at a tens/hundreds place and make it a floating point.  In addition,
	//          1%y% could be an array, so best not to tag that as non-expression.
	//          For that matter, %y%.0 could be an obscure kind of reverse-notation array itself.
	//          However, as of v1.0.29, things like %y%000 are allowed, e.g. Sleep %Seconds%000
	// 0x%y%  ; i.e. make it hex (too rare to check for, plus it could be an array).
	// %y%%z% ; i.e. concatenate two numbers to make a larger number (too rare to check for)
	LPTSTR cp = aArgText + (*aArgText == '-' || *aArgText == '+'); // i.e. +1 if second term evaluates to true.
	return *cp != g_DerefChar // If no deref, for simplicity assume it's an expression since any such non-numeric item would be extremely rare in pre-expression era.
		|| !aArgMap || *(aArgMap + (cp != aArgText)) // There's no literal-map or this deref char is not really a deref char because it's marked as a literal.
		|| !(cp = _tcschr(cp + 1, g_DerefChar)) // There is no next deref char.
		|| (cp[1] && !IsPureNumeric(cp + 1, false, true, true)); // But that next deref char is not the last char, which means this is not a single isolated deref. v1.0.29: Allow things like Sleep %Var%000.
		// Above does not need to check whether last deref char is marked literal in the
		// arg map because if it is, it would mean the first deref char lacks a matching
		// close-symbol, which will be caught as a syntax error below regardless of whether
		// this is an expression.
}



ResultType Script::AddLine(ActionTypeType aActionType, LPTSTR aArg[], int aArgc, LPTSTR aArgMap[])
// aArg must be a collection of pointers to memory areas that are modifiable, and there
// must be at least aArgc number of pointers in the aArg array.  In v1.0.40, a caller (namely
// the "macro expansion" for remappings such as "a::b") is allowed to pass a non-NULL value for
// aArg but a NULL value for aArgMap.
// Returns OK or FAIL.
{
#ifdef _DEBUG
	if (aActionType == ACT_INVALID)
		return ScriptError(_T("DEBUG: BAD AddLine"), aArgc > 0 ? aArg[0] : _T(""));
#endif

	bool do_update_labels;
	if (aArgc >= UCHAR_MAX) // Special signal from caller to avoid pointing any pending labels to this particular line.
	{
		aArgc -= UCHAR_MAX;
		do_update_labels = false;
	}
	else
		do_update_labels = !mNoUpdateLabels;

	Var *target_var;
	DerefType deref[MAX_DEREFS_PER_ARG];  // Will be used to temporarily store the var-deref locations in each arg.
	int deref_count;  // How many items are in deref array.
	ArgStruct *new_arg;  // We will allocate some dynamic memory for this, then hang it onto the new line.
	size_t operand_length;
	TCHAR orig_char;
	LPTSTR op_begin, op_end;
	LPTSTR this_aArgMap, this_aArg, cp;
	bool is_function, pending_function_is_new_op = false;

	//////////////////////////////////////////////////////////
	// Build the new arg list in dynamic memory.
	// The allocated structs will be attached to the new line.
	//////////////////////////////////////////////////////////
	if (!aArgc)
		new_arg = NULL;  // Just need an empty array in this case.
	else
	{
		if (   !(new_arg = (ArgStruct *)SimpleHeap::Malloc(aArgc * sizeof(ArgStruct)))   )
			return ScriptError(ERR_OUTOFMEM);

		int i, j;

		for (i = 0; i < aArgc; ++i)
		{
			////////////////
			// FOR EACH ARG:
			////////////////
			this_aArg = aArg[i];                        // For performance and convenience.
			this_aArgMap = aArgMap ? aArgMap[i] : NULL; // Same.
			ArgStruct &this_new_arg = new_arg[i];       // Same.
			this_new_arg.is_expression = false;         // Set default early, for maintainability.

			// Before allocating memory for this Arg's text, first check if it's a pure
			// variable.  If it is, we store it differently (and there's no need to resolve
			// escape sequences in these cases, since var names can't contain them):
			if (aActionType == ACT_LOOP && i == 1 && aArg[0] && !_tcsicmp(aArg[0], _T("Parse"))) // Verified.
				// i==1 --> 2nd arg's type is based on 1st arg's text.
				this_new_arg.type = ARG_TYPE_INPUT_VAR;
			else
				this_new_arg.type = Line::ArgIsVar(aActionType, i);

			// v1.0.29: Allow expressions in any parameter that starts with % followed by a space
			// or tab. This should be unambiguous because spaces and tabs are illegal in variable names.
			// Since there's little if any benefit to allowing input and output variables to be
			// dynamically built via expression, for now it is disallowed.  If ever allow it,
			// need to review other sections to ensure they will tolerate it.  Also, the following
			// would probably need revision to get it to be detected as an output-variable:
			// % Array%i% = value
			if (*this_aArg == g_DerefChar && !(this_aArgMap && *this_aArgMap) // It's a non-literal deref character.
				&& IS_SPACE_OR_TAB(this_aArg[1])) // Followed by a space or tab.
			{
				if (this_new_arg.type == ARG_TYPE_OUTPUT_VAR // Command requires a variable, not an expression.
					|| this_new_arg.type == ARG_TYPE_INPUT_VAR // Exclude NORMAL args from the check below.
						&& (aActionType == ACT_SORT || ACT_IS_IF(aActionType))) // Sort and If commands depend on arg 0 being a variable.
					return ScriptError(_T("Unexpected %"), this_aArg); // Short message since it's rare.
				this_new_arg.type = ARG_TYPE_NORMAL; // If this was an input var, make it a normal expression.
				this_new_arg.is_expression = true;
				// Omit the percent sign and the space after it from further consideration.
				this_aArg += 2;
				if (this_aArgMap)
					this_aArgMap += 2;
				// ACT_ASSIGN isn't capable of dealing with expressions because ExecUntil() does not
				// call ExpandArgs() automatically for it.  Thus its function, PerformAssign(), would
				// not be given the expanded result of the expression.
				if (aActionType == ACT_ASSIGN)
					aActionType = ACT_ASSIGNEXPR;
			}

			// Since some vars are optional, the below allows them all to be blank or
			// not present in the arg list.  If a mandatory var is blank at this stage,
			// it's okay because all mandatory args are validated to be non-blank elsewhere:
			if (this_new_arg.type != ARG_TYPE_NORMAL)
			{
				if (!*this_aArg)
					// An optional input or output variable has been omitted, so indicate
					// that this arg is not a variable, just a normal empty arg.  Functions
					// such as ListLines() rely on this having been done because they assume,
					// for performance reasons, that args marked as variables really are
					// variables.  In addition, ExpandArgs() relies on this having been done
					// as does the load-time validation for ACT_DRIVEGET:
					this_new_arg.type = ARG_TYPE_NORMAL;
				else
				{
					// Does this input or output variable contain a dereference?  If so, it must
					// be resolved at runtime (to support arrays, etc.).
					// Find the first non-escaped dereference symbol:
					for (j = 0; this_aArg[j] && (this_aArg[j] != g_DerefChar || (this_aArgMap && this_aArgMap[j])); ++j);
					if (!this_aArg[j])
					{
						// A non-escaped deref symbol wasn't found, therefore this variable does not
						// appear to be something that must be resolved dynamically at runtime.
						if (   !(target_var = FindOrAddVar(this_aArg))   )
							return FAIL;  // The above already displayed the error.
						// If this action type is something that modifies the contents of the var, ensure the var
						// isn't a special/reserved one:
						if (this_new_arg.type == ARG_TYPE_OUTPUT_VAR && VAR_IS_READONLY(*target_var))
							return ScriptError(ERR_VAR_IS_READONLY, this_aArg);
						// Rather than removing this arg from the list altogether -- which would disturb
						// the ordering and hurt the maintainability of the code -- the next best thing
						// in terms of saving memory is to store an empty string in place of the arg's
						// text if that arg is a pure variable (i.e. since the name of the variable is already
						// stored in the Var object, we don't need to store it twice):
						this_new_arg.text = _T("");
						this_new_arg.length = 0;
						this_new_arg.deref = (DerefType *)target_var;
						continue;
					}
					// else continue on to the below so that this input or output variable name's dynamic part
					// (e.g. array%i%) can be partially resolved.
				}
			}

			// Below will set the new var to be the constant empty string if the
			// source var is NULL or blank.
			// e.g. If WinTitle is unspecified (blank), but WinText is non-blank.
			// Using empty string is much safer than NULL because these args
			// will be frequently accessed by various functions that might
			// not be equipped to handle NULLs.  Rather than having to remember
			// to check for NULL in every such case, setting it to a constant
			// empty string here should make things a lot more maintainable
			// and less bug-prone.  If there's ever a need for the contents
			// of this_new_arg to be modifiable (perhaps some obscure API calls require
			// modifiable strings?) can malloc a single-char to contain the empty string.
			// 
			// So that it can be passed to Malloc(), first update the length to match what the text will be
			// (if the alloc fails, an inaccurate length won't matter because it's an program-abort situation).
			// The length must fit into a WORD, which it will since each arg is literal text from a script's line,
			// which is limited to LINE_SIZE. The length member was added in v1.0.44.14 to boost runtime performance.
			this_new_arg.length = (WORD)_tcslen(this_aArg);
			if (   !(this_new_arg.text = SimpleHeap::Malloc(this_aArg, this_new_arg.length))   )
				return FAIL;  // It already displayed the error for us.

			////////////////////////////////////////////////////
			// Build the list of dereferenced vars for this arg.
			////////////////////////////////////////////////////
			// Now that any escaped g_DerefChars have been marked, scan new_arg.text to
			// determine where the variable dereferences are (if any).  In addition to helping
			// runtime performance, this also serves to validate the script at load-time
			// so that some errors can be caught early.  Note: this_new_arg.text is scanned rather
			// than this_aArg because we want to establish pointers to the correct area of
			// memory:
			deref_count = 0;  // Init for each arg.

			// As of v1.0.25, pure numeric parameters can optionally be numeric expressions, so check for that:
			if (*this_new_arg.text // Not omitted.
				&& ArgIsNumeric(aActionType, g_act[aActionType].NumericParams, aArg, i, aArgc))
			{
				// It might be an expression so do the final checks.
				// Override the original false default of is_expression unless an exception applies.
				// Since ACT_ASSIGNEXPR, WHILE, FOR and UNTIL aren't legacy commands, don't call
				// LegacyArgIsExpression() for them because that would cause things like x:=%y% and
				// "while %x%" to behave the same as x:=y and "while x:, which would be inconsistent
				// with how expressions are supposed to work. ACT_RETURN should have been excluded
				// too; but it was left out for so long that it was thought best to document and keep
				// the unexpected behavior of "return %x%".
				// For other commands, if any telltale character is present it's definitely an
				// expression because this is an arg that's marked as a number-or-expression.
				// So telltales avoid the need for the complex check further below.
				if (aActionType == ACT_ASSIGNEXPR || aActionType >= ACT_FOR && aActionType <= ACT_UNTIL // i.e. FOR, WHILE or UNTIL
					|| aActionType == ACT_THROW
					|| StrChrAny(this_new_arg.text, EXPR_TELLTALES)) // See above.
					this_new_arg.is_expression = true;
				// For backward-compatibility, ignore the previous value if this isn't Transform:
				else if (  !(aActionType == ACT_TRANSFORM && this_new_arg.is_expression)  )
					this_new_arg.is_expression = LegacyArgIsExpression(this_new_arg.text, this_aArgMap);
			} // i is a mandatory-numeric arg

			// To help runtime performance, the below changes args to non-expressions if they contain
			// only a single numeric literal (or are entirely blank). At runtime, such args are expanded
			// normally rather than having to run them through the expression evaluator. Exceptions:
			//
			//	a)	ACT_FOR requires an expression; it is incapable of accepting a non-expression.
			//		Although we could treat things like "For x in 0" as load-time errors, it wouldn't
			//		be consistent with "For x in var_containing_num" and the rarity mightn't be worth
			//		the added code.
			//
			//	b)	ACT_ASSIGNEXPR should assign a cached binary integer in addition to the numeric literal.
			//		This might perform just as well overall but is more important for consistency, especially
			//		with COM objects which would otherwise treat the number as a string, potentially causing
			//		a type mismatch error.
			//
			if (this_new_arg.is_expression && IsPureNumeric(this_new_arg.text, true, true, true)
				&& aActionType != ACT_ASSIGNEXPR && aActionType != ACT_FOR && aActionType != ACT_THROW)
				this_new_arg.is_expression = false;

			if (this_new_arg.is_expression)
			{
				// L31: There used to be a section of code here for ensuring parentheses are balanced, but that is done in ExpressionToPostfix now.
				#define ERR_EXP_ILLEGAL_CHAR _T("The leftmost character above is illegal in an expression.") // "above" refers to the layout of the error dialog.
				// ParseDerefs() won't consider escaped percent signs to be illegal, but in this case
				// they should be since they have no meaning in expressions.  UPDATE for v1.0.44.11: The following
				// is now commented out because it causes false positives (and fixing that probably isn't worth the
				// performance & code size).  Specifically, the section below reports an error for escaped delimiters
				// inside quotes such as x := "`%".  More importantly, it defeats the continuation section's %
				// option; for example:
				//   MsgBox %
				//   (%  ; <<< This option here is defeated because it causes % to be replaced with `% at an early stage.
				//   "%"
				//   )
				//if (this_aArgMap) // This arg has an arg map indicating which chars are escaped/literal vs. normal.
				//	for (j = 0; this_new_arg.text[j]; ++j)
				//		if (this_aArgMap[j] && this_new_arg.text[j] == g_DerefChar)
				//			return ScriptError(ERR_EXP_ILLEGAL_CHAR, this_new_arg.text + j);

				// Resolve all operands (that aren't numbers) into variable references.  Doing this here at
				// load-time greatly improves runtime performance, especially for scripts that have a lot
				// of variables.
				for (op_begin = this_new_arg.text; *op_begin; op_begin = op_end)
				{
					if (*op_begin == '.' && op_begin[1] == '=') // v1.0.46.01: Support .=, but not any use of '.' because that is reserved as a struct/member operator.
						op_begin += 2;
					for (; *op_begin && _tcschr(EXPR_OPERAND_TERMINATORS_EX_DOT, *op_begin); ++op_begin); // Skip over whitespace, operators, and parentheses.
					if (!*op_begin) // The above loop reached the end of the string: No operands remaining.
						break;

					// Now op_begin is the start of an operand, which might be a variable reference, a numeric
					// literal, or a string literal.  If it's a string literal, it is left as-is:
					if (*op_begin == '"')
					{
						// Find the end of this string literal, noting that a pair of double quotes is
						// a literal double quote inside the string:
						for (op_end = op_begin + 1;; ++op_end)
						{
							if (!*op_end)
								return ScriptError(ERR_MISSING_CLOSE_QUOTE, op_begin);
							if (*op_end == '"') // If not followed immediately by another, this is the end of it.
							{
								++op_end;
								if (*op_end != '"') // String terminator or some non-quote character.
									break;  // The previous char is the ending quote.
								//else a pair of quotes, which resolves to a single literal quote.
								// This pair is skipped over and the loop continues until the real end-quote is found.
							}
						}
						// op_end is now set correctly to allow the outer loop to continue.
						continue; // Ignore this literal string, letting the runtime expression parser recognize it.
					}
					
					// Find the end of this operand (if *op_end is '\0', strchr() will find that too):
					for (op_end = op_begin + 1; !_tcschr(EXPR_OPERAND_TERMINATORS_EX_DOT, *op_end); ++op_end); // Find first whitespace, operator, or paren.
					if (*op_end == '=' && op_end[-1] == '.') // v1.0.46.01: Support .=, but not any use of '.' because that is reserved as a struct/member operator.
						--op_end;
					// Now op_end marks the end of this operand.  The end might be the zero terminator, an operator, etc.

					// Must be done only after op_end has been set above (since loop uses op_end):
					if (*op_begin == '.' && _tcschr(_T(" \t="), op_begin[1])) // If true, it can't be something like "5." because the dot inside would never be parsed separately in that case.  Also allows ".=" operator.
						continue;
					//else any '.' not followed by a space, tab, or '=' is likely a number without a leading zero,
					// so continue on below to process it.

					cp = omit_leading_whitespace(op_end);
					if (*cp == ':' && cp[1] != '=') // Maybe the "key:" in "{key: value}".
					{
						cp = omit_trailing_whitespace(this_new_arg.text, op_begin - 1);
						if (*cp == ',' || *cp == '{')
						{
							// This is either the key in a key-value pair in an object literal, or a syntax
							// error which will be caught at a later stage (since the ':' is missing its '?').
							cp = find_identifier_end(op_begin);
							if (*cp != '.') // i.e. exclude x.y as that should be parsed as normal for an expression.
							{
								if (cp != op_end) // op contains reserved characters.
									return ScriptError(_T("Quote marks are required around this key."), op_begin);
								continue;
							}
						}
					}

					operand_length = op_end - op_begin;

					// Check if it's AND/OR/NOT:
					if (operand_length < 4 && operand_length > 1) // Ordered for short-circuit performance.
					{
						if (operand_length == 2)
						{
							if ((*op_begin == 'o' || *op_begin == 'O') && (op_begin[1] == 'r' || op_begin[1] == 'R'))
							{	// "OR" was found.
								op_begin[0] = '|'; // v1.0.45: Transform into easier-to-parse symbols for improved
								op_begin[1] = '|'; // runtime performance and reduced code size.  v1.0.48: It no longer helps runtime performance, but it's kept because changing moving it to ExpressionToPostfix() isn't likely to have much benefit.
								continue;
							}
						}
						else // operand_length must be 3
						{
							switch (*op_begin)
							{
							case 'a':
							case 'A':
								if (   (op_begin[1] == 'n' || op_begin[1] == 'N') // Relies on short-circuit boolean order.
									&& (op_begin[2] == 'd' || op_begin[2] == 'D')   )
								{	// "AND" was found.
									op_begin[0] = '&'; // v1.0.45: Transform into easier-to-parse symbols for
									op_begin[1] = '&'; // improved runtime performance and reduced code size.  v1.0.48: It no longer helps runtime performance, but it's kept because changing moving it to ExpressionToPostfix() isn't likely to have much benefit.
									op_begin[2] = ' '; // A space is used lieu of the complexity of the below.
									// Above seems better than below even though below would make it look a little
									// nicer in ListLines.  BELOW CAN'T WORK because this_new_arg.deref[] can contain
									// offsets that would also need to be adjusted:
									//memmove(op_begin + 2, op_begin + 3, _tcslen(op_begin+3)+1 ... or some expression involving this_new_arg.length this_new_arg.text);
									//--this_new_arg.length;
									//--op_end; // Ensure op_end is set up properly for the for-loop's post-iteration action.
									continue;
								}
								break;

							case 'n': // v1.0.45: Unlike "AND" and "OR" above, this one is not given a substitute
							case 'N': // because it's not the same as the "!" operator. See SYM_LOWNOT for comments.
								if (   (op_begin[1] == 'o' || op_begin[1] == 'O') // Relies on short-circuit boolean order.
									&& (op_begin[2] == 't' || op_begin[2] == 'T')   )
									continue; // "NOT" was found.
								if (   (op_begin[1] == 'e' || op_begin[1] == 'E')
									&& (op_begin[2] == 'w' || op_begin[2] == 'W')
									&& IS_SPACE_OR_TAB(op_begin[3])   )
								{
									cp = omit_leading_whitespace(op_begin + 4);
									// This "new " is a keyword only if followed immediately by an operand,
									// such as "new ClassVar", "new Class.NestedClass()" or "new %Var%()",
									// but not "new := 1" or "x := new + 1" and not '"string" new "string"'.
									if (!_tcschr(EXPR_OPERAND_TERMINATORS, *cp) && *cp != '"')
									{
										// If this is "new ClassVar()", we need to avoid parsing "ClassVar()" as a
										// function deref.  The check below handles that.  Note that "new X.Y()" is
										// excluded, since it wouldn't be parsed as a function deref anyway.
										for ( ; !_tcschr(EXPR_OPERAND_TERMINATORS, *cp); ++cp); // Find end of var.
										pending_function_is_new_op = (*cp == '(');
										continue;
									}
								}
								break;
							}
						}
					} // End of check for AND/OR/NOT.

					// Temporarily terminate, which avoids at least the below issue:
					// Two or more extremely long var names together could exceed MAX_VAR_NAME_LENGTH
					// e.g. LongVar%LongVar2% would be too long to store in a buffer of size MAX_VAR_NAME_LENGTH.
					// This seems pretty darn unlikely, but perhaps doubling it would be okay.
					// UPDATE: Above is now not an issue since caller's string is temporarily terminated rather
					// than making a copy of it.
					orig_char = *op_end;
					*op_end = '\0';

					// Illegal characters are legal when enclosed in double quotes.  So the following is
					// done only after the above has ensured this operand is not one enclosed entirely in
					// double quotes.
					// The following characters are either illegal in expressions or reserved for future use.
					// Rather than forbidding g_delimiter and g_DerefChar, it seems best to assume they are at
					// their default values for this purpose.  Otherwise, if g_delimiter is an operator, that
					// operator would then become impossible inside the expression.
					if (cp = StrChrAny(op_begin, EXPR_ILLEGAL_CHARS))
						return ScriptError(ERR_EXP_ILLEGAL_CHAR, cp);

					// Below takes care of recognizing hexadecimal integers, which avoids the 'x' character
					// inside of something like 0xFF from being detected as the name of a variable:
					if (   !IsPureNumeric(op_begin, true, false, true) // Not a numeric literal...
						/*&& !(*op_begin == '?' && !op_begin[1])*/   ) // ...and not an isolated '?' operator.  Relies on short-circuit boolean order.
					{
						if (*op_begin == '.') // L31: Check for something like "obj .property" - scientific-notation literals such as ".123e+1" may also be handled here.
						{
							if (_tcschr(op_begin, g_DerefChar))
								return ScriptError(ERR_INVALID_DOT, op_begin);
							// Skip over this scientific-notation literal or string of one or more member-access operations.
							// This won't skip the signed exponent of a scientific-notation literal, but that should be OK
							// as it will be recognized as purely numeric in the next iteration of this loop.
							*op_end = orig_char;
							continue;
						}
						if (cp = _tcschr(op_begin + 1, '.')) // L31: Check for scientific-notation literal (as in previous versions) or something like "obj.property". Above has already handled "obj .property" and similar.
						{
							if (ctoupper(op_end[-1]) == 'E' && (orig_char == '+' || orig_char == '-')) // Listed first for short-circuit performance with the below.
							{
								 // v1.0.46.11: This item appears to be a scientific-notation literal with the OPTIONAL +/- sign PRESENT on the exponent (e.g. 1.0e+001), so check that before checking if it's a variable name.
								*op_end = orig_char; // Undo the temporary termination.
								do // Skip over the sign and its exponent; e.g. the "+1" in "1.0e+1".  There must be a sign in this particular sci-notation number or we would never have arrived here.
									++op_end;
								while (*op_end >= '0' && *op_end <= '9'); // Avoid isdigit() because it sometimes causes a debug assertion failure at: (unsigned)(c + 1) <= 256 (probably only in debug mode), and maybe only when bad data got in it due to some other bug.
								// No need to do the following because a number can't validly be followed by the ".=" operator:
								//if (*op_end == '=' && op_end[-1] == '.') // v1.0.46.01: Support .=, but not any use of '.' because that is reserved as a struct/member operator.
								//	--op_end;

								// Double-check it really is a floating-point literal with signed exponent.
								orig_char = *op_end;
								*op_end = '\0';
								if (IsPureNumeric(op_begin, true, false, true))
								{
									*op_end = orig_char;
									continue; // Pure number, which doesn't need any processing at this stage.
								}
							}
							// else this is NOT a scientific-notation literal with +/- sign present, so treat it as a member-access operation.
							// Resolve the part preceding '.' as a variable reference. The rest is handled later, in ExpressionToPostfix.
							if (_tcschr(cp, g_DerefChar))
								return ScriptError(ERR_INVALID_DOT, cp);
							operand_length = cp - op_begin;
							is_function = false;
						}
						else
							is_function = (orig_char == '(');

						if (is_function && pending_function_is_new_op)
							pending_function_is_new_op = is_function = false;

						// This operand must be a variable/function reference or string literal, otherwise it's
						// a syntax error.
						// Check explicitly for derefs since the vast majority don't have any, and this
						// avoids the function call in those cases:
						if (_tcschr(op_begin, g_DerefChar)) // This operand contains at least one double dereference.
						{
							// v1.0.47.06: Dynamic function calls are now supported.
							//if (is_function)
							//	return ScriptError("Dynamic function calls are not supported.", op_begin);
							int first_deref = deref_count;

							// The percent-sign derefs are parsed and added to the deref array at this stage (on a
							// per-operand basis) rather than all at once for the entire arg because
							// the deref array must contain both percent-sign derefs and non-%-derefs interspersed
							// and ordered according to their physical position inside the arg, but ParseDerefs
							// only handles percent-sign derefs, not expression derefs like x+y.  In the following
							// example, the order of derefs must be x,i,y: if (x = Array%i% and y = 3)
							if (!ParseDerefs(op_begin, this_aArgMap ? this_aArgMap + (op_begin - this_new_arg.text) : NULL
								, deref, deref_count))
								return FAIL; // It already displayed the error.  No need to undo temp. termination.
							// And now leave this operand "raw" so that it will later be dereferenced again.
							// In the following example, i made into a deref but the result (Array33) must be
							// dereferenced during a second stage at runtime: if (x = Array%i%).

							if (is_function) // Dynamic function call.
								deref[first_deref].param_count = 0; // L31: Parameters are now counted by ExpressionToPostfix instead of here, but this must be initialized.
						}
						else // This operand is a variable name or function name (single deref).
						{
							#define TOO_MANY_REFS _T("Too many var/func refs.") // Short msg since so rare.
							if (deref_count >= MAX_DEREFS_PER_ARG)
								return ScriptError(TOO_MANY_REFS, op_begin); // Indicate which operand it ran out of space at.
							// Store the deref's starting location, even for functions (leave it set to the start
							// of the function's name for use when doing error reporting at other stages -- i.e.
							// don't set it to the address of the first param or closing-paren-if-no-params):
							deref[deref_count].marker = op_begin;
							deref[deref_count].length = (DerefLengthType)operand_length;
							if (deref[deref_count].is_function = is_function) // It's a function not a variable.
								// Set to NULL to catch bugs.  It must and will be filled in at a later stage
								// because the setting of each function's mJumpToLine relies upon the fact that
								// functions are added to the linked list only upon being formally defined
								// so that the most recently defined function is always last in the linked
								// list, awaiting its mJumpToLine that will appear beneath it.
								deref[deref_count].func = NULL;
							else // It's a variable rather than a function.
								if (   !(deref[deref_count].var = FindOrAddVar(op_begin, operand_length))   )
									return FAIL; // The called function already displayed the error.
							++deref_count; // Since above didn't "continue" or "return".
						}
					}
					//else purely numeric or '?'.  Do nothing since pure numbers and '?' don't need any
					// processing at this stage.
					*op_end = orig_char; // Undo the temporary termination.
				} // expression pre-parsing loop.

				// Now that the derefs have all been recognized above, simplify any special cases --
				// such as single isolated derefs -- to enhance runtime performance.
				//
				// There used to be a section here that translated each expression that consisted
				// solely of a quoted/literal string into a non-expression (except Post/SendMessage).
				// However, that is no longer appropriate for ACT_ASSIGNEXPR (which was the main
				// beneficiary) because an optimization further below would wrongly apply
				// SetFormat to the assigning of a quoted/literal string like Var:="55".
				// Benchmarks show that performance of assigning quoted literal strings is only
				// slightly slower when is_expression==true; also, the savings in code size and the
				// fact that the translation made ListLines inaccurate (due to the omitted quotes)
				// seem to support getting rid of that section.
				//
				// Make things like "Sleep Var" and "Var := X" into non-expressions.  At runtime,
				// such args are expanded normally rather than having to run them through the
				// expression evaluator.  A simple test script shows that this one change can
				// double the runtime performance of certain commands such as EnvAdd:
				// Below is somewhat obsolete but kept for reference:
				// This policy is basically saying that expressions are allowed to evaluate to strings
				// everywhere appropriate, but that at the moment the only appropriate place is x := y
				// because all other expressions should resolve to a numeric value by virtue of the fact
				// that they *are* numeric parameters.  ValidateName() serves to eliminate cases where
				// a single deref is accompanied by literal numbers, strings, or operators, e.g.
				// Var := X + 1 ... Var := Var2 "xyz" ... Var := -Var2
				if (deref_count == 1 && Var::ValidateName(this_new_arg.text, DISPLAY_NO_ERROR)) // Single isolated deref.
				{
					// ACT_WHILE performs less than 4% faster as a non-expression in these cases, and keeping
					// it as an expression avoids an extra check in a performance-sensitive spot of ExpandArgs
					// (near mActionType <= ACT_LAST_OPTIMIZED_IF).
					if (aActionType < ACT_FOR || aActionType > ACT_UNTIL && aActionType != ACT_THROW) // If it is FOR, WHILE, UNTIL or THROW, it would be something like "while x" in this case. Keep those as expressions for the reason above. PerformLoopFor() requires FOR's expression arg to remain an expression.
						this_new_arg.is_expression = false; // In addition to being an optimization, doing this might also be necessary for things like "Var := ClipboardAll" to work properly.
					// But if aActionType is ACT_ASSIGNEXPR, it's left as ACT_ASSIGNEXPR vs. ACT_ASSIGN
					// because it might be necessary to avoid having AutoTrim take effect for := (which
					// it never should).  In addition, ACT_ASSIGNEXPR probably performs better than
					// ACT_ASSIGN when is_expression==false.
				}
				else if (deref_count && !StrChrAny(this_new_arg.text, EXPR_OPERAND_TERMINATORS)) // No spaces, tabs, etc.
				{
					// Adjust if any of the following special cases apply:
					// x := y  -> Mark as non-expression (after expression-parsing set up parsed derefs above)
					//            so that the y deref will be only a single-deref to be directly stored in x.
					//            This is done in case y contains a string.  Since an expression normally
					//            evaluates to a number, without this workaround, x := y would be useless for
					//            a simple assignment of a string.  This case is handled above.
					// x := %y% -> Mark the right-side arg as an input variable so that it will be doubly
					//             dereferenced, similar to StringTrimRight, Out, %y%, 0.  This seems best
					//             because there would be little or no point to having it behave identically
					//             to x := y.  It might even be confusing in light of the next case below.
					// CASE #3:
					// x := Literal%y%Literal%z%Literal -> Same as above.  This is done mostly to support
					// retrieving array elements whose contents are *non-numeric* without having to use
					// something like StringTrimRight.
					
					// Now we know it has at least one deref.  But if any operators or other characters disallowed
					// in variables are present, it all three cases are disqualified and kept as expressions.
					// This check is necessary for all three cases:

					// No operators of any kind anywhere.  Not even +/- prefix, since those imply a numeric
					// expression.  No chars illegal in var names except the percent signs themselves,
					// e.g. *no* whitespace.
					// Also, the first deref (indeed, all of them) should point to a percent sign, since
					// there should not be any way for non-percent derefs to get mixed in with cases
					// 2 or 3.
					if (!deref[0].is_function && *deref[0].marker == g_DerefChar // This appears to be case #2 or #3.
						&& (aActionType < ACT_FOR || aActionType > ACT_UNTIL)) // Nearly doubles the speed of "while %x%" and "while Array%i%" to leave WHILE as an expression.  But y:=%x% and y:=Array%i% are about the same speed either way, and "if %x%" never reaches this point because for compatibility(?), it's the same as "if x". Additionally, PerformLoopFor() requires its only expression arg to remain an expression.
					{
						// The comment below is probably obsolete -- and perhaps so is this entire optimization
						// because expressions are faster now.  But in case it's necessary for anything related
						// to backward compatibility, it's kept (it may also reduce memory utilization a little
						// because it avoids making simple things into expressions, which require extra memory).
						// OLD: Set it up so that x:=Array%i% behaves the same as StringTrimRight, Out, Array%i%, 0.
						this_new_arg.is_expression = false;
						this_new_arg.type = ARG_TYPE_INPUT_VAR;
					}
				}
			} // if (this_new_arg.is_expression)
			else // this arg does not contain an expression.
				if (!ParseDerefs(this_new_arg.text, this_aArgMap, deref, deref_count))
					return FAIL; // It already displayed the error.

			//////////////////////////////////////////////////////////////
			// Allocate mem for this arg's list of dereferenced variables.
			//////////////////////////////////////////////////////////////
			if (deref_count)
			{
				// +1 for the "NULL-item" terminator:
				if (   !(this_new_arg.deref = (DerefType *)SimpleHeap::Malloc((deref_count + 1) * sizeof(DerefType)))   )
					return ScriptError(ERR_OUTOFMEM);
				memcpy(this_new_arg.deref, deref, deref_count * sizeof(DerefType));
				// Terminate the list of derefs with a deref that has a NULL marker:
				this_new_arg.deref[deref_count].marker = NULL;
			}
			else
				this_new_arg.deref = NULL;
		} // for each arg.
	} // there are more than zero args.

	//////////////////////////////////////////////////////////////////////////////////////
	// Now the above has allocated some dynamic memory, the pointers to which we turn over
	// to Line's constructor so that they can be anchored to the new line.
	//////////////////////////////////////////////////////////////////////////////////////
	Line *the_new_line = new Line(mCurrFileIndex, mCombinedLineNumber, aActionType, new_arg, aArgc);
	if (!the_new_line)
		return ScriptError(ERR_OUTOFMEM);

	Line &line = *the_new_line;  // For performance and convenience.

	line.mPrevLine = mLastLine;  // Whether NULL or not.
	if (mFirstLine == NULL)
		mFirstLine = the_new_line;
	else
		mLastLine->mNextLine = the_new_line;
	// This must be done after the above:
	mLastLine = the_new_line;
	mCurrLine = the_new_line;  // To help error reporting.

	///////////////////////////////////////////////////////////////////
	// Do any post-add validation & handling for specific action types.
	///////////////////////////////////////////////////////////////////
#ifndef AUTOHOTKEYSC // For v1.0.35.01, some syntax checking is removed in compiled scripts to reduce their size.
	int value;    // For temp use during validation.
	double value_float;
	SYSTEMTIME st;  // same.
#endif
	// v1.0.38: The following should help reduce code size, and for some commands helps load-time
	// performance by avoiding multiple resolutions of a given macro:
	LPTSTR new_raw_arg1 = NEW_RAW_ARG1;
	LPTSTR new_raw_arg2 = NEW_RAW_ARG2;
	LPTSTR new_raw_arg3 = NEW_RAW_ARG3;
	LPTSTR new_raw_arg4 = NEW_RAW_ARG4;

	switch(aActionType)
	{
	// Fix for v1.0.35.02:
	// THESE FIRST FEW CASES MUST EXIST IN BOTH SELF-CONTAINED AND NORMAL VERSION since they alter the
	// attributes/members of some types of lines:
	case ACT_SUB:
		if (aArgc < 2) // Validate at loadtime so that at runtime, DETERMINE_NUMERIC_TYPES and ARG2_AS_INT64 don't have to check that mArgc > 1.
			return ScriptError(ERR_PARAM2_REQUIRED);
		// ** DON'T BREAK; FALL INTO NEXT SECTION **
	case ACT_ADD:  // ************ OR IT FELL INTO THIS SECTION FROM ABOVE ************
	case ACT_MULT:
	case ACT_DIV:
#ifndef AUTOHOTKEYSC // For v1.0.35.01, some syntax checking is removed in compiled scripts to reduce their size.
		if (aArgc > 2) // Then this is ACT_ADD OR ACT_SUB with a 3rd parameter (TimeUnits)
		{
			if (*new_raw_arg3 && !line.ArgHasDeref(3))
				if (!_tcschr(_T("SMHD"), ctoupper(*new_raw_arg3)))  // (S)econds, (M)inutes, (H)ours, or (D)ays
					return ScriptError(ERR_PARAM3_INVALID, new_raw_arg3);
			if (aActionType == ACT_SUB && *new_raw_arg2 && !line.ArgHasDeref(2))
				if (!YYYYMMDDToSystemTime(new_raw_arg2, st, true))
					return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
			new_arg[1].postfix = NULL; // It's necessary to indicate that there is no cached binary number for arg #2 in the 3-arg mode of ACT_ADD due to runtime logic that checks it.  For the others, this helps maintainability.
			break; // Don't allow processing to continue.  Other sections below rely on this.
		}
		if (aActionType == ACT_DIV && !line.ArgHasDeref(2) && !new_arg[1].is_expression) // i.e. don't validate the following until runtime:
			if (!ATOF(new_raw_arg2))                           // x/=y ... x/=(4/4)/4 (v1.0.46.01: added is_expression check for expressions with no variables or function-calls).
				return ScriptError(ERR_DIVIDEBYZERO, new_raw_arg2);
		// ** DON'T BREAK; FALL INTO NEXT SECTION **
#endif
	case ACT_ASSIGN: // **** OR IT FELL INTO THIS SECTION FROM ABOVE ****
	case ACT_ASSIGNEXPR:
	case ACT_IFEQUAL:
	case ACT_IFNOTEQUAL:
	case ACT_IFGREATER:
	case ACT_IFGREATEROREQUAL:
	case ACT_IFLESS:
	case ACT_IFLESSOREQUAL:
	case ACT_IFBETWEEN:
	case ACT_IFNOTBETWEEN:
		int arg_index;
		for (arg_index = 1; arg_index < aArgc; ++arg_index) // Up to two iterations: arg #2 and arg#3 (for If[Not]Between)).
		{
			if (   *new_arg[arg_index].text && !line.ArgHasDeref(arg_index+1)
				&& !new_arg[arg_index].is_expression  // Expressions don't make sense for this, plus they need their postfix member for other purposes.
				&& IsPureNumeric(new_arg[arg_index].text, true, false, false) // aAllowImpure==false even for ACT_ADD/SUB/MULT/DIV because those would see almost all impure *LITERAL* numbers like 123abc as variables (too rare anyway).  Check for purity to rule out floats and expressions consisting only of literals such as 1+2 (in case they can ever be encountered here).
				&& !((aActionType == ACT_ASSIGN || aActionType == ACT_ASSIGNEXPR) // Only these need extra checking because the display format of the number doesn't matter for ADD/SUB/IFEQUAL/IFGREATER/etc. because they treat anything that looks like a number (any format) as a pure number.
					&& (
							   *new_raw_arg2 == '0' || *new_raw_arg2 == '+' // Assign hex or any unusually-formatted integers the old way so that the format is retained in case its important to the operation of the script (e.g. x:="005", x:=005, x:="0x5", x:="+5", x:=+5).
							|| new_arg[1].length > 18 // See below.
							// Integers that are too long are probably intended to be a series of characters/digits,
							// so assign them the old way to keep all of the digits.  Fix for v1.0.48.01: Reduced the
							// limit from MAX_INTEGER_LENGTH (20) to 18 so that the assignment (:= and =) of integers
							// that are 19 or 20 digits long work as they did prior to v1.0.48 (some of such integers
							// would overflow a signed 64-bit value, so keep all of them as strings).
							//
							// The following can't happen anymore because x:="string" is no longer translated
							// into is_expression==false.  There are some reasons given in a section higher above:
							//|| IS_SPACE_OR_TAB(new_raw_arg2[new_arg[1].length-1]) // Trailing whitespace, which can happen from something like x:="abc ".
							//|| IS_SPACE_OR_TAB(*new_raw_arg2) // This can happen via translation of x:=" abc " to x:= abc at an earlier stage.
							// Any LITERAL whitespace around a LITERAL number has always been ignored/omitted,
							// so storing binary integers for things like "x = 1" and "x := 1" should behave
							// as before, with the exception of "SetFormat, Integer, Hex", which will now be obeyed
							// by such assignments when it wasn't before.
						))   )
			{
				if (   !(new_arg[arg_index].postfix = (ExprTokenType *)SimpleHeap::Malloc(sizeof(__int64)))   )
					return ScriptError(ERR_OUTOFMEM);
				*(__int64 *)new_arg[arg_index].postfix = ATOI64(new_arg[arg_index].text);
			}
			else
				new_arg[arg_index].postfix = NULL; // Indicate that there is no cached binary number.
		}
		break;

	case ACT_LOOP:
		// If possible, determine the type of loop so that the preparser can better
		// validate some things:
		switch (aArgc)
		{
		case 0:
			line.mAttribute = ATTR_LOOP_NORMAL;
			break;
		case 1: // With only 1 arg, it must be a normal loop, file-pattern loop, or registry loop.
			// v1.0.43.07: Added check for new_arg[0].is_expression so that an expression without any variables
			// it it works (e.g. Loop % 1+1):
			if (line.ArgHasDeref(1) || new_arg[0].is_expression) // Impossible to know now what type of loop (only at runtime).
				line.mAttribute = ATTR_LOOP_UNKNOWN;
			else
			{
				if (IsPureNumeric(new_raw_arg1, false))
					line.mAttribute = ATTR_LOOP_NORMAL;
				else
					line.mAttribute = line.RegConvertRootKey(new_raw_arg1) ? ATTR_LOOP_REG : ATTR_LOOP_FILEPATTERN;
			}
			break;
		default:  // has 2 or more args.
			if (line.ArgHasDeref(1)) // Impossible to know now what type of loop (only at runtime).
				line.mAttribute = ATTR_LOOP_UNKNOWN;
			else if (!_tcsicmp(new_raw_arg1, _T("Read")))
				line.mAttribute = ATTR_LOOP_READ_FILE;
			else if (!_tcsicmp(new_raw_arg1, _T("Parse")))
				line.mAttribute = ATTR_LOOP_PARSE;
			else if (!_tcsicmp(new_raw_arg1, _T("Reg")))
				line.mAttribute = ATTR_LOOP_NEW_REG;
			else if (!_tcsicmp(new_raw_arg1, _T("Files")))
				line.mAttribute = ATTR_LOOP_NEW_FILES;
			else // the 1st arg can either be a Root Key or a File Pattern, depending on the type of loop.
			{
				line.mAttribute = line.RegConvertRootKey(new_raw_arg1) ? ATTR_LOOP_REG : ATTR_LOOP_FILEPATTERN;
				if (line.mAttribute == ATTR_LOOP_FILEPATTERN)
				{
					// Validate whatever we can rather than waiting for runtime validation:
					if (!line.ArgHasDeref(2) && Line::ConvertLoopMode(new_raw_arg2) == FILE_LOOP_INVALID)
						return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
					if (*new_raw_arg3 && !line.ArgHasDeref(3))
						if (_tcslen(new_raw_arg3) > 1 || (*new_raw_arg3 != '0' && *new_raw_arg3 != '1'))
							return ScriptError(ERR_PARAM3_INVALID, new_raw_arg3);
				}
				else // Registry loop.
				{
					if (aArgc > 2 && !line.ArgHasDeref(3) && Line::ConvertLoopMode(new_raw_arg3) == FILE_LOOP_INVALID)
						return ScriptError(ERR_PARAM3_INVALID, new_raw_arg3);
					if (*new_raw_arg4 && !line.ArgHasDeref(4))
						if (_tcslen(new_raw_arg4) > 1 || (*new_raw_arg4 != '0' && *new_raw_arg4 != '1'))
							return ScriptError(ERR_PARAM4_INVALID, new_raw_arg4);
				}
			}
		}
		break; // Outer switch().

	case ACT_WHILE: // Lexikos: ATTR_LOOP_WHILE is used to differentiate ACT_WHILE from ACT_LOOP, allowing code to be shared.
		line.mAttribute = ATTR_LOOP_WHILE;
		break;

	case ACT_FOR:
		line.mAttribute = ATTR_LOOP_FOR;
		break;

	// This one alters g_persistent so is present in its entirety (for simplicity) in both SC an non-SC version.
	case ACT_GUI:
		// By design, scripts that use the GUI cmd anywhere are persistent.  Doing this here
		// also allows WinMain() to later detect whether this script should become #SingleInstance.
		// Note: Don't directly change g_AllowOnlyOneInstance here in case the remainder of the
		// script-loading process comes across any explicit uses of #SingleInstance, which would
		// override the default set here.
		g_persistent = true;
#ifndef AUTOHOTKEYSC // For v1.0.35.01, some syntax checking is removed in compiled scripts to reduce their size.
		if (aArgc > 0 && !line.ArgHasDeref(1))
		{
			LPTSTR command, name;
			ResolveGui(new_raw_arg1, command, &name);
			if (!name)
				return ScriptError(ERR_INVALID_GUI_NAME, new_raw_arg1);

			GuiCommands gui_cmd = line.ConvertGuiCommand(command);

			switch (gui_cmd)
			{
			case GUI_CMD_INVALID:
				return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
			case GUI_CMD_ADD:
				if (aArgc > 1 && !line.ArgHasDeref(2))
				{
					GuiControls control_type;
					if (   !(control_type = line.ConvertGuiControl(new_raw_arg2))   )
						return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
					if (control_type == GUI_CONTROL_TREEVIEW && aArgc > 3) // Reserve it for future use such as a tab-indented continuation section that lists the tree hierarchy.
						return ScriptError(ERR_PARAM4_MUST_BE_BLANK, new_raw_arg4);
				}
				break;
			case GUI_CMD_CANCEL:
			case GUI_CMD_MINIMIZE:
			case GUI_CMD_MAXIMIZE:
			case GUI_CMD_RESTORE:
			case GUI_CMD_DESTROY:
			case GUI_CMD_DEFAULT:
			case GUI_CMD_OPTIONS:
				if (aArgc > 1)
					return ScriptError(_T("Parameter #2 and beyond should be omitted in this case."), new_raw_arg2);
				break;
			case GUI_CMD_SUBMIT:
			case GUI_CMD_MENU:
			case GUI_CMD_LISTVIEW:
			case GUI_CMD_TREEVIEW:
			case GUI_CMD_FLASH:
				if (aArgc > 2)
					return ScriptError(_T("Parameter #3 and beyond should be omitted in this case."), new_raw_arg3);
				break;
			// No action for these since they have a varying number of optional params:
			//case GUI_CMD_NEW:
			//case GUI_CMD_SHOW:
			//case GUI_CMD_FONT:
			//case GUI_CMD_MARGIN:
			//case GUI_CMD_TAB:
			//case GUI_CMD_COLOR: No load-time param validation to avoid larger EXE size.
			}
		}
#endif
		break;

	case ACT_GROUPADD:
	case ACT_GROUPACTIVATE:
	case ACT_GROUPDEACTIVATE:
	case ACT_GROUPCLOSE:
		// For all these, store a pointer to the group to help performance.
		// We create a non-existent group even for ACT_GROUPACTIVATE, ACT_GROUPDEACTIVATE
		// and ACT_GROUPCLOSE because we can't rely on the ACT_GROUPADD commands having
		// been parsed prior to them (e.g. something like "Gosub, DefineGroups" may appear
		// in the auto-execute portion of the script).
		if (!line.ArgHasDeref(1))
			if (   !(line.mAttribute = FindGroup(new_raw_arg1, true))   ) // Create-if-not-found so that performance is enhanced at runtime.
				return FAIL;  // The above already displayed the error.
		if (aActionType == ACT_GROUPACTIVATE || aActionType == ACT_GROUPDEACTIVATE)
		{
			if (*new_raw_arg2 && !line.ArgHasDeref(2))
				if (_tcslen(new_raw_arg2) > 1 || ctoupper(*new_raw_arg2) != 'R')
					return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
		}
		else if (aActionType == ACT_GROUPCLOSE)
			if (*new_raw_arg2 && !line.ArgHasDeref(2))
				if (_tcslen(new_raw_arg2) > 1 || !_tcschr(_T("RA"), ctoupper(*new_raw_arg2)))
					return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
		break;

	case ACT_SETFORMAT: // Must be done even when AUTOHOTKEYSC is defined so that g_WriteCacheDisabledInt64/Double is properly updated.
		if (aArgc < 1)
			break;
		if (line.ArgHasDeref(1)) // Something like "SetFormat, %Var%, ..."
		{
			// For the following and other sections further below that disable the cache, can't wait until
			// runtime execution encounters SetFormat to disable caching because script might rely on the
			// *default* format being *immediately* written out prior to the script changing SetFormat at
			// some later time.
			g_WriteCacheDisabledInt64 = TRUE;
			g_WriteCacheDisabledDouble = TRUE;
		}
		else
		{
			if (!_tcsnicmp(new_raw_arg1, _T("Float"), 5))
			{
				if (_tcsicmp(new_raw_arg1 + 5, _T("Fast"))) // Cache is left enabled when the new FloatFast/IntegerFast mode is present.
					g_WriteCacheDisabledDouble = TRUE;
				if (aArgc > 1 && !line.ArgHasDeref(2))
				{
					if (!IsPureNumeric(new_raw_arg2, true, false, true, true) // v1.0.46.11: Allow impure numbers to support scientific notation; e.g. 0.6e or 0.6E.
						|| _tcslen(new_raw_arg2) >= _countof(g->FormatFloat) - 2)
						return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
				}
			}
			else if (!_tcsnicmp(new_raw_arg1, _T("Integer"), 7))
			{
				if (_tcsicmp(new_raw_arg1 + 7, _T("Fast"))) // Cache is left enabled when the new FloatFast/IntegerFast mode is present.
					g_WriteCacheDisabledInt64 = TRUE;
				if (aArgc > 1 && !line.ArgHasDeref(2) && ctoupper(*new_raw_arg2) != 'H' && ctoupper(*new_raw_arg2) != 'D')
					return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
			}
			else
				return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		}
		// Size must be less than sizeof() minus 2 because need room to prepend the '%' and append
		// the 'f' to make it a valid format specifier string:
		break;

	case ACT_STRINGSPLIT: // v1.0.48.04: Moved this section so that it is done even when AUTOHOTKEYSC is defined, because the steps below are necessary for both.
		if (*new_raw_arg1 && !line.ArgHasDeref(1)) // The output array must be a legal name.
		{
			// 1.0.46.10: Fixed to look up ArrayName0 in advance (here at loadtime) so that runtime can
			// know whether it's local or global.  This is necessary because only here at loadtime
			// is there any awareness of the current function's list of declared variables (to conserve
			// memory, that list is longer available at runtime).
			TCHAR temp_var_name[MAX_VAR_NAME_LENGTH + 10]; // Provide extra room for trailing "0", and to detect names that are too long.
			sntprintf(temp_var_name, _countof(temp_var_name), _T("%s0"), new_raw_arg1);
			if (   !(the_new_line->mAttribute = FindOrAddVar(temp_var_name))   )
				return FAIL;  // The above already displayed the error.
		}
		//else it's a dynamic array name.  Since that's very rare, just use the old runtime behavior for
		// backward compatibility.
		break;

#ifndef AUTOHOTKEYSC // For v1.0.35.01, some syntax checking is removed in compiled scripts to reduce their size.
	case ACT_RETURN:
		if (aArgc > 0 && !g->CurrentFunc)
			return ScriptError(_T("Return's parameter should be blank except inside a function."));
		break;

	case ACT_AUTOTRIM:
	case ACT_DETECTHIDDENWINDOWS:
	case ACT_DETECTHIDDENTEXT:
	case ACT_SETSTORECAPSLOCKMODE:
		if (aArgc > 0 && !line.ArgHasDeref(1) && !line.ConvertOnOff(new_raw_arg1))
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;

	case ACT_STRINGCASESENSE:
		if (aArgc > 0 && !line.ArgHasDeref(1) && line.ConvertStringCaseSense(new_raw_arg1) == SCS_INVALID)
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;

	case ACT_SETBATCHLINES:
		if (aArgc > 0 && !line.ArgHasDeref(1))
		{
			if (!tcscasestr(new_raw_arg1, _T("ms")) && !IsPureNumeric(new_raw_arg1, true, false)) // For simplicity and due to rarity, new_arg[0].is_expression isn't checked, so a line with no variables or function-calls like "SetBatchLines % 1+1" will be wrongly seen as a syntax error.
				return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		}
		break;

	case ACT_SUSPEND:
		if (aArgc > 0 && !line.ArgHasDeref(1) && !line.ConvertOnOffTogglePermit(new_raw_arg1))
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;

	case ACT_BLOCKINPUT:
		if (aArgc > 0 && !line.ArgHasDeref(1) && !line.ConvertBlockInput(new_raw_arg1))
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;

	case ACT_SENDMODE:
		if (aArgc > 0 && !line.ArgHasDeref(1) && line.ConvertSendMode(new_raw_arg1, SM_INVALID) == SM_INVALID)
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;

	case ACT_SENDLEVEL:
		if (aArgc > 0 && !line.ArgHasDeref(1) && !SendLevelIsValid(ATOI(new_raw_arg1)))
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;

	case ACT_PAUSE:
	case ACT_KEYHISTORY:
		if (aArgc > 0 && !line.ArgHasDeref(1) && !line.ConvertOnOffToggle(new_raw_arg1))
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;

	case ACT_SETNUMLOCKSTATE:
	case ACT_SETSCROLLLOCKSTATE:
	case ACT_SETCAPSLOCKSTATE:
		if (aArgc > 0 && !line.ArgHasDeref(1) && !line.ConvertOnOffAlways(new_raw_arg1))
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;

	case ACT_STRINGMID:
		if (aArgc > 4 && !line.ArgHasDeref(5) && _tcsicmp(NEW_RAW_ARG5, _T("L")))
			return ScriptError(ERR_PARAM5_INVALID, NEW_RAW_ARG5);
		break;

	case ACT_STRINGGETPOS:
		if (*new_raw_arg4 && !line.ArgHasDeref(4) && !_tcschr(_T("LR1"), ctoupper(*new_raw_arg4)))
			return ScriptError(ERR_PARAM4_INVALID, new_raw_arg4);
		break;

	case ACT_REGREAD:
		// v1.1.21: The undocumented and obsolete 5-param syntax from AutoIt v2 is no longer supported.
		// Even the earliest known version of AutoHotkey (v0.207) did not use the ValueType parameter.
		// Example of obsolete syntax:  RegRead, OutVar, REG_SZ, HKEY_CURRENT_USER, Software\Winamp,
		// The following detects it as an error, since REG_SZ is not a valid root key:
		if (*new_raw_arg2 && !line.ArgHasDeref(2) && !line.RegConvertKey(new_raw_arg2, REG_EITHER_SYNTAX))
			return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
		break;

	case ACT_SETREGVIEW:
		if (!line.ArgHasDeref(1) && line.RegConvertView(new_raw_arg1) == -1)
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;

	case ACT_REGWRITE:
		// Both of these checks require that at least two parameters be present.  Otherwise, the command
		// is being used in its registry-loop mode and is validated elsewhere:
		if (aArgc > 1)
		{
			if (*new_raw_arg1 && !line.ArgHasDeref(1) && !line.RegConvertValueType(new_raw_arg1))
				return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
			if (*new_raw_arg2 && !line.ArgHasDeref(2) && !line.RegConvertKey(new_raw_arg2, REG_EITHER_SYNTAX))
				return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
		}
		break;

	case ACT_REGDELETE:
		if (*new_raw_arg1 && !line.ArgHasDeref(1) && !line.RegConvertKey(new_raw_arg1, REG_EITHER_SYNTAX))
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;

	case ACT_SOUNDGET:
	case ACT_SOUNDSET:
		if (aActionType == ACT_SOUNDSET && aArgc > 0 && !line.ArgHasDeref(1))
		{
			// The value of catching syntax errors at load-time seems to outweigh the fact that this check
			// sees a valid no-deref expression such as 300-250 as invalid.
			value_float = ATOF(new_raw_arg1);
			if (value_float < -100 || value_float > 100)
				return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		}
		if (*new_raw_arg2 && !line.ArgHasDeref(2) && !line.SoundConvertComponentType(new_raw_arg2))
			return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
		if (*new_raw_arg3 && !line.ArgHasDeref(3) && line.SoundConvertControlType(new_raw_arg3) == MIXERCONTROL_CONTROLTYPE_INVALID)
			return ScriptError(ERR_PARAM3_INVALID, new_raw_arg3);
		break;

	case ACT_SOUNDSETWAVEVOLUME:
		if (aArgc > 0 && !line.ArgHasDeref(1))
		{
			// The value of catching syntax errors at load-time seems to outweigh the fact that this check
			// sees a valid no-deref expression such as 300-250 as invalid.
			value_float = ATOF(new_raw_arg1);
			if (value_float < -100 || value_float > 100)
				return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		}
		break;

	case ACT_SOUNDPLAY:
		if (*new_raw_arg2 && !line.ArgHasDeref(2) && _tcsicmp(new_raw_arg2, _T("wait")) && _tcsicmp(new_raw_arg2, _T("1")))
			return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
		break;

	case ACT_PIXELSEARCH:
	case ACT_IMAGESEARCH:
		if (!*new_raw_arg3 || !*new_raw_arg4 || !*NEW_RAW_ARG5 || !*NEW_RAW_ARG6 || !*NEW_RAW_ARG7)
			return ScriptError(_T("Parameters 3 through 7 must not be blank."));
		if (aActionType != ACT_IMAGESEARCH)
		{
			if (*NEW_RAW_ARG8 && !line.ArgHasDeref(8))
			{
				// The value of catching syntax errors at load-time seems to outweigh the fact that this check
				// sees a valid no-deref expression such as 300-200 as invalid.
				value = ATOI(NEW_RAW_ARG8);
				if (value < 0 || value > 255)
					return ScriptError(ERR_PARAM8_INVALID, NEW_RAW_ARG8);
			}
		}
		break;

	case ACT_COORDMODE:
		if (*new_raw_arg1 && !line.ArgHasDeref(1) && line.ConvertCoordModeCmd(new_raw_arg1) == -1)
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		if (*new_raw_arg2 && !line.ArgHasDeref(2) && line.ConvertCoordMode(new_raw_arg2) == -1)
			return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
		break;

	case ACT_SETDEFAULTMOUSESPEED:
		if (*new_raw_arg1 && !line.ArgHasDeref(1))
		{
			// The value of catching syntax errors at load-time seems to outweigh the fact that this check
			// sees a valid no-deref expression such as 1+2 as invalid.
			value = ATOI(new_raw_arg1);
			if (value < 0 || value > MAX_MOUSE_SPEED)
				return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		}
		break;

	case ACT_MOUSEMOVE:
		if (*new_raw_arg3 && !line.ArgHasDeref(3))
		{
			// The value of catching syntax errors at load-time seems to outweigh the fact that this check
			// sees a valid no-deref expression such as 200-150 as invalid.
			value = ATOI(new_raw_arg3);
			if (value < 0 || value > MAX_MOUSE_SPEED)
				return ScriptError(ERR_PARAM3_INVALID, new_raw_arg3);
		}
		if (*new_raw_arg4 && !line.ArgHasDeref(4) && ctoupper(*new_raw_arg4) != 'R')
			return ScriptError(ERR_PARAM4_INVALID, new_raw_arg4);
		if (!line.ValidateMouseCoords(new_raw_arg1, new_raw_arg2))
			return ScriptError(ERR_MOUSE_COORD, new_raw_arg1);
		break;

	case ACT_MOUSECLICK:
		if (*NEW_RAW_ARG5 && !line.ArgHasDeref(5))
		{
			// The value of catching syntax errors at load-time seems to outweigh the fact that this check
			// sees a valid no-deref expression such as 200-150 as invalid.
			value = ATOI(NEW_RAW_ARG5);
			if (value < 0 || value > MAX_MOUSE_SPEED)
				return ScriptError(ERR_PARAM5_INVALID, NEW_RAW_ARG5);
		}
		if (*NEW_RAW_ARG6 && !line.ArgHasDeref(6))
			if (_tcslen(NEW_RAW_ARG6) > 1 || !_tcschr(_T("UD"), ctoupper(*NEW_RAW_ARG6)))  // Up / Down
				return ScriptError(ERR_PARAM6_INVALID, NEW_RAW_ARG6);
		if (*NEW_RAW_ARG7 && !line.ArgHasDeref(7) && ctoupper(*NEW_RAW_ARG7) != 'R')
			return ScriptError(ERR_PARAM7_INVALID, NEW_RAW_ARG7);
		// Check that the button is valid (e.g. left/right/middle):
		if (*new_raw_arg1 && !line.ArgHasDeref(1) && !line.ConvertMouseButton(new_raw_arg1)) // Treats blank as "Left".
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		if (!line.ValidateMouseCoords(new_raw_arg2, new_raw_arg3))
			return ScriptError(ERR_MOUSE_COORD, new_raw_arg2);
		break;

	case ACT_MOUSECLICKDRAG:
		// Even though we check for blanks here at load-time, we don't bother to do so at runtime
		// (i.e. if a dereferenced var resolved to blank, it will be treated as a zero):
		if (!*new_raw_arg4 || !*NEW_RAW_ARG5)
			return ScriptError(_T("Parameter #4 and 5 required."));
		if (*NEW_RAW_ARG6 && !line.ArgHasDeref(6))
		{
			// The value of catching syntax errors at load-time seems to outweigh the fact that this check
			// sees a valid no-deref expression such as 200-150 as invalid.
			value = ATOI(NEW_RAW_ARG6);
			if (value < 0 || value > MAX_MOUSE_SPEED)
				return ScriptError(ERR_PARAM6_INVALID, NEW_RAW_ARG6);
		}
		if (*NEW_RAW_ARG7 && !line.ArgHasDeref(7) && ctoupper(*NEW_RAW_ARG7) != 'R')
			return ScriptError(ERR_PARAM7_INVALID, NEW_RAW_ARG7);
		if (!line.ArgHasDeref(1))
			if (!line.ConvertMouseButton(new_raw_arg1, false))
				return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		if (!line.ValidateMouseCoords(new_raw_arg2, new_raw_arg3))
			return ScriptError(ERR_MOUSE_COORD, new_raw_arg2);
		if (!line.ValidateMouseCoords(new_raw_arg4, NEW_RAW_ARG5))
			return ScriptError(ERR_MOUSE_COORD, new_raw_arg4);
		break;

	case ACT_CONTROLSEND:
	case ACT_CONTROLSENDRAW:
		// Window params can all be blank in this case, but characters to send should
		// be non-blank (but it's ok if its a dereferenced var that resolves to blank
		// at runtime):
		if (!*new_raw_arg2)
			return ScriptError(ERR_PARAM2_REQUIRED);
		break;

	case ACT_CONTROLCLICK:
		// Check that the button is valid (e.g. left/right/middle):
		if (*new_raw_arg4 && !line.ArgHasDeref(4)) // i.e. it's allowed to be blank (defaults to left).
			if (!line.ConvertMouseButton(new_raw_arg4)) // Treats blank as "Left".
				return ScriptError(ERR_PARAM4_INVALID, new_raw_arg4);
		break;

	case ACT_FILEINSTALL:
	case ACT_FILECOPY:
	case ACT_FILEMOVE:
	case ACT_FILECOPYDIR:
	case ACT_FILEMOVEDIR:
		if (*new_raw_arg3 && !line.ArgHasDeref(3))
		{
			// The value of catching syntax errors at load-time seems to outweigh the fact that this check
			// sees a valid no-deref expression such as 2-1 as invalid.
			value = ATOI(new_raw_arg3);
			bool is_pure_numeric = IsPureNumeric(new_raw_arg3, false, true); // Consider negatives to be non-numeric.
			if (aActionType == ACT_FILEMOVEDIR)
			{
				if (!is_pure_numeric && ctoupper(*new_raw_arg3) != 'R'
					|| is_pure_numeric && value > 2) // IsPureNumeric() already checked if value < 0. 
					return ScriptError(ERR_PARAM3_INVALID, new_raw_arg3);
			}
			else
			{
				if (!is_pure_numeric || value > 1) // IsPureNumeric() already checked if value < 0.
					return ScriptError(ERR_PARAM3_INVALID, new_raw_arg3);
			}
		}
		if (aActionType == ACT_FILEINSTALL)
		{
			if (aArgc > 0 && line.ArgHasDeref(1))
				return ScriptError(_T("Must not contain variables."), new_raw_arg1);
		}
		break;

	case ACT_FILEREMOVEDIR:
		if (*new_raw_arg2 && !line.ArgHasDeref(2))
		{
			// The value of catching syntax errors at load-time seems to outweigh the fact that this check
			// sees a valid no-deref expression such as 3-2 as invalid.
			value = ATOI(new_raw_arg2);
			if (!IsPureNumeric(new_raw_arg2, false, true) || value > 1) // IsPureNumeric() prevents negatives.
				return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
		}
		break;

	case ACT_FILESETATTRIB:
		if (*new_raw_arg1 && !line.ArgHasDeref(1))
		{
			for (LPTSTR cp = new_raw_arg1; *cp; ++cp)
				if (!_tcschr(_T("+-^RASHNOT"), ctoupper(*cp)))
					return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		}
		// For the next two checks:
		// The value of catching syntax errors at load-time seems to outweigh the fact that this check
		// sees a valid no-deref expression such as 300-200 as invalid.
		if (aArgc > 2 && !line.ArgHasDeref(3) && line.ConvertLoopMode(new_raw_arg3) == FILE_LOOP_INVALID)
			return ScriptError(ERR_PARAM3_INVALID, new_raw_arg3);
		if (*new_raw_arg4 && !line.ArgHasDeref(4))
			if (_tcslen(new_raw_arg4) > 1 || (*new_raw_arg4 != '0' && *new_raw_arg4 != '1'))
				return ScriptError(ERR_PARAM4_INVALID, new_raw_arg4);
		break;

	case ACT_FILEGETTIME:
		if (*new_raw_arg3 && !line.ArgHasDeref(3))
			if (_tcslen(new_raw_arg3) > 1 || !_tcschr(_T("MCA"), ctoupper(*new_raw_arg3)))
				return ScriptError(ERR_PARAM3_INVALID, new_raw_arg3);
		break;

	case ACT_FILESETTIME:
		if (*new_raw_arg1 && !line.ArgHasDeref(1))
			if (!YYYYMMDDToSystemTime(new_raw_arg1, st, true))
				return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		if (*new_raw_arg3 && !line.ArgHasDeref(3))
			if (_tcslen(new_raw_arg3) > 1 || !_tcschr(_T("MCA"), ctoupper(*new_raw_arg3)))
				return ScriptError(ERR_PARAM3_INVALID, new_raw_arg3);
		// For the next two checks:
		// The value of catching syntax errors at load-time seems to outweigh the fact that this check
		// sees a valid no-deref expression such as 300-200 as invalid.
		if (aArgc > 3 && !line.ArgHasDeref(4) && line.ConvertLoopMode(new_raw_arg4) == FILE_LOOP_INVALID)
			return ScriptError(ERR_PARAM4_INVALID, new_raw_arg4);
		if (*NEW_RAW_ARG5 && !line.ArgHasDeref(5))
			if (_tcslen(NEW_RAW_ARG5) > 1 || (*NEW_RAW_ARG5 != '0' && *NEW_RAW_ARG5 != '1'))
				return ScriptError(ERR_PARAM5_INVALID, NEW_RAW_ARG5);
		break;

	case ACT_FILEGETSIZE:
		if (*new_raw_arg3 && !line.ArgHasDeref(3))
			if (_tcslen(new_raw_arg3) > 1 || !_tcschr(_T("BKM"), ctoupper(*new_raw_arg3))) // Allow B=Bytes as undocumented.
				return ScriptError(ERR_PARAM3_INVALID, new_raw_arg3);
		break;

	case ACT_SETTITLEMATCHMODE:
		if (aArgc > 0 && !line.ArgHasDeref(1) && !line.ConvertTitleMatchMode(new_raw_arg1))
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;

	case ACT_TRANSFORM:
		if (aArgc > 1 && !line.ArgHasDeref(2))
		{
			TransformCmds trans_cmd = Line::ConvertTransformCmd(new_raw_arg2);
			if (trans_cmd == TRANS_CMD_INVALID)
				return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
#ifndef UNICODE
			if (trans_cmd == TRANS_CMD_UNICODE && !*line.mArg[0].text) // blank text means output-var is not a dynamically built one.
			{
				// If the output var isn't the clipboard, the mode is "retrieve clipboard text as UTF-8".
				// Therefore, Param#3 should be blank in that case to avoid unnecessary fetching of the
				// entire clipboard contents as plain text when in fact the command itself will be
				// directly accessing the clipboard rather than relying on the automatic parameter and
				// deref handling.
				if (VAR(line.mArg[0])->Type() == VAR_CLIPBOARD)
				{
					if (aArgc < 3)
						return ScriptError(ERR_PARAM3_MUST_NOT_BE_BLANK);
				}
				else
					if (aArgc > 2)
						return ScriptError(ERR_PARAM3_MUST_BE_BLANK, new_raw_arg3);
				break; // This type has been fully checked above.
			}
#endif

			// The value of catching syntax errors at load-time seems to outweigh the fact that this check
			// sees a valid no-deref expression such as 1+2 as invalid.
			if (!line.ArgHasDeref(3)) // "true" since it might have been made into an InputVar due to being a simple expression.
			{
				switch(trans_cmd)
				{
				case TRANS_CMD_CHR:
				case TRANS_CMD_BITNOT:
				case TRANS_CMD_BITSHIFTLEFT:
				case TRANS_CMD_BITSHIFTRIGHT:
				case TRANS_CMD_BITAND:
				case TRANS_CMD_BITOR:
				case TRANS_CMD_BITXOR:
					if (!IsPureNumeric(new_raw_arg3, true, false))
						return ScriptError(_T("Parameter #3 must be an integer in this case."), new_raw_arg3);
					break;

				case TRANS_CMD_MOD:
				case TRANS_CMD_EXP:
				case TRANS_CMD_ROUND:
				case TRANS_CMD_CEIL:
				case TRANS_CMD_FLOOR:
				case TRANS_CMD_ABS:
				case TRANS_CMD_SIN:
				case TRANS_CMD_COS:
				case TRANS_CMD_TAN:
				case TRANS_CMD_ASIN:
				case TRANS_CMD_ACOS:
				case TRANS_CMD_ATAN:
					if (!IsPureNumeric(new_raw_arg3, true, false, true))
						return ScriptError(_T("Parameter #3 must be a number in this case."), new_raw_arg3);
					break;

				case TRANS_CMD_POW:
				case TRANS_CMD_SQRT:
				case TRANS_CMD_LOG:
				case TRANS_CMD_LN:
					if (!IsPureNumeric(new_raw_arg3, false, false, true))
						return ScriptError(_T("Parameter #3 must be a positive integer in this case."), new_raw_arg3);
					break;

				// The following are not listed above because no validation of Parameter #3 is needed at this stage:
				// TRANS_CMD_ASC
				// TRANS_CMD_UNICODE
				// TRANS_CMD_HTML
				// TRANS_CMD_DEREF
				}
			}

			switch(trans_cmd)
			{
			case TRANS_CMD_ASC:
			case TRANS_CMD_CHR:
			case TRANS_CMD_DEREF:
#ifndef UNICODE
			case TRANS_CMD_UNICODE:
			case TRANS_CMD_HTML:
#endif
			case TRANS_CMD_EXP:
			case TRANS_CMD_SQRT:
			case TRANS_CMD_LOG:
			case TRANS_CMD_LN:
			case TRANS_CMD_CEIL:
			case TRANS_CMD_FLOOR:
			case TRANS_CMD_ABS:
			case TRANS_CMD_SIN:
			case TRANS_CMD_COS:
			case TRANS_CMD_TAN:
			case TRANS_CMD_ASIN:
			case TRANS_CMD_ACOS:
			case TRANS_CMD_ATAN:
			case TRANS_CMD_BITNOT:
				if (*new_raw_arg4)
					return ScriptError(ERR_PARAM4_MUST_BE_BLANK, new_raw_arg4);
				break;

			case TRANS_CMD_BITAND:
			case TRANS_CMD_BITOR:
			case TRANS_CMD_BITXOR:
				if (!line.ArgHasDeref(4) && !IsPureNumeric(new_raw_arg4, true, false))
					return ScriptError(_T("Parameter #4 must be an integer in this case."), new_raw_arg4);
				break;

			case TRANS_CMD_BITSHIFTLEFT:
			case TRANS_CMD_BITSHIFTRIGHT:
				if (!line.ArgHasDeref(4) && !IsPureNumeric(new_raw_arg4, false, false))
					return ScriptError(_T("Parameter #4 must be a positive integer in this case."), new_raw_arg4);
				break;

			case TRANS_CMD_ROUND:
#ifdef UNICODE
			case TRANS_CMD_HTML:
#endif
				if (*new_raw_arg4 && !line.ArgHasDeref(4) && !IsPureNumeric(new_raw_arg4, true, false))
					return ScriptError(_T("Parameter #4 must be blank or an integer in this case."), new_raw_arg4);
				break;

			case TRANS_CMD_MOD:
			case TRANS_CMD_POW:
				if (!line.ArgHasDeref(4) && !IsPureNumeric(new_raw_arg4, true, false, true))
					return ScriptError(_T("Parameter #4 must be a number in this case."), new_raw_arg4);
				break;

#ifdef _DEBUG
			default:
				return ScriptError(_T("DEBUG: Unhandled"), new_raw_arg2);  // To improve maintainability.
#endif
			}

			switch(trans_cmd)
			{
			case TRANS_CMD_CHR:
				if (!line.ArgHasDeref(3))
				{
					value = ATOI(new_raw_arg3);
					if (!IsPureNumeric(new_raw_arg3, false, false) || value > 255) // IsPureNumeric() checks for value < 0 too.
						return ScriptError(ERR_PARAM3_INVALID, new_raw_arg3);
				}
				break;
			case TRANS_CMD_MOD:
				if (!line.ArgHasDeref(4) && !ATOF(new_raw_arg4)) // Parameter is omitted or something that resolves to zero.
					return ScriptError(ERR_DIVIDEBYZERO, new_raw_arg4);
				break;
			}
		}
		break;

	case ACT_MENU:
		if (aArgc > 1 && !line.ArgHasDeref(2))
		{
			MenuCommands menu_cmd = line.ConvertMenuCommand(new_raw_arg2);

			switch(menu_cmd)
			{
			case MENU_CMD_TIP:
			//case MENU_CMD_ICON: // L17: Now valid for other menus, used to set menu item icons.
			//case MENU_CMD_NOICON:
			case MENU_CMD_MAINWINDOW:
			case MENU_CMD_NOMAINWINDOW:
			case MENU_CMD_CLICK:
			{
				bool is_tray = true;  // Assume true if unknown.
				if (aArgc > 0 && !line.ArgHasDeref(1))
					if (_tcsicmp(new_raw_arg1, _T("tray")))
						is_tray = false;
				if (!is_tray)
					return ScriptError(ERR_MENUTRAY, new_raw_arg1);
				break;
			}
			}

			switch (menu_cmd)
			{
			case MENU_CMD_INVALID:
				return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);

			case MENU_CMD_NODEFAULT:
			case MENU_CMD_STANDARD:
			case MENU_CMD_NOSTANDARD:
			case MENU_CMD_DELETEALL:
			//case MENU_CMD_NOICON: // L17: Parameter #3 is now used to specify a menu item whose icon should be removed.
			case MENU_CMD_MAINWINDOW:
			case MENU_CMD_NOMAINWINDOW:
				if (*new_raw_arg3 || *new_raw_arg4 || *NEW_RAW_ARG5 || *NEW_RAW_ARG6)
					return ScriptError(_T("Parameter #3 and beyond should be omitted in this case."), new_raw_arg3);
				break;

			case MENU_CMD_RENAME:
			case MENU_CMD_USEERRORLEVEL:
			case MENU_CMD_CHECK:
			case MENU_CMD_UNCHECK:
			case MENU_CMD_TOGGLECHECK:
			case MENU_CMD_ENABLE:
			case MENU_CMD_DISABLE:
			case MENU_CMD_TOGGLEENABLE:
			case MENU_CMD_DEFAULT:
			case MENU_CMD_DELETE:
			case MENU_CMD_TIP:
			case MENU_CMD_CLICK:
			case MENU_CMD_NOICON: // L17: See comment in section above.
				if (   menu_cmd != MENU_CMD_RENAME && (*new_raw_arg4 || *NEW_RAW_ARG5 || *NEW_RAW_ARG6)   )
					return ScriptError(_T("Parameter #4 and beyond should be omitted in this case."), new_raw_arg4);
				switch(menu_cmd)
				{
				case MENU_CMD_USEERRORLEVEL:
				case MENU_CMD_TIP:
				case MENU_CMD_DEFAULT:
				case MENU_CMD_DELETE:
				case MENU_CMD_NOICON:
					break;  // i.e. for commands other than the above, do the default below.
				default:
					if (!*new_raw_arg3)
						return ScriptError(ERR_PARAM3_MUST_NOT_BE_BLANK);
				}
				break;

			// These have a highly variable number of parameters, or are too rarely used
			// to warrant detailed load-time checking, so are not validated here:
			//case MENU_CMD_SHOW:
			//case MENU_CMD_ADD:
			//case MENU_CMD_COLOR:
			//case MENU_CMD_ICON:
			}
		}
		break;

	case ACT_THREAD:
		if (aArgc > 0 && !line.ArgHasDeref(1) && !line.ConvertThreadCommand(new_raw_arg1))
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;

	case ACT_CONTROL:
		if (aArgc > 0 && !line.ArgHasDeref(1))
		{
			ControlCmds control_cmd = line.ConvertControlCmd(new_raw_arg1);
			switch (control_cmd)
			{
			case CONTROL_CMD_INVALID:
				return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
			case CONTROL_CMD_STYLE:
			case CONTROL_CMD_EXSTYLE:
			case CONTROL_CMD_TABLEFT:
			case CONTROL_CMD_TABRIGHT:
			case CONTROL_CMD_ADD:
			case CONTROL_CMD_DELETE:
			case CONTROL_CMD_CHOOSE:
			case CONTROL_CMD_CHOOSESTRING:
			case CONTROL_CMD_EDITPASTE:
				if (control_cmd != CONTROL_CMD_TABLEFT && control_cmd != CONTROL_CMD_TABRIGHT && !*new_raw_arg2)
					return ScriptError(ERR_PARAM2_MUST_NOT_BE_BLANK);
				break;
			default: // All commands except the above should have a blank Value parameter.
				if (*new_raw_arg2)
					return ScriptError(ERR_PARAM2_MUST_BE_BLANK, new_raw_arg2);
			}
		}
		break;

	case ACT_CONTROLGET:
		if (aArgc > 1 && !line.ArgHasDeref(2))
		{
			ControlGetCmds control_get_cmd = line.ConvertControlGetCmd(new_raw_arg2);
			switch (control_get_cmd)
			{
			case CONTROLGET_CMD_INVALID:
				return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
			case CONTROLGET_CMD_FINDSTRING:
			case CONTROLGET_CMD_LINE:
				if (!*new_raw_arg3)
					return ScriptError(ERR_PARAM3_MUST_NOT_BE_BLANK);
				break;
			case CONTROLGET_CMD_LIST:
				break; // Simply break for any sub-commands that have an optional parameter 3.
			default: // All commands except the above should have a blank Value parameter.
				if (*new_raw_arg3)
					return ScriptError(ERR_PARAM3_MUST_BE_BLANK, new_raw_arg3);
			}
		}
		break;

	case ACT_GUICONTROL:
		if (!*new_raw_arg2) // ControlID
			return ScriptError(ERR_PARAM2_REQUIRED);
		if (aArgc > 0 && !line.ArgHasDeref(1))
		{
			LPTSTR command, name;
			ResolveGui(new_raw_arg1, command, &name);
			if (!name)
				return ScriptError(ERR_INVALID_GUI_NAME, new_raw_arg1);

			GuiControlCmds guicontrol_cmd = line.ConvertGuiControlCmd(command);
			switch (guicontrol_cmd)
			{
			case GUICONTROL_CMD_INVALID:
				return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
			case GUICONTROL_CMD_CONTENTS:
			case GUICONTROL_CMD_TEXT:
			case GUICONTROL_CMD_MOVEDRAW:
			case GUICONTROL_CMD_OPTIONS:
				break; // Do nothing for the above commands since Param3 is optional.
			case GUICONTROL_CMD_MOVE:
			case GUICONTROL_CMD_CHOOSE:
			case GUICONTROL_CMD_CHOOSESTRING:
				if (!*new_raw_arg3)
					return ScriptError(ERR_PARAM3_MUST_NOT_BE_BLANK);
				break;
			default: // All commands except the above should have a blank Text parameter.
				if (*new_raw_arg3)
					return ScriptError(ERR_PARAM3_MUST_BE_BLANK, new_raw_arg3);
			}
		}
		break;

	case ACT_GUICONTROLGET:
		if (aArgc > 1 && !line.ArgHasDeref(2))
		{
			LPTSTR command, name;
			ResolveGui(new_raw_arg2, command, &name);
			if (!name)
				return ScriptError(ERR_INVALID_GUI_NAME, new_raw_arg2);

			GuiControlGetCmds guicontrolget_cmd = line.ConvertGuiControlGetCmd(command);
			// This first check's error messages take precedence over the next check's:
			switch (guicontrolget_cmd)
			{
			case GUICONTROLGET_CMD_INVALID:
				return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
			case GUICONTROLGET_CMD_CONTENTS:
				break; // Do nothing, since Param4 is optional in this case.
			default: // All commands except the above should have a blank parameter here.
				if (*new_raw_arg4) // Currently true for all, since it's a FutureUse param.
					return ScriptError(ERR_PARAM4_MUST_BE_BLANK, new_raw_arg4);
			}
			if (guicontrolget_cmd == GUICONTROLGET_CMD_FOCUS || guicontrolget_cmd == GUICONTROLGET_CMD_FOCUSV)
			{
				if (*new_raw_arg3)
					return ScriptError(ERR_PARAM3_MUST_BE_BLANK, new_raw_arg3);
			}
			// else it can be optionally blank, in which case the output variable is used as the
			// ControlID also.
		}
		break;

	case ACT_DRIVE:
		if (aArgc > 0 && !line.ArgHasDeref(1))
		{
			DriveCmds drive_cmd = line.ConvertDriveCmd(new_raw_arg1);
			if (!drive_cmd)
				return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
			if (drive_cmd != DRIVE_CMD_EJECT && !*new_raw_arg2)
				return ScriptError(ERR_PARAM2_MUST_NOT_BE_BLANK);
			// For DRIVE_CMD_LABEL: Note that it is possible and allowed for the new label to be blank.
			// Not currently done since all sub-commands take a mandatory or optional ARG3:
			//if (drive_cmd != ... && *new_raw_arg3)
			//	return ScriptError(ERR_PARAM3_MUST_BE_BLANK, new_raw_arg3);
		}
		break;

	case ACT_DRIVEGET:
		if (!line.ArgHasDeref(2))  // Don't check "aArgc > 1" because of DRIVEGET_CMD_SETLABEL's param format.
		{
			DriveGetCmds drive_get_cmd = line.ConvertDriveGetCmd(new_raw_arg2);
			if (!drive_get_cmd)
				return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
			if (drive_get_cmd != DRIVEGET_CMD_LIST && drive_get_cmd != DRIVEGET_CMD_STATUSCD && !*new_raw_arg3)
				return ScriptError(ERR_PARAM3_MUST_NOT_BE_BLANK);
			if (drive_get_cmd != DRIVEGET_CMD_SETLABEL && (aArgc < 1 || line.mArg[0].type == ARG_TYPE_NORMAL))
				// The output variable has been omitted.
				return ScriptError(ERR_PARAM1_MUST_NOT_BE_BLANK);
		}
		break;

	case ACT_PROCESS:
		if (aArgc > 0 && !line.ArgHasDeref(1))
		{
			ProcessCmds process_cmd = line.ConvertProcessCmd(new_raw_arg1);
			if (process_cmd != PROCESS_CMD_PRIORITY && process_cmd != PROCESS_CMD_EXIST && !*new_raw_arg2)
				return ScriptError(ERR_PARAM2_MUST_NOT_BE_BLANK);
			switch (process_cmd)
			{
			case PROCESS_CMD_INVALID:
				return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
			case PROCESS_CMD_EXIST:
			case PROCESS_CMD_CLOSE:
				if (*new_raw_arg3)
					return ScriptError(ERR_PARAM3_MUST_BE_BLANK, new_raw_arg3);
				break;
			case PROCESS_CMD_PRIORITY:
				if (!*new_raw_arg3 || (!line.ArgHasDeref(3) && !_tcschr(PROCESS_PRIORITY_LETTERS, ctoupper(*new_raw_arg3))))
					return ScriptError(ERR_PARAM3_INVALID, new_raw_arg3);
				break;
			case PROCESS_CMD_WAIT:
			case PROCESS_CMD_WAITCLOSE:
				if (*new_raw_arg3 && !line.ArgHasDeref(3) && !IsPureNumeric(new_raw_arg3, false, true, true))
					return ScriptError(_T("If present, parameter #3 must be a positive number in this case."), new_raw_arg3);
				break;
			}
		}
		break;

	// For ACT_WINMOVE, don't validate anything for mandatory args so that its two modes of
	// operation can be supported: 2-param mode and normal-param mode.
	// For these, although we validate that at least one is non-blank here, it's okay at
	// runtime for them all to resolve to be blank, without an error being reported.
	// It's probably more flexible that way since the commands are equipped to handle
	// all-blank params.
	// Not these because they can be used with the "last-used window" mode:
	//case ACT_IFWINEXIST:
	//case ACT_IFWINNOTEXIST:
	// Not these because they can have their window params all-blank to work in "last-used window" mode:
	//case ACT_IFWINACTIVE:
	//case ACT_IFWINNOTACTIVE:
	//case ACT_WINACTIVATE:
	//case ACT_WINWAITCLOSE:
	//case ACT_WINWAITACTIVE:
	//case ACT_WINWAITNOTACTIVE:
	case ACT_WINACTIVATEBOTTOM:
		if (!*new_raw_arg1 && !*new_raw_arg2 && !*new_raw_arg3 && !*new_raw_arg4)
			return ScriptError(ERR_WINDOW_PARAM);
		break;

	case ACT_WINWAIT:
		if (!*new_raw_arg1 && !*new_raw_arg2 && !*new_raw_arg4 && !*NEW_RAW_ARG5) // ARG3 is omitted because it's the timeout.
			return ScriptError(ERR_WINDOW_PARAM);
		break;

	case ACT_WINMENUSELECTITEM:
		// Window params can all be blank in this case, but the first menu param should
		// be non-blank (but it's ok if its a dereferenced var that resolves to blank
		// at runtime):
		if (!*new_raw_arg3)
			return ScriptError(ERR_PARAM3_REQUIRED);
		break;

	case ACT_WINSET:
		if (aArgc > 0 && !line.ArgHasDeref(1))
		{
			switch(line.ConvertWinSetAttribute(new_raw_arg1))
			{
			case WINSET_TRANSPARENT:
				if (aArgc > 1 && !line.ArgHasDeref(2))
				{
					value = ATOI(new_raw_arg2);
					if (value < 0 || value > 255)
						return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
				}
				break;
			case WINSET_TRANSCOLOR:
				if (!*new_raw_arg2)
					return ScriptError(ERR_PARAM2_MUST_NOT_BE_BLANK);
				break;
			case WINSET_ALWAYSONTOP:
				if (aArgc > 1 && !line.ArgHasDeref(2) && !line.ConvertOnOffToggle(new_raw_arg2))
					return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
				break;
			case WINSET_BOTTOM:
			case WINSET_TOP:
			case WINSET_REDRAW:
			case WINSET_ENABLE:
			case WINSET_DISABLE:
				if (*new_raw_arg2)
					return ScriptError(ERR_PARAM2_MUST_BE_BLANK);
				break;
			case WINSET_INVALID:
				return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
			}
		}
		break;

	case ACT_WINGET:
		if (!line.ArgHasDeref(2) && !line.ConvertWinGetCmd(new_raw_arg2)) // It's okay if ARG2 is blank.
			return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
		break;

	case ACT_SYSGET:
		if (!line.ArgHasDeref(2) && !line.ConvertSysGetCmd(new_raw_arg2))
			return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
		break;

	case ACT_INPUTBOX:
		if (*NEW_RAW_ARG9)  // && !line.ArgHasDeref(9)
			return ScriptError(_T("Parameter #9 must be blank."), NEW_RAW_ARG9);
		break;

	case ACT_IFMSGBOX:
		if (aArgc > 0 && !line.ArgHasDeref(1) && !line.ConvertMsgBoxResult(new_raw_arg1))
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;

	case ACT_IFIS:
	case ACT_IFISNOT:
		if (aArgc > 1 && !line.ArgHasDeref(2) && !line.ConvertVariableTypeName(new_raw_arg2))
			// Don't refer to it as "Parameter #2" because this command isn't formatted/displayed that way.
			// Update: Param2 is more descriptive than the other (short) alternatives:
			return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
		break;

	case ACT_GETKEYSTATE:
		// v1.0.44.03: Don't validate single-character key names because although a character like ü might have no
		// matching VK in system's default layout, that layout could change to something which does have a VK for it.
		if (aArgc > 1 && !line.ArgHasDeref(2) && _tcslen(new_raw_arg2) > 1 && !TextToVK(new_raw_arg2) && !ConvertJoy(new_raw_arg2))
			return ScriptError(ERR_PARAM2_INVALID, new_raw_arg2);
		break;

	case ACT_KEYWAIT: // v1.0.44.03: See comment above.
		if (aArgc > 0 && !line.ArgHasDeref(1) && _tcslen(new_raw_arg1) > 1 && !TextToVK(new_raw_arg1) && !ConvertJoy(new_raw_arg1))
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;

	case ACT_FILEAPPEND:
		if (aArgc > 2 && !line.ArgHasDeref(3) && line.ConvertFileEncoding(new_raw_arg3) == -1)
			return ScriptError(ERR_PARAM3_INVALID, new_raw_arg3);
		break;
	case ACT_FILEENCODING:
		if (aArgc > 0 && !line.ArgHasDeref(1) && line.ConvertFileEncoding(new_raw_arg1) == -1)
			return ScriptError(ERR_PARAM1_INVALID, new_raw_arg1);
		break;
#endif  // The above section is in place only if when not AUTOHOTKEYSC.
	}

	if (mNextLineIsFunctionBody && do_update_labels) // do_update_labels: false for '#if expr' and 'static var:=expr', neither of which should be treated as part of the function's body.
	{
		g->CurrentFunc->mJumpToLine = the_new_line;
		mNextLineIsFunctionBody = false;
		if (g->CurrentFunc->mDefaultVarType == VAR_DECLARE_NONE)
			g->CurrentFunc->mDefaultVarType = VAR_DECLARE_LOCAL;  // Set default since no override was discovered at the top of the body.
	}

	// No checking for unbalanced blocks is done here.  That is done by PreparseBlocks() because
	// it displays more informative error messages:
	if (aActionType == ACT_BLOCK_BEGIN)
	{
		++mCurrentFuncOpenBlockCount; // It's okay to increment unconditionally because it is reset to zero every time a new function definition is entered.
		// It's only necessary to check the last func, not the one(s) that come before it, to see if its
		// mJumpToLine is NULL.  This is because our caller has made it impossible for a function
		// to ever have been defined in the first place if it lacked its opening brace.  Search on
		// "consecutive function" for more comments.  In addition, the following does not check
		// that mCurrentFuncOpenBlockCount is exactly 1, because: 1) Want to be able to support function
		// definitions inside of other function definitions (to help script maintainability); 2) If
		// mCurrentFuncOpenBlockCount is 0 or negative, that will be caught as a syntax error by PreparseBlocks(),
		// which yields a more informative error message that we could here.
		if (g->CurrentFunc && !g->CurrentFunc->mJumpToLine)
		{
			// The above check relies upon the fact that g->CurrentFunc->mIsBuiltIn must be false at this stage,
			// which is the case because any non-overridden built-in function won't get added until after all
			// lines have been added, namely PreparseBlocks().
			line.mAttribute = ATTR_TRUE;  // Flag this ACT_BLOCK_BEGIN as the opening brace of the function's body.
			// For efficiency, and to prevent ExecUntil from starting a new recursion layer for the function's
			// body, the function's execution should begin at the first line after its open-brace (even if that
			// first line is another open-brace or the function's close-brace (i.e. an empty function):
			mNextLineIsFunctionBody = true;
		}
	}
	else if (aActionType == ACT_BLOCK_END)
	{
		--mCurrentFuncOpenBlockCount; // It's okay to increment unconditionally because it is reset to zero every time a new function definition is entered.
		if (g->CurrentFunc && !mCurrentFuncOpenBlockCount) // Any negative mCurrentFuncOpenBlockCount is caught by a different stage.
		{
			Func &func = *g->CurrentFunc;
			// After this point, mGlobalVar is used to prevent dynamic variable references from
			// resolving to globals which aren't declared in this function (though in v1, this is
			// only done in force-local functions, added in v1.1.27).
			if (func.mGlobalVarCount && (func.mDefaultVarType & VAR_FORCE_LOCAL))
			{
				// Now that there can be no more "global" declarations, copy the list into persistent memory.
				Var **global_vars;
				if (  !(global_vars = (Var **)SimpleHeap::Malloc(func.mGlobalVarCount * sizeof(Var *)))  )
					return ScriptError(ERR_OUTOFMEM);
				memcpy(global_vars, func.mGlobalVar, func.mGlobalVarCount * sizeof(Var *));
				func.mGlobalVar = global_vars;
			}
			else
			{
				func.mGlobalVarCount = 0;
				func.mGlobalVar = NULL;
			}
			line.mAttribute = ATTR_TRUE;  // Flag this ACT_BLOCK_END as the ending brace of a function's body.
			g->CurrentFunc = NULL;
		}
	}

	// Above must be done prior to the below, since it sometimes sets mAttribute for use below.

	///////////////////////////////////////////////////////////////
	// Update any labels that should refer to the newly added line.
	///////////////////////////////////////////////////////////////
	// If the label most recently added doesn't yet have an anchor in the script, provide it.
	// UPDATE: In addition, keep searching backward through the labels until a non-NULL
	// mJumpToLine is found.  All the ones with a NULL should point to this new line to
	// support cases where one label immediately follows another in the script.
	// Example:
	// #a::  <-- don't leave this label with a NULL jumppoint.
	// LaunchA:
	// ...
	// return
	if (do_update_labels)
	{
		for (Label *label = mLastLabel; label != NULL && label->mJumpToLine == NULL; label = label->mPrevLabel)
		{
			if (line.mActionType == ACT_ELSE || line.mActionType == ACT_UNTIL || line.mActionType == ACT_CATCH)
				return ScriptError(_T("A label must not point to an ELSE or UNTIL or CATCH."));
			// The following is inaccurate; each block-end is in fact owned by its block-begin
			// and not the block that encloses them both, so this restriction is unnecessary.
			// THE COMMENT BELOW IS OBSOLETE:
			// Don't allow this because it may cause problems in a case such as this because
			// label1 points to the end-block which is at the same level (and thus normally
			// an allowable jumppoint) as the goto.  But we don't want to allow jumping into
			// a block that belongs to a control structure.  In this case, it would probably
			// result in a runtime error when the execution unexpectedly encounters the ELSE
			// after performing the goto:
			// goto, label1
			// if x
			// {
			//    ...
			//    label1:
			// }
			// else
			//    ...
			//
			// An alternate way to deal with the above would be to make each block-end be owned
			// by its block-begin rather than the block that encloses them both.
			//if (line.mActionType == ACT_BLOCK_END)
			//	return ScriptError(_T("A label must not point to the end of a block. For loops, use Continue vs. Goto."));
			label->mJumpToLine = the_new_line;
		}
	}

	return OK;
}



ResultType Script::ParseDerefs(LPTSTR aArgText, LPTSTR aArgMap, DerefType *aDeref, int &aDerefCount)
// Caller provides modifiable aDerefCount, which might be non-zero to indicate that there are already
// some items in the aDeref array.
// Returns FAIL or OK.
{
	size_t deref_string_length; // So that overflow can be detected, this is not of type DerefLengthType.

	// For each dereference found in aArgText:
	for (int j = 0;; ++j)  // Increment to skip over the symbol just found by the inner for().
	{
		// Find next non-literal g_DerefChar:
		for (; aArgText[j] && (aArgText[j] != g_DerefChar || (aArgMap && aArgMap[j])); ++j);
		if (!aArgText[j])
			break;
		// else: Match was found; this is the deref's open-symbol.
		if (aDerefCount >= MAX_DEREFS_PER_ARG)
			return ScriptError(TOO_MANY_REFS, aArgText); // Short msg since so rare.
		DerefType &this_deref = aDeref[aDerefCount];  // For performance.
		this_deref.marker = aArgText + j;  // Store the deref's starting location.
		// Find next g_DerefChar, even if it's a literal.
		for (++j; aArgText[j] && aArgText[j] != g_DerefChar; ++j);
		if (!aArgText[j])
			return ScriptError(_T("This parameter contains a variable name missing its ending percent sign."), aArgText);
		// Otherwise: Match was found; this should be the deref's close-symbol.
		if (aArgMap && aArgMap[j])  // But it's mapped as literal g_DerefChar.
			return ScriptError(_T("Invalid `%."), aArgText); // Short msg. since so rare.
		deref_string_length = aArgText + j - this_deref.marker + 1;
		if (deref_string_length == 2) // The percent signs were empty, e.g. %%
			return ScriptError(_T("Empty variable reference (%%)."), aArgText); // Short msg. since so rare.
		if (deref_string_length - 2 > MAX_VAR_NAME_LENGTH) // -2 for the opening & closing g_DerefChars
			return ScriptError(_T("Variable name too long."), aArgText); // Short msg. since so rare.
		this_deref.is_function = false;
		this_deref.length = (DerefLengthType)deref_string_length;
		if (   !(this_deref.var = FindOrAddVar(this_deref.marker + 1, this_deref.length - 2))   )
			return FAIL;  // The called function already displayed the error.
		++aDerefCount;
	} // for each dereference.

	return OK;
}



ResultType Script::DefineFunc(LPTSTR aBuf, Var *aFuncGlobalVar[])
// Returns OK or FAIL.
// Caller has already called ValidateName() on the function, and it is known that this valid name
// is followed immediately by an open-paren.  aFuncExceptionVar is the address of an array on
// the caller's stack that will hold the list of exception variables (those that must be explicitly
// declared as either local or global) within the body of the function.
{
	LPTSTR param_end, param_start = _tcschr(aBuf, '('); // Caller has ensured that this will return non-NULL.
	int insert_pos;
	
	if (mClassObjectCount) // Class method or property getter/setter.
	{
		Object *class_object = mClassObject[mClassObjectCount - 1];

		*param_start = '\0'; // Temporarily terminate, for simplicity.

		// Build the fully-qualified method name for A_ThisFunc and ListVars:
		// AddFunc() enforces a limit of MAX_VAR_NAME_LENGTH characters for function names, which is relied
		// on by FindFunc(), BIF_OnMessage() and perhaps others.  For simplicity, allow one extra char to be
		// printed and rely on AddFunc() detecting that the name is too long.
		TCHAR full_name[MAX_VAR_NAME_LENGTH + 1 + 1]; // Extra +1 for null terminator.
		_sntprintf(full_name, MAX_VAR_NAME_LENGTH + 1, _T("%s.%s"), mClassName, aBuf);
		full_name[MAX_VAR_NAME_LENGTH + 1] = '\0'; // Must terminate at this exact point if _sntprintf hit the limit.

		// Check for duplicates and determine insert_pos:
		Func *found_func;
		ExprTokenType found_item;
		if (!mClassProperty && class_object->GetItem(found_item, aBuf) // Must be done in addition to the below to detect conflicting var/method declarations.
			|| (found_func = FindFunc(full_name, 0, &insert_pos))) // Must be done to determine insert_pos.
		{
			return ScriptError(ERR_DUPLICATE_DECLARATION, aBuf); // The parameters are omitted due to temporary termination above.  This might make it more obvious why "X[]" and "X()" are considered duplicates.
		}
		
		*param_start = '('; // Undo temporary termination.

		// Below passes class_object for AddFunc() to store the func "by reference" in it:
		if (  !(g->CurrentFunc = AddFunc(full_name, 0, false, insert_pos, class_object))  )
			return FAIL;
	}
	else
	{
		Func *found_func = FindFunc(aBuf, param_start - aBuf, &insert_pos);
		if (found_func)
		{
			if (!found_func->mIsBuiltIn)
				return ScriptError(_T("Duplicate function definition."), aBuf); // Seems more descriptive than "Function already defined."
			else // It's a built-in function that the user wants to override with a custom definition.
			{
				found_func->mIsBuiltIn = false;  // Override built-in with custom.
				found_func->mParamCount = 0; // Revert to the default appropriate for non-built-in functions.
				found_func->mMinParams = 0;  //
				found_func->mJumpToLine = NULL; // Fixed for v1.0.35.12: Must reset for detection elsewhere.
				g->CurrentFunc = found_func;
			}
		}
		else
			// The value of g->CurrentFunc must be set here rather than by our caller since AddVar(), which we call,
			// relies upon it having been done.
			if (   !(g->CurrentFunc = AddFunc(aBuf, param_start - aBuf, false, insert_pos))   )
				return FAIL; // It already displayed the error.
	}

	mCurrentFuncOpenBlockCount = 0; // v1.0.48.01: Initializing this here makes function definitions work properly when they're inside a block.
	Func &func = *g->CurrentFunc; // For performance and convenience.
	size_t param_length, value_length;
	FuncParam param[MAX_FUNCTION_PARAMS];
	int param_count = 0;
	TCHAR buf[LINE_SIZE], *target;
	bool param_must_have_default = false;

	if (mClassObjectCount)
	{
		// Add the automatic/hidden "this" parameter.
		if (  !(param[0].var = FindOrAddVar(_T("this"), 4, VAR_DECLARE_LOCAL | VAR_LOCAL_FUNCPARAM))  )
			return FAIL;
		param[0].default_type = PARAM_DEFAULT_NONE;
		param[0].is_byref = false;
		++param_count;
		++func.mMinParams;

		if (mClassProperty && toupper(param_start[-3]) == 'S') // Set
		{
			// Unlike __Set, the property setters accept the value as the second parameter,
			// so that we can make it automatic here and make variadic properties easier.
			if (  !(param[1].var = FindOrAddVar(_T("value"), 5, VAR_DECLARE_LOCAL | VAR_LOCAL_FUNCPARAM))  )
				return FAIL;
			param[1].default_type = PARAM_DEFAULT_NONE;
			param[1].is_byref = false;
			++param_count;
			++func.mMinParams;
		}
	}

	for (param_start = omit_leading_whitespace(param_start + 1);;)
	{
		if (*param_start == ')') // No more params.
			break;

		// Must start the search at param_start, not param_start+1, so that something like fn(, x) will be properly handled:
		if (   !*param_start || !(param_end = StrChrAny(param_start, _T(", \t:=*)")))   ) // Look for first comma, space, tab, =, or close-paren.
			return ScriptError(ERR_MISSING_CLOSE_PAREN, aBuf);

		if (param_count >= MAX_FUNCTION_PARAMS)
			return ScriptError(_T("Too many params."), param_start); // Short msg since so rare.
		FuncParam &this_param = param[param_count]; // For performance and convenience.

		// To enhance syntax error catching, consider ByRef to be a keyword; i.e. that can never be the name
		// of a formal parameter:
		if (this_param.is_byref = !tcslicmp(param_start, _T("ByRef"), param_end - param_start)) // ByRef.
		{
			// Omit the ByRef keyword from further consideration:
			param_start = omit_leading_whitespace(param_end);
			if (   !*param_start || !(param_end = StrChrAny(param_start, _T(", \t:=*)")))   ) // Look for first comma, space, tab, =, or close-paren.
				return ScriptError(ERR_MISSING_CLOSE_PAREN, aBuf);
		}

		if (   !(param_length = param_end - param_start)   )
			return ScriptError(ERR_BLANK_PARAM, aBuf); // Reporting aBuf vs. param_start seems more informative since Vicinity isn't shown.

		if (this_param.var = FindVar(param_start, param_length, &insert_pos, FINDVAR_LOCAL))  // Assign.
			return ScriptError(_T("Duplicate parameter."), param_start);
		if (   !(this_param.var = AddVar(param_start, param_length, insert_pos, VAR_DECLARE_LOCAL | VAR_LOCAL_FUNCPARAM))   )	// Pass VAR_LOCAL_FUNCPARAM as last parameter to mean "it's a local but more specifically a function's parameter".
			return FAIL; // It already displayed the error, including attempts to have reserved names as parameter names.
		
		this_param.default_type = PARAM_DEFAULT_NONE;  // Set default.
		param_start = omit_leading_whitespace(param_end);

		if (func.mIsVariadic = (*param_start == '*'))
		{
			param_start = omit_leading_whitespace(param_start + 1);
			if (*param_start != ')')
				// Give vague error message since the user's intent isn't clear.
				return ScriptError(ERR_MISSING_CLOSE_PAREN, param_start);
			// Although this param must be counted since it needs a FuncParam in the array,
			// it doesn't count toward func.mMinParams or func.mParamCount.
			++param_count;
			break;
		}

		// v1.1.09: Allow optional parameters to be declared with := instead of =.
		if (*param_start == ':')
		{
			if (param_start[1] != '=')
				return ScriptError(ERR_BAD_OPTIONAL_PARAM, param_start);
			++param_start;
		}

		// v1.0.35: Check if a default value is specified for this parameter and set up for the next iteration.
		// The following section is similar to that used to support initializers for static variables.
		// So maybe maintain them together.
		if (*param_start == '=') // This is the default value of the param just added.
		{
			param_start = omit_leading_whitespace(param_start + 1); // Start of the default value.
			if (*param_start == '"') // Quoted literal string, or the empty string.
			{
				// v1.0.46.13: Added support for quoted/literal strings beyond simply "".
				// The following section is nearly identical to one in ExpandExpression().
				// Find the end of this string literal, noting that a pair of double quotes is
				// a literal double quote inside the string.
				for (target = buf, param_end = param_start + 1;;) // Omit the starting-quote from consideration, and from the resulting/built string.
				{
					if (!*param_end) // No matching end-quote. Probably impossible due to load-time validation.
						return ScriptError(ERR_MISSING_CLOSE_QUOTE, param_start); // Reporting param_start vs. aBuf seems more informative in the case of quoted/literal strings.
					if (*param_end == '"') // And if it's not followed immediately by another, this is the end of it.
					{
						++param_end;
						if (*param_end != '"') // String terminator or some non-quote character.
							break;  // The previous char is the ending quote.
						//else a pair of quotes, which resolves to a single literal quote. So fall through
						// to the below, which will copy of quote character to the buffer. Then this pair
						// is skipped over and the loop continues until the real end-quote is found.
					}
					//else some character other than '\0' or '"'.
					*target++ = *param_end++;
				}
				*target = '\0'; // Terminate it in the buffer.
				// The above has also set param_end for use near the bottom of the loop.
				ConvertEscapeSequences(buf, g_EscapeChar, false); // Raw escape sequences like `n haven't been converted yet, so do it now.
				this_param.default_type = PARAM_DEFAULT_STR;
				this_param.default_str = *buf ? SimpleHeap::Malloc(buf, target-buf) : _T("");
			}
			else // A default value other than a quoted/literal string.
			{
				if (!(param_end = StrChrAny(param_start, _T(", \t=)")))) // Somewhat debatable but stricter seems better.
					return ScriptError(ERR_MISSING_COMMA, aBuf); // Reporting aBuf vs. param_start seems more informative since Vicinity isn't shown.
				value_length = param_end - param_start;
				if (value_length > MAX_NUMBER_LENGTH) // Too rare to justify elaborate handling or error reporting.
					value_length = MAX_NUMBER_LENGTH;
				tcslcpy(buf, param_start, value_length + 1);  // Make a temp copy to simplify the below (especially IsPureNumeric).
				if (!_tcsicmp(buf, _T("false")))
				{
					this_param.default_type = PARAM_DEFAULT_INT;
					this_param.default_int64 = 0;
				}
				else if (!_tcsicmp(buf, _T("true")))
				{
					this_param.default_type = PARAM_DEFAULT_INT;
					this_param.default_int64 = 1;
				}
				else // The only things supported other than the above are integers and floats.
				{
					// Vars could be supported here via FindVar(), but only globals ABOVE this point in
					// the script would be supported (since other globals don't exist yet). So it seems
					// best to wait until full/comprehensive support for expressions is studied/designed
					// for both static initializers and parameter-default-values.
					switch(IsPureNumeric(buf, true, false, true))
					{
					case PURE_INTEGER:
						// It's always been somewhat inconsistent that for parameter default values,
						// numbers like 0xFF and 0123 do not preserve their formatting (unlike func(0123)
						// and y:=0xFF, which do preserve it). But for backward compatibility and
						// performance, it seems best to keep it this way.
						this_param.default_type = PARAM_DEFAULT_INT;
						this_param.default_int64 = ATOI64(buf);
						break;
					case PURE_FLOAT:
						this_param.default_type = PARAM_DEFAULT_FLOAT;
						this_param.default_double = ATOF(buf);
						break;
					default: // Not numeric (and also not a quoted string because that was handled earlier).
						return ScriptError(_T("Unsupported parameter default."), aBuf);
					}
				}
			}
			param_must_have_default = true;  // For now, all other params after this one must also have default values.
			// Set up for the next iteration:
			param_start = omit_leading_whitespace(param_end);
		}
		else // This parameter does not have a default value specified.
		{
			if (param_must_have_default)
				return ScriptError(_T("Parameter default required."), this_param.var->mName);
			++func.mMinParams;
		}
		++param_count;

		if (*param_start != ',' && *param_start != ')') // Something like "fn(a, b c)" (missing comma) would cause this.
			return ScriptError(ERR_MISSING_COMMA, aBuf); // Reporting aBuf vs. param_start seems more informative since Vicinity isn't shown.
		if (*param_start == ',')
		{
			param_start = omit_leading_whitespace(param_start + 1);
			if (*param_start == ')') // If *param_start is ',' it will be caught as an error by the next iteration.
				return ScriptError(ERR_BLANK_PARAM, aBuf); // Reporting aBuf vs. param_start seems more informative since Vicinity isn't shown.
		}
		//else it's ')', in which case the next iteration will handle it.
		// Above has ensured that param_start now points to the next parameter, or ')' if none.
	} // for() each formal parameter.

	if (param_start[1]) // Something follows the ')' other than OTB (which was handled at an earlier stage).
		return ScriptError(ERR_INVALID_FUNCDECL, aBuf);

	if (param_count)
	{
		// Allocate memory only for the actual number of parameters actually present.
		size_t size = param_count * sizeof(param[0]);
		if (   !(func.mParam = (FuncParam *)SimpleHeap::Malloc(size))   )
			return ScriptError(ERR_OUTOFMEM);
		func.mParamCount = param_count - func.mIsVariadic; // i.e. don't count the final "param*" of a variadic function.
		memcpy(func.mParam, param, size);
	}
	//else leave func.mParam/mParamCount set to their NULL/0 defaults.

	if (mLastLabel && !mLastLabel->mJumpToLine)
	{
		// Check all variants of all hotkeys, since there might be multiple variants
		// of various hotkeys defined in a row, such as:
		// ^a::
		// #IfWinActive B
		// ^a::
		// ^b::
		//    somefunc(){ ...
		for (int i = 0; i < Hotkey::sHotkeyCount; ++i)
		{
			for (HotkeyVariant *v = Hotkey::shk[i]->mFirstVariant; v; v = v->mNextVariant)
			{
				Label *label = v->mJumpToLabel->ToLabel(); // Might be a function.
				if (label && !label->mJumpToLine) // This hotkey label is pointing at this function.
				{
					// Update the hotkey to use this function instead of the label.
					v->mJumpToLabel = &func;

					// Remove this hotkey label from the list.  Each label is removed as the corresponding
					// hotkey variant is found so that any generic labels that might be mixed in are left
					// in the list and detected as errors later.
					RemoveLabel(label);
				}
			}
		}
		// Check hotstrings as well (even if a hotkey was found):
		for (int i = Hotstring::sHotstringCount - 1; i >= 0; --i) // Start with the last one defined, for performance.
		{
			Label *label = Hotstring::shs[i]->mJumpToLabel->ToLabel(); // Might be a function.
			if (!label || label->mJumpToLine)
				// This hotstring has a function or a label which has already been resolved.
				// Since hotstrings are listed in order of definition and we're iterating in
				// the reverse order, there's no need to continue.
				break;
			// See hotkey section above for comments.
			Hotstring::shs[i]->mJumpToLabel = &func;
			RemoveLabel(label);
		}
		if (mLastLabel && !mLastLabel->mJumpToLine)
			// There are one or more non-hotkey labels pointing at this function.
			return ScriptError(_T("A label must not point to a function."), mLastLabel->mName);
		// Since above didn't return, the label or labels must have been hotkey labels.
		if (func.mMinParams)
			return ScriptError(ERR_HOTKEY_FUNC_PARAMS, aBuf);
	}

	// Indicate success:
	func.mGlobalVar = aFuncGlobalVar; // Give func.mGlobalVar its address, to be used for any var declarations inside this function's body.
	return OK;
}



ResultType Script::DefineClass(LPTSTR aBuf)
{
	if (mClassObjectCount == MAX_NESTED_CLASSES)
		return ScriptError(_T("This class definition is nested too deep."), aBuf);

	LPTSTR cp, class_name = aBuf;
	Object *outer_class, *base_class = NULL;
	Object *&class_object = mClassObject[mClassObjectCount]; // For convenience.
	Var *class_var;
	ExprTokenType token;

	for (cp = aBuf; *cp && !IS_SPACE_OR_TAB(*cp); ++cp);
	if (*cp)
	{
		*cp = '\0'; // Null-terminate here for class_name.
		cp = omit_leading_whitespace(cp + 1);
		if (_tcsnicmp(cp, _T("extends"), 7) || !IS_SPACE_OR_TAB(cp[7]))
			return ScriptError(_T("Syntax error in class definition."), cp);
		LPTSTR base_class_name = omit_leading_whitespace(cp + 8);
		if (!*base_class_name)
			return ScriptError(_T("Missing class name."), cp);
		if (  !(base_class = FindClass(base_class_name))  )
		{
			// This class hasn't been defined yet, but it might be.  Automatically create the
			// class, but store it in the "unresolved" list.  When its definition is encountered,
			// it will be removed from the list.  If any classes remain in the list when the end
			// of the script is reached, an error will be thrown.
			if (mUnresolvedClasses && mUnresolvedClasses->GetItem(token, base_class_name))
			{
				// Some other class has already referenced it.  Use the existing object:
				base_class = (Object *)token.object;
			}
			else
			{
				if (  !mUnresolvedClasses && !(mUnresolvedClasses = Object::Create())
					|| !(base_class = Object::Create())
					// Storing the file/line index in "__Class" instead of something like "DBG" or
					// two separate fields helps to reduce code size and maybe memory fragmentation.
					// It will be overwritten when the class definition is encountered.
					|| !base_class->SetItem(_T("__Class"), ((__int64)mCurrFileIndex << 32) | mCombinedLineNumber)
					|| !mUnresolvedClasses->SetItem(base_class_name, base_class)  )
					return ScriptError(ERR_OUTOFMEM);
			}
		}
	}

	// Validate the name even if this is a nested definition, for consistency.
	if (!Var::ValidateName(class_name, DISPLAY_NO_ERROR))
		return ScriptError(_T("Invalid class name."), class_name);

	class_object = NULL; // This initializes the entry in the mClassObject array.
	
	if (mClassObjectCount) // Nested class definition.
	{
		outer_class = mClassObject[mClassObjectCount - 1];
		if (outer_class->GetItem(token, class_name))
			// At this point it can only be an Object() created by a class definition.
			class_object = (Object *)token.object;
	}
	else // Top-level class definition.
	{
		*mClassName = '\0'; // Init.
		if (  !(class_var = FindOrAddVar(class_name))  )
			return FAIL;
		if (class_var->IsObject())
			// At this point it can only be an Object() created by a class definition.
			class_object = (Object *)class_var->Object();
		else
			// Force the variable to be super-global rather than passing this flag to
			// FindOrAddVar: a prior reference to this variable may have created it as
			// an ordinary global.
			class_var->Scope() = VAR_DECLARE_SUPER_GLOBAL;
	}
	
	if (_tcslen(mClassName) + _tcslen(class_name) + 1 >= _countof(mClassName)) // +1 for '.'
		return ScriptError(_T("Full class name is too long."));
	if (*mClassName)
		_tcscat(mClassName, _T("."));
	_tcscat(mClassName, class_name);

	// For now, it seems more useful to detect a duplicate as an error rather than as
	// a continuation of the previous definition.  Partial definitions might be allowed
	// in future, perhaps via something like "Class Foo continued".
	if (class_object)
		return ScriptError(_T("Duplicate class definition."), aBuf);

	token.symbol = SYM_STRING;
	token.marker = mClassName;
	
	if (mUnresolvedClasses)
	{
		ExprTokenType result_token, *param = &token;
		result_token.symbol = SYM_STRING;
		result_token.marker = _T("");
		result_token.mem_to_free = NULL;
		// Search for class and simultaneously remove it from the unresolved list:
		mUnresolvedClasses->_Remove(result_token, &param, 1); // result_token := mUnresolvedClasses.Remove(token)
		// If a field was found/removed, it can only be SYM_OBJECT.
		if (result_token.symbol == SYM_OBJECT)
		{
			// Use this object as the class.  At least one other object already refers to it as mBase.
			// At this point class_object["__Class"] contains the file index and line number, but it
			// will be overwritten below.
			class_object = (Object *)result_token.object;
		}
	}

	if (   !(class_object || (class_object = Object::Create()))
		|| !(class_object->SetItem(_T("__Class"), token))
		|| !(mClassObjectCount
				? outer_class->SetItem(class_name, class_object) // Assign to super_class[class_name].
				: class_var->Assign(class_object))   ) // Assign to global variable named %class_name%.
		return ScriptError(ERR_OUTOFMEM);

	class_object->SetBase(base_class); // May be NULL.

	++mClassObjectCount;
	return OK;
}


ResultType Script::DefineClassProperty(LPTSTR aBuf)
{
	LPTSTR name_end = find_identifier_end(aBuf);
	if (*name_end == '.')
		return ScriptError(ERR_INVALID_LINE_IN_CLASS_DEF, aBuf);

	LPTSTR param_start = omit_leading_whitespace(name_end);
	if (*param_start == '[')
	{
		LPTSTR param_end = aBuf + _tcslen(aBuf);
		if (param_end[-1] != ']')
			return ScriptError(ERR_MISSING_CLOSE_BRACKET, aBuf);
		*param_start = '(';
		param_end[-1] = ')';
	}
	else
		param_start = _T("()");
	
	// Save the property name and parameter list for later use with DefineFunc().
	mClassPropertyDef = tmalloc(_tcslen(aBuf) + 7); // +7 for ".Get()\0"
	if (!mClassPropertyDef)
		return ScriptError(ERR_OUTOFMEM);
	_stprintf(mClassPropertyDef, _T("%.*s.Get%s"), int(name_end - aBuf), aBuf, param_start);

	Object *class_object = mClassObject[mClassObjectCount - 1];
	*name_end = 0; // Terminate for aBuf use below.
	if (class_object->GetItem(ExprTokenType(), aBuf))
		return ScriptError(ERR_DUPLICATE_DECLARATION, aBuf);
	mClassProperty = new Property();
	if (!mClassProperty || !class_object->SetItem(aBuf, mClassProperty))
		return ScriptError(ERR_OUTOFMEM);
	return OK;
}


ResultType Script::DefineClassVars(LPTSTR aBuf, bool aStatic)
{
	Object *class_object = mClassObject[mClassObjectCount - 1];
	LPTSTR item, item_end;
	TCHAR orig_char, buf[LINE_SIZE];
	size_t buf_used = 0;
	ExprTokenType temp_token, empty_token, int_token;
	empty_token.symbol = SYM_STRING;
	empty_token.marker = _T("");
	int_token.symbol = SYM_INTEGER; // Value used to mark instance variables.
	int_token.value_int64 = 1;      //
					
	for (item = omit_leading_whitespace(aBuf); *item;) // FOR EACH COMMA-SEPARATED ITEM IN THE DECLARATION LIST.
	{
		item_end = find_identifier_end(item);
		if (item_end == item)
			return ScriptError(ERR_INVALID_CLASS_VAR, item);
		orig_char = *item_end;
		*item_end = '\0'; // Temporarily terminate.
		bool item_exists = class_object->GetItem(temp_token, item);
		if (orig_char == '.')
		{
			*item_end = orig_char; // Undo termination.
			// This is something like "object.key := 5", which is only valid if "object" was
			// previously declared (and will presumably be assigned an object at runtime).
			// Ensure that at least the root class var exists; any further validation would
			// be impossible since the object doesn't exist yet.
			if (!item_exists)
				return ScriptError(_T("Unknown class var."), item);
			for (TCHAR *cp; *item_end == '.'; item_end = cp)
			{
				for (cp = item_end + 1; IS_IDENTIFIER_CHAR(*cp); ++cp);
				if (cp == item_end + 1)
					// This '.' wasn't followed by a valid identifier.  Leave item_end
					// pointing at '.' and allow the switch() below to report the error.
					break;
			}
		}
		else
		{
			if (item_exists)
				return ScriptError(ERR_DUPLICATE_DECLARATION, item);
			// Assign class_object[item] := "" to mark it as a class variable
			// and allow duplicate declarations to be detected:
			if (!class_object->SetItem(item, aStatic ? empty_token : int_token))
				return ScriptError(ERR_OUTOFMEM);
			*item_end = orig_char; // Undo termination.
		}
		size_t name_length = item_end - item;
						
		// This section is very similar to the one in ParseAndAddLine() which deals with
		// variable declarations, so maybe maintain them together:

		item_end = omit_leading_whitespace(item_end); // Move up to the next comma, assignment-op, or '\0'.
		switch (*item_end)
		{
		case ',':  // No initializer is present for this variable, so move on to the next one.
			item = omit_leading_whitespace(item_end + 1); // Set "item" for use by the next iteration.
			continue; // No further processing needed below.
		case '\0': // No initializer is present for this variable, so move on to the next one.
			item = item_end; // Set "item" for use by the loop's condition.
			continue;
		case '=': // Supported for consistency with v1 syntax; to be removed in v2.
			++item_end; // Point to the character after the "=".
			break;
		case ':':
			if (item_end[1] == '=')
			{
				item_end += 2; // Point to the character after the ":=".
				break;
			}
			// Otherwise, fall through to below:
		default:
			return ScriptError(ERR_INVALID_CLASS_VAR, item);
		}
						
		// Since above didn't "continue", this declared variable also has an initializer.
		// Append the class name, ":=" and initializer to pending_buf, to be turned into
		// an expression below, and executed at script start-up.
		item_end = omit_leading_whitespace(item_end);
		LPTSTR right_side_of_operator = item_end; // Save for use below.

		item_end += FindNextDelimiter(item_end); // Find the next comma which is not part of the initializer (or find end of string).

		// Append "ClassNameOrThis.VarName := Initializer, " to the buffer.
		int chars_written = _sntprintf(buf + buf_used, _countof(buf) - buf_used, _T("%s.%.*s := %.*s, ")
			, aStatic ? mClassName : _T("this"), (int)name_length, item, (int)(item_end - right_side_of_operator), right_side_of_operator);
		if (chars_written < 0)
			return ScriptError(_T("Declaration too long.")); // Short message since should be rare.
		buf_used += chars_written;

		// Set "item" for use by the next iteration:
		item = (*item_end == ',') // i.e. it's not the terminator and thus not the final item in the list.
			? omit_leading_whitespace(item_end + 1)
			: item_end; // It's the terminator, so let the loop detect that to finish.
	}
	if (buf_used)
	{
		// Above wrote at least one initializer expression into buf.
		buf[buf_used -= 2] = '\0'; // Remove the final ", "

		// The following section temporarily replaces mLastLine in order to insert script lines
		// either at the end of the list of static initializers (separate from the main script)
		// or at the end of the __Init method belonging to this class.  Save the current values:
		Line *script_first_line = mFirstLine, *script_last_line = mLastLine;
		Line *block_end;
		Func *init_func = NULL;
		
		if (!aStatic)
		{
			ExprTokenType token;
			if (class_object->GetItem(token, _T("__Init")) && token.symbol == SYM_OBJECT
				&& (init_func = dynamic_cast<Func *>(token.object))) // This cast SHOULD always succeed; done for maintainability.
			{
				// __Init method already exists, so find the end of its body.
				for (block_end = init_func->mJumpToLine;
					 block_end->mActionType != ACT_BLOCK_END || !block_end->mAttribute;
					 block_end = block_end->mNextLine);
			}
			else
			{
				// Create an __Init method for this class.
				TCHAR def[] = _T("__Init()");
				if (!DefineFunc(def, NULL) || !AddLine(ACT_BLOCK_BEGIN)
					|| (class_object->Base() && !ParseAndAddLine(_T("base.__Init()"), ACT_EXPRESSION))) // Initialize base-class variables first. Relies on short-circuit evaluation.
					return FAIL;
				
				mLastLine->mLineNumber = 0; // Signal the debugger to skip this line while stepping in/over/out.
				init_func = g->CurrentFunc;
				init_func->mDefaultVarType = VAR_DECLARE_GLOBAL; // Allow global variables/class names in initializer expressions.
				
				if (!AddLine(ACT_BLOCK_END)) // This also resets g->CurrentFunc to NULL.
					return FAIL;
				block_end = mLastLine;
				block_end->mLineNumber = 0; // See above.
				
				// These must be updated as one or both have changed:
				script_first_line = mFirstLine;
				script_last_line = mLastLine;
			}
			g->CurrentFunc = init_func; // g->CurrentFunc should be NULL prior to this.
			mLastLine = block_end->mPrevLine; // i.e. insert before block_end.
			mLastLine->mNextLine = NULL; // For maintainability; AddLine() should overwrite it regardless.
			mCurrLine = NULL; // Fix for v1.1.09.02: Leaving this non-NULL at best causes error messages to show irrelevant vicinity lines, and at worst causes a crash because the linked list is in an inconsistent state.
		}

		mNoUpdateLabels = true;
		if (!ParseAndAddLine(buf))
			return FAIL; // Above already displayed the error.
		mNoUpdateLabels = false;
		
		if (aStatic)
		{
			mLastLine->mAttribute = (AttributeType)mLastLine->mActionType;
			mLastLine->mActionType = ACT_STATIC; // Mark this line for the preparser.
		}
		else
		{
			if (init_func->mJumpToLine == block_end) // This can be true only for the first initializer of a class with no base-class.
				init_func->mJumpToLine = mLastLine;
			// Rejoin the function's block-end (and any lines following it) to the main script.
			mLastLine->mNextLine = block_end;
			block_end->mPrevLine = mLastLine;
			// mFirstLine should be left as it is: if it was NULL, it now contains a pointer to our
			// __init function's block-begin, which is now the very first executable line in the script.
			g->CurrentFunc = NULL;
			// Restore mLastLine so that any subsequent script lines are added at the correct point.
			mLastLine = script_last_line;
		}
	}
	return OK;
}


Object *Script::FindClass(LPCTSTR aClassName, size_t aClassNameLength)
{
	if (!aClassNameLength)
		aClassNameLength = _tcslen(aClassName);
	if (!aClassNameLength || aClassNameLength > MAX_CLASS_NAME_LENGTH)
		return NULL;

	LPTSTR cp, key;
	ExprTokenType token;
	Object *base_object = NULL;
	TCHAR class_name[MAX_CLASS_NAME_LENGTH + 2]; // Extra +1 for '.' to simplify parsing.
	
	// Make temporary copy which we can modify.
	tmemcpy(class_name, aClassName, aClassNameLength);
	class_name[aClassNameLength] = '.'; // To simplify parsing.
	class_name[aClassNameLength + 1] = '\0';

	// Get base variable; e.g. "MyClass" in "MyClass.MySubClass".
	cp = _tcschr(class_name + 1, '.');
	Var *base_var = FindVar(class_name, cp - class_name, NULL, FINDVAR_GLOBAL);
	if (!base_var)
		return NULL;

	// Although at load time only the "Object" type can exist, dynamic_cast is used in case we're called at run-time:
	if (  !(base_var->IsObject() && (base_object = dynamic_cast<Object *>(base_var->Object())))  )
		return NULL;

	// Even if the loop below has no iterations, it initializes 'key' to the appropriate value:
	for (key = cp + 1; cp = _tcschr(key, '.'); key = cp + 1) // For each key in something like TypeVar.Key1.Key2.
	{
		if (cp == key)
			return NULL; // ScriptError(_T("Missing name."), cp);
		*cp = '\0'; // Terminate at the delimiting dot.
		if (!base_object->GetItem(token, key))
			return NULL;
		base_object = (Object *)token.object; // See comment about Object() above.
	}

	return base_object;
}


Object *Object::GetUnresolvedClass(LPTSTR &aName)
// This method is only valid for mUnresolvedClass.
{
	if (!mFieldCount)
		return NULL;
	aName = mFields[0].key.s;
	return (Object *)mFields[0].object;
}

ResultType Script::ResolveClasses()
{
	LPTSTR name;
	Object *base = mUnresolvedClasses->GetUnresolvedClass(name);
	if (!base)
		return OK;
	// There is at least one unresolved class.
	ExprTokenType token;
	if (base->GetItem(token, _T("__Class")))
	{
		// In this case (an object in the mUnresolvedClasses list), it is always an integer
		// containing the file index and line number:
		mCurrFileIndex = int(token.value_int64 >> 32);
		mCombinedLineNumber = LineNumberType(token.value_int64);
	}
	mCurrLine = NULL;
	return ScriptError(_T("Unknown class."), name);
}




#ifndef AUTOHOTKEYSC
struct FuncLibrary
{
	LPTSTR path;
	DWORD_PTR length;
};

Func *Script::FindFuncInLibrary(LPTSTR aFuncName, size_t aFuncNameLength, bool &aErrorWasShown, bool &aFileWasFound, bool aIsAutoInclude)
// Caller must ensure that aFuncName doesn't already exist as a defined function.
// If aFuncNameLength is 0, the entire length of aFuncName is used.
{
	aErrorWasShown = false; // Set default for this output parameter.
	aFileWasFound = false;

	int i;
	LPTSTR char_after_last_backslash, terminate_here;
	DWORD attr;

	#define FUNC_LIB_EXT EXT_AUTOHOTKEY
	#define FUNC_LIB_EXT_LENGTH (_countof(FUNC_LIB_EXT) - 1)
	#define FUNC_LOCAL_LIB _T("\\Lib\\") // Needs leading and trailing backslash.
	#define FUNC_LOCAL_LIB_LENGTH (_countof(FUNC_LOCAL_LIB) - 1)
	#define FUNC_USER_LIB _T("\\AutoHotkey\\Lib\\") // Needs leading and trailing backslash.
	#define FUNC_USER_LIB_LENGTH (_countof(FUNC_USER_LIB) - 1)
	#define FUNC_STD_LIB _T("Lib\\") // Needs trailing but not leading backslash.
	#define FUNC_STD_LIB_LENGTH (_countof(FUNC_STD_LIB) - 1)

	#define FUNC_LIB_COUNT 3
	static FuncLibrary sLib[FUNC_LIB_COUNT] = {0};

	if (!sLib[0].path) // Allocate & discover paths only upon first use because many scripts won't use anything from the library. This saves a bit of memory and performance.
	{
		for (i = 0; i < FUNC_LIB_COUNT; ++i)
			if (   !(sLib[i].path = (LPTSTR) SimpleHeap::Malloc(MAX_PATH * sizeof(TCHAR)))   ) // Need MAX_PATH for to allow room for appending each candidate file/function name.
				return NULL; // Due to rarity, simply pass the failure back to caller.

		FuncLibrary *this_lib;

		// DETERMINE PATH TO "LOCAL" LIBRARY:
		this_lib = sLib; // For convenience and maintainability.
		this_lib->length = BIV_ScriptDir(NULL, _T(""));
		if (this_lib->length < MAX_PATH-FUNC_LOCAL_LIB_LENGTH)
		{
			this_lib->length = BIV_ScriptDir(this_lib->path, _T(""));
			_tcscpy(this_lib->path + this_lib->length, FUNC_LOCAL_LIB);
			this_lib->length += FUNC_LOCAL_LIB_LENGTH;
		}
		else // Insufficient room to build the path name.
		{
			*this_lib->path = '\0'; // Mark this library as disabled.
			this_lib->length = 0;   //
		}

		// DETERMINE PATH TO "USER" LIBRARY:
		this_lib++; // For convenience and maintainability.
		this_lib->length = BIV_MyDocuments(this_lib->path, _T(""));
		if (this_lib->length < MAX_PATH-FUNC_USER_LIB_LENGTH)
		{
			_tcscpy(this_lib->path + this_lib->length, FUNC_USER_LIB);
			this_lib->length += FUNC_USER_LIB_LENGTH;
		}
		else // Insufficient room to build the path name.
		{
			*this_lib->path = '\0'; // Mark this library as disabled.
			this_lib->length = 0;   //
		}

		// DETERMINE PATH TO "STANDARD" LIBRARY:
		this_lib++; // For convenience and maintainability.
		GetModuleFileName(NULL, this_lib->path, MAX_PATH); // The full path to the currently-running AutoHotkey.exe.
		char_after_last_backslash = 1 + _tcsrchr(this_lib->path, '\\'); // Should always be found, so failure isn't checked.
		this_lib->length = (DWORD)(char_after_last_backslash - this_lib->path); // The length up to and including the last backslash.
		if (this_lib->length < MAX_PATH-FUNC_STD_LIB_LENGTH)
		{
			_tcscpy(this_lib->path + this_lib->length, FUNC_STD_LIB);
			this_lib->length += FUNC_STD_LIB_LENGTH;
		}
		else // Insufficient room to build the path name.
		{
			*this_lib->path = '\0'; // Mark this library as disabled.
			this_lib->length = 0;   //
		}

		for (i = 0; i < FUNC_LIB_COUNT; ++i)
		{
			attr = GetFileAttributes(sLib[i].path); // Seems to accept directories that have a trailing backslash, which is good because it simplifies the code.
			if (attr == 0xFFFFFFFF || !(attr & FILE_ATTRIBUTE_DIRECTORY)) // Directory doesn't exist or it's a file vs. directory. Relies on short-circuit boolean order.
			{
				*sLib[i].path = '\0'; // Mark this library as disabled.
				sLib[i].length = 0;   //
			}
		}
	}
	// Above must ensure that all sLib[].path elements are non-NULL (but they can be "" to indicate "no library").

	if (!aFuncNameLength) // Caller didn't specify, so use the entire string.
		aFuncNameLength = _tcslen(aFuncName);

	TCHAR *dest, *first_underscore, class_name_buf[MAX_VAR_NAME_LENGTH + 1];
	LPTSTR naked_filename = aFuncName;               // Set up for the first iteration.
	size_t naked_filename_length = aFuncNameLength; //

	for (int second_iteration = 0; second_iteration < 2; ++second_iteration)
	{
		for (i = 0; i < FUNC_LIB_COUNT; ++i)
		{
			if (!*sLib[i].path) // Library is marked disabled, so skip it.
				continue;

			if (sLib[i].length + naked_filename_length >= MAX_PATH-FUNC_LIB_EXT_LENGTH)
				continue; // Path too long to match in this library, but try others.
			dest = (LPTSTR) tmemcpy(sLib[i].path + sLib[i].length, naked_filename, naked_filename_length); // Append the filename to the library path.
			_tcscpy(dest + naked_filename_length, FUNC_LIB_EXT); // Append the file extension.

			attr = GetFileAttributes(sLib[i].path); // Testing confirms that GetFileAttributes() doesn't support wildcards; which is good because we want filenames containing question marks to be "not found" rather than being treated as a match-pattern.
			if (attr == 0xFFFFFFFF || (attr & FILE_ATTRIBUTE_DIRECTORY)) // File doesn't exist or it's a directory. Relies on short-circuit boolean order.
				continue;

			aFileWasFound = true; // Indicate success for #include <lib>, which doesn't necessarily expect a function to be found.

			// Since above didn't "continue", a file exists whose name matches that of the requested function.
			// Before loading/including that file, set the working directory to its folder so that if it uses
			// #Include, it will be able to use more convenient/intuitive relative paths.  This is similar to
			// the "#Include DirName" feature.
			// Call SetWorkingDir() vs. SetCurrentDirectory() so that it succeeds even for a root drive like
			// C: that lacks a backslash (see SetWorkingDir() for details).
			terminate_here = sLib[i].path + sLib[i].length - 1; // The trailing backslash in the full-path-name to this library.
			*terminate_here = '\0'; // Temporarily terminate it for use with SetWorkingDir().
			SetWorkingDir(sLib[i].path); // See similar section in the #Include directive.
			*terminate_here = '\\'; // Undo the termination.

			if (mIncludeLibraryFunctionsThenExit && aIsAutoInclude)
			{
				// For each auto-included library-file, write out two #Include lines:
				// 1) Use #Include in its "change working directory" mode so that any explicit #include directives
				//    or FileInstalls inside the library file itself will work consistently and properly.
				// 2) Use #IncludeAgain (to improve performance since no dupe-checking is needed) to include
				//    the library file itself.
				// We don't directly append library files onto the main script here because:
				// 1) ahk2exe needs to be able to see and act upon FileInstall and #Include lines (i.e. library files
				//    might contain #Include even though it's rare).
				// 2) #IncludeAgain and #Include directives that bring in fragments rather than entire functions or
				//    subroutines wouldn't work properly if we resolved such includes in AutoHotkey.exe because they
				//    wouldn't be properly interleaved/asynchronous, but instead brought out of their library file
				//    and deposited separately/synchronously into the temp-include file by some new logic at the
				//    AutoHotkey.exe's code for the #Include directive.
				// 3) ahk2exe prefers to omit comments from included files to minimize size of compiled scripts.
				mIncludeLibraryFunctionsThenExit->Format(_T("#Include %-0.*s\n#IncludeAgain %s\n")
					, sLib[i].length, sLib[i].path, sLib[i].path);
				// Now continue on normally so that our caller can continue looking for syntax errors.
			}
			
			// Fix for v1.1.06.00: If the file contains any lib #includes, it must be loaded AFTER the
			// above writes sLib[i].path to the iLib file, otherwise the wrong filename could be written.
			if (!LoadIncludedFile(sLib[i].path, false, false)) // Fix for v1.0.47.05: Pass false for allow-dupe because otherwise, it's possible for a stdlib file to attempt to include itself (especially via the LibNamePrefix_ method) and thus give a misleading "duplicate function" vs. "func does not exist" error message.  Obsolete: For performance, pass true for allow-dupe so that it doesn't have to check for a duplicate file (seems too rare to worry about duplicates since by definition, the function doesn't yet exist so it's file shouldn't yet be included).
			{
				aErrorWasShown = true; // Above has just displayed its error (e.g. syntax error in a line, failed to open the include file, etc).  So override the default set earlier.
				return NULL;
			}

			// Now that a matching filename has been found, it seems best to stop searching here even if that
			// file doesn't actually contain the requested function.  This helps library authors catch bugs/typos.
			return FindFunc(aFuncName, aFuncNameLength);
		} // for() each library directory.

		// Now that the first iteration is done, set up for the second one that searches by class/prefix.
		// Notes about ambiguity and naming collisions:
		// By the time it gets to the prefix/class search, it's almost given up.  Even if it wrongly finds a
		// match in a filename that isn't really a class, it seems inconsequential because at worst it will
		// still not find the function and will then say "call to nonexistent function".  In addition, the
		// ability to customize which libraries are searched is planned.  This would allow a publicly
		// distributed script to turn off all libraries except stdlib.
		if (   !(first_underscore = _tcschr(aFuncName, '_'))   ) // No second iteration needed.
			break; // All loops are done because second iteration is the last possible attempt.
		naked_filename_length = first_underscore - aFuncName;
		if (naked_filename_length >= _countof(class_name_buf)) // Class name too long (probably impossible currently).
			break; // All loops are done because second iteration is the last possible attempt.
		naked_filename = class_name_buf; // Point it to a buffer for use below.
		tmemcpy(naked_filename, aFuncName, naked_filename_length);
		naked_filename[naked_filename_length] = '\0';
	} // 2-iteration for().

	// Since above didn't return, no match found in any library.
	return NULL;
}
#endif



Func *Script::FindFunc(LPCTSTR aFuncName, size_t aFuncNameLength, int *apInsertPos) // L27: Added apInsertPos for binary-search.
// Returns the Function whose name matches aFuncName (which caller has ensured isn't NULL).
// If it doesn't exist, NULL is returned.
{
	if (!aFuncNameLength) // Caller didn't specify, so use the entire string.
		aFuncNameLength = _tcslen(aFuncName);

	if (apInsertPos) // L27: Set default for maintainability.
		*apInsertPos = -1;

	// For the below, no error is reported because callers don't want that.  Instead, simply return
	// NULL to indicate that names that are illegal or too long are not found.  If the caller later
	// tries to add the function, it will get an error then:
	if (aFuncNameLength > MAX_VAR_NAME_LENGTH)
		return NULL;

	// The following copy is made because it allows the name searching to use _tcsicmp() instead of
	// strlicmp(), which close to doubles the performance.  The copy includes only the first aVarNameLength
	// characters from aVarName:
	TCHAR func_name[MAX_VAR_NAME_LENGTH + 1];
	tcslcpy(func_name, aFuncName, aFuncNameLength + 1);  // +1 to convert length to size.

	Func *pfunc;
	
	// Using a binary searchable array vs a linked list speeds up dynamic function calls, on average.
	int left, right, mid, result;
	for (left = 0, right = mFuncCount - 1; left <= right;)
	{
		mid = (left + right) / 2;
		result = _tcsicmp(func_name, mFunc[mid]->mName); // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales; 3) performance.
		if (result > 0)
			left = mid + 1;
		else if (result < 0)
			right = mid - 1;
		else // Match found.
			return mFunc[mid];
	}
	if (apInsertPos)
		*apInsertPos = left;

	// Since above didn't return, there is no match.  See if it's a built-in function that hasn't yet
	// been added to the function list.

	// Set defaults to be possibly overridden below:
	int min_params = 1;
	int max_params = 1;
	BuiltInFunctionType bif;
	LPTSTR suffix = func_name + 3;

	if (!_tcsnicmp(func_name, _T("LV_"), 3)) // As a built-in function, LV_* can only be a ListView function.
	{
		suffix = func_name + 3;
		if (!_tcsicmp(suffix, _T("GetNext")))
		{
			bif = BIF_LV_GetNextOrCount;
			min_params = 0;
			max_params = 2;
		}
		else if (!_tcsicmp(suffix, _T("GetCount")))
		{
			bif = BIF_LV_GetNextOrCount;
			min_params = 0; // But leave max at its default of 1.
		}
		else if (!_tcsicmp(suffix, _T("GetText")))
		{
			bif = BIF_LV_GetText;
			min_params = 2;
			max_params = 3;
		}
		else if (!_tcsicmp(suffix, _T("Add")))
		{
			bif = BIF_LV_AddInsertModify;
			min_params = 0; // 0 params means append a blank row.
			max_params = 10000; // An arbitrarily high limit that will never realistically be reached.
		}
		else if (!_tcsicmp(suffix, _T("Insert")))
		{
			bif = BIF_LV_AddInsertModify;
			// Leave min_params at 1.  Passing only 1 param to it means "insert a blank row".
			max_params = 10000; // An arbitrarily high limit that will never realistically be reached.
		}
		else if (!_tcsicmp(suffix, _T("Modify")))
		{
			bif = BIF_LV_AddInsertModify; // Although it shares the same function with "Insert", it can still have its own min/max params.
			// Leave min_params at 1 so that it can be called like LV_Modify(row, , col1, col2).
			max_params = 10000; // An arbitrarily high limit that will never realistically be reached.
		}
		else if (!_tcsicmp(suffix, _T("Delete")))
		{
			bif = BIF_LV_Delete;
			min_params = 0; // Leave max at its default of 1.
		}
		else if (!_tcsicmp(suffix, _T("InsertCol")))
		{
			bif = BIF_LV_InsertModifyDeleteCol;
			// Leave min_params at 1 because inserting a blank column ahead of the first column
			// does not seem useful enough to sacrifice the no-parameter mode, which might have
			// potential future uses.
			max_params = 3;
		}
		else if (!_tcsicmp(suffix, _T("ModifyCol")))
		{
			bif = BIF_LV_InsertModifyDeleteCol;
			min_params = 0;
			max_params = 3;
		}
		else if (!_tcsicmp(suffix, _T("DeleteCol")))
			bif = BIF_LV_InsertModifyDeleteCol; // Leave min/max set to 1.
		else if (!_tcsicmp(suffix, _T("SetImageList")))
		{
			bif = BIF_LV_SetImageList;
			max_params = 2; // Leave min at 1.
		}
		else
			return NULL;
	}
	else if (!_tcsnicmp(func_name, _T("TV_"), 3)) // As a built-in function, TV_* can only be a TreeView function.
	{
		suffix = func_name + 3;
		if (!_tcsicmp(suffix, _T("Add")))
		{
			bif = BIF_TV_AddModifyDelete;
			max_params = 3; // Leave min at its default of 1.
		}
		else if (!_tcsicmp(suffix, _T("Modify")))
		{
			bif = BIF_TV_AddModifyDelete;
			max_params = 3; // One-parameter mode is "select specified item".
		}
		else if (!_tcsicmp(suffix, _T("Delete")))
		{
			bif = BIF_TV_AddModifyDelete;
			min_params = 0;
		}
		else if (!_tcsicmp(suffix, _T("GetParent")) || !_tcsicmp(suffix, _T("GetChild")) || !_tcsicmp(suffix, _T("GetPrev")))
			bif = BIF_TV_GetRelatedItem;
		else if (!_tcsicmp(suffix, _T("GetCount")) || !_tcsicmp(suffix, _T("GetSelection")))
		{
			bif = BIF_TV_GetRelatedItem;
			min_params = 0;
			max_params = 0;
		}
		else if (!_tcsicmp(suffix, _T("GetNext"))) // Unlike "Prev", Next also supports 0 or 2 parameters.
		{
			bif = BIF_TV_GetRelatedItem;
			min_params = 0;
			max_params = 2;
		}
		else if (!_tcsicmp(suffix, _T("Get")) || !_tcsicmp(suffix, _T("GetText")))
		{
			bif = BIF_TV_Get;
			min_params = 2;
			max_params = 2;
		}
		else if (!_tcsicmp(suffix, _T("SetImageList")))
		{
			bif = BIF_TV_SetImageList;
			max_params = 2; // Leave min at 1.
		}
		else
			return NULL;
	}
	else if (!_tcsnicmp(func_name, _T("IL_"), 3)) // It's an ImageList function.
	{
		suffix = func_name + 3;
		if (!_tcsicmp(suffix, _T("Create")))
		{
			bif = BIF_IL_Create;
			min_params = 0;
			max_params = 3;
		}
		else if (!_tcsicmp(suffix, _T("Destroy")))
		{
			bif = BIF_IL_Destroy; // Leave Min/Max set to 1.
		}
		else if (!_tcsicmp(suffix, _T("Add")))
		{
			bif = BIF_IL_Add;
			min_params = 2;
			max_params = 4;
		}
		else
			return NULL;
	}
	else if (!_tcsicmp(func_name, _T("SB_SetText")))
	{
		bif = BIF_StatusBar;
		max_params = 3; // Leave min_params at its default of 1.
	}
	else if (!_tcsicmp(func_name, _T("SB_SetParts")))
	{
		bif = BIF_StatusBar;
		min_params = 0;
		max_params = 255; // 255 params allows for up to 256 parts, which is SB's max.
	}
	else if (!_tcsicmp(func_name, _T("SB_SetIcon")))
	{
		bif = BIF_StatusBar;
		max_params = 3; // Leave min_params at its default of 1.
	}
	else if (!_tcsicmp(func_name, _T("StrLen")))
		bif = BIF_StrLen;
	else if (!_tcsicmp(func_name, _T("SubStr")))
	{
		bif = BIF_SubStr;
		min_params = 2;
		max_params = 3;
	}
	else if (!_tcsicmp(func_name, _T("Trim")) || !_tcsicmp(func_name, _T("LTrim")) || !_tcsicmp(func_name, _T("RTrim"))) // L31
	{
		bif = BIF_Trim;
		min_params = 1;
		max_params = 2;
	}
	else if (!_tcsicmp(func_name, _T("InStr")))
	{
		bif = BIF_InStr;
		min_params = 2;
		max_params = 5;
	}
	else if (!_tcsicmp(func_name, _T("RegExMatch")))
	{
		bif = BIF_RegEx;
		min_params = 2;
		max_params = 4;
	}
	else if (!_tcsicmp(func_name, _T("RegExReplace")))
	{
		bif = BIF_RegEx;
		min_params = 2;
		max_params = 6;
	}
	else if (!_tcsicmp(func_name, _T("StrReplace")))
	{
		bif = BIF_StrReplace;
		min_params = 2;
		max_params = 5;
	}
	else if (!_tcsicmp(func_name, _T("StrSplit")))
	{
		bif = BIF_StrSplit;
		min_params = 1;
		max_params = 4;
	}
	else if (!_tcsnicmp(func_name, _T("GetKey"), 6))
	{
		suffix = func_name + 6;
		if (!_tcsicmp(suffix, _T("State")))
		{
			bif = BIF_GetKeyState;
			max_params = 2;
		}
		else if (!_tcsicmp(suffix, _T("Name")) || !_tcsicmp(suffix, _T("VK")) || !_tcsicmp(suffix, _T("SC")))
			bif = BIF_GetKeyName;
		else
			return NULL;
	}
	else if (!_tcsicmp(func_name, _T("Ord")) || !_tcsicmp(func_name, _T("Asc")))
		bif = BIF_Ord;
	else if (!_tcsicmp(func_name, _T("Chr")))
		bif = BIF_Chr;
	else if (!_tcsicmp(func_name, _T("Format")))
	{
		bif = BIF_Format;
		max_params = 10000; // An arbitrarily high limit that will never realistically be reached.
	}
	else if (!_tcsicmp(func_name, _T("StrGet")))
	{
		bif = BIF_StrGetPut;
		max_params = 3;
	}
	else if (!_tcsicmp(func_name, _T("StrPut")))
	{
		bif = BIF_StrGetPut;
		max_params = 4;
	}
	else if (!_tcsicmp(func_name, _T("NumGet")))
	{
		bif = BIF_NumGet;
		max_params = 3;
	}
	else if (!_tcsicmp(func_name, _T("NumPut")))
	{
		bif = BIF_NumPut;
		min_params = 2;
		max_params = 4;
	}
	else if (!_tcsicmp(func_name, _T("IsLabel")))
		bif = BIF_IsLabel;
	else if (!_tcsicmp(func_name, _T("Func")))
		bif = BIF_Func;
	else if (!_tcsicmp(func_name, _T("IsFunc")))
		bif = BIF_IsFunc;
	else if (!_tcsicmp(func_name, _T("IsByRef")))
		bif = BIF_IsByRef;
#ifdef ENABLE_DLLCALL
	else if (!_tcsicmp(func_name, _T("DllCall")))
	{
		bif = BIF_DllCall;
		max_params = 10000; // An arbitrarily high limit that will never realistically be reached.
	}
#endif
	else if (!_tcsicmp(func_name, _T("VarSetCapacity")))
	{
		bif = BIF_VarSetCapacity;
		max_params = 3;
	}
	else if (!_tcsicmp(func_name, _T("FileExist")))
		bif = BIF_FileExist;
	else if (!_tcsicmp(func_name, _T("WinExist")) || !_tcsicmp(func_name, _T("WinActive")))
	{
		bif = BIF_WinExistActive;
		min_params = 0;
		max_params = 4;
	}
	else if (!_tcsicmp(func_name, _T("Round")))
	{
		bif = BIF_Round;
		max_params = 2;
	}
	else if (!_tcsicmp(func_name, _T("Floor")) || !_tcsicmp(func_name, _T("Ceil")))
		bif = BIF_FloorCeil;
	else if (!_tcsicmp(func_name, _T("Mod")))
	{
		bif = BIF_Mod;
		min_params = 2;
		max_params = 2;
	}
	else if (!_tcsicmp(func_name, _T("Min")) || !_tcsicmp(func_name, _T("Max")))
	{
		bif = BIF_MinMax;
		max_params = 10000; // An arbitrarily high limit that will never realistically be reached.
	}
	else if (!_tcsicmp(func_name, _T("Abs")))
		bif = BIF_Abs;
	else if (!_tcsicmp(func_name, _T("Sin")))
		bif = BIF_Sin;
	else if (!_tcsicmp(func_name, _T("Cos")))
		bif = BIF_Cos;
	else if (!_tcsicmp(func_name, _T("Tan")))
		bif = BIF_Tan;
	else if (!_tcsicmp(func_name, _T("ASin")) || !_tcsicmp(func_name, _T("ACos")))
		bif = BIF_ASinACos;
	else if (!_tcsicmp(func_name, _T("ATan")))
		bif = BIF_ATan;
	else if (!_tcsicmp(func_name, _T("Exp")))
		bif = BIF_Exp;
	else if (!_tcsicmp(func_name, _T("Sqrt")) || !_tcsicmp(func_name, _T("Log")) || !_tcsicmp(func_name, _T("Ln")))
		bif = BIF_SqrtLogLn;
	else if (!_tcsicmp(func_name, _T("OnMessage")))
	{
		bif = BIF_OnMessage;
		max_params = 3;  // Leave min at 1.
		// By design, scripts that use OnMessage are persistent by default.  Doing this here
		// also allows WinMain() to later detect whether this script should become #SingleInstance.
		// Note: Don't directly change g_AllowOnlyOneInstance here in case the remainder of the
		// script-loading process comes across any explicit uses of #SingleInstance, which would
		// override the default set here.
		g_persistent = true;
	}
	else if (!_tcsicmp(func_name, _T("OnExit"))
		|| !_tcsicmp(func_name, _T("OnClipboardChange"))
		|| !_tcsicmp(func_name, _T("OnError")))
	{
		bif = BIF_On;
		max_params = 2;
	}
#ifdef ENABLE_REGISTERCALLBACK
	else if (!_tcsicmp(func_name, _T("RegisterCallback")))
	{
		bif = BIF_RegisterCallback;
		max_params = 4; // Leave min_params at 1.
	}
#endif
	else if (!_tcsicmp(func_name, _T("IsObject"))) // L31
	{
		bif = BIF_IsObject;
		max_params = 10000; // Leave min_params at 1.
	}
	else if (!_tcsnicmp(func_name, _T("Obj"), 3)) // L31: See script_object.cpp for details.
	{
		suffix = func_name + 3;

		if (!_tcsicmp(suffix, _T("ect"))) // i.e. "Object"
		{
			bif = BIF_ObjCreate;
			min_params = 0;
			max_params = 10000;
		}
#define BIF_OBJ_CASE(aCaseSuffix, aMinParams, aMaxParams) \
		else if (!_tcsicmp(suffix, _T(#aCaseSuffix))) \
		{ \
			bif = BIF_Obj##aCaseSuffix; \
			min_params = (1 + aMinParams); \
			max_params = (1 + aMaxParams); \
		}
		// All of these functions require the "object" parameter,
		// but it is excluded from the counts below for clarity:
		BIF_OBJ_CASE(Insert,		1, 10000) // [key,] value [, value2, ...]
		BIF_OBJ_CASE(InsertAt,		2, 10000) // index, value [, value2, ...]
		BIF_OBJ_CASE(Push,			1, 10000)
		BIF_OBJ_CASE(Delete, 		1, 2) // min_key [, max_key]
		BIF_OBJ_CASE(Remove, 		0, 2) // [min_key, max_key]
		BIF_OBJ_CASE(RemoveAt,      1, 2) // position [, count]
		BIF_OBJ_CASE(Pop,			0, 0)
		BIF_OBJ_CASE(Count, 		0, 0)
		BIF_OBJ_CASE(Length, 		0, 0)
		BIF_OBJ_CASE(MinIndex, 		0, 0)
		BIF_OBJ_CASE(MaxIndex, 		0, 0)
		BIF_OBJ_CASE(HasKey,		1, 1) // key
		BIF_OBJ_CASE(GetCapacity,	0, 1) // [key]
		BIF_OBJ_CASE(SetCapacity,	1, 2) // [key,] new_capacity
		BIF_OBJ_CASE(GetAddress,	1, 1) // key
		BIF_OBJ_CASE(NewEnum,		0, 0)
		BIF_OBJ_CASE(Clone,			0, 0)
		BIF_OBJ_CASE(BindMethod,	1, 10000) // obj, method [, param...]
#undef BIF_OBJ_CASE
		else if (!_tcsicmp(suffix, _T("AddRef")) || !_tcsicmp(suffix, _T("Release")))
			bif = BIF_ObjAddRefRelease;
		else if (!_tcsicmp(suffix, _T("RawSet")))
		{
			bif = BIF_ObjRaw;
			min_params = 3;
			max_params = 3;
		}
		else if (!_tcsicmp(suffix, _T("RawGet")))
		{
			bif = BIF_ObjRaw;
			min_params = 2;
			max_params = 2;
		}
		else if (!_tcsicmp(suffix, _T("GetBase")))
			bif = BIF_ObjBase;
		else if (!_tcsicmp(suffix, _T("SetBase")))
		{
			bif = BIF_ObjBase;
			min_params = 2;
			max_params = 2;
		}
		else return NULL;
	}
	else if (!_tcsicmp(func_name, _T("Array")))
	{
		bif = BIF_ObjArray;
		min_params = 0;
		max_params = 10000;
	}
	else if (!_tcsicmp(func_name, _T("FileOpen")))
	{
		bif = BIF_FileOpen;
		min_params = 2;
		max_params = 3;
	}
	else if (!_tcsnicmp(func_name, _T("ComObj"), 6))
	{
		suffix = func_name + 6;
		if	(!_tcsicmp(suffix, _T("Create")))
		{
			bif = BIF_ComObjCreate;
			max_params = 2;
		}
		else if	(!_tcsicmp(suffix, _T("Get")))
			bif = BIF_ComObjGet;
		else if	(!_tcsicmp(suffix, _T("Connect")))
		{
			bif = BIF_ComObjConnect;
			max_params = 2;
		}
		else if (!_tcsicmp(suffix, _T("Error")))
		{
			bif = BIF_ComObjError;
			min_params = 0;
		}
		else if (!_tcsicmp(suffix, _T("Type")))
		{
			bif = BIF_ComObjTypeOrValue;
			max_params = 2;
		}
		else if (!_tcsicmp(suffix, _T("Value")))
		{
			bif = BIF_ComObjTypeOrValue;
		}
		else if (!_tcsicmp(suffix, _T("Flags")))
		{
			bif = BIF_ComObjFlags;
			max_params = 3;
		}
		else if (!_tcsicmp(suffix, _T("Array")))
		{
			bif = BIF_ComObjArray;
			min_params = 2;
			max_params = 9; // up to 8 dimensions
		}
		else if (!_tcsicmp(suffix, _T("Query")))
		{
			bif = BIF_ComObjQuery;
			min_params = 2;
			max_params = 3;
		}
		else
		{
			// Fixed in v1.1.14.04: Avoid trying to make a Func with an invalid name. This fixes IsFunc("ComObj(") throwing an exception.
			if (!Var::ValidateName(func_name, DISPLAY_NO_ERROR))
				return NULL;
			bif = BIF_ComObjActive;
			min_params = 0;
			max_params = 3;
		}
	}
	else if (!_tcsicmp(func_name, _T("Exception")))
	{
		bif = BIF_Exception;
		max_params = 3;
	}
	else if (!_tcsicmp(func_name, _T("MenuGetHandle")) || !_tcsicmp(func_name, _T("MenuGetName")))
		bif = BIF_MenuGet;
	else if (!_tcsicmp(func_name, _T("LoadPicture")))
	{
		bif = BIF_LoadPicture;
		max_params = 3;
	}
	else if (!_tcsicmp(func_name, _T("Hotstring")))
	{
		bif = BIF_Hotstring;
		max_params = 3;
	}
	else
		return NULL; // Maint: There may be other lines above that also return NULL.

	// Since above didn't return, this is a built-in function that hasn't yet been added to the list.
	// Add it now:
	if (   !(pfunc = AddFunc(func_name, aFuncNameLength, true, left))   ) // L27: left contains the position within mFunc to insert the function.  Cannot use *apInsertPos as caller may have omitted it or passed NULL.
		return NULL;

	pfunc->mBIF = bif;
	pfunc->mMinParams = min_params;
	pfunc->mParamCount = max_params;

	return pfunc;
}



Func *Script::AddFunc(LPCTSTR aFuncName, size_t aFuncNameLength, bool aIsBuiltIn, int aInsertPos, Object *aClassObject)
// This function should probably not be called by anyone except FindOrAddFunc, which has already done
// the dupe-checking.
// Returns the address of the new function or NULL on failure.
// The caller must already have verified that this isn't a duplicate function.
{
	if (!aFuncNameLength) // Caller didn't specify, so use the entire string.
		aFuncNameLength = _tcslen(aFuncName);

	if (aFuncNameLength > MAX_VAR_NAME_LENGTH)
	{
		ScriptError(_T("Function name too long."), aFuncName);
		return NULL;
	}

	// Make a temporary copy that includes only the first aFuncNameLength characters from aFuncName:
	TCHAR func_name[MAX_VAR_NAME_LENGTH + 1];
	tcslcpy(func_name, aFuncName, aFuncNameLength + 1);  // See explanation above.  +1 to convert length to size.

	// In the future, it might be best to add another check here to disallow function names that consist
	// entirely of numbers.  However, this hasn't been done yet because:
	// 1) Not sure if there will ever be a good enough reason.
	// 2) Even if it's done in the far future, it won't break many scripts (pure-numeric functions should be very rare).
	// 3) Those scripts that are broken are not broken in a bad way because the pre-parser will generate a
	//    load-time error, which is easy to fix (unlike runtime errors, which require that part of the script
	//    to actually execute).
	if (!aClassObject && !Var::ValidateName(func_name, DISPLAY_FUNC_ERROR))  // Variable and function names are both validated the same way.
		// Above already displayed error for us.  This can happen at loadtime or runtime (e.g. StringSplit).
		return NULL;

	// Allocate some dynamic memory to pass to the constructor:
	LPTSTR new_name = SimpleHeap::Malloc(func_name, aFuncNameLength);
	if (!new_name)
		// It already displayed the error for us.
		return NULL;

	Func *the_new_func = new Func(new_name, aIsBuiltIn);
	if (!the_new_func)
	{
		ScriptError(ERR_OUTOFMEM);
		return NULL;
	}

	if (aClassObject)
	{
		LPTSTR key = _tcsrchr(new_name, '.');
		if (!key)
		{
			ScriptError(_T("Invalid method name."), new_name); // Shouldn't ever happen.
			return NULL;
		}
		if (mClassProperty)
		{
			if (toupper(key[1]) == 'G')
				mClassProperty->mGet = the_new_func;
			else
				mClassProperty->mSet = the_new_func;
		}
		else
			if (!aClassObject->SetItem(key + 1, the_new_func))
			{
				ScriptError(ERR_OUTOFMEM);
				return NULL;
			}
		aClassObject->AddRef(); // In case the script clears the class var.
		the_new_func->mClass = aClassObject;
		// Also add it to the script's list of functions, to support #Warn LocalSameAsGlobal
		// and automatic cleanup of objects in static vars on program exit.
	}
	
	if (mFuncCount == mFuncCountMax)
	{
		// Allocate or expand function list.
		int alloc_count = mFuncCountMax ? mFuncCountMax * 2 : 100;

		Func **temp = (Func **)realloc(mFunc, alloc_count * sizeof(Func *)); // If passed NULL, realloc() will do a malloc().
		if (!temp)
		{
			ScriptError(ERR_OUTOFMEM);
			return NULL;
		}
		mFunc = temp;
		mFuncCountMax = alloc_count;
	}

	if (aInsertPos != mFuncCount) // Need to make room at the indicated position for this variable.
		memmove(mFunc + aInsertPos + 1, mFunc + aInsertPos, (mFuncCount - aInsertPos) * sizeof(Func *));
	//else both are zero or the item is being inserted at the end of the list, so it's easy.
	mFunc[aInsertPos] = the_new_func;
	++mFuncCount;

	return the_new_func;
}



size_t Line::ArgIndexLength(int aArgIndex)
// This function is similar to ArgToInt(), so maintain them together.
// "ArgLength" is the arg's fully resolved, dereferenced length during runtime.
// Callers must call this only at times when sArgDeref and sArgVar are defined/meaningful.
// Caller must ensure that aArgIndex is 0 or greater.
// ArgLength() was added in v1.0.44.14 to help its callers improve performance by avoiding
// costly calls to _tcslen() (which is especially beneficial for huge strings).
{
#ifdef _DEBUG
	if (aArgIndex < 0)
	{
		LineError(_T("DEBUG: BAD"), WARN);
		aArgIndex = 0;  // But let it continue.
	}
#endif
	if (aArgIndex >= mArgc) // Arg doesn't exist, so don't try accessing sArgVar (unlike sArgDeref, it wouldn't be valid to do so).
		return 0; // i.e. treat it as the empty string.
	// The length is not known and must be calculated in the following situations:
	// - The arg consists of more than just a single isolated variable name (not possible if the arg is
	//   ARG_TYPE_INPUT_VAR).
	// - The arg is a built-in variable, in which case the length isn't known, so it must be derived from
	//   the string copied into sArgDeref[] by an earlier stage.
	// - The arg is a normal variable but it's VAR_ATTRIB_BINARY_CLIP. In such cases, our callers do not
	//   recognize/support binary-clipboard as binary and want the apparent length of the string returned
	//   (i.e. _tcslen(), which takes into account the position of the first binary zero wherever it may be).
	if (sArgVar[aArgIndex])
	{
		Var &var = *sArgVar[aArgIndex]; // For performance and convenience.
		if (   var.Type() == VAR_NORMAL  // This and below ordered for short-circuit performance based on types of input expected from caller.
			&& !(g_act[mActionType].MaxParamsAu2WithHighBit & 0x80) // Although the ones that have the highbit set are hereby omitted from the fast method, the nature of almost all of the highbit commands is such that their performance won't be measurably affected. See ArgMustBeDereferenced() for more info.
			&& (g_NoEnv || var.HasContents()) // v1.0.46.02: Recognize environment variables (when g_NoEnv==FALSE) by falling through to _tcslen() for them.
			&& &var != g_ErrorLevel   ) // Mostly for maintainability because the following situation is very rare: If it's g_ErrorLevel, use the deref version instead because if g_ErrorLevel is an input variable in the caller's command, and the caller changes ErrorLevel (such as to set a default) prior to calling this function, the changed/new ErrorLevel will be used rather than its original value (which is usually undesirable).
			//&& !var.IsBinaryClip())  // This check isn't necessary because the line below handles it.
			return var.LengthIgnoreBinaryClip(); // Do it the fast way (unless it's binary clipboard, in which case this call will internally call _tcslen()).
	}
	// Otherwise, length isn't known due to no variable, a built-in variable, or an environment variable.
	// So do it the slow way.
	return _tcslen(sArgDeref[aArgIndex]);
}



__int64 Line::ArgIndexToInt64(int aArgIndex)
// This function is similar to ArgIndexLength(), so maintain them together.
// Callers must call this only at times when sArgDeref and sArgVar are defined/meaningful.
// Caller must ensure that aArgIndex is 0 or greater.
{
#ifdef _DEBUG
	if (aArgIndex < 0)
	{
		LineError(_T("DEBUG: BAD"), WARN);
		aArgIndex = 0;  // But let it continue.
	}
#endif
	if (aArgIndex >= mArgc) // See ArgIndexLength() for comments.
		return 0; // i.e. treat it as ATOI64("").
	// SEE THIS POSITION IN ArgIndexLength() FOR IMPORTANT COMMENTS ABOUT THE BELOW.
	if (sArgVar[aArgIndex])
	{
		Var &var = *sArgVar[aArgIndex];
		if (   var.Type() == VAR_NORMAL  // See ArgIndexLength() for comments about this line and below.
			&& !(g_act[mActionType].MaxParamsAu2WithHighBit & 0x80)
			&& (g_NoEnv || var.HasContents())
			&& &var != g_ErrorLevel
			&& !var.IsBinaryClip()   )
			return var.ToInt64(FALSE);
	}
	// Otherwise:
	return ATOI64(sArgDeref[aArgIndex]); // See ArgIndexLength() for comments.
}



double Line::ArgIndexToDouble(int aArgIndex)
// This function is similar to ArgIndexLength(), so maintain them together.
// Callers must call this only at times when sArgDeref and sArgVar are defined/meaningful.
// Caller must ensure that aArgIndex is 0 or greater.
{
#ifdef _DEBUG
	if (aArgIndex < 0)
	{
		LineError(_T("DEBUG: BAD"), WARN);
		aArgIndex = 0;  // But let it continue.
	}
#endif
	if (aArgIndex >= mArgc) // See ArgIndexLength() for comments.
		return 0.0; // i.e. treat it as ATOF("").
	// SEE THIS POSITION IN ARGLENGTH() FOR IMPORTANT COMMENTS ABOUT THE BELOW.
	if (sArgVar[aArgIndex])
	{
		Var &var = *sArgVar[aArgIndex];
		if (   var.Type() == VAR_NORMAL  // See ArgIndexLength() for comments about this line and below.
			&& !(g_act[mActionType].MaxParamsAu2WithHighBit & 0x80)
			&& (g_NoEnv || var.HasContents())
			&& &var != g_ErrorLevel
			&& !var.IsBinaryClip()   )
			return var.ToDouble(FALSE);
	}
	// Otherwise:
	return ATOF(sArgDeref[aArgIndex]); // See ArgIndexLength() for comments.
}



Var *Line::ResolveVarOfArg(int aArgIndex, bool aCreateIfNecessary)
// Returns NULL on failure.  Caller has ensured that none of this arg's derefs are function-calls.
// Args that are input or output variables are normally resolved at load-time, so that
// they contain a pointer to their Var object.  This is done for performance.  However,
// in order to support dynamically resolved variables names like AutoIt2 (e.g. arrays),
// we have to do some extra work here at runtime.
// Callers specify false for aCreateIfNecessary whenever the contents of the variable
// they're trying to find is unimportant.  For example, dynamically built input variables,
// such as "StringLen, length, array%i%", do not need to be created if they weren't
// previously assigned to (i.e. they weren't previously used as an output variable).
// In the above example, the array element would never be created here.  But if the output
// variable were dynamic, our call would have told us to create it.
{
	// The requested ARG isn't even present, so it can't have a variable.  Currently, this should
	// never happen because the loading procedure ensures that input/output args are not marked
	// as variables if they are blank (and our caller should check for this and not call in that case):
	if (aArgIndex >= mArgc)
		return NULL;
	ArgStruct &this_arg = mArg[aArgIndex]; // For performance and convenience.

	// Since this function isn't inline (since it's called so frequently), there isn't that much more
	// overhead to doing this check, even though it shouldn't be needed since it's the caller's
	// responsibility:
	if (this_arg.type == ARG_TYPE_NORMAL) // Arg isn't an input or output variable.
		return NULL;
	if (!*this_arg.text) // The arg's variable is not one that needs to be dynamically resolved.
		return VAR(this_arg); // Return the var's address that was already determined at load-time.
	// The above might return NULL in the case where the arg is optional (i.e. the command allows
	// the var name to be omitted).  But in that case, the caller should either never have called this
	// function or should check for NULL upon return.  UPDATE: This actually never happens, see
	// comment above the "if (aArgIndex >= mArgc)" line.

	// Static to correspond to the static empty_var further below.  It needs the memory area
	// to support resolving dynamic environment variables.  In the following example,
	// the result will be blank unless the first line is present (without this fix here):
	//null = %SystemRoot%  ; bogus line as a required workaround in versions prior to v1.0.16
	//thing = SystemRoot
	//StringTrimLeft, output, %thing%, 0
	//msgbox %output%

	static TCHAR sVarName[MAX_VAR_NAME_LENGTH + 1];  // Will hold the dynamically built name.

	// At this point, we know the requested arg is a variable that must be dynamically resolved.
	// This section is similar to that in ExpandArg(), so they should be maintained together:
	LPTSTR pText = this_arg.text; // Start at the beginning of this arg's text.
	size_t var_name_length = 0;

	if (this_arg.deref) // There's at least one deref.
	{
		// Caller has ensured that none of these derefs are function calls (i.e. deref->is_function is alway false).
		for (DerefType *deref = this_arg.deref  // Start off by looking for the first deref.
			; deref->marker; ++deref)  // A deref with a NULL marker terminates the list.
		{
			// FOR EACH DEREF IN AN ARG (if we're here, there's at least one):
			// Copy the chars that occur prior to deref->marker into the buffer:
			for (; pText < deref->marker && var_name_length < MAX_VAR_NAME_LENGTH; sVarName[var_name_length++] = *pText++);
			if (var_name_length >= MAX_VAR_NAME_LENGTH && pText < deref->marker) // The variable name would be too long!
			{
				// This type of error is just a warning because this function isn't set up to cause a true
				// failure.  This is because the use of dynamically named variables is rare, and only for
				// people who should know what they're doing.  In any case, when the caller of this
				// function called it to resolve an output variable, it will see the the result is
				// NULL and terminate the current subroutine.
				#define DYNAMIC_TOO_LONG _T("This dynamically built variable name is too long.") \
					_T("  If this variable was not intended to be dynamic, remove the % symbols from it.")
				LineError(DYNAMIC_TOO_LONG, FAIL, this_arg.text);
				return NULL;
			}
			// Now copy the contents of the dereferenced var.  For all cases, aBuf has already
			// been verified to be large enough, assuming the value hasn't changed between the
			// time we were called and the time the caller calculated the space needed.
			if (deref->var->Get() > (VarSizeType)(MAX_VAR_NAME_LENGTH - var_name_length)) // The variable name would be too long!
			{
				LineError(DYNAMIC_TOO_LONG, FAIL, this_arg.text);
				return NULL;
			}
			var_name_length += deref->var->Get(sVarName + var_name_length);
			// Finally, jump over the dereference text. Note that in the case of an expression, there might not
			// be any percent signs within the text of the dereference, e.g. x + y, not %x% + %y%.
			pText += deref->length;
		}
	}

	// Copy any chars that occur after the final deref into the buffer:
	for (; *pText && var_name_length < MAX_VAR_NAME_LENGTH; sVarName[var_name_length++] = *pText++);
	if (var_name_length >= MAX_VAR_NAME_LENGTH && *pText) // The variable name would be too long!
	{
		LineError(DYNAMIC_TOO_LONG, FAIL, this_arg.text);
		return NULL;
	}
	
	if (!var_name_length)
	{
		LineError(_T("This dynamic variable is blank. If this variable was not intended to be dynamic,")
			_T(" remove the % symbols from it."), FAIL, this_arg.text);
		return NULL;
	}

	// Terminate the buffer, even if nothing was written into it:
	sVarName[var_name_length] = '\0';

	static Var empty_var(sVarName, (void *)VAR_NORMAL, false); // Must use sVarName here.  See comment above for why.

	Var *found_var;
	if (!aCreateIfNecessary)
	{
		// Now we've dynamically build the variable name.  It's possible that the name is illegal,
		// so check that (the name is automatically checked by FindOrAddVar(), so we only need to
		// check it if we're not calling that):
		if (!Var::ValidateName(sVarName))
			return NULL; // Above already displayed error for us.
		if (found_var = g_script.FindVar(sVarName, var_name_length)) // Assign.
			return found_var;
		// At this point, this is either a non-existent variable or a reserved/built-in variable
		// that was never statically referenced in the script (only dynamically), e.g. A_IPAddress%A_Index%
		if (!Script::GetBuiltInVar(sVarName))
			// If not found: for performance reasons, don't create it because caller just wants an empty variable.
			return &empty_var;
		//else it's the clipboard or some other built-in variable, so continue onward so that the
		// variable gets created in the variable list, which is necessary to allow it to be properly
		// dereferenced, e.g. in a script consisting of only the following:
		// Loop, 4
		//     StringTrimRight, IP, A_IPAddress%A_Index%, 0
	}
	// Otherwise, aCreateIfNecessary is true or we want to create this variable unconditionally for the
	// reason described above.
	if (   !(found_var = g_script.FindOrAddVar(sVarName, var_name_length))   )
		return NULL;  // Above will already have displayed the error.
	if (this_arg.type == ARG_TYPE_OUTPUT_VAR && VAR_IS_READONLY(*found_var))
	{
		LineError(ERR_VAR_IS_READONLY, FAIL, sVarName);
		return NULL;  // Don't return the var, preventing the caller from assigning to it.
	}
	else
		return found_var;
}



Var *Script::FindOrAddVar(LPTSTR aVarName, size_t aVarNameLength, int aScope)
// Caller has ensured that aVarName isn't NULL.
// Returns the Var whose name matches aVarName.  If it doesn't exist, it is created.
{
	if (!*aVarName)
		return NULL;
	int insert_pos;
	bool is_local; // Used to detect which type of var should be added in case the result of the below is NULL.
	Var *var;
	if (var = FindVar(aVarName, aVarNameLength, &insert_pos, aScope, &is_local))
		return var;
	// Otherwise, no match found, so create a new var.  This will return NULL if there was a problem,
	// in which case AddVar() will already have displayed the error:
	return AddVar(aVarName, aVarNameLength, insert_pos
		, (aScope & ~(VAR_LOCAL | VAR_GLOBAL)) | (is_local ? VAR_LOCAL : VAR_GLOBAL)); // When aScope == FINDVAR_DEFAULT, it contains both the "local" and "global" bits.  This ensures only the appropriate bit is set.
}



Var *Script::FindVar(LPTSTR aVarName, size_t aVarNameLength, int *apInsertPos, int aScope
	, bool *apIsLocal)
// Caller has ensured that aVarName isn't NULL.  It must also ignore the contents of apInsertPos when
// a match (non-NULL value) is returned.
// Returns the Var whose name matches aVarName.  If it doesn't exist, NULL is returned.
// If caller provided a non-NULL apInsertPos, it will be given a the array index that a newly
// inserted item should have to keep the list in sorted order (which also allows the ListVars command
// to display the variables in alphabetical order).
{
	if (!*aVarName)
		return NULL;
	if (!aVarNameLength) // Caller didn't specify, so use the entire string.
		aVarNameLength = _tcslen(aVarName);

	// For the below, no error is reported because callers don't want that.  Instead, simply return
	// NULL to indicate that names that are illegal or too long are not found.  When the caller later
	// tries to add the variable, it will get an error then:
	if (aVarNameLength > MAX_VAR_NAME_LENGTH)
		return NULL;

	// The following copy is made because it allows the various searches below to use _tcsicmp() instead of
	// strlicmp(), which close to doubles their performance.  The copy includes only the first aVarNameLength
	// characters from aVarName:
	TCHAR var_name[MAX_VAR_NAME_LENGTH + 1];
	tcslcpy(var_name, aVarName, aVarNameLength + 1);  // +1 to convert length to size.

	global_struct &g = *::g; // Reduces code size and may improve performance.
	bool search_local = (aScope & VAR_LOCAL) && g.CurrentFunc;

	// Above has ensured that g.CurrentFunc!=NULL whenever search_local==true.

	// Init for binary search loop:
	int left, right, mid, result;  // left/right must be ints to allow them to go negative and detect underflow.
	Var **var;  // An array of pointers-to-var.
	if (search_local)
	{
		var = g.CurrentFunc->mVar;
		right = g.CurrentFunc->mVarCount - 1;
	}
	else
	{
		var = mVar;
		right = mVarCount - 1;
	}

	// Binary search:
	for (left = 0; left <= right;) // "right" was already initialized above.
	{
		mid = (left + right) / 2;
		result = _tcsicmp(var_name, var[mid]->mName); // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales; 3) performance.
		if (result > 0)
			left = mid + 1;
		else if (result < 0)
			right = mid - 1;
		else // Match found.
			return var[mid];
	}

	// Since above didn't return, no match was found in the main list, so search the lazy list if there
	// is one.  If there's no lazy list, the value of "left" established above will be used as the
	// insertion point further below:
	if (search_local)
	{
		var = g.CurrentFunc->mLazyVar;
		right = g.CurrentFunc->mLazyVarCount - 1;
	}
	else
	{
		var = mLazyVar;
		right = mLazyVarCount - 1;
	}

	if (var) // There is a lazy list to search (and even if the list is empty, left must be reset to 0 below).
	{
		// Binary search:
		for (left = 0; left <= right;)  // "right" was already initialized above.
		{
			mid = (left + right) / 2;
			result = _tcsicmp(var_name, var[mid]->mName); // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales; 3) performance.
			if (result > 0)
				left = mid + 1;
			else if (result < 0)
				right = mid - 1;
			else // Match found.
				return var[mid];
		}
	}

	// Since above didn't return, no match was found and "left" always contains the position where aVarName
	// should be inserted to keep the list sorted.  The item is always inserted into the lazy list unless
	// there is no lazy list.
	// Set the output parameter, if present:
	if (apInsertPos) // Caller wants this value even if we'll be resorting to searching the global list below.
		*apInsertPos = left; // This is the index a newly inserted item should have to keep alphabetical order.
	
	if (apIsLocal) // Its purpose is to inform caller of type it would have been in case we don't find a match.
		*apIsLocal = search_local;

	// Since no match was found, if this is a local fall back to searching the list of globals at runtime
	// if the caller didn't insist on a particular type:
	if (search_local && aScope == FINDVAR_DEFAULT)
	{
		// In this case, callers want to fall back to globals when a local wasn't found.  However,
		// they want the insertion (if our caller will be doing one) to insert according to the
		// current assume-mode.  Therefore, if the mode is assume-global, pass the apIsLocal
		// and apInsertPos variables to FindVar() so that it will update them to be global.
		if (g.CurrentFunc->mDefaultVarType == VAR_DECLARE_GLOBAL)
			return FindVar(aVarName, aVarNameLength, apInsertPos, FINDVAR_GLOBAL, apIsLocal);
		// v1: Each *dynamic* variable reference may resolve to a global if one exists.
		bool force_local = g.CurrentFunc->mDefaultVarType & VAR_FORCE_LOCAL;
		if (mIsReadyToExecute && !force_local)
			return FindVar(aVarName, aVarNameLength, NULL, FINDVAR_GLOBAL);
		// Otherwise, caller only wants globals which are declared in *this* function:
		for (int i = 0; i < g.CurrentFunc->mGlobalVarCount; ++i)
			if (!_tcsicmp(var_name, g.CurrentFunc->mGlobalVar[i]->mName)) // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales; 3) performance.
				return g.CurrentFunc->mGlobalVar[i];
		if (force_local)
			return NULL;
		// As a last resort, check for a super-global:
		Var *gvar = FindVar(aVarName, aVarNameLength, NULL, FINDVAR_GLOBAL, NULL);
		if (gvar && gvar->IsSuperGlobal())
			return gvar;
	}
	// Otherwise, since above didn't return:
	return NULL; // No match.
}



Var *Script::AddVar(LPTSTR aVarName, size_t aVarNameLength, int aInsertPos, int aScope)
// Returns the address of the new variable or NULL on failure.
// Caller must ensure that g->CurrentFunc!=NULL whenever aIsLocal!=0.
// Caller must ensure that aVarName isn't NULL and that this isn't a duplicate variable name.
// In addition, it has provided aInsertPos, which is the insertion point so that the list stays sorted.
// Finally, aIsLocal has been provided to indicate which list, global or local, should receive this
// new variable, as well as the type of local variable.  (See the declaration of VAR_LOCAL etc.)
{
	if (!*aVarName) // Should never happen, so just silently indicate failure.
		return NULL;
	if (!aVarNameLength) // Caller didn't specify, so use the entire string.
		aVarNameLength = _tcslen(aVarName);

	if (aVarNameLength > MAX_VAR_NAME_LENGTH)
	{
		ScriptError(_T("Variable name too long."), aVarName);
		return NULL;
	}

	// Make a temporary copy that includes only the first aVarNameLength characters from aVarName:
	TCHAR var_name[MAX_VAR_NAME_LENGTH + 1];
	tcslcpy(var_name, aVarName, aVarNameLength + 1);  // See explanation above.  +1 to convert length to size.

	if (!Var::ValidateName(var_name))
		// Above already displayed error for us.  This can happen at loadtime or runtime (e.g. StringSplit).
		return NULL;

	bool aIsLocal = (aScope & VAR_LOCAL);

	// Not necessary or desirable to add built-in variables to a function's list of locals.  Always keep
	// built-in vars in the global list for efficiency and to keep them out of ListVars.  Note that another
	// section at loadtime displays an error for any attempt to explicitly declare built-in variables as
	// either global or local.
	VarEntry *builtin = GetBuiltInVar(var_name);
	if (aIsLocal && (builtin || !_tcsicmp(var_name, _T("ErrorLevel")))) // Attempt to create built-in variable as local.
	{
		if (  !(aScope & VAR_LOCAL_FUNCPARAM)  ) // It's not a UDF's parameter, so fall back to the global built-in variable of this name rather than displaying an error.
			return FindOrAddVar(var_name, aVarNameLength, FINDVAR_GLOBAL); // Force find-or-create of global.
		else // (aIsLocal & VAR_LOCAL_FUNCPARAM), which means "this is a local variable and a function's parameter".
		{
			ScriptError(_T("Illegal parameter name."), aVarName); // Short message since so rare.
			return NULL;
		}
	}

	// Allocate some dynamic memory to pass to the constructor:
	LPTSTR new_name = SimpleHeap::Malloc(var_name, aVarNameLength);
	if (!new_name)
		// It already displayed the error for us.
		return NULL;

	// Below specifically tests for VAR_LOCAL and excludes other non-zero values/flags:
	//   VAR_LOCAL_FUNCPARAM should not be made static.
	//   VAR_LOCAL_STATIC is already static.
	//   VAR_DECLARED indicates mDefaultVarType is irrelevant.
	if (aScope == VAR_LOCAL && (g->CurrentFunc->mDefaultVarType & VAR_LOCAL_STATIC))
		// v1.0.48: Lexikos: Current function is assume-static, so set static attribute.
		aScope |= VAR_LOCAL_STATIC;

	Var *the_new_var = new Var(new_name, builtin ? builtin->type : (void *)VAR_NORMAL, aScope);
	if (the_new_var == NULL)
	{
		ScriptError(ERR_OUTOFMEM);
		return NULL;
	}

	// If there's a lazy var list, aInsertPos provided by the caller is for it, so this new variable
	// always gets inserted into that list because there's always room for one more (because the
	// previously added variable would have purged it if it had reached capacity).
	Var **lazy_var = aIsLocal ? g->CurrentFunc->mLazyVar : mLazyVar;
	int &lazy_var_count = aIsLocal ? g->CurrentFunc->mLazyVarCount : mLazyVarCount; // Used further below too.
	if (lazy_var)
	{
		if (aInsertPos != lazy_var_count) // Need to make room at the indicated position for this variable.
			memmove(lazy_var + aInsertPos + 1, lazy_var + aInsertPos, (lazy_var_count - aInsertPos) * sizeof(Var *));
		//else both are zero or the item is being inserted at the end of the list, so it's easy.
		lazy_var[aInsertPos] = the_new_var;
		++lazy_var_count;
		// In a testing creating between 200,000 and 400,000 variables, using a size of 1000 vs. 500 improves
		// the speed by 17%, but when you subtract out the binary search time (leaving only the insert time),
		// the increase is more like 34%.  But there is a diminishing return after that: Going to 2000 only
		// gains 20%, and to 4000 only gains an addition 10%.  Therefore, to conserve memory in functions that
		// have so many variables that the lazy list is used, a good trade-off seems to be 2000 (8 KB of memory)
		// per function that needs it.
		#define MAX_LAZY_VARS 2000 // Don't make this larger than 90000 without altering the incremental increase of alloc_count further below.
		if (lazy_var_count < MAX_LAZY_VARS) // The lazy list hasn't yet reached capacity, so no need to merge it into the main list.
			return the_new_var;
	}

	// Since above didn't return, either there is no lazy list or the lazy list is full and needs to be
	// merged into the main list.

	// Create references to whichever variable list (local or global) is being acted upon.  These
	// references simplify the code:
	Var **&var = aIsLocal ? g->CurrentFunc->mVar : mVar; // This needs to be a ref. too in case it needs to be realloc'd.
	int &var_count = aIsLocal ? g->CurrentFunc->mVarCount : mVarCount;
	int &var_count_max = aIsLocal ? g->CurrentFunc->mVarCountMax : mVarCountMax;
	int alloc_count;

	// Since the above would have returned if the lazy list is present but not yet full, if the left side
	// of the OR below is false, it also means that lazy_var is NULL.  Thus lazy_var==NULL is implicit for the
	// right side of the OR:
	if ((lazy_var && var_count + MAX_LAZY_VARS > var_count_max) || var_count == var_count_max)
	{
		// Increase by orders of magnitude each time because realloc() is probably an expensive operation
		// in terms of hurting performance.  So here, a little bit of memory is sacrificed to improve
		// the expected level of performance for scripts that use hundreds of thousands of variables.
		if (!var_count_max)
			alloc_count = aIsLocal ? 100 : 1000;  // 100 conserves memory since every function needs such a block, and most functions have much fewer than 100 local variables.
		else if (var_count_max < 1000)
			alloc_count = 1000;
		else if (var_count_max < 9999) // Making this 9999 vs. 10000 allows an exact/whole number of lazy_var blocks to fit into main indices between 10000 and 99999
			alloc_count = 9999;
		else if (var_count_max < 100000)
		{
			alloc_count = 100000;
			// This is also the threshold beyond which the lazy list is used to accelerate performance.
			// Create the permanent lazy list:
			Var **&lazy_var = aIsLocal ? g->CurrentFunc->mLazyVar : mLazyVar;
			if (   !(lazy_var = (Var **)malloc(MAX_LAZY_VARS * sizeof(Var *)))   )
			{
				ScriptError(ERR_OUTOFMEM);
				return NULL;
			}
		}
		else if (var_count_max < 1000000)
			alloc_count = 1000000;
		else
			alloc_count = var_count_max + 1000000;  // i.e. continue to increase by 4MB (1M*4) each time.

		Var **temp = (Var **)realloc(var, alloc_count * sizeof(Var *)); // If passed NULL, realloc() will do a malloc().
		if (!temp)
		{
			ScriptError(ERR_OUTOFMEM);
			return NULL;
		}
		var = temp;
		var_count_max = alloc_count;
	}

	if (!lazy_var)
	{
		if (aInsertPos != var_count) // Need to make room at the indicated position for this variable.
			memmove(var + aInsertPos + 1, var + aInsertPos, (var_count - aInsertPos) * sizeof(Var *));
		//else both are zero or the item is being inserted at the end of the list, so it's easy.
		var[aInsertPos] = the_new_var;
		++var_count;
		return the_new_var;
	}
	//else the variable was already inserted into the lazy list, so the above is not done.

	// Since above didn't return, the lazy list is not only present, but full because otherwise it
	// would have returned higher above.

	// Since the lazy list is now at its max capacity, merge it into the main list (if the
	// main list was at capacity, this section relies upon the fact that the above already
	// increased its capacity by an amount far larger than the number of items contained
	// in the lazy list).

	// LAZY LIST: Although it's not nearly as good as hashing (which might be implemented in the future,
	// though it would be no small undertaking since it affects so many design aspects, both load-time
	// and runtime for scripts), this method of accelerating insertions into a binary search array is
	// enormously beneficial because it improves the scalability of binary-search by two orders
	// of magnitude (from about 100,000 variables to at least 5M).  Credit for the idea goes to Lazlo.
	// DETAILS:
	// The fact that this merge operation is so much faster than total work required
	// to insert each one into the main list is the whole reason for having the lazy
	// list.  In other words, the large memmove() that would otherwise be required
	// to insert each new variable into the main list is completely avoided.  Large memmove()s
	// become dramatically more costly than small ones because apparently they can't fit into
	// the CPU cache, so the operation would take hundreds or even thousands of times longer
	// depending on the speed difference between main memory and CPU cache.  But above and
	// beyond the CPU cache issue, the lazy sorting method results in vastly less memory
	// being moved than would have been required without it, so even if the CPU doesn't have
	// a cache, the lazy list method vastly increases performance for scripts that have more
	// than 100,000 variables, allowing at least 5 million variables to be created without a
	// dramatic reduction in performance.

	LPTSTR target_name;
	Var **insert_pos, **insert_pos_prev;
	int i, left, right, mid;

	// Append any items from the lazy list to the main list that are alphabetically greater than
	// the last item in the main list.  Above has already ensured that the main list is large enough
	// to accept all items in the lazy list.
	for (i = lazy_var_count - 1, target_name = var[var_count - 1]->mName
		; i > -1 && _tcsicmp(target_name, lazy_var[i]->mName) < 0
		; --i);
	// Above is a self-contained loop.
	// Now do a separate loop to append (in the *correct* order) anything found above.
	for (int j = i + 1; j < lazy_var_count; ++j) // Might have zero iterations.
		var[var_count++] = lazy_var[j];
	lazy_var_count = i + 1; // The number of items that remain after moving out those that qualified.

	// This will have zero iterations if the above already moved them all:
	for (insert_pos = var + var_count, i = lazy_var_count - 1; i > -1; --i)
	{
		// Modified binary search that relies on the fact that caller has ensured a match will never
		// be found in the main list for each item in the lazy list:
		for (target_name = lazy_var[i]->mName, left = 0, right = (int)(insert_pos - var - 1); left <= right;)
		{
			mid = (left + right) / 2;
			if (_tcsicmp(target_name, var[mid]->mName) > 0) // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales; 3) performance.
				left = mid + 1;
			else // it must be < 0 because caller has ensured it can't be equal (i.e. that there will be no match)
				right = mid - 1;
		}
		// Now "left" contains the insertion point and is known to be less than var_count due to a previous
		// set of loops above.  Make a gap there large enough to hold all items because that allows a
		// smaller total amount of memory to be moved by shifting the gap to the left in the main list,
		// gradually filling it as we go:
		insert_pos_prev = insert_pos;  // "prev" is the now the position of the beginning of the gap, but the gap is about to be shifted left by moving memory right.
		insert_pos = var + left; // This is where it *would* be inserted if we weren't doing the accelerated merge.
		memmove(insert_pos + i + 1, insert_pos, (insert_pos_prev - insert_pos) * sizeof(Var *));
		var[left + i] = lazy_var[i]; // Now insert this item at the far right side of the gap just created.
	}
	var_count += lazy_var_count;
	lazy_var_count = 0;  // Indicate that the lazy var list is now empty.

	return the_new_var;
}



VarEntry *Script::GetBuiltInVar(LPTSTR aVarName)
{
	VarEntry *biv;
	int count;
	// This array approach saves about 9KB on code size over the old approach
	// of a series of if's and _tcscmp calls, and performs about the same.
	// Two arrays are used so that common dynamic vars (without "A_" prefix)
	// don't require as long a search, and so that "A_" can be omitted from
	// each var name in the array (to reduce code size).
	if ((aVarName[0] == 'A' || aVarName[0] == 'a') && aVarName[1] == '_')
	{
		aVarName += 2;
		biv = g_BIV_A;
		count = _countof(g_BIV_A);
	}
	else
	{
		biv = g_BIV;
		count = _countof(g_BIV);
	}
	// Using binary search vs. linear search performs a bit better (notably for
	// rare/contrived cases like A_x%index%) and doesn't affect code size much.
	int left, right, mid, result;
	for (left = 0, right = count - 1; left <= right;)
	{
		mid = (left + right) / 2;
		result = _tcsicmp(aVarName, biv[mid].name);
		if (result > 0)
			left = mid + 1;
		else if (result < 0)
			right = mid - 1;
		else // Match found.
			return &biv[mid];
	}
	// Since above didn't return:
	return NULL;
}



WinGroup *Script::FindGroup(LPTSTR aGroupName, bool aCreateIfNotFound)
// Caller must ensure that aGroupName isn't NULL.  But if it's the empty string, NULL is returned.
// Returns the Group whose name matches aGroupName.  If it doesn't exist, it is created if aCreateIfNotFound==true.
// Thread-safety: This function is thread-safe (except when called with aCreateIfNotFound==true) even when
// the main thread happens to be calling AddGroup() and changing the linked list while it's being traversed here
// by the hook thread.  However, any subsequent changes to this function or AddGroup() must be carefully reviewed.
{
	if (!*aGroupName)
	{
		if (aCreateIfNotFound)
			// An error message must be shown in this case since or caller is about to
			// exit the current script thread (and we don't want it to happen silently).
			ScriptError(_T("Blank group name."));
		return NULL;
	}
	for (WinGroup *group = mFirstGroup; group != NULL; group = group->mNextGroup)
		if (!_tcsicmp(group->mName, aGroupName)) // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales; 3) performance.
			return group; // Match found.
	// Otherwise, no match found, so create a new group.
	if (!aCreateIfNotFound || AddGroup(aGroupName) != OK)
		return NULL;
	return mLastGroup;
}



ResultType Script::AddGroup(LPTSTR aGroupName)
// Returns OK or FAIL.
// The caller must already have verified that this isn't a duplicate group.
// This function is not thread-safe because it adds an entry to the quasi-global list of window groups.
// In addition, if this function is being called by one thread while another thread is calling FindGroup(),
// the thread-safety notes in FindGroup() apply.
{
	size_t aGroupName_length = _tcslen(aGroupName);
	if (aGroupName_length > MAX_VAR_NAME_LENGTH)
		return ScriptError(_T("Group name too long."), aGroupName);
	if (!Var::ValidateName(aGroupName, DISPLAY_NO_ERROR)) // Seems best to use same validation as var names.
		return ScriptError(_T("Illegal group name."), aGroupName);

	LPTSTR new_name = SimpleHeap::Malloc(aGroupName, aGroupName_length);
	if (!new_name)
		return FAIL;  // It already displayed the error for us.

	// The precise method by which the follows steps are done should be thread-safe even if
	// some other thread calls FindGroup() in the middle of the operation.  But any changes
	// must be carefully reviewed:
	WinGroup *the_new_group = new WinGroup(new_name);
	if (the_new_group == NULL)
		return ScriptError(ERR_OUTOFMEM);
	if (mFirstGroup == NULL)
		mFirstGroup = the_new_group;
	else
		mLastGroup->mNextGroup = the_new_group;
	// This must be done after the above:
	mLastGroup = the_new_group;
	return OK;
}



ResultType Script::PreparseExpressions(Line *aStartingLine)
{
	int i;
	DerefType *deref;
	for (Line *line = aStartingLine; line; line = line->mNextLine)
	{
		// Check if any of each arg's derefs are function calls.  If so, do some validation and
		// preprocessing to set things up for better runtime performance:
		for (i = 0; i < line->mArgc; ++i) // For each arg.
		{
			ArgStruct &this_arg = line->mArg[i]; // For performance and convenience.
			// Exclude the derefs of output and input vars from consideration, since they can't
			// be function calls:
			if (!this_arg.is_expression) // For now, only expressions are capable of calling functions. If ever change this, might want to add a check here for this_arg.type != ARG_TYPE_NORMAL (for performance).
				continue;
			if (this_arg.deref) // No derefs are present because the expression contains neither variables nor function calls.
			{
			for (deref = this_arg.deref; deref->marker; ++deref) // For each deref.
			{
				if (!deref->is_function)
					continue;
				if (   !(deref->func = FindFunc(deref->marker, deref->length))   )
				{
#ifndef AUTOHOTKEYSC
					bool error_was_shown, file_was_found;
					if (   !(deref->func = FindFuncInLibrary(deref->marker, deref->length, error_was_shown, file_was_found, true))   )
					{
						// When above already displayed the proximate cause of the error, it's usually
						// undesirable to show the cascade effects of that error in a second dialog:
						return error_was_shown ? FAIL : line->LineError(ERR_NONEXISTENT_FUNCTION, FAIL, deref->marker);
					}
#else
					return line->LineError(ERR_NONEXISTENT_FUNCTION, FAIL, deref->marker);
#endif
				}
				// L31: Parameter counting and validation was previously done in this section,
				//		but is now handled by ExpressionToPostfix.
			} // for each deref of this arg
			} // if (this_arg.deref)
			if (!line->ExpressionToPostfix(this_arg)) // At this stage, this_arg.is_expression is known to be true. Doing this here, after the script has been loaded, might improve the compactness/adjacent-ness of the compiled expressions in memory, which might improve performance due to CPU caching.
				return FAIL; // The function above already displayed the error msg.
		} // for each arg of this line
	} // for each line
	return OK;
}


ResultType Script::PreparseStaticLines(Line *aStartingLine)
// Combining this and PreparseExpressions() into one loop/function currently increases
// code size enough to affect the final EXE size, contrary to expectation.
{
	// Remove static var initializers and #if expression lines from the main line list so
	// that they won't be executed or interfere with PreparseBlocks().  This used to be done
	// at an earlier stage, but that required multiple PreparseBlocks() calls to account for
	// lines added by lib auto-includes.  Thus, it's now done after PreparseExpressions().
	for (Line *next_line, *prev_line, *line = aStartingLine; line; line = next_line)
	{
		next_line = line->mNextLine; // Save these since may be overwritten below.
		prev_line = line->mPrevLine; //

		switch (line->mActionType)
		{
		case ACT_STATIC:
			// Override mActionType so ACT_STATIC doesn't have to be handled at runtime:
			line->mActionType = (ActionTypeType)(UINT_PTR)line->mAttribute;
			// Add this line to the list of static initializers, which will be inserted above
			// the auto-execute section later.  The main line list will be corrected below.
			line->mPrevLine = mLastStaticLine;
			if (mLastStaticLine)
				mLastStaticLine->mNextLine = line;
			else
				mFirstStaticLine = line;
			mLastStaticLine = line;
			break;
		case ACT_HOTKEY_IF:
			// It's already been added to the hot-expr list, so just remove it from the script (below).
			// Override mActionType so ACT_HOTKEY_IF doesn't have to be handled by EvaluateCondition():
			line->mActionType = ACT_IFEXPR;
			break;
		default:
			continue;
		}
		// Since above didn't "continue", remove this line from the main script.
		if (prev_line)
			prev_line->mNextLine = next_line;
		else // Must have been the first line; set the new first line.
			mFirstLine = next_line;
		if (next_line)
			next_line->mPrevLine = prev_line;
		else
			mLastLine = prev_line;
	}
	return OK;
}



Line *Script::PreparseBlocks(Line *aStartingLine, ExecUntilMode aMode, Line *aParentLine, const AttributeType aLoopType)
// Will return NULL to the top-level caller if there's an error, or if
// mLastLine is NULL (i.e. the script is empty).
{
	Line *line_temp;

	for (Line *line = aStartingLine; line != NULL;)
	{
		// All lines in our recursion layer are assigned to the line or block that the caller specified:
		line->mParentLine = aParentLine; // Can be NULL.

		if (ACT_IS_IF(line->mActionType) // IF but not ELSE, which is handled by its IF, or as an error below.
			|| ACT_IS_LOOP(line->mActionType) // LOOP, WHILE or FOR.
			|| line->mActionType == ACT_TRY) // CATCH and FINALLY are excluded so they're caught as errors below.
		{
			if (line->mActionType == ACT_FOR)
			{
				ASSERT(line->mArgc == 3);
				// Now that this FOR's expression has been pre-parsed, exclude it from mArgc so that ExpandArgs()
				// won't evaluate it -- PerformLoopFor() needs to call ExpandExpression() directly in order to
				// receive the object reference which is the result of the expression.
				line->mArgc--;
			}

			line_temp = line->mNextLine;  // line_temp is now this IF's or LOOP's or TRY's action-line.
			// The following is commented out because all scripts end in ACT_EXIT:
			//if (line_temp == NULL) // This is an orphan IF/LOOP (has no action-line) at the end of the script.
			//	return line->PreparseError(_T("Q")); // Placeholder. Formerly "This if-statement or loop has no action."

			// Checks such as the following are omitted because any such errors are detected automatically
			// by recursion into this function; i.e. if this IF/LOOP/TRY's action is an ELSE, it will be
			// detected as having no associated IF.
			//#define IS_BAD_ACTION_LINE_TYPE(act) ((act) == ACT_ELSE || (act) == ACT_BLOCK_END || (act) == ACT_CATCH || (act) == ACT_FINALLY)
			//#define IS_BAD_ACTION_LINE(l) IS_BAD_ACTION_LINE_TYPE((l)->mActionType)
			//if (IS_BAD_ACTION_LINE(line_temp))
			//	return line->PreparseError(ERR_EXPECTED_BLOCK_OR_ACTION);

			// Lexikos: This section once maintained separate variables for file-pattern, registry, file-reading
			// and parsing loops. The intention seemed to be to validate certain commands such as FileAppend
			// differently depending on whether they're contained within a qualifying type of loop (even if some
			// other type of loop lies in between). However, that validation apparently wasn't implemented,
			// and implementing it now seems unnecessary. Doing so would also remove a useful capability:
			//
			//	Loop, Read, %InputFile%, %OutputFile%
			//	{
			//		MyFunc(A_LoopReadLine)
			//	}
			//	MyFunc(line) {
			//		... do some processing on %line% ...
			//		FileAppend, %line%	; This line could be considered an error, though it works in practice.
			//	}
			//

			// Recurse to group the line or lines which are this line's action or body as a
			// single entity and find the line below it.  This must be done even if line_temp
			// isn't an IF/ELSE/LOOP/BLOCK_BEGIN because all lines need mParentLine set by this
			// function, and some other types such as BREAK/CONTINUE also need special handling.
			line_temp = PreparseBlocks(line_temp, ONLY_ONE_LINE, line, line->mAttribute ? line->mAttribute : aLoopType);
			// If not an error, line_temp is now either:
			// 1) If this if's/loop's action was a BEGIN_BLOCK: The line after the end of the block.
			// 2) If this if's/loop's action was another IF or LOOP:
			//    a) the line after that if's else's action; or (if it doesn't have one):
			//    b) the line after that if's/loop's action
			// 3) If this if's/loop's action was some single-line action: the line after that action.
			// In all of the above cases, line_temp is now the line where we
			// would expect to find an ELSE for this IF, if it has one.

			// line_temp is NULL if an error occurred, but should never be NULL in any other
			// case because all scripts end in ACT_EXIT:
			if (line_temp == NULL)
				return NULL; // Error.

			// Set this line's mRelatedLine to the line after its action/body.  For an IF,
			// this is the line to which we jump at runtime when the IF is finished (whether
			// it's condition was true or false), thus skipping over any nested IF's that
			// aren't in blocks beneath it.  If there's no ELSE, the below value serves as
			// the jumppoint we go to when the if-statement is finished.  Example:
			// if x
			//   if y
			//     if z
			//       action1
			//     else
			//       action2
			// action3
			// x's jumppoint should be action3 so that all the nested if's under the
			// first one can be skipped after the "if x" line is recursively evaluated.
			// Because of this behavior (and the fact that all scripts end in ACT_EXIT),
			// all IFs will have a related line.
			line->mRelatedLine = line_temp;

			// Even if aMode == ONLY_ONE_LINE, an IF and its ELSE count as a single
			// statement (one line) due to its very nature (at least for this purpose),
			// so always continue on to evaluate the IF's ELSE, if present.  This also
			// applies to the CATCH or FINALLY belonging to a TRY or CATCH.
			for (bool line_belongs;; )
			{
				// Determine whether line_temp belongs to (is associated with) line.
				// At this point, line is either an IF/LOOP/TRY from above, or an
				// ELSE/CATCH/FINALLY handled by the previous iteration of this loop.
				// line_temp is the line after line's action/body.
				switch (line_temp->mActionType)
				{
				case ACT_ELSE: line_belongs = ACT_IS_IF(line->mActionType); break;
				case ACT_CATCH: line_belongs = (line->mActionType == ACT_TRY); break;
				case ACT_FINALLY: line_belongs = (line->mActionType == ACT_TRY || line->mActionType == ACT_CATCH); break;
				default: line_belongs = false; break;
				}
				if (!line_belongs)
					break;
				// Each line's mParentLine must be set appropriately for named loops to work.
				line_temp->mParentLine = line->mParentLine;
				// Set it up so that line is the ELSE/CATCH/FINALLY and line_temp is it's action.
				// Later, line_temp will be the line after the action/body, and will be checked
				// by the next iteration in case this is a TRY..CATCH and line_temp is FINALLY.
				line = line_temp; // ELSE/CATCH/FINALLY
				line_temp = line->mNextLine; // line's action/body.
				// The following case should be impossible because all scripts end in ACT_EXIT.
				// Thus, it's commented out:
				//if (line_temp == NULL) // An else with no action.
				//	return line->PreparseError(_T("This ELSE has no action."));
				//if (IS_BAD_ACTION_LINE(line_temp)) // See "#define IS_BAD_ACTION_LINE" for comments.
				//	return line->PreparseError(ERR_EXPECTED_BLOCK_OR_ACTION);
				// Assign to line_temp rather than line:
				line_temp = PreparseBlocks(line_temp, ONLY_ONE_LINE, line
					, line->mActionType == ACT_FINALLY ? ATTR_OBSCURE(aLoopType) : aLoopType);
				if (line_temp == NULL)
					return NULL; // Error.
				// Set this ELSE/CATCH/FINALLY's jumppoint.  This is similar to the jumppoint
				// set for an IF/LOOP/TRY, so see related comments above:
				line->mRelatedLine = line_temp;
			}
			if (line_temp->mActionType == ACT_UNTIL
				&& (line->mActionType == ACT_LOOP || line->mActionType == ACT_FOR)) // WHILE is excluded because PerformLoopWhile() doesn't handle UNTIL, due to rarity of need.
			{
				// For consistency/maintainability:
				line_temp->mParentLine = line->mParentLine;
				// Continue processing *after* UNTIL.
				line = line_temp->mNextLine;
			}
			else // continue processing from line_temp's position
				line = line_temp;

			// All cases above have ensured that line is now the first line beyond the
			// scope of the IF/LOOP/TRY and any associated statements.

			if (aMode == ONLY_ONE_LINE) // Return the next unprocessed line to the caller.
				return line;
			// Otherwise, continue processing at line's new location:
			continue;
		} // ActionType is IF/LOOP/TRY.

		// Since above didn't continue, do the switch:
		switch (line->mActionType)
		{
		case ACT_BLOCK_BEGIN:
			if (line->mAttribute == ATTR_TRUE) // This is the opening brace of a function definition.
			{
				if (aParentLine && aParentLine->mActionType != ACT_BLOCK_BEGIN) // Implies ACT_IS_LINE_PARENT(aParentLine->mActionType).  Functions are allowed inside blocks.
					// A function must not be defined directly below an IF/ELSE/LOOP because runtime evaluation won't handle it properly.
					return line->PreparseError(_T("Unexpected function"));
			}
			line_temp = PreparseBlocks(line->mNextLine, UNTIL_BLOCK_END, line, line->mAttribute ? 0 : aLoopType); // mAttribute usage: don't consider a function's body to be inside the loop, since it can be called from outside.
			// "line" is now either NULL due to an error, or the location of the END_BLOCK itself.
			if (line_temp == NULL)
				return NULL; // Error.
			// The BLOCK_BEGIN's mRelatedLine should point to the line *after* the BLOCK_END:
			line->mRelatedLine = line_temp->mNextLine;
			// Since any lines contained inside this block would already have been handled by
			// the recursion in the above call, continue searching from the end of this block:
			line = line_temp;
			break;
		case ACT_BLOCK_END:
			if (aMode == UNTIL_BLOCK_END)
				// Return line rather than line->mNextLine because, if we're at the end of
				// the script, it's up to the caller to differentiate between that condition
				// and the condition where NULL is an error indicator.
				return line;
			// Otherwise, we found an end-block we weren't looking for.
			return line->PreparseError(ERR_UNEXPECTED_CLOSE_BRACE);
		case ACT_BREAK:
		case ACT_CONTINUE:
			if (!aLoopType)
				return line->PreparseError(_T("Break/Continue must be enclosed by a Loop."));
			if (aLoopType == ATTR_LOOP_OBSCURED)
				return line->PreparseError(ERR_BAD_JUMP_INSIDE_FINALLY);
			break;

		case ACT_ELSE:
			// This happens if there's an extra ELSE in this scope level that has no IF:
			return line->PreparseError(ERR_ELSE_WITH_NO_IF);

		case ACT_UNTIL:
			// Similar to above.
			return line->PreparseError(ERR_UNTIL_WITH_NO_LOOP);

		case ACT_CATCH:
			// Similar to above.
			return line->PreparseError(ERR_CATCH_WITH_NO_TRY);

		case ACT_FINALLY:
			// Similar to above.
			return line->PreparseError(ERR_FINALLY_WITH_NO_PRECEDENT);
		} // switch()

		line = line->mNextLine; // If NULL due to physical end-of-script, the for-loop's condition will catch it.
		if (aMode == ONLY_ONE_LINE) // Return the next unprocessed line to the caller.
			// In this case, line shouldn't be (and probably can't be?) NULL because the line after
			// a single-line action shouldn't be the physical end of the script.  That's because
			// the loader has ensured that all scripts now end in ACT_EXIT.  And that final
			// ACT_EXIT should never be parsed here in ONLY_ONE_LINE mode because the only time
			// that mode is used is for the action of an IF, an ELSE, or possibly a LOOP.
			// In all of those cases, the final ACT_EXIT line in the script (which is explicitly
			// inserted by the loader) cannot be the line that was just processed by the
			// switch().  Therefore, the above assignment should not have set line to NULL
			// (which is good because NULL would probably be construed as "failure" by our
			// caller in this case):
			return line;
		// else just continue the for-loop at the new value of line.
	} // for()

	// End of script has been reached.  line is now NULL so don't dereference it.

	// If we were still looking for an EndBlock to match up with a begin, that's an error.
	// This indicates that at least one BLOCK_BEGIN is missing a BLOCK_END.  Let the error
	// message point at the most recent BLOCK_BEGIN (aParentLine) rather than at mLastLine,
	// which points to an EXIT which was added automatically by LoadFromFile().
	if (aMode == UNTIL_BLOCK_END)
		return aParentLine->PreparseError(ERR_MISSING_CLOSE_BRACE);

	// If we were told to process a single line, we were recursed and it should have returned above,
	// so it's an error here (can happen if we were called with aStartingLine == NULL?):
	if (aMode == ONLY_ONE_LINE)
		return mLastLine->PreparseError(_T("Q")); // Placeholder since probably impossible.  Formerly "The script ended while an action was still expected."

	// Otherwise, return something non-NULL to indicate success to the top-level caller:
	return mLastLine;
}



Line *Script::PreparseCommands(Line *aStartingLine)
// Preparse any commands which might rely on blocks having been fully preparsed,
// such as any command which has a jump target (label).
{
	bool in_function_body = false;

	for (Line *line = aStartingLine; line; line = line->mNextLine)
	{
		LPTSTR line_raw_arg1 = LINE_RAW_ARG1; // Resolve only once to help reduce code size.
		LPTSTR line_raw_arg2 = LINE_RAW_ARG2; //
		
		switch (line->mActionType)
		{
		case ACT_BLOCK_BEGIN:
			if (line->mAttribute == ATTR_TRUE) // This is the opening brace of a function definition.
				in_function_body = true; // Must be set only for mAttribute == ATTR_TRUE because functions can of course contain types of blocks other than the function's own block.
			break;
		case ACT_BLOCK_END:
			if (line->mAttribute == ATTR_TRUE) // This is the closing brace of a function definition.
				in_function_body = false; // Must be set only for the above condition because functions can of course contain types of blocks other than the function's own block.
			break;
		case ACT_BREAK:
		case ACT_CONTINUE:
			if (line->mArgc)
			{
				if (line->ArgHasDeref(1) || line->mArg->is_expression)
					// It seems unlikely that computing the target loop at runtime would be useful.
					// For simplicity, rule out things like "break %var%" and "break % func()":
					return line->PreparseError(ERR_PARAM1_INVALID); //_T("Target label of Break/Continue cannot be dynamic."));
				LPTSTR loop_name = line->mArg[0].text;
				Label *loop_label;
				Line *loop_line;
				bool is_numeric = IsPureNumeric(loop_name);
				// If loop_name is a label, find the innermost loop (#1) for validation purposes:
				int n = is_numeric ? _ttoi(loop_name) : 1;
				// Find the nth innermost loop which encloses this line:
				for (loop_line = line->mParentLine; loop_line; loop_line = loop_line->mParentLine)
					if (ACT_IS_LOOP(loop_line->mActionType)) // i.e. LOOP, FOR or WHILE.
						if (--n < 1)
							break;
				if (!loop_line || n != 0)
					return line->PreparseError(ERR_PARAM1_INVALID);
				if (!is_numeric)
				{
					// Target is a named loop.
					if ( !(loop_label = FindLabel(loop_name)) )
						return line->PreparseError(ERR_NO_LABEL, loop_name);
					Line *innermost_loop = loop_line;
					loop_line = loop_label->mJumpToLine;
					// Ensure the label points to a LOOP, FOR or WHILE which encloses this line.
					// Use innermost_loop as the starting-point of the "jump" to ensure the target
					// isn't a loop *inside* the current loop:
					if (   !ACT_IS_LOOP(loop_line->mActionType)
						|| !innermost_loop->IsJumpValid(*loop_label, true)   )
						return line->PreparseError(ERR_PARAM1_INVALID); //_T("Target label does not point to an appropriate Loop."));
					if (loop_line == innermost_loop)
						loop_line = NULL; // Since a normal break/continue will work in this case.
				}
				if (loop_line) // i.e. it wasn't determined to be this line's direct parent, which is always valid.
				{
					if (line->mActionType == ACT_BREAK)
					{
						// Find the line to which we'll actually jump when the loop breaks.
						loop_line = loop_line->mRelatedLine; // From LOOP itself to the line after the LOOP's body.
						if (loop_line->mActionType == ACT_UNTIL)
							loop_line = loop_line->mNextLine;
					}
					if (in_function_body && loop_line->IsOutsideAnyFunctionBody())
						return line->PreparseError(ERR_BAD_JUMP_OUT_OF_FUNCTION);
					if (!line->CheckValidFinallyJump(loop_line))
						return NULL; // Error already shown.
				}
				line->mRelatedLine = loop_line;
			}
			break;

		case ACT_GOSUB: // These two must be done here (i.e. *after* all the script lines have been added),
		case ACT_GOTO:  // so that labels both above and below each Gosub/Goto can be resolved.
			if (line->ArgHasDeref(1))
				// Since the jump-point contains a deref, it must be resolved at runtime:
				line->mRelatedLine = NULL;
  			else
			{
				if (!line->GetJumpTarget(false))
					return NULL; // Error was already displayed by called function.
				if (in_function_body && ((Label *)(line->mRelatedLine))->mJumpToLine->IsOutsideAnyFunctionBody()) // Relies on above call to GetJumpTarget() having set line->mRelatedLine.
				{
					if (line->mActionType == ACT_GOTO)
						return line->PreparseError(ERR_BAD_JUMP_OUT_OF_FUNCTION);
					// Since this Gosub and its target line are both inside a function, they must both
					// be in the same function because otherwise GetJumpTarget() would have reported
					// the target as invalid.
					line->mAttribute = ATTR_TRUE; // v1.0.48.02: To improve runtime performance, mark this Gosub as having a target that is outside of any function body.
				}
				if (line->mActionType == ACT_GOTO && !line->CheckValidFinallyJump(((Label *)(line->mRelatedLine))->mJumpToLine))
					return NULL; // Error already displayed above.
				//else leave mAttribute at its line-constructor default of ATTR_NONE.
			}
			break;

		case ACT_RETURN:
			for (Line *parent = line->mParentLine; parent; parent = parent->mParentLine)
				if (parent->mActionType == ACT_FINALLY)
					return line->PreparseError(ERR_BAD_JUMP_INSIDE_FINALLY);
			break;
			
		// These next 4 must also be done here (i.e. *after* all the script lines have been added),
		// so that labels both above and below this line can be resolved:
		case ACT_ONEXIT:
			if (*line_raw_arg1 && !line->ArgHasDeref(1))
				if (   !(line->mAttribute = FindLabel(line_raw_arg1))   )
					return line->PreparseError(ERR_NO_LABEL);
			break;

		case ACT_HOTKEY:
			if (!line->ArgHasDeref(1))
			{
				if (!_tcsnicmp(line_raw_arg1, _T("If"), 2))
				{
					LPTSTR cp = line_raw_arg1 + 2;
					if (!*cp) // Just "If"
					{
						if (line->mArgc > 2)
							return line->PreparseError(ERR_PARAM3_MUST_BE_BLANK);
						if (*line_raw_arg2 && !line->ArgHasDeref(2))
						{
							// Hotkey, If, Expression: Ensure the expression matches exactly an existing #If,
							// as required by the Hotkey command.  This seems worth doing since the current
							// behaviour might be unexpected (despite being documented), and because typos
							// are likely due to the fact that case and whitespace matter.
							for (HotkeyCriterion *cp = g_FirstHotExpr; ; cp = cp->NextCriterion)
							{
								if (!cp)
									return line->PreparseError(ERR_HOTKEY_IF_EXPR);
								if (!_tcscmp(line_raw_arg2, cp->WinTitle))
									break;
							}
						}
						break;
					}
					if (!_tcsnicmp(cp, _T("Win"), 3))
					{
						cp += 3;
						if (!_tcsnicmp(cp, _T("Not"), 3))
							cp += 3;
						if (!_tcsicmp(cp, _T("Active")) || !_tcsicmp(cp, _T("Exist")))
							break;
					}
					// Since above didn't break, it's something invalid starting with "If".
					return line->PreparseError(ERR_PARAM1_INVALID);
				}
				if (*line_raw_arg2 && !line->ArgHasDeref(2))
					if (   !(line->mAttribute = FindCallable(line_raw_arg2))   )
						if (!Hotkey::ConvertAltTab(line_raw_arg2, true))
							return line->PreparseError(ERR_NO_LABEL);
			}
			break;

		case ACT_SETTIMER:
			if (*line_raw_arg1 && !line->ArgHasDeref(1))
				if (   !(line->mAttribute = FindCallable(line_raw_arg1))   )
					return line->PreparseError(ERR_NO_LABEL);
			if (*line_raw_arg2 && !line->ArgHasDeref(2))
				if (!Line::ConvertOnOff(line_raw_arg2) && !IsPureNumeric(line_raw_arg2, true) // v1.0.46.16: Allow negatives to support the new run-only-once mode.
					&& _tcsicmp(line_raw_arg2, _T("Delete"))
					&& !line->mArg[1].is_expression) // v1.0.46.10: Don't consider expressions THAT CONTAIN NO VARIABLES OR FUNCTION-CALLS like "% 2*500" to be a syntax error.
					return line->PreparseError(ERR_PARAM2_INVALID);
			break;

		case ACT_GROUPADD: // This must be done here because it relies on all other lines already having been added.
			if (*LINE_RAW_ARG4 && !line->ArgHasDeref(4))
			{
				// If the label name was contained in a variable, that label is now resolved and cannot
				// be changed.  This is in contrast to something like "Gosub, %MyLabel%" where a change in
				// the value of MyLabel will change the behavior of the Gosub at runtime:
				Label *label = FindLabel(LINE_RAW_ARG4);
				if (!label)
					return line->PreparseError(ERR_NO_LABEL);
				line->mRelatedLine = (Line *)label; // The script loader has ensured that this can't be NULL.
				// Can't do this because the current line won't be the launching point for the
				// Gosub.  Instead, the launching point will be the GroupActivate rather than the
				// GroupAdd, so it will be checked by the GroupActivate or not at all (since it's
				// not that important in the case of a Gosub -- it's mostly for Goto's):
				//return IsJumpValid(label->mJumpToLine);
			}
			break;
		}
	} // for()
	// Return something non-NULL to indicate success:
	return mLastLine;
}



ResultType Line::ExpressionToPostfix(ArgStruct &aArg)
// Returns OK or FAIL.
{
	// Having a precedence array is required at least for SYM_POWER (since the order of evaluation
	// of something like 2**1**2 does matter).  It also helps performance by avoiding unnecessary pushing
	// and popping of operators to the stack. This array must be kept in sync with "enum SymbolType".
	// Also, dimensioning explicitly by SYM_COUNT helps enforce that at compile-time:
	static UCHAR sPrecedence[SYM_COUNT] =  // Performance: UCHAR vs. INT benches a little faster, perhaps due to the slight reduction in code size it causes.
	{
		0,0,0,0,0,0,0,0,0  // SYM_STRING, SYM_INTEGER, SYM_FLOAT, SYM_MISSING, SYM_VAR, SYM_OPERAND, SYM_OBJECT, SYM_DYNAMIC, SYM_BEGIN (SYM_BEGIN must be lowest precedence).
		, 82, 82         // SYM_POST_INCREMENT, SYM_POST_DECREMENT: Highest precedence operator so that it will work even though it comes *after* a variable name (unlike other unaries, which come before).
		, 86             // SYM_DOT
		, 4,4,4,4,4,4    // SYM_CPAREN, SYM_CBRACKET, SYM_CBRACE, SYM_OPAREN, SYM_OBRACKET, SYM_OBRACE (to simplify the code, parentheses/brackets/braces must be lower than all operators in precedence).
		, 6              // SYM_COMMA -- Must be just above SYM_OPAREN so it doesn't pop OPARENs off the stack.
		, 7,7,7,7,7,7,7,7,7,7,7,7  // SYM_ASSIGN_*. THESE HAVE AN ODD NUMBER to indicate right-to-left evaluation order, which is necessary for cascading assignments such as x:=y:=1 to work.
//		, 8              // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
		, 11, 11         // SYM_IFF_ELSE, SYM_IFF_THEN (ternary conditional).  HAS AN ODD NUMBER to indicate right-to-left evaluation order, which is necessary for ternaries to perform traditionally when nested in each other without parentheses.
//		, 12             // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
		, 16             // SYM_OR
		, 20             // SYM_AND
		, 25             // SYM_LOWNOT (the word "NOT": the low precedence version of logical-not).  HAS AN ODD NUMBER to indicate right-to-left evaluation order so that things like "not not var" are supports (which can be used to convert a variable into a pure 1/0 boolean value).
//		, 26             // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
		, 30, 30, 30     // SYM_EQUAL, SYM_EQUALCASE, SYM_NOTEQUAL (lower prec. than the below so that "x < 5 = var" means "result of comparison is the boolean value in var".
		, 34, 34, 34, 34 // SYM_GT, SYM_LT, SYM_GTOE, SYM_LTOE
		, 38             // SYM_CONCAT
		, 42             // SYM_BITOR -- Seems more intuitive to have these three higher in prec. than the above, unlike C and Perl, but like Python.
		, 46             // SYM_BITXOR
		, 50             // SYM_BITAND
		, 54, 54         // SYM_BITSHIFTLEFT, SYM_BITSHIFTRIGHT
		, 58, 58         // SYM_ADD, SYM_SUBTRACT
		, 62, 62, 62     // SYM_MULTIPLY, SYM_DIVIDE, SYM_FLOORDIVIDE
		, 67,67,67,67,67 // SYM_NEGATIVE (unary minus), SYM_HIGHNOT (the high precedence "!" operator), SYM_BITNOT, SYM_ADDRESS, SYM_DEREF
		// NOTE: THE ABOVE MUST BE AN ODD NUMBER to indicate right-to-left evaluation order, which was added in v1.0.46 to support consecutive unary operators such as !*var !!var (!!var can be used to convert a value into a pure 1/0 boolean).
//		, 68             // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
		, 72             // SYM_POWER (see note below).  Associativity kept as left-to-right for backward compatibility (e.g. 2**2**3 is 4**3=64 not 2**8=256).
		, 77, 77         // SYM_PRE_INCREMENT, SYM_PRE_DECREMENT (higher precedence than SYM_POWER because it doesn't make sense to evaluate power first because that would cause ++/-- to fail due to operating on a non-lvalue.
//		, 78             // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
//		, 82, 82         // RESERVED FOR SYM_POST_INCREMENT, SYM_POST_DECREMENT (which are listed higher above for the performance of YIELDS_AN_OPERAND().
		, 86             // SYM_FUNC -- Has special handling which ensures it stays tightly bound with its parameters as though it's a single operand for use by other operators; the actual value here is irrelevant.
		, 86             // SYM_NEW -- should be popped off the stack immediately after the pseudo function-call which follows it.
		, 36             // SYM_REGEXMATCH
	};
	// Most programming languages give exponentiation a higher precedence than unary minus and logical-not.
	// For example, -2**2 is evaluated as -(2**2), not (-2)**2 (the latter is unsupported by qmathPow anyway).
	// However, this rule requires a small workaround in the postfix-builder to allow 2**-2 to be
	// evaluated as 2**(-2) rather than being seen as an error.  v1.0.45: A similar thing is required
	// to allow the following to work: 2**!1, 2**not 0, 2**~0xFFFFFFFE, 2**&x.
	// On a related note, the right-to-left tradition of something like 2**3**4 is not implemented (maybe in v2).
	// Instead, the expression is evaluated from left-to-right (like other operators) to simplify the code.

	ExprTokenType infix[MAX_TOKENS], *postfix[MAX_TOKENS], *stack[MAX_TOKENS + 1];  // +1 for SYM_BEGIN on the stack.
	int infix_count = 0, postfix_count = 0, stack_count = 0;
	// Above dimensions the stack to be as large as the infix/postfix arrays to cover worst-case
	// scenarios and avoid having to check for overflow.  For the infix-to-postfix conversion, the
	// stack must be large enough to hold a malformed expression consisting entirely of operators
	// (though other checks might prevent this).  It must also be large enough for use by the final
	// expression evaluation phase, the worst case of which is unknown but certainly not larger
	// than MAX_TOKENS.

	///////////////////////////////////////////////////////////////////////////////////////////////
	// TOKENIZE THE INFIX EXPRESSION INTO AN INFIX ARRAY: Avoids the performance overhead of having
	// to re-detect whether each symbol is an operand vs. operator at multiple stages.
	///////////////////////////////////////////////////////////////////////////////////////////////
	// In v1.0.46.01, this section was simplified to avoid transcribing the entire expression into the
	// deref buffer.  In addition to improving performance and reducing code size, this also solves
	// obscure timing bugs caused by functions that have side-effects, especially in comma-separated
	// sub-expressions.  In these cases, one part of an expression could change a built-in variable
	// (indirectly or in the case of Clipboard, directly), an environment variable, or a double-def.
	// For example the dynamic components of a double-deref can be changed by other parts of an
	// expression, even one without commas.  Another example is: fn(clipboard, func_that_changes_clip()).
	// So now, built-in & environment variables and double-derefs are resolve when they're actually
	// encountered during the final/evaluation phase.
	// Another benefit to deferring the resolution of these types of items is that they become eligible
	// for short-circuiting, which further helps performance (they're quite similar to built-in
	// functions in this respect).
	LPTSTR op_end, cp;
	DerefType *deref, *this_deref, *deref_start, *deref_new;
	int derefs_in_this_double;
	int cp1; // int vs. char benchmarks slightly faster, and is slightly smaller in code size.

	for (cp = aArg.text, deref = aArg.deref // Start at the beginning of this arg's text and look for the next deref.
		;; ++deref, ++infix_count) // FOR EACH DEREF IN AN ARG:
	{
		this_deref = deref && deref->marker ? deref : NULL; // A deref with a NULL marker terminates the list (i.e. the final deref isn't a deref, merely a terminator of sorts.

		// BEFORE PROCESSING "this_deref" ITSELF, MUST FIRST PROCESS ANY LITERAL/RAW TEXT THAT LIES TO ITS LEFT.
		if (this_deref && cp < this_deref->marker // There's literal/raw text to the left of the next deref.
			|| !this_deref && *cp) // ...or there's no next deref, but there's some literal raw text remaining to be processed.
		{
			for (;; ++infix_count) // FOR EACH TOKEN INSIDE THIS RAW/LITERAL TEXT SECTION.
			{
				// Because neither the postfix array nor the stack can ever wind up with more tokens than were
				// contained in the original infix array, only the infix array need be checked for overflow:
				if (infix_count > MAX_TOKENS - 1) // No room for this operator or operand to be added.
					return LineError(ERR_EXPR_TOO_LONG);

				// Only spaces and tabs are considered whitespace, leaving newlines and other whitespace characters
				// for possible future use:
				cp = omit_leading_whitespace(cp);
				if (!*cp // Very end of expression...
					|| this_deref && cp >= this_deref->marker) // ...or no more literal/raw text left to process at the left side of this_deref.
					break; // Break out of inner loop so that bottom of the outer loop will process this_deref itself.

				ExprTokenType &this_infix_item = infix[infix_count]; // Might help reduce code size since it's referenced many places below.
				this_infix_item.deref = NULL; // Init needed for SYM_ASSIGN and related; a non-NULL deref means it should be converted to an object-assignment.

				// CHECK IF THIS CHARACTER IS AN OPERATOR.
				cp1 = cp[1]; // Improves performance by nearly 5% and appreciably reduces code size (at the expense of being less maintainable).
				switch (*cp)
				{
				// The most common cases are kept up top to enhance performance if switch() is implemented as if-else ladder.
				case '+':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_ASSIGN_ADD;
					}
					else
					{
						if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))
						{
							if (cp1 == '+')
							{
								// For consistency, assume that since the previous item is an operand (even if it's
								// ')'), this is a post-op that applies to that operand.  For example, the following
								// are all treated the same for consistency (implicit concatenation where the '.'
								// is omitted is rare anyway).
								// x++ y
								// x ++ y
								// x ++y
								// The following implicit concat is deliberately unsupported:
								//    "string" ++x
								// The ++ above is seen as applying to the string because it doesn't seem worth
								// the complexity to distinguish between expressions that can accept a post-op
								// and those that can't (operands other than variables can have a post-op;
								// e.g. (x:=y)++).
								++cp; // An additional increment to have loop skip over the operator's second symbol.
								this_infix_item.symbol = SYM_POST_INCREMENT;
							}
							else
								this_infix_item.symbol = SYM_ADD;
						}
						else if (cp1 == '+') // Pre-increment.
						{
							++cp; // An additional increment to have loop skip over the operator's second symbol.
							this_infix_item.symbol = SYM_PRE_INCREMENT;
						}
						else // Remove unary pluses from consideration since they do not change the calculation.
							--infix_count; // Counteract the loop's increment.
					}
					break;
				case '-':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_ASSIGN_SUBTRACT;
						break;
					}
					// Otherwise (since above didn't "break"):
					// Must allow consecutive unary minuses because otherwise, the following example
					// would not work correctly when y contains a negative value: var := 3 * -y
					if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))
					{
						if (cp1 == '-')
						{
							// See comments at SYM_POST_INCREMENT about this section.
							++cp; // An additional increment to have loop skip over the operator's second symbol.
							this_infix_item.symbol = SYM_POST_DECREMENT;
						}
						else
							this_infix_item.symbol = SYM_SUBTRACT;
					}
					else if (cp1 == '-') // Pre-decrement.
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_PRE_DECREMENT;
					}
					else // Unary minus.
					{
						// Set default for cases where the processing below this line doesn't determine
						// it's a negative numeric literal:
						this_infix_item.symbol = SYM_NEGATIVE;
						// v1.0.40.06: The smallest signed 64-bit number (-0x8000000000000000) wasn't properly
						// supported in previous versions because its unary minus was being seen as an operator,
						// and thus the raw number was being passed as a positive to _atoi64() or _strtoi64(),
						// neither of which would recognize it as a valid value.  To correct this, a unary
						// minus followed by a raw numeric literal is now treated as a single literal number
						// rather than unary minus operator followed by a positive number.
						//
						// To be a valid "literal negative number", the character immediately following
						// the unary minus must not be:
						// 1) Whitespace (atoi() and such don't support it, nor is it at all conventional).
						// 2) An open-parenthesis such as the one in -(x).
						// 3) Another unary minus or operator such as --x (which is the pre-decrement operator).
						// To cover the above and possibly other unforeseen things, insist that the first
						// character be a digit (even a hex literal must start with 0).
						if ((cp1 >= '0' && cp1 <= '9') || cp1 == '.') // v1.0.46.01: Recognize dot too, to support numbers like -.5.
						{
							for (op_end = cp + 2; !_tcschr(EXPR_OPERAND_TERMINATORS_EX_DOT, *op_end); ++op_end); // Find the end of this number (can be '\0').
							// 1.0.46.11: Due to obscurity, no changes have been made here to support scientific
							// notation followed by the power operator; e.g. -1.0e+1**5.
							if (!this_deref || op_end < this_deref->marker) // Detect numeric double derefs such as one created via "12%i% = value".
							{
								// Because the power operator takes precedence over unary minus, don't collapse
								// unary minus into a literal numeric literal if the number is immediately
								// followed by the power operator.  This is correct behavior even for
								// -0x8000000000000000 because -0x8000000000000000**2 would in fact be undefined
								// because ** is higher precedence than unary minus and +0x8000000000000000 is
								// beyond the signed 64-bit range.  SEE ALSO the comments higher above.
								// Use a temp variable because numeric_literal requires that op_end be set properly:
								LPTSTR pow_temp = omit_leading_whitespace(op_end);
								if (!(pow_temp[0] == '*' && pow_temp[1] == '*'))
									goto numeric_literal; // Goto is used for performance and also as a patch to minimize the chance of breaking other things via redesign.
								//else it's followed by pow.  Since pow is higher precedence than unary minus,
								// leave this unary minus as an operator so that it will take effect after the pow.
							}
							//else possible double deref, so leave this unary minus as an operator.
						}
					} // Unary minus.
					break;
				case ',':
					this_infix_item.symbol = SYM_COMMA; // Used to separate sub-statements and function parameters.
					break;
				case '/':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_ASSIGN_DIVIDE;
					}
					else if (cp1 == '/')
					{
						if (cp[2] == '=')
						{
							cp += 2; // An additional increment to have loop skip over the operator's 2nd & 3rd symbols.
							this_infix_item.symbol = SYM_ASSIGN_FLOORDIVIDE;
						}
						else
						{
							++cp; // An additional increment to have loop skip over the second '/' too.
							this_infix_item.symbol = SYM_FLOORDIVIDE;
						}
					}
					else
						this_infix_item.symbol = SYM_DIVIDE;
					break;
				case '*':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_ASSIGN_MULTIPLY;
					}
					else
					{
						if (cp1 == '*') // Python, Perl, and other languages also use ** for power.
						{
							++cp; // An additional increment to have loop skip over the second '*' too.
							this_infix_item.symbol = SYM_POWER;
						}
						else
						{
							// Differentiate between unary dereference (*) and the "multiply" operator:
							// See '-' above for more details:
							this_infix_item.symbol = (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))
								? SYM_MULTIPLY : SYM_DEREF;
						}
					}
					break;
				case '!':
					if (cp1 == '=') // i.e. != is synonymous with <>, which is also already supported by legacy.
					{
						++cp; // An additional increment to have loop skip over the '=' too.
						this_infix_item.symbol = SYM_NOTEQUAL;
					}
					else
						// If what lies to its left is a CPARAN or OPERAND, SYM_CONCAT is not auto-inserted because:
						// 1) Allows ! and ~ to potentially be overloaded to become binary and unary operators in the future.
						// 2) Keeps the behavior consistent with unary minus, which could never auto-concat since it would
						//    always be seen as the binary subtract operator in such cases.
						// 3) Simplifies the code.
						this_infix_item.symbol = SYM_HIGHNOT; // High-precedence counterpart of the word "not".
					break;
				case '(':
					// The below should not hurt any future type-casting feature because the type-cast can be checked
					// for prior to checking the below.  For example, if what immediately follows the open-paren is
					// the string "int)", this symbol is not open-paren at all but instead the unary type-cast-to-int
					// operator.
					if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol)
						&& _tcschr(_T(" \t)\""), cp[-1])) // For backward-compatibility, )( and "foo"(bar) are allowed.  Otherwise, a space/tab is required, as documented and so things like x[y]() don't need to deal with an extraneous SYM_CONCAT.
					{
						if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
							return LineError(ERR_EXPR_TOO_LONG);
						this_infix_item.symbol = SYM_CONCAT;
						++infix_count;
					}
					infix[infix_count].symbol = SYM_OPAREN; // MUST NOT REFER TO this_infix_item IN CASE ABOVE DID ++infix_count.
					break;
				case ')':
					this_infix_item.symbol = SYM_CPAREN;
					break;
				case '[': // L31
					if (infix_count && infix[infix_count - 1].symbol == SYM_DOT // obj.x[ ...
						&& *omit_leading_whitespace(cp + 1) != ']') // not obj.x[]
					{
						// L36: This is something like obj.x[y] or obj.x[y]:=z, which should be treated
						//		as a single operation such as ObjGet(obj,"x",y) or ObjSet(obj,"x",y,z).
						--infix_count;
						// Below will change the SYM_DOT token into SYM_OBRACKET, keeping the existing deref struct.
					}
					else
					{
						if (  !(deref_new = (DerefType *)SimpleHeap::Malloc(sizeof(DerefType)))  )
							return LineError(ERR_OUTOFMEM);
						if (  !(infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))  )
						{	// Array constructor; e.g. x := [1,2,3]
							deref_new->func = g_script.FindFunc(_T("Array"));
							deref_new->param_count = 0;
						}
						else
						{
							deref_new->func = &g_ObjGet; // This may be overridden by standard_pop_into_postfix.
							deref_new->param_count = 1; // Initially one parameter: the target object.
						}
						deref_new->marker = cp; // For error-reporting.
						deref_new->is_function = true;
						this_infix_item.deref = deref_new;
					}
					// This SYM_OBRACKET will be converted to SYM_FUNC after we determine what type of operation
					// is being performed.  SYM_FUNC requires a deref structure to point to the appropriate 
					// function; we will also use it to count parameters as each SYM_COMMA is encountered.
					// deref->func will be set at a later stage.  deref->is_function need not be set.
					// DO NOT USE this_infix_item in case above did --infix_count.
					infix[infix_count].symbol = SYM_OBRACKET;
					break;
				case ']': // L31
					this_infix_item.symbol = SYM_CBRACKET;
					this_infix_item.buf = cp; // Used to detect "obj[method_name](param)".
					break;
				case '{':
					if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))
						return LineError(_T("Unexpected \"{\""));
					if (  !(deref_new = (DerefType *)SimpleHeap::Malloc(sizeof(DerefType)))  )
						return LineError(ERR_OUTOFMEM);
					deref_new->func = g_script.FindFunc(_T("Object"));
					deref_new->is_function = true;
					deref_new->param_count = 0;
					deref_new->marker = cp; // For error-reporting.
					this_infix_item.deref = deref_new;
					this_infix_item.symbol = SYM_OBRACE;
					break;
				case '}':
					this_infix_item.symbol = SYM_CBRACE;
					this_infix_item.buf = cp; // Set mainly so CBRACE can share code with CBRACKET.
					break;
				case '=':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the other '=' too.
						this_infix_item.symbol = SYM_EQUALCASE;
					}
					else
						this_infix_item.symbol = SYM_EQUAL;
					break;
				case '>':
					switch (cp1)
					{
					case '=':
						++cp; // An additional increment to have loop skip over the '=' too.
						this_infix_item.symbol = SYM_GTOE;
						break;
					case '>':
						if (cp[2] == '=')
						{
							cp += 2; // An additional increment to have loop skip over the operator's 2nd & 3rd symbols.
							this_infix_item.symbol = SYM_ASSIGN_BITSHIFTRIGHT;
						}
						else
						{
							++cp; // An additional increment to have loop skip over the second '>' too.
							this_infix_item.symbol = SYM_BITSHIFTRIGHT;
						}
						break;
					default:
						this_infix_item.symbol = SYM_GT;
					}
					break;
				case '<':
					switch (cp1)
					{
					case '=':
						++cp; // An additional increment to have loop skip over the '=' too.
						this_infix_item.symbol = SYM_LTOE;
						break;
					case '>':
						++cp; // An additional increment to have loop skip over the '>' too.
						this_infix_item.symbol = SYM_NOTEQUAL;
						break;
					case '<':
						if (cp[2] == '=')
						{
							cp += 2; // An additional increment to have loop skip over the operator's 2nd & 3rd symbols.
							this_infix_item.symbol = SYM_ASSIGN_BITSHIFTLEFT;
						}
						else
						{
							++cp; // An additional increment to have loop skip over the second '<' too.
							this_infix_item.symbol = SYM_BITSHIFTLEFT;
						}
						break;
					default:
						this_infix_item.symbol = SYM_LT;
					}
					break;
				case '&':
					if (cp1 == '&')
					{
						++cp; // An additional increment to have loop skip over the second '&' too.
						this_infix_item.symbol = SYM_AND;
					}
					else if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_ASSIGN_BITAND;
					}
					else
					{
						// Differentiate between unary "take the address of" and the "bitwise and" operator:
						// See '-' above for more details:
						this_infix_item.symbol = (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))
							? SYM_BITAND : SYM_ADDRESS;
					}
					break;
				case '|':
					if (cp1 == '|')
					{
						++cp; // An additional increment to have loop skip over the second '|' too.
						this_infix_item.symbol = SYM_OR;
					}
					else if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_ASSIGN_BITOR;
					}
					else
						this_infix_item.symbol = SYM_BITOR;
					break;
				case '^':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_ASSIGN_BITXOR;
					}
					else
						this_infix_item.symbol = SYM_BITXOR;
					break;
				case '~':
					if (cp1 == '=')
					{
						++cp;
						if (   !(this_infix_item.deref = (DerefType *)SimpleHeap::Malloc(sizeof(DerefType)))   )
							return LineError(ERR_OUTOFMEM);
						this_infix_item.deref->func = g_script.FindFunc(_T("RegExMatch"));
						this_infix_item.deref->is_function = true;
						this_infix_item.deref->param_count = 2;
						this_infix_item.symbol = SYM_REGEXMATCH;
					}
					else
					// If what lies to its left is a CPARAN or OPERAND, SYM_CONCAT is not auto-inserted because:
					// 1) Allows ! and ~ to potentially be overloaded to become binary and unary operators in the future.
					// 2) Keeps the behavior consistent with unary minus, which could never auto-concat since it would
					//    always be seen as the binary subtract operator in such cases.
					// 3) Simplifies the code.
					this_infix_item.symbol = SYM_BITNOT;
					break;
				case '?':
					this_infix_item.symbol = SYM_IFF_THEN;
					break;
				case ':':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over the second '|' too.
						this_infix_item.symbol = SYM_ASSIGN;
					}
					else
						this_infix_item.symbol = SYM_IFF_ELSE;
					break;

				case '"': // QUOTED/LITERAL STRING.
					// Note that single and double-derefs are impossible inside string-literals
					// because the load-time deref parser would never detect anything inside
					// of quotes -- even non-escaped percent signs -- as derefs.
					if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
					{
						if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
							return LineError(ERR_EXPR_TOO_LONG);
						this_infix_item.symbol = SYM_CONCAT;
						++infix_count;
					}
					// The following section is nearly identical to one in DefineFunc().
					// Find the end of this string literal, noting that a pair of double quotes is
					// a literal double quote inside the string:
					for (op_end = ++cp;;) // Omit the starting-quote from consideration, and from the resulting/built string.
					{
						if (!*op_end) // No matching end-quote. Probably impossible due to load-time validation.
							return LineError(ERR_MISSING_CLOSE_QUOTE); // Since this error string is used in other places, compiler string pooling should result in little extra memory needed for this line.
						if (*op_end == '"') // And if it's not followed immediately by another, this is the end of it.
						{
							++op_end;
							if (*op_end != '"') // String terminator or some non-quote character.
								break;  // The previous char is the ending quote.
							//else a pair of quotes, which resolves to a single literal quote. So fall through
							// to the below, which will copy of quote character to the buffer. Then this pair
							// is skipped over and the loop continues until the real end-quote is found.
						}
						//else some character other than '\0' or '"'.
						++op_end;
					}
					// Since above didn't "goto", op_end is now the character after the ending '"'.

					// MUST NOT REFER TO this_infix_item IN CASE HIGHER ABOVE DID ++infix_count:
					infix[infix_count].symbol = SYM_STRING; // Marked explicitly as string vs. SYM_OPERAND to prevent it from being seen as a number, e.g. if (var == "12.0") would be false if var contains "12" with no trailing ".0".
					if (   !(infix[infix_count].marker = SimpleHeap::Malloc(cp, op_end - cp - 1))   ) // -1 to omit the ending quote.  cp was already adjusted to omit the starting quote.
						return LineError(ERR_OUTOFMEM);
					StrReplace(infix[infix_count].marker, _T("\"\""), _T("\""), SCS_SENSITIVE); // Resolve each "" into a single ".  Consequently, a little bit of memory in "marker" might be wasted, but it doesn't seem worth the code size to compensate for this.
					cp = omit_leading_whitespace(op_end); // Have the loop process whatever lies at op_end and beyond.
					
					if (*cp && _tcschr(_T("+-*&~!"), *cp) && cp[1] != '=' && (cp[1] != '&' || *cp != '&'))
					{
						// The symbol following this literal string is either a unary operator, or a
						// binary operator for which literal strings are not valid input.  Instead of
						// treating it as a syntax error (which may be difficult for the user to see),
						// we will insert a concat operator and allow the symbol to be interpreted as
						// a unary operator.  The most common cases where this helps are:
						//	MsgBox % "var's address is " &var
						//	MsgBox % "counter is now " ++var
						if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
							return LineError(ERR_EXPR_TOO_LONG);
						infix[++infix_count].symbol = SYM_CONCAT;
					}
					continue; // Continue vs. break to avoid the ++cp at the bottom. Above has already set cp to be the character after this literal string's close-quote.

				default: // NUMERIC-LITERAL, DOUBLE-DEREF, RELATIONAL OPERATOR SUCH AS "NOT", OR UNRECOGNIZED SYMBOL.
					if (*cp == '.') // This one must be done here rather than as a "case".  See comment below.
					{
						if (cp1 == '=')
						{
							++cp; // An additional increment to have loop skip over the operator's second symbol.
							this_infix_item.symbol = SYM_ASSIGN_CONCAT;
							break;
						}
						if (IS_SPACE_OR_TAB(cp1))
						{
							if (!infix_count || !IS_SPACE_OR_TAB(cp[-1])) // L31: Disallow things like "obj. name" since it seems ambiguous.  Checking infix_count ensures cp is not the beginning of the string (". " at the beginning of an expression would also be invalid).  Checking IS_SPACE_OR_TAB seems more appropriate than looking in EXPR_OPERAND_TERMINATORS; it enforces the previously unenforced but documented rule that concat-dot requires a space on either side.
								return LineError(ERR_INVALID_DOT, FAIL, cp);
							this_infix_item.symbol = SYM_CONCAT;
							break;
						}
						//else this is a '.' that isn't followed by a space, tab, or '='.  So it's probably
						// a number without a leading zero like .2, so continue on below to process it.
						// BACKWARD COMPATIBILITY: The above behavior creates ambiguity between a "pure"
						// concat operator (.) and numbers that begin with a decimal point.  I think that
						// came about because the concat operator was added after numbers like .5 had been
						// supported a long time.  In any case, it's documented that the '.' operator must
						// have a space on both sides to be valid, and maybe automatic/implicit concatenation
						// handles most such situations properly anyway (e.g. the expression "x .5").

						// L31
						// Rather than checking if this ".xxx" is numeric and trying to detect things like "obj.123" or ".123e+1",
						// use a simpler method: if . follows something which yields an operand, treat it as member access. As a
						// side-effect, implicit concatenation is no longer supported for floating point numbers beginning with ".".
						if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))
						{
							// Ensure at least enough room for this operand and the operator further below.
							if (infix_count > MAX_TOKENS - 2)
								return LineError(ERR_EXPR_TOO_LONG);
							
							// Skip this '.'
							++cp;

							// Find the end of the operand (".operand"):
							op_end = find_identifier_end(cp);
							if (!_tcschr(EXPR_OPERAND_TERMINATORS, *op_end))
								return LineError(ERR_INVALID_CHAR, FAIL, op_end);

							// Rather than trying to predict how something like "obj.-1" will be handled, treat it as a syntax error.
							// "obj.()" is allowed; it should mean "call the default method of obj" or "call the function object obj".
							if (op_end == cp && *op_end != '(')
								return LineError(ERR_INVALID_DOT, FAIL, cp-1); // Intentionally vague since the user's intention isn't clear.

							bool is_new_op = false;
							// For the '(' check below, determine if this op is part of a "new" operation, such as "new Class.NestedClass()".
							for (ExprTokenType *prev_infix = infix + infix_count - 1; prev_infix >= infix; prev_infix -= 2)
							{
								if (prev_infix->symbol != SYM_DOT)
								{
									if (IS_OPERAND(prev_infix->symbol) && prev_infix > infix)
										--prev_infix; // This is the target of the SYM_DOT, as in "target.foo".
									is_new_op = (prev_infix->symbol == SYM_NEW);
									break;
								}
							}

							// Output a SYM_OPERAND for the text following '.'
							infix[infix_count].symbol = SYM_OPERAND;
							if (   !(infix[infix_count].marker = SimpleHeap::Malloc(cp, op_end - cp))   )
								return LineError(ERR_OUTOFMEM);
							++infix_count;

							SymbolType new_symbol; // Type of token: SYM_FUNC or SYM_DOT (which must be treated differently as it doesn't have parentheses).
							DerefType *new_deref; // Holds a reference to the appropriate function, and parameter count.
							if (   !(new_deref = (DerefType *)SimpleHeap::Malloc(sizeof(DerefType)))   )
								return LineError(ERR_OUTOFMEM);
							new_deref->marker = cp - 1; // Not typically needed, set for error-reporting.
							new_deref->param_count = 2; // Initially two parameters: the object and identifier.
							new_deref->is_function = true;
							
							if (*op_end == '(' && !is_new_op)
							{
								new_symbol = SYM_FUNC;
								new_deref->func = &g_ObjCall;
								// DON'T DO THE FOLLOWING - must let next iteration handle '(' so it outputs a SYM_OPAREN:
								//++op_end;
							}
							else
							{
								new_symbol = SYM_DOT; // This will be changed to SYM_FUNC at a later stage.
								new_deref->func = &g_ObjGet; // Set default; may be overridden by standard_pop_into_postfix.
							}

							// Output the operator next - after the operand to avoid auto-concat.
							infix[infix_count].symbol = new_symbol;
							infix[infix_count].deref = new_deref;

							// Continue processing after this operand. Outer loop will do ++infix_count.
							cp = op_end;
							continue;
						}
					}

					// Find the end of this operand or keyword, even if that end extended into the next deref.
					// StrChrAny() is not used because if *op_end is '\0', the strchr() below will find it too:
					for (op_end = cp + 1; !_tcschr(EXPR_OPERAND_TERMINATORS, *op_end); ++op_end);
					// Now op_end marks the end of this operand or keyword.  That end might be the zero terminator
					// or the next operator in the expression, or just a whitespace.
					if (this_deref && op_end >= this_deref->marker)
						goto double_deref; // This also serves to break out of the inner for(), equivalent to a break.
					if (*op_end == '.')
					{
						// Since this isn't a double deref, it can probably only be a numeric literal with decimal portion.
						// Update op_end to include the decimal portion of the operand:
						for (++op_end; !_tcschr(EXPR_OPERAND_TERMINATORS, *op_end); ++op_end);
					}

					// Otherwise, this operand is a normal raw numeric-literal or a word-operator (and/or/not).
					// The section below is very similar to the one used at load-time to recognize and/or/not,
					// so it should be maintained with that section.  UPDATE for v1.0.45: The load-time parser
					// now resolves "OR" to || and "AND" to && to improve runtime performance and reduce code size here.
					// However, "NOT" must still be parsed here at runtime because it's not quite the same as the "!"
					// operator (different precedence), and it seemed too much trouble to invent some special
					// operator symbol for load-time to insert as a placeholder/substitute (especially since that
					// symbol would appear in ListLines).
					if (op_end-cp == 3
						&& (cp[0] == 'n' || cp[0] == 'N')
						&& *omit_leading_whitespace(op_end) != ':') // Exclude "not:" and "new:", which are either the key in a key-value pair or a syntax error.
					{
						if (   (cp1   == 'o' || cp1   == 'O')
							&& (cp[2] == 't' || cp[2] == 'T')   ) // "NOT" was found.
						{
							this_infix_item.symbol = SYM_LOWNOT;
							cp = op_end; // Have the loop process whatever lies at op_end and beyond.
							continue; // Continue vs. break to avoid the ++cp at the bottom (though it might not matter in this case).
						}
						if (   (cp1   == 'e' || cp1   == 'E')
							&& (cp[2] == 'w' || cp[2] == 'W')   ) // "NEW" was found.
						{
							if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))
							{
								if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
									return LineError(ERR_EXPR_TOO_LONG);
								this_infix_item.symbol = SYM_CONCAT;
								++infix_count;
							}
							// A previous stage ensured this "new" is followed by something which looks like
							// a function call.  Push this pseudo-operator onto the stack.  When the open-
							// parenthesis is encountered, symbol will be changed to SYM_FUNC.
							if (  !(deref_new = (DerefType *)SimpleHeap::Malloc(sizeof(DerefType)))  )
								return LineError(ERR_OUTOFMEM);
							deref_new->marker = cp; // For error-reporting.
							deref_new->param_count = 1; // Start counting at the class object, which precedes the open-parenthesis.
							deref_new->func = &g_ObjNew;
							deref_new->is_function = true;
							infix[infix_count].symbol = SYM_NEW;
							infix[infix_count].deref = deref_new;
							cp = op_end; // See comments above.
							continue;    //
						}
					}
numeric_literal:
					// Since above didn't "continue", this item is probably a raw numeric literal (either SYM_FLOAT
					// or SYM_INTEGER, to be differentiated later) because just about every other possibility has
					// been ruled out above.  For example, unrecognized symbols should be impossible at this stage
					// because load-time validation would have caught them.  And any kind of unquoted alphanumeric
					// characters (other than "NOT", which was detected above) wouldn't have reached this point
					// because load-time pre-parsing would have marked it as a deref/var, not raw/literal text.
					if (   (*op_end == '-' || *op_end == '+') && ctoupper(op_end[-1]) == 'E' // v1.0.46.11: It looks like scientific notation...
						&& !(cp[0] == '0' && ctoupper(cp[1]) == 'X') // ...and it's not a hex number (this check avoids falsely detecting hex numbers that end in 'E' as exponents). This line fixed in v1.0.46.12.
						&& !(cp[0] == '-' && cp[1] == '0' && ctoupper(cp[2]) == 'X') // ...and it's not a negative hex number (this check avoids falsely detecting hex numbers that end in 'E' as exponents). This line added as a fix in v1.0.47.03.
						)
					{
						// Since op_end[-1] is the 'E' or an exponent, the only valid things for op_end[0] to be
						// are + or - (it can't be a digit because the loop above would never have stopped op_end
						// at a digit).  If it isn't + or -, it's some kind of syntax error, so doing the following
						// seems harmless in any case:
						do // Skip over the sign and its exponent; e.g. the "+1" in "1.0e+1".  There must be a sign in this particular sci-notation number or we would never have arrived here.
							++op_end;
						while (*op_end >= '0' && *op_end <= '9'); // Avoid isdigit() because it sometimes causes a debug assertion failure at: (unsigned)(c + 1) <= 256 (probably only in debug mode), and maybe only when bad data got in it due to some other bug.
					}
					if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
					{
						if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
							return LineError(ERR_EXPR_TOO_LONG);
						this_infix_item.symbol = SYM_CONCAT;
						++infix_count;
					}
					// MUST NOT REFER TO this_infix_item IN CASE ABOVE DID ++infix_count:
					infix[infix_count].symbol = SYM_OPERAND;
					if (   !(infix[infix_count].marker = SimpleHeap::Malloc(cp, op_end - cp))   )
						return LineError(ERR_OUTOFMEM);
					cp = op_end; // Have the loop process whatever lies at op_end and beyond.
					continue; // "Continue" to avoid the ++cp at the bottom.
				} // switch() for type of symbol/operand.
				++cp; // i.e. increment only if a "continue" wasn't encountered somewhere above. Although maintainability is reduced to do this here, it avoids dozens of ++cp in other places.
			} // for each token in this section of raw/literal text.
		} // End of processing of raw/literal text (such as operators) that lie to the left of this_deref.

		if (!this_deref) // All done because the above just processed all the raw/literal text (if any) that
			break;       // lay to the right of the last deref.

		// THE ABOVE HAS NOW PROCESSED ANY/ALL RAW/LITERAL TEXT THAT LIES TO THE LEFT OF this_deref.
		// SO NOW PROCESS THIS_DEREF ITSELF.
		if (infix_count > MAX_TOKENS - 1) // No room for the deref item below to be added.
			return LineError(ERR_EXPR_TOO_LONG);
		DerefType &this_deref_ref = *this_deref; // Boosts performance slightly.
		if (this_deref_ref.is_function) // Above has ensured that at this stage, this_deref!=NULL.
		{
			if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
			{
				if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
					return LineError(ERR_EXPR_TOO_LONG);
				infix[infix_count++].symbol = SYM_CONCAT;
			}
			infix[infix_count].symbol = SYM_FUNC;
			infix[infix_count].deref = this_deref;
			// L31: Initialize param_count to zero to work with new method of parameter counting required for ObjGet/Set/Call. (See SYM_COMMA and SYM_'PAREN handling.)
			this_deref_ref.param_count = 0;
		}
		else // this_deref is a variable.
		{
			if (*this_deref_ref.marker == g_DerefChar) // A double-deref because normal derefs don't start with '%'.
			{
				// Find the end of this operand, even if that end extended into the next deref.
				// StrChrAny() is not used because if *op_end is '\0', the _tcschr() below will find it too:
				for (op_end = this_deref_ref.marker + this_deref_ref.length; !_tcschr(EXPR_OPERAND_TERMINATORS, *op_end); ++op_end);
				goto double_deref;
			}
			else
			{
				if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
				{
					if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
						return LineError(ERR_EXPR_TOO_LONG);
					infix[infix_count++].symbol = SYM_CONCAT;
				}
				infix[infix_count].var = this_deref_ref.var; // Set this first to allow optimizations below to override it.
				if (this_deref_ref.var->Type() == VAR_NORMAL // VAR_ALIAS is taken into account (and resolved) by Type().
					&& g_NoEnv) // v1.0.43.08: Added g_NoEnv.  Relies on short-circuit boolean order.
					// "!this_deref_ref.var->Get()" isn't checked here.  See comments in SYM_DYNAMIC evaluation.
				{
					// DllCall() and possibly others rely on this having been done to support changing the
					// value of a parameter (similar to by-ref).
					infix[infix_count].symbol = SYM_VAR; // Type() is VAR_NORMAL as verified above; but SYM_VAR can be the clipboard in the case of expression lvalues.  Search for VAR_CLIPBOARD further below for details.
					infix[infix_count].is_lvalue = FALSE; // Set default.  This simplifies #Warn ClassOverwrite (vs. storing it in the assignment token).
				}
				else // It's either a built-in variable (including clipboard) OR a possible environment variable.
				{
					// The following "variables" previously had optimizations in ExpandExpression(),
					// but since their values never change at run-time, it is better to do it here:
					if (this_deref_ref.var->mBIV == BIV_True_False)
					{
						infix[infix_count].symbol = SYM_INTEGER;
						infix[infix_count].value_int64 = (ctoupper(*this_deref_ref.marker) == 'T');
					}
					else if (this_deref_ref.var->mBIV == BIV_PtrSize)
					{
						infix[infix_count].symbol = SYM_INTEGER;
						infix[infix_count].value_int64 = sizeof(void*);
					}
					else if (this_deref_ref.var->mBIV == BIV_IsUnicode)
					{
#ifdef UNICODE
						infix[infix_count].symbol = SYM_INTEGER;
						infix[infix_count].value_int64 = 1;
#else
						infix[infix_count].symbol = SYM_STRING;
						infix[infix_count].marker = _T(""); // See BIV_IsUnicode for comments about why it is blank.
#endif
					}
					else
					{
						infix[infix_count].symbol = SYM_DYNAMIC;
						infix[infix_count].buf = NULL; // SYM_DYNAMIC requires that buf be set to NULL for non-double-deref vars (since there are two different types of SYM_DYNAMIC).
					}
				}
			}
		} // Handling of the var or function in this_deref.

		// Finally, jump over the dereference text. Note that in the case of an expression, there might not
		// be any percent signs within the text of the dereference, e.g. x + y, not %x% + %y% (unless they're
		// deliberately double-derefs).
		cp += this_deref_ref.length;
		// The outer loop will now do ++infix for us.

continue;     // To avoid falling into the label below. The label below is only reached by explicit goto.
double_deref: // Caller has set cp to be start and op_end to be the character after the last one of the double deref.
		if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol)) // If it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
		{
			if (infix_count > MAX_TOKENS - 2) // -2 to ensure room for this operator and the operand further below.
				return LineError(ERR_EXPR_TOO_LONG);
			infix[infix_count++].symbol = SYM_CONCAT;
		}

		infix[infix_count].symbol = SYM_DYNAMIC;
		if (   !(infix[infix_count].buf = SimpleHeap::Malloc(cp, op_end - cp))   ) // Example string: "Array%i%"
			return LineError(ERR_OUTOFMEM);

		// Set "deref" properly for the loop to resume processing at the item after this double deref.
		// Callers of double_deref have ensured that deref!=NULL and deref->marker!=NULL (because it
		// doesn't make sense to have a double-deref unless caller discovered the first deref that
		// belongs to this double deref, such as the "i" in Array%i%).
		for (deref_start = deref, ++deref; deref->marker && deref->marker < op_end; ++deref);
		derefs_in_this_double = (int)(deref - deref_start);
		--deref; // Compensate for the outer loop's ++deref.

		// There's insufficient room to shoehorn all the necessary data into the token (since circuit_token probably
		// can't be safely overloaded at this stage), so allocate a little bit of stack memory, just enough for the
		// number of derefs (variables) whose contents comprise the name of this double-deref variable (typically
		// there's only one; e.g. the "i" in Array%i%).
		if (   !(deref_new = (DerefType *)SimpleHeap::Malloc((derefs_in_this_double + 1) * sizeof(DerefType)))   ) // Provides one extra at the end as a terminator.
			return LineError(ERR_OUTOFMEM);
		memcpy(deref_new, deref_start, derefs_in_this_double * sizeof(DerefType));
		deref_new[derefs_in_this_double].marker = NULL; // Put a NULL in the last item, which terminates the array.
		for (deref_start = deref_new; deref_start->marker; ++deref_start)
			deref_start->marker = infix[infix_count].buf + (deref_start->marker - cp); // Point each to its position in the *new* buf.
		infix[infix_count].var = (Var *)deref_new; // Postfix evaluation uses this to build the variable's name dynamically.

		if (*op_end == '(' // i.e. dynamic function call (v1.0.47.06)
			&& !(infix_count && infix[infix_count-1].symbol == SYM_NEW)) // Not something like "new TestClass%n%()".
		{
			if (infix_count > MAX_TOKENS - 2) // No room for the following symbol to be added (plus the ++infix done that will be done by the outer loop).
				return LineError(ERR_EXPR_TOO_LONG);
			deref_start->is_function = true; // As a result of the loop above, deref_start is the null-marker deref which terminates the deref list.
			deref_start->func = NULL; // L31: MUST BE INITIALIZED for new parameter validation method.
			deref_start->param_count = deref_new->param_count; // param_count was set when the derefs were parsed.
			++infix_count; // THIS CREATES ANOTHER TOKEN for the function call itself.  Order in infix is SYM_DYNAMIC + SYM_FUNC + (parameter tokens/operators).
			infix[infix_count].symbol = SYM_FUNC;
			infix[infix_count].deref = deref_start; // See comment below.
			// The trick here is that this SYM_FUNC now points to one of the same deref items that the
			// corresponding SYM_DYNAMIC does. During postfix evaluation, that allows SYM_DYNAMIC to update
			// the attributes of that deref so that when the SYM_FUNC is executed, it will know which function
			// to call.
		}
		else
			deref_start->is_function = false;
		cp = op_end; // Must be done only after above is done using cp: Set things up for the next iteration.
		// The outer loop will now do ++infix for us.
	} // For each deref in this expression, and also for the final literal/raw text to the right of the last deref.

	// Terminate the array with a special item.  This allows infix-to-postfix conversion to do a faster
	// traversal of the infix array.
	if (infix_count > MAX_TOKENS - 1) // No room for the following symbol to be added.
		return LineError(ERR_EXPR_TOO_LONG);
	infix[infix_count].symbol = SYM_INVALID;

	////////////////////////////
	// CONVERT INFIX TO POSTFIX.
	////////////////////////////
	// SYM_BEGIN is the first item to go on the stack.  It's a flag to indicate that conversion to postfix has begun:
	ExprTokenType token_begin;
	token_begin.symbol = SYM_BEGIN;
	STACK_PUSH(&token_begin);

	SymbolType stack_symbol, infix_symbol, sym_prev;
	ExprTokenType *fwd_infix, *this_infix = infix;
	DerefType *in_param_list = NULL; // While processing the parameter list of a function-call, this points to its deref (for parameter counting and as an indicator).

	for (;;) // While SYM_BEGIN is still on the stack, continue iterating.
	{
		ExprTokenType *&this_postfix = postfix[postfix_count]; // Resolve early, especially for use by "goto". Reduces code size a bit, though it doesn't measurably help performance.
		infix_symbol = this_infix->symbol;                     //
		stack_symbol = stack[stack_count - 1]->symbol; // Frequently used, so resolve only once to help performance.

		// Put operands into the postfix array immediately, then move on to the next infix item:
		if (IS_OPERAND(infix_symbol)) // At this stage, operands consist of only SYM_OPERAND and SYM_STRING.
		{
			if (infix_symbol == SYM_DYNAMIC && SYM_DYNAMIC_IS_VAR_NORMAL_OR_CLIP(this_infix)) // Ordered for short-circuit performance.
			{
				// v1.0.46.01: If an environment variable is being used as an lvalue -- regardless
				// of whether that variable is blank in the environment -- treat it as a normal
				// variable instead.  This is because most people would want that, and also because
				// it's traditional not to directly support assignments to environment variables
				// (only EnvSet can do that, mostly for code simplicity).  In addition, things like
				// EnvVar.="string" and EnvVar+=2 aren't supported due to obscurity/rarity (instead,
				// such expressions treat EnvVar as blank). In light of all this, convert environment
				// variables that are targets of ANY assignments into normal variables so that they
				// can be seen as a valid lvalues when the time comes to do the assignment.
				// IMPORTANT: VAR_CLIPBOARD is made into SYM_VAR here, but only for assignments.
				// This allows built-in functions and other places in the code to treat SYM_VAR
				// as though it's always VAR_NORMAL, which reduces code size and improves maintainability.
				sym_prev = this_infix[1].symbol; // Resolve to help macro's code size and performance.
				if (IS_ASSIGNMENT_OR_POST_OP(sym_prev) // Post-op must be checked for VAR_CLIPBOARD (by contrast, it seems unnecessary to check it for others; see comments below).
					|| stack_symbol == SYM_PRE_INCREMENT || stack_symbol == SYM_PRE_DECREMENT) // Stack *not* infix.
				{
					this_infix->symbol = SYM_VAR; // Convert clipboard or environment variable into SYM_VAR.
					//this_infix->is_lvalue = TRUE; // No need to set this; it will be set when the assignment is detected.
				}
				// POST-INC/DEC: It seems unnecessary to check for these except for VAR_CLIPBOARD because
				// those assignments (and indeed any assignment other than .= and :=) will have no effect
				// on a ON A SYM_DYNAMIC environment variable.  This is because by definition, such
				// variables have an empty Var::Contents(), and AutoHotkey v1 does not allow
				// math operations on blank variables.  Thus, the result of doing a math-assignment
				// operation on a blank lvalue is almost the same as doing it on an invalid lvalue.
				// The main difference is that with the exception of post-inc/dec, assignments
				// wouldn't produce an lvalue unless we explicitly check for them all above.
				// An lvalue should be produced so that the following features are consistent
				// even for variables whose names happen to match those of environment variables:
				// - Pass an assignment byref or takes its address; e.g. &(++x).
				// - Cascading assignments; e.g. (++var) += 4 (rare to be sure).
				// - Possibly other lvalue behaviors that rely on SYM_VAR being present.
				// Above logic might not be perfect because it doesn't check for parens such as (var):=x,
				// and possibly other obscure types of assignments.  However, it seems adequate given
				// the rarity of such things and also because env vars are being phased out (scripts can
				// use #NoEnv to avoid all such issues).
			}
			this_postfix = this_infix++;
			this_postfix->circuit_token = NULL; // Set default. It's only ever overridden after it's in the postfix array.
			++postfix_count;
			continue; // Doing a goto to a hypothetical "standard_postfix_circuit_token" (in lieu of these last 3 lines) reduced performance and didn't help code size.
		}

		// Since above didn't "continue", the current infix symbol is not an operand, but an operator or other symbol.

		switch(infix_symbol)
		{
		case SYM_CPAREN: // Listed first for performance. It occurs frequently while emptying the stack to search for the matching open-parenthesis.
		case SYM_CBRACKET:	// Requires similar handling to CPAREN.
		case SYM_CBRACE:	// Requires similar handling to CPAREN.
		case SYM_COMMA:		// COMMA is handled here with CPAREN/BRACKET/BRACE for parameter counting and validation.
			if (infix_symbol != SYM_COMMA && !IS_OPAREN_MATCHING_CPAREN(stack_symbol, infix_symbol))
			{
				// This stack item is not the OPAREN/BRACKET/BRACE corresponding to this CPAREN/BRACKET/BRACE.
				if (stack_symbol == SYM_BEGIN // This can happen with bad expressions like "Var := 1 ? (:) :" even though theoretically it should mean that paren is closed without having been opened (currently impossible due to load-time balancing).
					|| IS_OPAREN_LIKE(stack_symbol)) // Mismatched parens/brackets/braces.
				{
					return LineError( (infix_symbol == SYM_CPAREN) ? ERR_UNEXPECTED_CLOSE_PAREN
									: (infix_symbol == SYM_CBRACKET) ? ERR_UNEXPECTED_CLOSE_BRACKET
									: ERR_UNEXPECTED_CLOSE_BRACE );
				}
				else // This stack item is an operator.
				{
					goto standard_pop_into_postfix;
					// By not incrementing i, the loop will continue to encounter SYM_CPAREN and thus
					// continue to pop things off the stack until the corresponding OPAREN is reached.
				}
			}
			// Otherwise, this infix item is a comma or a close-paren/bracket/brace whose corresponding
			// open-paren/bracket/brace is now at the top of the stack.  If a function parameter has just
			// been completed in postfix, we have extra work to do:
			//  a) Maintain and validate the parameter count.
			//  b) Where possible, allow empty parameters by inserting the parameter's default value.
			//  c) Optimize DllCalls by pre-resolving common function names.
			if (in_param_list && IS_OPAREN_LIKE(stack_symbol))
			{
				Func *func = in_param_list->func; // Can be NULL, e.g. for dynamic function calls.
				if (infix_symbol == SYM_COMMA || this_infix[-1].symbol != stack_symbol) // i.e. not an empty parameter list.
				{
					// Ensure the function can accept this many parameters.
					if (func && in_param_list->param_count >= func->mParamCount)
					{
						if (!func->mIsVariadic)
							return LineError(ERR_TOO_MANY_PARAMS, FAIL, in_param_list->marker);
						func = NULL; // Indicate that no validation can be done for this parameter.
					}

					// Accessing this_infix[-1] here is necessarily safe since in_param_list is
					// non-NULL, and that can only be the result of a previous SYM_OPAREN/BRACKET.
					if (this_infix[-1].symbol == SYM_COMMA || this_infix[-1].symbol == stack_symbol)
					{
						int num_blank_params = 0;
						// Also handle any missing params following this one, otherwise the this_infix[-1].symbol
						// check would fail next iteration because we've changed it from SYM_COMMA to SYM_MISSING.
						while (this_infix->symbol == SYM_COMMA) // For each missing parameter: (, or ,,
						{
							postfix[postfix_count] = this_infix++;
							postfix[postfix_count]->symbol = SYM_MISSING;
							postfix[postfix_count]->marker = _T(""); // Simplify some cases by letting it be treated as SYM_STRING.
							postfix[postfix_count]->circuit_token = NULL;
							++num_blank_params;
							++postfix_count;
						}
						// Below: Detects ,) and ,] as errors.  Since the loop above doesn't handle those cases,
						// the "continue" below would cause an infinite loop.  Additionally, it's not at all
						// useful and so likely to be an error.
						if (IS_CPAREN_LIKE(this_infix->symbol) // End of the list.
							|| func && in_param_list->param_count < func->mMinParams) // Omitted a required parameter.
							return LineError(ERR_BLANK_PARAM, FAIL, in_param_list->marker);
						// Now that we no longer need its old value, update param_count:
						in_param_list->param_count += num_blank_params;
						// Go back to the top to update the this_postfix ref.
						continue;
					}

					#ifdef ENABLE_DLLCALL
					if (func && func->mBIF == &BIF_DllCall // Implies mIsBuiltIn == true.
						&& in_param_list->param_count == 0) // i.e. this is the end of the first param.
					{
						// Optimise DllCall by resolving function addresses at load-time where possible.
						ExprTokenType &param1 = this_infix[-1]; // Safety note: this_infix is necessarily preceded by at least two tokens at this stage.
						if (param1.symbol == SYM_STRING && this_infix[-2].symbol == SYM_OPAREN) // i.e. the first param is a single literal string and nothing else.
						{
							void *function = GetDllProcAddress(param1.marker);
							if (function)
							{
								param1.symbol = SYM_INTEGER;
								param1.value_int64 = (__int64)function;
							}
							// Otherwise, one of the following is true:
							//  a) A file was specified and is not already loaded.  GetDllProcAddress avoids
							//     loading any dlls since doing so may have undesired side-effects.  The file
							//     might not exist at this stage, but might when DllCall is actually called.
							//	b) The function could not be found.  This is not considered an error since the
							//     absence of certain functions (e.g. IsWow64Process) may have some meaning to
							//     the script; or the function may be non-essential.
						}
					}
					#endif

					// This is SYM_COMMA or SYM_CPAREN/BRACKET/BRACE at the end of a parameter.
					++in_param_list->param_count;

					if (stack_symbol == SYM_OBRACE && (in_param_list->param_count & 1)) // i.e. an odd number of parameters, which means no "key:" was specified.
						return LineError(_T("Missing \"key:\" in object literal."));
				}

				// Enforce mMinParams:
				if (func && infix_symbol == SYM_CPAREN && in_param_list->param_count < func->mMinParams
					&& in_param_list->is_function != DEREF_VARIADIC) // Check this last since it will probably be rare.
					return LineError(ERR_TOO_FEW_PARAMS, FAIL, in_param_list->marker);
			}
				
			switch (infix_symbol)
			{
			case SYM_CPAREN: // implies stack_symbol == SYM_OPAREN.
				// See comments near the bottom of this (outer) case.  The first open-paren on the stack must be the one that goes with this close-paren.
				--stack_count; // Remove this open-paren from the stack, since it is now complete.
				++this_infix;  // Since this pair of parentheses is done, move on to the next token in the infix expression.

				in_param_list = (DerefType *)stack[stack_count]->buf; // Restore in_param_list to the value it had when SYM_OPAREN was pushed onto the stack.

				if (stack[stack_count-1]->symbol == SYM_FUNC) // i.e. topmost item on stack is SYM_FUNC.
				{
					goto standard_pop_into_postfix; // Within the postfix list, a function-call should always immediately follow its params.
				}
				break;
				
			case SYM_CBRACKET: // implies stack_symbol == SYM_OBRACKET.
			case SYM_CBRACE: // implies stack_symbol == SYM_OBRACE.
			{
				ExprTokenType &stack_top = *stack[stack_count - 1];
				//--stack_count; // DON'T DO THIS.
				stack_top.symbol = SYM_FUNC; // Change this OBRACKET to FUNC (see below).

				if (this_infix->buf[1] == '(') // i.e. "]("
				{
					// Appears to be a method call with a computed method name, such as x[y](prms).
					ASSERT(this_infix[1].symbol == SYM_OPAREN);
					if (infix_symbol == SYM_CBRACE // i.e. {...}(), seems best to reserve this for now.
						|| in_param_list->func != &g_ObjGet // i.e. it's something like x := [y,z]().
						|| in_param_list->param_count > 2) // i.e. x[y, ...]().
						return LineError(_T("Unsupported method call syntax."), FAIL, in_param_list->marker); // Error message is a bit vague since this can be x[y,z]() or x.y[z]().
					if (in_param_list->param_count == 1) // Just the target object; no method name: x[](...)
					{
						in_param_list->param_count++;
						postfix[postfix_count] = this_infix;
						postfix[postfix_count]->symbol = SYM_MISSING;
						postfix[postfix_count]->marker = _T(""); // Simplify some cases by letting it be treated as SYM_STRING.
						postfix[postfix_count]->circuit_token = NULL;
						++postfix_count;
					}
					stack_top.deref->func = &g_ObjCall; // Override the default now that we know this is a method-call.
					++this_infix; // Skip SYM_CBRACKET so this_infix points to SYM_OPAREN.
					this_infix->buf = stack_top.buf; // This contains the old value of in_param_list.
					// Push the open-paren over stack_top (which is now SYM_FUNC) so it will be handled
					// like an ordinary function call when a comma or the close-paren is encountered.
					STACK_PUSH(this_infix++);
					// The rest of the parameter list will be handled like any other function call,
					// except that in_param_list->param_count is already non-zero.
					break;
				}

				++this_infix; // Since this pair of brackets is done, move on to the next token in the infix expression.
				in_param_list = (DerefType *)stack_top.buf; // Restore in_param_list to the value it had when '[' was pushed onto the stack.					
				goto standard_pop_into_postfix; // Pop the token (now SYM_FUNC) into the postfix array to immediately follow its params.
			}

			default: // case SYM_COMMA:
				if (sPrecedence[stack_symbol] < sPrecedence[infix_symbol]) // i.e. stack_symbol is SYM_BEGIN or SYM_OPAREN/BRACKET/BRACE.
				{
					if (!in_param_list) // This comma separates statements rather than function parameters.
					{
						STACK_PUSH(this_infix);
						// v1.0.46.01: Treat ", var = expr" as though the "=" is ":=", even if there's a ternary
						// on the right side (for consistency and since such a ternary would be stand-alone,
						// which is a rare use for ternary).  Also cascade to the right to treat things like
						// x=y=z as assignments because its intuitiveness seems to outweigh other considerations.
						for (fwd_infix = this_infix + 1;; fwd_infix += 2)
						{
							// The following is checked first to simplify things and avoid any chance of reading
							// beyond the last item in the array. This relies on the fact that a SYM_INVALID token
							// exists at the end of the array as a terminator.
							if (fwd_infix->symbol == SYM_INVALID || fwd_infix[1].symbol != SYM_EQUAL) // Relies on short-circuit boolean order.
								break; // No further checking needed because there's no qualified equal-sign.
							// Otherwise, check what lies to the left of the equal-sign.
							if (fwd_infix->symbol == SYM_VAR)
							{
								fwd_infix[1].symbol = SYM_ASSIGN;
								continue; // Cascade to the right until the last qualified '=' operator is found.
							}
							// Otherwise, it's not a pure/normal variable.  But check if it's an environment var.
							if (fwd_infix->symbol != SYM_DYNAMIC || !SYM_DYNAMIC_IS_VAR_NORMAL_OR_CLIP(fwd_infix))
								break; // It qualifies as neither SYM_DYNAMIC nor SYM_VAR.
							// Otherwise, this is an environment variable being assigned something, so treat
							// it as a normal variable rather than an environment variable. This is because
							// by tradition (and due to the fact that not many people would want it),
							// direct assignment to environment variables isn't supported by anything other
							// than EnvSet.
							fwd_infix->symbol = SYM_VAR; // Convert dynamic to a normal variable, see above.
							//fwd_infix->is_lvalue = TRUE; // No need to set this; it will be set when the assignment is detected.
							fwd_infix[1].symbol = SYM_ASSIGN;
							// And now cascade to the right until the last qualified '=' operator is found.
						}
					}
					else
					{
						// It's a function comma, so don't put it in stack because function commas aren't
						// needed and they would probably prevent proper evaluation.  Only statement-separator
						// commas need to go onto the stack.
					}
					++this_infix; // Regardless of the outcome above, move rightward to the next infix item.
				}
				else
					goto standard_pop_into_postfix;
				break;
			} // end switch (infix_symbol)
			break;

		case SYM_FUNC:
			STACK_PUSH(this_infix++);
			// NOW FALL INTO THE OPEN-PAREN BELOW because load-time validation has ensured that each SYM_FUNC
			// is followed by a '('.
// ABOVE CASE FALLS INTO BELOW.
		case SYM_OPAREN:
			// Open-parentheses always go on the stack to await their matching close-parentheses.
			this_infix->buf = (LPTSTR)in_param_list; // L31: Save current value on the stack with this SYM_OPAREN.
			if (infix_symbol == SYM_FUNC)
				in_param_list = this_infix[-1].deref	; // Store this SYM_FUNC's deref.
			else if (stack_symbol == SYM_NEW)
			{
				// Now that the SYM_OPAREN of this SYM_NEW has been found, translate it to SYM_FUNC
				// so that it will be popped off the stack immediately after its parameter list:
				stack[stack_count - 1]->symbol = SYM_FUNC;
				in_param_list = stack[stack_count - 1]->deref;
			}
			else
				in_param_list = NULL; // Allow multi-statement commas, even in cases like Func((x,y)).
			STACK_PUSH(this_infix++);
			break;
			
		case SYM_OBRACKET:
		case SYM_OBRACE:
			this_infix->buf = (LPTSTR)in_param_list; // Save current value on the stack with this SYM_OBRACKET.
			in_param_list = this_infix->deref; // This deref holds param_count and other info for the current parameter list.
			STACK_PUSH(this_infix++); // Push this '[' onto the stack to await its ']'.
			break;

		case SYM_DOT: // x.y
			this_infix->symbol = SYM_FUNC;
			STACK_PUSH(this_infix++);
			goto standard_pop_into_postfix; // Pop it into postfix to immediately follow its operands.

		case SYM_IFF_ELSE: // i.e. this infix symbol is ':'.
			if (stack_symbol == SYM_BEGIN) // An ELSE with no matching IF/THEN.
				return LineError(_T("A \":\" is missing its \"?\"")); // Below relies on the above check having been done, to avoid underflow.
			if (in_param_list && stack_symbol == SYM_OBRACE)
			{	// End of key in something like {x: y}.
				++in_param_list->param_count;
				++this_infix;
				continue;
			}
			if (stack_symbol == SYM_OPAREN)
				return LineError(_T("Missing \")\" before \":\""));
			if (stack_symbol == SYM_OBRACKET)
				return LineError(_T("Missing \"]\" before \":\""));
			if (stack_symbol == SYM_OBRACE)
				return LineError(_T("Missing \"}\" before \":\""));
			// Otherwise:
			if (stack_symbol == SYM_IFF_THEN) // See comments near the bottom of this case. The first found "THEN" on the stack must be the one that goes with this "ELSE".
			{
				this_postfix = STACK_POP; // Pop this "THEN" into the postfix array.
				this_postfix->circuit_token = this_infix; // Point this "THEN" to its "ELSE" for use by short-circuit. This simplifies short-circuit by means such as avoiding the need to take notice of nested IFF's when discarding a branch (a different stage points the IFF's condition to its "THEN").
				++postfix_count;
				STACK_PUSH(this_infix++); // Push the ELSE onto the stack so that its operands will go into the postfix array before it.
				// Above also does ++i since this ELSE found its matching IF/THEN, so it's time to move on to the next token in the infix expression.
			}
			else // This stack item is an operator INCLUDE some other THEN's ELSE (all such ELSE's should be purged from the stack so that 1 ? 1 ? 2 : 3 : 4 creates postfix 112?3:?4: not something like 112?3?4::.
			{
				// By not incrementing i, the loop will continue to encounter SYM_IFF_ELSE and thus
				// continue to pop things off the stack until the corresponding SYM_IFF_THEN is reached.
				goto standard_pop_into_postfix;
			}
			break;

		case SYM_INVALID:
			if (stack_symbol == SYM_BEGIN) // Stack is basically empty, so stop the loop.
			{
				--stack_count; // Remove SYM_BEGIN from the stack, leaving the stack empty for use in postfix eval.
				goto end_of_infix_to_postfix; // Both infix and stack have been fully processed, so the postfix expression is now completely built.
			}
			else if (stack_symbol == SYM_OPAREN) // Open paren is never closed (currently impossible due to load-time balancing, but kept for completeness).
				return LineError(ERR_MISSING_CLOSE_PAREN); // Since this error string is used in other places, compiler string pooling should result in little extra memory needed for this line.
			else if (stack_symbol == SYM_OBRACKET) // L31
				return LineError(ERR_MISSING_CLOSE_BRACKET);
			else if (stack_symbol == SYM_OBRACE)
				return LineError(ERR_MISSING_CLOSE_BRACE);
			else // Pop item off the stack, AND CONTINUE ITERATING, which will hit this line until stack is empty.
				goto standard_pop_into_postfix;
			// ALL PATHS ABOVE must continue or goto.
			
		case SYM_MULTIPLY:
			if (in_param_list && (this_infix[1].symbol == SYM_CPAREN || this_infix[1].symbol == SYM_CBRACKET)) // Func(params*) or obj.foo[params*]
			{
				in_param_list->is_function = DEREF_VARIADIC;
				++this_infix;
				continue;
			}
			// DO NOT BREAK: FALL THROUGH TO BELOW

		default: // This infix symbol is an operator, so act according to its precedence.
			if (IS_ASSIGNMENT_OR_POST_OP(infix_symbol))
			{
				// Resolve the variable now, for validation after all files have been loaded.
				// Without this, the validation code would need to determine which postfix token
				// corresponds to an assignment's l-value, which would require larger code.
				if (this_infix > infix) // Must be checked, although always true in valid expressions.
				{
					if (this_infix[-1].symbol == SYM_VAR)
						this_infix[-1].is_lvalue = TRUE;
				}
			}
			//else: if it's pre-increment/decrement, looking to the right in infix is insufficient.
			// For cases like ++this.x, we must wait until the operator is popped from the stack.

			// If the symbol waiting on the stack has a lower precedence than the current symbol, push the
			// current symbol onto the stack so that it will be processed sooner than the waiting one.
			// Otherwise, pop waiting items off the stack (by means of i not being incremented) until their
			// precedence falls below the current item's precedence, or the stack is emptied.
			// Note: BEGIN and OPAREN are the lowest precedence items ever to appear on the stack (CPAREN
			// never goes on the stack, so can't be encountered there).
			if (   sPrecedence[stack_symbol] < sPrecedence[infix_symbol] + (sPrecedence[infix_symbol] % 2) // Performance: An sPrecedence2[] array could be made in lieu of the extra add+indexing+modulo, but it benched only 0.3% faster, so the extra code size it caused didn't seem worth it.
				|| IS_ASSIGNMENT_EXCEPT_POST_AND_PRE(infix_symbol) && stack_symbol != SYM_DEREF // See note 1 below. Ordered for short-circuit performance.
				|| stack_symbol == SYM_POWER && (infix_symbol >= SYM_NEGATIVE && infix_symbol <= SYM_DEREF // See note 2 below. Check lower bound first for short-circuit performance.
					|| infix_symbol == SYM_LOWNOT)   )
			{
				// NOTE 1: v1.0.46: The IS_ASSIGNMENT_EXCEPT_POST_AND_PRE line above was added in conjunction with
				// the new assignment operators (e.g. := and +=). Here's what it does: Normally, the assignment
				// operators have the lowest precedence of all (except for commas) because things that lie
				// to the *right* of them in the infix expression should be evaluated first to be stored
				// as the assignment's result.  However, if what lies to the *left* of the assignment
				// operator isn't a valid lvalue/variable (and not even a unary like -x can produce an lvalue
				// because they're not supposed to alter the contents of the variable), obeying the normal
				// precedence rules would be produce a syntax error due to "assigning to non-lvalue".
				// So instead, apply any pending operator on the stack (which lies to the left of the lvalue
				// in the infix expression) *after* the assignment by leaving it on the stack.  For example,
				// C++ and probably other languages (but not the older ANSI C) evaluate "true ? x:=1 : y:=1"
				// as a pair of assignments rather than as who-knows-what (probably a syntax error if you
				// strictly followed precedence).  Similarly, C++ evaluates "true ? var1 : var2 := 3" not as
				// "(true ? var1 : var2) := 3" (like ANSI C) but as "true ? var1 : (var2 := 3)".  Other examples:
				// -> not var:=5 ; It's evaluated contrary to precedence as: not (var:=5) [PHP does this too,
				//    and probably others]
				// -> 5 + var+=5 ; It's evaluated contrary to precedence as: 5 + (var+=5) [not sure if other
				//    languages do ones like this]
				// -> ++i := 5 ; Silly since increment has no lasting effect; so assign the 5 then do the pre-inc.
				// -> ++i /= 5 ; Valid, but maybe too obscure and inconsistent to treat it differently than
				//    the others (especially since not many people will remember that unlike i++, ++i produces
				//    an lvalue); so divide by 5 then do the increment.
				// -> i++ := 5 (and i++ /= 5) ; Postfix operator can't produce an lvalue, so do the assignment
				//    first and then the postfix op.
				// SYM_DEREF is the only exception to the above because there's a slight chance that
				// *Var:=X (evaluated strictly according to precedence as (*Var):=X) will be used for someday.
				// Also, SYM_FUNC seems unaffected by any of this due to its enclosing parentheses (i.e. even
				// if a function-call can someday generate an lvalue [SYM_VAR], the current rules probably
				// already support it.
				// Performance: Adding the above behavior reduced the expressions benchmark by only 0.6%; so
				// it seems worth it.
				//
				// NOTE 2: The SYM_POWER line above is a workaround to allow 2**-2 (and others in v1.0.45) to be
				// evaluated as 2**(-2) rather than being seen as an error.  However, as of v1.0.46, consecutive
				// unary operators are supported via the right-to-left evaluation flag above (formerly, consecutive
				// unaries produced a failure [blank value]).  For example:
				// !!x  ; Useful to convert a blank value into a zero for use with uninitialized variables.
				// not not x  ; Same as above.
				// Other examples: !-x, -!x, !&x, -&Var, ~&Var
				// And these deref ones (which worked even before v1.0.46 by different means: giving
				// '*' a higher precedence than the other unaries): !*Var, -*Var and ~*Var
				// !x  ; Supported even if X contains a negative number, since x is recognized as an isolated operand and not something containing unary minus.
				//
				// To facilitate short-circuit boolean evaluation, right before an AND/OR/IFF is pushed onto the
				// stack, point the end of it's left branch to it.  Note that the following postfix token
				// can itself be of type AND/OR/IFF, a simple example of which is "if (true and true and true)",
				// in which the first and's parent (in an imaginary tree) is the second "and".
				// But how is it certain that this is the final operator or operand of and AND/OR/IFF's left branch?
				// Here is the explanation:
				// Everything higher precedence than the AND/OR/IFF came off the stack right before it, resulting in
				// what must be a balanced/complete sub-postfix-expression in and of itself (unless the expression
				// has a syntax error, which is caught in various places).  Because it's complete, during the
				// postfix evaluation phase, that sub-expression will result in a new operand for the stack,
				// which must then be the left side of the AND/OR/IFF because the right side immediately follows it
				// within the postfix array, which in turn is immediately followed its operator (namely AND/OR/IFF).
				// Also, the final result of an IFF's condition-branch must point to the IFF/THEN symbol itself
				// because that's the means by which the condition is merely "checked" rather than becoming an
				// operand itself.
				if (infix_symbol <= SYM_AND && infix_symbol >= SYM_IFF_THEN && postfix_count) // Check upper bound first for short-circuit performance.
					postfix[postfix_count - 1]->circuit_token = this_infix; // In the case of IFF, this points the final result of the IFF's condition to its SYM_IFF_THEN (a different stage points the THEN to its ELSE).
				STACK_PUSH(this_infix++); // Push this_infix onto the stack and move rightward to the next infix item.
			}
			else // Stack item's precedence >= infix's (if equal, left-to-right evaluation order is in effect).
				goto standard_pop_into_postfix;
		} // switch(infix_symbol)

		continue; // Avoid falling into the label below except via explicit jump.  Performance: Doing it this way rather than replacing break with continue everywhere above generates slightly smaller and slightly faster code.
standard_pop_into_postfix: // Use of a goto slightly reduces code size.
		this_postfix = STACK_POP;
		this_postfix->circuit_token = NULL; // Set default. It's only ever overridden after it's in the postfix array.
		// Additional processing for syntax sugar:
		SymbolType postfix_symbol = this_postfix->symbol;
		switch (postfix_symbol)
		{
		case SYM_FUNC:
			infix_symbol = this_infix->symbol;
			// The sections below pre-process assignments to work with objects:
			//	x.y := z	->	x "y" z (set)
			//	x[y] += z	->	x y (get in-place, assume 2 params) z (add) (set)
			//	x.y[i] /= z	->	x "y" i 3 (get in-place, n params) z (div) (set)
			if (this_postfix->deref->func == &g_ObjGet || this_postfix->deref->func == &g_ObjCall) // Allow g_ObjCall for something like x.y(i) := z, since x.y(i) can act as x.y[i] for COM objects.
			{
				if (IS_ASSIGNMENT_EXCEPT_POST_AND_PRE(infix_symbol))
				{
					if (infix_symbol != SYM_ASSIGN)
					{
						int param_count = this_postfix->deref->param_count; // Number of parameters preceding the assignment operator.
						if (param_count != 2)
						{
							// Pass the actual param count via a separate token:
							this_postfix = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
							this_postfix->symbol = SYM_INTEGER;
							this_postfix->value_int64 = param_count;
							this_postfix->circuit_token = NULL;
							postfix_count++;
							param_count = 1; // ExpandExpression should consider there to be only one param; the others should be left on the stack.
						}
						else
						{
							param_count = 0; // Omit the "param count" token to indicate the most common case: two params.
						}
						ExprTokenType *&that_postfix = postfix[postfix_count]; // In case above did postfix_count++.
						that_postfix = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
						that_postfix->symbol = SYM_FUNC;
						if (  !(that_postfix->deref = (DerefType *)SimpleHeap::Malloc(sizeof(DerefType)))  ) // Must be persistent memory, unlike that_postfix itself.
							return LineError(ERR_OUTOFMEM);
						that_postfix->deref->func = &g_ObjGetInPlace;
						that_postfix->deref->is_function = true;
						that_postfix->deref->param_count = param_count;
						that_postfix->circuit_token = NULL;
					}
					else
					{
						--postfix_count; // Discard this token; the assignment op will be converted into SYM_FUNC later.
					}
					this_infix->deref = stack[stack_count]->deref; // Mark this assignment as an object assignment for the section below.
					this_infix->deref->func = &g_ObjSet;
					this_infix->deref->param_count++;
					// Now let this_infix be processed by the next iteration.
				}
				else if (!IS_OPERAND(infix_symbol))
				{
					// Post-increment/decrement has higher precedence, so check for it first:
					if (infix_symbol == SYM_POST_INCREMENT || infix_symbol == SYM_POST_DECREMENT)
					{
						// Replace the func with BIF_ObjIncDec to perform the operation. This has
						// the same effect as the section above with x.y(z) := 1; i.e. x.y(z)++ is
						// equivalent to x.y[z]++.  This is done for consistency, simplicity and
						// because x.y(z)++ would otherwise be a useless syntax error.
						this_postfix->deref->func = (infix_symbol == SYM_POST_INCREMENT ? &g_ObjPostInc : &g_ObjPostDec);
						++this_infix; // Discard this operator.
					}
					else
					{
						stack_symbol = stack[stack_count - 1]->symbol;
						if (stack_symbol == SYM_PRE_INCREMENT || stack_symbol == SYM_PRE_DECREMENT)
						{
							// See comments in the similar section above.
							this_postfix->deref->func = (stack_symbol == SYM_PRE_INCREMENT ? &g_ObjPreInc : &g_ObjPreDec);
							--stack_count; // Discard this operator.
						}
					}
				}
				// Otherwise, IS_OPERAND(infix_symbol) == true, which should only be possible
				// if this_infix[1] is SYM_DOT.  In that case, a later iteration should apply
				// the transformations above to that operator.
			}
			break;

		case SYM_NEW: // This is probably something like "new Class", without "()", otherwise an earlier stage would've handled it.
		case SYM_REGEXMATCH: // a ~= b  ->  RegExMatch(a, b)
			this_postfix->symbol = SYM_FUNC;
			break;

		case SYM_PRE_INCREMENT:
		case SYM_PRE_DECREMENT:
			// Mark the target variable as an l-value for #Warn ClassOverwrite.
			if (postfix_count && postfix[postfix_count-1]->symbol == SYM_VAR)
				postfix[postfix_count-1]->is_lvalue = TRUE;
			//else: It could still be something valid, like SYM_IFF_ELSE in ++(whichvar ? x : y).
			break;

		default:
			if (!IS_ASSIGNMENT_EXCEPT_POST_AND_PRE(postfix_symbol))
				break;
			if (this_postfix->deref) // The section above marked this as an object assignment in an earlier iteration.
			{
				ExprTokenType *assign_op = this_postfix;
				if (postfix_symbol != SYM_ASSIGN) // e.g. += or .=
				{
					switch (postfix_symbol)
					{
					case SYM_ASSIGN_ADD:           postfix_symbol = SYM_ADD; break;
					case SYM_ASSIGN_SUBTRACT:      postfix_symbol = SYM_SUBTRACT; break;
					case SYM_ASSIGN_MULTIPLY:      postfix_symbol = SYM_MULTIPLY; break;
					case SYM_ASSIGN_DIVIDE:        postfix_symbol = SYM_DIVIDE; break;
					case SYM_ASSIGN_FLOORDIVIDE:   postfix_symbol = SYM_FLOORDIVIDE; break;
					case SYM_ASSIGN_BITOR:         postfix_symbol = SYM_BITOR; break;
					case SYM_ASSIGN_BITXOR:        postfix_symbol = SYM_BITXOR; break;
					case SYM_ASSIGN_BITAND:        postfix_symbol = SYM_BITAND; break;
					case SYM_ASSIGN_BITSHIFTLEFT:  postfix_symbol = SYM_BITSHIFTLEFT; break;
					case SYM_ASSIGN_BITSHIFTRIGHT: postfix_symbol = SYM_BITSHIFTRIGHT; break;
					case SYM_ASSIGN_CONCAT:        postfix_symbol = SYM_CONCAT; break;
					}
					// Insert the concat or math operator before the assignment:
					this_postfix = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
					this_postfix->symbol = postfix_symbol;
					this_postfix->circuit_token = NULL;
					postfix[++postfix_count] = assign_op;
				}
				assign_op->symbol = SYM_FUNC; // An earlier stage already set up the func and param_count.
			}
		}
		++postfix_count;
	} // End of loop that builds postfix array from the infix array.
end_of_infix_to_postfix:

	// Create a new postfix array and attach it to this arg of this line.
	// SAVINGS/COMPRESSION: 4 bytes per struct could be saved by making symbol into a WORD and circuit_token
	// into a WORD/index/offset.  This was tried once and it didn't affect performance or code size very much,
	// but it did increase complexity and reduce maintainability quite a bit.  If ever try to do this, avoid
	// any 8-byte members like __int64 or double in the compressed struct because that would change the default
	// alignment to 64-bit vs. 32-bit, which would keep the struct size at 16 bytes rather than allowing it to
	// fall to 12 bytes.
	if (   !(aArg.postfix = (ExprTokenType *)SimpleHeap::Malloc((postfix_count+1)*sizeof(ExprTokenType)))   ) // +1 for the terminator item added below.
		return LineError(ERR_OUTOFMEM);

	int i, j;
	for (i = 0; i < postfix_count; ++i) // Copy the postfix array in physically sorted order into the new postfix array.
	{
		ExprTokenType &new_token = aArg.postfix[i];
		new_token = *postfix[i]; // Struct copy.  This also sets circuit_token to NULL for those circuit_tokens not overridden later below.
		if (new_token.symbol == SYM_OPERAND)
		{
			if (IsPureNumeric(new_token.marker, true, false, true) != PURE_INTEGER)
				new_token.buf = NULL; // Indicate that this SYM_OPERAND token LACKS a pre-converted binary integer.
			else // Pre-convert to binary integer, which can increase performance of complex expressions by up to 20%.
			{
				if (   !(new_token.buf = (LPTSTR) SimpleHeap::Malloc(sizeof(__int64)))   )
					return LineError(ERR_OUTOFMEM);
				*(__int64 *)new_token.buf = ATOI64(new_token.marker);
			}
		}
		if (new_token.circuit_token) // Adjust each circuit_token address to be relative to the new array rather than the temp/infix array.
		{
			for (j = i + 1; postfix[j] != new_token.circuit_token; ++j); // Should always be found, and always to the right in the postfix array, so no need to check postfix_count.
			new_token.circuit_token = aArg.postfix + j;
		}
	}
	aArg.postfix[postfix_count].symbol = SYM_INVALID;  // Special item to mark the end of the array.

	return OK;
}

//-------------------------------------------------------------------------------------

// Init static vars:
Line *Line::sLog[] = {NULL};  // Initialize all the array elements.
DWORD Line::sLogTick[]; // No initialization needed.
int Line::sLogNext = 0;  // Start at the first element.

#ifdef AUTOHOTKEYSC  // Reduces code size to omit things that are unused, and helps catch bugs at compile-time.
	LPTSTR Line::sSourceFile[1]; // No init needed.
#else
	LPTSTR *Line::sSourceFile = NULL; // Init to NULL for use with realloc() and for maintainability.
	int Line::sMaxSourceFiles = 0;
#endif
	int Line::sSourceFileCount = 0; // Zero source files initially.  The main script will be the first.

LPTSTR Line::sDerefBuf = NULL;  // Buffer to hold the values of any args that need to be dereferenced.
size_t Line::sDerefBufSize = 0;
int Line::sLargeDerefBufs = 0; // Keeps track of how many large bufs exist on the call-stack, for the purpose of determining when to stop the buffer-freeing timer.
LPTSTR Line::sArgDeref[MAX_ARGS]; // No init needed.
Var *Line::sArgVar[MAX_ARGS]; // Same.


void Line::FreeDerefBufIfLarge()
{
	if (sDerefBufSize > LARGE_DEREF_BUF_SIZE)
	{
		// Freeing the buffer should be safe even if the script's current quasi-thread is in the middle
		// of executing a command, since commands are all designed to make only temporary use of the
		// deref buffer (they make copies of anything they need prior to calling MsgSleep() or anything
		// else that might pump messages and thus result in a call to us here).
		free(sDerefBuf); // The above size-check has ensured this is non-NULL.
		SET_S_DEREF_BUF(NULL, 0);
		--sLargeDerefBufs;
		if (!sLargeDerefBufs)
			KILL_DEREF_TIMER
	}
	//else leave the timer running because some other deref buffer in a recursed ExpandArgs() layer
	// is still waiting to be freed (even if it isn't, it should be harmless to keep the timer running
	// just in case, since each call to ExpandArgs() will reset/postpone the timer due to the script
	// having demonstrated that it isn't idle).
}



// Maintain a circular queue of the lines most recently executed:
#define LOG_LINE(line) \
{ \
	sLog[sLogNext] = line; \
	sLogTick[sLogNext++] = GetTickCount(); \
	if (sLogNext >= LINE_LOG_SIZE) \
		sLogNext = 0; \
}



ResultType Line::ExecUntil(ExecUntilMode aMode, ExprTokenType *aResultToken, Line **apJumpToLine)
// Start executing at "this" line, stop when aMode indicates.
// RECURSIVE: Handles all lines that involve flow-control.
// aMode can be UNTIL_RETURN, UNTIL_BLOCK_END, ONLY_ONE_LINE.
// Returns OK, FAIL, EARLY_RETURN, EARLY_EXIT, or (perhaps only indirectly) LOOP_BREAK, or LOOP_CONTINUE.
// apJumpToLine is a pointer to Line-ptr (handle), which is an output parameter.  If NULL,
// the caller is indicating it doesn't need this value, so it won't (and can't) be set by
// the called recursion layer.
{
	Line *unused_jump_to_line;
	Line *&caller_jump_to_line = apJumpToLine ? *apJumpToLine : unused_jump_to_line; // Simplifies code in other places.
	// Important to init, since most of the time it will keep this value.
	// Tells caller that no jump is required (default):
	caller_jump_to_line = NULL;

	// The benchmark improvement of having the following variables declared outside the loop rather than inside
	// is about 0.25%.  Since that is probably not even statistically significant, the only reason for declaring
	// them here is in case compilers other than MSVC++ 7.1 benefit more -- and because it's an old silly habit.
	__int64 loop_iteration;
	WIN32_FIND_DATA *loop_file;
	RegItemStruct *loop_reg_item;
	LoopReadFileStruct *loop_read_file;
	LPTSTR loop_field;

	Line *jump_to_line; // Don't use *apJumpToLine because it might not exist.
	Label *jump_to_label;  // For use with Gosub & Goto & GroupActivate.
	BOOL jumping_from_inside_function_to_outside;
	ResultType if_condition, result;
	LONG_OPERATION_INIT
	global_struct &g = *::g; // Reduces code size and may improve performance. Eclipsing ::g with local g makes compiler remind/enforce the use of the right one.

	for (Line *line = this; line != NULL;)
	{
		// If a previous command (line) had the clipboard open, perhaps because it directly accessed
		// the clipboard via Var::Contents(), we close it here for performance reasons (see notes
		// in Clipboard::Open() for details):
		CLOSE_CLIPBOARD_IF_OPEN;

		// The below must be done at least when the keybd or mouse hook is active, but is currently
		// always done since it's a very low overhead call, and has the side-benefit of making
		// the app maximally responsive when the script is busy during high BatchLines.
		// This low-overhead call achieves at least two purposes optimally:
		// 1) Keyboard and mouse lag is minimized when the hook(s) are installed, since this single
		//    Peek() is apparently enough to route all pending input to the hooks (though it's inexplicable
		//    why calling MsgSleep(-1) does not achieve this goal, since it too does a Peek().
		//    Nevertheless, that is the testing result that was obtained: the mouse cursor lagged
		//    in tight script loops even when MsgSleep(-1) or (0) was called every 10ms or so.
		// 2) The app is maximally responsive while executing with a high or infinite BatchLines.
		// 3) Hotkeys are maximally responsive.  For example, if a user has game hotkeys, using
		//    a GetTickCount() method (which very slightly improves performance by cutting back on
		//    the number of Peek() calls) would introduce up to 10ms of delay before the hotkey
		//    finally takes effect.  10ms can be significant in games, where ping (latency) itself
		//    can sometimes be only 10 or 20ms. UPDATE: It looks like PeekMessage() yields CPU time
		//    automatically, similar to a Sleep(0), when our queue has no messages.  Since this would
		//    make scripts slow to a crawl, only do the Peek() every 5ms or so (though the timer
		//    granularity is 10ms on most OSes, so that's the true interval).
		// 4) Timed subroutines are run as consistently as possible (to help with this, a check
		//    similar to the below is also done for single commands that take a long time, such
		//    as URLDownloadToFile, FileSetAttrib, etc.
		LONG_OPERATION_UPDATE

		// If interruptions are currently forbidden, it's our responsibility to check if the number
		// of lines that have been run since this quasi-thread started now indicate that
		// interruptibility should be reenabled.  But if UninterruptedLineCountMax is negative, don't
		// bother checking because this quasi-thread will stay non-interruptible until it finishes.
		// v1.0.38.04: If g.ThreadIsCritical==true, no need to check or accumulate g.UninterruptedLineCount
		// because the script is now in charge of this thread's interruptibility.
		if (!g.AllowThreadToBeInterrupted && !g.ThreadIsCritical && g_script.mUninterruptedLineCountMax > -1) // Ordered for short-circuit performance.
		{
			// To preserve backward compatibility, ExecUntil() must be the one to check
			// g.UninterruptedLineCount and update g.AllowThreadToBeInterrupted, rather than doing
			// those things on-demand in IsInterruptible().  If those checks were moved to
			// IsInterruptible(), they might compare against a different/changed value of
			// g_script.mUninterruptedLineCountMax because IsInterruptible() is called only upon demand.
			// THIS SECTION DOES NOT CHECK g.ThreadStartTime because that only needs to be
			// checked on demand by callers of IsInterruptible().
			if (g.UninterruptedLineCount > g_script.mUninterruptedLineCountMax) // See above.
				g.AllowThreadToBeInterrupted = true;
			else
				// Incrementing this unconditionally makes it a cruder measure than g.LinesPerCycle,
				// but it seems okay to be less accurate for this purpose:
				++g.UninterruptedLineCount;
		}

		// The below handles the message-loop checking regardless of whether
		// aMode is ONLY_ONE_LINE (i.e. recursed) or not (i.e. we're using
		// the for-loop to execute the script linearly):
		if ((g.LinesPerCycle > -1 && g_script.mLinesExecutedThisCycle >= g.LinesPerCycle)
			|| (g.IntervalBeforeRest > -1 && tick_now - g_script.mLastScriptRest >= (DWORD)g.IntervalBeforeRest))
			// Sleep in between batches of lines, like AutoIt, to reduce the chance that
			// a maxed CPU will interfere with time-critical apps such as games,
			// video capture, or video playback.  Note: MsgSleep() will reset
			// mLinesExecutedThisCycle for us:
			MsgSleep(10);  // Don't use INTERVAL_UNSPECIFIED, which wouldn't sleep at all if there's a msg waiting.

		// At this point, a pause may have been triggered either by the above MsgSleep()
		// or due to the action of a command (e.g. Pause, or perhaps tray menu "pause" was selected during Sleep):
		while (g.IsPaused) // An initial "if (g.IsPaused)" prior to the loop doesn't make it any faster.
			MsgSleep(INTERVAL_UNSPECIFIED);  // Must check often to periodically run timed subroutines.

		// Do these only after the above has had its opportunity to spend a significant amount
		// of time doing what it needed to do.  i.e. do these immediately before the line will actually
		// be run so that the time it takes to run will be reflected in the ListLines log.
        g_script.mCurrLine = line;  // Simplifies error reporting when we get deep into function calls.

		if (g.ListLinesIsEnabled)
			LOG_LINE(line)

#ifdef CONFIG_DEBUGGER
		if (g_Debugger.IsConnected() && line->mActionType != ACT_WHILE) // L31: PreExecLine of ACT_WHILE is now handled in PerformLoopWhile() where inspecting A_Index will yield the correct result.
			g_Debugger.PreExecLine(line);
#endif

		// Do this only after the opportunity to Sleep (above) has passed, because during
		// that sleep, a new subroutine might be launched which would likely overwrite the
		// deref buffer used for arg expansion, below:
		// Expand any dereferences contained in this line's args.
		// Note: Only one line at a time be expanded via the above function.  So be sure
		// to store any parts of a line that are needed prior to moving on to the next
		// line (e.g. control stmts such as IF and LOOP).  Also, don't expand
		// ACT_ASSIGN because a more efficient way of dereferencing may be possible
		// in that case:
		if (line->mActionType != ACT_ASSIGN && line->mActionType != ACT_WHILE && line->mActionType != ACT_THROW)
		{
			result = line->ExpandArgs(aResultToken);
			// As of v1.0.31, ExpandArgs() will also return EARLY_EXIT if a function call inside one of this
			// line's expressions did an EXIT.
			if (result != OK)
				return result; // In the case of FAIL: Abort the current subroutine, but don't terminate the app.
		}

		if (ACT_IS_IF(line->mActionType))
		{
			++g_script.mLinesExecutedThisCycle;  // If and its else count as one line for this purpose.
			if_condition = line->EvaluateCondition();
#ifdef _DEBUG  // FAIL can be returned only in DEBUG mode.
			if (if_condition == FAIL)
				return FAIL;
#endif
			if (if_condition == CONDITION_TRUE)
			{
				// Under these conditions, line->mNextLine has already been verified non-NULL by the pre-parser,
				// so this dereference is safe:
				if (line->mNextLine->mActionType == ACT_BLOCK_BEGIN)
				{
					// If this IF/ELSE has a block under it rather than just a single line, take a shortcut
					// and directly execute the block.  This avoids one recursive call to ExecUntil()
					// for each iteration, which can speed up short/fast IFs/ELSEs by as much as 25%.
					// Another benefit is conservation of stack space, especially during "else if" ladders.
					//
					// At this point, line->mNextLine->mNextLine is already verified non-NULL by the pre-parser
					// because it checks that every IF/ELSE has a line under it (ACT_BLOCK_BEGIN in this case)
					// and that every ACT_BLOCK_BEGIN has at least one line under it.
					do
						result = line->mNextLine->mNextLine->ExecUntil(UNTIL_BLOCK_END, aResultToken, &jump_to_line);
					while (jump_to_line == line->mNextLine); // The above call encountered a Goto that jumps to the "{". See ACT_BLOCK_BEGIN in ExecUntil() for details.
				}
				else
					result = line->mNextLine->ExecUntil(ONLY_ONE_LINE, aResultToken, &jump_to_line);
				if (jump_to_line == line)
					// Since this IF's ExecUntil() encountered a Goto whose target is the IF
					// itself, continue with the for-loop without moving to a different
					// line.  Also: stay in this recursion layer even if aMode == ONLY_ONE_LINE
					// because we don't want the caller handling it because then it's cleanup
					// to jump to its end-point (beyond its own and any unowned elses) won't work.
					// Example:
					// if x  <-- If this layer were to do it, its own else would be unexpectedly encountered.
					//    label1:
					//    if y  <-- We want this statement's layer to handle the goto.
					//       goto, label1
					//    else
					//       ...
					// else
					//   ...
					continue;
				if (aMode == ONLY_ONE_LINE // See below.
					|| result != OK) // i.e. FAIL, EARLY_RETURN, EARLY_EXIT, LOOP_BREAK, or LOOP_CONTINUE.
				{
					// When jump_to_line!=NULL, the above call to ExecUntil() told us to jump somewhere.
					// But if we're in ONLY_ONE_LINE mode, our caller must handle it because only it knows how
					// to extricate itself from whatever it's doing.  Additionally, if result is LOOP_CONTINUE
					// or LOOP_BREAK and jump_to_line is not NULL, each ExecUntil() or PerformLoop() recursion
					// layer must pass jump_to_line to its caller, all the way up to the target loop which
					// will then know it should either BREAK or CONTINUE.
					//
					// EARLY_RETURN can occur if this if's action was a block and that block contained a RETURN,
					// or if this if's only action is RETURN.  It can't occur if we just executed a Gosub,
					// because that Gosub would have been done from a deeper recursion layer (and executing
					// a Gosub in ONLY_ONE_LINE mode can never return EARLY_RETURN).
					//
					caller_jump_to_line = jump_to_line; // Tell the caller to handle this jump (if applicable). jump_to_line==NULL is ok.
					return result;
				}
				// Now this if-statement, including any nested if's and their else's,
				// has been fully evaluated by the recursion above.  We must jump to
				// the end of this if-statement to get to the right place for
				// execution to resume.  UPDATE: Or jump to the goto target if the
				// call to ExecUntil told us to do that instead:
				if (jump_to_line != NULL)
				{
					if (jump_to_line->mParentLine != line->mParentLine)
					{
						caller_jump_to_line = jump_to_line; // Tell the caller to handle this jump.
						return OK;
					}
					line = jump_to_line;
				}
				else // Do the normal clean-up for an IF statement:
				{
					line = line->mRelatedLine; // The preparser has ensured that this is always non-NULL.
					if (line->mActionType == ACT_ELSE)
						line = line->mRelatedLine;
						// Now line is the ELSE's "I'm finished" jump-point, which is where
						// we want to be.  If line is now NULL, it will be caught when this
						// loop iteration is ended by the "continue" stmt below.  UPDATE:
						// it can't be NULL since all scripts now end in ACT_EXIT.
					// else the IF had NO else, so we're already at the IF's "I'm finished" jump-point.
				}
			}
			else // if_condition == CONDITION_FALSE
			{
				line = line->mRelatedLine; // The preparser has ensured that this is always non-NULL.
				if (line->mActionType == ACT_ELSE) // This IF has an else.
				{
					if (line->mNextLine->mActionType == ACT_BLOCK_BEGIN)
					{
						// For comments, see the "if_condition==CONDITION_TRUE" section higher above.
						do
							result = line->mNextLine->mNextLine->ExecUntil(UNTIL_BLOCK_END, aResultToken, &jump_to_line);
						while (jump_to_line == line->mNextLine);
					}
					else
						// Preparser has ensured that every ELSE has a non-NULL next line:
						result = line->mNextLine->ExecUntil(ONLY_ONE_LINE, aResultToken, &jump_to_line);

					if (aMode == ONLY_ONE_LINE || result != OK) // See the similar section above for comments.
					{
						caller_jump_to_line = jump_to_line;
						return result;
					}
					if (jump_to_line != NULL)
					{
						if (jump_to_line->mParentLine != line->mParentLine)
						{
							caller_jump_to_line = jump_to_line; // Tell the caller to handle this jump.
							return OK;
						}
						// jump to where the called function told us to go, rather than the end of our ELSE.
						line = jump_to_line;
					}
					else // Do the normal clean-up for an ELSE statement.
						line = line->mRelatedLine;
						// Now line is the ELSE's "I'm finished" jump-point, which is where
						// we want to be.  If line is now NULL, it will be caught when this
						// loop iteration is ended by the "continue" stmt below.  UPDATE:
						// it can't be NULL since all scripts now end in ACT_EXIT.
				}
				else if (aMode == ONLY_ONE_LINE)
					// Since this IF statement has no ELSE, and since it was executed
					// in ONLY_ONE_LINE mode, the IF-ELSE statement, which counts as
					// one line for the purpose of ONLY_ONE_LINE mode, has finished:
					return OK;
				// else we're already at the IF's "I'm finished" jump-point.
			} // if_condition == CONDITION_FALSE
			continue; // Let the for-loop process the new location specified by <line>.
		} // if (ACT_IS_IF)

		// If above didn't continue, it's not an IF, so handle the other
		// flow-control types:
		switch (line->mActionType)
		{
		case ACT_GOSUB:
			// A single gosub can cause an infinite loop if misused (i.e. recursive gosubs),
			// so be sure to do this to prevent the program from hanging:
			++g_script.mLinesExecutedThisCycle;
			if (line->mRelatedLine)
			{
				jump_to_label = (Label *)line->mRelatedLine;
				jumping_from_inside_function_to_outside = (line->mAttribute == ATTR_TRUE); // ATTR_TRUE was set by loadtime routines for any ACT_GOSUB that needs it.
			}
			else
			{
				// The label is a dereference, otherwise it would have been resolved at load-time.
				// So pass "true" below because don't want to update its mRelatedLine.  This is because
				// the label should be resolved every time through the loop in case the variable that
				// contains the label changes, e.g. Gosub, %MyLabel%
				if (   !(jump_to_label = line->GetJumpTarget(true))   )
					return FAIL; // Error was already displayed by called function.
				// Below is ordered for short-circuit performance.
				jumping_from_inside_function_to_outside = g.CurrentFunc && jump_to_label->mJumpToLine->IsOutsideAnyFunctionBody();
			}

			// v1.0.48.02: When a Gosub that lies inside a function body jumps outside of the function,
			// any references to dynamic variables should resolve to globals not locals. In addition,
			// GUI commands that lie inside such an external subroutine (such as GuiControl and
			// GuiControlGet) should behave as though they are not inside the function.
			if (jumping_from_inside_function_to_outside)
			{
				g.CurrentFuncGosub = g.CurrentFunc;
				g.CurrentFunc = NULL;
			}
			result = jump_to_label->Execute();
			if (jumping_from_inside_function_to_outside)
			{
				g.CurrentFunc = g.CurrentFuncGosub;
				g.CurrentFuncGosub = NULL; // Seems more maintainable to do it here vs. when the UDF returns, but debatable which is better overall for performance.
			}

			// Must do these return conditions in this specific order:
			if (result == FAIL || result == EARLY_EXIT)
				return result;
			if (aMode == ONLY_ONE_LINE)
				// This Gosub doesn't want its caller to know that the gosub's
				// subroutine returned early:
				return (result == EARLY_RETURN) ? OK : result;
			// If the above didn't return, the subroutine finished successfully and
			// we should now continue on with the line after the Gosub:
			line = line->mNextLine;
			continue;  // Resume looping starting at the above line.  "continue" is actually slight faster than "break" in these cases.

		case ACT_GOTO:
			// A single goto can cause an infinite loop if misused, so be sure to do this to
			// prevent the program from hanging:
			++g_script.mLinesExecutedThisCycle;
			if (   !(jump_to_label = (Label *)line->mRelatedLine)   )
				// The label is a dereference, otherwise it would have been resolved at load-time.
				// So send true because we don't want to update its mRelatedLine.  This is because
				// we want to resolve the label every time through the loop in case the variable
				// that contains the label changes, e.g. Gosub, %MyLabel%
				if (   !(jump_to_label = line->GetJumpTarget(true))   )
					return FAIL; // Error was already displayed by called function.
			// Now that the Goto is certain to occur:
			g.CurrentLabel = jump_to_label; // v1.0.46.16: Support A_ThisLabel.
			// One or both of these lines can be NULL.  But the preparser should have
			// ensured that all we need to do is a simple compare to determine
			// whether this Goto should be handled by this layer or its caller
			// (i.e. if this Goto's target is not in our nesting level, it MUST be the
			// caller's responsibility to either handle it or pass it on to its
			// caller).
			if (aMode == ONLY_ONE_LINE || line->mParentLine != jump_to_label->mJumpToLine->mParentLine)
			{
				caller_jump_to_line = jump_to_label->mJumpToLine; // Tell the caller to handle this jump.
				return OK;
			}
			// Otherwise, we will handle this Goto since it's in our nesting layer:
			line = jump_to_label->mJumpToLine;
			continue;  // Resume looping starting at the above line.  "continue" is actually slight faster than "break" in these cases.

		case ACT_GROUPACTIVATE: // Similar to ACT_GOSUB, which is why this section is here rather than in Perform().
		{
			++g_script.mLinesExecutedThisCycle; // Always increment for GroupActivate.
			WinGroup *group;
			if (   !(group = (WinGroup *)mAttribute)   )
				group = g_script.FindGroup(ARG1);
			result = OK; // Set default.
			ResultType activate_result = FAIL;
			if (group)
			{
				// Note: This will take care of DoWinDelay if needed:
				activate_result = group->Activate(*ARG2 && !_tcsicmp(ARG2, _T("R")), NULL, &jump_to_label);
				if (jump_to_label)
				{
					if (!line->IsJumpValid(*jump_to_label)) // Should be checked here rather than at the time that GroupAdd specified the label because it's from HERE that the jump will actually be done.
						return FAIL;

					// The section below is just like the Gosub code above, so maintain them together.
					jumping_from_inside_function_to_outside = g.CurrentFunc && jump_to_label->mJumpToLine->IsOutsideAnyFunctionBody();
					if (jumping_from_inside_function_to_outside)
					{
						g.CurrentFuncGosub = g.CurrentFunc;
						g.CurrentFunc = NULL;
					}
					result = jump_to_label->Execute();
					if (jumping_from_inside_function_to_outside)
					{
						g.CurrentFunc = g.CurrentFuncGosub;
						g.CurrentFuncGosub = NULL; // Seems more maintainable to do it here vs. when the UDF returns, but debatable which is better overall for performance.
					}
					if (result == FAIL || result == EARLY_EXIT)
						return result;
				}
			}
			//else no such group, so just proceed.
			SetErrorLevelOrThrowBool(!activate_result);
			if (aMode == ONLY_ONE_LINE)  // v1.0.45: These two lines were moved here from above to provide proper handling for GroupActivate that lacks a jump/gosub and that lies directly beneath an IF or ELSE.
				return (result == EARLY_RETURN) ? OK : result;
			line = line->mNextLine;
			continue;  // Resume looping starting at the above line.  "continue" is actually slight faster than "break" in these cases.
		}

		case ACT_RETURN:
			// Although a return is really just a kind of block-end, keep it separate
			// because when a return is encountered inside a block, it has a double function:
			// to first break out of all enclosing blocks and then return from the gosub.
			// NOTE: The return's ARG1 expression has been evaluated by ExpandArgs() above,
			// which is desirable *even* if aResultToken is NULL (i.e. the caller will be
			// ignoring the return value) in case the return's expression calls a function
			// which has side-effects.  For example, "return LogThisEvent()".
			if (aResultToken && aResultToken->symbol == SYM_STRING) // L31: Caller wants the return value, but no result has been set since caller set this default. (ExpandExpression does not use aResultToken for string values.)
			{
				if (ARGVAR1 && ARGVAR1->HasObject())
				{
					// This is a plain variable reference (not an expression) and the variable
					// contains an object.
					ARGVAR1->ToToken(*aResultToken);
				}
				else
				{
					// Even if this is a var containing a cached binary number, it also contains
					// a string which may have special formatting.  (It is certain that any var
					// at this point already contains a string, due to ExpandArgs() being called.)
					// So for compatibility and generally intuitive behaviour, return the string.
					aResultToken->symbol = SYM_STRING;
					aResultToken->marker = ARG1; // This sets it to blank if this return lacks an arg.
				}
			}
			//else the return value, if any, is discarded.
			// Don't count returns against the total since they should be nearly instantaneous. UPDATE: even if
			// the return called a function (e.g. return fn()), that function's lines would have been added
			// to the total, so there doesn't seem much problem with not doing it here.
			//++g_script.mLinesExecutedThisCycle;
			if (aMode != UNTIL_RETURN)
				// Tells the caller to return early if it's not the Gosub that directly
				// brought us into this subroutine.  i.e. it allows us to escape from
				// any number of nested blocks in order to get back out of their
				// recursive layers and back to the place this RETURN has meaning
				// to someone (at the right recursion layer):
				return EARLY_RETURN;
			return OK;

		case ACT_BREAK:
			if (line->mRelatedLine)
			{
				// Rather than having PerformLoop() handle LOOP_BREAK specifically, tell our caller to jump to
				// the line *after* the loop's body. This is always a jump our caller must handle, unlike GOTO:
				caller_jump_to_line = line->mRelatedLine;
				// Return OK instead of LOOP_BREAK to handle it like GOTO.  Returning LOOP_BREAK would
				// cause the following to behave incorrectly (as described in the comments):
				// If (True)  ; LOOP_BREAK takes precedence, causing this ExecUntil() layer to return.
				//     Loop, 1  ; Non-NULL jump-to-line causes the LOOP_BREAK result to propagate.
				//         Loop, 1
				//             Break, 2
				// MsgBox, This message would not be displayed.
				return OK;
			}
			return LOOP_BREAK;

		case ACT_CONTINUE:
			if (line->mRelatedLine)
			{
				// Signal any loops nested between this line and the target loop to return LOOP_CONTINUE:
				caller_jump_to_line = line->mRelatedLine; // Okay even if NULL.
			}
			return LOOP_CONTINUE;

		case ACT_LOOP:
		case ACT_WHILE: // Lexikos: mAttribute should be ATTR_LOOP_WHILE.
		case ACT_FOR: // Lexikos: mAttribute should be ATTR_LOOP_FOR.
		{
			HKEY root_key_type; // For registry loops, this holds the type of root key, independent of whether it is local or remote.
			AttributeType attr = line->mAttribute;
			switch ((size_t)attr)
			{
			case (size_t)ATTR_LOOP_REG:
				root_key_type = RegConvertRootKey(ARG1);
				break;
			case (size_t)ATTR_LOOP_NEW_REG:
				root_key_type = RegConvertKey(ARG2, REG_NEW_SYNTAX); // ARG1 is the word "Reg".
				break;
			case (size_t)ATTR_LOOP_UNKNOWN:
				// Since it couldn't be determined at load-time (probably due to derefs),
				// determine whether it's a file-loop, registry-loop or a normal/counter loop.
				// But don't change the value of line->mAttribute because that's our
				// indicator of whether this needs to be evaluated every time for
				// this particular loop (since the nature of the loop can change if the
				// contents of the variables dereferenced for this line change during runtime):
				switch (line->mArgc)
				{
				case 0:
					attr = ATTR_LOOP_NORMAL;
					break;
				case 1:
					// Unlike at loadtime, allow it to be negative at runtime in case it was a variable
					// reference that resolved to a negative number, to indicate that 0 iterations
					// should be performed.  UPDATE: Also allow floating point numbers at runtime
					// but not at load-time (since it doesn't make sense to have a literal floating
					// point number as the iteration count, but a variable containing a pure float
					// should be allowed):
					if (IsPureNumeric(ARG1, true, true, true))
						attr = ATTR_LOOP_NORMAL;
					else
					{
						root_key_type = RegConvertRootKey(ARG1);
						attr = root_key_type ? ATTR_LOOP_REG : ATTR_LOOP_FILEPATTERN;
					}
					break;
				default: // 2 or more args.
					if (!_tcsicmp(ARG1, _T("Read")))
						attr = ATTR_LOOP_READ_FILE;
					// Note that a "Parse" loop is not allowed to have it's first param be a variable reference
					// that resolves to the word "Parse" at runtime.  This is because the input variable would not
					// have been resolved in this case (since the type of loop was unknown at load-time),
					// and it would be complicated to have to add code for that, especially since there's
					// virtually no conceivable use for allowing it be a variable reference.
					else
					{
						root_key_type = RegConvertRootKey(ARG1);
						attr = root_key_type ? ATTR_LOOP_REG : ATTR_LOOP_FILEPATTERN;
					}
				}
			}

			// HANDLE ANY ERROR CONDITIONS THAT CAN ABORT THE LOOP:
			FileLoopModeType file_loop_mode;
			bool recurse_subfolders;
			switch ((size_t)attr)
			{
			case (size_t)ATTR_LOOP_FILEPATTERN:
				// Loop, FilePattern [, IncludeFolders?, Recurse?]
				file_loop_mode = ConvertLoopMode(ARG2);
				if (file_loop_mode == FILE_LOOP_INVALID)
					return line->LineError(ERR_PARAM2_INVALID, FAIL, ARG2);
				recurse_subfolders = (*ARG3 == '1' && !*(ARG3 + 1));
				break;
			case (size_t)ATTR_LOOP_REG:
			case (size_t)ATTR_LOOP_NEW_REG:
			case (size_t)ATTR_LOOP_NEW_FILES:
				if (attr == ATTR_LOOP_REG)
				{
					// Loop, RootKey [, Key, IncludeSubkeys?, Recurse?]
					file_loop_mode = ConvertLoopMode(ARG3);
					recurse_subfolders = (*ARG4 == '1' && !*(ARG4 + 1));
				}
				else
				{
					// Loop, Reg, RootKey\Key [, Mode]
					// Loop, Files, Pattern [, Mode]
					file_loop_mode = ConvertLoopModeString(ARG3);
					if (recurse_subfolders = (file_loop_mode & FILE_LOOP_RECURSE))
						file_loop_mode &= ~FILE_LOOP_RECURSE; // Eliminate the flag from further consideration.
				}
				if (file_loop_mode == FILE_LOOP_INVALID)
					return line->LineError(ERR_PARAM3_INVALID, FAIL, ARG3);
			}

			// ONLY AFTER THE ABOVE IS IT CERTAIN THE LOOP WILL LAUNCH (i.e. there was no error or early return).
			// So only now is it safe to make changes to global things like g.mLoopIteration.
			bool continue_main_loop = false; // Init these output parameters prior to starting each type of loop.
			jump_to_line = NULL;             //

			Line *finished_line = line->mRelatedLine;
			Line *until;
			if (finished_line->mActionType == ACT_UNTIL)
			{	// This loop has an additional post-condition.
				until = finished_line;
				// When finished, we'll jump to the line after UNTIL:
				finished_line = finished_line->mNextLine;
			}
			else
				until = NULL;

			// IN CASE THERE'S AN OUTER LOOP ENCLOSING THIS ONE, BACK UP THE A_LOOPXXX VARIABLES.
			// (See the "restore" section further below for comments.)
			loop_iteration = g.mLoopIteration;
			loop_file = g.mLoopFile;
			loop_reg_item = g.mLoopRegItem;
			loop_read_file = g.mLoopReadFile;
			loop_field = g.mLoopField;

			// INIT "A_INDEX" (one-based not zero-based). This is done here rather than in each PerformLoop()
			// function because it reduces code size and also because registry loops and file-pattern loops
			// can be intrinsically recursive (this is also related to the loop-recursion bugfix documented
			// for v1.0.20: fixes A_Index so that it doesn't wrongly reset to 0 inside recursive file-loops
			// and registry loops).
			if (attr != ATTR_LOOP_FOR) // PerformLoopFor() sets it later so its enumerator expression (which is evaluated only once) can refer to the A_Index of the outer loop.
				g.mLoopIteration = 1;

			// PERFORM THE LOOP:
			switch ((size_t)attr)
			{
			case (size_t)ATTR_LOOP_NORMAL: // Listed first for performance.
				bool is_infinite; // "is_infinite" is more maintainable and future-proof than using LLONG_MAX to simulate an infinite loop. Plus it gives peace-of-mind and the LLONG_MAX method doesn't measurably improve benchmarks (nor does BOOL vs. bool).
				__int64 iteration_limit;
				if (line->mArgc > 0) // At least one parameter is present.
				{
					// Note that a 0 means infinite in AutoIt2 for the REPEAT command; so the following handles
					// that too.
					iteration_limit = line->ArgToInt64(1); // If it's negative, zero iterations will be performed automatically.
					is_infinite = false;
				}
				else // It's ACT_LOOP without parameters.
				{
					iteration_limit = 0; // Avoids debug-mode's "used without having been defined" (though it's merely passed as a parameter, not ever used in this case).
					is_infinite = true;  // Override the default set earlier.
				}
				result = line->PerformLoop(aResultToken, continue_main_loop, jump_to_line, until
					, iteration_limit, is_infinite);
				break;
			case (size_t)ATTR_LOOP_WHILE: // Lexikos: ATTR_LOOP_WHILE is used to differentiate ACT_WHILE from ACT_LOOP, allowing code to be shared.
				result = line->PerformLoopWhile(aResultToken, continue_main_loop, jump_to_line);
				break;
			case (size_t)ATTR_LOOP_FOR:
				result = line->PerformLoopFor(aResultToken, continue_main_loop, jump_to_line, until);
				break;
			case (size_t)ATTR_LOOP_PARSE:
				// The phrase "csv" is unique enough since user can always rearrange the letters
				// to do a literal parse using C, S, and V as delimiters:
				if (_tcsicmp(ARG3, _T("CSV")))
					result = line->PerformLoopParse(aResultToken, continue_main_loop, jump_to_line, until);
				else
					result = line->PerformLoopParseCSV(aResultToken, continue_main_loop, jump_to_line, until);
				break;
			case (size_t)ATTR_LOOP_READ_FILE:
				{
					TextFile tfile;
					if (*ARG2 && tfile.Open(ARG2, DEFAULT_READ_FLAGS, g.Encoding & CP_AHKCP)) // v1.0.47: Added check for "" to avoid debug-assertion failure while in debug mode (maybe it's bad to to open file "" in release mode too).
					{
						result = line->PerformLoopReadFile(aResultToken, continue_main_loop, jump_to_line, until
							, &tfile, ARG3);
						tfile.Close();
					}
					else
						// The open of a the input file failed.  So just set result to OK since setting the
						// ErrorLevel isn't supported with loops (since that seems like it would be an overuse
						// of ErrorLevel, perhaps changing its value too often when the user would want
						// it saved -- in any case, changing that now might break existing scripts).
						result = OK;
				}
				break;
			case (size_t)ATTR_LOOP_FILEPATTERN:
			case (size_t)ATTR_LOOP_NEW_FILES:
				result = line->PerformLoopFilePattern(aResultToken, continue_main_loop, jump_to_line, until
					, file_loop_mode, recurse_subfolders, attr == ATTR_LOOP_FILEPATTERN ? ARG1 : ARG2);
				break;
			case (size_t)ATTR_LOOP_REG:
			case (size_t)ATTR_LOOP_NEW_REG:
				// This isn't the most efficient way to do things (e.g. the repeated calls to
				// RegConvertRootKey()), but it the simplest way for now.  Optimization can
				// be done at a later time:
				bool is_remote_registry;
				HKEY root_key;
				LPTSTR subkey;
				if (attr == ATTR_LOOP_REG)
					root_key = RegConvertRootKey(ARG1, &is_remote_registry), subkey = ARG2; // This will open the key if it's remote.
				else
					root_key = RegConvertKey(ARG2, REG_NEW_SYNTAX, &subkey, &is_remote_registry);
				if (root_key) 
				{
					// root_key_type needs to be passed in order to support A_LoopRegKey:
					result = line->PerformLoopReg(aResultToken, continue_main_loop, jump_to_line, until
						, file_loop_mode, recurse_subfolders, root_key_type, root_key, subkey);
					if (is_remote_registry)
						RegCloseKey(root_key);
				}
				else
					// The open of a remote key failed (we know it's remote otherwise it should have
					// failed earlier rather than here).  So just set result to OK since no ErrorLevel
					// setting is supported with loops (since that seems like it would be an overuse
					// of ErrorLevel, perhaps changing its value too often when the user would want
					// it saved.  But in any case, changing that now might break existing scripts).
					result = OK;
				break;
			}

			// RESTORE THE PREVIOUS A_LOOPXXX VARIABLES.  If there isn't an outer loop, this will set them
			// all to NULL/0, which is the most proper and also in keeping with historical behavior.
			// This backup/restore approach was adopted in v1.0.44.14 to simplify things and improve maintainability.
			// This change improved performance by only 1%, which isn't statistically significant.  More importantly,
			// it indirectly fixed the following bug:
			// When a "return" is executed inside a loop's body (and possibly an Exit or Fail too, but those are
			// covered more for simplicity and maintainability), the situations below require superglobals like
			// A_Index and A_LoopField to be restored for the outermost caller of ExecUntil():
			// 1) Var%A_Index% := func_that_returns_directly_from_inside_a_loop_body().
			//    The above happened because the return in the function's loop failed to restore A_Index for its
			//    caller because it had been designed to restore inter-line, not for intra-line activities like
			//    calling functions.
			// 2) A command that has expressions in two separate parameters and one of those parameters calls
			//    a function that returns directly from inside one of its loop bodies.
			//
			// This change was made feasible by making the A_LoopXXX attributes thread-specific, which prevents
			// interrupting threads from affecting the values our thread sees here.  So that change protects
			// against thread interruptions, and this backup/restore change here keeps the Loop variables in
			// sync with the current nesting level (braces, gosub, etc.)
			// 
			// The memory for structs like g.mLoopFile resides in the stack of an instance of PerformLoop(),
			// which is our caller or our caller's caller, etc.  In other words, it shouldn't be possible for
			// variables like g.mLoopFile to be non-NULL if there isn't a PerformLoop() beneath us in the stack.
			g.mLoopIteration = loop_iteration;
			g.mLoopFile = loop_file;
			g.mLoopRegItem = loop_reg_item;
			g.mLoopReadFile = loop_read_file;
			g.mLoopField = loop_field;

			if (result == FAIL || result == EARLY_RETURN || result == EARLY_EXIT)
				return result;
			// else result can be LOOP_BREAK or OK or LOOP_CONTINUE (but only if a loop-label was given).
			if (continue_main_loop) // It signaled us to do this:
				continue;

			if (jump_to_line)
			{
				if (jump_to_line->mParentLine != line->mParentLine)
				{
					// Our caller must handle the jump if it doesn't share the same parent as the
					// current line (i.e. it's not at the same nesting level) because that means
					// the jump target is at a more shallow nesting level than where we are now:
					caller_jump_to_line = jump_to_line; // Tell the caller to handle this jump (if applicable).
					return result; // If LOOP_CONTINUE, must be passed along so the target loop knows what to do.
				}
				// Since above didn't return, we're supposed to handle this jump.  So jump and then
				// continue execution from there:
				line = jump_to_line;
				continue; // end this case of the switch().
			}
			if (aMode == ONLY_ONE_LINE)
				return OK;
			// Since the above didn't return or break, either the loop has completed the specified
			// number of iterations or it was broken via the break command.  In either case, we jump
			// to the line after our loop's structure and continue there:
			line = finished_line;
			continue;  // Resume looping starting at the above line.  "continue" is actually slightly faster than "break" in these cases.
		} // case ACT_LOOP.

		case ACT_TRY:
		case ACT_CATCH:
		{
			ActionTypeType this_act = line->mActionType;
			int outer_excptmode = g.ExcptMode;

			if (this_act == ACT_CATCH)
			{
				// The following should never happen:
				//if (!g.ThrownToken)
				//	return line->LineError(_T("Attempt to catch nothing!"), CRITICAL_ERROR);

				Var* catch_var = ARGVARRAW1;
				ExprTokenType *our_token = g.ThrownToken;
				g.ThrownToken = NULL; // Assign() may cause script to execute via __Delete, so this must be cleared first.

				// Assign the thrown token to the variable if provided.
				result = catch_var ? catch_var->Assign(*our_token) : OK;
				g_script.FreeExceptionToken(our_token);
				if (!result)
					return FAIL;
			}
			else // (this_act == ACT_TRY)
			{
				g.ExcptMode |= EXCPTMODE_TRY; // Must use |= rather than = to avoid removing EXCPTMODE_CATCH, if present.
				if (line->mRelatedLine->mActionType != ACT_FINALLY) // Try without Catch/Finally acts like it has an empty Catch.
					g.ExcptMode |= EXCPTMODE_CATCH;
			}

			// The following section is similar to ACT_IF.
			jump_to_line = NULL;
			if (line->mNextLine->mActionType == ACT_BLOCK_BEGIN)
			{
				do
					result = line->mNextLine->mNextLine->ExecUntil(UNTIL_BLOCK_END, aResultToken, &jump_to_line);
				while (jump_to_line == line->mNextLine); // The above call encountered a Goto that jumps to the "{". See ACT_BLOCK_BEGIN in ExecUntil() for details.
			}
			else
				result = line->mNextLine->ExecUntil(ONLY_ONE_LINE, aResultToken, &jump_to_line);

			// Move to the next line after the 'try' or 'catch' block.
			line = line->mRelatedLine;

			if (this_act == ACT_TRY)
			{
				g.ExcptMode = outer_excptmode;
				bool bHasCatch = false;

				if (line->mActionType == ACT_CATCH)
				{
					bHasCatch = true;
					if (g.ThrownToken)
					{
						// An exception was thrown and we have a 'catch' block, so let the next
						// iteration handle it.  Implies result == FAIL && jump_to_line == NULL,
						// but result won't have any meaning for the next iteration.
						continue;
					}
					// Otherwise: no exception was thrown, so skip the 'catch' block.
					line = line->mRelatedLine;
				}
				if (line->mActionType == ACT_FINALLY)
				{
					// Let the section below handle the FINALLY block.
					this_act = ACT_CATCH;
				}
				else if (!bHasCatch && g.ThrownToken)
				{
					// An exception was thrown, but no 'catch' nor 'finally' is present.
					// In this case 'try' acts as a catch-all.
					g_script.FreeExceptionToken(g.ThrownToken);
					result = OK;
				}
			}
			if (this_act == ACT_CATCH && line->mActionType == ACT_FINALLY)
			{
				if (result == OK && !jump_to_line)
					// Execution is being allowed to flow normally into the finally block and
					// then the line after.  Let the next iteration handle the finally block.
					continue;

				// One of the following occurred:
				//  - An exception was thrown, and this try..(catch)..finally block didn't handle it.
				//  - A control flow statement such as break, continue or goto was used.
				// Recursively execute the finally block before continuing.
				ExprTokenType *thrown_token = g.ThrownToken;
				g.ThrownToken = NULL; // Must clear this temporarily to avoid arbitrarily re-throwing it.
				Line *invalid_jump; // Don't overwrite jump_to_line in case the try block used goto.
				PRIVATIZE_S_DEREF_BUF; // In case return was used and is returning the contents of the deref buf.
				ResultType res = line->ExecUntil(ONLY_ONE_LINE, NULL, &invalid_jump);
				DEPRIVATIZE_S_DEREF_BUF;
				if (res != OK || invalid_jump)
				{
					if (thrown_token) // The new error takes precedence over this one.
						g_script.FreeExceptionToken(thrown_token);
					if (res == FAIL || res == EARLY_EXIT)
						// Above: It's borderline whether Exit should be valid here, but it's allowed for
						// two reasons: 1) if the script was non-#Persistent it would have already terminated
						// anyway, and 2) it's only a question of whether to show a message before exiting.
						return res;
					// The remaining cases are all invalid jumps/control flow statements.  All such cases
					// should be detected at load time, but it seems best to keep this for maintainability:
					return g_script.mCurrLine->LineError(ERR_BAD_JUMP_INSIDE_FINALLY);
				}
				g.ThrownToken = thrown_token; // If non-NULL, this was thrown within the try block.
			}
			
			if (aMode == ONLY_ONE_LINE || result != OK)
			{
				caller_jump_to_line = jump_to_line;
				return result;
			}

			if (jump_to_line != NULL) // Implies goto or break/continue loop_name was used.
			{
				if (jump_to_line->mParentLine != line->mParentLine)
				{
					caller_jump_to_line = jump_to_line;
					return OK;
				}
				line = jump_to_line;
			}

			continue;
		}

		case ACT_THROW:
		{
			if (!line->mArgc)
				return line->ThrowRuntimeException(ERR_EXCEPTION);

			// ThrownToken should only be non-NULL while control is being passed up the
			// stack, which implies no script code can be executing.
			ASSERT(!g.ThrownToken);

			ExprTokenType* token = new ExprTokenType;
			if (!token) // Unlikely.
				return line->LineError(ERR_OUTOFMEM);

			// The following is based on code from PerformLoopFor()

			if (!sDerefBuf)
			{
				sDerefBufSize = (line->mArg[0].length < MAX_NUMBER_LENGTH ? MAX_NUMBER_LENGTH : line->mArg[0].length) + 1;
				if ( !(sDerefBuf = tmalloc(sDerefBufSize)) )
				{
					sDerefBufSize = 0;
					delete token;
					return line->LineError(ERR_OUTOFMEM);
				}
			}

			PRIVATIZE_S_DEREF_BUF;
			LPTSTR our_buf_marker = our_deref_buf;
			LPTSTR arg_deref[] = {0, 0};
			LPTSTR strVal;
			token->symbol = SYM_INVALID;
			strVal = line->ExpandExpression(0, result, token, our_buf_marker, our_deref_buf, our_deref_buf_size, arg_deref, 0);
			if (strVal == our_deref_buf)
				token->mem_to_free = strVal;
			else
			{
				token->mem_to_free = NULL;
				DEPRIVATIZE_S_DEREF_BUF;
			}

			if (!strVal)
			{
				// A script-function-call inside the expression returned EARLY_EXIT or FAIL.
				delete token;
				return result;
			}

			// Check if ExpandExpression has not returned a token at all
			if (token->symbol == SYM_INVALID)
			{
				// Store the returned string in the token
				token->symbol = SYM_STRING;
				token->marker = strVal;
			}

			// Throw the newly-created token
			g.ThrownToken = token;
			if (!(g.ExcptMode & EXCPTMODE_CATCH))
				g_script.UnhandledException(line);
			return FAIL;
		}

		case ACT_FINALLY:
		{
			// Directly execute next line
			line = line->mNextLine;
			continue;
		}

		case ACT_EXIT:
			// If this script has no hotkeys and hasn't activated one of the hooks, EXIT will cause the
			// the program itself to terminate.  Otherwise, it causes us to return from all blocks
			// and Gosubs (i.e. all the way out of the current subroutine, which was usually triggered
			// by a hotkey):
			if (IS_PERSISTENT)
				return EARLY_EXIT;  // It's "early" because only the very end of the script is the "normal" exit.
				// EARLY_EXIT needs to be distinct from FAIL for ExitApp() and AutoExecSection().
			// Otherwise, FALL THROUGH TO BELOW:
		case ACT_EXITAPP: // Unconditional exit.
			// This has been tested and it does yield to the OS the error code indicated in ARG1,
			// if present (otherwise it returns 0, naturally) as expected:
			return g_script.ExitApp(EXIT_EXIT, (int)line->ArgIndexToInt64(0));

		case ACT_BLOCK_BEGIN:
			if (line->mAttribute == ATTR_TRUE) // This is the ACT_BLOCK_BEGIN that starts a function's body.
			{
				// Anytime this happens at runtime it means a function has been defined inside the
				// auto-execute section, a block, or other place the flow of execution can reach
				// on its own.  This is not considered a call to the function.  Instead, the entire
				// body is just skipped over using this high performance method.  However, the function's
				// opening brace will show up in ListLines, but that seems preferable to the performance
				// overhead of explicitly removing it here.
				line = line->mRelatedLine; // Resume execution at the line following this functions end-block.
				continue;  // Resume looping starting at the above line.  "continue" is actually slight faster than "break" in these cases.
			}
			// Don't count block-begin/end against the total since they should be nearly instantaneous:
			//++g_script.mLinesExecutedThisCycle;
			// In this case, line->mNextLine is already verified non-NULL by the pre-parser:
			result = line->mNextLine->ExecUntil(UNTIL_BLOCK_END, aResultToken, &jump_to_line);
			if (jump_to_line == line)
				// Since this Block-begin's ExecUntil() encountered a Goto whose target is the
				// block-begin itself, continue with the for-loop without moving to a different
				// line.  Also: stay in this recursion layer even if aMode == ONLY_ONE_LINE
				// because we don't want the caller handling it because then it's cleanup
				// to jump to its end-point (beyond its own and any unowned elses) won't work.
				// Example:
				// if x  <-- If this layer were to do it, its own else would be unexpectedly encountered.
				// label1:
				// { <-- We want this statement's layer to handle the goto.
				//    if y
				//       goto, label1
				//    else
				//       ...
				// }
				// else
				//   ...
				continue;
			if (aMode == ONLY_ONE_LINE
				|| result != OK) // i.e. FAIL, EARLY_RETURN, EARLY_EXIT, LOOP_BREAK, or LOOP_CONTINUE.
			{
				// For more detailed comments, see the section (above this switch structure) which handles IF.
				caller_jump_to_line = jump_to_line; // Tell the caller to handle this jump (if applicable).  jump_to_line==NULL is ok.
				return result;
			}
			// Currently, all blocks are normally executed in ONLY_ONE_LINE mode because
			// they are the direct actions of an IF, an ELSE, or a LOOP.  So the
			// above will already have returned except when the user has created a
			// generic, standalone block with no associated control statement.
			// Check to see if we need to jump somewhere:
			if (jump_to_line != NULL)
			{
				if (line->mParentLine != jump_to_line->mParentLine)
				{
					caller_jump_to_line = jump_to_line; // Tell the caller to handle this jump (if applicable).
					return OK;
				}
				// Since above didn't return, jump to where the caller told us to go, rather than the end of
				// our block.
				line = jump_to_line;
			}
			else // Just go to the end of our block and continue from there.
				line = line->mRelatedLine;
				// Now line is the line after the end of this block.  Can be NULL (end of script).
				// UPDATE: It can't be NULL (not that it matters in this case) since the loader
				// has ensured that all scripts now end in an ACT_EXIT.
			continue;  // Resume looping starting at the above line.  "continue" is actually slightly faster than "break" in these cases.

		case ACT_BLOCK_END:
			// Don't count block-begin/end against the total since they should be nearly instantaneous:
			//++g_script.mLinesExecutedThisCycle;
			if (aMode != UNTIL_BLOCK_END)
				// Rajat found a way for this to happen that basically amounts to this:
				// If within a loop you gosub a label that is also inside of the block, and
				// that label sometimes doesn't return (i.e. due to a missing "return" somewhere
				// in its flow of control), the loop(s)'s block-end symbols will be encountered
				// by the subroutine, and these symbols don't have meaning to it.  In other words,
				// the subroutine has put us into a waiting-for-return state rather than a
				// waiting-for-block-end state, so when block-end's are encountered, that is
				// considered a runtime error:
				return line->LineError(_T("A \"return\" must be encountered prior to this \"}\"."));  // Former error msg was "Unexpected end-of-block (Gosub without Return?)."
			return OK; // It's the caller's responsibility to resume execution at the next line, if appropriate.

		// ACT_ELSE can happen when one of the cases in this switch failed to properly handle
		// aMode == ONLY_ONE_LINE.  But even if ever happens, it will just drop into the default
		// case, which will result in a FAIL (silent exit of thread) as an indicator of the problem.
		// So it's commented out:
		//case ACT_ELSE:
		//	// Shouldn't happen if the pre-parser and this function are designed properly?
		//	return line->LineError("Unexpected ELSE.");

		default:
			++g_script.mLinesExecutedThisCycle;
			result = line->Perform();
			if (!result || aMode == ONLY_ONE_LINE)
				// Thus, Perform() should be designed to only return FAIL if it's an error that would make
				// it unsafe to proceed in the subroutine we're executing now:
				return result; // Can be either OK or FAIL.
			line = line->mNextLine;
		} // switch()
	} // for each line

	// Above loop ended because the end of the script was reached.
	// At this point, it should be impossible for aMode to be
	// UNTIL_BLOCK_END because that would mean that the blocks
	// aren't all balanced (or there is a design flaw in this
	// function), but they are balanced because the preparser
	// verified that.  It should also be impossible for the
	// aMode to be ONLY_ONE_LINE because the function is only
	// called in that mode to execute the first action-line
	// beneath an IF or an ELSE, and the preparser has already
	// verified that every such IF and ELSE has a non-NULL
	// line after it.  Finally, aMode can be UNTIL_RETURN, but
	// that is normal mode of operation at the top level,
	// so probably shouldn't be considered an error.  For example,
	// if the script has no hotkeys, it is executed from its
	// first line all the way to the end.  For it not to have
	// a RETURN or EXIT is not an error.  UPDATE: The loader
	// now ensures that all scripts end in ACT_EXIT, so
	// this line should never be reached:
	return OK;
}



ResultType Line::EvaluateCondition() // __forceinline on this reduces benchmarks, probably because it reduces caching effectiveness by having code in the case that doesn't execute much in the benchmarks.
// Returns CONDITION_TRUE or CONDITION_FALSE (FAIL is returned only in DEBUG mode).
{
#ifdef _DEBUG
	if (!ACT_IS_IF(mActionType))
		return LineError(_T("DEBUG: EvaluateCondition() was called with a line that isn't a condition."));
#endif

	SymbolType var_is_pure_numeric, value_is_pure_numeric, value2_is_pure_numeric;
	int if_condition;
	BOOL arg2_has_binary_integer, arg3_has_binary_integer;
	Var *arg_var1, *arg_var2, *arg_var3;

	switch (mActionType)
	{
	case ACT_IFEXPR: // Listed first for performance.
		// The following is ordered for short-circuit performance. No need to check if it's g_ErrorLevel
		// (like ArgMustBeDereferenced() does) because ACT_IFEXPR doesn't internally change ErrorLevel.
		// Also, RAW is safe because loadtime validation ensured there is at least 1 arg.
		if_condition = (ARGVARRAW1 && !*ARG1 && ARGVARRAW1->Type() == VAR_NORMAL)
			? LegacyVarToBOOL(*ARGVARRAW1) // 30% faster than having ExpandArgs() resolve ARG1 even when it's a naked variable.
			: LegacyResultToBOOL(ARG1); // CAN'T simply check *ARG1=='1' because the loadtime routine has various ways of setting if_expresion to false for things that are normally expressions.
		break;

	case ACT_IFEQUAL:
	case ACT_IFNOTEQUAL:
		// For now, these seem to be the best rules to follow:
		// 1) If either one is non-empty and non-numeric, they're compared as strings.
		// 2) Otherwise, they're compared as numbers (with empty vars treated as zero).
		// In light of the above, two empty values compared to each other is the same as
		// "0 compared to 0".  e.g. if the clipboard is blank, the line "if clipboard ="
		// would be true.  However, the following are side-effects (are there any more?):
		// if var1 =    ; statement is true if var1 contains a literal zero (possibly harmful)
		// if var1 = 0  ; statement is true if var1 is blank (mostly harmless?)
		// if var1 !=   ; statement is false if var1 contains a literal zero (possibly harmful)
		// if var1 != 0 ; statement is false if var1 is blank (mostly harmless?)
		// In light of the above, the BOTH_ARE_NUMERIC macro has been altered to return
		// false if one of the items is a literal zero and the other is blank, so that
		// the two items will be compared as strings.  UPDATE: Altered it again because it
		// seems best to consider blanks to always be non-numeric (i.e. if either var is blank,
		// they will be compared as strings rather than as numbers):

		// Notes about the macros below:
		// Ordered for short-circuit performance. No need to check if it's g_ErrorLevel (like
		// ArgMustBeDereferenced() does) because the commands that use these macros don't internally
		// change ErrorLevel.
		// RAW is safe (for arg_var1) because loadtime validation ensured there is at least 1 arg.
		// For arg_var1, it isn't necessary to check ARGVARRAW1!=NULL because arg1 is ARG_TYPE_INPUT_VAR
		// for all the commands that use these macros, so loadtime validation ensures ARGVARRAW1!=NULL.
		#undef DETERMINE_NUMERIC_TYPES
		#define DETERMINE_NUMERIC_TYPES \
			if (arg2_has_binary_integer = mArgc > 1 && mArg[1].postfix && !mArg[1].is_expression)\
			{\
				arg_var2 = NULL;\
				value_is_pure_numeric = PURE_INTEGER;\
			}\
			else\
			{\
				arg_var2 = (mArgc > 1 && ARGVARRAW2 && !*ARG2 && ARGVARRAW2->Type() == VAR_NORMAL) ? ARGVARRAW2 : NULL;\
				value_is_pure_numeric = arg_var2 ? arg_var2->IsNonBlankIntegerOrFloat()\
					: IsPureNumeric(ARG2, true, false, true);\
			}\
			arg_var1 = (!*ARG1 && ARGVARRAW1->Type() == VAR_NORMAL) ? ARGVARRAW1 : NULL;\
			var_is_pure_numeric = arg_var1 ? arg_var1->IsNonBlankIntegerOrFloat()\
				: IsPureNumeric(ARG1, true, false, true);

		#define DETERMINE_NUMERIC_TYPES2 \
			DETERMINE_NUMERIC_TYPES \
			if (arg3_has_binary_integer = mArgc > 2 && mArg[2].postfix && !mArg[2].is_expression)\
			{\
				arg_var3 = NULL;\
				value2_is_pure_numeric = PURE_INTEGER;\
			}\
			else\
			{\
				arg_var3 = (mArgc > 2 && ARGVARRAW3 && !*ARG3 && ARGVARRAW3->Type() == VAR_NORMAL) ? ARGVARRAW3 : NULL;\
				value2_is_pure_numeric = arg_var3 ? arg_var3->IsNonBlankIntegerOrFloat()\
					: IsPureNumeric(ARG3, true, false, true);\
			}
		#define ARG1_AS_STRING (arg_var1 ? arg_var1->Contents() : ARG1)
		#define ARG2_AS_STRING (arg_var2 ? arg_var2->Contents() : ARG2)
		#define ARG3_AS_STRING (arg_var3 ? arg_var3->Contents() : ARG3)
		#define ARG1_AS_DOUBLE (arg_var1 ? arg_var1->ToDouble(TRUE) : ATOF(ARG1))
		#define ARG2_AS_DOUBLE (arg2_has_binary_integer ? (double)*(__int64*)mArg[1].postfix \
			: arg_var2 ? arg_var2->ToDouble(TRUE) : ATOF(ARG2))
		#define ARG3_AS_DOUBLE (arg3_has_binary_integer ? (double)*(__int64*)mArg[2].postfix \
			: arg_var3 ? arg_var3->ToDouble(TRUE) : ATOF(ARG3))
		#define ARG1_AS_INT64 (arg_var1 ? arg_var1->ToInt64(TRUE) : ATOI64(ARG1)) // Never has a binary integer because it's ARG_TYPE_INPUT_VAR.
		#define ARG2_AS_INT64 (arg2_has_binary_integer ? *(__int64*)mArg[1].postfix \
			: arg_var2 ? arg_var2->ToInt64(TRUE) : ATOI64(ARG2))
		#define ARG3_AS_INT64 (arg3_has_binary_integer ? *(__int64*)mArg[2].postfix \
			: arg_var3 ? arg_var3->ToInt64(TRUE) : ATOI64(ARG3))
		#define IF_EITHER_IS_NON_NUMERIC if (!value_is_pure_numeric || !var_is_pure_numeric)
		#define IF_EITHER_IS_NON_NUMERIC2 if (!value_is_pure_numeric || !value2_is_pure_numeric || !var_is_pure_numeric)
		#undef IF_EITHER_IS_FLOAT
		#define IF_EITHER_IS_FLOAT if (value_is_pure_numeric == PURE_FLOAT || var_is_pure_numeric == PURE_FLOAT)

		// In the below, it isn't necessary to check ARGVARRAW1!=NULL because arg1 is ARG_TYPE_INPUT_VAR
		// for all the commands that use these macros, so loadtime validation ensures ARGVARRAW1!=NULL.
		if (mArgc > 1 && ARGVARRAW1->IsBinaryClip() && ARGVARRAW2 && ARGVARRAW2->IsBinaryClip())
			if_condition = (ARGVARRAW1->Length() == ARGVARRAW2->Length()) // Accessing ARGVARRAW in all these places is safe due to the check mArgc > 1.
				&& !tmemcmp(ARGVARRAW1->Contents(), ARGVARRAW2->Contents(), ARGVARRAW1->Length());
		else
		{
			DETERMINE_NUMERIC_TYPES
			IF_EITHER_IS_NON_NUMERIC
				if_condition = !g_tcscmp(ARG1_AS_STRING, ARG2_AS_STRING);
			else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
				if_condition = ARG1_AS_DOUBLE == ARG2_AS_DOUBLE;
			else
				if_condition = ARG1_AS_INT64 == ARG2_AS_INT64;
		}
		if (mActionType == ACT_IFNOTEQUAL)
			if_condition = !if_condition;
		break;

	case ACT_IFLESS:
	case ACT_IFGREATEROREQUAL:
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_NON_NUMERIC
			if_condition = g_tcscmp(ARG1_AS_STRING, ARG2_AS_STRING) < 0;
		else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
			if_condition = ARG1_AS_DOUBLE < ARG2_AS_DOUBLE;
		else
			if_condition = ARG1_AS_INT64 < ARG2_AS_INT64;
		if (mActionType == ACT_IFGREATEROREQUAL)
			if_condition = !if_condition;
		break;
	case ACT_IFGREATER:
	case ACT_IFLESSOREQUAL:
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_NON_NUMERIC
			if_condition = g_tcscmp(ARG1_AS_STRING, ARG2_AS_STRING) > 0;
		else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
			if_condition = ARG1_AS_DOUBLE > ARG2_AS_DOUBLE;
		else
			if_condition = ARG1_AS_INT64 > ARG2_AS_INT64;
		if (mActionType == ACT_IFLESSOREQUAL)
			if_condition = !if_condition;
		break;

	case ACT_IFBETWEEN:
	case ACT_IFNOTBETWEEN:
		// Using the new macros is up to 62% faster than the old way that didn't exploit the ability of
		// one or more of the args involved to be variables that can cache binary numbers.
		DETERMINE_NUMERIC_TYPES2
		IF_EITHER_IS_NON_NUMERIC2
		{
			LPTSTR arg1_as_string = ARG1_AS_STRING; // Resolve only once.
			LPTSTR arg2_as_string = ARG2_AS_STRING; //
			LPTSTR arg3_as_string = ARG3_AS_STRING; //
			if (g->StringCaseSense == SCS_INSENSITIVE) // The most common mode is listed first for performance.
				if_condition = !(_tcsicmp(arg1_as_string, arg2_as_string) < 0 || _tcsicmp(arg1_as_string, arg3_as_string) > 0);
			else if (g->StringCaseSense == SCS_INSENSITIVE_LOCALE)
				if_condition = lstrcmpi(arg1_as_string, arg2_as_string) > -1 && lstrcmpi(arg1_as_string, arg3_as_string) < 1;
			else  // case sensitive
				if_condition = !(_tcscmp(arg1_as_string, arg2_as_string) < 0 || _tcscmp(arg1_as_string, arg3_as_string) > 0);
		}
		else IF_EITHER_IS_FLOAT
		{
			double arg1_as_double = ARG1_AS_DOUBLE;
			if_condition = arg1_as_double >= ARG2_AS_DOUBLE && arg1_as_double <= ARG3_AS_DOUBLE;
		}
		else
		{
			__int64 arg1_as_int64 = ARG1_AS_INT64;
			if_condition = arg1_as_int64 >= ARG2_AS_INT64 && arg1_as_int64 <= ARG3_AS_INT64;
		}
		if (mActionType == ACT_IFNOTBETWEEN)
			if_condition = !if_condition;
		break;

	case ACT_IFINSTRING:
	case ACT_IFNOTINSTRING:
	{
		// The most common mode is listed first for performance:
		if_condition = g_tcsstr(ARG1, ARG2) != NULL; // To reduce code size, resolve large macro only once for both these commands.
		if (mActionType == ACT_IFNOTINSTRING)
			if_condition = !if_condition;
		break;
	}

	case ACT_IFIN:
	case ACT_IFNOTIN:
		if_condition = IsStringInList(ARG1, ARG2, true);
		if (mActionType == ACT_IFNOTIN)
			if_condition = !if_condition;
		break;

	case ACT_IFCONTAINS:
	case ACT_IFNOTCONTAINS:
		if_condition = IsStringInList(ARG1, ARG2, false);
		if (mActionType == ACT_IFNOTCONTAINS)
			if_condition = !if_condition;
		break;

	case ACT_IFIS:
	case ACT_IFISNOT:
	{
		LPTSTR cp;
		VariableTypeType variable_type = ConvertVariableTypeName(ARG2);
		if (variable_type == VAR_TYPE_INVALID)
		{
			// Type is probably a dereferenced variable that resolves to an invalid type name.
			// It seems best to make the condition false in these cases, rather than pop up
			// a runtime error dialog:
			if_condition = false;
			break;
		}

		// Although ExpandArgs() has already flushed the binary-number cache for ACT_IFIS/ACT_IFISNOT
		// (since it doesn't optimize these due to them having too many subcommands/modes), can still
		// read from the binary number cache to avoid having to convert from text-to-number.
		// In the below, it isn't necessary to check ARGVARRAW1!=NULL because arg1 is ARG_TYPE_INPUT_VAR
		// for all the commands that use these macros, so loadtime validation ensures ARGVARRAW1!=NULL.
		arg_var1 = (ARGVARRAW1->Type() == VAR_NORMAL) ? ARGVARRAW1 : NULL;

		switch(variable_type)
		{
		case VAR_TYPE_NUMBER:
			if_condition = arg_var1 ? arg_var1->IsNonBlankIntegerOrFloat()
				: IsPureNumeric(ARG1, true, false, true);  // Floats are defined as being numeric.
			break;
		case VAR_TYPE_INTEGER:
			if_condition = arg_var1 ? (arg_var1->IsNonBlankIntegerOrFloat() == PURE_INTEGER) // Explicitly compare to PURE_INTEGER because IsNonBlankIntegerOrFloat() doesn't support aAllowFloat.
				: IsPureNumeric(ARG1, true, false, false);  // Passes false for aAllowFloat.
			break;
		case VAR_TYPE_FLOAT:
			if_condition = arg_var1 ? (arg_var1->IsNonBlankIntegerOrFloat() == PURE_FLOAT) // Explicitly compare to PURE_FLOAT.
				: (IsPureNumeric(ARG1, true, false, true) == PURE_FLOAT);
			break;
		case VAR_TYPE_TIME:
		{
			SYSTEMTIME st;
			// Also insist on numeric, because even though YYYYMMDDToFileTime() will properly convert a
			// non-conformant string such as "2004.4", for future compatibility, we don't want to
			// report that such strings are valid times:
			if_condition = IsPureNumeric(ARG1, false, false, false) && YYYYMMDDToSystemTime(ARG1, st, true); // Can't call IsNonBlankIntegerOrFloat() here because it doesn't support aAllowNegative.
			break;
		}
		case VAR_TYPE_DIGIT:
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				if (!_istdigit((UCHAR)*cp))
				{
					if_condition = false;
					break;
				}
			break;
		case VAR_TYPE_XDIGIT:
			cp = ARG1;
			if (!_tcsnicmp(cp, _T("0x"), 2)) // v1.0.44.09: Allow 0x prefix, which seems to do more good than harm (unlikely to break existing scripts).
				cp += 2;
			if_condition = true;
			for (; *cp; ++cp)
				if (!_istxdigit((UCHAR)*cp))
				{
					if_condition = false;
					break;
				}
			break;
		case VAR_TYPE_ALNUM:
			// Like AutoIt3, the empty string is considered to be alphabetic, which is only slightly debatable.
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				//if (!IsCharAlphaNumeric(*cp)) // Use this to better support chars from non-English languages.
				if (!aisalnum(*cp)) // But some users don't like it, Chinese users for example.
				{
					if_condition = false;
					break;
				}
			break;
		case VAR_TYPE_ALPHA:
			// Like AutoIt3, the empty string is considered to be alphabetic, which is only slightly debatable.
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				//if (!IsCharAlpha(*cp)) // Use this to better support chars from non-English languages.
				if (!aisalpha(*cp)) // But some users don't like it, Chinese users for example.
				{
					if_condition = false;
					break;
				}
			break;
		case VAR_TYPE_UPPER:
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				//if (!IsCharUpper(*cp)) // Use this to better support chars from non-English languages.
				if (!aisupper(*cp)) // But some users don't like it, Chinese users for example.
				{
					if_condition = false;
					break;
				}
			break;
		case VAR_TYPE_LOWER:
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				//if (!IsCharLower(*cp)) // Use this to better support chars from non-English languages.
				if (!aislower(*cp)) // But some users don't like it, Chinese users for example.
				{
					if_condition = false;
					break;
				}
			break;
		case VAR_TYPE_SPACE:
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				if (!_istspace(*cp))
				{
					if_condition = false;
					break;
				}
			break;
		}
		if (mActionType == ACT_IFISNOT)
			if_condition = !if_condition;
		break;
	}

	// For ACT_IFWINEXIST and ACT_IFWINNOTEXIST, although we validate that at least one
	// of their window params is non-blank during load, it's okay at runtime for them
	// all to resolve to be blank (due to derefs), without an error being reported.
	// It's probably more flexible that way, and in any event WinExist() is equipped to
	// handle all-blank params:
	case ACT_IFWINEXIST:
		// NULL-check this way avoids compiler warnings:
		if_condition = (WinExist(*g, FOUR_ARGS, false, true) != NULL);
		break;
	case ACT_IFWINNOTEXIST:
		if_condition = !WinExist(*g, FOUR_ARGS, false, true); // Seems best to update last-used even here.
		break;
	case ACT_IFWINACTIVE:
		if_condition = (WinActive(*g, FOUR_ARGS, true) != NULL);
		break;
	case ACT_IFWINNOTACTIVE:
		if_condition = !WinActive(*g, FOUR_ARGS, true);
		break;

	case ACT_IFEXIST:
		if_condition = DoesFilePatternExist(ARG1);
		break;
	case ACT_IFNOTEXIST:
		if_condition = !DoesFilePatternExist(ARG1);
		break;

	case ACT_IFMSGBOX:
	{
		int mb_result = ConvertMsgBoxResult(ARG1);
		// Seems best not to report runtime error for such a rare thing, just let it discover "false" further below:
		//if (!mb_result)
		//	return LineError(ERR_PARAM1_INVALID, FAIL, ARG1);
		if_condition = (g->MsgBoxResult == mb_result);
		break;
	}
#ifdef _DEBUG
	default: // Should never happen, but return an error if it does.
		return LineError(_T("DEBUG: EvaluateCondition(): Unhandled type of IF."));
#endif
	}
	return if_condition ? CONDITION_TRUE : CONDITION_FALSE;
}


// Evaluate an #If expression (in a thread created by the caller).
ResultType Line::EvaluateHotCriterionExpression()
{
	g_script.mCurrLine = this; // Added in v1.1.16 to fix A_LineFile and A_LineNumber.

	if (g->ListLinesIsEnabled)
		LOG_LINE(this)

	ResultType result = ExpandArgs();
	if (result == OK)
		result = EvaluateCondition();

	return result;
}


// Evaluate an #If expression or callback function.
// This is called by MainWindowProc when it receives an AHK_HOT_IF_EVAL message.
ResultType HotkeyCriterion::Eval(LPTSTR aHotkeyName)
{
	// Initialize a new quasi-thread to evaluate the expression. This may not be necessary for simple
	// expressions, but expressions which call user-defined functions may otherwise interfere with
	// whatever quasi-thread is running when the hook thread requests that this expression be evaluated.
	
	// Based on parts of MsgMonitor(). See there for comments.

	if (g_nThreads >= g_MaxThreadsTotal)
		return CONDITION_FALSE;

	bool prev_defer_messages = g_DeferMessagesForUnderlyingPump;
	// Force the use of PeekMessage() within MsgSleep() since GetMessage() is known to stall while
	// the system is waiting for our keyboard hook to return (last confirmed on Windows 10.0.18356).
	// This might relate to WM_TIMER being lower priority than the input hardware processing
	// performed by GetMessage().  MsgSleep() relies on WM_TIMER acting as a timeout for GetMessage().
	g_DeferMessagesForUnderlyingPump = true;

	// See MsgSleep() for comments about the following section.
	TCHAR ErrorLevel_saved[ERRORLEVEL_SAVED_SIZE];
	tcslcpy(ErrorLevel_saved, g_ErrorLevel->Contents(), _countof(ErrorLevel_saved));
	// Critical seems to improve reliability, either because the thread completes faster (i.e. before the timeout) or because we check for messages less often.
	InitNewThread(0, false, true, ACT_CRITICAL);
	ResultType result;

	// Update A_ThisHotkey, useful if #If calls a function to do its dirty work.
	LPTSTR prior_hotkey_name[] = { g_script.mThisHotkeyName, g_script.mPriorHotkeyName };
	DWORD prior_hotkey_time[] = { g_script.mThisHotkeyStartTime, g_script.mPriorHotkeyStartTime };
	g_script.mPriorHotkeyName = g_script.mThisHotkeyName;			// For consistency
	g_script.mPriorHotkeyStartTime = g_script.mThisHotkeyStartTime; //
	g_script.mThisHotkeyName = aHotkeyName;
	g_script.mThisHotkeyStartTime = // Updated for consistency.
	g_script.mLastScriptRest = g_script.mLastPeekTime = GetTickCount();

	// EVALUATE THE EXPRESSION OR CALL THE CALLBACK
	if (Type == HOT_IF_EXPR)
	{
		DEBUGGER_STACK_PUSH(_T("#If"))
		result = ExprLine->EvaluateHotCriterionExpression();
		DEBUGGER_STACK_POP()
	}
	else
	{
		ExprTokenType param = aHotkeyName;
		INT_PTR retval;
		result = LabelPtr(Callback)->ExecuteInNewThread(_T("#If"), &param, 1, &retval);
		if (result != FAIL)
			result = retval ? CONDITION_TRUE : CONDITION_FALSE;
	}

	// The following allows the expression to set the Last Found Window for the
	// hotkey subroutine, so that #if WinActive(T) and similar behave like #IfWin.
	// There may be some rare cases where the wrong hotkey gets this HWND (perhaps
	// if there are multiple hotkey messages in the queue), but there doesn't seem
	// to be any easy way around that.
	g_HotExprLFW = g->hWndLastUsed; // Even if above failed, for simplicity.

	// A_ThisHotkey must be restored else A_PriorHotkey will get an incorrect value later.
	g_script.mThisHotkeyName = prior_hotkey_name[0];
	g_script.mThisHotkeyStartTime = prior_hotkey_time[0];
	g_script.mPriorHotkeyName = prior_hotkey_name[1];
	g_script.mPriorHotkeyStartTime = prior_hotkey_time[1];

	ResumeUnderlyingThread(ErrorLevel_saved);

	g_DeferMessagesForUnderlyingPump = prev_defer_messages;

	return result;
}


ResultType Line::PerformLoop(ExprTokenType *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil
	, __int64 aIterationLimit, bool aIsInfinite) // bool performs better than BOOL in current benchmarks for this.
// This performs much better (by at least 7%) as a function than as inline code, probably because
// it's only called to set up the loop, not each time through the loop.
{
	ResultType result;
	Line *jump_to_line;
	global_struct &g = *::g; // Primarily for performance in this case.

	for (; aIsInfinite || g.mLoopIteration <= aIterationLimit; ++g.mLoopIteration)
	{
		// Execute once the body of the loop (either just one statement or a block of statements).
		// Preparser has ensured that every LOOP has a non-NULL next line.
		if (mNextLine->mActionType == ACT_BLOCK_BEGIN)
		{
			// Simple loops are about 6% faster by omitting the open-brace from ListLines, so
			// it seems worth it:
			//if (g.ListLinesIsEnabled)
			//{
			//	sLog[sLogNext] = mNextLine; // See comments in ExecUntil() about this section.
			//	sLogTick[sLogNext++] = GetTickCount();
			//	if (sLogNext >= LINE_LOG_SIZE)
			//		sLogNext = 0;
			//}

			// If this loop has a block under it rather than just a single line, take a shortcut
			// and directly execute the block.  This avoids one recursive call to ExecUntil()
			// for each iteration, which can speed up short/fast loops by as much as 30%.
			// Another benefit is conservation of stack space, especially during "else if" ladders.
			//
			// At this point, mNextLine->mNextLine is already verified non-NULL by the pre-parser
			// because it checks that every LOOP has a line under it (ACT_BLOCK_BEGIN in this case)
			// and that every ACT_BLOCK_BEGIN has at least one line under it.
			do
				result = mNextLine->mNextLine->ExecUntil(UNTIL_BLOCK_END, aResultToken, &jump_to_line);
			while (jump_to_line == mNextLine); // The above call encountered a Goto that jumps to the "{". See ACT_BLOCK_BEGIN in ExecUntil() for details.
		}
		else
			result = mNextLine->ExecUntil(ONLY_ONE_LINE, aResultToken, &jump_to_line);
		if (jump_to_line && !(result == LOOP_CONTINUE && jump_to_line == this)) // i.e. "goto somewhere" or "continue a_loop_which_encloses_this_one".
		{
			if (jump_to_line == this) 
				// Since this LOOP's ExecUntil() encountered a Goto whose target is the LOOP
				// itself, continue with the for-loop without moving to a different
				// line.  Also: stay in this recursion layer even if aMode == ONLY_ONE_LINE
				// because we don't want the caller handling it because then it's cleanup
				// to jump to its end-point (beyond its own and any unowned elses) won't work.
				// Example:
				// if x  <-- If this layer were to do it, its own else would be unexpectedly encountered.
				//    label1:
				//    loop  <-- We want this statement's layer to handle the goto.
				//       goto, label1
				// else
				//   ...
				// Also, signal all our callers to return until they get back to the original
				// ExecUntil() instance that started the loop:
				aContinueMainLoop = true;
			else // jump_to_line must be a line that's at the same level or higher as our Exec_Until's LOOP statement itself.
				aJumpToLine = jump_to_line; // Signal the caller to handle this jump.
			return result;
		}
		if (result != OK && result != LOOP_CONTINUE) // i.e. result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL)
			return result;
		if (aUntil && aUntil->EvaluateLoopUntil(result))
			return result;
		// Otherwise, the result of executing the body of the loop, above, was either OK
		// (the current iteration completed normally) or LOOP_CONTINUE (the current loop
		// iteration was cut short).  In both cases, just continue on through the loop.
	} // for()

	// The script's loop is now over.
	return OK;
}



// Lexikos: ACT_WHILE
ResultType Line::PerformLoopWhile(ExprTokenType *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine)
{
	ResultType result;
	Line *jump_to_line;
	global_struct &g = *::g; // Might slightly speed up the loop below.

	for (;; ++g.mLoopIteration)
	{
		g_script.mCurrLine = this; // For error-reporting purposes.
#ifdef CONFIG_DEBUGGER
		// L31: Let the debugger break at the 'While' line each iteration. Before this change,
		// a While loop with empty body such as While FuncWithSideEffect() {} would be "hit"
		// (via breakpoint or step) only once even if the loop had multiple iterations.
		// A_Index was also reported as "0"; it will now be reported correctly.
		if (g_Debugger.IsConnected())
			g_Debugger.PreExecLine(this);
#endif
		// Evaluate the expression only now that A_Index has been set.
		result = ExpandArgs();
		if (result != OK)
			return result;

		// Unlike if(expression), performance isn't significantly improved to make cases like
		// "while x" and "while %x%" into non-expressions (the latter actually performs much
		// better as an expression).  That is why the following check is much simpler than the
		// one used at at ACT_IFEXPR in EvaluateCondition():
		if (!LegacyResultToBOOL(ARG1))
			break;

		// CONCERNING ALL THE REST OF THIS FUNCTION: See comments in PerformLoop() for details.
		if (mNextLine->mActionType == ACT_BLOCK_BEGIN)
			do
				result = mNextLine->mNextLine->ExecUntil(UNTIL_BLOCK_END, aResultToken, &jump_to_line);
			while (jump_to_line == mNextLine);
		else
			result = mNextLine->ExecUntil(ONLY_ONE_LINE, aResultToken, &jump_to_line);
		if (jump_to_line && !(result == LOOP_CONTINUE && jump_to_line == this))
		{
			if (jump_to_line == this)
				aContinueMainLoop = true;
			else
				aJumpToLine = jump_to_line;
			return result;
		}
		if (result != OK && result != LOOP_CONTINUE)
			return result;

		// Before re-evaluating the condition, add it to the ListLines log again.  This is done
		// at the end of the loop rather than the beginning because ExecUntil already added the
		// line once immediately before the first iteration.
		if (g.ListLinesIsEnabled)
			LOG_LINE(this)
	} // for()
	return OK; // The script's loop is now over.
}



bool Line::EvaluateLoopUntil(ResultType &aResult)
{
	g_script.mCurrLine = this; // For error-reporting purposes.
	if (g->ListLinesIsEnabled)
		LOG_LINE(this);
#ifdef CONFIG_DEBUGGER
	// Let the debugger break at or step onto UNTIL.
	if (g_Debugger.IsConnected())
		g_Debugger.PreExecLine(this);
#endif
	aResult = ExpandArgs();
	if (aResult != OK)
		return true; // i.e. if it fails, break the loop.
	aResult = LOOP_BREAK; // Break out of any recursive PerformLoopXxx() calls.
	return LegacyResultToBOOL(ARG1); // See PerformLoopWhile() above for comments about this line.
}



ResultType Line::PerformLoopFor(ExprTokenType *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil)
{
	ResultType result;
	Line *jump_to_line;
	global_struct &g = *::g; // Might slightly speed up the loop below.

	// Save these pointers since they will be overwritten during the loop:
	Var *var[] = { ARGVARRAW1, ARGVARRAW2 };
	
	if (!sDerefBuf)
	{
		// This must be done in case ExpandExpression() needs the deref buf for temporary storage.
		sDerefBufSize = (mArg[2].length < MAX_NUMBER_LENGTH ? MAX_NUMBER_LENGTH : mArg[2].length) + 1; // See EXPR_BUF_SIZE macro in script_expression.cpp.
		if ( !(sDerefBuf = tmalloc(sDerefBufSize)) )
		{
			sDerefBufSize = 0;
			return LineError(ERR_OUTOFMEM);
		}
	}

	PRIVATIZE_S_DEREF_BUF;
	LPTSTR our_buf_marker = our_deref_buf;
	LPTSTR arg_deref[] = {0, 0}; // ExpandExpression checks these if it needs to expand the deref buffer.
	ExprTokenType object_token;
	object_token.symbol = SYM_INVALID; // Init in case ExpandExpression() resolves to a string, in which case it won't use enum_token.

	// Since expressions aren't normally capable of resolving to an object (except for RETURN), we need to
	// call ExpandExpression() directly and pass in a "result token" which will be used if the result is an
	// object or number. Load-time pre-parsing has ensured there are really three args, but mArgc == 2 so
	// this one hasn't been evaluated yet:
	if (ExpandExpression(2, result, &object_token, our_buf_marker, our_deref_buf, our_deref_buf_size, arg_deref, 0))
		result = OK;
	
	DEPRIVATIZE_S_DEREF_BUF

	if (result == FAIL || result == EARLY_EXIT)
		// A script-function-call inside the expression returned EARLY_EXIT or FAIL.
		return result;

	if (object_token.symbol != SYM_OBJECT)
		// The expression didn't resolve to an object, so no enumerator is available.
		return OK;
	
	TCHAR buf[MAX_NUMBER_SIZE]; // Small buffer which may be used by object->Invoke().
	
	ExprTokenType enum_token;
	ExprTokenType param_tokens[3];
	ExprTokenType *params[] = { param_tokens, param_tokens+1, param_tokens+2 };
	int param_count;

	// Set up enum_token the way Invoke expects:
	enum_token.symbol = SYM_STRING;
	enum_token.marker = _T("");
	enum_token.mem_to_free = NULL;
	enum_token.buf = buf;

	// Prepare to call object._NewEnum():
	param_tokens[0].symbol = SYM_STRING;
	param_tokens[0].marker = _T("_NewEnum");

	result = object_token.object->Invoke(enum_token, object_token, IT_CALL | IF_NEWENUM, params, 1);
	object_token.object->Release(); // This object reference is no longer needed.

	if (enum_token.mem_to_free)
		// Invoke returned memory for us to free.
		free(enum_token.mem_to_free);
	
	if (result == FAIL || result == EARLY_EXIT)
		return result;

	if (enum_token.symbol != SYM_OBJECT)
		// The object didn't return an enumerator, so nothing more we can do.
		return OK;

	// Prepare parameters for the loop below: enum.Next(var1 [, var2])
	param_tokens[0].marker = _T("Next");
	param_tokens[1].symbol = SYM_VAR;
	param_tokens[1].var = var[0];
	if (var[1])
	{
		// for x,y in z  ->  enum.Next(x,y)
		param_tokens[2].symbol = SYM_VAR;
		param_tokens[2].var = var[1];
		param_count = 3;
	}
	else
		// for x in z  ->  enum.Next(x)
		param_count = 2;

	IObject &enumerator = *enum_token.object; // Might perform better as a reference?

	ExprTokenType result_token;

	// Now that the enumerator expression has been evaluated, init A_Index:
	g.mLoopIteration = 1;

	for (;; ++g.mLoopIteration)
	{
		// Set up result_token the way Invoke expects; each Invoke() will change some or all of these:
		result_token.symbol = SYM_STRING;
		result_token.marker = _T("");
		result_token.mem_to_free = NULL;
		result_token.buf = buf;

		// Call enumerator.Next(var1, var2)
		result = enumerator.Invoke(result_token, enum_token, IT_CALL, params, param_count);
		if (result == FAIL || result == EARLY_EXIT)
			break;

		// Since any non-empty SYM_STRING is always considered "true", we need to change it to SYM_OPERAND
		// before calling TokenToBOOL; otherwise "return false" and "return 0" will both evaluate to true
		// since they are optimized to not be expressions (and only expressions return pure numeric values).
		if (result_token.symbol == SYM_STRING)
		{
			result_token.symbol = SYM_OPERAND;
			result_token.buf = NULL; // Indicate that this SYM_OPERAND token LACKS a pre-converted binary integer.
		}

		bool next_returned_true = TokenToBOOL(result_token);

		// Free any memory or object which may have been returned by Invoke:
		if (result_token.mem_to_free)
			free(result_token.mem_to_free);
		if (result_token.symbol == SYM_OBJECT)
			result_token.object->Release(); // Relies on the fact that TokenToBool() doesn't access the object.

		if (!next_returned_true)
		{	// The enumerator returned false, which means there are no more items.
			result = OK;
			break;
		}
		// Otherwise the enumerator already stored the next value(s) in the variable(s) we passed it via params.

		// CONCERNING ALL THE REST OF THIS FUNCTION: See comments in PerformLoop() for details.
		if (mNextLine->mActionType == ACT_BLOCK_BEGIN)
			do
				result = mNextLine->mNextLine->ExecUntil(UNTIL_BLOCK_END, aResultToken, &jump_to_line);
			while (jump_to_line == mNextLine);
		else
			result = mNextLine->ExecUntil(ONLY_ONE_LINE, aResultToken, &jump_to_line);
		if (jump_to_line && !(result == LOOP_CONTINUE && jump_to_line == this))
		{
			if (jump_to_line == this)
				aContinueMainLoop = true;
			else
				aJumpToLine = jump_to_line;
			break;
		}
		if (result != OK && result != LOOP_CONTINUE)
			break;
		if (aUntil && aUntil->EvaluateLoopUntil(result))
			break;
	} // for()
	enumerator.Release();
	return result; // The script's loop is now over.
}



ResultType Line::PerformLoopFilePattern(ExprTokenType *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil
	, FileLoopModeType aFileLoopMode, bool aRecurseSubfolders, LPTSTR aFilePattern)
// Note: Even if aFilePattern is just a directory (i.e. with not wildcard pattern), it seems best
// not to append "\\*.*" to it because the pattern might be a script variable that the user wants
// to conditionally resolve to various things at runtime.  In other words, it's valid to have
// only a single directory be the target of the loop.
{
	// Make a local copy of the path given in aFilePattern because as the lines of
	// the loop are executed, the deref buffer (which is what aFilePattern might
	// point to if we were called from ExecUntil()) may be overwritten --
	// and we will need the path string for every loop iteration.  We also need
	// to determine naked_filename_or_pattern:
	TCHAR file_path[MAX_PATH], naked_filename_or_pattern[MAX_PATH]; // Giving +3 extra for "*.*" seems fairly pointless because any files that actually need that extra room would fail to be retrieved by FindFirst/Next due to their inability to support paths much over 256.
	size_t file_path_length;
	tcslcpy(file_path, aFilePattern, _countof(file_path));
	LPTSTR last_backslash = _tcsrchr(file_path, '\\');
	if (last_backslash)
	{
		_tcscpy(naked_filename_or_pattern, last_backslash + 1); // Naked filename.  No danger of overflow due size of src vs. dest.
		*(last_backslash + 1) = '\0';  // Convert file_path to be the file's path, but use +1 to retain the final backslash on the string.
		file_path_length = _tcslen(file_path);
	}
	else
	{
		_tcscpy(naked_filename_or_pattern, file_path); // No danger of overflow due size of src vs. dest.
		*file_path = '\0'; // There is no path, so make it empty to use current working directory.
		file_path_length = 0;
	}

	// g->mLoopFile is the current file of the file-loop that encloses this file-loop, if any.
	// The below is our own current_file, which will take precedence over g->mLoopFile if this
	// loop is a file-loop:
	BOOL file_found;
	WIN32_FIND_DATA new_current_file;
	HANDLE file_search = FindFirstFile(aFilePattern, &new_current_file);
	for ( file_found = (file_search != INVALID_HANDLE_VALUE) // Convert FindFirst's return value into a boolean so that it's compatible with FindNext's.
		; file_found && FileIsFilteredOut(new_current_file, aFileLoopMode, file_path, file_path_length)
		; file_found = FindNextFile(file_search, &new_current_file));
	// file_found and new_current_file have now been set for use below.
	// Above is responsible for having properly set file_found and file_search.

	ResultType result;
	Line *jump_to_line;
	global_struct &g = *::g; // Primarily for performance in this case.

	for (; file_found; ++g.mLoopIteration)
	{
		g.mLoopFile = &new_current_file; // inner file-loop's file takes precedence over any outer file-loop's.
		// Other types of loops leave g.mLoopFile unchanged so that a file-loop can enclose some other type of
		// inner loop, and that inner loop will still have access to the outer loop's current file.

		// Execute once the body of the loop (either just one statement or a block of statements).
		// Preparser has ensured that every LOOP has a non-NULL next line.
		if (mNextLine->mActionType == ACT_BLOCK_BEGIN) // See PerformLoop() for comments about this section.
			do
				result = mNextLine->mNextLine->ExecUntil(UNTIL_BLOCK_END, aResultToken, &jump_to_line);
			while (jump_to_line == mNextLine);
		else
			result = mNextLine->ExecUntil(ONLY_ONE_LINE, aResultToken, &jump_to_line);
		if (jump_to_line && !(result == LOOP_CONTINUE && jump_to_line == this)) // See comments in PerformLoop() about this section.
		{
			if (jump_to_line == this)
				aContinueMainLoop = true;
			else
				aJumpToLine = jump_to_line; // Signal our caller to handle this jump.
			FindClose(file_search);
			return result;
		}
		if ( result != OK && result != LOOP_CONTINUE // i.e. result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL)
			|| (aUntil && aUntil->EvaluateLoopUntil(result)) )
		{
			FindClose(file_search);
			// Although ExecUntil() will treat the LOOP_BREAK result identically to OK, we
			// need to return LOOP_BREAK in case our caller is another instance of this
			// same function (i.e. due to recursing into subfolders):
			return result;
		}
		// Otherwise, the result of executing the body of the loop, above, was either OK
		// (the current iteration completed normally) or LOOP_CONTINUE (the current loop
		// iteration was cut short).  In both cases, just continue on through the loop.
		// But first do end-of-iteration steps:
		while ((file_found = FindNextFile(file_search, &new_current_file))
			&& FileIsFilteredOut(new_current_file, aFileLoopMode, file_path, file_path_length)); // Relies on short-circuit boolean order.
			// Above is a self-contained loop that keeps fetching files until there's no more files, or a file
			// is found that isn't filtered out.  It also sets file_found and new_current_file for use by the
			// outer loop.
	} // for()

	// The script's loop is now over.
	if (file_search != INVALID_HANDLE_VALUE)
		FindClose(file_search);

	// If aRecurseSubfolders is true, we now need to perform the loop's body for every subfolder to
	// search for more files and folders inside that match aFilePattern.  We can't do this in the
	// first loop, above, because it may have a restricted file-pattern such as *.txt and we want to
	// find and recurse into ALL folders:
	if (!aRecurseSubfolders) // No need to continue into the "recurse" section.
		return OK;

	// Since above didn't return, this is a file-loop and recursion into sub-folders has been requested.
	// Append *.* to file_path so that we can retrieve all files and folders in the aFilePattern
	// main folder.  We're only interested in the folders, but we have to use *.* to ensure
	// that the search will find all folder names:
	if (file_path_length > _countof(file_path) - 4) // v1.0.45.03: No room to append "*.*", so for simplicity, skip this folder (don't recurse into it).
		return OK; // This situation might be impossible except for 32000-capable paths because the OS seems to reserve room inside every directory for at least the maximum length of a short filename.
	LPTSTR append_pos = file_path + file_path_length;
	_tcscpy(append_pos, _T("*.*")); // Above has already verified that no overflow is possible.

	file_search = FindFirstFile(file_path, &new_current_file);
	if (file_search == INVALID_HANDLE_VALUE)
		return OK; // Nothing more to do.
	// Otherwise, recurse into any subdirectories found inside this parent directory.

	size_t path_and_pattern_length = file_path_length + _tcslen(naked_filename_or_pattern); // Calculated only once for performance.
	do
	{
		if (!(new_current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) // We only want directories (except "." and "..").
			|| new_current_file.cFileName[0] == '.' && (!new_current_file.cFileName[1]      // Relies on short-circuit boolean order.
				|| new_current_file.cFileName[1] == '.' && !new_current_file.cFileName[2])  //
			// v1.0.45.03: Skip over folders whose full-path-names are too long to be supported by the ANSI
			// versions of FindFirst/FindNext.  Without this fix, the section below formerly called PerformLoop()
			// with a truncated full-path-name, which caused the last_backslash-finding logic to find the wrong
			// backslash, which in turn caused infinite recursion and a stack overflow (i.e. caused by the
			// full-path-name getting truncated in the same spot every time, endlessly).
			|| path_and_pattern_length + _tcslen(new_current_file.cFileName) > _countof(file_path) - 2) // -2 to reflect: 1) the backslash to be added between cFileName and naked_filename_or_pattern; 2) the zero terminator.
			continue;
		// Build the new search pattern, which consists of the original file_path + the subfolder name
		// we just discovered + the original pattern:
		_stprintf(append_pos, _T("%s\\%s"), new_current_file.cFileName, naked_filename_or_pattern); // Indirectly set file_path to the new search pattern.  This won't overflow due to the check above.
		// Pass NULL for the 2nd param because it will determine its own current-file when it does
		// its first loop iteration.  This is because this directory is being recursed into, not
		// processed itself as a file-loop item (since this was already done in the first loop,
		// above, if its name matches the original search pattern):
		result = PerformLoopFilePattern(aResultToken, aContinueMainLoop, aJumpToLine, aUntil, aFileLoopMode, aRecurseSubfolders, file_path);
		// Above returns LOOP_CONTINUE for cases like "continue 2" or "continue outer_loop", where the
		// target is not this Loop but a Loop which encloses it. In those cases we want below to return:
		if (result != OK) // i.e. result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL)
		{
			FindClose(file_search);
			return result;  // Return even LOOP_BREAK, since our caller can be either ExecUntil() or ourself.
		}
		if (aContinueMainLoop // The call to PerformLoop() above signaled us to break & return.
			|| aJumpToLine)
			// Above: There's no need to check "aJumpToLine == this" because PerformLoop() would already have
			// handled it.  But if it set aJumpToLine to be non-NULL, it means we have to return and let our caller
			// handle the jump.
			break;
	} while (FindNextFile(file_search, &new_current_file));
	FindClose(file_search);

	return OK;
}



ResultType Line::PerformLoopReg(ExprTokenType *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil
	, FileLoopModeType aFileLoopMode, bool aRecurseSubfolders, HKEY aRootKeyType, HKEY aRootKey, LPTSTR aRegSubkey)
// aRootKeyType is the type of root key, independent of whether it's local or remote.
// This is used because there's no easy way to determine which root key a remote HKEY
// refers to.
{
	RegItemStruct reg_item(aRootKeyType, aRootKey, aRegSubkey);
	HKEY hRegKey;

	// Open the specified subkey.  Be sure to only open with the minimum permission level so that
	// the keys & values can be deleted or written to (though I'm not sure this would be an issue
	// in most cases):
	if (RegOpenKeyEx(reg_item.root_key, reg_item.subkey, 0, KEY_QUERY_VALUE | KEY_ENUMERATE_SUB_KEYS | g->RegView, &hRegKey) != ERROR_SUCCESS)
		return OK;

	// Get the count of how many values and subkeys are contained in this parent key:
	DWORD count_subkeys;
	DWORD count_values;
	if (RegQueryInfoKey(hRegKey, NULL, NULL, NULL, &count_subkeys, NULL, NULL
		, &count_values, NULL, NULL, NULL, NULL) != ERROR_SUCCESS)
	{
		RegCloseKey(hRegKey);
		return OK;
	}

	ResultType result;
	Line *jump_to_line;
	DWORD i;
	global_struct &g = *::g; // Primarily for performance in this case.

	// See comments in PerformLoop() for details about this section.
	// Note that &reg_item is passed to ExecUntil() rather than... (comment was never finished).
	#define MAKE_SCRIPT_LOOP_PROCESS_THIS_ITEM \
	{\
		g.mLoopRegItem = &reg_item;\
		if (mNextLine->mActionType == ACT_BLOCK_BEGIN)\
			do\
				result = mNextLine->mNextLine->ExecUntil(UNTIL_BLOCK_END, aResultToken, &jump_to_line);\
			while (jump_to_line == mNextLine);\
		else\
			result = mNextLine->ExecUntil(ONLY_ONE_LINE, aResultToken, &jump_to_line);\
		if (jump_to_line && !(result == LOOP_CONTINUE && jump_to_line == this))\
		{\
			if (jump_to_line == this)\
				aContinueMainLoop = true;\
			else\
				aJumpToLine = jump_to_line;\
			RegCloseKey(hRegKey);\
			return result;\
		}\
		if ( result != OK && result != LOOP_CONTINUE \
			|| (aUntil && aUntil->EvaluateLoopUntil(result)) ) \
		{\
			RegCloseKey(hRegKey);\
			return result;\
		}\
		++g.mLoopIteration;\
	}

	DWORD name_size;

	// First enumerate the values, which are analogous to files in the file system.
	// Later, the subkeys ("subfolders") will be done:
	if (count_values > 0 && aFileLoopMode != FILE_LOOP_FOLDERS_ONLY) // The caller doesn't want "files" (values) excluded.
	{
		reg_item.InitForValues();
		// Going in reverse order allows values to be deleted without disrupting the enumeration,
		// at least in some cases:
		for (i = count_values - 1;; --i) 
		{ 
			// Don't use CONTINUE in loops such as this due to the loop-ending condition being explicitly
			// checked at the bottom.
			name_size = _countof(reg_item.name);  // Must reset this every time through the loop.
			*reg_item.name = '\0';
			if (RegEnumValue(hRegKey, i, reg_item.name, &name_size, NULL, &reg_item.type, NULL, NULL) == ERROR_SUCCESS)
				MAKE_SCRIPT_LOOP_PROCESS_THIS_ITEM
			// else continue the loop in case some of the lower indexes can still be retrieved successfully.
			if (i == 0)  // Check this here due to it being an unsigned value that we don't want to go negative.
				break;
		}
	}

	// If the loop is neither processing subfolders nor recursing into them, don't waste the performance
	// doing the next loop:
	if (!count_subkeys || (aFileLoopMode == FILE_LOOP_FILES_ONLY && !aRecurseSubfolders))
	{
		RegCloseKey(hRegKey);
		return OK;
	}

	// Enumerate the subkeys, which are analogous to subfolders in the files system:
	// Going in reverse order allows keys to be deleted without disrupting the enumeration,
	// at least in some cases:
	reg_item.InitForSubkeys();
	TCHAR subkey_full_path[MAX_REG_ITEM_SIZE]; // But doesn't include the root key name, which is not only by design but testing shows that if it did, the length could go over MAX_REG_ITEM_SIZE.
	for (i = count_subkeys - 1;; --i) // Will have zero iterations if there are no subkeys.
	{
		// Don't use CONTINUE in loops such as this due to the loop-ending condition being explicitly
		// checked at the bottom.
		name_size = _countof(reg_item.name); // Must be reset for every iteration.
		if (RegEnumKeyEx(hRegKey, i, reg_item.name, &name_size, NULL, NULL, NULL, &reg_item.ftLastWriteTime) == ERROR_SUCCESS)
		{
			if (aFileLoopMode != FILE_LOOP_FILES_ONLY) // have the script's loop process this subkey.
				MAKE_SCRIPT_LOOP_PROCESS_THIS_ITEM
			if (aRecurseSubfolders) // Now recurse into the subkey, regardless of whether it was processed above.
			{
				// Build the new subkey name using the an area of memory on the stack that we won't need
				// after the recursive call returns to us.  Omit the leading backslash if subkey is blank,
				// which supports recursively searching the contents of keys contained within a root key
				// (fixed for v1.0.17):
				sntprintf(subkey_full_path, _countof(subkey_full_path), _T("%s%s%s"), reg_item.subkey
					, *reg_item.subkey ? _T("\\") : _T(""), reg_item.name);
				// This section is very similar to the one in PerformLoopFilePattern(), so see it for comments:
				result = PerformLoopReg(aResultToken, aContinueMainLoop, aJumpToLine, aUntil
					, aFileLoopMode , aRecurseSubfolders, aRootKeyType, aRootKey, subkey_full_path);
				if (result != OK)
				{
					RegCloseKey(hRegKey);
					return result;
				}
				if (aContinueMainLoop || aJumpToLine)
					break;
			}
		}
		// else continue the loop in case some of the lower indexes can still be retrieved successfully.
		if (i == 0)  // Check this here due to it being an unsigned value that we don't want to go negative.
			break;
	}
	RegCloseKey(hRegKey);
	return OK;
}



ResultType Line::PerformLoopParse(ExprTokenType *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil)
{
	if (!*ARG2) // Since the input variable's contents are blank, the loop will execute zero times.
		return OK;

	// The following will be used to hold the parsed items.  It needs to have its own storage because
	// even though ARG2 might always be a writable memory area, we can't rely upon it being
	// persistent because it might reside in the deref buffer, in which case the other commands
	// in the loop's body would probably overwrite it.  Even if the ARG2's contents aren't in
	// the deref buffer, we still can't modify it (i.e. to temporarily terminate it and thus
	// bypass the need for malloc() below) because that might modify the variable contents, and
	// that variable may be referenced elsewhere in the body of the loop (which would result
	// in unexpected side-effects).  So, rather than have a limit of 64K or something (which
	// would limit this feature's usefulness for parsing a large list of filenames, for example),
	// it seems best to dynamically allocate a temporary buffer large enough to hold the
	// contents of ARG2 (the input variable).  Update: Since these loops tend to be enclosed
	// by file-read loops, and thus may be called thousands of times in a short period,
	// it should help average performance to use the stack for small vars rather than
	// constantly doing malloc() and free(), which are much higher overhead and probably
	// cause memory fragmentation (especially with thousands of calls):
	size_t space_needed = ArgLength(2) + 1;  // +1 for the zero terminator.
	LPTSTR stack_buf, buf;
	#define FREE_PARSE_MEMORY if (buf != stack_buf) free(buf)  // Also used by the CSV version of this function.
	#define LOOP_PARSE_BUF_SIZE 40000                          //
	if (space_needed <= LOOP_PARSE_BUF_SIZE)
	{
		stack_buf = (LPTSTR)talloca(space_needed); // Helps performance.  See comments above.
		buf = stack_buf;
	}
	else
	{
		if (   !(buf = tmalloc(space_needed))   )
			// Probably best to consider this a critical error, since on the rare times it does happen, the user
			// would probably want to know about it immediately.
			return LineError(ERR_OUTOFMEM, FAIL, ARG2);
		stack_buf = NULL; // For comparison purposes later below.
	}
	_tcscpy(buf, ARG2); // Make the copy.

	// Make a copy of ARG3 and ARG4 in case either one's contents are in the deref buffer, which would
	// probably be overwritten by the commands in the script loop's body:
	TCHAR delimiters[512], omit_list[512];
	tcslcpy(delimiters, ARG3, _countof(delimiters));
	tcslcpy(omit_list, ARG4, _countof(omit_list));

	ResultType result;
	Line *jump_to_line;
	TCHAR *field, *field_end, saved_char;
	size_t field_length;
	global_struct &g = *::g; // Primarily for performance in this case.

	for (field = buf;;)
	{ 
		if (*delimiters)
		{
			if (   !(field_end = StrChrAny(field, delimiters))   ) // No more delimiters found.
				field_end = field + _tcslen(field);  // Set it to the position of the zero terminator instead.
		}
		else // Since no delimiters, every char in the input string is treated as a separate field.
		{
			// But exclude this char if it's in the omit_list:
			if (*omit_list && _tcschr(omit_list, *field))
			{
				++field; // Move on to the next char.
				if (!*field) // The end of the string has been reached.
					break;
				continue;
			}
			field_end = field + 1;
		}

		saved_char = *field_end;  // In case it's a non-delimited list of single chars.
		*field_end = '\0';  // Temporarily terminate so that GetLoopField() will see the correct substring.

		if (*omit_list && *field && *delimiters)  // If no delimiters, the omit_list has already been handled above.
		{
			// Process the omit list.
			field = omit_leading_any(field, omit_list, field_end - field);
			if (*field) // i.e. the above didn't remove all the chars due to them all being in the omit-list.
			{
				field_length = omit_trailing_any(field, omit_list, field_end - 1);
				field[field_length] = '\0';  // Terminate here, but don't update field_end, since saved_char needs it.
			}
		}

		g.mLoopField = field;

		if (mNextLine->mActionType == ACT_BLOCK_BEGIN) // See PerformLoop() for comments about this section.
			do
				result = mNextLine->mNextLine->ExecUntil(UNTIL_BLOCK_END, aResultToken, &jump_to_line);
			while (jump_to_line == mNextLine);
		else
			result = mNextLine->ExecUntil(ONLY_ONE_LINE, aResultToken, &jump_to_line);

		if (jump_to_line && !(result == LOOP_CONTINUE && jump_to_line == this)) // See comments in PerformLoop() about this section.
		{
			if (jump_to_line == this)
				aContinueMainLoop = true;
			else
				aJumpToLine = jump_to_line; // Signal our caller to handle this jump.
			FREE_PARSE_MEMORY;
			return result;
		}
		if ( result != OK && result != LOOP_CONTINUE // i.e. result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL)
			|| (aUntil && aUntil->EvaluateLoopUntil(result)) )
		{
			FREE_PARSE_MEMORY;
			return result;
		}

		if (!saved_char) // The last item in the list has just been processed, so the loop is done.
			break;
		*field_end = saved_char;  // Undo the temporary termination, in case the list of delimiters is blank.
		field = *delimiters ? field_end + 1 : field_end;  // Move on to the next field.
		++g.mLoopIteration;
	}
	FREE_PARSE_MEMORY;
	return OK;
}



ResultType Line::PerformLoopParseCSV(ExprTokenType *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil)
// This function is similar to PerformLoopParse() so the two should be maintained together.
// See PerformLoopParse() for comments about the below (comments have been mostly stripped
// from this function).
{
	if (!*ARG2) // Since the input variable's contents are blank, the loop will execute zero times.
		return OK;

	// See comments in PerformLoopParse() for details.
	size_t space_needed = ArgLength(2) + 1;  // +1 for the zero terminator.
	LPTSTR stack_buf, buf;
	if (space_needed <= LOOP_PARSE_BUF_SIZE)
	{
		stack_buf = (LPTSTR)talloca(space_needed); // Helps performance.  See comments above.
		buf = stack_buf;
	}
	else
	{
		if (   !(buf = tmalloc(space_needed))   )
			return LineError(ERR_OUTOFMEM, FAIL, ARG2);
		stack_buf = NULL; // For comparison purposes later below.
	}
	_tcscpy(buf, ARG2); // Make the copy.

	TCHAR omit_list[512];
	tcslcpy(omit_list, ARG4, _countof(omit_list));

	ResultType result;
	Line *jump_to_line;
	TCHAR *field, *field_end, saved_char;
	size_t field_length;
	bool field_is_enclosed_in_quotes;
	global_struct &g = *::g; // Primarily for performance in this case.

	for (field = buf;;)
	{
		if (*field == '"')
		{
			// For each field, check if the optional leading double-quote is present.  If it is,
			// skip over it since we always know it's the one that marks the beginning of
			// the that field.  This assumes that a field containing escaped double-quote is
			// always contained in double quotes, which is how Excel does it.  For example:
			// """string with escaped quotes""" resolves to a literal quoted string:
			field_is_enclosed_in_quotes = true;
			++field;
		}
		else
			field_is_enclosed_in_quotes = false;

		for (field_end = field;;)
		{
			if (   !(field_end = _tcschr(field_end, field_is_enclosed_in_quotes ? '"' : ','))   )
			{
				// This is the last field in the string, so set field_end to the position of
				// the zero terminator instead:
				field_end = field + _tcslen(field);
				break;
			}
			if (field_is_enclosed_in_quotes)
			{
				// The quote discovered above marks the end of the string if it isn't followed
				// by another quote.  But if it is a pair of quotes, replace it with a single
				// literal double-quote and then keep searching for the real ending quote:
				if (field_end[1] == '"')  // A pair of quotes was encountered.
				{
					tmemmove(field_end, field_end + 1, _tcslen(field_end + 1) + 1); // +1 to include terminator.
					++field_end; // Skip over the literal double quote that we just produced.
					continue; // Keep looking for the "real" ending quote.
				}
				// Otherwise, this quote marks the end of the field, so just fall through and break.
			}
			// else field is not enclosed in quotes, so the comma discovered above must be a delimiter.
			break;
		}

		saved_char = *field_end; // This can be the terminator, a comma, or a double-quote.
		*field_end = '\0';  // Terminate here so that GetLoopField() will see the correct substring.

		if (*omit_list && *field)
		{
			// Process the omit list.
			field = omit_leading_any(field, omit_list, field_end - field);
			if (*field) // i.e. the above didn't remove all the chars due to them all being in the omit-list.
			{
				field_length = omit_trailing_any(field, omit_list, field_end - 1);
				field[field_length] = '\0';  // Terminate here, but don't update field_end, since we need its pos.
			}
		}

		g.mLoopField = field;

		if (mNextLine->mActionType == ACT_BLOCK_BEGIN) // See PerformLoop() for comments about this section.
			do
				result = mNextLine->mNextLine->ExecUntil(UNTIL_BLOCK_END, aResultToken, &jump_to_line);
			while (jump_to_line == mNextLine);
		else
			result = mNextLine->ExecUntil(ONLY_ONE_LINE, aResultToken, &jump_to_line);

		if (jump_to_line && !(result == LOOP_CONTINUE && jump_to_line == this)) // See comments in PerformLoop() about this section.
		{
			if (jump_to_line == this)
				aContinueMainLoop = true;
			else
				aJumpToLine = jump_to_line; // Signal our caller to handle this jump.
			FREE_PARSE_MEMORY;
			return result;
		}
		if ( result != OK && result != LOOP_CONTINUE // i.e. result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL)
			|| (aUntil && aUntil->EvaluateLoopUntil(result)) )
		{
			FREE_PARSE_MEMORY;
			return result;
		}

		if (!saved_char) // The last item in the list has just been processed, so the loop is done.
			break;
		if (saved_char == ',') // Set "field" to be the position of the next field.
			field = field_end + 1;
		else // saved_char must be a double-quote char.
		{
			field = field_end + 1;
			if (!*field) // No more fields occur after this one.
				break;
			// Find the next comma, which must be a real delimiter since we're in between fields:
			if (   !(field = _tcschr(field, ','))   ) // No more fields.
				break;
			// Set it to be the first character of the next field, which might be a double-quote
			// or another comma (if the field is empty).
			++field;
		}
		++g.mLoopIteration;
	}
	FREE_PARSE_MEMORY;
	return OK;
}



ResultType Line::PerformLoopReadFile(ExprTokenType *aResultToken, bool &aContinueMainLoop, Line *&aJumpToLine, Line *aUntil
	, TextStream *aReadFile, LPTSTR aWriteFileName)
{
	LoopReadFileStruct loop_info(aReadFile, aWriteFileName);
	size_t line_length;
	ResultType result;
	Line *jump_to_line;
	global_struct &g = *::g; // Primarily for performance in this case.

	for (;; ++g.mLoopIteration)
	{ 
		if (  !(line_length = loop_info.mReadFile->ReadLine(loop_info.mCurrentLine, _countof(loop_info.mCurrentLine) - 1))  ) // -1 to ensure there's room for a null-terminator.
		{
			// We want to return OK except in some specific cases handled below (see "break").
			result = OK;
			break;
		}
		if (loop_info.mCurrentLine[line_length - 1] == '\n') // Remove newlines like FileReadLine does.
			--line_length;
		loop_info.mCurrentLine[line_length] = '\0';
		g.mLoopReadFile = &loop_info;
		if (mNextLine->mActionType == ACT_BLOCK_BEGIN) // See PerformLoop() for comments about this section.
			do
				result = mNextLine->mNextLine->ExecUntil(UNTIL_BLOCK_END, aResultToken, &jump_to_line);
			while (jump_to_line == mNextLine);
		else
			result = mNextLine->ExecUntil(ONLY_ONE_LINE, aResultToken, &jump_to_line);
		if (jump_to_line && !(result == LOOP_CONTINUE && jump_to_line == this)) // See comments in PerformLoop() about this section.
		{
			if (jump_to_line == this)
				aContinueMainLoop = true;
			else
				aJumpToLine = jump_to_line; // Signal our caller to handle this jump.
			break;
		}
		if (result != OK && result != LOOP_CONTINUE) // i.e. result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL)
			break;
		if (aUntil && aUntil->EvaluateLoopUntil(result))
			break;
	}

	if (loop_info.mWriteFile)
	{
		loop_info.mWriteFile->Close();
		delete loop_info.mWriteFile;
	}

	return result;
}



__forceinline ResultType Line::Perform() // As of 2/9/2009, __forceinline() reduces code size a little (since this function is called from only one place) and boosts performance a bit, though it's probably more due to the butterfly effect and cache hits/misses.
// Performs only this line's action.
// Returns OK or FAIL.
// The function should not be called to perform any flow-control actions such as
// Goto, Gosub, Return, Block-Begin, Block-End, If, Else, etc.
{
	TCHAR buf_temp[MAX_REG_ITEM_SIZE], *contents; // For registry and other things.
	WinGroup *group; // For the group commands.
	Var *arg_var2, *output_var = OUTPUT_VAR; // Okay if NULL. Users of it should only consider it valid if their first arg is actually an output_variable.
	global_struct &g = *::g; // Reduces code size due to replacing so many g-> with g. Eclipsing ::g with local g makes compiler remind/enforce the use of the right one.
	BOOL arg2_has_binary_integer;
	ToggleValueType toggle;  // For commands that use on/off/neutral.
	// Use signed values for these in case they're really given an explicit negative value:
	int start_char_num, chars_to_extract; // For String commands.
	size_t source_length; // For String commands.
	SymbolType var_is_pure_numeric, value_is_pure_numeric; // For math operations.
	vk_type vk; // For GetKeyState.
	__int64 device_id;  // For sound commands.  __int64 helps avoid compiler warning for some conversions.
	bool is_remote_registry; // For Registry commands.
	HKEY root_key; // For Registry commands.
	LPTSTR subkey, value_name, value;
	ResultType result;  // General purpose.

	// Even though the loading-parser already checked, check again, for now,
	// at least until testing raises confidence.  UPDATE: Don't this because
	// sometimes (e.g. ACT_ASSIGN/ADD/SUB/MULT/DIV) the number of parameters
	// required at load-time is different from that at runtime, because params
	// are taken out or added to the param list:
	//if (nArgs < g_act[mActionType].MinParams) ...

	switch (mActionType)
	{
	case ACT_ASSIGN:
		// Note: This line's args have not yet been dereferenced in this case (i.e. ExpandArgs() hasn't been
		// called).  The below function will handle that if it is needed.
		return PerformAssign();  // It will report any errors for us.

	case ACT_ASSIGNEXPR:
		// Currently, ACT_ASSIGNEXPR can occur even when mArg[1].is_expression==false, such as things like var:=5
		// and var:=Array%i%.  Search on "is_expression = " to find such cases in the script-loading/parsing
		// routines.
		if (mArgc > 1)
		{
			if (mArg[1].is_expression) // v1.0.45: ExpandExpression() already took care of it for us (for performance reasons).
				return OK;

			// Above must be checked prior to below since each uses "postfix" in a different way.
			if (mArg[1].postfix) // There is a cached binary integer.
				return output_var->Assign(*(__int64 *)mArg[1].postfix);

			// sArgVar is used to enhance performance, which would otherwise be poor for dynamic variables
			// such as Var:=Array%i% (which is an expression and handled by ACT_ASSIGNEXPR rather than
			// ACT_ASSIGN) because Array%i% would have to be resolved twice (once here and once
			// previously by ExpandArgs()) just to find out if it's IsBinaryClip()).
			// If ARG2 isn't blank, this ACT_ASSIGNEXPR is assigning an environment variable or g_ErrorLevel
			// (e.g. var:=Username); so can't apply this optimization.
			if (ARGVARRAW2 && !*ARG2) // See above.  Also, RAW is safe due to the above check of mArgc > 1.
			{
				switch(ARGVARRAW2->Type())
				{
				case VAR_NORMAL: // This can be reached via things like: x:=single_naked_var_including_binary_clip
					// Assign var to var in case ARGVARRAW2->IsBinaryClip(), and for others because
					// var-to-var has optimizations like retaining the copying over the cached binary number.
					// In the case of ARGVARRAW2->IsBinaryClip(), performance should be good since
					// IsBinaryClip() implies a single isolated deref, which would never have been copied
					// into the deref buffer.
					//
					// v1.0.46.01: ARGVARRAW2->IsBinaryClip() can be true because loadtime no longer translates
					// such statements into ACT_ASSIGN vs. ACT_ASSIGNEXPR.  Even without that change, it can also
					// be reached by something like:
					//    DynClipboardAll = ClipboardAll
					//    ClipSaved := %DynClipboardAll%
					return output_var->Assign(*ARGVARRAW2); // Var-to-var copy supports ARGVARRAW2 being binary clipboard, and also exploits caching of binary numbers, for performance.
				case VAR_CLIPBOARDALL:
					return output_var->AssignClipboardAll();
				//Otherwise it's VAR_CLIPBOARD or a read-only variable; continue on to do assign the normal way.
				}
			}
		}
		// Since above didn't return:
		// Note that simple assignments such as Var:="xyz" or Var:=Var2 are resolved to be
		// non-expressions at load-time.  In these cases, ARG2 would have been expanded
		// normally rather than evaluated as an expression.
		return output_var->Assign(ARG2); // ARG2 now contains the above or the evaluated result of the expression.

	case ACT_EXPRESSION:
		// Nothing needs to be done because the expression in ARG1 (which is the only arg) has already
		// been evaluated and its functions and subfunctions called.  Examples:
		//    fn(123, "string", var, fn2(y))
		//    x&=3
		//    var ? func() : x:=y
		return OK;

	// Like AutoIt2, if either output_var or ARG1 aren't purely numeric, they
	// will be considered to be zero for all of the below math functions:
	case ACT_ADD:
		// Notes about the macro below:
		// Ordered for short-circuit performance. No need to check if it's g_ErrorLevel (like
		// ArgMustBeDereferenced() does) because the commands that use it don't internally change ErrorLevel.
		// RAW is safe because loadtime validation ensured there are at least 2 args.
		// ACT_ADD/SUB/MULT/DIV are one of the few places that pass true to IsNonBlankIntegerOrFloat(true).
		// This is for backward compatibility.
		#define DEFINE_ARG_VAR2 arg_var2 = (ARGVARRAW2 && !*ARG2 && ARGVARRAW2->Type() == VAR_NORMAL) ? ARGVARRAW2 : NULL;
		#undef DETERMINE_NUMERIC_TYPES
		#define DETERMINE_NUMERIC_TYPES \
			if (arg2_has_binary_integer = mArg[1].postfix && !mArg[1].is_expression)\
			{\
				value_is_pure_numeric = PURE_INTEGER;\
				arg_var2 = NULL;\
			}\
			else\
			{\
				DEFINE_ARG_VAR2 \
				value_is_pure_numeric = arg_var2 ? arg_var2->IsNonBlankIntegerOrFloat(true)\
					: IsPureNumeric(ARG2, true, false, true, true);\
			}\
			var_is_pure_numeric = output_var->IsNonBlankIntegerOrFloat(true);
		#undef ARG2_AS_DOUBLE
		#undef ARG2_AS_INT64
		#define ARG2_AS_DOUBLE (arg2_has_binary_integer ? (double)*(__int64*)mArg[1].postfix \
			: arg_var2 ? arg_var2->ToDouble(FALSE) : ATOF(ARG2))
		#define ARG2_AS_INT64 (arg2_has_binary_integer ? *(__int64*)mArg[1].postfix \
			: (arg_var2 ? arg_var2->ToInt64(FALSE) : ATOI64(ARG2)))

		// Some performance can be gained by relying on the fact that short-circuit boolean
		// can skip the "var_is_pure_numeric" check whenever value_is_pure_numeric == PURE_FLOAT.
		// This is because var_is_pure_numeric is never directly needed here (unlike EvaluateCondition()).
		// However, benchmarks show that this makes such a small difference that it's not worth the
		// loss of maintainability and the slightly larger code size due to macro expansion:
		//#undef IF_EITHER_IS_FLOAT
		//#define IF_EITHER_IS_FLOAT if (value_is_pure_numeric == PURE_FLOAT \
		//	|| IsPureNumeric(output_var->Contents(), true, false, true, true) == PURE_FLOAT)

		DETERMINE_NUMERIC_TYPES

		if (!*ARG3 || !_tcschr(_T("SMHD"), ctoupper(*ARG3))) // ARG3 is absent or invalid, so do normal math (not date-time).
		{
			IF_EITHER_IS_FLOAT
				return output_var->Assign(output_var->ToDouble(FALSE) + ARG2_AS_DOUBLE);
			else // Non-numeric variables or values are considered to be zero for the purpose of the calculation.
				return output_var->Assign(output_var->ToInt64(FALSE) + ARG2_AS_INT64);
		}

		// Since above didn't return, the command is being used to add a value to a date-time.
		if (!value_is_pure_numeric) // It's considered to be zero, so the output_var is left unchanged:
			return OK;
		// Since above didn't return:
		// Use double to support a floating point value for days, hours, minutes, etc:
		double nUnits; // Declaring separate from initializing avoids compiler warning when not inside a block.
		nUnits = ARG2_AS_DOUBLE;
		FILETIME ft, ftNowUTC;
		if (*output_var->Contents())
		{
			if (!YYYYMMDDToFileTime(output_var->Contents(), ft))
				return output_var->Assign(_T("")); // Set to blank to indicate the problem.
		}
		else // The output variable is currently blank, so substitute the current time for it.
		{
			GetSystemTimeAsFileTime(&ftNowUTC);
			FileTimeToLocalFileTime(&ftNowUTC, &ft);  // Convert UTC to local time.
		}
		// Convert to 10ths of a microsecond (the units of the FILETIME struct):
		switch (ctoupper(*ARG3))
		{
		case 'S': // Seconds
			nUnits *= (double)10000000;
			break;
		case 'M': // Minutes
			nUnits *= ((double)10000000 * 60);
			break;
		case 'H': // Hours
			nUnits *= ((double)10000000 * 60 * 60);
			break;
		case 'D': // Days
			nUnits *= ((double)10000000 * 60 * 60 * 24);
			break;
		}
		// Convert ft struct to a 64-bit variable (maybe there's some way to avoid these conversions):
		ULARGE_INTEGER ul;
		ul.LowPart = ft.dwLowDateTime;
		ul.HighPart = ft.dwHighDateTime;
		// Add the specified amount of time to the result value:
		ul.QuadPart += (__int64)nUnits;  // Seems ok to cast/truncate in light of the *=10000000 above.
		// Convert back into ft struct:
		ft.dwLowDateTime = ul.LowPart;
		ft.dwHighDateTime = ul.HighPart;
		return output_var->Assign(FileTimeToYYYYMMDD(buf_temp, ft, false));

	case ACT_SUB:
		if (!*ARG3 || !_tcschr(_T("SMHD"), ctoupper(*ARG3))) // ARG3 is absent or invalid, so do normal math (not date-time).
		{
			DETERMINE_NUMERIC_TYPES
			IF_EITHER_IS_FLOAT
				return output_var->Assign(output_var->ToDouble(FALSE) - ARG2_AS_DOUBLE);
			else // Non-numeric variables or values are considered to be zero for the purpose of the calculation.
				return output_var->Assign(output_var->ToInt64(FALSE) - ARG2_AS_INT64);
		}

		// Since above didn't return, the command is being used to subtract date-time values.
		bool failed;
		// If either ARG2 or output_var->Contents() is blank, it will default to the current time:
		__int64 time_until; // Declaring separate from initializing avoids compiler warning when not inside a block.
		DEFINE_ARG_VAR2
		time_until = YYYYMMDDSecondsUntil(arg_var2 ? arg_var2->Contents() : ARG2
			, output_var->Contents(), failed);
		if (failed) // Usually caused by an invalid component in the date-time string.
			return output_var->Assign(_T(""));
		switch (ctoupper(*ARG3))
		{
		// Do nothing in the case of 'S' (seconds).  Otherwise:
		case 'M': time_until /= 60; break; // Minutes
		case 'H': time_until /= 60 * 60; break; // Hours
		case 'D': time_until /= 60 * 60 * 24; break; // Days
		}
		// Only now that any division has been performed (to reduce the magnitude of
		// time_until) do we cast down into an int, which is the standard size
		// used for non-float results (the result is always non-float for subtraction
		// of two date-times):
		return output_var->Assign(time_until); // Assign as signed 64-bit.

	case ACT_MULT:
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_FLOAT
			return output_var->Assign(output_var->ToDouble(FALSE) * ARG2_AS_DOUBLE);
		else // Non-numeric variables or values are considered to be zero for the purpose of the calculation.
			return output_var->Assign(output_var->ToInt64(FALSE) * ARG2_AS_INT64);

	case ACT_DIV:
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_FLOAT
		{
			double ARG2_as_float = ARG2_AS_DOUBLE;
			if (!ARG2_as_float)              // v1.0.46: Make behavior more consistent with expressions by
				return output_var->Assign(); // avoiding a runtime error dialog; just make the output variable blank.
			return output_var->Assign(output_var->ToDouble(FALSE) / ARG2_as_float);
		}
		else // Non-numeric variables or values are considered to be zero for the purpose of the calculation.
		{
			__int64 ARG2_as_int = ARG2_AS_INT64;
			if (!ARG2_as_int)                // v1.0.46: Make behavior more consistent with expressions by
				return output_var->Assign(); // avoiding a runtime error dialog; just make the output variable blank.
			return output_var->Assign(output_var->ToInt64(FALSE) / ARG2_as_int);
		}

	case ACT_STRINGLEFT:
		chars_to_extract = ArgToInt(3); // Use 32-bit signed to detect negatives and fit it VarSizeType.
		if (chars_to_extract < 0)
			// For these we don't report an error, since it might be intentional for
			// it to be called this way, in which case it will do nothing other than
			// set the output var to be blank.
			chars_to_extract = 0;
		else
		{
			source_length = ArgLength(2); // Should be quick because Arg2 is an InputVar (except when it's a built-in var perhaps).
			if (chars_to_extract > (int)source_length)
				chars_to_extract = (int)source_length; // Assign() requires a length that's <= the actual length of the string.
		}
		// It will display any error that occurs.
		return output_var->Assign(ARG2, chars_to_extract);

	case ACT_STRINGRIGHT:
		chars_to_extract = ArgToInt(3); // Use 32-bit signed to detect negatives and fit it VarSizeType.
		if (chars_to_extract < 0)
			chars_to_extract = 0;
		source_length = ArgLength(2);
		if ((UINT)chars_to_extract > source_length)
			chars_to_extract = (int)source_length;
		// It will display any error that occurs:
		return output_var->Assign(ARG2 + source_length - chars_to_extract, chars_to_extract);

	case ACT_STRINGMID:
		// v1.0.43.10: Allow chars-to-extract to be blank, which means "get all characters".
		// However, for backward compatibility, examine the raw arg, not ARG4.  That way, any existing
		// scripts that use a variable reference or expression that resolves to an empty string will
		// have the parameter treated as zero (as in previous versions) rather than "all characters".
		if (mArgc < 4 || !*mArg[3].text)
			chars_to_extract = INT_MAX;
		else
		{
			chars_to_extract = ArgToInt(4); // Use 32-bit signed to detect negatives and fit it VarSizeType.
			if (chars_to_extract < 1)
				return output_var->Assign();  // Set it to be blank in this case.
		}
		start_char_num = ArgToInt(3);
		if (ctoupper(*ARG5) == 'L')  // Chars to the left of start_char_num will be extracted.
		{
			// TRANSLATE "L" MODE INTO THE EQUIVALENT NORMAL MODE:
			if (start_char_num < 1) // Starting at a character number that is invalid for L mode.
				return output_var->Assign();  // Blank seems most appropriate for the L option in this case.
			start_char_num -= (chars_to_extract - 1);
			if (start_char_num < 1)
				// Reduce chars_to_extract to reflect the fact that there aren't enough chars
				// to the left of start_char_num, so we'll extract only them:
				chars_to_extract -= (1 - start_char_num);
		}
		// ABOVE HAS CONVERTED "L" MODE INTO NORMAL MODE, so "L" no longer needs to be considered below.
		// UPDATE: The below is also needed for the L option to work correctly.  Older:
		// It's somewhat debatable, but it seems best not to report an error in this and
		// other cases.  The result here is probably enough to speak for itself, for script
		// debugging purposes:
		if (start_char_num < 1)
			start_char_num = 1; // 1 is the position of the first char, unlike StringGetPos.
		source_length = ArgLength(2); // This call seems unavoidable in both "L" mode and normal mode.
		if (source_length < (UINT)start_char_num) // Source is empty or start_char_num lies to the right of the entire string.
			return output_var->Assign(); // No chars exist there, so set it to be blank.
		source_length -= (start_char_num - 1); // Fix for v1.0.44.14: Adjust source_length to be the length starting at start_char_num.  Otherwise, the length passed to Assign() could be too long, and it now expects an accurate length.
		if ((UINT)chars_to_extract > source_length)
			chars_to_extract = (int)source_length;
		return output_var->Assign(ARG2 + start_char_num - 1, chars_to_extract);

	case ACT_STRINGTRIMLEFT:
		chars_to_extract = ArgToInt(3); // Use 32-bit signed to detect negatives and fit it VarSizeType.
		if (chars_to_extract < 0)
			chars_to_extract = 0;
		source_length = ArgLength(2);
		if ((UINT)chars_to_extract > source_length) // This could be intentional, so don't display an error.
			chars_to_extract = (int)source_length;
		return output_var->Assign(ARG2 + chars_to_extract, (VarSizeType)(source_length - chars_to_extract));

	case ACT_STRINGTRIMRIGHT:
		chars_to_extract = ArgToInt(3); // Use 32-bit signed to detect negatives and fit it VarSizeType.
		if (chars_to_extract < 0)
			chars_to_extract = 0;
		source_length = ArgLength(2);
		if ((UINT)chars_to_extract > source_length) // This could be intentional, so don't display an error.
			chars_to_extract = (int)source_length;
		return output_var->Assign(ARG2, (VarSizeType)(source_length - chars_to_extract)); // It already displayed any error.

	case ACT_STRINGLOWER:
	case ACT_STRINGUPPER:
		contents = output_var->Contents(TRUE, TRUE); // Set default.	
		if (contents != ARG2 || output_var->Type() != VAR_NORMAL) // It's compared this way in case ByRef/aliases are involved.  This will detect even them.
		{
			// Clipboard is involved and/or source != dest.  Do it the more comprehensive way.
			// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
			// this call will set up the clipboard for writing.
			// Fix for v1.0.45.02: The v1.0.45 change where the value is assigned directly without sizing the
			// variable first doesn't work in cases when the variable is the clipboard.  This is because the
			// clipboard's buffer is changeable (for the case conversion later below) only when using the following
			// approach, not a simple "assign then modify its Contents()".
			if (output_var->AssignString(NULL, ArgLength(2)) != OK) // The length is in characters, so AssignString(NULL, ...)
				return FAIL;
			contents = output_var->Contents(); // Do this only after the above might have changed the contents mem address.
			// Copy the input variable's text directly into the output variable:
			_tcscpy(contents, ARG2);
		}
		//else input and output are the same, normal variable; so nothing needs to be copied over.  Just leave
		// contents at the default set earlier, then convert its case.
		if (*ARG3 && ctoupper(*ARG3) == 'T' && !*(ARG3 + 1)) // Convert to title case.
			StrToTitleCase(contents);
		else if (mActionType == ACT_STRINGLOWER)
			CharLower(contents);
		else
			CharUpper(contents);
		return output_var->Close();  // In case it's the clipboard.

	case ACT_STRINGLEN:
		return output_var->Assign((__int64)(ARGVARRAW2 && ARGVARRAW2->IsBinaryClip() // Load-time validation has ensured mArgc > 1.
			? ARGVARRAW2->Length() // Total size of the binary clip.
			: ArgLength(2)));
		// The above must be kept in sync with the StringLen() function elsewhere.

	case ACT_STRINGGETPOS:
	{
		LPTSTR arg4 = ARG4;
		int pos = -1; // Set default.
		int occurrence_number;
		if (*arg4 && _tcschr(_T("LR"), ctoupper(*arg4)))
			occurrence_number = *(arg4 + 1) ? ATOI(arg4 + 1) : 1;
		else
			occurrence_number = 1;
		// Intentionally allow occurrence_number to resolve to a negative, for scripting flexibility:
		if (occurrence_number > 0)
		{
			if (!*ARG3) // It might be intentional, in obscure cases, to search for the empty string.
				pos = 0;
				// Above: empty string is always found immediately (first char from left) regardless
				// of whether the search will be conducted from the right.  This is because it's too
				// rare to worry about giving it any more explicit handling based on search direction.
			else
			{
				LPTSTR found, haystack = ARG2, needle = ARG3;
				int offset = ArgToInt(5); // v1.0.30.03
				if (offset < 0)
					offset = 0;
				size_t haystack_length = ArgLength(2);
				if (offset < (int)haystack_length)
				{
					if (*arg4 == '1' || ctoupper(*arg4) == 'R') // Conduct the search starting at the right side, moving leftward.
					{
						// Want it to behave like in this example: If searching for the 2nd occurrence of
						// FF in the string FFFF, it should find the first two F's, not the middle two:
						found = tcsrstr(haystack, haystack_length - offset, needle, (StringCaseSenseType)g.StringCaseSense, occurrence_number);
					}
					else
					{
						// Want it to behave like in this example: If searching for the 2nd occurrence of
						// FF in the string FFFF, it should find position 3 (the 2nd pair), not position 2:
						size_t needle_length = ArgLength(3);
						int i;
						for (i = 1, found = haystack + offset; ; ++i, found += needle_length)
							if (!(found = g_tcsstr(found, needle)) || i == occurrence_number)
								break;
					}
					if (found)
						pos = (int)(found - haystack);
					// else leave pos set to its default value, -1.
				}
				//else offset >= haystack_length, so no match is possible in either left or right mode.
			}
		}
		g_ErrorLevel->Assign(pos < 0 ? ERRORLEVEL_ERROR : ERRORLEVEL_NONE);
		return output_var->Assign(pos); // Assign() already displayed any error that may have occurred.
	}

	case ACT_STRINGREPLACE:
		return StringReplace();

	case ACT_TRANSFORM:
		return Transform(ARG2, ARG3, ARG4);

	case ACT_STRINGSPLIT:
		return StringSplit(ARG1, ARG2, ARG3, ARG4);

	case ACT_SPLITPATH:
		return SplitPath(ARG1);

	case ACT_SORT:
		return PerformSort(ARG1, ARG2);

	case ACT_PIXELSEARCH:
		// ArgToInt() works on ARG7 (the color) because any valid BGR or RGB color has 0x00 in the high order byte:
		return PixelSearch(ArgToInt(3), ArgToInt(4), ArgToInt(5), ArgToInt(6), ArgToInt(7), ArgToInt(8), ARG9, false);
	case ACT_IMAGESEARCH:
		return ImageSearch(ArgToInt(3), ArgToInt(4), ArgToInt(5), ArgToInt(6), ARG7);
	case ACT_PIXELGETCOLOR:
		return PixelGetColor(ArgToInt(2), ArgToInt(3), ARG4);

	case ACT_SEND:
	case ACT_SENDRAW:
		SendKeys(ARG1, mActionType == ACT_SENDRAW ? SCM_RAW : SCM_NOT_RAW, g.SendMode);
		return OK;
	case ACT_SENDINPUT: // Raw mode is supported via {Raw} in ARG1.
		SendKeys(ARG1, SCM_NOT_RAW, g.SendMode == SM_INPUT_FALLBACK_TO_PLAY ? SM_INPUT_FALLBACK_TO_PLAY : SM_INPUT);
		return OK;
	case ACT_SENDPLAY: // Raw mode is supported via {Raw} in ARG1.
		SendKeys(ARG1, SCM_NOT_RAW, SM_PLAY);
		return OK;
	case ACT_SENDEVENT:
		SendKeys(ARG1, SCM_NOT_RAW, SM_EVENT);
		return OK;

	case ACT_CLICK:
		return PerformClick(ARG1);
	case ACT_MOUSECLICKDRAG:
		return PerformMouse(mActionType, SEVEN_ARGS);
	case ACT_MOUSECLICK:
		return PerformMouse(mActionType, THREE_ARGS, _T(""), _T(""), ARG5, ARG7, ARG4, ARG6);
	case ACT_MOUSEMOVE:
		return PerformMouse(mActionType, _T(""), ARG1, ARG2, _T(""), _T(""), ARG3, ARG4);

	case ACT_MOUSEGETPOS:
		return MouseGetPos(ArgToUInt(5));

	case ACT_WINACTIVATE:
	case ACT_WINACTIVATEBOTTOM:
		if (WinActivate(g, FOUR_ARGS, mActionType == ACT_WINACTIVATEBOTTOM))
			// It seems best to do these sleeps here rather than in the windowing
			// functions themselves because that way, the program can use the
			// windowing functions without being subject to the script's delay
			// setting (i.e. there are probably cases when we don't need
			// to wait, such as bringing a message box to the foreground,
			// since no other actions will be dependent on it actually
			// having happened:
			DoWinDelay;
		return OK;

	case ACT_WINMINIMIZE:
	case ACT_WINMAXIMIZE:
	case ACT_WINRESTORE:
	case ACT_WINHIDE:
	case ACT_WINSHOW:
	case ACT_WINCLOSE:
	case ACT_WINKILL:
	{
		// Set initial guess for is_ahk_group (further refined later).  For ahk_group, WinText,
		// ExcludeTitle, and ExcludeText must be blank so that they are reserved for future use
		// (i.e. they're currently not supported since the group's own criteria take precedence):
		bool is_ahk_group = !(_tcsnicmp(ARG1, _T("ahk_group"), 9) || *ARG2 || *ARG4);
		// The following is not quite accurate since is_ahk_group is only a guess at this stage, but
		// given the extreme rarity of the guess being wrong, this shortcut seems justified to reduce
		// the code size/complexity.  A wait_time of zero seems best for group closing because it's
		// currently implemented to do the wait after every window in the group.  In addition,
		// this makes "WinClose ahk_group GroupName" behave identically to "GroupClose GroupName",
		// which seems best, for consistency:
		int wait_time = is_ahk_group ? 0 : DEFAULT_WINCLOSE_WAIT;
		if (mActionType == ACT_WINCLOSE || mActionType == ACT_WINKILL) // ARG3 is contains the wait time.
		{
			if (*ARG3 && !(wait_time = (int)(1000 * ArgToDouble(3)))   )
				wait_time = 500; // Legacy (prior to supporting floating point): 0 is defined as 500ms, which seems more useful than a true zero.
			if (*ARG5)
				is_ahk_group = false;  // Override the default.
		}
		else
			if (*ARG3)
				is_ahk_group = false;  // Override the default.
		// Act upon all members of this group (WinText/ExcludeTitle/ExcludeText are ignored in this mode).
		if (is_ahk_group && (group = g_script.FindGroup(omit_leading_whitespace(ARG1 + 9)))) // Assign.
			return group->ActUponAll(mActionType, wait_time); // It will do DoWinDelay if appropriate.
		//else try to act upon it as though "ahk_group something" is a literal window title.
	
		// Since above didn't return, it's not "ahk_group", so do the normal single-window behavior.
		if (mActionType == ACT_WINCLOSE || mActionType == ACT_WINKILL)
		{
			if (WinClose(g, ARG1, ARG2, wait_time, ARG4, ARG5, mActionType == ACT_WINKILL)) // It closed something.
				DoWinDelay;
			return OK;
		}
		else
			return PerformShowWindow(mActionType, FOUR_ARGS);
	}

	case ACT_ENVGET:
		return EnvGet(ARG2);

	case ACT_ENVSET:
		// MSDN: "If [the 2nd] parameter is NULL, the variable is deleted from the current process’s environment."
		// My: Though it seems okay, for now, just to set it to be blank if the user omitted the 2nd param or
		// left it blank (AutoIt3 does this too).  Also, no checking is currently done to ensure that ARG2
		// isn't longer than 32K, since future OSes may support longer env. vars.  SetEnvironmentVariable()
		// might return 0(fail) in that case anyway.  Also, ARG1 may be a dereferenced variable that resolves
		// to the name of an Env. Variable.  In any case, this name need not correspond to any existing
		// variable name within the script (i.e. script variables and env. variables aren't tied to each other
		// in any way).  This seems to be the most flexible approach, but are there any shortcomings?
		// The only one I can think of is that if the script tries to fetch the value of an env. var (perhaps
		// one that some other spawned program changed), and that var's name corresponds to the name of a
		// script var, the script var's value (if non-blank) will be fetched instead.
		// Note: It seems, at least under WinXP, that env variable names can contain spaces.  So it's best
		// not to validate ARG1 the same way we validate script variables (i.e. just let\
		// SetEnvironmentVariable()'s return value determine whether there's an error).  However, I just
		// realized that it's impossible to "retrieve" the value of an env var that has spaces (until now,
		// since the EnvGet command is available).
		return SetErrorLevelOrThrowBool(!SetEnvironmentVariable(ARG1, ARG2));

	case ACT_ENVUPDATE:
	{
		// From the AutoIt3 source:
		// AutoIt3 uses SMTO_BLOCK (which prevents our thread from doing anything during the call)
		// vs. SMTO_NORMAL.  Since I'm not sure why, I'm leaving it that way for now:
		ULONG_PTR nResult;
		return SetErrorLevelOrThrowBool(!SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE,
										0, (LPARAM)_T("Environment"), SMTO_BLOCK, 15000, &nResult));
	}

	case ACT_URLDOWNLOADTOFILE:
		return URLDownloadToFile(TWO_ARGS);

	case ACT_RUNAS:
		if (!g_os.IsWin2000orLater()) // Do nothing if the OS doesn't support it.
			return OK;
		StringTCharToWChar(ARG1, g_script.mRunAsUser);
		StringTCharToWChar(ARG2, g_script.mRunAsPass);
		StringTCharToWChar(ARG3, g_script.mRunAsDomain);
		return OK;

	case ACT_RUN:
	{
		bool use_el = tcscasestr(ARG3, _T("UseErrorLevel"));
		result = g_script.ActionExec(ARG1, NULL, ARG2, !use_el, ARG3, NULL, use_el, true, ARGVAR4); // Be sure to pass NULL for 2nd param.
		if (use_el)
			// The special string ERROR is used, rather than a number like 1, because currently
			// RunWait might in the future be able to return any value, including 259 (STATUS_PENDING).
			result = g_ErrorLevel->Assign(result ? ERRORLEVEL_NONE : _T("ERROR"));
		// Otherwise, if result == FAIL, above already displayed the error (or threw an exception).
		return result;
	}

	case ACT_RUNWAIT:
	case ACT_CLIPWAIT:
	case ACT_KEYWAIT:
	case ACT_WINWAIT:
	case ACT_WINWAITCLOSE:
	case ACT_WINWAITACTIVE:
	case ACT_WINWAITNOTACTIVE:
		return PerformWait();

	case ACT_WINMOVE:
		return mArgc > 2 ? WinMove(EIGHT_ARGS) : WinMove(_T(""), _T(""), ARG1, ARG2);

	case ACT_WINMENUSELECTITEM:
		return WinMenuSelectItem(ELEVEN_ARGS);

	case ACT_CONTROLSEND:
	case ACT_CONTROLSENDRAW:
		return ControlSend(SIX_ARGS, mActionType == ACT_CONTROLSENDRAW ? SCM_RAW : SCM_NOT_RAW);

	case ACT_CONTROLCLICK:
		if (   !(vk = ConvertMouseButton(ARG4))   ) // Treats blank as "Left".
			return LineError(ERR_PARAM4_INVALID, FAIL, ARG4);
		return ControlClick(vk, *ARG5 ? ArgToInt(5) : 1, ARG6, ARG1, ARG2, ARG3, ARG7, ARG8);

	case ACT_CONTROLMOVE:
		return ControlMove(NINE_ARGS);
	case ACT_CONTROLGETPOS:
		return ControlGetPos(ARG5, ARG6, ARG7, ARG8, ARG9);
	case ACT_CONTROLGETFOCUS:
		return ControlGetFocus(ARG2, ARG3, ARG4, ARG5);
	case ACT_CONTROLFOCUS:
		return ControlFocus(FIVE_ARGS);
	case ACT_CONTROLSETTEXT:
		return ControlSetText(SIX_ARGS);
	case ACT_CONTROLGETTEXT:
		return ControlGetText(ARG2, ARG3, ARG4, ARG5, ARG6);
	case ACT_CONTROL:
		return Control(SEVEN_ARGS);
	case ACT_CONTROLGET:
		return ControlGet(ARG2, ARG3, ARG4, ARG5, ARG6, ARG7, ARG8);
	case ACT_STATUSBARGETTEXT:
		return StatusBarGetText(ARG2, ARG3, ARG4, ARG5, ARG6);
	case ACT_STATUSBARWAIT:
		return StatusBarWait(EIGHT_ARGS);
	case ACT_POSTMESSAGE:
	case ACT_SENDMESSAGE:
		return ScriptPostSendMessage(mActionType == ACT_SENDMESSAGE);
	case ACT_PROCESS:
		return ScriptProcess(THREE_ARGS);
	case ACT_WINSET:
		return WinSet(SIX_ARGS);
	case ACT_WINSETTITLE:
		return mArgc > 1 ? WinSetTitle(FIVE_ARGS) : WinSetTitle(_T(""), _T(""), ARG1);
	case ACT_WINGETTITLE:
		return WinGetTitle(ARG2, ARG3, ARG4, ARG5);
	case ACT_WINGETCLASS:
		return WinGetClass(ARG2, ARG3, ARG4, ARG5);
	case ACT_WINGET:
		return WinGet(ARG2, ARG3, ARG4, ARG5, ARG6);
	case ACT_WINGETTEXT:
		return WinGetText(ARG2, ARG3, ARG4, ARG5);
	case ACT_WINGETPOS:
		return WinGetPos(ARG5, ARG6, ARG7, ARG8);

	case ACT_SYSGET:
		return SysGet(ARG2, ARG3);

	case ACT_WINMINIMIZEALL:
		PostMessage(FindWindow(_T("Shell_TrayWnd"), NULL), WM_COMMAND, 419, 0);
		DoWinDelay;
		return OK;
	case ACT_WINMINIMIZEALLUNDO:
		PostMessage(FindWindow(_T("Shell_TrayWnd"), NULL), WM_COMMAND, 416, 0);
		DoWinDelay;
		return OK;

	case ACT_ONEXIT:
	{
		Label *target_label;
		// If it wasn't resolved at load-time, it must be a variable reference:
		if (   !(target_label = (Label *)mAttribute)   )
			if (   *ARG1 && !(target_label = g_script.FindLabel(ARG1))   )
					return LineError(ERR_NO_LABEL, FAIL, ARG1);
		g_script.mOnExitLabel = target_label;
		return OK;
	}

	case ACT_HOTKEY:
		// mAttribute is the label resolved at loadtime, if available (for performance).
		return Hotkey::Dynamic(THREE_ARGS, (IObject *)mAttribute, ARGVAR2);

	case ACT_SETTIMER: // A timer is being created, changed, or enabled/disabled.
	{
		IObject *target_label;
		// Note that only one timer per label is allowed because the label is the unique identifier
		// that allows us to figure out whether to "update or create" when searching the list of timers.
		if (   !(target_label = (IObject *)mAttribute)   ) // Since it wasn't resolved at load-time, it must be a variable reference.
		{
			if (   !(target_label = g_script.FindCallable(ARG1, ARGVAR1))   )
			{
				if (*ARG1)
					// ARG1 is a non-empty string and not the name of an existing label or function.
					return LineError(ERR_NO_LABEL, FAIL, ARG1);
				// Possible cases not ruled out by the above check:
				//   1) Label was omitted.
				//   2) Label was a single variable or non-expression which produced an empty value.
				//   3) Label was a single variable containing an incompatible function.
				//   4) Label was an expression which produced an empty value.
				//   5) Label was an expression which produced an object.
				// Case 3 is always an error.
				// Case 2 is arguably more likely to be an error (not intended to be empty) than meant as
				// an indicator to use the current label, so it seems safest to treat it as an error (and
				// also more consistent with Case 3).
				// Case 5 is currently not supported; the object reference was converted to an empty string
				// at an earlier stage, so it is indistinguishable from Case 4.  It seems rare that someone
				// would have a legitimate need for Case 4, so both cases are treated as an error.  This
				// covers cases like:
				//   SetTimer, % Func(a).Bind(b), xxx  ; Unsupported.
				//   SetTimer, % this.myTimerFunc, xxx  ; Unsupported (where myTimerFunc is an object).
				//   SetTimer, % this.MyMethod, xxx  ; Additional error: failing to bind "this" to MyMethod.
				// The following could be used to show "must not be blank" for Case 2, but it seems best
				// to reserve that message for when the parameter is really blank, not an empty variable:
				//if (mArgc > 0 && (mArg[0].is_expression /* Cases 4 & 5 */ || ARGVAR1 && ARGVAR1->HasObject() /* Case 3 */))
				if (*RAW_ARG1)
					return LineError(ERR_PARAM1_INVALID);
				if (g.CurrentLabel)
					// For backward-compatibility, use A_ThisLabel if set.  This can differ from CurrentTimer
					// when goto/gosub is used.  Some scripts apparently use this with subroutines that are
					// called both directly and by a timer.  The down side is that if a timer function uses
					// goto/gosub, A_ThisLabel must take predence; that may or may not be the user's intention.
					target_label = g.CurrentLabel;
				else if (g.CurrentTimer)
					// Default to the timer which launched the current thread.
					target_label = g.CurrentTimer->mLabel.ToObject();
				if (!target_label)
					// Either the thread was not launched by a timer or the timer has been deleted.
					return LineError(ERR_PARAM1_MUST_NOT_BE_BLANK);
			}
		}
		// And don't update mAttribute (leave it NULL) because we want ARG1 to be dynamically resolved
		// every time the command is executed (in case the contents of the referenced variable change).
		// In the data structure that holds the timers, we store the target label rather than the target
		// line so that a label can be registered independently as a timer even if there is another label
		// that points to the same line such as in this example:
		// Label1:
		// Label2:
		// ...
		// return
		if (!IsPureNumeric(ARG2, true, true, true)) // Allow it to be neg. or floating point at runtime.
		{
			toggle = Line::ConvertOnOff(ARG2);
			if (!toggle)
			{
				if (!_tcsicmp(ARG2, _T("Delete")))
				{
					g_script.DeleteTimer(target_label);
					return OK;
				}
				return LineError(ERR_PARAM2_INVALID, FAIL, ARG2);
			}
		}
		else
			toggle = TOGGLE_INVALID;
		// Below relies on distinguishing a true empty string from one that is sent to the function
		// as empty as a signal.  Don't change it without a full understanding because it's likely
		// to break compatibility or something else:
		switch(toggle)
		{
		case TOGGLED_ON:  
		case TOGGLED_OFF: g_script.UpdateOrCreateTimer(target_label, _T(""), ARG3, toggle == TOGGLED_ON, false); break;
		// Timer is always (re)enabled when ARG2 specifies a numeric period or is blank + there's no ARG3.
		// If ARG2 is blank but ARG3 (priority) isn't, tell it to update only the priority and nothing else:
		default: g_script.UpdateOrCreateTimer(target_label, ARG2, ARG3, true, !*ARG2 && *ARG3);
		}
		return OK;
	}

	case ACT_CRITICAL:
	{
		// v1.0.46: When the current thread is critical, have the script check messages less often to
		// reduce situations where an OnMesage or GUI message must be discarded due to "thread already
		// running".  Using 16 rather than the default of 5 solves reliability problems in a custom-menu-draw
		// script and probably many similar scripts -- even when the system is under load (though 16 might not
		// be enough during an extreme load depending on the exact preemption/timeslice dynamics involved).
		// DON'T GO TOO HIGH because this setting reduces response time for ALL messages, even those that
		// don't launch script threads (especially painting/drawing and other screen-update events).
		// Future enhancement: Could allow the value of 16 to be customized via something like "Critical 25".
		// However, it seems best not to allow it to go too high (say, no more than 2000) because that would
		// cause the script to completely hang if the critical thread never finishes, or takes a long time
		// to finish.  A configurable limit might also allow things to work better on Win9x because it has
		// a bigger tickcount granularity.
		// Some hardware has a tickcount granularity of 15 instead of 10, so this covers more variations.
		DWORD peek_frequency_when_critical_is_on = 16; // Set default.  See below.
		// v1.0.48: Below supports "Critical 0" as meaning "Off" to improve compatibility with A_IsCritical.
		// In fact, for performance, only the following are no recognized as turning on Critical:
		//     - "On"
		//     - ""
		//     - Integer other than 0.
		// Everything else, if considered to be "Off", including "Off", "Any non-blank string that
		// doesn't start with a non-zero number", and zero itself.
		g.ThreadIsCritical = !*ARG1 // i.e. a first arg that's omitted or blank is the same as "ON". See comments above.
			|| !_tcsicmp(ARG1, _T("ON"))
			|| (peek_frequency_when_critical_is_on = ArgToUInt(1)); // Non-zero integer also turns it on. Relies on short-circuit boolean order.
		if (g.ThreadIsCritical) // Critical has been turned on. (For simplicity even if it was already on, the following is done.)
		{
			g.PeekFrequency = peek_frequency_when_critical_is_on;
			g.AllowThreadToBeInterrupted = false;
			g.LinesPerCycle = -1;      // v1.0.47: It seems best to ensure SetBatchLines -1 is in effect because
			g.IntervalBeforeRest = -1; // otherwise it may check messages during the interval that it isn't supposed to.
		}
		else // Critical has been turned off.
		{
			// Since Critical is being turned off, allow thread to be immediately interrupted regardless of
			// any "Thread Interrupt" settings.
			g.PeekFrequency = DEFAULT_PEEK_FREQUENCY;
			g.AllowThreadToBeInterrupted = true;
		}
		// Once ACT_CRITICAL returns, the thread's interruptibility has been explicitly set; so the script
		// is now in charge of managing this thread's interruptibility.
		return OK;
	}

	case ACT_THREAD:
		switch (ConvertThreadCommand(ARG1))
		{
		case THREAD_CMD_PRIORITY:
			g.Priority = ArgToInt(2);
			break;
		case THREAD_CMD_INTERRUPT:
			// If either one is blank, leave that setting as it was before.
			if (*ARG2)
				g_script.mUninterruptibleTime = ArgToInt(2);  // 32-bit (for compatibility with DWORDs returned by GetTickCount).
			if (*ARG3)
				g_script.mUninterruptedLineCountMax = ArgToInt(3);  // 32-bit also, to help performance (since huge values seem unnecessary).
			break;
		case THREAD_CMD_NOTIMERS:
			g.AllowTimers = (*ARG2 && ArgToInt64(2) == 0);
			break;
		// If invalid command, do nothing since that is always caught at load-time unless the command
		// is in a variable reference (very rare in this case).
		}
		return OK;

	case ACT_GROUPADD: // Adding a WindowSpec *to* a group, not adding a group.
	{
		if (   !(group = (WinGroup *)mAttribute)   )
			if (   !(group = g_script.FindGroup(ARG1, true))   )  // Last parameter -> create-if-not-found.
				return FAIL;  // It already displayed the error for us.
		Label *target_label = NULL;
		if (*ARG4)
		{
			if (   !(target_label = (Label *)mRelatedLine)   ) // Jump target hasn't been resolved yet, probably due to it being a deref.
				if (   !(target_label = g_script.FindLabel(ARG4))   )
					return LineError(ERR_NO_LABEL, FAIL, ARG4);
			// Can't do this because the current line won't be the launching point for the
			// Gosub.  Instead, the launching point will be the GroupActivate rather than the
			// GroupAdd, so it will be checked by the GroupActivate or not at all (since it's
			// not that important in the case of a Gosub -- it's mostly for Goto's):
			//return IsJumpValid(label->mJumpToLine);
			group->mJumpToLabel = target_label;
		}
		return group->AddWindow(ARG2, ARG3, ARG5, ARG6);
	}

	// Note ACT_GROUPACTIVATE is handled by ExecUntil(), since it's better suited to do the Gosub.
	case ACT_GROUPDEACTIVATE:
		if (   !(group = (WinGroup *)mAttribute)   )
			group = g_script.FindGroup(ARG1);
		if (group)
			group->Deactivate(*ARG2 && !_tcsicmp(ARG2, _T("R")));  // Note: It will take care of DoWinDelay if needed.
		//else nonexistent group: By design, do nothing.
		return OK;

	case ACT_GROUPCLOSE:
		if (   !(group = (WinGroup *)mAttribute)   )
			group = g_script.FindGroup(ARG1);
		if (group)
			if (*ARG2 && !_tcsicmp(ARG2, _T("A")))
				group->ActUponAll(ACT_WINCLOSE, 0);  // Note: It will take care of DoWinDelay if needed.
			else
				group->CloseAndGoToNext(*ARG2 && !_tcsicmp(ARG2, _T("R")));  // Note: It will take care of DoWinDelay if needed.
		//else nonexistent group: By design, do nothing.
		return OK;

	case ACT_GETKEYSTATE:
		return GetKeyJoyState(ARG2, ARG3);

	case ACT_RANDOM:
	{
		if (!output_var) // v1.0.42.03: Special mode to change the seed.
		{
			init_genrand(ArgToUInt(2)); // It's documented that an unsigned 32-bit number is required.
			return OK;
		}
		bool use_float = IsPureNumeric(ARG2, true, false, true) == PURE_FLOAT
			|| IsPureNumeric(ARG3, true, false, true) == PURE_FLOAT;
		if (use_float)
		{
			double rand_min = *ARG2 ? ArgToDouble(2) : 0;
			double rand_max = *ARG3 ? ArgToDouble(3) : INT_MAX;
			// Seems best not to use ErrorLevel for this command at all, since silly cases
			// such as Max > Min are too rare.  Swap the two values instead.
			if (rand_min > rand_max)
			{
				double rand_swap = rand_min;
				rand_min = rand_max;
				rand_max = rand_swap;
			}
			return output_var->Assign((genrand_real1() * (rand_max - rand_min)) + rand_min);
		}
		else // Avoid using floating point, where possible, which may improve speed a lot more than expected.
		{
			int rand_min = *ARG2 ? ArgToInt(2) : 0;
			int rand_max = *ARG3 ? ArgToInt(3) : INT_MAX;
			// Seems best not to use ErrorLevel for this command at all, since silly cases
			// such as Max > Min are too rare.  Swap the two values instead.
			if (rand_min > rand_max)
			{
				int rand_swap = rand_min;
				rand_min = rand_max;
				rand_max = rand_swap;
			}
			// Do NOT use genrand_real1() to generate random integers because of cases like
			// min=0 and max=1: we want an even distribution of 1's and 0's in that case, not
			// something skewed that might result due to rounding/truncation issues caused by
			// the float method used above:
			// AutoIt3: __int64 is needed here to do the proper conversion from unsigned long to signed long:
			return output_var->Assign(   (int)(__int64(genrand_int32()
				% ((__int64)rand_max - rand_min + 1)) + rand_min)   );
		}
	}

	case ACT_DRIVESPACEFREE:
		return DriveSpace(ARG2, true);

	case ACT_DRIVE:
		return Drive(THREE_ARGS);

	case ACT_DRIVEGET:
		return DriveGet(ARG2, ARG3);

	case ACT_SOUNDGET:
	case ACT_SOUNDSET:
		return SoundSetGet(mActionType == ACT_SOUNDGET ? NULL : ARG1, ARG2, ARG3, ARG4);

	case ACT_SOUNDGETWAVEVOLUME:
	case ACT_SOUNDSETWAVEVOLUME:
		device_id = *ARG2 ? ArgToInt(2) - 1 : 0;
		if (device_id < 0)
			device_id = 0;
		return (mActionType == ACT_SOUNDGETWAVEVOLUME) ? SoundGetWaveVolume((HWAVEOUT)device_id)
			: SoundSetWaveVolume(ARG1, (HWAVEOUT)device_id);

	case ACT_SOUNDBEEP:
		// For simplicity and support for future/greater capabilities, no range checking is done.
		// It simply calls the function with the two DWORD values provided. It avoids setting
		// ErrorLevel because failure is rare and also because a script might want play a beep
		// right before displaying an error dialog that uses the previous value of ErrorLevel.
		Beep(*ARG1 ? ArgToUInt(1) : 523, *ARG2 ? ArgToUInt(2) : 150);
		return OK;

	case ACT_SOUNDPLAY:
		return SoundPlay(ARG1, *ARG2 && !_tcsicmp(ARG2, _T("wait")) || !_tcsicmp(ARG2, _T("1")));

	case ACT_FILEAPPEND:
		// Uses the read-file loop's current item filename was explicitly leave blank (i.e. not just
		// a reference to a variable that's blank):
		return FileAppend(ARG2, ARG1, (mArgc < 2) ? g.mLoopReadFile : NULL);

	case ACT_FILEREAD:
		return FileRead(ARG2);

	case ACT_FILEREADLINE:
		return FileReadLine(ARG2, ARG3);

	case ACT_FILEDELETE:
		return FileDelete(ARG1);

	case ACT_FILERECYCLE:
		return FileRecycle(ARG1);

	case ACT_FILERECYCLEEMPTY:
		return FileRecycleEmpty(ARG1);

	case ACT_FILEINSTALL:
		return FileInstall(THREE_ARGS);

	case ACT_FILECOPY:
	{
		int error_count = Util_CopyFile(ARG1, ARG2, ArgToInt(3) == 1, false, g.LastError);
		if (!error_count)
			return g_ErrorLevel->Assign(ERRORLEVEL_NONE);
		return SetErrorLevelOrThrowInt(error_count);
	}
	case ACT_FILEMOVE:
		return SetErrorLevelOrThrowInt(Util_CopyFile(ARG1, ARG2, ArgToInt(3) == 1, true, g.LastError));
	case ACT_FILECOPYDIR:
		return SetErrorLevelOrThrowBool(!Util_CopyDir(ARG1, ARG2, ArgToInt(3) == 1));
	case ACT_FILEMOVEDIR:
		if (ctoupper(*ARG3) == 'R')
		{
			// Perform a simple rename instead, which prevents the operation from being only partially
			// complete if the source directory is in use (due to being a working dir for a currently
			// running process, or containing a file that is being written to).  In other words,
			// the operation will be "all or none":
			return SetErrorLevelOrThrowBool(!MoveFile(ARG1, ARG2));
		}
		// Otherwise:
		return SetErrorLevelOrThrowBool(!Util_MoveDir(ARG1, ARG2, ArgToInt(3)));

	case ACT_FILECREATEDIR:
		return FileCreateDir(ARG1);
	case ACT_FILEREMOVEDIR:
		return SetErrorLevelOrThrowBool(!*ARG1 // Consider an attempt to create or remove a blank dir to be an error.
			|| !Util_RemoveDir(ARG1, ArgToInt(2) == 1)); // Relies on short-circuit evaluation.

	case ACT_FILEGETATTRIB:
		// The specified ARG, if non-blank, takes precedence over the file-loop's file (if any):
		#define USE_FILE_LOOP_FILE_IF_ARG_BLANK(arg) (*arg ? arg : (g.mLoopFile ? g.mLoopFile->cFileName : _T("")))
		return FileGetAttrib(USE_FILE_LOOP_FILE_IF_ARG_BLANK(ARG2));
	case ACT_FILESETATTRIB:
		FileSetAttrib(ARG1, USE_FILE_LOOP_FILE_IF_ARG_BLANK(ARG2), ConvertLoopMode(ARG3), ArgToInt(4) == 1);
		return !g.ThrownToken ? OK : FAIL;
	case ACT_FILEGETTIME:
		return FileGetTime(USE_FILE_LOOP_FILE_IF_ARG_BLANK(ARG2), *ARG3);
	case ACT_FILESETTIME:
		FileSetTime(ARG1, USE_FILE_LOOP_FILE_IF_ARG_BLANK(ARG2), *ARG3, ConvertLoopMode(ARG4), ArgToInt(5) == 1);
		return !g.ThrownToken ? OK : FAIL;
	case ACT_FILEGETSIZE:
		return FileGetSize(USE_FILE_LOOP_FILE_IF_ARG_BLANK(ARG2), ARG3);
	case ACT_FILEGETVERSION:
		return FileGetVersion(USE_FILE_LOOP_FILE_IF_ARG_BLANK(ARG2));

	case ACT_SETWORKINGDIR:
		SetWorkingDir(ARG1);
		return !g.ThrownToken ? OK : FAIL;

	case ACT_FILESELECTFILE:
		return FileSelectFile(ARG2, ARG3, ARG4, ARG5);

	case ACT_FILESELECTFOLDER:
		return FileSelectFolder(ARG2, ARG3, ARG4);

	case ACT_FILEGETSHORTCUT:
		return FileGetShortcut(ARG1);
	case ACT_FILECREATESHORTCUT:
		return FileCreateShortcut(NINE_ARGS);

	case ACT_KEYHISTORY:
#ifdef ENABLE_KEY_HISTORY_FILE
		if (*ARG1 || *ARG2)
		{
			switch (ConvertOnOffToggle(ARG1))
			{
			case NEUTRAL:
			case TOGGLE:
				g_KeyHistoryToFile = !g_KeyHistoryToFile;
				if (!g_KeyHistoryToFile)
					KeyHistoryToFile();  // Signal it to close the file, if it's open.
				break;
			case TOGGLED_ON:
				g_KeyHistoryToFile = true;
				break;
			case TOGGLED_OFF:
				g_KeyHistoryToFile = false;
				KeyHistoryToFile();  // Signal it to close the file, if it's open.
				break;
			// We know it's a variable because otherwise the loading validation would have caught it earlier:
			case TOGGLE_INVALID:
				return LineError(ERR_PARAM1_INVALID, FAIL, ARG1);
			}
			if (*ARG2) // The user also specified a filename, so update the target filename.
				KeyHistoryToFile(ARG2);
			return OK;
		}
#endif
		// Otherwise:
		return ShowMainWindow(MAIN_MODE_KEYHISTORY, false); // Pass "unrestricted" when the command is explicitly used in the script.
	case ACT_LISTLINES:
		if (   (toggle = ConvertOnOff(ARG1, NEUTRAL)) == NEUTRAL   )
			return ShowMainWindow(MAIN_MODE_LINES, false); // Pass "unrestricted" when the command is explicitly used in the script.
		// Otherwise:
		if (g.ListLinesIsEnabled)
		{
			// Since ExecUntil() just logged this ListLines On/Off in the line history, remove it to avoid
			// cluttering the line history with distracting lines that the user probably wouldn't want to see.
			// Might be especially useful in cases where a timer fires frequently (even if such a timer
			// used "ListLines Off" as its top line, that line itself would appear very frequently in the line
			// history).  v1.0.48.03: Fixed so that the below executes only when ListLines was previously "On".
			if (sLogNext > 0)
				--sLogNext;
			else
				sLogNext = LINE_LOG_SIZE - 1;
			sLog[sLogNext] = NULL; // Without this, one of the lines in the history would be invalid due to the circular nature of the line history array, which would also cause the line history to show the wrong chronological order in some cases.
		}
		g.ListLinesIsEnabled = (toggle == TOGGLED_ON);
		return OK;
	case ACT_LISTVARS:
		return ShowMainWindow(MAIN_MODE_VARS, false); // Pass "unrestricted" when the command is explicitly used in the script.
	case ACT_LISTHOTKEYS:
		return ShowMainWindow(MAIN_MODE_HOTKEYS, false); // Pass "unrestricted" when the command is explicitly used in the script.

	case ACT_MSGBOX:
	{
		int result;
		HWND dialog_owner = THREAD_DIALOG_OWNER; // Resolve macro only once to reduce code size.
		// If the MsgBox window can't be displayed for any reason, always return FAIL to
		// the caller because it would be unsafe to proceed with the execution of the
		// current script subroutine.  For example, if the script contains an IfMsgBox after,
		// this line, it's result would be unpredictable and might cause the subroutine to perform
		// the opposite action from what was intended (e.g. Delete vs. don't delete a file).
		if (!mArgc) // When called explicitly with zero params, it displays this default msg.
			result = MsgBox(_T("Press OK to continue."), MSGBOX_NORMAL, NULL, 0, dialog_owner);
		else if (mArgc == 1) // In the special 1-parameter mode, the first param is the prompt.
			result = MsgBox(ARG1, MSGBOX_NORMAL, NULL, 0, dialog_owner);
		else
			result = MsgBox(ARG3, ArgToInt(1), ARG2, ArgToDouble(4), dialog_owner); // dialog_owner passed via parameter to avoid internally-displayed MsgBoxes from being affected by script-thread's owner setting.
		// Above allows backward compatibility with AutoIt2's param ordering while still
		// permitting the new method of allowing only a single param.
		// v1.0.40.01: Rather than displaying another MsgBox in response to a failed attempt to display
		// a MsgBox, it seems better (less likely to cause trouble) just to abort the thread.  This also
		// solves a double-msgbox issue when the maximum number of MsgBoxes is reached.  In addition, the
		// max-msgbox limit is the most common reason for failure, in which case a warning dialog has
		// already been displayed, so there is no need to display another:
		//if (!result)
		//	// It will fail if the text is too large (say, over 150K or so on XP), but that
		//	// has since been fixed by limiting how much it tries to display.
		//	// If there were too many message boxes displayed, it will already have notified
		//	// the user of this via a final MessageBox dialog, so our call here will
		//	// not have any effect.  The below only takes effect if MsgBox()'s call to
		//	// MessageBox() failed in some unexpected way:
		//	LineError("The MsgBox could not be displayed.");
		// v1.1.09.02: If the MsgBox failed due to invalid options, it seems better to display
		// an error dialog than to silently exit the thread:
		if (!result && GetLastError() == ERROR_INVALID_MSGBOX_STYLE)
			return LineError(ERR_PARAM1_INVALID, FAIL, ARG1);
		return result ? OK : FAIL;
	}

	case ACT_INPUTBOX:
		return InputBox(output_var, ARG2, ARG3, ctoupper(*ARG4) == 'H' // 4th is whether to hide input.
			, *ARG5 ? ArgToInt(5) : INPUTBOX_DEFAULT  // Width
			, *ARG6 ? ArgToInt(6) : INPUTBOX_DEFAULT  // Height
			, *ARG7 ? ArgToInt(7) : INPUTBOX_DEFAULT  // Xpos
			, *ARG8 ? ArgToInt(8) : INPUTBOX_DEFAULT  // Ypos
			// ARG9: future use for Font name & size, e.g. "Courier:8"
			, ArgToDouble(10)  // Timeout
			, ARG11  // Initial default string for the edit field.
			);

	case ACT_SPLASHTEXTON:
		return SplashTextOn(*ARG1 ? ArgToInt(1) : 200, *ARG2 ? ArgToInt(2) : 0, ARG3, ARG4);
	case ACT_SPLASHTEXTOFF:
		DESTROY_SPLASH
		return OK;

	case ACT_PROGRESS:
		return Splash(FIVE_ARGS, _T(""), false);  // ARG6 is for future use and currently not passed.

	case ACT_SPLASHIMAGE:
		return Splash(ARG2, ARG3, ARG4, ARG5, ARG6, ARG1, true);  // ARG7 is for future use and currently not passed.

	case ACT_TOOLTIP:
		return ToolTip(FOUR_ARGS);

	case ACT_TRAYTIP:
		return TrayTip(FOUR_ARGS);

	case ACT_INPUT:
		return Input();


//////////////////////////////////////////////////////////////////////////

	case ACT_COORDMODE:
	{
		CoordModeType mode = ConvertCoordMode(ARG2);
		CoordModeType shift = ConvertCoordModeCmd(ARG1);
		if (shift != -1 && mode != -1) // Compare directly to -1 because unsigned.
			g.CoordMode = (g.CoordMode & ~(COORD_MODE_MASK << shift)) | (mode << shift);
		//else too rare to report an error, since load-time validation normally catches it.
		return OK;
	}

	case ACT_SETDEFAULTMOUSESPEED:
		g.DefaultMouseSpeed = (UCHAR)ArgToInt(1);
		// In case it was a deref, force it to be some default value if it's out of range:
		if (g.DefaultMouseSpeed < 0 || g.DefaultMouseSpeed > MAX_MOUSE_SPEED)
			g.DefaultMouseSpeed = DEFAULT_MOUSE_SPEED;
		return OK;

	case ACT_SENDMODE:
		g.SendMode = ConvertSendMode(ARG1, g.SendMode); // Leave value unchanged if ARG1 is invalid.
		return OK;

	case ACT_SENDLEVEL:
	{
		int sendLevel = ArgToInt(1);
		if (SendLevelIsValid(sendLevel))
			g.SendLevel = sendLevel;

		return OK;
	}

	case ACT_SETKEYDELAY:
		if (!_tcsicmp(ARG3, _T("Play")))
		{
			if (*ARG1)
				g.KeyDelayPlay = ArgToInt(1);
			if (*ARG2)
				g.PressDurationPlay = ArgToInt(2);
		}
		else
		{
			if (*ARG1)
				g.KeyDelay = ArgToInt(1);
			if (*ARG2)
				g.PressDuration = ArgToInt(2);
		}
		return OK;
	case ACT_SETMOUSEDELAY:
		if (!_tcsicmp(ARG2, _T("Play")))
			g.MouseDelayPlay = ArgToInt(1);
		else
			g.MouseDelay = ArgToInt(1);
		return OK;
	case ACT_SETWINDELAY:
		g.WinDelay = ArgToInt(1);
		return OK;
	case ACT_SETCONTROLDELAY:
		g.ControlDelay = ArgToInt(1);
		return OK;

	case ACT_SETBATCHLINES:
		// This below ensures that IntervalBeforeRest and LinesPerCycle aren't both in effect simultaneously
		// (i.e. that both aren't greater than -1), even though ExecUntil() has code to prevent a double-sleep
		// even if that were to happen.
		if (tcscasestr(ARG1, _T("ms"))) // This detection isn't perfect, but it doesn't seem necessary to be too demanding.
		{
			g.LinesPerCycle = -1;  // Disable the old BatchLines method in favor of the new one below.
			g.IntervalBeforeRest = ArgToInt(1);  // If negative, script never rests.  If 0, it rests after every line.
		}
		else
		{
			g.IntervalBeforeRest = -1;  // Disable the new method in favor of the old one below:
			// This value is signed 64-bits to support variable reference (i.e. containing a large int)
			// the user might throw at it:
			if (   !(g.LinesPerCycle = ArgToInt64(1))   )
				// Don't interpret zero as "infinite" because zero can accidentally
				// occur if the dereferenced var was blank:
				g.LinesPerCycle = 10;  // The old default, which is retained for compatibility with existing scripts.
		}
		return OK;

	case ACT_SETSTORECAPSLOCKMODE:
		if (   (toggle = ConvertOnOff(ARG1, NEUTRAL)) != NEUTRAL   )
			g.StoreCapslockMode = (toggle == TOGGLED_ON);
		return OK;

	case ACT_SETTITLEMATCHMODE:
		switch (ConvertTitleMatchMode(ARG1))
		{
		case FIND_IN_LEADING_PART: g.TitleMatchMode = FIND_IN_LEADING_PART; return OK;
		case FIND_ANYWHERE: g.TitleMatchMode = FIND_ANYWHERE; return OK;
		case FIND_REGEX: g.TitleMatchMode = FIND_REGEX; return OK;
		case FIND_EXACT: g.TitleMatchMode = FIND_EXACT; return OK;
		case FIND_FAST: g.TitleFindFast = true; return OK;
		case FIND_SLOW: g.TitleFindFast = false; return OK;
		}
		return LineError(ERR_PARAM1_INVALID, FAIL, ARG1);

	case ACT_SETFORMAT:
		// For now, it doesn't seem necessary to have runtime validation of the first parameter.
		// Just ignore the command if it's not valid:
		if (!_tcsnicmp(ARG1, _T("Float"), 5)) // "nicmp" vs. "icmp" so that Float and FloatFast are treated the same (loadtime validation already took notice of the Fast flag).
		{
			// -2 to allow room for the letter 'f' and the '%' that will be added:
			if (ArgLength(2) >= _countof(g.FormatFloat) - 2) // A variable that resolved to something too long.
				return OK; // Seems best not to bother with a runtime error for something so rare.
			// Make sure the formatted string wouldn't exceed the buffer size:
			__int64 width = ArgToInt64(2);
			LPTSTR dot_pos = _tcschr(ARG2, '.');
			__int64 precision = dot_pos ? ATOI64(dot_pos + 1) : 0;
			if (width + precision + 2 > MAX_NUMBER_LENGTH) // +2 to allow room for decimal point itself and leading minus sign.
				return OK; // Don't change it.
			// Create as "%ARG2f".  Note that %f can handle doubles in MSVC++:
			_stprintf(g.FormatFloat, _T("%%%s%s%s"), ARG2
				, dot_pos ? _T("") : _T(".") // Add a dot if none was specified so that "0" is the same as "0.", which seems like the most user-friendly approach; it's also easier to document in the help file.
				, IsPureNumeric(ARG2, true, true, true) ? _T("f") : _T("")); // If it's not pure numeric, assume the user already included the desired letter (e.g. SetFormat, Float, 0.6e).
		}
		else if (!_tcsnicmp(ARG1, _T("Integer"), 7)) // "nicmp" vs. "icmp" so that Integer and IntegerFast are treated the same (loadtime validation already took notice of the Fast flag).
		{
			switch(*ARG2)
			{
			case 'd':
			case 'D':
				g.FormatInt = 'D';
				break;
			case 'h':
			case 'H':
				g.FormatInt = (char) *ARG2;
				break;
			// Otherwise, since the first letter isn't recognized, do nothing since 99% of the time such a
			// probably would be caught at load-time.
			}
		}
		// Otherwise, ignore invalid type at runtime since 99% of the time it would be caught at load-time:
		return OK;

	case ACT_FORMATTIME:
		return FormatTime(ARG2, ARG3);

	case ACT_MENU:
		return g_script.PerformMenu(SIX_ARGS, ARGVAR4, ARGVAR5);

	case ACT_GUI:
		return g_script.PerformGui(FOUR_ARGS);

	case ACT_GUICONTROL:
		return GuiControl(THREE_ARGS, ARGVAR3);

	case ACT_GUICONTROLGET:
		return GuiControlGet(ARG2, ARG3, ARG4);

	////////////////////////////////////////////////////////////////////////////////////////
	// For these, it seems best not to report an error during runtime if there's
	// an invalid value (e.g. something other than On/Off/Blank) in a param containing
	// a dereferenced variable, since settings are global and affect all subroutines,
	// not just the one that we would otherwise report failure for:
	case ACT_SUSPEND:
		switch (ConvertOnOffTogglePermit(ARG1))
		{
		case NEUTRAL:
		case TOGGLE:
			ToggleSuspendState();
			break;
		case TOGGLED_ON:
			if (!g_IsSuspended)
				ToggleSuspendState();
			break;
		case TOGGLED_OFF:
			if (g_IsSuspended)
				ToggleSuspendState();
			break;
		case TOGGLE_PERMIT:
			// In this case do nothing.  The user is just using this command as a flag to indicate that
			// this subroutine should not be suspended.
			break;
		// We know it's a variable because otherwise the loading validation would have caught it earlier:
		case TOGGLE_INVALID:
			return LineError(ERR_PARAM1_INVALID, FAIL, ARG1);
		}
		return OK;
	case ACT_PAUSE:
		return ChangePauseState(ConvertOnOffToggle(ARG1), (bool)ArgToInt(2));
	case ACT_AUTOTRIM:
		if (   (toggle = ConvertOnOff(ARG1, NEUTRAL)) != NEUTRAL   )
			g.AutoTrim = (toggle == TOGGLED_ON);
		return OK;
	case ACT_STRINGCASESENSE:
		if ((g.StringCaseSense = ConvertStringCaseSense(ARG1)) == SCS_INVALID)
			g.StringCaseSense = SCS_INSENSITIVE; // For simplicity, just fall back to default if value is invalid (normally its caught at load-time; only rarely here).
		return OK;
	case ACT_DETECTHIDDENWINDOWS:
		if (   (toggle = ConvertOnOff(ARG1, NEUTRAL)) != NEUTRAL   )
			g.DetectHiddenWindows = (toggle == TOGGLED_ON);
		return OK;
	case ACT_DETECTHIDDENTEXT:
		if (   (toggle = ConvertOnOff(ARG1, NEUTRAL)) != NEUTRAL   )
			g.DetectHiddenText = (toggle == TOGGLED_ON);
		return OK;
	case ACT_BLOCKINPUT:
		switch (toggle = ConvertBlockInput(ARG1))
		{
		case TOGGLED_ON:
			ScriptBlockInput(true);
			break;
		case TOGGLED_OFF:
			ScriptBlockInput(false);
			break;
		case TOGGLE_SEND:
		case TOGGLE_MOUSE:
		case TOGGLE_SENDANDMOUSE:
		case TOGGLE_DEFAULT:
			g_BlockInputMode = toggle;
			break;
		case TOGGLE_MOUSEMOVE:
			g_BlockMouseMove = true;
			Hotkey::InstallMouseHook();
			break;
		case TOGGLE_MOUSEMOVEOFF:
			g_BlockMouseMove = false; // But the mouse hook is left installed because it might be needed by other things. This approach is similar to that used by the Input command.
			break;
		// default (NEUTRAL or TOGGLE_INVALID): do nothing.
		}
		return OK;

	////////////////////////////////////////////////////////////////////////////////////////
	// For these, it seems best not to report an error during runtime if there's
	// an invalid value (e.g. something other than On/Off/Blank) in a param containing
	// a dereferenced variable, since settings are global and affect all subroutines,
	// not just the one that we would otherwise report failure for:
	case ACT_SETNUMLOCKSTATE:
		return SetToggleState(VK_NUMLOCK, g_ForceNumLock, ARG1);
	case ACT_SETCAPSLOCKSTATE:
		return SetToggleState(VK_CAPITAL, g_ForceCapsLock, ARG1);
	case ACT_SETSCROLLLOCKSTATE:
		return SetToggleState(VK_SCROLL, g_ForceScrollLock, ARG1);

	case ACT_EDIT:
		g_script.Edit();
		return OK;
	case ACT_RELOAD:
		g_script.Reload(true);
		// Even if the reload failed, it seems best to return OK anyway.  That way,
		// the script can take some follow-on action, e.g. it can sleep for 1000
		// after issuing the reload command and then take action under the assumption
		// that the reload didn't work (since obviously if the process and thread
		// in which the Sleep is running still exist, it didn't work):
		return OK;

	case ACT_SLEEP:
	{
		// Only support 32-bit values for this command, since it seems unlikely anyone would to have
		// it sleep more than 24.8 days or so.  It also helps performance on 32-bit hardware because
		// MsgSleep() is so heavily called and checks the value of the first parameter frequently:
		int sleep_time = ArgToInt(1); // Keep it signed vs. unsigned for backward compatibility (e.g. scripts that do Sleep -1).

		// Do a true sleep on Win9x because the MsgSleep() method is very inaccurate on Win9x
		// for some reason (a MsgSleep(1) causes a sleep between 10 and 55ms, for example).
		// But only do so for short sleeps, for which the user has a greater expectation of
		// accuracy.  UPDATE: Do not change the 25 below without also changing it in Critical's
		// documentation.
		if (sleep_time < 25 && sleep_time > 0 && g_os.IsWin9x()) // Ordered for short-circuit performance. v1.0.38.05: Added "sleep_time > 0" so that Sleep -1/0 will work the same on Win9x as it does on other OSes.
			Sleep(sleep_time);
		else
			MsgSleep(sleep_time);
		return OK;
	}

	case ACT_INIREAD:
		return IniRead(ARG2, ARG3, ARG4, ARG5);
	case ACT_INIWRITE:
		return IniWrite(FOUR_ARGS);
	case ACT_INIDELETE:
		// To preserve maximum compatibility with existing scripts, only send NULL if ARG3
		// was explicitly omitted.  This is because some older scripts might rely on the
		// fact that a blank ARG3 does not delete the entire section, but rather does
		// nothing (that fact is untested):
		return IniDelete(ARG1, ARG2, mArgc < 3 ? NULL : ARG3);

	case ACT_REGREAD:
		if (mArgc < 2 && g.mLoopRegItem) // Uses the registry loop's current item.
			// If g.mLoopRegItem->name specifies a subkey rather than a value name, do this anyway
			// so that it will set ErrorLevel to ERROR and set the output variable to be blank.
			// Also, do not use RegCloseKey() on this, even if it's a remote key, since our caller handles that:
			return RegRead(g.mLoopRegItem->root_key, g.mLoopRegItem->subkey, g.mLoopRegItem->name);
		// Otherwise:
		root_key = RegConvertKey(ARG2, REG_EITHER_SYNTAX, &subkey, &is_remote_registry);
		if (!subkey) // Old syntax (root key without slash).
			subkey = ARG3, value_name = ARG4;
		else // New syntax (root key combined with subkey).
			value_name = ARG3;
		result = RegRead(root_key, subkey, value_name);
		if (is_remote_registry && root_key) // Never try to close local root keys, which the OS keeps always-open.
			RegCloseKey(root_key);
		return result;
	case ACT_REGWRITE:
		if (mArgc < 2 && g.mLoopRegItem) // Uses the registry loop's current item.
			// If g.mLoopRegItem->name specifies a subkey rather than a value name, do this anyway
			// so that it will set ErrorLevel to ERROR.  An error will also be indicated if
			// g.mLoopRegItem->type is an unsupported type:
			return RegWrite(g.mLoopRegItem->type, g.mLoopRegItem->root_key, g.mLoopRegItem->subkey, g.mLoopRegItem->name, ARG1);
		// Otherwise:
		root_key = RegConvertKey(ARG2, REG_EITHER_SYNTAX, &subkey, &is_remote_registry);
		if (!subkey) // Old syntax (root key without slash).
			subkey = ARG3, value_name = ARG4, value = ARG5;
		else // New syntax (root key combined with subkey).
			value_name = ARG3, value = ARG4;
		result = RegWrite(RegConvertValueType(ARG1), root_key, subkey, value_name, value);
		// If the value type or root_key are invalid, RegWrite() has set ErrorLevel rather than displaying a runtime error.
		if (is_remote_registry && root_key) // Never try to close local root keys, which the OS keeps always-open.
			RegCloseKey(root_key);
		return result;
	case ACT_REGDELETE:
		if (mArgc < 1 && g.mLoopRegItem) // Uses the registry loop's current item.
		{
			// In this case, if the current reg item is a value, just delete it normally.
			// But if it's a subkey, append it to the dir name so that the proper subkey
			// will be deleted as the user intended:
			if (g.mLoopRegItem->type == REG_SUBKEY)
			{
				if (*g.mLoopRegItem->subkey)
				{
					sntprintf(buf_temp, _countof(buf_temp), _T("%s\\%s"), g.mLoopRegItem->subkey, g.mLoopRegItem->name);
					subkey = buf_temp;
				}
				else // It's a direct subkey of root_key.
					subkey = g.mLoopRegItem->name;
				value_name = NULL;
			}
			else
			{
				subkey = g.mLoopRegItem->subkey;
				value_name = g.mLoopRegItem->name;
			}
			return RegDelete(g.mLoopRegItem->root_key, subkey, value_name);
		}
		// Otherwise:
		root_key = RegConvertKey(ARG1, REG_EITHER_SYNTAX, &subkey, &is_remote_registry);
		if (!subkey) // Old syntax (root key without slash).
			subkey = ARG2, value_name = ARG3;
		else // New syntax (root key combined with subkey).
			value_name = ARG2;
		// For backward-compatibility, the special phrase "ahk_default" indicates that the key's
		// default value (displayed as "(Default)" by RegEdit) should be deleted, while a blank
		// or omitted value deletes the entire subkey.  "RegDelete" without parameters needs to
		// delete the current item even if that item is the default value (A_LoopRegName = "")
		// or a value named "ahk_default", so the keyword is handled here, not in RegDelete():
		if (!*value_name) // Blank or omitted: delete the entire subkey.
			value_name = NULL;
		else if (!_tcsicmp(value_name, _T("ahk_default"))) // Delete the key's default value.
			value_name = _T("");
		result = RegDelete(root_key, subkey, value_name);
		if (is_remote_registry && root_key) // Never try to close local root keys, which the OS always keeps open.
			RegCloseKey(root_key);
		return result;
	case ACT_SETREGVIEW:
	{
		DWORD reg_view = RegConvertView(ARG1);
		// Validate the parameter even if it's not going to be used.
		if (reg_view == -1)
			return LineError(ERR_PARAM1_INVALID, FAIL, ARG1);
		// Since these flags cause the registry functions to fail on Win2k and have no effect on
		// any later 32-bit OS, ignore this command when the OS is 32-bit.  Leave A_RegView blank.
		if (IsOS64Bit())
			g.RegView = reg_view;
		return OK;
	}

	case ACT_OUTPUTDEBUG:
#ifndef CONFIG_DEBUGGER
		OutputDebugString(ARG1); // It does not return a value for the purpose of setting ErrorLevel.
#else
		g_Debugger.OutputDebug(ARG1);
#endif
		return OK;

	case ACT_SHUTDOWN:
		return Util_Shutdown(ArgToInt(1)) ? OK : FAIL; // Range of ARG1 is not validated in case other values are supported in the future.

	case ACT_FILEENCODING:
		UINT new_encoding = ConvertFileEncoding(ARG1);
		if (new_encoding == -1)
			return LineError(ERR_PARAM1_INVALID, FAIL, ARG1); // Probably a variable, otherwise load-time validation would've caught it.
		g.Encoding = new_encoding;
		return OK;
	} // switch()

	// Since above didn't return, this line's mActionType isn't handled here,
	// so caller called it wrong.  ACT_INVALID should be impossible because
	// Script::AddLine() forbids it.

#ifdef _DEBUG
	return LineError(_T("DEBUG: Perform(): Unhandled action type."));
#else
	return FAIL;
#endif
}



ResultType Line::Deref(Var *aOutputVar, LPTSTR aBuf)
// Similar to ExpandArg(), except it parses and expands all variable references contained in aBuf.
{
	aOutputVar = aOutputVar->ResolveAlias(); // Necessary for proper detection below of whether it's invalidly used as a source for itself.

	TCHAR var_name[MAX_VAR_NAME_LENGTH + 1];
	Var *var;
	VarSizeType expanded_length;
	size_t var_name_length;
	LPTSTR cp, cp1, dest;

	// Do two passes:
	// #1: Calculate the space needed so that aOutputVar can be given more capacity if necessary.
	// #2: Expand the contents of aBuf into aOutputVar.

	for (int which_pass = 0; which_pass < 2; ++which_pass)
	{
		if (which_pass) // Starting second pass.
		{
			// Set up aOutputVar, enlarging it if necessary.  If it is of type VAR_CLIPBOARD,
			// this call will set up the clipboard for writing:
			if (aOutputVar->AssignString(NULL, expanded_length) != OK)
				return FAIL;
			dest = aOutputVar->Contents();  // Init, and for performance.
		}
		else // First pass.
			expanded_length = 0; // Init prior to accumulation.

		for (cp = aBuf; ; ++cp)  // Increment to skip over the deref/escape just found by the inner for().
		{
			// Find the next escape char or deref symbol:
			for (; *cp && *cp != g_EscapeChar && *cp != g_DerefChar; ++cp)
			{
				if (which_pass) // 2nd pass
					*dest++ = *cp;  // Copy all non-variable-ref characters literally.
				else // just accumulate the length
					++expanded_length;
			}
			if (!*cp) // End of string while scanning/copying.  The current pass is now complete.
				break;
			if (*cp == g_EscapeChar)
			{
				if (which_pass) // 2nd pass
				{
					cp1 = cp + 1;
					switch (*cp1) // See ConvertEscapeSequences() for more details.
					{
						// Only lowercase is recognized for these:
						case 'a': *dest = '\a'; break;  // alert (bell) character
						case 'b': *dest = '\b'; break;  // backspace
						case 'f': *dest = '\f'; break;  // formfeed
						case 'n': *dest = '\n'; break;  // newline
						case 'r': *dest = '\r'; break;  // carriage return
						case 't': *dest = '\t'; break;  // horizontal tab
						case 'v': *dest = '\v'; break;  // vertical tab
						default:  *dest = *cp1; // These other characters are resolved just as they are, including '\0'.
					}
					++dest;
				}
				else
					++expanded_length;
				// Increment cp here and it will be incremented again by the outer loop, i.e. +2.
				// In other words, skip over the escape character, treating it and its target character
				// as a single character.
				++cp;
				continue;
			}
			// Otherwise, it's a dereference symbol, so calculate the size of that variable's contents
			// and add that to expanded_length (or copy the contents into aOutputVar if this is the
			// second pass).
			// Find the reference's ending symbol (don't bother with catching escaped deref chars here
			// -- e.g. %MyVar`% --  since it seems too troublesome to justify given how extremely rarely
			// it would be an issue):
			for (cp1 = cp + 1; *cp1 && *cp1 != g_DerefChar; ++cp1);
			if (!*cp1)    // Since end of string was found, this deref is not correctly terminated.
				continue; // For consistency, omit it entirely.
			var_name_length = cp1 - cp - 1;
			if (var_name_length && var_name_length <= MAX_VAR_NAME_LENGTH)
			{
				tcslcpy(var_name, cp + 1, var_name_length + 1);  // +1 to convert var_name_length to size.
				// Fixed for v1.0.34: Use FindOrAddVar() vs. FindVar() so that environment or built-in
				// variables that aren't directly referenced elsewhere in the script will still work:
				if (   !(var = g_script.FindOrAddVar(var_name, var_name_length))   )
					return FAIL; // Above already displayed the error.
				var = var->ResolveAlias();
				// Don't allow the output variable to be read into itself this way because its contents
				if (var != aOutputVar) // Both of these have had ResolveAlias() called, if required, to make the comparison accurate.
				{
					if (which_pass) // 2nd pass
						dest += var->Get(dest);
					else // just accumulate the length
						expanded_length += var->Get(); // Add in the length of the variable's contents.
				}
			}
			// else since the variable name between the deref symbols is blank or too long: for consistency in behavior,
			// it seems best to omit the dereference entirely (don't put it into aOutputVar).
			cp = cp1; // For the next loop iteration, continue at the char after this reference's final deref symbol.
		} // for()
	} // for() (first and second passes)

	*dest = '\0';  // Terminate the output variable.
	aOutputVar->SetCharLength((VarSizeType)_tcslen(aOutputVar->Contents())); // Update to actual in case estimate was too large.
	return aOutputVar->Close();  // In case it's the clipboard.
}



ResultType Script::DerefInclude(LPTSTR &aOutput, LPTSTR aBuf)
// For #Include and #IncludeAgain.
// Based on Line::Deref above, but with a few differences for backward-compatibility:
//  1) Percent signs that aren't part of a valid deref are not omitted.
//  2) Escape sequences aren't recognized (`; is handled elsewhere).
//  3) It is restricted to built-in vars to reduce the risk of breaking any scripts
//     that use percent sign literally in a filename.  Most other vars are empty anyway.
{
	aOutput = NULL; // Set default.

	TCHAR var_name[MAX_VAR_NAME_LENGTH + 1];
	VarSizeType expanded_length;
	size_t var_name_length;
	LPTSTR cp, cp1, dest;

	// Do two passes:
	// #1: Calculate the space needed.
	// #2: Expand the contents of aBuf into aOutput.

	for (int which_pass = 0; which_pass < 2; ++which_pass)
	{
		if (which_pass) // Starting second pass.
		{
			// Allocate a buffer to contain the result:
			if (  !(aOutput = tmalloc(expanded_length+1))  )
				return FAIL;
			dest = aOutput;
		}
		else // First pass.
			expanded_length = 0; // Init prior to accumulation.

		for (cp = aBuf; *cp; ++cp)  // Increment to skip over the deref/escape just found by the inner for().
		{
			if (*cp == g_DerefChar)
			{
				// It's a dereference symbol, so calculate the size of that variable's contents and add
				// that to expanded_length (or copy the contents into aOutputVar if this is the second pass).
				for (cp1 = cp + 1; *cp1 && *cp1 != g_DerefChar; ++cp1); // Find the reference's ending symbol.
				var_name_length = cp1 - cp - 1;
				if (*cp1 && var_name_length && var_name_length <= MAX_VAR_NAME_LENGTH)
				{
					tcslcpy(var_name, cp + 1, var_name_length + 1);  // +1 to convert var_name_length to size.
					VarEntry *biv = GetBuiltInVar(var_name);  // Only get built-in vars.
					if (biv)
					{
						if (which_pass) // 2nd pass
							dest += biv->type(dest, var_name);
						else
							expanded_length += biv->type(NULL, var_name);
						cp = cp1; // For the next loop iteration, continue at the char after this reference's final deref symbol.
						continue;
					}
				}
				// Since it wasn't a supported deref, copy the deref char into the output:
			}
			if (which_pass) // 2nd pass
				*dest++ = *cp;  // Copy all non-variable-ref characters literally.
			else // just accumulate the length
				++expanded_length;
		} // for()
	} // for() (first and second passes)

	*dest = '\0';  // Terminate the output variable.
	return OK;
}



LPTSTR Line::LogToText(LPTSTR aBuf, int aBufSize) // aBufSize should be an int to preserve negatives from caller (caller relies on this).
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Translates sLog into its text equivalent, putting the result into aBuf and
// returning the position in aBuf of its new string terminator.
// Caller has ensured that aBuf is non-NULL and that aBufSize is reasonable (at least 256).
{
	LPTSTR aBuf_orig = aBuf;

	// Store the position of where each retry done by the outer loop will start writing:
	LPTSTR aBuf_log_start = aBuf + sntprintf(aBuf, aBufSize, _T("Script lines most recently executed (oldest first).")
		_T("  Press [F5] to refresh.  The seconds elapsed between a line and the one after it is in parentheses to")
		_T(" the right (if not 0).  The bottommost line's elapsed time is the number of seconds since it executed.\r\n\r\n"));

	int i, lines_to_show, line_index, line_index2, space_remaining; // space_remaining must be an int to detect negatives.
#ifndef AUTOHOTKEYSC
	int last_file_index = -1;
#endif
	DWORD elapsed;
	bool this_item_is_special, next_item_is_special;

	// In the below, sLogNext causes it to start at the oldest logged line and continue up through the newest:
	for (lines_to_show = LINE_LOG_SIZE, line_index = sLogNext;;) // Retry with fewer lines in case the first attempt doesn't fit in the buffer.
	{
		aBuf = aBuf_log_start; // Reset target position in buffer to the place where log should begin.
		for (next_item_is_special = false, i = 0; i < lines_to_show; ++i, ++line_index)
		{
			if (line_index >= LINE_LOG_SIZE) // wrap around, because sLog is a circular queue
				line_index -= LINE_LOG_SIZE; // Don't just reset it to zero because an offset larger than one may have been added to it.
			if (!sLog[line_index]) // No line has yet been logged in this slot.
				continue; // ACT_LISTLINES and other things might rely on "continue" isntead of halting the loop here.
			this_item_is_special = next_item_is_special;
			next_item_is_special = false;  // Set default.
			if (i + 1 < lines_to_show)  // There are still more lines to be processed
			{
				if (this_item_is_special) // And we know from the above that this special line is not the last line.
					// Due to the fact that these special lines are usually only useful when they appear at the
					// very end of the log, omit them from the log-display when they're not the last line.
					// In the case of a high-frequency SetTimer, this greatly reduces the log clutter that
					// would otherwise occur:
					continue;

				// Since above didn't continue, this item isn't special, so display it normally.
				elapsed = sLogTick[line_index + 1 >= LINE_LOG_SIZE ? 0 : line_index + 1] - sLogTick[line_index];
				if (elapsed > INT_MAX) // INT_MAX is about one-half of DWORD's capacity.
				{
					// v1.0.30.02: Assume that huge values (greater than 24 days or so) were caused by
					// the new policy of storing WinWait/RunWait/etc.'s line in the buffer whenever
					// it was interrupted and later resumed by a thread.  In other words, there are now
					// extra lines in the buffer which are considered "special" because they don't indicate
					// a line that actually executed, but rather one that is still executing (waiting).
					// See ACT_WINWAIT for details.
					next_item_is_special = true; // Override the default.
					if (i + 2 == lines_to_show) // The line after this one is not only special, but the last one that will be shown, so recalculate this one correctly.
						elapsed = GetTickCount() - sLogTick[line_index];
					else // Neither this line nor the special one that follows it is the last.
					{
						// Refer to the line after the next (special) line to get this line's correct elapsed time.
						line_index2 = line_index + 2;
						if (line_index2 >= LINE_LOG_SIZE)
							line_index2 -= LINE_LOG_SIZE;
						elapsed = sLogTick[line_index2] - sLogTick[line_index];
					}
				}
			}
			else // This is the last line (whether special or not), so compare it's time against the current time instead.
				elapsed = GetTickCount() - sLogTick[line_index];
#ifndef AUTOHOTKEYSC
			// If the this line and the previous line are in different files, display the filename:
			if (last_file_index != sLog[line_index]->mFileIndex)
			{
				last_file_index = sLog[line_index]->mFileIndex;
				aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("---- %s\r\n"), sSourceFile[last_file_index]);
			}
#endif
			space_remaining = BUF_SPACE_REMAINING;  // Resolve macro only once for performance.
			// Truncate really huge lines so that the Edit control's size is less likely to be exhausted.
			// In v1.0.30.02, this is even more likely due to having increased the line-buf's capacity from
			// 200 to 400, therefore the truncation point was reduced from 500 to 200 to make it more likely
			// that the first attempt to fit the lines_to_show number of lines into the buffer will succeed.
			aBuf = sLog[line_index]->ToText(aBuf, space_remaining < 200 ? space_remaining : 200, true, elapsed, this_item_is_special);
			// If the line above can't fit everything it needs into the remaining space, it will fill all
			// of the remaining space, and thus the check against LINE_LOG_FINAL_MESSAGE_LENGTH below
			// should never fail to catch that, and then do a retry.
		} // Inner for()

		#define LINE_LOG_FINAL_MESSAGE _T("\r\nPress [F5] to refresh.") // Keep the next line in sync with this.
		#define LINE_LOG_FINAL_MESSAGE_LENGTH 24
		if (BUF_SPACE_REMAINING > LINE_LOG_FINAL_MESSAGE_LENGTH || lines_to_show < 120) // Either success or can't succeed.
			break;

		// Otherwise, there is insufficient room to put everything in, but there's still room to retry
		// with a smaller value of lines_to_show:
		lines_to_show -= 100;
		line_index = sLogNext + (LINE_LOG_SIZE - lines_to_show); // Move the starting point forward in time so that the oldest log entries are omitted.

	} // outer for() that retries the log-to-buffer routine.

	// Must add the return value, not LINE_LOG_FINAL_MESSAGE_LENGTH, in case insufficient room (i.e. in case
	// outer loop terminated due to lines_to_show being too small).
	return aBuf + sntprintf(aBuf, BUF_SPACE_REMAINING, LINE_LOG_FINAL_MESSAGE);
}



LPTSTR Line::VicinityToText(LPTSTR aBuf, int aBufSize) // aBufSize should be an int to preserve negatives from caller (caller relies on this).
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Caller has ensured that aBuf isn't NULL.
// Translates the current line and the lines above and below it into their text equivalent
// putting the result into aBuf and returning the position in aBuf of its new string terminator.
{
	LPTSTR aBuf_orig = aBuf;

	#define LINES_ABOVE_AND_BELOW 7

	// Determine the correct value for line_start and line_end:
	int i;
	Line *line_start, *line_end;
	for (i = 0, line_start = this
		; i < LINES_ABOVE_AND_BELOW && line_start->mPrevLine != NULL
		; ++i, line_start = line_start->mPrevLine);

	for (i = 0, line_end = this
		; i < LINES_ABOVE_AND_BELOW && line_end->mNextLine != NULL
		; ++i, line_end = line_end->mNextLine);

#ifdef AUTOHOTKEYSC
	if (!g_AllowMainWindow) // Override the above to show only a single line, to conceal the script's source code.
	{
		line_start = this;
		line_end = this;
	}
#endif

	// Now line_start and line_end are the first and last lines of the range
	// we want to convert to text, and they're non-NULL.
	aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("\tLine#\n"));

	int space_remaining; // Must be an int to preserve any negative results.

	// Start at the oldest and continue up through the newest:
	for (Line *line = line_start;;)
	{
		if (line == this)
			tcslcpy(aBuf, _T("--->\t"), BUF_SPACE_REMAINING);
		else
			tcslcpy(aBuf, _T("\t"), BUF_SPACE_REMAINING);
		aBuf += _tcslen(aBuf);
		space_remaining = BUF_SPACE_REMAINING;  // Resolve macro only once for performance.
		// Truncate large lines so that the dialog is more readable:
		aBuf = line->ToText(aBuf, space_remaining < 500 ? space_remaining : 500, false);
		if (line == line_end)
			break;
		line = line->mNextLine;
	}
	return aBuf;
}



LPTSTR Line::ToText(LPTSTR aBuf, int aBufSize, bool aCRLF, DWORD aElapsed, bool aLineWasResumed) // aBufSize should be an int to preserve negatives from caller (caller relies on this).
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Caller has ensured that aBuf isn't NULL.
// Translates this line into its text equivalent, putting the result into aBuf and
// returning the position in aBuf of its new string terminator.
{
	if (aBufSize < 3)
		return aBuf;
	else
		aBufSize -= (1 + aCRLF);  // Reserve one char for LF/CRLF after each line (so that it always get added).

	LPTSTR aBuf_orig = aBuf;

	aBuf += sntprintf(aBuf, aBufSize, _T("%03u: "), mLineNumber);
	if (aLineWasResumed)
		aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("STILL WAITING (%0.2f): "), (float)aElapsed / 1000.0);

	if (mActionType == ACT_IFBETWEEN || mActionType == ACT_IFNOTBETWEEN)
		aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("if %s %s %s and %s")
			, *mArg[0].text ? mArg[0].text : VAR(mArg[0])->mName  // i.e. don't resolve dynamic variable names.
			, g_act[mActionType].Name, RAW_ARG2, RAW_ARG3);
	else if (ACT_IS_ASSIGN(mActionType) || (ACT_IS_IF(mActionType) && mActionType < ACT_FIRST_COMMAND))
		aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("%s%s %s %s")
			, ACT_IS_IF(mActionType) ? _T("if ") : _T("")
			, *mArg[0].text ? mArg[0].text : VAR(mArg[0])->mName  // i.e. don't resolve dynamic variable names.
			, g_act[mActionType].Name, RAW_ARG2);
	else if (mActionType == ACT_FOR)
		aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("For %s,%s in %s")
			, *mArg[0].text ? mArg[0].text : VAR(mArg[0])->mName	  // i.e. don't resolve dynamic variable names.
			, *mArg[1].text || !VAR(mArg[1]) ? mArg[1].text : VAR(mArg[1])->mName  // can be omitted.
			, mArg[2].text);
	else
	{
		aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("%s"), g_act[mActionType].Name);
		for (int i = 0; i < mArgc; ++i)
			// This method a little more efficient than using snprintfcat().
			// Also, always use the arg's text for input and output args whose variables haven't
			// been resolved at load-time, since the text has everything in it we want to display
			// and thus there's no need to "resolve" dynamic variables here (e.g. array%i%).
			aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T(",%s"), (mArg[i].type != ARG_TYPE_NORMAL && !*mArg[i].text)
				? VAR(mArg[i])->mName : mArg[i].text);
	}
	if (aElapsed && !aLineWasResumed)
		aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T(" (%0.2f)"), (float)aElapsed / 1000.0);
	// UPDATE for v1.0.25: It seems that MessageBox(), which is the only way these lines are currently
	// displayed, prefers \n over \r\n because otherwise, Ctrl-C on the MsgBox copies the lines all
	// onto one line rather than formatted nicely as separate lines.
	// Room for LF or CRLF was reserved at the top of this function:
	if (aCRLF)
		*aBuf++ = '\r';
	*aBuf++ = '\n';
	*aBuf = '\0';
	return aBuf;
}



void Line::ToggleSuspendState()
{
	// If suspension is being turned on:
	// It seems unnecessary, and possibly undesirable, to purge any pending hotkey msgs from the msg queue.
	// Even if there are some, it's possible that they are exempt from suspension so we wouldn't want to
	// globally purge all messages anyway.
	g_IsSuspended = !g_IsSuspended;
	Hotstring::SuspendAll(g_IsSuspended);  // Must do this prior to ManifestAllHotkeysHotstringsHooks() to avoid incorrect removal of hook.
	Hotkey::ManifestAllHotkeysHotstringsHooks(); // Update the state of all hotkeys based on the complex interdependencies hotkeys have with each another.
	g_script.UpdateTrayIcon();
	CheckMenuItem(GetMenu(g_hWnd), ID_FILE_SUSPEND, g_IsSuspended ? MF_CHECKED : MF_UNCHECKED);
}



void Line::PauseUnderlyingThread(bool aTrueForPauseFalseForUnpause)
{
	if (g <= g_array) // Guard against underflow. This condition can occur when the script thread that called us is the AutoExec section or a callback running in the idle/0 thread.
		return;
	if (g[-1].IsPaused == aTrueForPauseFalseForUnpause) // It's already in the right state.
		return; // Return early because doing the updates further below would be wrong in this case.
	g[-1].IsPaused = aTrueForPauseFalseForUnpause; // Update the pause state to that specified by caller.
	if (aTrueForPauseFalseForUnpause) // The underlying thread is being paused when it was unpaused before.
		++g_nPausedThreads; // For this purpose the idle thread is counted as a paused thread.
	else // The underlying thread is being unpaused when it was paused before.
		--g_nPausedThreads;
}



ResultType Line::ChangePauseState(ToggleValueType aChangeTo, bool aAlwaysOperateOnUnderlyingThread)
// Currently designed to be called only by the Pause command (ACT_PAUSE).
// Returns OK or FAIL.
{
	switch (aChangeTo)
	{
	case TOGGLED_ON:
		break; // By breaking instead of returning, pause will be put into effect further below.
	case TOGGLED_OFF:
		// v1.0.37.06: The old method was to unpause the the nearest paused thread on the call stack;
		// but it was flawed because if the thread that made the flag true is interrupted, and the new
		// thread is paused via the pause command, and that thread is then interrupted, when the paused
		// thread resumes it would automatically and wrongly be unpaused (i.e. the unpause ticket would
		// be used at a level higher in the call stack than intended).
		// Flag this thread so that when it ends, the thread beneath it will be unpaused.  If that thread
		// (which can be the idle thread) isn't paused the following flag-change will be ignored at a later
		// stage. This method also relies on the fact that the current thread cannot itself be paused right
		// now because it is what got us here.
		PauseUnderlyingThread(false); // Necessary even for the "idle thread" (otherwise, the Pause command wouldn't be able to unpause it).
		return OK;
	case NEUTRAL: // the user omitted the parameter entirely, which is considered the same as "toggle"
	case TOGGLE:
		// Update for v1.0.37.06: "Pause" and "Pause Toggle" are more useful if they always apply to the
		// thread immediately beneath the current thread rather than "any underlying thread that's paused".
		if (g > g_array && g[-1].IsPaused) // Checking g>g_array avoids any chance of underflow, which might otherwise happen if this is called by the AutoExec section or a threadless callback running in thread #0.
		{
			PauseUnderlyingThread(false);
			return OK;
		}
		//ELSE since the underlying thread is not paused, continue onward to do the "pause enabled" logic below.
		// (This is the historical behavior because it allows a hotkey like F1::Pause to toggle the script's
		// pause state on and off -- even though what's really happening involves multiple threads.)
		break;
	default: // TOGGLE_INVALID or some other disallowed value.
		// We know it's a variable because otherwise the loading validation would have caught it earlier:
		return LineError(ERR_PARAM1_INVALID, FAIL, ARG1);
	}

	// Since above didn't return, pause should be turned on.
	if (aAlwaysOperateOnUnderlyingThread) // v1.0.37.06: Allow underlying thread to be directly paused rather than pausing the current thread.
	{
		PauseUnderlyingThread(true); // If the underlying thread is already paused, this flag change will be ignored at a later stage.
		return OK;
	}
	// Otherwise, pause the current subroutine (which by definition isn't paused since it had to be 
	// active to call us).  It seems best not to attempt to change the Hotkey mRunAgainAfterFinished
	// attribute for the current hotkey (assuming it's even a hotkey that got us here) or
	// for them all.  This is because it's conceivable that this Pause command occurred
	// in a background thread, such as a timed subroutine, in which case we wouldn't want the
	// pausing of that thread to affect anything else the user might be doing with hotkeys.
	// UPDATE: The above is flawed because by definition the script's quasi-thread that got
	// us here is now active.  Since it is active, the script will immediately become dormant
	// when this is executed, waiting for the user to press a hotkey to launch a new
	// quasi-thread.  Thus, it seems best to reset all the mRunAgainAfterFinished flags
	// in case we are in a hotkey subroutine and in case this hotkey has a buffered repeat-again
	// action pending, which the user probably wouldn't want to happen after the script is unpaused:
	Hotkey::ResetRunAgainAfterFinished();
	g->IsPaused = true;
	++g_nPausedThreads; // For this purpose the idle thread is counted as a paused thread.
	g_script.UpdateTrayIcon();
	return OK;
}



ResultType Line::ScriptBlockInput(bool aEnable)
// Always returns OK for caller convenience.
{
	// Must be running Win98/2000+ for this function to be successful.
	// We must dynamically load the function to retain compatibility with Win95 (program won't launch
	// at all otherwise).
	typedef void (CALLBACK *BlockInput)(BOOL);
	static BlockInput lpfnDLLProc = (BlockInput)GetProcAddress(GetModuleHandle(_T("user32")), "BlockInput");
	// Always turn input ON/OFF even if g_BlockInput says its already in the right state.  This is because
	// BlockInput can be externally and undetectably disabled, e.g. if the user presses Ctrl-Alt-Del:
	if (lpfnDLLProc)
		(*lpfnDLLProc)(aEnable ? TRUE : FALSE);
	g_BlockInput = aEnable;
	return OK;  // By design, it never returns FAIL.
}



Line *Line::PreparseError(LPTSTR aErrorText, LPTSTR aExtraInfo)
// Returns a different type of result for use with the Pre-parsing methods.
{
	// Make all preparsing errors critical because the runtime reliability
	// of the program relies upon the fact that the aren't any kind of
	// problems in the script (otherwise, unexpected behavior may result).
	// Update: It's okay to return FAIL in this case.  CRITICAL_ERROR should
	// be avoided whenever OK and FAIL are sufficient by themselves, because
	// otherwise, callers can't use the NOT operator to detect if a function
	// failed (since FAIL is value zero, but CRITICAL_ERROR is non-zero):
	LineError(aErrorText, FAIL, aExtraInfo);
	return NULL; // Always return NULL because the callers use it as their return value.
}


IObject *Line::CreateRuntimeException(LPCTSTR aErrorText, LPCTSTR aWhat, LPCTSTR aExtraInfo)
{
	// Build the parameters for Object::Create()
	ExprTokenType aParams[5*2]; int aParamCount = 4*2;
	ExprTokenType* aParam[5*2] = { aParams + 0, aParams + 1, aParams + 2, aParams + 3, aParams + 4
		, aParams + 5, aParams + 6, aParams + 7, aParams + 8, aParams + 9 };
	aParams[0].symbol = SYM_STRING;  aParams[0].marker = _T("What");
	aParams[1].symbol = SYM_STRING;  aParams[1].marker = aWhat ? (LPTSTR)aWhat : g_act[mActionType].Name;
	aParams[2].symbol = SYM_STRING;  aParams[2].marker = _T("File");
	aParams[3].symbol = SYM_STRING;  aParams[3].marker = Line::sSourceFile[mFileIndex];
	aParams[4].symbol = SYM_STRING;  aParams[4].marker = _T("Line");
	aParams[5].symbol = SYM_INTEGER; aParams[5].value_int64 = mLineNumber;
	aParams[6].symbol = SYM_STRING;  aParams[6].marker = _T("Message");
	aParams[7].symbol = SYM_STRING;  aParams[7].marker = (LPTSTR)aErrorText;
	if (aExtraInfo && *aExtraInfo)
	{
		aParamCount += 2;
		aParams[8].symbol = SYM_STRING; aParams[8].marker = _T("Extra");
		aParams[9].symbol = SYM_STRING; aParams[9].marker = (LPTSTR)aExtraInfo;
	}

	return Object::Create(aParam, aParamCount);
}


ResultType Line::ThrowRuntimeException(LPCTSTR aErrorText, LPCTSTR aWhat, LPCTSTR aExtraInfo)
{
	// ThrownToken should only be non-NULL while control is being passed up the
	// stack, which implies no script code can be executing.
	ASSERT(!g->ThrownToken);

	ExprTokenType *token;
	if (   !(token = new ExprTokenType)
		|| !(token->object = CreateRuntimeException(aErrorText, aWhat, aExtraInfo))   )
	{
		// Out of memory. It's likely that we were called for this very reason.
		// Since we don't even have enough memory to allocate an exception object,
		// just show an error message and exit the thread. Don't call LineError(),
		// since that would recurse into this function.
		if (token)
			delete token;
		MsgBox(ERR_OUTOFMEM ERR_ABORT);
		return FAIL;
	}

	token->symbol = SYM_OBJECT;
	token->mem_to_free = NULL;

	g->ThrownToken = token;
	if (!(g->ExcptMode & EXCPTMODE_CATCH))
		g_script.UnhandledException(this);

	// Returning FAIL causes each caller to also return FAIL, until either the
	// thread has fully exited or the recursion layer handling ACT_TRY is reached:
	return FAIL;
}

ResultType Script::ThrowRuntimeException(LPCTSTR aErrorText, LPCTSTR aWhat, LPCTSTR aExtraInfo)
{
	return g_script.mCurrLine->ThrowRuntimeException(aErrorText, aWhat, aExtraInfo);
}


ResultType Line::SetErrorLevelOrThrowBool(bool aError)
{
	if (!aError)
		return g_ErrorLevel->Assign(ERRORLEVEL_NONE);
	if (!g->InTryBlock())
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	// Otherwise, an error occurred and there is a try block, so throw an exception:
	return ThrowRuntimeException(ERRORLEVEL_ERROR);
}

ResultType Line::SetErrorLevelOrThrowStr(LPCTSTR aErrorValue)
{
	if ((*aErrorValue == '0' && !aErrorValue[1]) || !g->InTryBlock())
		return g_ErrorLevel->Assign(aErrorValue);
	// Otherwise, an error occurred and there is a try block, so throw an exception:
	return ThrowRuntimeException(aErrorValue);
}

ResultType Line::SetErrorLevelOrThrowInt(int aErrorValue)
{
	if (!aErrorValue || !g->InTryBlock())
		return g_ErrorLevel->Assign(aErrorValue);
	TCHAR buf[12];
	// Otherwise, an error occurred and there is a try block, so throw an exception:
	return ThrowRuntimeException(_itot(aErrorValue, buf, 10));
}

// Logic from the above functions is duplicated in the below functions rather than calling
// g_script.mCurrLine->SetErrorLevelOrThrow() to squeeze out a little extra performance for
// "success" cases.  These are done as overloads vs making aMessage optional to reduce code
// size, since they're called from numerous places.  (Even omitted parameters are passed
// "explicitly" in the compiled code.)

ResultType Script::SetErrorLevelOrThrowBool(bool aError)
{
	if (!aError)
		return g_ErrorLevel->Assign(ERRORLEVEL_NONE);
	if (!g->InTryBlock())
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	// Otherwise, an error occurred and there is a try block, so throw an exception:
	return ThrowRuntimeException(ERRORLEVEL_ERROR);
}

ResultType Script::SetErrorLevelOrThrowStr(LPCTSTR aErrorValue, LPCTSTR aWhat)
{
	if ((*aErrorValue == '0' && !aErrorValue[1]) || !g->InTryBlock())
		return g_ErrorLevel->Assign(aErrorValue);
	// Otherwise, an error occurred and there is a try block, so throw an exception:
	return ThrowRuntimeException(aErrorValue, aWhat);
}

ResultType Script::SetErrorLevelOrThrowInt(int aErrorValue, LPCTSTR aWhat)
{
	if (!aErrorValue || !g->InTryBlock())
		return g_ErrorLevel->Assign(aErrorValue);
	TCHAR buf[12];
	// Otherwise, an error occurred and there is a try block, so throw an exception:
	return ThrowRuntimeException(_itot(aErrorValue, buf, 10), aWhat);
}


ResultType Line::SetErrorsOrThrow(bool aError, DWORD aLastErrorOverride)
{
	// LastError is set even if we're going to throw an exception, for simplicity:
	g->LastError = aLastErrorOverride == -1 ? GetLastError() : aLastErrorOverride;
	return SetErrorLevelOrThrowBool(aError);
}


#define ERR_PRINT(fmt, ...) _ftprintf(stderr, fmt, __VA_ARGS__)

ResultType Line::LineError(LPCTSTR aErrorText, ResultType aErrorType, LPCTSTR aExtraInfo)
{
	if (!aErrorText)
		aErrorText = _T("");
	if (!aExtraInfo)
		aExtraInfo = _T("");

	if ((g->ExcptMode || g_script.mOnError.Count()) // OnError also needs an exception object.
		&& (aErrorType == FAIL || aErrorType == EARLY_EXIT)) // FAIL is most common, but EARLY_EXIT is used by ComError(). WARN and CRITICAL_ERROR are excluded.
		return ThrowRuntimeException(aErrorText, NULL, aExtraInfo);

	if (g_script.mErrorStdOut && !g_script.mIsReadyToExecute && aErrorType != WARN) // i.e. runtime errors are always displayed via dialog.
	{
		// JdeB said:
		// Just tested it in Textpad, Crimson and Scite. they all recognise the output and jump
		// to the Line containing the error when you double click the error line in the output
		// window (like it works in C++).  Had to change the format of the line to: 
		// printf("%s (%d) : ==> %s: \n%s \n%s\n",szInclude, nAutScriptLine, szText, szScriptLine, szOutput2 );
		// MY: Full filename is required, even if it's the main file, because some editors (EditPlus)
		// seem to rely on that to determine which file and line number to jump to when the user double-clicks
		// the error message in the output window.
		// v1.0.47: Added a space before the colon as originally intended.  Toralf said, "With this minor
		// change the error lexer of Scite recognizes this line as a Microsoft error message and it can be
		// used to jump to that line."
		#define STD_ERROR_FORMAT _T("%s (%d) : ==> %s\n")
		ERR_PRINT(STD_ERROR_FORMAT, sSourceFile[mFileIndex], mLineNumber, aErrorText); // printf() does not significantly increase the size of the EXE, probably because it shares most of the same code with sprintf(), etc.
		if (*aExtraInfo)
			ERR_PRINT(_T("     Specifically: %s\n"), aExtraInfo);
	}
	else
	{
		TCHAR buf[MSGBOX_TEXT_SIZE];
		FormatError(buf, _countof(buf), aErrorType, aErrorText, aExtraInfo, this
			// The last parameter determines the final line of the message:
			, (aErrorType == FAIL) ? (g_script.mIsReadyToExecute ? ERR_ABORT_NO_SPACES
									: (g_script.mIsRestart ? OLD_STILL_IN_EFFECT : WILL_EXIT))
			: (aErrorType == CRITICAL_ERROR) ? UNSTABLE_WILL_EXIT
			: (aErrorType == EARLY_EXIT) ? _T("Continue running the script?")
			: _T("For more details, read the documentation for #Warn."));
		
		g_script.mCurrLine = this;  // This needs to be set in some cases where the caller didn't.
		
#ifdef CONFIG_DEBUGGER
		if (g_Debugger.HasStdErrHook())
			g_Debugger.OutputDebug(buf);
		else
#endif
		if (MsgBox(buf, MB_TOPMOST | (aErrorType == EARLY_EXIT ? MB_YESNO : 0)) == IDNO)
			// The user was asked "Continue running the script?" and answered "No".
			// This will attempt to run the OnExit subroutine, which should be okay since that
			// subroutine will terminate the script if it encounters another runtime error:
			g_script.ExitApp(EXIT_ERROR);
	}

	if (aErrorType == CRITICAL_ERROR && g_script.mIsReadyToExecute)
		// Pass EXIT_DESTROY to ensure the program always exits, regardless of OnExit.
		g_script.ExitApp(EXIT_DESTROY);

	return aErrorType; // The caller told us whether it should be a critical error or not.
}



int Line::FormatError(LPTSTR aBuf, int aBufSize, ResultType aErrorType, LPCTSTR aErrorText, LPCTSTR aExtraInfo, Line *aLine, LPCTSTR aFooter)
{
	TCHAR source_file[MAX_PATH * 2];
	if (aLine && aLine->mFileIndex)
		sntprintf(source_file, _countof(source_file), _T(" in #include file \"%s\""), sSourceFile[aLine->mFileIndex]);
	else
		*source_file = '\0'; // Don't bother cluttering the display if it's the main script file.

	LPTSTR aBuf_orig = aBuf;
	// Error message:
	aBuf += sntprintf(aBuf, aBufSize, _T("%s%s:%s %-1.500s\n\n")  // Keep it to a sane size in case it's huge.
		, aErrorType == WARN ? _T("Warning") : (aErrorType == CRITICAL_ERROR ? _T("Critical Error") : _T("Error"))
		, source_file, *source_file ? _T("\n    ") : _T(" "), aErrorText);
	// Specifically:
	if (*aExtraInfo)
		// Use format specifier to make sure really huge strings that get passed our
		// way, such as a var containing clipboard text, are kept to a reasonable size:
		aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("Specifically: %-1.100s%s\n\n")
			, aExtraInfo, _tcslen(aExtraInfo) > 100 ? _T("...") : _T(""));
	// Relevant lines of code:
	if (aLine)
		aBuf = aLine->VicinityToText(aBuf, BUF_SPACE_REMAINING);
	// What now?:
	if (aFooter)
		aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("\n%s"), aFooter);
	
	return (int)(aBuf - aBuf_orig);
}



ResultType Script::ScriptError(LPCTSTR aErrorText, LPCTSTR aExtraInfo) //, ResultType aErrorType)
// Even though this is a Script method, including it here since it shares
// a common theme with the other error-displaying functions:
{
	if (mCurrLine)
		// If a line is available, do LineError instead since it's more specific.
		// If an error occurs before the script is ready to run, assume it's always critical
		// in the sense that the program will exit rather than run the script.
		// Update: It's okay to return FAIL in this case.  CRITICAL_ERROR should
		// be avoided whenever OK and FAIL are sufficient by themselves, because
		// otherwise, callers can't use the NOT operator to detect if a function
		// failed (since FAIL is value zero, but CRITICAL_ERROR is non-zero):
		return mCurrLine->LineError(aErrorText, FAIL, aExtraInfo);
	// Otherwise: The fact that mCurrLine is NULL means that the line currently being loaded
	// has not yet been successfully added to the linked list.  Such errors will always result
	// in the program exiting.
	if (!aErrorText)
		aErrorText = _T("Unk"); // Placeholder since it shouldn't be NULL.
	if (!aExtraInfo) // In case the caller explicitly called it with NULL.
		aExtraInfo = _T("");

	if (g_script.mErrorStdOut && !g_script.mIsReadyToExecute) // i.e. runtime errors are always displayed via dialog.
	{
		// See LineError() for details.
		ERR_PRINT(STD_ERROR_FORMAT, Line::sSourceFile[mCurrFileIndex], mCombinedLineNumber, aErrorText);
		if (*aExtraInfo)
			ERR_PRINT(_T("     Specifically: %s\n"), aExtraInfo);
	}
	else
	{
		TCHAR buf[MSGBOX_TEXT_SIZE], *cp = buf;
		int buf_space_remaining = (int)_countof(buf);

		cp += sntprintf(cp, buf_space_remaining, _T("Error at line %u"), mCombinedLineNumber); // Don't call it "critical" because it's usually a syntax error.
		buf_space_remaining = (int)(_countof(buf) - (cp - buf));

		if (mCurrFileIndex)
		{
			cp += sntprintf(cp, buf_space_remaining, _T(" in #include file \"%s\""), Line::sSourceFile[mCurrFileIndex]);
			buf_space_remaining = (int)(_countof(buf) - (cp - buf));
		}
		//else don't bother cluttering the display if it's the main script file.

		cp += sntprintf(cp, buf_space_remaining, _T(".\n\n"));
		buf_space_remaining = (int)(_countof(buf) - (cp - buf));

		if (*aExtraInfo)
		{
			cp += sntprintf(cp, buf_space_remaining, _T("Line Text: %-1.100s%s\nError: ")  // i.e. the word "Error" is omitted as being too noisy when there's no ExtraInfo to put into the dialog.
				, aExtraInfo // aExtraInfo defaults to "" so this is safe.
				, _tcslen(aExtraInfo) > 100 ? _T("...") : _T(""));
			buf_space_remaining = (int)(_countof(buf) - (cp - buf));
		}
		sntprintf(cp, buf_space_remaining, _T("%s\n\n%s"), aErrorText, mIsRestart ? OLD_STILL_IN_EFFECT : WILL_EXIT);

		//ShowInEditor();
#ifdef CONFIG_DEBUGGER
		if (g_Debugger.HasStdErrHook())
			g_Debugger.OutputDebug(buf);
		else
#endif
		MsgBox(buf);
	}
	return FAIL; // See above for why it's better to return FAIL than CRITICAL_ERROR.
}



ResultType Script::CriticalError(LPCTSTR aErrorText, LPCTSTR aExtraInfo)
{
	g->ExcptMode = EXCPTMODE_NONE; // Do not throw an exception.
	if (mCurrLine)
		mCurrLine->LineError(aErrorText, CRITICAL_ERROR, aExtraInfo);
	// mCurrLine should always be non-NULL during runtime, and CRITICAL_ERROR should
	// cause LineError() to exit even if an OnExit routine is present, so this is here
	// mainly for maintainability.
	TerminateApp(EXIT_DESTROY, CRITICAL_ERROR);
	return FAIL; // Never executed.
}



ResultType Script::UnhandledException(Line* aLine)
{
	LPCTSTR message = _T(""), extra = _T("");
	TCHAR extra_buf[MAX_NUMBER_SIZE], message_buf[MAX_NUMBER_SIZE];

	global_struct &g = *::g;

	// OnError: Allow the script to handle it via a global callback.
	static bool sOnErrorRunning = false;
	if (g_script.mOnError.Count() && !sOnErrorRunning)
	{
		ExprTokenType *token = g.ThrownToken;
		g.ThrownToken = NULL; // Allow the callback to execute correctly.
		sOnErrorRunning = true;
		bool returned_true = g_script.mOnError.Call(token, 1, INT_MAX) == CONDITION_TRUE;
		sOnErrorRunning = false;
		if (g.ThrownToken) // An exception was thrown by the callback.
		{
			// UnhandledException() has already been called recursively for g.ThrownToken,
			// so don't show a second error message.  This allows `throw param1` to mean
			// "abort all OnError callbacks and show default message now".
			FreeExceptionToken(token);
			return FAIL;
		}
		// Some callers rely on g.ThrownToken!=NULL to unwind the stack, so it is restored
		// rather than freeing it immediately.  If the exception object has __Delete, it
		// will be called after the stack unwinds.
		g.ThrownToken = token;
		if (returned_true)
			return FAIL;
	}

	if (Object *ex = dynamic_cast<Object *>(TokenToObject(*g.ThrownToken)))
	{
		// For simplicity and safety, we call into the Object directly rather than via Invoke().
		ExprTokenType t;
		if (ex->GetItem(t, _T("Message")))
			message = TokenToString(t, message_buf);
		if (ex->GetItem(t, _T("Extra")))
			extra = TokenToString(t, extra_buf);
		if (ex->GetItem(t, _T("Line")))
		{
			LineNumberType line_no = (LineNumberType)TokenToInt64(t);
			if (ex->GetItem(t, _T("File")))
			{
				LPCTSTR file = TokenToString(t);
				// Locate the line by number and file index, then display that line instead
				// of the caller supplied one since it's probably more relevant.
				int file_index;
				for (file_index = 0; file_index < Line::sSourceFileCount; ++file_index)
					if (!_tcsicmp(file, Line::sSourceFile[file_index]))
						break;
				Line *line;
				for (line = g_script.mFirstLine;
					line && (line->mLineNumber != line_no || line->mFileIndex != file_index);
					line = line->mNextLine);
				if (line)
					aLine = line;
			}
		}
	}
	else
	{
		// Assume it's a string or number.
		message = TokenToString(*g.ThrownToken, message_buf);
	}

	// If message is empty or numeric, display a default message for clarity.
	if (!*extra && IsPureNumeric(message, TRUE, TRUE, TRUE))
	{
		extra = message;
		message = _T("Unhandled exception.");
	}	

	TCHAR buf[MSGBOX_TEXT_SIZE];
	Line::FormatError(buf, _countof(buf), FAIL, message, extra, aLine
		, (g.ExcptMode & EXCPTMODE_DELETE) ? ERR_ABORT_DELETE : ERR_ABORT_NO_SPACES);
	MsgBox(buf);

	return FAIL;
}

void Script::FreeExceptionToken(ExprTokenType*& aToken)
{
	// If an object was thrown, release it.
	if (aToken->symbol == SYM_OBJECT)
		aToken->object->Release();
	// If a string was thrown and memory allocated for it, free it.
	if (aToken->mem_to_free)
		free(aToken->mem_to_free);
	// Free the token itself.
	delete aToken;
	// Clear caller's variable.
	aToken = NULL;
}


void Script::ScriptWarning(WarnMode warnMode, LPCTSTR aWarningText, LPCTSTR aExtraInfo, Line *line)
{
	if (warnMode == WARNMODE_OFF)
		return;

	if (!line) line = mCurrLine;
	int fileIndex = line ? line->mFileIndex : mCurrFileIndex;
	FileIndexType lineNumber = line ? line->mLineNumber : mCombinedLineNumber;

	TCHAR buf[MSGBOX_TEXT_SIZE], *cp = buf;
	int buf_space_remaining = (int)_countof(buf);
	
	#define STD_WARNING_FORMAT _T("%s (%d) : ==> Warning: %s\n")
	cp += sntprintf(cp, buf_space_remaining, STD_WARNING_FORMAT, Line::sSourceFile[fileIndex], lineNumber, aWarningText);
	buf_space_remaining = (int)(_countof(buf) - (cp - buf));

	if (*aExtraInfo)
	{
		cp += sntprintf(cp, buf_space_remaining, _T("     Specifically: %s\n"), aExtraInfo);
		buf_space_remaining = (int)(_countof(buf) - (cp - buf));
	}

	if (warnMode == WARNMODE_STDOUT)
#ifndef CONFIG_DEBUGGER
		_fputts(buf, stdout);
	else
		OutputDebugString(buf);
#else
		g_Debugger.FileAppendStdOut(buf) || _fputts(buf, stdout);
	else
		g_Debugger.OutputDebug(buf);
#endif

	// In MsgBox mode, MsgBox is in addition to OutputDebug
	if (warnMode == WARNMODE_MSGBOX)
	{
		if (!line)
			line = mCurrLine; // Call mCurrLine->LineError() vs ScriptError() to pass WARN.
		if (line)
			line->LineError(aWarningText, WARN, aExtraInfo);
		else
			// Realistically shouldn't happen.  If it does, the message might be slightly
			// misleading since ScriptError isn't equipped to display "warning" messages.
			ScriptError(aWarningText, aExtraInfo);
	}
}



void Script::WarnUninitializedVar(Var *var)
{
	bool isGlobal = !var->IsLocal();
	WarnMode warnMode = isGlobal ? g_Warn_UseUnsetGlobal : g_Warn_UseUnsetLocal;
	if (!warnMode)
		return;

	// Note: If warning mode is MsgBox, this method has side effect of marking the var initialized, so that
	// only a single message box gets raised per variable.  (In other modes, e.g. OutputDebug, the var remains
	// uninitialized because it may be beneficial to see the quantity and various locations of uninitialized
	// uses, and doesn't present the same user interface problem that multiple message boxes can.)
	if (warnMode == WARNMODE_MSGBOX)
		var->MarkInitialized();

	bool isNonStaticLocal = var->IsNonStaticLocal();
	LPCTSTR varClass = isNonStaticLocal ? _T("local") : (isGlobal ? _T("global") : _T("static"));
	LPCTSTR sameNameAsGlobal = (isNonStaticLocal && FindVar(var->mName, 0, NULL, FINDVAR_GLOBAL)) ? _T(" with same name as a global") : _T("");
	TCHAR buf[DIALOG_TITLE_SIZE], *cp = buf;

	int buf_space_remaining = (int)_countof(buf);
	sntprintf(cp, buf_space_remaining, _T("%s  (a %s variable%s)"), var->mName, varClass, sameNameAsGlobal);

	ScriptWarning(warnMode, WARNING_USE_UNSET_VARIABLE, buf);
}



void Script::MaybeWarnLocalSameAsGlobal(Func &func, Var &var)
// Caller has verified the following:
//  1) var is not a declared variable.
//  2) a global variable with the same name definitely exists.
{
	if (!g_Warn_LocalSameAsGlobal)
		return;

#ifdef ENABLE_DLLCALL
	if (IsDllArgTypeName(var.mName))
		// Exclude unquoted DllCall arg type names.  Although variable names like "str" and "ptr"
		// might be used for other purposes, it seems far more likely that both this var and its
		// global counterpart (if it exists) are blank vars which were used as DllCall arg types.
		return;
#endif

	Line *line = func.mJumpToLine;
	while (line && line->mActionType != ACT_BLOCK_BEGIN) line = line->mPrevLine;
	if (!line) line = func.mJumpToLine;

	TCHAR buf[DIALOG_TITLE_SIZE], *cp = buf;
	int buf_space_remaining = (int)_countof(buf);
	sntprintf(cp, buf_space_remaining, _T("%s  (in function %s)"), var.mName, func.mName);
	
	ScriptWarning(g_Warn_LocalSameAsGlobal, WARNING_LOCAL_SAME_AS_GLOBAL, buf, line);
}



void Script::PreprocessLocalVars(Func &aFunc, Var **aVarList, int &aVarCount)
{
	for (int v = 0; v < aVarCount; ++v)
	{
		Var &var = *aVarList[v];
		if (var.IsDeclared()) // Not a canditate for a super-global or warning.
			continue;
		Var *global_var = FindVar(var.mName, 0, NULL, FINDVAR_GLOBAL);
		if (!global_var) // No global variable with that name.
			continue;
		if (global_var->IsSuperGlobal())
		{
			// Make this local variable an alias for the super-global. Above has already
			// verified this var was not declared and therefore isn't a function parameter.
			var.UpdateAlias(global_var);
			// Remove the variable from the local list to prevent it from being shown in
			// ListVars or being reset when the function returns.
			memmove(aVarList + v, aVarList + v + 1, (--aVarCount - v) * sizeof(Var *));
			--v; // Counter the loop's increment.
		}
		else
		// Since this undeclared local variable has the same name as a global, there's
		// a chance the user intended it to be global. So consider warning the user:
		MaybeWarnLocalSameAsGlobal(aFunc, var);
	}
}



void Script::CheckForClassOverwrite()
{
	// Aside from class variables, A_Args is the only variable which can contain an object
	// at this stage.  Excluding it this way produces smaller code than checking for the
	// "__class" key within whatever object is found.
	Var *a_args = FindVar(_T("A_Args"));

	for (Line *line = mFirstLine; line; line = line->mNextLine)
	{
		for (int a = 0; a < line->mArgc; ++a)
		{
			ArgStruct &arg = line->mArg[a];
			if (arg.type == ARG_TYPE_OUTPUT_VAR)
			{
				if (!*arg.text) // The arg's variable is not one that needs to be dynamically resolved.
				{
					Var *target_var = VAR(arg);
					if (target_var->HasObject() && target_var != a_args)
						ScriptWarning(g_Warn_ClassOverwrite, WARNING_CLASS_OVERWRITE, target_var->mName, line);
				}
			}
			else if (arg.is_expression)
			{
				for (ExprTokenType *token = arg.postfix; token->symbol != SYM_INVALID; ++token)
				{
					if (token->symbol == SYM_VAR && token->is_lvalue && token->var->HasObject() && token->var != a_args)
						ScriptWarning(g_Warn_ClassOverwrite, WARNING_CLASS_OVERWRITE, token->var->mName, line);
				}
			}
		}
	}
}



LPTSTR Script::ListVars(LPTSTR aBuf, int aBufSize) // aBufSize should be an int to preserve negatives from caller (caller relies on this).
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Translates this script's list of variables into text equivalent, putting the result
// into aBuf and returning the position in aBuf of its new string terminator.
{
	LPTSTR aBuf_orig = aBuf;
	Func *current_func = g->CurrentFunc ? g->CurrentFunc : g->CurrentFuncGosub;
	if (current_func)
	{
		// This definition might help compiler string pooling by ensuring it stays the same for both usages:
		#define LIST_VARS_UNDERLINE _T("\r\n--------------------------------------------------\r\n")
		// Start at the oldest and continue up through the newest:
		aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("Local Variables for %s()%s"), current_func->mName, LIST_VARS_UNDERLINE);
		Func &func = *current_func; // For performance.
		for (int i = 0; i < func.mVarCount; ++i)
			if (func.mVar[i]->Type() == VAR_NORMAL) // Don't bother showing clipboard and other built-in vars.
				aBuf = func.mVar[i]->ToText(aBuf, BUF_SPACE_REMAINING, true);
	}
	// v1.0.31: The description "alphabetical" is kept even though it isn't quite true
	// when the lazy variable list exists, since those haven't yet been sorted into the main list.
	// However, 99.9% of scripts do not use the lazy list, so it seems too rare to worry about other
	// than document it in the ListVars command in the help file:
	aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("%sGlobal Variables (alphabetical)%s")
		, current_func ? _T("\r\n\r\n") : _T(""), LIST_VARS_UNDERLINE);
	// Start at the oldest and continue up through the newest:
	for (int i = 0; i < mVarCount; ++i)
		if (mVar[i]->Type() == VAR_NORMAL) // Don't bother showing clipboard and other built-in vars.
			aBuf = mVar[i]->ToText(aBuf, BUF_SPACE_REMAINING, true);
	return aBuf;
}



LPTSTR Script::ListKeyHistory(LPTSTR aBuf, int aBufSize) // aBufSize should be an int to preserve negatives from caller (caller relies on this).
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Translates this key history into text equivalent, putting the result
// into aBuf and returning the position in aBuf of its new string terminator.
{
	LPTSTR aBuf_orig = aBuf; // Needed for the BUF_SPACE_REMAINING macro.
	// I was initially concerned that GetWindowText() can hang if the target window is
	// hung.  But at least on newer OS's, this doesn't seem to be a problem: MSDN says
	// "If the window does not have a caption, the return value is a null string. This
	// behavior is by design. It allows applications to call GetWindowText without hanging
	// if the process that owns the target window is hung. However, if the target window
	// is hung and it belongs to the calling application, GetWindowText will hang the
	// calling application."
	HWND target_window = GetForegroundWindow();
	TCHAR win_title[100];
	if (target_window)
		GetWindowText(target_window, win_title, _countof(win_title));
	else
		*win_title = '\0';

	TCHAR timer_list[128];
	*timer_list = '\0';
	for (ScriptTimer *timer = mFirstTimer; timer != NULL; timer = timer->mNextTimer)
		if (timer->mEnabled)
			sntprintfcat(timer_list, _countof(timer_list) - 3, _T("%s "), timer->mLabel->Name()); // Allow room for "..."
	if (*timer_list)
	{
		size_t length = _tcslen(timer_list);
		if (length > (_countof(timer_list) - 5))
			tcslcpy(timer_list + length, _T("..."), _countof(timer_list) - length);
		else if (timer_list[length - 1] == ' ')
			timer_list[--length] = '\0';  // Remove the last space if there was room enough for it to have been added.
	}

	TCHAR LRtext[256];
	aBuf += sntprintf(aBuf, aBufSize,
		_T("Window: %s")
		//"\r\nBlocks: %u"
		_T("\r\nKeybd hook: %s")
		_T("\r\nMouse hook: %s")
		_T("\r\nEnabled Timers: %u of %u (%s)")
		//"\r\nInterruptible?: %s"
		_T("\r\nInterrupted threads: %d%s")
		_T("\r\nPaused threads: %d of %d (%d layers)")
		_T("\r\nModifiers (GetKeyState() now) = %s")
		_T("\r\n")
		, win_title
		//, SimpleHeap::GetBlockCount()
		, g_KeybdHook == NULL ? _T("no") : _T("yes")
		, g_MouseHook == NULL ? _T("no") : _T("yes")
		, mTimerEnabledCount, mTimerCount, timer_list
		//, INTERRUPTIBLE ? "yes" : "no"
		, g_nThreads > 1 ? g_nThreads - 1 : 0
		, g_nThreads > 1 ? _T(" (preempted: they will resume when the current thread finishes)") : _T("")
		, g_nPausedThreads - (g_array[0].IsPaused && !mAutoExecSectionIsRunning)  // Historically thread #0 isn't counted as a paused thread unless the auto-exec section is running but paused.
		, g_nThreads, g_nLayersNeedingTimer
		, ModifiersLRToText(GetModifierLRState(true), LRtext));
	GetHookStatus(aBuf, BUF_SPACE_REMAINING);
	aBuf += _tcslen(aBuf); // Adjust for what GetHookStatus() wrote to the buffer.
	return aBuf + sntprintf(aBuf, BUF_SPACE_REMAINING, g_KeyHistory ? _T("\r\nPress [F5] to refresh.")
		: _T("\r\nKey History has been disabled via #KeyHistory 0."));
}



ResultType Script::ActionExec(LPTSTR aAction, LPTSTR aParams, LPTSTR aWorkingDir, bool aDisplayErrors
	, LPTSTR aRunShowMode, HANDLE *aProcess, bool aUpdateLastError, bool aUseRunAs, Var *aOutputVar)
// Caller should specify NULL for aParams if it wants us to attempt to parse out params from
// within aAction.  Caller may specify empty string ("") instead to specify no params at all.
// Remember that aAction and aParams can both be NULL, so don't dereference without checking first.
// Note: For the Run & RunWait commands, aParams should always be NULL.  Params are parsed out of
// the aActionString at runtime, here, rather than at load-time because Run & RunWait might contain
// dereferenced variable(s), which can only be resolved at runtime.
{
	HANDLE hprocess_local;
	HANDLE &hprocess = aProcess ? *aProcess : hprocess_local; // To simplify other things.
	hprocess = NULL; // Init output param if the caller gave us memory to store it.  Even if caller didn't, other things below may rely on this being initialized.
	if (aOutputVar) // Same
		aOutputVar->Assign();

	// Launching nothing is always a success:
	if (!aAction || !*aAction) return OK;

	// Make sure this is set to NULL because CreateProcess() won't work if it's the empty string:
	if (aWorkingDir && !*aWorkingDir)
		aWorkingDir = NULL;

	#define IS_VERB(str) (   !_tcsicmp(str, _T("find")) || !_tcsicmp(str, _T("explore")) || !_tcsicmp(str, _T("open"))\
		|| !_tcsicmp(str, _T("edit")) || !_tcsicmp(str, _T("print")) || !_tcsicmp(str, _T("properties"))   )

	// Set default items to be run by ShellExecute().  These are also used by the error
	// reporting at the end, which is why they're initialized even if CreateProcess() works
	// and there's no need to use ShellExecute():
	LPTSTR shell_verb = NULL;
	LPTSTR shell_action = aAction;
	LPTSTR shell_params = NULL;
	
	///////////////////////////////////////////////////////////////////////////////////
	// This next section is done prior to CreateProcess() because when aParams is NULL,
	// we need to find out whether aAction contains a system verb.
	///////////////////////////////////////////////////////////////////////////////////
	if (aParams) // Caller specified the params (even an empty string counts, for this purpose).
	{
		if (IS_VERB(shell_action))
		{
			shell_verb = shell_action;
			shell_action = aParams;
		}
		else
			shell_params = aParams;
	}
	else // Caller wants us to try to parse params out of aAction.
	{
		// Find out the "first phrase" in the string to support the special "find" and "explore" operations.
		LPTSTR phrase;
		size_t phrase_len;
		// Set phrase_end to be the location of the first whitespace char, if one exists:
		LPTSTR phrase_end = StrChrAny(shell_action, _T(" \t")); // Find space or tab.
		if (phrase_end) // i.e. there is a second phrase.
		{
			phrase_len = phrase_end - shell_action;
			// Create a null-terminated copy of the phrase for comparison.
			phrase = tmemcpy(talloca(phrase_len + 1), shell_action, phrase_len);
			phrase[phrase_len] = '\0';
			// Firstly, treat anything following '*' as a verb, to support custom verbs like *Compile.
			if (*phrase == '*')
				shell_verb = phrase + 1;
			// Secondly, check for common system verbs like "find" and "edit".
			else if (IS_VERB(phrase))
				shell_verb = phrase;
			if (shell_verb)
				// Exclude the verb and its trailing space or tab from further consideration.
				shell_action += phrase_len + 1;
			// Otherwise it's not a verb, and may be re-parsed later.
		}
		// shell_action will be split into action and params at a later stage if ShellExecuteEx is to be used.
	}

	// This is distinct from hprocess being non-NULL because the two aren't always the
	// same.  For example, if the user does "Run, find D:\" or "RunWait, www.yahoo.com",
	// no new process handle will be available even though the launch was successful:
	bool success = false; // Separate from last_error for maintainability.
	DWORD last_error = 0;

	bool use_runas = aUseRunAs && (!mRunAsUser.IsEmpty() || !mRunAsPass.IsEmpty() || !mRunAsDomain.IsEmpty());
	if (use_runas && shell_verb)
	{
		if (aDisplayErrors)
			ScriptError(_T("System verbs unsupported with RunAs."));
		return FAIL;
	}
	
	size_t action_length = _tcslen(shell_action); // shell_action == aAction if shell_verb == NULL.
	if (action_length >= LINE_SIZE) // Max length supported by CreateProcess() is 32 KB. But there hasn't been any demand to go above 16 KB, so seems little need to support it (plus it reduces risk of stack overflow).
	{
        if (aDisplayErrors)
			ScriptError(_T("String too long.")); // Short msg since so rare.
		return FAIL;
	}

	// If the caller originally gave us NULL for aParams, always try CreateProcess() before
	// trying ShellExecute().  This is because ShellExecute() is usually a lot slower.
	// The only exception is if the action appears to be a verb such as open, edit, or find.
	// In that case, we'll also skip the CreateProcess() attempt and do only the ShellExecute().
	// If the user really meant to launch find.bat or find.exe, for example, he should add
	// the extension (e.g. .exe) to differentiate "find" from "find.exe":
	if (!shell_verb)
	{
		STARTUPINFO si = {0}; // Zero fill to be safer.
		si.cb = sizeof(si);
		// The following are left at the default of NULL/0 set higher above:
		//si.lpReserved = si.lpDesktop = si.lpTitle = NULL;
		//si.lpReserved2 = NULL;
		si.dwFlags = STARTF_USESHOWWINDOW;  // This tells it to use the value of wShowWindow below.
		si.wShowWindow = (aRunShowMode && *aRunShowMode) ? Line::ConvertRunMode(aRunShowMode) : SW_SHOWNORMAL;
		PROCESS_INFORMATION pi = {0};

		// Since CreateProcess() requires that the 2nd param be modifiable, ensure that it is
		// (even if this is ANSI and not Unicode; it's just safer):
		LPTSTR command_line;
		if (aParams && *aParams)
		{
			command_line = talloca(action_length + _tcslen(aParams) + 10); // +10 to allow room for space, terminator, and any extra chars that might get added in the future.
			_stprintf(command_line, _T("%s %s"), aAction, aParams);
		}
		else // We're running the original action from caller.
		{
			command_line = talloca(action_length + 1);
        	_tcscpy(command_line, aAction); // CreateProcessW() requires modifiable string.  Although non-W version is used now, it feels safer to make it modifiable anyway.
		}

		if (use_runas)
		{
			if (!DoRunAs(command_line, aWorkingDir, aDisplayErrors, si.wShowWindow  // wShowWindow (min/max/hide).
				, aOutputVar, pi, success, hprocess, last_error)) // These are output parameters it will set for us.
				return FAIL; // It already displayed the error, if appropriate.
		}
		else
		{
			// MSDN: "If [lpCurrentDirectory] is NULL, the new process is created with the same
			// current drive and directory as the calling process." (i.e. since caller may have
			// specified a NULL aWorkingDir).  Also, we pass NULL in for the first param so that
			// it will behave the following way (hopefully under all OSes): "the first white-space delimited
			// token of the command line specifies the module name. If you are using a long file name that
			// contains a space, use quoted strings to indicate where the file name ends and the arguments
			// begin (see the explanation for the lpApplicationName parameter). If the file name does not
			// contain an extension, .exe is appended. Therefore, if the file name extension is .com,
			// this parameter must include the .com extension. If the file name ends in a period (.) with
			// no extension, or if the file name contains a path, .exe is not appended. If the file name does
			// not contain a directory path, the system searches for the executable file in the following
			// sequence...".
			// Provide the app name (first param) if possible, for greater expected reliability.
			// UPDATE: Don't provide the module name because if it's enclosed in double quotes,
			// CreateProcess() will fail, at least under XP:
			//if (CreateProcess(aParams && *aParams ? aAction : NULL
			if (CreateProcess(NULL, command_line, NULL, NULL, FALSE, 0, NULL, aWorkingDir, &si, &pi))
			{
				success = true;
				if (pi.hThread)
					CloseHandle(pi.hThread); // Required to avoid memory leak.
				hprocess = pi.hProcess;
				if (aOutputVar)
					aOutputVar->Assign(pi.dwProcessId);
			}
			else
				last_error = GetLastError();
		}
	}

	// Since CreateProcessWithLogonW() was either not attempted or did not work, it's probably
	// best to display an error rather than trying to run it without the RunAs settings.
	// This policy encourages users to have RunAs in effect only when necessary:
	if (!success && !use_runas) // Either the above wasn't attempted, or the attempt failed.  So try ShellExecute().
	{
		SHELLEXECUTEINFO sei = {0};
		// sei.hwnd is left NULL to avoid potential side-effects with having a hidden window be the parent.
		// However, doing so may result in the launched app appearing on a different monitor than the
		// script's main window appears on (for multimonitor systems).  This seems fairly inconsequential
		// since scripted workarounds are possible.
		sei.cbSize = sizeof(sei);
		// Below: "indicate that the hProcess member receives the process handle" and not to display error dialog:
		sei.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI;
		sei.lpDirectory = aWorkingDir; // OK if NULL or blank; that will cause current dir to be used.
		sei.nShow = (aRunShowMode && *aRunShowMode) ? Line::ConvertRunMode(aRunShowMode) : SW_SHOWNORMAL;
		if (shell_verb)
		{
			sei.lpVerb = shell_verb;
			if (!_tcsicmp(shell_verb, _T("properties")))
				sei.fMask |= SEE_MASK_INVOKEIDLIST;  // Need to use this for the "properties" verb to work reliably.
		}
		if (!shell_params) // i.e. above hasn't determined the params yet.
		{
// Rather than just consider the first phrase to be the executable and the rest to be the param, we check it
// for a proper extension so that the user can launch a document name containing spaces, without having to
// enclose it in double quotes.  UPDATE: Want to be able to support executable filespecs without requiring them
// to be enclosed in double quotes.  Therefore, search the entire string, rather than just first_phrase, for
// the left-most occurrence of a valid executable extension.  This should be fine since the user can still
// pass in EXEs and such as params as long as the first executable is fully qualified with its real extension
// so that we can tell that it's the action and not one of the params.  UPDATE: Since any file type may
// potentially accept parameters (.lnk or .ahk files for instance), the first space-terminated substring which
// is either an existing file or ends in one of .exe,.bat,.com,.cmd,.hta is considered the executable and the
// rest is considered the param.  Remaining shortcomings of this method include:
//   -  It doesn't handle an extensionless executable such as "notepad test.txt"
//   -  It doesn't handle custom file types (scripts etc.) which don't exist in the working directory but can
//      still be executed due to %PATH% and %PATHEXT% even when our caller doesn't supply an absolute path.
// These limitations seem acceptable since the caller can allow even those cases to work by simply wrapping
// the action in quote marks.
			// Make a copy so that we can modify it (i.e. split it into action & params).
			// Using talloca ensures it will stick around until the function exits:
			LPTSTR parse_buf = talloca(action_length + 1);
			_tcscpy(parse_buf, shell_action);
			LPTSTR action_extension, action_end;
			// Let quotation marks be used to remove all ambiguity:
			if (*parse_buf == '"' && (action_end = _tcschr(parse_buf + 1, '"')))
			{
				shell_action = parse_buf + 1;
				*action_end = '\0';
				if (action_end[1])
				{
					shell_params = action_end + 1;
					// Omit the space which should follow, but only one, in case spaces
					// are meaningful to the target program.
					if (*shell_params == ' ')
						++shell_params;
				}
				// Otherwise, there's only the action in quotation marks and no params.
			}
			else
			{
				if (aWorkingDir) // Set current directory temporarily in case the action is a relative path:
					SetCurrentDirectory(aWorkingDir);
				// For each space which possibly delimits the action and params:
				for (action_end = parse_buf + 1; action_end = _tcschr(action_end, ' '); ++action_end)
				{
					// Find the beginning of the substring or file extension; if \ is encountered, this might be
					// an extensionless filename, but it probably wouldn't be meaningful to pass params to such a
					// file since it can't be associated with anything, so skip to the next space in that case.
					for ( action_extension = action_end - 1;
						  action_extension > parse_buf && !_tcschr(_T("\\/."), *action_extension);
						  --action_extension );
					if (*action_extension == '.') // Potential file extension; even if action_extension == parse_buf since ".ext" on its own is a valid filename.
					{
						*action_end = '\0'; // Temporarily terminate.
						// If action_extension is a common executable extension, don't call GetFileAttributes() since
						// the file might actually be in a location listed in %PATH% or the App Paths registry key:
						if ( (action_end-action_extension == 4 && tcscasestr(_T(".exe.bat.com.cmd.hta"), action_extension))
						// Otherwise the file might still be something capable of accepting params, like a script,
						// so check if what we have is the name of an existing file:
							|| !(GetFileAttributes(parse_buf) & FILE_ATTRIBUTE_DIRECTORY) ) // i.e. THE FILE EXISTS and is not a directory. This works because (INVALID_FILE_ATTRIBUTES & FILE_ATTRIBUTE_DIRECTORY) is non-zero.
						{	
							shell_action = parse_buf;
							shell_params = action_end + 1;
							break;
						}
						// What we have so far isn't an obvious executable file type or the path of an existing
						// file, so assume it isn't a valid action.  Unterminate and continue the loop:
						*action_end = ' ';
					}
				}
				if (aWorkingDir)
					SetCurrentDirectory(g_WorkingDir); // Restore to proper value.
			}
		}
		//else aParams!=NULL, so the extra parsing in the block above isn't necessary.

		// Not done because it may have been set to shell_verb above:
		//sei.lpVerb = NULL;
		sei.lpFile = shell_action;
		sei.lpParameters = shell_params; // NULL if no parameters were present.
		// Above was fixed v1.0.42.06 to be NULL rather than the empty string to prevent passing an
		// extra space at the end of a parameter list (this might happen only when launching a shortcut
		// [.lnk file]).  MSDN states: "If the lpFile member specifies a document file, lpParameters should
		// be NULL."  This implies that NULL is a suitable value for lpParameters in cases where you don't
		// want to pass any parameters at all.
		
		if (ShellExecuteEx(&sei)) // Relies on short-circuit boolean order.
		{
			typedef DWORD (WINAPI *GetProcessIDType)(HANDLE);
			// GetProcessID is only available on WinXP SP1 or later, so load it dynamically.
			static GetProcessIDType fnGetProcessID = (GetProcessIDType)GetProcAddress(GetModuleHandle(_T("kernel32.dll")), "GetProcessId");

			if (hprocess = sei.hProcess)
			{
				// A new process was created, so get its ID if possible.
				if (aOutputVar && fnGetProcessID)
					aOutputVar->Assign(fnGetProcessID(hprocess));
			}
			// Even if there's no process handle, it's considered a success because some
			// system verbs and file associations do not create a new process, by design.
			success = true;
		}
		else
			last_error = GetLastError();
	}

	if (!success) // The above attempt(s) to launch failed.
	{
		if (aUpdateLastError)
			g->LastError = last_error;

		if (aDisplayErrors)
		{
			TCHAR error_text[2048], verb_text[128], system_error_text[512];
			GetWin32ErrorText(system_error_text, _countof(system_error_text), last_error);
			if (shell_verb)
				sntprintf(verb_text, _countof(verb_text), _T("\nVerb: <%s>"), shell_verb);
			else // Don't bother showing it if it's just "open".
				*verb_text = '\0';
			if (!shell_params)
				shell_params = _T(""); // Expected to be non-NULL below.
			// Use format specifier to make sure it doesn't get too big for the error
			// function to display:
			sntprintf(error_text, _countof(error_text)
				, _T("%s\nAction: <%-0.400s%s>")
				_T("%s")
				_T("\nParams: <%-0.400s%s>")
				, use_runas ? _T("Launch Error (possibly related to RunAs):") : _T("Failed attempt to launch program or document:")
				, shell_action, _tcslen(shell_action) > 400 ? _T("...") : _T("")
				, verb_text
				, shell_params, _tcslen(shell_params) > 400 ? _T("...") : _T("")
				);
			ScriptError(error_text, system_error_text);
		}
		return FAIL;
	}

	// Otherwise, success:
	if (aUpdateLastError)
		g->LastError = 0; // Force zero to indicate success, which seems more maintainable and reliable than calling GetLastError() right here.

	// If aProcess isn't NULL, the caller wanted the process handle left open and so it must eventually call
	// CloseHandle().  Otherwise, we should close the process if it's non-NULL (it can be NULL in the case of
	// launching things like "find D:\" or "www.yahoo.com").
	if (!aProcess && hprocess)
		CloseHandle(hprocess); // Required to avoid memory leak.
	return OK;
}
