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
#include "window.h" // for a lot of things
#include "application.h" // for MsgSleep()
#include "TextIO.h"

#define NA MAX_FUNCTION_PARAMS
#define BIFn(name, minp, maxp, bif, ...) {_T(#name), bif, minp, maxp, FID_##name, __VA_ARGS__}
#define BIFi(name, minp, maxp, bif, id, ...) {_T(#name), bif, minp, maxp, id, __VA_ARGS__}
#define BIF1(name, minp, maxp, ...) {_T(#name), BIF_##name, minp, maxp, 0, __VA_ARGS__}
// The following array defines all built-in functions, their min and max parameter counts,
// and which real C++ function contains their implementation.  The macros above are used to
// reduce repetition and implement code-sharing: BIFs which share a C++ function are assigned
// an integer ID, defined in the BuiltInFunctionID enum.  Items must be in alphabetical order.
FuncEntry g_BIF[] =
{
	BIF1(Abs, 1, 1),
	BIFn(ACos, 1, 1, BIF_ASinACos),
	BIFn(ASin, 1, 1, BIF_ASinACos),
	BIF1(ATan, 1, 1),
	BIF1(CaretGetPos, 0, 2, {1, 2}),
	BIFn(Ceil, 1, 1, BIF_FloorCeil),
	BIF1(Chr, 1, 1),
	BIF1(Click, 0, 6),
#ifdef ENABLE_DLLCALL
	BIFn(ComCall, 2, NA, BIF_DllCall),
#endif
	BIF1(ComObjActive, 1, 1),
	BIF1(ComObjConnect, 1, 2),
	BIF1(ComObjFlags, 1, 3),
	BIFi(ComObjFromPtr, 1, 1, BIF_ComObj, 0),
	BIF1(ComObjGet, 1, 1),
	BIF1(ComObjQuery, 2, 3),
	BIF1(ComObjType, 1, 2),
	BIF1(ComObjValue, 1, 1),
	BIF1(Cos, 1, 1),
#ifdef ENABLE_DLLCALL
	BIFn(DllCall, 1, NA, BIF_DllCall),
#endif
	BIF1(Exp, 1, 1),
	BIF1(FileOpen, 2, 3),
	BIFn(Floor, 1, 1, BIF_FloorCeil),
	BIF1(Format, 1, NA),
	BIF1(FormatTime, 0, 2),
	BIFn(GetMethod, 1, 3, BIF_GetMethod),
	BIF1(HasBase, 2, 2),
	BIFn(HasMethod, 1, 3, BIF_GetMethod),
	BIF1(HasProp, 2, 2),
	BIF1(InStr, 2, 5),
	BIFi(IsAlnum, 1, 2, BIF_IsTypeish, VAR_TYPE_ALNUM),
	BIFi(IsAlpha, 1, 2, BIF_IsTypeish, VAR_TYPE_ALPHA),
	BIFi(IsDigit, 1, 1, BIF_IsTypeish, VAR_TYPE_DIGIT),
	BIFi(IsFloat, 1, 1, BIF_IsTypeish, VAR_TYPE_FLOAT),
	BIFi(IsInteger, 1, 1, BIF_IsTypeish, VAR_TYPE_INTEGER),
	BIFi(IsLower, 1, 2, BIF_IsTypeish, VAR_TYPE_LOWER),
	BIFi(IsNumber, 1, 1, BIF_IsTypeish, VAR_TYPE_NUMBER),
	BIF1(IsObject, 1, 1),
	BIFi(IsSetRef, 1, 1, BIF_IsSet, 0, {1}),
	BIFi(IsSpace, 1, 1, BIF_IsTypeish, VAR_TYPE_SPACE),
	BIFi(IsTime, 1, 1, BIF_IsTypeish, VAR_TYPE_TIME),
	BIFi(IsUpper, 1, 2, BIF_IsTypeish, VAR_TYPE_UPPER),
	BIFi(IsXDigit, 1, 1, BIF_IsTypeish, VAR_TYPE_XDIGIT),
	BIFn(Ln, 1, 1, BIF_SqrtLogLn),
	BIFn(Log, 1, 1, BIF_SqrtLogLn),
	BIFn(LTrim, 1, 2, BIF_Trim),
	BIFn(Max, 1, NA, BIF_MinMax),
	BIFn(Min, 1, NA, BIF_MinMax),
	BIF1(Mod, 2, 2),
	BIF1(NumGet, 2, 3),
	BIF1(NumPut, 3, NA),
	BIFn(ObjAddRef, 1, 1, BIF_ObjAddRefRelease),
	BIF1(ObjBindMethod, 1, NA),
	BIFn(ObjFromPtr, 1, 1, BIF_ObjPtr),
	BIFn(ObjFromPtrAddRef, 1, 1, BIF_ObjPtr),
	BIFn(ObjGetBase, 1, 1, BIF_Base),
	BIFn(ObjGetCapacity, 1, 1, BIF_ObjXXX),
	BIFn(ObjHasOwnProp, 2, 2, BIF_ObjXXX),
	BIFn(ObjOwnPropCount, 1, 1, BIF_ObjXXX),
	BIFn(ObjOwnProps, 1, 1, BIF_ObjXXX),
	BIFn(ObjPtr, 1, 1, BIF_ObjPtr),
	BIFn(ObjPtrAddRef, 1, 1, BIF_ObjPtr),
	BIFn(ObjRelease, 1, 1, BIF_ObjAddRefRelease),
	BIFn(ObjSetBase, 2, 2, BIF_Base),
	BIFn(ObjSetCapacity, 2, 2, BIF_ObjXXX),
	BIF1(Ord, 1, 1),
	BIF1(Random, 0, 2),
	BIFn(RegCreateKey, 0, 1, BIF_Reg),
	BIFn(RegDelete, 0, 2, BIF_Reg),
	BIFn(RegDeleteKey, 0, 1, BIF_Reg),
	BIFn(RegExMatch, 2, 4, BIF_RegEx, {3}),
	BIFn(RegExReplace, 2, 6, BIF_RegEx, {4}),
	BIFn(RegRead, 0, 3, BIF_Reg),
	BIFn(RegWrite, 0, 4, BIF_Reg),
	BIF1(Round, 1, 2),
	BIFn(RTrim, 1, 2, BIF_Trim),
	BIF1(RunWait, 1, 4, {4}),
	BIF1(Sin, 1, 1),
	BIF1(Sort, 1, 3),
	BIFn(SoundGetInterface, 1, 3, BIF_Sound),
	BIFn(SoundGetMute, 0, 2, BIF_Sound),
	BIFn(SoundGetName, 0, 2, BIF_Sound),
	BIFn(SoundGetVolume, 0, 2, BIF_Sound),
	BIFn(SoundSetMute, 1, 3, BIF_Sound),
	BIFn(SoundSetVolume, 1, 3, BIF_Sound),
	BIF1(SplitPath, 1, 6, {2, 3, 4, 5, 6}),
	BIFn(Sqrt, 1, 1, BIF_SqrtLogLn),
	BIF1(StrCompare, 2, 3),
	BIFn(StrGet, 1, 3, BIF_StrGetPut),
	BIF1(StrLen, 1, 1),
	BIFn(StrLower, 1, 1, BIF_StrCase),
	BIF1(StrPtr, 1, 1),
	BIFn(StrPut, 1, 4, BIF_StrGetPut),
	BIF1(StrReplace, 2, 6, {5}),
	BIFn(StrTitle, 1, 1, BIF_StrCase),
	BIFn(StrUpper, 1, 1, BIF_StrCase),
	BIF1(SubStr, 2, 3),
	BIF1(Tan, 1, 1),
	BIFn(Trim, 1, 2, BIF_Trim),
	BIF1(Type, 1, 1),
	BIF1(VarSetStrCapacity, 1, 2, {1}),
	BIF1(VerCompare, 2, 2),
	BIFn(WinActive, 0, 4, BIF_WinExistActive),
	BIFn(WinExist, 0, 4, BIF_WinExistActive),
};
#undef NA
#undef BIFn
#undef BIF1


#define A_x(name, fn) { _T(#name), fn, NULL }
#define A_(name) A_x(name, BIV_##name)
#define A_wx(name, fnget, fnset) { _T(#name), fnget, fnset }
#define A_w(name) A_wx(name, BIV_##name, BIV_##name##_Set)
// IMPORTANT: The following array must be kept in alphabetical order
// for binary search to work.  See Script::GetBuiltInVar for further comments.
// g_BIV_A: All built-in vars beginning with "A_".  The prefix is omitted from each
// name to reduce code size and speed up the comparisons.
VarEntry g_BIV_A[] =
{
	A_(AhkPath),
	A_(AhkVersion),
	A_w(AllowMainWindow),
	A_x(AppData, BIV_SpecialFolderPath),
	A_x(AppDataCommon, BIV_SpecialFolderPath),
	A_w(Clipboard),
	A_x(ComputerName, BIV_UserName_ComputerName),
	A_(ComSpec),
	A_wx(ControlDelay, BIV_xDelay, BIV_xDelay_Set),
	A_wx(CoordModeCaret, BIV_CoordMode, BIV_CoordMode_Set),
	A_wx(CoordModeMenu, BIV_CoordMode, BIV_CoordMode_Set),
	A_wx(CoordModeMouse, BIV_CoordMode, BIV_CoordMode_Set),
	A_wx(CoordModePixel, BIV_CoordMode, BIV_CoordMode_Set),
	A_wx(CoordModeToolTip, BIV_CoordMode, BIV_CoordMode_Set),
	A_(Cursor),
	A_x(DD, BIV_DateTime),
	A_x(DDD, BIV_MMM_DDD),
	A_x(DDDD, BIV_MMM_DDD),
	A_w(DefaultMouseSpeed),
	A_x(Desktop, BIV_SpecialFolderPath),
	A_x(DesktopCommon, BIV_SpecialFolderPath),
	A_w(DetectHiddenText),
	A_w(DetectHiddenWindows),
	A_(EndChar),
	A_w(EventInfo),
	A_w(FileEncoding),
	A_wx(HotkeyInterval, BIV_Hotkey, BIV_Hotkey_Set),
	A_wx(HotkeyModifierTimeout, BIV_Hotkey, BIV_Hotkey_Set),
	A_x(Hour, BIV_DateTime),
	A_(IconFile),
	A_w(IconHidden),
	A_(IconNumber),
	A_w(IconTip),
	A_wx(Index, BIV_LoopIndex, BIV_LoopIndex_Set),
	A_(InitialWorkingDir),
	A_(Is64bitOS),
	A_(IsAdmin),
	A_(IsCompiled),
	A_(IsCritical),
	A_(IsPaused),
	A_(IsSuspended),
	A_wx(KeyDelay, BIV_xDelay, BIV_xDelay_Set),
	A_wx(KeyDelayPlay, BIV_xDelay, BIV_xDelay_Set),
	A_wx(KeyDuration, BIV_xDelay, BIV_xDelay_Set),
	A_wx(KeyDurationPlay, BIV_xDelay, BIV_xDelay_Set),
	A_(Language),
	A_w(LastError),
	A_(LineFile),
	A_(LineNumber),
	A_w(ListLines),
	A_(LoopField),
	A_(LoopFileAttrib),
	A_(LoopFileDir),
	A_(LoopFileExt),
	A_(LoopFileFullPath),
	A_x(LoopFileName, BIV_LoopFileName),
	A_(LoopFilePath),
	A_x(LoopFileShortName, BIV_LoopFileName),
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
	A_(LoopRegTimeModified),
	A_(LoopRegType),
	A_wx(MaxHotkeysPerInterval, BIV_Hotkey, BIV_Hotkey_Set),
	A_x(MDay, BIV_DateTime),
	A_w(MenuMaskKey),
	A_x(Min, BIV_DateTime),
	A_x(MM, BIV_DateTime),
	A_x(MMM, BIV_MMM_DDD),
	A_x(MMMM, BIV_MMM_DDD),
	A_x(Mon, BIV_DateTime),
	A_wx(MouseDelay, BIV_xDelay, BIV_xDelay_Set),
	A_wx(MouseDelayPlay, BIV_xDelay, BIV_xDelay_Set),
	A_x(MSec, BIV_DateTime),
	A_(MyDocuments),
	A_x(Now, BIV_Now),
	A_x(NowUTC, BIV_Now),
	A_(OSVersion),
	A_(PriorHotkey),
	A_(PriorKey),
	A_x(ProgramFiles, BIV_SpecialFolderPath),
	A_x(Programs, BIV_SpecialFolderPath),
	A_x(ProgramsCommon, BIV_SpecialFolderPath),
	A_(PtrSize),
	A_w(RegView),
	A_(ScreenDPI),
	A_x(ScreenHeight, BIV_ScreenWidth_Height),
	A_x(ScreenWidth, BIV_ScreenWidth_Height),
	A_(ScriptDir),
	A_(ScriptFullPath),
	A_(ScriptHwnd),
	A_w(ScriptName),
	A_x(Sec, BIV_DateTime),
	A_w(SendLevel),
	A_w(SendMode),
	A_x(Space, BIV_Space_Tab),
	A_x(StartMenu, BIV_SpecialFolderPath),
	A_x(StartMenuCommon, BIV_SpecialFolderPath),
	A_x(Startup, BIV_SpecialFolderPath),
	A_x(StartupCommon, BIV_SpecialFolderPath),
	A_w(StoreCapsLockMode),
	A_x(Tab, BIV_Space_Tab),
	A_(Temp), // Debatably should be A_TempDir, but brevity seemed more popular with users, perhaps for heavy uses of the temp folder.,
	A_(ThisFunc),
	A_(ThisHotkey),
	A_(TickCount),
	A_x(TimeIdle, BIV_TimeIdle),
	A_x(TimeIdleKeyboard, BIV_TimeIdle),
	A_x(TimeIdleMouse, BIV_TimeIdle),
	A_x(TimeIdlePhysical, BIV_TimeIdle),
	A_(TimeSincePriorHotkey),
	A_(TimeSinceThisHotkey),
	A_w(TitleMatchMode),
	A_wx(TitleMatchModeSpeed, BIV_TitleMatchModeSpeed, BIV_TitleMatchMode_Set),
	A_(TrayMenu),
	A_x(UserName, BIV_UserName_ComputerName),
	A_x(WDay, BIV_DateTime),
	A_wx(WinDelay, BIV_xDelay, BIV_xDelay_Set),
	A_(WinDir),
	A_w(WorkingDir),
	A_x(YDay, BIV_DateTime),
	A_x(Year, BIV_DateTime),
	A_x(YWeek, BIV_DateTime),
	A_x(YYYY, BIV_DateTime)
};
#undef A_
#undef A_x
#undef A_w
#undef A_wx
#undef VF


// General note about the methods in here:
// Want to be able to support multiple simultaneous points of execution
// because more than one subroutine can be executing simultaneously
// (well, more precisely, there can be more than one script subroutine
// that's in a "currently running" state, even though all such subroutines,
// except for the most recent one, are suspended.  So keep this in mind when
// using things such as static data members or static local variables.


Script::Script()
	: mFirstLine(NULL), mLastLine(NULL), mCurrLine(NULL)
	, mThisHotkeyName(_T("")), mPriorHotkeyName(_T("")), mThisHotkeyStartTime(0), mPriorHotkeyStartTime(0)
	, mEndChar(0), mThisHotkeyModifiersLR(0)
	, mOnClipboardChangeIsRunning(false)
	, mFirstLabel(NULL), mLastLabel(NULL)
	, mLastHotFunc(nullptr), mUnusedHotFunc(nullptr)
	, mFirstTimer(NULL), mLastTimer(NULL), mTimerEnabledCount(0), mTimerCount(0)
	, mFirstMenu(NULL), mLastMenu(NULL), mMenuCount(0)
	, mOpenBlock(NULL), mNextLineIsFunctionBody(false), mNoUpdateLabels(false)
	, mClassObjectCount(0), mUnresolvedClasses(NULL), mClassProperty(NULL), mClassPropertyDef(NULL)
	, mCurrFileIndex(0), mCombinedLineNumber(0)
	, mFileSpec(_T("")), mFileDir(_T("")), mFileName(_T("")), mOurEXE(_T("")), mOurEXEDir(_T("")), mMainWindowTitle(_T(""))
	, mScriptName(NULL)
	, mIsReadyToExecute(false), mAutoExecSectionIsRunning(false)
	, mIsRestart(false), mErrorStdOut(false), mErrorStdOutCP(0)
#ifndef AUTOHOTKEYSC
	, mValidateThenExit(false)
	, mCmdLineInclude(NULL)
#endif
	, mUninterruptedLineCountMax(1000), mUninterruptibleTime(17)
	, mCustomIcon(NULL), mCustomIconSmall(NULL) // Normally NULL unless there's a custom tray icon loaded dynamically.
	, mCustomIconFile(NULL), mIconFrozen(false), mTrayIconTip(NULL) // Allocated on first use.
	, mCustomIconNumber(0)
{
	// v1.0.25: mLastScriptRest (removed in v2) and mLastPeekTime are now initialized
	// right before the auto-exec section of the script is launched, which avoids an
	// initial Sleep(10) in ExecUntil that would otherwise occur.
	ZeroMemory(&mNIC, sizeof(mNIC));  // Constructor initializes this, to be safe.
	mNIC.hWnd = NULL;  // Set this as an indicator that it tray icon is not installed.
	
	// This is done here rather than in Init() because by then CreateRootPrototypes()
	// definitely would have already caused functions to be added to the list.
	mFuncs.Alloc(100); // Avoid multiple reallocations for simple scripts.

#ifdef _DEBUG
	if (ID_FILE_EXIT < ID_MAIN_FIRST) // Not a very thorough check.
		ScriptError(_T("DEBUG: ID_FILE_EXIT is too large (conflicts with IDs reserved via ID_USER_FIRST)."));
	if (MAX_CONTROLS_PER_GUI > ID_USER_FIRST - 3)
		ScriptError(_T("DEBUG: MAX_CONTROLS_PER_GUI is too large (conflicts with IDs reserved via ID_USER_FIRST)."));
	if (g_ActionCount != ACT_LAST_NAMED_ACTION + 1)
		ScriptError(_T("DEBUG: g_act and enum_act are out of sync."));
	int LargestMaxParams, i;
	// Find the Largest value of MaxParams used by any command and make sure it
	// isn't something larger than expected by the parsing routines:
	for (LargestMaxParams = i = 0; i < g_ActionCount; ++i)
		if (g_act[i].MaxParams > LargestMaxParams)
			LargestMaxParams = g_act[i].MaxParams;
	if (LargestMaxParams > MAX_ARGS)
		ScriptError(_T("DEBUG: At least one command supports more arguments than allowed."));
	if (sizeof(ActionTypeType) == 1 && g_ActionCount > 256)
		ScriptError(_T("DEBUG: Since there are now more than 256 Action Types, the ActionTypeType")
			_T(" typedef must be changed."));
	// Ensure binary-search arrays are sorted correctly:
	for (int i = 1; i < _countof(g_BIF); ++i)
		if (_tcsicmp(g_BIF[i-1].mName, g_BIF[i].mName) >= 0)
			ScriptError(_T("DEBUG: g_BIF out of order."), g_BIF[i].mName);
	for (int i = 1; i < _countof(g_BIV_A); ++i)
		if (_tcsicmp(g_BIV_A[i-1].name, g_BIV_A[i].name) >= 0)
			ScriptError(_T("DEBUG: g_BIV_A out of order."), g_BIV_A[i].name);
#endif
	InitializeCriticalSection(&g_CriticalRegExCache); // v1.0.45.04: Must be done early so that it's unconditional, so that DeleteCriticalSection() in the script destructor can also be unconditional.
	OleInitialize(NULL);
}



Script::~Script() // Destructor.
{
	Hotkey::AllDestruct(); // Unregister hooks and hotkeys.

	if (mNIC.hWnd) // Tray icon is installed.
		Shell_NotifyIcon(NIM_DELETE, &mNIC); // Remove it.

	if (mOnClipboardChange.Count()) // Remove from viewer chain.
		EnableClipboardListener(false);

	// DestroyWindow() will cause MainWindowProc() to immediately receive and process the
	// WM_DESTROY msg, which should in turn result in any child windows being destroyed
	// and other cleanup being done:
	g_DestroyWindowCalled = true;
	DestroyWindow(g_hWnd);


	int i;
	// It is safer/easier to destroy the GUI windows prior to the menus (especially the menu bars).
	// This is because one GUI window might get destroyed and take with it a menu bar that is still
	// in use by an existing GUI window.  GuiType::Destroy() adheres to this philosophy by detaching
	// its menu bar prior to destroying its window.
	GuiType* gui;
	while (gui = g_firstGui) // Destroy any remaining GUI windows (due to e.g. circular references). Also: assignment.
		gui->Destroy();
	for (i = 0; i < GuiType::sFontCount; ++i) // Now that GUI windows are gone, delete all GUI fonts.
		if (GuiType::sFont[i].hfont)
			DeleteObject(GuiType::sFont[i].hfont);
	// The above might attempt to delete an HFONT from GetStockObject(DEFAULT_GUI_FONT), etc.
	// But that should be harmless:
	// MSDN: "It is not necessary (but it is not harmful) to delete stock objects by calling DeleteObject."

	// Above: Probably best to have removed icon from tray and destroyed any Gui windows that were
	// using it prior to getting rid of the script's custom icon below:
	if (mCustomIcon)
	{
		DestroyIcon(mCustomIcon);
		DestroyIcon(mCustomIconSmall); // Should always be non-NULL if mCustomIcon is non-NULL.
	}

	// Since they're not associated with a window, we must free the resources for all popup menus.
	// Update: Even if a menu is being used as a GUI window's menu bar, see note above for why menu
	// destruction is done AFTER the GUI windows are destroyed:
	for (UserMenu *m = mFirstMenu; m; m = m->mNextMenu)
		m->Dispose();

	// Since tooltip windows are unowned, they should be destroyed to avoid resource leak:
	for (i = 0; i < MAX_TOOLTIPS; ++i)
		if (g_hWndToolTip[i] && IsWindow(g_hWndToolTip[i]))
			DestroyWindow(g_hWndToolTip[i]);

	// Close any open sound item to prevent hang-on-exit in certain operating systems or conditions.
	// If there's any chance that a sound was played and not closed out, or that it is still playing,
	// this check is done.  Otherwise, the check is avoided since it might be a high overhead call,
	// especially if the sound subsystem part of the OS is currently swapped out or something:
	if (g_SoundWasPlayed)
	{
		TCHAR buf[MAX_PATH * 2]; // See "MAX_PATH note" in Line::SoundPlay for comments.
		mciSendString(_T("status ") SOUNDPLAY_ALIAS _T(" mode"), buf, _countof(buf), NULL);
		if (*buf) // "playing" or "stopped"
			mciSendString(_T("close ") SOUNDPLAY_ALIAS, NULL, 0, NULL);
	}

#ifdef CONFIG_DLL
	UnregisterClass(WINDOW_CLASS_MAIN, g_hInstance);
	UnregisterClass(WINDOW_CLASS_GUI, g_hInstance);
#endif

	DeleteCriticalSection(&g_CriticalRegExCache); // g_CriticalRegExCache is used elsewhere for thread-safety.
	OleUninitialize();
}



ResultType Script::Init(LPTSTR aScriptFilename)
// Returns OK or FAIL.
// Caller has provided an empty string for aScriptFilename if this is a compiled script.
// Otherwise, aScriptFilename can be NULL if caller hasn't determined the filename of the script yet.
{
	TCHAR buf[UorA(T_MAX_PATH, 2048)]; // Just to make sure we have plenty of room to do things with.
	size_t buf_length;

	// It may be better to get the module name this way rather than reading it from the registry
	// in case the user has moved it to a folder other than the install folder, hasn't installed it,
	// or has renamed the EXE file itself.
	if (buf_length = GetModuleFileName(NULL, buf, _countof(buf)))
	{
		// Any path longer than MAX_PATH is probably impossible as of 2018 since testing indicates
		// the program can't start if its path is longer than MAX_PATH-1 even with Windows 10 long
		// path awareness enabled.
		if (buf_length == _countof(buf)) // It was truncated.
			return FAIL; // Seems the safest option for this unlikely case.
		mOurEXE = SimpleHeap::Alloc(buf, buf_length);
		LPTSTR last_backslash = _tcsrchr(buf, '\\');
		if (last_backslash) // Probably always true due to the nature of GetModuleFileName().
			mOurEXEDir = SimpleHeap::Alloc(buf, last_backslash - buf);
	}

#ifdef AUTOHOTKEYSC
	mKind = ScriptKindResource;
	// Fix for v1.0.29: Override the caller's use of __argv[0] by using GetModuleFileName(),
	// so that when the script is started from the command line but the user didn't type the
	// extension, the extension will be included.  This necessary because otherwise
	// #SingleInstance wouldn't be able to detect duplicate versions in every case.
	// It also provides more consistency.
	aScriptFilename = buf;
#else
#ifdef _DEBUG
	if (!aScriptFilename)
		aScriptFilename = _T("Test\\Test.ahk");
#endif
	if (!aScriptFilename)
	{
		// Since no script-file was specified on the command line, use the default path,
		// <EXEDIR>\<EXENAME>.ahk.  This is intended for portable copies of AutoHotkey.
		// Users of a proper AutoHotkey installation should open .ahk files directly.
		LPTSTR suffix, dot;
		if (  (suffix = _tcsrchr(buf, '\\')) // Find name part of path.
			&& (dot = _tcsrchr(suffix, '.')) // Find extension part of name.
			&& dot - buf + 5 < _countof(buf)  ) // Enough space in buffer?
		{
			_tcscpy(dot, EXT_AUTOHOTKEY);
		}
		else // Very unlikely.
			return FAIL;

		aScriptFilename = buf; // Use the entire path, including the exe's directory.
	}
	if (*aScriptFilename == '*')
	{
		if (!aScriptFilename[1]) // Read script from stdin.
		{
			mKind = ScriptKindStdIn;
			buf[0] = '*';
			buf[1] = '\0';
			// Seems best to disable #SingleInstance for stdin scripts.
			g_AllowOnlyOneInstance = SINGLE_INSTANCE_OFF;
		}
		else
		{
			mKind = ScriptKindResource;
			g_AllowMainWindow = false;
		}
	}
	else
	{
		mKind = ScriptKindFile;
		if (aScriptFilename != buf)
		{
			// In case the script is a relative filespec (relative to current working dir):
			buf_length = GetFullPathName(aScriptFilename, _countof(buf), buf, NULL); // Succeeds even on nonexistent files.
			if (!buf_length || buf_length >= _countof(buf))
				return FAIL; // Due to rarity, no error msg, just abort.
		}
	}
#endif
	if (mKind != ScriptKindStdIn)
	{
		// Using the correct case not only makes it look better in title bar & tray tool tip,
		// it also helps with the detection of "this script already running" since otherwise
		// it might not find the dupe if the same script name is launched with different
		// lowercase/uppercase letters:
		ConvertFilespecToCorrectCase(buf, _countof(buf), buf_length); // This might change the length, e.g. due to expansion of 8.3 filename.
	}
	mFileSpec = SimpleHeap::Alloc(buf);  // The full spec is stored for convenience.
	LPTSTR filename_marker;
	if (filename_marker = _tcsrchr(buf, '\\'))
	{
		mFileDir = SimpleHeap::Alloc(buf, filename_marker - buf);
		++filename_marker;
	}
	else
	{
		// There is no slash in "*" (i.e. g_RunStdIn).  The only other known cause of this
		// condition is a path being too long for GetFullPathName to expand it into buf,
		// in which case buf and mFileSpec are now empty, and this will cause LoadFromFile()
		// to fail and the program to exit.
		mFileDir = g_WorkingDirOrig;
		filename_marker = buf;
	}
	mFileName = SimpleHeap::Alloc(filename_marker);
#ifdef AUTOHOTKEYSC
	// Omit AutoHotkey from the window title, like AutoIt3 does for its compiled scripts.
	// One reason for this is to reduce backlash if evil-doers create viruses and such
	// with the program.  buf already contains the full path, so no change is needed.
#else
	if (_tcscmp(aScriptFilename, SCRIPT_RESOURCE_SPEC))
		sntprintfcat(buf, _countof(buf), _T(" - %s"), mKind != ScriptKindResource ? T_AHK_NAME_VERSION : aScriptFilename);
#endif
	mMainWindowTitle = SimpleHeap::Alloc(buf);


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

	TCHAR class_name[64];
	HWND fore_win = GetForegroundWindow();
	bool do_minimize = !fore_win || (GetClassName(fore_win, class_name, _countof(class_name))
		&& !_tcsicmp(class_name, _T("Shell_TrayWnd"))); // Shell_TrayWnd is the taskbar's class on Win98/XP and probably the others too.

	// v2: Set the default size to 75% of the primary work area (which seems to be what
	// CW_USEDEFAULT does, at least on Windows 11); but limit the default width to the
	// total height, so the window won't be ridiculously wide on ultra-wide screens.
	RECT work;
	SystemParametersInfo(SPI_GETWORKAREA, 0, &work, 0);
	int default_width = (work.right - work.left) * 3 / 4;
	int default_height = work.bottom - work.top;
	if (default_width > default_height)
		default_width = default_height;
	default_height = default_height * 3 / 4;

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
		, default_width // width
		, default_height // height
		, NULL // parent window
		, NULL // Identifies a menu, or specifies a child-window identifier depending on the window style
		, g_hInstance // passed into WinMain
		, NULL))   ) // lpParam
	{
		MsgBox(_T("CreateWindow")); // Short msg since so rare.
		return FAIL;
	}

	// Now that all static initializers (such as for Object::sPrototype)
	// are guaranteed to have been executed, construct the Tray menu.
	mTrayMenu = new UserMenu(MENU_TYPE_POPUP);
	mTrayMenu->AppendStandardItems();

	// Stdin scripts leave the menu items enabled, to make it easier to imitate a standard script
	// (such as by handling the appropriate WM_COMMAND messages and overriding Reload/Edit).
	if (mKind == ScriptKindResource)
	{
		HMENU menu = GetMenu(g_hWnd);
		// Disable the Edit menu item, since it's not useful without a source file:
		EnableMenuItem(menu, ID_FILE_EDITSCRIPT, MF_DISABLED | MF_GRAYED);
		if (!g_AllowMainWindow) // Apply the compiled-script default value.
			EnableOrDisableViewMenuItems(menu, MF_DISABLED | MF_GRAYED);
	}

	if (    !(g_hWndEdit = CreateWindow(_T("Edit"), NULL, WS_CHILD | WS_VISIBLE | WS_BORDER
		| ES_LEFT | ES_MULTILINE | ES_READONLY | WS_VSCROLL // | WS_HSCROLL (saves space)
		, 0, 0, 0, 0, g_hWnd, (HMENU)1, g_hInstance, NULL))   )
	{
		MsgBox(_T("CreateWindow")); // Short msg since so rare.
		return FAIL;
	}
	// FONTS: The font used by default, at least on XP, is GetStockObject(SYSTEM_FONT).
	// Use something more appealing (monospaced seems preferable):
	HDC hdc = GetDC(g_hWndEdit);
	g_hFontEdit = CreateFont(FONT_POINT(hdc, 10), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
		, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, _T("Consolas"));
	ReleaseDC(g_hWndEdit, hdc);
	SendMessage(g_hWndEdit, WM_SETFONT, (WPARAM)g_hFontEdit, 0);

	// v1.0.30.05: Specifying a limit of zero opens the control to its maximum text capacity,
	// which removes the 32K size restriction.  Testing shows that this does not increase the actual
	// amount of memory used for controls containing small amounts of text.  All it does is allow
	// the control to allocate more memory as needed.
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

	g_hAccelTable = LoadAccelerators(g_hInstance, MAKEINTRESOURCE(IDR_ACCELERATOR1));

	if (g_NoTrayIcon)
		mNIC.hWnd = NULL;  // Set this as an indicator that tray icon is not installed.
	else
		// Even if the below fails, don't return FAIL in case the user is using a different shell
		// or something.  In other words, it is expected to fail under certain circumstances and
		// we want to tolerate that:
		CreateTrayIcon();

	return OK;
}



void Script::EnableClipboardListener(bool aEnable)
{
	static bool sEnabled = false;
	if (aEnable == sEnabled) // Simplifies BIF_On.
		return;
	if (aEnable)
		AddClipboardFormatListener(g_hWnd);
	else
		RemoveClipboardFormatListener(g_hWnd);
	sEnabled = aEnable;
}



void Script::AllowMainWindow(bool aAllow)
{
	if (g_AllowMainWindow == aAllow)
		return;
	g_AllowMainWindow = aAllow;
	EnableOrDisableViewMenuItems(GetMenu(g_hWnd)
		, aAllow ? MF_ENABLED : MF_DISABLED | MF_GRAYED);
	mTrayMenu->EnableStandardOpenItem(aAllow);
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
		// For simplicity, g_IconSmall isn't used since it might not match the current tray icon size.
		//: (icon == IDI_MAIN) ? g_IconSmall // Use the pre-loaded small icon for best quality.
		: NULL;
	if (!mNIC.hIcon)
	{
		// Retrieve the tray icon size each time in case the primary screen's DPI has changed.
		int cx = GetSystemTrayIconSize();
		// Use LR_SHARED for simplicity and performance more than to conserve memory in this case.
		// The documentation says "This function finds the first image in the cache with the requested
		// resource name, regardless of the size requested." -- but testing (on Windows 11) showed that
		// in reality, it caches one icon handle per size.
		mNIC.hIcon = (HICON)LoadImage(g_hInstance, MAKEINTRESOURCE(icon), IMAGE_ICON, cx, cx, LR_SHARED); 
	}
	// NIM_ADD is attempted as a fallback, since NIM_MODIFY can fail if the taskbar was
	// recreated or the icon has been deleted by a script calling Shell_NotifyIcon().
	if (Shell_NotifyIcon(NIM_MODIFY, &mNIC) || Shell_NotifyIcon(NIM_ADD, &mNIC))
	{
		icon_shows_paused = g->IsPaused;
		icon_shows_suspended = g_IsSuspended;
	}
	// else do nothing, just leave it in the same state.
}



bif_impl FResult TraySetIcon(optl<StrArg> aIconFile, optl<int> aIconNumber, optl<BOOL> aFreeze)
{
	return g_script.SetTrayIcon(aIconFile.value_or_null(), aIconNumber.value_or(1)
		, aFreeze.has_value() ? aFreeze.value() ? TOGGLED_ON : TOGGLED_OFF : NEUTRAL) ? OK : FR_FAIL;
}

ResultType Script::SetTrayIcon(LPCTSTR aIconFile, int aIconNumber, ToggleValueType aFreezeIcon)
{
	bool force_update = false;

	if (aFreezeIcon != NEUTRAL) // i.e. if it's blank, don't change the current setting of mIconFrozen.
	{
		if (mIconFrozen != (aFreezeIcon == TOGGLED_ON)) // It needs to be toggled.
		{
			mIconFrozen = !mIconFrozen;
			force_update = true; // Ensure the icon correctly reflects the current setting and suspend/pause status.
		}
	}
	
	if (aIconFile && *aIconFile == '*' && !aIconFile[1]) // Restore the standard icon.
	{
		if (mCustomIcon)
		{
			GuiType::DestroyIconsIfUnused(mCustomIcon, mCustomIconSmall); // v1.0.37.07: Solves reports of Gui windows losing their icons.
			// If the above doesn't destroy the icon, the GUI window(s) still using it are responsible for
			// destroying it later.
			mCustomIcon = NULL;  // To indicate that there is no custom icon.
			mCustomIconSmall = NULL;
			free(mCustomIconFile);
			mCustomIconFile = NULL;
			mCustomIconNumber = 0;
			force_update = true;
		}
		aIconFile = nullptr; // Update the icon and return, below.
	}
	
	if (!aIconFile) // No icon specified, or it was already reset to default above.
	{
		if (force_update)
			UpdateTrayIcon(true);
		return OK;
	}

	// v1.0.43.03: Load via LoadPicture() vs. ExtractIcon() because ExtractIcon harms the quality
	// of 16x16 icons inside .ico files by first scaling them to 32x32 (which then has to be scaled
	// back to 16x16 for the tray and for the SysMenu icon). I've visually confirmed that the
	// distortion occurs at least when a 16x16 icon is loaded by ExtractIcon() then put into the
	// tray.  It might not be the scaling itself that distorts the icon: the pixels are all in the
	// right places, it's just that some are the wrong color/shade. This implies that some kind of
	// unwanted interpolation or color tweaking is being done by ExtractIcon (and probably LoadIcon),
	// but not by LoadImage.
	// Also, load the icon at actual size so that when/if this icon is used for a GUI window, its
	// appearance in the alt-tab menu won't be unexpectedly poor due to having been scaled from its
	// native size down to 16x16.
	
	if (aIconNumber == 0) // Must validate for use in two places below.
		aIconNumber = 1; // Must be != 0 to tell LoadPicture that "icon must be loaded, never a bitmap".

	int image_type;
	// L17: For best results, load separate small and large icons.
	HICON new_icon_small;
	HICON new_icon = NULL; // Initialize to detect failure to load either icon.
	HMODULE icon_module = NULL; // Must initialize because it's not always set by LoadPicture().
	if (!_tcsnicmp(aIconFile, _T("HICON:"), 6) && aIconFile[6] != '*')
	{
		// Handle this here rather than in LoadPicture() because the first call would destroy the
		// original icon (due to specifying the width and height), causing the second call to fail.
		// Keep the original size for both icons since that sometimes produces better results than
		// CopyImage(), and it keeps the code smaller.
		new_icon_small = (HICON)(UINT_PTR)ATOI64(aIconFile + 6);
		new_icon = new_icon_small; // DestroyIconsIfUnused() handles this case by calling DestroyIcon() only once.
	}
	else if ( new_icon_small = (HICON)LoadPicture(aIconFile, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), image_type, aIconNumber, false) ) // Called with icon_number > 0, it guarantees return of an HICON/HCURSOR, never an HBITMAP.
		if ( !(new_icon = (HICON)LoadPicture(aIconFile, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), image_type, aIconNumber, false, NULL, &icon_module)) )
			DestroyIcon(new_icon_small);
	if ( !new_icon )
		return g_script.RuntimeError(ERR_LOAD_ICON, aIconFile);

	GuiType::DestroyIconsIfUnused(mCustomIcon, mCustomIconSmall); // This destroys it if non-NULL and it's not used by an GUI windows.

	mCustomIcon = new_icon;
	mCustomIconSmall = new_icon_small;
	mCustomIconNumber = aIconNumber;
	
	TCHAR full_path[MAX_PATH], *filename_marker;
	// If the icon was loaded from a DLL, relative->absolute conversion below may produce the
	// wrong result (i.e. in the typical case where the DLL is not in the working directory).
	// So in that case, get the path of the module which contained the icon (if available).
	// Get the full path in case it's a relative path.  This is documented and it's done in case
	// the script ever changes its working directory:
	if (   icon_module && GetModuleFileName(icon_module, full_path, _countof(full_path))
		|| GetFullPathName(aIconFile, _countof(full_path) - 1, full_path, &filename_marker)   )
		aIconFile = full_path;
	free(mCustomIconFile);
	mCustomIconFile = _tcsdup(aIconFile); // Failure isn't checked due to rarity and for simplicity; it'll be reported as empty in that case.

	if (icon_module)
		FreeLibrary(icon_module);

	if (!g_NoTrayIcon)
		UpdateTrayIcon(true);  // Need to use true in this case too.
	return OK;
}



ResultType Script::AutoExecSection()
// Returns FAIL if can't run due to critical error.  Otherwise returns OK.
{
	// Now that g_MaxThreadsTotal has been permanently set by the processing of script directives like
	// #MaxThreads, an appropriately sized array can be allocated:
	if (   !(g_array = (global_struct *)malloc((g_MaxThreadsTotal+TOTAL_ADDITIONAL_THREADS) * sizeof(global_struct)))   )
		return FAIL; // Due to rarity, just abort. It wouldn't be safe to run ExitApp() due to possibility of an OnExit function.
	CopyMemory(g_array, g, sizeof(global_struct)); // Copy the temporary/startup "g" into array[0].
	g = g_array; // Must be done after above.

	// v2: Ensure the Hotkey function defaults to no criterion rather than the last #HotIf WinActive/Exist().  Alternatively we
	// could replace CopyMemory() above with global_init(), but it would need to be changed back if ever we
	// want a directive to affect the default settings.
	g->HotCriterion = NULL;
	
	// Must be done before InitClasses(), otherwise destroying a Gui in a class constructor
	// would terminate the script:
	++g_nThreads;

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
		SET_AUTOEXEC_TIMER(100); // Currently this only marks the auto-execute thread as uninterruptible for the time indicated.
		mAutoExecSectionIsRunning = true;

		// v1.0.25: This is now done here, closer to the actual execution of the first line in the script,
		// to avoid an unnecessary Sleep(10) that would otherwise occur in ExecUntil:
		mLastPeekTime = GetTickCount();

		DEBUGGER_STACK_PUSH(g_AutoExecuteThreadDesc)
		ExecUntil_result = mFirstLine->ExecUntil(UNTIL_RETURN); // Might never return (e.g. infinite loop or ExitApp).
		DEBUGGER_STACK_POP()

		mAutoExecSectionIsRunning = false;
	}
	// REMEMBER: The ExecUntil() call above will never return if the AutoExec section never finishes
	// (e.g. infinite loop) or it uses Exit/ExitApp.

	--g_nThreads;

	// Check if an exception has been thrown
	if (g->ThrownToken)
		g_script.FreeExceptionToken(g->ThrownToken);

	// Now that the auto-execute thread has finished, ensure that the very first g-item is always
	// interruptible.  This avoids having to treat the first g-item as special in various places.
	global_maximize_interruptibility(*g); // See below.

	return ExecUntil_result;
}



bool Script::IsPersistent()
{
	// Consider the script "persistent" if any of the following conditions are true:
	if (Hotkey::sHotkeyCount || Hotstring::sHotstringCount // At least one hotkey or hotstring exists.
		// No attempt is made to determine if the hotkeys/hotstrings are enabled, since even if they
		// are, it's impossible to detect whether #HotIf will allow them to ever execute.
		|| g_persistent // Persistent() has been used somewhere in the script.
		|| g_script.mTimerEnabledCount // At least one script timer is currently enabled.
		|| mOnClipboardChange.Count() // The script is monitoring clipboard changes.
		|| g_input // At least one active InputHook.
		|| IsWindowVisible(g_hWnd))
		return true;
	// OnMessage does not make the script persistent because:
	//  1) It has too many uses to assume that the author would want the script to be persistent
	//     (which would probably only be in cases where OnMessage is the main purpose of the script,
	//     such as if it just idles until a command is received via WM_COPYDATA or other message).
	//  2) It's often used in conjunction with a GUI, and in such cases it's probably more convenient
	//     to rely on other logic to keep the script running while the GUI is visible, and allow it to
	//     exit when the GUI is closed.
	//  3) Message monitors that are intended to keep the script running probably won't be removed,
	//     so checking the number of active monitors here at runtime doesn't offer much benefit vs.
	//     requiring the script to call Persistent() if needed.
	//  4) The only alternatives to OnMessage are generally window subclassing and hooks, which come
	//     with various problems, such as spamming ListLines, interacting badly with Pause/A_IsPaused
	//     (due to thread-starting messages passing through the window proc), and others.
	// Adding custom items to A_TrayMenu does not make the script persistent because:
	//  1) It could be something like a slight modification to a standard item's behaviour,
	//     or something else that isn't intended to keep the script running.
	//  2) Custom items usually aren't removed, so having checks for them at runtime is not as helpful
	//     as some of the conditions above, like timers and visible windows.  When they are removed, it
	//     is always a direct action of the code, not some external event.
	//  3) If adding custom items made the script persistent, it can be worked around by intercepting
	//     right-clicks on the tray icon and showing some other menu, but replicating the Pause/Suspend
	//     checkmarks and AddStandard() behaviour is surprisingly complicated.  By contrast, calling
	//     Persistent() before/after adding the custom items is trivial.
	for (GuiType* gui = g_firstGui; gui; gui = gui->mNextGui)
		if (IsWindowVisible(gui->mHwnd)) // A GUI is visible.
			return true;
	// Otherwise, none of the above conditions are true; but there might still be
	// one or more script threads running.  Caller is responsible for checking that.
	return false;
}



void Script::ExitIfNotPersistent(ExitReasons aExitReason)
{
	if (g_nThreads || IsPersistent())
		return;
	g_script.ExitApp(aExitReason);
}



bif_impl void Edit(optl<StrArg> aFileName)
{
	g_script.Edit(aFileName.value_or_null());
}

ResultType Script::Edit(LPCTSTR aFileName)
{
#ifdef AUTOHOTKEYSC
	return OK; // Do nothing.
#else
	HWND hwnd = NULL;
	if (!aFileName)
	{
		if (mKind != ScriptKindFile)
			return OK;
		TitleMatchModes old_mode = g->TitleMatchMode;
		g->TitleMatchMode = FIND_ANYWHERE;
		hwnd = WinExist(*g, mFileSpec, _T(""), mMainWindowTitle, _T("")); // Exclude our own main window.
		if (!hwnd)
			hwnd = WinExist(*g, mFileName, _T(""), mMainWindowTitle, _T("")); // Try again without path.
		g->TitleMatchMode = old_mode;
		aFileName = mFileSpec;
	}
	if (hwnd)
	{
		TCHAR class_name[32];
		GetClassName(hwnd, class_name, _countof(class_name));
		if (!_tcscmp(class_name, _T("#32770")) || !_tcsnicmp(class_name, _T("AutoHotkey"), 10)) // MessageBox(), InputBox(), FileSelect(), or GUI/script-owned window.
			hwnd = NULL;  // Exclude it from consideration.
	}
	if (hwnd)  // File appears to already be open for editing, so use the current window.
		SetForegroundWindowEx(hwnd);
	else
	{
		TCHAR buf[T_MAX_PATH + 2]; // +2 for the two quote marks.
		// Enclose in double quotes anything that might contain spaces since the CreateProcess()
		// method, which is attempted first, is more likely to succeed.  This is because it uses
		// the command line method of creating the process, with everything all lumped together:
		sntprintf(buf, _countof(buf), _T("\"%s\""), aFileName);
		if (!ActionExec(_T("Edit"), buf, mFileDir, false))  // Since this didn't work, try notepad.
		{
			// Even though notepad properly handles filenames with spaces in them under WinXP,
			// even without double quotes around them, it seems safer and more correct to always
			// enclose the filename in double quotes for maximum compatibility with all OSes:
			if (!ActionExec(_T("notepad.exe"), buf, mFileDir, false))
				MsgBox(_T("Could not open script.")); // Short message since so rare.
		}
	}
	return OK;
#endif
}



bif_impl FResult Reload()
{
	// If Reload() returns failure, that means a failure to launch the new process at all,
	// and an error message was displayed or an exception was thrown, so we must return FR_FAIL.
	// If Reload() returns success but the script doesn't terminate even after some delay,
	// it can be assumed that the reload didn't work, and the script can take follow-on action.
	return g_script.Reload(true) ? OK : FR_FAIL;
}

ResultType Script::Reload(bool aDisplayErrors)
{
	// The new instance we're about to start will tell our process to stop, or it will display
	// a syntax error or some other error, in which case our process will still be running:
#ifdef AUTOHOTKEYSC
	// This is here in case a compiled script ever uses the Reload function.  Since the "Reload Script"
	// tray menu item is not available for compiled scripts, it can't be called from there.
	return g_script.ActionExec(mOurEXE, _T("/restart"), g_WorkingDirOrig, aDisplayErrors);
#else
	if (mKind == ScriptKindStdIn)
		return OK;
	TCHAR arg_string[UorA(MAX_WIDE_PATH, MAX_PATH * 2 + 16)]; // MAX_WIDEPATH coincides with the CreateProcess command line length limit (+1).
	if (mCmdLineInclude)
		sntprintf(arg_string, _countof(arg_string), _T("/include \"%s\" "), mCmdLineInclude);
	else
		*arg_string = '\0';
	sntprintfcat(arg_string, _countof(arg_string), _T("/restart /script \"%s\""), Line::sSourceFile[0]);
	return g_script.ActionExec(mOurEXE, arg_string, g_WorkingDirOrig, aDisplayErrors);
#endif
}



bif_impl ResultType Exit(optl<int> aExitCode)
{
	// Even if the script isn't persistent, this thread might've interrupted another which should
	// be allowed to complete normally.  This is especially important in v2 because a persistent
	// script can become non-persistent by disabling a timer, closing a GUI, etc.  So if there
	// are any threads below this one, only exit this thread:
	//if (g_nThreads > 1 || g_script.IsPersistent())
	// UPDATE: Handle it this way unconditionally so that the thread is properly exited prior to
	// the script terminating; i.e. any FINALLY statements are executed and __delete is called for
	// any objects in local variables or on the expression evaluation stack.
	// UPDATE: Also need to make sure that the exit code gets used (commit 40dd95f broke that).
	// For now it seems best to match the old behaviour of ignoring the exit code if Exit isn't
	// going to immediately terminate the script.  Avoid checking IsPersistent() here because
	// conditions can change during stack-unwind due to __delete or FINALLY.  Instead, this is
	// reset to 0 in ResumeUnderlyingThread().
	if (g_nThreads <= 1)
		g_script.mPendingExitCode = aExitCode.has_value() ? *aExitCode : 0;
	return EARLY_EXIT;
}

bif_impl ResultType ExitApp(optl<int> aExitCode)
{
	g_script.mPendingExitCode = aExitCode.has_value() ? *aExitCode : 0;
	return g_script.ExitApp(EXIT_EXIT);
}



ResultType Script::ExitApp(ExitReasons aExitReason)
// Normal exit (if aBuf is NULL), or a way to exit immediately on error (which is mostly
// for times when it would be unsafe to call MsgBox() due to the possibility that it would
// make the situation even worse).
{
	int aExitCode = mPendingExitCode; // It's done this way so Exit() can pass an exit code indirectly.

	// Note that currently, mOnExit.Count() can only be non-zero if the script is in a runnable
	// state (since registering an OnExit function requires that the script calls OnExit()).
	// If this ever changes, the !mIsReadyToExecute condition should be added below:
	if (!mOnExit.Count() || g_OnExitIsRunning)
	{
		// In the case of g_OnExitIsRunning == true:
		// There is another instance of this function beneath us on the stack.  Since we have
		// been called, this is a true exit condition and we exit immediately.
		// MUST NOT create a new thread when g_OnExitIsRunning because g_array allows only one
		// extra thread for ExitApp() (which allows it to run even when MAX_THREADS_EMERGENCY has
		// been reached).  See TOTAL_ADDITIONAL_THREADS.
		TerminateApp(aExitReason, aExitCode);
	}

	// Otherwise, the script contains the special OnExit function that we will run here instead
	// of exiting.  And since it does, we know that the script is in a ready-to-execute state
	// because that is the only way an OnExit function could have been defined in the first place.
	// The OnExit function may return true in order to abort the Exit sequence (e.g. it could
	// display an "Are you sure?" prompt, and if the user chooses "No", it returns true, letting
	// the thread we create below end normally).

	// Next, save the current state of the globals so that they can be restored just prior
	// to returning to our caller:
	InitNewThread(0, true, true); // Uninterruptibility is handled below. Since this special thread should always run, no checking of g_MaxThreadsTotal is done before calling this.

	// Turn on uninterruptibility to forbid any hotkeys, timers, or user defined menu items
	// to interrupt.  This is mainly done for peace-of-mind (since possible interactions due to
	// interruptions have not been studied) and the fact that this most users would not want this
	// subroutine to be interruptible (it usually runs quickly anyway).  Another reason to make
	// it non-interruptible is that some OnExit functions might destruct things used by the
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
	//    might be currently pending.  When it fires, it would make the OnExit function interruptible
	//    rather than the underlying subroutine.  The above fixes the first part of that problem.
	//    The 2nd part is fixed by reinstating the timer when the uninterruptible thread is resumed.
	//    This special handling is only necessary here -- not in other places where new threads are
	//    created -- because OnExit is the only type of thread that can interrupt an uninterruptible
	//    thread.
	BOOL g_AllowInterruption_prev = g_AllowInterruption;  // Save current setting.
	g_AllowInterruption = FALSE; // Mark the thread just created above as permanently uninterruptible (i.e. until it finishes and is destroyed).

	g_OnExitIsRunning = true;
	DEBUGGER_STACK_PUSH(_T("OnExit"))

	ExprTokenType param[] = { GetExitReasonString(aExitReason), (__int64)aExitCode };
	ResultType result = mOnExit.Call(param, _countof(param), 1);
	
	DEBUGGER_STACK_POP()

	if (result != CONDITION_TRUE // OnExit function did not return true to prevent exit.
		|| EXITREASON_MUST_EXIT(aExitReason)) // Caller requested we exit unconditionally.
		TerminateApp(aExitReason, aExitCode);

	// Otherwise:
	g_AllowInterruption = g_AllowInterruption_prev;  // Restore original setting.
	ResumeUnderlyingThread();
	// If this OnExit thread is the last script thread and the script is not persistent, the above
	// call recurses into this function.  g_OnExitIsRunning == true prevents infinite recursion
	// in that case.  It is now safe to reset:
	g_OnExitIsRunning = false;  // In case the user wanted the thread to end normally (see above).

	return EARLY_EXIT;
}



void ReleaseVarObjects(VarList &aVars)
{
	for (int v = 0; v < aVars.mCount; ++v)
	{
		Var &var = *aVars.mItem[v];
		if (var.IsObject() && var.Type() == VAR_NORMAL)
			var.ReleaseObject(); // ReleaseObject() vs Free() for performance (though probably not important at this point).
		// Otherwise, maybe best not to free it in case an object's __Delete meta-function uses it?
	}
}

void ReleaseStaticVarObjects(FuncList &aFuncs)
{
	for (int i = 0; i < aFuncs.mCount; ++i)
	{
		ASSERT(dynamic_cast<UserFunc*>(aFuncs.mItem[i]));
		auto &f = *(UserFunc *)aFuncs.mItem[i];
		// Since it doesn't seem feasible to release all var backups created by recursive function
		// calls and all tokens in the 'stack' of each currently executing expression, currently
		// only static and global variables are released.  It seems best for consistency to also
		// avoid releasing top-level non-static local variables (i.e. which aren't in var backups).
		ReleaseVarObjects(f.mStaticVars);
	}
}



void Script::TerminateApp(ExitReasons aExitReason, int aExitCode)
// Note that g_script's destructor takes care of most other cleanup work, such as destroying
// tray icons, menus, and unowned windows such as ToolTip.
{
	// L31: Release objects stored in variables, where possible.
	if (aExitReason != EXIT_CRITICAL) // i.e. Avoid making matters worse if EXIT_CRITICAL.
	{
		// Ensure the current thread is not paused and can't be interrupted
		// in case one or more objects need to call a __delete meta-function.
		g_AllowInterruption = FALSE;
		g->IsPaused = false;

		ReleaseVarObjects(mVars);
		ReleaseStaticVarObjects(mFuncs);
	}
#ifdef CONFIG_DEBUGGER // L34: Exit debugger *after* the above to allow debugging of any invoked __Delete handlers.
	g_Debugger.Exit(aExitReason);
#endif

	// PostQuitMessage() might be needed to prevent hang-on-exit.  Once this is done, no message boxes or
	// other dialogs can be displayed.  MSDN: "The exit value returned to the system must be the wParam
	// parameter of the WM_QUIT message."  In our case, PostQuitMessage() should announce the same exit code
	// that we will eventually call exit() with:
	PostQuitMessage(aExitCode);

	// I know this isn't the preferred way to exit the program.  However, due to unusual
	// conditions such as the script having MsgBoxes or other dialogs displayed on the screen
	// at the time the user exits (in which case our main event loop would be "buried" underneath
	// the event loops of the dialogs themselves), this is the only reliable way I've found to exit
	// so far.
	exit(aExitCode); // exit() is insignificant in code size.  It does more than ExitProcess(), but perhaps nothing more that this application actually requires.
}



UINT Script::LoadFromFile(LPCTSTR aFileSpec)
// Returns the number of non-comment lines that were loaded, or LOADING_FAILED on error.
{
	mIsReadyToExecute = mAutoExecSectionIsRunning = false;
	if (!aFileSpec)
		aFileSpec = mFileSpec;

#ifndef AUTOHOTKEYSC  // When not in stand-alone mode, read an external script file.
	DWORD attr = mKind == ScriptKindFile ? GetFileAttributes(aFileSpec) : 0; // v1.1.17: Don't check if reading script from stdin.
	if (attr == MAXDWORD) // File does not exist or lacking the authorization to get its attributes.
	{
		if (!g_script.mErrorStdOut)
		{
			TCHAR buf[T_MAX_PATH + 24];
			sntprintf(buf, _countof(buf), _T("%s\n%s"), ERR_SCRIPT_NOT_FOUND, aFileSpec);
			MsgBox(buf, MB_ICONHAND);
		}
		else
		{
			Line::sSourceFile = &mFileSpec;
			ScriptError(ERR_SCRIPT_NOT_FOUND);
		}
		return LOADING_FAILED;
	}
#endif

	// Load the main script file.  This will also load any files it includes with #Include.
	if (!LoadIncludedFile(mKind == ScriptKindStdIn ? _T("*") : aFileSpec, false, false))
		return LOADING_FAILED;
		
	g_SuspendExempt = false; // #SuspendExempt should not affect Hotkey()/Hotstring().

#ifdef ENABLE_DLLCALL
	// So that (the last occuring) "#DllLoad directory" doesn't affect calls to GetDllProcAddress for run time calls to DllCall
	// or DllCall optimizations in Line::ExpressionToPostfix.
	if (!SetDllDirectory(NULL))
	{
		ScriptError(ERR_INTERNAL_CALL);
		return LOADING_FAILED;
	}
#endif
	// Preparse all expressions and resolve all variable references.  The outer-most scope
	// is preparsed first, then each function, working inward through all nested functions.
	// All of a function's non-dynamic local variables are created before variable names
	// are resolved in nested functions.
	if (!PreparseExpressions(mFirstLine)
		|| !PreparseExpressions(mHotFuncs) // mHotFuncs first in case they have nested functions, which would be in mFuncs.
		|| !PreparseExpressions(mFuncs)
		|| !PreparseVarRefs())
		return LOADING_FAILED; // Error was already displayed by the above call.

	// Do some processing of local variables to support closures.
	// This must be done after PreparseExpressions() has resolved all variable references.
	if (!PreprocessLocalVars(mHotFuncs) // mHotFuncs first in case they have nested functions, which would be in mFuncs.
		|| !PreprocessLocalVars(mFuncs))
		return LOADING_FAILED;

	free(mHotFuncs.mItem); // Not needed beyond this point.

	// Resolve any unresolved base classes.
	if (mUnresolvedClasses)
	{
		if (!ResolveClasses())
			return LOADING_FAILED;
		mUnresolvedClasses->Release();
		mUnresolvedClasses = NULL;
	}

	// Set the working directory to the script's directory.  This must be done after the above
	// since the working dir may have been changed by the script's use of "#Include C:\Scripts".
	// LoadIncludedFile() also changes it, but any value other than mFileDir would have been
	// restored by "#Include" after LoadIncludedFile() returned.  Note that A_InitialWorkingDir
	// contains the startup-determined working directory, so no flexibility is lost.
	SetCurrentDirectory(mFileDir);
	g_WorkingDir.SetString(mFileDir);

	// Even if the last line of the script is already "exit", always add another
	// one in case the script ends in a label.  That way, every label will have
	// a non-NULL target, which simplifies other aspects of script execution.
	// Making sure that all scripts end with an EXIT ensures that if the script
	// file ends with ELSEless IF or an ELSE, that IF's or ELSE's mRelatedLine
	// will be non-NULL, which further simplifies script execution.
	if (!AddLine(ACT_EXIT))
		return LOADING_FAILED;

	if (!PreparseCommands(mFirstLine))
		return LOADING_FAILED; // Error was already displayed by the above calls.
	
#ifndef AUTOHOTKEYSC
	if (mValidateThenExit)
		return 0; // Tell our caller to do a normal exit.
#endif

	return TRUE; // Must be non-zero.
}



bool Script::IsFunctionDefinition(LPTSTR aBuf, LPTSTR aNextBuf)
// Helper function for LoadIncludedFile().
// Caller passes in an aBuf containing a candidate line such as "function(x, y)"
// Caller has ensured that aBuf is rtrim'd.
// Caller should pass NULL for aPendingFunctionHasBrace to indicate that function definitions (open-brace
// on same line as function) are not allowed.  When non-NULL *and* aBuf is a function call/def,
// *aPendingFunctionHasBrace is set to true if a brace is present at the end, or false otherwise.
// In addition, any open-brace is removed from aBuf in this mode.
{
	LPTSTR action_start = aBuf;
	LPTSTR action_end = find_identifier_end(aBuf);
	// Allow the "static" keyword for preventing a nested function from becoming a closure.
	if (IS_SPACE_OR_TAB(*action_end) && !_tcsnicmp(aBuf, _T("Static"), 6))
		action_end = find_identifier_end(action_start = omit_leading_whitespace(action_end));
	// Can't be a function definition or call without an open-parenthesis as first char found by the above.
	// action_end points at the first character which is not usable in an identifier, such as a space, tab
	// colon or other operator symbol.  As a result, it can't be:
	// 1) a hotstring, since they always start with at least one colon that would be caught immediately as 
	//    first-expr-char-is-not-open-parenthesis by the above.
	// 2) Any kind of math or assignment, such as var:=(x+y) or var+=(x+y).
	// The only things it could be other than a function call or function definition are:
	// Single-line hotkey such as KeyName::MsgBox.  But (:: is the only valid hotkey where *action_end == '(',
	// and that's handled by excluding action_end == aBuf.
	if (*action_end != '(' || action_end == action_start)
		return false;
	// Is it a control flow statement, such as "if(condition)"?
	*action_end = '\0';
	bool is_control_flow = ConvertActionType(aBuf);
	*action_end = '(';
	if (is_control_flow)
		return false;
	// It's not control flow.
	LPTSTR param_end = action_end + FindExprDelim(action_end, ')', 1);
	if (*param_end != ')')
		return false;
	LPTSTR next_token = omit_leading_whitespace(param_end + 1);
	return *next_token == 0 && *aNextBuf == '{' // Brace on next line.
		|| *next_token == '{' && next_token[1] == 0 // Brace on same line.
		|| *next_token == '=' && next_token[1] == '>'; // Fn() => expr
}



inline LPTSTR IsClassDefinition(LPTSTR aBuf)
{
	if (_tcsnicmp(aBuf, _T("Class"), 5) || !IS_SPACE_OR_TAB(aBuf[5])) // i.e. it's not "Class" followed by a space or tab.
		return NULL;
	LPTSTR class_name = omit_leading_whitespace(aBuf + 6);
	if (_tcschr(EXPR_ALL_SYMBOLS, *class_name))
		// It's probably something like "Class := GetClass()".
		return NULL;
	// Validation of the name is left up to the caller, for simplicity.
	return class_name;
}



void RemoveBufChar0(LPTSTR aBuf, size_t &aBufLength)
{
	LPTSTR cp = omit_leading_whitespace(aBuf + 1);
	aBufLength -= (cp - aBuf);
	if (aBufLength) // Some non-whitespace remains.
		tmemmove(aBuf, aBuf + 1, aBufLength);
	else
		*aBuf = '\0';
}



bool ClassHasOpenBrace(LPTSTR aBuf, size_t aBufLength, LPTSTR aNextBuf, size_t &aNextBufLength)
{
	if (aBuf[aBufLength - 1] == '{') // Brace on same line (OTB).
	{
		rtrim(aBuf, aBufLength - 1);
		return true;
	}
	if (*aNextBuf == '{') // Brace on next line.
	{
		// Remove '{' from aNextBuf since ACT_BLOCK_END is unwanted in this context.
		RemoveBufChar0(aNextBuf, aNextBufLength);
		return true;
	}
	return false;
}



ResultType Script::OpenIncludedFile(TextStream *&ts, LPCTSTR aFileSpec, bool aAllowDuplicateInclude, bool aIgnoreLoadFailure)
// Open the included file.  Returns CONDITION_TRUE if the file is to
// be loaded, otherwise OK (duplicate/already loaded) or FAIL (error).
// See "full_path" below for why this is separate to LoadIncludedFile().  
{
#ifndef AUTOHOTKEYSC
	if (!aFileSpec || !*aFileSpec) return FAIL;

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

	// Use of stack memory here to build the full path is the most efficient method,
	// but utilizes 64KB per buffer on Unicode builds.  There is virtually no cost
	// when used here, but if used directly in LoadIncludedFile(), this would mean
	// 64KB used *for each instance on the stack*, which significantly reduces the
	// recursion limit for #include inside #include.  Note that enclosing the buf
	// within a limited scope is insufficient, as the compiler will (or may) still
	// allocate the required stack space on entry to the function.
	TCHAR full_path[T_MAX_PATH];

	int source_file_index = Line::sSourceFileCount;
	if (!source_file_index && *aFileSpec != '*')
		// Since this is the first source file, it must be the main script file.  Just point it to the
		// location of the filespec already dynamically allocated:
		Line::sSourceFile[source_file_index] = mFileSpec;
	else
	{
		// Get the full path in case aFileSpec has a relative path.  This is done so that duplicates
		// can be reliably detected (we only want to avoid including a given file more than once):
		LPTSTR filename_marker;
		if (*aFileSpec == '*')
		{
			tcslcpy(full_path, aFileSpec, _countof(full_path));
			CharUpper(full_path); // Resource names must be upper-case.
		}
		else
			GetFullPathName(aFileSpec, _countof(full_path), full_path, &filename_marker);
		// Check if this file was already included.  If so, it's not an error because we want
		// to support automatic "include once" behavior.  So just ignore repeats:
		if (!aAllowDuplicateInclude)
			for (int f = 0; f < source_file_index; ++f) // Here, source_file_index==Line::sSourceFileCount
				if (!lstrcmpi(Line::sSourceFile[f], full_path)) // Case insensitive like the file system (testing shows that "Ä" == "ä" in the NTFS, which is hopefully how lstrcmpi works regardless of locale).
					return OK;
		// The path is copied into persistent memory further below, after the file has been opened,
		// in case the opening fails and aIgnoreLoadFailure==true.  Initialize for the check below.
		Line::sSourceFile[source_file_index] = NULL;
	}

	LPCTSTR filespec_to_open = aFileSpec;
	UINT codepage = g_DefaultScriptCodepage;

	TextMem::Buffer textbuf(nullptr, 0, false);
	HRSRC hRes;
	if (*aFileSpec == '*' && aFileSpec[1] && (hRes = FindResource(NULL, aFileSpec + 1, RT_RCDATA)))
	{
		HGLOBAL hResData = LoadResource(NULL, hRes);
		if (hResData && (textbuf.mBuffer = LockResource(hResData)))
		{
			textbuf.mLength = SizeofResource(NULL, hRes);
			filespec_to_open = textbuf;
			codepage = CP_UTF8;
			ts = new TextMem();
		}
	}
	else
		ts = new TextFile();

	if (!ts || !ts->Open(filespec_to_open, DEFAULT_READ_FLAGS, codepage))
	{
		if (aIgnoreLoadFailure)
			return OK;
		TCHAR msg_text[T_MAX_PATH + 64]; // T_MAX_PATH vs. MAX_PATH because the full length could be utilized with ErrorStdOut.
		sntprintf(msg_text, _countof(msg_text), _T("%s file \"%s\" cannot be opened.")
			, source_file_index ? _T("#Include") : _T("Script"), aFileSpec);
		return ScriptError(msg_text);
	}
	
	// This is done only after the file has been successfully opened in case aIgnoreLoadFailure==true:
	if (!Line::sSourceFile[source_file_index])
		Line::sSourceFile[source_file_index] = SimpleHeap::Alloc(full_path);
	//else the first file was already taken care of by another means.
	++Line::sSourceFileCount;

	if (!source_file_index)
	{
		// Load any pre-script resource.
		if (FindResource(NULL, SCRIPT_PRESOURCE_NAME, RT_RCDATA)
			&& !LoadIncludedFile(SCRIPT_PRESOURCE_SPEC, false, false))
			return FAIL;

		// Load any file specified with the /include switch.
		if (mCmdLineInclude
			&& !LoadIncludedFile(mCmdLineInclude, false, false))
			return FAIL;
	}
	
	// Set the working directory so that any #Include directives are relative to the directory
	// containing this file by default.  Call SetWorkingDir() vs. SetCurrentDirectory() so that it
	// succeeds even for a root drive like C: that lacks a backslash (see SetWorkingDir() for details).
	if (source_file_index)
	{
		LPTSTR terminate_here = _tcsrchr(full_path, '\\');
		if (terminate_here > full_path)
		{
			*terminate_here = '\0'; // Temporarily terminate it for use with SetWorkingDir().
			SetWorkingDir(full_path);
			*terminate_here = '\\'; // Undo the termination.
		}
		//else: probably impossible? Just leave the working dir as-is, for simplicity.
	}
	else
		SetWorkingDir(mFileDir);

#else // Stand-alone mode (there are no include files in this mode since all of them were merged into the main script at the time of compiling).

	TextMem::Buffer textbuf(NULL, 0, false);

	HRSRC hRes;
	HGLOBAL hResData;

	hRes = FindResource(NULL, SCRIPT_RESOURCE_NAME, RT_RCDATA);
	if ( !( hRes 
			&& (textbuf.mLength = SizeofResource(NULL, hRes))
			&& (hResData = LoadResource(NULL, hRes))
			&& (textbuf.mBuffer = LockResource(hResData)) ) )
	{
		return ScriptError(_T("Could not extract script from EXE."), aFileSpec);
	}

	ts = new TextMem();
	// NOTE: Ahk2Exe strips off the UTF-8 BOM.
	ts->Open(textbuf, TextStream::READ | TextStream::EOL_CRLF | TextStream::EOL_ORPHAN_CR, CP_UTF8);

	// Since this is a compiled script, there is only one script file.
	// Just point it to the location of the filespec already dynamically allocated:
	Line::sSourceFile[0] = mFileSpec;
	++Line::sSourceFileCount;

#endif
	
	// Since above did not continue, proceed with loading the file.
	return CONDITION_TRUE;
}



ResultType Script::LoadIncludedFile(LPCTSTR aFileSpec, bool aAllowDuplicateInclude, bool aIgnoreLoadFailure)
// Returns OK or FAIL.
{
	int source_file_index = Line::sSourceFileCount; // Set early in case of >PRESCRIPT<.
	TextStream *ts = nullptr;
	ResultType result = OpenIncludedFile(ts, aFileSpec, aAllowDuplicateInclude, aIgnoreLoadFailure);
	if (result == CONDITION_TRUE)
		result = LoadIncludedFile(ts, source_file_index);
	if (ts)
		delete ts;
	return result;
}



ResultType Script::LoadIncludedFile(TextStream *fp, int aFileIndex)
// Returns OK or FAIL.
{
	// Keep this var on the stack due to recursion, which allows newly created lines to be given the
	// correct file number even when some #include's have been encountered in the middle of the script:
	int source_file_index = aFileIndex;

	bool blocks_previously_open = mOpenBlock || mClassObjectCount; // For error detection.

	LineBuffer buf, next_buf;
	size_t &buf_length = buf.length, &next_buf_length = next_buf.length;
	bool buf_has_brace;

	// File is now open, read lines from it.

	bool has_continuation_section;
	TCHAR orig_char;

	LPTSTR hotkey_flag, cp, cp1, hotstring_start, hotstring_options;
	Hotkey *hk;
	LineNumberType saved_line_number;
	HookActionType hook_action;
	bool hook_is_mandatory, hotstring_execute;
	UCHAR no_suppress;
	ResultType hotkey_validity;

	// Init both for main file and any included files loaded by this function:
	mCurrFileIndex = source_file_index;  // source_file_index is kept on the stack due to recursion (from #include).

#ifdef AUTOHOTKEYSC
	// -1 (MAX_UINT in this case) to compensate for the fact that there is a comment containing
	// the version number added to the top of each compiled script:
	LineNumberType phys_line_number = -1;
#else
	LineNumberType phys_line_number = 0;
#endif
	
	if (!buf.Expand() || !next_buf.Expand())
		return ScriptError(ERR_OUTOFMEM);

	// buf is initialized this way rather than calling GetLine() to simplify handling of comment
	// sections beginning at the first line, and to reduce code size by having GetLine() only
	// called from one place:
	*buf = '\0';

	while (buf_length != -1)  // Compare directly to -1 since length is unsigned.
	{
		// For each whole line (a line with continuation section is counted as only a single line
		// for the purpose of this outer loop).

		// Keep track of this line's *physical* line number within its file for A_LineNumber and
		// error reporting purposes.  This must be done only in the outer loop so that it tracks
		// the topmost line of any set of lines merged due to continuation section/line(s)..
		mCombinedLineNumber = phys_line_number;

		// This must be reset for each iteration because a prior iteration may have changed it, even
		// indirectly by calling something that changed it:
		mCurrLine = NULL;  // To signify that we're in transition, trying to load a new one.

		if (!GetLineContinuation(fp, buf, next_buf, phys_line_number, has_continuation_section))
			return FAIL;

process_completed_line:
		// buf_length can't be -1 (though next_buf_length can) because outer loop's condition prevents it:
		if (!buf_length) // Done only after the line number increments above so that the physical line number is properly tracked.
			goto continue_main_loop; // In lieu of "continue", for performance.

		// Since neither of the above executed, or they did but didn't "continue",
		// buf now contains a non-commented line, either by itself or built from
		// any continuation sections/lines that might have been present.  Also note that
		// by design, phys_line_number will be greater than mCombinedLineNumber whenever
		// a continuation section/lines were used to build this combined line.

		hotstring_start = NULL;
		hotkey_flag = NULL;
		bool hotkey_uses_otb = false; // used for hotstrings too.
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
			// This is necessary for it to handle cases such as the following:
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
					// v2: Use the first non-escaped double-colon, not the last, since it seems more likely
					// that the user intends to produce text with "::" in it rather than typing "::" to trigger
					// the hotstring, and generally the trigger should be short.  By contrast, the v1 policy
					// behaved inconsistently with an odd number of colons, such as:
					//   ::foo::::bar  ; foo:: -> bar
					//   ::foo:::bar   ; foo -> :bar
					if (!hotkey_flag && *cp1 == ':') // Found a non-escaped double-colon, so this is the right one.
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
					case 's': *cp1 = ' '; break;   // space
					// Otherwise, if it's not one of the above, the escape-char is considered to
					// mark the next character as literal, regardless of what it is. Examples:
					// `` -> `
					// `: -> : (so `::: means a literal : followed by hotkey_flag)
					// `; -> ;
					// `c -> c (i.e. unknown escape sequences resolve to the char after the `)
				}
				// Below has a final +1 to include the terminator:
				tmemmove(cp, cp1, _tcslen(cp1) + 1);
				// v2: The following is not done because 1) it is counter-intuitive for ` to affect two
				// characters and 2) it hurts flexibility by preventing the escaping of a single colon
				// immediately prior to the double-colon, such as ::lbl`:::.  Older comment:
				// Since single colons normally do not need to be escaped, this increments one extra
				// for double-colons to skip over the entire pair so that its second colon
				// is not seen as part of the hotstring's final double-colon.  Example:
				// ::ahc```::::Replacement String
				//if (*cp == ':' && *cp1 == ':')
				//	++cp;
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
					// Above: Hotkeys like "^:::" and "l & :::" are not supported because: 1) some cases are
					// ambiguous, such as "^:::" legitimately remapping caret to colon; 2) retaining support
					// for colon as a remap target would require larger/more complicated code; 3) such hotkeys
					// are hard for a human to read/interpret.
				// v1.0.40: It appears to be a hotkey, but validate it as such before committing to processing
				// it as a hotkey.  If it fails validation as a hotkey, treat it as a command that just happens
				// to contain a double-colon somewhere.  This avoids the need to escape double colons in scripts.
				// Note: Hotstrings can't suffer from this type of ambiguity because a leading colon or pair of
				// colons makes them easier to detect.
				cp = omit_trailing_whitespace(buf, hotkey_flag); // For maintainability.
				orig_char = *cp;
				*cp = '\0'; // Temporarily terminate.
				hotkey_validity = Hotkey::TextInterpret(omit_leading_whitespace(buf), NULL); // Passing NULL calls it in validate-only mode.
				switch (hotkey_validity)
				{
				case FAIL:
					return FAIL; // It's an invalid hotkey and above already displayed the error message.
				case CONDITION_FALSE:
					hotkey_flag = NULL; // It doesn't look like valid hotkey syntax, so parse it as something else (so the error message won't be ERR_INVALID_KEYNAME).
					break;
				//case CONDITION_TRUE:
					// It's a key that doesn't exist on the current keyboard layout.  Leave hotkey_flag set
					// so that the section below handles it as a hotkey.  This ensures any same-line action
					// or trailing block is interpreted correctly.  A warning will be displayed below.
				}
				*cp = orig_char; // Undo the temp. termination above.
			}
		}

		if (hotkey_flag && hotkey_flag > buf) // It's a hotkey/hotstring label.
		{
			
			// Allow a current function if it is mLastHotFunc, this allows stacking,
			// x::			// mLastHotFunc created here
			// y::action	// parsing "y::" now.
			if ( (g->CurrentFunc && g->CurrentFunc != mLastHotFunc) || mClassObjectCount)
			{
				// The reason for not allowing hotkeys and hotstrings inside a function's body is that it
				// would create a nested function that could be called without ever calling the outer function.
				// If the hotkey function became a closure, activating the hotkey would at best merely raise an
				// error since it would not be associated with any particular call to the outer function.
				// Currently CreateHotFunc() isn't set up to permit it; e.g. it doesn't set mOuterFunc or
				// mDownVar, so a hotkey inside a function would reset g->CurrentFunc to nullptr for the
				// remainder of the outer function and would crash if the hotkey references any outer vars.
				return ScriptError(_T("Hotkeys/hotstrings are not allowed inside functions or classes."), buf);
			}

			*hotkey_flag = '\0'; // Terminate so that buf is now the hotkey's name.
			hotkey_flag += HOTKEY_FLAG_LENGTH;  // Now hotkey_flag is the hotkey's action, if any.
			
			LPTSTR next_nonspace = omit_leading_whitespace(hotkey_flag);
			hotkey_uses_otb = *next_nonspace == '{' && !next_nonspace[1]; // Has already been rtrimmed by GetLine().
			if (!hotstring_start || hotstring_execute)
			{
				// Mustn't use ltrim(hotkey_flag) because that would cause buf.length to become incorrect:
				hotkey_flag = next_nonspace;
			}
			// else don't trim auto-replace hotstrings since literal spaces in both substrings are significant.
			if (!hotstring_start)
			{
				// To use '{' as remap_dest, escape it.
				if (hotkey_flag[0] == g_EscapeChar && hotkey_flag[1] == '{')
					hotkey_flag++;

				cp = hotkey_flag; // Set default, conditionally overridden below (v1.0.44.07).
				vk_type remap_dest_vk;
				// v1.0.40: Check if this is a remap rather than hotkey:
				if (!hotkey_uses_otb   
					&& *hotkey_flag // This hotkey's action is on the same line as its trigger definition.
					&& (remap_dest_vk = hotkey_flag[1] ? TextToVK(cp = const_cast<LPTSTR>(Hotkey::TextToModifiers(hotkey_flag, NULL))) : 0xFF)   ) // And the action appears to be a remap destination rather than a command.
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
					
					vk_type remap_source_vk;
					TCHAR remap_source[32], remap_dest[32], remap_dest_modifiers[8]; // Must fit the longest key name (currently Browser_Favorites [17]), but buffer overflow is checked just in case.
					bool remap_source_is_combo, remap_source_is_mouse, remap_dest_is_mouse, remap_keybd_to_mouse;

					// These will be ignored in other stages if it turns out not to be a remap later below:
					LPCTSTR cp1;
					remap_source_vk = TextToVK(cp1 = Hotkey::TextToModifiers(buf, NULL)); // An earlier stage verified that it's a valid hotkey, though VK could be zero.
					remap_source_is_combo = _tcsstr(cp1, COMPOSITE_DELIMITER);
					remap_source_is_mouse = IsMouseVK(remap_source_vk);
					remap_dest_is_mouse = IsMouseVK(remap_dest_vk);
					remap_keybd_to_mouse = !remap_source_is_mouse && remap_dest_is_mouse;
					sntprintf(remap_source, _countof(remap_source), _T("%s%s%s")
						, remap_source_is_combo ? _T("") : _T("*") // v1.1.27.01: Omit * when the remap source is a custom combo.
						, _tcslen(cp1) == 1 && IsCharUpper(*cp1) ? _T("+") : _T("")  // Allow A::b to be different than a::b.
						, static_cast<LPTSTR>(buf)); // Include any modifiers too, e.g. ^b::c.
					if (*cp == '"' || *cp == g_EscapeChar) // Need to escape these.
					{
						*remap_dest = g_EscapeChar;
						remap_dest[1] = *cp;
						remap_dest[2] = '\0';
					}
					else
						tcslcpy(remap_dest, cp, _countof(remap_dest));  // But exclude modifiers here; they're wanted separately.
					tcslcpy(remap_dest_modifiers, hotkey_flag, _countof(remap_dest_modifiers));
					if (cp - hotkey_flag < _countof(remap_dest_modifiers)) // Avoid reading beyond the end.
						remap_dest_modifiers[cp - hotkey_flag] = '\0';   // Terminate at the proper end of the modifier string.
					
					if (!*remap_dest_modifiers	
						&& (remap_dest_vk == VK_PAUSE)	
						&& (!_tcsicmp(remap_dest, _T("Pause")))) // Specifically "Pause", not "vk13".
					{
						// In the unlikely event that the dest key has the same name as a command, disqualify it
						// from being a remap (as documented). 
						// v1.0.40.05: If the destination key has any modifiers,
						// it is unambiguously a key name rather than a command.
					}
					else
					{
						// It is a remapping. Create one "down" and one "up" hotkey,
						// eg, "x::y" yields,
						// *x::
						// {
						// SetKeyDelay(-1), Send("{Blind}{y DownR}")
						// }
						// *x up::
						// {
						// SetKeyDelay(-1), Send("{Blind}{y Up}")
						// }
						// Using one line to facilitate code.

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

						if (mLastHotFunc)		
							// Checking this to disallow stacking, eg
							// x::
							// y::z
							// which would cause x:: to just do the "down"
							// part of y::z.
							return ScriptError(ERR_HOTKEY_MISSING_BRACE);

						auto make_remap_hotkey = [&](LPTSTR aKey)
						{
							if (!CreateHotFunc())
									return FAIL;
							hk = Hotkey::FindHotkeyByTrueNature(aKey, no_suppress, hook_is_mandatory);
							if (hk)
							{
								if (!hk->AddVariant(mLastHotFunc, no_suppress))
									return FAIL;
							}
							else if (!Hotkey::AddHotkey(mLastHotFunc, HK_NORMAL, aKey, no_suppress))
								return FAIL;
							return OK;
						};
						// Start with the "down" hotkey:
						if (!make_remap_hotkey(remap_source)) 
							return FAIL;
						
						TCHAR remap_buf[LINE_SIZE];
						cp = remap_buf;
						cp += _stprintf(cp
							, _T("%s(-1),") // Does NOT need to be "-1, -1" for SetKeyDelay (see below).
							, remap_dest_is_mouse ? _T("SetMouseDelay") : _T("SetKeyDelay") // Using the full function names vs. Set%sDelay might help size due to string pooling.
						);
						// It seems unnecessary to set press-duration to -1 even though the auto-exec section might
						// have set it to something higher than -1 because:
						// 1) Press-duration doesn't apply to normal remappings since they use down-only and up-only events.
						// 2) Although it does apply to remappings such as a::B and a::^b (due to press-duration being
						//    applied after a change to modifier state), those remappings are fairly rare and supporting
						//    a non-negative-one press-duration (almost always 0) probably adds a degree of flexibility
						//    that may be desirable to keep.
						// 3) SendInput may become the predominant SendMode, so press-duration won't often be in effect anyway.
						// 4) It has been documented that remappings use the auto-execute section's press-duration.
						// The primary reason for adding Key/MouseDelay -1 is to minimize the chance that a one of
						// these hotkey threads will get buried under some other thread such as a timer, which
						// would disrupt the remapping if #MaxThreadsPerHotkey is at its default of 1.
						if (remap_keybd_to_mouse)
						{
							// Since source is keybd and dest is mouse, prevent keyboard auto-repeat from auto-repeating
							// the mouse button (since that would be undesirable 90% of the time).  This is done
							// by inserting a single extra IF-statement above the Send that produces the down-event:
							cp += _stprintf(cp, _T("!GetKeyState(\"%s\")&&"), remap_dest); // Should be no risk of buffer overflow due to prior validation.
						}
						// Otherwise, remap_keybd_to_mouse==false.

						TCHAR blind_mods[5], * next_blind_mod = blind_mods, * this_mod, * found_mod;
						for (this_mod = _T("!#^+"); *this_mod; ++this_mod)
						{
							found_mod = _tcschr(remap_source, *this_mod);
							if (found_mod && found_mod[1]) // Exclude the last char for !:: and similar.
								*next_blind_mod++ = *this_mod;
						}
						*next_blind_mod = '\0';
						LPTSTR extra_event = _T(""); // Set default.
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
						cp += _stprintf(cp
							, _T("Send(\"{Blind%s}%s%s{%s DownR}\")") // DownR vs. Down. See Send's DownR handler for details.
							, blind_mods, extra_event, remap_dest_modifiers, remap_dest);

						auto define_remap_func = [&]()
						{
							if (!AddLine(ACT_BLOCK_BEGIN)
								|| !ParseAndAddLine(remap_buf)
								|| !AddLine(ACT_BLOCK_END))
								return FAIL;
							return OK;
						};
						if (!define_remap_func()) // the "down" function.
							return FAIL;
						//
						// "Down" is finished, proceed with "Up":
						//
						_stprintf(remap_buf
							, _T("%s up") // Key-up hotkey, e.g. *LButton up::
							, remap_source);
						if (!make_remap_hotkey(remap_buf)) 
							return FAIL;
						_stprintf(remap_buf
							, _T("%s(-1),")
							_T("Send(\"{Blind}{%s Up}\")\n") // Unlike the down-event above, remap_dest_modifiers is not included for the up-event; e.g. ^{b up} is inappropriate.
							, remap_dest_is_mouse ? _T("SetMouseDelay") : _T("SetKeyDelay") // Using the full function names vs. Set%sDelay might help size due to string pooling.
							, remap_dest
						);
						if (!define_remap_func()) // define the "up" function.
							return FAIL;
						goto continue_main_loop;
					}
					// Since above didn't goto this is not a remap after all:
				}
			}
			
			auto set_last_hotfunc = [&]()
			{
				if (!mLastHotFunc)
					return CreateHotFunc();
				else
					return mLastHotFunc;
			};
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

				if (hotstring_execute && (!*hotkey_flag || hotkey_uses_otb))
					// Do not allow execute option with blank line or OTB.
					// Without this check, this
					// :X:x::
					// {
					// }
					// would execute the block. But X is supposed to mean "execute this line".
					return ScriptError(ERR_EXPECTED_ACTION);

				if (hotkey_uses_otb)
				{
					// Never use otb if text or raw mode is in effect for this hotstring.
					// Either explicitly or via #hotstring.
					bool uses_text_or_raw_mode = g_HSSendRaw;
					for (LPTSTR cp = hotstring_options; cp < hotstring_start; ++cp)
						switch (ctoupper(*cp))
						{
						case 'T':
						case 'R':
							uses_text_or_raw_mode = cp[1] != '0';
						}
					if (uses_text_or_raw_mode)
						hotkey_uses_otb = false;
				}

				// The hotstring never uses otb if it uses X or T options (either explicitly or via #hotstring).
				if ( (!*hotkey_flag || hotkey_uses_otb) || hotstring_execute)
				{
					// It is not auto-replace
					if (!set_last_hotfunc())
						return FAIL;
				} 
				else if (mLastHotFunc)
				{
					// It is autoreplace but an earlier hotkey or hotstring
					// is "stacked" above, treat it as and error as it doesn't
					// make sense. Otherwise one could write something like:
					/*
					::abc:: 
					::def::text
					x::action
					 which would work as,
					::def::text
					::abc::
					x::action
					*/
					// Note that if it is ":X:def::action" instead, we do not end up here and
					// "::abc::" will also trigger "action".
					mCombinedLineNumber--;	// It must be the previous line.
					return ScriptError(ERR_HOTKEY_MISSING_BRACE);
				}
				
				if (!Hotstring::AddHotstring(buf, mLastHotFunc, hotstring_options
					, hotstring_start, hotstring_execute || hotkey_uses_otb ? _T("") : hotkey_flag, has_continuation_section))
					return FAIL;
				if (!mLastHotFunc)
					goto continue_main_loop;
			}
			else // It's a hotkey vs. hotstring.
			{
				hook_action = Hotkey::ConvertAltTab(hotkey_flag, false);
				if (hk = Hotkey::FindHotkeyByTrueNature(buf, no_suppress, hook_is_mandatory)) // Parent hotkey found.  Add a child/variant hotkey for it.
				{
					if (hook_action) // no_suppress has always been ignored for these types (alt-tab hotkeys).
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
						if (hk->FindVariant()) // See if there's already a variant matching the current criteria (no_suppress does not make variants distinct form each other because it would require firing two hotkey IDs in response to pressing one hotkey, which currently isn't in the design).
						{
							mCurrLine = NULL;  // Prevents showing unhelpful vicinity lines.
							return ScriptError(_T("Duplicate hotkey."), buf);
						}
						if (!set_last_hotfunc())
							return FAIL;
						if (!hk->AddVariant(mLastHotFunc, no_suppress))
							return ScriptError(ERR_OUTOFMEM, buf);
						if (hook_is_mandatory || g_ForceKeybdHook)
						{
							// Require the hook for all variants of this hotkey if any variant requires it.
							// This seems more intuitive than the old behaviour, which required $ or #UseHook
							// to be used on the *first* variant, even though it affected all variants.
							hk->mKeybdHookMandatory = true;
						}
					}
				}
				else // No parent hotkey yet, so create it.
				{
					if (hook_action != HK_NORMAL && mLastHotFunc)
						// A hotkey is stacked above, eg,
						// x::
						// y & z::altTab
						// Not supported.
						return ScriptError(ERR_HOTKEY_MISSING_BRACE);
					
					if (hook_action == HK_NORMAL 
						&& !set_last_hotfunc())
						return FAIL;
					
					hk = Hotkey::AddHotkey(mLastHotFunc, hook_action, buf, no_suppress);
					if (!hk)
					{
						if (hotkey_validity != CONDITION_TRUE)
							return FAIL; // It already displayed the error.
						// This hotkey uses a single-character key name, which could be valid on some other
						// keyboard layout.  Allow the script to start, but warn the user about the problem.
						// Note that this hotkey's label is still valid even though the hotkey wasn't created.
#ifndef AUTOHOTKEYSC
						if (!mValidateThenExit) // Current keyboard layout is not relevant in /validate mode.
#endif
						{
							TCHAR msg_text[128];
							sntprintf(msg_text, _countof(msg_text), _T("Note: The hotkey %s will not be active because it does not exist in the current keyboard layout."), static_cast<LPTSTR>(buf));
							MsgBox(msg_text);
						}
					}
				}
			}
			if (*hotkey_flag) // This hotkey's/hotstring's action is on the same line as its label.
			{
				if (hotkey_uses_otb) {
					// x::{
					//	; code
					// }
					if (!AddLine(ACT_BLOCK_BEGIN))
						return FAIL;
				}
				// Don't add AltTab or similar as a line, since it has no meaning as a script command.
				else if (hotstring_start ? hotstring_execute : !hook_action)
				{
					// Eg, ":X:abc::msgbox" or "x::msgbox", 
					// x::
					// {
					// msgbox
					// }
					ASSERT(mLastHotFunc && mLastHotFunc == g->CurrentFunc);
					// Remove the hotkey from buf.
					buf_length -= hotkey_flag - buf;
					tmemmove(buf, hotkey_flag, buf_length);
					buf[buf_length] = '\0';
					// Before adding the line, apply expression line-continuation logic, which hasn't
					// been applied yet because hotkey labels can contain unbalanced ()[]{}:
					if (   !GetLineContExpr(fp, buf, next_buf, phys_line_number, has_continuation_section)
						|| !AddLine(ACT_BLOCK_BEGIN)			// Implicit start of function
						|| !ParseAndAddLine(buf)				// Function body - one line
						|| !AddLine(ACT_BLOCK_END))				// Implicit end of function
						return FAIL;
				}
			}
			goto continue_main_loop; // In lieu of "continue", for performance.
		} // if (hotkey_flag && hotkey_flag > buf)

		// Otherwise, not a hotkey or hotstring.  Check if it's a generic, non-hotkey label:
		if (buf[buf_length - 1] == ':' // Labels must end in a colon (buf was previously rtrimmed).
			&& (!mClassObjectCount || g->CurrentFunc)) // Not directly inside a class body or property definition (but inside a method is okay).
		{
			// Labels (except hotkeys) must contain no whitespace, delimiters, or escape-chars.
			// This is to avoid problems where a legitimate action-line ends in a colon,
			// such as "#Include c:", and for sanity (so labels can be readily recognized).
			// We allow hotkeys to violate this since they may contain commas, but they must
			// also be a valid combination of symbols and key names, so there is no ambiguity.
			// v2.0: Require label names to use the same set of characters as other identifiers.
			// Aside from consistency and ensuring readability, this might enable future changes
			// to the parser or new syntax.
			cp = find_identifier_end<LPTSTR>(buf);
			if ((cp - buf + 1) == buf_length && cp > buf)
			{
				buf[--buf_length] = '\0';  // Remove the trailing colon.
				if (!_tcsicmp(buf, _T("Default")) && mOpenBlock && mOpenBlock->mPrevLine // "Default:" case.
					&& mOpenBlock->mPrevLine->mActionType == ACT_SWITCH) // It's a normal label in any other case.
				{
					if (!AddLine(ACT_CASE))
						return FAIL;
				}
				else
				{
					if (mLastHotFunc) // It is a label in a "stack" of hotkeys.
						return ScriptError(ERR_HOTKEY_MISSING_BRACE);
					if (!AddLabel(buf, false))
						return FAIL;
				}
				goto continue_main_loop; // In lieu of "continue", for performance.
			}
		}
		// Since above didn't "goto", it's not a label.
		if (*buf == '#')
		{
			if (!_tcsnicmp(buf, _T("#HotIf"), 6) && IS_SPACE_OR_TAB(buf[6]))
			{
				// Allow an expression enclosed in ()/[]/{} to span multiple lines:
				if (!GetLineContExpr(fp, buf, next_buf, phys_line_number, has_continuation_section))
					return FAIL;
			}
			saved_line_number = mCombinedLineNumber; // Backup in case IsDirective() processes an include file, which would change mCombinedLineNumber's value.
			switch(IsDirective(buf)) // Note that it may alter the contents of buf
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
			if (mClassObjectCount && !g->CurrentFunc)
			{
				if (*buf == '{')
					return ScriptError(ERR_UNEXPECTED_OPEN_BRACE, buf);

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
			}
			else // Normal block begin/end.
			{
				if (!AddLine(*buf == '{' ? ACT_BLOCK_BEGIN : ACT_BLOCK_END))
					return FAIL;
			}
			// Allow the remainder of the line to be treated as a separate line:
			LPTSTR cp = omit_leading_whitespace(buf + 1);
			if (*cp)
			{
				buf_length -= (cp - buf);
				tmemmove(buf, cp, buf_length + 1);
				mCurrLine = NULL;  // To signify that we're in transition, trying to load a new line.
				goto process_completed_line; // Have the main loop process the contents of "buf" as though it came in from the script.
			}
			goto continue_main_loop; // It's just a naked "{" or "}", so no more processing needed for this line.
		}

		// Handle this first so that GetLineContExpr() doesn't need to detect it for OTB exclusion:
		if (LPTSTR class_name = IsClassDefinition(buf))
		{
			if (g->CurrentFunc)
				return ScriptError(_T("Functions cannot contain classes."), buf);
			if (!ClassHasOpenBrace(buf, buf_length, next_buf, next_buf_length))
				return ScriptError(ERR_MISSING_OPEN_BRACE, buf);
			if (!DefineClass(class_name))
				return FAIL;
			goto continue_main_loop;
		}

		// Aside from goto/break/continue, anything not already handled above is either an expression
		// or something with similar lexical requirements (i.e. balanced parentheses/brackets/braces).
		// The following call allows any expression enclosed in ()/[]/{} to span multiple lines:
		if (!GetLineContExpr(fp, buf, next_buf, phys_line_number, has_continuation_section))
			return FAIL;

		if (mClassProperty && !g->CurrentFunc) // This is checked before IsFunction() to prevent method definitions inside a property.
		{
			if (!_tcsnicmp(buf, _T("Get"), 3) || !_tcsnicmp(buf, _T("Set"), 3))
			{
				LPTSTR cp = omit_leading_whitespace(buf + 3);
				if (!*cp && *next_buf == '{' // Brace on next line
					|| *cp == '{' && !cp[1] // Brace on this line
					|| *cp == '=' && cp[1] == '>') // => expr
				{
					LPTSTR dot = _tcschr(mClassPropertyDef, '.');
					dot[1] = *buf; // Replace the x in property.xet(params).
					if (!DefineClassPropertyXet(buf, cp))
						return FAIL;
					goto continue_main_loop;
				}
			}
			return ScriptError(ERR_INVALID_LINE_IN_PROPERTY_DEF, buf);
		}

		if (mClassObjectCount && !g->CurrentFunc) // Inside a class definition (and not inside a method).
		{
			LPTSTR id = buf;
			bool is_static = false;
			if (!_tcsnicmp(id, _T("Static"), 6) && IS_SPACE_OR_TAB(id[6]))
			{
				id = omit_leading_whitespace(id + 7);
				if (IS_IDENTIFIER_CHAR(*id))
					is_static = true;
				else
					id = buf;
			}
			if (IsFunctionDefinition(id, next_buf))
			{
				if (!DefineFunc(id, is_static))
					return FAIL;
				goto continue_main_loop;
			}
			for (cp = id; IS_IDENTIFIER_CHAR(*cp) || *cp == '.'; ++cp);
			if (cp > id) // i.e. buf begins with an identifier.
			{
				cp = omit_leading_whitespace(cp);
				if (*cp == ':' && cp[1] == '=') // This is an assignment.
				{
					if (!DefineClassVars(id, is_static))
						return FAIL;
					goto continue_main_loop;
				}
				if (!*cp || *cp == '[' || *cp == '{' || (*cp == '=' && cp[1] == '>')) // Property or invalid.
				{
					if (!DefineClassProperty(id, is_static, buf_has_brace))
						return FAIL;
					if (!buf_has_brace)
					{
						if (*next_buf != '{')
							return ScriptError(ERR_UNRECOGNIZED_ACTION, buf); // Vague message because user's intention is unknown.
						RemoveBufChar0(next_buf, next_buf_length);
					}
					goto continue_main_loop;
				}
			}
			// Anything not already handled above is not valid directly inside a class definition.
			return ScriptError(ERR_INVALID_LINE_IN_CLASS_DEF, buf);
		}
		else if (IsFunctionDefinition(buf, next_buf))
		{
			if (!DefineFunc(buf))
				return FAIL;
			goto continue_main_loop;
		}

		// Parse the command, assignment or expression, including any same-line open brace or sub-action
		// for ELSE, TRY, CATCH or FINALLY.  Unlike braces at the start of a line (processed above), this
		// does not allow directives or labels to the right of the command.
		if (!ParseAndAddLine(buf))
			return FAIL;

continue_main_loop: // This method is used in lieu of "continue" for performance and code size reduction.
		// Since above didn't "continue", resume loading script line by line:
		swap(buf.p, next_buf.p);
		swap(buf.size, next_buf.size);
		buf_length = next_buf_length;
		// The lines above alternate buffers (toggles next_buf to be the unused buffer), which helps
		// performance because it avoids memcpy from buf2 to buf1.
	} // for each whole/constructed line.

	// The following checks detect errors where a "}" was missed or there's an If/Else/etc.
	// or hotkey at the end of the file with no corresponding body.  Although it's possible
	// that another file might follow this one and fill in the missing part, it's much more
	// likely to be an error.  blocks_previously_open is checked rather than source_file_index
	// so that errors in auto-includes are detected.

	if (!blocks_previously_open && (mOpenBlock || mClassObjectCount))
	{
		if (mOpenBlock)
			return mOpenBlock->LineError(ERR_MISSING_CLOSE_BRACE);
		// A class definition has not been closed with "}".  Previously this was detected by adding
		// the open and close braces as lines, but this way is simpler and has less overhead.
		// The downside is that the line number won't be shown; however, the class name will.
		// Seems okay not to show mClassProperty->mName since the class is missing "}" as well.
		return ScriptError(ERR_MISSING_CLOSE_BRACE, mClassName);
	}

	if (mPendingParentLine) // If, Else, Loop, etc. without a block or action.
		return mPendingParentLine->LineUnexpectedError();

	if (mLastHotFunc)
		// A hotkey at the end of the file, with no function body.  Checking for this here
		// ensures mJumpToLine is always non-null for later stages.
		return ScriptError(ERR_HOTKEY_MISSING_BRACE);

	++mCombinedLineNumber; // L40: Put the implicit ACT_EXIT on the line after the last physical line (for the debugger).
	return OK;
}



bool Script::EndsWithOperator(LPTSTR aBuf, LPTSTR aBuf_marker)
// Returns true if aBuf_marker is the end of an operator, excluding ++ and --.
{
	LPTSTR cp = aBuf_marker; // Caller has omitted trailing whitespace.
	if (_tcschr(EXPR_OPERATOR_SYMBOLS, *cp) // It's a binary operator or ++ or --.
		&& !((*cp == '+' || *cp == '-') && cp > aBuf && cp[-1] == *cp)) // Not ++ or --.
		return true;
	LPTSTR word;
	for (word = cp; word > aBuf && IS_IDENTIFIER_CHAR(word[-1]); --word);
	if (word > aBuf && word[-1] == '.')
		return false; // Reserved words are permitted as property names, so `a.as` isn't an operator.
	switch (ConvertWordOperator(word, cp - word + 1))
	{
	case SYM_OR:
	case SYM_AND:
	case SYM_IS:
	case SYM_LOWNOT:
	case SYM_RESERVED_OPERATOR:
		return true;
	}
	return false;
}



ResultType Script::LineBuffer::EnsureCapacity(size_t aLength)
{
	aLength += RESERVED_SPACE;
	if (size >= aLength)
		return OK;
	// See comments in Expand() regarding buffer growth.
	size_t newsize = size ? size : INITIAL_SIZE;
	while (newsize < aLength)
		newsize *= 2;
	return Realloc(newsize);
}

ResultType Script::LineBuffer::Expand()
{
	// The buffer typically needs to grow incrementally while joining lines in a continuation section.
	// Expanding in large increments avoids multiple reallocs in most files.  Expanding exponentially
	// scales better in theory, though some testing showed it to have virtually no impact on load time
	// even with very long lines.  Memory usage isn't a concern since the buffer will be freed after
	// the script file is read.
	return Realloc(size ? size * 2 : INITIAL_SIZE);
}

ResultType Script::LineBuffer::Realloc(size_t aNewSize)
{
	LPTSTR newp = (LPTSTR)realloc(p, sizeof(TCHAR) * aNewSize);
	if (!newp)
		return FAIL;
	p = newp;
	size = aNewSize;
	return OK;
}



ResultType Script::GetLineContExpr(TextStream *fp, LineBuffer &buf, LineBuffer &next_buf
	, LineNumberType &phys_line_number, bool &has_continuation_section)
// Determine whether buf has an unclosed (/[/{, and if so, complete the expression by
// appending subsequent lines until the number of opening and closing symbols balances out.
// Continuation lines and sections are assumed to have already been handled (and appended to
// buf) by a previous call to GetLineContinuation(), but **only up to the unbalanced line**;
// any subsequent contination is handled by this function.
{
	ActionTypeType action_type = ACT_INVALID; // Set default.
	ptrdiff_t action_end_pos = 0;

	size_t &buf_length = buf.length;
	size_t &next_buf_length = next_buf.length;

	TCHAR expect[MAX_BALANCEEXPR_DEPTH], open_quote = 0;
	int balance = BalanceExpr(buf, 0, expect);
	if (balance < 0 // Invalid.
		|| !buf_length // Blank (shouldn't happen as buf must already qualify as expression-like).
		|| next_buf_length == -1) // End of file.
		return balance == 0 ? OK : BalanceExprError(balance, expect, buf);
	
	// Perform rough checking for this line's action type.
	for (LPTSTR action_start = buf; ; )
	{
		LPTSTR action_end = find_identifier_end(action_start);
		if (action_end > action_start)
		{
			TCHAR orig_char = *action_end;
			// This relies on names of control flow statements being invalid for use as var/func names:
			if (IS_SPACE_OR_TAB(orig_char) || orig_char == '(' || orig_char == '{') // '{' supports "else{".
			{
				*action_end = '\0';
				action_type = ConvertActionType(action_start);
				*action_end = orig_char;

				if (action_type == ACT_ELSE || action_type == ACT_TRY || action_type == ACT_FINALLY)
				{
					action_start = omit_leading_whitespace(action_end);
					if (*action_start == '{')
						return OK; // This is unconditionally a block-begin, not an expression.
					continue; // Parse any same-line action instead.
				}

				if (mClassObjectCount && !g->CurrentFunc // In a class body.
					&& (action_end - action_start == 6) && !_tcsnicmp(action_start, _T("Static"), 6)) // Ignore "Static" modifier.
				{
					action_start = omit_leading_whitespace(action_end);
					continue;
				}
			}
		}
		action_end_pos = action_end - buf; // Use position since action_end will be invalidated if buf is reallocated.
		break;
	}

	do
	{
		if (balance == 0 && buf[buf_length - 1] == ':')
		{
			if (action_type == ACT_CASE && FindExprDelim(buf, ':') == buf_length - 1)
			{
				// This is the colon terminating a case statement.
				return OK;
			}
			//else this colon qualifies for line continuation.
		}
		else if (balance == 0 && !EndsWithOperator(buf, buf + (buf_length - 1)) && !IsSOLContExpr(next_buf))
		{
			// There's no continuation by enclosure and no continuation operator, so we're done.
			return OK;
		}
		// Before appending each line, check whether the last line ended with OTB '{'.
		// It can't be OTB if balance > 1 since that would mean another unclosed (/[/{.
		// balance == 1 implies buf_length >= 1, but below requires buf_length >= 2.
		// The shortest valid OTB is for the property "p{".  Finding '{' usually implies
		// buf_length >= 2 because '{' at the start of buf is usually handled by the
		// caller, but that isn't always the case (e.g. for one-line hotkeys).
		if (balance == 1 && buf[buf_length - 1] == '{' && buf_length >= 2)
		{
			// Some common OTB constructs:
			//   myfn() {
			//   if (cond) {
			//   while (cond) {
			//   loop {
			//   if cond {
			// For the first few cases, *cp == ')' is sufficient.  There is no need to verify
			// that this is a function definition because ") {" is not valid in an expression
			// (it is reserved for future use with anonymous functions).  Similarly, "] {" is
			// either a property definition or invalid.
			// For other cases, checking the action_type is the only way to resolve the ambiguity
			// between "loop {" and "return {".  Since valid OTB can't be preceded by an operator
			// such as ":= {", also check that case to improve flexibility.
			LPTSTR cp = omit_trailing_whitespace(buf, buf + (buf_length - 2));
			if (   *cp == ')' // Function/method definition or reserved.
				|| *cp == ']' // Property definition or reserved.
				|| ACT_IS_LINE_PARENT(action_type) && !EndsWithOperator(buf, cp)
				|| mClassObjectCount && !g->CurrentFunc && (cp - buf) < action_end_pos) // "Property {" (get/set was already handled by caller).
				return OK;
		}
		if (next_buf_length) // Skip empty/comment lines.
		{
			if (!buf.EnsureCapacity(next_buf_length + buf_length + 1))
				return ScriptError(ERR_OUTOFMEM);
			balance = BalanceExpr(next_buf, balance, expect, &open_quote); // Adjust balance based on what we're about to append.
			buf[buf_length++] = ' '; // To ensure two distinct tokens aren't joined together.  ' ' vs. '\n' because DefineFunc() currently doesn't permit '\n'.
			tmemcpy(buf + buf_length, next_buf, next_buf_length); // Append next_buf to this line.
			buf_length += next_buf_length;
			buf[buf_length] = '\0';
		}
		auto prev_length = buf_length; // Store length vs. pointer since buf may be reallocated below.
		// This serves to get the next non-blank, non-comment line into next_buf but also handles any
		// continuation sections which follow the line just appended:
		if (!GetLineContinuation(fp, buf, next_buf, phys_line_number, has_continuation_section))
			return FAIL;
		if (buf_length > prev_length && balance >= 0)
			balance = BalanceExpr(buf + prev_length, balance, expect, &open_quote); // Adjust balance based on what was appended.
		if (open_quote)
		{
			// Unterminated string (continuation expressions should not automatically permit strings to span lines).
			expect[0] = open_quote;
			expect[1] = 0;
			balance = -1;
		}
	} // do
	while (balance >= 0 && next_buf_length != -1);
	if (balance != 0)
	{
		// buf might include some lines that the author did not intend to be merged, so report
		// the error immediately.  Leaving it to later might obscure it with some other problem,
		// such as control flow statements being interpreted as invalid variable references.
		return BalanceExprError(balance, expect, buf);
	}
	return OK;
}



bool Script::IsSOLContExpr(LineBuffer &next_buf)
{
	LPTSTR cp, hotkey_flag;

	switch(ctoupper(*next_buf)) // Above has ensured *next_buf != '\0' (toupper might have problems with '\0').
	{
	case 'A': // AND, AS (future use)
	case 'O': // OR
	case 'I': // IS, IN, (unwanted) ISSET
	case 'C': // CONTAINS (future use)
		// See comments in the default section further below.
		cp = find_identifier_end<LPTSTR>(next_buf);
		// Although (x)and(y) is technically valid, it's quite unusual.  The space or tab requirement is kept
		// as the simplest way to allow method definitions to use these as names (when called, the leading dot
		// ensures there is no ambiguity).  Note that checking if we're inside a class definition is not
		// sufficient because multi-line expressions are valid there too (i.e. for var initializers).
		// This also rules out valid double-derefs such as and%suffix% := 1.
		if (IS_SPACE_OR_TAB(*cp))
		{
			// Unlike in v1, there's no check for an operator after AND/OR (such as AND := 1) because they
			// should never be used as variable names.
			auto op = ConvertWordOperator(next_buf, cp - next_buf);
			if (op && op != SYM_ISSET) // IsSet is excluded for property definitions like IsSet => true.
				return true;
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
		// The logic here isn't smart enough to differentiate between a leading - that's meant as a
		// continuation character and one that isn't. Even if it were, it would still be ambiguous in
		// some cases because the author's intent isn't known; for example, the leading minus sign on
		// the second line below is ambiguous:
		//    x := y 
		//    -z ? a:=1 : func() 
		if ((*next_buf == ':' || *next_buf == '+' || *next_buf == '-') && next_buf[1] == *next_buf // See above.
			|| !_tcschr(CONTINUATION_LINE_SYMBOLS, *next_buf)) // Line doesn't start with a continuation char.
			break;
		// Some of the above checks must be done before the next ones.
		if (   !(hotkey_flag = _tcsstr(next_buf, HOTKEY_FLAG))   ) // Without any "::", it can't be a hotkey or hotstring.
			return true;
		if (*next_buf == ':') // First char is ':', so it's more likely a hotstring than a hotkey.
		{
			// Remember that hotstrings can contain what *appear* to be quoted literal strings,
			// so detecting whether a "::" is in a quoted/literal string in this case would
			// be more complicated.  That's one reason this other method is used.
			bool hotstring_options_all_valid = true;
			for (cp = next_buf + 1; *cp && *cp != ':'; ++cp)
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
		// because above already discovered that it contains "::" somewhere.  The v1 policy was
		// to do some rough checks for comma hotkeys and quote marks.  For v2, we use a much
		// clearer rule: if it's valid hotkey syntax (even with invalid key names), it's not
		// continuation.  There is at least one pathological case that could be an expression,
		// but it's much more likely to be intended as a hotkey:
		//  MsgBox
		//  +!'::'  ; 0
		*hotkey_flag = '\0';
		bool valid_hotkey = Hotkey::TextInterpret(next_buf, NULL, true);
		*hotkey_flag = *HOTKEY_FLAG;
		if (!valid_hotkey)
			return true; // It's not valid hotkey syntax, so treat it as continuation even if it's ultimately a syntax error.
	} // switch(toupper(*next_buf))
	return false;
}



ResultType Script::BalanceExprError(int aBalance, TCHAR aExpect[], LPTSTR aLineText)
{
	TCHAR expected, found;
	if (aBalance < 0)
	{
		expected = aExpect[0];
		found = aExpect[1];
	}
	else
	{
		expected = aBalance < MAX_BALANCEEXPR_DEPTH ? aExpect[aBalance - 1] : 0;
		found = 0;
	}
	TCHAR msgbuf[40];
	LPTSTR msgfmt;
	if (expected && found)
		msgfmt = _T("Missing \"%c\" before \"%c\"");
	else if (found)
		msgfmt = _T("Unexpected \"%c\""), expected = found;
	else if (expected)
		msgfmt = _T("Missing \"%c\"");
	else
		msgfmt = _T("Missing symbol"); // Rare case (expression too deep to keep track of what's missing).
	sntprintf(msgbuf, _countof(msgbuf), msgfmt, expected, found);
	return ScriptError(msgbuf, aLineText);
}



ResultType Script::GetLineContinuation(TextStream *fp, LineBuffer &buf, LineBuffer &next_buf
	, LineNumberType &phys_line_number, bool &has_continuation_section)
{
	bool do_rtrim, literal_escapes, literal_quotes;
	#define CONTINUATION_SECTION_WITHOUT_COMMENTS 1 // MUST BE 1 because it's the default set by anything that's boolean-true.
	#define CONTINUATION_SECTION_WITH_COMMENTS    2 // Zero means "not in a continuation section".
	int in_continuation_section, indent_level;
	ToggleValueType do_ltrim;
	TCHAR indent_char, suffix[16];
	size_t suffix_length;

	LPTSTR next_option, option_end, cp;
	TCHAR orig_char, quote_char;
	bool in_comment_section;
	int quote_pos;
	int continuation_line_count;

	auto &buf_length = buf.length;
	auto &next_buf_length = next_buf.length;

	// Read in the next line (if that next line is the start of a continuation section, append
	// it to the line currently being processed:
	for (has_continuation_section = in_comment_section = false, in_continuation_section = 0;;)
	{
		// This increment relies on the fact that this loop always has at least one iteration:
		++phys_line_number; // Tracks phys. line number in *this* file (independent of any recursion caused by #Include).
		next_buf_length = GetLine(next_buf, in_continuation_section, in_comment_section, fp);
		if (!in_continuation_section)
		{
			// v2: The comment-end is allowed at the end of the line (vs. just the start) to reduce
			// confusion for users expecting C-like behaviour, but unlike v1.1, the end flag is not
			// allowed outside of comments, since allowing and ignoring */ at the end of any line
			// seems to have risk of ambiguity.
			// If this policy is changed to ignore an orphan */, remember to allow */:: as a hotkey.

			// Check for /* first in case */ appears on the same line.  There's no need to check
			// in_comment_section since this has the same effect either way (and it's usually false).
			if (!_tcsncmp(next_buf, _T("/*"), 2))
			{
				in_comment_section = true;
				// But don't "continue;" since there might be a */ on this same line.
			}

			if (in_comment_section)
			{
				if (next_buf_length == -1) // Compare directly to -1 since length is unsigned.
					break; // By design, it's not an error.  This allows "/*" to be used to comment out the bottommost portion of the script without needing a matching "*/".

				if (!_tcsncmp(next_buf, _T("*/"), 2))
				{
					in_comment_section = false;
					next_buf_length -= 2; // Adjust for removal of */ from the beginning of the string.
					tmemmove(next_buf, next_buf + 2, next_buf_length + 1);  // +1 to include the string terminator.
					next_buf_length = ltrim(next_buf, next_buf_length); // Get rid of any whitespace that was between the comment-end and remaining text.
					if (!*next_buf) // The rest of the line is empty, so it was just a naked comment-end.
						continue;
				}
				else
				{
					// This entire line is part of the comment, but there might be */ at the end of the line.
					if (next_buf_length >= 2 && !_tcscmp(next_buf + (next_buf_length - 2), _T("*/")))
						in_comment_section = false;
					continue;
				}
			}

			// v1.0.38.06: The following has been fixed to exclude "(:" and "(::".  These should be
			// labels/hotkeys, not the start of a continuation section.
			// v2: Unlike v1, a line that starts with '(' but that ends with ':' is not a valid label,
			// so should be treated as a continuation section, for cases like "(Join:".
			if (   !(in_continuation_section = (next_buf_length != -1 && *next_buf == '(' && next_buf[1] != ':'))   ) // Compare directly to -1 since length is unsigned.
			{
				if (!next_buf_length)
					// It is permitted to have blank lines and comment lines in between the line above
					// and any continuation section/line that might come after the end of the
					// comment/blank lines:
					continue;
				// Since there is no continuation section and this line isn't blank, no further searching
				// is needed.  This also covers the case of EOF (next_buf_length == -1).
				break;
			} // if (!in_continuation_section)

			// OTHERWISE in_continuation_section != 0, so the above has found the first line of a new
			// continuation section.
			continuation_line_count = 0; // Reset for this new section.
			// Parse options.  First set the defaults, which can be individually overridden
			// by any options actually present.  RTrim defaults to ON for two reasons:
			// 1) Whitespace often winds up at the end of a lines in a text editor by accident.  In addition,
			//    whitespace at the end of any consolidated/merged line will be rtrim'd anyway, since that's
			//    how command parsing works.
			// 2) Copy & paste from the forum and perhaps other web sites leaves a space at the end of each
			//    line.  Although this behavior is probably site/browser-specific, it's a consideration.
			do_ltrim = NEUTRAL; // Start off at neutral/smart-trim.
			do_rtrim = true; // Seems best to rtrim even if this line is a hotstring, since it is very rare that trailing spaces and tabs would ever be desirable.
			// For hotstrings (which could be detected via *buf==':'), it seems best not to default the
			// escape character (`) to be literal because the ability to have `t `r and `n inside the
			// hotstring continuation section seems more useful/common than the ability to use the
			// accent character by itself literally (which seems quite rare in most languages).
			literal_escapes = false;
			//literal_quotes is set below based on the contents of buf.
			// The default is linefeed because:
			// 1) It's the best choice for hotstrings, for which the line continuation mechanism is well suited.
			// 2) It's good for FileAppend.
			// 3) Minor: Saves memory in large sections by being only one character instead of two.
			suffix[0] = '\n';
			suffix[1] = '\0';
			suffix_length = 1;
			LPTSTR invalid_option = nullptr;
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
					ConvertEscapeSequences(suffix, NULL);
					suffix_length = _tcslen(suffix);
				}
				else if (!_tcsnicmp(next_option, _T("LTrim"), 5))
					do_ltrim = (next_option[5] == '0') ? TOGGLED_OFF : TOGGLED_ON;  // i.e. Only an explicit zero will turn it off.
				else if (!_tcsnicmp(next_option, _T("RTrim"), 5))
					do_rtrim = (next_option[5] != '0');
				else if (!_tcsnicmp(next_option, _T("Comments"), option_end - next_option)) // Any prefix is supported, but the documented options are: C, Com, Comment, Comments
					in_continuation_section = CONTINUATION_SECTION_WITH_COMMENTS; // Override the default, which is boolean true (i.e. 1).
				else if (*next_option == g_EscapeChar && !next_option[1])
					literal_escapes = true;
				else
				{
					// If ')' is present in any as yet unrecognized option, treat this line as an expression and not
					// the start of a continuation section.  This allows expressions to start the line with '(' to
					// control order of precedence, like (x.y)() or (x=y) && z().  Do the same for '(', for symmetry
					// and to allow multi-line expressions starting with (( and (MyFunc(.
					for (cp = next_option; *cp; ++cp)
						if (*cp == '(' || *cp == ')')
						{
							in_continuation_section = 0;
							*option_end = orig_char; // Undo the temporary termination.
							return OK;
						}
					// All other characters are reserved for possible future options or to detect errors.  One reason
					// not to assume any invalid continuation section was intended to be an expression is that something
					// like the following would result in a 'Missing "' error instead of reporting the invalid option:
					//  sometext := "
					//  (Jion`r`n
					//  ...
					// Mark the first invalid option, but don't report it until any other options are checked for ().
					if (!invalid_option)
						invalid_option = next_option;
				}

				// If the item was not handled by the above, ignore it because it is unknown.
				*option_end = orig_char; // Undo the temporary termination.
			} // for() each item in option list

			if (invalid_option)
				return ScriptError(ERR_INVALID_OPTION, invalid_option);

			// This section determines the default value for literal_quotes based on whether this
			// continuation section appears to start inside a quoted string.  This is done because
			// unconditionally defaulting to true would make it possible to start a quoted string
			// but impossible to end it within the same section, which seems counter-intuitive.
			// It's okay if buf contains something other than an expression, because in that case
			// it won't matter whether the quote marks are escaped or not.
			if (!has_continuation_section)
			{
				quote_pos = 0; // Init.
				quote_char = 0;
			}
			for (;;)
			{
				if (quote_char) // quote_marker is within a string.
				{
					quote_pos += FindTextDelim(buf + quote_pos, quote_char);
					if (!buf[quote_pos]) // No end quote yet.
						break;
					++quote_pos; // Skip the end quote.
				}
				for ( ; buf[quote_pos] && buf[quote_pos] != '"' && buf[quote_pos] != '\''; ++quote_pos);
				if (  !(quote_char = buf[quote_pos])  )
					break;
				++quote_pos; // Continue scanning after the start quote.
			}
			literal_quotes = quote_char != 0; // true if this section starts inside a quoted string.
			// quote_pos and quote_char are retained between iterations of the outer loop to ensure
			// correct detection when there are multiple continuation sections, and to avoid re-scanning
			// the entire buf for each new section.

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
			next_buf_length = rtrim(next_buf, next_buf_length); // Done because GetLine() wouldn't have done it due to have told it we're in a continuation section.
			// Anything that lies to the right of the close-parenthesis gets appended verbatim, with
			// no trimming (for flexibility) and no options-driven translation:
			cp = next_buf + 1;  // Use temp var cp to avoid a memcpy.
			--next_buf_length;  // This is now the length of cp, not next_buf.
		}
		else
		{
			// To support continuation sections following a naked action name such as "return", add a space.
			// Checking for an action name at the start of buf would be insufficient due to try/else/finally,
			// hotkeys with same-line action, etc. so perform only very basic checking.  This probably also
			// benefits other expressions, such as if buf ends with a variable name.  The quote_char check
			// prevents unwanted spaces in quoted strings at the expense of possible inconsistency in some
			// vanishingly rare cases, such as:
			//  - Auto-replace hotstring with an odd number of quote marks on the first line, a word char
			//    at the end of the first line *and* a continuation section following it.
			//  - Other very unconventional cases, such as labels or directives which also meet the above
			//    criteria (it's very unlikely for continuation sections to be used with these at all).
			if (!continuation_line_count && !quote_char // First content line and we're not in a quoted string.
				&& buf_length && IS_IDENTIFIER_CHAR(buf[buf_length - 1]) // buf ends with a possible var/action name.
				&& buf_length + 1 < LINE_SIZE)
			{
				buf[buf_length++] = ' ';
				buf[buf_length] = '\0';
			}
			// The following are done in this block only because anything that comes after the closing
			// parenthesis (i.e. the block above) is exempt from translations and custom trimming.
			// This means that commas are always delimiters and percent signs are always deref symbols
			// in the previous block.
			if (do_rtrim)
				next_buf_length = rtrim(next_buf, next_buf_length);
			if (do_ltrim == NEUTRAL)
			{
				// Neither "LTrim" nor "LTrim0" was present in this section's options, so
				// trim the continuation section based on the indentation of the first line.
				if (!continuation_line_count)
				{
					// This is the first line.
					indent_char = *next_buf;
					if (IS_SPACE_OR_TAB(indent_char))
					{
						// For simplicity, require that only one type of indent char is used. Otherwise
						// we'd have to provide some way to set the width (in spaces) of a tab char.
						for (indent_level = 1; next_buf[indent_level] == indent_char; ++indent_level);
						// Let the section below actually remove the indentation on this and subsequent lines.
					}
					else
						indent_level = 0; // No trimming is to be done.
				}
				if (indent_level)
				{
					int i;
					for (i = 0; i < indent_level && next_buf[i] == indent_char; ++i);
					if (i == indent_level)
					{
						// LTrim exactly (indent_level) occurrences of (indent_char).
						tmemmove(next_buf, next_buf + i, next_buf_length - i + 1); // +1 for null terminator.
						next_buf_length -= i;
					}
					// Otherwise, the indentation on this line is inconsistent with the first line,
					// so just leave it as is.
				}
			}
			else if (do_ltrim == TOGGLED_ON)
				// Trim all leading whitespace.
				next_buf_length = ltrim(next_buf, next_buf_length);
			// Insert escape characters as needed for escape characters or quote marks to be interpreted
			// literally, as per continuation section options or detection of enclosing quote marks.
			// For simplicity, allow for worst-case expansion.  In most cases next_buf won't need to be
			// expanded, and if it is, that space may be reused for subsequent lines.  Using the mode of
			// StrReplace that allocates memory would require larger code and might increase the number
			// of buffer expansions needed.
			if ((literal_escapes || literal_quotes) && !next_buf.EnsureCapacity(next_buf_length * 2))
				return ScriptError(ERR_OUTOFMEM);
			if (literal_escapes) // literal_escapes must be done FIRST because otherwise it would also replace any accents added by other options.
				StrReplace(next_buf, _T("`"), _T("``"), SCS_SENSITIVE, UINT_MAX, -1, nullptr, &next_buf_length);
			if (literal_quotes)
			{
				StrReplace(next_buf, _T("'"), _T("`'"), SCS_SENSITIVE, UINT_MAX, -1, nullptr, &next_buf_length);
				StrReplace(next_buf, _T("\""), _T("`\""), SCS_SENSITIVE, UINT_MAX, -1, nullptr, &next_buf_length);
			}
			cp = next_buf;
		} // Handling of a normal line within a continuation section.

		// Must check the combined length only after anything that might have expanded the string above.
		if (!buf.EnsureCapacity(buf_length + next_buf_length + suffix_length))
			return ScriptError(ERR_OUTOFMEM);

		++continuation_line_count;
		// Append this continuation line onto the primary line.
		// The suffix for the previous line gets written immediately prior to writing this next line,
		// which allows the suffix to be omitted for the final line.  But if this is the first line,
		// No suffix is written because there is no previous line in the continuation section.
		// In addition, if cp!=next_buf, this is the special line whose text occurs to the right of the
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
	return OK;
}



size_t Script::GetLine(LineBuffer &aBuf, int aInContinuationSection, bool aInBlockComment, TextStream *ts)
{
	size_t aBuf_length = 0;
	for (;;)
	{
		auto aBuf_capacity = aBuf.Capacity();
		auto read_length = ts->ReadLine(aBuf + aBuf_length, (DWORD)(aBuf_capacity - aBuf_length));
		if (!read_length && !aBuf_length) // End of file was reached and there's no line text from a previous iteration.
		{
			aBuf[0] = '\0';
			return -1;
		}
		aBuf_length += read_length;
		if (aBuf[aBuf_length - 1] == '\n')
			--aBuf_length;
		if (aBuf_length < aBuf_capacity)
			break; // It read one complete line, either finding a line ending or reaching end of file.
		// This line is at least the size of aBuf, so expand aBuf to read more.
		if (!aBuf.Expand())
			return -1;
	}
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
			if (*cp == g_CommentChar)
			{
				*aBuf = '\0'; // Since this line is a comment, have the caller ignore it.
				return -2; // Callers tolerate -2 only when in a continuation section.  -2 indicates, "don't include this line at all, not even as a blank line to which the JOIN string (default "\n") will apply.
			}
			if (*cp == ')') // This is the last line of the continuation section.
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
		if (aInBlockComment)
		{
			if (  !(*aBuf == '*' && aBuf[1] == '/')  )
				// Avoid stripping ;comments since that would prevent detection of the comment-end
				// in cases like "; xxx */".
				return aBuf_length;
			//else the block comment ends at */ and there may be a ;line comment following it.
		}
		else if (*aBuf == g_CommentChar)
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

	// Handle comment-flags that appear to the right of a valid line:
	LPTSTR cp, prevp;
	for (cp = _tcschr(aBuf, g_CommentChar); cp; cp = _tcschr(cp + 1, g_CommentChar))
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
// aBuf must be a modifiable string as it may be temporarily terminated at different points.
// Returns CONDITION_TRUE, CONDITION_FALSE, or FAIL.
// Note: Don't assume that every line in the script that starts with '#' is a directive
// because hotkeys can legitimately start with that as well.  i.e., the following line should
// not be unconditionally ignored, just because it starts with '#', since it is a valid hotkey:
// #y::run, notepad
{
	TCHAR end_flags[] = {' ', '\t', '\0'}; // '\0' must be last.
	LPTSTR directive_end, parameter;
	if (   !(directive_end = StrChrAny(aBuf, end_flags))   )
	{
		directive_end = aBuf + _tcslen(aBuf); // Point it to the zero terminator.
		parameter = NULL;
	}
	else
		if (!*(parameter = omit_leading_whitespace(directive_end)))
			parameter = NULL;

	// Make a null-terminated copy of the potential directive name, if not already terminated,
	// so that overlapping names such as #MaxThreads and #MaxThreadsPerHotkey won't get mixed up.
	// This significantly reduces code size vs. using tcslicmp() in IS_DIRECTIVE_MATCH().
	LPCTSTR directive_name;
	TCHAR directive_name_buf[32];
	if (directive_end - aBuf >= _countof(directive_name_buf))
		return CONDITION_FALSE;
	if (*directive_end)
	{
		tcslcpy(directive_name_buf, aBuf, directive_end - aBuf + 1);
		directive_name = directive_name_buf;
	}
	else
		directive_name = aBuf;
	#define IS_DIRECTIVE_MATCH(directive) (!_tcsicmp(directive_name, directive))

	bool is_include_again = false; // Set default in case of short-circuit boolean.
	if (IS_DIRECTIVE_MATCH(_T("#Include")) || (is_include_again = IS_DIRECTIVE_MATCH(_T("#IncludeAgain"))))
	{
		// Standalone EXEs ignore this directive since the included files were already merged in
		// with the main file when the script was compiled.  These should have been removed
		// or commented out by Ahk2Exe, but just in case, it's safest to ignore them:
#ifdef AUTOHOTKEYSC
		return CONDITION_TRUE;
#else
		if (!parameter)
			return ScriptError(ERR_PARAM1_REQUIRED, aBuf);
		
		// Permit quote marks around the parameter.  In future, quote marks might be required,
		// or the parameter might be evaluated as an expression.
		parameter = strip_quote_marks(parameter);
		if (!*parameter)
			return ScriptError(ERR_PARAM1_INVALID);

		// v1.0.32:
		bool ignore_load_failure = (parameter[0] == '*' && ctoupper(parameter[1]) == 'I' && IS_SPACE_OR_TAB(parameter[2])); // Relies on short-circuit boolean order.
		if (ignore_load_failure)
			parameter += 3; // Skip over at most one space or tab, since others might be a literal part of the filename.

		if (*parameter == '<') // Support explicitly-specified <standard_lib_name>.
		{
			LPTSTR parameter_end = _tcschr(parameter, '>');
			if (parameter_end && !parameter_end[1])
			{
				++parameter; // Remove '<'.
				*parameter_end = '\0'; // Remove '>'.
				bool error_was_shown, file_was_found;
				// Save the working directory; see the similar line below for details.
				LPTSTR prev_dir = GetWorkingDir();
				// Attempt to include a script file from a Lib folder:
				IncludeLibrary(parameter, parameter_end - parameter, error_was_shown, file_was_found);
				// Restore the working directory.
				if (prev_dir)
				{
					SetCurrentDirectory(prev_dir);
					free(prev_dir);
				}
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
		// Save the working directory because LoadIncludedFile() changes it, and we want to retain any directory
		// set by a previous instance of "#Include DirPath" for any other instances of #Include below this line.
		LPTSTR prev_dir = GetWorkingDir();
		// Since above didn't return, it's a file (or non-existent file, in which case the below will display
		// the error).  This will also display any other errors that occur:
		ResultType result = LoadIncludedFile(include_path, is_include_again, ignore_load_failure) ? CONDITION_TRUE : FAIL;
		// Restore the working directory.
		if (prev_dir)
		{
			SetCurrentDirectory(prev_dir);
			free(prev_dir);
		}
		free(include_path);
		return result;
#endif
	}

	if (IS_DIRECTIVE_MATCH(_T("#DllLoad")))
	{
		if (!parameter)
		{
			// an empty #DllLoad restores the default search order.
			if (!SetDllDirectory(NULL))
				return ScriptError(ERR_INTERNAL_CALL);
			return CONDITION_TRUE;
		}

		// Permit quote marks around the parameter.  In future, quote marks might be required,
		// or the parameter might be evaluated as an expression.
		parameter = strip_quote_marks(parameter);

		// ignore failure if path is preceeded by "*i".
		bool ignore_load_failure = (parameter[0] == '*' && ctoupper(parameter[1]) == 'I'); // Relies on short-circuit boolean order.
		if (ignore_load_failure)
		{
			parameter += 2;
			if (IS_SPACE_OR_TAB(*parameter)) // Skip over at most one space or tab, since others might be a literal part of the filename.
				++parameter;
		}
		LPTSTR library_path;	// needs to be freed before return, unless DerefInclude return FAIL.
		if (!DerefInclude(library_path, parameter))
			return FAIL;
		DWORD attr = GetFileAttributes(library_path);
		if (attr != 0xFFFFFFFF && (attr & FILE_ATTRIBUTE_DIRECTORY)) // File exists and its a directory
		{
			BOOL result = SetDllDirectory(library_path);
			free(library_path);
			if (!result)
				return ScriptError(ERR_INTERNAL_CALL);
			return CONDITION_TRUE;
		}
		ResultType result = CONDITION_TRUE;	// set default
		HMODULE hmodule = LoadLibrary(library_path);

		if (hmodule)
			// "Pin" the dll so that the script cannot unload it with FreeLibrary.
			// This is done to avoid undefined behaviour when DllCall optimizations
			// resolves a proc address in a dll loaded by this directive.
			GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_PIN, library_path, &hmodule);  // MSDN regarding hmodule: "If the function fails, this parameter is NULL."

		if (!hmodule					// the library couldn't be loaded
			&& !ignore_load_failure)	// no *i was specified
			result = ScriptError(_T("Failed to load DLL."), library_path);
		free(library_path);
		return result;
	} // end #DllLoad

	if (IS_DIRECTIVE_MATCH(_T("#NoTrayIcon")))
	{
		g_NoTrayIcon = true;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#SingleInstance")))
	{
		// Since "Prompt" is the default mode when #SingleInstance is not present, make #SingleInstance
		// on its own activate "Force" mode.  But also allow "Force" and "Prompt" explicitly, since it
		// might help clarity (and may be useful for overriding an #include?).
		if (!parameter || !_tcsicmp(parameter, _T("Force")))
			g_AllowOnlyOneInstance = SINGLE_INSTANCE_REPLACE;
		else if (!_tcsicmp(parameter, _T("Prompt")))
			g_AllowOnlyOneInstance = SINGLE_INSTANCE_PROMPT;
		else if (!_tcsicmp(parameter, _T("Ignore")))
			g_AllowOnlyOneInstance = SINGLE_INSTANCE_IGNORE;
		else if (!_tcsicmp(parameter, _T("Off")))
			g_AllowOnlyOneInstance = SINGLE_INSTANCE_OFF;
		else // Since it could be a typo like "Of", alert the user:
			return ScriptError(ERR_PARAM1_INVALID, aBuf);
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#UseHook")))
	{
		if (!ConvertDirectiveBool(parameter, g_ForceKeybdHook, true))
			return ScriptError(ERR_PARAM1_INVALID, parameter);
		return CONDITION_TRUE;
	}

	// L4: Handle #HotIf (expression) directive.
	if (IS_DIRECTIVE_MATCH(_T("#HotIf")))
	{
		// Disallow #HotIf inside a function or class, but allow it interspersed with hotkey labels;
		// i.e. g->CurrentFunc and mLastHotFunc must both be null or the same non-null value.
		if (mClassObjectCount || g->CurrentFunc != mLastHotFunc) 
			return ScriptError(ERR_INVALID_USAGE);

		if (!parameter) // The omission of the parameter indicates that any existing criteria should be turned off.
		{
			g->HotCriterion = NULL; // Indicate that no criteria are in effect for subsequent hotkeys.
			return CONDITION_TRUE;
		}

		// Check for a duplicate #HotIf expression;
		//  - Prevents duplicate hotkeys under separate copies of the same expression.
		//  - HotIf would only be able to select the first expression with the given source code.
		//  - Conserves memory.
		if (g->HotCriterion = FindHotkeyIfExpr(parameter))
			return CONDITION_TRUE;

		auto pending_hotfunc = mLastHotFunc;
		
		// Create a function to return the result of the expression
		// specified by "parameter":

		CreateHotFunc();
		
		auto fr = Hotkey::IfExpr(mLastHotFunc); // Set the new criterion.
		if (fr != OK) // Only OK and FR_E_OUTOFMEM should be possible given how it's called.
			return fr == FR_E_OUTOFMEM ? ScriptError(ERR_OUTOFMEM) : FAIL;

		g->HotCriterion->OriginalExpr = SimpleHeap::Alloc(parameter);
		
		auto func = mLastHotFunc; // AddLine will set mLastHotFunc to nullptr below
		
		if (!AddLine(ACT_BLOCK_BEGIN)
			|| !ParseAndAddLine(parameter, ACT_HOTKEY_IF) // PreparseExpressions will change this to ACT_RETURN
			|| !AddLine(ACT_BLOCK_END))
			return ScriptError(ERR_OUTOFMEM);

		func->mJumpToLine->mAttribute = g->HotCriterion;	// Must be set for PreparseHotkeyIfExpr

		ASSERT(!g->CurrentFunc && !mLastHotFunc); // Should be null due to ACT_BLOCK_END and prior checks.
		g->CurrentFunc = mLastHotFunc = pending_hotfunc; // Restore any pending hotkey function.
		
		return CONDITION_TRUE;
	}

	// L4: Allow #HotIf timeout to be adjusted.
	if (IS_DIRECTIVE_MATCH(_T("#HotIfTimeout")))
	{
		if (parameter)
			g_HotExprTimeout = ATOU(parameter);
		return CONDITION_TRUE;
	}

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
				ConvertEscapeSequences(g_EndChars, NULL);
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
				, g_HSDetectWhenInsideWord, g_HSDoReset, g_HSSameLineAction, g_SuspendExemptHS);
		}
		return CONDITION_TRUE;
	}

	if (IS_DIRECTIVE_MATCH(_T("#MaxThreadsPerHotkey")))
	{
		if (parameter)
		{
			// Use value as a temp holder since it's int vs. UCHAR and can thus detect very large or negative values:
			int value = ATOI(parameter);  // parameter was set to the right position by the above macro
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
		if (!ConvertDirectiveBool(parameter, g_MaxThreadsBuffer, true))
			return ScriptError(ERR_PARAM1_INVALID, parameter);
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#MaxThreads")))
	{
		if (parameter)
		{
			int value = ATOI(parameter);  // parameter was set to the right position by the above macro
			if (value > MAX_THREADS_LIMIT) // For now, keep this limited to prevent stack overflow due to too many pseudo-threads.
				value = MAX_THREADS_LIMIT; // UPDATE: To avoid array overflow, this limit must by obeyed except where otherwise documented.
			else if (value < 1)
				value = 1;
			g_MaxThreadsTotal = value;
		}
		return CONDITION_TRUE;
	}

	if (IS_DIRECTIVE_MATCH(_T("#SuspendExempt")))
	{
		if (!ConvertDirectiveBool(parameter, g_SuspendExempt, true))
			return ScriptError(ERR_PARAM1_INVALID, aBuf);
		return CONDITION_TRUE;
	}

	if (IS_DIRECTIVE_MATCH(_T("#ClipboardTimeout")))
	{
		if (parameter)
			g_ClipboardTimeout = ATOI(parameter);  // parameter was set to the right position by the above macro
		return CONDITION_TRUE;
	}

	if (IS_DIRECTIVE_MATCH(_T("#WinActivateForce")))
	{
		g_WinActivateForce = true;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH(_T("#ErrorStdOut")))
	{
		// Permit quote marks around the parameter.  In future, quote marks might be required,
		// or the parameter might be evaluated as an expression.
		parameter = strip_quote_marks(parameter);
		SetErrorStdOut(parameter);
		if (mErrorStdOutCP == -1)
			return ScriptError(ERR_PARAM1_INVALID, parameter);
		return CONDITION_TRUE;
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

		static const LPCTSTR sWarnTypes[] = { WARN_TYPE_STRINGS };
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

		static const LPCTSTR sWarnModes[] = { WARN_MODE_STRINGS };
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

		if (warnType == WARN_LOCAL_SAME_AS_GLOBAL || warnType == WARN_ALL)
			g_Warn_LocalSameAsGlobal = warnMode;

		if (warnType == WARN_UNREACHABLE || warnType == WARN_ALL)
			g_Warn_Unreachable = warnMode;

		if (warnType == WARN_VAR_UNSET || warnType == WARN_ALL)
			g_Warn_VarUnset = warnMode;

		return CONDITION_TRUE;
	}

	if (IS_DIRECTIVE_MATCH(_T("#Requires")))
	{
#ifdef AUTOHOTKEYSC
		return CONDITION_TRUE; // Omit the checks for compiled scripts to reduce code size.
#else
		if (!parameter)
			return ScriptError(ERR_PARAM1_REQUIRED);

		if (!_tcsnicmp(parameter, _T("AutoHotkey"), 10))
		{
			if (!parameter[10]) // Just #requires AutoHotkey; would seem silly to warn the user in this case.
				return CONDITION_TRUE;

			if (IS_SPACE_OR_TAB(parameter[10]))
			{
				TCHAR word[32];
				for (LPCTSTR end, cp = parameter + 11; ; cp = end)
				{
					cp = omit_leading_whitespace(cp);
					if (!*cp)
						return CONDITION_TRUE;
					
					for (end = cp; *end && !IS_SPACE_OR_TAB(*end); ++end);
					tcslcpy(word, cp, min(_countof(word), end - cp + 1));

					// Allow these words when appropriate: 32-bit, 64-bit
					if (!_tcsicmp(word, _T(AHK_BIT)))
						continue;

					// It's either an unment requirement or a version number.
					if (VersionSatisfies(T_AHK_VERSION, word))
						continue;

					break;
				}
			}
		}
		// Unmet or unrecognized requirement.
		return RequirementError(parameter);
#endif
	}

	// Otherwise, report that this line isn't a directive:
	return CONDITION_FALSE;
}



ResultType Script::ConvertDirectiveBool(LPTSTR aBuf, bool &aResult, bool aDefault)
{
	if (!aBuf || !*aBuf)
		aResult = aDefault;
	else if (!_tcsicmp(aBuf, _T("True")) || *aBuf == '1' && !aBuf[1])
		aResult = true;
	else if (!_tcsicmp(aBuf, _T("False")) || *aBuf == '0' && !aBuf[1])
		aResult = false;
	else
		return FAIL;
	return OK;
}



ResultType Script::RequirementError(LPCTSTR aRequirement)
{
	TCHAR buf[512];
	sntprintf(buf, _countof(buf), _T("This script requires %s.\n\nCurrent interpreter: %s v%s %s\n%s")
		, aRequirement, T_AHK_NAME, T_AHK_VERSION, _T(AHK_BIT), mOurEXE);
	return ScriptError(buf);
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



ResultType Script::UpdateOrCreateTimer(IObject *aCallback
	, bool aUpdatePeriod, __int64 aPeriod, bool aUpdatePriority, int aPriority)
// Caller should specific a blank aPeriod to prevent the timer's period from being changed
// (i.e. if caller just wants to turn on or off an existing timer).  But if it does this
// for a non-existent timer, that timer will be created with the default period as specified in
// the constructor.
{
	ScriptTimer *timer;
	for (timer = mFirstTimer; timer != NULL; timer = timer->mNextTimer)
		if (timer->mCallback == aCallback) // Match found.
			break;
	bool timer_existed = (timer != NULL);
	if (!timer_existed)  // Create it.
	{
		timer = new ScriptTimer(aCallback);
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

	if (!timer->mEnabled)
	{
		timer->mEnabled = true;
		++mTimerEnabledCount;
		SET_MAIN_TIMER  // Ensure the API timer is always running when there is at least one enabled timed subroutine.
		if (timer_existed)
		{
			// mEnabled is currently used to mark a running timer for deletion upon return.
			// Since the timer could be recreated by a different thread which just happens
			// to interrupt the timer thread, it seems best for consistency to reset Period
			// and Priority as though it had already been deleted.
			aUpdatePeriod = true;
			aUpdatePriority = true;
		}
	}

	if (aUpdatePeriod)
	{
		if (aPeriod < 0) // Support negative periods to mean "run only once".
		{
			timer->mRunOnlyOnce = true;
			timer->mPeriod = (DWORD)-aPeriod;
		}
		else // Positive number.
		{
			timer->mPeriod = (DWORD)aPeriod;
			timer->mRunOnlyOnce = false;
		}
		if (timer->mPeriod == 1)
			timer->mPeriod = 0; // Allow execution within the same tick (though not guaranteed).
	}

	if (aUpdatePriority)
		timer->mPriority = aPriority;

	if (aUpdatePeriod || !aUpdatePriority)
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
		if (timer->mCallback == aLabel) // Match found.
		{
			// Disable it, even if it's not technically being deleted yet.
			if (timer->mEnabled)
				timer->Disable(); // Keeps track of mTimerEnabledCount and whether the main timer is needed.
			if (timer->mExistingThreads // This condition differs from g->CurrentTimer == timer, which only detects the "top-most" timer.
				|| timer->mDeleteLocked)
			{
				// In this case we can't delete the timer yet, but CheckScriptTimers()
				// will check mEnabled after the callback returns and will delete it then.
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



Label *Script::FindLabel(LPCTSTR aLabelName)
// Returns the first label whose name matches aLabelName, or NULL if not found.

// If duplicates labels are now possible, callers must be aware that only the first match is returned.
// This helps performance by requiring on average only half the labels to be searched before
// a match is found.
{
	if (!aLabelName || !*aLabelName) return NULL;
	Label *label;
	if (g->CurrentFunc)
		label = g->CurrentFunc->mFirstLabel; // Search only local labels, since global labels aren't valid jump targets in this case.
	else
		label = mFirstLabel;
	for ( ; label; label = label->mNextLabel)
		if (!_tcsicmp(label->mName, aLabelName)) // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales; 3) performance.
			return label; // Match found.
	return NULL; // No match found.
}



ResultType Script::AddLabel(LPTSTR aLabelName, bool aAllowDupe)
// Returns OK or FAIL.
{
	if (!*aLabelName)
		return FAIL; // For now, silent failure because callers should check this beforehand.
	Label *&first_label = g->CurrentFunc ? g->CurrentFunc->mFirstLabel : mFirstLabel;
	Label *&last_label  = g->CurrentFunc ? g->CurrentFunc->mLastLabel  : mLastLabel;
	if (!aAllowDupe && FindLabel(aLabelName))
	{
		// Don't attempt to dereference label->mJumpToLine because it might not
		// exist yet.  Example:
		// label1:
		// label1:  <-- This would be a dupe-error but it doesn't yet have an mJumpToLine.
		// return
		return ScriptError(_T("Duplicate label."), aLabelName);
	}
	LPTSTR new_name = SimpleHeap::Alloc(aLabelName);
	Label *the_new_label = new Label(new_name); // Pass it the dynamic memory area we created.
#ifdef CONFIG_DLL
	++mLabelCount;
	the_new_label->mLineNumber = mCombinedLineNumber;
	the_new_label->mFileIndex = mCurrFileIndex;
#endif
	the_new_label->mPrevLabel = last_label;  // Whether NULL or not.
	if (first_label == NULL)
		first_label = the_new_label;
	else
		last_label->mNextLabel = the_new_label;
	// This must be done after the above:
	last_label = the_new_label;
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



ResultType Script::ParseAndAddLine(LPTSTR aLineText, ActionTypeType aActionType, LPTSTR aLiteralMap, size_t aLiteralMapLength)
// Returns OK or FAIL.
// aLineText needs to be a string whose contents are modifiable (though the string won't be made any
// longer than it is now, so it doesn't have to be of size LINE_SIZE). This helps performance by
// allowing the string to be split into sections without having to make temporary copies.
// aLineText must point to a buffer with room to append "()" if it may contain a function call statement
// (which is only possible when aActionType is ACT_INVALID/omitted or an ACT with possible subaction).
{
#ifdef _DEBUG
	if (!aLineText || !*aLineText && !aActionType)
		return ScriptError(_T("DEBUG: ParseAndAddLine() called incorrectly."));
#endif
	size_t line_length = _tcslen(aLineText); // Length is needed in a couple of places.

	TCHAR action_name[MAX_VAR_NAME_LENGTH + 1], *end_marker;
	if (aActionType) // Currently can be ACT_EXPRESSION or ACT_HOTKEY_IF.
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
			if (!_tcsnicmp(aLineText, _T("global"), 6))
			{
				cp = aLineText + 6; // The character after the declaration word.
				declare_type = VAR_DECLARE_GLOBAL;
			}
			else
			{
				if (!_tcsnicmp(aLineText, _T("local"), 5))
				{
					cp = aLineText + 5; // The character after the declaration word.
					declare_type = VAR_DECLARE_LOCAL;
				}
				else if (!_tcsnicmp(aLineText, _T("static"), 6)) // Static also implies local (for functions that default to global).
				{
					cp = aLineText + 6; // The character after the declaration word.
					declare_type = VAR_DECLARE_STATIC;
				}
				else // It's not the word "global", "local", or static, so no further checking is done.
					break;
			}

			if (*cp && !IS_SPACE_OR_TAB(*cp)) // There is a character following the word local but it's not a space or tab.
				break; // It doesn't qualify as being the global or local keyword because it's something like global2.
			if (*cp && *(cp = omit_leading_whitespace(cp))) // Second *cp is probably always non-null since caller rtrimmed it, but even if not it's handled correctly.
			{
				if (!IS_IDENTIFIER_CHAR(*cp))
					break;
			}
			else // It's the word "global", "local", "static" by itself.
			{
				// Any combination of declarations is allowed here for simplicity, but only declarations can
				// appear above this line:
				if (mNextLineIsFunctionBody && declare_type != VAR_DECLARE_LOCAL)
				{
					g->CurrentFunc->mDefaultVarType = declare_type;
					// No further action is required.
					return OK;
				}
				// Otherwise, it occurs too far down in the body.
				return ScriptError(ERR_UNEXPECTED_DECL, aLineText);
			}
			
			// Since above didn't break or return, a variable is being declared.

			if (!g->CurrentFunc && (declare_type & VAR_LOCAL))
				return ScriptError(ERR_UNEXPECTED_DECL, aLineText);

			bool open_brace_was_added, belongs_to_line_above;
			size_t var_name_length;
			LPTSTR item;

			for (belongs_to_line_above = mPendingParentLine
				, open_brace_was_added = false, item = cp
				; *item;) // FOR EACH COMMA-SEPARATED ITEM IN THE DECLARATION LIST.
			{
				LPTSTR item_end = find_identifier_end(item + 1);
				var_name_length = (VarSizeType)(item_end - item);

				Var *var = nullptr;
				Var *global_var = nullptr;
				if (declare_type & VAR_GLOBAL)
				{
					// This is done separately since it may also need to be added to mStaticVars:
					if (  !(global_var = FindOrAddVar(item, var_name_length, declare_type))  )
						return FAIL; // It already displayed the error.
				}

				if (g->CurrentFunc)
				{
					int insert_pos;
					VarList *varlist;
					ResultType result = OK;
					// Search both locals and statics so any conflicts are detected, but set insert_pos
					// according to which list we want to insert into.  For "global", we want to insert
					// into the static list to make it accessible to the function.
					int insert_into = global_var ? (VAR_LOCAL_STATIC | VAR_LOCAL) : declare_type;
					var = FindVar(item, var_name_length, insert_into, &varlist, &insert_pos, &result);
					if (!var)
					{
						if (!result)
							return result;
						if (global_var) // Implies (declare_type & VAR_GLOBAL) != 0.
						{
							// Insert could be skipped if the function is assume-global, but isn't because:
							//  - When !mNextLineIsFunctionBody, mDefaultVarType might still be changed.
							//  - It seems harmless to add any explicitly declared globals here.
							//  - There might be some reason to refer to the declared globals in future.
							if (!varlist->Insert(global_var, insert_pos))
								return ScriptError(ERR_OUTOFMEM);
							var = global_var;
						}
						else
						{
							var = AddVar(item, var_name_length, varlist, insert_pos, declare_type);
							if (!var)
								return FAIL; // It already displayed the error.
						}
					}
					// Detect the following conflicts:
					//  - Declaring local but found a static/global in conflicts, or vice versa.
					//  - Declaring static but found a global in varlist, or vice versa.
					//  - Declaring local but found a parameter in varlist.
					//  - Declaring a built-in variable as local or static.
					// But permit the following:
					//  - Exact duplicate declarations, such as for two different code paths.
					if (var->Scope() != declare_type)
						return ConflictingDeclarationError(Var::DeclarationType(declare_type), var);
				}
				else
					var = global_var;

				item_end = omit_leading_whitespace(item_end); // Move up to the next comma, assignment-op, or '\0'.

				LPTSTR the_operator = item_end;
				switch(*item_end)
				{
				case ',':  // No initializer is present for this variable, so move on to the next one.
					item = omit_leading_whitespace(item_end + 1); // Set "item" for use by the next iteration.
					continue; // No further processing needed below.
				case '\0': // No initializer is present for this variable, so move on to the next one.
					item = item_end; // Set "item" for use by the next iteration.
					continue;
				case ':':
					if (item_end[1] == '=')
					{
						item_end += 2; // Point to the character after the ":=".
						break;
					}
					// Colon with no following '='. Fall through to the default case.
				default:
					// Since this variable might already have a value, compound assignments are
					// permitted (although not very conventional).  This is mostly for globals,
					// which might only be referenced on this one line (and outside the function),
					// but might have some utility with "static" due to its once-only nature.
					// += -= *= /= .= |= &= ^=
					if (_tcschr(_T("+-*/.|&^"), *item_end) && item_end[1] == '=')
						break;
					// //= <<= >>= >>>=
					if (_tcschr(_T("/<>"), *item_end) && item_end[1] == *item_end && (item_end[2] == '=' // //= <<= >>=
						|| (*item_end == '>' && item_end[2] == *item_end && item_end[3] == '='))) // >>>=
						break;
					return ScriptError(ERR_INVALID_VARDECL, item);
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

				item_end += FindExprDelim(item_end); // FIND THE NEXT "REAL" COMMA (or the end of the string).
				
				// Above has now found the final comma of this sub-statement (or the terminator if there is no comma).
				LPTSTR terminate_here = omit_trailing_whitespace(item, item_end-1) + 1; // v1.0.47.02: Fix the fact that "x=5 , y=6" would preserve the whitespace at the end of "5".  It also fixes wrongly showing a syntax error for things like: static d="xyz"  , e = 5
				TCHAR orig_char = *terminate_here;
				*terminate_here = '\0'; // Temporarily terminate (it might already be the terminator, but that's harmless).

				// Performance note: Simple assignments of a literal number or string to a variable currently
				// perform better standalone than combined with the comma operator.  This is almost certainly
				// due to optimizations which avoid the need to call ExpandExpression for those cases.
				// If this ever changes, note that the braces added below are needed for ACT_STATIC regardless
				// of how many initializers there are, due to the way it removes its own Line after execution.
				if (belongs_to_line_above && !open_brace_was_added) // v1.0.46.01: Put braces to allow initializers to work even directly under an IF/ELSE/LOOP.
				{
					if (!AddLine(ACT_BLOCK_BEGIN))
						return FAIL;
					open_brace_was_added = true;
				}
				// Call Parse() vs. AddLine() because it detects and optimizes simple assignments into
				// non-expressions for faster runtime execution.
				if (!ParseAndAddLine(item, declare_type == VAR_DECLARE_STATIC ? ACT_STATIC : ACT_INVALID))
					return FAIL; // Above already displayed the error.

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
		end_marker = ParseActionType(action_name, aLineText, true);
	}
	
	// Above has ensured that end_marker is the address of the last character of the action name,
	// or NULL if there is no action name.
	// Find the arguments (not to be confused with exec_params) of this action, if it has any:
	LPTSTR action_args;
	bool could_be_named_action;
	if (end_marker)
	{
		action_args = omit_leading_whitespace(end_marker);
		TCHAR end_char = *end_marker;
		could_be_named_action = (!end_char || IS_SPACE_OR_TAB(end_char) || end_char == '(')
			&& *action_args != '='; // Allow for a more specific error message for `x = y`, to ease transition from v1.
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
	bool all_args_are_expressions = false;

	if (!aActionType) // i.e. the caller hasn't yet determined this line's action type.
	{
		//////////////////////////////////////////////////////
		// Detect operators and assignments such as := and +=
		//////////////////////////////////////////////////////
		// This section is done before the section that checks whether action_name is a valid command
		// because it avoids ambiguity in a line such as the following:
		//    Input := test  ; Would otherwise be confused with the Input command.

		if (*aLineText == '"' || *aLineText == '\'')
		{
			action_args = aLineText + FindTextDelim(aLineText, *aLineText, 1, aLiteralMap);
			if (*action_args)
			{
				action_args = omit_leading_whitespace(action_args + 1);
				if (*action_args != '.') // Don't treat "" := x as valid.
					action_args = aLineText;
			}
		}

		TCHAR action_args_2nd_char = action_args[1];

		switch(*action_args)
		{
		case ':':
			// v1.0.40: Allow things like "MsgBox :: test" to be valid by insisting that '=' follows ':'.
			// v2.0: The example above is invalid, but it's still best to verify this is really ':='.
			if (action_args_2nd_char == '=') // i.e. :=
				aActionType = ACT_ASSIGNEXPR;
			break;
		case '+':
		case '-':
			// Support for ++i and --i.  Do a complete validation/recognition of the operator to allow
			// a line such as the following, which omits the first optional comma, to still be recognized
			// as a command rather than a variable-with-operator:
			// SetWinDelay -1
			// v2.0: Cases like `x ++` aren't currently supported due to the risk of misinterpreting
			// something like `myfn ++x` or other contrived-but-valid calls like `myfn ++(b ? x : y)`
			// or `myfn ++++x`.  Unlike ternary, it's probably more common to write ++ without the space.
			if (!end_marker && action_args_2nd_char == *action_args // i.e. the pre-increment/decrement operator; e.g. ++index or --index.
				|| action_args_2nd_char == '=') // i.e. x+=y or x-=y (by contrast, post-increment/decrement is recognized only after we check for a command name to cut down on ambiguity).
				aActionType = ACT_EXPRESSION; // Mark this line as a stand-alone expression.
			break;
		case '*': // *=
		case '|': // |=
		case '&': // &=
		case '^': // ^=
			if (action_args_2nd_char == '=')
				aActionType = ACT_EXPRESSION; // Mark this line as a stand-alone expression.
			break;
		case '/':
			if (action_args_2nd_char == '=' // i.e. /=
				|| action_args_2nd_char == '/' && action_args[2] == '=') // i.e. //=
				aActionType = ACT_EXPRESSION; // Mark this line as a stand-alone expression.
			break;
		case '?': // Stand-alone ternary beginning with a variable, such as true ? fn1() : fn2().
			// v2.0: Even if this is a valid function name (which is impossible to determine for
			// user-defined functions at this stage due), this can't be a valid function call stmt
			// since its first parameter would begin with the '?' operator.
			aActionType = ACT_EXPRESSION; // Mark this line as a stand-alone expression.
			break;
		case '>':
		case '<':
			if (action_args_2nd_char == *action_args && ( action_args[2] == '='	// i.e. >>= and <<=
				|| (action_args_2nd_char == '>' && action_args[2] == '>' && action_args[3] == '=') )) // >>>=
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
				LPTSTR cp, id_end;
				bool has_space_or_tab;
				for (;;) // L35: Loop to fix x.y.z() and similar.
				{
					id_end = find_identifier_end(id_begin);
					if (*id_end == '(' // Allow function/method Call as standalone expression.
						|| *id_end == g_DerefChar) // Allow dynamic property/method access (too hard to validate what's to the right of %).
					{
						aActionType = ACT_EXPRESSION;
						break;
					}
					if (id_end == id_begin)
						// No valid identifier, doesn't look like a valid expression.
						break;
					has_space_or_tab = IS_SPACE_OR_TAB(*id_end);
					cp = omit_leading_whitespace(id_end);
					if (*cp) // Avoid checking cp[1] and cp[2] when !*cp.
					if (*cp == '[' // x.y[z]
						|| cp[1] == '=' && _tcschr(_T(":+-*/|&^."), cp[0]) // Two-char assignment operator.
						// or there are two repeated characters 
						|| cp[1] == cp[0]
						&& ( ( _tcschr(_T("/<>"), cp[0]) && cp[2] == '=' // //=, <<= or >>=
									|| *cp == '+' || *cp == '-' ) // x.y++ or x.y--
							// or three repeated characters:
							|| (cp[0] == '>' && cp[2] == '>' && cp[3] == '=') )	) // >>>=
					{	// Allow Set and bracketed Get as standalone expression.
						aActionType = ACT_EXPRESSION;
						break;
					}
					if (*cp != '.')
					{
						if (!*cp || has_space_or_tab)
						{
							end_marker = id_end;
							could_be_named_action = true;
							// Let the parentheses be added in the section below.
						}
						//else: Neither a command nor a legal standalone expression.
						break;
					}
					id_begin = cp + 1;
				}
			}
			break;
		//default: Leave aActionType set to ACT_INVALID. This also covers case '\0' in case that's possible.
		} // switch()

		if (aActionType) // An assignment or other type of action was discovered above.
		{
			if (aActionType == ACT_ASSIGNEXPR)
			{
				// Find the first non-function comma.
				// This is done because ACT_ASSIGNEXPR needs to make comma-separated sub-expressions
				// into one big ACT_EXPRESSION so that the leftmost sub-expression will get evaluated
				// prior to the others (for consistency and as documented).  However, this has at
				// least one side-effect; namely that if expression evaluation is aborted for some
				// reason, the assignment is skipped completely rather than assigning a blank value.
				// ALSO: ACT_ASSIGNEXPR is made into ACT_EXPRESSION *only* when multi-statement
				// commas are present because it performs much better for trivial assignments,
				// even some which aren't optimized to become non-expressions.
				LPTSTR cp = action_args + FindExprDelim(action_args, g_delimiter, 2);
				if (*cp) // Found a delimiting comma other than one in a sub-statement or function. Shouldn't need to worry about unquoted escaped commas since they don't make sense with += and -=.
				{
					// Any non-function comma qualifies this as multi-statement.
					aActionType = ACT_EXPRESSION;
				}
				else
				{
					// The following converts:
					// x := 2 -> ACT_ASSIGNEXPR, x, 2
					*action_args = g_delimiter; // Replace the ":" with a delimiter for later parsing.
					action_args[1] = ' '; // Remove the "=" from consideration.
				}
			}
			//else it's already an isolated expression, so no changes are desired.
			action_args = aLineText; // Since this is an assignment and/or expression, use the line's full text for later parsing.
		} // if (aActionType)

	} // Handling of assignments and other operators.
	//else aActionType was already determined by the caller.

	if (!aActionType && could_be_named_action) // Caller nor logic above has yet determined the action.
	{
		aActionType = ConvertActionType(action_name); // Is this line a control flow statement?

		if (!aActionType && *end_marker != '(')
		{
			if (*action_name < '0' || *action_name > '9') // Exclude numbers, since no function name can start with a number.
			{
				// Convert function/method call statements to function/method calls.
				if (*end_marker) // Replace space or tab with parenthesis.
					*end_marker = '(';
				else
					aLineText[line_length++] = '(';
				aLineText[line_length++] = ')';
				aLineText[line_length] = '\0';
				action_args = aLineText;
				aActionType = ACT_EXPRESSION;
			}
		}
	}

	if (!aActionType) // Didn't find any action or command in this line.
	{
		// v1.0.41: Support one-true brace style even if there's no space, but make it strict so that
		// things like "Loop{ string" are reported as errors (in case user intended an object literal).
		if (*action_args == '{')
		{
			switch (ActionTypeType otb_act = ConvertActionType(action_name))
			{
			case ACT_LOOP:
			case ACT_SWITCH:
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
		else if (*action_args == ':')
		{
			// Default isn't handled as its own action type because that would only cover the cases
			// where there is a space after the name.  This covers "Default: subaction", "Default :"
			// and "Default : subaction".  "Default:" on its own is handled by label-parsing.
			if (!_tcsicmp(action_name, _T("Default")))
			{
				if (!AddLine(ACT_CASE))
					return FAIL;
				action_args = omit_leading_whitespace(action_args + 1);
				if (!*action_args)
					return OK;
				return ParseAndAddLine(action_args);
			}
		}
		if (!aActionType && _tcschr(EXPR_ALL_SYMBOLS, *action_args))
		{
			LPTSTR question_mark;
			if ((*action_args == '+' || *action_args == '-') && action_args[1] == *action_args) // Post-inc/dec. See comments further below.
			{
				aActionType = ACT_EXPRESSION; // Mark this line as a stand-alone expression.
				action_args = aLineText; // Since this is an expression, use the line's full text for later parsing.
			}
			else if ((question_mark = _tcschr(action_args,'?')) && _tcschr(question_mark,':')) // Rough check (see comments below). Relies on short-circuit boolean order.
			{
				// To avoid hindering load-time error detection such as misspelled command names, allow stand-alone
				// expressions only for things that can produce a side-effect (currently only ternaries like
				// the ones mentioned later below need to be checked since the following other things were
				// previously recognized as ACT_EXPRESSION if appropriate: function-calls, post- and
				// pre-inc/dec (++/--), and assignment operators like := += *=
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
			if (end_marker && (*end_marker == '(' || *end_marker == '[') // v1.0.46.11: Recognize as multi-statements that start with a function, like "fn(), x:=4".  v1.0.47.03: Removed the following check to allow a close-brace to be followed by a comma-less function-call: strchr(action_args, g_delimiter).
				|| *action_args == g_DerefChar // Something beginning with a double-deref (which may contain expressions).
				|| *aLineText == '(' // Probably an expression with parentheses to control order of evaluation. In this case, aLineText == action_args, so it was already handled by the condition above.
				)
			{
				aActionType = ACT_EXPRESSION; // Mark this line as a stand-alone expression.
				action_args = aLineText; // Since this is a function-call followed by a comma and some other expression, use the line's full text for later parsing.
			}
			else if (*action_args == '=')
				// v2: Give a more specific error message since the user probably meant to do an old-style assignment.
				return ScriptError(_T("Syntax error. Did you mean to use \":=\"?"), aLineText);
			else if (*action_args == g_delimiter)
				return ScriptError(_T("Function calls require a space or \"(\".  Use comma only between parameters."), aLineText);
			else
				// v1.0.40: Give a more specific error message now that hotkeys can make it here due to
				// the change that avoids the need to escape double-colons:
				return ScriptError(_tcsstr(aLineText, HOTKEY_FLAG) ? ERR_INVALID_HOTKEY : ERR_UNRECOGNIZED_ACTION, aLineText);
		}
	} // if (!aActionType)
	else if (aActionType == ACT_LOOP)
	{
		LPTSTR cp = StrChrAny(action_args, EXPR_OPERAND_TERMINATORS);
		if (cp && cp > action_args && (IS_SPACE_OR_TAB(*cp) || *cp == g_delimiter))
		{
			LPTSTR next_nonspace = omit_leading_whitespace(cp);
			bool comma_was_used = *next_nonspace == g_delimiter;
			// Now cp points at the next character after the first word, while next_nonspace points
			// at the next non-whitespace character after that.  If comma_was_used, there is also
			// a comma and potentially more whitespace to skip if a valid loop keyword is used.
			TCHAR orig_char = *cp;
			*cp = '\0'; // Terminate the sub-command name.
			if (!_tcsicmp(_T("Files"), action_args))
				aActionType = ACT_LOOP_FILE;
			else if (!_tcsicmp(_T("Reg"), action_args))
				aActionType = ACT_LOOP_REG;
			else if (!_tcsicmp(_T("Read"), action_args))
				aActionType = ACT_LOOP_READ;
			else if (!_tcsicmp(_T("Parse"), action_args))
				aActionType = ACT_LOOP_PARSE;
			*cp = orig_char; // Undo the temporary termination.
			if (aActionType != ACT_LOOP) // Loop sub-type discovered above.
				// Set action_args to the start of the actual args, skipping the optional delimiter.
				action_args = comma_was_used ? omit_leading_whitespace(next_nonspace + 1) : next_nonspace;
			else if (comma_was_used)
				// This is a plain word or variable reference followed by a comma.  It could be a
				// multi-statement expression, but in this case the first sub-statement would have
				// no effect, so it's probably an error.
				return ScriptError(_T("Invalid loop type."), aLineText);
		}
	}
	int mark;
	LPTSTR subaction_start = NULL;
	// Perform some pre-processing to allow the parameter list of a control flow statement
	// to be enclosed in parentheses.  For goto/break/continue, this changes the parameter
	// from a literal label to an expression.  For others it is just a matter of coding style.
	// No pre-processing is needed for statements with exactly one parameter, since they will
	// be interpreted correctly with or without parentheses.
	switch (aActionType)
	{
	// Cases not handled here because they do not require pre-processing:
	//case ACT_IF:
	//case ACT_WHILE:
	//case ACT_UNTIL:
	// Cases omitted to prevent empty "()" from being allowed:
	//case ACT_LOOP:
	//case ACT_RETURN:
	case ACT_FOR:
	case ACT_CATCH:
		end_marker = action_args; // Handle both "for(...)" and "for (...)" below.
	case ACT_GOTO:
	case ACT_BREAK:    // These only support constant expressions (checked later).
	case ACT_CONTINUE: //
		if (end_marker && *end_marker == '(')
		{
			LPTSTR last_char = aLineText + line_length - 1;
			if (*last_char == '{' && (aActionType == ACT_FOR || aActionType == ACT_CATCH))
			{
				add_openbrace_afterward = true;
				last_char = omit_trailing_whitespace(end_marker, last_char - 1);
			}
			if (*last_char == ')')
			{
				// Remove the parentheses (and possible open brace) and trailing space.
				ASSERT(action_args == end_marker);
				++action_args;
				last_char = omit_trailing_whitespace(end_marker, last_char - 1);
				last_char[1] = '\0';
				// Treat this like a function call: all parameters are sub-expressions.
				all_args_are_expressions = true;
			}
		}
		if (aActionType != ACT_CATCH || *action_args != '{') // i.e. "Catch varname" must be handled a different way.
			break;
		// Otherwise, fall through to handle "Catch {":
	case ACT_ELSE:
	case ACT_TRY:
	case ACT_FINALLY:
		if (end_marker && *end_marker == '(') // Seems best to treat this as an error, perhaps reserve for future use.
			return ScriptError(ERR_EXPR_SYNTAX, aLineText);
		if (*action_args == '{')
		{
			add_openbrace_afterward = true;
			action_args = omit_leading_whitespace(action_args + 1);
		}
		if (*action_args)
		{
			subaction_start = action_args;
			*--action_args = '\0'; // Relies on there being a space, tab or brace at this position.
		}
		break;
	case ACT_CASE:
		mark = FindExprDelim(action_args, ':');
		if (!action_args[mark])
			return ScriptError(ERR_MISSING_COLON, aLineText);
		action_args[mark] = '\0';
		rtrim(action_args);
		subaction_start = omit_leading_whitespace(action_args + mark + 1);
		if (!*subaction_start) // Nothing to the right of ':'.
			subaction_start = NULL;
		break;
	}

	Action &this_action = g_act[aActionType];

	//////////////////////////////////////////////////////////////////////////////////////////////
	// Handle escaped-sequences (escaped delimiters and all others except variable deref symbols).
	// This section must occur after all other changes to the pointer value action_args have
	// occurred above.
	//////////////////////////////////////////////////////////////////////////////////////////////
	LPTSTR literal_map;
	if (aLiteralMap)
	{
		// Caller's aLiteralMap starts at aLineText, so adjust it so that it starts at the newly
		// found position of action_args instead:
		literal_map = aLiteralMap + (action_args - aLineText);
	}
	else
	{
		literal_map = (LPTSTR)_alloca(sizeof(TCHAR) * line_length);
		ZeroMemory(literal_map, sizeof(TCHAR) * line_length);  // Must be fully zeroed for this purpose.
		// Resolve escaped sequences and make a map of which characters in the string should
		// be interpreted literally rather than as their native function.  In other words,
		// convert any escape sequences in order from left to right.  This part must be
		// done *after* checking for comment-flags that appear to the right of a valid line, above.
		// How literal comment-flags (e.g. semicolons) work:
		//string1; string2 <-- not a problem since string2 won't be considered a comment by the above.
		//string1 ; string2  <-- this would be a user mistake if string2 wasn't supposed to be a comment.
		//string1 `; string 2  <-- since esc seq. is resolved *after* checking for comments, this behaves as intended.
		ConvertEscapeSequences(action_args, literal_map);
	}

	// v2: All statements accept expressions by default, except goto/break/continue.
	if (aActionType > ACT_LAST_JUMP || aActionType < ACT_FIRST_JUMP)
		all_args_are_expressions = true;

	if (aActionType == ACT_FOR)
	{
		// Validate "For" syntax and translate to conventional command syntax.
		// "For x,y in z" -> "For x,y, z"
		// "For x in y"   -> "For x,, y"
		LPTSTR cp;
		for (cp = action_args; ; )
		{
			if (ctolower(cp[0]) == 'i' && ctolower(cp[1]) == 'n' && IS_SPACE_OR_TAB(cp[2]))
				break;
			if (!IS_LEADING_IDENTIFIER_CHAR(*cp) && *cp != g_delimiter)
				return ScriptError(ERR_EXPR_SYNTAX, aLineText);
			while (IS_IDENTIFIER_CHAR(*cp)) ++cp;
			if (*cp == g_delimiter) ++cp;
			while (IS_SPACE_OR_TAB(*cp)) ++cp;
		}
		// Replace "in" with a normal arg delimiter, or space if there are no vars.
		cp[0] = cp == action_args ? ' ' : g_delimiter;
		cp[1] = ' ';
	}
	else if (aActionType == ACT_CATCH)
	{
		// Add as 1 arg and parse later.
		all_args_are_expressions = false;
	}

	/////////////////////////////////////////////////////////////
	// Parse the parameter string into a list of separate params.
	/////////////////////////////////////////////////////////////
	// MaxParams has already been verified as being <= MAX_ARGS.
	int nArgs;
	LPTSTR arg[MAX_ARGS], arg_map[MAX_ARGS];
	int max_params = this_action.MaxParams;
	int max_params_minus_one = max_params - 1;
	bool is_expression;

	for (nArgs = mark = 0; action_args[mark] && nArgs < max_params; ++nArgs)
	{
		arg[nArgs] = action_args + mark;
		arg_map[nArgs] = literal_map + mark;

		is_expression = all_args_are_expressions; // This would be false for goto/break/continue.

		// Find the end of the above arg:
		if (is_expression)
			// Find the next delimiter, taking into account quote marks, parentheses, etc.
			mark = FindExprDelim(action_args, g_delimiter, mark, literal_map);
		else
			// Find the next non-literal delimiter.
			mark = FindTextDelim(action_args, g_delimiter, mark, literal_map);

		if (action_args[mark])  // A non-literal delimiter was found.
		{
			if (nArgs == max_params_minus_one) // There should be no delimiter, since this is the last arg.
			{
				++nArgs;
				break; // Let section below handle it as an error (if it's not ACT_EXPRESSION).
			}
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
	if (action_args[mark] && !(aActionType == ACT_EXPRESSION || aActionType == ACT_CATCH))
	{
		sntprintf(error_msg, _countof(error_msg), _T("\"%s\" accepts at most %d parameter%s.")
			, this_action.Name, max_params
			, max_params > 1 ? _T("s") : _T(""));
		return ScriptError(error_msg, aLineText);
	}
	for (int i = 0, non_blank_params = aActionType == ACT_CASE ? nArgs : aActionType == ACT_CATCH ? nArgs - 1 : this_action.MinParams
		; i < non_blank_params; ++i) // It's only safe to do this after the above.
		if (!*arg[i])
		{
			sntprintf(error_msg, _countof(error_msg), _T("\"%s\" requires that parameter #%u be non-blank.")
				, this_action.Name, i + 1);
			return ScriptError(error_msg, aLineText);
		}

	// Handle one-true-brace (OTB).
	// Else, Try and Finally were already handled, since they also accept a sub-action.
	if (nArgs && (ACT_IS_LOOP(aActionType) || aActionType == ACT_IF || aActionType == ACT_CATCH || aActionType == ACT_SWITCH)
		&& !add_openbrace_afterward) // It wasn't already handled.
	{
		LPTSTR arg1 = arg[nArgs - 1]; // For readability and possibly performance.
		LPTSTR arg1_last_char = arg1 + _tcslen(arg1) - 1;
		if (arg1_last_char >= arg1 // Checked for maintainability.  Probably unnecessary since if true, arg1_last_char would be the null-terminator of the previous arg (never '{').
			&& *arg1_last_char == '{') // Relies on short-circuit boolean evaluation.
		{
			add_openbrace_afterward = true;
			*arg1_last_char = '\0';  // Since it will be fully handled here, remove the brace from further consideration.
			if (!rtrim(arg1)) // Trimmed down to nothing, so only a brace was present: remove the arg completely.
				if (aActionType == ACT_LOOP || aActionType == ACT_CATCH || aActionType == ACT_SWITCH)
					--nArgs;
				else // ACT_WHILE, ACT_FOR or ACT_IF
					return ScriptError(ERR_PARAM1_REQUIRED, aLineText);
		}
	}

	if (aActionType == ACT_SWITCH && nArgs && !*arg[0]) // Switch with parameters - do this only after the above, since it might remove one arg.
		ScriptError(ERR_PARAM1_MUST_NOT_BE_BLANK); // "Switch , y" or "Switch , "

	if (!AddLine(aActionType, arg, nArgs, arg_map, all_args_are_expressions))
		return FAIL;
	if (add_openbrace_afterward)
		if (!AddLine(ACT_BLOCK_BEGIN))
			return FAIL;
	if (!subaction_start) // There is no subaction in this case.
		return OK;
	return ParseAndAddLine(subaction_start); // Escape sequences in the subaction haven't been translated yet.
}



inline LPTSTR Script::ParseActionType(LPTSTR aBufTarget, LPTSTR aBufSource, bool aDisplayErrors)
// inline since it's called so often.
// aBufTarget should be at least MAX_VAR_NAME_LENGTH + 1 in size.
// Returns NULL if it's not a named action; otherwise, the address of the first
// character after the action name in aBufSource (possibly the null-terminator).
{
	////////////////////////////////////////////////////////
	// Find the action name and the start of the param list.
	////////////////////////////////////////////////////////
	LPTSTR end_marker = find_identifier_end(aBufSource);
	size_t action_name_length = end_marker - aBufSource;
	if (!action_name_length || action_name_length > MAX_VAR_NAME_LENGTH)
	{
		// For the "too long" case, it will ultimately be treated as an unrecognized action.
		*aBufTarget = '\0';
		return NULL;
	}
	tcslcpy(aBufTarget, aBufSource, action_name_length + 1);
	return end_marker;
}



inline ActionTypeType Script::ConvertActionType(LPCTSTR aActionTypeString)
// inline since it's called so often, but don't keep it in the .h due to #include issues.
{
	// For the loop's index:
	// Use an int rather than ActionTypeType since it's sure to be large enough to go beyond
	// 256 if there happen to be exactly 256 actions in the array:
 	for (int action_type = ACT_FIRST_NAMED_ACTION; action_type <= ACT_LAST_NAMED_ACTION; ++action_type)
		if (!_tcsicmp(aActionTypeString, g_act[action_type].Name)) // Match found.
			return action_type;
	return ACT_INVALID;  // On failure to find a match.
}



ResultType Script::AddLine(ActionTypeType aActionType, LPTSTR aArg[], int aArgc, LPTSTR aArgMap[], bool aAllArgsAreExpressions)
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
	if (mLastHotFunc)
	{
		if (aActionType != ACT_BLOCK_BEGIN)
			return ScriptError(ERR_HOTKEY_MISSING_BRACE);
		mLastHotFunc = nullptr;
	}
	if (aActionType == ACT_BLOCK_BEGIN && mIgnoreNextBlockBegin)
	{
		mIgnoreNextBlockBegin = false;
		return OK;
	}

	DerefList deref;  // Will be used to temporarily store the var-deref locations in each arg.
	ArgStruct *new_arg;  // We will allocate some dynamic memory for this, then hang it onto the new line.
	LPTSTR this_aArgMap, this_aArg;

	//////////////////////////////////////////////////////////
	// Build the new arg list in dynamic memory.
	// The allocated structs will be attached to the new line.
	//////////////////////////////////////////////////////////
	if (!aArgc)
		new_arg = NULL;  // Just need an empty array in this case.
	else
	{
		new_arg = SimpleHeap::Alloc<ArgStruct>(aArgc);

		int i;

		for (i = 0; i < aArgc; ++i)
		{
			////////////////
			// FOR EACH ARG:
			////////////////
			this_aArg = aArg[i];                        // For performance and convenience.
			this_aArgMap = aArgMap ? aArgMap[i] : NULL; // Same.
			ArgStruct &this_new_arg = new_arg[i];       // Same.
			this_new_arg.postfix = NULL;                // Set default early, for maintainability.

			// Determine whether this arg is ARG_TYPE_NORMAL or ARG_TYPE_OUTPUT_VAR.  Blank args are
			// set to ARG_TYPE_NORMAL for maintainability, so ARG_TYPE_OUTPUT_VAR always has a non-null
			// var at runtime.  ARG_TYPE_INPUT_VAR is never encountered at this stage; it is only used
			// by certain optimizations in ExpressionToPostfix(), which comes later.
			this_new_arg.type = *this_aArg ? Line::ArgIsVar(aActionType, i, aArgc) : ARG_TYPE_NORMAL;
			
			// All non-blank args are handled by ExpressionToPostfix, except where caller has indicated
			// this action type does not accept an expression (e.g. "goto label").
			this_new_arg.is_expression = *this_aArg && aAllArgsAreExpressions;

			// So that it can be passed to Alloc(), first update the length to match what the text will be
			// (if the alloc fails, an inaccurate length won't matter because it's a program-abort situation).
			// The length member was added in v1.0.44.14 to boost runtime performance, but is currently only
			// used to estimate how much deref buf space is needed; see EXPR_BUF_SIZE macro.
			this_new_arg.length = (ArgLengthType)_tcslen(this_aArg);
			
			// Create a copy of arg text in persistent memory.
			this_new_arg.text = SimpleHeap::Alloc(this_aArg, this_new_arg.length);

			///////////////////////////////////////////
			// Build the list of operands for this arg.
			///////////////////////////////////////////
			// Now that any escaped quote marks etc. have been marked, pre-parse all operands
			// in this_new_arg.text.  Review is needed to determine whether var derefs still
			// need to be resolved at this stage (and whether this is optimal).  Other operands
			// are marked at this stage to avoid some redundant processing later, and so that we
			// don't need to make this_aArgMap persistent.
			// Note: this_new_arg.text is scanned rather than this_aArg because we want to
			// establish pointers to the correct area of memory:
			deref.count = 0;  // Init for each arg.

			if (this_new_arg.is_expression)
			{
				if (!ParseOperands(this_new_arg.text, this_aArgMap, deref))
					return FAIL;
			}
			// Otherwise, this arg does not contain an expression and therefore cannot contain
			// any derefs.  This is currently limited to ACT_GOTO/ACT_BREAK/ACT_CONTINUE,
			// which require either a literal label name or () to force an expression.

			//////////////////////////////////////////////////////////////
			// Allocate mem for this arg's list of dereferenced variables.
			//////////////////////////////////////////////////////////////
			if (deref.count)
			{
				this_new_arg.deref = SimpleHeap::Alloc<DerefType>(deref.count + 1); // +1 for the "NULL-item" terminator.
				memcpy(this_new_arg.deref, deref.items, deref.count * sizeof(DerefType));
				// Terminate the list of derefs with a deref that has a NULL marker:
				this_new_arg.deref[deref.count].marker = NULL;
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
	// v2.0: Currently almost none of the v1 validation is done here, since built-in functions
	// are now fully parsed as expressions and not as unique action types.  This old method
	// was never applicable to expressions/functions, which are more complicated to validate.
	// This should be replaced with better tools (separate to the program), such as syntax
	// checkers which can operate while the user is editing, before they try to run the script.
	if (aActionType == ACT_RETURN)
	{
		if (aArgc > 0 && !g->CurrentFunc)
			return ScriptError(_T("Return's parameter should be blank except inside a function."));
	}

	// This next section and the block-begin/end sections below are responsible
	// for setting up the relationships of lines (mParentLine and mRelatedLine).
	// This was previously done as a separate pass, but doing it here enables
	// all subsequent stages to rely on the relationships already being defined.
	// It also reduced code size.
	switch (aActionType)
	{
	case ACT_CASE:
		if (!(mOpenBlock && mOpenBlock->mPrevLine && mOpenBlock->mPrevLine->mActionType == ACT_SWITCH)
			|| mPendingParentLine)
			return line.LineUnexpectedError();
		Line *last_case;
		if ((last_case = mOpenBlock->mNextLine) && last_case != &line)
		{
			// Form a linked list of CASE lines for this SWITCH.
			while (last_case->mRelatedLine)
				last_case = last_case->mRelatedLine;
			last_case->mRelatedLine = &line;
		}
		break;
	case ACT_ELSE:
	case ACT_UNTIL:
	case ACT_CATCH:
	case ACT_FINALLY:
		bool expected = false;
		Line *parent = mPendingRelatedLine;
		for (;; parent = parent->mParentLine)
		{
			if (!parent)
				return line.LineUnexpectedError();
			enum_act parent_act = (enum_act)parent->mActionType;
			switch (aActionType)
			{
			case ACT_ELSE:
				if (parent_act == ACT_TRY)
					// Prevent the Else in Try..Else from attaching to a previous statement, since some users
					// might expect Try..Else to be synonmous with Try..Catch Error..Else (which is currently
					// disabled for code size, and because "try x else y" sounds more like y could be executed
					// if x was NOT successful).
					return line.LineUnexpectedError();
				expected = parent_act == ACT_IF || parent_act == ACT_CATCH || ACT_IS_LOOP(parent_act);
				break;
			case ACT_UNTIL: expected = ACT_IS_LOOP_EXCLUDING_WHILE(parent_act); break;
			case ACT_FINALLY:
				if (parent_act == ACT_ELSE)
				{
					Line *else_line = parent; // But what kind of ELSE?
					parent = parent->mPrevLine; // This is potentially the last line of a compound statement above the ELSE.
					do // Find the IF/CATCH at the same level as the ELSE.
						parent = parent->mParentLine;
					while (parent->mParentLine != else_line->mParentLine);
					expected = parent->mActionType == ACT_CATCH;
					break;
				}
				// Fall through:
			case ACT_CATCH:
				expected = parent_act == ACT_TRY || parent_act == ACT_CATCH;
				break;
			}
			if (expected)
				break;
		}
		line.mParentLine = parent->mParentLine; // Not parent itself, since an ELSE can't jump into its IF's body.
		break;
	}
	while (mPendingRelatedLine) // A completed block or a parent line with a single-line action, waiting for its mRelatedLine.
	{
		bool break_time = aActionType == ACT_UNTIL && ACT_IS_LOOP_EXCLUDING_WHILE(mPendingRelatedLine->mActionType); // Until must "relate to" only a single statement.
		mPendingRelatedLine->mRelatedLine = &line;
		// Regardless of whether mPendingRelatedLine is a block-begin or parent line,
		// its own parent also needs its mRelatedLine set, unless that's an open block.
		mPendingRelatedLine = mPendingRelatedLine->mParentLine;
		if (mPendingRelatedLine && mPendingRelatedLine->mActionType == ACT_BLOCK_BEGIN)
			mPendingRelatedLine = nullptr;
		if (break_time)
			break;
	}
	if (mPendingParentLine) // A line waiting for block or single-line action.
	{
		line.mParentLine = mPendingParentLine;
		if (aActionType != ACT_BLOCK_BEGIN)
			mPendingRelatedLine = mPendingParentLine; // The next line will be the parent's mRelatedLine.
		mPendingParentLine = nullptr;
	}
	if (ACT_IS_LINE_PARENT(aActionType))
	{
		mPendingParentLine = &line;
		mPendingRelatedLine = nullptr;
	}
	// If no parent line was determined above, this line belongs to the current block.
	if (!line.mParentLine)
		line.mParentLine = mOpenBlock;

	if (mNextLineIsFunctionBody)
	{
		g->CurrentFunc->mJumpToLine = the_new_line;
		mNextLineIsFunctionBody = false;
	}

	// Now that we've finished using the previous value of mOpenBlock, check the following:
	if (aActionType == ACT_BLOCK_BEGIN)
	{
		mOpenBlock = the_new_line;
		// It's only necessary to check the last func, not the one(s) that come before it, to see if its
		// mJumpToLine is NULL.  This is because our caller has made it impossible for a function
		// to ever have been defined in the first place if it lacked its opening brace.  Search on
		// "consecutive function" for more comments.
		if (g->CurrentFunc && !g->CurrentFunc->mJumpToLine)
		{
			line.mAttribute = g->CurrentFunc;  // Flag this ACT_BLOCK_BEGIN as the opening brace of the function's body.
			// For efficiency, and to prevent ExecUntil from starting a new recursion layer for the function's
			// body, the function's execution should begin at the first line after its open-brace (even if that
			// first line is another open-brace or the function's close-brace (i.e. an empty function):
			mNextLineIsFunctionBody = true;
		}
	}
	// See further below for ACT_BLOCK_END.

	// Above must be done prior to the below, since it sometimes sets mAttribute for use below.

	///////////////////////////////////////////////////////////////
	// Update any labels that should refer to the newly added line.
	///////////////////////////////////////////////////////////////
	// If the label most recently added doesn't yet have an anchor in the script, provide it.
	// UPDATE: In addition, keep searching backward through the labels until a non-NULL
	// mJumpToLine is found.  All the ones with a NULL should point to this new line to
	// support cases where one label immediately follows another in the script.
	if (!mNoUpdateLabels)
	{
		for (Label *label = g->CurrentFunc ? g->CurrentFunc->mLastLabel : mLastLabel;
			label != NULL && label->mJumpToLine == NULL; label = label->mPrevLabel)
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

	// Above must be done prior to the below, because otherwise a function-local label's mJumpToLine
	// would not be set in a case like the following:
	// Func(){
	//    ...
	//    local_label:
	// }

	if (aActionType == ACT_BLOCK_END)
	{
		if (!mOpenBlock || line.mParentLine != mOpenBlock)
			return line.LineUnexpectedError();
		mPendingRelatedLine = mOpenBlock; // The next line will be the block-begin's mRelatedLine, and possibly its parent's mRelatedLine.

		// Rather than checking whether each newly added line is being added into a Switch,
		// just verify here that the first line of each Switch is valid:
		if (mOpenBlock->mPrevLine && mOpenBlock->mPrevLine->mActionType == ACT_SWITCH
			&& mOpenBlock->mNextLine->mActionType != ACT_CASE
			&& mOpenBlock->mNextLine != &line)
			return mOpenBlock->mNextLine->LineError(_T("Expected Case/Default."));

		if (g->CurrentFunc && g->CurrentFunc == mOpenBlock->mAttribute)
		{
			line.mAttribute = g->CurrentFunc;  // Flag this ACT_BLOCK_END as the ending brace of this function's body.
			g->CurrentFunc = g->CurrentFunc->mOuterFunc;  // Step out of this function.
			if (g->CurrentFunc && !g->CurrentFunc->mJumpToLine)
				// The outer function has no body yet, so it probably began with one or more nested functions.
				mNextLineIsFunctionBody = true;
		}
		do
			mOpenBlock = mOpenBlock->mParentLine;
		while (mOpenBlock && mOpenBlock->mActionType != ACT_BLOCK_BEGIN);
	}

	return OK;
}



ResultType DerefList::Push()
{
	if (count == size)
	{
		const int block_size = 128; // In most cases one allocation will be enough.
		DerefType *p = (DerefType *)realloc(items, (size + block_size) * sizeof(DerefType));
		if (!p)
			return MemoryError();
		items = p;
		size += block_size;
	}
	++count;
	return OK;
}



ResultType Script::ParseOperands(LPTSTR aArgText, LPTSTR aArgMap, DerefList &aDeref, int *aPos, TCHAR aEndChar)
{
	LPTSTR op_begin, op_end;
	size_t operand_length;
	TCHAR close_char, *cp;
	int j;
	bool is_function;
	bool is_double_deref;
	SymbolType wordop;

	// Escape sequences outside of literal strings have no meaning, so could be considered illegal
	// to improve error detection and reserve ` for future use.  However, in that case ` must still
	// be allowed inside quoted strings and for `) in continuation sections, so this is inadequate:
	//if (this_aArgMap) // This arg has an arg map indicating which chars are escaped/literal vs. normal.
	//	for (j = 0; this_new_arg.text[j]; ++j)
	//		if (this_aArgMap[j])
	//			return ScriptError(ERR_EXP_ILLEGAL_CHAR, this_new_arg.text + j);

	op_begin = aArgText;
	if (aPos)
		op_begin += *aPos;
	for (; *op_begin; op_begin = op_end)
	{
		while (*op_begin && _tcschr(EXPR_OPERAND_TERMINATORS, *op_begin)) // Skip over whitespace, operators, and parentheses.
		{
			if (*op_begin == aEndChar)
    		{
    			if (aPos)
    				*aPos = int(op_begin - aArgText);
    			return OK;
    		}
			if (*op_begin == g_DerefChar)
				// This is the start of a double-deref with no initial literal part.  Double-derefs are
				// handled outside the loop for simplicity.  Breaking out of the loop means that this
				// will be seen as an operand of zero length.
				break;
			if (*op_begin == '.')
			{
				// This case also handles the dot in `.=` (the `=` is skipped by the next iteration).
				// Skip the numeric literal or identifier to the right of this, if there is one.
				// This won't skip the signed exponent of a scientific-notation literal, but that should
				// be OK as it will be recognized as purely numeric in the next iteration of this loop.
				op_end = find_identifier_end(op_begin + 1);
				if (*op_end != g_DerefChar || aEndChar == g_DerefChar)
				{
					op_begin = op_end;
					continue;
				}
				// Otherwise, it's a dynamic property or method call.
				++op_begin; // Skip the dot.
				break;
			}
			switch (*op_begin)
			{
			default:
				op_begin++;
				continue;

			case '\n':
				// Allow `n unquoted in expressions so that continuation sections can be used to
				// spread an expression across multiple lines without requiring the "Join" option;
				// but replace them with spaces now so they don't need to be handled later on.
				*op_begin++ = ' ';
				continue;

			// These errors are detected by GetLineContExpr()/BalanceExpr() or later by ExpressionToPostfix:
			//case ')':
			//case ']':
			//case '}':
			//	return ScriptError(ERR_EXPR_SYNTAX, op_begin);
			case '(': close_char = ')'; break;
			case '[': close_char = ']'; break;
			case '{': close_char = '}'; break;
			}
			// Since above didn't "continue", recurse to handle nesting:
			j = (int)(op_begin - aArgText + 1);
			if (close_char == ')')
			{
				// Before parsing derefs, check if this is the parameter list of an inline function.
				LPTSTR close_paren = aArgText + FindExprDelim(aArgText, close_char, j, aArgMap);
				if (*close_paren)
				{
					cp = omit_leading_whitespace(close_paren + 1);
					if (*cp == '=' && cp[1] == '>') // () => Fat arrow function.
					{
						if (aDeref.count)
						{
							DerefType &d = *aDeref.Last();
							if (d.type == DT_VAR && d.marker + d.length == op_begin)
								op_begin = d.marker;
						}
						if (!ParseFatArrow(aArgText, aArgMap, aDeref, op_begin, close_paren, cp + 2, op_begin))
							return FAIL;
						continue;
					}
				}
			}
			if (!ParseOperands(aArgText, aArgMap, aDeref, &j, close_char))
				return FAIL;
			op_begin = aArgText + j;
			if (*op_begin)
				op_begin++;
			//else:
				// Missing close paren/bracket/brace. Let ExpressionToPostfix() handle it.
		}
		if (!*op_begin) // The above loop reached the end of the string: No operands remaining.
			break;
		// Now op_begin is the start of an operand, which might be a variable reference, a numeric
		// literal, or a string literal.

		if (*op_begin == '"' || *op_begin == '\'')
		{
			op_end = aArgText + FindTextDelim(aArgText, *op_begin, int(op_begin - aArgText + 1), aArgMap);
			if (!*op_end)
				return ScriptError(ERR_MISSING_CLOSE_QUOTE, op_begin);
			if (!aDeref.Push())
				return FAIL;
			DerefType &this_deref = *aDeref.Last();
			this_deref.type = DT_QSTRING;
			this_deref.terminal = true;
			this_deref.marker = op_begin + 1;
			this_deref.length = DerefLengthType(op_end - (op_begin + 1));
			this_deref.substring_count = 1;
			++op_end;
			// op_end is now set correctly to allow the outer loop to continue.
			continue;
		}

		if (*op_begin <= '9' && *op_begin >= '0') // Numeric literal.  Numbers starting with a decimal point are handled under "case '.'".
		{
			// Behaviour here should match the similar section in ExpressionToPostfix().
			istrtoi64(op_begin, &op_end);
			if (!IS_HEX(op_begin)) // This check is probably only needed on VC++ 2015 and later, where _tcstod allows hex.
			{
				LPTSTR d_end;
				_tcstod(op_begin, &d_end);
				if (op_end < d_end && _tcschr(EXPR_OPERAND_TERMINATORS, *d_end))
				{
					op_end = d_end;
					continue;
				}
			}
			if (_tcschr(EXPR_OPERAND_TERMINATORS, *op_end))
			{
				// Do nothing further since pure numbers don't need any processing at this stage.
				// If this number has a fractional part, it is handled by "case '.'" above.
				continue;
			}
			//else: It's not valid, but let it pass through to Var::ValidateName() to generate an error message.
		}

		// Find the end of this operand (the position immediately after the operand):
		op_end = find_identifier_end(op_begin);
		// Now op_end marks the end of this operand.  The end might be the zero terminator, an operator, etc.
		operand_length = op_end - op_begin;

		if (is_double_deref = (*op_end == g_DerefChar && aEndChar != g_DerefChar))
		{
			// This operand is the leading literal part of a double dereference.
			j = (int)(op_begin - aArgText);
			if (!ParseDoubleDeref(aArgText, aArgMap, aDeref, &j))
				return FAIL;
			op_end = aArgText + j;
			is_function = *op_end == '('; // Dynamic function call.
		}
		else if (!operand_length) // Found an illegal char.
		{
			// Due to other checks above, this should be possible only when *op_end is an illegal
			// character (since it is not null, an operand terminator or identifier character).
			// All characters are permitted in quoted strings, which were already handled above.
			// An invalid identifier such as "bad\01" is processed as a valid var "bad" followed
			// by a zero-length operand (triggering this section).
			return ScriptError(ERR_EXP_ILLEGAL_CHAR, op_end);
		}
		else
		{
			// Before checking for word operators, check if this is the "key" in {key: value}.
			// Things like {new: true} and obj.new are allowed because there is no ambiguity.
			// This is done after the check for '.' above so that {x.y:z} is parsed as {(x.y):z}.
			cp = omit_leading_whitespace(op_end);
			if (*cp == ':' && cp[1] != '=') // Not :=
			{
				cp = omit_trailing_whitespace(aArgText, op_begin - 1);
				if (*cp == ',' || *cp == '{')
				{
					// This is either the key in a key-value pair in an object literal, or a syntax
					// error which will be caught at a later stage (since the ':' is missing its '?').
					continue; // Leave it unmarked so ExpressionToPostfix() handles it as a literal.
				}
			}
			else if (*cp == '=' && cp[1] == '>') // one_param => Fat arrow function.
			{
				if (!ParseFatArrow(aArgText, aArgMap, aDeref, op_begin, op_end, cp + 2, op_end))
					return FAIL;
				continue;
			}
			is_function	= *op_end == '(';
		}

		if (!aDeref.Push())
			return FAIL;
		auto &this_deref = *aDeref.Last();

		if (is_double_deref)
		{
			// Mark the end of this dynamic reference.
			this_deref.marker = op_end;
			this_deref.length = 0;
			if (op_begin > aArgText && op_begin[-1] == '.')
				this_deref.type = DT_DOTPERCENT;
			else
				this_deref.type = DT_DOUBLE;
		}
		else if (  operand_length < 9 && (wordop = ConvertWordOperator(op_begin, operand_length))  )
		{
			// It's a word operator like AND/OR/NOT.
			if (SYM_IS_RESERVED(wordop))
			{
				return ScriptError(_T("Unexpected reserved word."), op_begin);
			}
			// Mark this word as an operator.  Unlike the old method of replacing "OR" with "||",
			// this leaves ListLines more accurate.  More importantly, it allows "Hotkey, If" to
			// recognize an expression which uses AND/OR.  Additionally, the "NEW" operator will
			// require a DerefType struct in the postfix expression phase anyway.
			this_deref.marker = op_begin;
			this_deref.length = (DerefLengthType)operand_length;
			this_deref.type = DT_WORDOP;
			this_deref.symbol = wordop;
		}
		else if (operand_length == 4 && !_tcsnicmp(op_begin, _T("True"), 4)
			|| operand_length == 5 && !_tcsnicmp(op_begin, _T("False"), 5))
		{
			this_deref.marker = op_begin;
			this_deref.length = (DerefLengthType)operand_length;
			this_deref.type = DT_CONST_INT;
			this_deref.int_value = toupper(*op_begin) == 'T';
		}
		else // This operand is a variable name or function name.
		{
			// Store the deref's starting location, even for functions (leave it set to the start
			// of the function's name for use when doing error reporting at other stages -- i.e.
			// don't set it to the address of the first param or closing-paren-if-no-params):
			this_deref.marker = op_begin;
			this_deref.length = (DerefLengthType)operand_length;
			this_deref.type = DT_VAR;
			this_deref.var = nullptr; // To be resolved later.
		}
	}
	if (aPos)
    	*aPos = int(op_begin - aArgText);
	return OK;
}


ResultType Script::ParseDoubleDeref(LPTSTR aArgText, LPTSTR aArgMap, DerefList &aDeref, int *aPos)
{
	LPTSTR op_begin, dd_begin;
	LPTSTR op_end;
	int count = 0;
	for (op_begin = dd_begin = aArgText + *aPos; ; op_begin = op_end + 1)
	{
		op_end = find_identifier_end(op_begin);
		
		// String "derefs" are needed to pair up to the '%' markers, allowing both
		// to be optimized out if the string is empty (as on either side of %a%).
		if (!aDeref.Push())
			return FAIL;
		DerefType &this_deref = *aDeref.Last();
		this_deref.type = DT_STRING;
		this_deref.terminal = *op_end != g_DerefChar;
		this_deref.marker = op_begin;
		this_deref.length = DerefLengthType(op_end - op_begin);
		this_deref.substring_count = ++count;

		*aPos = int(op_end - aArgText);
		if (this_deref.terminal)
			break;
		(*aPos)++;
		if (!ParseOperands(aArgText, aArgMap, aDeref, aPos, g_DerefChar))
			return FAIL;
		op_end = aArgText + *aPos;
		if (*op_end != g_DerefChar)
			return ScriptError(_T("Missing ending \"%\""), op_begin);
	}
	return OK;
}


SymbolType Script::ConvertWordOperator(LPCTSTR aWord, size_t aLength)
{
	struct WordOp
	{
		LPCTSTR word;
		SymbolType op;
	};
	static WordOp sWordOp[] =
	{
		{ _T("or"), SYM_OR },
		{ _T("and"), SYM_AND },
		{ _T("not"), SYM_LOWNOT },
		{ _T("is"), SYM_IS },
		{ _T("IsSet"), SYM_ISSET },
		{ SUPER_KEYWORD, SYM_SUPER },
		{ _T("as"), SYM_RESERVED_OPERATOR },
		{ _T("contains"), SYM_RESERVED_OPERATOR },
		{ _T("in"), SYM_RESERVED_OPERATOR },
		{ _T("unset"), SYM_MISSING },
	};
	for (int i = 0; i < _countof(sWordOp); ++i)
	{
		if (!_tcsnicmp(sWordOp[i].word, aWord, aLength) && !sWordOp[i].word[aLength])
		{
			return sWordOp[i].op;
		}
	}
	return (SymbolType)FALSE;
}


ResultType Script::ParseFatArrow(LPTSTR aArgText, LPTSTR aArgMap, DerefList &aDeref
	, LPTSTR aPrmStart, LPTSTR aPrmEnd, LPTSTR aExpr, LPTSTR &aExprEnd)
{
	int expr_start_pos = int(aExpr - aArgText);
	int expr_end_pos = FindExprDelim(aArgText, 0, expr_start_pos, aArgMap);
	if (!ParseFatArrow(aDeref, aPrmStart, aPrmEnd, aExpr, aArgText + expr_end_pos, aArgMap + expr_start_pos))
		return FAIL;
	aExprEnd = aArgText + expr_end_pos;
	return OK;
}


ResultType Script::ParseFatArrow(DerefList &aDeref, LPTSTR aPrmStart, LPTSTR aPrmEnd, LPTSTR aExpr, LPTSTR aExprEnd, LPTSTR aExprMap)
{
	TCHAR orig_end;

	// Avoid pointing any pending labels at the fat arrow function's body (let the caller
	// finish adding the line which contains this expression and make it the label's target).
	bool nolabels = mNoUpdateLabels; // Could be true if this line is a static declaration.
	Line *parent_line = mPendingParentLine;
	mNoUpdateLabels = true;

	if (*aPrmEnd == ')') // `() => e` or `fn() => e`, not `v => e`.
	{
		orig_end = aPrmEnd[1];
		aPrmEnd[1] = '\0';
		if (!DefineFunc(aPrmStart, false, true))
			return FAIL;
		aPrmEnd[1] = orig_end;
	}
	else // aPrmStart is a variable name and aPrmEnd is the next character after it.
	{
		// Format the parameter list as needed for DefineFunc().
		TCHAR prm[MAX_VAR_NAME_LENGTH + 4];
		sntprintf(prm, _countof(prm), _T("(%.*s)"), aPrmEnd - aPrmStart, aPrmStart);
		if (!DefineFunc(prm, false, true))
			return FAIL;
	}

	if (!AddLine(ACT_BLOCK_BEGIN))
		return FAIL;
	auto block_begin = mLastLine;

	for (; aExprEnd > aExpr && IS_SPACE_OR_TAB(aExprEnd[-1]); --aExprEnd);
	orig_end = *aExprEnd;
	*aExprEnd = '\0';
	if (!ParseAndAddLine(aExpr, ACT_RETURN, aExprMap, aExprEnd - aExpr))
		return FAIL;
	*aExprEnd = orig_end;

	auto name = g->CurrentFunc->mName; // Must be done before AddLine(ACT_BLOCK_END).

	if (!AddLine(ACT_BLOCK_END))
		return FAIL;

	if (!*name) // Anonymous function is not yet preceded by a DT_VAR.
	{
		if (!aDeref.Push())
			return FAIL;
	}
	else // Replace the DT_VAR preceding this fat arrow function with DT_FUNCREF.
	{
		ASSERT(aDeref.count && aDeref.Last()->type == DT_VAR && aDeref.Last()->marker == aPrmStart);
	}
	auto &fr = *aDeref.Last();
	fr.type = DT_FUNCREF;
	fr.marker = aPrmStart; // Mark the entire fat arrow expression as a function deref.
	fr.length = DerefLengthType(aExprEnd - aPrmStart); // Relies on the fact that an arg can't be longer than the max deref length (because ArgLengthType == DerefLengthType).
	if (*name)
		fr.var = FindVar(name, 0); // Find the Var already added by AddFunc().
	else // No name, so AddFunc() inserted it at position 0.
		fr.var = g->CurrentFunc ? g->CurrentFunc->mVars.mItem[0] : GlobalVars()->mItem[0];

	mNoUpdateLabels = nolabels;
	if (parent_line)
	{
		mPendingParentLine = parent_line;
		mPendingRelatedLine = block_begin; // This ensures block_begin->mRelatedLine is set for PreparseExpressions.
	}

	return OK;
}


ResultType Script::DefineFunc(LPTSTR aBuf, bool aStatic, bool aIsInExpression)
// Returns OK or FAIL.
// Caller has already called ValidateName() on the function, and it is known that this valid name
// is followed immediately by an open-paren.  aFuncExceptionVar is the address of an array on
// the caller's stack that will hold the list of exception variables (those that must be explicitly
// declared as either local or global) within the body of the function.
{
	LPTSTR param_end, param_start = _tcschr(aBuf, '('); // Caller has ensured that this will return non-NULL.
	int insert_pos;
	
	auto last_hotfunc = mLastHotFunc;
	if (last_hotfunc)
	{
		// This means we are defining a function under a "trigger::".
		// Then mLastHotFunc will not be used, instead all variants which have set it as its
		// mJumpLabel will replace it with the new function defined in this call.
		ASSERT(last_hotfunc == g->CurrentFunc);
		g->CurrentFunc =					// To avoid nesting the new function inside it.
			mLastHotFunc = nullptr;			// To avoid the next call AddLine demanding a ACT_BLOCK_BEGIN.										
	}

	bool is_method = mClassObjectCount && !g->CurrentFunc && !aIsInExpression;
	if (is_method) // Class method or property getter/setter.
	{
		Object *class_object = mClassObject[mClassObjectCount - 1];
		if (!aStatic)
			class_object = (Object *)class_object->GetOwnPropObj(_T("Prototype"));

		*param_start = '\0'; // Temporarily terminate, for simplicity.

		// Build the fully-qualified method name for A_ThisFunc and ListVars:
		// AddFunc() enforces a limit of MAX_VAR_NAME_LENGTH characters for function names.
		// For simplicity, allow one extra char to be printed and rely on AddFunc() detecting
		// that the name is too long.
		TCHAR full_name[MAX_VAR_NAME_LENGTH + 1 + 1]; // Extra +1 for null terminator.
		_sntprintf(full_name, MAX_VAR_NAME_LENGTH + 1, aStatic ? _T("%s.%s") : _T("%s.Prototype.%s"), mClassName, aBuf);
		full_name[MAX_VAR_NAME_LENGTH + 1] = '\0'; // Must terminate at this exact point if _sntprintf hit the limit.
		
		*param_start = '('; // Undo temporary termination.

		// Below passes class_object for AddFunc() to store the func "by reference" in it:
		if (  !(g->CurrentFunc = AddFunc(full_name, -1, class_object))  )
			return FAIL;
	}
	else
	{
		// The value of g->CurrentFunc must be set here rather than by our caller since AddVar(), which we call,
		// relies upon it having been done.
		if (   !(g->CurrentFunc = AddFunc(aBuf, param_start - aBuf))   )
			return FAIL; // It already displayed the error.
	}

	// mNextLineIsFunctionBody can be true in cases where a function begins with a nested function.
	// Reset it to false for this new function prior to any AddLine() calls; it will be set back to
	// true if appropriate when the block-end of this function is found.
	mNextLineIsFunctionBody = false;

	auto &func = *g->CurrentFunc; // For performance and convenience.

	size_t param_length, value_length;
	FuncParam param[MAX_FUNCTION_PARAMS];
	int param_count = 0;
	TCHAR buf[LINE_SIZE];
	bool param_must_have_default = false;
	bool at_least_one_default_expr = false;

#ifdef CONFIG_DLL
	func.mLineNumber = mCombinedLineNumber;
	func.mFileIndex = mCurrFileIndex;
#endif

	func.mIsFuncExpression = aIsInExpression;

	if (is_method)
	{
		// Add the automatic/hidden "this" parameter.
		if (  !(param[0].var = AddVar(_T("this"), 4, &func.mVars, 0, VAR_DECLARE_LOCAL | VAR_LOCAL_FUNCPARAM))  )
			return FAIL;
		param[0].default_type = PARAM_DEFAULT_NONE;
		param[0].is_byref = false;
		++param_count;
		++func.mMinParams;

		if (mClassProperty && toupper(param_start[-3]) == 'S') // Set
		{
			// Unlike __Set, the property setters accept the value as the second parameter,
			// so that we can make it automatic here and make variadic properties easier.
			if (  !(param[1].var = AddVar(_T("value"), 5, &func.mVars, 1, VAR_DECLARE_LOCAL | VAR_LOCAL_FUNCPARAM))  )
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

		if (param_count >= MAX_FUNCTION_PARAMS)
			return ScriptError(_T("Too many params."), param_start); // Short msg since so rare.
		FuncParam &this_param = param[param_count]; // For performance and convenience.

		if (this_param.is_byref = (*param_start == '&')) // ByRef.
			param_start = omit_leading_whitespace(param_start + 1);
		
		param_end = find_identifier_end(param_start);
		param_length = param_end - param_start;
		if (param_length)
		{
			VarList *varlist; // FindVar() should always set this to &func.mVars, but best to pass it for maintainability.
			if (this_param.var = FindVar(param_start, param_length, FINDVAR_LOCAL, &varlist, &insert_pos))
				return ConflictingDeclarationError(_T("parameter"), this_param.var);
			if (   !(this_param.var = AddVar(param_start, param_length, varlist, insert_pos, VAR_DECLARE_LOCAL | VAR_LOCAL_FUNCPARAM))   )	// Pass VAR_LOCAL_FUNCPARAM as last parameter to mean "it's a local but more specifically a function's parameter".
				return FAIL; // It already displayed the error, including attempts to have reserved names as parameter names.
			param_start = omit_leading_whitespace(param_end);
		}
		else
		{
			if (*param_start != '*')
				return ScriptError(ERR_MISSING_PARAM_NAME, aBuf); // Reporting aBuf vs. param_start seems more informative since Vicinity isn't shown.
			// fn(a,*) permits surplus parameters but does not store them.
			this_param.var = NULL;
		}
		
		this_param.default_type = PARAM_DEFAULT_NONE;  // Set default.

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

		if (*param_start == '=') // This will probably be a common error for users transitioning from v1.
			return ScriptError(ERR_BAD_OPTIONAL_PARAM, param_start);

		// v1.0.35: Check if a default value is specified for this parameter and set up for the next iteration.
		// The following section is similar to that used to support initializers for static variables.
		// So maybe maintain them together.
		if (*param_start == ':') // This is the default value of the param just added.
		{
			if (param_start[1] != '=')
				return ScriptError(ERR_BAD_OPTIONAL_PARAM, param_start);

			param_start = omit_leading_whitespace(param_start + 2); // Start of the default value.
			if (*param_start == '"' || *param_start == '\'') // Quoted literal string, or the empty string.
			{
				// Find the end of this string literal, taking into account quote type and escape chars.
				param_end = param_start + FindTextDelim(param_start, *param_start, 1);
				if (*param_end) // Close-quote was found.
				{
					auto next_non_white = *omit_leading_whitespace(param_end + 1);
					if ((next_non_white == ',' || next_non_white == ')') // Just a literal string.
						&& param_end - param_start < _countof(buf)) // If it's too long for the buffer, let it raise an error in the next section.
					{
						tcslcpy(buf, param_start + 1, param_end - param_start); // +1 skips the opening quote and also accounts for the terminator for tcslcpy.
						ConvertEscapeSequences(buf, NULL); // Raw escape sequences like `n haven't been converted yet, so do it now.
						this_param.default_type = PARAM_DEFAULT_STR;
						this_param.default_str = SimpleHeap::Alloc(buf);
						++param_end; // Move beyond the close-quote, for below.
					}
				}
			}
			if (this_param.default_type != PARAM_DEFAULT_STR)
			{
				value_length = FindExprDelim(param_start, 0);
				if (value_length >= _countof(buf)) // Default expressions shouldn't realistically be longer than LINE_SIZE.
					return ScriptError(ERR_INVALID_FUNCDECL, aBuf);
				if (value_length <= MAX_NUMBER_SIZE)
				{
					tcslcpy(buf, param_start, value_length + 1); // Make a temp copy to simplify the below (especially IsNumeric).
					value_length = rtrim(buf, value_length); // trim white space for _tcsicmp below
				}
				else
					*buf = '\0'; // It can't be any of the optimized defaults, so skip the copy.
				if (!_tcsicmp(buf, _T("False")))
				{
					this_param.default_type = PARAM_DEFAULT_INT;
					this_param.default_int64 = 0;
				}
				else if (!_tcsicmp(buf, _T("True")))
				{
					this_param.default_type = PARAM_DEFAULT_INT;
					this_param.default_int64 = 1;
				}
				else if (!_tcsicmp(buf, _T("unset")))
				{
					this_param.default_type = PARAM_DEFAULT_UNSET;
				}
				else // The only things supported other than the above are integers and floats.
				{
					switch(IsNumeric(buf, true, false, true))
					{
					case PURE_INTEGER:
						this_param.default_type = PARAM_DEFAULT_INT;
						this_param.default_int64 = ATOI64(buf);
						break;
					case PURE_FLOAT:
						this_param.default_type = PARAM_DEFAULT_FLOAT;
						this_param.default_double = _tstof(buf); // _tstof() vs. ATOF() because PURE_FLOAT is never hexadecimal.
						break;
					default: // Not numeric (and also not a quoted string because that was handled earlier).
						sntprintf(buf, _countof(buf), _T("%s ?? %s := %.*s"), this_param.var->mName, this_param.var->mName, value_length, param_start);
						if (!at_least_one_default_expr)
						{
							at_least_one_default_expr = true;
							if (!AddLine(ACT_BLOCK_BEGIN))
								return FAIL;
						}
						if (!ParseAndAddLine(buf, ACT_EXPRESSION))
							return FAIL;
						this_param.default_type = PARAM_DEFAULT_UNSET;
					}
				}
				param_end = param_start + value_length;
			}
			param_must_have_default = true;  // For now, all other params after this one must also have default values.
			// Set up for the next iteration:
			param_start = omit_leading_whitespace(param_end);
		}
		else if (*param_start == '?')
		{
			this_param.default_type = PARAM_DEFAULT_UNSET;
			param_must_have_default = true;  // For now, all other params after this one must also have default values.
			// Set up for the next iteration:
			param_start = omit_leading_whitespace(param_start + 1);
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
				return ScriptError(ERR_MISSING_PARAM_NAME, aBuf); // Reporting aBuf vs. param_start seems more informative since Vicinity isn't shown.
		}
		//else it's ')', in which case the next iteration will handle it.
		// Above has ensured that param_start now points to the next parameter, or ')' if none.
	} // for() each formal parameter.

	if (at_least_one_default_expr)
		mIgnoreNextBlockBegin = true; // This is only set after all parameters are parsed, in case they contain fat arrow functions.

	if (param_count)
	{
		// Allocate memory only for the actual number of parameters actually present.
		func.mParam = SimpleHeap::Alloc<FuncParam>(param_count);
		func.mParamCount = param_count - func.mIsVariadic; // i.e. don't count the final "param*" of a variadic function.
		memcpy(func.mParam, param, param_count * sizeof(FuncParam));
	}
	//else leave func.mParam/mParamCount set to their NULL/0 defaults.
	
	if (last_hotfunc)
	{
		// Check all hotkeys, since there might be multiple
		// hotkeys stacked:
		// ^a::
		// ^b::
		//    somefunc(){ ...
		for (int i = 0; i < Hotkey::sHotkeyCount; ++i)
		{
			for (HotkeyVariant* v = Hotkey::shk[i]->mFirstVariant; v; v = v->mNextVariant)
				if (v->mCallback == last_hotfunc)
				{
					v->mCallback = &func;
					v->mOriginalCallback = &func;	// To make scripts more maintainable and to
													// make the hotkey() function more consistent.
													// For example,
													// x::{
													// }
													// Hotkey 'x', 'x', 'off'
													// should be valid if the script is changed to,
													// x::
													// myFunc(){
													// }
													// Hotkey 'x', 'x', 'off'

					break;	// Only one variant possible, per hotkey.
				}
		}
		// Check hotstrings as well (even if a hotkey was found):
		for (int i = Hotstring::sHotstringCount - 1; i >= 0; --i) // Start with the last one defined, for performance.
		{
			if (Hotstring::shs[i]->mCallback != last_hotfunc)
				// This hotstring has a function or is auto-replace.
				// Since hotstrings are listed in order of definition and we're iterating in
				// the reverse order, there's no need to continue.
				break;
			Hotstring::shs[i]->mCallback = &func;
		}

		if (func.mMinParams > 1 || (func.mParamCount == 0 && !func.mIsVariadic))
			// func must accept at least one parameter (or use *), any remaining parameters must be optional.
			return ScriptError(func.mParamCount == 0 ? ERR_PARAM_REQUIRED : ERR_HOTKEY_FUNC_PARAMS, aBuf);
		
		mUnusedHotFunc = last_hotfunc;	// This function not used now, so reuse it for the 
										// next hotkey which needs a function.
		mHotFuncs.mCount--;				// Consider it gone from the list.
	}

	// At this point, param_start points to ')'.
	ASSERT(*param_start == ')');
	param_start = omit_leading_whitespace(param_start + 1);
	if (*param_start == '{') // OTB.
	{
		if (!AddLine(ACT_BLOCK_BEGIN))
			return FAIL;
		ASSERT(!param_start[1]); // IsFunctionDefinition() verified this.
	}
	else if (*param_start == '=' && '>' == param_start[1]) // => expr
	{
		param_start = omit_leading_whitespace(param_start + 2);
		if (!*param_start)
			return ScriptError(ERR_INVALID_FUNCDECL, aBuf);
		if (!AddLine(ACT_BLOCK_BEGIN)
			|| !ParseAndAddLine(param_start, ACT_RETURN)
			|| !AddLine(ACT_BLOCK_END))
			return FAIL;
	}
	else
	{
		// IsFunctionDefinition() permits only {, => or \0.
		ASSERT(!*param_start);
	}
	return OK;
}



ResultType Script::DefineClass(LPTSTR aBuf)
{
	if (mClassObjectCount == MAX_NESTED_CLASSES)
		return ScriptError(_T("This class definition is nested too deep."), aBuf);

	LPTSTR cp, class_name = aBuf, base_class_name = nullptr;
	Object *outer_class, *base_class = Object::sClass, *base_prototype = Object::sPrototype;
	Var *class_var;
	ExprTokenType token;

	for (cp = aBuf; *cp && !IS_SPACE_OR_TAB(*cp); ++cp);
	size_t class_name_length = cp - aBuf;
	if (*cp)
	{
		*cp = '\0'; // Null-terminate here for class_name.
		cp = omit_leading_whitespace(cp + 1);
		if (_tcsnicmp(cp, _T("extends"), 7) || !IS_SPACE_OR_TAB(cp[7]))
			return ScriptError(_T("Syntax error in class definition."), cp);
		base_class_name = omit_leading_whitespace(cp + 8);
		if (!*base_class_name)
			return ScriptError(_T("Missing class name."), cp);
		base_class = FindClass(base_class_name);
		if (!base_class)
		{
			// This class hasn't been defined yet, but it might be.  Automatically create the
			// class, but store it in the "unresolved" list.  When its definition is encountered,
			// it will be removed from the list.  If any classes remain in the list when the end
			// of the script is reached, an error will be thrown.
			if (mUnresolvedClasses)
				base_class = (Object *)mUnresolvedClasses->GetOwnPropObj(base_class_name);
			else
				mUnresolvedClasses = Object::Create();
			if (!base_class)
			{
				if (   !(base_prototype = Object::CreatePrototype(base_class_name))
					|| !(base_class = Object::CreateClass(base_prototype))
					// This property will be removed when the class definition is encountered:
					|| !base_class->SetOwnProp(_T("Line"), ((__int64)mCurrFileIndex << 32) | mCombinedLineNumber)
					|| !mUnresolvedClasses->SetOwnProp(base_class_name, base_class)  )
					return ScriptError(ERR_OUTOFMEM);
			}
		}
		base_prototype = (Object *)base_class->GetOwnPropObj(_T("Prototype"));
	}

	// Validate the name even if this is a nested definition, for consistency.
	if (!Var::ValidateName(class_name, DISPLAY_CLASS_ERROR))
		return FAIL;

	Object *&class_object = mClassObject[mClassObjectCount]; // For convenience.
	class_object = nullptr;
	bool conflict_found = false;
	
	if (mClassObjectCount) // Nested class definition.
	{
		outer_class = mClassObject[mClassObjectCount - 1];
		if (outer_class->GetOwnProp(token, class_name))
		{
			conflict_found = true;
			// If "continuing" a predefined class was permitted, this is where we would
			// set class_object; but only after confirming that token contains a class.
			// It could be a previously declared instance/static var.
		}
	}
	else // Top-level class definition.
	{
		*mClassName = '\0'; // Init.
		VarList *varlist = GlobalVars();
		int insert_pos;
		class_var = varlist->Find(class_name, &insert_pos);
		if (class_var)
		{
			if (class_var->IsDeclared())
				return ConflictingDeclarationError(_T("class"), class_var);
		}
		else if (  !(class_var = AddVar(class_name, class_name_length, varlist, insert_pos, VAR_DECLARE_GLOBAL))  )
			return FAIL;
	}
	
	size_t length = _tcslen(mClassName), extra_length = class_name_length + 1; // +1 for '.'
	if (length + extra_length >= _countof(mClassName))
		return ScriptError(_T("Full class name is too long."));
	if (length)
		mClassName[length++] = '.';
	tmemcpy(mClassName + length, class_name, extra_length); // Includes null-terminator.
	length += extra_length - 1; // -1 to exclude the null-terminator.

	// For now, it seems more useful to detect a duplicate as an error rather than as
	// a continuation of the previous definition.  Partial definitions might be allowed
	// in future, perhaps via something like "Class Foo continued".
	if (conflict_found)
		return ScriptError(ERR_DUPLICATE_DECLARATION, aBuf);

	if (mUnresolvedClasses)
	{
		token.SetValue(mClassName, length);
		ExprTokenType *param = &token;
		ResultToken result_token;
		result_token.symbol = SYM_STRING; // Init for detecting SYM_OBJECT below.
		// Search for class and simultaneously remove it from the unresolved list:
		mUnresolvedClasses->DeleteProp(result_token, 0, IT_CALL, &param, 1); // result_token := mUnresolvedClasses.Delete(token)
		// If a field was found/removed, it can only be SYM_OBJECT.
		if (result_token.symbol == SYM_OBJECT)
		{
			// Use this object as the class.  At least one other object already refers to it as mBase.
			class_object = (Object *)result_token.object;
			class_object->DeleteOwnProp(_T("Line")); // Remove the error reporting info.
		}
	}

	Object *prototype;
	if (class_object)
		prototype = (Object *)class_object->GetOwnPropObj(_T("Prototype"));
	else
		class_object = Object::CreateClass(prototype = Object::CreatePrototype(mClassName));

	if (mClassObjectCount)
	{
		// Define property outer_class.%class_name%.
		outer_class->DefineClass(class_name, class_object);
	}
	else
	{
		class_var->Assign(class_object); // Assign to global variable named %class_name%.
		class_var->MakeReadOnly();
		class_var->MarkUninitialized(); // This enables the constructor to be called on first use.
	}

	prototype->SetBase(base_prototype);
	class_object->SetBase(base_class);

	if (mClassObjectCount ? !DefineClassVarInit(mClassName, true, outer_class, ACT_EXPRESSION) : !ParseAndAddLine(mClassName, ACT_EXPRESSION))
		return FAIL;
	++mClassObjectCount;
	// Define __Init unconditionally so that classes without static vars do not inherit __Init.
	if (!DefineClassInit(true))
		return FAIL;
	if (base_class_name)
	{
		// Ensure the base class is initialized by inserting a reference to it into static __Init.
		// Otherwise, the base class could be accessed via .base without triggering initialization.
		if (!DefineClassVarInit(base_class_name, true, class_object, ACT_EXPRESSION))
			return FAIL;
	}
	
	// This line enables a class without any static methods to be freed at program exit,
	// or sooner if it's a nested class and the script removes it from the outer class.
	// Classes with static methods are never freed, since the method itself retains a
	// reference to the class.
	class_object->Release();

	return OK;
}


ResultType Script::DefineClassProperty(LPTSTR aBuf, bool aStatic, bool &aBufHasBraceOrNotNeeded)
{
	LPTSTR name_end = find_identifier_end(aBuf);
	LPTSTR param_start = omit_leading_whitespace(name_end);
	LPTSTR param_end, next_token;
	if (*param_start == '[')
	{
		param_start = omit_leading_whitespace(param_start + 1);
		if (*param_start == ']')
			return ScriptError(ERR_PROPERTY_EMPTY_BRACKETS, aBuf);
		param_end = param_start + FindExprDelim(param_start, ']');
		if (!*param_end)
			return ScriptError(ERR_MISSING_CLOSE_BRACKET, aBuf);
		next_token = omit_leading_whitespace(param_end + 1);
	}
	else
	{
		param_end = param_start;
		next_token = param_start;
	}
	if (!*next_token)
		aBufHasBraceOrNotNeeded = false;
	else if (*next_token == '{' && !next_token[1]
		|| *next_token == '=' && next_token[1] == '>')
		aBufHasBraceOrNotNeeded = true;
	else
		return ScriptError(ERR_INVALID_LINE_IN_CLASS_DEF, aBuf);

	// Save the property name and parameter list for later use with DefineFunc().
	int name_length = int(name_end - aBuf);
	int param_length = int(param_end - param_start);
	mClassPropertyDef = tmalloc(name_length + param_length + 7); // +7 for ".Get()\0"
	if (!mClassPropertyDef)
		return ScriptError(ERR_OUTOFMEM);
	_stprintf(mClassPropertyDef, _T("%.*s.Get(%.*s)"), name_length, aBuf, param_length, param_start);

	Object *class_object = mClassObject[mClassObjectCount - 1];
	if (!aStatic)
		class_object = (Object *)class_object->GetOwnPropObj(_T("Prototype"));
	TCHAR end_char = *name_end; // In case there's no space before =>.
	*name_end = 0; // Terminate for aBuf use below.
	switch (class_object->GetOwnPropType(aBuf))
	{
	case Object::PropType::Object:
	case Object::PropType::DynamicValue: // get/set
	case Object::PropType::DynamicMixed: // get/set and call
		return ScriptError(ERR_DUPLICATE_DECLARATION, aBuf);
	}
	mClassProperty = class_object->DefineProperty(aBuf);
	mClassPropertyStatic = aStatic;
	if (!mClassProperty)
		return ScriptError(ERR_OUTOFMEM);

	*name_end = end_char;
	if (*next_token == '=') // => expr
	{
		// mClassPropertyDef is already set up for "Get".
		if (!DefineClassPropertyXet(aBuf, next_token))
			return FAIL;
		// Immediately close this property definition.
		mClassProperty = NULL;
		free(mClassPropertyDef);
		mClassPropertyDef = NULL;
	}

	return OK;
}


ResultType Script::DefineClassPropertyXet(LPTSTR aBuf, LPTSTR aEnd)
{
	// For simplicity, pass the property definition to DefineFunc instead of the actual
	// line text, even though it makes some error messages a bit inaccurate. (That would
	// happen anyway when DefineFunc() finds a syntax error in the parameter list.)
	if (!DefineFunc(mClassPropertyDef, mClassPropertyStatic))
		return FAIL;
	if (mClassProperty->MinParams == -1)
	{
		int hidden_params = ctoupper(*aBuf) == 'S' ? 2 : 1;
		mClassProperty->MinParams = g->CurrentFunc->mMinParams - hidden_params;
		mClassProperty->MaxParams = g->CurrentFunc->mIsVariadic ? INT_MAX : g->CurrentFunc->mParamCount - hidden_params;
	}
	if (*aEnd && !AddLine(ACT_BLOCK_BEGIN)) // *aEnd is '{' or '='.
		return FAIL;
	if (*aEnd == '=') // => expr
	{
		aEnd = omit_leading_whitespace(aEnd + 2);
		if (!ParseAndAddLine(aEnd, ACT_RETURN)
			|| !AddLine(ACT_BLOCK_END))
			return FAIL;
	}
	return OK;
}


ResultType Script::DefineClassVars(LPTSTR aBuf, bool aStatic)
{
	Object *class_object = mClassObject[mClassObjectCount - 1];
	if (!aStatic)
		class_object = (Object *)class_object->GetOwnPropObj(_T("Prototype"));

	LPTSTR item, item_end;
	TCHAR orig_char, buf[LINE_SIZE];
	size_t buf_used = 0;
	ExprTokenType empty_token;
	empty_token.symbol = SYM_MISSING;

	for (item = omit_leading_whitespace(aBuf); *item;) // FOR EACH COMMA-SEPARATED ITEM IN THE DECLARATION LIST.
	{
		item_end = find_identifier_end(item);
		if (item_end == item)
			return ScriptError(ERR_INVALID_CLASS_VAR, item);
		orig_char = *item_end;
		*item_end = '\0'; // Temporarily terminate.
		ExprTokenType existing;
		auto item_exists = class_object->GetOwnPropType(item);
		bool item_name_has_dot = (orig_char == '.');
		if (item_name_has_dot)
		{
			*item_end = orig_char; // Undo termination.
			// This is something like "object.key := 5", which is only valid if "object" was
			// previously declared (and will presumably be assigned an object at runtime).
			// Ensure that at least the root class var exists; any further validation would
			// be impossible since the object doesn't exist yet.
			if (item_exists == Object::PropType::None)
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
			switch (item_exists)
			{
			case Object::PropType::Value:
			case Object::PropType::Object: // Prototype or nested class.
				return ScriptError(ERR_DUPLICATE_DECLARATION, item);
			case Object::PropType::None:
				// Assign class_object[item] := "" to mark it as a value property
				// and allow duplicate declarations to be detected:
				if (!class_object->SetOwnProp(item, empty_token))
					return ScriptError(ERR_OUTOFMEM);
			// But for PropType::Dynamic, we want this line to assign to the property, so don't overwrite it.
			}
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
		// Build an expression which can be executed to initialize this class variable.
		item_end = omit_leading_whitespace(item_end);
		LPTSTR right_side_of_operator = item_end; // Save for use below.

		item_end += FindExprDelim(item_end); // Find the next comma which is not part of the initializer (or find end of string).

		// Append "ClassNameOrThis.VarName := Initializer, " to the buffer.
		LPCTSTR initializer = _T("%s.%.*s := %.*s, ");
		int chars_written = _sntprintf(buf + buf_used, _countof(buf) - buf_used, initializer
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
		return DefineClassVarInit(buf, aStatic, class_object);
	}
	return OK;
}


ResultType Script::DefineClassVarInit(LPTSTR aBuf, bool aStatic, Object *aObject, ActionTypeType aActionType)
{
	Line *script_last_line = mLastLine;
	Line *block_end;
	auto init_func = (UserFunc *)aObject->GetOwnPropMethod(_T("__Init")); // Can only be a user-defined function or nullptr at this stage.
	if (init_func)
	{
		// __Init method already exists, so find the end of its body.
		for (block_end = init_func->mJumpToLine;
			block_end->mActionType != ACT_BLOCK_END || block_end->mAttribute != init_func;
			block_end = block_end->mNextLine);
	}
	else
	{
		// Create an __Init method for this class/prototype.
		init_func = DefineClassInit(aStatic);
		if (!init_func)
			return FAIL;
		script_last_line = block_end = mLastLine;
	}
	
	// Temporarily replace mLastLine to insert script lines at the end of __Init.
	g->CurrentFunc = init_func; // g->CurrentFunc should be NULL prior to this.
	mLastLine = block_end->mPrevLine; // i.e. insert before block_end.
	mLastLine->mNextLine = nullptr; // For maintainability; AddLine() should overwrite it regardless.
	mCurrLine = nullptr; // Fix for v1.1.09.02: Leaving this non-NULL at best causes error messages to show irrelevant vicinity lines, and at worst causes a crash because the linked list is in an inconsistent state.

	Line *prl = mPendingRelatedLine; // This was set by AddLine(ACT_BLOCK_END) if DefineClassInit() was just called.
	mPendingRelatedLine = nullptr;
	mNoUpdateLabels = true;
	if (!ParseAndAddLine(aBuf, aActionType))
		return FAIL; // Above already displayed the error.
	mNoUpdateLabels = false;
	mPendingRelatedLine = prl;

	if (init_func->mJumpToLine == block_end) // This can be true only for the first static initializer.
		init_func->mJumpToLine = mLastLine;
	// Rejoin the function's block-end (and any lines following it) to the main script.
	mLastLine->mNextLine = block_end;
	block_end->mPrevLine = mLastLine;
	// mFirstLine should be left as it is: if it was NULL, it now contains a pointer to our
	// __init function's block-begin, which is now the very first executable line in the script.
	g->CurrentFunc = nullptr;
	// Restore mLastLine so that any subsequent script lines are added at the correct point.
	mLastLine = script_last_line;
	return OK;
}


UserFunc *Script::DefineClassInit(bool aStatic)
{
	TCHAR def[] = _T("__Init(){");
	if (!DefineFunc(def, aStatic))
		return nullptr;
	if (!aStatic)
	{
		if (!ParseAndAddLine(SUPER_KEYWORD _T(".__Init()"), ACT_EXPRESSION)) // Initialize base-class variables first. Relies on short-circuit evaluation.
			return nullptr;
		mLastLine->mLineNumber = 0; // Signal the debugger to skip this line while stepping in/over/out.
	}
	auto init_func = g->CurrentFunc;
	if (!AddLine(ACT_BLOCK_END)) // This also resets g->CurrentFunc to NULL.
		return nullptr;
	mLastLine->mLineNumber = 0; // See above.
	return init_func;
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
	// To reserve for possible future use with local classes, resolve the class variable according
	// to the current scope (which is always global for class..extends, but often local for Catch).
	// This also avoids confusion due to inconsistency between `catch x` and other references to x.
	Var *base_var = FindVar(class_name, cp - class_name);
	if (!base_var)
		return NULL;

	// Although at load time only the "Object" type can exist, dynamic_cast is used in case we're called at run-time:
	if (  !(base_var->IsObject() && (base_object = dynamic_cast<Object *>(base_var->Object()))
		&& base_object->IsDerivedFrom(Object::sClassPrototype))  ) // Rule out something silly like "class x extends MsgBox".
		return NULL;

	// Even if the loop below has no iterations, it initializes 'key' to the appropriate value:
	for (key = cp + 1; cp = _tcschr(key, '.'); key = cp + 1) // For each key in something like TypeVar.Key1.Key2.
	{
		if (cp == key)
			return NULL; // ScriptError(_T("Missing name."), cp);
		*cp = '\0'; // Terminate at the delimiting dot.
		// Invoking at this stage isn't safe since aClassName could include the name of a
		// user-defined property getter, or the class might implement __Get, and the script
		// isn't necessarily ready to execute.  Get it this way instead:
		auto getter = dynamic_cast<BuiltInFunc*>(base_object->GetOwnPropGetter(key));
		if (!getter || getter->mBIF != Class_GetNestedClass)
			return NULL;
		base_object = ((NestedClassInfo*)getter->mData)->class_object;
	}

	return base_object;
}


Object *Object::GetUnresolvedClass(LPTSTR &aName)
// This method is only valid for mUnresolvedClass.
{
	if (!mFields.Length())
		return NULL;
	aName = mFields[0].name;
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
	if (base->GetOwnProp(token, _T("Line")))
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

#define FUNC_LIB_EXT EXT_AUTOHOTKEY
#define FUNC_LIB_EXT_LENGTH (_countof(FUNC_LIB_EXT) - 1)
#define FUNC_LOCAL_LIB _T("\\Lib\\") // Needs leading and trailing backslash.
#define FUNC_LOCAL_LIB_LENGTH (_countof(FUNC_LOCAL_LIB) - 1)
#define FUNC_USER_LIB _T("\\AutoHotkey\\Lib\\") // Needs leading and trailing backslash.
#define FUNC_USER_LIB_LENGTH (_countof(FUNC_USER_LIB) - 1)
#define FUNC_STD_LIB _T("\\Lib\\") // Needs leading and trailing backslash.
#define FUNC_STD_LIB_LENGTH (_countof(FUNC_STD_LIB) - 1)

#define FUNC_LIB_COUNT 3

void Script::InitFuncLibraries(FuncLibrary aLib[])
{
	// Local lib in script's directory.
	InitFuncLibrary(aLib[0], mFileDir, FUNC_LOCAL_LIB);

	// User lib in Documents folder.
	PWSTR docs_dir = GetDocumentsFolder();
	InitFuncLibrary(aLib[1], docs_dir, FUNC_USER_LIB);
	CoTaskMemFree(docs_dir);

	// Std lib in AutoHotkey directory.
	InitFuncLibrary(aLib[2], mOurEXEDir, FUNC_STD_LIB);
}

void Script::InitFuncLibrary(FuncLibrary &aLib, LPTSTR aPathBase, LPTSTR aPathSuffix)
{
	TCHAR buf[T_MAX_PATH + 1]; // +1 to ensure truncated paths are filtered out due to being too long for GetFileAttributes().  It would probably have to be intentional due to MAX_PATH limits elsewhere, but this +1 costs nothing.
	int length = sntprintf(buf, _countof(buf), _T("%s%s"), aPathBase, aPathSuffix);
	DWORD attr = GetFileAttributes(buf); // Seems to accept directories that have a trailing backslash, which is good because it simplifies the code.
	if (attr == 0xFFFFFFFF || !(attr & FILE_ATTRIBUTE_DIRECTORY)) // Directory doesn't exist or it's a file vs. directory. Relies on short-circuit boolean order.
	{
		aLib.path = _T("");
		return;
	}
	// Allow room for appending each candidate file/function name.  This could be exactly
	// MAX_PATH on ANSI builds since no path longer than that would work, but doing it this
	// way simplifies length checks later (and Unicode builds support much longer paths).
	size_t buf_size = length + MAX_VAR_NAME_LENGTH + FUNC_LIB_EXT_LENGTH + 1;
	aLib.path = SimpleHeap::Alloc<TCHAR>(buf_size);
	tmemcpy(aLib.path, buf, length + 1);
	aLib.length = length;
}

void Script::IncludeLibrary(LPTSTR aFuncName, size_t aFuncNameLength, bool &aErrorWasShown, bool &aFileWasFound)
// Caller must ensure that aFuncName doesn't already exist as a defined function.
// If aFuncNameLength is 0, the entire length of aFuncName is used.
{
	aErrorWasShown = false; // Set default for this output parameter.
	aFileWasFound = false;

	int i;
	DWORD attr;

	static FuncLibrary sLib[FUNC_LIB_COUNT] = {0};

	if (!sLib[0].path) // Allocate & discover paths only upon first use because many scripts won't use anything from the library. This saves a bit of memory and performance.
		InitFuncLibraries(sLib);
	// Above must ensure that all sLib[].path elements are non-NULL (but they can be "" to indicate "no library").

	if (!aFuncNameLength) // Caller didn't specify, so use the entire string.
		aFuncNameLength = _tcslen(aFuncName);
	if (aFuncNameLength > MAX_VAR_NAME_LENGTH) // Too long to fit in the allowed space, and also too long to be a valid function name.
		return;

	TCHAR *dest, *first_underscore, class_name_buf[MAX_VAR_NAME_LENGTH + 1];
	LPTSTR naked_filename = aFuncName;               // Set up for the first iteration.
	size_t naked_filename_length = aFuncNameLength; //

	for (int second_iteration = 0; second_iteration < 2; ++second_iteration)
	{
		for (i = 0; i < FUNC_LIB_COUNT; ++i)
		{
			if (!*sLib[i].path) // Library is marked disabled, so skip it.
				continue;

			dest = (LPTSTR) tmemcpy(sLib[i].path + sLib[i].length, naked_filename, naked_filename_length); // Append the filename to the library path.
			_tcscpy(dest + naked_filename_length, FUNC_LIB_EXT); // Append the file extension.

			attr = GetFileAttributes(sLib[i].path); // Testing confirms that GetFileAttributes() doesn't support wildcards; which is good because we want filenames containing question marks to be "not found" rather than being treated as a match-pattern.
			if (attr == 0xFFFFFFFF || (attr & FILE_ATTRIBUTE_DIRECTORY)) // File doesn't exist or it's a directory.
				continue;
			// Since above didn't "continue", a file exists whose name matches that of the requested function.
			aFileWasFound = true; // Indicate success for #include <lib>, which doesn't necessarily expect a function to be found.

			// g->CurrentFunc is non-NULL when the function-call being resolved is inside
			// a function.  Save and reset it for correct behaviour in the include file:
			auto current_func = g->CurrentFunc;
			g->CurrentFunc = NULL;

			// Fix for v1.1.06.00: If the file contains any lib #includes, it must be loaded AFTER the
			// above writes sLib[i].path to the iLib file, otherwise the wrong filename could be written.
			if (!LoadIncludedFile(sLib[i].path, false, false)) // Fix for v1.0.47.05: Pass false for allow-dupe because otherwise, it's possible for a stdlib file to attempt to include itself (especially via the LibNamePrefix_ method) and thus give a misleading "duplicate function" vs. "func does not exist" error message.  Obsolete: For performance, pass true for allow-dupe so that it doesn't have to check for a duplicate file (seems too rare to worry about duplicates since by definition, the function doesn't yet exist so it's file shouldn't yet be included).
				aErrorWasShown = true; // Above has just displayed its error (e.g. syntax error in a line, failed to open the include file, etc).  So override the default set earlier.
			
			g->CurrentFunc = current_func; // Restore.
			return; // A file was found, so look no further.
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
}

#endif



template<typename T, int S>
T *ScriptItemList<T, S>::Find(LPCTSTR aName, int *apInsertPos)
{
	// Using a binary searchable array vs a linked list speeds up dynamic function calls, on average.
	int left, right, mid, result;
	for (left = 0, right = mCount - 1; left <= right;)
	{
		mid = (left + right) / 2;
		result = _tcsicmp(aName, mItem[mid]->mName); // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales; 3) performance.
		if (result > 0)
			left = mid + 1;
		else if (result < 0)
			right = mid - 1;
		else // Match found.
		{
			if (apInsertPos)
				*apInsertPos = mid;
			return mItem[mid];
		}
	}
	if (apInsertPos)
		*apInsertPos = left;
	return NULL;
}



template<typename T, int S>
T *ScriptItemList<T, S>::Find(LPCTSTR aName, size_t aNameLength, int *apInsertPos)
{
	if (!aName[aNameLength])
		return Find(aName, apInsertPos);

	// For the below, no error is reported because callers don't want that.  Instead, simply return
	// nullptr to indicate that names that are illegal or too long are not found.  When the caller
	// later tries to add the item, it will get an error then:
	if (aNameLength > MAX_VAR_NAME_LENGTH)
		return nullptr;

	// The following copy is made because it allows the search to use _tcsicmp() instead of _tcsnicmp(),
	// which improves performance.  The copy includes only the first aNameLength characters from aName:
	TCHAR name[MAX_VAR_NAME_LENGTH + 1];
	tcslcpy(name, aName, aNameLength + 1);  // +1 to convert length to size.

	return Find(name, apInsertPos);
}



Func *Script::FindGlobalFunc(LPCTSTR aFuncName, size_t aFuncNameLength)
{
	if (Var *var = FindVar(aFuncName, aFuncNameLength, VAR_GLOBAL))
		if (var->Type() == VAR_CONSTANT)
			return dynamic_cast<Func *>(var->ToObject());
	return nullptr;
}



BuiltInFunc::BuiltInFunc(FuncEntry &bif) : BuiltInFunc(bif.mName)
{
	mBIF = bif.mBIF;
	mMinParams = bif.mMinParams;
	mIsVariadic = bif.mMaxParams == MAX_FUNCTION_PARAMS;
	mParamCount = mIsVariadic ? bif.mMinParams : bif.mMaxParams;
	mFID = (BuiltInFunctionID)bif.mID;
	mOutputVars = bif.mOutputVars;
}



Func *Script::GetBuiltInFunc(LPTSTR aFuncName)
{
	int left, right, mid, result;
	for (left = 0, right = _countof(g_BIF) - 1; left <= right;)
	{
		mid = (left + right) / 2;
		result = _tcsicmp(aFuncName, g_BIF[mid].mName);
		if (result > 0)
			left = mid + 1;
		else if (result < 0)
			right = mid - 1;
		else // Match found.
			return new BuiltInFunc(g_BIF[mid]);
	}
	return GetBuiltInMdFunc(aFuncName);
}



UserFunc *Script::AddFunc(LPCTSTR aFuncName, size_t aFuncNameLength, Object *aClassObject)
// Returns the address of the new function or NULL on failure.
// The caller must already have verified that this isn't a duplicate function.
{
	bool is_static = !_tcsnicmp(aFuncName, _T("Static"), 6) && IS_SPACE_OR_TAB(aFuncName[6]);
	if (is_static)
	{
		if (aClassObject || !g->CurrentFunc)
		{
			ScriptError(ERR_INVALID_FUNCDECL, aFuncName); // Uses a generic message to minimize code size.
			return nullptr;
		}
		size_t n;
		for (n = 7; IS_SPACE_OR_TAB(aFuncName[n]); ++n);
		if (aFuncNameLength > n) // Checked for maintainability; should always be true.
		{
			aFuncNameLength -= n;
			aFuncName += n;
		}
		//else (shouldn't happen): it will fail to validate below.
	}

	if (aFuncNameLength == -1) // Caller didn't specify, so use the entire string.
		aFuncNameLength = _tcslen(aFuncName);

	if (aFuncNameLength > MAX_VAR_NAME_LENGTH) // Various parts of the code rely on this limit being enforced.
	{
		ScriptError(_T("Function name too long."), aFuncName);
		return nullptr;
	}

	// ValidateName requires that the name be null-terminated, but it isn't in this case.
	// Doing this first saves doing tcslcpy() into a temporary buffer, and won't leak memory
	// since the script currently always exits if an error occurs anywhere below:
	LPTSTR new_name = SimpleHeap::Alloc(aFuncName, aFuncNameLength);

	if (!aClassObject && *new_name && !Var::ValidateName(new_name, DISPLAY_FUNC_ERROR))  // Variable and function names are both validated the same way.
		return nullptr; // Above already displayed the error for us.

	auto the_new_func = new UserFunc(new_name);

	if (aClassObject)
	{
		LPTSTR key = _tcsrchr(new_name, '.'); // DefineFunc() always passes "ClassName.MethodName".
		++key;
		if (!Var::ValidateName(key, DISPLAY_METHOD_ERROR))
			return nullptr;
		if (mClassProperty)
		{
			bool is_getter = ctoupper(*key) == 'G';
			if (is_getter ? mClassProperty->Getter() : mClassProperty->Setter())
			{
				ScriptError(ERR_DUPLICATE_DECLARATION, new_name);
				return nullptr;
			}
			if (is_getter)
				mClassProperty->SetGetter(the_new_func);
			else
				mClassProperty->SetSetter(the_new_func);
		}
		else
		{
			switch (aClassObject->GetOwnPropType(key))
			{
			case Object::PropType::Object:
			case Object::PropType::DynamicMethod: // call
			case Object::PropType::DynamicMixed: // get/set and call
				ScriptError(ERR_DUPLICATE_DECLARATION, new_name);
				return nullptr;
			}
			if (!aClassObject->DefineMethod(key, the_new_func))
			{
				ScriptError(ERR_OUTOFMEM);
				return nullptr;
			}
		}
		aClassObject->AddRef(); // In case the script clears the class var.
		the_new_func->mClass = aClassObject;
		// Also add it to the script's list of functions, to support #Warn LocalSameAsGlobal
		// and automatic cleanup of objects in static vars on program exit.
	}
	else
	{
		ResultType result = OK;
		VarList *varlist;
		int insert_pos;
		int scope = (g->CurrentFunc ? VAR_DECLARE_LOCAL : VAR_DECLARE_GLOBAL);
		if (is_static)
		{
			the_new_func->mIsStatic = true;
			scope |= VAR_LOCAL_STATIC;
		}
		Var *var = aFuncNameLength ? FindVar(aFuncName, aFuncNameLength, FINDVAR_NO_BIF | scope, &varlist, &insert_pos, &result) : NULL;
		if (var)
		{
			// Anything found at this early stage must have been created by a prior declaration
			// or predefined.  Treat it as an error even if var is uninitialized, otherwise:
			//  1) "static x" or "global x" prior to the function would affect its scope/duration
			//     in a way that is problematic for closures and probably unintentional.
			//  2) It would be possible to declare the function as local/global prior to defining
			//     it, but not afterward (as the conflict would be detected by ParseAndAddLine).
			ConflictingDeclarationError(_T("function"), var);
			return nullptr;
		}
		if (!aFuncNameLength)
		{
			insert_pos = 0;
			varlist = g->CurrentFunc ? &g->CurrentFunc->mVars : GlobalVars();
		}
		var = AddVar(aFuncName, aFuncNameLength, varlist, insert_pos, scope);
		if (!var)
			return nullptr;
		var->Assign(the_new_func);
		var->MakeReadOnly();
	}

	the_new_func->mOuterFunc = g->CurrentFunc;
	
	// The general rule is that a nested function's namespace includes all of the names that
	// were in the outer function's namespace by default.  In other words, if the outer function
	// is assume-global, the nested function should also default to assume-global.
	if (the_new_func->mOuterFunc && the_new_func->mOuterFunc->IsAssumeGlobal())
		the_new_func->mDefaultVarType = VAR_GLOBAL; // Not VAR_DECLARE_GLOBAL, which would prevent referencing of outer vars.

	if (!mFuncs.Insert(the_new_func, mFuncs.mCount))
	{
		ScriptError(ERR_OUTOFMEM);
		return nullptr;
	}

	return the_new_func;
}



template<typename T, int INITIAL_SIZE>
ResultType ScriptItemList<T, INITIAL_SIZE>::Insert(T *aFunc, int aInsertPos)
{
	if (mCount == mCountMax)
	{
		// Allocate or expand function list.
		if (!Alloc(mCountMax ? mCountMax * 2 : INITIAL_SIZE))
			return FAIL;
	}

	if (aInsertPos != mCount) // Need to make room at the indicated position for this variable.
		memmove(mItem + aInsertPos + 1, mItem + aInsertPos, (mCount - aInsertPos) * sizeof(Func *));
	//else both are zero or the item is being inserted at the end of the list, so it's easy.
	mItem[aInsertPos] = aFunc;
	++mCount;
	return OK;
}



template<typename T, int S>
ResultType ScriptItemList<T,S>::Alloc(int aAllocCount)
{
	T **temp = (T **)realloc(mItem, aAllocCount * sizeof(T *)); // If passed NULL, realloc() will do a malloc().
	if (!temp)
		return FAIL;
	mItem = temp;
	mCountMax = aAllocCount;
	return OK;
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
	if (!mArg[aArgIndex].is_expression && mArg[aArgIndex].postfix)
	{
		// This is a simple constant value or variable reference.
		switch (mArg[aArgIndex].postfix->symbol)
		{
		case SYM_STRING:
			return mArg[aArgIndex].postfix->marker_length;
		case SYM_VAR:
			return mArg[aArgIndex].postfix->var->Length();
		}
	}
	// Otherwise, length isn't known, so do it the slow way.
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
	if (!mArg[aArgIndex].is_expression && mArg[aArgIndex].postfix)
		return TokenToInt64(*mArg[aArgIndex].postfix);
	// Otherwise:
	return ATOI64(sArgDeref[aArgIndex]);
}


Var *Script::FindOrAddVar(LPCTSTR aVarName, size_t aVarNameLength, int aScope)
// Caller has ensured that aVarName isn't NULL.
// Returns the Var whose name matches aVarName.  If it doesn't exist, it is created.
{
	if (!*aVarName)
		return NULL;
	if (!aVarNameLength) // Caller didn't specify, so use the entire string.
		aVarNameLength = _tcslen(aVarName);
	ResultType result = OK; // Must set default.
	VarList *varlist;
	int insert_pos;
	Var *var;
	if (var = FindVar(aVarName, aVarNameLength, aScope, &varlist, &insert_pos, &result))
		return var;
	if (!result) // An error was displayed.
		return nullptr;
	// Otherwise, no match found, so create a new var.
	bool is_local = varlist != GlobalVars();
	// This will return NULL if there was a problem, in which case AddVar() will already have displayed the error:
	return AddVar(aVarName, aVarNameLength, varlist, insert_pos
		, (aScope & ~(VAR_LOCAL | VAR_GLOBAL)) | (is_local ? VAR_LOCAL : VAR_GLOBAL)); // When aScope == FINDVAR_DEFAULT, it contains both the "local" and "global" bits.  This ensures only the appropriate bit is set.
}



Var *Script::FindVar(LPCTSTR aVarName, size_t aVarNameLength, int aScope
	, VarList **apList, int *apInsertPos, ResultType *aDisplayError)
// Caller has ensured that aVarName isn't NULL.  It must also ignore the contents of apInsertPos when
// a match (non-NULL value) is returned.
// Returns the Var whose name matches aVarName.  If it doesn't exist, NULL is returned.
// If caller provided a non-NULL apInsertPos, it will be given the array index that a newly
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

	if (search_local)
	{
		auto &func = *g.CurrentFunc;
		int nonstatic_pos, static_pos;
		if (Var *found = func.mVars.Find(var_name, &nonstatic_pos)) return found;
		if (Var *found = func.mStaticVars.Find(var_name, &static_pos)) return found;
		bool add_static = (aScope & VAR_LOCAL_STATIC)
			|| !(aScope & VAR_DECLARED) && (func.mDefaultVarType & VAR_LOCAL_STATIC);
		if (apList) *apList = add_static ? &func.mStaticVars : &func.mVars;
		if (apInsertPos) *apInsertPos = add_static ? static_pos : nonstatic_pos;
	}
	else
	{
		auto varlist = GlobalVars();
		int insert_pos;
		if (Var *found = varlist->Find(var_name, &insert_pos)) return found;
		if (apList) *apList = varlist;
		if (apInsertPos) *apInsertPos = insert_pos;

		if (!(aScope & FINDVAR_NO_BIF))
		// Built-in functions can be shadowed, so are checked only in this section.
		if (auto *func = GetBuiltInFunc(var_name))
		{
			Var *var = AddVar(var_name, aVarNameLength, varlist, insert_pos, VAR_DECLARE_GLOBAL);
			if (!var)
			{
				if (aDisplayError)
					*aDisplayError = FAIL;
				func->Release();
				return nullptr;
			}
			var->AssignSkipAddRef(func);
			var->MakeReadOnly();
			return var;
		}
	}

	// Since no match was found, if this is a local fall back to searching outer functions and the
	// list of globals if the caller didn't insist on a particular type:
	if (search_local && (aScope & VAR_GLOBAL))
	{
		// Search outer functions, if applicable:
		if (g.CurrentFunc->mOuterFunc)
			if (Var *found = FindUpVar(var_name, aVarNameLength, *g.CurrentFunc, aDisplayError))
				return found;
		// No local var was found, nor was a global declared in this function (otherwise it would
		// have been found in g.CurrentFunc->mStaticVars), so caller wants to fall back to a global
		// in this case only if the function is assume-global or FINDVAR_GLOBAL_FALLBACK was used.
		// However, they want the insertion (if our caller will be doing one) to insert according to
		// the current assume-mode.  Therefore, if the mode is assume-global, pass apList/apInsertPos
		// to FindVar() so that it will update them to be global.
		if (g.CurrentFunc->IsAssumeGlobal())
			return FindVar(aVarName, aVarNameLength, FINDVAR_GLOBAL, apList, apInsertPos);
		if (aScope & FINDVAR_GLOBAL_FALLBACK)
			if (Var *found = FindGlobalVar(aVarName, aVarNameLength))
				return found;
	}
	// Otherwise, since above didn't return:
	if (auto *builtin = GetBuiltInVar(var_name))
	{
		Var *var = FindOrAddBuiltInVar(var_name, builtin);
		if (!var && aDisplayError)
			*aDisplayError = FAIL;
		return var;
	}
	return nullptr;
}



Var *Script::FindUpVar(LPCTSTR aVarName, size_t aVarNameLength, UserFunc &aInner, ResultType *aDisplayError)
{
	// Do not search outer if inner is *explicitly* assume-global, since in that case a
	// global var should be returned in preference to anything defined by outer.
	if (aInner.mDefaultVarType == VAR_DECLARE_GLOBAL)
		return nullptr;
	auto &outer = *aInner.mOuterFunc;
	Var *outer_var;
	if (  (outer_var = outer.mStaticVars.Find(aVarName)) || aInner.mIsStatic  )
		return outer_var; // Can be nullptr if aInner.mIsStatic.
	if (  !(outer_var = outer.mVars.Find(aVarName))  )
	{
		if (  !(outer.mOuterFunc && (outer_var = FindUpVar(aVarName, aVarNameLength, outer, aDisplayError)))  )
			return nullptr;
		// Static or declared-global variables of outer would return above by virtue of being
		// in mStaticVars, but variables pulled in from further out must still be checked:
		if (!outer_var->IsNonStaticLocal())
			return outer_var;
	}
	// At this point, all var refs used in declarations, assignments or &var in the outer
	// function should have already been parsed, while it's possible that some read-refs
	// have not.  Ignore all variables that lack an assignment, &var or declaration.
	//  - Avoids an issue where only read-refs ABOVE the nested function are accessible
	//    (which could also be fixed by changing PreparseVarRefs to parse one function
	//    at a time, parsing outer in its entirety before inner).
	//  - Avoids inconsistency between read-refs and write-refs: when the outer var is created
	//    by a read-ref, closures which only read would capture it, while closures which write
	//    would create their own local.
	//  - Avoids having behaviour depend on whether a global exists: when outer and inner both
	//    only use read-refs (but maybe assign dynamically), whether the nested function becomes
	//    a closure might depend on whether the var is global or local to outer.
	//  - Having a non-dynamic assignment (and no global declaration) makes it easier to see
	//    that the var is local by looking at just the outer function, and avoids warnings.
	//  - A clear rule is easier to remember.
	if (  !(outer_var->IsAssignedSomewhere() || outer_var->IsDeclared())  )
		return nullptr;
	// Since this local variable is not static, things need to be set up so that the inner
	// function will capture it as an upvar at the appropriate time.
	if (mIsReadyToExecute)
	{
		// This dynamic reference resolved to a variable that hasn't been captured by the
		// current function, so is not safe to return.
		if (aDisplayError)
			*aDisplayError = RuntimeError(ERR_DYNAMIC_UPVAR, aVarName);
		return nullptr;
	}
	ASSERT(!outer.mDownVar && !aInner.mUpVar); // PreprocessLocalVars has not yet been called.
	if (!(outer_var->Scope() & VAR_DOWNVAR))
	{
		outer.mDownVarCount++;
		outer_var->Scope() |= VAR_DOWNVAR;
	}
	// At this point aInner doesn't have a variable by this name, so create one:
	int insert_pos;
	aInner.mVars.Find(aVarName, &insert_pos);
	Var *inner_var = AddVar(aVarName, aVarNameLength, &aInner.mVars, insert_pos, VAR_LOCAL | VAR_DECLARED); // VAR_DECLARED suppresses checking for a global for #Warn LocalSameAsGlobal or when #Warn UseUnset is triggered.
	if (!inner_var)
	{
		if (aDisplayError)
			*aDisplayError = FAIL; // Error already displayed.
		return nullptr;
	}
	ASSERT(outer_var->Type() != VAR_CONSTANT || outer_var->HasObject() && dynamic_cast<UserFunc*>(outer_var->Object()));
	bool indefinite = outer_var->Type() == VAR_CONSTANT && !((UserFunc *)outer_var->Object())->mUpVarCount;
	if (!indefinite && ++aInner.mUpVarCount == 1) // This function's first definite upvar.
	{
		// Count all references to aInner as upvars (except the one in outer).
		CountNestedFuncRefs(outer, aInner.mName);
	}
	inner_var->SetAliasDirect(outer_var); // Temporarily point the upvar to its downvar for later processing.
	inner_var->MarkAssignedSomewhere(); // Upvars are not candidates for #Warn VarUnset.
	return inner_var;
}



void Script::CountNestedFuncRefs(UserFunc &aWithin, LPCTSTR aFuncName)
{
	for (int i = 0; i < aWithin.mVars.mCount; ++i) // There's no list of nested functions, so we look for them within mVars.
	{
		Var &nfv = *aWithin.mVars.mItem[i];
		if (nfv.Type() == VAR_CONSTANT && !nfv.IsAlias()) // nfv is a reference to a function defined within aWithin (an alias would be a reference to something from outside).
		{
			ASSERT(nfv.HasObject() && dynamic_cast<UserFunc*>(nfv.Object()));
			auto &nf = *(UserFunc *)nfv.Object();
			Var *uv = nf.mVars.Find(aFuncName);
			if (uv && uv->IsAlias()) // nf refers to the outer aFuncName (if not an alias, it would be a nested function with the same name).
			{
				nf.mUpVarCount++; // Count uv as an upvar, as caller has determined aFuncName is now a closure.
				if (nf.mUpVarCount == 1) // nf's first definite upvar, so nf is also now known to be a closure.
					CountNestedFuncRefs(*nf.mOuterFunc, nf.mName); // Count all references to nf as upvars (except the one in nf.mOuterFunc).
				if (uv->Scope() & VAR_DOWNVAR) // uv is also referenced by functions nested within nf,
					CountNestedFuncRefs(nf, aFuncName); // so recurse.
			}
		}
	}
}



Var *Script::AddVar(LPCTSTR aVarName, size_t aVarNameLength, VarList *aList, int aInsertPos, int aScope)
// Returns the address of the new variable or NULL on failure.
// Caller must ensure that aVarName isn't NULL and that this isn't a duplicate variable name.
// In addition, it has provided aInsertPos, which is the insertion point so that the list stays sorted.
{
	if (aVarNameLength > MAX_VAR_NAME_LENGTH)
	{
		ScriptError(_T("Variable name too long."), aVarName);
		return NULL;
	}

	// Make a temporary copy that includes only the first aVarNameLength characters from aVarName:
	TCHAR var_name[MAX_VAR_NAME_LENGTH + 1];
	tcslcpy(var_name, aVarName, aVarNameLength + 1);  // See explanation above.  +1 to convert length to size.

	if (*var_name && !Var::ValidateName(var_name))
		// Above already displayed error for us.  This can happen at loadtime or runtime (e.g. StringSplit).
		return NULL;

	if ((aScope & (VAR_LOCAL | VAR_DECLARED)) == VAR_LOCAL // This is an implicit local.
		&& g_Warn_LocalSameAsGlobal && !mIsReadyToExecute // Not enabled at runtime because of overlap with #Warn UseUnset, and dynamic assignments in assume-local mode are less likely to be intended global.
		&& FindGlobalVar(var_name, aVarNameLength))
		WarnLocalSameAsGlobal(var_name);

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

	Var *the_new_var = new Var(new_name, aScope);
	if (!the_new_var || !aList->Insert(the_new_var, aInsertPos))
	{
		MemoryError();
		return NULL;
	}

	if (aScope & VAR_LOCAL_FUNCPARAM)
		the_new_var->MarkAssignedSomewhere();

	return the_new_var;
}



Var *Script::FindOrAddBuiltInVar(LPCTSTR aVarName, VarEntry *aVarEntry)
{
	int insert_pos;
	if (Var *found = mVars.Find(aVarName, &insert_pos))
		return found;
	LPTSTR name = SimpleHeap::Malloc(aVarName);
	if (!name)
		return nullptr;
	Var *the_new_var = new Var(name, aVarEntry, VAR_DECLARE_GLOBAL);
	if (!the_new_var || !mVars.Insert(the_new_var, insert_pos))
	{
		MemoryError();
		return nullptr;
	}
	return the_new_var;
}



VarEntry *Script::GetBuiltInVar(LPCTSTR aVarName)
{
	// This array approach saves about 9KB on code size over the old approach
	// of a series of if's and _tcscmp calls, and performs about the same.
	// The "A_" prefix is omitted from each name string to reduce code size.
	if ((aVarName[0] == 'A' || aVarName[0] == 'a') && aVarName[1] == '_')
		aVarName += 2;
	else
		return nullptr;
	// Using binary search vs. linear search performs a bit better (notably for
	// rare/contrived cases like A_x%index%) and doesn't affect code size much.
	int left, right, mid, result;
	for (left = 0, right = _countof(g_BIV_A) - 1; left <= right;)
	{
		mid = (left + right) / 2;
		result = _tcsicmp(aVarName, g_BIV_A[mid].name);
		if (result > 0)
			left = mid + 1;
		else if (result < 0)
			right = mid - 1;
		else // Match found.
			return &g_BIV_A[mid];
	}
	// Since above didn't return:
	return nullptr;
}



WinGroup *Script::FindGroup(LPCTSTR aGroupName, bool aCreateIfNotFound)
// Caller must ensure that aGroupName isn't NULL.  But if it's the empty string, NULL is returned.
// Returns the Group whose name matches aGroupName.  If it doesn't exist, it is created if aCreateIfNotFound==true.
// Thread-safety: This function is thread-safe (except when called with aCreateIfNotFound==true) even when
// the main thread happens to be calling AddGroup() and changing the linked list while it's being traversed here
// by the hook thread.  However, any subsequent changes to this function or AddGroup() must be carefully reviewed.
{
	if (!*aGroupName)
	{
		if (aCreateIfNotFound)
			// An error message must be shown in this case since our caller is about to
			// exit the current script thread (and we don't want it to happen silently).
			ValueError(_T("Blank group name."), nullptr, FAIL);
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



ResultType Script::AddGroup(LPCTSTR aGroupName)
// Returns OK or FAIL.
// The caller must already have verified that this isn't a duplicate group.
// This function is not thread-safe because it adds an entry to the quasi-global list of window groups.
// In addition, if this function is being called by one thread while another thread is calling FindGroup(),
// the thread-safety notes in FindGroup() apply.
{
	size_t aGroupName_length = _tcslen(aGroupName);
	if (aGroupName_length > MAX_VAR_NAME_LENGTH)
		return ValueError(_T("Group name too long."), aGroupName, FAIL);
	if (!Var::ValidateName(aGroupName, DISPLAY_GROUP_ERROR)) // Seems best to use same validation as var names.
		return FAIL;

	LPTSTR new_name = SimpleHeap::Malloc(aGroupName, aGroupName_length);
	if (!new_name)
		return FAIL;  // It already displayed the error for us.

	// The precise method by which the follows steps are done should be thread-safe even if
	// some other thread calls FindGroup() in the middle of the operation.  But any changes
	// must be carefully reviewed:
	WinGroup *the_new_group = new WinGroup(new_name);
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
	for (Line *line = aStartingLine; line; line = line->mNextLine) // For each line.
	{
		// Skip any function bodies.
		while (line->mActionType == ACT_BLOCK_BEGIN && line->mAttribute)
			if (  !(line = line->mRelatedLine)  )
				return OK;
		if (line->mActionType == ACT_BLOCK_END && line->mAttribute)
			return OK; // End of this function.

		mCurrLine = line; // For error reporting in FindVar() and perhaps other places.

		if (line->mActionType == ACT_CATCH)
		{
			// Preparse Catch's output var at this stage so it has the proper assign-local effect.
			// Don't parse the class names yet since that might parse a given name as global when it
			// should be local due to a subsequent assignment.  Delaying it ensures CATCH's classes
			// are always consistent with other references in the function, and ensures that adding
			// local classes in future won't necessitate a change in behaviour.
			if (!PreparseCatchVar(line))
				return FAIL;
			continue;
		}

		for (int i = 0; i < line->mArgc; ++i) // For each arg.
		{
			ArgStruct &this_arg = line->mArg[i]; // For performance and convenience.
			if (!this_arg.is_expression) // Plain text; i.e. goto/break/continue label.
				continue;
			ASSERT(!this_arg.postfix);
			// Otherwise, convert the expression text to postfix form and set is_expression
			// based on whether the arg should be evaluated by ExpandExpression():
			if (!line->ExpressionToPostfix(this_arg)) // Doing this here, after the script has been loaded, might improve the compactness/adjacent-ness of the compiled expressions in memory, which might improve performance due to CPU caching.
				return FAIL; // The function above already displayed the error msg.
		}
	}
	return OK;
}



ResultType Script::PreparseExpressions(FuncList &aFuncs)
{
	for (int i = 0; i < aFuncs.mCount; ++i)
	{
		ASSERT(dynamic_cast<UserFunc*>(aFuncs.mItem[i]));
		auto &func = *(UserFunc *)aFuncs.mItem[i];
		g->CurrentFunc = &func;
		if (!PreparseExpressions(func.mJumpToLine)) // Preparse this function's body.
			return FAIL;
		// Nested functions will be preparsed next, due to the fact that they immediately
		// follow the outer function in aFuncs.
	}
	g->CurrentFunc = nullptr;
	return OK;
}



Line *Script::PreparseCommands(Line *aStartingLine)
// Preparse any commands which might rely on blocks having been fully preparsed,
// such as any command which has a jump target (label).
// Also perform some late-stage optimizations and validation.
{
	for (Line *line = aStartingLine; line; line = line->mNextLine)
	{
		LPTSTR line_raw_arg1 = LINE_RAW_ARG1; // Resolve only once to help reduce code size.
		LPTSTR line_raw_arg2 = LINE_RAW_ARG2; //
		
		switch (line->mActionType)
		{
		case ACT_BLOCK_BEGIN:
			if (line->mAttribute) // This is the opening brace of a function definition.
				g->CurrentFunc = (UserFunc *)line->mAttribute; // Must be set only for the above condition because functions can of course contain types of blocks other than the function's own block.
			break;
		case ACT_BLOCK_END:
			if (line->mAttribute) // This is the closing brace of a function definition.
			{
				auto &func = *(UserFunc *)line->mAttribute;

				g->CurrentFunc = func.mOuterFunc;

				Line *block_begin = line->mParentLine;
				Line *parent = block_begin->mParentLine;

				if (func.mIsFuncExpression // =>
					&& block_begin->mParentLine
					&& block_begin->mParentLine->mActionType != ACT_BLOCK_BEGIN)
				{
					if (line->mNextLine->mActionType == ACT_BLOCK_BEGIN) // It could only be a fat arrow block-begin under these conditions.
						// There's another =>function after this one (defined within the same
						// expression), so just continue until the last =>function is found.
						continue;
					// This fat arrow function's parent line is a statement with a single-line
					// action, but that action is currently separated from its parent by one or
					// more fat arrow functions.  It won't work that way because If/Else/Loop/etc.
					// all skip an initial ACT_BLOCK_BEGIN (to avoid an extra ExecUntil call),
					// which would result in executing the function's body instead of skipping it.
					Line *body = line->mNextLine;
				#ifdef KEEP_FAT_ARROW_FUNCTIONS_IN_LINE_LIST // Currently unused.
					block_begin = parent->mNextLine; // In case there are multiple fat arrow functions on one line.
					Line *after_body = parent->mRelatedLine;
					Line *body_end = after_body->mPrevLine; // In case body is multiple lines (such as a nested IF or LOOP).
					// Swap the statement body and fat arrow functions around to make it work:
					parent   ->mNextLine = body       , body       ->mPrevLine = parent;
					body_end ->mNextLine = block_begin, block_begin->mPrevLine = body_end;
					line     ->mNextLine = after_body , after_body ->mPrevLine = line;
				#else
					// Remove the fat arrow functions to allow the correct body to execute.
					// This relies on there being no need for the fat arrow functions to
					// remain in the Line list after this point (so for instance, there's
					// no possibility of setting a breakpoint in one of these functions).
					parent->mNextLine = body, body->mPrevLine = parent;
				#endif
				}
				else if (parent && parent->mActionType != ACT_BLOCK_BEGIN)
				{
					// This is a normal function definition directly under If/Else/Loop/etc.
					// It would execute incorrectly as described above, so raise an error.
					return block_begin->PreparseError(_T("Unexpected function"));
				}
			}
			break;
		case ACT_BREAK:
		case ACT_CONTINUE:
			{
				if (line->ArgHasDeref(1))
					// It seems unlikely that computing the target loop at runtime would be useful.
					// For simplicity, rule out things like "break(var)" and "break(func())":
					return line->PreparseError(ERR_PARAM1_INVALID); //_T("Target label of Break/Continue cannot be dynamic."));
				LPTSTR loop_name = line_raw_arg1;
				Label *loop_label = nullptr;
				Line *loop_line, *innermost_loop = nullptr;
				if (*loop_name && !IsNumeric(loop_name))
				{
					// Target is a named loop.
					if ( !(loop_label = FindLabel(loop_name)) )
						return line->PreparseError(ERR_NO_LABEL, loop_name);
					// loop_label->mJumpToLine is used only to identify the line in the loop below.
					// This approach ensures the loop is a valid target.  The previous approach using
					// innermost_loop->IsJumpValid() failed to detect cases where the target loop is
					// a valid goto target but doesn't enclose this BREAK/CONTINUE.
				}
				int n = loop_label ? 0 : *loop_name ? _ttoi(loop_name) : 1;
				// For n > 0, find the nth innermost loop which encloses this line.
				// Otherwise, determine whether loop_label points to a valid loop.
				for (loop_line = line->mParentLine; loop_line; loop_line = loop_line->mParentLine)
				{
					if (ACT_IS_LOOP(loop_line->mActionType)) // Any type of LOOP, FOR or WHILE.
					{
						if (!innermost_loop)
							innermost_loop = loop_line;
						if (--n == 0) // == not <=, so "break 0" and negative values are detected as invalid.
							break;
						if (loop_label && loop_label->mJumpToLine == loop_line)
							break;
					}
					if (loop_line->mActionType == ACT_BLOCK_BEGIN && loop_line->mAttribute
						|| loop_line->mActionType == ACT_FINALLY)
					{
						loop_line = nullptr;
						break; // Stop the search here since any outer loop would be an invalid target.
					}
				}
				if (!innermost_loop)
					return line->PreparseError(_T("Break/Continue must be enclosed by a Loop."));
				if (!loop_line)
					return line->PreparseError(ERR_PARAM1_INVALID);
				if (loop_line != innermost_loop)
				{
					if (line->mActionType == ACT_BREAK)
					{
						// Find the line to which we'll actually jump when the loop breaks.
						loop_line = loop_line->mRelatedLine; // From LOOP itself to the line after the LOOP's body.
						if (loop_line->mActionType == ACT_UNTIL)
							loop_line = loop_line->mNextLine;
					}
					line->mRelatedLine = loop_line;
				}
			}
			break;

		case ACT_GOTO:
			// This must be done here (i.e. *after* all the script lines have been added),
			// so that labels both above and below each Goto can be resolved.
			if (line->ArgHasDeref(1))
				// Since the jump-point contains a deref, it must be resolved at runtime:
				line->mRelatedLine = NULL;
  			else if (!line->GetJumpTarget(false))
				return NULL; // Error was already displayed by the called function.
			break;

		case ACT_HOTKEY_IF:
			PreparseHotkeyIfExpr(line);
			line->mActionType = ACT_RETURN;
			break; // No need for the ACT_RETURN validation performed below in this case.

		case ACT_RETURN:
			for (Line *parent = line->mParentLine; parent; parent = parent->mParentLine)
				if (parent->mActionType == ACT_FINALLY)
					return line->PreparseError(ERR_BAD_JUMP_INSIDE_FINALLY);
				else if (parent->mActionType == ACT_BLOCK_BEGIN && parent->mAttribute)
					break; // Allow "return" to be used inside a nested function or fat arrow function.
			break;

		case ACT_CATCH:
			if (!PreparseCatchClass(line))
				return nullptr;
			break;
		}

		// Finalize and optimize postfix expressions.
		for (int i = 0; i < line->mArgc; ++i)
			if (line->mArg[i].postfix && !line->FinalizeExpression(line->mArg[i]))
				return nullptr;

		// Check for unreachable code.
		if (g_Warn_Unreachable)
		switch (line->mActionType)
		{
		case ACT_RETURN:
		case ACT_BREAK:
		case ACT_CONTINUE:
		case ACT_GOTO:
		case ACT_THROW:
		// v2: ACT_EXIT is always from AddLine(ACT_EXIT), since a script calling Exit would produce ACT_EXPRESSION.
		// Exit could be parsed as both ACT_EXIT and as a function, but that would create minor inconsistencies
		// between "Exit" and "(x && Exit())", or Exit called directly vs. called via a separate function.
		case ACT_EXIT:
		//case ACT_EXITAPP: // Excluded since it's just a function in v2, and there can't be any expectation that the code following it will execute anyway.
			Line *next_line = line->mNextLine;
			if (!next_line // line is the script's last line.
				|| next_line->mParentLine != line->mParentLine) // line is the one-line action of if/else/loop/etc.
				break;
			while (next_line->mActionType == ACT_BLOCK_BEGIN && next_line->mAttribute)
				next_line = next_line->mRelatedLine; // Skip function body.
			switch (next_line->mActionType)
			{
			case ACT_EXIT: // v2: It's from an automatic AddLine(), so should be excluded.
			case ACT_BLOCK_END: // There's nothing following this line in the same block.
			case ACT_CASE:
				continue;
			}
			if (IsLabelTarget(next_line))
				break;
			TCHAR buf[64];
			sntprintf(buf, _countof(buf), _T("This line will never execute, due to %s preceding it."), g_act[line->mActionType].Name);
			ScriptWarning(g_Warn_Unreachable, buf, _T(""), next_line);
		}
	} // for()
	// Return something non-NULL to indicate success:
	return mLastLine;
}



ResultType Script::PreparseCatchVar(Line *aLine)
{
	if (!aLine->mArgc)
		return OK;
	auto args = SimpleHeap::Alloc<CatchStatementArgs>();
	args->output_var = nullptr;
	args->prototype = nullptr;
	args->prototype_count = 0;
	aLine->mAttribute = args;
	LPTSTR cp = aLine->mArg->text + aLine->mArg->length;
	for (LPTSTR min_space_position = aLine->mArg->text + 2;;)
	{
		--cp;
		if (cp < min_space_position)
			return OK;
		if (IS_SPACE_OR_TAB(*cp))
			break;
		if (!IS_IDENTIFIER_CHAR(*cp))
			//return *cp == '.' ? OK : aLine->LineError(ERR_EXPR_SYNTAX);
			return OK; // Leave error reporting to PreparseCatchClass().
	}
	LPTSTR var_name = cp + 1;
	cp = omit_trailing_whitespace(aLine->mArg->text + 1, cp);
	if (ctolower(*cp) == 's' && ctolower(cp[-1]) == 'a' && (cp - 1 == aLine->mArg->text || IS_SPACE_OR_TAB(cp[-2])))
	{
		Var *var;
		if (  !(var = FindOrAddVar(var_name, 0, FINDVAR_FOR_WRITE))
			|| !aLine->ValidateVarUsage(var, VARREF_OUTPUT_VAR)  )
			return FAIL;
		var->MarkAssignedSomewhere();
		args->output_var = var;
		aLine->mArg->length = ArgLengthType((cp - 1) - aLine->mArg->text);
	}
	return OK;
}



ResultType Script::PreparseCatchClass(Line *aLine)
{
	if (!aLine->mArgc || !aLine->mArg->length) // PreparseCatchVar() may have already set length to exclude " as output_var".
		return OK;
	IObject *prototype[MAX_ARGS - 1]; // Set an arbitrary limit to simplify the code; using MAX_ARGS to keep things consistent, -1 for the output var.
	int prototype_count = 0;
	for (LPTSTR cp = aLine->mArg->text, cp_end = cp + aLine->mArg->length; cp < cp_end; )
	{
		if (prototype_count)
		{
			if (*cp != g_delimiter)
				return aLine->LineError(ERR_EXPR_SYNTAX);
			if (prototype_count == _countof(prototype))
				return aLine->LineError(_T("Too many classes."));
			cp = omit_leading_whitespace(cp + 1);
		}
		LPTSTR end = cp;
		while (IS_IDENTIFIER_CHAR(*end) || *end == '.') ++end;
		LPTSTR next = omit_leading_whitespace(end);
		if (end == cp)
			return aLine->LineError(ERR_EXPR_SYNTAX);
		auto cls = FindClass(cp, end - cp);
		if (  !cls || !(prototype[prototype_count++] = cls->GetOwnPropObj(_T("Prototype")))  )
			return aLine->LineError(_T("Invalid class."), FAIL, cp);
		cp = next;
	}
	if (prototype_count) // Currently should always be non-zero due to the length check.
	{
		auto args = (CatchStatementArgs *)aLine->mAttribute;
		args->prototype = SimpleHeap::Alloc<IObject *>(prototype_count);
		args->prototype_count = prototype_count;
		memcpy(args->prototype, prototype, prototype_count * sizeof(IObject *));
	}
	return OK;
}



bool Script::IsLabelTarget(Line *aLine)
{
	Label *lbl = g->CurrentFunc ? g->CurrentFunc->mFirstLabel : mFirstLabel;
	for ( ; lbl && lbl->mJumpToLine != aLine; lbl = lbl->mNextLabel);
	return lbl;
}



ResultType Line::ExpressionToPostfix(ArgStruct &aArg)
{
	ExprTokenType *infix = NULL;
	ResultType result = ExpressionToPostfix(aArg, infix);
	// This approach produces smaller code than using RAII in the function below:
	free(infix);
	return result;
}

ResultType Line::ExpressionToPostfix(ArgStruct &aArg, ExprTokenType *&aInfix)
// Returns OK or FAIL.
{
	// Having a precedence array is required at least for SYM_POWER (since the order of evaluation
	// of something like 2**1**2 does matter).  It also helps performance by avoiding unnecessary pushing
	// and popping of operators to the stack. This array must be kept in sync with "enum SymbolType".
	// Also, dimensioning explicitly by SYM_COUNT helps enforce that at compile-time:
	static UCHAR sPrecedence[SYM_COUNT] =  // Performance: UCHAR vs. INT benches a little faster, perhaps due to the slight reduction in code size it causes.
	{
		0,0,0,0,0,0,0,0,0// SYM_STRING, SYM_INTEGER, SYM_FLOAT, SYM_MISSING, SYM_VAR, SYM_OBJECT, SYM_DYNAMIC, SYM_SUPER, SYM_BEGIN (SYM_BEGIN must be lowest precedence).
		, 82, 82         // SYM_POST_INCREMENT, SYM_POST_DECREMENT: Highest precedence operator so that it will work even though it comes *after* a variable name (unlike other unaries, which come before).
		, 86             // SYM_DOT
		, 2,2,2,2,2,2    // SYM_CPAREN, SYM_CBRACKET, SYM_CBRACE, SYM_OPAREN, SYM_OBRACKET, SYM_OBRACE (to simplify the code, parentheses/brackets/braces must be lower than all operators in precedence).
		, 6              // SYM_COMMA -- Must be just above SYM_OPAREN so it doesn't pop OPARENs off the stack.
		, 7,7,7,7,7,7,7,7,7,7,7,7,7  // SYM_ASSIGN_*. THESE HAVE AN ODD NUMBER to indicate right-to-left evaluation order, which is necessary for cascading assignments such as x:=y:=1 to work.
//		, 8              // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
		, 11, 11         // SYM_IFF_ELSE, SYM_IFF_THEN (ternary conditional).  HAS AN ODD NUMBER to indicate right-to-left evaluation order, which is necessary for ternaries to perform traditionally when nested in each other without parentheses.
//		, 12             // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
		, 15             // SYM_OR_MAYBE -- Right-associative as below.
		, 17             // SYM_OR -- Right-associative so that short-circuit skips the entire right branch, instead of evaluating each one in sequence with the same result.
		, 21             // SYM_AND -- As above.
//		, 25             // Reserved for SYM_LOWNOT.
//		, 26             // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
		, 28             // SYM_IS
		, 30, 30, 30, 30 // SYM_EQUAL, SYM_EQUALCASE, SYM_NOTEQUAL, SYM_NOTEQUALCASE (lower prec. than the below so that "x < 5 = var" means "result of comparison is the boolean value in var".
		, 34, 34, 34, 34 // SYM_GT, SYM_LT, SYM_GTOE, SYM_LTOE
		, 36             // SYM_REGEXMATCH
		, 38             // SYM_CONCAT
		, 4				 // SYM_LOW_CONCAT
		, 42             // SYM_BITOR -- Seems more intuitive to have these three higher in prec. than the above, unlike C and Perl, but like Python.
		, 46             // SYM_BITXOR
		, 50             // SYM_BITAND
		, 54, 54, 54     // SYM_BITSHIFTLEFT, SYM_BITSHIFTRIGHT, SYM_BITSHIFTRIGHT_LOGICAL
		, 62			 // SYM_INTEGERDIVIDE
		, 58, 58         // SYM_ADD, SYM_SUBTRACT
		, 62, 62	     // SYM_MULTIPLY, SYM_DIVIDE
		, 73             // SYM_POWER (see note below). Changed to right-associative for v2.
		, 25             // SYM_LOWNOT (the word "NOT": the low precedence version of logical-not).  HAS AN ODD NUMBER to indicate right-to-left evaluation order so that things like "not not var" are supports (which can be used to convert a variable into a pure 1/0 boolean value).
		, 67,67,67,67,67 // SYM_NEGATIVE (unary minus), SYM_POSITIVE (unary plus), SYM_REF, SYM_HIGHNOT (the high precedence "!" operator), SYM_BITNOT
		// NOTE: THE ABOVE MUST BE AN ODD NUMBER to indicate right-to-left evaluation order, which was added in v1.0.46 to support consecutive unary operators such as !*var !!var (!!var can be used to convert a value into a pure 1/0 boolean).
//		, 68             // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
		, 86             // SYM_ISSET
		, 77, 77         // SYM_PRE_INCREMENT, SYM_PRE_DECREMENT (higher precedence than SYM_POWER because it doesn't make sense to evaluate power first because that would cause ++/-- to fail due to operating on a non-lvalue.
//		, 78             // THIS VALUE MUST BE LEFT UNUSED so that the one above can be promoted to it by the infix-to-postfix routine.
//		, 82, 82         // RESERVED FOR SYM_POST_INCREMENT, SYM_POST_DECREMENT (which are listed higher above for the performance of YIELDS_AN_OPERAND().
		, 86             // SYM_FUNC -- Has special handling which ensures it stays tightly bound with its parameters as though it's a single operand for use by other operators; the actual value here is irrelevant.
		, 0, 0           // SYM_RESERVED_*
	};
	// Most programming languages give exponentiation a higher precedence than unary minus and logical-not.
	// For example, -2**2 is evaluated as -(2**2), not (-2)**2 (the latter is unsupported by qmathPow anyway).
	// However, this rule requires a small workaround in the postfix-builder to allow 2**-2 to be
	// evaluated as 2**(-2) rather than being seen as an error.  v1.0.45: A similar thing is required
	// to allow the following to work: 2**!1, 2**not 0, 2**~0xFFFFFFFE, 2**&x.

	// IsSet is constructed here because it's a sort of intrinsic function with its own
	// special rules.  ExprOp<> isn't used because it doesn't have parameter count limits
	// or a name (which is displayed when the parameter count is invalid, for instance).
	static BuiltInFunc *sIsSetFunc = new BuiltInFunc { _T("IsSet"), BIF_IsSet, 1, 1 };

	ExprTokenType *infix = NULL;
	int infix_size = 0, infix_count = 0, allow_for_extra_postfix = 0;
	const int INFIX_GROWTH = 128; // Amount to grow by each time expansion is needed.  Rarely needed more than once (from zero).
	const int INFIX_MIN_SPACE = 5; // Minimum space to allow prior to each iteration or deref-processing.  Room for auto-concat, two tokens for `.id`, `. ("s")` for DT_STRING, and `f(fn)` for DT_FUNCREF; plus the final SYM_INVALID.

	///////////////////////////////////////////////////////////////////////////////////////////////
	// TOKENIZE THE INFIX EXPRESSION INTO AN INFIX ARRAY: Avoids the performance overhead of having
	// to re-detect whether each symbol is an operand vs. operator at multiple stages.
	///////////////////////////////////////////////////////////////////////////////////////////////
	LPTSTR op_end, cp;
	DerefType *deref, *this_deref;
	int cp1; // int vs. char benchmarks slightly faster, and is slightly smaller in code size.
	TCHAR number_buf[MAX_NUMBER_SIZE];

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
				if (infix_count + INFIX_MIN_SPACE > infix_size)
				{
					infix_size += INFIX_GROWTH;
					if (void *p = realloc(infix, infix_size * sizeof(ExprTokenType)))
						aInfix = infix = (ExprTokenType *)p;
					else
						return LineError(ERR_OUTOFMEM);
				}

				cp = omit_leading_whitespace(cp);
				if (!*cp // Very end of expression...
					|| this_deref && cp >= this_deref->marker) // ...or no more literal/raw text left to process at the left side of this_deref.
					break; // Break out of inner loop so that bottom of the outer loop will process this_deref itself.

				ExprTokenType &this_infix_item = infix[infix_count]; // Might help reduce code size since it's referenced many places below.
				this_infix_item.callsite = nullptr; // Init needed for SYM_ASSIGN and related; a non-NULL value means it should be converted to an object-assignment.
				// Above: SYM_REF and possibly others may rely on this zero-initializing the union.
				this_infix_item.error_reporting_marker = cp; // Used for reporting syntax errors with unary and binary operators, and may be overwritten via union for other symbols.

				// Auto-concat requires a space or tab for the following reasons:
				//  - Readability.
				//  - Compliance with the documentation.
				//  - To simplify handling of expressions like x[y]().
				//  - To reserve other combinations for future use; e.g. R"raw string".
#define CHECK_AUTO_CONCAT \
				if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol)) \
				{ \
					if (!IS_SPACE_OR_TAB(cp[-1])) \
						return LineError(ERR_BAD_AUTO_CONCAT, FAIL, cp); \
					infix[infix_count++].symbol = SYM_CONCAT; \
				}

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
								// UPDATE: Quoted literal strings now explicitly support the above, but not
								// other operands such as 123 ++x, which would be vanishingly rare.
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
						else // Unary plus.
							this_infix_item.symbol = SYM_POSITIVE; // Added in v2 for symmetry with SYM_NEGATIVE; i.e. if -x produces NaN, so should +x.
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
							_tcstod(cp, &op_end); // Find the end of this number, the easy way.
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
								// Use a temp variable because unquoted_literal requires that op_end be set properly:
								LPTSTR pow_temp = omit_leading_whitespace(op_end);
								if (!(pow_temp[0] == '*' && pow_temp[1] == '*'))
									goto unquoted_literal; // Goto is used for performance and also as a patch to minimize the chance of breaking other things via redesign.
								//else it's followed by pow.  Since pow is higher precedence than unary minus,
								// leave this unary minus as an operator so that it will take effect after the pow.
							}
							//else it appears to be a double deref; invalid, since names beginning with a digit are
							// prohibited.  Either way it would be detected as an error, but interpreting it as a
							// double-deref gives it a more intuitive error message (than ERR_BAD_AUTO_CONCAT).
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
							this_infix_item.symbol = SYM_ASSIGN_INTEGERDIVIDE;
						}
						else
						{
							++cp; // An additional increment to have loop skip over the second '/' too.
							this_infix_item.symbol = SYM_INTEGERDIVIDE;
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
							this_infix_item.symbol = SYM_MULTIPLY;
					}
					break;
				case '!':
					if (cp1 == '=') // i.e. !=
						// An additional increment for each '=' to have loop skip over the '=' too.
						++cp, this_infix_item.symbol	= cp[1] == '=' // note, cp[1] is not equal to cp1 here due to ++cp
														? (++cp, SYM_NOTEQUALCASE)	// !==
														: SYM_NOTEQUAL;				// != 
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
						&& IS_SPACE_OR_TAB(cp[-1])) // If there's no space, assume it's something valid like "new Class()" until it can be proven otherwise.
					{
						infix[infix_count++].symbol = SYM_CONCAT;
					}
					infix[infix_count].symbol = SYM_OPAREN; // MUST NOT REFER TO this_infix_item IN CASE ABOVE DID ++infix_count.
					infix[infix_count].marker = cp;
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
						// Below will change the SYM_DOT token into SYM_OBRACKET, keeping the existing callsite struct.
					}
					else
					{
						auto callsite = new CallSite();
						if (  !(infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))  )
							callsite->func = ExprOp<Op_Array, 0>();
						else
							callsite->flags = IT_GET; // This may be overridden by standard_pop_into_postfix.
						this_infix_item.callsite = callsite;
					}
					// DO NOT USE this_infix_item in case above did --infix_count.
					infix[infix_count].symbol = SYM_OBRACKET;
					break;
				case ']': // L31
					this_infix_item.symbol = SYM_CBRACKET;
					break;
				case '{':
					if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))
						return LineError(_T("Unexpected \"{\""));
					this_infix_item.callsite = new CallSite();
					this_infix_item.callsite->func = ExprOp<Op_Object, 0>();
					this_infix_item.symbol = SYM_OBRACE;
					break;
				case '}':
					this_infix_item.symbol = SYM_CBRACE;
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
							if (cp[1] == '>') // look for a third '>'
							{
								++cp; // to have the loop skip the third '>'
								// it is SYM_BITSHIFTRIGHT_LOGICAL or SYM_ASSIGN_BITSHIFTRIGHT_LOGICAL
								this_infix_item.symbol = cp[1] == '=' ? (cp++, SYM_ASSIGN_BITSHIFTRIGHT_LOGICAL) : SYM_BITSHIFTRIGHT_LOGICAL;
							}
							else
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
						// Differentiate between unary "take a reference to" and the "bitwise and" operator:
						// See '-' above for more details:
						if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))
							this_infix_item.symbol = SYM_BITAND;
						else
						{
							this_infix_item.symbol = SYM_REF;
							this_infix_item.var_usage = VARREF_READ; // Set default: assume a VarRef is needed.
						}
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
						this_infix_item.callsite = new CallSite();
						this_infix_item.callsite->func = ExprOp<BIF_RegEx, FID_RegExMatch>();
						this_infix_item.callsite->param_count = 2;
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
					if (cp1 == '?')
					{
						if ( !(infix_count && (infix[infix_count - 1].symbol == SYM_VAR || infix[infix_count - 1].symbol == SYM_DYNAMIC)) )
							return LineError(ERR_EXPR_SYNTAX, FAIL, cp);
						infix[infix_count - 1].var_usage = VARREF_READ_MAYBE; // `var ??` implies that no error should be raised if var is unset.
						++cp; // An additional increment to have loop skip over the operator's second symbol.
						this_infix_item.symbol = SYM_OR_MAYBE;
						break;
					}
					if (IS_SPACE_OR_TAB(cp1))
						cp1 = *omit_leading_whitespace(cp + 1);
					if (cp1 == ')' || cp1 == ',' || cp1 == ']' || cp1 == '}' || cp1 == ':')
					{
						// Under the conditions checked above, this '?' can't be a valid ternary operator.
						// If there's a variable to the left, this is probably a valid "maybe" reference.
						// Future use: other operations such as x[y]?.
						if ( !(infix_count && (infix[infix_count - 1].symbol == SYM_VAR || infix[infix_count - 1].symbol == SYM_DYNAMIC)) )
							return LineError(ERR_EXPR_SYNTAX, FAIL, cp);
						infix[infix_count - 1].var_usage = VARREF_READ_MAYBE;
						++cp; // Discard this '?'.
						--infix_count; // Counter the loop's increment.
						continue;
					}
					this_infix_item.symbol = SYM_IFF_THEN;
					this_infix_item.marker = cp; // For error-reporting.
					break;
				case ':':
					if (cp1 == '=')
					{
						++cp; // An additional increment to have loop skip over '=' too.
						this_infix_item.symbol = SYM_ASSIGN;
					}
					else
					{
						this_infix_item.symbol = SYM_IFF_ELSE;
						this_infix_item.marker = cp; // For detection of invalid object literals, and error-reporting.
					}
					break;

				case '"': // QUOTED/LITERAL STRING.
				case '\'':
					// This is the opening quote mark of a literal string.  The string itself will be
					// handled via the "deref" list.
					infix_count--; // Counter the loop's increment.
					break;

				case g_DerefChar:
					// Deref char within a string or double-deref; acts as concat.
					this_infix_item.symbol = SYM_LOW_CONCAT;
					break;

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
								return LineError(ERR_EXPR_SYNTAX, FAIL, cp);
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
							if (this_deref && this_deref->marker == cp + 1) // Dynamic member such as x.%y% or x.y%z%.
							{
								--infix_count; // Counter the loop's increment.
								break;
							}

							// Skip this '.'
							++cp;

							// Find the end of the operand (".operand"):
							op_end = find_identifier_end(cp);
							if (!_tcschr(EXPR_OPERAND_TERMINATORS, *op_end))
								return LineError(ERR_EXP_ILLEGAL_CHAR, FAIL, op_end);

							if (op_end == cp) // Missing identifier.
								return LineError(ERR_EXPR_SYNTAX, FAIL, cp-1); // Intentionally vague since the user's intention isn't clear.

							auto callsite = new CallSite();
							callsite->member = SimpleHeap::Alloc(cp, op_end - cp);

							SymbolType new_symbol; // Type of token: SYM_FUNC or SYM_DOT (which must be treated differently as it doesn't have parentheses).
							if (*op_end == '(')
							{
								new_symbol = SYM_FUNC;
								callsite->flags = IT_CALL;
								// DON'T DO THE FOLLOWING - must let next iteration handle '(' so it outputs a SYM_OPAREN:
								//++op_end;
							}
							else
							{
								new_symbol = SYM_DOT; // This will be changed to SYM_FUNC at a later stage.
								callsite->flags = IT_GET; // Set default; may be overridden by standard_pop_into_postfix.
							}

							// Output the operator next - after the operand to avoid auto-concat.
							infix[infix_count].symbol = new_symbol;
							infix[infix_count].callsite = callsite;

							// Continue processing after this operand. Outer loop will do ++infix_count.
							cp = op_end;
							continue;
						}
					}

unquoted_literal:
					// This operand is a normal raw numeric-literal, or an unquoted literal string/key in
					// an object literal, such as "{key: value}".  Word operators such as AND/OR/NOT/NEW
					// and variable/function references don't reach this point as they are pre-parsed by
					// ParseOperands() and placed into the "deref" array.  Unrecognized symbols should be
					// impossible at this stage because prior validation would have caught them.
					CHECK_AUTO_CONCAT;
					// MUST NOT REFER TO this_infix_item IN CASE ABOVE DID ++infix_count:
					ExprTokenType &this_literal = infix[infix_count];

					if (*cp <= '9' && (*cp >= '0' || *cp == '+' || *cp == '-' || *cp == '.'))
					{
						// Looks like a number.  The checks above and IsHex() below rule out the strings
						// "inf", "infinity", "nan" and "nanxxx", and hexadecimal floating-point numbers,
						// which would otherwise be interpreted as valid if supported by the compiler
						// (VC++ 2015 and later).  This is done because we don't support those values
						// elsewhere in the code, such as in IsNumeric().
						LPTSTR i_end, d_end;
						__int64 i = istrtoi64(cp, &i_end);
						if (!IsHex(cp))
						{
							double d = _tcstod(cp, &d_end);
							if (d_end > i_end && _tcschr(EXPR_OPERAND_TERMINATORS, *d_end))
							{
								this_literal.symbol = SYM_FLOAT;
								this_literal.value_double = d;
								cp = d_end; // Have the loop process whatever lies at d_end and beyond.
								continue;
							}
						}
						if (*cp == '.') // Must be checked to avoid `(.foo)` being interpreted as `((0).foo)`.
							return LineError(ERR_EXPR_SYNTAX, FAIL, cp);
						if (_tcschr(EXPR_OPERAND_TERMINATORS, *i_end)
							// Exclude property names composed of digits, in an object literal:
							&& !(*omit_leading_whitespace(i_end) == ':' && infix_count
								&& (infix[infix_count-1].symbol == SYM_OBRACE || infix[infix_count-1].symbol == SYM_COMMA)))
						{
							this_literal.symbol = SYM_INTEGER;
							this_literal.value_int64 = i;
							cp = i_end; // Have the loop process whatever lies at i_end and beyond.
							continue;
						}
					}
					op_end = find_identifier_end(cp);
					// SYM_STRING: either the "key" in "{key: value}" or a syntax error (might be impossible).
					LPTSTR str = SimpleHeap::Alloc(cp, op_end - cp);
					this_literal.SetValue(str, op_end - cp);
					cp = op_end; // Have the loop process whatever lies at op_end and beyond.
					continue; // "Continue" to avoid the ++cp at the bottom.
				} // switch() for type of symbol/operand.
				if (IS_ASSIGNMENT_EXCEPT_POST_AND_PRE(this_infix_item.symbol))
					++allow_for_extra_postfix;
				else if (this_infix_item.symbol == SYM_POST_INCREMENT || this_infix_item.symbol == SYM_POST_DECREMENT)
					allow_for_extra_postfix += 4;
				else if (this_infix_item.symbol == SYM_PRE_INCREMENT || this_infix_item.symbol == SYM_PRE_DECREMENT)
					allow_for_extra_postfix += 2;
				++cp; // i.e. increment only if a "continue" wasn't encountered somewhere above. Although maintainability is reduced to do this here, it avoids dozens of ++cp in other places.
			} // for each token in this section of raw/literal text.
		} // End of processing of raw/literal text (such as operators) that lie to the left of this_deref.

		if (!this_deref) // All done because the above just processed all the raw/literal text (if any) that
			break;       // lay to the right of the last deref.

		if (infix_count + INFIX_MIN_SPACE > infix_size)
		{
			infix_size += INFIX_GROWTH;
			if (void *p = realloc(infix, infix_size * sizeof(ExprTokenType)))
				aInfix = infix = (ExprTokenType *)p;
			else
				return LineError(ERR_OUTOFMEM);
		}

		// THE ABOVE HAS NOW PROCESSED ANY/ALL RAW/LITERAL TEXT THAT LIES TO THE LEFT OF this_deref.
		// SO NOW PROCESS THIS_DEREF ITSELF.
		DerefType &this_deref_ref = *this_deref; // Boosts performance slightly.
		if (this_deref_ref.type == DT_STRING || this_deref_ref.type == DT_QSTRING)
		{
			bool is_end_of_string = this_deref_ref.terminal;
			bool is_start_of_string = this_deref_ref.substring_count == 1;
			bool require_paren = this_deref_ref.substring_count > 1 || !this_deref_ref.terminal;

			cp = this_deref_ref.marker; // This is done in case omit_leading_whitespace() skipped over leading whitespace of a string.

			if (infix_count && YIELDS_AN_OPERAND(infix[infix_count - 1].symbol))
			{ 
				LPTSTR start_of_literal = cp - (this_deref_ref.type == DT_QSTRING);
				TCHAR c = start_of_literal[-1];
				// See CHECK_AUTO_CONCAT macro definition for comments.
				if (c != '.') // i.e. it's not a dynamic member like .my%x%
				{
					if (!IS_SPACE_OR_TAB(c))
						return LineError(ERR_BAD_AUTO_CONCAT, FAIL, start_of_literal);
					infix[infix_count++].symbol = SYM_CONCAT;
				}
			}

			bool can_be_optimized_out = this_deref_ref.length == 0 && this_deref_ref.type != DT_QSTRING;
			if (can_be_optimized_out)
			{
				// This substring is empty; can we optimize it out along with one concat op?
				if (*cp == g_DerefChar)
					// Skip the deref char, which otherwise produces SYM_LOW_CONCAT.  Can't use cp++ because we need
					// to ensure the value stored infix[].marker below doesn't end up pointing to an open parenthesis,
					// such as for obj.%(methodexpr)%().  length can be repurposed because this deref will be discarded.
					this_deref_ref.length = 1; 
				else if (infix_count && infix[infix_count - 1].symbol == SYM_LOW_CONCAT)
					infix_count--; // Undo the concat op.
				else
					can_be_optimized_out = false;
			}

			if (require_paren && is_start_of_string)
			{
				infix[infix_count].symbol = SYM_OPAREN;
				infix[infix_count].marker = cp;
				infix_count++;
			}

			if (!can_be_optimized_out)
			{
				LPTSTR str = SimpleHeap::Alloc(this_deref_ref.marker, this_deref_ref.length);
				infix[infix_count].SetValue(str, this_deref_ref.length);
				infix_count++;
			}
			cp += this_deref_ref.length;

			if (is_end_of_string)
			{
				if (this_deref_ref.type == DT_QSTRING)
				{
					cp = omit_leading_whitespace(cp + 1);
					if (*cp && _tcschr(_T("+-*&~!"), *cp) && cp[1] != '=' && (cp[1] != '&' || *cp != '&'))
					{
						// The symbol following this literal string is either a unary operator or a
						// binary operator which can't (at least logically) be applied to a literal
						// string. Since the user's intention isn't clear, treat it as a syntax error.
						// The most common cases where this helps are:
						//	MsgBox % "var's address is " &var  ; Misinterpreted as SYM_BITAND.
						//	MsgBox % "counter is now " ++var   ; Misinterpreted as SYM_POST_INCREMENT.
						return LineError(_T("Unexpected operator following literal string."), FAIL, cp);
					}
				}
				if (require_paren)
				{
    				infix[infix_count].symbol = SYM_CPAREN;
					infix_count++;
				}
			}
			else
			{
				// *cp must be either % (indicating an expression within the string) or the end of
				// the arg (implying this is a top-level unquoted string and !aArg.is_expression),
				// except that an optimization above may have skipped the '%'.
				ASSERT(*cp == g_DerefChar || !*cp || can_be_optimized_out && cp[-1] == g_DerefChar);
			}
			// Counter the loop's increment.  It's done this way for simplicity, since a variable
			// number of tokens are created by this section and the conditions are a bit convoluted.
			infix_count--;
			continue;
		}
		else if (this_deref_ref.type == DT_DOUBLE) // Marks the end of a var double-dereference.
		{
			infix[infix_count].symbol = SYM_DYNAMIC;
			infix[infix_count].var = nullptr; // Indicate this is a double-deref.
			infix[infix_count].var_usage = VARREF_READ; // Set default.
		}
		else if (this_deref_ref.type == DT_DOTPERCENT)
		{
			auto callsite = new CallSite();
			if (*this_deref_ref.marker == '(')
			{
				infix[infix_count].symbol = SYM_FUNC;
				callsite->flags = IT_CALL | EIF_STACK_MEMBER;
			}
			else
			{
				infix[infix_count].symbol = SYM_DOT;
				callsite->flags = IT_GET | EIF_STACK_MEMBER;
			}
			infix[infix_count].callsite = callsite;
		}
		else if (this_deref_ref.type == DT_WORDOP)
		{
			if (this_deref_ref.symbol == SYM_SUPER)
			{
				// These checks ensure SYM_SUPER doesn't need to be specifically handled by
				// ExpandExpression(); it will be pushed onto the stack due to IS_OPERAND()
				// and handled directly by SYM_FUNC.
				LPTSTR next_op = cp + this_deref_ref.length;
				if (*next_op != '(') // Permit super() to mean super.call(), but do not permit "super (1)".
				{
					next_op = omit_leading_whitespace(next_op); // The "." and "[]" operators permit leading spaces.
					if (*next_op != '.' && *next_op != '[') // Permit super.x and super[i].
						return LineError(ERR_EXPR_SYNTAX, FAIL, cp);
				}
				if (!g->CurrentFunc || !g->CurrentFunc->mClass)
					return LineError(_T("\"") SUPER_KEYWORD _T("\" is valid only inside a class."), FAIL, cp);
				CHECK_AUTO_CONCAT;
			}
			else if (this_deref_ref.symbol == SYM_ISSET)
			{
				if (cp[this_deref_ref.length] != '(')
					return LineError(ERR_MISSING_OPEN_PAREN, FAIL, cp);
				CHECK_AUTO_CONCAT;
				infix[infix_count].callsite = new CallSite();
				infix[infix_count].callsite->func = sIsSetFunc;
				this_deref_ref.symbol = SYM_FUNC;
			}
			infix[infix_count].symbol = this_deref_ref.symbol;
			infix[infix_count].error_reporting_marker = this_deref_ref.marker;
		}
		else if (this_deref_ref.type == DT_CONST_INT)
		{
			CHECK_AUTO_CONCAT;
			infix[infix_count].SetValue(this_deref_ref.int_value);
		}
		else // this_deref is a variable (DT_VAR) or fat arrow function (DT_FUNCREF).
		{
			CHECK_AUTO_CONCAT;
			// Use SYM_VAR for both built-in and normal vars at this stage, since some normal vars
			// might become read-only later on anyway.  This simplifies some parts, as non-dynamic
			// references are always SYM_VAR and SYM_DYNAMIC has only one meaning.  A later stage
			// will validate and translate each SYM_VAR as needed.
			infix[infix_count].symbol = SYM_VAR;
			infix[infix_count].var_deref = &this_deref_ref;
			infix[infix_count].var_usage = VARREF_READ; // Set default, for detection of how this var ref is used.
		} // Handling of the var or function in this_deref.

		// Finally, jump over the dereference text. Note that in the case of an expression, there might not
		// be any percent signs within the text of the dereference, e.g. x + y, not %x% + %y% (unless they're
		// deliberately double-derefs).
		cp += this_deref_ref.length;
		// The outer loop will now do ++infix for us.
	} // For each deref in this expression, and also for the final literal/raw text to the right of the last deref.

	if (infix_count == 0)
		// Probably something like "Loop Parse, foo, `n" since an empty expression wouldn't make it this far.
		// infix_size is 0 in this case, so cannot continue.  An alternative would be to treat this as a blank
		// parameter, but it's probably more useful to treat it as an error (maybe "`n" was intended).
		return LineError(ERR_EXPR_SYNTAX);

	if (aArg.type == ARG_TYPE_OUTPUT_VAR)
	{
		if (infix_count != 1 || infix->symbol != SYM_VAR)
			return LineError(ERR_EXPR_SYNTAX, FAIL, aArg.text);
		infix->var_usage = VARREF_OUTPUT_VAR;
	}

	// Terminate the array with a special item.  This allows infix-to-postfix conversion to do a faster
	// traversal of the infix array.
	ASSERT(infix_count < infix_size);
	infix[infix_count].symbol = SYM_INVALID;

	////////////////////////////
	// CONVERT INFIX TO POSTFIX.
	////////////////////////////

	// postfix may need to be a bit bigger than infix: Allow an extra slot for each compound assignment,
	// in case it targets an object/property and therefore needs to be expanded.  Also allow an extra slot
	// to be temporarily occupied by a short-circuit operator popped from the stack.
	int max_postfix_count = infix_count + allow_for_extra_postfix + 1;
	ExprTokenType **postfix = (ExprTokenType **)_alloca(max_postfix_count * sizeof(ExprTokenType *));
	ExprTokenType **stack = (ExprTokenType **)_alloca((infix_count + 1) * sizeof(ExprTokenType *));  // +1 for SYM_BEGIN on the stack.
	int postfix_count = 0, stack_count = 0;
	// Above dimensions the stack to be as large as the infix/postfix arrays to cover worst-case
	// scenarios and avoid having to check for overflow.  For the infix-to-postfix conversion, the
	// stack must be large enough to hold a malformed expression consisting entirely of operators
	// (though other checks might prevent this).

#ifdef _DEBUG
#undef STACK_PUSH
#define STACK_PUSH(token_ptr) (ASSERT(stack_count < infix_count), stack[stack_count++] = (token_ptr))
#endif

	// SYM_BEGIN is the first item to go on the stack.  It's a flag to indicate that conversion to postfix has begun:
	ExprTokenType token_begin;
	token_begin.symbol = SYM_BEGIN;
	STACK_PUSH(&token_begin);

	SymbolType stack_symbol, infix_symbol, sym_prev, sym_next;
	ExprTokenType *this_infix = infix;
	CallSite *in_param_list = nullptr; // The function call site of the parameter list which directly contains the current token.

	for (;;) // While SYM_BEGIN is still on the stack, continue iterating.
	{
		ASSERT(postfix_count < max_postfix_count);
		ExprTokenType *&this_postfix = postfix[postfix_count]; // Resolve early, especially for use by "goto". Reduces code size a bit, though it doesn't measurably help performance.
		infix_symbol = this_infix->symbol;                     //
		stack_symbol = stack[stack_count - 1]->symbol; // Frequently used, so resolve only once to help performance.

		// Put operands into the postfix array immediately, then move on to the next infix item:
		if (IS_OPERAND(infix_symbol))
		{
			this_postfix = this_infix++;
			++postfix_count;
			continue; // Doing a goto to a hypothetical "standard_postfix" (in lieu of these last 3 lines) reduced performance and didn't help code size.
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
				if (stack_symbol == SYM_BEGIN // Not sure if this is possible.
					|| IS_OPAREN_LIKE(stack_symbol)) // Mismatched parens/brackets/braces.
				{
					// This should never happen due to balancing done by GetLineContExpr()/BalanceExpr().
					return LineError(ERR_EXPR_SYNTAX);
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
			//  b) Allow empty parameters by inserting the SYM_MISSING marker.
			//  c) Optimize DllCalls by pre-resolving common function names.
			if (!in_param_list)
			{
				sym_prev = this_infix[-1].symbol; // There's always at least one token preceding this one.
				if (sym_prev == SYM_OPAREN || sym_prev == SYM_COMMA) // () or ,)
					return LineError(ERR_EXPR_SYNTAX, FAIL, this_infix->error_reporting_marker);
			}
			else if (IS_OPAREN_LIKE(stack_symbol))
			{
				if (infix_symbol == SYM_COMMA || this_infix[-1].symbol != stack_symbol) // i.e. not an empty parameter list.
				{
					// Accessing this_infix[-1] here is necessarily safe since in_param_list is
					// non-NULL, and that can only be the result of a previous SYM_OPAREN/BRACKET.
					SymbolType prev_sym = this_infix[-1].symbol;
					if (prev_sym == SYM_COMMA || prev_sym == stack_symbol) // Empty parameter.
					{
						int num_blank_params = 0;
						while (this_infix->symbol == SYM_COMMA)
						{
							++this_infix;
							++num_blank_params;
						}
						infix_symbol = this_infix->symbol; // In case this_infix changed above.
						if (!IS_OPAREN_MATCHING_CPAREN(stack_symbol, infix_symbol))
						{
							for (int i = 0; i < num_blank_params; ++i)
							{
								// Do not reuse the infix token, since checks in other places depend on it being
								// identified as SYM_COMMA.  Using _alloca() inexplicably produces somewhat smaller
								// code than setting it up like token_begin, at least at the time of this comment.
								ExprTokenType *missing = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
								missing->symbol = SYM_MISSING;
								postfix[postfix_count++] = missing;
							}
							in_param_list->param_count += num_blank_params;
							// Go back to the top to update the this_postfix ref.
							continue;
						}
						// Since above didn't "continue", the last blank parameter is at the end of the list (and that's
						// terminated with the correct symbol), so there's no need to put anything in postfix or adjust
						// param_count.  this_infix has already been adjusted to discard the blank parameters.
					}
					else
					{
						// This is SYM_COMMA or SYM_CPAREN/BRACKET/BRACE at the end of a parameter.
						++in_param_list->param_count;

						if (stack_symbol == SYM_OBRACE && (in_param_list->param_count & 1)) // i.e. an odd number of parameters, which means no "key:" was specified.
							return LineError(_T("Missing \"propertyname:\" in object literal."));
					}
				}
			}
				
			switch (infix_symbol)
			{
			case SYM_CPAREN: // implies stack_symbol == SYM_OPAREN.
				// See comments near the bottom of this (outer) case.  The first open-paren on the stack must be the one that goes with this close-paren.
				--stack_count; // Remove this open-paren from the stack, since it is now complete.
				++this_infix;  // Since this pair of parentheses is done, move on to the next token in the infix expression.

				in_param_list = stack[stack_count]->outer_param_list; // Restore in_param_list to the value it had when SYM_OPAREN was pushed onto the stack.

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
				++this_infix; // Since this pair of brackets is done, move on to the next token in the infix expression.
				in_param_list = stack_top.outer_param_list; // Restore in_param_list to the value it had when '[' was pushed onto the stack.					
				goto standard_pop_into_postfix; // Pop the token (now SYM_FUNC) into the postfix array to immediately follow its params.
			}

			default: // case SYM_COMMA:
				if (sPrecedence[stack_symbol] < sPrecedence[infix_symbol]) // i.e. stack_symbol is SYM_BEGIN or SYM_OPAREN/BRACKET/BRACE.
				{
					if (!in_param_list) // This comma separates statements rather than function parameters.
					{
						STACK_PUSH(this_infix++);
						// Pop this comma immediately into postfix so that when it is encountered at
						// run-time, it will pop and discard the result of its left-hand sub-statement.
						goto standard_pop_into_postfix;
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
			this_infix->outer_param_list = in_param_list; // Save current value on the stack with this SYM_OPAREN.
			if (infix_symbol == SYM_FUNC)
				in_param_list = this_infix[-1].callsite; // Store this SYM_FUNC's deref.
			else if (this_infix > infix && YIELDS_AN_OPERAND(this_infix[-1].symbol)
				&& *this_infix->marker == '(') // i.e. it's not an implicit SYM_OPAREN generated by DT_STRING.
			{
				// This is a function call with something other than a known function name,
				// so this_infix[-1] is the object to be called (or an error).
				auto token = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
				token->symbol = SYM_FUNC;
				token->callsite = in_param_list = new CallSite();
				if (!in_param_list)
					return LineError(ERR_OUTOFMEM);
				STACK_PUSH(token);
			}
			else
				in_param_list = NULL; // Allow multi-statement commas, even in cases like Func((x,y)).
			STACK_PUSH(this_infix++);
			break;
			
		case SYM_OBRACKET:
		case SYM_OBRACE:
			this_infix->outer_param_list = in_param_list; // Save current value on the stack with this SYM_OBRACKET.
			in_param_list = this_infix->callsite;
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
			{
				// This should be the end of a property name in something like {x: y}.
				if (postfix_count)
				{
					LPTSTR cp;
					switch (postfix[postfix_count - 1]->symbol)
					{
					case SYM_DYNAMIC:
						// This is the end of a dynamic property name in something like {%x%: y}.
						postfix_count--;
						break;
					case SYM_STRING:
						// Is this a quoted string (invalid here) or unquoted property name?
						for (cp = this_infix->marker - 1; IS_SPACE_OR_TAB(*cp); --cp);
						if (IS_IDENTIFIER_CHAR(*cp))
							break;
					default:
						return LineError(_T("Invalid property name in object literal."));
					}
				}
				++in_param_list->param_count;
				++this_infix;
				continue;
			}
			if (stack_symbol == SYM_OPAREN)
				return LineError(_T("Unexpected \":\"")); // No reference to ")" since it might be a function call statement.
			if (stack_symbol == SYM_OBRACKET)
				return LineError(_T("Missing \"]\" before \":\""));
			if (stack_symbol == SYM_OBRACE)
				return LineError(_T("Missing \"}\" before \":\""));
			// Otherwise:
			if (stack_symbol == SYM_IFF_THEN) // See comments near the bottom of this case. The first found "THEN" on the stack must be the one that goes with this "ELSE".
			{
				// An earlier stage put THEN directly into the postfix array before also pushing it onto the stack.
				// Pop this THEN off the stack and point it at its ELSE, but don't write it into postfix again:
				STACK_POP->circuit_token = this_infix;
				// Since this ELSE also marks the end of the THEN branch, write it to postfix (in place of the above)
				// so that after the THEN branch is evaluated it will hit this ELSE and skip over the ELSE branch:
				this_postfix = this_infix;
				++postfix_count;
				// Also push this ELSE onto the stack. Once the end of branch is found, ELSE will be popped off the
				// stack, updated to point to the end of the ELSE branch and not written to postfix a second time:
				STACK_PUSH(this_infix++);
				// Above also increments this_infix: this ELSE found its matching IF/THEN, so it's time to move on to the next token in the infix expression.
			}
			else // This stack item is an operator, possibly even some other THEN's ELSE (all such ELSE's should be purged from the stack to properly support nested ternary).
			{
				// By not incrementing this_infix, the loop will continue to encounter SYM_IFF_ELSE and thus
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
			else if (  stack_symbol == SYM_OPAREN // Open paren is never closed (currently impossible due to load-time balancing, but kept for completeness).
					|| stack_symbol == SYM_OBRACKET
					|| stack_symbol == SYM_OBRACE  )
				return LineError(ERR_EXPR_SYNTAX);
			else // Pop item off the stack, AND CONTINUE ITERATING, which will hit this line until stack is empty.
				goto standard_pop_into_postfix;
			// ALL PATHS ABOVE must continue or goto.
			
		case SYM_MULTIPLY:
			if (in_param_list && (this_infix[1].symbol == SYM_CPAREN || this_infix[1].symbol == SYM_CBRACKET)) // Func(params*) or obj.foo[params*]
			{
				in_param_list->is_variadic(true);
				++this_infix;
				continue;
			}
			// DO NOT BREAK: FALL THROUGH TO BELOW

		default: // This infix symbol is an operator, so act according to its precedence.
			// If the symbol waiting on the stack has a lower precedence than the current symbol, push the
			// current symbol onto the stack so that it will be processed sooner than the waiting one.
			// Otherwise, pop waiting items off the stack (by means of i not being incremented) until their
			// precedence falls below the current item's precedence, or the stack is emptied.
			// Note: BEGIN and OPAREN are the lowest precedence items ever to appear on the stack (CPAREN
			// never goes on the stack, so can't be encountered there).
			if (   sPrecedence[stack_symbol] < sPrecedence[infix_symbol] + (sPrecedence[infix_symbol] % 2) // Performance: An sPrecedence2[] array could be made in lieu of the extra add+indexing+modulo, but it benched only 0.3% faster, so the extra code size it caused didn't seem worth it.
				|| IS_ASSIGNMENT_EXCEPT_POST_AND_PRE(infix_symbol) // See note 1 below. Ordered for short-circuit performance.
				|| stack_symbol == SYM_POWER && SYM_OVERRIDES_POWER_ON_STACK(infix_symbol)   ) // See note 2 below.
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

				// Perform some rough checks to detect most syntax errors.  This is done after the
				// precedence check so that it isn't done multiple times for a single token when
				// the stack contains one or more higher-precedence operators, and also so that
				// the left operand (if this is a binary operator) has been popped into postfix.
				sym_prev = this_infix > infix ? this_infix[-1].symbol : SYM_INVALID;
				sym_next = this_infix[1].symbol; // It will be SYM_INVALID if there are no more.
				SymbolType sym_postfix = postfix_count ? postfix[postfix_count-1]->symbol : SYM_INVALID;
				if (IS_ASSIGNMENT_OR_POST_OP(infix_symbol))
				{
					// Assignment and postfix operators must be preceded by a variable, except for
					// assignment operators which have been marked with a non-NULL deref, indicating
					// that the target is an object's property.  Postfix operators which apply to an
					// object's property are fully handled in the standard_pop_into_postfix section.
					if (this_infix->callsite) // Object property.  Takes precedence over the next checks.
					{}  // Nothing needed here.
					else if (sym_postfix == SYM_VAR || sym_postfix == SYM_DYNAMIC)
					{
						ExprTokenType &target = *postfix[postfix_count - 1];
						target.var_usage = VARREF_LVALUE; // Mark this as the target of an assignment.
					}
					else if (!IS_OPERATOR_VALID_LVALUE(sym_postfix))
						return LineError(ERR_INVALID_ASSIGNMENT, FAIL, this_infix->error_reporting_marker);
				}
				else
				{
					// Prefix operators must not be preceded by an operand.
					// Binary operators must be preceded by an operand.
					// If sym_prev == SYM_FUNC, that can only be the result of SYM_DOT (x.y).
					if ((YIELDS_AN_OPERAND(sym_prev) || sym_prev == SYM_FUNC) == IS_PREFIX_OPERATOR(infix_symbol))
						return LineError(ERR_EXPR_SYNTAX, FAIL, this_infix->error_reporting_marker);
					// If it's pre-increment/decrement, looking to the right in infix is insufficient.
					// For cases like ++this.x, we must wait until the operator is popped from the stack.
				}
				if (!IS_POSTFIX_OPERATOR(infix_symbol))
				{
					// Prefix and binary operators must be followed by an operand.
					if (  !(IS_OPERAND(sym_next)
						 || sym_next == SYM_FUNC
						 || IS_OPAREN_LIKE(sym_next)
						 || IS_PREFIX_OPERATOR(sym_next))  )
						return LineError(ERR_EXPR_MISSING_OPERAND, FAIL, this_infix->error_reporting_marker);
				}

				if (IS_SHORT_CIRCUIT_OPERATOR(infix_symbol))
				{
					// Short-circuit boolean evaluation works as follows:
					//
					// When an AND/OR/IFF is encountered in infix, it is immediately put into postfix to act as a
					// unary operator on its left operand and something like a conditional jump instruction. It is
					// also pushed onto the stack so that the end of its right branch can be found in accordance with
					// precedence rules; when it is popped off the stack, circuit_token is set to the last token in
					// postfix at that point, which is necessarily the end of the operator's right branch.
					//
					// x AND y
					// x OR y
					// x IFF_THEN y IFF_ELSE z
					//
					// AND/OR: Pop operand off stack. If short-circuit condition is satified, push the operand back
					// onto the stack and jump over the operator's right branch; otherwise simply continue, allowing
					// the right branch to be evaluated and its result used as the overall result of the AND/OR.
					//
					// IFF_THEN: Pop operand off stack. If false, jump to the "else" branch; otherwise continue.
					// IFF_ELSE: Encountered only after evaluating the "then" branch; jumps over the "else" branch.
					//
					// This new approach has the following benefits over the old approach:
					// 1) circuit_token doesn't need to be checked every time a value is pushed onto the stack.
					// 2) circuit_token doesn't need to propagate from an operator or function-call to its result.
					// 3) Cascading of short-circuit operators does not need to be handled specifically
					//    as it is naturally supported via infix-to-postfix conversion.
					// 4) Since circuit_token is only needed in the actual operators, it can overlap with
					//    the value fields -- i.e. the size of ExprTokenType may be reduced (assuming
					//    the other uses of circuit_token can be eliminated).
					//
					// OBSOLETE COMMENTS:
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
					//
					this_postfix = this_infix;
					++postfix_count;
				}
				STACK_PUSH(this_infix++); // Push this_infix onto the stack and move rightward to the next infix item.
			}
			else // Stack item's precedence >= infix's (if equal, left-to-right evaluation order is in effect).
				goto standard_pop_into_postfix;
		} // switch(infix_symbol)

		continue; // Avoid falling into the label below except via explicit jump.  Performance: Doing it this way rather than replacing break with continue everywhere above generates slightly smaller and slightly faster code.
standard_pop_into_postfix: // Use of a goto slightly reduces code size.
		this_postfix = STACK_POP;
		// Additional processing for short-circuit evaluation and syntax sugar:
		SymbolType postfix_symbol = this_postfix->symbol;
		switch (postfix_symbol)
		{
		case SYM_LOW_CONCAT:
			// Now that operator precedence has been fully handled, change the symbol to simplify runtime evaluation.
			this_postfix->symbol = SYM_CONCAT;
			break;

		case SYM_FUNC:
			infix_symbol = this_infix->symbol;
			// The sections below pre-process assignments to work with objects:
			//	x.y := z	->	x "y" z (set)
			//	x[y] += z	->	x y (get in-place, assume 2 params) z (add) (set)
			//	x.y[i] /= z	->	x "y" i 3 (get in-place, n params) z (div) (set)
			if ((this_postfix->callsite->flags & IT_BITMASK) == IT_GET)
			{
				bool square_brackets = !this_postfix->callsite->member;
				
				if (IS_ASSIGNMENT_EXCEPT_POST_AND_PRE(infix_symbol))
				{
					auto callsite = this_postfix->callsite; // To be reused below.
					if (infix_symbol != SYM_ASSIGN)
					{
						// Give it a new callsite since the original one will be reused below
						// (so that there's no redundant allocation for the SYM_ASSIGN case).
						this_postfix->callsite = new CallSite();
						this_postfix->callsite->flags		= callsite->flags | EIF_LEAVE_PARAMS;
						this_postfix->callsite->member		= callsite->member;
						this_postfix->callsite->param_count	= callsite->param_count;
					}
					else
					{
						--postfix_count; // Discard this token; the assignment op will be converted into SYM_FUNC later.
					}
					callsite->flags = IT_SET | (callsite->flags & ~IT_GET);
					callsite->param_count++; // Include the value to be assigned.
					this_infix->callsite = callsite;
					// Now let this_infix be processed by the next iteration, and eventually
					// have its symbol changed to SYM_FUNC.
				}
				else if (!IS_OPERAND(infix_symbol))
				{
					stack_symbol = stack[stack_count - 1]->symbol;
					// Post-increment/decrement has higher precedence, so check for it first:
					bool is_post_op = (infix_symbol == SYM_POST_INCREMENT || infix_symbol == SYM_POST_DECREMENT);
					if (   is_post_op
						|| infix_symbol != SYM_DOT && infix_symbol != SYM_OPAREN // Do not apply the "++" to the "x.y" part of "++x.y.z" (SYM_DOT) or "++x.y.%z%" (SYM_OPAREN).
						&& (stack_symbol == SYM_PRE_INCREMENT || stack_symbol == SYM_PRE_DECREMENT)   )
					{
						auto get_token = this_postfix;
						auto get_callsite = get_token->callsite;

						auto set_callsite = new CallSite();
						set_callsite->member      = get_callsite->member;
						set_callsite->flags       = IT_SET | (get_callsite->flags & ~IT_GET);
						set_callsite->param_count = get_callsite->param_count + 1;
						
						get_callsite->flags |= EIF_LEAVE_PARAMS;

						bool is_increment = is_post_op ? infix_symbol == SYM_POST_INCREMENT
							: stack_symbol == SYM_PRE_INCREMENT;
						// Drop the pre/post operator since it's fully handled here, but reuse the token:
						auto n_token = is_post_op ? this_infix++ : stack[--stack_count];
						n_token->SetValue(is_increment ? 1 : -1);

						postfix[postfix_count + 1] = n_token;
						postfix[postfix_count + 2] = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
						postfix[postfix_count + 2]->symbol = SYM_ADD;
						postfix[postfix_count + 3] = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
						postfix[postfix_count + 3]->symbol = SYM_FUNC;
						postfix[postfix_count + 3]->callsite = set_callsite;
						if (is_post_op)
						{
							// Subtract from the result to get back to the original value.
							postfix[postfix_count + 4] = n_token;
							postfix[postfix_count + 5] = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
							postfix[postfix_count + 5]->symbol = SYM_SUBTRACT;
							postfix_count += 5; // Account for the extra tokens; a final ++ is done below.
						}
						else
							postfix_count += 3;
					}
				}
				// Otherwise, IS_OPERAND(infix_symbol) == true, which should only be possible
				// if this_infix[1] is SYM_DOT.  In that case, a later iteration should apply
				// the transformations above to that operator.
			}
			else if ((this_postfix->callsite->flags & IT_BITMASK) == IT_CALL)
			{
				if (this_postfix->callsite->param_count == 1 && this_postfix->callsite->func == sIsSetFunc)
				{
					if (postfix[postfix_count - 1]->symbol != SYM_VAR && postfix[postfix_count - 1]->symbol != SYM_DYNAMIC)
					{
						return LineError(_T("IsSet requires a variable."), FAIL, this_postfix->error_reporting_marker);
					}
					postfix[postfix_count - 1]->var_usage = VARREF_ISSET;
				}
			}
			break;

		case SYM_REGEXMATCH: // a ~= b  ->  RegExMatch(a, b)
			this_postfix->symbol = SYM_FUNC;
			break;

		case SYM_AND:
		case SYM_OR:
		case SYM_OR_MAYBE:
		case SYM_IFF_ELSE:
		{
			// Point this short-circuit operator to the end of its right operand.
			ExprTokenType *iff_end = postfix[postfix_count - 1];
			if (this_postfix == iff_end) // i.e. the last token is the operator itself.
				return LineError(ERR_EXPR_MISSING_OPERAND, FAIL, this_postfix->marker);
			// Point the original token already in postfix to the end of its right branch:
			this_postfix->circuit_token = iff_end;
			continue; // This token was already put into postfix by an earlier stage, so skip it this time.
		}
		case SYM_IFF_THEN:
			return LineError(_T("A \"?\" is missing its \":\""), FAIL, this_postfix->marker);

		case SYM_REF:
			postfix_symbol = postfix[postfix_count - 1]->symbol;
			if (postfix_symbol == SYM_VAR || postfix_symbol == SYM_DYNAMIC)
				postfix[postfix_count - 1]->var_usage = VARREF_REF;
			else if (!IS_OPERATOR_VALID_LVALUE(postfix_symbol))
				return LineError(_T("\"&\" requires a variable."));
			break;

		case SYM_PRE_INCREMENT:
		case SYM_PRE_DECREMENT:
			if (postfix_count)
			{
				ExprTokenType &target = *postfix[postfix_count - 1];
				// This is nearly identical to the section for assignments under "if (IS_ASSIGNMENT_OR_POST_OP(infix_symbol))":
				if (target.symbol == SYM_VAR || target.symbol == SYM_DYNAMIC)
					target.var_usage = VARREF_LVALUE; // Mark this as the target of an assignment.
				else if (!IS_OPERATOR_VALID_LVALUE(target.symbol))
					return LineError(ERR_INVALID_ASSIGNMENT, FAIL, this_postfix->error_reporting_marker);
			}
			break;

		default:
			if (!IS_ASSIGNMENT_EXCEPT_POST_AND_PRE(postfix_symbol))
				break;
			if (this_postfix->callsite)
			{
				// An earlier iteration of the SYM_FUNC section above used callsite to mark this as an object assignment.
				ExprTokenType *assign_op = this_postfix;
				if (postfix_symbol != SYM_ASSIGN) // e.g. += or .=
				{
					switch (postfix_symbol)
					{
					case SYM_ASSIGN_ADD:					postfix_symbol = SYM_ADD; break;
					case SYM_ASSIGN_SUBTRACT:				postfix_symbol = SYM_SUBTRACT; break;
					case SYM_ASSIGN_MULTIPLY:				postfix_symbol = SYM_MULTIPLY; break;
					case SYM_ASSIGN_DIVIDE:					postfix_symbol = SYM_DIVIDE; break;
					case SYM_ASSIGN_INTEGERDIVIDE:			postfix_symbol = SYM_INTEGERDIVIDE; break;
					case SYM_ASSIGN_BITOR:					postfix_symbol = SYM_BITOR; break;
					case SYM_ASSIGN_BITXOR:					postfix_symbol = SYM_BITXOR; break;
					case SYM_ASSIGN_BITAND:					postfix_symbol = SYM_BITAND; break;
					case SYM_ASSIGN_BITSHIFTLEFT:			postfix_symbol = SYM_BITSHIFTLEFT; break;
					case SYM_ASSIGN_BITSHIFTRIGHT:			postfix_symbol = SYM_BITSHIFTRIGHT; break;
					case SYM_ASSIGN_BITSHIFTRIGHT_LOGICAL:	postfix_symbol = SYM_BITSHIFTRIGHT_LOGICAL; break;
					case SYM_ASSIGN_CONCAT:					postfix_symbol = SYM_CONCAT; break;
					}
					// Insert the concat or math operator before the assignment:
					this_postfix = (ExprTokenType *)_alloca(sizeof(ExprTokenType));
					this_postfix->symbol = postfix_symbol;
					postfix[++postfix_count] = assign_op;
				}
				assign_op->symbol = SYM_FUNC; // An earlier stage already set up assign_op->callsite.
			}
		}
		++postfix_count;
	} // End of loop that builds postfix array from the infix array.
end_of_infix_to_postfix:

	if (!postfix_count) // The code below relies on this check.  This can't be an empty (omitted) expression because an earlier check would've turned it into a non-expression.
		return LineError(ERR_EXPR_SYNTAX, FAIL, mArgc > 1 ? aArg.text : _T(""));

	// The following enables ExpandExpression() to be skipped in common cases for ACT_ASSIGNEXPR
	// and ACT_RETURN.  A similar optimization used to be done for simple literal integers by
	// storing an __int64* in arg.postfix, but this new approach allows floating-point numbers
	// and literal strings to be optimized, and the pure numeric status of floats to be retained.
	// Unlike the old method, numeric literals are reformatted to look exactly like they would if
	// this remained an expression; for example, MsgBox 1.00 shows "1.0" instead of "1.00".
	ExprTokenType &only_token = *postfix[0];
	SymbolType only_symbol = only_token.symbol;
	if (   postfix_count == 1 && IS_OPERAND(only_symbol) // This expression is a lone operand, like (1) or "string".
		&& (mActionType < ACT_FOR || mActionType > ACT_UNTIL) // It's not WHILE or UNTIL, which currently perform better as expressions, or FOR, which performs the same but currently expects aResultToken to always be set.
		&& (mActionType != ACT_SWITCH && mActionType != ACT_CASE) // It's not SWITCH or CASE, which require a proper postfix expression.
		&& (mActionType != ACT_THROW) // Exclude THROW to simplify variable handling (ensures vars are always dereferenced).
		&& (mActionType != ACT_HOTKEY_IF) // #HotIf requires the expression text not be modified.
		&& (only_symbol != SYM_VAR || mActionType != ACT_RETURN) // "return var" is kept as an expression for correct handling of local vars (see "ToReturnValue") and ByRef.
		)
	{
		// The checks above leave: ACT_ASSIGNEXPR, ACT_EXPRESSION? (ineffectual), ACT_IF,
		// ACT_LOOP and co., ACT_GOTO, ACT_RETURN (if not var).
		switch (only_symbol)
		{
		case SYM_INTEGER:
		case SYM_FLOAT:
			// Convert this numeric literal back into a string to ensure the format is consistent.
			// This also ensures parentheses are not present in the output, for cases like MsgBox (1.0).
			aArg.text = SimpleHeap::Alloc(TokenToString(only_token, number_buf));
			aArg.is_expression = false;
			break;
		case SYM_STRING:
			// If this arg will be expanded normally, it needs a pointer to the final string,
			// without leading/trailing quotation marks.  This doesn't apply to ACT_ASSIGNEXPR or
			// ACT_RETURN, since they use the string token in postfix.  Line::ToText compensates
			// for this by adding quotes.
			aArg.text = only_token.marker;
			aArg.is_expression = false;
			break;
		case SYM_VAR:
			// Since is_expression is still true, this is only a hint for a later stage:
			if (aArg.type != ARG_TYPE_OUTPUT_VAR)
				aArg.type = ARG_TYPE_INPUT_VAR;
			break;
		}
	}

	// Create a new postfix array and attach it to this arg of this line.
	aArg.postfix = SimpleHeap::Alloc<ExprTokenType>(postfix_count + 1); // +1 for the terminator item added below.

	int i, j, max_stack = 0, max_alloc = 0;
	for (i = 0; i < postfix_count; ++i) // Copy the postfix array in physically sorted order into the new postfix array.
	{
		ExprTokenType &new_token = aArg.postfix[i];
		new_token.CopyExprFrom(*postfix[i]);
		ASSERT((UINT)new_token.symbol < SYM_COUNT);
		if (SYM_USES_CIRCUIT_TOKEN(new_token.symbol)) // Adjust each circuit_token address to be relative to the new array rather than the temp/infix array.
		{
			// circuit_token should always be non-NULL at this point.
			for (j = i + 1; postfix[j] != new_token.circuit_token; ++j); // Should always be found, and always to the right in the postfix array, so no need to check postfix_count.
			new_token.circuit_token = aArg.postfix + j;
		}
		// Simple calculation: only operands and SYM_FUNC can increase the stack count,
		// so this finds the worst-case stack requirement (or slightly higher).
		if (IS_OPERAND(new_token.symbol) || new_token.symbol == SYM_FUNC)
			++max_stack;
		// Resolve variable references for assignments and output vars.  Leave all others
		// for later to facilitate detecting variables which are not assigned anywhere.
		if (new_token.symbol == SYM_VAR && VARREF_IS_WRITE(new_token.var_usage))
		{
			new_token.var = g_script.FindOrAddVar(new_token.var_deref->marker, new_token.var_deref->length, FINDVAR_FOR_WRITE);
			if (!new_token.var)
				return FAIL;
			// This is currently left to FinalizeExpression() to reduce code size:
			//if (new_token.var_usage != VARREF_REF // Validation of REF is left to FinalizeExpression().
			//	&& !ValidateVarUsage(new_token.var, new_token.var_usage))
			//	return FAIL;
			new_token.var->MarkAssignedSomewhere();
			if (aArg.type == ARG_TYPE_OUTPUT_VAR)
			{
				aArg.deref = (DerefType *)new_token.var;
				aArg.is_expression = false;
			}
		}
		// Count the tokens which potentially use to_free[].
		if (new_token.symbol == SYM_DYNAMIC || new_token.symbol == SYM_FUNC
			|| new_token.symbol == SYM_CONCAT || new_token.symbol == SYM_REF)
			++max_alloc;
	}
	aArg.postfix[postfix_count].symbol = SYM_INVALID;  // Special item to mark the end of the array.
	aArg.max_stack = max_stack;
	aArg.max_alloc = max_alloc;

	return OK;
}



ResultType Line::FinalizeExpression(ArgStruct &aArg)
{
	auto stack = (ExprTokenType **)_alloca(aArg.max_stack * sizeof(ExprTokenType **));
	int stack_count = 0;
	// stack_count checks: Since missing operands are caught at an earlier stage, it doesn't
	// seem possible to produce any situation where the stack is unbalanced, but due to the
	// overall complexity it seems best to assume that it could happen.  Having these checks
	// here might help catch future (or present) bugs.

#define TOKEN_MAY_MISS(token) ((token)->symbol == SYM_MISSING || (token)->symbol == SYM_VAR && (token)->var_usage == VARREF_READ_MAYBE)

	for (auto this_postfix = aArg.postfix; this_postfix->symbol != SYM_INVALID; ++this_postfix)
	{
		auto postfix_symbol = this_postfix->symbol;
		ASSERT((UINT)postfix_symbol < SYM_COUNT);
		if (IS_OPERAND(postfix_symbol))
		{
			if (postfix_symbol == SYM_DYNAMIC)
			{
				if (stack_count < 1 || TOKEN_MAY_MISS(stack[stack_count - 1]))
					return LineError(ERR_EXPR_SYNTAX);
				--stack_count;
			}
			stack[stack_count++] = this_postfix;
		}
		else if (IS_POSTFIX_OPERATOR(postfix_symbol) || IS_PREFIX_OPERATOR(postfix_symbol))
		{
			if (stack_count < 1 || TOKEN_MAY_MISS(stack[stack_count - 1]))
				return LineError(ERR_EXPR_SYNTAX);
			stack[stack_count - 1] = this_postfix;
		}
		else if (SYM_USES_CIRCUIT_TOKEN(postfix_symbol))
		{
			if (stack_count < 1 || TOKEN_MAY_MISS(stack[stack_count - 1]) && postfix_symbol != SYM_IFF_ELSE && postfix_symbol != SYM_OR_MAYBE)
				return LineError(ERR_EXPR_SYNTAX);
			// Pop the result of the left branch of this short-circuit operator, then just allow the right branch
			// to be evaluated and its value/token used as the result.  A consequence of this simple approach is
			// that when a short-circuit expression is passed as a parameter, only the right branch...
			//  1) is permitted to be something like &A_Index, being passed to a built-in output var param.
			//  2) is resolved to a function address in something like DllCall(a || "B") or DllCall(a?"B":"C").
			// If this is changed, be sure that both branches are "finalized"; i.e. don't just skip over the right
			// branch, since it might contain invalid function calls.  This seems difficult to do in one pass since
			// the end of the right branch isn't marked in any way, but it can be done by having a recursive call
			// finalize from [this_postfix] to [this_postfix->circuit_token], then having this layer skip it.
			--stack_count;
		}
		else if (postfix_symbol == SYM_COMMA)
		{
			if (stack_count < 1 || TOKEN_MAY_MISS(stack[stack_count - 1]))
				return LineError(ERR_EXPR_SYNTAX);
			--stack_count;
		}
		else if (postfix_symbol != SYM_FUNC)
		{
			if (stack_count < 2
				|| TOKEN_MAY_MISS(stack[stack_count - 1]) && postfix_symbol != SYM_ASSIGN
				|| TOKEN_MAY_MISS(stack[stack_count - 2]))
				return LineError(ERR_EXPR_SYNTAX);
			--stack_count; // Pop RHS
			stack[stack_count - 1] = this_postfix; // Replace LHS
		}
		else // SYM_FUNC
		{
			int prev_stack_count = stack_count;
			int param_count = this_postfix->callsite->param_count;
			stack_count -= param_count;
			auto param = stack + stack_count;
			bool call_call = !this_postfix->callsite->member && IT_CALL == (this_postfix->callsite->flags & IT_BITMASK);
			if (this_postfix->callsite->flags & EIF_STACK_MEMBER)
			{
				--stack_count;
				call_call = false;
			}
			if (stack_count < 0)
				return LineError(ERR_EXPR_SYNTAX);
			Func *func = nullptr;
			if (this_postfix->callsite->func)
				func = this_postfix->callsite->func;
			else
			{
				if (stack_count < 1)
					return LineError(ERR_EXPR_SYNTAX);
				auto func_op = stack[--stack_count];
				func = dynamic_cast<Func*>(TokenToObject(*func_op));
			}
			if (this_postfix->callsite->flags & EIF_LEAVE_PARAMS)
				stack_count = prev_stack_count;
			if (func && call_call)
			{
				if (this_postfix->callsite->is_variadic())
					--param_count; // Exclude the array parameter, which resolves to 0 or more parameters.
				if (param_count < func->mMinParams && !this_postfix->callsite->is_variadic())
					return LineError(ERR_TOO_FEW_PARAMS, FAIL, func->mName);
				if (param_count > func->mParamCount && !func->mIsVariadic)
					return LineError(ERR_TOO_MANY_PARAMS, FAIL, func->mName);

				auto func_as_bif = dynamic_cast<BuiltInFunc*>(func);
				auto *bif = func_as_bif ? func_as_bif->mBIF : nullptr;

				for (int i = 0; i < param_count; ++i)
				{
					switch (param[i]->symbol)
					{
					case SYM_MISSING:
						if (i < func->mMinParams)
							return LineError(ERR_PARAM_REQUIRED);
						break;
					case SYM_REF:
						if (func->ArgIsOutputVar(i))
						{
							// Override the SYM_REF's default var_usage of VARREF_READ to indicate 1) that SYM_VAR can be
							// passed directly without the need to allocate a VarRef, and 2) which var types are allowed.
							// Permit VAR_VIRTUAL for all built-in functions except VarSetStrCapacity.
							param[i]->var_usage =
								(!bif || bif == BIF_VarSetStrCapacity) ? VARREF_REF
								:										 VARREF_OUTPUT_VAR;
							if (!bif)
								param[i]->object = func; // Used during runtime to detect when a VarRef is needed due to recursion (func->mInstances).
							if (param[i][-1].symbol == SYM_DYNAMIC || param[i][-1].symbol == SYM_VAR)
								param[i][-1].var_usage = param[i]->var_usage;
						}
						break;
					}
				}

				if (bif == &BIF_DllCall && param_count)
				{
					if (param[0]->symbol == SYM_STRING)
					{
						if (void *function = GetDllProcAddress(param[0]->marker))
							param[0]->SetValue((__int64)function);
					}
				}
			}
			stack[stack_count++] = this_postfix;
		} // SYM_FUNC
	} // postfix loop

	if (stack_count != 1 || TOKEN_MAY_MISS(stack[stack_count - 1]) && mActionType != ACT_ASSIGNEXPR)
		return LineError(ERR_EXPR_SYNTAX);

#undef TOKEN_MAY_MISS

	for (auto this_postfix = aArg.postfix; this_postfix->symbol != SYM_INVALID; ++this_postfix)
	{
		if (this_postfix->symbol != SYM_VAR)
			continue;
		// Now that var_usage has been set for any SYM_REF operators in postfix that are passed
		// to functions, ensure the var being referenced is valid for the given var_usage.
		if (!ValidateVarUsage(this_postfix->var, this_postfix->var_usage))
			return FAIL;
		if (!this_postfix->var->IsAssignedSomewhere()
			&& this_postfix->var_usage == VARREF_READ) // i.e. VARREF_READ_MAYBE (var?, var ?? value) generates no warning but doesn't mark var as assigned.
			g_script.WarnUnassignedVar(this_postfix->var, this);
	}

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



ResultType Line::ExecUntil(ExecUntilMode aMode, ResultToken *aResultToken, Line **apJumpToLine)
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
	LoopFilesStruct *loop_file;
	RegItemStruct *loop_reg_item;
	LoopReadFileStruct *loop_read_file;
	LPTSTR loop_field;

	Line *jump_to_line; // Don't use *apJumpToLine because it might not exist.
	Label *jump_to_label;  // For use with Goto.
	ResultType if_condition, result;
	LONG_OPERATION_INIT
	global_struct &g = *::g; // Reduces code size and may improve performance. Eclipsing ::g with local g makes compiler remind/enforce the use of the right one.

	for (Line *line = this; line != NULL;)
	{
		// The below must be done at least when the keybd or mouse hook is active, but is currently
		// always done since it's a very low overhead call, and has the side-benefit of making
		// the app maximally responsive when the script is busy.
		// This low-overhead call achieves at least two purposes optimally:
		// 1) Keyboard and mouse lag is minimized when the hook(s) are installed, since this single
		//    Peek() is apparently enough to route all pending input to the hooks (though it's inexplicable
		//    why calling MsgSleep(-1) does not achieve this goal, since it too does a Peek().
		//    Nevertheless, that is the testing result that was obtained: the mouse cursor lagged
		//    in tight script loops even when MsgSleep(-1) or (0) was called every 10ms or so.
		// 2) The app is maximally responsive while executing in a tight loop.
		// 3) Hotkeys are maximally responsive.  For example, if a user has game hotkeys, using
		//    a GetTickCount() method (which very slightly improves performance by cutting back on
		//    the number of Peek() calls) would introduce up to 10ms of delay before the hotkey
		//    finally takes effect.  10ms can be significant in games, where ping (latency) itself
		//    can sometimes be only 10 or 20ms. UPDATE: It looks like PeekMessage() yields CPU time
		//    automatically, similar to a Sleep(0), when our queue has no messages.  Since this would
		//    make scripts slow to a crawl, only do the Peek() every 5ms or so (though the timer
		//    granularity is 10ms on most OSes, so that's the true interval).
		// 4) Timed subroutines are run as consistently as possible (to help with this, a check
		//    similar to the below is also done for single commmands that take a long time, such
		//    as Download, FileSetAttrib, etc.
		LONG_OPERATION_UPDATE

		// If interruptions are currently forbidden, it's our responsibility to check if the number
		// of lines that have been run since this quasi-thread started now indicate that
		// interruptibility should be reenabled.  g.UninterruptedLineCount must also be tracked
		// for IsInterruptible() to detect execution of the thread's first line.
		// v1.0.38.04: If g.ThreadIsCritical==true, no need to check or accumulate g.UninterruptedLineCount
		// because the script is now in charge of this thread's interruptibility.
		if (!g.AllowThreadToBeInterrupted && !g.ThreadIsCritical) // Ordered for short-circuit performance.
		{
			// Incrementing this unconditionally makes it a relatively crude measure,
			// but it seems okay to be less accurate for this purpose:
			++g.UninterruptedLineCount;
			// To preserve backward compatibility, ExecUntil() must be the one to check
			// g.UninterruptedLineCount and update g.AllowThreadToBeInterrupted, rather than doing
			// those things on-demand in IsInterruptible().  If those checks were moved to
			// IsInterruptible(), they might compare against a different/changed value of
			// g_script.mUninterruptedLineCountMax because IsInterruptible() is called only upon demand.
			// THIS SECTION DOES NOT CHECK g.ThreadStartTime because that only needs to be
			// checked on demand by callers of IsInterruptible().
			if (g.UninterruptedLineCount > g_script.mUninterruptedLineCountMax // See above.
				&& g_script.mUninterruptedLineCountMax > -1)
				g.AllowThreadToBeInterrupted = true;
		}

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
		// line (e.g. control stmts such as IF and LOOP).
		if (!ACT_EXPANDS_ITS_OWN_ARGS(line->mActionType)) // Not ACT_ASSIGNEXPR, ACT_WHILE or ACT_THROW.
		{
			result = line->ExpandArgs(line->mActionType == ACT_RETURN ? aResultToken : NULL);
			// As of v1.0.31, ExpandArgs() will also return EARLY_EXIT if a function call inside one of this
			// line's expressions did an EXIT.
			if (result != OK)
				return result; // In the case of FAIL: Abort the current subroutine, but don't terminate the app.
		}

		switch (line->mActionType)
		{
		case ACT_IF:
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
					// or if this if's only action is RETURN.
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

		case ACT_GOTO:
			// A single goto can cause an infinite loop if misused, so be sure to do this to
			// prevent the program from hanging:
			if (   !(jump_to_label = (Label *)line->mRelatedLine)   )
				// The label is a dereference, otherwise it would have been resolved at load-time.
				// So send true because we don't want to update its mRelatedLine.  This is because
				// we want to resolve the label every time through the loop in case the variable
				// that contains the label changes, e.g. Goto(VarContainingLabelName)
				if (   !(jump_to_label = line->GetJumpTarget(true))   )
					return FAIL; // Error was already displayed by the called function.
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

		case ACT_RETURN:
			// Although a return is really just a kind of block-end, keep it separate
			// because when a return is encountered inside a block, it has a double function:
			// to first break out of all enclosing blocks and then return from the function.
			// NOTE: The return's ARG1 expression has been evaluated by ExpandArgs() above,
			// which is desirable *even* if aResultToken is NULL (i.e. the caller will be
			// ignoring the return value) in case the return's expression calls a function
			// which has side-effects.  For example, "return LogThisEvent()".
			if (aMode != UNTIL_RETURN)
				// Tells the caller to return early if it's not the function that directly
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
		case ACT_LOOP_FILE:
		case ACT_LOOP_REG:
		case ACT_LOOP_READ:
		case ACT_LOOP_PARSE:
		case ACT_FOR:
		case ACT_WHILE:
		{
			// HANDLE ANY ERROR CONDITIONS THAT CAN ABORT THE LOOP:
			FileLoopModeType file_loop_mode;
			HKEY root_key_type; // For registry loops, this holds the type of root key, independent of whether it is local or remote.
			if (line->mActionType == ACT_LOOP_FILE || line->mActionType == ACT_LOOP_REG)
			{
				if (line->mActionType == ACT_LOOP_REG)
					if (  !(root_key_type = RegConvertKey(ARG1))  )
						return line->LineError(ERR_PARAM1_INVALID, FAIL, ARG1);

				file_loop_mode = (line->mArgc <= 1) ? FILE_LOOP_FILES_ONLY : ConvertLoopMode(ARG2);
				if (file_loop_mode == FILE_LOOP_INVALID)
					return line->LineError(ERR_PARAM2_INVALID, FAIL, ARG2);
			}

			// ONLY AFTER THE ABOVE IS IT CERTAIN THE LOOP WILL LAUNCH (i.e. there was no error or early return).
			// So only now is it safe to make changes to global things like g.mLoopIteration.
			jump_to_line = NULL; // Init this output parameter prior to starting each type of loop.

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
			if (line->mActionType != ACT_FOR) // PerformLoopFor() sets it later so its enumerator expression (which is evaluated only once) can refer to the A_Index of the outer loop.
				g.mLoopIteration = 1;

			// PERFORM THE LOOP:
			switch (line->mActionType)
			{
			case ACT_LOOP: // Listed first for performance.
				bool is_infinite; // "is_infinite" is more maintainable and future-proof than using LLONG_MAX to simulate an infinite loop. Plus it gives peace-of-mind and the LLONG_MAX method doesn't measurably improve benchmarks (nor does BOOL vs. bool).
				__int64 iteration_limit;
				if (line->mArgc > 0) // At least one parameter is present.
				{
					iteration_limit = line->ArgToInt64(1); // If it's negative, zero iterations will be performed automatically.
					is_infinite = false;
				}
				else // No parameters are present.
				{
					iteration_limit = 0; // Avoids debug-mode's "used without having been defined" (though it's merely passed as a parameter, not ever used in this case).
					is_infinite = true;  // Override the default set earlier.
				}
				result = line->PerformLoop(aResultToken, jump_to_line, until, iteration_limit, is_infinite);
				break;
			case ACT_WHILE:
				result = line->PerformLoopWhile(aResultToken, jump_to_line);
				break;
			case ACT_FOR:
				result = line->PerformLoopFor(aResultToken, jump_to_line, until);
				break;
			case ACT_LOOP_PARSE:
				// The phrase "csv" is unique enough since user can always rearrange the letters
				// to do a literal parse using C, S, and V as delimiters:
				if (_tcsicmp(ARG2, _T("CSV")))
					result = line->PerformLoopParse(aResultToken, jump_to_line, until);
				else
					result = line->PerformLoopParseCSV(aResultToken, jump_to_line, until);
				break;
			case ACT_LOOP_READ:
				result = line->PerformLoopReadFile(aResultToken, jump_to_line, until, ARG1, ARG2);
				break;
			case ACT_LOOP_FILE:
				result = line->PerformLoopFilePattern(aResultToken, jump_to_line, until
					, (file_loop_mode & ~FILE_LOOP_RECURSE), (file_loop_mode & FILE_LOOP_RECURSE), ARG1);
				break;
			case ACT_LOOP_REG:
				// This isn't the most efficient way to do things (e.g. the repeated calls to
				// RegConvertRootKey()), but it the simplest way for now.  Optimization can
				// be done at a later time:
				bool is_remote_registry;
				HKEY root_key;
				LPTSTR subkey;
				if (root_key = RegConvertKey(ARG1, &subkey, &is_remote_registry)) // This will open the key if it's remote.
				{
					// root_key_type needs to be passed in order to support GetLoopRegKey():
					result = line->PerformLoopReg(aResultToken, jump_to_line, until
						, (file_loop_mode & ~FILE_LOOP_RECURSE), (file_loop_mode & FILE_LOOP_RECURSE), root_key_type, root_key, subkey);
					if (is_remote_registry)
						RegCloseKey(root_key);
				}
				else
					// The open of a remote key failed (we know it's remote otherwise it should have
					// failed earlier rather than here).
					result = g_script.Win32Error();
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
			// sync with the current nesting level (braces, function calls, etc.)
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
			if (finished_line->mActionType == ACT_ELSE)
			{
				if (result == CONDITION_FALSE)
				{
					line = finished_line->mNextLine;
					continue;
				}
				finished_line = finished_line->mRelatedLine;
			}
			if (aMode == ONLY_ONE_LINE)
				return OK;
			// Since the above didn't return or break, either the loop has completed the specified
			// number of iterations or it was broken via the break command.  In either case, we jump
			// to the line after our loop's structure and continue there:
			line = finished_line;
			continue;  // Resume looping starting at the above line.  "continue" is actually slightly faster than "break" in these cases.
		} // case ACT_LOOP, ACT_LOOP_*, ACT_FOR, ACT_WHILE.

		case ACT_TRY:
		case ACT_CATCH:
		{
			ActionTypeType this_act = line->mActionType;
			int outer_excptmode = g.ExcptMode;

			ResultToken *our_token = g.ThrownToken; // Should always be non-null for CATCH and null for TRY.
			g.ThrownToken = nullptr; // Must be done before assigning the output var in case Assign() causes script to execute via __Delete.

			if (this_act == ACT_CATCH)
			{
				ASSERT(our_token);
				g.ExcptMode |= EXCPTMODE_CAUGHT;

				// Assign the thrown token to the variable if provided.
				result = OK;
				if (auto args = (CatchStatementArgs *)line->mAttribute)
					if (args->output_var)
						result = args->output_var->Assign(*our_token);
				if (!result)
				{
					g_script.FreeExceptionToken(our_token);
					return FAIL;
				}
			}
			else // (this_act == ACT_TRY)
			{
				//g.ExcptMode |= EXCPTMODE_TRY; // Currently unused.  Must use |= rather than = to avoid removing EXCPTMODE_CATCH, if present.
				if (line->mRelatedLine->mActionType != ACT_FINALLY) // Try without Catch/Finally acts like it has an empty Catch.
				{
					g.ExcptMode |= EXCPTMODE_CATCH;
					// Disable re-throw when catch is present, since the inner layer won't have the Error object to match
					// against type filters such as `try..catch SpecificError`, and fixing that would increase complexity
					// or the size of global_struct.
					g.ExcptMode &= ~EXCPTMODE_CAUGHT;
				}
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

			bool UnhandledException_was_not_called = (g.ExcptMode & EXCPTMODE_CATCH); // Because CATCH is present (or TRY without CATCH or FINALLY).
			if (our_token) // Implies this is CATCH, or TRY but somehow g.ThrownToken was already non-null due to a prior exception.
			{
				if (!result && !g.ThrownToken) // Assume this is a re-throw, since any other failure should have generated a new exception.
				{
					g.ThrownToken = our_token, our_token = nullptr;
					UnhandledException_was_not_called = true; // Because THROW didn't have our_token to pass to it.
				}
				else
					g_script.FreeExceptionToken(our_token); // Get this out of the way now in case of early "goto" or "return".
			}
			g.ExcptMode = outer_excptmode; // Reset it now in case of early "goto" or "return".

			// Move to the next line after the 'try' or 'catch' block.
			line = line->mRelatedLine;

			bool bHasCatch = false;
			if (this_act == ACT_TRY)
			{
				while (line->mActionType == ACT_CATCH)
				{
					bHasCatch = true;
					if (g.ThrownToken && line->CatchThis(*g.ThrownToken))
					{
						// An exception was thrown and we have a 'catch' block, so let the next
						// iteration handle it.  Implies result == FAIL && jump_to_line == NULL,
						// but result won't have any meaning for the next iteration.
						goto continue_ExecUntil_loop;
					}
					// Otherwise: an exception of appropriate type was not thrown, so skip this 'catch'.
					line = line->mRelatedLine;
				}
				if (!bHasCatch && g.ThrownToken && line->mActionType != ACT_FINALLY
					&& Object::HasBase(*g.ThrownToken, ErrorPrototype::Error)) // Make `try {}` equivalent to `try {} catch {}` without a class.
				{
					// An exception was thrown, but no 'catch' nor 'finally' is present.
					// In this case 'try' acts as a catch-all.
					g_script.FreeExceptionToken(g.ThrownToken);
					result = OK;
				}
				if (!g.ThrownToken && line->mActionType == ACT_ELSE && result == OK && !jump_to_line)
				{
					// Since no exception was thrown, execute the ELSE (by jumping to its first statement).
					line = line->mNextLine;
					continue;
				}
			}
			else // this_act == ACT_CATCH
			{
				// Even if g.ThrownToken, no other 'catch' statements at this level should execute.
				while (line->mActionType == ACT_CATCH)
					line = line->mRelatedLine;
			}
			if (line->mActionType == ACT_ELSE)
				line = line->mRelatedLine;
			if (line->mActionType == ACT_FINALLY)
			{
				if (result == OK && !jump_to_line)
					// Execution is being allowed to flow normally into the finally block and
					// then the line after.  Let the next iteration handle the finally block.
					continue;

				// One of the following occurred:
				//  - An exception was thrown, and this try..(catch)..finally block didn't handle it.
				//  - A control flow statement such as break, continue or goto was used.
				// Recursively execute the finally block before continuing.
				ResultToken *thrown_token = g.ThrownToken;
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
						// two reasons: 1) if the script was non-persistent it could have already terminated
						// anyway, and 2) it's only a question of whether to show a message before exiting.
						return res;
					// The remaining cases are all invalid jumps/control flow statements.  All such cases
					// should be detected at load time, but it seems best to keep this for maintainability:
					return g_script.mCurrLine->LineError(ERR_BAD_JUMP_INSIDE_FINALLY);
				}
				g.ThrownToken = thrown_token; // If non-NULL, this was thrown within the try block.
			}
			if (g.ThrownToken && UnhandledException_was_not_called && !(g.ExcptMode & EXCPTMODE_CATCH))
			{
				// Known limitation: UnhandledException wasn't called when the error was thrown since a CATCH
				// was present, but it wasn't caught due to the type filter, so we have to do this here so the
				// error will be reported.  The idea of calling UnhandledException early is that OnError can
				// attach a debugger which can inspect the stack before it unwinds, but at this point it has
				// already partially unwound.
				// mCurrLine is more likely to be relevant than `line`, which at this point is probably the
				// line after the TRY/ELSE.  This is only used if ThrownToken doesn't have valid File/Line
				// properties; i.e. it's not a proper Error, so it's something not intended for error-reporting.
				return g_script.UnhandledException(g_script.mCurrLine);
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
			continue_ExecUntil_loop:
			continue;
		}

		case ACT_THROW:
		{
			if (!line->mArgc)
			{
				// If we're executing within a CATCH statement, signal it to re-throw its exception.
				// Otherwise, throw a new generic Error.
				return (g.ExcptMode & EXCPTMODE_CAUGHT) ? FAIL : line->ThrowRuntimeException(ERR_EXCEPTION);
			}

			// ThrownToken should only be non-NULL while control is being passed up the
			// stack, which implies no script code can be executing.
			ASSERT(!g.ThrownToken);

			ResultToken* token = new ResultToken;
			token->symbol = SYM_STRING; // Set default. ExpandArgs() mightn't set it.
			token->mem_to_free = NULL;
			token->marker_length = -1;

			result = line->ExpandArgs(token);
			if (result != OK)
			{
				// A script-function-call inside the expression returned EARLY_EXIT or FAIL.
				delete token;
				return result;
			}

			if (token->symbol == SYM_STRING && !token->mem_to_free)
			{
				// There are a couple of shortcuts we could take here, but they aren't taken
				// because throwing a string is pretty rare and not a performance priority.
				// Shortcut #1: Assign _T("") if ARG1 is empty.
				// Shortcut #2: Take over sDerefBuf if ARG1 == sDerefBuf.
				if (!token->Malloc(token->marker, token->marker_length))
				{
					delete token;
					// See ThrowRuntimeException() for comments.
					MsgBox(ERR_OUTOFMEM ERR_ABORT);
					return FAIL;
				}
			}

			// Throw the newly-created token
			line->SetThrownToken(g, token);
			return FAIL;
		}

		case ACT_FINALLY:
		{
			// Directly execute next line
			line = line->mNextLine;
			continue;
		}

		case ACT_SWITCH:
		{
			Line *line_to_execute = NULL;

			// Privatize our deref buf so that any function calls within any of the SWITCH/CASE
			// expressions can allocate and use their own separate deref buf.  Our deref buf
			// will be reused for each expression evaluation.
			PRIVATIZE_S_DEREF_BUF;

			size_t switch_value_mem_size = 0;
			ResultToken switch_value;
			switch_value.mem_to_free = NULL;
			SymbolType switch_is_numeric = PURE_NOT_NUMERIC;
			StringCaseSenseType string_case_sense = SCS_INVALID; // In this case, the default value indicates numeric comparison should be used when appropriate.
			if (!line->mArgc) // Switch with no value: find the first 'true' case.
			{
				switch_value.symbol = SYM_INVALID;
				result = OK;
			}
			else
			{
				// Load time validation has ensured that at least the first parameter is present in this case.
				result = line->ExpandSingleArg(0, switch_value, our_deref_buf, our_deref_buf_size);
				if (result == OK  // Do not bother with the CaseSense parameter if the switch value expression failed.
					&& line->mArgc == 2)
				{
					// The CaseSense parameter was specified.
					ResultToken case_sense_value;
					case_sense_value.mem_to_free = NULL;
					size_t mem_size = 0;
					result = line->ExpandSingleArg(1, case_sense_value, case_sense_value.mem_to_free, mem_size);
					if (result == OK)
					{
						string_case_sense = TokenToStringCase(case_sense_value);

						result = string_case_sense >= SCS_INSENSITIVE_LOGICAL ? line->LineError(ERR_PARAM2_INVALID) : OK;
					}
					case_sense_value.Free();
				}
				if (result == OK)
				{
					if (string_case_sense == SCS_INVALID)
					{
						// In this mode, numeric comparison is performed if both switch_value and case_value are
						// numeric.  For performance, call IsNumeric() for switch_value only once.  If non-numeric,
						// IsNumeric() calls can be avoided for each case_value.  If it's an object, the comparison
						// result will be false (not an error) if case_value is not an object, as for a == b.
						// If string comparison ends up being performed, 
						switch (switch_value.symbol)
						{
						case SYM_STRING: switch_is_numeric = IsNumeric(switch_value.marker, TRUE, FALSE, TRUE); break;
						case SYM_OBJECT: break; // Leave switch_is_numeric at its default value, PURE_NOT_NUMERIC.
						default: switch_is_numeric = switch_value.symbol; break;
						}
					}
					else
					{
						// Since the CaseSense parameter was specified, perform string comparison for every case.
						// Avoid calling IsNumeric(), for performance and because its result is not relevant.
						// It doesn't make sense to specify case sensitivity for objects, so treat that as an error
						// (otherwise EvaluateSwitchCase would ignore string_case_sense). 
						if (switch_value.symbol == SYM_OBJECT)
							result = TypeError(_T("String"), switch_value);
					}
				}
			}
			if (result == OK)
			{
				if (switch_value.symbol == SYM_STRING && switch_value.marker == our_deref_buf)
				{
					// Prevent the case expressions from reusing our_deref_buf, since we'll need it.
					// A new buf will be allocated by ExpandSingleArg() if required, which would only
					// be if at least one case expression is not a literal number/quoted string.
					switch_value.mem_to_free = our_deref_buf;
					switch_value_mem_size = our_deref_buf_size;
					our_deref_buf = NULL;
					our_deref_buf_size = 0;
				}
				// For each CASE:
				for (Line *case_line = line->mNextLine->mNextLine; case_line && case_line->mActionType == ACT_CASE; case_line = case_line->mRelatedLine)
				{
					int arg, arg_count = case_line->mArgc;
					if (!arg_count) // The default case.
					{
						line_to_execute = case_line->mNextLine;
						continue;
					}
					for (arg = 0; arg < arg_count; ++arg)
					{
						ResultToken case_value;
						result = case_line->ExpandSingleArg(arg, case_value, our_deref_buf, our_deref_buf_size);
						if (result != OK)
						{
							line_to_execute = NULL; // Do not execute default case.
							break;
						}
						bool found;
						if (switch_value.symbol != SYM_INVALID)
						{
							result = EvaluateSwitchCase(switch_value, switch_is_numeric, case_value, string_case_sense);
							if (result != CONDITION_TRUE && result != CONDITION_FALSE)
							{
								line_to_execute = NULL; // Do not execute default case.
								break;
							}
							found = result == CONDITION_TRUE;
							result = OK;
						}
						else
							found = TokenToBOOL(case_value);
						if (case_value.symbol == SYM_OBJECT)
							case_value.object->Release();
						if (found)
						{
							line_to_execute = case_line->mNextLine;
							break;
						}
					}
					if (arg < arg_count)
						break;
				}
				if (switch_value_mem_size)
				{
					if (our_deref_buf)
					{
						// Free the newly allocated deref buf.
						free(our_deref_buf);
						if (our_deref_buf_size > LARGE_DEREF_BUF_SIZE)
							--sLargeDerefBufs;
					}
					// Restore original deref buf.
					our_deref_buf = switch_value.mem_to_free;
					our_deref_buf_size = switch_value_mem_size;
				}
				else
					switch_value.Free();
			}

			DEPRIVATIZE_S_DEREF_BUF;

			if (line_to_execute)
			{
				// Above found a matching CASE.  Execute the lines between it and the next CASE or block-end.
				result = line_to_execute->ExecUntil(UNTIL_BLOCK_END, aResultToken, &jump_to_line);
			}
			else
				jump_to_line = NULL;

			if (result != OK || aMode == ONLY_ONE_LINE)
			{
				caller_jump_to_line = jump_to_line;
				return result;
			}
			
			if (jump_to_line != NULL)
			{
				if (jump_to_line->mParentLine != line->mParentLine)
				{
					caller_jump_to_line = jump_to_line;
					return OK;
				}
				line = jump_to_line;
				continue;
			}
			
			// Continue execution at the line following the block-end.
			line = line->mRelatedLine;
			continue;
		}

		case ACT_CASE:
			// This is the next CASE after one that matched, so we're done.
			return OK;

#ifndef CONFIG_DLL
#endif
		case ACT_BLOCK_BEGIN:
			if (line->mAttribute) // This is the ACT_BLOCK_BEGIN that starts a function's body.
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
			// v2: This check is disabled to reduce code size, as it doesn't seem to be needed
			// now that GOSUB has been removed.  Validation in PreparseBlocks() should make it
			// impossible to produce this condition:
			//if (aMode != UNTIL_BLOCK_END)
			//	return line->LineError(_T("A \"return\" must be encountered prior to this \"}\"."));
			return OK; // It's the caller's responsibility to resume execution at the next line, if appropriate.

		// ACT_ELSE can happen when one of the cases in this switch failed to properly handle
		// aMode == ONLY_ONE_LINE.  But even if ever happens, it will just drop into the default
		// case, which will result in a FAIL (silent exit of thread) as an indicator of the problem.
		// So it's commented out:
		//case ACT_ELSE:
		//	// Shouldn't happen if the pre-parser and this function are designed properly?
		//	return line->LineError("Unexpected ELSE.");

		case ACT_ASSIGNEXPR:
			result = line->PerformAssign();
			// Fall through:
		case ACT_EXPRESSION:
			// Nothing else needs to be done because the expression in ARG1 (which is the only arg)
			// has already been evaluated and its functions and subfunctions called.  Examples:
			//    fn(123, "string", var, fn2(y))
			//    x&=3
			//    var ? func() : x:=y
			if (result != OK || aMode == ONLY_ONE_LINE)
				return result; // Usually OK or FAIL; can also be EARLY_EXIT.
			line = line->mNextLine;
			continue;

		case ACT_STATIC:
			if (result != OK)
				return result;
			// If this is the function's first line, must update mJumpToLine so that it will skip this one.
			if (g.CurrentFunc && g.CurrentFunc->mJumpToLine == line)
				g.CurrentFunc->mJumpToLine = line->mNextLine;
			// Update any If/Else/Loop or similar (there can be multiple nested) that immediately precedes
			// this line so that they do not jump to this one after executing their body.
			for (Line *parent = line->mPrevLine->mParentLine; parent && parent->mRelatedLine == line; parent = parent->mParentLine)
				parent->mRelatedLine = line->mNextLine;
			// Remove this line from the main list, so that it won't be executed again during normal flow.
			line->mPrevLine->mNextLine = line->mNextLine;
			line->mNextLine->mPrevLine = line->mPrevLine;
			if (aMode == ONLY_ONE_LINE)
				return result;
			line = line->mNextLine;
			continue;

		case ACT_EXIT: // In this context it's the ACT_EXIT added automatically by LoadFromFile().
			return EARLY_EXIT;

#ifdef _DEBUG
		default:
			return LineError(_T("DEBUG: ExecUntil(): Unhandled action type."));
#endif
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



ResultType Line::EvaluateCondition()
// Returns CONDITION_TRUE or CONDITION_FALSE (FAIL is returned only in DEBUG mode).
{
#ifdef _DEBUG
	if (mActionType != ACT_IF)
		return LineError(_T("DEBUG: EvaluateCondition() was called with a line that isn't a condition."));
#endif

	// The following is ordered for short-circuit performance.
	Var *var = mArg[0].type == ARG_TYPE_INPUT_VAR ? VAR(mArg[0]) : nullptr;
	bool if_condition = var ? VarToBOOL(*var) // Not verified recently: 30% faster than having ExpandArgs() resolve ARG1 even when it's a naked variable.
		: ResultToBOOL(ARG1); // CAN'T simply check *ARG1=='1' because the loadtime routine has various ways of setting if_expresion to false for things that are normally expressions.

	return if_condition ? CONDITION_TRUE : CONDITION_FALSE;
}



ResultType Line::EvaluateSwitchCase(ExprTokenType &aSwitch, SymbolType aSwitchIsNumeric, ExprTokenType &aCase, StringCaseSenseType aStringCaseSense)
// Performs comparison of two tokens for Switch().
{
	SymbolType case_is_numeric = PURE_NOT_NUMERIC;
	if (aStringCaseSense == SCS_INVALID)
	{
		// Perform numeric comparison if both aSwitch and aCase are numeric and at least one
		// is pure numeric, not a string.  This should match the behaviour of SYM_EQUALCASE.
		if (aSwitchIsNumeric && !(aSwitch.symbol == SYM_STRING && aCase.symbol == SYM_STRING))
			case_is_numeric = TokenIsNumeric(aCase);
		//else no need to call TokenIsNumeric(); just leave case_is_numeric set to its default
		// value so that string comparison will be used.
	}
	else // Since CaseSense was specified, unconditionally use string comparison.
	{
		if (aCase.symbol == SYM_OBJECT)
			return TypeError(_T("String"), aCase);
		// Caller has unconditionally set aSwitchIsNumeric to PURE_NOT_NUMERIC in this case,
		// and has already verified aSwitch.symbol != SYM_OBJECT.
	}
	bool equal;
	if (aSwitch.symbol == SYM_OBJECT || aCase.symbol == SYM_OBJECT)
		equal = aSwitch.symbol == aCase.symbol && aSwitch.object == aCase.object;
	else if (!case_is_numeric)
	{
		TCHAR left_buf[MAX_NUMBER_SIZE], *left_string = TokenToString(aSwitch, left_buf);
		TCHAR right_buf[MAX_NUMBER_SIZE], *right_string = TokenToString(aCase, right_buf);
		switch (aStringCaseSense)
		{
		case SCS_INSENSITIVE:			equal = !_tcsicmp(left_string, right_string); break;
		case SCS_INSENSITIVE_LOCALE:	equal = !lstrcmpi(left_string, right_string); break;
		default:						equal = !_tcscmp(left_string, right_string); break;
		}
	}
	else if (aSwitchIsNumeric == PURE_INTEGER && case_is_numeric == PURE_INTEGER)
		equal = TokenToInt64(aSwitch) == TokenToInt64(aCase);
	else
		equal = TokenToDouble(aSwitch) == TokenToDouble(aCase);
	return equal ? CONDITION_TRUE : CONDITION_FALSE;
}



// Evaluate an #HotIf expression or callback function.
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
	// Critical seems to improve reliability, either because the thread completes faster (i.e. before the timeout) or because we check for messages less often.
	InitNewThread(0, false, true, true);

	// Let HotIf default to the criterion currently being evaluated, in case Hotkey() is called.
	g->HotCriterion = this;

	// Update A_ThisHotkey, useful if #HotIf calls a function to do its dirty work.
	LPTSTR prior_hotkey_name[] = { g_script.mThisHotkeyName, g_script.mPriorHotkeyName };
	DWORD prior_hotkey_time[] = { g_script.mThisHotkeyStartTime, g_script.mPriorHotkeyStartTime };
	g_script.mPriorHotkeyName = g_script.mThisHotkeyName;			// For consistency
	g_script.mPriorHotkeyStartTime = g_script.mThisHotkeyStartTime; //
	g_script.mThisHotkeyName = aHotkeyName;
	g_script.mThisHotkeyStartTime = // Updated for consistency.
	g_script.mLastPeekTime = GetTickCount();

	// CALL THE CALLBACK
	ExprTokenType param = aHotkeyName;
	auto result = IObjectPtr(Callback)->ExecuteInNewThread(_T("#HotIf"), &param, 1, true);

	// The following allows the expression to set the Last Found Window for the
	// hotkey function.
	// There may be some rare cases where the wrong hotkey gets this HWND (perhaps
	// if there are multiple hotkey messages in the queue), but there doesn't seem
	// to be any easy way around that.
	g_HotExprLFW = g->hWndLastUsed; // Even if above failed, for simplicity.

	// A_ThisHotkey must be restored else A_PriorHotkey will get an incorrect value later.
	g_script.mThisHotkeyName = prior_hotkey_name[0];
	g_script.mThisHotkeyStartTime = prior_hotkey_time[0];
	g_script.mPriorHotkeyName = prior_hotkey_name[1];
	g_script.mPriorHotkeyStartTime = prior_hotkey_time[1];

	ResumeUnderlyingThread();

	g_DeferMessagesForUnderlyingPump = prev_defer_messages;

	return result;
}



// The following macro is used to reduce repetition, in lieu of an inline function,
// which wouldn't be fully inlined.  Benchmarks showed this affected performance of
// very simple loops by a few percent.  Making it a macro also allows repetition of
// the result checks to be avoided.
//
// Simple loops are about 6% faster by omitting the open-brace from ListLines, so
// it seems worth it:
//if (g.ListLinesIsEnabled)
//	LOG_LINE(mNextLine)
//
// If this loop has a block under it rather than just a single line, take a shortcut
// and directly execute the block.  This avoids one recursive call to ExecUntil()
// for each iteration, which can speed up short/fast loops by as much as 30%.
// Another benefit is conservation of stack space, especially during "else if" ladders.
//
// At this point, mNextLine is already verified non-NULL by the pre-parser because it
// checks that every LOOP has a line under it (ACT_BLOCK_BEGIN in this case) and that
// every ACT_BLOCK_BEGIN has at least one line under it.
#define PERFORMLOOP_EXECUTE_BODY \
	if (mNextLine->mActionType == ACT_BLOCK_BEGIN) \
	{ \
		do \
			result = mNextLine->mNextLine->ExecUntil(UNTIL_BLOCK_END, aResultToken, &jump_to_line); \
		while (jump_to_line == mNextLine); \
	} \
	else \
		result = mNextLine->ExecUntil(ONLY_ONE_LINE, aResultToken, &jump_to_line); /* i.e. "goto somewhere" or "continue a_loop_which_encloses_this_one". */ \
	if (jump_to_line && !(result == LOOP_CONTINUE && jump_to_line == this)) \
	{ /* jump_to_line must be a line that's at the same level or higher as our Exec_Until's LOOP statement itself. */ \
		aJumpToLine = jump_to_line; /* Signal the caller to handle this jump. */ \
		break; \
	} \
	if (result == LOOP_CONTINUE) /* Continue this loop (outer loops were handled via jump_to_line above). */ \
		result = OK; /* Don't propagate the LOOP_CONTINUE result since it has already achieved its purpose. */ \
	else if (result != OK) /* i.e. result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL) */ \
		break;
#define PERFORMLOOP_EVALUATE_UNTIL \
	if (aUntil && aUntil->EvaluateLoopUntil(result)) \
		break;



ResultType Line::PerformLoop(ResultToken *aResultToken, Line *&aJumpToLine, Line *aUntil
	, __int64 aIterationLimit, bool aIsInfinite) // bool performs better than BOOL in current benchmarks for this.
// This performs much better (by at least 7%) as a function than as inline code, probably because
// it's only called to set up the loop, not each time through the loop.
{
	ResultType result = CONDITION_FALSE;
	Line *jump_to_line = nullptr;
	global_struct &g = *::g; // Primarily for performance in this case.

	for (; aIsInfinite || g.mLoopIteration <= aIterationLimit; ++g.mLoopIteration)
	{
		// Execute the body of the loop (either just one statement or a block of statements).
		// Preparser has ensured that every LOOP has a non-NULL next line.
		PERFORMLOOP_EXECUTE_BODY
		PERFORMLOOP_EVALUATE_UNTIL
		// Since the macro didn't "break", the result of executing the body of the loop was either OK
		// (the current iteration completed normally) or LOOP_CONTINUE (the current loop iteration
		// was cut short).  In both cases, just continue on through the loop.
	} // for()

	// The script's loop is now over.
	return result;
}



// Lexikos: ACT_WHILE
ResultType Line::PerformLoopWhile(ResultToken *aResultToken, Line *&aJumpToLine)
{
	ResultType result = CONDITION_FALSE;
	Line *jump_to_line = nullptr;
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
		auto args_result = ExpandArgs();
		if (args_result != OK)
			return args_result;

		// Unlike if(expression), performance isn't significantly improved to make cases like
		// "while x" and "while %x%" into non-expressions (the latter actually performs much
		// better as an expression).  That is why the following check is much simpler than the
		// one used at ACT_IF in EvaluateCondition():
		if (!ResultToBOOL(ARG1))
			break;

		PERFORMLOOP_EXECUTE_BODY

		// Before re-evaluating the condition, add it to the ListLines log again.  This is done
		// at the end of the loop rather than the beginning because ExecUntil already added the
		// line once immediately before the first iteration.
		if (g.ListLinesIsEnabled)
			LOG_LINE(this)
	} // for()
	return result; // The script's loop is now over.
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
	return ResultToBOOL(ARG1); // See PerformLoopWhile() above for comments about this line.
}



ResultType Line::PerformLoopFor(ResultToken *aResultToken, Line *&aJumpToLine, Line *aUntil)
{
	ResultType result;
	Line *jump_to_line;
	global_struct &g = *::g; // Might slightly speed up the loop below.

	ResultToken enum_token;
	enum_token.InitResult(nullptr); // buf can be null because this isn't ACT_RETURN.
	PRIVATIZE_S_DEREF_BUF;
	result = ExpandSingleArg(mArgc - 1, enum_token, our_deref_buf, our_deref_buf_size);
	DEPRIVATIZE_S_DEREF_BUF;
	if (result != OK)
		// A script-function-call inside the expression returned EARLY_EXIT or FAIL.
		return result;
	
	int var_count = mArgc - 1;
	
	IObject *enumerator;
	result = GetEnumerator(enumerator, enum_token, var_count, true);
	enum_token.Free();
	if (result == FAIL || result == EARLY_EXIT)
		return result;

	// "Localize" the loop variables.
	auto var_bkp = (VarBkp *)_alloca(sizeof(VarBkp) * var_count);
	auto var_param = (ExprTokenType**)_alloca(sizeof(ExprTokenType*) * var_count);
	for (int i = 0; i < var_count; ++i)
	{
		ExprTokenType &token = *(ExprTokenType*)_alloca(sizeof(ExprTokenType));
		var_param[i] = &token;
		if (mArg[i].type == ARG_TYPE_OUTPUT_VAR)
		{
			auto var = VAR(mArg[i]);
			var->Backup(var_bkp[i]);
			token.symbol = SYM_OBJECT;
			token.object = var->GetRef();
			if (token.object)
				continue;
			result = FAIL;
			var_count = i + 1; // Restore var_bkp and free any previous refs.
		}
		// Omitted.
		token.symbol = SYM_MISSING;
	}

	// Now that the enumerator expression has been evaluated, init A_Index:
	g.mLoopIteration = 1;

	if (result != FAIL)
	for (result = CONDITION_FALSE;; ++g.mLoopIteration)
	{
		auto enum_result = CallEnumerator(enumerator, var_param, var_count, true);
		if (enum_result == FAIL || enum_result == EARLY_EXIT)
		{
			result = enum_result;
			break;
		}
		if (enum_result != CONDITION_TRUE) // The enumerator returned false, which means there are no more items.
			break;
		// Otherwise the enumerator already stored the next value(s) in the variable(s) we passed it via params.

		PERFORMLOOP_EXECUTE_BODY
		PERFORMLOOP_EVALUATE_UNTIL
	} // for()
	enumerator->Release();
	for (int i = 0; i < var_count; ++i)
	{
		if (mArg[i].type == ARG_TYPE_OUTPUT_VAR)
		{
			Var *var = VAR(mArg[i]);
			var->Free(VAR_NEVER_FREE | VAR_CLEAR_ALIASES); // Release var's reference to the VarRef.
			var->Restore(var_bkp[i]);
		}
		if (var_param[i]->symbol == SYM_OBJECT)
			var_param[i]->object->Release();
	}
	return result; // The script's loop is now over.
}



ResultType Line::PerformLoopFilePattern(ResultToken *aResultToken, Line *&aJumpToLine, Line *aUntil
	, FileLoopModeType aFileLoopMode, bool aRecurseSubfolders, LPTSTR aFilePattern)
{
	ResultType result = CONDITION_FALSE;
	// LoopFilesStruct is currently about 128KB, so it's probably best not to put it on the stack.
	// 128KB temporary usage per Loop (for all iterations) seems a reasonable trade-off for supporting
	// long paths while keeping code size minimal.
	LoopFilesStruct *plfs = new LoopFilesStruct;
	if (!plfs)
		return MemoryError();
	// Parse aFilePattern into its components and copy into *plfs.  Copies are taken because:
	//  - As the lines of the loop are executed, the deref buffer (which is what aFilePattern might
	//    point to if we were called from ExecUntil()) may be overwritten -- and we will need the path
	//    string for every loop iteration.
	//  - We relative paths resolved to full paths in case the working directory is changed after the
	//    loop begins but before FindFirstFile() is called for a sub-directory.
	//  - Several built-in variables rely on the paths stored in *plfs and updated on each iteration.
	//    This approach means a little up-front work that mightn't be needed, but greatly improves
	//    performance in some common cases, such as if A_LoopFileLongPath is used on each iteration.
	if (ParseLoopFilePattern(aFilePattern, *plfs, result))
		result = PerformLoopFilePattern(aResultToken, aJumpToLine, aUntil, aFileLoopMode, aRecurseSubfolders, *plfs);
	//else: leave result == CONDITION_FALSE, since in effect, no files were found.
	delete plfs;
	return result;
}



bool Line::ParseLoopFilePattern(LPTSTR aFilePattern, LoopFilesStruct &lfs, ResultType &aResult)
// Parse aFilePattern and initialize lfs.
{
	if (!*aFilePattern) // Some checks below may rely on empty aFilePattern having been excluded.
		// FindFirstFile() would find nothing in this case, but continuing may cause unnecessary
		// rescursion through all subdirectories of the working directory.
		return false;
	// Note: Even if aFilePattern is just a directory (i.e. with no wildcard pattern), it seems best
	// not to append "\\*.*" to it because the pattern might be a script variable that the user wants
	// to conditionally resolve to various things at runtime.  In other words, it's valid to have
	// only a single directory be the target of the loop.
	
	// v1.1.31.00: This function was revised.
	//  - Resolve aFilePattern to a full path immediately so that changing the working directory
	//    does not disrupt recursion or A_LoopFileLongPath/ShortPath.
	//  - Take care that path lengths are not limited to below what the system allows.
	//    That is, directory+pattern can be up to MAX_PATH on ANSI and 32767 on Unicode,
	//    although the latter requires the \\?\ prefix or Windows 10 long path awareness.
	//  - Optimize A_LoopFileLongPath and A_LoopFileShortPath by resolving the path prefix
	//    up-front and utilizing the names returned by FindFirstFile() during the loop.
	//    This can be much faster because file system access is minimized.  Benchmarks
	//    showed improvement even for single-iteration loops which use A_LoopFileLongPath.
	
	// Prior to v1.1.31.00, A_LoopFileLongPath worked as follows:
	//  - Each reference to A_LoopFileLongPath results in two BIV_LoopFileLongPath calls.
	//  - BIV_LoopFileLongPath calls GetFullPathName() and ConvertFilespecToCorrectCase().
	//  - CFTCC() calls FindFirstFile() once for each slash-delimited name.
	//  - For example, with "c:\foo\bar\baz.txt", A_LoopFileLongPath results in SIX calls to
	//    FindFirstFile().

	// Find the final name or pattern component of aFilePattern.
	LPTSTR name_part, cp;
	for (name_part = cp = aFilePattern; *cp; cp++)
		if (*cp == '\\' || *cp == '/')
			name_part = cp + 1;

	if (name_part == aFilePattern && *name_part && name_part[1] == ':') // Single character followed by ':' but not '\\' or '/'.
		name_part += 2;
	
	size_t pattern_length = cp - name_part;
	if (pattern_length > _countof(lfs.pattern)) // Most likely too long to match a real path/filename.
		return false; 
	tmemcpy(lfs.pattern, name_part, pattern_length + 1);
	lfs.pattern_length = pattern_length;

	size_t orig_dir_length = name_part - aFilePattern;
	if (orig_dir_length)
	{
		if (  !(lfs.orig_dir = tmalloc(orig_dir_length + 1))  )
		{
			aResult = MemoryError();
			return false;
		}
		tmemcpy(lfs.orig_dir, aFilePattern, orig_dir_length);
		lfs.orig_dir[orig_dir_length] = '\0';
		lfs.orig_dir_length = orig_dir_length;
	}
	else
	{
		lfs.orig_dir = _T("");
		lfs.orig_dir_length = 0;
	}

	// The following aren't checked because A_LoopFileLongPath requires that GetFullPathName()
	// be called unconditionally so that relative references such as "x\..\y" are resolved:
	//if (   !(aFilePattern[1] == ':' && aFilePattern[2] == '\\') // Not an absolute path with drive letter (must have slash, as "C:xxx" is not fully qualified).
	//	&& !(*aFilePattern == '\\' && aFilePattern[1] == '\\')   ) // Not a UNC path.

	// Testing shows that GetFullPathNameW() supports longer than MAX_PATH even on Windows XP.
	// MSDN: "In the ANSI version of this function, the name is limited to MAX_PATH characters.
	//  To extend this limit to 32,767 wide characters, call the Unicode version of the function
	//  (GetFullPathNameW), and prepend "\\?\" to the path. "
	// But that's obviously incorrect, since prepending "\\?\" would make it an absolute path.
	// Instead, we just call it without the prefix, and this works (on Unicode builds).
	LPCTSTR dir_to_resolve = *lfs.orig_dir ? lfs.orig_dir : _T(".\\"); // Include a trailing slash so there will be one in the result.
	lfs.file_path_length = GetFullPathName(dir_to_resolve, _countof(lfs.file_path), lfs.file_path, NULL);
	if (!lfs.file_path_length)
		// It's unclear under what conditions GetFullPathName() can fail, but the most likely
		// cause is that the buffer is too small (and even this is unlikely on Unicode builds).
		// With current buffer sizes, that implies the path is too long for FindFirstFile().
		return false;

	if (lfs.file_path[lfs.file_path_length - 1] != '\\') // aFilePattern was "x:pattern" with no slash.
	{
		lfs.file_path[lfs.file_path_length++] = '\\';
		lfs.file_path[lfs.file_path_length] = '\0';
	}
	
	// Mark the part of file_path which will contain discovered sub-directories/files.
	// This will be appended to orig_dir to get the value of A_LoopFilePath.
	lfs.file_path_suffix = lfs.file_path + lfs.file_path_length;
	
	// Correct case and convert any short names to long names for A_LoopFileLongPath.
	LPTSTR long_dir;
	TCHAR long_dir_buf[MAX_WIDE_PATH]; // Max expanded length is probably about 8k on ANSI, but this covers all builds.
	if (  !(long_dir = ConvertFilespecToCorrectCase(lfs.file_path, long_dir_buf, _countof(long_dir_buf), lfs.long_dir_length))  )
	{
		// Conversion failed.  In theory, FindFirstFile() can fail for a parent directory
		// but succeed for the full path, so use file_path as-is and attempt to continue.
		// Since this is rare, _tcsdup() is still used below rather than aliasing file_path
		// (which would require avoiding free() in that case).
		long_dir = lfs.file_path;
		lfs.long_dir_length = lfs.file_path_length;
	}
	if (  !(lfs.long_dir = _tcsdup(long_dir))  ) // Conserve memory during loop execution by not using the full MAX_WIDE_PATH.
	{
		aResult = MemoryError();
		return false;
	}

	// For simplicity/code size, get the short path unconditionally, even though it's more
	// rarely used than the long path (and not required for the loop to function).  Having the
	// short path built in advance helps performance for loops which use it.
	// Testing and research shows that GetShortPathName() uses the long name for a directory
	// or file if no short name exists, so it can be as long as orig_dir (also, see below).
	lfs.short_path_length = GetShortPathName(lfs.orig_dir, lfs.short_path, _countof(lfs.short_path));
	if (lfs.short_path_length > _countof(lfs.short_path)) // Buffer was too small.
		// This can only occur if the short path is longer than the long path (file_path),
		// which is possible for names which are short but don't comply with 8.3 rules,
		// such as ".vs", which might have the 8.3 name "VS4DA5~1".  Even so, file_path
		// must be near the limit for short_path to exceed it.
		return false; // short_path and short_path_length aren't valid, so just abort.

	return true; // aFilePattern parsed and all members initialized okay.
}



ResultType Line::PerformLoopFilePattern(ResultToken *aResultToken, Line *&aJumpToLine, Line *aUntil
	, FileLoopModeType aFileLoopMode, bool aRecurseSubfolders, LoopFilesStruct &lfs)
// This is the recursive part, called for each sub-directory when aRecurseSubfolders is true.
// Caller has allocated buffers (lfs) and filled in the initial paths and filename/pattern.
{
	// Save current lengths before modification.
	size_t file_path_length = lfs.file_path_length;
	size_t short_path_length = lfs.short_path_length;
	lfs.dir_length = file_path_length; // During the loop, lfs.file_path_length will include the filename.
	
	if (!lfs.pattern_length && lfs.orig_dir_length == 2 && lfs.orig_dir[1] == ':')
	{
		// Handle "C:" by changing the pattern to "." to match the directory itself,
		// otherwise it would fail since lfs.file_path contains a trailing slash.
		// Search for "=C:" in SetWorkingDir() for explanation of "C:" vs. "C:\".
		lfs.pattern[0] = '.';
		lfs.pattern[1] = '\0';
		lfs.pattern_length = 1;
		// Disable recursion since "." would otherwise be found in every sub-directory.
		aRecurseSubfolders = false;
	}

	LPTSTR file_path_end = lfs.file_path + file_path_length;
	size_t file_space_remaining = _countof(lfs.file_path) - file_path_length;
	if (lfs.pattern_length >= file_space_remaining)
		return CONDITION_FALSE;
	tmemcpy(file_path_end, lfs.pattern, lfs.pattern_length + 1); // file_path already includes the slash.

	BOOL file_found;
	HANDLE file_search = FindFirstFile(lfs.file_path, &lfs);
	for ( file_found = (file_search != INVALID_HANDLE_VALUE) // Convert FindFirst's return value into a boolean so that it's compatible with FindNext's.
		; file_found && FileIsFilteredOut(lfs, aFileLoopMode)
		; file_found = FindNextFile(file_search, &lfs));
	// file_found and lfs have now been set for use below.
	// Above is responsible for having properly set file_found and file_search.

	ResultType result = CONDITION_FALSE; // Default to "no iterations executed".
	Line *jump_to_line = nullptr;
	global_struct &g = *::g; // Primarily for performance in this case.

	g.mLoopFile = &lfs; // inner file-loop's file takes precedence over any outer file-loop's.
	// Other types of loops leave g.mLoopFile unchanged so that a file-loop can enclose some other type of
	// inner loop, and that inner loop will still have access to the outer loop's current file.

	for (; file_found; ++g.mLoopIteration)
	{
		PERFORMLOOP_EXECUTE_BODY
		PERFORMLOOP_EVALUATE_UNTIL
		
		// The macro above will "break", depending on result and jump_to_line.
		// Otherwise, the result of executing the body of the loop, above, was either OK
		// (the current iteration completed normally) or LOOP_CONTINUE (the current loop
		// iteration was cut short).  In both cases, just continue on through the loop.
		// But first do end-of-iteration steps:
		while ((file_found = FindNextFile(file_search, &lfs))
			&& FileIsFilteredOut(lfs, aFileLoopMode)); // Relies on short-circuit boolean order.
			// Above is a self-contained loop that keeps fetching files until there's no more files, or a file
			// is found that isn't filtered out.  It also sets file_found and lfs for use by the outer loop.
	} // for()

	// The script's loop is now over.
	if (file_search != INVALID_HANDLE_VALUE)
		FindClose(file_search);

	if (result != OK && result != CONDITION_FALSE || aJumpToLine)
		return result;

	// If aRecurseSubfolders is true, we now need to perform the loop's body for every subfolder to
	// search for more files and folders inside that match aFilePattern.  We can't do this in the
	// first loop, above, because it may have a restricted file-pattern such as *.txt and we want to
	// find and recurse into ALL folders:
	if (!aRecurseSubfolders) // No need to continue into the "recurse" section.
		return result;

	// Since above didn't return, this is a file-loop and recursion into sub-folders has been requested.
	// Append * to file_path so that we can retrieve all files and folders in the aFilePattern main dir.
	// We're only interested in the folders, but there's no special pattern that would filter out files.
	file_path_end[0] = '*'; // There's always room for this since it's shorter than lfs.pattern.
	file_path_end[1] = '\0';
	file_search = FindFirstFile(lfs.file_path, &lfs);
	if (file_search == INVALID_HANDLE_VALUE)
		return result; // Nothing more to do.
	// Otherwise, recurse into any subdirectories found inside this parent directory.

	LPTSTR short_path_end = lfs.short_path + short_path_length;
	size_t short_space_remaining = _countof(lfs.short_path) - short_path_length;

	do
	{
		if (!(lfs.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) // We only want directories (except "." and "..").
			|| lfs.cFileName[0] == '.' && (!lfs.cFileName[1]      // Relies on short-circuit boolean order.
				|| lfs.cFileName[1] == '.' && !lfs.cFileName[2])) //
			continue;

		size_t this_dir_length = _tcslen(lfs.cFileName);
		if (this_dir_length + 1 >= file_space_remaining)
			// This should be virtually impossible because:
			//  - Unicode builds allow 32767 chars, which is also the limit for most APIs.
			//  - ANSI builds allow MAX_PATH*2 (520), but FindFirstFileA() would fail for any
			//    file pattern longer than MAX_PATH, and cFileName itself is MAX_PATH chars max.
			continue;

		// Append the directory name\ to file_path.
		tmemcpy(file_path_end, lfs.cFileName, this_dir_length);
		file_path_end[this_dir_length] = '\\';
		file_path_end[this_dir_length + 1] = '\0';
		lfs.file_path_length = file_path_length + this_dir_length + 1;

		// Append the directory's short (8.3) name to short_path.
		LPTSTR short_name = lfs.cAlternateFileName;
		size_t short_name_length = _tcslen(short_name);
		if (!short_name_length)
		{
			short_name = lfs.cFileName; // See BIV_LoopFileName for comments about why cFileName is used.
			short_name_length = this_dir_length;
		}
		if (short_name_length + 1 >= short_space_remaining) // Should realistically never happen, but it's possible for an 8.3 name to be longer than the non-8.3 name.
			continue;
		tmemcpy(short_path_end, short_name, short_name_length);
		short_path_end[short_name_length] = '\\';
		short_path_end[short_name_length + 1] = '\0';
		lfs.short_path_length = short_path_length + short_name_length + 1;

		auto sub_result = PerformLoopFilePattern(aResultToken, aJumpToLine, aUntil, aFileLoopMode, aRecurseSubfolders, lfs);
		// Above returns LOOP_CONTINUE for cases like "continue 2" or "continue outer_loop", where the
		// target is not this Loop but a Loop which encloses it. In those cases we want below to return:
		if (sub_result != CONDITION_FALSE)
		{
			result = sub_result;
			if (result != OK) // i.e. result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL
				break;
			if (aJumpToLine)
				break; // Return and let our caller handle the jump.
		}
		// Otherwise, the recursive call performed no iterations, but result could already be
		// OK or CONDITION_FALSE depending on whether iterations were performed by this layer.
	} while (FindNextFile(file_search, &lfs));
	
	FindClose(file_search);
	return result;  // Return even LOOP_BREAK, since our caller can be either ExecUntil() or ourself.
}



ResultType Line::PerformLoopReg(ResultToken *aResultToken, Line *&aJumpToLine, Line *aUntil
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
		return CONDITION_FALSE;

	// Get the count of how many values and subkeys are contained in this parent key:
	DWORD count_subkeys;
	DWORD count_values;
	if (RegQueryInfoKey(hRegKey, NULL, NULL, NULL, &count_subkeys, NULL, NULL
		, &count_values, NULL, NULL, NULL, NULL) != ERROR_SUCCESS)
	{
		RegCloseKey(hRegKey);
		return CONDITION_FALSE;
	}

	ResultType result = CONDITION_FALSE;
	Line *jump_to_line = nullptr;
	DWORD i;
	global_struct &g = *::g; // Primarily for performance in this case.

	// See comments in PerformLoop() for details about this section.
	#define MAKE_SCRIPT_LOOP_PROCESS_THIS_ITEM \
	{\
		g.mLoopRegItem = &reg_item;\
		PERFORMLOOP_EXECUTE_BODY \
		PERFORMLOOP_EVALUATE_UNTIL \
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
		if (result != OK && result != CONDITION_FALSE || aJumpToLine)
		{
			RegCloseKey(hRegKey);
			return result;
		}
	}

	// If the loop is neither processing subfolders nor recursing into them, don't waste the performance
	// doing the next loop:
	if (!count_subkeys || (aFileLoopMode == FILE_LOOP_FILES_ONLY && !aRecurseSubfolders))
	{
		RegCloseKey(hRegKey);
		return result;
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
				auto sub_result = PerformLoopReg(aResultToken, aJumpToLine, aUntil
					, aFileLoopMode , aRecurseSubfolders, aRootKeyType, aRootKey, subkey_full_path);
				if (sub_result != CONDITION_FALSE)
				{
					result = sub_result;
					if (result != OK) // i.e. result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL
						break;
					if (aJumpToLine)
						break; // Return and let our caller handle the jump.
				}
				// Otherwise, the recursive call performed no iterations, but result could already be
				// OK or CONDITION_FALSE depending on whether iterations were performed by this layer.
			}
		}
		// else continue the loop in case some of the lower indexes can still be retrieved successfully.
		if (i == 0)  // Check this here due to it being an unsigned value that we don't want to go negative.
			break;
	}
	RegCloseKey(hRegKey);
	return result;
}



ResultType Line::PerformLoopParse(ResultToken *aResultToken, Line *&aJumpToLine, Line *aUntil)
{
	if (!*ARG1) // Since the input variable's contents are blank, the loop will execute zero times.
		return CONDITION_FALSE;

	// The following will be used to hold the parsed items.  It needs to have its own storage because
	// even though ARG1 might always be a writable memory area, we can't rely upon it being
	// persistent because it might reside in the deref buffer, in which case the other commands
	// in the loop's body would probably overwrite it.  Even if the ARG1's contents aren't in
	// the deref buffer, we still can't modify it (i.e. to temporarily terminate it and thus
	// bypass the need for malloc() below) because that might modify the variable contents, and
	// that variable may be referenced elsewhere in the body of the loop (which would result
	// in unexpected side-effects).  So, rather than have a limit of 64K or something (which
	// would limit this feature's usefulness for parsing a large list of filenames, for example),
	// it seems best to dynamically allocate a temporary buffer large enough to hold the
	// contents of ARG1 (the input variable).  Update: Since these loops tend to be enclosed
	// by file-read loops, and thus may be called thousands of times in a short period,
	// it should help average performance to use the stack for small vars rather than
	// constantly doing malloc() and free(), which are much higher overhead and probably
	// cause memory fragmentation (especially with thousands of calls):
	size_t space_needed = ArgLength(1) + 1;  // +1 for the zero terminator.
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
			return MemoryError();
		stack_buf = NULL; // For comparison purposes later below.
	}
	_tcscpy(buf, ARG1); // Make the copy.

	// Make a copy of ARG2 and ARG3 in case either one's contents are in the deref buffer, which would
	// probably be overwritten by the commands in the script loop's body:
	TCHAR delimiters[512], omit_list[512];
	tcslcpy(delimiters, ARG2, _countof(delimiters));
	tcslcpy(omit_list, ARG3, _countof(omit_list));

	ResultType result = CONDITION_FALSE;
	Line *jump_to_line = nullptr;
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

		PERFORMLOOP_EXECUTE_BODY
		PERFORMLOOP_EVALUATE_UNTIL

		if (!saved_char) // The last item in the list has just been processed, so the loop is done.
			break;
		*field_end = saved_char;  // Undo the temporary termination, in case the list of delimiters is blank.
		field = *delimiters ? field_end + 1 : field_end;  // Move on to the next field.
		++g.mLoopIteration;
	}
	FREE_PARSE_MEMORY;
	return result;
}



ResultType Line::PerformLoopParseCSV(ResultToken *aResultToken, Line *&aJumpToLine, Line *aUntil)
// This function is similar to PerformLoopParse() so the two should be maintained together.
// See PerformLoopParse() for comments about the below (comments have been mostly stripped
// from this function).
{
	if (!*ARG1) // Since the input variable's contents are blank, the loop will execute zero times.
		return CONDITION_FALSE;

	// See comments in PerformLoopParse() for details.
	size_t space_needed = ArgLength(1) + 1;  // +1 for the zero terminator.
	LPTSTR stack_buf, buf;
	if (space_needed <= LOOP_PARSE_BUF_SIZE)
	{
		stack_buf = (LPTSTR)talloca(space_needed); // Helps performance.  See comments above.
		buf = stack_buf;
	}
	else
	{
		if (   !(buf = tmalloc(space_needed))   )
			return MemoryError();
		stack_buf = NULL; // For comparison purposes later below.
	}
	_tcscpy(buf, ARG1); // Make the copy.

	TCHAR omit_list[512];
	tcslcpy(omit_list, ARG3, _countof(omit_list));

	ResultType result = CONDITION_FALSE;
	Line *jump_to_line = nullptr;
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

		PERFORMLOOP_EXECUTE_BODY
		PERFORMLOOP_EVALUATE_UNTIL

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
	return result;
}



ResultType Line::PerformLoopReadFile(ResultToken *aResultToken, Line *&aJumpToLine, Line *aUntil
	, LPTSTR aReadFileName, LPTSTR aWriteFileName)
{
	TextFile tfile;
	bool file_is_open = tfile.Open(aReadFileName, DEFAULT_READ_FLAGS, g->Encoding & CP_AHKCP);
	if (!file_is_open)
	{
		// Failed to open the input file.  If an ELSE is present, executing it if the file wasn't found
		// seems much more useful than executing ELSE only for empty files; but if there's no ELSE, it's
		// probably best to throw an error.
		DWORD error = GetLastError();
		if (  !(mRelatedLine->mActionType == ACT_ELSE && (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND))  )
			return g_script.Win32Error(error);
		// Otherwise, execute the ELSE below.
	}

	// Make a persistent copy in case aWriteFileName's contents are in the deref buffer:
	if (  !(aWriteFileName = _tcsdup(aWriteFileName))  )
		return MemoryError();

	LoopReadFileStruct loop_info { aWriteFileName };
	size_t line_length;
	ResultType result = CONDITION_FALSE;
	Line *jump_to_line = nullptr;
	global_struct &g = *::g; // Primarily for performance in this case.
	g.mLoopReadFile = &loop_info;

	if (file_is_open)
	for (;; ++g.mLoopIteration)
	{ 
		if (  !(line_length = tfile.ReadLine(loop_info.mCurrentLine, _countof(loop_info.mCurrentLine) - 1))  ) // -1 to ensure there's room for a null-terminator.
			break;
		if (loop_info.mCurrentLine[line_length - 1] == '\n') // Remove end-of-line character.
			--line_length;
		loop_info.mCurrentLine[line_length] = '\0';

		PERFORMLOOP_EXECUTE_BODY
		PERFORMLOOP_EVALUATE_UNTIL
	}

	if (result == CONDITION_FALSE)
	{
		// No lines were read, so execute this loop's ELSE if present.  It's done here rather than
		// in ExecUntil so that FileAppend can be used within the ELSE to write to mWriteFile.
		if (mRelatedLine->mActionType == ACT_ELSE)
			result = mRelatedLine->mNextLine->ExecUntil(ONLY_ONE_LINE, aResultToken, &aJumpToLine);
		else
			result = OK; // Never return CONDITION_FALSE, otherwise ExecUntil() will execute the ELSE a second time.
	}

	if (loop_info.mWriteFile)
	{
		loop_info.mWriteFile->Close();
		delete loop_info.mWriteFile;
	}

	return result;
}



ResultType Line::PerformAssign()
{
	// Note: This line's args have not yet been dereferenced in this case (i.e. ExpandArgs() hasn't been called).
	// Currently, ACT_ASSIGNEXPR can occur even when mArg[1].is_expression==false, such as things like var:=5
	// and var:="string".  Search on "is_expression = " to find such cases in the script-loading/parsing routines.
	
	if (!mArg[1].is_expression)
	{
		ASSERT(mArg[1].postfix);
		// Examples of assignments this covers:
		//		x := 123
		//		x := 1.0
		//		x := "quoted literal string"
		//		x := normal_var  ; but not VAR_VIRTUAL or VAR_CONSTANT
		if (mArg[1].postfix->symbol == SYM_VAR && mArg[1].postfix->var->IsUninitialized()
			&& mArg[1].postfix->var_usage == VARREF_READ)
			return g_script.VarUnsetError(mArg[1].postfix->var); // !is_expression implies VAR_NORMAL, so InitializeConstant() isn't needed here.
		Var *output_var = VAR(mArg[0]);
		return output_var->Assign(*mArg[1].postfix);
	}

	return ExpandArgs(); // ExpandExpression() will also take care of the assignment (for performance).
}



ResultType Script::DerefInclude(LPTSTR &aOutput, LPTSTR aBuf)
// For #Include, #IncludeAgain and #DllLoad.
// Based on Line::Deref above, but with a few differences for backward-compatibility:
//  1) Percent signs that aren't part of a valid deref are not omitted.
//  2) Escape sequences aren't recognized (`; is handled elsewhere).
//  3) It is restricted to built-in vars to reduce the risk of breaking any scripts
//     that use percent sign literally in a filename.  Most other vars are empty anyway.
{
	aOutput = nullptr; // Set default.

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
					Var *var = FindGlobalVar(cp + 1, var_name_length);
					if (var && var->Type() == VAR_VIRTUAL)
					{
						if (which_pass) // 2nd pass
						{
							size_t var_length = var->CharLength();
							tmemcpy(dest, var->Contents(), var_length);
							var->Free();
							dest += var_length;
						}
						else
						{
							if (!var->PopulateVirtualVar())
								return FAIL;
							expanded_length += var->CharLength();
						}
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
				continue; // ACT_LISTLINES and other things might rely on "continue" instead of halting the loop here.
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



LPTSTR Line::ToText(LPTSTR aBuf, int aBufSize, bool aCRLF, DWORD aElapsed, bool aLineWasResumed, bool aLineNumber) // aBufSize should be an int to preserve negatives from caller (caller relies on this).
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

	if (aLineNumber)
		aBuf += sntprintf(aBuf, aBufSize, _T("%03u: "), mLineNumber);
	if (aLineWasResumed)
		aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("STILL WAITING (%0.2f): "), (float)aElapsed / 1000.0);

	if (mActionType == ACT_CASE)
	{
		if (mArgc)
			aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("%s %s:"), g_act[mActionType].Name
				, mArg[0].text);
		else
			aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("Default:"));
	}
	else
	{
		int i = 0;
		if (mActionType == ACT_ASSIGNEXPR)
			aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("%s := "), mArg[i++].text);
		else if (mActionType != ACT_EXPRESSION)
			aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("%s"), g_act[mActionType].Name);
		for (; i < mArgc; ++i)
		{
			bool quote = mArg[i].type == ARG_TYPE_NORMAL
				&& !mArg[i].is_expression && mArg[i].postfix && mArg[i].postfix->symbol == SYM_STRING;
			LPCTSTR delim = mActionType <= ACT_EXPRESSION ? _T("") : i == 0 ? _T(" ")
				: (i + 1 == mArgc && mActionType == ACT_FOR) ? _T(" in ")
				: _T(", ");
			// This method is a little more efficient than using snprintfcat().
			aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, quote ? _T("%s\"%s\"") : _T("%s%s"), delim, mArg[i].text);
		}
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



void ToggleSuspendState()
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



bif_impl FResult Suspend(optl<int> aMode)
{
	auto toggle = Line::Convert10Toggle(aMode);
	if (toggle == TOGGLE_INVALID)
		return FR_E_ARG(0);
	if (toggle >= TOGGLE // i.e. TOGGLE or NEUTRAL (omitted)
		|| ((toggle == TOGGLED_ON) != g_IsSuspended))
		ToggleSuspendState();
	return OK;
}



void PauseUnderlyingThread(bool aTrueForPauseFalseForUnpause)
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


bif_impl FResult Pause(optl<int> aNewState)
{
	auto toggle = Line::Convert10Toggle(aNewState);
	switch (toggle)
	{
	case NEUTRAL:
		PauseCurrentThread();
		return OK;
	case TOGGLE:
		// Update for v2: "Pause -1" is more useful if it always applies to the thread immediately beneath
		// the current thread, since pausing the current thread would prevent a hotkey from unpausing itself
		// by default (there's no longer a special case allowing Pause hotkeys a second thread).
		if (g > g_array && g[-1].IsPaused) // Checking g>g_array avoids any chance of underflow, which might otherwise happen if this is called by the AutoExec section or a threadless callback running in thread #0.
			toggle = TOGGLED_OFF;
		// Fall through:
	case TOGGLED_ON:
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
		PauseUnderlyingThread(toggle != TOGGLED_OFF); // i.e. pause if ON or fell through from TOGGLE.
		return OK;
	default: // TOGGLE_INVALID or some other disallowed value.
		return FR_E_ARG(0);
	}
}


void PauseCurrentThread()
{
	// Pause the current subroutine (which by definition isn't paused since it had to be  active
	// to call us).  It seems best not to attempt to change the Hotkey mRunAgainAfterFinished
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
	auto &g = *::g;
	g.IsPaused = true;
	++g_nPausedThreads; // For this purpose the idle thread is counted as a paused thread.
	g_script.UpdateTrayIcon();
	// Don't return until the script is unpaused.  This is done for two reasons:
	// 1) Pause() can be called as part of an expression, in which case it seems more intuitive for
	//    the script to pause immediately rather than after evaluating more of the expression.
	// 2) If `return Pause()` is used, the thread might end before checking g.IsPaused, in which
	//    case g_nPausedThreads would not be adjusted and timers would forever be disabled.
	while (g.IsPaused)
		MsgSleep(INTERVAL_UNSPECIFIED);
}



ResultType Script::PreprocessLocalVars(FuncList &aFuncs)
{
	for (int i = 0; i < aFuncs.mCount; ++i)
	{
		ASSERT(dynamic_cast<UserFunc*>(aFuncs.mItem[i]));
		auto &func = *(UserFunc *)aFuncs.mItem[i];
		if (!PreprocessLocalVars(func))
			return FAIL;
		// Nested functions will be preparsed next, due to the fact that they immediately
		// follow the outer function in aFuncs.
	}
	return OK;
}



ResultType Script::PreprocessLocalVars(UserFunc &aFunc)
{
	Var **vars = aFunc.mVars.mItem;
	int var_count = aFunc.mVars.mCount;
	int closure_count = 0, non_closure_count = 0;
	for (int v = 0; v < var_count; ++v)
	{
		Var &var = *vars[v];
		if (var.Type() == VAR_CONSTANT)
		{
			// Currently only nested functions create local constants, but this could be
			// a local constant or an upvar alias for a non-local constant.
			ASSERT(var.HasObject() && dynamic_cast<UserFunc *>(var.Object()));
			auto &func = *(UserFunc *)var.Object();
			if (!func.mUpVarCount)
			{
				// This function doesn't need to be a closure, so convert it to static.
				int insert_pos;
				aFunc.mStaticVars.Find(var.mName, &insert_pos);
				aFunc.mStaticVars.Insert(&var, insert_pos);
				var.Scope() |= VAR_LOCAL_STATIC;
				++non_closure_count;
				if (var.Scope() & VAR_DOWNVAR)
				{
					var.Scope() &= ~VAR_DOWNVAR;
					--aFunc.mDownVarCount;
				}
			}
			else if (!var.IsAlias()) // At this stage only upvars are aliases.
			{
				// This is a closure of aFunc and not an alias for an outer function's
				// nested function constant.
				++closure_count;
			}
		}
	}
	if (non_closure_count) // One or more static vars are to be removed from mVars.
	{
		int dst = 0;
		for (int src = 0; src < var_count; ++src)
			if (!vars[src]->IsStatic())
				vars[dst++] = vars[src];
		aFunc.mVars.mCount = var_count = dst;
	}
	if (closure_count)
	{
		aFunc.mClosure = SimpleHeap::Alloc<ClosureInfo>(closure_count);
		for (int v = 0; v < var_count; ++v)
		{
			Var &var = *vars[v];
			if (!var.IsAlias() && var.Type() == VAR_CONSTANT)
			{
				ASSERT(var.IsObject() && dynamic_cast<UserFunc *>(var.Object()));
				auto &ci = aFunc.mClosure[aFunc.mClosureCount++];
				ci.var = &var;
				ci.func = (UserFunc *)var.Object();
			}
		}
		ASSERT(aFunc.mClosureCount == closure_count);
	}
	if (aFunc.mUpVarCount)
	{
		// Upvars are vars local to an outer function which are referenced by aFunc.
		// They might be local to aFunc.mOuterFunc or one further out.  In the latter
		// case, aFunc.mOuterFunc contains a downvar which is also one of its upvars.
		aFunc.mUpVar = SimpleHeap::Alloc<Var*>(aFunc.mUpVarCount);
		aFunc.mUpVarIndex = SimpleHeap::Alloc<int>(aFunc.mUpVarCount);

		auto &outer = *aFunc.mOuterFunc; // Always non-null when aFunc.mUpVarCount > 0.
		int upvar_count = 0;
		for (int v = 0; v < var_count; ++v)
		{
			Var &var = *vars[v];
			if (!var.IsAlias()) // At this stage only upvars are aliases.
				continue;

			Var *downvar = var.GetAliasFor(); // FindUpVar() set var to be an alias of the corresponding downvar.
			int d;
			for (d = 0; ; ++d)
			{
				ASSERT(d < outer.mDownVarCount);
				if (outer.mDownVar[d] == downvar)
					break;
			}

			ASSERT(upvar_count < aFunc.mUpVarCount);
			aFunc.mUpVar[upvar_count] = &var;
			aFunc.mUpVarIndex[upvar_count] = d;
			++upvar_count;
		}
		ASSERT(upvar_count == aFunc.mUpVarCount);
	}
	if (aFunc.mDownVarCount)
	{
		// Downvars are vars local to aFunc which are referenced by a nested function.
		aFunc.mDownVar = SimpleHeap::Alloc<Var*>(aFunc.mDownVarCount);
		
		int downvar_count = 0;
		for (int v = 0; v < var_count; ++v)
		{
			if (vars[v]->Scope() & VAR_DOWNVAR)
			{
				ASSERT(downvar_count < aFunc.mDownVarCount);
				aFunc.mDownVar[downvar_count++] = vars[v];
			}
		}
		ASSERT(downvar_count == aFunc.mDownVarCount);
	}
	return OK;
}



ResultType Script::PreparseVarRefs()
{
	// AddLine() and ExpressionToPostfix() currently leave validation of output variables
	// and lvalues to this function, to reduce code size.  Search for "VarIsReadOnlyError".
	for (Line *line = mFirstLine; line; line = line->mNextLine)
	{
		switch (line->mActionType)
		{ // Establish context for FindOrAddVar:
		case ACT_BLOCK_BEGIN: if (line->mAttribute) g->CurrentFunc = (UserFunc *)line->mAttribute; break;
		case ACT_BLOCK_END: if (line->mAttribute) g->CurrentFunc = g->CurrentFunc->mOuterFunc; break;
		}
		
		mCurrLine = line; // For error-reporting.

		for (int a = 0; a < line->mArgc; ++a)
		{
			ArgStruct &arg = line->mArg[a];
			if (!arg.is_expression)
				continue;
			for (ExprTokenType *token = arg.postfix; token->symbol != SYM_INVALID; ++token)
			{
				if (token->symbol != SYM_VAR // Not a var.
					|| VARREF_IS_WRITE(token->var_usage)) // Already resolved by ExpressionToPostfix.
					continue;
				if (token->var_deref->type == DT_FUNCREF)
				{
					token->var = token->var_deref->var;
					continue;
				}
				if (  !(token->var = FindOrAddVar(token->var_deref->marker, token->var_deref->length, FINDVAR_FOR_READ))  )
					return FAIL;
				if (token->var->IsAlias()) // Upvar.
					continue;
				switch (token->var->Type())
				{
				case VAR_CONSTANT:
					// Resolve constants to their values, except in cases like IsSet(SomeClass), or when
					// the constant is a closure (which may change each time the outer function is called).
					if (!token->var->IsLocal() && VARREF_IS_READ(token->var_usage) && !token->var->IsUninitialized())
						token->var->ToToken(*token);
					continue;
				case VAR_VIRTUAL:
					if (VARREF_IS_READ(token->var_usage))
						++arg.max_alloc; // Reserve a to_free[] slot for it in ExpandExpression().
					break;
				default:
					// Suppress any VarUnset warnings for IsSet(var) so that it can be used to determine if an
					// optionally-set global variable has been set, such as a class in an optional #Include.
					// MarkAssignedSomewhere() isn't used for this because PreprocessLocalVars() hasn't been
					// called yet, and it relies on the attribute being accurate.
					if (token->var_usage == VARREF_ISSET)
						token->var->MarkAlreadyWarned();
					// It's too early to show VarUnset warnings, since not all VARREF_ISSET references have been marked.
					// The effect of IsSet(v) for suppressing the warning shouldn't be positional since it's feasible for
					// a check in one function to guard evaluation of a VARREF_READ in some other function.
				}
			}
			if (arg.type == ARG_TYPE_INPUT_VAR)
			{
				if (arg.postfix->symbol != SYM_VAR || arg.postfix->var->Type() != VAR_NORMAL)
				{
					// Can't be ARG_TYPE_INPUT_VAR after all, as VAR_VIRTUAL and VAR_CONSTANT require ExpandExpression
					// (unless it's a constant which was converted to SYM_OBJECT above).
					arg.type = ARG_TYPE_NORMAL;
				}
				else
				{
					// This arg can be optimized by avoiding ExpandExpression.
					arg.deref = (DerefType *)arg.postfix->var;
					arg.is_expression = false;
				}
			}
		}
	}
	return OK;
}



LPTSTR ListVarsHelper(LPTSTR aBuf, int aBufSize, LPTSTR aBuf_orig, VarList &aVars)
{
	for (int i = 0; i < aVars.mCount; ++i)
		if (aVars.mItem[i]->Type() == VAR_NORMAL) // Don't bother showing VAR_CONSTANT; ToText() doesn't support VAR_VIRTUAL.
			aBuf = aVars.mItem[i]->ToText(aBuf, BUF_SPACE_REMAINING, true);
	return aBuf;
}

LPTSTR Script::ListVars(LPTSTR aBuf, int aBufSize) // aBufSize should be an int to preserve negatives from caller (caller relies on this).
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Translates this script's list of variables into text equivalent, putting the result
// into aBuf and returning the position in aBuf of its new string terminator.
{
	LPTSTR aBuf_orig = aBuf;
	auto current_func = g->CurrentFunc;
	if (current_func)
	{
		// This definition might help compiler string pooling by ensuring it stays the same for all usages:
		#define LIST_VARS_UNDERLINE _T("\r\n--------------------------------------------------\r\n")
		if (current_func->mVars.mCount || !current_func->mStaticVars.mCount)
		{
			aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("Local Variables for %s()%s")
				, current_func->mName, LIST_VARS_UNDERLINE);
			aBuf = ListVarsHelper(aBuf, aBufSize, aBuf_orig, current_func->mVars);
		}
		do {
			if (current_func->mStaticVars.mCount)
			{
				aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("%sStatic Variables for %s()%s")
					, (aBuf > aBuf_orig) ? _T("\r\n\r\n") : _T(""), current_func->mName, LIST_VARS_UNDERLINE);
				aBuf = ListVarsHelper(aBuf, aBufSize, aBuf_orig, current_func->mStaticVars);
			}
		} while (current_func = current_func->mOuterFunc);
	}
	aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("%sGlobal Variables (alphabetical)%s")
		, (aBuf > aBuf_orig) ? _T("\r\n\r\n") : _T(""), LIST_VARS_UNDERLINE);
	aBuf = ListVarsHelper(aBuf, aBufSize, aBuf_orig, mVars);
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
			sntprintfcat(timer_list, _countof(timer_list) - 3, _T("%s "), timer->mCallback->Name()); // Allow room for "..."
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
		: _T("\r\nKey History has been disabled via KeyHistory(0)."));
}



ResultType Script::ActionExec(LPCTSTR aAction, LPCTSTR aParams, LPCTSTR aWorkingDir, bool aDisplayErrors
	, LPCTSTR aRunShowMode, HANDLE *aProcess, bool aUpdateLastError, bool aUseRunAs)
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

	// Launching nothing is always a success:
	if (!aAction || !*aAction) return OK;

	if (aWorkingDir)
	{
		if (*aWorkingDir)
		{
			// aWorkingDir is validated for two main reasons:
			//  1) To make the cause of failure much clearer when both CreateProcess and ShellExecuteEx fail.
			//  2) To prevent ShellExecuteEx from wrongly executing a file:
			//     a) with aAction relative to the wrong working directory; and/or
			//     b) with the child process having the wrong working directory.
			auto attrib = GetFileAttributes(aWorkingDir);
			if (attrib == INVALID_FILE_ATTRIBUTES || !(attrib & FILE_ATTRIBUTE_DIRECTORY))
				return aDisplayErrors ? RuntimeError(ERR_PARAM2_INVALID, aWorkingDir) : FAIL;
		}
		else
			// Set it to NULL because CreateProcess() won't work if it's the empty string.
			aWorkingDir = NULL;
	}

	#define IS_VERB(str) (   !_tcsicmp(str, _T("find")) || !_tcsicmp(str, _T("explore")) || !_tcsicmp(str, _T("open"))\
		|| !_tcsicmp(str, _T("Edit")) || !_tcsicmp(str, _T("print")) || !_tcsicmp(str, _T("properties"))   )

	// Set default items to be run by ShellExecute().  These are also used by the error
	// reporting at the end, which is why they're initialized even if CreateProcess() works
	// and there's no need to use ShellExecute():
	LPCTSTR shell_verb = NULL;
	LPCTSTR shell_action = aAction;
	LPCTSTR shell_params = NULL;
	
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
		LPCTSTR phrase_end = StrChrAny(shell_action, _T(" \t")); // Find space or tab.
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
			return RuntimeError(_T("System verbs unsupported with RunAs."));
		return FAIL;
	}
	
	size_t action_length = _tcslen(shell_action); // shell_action == aAction if shell_verb == NULL.
	if (action_length >= LINE_SIZE) // Max length supported by CreateProcess() is 32 KB. But there hasn't been any demand to go above 16 KB, so seems little need to support it (plus it reduces risk of stack overflow).
	{
        if (aDisplayErrors)
			return RuntimeError(_T("String too long.")); // Short msg since so rare.
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
			command_line = talloca(action_length + _tcslen(aParams) + 10); // +10 to allow room for quotes, space, terminator, and any extra chars that might get added in the future.
			_stprintf(command_line, _T("\"%s\" %s"), aAction, aParams);
		}
		else // We're running the original action from caller.
		{
			command_line = talloca(action_length + 1);
        	_tcscpy(command_line, aAction); // CreateProcessW() requires modifiable string.  Although non-W version is used now, it feels safer to make it modifiable anyway.
		}

		if (use_runas)
		{
			if (!DoRunAs(command_line, aWorkingDir, aDisplayErrors, si.wShowWindow  // wShowWindow (min/max/hide).
				, pi, success, hprocess, last_error)) // These are output parameters it will set for us.
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
			hprocess = sei.hProcess;
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
			return RuntimeError(error_text, system_error_text);
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
