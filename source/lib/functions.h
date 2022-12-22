
#ifndef MD_WINTITLE_ARGS
#define MD_WINTITLE_ARGS (In_Opt, Variant, WinTitle), (In_Opt, String, WinText), (In_Opt, String, ExcludeTitle), (In_Opt, String, ExcludeText)
#define MD_CONTROL_ARGS (In, Variant, Control), MD_WINTITLE_ARGS
#define MD_CONTROL_ARGS_OPT (In_Opt, Variant, Control), MD_WINTITLE_ARGS
#endif

md_func_x(BlockInput, ScriptBlockInput, FResult, (In, String, Mode))

#ifdef ENABLE_REGISTERCALLBACK
md_func(CallbackCreate, (In, Object, Function), (In_Opt, String, Options), (In_Opt, Int32, ParamCount), (Ret, UIntPtr, RetVal))
md_func(CallbackFree, (In, UIntPtr, Callback))
#endif

md_func(ClipWait, (In_Opt, Float64, Timeout), (In_Opt, Int32, AnyType), (Ret, Bool32, RetVal))

md_func(ControlAddItem, (In, String, Value), MD_CONTROL_ARGS, (Ret, IntPtr, Index))
md_func(ControlChooseIndex, (In, IntPtr, Index), MD_CONTROL_ARGS)
md_func(ControlChooseString, (In, String, Value), MD_CONTROL_ARGS, (Ret, IntPtr, Index))
md_func(ControlClick, (In_Opt, Variant, Control), (In_Opt, Variant, WinTitle), (In_Opt, String, WinText), (In_Opt, String, Button), (In_Opt, Int32, Count), (In_Opt, String, Options), (In_Opt, String, ExcludeTitle), (In_Opt, String, ExcludeText))
md_func(ControlDeleteItem, (In, IntPtr, Index), MD_CONTROL_ARGS)
md_func(ControlFindItem, (In, String, Item), MD_CONTROL_ARGS, (Ret, IntPtr, Index))
md_func(ControlFocus, MD_CONTROL_ARGS)
md_func(ControlGetChecked, MD_CONTROL_ARGS, (Ret, Bool32, RetVal))
md_func(ControlGetChoice, MD_CONTROL_ARGS, (Ret, String, RetVal))
md_func(ControlGetClassNN, MD_CONTROL_ARGS, (Ret, String, ClassNN))
md_func(ControlGetEnabled, MD_CONTROL_ARGS, (Ret, Bool32, RetVal))
md_func(ControlGetExStyle, MD_CONTROL_ARGS, (Ret, UInt32, RetVal))
md_func(ControlGetFocus, MD_WINTITLE_ARGS, (Ret, UInt32, Hwnd))
md_func(ControlGetHwnd, MD_CONTROL_ARGS, (Ret, UInt32, Hwnd))
md_func(ControlGetIndex, MD_CONTROL_ARGS, (Ret, IntPtr, Index))
md_func(ControlGetItems, MD_CONTROL_ARGS, (Ret, Object, StringArray))
md_func(ControlGetPos, (Out_Opt, Int32, X), (Out_Opt, Int32, Y), (Out_Opt, Int32, Width), (Out_Opt, Int32, Height), MD_CONTROL_ARGS)
md_func(ControlGetStyle, MD_CONTROL_ARGS, (Ret, UInt32, RetVal))
md_func(ControlGetText, MD_CONTROL_ARGS, (Ret, String, Text))
md_func(ControlGetVisible, MD_CONTROL_ARGS, (Ret, Bool32, RetVal))
md_func(ControlHide, MD_CONTROL_ARGS)
md_func(ControlHideDropDown, MD_CONTROL_ARGS)
md_func(ControlMove, (In_Opt, Int32, X), (In_Opt, Int32, Y), (In_Opt, Int32, Width), (In_Opt, Int32, Height), MD_CONTROL_ARGS)
md_func(ControlSend, (In, String, Keys), MD_CONTROL_ARGS_OPT)
md_func(ControlSendText, (In, String, Keys), MD_CONTROL_ARGS_OPT)
md_func(ControlSetChecked, (In, Int32, Value), MD_CONTROL_ARGS)
md_func(ControlSetEnabled, (In, Int32, Value), MD_CONTROL_ARGS)
md_func(ControlSetExStyle, (In, String, Value), MD_CONTROL_ARGS)
md_func(ControlSetStyle, (In, String, Value), MD_CONTROL_ARGS)
md_func(ControlSetText, (In, String, NewText), MD_CONTROL_ARGS)
md_func(ControlShow, MD_CONTROL_ARGS)
md_func(ControlShowDropDown, MD_CONTROL_ARGS)

md_func(CoordMode, (In, String, TargetType), (In_Opt, String, RelativeTo), (Ret, String, RetVal))
md_func_v(Critical, (In_Opt, String, OnOffNumber), (Ret, Int32, RetVal))

md_func(DateAdd, (In, String, DateTime), (In, Float64, Time), (In, String, TimeUnits), (Ret, String, RetVal))
md_func(DateDiff, (In, String, DateTime1), (In, String, DateTime2), (In, String, TimeUnits), (Ret, Int64, RetVal))

md_func_v(DetectHiddenText, (In, Bool32, Mode), (Ret, Bool32, RetVal))
md_func_v(DetectHiddenWindows, (In, Bool32, Mode), (Ret, Bool32, RetVal))

md_func(DirCopy, (In, String, Source), (In, String, Dest), (In_Opt, Int32, Overwrite))
md_func(DirCreate, (In, String, Path))
md_func(DirDelete, (In, String, Path), (In_Opt, Bool32, Recurse))
md_func_v(DirExist, (In, String, Pattern), (Ret, String, RetVal))
md_func(DirMove, (In, String, Source), (In, String, Dest), (In_Opt, String, Flag))
md_func(DirSelect, (In_Opt, String, StartingFolder), (In_Opt, Int32, Options), (In_Opt, String, Prompt), (Ret, String, RetVal))

md_func(Download, (In, String, URL), (In, String, Path))

md_func(DriveEject, (In_Opt, String, Drive))
md_func(DriveGetCapacity, (In, String, Path), (Ret, Int64, RetVal))
md_func(DriveGetFilesystem, (In, String, Drive), (Ret, String, RetVal))
md_func(DriveGetLabel, (In, String, Drive), (Ret, String, RetVal))
md_func(DriveGetList, (In_Opt, String, Type), (Ret, String, RetVal))
md_func(DriveGetSerial, (In, String, Drive), (Ret, Int64, RetVal))
md_func(DriveGetSpaceFree, (In, String, Path), (Ret, Int64, RetVal))
md_func(DriveGetStatus, (In, String, Path), (Ret, String, RetVal))
md_func(DriveGetStatusCD, (In_Opt, String, Drive), (Ret, String, RetVal))
md_func(DriveGetType, (In, String, Path), (Ret, String, RetVal))
md_func(DriveLock, (In, String, Drive))
md_func(DriveRetract, (In_Opt, String, Drive))
md_func(DriveSetLabel, (In, String, Drive), (In_Opt, String, Label))
md_func(DriveUnlock, (In, String, Drive))

md_func_v(Edit, (In_Opt, String, Filename))

md_func(EditGetCurrentCol, MD_CONTROL_ARGS, (Ret, UInt32, Index))
md_func(EditGetCurrentLine, MD_CONTROL_ARGS, (Ret, UIntPtr, Index))
md_func(EditGetLine, (In, IntPtr, Index), MD_CONTROL_ARGS, (Ret, String, RetVal))
md_func(EditGetLineCount, MD_CONTROL_ARGS, (Ret, UIntPtr, Index))
md_func(EditGetSelectedText, MD_CONTROL_ARGS, (Ret, String, RetVal))
md_func(EditPaste, (In, String, Value), MD_CONTROL_ARGS)

md_func(EnvGet, (In, String, VarName), (Ret, String, RetVal))
md_func(EnvSet, (In, String, VarName), (In_Opt, String, Value))

md_func_x(Exit, Exit, ResultType, (In_Opt, Int32, ExitCode))
md_func_x(ExitApp, ExitApp, ResultType, (In_Opt, Int32, ExitCode))

md_func(FileAppend, (In, Variant, Value), (In_Opt, String, Path), (In_Opt, String, Options))
md_func(FileCopy, (In, String, Source), (In, String, Dest), (In_Opt, Int32, Overwrite))
md_func(FileCreateShortcut, (In, String, Target), (In, String, LinkFile), (In_Opt, String, WorkingDir),
	(In_Opt, String, Args), (In_Opt, String, Description), (In_Opt, String, IconFile),
	(In_Opt, String, ShortcutKey), (In_Opt, Int32, IconNumber), (In_Opt, Int32, RunState))
md_func(FileDelete, (In, String, Pattern))
md_func(FileEncoding, (In, String, Encoding), (Ret, Variant, RetVal))
md_func_v(FileExist, (In, String, Pattern), (Ret, String, RetVal))
md_func(FileGetAttrib, (In_Opt, String, Path), (Ret, String, RetVal))
md_func(FileGetShortcut, (In, String, LinkFile), (Out_Opt, String, Target), (Out_Opt, String, WorkingDir),
	(Out_Opt, String, Args), (Out_Opt, String, Description), (Out_Opt, String, IconFile),
	(Out_Opt, Variant, IconNum), (Out_Opt, Int32, RunState))
md_func(FileGetSize, (In_Opt, String, Path), (In_Opt, String, Units), (Ret, Int64, RetVal))
md_func(FileGetTime, (In_Opt, String, Path), (In_Opt, String, WhichTime), (Ret, String, RetVal))
md_func(FileGetVersion, (In_Opt, String, Path), (Ret, String, RetVal))
md_func(FileInstall, (In, String, Source), (In, String, Dest), (In_Opt, Int32, Overwrite))
md_func(FileMove, (In, String, Source), (In, String, Dest), (In_Opt, Int32, Overwrite))
md_func(FileRead, (In, String, Path), (In_Opt, String, Options), (Ret, Variant, RetVal))
md_func(FileRecycle, (In, String, Pattern))
md_func(FileRecycleEmpty, (In_Opt, String, Drive))
md_func(FileSelect, (In_Opt, String, Options), (In_Opt, String, RootDirFileName), (In_Opt, String, Title), (In_Opt, String, Filter), (Ret, Variant, RetVal))
md_func(FileSetAttrib, (In, String, Attributes), (In_Opt, String, Pattern), (In_Opt, String, Mode))
md_func(FileSetTime, (In_Opt, String, YYYYMMDD), (In_Opt, String, Pattern), (In_Opt, String, WhichTime), (In_Opt, String, Mode))

md_func_v(GetKeyName, (In, String, KeyName), (Ret, String, RetVal))
md_func_x(GetKeySC, GetKeySC, Int32, (In, String, KeyName))
md_func(GetKeyState, (In, String, KeyName), (In_Opt, String, Mode), (Ret, Variant, RetVal))
md_func_x(GetKeyVK, GetKeyVK, Int32, (In, String, KeyName))

md_func(GroupActivate, (In, String, GroupName), (In_Opt, String, Mode), (Ret, UInt32, RetVal))
md_func(GroupAdd, (In, String, GroupName), (In_Opt, String, WinTitle), (In_Opt, String, WinText), (In_Opt, String, ExcludeTitle), (In_Opt, String, ExcludeText))
md_func(GroupClose, (In, String, GroupName), (In_Opt, String, Mode))
md_func(GroupDeactivate, (In, String, GroupName), (In_Opt, String, Mode))

md_func_v(GuiCtrlFromHwnd, (In, UInt32, Hwnd), (Ret, Object, Gui))
md_func_v(GuiFromHwnd, (In, UInt32, Hwnd), (In_Opt, Bool32, Recurse), (Ret, Object, Gui))

md_func(HotIf, (In_Opt, Variant, Criterion))
md_func(HotIfWinActive, (In_Opt, String, WinTitle), (In_Opt, String, WinText))
md_func(HotIfWinExist, (In_Opt, String, WinTitle), (In_Opt, String, WinText))
md_func(HotIfWinNotActive, (In_Opt, String, WinTitle), (In_Opt, String, WinText))
md_func(HotIfWinNotExist, (In_Opt, String, WinTitle), (In_Opt, String, WinText))
md_func_x(Hotkey, BIF_Hotkey, FResult, (In, String, KeyName), (In_Opt, Variant, Action), (In_Opt, String, Options))
md_func_x(Hotstring, BIF_Hotstring, FResult, (In, String, String), (In_Opt, Variant, Replacement), (In_Opt, String, OnOffToggle), (Ret, Variant, RetVal))

md_func(IL_Add, (In, UIntPtr, ImageList), (In, String, Filename), (In_Opt, Int32, IconNumber), (In_Opt, Bool32, ResizeNonIcon), (Ret, Int32, Index))
md_func_x(IL_Create, IL_Create, UIntPtr, (In_Opt, Int32, InitialCount), (In_Opt, Int32, GrowCount), (In_Opt, Bool32, LargeIcons))
md_func_x(IL_Destroy, IL_Destroy, Bool32, (In, UIntPtr, ImageList))

md_func(ImageSearch, (Out, Variant, X), (Out, Variant, Y), (In, Int32, X1), (In, Int32, Y1), (In, Int32, X2), (In, Int32, Y2), (In, String, Image), (Ret, Bool32, Found))

md_func(IniDelete, (In, String, Path), (In, String, Section), (In_Opt, String, Key))
md_func(IniRead, (In, String, Path), (In_Opt, String, Section), (In_Opt, String, Key), (In_Opt, String, Default), (Ret, String, RetVal))
md_func(IniWrite, (In, String, Value), (In, String, Path), (In, String, Section), (In_Opt, String, Key))

md_func(InputBox, (In_Opt, String, Prompt), (In_Opt, String, Title), (In_Opt, String, Options), (In_Opt, String, Default), (Ret, Object, RetVal))

md_func_v(InstallKeybdHook, (In_Opt, Bool32, Install), (In_Opt, Bool32, Force))
md_func_v(InstallMouseHook, (In_Opt, Bool32, Install), (In_Opt, Bool32, Force))

md_func_x(IsLabel, IsLabel, Bool32, (In, String, Name))

md_func(KeyHistory, (In_Opt, Int32, MaxEvents))

md_func(KeyWait, (In, String, KeyName), (In_Opt, String, Options), (Ret, Bool32, RetVal))

md_func_v(ListHotkeys, md_arg_none)
md_func_v(ListLines, (In_Opt, Int32, Mode), (Ret, Int32, RetVal))
md_func_v(ListVars, md_arg_none)

md_func(ListViewGetContent, (In_Opt, String, Options), MD_CONTROL_ARGS, (Ret, Variant, RetVal))

md_func(LoadPicture, (In, String, Filename), (In_Opt, String, Options), (Out_Opt, Int32, ImageType), (Ret, UIntPtr, Handle))

md_func_v(MenuFromHandle, (In, UIntPtr, Handle), (Ret, Object, Menu))
md_func(MenuSelect, (In_Opt, Variant, WinTitle), (In_Opt, String, WinText), (In, String, Menu)
	, (In_Opt, String, SubMenu1), (In_Opt, String, SubMenu2), (In_Opt, String, SubMenu3), (In_Opt, String, SubMenu4), (In_Opt, String, SubMenu5), (In_Opt, String, SubMenu6)
	, (In_Opt, String, ExcludeTitle), (In_Opt, String, ExcludeText))

md_func(MonitorGet, (In_Opt, Int32, N), (Out_Opt, Int32, Left), (Out_Opt, Int32, Top), (Out_Opt, Int32, Right), (Out_Opt, Int32, Bottom), (Ret, Int32, RetVal))
md_func_x(MonitorGetCount, MonitorGetCount, Int32, md_arg_none)
md_func(MonitorGetName, (In_Opt, Int32, N), (Ret, String, RetVal))
md_func_x(MonitorGetPrimary, MonitorGetPrimary, Int32, md_arg_none)
md_func(MonitorGetWorkArea, (In_Opt, Int32, N), (Out_Opt, Int32, Left), (Out_Opt, Int32, Top), (Out_Opt, Int32, Right), (Out_Opt, Int32, Bottom), (Ret, Int32, RetVal))

md_func(MouseClick,
	(In_Opt, String, Button),
	(In_Opt, Int32, X), (In_Opt, Int32, Y),
	(In_Opt, Int32, ClickCount),
	(In_Opt, Int32, Speed),
	(In_Opt, String, DownOrUp),
	(In_Opt, String, Relative))
md_func(MouseClickDrag,
	(In_Opt, String, Button),
	(In, Int32, X1), (In, Int32, Y1),
	(In, Int32, X2), (In, Int32, Y2),
	(In_Opt, Int32, Speed),
	(In_Opt, String, Relative))
md_func(MouseGetPos, (Out_Opt, Int32, X), (Out_Opt, Int32, Y), (Out_Opt, Variant, Win), (Out_Opt, Variant, Control), (In_Opt, Int32, Flag))
md_func(MouseMove,
	(In, Int32, X), (In, Int32, Y),
	(In_Opt, Int32, Speed),
	(In_Opt, String, Relative))

md_func(MsgBox, (In_Opt, String, Text), (In_Opt, String, Title), (In_Opt, String, Options), (Ret, String, RetVal))

md_func(OnClipboardChange, (In, Object, Function), (In_Opt, Int32, AddRemove))
md_func(OnError, (In, Object, Function), (In_Opt, Int32, AddRemove))
md_func(OnExit, (In, Object, Function), (In_Opt, Int32, AddRemove))
md_func(OnMessage, (In, UInt32, Number), (In, Object, Function), (In_Opt, Int32, MaxThreads))

md_func_v(OutputDebug, (In, String, Text))

md_func(Pause, (In_Opt, Int32, NewState))

md_func_v(Persistent, (In_Opt, Bool32, NewValue), (Ret, Bool32, OldValue))

md_func(PixelGetColor, (In, Int32, X), (In, Int32, Y), (In_Opt, String, Mode), (Ret, String, Color))
md_func(PixelSearch, (Ret, Bool32, Found), (Out, Variant, X), (Out, Variant, Y), (In, Int32, X1), (In, Int32, Y1), (In, Int32, X2), (In, Int32, Y2), (In, UInt32, Color), (In_Opt, Int32, Variation))

#undef PostMessage
md_func_x(PostMessage, ScriptPostMessage, FResult, (In, UInt32, Msg), (In_Opt, Variant, wParam), (In_Opt, Variant, lParam), MD_CONTROL_ARGS_OPT)

md_func(ProcessClose, (In, String, Process), (Ret, UInt32, ClosedPID))
md_func_x(ProcessExist, ProcessExist, UInt32, (In_Opt, String, Process))
md_func(ProcessGetName, (In_Opt, String, Process), (Ret, String, Name))
md_func_x(ProcessGetParent, ProcessGetParent, UInt32, (In_Opt, String, Process))
md_func(ProcessGetPath, (In_Opt, String, Process), (Ret, String, Path))
md_func(ProcessSetPriority, (In, String, Priority), (In_Opt, String, Process), (Ret, UInt32, FoundPID))
md_func(ProcessWait, (In, String, Process), (In_Opt, Float64, Timeout), (Ret, UInt32, FoundPID))
md_func(ProcessWaitClose, (In, String, Process), (In_Opt, Float64, Timeout), (Ret, UInt32, UnclosedPID))

md_func(Reload, md_arg_none)

md_func(Run, (In, String, Target), (In_Opt, String, WorkingDir), (In_Opt, String, Options), (Out_Opt, Variant, PID))
md_func_v(RunAs, (In_Opt, String, User), (In_Opt, String, Password), (In_Opt, String, Domain))

md_func_v(Send, (In, String, Keys))
md_func_v(SendEvent, (In, String, Keys))
md_func_v(SendInput, (In, String, Keys))
md_func(SendLevel, (In, Int32, Level), (Ret, Int32, RetVal))

#undef SendMessage
md_func_x(SendMessage, ScriptSendMessage, FResult, (In, UInt32, Msg), (In_Opt, Variant, wParam), (In_Opt, Variant, lParam), MD_CONTROL_ARGS_OPT, (In_Opt, Int32, Timeout), (Ret, UIntPtr, RetVal))

md_func(SendMode, (In, String, Mode), (Ret, Variant, RetVal))
md_func_v(SendPlay, (In, String, Keys))
md_func_v(SendText, (In, String, Text))

md_func(SetCapsLockState, (In_Opt, String, State))
md_func(SetControlDelay, (In, Int32, Delay), (Ret, Int32, RetVal))
md_func(SetDefaultMouseSpeed, (In, Int32, Speed), (Ret, Int32, RetVal))
md_func(SetKeyDelay, (In_Opt, Int32, Delay), (In_Opt, Int32, Duration), (In_Opt, String, Mode))
md_func(SetMouseDelay, (In, Int32, Delay), (In_Opt, String, Mode), (Ret, Int32, RetVal))
md_func(SetNumLockState, (In_Opt, String, State))
md_func(SetRegView, (In, String, RegView), (Ret, Variant, RetVal))
md_func(SetScrollLockState, (In_Opt, String, State))
md_func_v(SetStoreCapsLockMode, (In, Bool32, Mode), (Ret, Bool32, RetVal))

md_func(SetTimer, (In_Opt, Object, Function), (In_Opt, Int64, Period), (In_Opt, Int32, Priority))

md_func(SetTitleMatchMode, (In, String, Mode), (Ret, Variant, RetVal))
md_func(SetWinDelay, (In, Int32, Delay), (Ret, Int32, RetVal))
md_func(SetWorkingDir, (In, String, Path))

md_func(Shutdown, (In, Int32, Flags))

md_func_x(Sleep, ScriptSleep, Void, (In, Int32, Delay))

md_func_v(SoundBeep, (In_Opt, Int32, Duration), (In_Opt, Int32, Frequency))
md_func(SoundPlay, (In, String, Path), (In_Opt, String, Wait))

md_func(StatusBarGetText, (In_Opt, Int32, Part), MD_WINTITLE_ARGS, (Ret, String, RetVal))
md_func(StatusBarWait, (In_Opt, String, Text), (In_Opt, Float64, Timeout), (In_Opt, Int32, Part), (In_Opt, Variant, WinTitle), (In_Opt, String, WinText), (In_Opt, Int32, Interval), (In_Opt, String, ExcludeTitle), (In_Opt, String, ExcludeText), (Ret, Bool32, RetVal))

md_func(StrSplit, (In, String, String), (In_Opt, Variant, Delimiters), (In_Opt, String, OmitChars), (In_Opt, Int32, MaxParts), (Ret, Object, RetVal))

md_func(Suspend, (In_Opt, Int32, Mode))

md_func_x(SysGet, SysGet, Int32, (In, Int32, Index))
md_func(SysGetIPAddresses, (Ret, Object, RetVal))

md_func(Thread, (In, String, Command), (In_Opt, Int32, Value1), (In_Opt, Int32, Value2))

md_func(ToolTip, (In_Opt, String, Text), (In_Opt, Int32, X), (In_Opt, Int32, Y), (In_Opt, Int32, Index), (Ret, UInt32, Hwnd))

md_func(TraySetIcon, (In_Opt, String, File), (In_Opt, Int32, Number), (In_Opt, Bool32, Freeze))
md_func(TrayTip, (In_Opt, String, Text), (In_Opt, String, Title), (In_Opt, String, Options))

md_func(WinActivate, MD_WINTITLE_ARGS)
md_func(WinActivateBottom, MD_WINTITLE_ARGS)
md_func(WinClose, (In_Opt, Variant, WinTitle), (In_Opt, String, WinText), (In_Opt, Float64, WaitTime), (In_Opt, String, ExcludeTitle), (In_Opt, String, ExcludeText))
md_func(WinGetAlwaysOnTop, MD_WINTITLE_ARGS, (Ret, Bool32, RetVal))
md_func(WinGetClass, MD_WINTITLE_ARGS, (Ret, String, Class))
md_func(WinGetClientPos, (Out_Opt, Int32, X), (Out_Opt, Int32, Y), (Out_Opt, Int32, Width), (Out_Opt, Int32, Height), MD_WINTITLE_ARGS)
md_func(WinGetControls, MD_WINTITLE_ARGS, (Ret, Object, ClassNNArray))
md_func(WinGetControlsHwnd, MD_WINTITLE_ARGS, (Ret, Object, HwndArray))
md_func(WinGetCount, MD_WINTITLE_ARGS, (Ret, Int32, Count))
md_func(WinGetEnabled, MD_WINTITLE_ARGS, (Ret, Bool32, Enabled))
md_func(WinGetExStyle, MD_WINTITLE_ARGS, (Ret, UInt32, ExStyle))
md_func(WinGetID, MD_WINTITLE_ARGS, (Ret, UInt32, Hwnd))
md_func(WinGetIDLast, MD_WINTITLE_ARGS, (Ret, UInt32, Hwnd))
md_func(WinGetList, MD_WINTITLE_ARGS, (Ret, Object, HwndArray))
md_func(WinGetMinMax, MD_WINTITLE_ARGS, (Ret, Int32, RetVal))
md_func(WinGetPID, MD_WINTITLE_ARGS, (Ret, UInt32, PID))
md_func(WinGetPos, (Out_Opt, Int32, X), (Out_Opt, Int32, Y), (Out_Opt, Int32, Width), (Out_Opt, Int32, Height), MD_WINTITLE_ARGS)
md_func(WinGetProcessName, MD_WINTITLE_ARGS, (Ret, String, Name))
md_func(WinGetProcessPath, MD_WINTITLE_ARGS, (Ret, String, Path))
md_func(WinGetStyle, MD_WINTITLE_ARGS, (Ret, UInt32, Style))
md_func(WinGetText, MD_WINTITLE_ARGS, (Ret, String, Title))
md_func(WinGetTitle, MD_WINTITLE_ARGS, (Ret, String, Title))
md_func(WinGetTransColor, MD_WINTITLE_ARGS, (Ret, String, Color))
md_func(WinGetTransparent, MD_WINTITLE_ARGS, (Ret, Variant, RetVal))
md_func(WinHide, MD_WINTITLE_ARGS)
md_func(WinKill, (In_Opt, Variant, WinTitle), (In_Opt, String, WinText), (In_Opt, Float64, WaitTime), (In_Opt, String, ExcludeTitle), (In_Opt, String, ExcludeText))
md_func(WinMaximize, MD_WINTITLE_ARGS)
md_func(WinMinimize, MD_WINTITLE_ARGS)
md_func_v(WinMinimizeAll, md_arg_none)
md_func_v(WinMinimizeAllUndo, md_arg_none)
md_func(WinMove, (In_Opt, Int32, X), (In_Opt, Int32, Y), (In_Opt, Int32, Width), (In_Opt, Int32, Height), MD_WINTITLE_ARGS)
md_func(WinMoveBottom, MD_WINTITLE_ARGS)
md_func(WinMoveTop, MD_WINTITLE_ARGS)
md_func(WinRedraw, MD_WINTITLE_ARGS)
md_func(WinRestore, MD_WINTITLE_ARGS)
md_func(WinSetAlwaysOnTop, (In_Opt, Int32, Value), MD_WINTITLE_ARGS)
md_func(WinSetEnabled, (In, Int32, Value), MD_WINTITLE_ARGS)
md_func(WinSetExStyle, (In, String, Style), MD_WINTITLE_ARGS)
md_func(WinSetRegion, (In_Opt, String, Options), MD_WINTITLE_ARGS)
md_func(WinSetStyle, (In, String, Style), MD_WINTITLE_ARGS)
md_func(WinSetTitle, (In, String, NewTitle), MD_WINTITLE_ARGS)
md_func(WinSetTransColor, (In, Variant, Value), MD_WINTITLE_ARGS)
md_func(WinSetTransparent, (In, String, Value), MD_WINTITLE_ARGS)
md_func(WinShow, MD_WINTITLE_ARGS)
md_func(WinWait, (In_Opt, Variant, WinTitle), (In_Opt, String, WinText), (In_Opt, Float64, Timeout), (In_Opt, String, ExcludeTitle), (In_Opt, String, ExcludeText), (Ret, UInt32, Hwnd))
md_func(WinWaitActive, (In_Opt, Variant, WinTitle), (In_Opt, String, WinText), (In_Opt, Float64, Timeout), (In_Opt, String, ExcludeTitle), (In_Opt, String, ExcludeText), (Ret, UInt32, Hwnd))
md_func(WinWaitClose, (In_Opt, Variant, WinTitle), (In_Opt, String, WinText), (In_Opt, Float64, Timeout), (In_Opt, String, ExcludeTitle), (In_Opt, String, ExcludeText), (Ret, UInt32, RetVal))
md_func(WinWaitNotActive, (In_Opt, Variant, WinTitle), (In_Opt, String, WinText), (In_Opt, Float64, Timeout), (In_Opt, String, ExcludeTitle), (In_Opt, String, ExcludeText), (Ret, UInt32, RetVal))
