
md_func(DriveGetCapacity, (In, String, Path), (Ret, Int64, RetVal))
md_func(DriveGetFilesystem, (In, String, Drive), (Ret, String, RetVal))
md_func(DriveGetLabel, (In, String, Drive), (Ret, String, RetVal))
md_func(DriveGetSerial, (In, String, Drive), (Ret, Int64, RetVal))
md_func(DriveGetSpaceFree, (In, String, Path), (Ret, Int64, RetVal))
md_func(DriveGetStatus, (In, String, Path), (Ret, String, RetVal))
md_func(DriveGetStatusCD, (In_Opt, String, Drive), (Ret, String, RetVal))
md_func(DriveGetType, (In, String, Path), (Ret, String, RetVal))
md_func(DriveGetList, (In_Opt, String, Type), (Ret, String, RetVal))

md_func(DriveEject, (In_Opt, String, Drive))
md_func(DriveRetract, (In_Opt, String, Drive))

md_func(DriveLock, (In, String, Drive))
md_func(DriveUnlock, (In, String, Drive))

md_func(DriveSetLabel, (In, String, Drive), (In_Opt, String, Label))


md_func(EnvGet, (In, String, VarName), (Ret, String, RetVal))
md_func_x(EnvSet, EnvSet, NzIntWin32, (In, String, VarName), (In_Opt, String, Value))


md_func(FileAppend, (In, Variant, Value), (In_Opt, String, Path), (In_Opt, String, Options))
md_func(FileCopy, (In, String, Source), (In, String, Dest), (In_Opt, Int32, Overwrite))
md_func(FileCreateShortcut, (In, String, Target), (In, String, LinkFile), (In_Opt, String, WorkingDir),
	(In_Opt, String, Args), (In_Opt, String, Description), (In_Opt, String, IconFile),
	(In_Opt, String, ShortcutKey), (In_Opt, Int32, IconNumber), (In_Opt, Int32, RunState))
md_func(FileDelete, (In, String, Pattern))
md_func_x(FileExist, FileExist, Void, (In, String, Pattern), (Ret, String, RetVal))
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


md_func(DirCopy, (In, String, Source), (In, String, Dest), (In_Opt, Int32, Overwrite))
md_func(DirCreate, (In, String, Path))
md_func(DirDelete, (In, String, Path), (In_Opt, Bool32, Recurse))
md_func_x(DirExist, DirExist, Void, (In, String, Pattern), (Ret, String, RetVal))
md_func(DirMove, (In, String, Source), (In, String, Dest), (In_Opt, String, Flag))
md_func(DirSelect, (In_Opt, String, StartingFolder), (In_Opt, Int32, Options), (In_Opt, String, Prompt), (Ret, String, RetVal))


md_func(Download, (In, String, URL), (In, String, Path))


md_func_x(InstallKeybdHook, InstallKeybdHook, Void, (In_Opt, Bool32, Install), (In_Opt, Bool32, Force))
md_func_x(InstallMouseHook, InstallMouseHook, Void, (In_Opt, Bool32, Install), (In_Opt, Bool32, Force))


md_func(MonitorGet,
	(In_Opt, Int32, N),
	(Out_Opt, Int32, Left),
	(Out_Opt, Int32, Top),
	(Out_Opt, Int32, Right),
	(Out_Opt, Int32, Bottom),
	(Ret, Int32, RetVal))

md_func(MonitorGetWorkArea,
	(In_Opt, Int32, N),
	(Out_Opt, Int32, Left),
	(Out_Opt, Int32, Top),
	(Out_Opt, Int32, Right),
	(Out_Opt, Int32, Bottom),
	(Ret, Int32, RetVal))

md_func_x(MonitorGetCount, MonitorGetCount, Int32, md_arg_none)
md_func_x(MonitorGetPrimary, MonitorGetPrimary, Int32, md_arg_none)

md_func(MonitorGetName, (In_Opt, Int32, N), (Ret, String, RetVal))


md_func_x(Persistent, Persistent, Void, (In_Opt, Bool32, NewValue), (Ret, Bool32, OldValue))


md_func_x(RunAs, RunAs, Void, (In_Opt, String, User), (In_Opt, String, Password), (In_Opt, String, Domain))


md_func_x(SysGet, SysGet, Int32, (In, Int32, Index))
md_func(SysGetIPAddresses, (Ret, Object, RetVal))
