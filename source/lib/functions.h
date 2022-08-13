
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


md_func_x(SysGet, SysGet, Int32, (In, Int32, Index))
