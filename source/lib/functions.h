
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
