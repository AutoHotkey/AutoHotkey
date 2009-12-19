#NoEnv
Path = %1%
FileRead, bin, %Path%

; Clear CheckSum of NT header, since Ahk2Exe does not support non-zero values.
NumPut(0, bin, NumGet(bin, 60, "int") + 88, "int")

file := DllCall("CreateFile", "str", Path, "uint", 0x40000000, "uint", 0, "uint", 0, "uint", 3, "uint", 0, "uint", 0)
if file = -1
{
    MsgBox, 16, Error, Failed to open AutoHotkeySC.bin to correct CheckSum.
    ExitApp % A_LastError
}

if !DllCall("WriteFile", "uint", file, "uint", &bin, "uint", VarSetCapacity(bin), "uint*", 0, "uint", 0)
{
    MsgBox, 16, Error, Failed to write to AutoHotkeySC.bin to correct CheckSum.
    ExitApp % A_LastError
}

DllCall("CloseHandle", "uint", file)