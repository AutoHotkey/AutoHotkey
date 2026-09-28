#NoTrayIcon
#SingleInstance Off

#include UX\install.ahk

if A_Args.Length {
    Install_Main
    ExitApp
}

#include UX\ui-setup.ahk

UnpackFiles(installDir) {
    DirCreate dir := installDir "\.staging\" A_ScriptName
    SetWorkingDir dir
    OnExit cleanup
    DirCreate "UX"
    DirCreate "UX\inc"
    DirCreate "UX\Templates"
    DllCall("EnumResourceNames", "ptr", 0, "ptr", 10, "ptr", CallbackCreate(enumProc, "F"), "ptr", 0)
    enumProc(hmod, lpType, lpName, lParam) {
        if (lpName & ~0xffff) {
            f := StrGet(lpName)
            (FileInstall)(f, f, 1)
        }
        return true
    }
    return dir
    cleanup(*) {
        SetWorkingDir A_ScriptDir
        DirDelete dir, true
        try DirDelete installDir "\.staging"
    }
}
