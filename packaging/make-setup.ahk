#Requires AutoHotkey v2.0

OnError (e, *) => ExitApp(e.Line)

; Files to include in the installer
FileCopy "..\license.txt", "content", 1
FileCopy "..\bin\AutoHotkey32.exe", "content", 1
FileCopy "..\bin\AutoHotkey64.exe", "content", 1
FileCopy "..\help\AutoHotkey.chm", "content", 1

; Base file of the installer itself
; (AutoHotkey32.exe with custom version info)
FileCopy "..\bin\AutoHotkey_setup.exe", ".", 1

SetWorkingDir "content"

fa := []
Loop Read A_ScriptDir "\files.txt" {
    Loop Files A_LoopReadLine, "F"
        fa.Push(A_LoopFilePath)
    else Loop Files A_LoopReadLine "\*", "FR"
        fa.Push(A_LoopFilePath)
}

EmbedFiles("..\AutoHotkey_setup.exe", [[1, "..\setup-stub.ahk"], fa*])

EmbedFiles(ExePath, FilePaths) {
    if !FilePaths.Length
        throw ValueError("Empty FilePaths")
    if !(hupd := DllCall("BeginUpdateResource", "str", ExePath, "int", false))
        throw OSError()
    try {
        for fp in FilePaths {
            if fp is Array ; [id, path]
                rn := fp[1], fp := fp[2]
            else
                rn := fp
            data := FileRead(fp, "RAW")
            if !DllCall("UpdateResource", "ptr", hupd, "ptr", 10 ; RCDATA
                        , "ptr", rn is String ? StrPtr(rn) : rn ; Resource name or ID
                        , "ushort", 1033, "ptr", data, "uint", data.size)
                throw OSError()
        }
        updated := IsSet(data)
    }
    finally {
        updated := DllCall("EndUpdateResource", "ptr", hupd, "int", !IsSet(updated))
    }
    if !updated
        throw OSError()
}