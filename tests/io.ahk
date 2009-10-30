FileRead OutVar, *P65001 %A_ScriptDir%\vectors\UTF-8-nobom.txt
MsgBox %OutVar%

FileRead OutVar, %A_ScriptDir%\vectors\UTF-8.txt
MsgBox %OutVar%

FileRead OutVar, %A_ScriptDir%\vectors\UTF-16.txt
MsgBox %OutVar%

FileDelete %A_ScriptDir%\vectors\output.txt
FileAppend %OutVar%, %A_ScriptDir%\vectors\output.txt, UTF-8-RAW
