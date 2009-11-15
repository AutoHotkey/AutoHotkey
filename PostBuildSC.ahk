#NoEnv
; This file clear the checksum of AutoHotkeySC.bin.
; It also a demo of the new, object oriented, advanced file IO.

; object := FileOpen(filename, flags, codepage)
file := FileOpen(A_ScriptDir . "\SC (self-contained)\AutoHotkeySC.bin", 3) ; UPDATE = 3

; object.Seek(distance, origin)
; origin: SEEK_SET = 0, SEEK_CUR = 1, SEEK_END = 2
file.Seek(60, 0) ; SEEK_SET = 0

VarSetCapacity(offset, 4) ; This is not required actually.

; object.RawRead(VarOrAddress, bytes). Read raw binary data.
file.RawRead(offset, 4)

offset := NumGet(offset, 0, "int") + 88

file.Seek(offset, 0)

VarSetCapacity(csum, 4, 0)

; object.RawWrite(VarOrAddress, bytes). Write raw binary data.
file.RawWrite(csum, 4)

; object.Close(). Close the file, it should be closed with the object is freed, too.
file.Close()
