#NoEnv
; This file clears the checksum of the AutoHotkeySC.bin files.
; It also a demo of the new, object oriented, advanced file IO.

FixSC("SC (self-contained)")
FixSC("SC (minimum size)")

FixSC(folder){
	scbin := A_ScriptDir . folder . "\AutoHotkeySC.bin"
	IfNotExist %scbin%
		return
	; Open a file
	; object := FileOpen(filename, flags[, codepage])
	file := FileOpen(scbin, 3) ; UPDATE = 3
	if (!IsObject(file))
		return ; Can not open the file

	; object.Seek(distance[, origin=0])
	; origin: SEEK_SET = 0, SEEK_CUR = 1, SEEK_END = 2
	file.Seek(60)

	; Reserve memory to read the integer
	VarSetCapacity(offset, 4)

	; Read raw binary data:
	; object.RawRead(VarOrAddress, bytes)
	file.RawRead(offset, 4)

	offset := NumGet(offset, 0, "int") + 88

	file.Seek(offset)

	VarSetCapacity(csum, 4, 0)
	
	; object.RawWrite(VarOrAddress, bytes). Write raw binary data.
	file.RawWrite(csum, 4)
	
	; Close the file. It is also closed when the object is freed.
	; object.Close()
	file.Close()
}
