VarSetCapacity(x,2,1)
;MsgBox, % NumGet(x, 0, "UChar")
;MsgBox, % NumGet(&x, 0, "UChar")

;MsgBox, % NumGet(x, 0, "UShort")
;MsgBox, % NumGet(&x, 0, "UShort")

NumPut(30, x, 0, "UChar")
NumPut(20, x, 1, "UChar")
MsgBox, % NumGet(x, 0, "UChar")
MsgBox, % NumGet(x, 1, "UChar")
MsgBox, % NumGet(x, 0, "UShort")


;MsgBox % VarSetCapacity(x,3) "|" VarSetCapacity(y,4)
;MsgBox % VarSetCapacity(z,1,1) "|" NumGet(&z)
