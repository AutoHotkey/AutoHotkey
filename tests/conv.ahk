Transform StrUTF8, ToCodePage, 65001, "にほん"
Transform StrNative, FromCodepage, 65001, %StrUTF8%
MsgBox "%StrNative%"
