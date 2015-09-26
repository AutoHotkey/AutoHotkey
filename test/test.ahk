#SingleInstance force
;~ #include <Attach>
;~ Gui, Margin, 0, 0
try {
	Gui,+Scroll +Resize ;0x300000
	
	;~ Gui,Add, Edit, ah1/3 w300 h20 hwndt1,Test1
	;~ Gui,Add, Edit, ay1/3 ah1/3 w300 h20 hwndt2,Test2 ; y25 of 100 = 1/4
	;~ Gui,Add, Edit, ay1/3 ah1/3 w300 h20 hwndt3,Test3
	;Gui,Add, Edit, ay1/3 ah1/3 w300 h20 hwndt4,Test4
	Gui,Add, Edit, ay1/3 ayr ah x+10 ym0 w300 h100,Test5

	/*
	Gui,Add, Edit, avg1 ah w300 h20 hwndt1,Test1 ; grew by 10px, moved down by 10px g1 = 0
	Gui,Add, Edit, avg1 ah2  w300 h20 hwndt2,Test2 ; g1 = 20. grew by 10, moved down by 10
	Gui,Add, Edit, avg1 w300 h20 hwndt3,Test3 ; g1 = 40
	Gui,Add, Edit, am5 ag1 ah w300 h20 hwndt4,Test4
	Gui,
	*/
} catch {
	;Gui +0x300000 Resize
	Gui +Resize
	Gui,Add, Edit, w300 h20 hwndt1,Test1
	Gui,Add, Edit, w300 h20 hwndt2,Test2 ; y25 of 100 = 1/4
	Gui,Add, Edit, w300 h20 hwndt3,Test3
	Gui,Add, Edit, w300 h20 hwndt4,Test4
	;~ Attach(t1,"h")
	;~ Attach(t2,"y h")
	;~ Attach(t3,"y h")
	;~ Attach(t4,"y h")

}
;Gui,Show, x0 w320 h212
Gui,Show, x0 ;w300 h100
;~ Gui,Show, x0 w300 h300
return
GuiClose:
ExitApp