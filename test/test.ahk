
#SingleInstance force

Gui +Resize +Scroll
Gui, Margin, 5, 5

Gui, Add, Edit, xm w100 aw1/4 ym h40 ah1/4, Test01
Gui, Add, Edit, xm w100 aw1/4 h40 ay1/4 ah1/4, Test02
Gui, Add, Edit, xm w100 aw1/4 h40 ay1/4 ah1/4, Test03
Gui, Add, Edit, xm w100 aw1/4 h40 ay1/4 ah1/4, Test04

Gui, Add, Edit, xp+105 ax1/4 ayr w100 aw1/4 h55 ym ah1/3, Test05
Gui, Add, Edit, xp axp w100 aw1/4 h55 yp+60 ah1/3 ay1/3, Test06
Gui, Add, Edit, xp axp w100 aw1/4 h55 yp+60 ah1/3 ay1/3, Test07

Gui, Add, Edit, xp+105 ax1/4 w100 aw1/4 h85 ym ayr ah1/2, Test08
Gui, Add, Edit, xp axp w100 aw1/4 h85 yp+90 ah1/2 ay1/2, Test09

Gui, Add, Edit, xp+105 ax1/4 w100 aw1/4 h175 ym ah, Test10

Gui, Add, Edit, xm w100 ayr axr ay1 aw1/4 h175 ah1/2, Test11
Gui, Add, Edit, xp+105 w100 ax1/4 aw1/4 ayp h175 ah1/2, Test12
Gui, Add, Edit, xp+105 w205 ax1/4 aw1/2 ayp h175 ah1/2, Test13

Gui, Add, Edit, xm w415 aw1 ay1/2, Test14

Gui, Show, x0

return
GuiClose:
ExitApp
