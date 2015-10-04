
#SingleInstance force

Gui +Resize +Scroll
Gui, Margin, 5, 5
Gui, Add, Edit, xm w100 aw1/4 ym h40 ah1/4, Test1
Gui, Add, Edit, xm w100 aw1/4 h40 ay1/4 ah1/4, Test2
Gui, Add, Edit, xm w100 aw1/4 h40 ay1/4 ah1/4, Test3
Gui, Add, Edit, xm w100 aw1/4 h40 ay1/4 ah1/4, Test4

Gui, Add, Edit, xp+105 ax1/4 ayr w100 aw1/4 h55 ym ah1/3, Test5
Gui, Add, Edit, xp axp w100 aw1/4 h55 yp+60 ah1/3 ay1/3, Test6
Gui, Add, Edit, xp axp w100 aw1/4 h55 yp+60 ah1/3 ay1/3, Test7

Gui, Add, Edit, xp+105 ax1/4 w100 aw1/4 h85 ym ayr ah1/2, Test8
Gui, Add, Edit, xp axp w100 aw1/4 h85 yp+90 ah1/2 ay1/2, Test9

Gui, Add, Edit, xp+105 ax1/4 w100 aw1/4 h175 ym ah, Test10
Gui, Show, x0

return
GuiClose:
ExitApp
