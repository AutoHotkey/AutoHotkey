((name voxlyn))
#Persistent
SetTimer, Fish, 1000
return

Fish:
Click
return#Persistent
SetTimer, Fish, 1000
return

Fish:
Random, delay, 800, 1200
Click
Sleep, delay
return#Persistent
SetTimer, Fish, 1000
return

Fish:
Random, delay, 800, 1200
Click
Sleep, delay

; random stop
Random, stopChance, 1, 100
if (stopChance < 10) {
    Sleep, 5000
}
return#Persistent
SetTitleMatchMode, 2
CoordMode, Mouse, Screen

F8::Toggle := !Toggle

SetTimer, Fish, 100
return

Fish:
if (!Toggle)
    return

Random, delay, 600, 1200
Click
Sleep, delay
return
