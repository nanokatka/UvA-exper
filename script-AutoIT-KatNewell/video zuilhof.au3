$Imagesoft="Origin"
$namex="12a10-video-od15a10s100x_"
$NN=100

for $j=1 to $NN step 1
$name=$namex & $j
Sleep(100)

WinActivate($Imagesoft,"")
WinWaitActive($Imagesoft,"")
MouseClick("left", 1677, 189, 1)
Send("!fms")
Send("{ENTER 1}")
Send($name & ".txt")
Sleep(500)
Send("!o")
Sleep(3000)

MouseClick("left", 2303, 495, 1)
Sleep(1000)
MouseClick("left", 1518, 215, 1)
Send("^c")
Sleep(3000)
MouseClick("left", 2005, 567, 1)
Send("^v")
Sleep(10000)
MouseClick("left", 1658, 544, 1)
Send("!fe1")
Sleep(500)
Send("!n")
Sleep(12000)


Next
