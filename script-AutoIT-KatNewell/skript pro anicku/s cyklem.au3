$Imagesoft="ImageJ"
$namex="20110210_Si_24_100_09l_"
$NN=40
;$Nname="0"
;$name="0"

for $j=0 to $NN step 1
	
if StringLen($j)<2 then 
	$Nname="0" & $j
Else
	$Nname=$j
	EndIf

$name=$namex & $Nname
Sleep(100)

WinActivate($Imagesoft,"")
WinWaitActive($Imagesoft,"")
Send("^o")
Send($name & ".tif")
Sleep(500)
Send("!o")
Sleep(1000)

WinActivate($Imagesoft,"")
WinWaitActive($Imagesoft,"")
Send("{ALT}")
Sleep(200)
Send("{RIGHT 2}")
Send("{DOWN 5}")
Send("{RIGHT 1}")
Send("{ENTER 1}")
Sleep(1000)

WinActivate($Imagesoft,"")
WinWaitActive($Imagesoft,"")
Sleep(500)
Send("^s")
Send("{ENTER 1}")
Sleep(1000)

WinActivate($Imagesoft,"")
WinWaitActive($Imagesoft,"")
Send("{ALT}")
Sleep(200)
Send("{RIGHT 3}")
Send("{DOWN 2}")
Send("{ENTER 1}")
Sleep(500)
Send("^+s")
Sleep(500)

WinActivate($Imagesoft,"")
WinWaitActive($Imagesoft,"")
Send("{ALT}")
Sleep(200)
Send("{DOWN 9}")
Send("{RIGHT 1}")
Send("{ENTER 1}")
Sleep(500)
Send($name & "a" &".tif")
Sleep(500)
Send("{ENTER 1}")
Sleep(1000)
Send("^w")
Sleep(300)

WinActivate($Imagesoft,"")
WinWaitActive($Imagesoft,"")
Send("^s")
Send("{ENTER 1}")
Sleep(1000)
Send("^w")
Sleep(300)

WinActivate($Imagesoft,"")
WinWaitActive($Imagesoft,"")
Send("^w")
Sleep(300)

WinActivate($Imagesoft,"")
WinWaitActive($Imagesoft,"")
Send("^o")
Send($name & ".tif (blue).tif")
Sleep(300)
Send("!o")
Sleep(1000)

WinActivate($Imagesoft,"")
WinWaitActive($Imagesoft,"")
Send("^o")
Send($name & ".tif (green).tif")
Sleep(300)
Send("!o")
Sleep(1000)

WinActivate($Imagesoft,"")
WinWaitActive($Imagesoft,"")
Send("{ALT}")
Sleep(200)
Send("{RIGHT 2}")
Send("{DOWN 5}")
Send("{RIGHT 1}")
Send("{DOWN 1}")
Send("{ENTER 1}")
Sleep(500)
Send("{ENTER 1}")

WinActivate($Imagesoft,"")
WinWaitActive($Imagesoft,"")
Send("{ALT}")
Sleep(200)
Send("{RIGHT 2}")
Send("{DOWN 1}")
Send("{RIGHT 1}")
Send("{DOWN 4}")
Send("{ENTER 1}")
Sleep(500)

WinActivate($Imagesoft,"")
WinWaitActive($Imagesoft,"")
Send("{ALT}")
Sleep(200)
Send("{DOWN 9}")
Send("{RIGHT 1}")
Send("{ENTER 1}")
Sleep(500)
Send($Nname & ".tif")
Sleep(500)
Send("{ENTER 1}")
Sleep(1000)
Send("^w")
Sleep(500)
Send("^w")
Sleep(500)

Next
