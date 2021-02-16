$CCD="C7557-01"

AutoItSetOption("WinTitleMatchMode",1)

WinActivate($CCD,"")
WinWaitActive($CCD,"")
$vslses="vsl"
$step=100
$min=0
$max=4000
For $j=$min to $max step $step  
Send("!fo")
Sleep(500)
Send($vslses & $j & ".BIN")
WinActivate("Open")
WinWaitActive("Open")
Send("{ENTER 1}")
;Send("!o")
Sleep(1500)
Send("!fa")
Sleep(500)
Send($vslses & $j & ".TXT")
;Send("!t")
;Send("{DOWN 2}")
;Send("{ENTER 1}")
;Send("!n")
;Send($vslses & $j)
Sleep(1500)
Send("{ENTER 1}")
Sleep(1500)
;WinActivate("Save As...","")
;WinWaitActive("Save As...","")
;Sleep(500)
;$j=$j-$zeroSES
;While StringLen($j)<2
;$j="0" & $j
;WEnd
;$a=$j*0.5
;Send($vsl_ses & $a)
;Send("{ENTER 1}")
;WinWaitActive($CCD)
;$j=$j+$zeroSES
;$imotor=Round($imotor+$imotorstep,2)
Next

