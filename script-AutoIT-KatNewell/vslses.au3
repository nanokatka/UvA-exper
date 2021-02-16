$CCD="C7557-01"
$motor="SMCView"

$vsl_ses=$CmdLine[1]  ; co sa meria (vsl alebo ses)
$length=$CmdLine[2]    ;dlzka pruzku v mikrometrech
$step=$CmdLine[3]        ;krok v mikrometrech
$integrtime=$CmdLine[4]   ; pocet akumulacii 
$zero=$CmdLine[5]

AutoItSetOption("WinTitleMatchMode",3)  ;exact title match

$isleep=$integrtime+5000  ; cas mereni, gate width + gate delay nesmi byt vetsi nez 0.1s!!!!!
$imotor=$zero  ;hodnota na motorku na zacatku
$imotorstep=$step
$lengthzero=$length+$zero

;hleda okno motorku
For $j=$zero to $lengthzero step $step  
    $var=WinList()   
    for $i=1 to $var[0][0]
	   if $var[$i][0]=$motor Then
		 $Handle=$var[$i][1] 
	   EndIf
	Next

AutoItSetOption("WinTitleMatchMode",1)

$var=WinGetTitle($Handle)
WinActivate($Handle,"")
Sleep(500)
MouseClick("left",255,121,1,1)  ;klikne na polohu motorku
Sleep(300)
Send("{BS 6}")  ; vymaze polohu na motorku
Send("{DEL 6}") ; pro jistotu, aby to vymazal
Sleep(300)
;$smotor=StringSplit($imotor,".")
;StringReplace($smotor,".",",")
;Send($smotor[1])
Send($j)
;If $smotor[0]>1 Then
;	Send("," & $smotor[2])
;EndIf	

Sleep(500)
MouseClick("left",183,145,1,1)  ;klikne na start
Sleep(2000)

WinActivate($CCD,"")
WinWaitActive($CCD,"")

Send("{F1}")  ; zacne meranie
Sleep($isleep)
Send("!fa")
Sleep(200)
Send("!t")
Sleep(200)
Send("{DOWN 2}")
Sleep(200)
Send("{ENTER 1}")
Sleep(200)
Send("!n")
;WinActivate("Save As...","")
;WinWaitActive("Save As...","")
Sleep(150)
While StringLen($j)<3
	$j="0" & $j
WEnd
Send($vsl_ses & $j)
Send("{ENTER 1}")
WinWaitActive($CCD)

$imotor=Round($imotor+$imotorstep,2)
Next

