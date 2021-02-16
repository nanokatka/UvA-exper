$CCD="C7557-01"
$motor="SMCView"

$vsl_ses=$CmdLine[1]  ; co sa meria (vsl alebo ses)
$length=$CmdLine[2]    ;dlzka pruzku v mikrometrech
$step=$CmdLine[3]        ;krok v mikrometrech
$integrtime=$CmdLine[4]   ; pocet akumulacii 
$zeroSES=$CmdLine[5]      ;ses spot length (um)

AutoItSetOption("WinTitleMatchMode",3)  ;exact title match

$isleep=$integrtime+5000  ; cas mereni, gate width + gate delay nesmi byt vetsi nez 0.1s!!!!!
$imotor=$zeroSES*0.001   ;hodnota na motorku na zacatku
$imotorstep=$step*0.001
$lengthtotal=$length+$zeroSES

;hleda okno motorku
For $j=$zeroSES to $lengthtotal step $step  
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
MouseClick("left",274,122,1,1)  ;klikne na polohu motorku
Sleep(300)
Send("{BS 6}")  ; vymaze polohu na motorku
Send("{DEL 6}") ; pro jistotu, aby to vymazal
Sleep(300)

$smotor=StringSplit($imotor,",")
StringReplace($smotor,".",",")
Send($smotor[1])
If $smotor[0]>1 Then
	Send("," & $smotor[2])
EndIf	

Sleep(500)
MouseClick("left",185,146,1,1)  ;klikne na start
Sleep(2000)

WinActivate($CCD,"")
WinWaitActive($CCD,"")

Send("{F1}")  ; zacne meranie
Sleep($isleep)
Send("!fa")
Send("!t")
Send("{DOWN 2}")
Send("{ENTER 1}")
Send("!n")
;WinActivate("Save As...","")
;WinWaitActive("Save As...","")
Sleep(150)
$j=$j-$zeroSES
While StringLen($j)<2
	$j="0" & $j
WEnd
Send($vsl_ses & $j)
Send("{ENTER 1}")
WinWaitActive($CCD)
$j=$j+$zeroSES

$imotor=Round($imotor+$imotorstep,2)
Next

