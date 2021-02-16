;Program na ovladanie merania vsl/ses s motorom.
;posuva motor a meria spektra (iCCD) so zadanym krokom, namerana data exportuje
;okna motoru a CCD musia byt otvorene, okno motoru musi byt v spravnej polohe na obrazovke (stlacit REFRESH)
;na zaciatku je treba nastavit adresar, do ktoreho sa exportuju data

$andor="Andor iStar"
$motor="SMCView"

$length=$CmdLine[2]    ;dlzka pruzku v mikrometroch
$step=$CmdLine[3]        ;krok
$vsl_ses=$CmdLine[1]  ; co sa meria (vsl alebo ses)
$accum=$CmdLine[4]   ; pocet akumulacii 

AutoItSetOption("WinTitleMatchMode",3)

$isleep=100*$accum+5000  ; cas merania, gate width + gate delay nesmie byt vacsie ako 0.1s!!!!!
$imotor="0"  ;hodnota na motorku na zaciatku
$imotorstep=0.002*$step
For $j=0 to $length step $step
;k motoru pristupuje cez handly
$var=WinList()   ;hleda okno motorku - potrebuje najit handle, jinak dela problemy
for $i=1 to $var[0][0]
	if $var[$i][0]=$motor Then
		$Handle=$var[$i][1]
		;MsgBox(0,"Motor","found")
	EndIf
	;MsgBox(0,"window",$var[$i][0])
Next

AutoItSetOption("WinTitleMatchMode",1)

$var=WinGetTitle($Handle)
WinActivate($Handle,"")
Sleep(500)
MouseClick("left",276,167,1,1)  ;klikne na polohu motorku
Sleep(300)
Send("{BS 6}")  ; vymaze polohu na motorku
Send("{DEL 6}") ; pro jistotu, aby to vymazal
Sleep(300)
$smotor=StringSplit($imotor,".")
StringReplace($smotor,".",",")
Send($smotor[1])
If $smotor[0]>1 Then
	Send("," & $smotor[2])
EndIf	
Sleep(500)
MouseClick("left",193,189,1,1)  ;klikne na start
Sleep(2000)
	
WinActivate($andor,"")
WinWaitActive($andor,"")

Send("{F5}")  ; zacne meranie
Sleep($isleep)
Send("!fe")
WinActivate("Export As...","")
WinWaitActive("Export As...","")
Sleep(150)
While StringLen($j)<4
	$j="0" & $j
WEnd
Send($vsl_ses & $j)
Send("{ENTER 2}")
WinWaitActive($andor)

$imotor=Round($imotor+$imotorstep,2)
Next


