
$CCD="C7557-01"
$motor="SMCView"
$spectro="M266-4"

$length=$CmdLine[2]    ;dlzka pruzku v mikrometroch
$step=$CmdLine[3]        ;krok
$vsl_ses=$CmdLine[1]  ; co sa meria (vsl alebo ses)
$integrtime=$CmdLine[4]   ; pocet akumulacii 

AutoItSetOption("WinTitleMatchMode",3)  ;exact title match


$isleep=100*$integrtime+5000  ; cas mereni, gate width + gate delay nesmi byt vetsi nez 0.1s!!!!!
$imotor="0"  ;hodnota na motorku na zacatku
$imotorstep=0.002*$step

For $j=0 to $length step $step  ;k motoru pristupuje cez handly
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
MouseClick("left",276,17,1,1)  ;klikne na polohu motorku
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
MouseClick("left",170,42,1,1)  ;klikne na start
Sleep(2000)

WinActivate($CCD,"")
WinWaitActive($CCD,"")

Send("{F1}")  ; zacne meranie
Sleep($isleep)
Send("!fe")
WinActivate("Save As...","")
WinWaitActive("Save As...","")
Sleep(150)
While StringLen($j)<4
	$j="0" & $j
WEnd
Send($vsl_ses & $j)
Send("{ENTER 2}")
WinWaitActive($CCD)

$imotor=Round($imotor+$imotorstep,2)
Next

