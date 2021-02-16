#include <GUIConstants.au3>
#Include <GuiComboBox.au3>
$windowPL = "Andor MCD"
$windowM = "SMCView"
$andor = "C:\Program Files\Andor Technology\AndorMCD.exe"
$andorPath = "C:\Program Files\Andor Technology\"
$smc = "C:\Program Files\SMCVieW\SMCVieW.exe"
$delay = 2600
$digits = 5

Global $time, $acc, $from, $to, $step, $path, $name, $name2, $cbautoname, $nameCombo, $window, $pos, $gui, $handleM, $smcWasUp, $newPos, $GoOn, $measureOn, $cbOpen, $cFilter, $cbShutDown, $rOneMeas, $rMeasSeries, $autowindow, $cbauto, $nameCount, $expWarning
HotKeySet("{ESC}","CancelMeas")
HotKeySet("^m","Measure")
HotKeySet("^o","OpenLaser")
HotKeySet("^c","CloseLaser")

Func CreateAndorSettings()
GUICtrlCreateButton("?",380,5,18,18)
GUICtrlSetOnEvent(-1,"ShowHelp")
GUICtrlCreateLabel("Andor settings",5,5,150,30)
GUICtrlSetFont(-1,10,600)
GUICtrlCreateLabel("Accumulation time: ",5,30)
$time = GUICtrlCreateInput("0.019",120,28,40,20)
GUICtrlCreateLabel("s",165,30)
GUICtrlCreateLabel("Accumulation number: ",5,52)
$acc = GUICtrlCreateInput("100",120,50,40,20)
GUICtrlCreateButton("Set",200,33,50)
GUICtrlSetOnEvent(-1,"SetSettings")

GUICtrlCreateButton("Switch Off",327,25,70)
GUICtrlSetOnEvent(-1,"SwitchOff")
$cbShutDown=GUICtrlCreateCheckbox("Shut down computer",275,50)
GUICtrlSetState($cbShutDown,$GUI_CHECKED)
EndFunc                          

Func CreateSavingPath()
GUICtrlCreateLabel("Saving path",5,80,150,30)
GUICtrlSetFont(-1,10,600)
GUICtrlCreateLabel("Path: ",5,100,150,30)
GUICtrlCreateLabel("Name: ",165,100,150,30)
GUICtrlCreateLabel("Window: ",330,100,150,30)
$path = GUICtrlCreateInput("C:\Mereni\" & @YEAR & "\" & @MON & "\" & @MON & @MDAY & "\",5,120,150,20)
$name = GUICtrlCreateInput("",165,120,75,20)
$name2 = GUICtrlCreateInput("",240,120,60,20)
GUICtrlCreateLabel("-pos-w",303,120,55,20)
GUICtrlSetFont(-1,9,575)
$window = GUICtrlCreateInput("1",350,120,20,20)

$cbautoname = GUICtrlCreateCheckbox("Auto:",117,145,40,20)
GUICtrlSetState(-1,$GUI_CHECKED)
$filenames = FileOpen("filenames.txt", 0)  ;read filenames settings
$nameCombo = GUICtrlCreateCombo("",167,145,80,20)
$s = ""
While 1
    $line = FileReadLine($filenames)
    If @error = -1 Then ExitLoop
	$s = $s & "|" & $line	
Wend
$line = FileReadLine($filenames ,1)   ;first line: default
FileClose($filenames)
GUICtrlSetData(-1, $s, $line)
GUICtrlCreateButton("->",245,145,20,20)
GUICtrlSetOnEvent(-1,"ChangeFilename")
ChangeFilename()                     ;setting the first item
$nameCount = _GUICtrlComboBox_GetCount($nameCombo)

$autowindow = GUICtrlCreateInput("12",325,145,50,20)
$cbauto = GUICtrlCreateCheckbox("Auto:",280,145,40,20)
GUICtrlSetState(-1,$GUI_CHECKED)
GUICtrlCreateButton("->",375,145,20,20)
GUICtrlSetOnEvent(-1,"ChangeWindow")
;GUICtrlCreateLabel("(position will be inserted in front of the window number)",245,140,100,50)
EndFunc

Func CreateMeasurementSettings()
GUICtrlCreateLabel("Measurement settings",5,190,200,20)
GUICtrlSetFont(-1,10,600)
GUICtrlCreateLabel("Laser:",5,215)
GUICtrlCreateButton("&Open",45,208,45)
GUICtrlSetOnEvent(-1,"OpenLaser")
GUICtrlCreateButton("&Close",95,208,45)
GUICtrlSetOnEvent(-1,"CloseLaser")
GUICtrlCreateLabel("Use a special fiter:",170,215)
$cFilter=GUICtrlCreateCombo("",260,210,70,130)
GUICtrlSetData($cFilter,"None|0.9|0.5|0.1|0.01|0.001|0.0001","None")
GUICtrlCreateLabel("One measurement:",5,242)
GUICtrlCreateLabel("New position:",5,255)
GUICtrlCreateButton("Go to:",5,270,45)
GUICtrlSetOnEvent(-1,"GoToNewPosition")
$newPos = GUICtrlCreateInput("0",60,270,40,20)
$cbOpen=GUICtrlCreateCheckbox("Leave open after measurement",150,333)
GUICtrlSetState(-1,$GUI_CHECKED)
GUICtrlCreateButton("&Measure",200,310,70)
GUICtrlSetOnEvent(-1,"Measure")
GUICtrlCreateGroup ("", 150, 260, 130, 45)
$rOneMeas = GUICtrlCreateRadio ("One measurement", 155, 270, 120, 15)
GUICtrlSetState(-1,$GUI_CHECKED)
$rMeasSeries = GUICtrlCreateRadio ("Mesurement series", 155, 285, 120, 15)
GUICtrlCreateGroup ("",-99,-99,1,1)  ;close group
GUICtrlCreateButton("Export",315,310,70)
GUICtrlSetOnEvent(-1,"OneExport")
GUICtrlCreateLabel("Homogeneity series:",5,302)
GUICtrlCreateLabel("From: ",5,315)
GUICtrlCreateLabel("To: ",50,315)
GUICtrlCreateLabel("Step: ",95,315)
$from = GUICtrlCreateInput("2000",5,330,40,20)
$to = GUICtrlCreateInput("8000",50,330,40,20)
$step = GUICtrlCreateInput("50",95,330,40,20)
;GUICtrlCreateButton("Measure",200,330,70)
;GUICtrlSetOnEvent(-1,"Measure")
GUICtrlCreateLabel("(measurement of a series can be stopped by hitting the Esc key)",170,350,160,30)
$expWarning=GUICtrlCreateLabel("",340,350,50,40)
EndFunc

Func CancelMeas()
	if $MeasureOn Then
		$GoOn=False
	    MsgBox(0,"Esc hit: Stopping the measurement","The measurement will be stopped after finishing the last acquisition.",3)
	EndIf	
EndFunc	

Func ShowHelp()
	Run("notepad.exe _readme.txt")
EndFunc	

Func Close()
	Exit
EndFunc	

Func SetDelay()
	$delay = (GUICtrlRead($time)*GUICtrlRead($acc))*1000+700
EndFunc

Func MeasSave()   ;zmeri a ulozi
	Opt("WinTitleMatchMode",2)  ;match any substring in the title
	WinActivate($windowPL)
    WinWaitActive($windowPL)
    Send("{F5}") ;mereni
	ControlClick($windowPL,"","OWL_Notetab1","left",1,23,15)   ;ukaz mereni
    Sleep($delay)
    Send("!fa") ;save as
    WinWaitActive("Uložit jako")
	$s = $pos
	While StringLen($s)<$digits
		$s = "0" & $s
	WEnd	
	$file = GUICtrlRead($path)
	If StringRight($file,1) <> "\" Then $file = $file & "\"
	$s1 = GUICtrlRead($name)	
	If StringLen($s1)>0 Then $file=$file & $s1 & "-"
	$s1 = GUICtrlRead($name2)	
	If StringLen($s1)>0 Then $file=$file & $s1 & "-"	
    $file = $file & $s & "-w" & GUICtrlRead($window)
    Send($file)
    Send("{ENTER}")
    WinWaitActive($windowPL)
	Send("^{F4}")
    WinWaitActive($windowPL)
	Opt("WinTitleMatchMode",1) ;match title from start
EndFunc	

Func BckgdMeasSave()  ;zmeri pozadi, signal a ulozi
	GoToKolo(7)                   ;close
	WinActivate($windowPL)
    WinWaitActive($windowPL)
	;WinWaitActive($windowPL)
    Send("^b") ;pozadi
	ControlClick($windowPL,"","OWL_Notetab1","left",1,53,15)   ;ukaz pozadi
    Sleep($delay)
	Send("!wh")
	WinWaitActive($windowPL)
	OpenFilter()   ;otevrit
    MeasSave()
	;ExportWarning(True)
EndFunc

Func SetSettings()
	WinActivate($windowPL)
	WinWaitActive($windowPL)
	Send("!at")
	WinWaitActive("Setup Acquisition")
	ControlClick("Setup Acquisition","","Button2")  ;Single Scan: resetting cosmic rays
	WinWaitActive("Setup Acquisition")
	Send("{ENTER}")
	WinWaitActive($windowPL)
	Send("!at")
	WinWaitActive("Setup Acquisition")
 	ControlClick("Setup Acquisition","","Button4")  ;Accumulate
	WinWaitActive("Setup Acquisition")
	ControlClick("Setup Acquisition","","Edit1")
	WinWaitActive("Setup Acquisition")
	Send("+{HOME}")
	WinWaitActive("Setup Acquisition")
	Send("{DELETE}")
	WinWaitActive("Setup Acquisition")
	$s = GUICtrlRead($time)
	Send($s)
	WinWaitActive("Setup Acquisition")
	ControlClick("Setup Acquisition","","Edit3")
	WinWaitActive("Setup Acquisition")
	Send("+{HOME}")
    WinWaitActive("Setup Acquisition")
	Send("{DELETE}")
	WinWaitActive("Setup Acquisition")
	$s = GUICtrlRead($acc)
	Send($s)
	Sleep(300)
	ControlClick("Setup Acquisition","","Button23")  ;remove cosmic rays
	Sleep(500)
	WinWaitActive("Setup Acquisition")
	Send("{ENTER}")
	WinWaitActive($windowPL)
	Send("!ad")
	WinWaitActive("Data Type")
	Sleep(300)
	Send("b")                                       ;background corrected
	Sleep(500)
	Send("{ENTER}")
	WinWaitActive($windowPL)
	SetDelay()
	WinActivate($gui)
EndFunc	

Func GoToX($pos,$st)                ;posune na pozici x, $st je krok v mikronech, zadava se kvuli cekani na konci
	If $pos<0 Then 
		$pos=0
		MsgBox(0,"Error:","Entered position too low. Resetting to 0.",10)
		GUICtrlSetData($NewPos,"0")
	EndIf	
	If $pos>10000 Then  
		$pos=10000
		MsgBox(0,"Error:","Entered position too high. Resetting to 10000.",10)
		GUICtrlSetData($NewPos,"10000")
	EndIf	
	WinActivate($handleM)
	Sleep(300)
	MouseClick("left",300,110,1,0)  ;enter position edit
	WinWaitActive($handleM)
	Send("{END}")
	WinWaitActive($handleM)
	Send("+{HOME}")
	WinWaitActive($handleM)
	Send("{DEL}")
	WinWaitActive($handleM)
		Send($pos)                       ;enter position
	WinWaitActive($handleM)
	MouseClick("left",180,135,1,0)
	Sleep($st*1.3+500)
EndFunc

Func GoToNewPosition()
	$pom = GUICtrlRead($NewPos)
	GoToX($pom,1500)
	WinActivate($gui)
EndFunc	
	

Func GoToKolo($pos)
	WinActivate($handleM)
	Sleep(300)
	MouseClick("left",300,245,1,0)  ;enter kolo edit
	WinWaitActive($handleM)
	Send("{END}")
	WinWaitActive($handleM)
	Send("+{HOME}")
	WinWaitActive($handleM)
	Send("{DEL}")
	WinWaitActive($handleM)
	Send($pos)                       ;enter number
	WinWaitActive($handleM)
	MouseClick("left",180,265,1,0)
    Sleep(2000)
EndFunc

Func OpenFilter()
		$pom = GUICtrlRead($cFilter)
	Select
	case $pom=="None" 
			$pom1 = 0		
	case $pom=="0.9" 
			$pom1 = 6
	case $pom=="0.5" 
			$pom1 = 5		
	case $pom=="0.1" 
			$pom1 = 4		
	case $pom=="0.01" 
			$pom1 = 3		
	case $pom=="0.001" 
			$pom1 = 2		
	case $pom=="0.0001" 
			$pom1 = 1					
	EndSelect	
	GoToKolo($pom1)
EndFunc	

Func OpenLaser()
    OpenFilter()
    WinActivate($gui)	
EndFunc	

Func CloseLaser()
	GoToKolo(7)
	WinActivate($gui)
EndFunc	

Func OneExport()
	Export()
	WinActivate($gui)
EndFunc	

Func Measure()
	Select 
	Case GUICtrlRead($rOneMeas)=$GUI_CHECKED 			;one measurement
		WinMove($handleM,"",0,0)
		CreateDir()
		WinMove($handleM,"",0,0)
		SetDelay()
		Sleep(150)
		$pos = GUICtrlRead($newPos)
		BckgdMeasSave()
		Sleep(300)
	Case Else	                                         ;measurement series
		$GoOn=True
		$MeasureOn=True
		CreateDir()
		WinMove($handleM,"",0,0)
		SetDelay()
		Sleep(150)
		$pos = GUICtrlRead($from)
		GoToX($pos,$pos)
		BckgdMeasSave()
		$st = GUICtrlRead($step)
		$end = GUICtrlRead($to)
		While ($st>0) and ($pos < $end)
			$pos = $pos + $st
			GoToX($pos,$st)
			MeasSave()
			If Not $GoOn Then
				ExitLoop
			EndIf	
		WEnd
		Export()
		;GoToKolo(7)   ;zavrit
		GoToX(0,$pos)
		$MeasureOn=False
	EndSelect
	If (GUICtrlRead($cbOpen)=$GUI_UNCHECKED) Then GoToKolo(7)   
	$last = False
	$showMessage = False
	WinActivate($gui)
	If (GUICtrlRead($cbauto)=$GUI_CHECKED) Then ;zmenit okno
		WinWaitActive("Fotoluminiscence: ")
		ChangeWindow()
		$s = GUICtrlRead($window)
		$s1 = GUICtrlRead($autowindow)
		$s1 = StringLeft($s1,1)
		$showMessage=True	
		If ($s==$s1) Then                  ;okno 1
			If (GUICtrlRead($cbautoname)=$GUI_CHECKED) Then 
				$last=IncrementFilename() ;meneni souboru	
			Else 
				$showMessage=False					
			EndIf
		EndIf	
		If $showMessage Then
			If $last Then
				GUICtrlSetData($name,"")
				MsgBox(0,"Measurements finished:","List finished. Exporting...",10)
				Export()
			Else	
				$s1 = GUICtrlRead($name)	
				$s="sample " & $s1 & ", window " & $s   ;comment: okno + sampleM
				MsgBox(0,"New measurement:",$s,10)
			EndIf	
		EndIf
		
	EndIf	
	WinActivate($gui)
EndFunc	

Func StartAndor()
	MsgBox(0,"Start Andor","Andor is starting up",1)
	If Not WinExists($windowPL) Then
		Run($andor,$andorPath)
		WinWaitActive($windowPL)	
	EndIf	
	WinActivate($windowPL)
	WinWaitActive($windowPL)
	ControlClick($windowPL,"","OWL_Gauge1")     ;chlazeni CCD
	WinWaitActive("Temperature")
	ControlClick("Temperature","","Button3")
	ControlClick("Temperature","","Edit1")
	Send("{END}")
	Send("+{HOME}")
	Send("{DEL}")
 	Send("-40")
	WinWaitActive("Temperature")
	Send("{ENTER}")
	WinWaitActive($windowPL)
EndFunc	

Func StartSMC()
$smcWasUp=ProcessExists("SMCView.exe")	
If Not $smcWasUp Then
 Run($smc)
 MsgBox(0,"Start SMC","SMC is starting up, please wait.",11.5)
EndIf 
$handles = WinList ()
$i = 1
While (StringInStr($handles [$i][0],$windowM) = 0) and ($i<$handles[0][0])
	$i = $i +1
WEnd	
If $handles[$i][0] = $windowM  Then 
	$handleM = $handles[$i][1]
Else 
	MsgBox(0,"Error:","SMC sofware not found.")
    Exit	
EndIf	
WinMove($handleM,"",0,0)
EndFunc

Func InitiateSMC()
WinActivate($handleM)
Sleep(300)
MouseClick("left",850,135,1,0)    ;setup X
Sleep(500)
MouseClick("left",500,340,1,0)
WinWaitActive("Choose")
ControlClick("Choose","","Edit1")
Sleep(300)
Send("x-posuv.usm")
Send("{ENTER}")
Sleep(500)
WinWaitClose("Choose")
Sleep(300)
MouseClick("left",805,340,1,0)
Sleep(500)
MouseClick("left",180,135,1,0)
Sleep(500)
MouseClick("left",180,135,1,0)
Sleep(300)
MouseClick("left",850,250,1,0)    ;setup kolo
Sleep(500)
MouseClick("left",500,340,1,0)
WinWaitActive("Choose")
ControlClick("Choose","","Edit1")
Sleep(300)
Send("kolo.usm")
Send("{ENTER}")
Sleep(500)
WinWaitClose("Choose")
Sleep(300)
MouseClick("left",805,340,1,0)
Sleep(500)
MouseClick("left",180,265,1,0)
EndFunc	

Func Export()
	Opt("MouseCoordMode",0) ;relative to active window
	WinActivate($windowPL)
	WinWaitActive($windowPL)
	Send("!f")
	Send("{DOWN 7}")
	Send("{ENTER}")
	WinWaitActive("Batch Conversion")
	Opt("MouseClickDragDelay",1000)
	MouseClickDrag("left",270,90,10,220,10)
    Send("!o")
	WinWaitActive("Ascii Separator")
	ControlClick("Ascii Separator","","TRadioButton1")
	Send("{ENTER}")
	Sleep(1000)
	If WinExists("File Already Exists") Then
		Send("!a") 
	EndIf	
	Sleep(5000)
	WinWaitActive("Progress")
	Send("!c")
	Sleep(1000)
	;MouseClick("left")
	WinActivate("Batch Conversion")
	WinWaitActive("Batch Conversion")
	Send("!c")
	;ExportWarning(False)
	WinWaitActive($windowPL)
	Opt("MouseCoordMode",1) ;relative to screen (absolute)
EndFunc	

Func CreateDir()
	$s = GUICtrlRead($path)
	If not FileExists($s) Then DirCreate($s)
EndFunc	

Func SwitchOff()
	If WinExists($windowPL) Then
	 WinActivate($windowPL)
	 WinWaitActive($windowPL)
	 ControlClick($windowPL,"","OWL_Gauge1")     ;chlazeni CCD
	 WinWaitActive("Temperature")
	 ControlClick("Temperature","","Button4")
	 WinWaitActive("Temperature")
	 Send("{ENTER}")
	 WinWaitActive($windowPL)
	EndIf 
	
	If WinExists($handleM) Then
	 WinActivate($handleM)
	 WinWaitActive($handleM)
	 GoToKolo(7)
	 GoToX(0,10000)
	 MsgBox(0,"Switching off:","Closing SMC",2)
	 MouseClick("left",100,320,1,0)
	 Sleep(500)
	 MouseClick("left",125,245,1,0)
	 Sleep(500)
	 MouseClick("left",930,118)
	 Sleep(500)
    EndIf
 
    If WinExists($windowPL) Then
	 WinActivate($windowPL)
	 WinWaitActive($windowPL)
	 $col = PixelGetColor(1180,955)
     While $col=16711680             ;cervena barva u vypinani
	 	Sleep(500)
	 	MsgBox(0,"Switching off:","Waiting for Andor to heat up.",4)
	 	$col = PixelGetColor(1180,955)
	 WEnd
	WinActivate($windowPL)
	WinWaitActive($windowPL)
    WinClose($windowPL)
	Sleep(700)
	Send("{RIGHT}")
	Sleep(300)
	Send("{ENTER}")
    EndIf
 
    If (GUICtrlRead($cbShutDown)==$GUI_CHECKED) Then
		Shutdown(1)
	Else 
		Exit 
	EndIf	
EndFunc

Func ChangeWindow()
	$a=GUICtrlRead($autowindow)
	$b=GUICtrlRead($window)
	$n=StringInStr($a,$b)
	$n=$n+1
	If $n>StringLen($a) Then $n=1
    $a=StringLeft($a,$n)          ;last n digits
	$a=StringRight($a,1)          ; very last digit is the nth
	GUICtrlSetData($window,$a)	
EndFunc	

Func ChangeFilename()
	$a=GUICtrlRead($nameCombo)
	GUICtrlSetData($name,$a)
EndFunc	

Func IncrementFilename()
	$s=GUICtrlRead($name)
	$i=_GUICtrlComboBox_FindString($nameCombo,$s,-1)   ;selected item: first item = 0
	$i=$i+1
	$i=_GUICtrlComboBox_SetCurSel($nameCombo,$i)
	ChangeFilename()
 	If ($i==-1) Then                 ;last, i.e. next doesnt exist: return True
		Return True
	Else 
		Return False
	EndIf	
EndFunc	


Func ExportWarning($t)
If $t Then
	GUICtrlSetData($expWarning,"not exported")
	GUICtrlSetColor(-1,$CLR_RED)
Else 	
	GUICtrlSetData($expWarning,"export OK")
	GUICtrlSetColor(-1,$CLR_RED)
EndIf	
EndFunc	

Opt("WinTitleMatchMode",1) ;match title from start
Opt("CaretCoordMode",0) ;coord relative to active window

Opt("GUIOnEventMode", 1)  
$gui = GUICreate("Fotoluminiscence: PL v1j")
GUISetOnEvent($GUI_EVENT_CLOSE, "Close")

$GoOn = True
$measureOn = False

StartSMC()
CreateAndorSettings()
CreateSavingPath()
CreateMeasurementSettings()

Sleep(5000)
StartAndor()
SetSettings()
If Not $smcWasUp Then
	InitiateSMC()
EndIf	

GUISetState(@SW_SHOW)

While 1
  Sleep(100)  ; Idle around
WEnd

