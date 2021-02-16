; -------------------
; | VSLSESScript v2 |
; | By Leon Webbers |
; -------------------
; Based on the original version by Katerina Dohnalova

AutoItSetOption("SendKeyDelay", 10) ; increase this value if save operation fails (especially the first time)
; global variable definition
Const $CCD = "C7557-01" ; window title of CCD app
Const $motor = "[TITLE:SMCVieW; CLASS:LVDChild]" ; window title and class of motor app (needs to be specific, because there are actually two visible windows with the title 'SMCVieW'. Also note the case sensitivity)
Const $usage = "Usage: vslsesscript prefix stripeLength stepSize integrationTime motorRef [startAt]" & @CRLF & _
	@TAB & "prefix - filename prefix (typically VSL or SES)" & @CRLF & _
	@TAB & "stripeLength - length of the stripe (typically 4000 um)" & @CRLF & _
	@TAB & "stepSize - increase of slit width (VSL) or shift of excitation spot (SES) per iteration (typically 100 um)" & @CRLF & _
	@TAB & "integrationTime - duration of each iteration, i.e. the value used for 'Start Interval' in the CCD app" & @CRLF & _
	@TAB & "motorRef - the motor value that should be treated as zero" & @CRLF & _
	@TAB & "startAt - optional parameter specifying starting position (useful when a previous cycle was interrupted for some reason, and you want to continue where you left off)"
; the following coordinates are all relative to their respective window edges
Const $btnSTARTx = 180 ; x position of the START button in the motor app window
Const $btnSTARTy = 140 ; y position of the same
Const $nudDestx = 295 ; x position of the Destination numericupdown in the motor app window
Const $nudDesty = 115 ; y position of the same
Const $lblPowerx = 755 ; x position of motor power label
Const $lblPowery = 110 ; y position the same
Const $btnOnOffx = 740 ; x position of ON/OFF switch
Const $btnOnOffy = 145 ; y position of the same
Const $txtSIx = 150 ;  x position of Start Interval textbox
Const $txtSIy = 80; y position of the same

; perform a few checks on arguments to enhance user experience
If $CmdLine[0] < 5 Then
	MsgBox(16, "Error", "Incorrect number of arguments." & @CRLF & @CRLF & $usage)
	Exit
EndIf

; retrieve arguments
Dim $prefix = $CmdLine[1]
Dim $length = $CmdLine[2]
Dim $step = $CmdLine[3]
Dim $integrtime = $CmdLine[4]
Dim $zero = $CmdLine[5]
Dim $startAt = 0
If $CmdLine[0] = 6 Then
	$startAt = $CmdLine[6]
EndIf

; additional checks
If (Number($length) < 1) Or (Number($step) < 1) Or (Number($integrtime) < 1) Then
	MsgBox(16, "Error", "Invalid parameter. Should be an integral number greater than zero." & @CRLF & @CRLF & $usage)
	Exit
EndIf
If Not WinExists($CCD) Then
	MsgBox(16, "Error", "CCD app cannot be found. Please start the C7557-01 Application Software and try again.")
	; todo: auto start CCD app
	Exit
EndIf
If Not WinExists($motor) Then
	MsgBox(16, "Error", "Motor app cannot be found. Please start the SMCVieW application and try again.")
	; todo: auto start motor app
	Exit
EndIf
If MsgBox(36, "Confirm configuration", "Have you remembered to change the slit configuration from SES to VSL or vice versa?") = 7 Then
	Exit
EndIf

; move motor app window to top left position
WinMove($motor, "", 0, 0)
Dim $size = WinGetPos($motor)
; move CCD app window below motor app window to make sure they do not overlap
WinMove($CCD, "", 0, $size[3])

; ask in which direction data should be saved (remembers previously used location stored in ini file)
Dim $folder = FileSelectFolder("Choose a location to save the data files.", "", 3, IniRead("vslsesscript.ini", "settings", "folder", ""))
If $folder = "" Then
	Exit ; user pressed cancel
EndIf
IniWrite("vslsesscript.ini", "settings", "folder", $folder) ; save new location

; make sure 'Start Interval' matches integrationtime parameter
WinActivate($CCD)
;MouseClick("left", $txtSIx, $size[3] + $txtSIy)
; if the following does not work properly, uncomment the previous line and replace all instances of ControlSend below with their equivalent Send calls
ControlFocus($CCD, "", 1000) ; focus edit control
ControlSend($CCD, "", 1000, "{END}")
ControlSend($CCD, "", 1000, "+{HOME}")
ControlSend($CCD, "", 1000, $integrtime)

; define some more variables
Dim $lengthtotal = $length + $zero
Dim $totalTime = ($integrtime + 6000) * $length / (1000 * $step)
Dim $jStart = $zero + $startAt

; show progress window in lower right corner
Dim $tb = WinGetPos("[CLASS:Shell_TrayWnd]") ; get taskbar size
ProgressOn("Script Progress", "Script Progress", "", $tb[2] - 306, $tb[1] - 126, 1)

For $j=$jStart to $lengthtotal step $step
	$jAbs = $j - $zero ; for later reference
	; input next position in motor app
	WinActivate($motor)
	MouseClick("left", $nudDestx, $nudDesty, 1, 0)
	Sleep(200)
	Send("{END}")
	Send("+{HOME}")
	Send($j)
	Sleep(100)
	; if power label is green, the power is switched off, so turn it back on
	If Hex(PixelGetColor($lblPowerx, $lblPowery), 6) = "00FF00" Then
		MouseClick("left", $btnOnOffx, $btnOnOffy, 1, 0) ; turn power on
		Sleep(200)
	EndIf
	; start moving motor to next position
	MouseClick("left", $btnSTARTx, $btnSTARTy, 1, 0)
	Sleep(500)
	; wait unill it's finished (power label will turn yellow)
	; Note: AutoIt prior to v3.0.102 uses BGR format, so swap the FF with the 00 when using this script on those versions
	While Hex(PixelGetColor($lblPowerx, $lblPowery), 6) <> "FFFF00"
		Sleep(500)
	WEnd
	; activate CCD app
	WinActivate($CCD)
	WinWaitActive($CCD)
	Sleep(500)
	; for some reason none of these are working:
	;WinMenuSelectItem($CCD, "", "&Measurement", "&Start")
	;ControlClick($CCD, "", "[CLASS:Button; INSTANCE:7]")
	;ControlFocus($CCD, "", 32773)
	;ControlSend($CCD, "", 32773, "{ENTER}")
	;ControlSend($CCD, "", 32773, "{F1}")
	; so use this: less robust, but as long as the CCD window is properly activated, it should work OK
	Send("{F1}")
	; wait for the measurement to complete
	Sleep($integrtime + 200)
	; wait until the CCD window responds again and the measuring label is no longer visible (in case it was not done already)
	; While SendMessageTimeout(mainWindowHandle, 0, IntPtr.Zero, IntPtr.Zero, 2, 0x1388, out ptr2) <> IntPtr.Zero
	While (ControlCommand($CCD, "", 1060, "IsVisible", "") = 1) Or (@error = 1)
		Sleep(200)
	WEnd
	; now make doubly sure the CCD app window is still active
	WinActivate($CCD)
	WinWaitActive($CCD)
	; save data
	WinMenuSelectItem($CCD, "", "&File", "Save &As...")
	WinWaitActive("Save As")
	Sleep(500)
	; enter filename and path. By using the .txt extension, it will automatically choose text format for output (so no need to select the text filetype)
	If $jAbs > 0 Then
		ControlSend("Save As", "", 1152, $prefix & $jAbs & "um.txt{ENTER}")
	Else
		ControlSend("Save As", "", 1152, $folder & "\" & $prefix & $jAbs & "um.txt{ENTER}")
	EndIf
	;ControlClick("Save As", "", 1) ; didn't work, so appended {ENTER} to the text to be sent to the filename edit
	; update progress window
	ProgressSet(Round($jAbs  / $length * 100), "Time remaining: ~" & Round(($totalTime - (($integrtime + 6000) * $jAbs / (1000 * $step))) / 60, 2) & " min.")
	WinWaitClose("Save As")
Next

; reset motor to starting position
WinActivate($motor)
MouseClick("left", $nudDestx, $nudDesty, 1, 0)
Send("{END}")
Send("+{HOME}")
Send($zero)
Sleep(100)
; if power label is green, the power is switched off, so turn it back on
If Hex(PixelGetColor($lblPowerx, $lblPowery), 6) = "00FF00" Then
	MouseClick("left", $btnOnOffx, $btnOnOffy, 1, 0) ; turn power on
EndIf
MouseClick("left", $btnSTARTx, $btnSTARTy, 1, 0)

; hide progress window
ProgressOff()

; done!