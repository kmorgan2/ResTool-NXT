#Region
#AutoIt3Wrapper_Icon=ResTool NXT.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Compile_both=y
#AutoIt3Wrapper_Res_Description=Automation tool for Residential Technology Helpdesk of Northern Illinois University
#AutoIt3Wrapper_Res_Fileversion=0.1.141023.0
#EndRegion
#Region ###Includes
#RequireAdmin
#include <GUIConstants.au3>
#include <MsgBoxConstants.au3>
#include <Date.au3>
#include <Inet.au3>
#include <IE.au3>
#include <File.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiStatusBar.au3>
#include <ProgressConstants.au3>
#include <StaticConstants.au3>
#include <TabConstants.au3>
#include <WindowsConstants.au3>
#EndRegion ###Includes
#Region ###Program Init and Loop
Opt("GUIOnEventMode", 1)
#Region ### START Koda GUI section ### Form=C:\Users\ResTech\Desktop\ResToolNxt_Form_New.kxf
$form = GUICreate("ResTool NXT: Beta!", 330, 460)
$progbar = GUICtrlCreateProgress(10, 390, 310, 20)
GUICtrlSetStyle($progbar, 8)
GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
$proglabel = GUICtrlCreateLabel("Initializing ResTool", 10, 420, 310, 17, $ss_center)
$tab = GUICtrlCreateTab(0, 0, 330, 380)
$tabsheet1 = GUICtrlCreateTabItem("Virus Removal")
GUICtrlSetState(-1, $gui_show)
$scan = GUICtrlCreateGroup("Antivirus Scans", 10, 42, 95, 310)
$combofix = GUICtrlCreateButton("ComboFix", 20, 67, 75, 25)
$malwarebytes = GUICtrlCreateButton("Malwarebytes", 20, 117, 75, 25)
$eset = GUICtrlCreateButton("ESET", 20, 167, 75, 25)
$spybot = GUICtrlCreateButton("Spybot", 20, 217, 75, 25)
$sas = GUICtrlCreateButton("SAS", 20, 267, 75, 25)
$housecall = GUICtrlCreateButton("HouseCall", 20, 317, 75, 25)
$clean = GUICtrlCreateGroup("System Cleanup", 115, 42, 95, 160)
$ccleaner = GUICtrlCreateButton("CCleaner", 125, 67, 75, 25)
$programs = GUICtrlCreateButton("Programs", 125, 117, 75, 25)
$temp = GUICtrlCreateButton("Temp Files", 123, 167, 75, 25)
$uguu = GUICtrlCreateGroup("Miscellaneous", 220, 42, 95, 160)
$rmscan = GUICtrlCreateButton("Rem. Scans", 230, 67, 75, 25)
$mse = GUICtrlCreateButton("Install MSE", 230, 117, 75, 25)
$ticket = GUICtrlCreateButton("Ticket", 228, 167, 75, 25)
$rootkit = GUICtrlCreateGroup("Rootkit", 115, 207, 96, 115)
$mwbar = GUICtrlCreateButton("MWBAR", 125, 232, 75, 25)
$tdss = GUICtrlCreateButton("TDSS", 125, 282, 75, 25)
$pic1 = GUICtrlCreatePic(@ScriptDir & "\clippy-300x300.jpg", 220, 212, 97, 110)
GUICtrlSetTip(-1, "Hi! I'm Clippy, and I'm your ResTool Assistant. Just kidding. I'm a total easter egg.")
$tabsheet2 = GUICtrlCreateTabItem("OS Troubleshooting")
$net = GUICtrlCreateGroup("Network", 10, 42, 95, 310)
$ipconfig = GUICtrlCreateButton("IPConfig", 20, 67, 75, 25)
$winsock = GUICtrlCreateButton("Winsock", 20, 117, 75, 25)
$hidadapt = GUICtrlCreateButton("Hid. Adapters", 20, 167, 75, 25)
$wifi = GUICtrlCreateButton("NIUwireless", 20, 217, 75, 25)
$speed = GUICtrlCreateButton("Speedtest", 20, 267, 75, 25)
$ncprop = GUICtrlCreateButton("Adapter Prop.", 20, 317, 75, 25)
$system = GUICtrlCreateGroup("System", 115, 42, 95, 310)
$sfc = GUICtrlCreateButton("SFC", 125, 67, 75, 25)
$dism = GUICtrlCreateButton("DISM", 125, 117, 75, 25)
$chkdsk = GUICtrlCreateButton("Disk Check", 125, 267, 75, 25)
$aio = GUICtrlCreateButton("All In One", 125, 167, 75, 25)
$defrag = GUICtrlCreateButton("Defragment", 125, 217, 75, 25)
$devmgr = GUICtrlCreateButton("Device Mgmt.", 125, 317, 75, 25)
$misc = GUICtrlCreateGroup("Miscellaneous", 220, 42, 95, 310)
$rmnac = GUICtrlCreateButton("Remove NAC", 230, 67, 75, 25)
$rmmse = GUICtrlCreateButton("Remove MSE", 230, 117, 75, 25)
$msconfig = GUICtrlCreateButton("MSConfig", 230, 267, 75, 25)
$cpl = GUICtrlCreateButton("Control Panel", 230, 167, 75, 25)
$regedit = GUICtrlCreateButton("Registry", 230, 217, 75, 25)
$nada = GUICtrlCreateButton("Nothing", 230, 317, 75, 25)
GUICtrlSetTip(-1, "This button does nothing!")
$tabsheet3 = GUICtrlCreateTabItem("Removal Tools")
$ni = GUICtrlCreateLabel("NOT IMPLEMENTED", 0, 173, 320, 28, $ss_center)
$tabsheet4 = GUICtrlCreateTabItem("Auto")
$label1 = GUICtrlCreateLabel("DON'T CLICK THIS BUTTON! SRSLY!", 4, 173, 320, 28, $ss_center)
GUICtrlSetFont(-1, 16, 400, 0, "MS Sans Serif")
GUICtrlSetColor(-1, 16711680)
$auto = GUICtrlCreateButton(" FIX EVERYTHING!", 52, 213, 200, 100)
GUICtrlSetTip(-1, "You've been warned!")
If (Not (Random(0, 499, 1) == 217)) Then
	GUICtrlSetState($pic1, 32)
EndIf
$statusbar1 = _GUICtrlStatusBar_Create($form)
Dim $statusbar1_partswidth[3] = [150, 250, -1]
_GUICtrlStatusBar_SetParts($statusbar1, $statusbar1_partswidth)
_GUICtrlStatusBar_SetMinHeight($statusbar1, 20)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###
Global $ip = @IPAddress1
Global $osv = @OSVersion
Global $osa = @OSArch
Global $ticketno = RegRead("HKLM\SOFTWARE\ResTech", "TicketNo")
Global $32progfiledir = @ProgramFilesDir & " (x86)"
_GUICtrlStatusBar_SetText($statusbar1, _readableosv(), 0)
_GUICtrlStatusBar_SetText($statusbar1, $ip, 1)
If ($osa = "X86") Then
	$32progfiledir = @ProgramFilesDir
EndIf
If ($osv = "WIN_81") Then
	GUICtrlSetStyle($combofix, $ws_disabled)
EndIf
If (Not $ticketno = "") Then
	If (_DateDiff("D", RegRead("HKLM\SOFTWARE\ResTech", "LastOpen"), _NowCalc()) >= 7) Then
		If ($idno = MsgBox($mb_yesno, "Ticket Number Found", "This computer has a ticket number recorded already. Is the current ticket for this computer still TT" & $ticketno & "?")) Then
			_setticket()
		Else
			_updatelastopen()
		EndIf
	Else
		_updatelastopen()
	EndIf
Else
	_setticket()
	SetError(0)
EndIf
_GUICtrlStatusBar_SetText($statusbar1, "TT" & $ticketno, 2)
GUISetOnEvent($gui_event_close, "_Close")
GUICtrlSetOnEvent($combofix, "_RunCF")
GUICtrlSetOnEvent($malwarebytes, "_RunMWB")
GUICtrlSetOnEvent($eset, "_RunESET")
GUICtrlSetOnEvent($spybot, "_RunSB")
GUICtrlSetOnEvent($sas, "_RunSAS")
GUICtrlSetOnEvent($housecall, "_RunHC")
GUICtrlSetOnEvent($ccleaner, "_RunCC")
GUICtrlSetOnEvent($programs, "_ProgFeat")
GUICtrlSetOnEvent($temp, "")
GUICtrlSetStyle($temp, $ws_disabled)
GUICtrlSetOnEvent($rmscan, "")
GUICtrlSetStyle($rmscan, $ws_disabled)
GUICtrlSetOnEvent($mse, "_GetMSE")
GUICtrlSetOnEvent($mwbar, "")
GUICtrlSetStyle($mwbar, $ws_disabled)
GUICtrlSetOnEvent($tdss, "")
GUICtrlSetStyle($tdss, $ws_disabled)
GUICtrlSetOnEvent($ipconfig, "_WebReset")
GUICtrlSetOnEvent($winsock, "_Winsock")
GUICtrlSetOnEvent($hidadapt, "_HidNetRemove")
GUICtrlSetOnEvent($wifi, "_WifiProfileAdd")
GUICtrlSetOnEvent($speed, "_SpeedTest")
GUICtrlSetOnEvent($ncprop, "_NetAdapterProperties")
GUICtrlSetOnEvent($sfc, "_SFC")
GUICtrlSetOnEvent($dism, "_DISM")
GUICtrlSetOnEvent($chkdsk, "_Chkdsk")
GUICtrlSetOnEvent($aio, "")
GUICtrlSetStyle($aio, $ws_disabled)
GUICtrlSetOnEvent($defrag, "_Defraggle")
GUICtrlSetOnEvent($devmgr, "_Devmgmt")
GUICtrlSetOnEvent($rmnac, "")
GUICtrlSetStyle($rmnac, $ws_disabled)
GUICtrlSetOnEvent($msconfig, "")
GUICtrlSetStyle($msconfig, $ws_disabled)
GUICtrlSetOnEvent($regedit, "")
GUICtrlSetStyle($regedit, $ws_disabled)
GUICtrlSetOnEvent($nada, "")
GUICtrlSetOnEvent($auto, "_EasterEgg")
GUICtrlSetOnEvent($rmmse, "")
GUICtrlSetStyle($rmmse, $ws_disabled)
GUICtrlSetOnEvent($cpl, "_OpenControlPanel")
GUICtrlSetData($proglabel, "ResTool Ready...")
GUICtrlSetStyle($progbar, 1)
GUICtrlSetData($progbar, 0)
Run(@ComSpec & " /c " & "mkdir " & "C:\ResTech", "", @SW_HIDE)
While 1
	Sleep(100)
WEnd
#EndRegion ###Program Init and Loop
#Region ###Helpers
Func _close()
	Exit
EndFunc   ;==>_close

Func _readableosv()
	Local $r = ""
	Local $s = ""
	If (@OSVersion = "WIN_7") Then
		$r = "Windows 7"
	ElseIf (@OSVersion = "WIN_81") Then
		$r = "Windows 8.1"
	ElseIf (@OSVersion = "WIN_8") Then
		$r = "Windows 8"
	ElseIf (@OSVersion = "WIN_VISTA") Then
		$r = "Windows Vista"
	Else
		$r = "UNSUPPORTED OS"
	EndIf
	If (@OSArch = "X86") Then
		$s = "32-bit"
	ElseIf (@OSVersion = "IA64") Then
		$s = "64-bit Itanium"
	Else
		$s = "64-bit"
	EndIf
	Return $r & " " & $s
EndFunc   ;==>_readableosv

Func _runau3($sfilepath, $sworkingdir = "", $ishowflag = @SW_SHOW, $ioptflag = 0)
	Return Run('"' & @AutoItExe & '" /AutoIt3ExecuteScript "' & $sfilepath & '"', $sworkingdir, $ishowflag, $ioptflag)
EndFunc   ;==>_runau3

Func _setticket()
	Local $newtix = InputBox("Enter Ticket Number", "Please enter the computer's current ticket number:")
	If (Not $newtix = "") Then
		RegWrite("HKLM\SOFTWARE\ResTech", "TicketNo", "REG_SZ", $newtix)
		_updatelastopen()
		$ticketno = $newtix
	Else
		MsgBox(48, "Invalid Ticket", "You entered an invalid ticket number or did not enter a number at all. Please try again.")
		_setticket()
	EndIf
EndFunc   ;==>_setticket

Func _updatelastopen()
	RegWrite("HKLM\SOFTWARE\ResTech", "LastOpen", "REG_SZ", _NowCalc())
EndFunc   ;==>_updatelastopen

Func _appendlog($code, $prog, $msg = "")
	Local $log = StringLeft(@ScriptDir, 3) & "Logs\" & $ticketno & ".txt"
	If (Not FileExists($log)) Then
		_FileCreate($log)
	EndIf
	$logfile = FileOpen($log, 1)
	Local $logentry = _NowCalc()
	Switch $code
		Case 1
			$logentry = $logentry & " Starting  " & $prog & "."
		Case 2
			$logentry = $logentry & " Error occurred while running " & $prog & "."
		Case 3
			$logentry = $logentry & " " & $prog & " generated the following results: " & $msg
		Case 4
			$logentry = $logentry & " " & $prog & " completed successfully."
		Case 5
			$logentry = $logentry & " " & $prog & " generated the following message: " & $msg
	EndSwitch
	If (Not ($msg = "" Or $code = 5 Or $code = 3)) Then
		$logentry = $logentry & " The program also generated the following message: " & $msg
	EndIf
	FileWriteLine($logfile, $logentry)
	FileClose($logfile)
EndFunc   ;==>_appendlog

Func _writeregstats($shortcode, $value)
	RegWrite("HKLM\SOFTWARE\ResTech\AutoStats", $shortcode, "REG_SZ", $value)
EndFunc   ;==>_writeregstats

Func _readregstats($shortcode)
	Return Number(RegRead("HKLM\SOFTWARE\ResTech\AutoStats", $shortcode))
EndFunc   ;==>_readregstats

Func _winwaitnotify($title, $text, $controltitle, $timeout = 0)
	If (WinWait($title, $text, $timeout) = 0) Then
		MsgBox($mb_yesno, "Couldn't find " & $title, "ResTool could not find the window titled " & $title & "." & @CRLF & "Please continue to this window and click " & $controltitle & ".")
	EndIf
EndFunc   ;==>_winwaitnotify

Func _queuestartup()
	FileCreateShortcut(@ScriptFullPath, @StartupCommonDir & "\RT.lnk")
EndFunc   ;==>_queuestartup

Func _waitclick($title, $text, $control, $timeout = 0)
	If (WinWait($title, $text, $timeout) = 0) Then
		If (MsgBox($mb_yesno, "Couldn't find " & $title, "ResTool could not find the window titled " & $title & "." & @CRLF & "Continue to this window and click " & $control & @CRLF & ". Continue execution?") = $idyes) Then
		Else
			Run(@ScriptDir & "\ResTool NXT_x64.exe")
			Exit
		EndIf
	Else
		ControlClick($title, $text, $control, "left", 1)
	EndIf
EndFunc   ;==>_waitclick

#EndRegion ###Helpers
#Region ###Scanners

Func _runcf()
	Local $updated = 0
	_appendlog(1, "ComboFix")
	GUICtrlSetData($proglabel, "ComboFix Running")
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	Run(@ScriptDir & "/Script/Scanners/CF/CF.exe")
	While (Not WinExists("ComboFix"))
		If (WinExists("ComboFix: Disclaimer")) Then
			WinActivate("ComboFix: Disclaimer")
			Send("{ENTER}")
			$updated = 1
		EndIf
		Sleep(100)
	WEnd
	While (WinExists("ComboFix"))
		If (WinExists("ComboFix", "Yes")) Then
			Sleep(100)
			Send("{ENTER}")
		EndIf
		Sleep(100)
	WEnd
	If ($updated = 1) Then
		_runcf()
	Else
		While (Not WinExists("Administrator:  ."))
			Sleep(100)
			If (WinExists("Warning !!")) Then
				WinActivate("Warning !!")
				Send("{ENTER}")
			EndIf
		WEnd
		While (WinExists("Administrator:  ."))
			Sleep(100)
		WEnd
		_appendlog(4, "ComboFix")
	EndIf
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc   ;==>_runcf

Func _runmwb()
	_appendlog(1, "Malwarebytes Anti-Malware")
	GUICtrlSetData($proglabel, "Malwarebytes Running")
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	If (Not FileExists(@ProgramsCommonDir & "\Malwarebytes' Anti-Malware\Malwarebytes Anti-Malware.lnk")) Then
		_installmwb()
	EndIf
	ShellExecute($32progfiledir & "\Malwarebytes' Anti-Malware\mbam.exe")
	While (Not ProcessExists("mbam.exe"))
		Sleep(100)
	WEnd
	Sleep(1000)
	If (WinActive("Malwarebytes Anti-Malware", "outdated")) Then
		Send("{ENTER}")
		Sleep(20000)
		If (WinActive("Malwarebytes Anti-Malware", "An error has occurred")) Then
			_appendlog(3, "Malwarebytes Anti-Malware", "Failed to Update")
			If ($idno == MsgBox($mb_yesno, "ResFail: Net Connection", "The computer is not connected to the proxy and could not update. Continue with MWB scan?")) Then
				Send("{ENTER}")
				Sleep(1000)
				Send("!{F4}")
				Exit
			Else
				Send("{ENTER}")
			EndIf
		ElseIf (WinActive("Malwarebytes Anti-Malware", "latest version")) Then
			ControlClick("Malwarebytes Anti-Malware", "", 2)
			_winwaitnotify("Malwarebytes Anti-Malware", "database was successfully", "OK")
			Send("{ENTER}")
		Else
			_winwaitnotify("Malwarebytes Anti-Malware", "database was successfully", "OK")
			Send("{ENTER}")
		EndIf
	EndIf
	Sleep(10000)
	If (WinActive("Malwarebytes Anti-Malware", "latest version")) Then
		ControlClick("Malwarebytes Anti-Malware", "", 2)
	EndIf
	_winwaitnotify("Malwarebytes Anti-Malware", "Perform", "Full Scan and Start")
	ControlClick("Malwarebytes Anti-Malware", "Perform", 69)
	ControlClick("Malwarebytes Anti-Malware", "Perform", 71)
	_winwaitnotify("Full scan", "", "Drive selection")
	Send("{ENTER}")
	While (Not ProcessExists("notepad.exe"))
		Sleep(100)
	WEnd
	_appendlog(4, "Malwarebytes Anti-Malware")
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc   ;==>_runmwb

Func _installmwb()
	Run(@ScriptDir & "\Script\Scanners\MWB\MWB.exe")
	_winwaitnotify("Select Setup Language", "", "Next")
	Send("{ENTER}")
	_winwaitnotify("Setup - Malwarebytes Anti-Malware", "Welcome", "Next")
	Send("{ENTER}")
	_winwaitnotify("Setup - Malwarebytes Anti-Malware", "License", "Next")
	Send("{TAB}A{ENTER}")
	_winwaitnotify("Setup - Malwarebytes Anti-Malware", "Information", "Next")
	Send("{ENTER}")
	_winwaitnotify("Setup - Malwarebytes Anti-Malware", "Select Destination", "Next")
	Send("{ENTER}")
	Sleep(1000)
	If WinExists("Folder Exists") Then
		Send("{ENTER}")
	EndIf
	_winwaitnotify("Setup - Malwarebytes Anti-Malware", "Select Start", "Next")
	Send("{ENTER}")
	_winwaitnotify("Setup - Malwarebytes Anti-Malware", "Select Additional", "Next")
	Send("{ENTER}")
	_winwaitnotify("Setup - Malwarebytes Anti-Malware", "Ready", "Next")
	Send("{ENTER}")
	Sleep(1000)
	If WinExists("Setup - Malwarebytes Anti-Malware", "Preparing") Then
		MsgBox(48, "Restart Required", "The computer is going to restart in order to complete a previous Malwarebytes install action.")
		Sleep(10000)
		Send("{ENTER}")
	EndIf
	_winwaitnotify("Setup - Malwarebytes Anti-Malware", "Completing", "Finish")
	Send(" ")
	Send("{TAB}")
	Send(" ")
	Send("{TAB}")
	Send(" ")
	Send("{ENTER}")
EndFunc   ;==>_installmwb

Func _runeset()
	_appendlog(1, "ESET Online Scanner")
	GUICtrlSetData($proglabel, "ESET Running")
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	Run(@ScriptDir & "\Script\Scanners\ESET\ESET.exe")
	_winwaitnotify("Terms of use", "", "Accept and Next")
	ControlClick("Terms of use", "", 107)
	ControlClick("Terms of use", "", 105)
	Sleep(10000)
	If WinExists("Downloading ESET Online Scanner...", "Use custom proxy settings") Then
		If (_GetIP() = -1) Then
			While (_GetIP() = -1)
				If (MsgBox(5, "Updating...", "Please plug in the Proxy") == 1) Then
					ControlClick("Downloading ESET Online Scanner...", "", 109)
				EndIf
			WEnd
		Else
			ControlClick("Downloading ESET Online Scanner...", "", 109)
		EndIf
	EndIf
	_winwaitnotify("ESET Online Scanner", "unwanted", "Enable detection and Start")
	ControlClick("ESET Online Scanner", "unwanted", 346)
	ControlClick("ESET Online Scanner", "unwanted", 321)
	WinWaitActive("ESET Online Scanner", "Finished")
	$results = StringSplit(WinGetText("ESET Online Scanner"), @LF, 1)[8]
	ControlClick("ESET Online Scanner", "", 220)
	MsgBox(0, "ESET Results", "ESET found " & $results & " Infected Files.")
	_appendlog(4, "ESET Online Scanner", $results & " Infected Files.")
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc   ;==>_runeset

Func _runsb()
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	If (Not FileExists($32progfiledir & "\Spybot - Search & Destroy 2\SDUpdate.exe")) Then
		GUICtrlSetData($proglabel, "Installing Spybot")
		Run(@ScriptDir & "\Script\Scanners\SB\SB.exe")
		_winwaitnotify("Select Setup Language", "", "Next")
		Send("{ENTER}")
		_winwaitnotify("Setup - Spybot - Search & Destroy", "", "Next")
		Send("{ENTER}")
		_winwaitnotify("Setup - Spybot - Search & Destroy", "Donations", "Next")
		Send("{ENTER}")
		_winwaitnotify("Setup - Spybot - Search & Destroy", "Installation &", "Next")
		Send("{ENTER}")
		_winwaitnotify("Setup - Spybot - Search & Destroy", "License", "Accept and Next")
		Send("{TAB}")
		Send("A")
		Send("{ENTER}")
		_winwaitnotify("Setup - Spybot - Search & Destroy", "Ready", "Next")
		Send("{ENTER}")
		If (WinExists("Setup - Spybot - Search & Destroy", "Preparing to Install")) Then
			If ($idyes = MsgBox($mb_yesno, "Restart required", "Restart to complete SB installation?")) Then
				Send("{ENTER}")
			Else
				Send("N")
				Send("{ENTER}")
				_appendlog(3, "Spybot Search & Destroy", "Required a restart to complete installation")
			EndIf
		EndIf
		_winwaitnotify("Setup - Spybot - Search & Destroy", "Completing", "Finish")
		Send("{ENTER}")
	EndIf
	If (ProcessExists("SDWelcome.exe")) Then
		WinActivate("Start Center (Spybot - Search & Destroy 2.1)")
		Send("!{F4}")
	EndIf
	GUICtrlSetData($proglabel, "Updating Spybot")
	ShellExecute($32progfiledir & "\Spybot - Search & Destroy 2\SDUpdate.exe")
	WinWait("Update (Spybot - Search & Destroy 2.1, administrator privileges)")
	Sleep(1000)
	WinWait("Update (Spybot - Search & Destroy 2.1, administrator privileges)", "Status check complete.")
	Send("{TAB}{ENTER}")
	WinWait("Update (Spybot - Search & Destroy 2.1, administrator privileges)", "+++")
	WinClose("Update (Spybot - Search & Destroy 2.1, administrator privileges)")
	GUICtrlSetData($proglabel, "Spybot Scan Started")
	ShellExecute($32progfiledir & "\Spybot - Search & Destroy 2\SDScan.exe")
	WinWait("System Scan (Spybot - Search & Destroy 2.1, administrator privileges)")
	ControlClick("System Scan (Spybot - Search & Destroy 2.1, administrator privileges)", "", "[CLASS:TButton; INSTANCE:3]")
	WinWait("System Scan (Spybot - Search & Destroy 2.1, administrator privileges)", "Scan took ")
	If (WinExists("System Scan (Spybot - Search & Destroy 2.1, administrator privileges)", "The antispyware scan came up without results.")) Then
		WinClose("System Scan (Spybot - Search & Destroy 2.1, administrator privileges)")
		_appendlog(4, "Spybot Search and Destroy", "No Infected Files")
	Else
		ControlClick("System Scan (Spybot - Search & Destroy 2.1, administrator privileges)", "", "[CLASS:TButton; INSTANCE:3]")
		ControlClick("System Scan (Spybot - Search & Destroy 2.1, administrator privileges)", "Fix selected", "[CLASS:TButton; INSTANCE:1]")
		_appendlog(4, "Spybot Search and Destroy", "Infected Files were Found.")
	EndIf
	GUICtrlSetData($proglabel, "ResTool Ready...")
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
EndFunc   ;==>_runsb

Func _runsas()
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "SUPERAntiSpyware Running")
	If (Not FileExists(@ProgramFilesDir & "\SUPERAntiSpyware\SUPERAntiSpyware.exe")) Then
		Run(@ScriptDir & "\Script\Scanners\SAS\SAS.exe")
		_winwaitnotify("SUPERAntiSpyware Setup", "", "Custom Install")
		If (1 = MsgBox(0, "ResFail: No Automation", "Please click the Custom Install button.")) Then
		EndIf
		_winwaitnotify("SUPERAntiSpyware Free Edition Setup", "", "Next")
		ControlClick("SUPERAntiSpyware Free Edition Setup", "", 1)
		ControlClick("SUPERAntiSpyware Free Edition Setup", "", 1)
		ControlClick("SUPERAntiSpyware Free Edition Setup", "", 1)
		_winwaitnotify("SUPERAntiSpyware Free Edition Setup", "Configuration", "Next")
		ControlClick("SUPERAntiSpyware Free Edition Setup", "", 1)
		ControlClick("SUPERAntiSpyware Free Edition Setup", "", 2)
		_winwaitnotify("SUPERAntiSpyware Professional Trial", "", "Decline")
		ControlClick("SUPERAntiSpyware Professional Trial", "", 2)
	Else
		Run(@ProgramFilesDir & "\SUPERAntiSpyware\SUPERAntiSpyware.exe")
		Run(@ProgramFilesDir & "\SUPERAntiSpyware\SUPERAntiSpyware.exe")
	EndIf
	_appendlog(1, "SUPERAntiSpyware")
	_winwaitnotify("SUPERAntiSpyware Free Edition", "Complete Scan", "Update, Complete Scan and Scan your Computer...")
	ControlClick("SUPERAntiSpyware Free Edition", "", 1011)
	Sleep(60000)
	ControlClick("SUPERAntiSpyware Free Edition", "", 1073)
	ControlClick("SUPERAntiSpyware Free Edition", "", 1090)
	_winwaitnotify("SUPERAntiSpyware Free Edition", "Scan Boost", "High Boost and Start Complete Scan.")
	ControlClick("SUPERAntiSpyware Free Edition", "", "[CLASS:Button; INSTANCE:19]")
	Send("{+}")
	ControlClick("SUPERAntiSpyware Free Edition", "", 1008)
	ControlClick("SUPERAntiSpyware Free Edition", "", 6)
	WinWaitActive("SUPERAntiSpyware Scan Summary")
	_appendlog(4, "SUPERAntiSpyware")
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc   ;==>_runsas

Func _runhc()
	If ($osa = "X86") Then
		_runhc32()
	Else
		_runhc64()
	EndIf
EndFunc   ;==>_runhc

Func _runhc32()
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "HouseCall Running")
	_appendlog(1, "Trend Micro HouseCall")
	Run(@ScriptDir & "/Script/Scanners/HC/HC32.exe")
	_winwaitnotify("HouseCall Download", "", "OK")
	Sleep(1000)
	If (WinActive("Trend Micro HouseCall", "Unable to complete the download.")) Then
		MsgBox(48, "Network Connection", "Plug in the proxy and press OK.")
		_runhc32()
	Else
		WinWaitClose("HouseCall Download")
		Sleep(10000)
		ControlClick("", "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "left", 1, 33, 313)
		Sleep(1000)
		ControlClick("", "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "left", 1, 564, 365)
		Sleep(1000)
		ControlClick("", "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "left", 1, 570, 133)
		MsgBox(0, "Housecall is Evil!", "Housecall won't tell me when it finishes. So I'm going to make your life extra hard and make you close it for me. Thanks!")
		While (ProcessExists("housecall.bin"))
			Sleep(10000)
		WEnd
		GUICtrlSetStyle($progbar, 1)
		GUICtrlSetData($progbar, 0)
		GUICtrlSetData($proglabel, "ResTool Ready...")
		_appendlog(4, "Trend Micro HouseCall")
	EndIf
EndFunc   ;==>_runhc32

Func _runhc64()
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "HouseCall Running")
	_appendlog(1, "Trend Micro HouseCall")
	Run(@ScriptDir & "\Script\Scanners\HC\HC64.exe")
	_winwaitnotify("HouseCall Download", "", "OK")
	Sleep(1000)
	If (WinActive("Trend Micro HouseCall", "Unable to complete the download.")) Then
		MsgBox(48, "Network Connection", "Plug in the proxy and press OK.")
		_runhc32()
	Else
		WinWaitClose("HouseCall Download")
		Sleep(1000)
		$handle = _IEAttach(WinWaitActive("[CLASS:#32770]"), "embedded")
		$accept = _IEGetObjById($handle, "r_eula_0")
		$next = _IEGetObjById($handle, "btn_next")
		_IEAction($accept, "click")
		_IEAction($next, "click")
		Sleep(1000)
		$settings = _IEGetObjById($handle, "Settings")
		$scan = _IEGetObjById($handle, "ScanNow")
		_IEAction($settings, "click")
		$settingswin = WinGetHandle("[CLASS:#32770]")
		$settingshandle = _IEAttach($settingswin, "embedded")
		$full = _IEGetObjById($settingshandle, "type1")
		$ok = _IEGetObjById($settingshandle, "btn_ok")
		_IEAction($full, "click")
		_IEAction($ok, "click")
		WinWaitClose("Settings")
		_IEAction($scan, "click")
		ProcessWaitClose("housecall.bin")
		GUICtrlSetStyle($progbar, 1)
		GUICtrlSetData($progbar, 0)
		GUICtrlSetData($proglabel, "ResTool Ready...")
		_appendlog(4, "Trend Micro HouseCall")
	EndIf
EndFunc   ;==>_runhc64

Func _runcc()
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "CCleaner Running")
	If (Not FileExists(@ProgramFilesDir & "\CCleaner\CCleaner.exe")) Then
		Run(@ScriptDir & "\Script\CC.exe")
		_winwaitnotify("CCleaner Professional Setup", "Welcome", "Next")
		ControlClick("CCleaner Professional Setup", "", 1)
		_winwaitnotify("CCleaner Professional Setup", "Automatically check for updates to CCleaner", "Next")
		Send("{ENTER}")
		If (WinExists("CCleaner Professional Setup", "Free! Google")) Then
			ControlClick("CCleaner Professional Setup", "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "left", 1, 29, 156)
		EndIf
		Send("{ENTER}")
		_winwaitnotify("CCleaner Professional Setup", "Release notes", "Finish", 90)
		ControlClick("CCleaner Professional Setup", "", 1204)
		ControlClick("CCleaner Professional Setup", "", 1)
		_winwaitnotify("CCleaner Professional", "", "Close")
		WinClose("CCleaner Professional")
	Else
		If ($osv = "X86") Then
			Run(@ProgramFilesDir & "\CCleaner\CCleaner.exe")
		Else
			Run(@ProgramFilesDir & "\CCleaner\CCleaner64.exe")
		EndIf
	EndIf
	_winwaitnotify("CCleaner Professional", "", "Close")
	WinClose("CCleaner Professional")
	_appendlog(1, "CCleaner File Cleaner")
	_winwaitnotify("Piriform CCleaner - Professional Edition", "", "Run Cleaner")
	ControlClick("Piriform CCleaner - Professional Edition", "", 1021)
	_winwaitnotify("", "This process will ", "OK")
	Send("{ENTER}")
	_winwaitnotify("Piriform CCleaner - Professional Edition", "CLEANING", "Registry")
	$results = StringSplit(WinGetText("Piriform CCleaner - Professional Edition"), @LF, 1)[8]
	MsgBox(0, "", $results)
	_appendlog(4, "CCleaner File Cleaner", $results)
	Send("G")
	Local $rresults = 1
	$title = "Piriform CCleaner - Professional Edition"
	$text = "Scan for Issues"
	While (Not $rresults = 0)
		_appendlog(1, "CCleaner Registry Scan")
		WinWait($title, $text)
		WinActivate($title, $text)
		ControlClick($title, $text, "Button2")
		Sleep(100)
		While (Not ControlGetText($title, $text, "Button2") = "&Scan for Issues")
			Sleep(100)
		WEnd
		If ControlCommand($title, $text, "Button3", "IsEnabled") Then
			ControlClick($title, $text, "Button3")
			_winwaitnotify("CCleaner", "", "Yes")
			ControlClick("CCleaner", "", 6)
			_winwaitnotify("Save As", "", "Save")
			Sleep(1000)
			Local $ccrfile = StringSplit(WinGetText("Save As"), @LF)[5]
			Send("{ENTER}")
			$text = "Fix All Selected Issues"
			WinWait("", $text)
			Sleep(100)
			ControlClick("", $text, "Button4")
			Sleep(1000)
			While Not ControlCommand("", $text, "Button1", "IsEnabled")
				Sleep(500)
			WEnd
			ControlClick("", $text, "Button5")
			$rresults = UBound(StringSplit(FileRead(@UserProfileDir & "\Documents\" & $ccrfile & ".reg"), "[HK", 1))
			MsgBox(0, "", $rresults & " Registry Entries")
		Else
			MsgBox(0, "", "0 Registry Entries")
			$rresults = 0
		EndIf
		$text = "Scan for Issues"
		_appendlog(4, "CCleaner Registry Cleaner", $rresults & " Entries")
	WEnd
	WinClose("Piriform CCleaner - Professional Edition")
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc   ;==>_runcc

#EndRegion ###Scanners
#Region ###OS

Func _speedtest()
	ShellExecute("iexplore.exe", "http://speedtest.niu.edu")
EndFunc   ;==>_speedtest

Func _netadapterproperties()
	ShellExecute("Ncpa.cpl")
EndFunc   ;==>_netadapterproperties

Func _progfeat()
	ShellExecute("appwiz.cpl")
EndFunc   ;==>_progfeat

Func _devmgmt()
	ShellExecute("devmgmt.msc")
EndFunc   ;==>_devmgmt

Func _defraggle()
	_appendlog(1, "Disk Defragmenter")
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "Defragmenting")
	GUISetState(@SW_SHOW)
	ShellExecute("defrag", "/C /U", "", "", @SW_HIDE)
	Sleep(1000)
	While (ProcessExists("Defrag.exe"))
		Sleep(100)
	WEnd
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	_appendlog(4, "Disk Defragmenter")
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc   ;==>_defraggle

Func _webreset()
	_appendlog(1, "IPConfig Reset")
	GUICtrlSetData($proglabel, "Initializing IP Config Reset")
	ShellExecuteWait("ipconfig", "/release", "", "", @SW_HIDE)
	Sleep(100)
	GUICtrlSetData($progbar, 33)
	GUICtrlSetData($proglabel, "Step 1 of 3 Complete")
	ShellExecuteWait("ipconfig", "/flushdns", "", "", @SW_HIDE)
	Sleep(100)
	GUICtrlSetData($progbar, 66)
	GUICtrlSetData($proglabel, "Step 2 of 3 Complete")
	ShellExecute("ipconfig", "/renew", "", "", @SW_HIDE)
	Sleep(100)
	While (ProcessExists("ipconfig.exe"))
		Sleep(100)
	WEnd
	GUICtrlSetData($progbar, 100)
	GUICtrlSetData($proglabel, "Step 3 of 3 Complete")
	Sleep(1000)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
	_appendlog(4, "IPConfig Reset")
EndFunc   ;==>_webreset

Func _winsock()
	_appendlog(1, "Network Services Reset")
	ShellExecuteWait("netsh", "winsock reset", "", "", @SW_HIDE)
	GUICtrlSetData($proglabel, "Step 1 of 5 Complete")
	GUICtrlSetData($progbar, 20)
	ShellExecuteWait("netsh", "winsock reset catalog", "", "", @SW_HIDE)
	GUICtrlSetData($proglabel, "Step 2 of 5 Complete")
	GUICtrlSetData($progbar, 40)
	ShellExecuteWait("netsh", "interface ip reset", "", "", @SW_HIDE)
	GUICtrlSetData($proglabel, "Step 3 of 5 Complete")
	GUICtrlSetData($progbar, 60)
	ShellExecuteWait("netsh", "interface reset all", "", "", @SW_HIDE)
	GUICtrlSetData($proglabel, "Step 4 of 5 Complete")
	GUICtrlSetData($progbar, 80)
	Select
		Case $osv = "WIN_XP"
			ShellExecuteWait("netsh", "firewall reset", "", "", @SW_HIDE)
		Case Else
			ShellExecuteWait("netsh", "advfirewall reset", "", "", @SW_HIDE)
	EndSelect
	GUICtrlSetData($proglabel, "Step 5 of 5 Complete")
	GUICtrlSetData($progbar, 100)
	Sleep(1000)
	GUICtrlSetData($proglabel, "ResTool Ready...")
	GUICtrlSetData($progbar, 0)
	_appendlog(4, "Network Services Reset")
EndFunc   ;==>_winsock

Func _hidnetremove()
	_appendlog(1, "Hidden Network Adapter Removal")
	Local $devcon = @ScriptDir & "\Script\OOB\64"
	If ($osa == "X86") Then
		$devcon = @ScriptDir & "\Script\OOB\32"
	EndIf
	GUICtrlSetData($proglabel, "Hidden Adapter: Step 1 of 4")
	ShellExecuteWait("devcon", '-r remove "@ROOT\*TEREDO\0000"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*TEREDO\0001"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*TEREDO\0002"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*TEREDO\0003"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*TEREDO\0004"', $devcon)
	GUICtrlSetData($progbar, 25)
	GUICtrlSetData($proglabel, "Hidden Adapter: Step 2 of 4")
	ShellExecuteWait("devcon", '-r remove "@ROOT\*ISATAP\0000"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*ISATAP\0001"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*ISATAP\0002"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*ISATAP\0003"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*ISATAP\0004"', $devcon)
	GUICtrlSetData($progbar, 50)
	GUICtrlSetData($proglabel, "Hidden Adapter: Step 3 of 4")
	ShellExecuteWait("devcon", '-r remove "@ROOT\*TUNMP\0000"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*TUNMP\0001"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*TUNMP\0002"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*TUNMP\0003"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*TUNMP\0004"', $devcon)
	GUICtrlSetData($progbar, 75)
	GUICtrlSetData($proglabel, "Hidden Adapter: Step 4 of 4")
	ShellExecuteWait("devcon", '-r remove "@ROOT\*6TO4\0000"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*6TO4\0001"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*6TO4\0002"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*6TO4\0003"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*6TO4\0004"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*6TO4\0005"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*6TO4\0006"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*6TO4\0007"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*6TO4\0008"', $devcon)
	ShellExecuteWait("devcon", '-r remove "@ROOT\*6TO4\0009"', $devcon)
	GUICtrlSetData($progbar, 100)
	Sleep(1000)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
	_appendlog(4, "Hidden Network Adapter Removal")
EndFunc   ;==>_hidnetremove

Func _chkdsk()
	_appendlog(5, "Check Disk", "This program will run at next reboot. Results unknown.")
	ShellExecuteWait("fsutil", "dirty set c:")
	If (MsgBox($mb_yesno, "Chkdsk queued", "Chkdsk is scheduled to scan the hard drive at next boot. Would you like to restart now?") == $idyes) Then
		ShellExecute("shutdown", "/r /t 0", "", "", @SW_HIDE)
	Else
		MsgBox(0, "Chkdsk queued", "Chkdsk will run the next time the computer restarts.")
	EndIf
	GUICtrlSetData($proglabel, "Chkdsk scan successfully queued.")
EndFunc   ;==>_chkdsk

Func _sfc()
	_appendlog(1, "System File Checker")
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "Running System File Check")
	ShellExecuteWait("C:\Windows\System32\sfc.exe", "/scannow", "", "", @SW_HIDE)
	Sleep(1000)
	While (ProcessExists("sfc.exe"))
		Sleep(100)
	WEnd
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "System File Check Complete")
	$failcount = StringSplit(FileRead(@WindowsDir & "\Logs\CBS\CBS.log"), "[SR] Cannot", 1)[0] - 1
	$repcount = StringSplit(FileRead(@WindowsDir & "\Logs\CBS\CBS.log"), "[SR] Rep", 1)[0] - 1
	If (Not $failcount = 0) Then
		_appendlog(3, "System File Checker", "SFC encountered filesystem corruption it was unable to repair.")
		MsgBox($mb_iconwarning, "SFC Results", "SFC may have failed to repair issues. See KB2965 to verify.")
	Else
		If (Not $repcount = 0) Then
			MsgBox($mb_iconinformation, "SFC Results", "SFC completed successfully and may have made repairs. See KB2965 to verify.")
			_appendlog(4, "System File Checker", "Corruption Repaired.")
		Else
			MsgBox(64, "SFC Resluts", "SFC completed successfully and has not found any system file corruption.")
			_appendlog(4, "System File Checker", "No Corruption Found.")
		EndIf
	EndIf
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc   ;==>_sfc

Func _wifiprofileadd()
	_appendlog(1, "NIUwireless Profile Import")
	ShellExecuteWait("netsh", "wlan delete profile name=NIUwireless", "", "", @SW_HIDE)
	GUICtrlSetData($progbar, 20)
	GUICtrlSetData($proglabel, "Removed old NIUwireless profile")
	ShellExecuteWait("netsh", "wlan delete profile name=NIUwirelessInstructions", "", "", @SW_HIDE)
	GUICtrlSetData($progbar, 40)
	GUICtrlSetData($proglabel, "Removed NIUwirelessInstructions")
	ShellExecuteWait("netsh", "wlan add profile filename=" & @ScriptDir & "\Script\WiFi.xml user=all", "", "", @SW_HIDE)
	GUICtrlSetData($progbar, 60)
	GUICtrlSetData($proglabel, "Added new NIUwireless profile")
	_webreset()
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
	_appendlog(4, "NIUwireless Profile Import")
EndFunc   ;==>_wifiprofileadd

Func _opencontrolpanel()
	GUICtrlSetData($proglabel, "Opening Control Panel")
	ShellExecuteWait("control", "", "")
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc   ;==>_opencontrolpanel

Func _dism()
	_appendlog(1, "Deployment Image Servicing and Management Tool")
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "Running DISM")
	ShellExecuteWait("dism", "/Online /Cleanup-Image /RestoreHealth", "", "", @SW_HIDE)
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	$failcount = StringSplit(FileRead(@WindowsDir & "\Logs\dism\dism.log"), "(restorehealth)", 1)[0] - 1
	If (Not $failcount = 0) Then
		_appendlog(2, "Deployment Image Servicing and Management Tool", "DISM can't run on this computer")
		MsgBox($mb_iconwarning, "DISM", "DISM cannot run on this computer.")
	Else
		_appendlog(4, "Deployment Image Servicing and Management Tool")
		MsgBox(0, "DISM", "DISM Completed without error, but the results are unknown")
	EndIf
	GUICtrlSetData($proglabel, "ResTool Ready.")
EndFunc   ;==>_dism

Func _runaio()
EndFunc   ;==>_runaio

Func _getmse()
	If (Not ($osv = "WIN_81") And Not ($osv = "WIN_8")) Then
		If ($osa == "X86") Then
		Else
		EndIf
	Else
	EndIf
EndFunc   ;==>_getmse

Func _getprint()
EndFunc   ;==>_getprint

#EndRegion ###OS
#Region ###Removal

Func _remnor()
EndFunc   ;==>_remnor

Func _remavg()
EndFunc   ;==>_remavg

Func _remavt()
EndFunc   ;==>_remavt

Func _remkas()
EndFunc   ;==>_remkas

Func _remmca()
EndFunc   ;==>_remmca

#EndRegion ###Removal
#Region ###Other

Func _easteregg()
	Local $ovum[11] = ["Formatting the system drive", "Emailing nasty comments to Kuba", "Pestering ITS", "Insulting the user", "Wiping client's social media", "Reconfiguring proxy", "Installing µTorrent", "Breaking the law", "Reticulating Splines", "Kicking Puppies", "Doing Something Evil"]
	For $i = 0 To 20 Step 1
		Local $uguuu = $ovum[Random(0, 10, 1)]
		GUICtrlSetData($proglabel, $uguuu)
		For $j = 0 To 100 Step 1
			GUICtrlSetData($progbar, $j)
			$sleep = Random(10, 100, 1)
			Sleep($sleep)
		Next
	Next
EndFunc   ;==>_easteregg

#EndRegion ###Other
