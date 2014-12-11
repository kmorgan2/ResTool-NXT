#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ResTool NXT.ico
#AutoIt3Wrapper_Outfile_x64=J:\ResTool NXT_x64.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_Res_Description=Automation tool for Residential Technology Helpdesk of Northern Illinois University
#AutoIt3Wrapper_Res_Fileversion=0.1.141211.0
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region ###Includes and Compiler Directives
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
#include <ColorConstants.au3>
#EndRegion ###Includes
#Region ###Program Init and Loop
Opt("GUIOnEventMode", 1)
#Region ### START Koda GUI section ### Form=C:\Users\ResTech\Desktop\ResToolNxt_Form_New.kxf
$form = GUICreate("ResTool NXT: Beta!", 330, 460)
$progbar = GUICtrlCreateProgress(10, 390, 310, 20)
;set the progress bar to loading while we are waiting for program to init
GUICtrlSetStyle($progbar, 8)
GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
$proglabel = GUICtrlCreateLabel("Initializing ResTool", 10, 420, 310, 17, $ss_center)
$tab = GUICtrlCreateTab(0, 0, 330, 380)
$tabsheet1 = GUICtrlCreateTabItem("Virus Removal")
GUICtrlSetState(-1, $gui_show) ;define the default tab
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
$print = GUICtrlCreateButton("Pharos", 230, 317, 75, 25)
$tabsheet3 = GUICtrlCreateTabItem("Removal Tools")
$ni = GUICtrlCreateLabel("NOT IMPLEMENTED", 0, 173, 320, 28, $ss_center)
$tabsheet4 = GUICtrlCreateTabItem("Auto")
$label1 = GUICtrlCreateLabel("DON'T CLICK THIS BUTTON! SRSLY!", 4, 173, 320, 28, $ss_center)
GUICtrlSetFont(-1, 16, 400, 0, "MS Sans Serif")
GUICtrlSetColor(-1, 16711680)
$auto = GUICtrlCreateButton(" FIX EVERYTHING!", 52, 213, 200, 100)
GUICtrlSetTip(-1, "You've been warned!")
$statusbar1 = _GUICtrlStatusBar_Create($form)
Dim $statusbar1_partswidth[3] = [150, 250, -1]
_GUICtrlStatusBar_SetParts($statusbar1, $statusbar1_partswidth)
_GUICtrlStatusBar_SetMinHeight($statusbar1, 20)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###
;define a couple things that are freuquently used
Global $ip = @IPAddress1
Global $osv = @OSVersion
Global $osa = @OSArch
Global $ticketno = RegRead("HKLM\SOFTWARE\ResTech", "TicketNo")
;define the directory where 32-bit installs reside
Global $32progfiledir = @ProgramFilesDir & " (x86)"
If ($osa = "X86") Then
	$32progfiledir = @ProgramFilesDir
EndIf
;update or define the ticket number if necessary
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
;Put things on the status bar
_GUICtrlStatusBar_SetText($statusbar1, _readableosv(), 0)
_GUICtrlStatusBar_SetText($statusbar1, $ip, 1)
_GUICtrlStatusBar_SetText($statusbar1, "TT" & $ticketno, 2)
;if Auto MSConfig has been toggled to Safe Mode, make the button red to indicate as such
If(_readregstats("MSCFG") == 1) Then
	GUICtrlSetColor($msconfig, $COLOR_RED)
EndIf
;Remove Startup shortcut to ResTool if it exists
If (FileExists(@StartupCommonDir & "\RT.lnk")) Then
	FileDelete(@StartupCommonDir & "\RT.lnk")
EndIf
;Attach clicks to functions and disable nonworking fxns
GUISetOnEvent($gui_event_close, "_Close")
GUICtrlSetOnEvent($combofix, "_RunCF")
;disable combofix if Windows 8.1 - incompatible
If ($osv = "WIN_81") Then
	GUICtrlSetStyle($combofix, $ws_disabled)
	GUICtrlSetTip($combofix, "Incompatible with Windows 8.1")
EndIf
GUICtrlSetOnEvent($malwarebytes, "_RunMWB")
GUICtrlSetOnEvent($eset, "_RunESET")
GUICtrlSetOnEvent($spybot, "_RunSB")
GUICtrlSetOnEvent($sas, "_RunSAS")
GUICtrlSetOnEvent($housecall, "_RunHC")
GUICtrlSetOnEvent($ccleaner, "_RunCC")
GUICtrlSetOnEvent($programs, "_ProgFeat")
GUICtrlSetOnEvent($temp, "_tempfr")
GUICtrlSetOnEvent($mwbar, "_runmwbar")
;GUICtrlSetStyle($mwbar, $ws_disabled)
GUICtrlSetOnEvent($tdss, "_runtdss")
GUICtrlSetOnEvent($rmscan, "") ;_remscans
GUICtrlSetStyle($rmscan, $ws_disabled)
GUICtrlSetOnEvent($mse, "_GetMSE")
GUICtrlSetOnEvent($ticket, "_openticket")
GUICtrlSetOnEvent($ipconfig, "_WebReset")
GUICtrlSetOnEvent($winsock, "_Winsock")
GUICtrlSetOnEvent($hidadapt, "_HidNetRemove")
GUICtrlSetOnEvent($wifi, "_WifiProfileAdd")
GUICtrlSetOnEvent($speed, "_SpeedTest")
GUICtrlSetOnEvent($ncprop, "_NetAdapterProperties")
GUICtrlSetOnEvent($sfc, "_SFC")
GUICtrlSetOnEvent($dism, "_DISM")
If Not(($OSV = "WIN_8_1") Or ($OSV = "WIN_8")) Then
	GuiCtrlSetStyle($dism, $ws_disabled)
	GUICtrlSetTip($dism, "Incompatible with Windows 7 or earlier")
EndIf
GUICtrlSetOnEvent($chkdsk, "_Chkdsk")
GUICtrlSetOnEvent($aio, "_runaio")
;GUICtrlSetStyle($aio, $ws_disabled)
GUICtrlSetOnEvent($defrag, "_Defraggle")
GUICtrlSetOnEvent($devmgr, "_Devmgmt")
GUICtrlSetOnEvent($rmnac, "_rmnac") ;_rmnac
;GUICtrlSetStyle($rmnac, $ws_disabled)
GUICtrlSetOnEvent($rmmse, "") ;_rmmse
GUICtrlSetStyle($rmmse, $ws_disabled)
GUICtrlSetOnEvent($cpl, "_OpenControlPanel")
GUICtrlSetOnEvent($regedit, "_regedit")
GUICtrlSetOnEvent($msconfig, "_togglemsconfig")
GUICtrlSetOnEvent($print, "_anyprint")
GUICtrlSetOnEvent($auto, "_EasterEgg")
GUICtrlSetData($proglabel, "ResTool Ready...")
GUICtrlSetStyle($progbar, 1)
GUICtrlSetData($progbar, 0)
;main program loop.
While 1
	Sleep(100)
WEnd
#EndRegion ###Program Init and Loop
#Region ###AutoHelpers
#cs ---------------------------------------------------------------------------
FUNCTION: _automode()

PURPOSE: Controls the operation of AutoMode. Currently not invoked, waiting
		 for Auto Mode functions to be ready. Will be bound to fix everything
		 button on Auto tab of GUI.

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _automode()
	_appendLog(1, "Automated Removal Mode")
	While (Not _autoallclear)
		$next = _getnextscan()
		If (_readregstats($next) = 0) Then
			_writelastscan($next)
		Else
			_autorun($next)
		EndIf

	_appendLog(4, "Automated Removal Mode")
	_autoCC()
	_autoSFC()
	WEnd
EndFunc
#cs ---------------------------------------------------------------------------
FUNCTION: _autoappendlog()

PURPOSE: Appends to a log file stored on disk with title, called by _auto(scan)
		 methods

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _autoappendlog($shortcode, $results = -1, $message = "")
	Local $log = StringLeft(@ScriptDir, 3) & "Logs\Auto" & $ticketno & ".txt"
	If (Not FileExists($log)) Then
		_FileCreate($log)
	EndIf
	$logfile = FileOpen($log, 1)
	If $results = -1 Then
		FileWriteLine($logfile, $shortcode & " - " & _readregstats($shortcode))
	ElseIf $message = "" Then
		FileWriteLine($logfile, $shortcode & " - " & $results)
	Else
		FileWriteLine($logfile, $shortcode & " - " & $message)
	EndIf
EndFunc
#cs ---------------------------------------------------------------------------
FUNCTION: _autoallclear()

PURPOSE: Returns true when all scans have reached 0 0

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _autoallclear()
	$cfstat = Not(_readregstats("CF"))
	$mwbstat = Not(_readregstats("MWB"))
	$esetstat = Not(_readregstats("ESET"))
	$sbstat = Not(_readregstats("SB"))
	$sasstat = Not(_readregstats("SAS"))
	$hcstat = Not(_readregstats("HC"))
	Return $cfstat & $mwbstat & $esetstat & $sbstat & $sasstat & $hcstat
EndFunc
#cs ---------------------------------------------------------------------------
FUNCTION: _autorun()

PURPOSE: Converts shortcode notation to function calls

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _autorun($shortcode)
	If $shortcode = "CF" Then
		_autocf()
	ElseIf $shortcode = "MWB" Then
		_automwb()
	ElseIf $shortcode = "ESET" Then
		_autoeset()
	ElseIf $shortcode = "SB" Then
		_autosb()
	ElseIf $shortcode = "SAS" Then
		_autosas()
	ElseIf $shortcode = "HC" Then
		_autohc()
	Else
		_autocf()
		$shortcode = "CF"
	EndIf
	_writelastscan($shortcode)
EndFunc
#cs ---------------------------------------------------------------------------
FUNCTION: _getnextscan()

PURPOSE: Determines the next scan to run based on the last scan ran/skipped

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _getnextscan()
	SetError(0)
	$shortcode = RegRead("HKLM\SOFTWARE\ResTech\AutoStats" , "LastScan")
	If (@error = 0) Then
		If $shortcode = "CF" Then
			Return "MWB"
		ElseIf $shortcode = "MWB" Then
			Return "ESET"
		ElseIf $shortcode = "ESET" Then
			Return "SB"
		ElseIf $shortcode = "SB" Then
			Return "SAS"
		ElseIf $shortcode = "SAS" Then
			Return "HC"
		ElseIf $shortcode = "HC" Then
			Return "CF"
		Else
			Return "CF"
		EndIf
	Else
		Return "CF"
	EndIf
EndFunc
#cs ---------------------------------------------------------------------------
FUNCTION: _writelastscan()

PURPOSE: Stores information indicating the last scan ran using the registry

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _writelastscan($shortcode)
	RegWrite("HKLM\SOFTWARE\ResTech\AutoStats", "LastScan", "REG_SZ", $shortcode)
EndFunc
#cs -----------------------------------------------------------------------------
FUNCTION: _writeregstats()

PURPOSE: Creates and or modifies a registry entry of the given $shortcode

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Not used. Implemented for future use in Auto mode. Log scans by using
	   shortcodes like MWB, ESET, etc. Value should be numerical
#ce -----------------------------------------------------------------------------
Func _writeregstats($shortcode, $value)
	RegWrite("HKLM\SOFTWARE\ResTech\AutoStats", $shortcode, "REG_SZ", $value)
EndFunc   ;==>_writeregstats
#cs -----------------------------------------------------------------------------
FUNCTION: _readregstats()

PURPOSE: Gets the data from a registry entry of the given shortcode

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Not used. Implemented for future use in Auto mode. Read scan data by
	   using shortcodes like MWB, ESET, etc. Will fail if no key or data bad.
#ce -----------------------------------------------------------------------------
Func _readregstats($shortcode)
	SetError(0)
	$value = RegRead("HKLM\SOFTWARE\ResTech\AutoStats", $shortcode)
	If @error = 0 Then
		Return Number($value)
	Else
		Return -1
	EndIf
EndFunc   ;==>_readregstats
#EndRegion ###AutoHelpers
#Region ###Helpers
#cs ---------------------------------------------------------------------------
FUNCTION: _start()

PURPOSE: Handle GUI updates and log entries to start a function

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _start($program)
	_appendlog(1, $program)
	GUICtrlSetData($proglabel, $program & " Running")
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
EndFunc
#cs ---------------------------------------------------------------------------
FUNCTION: _end()

PURPOSE: Handle GUI updates and log entries to finish a function

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _end($program)
	_appendlog(4, $program)
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc
#cs ---------------------------------------------------------------------------
FUNCTION: _close()

PURPOSE: Exits the program. Invoked by $GUI_EVENT_CLOSE

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES:
#ce -----------------------------------------------------------------------------
Func _close()
	Exit
EndFunc   ;==>_close

#cs -----------------------------------------------------------------------------
FUNCTION: _readableosv()

PURPOSE: Generates a user-readable string containing OS Version and Architecture

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Will point out if you are using anything XP or older or we've never seen
	   Gives UNSUPPORTED OS [arch] as string.
#ce -----------------------------------------------------------------------------
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

#cs -----------------------------------------------------------------------------
FUNCTION: _runau3()

PURPOSE: Runs an external uncompiled AutoIt Script, uses Run() syntax

AUTHOR: Kevin Morgan via some code from the Intertubes

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Not used, but not deprecated due to usefulness for testing.
#ce -----------------------------------------------------------------------------
Func _runau3($sfilepath, $sworkingdir = "", $ishowflag = @SW_SHOW, $ioptflag = 0)
	Return Run('"' & @AutoItExe & '" /AutoIt3ExecuteScript "' & $sfilepath & '"', $sworkingdir, $ishowflag, $ioptflag)
EndFunc   ;==>_runau3

#cs -----------------------------------------------------------------------------
FUNCTION: _setticket()

PURPOSE: Logs the computer's ticket number in the registry

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Will keep prompting you until you enter something, no more error checking
#ce -----------------------------------------------------------------------------
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

#cs -----------------------------------------------------------------------------
FUNCTION: _updatelastopen()

PURPOSE: Updates the timestamp on ResTool's last run

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Relies on System RTC, if that is incorrect, this will be too.
#ce -----------------------------------------------------------------------------
Func _updatelastopen()
	RegWrite("HKLM\SOFTWARE\ResTech", "LastOpen", "REG_SZ", _NowCalc())
EndFunc   ;==>_updatelastopen

#cs -----------------------------------------------------------------------------
FUNCTION: _appendlog()

PURPOSE: Appends the logfile stored on ResFlash

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Takes 3 arguments: update code, program name, [message]
	   Codes:	1: Program/Process start
				2: Error running program/process
				3: Results of program/process (message required)
				4: Program/Process completed
				5: Miscellaneous message (message required)
		No entry will be created for malformed data, specifically invalid code
#ce -----------------------------------------------------------------------------
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
#cs -----------------------------------------------------------------------------
FUNCTION: _winwaitnotify()

PURPOSE: Wrapper for WinWait that pops up a messagebox if the timeout is reached

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Due to incompatibilities with some processes, default timeout is set to
	   0/unlimited, which renders this function pretty much useless.
	   a better implementation is in the works with _waitclick
#ce -----------------------------------------------------------------------------
Func _winwaitnotify($title, $text, $controltitle, $timeout = 0)
	If (WinWait($title, $text, $timeout) = 0) Then
		MsgBox($mb_yesno, "Couldn't find " & $title, "ResTool could not find the window titled " & $title & "." & @CRLF & "Please continue to this window and click " & $controltitle & ".")
	EndIf
EndFunc   ;==>_winwaitnotify

#cs -----------------------------------------------------------------------------
FUNCTION: _queuestartup()

PURPOSE: Creates a shortcut to run ResTool at startup

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Really need an _unqueuestartup. Not Used. Planned for forced-restarts
	   like in MWB and chkdsk /f
#ce -----------------------------------------------------------------------------
Func _queuestartup()
	FileCreateShortcut(@ScriptFullPath, @StartupCommonDir & "\RT.lnk")
EndFunc   ;==>_queuestartup

#cs -----------------------------------------------------------------------------
FUNCTION: _waitclick()

PURPOSE: Handles a WinWait/ControlClick combo where there is a possibility that
		 AutoIt may miss the window

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: Somewhat Recently

NOTES: Not yet used. Default timeout is useless. Gives user option to terminate
	   script. Needs timings or a better way to detect control clickability.
#ce -----------------------------------------------------------------------------
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
#Region ###AutoScanners
#cs -----------------------------------------------------------------------------
FUNCTION: _autocf()

PURPOSE: Runs ComboFix for Auto Mode

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _autocf()
EndFunc
#cs -----------------------------------------------------------------------------
FUNCTION: _automwb()

PURPOSE: Runs Malwarebytes for Auto Mode

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _automwb()
EndFunc
#cs -----------------------------------------------------------------------------
FUNCTION: _autoeset()

PURPOSE: Runs Eset for Auto Mode

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _autoeset()
EndFunc
#cs -----------------------------------------------------------------------------
FUNCTION: _autosb()

PURPOSE: Runs SpyBot for Auto Mode

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _autosb()
EndFunc
#cs -----------------------------------------------------------------------------
FUNCTION: _autosas()

PURPOSE: Runs SuperAntiSpyware for Auto Mode

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _autosas()
EndFunc
#cs -----------------------------------------------------------------------------
FUNCTION: _autohc()

PURPOSE: Runs HouseCall for Auto Mode

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _autohc()
EndFunc
#cs -----------------------------------------------------------------------------
FUNCTION: _autocc()

PURPOSE: Runs CCleaner for Auto Mode

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _autocc()
EndFunc
#cs -----------------------------------------------------------------------------
FUNCTION: _autosfc()

PURPOSE: Runs System File Check for Auto Mode

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _autosfc()
EndFunc
#EndRegion ###AutoScanners
#Region ###Scanners
#cs -----------------------------------------------------------------------------
FUNCTION: _runcf()

PURPOSE: Runs ComboFix like a boss

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/11/14

NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _runcf()
	_appendlog(1, "ComboFix")
	GUICtrlSetData($proglabel, "ComboFix Running")
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	;run CF
	Run(@ScriptDir & "/Script/Scanners/CF/CF.exe")
	;Handle startup disclaimers
	While (Not WinExists("ComboFix"))
		If (WinExists("ComboFix: Disclaimer")) Then
			WinActivate("ComboFix: Disclaimer")
			Send("{ENTER}")
		EndIf
		Sleep(1000)
	WEnd
	;Handle update Dialogs
	While (WinExists("ComboFix"))
		If (WinExists("ComboFix", "Yes")) Then
			Sleep(100)
			Send("{ENTER}")
		EndIf
		Sleep(100)
	WEnd
	;there may be an agree somewhere in here that is unhandled?
	;Handle AV Warnings;
	While (Not WinExists("Administrator:  ."))
		Sleep(100)
		If (WinExists("Warning !!")) Then
			WinActivate("Warning !!")
			Send("{ENTER}")
		EndIf
	WEnd
	;wait for autoit to close, handle repeated open/close of conhost
	While(True)
		If(ProcessExists("conhost.exe")) Then
			ProcessWaitClose("conhost.exe")
		Else
			Sleep(1000)
			If Not(ProcessExists("conhost.exe")) Then
				ExitLoop
			EndIf
		EndIf
	WEnd
	;parse logfile

	;open file
	Dim $cflog = FileOpen("C:\COMBOFIX.TXT")
	;advance to the Other Deletions portion of the scan
	If(Not ($cflog = -1)) Then
		Dim $str = FileReadLine($cflog)
		While (StringInStr($str, "Other Deletions") = 0 And Not (@error = -1))
			$str = FileReadLine($cflog)
		WEnd
		;count entries between Other Deletions and files created
		Dim $count = 0
		If Not(@error = -1) Then
			$str = FileReadLine($cflog)
			While (StringInStr($str, "(((((((((((") = 0 And Not (@error = -1))
				$str = FileReadLine($cflog)
				If (Not(StringLeft($str, 1) = ".") And Not(StringLeft($str, 1) = "(")) Then
					$count += 1
				EndIf
			WEnd
		EndIf
		MsgBox(0, "ComboFix Log Results", "ComboFix made " & $count & " deletion(s).")
		_appendLog(3, "ComboFix", $count & " deletion(s).")
	Else
		MsgBox(0, "ComboFix Log Error", "Unable to open ComboFix log file.")
		_appendLog(2, "ComboFix", "Log file not found")
	EndIf
	If ProcessExists("notepad.exe") Then
		ProcessClose("notepad.exe")
	EndIf
	_appendlog(4, "ComboFix")
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc   ;==>_runcf

#cs -----------------------------------------------------------------------------
FUNCTION: _runmwb()

PURPOSE: Runs Malwarebytes

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Doesn't handle the log, only completes after stuff is removed manually
#ce -----------------------------------------------------------------------------
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

#cs -----------------------------------------------------------------------------
FUNCTION: _installmwb()

PURPOSE: Handles the malwarebytes installation

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Might miss a dialog once in a blue moon, but is otherwise super reliable
#ce -----------------------------------------------------------------------------
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

#cs -----------------------------------------------------------------------------
FUNCTION: _runeset()

PURPOSE: Runs ESET Online Scanner. Logs Results as well

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Currently misses the Enable Detection of PUPs screen on most machines
#ce -----------------------------------------------------------------------------
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
	Sleep(1000)
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

#cs -----------------------------------------------------------------------------
FUNCTION: _runsb()

PURPOSE: Runs SpyBot Search and Destroy

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Hangs in a lot of places, really need to optimize.
#ce -----------------------------------------------------------------------------
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

#cs -----------------------------------------------------------------------------
FUNCTION: _runsas()

PURPOSE: Runs Super Anti Spyware

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Really struggles with the update dialog.
#ce -----------------------------------------------------------------------------
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
#cs -----------------------------------------------------------------------------
FUNCTION: _runhc()

PURPOSE: Runs Housecall

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 11/24/14

NOTES: Now works perfectly fine, but requires a slight bit of cheating
#ce -----------------------------------------------------------------------------
Func _runhc()
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "HouseCall Running")
	If ($osv = "X86") Then
		$ver = "32"
	Else
		$ver = "64"
	EndIf
	_appendlog(1, "Trend Micro HouseCall")
	Run(@ScriptDir & "\Script\Scanners\HC\HC" & $ver & ".exe")
	_winwaitnotify("HouseCall Download", "", "OK")
	Sleep(1000)
	If (WinActive("Trend Micro HouseCall", "Unable to complete the download.")) Then
		MsgBox(48, "Network Connection", "Plug in the proxy and press OK.")
		_runhc()
	Else
		WinWaitClose("HouseCall Download")
		Sleep(1000)
		$handle = _IEAttach(WinWaitActive("[CLASS:#32770]"), "embedded")
		$accept = _IEGetObjById($handle, "r_eula_0")
		$next = _IEGetObjById($handle, "btn_next")
		;handle EULA already accepted (above calls set @error to 7 if so)
		If Not(@error = 7) Then
			_IEAction($accept, "click")
			_IEAction($next, "click")
		Else
			SetError(0)
		EndIf
		Sleep(1000)
		$scan = _IEGetObjById($handle, "ScanNow")
		;must manually click this since it causes focus loss on the main window
		ControlClick("[CLASS:#32770]", "", "Internet Explorer_Server1", "left", 2, 463, 188)
		Sleep(1000)
		$settingshandle = _IEAttach(WinGetHandle("Settings"), "embedded")
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

#cs -----------------------------------------------------------------------------
FUNCTION: _runcc()

PURPOSE: Runs CCleaner

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 11/18/14

NOTES:
#ce -----------------------------------------------------------------------------
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
			$rresults = 0
			While (ControlCommand("", $text, "Button3", "IsEnabled"))
				ControlClick("", $text, "Button3")
				$rresults +=1
				Sleep(10)
			WEnd
			ControlClick("", $text, "Button5")
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
#cs -----------------------------------------------------------------------------
FUNCTION: _runmwbar()

PURPOSE: Runs Malwarebytes' Anti Malware

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/4/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _runmwbar()
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "MWBAR Running")
	_appendlog(1, "MWBAR")
	Run(@ScriptDir & "\Script\Scanners\MWBAR.exe")
	WinWait("Malwarebytes Anti-Rootkit", "")
	ControlClick("Malwarebytes Anti-Rootkit", "", "Button2")
	WinWait("Malwarebytes Anti-Rootkit BETA v1.08.2.1001", "")
	Send("!n")
	Sleep(1000)
	Send("{TAB}{TAB}{ENTER}")
	Sleep(60000);wait for update for an entire minute
	Send("{TAB}{TAB}{TAB}{ENTER}")
	MsgBox($MB_TOPMOST, "Wait for scan to complete", "Close this message box when MWBAR is finished scanning completely.")
	Send("!n")
	MsgBox($MB_TOPMOST, "View Results", "Results should be showing. Record them and then close this dialog")
	MsgBox($MB_TOPMOST, "Confirm Close", "Don't close this until you've actually recorded the results!")
	Send("!e")
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
	_appendlog(4, "MWBAR")
EndFunc
#cs -----------------------------------------------------------------------------
FUNCTION: _runtdss()

PURPOSE: Runs TDSSKiller

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 11/20/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _runtdss()
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "TDSSKiller Running")
	_appendlog(1, "TDSSKiller")
	Local $aTDSSLog
	Dim $suscount = 0, $infcount = 0
	ShellExecute(@ScriptDir & "\Script\Scanners\TDSS.exe", "-silent -l C:\TDSSAI.txt -accepteula -accepteulaksn")
	If FileExists("C:\TDSSAI.txt") Then
            If Not _FileReadToArray("C:\TDSSAI.txt", $aTDSSLog) Then
                MsgBox(16, "Error", "Error parsing results")
                _appendlog(2, "TDSSKiller", "Error parsing results")
            Else
                For $i = 1 To UBound($aTDSSLog) - 1
                    If StringInStr($aTDSSLog[$i], "Suspicious") Then
                        $suscount += 1
                    ElseIf StringInStr($aTDSSLog[$i], "infected") Then
                        $infcount += 1
                    EndIf
                Next
				MsgBox(0, "Results", $suscount & " suspicious files, and " & $infcount & " infected files.")
				_appendlog(3, "TDSSKiller", $suscount & " suspicious / " & $infcount & " infected")
            EndIf
        EndIf
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
	_appendlog(4, "TDSSKiller")
EndFunc
#EndRegion ###Scanners
#Region ###OS
#cs -----------------------------------------------------------------------------
FUNCTION: _rmnac

PURPOSE: Uninstalls Cisco NAC Agent

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 12/2/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _rmnac()
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "Removing Cisco NAC Agent")
	_appendlog(1, "NAC Uninstall")
	ShellExecute("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Cisco\Cisco NAC Agent\Uninstall Cisco NAC Agent.lnk")
	WinWait("Windows Installer", "")
	ControlClick("Windows Installer", "", "Button1")
	Sleep(1000)
	WinWait("Windows Installer", "")
	WinWait("Cisco NAC Agent ", "")
	WinWaitClose("Cisco NAC Agent ", "")
	_appendlog(4, "NAC Uninstall")
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc
#cs -----------------------------------------------------------------------------
FUNCTION: _anyprint()

PURPOSE: Installs Anywhere Printer support

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 11/20/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _anyprint()
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "Installing Anywhere Printer")
	_appendlog(1, "Pharos Install")
	Run(@ScriptDir & "\Script\Installers\Pharos.exe")
	WinWait("Package ""Windows AnywherePrint Client"" installer.", "")
	WinActivate("[LAST]")
	ControlClick("Package ""Windows AnywherePrint Client"" installer.", "", "Button1")
	WinWait("Package ""Windows AnywherePrint Client"" installer.", "Install")
	ControlClick("Package ""Windows AnywherePrint Client"" installer.", "", "Button1")
	WinWait("Package ""Windows AnywherePrint Client"" installer.", "Finish")
	ControlClick("Package ""Windows AnywherePrint Client"" installer.", "", "Button2")
	_appendlog(4, "Pharos Install")
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc
#cs -----------------------------------------------------------------------------
FUNCTION: _tempfr()

PURPOSE: Uses the Disk Cleanup tool to remove temp files

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 11/20/14

NOTES: Removes default selected files, since in most cases unchecked entries
	   only add up to a couple mb of space or are handled by CCleaner
#ce -----------------------------------------------------------------------------
Func _tempfr()
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "Running Disk Cleanup")
	_appendlog(1, "Disk Cleanup")
	Run(@SystemDir & "\cleanmgr.exe")
	WinWait("Disk Cleanup", "")
	WinWait("Disk Cleanup for  (C:)", "")
	WinActivate("Disk Cleanup for  (C:)", "")
	ControlClick("Disk Cleanup for  (C:)", "", "Button4")
	WinWait("Disk Cleanup", "Delete Files")
	ControlClick("Disk Cleanup", "Delete Files", "Button1")
	ProcessWaitClose("cleanmgr.exe")
	_appendlog(4, "Disk Cleanup")
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc

#cs -----------------------------------------------------------------------------
FUNCTION: _openticket()

PURPOSE: Opens the ticket

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 11/20/14

NOTES: Opens in default browser
#ce -----------------------------------------------------------------------------
Func _openticket()
	ShellExecute("http://intranet.restech.niu.edu/tickets/" & $ticketno)
EndFunc
#cs -----------------------------------------------------------------------------
FUNCTION: _regedit()

PURPOSE: Opens the Registry Editor

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 11/20/14

NOTES: Opens the registry associated with program arch. 64 gets 64 reg, 32 to 32
#ce -----------------------------------------------------------------------------
Func _regedit()
	ShellExecute("regedit")
EndFunc
#cs -----------------------------------------------------------------------------
FUNCTION: _togglemsconfig()

PURPOSE: Toggles safe boot

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 11/18/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _togglemsconfig()
	Run("msconfig")
	WinWait("System Configuration", "" )
	WinActivate("System Configuration", "")
	Sleep(100)
	Send("^{TAB}")
	IF(Not(_readregstats("MSCFG") == 1)) Then
		ControlClick("System Configuration", "", "Button5")
		Sleep(100)
		ControlClick("System Configuration", "", "Button9")
		Sleep(100)
		ControlClick("System Configuration", "", "Button30")
		WinWait("System Configuration", "You may need to restart your computer")
		ControlClick("System Configuration", "", "Button3")
		_writeregstats("MSCFG", 1)
		GUICtrlSetColor($msconfig, $COLOR_RED)
	Else
		ControlClick("System Configuration", "", "Button5")
		_writeregstats("MSCFG", 0)
		Sleep(100)
		ControlClick("System Configuration", "", "Button30")
		WinWait("System Configuration", "You may need to restart your computer")
		ControlClick("System Configuration", "", "Button3")
		GUICtrlSetColor($msconfig, $COLOR_BLACK)
	EndIf
EndFunc
#cs -----------------------------------------------------------------------------
FUNCTION: _speedtest

PURPOSE: Opens Speedtest in Internet Explorer

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Opens in default browser
#ce -----------------------------------------------------------------------------
Func _speedtest()
	ShellExecute("http://speedtest.niu.edu")
EndFunc   ;==>_speedtest

#cs ---------------------------------s--------------------------------------------
FUNCTION: _netadapterproperties()

PURPOSE: Opens Network Adapter Settings

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES:
#ce -----------------------------------------------------------------------------
Func _netadapterproperties()
	ShellExecute("Ncpa.cpl")
EndFunc   ;==>_netadapterproperties

#cs -----------------------------------------------------------------------------
FUNCTION: _progfeat()

PURPOSE: Opens Programs and Features

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES:
#ce -----------------------------------------------------------------------------
Func _progfeat()
	ShellExecute("appwiz.cpl")
EndFunc   ;==>_progfeat

#cs -----------------------------------------------------------------------------
FUNCTION: _devmgmt()

PURPOSE: Opens Device Management in MMC

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES:
#ce -----------------------------------------------------------------------------
Func _devmgmt()
	ShellExecute("devmgmt.msc")
EndFunc   ;==>_devmgmt

#cs -----------------------------------------------------------------------------
FUNCTION: _defraggle

PURPOSE: Defragments the system drive

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES:
#ce -----------------------------------------------------------------------------
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

#cs -----------------------------------------------------------------------------
FUNCTION: _webreset

PURPOSE: Performs an IPConfig reset

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES:
#ce -----------------------------------------------------------------------------
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

#cs -----------------------------------------------------------------------------
FUNCTION: _winsock()

PURPOSE: Runs a Winsock Reset

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES:
#ce -----------------------------------------------------------------------------
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

#cs -----------------------------------------------------------------------------
FUNCTION: _hidnetremove

PURPOSE: Removes hidden network adapters using the Device Console

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES:
#ce -----------------------------------------------------------------------------
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

#cs -----------------------------------------------------------------------------
FUNCTION: _chkdsk()

PURPOSE: Catalogs a disk check for next startup

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 11/25/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _chkdsk()
	_appendlog(1, "Disk Check")
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "Running System File Check")
	$var = Run(@ComSpec & " /c" & '%systemroot%\system32\chkdsk.exe C: > ' & @ScriptDir & "\Logs\chkdsk.txt","", @SW_HIDE)
	ProcessWaitClose($var)
	$fd = FileOpen(@ScriptDir & "\Logs\chkdsk.txt")
	$text = FileRead($fd)
	If (StringInStr($text, "problems") = 0) Then
		MsgBox(0, "No Errors", "Chkdsk did not find any errors")
		_appendlog(4, "Disk Check")
	Else
		If (MsgBox($mb_yesno, "Errors Found", "Should I schedule a repair at next boot?") == $idyes) Then
			ShellExecuteWait("fsutil", "dirty set c:")
			If (MsgBox($mb_yesno, "Chkdsk queued", "Chkdsk is scheduled to repair the hard drive at next boot. Would you like to restart now?") == $idyes) Then
				_queuestartup()
				_appendlog(3, "Disk Check", "Errors Found, Restarting to fix.")
				ShellExecute("shutdown", "/r /t 0", "", "", @SW_HIDE)
			Else
				_appendLog(3, "Disk Check", "Errors found, Will fix at next restart.")
				MsgBox(0, "Chkdsk queued", "Chkdsk will run the next time the computer restarts.")
			EndIf
		Else
		_appendlog(4, "Disk Check", "Errors Found but no repair queued.")
		MsgBox(0, "Chkdsk Not Queued", "Chkdsk will not fix disk errors.")
		EndIf
	EndIf
	FileClose($fd)
	FileDelete(@ScriptDir & "\Logs\chkdsk.txt")
EndFunc   ;==>_chkdsk

#cs -----------------------------------------------------------------------------
FUNCTION: _sfc()

PURPOSE: Runs SFC

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Currently doesn't check against date stamps, which it should to verify
	   results are from current scan. Additionally, I could diff before/after
#ce -----------------------------------------------------------------------------
Func _sfc()
	_appendlog(1, "System File Checker")
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "Running System File Check")
	$var = Run(@ComSpec & " /C sfc /scannow", "", @SW_HIDE, $STDOUT_CHILD)
	ProcessWaitClose($var)
	$text = StdoutRead($var)
	If Not(StringInStr($text, "pending") = 0) Then
		MsgBox(0, "SFC Failed", "SFC repairs or Windows Updates required to continue.")
		_appendlog(2, "System File Checker", "Failed due to pending updates")
	ElseIf Not(StringInStr($text, "not find any integrity violations") = 0) Then
		MsgBox(0, "SFC Finished", "SFC did not find any corrupt files")
		_appendlog(4, "System File Checker", "Did not find any corrupt files")
	ElseIf Not(StringInStr($text, "unable to fix") = 0) Then
		MsgBox(0, "SFC Failed", "SFC found corrupt files but could not repair them.")
		_appendlog(2, "System File Checker", "Found corrupt files but could not repair")
	ElseIf Not(StringInStr($text, "successfully repaired") = 0) Then
		MsgBox(0, "SFC Repairs Finished", "SFC found and repaired corrupt files.")
		_appendlog(3, "System File Checker", "Repaired corrupted files")
	Else
		MsgBox(0, "SFC Status Unknown", $text)
		_appendlog(2, "System File Checker", "Failed with unknown status")
		_appendlog(5, "DEBUG INFO -SFC-", $text)
	EndIf
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc   ;==>_sfc

#cs -----------------------------------------------------------------------------
FUNCTION: _wifiprofileadd

PURPOSE: Adds the WiFi

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Not sure if I actually have the xml bundled with ResTool
#ce -----------------------------------------------------------------------------
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

#cs -----------------------------------------------------------------------------
FUNCTION: _opencontrolpanel

PURPOSE: Opens Control Panel

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: A Long Time Ago

NOTES:
#ce -----------------------------------------------------------------------------
Func _opencontrolpanel()
	GUICtrlSetData($proglabel, "Opening Control Panel")
	ShellExecuteWait("control", "", "")
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc   ;==>_opencontrolpanel
#cs -----------------------------------------------------------------------------
FUNCTION: _dism

PURPOSE: runs a DISM scan

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 11/25/14

NOTES:
#ce -----------------------------------------------------------------------------
Func _dism()
	_appendlog(1, "Deployment Image Servicing and Management Tool")
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "Running DISM")
	ShellExecuteWait("dism", "/Online /Cleanup-Image /RestoreHealth", "", "", @SW_HIDE)
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	If (StringInStr(FileRead(@WindowsDir & "\Logs\dism\dism.log"), "Failed to restore the image health", 1)) Then
		_appendlog(2, "Deployment Image Servicing and Management Tool", "DISM failed to reapair system file corruption")
		MsgBox($mb_iconwarning, "DISM", "DISM failed to restore the image health")
	Else
		_appendlog(4, "Deployment Image Servicing and Management Tool")
		MsgBox(0, "DISM", "DISM Completed without error, but whether it made repairs or not is unknown")
	EndIf
	GUICtrlSetData($proglabel, "ResTool Ready.")
EndFunc   ;==>_dism
#cs -----------------------------------------------------------------------------
FUNCTION: _runaio

PURPOSE: Runs Tweaking.com - Windows Repair All in One

AUTHOR: Nobody

DATE OF LAST UPDATE: 11/25/14

NOTES: Uses silent mode. Not logging compliant.
#ce -----------------------------------------------------------------------------
Func _runaio()
	_appendlog(1, "All-in-One")
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "Running AIO")
	;if not installed
	If Not(FileExists($32progfiledir & "\Tweaking.com\Windows Repair (All in One)\Repair_Windows.exe")) Then
		Run(@ScriptDir & "\Script\AIO.exe")
		WinWait("Tweaking.com - Windows Repair (All in One) Setup", "")
		ControlClick("Tweaking.com - Windows Repair (All in One) Setup", "", "Button3")
		WinWait("Tweaking.com - Windows Repair (All in One) Setup", "C&hange...")
		ControlClick("Tweaking.com - Windows Repair (All in One) Setup", "", "Button1")
		WinWait("Tweaking.com - Windows Repair (All in One) Setup", "Install shortcuts")
		ControlClick("Tweaking.com - Windows Repair (All in One) Setup", "", "Button1")
		Sleep(500)
		ControlClick("Tweaking.com - Windows Repair (All in One) Setup", "", "Button1")
		WinWait("Tweaking.com - Windows Repair (All in One) Setup", "Create Desktop")
		;need to actually do these correctly
		Send("!n")
		Sleep(100)
		Send("!f")
	EndIf
	ShellExecuteWait($32progfiledir & "\Tweaking.com\Windows Repair (All in One)\Repair_Windows.exe", "/silent")
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	_appendlog(4, "MSE Install")
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc   ;==>_runaio
#cs -----------------------------------------------------------------------------
FUNCTION: _getmse()

PURPOSE: Installs MSE

AUTHOR: Kevin Morgan

DATE OF LAST UPDATE: 11/20/14

NOTES: Does not activate yet on Win 8.1
#ce -----------------------------------------------------------------------------
Func _getmse()
	_appendlog(1, "MSE Install")
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	GUICtrlSetData($proglabel, "Installing MSE")
	If (Not ($osv = "WIN_81") And Not ($osv = "WIN_8")) Then
		If ($osa == "X86") Then
			Run(@ScriptDir & "\Script\Installers\MSE32.exe")
		Else
			Run(@ScriptDir & "\Script\Installers\MSE64.exe")
		EndIf
			WinWait("Microsoft Security Essentials", "")
			ControlClick("Microsoft Security Essentials", "", "Button1")
			WinWait("Microsoft Security Essentials", "accept")
			ControlClick("Microsoft Security Essentials", "", "Button1")
			WinWait("Microsoft Security Essentials", "join the program")
			ControlClick("Microsoft Security Essentials", "", "Button2")
			Sleep(100)
			ControlClick("Microsoft Security Essentials", "", "Button4")
			WinWait("Microsoft Security Essentials", "optimize")
			ControlClick("Microsoft Security Essentials", "", "Button2")
			WinWait("Microsoft Security Essentials", "may conflict with")
			ControlClick("Microsoft Security Essentials", "", "Button1")
			WinWait("Microsoft Security Essentials", "Finish")
			ControlClick("Microsoft Security Essentials", "", "Button1")
	Else
		;Need to enable because Win 8.
	EndIf
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	_appendlog(4, "MSE Install")
	GUICtrlSetData($proglabel, "ResTool Ready...")
EndFunc   ;==>_getmse
#cs -----------------------------------------------------------------------------
FUNCTION: _getprint

PURPOSE: Currently does nothing

AUTHOR: Nobody

DATE OF LAST UPDATE: Never

NOTES: Will install AnywherePrint client
#ce -----------------------------------------------------------------------------

Func _getprint()
EndFunc   ;==>_getprint

#EndRegion ###OS
#Region ###Removal
#cs -----------------------------------------------------------------------------
FUNCTION: _remnor()

PURPOSE: Nothing

AUTHOR: Nobody

DATE OF LAST UPDATE: Never

NOTES: Should run Norton Removal Tool
#ce -----------------------------------------------------------------------------
Func _remnor()

EndFunc   ;==>_remnor
#cs -----------------------------------------------------------------------------
FUNCTION: _remavg()

PURPOSE: Nothing

AUTHOR: Nobody

DATE OF LAST UPDATE: Never

NOTES: Should run AVG Removal Tool
#ce -----------------------------------------------------------------------------
Func _remavg()

EndFunc   ;==>_remavg
#cs -----------------------------------------------------------------------------
FUNCTION: _remavt()

PURPOSE: Nothing

AUTHOR: Nobody

DATE OF LAST UPDATE: Never

NOTES: Should remove Avista, but I don't think anyone actually uses that ever
#ce -----------------------------------------------------------------------------
Func _remavt()

EndFunc   ;==>_remavt
#cs -----------------------------------------------------------------------------
FUNCTION: _remkas()

PURPOSE: Nothing

AUTHOR: Nobody

DATE OF LAST UPDATE: Never

NOTES: Should run Kaspersky Removal Tool
#ce -----------------------------------------------------------------------------
Func _remkas()

EndFunc   ;==>_remkas
#cs -----------------------------------------------------------------------------
FUNCTION: _remmca()

PURPOSE: Nothing

AUTHOR: Nobody

DATE OF LAST UPDATE: Never

NOTES: Should run McAfee Removal Tools
#ce -----------------------------------------------------------------------------
Func _remmca()

EndFunc   ;==>_remmca
#EndRegion ###Removal
#Region ###Other
#cs -----------------------------------------------------------------------------
FUNCTION: _easteregg()

PURPOSE: Totally does not facilitate an easter egg in the program. Not at all

AUTHOR: It's a Secret

DATE OF LAST UPDATE: A Long Time Ago

NOTES: Okay, it might actually be an easter egg. Blame Dan.
#ce -----------------------------------------------------------------------------

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
