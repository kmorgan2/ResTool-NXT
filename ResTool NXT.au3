#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ResTool NXT.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_Res_Description=ResTool NXT
#AutoIt3Wrapper_Res_Fileversion=0.1.150206.0
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region ###Includes
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
#include <ITaskBarList.au3>
#include <String.au3>
#include <Array.au3>
#EndRegion ###Includes
#Region ###Program Init and Loop

;set the program to GUIOnEventMode which provides a more robust pre-implemented GUI event manager.
Opt("GUIOnEventMode", 1)

#Region ### START Koda GUI section ### Form=C:\Users\ResTech\Desktop\ResToolNxt_Form_New.kxf
;create basic form and form elements
$form = GUICreate("ResTool NXT: Beta!", 330, 460)
$progbar = GUICtrlCreateProgress(10, 390, 310, 20)

;set the progress bar to loading while we are waiting for program to init
GUICtrlSetStyle($progbar, 8)
GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
$proglabel = GUICtrlCreateLabel("Initializing ResTool", 10, 410, 310, 20, $ss_center)
$tab = GUICtrlCreateTab(0, 0, 330, 380)

;create first tabsheet - 'Virus Removal' and buttons
$tabsheet1 = GUICtrlCreateTabItem("Virus Removal")
GUICtrlSetState(-1, $gui_show) ;defines this as the default tab
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

;create second tabsheet - 'OS Troubleshooting' and buttons
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

;create third tabsheet - 'Removal Tools' and a placeholder
$tabsheet3 = GUICtrlCreateTabItem("Removal Tools")
$ni = GUICtrlCreateLabel("NOT IMPLEMENTED", 0, 173, 320, 28, $ss_center)

;create fourth tabsheet - 'Auto' and a button. Add some element to provide scan data
$tabsheet4 = GUICtrlCreateTabItem("Auto")
$label1 = GUICtrlCreateLabel("DON'T CLICK THIS BUTTON! SRSLY!", 4, 173, 320, 28, $ss_center)
GUICtrlSetFont(-1, 16, 400, 0, "MS Sans Serif")
GUICtrlSetColor(-1, 16711680);this color got converted to decimal somehow, but it's just red
$auto = GUICtrlCreateButton(" FIX EVERYTHING!", 52, 213, 200, 100)
GUICtrlSetTip(-1, "You've been warned!")

;create the status bar
$statusbar1 = _GUICtrlStatusBar_Create($form)
Dim $statusbar1_partswidth[3] = [150, 250, -1]
_GUICtrlStatusBar_SetParts($statusbar1, $statusbar1_partswidth)
_GUICtrlStatusBar_SetMinHeight($statusbar1, 20)

;set the state of the entire GUI to visible.
GUISetState(@SW_SHOW)

;enable thumbnail progress
$oTaskbar = _ITaskBar_CreateTaskBarObj()
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
		If ($idno = MsgBox($mb_yesno, "ResTool NXT: Ticket Number", "This computer has a ticket number recorded already. Is the current ticket for this computer still TT" & $ticketno & "?")) Then
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

;Put dynamic values on the status bar
_GUICtrlStatusBar_SetText($statusbar1, _readableosv(), 0)
_GUICtrlStatusBar_SetText($statusbar1, $ip, 1)
_GUICtrlStatusBar_SetText($statusbar1, "TT" & $ticketno, 2)

;if Auto MSConfig has been toggled to Safe Mode, make the button red to indicate as such
If (_readregstats("MSCFG") == 1) Then
	GUICtrlSetColor($msconfig, $COLOR_RED)
EndIf

;Remove Startup shortcut to ResTool if it exists
If (FileExists(@StartupCommonDir & "\RT.lnk")) Then
	FileDelete(@StartupCommonDir & "\RT.lnk")
EndIf

;Map GUI Events to Functions (and disable if necessary)
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
GUICtrlSetOnEvent($tdss, "_runtdss")

;currently scanner/cc removal not implemented
GUICtrlSetOnEvent($rmscan, "") ;Func _remscans
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

;disable DISM if we are not on Windows 8
If Not (($osv = "WIN_8_1") Or ($osv = "WIN_8")) Then
	GUICtrlSetStyle($dism, $ws_disabled)
	GUICtrlSetTip($dism, "Incompatible with Windows 7 or earlier")
EndIf
GUICtrlSetOnEvent($chkdsk, "_Chkdsk")
GUICtrlSetOnEvent($aio, "_runaio")
GUICtrlSetOnEvent($defrag, "_Defraggle")
GUICtrlSetOnEvent($devmgr, "_Devmgmt")
GUICtrlSetOnEvent($rmnac, "_rmnac")

;currently MSE removal is not implemented
GUICtrlSetOnEvent($rmmse, "") ;_rmmse
GUICtrlSetStyle($rmmse, $ws_disabled)
GUICtrlSetOnEvent($cpl, "_OpenControlPanel")
GUICtrlSetOnEvent($regedit, "_regedit")
GUICtrlSetOnEvent($msconfig, "_togglemsconfig")
GUICtrlSetOnEvent($print, "_anyprint")
GUICtrlSetOnEvent($auto, "_EasterEgg")
GUICtrlSetStyle($progbar, 1)
GUICtrlSetData($progbar, 0)
GUICtrlSetData($proglabel, "ResTool Ready...")

;Main program loop. Really complicated.
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
	_appendlog(1, "Automated Removal Mode") ;Log AutoMode start
	While (Not _autoallclear) ;While there are remaining scans necessary get and run next scan
		$next = _getnextscan()
		If (_readregstats($next) = 0) Then
			_writelastscan($next)
		Else
			_autorun($next)
		EndIf
	WEnd
	_autocc() ;run cc after scans finished
	_autosfc() ;run sfc after cc finished
	_appendlog(4, "Automated Removal Mode") ;Log AutoMode Completion
EndFunc   ;==>_automode
#cs ---------------------------------------------------------------------------
	FUNCTION: _autoappendlog()

	PURPOSE: Appends to a log file stored on disk with title, called by _auto(scan)
	methods

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _autoappendlog($shortcode, $results = -1, $message = "")

	;dynamically build the logfile name: "'X':\Logs\Auto\ticketno.txt"
	Local $log = StringLeft(@ScriptDir, 3) & "Logs\Auto" & $ticketno & ".txt"

	;If the logfile does not exist, make it
	If (Not FileExists($log)) Then
		_FileCreate($log)
	EndIf

	;open the logfile for writing, write the results, and close the log
	$logfile = FileOpen($log, 1)
	If $results = -1 Then
		FileWriteLine($logfile, $shortcode & " - " & _readregstats($shortcode))
	ElseIf $message = "" Then
		FileWriteLine($logfile, $shortcode & " - " & $results)
	Else
		FileWriteLine($logfile, $shortcode & " - " & $message)
	EndIf
	FileClose($logfile)
EndFunc   ;==>_autoappendlog
#cs ---------------------------------------------------------------------------
	FUNCTION: _autoallclear()

	PURPOSE: Returns true when all scans have reached 0 0

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _autoallclear()

	;get Scan = 0 for all scans
	$cfstat = Not (_readregstats("CF"))
	$mwbstat = Not (_readregstats("MWB"))
	$esetstat = Not (_readregstats("ESET"))
	$sbstat = Not (_readregstats("SB"))
	$sasstat = Not (_readregstats("SAS"))
	$hcstat = Not (_readregstats("HC"))

	;return all scans = 0
	Return $cfstat & $mwbstat & $esetstat & $sbstat & $sasstat & $hcstat
EndFunc   ;==>_autoallclear
#cs ---------------------------------------------------------------------------
	FUNCTION: _autorun()

	PURPOSE: Convenience function to replace shortcode with method call

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
EndFunc   ;==>_autorun
#cs ---------------------------------------------------------------------------
	FUNCTION: _getnextscan()

	PURPOSE: Determines the next scan to run based on the last scan ran/skipped

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _getnextscan()

	;set error to 0 so we can make sure the key was opened successfully
	SetError(0)

	;grab the LastScan key from HKEY_LOCAL_MACHINE\SOFTWARE\ResTech\AutoStats
	$shortcode = RegRead("HKLM\SOFTWARE\ResTech\AutoStats", "LastScan")

	;if valid shortcode, use it to determine next scan from last scan
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
		Return "CF" ;start from the beginning if nothing has been defined, just in case
	EndIf
EndFunc   ;==>_getnextscan
#cs ---------------------------------------------------------------------------
	FUNCTION: _writelastscan()

	PURPOSE: Stores information indicating the last scan ran using the registry

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _writelastscan($shortcode)

	;write the shortcode of the last scan to HKEY_LOCAL_MACHINE\SOFTWARE\ResTech\AutoStats as key LastScan
	RegWrite("HKLM\SOFTWARE\ResTech\AutoStats", "LastScan", "REG_SZ", $shortcode)
EndFunc   ;==>_writelastscan
#cs -----------------------------------------------------------------------------
	FUNCTION: _writeregstats()

	PURPOSE: Creates and or modifies a registry entry of the given $shortcode

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: A Long Time Ago

	NOTES: Not used. Implemented for future use in Auto mode. Log scans by using
	shortcodes like MWB, ESET, etc. Value should be numerical
#ce -----------------------------------------------------------------------------
Func _writeregstats($shortcode, $value)

	;write the results of last scan to HKEY_LOCAL_MACHINE\SOFTWARE\ResTech\AutoScans by shortcode
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

	;set error to 0 to determine successful regread
	SetError(0)

	;read the key from HKEY_LOCAL_MACHINE\SOFTWARE\ResTech\AutoStats by shortcode
	$value = RegRead("HKLM\SOFTWARE\ResTech\AutoStats", $shortcode)

	;If valid read, return it as a number (strings are for losers, but registry readibility)
	If @error = 0 Then
		Return Number($value)
	Else
		Return -1
	EndIf
EndFunc   ;==>_readregstats
#EndRegion ###AutoHelpers
#Region ###Helpers
#Region ###Helpers-ErrorHandling
#cs ---------------------------------------------------------------------------
	FUNCTION: _debugwaitclick()

	PURPOSE: Perform a WinWait and ControlClick in succession with error handling
	for any issues that arise

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _debugwaitclick($title, $control, $text = "", $timeout = 300000, $button = "left", $clicks = 1, $x = 0, $y = 0)

	;wait for the window and grab its handle to verify
	$hWnd = WinWait($title, $text, $timeout)

	;if no window found, log it and return failure (Returncode 0)
	If $hWnd = 0 Then
		_appendlog(2, "AutoIt", 'No window found with Title "' & $title & '" and text "' & $text & '"')
		Return 0
	EndIf

	;Activate the window for compatibility, then wait for it to become active/update
	WinActivate($title, $text)
	Sleep(100)

	;Perform a ControlClick and record what it returns
	$ccresult = ControlClick($title, $text, $control, $button, $clicks, $x, $y)

	;if controlclick failed, log it and return failure (0)
	If $ccresult = 0 Then
		_appendlog(2, "AutoIt", 'No control found with ClassnameNN "' & $control & '", in window with Title "' & $title & '" and text "' & $text & '"')
		Return 0
	EndIf

	;if we've made it this far, return success (Returncode 1)
	Return 1
EndFunc   ;==>_debugwaitclick
#cs ---------------------------------------------------------------------------
	FUNCTION: _debugprocesswait()

	PURPOSE: Handle errors in a ProcessWait statement

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _debugprocesswait($process, $timeout = 60000)

	;get the pid of our open process
	$proc = ProcessWait($process, $timeout)

	;if it isn't open or it's hiding from AutoIt, Log and fail
	If $proc = 0 Then
		_appendlog(2, "AutoIt", $process & ' did not open in ' & ($timeout / 1000) & ' seconds')
		Return 0
	EndIf

	;otherwise return success
	Return 1
EndFunc   ;==>_debugprocesswait
#cs ---------------------------------------------------------------------------
	FUNCTION: _debugprocesswaitclose()

	PURPOSE: Handle errors in a ProcessWaitClose statement

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _debugprocesswaitclose($process, $timeout = 1800000)

	;get the pid of our non-open process
	$proc = ProcessWaitClose($process, $timeout)

	;if the proc didn't open log and fail.
	If $proc = 0 Then
		_appendlog(2, "AutoIt", $process & ' did not close in ' & ($timeout / 6000) & ' minutes')
		Return 0

		;if zero was not returned, determine if process existed, reset error, log and fail
	ElseIf @extended = 0xCCCCCCCC Then
		_appendlog(2, "AutoIt", $process & ' did not exist.')
		SetError(0)
		Return 0
	EndIf

	;otherwise return success
	Return 1
EndFunc   ;==>_debugprocesswaitclose
#cs ---------------------------------------------------------------------------
	FUNCTION: _debugwinwait()

	PURPOSE: Perform a WinWait and ControlClick in succession with error handling
	for any issues that arise

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _debugwinwait($title, $text = "", $timeout = 300000)

	;get handle of window we're waiting for
	$hWnd = WinWait($title, $text, $timeout)

	;if it doesn't open, log and fail
	If $hWnd = 0 Then
		_appendlog(2, "AutoIt", 'No window found with Title "' & $title & '" and text "' & $text & '"')
		Return 0
	EndIf

	;otherwise success
	Return 1
EndFunc   ;==>_debugwinwait
#EndRegion ###Helpers-ErrorHandling
#cs ---------------------------------------------------------------------------
	FUNCTION: _start()

	PURPOSE: Handle GUI updates and log entries to start a function

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _start($program)
	;log program start
	_appendlog(1, $program)
	;update and marquee the progress bar and taskbar thumb, update label
	GUICtrlSetData($proglabel, $program & " Running")
	GUICtrlSetStyle($progbar, 8)
	GUICtrlSendMsg($progbar, $pbm_setmarquee, True, 20)
	_ITaskBar_SetProgressState($form, 1)
EndFunc   ;==>_start
#cs ---------------------------------------------------------------------------
	FUNCTION: _end()

	PURPOSE: Handle GUI updates and log entries to finish a function

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _end($program)
	;update and demarquee progbar and taskbar thumb, update label.
	_appendlog(4, $program)
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
	_ITaskBar_SetProgressValue($form, 0)
	_ITaskBar_SetProgressState($form, 0)
EndFunc   ;==>_end
Func _error()
	_ITaskBar_SetProgressState($form, 4)
	_ITaskBar_SetProgressValue($form, 100)
EndFunc

Func _unerror()
	_ITaskBar_SetProgressState($form, 1)
EndFunc
#cs ---------------------------------------------------------------------------
	FUNCTION: _close()

	PURPOSE: Exits the program. Invoked by $GUI_EVENT_CLOSE

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: A Long Time Ago

	NOTES:
#ce -----------------------------------------------------------------------------
Func _close()
	Exit ;that was complex
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
	;define vars to hold version and arch
	Local $r = ""
	Local $s = ""

	;generate version string from OSVersion
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

	;generate architecture from OSArch
	If (@OSArch = "X86") Then
		$s = "32-bit"
	ElseIf (@OSVersion = "IA64") Then
		$s = "64-bit Itanium"
	Else
		$s = "64-bit"
	EndIf

	;Concatenate and return
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
	;create an input box to get a number
	Local $newtix = InputBox("Enter Ticket Number", "Please enter the computer's current ticket number:")

	;run validation on string
	If (Not Number($newtix) = 0) Then

		;write it to registry and update everything
		RegWrite("HKLM\SOFTWARE\ResTech", "TicketNo", "REG_SZ", $newtix)
		_updatelastopen()
		$ticketno = $newtix
	Else
		;pester the user until they enter a good ticket
		MsgBox(48, "ResTool NXT: Ticket Number", "You entered an invalid ticket number or did not enter a number at all. Please try again.")
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
	;put the current date & time in as last open time for ResTool
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

	;get logfile "'X':\Logs\ticketno.txt"
	Local $log = StringLeft(@ScriptDir, 3) & "Logs\" & $ticketno & ".txt"

	;if file doesn't exist, create it
	If (Not FileExists($log)) Then
		_FileCreate($log)
	EndIf

	;open the logfile
	$logfile = FileOpen($log, 1)
	If (Not $logfile = -1) Then

		;get the timestamp for the start of our log entry
		Local $logentry = _NowCalc()

		;complete the rest of the entry & concatenate
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
		;add extra messages if necessary
		If (Not ($msg = "" Or $code = 5 Or $code = 3)) Then
			$logentry = $logentry & " The program also generated the following message: " & $msg
		EndIf

		;write to file and close
		FileWriteLine($logfile, $logentry)
		FileClose($logfile)
	Else

		;tell the user that everything is broken
		MsgBox(0, "ResTool NXT: Logfile", "Error opening logfile. You should probably restart ResTool.")
	EndIf
EndFunc   ;==>_appendlog
#cs -----------------------------------------------------------------------------
	FUNCTION: _queuestartup()

	PURPOSE: Creates a shortcut to run ResTool at startup

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: A Long Time Ago

	NOTES: Really need an _unqueuestartup. Not Used. Planned for forced-restarts
	like in MWB and chkdsk /f
#ce -----------------------------------------------------------------------------
Func _queuestartup()
	;create a shortcut in the public startup directory to ResTool
	FileCreateShortcut(@ScriptFullPath, @StartupCommonDir & "\RT.lnk")
EndFunc   ;==>_queuestartup
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
EndFunc   ;==>_autocf
#cs -----------------------------------------------------------------------------
	FUNCTION: _automwb()

	PURPOSE: Runs Malwarebytes for Auto Mode

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _automwb()
EndFunc   ;==>_automwb
#cs -----------------------------------------------------------------------------
	FUNCTION: _autoeset()

	PURPOSE: Runs Eset for Auto Mode

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _autoeset()
EndFunc   ;==>_autoeset
#cs -----------------------------------------------------------------------------
	FUNCTION: _autosb()

	PURPOSE: Runs SpyBot for Auto Mode

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _autosb()
EndFunc   ;==>_autosb
#cs -----------------------------------------------------------------------------
	FUNCTION: _autosas()

	PURPOSE: Runs SuperAntiSpyware for Auto Mode

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _autosas()
EndFunc   ;==>_autosas
#cs -----------------------------------------------------------------------------
	FUNCTION: _autohc()

	PURPOSE: Runs HouseCall for Auto Mode

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _autohc()
EndFunc   ;==>_autohc
#cs -----------------------------------------------------------------------------
	FUNCTION: _autocc()

	PURPOSE: Runs CCleaner for Auto Mode

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _autocc()
EndFunc   ;==>_autocc
#cs -----------------------------------------------------------------------------
	FUNCTION: _autosfc()

	PURPOSE: Runs System File Check for Auto Mode

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/11/14

	NOTES: Can't currently handle the log.
#ce -----------------------------------------------------------------------------
Func _autosfc()
EndFunc   ;==>_autosfc
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
	_start("ComboFix")
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
	While (True)
		If (ProcessExists("conhost.exe")) Then
			ProcessWaitClose("conhost.exe")
		Else
			Sleep(1000)
			If Not (ProcessExists("conhost.exe")) Then
				ExitLoop
			EndIf
		EndIf
	WEnd
	;parse logfile

	;open file
	Dim $cflog = FileOpen("C:\COMBOFIX.TXT")
	;advance to the Other Deletions portion of the scan
	If (Not ($cflog = -1)) Then
		Dim $str = FileReadLine($cflog)
		While (StringInStr($str, "Other Deletions") = 0 And Not (@error = -1))
			$str = FileReadLine($cflog)
		WEnd
		;count entries between Other Deletions and files created
		Dim $count = 0
		If Not (@error = -1) Then
			$str = FileReadLine($cflog)
			While (StringInStr($str, "(((((((((((") = 0 And Not (@error = -1))
				$str = FileReadLine($cflog)
				If (Not (StringLeft($str, 1) = ".") And Not (StringLeft($str, 1) = "(")) Then
					$count += 1
				EndIf
			WEnd
		EndIf
		MsgBox(0, "ResTool NXT: ComboFix", "ComboFix made " & $count & " deletion(s).")
		_appendlog(3, "ComboFix", $count & " deletion(s).")
	Else
		MsgBox(0, "ResTool NXT: ComboFix", "Unable to open ComboFix log file.")
		_appendlog(2, "ComboFix", "Log file not found")
	EndIf
	If ProcessExists("notepad.exe") Then
		ProcessClose("notepad.exe")
	EndIf
	_end("ComboFix")
EndFunc   ;==>_runcf

#cs -----------------------------------------------------------------------------
	FUNCTION: _runmwb()

	PURPOSE: Runs Malwarebytes

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 1/30/15

	NOTES: Now with logging support!
#ce -----------------------------------------------------------------------------
Func _runmwb()
	_start("Malwarebytes")
	;install if necessary
	If (Not FileExists(@ProgramsCommonDir & "\Malwarebytes' Anti-Malware\Malwarebytes Anti-Malware.lnk")) Then
		_installmwb()
	EndIf
	;run the mwb executable and wait for it to open
	ShellExecute($32progfiledir & "\Malwarebytes' Anti-Malware\mbam.exe")
	While (Not ProcessExists("mbam.exe"))
		Sleep(100)
	WEnd
	WinWaitActive("Malwarebytes")
	If (WinActive("Malwarebytes Anti-Malware", "latest version")) Then
		Sleep(100)
		Send("{ESC}")
	EndIf
	;handle the update dialogs, if present. First select update to latest version
	If (WinActive("Malwarebytes Anti-Malware", "outdated")) Then
		Send("{ENTER}")
		Sleep(5000);wait 20 seconds
		;if error, log it and try to fix it
		If (WinActive("Malwarebytes Anti-Malware", "An error has occurred")) Then
			_error()
			_appendlog(3, "Malwarebytes", "Failed to Update")
			;if user doesn't wish to continue, tie up loose ends
			If ($idno == MsgBox($mb_yesno, "ResTool NXT: Malwarebytes", "The computer is not connected to the proxy and could not update. Continue with MWB scan?")) Then
				Send("{ENTER}")
				Sleep(1000)
				Send("!{F4}")
				_end("Malwarebytes")
				Return 0
			Else
				Send("{ENTER}")
			EndIf
			_unerror()
		;if no error, get rid of the update to 2.x dialog and start the real update
		ElseIf (WinActive("Malwarebytes Anti-Malware", "latest version")) Then
			ControlClick("Malwarebytes Anti-Malware", "", 2)
			WinWait("Malwarebytes Anti-Malware", "database was successfully")
			Send("{ENTER}")

		;for the random case where the 2.x dialog doesn't show up
		Else
			WinWait("Malwarebytes Anti-Malware", "database was successfully")
			Send("{ENTER}")
		EndIf
	Sleep(5000)
	EndIf

	;wait for update completion

	;get rid of an occasional 2.x dialog that pops up again
	If (WinActive("Malwarebytes Anti-Malware", "latest version")) Then
		ControlClick("Malwarebytes Anti-Malware", "", 2)
	EndIf

	;start a full scan
	WinWait("Malwarebytes Anti-Malware", "Perform")
	ControlClick("Malwarebytes Anti-Malware", "Perform", 69)
	ControlClick("Malwarebytes Anti-Malware", "Perform", 71)
	WinWait("Full scan", "")
	Send("{ENTER}")

	;wait for it to finish (it opens notepad for results)
	While (Not ProcessExists("notepad.exe"))
		Sleep(100)
	WEnd
	If (WinExists("Malwarebytes Anti-Malware", "No malicious items")) Then
		WinClose("Malwarebytes Anti-Malware", "No malicious items");close dialog
		;we're clear; log results and exit
		MsgBox(0, "ResTool NXT: Malwarebytes", "MWB found 0 infected files.")
		_appendlog(3, "Malwarebytes", "0 infected files")
		;exit
	ElseIf (WinExists("Malwarebytes Anti-Malware", "Show Results")) Then
		ControlClick("Malwarebytes Anti-Malware", "Show Results", "Button1");close dialog
		;we've got items. parse log to determine how many
		Local $wt = WinGetTitle("mbam")
		Local $fname = @AppDataDir & "\Malwarebytes\Malwarebytes' Anti-Malware\Logs\" & StringLeft($wt, StringLen($wt) - 10)& '.txt'
		Local $fd = FileOpen($fname)
		Local $str = FileRead($fd)
		$str = StringTrimLeft($str, StringInStr($str, "Memory P"))
		Local $astr = _StringBetween($str, ": ", @CRLF); get an array of # every ': # @CRLF" match
		Local $result = 0
		For $i = 0 To UBound($astr) -1; sum all array matches via a loop
			$result += Int($astr[$i])
		Next
		MsgBox(0, "ResTool NXT: Malwarebytes", "MWB found " & $result & " infected files.")
		_appendlog(3, "Malwarebytes", $result & " infected files")
		WinClose("mbam")
		;Remove all
		ControlClick("Malwarebytes Anti-Malware", "", "ThunderRT6CommandButton2");show results
		ControlClick("Malwarebytes Anti-Malware", "", "ThunderRT6UserControlDC1", "right");context menu
		Sleep (100)
		Send("c");select check all
		Sleep (100)
		Send("{ENTER}");enter
		ControlClick("Malwarebytes Anti-Malware", "", "ThunderRT6CommandButton2");remove selected
		EndIf
	_end("Malwarebytes")
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

	;fast-forward through the boring stuff
	WinWait("Select Setup Language", "")
	Send("{ENTER}")
	WinWait("Setup - Malwarebytes Anti-Malware", "Welcome")
	Send("{ENTER}")
	WinWait("Setup - Malwarebytes Anti-Malware", "License")
	Send("{TAB}A{ENTER}");tabs to radios, hits A, hits enter
	WinWait("Setup - Malwarebytes Anti-Malware", "Information")
	Send("{ENTER}")
	WinWait("Setup - Malwarebytes Anti-Malware", "Select Destination")
	Send("{ENTER}")
	Sleep(1000)

	;someone already made a MWB folder... uh-oh!
	If WinExists("Folder Exists") Then
		Send("{ENTER}")
	EndIf

	;more minutae to {ENTER} through
	WinWait("Setup - Malwarebytes Anti-Malware", "Select Start")
	Send("{ENTER}")
	WinWait("Setup - Malwarebytes Anti-Malware", "Select Additional")
	Send("{ENTER}")
	WinWait("Setup - Malwarebytes Anti-Malware", "Ready")
	Send("{ENTER}")
	Sleep(1000)

	;we've got a restart on our hands.
	If WinExists("Setup - Malwarebytes Anti-Malware", "Preparing") Then
		MsgBox(48, "ResTool NXT: Malwarebytes", "The computer is going to restart in order to complete a previous Malwarebytes install action.")
		Sleep(10000)
		_queuestartup() ;make ResTool start next boot.
		_appendlog(5, "Malwarebytes", "Requires restart to install")
		Send("{ENTER}")
	EndIf
	;finish up and make sure not to start uncontrollable opens/updates
	WinWait("Setup - Malwarebytes Anti-Malware", "Completing")
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
	_start("ESET Online Scanner")

	;run the executable
	Run(@ScriptDir & "\Script\Scanners\ESET\ESET.exe")

	;accept ToU
	WinWait("Terms of use", "")
	ControlClick("Terms of use", "", 107)
	ControlClick("Terms of use", "", 105)
	Sleep(10000)

	;someone forgot to feed the computer its internets
	If WinExists("Downloading ESET Online Scanner...", "Use custom proxy settings") Then
		If (_GetIP() = -1) Then
			_error()
			;pester the user until we've got a valid PUBLIC ip (good way to verify Net Connection)
			If $IDRETRY = MsgBox($MB_RETRYCANCEL, "ResTool NXT:  ESET Online Scanner", "Please connect the computer to the internet to continue") Then
				Sleep (1000)
				If (@IPAddress1 = "127.0.0.1") Then
					WinClose("Downloading ESET Online Scanner...")
					MsgBox(0, "ResTool NXT: ESET Online Scanner", "ESET scan aborted due to no network connection.")
					_appendlog(2, "ESET Online Scanner", "No Internet connection. Scan aborted.")
					_end ("ESET Online Scanner")
					Return 0
				Else
					ControlClick("Downloading ESET Online Scanner...", "", 109)
					_unerror()
				EndIf
			Else
				WinClose("Downloading ESET Online Scanner...")
				MsgBox(0, "ResTool NXT: ESET Online Scanner", "ESET scan aborted due to no network connection.")
				_appendlog(2, "ESET Online Scanner", "No Internet connection. Scan aborted.")
				_end ("ESET Online Scanner")
				Return 0
			EndIf
		Else
			;try again if valid ip (not sure if this is the best option)
			ControlClick("Downloading ESET Online Scanner...", "", 109)
		EndIf
	EndIf

	;fill out PUP dialog and start
	WinWait("ESET Online Scanner", "unwanted")
	Sleep(1000)
	ControlClick("ESET Online Scanner", "unwanted", 346)
	ControlClick("ESET Online Scanner", "unwanted", 321)

	;wait for completion
	$esethWnd = WinWaitActive("ESET Online Scanner", "Finished")
	;grab the screen text, then get the 9th line, which has the results
	$results = StringSplit(WinGetText("ESET Online Scanner"), @LF, 1)[8]
	;finish eset
	ControlClick("ESET Online Scanner", "", 220)
	MsgBox(0, "ResTool NXT: ESET Online Scanner", "ESET found " & $results & " Infected Files.")
	_appendlog(4, "ESET Online Scanner", $results & " Infected Files.")
	WinClose($esethWnd)
	_end("ESET Online Scanner")
EndFunc   ;==>_runeset

#cs -----------------------------------------------------------------------------
	FUNCTION: _runsb()

	PURPOSE: Runs SpyBot Search and Destroy

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 2/5/15

	NOTES: Not fully tested. Think the results might not work
#ce -----------------------------------------------------------------------------
Func _runsb()
	;if not installed, make it so (the install is pretty self explanatory)
	If (Not FileExists($32progfiledir & "\Spybot - Search & Destroy 2\SDUpdate.exe")) Then
		If (_installsb() = 0) Then
			Return 0
		EndIf
	EndIf
	_start("Spybot")
	;Begin Updates
	GUICtrlSetData($proglabel, "Updating Spybot")
	ShellExecute($32progfiledir & "\Spybot - Search & Destroy 2\SDUpdate.exe")
	While (ControlCommand("Update (Spybot", "", "TButton1", "IsEnabled") = 0)
		Sleep(100)
	WEnd
	Sleep(2000)
	If (WinExists("Update (Spybot", "There are updates available!")) Then
		_debugwaitclick("Update (Spybot", "TButton1")
		_debugwinwait("Update (Spybot", "+++")
		_appendlog(3, "Spybot", "Updates were installed")
		WinClose("Update (Spybot")
	Else
		WinClose("Update (Spybot")
		If (_GetIP() = -1) Then
			_error()
			MsgBox(0, "ResTool NXT: Spybot", "Spybot could not update because the network was not connected. Please connect the computer to the Internet and restart Spybot.")
			_appendlog(2, "Spybot", "Updates failed due to no network connection")
			_end("Spybot")
			Return 0
		Else
			_appendlog(3, "Spybot", "No Updates were found")
		EndIf
	EndIf

	;Begin Scans
	GUICtrlSetData($proglabel, "Spybot Scan Started")
	ShellExecute($32progfiledir & "\Spybot - Search & Destroy 2\SDScan.exe")
	_debugwinwait("System Scan (Spybot", "Start a scan")
	_debugwaitclick("System Scan (Spybot", "TButton3")
	Sleep(1040)
	;Handle "Do you need help?" dialog
	If (WinExists("Spybot - Search & Destroy 2", "JSDialog")) Then
		_debugwaitclick("Spybot - Search & Destroy 2", "TButton1")
	EndIf
	;Begin Cleanup
	_debugwinwait("System Scan (Spybot", "Scan Took", 3600000)
	If (WinExists("System Scan (Spybot", "The antispyware scan came up without results.")) Then
			MsgBox(0, "ResTool NXT: Spybot", "Spybot found 0 results.")
			_appendlog(3, "Spybot", "Spybot found 0 results.")
	Else
		;read result scans
		Local $text = WinGetText("System Scan (Spybot")
		$text = StringRight($text, StringInStr($text, "The antispyware scan found ") + 27)
		$text = StringLeft($text, StringInStr($text, "results") - 1)
		MsgBox(0, "ResTool NXT: Spybot", "Spybot found " & $text & " results.")
		_appendlog(3, "Spybot", "Spybot found " & $text & " results.")
		_debugwaitclick("System Scan (Spybot", "TButton3")
		Sleep(1000)
			_debugwaitclick("System Scan (Spybot", "TButton1", "Scan finished.")
			;wait for cleaning to finish
			While(ControlCommand("System Scan (Spybot", "", "TButton1", "IsEnabled"))
				Sleep(100)
			WEnd
	EndIf
	WinClose("System Scan (Spybot")
	_end("Spybot")
EndFunc   ;==>_runsb
#cs -----------------------------------------------------------------------------
	FUNCTION: _installsb()

	PURPOSE: Installs Super Anti Spyware

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 2/5/15

	NOTES: It's a pretty standard install. Handles the restart required dialog that
		   will only occur if the program is uninstalled and installed before restart.
#ce -----------------------------------------------------------------------------
Func _installsb()
	_start("Spybot Install")
	_debugwaitclick("Select Setup Language", "TNewButton1")
	_debugwaitclick("Setup - Spybot - Search & Destroy", "TNewButton1", "Welcome")
	_debugwaitclick("Setup - Spybot - Search & Destroy", "TNewButton2", "Donations")
	_debugwaitclick("Setup - Spybot - Search & Destroy", "TNewButton2", "Installation")
	_debugwaitclick("Setup - Spybot - Search & Destroy", "TNewRadioButton1", "License")
	_debugwaitclick("Setup - Spybot - Search & Destroy", "TNewButton2", "License")
	_debugwaitclick("Setup - Spybot - Search & Destroy", "TNewButton2", "Ready to")
	If (WinExists("Setup - Spybot - Search & Destroy", "Preparing to Install")) Then
		_appendlog(2, "Spybot", "Restart Required")
		If($IDYES = MsgBox($MB_YESNO, "ResTool NXT: Spybot Requires Restart", "You must restart to install Spybot. Restart Now?")) Then
			_queuestartup()
			ShellExecute("shutdown", "/r /t 0", "", "", @SW_HIDE)
		Else
			_end("Spybot Install")
			Return 0
		EndIf
	EndIf
	_debugwaitclick("Setup - Spybot - Search & Destroy", "TNewCheckListBox1", "Completing", 300000, "left", 1, 10, 10)
	_debugwaitclick("Setup - Spybot - Search & Destroy", "TNewButton2", "Completing")
	_end("Spybot Install")
EndFunc
#cs -----------------------------------------------------------------------------
	FUNCTION: _runsas()

	PURPOSE: Runs Super Anti Spyware

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 2/6/15

	NOTES: Not tested, but I think it should work fully.
#ce -----------------------------------------------------------------------------
Func _runsas()
	_start("SuperAntiSpyware")
	If (Not FileExists("C:\Program Files\SUPERAntiSpyware\SUPERAntiSpyware.exe")) Then
		_debugwinwait("SUPERAntiSpyware Setup", "Custom")
		;Click the custom install button on the main window
		WinMove("SUPERAntiSpyware Setup", "", 0, 0)
		WinActivate("SUPERAntiSpyware Setup", "")
		Sleep(100)
		MouseClick("left", 355, 324)
		;Go through the setup dialogs: Welcome
		_debugwaitclick("SUPERAntiSpyware Free Edition Setup", "Button8", "WARNING:")
		;Go through the setup dialogs: EULA
		_debugwaitclick("SUPERAntiSpyware Free Edition Setup", "Button8", "license agreement below to")
		;Go through the setup dialogs: User Selection
		_debugwaitclick("SUPERAntiSpyware Free Edition Setup", "Button8", "settings for this application")
		;Go through the setup dialogs: Location
		_debugwaitclick("SUPERAntiSpyware Free Edition Setup", "Button8", "Setup will install the files")
		;Actual install progress is here
		;Go through the setup dialogs: Configuration
		_debugwaitclick("SUPERAntiSpyware Free Edition Setup", "Button1", "following options");Disable Post-install update. We'll do it ourselves
		_debugwaitclick("SUPERAntiSpyware Free Edition Setup", "Button2", "following options");Disable Automatic Update (Pro bloatware)
		_debugwaitclick("SUPERAntiSpyware Free Edition Setup", "Button8", "following options")
		;Go through the setup dialogs: Finish
		_debugwaitclick("SUPERAntiSpyware Free Edition Setup", "Button8", "Finished")
	EndIf
	;update
	ProcessClose("SUPERAntiSpyware.exe")
	Local $inetObj = InetGet("http://cdn.superantispyware.com/SASDEFINITIONS.EXE", @TempDir & "\SASDEFS.EXE", 1, 0)
	InetClose($inetObj)
	Run(@TempDir & "\SASDEFS.EXE")
	_debugwaitclick("SUPERAntiSpyware Definition Update", "Button1", "OK")
	_debugwaitclick("Action Required", "Button1", "Please restart")
	Sleep(1000)
	;start SAS
	Run("C:\Program Files\1SUPERAntiSpyware\SUPERAntiSpyware.exe")
	Run("C:\Program Files\SUPERAntiSpyware\SUPERAntiSpyware.exe")
	;Actually Scan
	_debugwinwait("SUPERAntiSpyware Definition Update")
	_debugwaitclick("SUPERAntiSpyware Free Edition", "Button3", "Scans your computer");select complete scan
	_debugwaitclick("SUPERAntiSpyware Free Edition", "Button1");Press Start Scan...
	_debugwaitclick("SUPERAntiSpyware Free Edition", "Button23", "all running items");Activate High Scan Boost
	_debugwaitclick("SUPERAntiSpyware Free Edition", "Button21");Start Scan
	;wait up to an hour for the scan to complete
	_debugwinwait("SUPERAntiSpyware Scan Summary", "SUPER", 3600000);threats detected window
	#include <String.au3>
	Local $results = WinGetText("SUPERAntiSpyware Scan Summary", "")
	$results = StringTrimLeft($results, (StringInStr($results, "Detected:") + 9))
	$results = StringLeft($results, StringInStr($results, @LF))
	MsgBox(0, "ResTool NXT: SuperAntiSpyware", "SAS Found " & $results & "threats.")
	_debugwaitclick("SUPERAntiSpyware Scan Summary", "Button1")
	_debugwaitclick("SUPERAntiSpyware Free Edition", "Button34", "Place a checkmark");remove selected items
	_debugwaitclick("SUPERAntiSpyware Quarantine Complete", "Button1", "following items")
	_debugwaitclick("SuperAntiSpyware Free Edition", "Button36", "Processed Items")
	WinClose("SUPERAntiSpyware Free Edition")
	_end("SuperAntiSpyware")
EndFunc   ;==>_runsas
#cs -----------------------------------------------------------------------------
	FUNCTION: _runhc()

	PURPOSE: Runs Housecall

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 11/24/14

	NOTES: Now works perfectly fine, but requires a slight bit of cheating
#ce -----------------------------------------------------------------------------
Func _runhc()
	_start("TrendMicro HouseCall")
	;determine version to run
	If ($osv = "X86") Then
		$ver = "32"
	Else
		$ver = "64"
	EndIf
	;start the right version
	Run(@ScriptDir & "\Script\Scanners\HC\HC" & $ver & ".exe")
	WinWait("HouseCall Download", "")
	Sleep(1000)
	If (WinActive("Trend Micro HouseCall", "Unable to complete the download.")) Then
		MsgBox(48, "ResTool NXT: Network", "Connect computer to the internet and click OK")
		_runhc()
	Else
		WinWaitClose("HouseCall Download")
		Sleep(1000)
		;grab the window as an Internet Explorer object for manipulation
		$handle = _IEAttach(WinWaitActive("[CLASS:#32770]"), "embedded")
		$accept = _IEGetObjById($handle, "r_eula_0")
		$next = _IEGetObjById($handle, "btn_next")
		;handle EULA already accepted (above calls set @error to 7 if so)
		If Not (@error = 7) Then
			_IEAction($accept, "click")
			_IEAction($next, "click")
		Else
			;set error to 0 for compatibility
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
		_end("TrendMicro HouseCall")
	EndIf
EndFunc   ;==>_runhc

#cs -----------------------------------------------------------------------------
	FUNCTION: _runcc()

	PURPOSE: Runs CCleaner

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 11/18/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _runcc()
	_start("CCleaner")
	;install CCleaner if necessary
	If (Not FileExists(@ProgramFilesDir & "\CCleaner\CCleaner.exe")) Then
		Run(@ScriptDir & "\Script\CC.exe")
		WinWait("CCleaner Professional Setup", "Welcome")
		ControlClick("CCleaner Professional Setup", "", 1)
		WinWait("CCleaner Professional Setup", "Automatically check for updates to CCleaner")
		Send("{ENTER}")
		If (WinExists("CCleaner Professional Setup", "Free! Google")) Then
			ControlClick("CCleaner Professional Setup", "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "left", 1, 29, 156)
		EndIf
		Send("{ENTER}")
		WinWait("CCleaner Professional Setup", "Release notes")
		ControlClick("CCleaner Professional Setup", "", 1204)
		ControlClick("CCleaner Professional Setup", "", 1)
		WinWait("CCleaner Professional", "")
		WinClose("CCleaner Professional")
	;otherwise run the correct architecture
	Else
		If ($osv = "X86") Then
			Run(@ProgramFilesDir & "\CCleaner\CCleaner.exe")
		Else
			Run(@ProgramFilesDir & "\CCleaner\CCleaner64.exe")
		EndIf
	EndIf
	;perform file cleanup
	WinWait("CCleaner Professional", "")
	WinClose("CCleaner Professional")
	WinWait("Piriform CCleaner - Professional Edition", "")
	ControlClick("Piriform CCleaner - Professional Edition", "", 1021)
	WinWait("", "This process will ")
	Send("{ENTER}")
	WinWait("Piriform CCleaner - Professional Edition", "CLEANING")

	;Results are located on the 9th line of text grabbed from the window
	$results = StringSplit(WinGetText("Piriform CCleaner - Professional Edition"), @LF, 1)[8]

	MsgBox(0, "ResTool NXT: CCleaner", $results)
	_appendlog(4, "CCleaner File Cleaner", $results) ; mark these results in the log
	Send("g");switches to registry (on a good day)
	Local $rresults = 1
	$title = "Piriform CCleaner - Professional Edition"
	$text = "Scan for Issues"
	;While we haven't gotten a 0 entry scan, run a scan
	While (Not $rresults = 0)
		WinWait($title, $text)
		WinActivate($title, $text)
		ControlClick($title, $text, "Button2")
		Sleep(100)
		;wait for scan to complete
		While (Not ControlGetText($title, $text, "Button2") = "&Scan for Issues")
			Sleep(100)
		WEnd
		;if Fix Results is enabled, we will do that
		If ControlCommand($title, $text, "Button3", "IsEnabled") Then
			ControlClick($title, $text, "Button3")
			WinWait("CCleaner", "")
			ControlClick("CCleaner", "", 6)

			;save the file in default location (usually My Docs)
			WinWait("Save As", "")
			Sleep(1000)
			Send("{ENTER}")

			;start fixing issues
			$text = "Fix All Selected Issues"
			WinWait("", $text)
			Sleep(100)
			$rresults = 0
			;For each entry, click the Fix button and add a count to rresults
			While (ControlCommand("", $text, "Button3", "IsEnabled"))
				ControlClick("", $text, "Button3")
				$rresults += 1
				Sleep(10)
			WEnd
			;close the dialog once all entries fixed
			ControlClick("", $text, "Button5")
			;return result to user
			MsgBox(0, "ResTool NXT: CCleaner", $rresults & " Registry Entries")
		;Fix Results not enabled, no errors to fix
		Else
			;set value to 0 and return to user
			MsgBox(0, "ResTool NXT: CCleaner", "0 Registry Entries")
			$rresults = 0
		EndIf
		;reset text back to what it was at the start of the loop.
		$text = "Scan for Issues"
		;log things
		_appendlog(4, "CCleaner Registry Cleaner", $rresults & " Entries")
	WEnd
	_end("CCleaner")
EndFunc   ;==>_runcc
#cs -----------------------------------------------------------------------------
	FUNCTION: _runmwbar()

	PURPOSE: Runs Malwarebytes' Anti Malware

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 12/4/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _runmwbar()
	_start("Malwarebytes Anti Rootkit")
	;run the executable
	Run(@ScriptDir & "\Script\Scanners\MWBAR.exe")
	;I honesstly have no clue what I did here, but it works in most cases
	;It's mostly boring window-handling stuff, but if someone wants to document it...
	WinWait("Malwarebytes Anti-Rootkit", "")
	ControlClick("Malwarebytes Anti-Rootkit", "", "Button2")
	WinWait("Malwarebytes Anti-Rootkit BETA v1.08.2.1001", "")
	Send("!n")
	Sleep(1000)
	Send("{TAB}{TAB}{ENTER}")
	Sleep(60000);wait for update for an entire minute
	Send("{TAB}{TAB}{TAB}{ENTER}")
	MsgBox($MB_TOPMOST, "ResTool NXT: MWBAR", "Close this message box when MWBAR is finished scanning completely.")
	Send("!n")
	;there was ABSOLUTELY NO WAY to get the results from the window, and MWBAR is stupid about logfiles.
	;I'm sure there is a way, but we don't use it enough for it to necessitate the time to find one.
	MsgBox($MB_TOPMOST, "ResTool NXT: MWBAR", "Results should be showing. Record them and then close this dialog")
	MsgBox($MB_TOPMOST, "ResTool NXT: MWBAR", "Don't close this until you've actually recorded the results!")
	Send("!e")
	_end("Malwarebytes Anti Rootkit")
EndFunc   ;==>_runmwbar
#cs -----------------------------------------------------------------------------
	FUNCTION: _runtdss()

	PURPOSE: Runs TDSSKiller

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 11/20/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _runtdss()
	_start("TDSSKiller")
	;make some variables that we will use later.
	Local $aTDSSLog
	Dim $suscount = 0, $infcount = 0
	;open the executable with some flags to expedite the process (no windows!)
	ShellExecuteWait(@ScriptDir & "\Script\Scanners\TDSS.exe", "-silent -l C:\TDSSAI.txt -accepteula -accepteulaksn")
	;if we've got a log as expected
	If FileExists("C:\TDSSAI.txt") Then
		;if we can't read it, log failure
		If Not _FileReadToArray("C:\TDSSAI.txt", $aTDSSLog) Then
			MsgBox(16, "ResTool NXT: TDSSKiller", "Error parsing results")
			_appendlog(2, "TDSSKiller", "Error parsing results")
		;else parse it
		Else
			;for every Suspicious or Infected file, add a counter respectively
			For $i = 1 To UBound($aTDSSLog) - 1
				If StringInStr($aTDSSLog[$i], "Suspicious") Then
					$suscount += 1
				ElseIf StringInStr($aTDSSLog[$i], "infected") Then
					$infcount += 1
				EndIf
			Next
			;return these results nicely and log them
			MsgBox(0, "ResTool NXT: TDSSKiller", $suscount & " suspicious files, and " & $infcount & " infected files.")
			_appendlog(3, "TDSSKiller", $suscount & " suspicious / " & $infcount & " infected")
		EndIf
	EndIf
	_end("TDSSKiller")
EndFunc   ;==>_runtdss
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
	_start("NAC Uninstall")
	;Clicks buttons in the uninstaller. Pretty straightforward.
	ShellExecute("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Cisco\Cisco NAC Agent\Uninstall Cisco NAC Agent.lnk")
	WinWait("Windows Installer", "")
	ControlClick("Windows Installer", "", "Button1")
	Sleep(1000)
	WinWait("Windows Installer", "")
	WinWait("Cisco NAC Agent ", "")
	WinWaitClose("Cisco NAC Agent ", "")
	_end("NAC Uninstall")
EndFunc   ;==>_rmnac
#cs -----------------------------------------------------------------------------
	FUNCTION: _anyprint()

	PURPOSE: Installs Anywhere Printer support

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 11/20/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _anyprint()
	_start("Pharos Install")
	;Clicks buttons in an installer. Pretty straightforward
	Run(@ScriptDir & "\Script\Installers\Pharos.exe")
	WinWait("Package ""Windows AnywherePrint Client"" installer.", "")
	WinActivate("[LAST]")
	ControlClick("Package ""Windows AnywherePrint Client"" installer.", "", "Button1")
	WinWait("Package ""Windows AnywherePrint Client"" installer.", "Install")
	ControlClick("Package ""Windows AnywherePrint Client"" installer.", "", "Button1")
	WinWait("Package ""Windows AnywherePrint Client"" installer.", "Finish")
	ControlClick("Package ""Windows AnywherePrint Client"" installer.", "", "Button2")
	_end("Pharos Install")
EndFunc   ;==>_anyprint
#cs -----------------------------------------------------------------------------
	FUNCTION: _tempfr()

	PURPOSE: Uses the Disk Cleanup tool to remove temp files

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 11/20/14

	NOTES: Removes default selected files, since in most cases unchecked entries
	only add up to a couple mb of space or are handled by CCleaner
#ce -----------------------------------------------------------------------------
Func _tempfr()
	_start("Disk Cleanup")
	;run the file removal thing
	Run(@SystemDir & "\cleanmgr.exe")
	;select the default selection, hope no one has a weirdly named sys drive
	WinWait("Disk Cleanup", "")
	WinWait("Disk Cleanup for  (C:)", "")
	WinActivate("Disk Cleanup for  (C:)", "")
	;start the cleanup
	ControlClick("Disk Cleanup for  (C:)", "", "Button4")
	WinWait("Disk Cleanup", "Delete Files")
	;delete the files we don't want
	ControlClick("Disk Cleanup", "Delete Files", "Button1")
	ProcessWaitClose("cleanmgr.exe")
	_end("Disk Cleanup")
EndFunc   ;==>_tempfr

#cs -----------------------------------------------------------------------------
	FUNCTION: _openticket()

	PURPOSE: Opens the ticket

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 11/20/14

	NOTES: Opens in default browser
#ce -----------------------------------------------------------------------------
Func _openticket()
	;A great example of how awesome ShellExecute is. It literally will open anything
	;Just like Windows would.
	ShellExecute("http://intranet.restech.niu.edu/tickets/" & $ticketno)
EndFunc   ;==>_openticket
#cs -----------------------------------------------------------------------------
	FUNCTION: _regedit()

	PURPOSE: Opens the Registry Editor

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 11/20/14

	NOTES: Opens the registry associated with program arch. 64 gets 64 reg, 32 to 32
#ce -----------------------------------------------------------------------------
Func _regedit()
	;most complex line of code in the entire program
	ShellExecute("regedit")
EndFunc   ;==>_regedit
#cs -----------------------------------------------------------------------------
	FUNCTION: _togglemsconfig()

	PURPOSE: Toggles safe boot

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 11/18/14

	NOTES:
#ce -----------------------------------------------------------------------------
Func _togglemsconfig()
	;open msconfig
	Run("msconfig")
	;wait for it to pop up, then foreground it
	WinWait("System Configuration", "")
	WinActivate("System Configuration", "")
	Sleep(100)
	;move to next tab with Alt-Tab (at least I think ^ means alt)
	Send("^{TAB}")
	;if we're enabling it
	If (Not (_readregstats("MSCFG") == 1)) Then
		;click some buttons
		ControlClick("System Configuration", "", "Button5")
		Sleep(100)
		ControlClick("System Configuration", "", "Button9")
		Sleep(100)
		ControlClick("System Configuration", "", "Button30")
		;tell it not to restart, because we do what we want
		WinWait("System Configuration", "You may need to restart your computer")
		ControlClick("System Configuration", "", "Button3")
		;record it in the registry, set the color in the GUI
		_writeregstats("MSCFG", 1)
		GUICtrlSetColor($msconfig, $COLOR_RED)
	;if we're disabling it
	Else
		;click some buttons
		ControlClick("System Configuration", "", "Button5")
		;update the registry
		_writeregstats("MSCFG", 0)
		Sleep(100)
		ControlClick("System Configuration", "", "Button30")
		;tell it not to restart, because we do what we want.
		WinWait("System Configuration", "You may need to restart your computer")
		ControlClick("System Configuration", "", "Button3")
		;Eliminate color from the GUI, because nobody gets a happy GUI
		GUICtrlSetColor($msconfig, $COLOR_BLACK)
	EndIf
EndFunc   ;==>_togglemsconfig
#cs -----------------------------------------------------------------------------
	FUNCTION: _speedtest

	PURPOSE: Opens Speedtest in Internet Explorer

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: A Long Time Ago

	NOTES: Opens in default browser
#ce -----------------------------------------------------------------------------
Func _speedtest()
	ShellExecute("http://speedtest.niu.edu") ;self explanatory
EndFunc   ;==>_speedtest

#cs ---------------------------------s--------------------------------------------
	FUNCTION: _netadapterproperties()

	PURPOSE: Opens Network Adapter Settings

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: A Long Time Ago

	NOTES:
#ce -----------------------------------------------------------------------------
Func _netadapterproperties()
	ShellExecute("Ncpa.cpl") ;Ncpa is cpl module for Network Control Panel (Adapter Settings)
EndFunc   ;==>_netadapterproperties

#cs -----------------------------------------------------------------------------
	FUNCTION: _progfeat()

	PURPOSE: Opens Programs and Features

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: A Long Time Ago

	NOTES:
#ce -----------------------------------------------------------------------------
Func _progfeat()
	ShellExecute("appwiz.cpl") ;appwiz is cpl module for Programs and Features
EndFunc   ;==>_progfeat

#cs -----------------------------------------------------------------------------
	FUNCTION: _devmgmt()

	PURPOSE: Opens Device Management in MMC

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: A Long Time Ago

	NOTES:
#ce -----------------------------------------------------------------------------
Func _devmgmt()
	ShellExecute("devmgmt.msc") ;devmgmt is MMC module for Device Management
EndFunc   ;==>_devmgmt

#cs -----------------------------------------------------------------------------
	FUNCTION: _defraggle

	PURPOSE: Defragments the system drive

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: A Long Time Ago

	NOTES:
#ce -----------------------------------------------------------------------------
Func _defraggle()
	_start("Disk Defragmenter")
	;Start the defragmenter without a window
	ShellExecute("defrag", "/C /U", "", "", @SW_HIDE)
	Sleep(1000)
	;Wait for defrag to do its thing
	While (ProcessExists("Defrag.exe"))
		Sleep(100)
	WEnd
	_end("Disk Defragmenter")
EndFunc   ;==>_defraggle

#cs -----------------------------------------------------------------------------
	FUNCTION: _webreset

	PURPOSE: Performs an IPConfig reset

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: A Long Time Ago

	NOTES:
#ce -----------------------------------------------------------------------------
Func _webreset()
	;all input/output stuff has to be printed explicitly, since it is by stage
	_appendlog(1, "IPConfig Reset")
	GUICtrlSetData($proglabel, "Initializing IP Config Reset")
	;release
	ShellExecuteWait("ipconfig", "/release", "", "", @SW_HIDE)
	Sleep(100)
	GUICtrlSetData($progbar, 33)
	_ITaskBar_SetProgressValue($form, 33)
	GUICtrlSetData($proglabel, "Step 1 of 3 Complete")
	;flushdns
	ShellExecuteWait("ipconfig", "/flushdns", "", "", @SW_HIDE)
	Sleep(100)
	GUICtrlSetData($progbar, 66)
	_ITaskBar_SetProgressValue($form, 66)
	GUICtrlSetData($proglabel, "Step 2 of 3 Complete")
	;renew
	ShellExecute("ipconfig", "/renew", "", "", @SW_HIDE)
	Sleep(100)
	;wait for renew to complete
	While (ProcessExists("ipconfig.exe"))
		Sleep(100)
	WEnd
	;fix the GUI and log stuff
	GUICtrlSetData($progbar, 100)
	_ITaskBar_SetProgressValue($form, 100)
	GUICtrlSetData($proglabel, "Step 3 of 3 Complete")
	Sleep(1000) ;pause so we see the stuff
	GUICtrlSetData($progbar, 0)
	_ITaskBar_SetProgressValue($form, 0)
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
	;handle log and GUI data explicitly
	_appendlog(1, "Network Services Reset")
	ShellExecuteWait("netsh", "winsock reset", "", "", @SW_HIDE) ;winsock reset
	GUICtrlSetData($proglabel, "Step 1 of 5 Complete")
	GUICtrlSetData($progbar, 20)
	_ITaskBar_SetProgressValue($form, 20)
	ShellExecuteWait("netsh", "winsock reset catalog", "", "", @SW_HIDE) ;winsock catalog reset
	GUICtrlSetData($proglabel, "Step 2 of 5 Complete")
	GUICtrlSetData($progbar, 40)
	_ITaskBar_SetProgressValue($form, 40)
	ShellExecuteWait("netsh", "int ip reset", "", "", @SW_HIDE) ;ip settings reset
	GUICtrlSetData($proglabel, "Step 3 of 5 Complete")
	GUICtrlSetData($progbar, 60)
	_ITaskBar_SetProgressValue($form, 60)
	ShellExecuteWait("netsh", "interface ipv4 reset", "", "", @SW_HIDE) ;network interface reset
	GUICtrlSetData($proglabel, "Step 4 of 5 Complete")
	GUICtrlSetData($progbar, 80)
	_ITaskBar_SetProgressValue($form, 80)
	Select
		Case $osv = "WIN_XP"
			ShellExecuteWait("netsh", "firewall reset", "", "", @SW_HIDE);reset old firewall if xp
		Case Else
			ShellExecuteWait("netsh", "advfirewall reset", "", "", @SW_HIDE);reset firewall if newer
	EndSelect
	;finish the GUI updating stuff
	GUICtrlSetData($proglabel, "Step 5 of 5 Complete")
	GUICtrlSetData($progbar, 100)
	_ITaskBar_SetProgressValue($form, 100)
	Sleep(1000);wait so we see the stuff
	GUICtrlSetData($proglabel, "ResTool Ready...")
	GUICtrlSetData($progbar, 0)
	_ITaskBar_SetProgressValue($form, 0)
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
	;manual log and progress bar since steps
	_appendlog(1, "Hidden Network Adapter Removal")
	;determine version of devcon to use
	Local $devcon = @ScriptDir & "\Script\OOB\64\devcon.exe"
	If ($osa == "X86") Then
		$devcon = @ScriptDir & "\Script\OOB\32\devcon.exe"
	EndIf
	;remove Teredo Tunneling Pseudo-interface adapters
	GUICtrlSetData($proglabel, "Hidden Adapter: Step 1 of 4")
	ShellExecuteWait($devcon, '-r remove "@ROOT\*TEREDO\0000"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*TEREDO\0001"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*TEREDO\0002"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*TEREDO\0003"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*TEREDO\0004"', $devcon)
	GUICtrlSetData($progbar, 25)
	;remove ISATAP Pseudo-interface adapters
	_ITaskBar_SetProgressValue($form, 25)
	GUICtrlSetData($proglabel, "Hidden Adapter: Step 2 of 4")
	ShellExecuteWait($devcon, '-r remove "@ROOT\*ISATAP\0000"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*ISATAP\0001"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*ISATAP\0002"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*ISATAP\0003"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*ISATAP\0004"', $devcon)
	GUICtrlSetData($progbar, 50)
	;remove direct Tunneling adapters
	_ITaskBar_SetProgressValue($form, 50)
	GUICtrlSetData($proglabel, "Hidden Adapter: Step 3 of 4")
	ShellExecuteWait($devcon, '-r remove "@ROOT\*TUNMP\0000"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*TUNMP\0001"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*TUNMP\0002"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*TUNMP\0003"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*TUNMP\0004"', $devcon)
	GUICtrlSetData($progbar, 75)
	_ITaskBar_SetProgressValue($form, 75)
	;remove 6to4 conversion pseudo-interface adapters
	GUICtrlSetData($proglabel, "Hidden Adapter: Step 4 of 4")
	ShellExecuteWait($devcon, '-r remove "@ROOT\*6TO4\0000"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*6TO4\0001"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*6TO4\0002"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*6TO4\0003"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*6TO4\0004"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*6TO4\0005"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*6TO4\0006"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*6TO4\0007"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*6TO4\0008"', $devcon)
	ShellExecuteWait($devcon, '-r remove "@ROOT\*6TO4\0009"', $devcon)
	;fix the GUI data
	GUICtrlSetData($progbar, 100)
	_ITaskBar_SetProgressValue($form, 100)
	Sleep(1000)
	GUICtrlSetData($progbar, 0)
	_ITaskBar_SetProgressValue($form, 0)
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
	_start("Disk Check")
	;run and wait for close of chkdsk. Must  be done this way
	$var = Run(@ComSpec & " /c" & '%systemroot%\system32\chkdsk.exe C: > ' & @ScriptDir & "\Logs\chkdsk.txt", "", @SW_HIDE)
	ProcessWaitClose($var)
	;open the checkdisk log, read it and parse for errors
	$fd = FileOpen(@ScriptDir & "\Logs\chkdsk.txt")
	$text = FileRead($fd)
	If (StringInStr($text, "problems") = 0) Then
		MsgBox(0, "ResTool NXT: Disk Check", "Chkdsk did not find any errors")
		_appendlog(4, "Disk Check")
	Else
		;if errors, tell the drive it needs to be checked and ask about reboot
		If (MsgBox($mb_yesno, "ResTool NXT: Disk Check", "Should I schedule a repair at next boot?") == $idyes) Then
			ShellExecuteWait("fsutil", "dirty set c:")
			;if user wants to reboot, do it. Otherwise don't; Repair happens at next restart
			If (MsgBox($mb_yesno, "ResTool NXT: Disk Check", "Chkdsk is scheduled to repair the hard drive at next boot. Would you like to restart now?") == $idyes) Then
				_queuestartup()
				_appendlog(3, "Disk Check", "Errors Found, Restarting to fix.")
				ShellExecute("shutdown", "/r /t 0", "", "", @SW_HIDE)
			Else
				_appendlog(3, "Disk Check", "Errors found, Will fix at next restart.")
				MsgBox(0, "ResTool NXT: Disk Check", "Chkdsk will run the next time the computer restarts.")
			EndIf
		Else
			_appendlog(4, "Disk Check", "Errors Found but no repair queued.")
			MsgBox(0, "ResTool NXT: Disk Check", "Chkdsk will not fix disk errors.")
		EndIf
	EndIf
	FileClose($fd)
	FileDelete(@ScriptDir & "\Logs\chkdsk.txt")
	_end("Disk Check")
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
	_start("System File Check")
	;build our output handler for SFC data
	$var = Run(@ComSpec & " /C sfc /scannow", "", @SW_HIDE, $STDOUT_CHILD)
	ProcessWaitClose($var)
	$text = StdoutRead($var)
	;Parse the data for most-to-least unique information:
	;Failure to run/Success/Failure/Repair/None of the above
	If Not (StringInStr($text, "pending") = 0) Then
		MsgBox(0, "ResTool NXT: SFC", "SFC repairs or Windows Updates are required to continue.")
		_appendlog(2, "System File Check", "Failed due to pending updates")
	ElseIf Not (StringInStr($text, "not find any integrity violations") = 0) Then
		MsgBox(0, "ResTool NXT: SFC", "SFC did not find any corrupt files")
		_appendlog(4, "System File Check", "Did not find any corrupt files")
	ElseIf Not (StringInStr($text, "unable to fix") = 0) Then
		MsgBox(0, "ResTool NXT: SFC", "SFC found corrupt files but could not repair them.")
		_appendlog(2, "System File Check", "Found corrupt files but could not repair")
	ElseIf Not (StringInStr($text, "successfully repaired") = 0) Then
		MsgBox(0, "ResTool NXT: SFC", "SFC found and repaired corrupt files.")
		_appendlog(3, "System File Check", "Repaired corrupted files")
	Else
		MsgBox(0, "ResTool NXT: SFC (STATUS UNKNOWN)", $text)
		_appendlog(2, "System File Checker", "Failed with unknown status")
		_appendlog(5, "DEBUG INFO -SFC-", $text)
	EndIf
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	GUICtrlSetData($proglabel, "ResTool Ready...")
	_end("System File Check")
EndFunc   ;==>_sfc

#cs -----------------------------------------------------------------------------
	FUNCTION: _wifiprofileadd

	PURPOSE: Adds the WiFi

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: A Long Time Ago

	NOTES: Not sure if I actually have the xml bundled with ResTool
#ce -----------------------------------------------------------------------------
Func _wifiprofileadd()
	;requires manual GUI control
	_appendlog(1, "NIUwireless Profile Import")
	;remove NIUWireless
	ShellExecuteWait("netsh", "wlan delete profile name=NIUwireless", "", "", @SW_HIDE)
	GUICtrlSetData($progbar, 33)
	_ITaskBar_SetProgressValue($form, 33)
	GUICtrlSetData($proglabel, "Removed old NIUwireless profile")
	;remove NIUWirelessInstructions
	ShellExecuteWait("netsh", "wlan delete profile name=NIUwirelessInstructions", "", "", @SW_HIDE)
	GUICtrlSetData($progbar, 66)
	_ITaskBar_SetProgressValue($form, 66)
	GUICtrlSetData($proglabel, "Removed NIUwirelessInstructions")
	;add profile from XML
	ShellExecuteWait("netsh", "wlan add profile filename=" & @ScriptDir & "\Script\WiFi.xml user=all", "", "", @SW_HIDE)
	GUICtrlSetData($progbar, 100)
	_ITaskBar_SetProgressValue($form, 100)
	GUICtrlSetData($proglabel, "Added new NIUwireless profile.")
	Sleep(100)
	GuiCtrlSetData($proglabel, "Resetting network.")
	Sleep(1000)
	_webreset()
	GUICtrlSetData($progbar, 0)
	_ITaskBar_SetProgressValue($form, 0)
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
	ShellExecuteWait("control", "", "");runs Control Panel
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
	_start("DISM")
	;run DISM in hidden window - wait for completion
	ShellExecuteWait("dism", "/Online /Cleanup-Image /RestoreHealth", "", "", @SW_HIDE)
	GUICtrlSetStyle($progbar, 1)
	GUICtrlSetData($progbar, 0)
	;Catch errors and stuff: Failed, Fixed, No Error
	If (StringInStr(FileRead(@WindowsDir & "\Logs\dism\dism.log"), "Failed to restore the image health", 1)) Then
		_appendlog(2, "DISM", "DISM failed to reapair system file corruption.")
		MsgBox($mb_iconwarning, "ResTool NXT: DISM", "DISM failed to restore the image health.")
	Else
		_appendlog(4, "DISM")
		MsgBox(0, "ResTool NXT: DISM", "DISM Completed without error, but whether it made repairs or not is unknown.")
	EndIf
	_end("DISM")
EndFunc   ;==>_dism
#cs -----------------------------------------------------------------------------
	FUNCTION: _runaio

	PURPOSE: Runs Tweaking.com - Windows Repair All in One

	AUTHOR: Nobody

	DATE OF LAST UPDATE: 11/25/14

	NOTES: Uses silent mode. Not logging compliant.
#ce -----------------------------------------------------------------------------
Func _runaio()
	_start("All In One")
	;if not installed, install. Pretty Standard install
	If Not (FileExists($32progfiledir & "\Tweaking.com\Windows Repair (All in One)\Repair_Windows.exe")) Then
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
	;run it in silent mode. It should do everything by itself, unless it decides that it doesn't like something
	ShellExecuteWait($32progfiledir & "\Tweaking.com\Windows Repair (All in One)\Repair_Windows.exe", "/silent")
	_end("All In One")
EndFunc   ;==>_runaio
#cs -----------------------------------------------------------------------------
	FUNCTION: _getmse()

	PURPOSE: Installs MSE

	AUTHOR: Kevin Morgan

	DATE OF LAST UPDATE: 11/20/14

	NOTES: Does not activate yet on Win 8.1
#ce -----------------------------------------------------------------------------
Func _getmse()
	_start("MSE Install")
	;if MSE shouldn't be installed by default, run an install
	If (Not ($osv = "WIN_81") And Not ($osv = "WIN_8")) Then
		If ($osa == "X86") Then
			Run(@ScriptDir & "\Script\Installers\MSE32.exe")
		Else
			Run(@ScriptDir & "\Script\Installers\MSE64.exe")
		EndIf
		WinWait("Microsoft Security Essentials", "");first window
		ControlClick("Microsoft Security Essentials", "", "Button1")
		WinWait("Microsoft Security Essentials", "accept");EULA
		ControlClick("Microsoft Security Essentials", "", "Button1")
		WinWait("Microsoft Security Essentials", "join the program");Customer Improvement Program
		ControlClick("Microsoft Security Essentials", "", "Button2");checkbutton
		Sleep(100)
		ControlClick("Microsoft Security Essentials", "", "Button4");next
		WinWait("Microsoft Security Essentials", "optimize");update
		ControlClick("Microsoft Security Essentials", "", "Button2")
		WinWait("Microsoft Security Essentials", "may conflict with");AV warning
		ControlClick("Microsoft Security Essentials", "", "Button1")
		WinWait("Microsoft Security Essentials", "Finish");Finish
		ControlClick("Microsoft Security Essentials", "", "Button1")
	Else
		;Need to enable because Win 8. Not implemented
		MsgBox(0, "ResTool NXT: MSE Install", "MSE already installed as Windows Defender. Please enable from Control Panel.")
	EndIf
	_end("MSE Install")
EndFunc   ;==>_getmse
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
