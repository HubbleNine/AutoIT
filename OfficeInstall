;#SCRIPT# ======================================================================================================================
; Title .........: OfficeInstallTest
; AutoIt Version : 1.0.0.1
; Description ...: Installs O365 32 bit
; Return Codes ..: 0
; Author(s) .....: Me, Friend of mine
; Last Update ...: 21 March 2017
; #CHANGELOG# ===================================================================================================================
;
;================================================================================================================================
;Comments
;================================================================================================================================
;This script is used to uninstall Office 2003, 2007, 2010, 2013 and 2016 and then installs Office 365 apps
;================================================================================================================================
;Addition: Added OS Version checking to verify deployment will not progress on any machine with Server OS installed.
;================================================================================================================================

AutoItSetOption('MustDeclareVars',1)
AutoItSetOption('TrayIconHide',1)
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Constants.au3>
#include <TrayConstants.au3>

;***Checks if OS version for Server OS, if so, it will exit the script
If @OSVersion = "WIN_2016" Or @OSVersion = "WIN_2012R2" Or @OSVersion = "WIN_2012" Or @OSVersion = "WIN_2008R2" Or @OSVersion = "WIN_2008" Or @OSVersion = "WIN_2003" Then
Exit
EndIf

;***Checks if OS is 64 bit, if so, it will disable Program File redirection
;If @OSArch = "x64" Then
;_WinAPI_Wow64EnableWow64FsRedirection (False)
;EndIf

;***Compiles all necessary install bits and log file inside the executable
FileInstall ( ".\OffScrub03.vbs", @TempDir & "\OffScrub03.vbs", 1)
FileInstall ( ".\OffScrub07.vbs", @TempDir & "\OffScrub07.vbs", 1)
FileInstall ( ".\OffScrub10.vbs", @TempDir & "\OffScrub10.vbs", 1)
FileInstall ( ".\OffScrub13.vbs", @TempDir & "\OffScrub13.vbs", 1)
FileInstall ( ".\OffScrub16.vbs", @TempDir & "\OffScrub16.vbs", 1)
FileInstall ( ".\OffScrubC2R.vbs", @TempDir & "\OffScrubC2R.vbs", 1)
FileInstall ( ".\OfficeProPlus.exe", @TempDir & "\OfficeProPlus.exe", 1)

;***Runs uninstall scrub scripts for Office 2016 - 2003
;***and logs the success or failure of each uninstall
;***This runs a splash text during the uninstall
AutoItSetOption('TrayIconHide', 0)
SplashTextOn('Uninstalling Microsoft Office', 'Please wait while Microsoft Office is uninstalled', 400, 60, -1, -1, $DLG_MOVEABLE + $DLG_TEXTVCENTER, '', 10)
TrayTip('Microsoft Office Uninstall','Uninstall started.', 5, $TIP_ICONASTERISK)
ShellExecuteWait ('cscript.exe', @TempDir & '\OffScrub16.vbs ProPlus, Standard /Q /S /NoCancel', @SystemDir, "", @SW_HIDE)
ShellExecuteWait ('cscript.exe', @TempDir & '\OffScrub13.vbs ProPlus, Standard /Q /S /NoCancel', @SystemDir, "", @SW_HIDE)
ShellExecuteWait ('cscript.exe', @TempDir & '\OffScrub10.vbs ProPlus, Standard /Q /S /NoCancel', @SystemDir, "", @SW_HIDE)
ShellExecuteWait ('cscript.exe', @TempDir & '\OffScrub07.vbs ProPlus, Standard /Q /S /NoCancel', @SystemDir, "", @SW_HIDE)
ShellExecuteWait ('cscript.exe', @TempDir & '\OffScrub03.vbs ProPlus, Standard /Q /S /NoCancel', @SystemDir, "", @SW_HIDE)
ShellExecuteWait ('cscript.exe', @TempDir & '\OffScrubC2R.vbs ProPlus, Standard /Q /S /NoCancel', @SystemDir, "", @SW_HIDE)
SplashOff()

;***Runs the install msi for O365 Office 2016 Applications
;***and logs the success or failure to a network share
Local $rtnValue = RunWait (@TempDir & '\OfficeProPlus.exe')
	If $rtnValue = 0 Or $rtnValue = 3010 Then
		FileWrite ( "\\[path you want logs to go to]\Success.txt", @ComputerName & " installed O365" & @CRLF )
	Else
		FileWrite ( "\\[path you want logs to go to]\Failed.txt", @ComputerName & " failed to install O365" & @CRLF )
	EndIf

Shutdown(6)

Exit
