AutoItSetOption('MustDeclareVars', 1)
AutoItSetOption('TrayIconHide', 1)
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Constants.au3>
#include <TrayConstants.au3>
#include <Date.au3>
Global $SunWknr = 0

;checks if Sunday
If @WDAY <> 1 Then Exit

Func thirdSunday()
	; check sunday number in this month
	If @MDAY < 8 Then
		$SunWknr = 1
	ElseIf @MDAY < 15 Then
		$SunWknr = 2
	ElseIf @MDAY < 22 Then
		$SunWknr = 3
	ElseIf @MDAY < 29 Then
		$SunWknr = 4
	Else
		$SunWknr = 5
	EndIf
EndFunc   ;==>thirdSunday

thirdSunday()

If $SunWknr <> 3 Then
	Exit
Else
	Shutdown(6)
EndIf

Exit
