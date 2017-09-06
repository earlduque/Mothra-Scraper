; If I'm looking at a putty window, modify the behavious or Control+C and Control+V	
#IfWinActive, ahk_exe putty.exe
{
	^v::
		send, +{Insert} ;Turns ctrl+v into shift+insert
		Return
	^c::Return ;doesn't allow control+c
}




; Windows+M does Mothra Magic, make sure to have a mothra window active on the mailid screen (with a valid mothra id visible) and the Mothra Screens.xlsm workbook open
#m::
	WinGetPos, X, Y, Width, Height
	WinGetActiveTitle, Title
	WinRestore, %Title%
	WinMove, %Title%,,,, 675, 425
	IfWinNotExist, Mothra Screens.xlsm - Excel
	{
		MsgBox Please open Mothra Screens.xlsm
		Return
	}
	OriginalClipboard = %clipboard%
	clipboard := ""
	MouseClickDrag, Left, 17, 56, 64, 57
	VerifyCorrectScreen = %clipboard%
	MouseClickDrag, Left, 115, 120, 175, 120
	MothraID = %clipboard%
	If (VerifyCorrectScreen != "MailID" or MothraID = "________")
	{
		MsgBox Uh-oh`n`nScreen should be "MailID", Screen = "%VerifyCorrectScreen%"`nMothraID should't be "________", MothraID = "%MothraID%"
		Return
	}
	MouseClickDrag, Left, 9, 39, 647, 393
	WinActivate, Mothra Screens.xlsm - Excel
	Xl := ComObjActive("Excel.Application") ;creates a handle to your Application object
	xl.Run("CopyLastSearch")
	xl.Run("ClearData")
	Sleep, 150
	xl.Run("DisplayPeople")
	Sleep,300
	WinActivate, AutoMothra
	Send, mp%MothraID%{Enter}
	MouseClickDrag, Left, 9, 39, 647, 393
	WinActivate, Mothra Screens.xlsm - Excel
	xl.Run("DisplayPeople")
	Sleep, 300
	WinActivate, AutoMothra
	Send, mim%MothraID%{Enter}
	MouseClickDrag, Left, 9, 39, 647, 393
	WinActivate, Mothra Screens.xlsm - Excel
	xl.Run("DisplayPeople")
	Sleep, 300
	WinActivate, AutoMothra
	Send, mo%MothraID%{Enter}
	MouseClickDrag, Left, 9, 39, 647, 393
	WinActivate, Mothra Screens.xlsm - Excel
	xl.Run("DisplayPeople")
	Sleep, 300
	xl.Run("MoveOScreenChecktoAHK")
	NextScreen = %Clipboard%
	NextScreenGo = More
	ifInString, NextScreen, %NextScreenGo%
	{
		WinActivate, AutoMothra
		Send, n
		MouseClickDrag, Left, 9, 106, 647, 345
		WinActivate, Mothra Screens.xlsm - Excel
		xl.Run("MoreObjects")
		Sleep, 300
	}
	xl.Run("MoveLoginIDtoAHK")
	Sleep, 150
	LoginID = %clipboard%
	WinActivate, AutoMothra
	Send, qma%LoginID%
	MouseClickDrag, Left, 9, 39, 647, 393
	WinActivate, Mothra Screens.xlsm - Excel
	xl.Run("DisplayPeople")
	Sleep, 300
	xl.Run("MoveAScreenChecktoAHK")
	Sleep, 150
	NextScreen = %clipboard%
	WinActivate, AutoMothra
	Sleep, 300
	ifInString, NextScreen, %NextScreenGo%
	{
		Send, n
		MouseClickDrag, Left, 9, 233, 647, 393
		WinActivate, Mothra Screens.xlsm - Excel
		xl.Run("MorePermits")
	}
	xl.Run("ClearSSN")
	clipboard = %OriginalClipboard%
	WinActivate, AutoMothra
	WinMove, %Title%,, %X%, %Y%, %Width%, %Height%
	Return
	
	
	
	
; Control+A selects all on the mothra window
^a::
	MouseClickDrag, Left, 9, 39, 647, 393
	Return
