InputBox, MothraWindow, , Please enter the exact name of your Mothra window., , , , , , , ,%MothraWindow%
if ErrorLevel {
    MsgBox, You pressed cancel. Exiting program.
	exitapp
	}
else
    MsgBox, You entered "%MothraWindow%"`n`nPlease make sure Mothra Screens.xlsm is also open`n`nWindows+M (on mailID screen) will run the scraper`nControl+a will select all in Mothra`nControl+c in the Excel workbook on "Copy Screens" will grab the Mothra Screens and put it in the clipboard.

	
	


; If I'm looking at a putty window, modify the behavious or Control+C and Control+V	
#If WinActive(MothraWindow)
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
	
	;mailid and initialization for excel
	MouseClickDrag, Left, 9, 39, 647, 393
	WinActivate, Mothra Screens.xlsm - Excel
	Xl := ComObjActive("Excel.Application") ;creates a handle to your Application object
	xl.Run("CopyLastSearch")
	xl.Run("ClearData")
	Sleep, 150
	xl.Run("DisplayPeople")
	Sleep,500
	
	;people
	WinActivate, %MothraWindow%
	Send, mp%MothraID%{Enter}
	MouseClickDrag, Left, 9, 39, 647, 393
	WinActivate, Mothra Screens.xlsm - Excel
	xl.Run("DisplayPeople")
	Sleep, 500
	
	;IDs
	WinActivate, %MothraWindow%
	Send, mim%MothraID%{Enter}
	MouseClickDrag, Left, 9, 39, 647, 393
	WinActivate, Mothra Screens.xlsm - Excel
	xl.Run("DisplayPeople")
	Sleep, 500
	
	;objects
	WinActivate, %MothraWindow%
	Send, mo%MothraID%{Enter}
	MouseClickDrag, Left, 9, 39, 647, 393
	WinActivate, Mothra Screens.xlsm - Excel
	xl.Run("DisplayPeople")
	Sleep, 500
	xl.Run("MoveOScreenChecktoAHK")
	NextScreen = %Clipboard%
	NextScreenGo = More
	ifInString, NextScreen, %NextScreenGo%
	{
		WinActivate, %MothraWindow%
		Send, n
		MouseClickDrag, Left, 9, 106, 647, 345
		WinActivate, Mothra Screens.xlsm - Excel
		xl.Run("MoreObjects")
		Sleep, 500
	}
	xl.Run("MoveLoginIDtoAHK")
	Sleep, 150
	LoginID = %clipboard%
	
	;LoginID
	WinActivate, %MothraWindow%
	Send, qma%LoginID%
	MouseClickDrag, Left, 9, 39, 647, 393
	WinActivate, Mothra Screens.xlsm - Excel
	xl.Run("DisplayPeople")
	Sleep, 500
	xl.Run("MoveAScreenChecktoAHK")
	Sleep, 150
	NextScreen = %clipboard%
	WinActivate, %MothraWindow%
	Sleep, 300
	ifInString, NextScreen, %NextScreenGo%
	{
		Send, n
		MouseClickDrag, Left, 9, 233, 647, 393
		WinActivate, Mothra Screens.xlsm - Excel
		xl.Run("MorePermits")
	}
	
	;history
	WinActivate, %MothraWindow%
	Send, qmh%MothraID%{Enter}
	MouseClickDrag, Left, 9, 39, 647, 393
	WinActivate, Mothra Screens.xlsm - Excel
	xl.Run("DisplayPeople")
	Sleep, 500
	xl.Run("MoveHScreenChecktoAHK")
	NextScreen = %clipboard%
	ifInString, NextScreen, %NextScreenGo%
	{
		WinActivate, %MothraWindow%
		send, n
		MouseClickDrag, Left, 9, 120, 647, 345
		WinActivate, Mothra Screens.xlsm - Excel
		xl.Run("MoreHistory")
		Sleep, 300
	}
	
	;conclude
	xl.Run("ClearPII")
	clipboard = %OriginalClipboard%
	WinActivate, %MothraWindow%
	WinMove, %Title%,, %X%, %Y%, %Width%, %Height%
	;WinActivate, Mothra Screens.xlsm - Excel
	Return
	
	
	
	
; Control+A selects all on the mothra window
^a::
	WinGetPos, X, Y, Width, Height
	WinGetActiveTitle, Title
	WinRestore, %Title%
	WinMove, %Title%,,,, 675, 425
	MouseClickDrag, Left, 9, 39, 647, 403
	NewSelectAll = %Clipboard%
	StringReplace, NewSelectAll, NewSelectAll, `r`n`r`n, `r`n, UseErrorLevel
	StringReplace, NewSelectAll, NewSelectAll, `r`n`r`n, `r`n, UseErrorLevel
	StringReplace, NewSelectAll, NewSelectAll, `r`n`r`n, `r`n, UseErrorLevel
	Clipboard = %NewSelectAll%
	WinMove, %Title%,, %X%, %Y%, %Width%, %Height%
	Return




; Control+S saves the script and if the script is in focus, reloads the script
#IfWinActive, *C:\Users\eduque\Downloads\MyScript.ahk - Notepad++
~^s::
	sleep 100
	reload
	return
	

	
	
; If Mothra Screens.xlsm - Excel active do these things
#IfWinActive, Mothra Screens.xlsm
~^c::
	sleep, 150
	Xl := ComObjActive("Excel.Application") ;creates a handle to your Application object
	CheckCopyScreens = %clipboard%
	DisplayPeople = p-Display People
	DisplayMailID = F-Display MailID
	DisplayLoginID = A-Display LoginID
	DisplayIDs = I-Display IDs
	ListOwnedObjects = O-List Owned Objects
	NewCopy := ""
	NewCopyGo = 0
	ifInString, CheckCopyScreens, %DisplayPeople%
	{
		xl.Run("CopyP")
		NewCopy = %NewCopy%%Clipboard%
		NewCopyGo := true
	}
	ifInString, CheckCopyScreens, %DisplayMailID%
	{
		xl.Run("CopyF")
		NewCopy = %NewCopy%%Clipboard%
		NewCopyGo := true
	}
	ifInString, CheckCopyScreens, %DisplayLoginID%
	{
		xl.Run("CopyA")
		NewCopy = %NewCopy%%Clipboard%
		xl.Run("CopyA2")
		NewCopy = %NewCopy%%Clipboard%
		NewCopyGo := true
	}
	ifInString, CheckCopyScreens, %DisplayIDs%
	{
		xl.Run("CopyI")
		NewCopy = %NewCopy%%Clipboard%
		NewCopyGo := true
	}
	ifInString, CheckCopyScreens, %ListOwnedObjects%
	{
		xl.Run("CopyO")
		NewCopy = %NewCopy%%Clipboard%
		NewCopyGo := true
	}
	if NewCopyGo
	{
	StringReplace, NewCopy, NewCopy, `r`n`r`n, `r`n, UseErrorLevel
	clipboard = %NewCopy%
	}