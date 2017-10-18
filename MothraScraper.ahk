CoordMode, Mouse, Client
#Esc::
	exit

; If I'm looking at a putty window, modify the behavious or Control+C and Control+V	
#IfWinActive, ahk_exe putty.exe
{
	^v::
		send, +{Insert} ;Turns ctrl+v into shift+insert
		Return
	^c::Return ;doesn't allow control+c





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
	;MouseClickDrag, Left, 6, 24, 55, 25
	;VerifyCorrectScreen = %clipboard%
	;MouseClickDrag, Left, 108, 87, 168, 88
	;MothraID = %clipboard%
	;If (VerifyCorrectScreen != "MailID" or MothraID = "________")
	;{
	;	MsgBox Uh-oh`n`nScreen should be "MailID", Screen = "%VerifyCorrectScreen%"`nMothraID should't be "________", MothraID = "%MothraID%"
	;	Return
	;}
	
	;mailid and initialization for excel
	GrabScreen()
	WinActivate, Mothra Screens.xlsm - Excel
	
	Xl := ComObjActive("Excel.Application") ;creates a handle to your Application object
	xl.Run("CopyLastSearch")
	xl.Run("ClearData")
	Sleep, 150
	xl.Run("DisplayPeople")
	Sleep,500
	
	xl.Run("CheckForMothraID")
	MothraIDcheck = %Clipboard%
	MothraIDcheck2 = More
	ifInString, MothraIDcheck, %MothraIDcheck2% 
	{
		MsgBox, "No Mothra ID visble! Are you on the mailID screen?"
		Return
	}
	else 
		xl.Run("MoveMothraIDtoAHK")
	
	MothraID = %clipboard%
	;people
	WinActivate, AutoMothra
	Send, mp%MothraID%{Enter}
	GrabScreen()
	WinActivate, Mothra Screens.xlsm - Excel
	xl.Run("DisplayPeople")
	Sleep, 750
	
	;IDs
	WinActivate, AutoMothra
	Send, mim%MothraID%{Enter}
	GrabScreen()
	WinActivate, Mothra Screens.xlsm - Excel
	xl.Run("DisplayPeople")
	Sleep, 500
	
	;objects
	WinActivate, AutoMothra
	Send, mo%MothraID%{Enter}
	Sleep, 500
	GrabScreen()
	WinActivate, Mothra Screens.xlsm - Excel
	xl.Run("DisplayPeople")
	Sleep, 500
	xl.Run("MoveOScreenChecktoAHK")
	NextScreen = %Clipboard%
	NextScreenGo = More
	ifInString, NextScreen, %NextScreenGo%
	{
		WinActivate, AutoMothra
		Send, n
		GrabScreen()
		WinActivate, Mothra Screens.xlsm - Excel
		xl.Run("MoreObjects")
		Sleep, 500
	}
	xl.Run("MoveLoginIDtoAHK")
	Sleep, 150
	LoginID = %clipboard%
	
	;LoginID
	WinActivate, AutoMothra
	Send, qma%LoginID%
	GrabScreen()
	WinActivate, Mothra Screens.xlsm - Excel
	xl.Run("DisplayPeople")
	Sleep, 500
	xl.Run("MoveAScreenChecktoAHK")
	Sleep, 150
	NextScreen = %clipboard%
	WinActivate, AutoMothra
	Sleep, 300
	ifInString, NextScreen, %NextScreenGo%
	{
		Send, n
		GrabScreen()
		WinActivate, Mothra Screens.xlsm - Excel
		xl.Run("MorePermits")
	}
	
	;history
	WinActivate, AutoMothra
	Send, qmh%MothraID%{Enter}
	GrabScreen()
	WinActivate, Mothra Screens.xlsm - Excel
	xl.Run("DisplayPeople")
	Sleep, 500
	xl.Run("MoveHScreenChecktoAHK")
	NextScreen = %clipboard%
	ifInString, NextScreen, %NextScreenGo%
	{
		WinActivate, AutoMothra
		send, n
		GrabScreen()
		WinActivate, Mothra Screens.xlsm - Excel
		xl.Run("MoreHistory")
		Sleep, 300
	}
	
	;conclude
	xl.Run("ClearPII")
	clipboard = %OriginalClipboard%
	WinActivate, AutoMothra
	WinMove, %Title%,, %X%, %Y%, %Width%, %Height%
	WinActivate, Mothra Screens.xlsm - Excel
	Return
	
	
	
	
; Control+A selects all on the mothra window
^a::
	WinGetPos, X, Y, Width, Height
	WinGetActiveTitle, Title
	WinRestore, %Title%
	WinMove, %Title%,,,, 675, 425
	GrabScreen()
	StringReplace, Clipboard, Clipboard, `r`n`r`n, `r`n, UseErrorLevel
	StringReplace, Clipboard, Clipboard, `r`n`r`n, `r`n, UseErrorLevel
	StringReplace, Clipboard, Clipboard, `r`n`r`n, `r`n, UseErrorLevel
	WinMove, %Title%,, %X%, %Y%, %Width%, %Height%
	Return

}


; Control+S saves the script and if the script is in focus, reloads the script
#IfWinActive, *C:\Users\eduque\OneDrive - UC Davis\Scripts\MothraScraperTest.ahk - Notepad++
~^s::
	sleep 100
	reload
	return
	




	WindowGetRect(AutoMothra) {
    if hwnd := WinExist(AutoMothra) {
        VarSetCapacity(rect, 16, 0)
        DllCall("GetClientRect", "Ptr", hwnd, "Ptr", &rect)
        return {width: NumGet(rect, 8, "Int"), height: NumGet(rect, 12, "Int")}
		}
	}
	
	
	GrabScreen(){
		rect := WindowGetRect("AutoMothra")
		Sleep, 300
		MouseClickDrag, left, 1, 1, rect.width, rect.height
		NewSelectAll = %Clipboard%
		StillLoading := "__"
		ObjectsLoading := "Enter Searchable ID Number"
		Loop, 3
		{
			ifInString, NewSelectAll, %ObjectsLoading%,
			{
				sleep, 300
				MouseClickDrag, left, 1, 1, rect.width, rect.height
				NewSelectAll = %Clipboard%
				continue
			}
			ifInString, NewSelectAll, %StillLoading%
			{
				sleep, 300
				MouseClickDrag, left, 1, 1, rect.width, rect.height
				NewSelectAll = %Clipboard%
				continue
			}
			else 
				break
		}
	Clipboard = %NewSelectAll%
	}
	
; If Mothra Screens.xlsm - Excel active do these things
#IfWinActive, Mothra Screens.xlsm
{
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
	return

}