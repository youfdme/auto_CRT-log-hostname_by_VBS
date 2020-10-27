#$language = "VBScript"
#$interface = "1.0"


' This automatically generated script may need to be
' edited in order to work correctly.


OD = array("show run","show ver","show ip int brief","dir","show fan")
'For i=0 To UBound(arr)-LBound(arr)
'MsgBox OD(i)
'Next

Sub Check_more()

	if (crt.Screen.WaitForString (" --More-- ",1)<>False) Then
	crt.Screen.Send "                              " & chr(13)
	Else
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	end if
end sub

sub Start_log()
	crt.Screen.Synchronous = True
	Call crt.Session.Log(False)
	Dim logfile
	logfile = "log\%H_%M%D_%h%m%s.log"
	crt.Session.LogFileName = logfile
	call crt.session.log (False)
	call crt.session.log (True)
	
	crt.sleep 1000
end sub

sub stop_log()
	Dim faile
	faile = "failed"
	if (crt.Screen.WaitForString ("#",3)<>False) Then
	call crt.session.log (false)
	MsgBox "Success log over!"
	else 
		MsgBox faile
		end if 
end sub

Sub Main
	Start_log()
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "#"
	For i=0 To UBound(OD)-LBound(OD)
	crt.Screen.WaitForString "#"
	order= OD(i)
	crt.Screen.Send order & chr(13)
	Check_more()
	crt.Screen.WaitForString "#"
	Next
	crt.Screen.WaitForString "#"
	stop_log()
	
End Sub
