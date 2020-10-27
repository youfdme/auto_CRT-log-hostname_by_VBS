#$language = "VBScript"
#$interface = "1.0"




' This automatically generated script may need to be
' edited in order to work correctly.

Const ForReading  =  1, ForWriting  =  2, ForAppending  =  8

Mt="#"
'便于其他设备修改提示符

	Dim logfile
	logfile = "log\%H_%M%D_%h%m%s.log"

Sub Check_more()
	if (crt.Screen.WaitForString (" --More-- ",1)<>False) Then
	crt.Screen.Send "                              " & chr(13)
	Check_more()
	Else
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	end if
end sub

sub Start_log()
	crt.Screen.Synchronous = True
	Call crt.Session.Log(False)
	crt.Session.LogFileName = logfile
	call crt.session.log (False)
	call crt.session.log (True)
	crt.sleep 200
end sub

sub stop_log()
	Dim faile
	faile = "failed"
	success = "Successful ! "
	if (crt.Screen.WaitForString (Mt,3)<>False) Then
	call crt.session.log (false)
	MsgBox success
	else 
		MsgBox faile
		end if 
end sub


Sub Main
	Start_log()
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString Mt
	Dim fso
	Set fso  =  CreateObject("Scripting.FileSystemObject")
	Set file1  =  fso.OpenTextFile("orders.txt",Forreading, False)
	DO While file1.AtEndOfStream <> True
	line  =  file1.ReadLine
	crt.Screen.Send line & chr(13)
	Check_more()
	crt.Screen.WaitForString Mt
	loop
	crt.Screen.WaitForString Mt
	stop_log()
	
End Sub



