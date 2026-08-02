Dim objShell, objWMIService, colProcess, objProcess
Set objShell = CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'WINWORD.EXE'")

If colProcess.Count > 0 Then
    ' اگر Word در حال اجرا است
    MsgBox "Microsoft Word is currently running. Please close it before proceeding.", vbExclamation, "Warning"
Else
    ' اگر Word در حال اجرا نیست
    objShell.Run "winword"
End If

Set objShell = Nothing
Set objWMIService = Nothing
Set colProcess = Nothing