Attribute VB_Name = "Module2"
Option Compare Database


Function StartServer()
    x = Shell("POWERSHELL.exe " & CurrentProject.Path & "\Start-WebServer_StaticFiles.ps1", vbNormalFocus)
End Function

Function StopServer()
    Call DoCmd.OpenForm("Quit")
    DoEvents
    Call DoCmd.Close(acForm, "Quit")
End Function

'Function StartServer()
' x = Shell("POWERSHELL.exe " & "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\GetObjectVBS.ps1", vbHide)
'End Function

