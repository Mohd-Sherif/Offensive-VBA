Sub AutoOpen()
    'Call winScript
    'Call asyncExec
    'Call RunPowerShell
    'Call inMemExec
    Call RunShell
End Sub

Sub RunShell()
    Shell "notepad.exe", vbMaximizedFocus
End Sub

Sub inMemExec()
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    
    objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ""[System.Diagnostics.Process]::Start('notepad.exe')""", 0, True
End Sub

Sub RunPowerShell()
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    
    objShell.Run "powershell.exe -Command ""Get-Process | Out-File 'C:\Users\Mohamed Sherif\Desktop\processes.txt'", vbHide, True
End Sub

Sub asyncExec()
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "calc.exe", 1, True
End Sub

Sub winScript()
    Dim output As String
    Dim objExec As Object
    Dim objShell As Object
    
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("cmd.exe /c ipconfig")
    
    Do While Not objExec.StdOut.AtEndOfStream
        output = output & objExec.StdOut.ReadLine() & vbCrLf
    Loop
    
    MsgBox output
End Sub
