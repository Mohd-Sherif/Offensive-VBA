' --------------------------------------------
' VBA Program Execution Demo - Comprehensive Examples
' --------------------------------------------
' This file demonstrates how to execute external programs,
' run PowerShell commands, and capture output using VBA.
' Use the MainMenu to choose which example to run.
' --------------------------------------------

Option Explicit

' --------------------------------------------
' Main Menu
' --------------------------------------------
Sub MainMenu()
    Dim choice As Integer
    choice = InputBox("Welcome to the VBA Program Execution Demo!" & vbCrLf & _
                      "Choose an example to run:" & vbCrLf & _
                      "1. Run Notepad (Shell)" & vbCrLf & _
                      "2. Run Notepad via PowerShell (inMemExec)" & vbCrLf & _
                      "3. Save Process List to File (PowerShell)" & vbCrLf & _
                      "4. Open Calculator (asyncExec)" & vbCrLf & _
                      "5. Run ipconfig and Display Output (winScript)" & vbCrLf & _
                      "6. Exit", "VBA Program Execution Demo")
    
    Select Case choice
        Case 1: Call RunShell
        Case 2: Call inMemExec
        Case 3: Call RunPowerShell
        Case 4: Call asyncExec
        Case 5: Call winScript
        Case 6: Exit Sub
        Case Else: MsgBox "Invalid choice! Please try again.", vbExclamation, "Error"
    End Select
    
    ' Return to the main menu after running an example
    Call MainMenu
End Sub

' --------------------------------------------
' Program Execution Examples
' --------------------------------------------

' Demonstrates running Notepad using the Shell function
Sub RunShell()
    On Error Resume Next
    Shell "notepad.exe", vbMaximizedFocus
    If Err.Number = 0 Then
        MsgBox "Notepad opened successfully.", vbInformation, "RunShell"
    Else
        MsgBox "Error: Failed to open Notepad. Error Code: " & Err.Number, vbCritical, "RunShell"
    End If
    On Error GoTo 0
End Sub

' Demonstrates running Notepad via PowerShell in memory
Sub inMemExec()
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    
    On Error Resume Next
    objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ""[System.Diagnostics.Process]::Start('notepad.exe')""", 0, True
    If Err.Number = 0 Then
        MsgBox "Notepad started in memory via PowerShell.", vbInformation, "inMemExec"
    Else
        MsgBox "Error: Failed to execute PowerShell command. Error Code: " & Err.Number, vbCritical, "inMemExec"
    End If
    On Error GoTo 0
End Sub

' Demonstrates saving a process list to a file using PowerShell
Sub RunPowerShell()
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    
    On Error Resume Next
    objShell.Run "powershell.exe -Command ""Get-Process | Out-File 'C:\Users\Mohamed Sherif\Desktop\processes.txt'""", vbHide, True
    If Err.Number = 0 Then
        MsgBox "Process list saved to 'processes.txt' on Desktop.", vbInformation, "RunPowerShell"
    Else
        MsgBox "Error: Failed to save process list. Error Code: " & Err.Number, vbCritical, "RunPowerShell"
    End If
    On Error GoTo 0
End Sub

' Demonstrates opening Calculator asynchronously
Sub asyncExec()
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    
    On Error Resume Next
    objShell.Run "calc.exe", 1, True
    If Err.Number = 0 Then
        MsgBox "Calculator opened successfully.", vbInformation, "asyncExec"
    Else
        MsgBox "Error: Failed to open Calculator. Error Code: " & Err.Number, vbCritical, "asyncExec"
    End If
    On Error GoTo 0
End Sub

' Demonstrates running ipconfig and displaying the output
Sub winScript()
    Dim output As String
    Dim objExec As Object
    Dim objShell As Object
    
    On Error Resume Next
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("cmd.exe /c ipconfig")
    
    If Err.Number = 0 Then
        Do While Not objExec.StdOut.AtEndOfStream
            output = output & objExec.StdOut.ReadLine() & vbCrLf
        Loop
        MsgBox "ipconfig executed successfully. Output: " & vbCrLf & output, vbInformation, "winScript"
    Else
        MsgBox "Error: Failed to execute ipconfig. Error Code: " & Err.Number, vbCritical, "winScript"
    End If
    On Error GoTo 0
End Sub

' --------------------------------------------
' Run the Main Menu when the file is opened
' --------------------------------------------
Sub AutoOpen()
    Call MainMenu
End Sub
