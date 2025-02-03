Sub MessageWithButtons()
    Dim response As Integer
    response = MsgBox("Do you want to continue?", vbYesNo + vbQuestion, "Confirmation")
    
    If response = vbYes Then
        MsgBox "Continue..."
    Else
        MsgBox "Abort..."
    End If

End Sub
