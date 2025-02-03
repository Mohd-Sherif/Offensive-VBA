Public globalVar As String

Sub FirstProcedure()
    varA = 10 ' Implicit Declaration
    Dim varB As Integer
    varB = 20
    Dim localVar As String
    localVar = "Local Variable"
    globalVar = "Sherif"
    MsgBox globalVar, vbInformation, "Global variable"
    MsgBox localVar, vbInformation, "Local variable"
End Sub
