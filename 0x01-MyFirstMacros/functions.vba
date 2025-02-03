Sub basicVBA()
    Dim result As Integer
    result = AddNumbers(10, 40)
    MsgBox "The Sum is " & result
End Sub

Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function
