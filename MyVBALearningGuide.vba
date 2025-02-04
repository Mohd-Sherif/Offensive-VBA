' --------------------------------------------
' VBA Learning Guide - Comprehensive Examples
' --------------------------------------------
' This file contains examples of various VBA concepts.
' Use the MainMenu to choose which example to run.
' --------------------------------------------

Public globalVar As String ' Global variable

Option Explicit

' --------------------------------------------
' Main Menu
' --------------------------------------------
Sub MainMenu()
    Dim choice As Integer
    choice = InputBox("Welcome to the VBA Learning Guide!" & vbCrLf & _
                      "Choose an example to run:" & vbCrLf & _
                      "1. Error Message" & vbCrLf & _
                      "2. Print Numbers (For Loop)" & vbCrLf & _
                      "3. Add Numbers (Function)" & vbCrLf & _
                      "4. Input Box Example" & vbCrLf & _
                      "5. Message with Buttons (Conditional)" & vbCrLf & _
                      "6. Procedure Calls" & vbCrLf & _
                      "7. Variables Example (Local vs Global)" & vbCrLf & _
                      "8. Error Handling Example" & vbCrLf & _
                      "9. Interactive Add Numbers" & vbCrLf & _
                      "10. For Each Loop (Collections)" & vbCrLf & _
                      "11. Do While Loop" & vbCrLf & _
                      "12. Do Until Loop" & vbCrLf & _
                      "13. While Wend Loop" & vbCrLf & _
                      "14. Nested Loops" & vbCrLf & _
                      "15. Select Case Example" & vbCrLf & _
                      "16. Working with Arrays" & vbCrLf & _
                      "17. Exit", "VBA Learning Guide")
    
    Select Case choice
        Case 1: Call ErrorMessage
        Case 2: Call PrintNumbers
        Case 3: Call basicVBA
        Case 4: Call InputBoxExample
        Case 5: Call MessageWithButtons
        Case 6: Call FirstProcedure
        Case 7: Call FirstProcedureWithVariables
        Case 8: Call DivideNumbers
        Case 9: Call AddNumbersInteractive
        Case 10: Call ForEachLoopExample
        Case 11: Call DoWhileLoopExample
        Case 12: Call DoUntilLoopExample
        Case 13: Call WhileWendLoopExample
        Case 14: Call NestedLoopsExample
        Case 15: Call SelectCaseExample
        Case 16: Call WorkingWithArrays
        Case 17: Exit Sub
        Case Else: MsgBox "Invalid choice! Please try again.", vbExclamation, "Error"
    End Select
    
    ' Return to the main menu after running an example
    Call MainMenu
End Sub

' --------------------------------------------
' Message Boxes and User Interaction
' --------------------------------------------

' Displays a simple error message
Sub ErrorMessage()
    MsgBox "Something went wrong!", vbCritical + vbOKOnly, "Error"
End Sub

' Demonstrates how to use an InputBox to get user input
Sub InputBoxExample()
    Dim userName As String
    userName = InputBox("What is your name", "User Input")
    MsgBox "Hello, " & userName & "!"
End Sub

' Shows a message box with Yes/No buttons and handles the response
Sub MessageWithButtons()
    Dim response As Integer
    response = MsgBox("Do you want to continue?", vbYesNo + vbQuestion, "Confirmation")
    
    If response = vbYes Then
        MsgBox "Continue..."
    Else
        MsgBox "Abort..."
    End If
End Sub

' --------------------------------------------
' Loops and Iterations
' --------------------------------------------

' Demonstrates a simple For loop
Sub PrintNumbers()
    Dim i As Integer
    For i = 1 To 5
        MsgBox "Message Box Number: " & i
    Next i
End Sub

' Demonstrates a For Each loop with a collection
Sub ForEachLoopExample()
    Dim cell As Range
    For Each cell In Range("A1:A5")
        MsgBox "Cell Value: " & cell.Value
    Next cell
End Sub

' Demonstrates a Do While loop
Sub DoWhileLoopExample()
    Dim i As Integer
    i = 1
    Do While i <= 5
        MsgBox "Do While Loop: " & i
        i = i + 1
    Loop
End Sub

' Demonstrates a Do Until loop
Sub DoUntilLoopExample()
    Dim i As Integer
    i = 1
    Do Until i > 5
        MsgBox "Do Until Loop: " & i
        i = i + 1
    Loop
End Sub

' Demonstrates a While Wend loop
Sub WhileWendLoopExample()
    Dim i As Integer
    i = 1
    While i <= 5
        MsgBox "While Wend Loop: " & i
        i = i + 1
    Wend
End Sub

' Demonstrates nested loops
Sub NestedLoopsExample()
    Dim i As Integer, j As Integer
    For i = 1 To 3
        For j = 1 To 2
            MsgBox "Nested Loop: i = " & i & ", j = " & j
        Next j
    Next i
End Sub

' --------------------------------------------
' Functions and Procedures
' --------------------------------------------

' Calls a function to add two numbers
Sub basicVBA()
    Dim result As Integer
    result = AddNumbers(10, 40)
    MsgBox "The Sum is " & result
End Sub

' Function to add two numbers
Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function

' Demonstrates calling another procedure
Sub FirstProcedure()
    Call SecondProcedure
End Sub

Sub SecondProcedure()
    MsgBox "Second Procedure was called!"
End Sub

' --------------------------------------------
' Variables and Data Types
' --------------------------------------------

Sub FirstProcedureWithVariables()
    varA = 10 ' Implicit Declaration
    Dim varB As Integer
    varB = 20
    Dim localVar As String
    localVar = "Local Variable"
    globalVar = "Mohamed Sherif"
    MsgBox globalVar, vbInformation, "Global variable"
    MsgBox localVar, vbInformation, "Local variable"
End Sub

' --------------------------------------------
' Error Handling
' --------------------------------------------

' Demonstrates error handling with division
Sub DivideNumbers()
    On Error GoTo ErrorHandler
    Dim num1 As Integer, num2 As Integer, result As Double
    num1 = InputBox("Enter the first number", "Divide Numbers")
    num2 = InputBox("Enter the second number", "Divide Numbers")
    result = num1 / num2
    MsgBox "The result is " & result
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
End Sub

' --------------------------------------------
' Conditional Statements
' --------------------------------------------

' Demonstrates Select Case
Sub SelectCaseExample()
    Dim day As Integer
    day = InputBox("Enter a number between 1 and 7", "Day of the Week")
    
    Select Case day
        Case 1: MsgBox "Monday"
        Case 2: MsgBox "Tuesday"
        Case 3: MsgBox "Wednesday"
        Case 4: MsgBox "Thursday"
        Case 5: MsgBox "Friday"
        Case 6: MsgBox "Saturday"
        Case 7: MsgBox "Sunday"
        Case Else: MsgBox "Invalid day!"
    End Select
End Sub

' --------------------------------------------
' Arrays
' --------------------------------------------

' Demonstrates working with arrays
Sub WorkingWithArrays()
    Dim arr(3) As String
    arr(0) = "Apple"
    arr(1) = "Banana"
    arr(2) = "Cherry"
    
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        MsgBox "Array Value: " & arr(i)
    Next i
End Sub

' --------------------------------------------
' Interactive Examples
' --------------------------------------------

' Allows the user to input numbers and see the sum
Sub AddNumbersInteractive()
    Dim num1 As Integer, num2 As Integer
    num1 = InputBox("Enter the first number", "Add Numbers")
    num2 = InputBox("Enter the second number", "Add Numbers")
    MsgBox "The sum is " & AddNumbers(num1, num2)
End Sub

' --------------------------------------------
' Run the Main Menu when the file is opened
' --------------------------------------------
Sub AutoOpen()
    Call MainMenu
End Sub
