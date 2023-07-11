Option Explicit


'Write a macro that prompts the user to enter a number, and then displays all the even numbers up to that number using a Do Until loop.
Sub ULoop1()

Dim i As Integer
i = InputBox("Enter a number")

Dim x As Integer
x = 1

Do Until x >= i
    If x Mod 2 = 0 Then
    MsgBox (x)
    End If
x = x + 1
Loop


End Sub