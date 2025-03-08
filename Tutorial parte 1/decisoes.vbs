Option Explicit
Dim num1
Dim num2
Dim decision

num1 = 20
num2 = 20

If (num1 < num2) Then
    decision = 1
ElseIf (num2 < num1) Then
    decision = 2
Else
End If

Select Case decision
    Case 1
        MsgBox "Number 1 is the smaller value"
    Case 2
        MsgBox "Number 2 is the smaller value"
    Case Else
        MsgBox "The numbers are equal"
End Select