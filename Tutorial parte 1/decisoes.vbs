Option Explicit
Dim num1
Dim num2
Dim decisao

num1 = 20
num2 = 20

If (num1 < num2) Then
    decisao = 1
    ElseIf (num2 < num1) Then
    decisao = 2
    Else
End If
        
Select Case decisao
    Case 1
        MsgBox "O numero 1 é o menor valor"
    Case 2
        MsgBox "O numero 2 é o menor valor"
    Case Else
        MsgBox "Os numeros são iguais"
End Select