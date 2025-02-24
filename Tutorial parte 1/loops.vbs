Option Explicit

Dim num1, num2, i
num1 = 1
num2 = 10

Do While num1 < num2
    num1 = num1 + 2
    If (num1 = 7) Then 
    WScript.Echo "Fim do loop em 7"
    Exit Do
    End If
    WScript.Echo num1
Loop

Do Until num1 = 10
    num1 = num1 + 1
    WScript.Echo num1
Loop

For i = 1 To 10
    If (i = 5) Then 
    WScript.Echo "metade"
    End If
    WScript.Echo i
Next

