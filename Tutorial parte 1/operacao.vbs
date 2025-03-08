Option Explicit
Dim num1, num2
num1 = 10
num2 = 2

Dim sum, subtraction, multiplication, division, remainder, power
sum = num1 + num2
subtraction = num1 - num2
multiplication = num1 * num2
division = num1 / num2
remainder = num1 Mod num2
power = num1 ^ num2

MsgBox("Sum: " & sum & vbCrLf & "Subtraction: " & subtraction)
MsgBox("Multiplication: " & multiplication & vbCrLf & "Division: " & division)
MsgBox("Remainder of division: " & remainder & vbCrLf & "Power: " & power)


MsgBox("One of the nums is less than 10:" & (num1<10 or num2<10) & vbCrLf & "Only one of the nums is equal to 10:" & (num1<>10 xor num2<>10))