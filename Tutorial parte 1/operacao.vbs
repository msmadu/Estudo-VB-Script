Option Explicit
Dim num1, num2
num1 = 10
num2 = 2

Dim soma, subtracao, multiplicacao, divisao, resto, potencia
soma = num1 + num2
subtracao = num1 - num2
multiplicacao = num1 * num2
divisao = num1 / num2
resto = num1 Mod num2
potencia = num1 ^ num2

MsgBox("Soma: " & soma & vbCrLf & "Subtração: " & subtracao)
MsgBox("Multiplicação: " & multiplicacao & vbCrLf & "Divisão: " & divisao)
MsgBox("Resto da divisão: " & resto & vbCrLf & "Potência: " & potencia)


MsgBox("Um dos num é menor que 10:" & (num1<10 or num2<10) & vbCrLf & "Apenas um dos num é igual a 10:" & (num1<>10 xor num2<>10))