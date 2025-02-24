' Ter variaveis declaradas dentro e fora de algum contexto, Dim e Public 
Option Explicit
Public num1
Public subt 

Function soma()
    Dim num2
    num1 = 10
    num2 = 20
    soma = num1 + num2
End Function
Call MsgBox(soma())

subt = num2 - num1 
MsgBox(subt)
'RESULTADO: Erro de execução do Microsoft VBScript: Variável de objeto não definida: 'num2'
'Porque a variável num2 foi declarada dentro da função e não pode ser acessada fora dela
