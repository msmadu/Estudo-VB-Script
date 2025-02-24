Option Explicit
Dim palavra
Dim letras
Dim maxErros
Dim erros
Dim tentativa 
Dim exibirPalavra
palavra = Lcase(InputBox("Digite uma palavra para seu amigo acertar"))
maxErros = 5

ReDim letras(Len(palavra) - 1)
ReDim exibirPalavra(Len(palavra) - 1)

Dim i
For i = 1 To Len(palavra)
    letras(i - 1) = Mid(palavra, i, 1)
    exibirPalavra(i - 1) = "*"  
Next

WScript.Echo "A palavra tem " & Join(exibirPalavra, "") & " letras."

Do While erros < maxErros
    WScript.Echo "" & vbCrLf
    WScript.StdOut.WriteLine "Digite uma letra:"
    tentativa = LCase(WScript.StdIn.ReadLine())
    If Len(tentativa) <> 1 Then 
        WScript.Echo "Digite apenas uma letra"
    ElseIf InStr(palavra, tentativa) Then
        VerificarLetra(tentativa)
    Else
        erros = erros + 1
        WScript.Echo "Erros: " & erros
        DesenharBoneco(erros)
    End If
    

    Dim encontrouAsterisco
    encontrouAsterisco = False

    For i = 0 To UBound(exibirPalavra)
        If exibirPalavra(i) = "*" Then
            encontrouAsterisco = True
            Exit For
        End If
    Next

    If Not encontrouAsterisco Then
        WScript.Echo "Acertou a palavra!!!"
        Exit Do
    End If
Loop

Function VerificarLetra(tentativa)
    Dim i 
    For i = 0 To Len(palavra) - 1
        If letras(i) = tentativa Then
            exibirPalavra(i) = tentativa
        End If
    Next
    WScript.Echo "Palavra Secreta: " & Join(exibirPalavra, "")
End Function

Function DesenharBoneco(erros)
    Select Case erros
        Case 1 
            WScript.Echo  " O "  & "   Palavra Secreta:" & Join(exibirPalavra, "")    
        Case 2
            WScript.Echo " O "
            WScript.Echo "/|"   & "   Palavra Secreta:" &  Join(exibirPalavra, "")
        Case 3
            WScript.Echo " O "
            WScript.Echo "/|\"  & "   Palavra Secreta:" & Join(exibirPalavra, "")  
        Case 4
            WScript.Echo " O "
            WScript.Echo "/|\"
            WScript.Echo "/"    & "   Palavra Secreta:" &  Join(exibirPalavra, "")  
        Case 5
            WScript.Echo " O "
            WScript.Echo "/|\"
            WScript.Echo "/ \"  & "   Palavra Secreta:" & Join(exibirPalavra, "")  
    End Select
End Function