'' Two Sum
'' Given an array of integers, return indices of the two numbers such that they add up to a specific target.
'' You may assume that each input would have exactly one solution, and you may not use the same element twice.

Option Explicit

dim target
dim listaArray
dim receberLista

WScript.StdOut.WriteLine "Digite os numeros separados:"
receberLista = WScript.StdIn.ReadLine() 
listaArr = Split(lista, " ")

WScript.StdOut.WriteLine "Digite o numero alvo:"
target = WScript.StdIn.ReadLine()

Call twoSum(listaArray, target)

Function twoSum(listaArray, num)
    dim i 
    dim j
    dim maior
    maior = UBound(listaArray)
    for i = 0 to maior
        for j = i + 1 to maior
            if CInt(listaArray(i)) + CInt(listaArray(j)) = CInt(num) Then
                WScript.Echo "[" & i & "," & j & "]"
                Exit Function
            End If
        Next
    Next
    WScript.Echo "Nenhuma solução encontrada."
End Function