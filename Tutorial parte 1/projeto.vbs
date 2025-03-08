Option Explicit
Dim word
Dim letters
Dim maxErrors
Dim errors
Dim attempt
Dim displayWord
word = LCase(InputBox("Enter a word for your friend to guess"))
maxErrors = 5

ReDim letters(Len(word) - 1)
ReDim displayWord(Len(word) - 1)

Dim i
For i = 1 To Len(word)
    letters(i - 1) = Mid(word, i, 1)
    displayWord(i - 1) = "*"
Next

WScript.Echo "The word has " & Join(displayWord, "") & " letters."

Do While errors < maxErrors
    WScript.Echo "" & vbCrLf
    WScript.StdOut.WriteLine "Enter a letter:"
    attempt = LCase(WScript.StdIn.ReadLine())
    If Len(attempt) <> 1 Then
        WScript.Echo "Enter only one letter"
    ElseIf InStr(word, attempt) Then
        CheckLetter(attempt)
    Else
        errors = errors + 1
        WScript.Echo "Errors: " & errors
        DrawMan(errors)
    End If

    Dim foundAsterisk
    foundAsterisk = False

    For i = 0 To UBound(displayWord)
        If displayWord(i) = "*" Then
            foundAsterisk = True
            Exit For
        End If
    Next

    If Not foundAsterisk Then
        WScript.Echo "You guessed the word!!!"
        Exit Do
    End If
Loop

Function CheckLetter(attempt)
    Dim i
    For i = 0 To Len(word) - 1
        If letters(i) = attempt Then
            displayWord(i) = attempt
        End If
    Next
    WScript.Echo "Secret Word: " & Join(displayWord, "")
End Function

Function DrawMan(errors)
    Select Case errors
        Case 1
            WScript.Echo " O " & " Secret Word:" & Join(displayWord, "")
        Case 2
            WScript.Echo " O "
            WScript.Echo "/|" & " Secret Word:" & Join(displayWord, "")
        Case 3
            WScript.Echo " O "
            WScript.Echo "/|\" & " Secret Word:" & Join(displayWord, "")
        Case 4
            WScript.Echo " O "
            WScript.Echo "/|\"
            WScript.Echo "/" & " Secret Word:" & Join(displayWord, "")
        Case 5
            WScript.Echo " O "
            WScript.Echo "/|\"
            WScript.Echo "/ \" & " Secret Word:" & Join(displayWord, "")
    End Select
End Function