' Two Sum
' Given an array of integers, return indices of the two numbers such that they add up to a specific target.
' You may assume that each input would have exactly one solution, and you may not use the same element twice.

Option Explicit

dim target
dim listArray
dim getList

WScript.StdOut.WriteLine "Enter the numbers separated by spaces:"
getList = WScript.StdIn.ReadLine()
listArr = Split(getList, " ")

WScript.StdOut.WriteLine "Enter the target number:"
target = WScript.StdIn.ReadLine()

Call twoSum(listArray, target)

Function twoSum(listArray, num)
    dim i
    dim j
    dim greater
    greater = UBound(listArray)
    for i = 0 to greater
        for j = i + 1 to greater
            if CInt(listArray(i)) + CInt(listArray(j)) = CInt(num) Then
                WScript.Echo "[" & i & "," & j & "]"
                Exit Function
            End If
        Next
    Next
    WScript.Echo "No solution found."
End Function