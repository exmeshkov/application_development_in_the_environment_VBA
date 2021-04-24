Attribute VB_Name = "Module1"
Option Explicit
Function next_prime_number(n As Integer) As Integer

Dim i As Integer, flag As Boolean

Do While (True)

n = n + 1

flag = True

For i = 2 To n - 1
If (n Mod i = 0) Then
    flag = False
    Exit For
End If
Next

If (flag = True) Then
Exit Do
End If
Loop

next_prime_number = n

End Function


Sub var12()

Dim n As Integer, result As Integer

n = InputBox("¬ведите простое число")
result = next_prime_number(n)

MsgBox "—ледующее простое число: " & Str(result)

End Sub
