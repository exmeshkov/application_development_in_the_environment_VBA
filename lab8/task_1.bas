Attribute VB_Name = "Module1"
Sub var12()

Dim A() As Integer, i As Integer, L As Integer, M As Integer, result As String, k As Integer

n = InputBox("¬ведите количество элементов массива", "ќпределение размера массива")
ReDim A(n)
For i = 0 To n
A(i) = Int(100 * Rnd + 1)
Next

1: L = InputBox("¬ведите число L, которое больше 0 и меньше либо равно M - 1")
M = InputBox("¬ведите число M, которое больше либо равно L + 1")
If (L = 0) Or (M - 1 < L) Then
    MsgBox ("ќшибка, повторите ввод данных")
    GoTo 1
End If

result = ""
For k = 0 To n
If (A(k) Mod M = L) Then
    result = result & Str(A(k))
End If
Next

MsgBox result

End Sub
