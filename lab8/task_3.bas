Attribute VB_Name = "Module2"
Sub var12()

Dim A() As Integer, max_dist(3) As Integer, min_dist(3) As Integer, i As Integer, j As Integer, n As Integer, ar As String


1: n = InputBox("Введите размер массива, кратный двум", "Ввод данных")
If n Mod 2 <> 0 Then
    MsgBox ("Ошбика. Повторите ввод")
    GoTo 1
End If
ReDim A(n)
For i = 0 To n - 1
A(i) = (100 * Rnd + 1)
ar = ar & Str(A(i))
Next
min_dist(0) = 0
min_dist(1) = 1
min_dist(2) = 2
min_dist(3) = 3

max_dist(0) = 0
max_dist(1) = 1
max_dist(2) = 2
max_dist(3) = 3


For i = 0 To n - 1 Step 2
    For j = 0 To n - 1 Step 2
        If i <> j Then
            If ((A(j) - A(i)) ^ 2 + (A(j + 1) - A(i + 1)) ^ 2) ^ 0.5 < ((A(min_dist(2)) - A(min_dist(0))) ^ 2 + (A(min_dist(3)) - A(min_dist(1))) ^ 2) ^ 0.5 Then
                min_dist(0) = i
                min_dist(1) = i + 1
                min_dist(2) = j
                min_dist(3) = j + 1
            ElseIf ((A(j) - A(i)) ^ 2 + (A(j + 1) - A(i + 1)) ^ 2) ^ 0.5 > ((A(max_dist(2)) - A(max_dist(0))) ^ 2 + (A(max_dist(3)) - A(max_dist(1))) ^ 2) ^ 0.5 Then
                max_dist(0) = i
                max_dist(1) = i + 1
                max_dist(2) = j
                max_dist(3) = j + 1
            End If
        End If
    Next
Next
MsgBox "Массив: " & ar & vbLf & "Максимальное расстоение между точками с индексами :" & Str(min_dist(0)) & "," & Str(min_dist(1)) & " и " & Str(min_dist(2)) & "," & Str(min_dist(3)) & vbLf & "Минимальное расстоение между точками с индексами :" & Str(max_dist(0)) & "," & Str(max_dist(1)) & " и " & Str(max_dist(2)) & "," & Str(max_dist(3))

End Sub


Sub var2()

Dim A() As Integer, r As Double, i As Integer, ar As String

1: n = InputBox("Введите размер массива, кратный двум", "Ввод данных")
If n Mod 2 <> 0 Then
    MsgBox ("Ошбика. Повторите ввод")
    GoTo 1
End If

ReDim A(n)
For i = 0 To n - 1
A(i) = (100 * Rnd + 1)
ar = ar & Str(A(i))
Next

r = 0
For i = 0 To n - 1 Step 2
If (A(i) ^ 2 + A(i + 1) ^ 2) ^ 0.5 > r Then
    r = (A(i) ^ 2 + A(i + 1) ^ 2) ^ 0.5
End If
Next

MsgBox "Массив: " & ar & vbLf & "Минимальный радиус окружности, вмещающий все точки = " & Str(r)
End Sub
