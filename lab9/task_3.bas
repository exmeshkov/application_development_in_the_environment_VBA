Attribute VB_Name = "Module1"
Sub sort(A() As Integer, n As Integer)
Dim i As Integer, j As Integer, Z As Integer
    For i = 0 To n - 1
        For j = i + 1 To n - 1
            If A(i) < A(j) Then
            Z = A(j)
            A(j) = A(i)
            A(i) = Z
            End If
        Next
    Next
End Sub
Sub var12()


Dim A() As Integer, i As Integer, sort_str As String, s As String, n As Integer, sort_str2 As String, j As Integer

n = InputBox("Введите длинну массива")

ReDim A(n) As Integer
s = ""
For i = 0 To n - 1
    A(i) = Int(100 * Rnd + 1)
    s = s & Hex(A(i)) & " "
Next

sort A:=A, n:=n

For i = 0 To n - 1
    sort_str = sort_str & Hex(A(i)) & " "
    sort_str2 = sort_str2 & Str(A(i)) & " "
Next

ActiveDocument.Range.Text = ActiveDocument.Range.Text & "Начальный массив:" & vbLf & s & vbLf & "Отсортированный массив по убыванию:" & vbLf & sort_str & vbLf & "Отсортированный массив в десятичном виде:" & vbLf & sort_str2
End Sub
