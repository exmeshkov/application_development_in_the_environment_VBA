Attribute VB_Name = "Module3"
Sub var12()

Dim A() As Integer, n As Integer, i As Integer, j As Integer, st As String

n = InputBox("¬ведите размерность матрицы", "¬вод данных")

ReDim A(n - 1, n - 1)




For i = 0 To n - 1
    For j = i To n - 1
        A(j, n - i - 1) = i + 1
    Next
Next

For i = 0 To n - 1
    For j = 0 To n - 1
        st = st + Str(A(i, j)) + "    "
    Next
    ActiveDocument.Range.Text = ActiveDocument.Range.Text & st '
    st = ""
Next

End Sub


Sub var12_2()

Dim A() As Integer, n As Integer, i As Integer, j As Integer, st As String, k As Integer

n = InputBox("¬ведите размерность матрицы", "¬вод даннных")

ReDim A(n - 1, n - 1)
k = 0
For i = 0 To n - 1
    For j = 0 To n - 1
        If i = j Then
            A(i, j) = k
            k = k + 1
        End If
    Next
Next

For i = 0 To n - 1
    For j = 0 To n - 1
        st = st + Str(A(i, j)) + "    "
    Next
    ActiveDocument.Range.Text = ActiveDocument.Range.Text & st '
    st = ""
Next

End Sub
