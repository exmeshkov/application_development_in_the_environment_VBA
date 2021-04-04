Attribute VB_Name = "Module4"
Sub var12()

Dim A() As Integer, i As Integer, j As Integer, n As Integer, M As Integer, st As String, max As Integer, max_index(1) As Integer, tmp As Integer, k As Integer


n = InputBox("Введите кол-во строк в матрице", "Ввод данных")
M = InputBox("Введите кол-во столбцов в матрице", "Ввод данных")
k = InputBox("Введите на пересечении какого столбца и строки должен находится максимальный элемент")

ReDim A(n - 1, M - 1)

ActiveDocument.Range.Text = ActiveDocument.Range.Text & "Начальная матрица: " '

st = ""
For i = 0 To n - 1
    For j = 0 To M - 1
        A(i, j) = 100 * Rnd + 1
        st = st + Str(A(i, j)) + "    "
    Next
    ActiveDocument.Range.Text = ActiveDocument.Range.Text & st '
    st = ""
Next


max = 0

For i = 0 To n - 1
    For j = 0 To M - 1
        If Abs(A(i, j)) > max Then
            max = Abs(A(i, j))
            max_index(0) = i
            max_index(1) = j
        End If
    Next
Next

ActiveDocument.Range.Text = ActiveDocument.Range.Text & "Максимальный эл-т = " & Str(max) '

k = k - 1

For i = 0 To n - 1
    tmp = A(k, i)
    A(k, i) = A(max_index(0), i)
    A(max_index(0), i) = tmp
Next

For j = 0 To n - 1
    tmp = A(j, k)
    A(j, k) = A(j, max_index(1))
    A(j, max_index(1)) = tmp
Next

ActiveDocument.Range.Text = ActiveDocument.Range.Text & "Новая матрица: " '

For i = 0 To n - 1
    For j = 0 To M - 1
        st = st + Str(A(i, j)) + "    "
    Next
    ActiveDocument.Range.Text = ActiveDocument.Range.Text & st '
    st = ""
Next

End Sub

Sub var24()

Dim A() As Double, i As Integer, j As Integer, n As Integer, max As Double, max_index(1) As Integer, M() As Double, st As String

n = InputBox("Введите размерность матрицы", "Ввод данных")

ReDim A(n - 1, n - 1)
ReDim M(n - 2, n - 2)

ActiveDocument.Range.Text = ActiveDocument.Range.Text & "Начальная матрица: " '

st = ""
For i = 0 To n - 1
    For j = 0 To n - 1
        A(i, j) = 100 * Rnd + 1
        st = st + Str(A(i, j)) + "    "
    Next
    ActiveDocument.Range.Text = ActiveDocument.Range.Text & st '
    st = ""
Next

max = 0

For i = 0 To n - 1
    For j = 0 To n - 1
        If Abs(A(i, j)) > max Then
            max = Abs(A(i, j))
            max_index(0) = i
            max_index(1) = j
        End If
    Next
Next
ActiveDocument.Range.Text = ActiveDocument.Range.Text & "Максимальный эл-т = " & Str(max) '


ActiveDocument.Range.Text = ActiveDocument.Range.Text & "Новая матрица: " '
st = ""
For i = 0 To max_index(0) - 1
    For j = 0 To max_index(1) - 1
        If i <> max_index(0) Then
            If j <> max_index(1) Then
                M(i, j) = A(i, j)
                st = st + Str(M(i, j)) + "    "
            End If
        End If
    Next
    ActiveDocument.Range.Text = ActiveDocument.Range.Text & st '
    st = ""
Next

For i = max_index(0) + 1 To n - 1
    For j = max_index(1) + 1 To n - 1
        If i <> max_index(0) Then
            If j <> max_index(1) Then
                M(i - 1, j - 1) = A(i, j)
                st = st + Str(M(i - 1, j - 1)) + "    "
            End If
        End If
    Next
    ActiveDocument.Range.Text = ActiveDocument.Range.Text & st '
    st = ""
Next


End Sub
