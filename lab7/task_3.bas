Attribute VB_Name = "Module5"
Sub var12()

Dim i As Integer, a As Integer, b As Integer, h As Integer

a = InputBox("¬ведите значение а")
b = InputBox("¬ведите значение b")
h = InputBox("¬ведите значение h")

i = a

Do While (i <= b)
    ActiveDocument.Range.Text = ActiveDocument.Range.Text & Str(i) & "           |           " & Str(-Cos(2 * i))
    i = i + h
Loop


End Sub

Sub var24()

Dim i As Integer, a As Integer, b As Integer, h As Integer

a = InputBox("¬ведите значение а")
b = InputBox("¬ведите значение b")
h = InputBox("¬ведите значение h")

i = a

Do While (i <= b)
    ActiveDocument.Range.Text = ActiveDocument.Range.Text & Str(i) & "           |           " & Str(i / Cos(i))
    i = i + h
Loop


End Sub
