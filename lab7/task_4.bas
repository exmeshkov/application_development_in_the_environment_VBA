Attribute VB_Name = "Module6"
Sub var12()

Dim i As Integer, n As Integer, flag As Boolean, s As Long, j As Integer


n = InputBox("¬ведите число n")

For i = 2 To n
    s = 0
    For j = 1 To i - 1
    If (i Mod j = 0) Then
        s = s + j
    End If
    Next
    If (i = s) Then
        ActiveDocument.Range.Text = ActiveDocument.Range.Text & Str(i)
    End If
Next

End Sub


Sub var24()

Dim n As Integer, count As Integer, i As Integer

n = InputBox("¬ведите число n")
count = 0
For i = 1 To n
    If ((i Mod 2 <> 0) And (i Mod 3 <> 0) And (i Mod 5 <> 0)) Then
        count = count + 1
    End If
Next

MsgBox Str(count)

End Sub
