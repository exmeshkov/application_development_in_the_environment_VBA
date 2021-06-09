Attribute VB_Name = "Module2"
Sub var24()

Dim n As String, i As Integer, res As Integer, tmp As Integer

s = InputBox("¬ведите число, не превышающее 99999")

res = 1

For i = 1 To Len(s)
    res = res * Int(Mid(s, i, 1))
Next

MsgBox Str(res)

End Sub
