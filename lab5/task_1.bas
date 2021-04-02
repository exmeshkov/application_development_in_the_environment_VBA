Attribute VB_Name = "Module1"
Sub task_1()

Dim a As Integer, b As Integer, c As Integer, d As Integer

a = InputBox("¬ведите a")
b = InputBox("¬ведите b")
1: c = InputBox("¬ведите c")
d = InputBox("¬ведите d")
If c + d = 0 Then
    MsgBox "ƒеление на ноль, повторите ввод"
    GoTo 1
End If

MsgBox (a + b) / (c + d)


End Sub
