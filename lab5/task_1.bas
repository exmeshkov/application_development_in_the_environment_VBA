Attribute VB_Name = "Module1"
Sub task_1()

Dim a As Integer, b As Integer, c As Integer, d As Integer

a = InputBox("������� a")
b = InputBox("������� b")
1: c = InputBox("������� c")
d = InputBox("������� d")
If c + d = 0 Then
    MsgBox "������� �� ����, ��������� ����"
    GoTo 1
End If

MsgBox (a + b) / (c + d)


End Sub
