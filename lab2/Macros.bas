Attribute VB_Name = "NewMacros"
Sub task1()

Dim msg As String, style As Integer, title As String, a As Integer, b As Integer
Dim default

msg = "������� ����� a: "
title = "���� ����� a"
default = 25
a = InputBox(msg, title, default, 1000, 7000)

msg = "������� ����� b: "
title = "���� ����� b"
default = 5
b = InputBox(msg, title, default, 17000, 7000)

MsgBox a + b, vbYesNoCancel + vbExclamation, "����� ����� a � b = "
MsgBox a - b, vbCritical, "�������� ����� a � b = "
MsgBox a / b, vbYesNo + vbInformation, "������� ����� a � b = "
MsgBox a Mod b, vbAbortRetryIgnore + vbQuestion, "������� �� ������� a �� b = "
End Sub

Sub task2()
Dim name As String
name = InputBox("������� ���� ���", , "���������")
MsgBox "������������, " & name

End Sub
