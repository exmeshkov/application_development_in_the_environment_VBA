Attribute VB_Name = "Module2"
Sub task_2()

Dim day As Integer, month As Integer, year As Integer

year = InputBox("������� ���")
1: month = InputBox("������� �����")
day = InputBox("������� ����")

If day < 1 Or day > 31 Or month < 1 Or month > 12 Then
    MsgBox "����������� ����, ��������� ����"
    GoTo 1
End If

MsgBox "����: " & Str(day) & vbLf & "�����: " & Str(month) & vbLf & "���: " & Str(year)

End Sub

