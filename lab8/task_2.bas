Attribute VB_Name = "Module1"
Option Base 1

Sub var12()

Dim A() As Integer, i As Integer, n As Integer, result As Integer, j As Integer

n = InputBox("������� ���-�� ������� � �������, �� ������ 1", "���� ������")
1: i = InputBox("������� ����� ��������, ����� ���������� � ������� �������� �� ������ ������, �� ������ n", "���� ������")
result = 0
If i > n Then
    MsgBox "������. ��������� ����"
    GoTo 1
End If
For j = 1 To i
result = result + j
Next

MsgBox Str(result)

End Sub
