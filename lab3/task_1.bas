Attribute VB_Name = "NewMacros"
Sub var12()

Dim x As Single, res As Single
'���������� ���� ��������;
x = Val(InputBox("������� �����", "���� �����"))
'�������� ����������� ��������;
res = (x ^ 2 - 7 * x + 10) / (x ^ 2 - 8 * x + 12)
'������� ��������� �� �����;
MsgBox "��������� = " & Str(res)

End Sub

Sub var24()

Dim x As Single, res As Single
'���������� ���� ��������;
x = Val(InputBox("������� �����", "���� �����"))
'�������� ����������� ��������;
res = x - 10 * Sin(x) + Abs(x ^ 4 - x ^ 5)
'������� ��������� �� �����;
MsgBox "��������� = " & Str(res)
End Sub
