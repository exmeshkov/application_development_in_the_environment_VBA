Attribute VB_Name = "Module3"
Sub var12()

Dim x As Double, y As Double, res As Boolean

'���� ��������;
x = Val(InputBox("������� ���������� x", "���� ��������"))
y = Val(InputBox("������� ���������� y", "���� ��������"))
'�������� �������������� �������������� �������;
If (x ^ 2 + y ^ 2 <= 36 And x > 0 And y > 0) Or (x - 6 <= y And x > 0 And y < 0) Then
    res = True
Else
    res = False
End If
'����� ����������;
MsgBox res

End Sub

Sub var24()

Dim x As Double, y As Double, res As Boolean

'���� ��������;
x = Val(InputBox("������� ���������� x", "���� ��������"))
y = Val(InputBox("������� ���������� y", "���� ��������"))
'�������� �������������� �������������� �������;
If (y >= 0) And (x <= -2 Or x >= 2 Or y >= 2) And (-5 <= x <= 5 And y <= 6) Then
    res = True
Else
    res = False
End If
'����� ����������;
MsgBox res

End Sub
