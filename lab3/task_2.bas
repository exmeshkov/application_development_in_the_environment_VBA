Attribute VB_Name = "Module1"
Sub var12()

Dim a As Double, b As Double, alpha As Double, S As Double, pi As Double

'���������� ���� ��������;
a = Val(InputBox("������� ������� ���������", "���� �����"))
b = Val(InputBox("������� ������� ���������", "���� �����"))
alpha = Val(InputBox("������� ���� ��� ������� ���������", "���� �����"))
pi = 4 * Atn(1)
alpha = alpha * pi / 180
'�������� ����������� ��������;
res = 1 / 2 * (a ^ 2 - b ^ 2) * Tan(alpha)
'������� ��������� �� �����;
MsgBox "������� ��������������� �������� = " & Str(res)

End Sub

Sub var24()

Dim H As Double, R As Double, V� As Double, V� As Double, pi As Double
'���������� ���� ��������;
H = Val(InputBox("������� ������", "���� �����"))
R = Val(InputBox("������� ������ ���������", "���� �����"))
pi = 4 * Atn(1)
'�������� ����� ������;
V� = 1 / 3 * pi * R ^ 2 * H
'�������� ����� ��������;
V� = pi * R ^ 2 * H
'������� ��������� �� �����;
MsgBox "����� ������ = " & Str(V�) & "; ����� �������� �����" & Str(V�)

End Sub
