Attribute VB_Name = "Module2"
Sub var12()

Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double, ax As Double, ay As Double, res As Boolean

'���������� ���� ��������;
x1 = Val(InputBox("������� �������� ���������� x ����� ������� ������� ��������������", "���� �����"))
y1 = Val(InputBox("������� �������� ���������� y ����� ������� ������� ��������������", "���� �����"))

x2 = Val(InputBox("������� �������� ���������� x ������ ������ ������� ��������������", "���� �����"))
y2 = Val(InputBox("������� �������� ���������� y ������ ������ ������� ��������������", "���� �����"))

ax = Val(InputBox("������� �������� ���������� x ����� � ", "���� �����"))
ay = Val(InputBox("������� �������� ���������� y ����� � ", "���� �����"))

'�������� �������������� ����� � ��������������;
If x1 <= ax <= x2 And y1 <= ay <= y2 Then
    res = True
Else
    res = False
End If

'����� ����������;
MsgBox res
    

End Sub

Sub var24()

Dim a As Integer, b As Integer, c As Integer, res As Boolean

'���������� ���� ��������;
a = Val(InputBox("������� ����� a", "���� �����"))
b = Val(InputBox("������� ����� b", "���� �����"))
c = Val(InputBox("������� ����� c", "���� �����"))

'��������� �������� �������;
If b = 0 Or c = 0 Then
    MsgBox "������� ������������ �������� (��������� ������� �� 0), ����������, ���������� ��� ���"
Else
    '��������� �������� �������;
    If a Mod b = 0 And a Mod c <> 0 Then
        res = True
    Else
        res = False
    End If
    '����� ����������
    MsgBox res
End If

End Sub
