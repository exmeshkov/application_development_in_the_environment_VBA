Attribute VB_Name = "Module5"
Private Function max(ByVal x As Double, ByVal y As Double, ByVal z As Double)
If x >= y And x >= z Then
    max = x
ElseIf y >= x And y >= z Then
    max = y
ElseIf z >= y And z >= x Then
    max = z
End If
End Function
Private Function min(ByVal x As Double, ByVal y As Double, ByVal z As Double)
If x <= y And x <= z Then
    min = x
ElseIf y <= x And y <= z Then
    min = y
ElseIf z <= y And z <= x Then
    min = z
End If
End Function

Sub var12()

Dim x As Double, y As Double, z As Double, res As Double

x = InputBox("������� ����� x", "���� �����")
y = InputBox("������� ����� y", "���� �����")
z = InputBox("������� ����� z", "���� �����")

MsgBox (max(x, y, z) ^ 2 - 2 ^ x * min(x, y, z)) / (Sin(2 * x) + max(x, y, z) / min(x, y, z))

End Sub

Sub var4()

Dim a As Double, b As Double, c As Double, d As Double

a = InputBox("������� ����� a", "���� �����")
b = InputBox("������� ����� b", "���� �����")
c = InputBox("������� ����� c", "���� �����")
d = InputBox("������� ����� d", "���� �����")

If a = d Then
    MsgBox "����� a ����� ����� d"
ElseIf b = d Then
    MsgBox "����� b ����� ����� d"
ElseIf c = d Then
    MsgBox "����� � ����� ����� d"
Else
    MsgBox max(d - a, d - b, d - c)
End If

End Sub
