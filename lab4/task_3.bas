Attribute VB_Name = "Module6"
Sub var12()

Dim x As Double

x = InputBox("������� �������� ���������� x", "���� �����")

If x < 0 Then
    MsgBox Sin(x)
Else
    MsgBox Cos(x)

End If

End Sub

