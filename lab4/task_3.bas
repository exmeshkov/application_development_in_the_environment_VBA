Attribute VB_Name = "Module6"
Sub var12()

Dim x As Double

x = InputBox("¬ведите значение переменной x", "¬вод числа")

If x < 0 Then
    MsgBox Sin(x)
Else
    MsgBox Cos(x)

End If

End Sub

