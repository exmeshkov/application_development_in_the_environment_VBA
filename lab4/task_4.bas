Attribute VB_Name = "Module7"
Sub var4()

Dim x As Double, y As Double, z As Double

x = InputBox("¬ведите число x", "¬вод числа")
y = InputBox("¬ведите число y", "¬вод числа")
z = InputBox("¬ведите число z", "¬вод числа")

If x + y + z < 1 Then
    If x < y And x < z Then
        x = (y + z) / 2
    ElseIf y < x And y < z Then
        y = (x + z) / 2
    Else
        z = (x + y) / 2
    End If
Else
    If x < y Then
        x = (y + z) / 2
    Else
        y = (x + z) / 2
    End If
End If

MsgBox "x=" & Str(x) & " y=" & Str(y) & " z=" & Str(z)

End Sub
