Attribute VB_Name = "Module4"
Sub var12()

Dim a As Double, b As Double, c As Double

a = InputBox("¬ведите число a", "¬вод числа")
b = InputBox("¬ведите число b", "¬вод числа")
c = InputBox("¬ведите число c", "¬вод числа")

If a = -b Or a = -c Or b = -c Then
    MsgBox True
Else
    MsgBox False
End If

End Sub

Sub var1()

Dim a As Double, b As Double, c As Double

a = InputBox("¬ведите число a", "¬вод числа")
b = InputBox("¬ведите число b", "¬вод числа")
c = InputBox("¬ведите число c", "¬вод числа")

If a < 0 Then
    MsgBox a ^ 4
Else
 MsgBox a ^ 2
End If

If b < 0 Then
    MsgBox b ^ 4
Else
 MsgBox b ^ 2
End If

If c < 0 Then
    MsgBox c ^ 4
Else
 MsgBox c ^ 2
End If
 

End Sub
