Attribute VB_Name = "Module3"
Sub var12()

Dim x As Double, y As Double, res As Boolean

'Ввод значений;
x = Val(InputBox("Введите координату x", "Ввод значения"))
y = Val(InputBox("Введите координату y", "Ввод значения"))
'Проверка принадлежности заштрихованной области;
If (x ^ 2 + y ^ 2 <= 36 And x > 0 And y > 0) Or (x - 6 <= y And x > 0 And y < 0) Then
    res = True
Else
    res = False
End If
'Вывод результата;
MsgBox res

End Sub

Sub var24()

Dim x As Double, y As Double, res As Boolean

'Ввод значений;
x = Val(InputBox("Введите координату x", "Ввод значения"))
y = Val(InputBox("Введите координату y", "Ввод значения"))
'Проверка принадлежности заштрихованной области;
If (y >= 0) And (x <= -2 Or x >= 2 Or y >= 2) And (-5 <= x <= 5 And y <= 6) Then
    res = True
Else
    res = False
End If
'Вывод результата;
MsgBox res

End Sub
