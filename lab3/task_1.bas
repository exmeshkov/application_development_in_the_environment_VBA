Attribute VB_Name = "NewMacros"
Sub var12()

Dim x As Single, res As Single
'Осуществим ввод значений;
x = Val(InputBox("Введите число", "Ввод числа"))
'Вычислим необходимое значение;
res = (x ^ 2 - 7 * x + 10) / (x ^ 2 - 8 * x + 12)
'Выведем результат на экран;
MsgBox "Результат = " & Str(res)

End Sub

Sub var24()

Dim x As Single, res As Single
'Осуществим ввод значений;
x = Val(InputBox("Введите число", "Ввод числа"))
'Вычислим необходимое значение;
res = x - 10 * Sin(x) + Abs(x ^ 4 - x ^ 5)
'Выведем результат на экран;
MsgBox "Результат = " & Str(res)
End Sub
