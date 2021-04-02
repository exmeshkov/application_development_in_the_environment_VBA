Attribute VB_Name = "Module1"
Sub var12()

Dim a As Double, b As Double, alpha As Double, S As Double, pi As Double

'Осуществим ввод значений;
a = Val(InputBox("Введите большое основание", "Ввод числа"))
b = Val(InputBox("Введите меньшее основание", "Ввод числа"))
alpha = Val(InputBox("Введите угол при большем основании", "Ввод числа"))
pi = 4 * Atn(1)
alpha = alpha * pi / 180
'Вычислим необходимое значение;
res = 1 / 2 * (a ^ 2 - b ^ 2) * Tan(alpha)
'Выведем результат на экран;
MsgBox "Площадь раввнобедренной трапеции = " & Str(res)

End Sub

Sub var24()

Dim H As Double, R As Double, Vк As Double, Vц As Double, pi As Double
'Осуществим ввод значений;
H = Val(InputBox("Введите высоту", "Ввод числа"))
R = Val(InputBox("Введите радиус основания", "Ввод числа"))
pi = 4 * Atn(1)
'Вычислим объём конуса;
Vк = 1 / 3 * pi * R ^ 2 * H
'Вычислим объём цилиндра;
Vц = pi * R ^ 2 * H
'Выведем результат на экран;
MsgBox "Объем конуса = " & Str(Vк) & "; Объём цилиндра равен" & Str(Vц)

End Sub
