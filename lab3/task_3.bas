Attribute VB_Name = "Module2"
Sub var12()

Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double, ax As Double, ay As Double, res As Boolean

'Осуществим ввод значений;
x1 = Val(InputBox("Введите значение координаты x левой верхней вершины прямоугольника", "Ввод числа"))
y1 = Val(InputBox("Введите значение координаты y левой верхней вершины прямоугольника", "Ввод числа"))

x2 = Val(InputBox("Введите значение координаты x правой нижней вершины прямоугольника", "Ввод числа"))
y2 = Val(InputBox("Введите значение координаты y правой нижней вершины прямоугольника", "Ввод числа"))

ax = Val(InputBox("Введите значение координаты x точки А ", "Ввод числа"))
ay = Val(InputBox("Введите значение координаты y точки А ", "Ввод числа"))

'Проверка пренадлежности точки к прямоугольнику;
If x1 <= ax <= x2 And y1 <= ay <= y2 Then
    res = True
Else
    res = False
End If

'Вывод результата;
MsgBox res
    

End Sub

Sub var24()

Dim a As Integer, b As Integer, c As Integer, res As Boolean

'Осуществим ввод значений;
a = Val(InputBox("Введите число a", "Ввод числа"))
b = Val(InputBox("Введите число b", "Ввод числа"))
c = Val(InputBox("Введите число c", "Ввод числа"))

'Выполняем проверку условий;
If a Mod b = 0 And a Mod c <> 0 Then
    res = True
Else
    res = False
End If
    
'Вывод результата
MsgBox res

End Sub
