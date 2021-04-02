Attribute VB_Name = "Module1"
Option Base 1

Sub var12()

Dim A() As Integer, i As Integer, n As Integer, result As Integer, j As Integer

n = InputBox("Введите кол-во человек в очереди, не меньше 1", "Ввод данных")
1: i = InputBox("Введите номер человека, время пребывания в очереди которого вы хотите узнать, не больше n", "Ввод данных")
result = 0
If i > n Then
    MsgBox "Ошибка. Повторите ввод"
    GoTo 1
End If
For j = 1 To i
result = result + j
Next

MsgBox Str(result)

End Sub
