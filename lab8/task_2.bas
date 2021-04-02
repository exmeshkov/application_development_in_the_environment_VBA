Attribute VB_Name = "Module1"
Option Base 1

Sub var12()

Dim A() As Integer, i As Integer, n As Integer, result As Integer, j As Integer

n = InputBox("¬ведите кол-во человек в очереди, не меньше 1", "¬вод данных")
1: i = InputBox("¬ведите номер человека, врем¤ пребывани¤ в очереди которого вы хотите узнать, не больше n", "¬вод данных")
result = 0
If i > n Then
    MsgBox "ќшибка. ѕовторите ввод"
    GoTo 1
End If
For j = 1 To i
result = result + j
Next

MsgBox Str(result)

End Sub
