Attribute VB_Name = "Module2"
Sub task_2()

Dim day As Integer, month As Integer, year As Integer

year = InputBox("Введите год")
1: month = InputBox("Введите месяц")
day = InputBox("Введите день")

If day < 1 Or day > 31 Or month < 1 Or month > 12 Then
    MsgBox "Невозможная дата, повторите ввод"
    GoTo 1
End If

MsgBox "День: " & Str(day) & vbLf & "Месяц: " & Str(month) & vbLf & "Год: " & Str(year)

End Sub

