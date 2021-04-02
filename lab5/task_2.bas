Attribute VB_Name = "Module2"
Sub task_2()

Dim day As Integer, month As Integer, year As Integer

year = InputBox("¬ведите год")
1: month = InputBox("¬ведите мес¤ц")
day = InputBox("¬ведите день")

If day < 1 Or day > 31 Or month < 1 Or month > 12 Then
    MsgBox "Ќевозможна¤ дата, повторите ввод"
    GoTo 1
End If

MsgBox "ƒень: " & Str(day) & vbLf & "ћес¤ц: " & Str(month) & vbLf & "√од: " & Str(year)

End Sub

