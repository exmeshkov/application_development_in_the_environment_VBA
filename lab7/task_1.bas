Attribute VB_Name = "Module3"
Sub var12()

Dim s As Integer, count As Integer

s = InputBox("¬ведите стоимость покупки")
count = 0

Do While s > 500
    s = s - 500
    count = count + 1
Loop

Do While s > 100
    s = s - 100
    count = count + 1
Loop

Do While s > 50
    s = s - 50
    count = count + 1
Loop

Do While s > 10
    s = s - 10
    count = count + 1
Loop

Do While s > 5
    s = s - 5
    count = count + 1
Loop

Do While s > 2
    s = s - 2
    count = count + 1
Loop

Do While s > 1
    s = s - 1
    count = count + 1
Loop

MsgBox Str(count)

End Sub
