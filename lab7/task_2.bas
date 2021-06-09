Attribute VB_Name = "Module4"
Sub var12_1()

Dim n As Integer, i As Integer, P As Double

n = InputBox("¬ведите число n")
P = 1

For i = 2 To n
    P = P * (1 - 1 / i ^ 2)
Next

MsgBox Str(P)

End Sub

Sub var12_2()

Dim n As Integer, i As Integer, P As Double

n = InputBox("¬ведите число n")
P = 1
i = 2

Do While (i <= n)
    P = P * (1 - 1 / i ^ 2)
    i = i + 1
Loop

MsgBox P

End Sub

Sub var2_1()

Dim n As Integer, i As Integer, S As Double, tmp As Double, j As Integer

n = InputBox("¬ведите число n")
S = 0

For i = 1 To n
    tmp = 0
    For j = 1 To i
        tmp = tmp + Sin(j)
    Next
    S = S + 1 / tmp
Next

MsgBox Str(S)


End Sub

Sub var2_2()

Dim n As Integer, i As Integer, S As Double, tmp As Double, j As Integer

n = InputBox("¬ведите число n")
S = 0
i = 1

Do While (i <= n)
    tmp = 0
    j = 1
    Do While (j <= i)
        tmp = tmp + Sin(j)
        j = j + 1
    Loop
    S = S + 1 / tmp
    i = i + 1
Loop

MsgBox Str(S)

End Sub
