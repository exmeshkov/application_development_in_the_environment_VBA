Attribute VB_Name = "Module1"
Sub var12()

Dim n As Integer

1: n = InputBox("¬ведите число n (кол-во лет)")

If (n < 1) Or (n > 99) Then
    MsgBox "ƒопустимы значени¤ от 1 до 99"
    GoTo 1
End If

If (n Mod 10 = 1) Then
    MsgBox "ћне" & Str(n) & " год"
Else
    If (n Mod 10 >= 2 And n Mod 10 <= 4 And (n < 10 Or n > 20)) Then
        MsgBox "ћне " & Str(n) & " года"
    Else
        MsgBox "ћне " & Str(n) & " лет"
    End If
End If


End Sub
