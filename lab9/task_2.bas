Function max_of_string(s As String)

Dim m As Integer, i As Integer

m = 0
For i = 0 To Len(s)
    If (Val(Mid(s, i + 1, 1)) > m) Then
        m = Val(Mid(s, i + 1, 1))
    End If
Next

max_of_string = m

End Function

Function generate_increasing_sequence(n As Integer, prefix As String) As Integer
Dim i As Integer

If (n = 0) Then
    ActiveDocument.Range.Text = ActiveDocument.Range.Text & prefix '
    Exit Function
End If


For i = 1 To 9
    If ((Len(prefix) > 0 And i > max_of_string(prefix)) Or (Len(prefix) = 0)) Then
        prefix = prefix & Str(i)
        generate_increasing_sequence = generate_increasing_sequence(n - 1, prefix)
        prefix = Replace(prefix, Str(i), "")
    End If
Next

End Function
Sub var12()

Dim n As Integer

n = InputBox("Введите число n")

n = generate_increasing_sequence(n, "")

End Sub
Sub var2()

Dim n As Integer

1: n = InputBox("Введите число n")
If (n Mod 10 = 0) Then
    MsgBox "Обратное число не может начинаться с 0. Повторите ввод"
    GoTo 1
End If


MsgBox StrReverse(Str(n))

End Sub
