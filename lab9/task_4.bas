Attribute VB_Name = "Module1"
Function is_palindrom(s As String, i As Integer, j As Integer) As Boolean

If (i >= j) Then
    is_palindrom = True
    Exit Function
End If

If (Mid(s, i, 1) = Mid(s, j, 1)) Then
    is_palindrom = is_palindrom(s, i + 1, j - 1)
Else
    is_palindrom = False
    Exit Function
End If


End Function

Sub var3()

Dim s As String, i As Integer, j As Integer

s = InputBox("¬ведите строку")
i = InputBox("¬ведите левую границу провер¤емой подстроки")
j = InputBox("¬ведите правую границу провер¤емой подстроки")

MsgBox Str(is_palindrom(s, i, j))

End Sub
