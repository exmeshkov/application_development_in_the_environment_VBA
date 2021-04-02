Attribute VB_Name = "NewMacros"
Sub task1()

Dim msg As String, style As Integer, title As String, a As Integer, b As Integer
Dim default

msg = "¬ведите число a: "
title = "¬вод числа a"
default = 25
a = InputBox(msg, title, default, 1000, 7000)

msg = "¬ведите число b: "
title = "¬вод числа b"
default = 5
b = InputBox(msg, title, default, 17000, 7000)

MsgBox a + b, vbYesNoCancel + vbExclamation, "—умма чисел a и b = "
MsgBox a - b, vbCritical, "–азность чисел a и b = "
MsgBox a / b, vbYesNo + vbInformation, "„астное чисел a и b = "
MsgBox a Mod b, vbAbortRetryIgnore + vbQuestion, "ќстаток от делени¤ a на b = "
End Sub

Sub task2()
Dim name As String
name = InputBox("¬ведите ¬аше им¤", , "јлександр")
MsgBox "«дравствуйте, " & name

End Sub
