Attribute VB_Name = "NewMacros"
Sub task1()

Dim msg As String, style As Integer, title As String, a As Integer, b As Integer
Dim default

msg = "Введите число a: "
title = "Ввод числа a"
default = 25
a = InputBox(msg, title, default, 1000, 7000)

msg = "Введите число b: "
title = "Ввод числа b"
default = 5
b = InputBox(msg, title, default, 17000, 7000)

MsgBox a + b, vbYesNoCancel + vbExclamation, "Сумма чисел a и b = "
MsgBox a - b, vbCritical, "Разность чисел a и b = "
MsgBox a / b, vbYesNo + vbInformation, "Частное чисел a и b = "
MsgBox a Mod b, vbAbortRetryIgnore + vbQuestion, "Остаток от деления a на b = "
End Sub

Sub task2()
Dim name As String
name = InputBox("Введите Ваше имя", , "Александр")
MsgBox "Здравствуйте, " & name

End Sub
