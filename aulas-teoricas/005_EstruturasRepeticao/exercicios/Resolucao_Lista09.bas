Attribute VB_Name = "Resolucao_Lista09"
Option Explicit

Sub exercicio69()

Planilha2.Activate
Dim i As Integer

For i = 1 To 100
    Cells(i, 1).Value = i
Next

End Sub

Sub exercicio70()

Planilha2.Activate
Dim i As Integer, j As Integer
i = 2: j = 1
Do While i <= 500
    Cells(j, 2) = i
    j = j + 1
    i = i + 2
Loop

End Sub

Sub exercicio71()

Dim i As Integer, j As Integer
j = 1

For i = 500 To 1 Step -1
    Cells(j, 3) = i
    j = j + 1
Next

End Sub

Sub exercicio72()
Dim i As Integer, j As Integer

i = 500: j = 1
Do Until i = 2
    Cells(j, 4) = i
    i = i - 2
    j = j + 1
Loop

End Sub

Sub exercicio73()

Dim qntdNumeros As Integer, i As Integer, j As Integer, soma As Integer
qntdNumeros = 10: j = 2

For i = 1 To qntdNumeros
    soma = soma + j
    Cells(i, 5) = j
    j = j + 2
Next

Cells(i, 5) = soma

End Sub

Sub exercicio74()
Dim qntdNumeros As Integer, i As Integer, j As Integer, soma As Integer
qntdNumeros = 10: j = 1

For i = 1 To qntdNumeros
    If j Mod 2 <> 0 Then
        soma = soma + j
        Cells(i, 6) = j
        j = j + 1
    End If
    j = j + 1
Next

Cells(i, 6) = soma

End Sub

Sub exercicio75()

Dim i As Integer, j As Integer, soma As Integer
 i = 1: j = 1
 
Do While soma < 200
    soma = soma + i
    Cells(i, 7) = j
    i = i + 1
    j = j + 1
Loop

End Sub
Sub exercicio76()
Dim i As Integer, j As Integer, soma As Integer
i = 1: j = 1

Do While soma <= 200
    soma = soma + j
    If soma < 200 Then
        Cells(i, 8) = j
        i = i + 1
        j = j + 3
    End If
Loop

End Sub

Sub exercicio77()

Dim i As Integer

For i = 1 To 7
    Cells(i, 10) = WeekdayName(i, False)
Next

End Sub

Sub exercicio78()
Dim i As Integer

For i = 1 To 12
    Cells(i, 11) = MonthName(i, False)
Next

End Sub
