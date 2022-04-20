Attribute VB_Name = "Resolucao_Lista10"
Option Explicit

Sub exercicio79()

Dim rng As Range, C As Range, i As Integer: i = 1
Set rng = Range("A1:H20")

For Each C In rng
    C.Value = i
    i = i + 1
Next

End Sub

Sub exercicio80()

Dim i As Integer

For i = 1 To 20
    Sheets.Add
Next

End Sub


Sub exercicio81()

Dim w As Worksheet, i As Integer: i = 1

For Each w In Worksheets
    Cells(i, 1) = w.Name
    i = i + 1
Next

End Sub

Sub exercicio82()

Dim i As Integer

For i = 1 To 20
    Sheets(i).Delete
Next

End Sub

Sub exercicio83()


Dim i As Integer, j As Integer, soma As Integer
i = Planilha2.Range("A1").CurrentRegion.Rows.Count

For j = 2 To i
    If Cells(j, 3).Value = "Sudeste" Then
        soma = soma + 1
    End If
Next
    MsgBox soma
End Sub

Sub exercicio84()

Dim i As Integer, j As Integer, soma As Integer
i = Planilha2.Range("A1").CurrentRegion.Rows.Count

For j = 2 To i
    If Cells(j, 3).Value = "Sul" Or Cells(j, 3).Value = "Norte" Then
        soma = soma + 1
    End If
Next
    MsgBox soma
End Sub

Sub exercicio85()
Dim i As Integer, j As Integer, z As Integer: z = 2
i = Planilha2.Range("A1").CurrentRegion.Rows.Count
For j = 2 To i
    If Cells(j, 2) = "Fisioterapia" Then
        Cells(z, 5).Value = Cells(j, 1).Value
        Cells(z, 6).Value = Cells(j, 2).Value
        z = z + 1
    End If
Next

End Sub

Sub exercicio86()
Dim i As Integer, j As Integer

i = Planilha4.Range("A1").CurrentRegion.Rows.Count

For j = 2 To i
    If Cells(j, 3) = "Sudeste" Then
        Cells(j, 3).Value = "Sul"
    End If
Next

End Sub

Sub exercicio87()
Dim i As Integer, j As Integer, k As Integer
i = Planilha6.Range("A1").CurrentRegion.Rows.Count
k = 2
For j = 2 To i
    If Cells(k, 3).Value <> "Norte" Then
        Range(Cells(k, 1), Cells(k, 3)).Delete
    Else
        k = k + 1
    End If
Next
End Sub

Sub exercicio88()
Dim i As Integer, j As Integer, k As Integer
i = Planilha7.Range("A1").CurrentRegion.Rows.Count
k = 2
For j = 2 To i
    If Cells(k, 3).Value = "Sudeste" Then
        Range(Cells(k, 1), Cells(k, 3)).Delete
    Else
        k = k + 1
    End If
Next
End Sub
