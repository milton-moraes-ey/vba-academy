Attribute VB_Name = "Resolucao_Lista05"
Option Explicit

Sub exercicio41()
    Dim rng As Range
    Planilha3.Activate
    Set rng = Range("B2").CurrentRegion
    MsgBox rng.Address
End Sub

Sub exercicio42()
    Planilha4.Activate
    
    Dim ultimaLinha As Integer
    ultimaLinha = Range("B2").CurrentRegion.Rows.Count + 1
    
    Dim tabela As Range
    Set tabela = Range(Cells(3, 2), Cells(ultimaLinha, 6))
    
    tabela.Select
End Sub

Sub exercicio43()
    Planilha4.Activate
    
    Dim ultimaLinha As Integer
    ultimaLinha = Range("B2").CurrentRegion.Rows.Count + 1
    
    Dim tabela As Range
    Set tabela = Range(Cells(3, 2), Cells(ultimaLinha, 6))
    
    MsgBox tabela.Address
End Sub

Sub exercicio44()
    Planilha7.Activate

    Dim tabela As Range
    Set tabela = Range("B2").CurrentRegion
    
    tabela.ClearFormats
End Sub

Sub erxercicio45()
    Planilha8.Activate
    
    Dim tabela As Range
    Set tabela = Range("B2").CurrentRegion
    
    Dim ultimaLinha As Integer
    ultimaLinha = tabela.Rows.Count
    
    tabela.Rows(ultimaLinha).Select
End Sub

Sub exercicio46()
    Planilha9.Activate
    
    Dim tabela As Range
    Set tabela = Range("B2").CurrentRegion
    
    tabela.Columns(1).Select
End Sub
Sub exercicio47()
    Planilha10.Activate
    
    Dim ultimaLinha As Integer
    ultimaLinha = Range("B2").CurrentRegion.Rows.Count + 1
    
    Dim tabela As Range
    Set tabela = Range(Cells(3, 2), Cells(ultimaLinha, 2))
    
    tabela.Select
End Sub

Sub exercicio48()
    Planilha11.Activate
    
    Dim tabela As Range
    Set tabela = Range("B2").CurrentRegion
    
    Dim primeiraLinhaVazia As Integer
    primeiraLinhaVazia = tabela.Rows.Count + 1
    
    tabela.Rows(primeiraLinhaVazia).Select
End Sub

Sub exercicio49()
    Planilha12.Activate
    Dim lin As Integer
    lin = Range("linha").Value
    
    Dim tabela As Range
    Set tabela = Range("B2").CurrentRegion
    tabela.Rows(lin).Select
    
    'Dim add As String
    'add = tabela.Rows(lin).Address
    'MsgBox add
End Sub

Sub exercicio50()
    Planilha13.Activate
    
    Dim lin As Integer
    lin = Range("lin").Value
    
    Dim tabela As Range
    Set tabela = Range("B2").CurrentRegion
    
    Dim destino As Range
    Set destino = Range("B25:F25")
    destino = tabela.Rows(lin).Value
End Sub
