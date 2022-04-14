Attribute VB_Name = "Resolucao_Lista06"
Option Explicit

Sub exercicio51()
    Planilha8.Activate
    
    Dim tabela As Range
    Set tabela = Range("B2").CurrentRegion
    
    Dim primeiraLinhaVazia As Integer
    primeiraLinhaVazia = tabela.Rows.Count + 2
    
    Dim rangeDeValores As Range
    Set rangeDeValores = Range("F3:H3")
    
    Range(Cells(primeiraLinhaVazia, 2), Cells(primeiraLinhaVazia, 4)) = rangeDeValores.Value
End Sub

Sub exercicio52()
    Planilha9.Activate
    
    Dim tabela As Range
    Set tabela = Range("B2").CurrentRegion
    
    Dim novosValores As Range
    Set novosValores = Range("F3:H3")
    
    Dim nome As String
    nome = novosValores.Cells(1, 1)
        
    Dim linhaNome As String
    linhaNome = tabela.Find(nome, , , xlWhole).Row
    
    Range(Cells(linhaNome, 2), Cells(linhaNome, 4)) = novosValores.Value
End Sub

Sub exercicio53()
    Planilha10.Activate
    
    Dim tabela As Range
    Set tabela = Range("B2").CurrentRegion
    
    Dim nomeDeletar As Range
    Set nomeDeletar = Range("F3")
        
    Dim linhaNome As String
    linhaNome = tabela.Find(nomeDeletar, , , xlWhole).Row
    
    tabela.Rows(linhaNome - 1).Delete xlUp
End Sub

Sub exercicio54()
    Planilha11.Activate

    Dim ultimaLinha As Integer
    ultimaLinha = Planilha11.Range("B2").CurrentRegion.Rows.Count + 1
        
    Dim tabela As Range
    Set tabela = Range(Cells(3, 2), Cells(ultimaLinha, 4))
   
   
   Planilha2.Activate
   ultimaLinha = Planilha2.Range("A1").CurrentRegion.Rows.Count + 1
   tabela.Copy Planilha2.Cells(ultimaLinha, 1)
   
   ''ultimaLinha = Planilha2.Range("A1").CurrentRegion.Rows.Count
   Planilha2.Range("A1").CurrentRegion.Rows(ultimaLinha).Select
End Sub

Sub exercicio55()
    Planilha8.Activate
    Dim tabela As Range
    Set tabela = Range("B2").CurrentRegion
    
    Planilha8.Sort.SortFields.Clear
    
    Planilha8.Sort.SortFields.Add tabela.Columns(3), , xlAscending
    Planilha8.Sort.SortFields.Add tabela.Columns(2), , xlAscending
    Planilha8.Sort.SortFields.Add tabela.Columns(1), , xlAscending
    
    Planilha8.Sort.SetRange tabela
    Planilha8.Sort.Header = xlYes
    Planilha8.Sort.Apply

End Sub
