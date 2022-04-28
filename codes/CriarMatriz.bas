Attribute VB_Name = "CriarMatriz"
Option Explicit
Global P()

Sub populateListBox()
    Dim qntdItensNaMatriz As Integer, i As Integer, preco As Integer
    
    qntdItensNaMatriz = WorksheetFunction.RandBetween(5, 30)
    
    
    ReDim P(1 To qntdItensNaMatriz, 1 To 2)
    
    For i = 1 To qntdItensNaMatriz
        preco = WorksheetFunction.RandBetween(10, 100)
        P(i, 1) = "Item " & i
        P(i, 2) = VBA.FormatCurrency(preco, 2)
    Next i

End Sub
