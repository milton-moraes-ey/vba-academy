Attribute VB_Name = "Resolucao_Lista04"
Option Explicit

Sub exercicio31()
    
    Planilha2.Activate
    Range("B2:F13").Select

End Sub

Sub exercicio32()

    Planilha3.Activate
    Range("dados").Select

End Sub


Sub exercicio33()
    Planilha4.Activate
    
    Dim tabela As Range
    Set tabela = Range("B2").CurrentRegion
    tabela.Select

End Sub

Sub exercicio34()

    Planilha6.Activate
    Range("B3:D10, F3:H10, B13:D20, F13:H20").Select
End Sub

Sub exercicio35()
    Planilha6.Activate
    Range("A:A").Select
End Sub

Sub exercicio36()
    Planilha6.Activate
    Range("A:D").Select
End Sub

Sub exercicio37()
    Planilha6.Activate
    Range("A:D, G:J").Select
End Sub

Sub exercicio38()
    Planilha6.Activate
    Range("1:3").Select
End Sub

Sub exercicio39()
    Planilha6.Activate
    Range("1:3,8:10").Select
End Sub

Sub exercicio40()
    Dim numeroDeLinhas As Integer
     numeroDeLinhas = Range("b2").CurrentRegion.Rows.Count - 1
     MsgBox numeroDeLinhas
End Sub
