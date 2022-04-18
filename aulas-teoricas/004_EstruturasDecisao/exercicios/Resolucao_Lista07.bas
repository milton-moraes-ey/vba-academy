Attribute VB_Name = "Resolucao_Lista07"
Option Explicit

Sub exercicio56()
    Dim a As Integer, b As Integer
    a = InputBox("Informe um n�mero")
    b = InputBox("Informe outro n�mero")
    
    If a > b Then
        MsgBox a
    ElseIf a = b Then
        MsgBox "S�o iguais"
    Else
        MsgBox b
    End If
End Sub

Sub exercicio57()
    Dim a As Integer
    a = InputBox("Informe um n�mero")
    If a Mod 2 = 0 Then
        MsgBox a & " � par"
    Else
        MsgBox a & " � �mpar"
    End If
End Sub

Sub exercicio58()
    Dim a As Integer
    a = InputBox("Informe um n�mero")
    If a > 0 Then
        MsgBox "O n�mero informado � POSITIVO"
    ElseIf a < 0 Then
        MsgBox "O n�mero informado � NEGATIVO"
    Else
        MsgBox "ZERO"
    End If
End Sub

Sub exercicio59()
    Dim a As String
    a = InputBox("Informe seu sexo digitando F (feminino) ou M (masculino)")
    
    If a = "M" Or a = "m" Then
        MsgBox "Masculino"
    ElseIf a = "F" Or a = "f" Then
        MsgBox "Feminino"
    Else
        MsgBox "Entre somente com F ou M", vbCritical, "ENTRADA INV�LIDA"
    End If
End Sub

Sub exercicio60()

    Dim letra As String
    letra = InputBox("Informe uma letra")
    
    Select Case letra
        Case Is = "a", "e", "i", "o", "i", "A", "E", "I", "O", "U"
            MsgBox "VOGAL"
        Case Is = "b", "c", "d", "f", "g", "h", "j", "k", "l", "m", "n", "p", "q", "r", "s", "t" _
    , "v", "x", "z", "w", "y"
            MsgBox "CONSOANTE"
        Case Is = 1, 2, 3, 4, 5, 6, 7, 8, 9, 0
            MsgBox "N�MERO"
        Case Else
            MsgBox "OUTRO"
    End Select
End Sub

Sub exercicio61()
    Dim nota1 As Double, nota2 As Double, nota3 As Double
    nota1 = InputBox("Insira uma nota")
    nota2 = InputBox("Insira uma nota")
    nota3 = InputBox("Insira uma nota")
    
    Dim media As Double
    media = (nota1 + nota2 + nota3) / 3
    
    If media >= 7 Then
        MsgBox "APROVADO: Sua m�dia foi igual a " & media
    ElseIf media < 4 Then
        MsgBox "REPROVADO: Sua m�dia foi igual a " & media
    Else
        MsgBox "RECUPERA��O: Sua m�dia foi igual a " & media
    End If
End Sub

Sub exercicio62()

    Planilha2.Activate
    
    Dim prova1 As Integer, prova2 As Integer, prova3 As Integer
    prova1 = Range("C4").Value
    prova2 = Range("C5").Value
    prova3 = Range("C6").Value
    
    Dim totalPresenca As Integer, totalFaltas As Integer
    totalPresenca = Range("F4").Value
    totalFaltas = Range("F5").Value
    
    Dim notaTotal As Integer
    notaTotal = prova1 + prova2 + prova3
    
    Dim media As Double
    media = notaTotal / 3
    
    Dim porcentagemPresenca As Double
    porcentagemPresenca = totalPresenca / (totalFaltas + totalPresenca)

    
    Dim situacao As String
    situacao = "REPROVADO"
    
    If porcentagemPresenca < 0.75 Then
        MsgBox situacao
        Exit Sub
    End If
    
    If media >= 7 Then
        situacao = "APROVADO"
        MsgBox situacao
    ElseIf media < 4 Then
        MsgBox situacao
    Else
        situacao = "RECUPERA��O"
        MsgBox situacao
    End If

End Sub

Sub exercicio63()

    Planilha3.Activate

    Dim gasolinaPorLitro As Double, alcoolPorLitro As Double, desconto As Double, combustivel As String, quantidadeCombustivel As Integer, totalPagar As Currency
    gasolinaPorLitro = Range("D8").Value
    alcoolPorLitro = Range("D9").Value
    combustivel = Range("J8").Value
    quantidadeCombustivel = Range("K8").Value
    desconto = 0
    
    If combustivel = "Gasolina" Then
    
        If quantidadeCombustivel > 10 And quantidadeCombustivel <= 20 Then
            desconto = 0.03
        End If
        
        If quantidadeCombustivel > 20 Then
            desconto = 0.05
        End If
        
        totalPagar = gasolinaPorLitro * quantidadeCombustivel * (1 - desconto)
        Range("L8").Value = totalPagar
        
    Else
    
        If quantidadeCombustivel > 10 And quantidadeCombustivel <= 20 Then
            desconto = 0.04
        End If
        
        If quantidadeCombustivel > 20 Then
            desconto = 0.06
        End If
    
        totalPagar = alcoolPorLitro * quantidadeCombustivel * (1 - desconto)
        Range("L8").Value = totalPagar
        
    End If
    
End Sub
