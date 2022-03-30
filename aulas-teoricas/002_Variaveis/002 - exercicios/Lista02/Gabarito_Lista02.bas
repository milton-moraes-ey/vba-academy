Attribute VB_Name = "Gabarito_Lista02"
Option Explicit

Sub exercicio11()

    Dim a As Variant, b As Variant, helper As Variant
    a = InputBox("Digite um valor: ")
    b = InputBox("Digite outro valor: ")
    
    helper = a
    a = b
    b = helper
    
    MsgBox a
    MsgBox b

End Sub

Sub exercicio12()

    Dim a As Integer, b As Integer, maiorValor As Integer
    a = InputBox("Digite um valor")
    b = InputBox("Digite outro valor")
    
    maiorValor = -(a > b) * a - (a < b) * b
    MsgBox maiorValor
    
    
End Sub

Sub exercicio13()

    Dim peso As Single, altura As Single, IMC As Single
    
    peso = InputBox("Informe seu peso")
    altura = InputBox("Informe sua altura")
    
    IMC = peso / altura ^ 2
    
    MsgBox Round(IMC, 2)

End Sub

Sub exercicio14()

    Dim precoFixo As Integer, precoPorKm As Single, km As Integer, totalDaCorrida As Single
    
    precoFixo = 5
    precoPorKm = 4.5
    km = InputBox("Informa a distancia a percorrer")

    totalDaCorrida = precoFixo + precoPorKm * km
    MsgBox "Total da corrida foi: R$" & totalDaCorrida
    
End Sub

Sub exercicio15()

    Dim ganhoPorHora As Single, horasTrabalhadasPorDia As Single, diasPorSemana As Integer, salario As Double
    ganhoPorHora = InputBox("Informe quanto você ganha por hora")
    horasTrabalhadasPorDia = InputBox("Informe a quantidade de horas que você trabalha por dia")
    diasPorSemana = InputBox("Informe quantos dias você trabalha na semana")
    
    salario = ganhoPorHora * horasTrabalhadasPorDia * diasPorSemana * 4
    MsgBox "Seu salario no mês foi de: R$" & salario

End Sub

Sub exercicio16()
    Dim investimento As Double, taxaDeJurosAoMes As Single, quantidadeDeMeses As Integer, rendimento As Double
    investimento = InputBox("Iforme o valor do investimento em reais (R$)")
    taxaDeJurosAoMes = InputBox("Informe a taxa de juros ao mes (%)") / 100
    quantidadeDeMeses = InputBox("Informe a quantidade de meses")
    
    rendimento = investimento * (1 + taxaDeJurosAoMes) ^ quantidadeDeMeses
    MsgBox "O rendimento do investimento será de R$" & Round(rendimento, 2)
End Sub
Sub exercicio17()
    
    Dim ws As Worksheet, remuneracaoPorHora As Currency, horasTrabalhadasPorDia As Single, diasPorSemana As Integer
    Dim impostoDeRenda As Integer, INSS As Integer, sindicato As Integer, salarioBruto As Currency, salarioLiquido As Currency, descontos As Double
    
    Set ws = Planilha1
    With ws
     remuneracaoPorHora = .Range("D9").Value
     horasTrabalhadasPorDia = .Range("D10").Value
     diasPorSemana = .Range("D11").Value
    End With
    
    salarioBruto = remuneracaoPorHora * horasTrabalhadasPorDia * diasPorSemana * 4
    impostoDeRenda = salarioBruto * 0.11
    INSS = salarioBruto * 0.08
    sindicato = salarioBruto * 0.05
    
    descontos = impostoDeRenda + INSS + sindicato
    salarioLiquido = salarioBruto - descontos
    
    
    With ws
        .Range("G9") = salarioBruto
        .Range("G10") = impostoDeRenda
        .Range("G11") = INSS
        .Range("G12") = sindicato
        .Range("G13") = salarioLiquido
    End With

End Sub

Sub exercicio18()

    Dim number As String
    number = InputBox("Digite um número de três algarismos")
    number = StrReverse(number)
    MsgBox number

End Sub

Sub exercicio19()

    Dim ws As Worksheet, valorDaCompra As Double, taxaDeImportacao As Single, totalDaCompra As Double
    
    Set ws = Planilha2
    With ws
        valorDaCompra = .Range("D9").Value
        taxaDeImportacao = .Range("D10").Value
    End With
    
    If valorDaCompra > 500 Then
        totalDaCompra = valorDaCompra + valorDaCompra * taxaDeImportacao
        ws.Range("D11").Value = Round(totalDaCompra, 2)
        Exit Sub
    End If
    
    totalDaCompra = valorDaCompra
    ws.Range("D11").Value = Round(totalDaCompra, 2)
End Sub

Sub exercicio20()

    Dim precoPorUnidade As Integer, totalDeUnidades As Integer, desconto As Single, totalDaCompra As Double
    precoPorUnidade = 5
    totalDeUnidades = InputBox("Informe o total de unidades compradas")
    desconto = 0.1
    
    If totalDeUnidades > 50 Then
        totalDaCompra = precoPorUnidade * totalDeUnidades
        totalDaCompra = totalDaCompra - (totalDaCompra * desconto)
        MsgBox Round(totalDaCompra)
        Exit Sub
    End If
    
    totalDaCompra = precoPorUnidade * totalDeUnidades
    MsgBox totalDaCompra

End Sub
