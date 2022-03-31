Attribute VB_Name = "Gabarito_Lista01"
Option Explicit

Sub exercicio01()

    Dim a As Integer, b As Integer, soma As Integer
    
    a = InputBox("Digite um número: ")
    b = InputBox("Digite outro número: ")
    soma = a + b
    
    MsgBox a & " + " & b & " = " & soma

End Sub

Sub exercicio02()

    Dim a As Integer, b As Integer
    
    a = InputBox("Digite um número: ")
    b = InputBox("Digite outro número: ")
    
    MsgBox "MÉDIA: " & (a + b) / 2

End Sub

Sub exercicio03()

    Dim nome1 As String, nome2 As String
    nome1 = InputBox("Digite um nome:")
    nome2 = InputBox("Digite outro nome:")
    
    MsgBox nome1 & " - " & nome2

End Sub


Sub exercicio04()

    Dim anoNascimento As Integer, idade As Integer, anoAtual As Integer
    
    anoNascimento = InputBox("Digite seu ano de nascimento:")
    anoAtual = Year(Date)
    idade = anoAtual - anoNascimento
    MsgBox "Sua idade é: " & idade

End Sub

Sub exercicio05()

    Dim tipoBoolean As Boolean
    Dim num As Integer
    num = InputBox("Digite 0 ou -1")
    
    tipoBoolean = num
    MsgBox tipoBoolean

End Sub

Sub exercicio06()

    Dim real As Double, dolar As Double
    real = InputBox("Digite uma quantidade em reais (R$)")
    
    dolar = real / 3.23
    MsgBox "R$" & real & " convertido em dólar é $" & Round(dolar, 2)
    
End Sub

Sub exercicio07()
    
    Dim distancia As Integer, precoGasolina As Double, valorGastoComGasolina As Double

    distancia = InputBox("Digite a distância em (km)")
    precoGasolina = 3.1
    
    valorGastoComGasolina = (distancia / 10) * precoGasolina
    MsgBox "Para a distancia mencionada " & " (" & distancia & "km) " & " serão gastos R$" & valorGastoComGasolina & " com gasolina."
End Sub

Sub exercicio08()

    Dim horaDeEntrada As Date, horaDeSaida As Date, horasTrabalhadas As Date
    horaDeEntrada = InputBox("Digite a hora de entrada")
    horaDeSaida = InputBox("Digite a hora de saída")

    horasTrabalhadas = horaDeSaida - horaDeEntrada
    MsgBox "O total de horas trabalhadas foi de: " & horasTrabalhadas & " horas."

End Sub

Sub exercicio09()

    Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer
    a = InputBox("Digite um número")
    b = InputBox("Digite um número")
    c = InputBox("Digite um número")
    d = InputBox("Digite um número")
    e = InputBox("Digite um número")
    MsgBox a * b * c * d * e
    
End Sub


Sub exercicio10()
    
    Dim nome As String * 3
    nome = InputBox("Digite seu nome")
    MsgBox nome

End Sub
