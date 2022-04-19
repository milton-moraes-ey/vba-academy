Attribute VB_Name = "Resolucao_Lista08"
Option Explicit

Sub exercicio64()

    Planilha2.Activate
    
    Dim alfabetizado As String, idade As Integer, situacao As String

    alfabetizado = Range("C4").Value
    idade = Range("C5").Value
    situacao = "OBRIGATÓRIO"
    
    If alfabetizado = "NÃO" And idade < 16 Or alfabetizado = "SIM" And idade < 16 Then
        situacao = "NÃO VOTA"
        Range("C6").Value = situacao
        Exit Sub
    End If
    
    If alfabetizado = "NÃO" And idade >= 16 Then
        situacao = "FACULTATIVO"
        Range("C6").Value = situacao
        Exit Sub
    End If
    
    If alfabetizado = "SIM" And idade >= 16 And idade < 18 Or idade > 70 Then
        situacao = "FACULTATIVO"
        Range("C6").Value = situacao
        Exit Sub
    End If
    
    Range("C6").Value = situacao
    
End Sub

Sub exercicio65()
    Planilha3.Activate
    
    Dim brasil As Integer, argentina As Integer, resultado As String
    
    brasil = Range("E7").Value
    argentina = Range("G7").Value
    
    If brasil > argentina Then
        resultado = "BRASIL"
        Range("D11").Value = resultado
        Exit Sub
    End If
    
    If argentina > brasil Then
        resultado = "ARGENTINA"
        Range("D11").Value = resultado
        Exit Sub
    End If
    
    resultado = "EMPATE"
    Range("D11").Value = resultado
    
   ' If brasil > argentina Then
   '    resultado = "BRASIL"
   ' ElseIf brasil = argentina Then
   '    resultado = "EMPATE"
   ' Else
   '    resultado = "ARGENTINA"
   ' End If
   ' Range("D11").Value = resultado
    
    
End Sub

Sub exercicio66()
'SENHA 123 LOGIN ADMIN

Dim login As String, senha As String
login = UserForm1.TextBox1.Value: senha = UserForm1.TextBox2.Value
login = LCase(login)

If login = "admin" And senha = "123" Then
    MsgBox "Bem vindo"
ElseIf login <> "admin" And senha <> "123" Then
    MsgBox "Usuário e senha incorretos"
ElseIf senha <> "123" Then
    MsgBox "Senha incorreta"
Else
    MsgBox "Usuário incorreto"
End If

End Sub

Sub exercicio67()

    Dim total As Double, totalFinal As Double
    total = UserForm2.TextBox1.Value
    
    If UserForm2.OptionButton1 = True Then
        totalFinal = total * (1 - 0.05)
        totalFinal = VBA.FormatCurrency(totalFinal, 2)
        UserForm2.TextBox2.Value = "R$" & totalFinal
    ElseIf UserForm2.OptionButton2 = True Then
        totalFinal = total * (1 + 0.05)
        totalFinal = VBA.FormatCurrency(totalFinal, 2)
        UserForm2.TextBox2.Value = "R$" & totalFinal
    Else
        MsgBox "Selecione uma opção de pagamento"
        Exit Sub
    End If
    

End Sub

Sub exercicio68()
    Dim mes As String, trimestre As String
    mes = UserForm3.ComboBox1.Value
    
   If mes = "" Then
   MsgBox "SELECIONE UM MES"
   Exit Sub
   Else
     Select Case mes
        Case Is = "janeiro", "fevereiro", "março"
        trimestre = "1° Trimestre"
        UserForm3.Label2.Caption = trimestre
        
        Case Is = "abril", "maio", "junho"
        trimestre = "2° Trimestre"
        UserForm3.Label2.Caption = trimestre
        
        Case Is = "julho", "agosto", "setembro"
        trimestre = "3° Trimestre"
        UserForm3.Label2.Caption = trimestre
        
        Case Is = "outubro", "novembro", "dezembro"
        trimestre = "4° Trimestre"
        UserForm3.Label2.Caption = trimestre
    End Select
   End If
    
End Sub
