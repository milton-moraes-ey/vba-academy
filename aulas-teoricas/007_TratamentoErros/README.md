# Tratamento de Erros

- Tentar prever os erros que o Usuário pode cometer
- Duas formas de tratativas de erro no VBA: `On Error GoTo <RÓTULO_DE_ERRO >:` ou `On Error Resume Next:`
- Cada uma tem uma abordagem diferente

Exemplos:

### On error GoTo

```visual-basic
Private Sub CommandButton1_Click()
On Error GoTo erro:
    
    Dim n As Integer, R As Double
    
    If IsNumeric(TextBox1.Value) = False Then
        MsgBox "Não é um número": TextBox1.SetFocus
        TextBox1.Value = ""
        Exit Sub
    ElseIf TextBox1.Value < 0 Then
        MsgBox "Coloque u m número maor que zero": TextBox1.SetFocus
        TextBox1.Value = ""
        Exit Sub
    End If
    
    n = TextBox1.Value
    R = Sqr(n)
    
    Label1.Caption = R
    
Exit Sub
erro:
    MsgBox "Houve ume erro", vbCritical, "ERRO" 'Tratamento do erro On Error GoTo
End Sub
```

No código acima fazemos um cálculo da raiz quadrada de um número inputado por um usuário via *TextBox* de um *UserForm*. Para tratativa de erros nesse tipo de cenário podemos:

- Fazer verificações via IF, prevendo alguns possíveis atitudes do usuário que resultaria em erro - vide bloco IF - ELSEIF do código acima
- Adicionar uma camada de tratamento de erro extra com a expressão `On Error GoTo` para fazer qualquer tipo de tratativa excepcional de erro.
- Essa forma de fazer o tratamento do erro conduz o código da aplicação para o rótulo (no exemplo **erro:**) e a partir dali faz-se o tratamento do erro conforme a necessidade. No caso apresentado, a tratativa do erro foi simplesmente uma mensagem (**MsgBox**) no final do código ‘

## On Error Resume Next

- Ao cair em uma situação de erro, esse comando simplesmente ignora o erro e segue para os próximos comandos contidos na sequência do código.
- Muito utilizado em Loops.

```visual-basic
Sub calcular()
Dim rng As Range
On Error Resume Next
For Each rng In Selection
    rng.Value = Sqr(rng.Value)
Next
End Sub
```