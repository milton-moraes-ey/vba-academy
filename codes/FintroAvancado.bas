' Exemplo da função
Function IntervaloDados() As String
Dim baseDeDados As Range, intervaloDeCriterios As Range, destino As Range

Set baseDeDados = Planilha1.Range("A1").CurrentRegion
Set intervaloDeCriterios = Planilha2.Range("A1:G2")
Set destino = Planilha2.Range("A4:F4")

baseDeDados.AdvancedFilter xlFilterCopy, intervaloDeCriterios, destino

IntervaloDados = destino.CurrentRegion.Offset(1, 0).Address(, , , True)
End Function

' ---------------------------- * ----------------------------
' ---------------------------- * ----------------------------
' Exemplo do Formulário:
Private Sub btnFiltrar_Click()

Dim dataInicial As Date, dataFinal As Date
dataInicial = txtDataInicial.Value
dataFinal = txtDataFinal.Value

' Limpar os dados do filtro anterior
Planilha2.Range("A2:G2").Clear

' Associar os campos do range = intervalo de critérios com os txt's do Formulário:
With Planilha2
    .Range("A2").Value = txtNome.Value
    .Range("C2").Value = txtNacionalidade.Value
    .Range("D2").Value = txtTime.Value
    .Range("F2").Value = ">=" & VBA.Format(dataInicial, "mm/dd/yyyy")
    .Range("G2").Value = "<=" & VBA.Format(dataFinal, "mm/dd/yyyy")
End With

' Referenciar a ListBox através do RowSource o Intervalo de Dados retornado pela Função
ListBox.RowSource = IntervaloDados

Label6.Caption = "Total de registros " & ListBox.ListCount - 1
    
End Sub