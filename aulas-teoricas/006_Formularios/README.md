# Formulários e Controles de Formulários

# List Box

## Configurações básicas

1. Propriedade **RowSource** ⇒ Adiciona dados na *ListBox*. Manual ou em tempo de execução. Preenche os dados baseado nos valores de um range de células. 
    
    <aside>
    💡 Existem diferentes formas de inserir os dados em uma ListBox. Quando formos usar a propriedade **RowSource** recomenda-se inserir os dados em tempo de execução através do evento *Initialize* do formulário.
    
    </aside>
    
2. Após inserir os dados na *ListBox*, devemos sempre organizar o número de colunas a través da propriedade **ColumnCount.** 
3. Podemos ajustar a largura das colunas através da propriedade  **ColumnWidths.**
4. Por padrão, o controle *ListBox* vem com cabeçalhos com false. Para ajustar os cabeçalhos, devemos: 
    1.  Deixar a propriedade **ColumHeads** como *`true`*.
    2. Alterar o range de onde nossos dados começam. Por exemplo, para uma tabela que ocupa o range  `A1:D300`, cujo `A1:D1` é referente ao cabeçalho dos dados, na propriedade **RowSource** da *ListBox* devemos delimitar os dados para: `A2:D300`. Dessa forma, com a propriedade **`ColumHeads = true`** o VBA já entende que a primeira linha da fonte de dados é o cabeçalho.

---

## Adicionando Dados pelo AddItem e List

O método **AddItem** adiciona **LINHAS** à *ListBox.* É importante entender que a *ListBox* atua como se fosse uma planilha, nesse caso, cada interseção de linha x coluna forma uma “célula”. Dessa forma, podemos adicionar itens através da propriedade **List** que recebe dois argumentos: linha e coluna. **NOTA: NA LIST BOX O INDEX INICIAL É 0.** 

Porém, para conseguirmos adicionar um item com a propriedade **List** antes devemos chamar o método **AddItem** que é o responsável por adicionar Linhas à nossa *ListBox.*

**NOTA IMPORTANTE: Através do método AddItem, o VBA tem uma limitação que, podemos adicionar somente até 10 colunas na ListBox.**

Veja exemplo de:

```VB
sub addDados()
Me.ListBox1.List(0, 0) = "Janeiro"
Me.ListBox1.List(0, 1) = "Fevereiro"
Me.ListBox1.List(0, 2) = "Março"
Me.ListBox1.List(0, 3) = "Abril"
end sub
```

Exemplo contornando o problema das 10 colunas com a propriedade List:Me.ListBox1.List = Planilha2.Range("A1:L12").Value

```VB
sub addMes()
	Me.ListBox1.List = Planilha2.Range("A1:L12").Value
end sub
```

**Utilizando AddItem e List combinados para preencher dados do listbox:**

Dias da semana com um índice.

```VB
sub addDiasSemana()
	Dim i As Integer
	For i = 1 To 7
	    ListBox1.AddItem
	    ListBox1.List(ListBox1.ListCount - 1, 0) = i
	    ListBox1.List(ListBox1.ListCount - 1, 1) = WeekdayName(i)
	Next
end sub
```

**NOTA: O ListCount é adicionado para construir a lógica de inserção de dados em cada linha, pois, o index de referência da ListBox é 0. Dessa forma, o método AddItem irá adicionar uma linha na ListBox, fazendo com que o ListCount seja = 1 para a primeira iteração, subtraindo 1, temos que o primeiro argumento da propriedade List seja igual a 0. Portanto, a cada iteração o ListCount terá um inteiro a mais que o valor de referência da ListBox, logo subtraímos 1 para compensar essa diferença e assim, viabilizar a inserção de dados de forma correta na ListBox.**

---

## Preenchendo uma ListBox - Matriz

Primeiro exemplo:

A partir de uma matriz em um intervalo de céluas

```VB
Private Sub CommandButton1_Click()
Dim m() As Variant

m = Range("A1").CurrentRegion.Value
Me.ListBox1.List = m

End Sub
```

Segundo exemplo

A partir de uma matriz criada em código:

```VB
Private Sub CommandButton2_Click()
Dim d(1 To 7) As String
Dim i As Integer

For i = 1 To 7
    d(i) = WeekdayName(i)
Next

Me.ListBox1.List = d
End Sub
```

Como contar o Número da Linha que cliquei?

OBS: Semelhante ao Target.

OBS: O ListIndex é diferente ao ListCount que conta o número TOTAL de linhas da *ListBox*

```VB
sub numeroDaLinha()
	dim nLinha as integer
	nLinha = me.ListBox1.ListIndex
	msgbox nLinha
end sub
```

Como alterar o valor de um determinado dado da ListBox?

OBS: Isso só é possível se o dado a ser alterado for adicionado à *ListBox* através de uma Matriz ou através do método **AddItem** e **List.**  Da mesma forma, podemos usar o método **RemoveItem** para apagar determinado dado da *ListBox.*

```VB

Sub novoValor()

Dim nLinha As Integer
nLinha = Me.ListBox1.ListIndex

Dim novoValor As String
novoValor = InputBox("Digite um novo valor")

Me.ListBox1.List(nLinha, 0) = novoValor

End Sub
```

---

## Seleção múltipla na ListBox

Nas propriedades da *ListBox* temos uma propriedade chamada **MultiSelect,** que disponibiliza três opções de valores:

1. 0 - **fmMulltiSelectSingle** ⇒ Valor padrão, que nos permite selecionar um único item por vez
2. 1 - **fmMulltiSelectMulti** ⇒ Permite escolher mais de um item da ListBox ao clicar sobre.
3. 2 - **fmMulltiSelectExtended** ⇒ Permite que o usuário clique e arraste o mouse para selecionar vários itens da *ListBox*. Porém, só permite múltiplas seleções ao clicar **se a tecla CTRL estiver pressionada**.

Em tempo de execução, podemos fazer o tratamento de múltiplos itens selecionados através da propriedade **Selected**. Veja um exemplo de código que selecionamos múltiplos itens da lista e guardamos a soma dos itens em uma label utilizando o evento Change da *ListBox.*

```VB
Private Sub ListBox1_Change()
Dim i As Integer, n As Integer
    n = 0
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            n = n + ListBox1.List(i, 1)
        End If
    Next i
    Label1.Caption = VBA.FormatCurrency(n, 2)
End Sub
```

---

## Filtro avançado

**A Macro:**

```VB
Sub filtroAvancado()

Dim baseDeDados As Range, intervaloCriterios As Range, destinoFiltro As Range
Dim novaBase As Range
Set baseDeDados = Range("A1").CurrentRegion
Set intervaloCriterios = Range("G1:J2")
Set destinoFiltro = Range("G4:J4")

Range("I2").Value = UserForm.CB_Criterio1.Value
Range("J2").Value = UserForm.CB_Criterio2.Value

baseDeDados.AdvancedFilter xlFilterCopy, intervaloCriterios, destinoFiltro

Set novaBase = destinoFiltro.CurrentRegion.Offset(1, 0)
UserForm.ListBox.RowSource = novaBase.Address
UserForm.Label.Caption = UserForm.ListBox.ListCount - 1 _
& " Registros encontrados"

End Sub
```

**O Formulário:**

```VB
Private Sub ComboBox1_Change()
Call filtroAvancado
End Sub

Private Sub ComboBox2_Change()
Call filtroAvancado
End Sub

Private Sub UserForm_Initialize()
Me.Height = 500
Me.Width = 600

Dim base As Range
Set base = Range("A1").CurrentRegion

ListBox.RowSource = base.Offset(1, 0).Address
End Sub
```

## Filtro Avançado com Intervalo de Datas

**A Função:**

```VB
Function IntervaloDados() As String

Dim baseDeDados As Range, intervaloDeCriterios As Range, destino As Range

Set baseDeDados = Planilha1.Range("A1").CurrentRegion
Set intervaloDeCriterios = Planilha2.Range("A1:G2")
Set destino = Planilha2.Range("A4:F4")

baseDeDados.AdvancedFilter xlFilterCopy, intervaloDeCriterios, destino

IntervaloDados = destino.CurrentRegion.Offset(1, 0).Address(, , , True)

End Function
```

**O formulário:**

```VB
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

Private Sub UserForm_Initialize()
Me.Height = 419
Me.Width = 604
End Sub
```