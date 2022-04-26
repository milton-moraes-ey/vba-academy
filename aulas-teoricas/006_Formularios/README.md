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

```visual-basic
sub addDados()
Me.ListBox1.List(0, 0) = "Janeiro"
Me.ListBox1.List(0, 1) = "Fevereiro"
Me.ListBox1.List(0, 2) = "Março"
Me.ListBox1.List(0, 3) = "Abril"
end sub
```

Exemplo contornando o problema das 10 colunas com a propriedade List:Me.ListBox1.List = Planilha2.Range("A1:L12").Value

```visual-basic
sub addMes()
	Me.ListBox1.List = Planilha2.Range("A1:L12").Value
end sub
```

**Utilizando AddItem e List combinados para preencher dados do listbox:**

Dias da semana com um índice.

```visual-basic
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

```visual-basic
Private Sub CommandButton1_Click()
Dim m() As Variant

m = Range("A1").CurrentRegion.Value
Me.ListBox1.List = m

End Sub
```

Segundo exemplo

A partir de uma matriz criada em código:

```visual-basic
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

```visual-basic
sub numeroDaLinha()
	dim nLinha as integer
	nLinha = me.ListBox1.ListIndex
	msgbox nLinha
end sub
```

Como alterar o valor de um determinado dado da ListBox?

OBS: Isso só é possível se o dado a ser alterado for adicionado à *ListBox* através de uma Matriz ou através do método **AddItem** e **List.**  Da mesma forma, podemos usar o método **RemoveItem** para apagar determinado dado da *ListBox.*

```visual-basic

Sub novoValor()

Dim nLinha As Integer
nLinha = Me.ListBox1.ListIndex

Dim novoValor As String
novoValor = InputBox("Digite um novo valor")

Me.ListBox1.List(nLinha, 0) = novoValor

End Sub
```