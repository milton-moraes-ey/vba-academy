# 003 - Objeto Range

### **Referências às células**

**Objetos do tipo range?**

- Células
- Conjunto de células

```vbnet

Sub exemplos()

'Fazer referencia a uma unica celula
Range("C1").Value = "vba"

' Fazer referência a duas céluas ao mesmo tempo:
Range("b4, b7").Value = "mesma referencia"

' Fazer referencia a um intervalo contínuo
Range("C4:C10").Value = "intervalo continuo"

' Apagando os valores
Range("A1:Z20").Clear

' Selecionando dois intervalos distintos
Range("B4:C9, E5:F9").Value = "sds"

' Fazendo referência a um intervalo nomeado
Planilha4.Range("DIAS").Select

' Fazendo referencia a uma linha inteira
Range("2:4, 7:10").Select

' Mesma coisa para colunas
Range("B:B, E:F, I:H").Select

Planilha4.Range("A4") = "Novo valor pra essa plan"

End Sub

```
---

### **Numero de linhas de um intervalo**

````VB

Sub numeroDeLinhasDeUmIntervalo()

' Como contar o número de linhas de um intervalo
Dim n As Integer
n = Range("A1").CurrentRegion.Rows.Count
MsgBox n

End Sub
````
---

### Método Find (localizar)

```vbnet
		Dim txt As String
    txt = "leite"
    Cells.Find(txt, , , xlWhole).Select
```

Procura a primeira palavra “leite” em uma worksheet. O termo “xlWhole” retorna exatamente a palavra procurada - lembrando que o VBA não é *case sensitive*

```vbnet
		Dim txt As String
    txt = "leite"
    Cells.Find(txt, , , xlPart).Select
```

Com o “xlPart” Procura a primeira palavra “leite”, porém ele pode encontrar palavras compostas, como por exemplo “Café com leite”. 

Um caso prático do método Find, é quando precisamos fazer algum tipo de alteração em determinado elemento em um sistema. Suponhamos que dentro de um sistema de cadastro de funcionários que temos uma tabela com ID, Nome, Sobrenome e Função. Agora, queremos editar a função do funcionário cujo ID é 24569. Como o ID é um valor único, podemos usar o método Find para encontrar esse valor e a linha em que esse funcionário está registrado, e posteriormente prosseguir com a lógica que faça a alteração na coluna Função. Exemplo de como encontrar a Linha desse registro:

```vbnet
Sub localizandoLinhaDeDeterminadoValor()
    Dim ID As String
    Dim tabelaDeFuncionarios As Range
    Dim nLinha As Integer
    
    Set tabelaDeFuncionarios = Range("A1").CurrentRegion
    ID= "24569"
    
    nLinha  = tabela.Columns(4).Find(txt, , , xlWhole).Row
    
    MsgBox nLinha 
End Sub
```

---

