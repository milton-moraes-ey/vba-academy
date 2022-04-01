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
