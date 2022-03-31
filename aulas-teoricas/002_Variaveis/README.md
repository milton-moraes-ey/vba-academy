# 002 Variáveis em VBA

### **Declaração de variáveis**
Tabela representativa da quantidade de memória utilizada por tipo de variável e sua respectiva faixa de valores do Visual Basic for Applications.

<div align="center">

Tipo de dados | Bytes Usados | Faixa de Valores 
--- | --- | --- 
Boolean | 2 | Verdadeiro ou Falso 
Integer | 2 | -32.768 a +32.767
Long | 4 | -2.147.483.648 a + 2.147.483.647
Single |	4	| - 3,4E38 a +3,4E38
Double |	8 |	-1,7E308 a +1,7E308
Currency (Moeda) |	8	|-9223372036854,5808 a 9223372036854,5807
Date|	8|	01/01/100 a 31/12/9999
String|	1 por caractere |	aprox. 65500
Object|	4	|Qualquer objeto
Variant|16|	Quaquer tipo de dado

</div>

### **Como obter o maior e menor valor sem utilizar IFs**

````VB

Sub obter_maior_e_menor_valor(a As Integer, b As Integer)

  ' Menor valor
  Dim menorValor As Integer
  menorValor = -(a < b) * a - (a > b) * b
  MsgBox "Menor valor é: " & menorValor

  ' Maior valor
  Dim maiorValor As Integer
  maiorValor = -(a > b) * a - (a < b) * b
  MsgBox "Maior valor é: " & maiorValor

End Sub
````

Em VBA valores *truth* sáo iguais a -1 e valores *falsy* sáo 0:
<div align="center">

| Booleans | Value
---| ---|
True | -1 
False | 0

</div>
---

### **Escopo de Variáveis**

<div align="center">

| Tipos de Variáveis | Escopo
---| ---|
Dim | Procedimento / Módulo
Private | Apenas Módulo
Public | Todos os Módulos
Global | Todos os Módulos
Static | Apenas Procedimento

</div>


*obs: Static retêm o valor da variável do Procedimento após o término da sua execução.*

---

### Arrays - **Relacionar Matriz com Intervalos de células**

```vbnet
Sub macro()
' Dado o vetor
Dim dSemana(1 To 7) As String 'Vetor => Uma lina
dSemana(1) = "Domingo"
dSemana(2) = "Segunda"
dSemana(3) = "Terça"
dSemana(4) = "Quarta"
dSemana(5) = "Quinta"
dSemana(6) = "Sexta"
dSemana(7) = "Sábado"

' Atribuindo valores de um vetor à céluas de uma planilha. Lembrando que vetores são uma linha (Unidimensional)
Range("B3:h3").Value = dSemana

' Atribuindo os valores de um vetor, à células de uma planilha que estão dispostas na Vertical (coluna)
' Para executar essa ação, deve-se utilizar, dentro do VBA a função "transpor" do excel
Range("B5:B11").Value = Application.Transpose(dSemana)
End Sub
```

```vbnet

Sub obtendo_valores_da_planilha_salvando_em_vetor()

    ' Não colocar tipoo de dado e o vetor tem que ser dinâmico
    Dim dSemana()
    dSemana() = Range("B3:H3").Value
    
    ' No caso desse tipo de situação, apesar de termos declarado um "vetor", ao puxar os dados
    ' da planilha para o código, o VBA entende como uma MATRIZ - apesar dos valores estarem somente em uma
    ' única linha
    
    ' Não funciona
    MsgBox dSemana(1)
    
    ' Para pegar o valor, devemos especificar linha e a coluna, pois o VBA entende como MATRIZ.
    MsgBox dSemana(1, 1)
    

End Sub

```
---

### **PROCV com Matrizes no VBA**

```vbnet
Sub Procv()

    Dim matrizTabela()
    Dim valorProcurado()
    Dim valorResultado As Range
    
    matrizTabela() = Range("B2:E5571").Value
    valorProcurado() = Range("H3:H771").Value
    Set valorResultado = Range("I3:I62")
    
    valorResultado = WorksheetFunction.VLookup(valorProcurado, matrizTabela, 2, 0)
    

End Sub
```
---

### **Arrays - LBound & UBound**

**Como verificar o comprimento de um vetor?**

```vbnet
Sub tamanhoDoVetor()

    Dim Arr(5) As String
    Dim arrLen As Integer
    
    ' Mostrar o índice superior do vetor Arr
    MsgBox UBound(Arr) ' retorna 5
    
    ' Mostrar o índice inferior do vetor Arr
    MsgBox LBound(Arr) ' retorna 0
    
    ' Tamanho do vetor se seu índice inferior seja = 0:
    If LBound(Arr) = 0 Then
        arrLen = UBound(Arr) + 1
    End If

    
    ' Tamanho de um vetor se o índice inferior for diferente de 0:
    If LBound(Arr) <> 0 Then
        arrLen = UBound(Arr) - LBound(Arr) + 1
    End If

        
    ' Construindo a mesma lógica em um unico if:
    If LBound(Arr) = 0 Then
        arrLen = UBound(Arr) + 1
    Else
        arrLen = UBound(Arr) - LBound(Arr) + 1
    End If
        MsgBox arrLen

End Sub
```

```vbnet

Sub limitesDaMatriz()

    Dim mtz(5, 4) As String
    
    
    'no caso de matrizes, devemos especificar qual dimensão iremos analisar seus limites
    
    ' Verificar o menor índice da linha (índice 1 = primeira dimensão)
    MsgBox LBound(mtz, 1) ' retorna 0
    
    ' Verificar o menor índice da coluna (índice 2 = segunda dimensão)
    MsgBox LBound(mtz, 2) ' retorna 0

     ' Verificar o maior índice da linha (índice 1 = primeira dimensão)
    MsgBox UBound(mtz, 1) ' retorna 5
    
    ' Verificar o maior índice da coluna (índice 2 = segunda dimensão)
    MsgBox UBound(mtz, 2) ' retorna 4
    
End Sub
```

---
## Listas de exercícios

- [Lista 1](/aulas-teoricas/002_Variaveis/exercicios/Lista01/)
- [Lista 2](/aulas-teoricas/002_Variaveis/exercicios/Lista02/)
- [Lista 3](/aulas-teoricas/002_Variaveis/exercicios/Lista03/)