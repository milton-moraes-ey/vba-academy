# 002 Variáveis em VBA

### **Declaração de variáveis**
Tabela representativa da quantidade de memória utilizada por tipo de variável e sua respectiva faixa de valores do Visual Basic for Applications.

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

| Booleans | Value
---| ---
True | -1 
False | 0







---
## Listas de exercícios

- [Lista 1](/aulas-teoricas/002_Variaveis/exercicios/Lista01/)
- [Lista 2](/aulas-teoricas/002_Variaveis/exercicios/Lista02/)
- [Lista 3](/aulas-teoricas/002_Variaveis/exercicios/Lista03/)