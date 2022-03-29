# 001 - Introdução ao VBA

## Referências Absolutas vs Referências Relativas

**Macro Absoluta:** os resultados acontecem exatamente como foi determinado na gravação, inclusive, nas mesmas células.

**Macro Relativa:** os comandos são executados da maneira que foi gravada, mas a partir da célula ativa, ou seja, da célula que está selecionada no momento. 

## Noções básicas sobre objetos, métodos,m propriedades e eventos

### **Objetos e coleções**
Um objeto representa um elemento de um aplicativo, como uma planilha, uma célula, um gráfico, um formulário ou um relatório. No código do Visual Basic, você deve identificar um objeto antes de aplicar um dos métodos do objeto ou alterar o valor de uma das suas propriedades.

Uma coleção é um objeto que contém vários objetos, normalmente do mesmo tipo, mas nem sempre. No Microsoft Excel, por exemplo, o objeto Workbooks contém todos os objetos Workbook abertos. No Visual Basic, a coleção Forms contém todos os objetos Form em um aplicativo.

Os itens em uma coleção podem ser identificados por número ou por nome. Por exemplo, o seguinte procedimento identifica o primeiro objeto Workbook aberto.

```VB
Sub CloseFirst() 
 Workbooks(1).Close 
End Sub
```
O procedimento seguinte usa um nome especificado como uma cadeia de caracteres para identificar um objeto Form.

```VB

Sub CloseForm() 
 Forms("MyForm.frm").Close 
End Sub
```

Também é possível manipular uma coleção inteira de objetos se eles compartilharem os mesmos métodos. Por exemplo, o procedimento a seguir fecha todos os formulários abertos.

```VB

Copiar
Sub CloseAll() 
 Forms.Close 
End Sub
```

**Retornar objetos**


Todos os aplicativos possuem uma maneira de retornar os objetos que contêm. No entanto, eles não são todos iguais. Assim, você deve consultar o tópico da Ajuda para o objeto ou coleção que está usando no aplicativo para ver como retornar o objeto.

### **Métodos**
Um método é uma ação que um objeto pode executar. Por exemplo, Add é um método do objeto ComboBox, porque ele adiciona uma nova entrada para uma caixa de combinação.

O procedimento a seguir usa o método Add para adicionar um novo item à uma ComboBox.

```VB
Sub AddEntry(newEntry as String) 
 Combo1.Add newEntry 
End Sub
```
### **Propriedades**
Uma propriedade é um atributo que define uma das características do objeto, como tamanho, cor ou localização na tela, ou um aspecto do comportamento dele, como se o objeto está habilitado ou visível. Para alterar as características de um objeto, deve-se alterar os valores de suas propriedades.

Para definir o valor de uma propriedade, siga a referência a um objeto com um período, o nome da propriedade, um sinal de igual (=) e o novo valor da propriedade. Por exemplo, o procedimento a seguir altera a legenda de um formulário do Visual Basic, definindo a propriedade Caption.

```VB

Sub ChangeName(newTitle) 
 myForm.Caption = newTitle 
End Sub
```
Algumas propriedades não podem ser configuradas. O tópico da Ajuda para cada propriedade indica se você pode configurar a propriedade (leitura/gravação), apenas ler a propriedade (somente leitura) ou apenas gravar na propriedade (somente gravação).

É possível recuperar as informações sobre um objeto retornando o valor de uma de suas propriedades. O procedimento a seguir usa uma caixa de mensagem para exibir o título que aparece na parte superior do formulário ativo no momento.

```VB
Sub GetFormName() 
 formName = Screen.ActiveForm.Caption 
 MsgBox formName 
End Sub
```
### **Eventos**


Um evento é uma ação reconhecida por um objeto, como clicar com o mouse ou pressionar uma tecla e, para a qual é possível escrever um código para responder. Os eventos podem ocorrer como resultado de uma ação do usuário ou de um código do programa, ou podem ser disparados pelo sistema.