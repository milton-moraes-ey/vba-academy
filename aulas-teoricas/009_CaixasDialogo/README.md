# Caixas de diálogo e interação com usuário

- InputBox ⇒ Caixa de diálogo que obtêm valores imputados pelo usuários
- MsgBox ⇒ Função que exibe mensagens na tela

Opções de botões, ícones do MsgBox:

| Constante | Valor | O que ela faz |
| --- | --- | --- |
| vbOkOnly | 0 | Exibe apenas o botão OK |
| vbOkCancel | 1 | Exibe os botões OK e Cancelar |
| vbAbortRetryIgnore | 2 | Exibe
  os botões Anular, Repetir, Ignorar |
| vbYesNoCancel | 3 | Exibe os botões Sim, Não, Cancelar |
| vbYesNo | 4 | Exibe
  os botões Sim e Não |
| vbRetryCancel | 5 | Exibe os botões Repetir e Cancelar |
| vbCritical | 16 | Exibe
  o ícone de mensagem Crítica |
| vbQuestion | 32 | Exibe o ícone de interrogação |
| vbExclamation | 48 | Exibe
  o ícone de exclamação |
| vbInformation | 64 | Exibe o ícone de informação |
| vbDefaulttButton1 | 0 | O
  botão padrão é o primeiro |
| vbDefaulttButton2 | 256 | O botão padrão é o segundo |
| vbDefaulttButton3 | 512 | O
  botão padrão é o terceiro |
| vbDefaulttButton4 | 768 | O botão padrão é o quarto |

⇒ No msgbox podemos juntar botões para formar configurações específicas de mensagens para usuário, somando os nomes dos componentes (também pode usar os valores numéricos para fazer a mesma lógica abaixo). Exemplo:

```VB
sub msgbox()
	MsgBox "Texto para o usuário", vbInformation + vbYesNo
end sub
```

⇒ Nota que não podemos usar isso para juntar dois ícones, ou seja dois tipos iguais de componente (ex.: ícone + ícone).

### Interagindo com as MsgBox:

```VB
Sub msgBox()

Dim Resp As VbMsgBoxResult

Resp = MsgBox("Deseja Deletar?", vbQuestion + vbYesNoCancel)

If Resp = vbYes Then
    MsgBox "clicou em sim"
ElseIf Resp = vbCancel Then
    MsgBox "clicou em cancelar"
Else
    MsgBox "clicou em nao"
End If

End Sub
```

---

### Abrindo arquivos com Excel: Método GetOpenFileName

**Exemplos básicos:**

1. Seleciona uma imagem e carrega em um Controle Image de um Formulário:

```VB

Private Sub CommandButton1_Click()

Dim c As Variant, TIPO As String
TIPO = "Arquivos BMP,*.bmp,Arquivos JPEG,*.JPEG"

c = Application.GetOpenFilename(TIPO)
If c = False Then Exit Sub

Image1.Picture = LoadPicture(c)

End Sub
```

1. Seleciona um arquivo especificado na variável TIPO e abre o arquivo no EXCEL

```VB
Sub getOpenFileNameExemplo()

Dim c As Variant
Dim TIPO As String
TIPO = "Arquivos Excel,*.xlsx,Arquivos txt,*.txt,Arquivos PDF,*.PDF "

c = Application.GetOpenFilename(TIPO, , "VBA ACADEMY", , False)

If c = False Then Exit Sub

Workbooks.Open c

End Sub
```

**Exemplos avançados:**

1. Abrindo arquivos que não sou do tipo xlsx (ou outro comátivel com Excel) em seus próprios tipos de arquivo. (abre arquivos .TXT no notepad por exemplo)

```VB
Sub flwHyperlink()

Dim C As Variant

C = Application.GetOpenFilename("Arquivo PDF,*.PDF, Arquivo TXT, *.txt", , , , False)

If C = False Then Exit Sub

ThisWorkbook.FollowHyperlink C

End Sub
```

1. Trabalhando com seleção múltipla de arquivos

```VB
Sub multiplosArquivos()

Dim C As Variant

C = Application.GetOpenFilename(, , , , True)

If IsArray(C) = True Then
    For i = 1 To UBound(C)
        Cells(i, 1).Value = C(i)
    Next
End If

End Sub
```

---

### Obter caminho de uma pasta

```VB
Sub GetFolder()

Dim pasta As String

With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "SELECIONE UMA PASTA"
    .Show
    If .SelectedItems.Count = 0 Then Exit Sub
    pasta = .SelectedItems(1)
    MsgBox pasta
End With

End Sub
```