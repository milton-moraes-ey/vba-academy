# Eventos

- Os eventos são utilizados para execução de macros automaticamente, quando determinada ação for realizada. Por exemplo, o evento **Open** do *WorkBook* executa uma macro toda vez que o Excel abrir.
- Eventos WindowActivate e WindowDeactivate ⇒ Só funcionam quando trocamos entre  janelas do EXCEL

### Eventos mais importantes de uma planilha:

- Change: Executa uma macro quando uma determinada célula for alterada.
- SelectionChange: Executa uma macro quando selecionar determinado range

Funcionalidade importante: Interseção entre dois intervalos. Suponhamos que queremos executar uma macro que execute quando determinada célula for alterada dentro de um conjunto de células específicos. 

```VB
Private Sub Worksheet_Change(ByVal Target As Range)

If Not Intersect(Target, Range("C2:D10")) Is Nothing Then
    MsgBox "VC ALTEROU o intervalo 1"
ElseIf Not Intersect(Target, Range("f2:g10")) Is Nothing Then
    MsgBox "VC ALTEROU o intervalo 2"
End If

End Sub
```