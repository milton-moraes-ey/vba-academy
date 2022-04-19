# Estruturas de Repetição em VBA

## For Next

**Syntaxe**:

```visual-basic
For {VAR_NUMERICA} = {VALOR} To {NUM_REPETICOES}
	'
  'CÓDIGO A SER EXECUTADO DENTRO DO LOOP
	'
Next
```

**Exemplos:**

Tanto a sub ex01 e a sub ex02 preenchem as celulas A1:A5 com os números de 1 a 5

```visual-basic
Sub ex01()
    Dim i As Integer, n As Integer: n = 1
    For i = 1 To 5
        Cells(n, 1).Value = n
        n = n + 1
    Next
End Sub
```

```visual-basic
Sub ex02()
    Dim i As Integer
    For i = 1 To 5
      Cells(i, 1).Value = i
    Next
End Sub
```

```visual-basic
Sub ex03()

Dim i As Integer, n As Integer

n = Sheets.Count

For i = 1 To n
    Cells(i, 1).Value = Sheets(i).Name
Next

End Sub
```

- Adicionando os meses do ano e os dias da semana em uma Combobox:

```visual-basic
sub addDiaSemana()

Dim i As Integer
For i = 1 To 7
    Me.ComboBox1.AddItem (WeekdayName(i))
Next

end sub
```

```visual-basic
sub addMesesComboBox

Dim i As Integer
For i = 1 To 12
    Me.ComboBox1.AddItem (MonthName(i))
Next

end sub
```