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

---

## Do While

**Exemplos:**

```visual-basic
Sub exemplo01()

Dim i As Integer

i = 1
Do While Cells(i, 1).Value <> ""
    i = 1 + i
Loop

Cells(i, 1).Value = Date

End Sub
```

```visual-basic
Sub exemplo02()

Dim i As Integer

i = 1
Do While Cells(i, 1).Value <> ""
    If Cells(i, 1).Value = "ter" Then
        Cells(i, 1).Insert xlDown
        Cells(i, 1).Value = "ter"
        i = i + 1
    End If
    i = i + 1
Loop

End Sub
```

---

## Do Until

**Exemplos:**

```visual-basic
Sub EXEMPLO01()

Dim i As Integer

i = 1
Do Until Cells(i, 1).Value = ""
    i = i + 1
Loop

Cells(i, 1).Value = Date

End Sub
```

```visual-basic
Sub exemplo02()

Dim i As Integer
Dim R As Integer

i = WorksheetFunction.RandBetween(1, 100)

Do
    R = InputBox("Escolha um número")
    If R > i Then
        MsgBox "Escolha um número menor"
    ElseIf R < i Then
        MsgBox "Escolha um número Maior"
    End If
Loop Until R = i

MsgBox "Você acertou"

End Sub
```

Sub exemplo02()

Dim i As Integer
Dim R As Integer

i = WorksheetFunction.RandBetween(1, 100)

Do
R = InputBox("Escolha um número")
If R > i Then
MsgBox "Escolha um número menor"
ElseIf R < i Then
MsgBox "Escolha um número Maior"
End If
Loop Until R = i

MsgBox "Você acertou"

End Sub

Sub exemplo02()

Dim i As Integer
Dim R As Integer

i = WorksheetFunction.RandBetween(1, 100)

Do
R = InputBox("Escolha um número")
If R > i Then
MsgBox "Escolha um número menor"
ElseIf R < i Then
MsgBox "Escolha um número Maior"
End If
Loop Until R = i

MsgBox "Você acertou"

End Sub

Sub exemplo02()

Dim i As Integer
Dim R As Integer

i = WorksheetFunction.RandBetween(1, 100)

Do
R = InputBox("Escolha um número")
If R > i Then
MsgBox "Escolha um número menor"
ElseIf R < i Then
MsgBox "Escolha um número Maior"
End If
Loop Until R = i

MsgBox "Você acertou"

End Sub

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

---

## Do While

**Exemplos:**

```visual-basic
Sub exemplo01()

Dim i As Integer

i = 1
Do While Cells(i, 1).Value <> ""
    i = 1 + i
Loop

Cells(i, 1).Value = Date

End Sub
```

```visual-basic
Sub exemplo02()

Dim i As Integer

i = 1
Do While Cells(i, 1).Value <> ""
    If Cells(i, 1).Value = "ter" Then
        Cells(i, 1).Insert xlDown
        Cells(i, 1).Value = "ter"
        i = i + 1
    End If
    i = i + 1
Loop

End Sub
```

---

## Do Until

**Exemplos:**

```visual-basic
Sub EXEMPLO01()

Dim i As Integer

i = 1
Do Until Cells(i, 1).Value = ""
    i = i + 1
Loop

Cells(i, 1).Value = Date

End Sub
```

```visual-basic
Sub exemplo02()

Dim i As Integer
Dim R As Integer

i = WorksheetFunction.RandBetween(1, 100)

Do
    R = InputBox("Escolha um número")
    If R > i Then
        MsgBox "Escolha um número menor"
    ElseIf R < i Then
        MsgBox "Escolha um número Maior"
    End If
Loop Until R = i

MsgBox "Você acertou"

End Sub
```

Sub exemplo02()

Dim i As Integer
Dim R As Integer

i = WorksheetFunction.RandBetween(1, 100)

Do
R = InputBox("Escolha um número")
If R > i Then
MsgBox "Escolha um número menor"
ElseIf R < i Then
MsgBox "Escolha um número Maior"
End If
Loop Until R = i

MsgBox "Você acertou"

End Sub

Sub exemplo02()

Dim i As Integer
Dim R As Integer

i = WorksheetFunction.RandBetween(1, 100)

Do
R = InputBox("Escolha um número")
If R > i Then
MsgBox "Escolha um número menor"
ElseIf R < i Then
MsgBox "Escolha um número Maior"
End If
Loop Until R = i

MsgBox "Você acertou"

End Sub

Sub exemplo02()

Dim i As Integer
Dim R As Integer

i = WorksheetFunction.RandBetween(1, 100)

Do
R = InputBox("Escolha um número")
If R > i Then
MsgBox "Escolha um número menor"
ElseIf R < i Then
MsgBox "Escolha um número Maior"
End If
Loop Until R = i

MsgBox "Você acertou"

End Sub