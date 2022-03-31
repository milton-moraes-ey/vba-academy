Attribute VB_Name = "Resolucao_Lista03"
Option Explicit

Sub exercicio21()

    Dim arr(2) As String
    arr(0) = InputBox("Informe um nome")
    arr(1) = InputBox("Informe um nome")
    arr(2) = InputBox("Informe um nome")
    
    MsgBox arr(0)
    MsgBox arr(1)
    MsgBox arr(2)

End Sub

Sub exercicio22()

    Dim num As Integer
    Dim arr() As String
    Dim i As Integer
    Dim name As String
    
    num = InputBox("Informe um numero de 2 a 5")
    
    ReDim arr(num) As String
    
    For i = 0 To num
        arr(i) = InputBox("Informe um nome")
        name = name & arr(i) & " "
    Next
    
    MsgBox name

End Sub

Sub exercicio23()

    Dim arr() As String
    Dim num As Integer, i As Integer
    ReDim arr(2) As String
    
    arr(0) = "Milton"
    arr(1) = "Sara"
    arr(2) = "Soares"
    
    num = InputBox("Informe um número de 1 a 3")
    num = num + UBound(arr) + 1
    
    ReDim Preserve arr(num) As String
    
    For i = 4 To num
        arr(i) = InputBox("Informe Nomes")
    Next
    
    MsgBox arr(0)

End Sub

Sub exercicio24()

    Dim mtz(3, 3) As Integer
    
    mtz(1, 1) = 11
    mtz(1, 2) = 12
    mtz(1, 3) = 13
    
    mtz(2, 1) = 21
    mtz(2, 2) = 22
    mtz(2, 3) = 23
    
    mtz(3, 1) = 31
    mtz(3, 2) = 32
    mtz(3, 3) = 33
    
    MsgBox mtz(3, 3)

End Sub

Sub exercicio25()

    Dim meses(11) As String
    
    meses(0) = "JANEIRO"
    meses(1) = "FEVEREIRO"
    meses(2) = "MARÇO"
    meses(3) = "ABRIL"
    meses(4) = "MAIO"
    meses(5) = "JUNHO"
    meses(6) = "JULHO"
    meses(7) = "AGOSTO"
    meses(8) = "SETEMBRO"
    meses(9) = "OUTUBRO"
    meses(10) = "NOVEMBRO"
    meses(11) = "DEZEMBRO"
    
    Range("C9:N9").Value = meses()

End Sub

Sub exercicio26()

    Dim meses(11) As String
    
    meses(0) = "JANEIRO"
    meses(1) = "FEVEREIRO"
    meses(2) = "MARÇO"
    meses(3) = "ABRIL"
    meses(4) = "MAIO"
    meses(5) = "JUNHO"
    meses(6) = "JULHO"
    meses(7) = "AGOSTO"
    meses(8) = "SETEMBRO"
    meses(9) = "OUTUBRO"
    meses(10) = "NOVEMBRO"
    meses(11) = "DEZEMBRO"

    Range("C9:C20").Value = Application.Transpose(meses())

End Sub

Sub exercicio27()

    Dim matriz(), rng As Range
    
    matriz() = rng.Value
    
    MsgBox matriz(6, 5)

End Sub

Sub exercicio28()

    Dim matriz()
    matriz() = Range("C9").CurrentRegion
    MsgBox UBound(matriz)
    MsgBox UBound(matriz, 2)
    
End Sub

Sub exercicio29()
    
    Dim matriz()
    matriz() = Range("C9").CurrentRegion
    
    MsgBox matriz(1, 1)
    MsgBox matriz(1, UBound(matriz, 2)) ' (1, C)
    MsgBox matriz(UBound(matriz), 1) ' (L, 1)
    MsgBox matriz(UBound(matriz), UBound(matriz, 2)) ' (L, C)
    
End Sub

Sub exercicio30()

    Dim arr(2 To 6) As Worksheet, num As Integer
    
    Set arr(2) = Sheets(2)
    Set arr(3) = Sheets(3)
    Set arr(4) = Sheets(4)
    Set arr(5) = Sheets(5)
    Set arr(6) = Sheets(6)
    
    num = InputBox("Informe um numero de 2 a 6")
    
    
    MsgBox arr(num).name
    
End Sub
