' Esse é um script que é utilizando no evento Click de um Toggle Button, construído dentro de um UserForm
' Além disso, para o correto funcionamento, é necessário criar um Frame, no caso referenciado como Menu

Private Sub ToggleButton1_Click()
Application.ScreenUpdating = False

    Dim leftFinal As Long, leftIncial As Long, cont As Integer
    leftIncial = Me.Frame1.Width * -1
    leftFinal = 0
    
    
    If Me.ToggleButton1.Value = True Then
        For cont = leftIncial To leftFinal
            Me.Frame1.Left = cont
        Next cont
    Else
        For cont = leftFinal To leftIncial Step -1
            Me.Frame1.Left = cont
        Next cont
    End If
        
End Sub