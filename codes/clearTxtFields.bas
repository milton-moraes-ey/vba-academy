Private sub btnClear_Click()
Dim ctrl as Control
  For Each ctrl in Me.Controls
        if TypeName(ctrl) = "TextBox" then
            ctrl.Value = ""
        end if
  next
end sub