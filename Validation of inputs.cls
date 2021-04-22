
if IsNumeric(Right(TextBox82, 1)) Then
TextBox82.BackColor = &HFF&
MsgBox (""Only letters allowed in field"")
TextBox82 = Left(TextBox82, Len(TextBox82) - 1)
Else
TextBox82.BackColor = &H80000005
End If
End Sub
