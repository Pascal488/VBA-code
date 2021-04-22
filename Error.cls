Private Sub passconfig()
Dim rs As Recordset
Set rs = CurrentDb.OpenRecordset("UserTable", dbOpenSnapshot, dbReadOnly)
rs.FindFirst "Password='" & Me.Passwordtxt & "'"
rs.MoveLast
If rs("Password") <> Nz(Me.Passwordtxt, "") Then
    incorectpass
    Exit Sub
End If
    Me.Passwordtxt.Visible = False
End Sub

Private Sub incorectpass()
     Me.Passwordtxt.Visible = True
     Me.Usernametxt.SetFocus
     Me.Usernametxt.BorderColor = RGB(255, 0, 0)
     Exit Sub
End Sub
