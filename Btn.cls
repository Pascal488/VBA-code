Private Sub LOGBTN_Click()
Dim zuchu As DAO.Recordset
Dim SQL As String
SQL = "SELECT * FROM UserTable WHERE UserName =' + Me.Usernametxt.Value + " '"

Set zuchu = CurrentDb.OpenRecordset(SQL)

If zuchu.EOF Then
  incorrectseco
  Exit Sub
End If

zuchu.MoveLast
If zuchu("Password") <> Nz(Me.Passwordtxt, "") Then
    incorectpass
    Exit Sub
End If
    
 
End Sub
Private Sub incorectpass()
     Me.Passwordtxt.Visible = True
     Me.Usernametxt.SetFocus
     Me.Usernametxt.BorderColor = RGB(255, 0, 0)
     Exit Sub
End Sub
