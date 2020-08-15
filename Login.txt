Dim rs As Recordset
Set rs = CurrentDb.OpenRecordset("UserTable", dbOpenSnapshot, dbReadOnly)
rs.FindFirst "Username='" & Me.Usernametxt & "'"

If rs.NoMatch = True Then
  incorrectseco
  Exit Sub
End If
If rs.NoMatch = False Then
    DoCmd.OpenForm "MAIN"
    MsgBox "Great login process completed"
    Exit Sub
End If
