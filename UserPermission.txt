If DLookup("[FieldName]", "TableName", "username =[usernametxt] AND password = [passwordtxt]") >= 11 Then
DoCmd.OpenForm "FormName", , acNormal
MsgBox "Welcome"
End If
'This help to Authorize different type of users in Ms Access

