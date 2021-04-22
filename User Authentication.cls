TempVars("UserType") = rs!UserID.Value
 If rs!UserID = "3" Then
   Dim Prop As Property
   On Error GoTo SetProperty
   Set Prop = CurrentDb.CreateProperty("Allow ByPass Key", dbBoolean, False)
   CurrentDb.Properties.Append Prop
 End If
SetProperty:
   If MsgBox("Would you liike to ByPass key?", vbYesNo, "Allow Bypass?") = vbYes Then
   CurrentDb.Properties("Allow Bypass Key") = True
   Else
   CurrentDb.Properties("Allow Bypass Key") = False
   Exit Sub

End If
