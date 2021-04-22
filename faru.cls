Option Explicit
Dim rs As Recordset
set rs = currentDb.openrecordset("UserTable", dbOpenSnapshot, dbReadOnly)

rs.FindFirst "Username='" & Me.Usernametxt & "'"
if rs.Nomatch =True then
me.labelname.visible =true
me.componentname.setfocus
exit sub
end if
me.labelname.visible = false
