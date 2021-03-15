<include modForm_BeforeInsert_CarryOver.bas>
'http://allenbrowne.com/ser-24.html
Dim strMsg As String
CarryOver Me, strMsg
If strMsg <> vbNullString Then
	MsgBox strMsg, vbInformation
End If