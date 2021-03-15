Attribute VB_Name = "modAccessFormData"
Public Function FormHasData(frm As Form) As Boolean
'Test if the form has any records (other than new one).
'           Return False for unbound forms, and forms with no records.
'Note:      Avoids the bug in Access 2007 where text boxes cannot use:
'               [Forms].[Form1].[Recordset].[RecordCount]
'http://allenbrowne.com/RecordCountError.html
    On Error Resume Next    'To handle unbound forms.
    FormHasData = (frm.Recordset.RecordCount <> 0&)
End Function

