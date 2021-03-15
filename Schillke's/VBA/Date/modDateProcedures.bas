Attribute VB_Name = "modDateProcedures"
Function IsRealDate(varDate As Variant) As Boolean
'Test if variant variable has a value that may be considered a date
' RequiredLevel:3
    IsRealDate = False
    On Error Resume Next
    If IsDate(varDate) Then IsRealDate = True
    If Year(varDate) = "" Then IsRealDate = False
    If Year(varDate) <= 1900 Then IsRealDate = False
End Function
