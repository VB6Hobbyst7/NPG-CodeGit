Attribute VB_Name = "modNNZ"
Public Function nnz(testvalue As Variant) As Variant
'see http://access.mvps.org/Access/forms/frm0022.htm for use
'Not Numeric return zero
    If Not (IsNumeric(testvalue)) Then
        nnz = 0
    Else
        nnz = testvalue
    End If
End Function
