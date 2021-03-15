Attribute VB_Name = "modArraySwapProcedures"
'<include modArrayIsProcedures.bas><include modArraySizeProcedures.bas>

Function SwapArrayRows(Arr As Variant, Row1 As Long, Row2 As Long) As Variant
'Returns an array based on Arr with Row1 and Row2 swapped.
' It returns the result array or NULL if an error occurred.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim V As Variant
Dim Result As Variant
Dim RowNdx As Long
Dim ColNdx As Long

'''''''''''''''''''''''''
' Ensure Arr is an array.
'''''''''''''''''''''''''
If IsArray(Arr) = False Then
    SwapArrayRows = Null
    Exit Function
End If

''''''''''''''''''''''''''''''''
' Set Result to Arr
''''''''''''''''''''''''''''''''
Result = Arr

''''''''''''''''''''''''''''''''
' Ensure Arr is two-dimensional.
''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=Arr) <> 2 Then
    SwapArrayRows = Null
    Exit Function
End If

''''''''''''''''''''''''''''''''
' Ensure Row1 and Row2 are less
' than or equal to the number of
' rows.
''''''''''''''''''''''''''''''''
If (Row1 > UBound(Arr, 1)) Or (Row2 > UBound(Arr, 1)) Then
    SwapArrayRows = Null
    Exit Function
End If
    
'''''''''''''''''''''''''''''''''
' If Row1 = Row2, just return the
' array and exit. Nothing to do.
'''''''''''''''''''''''''''''''''
If Row1 = Row2 Then
    SwapArrayRows = Arr
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''
' Redim V to the number of columns.
'''''''''''''''''''''''''''''''''''''''''
ReDim V(LBound(Arr, 2) To UBound(Arr, 2))
'''''''''''''''''''''''''''''''''''''''''
' Put Row1 in V
'''''''''''''''''''''''''''''''''''''''''
For ColNdx = LBound(Arr, 2) To UBound(Arr, 2)
    V(ColNdx) = Arr(Row1, ColNdx)
    Result(Row1, ColNdx) = Arr(Row2, ColNdx)
    Result(Row2, ColNdx) = V(ColNdx)
Next ColNdx

SwapArrayRows = Result

End Function


Function SwapArrayColumns(Arr As Variant, Col1 As Long, Col2 As Long) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SwapArrayColumns
' This function returns an array based on Arr with Col1 and Col2 swapped.
' It returns the result array or NULL if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim V As Variant
Dim Result As Variant
Dim RowNdx As Long
Dim ColNdx As Long

'''''''''''''''''''''''''
' Ensure Arr is an array.
'''''''''''''''''''''''''
If IsArray(Arr) = False Then
    SwapArrayColumns = Null
    Exit Function
End If

''''''''''''''''''''''''''''''''
' Set Result to Arr
''''''''''''''''''''''''''''''''
Result = Arr

''''''''''''''''''''''''''''''''
' Ensure Arr is two-dimensional.
''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=Arr) <> 2 Then
    SwapArrayColumns = Null
    Exit Function
End If

''''''''''''''''''''''''''''''''
' Ensure Row1 and Row2 are less
' than or equal to the number of
' rows.
''''''''''''''''''''''''''''''''
If (Col1 > UBound(Arr, 2)) Or (Col2 > UBound(Arr, 2)) Then
    SwapArrayColumns = Null
    Exit Function
End If
    
'''''''''''''''''''''''''''''''''
' If Col1 = Col2, just return the
' array and exit. Nothing to do.
'''''''''''''''''''''''''''''''''
If Col1 = Col2 Then
    SwapArrayColumns = Arr
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''
' Redim V to the number of columns.
'''''''''''''''''''''''''''''''''''''''''
ReDim V(LBound(Arr, 1) To UBound(Arr, 1))
'''''''''''''''''''''''''''''''''''''''''
' Put Col2 in V
'''''''''''''''''''''''''''''''''''''''''
For RowNdx = LBound(Arr, 1) To UBound(Arr, 1)
    V(RowNdx) = Arr(RowNdx, Col1)
    Result(RowNdx, Col1) = Arr(RowNdx, Col2)
    Result(RowNdx, Col2) = V(RowNdx)
Next RowNdx

SwapArrayColumns = Result

End Function
