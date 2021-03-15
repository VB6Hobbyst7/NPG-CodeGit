Attribute VB_Name = "modArrayGetValuesProcedures"
'<include modArrayIsProcedures.bas><include modArraySizeProcedures.bas>

Function GetColumn(Arr As Variant, ResultArr As Variant, ColumnNumber As Long) As Boolean
'Populates ResultArr with a one-dimensional array that is the specified column of Arr. 
'The existing contents of ResultArr are destroyed. ResultArr must be a dynamic array.
' Returns True or False indicating success.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RowNdx As Long
''''''''''''''''''''''''''''''
' Ensure Arr is an array.
''''''''''''''''''''''''''''''
If IsArray(Arr) = False Then
    GetColumn = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure Arr is a two-dimensional
' array.
''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr) <> 2 Then
    GetColumn = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure ResultArr is a dynamic
' array.
''''''''''''''''''''''''''''''''''
If IsArrayDynamic(ResultArr) = False Then
    GetColumn = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''
' Ensure ColumnNumber is less than
' or equal to the number of columns.
''''''''''''''''''''''''''''''''''''
If UBound(Arr, 2) < ColumnNumber Then
    GetColumn = False
    Exit Function
End If
If LBound(Arr, 2) > ColumnNumber Then
    GetColumn = False
    Exit Function
End If

Erase ResultArr
ReDim ResultArr(LBound(Arr, 1) To UBound(Arr, 1))
For RowNdx = LBound(ResultArr) To UBound(ResultArr)
    ResultArr(RowNdx) = Arr(RowNdx, ColumnNumber)
Next RowNdx

GetColumn = True

End Function


Function GetRow(Arr As Variant, ResultArr As Variant, RowNumber As Long) As Boolean
' Populates ResultArr with a one-dimensional array that is the specified row of Arr. 
'The existing contents of ResultArr are destroyed. ResultArr must be a dynamic array.
' Returns True or False indicating success.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ColNdx As Long
''''''''''''''''''''''''''''''
' Ensure Arr is an array.
''''''''''''''''''''''''''''''
If IsArray(Arr) = False Then
    GetRow = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure Arr is a two-dimensional
' array.
''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr) <> 2 Then
    GetRow = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure ResultArr is a dynamic
' array.
''''''''''''''''''''''''''''''''''
If IsArrayDynamic(ResultArr) = False Then
    GetRow = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''
' Ensure ColumnNumber is less than
' or equal to the number of columns.
''''''''''''''''''''''''''''''''''''
If UBound(Arr, 1) < RowNumber Then
    GetRow = False
    Exit Function
End If
If LBound(Arr, 1) > RowNumber Then
    GetRow = False
    Exit Function
End If

Erase ResultArr
ReDim ResultArr(LBound(Arr, 2) To UBound(Arr, 2))
For ColNdx = LBound(ResultArr) To UBound(ResultArr)
    ResultArr(ColNdx) = Arr(RowNumber, ColNdx)
Next ColNdx

GetRow = True

End Function

