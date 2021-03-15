Attribute VB_Name = "modArrayTransposeProcedures"
'<include modArrayIsProcedures.bas><include modArraySizeProcedures.bas>

Public Function TransposeArray(InputArr As Variant, OutputArr As Variant) As Boolean
'Transposes a two-dimensional array. 
' It returns True if successful or False if an error occurs. InputArr must be two-dimensions. 
' OutputArr must be a dynamic array. It will be Erased and resized, so any existing content will
' be destroyed.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim RowNdx As Long
Dim ColNdx As Long
Dim LB1 As Long
Dim LB2 As Long
Dim UB1 As Long
Dim UB2 As Long

'''''''''''''''''''''''''''''''''''
' Ensure InputArr and OutputArr
' are arrays.
'''''''''''''''''''''''''''''''''''
If (IsArray(InputArr) = False) Or (IsArray(OutputArr) = False) Then
    TransposeArray = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''
' Ensure OutputArr is a dynamic
' array.
'''''''''''''''''''''''''''''''''''
If IsArrayDynamic(Arr:=OutputArr) = False Then
    TransposeArray = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure InputArr is two-dimensions,
' no more, no lesss.
''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=InputArr) <> 2 Then
    TransposeArray = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''
' Get the Lower and Upper bounds of
' InputArr.
'''''''''''''''''''''''''''''''''''''''
LB1 = LBound(InputArr, 1)
LB2 = LBound(InputArr, 2)
UB1 = UBound(InputArr, 1)
UB2 = UBound(InputArr, 2)

'''''''''''''''''''''''''''''''''''''''''
' Erase and ReDim OutputArr
'''''''''''''''''''''''''''''''''''''''''
Erase OutputArr
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Redim the Output array. Not the that the LBound and UBound
' values are preserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReDim OutputArr(LB2 To LB2 + UB2 - LB2, LB1 To LB1 + UB1 - LB1)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Loop through the elemetns of InputArr and put each value
' in the proper element of the tranposed array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For RowNdx = LBound(InputArr, 2) To UBound(InputArr, 2)
    For ColNdx = LBound(InputArr, 1) To UBound(InputArr, 1)
        OutputArr(RowNdx, ColNdx) = InputArr(ColNdx, RowNdx)
    Next ColNdx
Next RowNdx

'''''''''''''''''''''''''
' Success -- return True.
'''''''''''''''''''''''''
TransposeArray = True

End Function
