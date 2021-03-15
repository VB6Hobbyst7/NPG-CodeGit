Attribute VB_Name = "modArrayCompareProcedures"
'<include modArrayIsProcedures.bas><include modArraySizeProcedures.bas>

Public Function CompareArrays(Array1 As Variant, Array2 As Variant, _
    ResultArray As Variant, Optional CompareMode As VbCompareMethod = vbTextCompare) As Boolean
' Compares two arrays, Array1 and Array2, element by element, and puts the results of
' the comparisons in ResultArray. Each element of ResultArray will be -1, 0, or +1. A -1 indicates that
' the element in Array1 was less than the corresponding element in Array2. A 0 indicates that the
' elements are equal, and +1 indicates that the element in Array1 is greater than Array2. Both
' Array1 and Array2 must be allocated single-dimensional arrays, and ResultArray must be dynamic array
' of a numeric data type (typically Longs). Array1 and Array2 must contain the same number of elements,
' and have the same lower bound. The LBound of ResultArray will be the same as the data arrays.
'RequiredLevel:4 http://www.cpearson.com/excel/vbaarrays.htm
' An error will occur if Array1 or Array2 contains an Object or User Defined Type.
'
' When comparing elements, the procedure does the following:
' If both elements are numeric data types, they are compared arithmetically.

' If one element is a numeric data type and the other is a string and that string is numeric,
' then both elements are converted to Doubles and compared arithmetically. If the string is not
' numeric, both elements are converted to strings and compared using StrComp, with the
' compare mode set by CompareMode.
'
' If both elements are numeric strings, they are converted to Doubles and compared arithmetically.
'
' If either element is not a numeric string, the elements are converted and compared with StrComp.
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx1 As Long
Dim Ndx2 As Long
Dim ResNdx As Long
Dim S1 As String
Dim S2 As String
Dim D1 As Double
Dim D2 As Double
Dim Done As Boolean
Dim Compare As VbCompareMethod
Dim LB As Long

''''''''''''''''''''''''''''''''''''
' Set the default return value.
''''''''''''''''''''''''''''''''''''
CompareArrays = False

''''''''''''''''''''''''''''''''''''
' Ensure we have a Compare mode
' value.
''''''''''''''''''''''''''''''''''''
If CompareMode = vbBinaryCompare Then
    Compare = vbBinaryCompare
Else
    Compare = vbTextCompare
End If


''''''''''''''''''''''''''''''''''''
' Ensure we have arrays.
''''''''''''''''''''''''''''''''''''
If IsArray(Array1) = False Then
    Exit Function
End If
If IsArray(Array2) = False Then
    Exit Function
End If
If IsArray(ResultArray) = False Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''''
' Ensure ResultArray is dynamic
'''''''''''''''''''''''''''''''''''
If IsArrayDynamic(Arr:=ResultArray) = False Then
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Ensure the arrays are single-dimensional.
''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=Array1) <> 1 Then
    Exit Function
End If
If NumberOfArrayDimensions(Arr:=Array2) <> 1 Then
    Exit Function
End If
If NumberOfArrayDimensions(Arr:=Array1) > 1 Then 'allow 0 indicating non-allocated array
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Ensure the LBounds are the same
''''''''''''''''''''''''''''''''''''''''''
If LBound(Array1) <> LBound(Array2) Then
    Exit Function
End If
    

''''''''''''''''''''''''''''''''''''''''''
' Ensure the arrays are the same size.
''''''''''''''''''''''''''''''''''''''''''
If (UBound(Array1) - LBound(Array1)) <> (UBound(Array2) - LBound(Array2)) Then
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''
' Redim ResultArray to the numbr of elements
' in Array1.
''''''''''''''''''''''''''''''''''''''''''''''
ReDim ResultArray(LBound(Array1) To UBound(Array1))

Ndx1 = LBound(Array1)
Ndx2 = LBound(Array2)

''''''''''''''''''''''''''''''''''''''''''''''
' Scan each array to see if it contains objects
' or User-Defined Types. If found, exit with
' False.
''''''''''''''''''''''''''''''''''''''''''''''
For Ndx1 = LBound(Array1) To UBound(Array1)
    If IsObject(Array1(Ndx1)) = True Then
        Exit Function
    End If
    If VarType(Array1(Ndx1)) >= vbArray Then
        Exit Function
    End If
    If VarType(Array1(Ndx1)) = vbUserDefinedType Then
        Exit Function
    End If
Next Ndx1

For Ndx1 = LBound(Array2) To UBound(Array2)
    If IsObject(Array2(Ndx1)) = True Then
        Exit Function
    End If
    If VarType(Array2(Ndx1)) >= vbArray Then
        Exit Function
    End If
    If VarType(Array2(Ndx1)) = vbUserDefinedType Then
        Exit Function
    End If
Next Ndx1

Ndx1 = LBound(Array1)
Ndx2 = Ndx1
ResNdx = LBound(ResultArray)
Done = False
Do Until Done = True
''''''''''''''''''''''''''''''''''''
' Loop until we reach the end of
' the array.
''''''''''''''''''''''''''''''''''''
    If IsNumeric(Array1(Ndx1)) = True And IsNumeric(Array2(Ndx2)) Then
        D1 = CDbl(Array1(Ndx1))
        D2 = CDbl(Array2(Ndx2))
        If D1 = D2 Then
            ResultArray(ResNdx) = 0
        ElseIf D1 < D2 Then
            ResultArray(ResNdx) = -1
        Else
            ResultArray(ResNdx) = 1
        End If
    Else
        S1 = CStr(Array1(Ndx1))
        S2 = CStr(Array2(Ndx1))
        ResultArray(ResNdx) = StrComp(S1, S2, Compare)
    End If
        
    ResNdx = ResNdx + 1
    Ndx1 = Ndx1 + 1
    Ndx2 = Ndx2 + 1
    ''''''''''''''''''''''''''''''''''''''''
    ' If Ndx1 is greater than UBound(Array1)
    ' we've hit the end of the arrays.
    ''''''''''''''''''''''''''''''''''''''''
    If Ndx1 > UBound(Array1) Then
        Done = True
    End If
Loop

CompareArrays = True
End Function



