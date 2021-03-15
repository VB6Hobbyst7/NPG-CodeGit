Attribute VB_Name = "modArrayCombineProcedures"
'<include modArraySizeProcedures.bas>

Public Function CombineTwoDArrays(Arr1 As Variant, _
    Arr2 As Variant) As Variant
'Takes two 2-dimensional arrays, Arr1 and Arr2, and returns an array combining the two. 
'The number of Rows RequiredLevel:4 http://www.cpearson.com/excel/vbaarrays.htm
' in the result is NumRows(Arr1) + NumRows(Arr2). Arr1 and
' Arr2 must have the same number of columns, and the result
' array will have that many columns. All the LBounds must
' be the same. E.g.,
' The following arrays are legal:
'        Dim Arr1(0 To 4, 0 To 10)
'        Dim Arr2(0 To 3, 0 To 10)
'
' The following arrays are illegal
'        Dim Arr1(0 To 4, 1 To 10)
'        Dim Arr2(0 To 3, 0 To 10)
'
' The returned result array is Arr1 with additional rows
' appended from Arr2. For example, the arrays
'    a    b        and     e    f
'    c    d                g    h
' become
'    a    b
'    c    d
'    e    f
'    g    h
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''
' Upper and lower bounds of Arr1.
'''''''''''''''''''''''''''''''''
Dim LBoundRow1 As Long
Dim UBoundRow1 As Long
Dim LBoundCol1 As Long
Dim UBoundCol1 As Long

'''''''''''''''''''''''''''''''''
' Upper and lower bounds of Arr2.
'''''''''''''''''''''''''''''''''
Dim LBoundRow2 As Long
Dim UBoundRow2 As Long
Dim LBoundCol2 As Long
Dim UBoundCol2 As Long

'''''''''''''''''''''''''''''''''''
' Upper and lower bounds of Result.
'''''''''''''''''''''''''''''''''''
Dim LBoundRowResult As Long
Dim UBoundRowResult As Long
Dim LBoundColResult As Long
Dim UBoundColResult As Long

'''''''''''''''''
' Index Variables
'''''''''''''''''
Dim RowNdx1 As Long
Dim ColNdx1 As Long
Dim RowNdx2 As Long
Dim ColNdx2 As Long
Dim RowNdxResult As Long
Dim ColNdxResult As Long


'''''''''''''
' Array Sizes
'''''''''''''
Dim NumRows1 As Long
Dim NumCols1 As Long

Dim NumRows2 As Long
Dim NumCols2 As Long

Dim NumRowsResult As Long
Dim NumColsResult As Long

Dim Done As Boolean
Dim Result() As Variant
Dim ResultTrans() As Variant

Dim V As Variant


'''''''''''''''''''''''''''''''
' Ensure that Arr1 and Arr2 are
' arrays.
''''''''''''''''''''''''''''''
If (IsArray(Arr1) = False) Or (IsArray(Arr2) = False) Then
    CombineTwoDArrays = Null
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure both arrays are allocated
' two dimensional arrays.
''''''''''''''''''''''''''''''''''
If (NumberOfArrayDimensions(Arr1) <> 2) Or (NumberOfArrayDimensions(Arr2) <> 2) Then
    CombineTwoDArrays = Null
    Exit Function
End If
    
'''''''''''''''''''''''''''''''''''''''
' Ensure that the LBound and UBounds
' of the second dimension are the
' same for both Arr1 and Arr2.
'''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''
' Get the existing bounds.
''''''''''''''''''''''''''
LBoundRow1 = LBound(Arr1, 1)
UBoundRow1 = UBound(Arr1, 1)

LBoundCol1 = LBound(Arr1, 2)
UBoundCol1 = UBound(Arr1, 2)

LBoundRow2 = LBound(Arr2, 1)
UBoundRow2 = UBound(Arr2, 1)

LBoundCol2 = LBound(Arr2, 2)
UBoundCol2 = UBound(Arr2, 2)

''''''''''''''''''''''''''''''''''''''''''''''''''
' Get the total number of rows for the result
' array.
''''''''''''''''''''''''''''''''''''''''''''''''''
NumRows1 = UBoundRow1 - LBoundRow1 + 1
NumCols1 = UBoundCol1 - LBoundCol1 + 1
NumRows2 = UBoundRow2 - LBoundRow2 + 1
NumCols2 = UBoundCol2 - LBoundCol2 + 1

'''''''''''''''''''''''''''''''''''''''''
' Ensure the number of columns are equal.
'''''''''''''''''''''''''''''''''''''''''
If NumCols1 <> NumCols2 Then
    CombineTwoDArrays = Null
    Exit Function
End If

NumRowsResult = NumRows1 + NumRows2

'''''''''''''''''''''''''''''''''''''''
' Ensure that ALL the LBounds are equal.
''''''''''''''''''''''''''''''''''''''''
If (LBoundRow1 <> LBoundRow2) Or _
    (LBoundRow1 <> LBoundCol1) Or _
    (LBoundRow1 <> LBoundCol2) Then
    CombineTwoDArrays = Null
    Exit Function
End If
'''''''''''''''''''''''''''''''
' Get the LBound of the columns
' of the result array.
'''''''''''''''''''''''''''''''
LBoundColResult = LBoundRow1
'''''''''''''''''''''''''''''''
' Get the UBound of the columns
' of the result array.
'''''''''''''''''''''''''''''''
UBoundColResult = UBoundCol1

UBoundRowResult = LBound(Arr1, 1) + NumRows1 + NumRows2 - 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Redim the Result array to have number of rows equal to
' number-of-rows(Arr1) + number-of-rows(Arr2)
' and number-of-columns equal to number-of-columns(Arr1)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReDim Result(LBoundRow1 To UBoundRowResult, LBoundColResult To UBoundColResult)

RowNdxResult = LBound(Result, 1) - 1

Done = False
Do Until Done
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Copy elements of Arr1 to Result
    ''''''''''''''''''''''''''''''''''''''''''''
    For RowNdx1 = LBound(Arr1, 1) To UBound(Arr1, 1)
        RowNdxResult = RowNdxResult + 1
        For ColNdx1 = LBound(Arr1, 2) To UBound(Arr1, 2)
            V = Arr1(RowNdx1, ColNdx1)
            Result(RowNdxResult, ColNdx1) = V
        Next ColNdx1
    Next RowNdx1

    '''''''''''''''''''''''''''''''''''''''''''''
    ' Copy elements of Arr2 to Result
    '''''''''''''''''''''''''''''''''''''''''''''
    For RowNdx2 = LBound(Arr2, 1) To UBound(Arr2, 1)
        RowNdxResult = RowNdxResult + 1
        For ColNdx2 = LBound(Arr2, 2) To UBound(Arr2, 2)
            V = Arr2(RowNdx2, ColNdx2)
            Result(RowNdxResult, ColNdx2) = V
        Next ColNdx2
    Next RowNdx2
    
    If RowNdxResult >= UBound(Result, 1) + (LBoundColResult = 1) Then
        Done = True
    End If
'''''''''''''
' End Of Loop
'''''''''''''
Loop
'''''''''''''''''''''''''
' Return the Result
'''''''''''''''''''''''''
CombineTwoDArrays = Result

End Function

Public Function ConcatenateArrays(ResultArray As Variant, ArrayToAppend As Variant, _
        Optional NoCompatabilityCheck As Boolean = False) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Appends ArrayToAppend to the end of ResultArray, increasing the size of ResultArray as needed. 
'ResultArray must be a dynamic array, but it need not be allocated. ArrayToAppend http://www.cpearson.com/excel/vbaarrays.htm
' may be either static or dynamic, and if dynamic it may be unallocted. If ArrayToAppend is
' unallocated, ResultArray is left unchanged. RequiredLevel:4
'
' The data types of ResultArray and ArrayToAppend must be either the same data type or
' compatible numeric types. A compatible numeric type is a type that will not cause a loss of
' precision or cause an overflow. For example, ReturnArray may be Longs, and ArrayToAppend amy
' by Longs or Integers, but not Single or Doubles because information might be lost when
' converting from Double to Long (the decimal portion would be lost). To skip the compatability
' check and allow any variable type in ResultArray and ArrayToAppend, set the NoCompatabilityCheck
' parameter to True. If you do this, be aware that you may loose precision and you may will
' get an overflow error which will cause a result of 0 in that element of ResultArra.
'
' Both ReaultArray and ArrayToAppend must be one-dimensional arrays.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim VTypeResult As VbVarType
Dim Ndx As Long
Dim Res As Long
Dim NumElementsToAdd As Long
Dim AppendNdx As Long
Dim VTypeAppend As VbVarType
Dim ResultLB As Long
Dim ResultUB As Long
Dim ResultWasAllocated As Boolean

'''''''''''''''''''''''''''''''''
' Set the default result.
''''''''''''''''''''''''''''''''
ConcatenateArrays = False

'''''''''''''''''''''''''''''''''
' Ensure ResultArray is an array.
'''''''''''''''''''''''''''''''''
If IsArray(ResultArray) = False Then
    Exit Function
End If
''''''''''''''''''''''''''''''''''
' Ensure ArrayToAppend is an array.
''''''''''''''''''''''''''''''''''
If IsArray(ArrayToAppend) = False Then
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure both arrays are single
' dimensional.
''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(ResultArray) > 1 Then
    Exit Function
End If
If NumberOfArrayDimensions(ArrayToAppend) > 1 Then
    Exit Function
End If
'''''''''''''''''''''''''''''''''''
' Ensure ResultArray is dynamic.
'''''''''''''''''''''''''''''''''''
If IsArrayDynamic(Arr:=ResultArray) = False Then
    Exit Function
End If

''''''''''''''''''''''''''''''''''''
' Ensure ArrayToAppend is allocated.
' If ArrayToAppend is not allocated,
' we have nothing to append, so
' exit with a True result.
''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(Arr:=ArrayToAppend) = False Then
    ConcatenateArrays = True
    Exit Function
End If


If NoCompatabilityCheck = False Then
    ''''''''''''''''''''''''''''''''''''''
    ' Ensure the array are compatible
    ' data types.
    ''''''''''''''''''''''''''''''''''''''
    If AreDataTypesCompatible(DestVar:=ResultArray, SourceVar:=ArrayToAppend) = False Then
        '''''''''''''''''''''''''''''''''''''''''''
        ' The arrays are not compatible data types.
        '''''''''''''''''''''''''''''''''''''''''''
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''
    ' If one array is an array of
    ' objects, ensure the other contains
    ' all objects (or Nothing)
    ''''''''''''''''''''''''''''''''''''
    If VarType(ResultArray) - vbArray = vbObject Then
        If IsArrayAllocated(ArrayToAppend) = True Then
            For Ndx = LBound(ArrayToAppend) To UBound(ArrayToAppend)
                If IsObject(ArrayToAppend(Ndx)) = False Then
                    Exit Function
                End If
            Next Ndx
        End If
    End If
End If
    
    
'''''''''''''''''''''''''''''''''''''''
' Get the number of elements in
' ArrrayToAppend
'''''''''''''''''''''''''''''''''''''''
NumElementsToAdd = UBound(ArrayToAppend) - LBound(ArrayToAppend) + 1
''''''''''''''''''''''''''''''''''''''''
' Get the bounds for resizing the
' ResultArray. If ResultArray is allocated
' use the LBound and UBound+1. If
' ResultArray is not allocated, use
' the LBound of ArrayToAppend for both
' the LBound and UBound of ResultArray.
''''''''''''''''''''''''''''''''''''''''

If IsArrayAllocated(Arr:=ResultArray) = True Then
    ResultLB = LBound(ResultArray)
    ResultUB = UBound(ResultArray)
    ResultWasAllocated = True
    ReDim Preserve ResultArray(ResultLB To ResultUB + NumElementsToAdd)
Else
    ResultUB = UBound(ArrayToAppend)
    ResultWasAllocated = False
    ReDim ResultArray(LBound(ArrayToAppend) To UBound(ArrayToAppend))
End If

''''''''''''''''''''''''''''''''''''''''
' Copy the data from ArrayToAppend to
' ResultArray.
''''''''''''''''''''''''''''''''''''''''
If ResultWasAllocated = True Then
    ''''''''''''''''''''''''''''''''''''''''''
    ' If ResultArray was allocated, we
    ' have to put the data from ArrayToAppend
    ' at the end of the ResultArray.
    ''''''''''''''''''''''''''''''''''''''''''
    AppendNdx = LBound(ArrayToAppend)
    For Ndx = ResultUB + 1 To UBound(ResultArray)
        If IsObject(ArrayToAppend(AppendNdx)) = True Then
            Set ResultArray(Ndx) = ArrayToAppend(AppendNdx)
        Else
            ResultArray(Ndx) = ArrayToAppend(AppendNdx)
        End If
        AppendNdx = AppendNdx + 1
        If AppendNdx > UBound(ArrayToAppend) Then
            Exit For
        End If
    Next Ndx
Else
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' If ResultArray was not allocated, we simply
    ' copy element by element from ArrayToAppend
    ' to ResultArray.
    ''''''''''''''''''''''''''''''''''''''''''''''
    For Ndx = LBound(ResultArray) To UBound(ResultArray)
        If IsObject(ArrayToAppend(Ndx)) = True Then
            Set ResultArray(Ndx) = ArrayToAppend(Ndx)
        Else
            ResultArray(Ndx) = ArrayToAppend(Ndx)
        End If
    Next Ndx

End If
'''''''''''''''''''''''
' Success. Return True.
'''''''''''''''''''''''
ConcatenateArrays = True

End Function

Public Function VectorsToArray(Arr As Variant, ParamArray Vectors()) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Takes 1 or more single-dimensional arrays and converts them into a single multi-dimensional array. 
'Each array in Vectors comprises one row of the new array. The number of columns in the
' new array is the maximum of the number of elements in each vector.
' Arr MUST be a dynamic array of a data type compatible with ALL the
' elements in each Vector. The code does NOT trap for an error RequiredLevel:4 http://www.cpearson.com/excel/vbaarrays.htm
' 13 - Type Mismatch.
'
' If the Vectors are of differing sizes, Arr is sized to hold the
' maximum number of elements in a Vector. The procedure Erases the
' Arr array, so when it is reallocated with Redim, all elements will
' be the reset to their default value (0 or vbNullString or Empty).
' Unused elements in the new array will remain the default value for
' that data type.
'
' Each Vector in Vectors must be a single dimensional array, but
' the Vectors may be of different sizes and LBounds.
'
' Each element in each Vector must be a simple data type. The elements
' may NOT be Object, Arrays, or User-Defined Types.
'
' The rows and columns of the result array are 0-based, regardless of
' the LBound of each vector and regardless of the Option Base statement.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Vector As Variant
Dim VectorNdx As Long
Dim NumElements As Long
Dim NumRows As Long
Dim NumCols As Long
Dim RowNdx As Long
Dim ColNdx As Long
Dim VType As VbVarType

'''''''''''''''''''''''''''''''''''
' Ensure we have an Array
''''''''''''''''''''''''''''''''''
If IsArray(Arr) = False Then
    VectorsToArray = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure we have a dynamic array
''''''''''''''''''''''''''''''''''
If IsArrayDynamic(Arr) = False Then
    VectorsToArray = False
    Exit Function
End If
    
'''''''''''''''''''''''''''''''''
' Ensure that at least one vector
' was passed in Vectors
'''''''''''''''''''''''''''''''''
If IsMissing(Vectors) = True Then
    VectorsToArray = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''''
' Loop through Vectors to determine the
' size of the result array. We do this
' loop first to prevent having to do
' a Redim Preserve. This requires looping
' through Vectors a second time, but this
' is still faster than doing Redim Preserves.
'''''''''''''''''''''''''''''''''''''''''''''''
For Each Vector In Vectors
    ''''''''''''''''''''''''''''
    ' Ensure Vector is single
    ' dimensional array. This
    ' will take care of the case
    ' if Vector is an unallocated
    ' array (NumberOfArrayDimensions = 0
    ' for an unallocated array).
    ''''''''''''''''''''''''''''
    If NumberOfArrayDimensions(Vector) <> 1 Then
        VectorsToArray = False
        Exit Function
    End If
    '''''''''''''''''''''''''''''''''''''
    ' Ensure that Vector is not an array.
    '''''''''''''''''''''''''''''''''''''
    If IsArray(Vector) = False Then
        VectorsToArray = False
        Exit Function
    End If
    '''''''''''''''''''''''''''''''''
    ' Increment the number of rows.
    ' Each Vector is one row or the
    ' result array. Test the size
    ' of Vector. If it is larger
    ' than the existing value of
    ' NumCols, set NumCols to the
    ' new, larger, value.
    '''''''''''''''''''''''''''''''''
    NumRows = NumRows + 1
    If NumCols < UBound(Vector) - LBound(Vector) + 1 Then
        NumCols = UBound(Vector) - LBound(Vector) + 1
    End If
Next Vector
''''''''''''''''''''''''''''''''''''''''''''
' Redim Arr to the appropriate size. Arr
' is 0-based in both directions, regardless
' of the LBound of the original Arr and
' regardless of the LBounds of the Vectors.
''''''''''''''''''''''''''''''''''''''''''''
ReDim Arr(0 To NumRows - 1, 0 To NumCols - 1)

'''''''''''''''''''''''''''''''
' Loop row-by-row.
For RowNdx = 0 To NumRows - 1
    ''''''''''''''''''''''''''''''''
    ' Loop through the columns.
    ''''''''''''''''''''''''''''''''
    For ColNdx = 0 To NumCols - 1
        ''''''''''''''''''''''''''''
        ' Set Vector (a Variant) to
        ' the Vectors(RowNdx) array.
        ' We declare Vector as a
        ' variant so it can take an
        ' array of any simple data
        ' type.
        ''''''''''''''''''''''''''''
        Vector = Vectors(RowNdx)
        '''''''''''''''''''''''''''''
        ' The vectors need not ber
        If ColNdx < UBound(Vector) - LBound(Vector) + 1 Then
            VType = VarType(Vector(LBound(Vector) + ColNdx))
            If VType >= vbArray Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Test for VType >= vbArray. The VarType of an array
                ' is vbArray + VarType(element of array). E.g., the
                ' VarType of an array of Longs equal vbArray + vbLong.
                ' Anything greater than or equal to vbArray is an
                ' array of some time.
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                VectorsToArray = False
                Exit Function
            End If
            If VType = vbObject Then
                VectorsToArray = False
                Exit Function
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Vector(LBound(Vector) + ColNdx) is
            ' a simple data type. If Vector(LBound(Vector) + ColNdx)
            ' is not a compatible data type with Arr, then a Type
            ' Mismatch error will occur. We do NOT trap this error.
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Arr(RowNdx, ColNdx) = Vector(LBound(Vector) + ColNdx)
        End If
    Next ColNdx
Next RowNdx

VectorsToArray = True

End Function

