Attribute VB_Name = "modArrayIsProcedures"

Public Function IsArrayAllDefault(InputArray As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests if the array contains all default values for its data type.
'   Variable Type           Value
'   -------------           -------------------
'   Variant                 Empty
'   String                  vbNullString
'   Numeric                 0
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
Dim DefaultValue As Variant

'''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''
IsArrayAllDefault = False

'''''''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
'''''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    IsArrayAllDefault = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure array is allocated. An
' unallocated is considered to be
' all the same type. Return True.
''''''''''''''''''''''''''''''''''
If IsArrayAllocated(Arr:=InputArray) = False Then
    IsArrayAllDefault = True
    Exit Function
End If
    
''''''''''''''''''''''''''''''''''
' Test the type of variable
''''''''''''''''''''''''''''''''''
Select Case VarType(InputArray)
    Case vbArray + vbVariant
        DefaultValue = Empty
    Case vbArray + vbString
        DefaultValue = vbNullString
    Case Is > vbArray
        DefaultValue = 0
End Select
For Ndx = LBound(InputArray) To UBound(InputArray)
    If IsObject(InputArray(Ndx)) Then
        If Not InputArray(Ndx) Is Nothing Then
            Exit Function
        Else
            
        End If
    Else
        If VarType(InputArray(Ndx)) <> vbEmpty Then
            If InputArray(Ndx) <> DefaultValue Then
                Exit Function
            End If
        End If
    End If
Next Ndx

'''''''''''''''''''''''''''''''
' If we make it out of the loop,
' the array is all defaults.
' Return True.
'''''''''''''''''''''''''''''''
IsArrayAllDefault = True

End Function

Public Function IsArrayAllNumeric(Arr As Variant, _
    Optional AllowNumericStrings As Boolean = False) As Boolean
'Tests if the Array is entirely numeric. 
'False otherwise. The AllowNumericStrings
' parameter indicates whether strings containing numeric data are considered numeric. If this
' parameter is True, a numeric string is considered a numeric variable. If this parameter is
' omitted or False, a numeric string is not considered a numeric variable.
' Variants that are numeric or Empty are allowed. Variants that are arrays, objects, or
' non-numeric data are not allowed.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long

''''''''''''''''''''''''''''
' Ensure Arr is an array.
''''''''''''''''''''''''''''
If IsArray(Arr) = False Then
    IsArrayAllNumeric = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Ensure Arr is allocated (non-empty).
''''''''''''''''''''''''''''''''''''''
If IsArrayEmpty(Arr:=Arr) = True Then
    IsArrayAllNumeric = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Loop through the array.
'''''''''''''''''''''''''''''''''''''
For Ndx = LBound(Arr) To UBound(Arr)
    Select Case VarType(Arr(Ndx))
        Case vbInteger, vbLong, vbDouble, vbSingle, vbCurrency, vbDecimal, vbEmpty
            ' all valid numeric types
        
        Case vbString
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' For strings, check the AllowNumericStrings parameter.
            ' If True and the element is a numeric string, allow it.
            ' If it is a non-numeric string, exit with False.
            ' If AllowNumericStrings is False, all strings, even
            ' numeric strings, will cause a result of False.
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If AllowNumericStrings = True Then
                '''''''''''''''''''''''''''''''''
                ' Allow numeric strings.
                '''''''''''''''''''''''''''''''''
                If IsNumeric(Arr(Ndx)) = False Then
                    IsArrayAllNumeric = False
                    Exit Function
                End If
            Else
                IsArrayAllNumeric = False
                Exit Function
            End If
        Case vbVariant
            '''''''''''''''''''''''''''''''''''''''''''''
            ' For Variants, disallow Arrays and Objects.
            ' If the element is not an array or an object,
            ' test whether it is numeric. Allow numeric
            ' Varaints.
            '''''''''''''''''''''''''''''''''''''''''''''
            If IsArray(Arr(Ndx)) = True Then
                IsArrayAllNumeric = False
                Exit Function
            End If
            If IsObject(Arr(Ndx)) = True Then
                IsArrayAllNumeric = False
                Exit Function
            End If
            
            If IsNumeric(Arr(Ndx)) = False Then
                IsArrayAllNumeric = False
                Exit Function
            End If
                
        Case Else
            ' any other data type returns False
            IsArrayAllNumeric = False
            Exit Function
    End Select
Next Ndx

IsArrayAllNumeric = True

End Function

Public Function IsArrayAllocated(Arr As Variant) As Boolean
'Tests if the array is allocated - either a static array or a dynamic array that has been sized with Redim
' FALSE if the array is not allocated (a dynamic that has not yet
' been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
' allocated.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is just the reverse of IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim N As Long
On Error Resume Next

' if Arr is not an array, return FALSE and get out.
If IsArray(Arr) = False Then
    IsArrayAllocated = False
    Exit Function
End If

' Attempt to get the UBound of the array. If the array has not been allocated,
' an error will occur. Test Err.Number to see if an error occurred.
N = UBound(Arr, 1)
If (Err.Number = 0) Then
    ''''''''''''''''''''''''''''''''''''''
    ' Under some circumstances, if an array
    ' is not allocated, Err.Number will be
    ' 0. To acccomodate this case, we test
    ' whether LBound <= Ubound. If this
    ' is True, the array is allocated. Otherwise,
    ' the array is not allocated.
    '''''''''''''''''''''''''''''''''''''''
    If LBound(Arr) <= UBound(Arr) Then
        ' no error. array has been allocated.
        IsArrayAllocated = True
    Else
        IsArrayAllocated = False
    End If
Else
    ' error. unallocated array
    IsArrayAllocated = False
End If

End Function

Public Function IsArrayDynamic(ByRef Arr As Variant) As Boolean
' Tests whether Array is a dynamic array.
' Note that if you attempt to ReDim a static array in the same procedure in which it is
' declared, you'll get a compiler error and your code won't run at all.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim LUBound As Long

' If we weren't passed an array, get out now with a FALSE result
If IsArray(Arr) = False Then
    IsArrayDynamic = False
    Exit Function
End If

' If the array is empty, it hasn't been allocated yet, so we know
' it must be a dynamic array.
If IsArrayEmpty(Arr:=Arr) = True Then
    IsArrayDynamic = True
    Exit Function
End If

' Save the UBound of Arr.
' This value will be used to restore the original UBound if Arr
' is a single-dimensional dynamic array. Unused if Arr is multi-dimensional,
' or if Arr is a static array.
LUBound = UBound(Arr)

On Error Resume Next
Err.Clear

' Attempt to increase the UBound of Arr and test the value of Err.Number.
' If Arr is a static array, either single- or multi-dimensional, we'll get a
' C_ERR_ARRAY_IS_FIXED_OR_LOCKED error. In this case, return FALSE.
'
' If Arr is a single-dimensional dynamic array, we'll get C_ERR_NO_ERROR error.
'
' If Arr is a multi-dimensional dynamic array, we'll get a
' C_ERR_SUBSCRIPT_OUT_OF_RANGE error.
'
' For either C_NO_ERROR or C_ERR_SUBSCRIPT_OUT_OF_RANGE, return TRUE.
' For C_ERR_ARRAY_IS_FIXED_OR_LOCKED, return FALSE.

ReDim Preserve Arr(LBound(Arr) To LUBound + 1)

Select Case Err.Number
    Case C_ERR_NO_ERROR
        ' We successfully increased the UBound of Arr.
        ' Do a ReDim Preserve to restore the original UBound.
        ReDim Preserve Arr(LBound(Arr) To LUBound)
        IsArrayDynamic = True
    Case C_ERR_SUBSCRIPT_OUT_OF_RANGE
        ' Arr is a multi-dimensional dynamic array.
        ' Return True.
        IsArrayDynamic = True
    Case C_ERR_ARRAY_IS_FIXED_OR_LOCKED
        ' Arr is a static single- or multi-dimensional array.
        ' Return False
        IsArrayDynamic = False
    Case Else
        ' We should never get here.
        ' Some unexpected error occurred. Be safe and return False.
        IsArrayDynamic = False
End Select

End Function

Public Function IsArrayEmpty(Arr As Variant) As Boolean
'Tests whether the array is empty (unallocated). 
'Returns TRUE or FALSE.
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
' This function is really the reverse of IsArrayAllocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim LB As Long
Dim UB As Long

Err.Clear
On Error Resume Next
If IsArray(Arr) = False Then
    ' we weren't passed an array, return True
    IsArrayEmpty = True
End If

' Attempt to get the UBound of the array. If the array is
' unallocated, an error will occur.
UB = UBound(Arr, 1)
If (Err.Number <> 0) Then
    IsArrayEmpty = True
Else
    ''''''''''''''''''''''''''''''''''''''''''
    ' On rare occassion, under circumstances I
    ' cannot reliably replictate, Err.Number
    ' will be 0 for an unallocated, empty array.
    ' On these occassions, LBound is 0 and
    ' UBoung is -1.
    ' To accomodate the weird behavior, test to
    ' see if LB > UB. If so, the array is not
    ' allocated.
    ''''''''''''''''''''''''''''''''''''''''''
    Err.Clear
    LB = LBound(Arr)
    If LB > UB Then
        IsArrayEmpty = True
    Else
        IsArrayEmpty = False
    End If
End If

End Function

Public Function IsArrayObjects(InputArray As Variant, _
    Optional AllowNothing As Boolean = True) As Boolean
'Tests if InputArray is entirely objects (Nothing objects are optionally allowed -- default it true, allow Nothing objects). 
'Set the AllowNothing to true or false to indicate whether Nothing objects are allowed.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long

''''''''''''''''''''''''''''''''
' Set the default return value.
''''''''''''''''''''''''''''''''
IsArrayObjects = False

''''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''
' Ensure we have a single dimensional
' array.
'''''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(Arr:=InputArray)
    Case 0
        ''''''''''''''''''''''''''''''''''
        ' Unallocated dynamic array.
        ' Not allowed.
        ''''''''''''''''''''''''''''''''''
        Exit Function
    Case 1
         '''''''''''''''''''''''''''''''''
         ' OK
         '''''''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''''''
        ' Multi-dimensional array.
        ' Not allowed.
        ''''''''''''''''''''''''''''''''
        Exit Function
End Select

For Ndx = LBound(InputArray) To UBound(InputArray)
    If IsObject(InputArray(Ndx)) = False Then
        Exit Function
    End If
    If InputArray(Ndx) Is Nothing Then
        If AllowNothing = False Then
            Exit Function
        End If
    End If
Next Ndx

IsArrayObjects = True

End Function

Public Function IsArraySorted(TestArray As Variant, _
    Optional Descending As Boolean = False) As Variant
'Determines whether a single-dimensional array is sorted. 
'Because sorting is an expensive operation, especially so on large array of Variants,
' you may want to determine if an array is already in sorted order prior to doing an actual sort.
' This function returns True if an array is in sorted order (either ascending or
' descending order, depending on the value of the Descending parameter -- default
' is false = Ascending). The decision to do a string comparison (with StrComp) or
' a numeric comparison (with < or >) is based on the data type of the first
' element of the array.
' If TestArray is not an array, is an unallocated dynamic array, or has more than
' one dimension, or the VarType of TestArray is not compatible, the function
' returns NULL.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim StrCompResultFail As Long
Dim NumericResultFail As Boolean
Dim Ndx As Long
Dim NumCompareResult As Boolean
Dim StrCompResult As Long

Dim IsString As Boolean
Dim VType As VbVarType

''''''''''''''''''''''''''''''''''
' Ensure TestArray is an array.
''''''''''''''''''''''''''''''''''
If IsArray(TestArray) = False Then
    IsArraySorted = Null
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a single dimensional array.
''''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=TestArray) <> 1 Then
    IsArraySorted = Null
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''
' The following code sets the values of
' comparison that will indicate that the
' array is unsorted. It the result of
' StrComp (for strings) or ">=" (for
' numerics) equals the value specified
' below, we know that the array is
' unsorted.
'''''''''''''''''''''''''''''''''''''''''''''
If Descending = True Then
    StrCompResultFail = -1
    NumericResultFail = False
Else
    StrCompResultFail = 1
    NumericResultFail = True
End If

''''''''''''''''''''''''''''''''''''''''''''''
' Determine whether we are going to do a string
' comparison or a numeric comparison.
''''''''''''''''''''''''''''''''''''''''''''''
VType = VarType(TestArray(LBound(TestArray)))
Select Case VType
    Case vbArray, vbDataObject, vbEmpty, vbError, vbNull, vbObject, vbUserDefinedType
    '''''''''''''''''''''''''''''''''
    ' Unsupported types. Reutrn Null.
    '''''''''''''''''''''''''''''''''
        IsArraySorted = Null
        Exit Function
    Case vbString, vbVariant
    '''''''''''''''''''''''''''''''''
    ' Compare as string
    '''''''''''''''''''''''''''''''''
        IsString = True
    Case Else
    '''''''''''''''''''''''''''''''''
    ' Compare as numeric
    '''''''''''''''''''''''''''''''''
        IsString = False
End Select

For Ndx = LBound(TestArray) To UBound(TestArray) - 1
    If IsString = True Then
        StrCompResult = StrComp(TestArray(Ndx), TestArray(Ndx + 1))
        If StrCompResult = StrCompResultFail Then
            IsArraySorted = False
            Exit Function
        End If
    Else
        NumCompareResult = (TestArray(Ndx) >= TestArray(Ndx + 1))
        If NumCompareResult = NumericResultFail Then
            IsArraySorted = False
            Exit Function
        End If
    End If
Next Ndx


''''''''''''''''''''''''''''
' If we made it out of  the
' loop, then the array is
' in sorted order. Return
' True.
''''''''''''''''''''''''''''
IsArraySorted = True
End Function

Public Function IsNumericDataType(TestVar As Variant) As Boolean
'Indicates whether the data type of a variable is a numeric data type. 
'It will return TRUE for all of the following data types:
'       vbCurrency
'       vbDecimal
'       vbDouble
'       vbInteger
'       vbLong
'       vbSingle
'
' It will return FALSE for any other data type, including empty Variants and objects.
' If TestVar is an allocated array, it will test data type of the array
' and return TRUE or FALSE for that data type. If TestVar is an allocated
' array, it tests the data type of the first element of the array. If
' TestVar is an array of Variants, the function will indicate only whether
' the first element of the array is numeric. Other elements of the array
' may not be numeric data types. To test an entire array of variants
' to ensure they are all numeric data types, use the IsVariantArrayNumeric
' function.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
' It will return FALSE for any other data type. Use this procedure
' instead of VBA's IsNumeric function because IsNumeric will return
' TRUE if the variable is a string containing numeric data. This
' will cause problems with code like
'        Dim V1 As Variant
'        Dim V2 As Variant
'        V1 = "1"
'        V2 = "2"
'        If IsNumeric(V1) = True Then
'            If IsNumeric(V2) = True Then
'                Debug.Print V1 + V2
'            End If
'        End If
'
' The output of the Debug.Print statement will be "12", not 3,
' because V1 and V2 are strings and the '+' operator acts like
' the '&' operator when used with strings. This can lead to
' unexpected results.
'
' IsNumeric should only be used to test strings for numeric content
' when converting a string value to a numeric variable.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Element As Variant
    Dim NumDims As Long
    If IsArray(TestVar) = True Then
        NumDims = NumberOfArrayDimensions(Arr:=TestVar)
        If NumDims > 1 Then
            '''''''''''''''''''''''''''''''''
            ' this procedure does not support
            ' multi-dimensional arrays.
            '''''''''''''''''''''''''''''''''
            IsNumericDataType = False
            Exit Function
        End If
        If IsArrayAllocated(Arr:=TestVar) = True Then
            Element = TestVar(LBound(TestVar))
            Select Case VarType(Element)
                Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                    IsNumericDataType = True
                    Exit Function
                Case Else
                    IsNumericDataType = False
                    Exit Function
            End Select
        Else
            Select Case VarType(TestVar) - vbArray
                Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                    IsNumericDataType = True
                    Exit Function
                Case Else
                    IsNumericDataType = False
                    Exit Function
            End Select
        End If
    End If
    Select Case VarType(TestVar)
        Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
            IsNumericDataType = True
        Case Else
            IsNumericDataType = False
    End Select
End Function

Public Function IsVariantArrayConsistent(Arr As Variant) As Boolean
'Indicates whether an array of variants contains all the same data types. 
'Returns FALSE under the following
' circumstances:
'       Arr is not an array
'       Arr is an array but is unallocated
'       Arr is a multidimensional array
'       Arr is allocated but does not contain consistant data types.
'
' If Arr is an array of objects, objects that are Nothing are ignored.
' As long as all non-Nothing objects are the same object type, the
' function returns True.
'
' It returns TRUE if all the elements of the array have the same
' data type. If Arr is an array of a specific data types, not variants,
' (E.g., Dim V(1 To 3) As Long), the function will return True. If
' an array of variants contains an uninitialized element (VarType =
' vbEmpty) that element is skipped and not used in the comparison. The
' reasoning behind this is that an empty variable will return the
' data type of the variable to which it is assigned (e.g., it will
' return vbNullString to a String and 0 to a Double).
'
' The function does not support arrays of User Defined Types.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim FirstDataType As VbVarType
Dim Ndx As Long
'''''''''''''''''''''''''''''''''''''''''
' Exit with False if Arr is not an array.
'''''''''''''''''''''''''''''''''''''''''
If IsArray(Arr) = False Then
    IsVariantArrayConsistent = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Exit with False if Arr is not allocated.
''''''''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(Arr) = False Then
    IsVariantArrayConsistent = False
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''
' Exit with false on multi-dimensional
' arrays.
''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr) <> 1 Then
    IsVariantArrayConsistent = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Test if we have an array of a specific
' type rather than Variants. If so,
' return TRUE and get out.
''''''''''''''''''''''''''''''''''''''''''
If (VarType(Arr) <= vbArray) And _
    (VarType(Arr) <> vbVariant) Then
    IsVariantArrayConsistent = True
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' Get the data type of the first element.
''''''''''''''''''''''''''''''''''''''''''
FirstDataType = VarType(Arr(LBound(Arr)))
''''''''''''''''''''''''''''''''''''''''''
' Loop through the array and exit if
' a differing data type if found.
''''''''''''''''''''''''''''''''''''''''''
For Ndx = LBound(Arr) + 1 To UBound(Arr)
    If VarType(Arr(Ndx)) <> vbEmpty Then
        If IsObject(Arr(Ndx)) = True Then
            If Not Arr(Ndx) Is Nothing Then
                If VarType(Arr(Ndx)) <> FirstDataType Then
                    IsVariantArrayConsistent = False
                    Exit Function
                End If
            End If
        Else
            If VarType(Arr(Ndx)) <> FirstDataType Then
                IsVariantArrayConsistent = False
                Exit Function
            End If
        End If
    End If
Next Ndx

''''''''''''''''''''''''''''''''''''''''''
' If we make it out of the loop,
' then the array is consistent.
''''''''''''''''''''''''''''''''''''''''''
IsVariantArrayConsistent = True

End Function

Public Function IsVariantArrayNumeric(TestArray As Variant) As Boolean
'Tests if all the elements of an array of variants are numeric data types. 
'They need not all be the same data
' type. You can have a mix of Integer, Longs, Doubles, and Singles.
' As long as they are all numeric data types, the function will
' return TRUE. If a non-numeric data type is encountered, the
' function will return FALSE. Also, it will return FALSE if
' TestArray is not an array, or if TestArray has not been
' allocated. TestArray may be a multi-dimensional array. This
' procedure uses the IsNumericDataType function to determine whether
' a variable is a numeric data type. If there is an uninitialized
' variant (VarType = vbEmpty) in the array, it is skipped and not
' used in the comparison (i.e., Empty is considered a valid numeric
' data type since you can assign a number to it).
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim DimNdx As Long
Dim NumDims As Long

''''''''''''''''''''''''''''''''
' Ensure TestArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(TestArray) = False Then
    IsVariantArrayNumeric = False
    Exit Function
End If
''''''''''''''''''''''''''''''''
' Ensure that TestArray has been
' allocated.
''''''''''''''''''''''''''''''''
If IsArrayAllocated(Arr:=TestArray) = False Then
    IsVariantArrayNumeric = False
    Exit Function
End If
''''''''''''''''''''''''''''''''''''
' Ensure the array is a one
' dimensional array. This procedure
' will not work on multi-dimensional
' arrays.
''''''''''''''''''''''''''''''''''''
'If NumberOfArrayDimensions(Arr:=TestArray) > 1 Then
'    IsVariantArrayNumeric = False
'    Exit Function
'End If
    
NumDims = NumberOfArrayDimensions(Arr:=TestArray)
If NumDims = 1 Then
    '''''''''''''''''''''''''''''''''''
    ' single dimensional array
    '''''''''''''''''''''''''''''''''''
    For Ndx = LBound(TestArray) To UBound(TestArray)
        If IsObject(TestArray(Ndx)) = True Then
            IsVariantArrayNumeric = False
            Exit Function
        End If
        
        If VarType(TestArray(Ndx)) <> vbEmpty Then
            If IsNumericDataType(TestVar:=TestArray(Ndx)) = False Then
                IsVariantArrayNumeric = False
                Exit Function
            End If
        End If
    Next Ndx
Else
    ''''''''''''''''''''''''''''''''''''
    ' multi-dimensional array
    ''''''''''''''''''''''''''''''''''''
    For DimNdx = 1 To NumDims
        For Ndx = LBound(TestArray, DimNdx) To UBound(TestArray, DimNdx)
            If VarType(TestArray(Ndx, DimNdx)) <> vbEmpty Then
                If IsNumericDataType(TestVar:=TestArray(Ndx, DimNdx)) = False Then
                    IsVariantArrayNumeric = False
                    Exit Function
                End If
            End If
        Next Ndx
    Next DimNdx
End If

'''''''''''''''''''''''''''''''''''''''
' If we made it out of the loop, then
' the array is entirely numeric.
'''''''''''''''''''''''''''''''''''''''
IsVariantArrayNumeric = True

End Function

Private Function NumberOfArrayDimensions(Arr As Variant) As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumberOfArrayDimensions
' This function returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Integer
Dim Res As Integer
On Error Resume Next
' Loop, increasing the dimension index Ndx, until an error occurs.
' An error will occur when Ndx exceeds the number of dimension
' in the array. Return Ndx - 1.
Do
    Ndx = Ndx + 1
    Res = UBound(Arr, Ndx)
Loop Until Err.Number <> 0

NumberOfArrayDimensions = Ndx - 1

End Function
