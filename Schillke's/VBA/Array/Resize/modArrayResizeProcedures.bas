Attribute VB_Name = "modArrayResizeProcedures"
'<include modArrayIsProcedures.bas><include modArraySizeProcedures.bas>

Public Function DeleteArrayElement(InputArray As Variant, ElementNumber As Long, _
    Optional ResizeDynamic As Boolean = False) As Boolean
'Deletes an element from InputArray, and shifts elements that are to the
' right of the deleted element to the left. If InputArray is a dynamic array, and the
' ResizeDynamic parameter is True, the array will be resized one element smaller. Otherwise,
' the right-most entry in the array is set to the default value appropriate to the data
' type of the array (0, vbNullString, Empty, or Nothing). If the array is an array of Variant
' types, the default data type is the data type of the last element in the array.
' The function returns True if the elememt was successfully deleted, or False if an error
' occurrred. This procedure works only on single-dimensional
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim VType As VbVarType

''''''''''''''''''''''''''''''''
' Set the default result
''''''''''''''''''''''''''''''''
DeleteArrayElement = False

''''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a single dimensional array
'''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=InputArray) <> 1 Then
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a valid ElementNumber
''''''''''''''''''''''''''''''''''''''''''''''
If (ElementNumber < LBound(InputArray)) Or (ElementNumber > UBound(InputArray)) Then
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''
' If we have a single element array, Erase it.
''''''''''''''''''''''''''''''''''''''''''''''
If LBound(InputArray) = UBound(InputArray) Then
    Erase InputArray
    Exit Function
End If


''''''''''''''''''''''''''''''''''''''''''''''
' Get the variable data type of the element
' we're deleting.
''''''''''''''''''''''''''''''''''''''''''''''
VType = VarType(InputArray(UBound(InputArray)))
If VType >= vbArray Then
    VType = VType - vbArray
End If
''''''''''''''''''''''''''''''''''''''''''''''
' Shift everything to the left
''''''''''''''''''''''''''''''''''''''''''''''
For Ndx = ElementNumber To UBound(InputArray) - 1
    InputArray(Ndx) = InputArray(Ndx + 1)
Next Ndx
''''''''''''''''''''''''''''''''''''''''''''''
' If ResizeDynamic is True, resize the array
' if it is dynamic.
''''''''''''''''''''''''''''''''''''''''''''''
If IsArrayDynamic(Arr:=InputArray) = True Then
    If ResizeDynamic = True Then
        ''''''''''''''''''''''''''''''''
        ' Resize the array and get out.
        ''''''''''''''''''''''''''''''''
        ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray) - 1)
        DeleteArrayElement = True
        Exit Function
    End If
End If
'''''''''''''''''''''''''''''
' Set the last element of the
' InputArray to the proper
' default value.
'''''''''''''''''''''''''''''
Select Case VType
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbDate, vbCurrency, vbDecimal
        InputArray(UBound(InputArray)) = 0
    Case vbString
        InputArray(UBound(InputArray)) = vbNullString
    Case vbArray, vbVariant, vbEmpty, vbError, vbNull, vbUserDefinedType
        InputArray(UBound(InputArray)) = Empty
    Case vbBoolean
        InputArray(UBound(InputArray)) = False
    Case vbObject
        Set InputArray(UBound(InputArray)) = Nothing
    Case Else
        InputArray(UBound(InputArray)) = 0
End Select

DeleteArrayElement = True

End Function

Function ExpandArray(Arr As Variant, WhichDim As Long, AdditionalElements As Long, _
        FillValue As Variant) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ExpandArray
' This expands a two-dimensional array in either dimension. It returns the result
' array if successful, or NULL if an error occurred. The original array is never
' changed.
' Paramters:
' --------------------
' Arr                   is the array to be expanded.
'
' WhichDim              is either 1 for additional rows or 2 for
'                       additional columns.
'
' AdditionalElements    is the number of additional rows or columns
'                       to create.
'
' FillValue             is the value to which the new array elements should be
'                       initialized.
'
' You can nest calls to Expand array to expand both the number of rows and
' columns. E.g.,
'
' C = ExpandArray(ExpandArray(Arr:=A, WhichDim:=1, AdditionalElements:=3, FillValue:="R"), _
'    WhichDim:=2, AdditionalElements:=4, FillValue:="C")
' This first adds three rows at the bottom of the array, and then adds four
' columns on the right of the array.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Result As Variant
Dim RowNdx As Long
Dim ColNdx As Long
Dim ResultRowNdx As Long
Dim ResultColNdx As Long
Dim NumRows As Long
Dim NumCols As Long
Dim NewUBound As Long

Const ROWS_ As Long = 1
Const COLS_ As Long = 2


''''''''''''''''''''''''''''
' Ensure Arr is an array.
''''''''''''''''''''''''''''
If IsArray(Arr) = False Then
    ExpandArray = Null
    Exit Function
End If
'''''''''''''''''''''''''''''''''
' Ensure Arr has two dimenesions.
'''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=Arr) <> 2 Then
    ExpandArray = Null
    Exit Function
End If
'''''''''''''''''''''''''''''''''
' Ensure the dimension is 1 or 2.
'''''''''''''''''''''''''''''''''
Select Case WhichDim
    Case 1, 2
    Case Else
        ExpandArray = Null
        Exit Function
End Select

''''''''''''''''''''''''''''''''''''
' Ensure AdditionalElements is > 0.
' If AdditionalElements  < 0, return NULL.
' If AdditionalElements  = 0, return Arr.
''''''''''''''''''''''''''''''''''''
If AdditionalElements < 0 Then
    ExpandArray = Null
    Exit Function
End If
If AdditionalElements = 0 Then
    ExpandArray = Arr
    Exit Function
End If
    
NumRows = UBound(Arr, 1) - LBound(Arr, 1) + 1
NumCols = UBound(Arr, 2) - LBound(Arr, 2) + 1
   
If WhichDim = ROWS_ Then
    '''''''''''''''
    ' Redim Result.
    '''''''''''''''
    ReDim Result(LBound(Arr, 1) To UBound(Arr, 1) + AdditionalElements, LBound(Arr, 2) To UBound(Arr, 2))
    ''''''''''''''''''''''''''''''
    ' Transfer Arr array to Result
    ''''''''''''''''''''''''''''''
    For RowNdx = LBound(Arr, 1) To UBound(Arr, 1)
        For ColNdx = LBound(Arr, 2) To UBound(Arr, 2)
            Result(RowNdx, ColNdx) = Arr(RowNdx, ColNdx)
        Next ColNdx
    Next RowNdx
    '''''''''''''''''''''''''''''''
    ' Fill the rest of the result
    ' array with FillValue.
    '''''''''''''''''''''''''''''''
    For RowNdx = UBound(Arr, 1) + 1 To UBound(Result, 1)
        For ColNdx = LBound(Arr, 2) To UBound(Arr, 2)
            Result(RowNdx, ColNdx) = FillValue
        Next ColNdx
    Next RowNdx
Else
    '''''''''''''''
    ' Redim Result.
    '''''''''''''''
    ReDim Result(LBound(Arr, 1) To UBound(Arr, 1), UBound(Arr, 2) + AdditionalElements)
    ''''''''''''''''''''''''''''''
    ' Transfer Arr array to Result
    ''''''''''''''''''''''''''''''
    For RowNdx = LBound(Arr, 1) To UBound(Arr, 1)
        For ColNdx = LBound(Arr, 2) To UBound(Arr, 2)
            Result(RowNdx, ColNdx) = Arr(RowNdx, ColNdx)
        Next ColNdx
    Next RowNdx
    '''''''''''''''''''''''''''''''
    ' Fill the rest of the result
    ' array with FillValue.
    '''''''''''''''''''''''''''''''
    For RowNdx = LBound(Arr, 1) To UBound(Arr, 1)
        For ColNdx = UBound(Arr, 2) + 1 To UBound(Result, 2)
            Result(RowNdx, ColNdx) = FillValue
        Next ColNdx
    Next RowNdx
    
End If
''''''''''''''''''''
' Return the result.
''''''''''''''''''''
ExpandArray = Result

End Function

Public Function ChangeBoundsOfArray(InputArr As Variant, _
    NewLowerBound As Long, NewUpperBound) As Boolean
'Changes the upper and lower bounds of the specified array. 
'InputArr MUST be a single-dimensional dynamic array.
' If the new size of the array (NewUpperBound - NewLowerBound + 1)
' is greater than the original array, the unused elements on
' right side of the array are the default values for the data type
' of the array. If the new size is less than the original size,
' only the first (left-most) N elements are included in the new array.
' The elements of the array may be simple variables (Strings, Longs, etc)
' Object, or Arrays. User-Defined Types are not supported.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
' The function returns True if successful, False otherwise.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim TempArr() As Variant
Dim InNdx As Long
Dim OutNdx As Long
Dim TempNdx As Long
Dim FirstIsObject As Boolean

''''''''''''''''''''''''''''''''''''
' Ensure we have an array.
''''''''''''''''''''''''''''''''''''
If IsArray(InputArr) = False Then
    ChangeBoundsOfArray = False
    Exit Function
End If
''''''''''''''''''''''''''''''''''''
' Ensure the array is dynamic.
''''''''''''''''''''''''''''''''''''
If IsArrayDynamic(InputArr) = False Then
    ChangeBoundsOfArray = False
    Exit Function
End If
''''''''''''''''''''''''''''''''''''
' Ensure the array is allocated.
''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(InputArr) = False Then
    ChangeBoundsOfArray = False
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''
' Ensure the NewLowerBound > NewUpperBound.
'''''''''''''''''''''''''''''''''''''''''''
If NewLowerBound > NewUpperBound Then
    ChangeBoundsOfArray = False
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''
' Ensure Arr is a single dimensional array.
'''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(InputArr) <> 1 Then
    ChangeBoundsOfArray = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''
' We need to save the IsObject status of the first
' element of the InputArr to properly handle
' the Empty variables is we are making the array
' larger than it was before.
'''''''''''''''''''''''''''''''''''''''''''''''''''
FirstIsObject = IsObject(InputArr(LBound(InputArr)))


''''''''''''''''''''''''''''''''''''''''''''
' Resize TempArr and save the values in
' InputArr in TempArr. TempArr will have
' an LBound of 1 and a UBound of the size
' of (NewUpperBound - NewLowerBound +1)
'''''''''''''''''''''''''''''''''''''''''''
ReDim TempArr(1 To (NewUpperBound - NewLowerBound + 1))
'''''''''''''''''''''''''''''''''''''''''''
' Load up TempArr
'''''''''''''''''''''''''''''''''''''''''''
TempNdx = 0
For InNdx = LBound(InputArr) To UBound(InputArr)
    TempNdx = TempNdx + 1
    If TempNdx > UBound(TempArr) Then
        Exit For
    End If
    
    If (IsObject(InputArr(InNdx)) = True) Then
        If InputArr(InNdx) Is Nothing Then
            Set TempArr(TempNdx) = Nothing
        Else
            Set TempArr(TempNdx) = InputArr(InNdx)
        End If
    Else
        TempArr(TempNdx) = InputArr(InNdx)
    End If
Next InNdx

''''''''''''''''''''''''''''''''''''
' Now, Erase InputArr, resize it to the
' new bounds, and load up the values from
' TempArr to the new InputArr.
''''''''''''''''''''''''''''''''''''
Erase InputArr
ReDim InputArr(NewLowerBound To NewUpperBound)
OutNdx = LBound(InputArr)
For TempNdx = LBound(TempArr) To UBound(TempArr)
    If OutNdx <= UBound(InputArr) Then
        If IsObject(TempArr(TempNdx)) = True Then
            Set InputArr(OutNdx) = TempArr(TempNdx)
        Else
            If FirstIsObject = True Then
                If IsEmpty(TempArr(TempNdx)) = True Then
                    Set InputArr(OutNdx) = Nothing
                Else
                    Set InputArr(OutNdx) = TempArr(TempNdx)
                End If
            Else
                InputArr(OutNdx) = TempArr(TempNdx)
            End If
        End If
    Else
        Exit For
    End If
    OutNdx = OutNdx + 1
Next TempNdx

'''''''''''''''''''''''''''''
' Success -- Return True
'''''''''''''''''''''''''''''
ChangeBoundsOfArray = True


End Function


