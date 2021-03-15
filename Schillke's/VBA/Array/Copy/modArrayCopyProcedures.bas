Attribute VB_Name = "modArrayCopyProcedures"
'<include modArrayIsProcedures.bas><include modArraySizeProcedures.bas><include modArrayDataTypeProcedures.bas>

Public Function CopyArray(DestinationArray As Variant, SourceArray As Variant, _
        Optional NoCompatabilityCheck As Boolean = False) As Boolean
' Copies the contents of SourceArray to the DestinationaArray. Both SourceArray
' and DestinationArray may be either static or dynamic and either or both may be unallocated.
' RequiredLevel:4
' If DestinationArray is dynamic, it is resized to match SourceArray. The LBound and UBound
' of DestinationArray will be the same as SourceArray, and all elements of SourceArray will
' be copied to DestinationArray.
'
' If DestinationArray is static and has more elements than SourceArray, all of SourceArray
' is copied to DestinationArray and the right-most elements of DestinationArray are left
' intact.
'
' If DestinationArray is static and has fewer elements that SourceArray, only the left-most
' elements of SourceArray are copied to fill out DestinationArray.
'
' If SourceArray is an unallocated array, DestinationArray remains unchanged and the procedure
' terminates.
'
' If both SourceArray and DestinationArray are unallocated, no changes are made to either array
' and the procedure terminates.
'
' SourceArray may contain any type of data, including Objects and Objects that are Nothing
' (the procedure does not support arrays of User Defined Types since these cannot be coerced
' to Variants -- use classes instead of types).
'
' The function tests to ensure that the data types of the arrays are the same or are compatible.
' See the function AreDataTypesCompatible for information about compatible data types. To skip
' this compability checking, set the NoCompatabilityCheck parameter to True. Note that you may
' lose information during data conversion (e.g., losing decimal places when converting a Double
' to a Long) or you may get an overflow (storing a Long in an Integer) which will result in that
' element in DestinationArray having a value of 0.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim VTypeSource As VbVarType
Dim VTypeDest As VbVarType
Dim SNdx As Long
Dim DNdx As Long


'''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''
CopyArray = False

''''''''''''''''''''''''''''''''''
' Ensure both DestinationArray and
' SourceArray are arrays.
''''''''''''''''''''''''''''''''''
If IsArray(DestinationArray) = False Then
    Exit Function
End If
If IsArray(SourceArray) = False Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''
' Ensure DestinationArray and
' SourceArray are single-dimensional.
' 0 indicates an unallocated array,
' which is allowed.
'''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=SourceArray) > 1 Then
    Exit Function
End If
If NumberOfArrayDimensions(Arr:=DestinationArray) > 1 Then
    Exit Function
End If

''''''''''''''''''''''''''''''''''''
' If SourceArray is not allocated,
' leave DestinationArray intact and
' return a result of True.
''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(Arr:=SourceArray) = False Then
    CopyArray = True
    Exit Function
End If

If NoCompatabilityCheck = False Then
    ''''''''''''''''''''''''''''''''''''''
    ' Ensure both arrays are the same
    ' type or compatible data types. See
    ' the function AreDataTypesCompatible
    ' for information about compatible
    ' types.
    ''''''''''''''''''''''''''''''''''''''
    If AreDataTypesCompatible(DestVar:=DestinationArray, SourceVar:=SourceArray) = False Then
        CopyArray = False
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''
    ' If one array is an array of
    ' objects, ensure the other contains
    ' all objects (or Nothing)
    ''''''''''''''''''''''''''''''''''''
    If VarType(DestinationArray) - vbArray = vbObject Then
        If IsArrayAllocated(SourceArray) = True Then
            For SNdx = LBound(SourceArray) To UBound(SourceArray)
                If IsObject(SourceArray(SNdx)) = False Then
                    Exit Function
                End If
            Next SNdx
        End If
    End If
End If

If IsArrayAllocated(Arr:=DestinationArray) = True Then
    If IsArrayAllocated(Arr:=SourceArray) = True Then
        '''''''''''''''''''''''''''''''''''''''''''''''''
        ' If both arrays are allocated, copy from
        ' SourceArray to DestinationArray. If
        ' SourceArray is smaller that DesetinationArray,
        ' the right-most elements of DestinationArray
        ' are left unchanged. If SourceArray is larger
        ' than DestinationArray, the right most elements
        ' of SourceArray are not copied.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        DNdx = LBound(DestinationArray)
        On Error Resume Next
        For SNdx = LBound(SourceArray) To UBound(SourceArray)
            If IsObject(SourceArray(SNdx)) = True Then
                Set DestinationArray(DNdx) = SourceArray(DNdx)
            Else
                DestinationArray(DNdx) = SourceArray(DNdx)
            End If
            DNdx = DNdx + 1
            If DNdx > UBound(DestinationArray) Then
                Exit For
            End If
        Next SNdx
        On Error GoTo 0
    Else
        '''''''''''''''''''''''''''''''''''''''''''''''
        ' If SourceArray is not allocated, so we have
        ' nothing to copy. Exit with a result
        ' of True. Leave DestinationArray intact.
        '''''''''''''''''''''''''''''''''''''''''''''''
        CopyArray = True
        Exit Function
    End If
        
Else
    If IsArrayAllocated(Arr:=SourceArray) = True Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' If Destination array is not allocated and
        ' SourceArray is allocated, Redim DestinationArray
        ' to the same size as SourceArray and copy
        ' the elements from SourceArray to DestinationArray.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        On Error Resume Next
        ReDim DestinationArray(LBound(SourceArray) To UBound(SourceArray))
        For SNdx = LBound(SourceArray) To UBound(SourceArray)
            If IsObject(SourceArray(SNdx)) = True Then
                Set DestinationArray(SNdx) = SourceArray(SNdx)
            Else
                DestinationArray(SNdx) = SourceArray(SNdx)
            End If
        Next SNdx
        On Error GoTo 0
    Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' If both SourceArray and DestinationArray are
        ' unallocated, we have nothing to copy (this condition
        ' is actually detected above, but included here
        ' for consistancy), so get out with a result of True.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        CopyArray = True
        Exit Function
    End If
End If

'''''''''''''''''''''''
' Success. Return True.
'''''''''''''''''''''''
CopyArray = True

End Function



Public Function CopyArraySubSetToArray(InputArray As Variant, ResultArray As Variant, _
    FirstElementToCopy As Long, LastElementToCopy As Long, DestinationElement As Long) As Boolean
' Copies elements of InputArray to ResultArray. 
'It takes the elements from FirstElementToCopy to LastElementToCopy (inclusive) from InputArray and
' copies them to ResultArray, starting at DestinationElement. Existing data in
' ResultArray will be overwrittten. If ResultArray is a dynamic array, it will
' be resized if needed. If ResultArray is a static array and it is not large
' enough to copy all the elements, no elements are copied and the function
' returns False. ' RequiredLevel:4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
Dim SrcNdx As Long
Dim DestNdx As Long
Dim NumElementsToCopy As Long

'''''''''''''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''''''''''''
CopyArraySubSetToArray = False

''''''''''''''''''''''''''''''''''''''''''
' Ensure InputArray and ResultArray are
' arrays.
''''''''''''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    Exit Function
End If
If IsArray(ResultArray) = False Then
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''
' Ensure InputArray is single dimensional.
'''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=InputArray) <> 1 Then
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''
' Ensure ResultArray is unallocated or
' single dimensional.
'''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=ResultArray) > 1 Then
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''
' Ensure the bounds and indexes are valid.
''''''''''''''''''''''''''''''''''''''''''''
If FirstElementToCopy < LBound(InputArray) Then
    Exit Function
End If
If LastElementToCopy > UBound(InputArray) Then
   Exit Function
End If
If FirstElementToCopy > LastElementToCopy Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''
' Calc the number of elements we'll copy
' from InputArray to ResultArray.
'''''''''''''''''''''''''''''''''''''''''
NumElementsToCopy = LastElementToCopy - FirstElementToCopy + 1

If IsArrayDynamic(Arr:=ResultArray) = False Then
    If (DestinationElement + NumElementsToCopy - 1) > UBound(ResultArray) Then
        '''''''''''''''''''''''''''''''''''''''''''''
        ' ResultArray is static and can't be resized.
        ' There is not enough room in the array to
        ' copy all the data.
        '''''''''''''''''''''''''''''''''''''''''''''
        Exit Function
    End If
Else
    ''''''''''''''''''''''''''''''''''''''''''''
    ' ResultArray is dynamic and can be resized.
    ' Test whether we need to resize the array,
    ' and resize it if required.
    '''''''''''''''''''''''''''''''''''''''''''''
    If IsArrayEmpty(Arr:=ResultArray) = True Then
        '''''''''''''''''''''''''''''''''''''''
        ' ResultArray is unallocated. Resize it
        ' to DestinationElement + NumElementsToCopy - 1.
        ' This provides empty elements to the left
        ' of the DestinationElement and room to
        ' copy NumElementsToCopy.
        '''''''''''''''''''''''''''''''''''''''''
        ReDim ResultArray(1 To DestinationElement + NumElementsToCopy - 1)
    Else
        '''''''''''''''''''''''''''''''''''''''''''''''''
        ' ResultArray is allocated. If there isn't room
        ' enough in ResultArray to hold NumElementsToCopy
        ' starting at DestinationElement, we need to
        ' resize the array.
        '''''''''''''''''''''''''''''''''''''''''''''''''
        If (DestinationElement + NumElementsToCopy - 1) > UBound(ResultArray) Then
            If DestinationElement + NumElementsToCopy > UBound(ResultArray) Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Resize the ResultArray.
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If NumElementsToCopy + DestinationElement > UBound(ResultArray) Then
                    ReDim Preserve ResultArray(LBound(ResultArray) To UBound(ResultArray) + DestinationElement - 1)
                Else
                    ReDim Preserve ResultArray(LBound(ResultArray) To UBound(ResultArray) + NumElementsToCopy)
                End If
            Else
                ''''''''''''''''''''''''''''''''''''''''''''
                ' Resize the array to hold NumElementsToCopy
                ' starting at DestinationElement.
                ''''''''''''''''''''''''''''''''''''''''''''
                ReDim Preserve ResultArray(LBound(ResultArray) To UBound(ResultArray) + NumElementsToCopy - DestinationElement + 2)
            End If
        Else
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' The ResultArray is large enough to hold
            ' NumberOfElementToCopy starting at DestinationElement.
            ' No need to resize the array.
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
    End If
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''
' Copy the elements from InputArray to ResultArray
' Note that there is no type compatibility checking
' when copying the elements.
'''''''''''''''''''''''''''''''''''''''''''''''''''
DestNdx = DestinationElement
For SrcNdx = FirstElementToCopy To LastElementToCopy
    If IsObject(InputArray(SrcNdx)) = True Then
        Set ResultArray(DestNdx) = InputArray(DestNdx)
    Else
        On Error Resume Next
        ResultArray(DestNdx) = InputArray(SrcNdx)
        On Error GoTo 0
    End If
    DestNdx = DestNdx + 1
Next SrcNdx
    
CopyArraySubSetToArray = True
    
End Function

Public Function CopyNonNothingObjectsToArray(ByRef SourceArray As Variant, _
    ByRef ResultArray As Variant, Optional NoAlerts As Boolean = False) As Boolean
'Copies all objects that are not Nothing from SourceArray to ResultArray. 
'ResultArray MUST be a dynamic array of type Object or Variant.
' E.g.,
'       Dim ResultArray() As Object ' Or
'       Dim ResultArray() as Variant
'
' ResultArray will be Erased and then resized to hold the non-Nothing elements
' from SourceArray. The LBound of ResultArray will be the same as the LBound
' of SourceArray, regardless of what its LBound was prior to calling this
' procedure. ' RequiredLevel:4
'
' This function returns True if the operation was successful or False if an
' an error occurs. If an error occurs, a message box is displayed indicating
' the error. To suppress the message boxes, set the NoAlerts parameter to
' True.
'
' This function uses the following procedures. They are declared as Private
' procedures at the end of this module.
'       IsArrayDynamic
'       IsArrayEmpty
'       NumberOfArrayDimensions
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ResNdx As Long
Dim InNdx  As Long

'''''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''''
CopyNonNothingObjectsToArray = False

'''''''''''''''''''''''''''''''''''
' Ensure SourceArray is an array.
'''''''''''''''''''''''''''''''''''
If IsArray(SourceArray) = False Then
    If NoAlerts = False Then
        MsgBox "SourceArray is not an array."
    End If
    Exit Function
End If
'''''''''''''''''''''''''''''''''''
' Ensure SourceArray is a single
' dimensional array.
'''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(Arr:=SourceArray)
    Case 0
        '''''''''''''''''''''''''''''
        ' Unallocated dynamic array.
        ' Not Allowed.
        '''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "SourceArray is an unallocated array."
        End If
        Exit Function
        
    Case 1
        '''''''''''''''''''''''''''''
        ' Single-dimensional array.
        ' This is OK.
        '''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''
        ' Multi-dimensional array.
        ' This is not allowed.
        '''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "SourceArray is a multi-dimensional array. This is not allowed."
        End If
        Exit Function
End Select
'''''''''''''''''''''''''''''''''''
' Ensure ResultArray is an array.
'''''''''''''''''''''''''''''''''''
If IsArray(ResultArray) = False Then
    If NoAlerts = False Then
        MsgBox "ResultArray is not an array."
    End If
    Exit Function
End If
'''''''''''''''''''''''''''''''''''
' Ensure ResultArray is an dynamic.
'''''''''''''''''''''''''''''''''''
If IsArrayDynamic(Arr:=ResultArray) = False Then
    If NoAlerts = False Then
        MsgBox "ResultArray is not a dynamic array."
    End If
    Exit Function
End If
'''''''''''''''''''''''''''''''''''
' Ensure ResultArray is a single
' dimensional array.
'''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(Arr:=ResultArray)
    Case 0
        '''''''''''''''''''''''''''''
        ' Unallocated dynamic array.
        ' This is OK.
        '''''''''''''''''''''''''''''
    Case 1
        '''''''''''''''''''''''''''''
        ' Single-dimensional array.
        ' This is OK.
        '''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''
        ' Multi-dimensional array.
        ' This is not allowed.
        '''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "SourceArray is a multi-dimensional array. This is not allowed."
        End If
        Exit Function
End Select

'''''''''''''''''''''''''''''''''
' Ensure that all the elements of
' SourceArray are in fact objects.
'''''''''''''''''''''''''''''''''
For InNdx = LBound(SourceArray) To UBound(SourceArray)
    If IsObject(SourceArray(InNdx)) = False Then
        If NoAlerts = False Then
            MsgBox "Element " & CStr(InNdx) & " of SourceArray is not an object."
        End If
        Exit Function
    End If
Next InNdx

''''''''''''''''''''''''''''''
' Erase the ResultArray. Since
' ResultArray is dynamic, this
' will relase the memory used
' by ResultArray and return
' the array to an unallocated
' state.
''''''''''''''''''''''''''''''
Erase ResultArray
''''''''''''''''''''''''''''''
' Now, size ResultArray to the
' size of SourceArray. After
' moving all the non-Nothing
' elements, we'll do another
' resize to get ResultArray
' to the used size. This method
' allows us to avoid Redim
' Preserve for every element.
'''''''''''''''''''''''''''''
ReDim ResultArray(LBound(SourceArray) To UBound(SourceArray))

ResNdx = LBound(SourceArray)
For InNdx = LBound(SourceArray) To UBound(SourceArray)
    If Not SourceArray(InNdx) Is Nothing Then
        Set ResultArray(ResNdx) = SourceArray(InNdx)
        ResNdx = ResNdx + 1
    End If
Next InNdx
''''''''''''''''''''''''''''''''''''''''''
' Now that we've copied all the
' non-Nothing elements from SourceArray
' to ResultArray, we call Redim Preserve
' to resize the ResultArray to the size
' actually used. Test ResNdx to see
' if we actually copied any elements.
''''''''''''''''''''''''''''''''''''''''''
If ResNdx > LBound(SourceArray) Then
    '''''''''''''''''''''''''''''''''''''''
    ' If ResNdx > LBound(SourceArray) then
    ' we copied at least one element out of
    ' SourceArray.
    '''''''''''''''''''''''''''''''''''''''
    ReDim Preserve ResultArray(LBound(ResultArray) To ResNdx - 1)
Else
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Otherwise, we didn't copy any elements
    ' from SourceArray (all elements in SourceArray
    ' were Nothing). In this case, Erase ResultArray.
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Erase ResultArray
End If
'''''''''''''''''''''''''''''
' No errors were encountered.
' Return True.
'''''''''''''''''''''''''''''
CopyNonNothingObjectsToArray = True


End Function



