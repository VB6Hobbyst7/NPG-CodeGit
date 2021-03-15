Attribute VB_Name = "modArraySortProcedures"
'<include modArrayIsProcedures.bas><include modArraySizeProcedures.bas>

Public Function FirstNonEmptyStringIndexInArray(InputArray As Variant) As Long
'Returns the index into InputArray of the first non-empty string.
' This is generally used when InputArray is the result of a sort operation,
' which puts empty strings at the beginning of the array.
' Returns -1 is an error occurred or if the entire array is empty strings.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long

If IsArray(InputArray) = False Then
    FirstNonEmptyStringIndexInArray = -1
    Exit Function
End If
   
Select Case NumberOfArrayDimensions(Arr:=InputArray)
    Case 0
        '''''''''''''''''''''''''''''''''''''''''
        ' indicates an unallocated dynamic array.
        '''''''''''''''''''''''''''''''''''''''''
        FirstNonEmptyStringIndexInArray = -1
        Exit Function
    Case 1
        '''''''''''''''''''''''''''''''''''''''''
        ' single dimensional array. OK.
        '''''''''''''''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''''''''''''''
        ' multidimensional array. Invalid.
        '''''''''''''''''''''''''''''''''''''''''
        FirstNonEmptyStringIndexInArray = -1
        Exit Function
End Select

For Ndx = LBound(InputArray) To UBound(InputArray)
    If InputArray(Ndx) <> vbNullString Then
        FirstNonEmptyStringIndexInArray = Ndx
        Exit Function
    End If
Next Ndx

FirstNonEmptyStringIndexInArray = -1
End Function

Public Function MoveEmptyStringsToEndOfArray(InputArray As Variant) As Boolean
' This procedure takes the SORTED array InputArray, which, if sorted in
' ascending order, will have all empty strings at the front of the array.
' This procedure moves those strings to the end of the array, shifting
' the non-empty strings forward in the array.
' Note that InputArray MUST be sorted in ascending order.
' Returns True if the array was correctly shifted (if necessary) and False
' if an error occurred.
' This function uses the following functions, which are included as Private
' procedures at the end of this module.
'       FirstNonEmptyStringIndexInArray
'       NumberOfArrayDimensions
'       IsArrayAllocated

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Temp As String
Dim Ndx As Long
Dim Ndx2 As Long
Dim NonEmptyNdx As Long
Dim FirstNonEmptyNdx As Long


''''''''''''''''''''''''''''''''
' Ensure InpuyArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    MoveEmptyStringsToEndOfArray = False
    Exit Function
End If


''''''''''''''''''''''''''''''''''''
' Ensure that the array is allocated
' (not an empty array).
''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(Arr:=InputArray) = False Then
    MoveEmptyStringsToEndOfArray = False
    Exit Function
End If


FirstNonEmptyNdx = FirstNonEmptyStringIndexInArray(InputArray:=InputArray)
If FirstNonEmptyNdx <= LBound(InputArray) Then
    ''''''''''''''''''''''''''''''''''''''''''
    ' No empty strings at the beginning of the
    ' array. Get out now.
    ''''''''''''''''''''''''''''''''''''''''''
    MoveEmptyStringsToEndOfArray = True
    Exit Function
End If


''''''''''''''''''''''''''''''''''''''''''''''''
' Loop through the array, swapping vbNullStrings
' at the beginning with values at the end.
''''''''''''''''''''''''''''''''''''''''''''''''
NonEmptyNdx = FirstNonEmptyNdx
For Ndx = LBound(InputArray) To UBound(InputArray)
    If InputArray(Ndx) = vbNullString Then
        InputArray(Ndx) = InputArray(NonEmptyNdx)
        InputArray(NonEmptyNdx) = vbNullString
        NonEmptyNdx = NonEmptyNdx + 1
        If NonEmptyNdx > UBound(InputArray) Then
            Exit For
        End If
    End If
Next Ndx
''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set entires (Ndx+1) to UBound(InputArray) to
' vbNullStrings.
''''''''''''''''''''''''''''''''''''''''''''''''''''
For Ndx2 = Ndx + 1 To UBound(InputArray)
    InputArray(Ndx2) = vbNullString
Next Ndx2
MoveEmptyStringsToEndOfArray = True

End Function

