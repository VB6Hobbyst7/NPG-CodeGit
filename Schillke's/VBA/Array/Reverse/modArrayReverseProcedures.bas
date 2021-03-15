Attribute VB_Name = "modArrayReverseProcedures"
'<include modArrayIsProcedures.bas><include modArraySizeProcedures.bas>

Public Function ReverseArrayInPlace(InputArray As Variant, _
    Optional NoAlerts As Boolean = False) As Boolean
'Rreverses the order of an array in place -- this is, the array variable in the calling procedure is reversed. 
'This works only on single-dimensional arrays
' of simple data types (String, Single, Double, Integer, Long). It will not work
' on arrays of objects. Use ReverseArrayOfObjectsInPlace to reverse an array of objects.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As Variant
Dim Ndx As Long
Dim Ndx2 As Long


'''''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''''
ReverseArrayInPlace = False

'''''''''''''''''''''''''''''''''
' ensure we have an array
'''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    If NoAlerts = False Then
        MsgBox "The InputArray parameter is not an array."
    End If
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Test the number of dimensions of the
' InputArray. If 0, we have an empty,
' unallocated array. Get out with
' an error message. If greater than
' one, we have a multi-dimensional
' array, which is not allowed. Only
' an allocated 1-dimensional array is
' allowed.
''''''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(InputArray)
    Case 0
        If NoAlerts = False Then
            MsgBox "The input array is an empty, unallocated array."
        End If
        Exit Function
    Case 1
        ' ok
    Case Else
        If NoAlerts = False Then
            MsgBox "The input array is multi-dimensional. ReverseArrayInPlace works only " & _
                   "on single-dimensional arrays."
        End If
        Exit Function
End Select

Ndx2 = UBound(InputArray)
''''''''''''''''''''''''''''''''''''''
' loop from the LBound of InputArray to
' the midpoint of InputArray
''''''''''''''''''''''''''''''''''''''
For Ndx = LBound(InputArray) To ((UBound(InputArray) - LBound(InputArray) + 1) \ 2)
    'swap the elements
    Temp = InputArray(Ndx)
    InputArray(Ndx) = InputArray(Ndx2)
    InputArray(Ndx2) = Temp
    ' decrement the upper index
    Ndx2 = Ndx2 - 1
Next Ndx

''''''''''''''''''''''''''''''''''''''
' OK - Return True
''''''''''''''''''''''''''''''''''''''
ReverseArrayInPlace = True

End Function



Public Function ReverseArrayOfObjectsInPlace(InputArray As Variant, _
    Optional NoAlerts As Boolean = False) As Boolean
'Reverses the order of an array in place -- this is, the array variable in the calling procedure is reversed. 
'This works only with arrays of objects. It does not work on simple variables. 
'Use ReverseArrayInPlace for simple variables. An error
' will occur if an element of the array is not an object.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As Variant
Dim Ndx As Long
Dim Ndx2 As Long


'''''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''''
ReverseArrayOfObjectsInPlace = False

'''''''''''''''''''''''''''''''''
' ensure we have an array
'''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    If NoAlerts = False Then
        MsgBox "The InputArray parameter is not an array."
    End If
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Test the number of dimensions of the
' InputArray. If 0, we have an empty,
' unallocated array. Get out with
' an error message. If greater than
' one, we have a multi-dimensional
' array, which is not allowed. Only
' an allocated 1-dimensional array is
' allowed.
''''''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(InputArray)
    Case 0
        If NoAlerts = False Then
            MsgBox "The input array is an empty, unallocated array."
        End If
        Exit Function
    Case 1
        ' ok
    Case Else
        If NoAlerts = False Then
            MsgBox "The input array is multi-dimensional. ReverseArrayInPlace works only " & _
                   "on single-dimensional arrays."
        End If
        Exit Function
End Select

Ndx2 = UBound(InputArray)

'''''''''''''''''''''''''''''''''''''
' ensure the entire array consists
' of objects (Nothing objects are
' allowed).
'''''''''''''''''''''''''''''''''''''
For Ndx = LBound(InputArray) To UBound(InputArray)
    If IsObject(InputArray(Ndx)) = False Then
        If NoAlerts = False Then
            MsgBox "Array item " & CStr(Ndx) & " is not an object."
        End If
        Exit Function
    End If
Next Ndx

''''''''''''''''''''''''''''''''''''''
' loop from the LBound of InputArray to
' the midpoint of InputArray
''''''''''''''''''''''''''''''''''''''
For Ndx = LBound(InputArray) To ((UBound(InputArray) - LBound(InputArray) + 1) \ 2)
    Set Temp = InputArray(Ndx)
    Set InputArray(Ndx) = InputArray(Ndx2)
    Set InputArray(Ndx2) = Temp
    ' decrement the upper index
    Ndx2 = Ndx2 - 1
Next Ndx

''''''''''''''''''''''''''''''''''''''
' OK - Return True
''''''''''''''''''''''''''''''''''''''
ReverseArrayOfObjectsInPlace = True

End Function


