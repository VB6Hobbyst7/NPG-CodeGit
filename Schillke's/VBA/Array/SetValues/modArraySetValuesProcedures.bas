Attribute VB_Name = "modArraySetValuesProcedures"
'<include modArrayIsProcedures.bas><include modArraySizeProcedures.bas>

Public Function ResetVariantArrayToDefaults(InputArray As Variant) As Boolean
' Resets all the elements of an array of Variants back to their appropriate
' default values. The elements of the array may be of mixed types (e.g., some Longs,
' some Objects, some Strings, etc). Each data type will be set to the appropriate
' default value (0, vbNullString, Empty, or Nothing). It returns True if the
' array was set to defautls, or False if an error occurred. InputArray must be
' an allocated single-dimensional array. This function differs from the Erase
' function in that it preserves the original data types, while Erase sets every
' element to Empty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
'''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''
ResetVariantArrayToDefaults = False

'''''''''''''''''''''''''''''''
' Ensure InputArray is an array
'''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    Exit Function
End If

'''''''''''''''''''''''''''''''
' Ensure InputArray is a single
' dimensional allocated array.
'''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=InputArray) <> 1 Then
    Exit Function
End If

For Ndx = LBound(InputArray) To UBound(InputArray)
    SetVariableToDefault InputArray(Ndx)
Next Ndx

ResetVariantArrayToDefaults = True

End Function

 
Public Function SetObjectArrayToNothing(InputArray As Variant) As Boolean
'Sets all the elements of InputArray to Nothing. 
'Use this function rather than Erase because if InputArray is an array of Variants, Erase
' will set each element to Empty, not Nothing, and the element will cease  to be an object.
' The function returns True if successful, False otherwise.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim N As Long

''''''''''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
''''''''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    SetObjectArrayToNothing = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a single-dimensional array.
''''''''''''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=InputArray) <> 1 Then
    SetObjectArrayToNothing = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure the array is allocated and that each
' element is an object (or Nothing). If the
' array is not allocated, return True.
' We do this test before setting any element
' to Nothing so we don't end up with an array
' that is a mix of Empty and Nothing values.
' This means looping through the array twice,
' but it ensures all or none of the elements
' get set to Nothing.
''''''''''''''''''''''''''''''''''''''''''''''''
If IsArrayAllocated(Arr:=InputArray) = True Then
    For N = LBound(InputArray) To UBound(InputArray)
        If IsObject(InputArray(N)) = False Then
            SetObjectArrayToNothing = False
            Exit Function
        End If
    Next N
Else
    SetObjectArrayToNothing = True
    Exit Function
End If


'''''''''''''''''''''''''''''''''''''''''''''
' Set each element of InputArray to Nothing.
'''''''''''''''''''''''''''''''''''''''''''''
For N = LBound(InputArray) To UBound(InputArray)
    Set InputArray(N) = Nothing
Next N

SetObjectArrayToNothing = True

End Function


Public Sub SetVariableToDefault(ByRef Variable As Variant)
'Sets Variable to the appropriate default value for its data type. 
'Note that it cannot change User-Defined
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
' Types.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsObject(Variable) Then
    '''''''''''''''''''''''''''''''''''''''
    ' We test with IsObject here so that
    ' the object itself, not the default
    ' property of the object, is evaluated.
    ''''''''''''''''''''''''''''''''''''''''
    Set Variable = Nothing
Else
    Select Case VarType(Variable)
        Case Is >= vbArray
            ''''''''''''''''''''''''''''''''''''''''''''
            ' The VarType of an array is
            ' equal to vbArray + VarType(ArrayElement).
            ' Here we check for anything >= vbArray
            ''''''''''''''''''''''''''''''''''''''''''''
            Erase Variable
        Case vbBoolean
            Variable = False
        Case vbByte
            Variable = CByte(0)
        Case vbCurrency
            Variable = CCur(0)
        Case vbDataObject
            Set Variable = Nothing
        Case vbDate
            Variable = CDate(0)
        Case vbDecimal
            Variable = CDec(0)
        Case vbDouble
            Variable = CDbl(0)
        Case vbEmpty
            Variable = Empty
        Case vbError
            Variable = Empty
        Case vbInteger
            Variable = CInt(0)
        Case vbLong
            Variable = CLng(0)
        Case vbNull
            Variable = Empty
        Case vbObject
            Set Variable = Nothing
        Case vbSingle
            Variable = CSng(0)
        Case vbString
            Variable = vbNullString
        Case vbUserDefinedType
            '''''''''''''''''''''''''''''''''
            ' User-Defined-Types cannot be
            ' set to a general default value.
            ' Each element must be explicitly
            ' set to its default value. No
            ' assignment takes place in this
            ' procedure.
            ''''''''''''''''''''''''''''''''''
        Case vbVariant
            ''''''''''''''''''''''''''''''''''''''''''''''''
            ' This case is included for constistancy,
            ' but we will never get here. If the Variant
            ' contains data, VarType returns the type of
            ' that data. An Empty Variant is type vbEmpty.
            ''''''''''''''''''''''''''''''''''''''''''''''''
            Variable = Empty
    End Select
End If

End Sub

