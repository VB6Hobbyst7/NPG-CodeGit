Attribute VB_Name = "modArrayDataTypeProcedures"
Public Function DataTypeOfArray(Arr As Variant) As VbVarType
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DataTypeOfArray
'
' Returns a VbVarType value indicating data type of the elements of
' Arr.
'
' The VarType of an array is the value vbArray plus the VbVarType value of the
' data type of the array. For example the VarType of an array of Longs is 8195,
' which equal to vbArray + vbLong. This code subtracts the value of vbArray to
' return the native data type.
'
' If Arr is a simple array, either single- or mulit-
' dimensional, the function returns the data type of the array. Arr
' may be an unallocated array. We can still get the data type of an unallocated
' array.
'
' If Arr is an array of arrays, the function returns vbArray. To retrieve
' the data type of a subarray, pass into the function one of the sub-arrays. E.g.,
' Dim R As VbVarType
' R = DataTypeOfArray(A(LBound(A)))
'
' This function support single and multidimensional arrays. It does not
' support user-defined types. If Arr is an array of empty variants (vbEmpty)
' it returns vbVariant
'
' Returns -1 if Arr is not an array.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Element As Variant
Dim NumDimensions As Long

' If Arr is not an array, return
' vbEmpty and get out.
If IsArray(Arr) = False Then
    DataTypeOfArray = -1
    Exit Function
End If

If IsArrayEmpty(Arr) = True Then
    ' If the array is unallocated, we can still get its data type.
    ' The result of VarType of an array is vbArray + the VarType
    ' of elements of the array (e.g., the VarType of an array of Longs
    ' is 8195, which is vbArray + vbLong). Thus, to get the basic data
    ' type of the array, we subtract the value vbArray.
    DataTypeOfArray = VarType(Arr) - vbArray
Else
    ' get the number of dimensions in the array.
    NumDimensions = NumberOfArrayDimensions(Arr)
    ' set variable Element to first element of the first dimension
    ' of the array
    If NumDimensions = 1 Then
        If IsObject(Arr(LBound(Arr))) = True Then
            DataTypeOfArray = vbObject
            Exit Function
        End If
        Element = Arr(LBound(Arr))
    Else
        If IsObject(Arr(LBound(Arr), 1)) = True Then
            DataTypeOfArray = vbObject
            Exit Function
        End If
        Element = Arr(LBound(Arr), 1)
    End If
    ' if we were passed an array of arrays, IsArray(Element) will
    ' be true. Therefore, return vbArray. If IsArray(Element) is false,
    ' we weren't passed an array of arrays, so simply return the data type of
    ' Element.
    If IsArray(Element) = True Then
        DataTypeOfArray = vbArray
    Else
        If VarType(Element) = vbEmpty Then
            DataTypeOfArray = vbVariant
        Else
            DataTypeOfArray = VarType(Element)
        End If
    End If
End If

End Function


Public Function AreDataTypesCompatible(DestVar As Variant, SourceVar As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' AreDataTypesCompatible
' This function determines if SourceVar is compatiable with DestVar. If the two
' data types are the same, they are compatible. If the value of SourceVar can
' be stored in DestVar with no loss of precision or an overflow, they are compatible.
' For example, if DestVar is a Long and SourceVar is an Integer, they are compatible
' because an integer can be stored in a Long with no loss of information. If DestVar
' is a Long and SourceVar is a Double, they are not compatible because information
' will be lost converting from a Double to a Long (the decimal portion will be lost).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim SVType As VbVarType
Dim DVType As VbVarType

'''''''''''''''''''''''''''''''''''
' Set the default return type.
'''''''''''''''''''''''''''''''''''
AreDataTypesCompatible = False

'''''''''''''''''''''''''''''''''''
' If DestVar is an array, get the
' type of array. If it is an array
' its VarType is vbArray + VarType(element)
' so we subtract vbArray to get then
' data type of the aray. E.g.,
' the VarType of an array of Longs
' is 8195 = vbArray + vbLong,
' 8195 - vbArray = vbLong (=3).
'''''''''''''''''''''''''''''''''''
If IsArray(DestVar) = True Then
    DVType = VarType(DestVar) - vbArray
Else
    DVType = VarType(DestVar)
End If
'''''''''''''''''''''''''''''''''''
' If SourceVar is an array, get the
' type of array.
'''''''''''''''''''''''''''''''''''
If IsArray(SourceVar) = True Then
    SVType = VarType(SourceVar) - vbArray
Else
    SVType = VarType(SourceVar)
End If

''''''''''''''''''''''''''''''''''''
' If one variable is an array and
' the other is not an array, they
' are incompatible.
''''''''''''''''''''''''''''''''''''
If ((IsArray(DestVar) = True) And (IsArray(SourceVar) = False) Or _
    (IsArray(DestVar) = False) And (IsArray(SourceVar) = True)) Then
    Exit Function
End If


''''''''''''''''''''''''''''''''''''
' Test the data type of DestVar
' and return a result if SourceVar
' is compatible with that type.
''''''''''''''''''''''''''''''''''''
If SVType = DVType Then
    '''''''''''''''''''''''''''''''''
    ' The the variable types are the
    ' same, they are compatible.
    ''''''''''''''''''''''''''''''''
    AreDataTypesCompatible = True
    Exit Function
Else
    '''''''''''''''''''''''''''''''''''''''''
    ' If the data types are not the same,
    ' determine whether they are compatible.
    '''''''''''''''''''''''''''''''''''''''''
    Select Case DVType
        Case vbInteger
            Select Case SVType
                Case vbInteger
                    AreDataTypesCompatible = True
                    Exit Function
                Case Else
                    AreDataTypesCompatible = False
                    Exit Function
            End Select
        
        Case vbLong
            Select Case SVType
                Case vbInteger, vbLong
                    AreDataTypesCompatible = True
                    Exit Function
                Case Else
                    AreDataTypesCompatible = False
                    Exit Function
            End Select
        Case vbSingle
            Select Case SVType
                Case vbInteger, vbLong, vbSingle
                    AreDataTypesCompatible = True
                    Exit Function
                Case Else
                    AreDataTypesCompatible = False
                    Exit Function
            End Select
        Case vbDouble
            Select Case SVType
                Case vbInteger, vbLong, vbSingle, vbDouble
                    AreDataTypesCompatible = True
                    Exit Function
                Case Else
                    AreDataTypesCompatible = False
                    Exit Function
            End Select
        Case vbString
            Select Case SVType
                Case vbString
                    AreDataTypesCompatible = True
                    Exit Function
                Case Else
                    AreDataTypesCompatible = False
                    Exit Function
            End Select
        Case vbObject
            Select Case SVType
                Case vbObject
                    AreDataTypesCompatible = True
                    Exit Function
                Case Else
                    AreDataTypesCompatible = False
                    Exit Function
            End Select
        Case vbBoolean
            Select Case SVType
                Case vbBoolean, vbInteger
                    AreDataTypesCompatible = True
                    Exit Function
                Case Else
                    AreDataTypesCompatible = False
                    Exit Function
            End Select
         Case vbByte
            Select Case SVType
                Case vbByte
                    AreDataTypesCompatible = True
                    Exit Function
                Case Else
                    AreDataTypesCompatible = False
                    Exit Function
            End Select
        Case vbCurrency
            Select Case SVType
                Case vbInteger, vbLong, vbSingle, vbDouble
                    AreDataTypesCompatible = True
                    Exit Function
                Case Else
                    AreDataTypesCompatible = False
                    Exit Function
            End Select
        Case vbDecimal
            Select Case SVType
                Case vbInteger, vbLong, vbSingle, vbDouble
                    AreDataTypesCompatible = True
                    Exit Function
                Case Else
                    AreDataTypesCompatible = False
                    Exit Function
            End Select
        Case vbDate
            Select Case SVType
                Case vbLong, vbSingle, vbDouble
                    AreDataTypesCompatible = True
                    Exit Function
                Case Else
                    AreDataTypesCompatible = False
                    Exit Function
            End Select
        
         Case vbEmpty
            Select Case SVType
                Case vbVariant
                    AreDataTypesCompatible = True
                    Exit Function
                Case Else
                    AreDataTypesCompatible = False
                    Exit Function
            End Select
         Case vbError
            AreDataTypesCompatible = False
            Exit Function
         Case vbNull
            AreDataTypesCompatible = False
            Exit Function
         Case vbObject
            Select Case SVType
                Case vbObject
                    AreDataTypesCompatible = True
                    Exit Function
                Case Else
                    AreDataTypesCompatible = False
                    Exit Function
            End Select
         Case vbVariant
            AreDataTypesCompatible = True
            Exit Function
        
    End Select
End If


End Function


