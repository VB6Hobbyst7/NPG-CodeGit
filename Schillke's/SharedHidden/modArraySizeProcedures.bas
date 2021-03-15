Attribute VB_Name = "modArraySizeProcedures"
Public Function NumberOfArrayDimensions(Arr As Variant) As Integer
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
 



Public Function NumElements(Arr As Variant, Optional Dimension = 1) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumElements
' Returns the number of elements in the specified dimension (Dimension) of the array in
' Arr. If you omit Dimension, the first dimension is used. The function will return
' 0 under the following circumstances:
'     Arr is not an array, or
'     Arr is an unallocated array, or
'     Dimension is greater than the number of dimension of Arr, or
'     Dimension is less than 1.
'
' This function does not support arrays of user-defined Type variables.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim NumDimensions As Long

' if Arr is not an array, return 0 and get out.
If IsArray(Arr) = False Then
    NumElements = 0
    Exit Function
End If

' if the array is unallocated, return 0 and get out.
If IsArrayEmpty(Arr) = True Then
    NumElements = 0
    Exit Function
End If

' ensure that Dimension is at least 1.
If Dimension < 1 Then
    NumElements = 0
    Exit Function
End If

' get the number of dimensions
NumDimensions = NumberOfArrayDimensions(Arr)
If NumDimensions < Dimension Then
    NumElements = 0
    Exit Function
End If

' returns the number of elements in the array
NumElements = UBound(Arr, Dimension) - LBound(Arr, Dimension) + 1

End Function

