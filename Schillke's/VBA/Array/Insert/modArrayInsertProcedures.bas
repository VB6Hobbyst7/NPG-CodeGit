Attribute VB_Name = "modArrayInsertProcedures"
'<include modArrayIsProcedures.bas><include modArraySizeProcedures.bas><include modArraySortProcedures.bas>
Option Explicit
Option Compare Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modArraySupport
' By Chip Pearson, chip@cpearson.com, www.cpearson.com
'
' This module contains procedures that provide information about and manipulate
' VB/VBA arrays. NOTE: These functions call one another. It is strongly suggested
' that you Import this entire module to a VBProject rather then copy/pasting
' individual procedures.
'
' For details on these functions, see www.cpearson.com/excel/VBAArrays.htm
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function InsertElementIntoArray(InputArray As Variant, Index As Long, _
    Value As Variant) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Inserts an element with a value of Value into InputArray at locatation Index.
' InputArray must be a dynamic array. The Value is stored in location Index, and everything
' to the right of Index is shifted to the right. The array is resized to make room for
' the new element. The value of Index must be greater than or equal to the LBound of
' InputArray and less than or equal to UBound+1. If Index is UBound+1, the Value is
' placed at the end of the array.
'http://www.cpearson.com/excel/vbaarrays.htm RequiredLevel:4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long

'''''''''''''''''''''''''''''''
' Set the default return value.
'''''''''''''''''''''''''''''''
InsertElementIntoArray = False

''''''''''''''''''''''''''''''''
' Ensure InputArray is an array.
''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    Exit Function
End If

''''''''''''''''''''''''''''''''
' Ensure InputArray is dynamic.
''''''''''''''''''''''''''''''''
If IsArrayDynamic(Arr:=InputArray) = False Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''
' Ensure InputArray is allocated.
'''''''''''''''''''''''''''''''''
If IsArrayAllocated(Arr:=InputArray) = False Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''
' Ensure InputArray is a single
' dimensional array.
'''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=InputArray) <> 1 Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''
' Ensure Index is a valid element index.
' We allow Index to be equal to
' UBound + 1 to facilitate inserting
' a value at the end of the array. E.g.,
' InsertElementIntoArray(Arr,UBound(Arr)+1,123)
' will insert 123 at the end of the array.
'''''''''''''''''''''''''''''''''''''''''
If (Index < LBound(InputArray)) Or (Index > UBound(InputArray) + 1) Then
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''
' Resize the array
'''''''''''''''''''''''''''''''''''''''''''''
ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray) + 1)
'''''''''''''''''''''''''''''''''''''''''''''
' First, we set the newly created last element
' of InputArray to Value. This is done to trap
' an error 13, type mismatch. This last entry
' will be overwritten when we shift elements
' to the right, and the Value will be inserted
' at Index.
'''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
Err.Clear
InputArray(UBound(InputArray)) = Value
If Err.Number <> 0 Then
    ''''''''''''''''''''''''''''''''''''''
    ' An error occurred, most likely
    ' an error 13, type mismatch.
    ' Redim the array back to its original
    ' size and exit the function.
    '''''''''''''''''''''''''''''''''''''''
    ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray) - 1)
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''''''''''
' Shift everything to the right.
'''''''''''''''''''''''''''''''''''''''''''''
For Ndx = UBound(InputArray) To Index + 1 Step -1
    InputArray(Ndx) = InputArray(Ndx - 1)
Next Ndx

'''''''''''''''''''''''''''''''''''''''''''''
' Insert Value at Index
'''''''''''''''''''''''''''''''''''''''''''''
InputArray(Index) = Value
   
InsertElementIntoArray = True

End Function

