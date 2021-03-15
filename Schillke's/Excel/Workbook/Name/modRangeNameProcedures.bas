Attribute VB_Name = "modRangeNameProcedures"

Public Function RangeNameExists(strName, Optional WB As Workbook) As Boolean
'Test if the range name exists in active or specified workbook
'http://codevba.com/excel/range.htm#RangeNameExists
Dim N As Name
    If WB Is Nothing Then
        Set WB = ActiveWorkbook
    End If
    RangeNameExists = False
    For Each N In WB.Names
        If UCase(N.Name) = UCase(strName) Then
            RangeNameExists = True
            Exit Function
        End If
    Next N
End Function
