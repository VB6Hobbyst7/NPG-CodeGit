Attribute VB_Name = "modWorkbookOpenProcedures"
Public Function WorkbookIsOpen(wbname As String) As Boolean
'Test if there is an open workbook with certain name
'http://codevba.com/excel/test_if_open_workbook.htm
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(wbname)
    If Err = 0 Then
        WorkbookIsOpen = True
    Else
        WorkbookIsOpen = False
    End If
End Function

