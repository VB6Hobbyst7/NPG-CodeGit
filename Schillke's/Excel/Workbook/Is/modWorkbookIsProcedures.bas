Attribute VB_Name = "modWorkbookIsProcedures"
Function IsOpenWorkbook(strWorkbookName As String, Optional bFullname As Boolean = False) As Boolean
'Check if there is already a file open with the same name
'http://codevba.com/excel/test_if_open_workbook.htm
'Note: this function works correctly independent of Windows 'Show Extension' View setting
'Testcases:
'debug.Print isOpenWorkbook(thisworkbook.Name)          = True '"MySavedFile.xls" - saved file
'debug.Print isOpenWorkbook(thisworkbook.Name,True)     = False '"MySavedFile.xls" - saved file
'debug.Print isOpenWorkbook(thisworkbook.FullName,False)= False '"..\MySavedFile.xls" - saved file
'debug.Print isOpenWorkbook(thisworkbook.FullName,True) = True '"..\MySavedFile.xls" - saved file
'debug.Print isOpenWorkbook("MySavedFile")              = False
'debug.Print isOpenWorkbook("book1")                    = True
'debug.Print isOpenWorkbook("book1.xls")                = False
Dim wb As Workbook
Dim strName As String
    IsOpenWorkbook = False
    For Each wb In Workbooks
        If bFullname = False Then
            strName = wb.Name
        Else
            strName = wb.FullName
        End If
        If (StrComp(strName, strWorkbookName, vbTextCompare) = 0) Then
            IsOpenWorkbook = True
            Exit Function
        End If
    Next
End Function
