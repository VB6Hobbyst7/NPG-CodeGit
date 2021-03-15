Attribute VB_Name = "modRangeMethodsFindProcedures"
Option Explicit
Option Compare Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modFoundCells
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
'
' This module contains the GetFoundCells function. GetFoundCells will return
' a range object that contains all the cells which contain specified values. The
' parameters to GetFoundCells are the same as the parameters to the Find method of the
' Range object.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function GetFoundCells(SearchRange As Range, FindWhat As Variant, _
    Optional LookIn As XlFindLookIn = xlValues, Optional LookAt As XlLookAt = xlWhole, _
    Optional SearchOrder As XlSearchOrder = xlByRows, _
    Optional MatchCase As Boolean = False) As Range
'returns a Range object that contains all the cells in SearchRange in which FindWhat was found. 
'The parameters to the function have the same meaning as they do for the
' Find method of the Range object. If no cells were found, the result of this function
' is Nothing. RequiredLevel:4
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim FoundCell As Range
Dim FoundCells As Range
Dim LastCell As Range
Dim FirstAddr As String
With SearchRange
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' In order to have Find search for the FindWhat value
    ' starting at the first cell in the SearchRange, we
    ' have to find the last cell in SearchRange and use
    ' that as the cell after which the Find will search.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set LastCell = .Cells(.Cells.Count)
End With

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Do the initial Find. If we don't find FindWhat in the first Find,
' we won't even go into the code which searches for subsequent
' occurances.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set FoundCell = SearchRange.Find(what:=FindWhat, after:=LastCell, _
    LookIn:=LookIn, LookAt:=LookAt, SearchOrder:=SearchOrder, MatchCase:=MatchCase)
If Not FoundCell Is Nothing Then
    ''''''''''''''''''''''''''''''
    ' Set the FoundCells range
    ' to the first FoundCell.
    ''''''''''''''''''''''''''''''
    Set FoundCells = FoundCell
    ''''''''''''''''''''''''''''
    ' FirstAddr will contain the
    ' address of the first found
    ' cell. We test each FoundCell
    ' to this address to prevent
    ' the Find from looping back
    ' through the range it has
    ' already searched.
    ''''''''''''''''''''''''''''
    FirstAddr = FoundCell.Address
    Do
        ''''''''''''''''''''''''''''''''
        ' Loop calling FindNext until
        ' FoundCell is nothing or
        ' we wrap around the first
        ' found cell (address is in
        ' FirstAddr).
        '''''''''''''''''''''''''''''''
        Set FoundCells = Application.Union(FoundCells, FoundCell)
        Set FoundCell = SearchRange.FindNext(after:=FoundCell)
    Loop Until (FoundCell Is Nothing) Or (FoundCell.Address = FirstAddr)
End If

''''''''''''''''''''
' Return the result.
''''''''''''''''''''
If FoundCells Is Nothing Then
    Set GetFoundCells = Nothing
Else
    Set GetFoundCells = FoundCells
End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modFindAll
' By Chip Peasron, chip@cpearson.com. www.cpearson.com
' 24-October-2007
' This module is described at www.cpearson.com/Excel/FindAll.aspx
' Requires Excel 2000 or later.
'
' This module contains two functions, FindAll and FindAllOnWorksheets that are use
' to find values on a worksheet or multiple worksheets.
'
' FindAll searches a range and returns a range containing the cells in which the
'   searched for text was found. If the string was not found, it returns Nothing.

' FindAllOnWorksheets searches the same range on one or more workshets. It return
'   an array of ranges, each of which is the range on that worksheet in which the
'   value was found. If the value was not found on a worksheet, that worksheet's
'   element in the returned array will be Nothing.
'
' In both functions, the parameters that control the search have the same meaning
' and effect as they do in the Range.Find method.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function FindAll(SearchRange As Range, _
                FindWhat As Variant, _
                Optional LookIn As XlFindLookIn = xlValues, _
                Optional LookAt As XlLookAt = xlWhole, _
                Optional SearchOrder As XlSearchOrder = xlByRows, _
                Optional MatchCase As Boolean = False) As Range
'searches the range specified by SearchRange and returns a Range object that contains all the cells in which FindWhat was found. 
'The search parameters to this function have the same meaning and effect as they do with the
' Range.Find method. If the value was not found, the function return Nothing.RequiredLevel:3
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim FoundCell As Range
Dim FirstFound As Range
Dim LastCell As Range
Dim ResultRange As Range

With SearchRange
    Set LastCell = .Cells(.Cells.Count)
End With
'On Error Resume Next
On Error GoTo 0
Set FoundCell = SearchRange.Find(what:=FindWhat, _
        after:=LastCell, _
        LookIn:=LookIn, _
        LookAt:=LookAt, _
        SearchOrder:=SearchOrder, _
        MatchCase:=MatchCase)

If Not FoundCell Is Nothing Then
    Set FirstFound = FoundCell
    Set ResultRange = FoundCell
    Set FoundCell = SearchRange.FindNext(after:=FoundCell)
    Do Until False ' Loop forever. We'll "Exit Do" when necessary.
        If (FoundCell Is Nothing) Then
            Exit Do
        End If
        If (FoundCell.Address = FirstFound.Address) Then
            Exit Do
        End If
        Set ResultRange = Application.Union(ResultRange, FoundCell)
        Set FoundCell = SearchRange.FindNext(after:=FoundCell)
    Loop
End If
    
Set FindAll = ResultRange

End Function

Public Function FindAllOnWorksheets(InWorkbook As Workbook, _
                InWorksheets As Variant, _
                SearchAddress As String, _
                FindWhat As Variant, _
                Optional LookIn As XlFindLookIn = xlValues, _
                Optional LookAt As XlLookAt = xlWhole, _
                Optional SearchOrder As XlSearchOrder, _
                Optional MatchCase As Boolean = False) As Variant
'Searches a range on one or more worksheets, in the range specified by SearchAddress.
'RequiredLevel:4
' InWorkbook specifies the workbook in which to search. If this is Nothing, the active
'   workbook is used.
'
' InWorksheets specifies what worksheets to search. InWorksheets can be any of the
' following:
'   - Empty: This will search all worksheets of the workbook.
'   - String: The name of the worksheet to search.
'   - String: The names of the worksheets to search, separated by a ':' character.
'   - Array: A one dimensional array whose elements are any of the following:
'           - Object: A worksheet object to search. This must be in the same workbook
'               as InWorkbook.
'           - String: The name of the worksheet to search.
'           - Number: The index number of the worksheet to search.
' If any one of the specificed worksheets is not found in InWorkbook, no search is
' performed. The search takes place only after everything has been validated.
'
' The other parameters have the same meaning and effect on the search as they do
' in the Range.Find method.
'
' Most of the code in this procedure deals with the InWorksheets parameter to give
' the absolute maximum flexibility in specifying which sheet to search.
'
' This function requires the FindAll procedure, also in this module or avaialable
' at www.cpearson.com/Excel/FindAll.aspx.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim WSArray() As String
Dim WS As Worksheet
Dim WB As Workbook
Dim ResultRange() As Range
Dim WSNdx As Long
Dim R As Range
Dim SearchRange As Range
Dim FoundRange As Range
Dim WSS As Variant
Dim N As Long


'''''''''''''''''''''''''''''''''''''''''''
' Determine what Workbook to search.
'''''''''''''''''''''''''''''''''''''''''''
If InWorkbook Is Nothing Then
    Set WB = ActiveWorkbook
Else
    Set WB = InWorkbook
End If

'''''''''''''''''''''''''''''''''''''''''''
' Determine what sheets to search
'''''''''''''''''''''''''''''''''''''''''''
If IsEmpty(InWorksheets) = True Then
    ''''''''''''''''''''''''''''''''''''''''''
    ' Empty. Search all sheets.
    ''''''''''''''''''''''''''''''''''''''''''
    With WB.Worksheets
        ReDim WSArray(1 To .Count)
        For WSNdx = 1 To .Count
            WSArray(WSNdx) = .Item(WSNdx).Name
        Next WSNdx
    End With

Else
    '''''''''''''''''''''''''''''''''''''''
    ' If Object, ensure it is a Worksheet
    ' object.
    ''''''''''''''''''''''''''''''''''''''
    If IsObject(InWorksheets) = True Then
        If TypeOf InWorksheets Is Excel.Worksheet Then
            ''''''''''''''''''''''''''''''''''''''''''
            ' Ensure Worksheet is in the WB workbook.
            ''''''''''''''''''''''''''''''''''''''''''
            If StrComp(InWorksheets.Parent.Name, WB.Name, vbTextCompare) <> 0 Then
                ''''''''''''''''''''''''''''''
                ' Sheet is not in WB. Get out.
                ''''''''''''''''''''''''''''''
                Exit Function
            Else
                ''''''''''''''''''''''''''''''
                ' Same workbook. Set the array
                ' to the worksheet name.
                ''''''''''''''''''''''''''''''
                ReDim WSArray(1 To 1)
                WSArray(1) = InWorksheets.Name
            End If
        Else
            '''''''''''''''''''''''''''''''''''''
            ' Object is not a Worksheet. Get out.
            '''''''''''''''''''''''''''''''''''''
        End If
    Else
        '''''''''''''''''''''''''''''''''''''''''''
        ' Not empty, not an object. Test for array.
        '''''''''''''''''''''''''''''''''''''''''''
        If IsArray(InWorksheets) = True Then
            '''''''''''''''''''''''''''''''''''''''
            ' It is an array. Test if each element
            ' is an object. If it is a worksheet
            ' object, get its name. Any other object
            ' type, get out. Not an object, assume
            ' it is the name.
            ''''''''''''''''''''''''''''''''''''''''
            ReDim WSArray(LBound(InWorksheets) To UBound(InWorksheets))
            For WSNdx = LBound(InWorksheets) To UBound(InWorksheets)
                If IsObject(InWorksheets(WSNdx)) = True Then
                    If TypeOf InWorksheets(WSNdx) Is Excel.Worksheet Then
                        ''''''''''''''''''''''''''''''''''''''
                        ' It is a worksheet object, get name.
                        ''''''''''''''''''''''''''''''''''''''
                        WSArray(WSNdx) = InWorksheets(WSNdx).Name
                    Else
                        ''''''''''''''''''''''''''''''''
                        ' Other type of object, get out.
                        ''''''''''''''''''''''''''''''''
                        Exit Function
                    End If
                Else
                    '''''''''''''''''''''''''''''''''''''''''''
                    ' Not an object. If it is an integer or
                    ' long, assume it is the worksheet index
                    ' in workbook WB.
                    '''''''''''''''''''''''''''''''''''''''''''
                    Select Case UCase(TypeName(InWorksheets(WSNdx)))
                        Case "LONG", "INTEGER"
                            Err.Clear
                            '''''''''''''''''''''''''''''''''''
                            ' Ensure integer if valid index.
                            '''''''''''''''''''''''''''''''''''
                            Set WS = WB.Worksheets(InWorksheets(WSNdx))
                            If Err.Number <> 0 Then
                                '''''''''''''''''''''''''''''''
                                ' Invalid index.
                                '''''''''''''''''''''''''''''''
                                Exit Function
                            End If
                            ''''''''''''''''''''''''''''''''''''
                            ' Valid index. Get name.
                            ''''''''''''''''''''''''''''''''''''
                            WSArray(WSNdx) = WB.Worksheets(InWorksheets(WSNdx)).Name
                        Case "STRING"
                            Err.Clear
                            '''''''''''''''''''''''''''''''''''''
                            ' Ensure valid name.
                            '''''''''''''''''''''''''''''''''''''
                            Set WS = WB.Worksheets(InWorksheets(WSNdx))
                            If Err.Number <> 0 Then
                                '''''''''''''''''''''''''''''''''
                                ' Invalid name, get out.
                                '''''''''''''''''''''''''''''''''
                                Exit Function
                            End If
                            WSArray(WSNdx) = InWorksheets(WSNdx)
                    End Select
                End If
                'WSArray(WSNdx) = InWorksheets(WSNdx)
            Next WSNdx
        Else
            ''''''''''''''''''''''''''''''''''''''''''''
            ' InWorksheets is neither an object nor an
            ' array. It is either the name or index of
            ' the worksheet.
            ''''''''''''''''''''''''''''''''''''''''''''
            Select Case UCase(TypeName(InWorksheets))
                Case "INTEGER", "LONG"
                    '''''''''''''''''''''''''''''''''''''''
                    ' It is a number. Ensure sheet exists.
                    '''''''''''''''''''''''''''''''''''''''
                    Err.Clear
                    Set WS = WB.Worksheets(InWorksheets)
                    If Err.Number <> 0 Then
                        '''''''''''''''''''''''''''''''
                        ' Invalid index, get out.
                        '''''''''''''''''''''''''''''''
                        Exit Function
                    Else
                        WSArray = Array(WB.Worksheets(InWorksheets).Name)
                    End If
                Case "STRING"
                    '''''''''''''''''''''''''''''''''''''''''''''''''''
                    ' See if the string contains a ':' character. If
                    ' so, the InWorksheets contains a string of multiple
                    ' worksheets.
                    '''''''''''''''''''''''''''''''''''''''''''''''''''
                    If InStr(1, InWorksheets, ":", vbBinaryCompare) > 0 Then
                        ''''''''''''''''''''''''''''''''''''''''''
                        ' ":" character found. split apart sheet
                        ' names.
                        ''''''''''''''''''''''''''''''''''''''''''
                        WSS = Split(InWorksheets, ":")
                        Err.Clear
                        N = LBound(WSS)
                        If Err.Number <> 0 Then
                            '''''''''''''''''''''''''''''
                            ' Unallocated array. Get out.
                            '''''''''''''''''''''''''''''
                            Exit Function
                        End If
                        If LBound(WSS) > UBound(WSS) Then
                            '''''''''''''''''''''''''''''
                            ' Unallocated array. Get out.
                            '''''''''''''''''''''''''''''
                            Exit Function
                        End If
                            
                                                
                        ReDim WSArray(LBound(WSS) To UBound(WSS))
                        For N = LBound(WSS) To UBound(WSS)
                            Err.Clear
                            Set WS = WB.Worksheets(WSS(N))
                            If Err.Number <> 0 Then
                                Exit Function
                            End If
                            WSArray = WSS(N)
                         Next N
                    Else
                        Err.Clear
                        Set WS = WB.Worksheets(InWorksheets)
                        If Err.Number <> 0 Then
                            '''''''''''''''''''''''''''''''''
                            ' Invalid name, get out.
                            '''''''''''''''''''''''''''''''''
                            Exit Function
                        Else
                            WSArray = Array(InWorksheets)
                        End If
                    End If
            End Select
        End If
    End If
End If
'''''''''''''''''''''''''''''''''''''''''''
' Ensure SearchAddress is valid
'''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
For WSNdx = LBound(WSArray) To UBound(WSArray)
    Err.Clear
    Set WS = WB.Worksheets(WSArray(WSNdx))
    ''''''''''''''''''''''''''''''''''''''''
    ' Worksheet does not exist
    ''''''''''''''''''''''''''''''''''''''''
    If Err.Number <> 0 Then
        Exit Function
    End If
    Err.Clear
    Set R = WB.Worksheets(WSArray(WSNdx)).Range(SearchAddress)
    If Err.Number <> 0 Then
        ''''''''''''''''''''''''''''''''''''
        ' Invalid Range. Get out.
        ''''''''''''''''''''''''''''''''''''
        Exit Function
    End If
Next WSNdx

''''''''''''''''''''''''''''''''''''''''
' SearchAddress is valid for all sheets.
' Call FindAll to search the range on
' each sheet.
''''''''''''''''''''''''''''''''''''''''
ReDim ResultRange(LBound(WSArray) To UBound(WSArray))
For WSNdx = LBound(WSArray) To UBound(WSArray)
    Set WS = WB.Worksheets(WSArray(WSNdx))
    Set SearchRange = WS.Range(SearchAddress)
    Set FoundRange = FindAll(SearchRange:=SearchRange, _
                    FindWhat:=FindWhat, _
                    LookIn:=LookIn, LookAt:=LookAt, _
                    SearchOrder:=SearchOrder, _
                    MatchCase:=MatchCase)
    If FoundRange Is Nothing Then
        Set ResultRange(WSNdx) = Nothing
    Else
        Set ResultRange(WSNdx) = FoundRange
    End If
Next WSNdx

FindAllOnWorksheets = ResultRange

End Function




