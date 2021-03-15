Attribute VB_Name = "modELookup"
Public Function ELookup(Expr As String, Domain As String, Optional Criteria As Variant, _
    Optional OrderClause As Variant) As Variant
'Faster and more flexible replacement for DLookup()
'http://allenbrowne.com/ser-42.html RequiredLevel:4

On Error GoTo Err_ELookup
    'Arguments: Same as DLookup, with additional Order By option.
    'Return:    Value of the Expr if found, else Null.
    '           Delimited list for multi-value field.
    'Author:    Allen Browne. allen@allenbrowne.com
    'Updated:   December 2006, to handle multi-value fields (Access 2007 and later.)
    'Examples:
    '           1. To find the last value, include DESC in the OrderClause, e.g.:
    '               ELookup("[Surname] & [FirstName]", "tblClient", , "ClientID DESC")
    '           2. To find the lowest non-null value of a field, use the Criteria, e.g.:
    '               ELookup("ClientID", "tblClient", "Surname Is Not Null" , "Surname")
    'Note:      Requires a reference to the DAO library.
    Dim db As DAO.Database          'This database.
    Dim rs As DAO.Recordset         'To retrieve the value to find.
    Dim rsMVF As DAO.Recordset      'Child recordset to use for multi-value fields.
    Dim varResult As Variant        'Return value for function.
    Dim strSql As String            'SQL statement.
    Dim strOut As String            'Output string to build up (multi-value field.)
    Dim lngLen As Long              'Length of string.
    Const strcSep = ","             'Separator between items in multi-value list.

    'Initialize to null.
    varResult = Null

    'Build the SQL string.
    strSql = "SELECT TOP 1 " & Expr & " FROM " & Domain
    If Not IsMissing(Criteria) Then
        strSql = strSql & " WHERE " & Criteria
    End If
    If Not IsMissing(OrderClause) Then
        strSql = strSql & " ORDER BY " & OrderClause
    End If
    strSql = strSql & ";"

    'Lookup the value.
    Set db = DBEngine(0)(0)
    Set rs = db.OpenRecordset(strSql, dbOpenForwardOnly)
    If rs.RecordCount > 0 Then
        'Will be an object if multi-value field.
        If VarType(rs(0)) = vbObject Then
            Set rsMVF = rs(0).Value
            Do While Not rsMVF.EOF
                If rs(0).Type = 101 Then        'dbAttachment
                    strOut = strOut & rsMVF!FileName & strcSep
                Else
                    strOut = strOut & rsMVF![Value].Value & strcSep
                End If
                rsMVF.MoveNext
            Loop
            'Remove trailing separator.
            lngLen = Len(strOut) - Len(strcSep)
            If lngLen > 0& Then
                varResult = Left(strOut, lngLen)
            End If
            Set rsMVF = Nothing
        Else
            'Not a multi-value field: just return the value.
            varResult = rs(0)
        End If
    End If
    rs.Close

    'Assign the return value.
    ELookup = varResult

Exit_ELookup:
    Set rs = Nothing
    Set db = Nothing
    Exit Function

Err_ELookup:
    MsgBox Err.Description, vbExclamation, "ELookup Error " & Err.Number
    Resume Exit_ELookup
End Function

