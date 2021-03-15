Attribute VB_Name = "modForm_BeforeInsert_CarryOver"
Option Compare Database
Option Explicit

Public Function CarryOver(frm As Form, strErrMsg As String, ParamArray avarExceptionList()) As Long
On Error GoTo Err_Handler
    'Purpose: Carry over the same fields to a new record, based on the last record in the form.
    'Arguments: frm               = the form to copy the values on.
    '           strErrMsg         = string to append error messages to.
    '           avarExceptionList = list of control names NOT to copy values over to.
    'Return:    Count of controls that had a value assigned.
    'Usage:     In a form's BeforeInsert event, excluding Surname and City controls:
    '               Call CarryOver(Me, strMsg, "Surname", City")
    Dim rs As DAO.Recordset         'Clone of form.
    Dim ctl As Control              'Each control on form.
    Dim strForm As String           'Name of form (for error handler.)
    Dim strControl As String        'Each control in the loop
    Dim strActiveControl As String  'Name of the active control. Don't assign this as user is typing in it.
    Dim strControlSource As String  'ControlSource property.
    Dim lngI As Long                'Loop counter.
    Dim lngLBound As Long           'Lower bound of exception list array.
    Dim lngUBound As Long           'Upper bound of exception list array.
    Dim bCancel As Boolean          'Flag to cancel this operation.
    Dim bSkip As Boolean            'Flag to skip one control.
    Dim lngKt As Long               'Count of controls assigned.

    'Initialize.
    strForm = frm.Name
    strActiveControl = frm.ActiveControl.Name
    lngLBound = LBound(avarExceptionList)
    lngUBound = UBound(avarExceptionList)

    'Must not assign values to the form's controls if it is not at a new record.
    If Not frm.NewRecord Then
        bCancel = True
        strErrMsg = strErrMsg & "Cannot carry values over. Form '" & strForm & "' is not at a new record." & vbCrLf
    End If
    'Find the record to copy, checking there is one.
    If Not bCancel Then
        Set rs = frm.RecordsetClone
        If rs.RecordCount <= 0& Then
            bCancel = True
            strErrMsg = strErrMsg & "Cannot carry values over. Form '" & strForm & "' has no records." & vbCrLf
        End If
    End If

    If Not bCancel Then
        'The last record in the form is the one to copy.
        rs.MoveLast
        'Loop the controls.
        For Each ctl In frm.Controls
            bSkip = False
            strControl = ctl.Name
            'Ignore the active control, those without a ControlSource, and those in the exception list.
            If (strControl <> strActiveControl) And HasProperty(ctl, "ControlSource") Then
                For lngI = lngLBound To lngUBound
                    If avarExceptionList(lngI) = strControl Then
                        bSkip = True
                        Exit For
                    End If
                Next
                If Not bSkip Then
                    'Examine what this control is bound to. Ignore unbound, or bound to an expression.
                    strControlSource = ctl.ControlSource
                    If (strControlSource <> vbNullString) And Not (strControlSource Like "=*") Then
                        'Ignore calculated fields (no SourceTable), autonumber fields, and null values.
                        With rs(strControlSource)
                            If (.SourceTable <> vbNullString) And ((.Attributes And dbAutoIncrField) = 0&) _
                                And Not (IsCalcTableField(rs(strControlSource)) Or IsNull(.Value)) Then
                                If ctl.Value = .Value Then
                                    'do nothing. (Skipping this can cause Error 3331.)
                                Else
                                    ctl.Value = .Value
                                    lngKt = lngKt + 1&
                                End If
                            End If
                        End With
                    End If
                End If
            End If
        Next
    End If

    CarryOver = lngKt

Exit_Handler:
    Set rs = Nothing
    Exit Function

Err_Handler:
    strErrMsg = strErrMsg & Err.Description & vbCrLf
    Resume Exit_Handler
End Function

Private Function IsCalcTableField(fld As DAO.Field) As Boolean
    'Purpose: Returns True if fld is a calculated field (Access 2010 and later only.)
On Error GoTo ExitHandler
    Dim strExpr As String

    strExpr = fld.Properties("Expression")
    If strExpr <> vbNullString Then
        IsCalcTableField = True
    End If

ExitHandler:
End Function

Private Function HasProperty(obj As Object, strPropName As String) As Boolean
    'Purpose: Return true if the object has the property.
    Dim varDummy As Variant

    On Error Resume Next
    varDummy = obj.Properties(strPropName)
    HasProperty = (Err.Number = 0)
End Function


