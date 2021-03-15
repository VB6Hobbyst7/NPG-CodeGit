Attribute VB_Name = "modOnCloseKeep1Open"
Public Function Keep1Open(objMe As Object)
On Error GoTo Err_Keep1Open
    'Purpose:   Open the Switchboard if nothing else is visible.
    'Argument:  The object being closed.
    'Usage:     In the OnClose property of forms and reports:
    '               =Keep1Open([Form])
    '               =Keep1Open([Report])
    'Note:      Replace "Switchboard" with the name of your switchboard form.
    Dim frm As Form         'an open form.
    Dim rpt As Report       'an open report.
    Dim bFound As Boolean   'Flag not to open the switchboard.
    
    'Any other visible forms?
    If Not bFound Then
        For Each frm In Forms
            If (frm.Hwnd <> objMe.Hwnd) And (frm.Visible) Then
                bFound = True
                Exit For
            End If
        Next
    End If
    
    'Any other visible reports?
    If Not bFound Then
        For Each rpt In Reports
            If (rpt.Hwnd <> objMe.Hwnd) And (rpt.Visible) Then
                bFound = True
                Exit For
            End If
        Next
    End If
    
    'If none found, open the switchboard.
    If Not bFound Then
        DoCmd.OpenForm "Switchboard"
    End If
    
Exit_Keep1Open:
    Set frm = Nothing
    Set rpt = Nothing
    Exit Function
    
Err_Keep1Open:
    If Err.Number <> 2046& Then     'OpenForm is not available when closing database.
        'Call LogError(Err.Number, Err.Description, ".Keep1Open()")
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Keep1Open()"
    End If
End Function
