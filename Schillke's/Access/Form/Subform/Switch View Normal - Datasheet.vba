With Me.fsub 'Name of the subform control
    .SetFocus
    If .Form.CurrentView = 2 Then 'acCurViewDatasheet
        <Application.>DoCmd.RunCommand 462 'acCmdSubformFormView
    Else
        <Application.>DoCmd.RunCommand 463 'acCmdSubformDatasheetView
    End If
End With
