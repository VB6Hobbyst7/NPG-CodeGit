'Schillke

Sub Loop_Example()
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim Lrow As Long
    Dim CalcMode As Long
    Dim ViewMode As Long

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    'We use the ActiveSheet but you can replace this with
    'Sheets("MySheet")if you want
    With ActiveSheet

        'We select the sheet so we can change the window view
        .Select

        'If you are in Page Break Preview Or Page Layout view go
        'back to normal view, we do this for speed
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

        'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False

        'Set the first and last row to loop through
        Firstrow = .UsedRange.Cells(1).Row
        Lastrow = .UsedRange.Rows(.UsedRange.Rows.Count).Row

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow = Lastrow To Firstrow Step -1

            'We check the values in the A column in this example
            With .Cells(Lrow, "A")

                If Not IsError(.Value) Then

                    If .Value = "ron" Then .EntireRow.Delete
                    'This will delete each row with the Value "ron"
                    'in Column A, case sensitive.

                End If

            End With

        Next Lrow

    End With

    ActiveWindow.View = ViewMode
    With Application
        .ScreenUpdating = True
        .Calculation = CalcMode
    End With

End Sub