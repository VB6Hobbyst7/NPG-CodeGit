<include modKeyState.bas>
Do While IsShiftKeyDown()
    DoEvents
Loop
Set <Workbook> = <Application.>Workbooks.Open(<String:Filename>)