<include modDatePicker.bas>
Dim datEmpty As Date, datPicked As Date
datPicked = PickDate(BeginDate:=#4/15/2015#)
If datPicked > datEmpty Then
    <cursor>
End If
