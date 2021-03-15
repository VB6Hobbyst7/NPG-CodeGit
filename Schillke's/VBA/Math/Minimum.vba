<Variant:List> = 'Array(1,3,2)
Dim i As Long,
Dim Minimum as Variant
Minimum = <Variant:List>(LBound(<Variant:List>))
For i = LBound(<Variant:List>) To UBound(<Variant:List>)
    If <Variant:List>(i) < Minimum Then Minimum = <Variant:List>(i)
Next i