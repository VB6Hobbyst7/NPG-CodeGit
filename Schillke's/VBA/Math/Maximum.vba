<Variant:List> = 'Array(1,3,2)
Dim i As Long,
Dim Maximum as Variant
Maximum = <Variant:List>(LBound(<Variant:List>))
For i = LBound(<Variant:List>) To UBound(<Variant:List>)
    If <Variant:List>(i) > Maximum Then Maximum = <Variant:List>(i)
Next i