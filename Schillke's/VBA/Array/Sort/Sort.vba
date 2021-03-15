<RequiredLevel:3>
'<Array:List> = 'Array("b", "a")
Dim First As Long, Last As Long
Dim i As Long, j As Long
Dim Temp As String
First = LBound(<Array:List>)
Last = UBound(<Array:List>)
For i = First To Last - 1
    For j = i + 1 To Last
        If UCase(<Array:List>(i)) > UCase(<Array:List>(j)) Then
            Temp = <Array:List>(j)
            <Array:List>(j) = <Array:List>(i)
            <Array:List>(i) = Temp
        End If
    Next j
Next i