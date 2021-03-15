Attribute VB_Name = "modRangeIsProcedures"

Public Function IsInRange(rng1, rng2) As Boolean
'Returns True if rng1 is a subset of rng2
'http://codevba.com/excel/range.htm#IsInRange
'?IsInRange(Range("A1"), Range("A1:A2"))=True
'?IsInRange(Range("A1:A2"), Range("A1"))=False
    IsInRange = False
    If rng1.Parent.Parent.Name = rng2.Parent.Parent.Name Then
        If rng1.Parent.Name = rng2.Parent.Name Then
            If Union(rng1, rng2).Address = rng2.Address Then
                IsInRange = True
            End If
        End If
    End If
End Function
