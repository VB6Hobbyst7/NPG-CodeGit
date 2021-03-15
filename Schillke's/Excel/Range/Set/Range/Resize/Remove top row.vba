<RequiredLevel:3>
Dim rngTopRowRemoved As Range
Set rngTopRowRemoved = <Range>.Offset(1, 0).Resize(<Range>.Rows.Count - 1, <Range>.Columns.Count)