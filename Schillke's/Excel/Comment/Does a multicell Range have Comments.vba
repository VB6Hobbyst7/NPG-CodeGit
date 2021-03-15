On Error Resume Next
<Boolean:RangeHasComments> = <Range>.SpecialCells(-4144).Cells.Count > 0 'xlCellTypeComments = -4144 