Public Sub cleanBothEnds()

Dim iRow As Long
Dim iCol As Long
Dim nRows As Long
Dim nCols As Long


nRows = ActiveSheet.UsedRange.Rows.Count
nCols = ActiveSheet.UsedRange.Columns.Count

Call replaceNbsp   ' call another function.

For iCol = 1 To nCols
    For iRow = 1 To nRows
        ActiveSheet.Cells(iRow, iCol) = Trim(ActiveSheet.Cells(iRow, iCol))
    Next iRow
Next iCol

End Sub
