Public Sub replaceNbsp()

'&nbsp; is from the web, which can't be trimmed by Trim function.

Dim iRow As Long
Dim iCol As Long
Dim nRows As Long
Dim nCols As Long

nRows = ActiveSheet.UsedRange.Rows.Count
nCols = ActiveSheet.UsedRange.Columns.Count
For iCol = 1 To nCols
    For iRow = 1 To nRows
        ActiveSheet.Cells(iRow, iCol) = Replace(ActiveSheet.Cells(iRow, iCol), Chr(160), "")
    Next iRow
Next iCol

End Sub
