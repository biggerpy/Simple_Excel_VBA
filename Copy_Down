Public Sub downXTimes()

Dim sElected As Variant
Dim offSet As Long
Dim nRows As Long


nRows = ActiveSheet.UsedRange.Rows.Count
offSet = InputBox("How many cells?", , 9999)
If offSet = 9999 Then  'if no value is entered, use the default value
    offSet = nRows - ActiveCell.Row  ' fill all the cells to the end of the last row.
End If
sElected = ActiveCell.Value
ActiveSheet.Range(ActiveCell.Cells, ActiveCell.Cells.offSet(offSet, 0)).Select
Selection.Value = sElected

End Sub
