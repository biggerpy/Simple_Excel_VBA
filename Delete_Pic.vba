Public Sub paint_white()
Dim i As Integer
Dim j As Integer
Dim nRows As Integer
Dim nCols As Integer

nRows = ActiveSheet.UsedRange.Rows.Count
For i = 1 To nRows
    If Cells(i, 8).Interior.ColorIndex = -4124 And Trim(Cells(i, 8)) <> "Translation" And Trim(Cells(i, 8)) <> "English" And Trim(Cells(i, 8)) <> "" Then
        Cells(i, 8).Font.ColorIndex = 2
        Cells(i, 9).Font.ColorIndex = 2
    End If
Next i
        


End Sub

Public Sub test()
Dim i As Integer
Dim nRows As Integer

Dim pic As Shape
nRows = ActiveSheet.UsedRange.Rows.Count
For Each pic In ActiveSheet.Shapes
    pic.TopLeftCell.Activate  'located the picture
    Application.Goto ActiveCell.EntireRow, True   'scroll the window to the top
    pic.Select  ' select the pic
    choice = MsgBox("Delete?", vbYesNo, "Choice")
    If choice = vbYes Then
        pic.Delete
    End If
Next pic


End Sub
