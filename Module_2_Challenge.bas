Attribute VB_Name = "Module1"
Sub get_unique_ticker_values()
Dim row As Long
row = Cells(Rows.Count, "A").End(xlUp).row
ActiveSheet.Range("A1:A" & row).AdvancedFilter _
Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("K1"), _
Unique:=True
End Sub

Sub setup_data_formatting()
With Sheets("2018")
    Range("M2:M4001").NumberFormat = "0.00%"
End With
End Sub

Sub calc_range_of_unique_ticker()
Dim startRow As Long
Dim endRow As Long
Dim lastRow As Long
With Sheets("2018")
lastRow = .Cells(Rows.Count, "K").End(xlUp).row

For Each I In .Range("K2", "K" & lastRow)
    startRow = .Range("A:A").Find(what:=I, after:=.Range("A1")).row
    endRow = .Range("A:A").Find(what:=I, after:=.Range("A1"), lookat:=xlWhole, searchdirection:=xlPrevious).row
    Call calc_total_volume(startRow, endRow)
    Call calc_percentage_change(startRow, endRow)
    Call calc_yearly_change(startRow, endRow)
Next I
End With
End Sub

Sub calc_total_volume(startRow As Long, endRow As Long)
Dim sum As Double
Dim lastRow As Long
With Sheets("2018")
    lastRow = .Cells(Rows.Count, "N").End(xlUp).row + 1
    .Cells(lastRow, "N").Value = Application.WorksheetFunction.sum(Range("G" & startRow, "G" & endRow))
End With
End Sub

Sub calc_yearly_change(startRow As Long, endRow As Long)
With Sheets("2018")
    Set openAmount = .Cells(startRow, "C")
    Set closeAmount = .Cells(endRow, "F")
    lastRow = .Cells(Rows.Count, "L").End(xlUp).row + 1
    .Cells(lastRow, "L").Value = closeAmount - openAmount
End With
End Sub

Sub calc_percentage_change(startRow As Long, endRow As Long)
With Sheets("2018")
    Set openAmount = .Cells(startRow, "C")
    Set closeAmount = .Cells(endRow, "F")
    lastRow = .Cells(Rows.Count, "M").End(xlUp).row + 1
    .Cells(lastRow, "M").Value = (closeAmount - openAmount) / openAmount
End With
End Sub


