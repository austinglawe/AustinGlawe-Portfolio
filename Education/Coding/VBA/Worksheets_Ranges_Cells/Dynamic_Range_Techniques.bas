' -----------------------------------------
' VBA Worksheets_Ranges_Cells:
' Dynamic range techniques
' -----------------------------------------
'
' UsedRange:
'   Dim rng As Range
'   Set rng = Worksheets("Sheet1").UsedRange
'
' Find last used row in a column:
'   Dim lastRow As Long
'   lastRow = Worksheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row
'
' Find last used column in a row:
'   Dim lastCol As Long
'   lastCol = Worksheets("Sheet1").Cells(1, Columns.Count).End(xlToLeft).Column
'
' Define dynamic range:
'   Dim dataRange As Range
'   Set dataRange = Worksheets("Sheet1").Range("A1:A" & lastRow)
'
' Find last used cell anywhere:
'   Dim lastRow As Long, lastCol As Long
'   lastRow = Worksheets("Sheet1").Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
'   lastCol = Worksheets("Sheet1").Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
'
' Best practices:
' - Use .End(xlUp) / .End(xlToLeft) for fast column/row scanning.
' - Use Find("*") for true last-used detection on messy sheets.
' - UsedRange can be inaccurate if sheet history isn’t "clean."
' - Avoid hardcoding row or column counts: adapt to your data!
'
' -----------------------------------------
