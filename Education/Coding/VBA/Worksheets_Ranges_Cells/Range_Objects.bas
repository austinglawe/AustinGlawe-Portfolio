' -----------------------------------------
' VBA Worksheets_Ranges_Cells:
' Working with Range objects
' -----------------------------------------
'
' Referencing a single cell:
'   Dim rng As Range
'   Set rng = Worksheets("Sheet1").Range("A1")
'
' Referencing multiple cells:
'   Set rng = Worksheets("Sheet1").Range("A1:C3")
'
' Referencing entire row:
'   Set rng = Worksheets("Sheet1").Rows(1)
'
' Referencing entire column:
'   Set rng = Worksheets("Sheet1").Columns("A")
'
' Referencing named range:
'   Set rng = Worksheets("Sheet1").Range("MyNamedRange")
'
' Using Cells(row, column):
'   Set rng = Worksheets("Sheet1").Cells(1, 1)  ' Equivalent to A1
'
' Example with loop:
'   Dim i As Integer
'   For i = 1 To 10
'       Worksheets("Sheet1").Cells(i, 1).Value = i
'   Next i
'
' Offset example:
'   Set rng = Worksheets("Sheet1").Range("A1").Offset(1, 2)  ' C2
'
' Resize example:
'   Set rng = Worksheets("Sheet1").Range("A1").Resize(3, 2)  ' 3 rows x 2 columns starting from A1
'
' Best practices:
' - Fully qualify all Range references to avoid context issues.
' - Range("A1") without worksheet refers to ActiveSheet (be cautious).
' - Offset and Resize provide flexible ways to work relative to an anchor cell.
'
' -----------------------------------------
