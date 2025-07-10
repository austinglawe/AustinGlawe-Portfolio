' -----------------------------------------
' VBA Worksheets_Ranges_Cells:
' Reading/writing cell values
' -----------------------------------------
'
' Reading a cell value:
'   Dim val As Variant
'   val = Worksheets("Sheet1").Range("A1").Value
'
' Writing a value:
'   Worksheets("Sheet1").Range("A1").Value = 100
'
' Using a Worksheet object:
'   Dim ws As Worksheet
'   Set ws = Worksheets("Sheet1")
'   ws.Range("A1").Value = "Hello world"
'
' Writing a formula:
'   ws.Range("B1").Formula = "=A1*2"
'
' Reading displayed text:
'   Dim displayText As String
'   displayText = ws.Range("A1").Text
'
' Bulk read (array of values):
'   Dim arr As Variant
'   arr = ws.Range("A1:C3").Value
'
' Bulk write (array of values):
'   Dim arr(1 To 2, 1 To 3) As Variant
'   arr(1, 1) = "A"
'   arr(1, 2) = "B"
'   arr(1, 3) = "C"
'   arr(2, 1) = 1
'   arr(2, 2) = 2
'   arr(2, 3) = 3
'   ws.Range("A1:C2").Value = arr
'
' Notes:
' - .Value: Actual value.
' - .Text: Displayed text (read-only, formatting-dependent).
' - .Formula: Formula string (e.g., "=A1+B1").
'
' Best practices:
' - Fully qualify all references.
' - Prefer array reads/writes for performance on large ranges.
' - .Text is read-only and should not be confused with .Value.
'
' -----------------------------------------
