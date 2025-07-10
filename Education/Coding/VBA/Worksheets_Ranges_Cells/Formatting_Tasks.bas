' -----------------------------------------
' VBA Worksheets_Ranges_Cells:
' Common formatting tasks
' -----------------------------------------
'
' Font properties:
'   With Worksheets("Sheet1").Range("A1")
'       .Font.Bold = True
'       .Font.Italic = True
'       .Font.Size = 12
'       .Font.Name = "Calibri"
'   End With
'
' Font and fill colors:
'   With Worksheets("Sheet1").Range("A1")
'       .Font.Color = RGB(255, 0, 0)        ' Red font
'       .Interior.Color = RGB(255, 255, 0)  ' Yellow fill
'   End With
'
' Borders:
'   Dim rng As Range
'   Set rng = Worksheets("Sheet1").Range("A1:B2")
'
'   With rng.Borders
'       .LineStyle = xlContinuous
'       .Weight = xlThin
'       .Color = RGB(0, 0, 0)
'   End With
'
'   Individual border example:
'       rng.Borders(xlEdgeBottom).LineStyle = xlContinuous
'
' Number formats:
'   With Worksheets("Sheet1").Range("A1")
'       .NumberFormat = "$#,##0.00"  ' Currency
'   End With
'
' Common formats:
' - "$#,##0.00" : Currency with 2 decimals
' - "0%" : Percentage
' - "mm/dd/yyyy" : Date format
' - "@" : Text
' - "0.00" : 2 decimal places
'
' Best practices:
' - Use With blocks for clean formatting code.
' - Fully qualify ranges to avoid ambiguity.
' - Prefer RGB() for specifying colors.
' - Format directly; avoid using Select.
'
' -----------------------------------------
