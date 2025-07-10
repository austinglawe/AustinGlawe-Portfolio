' -----------------------------------------
' VBA Reports_Automation:
' Formatting worksheets and ranges
' -----------------------------------------
'
' Format fonts:
'   With Range("A1:D10")
'       .Font.Name = "Calibri"
'       .Font.Size = 12
'       .Font.Bold = True
'   End With
'
' Format cell interior color:
'   .Interior.Color = RGB(220, 230, 241)
'
' Apply borders:
'   .Borders.LineStyle = xlContinuous
'   .Borders.Weight = xlThin
'
' Selective border example:
'   With .Borders(xlEdgeBottom)
'       .LineStyle = xlContinuous
'       .Weight = xlMedium
'       .Color = RGB(0,0,0)
'   End With
'
' Alignment and wrap text:
'   .HorizontalAlignment = xlCenter
'   .VerticalAlignment = xlCenter
'   .WrapText = True
'
' Number formatting:
'   .NumberFormat = "$#,##0.00"
'
' Merging cells and autofit:
'   With Worksheets("Report")
'       .Range("A1:D1").Merge
'       .Range("A1:D1").HorizontalAlignment = xlCenter
'       .Columns("A:D").AutoFit
'   End With
'
' Best practices:
' - Consistent styling for professionalism.
' - Use RGB for custom colors.
' - Wrap text for long text fields.
' - Autofit columns after populating data.
'
' -----------------------------------------
