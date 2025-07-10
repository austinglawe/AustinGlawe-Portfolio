' -----------------------------------------
' VBA Reports_Automation:
' Automating printing
' -----------------------------------------
'
' Print preview worksheet:
'   Worksheets("Report").PrintPreview
'
' Print worksheet directly:
'   Worksheets("Report").PrintOut Copies:=1, Collate:=True
'
' Print specific range:
'   Worksheets("Report").Range("A1:H50").PrintOut Copies:=1
'
' Print entire workbook:
'   ThisWorkbook.PrintOut Copies:=1
'
' Best practices:
' - Use PrintPreview to verify before printing.
' - Ensure printer settings are correct.
' - Use Collate:=True when printing multiple copies.
' - Handle printer errors gracefully.
'
' -----------------------------------------
