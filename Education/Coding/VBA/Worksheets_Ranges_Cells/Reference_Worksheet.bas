' -----------------------------------------
' VBA Worksheets_Ranges_Cells:
' Referencing worksheets
' -----------------------------------------
'
' Purpose:
' - Access worksheets programmatically.
'
' Collections:
' - Worksheets: only worksheet-type sheets.
' - Sheets: all sheets (worksheets, chart sheets, macro sheets).
'
' Reference by name:
'   Worksheets("Sheet1").Activate
'
' Reference by index:
'   Worksheets(1).Activate  ' First worksheet (by order)
'
' Reference in a workbook:
'   Workbooks("MyWorkbook.xlsx").Worksheets("Sheet1").Activate
'
' Using object variable:
'   Dim ws As Worksheet
'   Set ws = Worksheets("Data")
'   ws.Range("A1").Value = "Hello"
'
' Looping through worksheets:
'   Dim ws As Worksheet
'   For Each ws In Worksheets
'       Debug.Print ws.Name
'   Next ws
'
' Best practices:
' - Use Worksheet object variables for clarity and performance.
' - Prefer Worksheets("Name") when sheet names are known.
' - Reference workbook explicitly when working with multiple workbooks.
'
' -----------------------------------------
