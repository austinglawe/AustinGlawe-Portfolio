' -----------------------------------------
' VBA Charts_Graphs_Visualization:
' Creating basic charts from ranges or tables
' -----------------------------------------
'
' Steps:
' 1. Define worksheet and data range.
' 2. Add ChartObject to worksheet.
' 3. Set chart source data.
' 4. Specify chart type.
'
' Example:
'   Dim ws As Worksheet
'   Dim chartObj As ChartObject
'   Dim dataRange As Range
'
'   Set ws = Worksheets("Report")
'   Set dataRange = ws.Range("A1:B6")
'
'   Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=50, Width:=400, Height:=300)
'
'   With chartObj.Chart
'       .SetSourceData Source:=dataRange
'       .ChartType = xlColumnClustered
'       .HasTitle = True
'       .ChartTitle.Text = "Sales Data"
'   End With
'
' Best practices:
' - Use named ranges or tables for dynamic data.
' - Position and size charts well.
' - Use descriptive titles.
'
' -----------------------------------------
