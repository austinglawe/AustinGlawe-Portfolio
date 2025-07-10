' -----------------------------------------
' VBA Charts_Graphs_Visualization:
' Exporting charts as images or copying
' -----------------------------------------
'
' Export chart as image:
'   Dim cht As ChartObject
'   Set cht = Worksheets("Report").ChartObjects(1)
'   cht.Chart.Export Filename:="C:\Reports\SalesChart.png", FilterName:="PNG"
'
' Copy chart to clipboard:
'   cht.Chart.CopyPicture Appearance:=xlScreen, Format:=xlPicture
'
' Paste chart image to another sheet:
'   Worksheets("Sheet2").Paste Destination:=Worksheets("Sheet2").Range("A1")
'
' Best practices:
' - Use meaningful filenames and paths.
' - Verify export folders exist.
' - PNG format preferred for quality.
' - Use copy/paste to embed charts in other Office apps.
'
' -----------------------------------------
