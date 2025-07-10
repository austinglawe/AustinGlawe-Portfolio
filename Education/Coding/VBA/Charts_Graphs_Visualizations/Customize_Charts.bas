' -----------------------------------------
' VBA Charts_Graphs_Visualization:
' Customizing chart types and data series
' -----------------------------------------
'
' Change chart type:
'   chartObj.Chart.ChartType = xlLineMarkers
'
' Modify series:
'   With chartObj.Chart.SeriesCollection(1)
'       .Name = "Sales 2025"
'       .Values = Worksheets("Report").Range("B2:B7")
'       .XValues = Worksheets("Report").Range("A2:A7")
'       .ChartType = xlColumnClustered
'   End With
'
' Add new series:
'   chartObj.Chart.SeriesCollection.NewSeries
'   With chartObj.Chart.SeriesCollection(2)
'       .Name = "Expenses 2025"
'       .Values = Worksheets("Report").Range("C2:C7")
'       .XValues = Worksheets("Report").Range("A2:A7")
'       .ChartType = xlLine
'   End With
'
' Combo charts:
'   With chartObj.Chart
'       .ChartType = xlColumnClustered
'       .SeriesCollection(2).ChartType = xlLine
'   End With
'
' Best practices:
' - Clear series names for legends.
' - Update Values and XValues dynamically.
' - Use combo charts for mixed data.
' - Remove unused series.
'
' -----------------------------------------
