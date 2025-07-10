' -----------------------------------------
' VBA Charts_Graphs_Visualization:
' Formatting chart elements (titles, labels, axes)
' -----------------------------------------
'
' Chart title:
'   With chartObj.Chart
'       .HasTitle = True
'       .ChartTitle.Text = "Monthly Sales Overview"
'       .ChartTitle.Font.Size = 14
'       .ChartTitle.Font.Bold = True
'   End With
'
' Axis titles:
'   With chartObj.Chart.Axes(xlCategory)
'       .HasTitle = True
'       .AxisTitle.Text = "Month"
'   End With
'   With chartObj.Chart.Axes(xlValue)
'       .HasTitle = True
'       .AxisTitle.Text = "Sales ($)"
'   End With
'
' Data labels:
'   With chartObj.Chart.SeriesCollection(1)
'       .HasDataLabels = True
'       .DataLabels.Position = xlLabelPositionAbove
'       .DataLabels.Font.Size = 10
'   End With
'
' Legend formatting:
'   With chartObj.Chart.Legend
'       .Position = xlLegendPositionBottom
'       .Font.Size = 10
'   End With
'
' Gridlines:
'   chartObj.Chart.Axes(xlValue).HasMajorGridlines = True
'   chartObj.Chart.Axes(xlValue).HasMinorGridlines = False
'
' Best practices:
' - Use clear titles for context.
' - Position labels/legends to avoid clutter.
' - Format fonts/colors consistently.
' - Use gridlines to aid readability only.
'
' -----------------------------------------
