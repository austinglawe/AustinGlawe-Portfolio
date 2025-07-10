' -----------------------------------------
' VBA Charts_Graphs_Visualization:
' Adding trendlines and data labels
' -----------------------------------------
'
' Add trendline:
'   With chartObj.Chart.SeriesCollection(1)
'       .Trendlines.Add Type:=xlLinear, Forward:=0, Backward:=0, DisplayEquation:=True, DisplayRSquared:=True
'   End With
'
' Customize data labels:
'   With chartObj.Chart.SeriesCollection(1).DataLabels
'       .ShowValue = True
'       .ShowCategoryName = False
'       .ShowSeriesName = False
'       .Font.Bold = True
'       .Font.Size = 11
'   End With
'
' Position data labels:
'   .DataLabels.Position = xlLabelPositionAbove
'
' Best practices:
' - Use trendlines to show trends.
' - Show R² only for knowledgeable audiences.
' - Keep data labels uncluttered.
' - Position labels for readability.
'
' -----------------------------------------
