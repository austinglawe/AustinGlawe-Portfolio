' -----------------------------------------
' VBA Charts_Graphs_Visualization:
' Updating and refreshing charts dynamically
' -----------------------------------------
'
' Update chart source data:
'   With chartObj.Chart
'       .SetSourceData Source:=Worksheets("Report").Range("A1:B20")
'   End With
'
' Update series values and categories:
'   With chartObj.Chart.SeriesCollection(1)
'       .Values = Worksheets("Report").Range("B2:B20")
'       .XValues = Worksheets("Report").Range("A2:A20")
'   End With
'
' Refresh chart:
'   chartObj.Chart.Refresh
'
' Best practices:
' - Use named ranges or tables for dynamic data.
' - Update Values and XValues for precise control.
' - Update existing charts, don’t recreate.
' - Maintain formatting consistency.
'
' -----------------------------------------
