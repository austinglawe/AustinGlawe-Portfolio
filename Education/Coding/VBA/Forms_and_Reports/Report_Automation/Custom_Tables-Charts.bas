' -----------------------------------------
' VBA Reports_Automation:
' Adding and customizing tables and charts
' -----------------------------------------
'
' Adding a table:
'   Dim tbl As ListObject
'   Set tbl = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
'   tbl.Name = "SalesTable"
'   tbl.TableStyle = "TableStyleMedium9"
'
' Customize table:
'   tbl.ShowTotals = True
'   tbl.ShowTableStyleRowStripes = True
'   tbl.HeaderRowRange.Font.Bold = True
'
' Adding a chart:
'   Dim cht As ChartObject
'   Set cht = ws.ChartObjects.Add(Left:=300, Top:=50, Width:=400, Height:=250)
'   With cht.Chart
'       .SetSourceData Source:=tbl.Range
'       .ChartType = xlColumnClustered
'       .HasTitle = True
'       .ChartTitle.Text = "Sales by Region"
'       .Axes(xlCategory).HasTitle = True
'       .Axes(xlCategory).AxisTitle.Text = "Region"
'       .Axes(xlValue).HasTitle = True
'       .Axes(xlValue).AxisTitle.Text = "Sales Amount"
'   End With
'
' Formatting chart:
'   With cht.Chart.SeriesCollection(1)
'       .Format.Fill.ForeColor.RGB = RGB(0, 112, 192)
'   End With
'
' Best practices:
' - Use tables for dynamic chart ranges.
' - Name tables and charts clearly.
' - Position and size charts for readability.
' - Include titles and axis labels.
' - Apply consistent styling.
'
' -----------------------------------------
