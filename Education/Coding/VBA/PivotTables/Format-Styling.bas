' -----------------------------------------
' VBA PivotTables_Reporting:
' Formatting and styling PivotTables
' -----------------------------------------
'
' Set PivotTable style:
'   pt.TableStyle2 = "PivotStyleMedium9"
'
' Adjust column widths:
'   pt.TableRange2.Columns.AutoFit
'
' Format data field numbers:
'   Dim df As PivotField
'   Set df = pt.DataFields(1)
'   df.NumberFormat = "$#,##0.00"
'
' Show/hide subtotals and grand totals:
'   pt.RowAxisLayout xlTabularRow
'   pt.RowFields("Region").Subtotals(1) = False
'   pt.ColumnGrand = False
'   pt.RowGrand = True
'
' Format font and style:
'   With pt.TableRange1.Font
'       .Name = "Calibri"
'       .Size = 11
'       .Bold = True
'   End With
'
' Best practices:
' - Use built-in styles for consistency.
' - Format after PivotTable creation/refresh.
' - Use TableRange1 and TableRange2 appropriately.
' - Clear old formatting when regenerating.
'
' -----------------------------------------
