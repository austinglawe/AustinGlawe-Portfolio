' -----------------------------------------
' VBA PivotTables_Reporting:
' Automating report generation with multiple PivotTables
' -----------------------------------------
'
' Tips:
' - Clear or create sheets to avoid duplicates.
' - Store PivotTable references for management.
' - Use loops to create multiple PivotTables dynamically.
' - Apply formatting and filtering consistently.
' - Refresh all PivotTables after creation.
'
' Example:
'   Dim categories As Variant
'   categories = Array("Electronics", "Clothing", "Toys")
'
'   For i = LBound(categories) To UBound(categories)
'       Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
'       Set pivotPos = wsReport.Cells(3 + (i - LBound(categories)) * 15, 1)
'       Set pt = pc.CreatePivotTable(TableDestination:=pivotPos, TableName:="Pivot_" & categories(i))
'       With pt
'           .PivotFields("Category").ClearAllFilters
'           .PivotFields("Category").CurrentPage = categories(i)
'           .AddDataField .PivotFields("Sales"), "Sum of Sales", xlSum
'       End With
'   Next i
'
' Best practices:
' - Use unique PivotTable names.
' - Plan worksheet layout for multiple PivotTables.
' - Modularize repetitive code.
'
' -----------------------------------------
