' -----------------------------------------
' VBA PivotTables_Reporting:
' Adding and configuring PivotFields
' -----------------------------------------
'
' Add fields to PivotTable (pt):
'   With pt
'       ' Row field
'       .PivotFields("Region").Orientation = xlRowField
'       .PivotFields("Region").Position = 1
'
'       ' Column field
'       .PivotFields("Year").Orientation = xlColumnField
'       .PivotFields("Year").Position = 1
'
'       ' Filter (Page) field
'       .PivotFields("Category").Orientation = xlPageField
'       .PivotFields("Category").Position = 1
'
'       ' Data field with sum aggregation
'       .AddDataField .PivotFields("Sales"), "Sum of Sales", xlSum
'   End With
'
' Notes:
' - Use xlRowField, xlColumnField, xlPageField for placement.
' - Use AddDataField for values with summary functions (xlSum, xlCount, etc.).
' - Set Position to order fields.
'
' Best practices:
' - Check field existence before adding.
' - Clear fields if regenerating.
' - Use descriptive captions for data fields.
'
' -----------------------------------------
