' -----------------------------------------
' VBA PivotTables_Reporting:
' Creating PivotCaches and PivotTables
' -----------------------------------------
'
' Steps:
' 1. Define source data range.
' 2. Create PivotCache from source data.
' 3. Create PivotTable on target sheet and cell.
'
' Example:
'   Dim wb As Workbook
'   Dim wsData As Worksheet, wsReport As Worksheet
'   Dim pc As PivotCache
'   Dim pt As PivotTable
'   Dim dataRange As Range
'   Dim pivotDest As Range
'
'   Set wb = ThisWorkbook
'   Set wsData = wb.Worksheets("Data")
'   Set wsReport = wb.Worksheets("Report")
'   Set dataRange = wsData.Range("A1:D100")
'   Set pivotDest = wsReport.Range("A3")
'
'   Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
'   Set pt = pc.CreatePivotTable(TableDestination:=pivotDest, TableName:="SalesPivot")
'
' Notes:
' - SourceType:=xlDatabase indicates data is from a worksheet range.
' - Use dynamic ranges or Excel Tables for flexible source.
' - PivotTable name must be unique.
'
' Best practices:
' - Delete or clear old PivotTables if regenerating.
' - Use error handling to manage duplicates.
'
' -----------------------------------------
