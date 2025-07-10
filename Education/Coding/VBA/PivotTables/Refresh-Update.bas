' -----------------------------------------
' VBA PivotTables_Reporting:
' Refreshing PivotTables and updating source data
' -----------------------------------------
'
' Refresh a single PivotTable:
'   pt.RefreshTable
'
' Refresh all PivotTables in workbook:
'   Dim ws As Worksheet, pt As PivotTable
'   For Each ws In ThisWorkbook.Worksheets
'       For Each pt In ws.PivotTables
'           pt.RefreshTable
'       Next pt
'   Next ws
'
' Update source data dynamically:
'   Dim dataRange As Range
'   Set dataRange = wsData.Range("A1").CurrentRegion
'   pt.ChangePivotCache ThisWorkbook.PivotCaches.Create( _
'       SourceType:=xlDatabase, _
'       SourceData:=dataRange)
'   pt.RefreshTable
'
' Use Excel Tables as source (auto-expand):
'   pt.ChangePivotCache ThisWorkbook.PivotCaches.Create( _
'       SourceType:=xlDatabase, _
'       SourceData:="Table1")
'   pt.RefreshTable
'
' Best practices:
' - Use dynamic ranges or Tables for flexible source data.
' - Always refresh after data changes.
' - Refresh all PivotTables if multiple reports exist.
'
' -----------------------------------------
