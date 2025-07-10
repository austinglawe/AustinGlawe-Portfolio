' -----------------------------------------
' VBA Reports_Automation:
' Exporting reports to PDF and other formats
' -----------------------------------------
'
' Export worksheet as PDF:
'   Worksheets("Report").ExportAsFixedFormat _
'       Type:=xlTypePDF, _
'       Filename:="C:\Reports\SalesReport.pdf", _
'       Quality:=xlQualityStandard, _
'       IncludeDocProperties:=True, _
'       IgnorePrintAreas:=False, _
'       OpenAfterPublish:=True
'
' Export entire workbook as PDF:
'   ThisWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:="C:\Reports\FullWorkbook.pdf"
'
' Export worksheet as XLSX copy:
'   Worksheets("Report").Copy
'   ActiveWorkbook.SaveAs Filename:="C:\Reports\ReportCopy.xlsx", FileFormat:=xlOpenXMLWorkbook
'   ActiveWorkbook.Close False
'
' Export chart as image:
'   Dim cht As ChartObject
'   Set cht = Worksheets("Report").ChartObjects(1)
'   cht.Chart.Export Filename:="C:\Reports\Chart1.png", FilterName:="PNG"
'
' Best practices:
' - Define print areas before export.
' - Use IgnorePrintAreas:=False to respect print settings.
' - Verify export paths exist.
' - Use descriptive, timestamped filenames.
' - Close temporary workbooks after saving.
'
' -----------------------------------------
