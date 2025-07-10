' -----------------------------------------
' VBA Reports_Automation:
' Setting up page layout (margins, orientation, headers/footers)
' -----------------------------------------
'
' Page orientation and margins:
'   With Worksheets("Report").PageSetup
'       .Orientation = xlLandscape
'       .TopMargin = Application.InchesToPoints(0.75)
'       .BottomMargin = Application.InchesToPoints(0.75)
'       .LeftMargin = Application.InchesToPoints(0.7)
'       .RightMargin = Application.InchesToPoints(0.7)
'       .HeaderMargin = Application.InchesToPoints(0.3)
'       .FooterMargin = Application.InchesToPoints(0.3)
'   End With
'
' Headers and footers:
'   With Worksheets("Report").PageSetup
'       .LeftHeader = "&D"        ' Date
'       .CenterHeader = "Sales Report"
'       .RightHeader = "&T"       ' Time
'       .LeftFooter = "Confidential"
'       .CenterFooter = "Page &P of &N"
'       .RightFooter = "&F"       ' File name
'   End With
'
' Print area and titles:
'   With Worksheets("Report").PageSetup
'       .PrintArea = "$A$1:$H$50"
'       .PrintTitleRows = "$1:$1"
'       .PrintTitleColumns = "$A:$A"
'   End With
'
' Best practices:
' - Consistent margins and orientation.
' - Useful headers/footers for context.
' - Precise print area.
' - Print titles for readability.
' - Test print preview.
'
' -----------------------------------------
