' -----------------------------------------
' VBA PivotTables_Reporting:
' Applying filters and slicers programmatically
' -----------------------------------------
'
' Filtering PivotFields:
'   With pt.PivotFields("Region")
'       .ClearAllFilters
'       ' For PageFields:
'       .CurrentPage = "West"
'
'       ' For Row/Column fields:
'       Dim pi As PivotItem
'       For Each pi In .PivotItems
'           pi.Visible = (pi.Name = "West")
'       Next pi
'   End With
'
' Adding slicers (Excel 2010+):
'   Dim slicerCache As SlicerCache
'   Set slicerCache = wb.SlicerCaches.Add(pt, "Region")
'
'   Dim slicer As Slicer
'   Set slicer = slicerCache.Slicers.Add(wsReport, , "Region", "Region Slicer", 10, 10, 144, 144)
'
' Controlling slicer selection:
'   Dim item As SlicerItem
'   For Each item In slicer.SlicerItems
'       item.Selected = (item.Name = "West")
'   Next item
'
' Best practices:
' - Clear filters before applying new.
' - Use Visible property carefully.
' - Position slicers logically.
'
' -----------------------------------------
