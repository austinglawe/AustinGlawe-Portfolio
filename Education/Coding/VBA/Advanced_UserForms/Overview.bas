' -----------------------------------------
' VBA UserForms_AdvancedControls:
' Advanced UserForm controls and customization
' -----------------------------------------
'
' Adding controls:
' - Tools > Additional Controls in VBA Editor
' - Microsoft TreeView Control
' - Microsoft ListView Control
' - Microsoft Calendar Control (may need separate install)
'
' TreeView example:
'   With Me.TreeView1.Nodes
'       .Clear
'       .Add Key:="root", Text:="Root Node"
'       .Add Key:="child1", Text:="Child Node 1", Parent:="root"
'       .Add Key:="child2", Text:="Child Node 2", Parent:="root"
'   End With
'
' ListView example:
'   With Me.ListView1
'       .View = lvwReport
'       .Gridlines = True
'       .FullRowSelect = True
'       .ColumnHeaders.Add , , "Name", 100
'       .ColumnHeaders.Add , , "Value", 70
'       .ListItems.Add , , "Item 1"
'       .ListItems(1).ListSubItems.Add , , "123"
'   End With
'
' Best practices:
' - Verify MSCOMCTL availability on target machines.
' - Handle events for interactivity.
' - Design for usability.
' - Test on all target Office versions and bitness.
'
' -----------------------------------------
