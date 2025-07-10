' -----------------------------------------
' VBA Events_Macros_Automation:
' Workbook-level events
' -----------------------------------------
'
' Purpose:
' - Run code automatically on workbook actions.
'
' Common events in ThisWorkbook module:
' - Workbook_Open: Runs when workbook opens.
' - Workbook_BeforeClose(Cancel As Boolean): Runs before workbook closes.
' - Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range): Runs on sheet changes.
' - Workbook_SheetActivate(ByVal Sh As Object): Runs when a sheet is activated.
'
' Example Workbook_Open:
'   Private Sub Workbook_Open()
'       MsgBox "Welcome!"
'   End Sub
'
' Example Workbook_BeforeClose:
'   Private Sub Workbook_BeforeClose(Cancel As Boolean)
'       Dim answer As VbMsgBoxResult
'       answer = MsgBox("Save before closing?", vbYesNoCancel)
'       If answer = vbYes Then ThisWorkbook.Save
'       If answer = vbCancel Then Cancel = True
'   End Sub
'
' Best practices:
' - Place code in ThisWorkbook module.
' - Keep code efficient.
' - Use Cancel parameter to control actions.
' - Avoid recursive event calls.
'
' -----------------------------------------
