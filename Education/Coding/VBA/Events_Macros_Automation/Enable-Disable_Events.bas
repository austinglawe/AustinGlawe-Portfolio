' -----------------------------------------
' VBA Events_Macros_Automation:
' Enabling and disabling events
' -----------------------------------------
'
' Purpose:
' - Control Excel event handling to avoid recursive triggers.
'
' Property:
'   Application.EnableEvents = True or False
'
' Usage:
'   Application.EnableEvents = False
'   ' Code that modifies sheets or cells
'   Application.EnableEvents = True
'
' Example:
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       On Error GoTo Cleanup
'       Application.EnableEvents = False
'       ' Event code that changes cells
'   Cleanup:
'       Application.EnableEvents = True
'   End Sub
'
' Best practices:
' - Use error handling to guarantee re-enabling events.
' - Never leave events disabled unintentionally.
'
' -----------------------------------------
