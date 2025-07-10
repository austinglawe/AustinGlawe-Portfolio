' -----------------------------------------
' VBA Events_Macros_Automation:
' Worksheet-level events
' -----------------------------------------
'
' Purpose:
' - Run code automatically on worksheet actions.
'
' Common worksheet events (code in sheet module):
' - Worksheet_Change(ByVal Target As Range): Fires when cells change.
' - Worksheet_SelectionChange(ByVal Target As Range): Fires on selection change.
' - Worksheet_Activate(): Fires when worksheet activated.
' - Worksheet_Deactivate(): Fires when worksheet deactivated.
'
' Example Worksheet_Change:
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       If Not Intersect(Target, Me.Range("A1:A10")) Is Nothing Then
'           MsgBox "Changed cell in A1:A10"
'       End If
'   End Sub
'
' Example Worksheet_SelectionChange:
'   Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'       Me.Range("B1").Value = "Selected: " & Target.Address
'   End Sub
'
' Best practices:
' - Keep code efficient to avoid slowdowns.
' - Use Intersect to restrict to relevant ranges.
' - Prevent recursive event calls using Application.EnableEvents.
'
' -----------------------------------------
