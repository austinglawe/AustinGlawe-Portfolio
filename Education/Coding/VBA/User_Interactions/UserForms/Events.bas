' -----------------------------------------
' UserForm events overview
' -----------------------------------------
'
' UserForms support built-in events that allow code to run automatically
' when certain actions occur on the form or its controls.
'
' Common UserForm-level events:
'
' Initialize
' - Fires when the form is first created (before showing).
' - Used for setup (e.g., populating ComboBoxes, setting defaults).
' - Example:
'     Private Sub UserForm_Initialize()
'         ComboBox1.AddItem "Option 1"
'         ComboBox1.AddItem "Option 2"
'     End Sub
'
' Activate
' - Fires every time the form becomes active.
'
' Deactivate
' - Fires when the form loses focus.
'
' QueryClose
' - Fires when user attempts to close the form (e.g., clicking "X").
' - Allows cancellation of close by setting Cancel = True.
' - Example:
'     Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'         If MsgBox("Are you sure you want to close?", vbYesNo) = vbNo Then
'             Cancel = True
'         End If
'     End Sub
'
' Terminate
' - Fires when the form is fully unloaded from memory (after Unload).
'
' Control-specific events:
' - Example:
'     CommandButton1_Click — Fires when button clicked.
'     TextBox1_Change — Fires when text changes.
'     ComboBox1_Change — Fires when selection changes.
'
' Notes:
' - Initialize fires only once per instance.
' - Activate/Deactivate can fire multiple times as user switches focus.
' - QueryClose allows confirmation before closing.
' - Terminate is final cleanup after Unload.
'
' Best practices:
' - Use Initialize for preparing form contents.
' - Use QueryClose for close confirmation logic.
' - Handle individual control events for user interactions.
'
' -----------------------------------------

