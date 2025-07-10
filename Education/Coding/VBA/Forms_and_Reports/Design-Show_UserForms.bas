' -----------------------------------------
' VBA UserForms_Basics:
' Designing and showing UserForms
' -----------------------------------------
'
' What is a UserForm?
' - Custom dialog box for user interaction.
'
' Designing:
' - Insert > UserForm in VBA Editor.
' - Use Toolbox to add controls.
' - Set control properties (Name, Caption, etc.).
'
' Showing a UserForm:
'   Sub ShowForm()
'       UserForm1.Show
'   End Sub
'
' Closing a UserForm:
' - From form code: Unload Me  ' Completely closes
' - Or Me.Hide              ' Hides form, preserves state
'
' Modal vs Modeless:
' - UserForm.Show (modal) — blocks Excel until closed.
' - UserForm.Show vbModeless — allows switching.
'
' Best practices:
' - Use clear, descriptive names.
' - Design intuitive layouts.
' - Use code behind for validation and events.
'
' -----------------------------------------
