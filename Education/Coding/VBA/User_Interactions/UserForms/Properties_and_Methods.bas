' -----------------------------------------
' UserForm properties and methods overview
' -----------------------------------------
'
' Common properties:
'
' Name
' - The internal name used to reference the form in VBA code.
' - Best practice: rename from default (e.g., UserForm1) to something descriptive.
'
' Caption
' - Text displayed in the UserForm title bar.
'
' Height / Width
' - Size of the UserForm in points.
'
' Top / Left
' - Position of UserForm relative to screen.
'
' Enabled
' - True/False: whether the UserForm responds to interaction.
'
' Visible
' - True/False: whether the UserForm is currently displayed.
'
' Tag
' - Free-form string for storing custom metadata or temporary data.
'
' ShowModal
' - True/False: determines if form is modal (default True) or modeless.
'
' Common methods:
'
' Show
' - Displays the UserForm.
' - Example: UserForm1.Show
' - Example: UserForm1.Show vbModeless
'
' Hide
' - Hides the UserForm but keeps it loaded in memory.
'
' Unload
' - Completely unloads the UserForm from memory.
' - Example: Unload UserForm1
'
' Usage notes:
' - Modal UserForm blocks interaction with Excel until closed.
' - Modeless UserForm allows Excel to remain usable while form is open.
' - Hide keeps state intact (user input stays on form).
' - Unload clears all state and triggers UserForm_Terminate event.
'
' -----------------------------------------

