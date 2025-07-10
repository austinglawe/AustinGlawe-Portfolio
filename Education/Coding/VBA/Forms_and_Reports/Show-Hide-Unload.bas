' -----------------------------------------
' VBA UserForms_Basics:
' Showing, hiding, and unloading UserForms
' -----------------------------------------
'
' Showing a UserForm:
'   UserForm1.Show          ' Modal by default
'   UserForm1.Show vbModeless  ' Modeless (Excel usable)
'
' Hiding a UserForm:
'   Me.Hide                 ' Keeps form loaded, preserves data
'
' Unloading a UserForm:
'   Unload Me              ' Closes and frees form, resets controls
'
' Differences:
' - Show: display form
' - Hide: temporary hide, keeps form in memory
' - Unload: permanently close, frees memory
'
' Best practices:
' - Unload when done to free resources
' - Hide if you plan to reuse with existing data
' - Use modal for focused input, modeless for multitasking
'
' -----------------------------------------
