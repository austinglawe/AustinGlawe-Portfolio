' -----------------------------------------
' VBA UserForms_AdvancedControls:
' Multi-page UserForms and navigation
' -----------------------------------------
'
' MultiPage control:
' - Adds tabbed pages for organizing form content
'
' Initialize pages:
'   Me.MultiPage1.Pages(0).Caption = "Personal Info"
'   Me.MultiPage1.Pages.Add
'   Me.MultiPage1.Pages(1).Caption = "Settings"
'
' Navigate between pages:
'   Me.MultiPage1.Value = 1  ' Switch to second page
'
' Enable/disable or hide pages:
'   Me.MultiPage1.Pages(1).Enabled = False
'   Me.MultiPage1.Pages(1).Visible = False
'
' Navigation buttons example:
'   Private Sub cmdNext_Click()
'       If Me.MultiPage1.Value < Me.MultiPage1.Pages.Count - 1 Then
'           Me.MultiPage1.Value = Me.MultiPage1.Value + 1
'       End If
'   End Sub
'
'   Private Sub cmdPrev_Click()
'       If Me.MultiPage1.Value > 0 Then
'           Me.MultiPage1.Value = Me.MultiPage1.Value - 1
'       End If
'   End Sub
'
' Best practices:
' - Group pages logically.
' - Provide clear navigation.
' - Validate input before next page.
' - Use consistent naming and design.
'
' -----------------------------------------
