' -----------------------------------------
' VBA UserForms_AdvancedControls:
' Complex user interactions and validation
' -----------------------------------------
'
' Importance of validation:
' - Ensures valid, complete data
' - Prevents errors
' - Improves UX
'
' Validation types:
' - Required fields
' - Data type (IsNumeric, IsDate)
' - Range checks
' - Cross-field validation
'
' Example: Required field
'   Private Sub cmdSubmit_Click()
'       If Trim(Me.txtName.Value) = "" Then
'           MsgBox "Name is required.", vbExclamation
'           Me.txtName.SetFocus
'           Exit Sub
'       End If
'       Unload Me
'   End Sub
'
' Real-time feedback example:
'   Private Sub txtEmail_Change()
'       If IsValidEmail(Me.txtEmail.Value) Then
'           Me.lblEmailStatus.Caption = "Valid"
'           Me.lblEmailStatus.ForeColor = vbGreen
'       Else
'           Me.lblEmailStatus.Caption = "Invalid"
'           Me.lblEmailStatus.ForeColor = vbRed
'       End If
'   End Sub
'
' Best practices:
' - Validate on submit and optionally on change
' - Provide clear error messages
' - Modular validation code
' - Use SetFocus to guide correction
'
' -----------------------------------------
