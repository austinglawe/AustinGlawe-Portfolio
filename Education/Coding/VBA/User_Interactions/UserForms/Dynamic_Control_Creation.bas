' -----------------------------------------
' Dynamic control creation on UserForms
' -----------------------------------------
'
' VBA allows adding controls at runtime using Controls.Add.
' This makes forms flexible and data-driven when the number/type of controls
' cannot be determined at design time.
'
' Syntax:
'   Set ctl = Me.Controls.Add("Forms.ControlType.1", "ControlName", Visible)
'
' Common ProgIDs:
'   Label          "Forms.Label.1"
'   TextBox        "Forms.TextBox.1"
'   CommandButton  "Forms.CommandButton.1"
'   CheckBox       "Forms.CheckBox.1"
'   OptionButton   "Forms.OptionButton.1"
'   ComboBox       "Forms.ComboBox.1"
'   ListBox        "Forms.ListBox.1"
'   Frame          "Forms.Frame.1"
'   Image          "Forms.Image.1"
'   ScrollBar      "Forms.ScrollBar.1"
'   SpinButton     "Forms.SpinButton.1"
'
' Example: add 3 TextBoxes dynamically
'   Private Sub UserForm_Initialize()
'       Dim i As Integer
'       Dim tb As MSForms.TextBox
'       For i = 1 To 3
'           Set tb = Me.Controls.Add("Forms.TextBox.1", "DynamicTextBox" & i, True)
'           tb.Top = 20 + (i - 1) * 30
'           tb.Left = 100
'           tb.Width = 100
'       Next i
'   End Sub
'
' Behavior:
' - Dynamically added controls are part of the Controls collection.
' - Position and size must be set manually.
' - Event handling for dynamic controls requires advanced techniques (e.g., WithEvents in a class).
'
' When to use:
' - When control count depends on runtime data.
' - When designing flexible or data-driven forms.
'
' Limitations:
' - Cannot declare event handlers directly without additional class wrapper code.
' - Manual layout and positioning required.
'
' -----------------------------------------

