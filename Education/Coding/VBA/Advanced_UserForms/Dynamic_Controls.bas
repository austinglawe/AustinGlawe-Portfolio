' -----------------------------------------
' VBA UserForms_AdvancedControls:
' Dynamic control manipulation and custom events
' -----------------------------------------
'
' Create controls dynamically:
'   Set txtDynamic = Me.Controls.Add("Forms.TextBox.1", "txtDynamic1", True)
'   With txtDynamic
'       .Left = 20
'       .Top = 50
'       .Width = 100
'       .Text = "Dynamic TextBox"
'   End With
'
' Handling events for dynamic controls using class module:
' - Create class module clsTextBoxEvents
'   Public WithEvents txtBox As MSForms.TextBox
'   Private Sub txtBox_Change()
'       MsgBox "Text changed: " & txtBox.Text
'   End Sub
'
' - In UserForm:
'   Dim colTextBoxes As Collection
'   Private Sub UserForm_Initialize()
'       Set colTextBoxes = New Collection
'       Set newTxt = Me.Controls.Add("Forms.TextBox.1", "txtDynamic1", True)
'       Set txtEvent = New clsTextBoxEvents
'       Set txtEvent.txtBox = newTxt
'       colTextBoxes.Add txtEvent
'   End Sub
'
' Best practices:
' - Keep event handlers in collection to prevent loss.
' - Use clear, unique control names.
' - Clean up dynamic controls properly.
' - Document event logic well.
'
' -----------------------------------------
