' -----------------------------------------
' Advanced UserForm customization:
' MultiPage advanced usage
' -----------------------------------------
'
' Purpose:
' - Organize large forms into tabbed sections.
'
' Key properties:
' - MultiPage.Pages: collection of Page objects.
' - MultiPage.Value: index of currently active page (0-based).
' - Page.Caption: tab label text.
'
' Example: programmatically switch page:
'   MultiPage1.Value = 1  ' Switch to page 2
'
' Adding controls at runtime to a specific page:
'   MultiPage1.Pages(0).Controls.Add "Forms.TextBox.1", "txtDynamic", True
'
' Handling Change event (fires when user switches tabs):
'   Private Sub MultiPage1_Change()
'       MsgBox "User switched to page " & (MultiPage1.Value + 1)
'   End Sub
'
' Example: dynamically populate a page when selected:
'   Private Sub MultiPage1_Change()
'       Dim pgIndex As Integer
'       pgIndex = MultiPage1.Value
'
'       If pgIndex = 1 Then
'           If MultiPage1.Pages(1).Controls.Count = 0 Then
'               Dim txt As MSForms.TextBox
'               Set txt = MultiPage1.Pages(1).Controls.Add("Forms.TextBox.1", "txtOnPage2", True)
'               txt.Top = 20
'               txt.Left = 20
'               txt.Width = 100
'               txt.Value = "Dynamically added on page 2"
'           End If
'       End If
'   End Sub
'
' Notes:
' - Page indices are 0-based.
' - Pages can be hidden by setting Page.Visible = False.
' - Pages cannot be removed at runtime, only hidden.
'
' Best practices:
' - Use MultiPage to simplify UI layout for forms with many fields.
' - Populate controls dynamically on-demand for performance optimization.
'
' -----------------------------------------
