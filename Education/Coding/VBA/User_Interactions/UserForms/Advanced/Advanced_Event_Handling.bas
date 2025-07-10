' -----------------------------------------
' Advanced UserForm customization:
' Hiding the title bar
' -----------------------------------------
'
' Purpose:
' - VBA UserForms do not provide a built-in property to hide the title bar.
' - Use Windows API to remove WS_CAPTION style to hide the title bar.
'
' API declarations:
'   Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
'   Private Declare PtrSafe Function GetWindowLongPtrA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
'   Private Declare PtrSafe Function SetWindowLongPtrA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
'
' Constants:
'   Private Const GWL_STYLE As Long = -16
'   Private Const WS_CAPTION As Long = &HC00000
'
' Example procedure:
'   Sub HideTitleBar(uf As Object)
'       Dim hWnd As LongPtr
'       hWnd = FindWindowA(vbNullString, uf.Caption)
'       If hWnd <> 0 Then
'           Dim lStyle As LongPtr
'           lStyle = GetWindowLongPtrA(hWnd, GWL_STYLE)
'           lStyle = lStyle And Not WS_CAPTION
'           SetWindowLongPtrA hWnd, GWL_STYLE, lStyle
'       End If
'   End Sub
'
' Usage:
' - Call HideTitleBar(Me) from UserForm_Initialize.
'
' Notes:
' - Removes both title bar and close button.
' - Prevents dragging unless coded separately.
' - Provide your own Close button for usability.
'
' -----------------------------------------
