' -----------------------------------------
' Advanced UserForm customization:
' Preventing UserForm resizing
' -----------------------------------------
'
' Purpose:
' - VBA UserForms are non-resizable by default.
' - If resizing has been enabled (e.g., via WS_THICKFRAME style), you can use Windows API to lock resizing again.
'
' API declarations:
'   Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
'   Private Declare PtrSafe Function GetWindowLongPtrA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
'   Private Declare PtrSafe Function SetWindowLongPtrA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
'
' Constants:
'   Private Const GWL_STYLE As Long = -16
'   Private Const WS_THICKFRAME As Long = &H40000
'
' Example procedure:
'   Sub PreventResizing(uf As Object)
'       Dim hWnd As LongPtr
'       hWnd = FindWindowA(vbNullString, uf.Caption)
'       If hWnd <> 0 Then
'           Dim lStyle As LongPtr
'           lStyle = GetWindowLongPtrA(hWnd, GWL_STYLE)
'           lStyle = lStyle And Not WS_THICKFRAME
'           SetWindowLongPtrA hWnd, GWL_STYLE, lStyle
'       End If
'   End Sub
'
' Usage:
' - Call PreventResizing(Me) from UserForm_Initialize.
'
' Notes:
' - VBA UserForms are fixed-size by default unless WS_THICKFRAME is applied via API.
' - This technique explicitly removes the resizing style if present.
'
' -----------------------------------------
