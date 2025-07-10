' -----------------------------------------
' Advanced UserForm customization:
' Removing the close button ("X") from title bar
' -----------------------------------------
'
' Purpose:
' - Prevent user from closing UserForm via the "X" button while retaining title bar visibility.
'
' API declarations:
'   Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
'   Private Declare PtrSafe Function GetWindowLongPtrA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
'   Private Declare PtrSafe Function SetWindowLongPtrA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
'
' Constants:
'   Private Const GWL_STYLE As Long = -16
'   Private Const WS_SYSMENU As Long = &H80000
'
' Example procedure:
'   Sub RemoveCloseButton(uf As Object)
'       Dim hWnd As LongPtr
'       hWnd = FindWindowA(vbNullString, uf.Caption)
'       If hWnd <> 0 Then
'           Dim lStyle As LongPtr
'           lStyle = GetWindowLongPtrA(hWnd, GWL_STYLE)
'           lStyle = lStyle And Not WS_SYSMENU
'           SetWindowLongPtrA hWnd, GWL_STYLE, lStyle
'       End If
'   End Sub
'
' Usage:
' - Call RemoveCloseButton(Me) from UserForm_Initialize.
'
' Notes:
' - Title bar remains visible (can still drag form).
' - "X" button is disabled (greyed out).
' - For additional safety, trap UserForm_QueryClose to block closure:
'     Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'         Cancel = True
'     End Sub
'
' -----------------------------------------
