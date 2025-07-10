' -----------------------------------------
' VBA API_Calls:
' Window management — focus, size, positioning
' -----------------------------------------
'
' Key API functions:
' - FindWindow: get window handle by class/title
' - SetForegroundWindow: bring window to front
' - SetWindowPos: move/resize window
' - ShowWindow: show, hide, minimize window
'
' Declarations (64-bit safe):
'   Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (...)
'   Declare PtrSafe Function SetForegroundWindow Lib "user32" (...)
'   Declare PtrSafe Function SetWindowPos Lib "user32" (...)
'   Declare PtrSafe Function ShowWindow Lib "user32" (...)
'
' Example:
'   hwnd = FindWindow("Notepad", vbNullString)
'   If hwnd <> 0 Then
'       SetForegroundWindow hwnd
'       SetWindowPos hwnd, 0, 100, 100, 800, 600, 0
'       ShowWindow hwnd, 5
'   End If
'
' ShowWindow constants:
' - 0 = Hide
' - 5 = Show
' - 6 = Minimize
'
' Best practices:
' - Validate hwnd before use
' - Use accurate window class/title in FindWindow
' - Avoid stealing user focus unexpectedly
' - Test across Windows versions
'
' -----------------------------------------
