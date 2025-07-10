' -----------------------------------------
' Advanced UserForm customization:
' Center UserForm over Excel application window
' -----------------------------------------
'
' Purpose:
' - By default, UserForms center on screen.
' - This technique centers a UserForm over the Excel window itself.
'
' API declarations:
'   Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
'   Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
'
' Type declaration:
'   Private Type RECT
'       Left As Long
'       Top As Long
'       Right As Long
'       Bottom As Long
'   End Type
'
' Example procedure:
'   Sub CenterUserFormOverExcel(uf As Object)
'       Dim hWnd As LongPtr
'       Dim rect As RECT
'       Dim formWidth As Long
'       Dim formHeight As Long
'
'       hWnd = FindWindowA("XLMAIN", Application.Caption)
'       If hWnd = 0 Then Exit Sub
'
'       GetWindowRect hWnd, rect
'
'       formWidth = uf.Width * (96 / 72)  ' Approximate conversion from points to pixels
'       formHeight = uf.Height * (96 / 72)
'
'       uf.Left = ((rect.Right + rect.Left) / 2) - (formWidth / 2)
'       uf.Top = ((rect.Bottom + rect.Top) / 2) - (formHeight / 2)
'   End Sub
'
' Usage:
' - Call CenterUserFormOverExcel(Me) from UserForm_Initialize.
'
' Notes:
' - Handles Excel window position dynamically.
' - Conversion assumes 96 DPI screen; may require adjustment for high-DPI.
'
' -----------------------------------------
