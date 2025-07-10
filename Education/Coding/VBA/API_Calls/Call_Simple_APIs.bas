' -----------------------------------------
' VBA API_Calls:
' Calling simple API functions — MessageBox
' -----------------------------------------
'
' Declare:
'   Declare PtrSafe Function MessageBox Lib "user32" Alias "MessageBoxA" ( _
'       ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long) As Long
'
' Usage:
'   Dim ret As Long
'   ret = MessageBox(0, "Hello from Windows API!", "API MessageBox", 0)
'   MsgBox "Clicked button number: " & ret
'
' uType examples:
' - 0 = OK button
' - 1 = OK + Cancel buttons
' - 16 = Critical icon
' - 32 = Question icon
' Combine with + for buttons + icons
'
' Best practices:
' - Use for advanced dialogs.
' - Check return value for user response.
' - Test on 32- and 64-bit Office.
'
' -----------------------------------------
