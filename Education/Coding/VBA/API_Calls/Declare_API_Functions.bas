' -----------------------------------------
' VBA API_Calls:
' Declaring Windows API functions with PtrSafe
' -----------------------------------------
'
' Use PtrSafe for 64-bit compatibility:
'   Declare PtrSafe Function MessageBox Lib "user32" Alias "MessageBoxA" ( _
'       ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long) As Long
'
' Notes:
' - PtrSafe indicates 64-bit safe declaration.
' - Use LongPtr for pointers/handles.
' - Strings usually passed ByVal.
' - Lib specifies DLL containing function.
' - Alias used if function name differs.
'
' Conditional compilation for 32/64-bit:
'   #If VBA7 Then
'       Declare PtrSafe Function ...
'   #Else
'       Declare Function ...
'   #End If
'
' Best practices:
' - Always use PtrSafe in modern VBA.
' - Use LongPtr for pointers/handles.
' - Wrap with conditional compilation for compatibility.
' - Comment declarations.
'
' -----------------------------------------
