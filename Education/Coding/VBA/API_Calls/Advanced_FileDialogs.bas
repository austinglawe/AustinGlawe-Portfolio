' -----------------------------------------
' VBA API_Calls:
' Using API for advanced file dialogs
' -----------------------------------------
'
' Advantages:
' - More customization (filters, multi-select, default folders)
' - Native Windows dialogs with full features
'
' Key APIs:
' - GetOpenFileName, GetSaveFileName from comdlg32.dll
'
' Declarations:
' #If VBA7 Then
'   Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
' #Else
'   Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
' #End If
'
' Requires:
' - Defining OPENFILENAME structure in VBA
' - Careful buffer and memory management
'
' Best practices:
' - Use only if VBA FileDialog lacks needed features
' - Handle strings/buffers carefully to avoid crashes
' - Test on all target Office and Windows bitness versions
'
' -----------------------------------------
