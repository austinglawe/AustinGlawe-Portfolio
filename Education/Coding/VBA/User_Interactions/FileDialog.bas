' -----------------------------------------
' FileDialog overview, usage, and guidance
' -----------------------------------------
'
' The FileDialog object allows VBA to show standard Windows dialogs so users can select files or folders.
' It has four main dialog types:
'
' 1️⃣ msoFileDialogFilePicker:
'     Use this when you want the user to pick one or more existing files.
'     It’s a general-purpose dialog that doesn’t mimic Excel’s native Open screen.
'     Best for scenarios like “Please choose a file to import.”
'
' 2️⃣ msoFileDialogOpen:
'     Looks and behaves like Excel’s built-in File > Open dialog.
'     Use this if you want a consistent Office-style user experience for opening files,
'     possibly including access to recent files and shared locations (OneDrive, etc.).
'
' 3️⃣ msoFileDialogSaveAs:
'     Looks and behaves like Excel’s File > Save As dialog.
'     Use this when you want the user to choose a path and name for saving a file.
'     Important: this dialog does NOT actually save the file itself — you must handle that in your code.
'
' 4️⃣ msoFileDialogFolderPicker:
'     Used when you want the user to pick a folder rather than a file.
'     Perfect for workflows like selecting a folder to save multiple files or process an entire directory.
'
' General guidance:
' - Use FilePicker for general file selection when a simple UI is sufficient.
' - Use Open for an Office-style Open dialog if consistency matters.
' - Use SaveAs when you need the user to provide a file path for saving.
' - Use FolderPicker when a folder path is needed.
'
' -----------------------------------------
' Key properties:
'
' .Title
'     Sets the window title for context (always recommended).
'
' .AllowMultiSelect
'     True/False (default False). Allows selecting multiple files (useful with FilePicker/Open).
'
' .InitialFileName
'     Sets the starting folder OR suggested file name (especially useful for SaveAs).
'
' .Filters
'     Collection used to filter displayed file types.
'     Use .Filters.Clear to clear defaults.
'     Use .Filters.Add to add filters (e.g., "*.xlsx", "*.txt").
'
' .FilterIndex
'     Determines which filter appears first (1-based index).
'
' .ButtonName
'     Allows customizing the text of the action button (e.g., “Import” instead of “Open”).
'
' .SelectedItems
'     Collection of selected file(s) or folder(s) after dialog closes (1-based index).
'     Only populated if user clicks OK (i.e., if .Show = -1).
'
' -----------------------------------------
' Main method:
'
' .Show
'     Displays the dialog.
'     Returns -1 if OK clicked, 0 if Cancel clicked.
'
' Example workflow:
'     Dim fd As FileDialog
'     Dim filePath As String
'
'     Set fd = Application.FileDialog(msoFileDialogFilePicker)
'     With fd
'         .Title = "Select a report file"
'         .AllowMultiSelect = False
'         .InitialFileName = "C:\Reports"
'         .Filters.Clear
'         .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
'         .Filters.Add "All Files", "*.*"
'
'         If .Show = -1 Then
'             filePath = .SelectedItems(1)
'             MsgBox "You selected: " & filePath
'         Else
'             MsgBox "No file selected."
'         End If
'     End With
'
' -----------------------------------------
' Summary:
' The FileDialog object gives VBA a professional, native dialog experience.
' Remember:
' - Pick the right dialog type for your scenario.
' - Set .Title and use .Filters for user-friendly experience.
' - Always check .Show result (-1 = OK, 0 = Cancel).
' - After .Show, get file(s) or folder from .SelectedItems(1).
'
' -----------------------------------------
