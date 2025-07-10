' -----------------------------------------
' VBA File_Handling overview
' -----------------------------------------
'
' Purpose:
' - Automate reading, writing, managing files and folders from VBA.
'
' Two primary approaches:
'
' 1. Native VBA file handling:
' - Dir: list files/folders, check existence.
' - Open: open files for reading/writing.
' - Input, Line Input: read file contents.
' - Write, Print: write to files.
' - Close: close file handles.
' - Kill: delete files.
' - Name: rename or move files.
'
' 2. FileSystemObject (FSO) approach:
' - Part of Microsoft Scripting Runtime.
' - Object-oriented API for files and folders.
'
'   Common FSO methods/properties:
'   - FileSystemObject.GetFile, GetFolder.
'   - FileSystemObject.CreateTextFile, OpenTextFile.
'   - FileSystemObject.FileExists, FolderExists.
'   - FileSystemObject.DeleteFile, DeleteFolder.
'
' When to use:
' - Use native VBA for lightweight simple tasks.
' - Use FSO for rich file/folder manipulation.
'
' Example scenarios:
' - Check if file exists.
' - Read/write text files.
' - Iterate through files in a folder.
' - Create folders dynamically.
'
' -----------------------------------------
