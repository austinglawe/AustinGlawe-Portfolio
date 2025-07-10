' -----------------------------------------
' VBA File_Handling:
' FileSystemObject (FSO) basics
' -----------------------------------------
'
' Purpose:
' - Provide modern, object-oriented file and folder handling in VBA.
'
' What is FileSystemObject:
' - Part of Microsoft Scripting Runtime (scrrun.dll).
' - Exposes File, Folder, and Drive objects.
'
' How to initialize:
'
' Late binding (no reference required, portable):
'   Dim fso As Object
'   Set fso = CreateObject("Scripting.FileSystemObject")
'
' Early binding (requires Tools > References > Microsoft Scripting Runtime):
'   Dim fso As Scripting.FileSystemObject
'   Set fso = New FileSystemObject
'
' Common methods:
' - FileExists(path): check if file exists.
' - FolderExists(path): check if folder exists.
' - CreateFolder(path): create a new folder.
' - GetFile(path): return a File object.
' - GetFolder(path): return a Folder object.
'
' Examples:
'   If fso.FileExists("C:\Test\file.txt") Then MsgBox "File exists!"
'   If fso.FolderExists("C:\Test") Then MsgBox "Folder exists!"
'   fso.CreateFolder "C:\NewFolder"
'
' Best practices:
' - Use late binding for portability (avoids library dependency).
' - Use FSO for rich file/folder management tasks.
' - For quick existence checks: Dir may be faster/simpler.
'
' -----------------------------------------
