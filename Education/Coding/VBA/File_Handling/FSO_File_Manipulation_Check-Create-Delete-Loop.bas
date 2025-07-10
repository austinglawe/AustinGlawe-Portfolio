' -----------------------------------------
' VBA File_Handling:
' FileSystemObject folder manipulation
' -----------------------------------------
'
' Purpose:
' - Work with folders/directories using FileSystemObject.
'
' Initialization:
'   Dim fso As Object
'   Set fso = CreateObject("Scripting.FileSystemObject")
'
' Check if folder exists:
'   If fso.FolderExists("C:\Test") Then MsgBox "Folder exists!"
'
' Create folder:
'   If Not fso.FolderExists("C:\Test\NewFolder") Then
'       fso.CreateFolder "C:\Test\NewFolder"
'   End If
'
' Delete folder:
'   If fso.FolderExists("C:\Test\OldFolder") Then
'       fso.DeleteFolder "C:\Test\OldFolder"
'   End If
'
' Get folder object:
'   Dim fld As Object
'   Set fld = fso.GetFolder("C:\Test")
'
' Folder properties:
'   Debug.Print fld.Name
'   Debug.Print fld.Path
'   Debug.Print fld.DateCreated
'   Debug.Print fld.Size
'
' Loop through files in folder:
'   Dim fil As Object
'   For Each fil In fld.Files
'       Debug.Print fil.Name
'   Next fil
'
' Loop through subfolders:
'   Dim sf As Object
'   For Each sf In fld.SubFolders
'       Debug.Print sf.Name
'   Next sf
'
' Best practices:
' - Always check FolderExists before using.
' - CreateFolder only creates one folder level (not recursive).
' - Use Folder object for metadata and child collections (Files, SubFolders).
'
' -----------------------------------------
