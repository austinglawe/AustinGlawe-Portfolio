' -----------------------------------------
' VBA File_Handling:
' Kill and Name overview
' -----------------------------------------
'
' Purpose:
' - Delete, rename, or move files/folders natively in VBA.
'
' Kill:
' - Syntax:
'     Kill pathname
'
' - Example:
'     Kill "C:\Test\oldfile.txt"
'
' - Supports wildcards:
'     Kill "C:\Test\*.bak"
'
' - Best practice:
'     If Dir("C:\Test\oldfile.txt") <> "" Then
'         Kill "C:\Test\oldfile.txt"
'     End If
'
' - Raises error if file does not exist unless error handling is applied.
'
' Name:
' - Syntax:
'     Name oldpathname As newpathname
'
' - Rename file example:
'     Name "C:\Test\oldname.txt" As "C:\Test\newname.txt"
'
' - Move file example (same drive only):
'     Name "C:\Test\file.txt" As "C:\Archive\file.txt"
'
' - Rename folder example:
'     Name "C:\OldFolder" As "C:\NewFolder"
'
' Notes:
' - Name cannot move files across different drives.
' - Check existence with Dir or use error handling.
'
' Best practices:
' - Use Kill for simple file deletion (wrap in existence check or error handler).
' - Use Name for renaming or same-drive moves.
' - For advanced/cross-drive moves: prefer FileSystemObject.
'
' -----------------------------------------

