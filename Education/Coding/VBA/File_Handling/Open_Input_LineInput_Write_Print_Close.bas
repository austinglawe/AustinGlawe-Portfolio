' -----------------------------------------
' VBA File_Handling:
' Open/Input/Line Input/Write/Print/Close overview
' -----------------------------------------
'
' Purpose:
' - Read and write text files using native VBA (no external libraries required).
'
' Syntax:
'   Open pathname For mode As #filenumber
'
' Modes:
' - For Input: read-only.
' - For Output: write (overwrite if exists).
' - For Append: add to end of file.
' - For Binary: binary read/write.
' - For Random: record-based access.
'
' Key statements:
' - Input: read comma-separated values.
' - Line Input: read a single text line.
' - Print: write plain text (adds CR/LF automatically).
' - Write: write values with delimiters (adds quotes/commas).
' - Close: close file handle.
'
' Examples:
'
' 1. Read file line by line:
'   Dim fileNum As Integer
'   Dim lineText As String
'
'   fileNum = FreeFile
'   Open "C:\Test\sample.txt" For Input As #fileNum
'
'   Do While Not EOF(fileNum)
'       Line Input #fileNum, lineText
'       Debug.Print lineText
'   Loop
'
'   Close #fileNum
'
' 2. Write new file (overwrite if exists):
'   fileNum = FreeFile
'   Open "C:\Test\output.txt" For Output As #fileNum
'   Print #fileNum, "Hello, world!"
'   Print #fileNum, "Line 2"
'   Close #fileNum
'
' 3. Append to file:
'   fileNum = FreeFile
'   Open "C:\Test\output.txt" For Append As #fileNum
'   Print #fileNum, "Appended line"
'   Close #fileNum
'
' 4. Read comma-separated values:
'   Dim name As String
'   Dim age As Integer
'
'   fileNum = FreeFile
'   Open "C:\Test\data.txt" For Input As #fileNum
'   Input #fileNum, name, age
'   Close #fileNum
'
' Best practices:
' - Always use FreeFile to obtain next available file number.
' - Always close files after use.
' - Line Input reads whole line as string.
' - Input reads comma-separated values into variables.
'
' -----------------------------------------
