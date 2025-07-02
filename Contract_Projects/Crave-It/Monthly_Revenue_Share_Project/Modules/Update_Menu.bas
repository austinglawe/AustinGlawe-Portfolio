Sub Update_MealsLookup()
' Last Updated: 2025.07.02

Dim wbMacro As Workbook
Dim CorrectPassword As String
Dim Password As String
Dim fdMeals As FileDialog
Dim FilePath As String
Dim wbTemp As Workbook
Dim wsTemp As Worksheet
Dim Headers As Variant
Dim HeaderMismatch As Boolean
Dim LastRowMenu As Long
Dim wsMealsLookup As Worksheet
Dim ws As Worksheet
Dim CodeLines As Variant
Dim TotalLines As Long
Dim MaxLinesPerSub As Long
Dim SubCount As Long
Dim LineIndex As Long
Dim CodeMod As Object
Dim j As Long


' Set the 'wbMacro' variable to this workbook.
    Set wbMacro = ThisWorkbook

' First ask the user for a password - Store the correct password in a variable called 'CorrectPassword'
    CorrectPassword = "Test123"

    ' Prompt user for password
        Password = InputBox("Enter the password to use this macro:", "Authorization Required")

    ' If incorrect password, exit macro
        If Password <> CorrectPassword Then
            MsgBox "Incorrect password. Access denied.", vbCritical
            Exit Sub
        End If

' Ask the user to select the Meals Lookup file
    Set fdMeals = Application.FileDialog(msoFileDialogFilePicker)
    
    With fdMeals
        .Title = "Select the Meals Lookup File"
        .Filters.Clear
    ' If operation was canceled, give the user a message and exit the sub.
        If .Show <> -1 Then
            MsgBox "No file selected. Please locate the correct file and try again.", vbExclamation, "No File Selected"
            Exit Sub
        End If
    ' Store the selected file in the variable 'FilePath'
        FilePath = .SelectedItems(1)
    End With

    ' Open the selected Meals Lookup file
        Set wbTemp = Workbooks.Open(FilePath)
        Set wsTemp = wbTemp.Sheets(1)
    
    ' Check to make sure the meals lookup file is correct.
        ' Create a variable to hold what all of the column headers should be from the raw report. Store it in a vairable called 'Headers'.
        ' (Columns A:IC)
            Headers = Array("School ID", "School Name", "ID", "Ordered", "Item Preface", " ""Item Name""", "Gluten Free", "Kosher", "Organic", "Veggie", "Dairy Free", "Soy Free", "Contains Nuts", "Contains Shellfish", "Vegan", "Egg Free", "Spicy", "Contains Fish", "Description", "Set 1 Options Label", "Set 1 Option 1", "Set 1 Option 2", "Set 1 Option 3", "Set 1 Option 4", "Set 1 Option 5", "Set 1 Option 6", "Set 1 Option 7", _
                "Set 1 Option 8", "Set 1 Option 9", "Set 1 Option 10", "Set 1 Option 11", "Set 1 Option 12", "Set 1 Option 13", "Set 1 Option 14", "Set 1 Option 15", "Set 1 Option 16", "Set 1 Option 17", "Set 1 Option 18", "Set 1 Option 19", "Set 1 Option 20", "Set 1 Option 21", "Set 1 Option 22", "Set 1 Option 23", "Set 1 Option 24", "Set 1 Option 25", "Set 1 Option 26", "Set 1 Option 27", _
                "Set 1 Option 28", "Set 1 Option 29", "Set 1 Option 30", "Set 1 Option MIN", "Set 1 Option MAX", "Set 2 Options Label", "Set 2 Option 1", "Set 2 Option 2", "Set 2 Option 3", "Set 2 Option 4", "Set 2 Option 5", "Set 2 Option 6", "Set 2 Option 7", "Set 2 Option 8", "Set 2 Option 9", "Set 2 Option 10", "Set 2 Option 11", "Set 2 Option 12", "Set 2 Option 13", "Set 2 Option 14", "Set 2 Option 15", _
                "Set 2 Option 16", "Set 2 Option 17", "Set 2 Option 18", "Set 2 Option 19", "Set 2 Option 20", "Set 2 Option 21", "Set 2 Option 22", "Set 2 Option 23", "Set 2 Option 24", "Set 2 Option 25", "Set 2 Option 26", "Set 2 Option 27", "Set 2 Option 28", "Set 2 Option 29", "Set 2 Option 30", "Set 2 Option MIN", "Set 2 Option MAX", "Set 3 Options Label", "Set 3 Option 1", "Set 3 Option 2", _
                "Set 3 Option 3", "Set 3 Option 4", "Set 3 Option 5", "Set 3 Option 6", "Set 3 Option 7", "Set 3 Option 8", "Set 3 Option 9", "Set 3 Option 10", "Set 3 Option 11", "Set 3 Option 12", "Set 3 Option 13", "Set 3 Option 14", "Set 3 Option 15", "Set 3 Option 16", "Set 3 Option 17", "Set 3 Option 18", "Set 3 Option 19", "Set 3 Option 20", "Set 3 Option 21", "Set 3 Option 22", "Set 3 Option 23", _
                "Set 3 Option 24", "Set 3 Option 25", "Set 3 Option 26", "Set 3 Option 27", "Set 3 Option 28", "Set 3 Option 29", "Set 3 Option 30", "Set 3 Option MIN", "Set 3 Option MAX", "Set 4 Options Label", "Set 4 Option 1", "Set 4 Option 2", "Set 4 Option 3", "Set 4 Option 4", "Set 4 Option 5", "Set 4 Option 6", "Set 4 Option 7", "Set 4 Option 8", "Set 4 Option 9", "Set 4 Option 10", "Set 4 Option 11", _
                "Set 4 Option 12", "Set 4 Option 13", "Set 4 Option 14", "Set 4 Option 15", "Set 4 Option 16", "Set 4 Option 17", "Set 4 Option 18", "Set 4 Option 19", "Set 4 Option 20", "Set 4 Option 21", "Set 4 Option 22", "Set 4 Option 23", "Set 4 Option 24", "Set 4 Option 25", "Set 4 Option 26", "Set 4 Option 27", "Set 4 Option 28", "Set 4 Option 29", "Set 4 Option 30", "Set 4 Option MIN", _
                "Set 4 Option MAX", "Set 5 Options Label", "Set 5 Option 1", "Set 5 Option 2", "Set 5 Option 3", "Set 5 Option 4", "Set 5 Option 5", "Set 5 Option 6", "Set 5 Option 7", "Set 5 Option 8", "Set 5 Option 9", "Set 5 Option 10", "Set 5 Option 11", "Set 5 Option 12", "Set 5 Option 13", "Set 5 Option 14", "Set 5 Option 15", "Set 5 Option 16", "Set 5 Option 17", "Set 5 Option 18", "Set 5 Option 19", _
                "Set 5 Option 20", "Set 5 Option 21", "Set 5 Option 22", "Set 5 Option 23", "Set 5 Option 24", "Set 5 Option 25", "Set 5 Option 26", "Set 5 Option 27", "Set 5 Option 28", "Set 5 Option 29", "Set 5 Option 30", "Set 5 Option MIN", "Set 5 Option MAX", "Item Type", "Is this a pizza (per slice)?", "Price A", "Price A Label", "Price for 2 Slices", "Price for 3 Slices", "Use Price B?", _
                "Price B", "Price B Label", "Use Price C?", "Price C", "Price C Label", "Choice of Drink", "Drink: Select up to #", "Drink: Select at least #", "Drink Label", "Choice of side", "Side: Select up to #", "Side: Select at least #", "Side Label", "A la carte only?", "Cannot order without entree or combo", "Do not offer a la carte", "Cannot order without at least one of", _
                "Free/Reduced for Students?", "Reduced Student Price", "Free/Reduced for Staff?", "Reduced Staff Price", "Hide item from menu?", "Item Code", "Report Code", "URL", "Vendor Name", "Item Cost A:1", "Item Cost A:2", "Item Cost A:3", "Item Cost B:1", "Item Cost B:2", "Item Cost B:3", "Item Cost C:1", "Item Cost C:2", "Item Cost C:3", "Item Cost: Pizza/2 Slices", _
                "Item Cost: Pizza/3 Slices", "Can Order in Qty?", "Min Qty", "Max Qty", "Available at Check-In", "Enable Block Ordering", "Inventory Restriction", "Max Allowed", "Message Display at #", "One Time or Daily")
                        
        ' Check the column headers match with the position they should in the raw report.
            ' Create a variable called 'HeaderMismatch' for a Boolean value in the lookup.
                HeaderMismatch = False
            ' Create a For loop to make sure each column header is the correct position.
                For i = LBound(Headers) To UBound(Headers)
                ' Compare each header to the value in Row 1 of the worksheet
                    If wsTemp.Cells(1, i + 1).Value <> Headers(i) Then
                ' If any header does not match, change 'HeaderMismatch' to true and jump out of the For loop
                        HeaderMismatch = True
                        Exit For
                    End If
                Next i

            ' If any column headers do not match, provide a message to the user and close the file they selected.
                If HeaderMismatch Then
                    MsgBox "The selected 'Menu_Library' file does not appear to be the correct file. Please locate the correct file and try again.", vbCritical, "Mismatching Column Headers"
                    wbTemp.Close SaveChanges:=False
                    Exit Sub
                End If
            
            ' If they all match, then populate 5 columns at the end of the data set.
                ' Create the columns Headers (ID:IH)
                    wsTemp.Range("ID1").Value = "Menu Item [ID]"
                    wsTemp.Range("IE1").Value = "Lookup - [School Name] | [Item Code]"
                    wsTemp.Range("IF1").Value = "Item Type"
                    wsTemp.Range("IG1").Value = "Price [A]"
                    wsTemp.Range("IH1").Value = "Duplicates with mismatching prices"
                
                ' Add in the formulas (ID:IH)
                    wsTemp.Range("ID2").Formula = "=C2"
                    wsTemp.Range("IE2").Formula = "=TRIM(B2&"" | ""&HF2)"
                    wsTemp.Range("IF2").Formula = "=IF(GC2=""E"",""Entree"",IF(GC2=""S"",""Side"",IF(GC2=""D"",""Drink"",IF(GC2=""O"",""Other"",""Check""))))"
                    wsTemp.Range("IG2").Formula = "=GE2"
                    wsTemp.Range("IH2").Formula = "=IF(COUNT(UNIQUE(FILTER(IG:IG,IE:IE=IE2)))>1, ""Check"","""")"
                    
                ' Find the last row of the worksheet and store it in a variable called 'LastRowMenu'
                    LastRowMenu = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
                    
                ' Fill Formulas Down (ID:IH)
                    wsTemp.Range("ID2:IH" & LastRowMenu).FillDown
                    
                ' Copy and Paste the values only (ID:IG)
                    wsTemp.Range("ID:IH").Copy
                    wsTemp.Range("ID:IH").PasteSpecial xlPasteValues
                    
                ' Get rid of the clipboard
                    Application.CutCopyMode = False
                    
                ' Delete all columns before column "ID" - Columns ID:IH will become A:E
                    wsTemp.Columns("A:IC").Delete
                    
                ' Sort data by 'Lookup - [School Name] | [Item Code]' (Column B), then sort it by 'Price [A]' (Column D)
                ' Note: VBA sorting is hierarchical, meaning it applies the first sort field, then within those results applies the second, and so on
                    With wsTemp.Sort
                        .SortFields.Clear
                    ' Sort A-Z for Column B, then lowest to highest price (Column D)
                        .SortFields.Add Key:=wsTemp.Range("B2:B" & LastRowMenu), Order:=xlAscending
                        .SortFields.Add Key:=wsTemp.Range("D2:D" & LastRowMenu), Order:=xlAscending
                    ' Set the range of data to sort
                        .SetRange wsTemp.Range("A1:E" & LastRowMenu)
                        .Header = xlYes
                        .Apply
                    End With
                    
                ' Create a new worksheet and store it in a variable called 'wsMealsLookup'. Name the worksheet "Updated Meals Lookup"
                    Set wsMealsLookup = wbMacro.Worksheets.Add(Before:=wbMacro.Worksheets(1))
                    wsMealsLookup.Name = "Updated Meals Lookup"
                
                ' Copy the data over into the 'wbMacro' Workbook.
                    wsTemp.Range("A:E").Copy Destination:=wsMealsLookup.Range("A1")
                
                ' Get rid of the clipboard
                    Application.CutCopyMode = False
                    
                ' Close 'wbTemp' without saving changes
                    wbTemp.Close SaveChanges:=False
                
                ' Turn off Alerts, so when the worksheet is deleted, it does not bring up an alert.
                    Application.DisplayAlerts = False
                
                ' In the 'wbMacro' workbook, if there is a worksheet called 'Meals Lookup' delete it.
                    For Each ws In wbMacro.Worksheets
                        If ws.Name = "Meals Lookup" Then
                            ws.Visible = xlSheetVisible
                            ws.Delete
                            wsMealsLookup.Name = "Meals Lookup"
                        End If
                    Next ws
                    
                ' Turn back on alerts
                   Application.DisplayAlerts = True
                   
                
                ' Start building the update for the 'MealsLookup' module
                    ' Place the first formula in F1
                        wsMealsLookup.Range("F1").Formula2 = "=""wsLookup.Range(""""A""&ROW(A1)&"""""").Value = """"""&TEXTJOIN(""^|^"",FALSE,A1,B1,C1,D1,E1)"
                        
                    ' Fill it down (using the 'LastRowMenu' variable for the last row)
                        wsMealsLookup.Range("F1:F" & LastRowMenu).FillDown
                    
                    ' Read formulas from wsMealsLookup column F
                        CodeLines = wsMealsLookup.Range("F1:F" & LastRowMenu).Value

                    ' Settings
                        TotalLines = LastRowMenu
                        MaxLinesPerSub = 1000
                        SubCount = Application.WorksheetFunction.Ceiling_Math(TotalLines / MaxLinesPerSub, 1)

                    ' Target code module
                        Set CodeMod = wbMacro.VBProject.VBComponents("CraveIt_Menu")

                    ' Clear MealsLookup module
                        With CodeMod.CodeModule
                            .DeleteLines 1, .CountOfLines
                        End With

                    ' Loop to build the subs
                        LineIndex = 1
                        For i = 1 To SubCount
                        ' If it is the first sub, have it create the "Meals Lookup" worksheet
                            If i = 1 Then
                                With CodeMod.CodeModule
                                    .InsertLines .CountOfLines + 1, "Sub MealsLookup_" & i & "()"
                                    .InsertLines .CountOfLines + 1, "' Last Updated: " & Format(Now, "YYYY.MM.DD \at hh:mm AM/PM")
                                    .InsertLines .CountOfLines + 1, ""
                                    .InsertLines .CountOfLines + 1, "Dim wsLookup As Worksheet"
                                    .InsertLines .CountOfLines + 1, ""
                                    .InsertLines .CountOfLines + 1, Space(4) & "' Create a worksheet called 'Meals Lookup' and store it in a variable called 'wsLookup'"
                                    .InsertLines .CountOfLines + 1, Space(8) & "Set wsLookup = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(""Selection Page""))"
                                    .InsertLines .CountOfLines + 1, Space(8) & "wsLookup.Name = ""Meals Lookup"""
                                    .InsertLines .CountOfLines + 1, ""
                                    .InsertLines .CountOfLines + 1, Space(4) & "' Populate the 'wsLookup' worksheet"
                                End With
                            Else
                            ' Otherwise reference "Meals Lookup" as the worksheet to hold in the variable 'wsLookup'
                                With CodeMod.CodeModule
                                    .InsertLines .CountOfLines + 1, "Sub MealsLookup_" & i & "()"
                                    .InsertLines .CountOfLines + 1, "' Last Updated: " & Format(Now, "YYYY.MM.DD \at hh:mm AM/PM")
                                    .InsertLines .CountOfLines + 1, ""
                                    .InsertLines .CountOfLines + 1, "Dim wsLookup As Worksheet"
                                    .InsertLines .CountOfLines + 1, ""
                                    .InsertLines .CountOfLines + 1, Space(4) & "' Create a variable called 'wsLookup' to hold the value of the worksheet called: 'Meals Lookup'"
                                    .InsertLines .CountOfLines + 1, Space(8) & "Set wsLookup = ThisWorkbook.Sheets(""Meals Lookup"")"
                                    .InsertLines .CountOfLines + 1, ""
                                    .InsertLines .CountOfLines + 1, Space(4) & "' Populate the 'wsLookup' worksheet"
                                End With
                            End If
                            
                            ' Insert the formulas
                                For j = 1 To MaxLinesPerSub
                                    If LineIndex > TotalLines Then Exit For
                                    With CodeMod.CodeModule
                                        .InsertLines .CountOfLines + 1, Space(8) & CodeLines(LineIndex, 1)
                                    End With
                                    LineIndex = LineIndex + 1
                                Next j
                    
                            ' Call next sub or wrap up
                                If i < SubCount Then
                                    With CodeMod.CodeModule
                                        .InsertLines .CountOfLines + 1, ""
                                        .InsertLines .CountOfLines + 1, "' Call the next sub"
                                        .InsertLines .CountOfLines + 1, Space(4) & "MealsLookup_" & (i + 1)
                                        .InsertLines .CountOfLines + 1, ""
                                        .InsertLines .CountOfLines + 1, "End Sub"
                                        .InsertLines .CountOfLines + 1, ""
                                    End With
                                Else
                                ' Final sub with cleanup
                                    With CodeMod.CodeModule
                                        .InsertLines .CountOfLines + 1, ""
                                        .InsertLines .CountOfLines + 1, Space(4) & "' Final Cleanup Steps"
                                        .InsertLines .CountOfLines + 1, Space(8) & "' Create a formula to split out the text."
                                        .InsertLines .CountOfLines + 1, Space(12) & "wsLookup.Range(""B1"").Formula2 = ""=TEXTSPLIT(A1,""""^|^"""",,FALSE)"""
                                        .InsertLines .CountOfLines + 1, ""
                                        .InsertLines .CountOfLines + 1, Space(8) & "' Find the last row of the 'wsLookup' worksheet."
                                        .InsertLines .CountOfLines + 1, Space(12) & "Dim LastRowLookup As Long"
                                        .InsertLines .CountOfLines + 1, Space(12) & "LastRowLookup = wsLookup.Cells(wsLookup.Rows.Count, 1).End(xlUp).Row"
                                        .InsertLines .CountOfLines + 1, ""
                                        .InsertLines .CountOfLines + 1, Space(8) & "' Fill down the formula in column B."
                                        .InsertLines .CountOfLines + 1, Space(12) & "wsLookup.Range(""B1:B"" & LastRowLookup).FillDown"
                                        .InsertLines .CountOfLines + 1, ""
                                        .InsertLines .CountOfLines + 1, Space(8) & "' Copy and paste the values of columns B:F."
                                        .InsertLines .CountOfLines + 1, Space(12) & "wsLookup.Range(""B:F"").Copy"
                                        .InsertLines .CountOfLines + 1, Space(12) & "wsLookup.Range(""B:F"").PasteSpecial xlPasteValues"
                                        .InsertLines .CountOfLines + 1, ""
                                        .InsertLines .CountOfLines + 1, Space(8) & "' Delete column A."
                                        .InsertLines .CountOfLines + 1, Space(12) & "wsLookup.Columns(1).Delete"
                                        .InsertLines .CountOfLines + 1, ""
                                        .InsertLines .CountOfLines + 1, Space(8) & "' Make columns A and D number values rather than text"
                                        .InsertLines .CountOfLines + 1, Space(12) & "wsLookup.Range(""A2:A"" & LastRowLookup).Value = wsLookup.Range(""A2:A"" & LastRowLookup).Value"
                                        .InsertLines .CountOfLines + 1, Space(12) & "wsLookup.Range(""D2:D"" & LastRowLookup).Value = wsLookup.Range(""D2:D"" & LastRowLookup).Value"
                                        .InsertLines .CountOfLines + 1, ""
                                        .InsertLines .CountOfLines + 1, Space(8) & "' Remove the clipboard"
                                        .InsertLines .CountOfLines + 1, Space(12) & "Application.CutCopyMode = False"
                                        .InsertLines .CountOfLines + 1, ""
                                        .InsertLines .CountOfLines + 1, Space(8) & "' AutoFilter and AutoFit columns A:E."
                                        .InsertLines .CountOfLines + 1, Space(12) & "wsLookup.Range(""A1:E1"").AutoFilter"
                                        .InsertLines .CountOfLines + 1, Space(12) & "wsLookup.Columns(""A:E"").AutoFit"
                                        .InsertLines .CountOfLines + 1, ""
                                        .InsertLines .CountOfLines + 1, Space(8) & "' Hide the 'wsLookup' worksheet."
                                        .InsertLines .CountOfLines + 1, Space(12) & "wsLookup.Visible = xlSheetHidden"
                                        .InsertLines .CountOfLines + 1, ""
                                        .InsertLines .CountOfLines + 1, "End Sub"
                                    End With
                                End If
                        Next i

                    ' Get rid of the clipboard
                        Application.CutCopyMode = False
                        
                        
        ' Clean up the Current 'wsMealsLookup' ("Meals Lookup") worksheet.
            ' Delete column F
                wsMealsLookup.Columns(6).Delete
            
            ' Add an AutoFilter and AutoFit
                wsMealsLookup.Range("A1:E1").AutoFilter
                wsMealsLookup.Columns("A:E").AutoFit
                
            ' Hide the 'wsMealsLookup' worksheet.
                wsMealsLookup.Visible = xlSheetHidden
                                        
        ' Give the user a message, letting them know the process is completed.
            MsgBox "Meals Menu updated successfully. Thank you for your patience!", vbInformation, "Update successful"





End Sub
