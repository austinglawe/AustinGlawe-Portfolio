Sub Add_Multiple_CraveIt_Reports()
' Last Updated: 2025.07.02

Dim wbMacro As Workbook
Dim UserResponse As VbMsgBoxResult
Dim AllowConsolidatedWorksheet As Boolean
Dim fd As FileDialog
Dim FolderPath As String
Dim NewFolderPathRaw As String
Dim NewFolderPathEdit As String
Dim Files As String
Dim FileList() As String
Dim FileCount As Long
Dim ConsolidatedWorksheetExists As Boolean
Dim wsConsolidated As Worksheet
Dim i As Long
Dim FileName As String
Dim WorksheetExistsMeals As Boolean
Dim WorksheetExists As Boolean
Dim SchoolPresent As Boolean
Dim DateRangePresent As Boolean
Dim wbTemp As Workbook
Dim wsTemp As Worksheet
Dim School As String
Dim DateRange As String
Dim SchoolYearMonth As String
Dim NewFileNameRaw As String
Dim LastRow As Long
Dim Row As Long
Dim wsNew As Worksheet
Dim ws As Worksheet
Dim ConsolidatedLastRow As Long
Dim CheckRow As Long
Dim NewFileNameEdit As String
Dim ConsolidatedLastRow2 As Long


    ' Set this workbook into the variable 'wbMacro'
        Set wbMacro = ThisWorkbook
        
    ' Ask user for their preference on if they would like to have a consolidated worksheet or not. Store it in a variable called 'AllowConsolidatedWorksheet'.
        UserResponse = MsgBox("Before getting started, would you like to generate an additional worksheet with all of the data combined?", vbYesNo + vbQuestion, "Consolidated Worksheet")

        If UserResponse = vbYes Then
            AllowConsolidatedWorksheet = True
        Else
            AllowConsolidatedWorksheet = False
        End If

    
    ' Ask user to select a folder with the stored reports
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        With fd
            .Title = "Select 'Crave It (All days in Range)' Reports Folder"
            If .Show <> -1 Then
                MsgBox "No folder selected. Please locate the correct folder and try again.", vbExclamation
                Exit Sub
            End If
            FolderPath = .SelectedItems(1)
        End With

    ' Check to make sure the folder has the name "Crave It" in the name
        If InStr(1, FolderPath, "Crave It", vbTextCompare) = 0 Then
            MsgBox "The folder name must include 'Crave It'. Please locate the correct folder and try again.", vbExclamation
            Exit Sub
        End If

    ' Create a new two variables to hold a new folder name for renaming the: raw reports ('NewFolderPathRaw') and one for renaming the edited reports ('NewFolderPathEdit')
        NewFolderPathRaw = FolderPath & "\Renamed Crave-It Files (Raw)\"
        NewFolderPathEdit = FolderPath & "\Renamed Crave-It Files (Edited)\"

    ' Check to see if there is a folder in the original folder path named "Renamed Crave-It Files (Raw)"
        ' If there is not, then create it.
            If Dir(NewFolderPathRaw, vbDirectory) = "" Then
                MkDir NewFolderPathRaw
            End If
        ' Otherwise move on.

    ' Check to see if there is a folder in the original folder path named "Renamed Crave-It Files (Edited)"
        ' If there is not, then create it.
            If Dir(NewFolderPathEdit, vbDirectory) = "" Then
                MkDir NewFolderPathEdit
            End If
        ' Otherwise move on.
        
    ' Get the all of the file names from the user selected folder and store it in a list.
        Files = Dir(FolderPath & "\*.*")
    
        Do While Files <> ""
            ReDim Preserve FileList(FileCount)
            FileList(FileCount) = Files
            FileCount = FileCount + 1
            Files = Dir
        Loop
    
    ' If AllowConsolidatedWorksheet is set to 'True', set it up.
        If AllowConsolidatedWorksheet = True Then
            ' Start by setting the variable to False
                ConsolidatedWorksheetExists = False
                
            ' Check to see if a 'Consolidated Reports' worksheet exists by looping through each worksheet.
                    For Each ws In wbMacro.Worksheets
                        If ws.Name = "Consolidated Reports" Then
                        ' If it exists, set 'ConsolidatedWorksheetExists' to 'True'
                            ConsolidatedWorksheetExists = True
                            Exit For
                        End If
                    Next ws
                    ' If it does not already exist, create it:
                        If ConsolidatedWorksheetExists = False Then
                            Set wsConsolidated = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets(1))
                            wsConsolidated.Name = "Consolidated Reports"
                            ' Add the column Headers
                                wsConsolidated.Range("A1:K1").Value = Array("School Name", "Date Range", "Item Name", "Item Type", "User Type", "Status", "Price", "# Orders", "Actual Price", "Revenue", "Revenue Share")
                            ' Make Bold
                                wsConsolidated.Range("A1:K1").Font.Bold = True
                        Else
                        ' If it exists store the worksheet in the variable called 'wsConsolidated'
                            Set wsConsolidated = wbMacro.Worksheets("Consolidated Reports")
                        End If
        End If
        
    ' Check if the "Meals Lookup" worksheet exists.
        WorksheetExistsMeals = False
        
        For Each ws In wbMacro.Worksheets
            If ws.Name = "Meals Lookup" Then
                WorksheetExistsMeals = True
            End If
        Next ws
        
        ' If it does not yet exist, create it.
            If WorksheetExistsMeals <> True Then
                MealsLookup_1
            End If
    
    ' Loop through each file in the folder the user gave.
        For i = 0 To UBound(FileList)
        
            ' Set the 'FileName' variable equal to the name in the loop list
                FileName = FileList(i)
            
            ' Set the variable 'WorksheetExists', 'SchoolPresent', and 'DateRangePresent' to False - these will be used later to make sure no data is duplicated.
                WorksheetExists = False
                SchoolPresent = False
                DateRangePresent = False
                
            ' Set the 'FilePath' based on the 'FolderPath' and 'FileName' variable. - this is the file to use.
                FilePath = FolderPath & "\" & FileName
        
            ' Check file to make sure it is an Excel file (.xls, .xlsx)
                ' If it does not have the extension ".xls" or ".xlsx" then go to the next file.
                    If Not (LCase(Right(FilePath, 4)) = ".xls" Or LCase(Right(FilePath, 5)) = ".xlsx") Then
                        GoTo NextFile
                    End If
                ' Otherwise proceed:
                
            ' Open the file and hold the workbook in a variable called 'wbTemp'
                Set wbTemp = Workbooks.Open(FilePath)
    
            ' Set the first sheet of 'wbTemp' into the variable 'wsTemp'
                Set wsTemp = wbTemp.Worksheets(1)
                
            ' Check if it is the right report, based on the following cells: A1 = "Served Report", A9 = "Items", I9 = "User Type", L9 = "Status", P9 = "Price"
                ' If it is does not contain these values in the cells, close the file and go to the next file.
                    If wsTemp.Range("A1").Value <> "Served Report" Or wsTemp.Range("A9").Value <> "Items" Or wsTemp.Range("I9").Value <> "User Type" Or _
                      wsTemp.Range("L9").Value <> "Status" Or wsTemp.Range("P9").Value <> "Price" Then
                        wbTemp.Close SaveChanges:=False
                        GoTo NextFile
                    End If
                ' If it is the correct report. Proceed:
                
            ' Store the School name from A4 in a variable called: "School"
                School = wsTemp.Range("A4").Value
    
            ' Store the Date Range from U4 in a variable called: "DateRange"
                DateRange = wsTemp.Range("U4").Value
    
            ' Bring the year and month with the school name into a variable and call it: "SchoolYearMonth"
                SchoolYearMonth = School & " - " & Right(DateRange, 4) & "." & IIf(Right(Left(DateRange, 2), 1) = "/", "0" & Left(DateRange, 1), Left(DateRange, 2))
    
            ' To hold the raw file name/full path, create a new variable called 'NewFileNameRaw'
                NewFileNameRaw = NewFolderPathRaw & SchoolYearMonth & " - Raw.xlsx"
                
            ' Check if the raw file name exists
                ' If it does not exist, save the file with the 'NewFileNameRaw' variable.
                    If Dir(NewFileNameRaw) = "" Then
                        wbTemp.SaveAs FileName:=NewFileNameRaw, FileFormat:=xlOpenXMLWorkbook
                    End If
                ' If it already exists, proceed:
            
            ' Start formatting the data into a worksheet.
                ' Unmerge all cells
                    wsTemp.Cells.UnMerge
                
                ' Delete Rows 1-8.
                    wsTemp.Rows("1:8").Delete
            
                ' Find the last row of the worksheet using column A and store it in a variable called 'LastRow'
                    ' LastRow = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
                    For Each cell In wsTemp.Range("A1:A" & wsTemp.Cells(wsTemp.Rows.Count, "A").End(xlUp).Row)
                        If Trim(cell.Value) = "Grand Total:" Then
                            LastRow = cell.Row
                            Exit For
                        End If
                    Next cell
                
                ' Delete the last 2 rows of the worksheet.
                    wsTemp.Rows((LastRow - 1) & ":" & (LastRow + 20)).Delete
                
                ' Delete Columns R:AD, M:O, J:K, C:H
                    wsTemp.Columns("R:AD").Delete
                    wsTemp.Columns("M:O").Delete
                    wsTemp.Columns("J:K").Delete
                    wsTemp.Columns("C:H").Delete
        
                ' Cut Column A and paste it into Column B
                    wsTemp.Columns("A").Cut Destination:=wsTemp.Columns("B")
            
                ' Change the Column Header in column B to "Item Type"
                    wsTemp.Range("B1").Value = "Item Type"
                
                ' Make the column header in column A "Item Name"
                    wsTemp.Range("A1").Value = "Item Name"
                
                ' Find the new last row, using column B.
                    LastRow = wsTemp.Cells(wsTemp.Rows.Count, 2).End(xlUp).Row
                
                ' Create a loop to pull the Item Name up to the same row as "Item Type"
                    For Row = 2 To LastRow Step 2
                        wsTemp.Range("A" & Row).Value = Trim(wsTemp.Range("B" & (Row + 1)).Value)
                        wsTemp.Range("B" & (Row + 1)).ClearContents
                    Next Row
    
                ' Delete all empty rows
                    For Row = LastRow To 3 Step -2
                        wsTemp.Rows(Row).Delete
                    Next Row
                
                ' Find the new last row based on column A
                    LastRow = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
            
                ' Add a column in front of column A - for 'DateRange'
                    wsTemp.Columns(1).Insert Shift:=xlToRight
                    
                    ' Put the column header "Date Range"
                        wsTemp.Range("A1").Value = "Date Range"
                    
                    ' Put the 'DateRange' value in row of the worksheet
                        wsTemp.Range("A2:A" & LastRow).Value = DateRange
                    
                ' Add another column in front of column A - for 'School'
                    wsTemp.Columns(1).Insert Shift:=xlToRight
                    
                    ' Put the column header "School Name"
                        wsTemp.Range("A1").Value = "School Name"
                    
                    ' Put the 'School' value in row of the worksheet
                        wsTemp.Range("A2:A" & LastRow).Value = School
                        
                ' Loop through column D line by line and change the abbreviated Item type into the full type.
                    For Row = 2 To LastRow
                        If wsTemp.Range("D" & Row).Value = "D:" Then
                            wsTemp.Range("D" & Row).Value = "Drink"
                        ElseIf wsTemp.Range("D" & Row).Value = "E:" Then
                            wsTemp.Range("D" & Row).Value = "Entree"
                        ElseIf wsTemp.Range("D" & Row).Value = "S:" Then
                            wsTemp.Range("D" & Row).Value = "Side"
                        ElseIf wsTemp.Range("D" & Row).Value = "O:" Then
                            wsTemp.Range("D" & Row).Value = "Other"
                        End If
                    Next Row
                
                ' Add the column headers to column I, J, and K
                    wsTemp.Range("I1").Value = "Actual Price"
                    wsTemp.Range("J1").Value = "Revenue"
                    wsTemp.Range("K1").Value = "Revenue Share"
                
                ' Add the formulas to column I and J and fill them down
                    ' Add Formulas
                        wsTemp.Range("I2").Formula = "=IF(G2<>0,G2,IF(D2=""Entree"",IF(ISNUMBER(SEARCH(""w/ milk"",C2)),-3.75,IF(AND(OR(A2=""BASIS Jack Lewis Jr."",A2=""BASIS Med Center"",A2=""BASIS Northeast"",A2=""BASIS Shavano""),AND(NOT(ISNUMBER(SEARCH(""burger"",C2))),NOT(ISNUMBER(SEARCH(""V:"",C2))))),-4.5,-5)),IF(ISNUMBER(SEARCH(""Milk"",C2)),-0.85,IF(ISNUMBER(SEARCH(""Water"",C2)),-0.5,""Check""))))"
                        wsTemp.Range("J2").Formula = "=I2*H2"
                        wsTemp.Range("K2").Formula = "'=IF(I2<0,J2,IF(OR(D2=""Drink"",D2=""Side"",AND(G2<>0,ISNUMBER(SEARCH(""w/ milk"",C2)))),J2*0.1,IF(D2=""Entree"",IF(F2=""Regular"",H2,IF(F2=""Free"",H2,IF(F2=""Reduced"",IF(ISNUMBER(SEARCH(""QTY"",C2)),IF(XLOOKUP(A2&"" | ""&LEFT(C2,SEARCH(""QTY"",C2)-2),'Meals Lookup'!B:B,'Meals Lookup'!E:E)=""Check"",""Check"",IF(ROUND(XLOOKUP(A2&"" | ""&LEFT(C2,SEARCH(""QTY"",C2)-2),'Meals Lookup'!B:B,'Meals Lookup'!D:D),2)<>G2,0,H2)),IF(XLOOKUP(A2&"" | ""&C2,'Meals Lookup'!B:B,'Meals Lookup'!E:E)=""Check"",""Check"",IF(ROUND(XLOOKUP(A2&"" | ""&C2,'Meals Lookup'!B:B,'Meals Lookup'!D:D),2)<>G2,0,H2))),""Check""))),""Check"")))"
                        
                    ' Fill Down
                        If LastRow > 2 Then
                            wsTemp.Range("I2:J" & LastRow).FillDown
                        End If
                        
    
                ' Format the cells
                    ' Remove Bold font and any border lines
                        With wsTemp.Cells
                            .Font.Bold = False
                            .Borders.LineStyle = xlNone
                        End With
            
                    ' Make header row bold and make the text start from the left side of the cell
                        With wsTemp.Range("A1:K1")
                            .Font.Bold = True
                            .HorizontalAlignment = xlLeft
                        End With
                        
                    ' Format columns I and J as Currency
                        wsTemp.Columns("I:K").NumberFormat = "$#,##0.00"
                        
                    ' Remove text wrapping from the new worksheet ('wsNew')
                        wsTemp.Cells.WrapText = False
                        
                    ' Apply AutoFilter to range A1:H1
                        wsTemp.Range("A1:K1").AutoFilter
                    
                    ' AutoFit columns A:H
                        wsTemp.Columns("A:K").AutoFit
                    
                    ' Sort Items in the table. Using Item Name(Column C), then Status (Column F), lastly by User Type (Column E)
                        ' Note: Items are sorted in a heirarchy - (each column in the hierarchy is only used if the previous one's results are equal)
                        With wsTemp.Sort
                            .SortFields.Clear
                        ' Sort by Item Name (Column C)
                            .SortFields.Add Key:=wsTemp.Range("C2:C" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                        ' Sort by Status (Column F)
                            .SortFields.Add Key:=wsTemp.Range("F2:F" & LastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                        ' Sort by User Type (Column E)
                            .SortFields.Add Key:=wsTemp.Range("E2:E" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                        ' Set Range and other filter parameters
                            .SetRange wsTemp.Range("A1:K" & LastRow)
                            .Header = xlYes
                            .MatchCase = False
                            .Orientation = xlTopToBottom
                            .Apply
                        End With
                    
                    ' Delete all rows with "Add Funds" in column C
                        For Row = LastRow To 2 Step -1
                            If Trim(wsTemp.Range("C" & Row).Value) = "Add Funds" Then
                                wsTemp.Rows(Row).Delete
                            End If
                        Next Row

                    ' update the 'LastRow' variable
                        LastRow = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row

                    
            ' Check if a worksheet exists for this data using the 'SchoolYearMonth' as the name the worksheet is named.
                ' If it does, then store "True" in the 'WorksheetExists' variable. - this will make it so the data is not pulled into the 'wbMacro' workboook.
                    For Each ws In wbMacro.Worksheets
                        If ws.Name = SchoolYearMonth Then
                            WorksheetExists = True
                            Exit For
                        End If
                    Next ws
                ' If it does not exist yet, create it:
                        
            ' Check if 'WorksheetExists' is False or True - if it does not yet exist, copy the data into the 'wbMacro' workbook.
                If WorksheetExists = False Then
                    ' Create a new worksheet in 'wbMacro' and store it in a variable called 'wsNew'
                        Set wsNew = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets(wbMacro.Worksheets.Count))
                        
                        ' Rename the worksheet based on the 'SchoolYearMonth' variable with an adjustment - first 20 characters before the "-" and all characters after
                            wsNew.Name = Left(SchoolYearMonth, Application.Min(20, InStrRev(SchoolYearMonth, "-") - 1)) & Mid(SchoolYearMonth, InStrRev(SchoolYearMonth, "-"))

                    ' Copy the 'wsTemp' worksheet into the 'wbMacro' workbook
                        wsTemp.Range("A1:K" & LastRow).Copy Destination:=wsNew.Range("A1")
                        
                    ' Put the formula in column K to find revenue share.
                        wsNew.Range("K2").Formula = "=IF(I2<0,J2,IF(OR(D2=""Drink"",D2=""Side"",AND(G2<>0,ISNUMBER(SEARCH(""w/ milk"",C2)))),J2*0.1,IF(D2=""Entree"",IF(F2=""Regular"",H2,IF(F2=""Free"",H2,IF(F2=""Reduced"",IF(ISNUMBER(SEARCH(""QTY"",C2)),IF(XLOOKUP(A2&"" | ""&LEFT(C2,SEARCH(""QTY"",C2)-2),'Meals Lookup'!B:B,'Meals Lookup'!E:E)=""Check"",""Check"",IF(ROUND(XLOOKUP(A2&"" | ""&LEFT(C2,SEARCH(""QTY"",C2)-2),'Meals Lookup'!B:B,'Meals Lookup'!D:D),2)<>G2,0,H2)),IF(XLOOKUP(A2&"" | ""&C2,'Meals Lookup'!B:B,'Meals Lookup'!E:E)=""Check"",""Check"",IF(ROUND(XLOOKUP(A2&"" | ""&C2,'Meals Lookup'!B:B,'Meals Lookup'!D:D),2)<>G2,0,H2))),""Check""))),""Check"")))"
                    ' Fill down the formula
                        If LastRow > 2 Then
                            wsNew.Range("K2:K" & LastRow).FillDown
                        End If
                        
                    ' Remove text wrapping from the new worksheet ('wsNew')
                        wsNew.Cells.WrapText = False
                    
                    ' Apply AutoFilter to range A1:K1
                        wsNew.Range("A1:K1").AutoFilter
                
                    ' AutoFit columns A:K
                        wsNew.Columns("A:K").AutoFit
                End If
            
            ' Check if 'AllowConsolidatedWorksheet' is True - if it is check to see if the worksheet data is already in the worksheet
                If AllowConsolidatedWorksheet = True Then
                    ' Find the last row of the 'wsConsolidated' worksheet and store it in a variable called 'ConsolidatedLastRow'
                        ConsolidatedLastRow = wsConsolidated.Cells(wsConsolidated.Rows.Count, 1).End(xlUp).Row + 1
                    ' Loop through rows to check if School and DateRange exist together
                        For CheckRow = 2 To ConsolidatedLastRow
                            If wsConsolidated.Cells(CheckRow, 1).Value = School And wsConsolidated.Cells(CheckRow, 2).Value = DateRange Then
                                SchoolPresent = True
                                DateRangePresent = True
                                Exit For
                            End If
                        Next CheckRow
                    ' If 'SchoolPresent' AND 'DateRangePresent' are both true, then do not add them to the 'wsConsolidated' otherwise, add them starting in "A" & ConsolidatedLastRow
                        If SchoolPresent = False And DateRangePresent = False Then
                            wsTemp.Range("A2:K" & LastRow).Copy Destination:=wsConsolidated.Range("A" & ConsolidatedLastRow)
                        End If
                End If
                
            ' Create a new file for the newly formatted data, using the variable 'NewFileNameEdit'
                NewFileNameEdit = NewFolderPathEdit & SchoolYearMonth & " - Edited.xlsx"
                
                ' Check if the file name exists in the Edited files folder.
                    ' If it does not yet exist, then save the file to the folder with the new in the 'NewFileNameEdit' variable
                        If Dir(NewFileNameEdit) = "" Then
                            wbTemp.SaveAs FileName:=NewFileNameEdit, FileFormat:=xlOpenXMLWorkbook
                        End If
                    ' If it already exists:
                   
            ' Close the workbook without saving changes.
                wbTemp.Close SaveChanges:=False
                
                
NextFile:
            ' Move to next file
                Next i
                
    ' Check if 'AllowConsolidatedWorksheet' is True - if it is, then autofit the columns and autofilter row 1
        If AllowConsolidatedWorksheet = True Then
            ' Find the new last row
                ConsolidatedLastRow2 = wsConsolidated.Cells(wsConsolidated.Rows.Count, 1).End(xlUp).Row
                
            ' Put in the formula for column K
                wsConsolidated.Range("K2").Formula = "=IF(I2<0,J2,IF(OR(D2=""Drink"",D2=""Side"",AND(G2<>0,ISNUMBER(SEARCH(""w/ milk"",C2)))),J2*0.1,IF(D2=""Entree"",IF(F2=""Regular"",H2,IF(F2=""Free"",H2,IF(F2=""Reduced"",IF(ISNUMBER(SEARCH(""QTY"",C2)),IF(XLOOKUP(A2&"" | ""&LEFT(C2,SEARCH(""QTY"",C2)-2),'Meals Lookup'!B:B,'Meals Lookup'!E:E)=""Check"",""Check"",IF(ROUND(XLOOKUP(A2&"" | ""&LEFT(C2,SEARCH(""QTY"",C2)-2),'Meals Lookup'!B:B,'Meals Lookup'!D:D),2)<>G2,0,H2)),IF(XLOOKUP(A2&"" | ""&C2,'Meals Lookup'!B:B,'Meals Lookup'!E:E)=""Check"",""Check"",IF(ROUND(XLOOKUP(A2&"" | ""&C2,'Meals Lookup'!B:B,'Meals Lookup'!D:D),2)<>G2,0,H2))),""Check""))),""Check"")))"
                ' Fill it down
                    If ConsolidatedLastRow2 > 2 Then
                        wsConsolidated.Range("K2:K" & ConsolidatedLastRow2).FillDown
                    End If
                    
            ' AutoFilter and AutoFit column A:K
                wsConsolidated.Range("A1:K1").AutoFilter
                wsConsolidated.Columns("A:K").AutoFit
        End If
        
    ' Add a message to let the user know the process is completed.
        MsgBox "The files have successfully been added to this workbook. Thank you for your patience!", vbInformation, "Process completed"
            
End Sub
