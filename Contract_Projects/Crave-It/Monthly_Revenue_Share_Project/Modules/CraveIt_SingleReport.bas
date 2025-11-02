Sub Add_Single_CraveIt_Report()
' Last Updated: 2025.10.31

Dim wbMacro As Workbook
Dim fd As FileDialog
Dim FilePath As String
Dim wbTemp As Workbook
Dim wsTemp As Worksheet
Dim wsNew As Worksheet
Dim ws As Worksheet
Dim School As String
Dim DateRange As String
Dim SchoolYearMonth As String
Dim WorksheetExists As Boolean
Dim lastRow As Long
Dim Row As Long

    ' Set this workbook into the variable 'wbMacro'
        Set wbMacro = ThisWorkbook

    ' Ask the user to select a 'Crave It' report file
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .Title = "Select 'Crave It (All days in Range)' Report"
            .Filters.Clear
            .Filters.Add "Excel Files", "*.xls; *.xlsx"
            If .Show <> -1 Then
                MsgBox "No file selected. Please select a report and try again.", vbExclamation
                Exit Sub
            End If
            FilePath = .SelectedItems(1)
        End With

    ' Open the file and hold the workbook in a variable called 'wbTemp'
        Set wbTemp = Workbooks.Open(FilePath)

    ' Set the first sheet of 'wbTemp' into the variable 'wsTemp'
        Set wsTemp = wbTemp.Worksheets(1)

    ' Check if it is the right report, based on the following cells: A1 = "Served Report", A9 = "Items", I9 = "User Type", L9 = "Status", P9 = "Price"
        If wsTemp.Range("A1").Value <> "Served Report" Or wsTemp.Range("A9").Value <> "Items" Or wsTemp.Range("I9").Value <> "User Type" Or _
           wsTemp.Range("L9").Value <> "Status" Or wsTemp.Range("P9").Value <> "Price" Then
            MsgBox "The selected report is not in the correct format. Please select the correct report.", vbExclamation
            wbTemp.Close SaveChanges:=False
            Exit Sub
        End If

    ' Store the School name from A4 in a variable called 'School'
        School = wsTemp.Range("A4").Value

    ' Store the Date Range from U4 in a variable called 'DateRange'
        DateRange = wsTemp.Range("U4").Value

    ' Bring the year and month with the school name into a variable and call it: 'SchoolYearMonth'
        SchoolYearMonth = School & " - " & Right(DateRange, 4) & "." & IIf(Right(Left(DateRange, 2), 1) = "/", "0" & Left(DateRange, 1), Left(DateRange, 2))

    ' Set 'WorksheetExists' to False to start - this will be used later to check if a sheet already exists
        WorksheetExists = False

    ' Check if a worksheet exists for this data using the 'SchoolYearMonth' as the name
        For Each ws In wbMacro.Worksheets
            If ws.Name = SchoolYearMonth Then
                WorksheetExists = True
                Exit For
            End If
        Next ws

    ' Start formatting the data into a worksheet
        ' Unmerge all cells
            wsTemp.Cells.UnMerge

        ' Delete Rows 1-8
            wsTemp.Rows("1:8").Delete

        ' Find the last row of the worksheet using column A and store it in 'LastRow'
            For Each cell In wsTemp.Range("A1:A" & wsTemp.Cells(wsTemp.Rows.Count, "A").End(xlUp).Row)
                If Trim(cell.Value) = "Grand Total:" Then
                    lastRow = cell.Row
                    Exit For
                End If
            Next cell

        ' Delete the last 2 rows of the worksheet
            wsTemp.Rows((lastRow - 1) & ":" & (lastRow + 20)).Delete

        ' Delete columns R:AD, M:O, J:K, C:H
            wsTemp.Columns("R:AD").Delete
            wsTemp.Columns("M:O").Delete
            wsTemp.Columns("J:K").Delete
            wsTemp.Columns("C:H").Delete

        ' Cut Column A and paste it into Column B
            wsTemp.Columns("A").Cut Destination:=wsTemp.Columns("B")

        ' Change the column header in column B to "Item Type"
            wsTemp.Range("B1").Value = "Item Type"

        ' Make the column header in column A "Item Name"
            wsTemp.Range("A1").Value = "Item Name"

        ' Find the new last row, using column B
            lastRow = wsTemp.Cells(wsTemp.Rows.Count, 2).End(xlUp).Row

        ' Loop to pull the Item Name up to the same row as "Item Type"
            For Row = 2 To lastRow Step 2
                wsTemp.Range("A" & Row).Value = Trim(wsTemp.Range("B" & (Row + 1)).Value)
                wsTemp.Range("B" & (Row + 1)).ClearContents
            Next Row

        ' Delete all empty rows
            For Row = lastRow To 3 Step -2
                wsTemp.Rows(Row).Delete
            Next Row

        ' Add a column in front of column A for 'DateRange'
            wsTemp.Columns(1).Insert Shift:=xlToRight
            wsTemp.Range("A1").Value = "Date Range"
            lastRow = wsTemp.Cells(wsTemp.Rows.Count, 2).End(xlUp).Row
            wsTemp.Range("A2:A" & lastRow).Value = DateRange

        ' Add a column in front of column A for 'School'
            wsTemp.Columns(1).Insert Shift:=xlToRight
            wsTemp.Range("A1").Value = "School Name"
            wsTemp.Range("A2:A" & lastRow).Value = School

        ' Loop through column D and update the Item Type abbreviations to full names
            For Row = 2 To lastRow
                Select Case wsTemp.Range("D" & Row).Value
                    Case "D:": wsTemp.Range("D" & Row).Value = "Drink"
                    Case "E:": wsTemp.Range("D" & Row).Value = "Entree"
                    Case "S:": wsTemp.Range("D" & Row).Value = "Side"
                    Case "O:": wsTemp.Range("D" & Row).Value = "Other"
                End Select
            Next Row

        ' Add the column headers for columns I, J, and K
            wsTemp.Range("I1").Value = "Actual Price"
            wsTemp.Range("J1").Value = "Revenue"
            wsTemp.Range("K1").Value = "Revenue Share"

        ' Add the formulas to columns I and J
            wsTemp.Range("I2").Formula = "=IF(G2<>0,G2,IF(D2=""Entree"",IF(ISNUMBER(SEARCH(""w/ milk"",C2)),-3.75,IF(OR(A2=""BASIS Jack Lewis Jr."",A2=""BASIS Med Center"",A2=""BASIS Northeast"",A2=""BASIS Shavano""),-4.5,-5)),IF(ISNUMBER(SEARCH(""Milk"",C2)),-0.85,IF(ISNUMBER(SEARCH(""Water"",C2)),-0.5,""Check""))))"
            wsTemp.Range("J2").Formula = "=I2*H2"

        ' Fill down the formulas if there is more than one row of data
            If lastRow > 2 Then
                wsTemp.Range("I2:J" & lastRow).FillDown
            End If

        ' Remove bold font and borders
            With wsTemp.Cells
                .Font.Bold = False
                .Borders.LineStyle = xlNone
                .WrapText = False
            End With

        ' Make header row bold and left-aligned
            With wsTemp.Range("A1:K1")
                .Font.Bold = True
                .HorizontalAlignment = xlLeft
            End With

        ' Format columns I:K as currency
            wsTemp.Columns("I:K").NumberFormat = "$#,##0.00"

        ' Apply AutoFilter to the header row
            wsTemp.Range("A1:K1").AutoFilter

        ' AutoFit the columns
            wsTemp.Columns("A:K").AutoFit

        ' Sort items by Item Name (C), Status (F), and User Type (E)
            With wsTemp.Sort
                .SortFields.Clear
                .SortFields.Add Key:=wsTemp.Range("C2:C" & lastRow), Order:=xlAscending
                .SortFields.Add Key:=wsTemp.Range("F2:F" & lastRow), Order:=xlDescending
                .SortFields.Add Key:=wsTemp.Range("E2:E" & lastRow), Order:=xlAscending
                .SetRange wsTemp.Range("A1:K" & lastRow)
                .Header = xlYes
                .Apply
            End With

        ' Delete all rows with "Add Funds" in column C
            For Row = lastRow To 2 Step -1
                If Trim(wsTemp.Range("C" & Row).Value) = "Add Funds" Then
                    wsTemp.Rows(Row).Delete
                End If
            Next Row
    
    ' Update the 'LastRow' variable
        lastRow = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row

    ' Add a new worksheet into 'wbMacro' after all other worksheets
        Set wsNew = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets(wbMacro.Worksheets.Count))

    ' If the worksheet does not already exist, rename it using 'SchoolYearMonth' with adjustment (first 18 characters before the last hyphen, plus all after)
        If WorksheetExists = False Then
            wsNew.Name = Left(SchoolYearMonth, Application.Min(20, InStrRev(SchoolYearMonth, "-") - 1)) & Mid(SchoolYearMonth, InStrRev(SchoolYearMonth, "-"))
        End If

    ' Copy the formatted data from 'wsTemp' into the new worksheet
        wsTemp.Range("A1:K" & lastRow).Copy Destination:=wsNew.Range("A1")
    
    ' Check if the "Meals Lookup" worksheet exists.
        For Each ws In wbMacro.Worksheets
            If ws.Name = "Meals Lookup" Then
                WorksheetExists = True
            End If
        Next ws
        
        ' If it does not yet exist, create it.
            If WorksheetExists <> True Then
                MealsLookup_1
            End If
        
    ' Put the formula in column K to find revenue share.
        wsNew.Range("K2").Formula = "=LET(School,A2, ItemType,D2, UserType,E2, ItemName,C2, " & _
                "ItemQty,ISNUMBER(SEARCH(""QTY"",ItemName)), " & _
                "ItemBase,IF(ItemQty,LEFT(ItemName,SEARCH(""QTY"",ItemName)-2),ItemName), SchoolItemLookup,School&"" | ""&ItemBase, " & _
                "MenuPrice,IFERROR(XLOOKUP(SchoolItemLookup,'Meals Lookup'!B:B,'Meals Lookup'!D:D),""""), " & _
                "MenuPriceFlagged,IFERROR(XLOOKUP(SchoolItemLookup,'Meals Lookup'!B:B,'Meals Lookup'!E:E),""Check""), " & _
                "Flagged,IFERROR(MenuPriceFlagged=""Check"",TRUE), " & _
                "PriceMatches,IFERROR(ROUND(MenuPrice,2)=G2,FALSE), " & _
                "SideOrDrink,OR(ItemType=""Drink"",ItemType=""Side""), " & _
                "Breakfast,ISNUMBER(SEARCH(""w/ milk"",ItemName)), " & _
                "ValidatedPrice,IF(Flagged,""Check"",IF(PriceMatches,H2,-1)), " & _
                "" & _
                "result,IF(School=""Central Texas Christian"",IF(SideOrDrink,J2*0.10,IF(UserType<>""Staff"",H2,ValidatedPrice)), " & _
                    "IF(I2<0, " & _
                        "J2, " & _
                        "IF(OR(SideOrDrink,AND(G2<>0,Breakfast)), " & _
                            "J2*0.15, " & _
                            "IF(ItemType=""Entree"", " & _
                                "IF(OR(F2=""Regular"",F2=""Free""), " & _
                                    "H2, " & _
                                    "IF(F2=""Reduced"", " & _
                                        "ValidatedPrice, " & _
                                        """Check"")), " & _
                                """Check"")))), " & _
                "result)"
        
        ' Fill down the formula
            If lastRow > 2 Then
                wsNew.Range("K2:K" & lastRow).FillDown
            End If
        
    ' Remove text wrapping from the new worksheet ('wsNew')
        wsNew.Cells.WrapText = False

    ' Apply AutoFilter and AutoFit to the new worksheet.
        wsNew.Range("A1:K1").AutoFilter
        wsNew.Columns("A:K").AutoFit

    ' Close the temp workbook without saving
        wbTemp.Close SaveChanges:=False
    
    ' Add a message to let the user know the process is completed.
        MsgBox "The file has successfully been added to this workbook.", vbInformation, "Process completed"

End Sub


