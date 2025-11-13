' Blackbaud 'AR' Reconciliation Process:
    ' Step 1: Download the GL for the approriate account from Intacct. (12010-[FY]) -- For example: 12010-2425
        ' Run the report on Intacct with a start date of Jan 1 prior to the fiscal year. -- For example: 1/1/2024 for the 12010-2425 account
    ' Step 2: Download all of the A/R Balance Reports from Blackbaud.
        ' They are found by going into the appropriate school. Then to 'Reports'. Then Download 'A/R Balance' Reports (under 'Accrual' heading)
            ' These need to be downloaded on a monthly basis. Use the parameters "Family" for 'Report Basis' and "Annual" for the 'Date Range Type'
                ' To pull the data from inception through the month you want, use the 1st of the next month-- For example to pull from inception through March 2025. Select April 1, 2025
            ' Once downloaded, put all A/R reports into a single folder.
    ' Step 3: Create a new file called: [School] "Blackbaud AR Reconciliation" [FY(s)]
        ' Now go into the GL Report
            ' Unmerge all cells
            ' Unwrap all cells
            ' Copy and Paste in the columns (and in this order): ("Posted Date", "Doc dt.", "Doc", "Memo/Description", "Location", "Division", "Funding Source", "Debt Services Series", "Txn No", "JNL", "Debit", "Credit", "Balance")
            ' Close out of the GL file.
        ' Start editing the GL data in the new file.
            ' Rename the worksheet the GL Data was pulled into "12010-[FY] GL"
            ' Highlight all of the 'Website Deposits' in light blue.
            ' Highlight all of the 'Inschool Deposits' in light green.
            ' Put a borderline beneath the end of each month (to help separate them and improve readability)
            ' Start adding up the numbers for "Billings", "Website Deposits", "Inschool Deposits", and "Totals" by month
        ' Now Create a New Worksheet called: [School] "Variance Analysis" [FY]
            ' Use the analysis page from previous 'Blackbaud AR Recon' Files.
                ' Populate all of the appropriate data and clear all of the non-formulated cells
            ' Start opening the 'A/R Balance' Reports that were downloaded.
                ' Go to the last row of the 'A/R Balance' Report.
                    ' Copy columns D:I and paste the values into the 'Variance Analysis' worksheet for the approriate month
                        ' The values will be placed in the row above the grey row.
                ' Do this for all of the files.
            ' Add the appropriate totals from the 'GL' worksheet into columns O:Q for the appropriate months.
        ' Now Check for variances between Blackbaud and Intacct, using columns S:U
    ' Step 4: Figure out where the variances are coming from.
    ' Step 5: Finalize Report.

Sub BlackbaudARRecon_Part1()

Dim UserResponse As VbMsgBoxResult

Dim ws As Worksheet
Dim wsReset As Worksheet
Dim ResetExists As Boolean


Dim wbMacro As Workbook

Dim fd As FileDialog
Dim FilePath As String
Dim FolderPath As String
Dim wbGL As Workbook
Dim wsGL As Worksheet

Dim PosteddtColumn As Long
Dim DocdtColumn As Long
Dim DocColumn As Long
Dim MemoColumn As Long
Dim LocationColumn As Long
Dim DivisionColumn As Long
Dim FundingSourceColumn As Long
Dim DebtServicesColumn As Long
Dim TxnNoColumn As Long
Dim JNLColumn As Long
Dim DebitColumn As Long
Dim CreditColumn As Long
Dim BalanceColumn As Long

Dim Col As Long
Dim HeaderName As String
Dim NotFound As String

Dim AccountPlusYear As String
Dim FY As String

Dim School As String
Dim SchoolName As String

Dim SheetExists As Boolean

Dim GLEndRow As Long
Dim GLWorksheetName As String
Dim wsGLAnalysis As Worksheet

Dim GLAnalysisLastRow As Long
Dim CurrentRow As Long

Dim wsVarianceAnalysis As Worksheet
 
Dim LoopNumber As Integer
Dim GLColumn As String
Dim GLBillings As Integer
Dim GLWebsiteDeposits As Integer
Dim GLInSchoolDeposits As Integer
                                
Dim FYLoop As Integer
Dim LastRow As Long
Dim LoopMonth As Integer



' Ask user if they are sure they want to start the converter
    UserResponse = MsgBox("Are you sure you want to start the 'Blackbaud AR Reconciliation' Converter?", vbYesNo + vbQuestion, "Confirmation to start the 'Blackbaud AR Recon' Converter")


' Check the user's response - If they choose no, end the sub, otherwise continue to the next steps.
    If UserResponse = vbNo Then
        Exit Sub
    End If
'______________________________________________________________________________________________________________________________________________


' Check if "COMPLETE RESET" worksheet already exists
ResetExists = False
For Each ws In ThisWorkbook.Worksheets
    If ws.Name = "COMPLETE RESET" Then
        ResetExists = True
        Exit For
    End If
Next ws

    ' If it doesn't exist, and we’re on the Selection Page, then create it and delete Selection Page
    If Not ResetExists And ActiveSheet.Name = "Converter Selection Page" Then
    ' Create the Reset Worksheet
        Reset.Create_Reset_Worksheet
    End If

'______________________________________________________________________________________________________________________________________________

' Set the macro workbook to the variable 'wbMacro'
    Set wbMacro = ThisWorkbook




    ' Ask user to select the GL File
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
            With fd
                .Title = "Select 12010--GL File"
                If .Show <> -1 Then
                    MsgBox "No File selected. Please locate the file and try again."
                    Exit Sub
                End If
                FilePath = .SelectedItems(1)
            End With
    ' Hold the 'FolderPath' to be able to rename the 'FilePath' file.
        FolderPath = Left(FilePath, InStrRev(FilePath, "\"))
    
    ' Open the selected file.
        Set wbGL = Workbooks.Open(FilePath)
        ' Create a variable to hold the GL worksheet as 'wsTempGL'
            Set wsGL = wbGL.Worksheets(1)
        ' Unmerge Cells
            wsGL.Cells.UnMerge
        ' Unwrap Cells
            wsGL.Cells.WrapText = False
        ' Check to make sure all of the appropriate information is included, if not at any point in this process, notify the user what issue has occurred and help them to get the correct report.
            ' In Row 7 Check that the following column headers exist within the file:
                ' "Posted dt.", "Doc dt.", "Doc", "Memo/Description", "Location", "Division", "Funding Source", "Debt Service Series", "Txn No", "JNL", "Debit", "Credit", "Balance"
                ' If they do, hold them in a variable to remember their column numbers
                For Col = 1 To wsGL.Cells(7, Columns.Count).End(xlToLeft).Column
                HeaderName = Trim(wsGL.Cells(7, Col).Value)
                
                Select Case LCase(HeaderName)
                    Case "posted dt."
                        PosteddtColumn = Col
                    Case "doc dt."
                        DocdtColumn = Col
                    Case "doc"
                        DocColumn = Col
                    Case "memo/description"
                        MemoColumn = Col
                    Case "location"
                        LocationColumn = Col
                    Case "division"
                        DivisionColumn = Col
                    Case "funding source"
                        FundingSourceColumn = Col
                    Case "debt service series"
                        DebtServicesColumn = Col
                    Case "txn no"
                        TxnNoColumn = Col
                    Case "jnl"
                        JNLColumn = Col
                    Case "debit"
                        DebitColumn = Col
                    Case "credit"
                        CreditColumn = Col
                    Case "balance"
                        BalanceColumn = Col
                End Select
            Next Col
            
            ' Now check if any are missing
                NotFound = ""
                
                If PosteddtColumn = 0 Then NotFound = NotFound & "Posted dt., "
                If DocdtColumn = 0 Then NotFound = NotFound & "Doc dt., "
                If DocColumn = 0 Then NotFound = NotFound & "Doc, "
                If MemoColumn = 0 Then NotFound = NotFound & "Memo/Description, "
                If LocationColumn = 0 Then NotFound = NotFound & "Location, "
                If DivisionColumn = 0 Then NotFound = NotFound & "Division, "
                If FundingSourceColumn = 0 Then NotFound = NotFound & "Funding Source, "
                If DebtServicesColumn = 0 Then NotFound = NotFound & "Debt Services Series, "
                If TxnNoColumn = 0 Then NotFound = NotFound & "Txn No, "
                If JNLColumn = 0 Then NotFound = NotFound & "JNL, "
                If DebitColumn = 0 Then NotFound = NotFound & "Debit, "
                If CreditColumn = 0 Then NotFound = NotFound & "Credit, "
                If BalanceColumn = 0 Then NotFound = NotFound & "Balance, "
                
                If Len(NotFound) > 0 Then
                    MsgBox "The following required columns were NOT found in Row 7: " & vbNewLine & NotFound, vbCritical, "Missing Columns"
                    wbGL.Close SaveChanges:=False
                    Exit Sub
                End If
                    
                ' If they all exist:
                    ' Check "A8" for its value (this is where the account number is)
                        ' Get the first 10 characters of A8. Store it in a variable 'AccountPlusYear'
                            AccountPlusYear = Left(wsGL.Range("A8").Value, 10)
                        
                        ' Make sure the first 5 characters are "12010" (Blackbaud Receivables account)
                        If Left(AccountPlusYear, 5) = "12010" Then
                        ' Take the last 2 characters of 'AccountPlusYear' and store it in a variable called 'FY' (Fiscal Year)
                            FY = Right(AccountPlusYear, 2)
                        Else
                            MsgBox "Account does not start with 12010. Please check the GL file.", vbCritical, "Invalid Account Number"
                            wbGL.Close SaveChanges:=False
                            Exit Sub
                        End If
        ' Use cell "B6" to determine the school. Hold it in a variable called 'School'
            School = wsGL.Range("B6").Value
            ' Pass it through a loop to determine appropriate naming convention for 'Blackbaud AR Reconciliation'. Store it in a variable called 'SchoolName'
                If School = "101--San Antonio Primary Medical Center" Then
                    SchoolName = "101 SAMC"
                ElseIf School = "102--San Antonio Primary North Central" Then
                    SchoolName = "102 SANC"
                ElseIf School = "103--San Antonio Shavano Campus" Then
                    SchoolName = "103 SASH"
                ElseIf School = "104--San Antonio Primary Northeast" Then
                    SchoolName = "104 SPNE"
                ElseIf School = "105--Austin Primary" Then
                    SchoolName = "105 AUSP"
                ElseIf School = "106--San Antonio Northeast" Then
                    SchoolName = "106 SANE"
                ElseIf School = "107--Austin" Then
                    SchoolName = "107 AUS"
                ElseIf School = "108--Pflugerville Primary" Then
                    SchoolName = "108 PFLP"
                ElseIf School = "109--Jack Lewis, Jr. Primary" Then
                    SchoolName = "109 JLJP"
                ElseIf School = "110--Benbrook" Then
                    SchoolName = "110 BEN"
                ElseIf School = "111--Pflugerville" Then
                    SchoolName = "111 PFL"
                ElseIf School = "112--Jack Lewis, Jr." Then
                    SchoolName = "112 JLJ"
                ElseIf School = "113--Cedar Park Primary" Then
                    SchoolName = "113 CPKP"
                ElseIf School = "114--Cedar Park" Then
                    SchoolName = "114 CPK"
                ElseIf School = "201--Washington DC (School)" Then
                    SchoolName = "201 BDC"
                ElseIf School = "401--Ahwatukee" Then
                    SchoolName = "401 AH"
                ElseIf School = "402--Chandler" Then
                    SchoolName = "402 CH"
                ElseIf School = "403--Chandler Primary South" Then
                    SchoolName = "403 CPS"
                ElseIf School = "404--Chander Primary North" Then
                    SchoolName = "404 CPN"
                ElseIf School = "405--Flagstaff" Then
                    SchoolName = "405 FL"
                ElseIf School = "406--Goodyear" Then
                    SchoolName = "406 GO"
                ElseIf School = "407--Mesa" Then
                    SchoolName = "407 ME"
                ElseIf School = "408--Oro Valley" Then
                    SchoolName = "408 OV"
                ElseIf School = "409--Peoria" Then
                    SchoolName = "409 PEO"
                ElseIf School = "410--Phoenix" Then
                    SchoolName = "410 PHX"
                ElseIf School = "411--Phoenix Central" Then
                    SchoolName = "411 PC"
                ElseIf School = "412--Prescott" Then
                    SchoolName = "412 PR"
                ElseIf School = "413--Scottsdale" Then
                    SchoolName = "413 SC"
                ElseIf School = "414--Scottsdale Primary East" Then
                    SchoolName = "414 SPE"
                ElseIf School = "415--Tucson North" Then
                    SchoolName = "415 TN"
                ElseIf School = "416--Tucson Primary" Then
                    SchoolName = "416 TP"
                ElseIf School = "417--Goodyear Primary" Then
                    SchoolName = "417 GYP"
                ElseIf School = "418--Oro Valley Primary" Then
                    SchoolName = "418 OP"
                ElseIf School = "419--Peoria Primary" Then
                    SchoolName = "419 PP"
                ElseIf School = "420--Phoenix South" Then
                    SchoolName = "420 PS"
                ElseIf School = "421--Phoenix Primary" Then
                    SchoolName = "421 PXP"
                ElseIf School = "422--Scottsdale Primary West" Then
                    SchoolName = "422 SCW"
                ElseIf School = "423--Phoenix North" Then
                    SchoolName = "423 PHXN"
                ElseIf School = "701--Baton Rouge - Materra" Then
                    SchoolName = "701 MA"
                ElseIf School = "702--Baton Rouge - Mid City" Then
                    SchoolName = "702 MC"
                Else
                    SchoolName = Left(School, 3)
                End If
        ' Now find the last row of the worksheet using column "A" as the row lookup. Store it in a variable called 'GLEndRow'
            GLEndRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
            
    ' Create a variable called 'GLWorksheetName' for use later in the macro and for naming the 'wsGLAnalysis'
        GLWorksheetName = SchoolName & " " & AccountPlusYear & " GL"
        
    ' Create a worksheet with the variable name 'wsGLAnalysis'. Use the following naming convention for the worksheet name: [SchoolName] 'AccountPlusYear' & "GL" in the macro workbook.
        ' Check if a worksheet with the name from the variable 'GLWorksheetName', exists. If it does, close the temporary workbook and exit the sub.
            For Each ws In ThisWorkbook.Worksheets
                If ws.Name = GLWorksheetName Then
                    wbGL.Close SaveChanges:=False
                    MsgBox "This GL has already been pulled into this file. Please find a different one, and try again."
                    Exit Sub
                End If
            Next ws
            
        ' Check if ""Add or Split Out Reports"" worksheet exists.
            SheetExists = False
            For Each ws In ThisWorkbook.Worksheets
                If ws.Name = "Add or Split Out Reports" Then
                    SheetExists = True
                    Exit For
                End If
            Next ws
            
        ' If it doesn't exist, place the new worksheet as the last sheet, otherwise, place the new worksheet before the "Add or Split Out Reports" Worksheet.
            If Not SheetExists Then
                Set wsGLAnalysis = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets(wbMacro.Worksheets.Count))
            Else
                Set wsGLAnalysis = wbMacro.Worksheets.Add(Before:=wbMacro.Worksheets("Add or Split Out Reports"))
            End If
            
        ' Rename the worksheet using the 'GLWorksheetName' Variable.
        wsGLAnalysis.Name = GLWorksheetName
        
        ' Go back to the GL worksheet/workbook, copy "A1:B6" into the new GL worksheet in cell "A20"
            wsGL.Range("A1:B6").Copy wsGLAnalysis.Range("A20")
            Application.CutCopyMode = False
            
        ' Copy the columns: "Posted dt.", "Doc dt.", "Doc", "Memo/Description", "Location", "Division", "Funding Source", "Debt Services Series", "Txn No", "JNL", "Debit", "Credit", "Balance"
            ' In that order
            ' From row 7 to the end row ('GLEndRow')
            ' Paste into the new GL worksheet starting in row 27.
                ' Posted dt.
                    wsGL.Range(wsGL.Cells(7, PosteddtColumn), wsGL.Cells(GLEndRow, PosteddtColumn)).Copy wsGLAnalysis.Range("A27")
                ' Doc dt.
                    wsGL.Range(wsGL.Cells(7, DocdtColumn), wsGL.Cells(GLEndRow, DocdtColumn)).Copy wsGLAnalysis.Range("B27")
                ' Doc
                    wsGL.Range(wsGL.Cells(7, DocColumn), wsGL.Cells(GLEndRow, DocColumn)).Copy wsGLAnalysis.Range("C27")
                ' Memo/Description
                    wsGL.Range(wsGL.Cells(7, MemoColumn), wsGL.Cells(GLEndRow, MemoColumn)).Copy wsGLAnalysis.Range("D27")
                ' Location
                    wsGL.Range(wsGL.Cells(7, LocationColumn), wsGL.Cells(GLEndRow, LocationColumn)).Copy wsGLAnalysis.Range("E27")
                ' Division
                    wsGL.Range(wsGL.Cells(7, DivisionColumn), wsGL.Cells(GLEndRow, DivisionColumn)).Copy wsGLAnalysis.Range("F27")
                ' Funding Source
                    wsGL.Range(wsGL.Cells(7, FundingSourceColumn), wsGL.Cells(GLEndRow, FundingSourceColumn)).Copy wsGLAnalysis.Range("G27")
                ' Debt Services Series
                    wsGL.Range(wsGL.Cells(7, DebtServicesColumn), wsGL.Cells(GLEndRow, DebtServicesColumn)).Copy wsGLAnalysis.Range("H27")
                ' Txn No
                    wsGL.Range(wsGL.Cells(7, TxnNoColumn), wsGL.Cells(GLEndRow, TxnNoColumn)).Copy wsGLAnalysis.Range("I27")
                ' JNL
                    wsGL.Range(wsGL.Cells(7, JNLColumn), wsGL.Cells(GLEndRow, JNLColumn)).Copy wsGLAnalysis.Range("J27")
                ' Debit
                    wsGL.Range(wsGL.Cells(7, DebitColumn), wsGL.Cells(GLEndRow, DebitColumn)).Copy wsGLAnalysis.Range("K27")
                ' Credit
                    wsGL.Range(wsGL.Cells(7, CreditColumn), wsGL.Cells(GLEndRow, CreditColumn)).Copy wsGLAnalysis.Range("L27")
                ' Balance
                    wsGL.Range(wsGL.Cells(7, BalanceColumn), wsGL.Cells(GLEndRow, BalanceColumn)).Copy wsGLAnalysis.Range("M27")
                
            ' Turn off CutCopyMode
            Application.CutCopyMode = False

        ' After those are copied over, close the 'wsGL' worksheet without saving changes. If it already exists, don't save it.
            If Dir(FolderPath & SchoolName & " " & AccountPlusYear & " GL - " & Format(wsGL.Range("B4").Value, "YYYY.MM.DD") & " - " & Format(wsGL.Range("B5").Value, "YYYY.MM.DD") & ".xlsx") = "" Then
                wbGL.SaveAs FileName:=FolderPath & SchoolName & " " & AccountPlusYear & " GL - " & Format(wsGL.Range("B4").Value, "YYYY.MM.DD") & " - " & Format(wsGL.Range("B5").Value, "YYYY.MM.DD") & ".xlsx", FileFormat:=xlOpenXMLWorkbook
            End If
            wbGL.Close SaveChanges:=False
            
        ' Loop through column A to find the last row of each month to add a borderline beneath the month (to help separate them and improve readability) - borderline should be for columns A:M
            ' Find last used row in column A then subtract 4. That is all the data that will be worked with.
                GLAnalysisLastRow = wsGLAnalysis.Cells(wsGLAnalysis.Rows.Count, "A").End(xlUp).Row - 4
            ' Start at row 29 and go to the last row ('GLAnalysisLastRow') - 1 (it will cause an error with the data below the last row since it is not a date value)
                For CurrentRow = 29 To GLAnalysisLastRow - 1
            ' Check if the current Row's month is equal to the previous row's month.
                    If Month(wsGLAnalysis.Cells(CurrentRow, "A").Value) <> Month(wsGLAnalysis.Cells(CurrentRow + 1, "A").Value) Then
            ' If it changed. Put a borderline in columns A:M for that row.
                        With wsGLAnalysis.Range("A" & CurrentRow & ":M" & CurrentRow).Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .Color = RGB(0, 0, 0)
                        End With
                    End If
                Next CurrentRow

        ' Highlight all of the "Inschool Deposits" in light green (Columns A:M)
            ' Condition: Column D contains "*In-School Deposit"
            For CurrentRow = 27 To GLAnalysisLastRow
                If InStr(1, wsGLAnalysis.Cells(CurrentRow, "D").Value, "In-School Deposit", vbTextCompare) > 0 Then
                    wsGLAnalysis.Range("A" & CurrentRow & ":M" & CurrentRow).Interior.Color = RGB(198, 224, 180)
                End If
            Next CurrentRow


        ' Highlight all of the "Website Deposits" in light blue (Columns A:M)
            ' Condtions: (Column D does not contain "*In-School Deposit") AND ((Column C is not empty) OR (Column D contains "*This is a duplicate" - {Column C may be empty}))
            For CurrentRow = 27 To GLAnalysisLastRow
                If InStr(1, wsGLAnalysis.Cells(CurrentRow, "D").Value, "In-School Deposit", vbTextCompare) = 0 Then
                    If (wsGLAnalysis.Cells(CurrentRow, "C").Value <> "") Or (InStr(1, wsGLAnalysis.Cells(CurrentRow, "D").Value, "This is a duplicate", vbTextCompare) > 0) Then
                        wsGLAnalysis.Range("A" & CurrentRow & ":M" & CurrentRow).Interior.Color = RGB(221, 235, 247)
                    End If
                End If
            Next CurrentRow
            
        
        
        ' Start Populating the top portion of the worksheet.
            ' Make 'Verdana', Font Size 9 - the default for the worksheet.
                With wsGLAnalysis.Cells
                    .Font.Name = "Verdana"
                    .Font.Size = 9
                End With
                
            ' Rows 1:5 are dedicated to Pre-FY (Fiscal Year) Billings
                ' A1: "Pre-FY" - Make Bold and Underline it
                    With wsGLAnalysis.Range("A1")
                        .Value = "Pre-FY"
                        .Font.Bold = True
                        .Font.Underline = xlUnderlineStyleSingle
                    End With
                ' A2: "Billings"
                    wsGLAnalysis.Range("A2").Value = "Billings"
                ' A3: "Website Deposits" - Color cell light blue
                    With wsGLAnalysis.Range("A3")
                        .Value = "Website Deposits"
                        .Interior.Color = RGB(221, 235, 247)
                    End With
                ' A4: "Inschool Deposits" - Color cell light green
                    With wsGLAnalysis.Range("A4")
                        .Value = "Inschool Deposits"
                        .Interior.Color = RGB(198, 224, 180)
                    End With
                ' A5: "Totals"
                    wsGLAnalysis.Range("A5").Value = "Total"
                    
                ' Create the headers for each month:
                    ' B1:
                        wsGLAnalysis.Range("B1").Value = "'Jan 20" & (FY - 1)
                    ' C1:
                        wsGLAnalysis.Range("C1").Value = "'Feb 20" & (FY - 1)
                    ' D1:
                        wsGLAnalysis.Range("D1").Value = "'Mar 20" & (FY - 1)
                    ' E1:
                        wsGLAnalysis.Range("E1").Value = "'Apr 20" & (FY - 1)
                    ' F1:
                        wsGLAnalysis.Range("F1").Value = "'May 20" & (FY - 1)
                    ' G1:
                        wsGLAnalysis.Range("G1").Value = "'Jun 20" & (FY - 1)
                    ' H1 (The totals for Pre-FY):
                        wsGLAnalysis.Range("H1").Value = "Balance"
                
                ' Create the formulas for each
                    ' B2:
                        wsGLAnalysis.Range("B2").Formula = "=SUMIFS($K27:$K" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B1,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B1,0)+1,$C27:$C" & GLAnalysisLastRow & ","""",$D27:$D" & GLAnalysisLastRow & ",""<>*In-School Deposit*"",$D27:$D" & GLAnalysisLastRow & ",""<>*This is a duplicate*"")-SUMIFS($L27:$L" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B1,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B1,0)+1,$C27:$C" & GLAnalysisLastRow & ","""",$D27:$D" & GLAnalysisLastRow & ",""<>*In-School Deposit*"")"
                    ' B3:
                        wsGLAnalysis.Range("B3").Formula = "=SUMIFS($L27:$L" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B1,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B1,0)+1,$C27:$C" & GLAnalysisLastRow & ",""<>"",$D27:$D" & GLAnalysisLastRow & ",""<>*In-School Deposit*"")-SUMIFS($K27:$K" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B1,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B1,0)+1,$C27:$C" & GLAnalysisLastRow & ",""<>"",$D27:$D" & GLAnalysisLastRow & ",""<>*In-School Deposit*"")-SUMIFS($K27:$K" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B1,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B1,0)+1,$D27:$D" & GLAnalysisLastRow & ",""=*This is a duplicate*"")"
                    ' B4:
                        wsGLAnalysis.Range("B4").Formula = "=SUMIFS($L27:$L" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B1,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B1,0)+1,$D27:$D" & GLAnalysisLastRow & ",""=*In-School Deposit*"")+SUMIFS($K27:$K" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B1,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B1,0)+1,$D27:$D" & GLAnalysisLastRow & ",""=*In-School Deposit*"")"
                    ' B5: =B2-B3-B4
                        wsGLAnalysis.Range("B5").Formula = "=B2-B3-B4"
                        ' Fill these formulas across from column B over to column G
                            wsGLAnalysis.Range("B2:B5").AutoFill Destination:=wsGLAnalysis.Range("B2:G5")
                            
                    ' Totals
                        ' H2:
                            wsGLAnalysis.Range("H2").Formula = "=SUM(B2:G2)"
                        ' H3:
                            wsGLAnalysis.Range("H3").Formula = "=SUM(B3:G3)"
                        ' H4:
                            wsGLAnalysis.Range("H4").Formula = "=SUM(B4:G4)"
                        ' H5:
                            wsGLAnalysis.Range("H5").Formula = "=SUM(B5:G5)"
                ' Make H1:H5 Bold
                    wsGLAnalysis.Range("H1:H5").Font.Bold = True
                
                ' Add a Borderline under A4:H4
                    With wsGLAnalysis.Range("A4:H4").Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .Color = RGB(0, 0, 0)
                    End With
                
                ' Change B2:H5 to 'Accounting' Format
                    wsGLAnalysis.Range("B2:H5").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                
                
            ' Rows 7:11 are dedicated to FY (Fiscal Year) Billings
                ' A7: "FY" - Make Bold and Underline it
                    With wsGLAnalysis.Range("A7")
                        .Value = "FY"
                        .Font.Bold = True
                        .Font.Underline = xlUnderlineStyleSingle
                    End With
                ' A8: "Billings"
                    wsGLAnalysis.Range("A8").Value = "Billings"
                ' A9: "Website Deposits" - Color cell light blue
                    With wsGLAnalysis.Range("A9")
                        .Value = "Website Deposits"
                        .Interior.Color = RGB(221, 235, 247)
                    End With
                ' A10: "Inschool Deposits" - Color cell light green
                    With wsGLAnalysis.Range("A10")
                        .Value = "Inschool Deposits"
                        .Interior.Color = RGB(198, 224, 180)
                    End With
                ' A11: "Totals"
                    wsGLAnalysis.Range("A11").Value = "Total"

                ' Create the headers for each month:
                    ' B7:
                        wsGLAnalysis.Range("B7").Value = "'Jul 20" & FY - 1
                    ' C7:
                        wsGLAnalysis.Range("C7").Value = "'Aug 20" & FY - 1
                    ' D7:
                        wsGLAnalysis.Range("D7").Value = "'Sep 20" & FY - 1
                    ' E7:
                        wsGLAnalysis.Range("E7").Value = "'Oct 20" & FY - 1
                    ' F7:
                        wsGLAnalysis.Range("F7").Value = "'Nov 20" & FY - 1
                    ' G7:
                        wsGLAnalysis.Range("G7").Value = "'Dec 20" & FY - 1
                    ' H7:
                        wsGLAnalysis.Range("H7").Value = "'Jan 20" & FY
                    ' I7:
                        wsGLAnalysis.Range("I7").Value = "'Feb 20" & FY
                    ' J7:
                        wsGLAnalysis.Range("J7").Value = "'Mar 20" & FY
                    ' K7:
                        wsGLAnalysis.Range("K7").Value = "'Apr 20" & FY
                    ' L7:
                        wsGLAnalysis.Range("L7").Value = "'May 20" & FY
                    ' M7:
                        wsGLAnalysis.Range("M7").Value = "'Jun 20" & FY
                    ' N7 (The totals for 'Pre-FY' AND 'FY'):
                        wsGLAnalysis.Range("N7").Value = "Balance"
                        
                ' Create the formulas for each
                    ' B8:
                        wsGLAnalysis.Range("B8").Formula = "=SUMIFS($K27:$K" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B7,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B7,0)+1,$C27:$C" & GLAnalysisLastRow & ","""",$D27:$D" & GLAnalysisLastRow & ",""<>*In-School Deposit*"",$D27:$D" & GLAnalysisLastRow & ",""<>*This is a duplicate*"")-SUMIFS($L27:$L" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B7,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B7,0)+1,$C27:$C" & GLAnalysisLastRow & ","""",$D27:$D" & GLAnalysisLastRow & ",""<>*In-School Deposit*"")"
                    ' B9:
                        wsGLAnalysis.Range("B9").Formula = "=SUMIFS($L27:$L" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B7,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B7,0)+1,$C27:$C" & GLAnalysisLastRow & ",""<>"",$D27:$D" & GLAnalysisLastRow & ",""<>*In-School Deposit*"")-SUMIFS($K27:$K" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B7,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B7,0)+1,$C27:$C" & GLAnalysisLastRow & ",""<>"",$D27:$D" & GLAnalysisLastRow & ",""<>*In-School Deposit*"")-SUMIFS($K27:$K" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B7,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B7,0)+1,$D27:$D" & GLAnalysisLastRow & ",""=*This is a duplicate*"")"
                    ' B10:
                        wsGLAnalysis.Range("B10").Formula = "=SUMIFS($L27:$L" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B7,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B7,0)+1,$D27:$D" & GLAnalysisLastRow & ",""=*In-School Deposit*"")+SUMIFS($K27:$K" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B7,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B7,0)+1,$D27:$D" & GLAnalysisLastRow & ",""=*In-School Deposit*"")"
                    ' B11:
                        wsGLAnalysis.Range("B11").Formula = "=B8-B9-B10"
                        ' Fill these formulas across from column B over to column M
                            wsGLAnalysis.Range("B8:B11").AutoFill Destination:=wsGLAnalysis.Range("B8:M11")
                    ' Totals
                        ' N8:
                            wsGLAnalysis.Range("N8").Value = "=H2+SUM(B8:M8)"
                        ' N9:
                            wsGLAnalysis.Range("N9").Value = "=H3+SUM(B9:M9)"
                        ' N10:
                            wsGLAnalysis.Range("N10").Value = "=H4+SUM(B10:M10)"
                        ' N11:
                            wsGLAnalysis.Range("N11").Value = "=H5+SUM(B11:M11)"
                ' Make N7:N11 Bold
                    wsGLAnalysis.Range("N7:N11").Font.Bold = True
                ' Add a Borderline under A10:N10
                    With wsGLAnalysis.Range("A10:N10").Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .Color = RGB(0, 0, 0)
                    End With
                ' Change B8:N11 to 'Accounting' Format
                    wsGLAnalysis.Range("B8:N11").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                
            ' Rows 13:17 are dedicated to Post-FY (Fiscal Year) Billings
                ' A13: "Post-FY" - Make Bold and Underline it
                    With wsGLAnalysis.Range("A13")
                        .Value = "Post-FY"
                        .Font.Bold = True
                        .Font.Underline = xlUnderlineStyleSingle
                    End With
                ' A14: "Billings"
                    wsGLAnalysis.Range("A14").Value = "Billings"
                ' A15: "Website Deposits" - Color cell light blue
                    With wsGLAnalysis.Range("A15")
                        .Value = "Website Deposits"
                        .Interior.Color = RGB(221, 235, 247)
                    End With
                ' A16: "Inschool Deposits" - Color cell light green
                    With wsGLAnalysis.Range("A16")
                        .Value = "Inschool Deposits"
                        .Interior.Color = RGB(198, 224, 180)
                    End With
                ' A17: "Totals"
                    wsGLAnalysis.Range("A17").Value = "Total"

                ' Create the headers for each month:
                    ' B13:
                        wsGLAnalysis.Range("B13").Value = "'Jul 20" & FY
                    ' C13:
                        wsGLAnalysis.Range("C13").Value = "'Aug 20" & FY
                    ' D13:
                        wsGLAnalysis.Range("D13").Value = "'Sep 20" & FY
                    ' E13:
                        wsGLAnalysis.Range("E13").Value = "'Oct 20" & FY
                    ' F13:
                        wsGLAnalysis.Range("F13").Value = "'Nov 20" & FY
                    ' G13:
                        wsGLAnalysis.Range("G13").Value = "'Dec 20" & FY
                    ' H13:
                        wsGLAnalysis.Range("H13").Value = "'Jan 20" & FY + 1
                    ' I13:
                        wsGLAnalysis.Range("I13").Value = "'Feb 20" & FY + 1
                    ' J13:
                        wsGLAnalysis.Range("J13").Value = "'Mar 20" & FY + 1
                    ' K13:
                        wsGLAnalysis.Range("K13").Value = "'Apr 20" & FY + 1
                    ' L13:
                        wsGLAnalysis.Range("L13").Value = "'May 20" & FY + 1
                    ' M13:
                        wsGLAnalysis.Range("M13").Value = "'Jun 20" & FY + 1
                    ' N13(The totals for 'Pre-FY' AND 'FY' AND 'POST-FY'):
                        wsGLAnalysis.Range("N13").Value = "Balance"
                        
                ' Create the formulas for each
                    ' B14:
                        wsGLAnalysis.Range("B14").Formula = "=SUMIFS($K27:$K" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B13,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B13,0)+1,$C27:$C" & GLAnalysisLastRow & ","""",$D27:$D" & GLAnalysisLastRow & ",""<>*In-School Deposit*"",$D27:$D" & GLAnalysisLastRow & ",""<>*This is a duplicate*"")-SUMIFS($L27:$L" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B13,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B13,0)+1,$C27:$C" & GLAnalysisLastRow & ","""",$D27:$D" & GLAnalysisLastRow & ",""<>*In-School Deposit*"")"
                    ' B15:
                        wsGLAnalysis.Range("B15").Formula = "=SUMIFS($L27:$L" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B13,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B13,0)+1,$C27:$C" & GLAnalysisLastRow & ",""<>"",$D27:$D" & GLAnalysisLastRow & ",""<>*In-School Deposit*"")-SUMIFS($K27:$K" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B13,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B13,0)+1,$C27:$C" & GLAnalysisLastRow & ",""<>"",$D27:$D" & GLAnalysisLastRow & ",""<>*In-School Deposit*"")-SUMIFS($K27:$K" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B13,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B13,0)+1,$D27:$D" & GLAnalysisLastRow & ",""=*This is a duplicate*"")"
                    ' B16:
                        wsGLAnalysis.Range("B16").Formula = "=SUMIFS($L27:$L" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B13,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B13,0)+1,$D27:$D" & GLAnalysisLastRow & ",""=*In-School Deposit*"")+SUMIFS($K27:$K" & GLAnalysisLastRow & ",$B27:$B" & GLAnalysisLastRow & ","">=""&EOMONTH(B13,-1)+1,$B27:$B" & GLAnalysisLastRow & ",""<""&EOMONTH(B13,0)+1,$D27:$D" & GLAnalysisLastRow & ",""=*In-School Deposit*"")"
                    ' B17:
                        wsGLAnalysis.Range("B17").Formula = "=B14-B15-B16"
                        ' Fill these formulas across from column B over to column M
                            wsGLAnalysis.Range("B14:B17").AutoFill Destination:=wsGLAnalysis.Range("B14:M17")
                    ' Totals
                    ' N14:
                        wsGLAnalysis.Range("N14").Value = "=N8+SUM(B14:M14)"
                    ' N15:
                        wsGLAnalysis.Range("N15").Value = "=N9+SUM(B15:M15)"
                    ' N16:
                        wsGLAnalysis.Range("N16").Value = "=N10+SUM(B16:M16)"
                    ' N17:
                        wsGLAnalysis.Range("N17").Value = "=N11+SUM(B17:M17)"
                ' Make N13:N17 Bold
                    wsGLAnalysis.Range("N13:N17").Font.Bold = True
                ' Add a Borderline under A16:N16
                    With wsGLAnalysis.Range("A16:N16").Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .Color = RGB(0, 0, 0)
                    End With
                ' Change B14:N17 to 'Accounting' Format
                    wsGLAnalysis.Range("B14:N17").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                
            ' Change Column Width of A:D and G:H and K:M
                ' Change Column Width of A:D and G:H
                    wsGLAnalysis.Columns("A").ColumnWidth = 16.22
                    wsGLAnalysis.Columns("B").ColumnWidth = 22
                    wsGLAnalysis.Columns("C").ColumnWidth = 27
                    wsGLAnalysis.Columns("D").ColumnWidth = 29
                    wsGLAnalysis.Columns("G").ColumnWidth = 27
                    wsGLAnalysis.Columns("H").ColumnWidth = 19
                    wsGLAnalysis.Columns("K:M").ColumnWidth = 16.89
  
            ' Group 'Pre-FY' together (Rows 1:5)
                wsGLAnalysis.Rows("1:5").Group
            
            ' Group 'FY' together (Rows 7:11)
                wsGLAnalysis.Rows("7:11").Group
            
            ' Group 'Post-FY' together (Rows 13:17)
                wsGLAnalysis.Rows("13:17").Group
            
            ' Add 'Download Date:' to A19 - Make bold
                With wsGLAnalysis.Range("A19")
                    .Value = "Download Date:"
                    .Font.Bold = True
                End With
                
            ' Add Subtotals to columns K, L, and M above the "Debit", "Credit", "Balance" headers. Also format as 'Accounting'. Change the background cell color to a royal blue and the font to white.
                With wsGLAnalysis.Range("K26:M26")
                    ' Apply subtotal formulas
                    .Cells(1, 1).Formula = "=SUBTOTAL(9,K29:K" & GLAnalysisLastRow & ")"
                    .Cells(1, 2).Formula = "=SUBTOTAL(9,L29:L" & GLAnalysisLastRow & ")"
                    .Cells(1, 3).Formula = "=K26-L26"
                
                    ' Apply 'Accounting' format
                    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                
                    ' Change background to blue and font to white
                    .Interior.Color = RGB(65, 105, 225)
                    .Font.Color = RGB(255, 255, 255)
                    .Font.Bold = True
                End With

            
            ' Freeze Rows 28 and above.
                With wsGLAnalysis
                    .Activate
                    .Range("A29").Select
                    ActiveWindow.FreezePanes = True
                End With
                
            ' Create a conditional format for if B19 is blank. Make it red
                With wsGLAnalysis.Range("B19").FormatConditions
                ' If B19 is blank (empty)
                    .Add Type:=xlExpression, Formula1:="=ISBLANK(B19)"
                ' Then make the cell red.
                    .Item(1).Interior.Color = RGB(255, 0, 0)
                End With


'____________________________________________________________________________________________________________________________________
    ' Create a new worksheet with the naming convention of: SchoolName & "Variance Analysis" & ['FY'-1 & "." & 'FY']
        ' Using the variable of 'SheetExists' (Telling whether the "Add or Split Out Reports" worksheet exists) - If it exists, put the new worksheet before it, otherwise, add it to the end of the worksheets list.
            If Not SheetExists Then
                Set wsVarianceAnalysis = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets(wbMacro.Worksheets.Count))
            Else
                Set wsVarianceAnalysis = wbMacro.Worksheets.Add(Before:=wbMacro.Worksheets("Add or Split Out Reports"))
            End If
        ' Rename the new worksheet.
            wsVarianceAnalysis.Name = Replace(SchoolName, " ", "") & " Variance Analysis " & (FY - 1) & "." & FY

            
        ' Start populating everything.
            ' Make the default font "Arial" and the size 10 Point.
                With wsVarianceAnalysis.Cells.Font
                    .Name = "Arial"
                    .Size = 10
                End With
                
            ' A1:B1
                With wsVarianceAnalysis.Range("A1:B1")
                ' Merge and center cells together.
                    .MergeCells = True
                    .HorizontalAlignment = xlCenter
                ' Add 'SchoolName'.
                    .Value = SchoolName
                ' Change Font to bold and 16 point size
                    .Font.Bold = True
                    .Font.Size = 16
                ' Change cell color to light blue.
                    .Interior.Color = RGB(175, 255, 255)
                ' Add border all around the cells
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .Borders.Color = RGB(105, 105, 105)
                ' Change cell size
                    .ColumnWidth = 12.14
                End With
                
            ' Set up C1:M1
                With wsVarianceAnalysis.Range("C1:M1")
                ' Merge and center cells together
                    .MergeCells = True
                    .HorizontalAlignment = xlCenter
                ' Add "BLACKBAUD DETAILS" to it.
                    .Value = "BLACKBAUD DETAILS"
                ' Change Font to bold and 16 point size
                    .Font.Bold = True
                    .Font.Size = 16
                ' Change cell color to light green.
                    .Interior.Color = RGB(198, 224, 180)
                ' Add border all around the cells
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .Borders.Color = RGB(105, 105, 105)
                ' Change cell size
                    .ColumnWidth = 13.43
                End With
                
            ' Set up O1:Q1
                With wsVarianceAnalysis.Range("O1:Q1")
                ' Merge and center cells together
                    .MergeCells = True
                    .HorizontalAlignment = xlCenter
                ' Add "INTACCT DETAILS" to it.
                    .Value = "INTACCT DETAILS"
                ' Change Font to bold and 16 point size
                    .Font.Bold = True
                    .Font.Size = 16
                ' Change cell color to yellow.
                    .Interior.Color = RGB(255, 255, 0)
                ' Add border all around the cells
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .Borders.Color = RGB(105, 105, 105)
                ' Change cell size
                    .ColumnWidth = 14#
                End With
                
            ' Set up S1:U1
            With wsVarianceAnalysis.Range("S1:U1")
            ' Merge and center cells together
                .MergeCells = True
                .HorizontalAlignment = xlCenter
            ' Add "VARIANCES" to it.
                .Value = "VARIANCES"
            ' Change Font to bold and 16 point size
                .Font.Bold = True
                .Font.Size = 16
            ' Change cell color to yellow.
                .Interior.Color = RGB(0, 176, 240)
            ' Add border all around the cells
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Borders.Color = RGB(105, 105, 105)
            ' Change cell size
                .ColumnWidth = 13.86
            End With
            
            ' Change column width of J, N, P, R, V, and W
                wsVarianceAnalysis.Range("J1").ColumnWidth = 2
                wsVarianceAnalysis.Range("N1").ColumnWidth = 1.71
                wsVarianceAnalysis.Range("P1").ColumnWidth = 15.57
                wsVarianceAnalysis.Range("R1").ColumnWidth = 1.71
                wsVarianceAnalysis.Range("V1").ColumnWidth = 1.43
                wsVarianceAnalysis.Range("W1").ColumnWidth = 14.71
            
            ' Wrap text data in row 2
                wsVarianceAnalysis.Rows(2).WrapText = True
            ' Fill in the headers for row 2
                ' A2
                    wsVarianceAnalysis.Range("A2").Value = "Month"
                ' B2
                    wsVarianceAnalysis.Range("B2").Value = "Fiscal Year"
                ' C2:
                    wsVarianceAnalysis.Range("C2").Value = "Total Billed"
                ' D2: **Note** Gloria suggested a change from 'Smart' to "Blackbaud'
                    wsVarianceAnalysis.Range("D2").Value = "Total Paid to Smart (settled)"
                ' E2: **Note** Gloria suggested a change from 'Smart' to "Blackbaud'
                    wsVarianceAnalysis.Range("E2").Value = "Total Paid to Smart (pending transfer)"
                ' F2:
                    wsVarianceAnalysis.Range("F2").Value = "Total Paid In School (settled)"
                ' G2:
                    wsVarianceAnalysis.Range("G2").Value = "Total Paid In School (in-process)"
                ' H2:
                    wsVarianceAnalysis.Range("H2").Value = "Balance Due"
                ' I2:
                    wsVarianceAnalysis.Range("I2").Value = "Total Balance Due"
                    
                ' K2:
                    wsVarianceAnalysis.Range("K2").Value = "Total Billing in Period"
                ' L2:
                    wsVarianceAnalysis.Range("L2").Value = "In School Pmts in Period"
                ' M2: **Note** Gloria suggested a change from 'SMART' to "BLACKBAUD'
                    wsVarianceAnalysis.Range("M2").Value = "SMART Payments in Period"
                    
                ' O2:
                    wsVarianceAnalysis.Range("O2").Value = "Total Billings on GL"
                ' P2:
                    wsVarianceAnalysis.Range("P2").Value = "In School Pmts on GL"
                ' Q2:
                    wsVarianceAnalysis.Range("Q2").Value = "Website Deposit"
                
                ' S2:
                    wsVarianceAnalysis.Range("S2").Value = "Total Billing Variance"
                ' T2:
                    wsVarianceAnalysis.Range("T2").Value = "In School Pmts Variance"
                ' U2:
                    wsVarianceAnalysis.Range("U2").Value = "Website Deposit Variance"
                
            ' Format Row 2
                ' A2:M2
                    With wsVarianceAnalysis.Range("A2:M2")
                    ' Center Horizontally and Vertically.
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    ' Change Font to bold
                        .Font.Bold = True
                    ' Add border all around the each cell
                        .Borders.LineStyle = xlContinuous
                        .Borders.Weight = xlThin
                        .Borders.Color = RGB(105, 105, 105)
                    End With
                    
                ' O2:Q2
                    With wsVarianceAnalysis.Range("O2:Q2")
                    ' Center Horizontally and Vertically.
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    ' Change Font to bold
                        .Font.Bold = True
                    ' Add border all around the each cell
                        .Borders.LineStyle = xlContinuous
                        .Borders.Weight = xlThin
                        .Borders.Color = RGB(105, 105, 105)
                    End With
                    
                ' S2:U2
                    With wsVarianceAnalysis.Range("S2:U2")
                    ' Center Horizontally and Vertically.
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    ' Change Font to bold
                        .Font.Bold = True
                    ' Add border all around the each cell
                        .Borders.LineStyle = xlContinuous
                        .Borders.Weight = xlThin
                        .Borders.Color = RGB(105, 105, 105)
                    End With
                    
                ' Freeze (Columns A and B) AND (Rows 1 and 2)
                    With wsVarianceAnalysis
                        .Activate
                            With ActiveWindow
                                .SplitColumn = 2
                                .SplitRow = 2
                                .FreezePanes = True
                            End With
                    End With
                    
            ' Set up the formatting and formulas within the table.
                ' Create a variable for the last row of the table call it: 'LastRow'
                    ' The purpose of this variable is to determine the row we want to end on for the table. It will be a variable, so it can be change and all formulas will update with the change.
                    LastRow = 92
                ' Create a variable called 'FYLoop', to help with naming the months in the table. **Note** This will update within the For Loop coming up.
                    FYLoop = FY - 1
                ' Create a variable for 'LoopMonth', Starting at month 2 (in the upcoming For Loop this variable will be used to go back 1 day from the start 'LoopMonth', allowing for the date to be the last day of each month)
                    LoopMonth = 2
                ' Set 'LoopNumber' equal to 0. This will change for the month the loop is on. It will change to help create the formulas needed to make everything dynamic.
                    LoopNumber = 0
                    
                ' Create a For Loop going from row 3 to the 'LastRow' Variable.
                    ' Use a For Loop where there are 3 different types of rows. Similiar to the current Recon. 2 light blue (Case 0 and Case 1), and 1 grey (Case 2)
                    ' To do this we can use 'Mod'. Or 'Mod'ulus Operator returns the value of a remainder. For example 7 Mod 3 will return 1 (it means 7/3 = __) the answer is [2 Remainder 1]. Mod will only return 1
                    ' The loop will hold a value 'i' and represents the row number it is currently working on. 'Mod' will determine the 'Case' or the row type and how to populate it.
                    ' Case 0 is mostly used as a separation row. Case 1 will hold the 'A/R Balance' Report values. Case 2 (Grey) will hold the totals and most of the formulas
                    For i = 3 To LastRow
                            Select Case (i - 3) Mod 3
                        ' Case 0 (no remainder) - 1st light blue row (mostly used as a separation row)
                        Case 0
                            ' In columns C:H color the cells with light blue.
                                With wsVarianceAnalysis.Range(wsVarianceAnalysis.Cells(i, "C"), wsVarianceAnalysis.Cells(i, "H"))
                                ' Color the cells with light blue
                                    .Interior.Color = RGB(189, 215, 238)
                                End With
                                
                            ' In column M have a border on the right edge of the cell.
                                With wsVarianceAnalysis.Range("M" & i).Borders(xlEdgeRight)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(0, 0, 0)
                                End With
                                
                            ' In column O have a border on the left edge of the cell.
                                With wsVarianceAnalysis.Range("O" & i).Borders(xlEdgeLeft)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(0, 0, 0)
                                End With
                            ' In column Q have a border on the right edge of the cell.
                                With wsVarianceAnalysis.Range("Q" & i).Borders(xlEdgeRight)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(0, 0, 0)
                                End With
                                
                            ' In column S have a border on the left edge of the cell.
                                With wsVarianceAnalysis.Range("S" & i).Borders(xlEdgeLeft)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(0, 0, 0)
                                End With
                                
                            ' In column U have a border on the right edge of the cell.
                                With wsVarianceAnalysis.Range("U" & i).Borders(xlEdgeRight)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(0, 0, 0)
                                End With
                                
                            ' If the row number is greater than 3
                                If i <> 3 Then
                                    With wsVarianceAnalysis.Range("I" & i)
                                        .Formula = "=E" & i & "+H" & i
                                    ' Change cell format to 'Accounting'
                                        .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                                    End With
                                End If
                                
                                
                                
                        ' Case 1 (Remainder 1) - 2nd light blue row ('A/R Balance' report values) Row
                        Case 1
                            ' Set the values of LoopNumber, GLColumn, GLBillings, GLWebsiteDeposits, GLInSchoolDeposits
                                LoopNumber = LoopNumber + 1
                                
                                ' Check 'LoopNumber' value to determine which column letter to use from '12010 GL' worksheet.
                                    If (LoopNumber = 1) Or (LoopNumber = 7) Or (LoopNumber = 19) Then
                                        GLColumn = "B"
                                    ElseIf (LoopNumber = 2) Or (LoopNumber = 8) Or (LoopNumber = 20) Then
                                        GLColumn = "C"
                                    ElseIf (LoopNumber = 3) Or (LoopNumber = 9) Or (LoopNumber = 21) Then
                                        GLColumn = "D"
                                    ElseIf (LoopNumber = 4) Or (LoopNumber = 10) Or (LoopNumber = 22) Then
                                        GLColumn = "E"
                                    ElseIf (LoopNumber = 5) Or (LoopNumber = 11) Or (LoopNumber = 23) Then
                                        GLColumn = "F"
                                    ElseIf (LoopNumber = 6) Or (LoopNumber = 12) Or (LoopNumber = 24) Then
                                        GLColumn = "G"
                                    ElseIf (LoopNumber = 13) Or (LoopNumber = 25) Then
                                        GLColumn = "H"
                                    ElseIf (LoopNumber = 14) Or (LoopNumber = 26) Then
                                        GLColumn = "I"
                                    ElseIf (LoopNumber = 15) Or (LoopNumber = 27) Then
                                        GLColumn = "J"
                                    ElseIf (LoopNumber = 16) Or (LoopNumber = 28) Then
                                        GLColumn = "K"
                                    ElseIf (LoopNumber = 17) Or (LoopNumber = 29) Then
                                        GLColumn = "L"
                                    ElseIf (LoopNumber = 18) Or (LoopNumber = 30) Then
                                        GLColumn = "M"
                                    End If
                                    
                                
                                ' Check 'LoopNumber' again, this time to determine the row needed for the formulas to populate the data from '12010 GL' Worksheet.
                                    If LoopNumber < 7 Then
                                        GLBillings = 2
                                        GLWebsiteDeposits = 3
                                        GLInSchoolDeposits = 4
                                    ElseIf (LoopNumber > 6) And (LoopNumber < 19) Then
                                        GLBillings = 8
                                        GLWebsiteDeposits = 9
                                        GLInSchoolDeposits = 10
                                    Else
                                        GLBillings = 14
                                        GLWebsiteDeposits = 15
                                        GLInSchoolDeposits = 16
                                    End If

                            ' In Column A
                                ' Make font bold
                                    wsVarianceAnalysis.Range("A" & i).Font.Bold = True
                                ' Place through date: 1/31/[YEAR] (For Example: for January... 1/31/2025)
                                    If LoopMonth = 13 Then
                                      wsVarianceAnalysis.Range("A" & i).Value = Application.WorksheetFunction.EoMonth(DateValue("12/1/20" & FYLoop), 0)
                                      wsVarianceAnalysis.Range("A" & i).NumberFormat = "m/d/yyyy"
                                    ElseIf LoopMonth = 14 Then
                                        LoopMonth = 2
                                        FYLoop = FYLoop + 1
                                        wsVarianceAnalysis.Range("A" & i).Value = DateValue(LoopMonth & "/1/20" & FYLoop) - 1
                                    Else
                                        wsVarianceAnalysis.Range("A" & i).Value = DateValue(LoopMonth & "/1/20" & FYLoop) - 1
                                    End If
                                    LoopMonth = LoopMonth + 1
                                    
                            ' In Column B put in the FY (For Example: 24/25)
                                wsVarianceAnalysis.Range("B" & i).Value = FY - 1 & "/" & FY
                                
                            ' In columns C:H
                                With wsVarianceAnalysis.Range(wsVarianceAnalysis.Cells(i, "C"), wsVarianceAnalysis.Cells(i, "H"))
                                ' Color the cells with light blue
                                    .Interior.Color = RGB(189, 215, 238)
                                ' Change format for the cells to 'Accounting'
                                    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                                End With
                                
                            ' In column I:
                                wsVarianceAnalysis.Range("I" & i).Formula = "=E" & i & "+H" & i
                                
                            ' In column K:
                                With wsVarianceAnalysis.Range("K" & i)
                                ' Change format for the cell to 'Accounting'
                                    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                                    .Formula = "=IF(C" & (i + 1) & "=0, 0, C" & (i + 1) & "-C" & (i - 2) & ")"
                                End With
                                
                            ' In column O:
                                With wsVarianceAnalysis.Range("O" & i)
                                ' Have a border on the left edge
                                    With .Borders(xlEdgeLeft)
                                        .LineStyle = xlContinuous
                                        .Weight = xlThin
                                        .Color = RGB(0, 0, 0)
                                    End With
                                ' Color the cell light blue
                                    .Interior.Color = RGB(189, 215, 238)
                                ' Change format for the cell to 'Accounting'
                                    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                                    .Formula = "='" & GLWorksheetName & "'!" & GLColumn & GLBillings
                                End With
                                
                            ' In column S
                                With wsVarianceAnalysis.Range("S" & i)
                                ' Have a border on the left edge
                                    With .Borders(xlEdgeLeft)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(0, 0, 0)
                                    End With
                                ' Change format for the cell to 'Accounting'
                                    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                                    .Formula = "=K" & i & "-O" & i
                                End With
                                
                            ' In column M have a border on the right edge of the cell
                                With wsVarianceAnalysis.Range("M" & i).Borders(xlEdgeRight)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(0, 0, 0)
                                End With
                                
                            ' In column Q have a border on the right edge of the cell
                                With wsVarianceAnalysis.Range("Q" & i).Borders(xlEdgeRight)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(0, 0, 0)
                                End With
                                
                            ' In column U have a border on the right edge of the cell
                                With wsVarianceAnalysis.Range("U" & i).Borders(xlEdgeRight)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(0, 0, 0)
                                End With
                                
                                
                        ' Case 2 (Remainder 2) - Grey row (Where all formulas and totals populate)
                        Case 2
                            'In columns A:I
                                With wsVarianceAnalysis.Range(wsVarianceAnalysis.Cells(i, "A"), wsVarianceAnalysis.Cells(i, "I"))
                                ' Color the cells with grey
                                    .Interior.Color = RGB(217, 217, 217)
                                ' Change format for the cells to 'Accounting'
                                    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                                ' Make the font bold
                                    .Font.Bold = True
                                End With
                            ' In column M have a border on the right edge
                                With wsVarianceAnalysis.Range("M" & i).Borders(xlEdgeRight)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(0, 0, 0)
                                End With
                                
                            ' In column O have a border on the left edge
                                With wsVarianceAnalysis.Range("O" & i).Borders(xlEdgeLeft)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(0, 0, 0)
                                End With
                                
                            ' In column Q have a border on the right edge
                                With wsVarianceAnalysis.Range("Q" & i).Borders(xlEdgeRight)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(0, 0, 0)
                                End With
                                
                            ' In column S have a border on the left edge
                                With wsVarianceAnalysis.Range("S" & i).Borders(xlEdgeLeft)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(0, 0, 0)
                                End With
                                
                            ' In column U have a border on the right edge
                                With wsVarianceAnalysis.Range("U" & i).Borders(xlEdgeRight)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(0, 0, 0)
                                End With
                                
                            ' In column L
                                With wsVarianceAnalysis.Range("L" & i)
                                    .Formula = "=IF(F" & i & "=0, 0, F" & i & "-F" & (i - 3) & ")"
                                ' Change format for the cells to 'Accounting'
                                    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                                End With
                                
                            ' In column M
                                With wsVarianceAnalysis.Range("M" & i)
                                    .Formula = "=IF(D" & i & "=0, 0, D" & i & "-D" & (i - 3) & ")"
                                ' Change format for the cells to 'Accounting'
                                    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                                End With
                                
                            ' In column P
                                With wsVarianceAnalysis.Range("P" & i)
                                   .Formula = "='" & GLWorksheetName & "'!" & GLColumn & GLInSchoolDeposits
                                ' Change cell color to light blue
                                    .Interior.Color = RGB(189, 215, 238)
                                ' Change format for the cells to 'Accounting'
                                    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                                End With
                                
                            ' In column Q
                                With wsVarianceAnalysis.Range("Q" & i)
                                   .Formula = "='" & GLWorksheetName & "'!" & GLColumn & GLWebsiteDeposits
                                ' Change cell color to light blue
                                    .Interior.Color = RGB(189, 215, 238)
                                ' Change format for the cells to 'Accounting'
                                    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                                End With
                                
                            ' In column T
                                wsVarianceAnalysis.Range("T" & i).Formula = "=L" & i & "-P" & i
                                
                            ' In column U
                                wsVarianceAnalysis.Range("U" & i).Formula = "=M" & i & "-Q" & i
                                
                            ' In column B have the cell value = "Total
                                wsVarianceAnalysis.Cells(i, "B").Value = "Total"
                                
                            ' Create a For loop to populate the cells in columns C:I to add the total from the two cells above that row with the corresponding column. (For example: in C8 (grey row) the formula would take C6+C7 to give the value for C8)
                                For Col = 3 To 9 ' C = 3, D = 4, ..., I = 9
                                    wsVarianceAnalysis.Cells(i, Col).Formula = "=" & wsVarianceAnalysis.Cells(i - 2, Col).Address(False, False) & "+" & wsVarianceAnalysis.Cells(i - 1, Col).Address(False, False)
                                Next Col

                End Select
            Next i
                                
                                
                                
                                
            ' Add the borderlines below the table with data
                ' In columns A:M make a borderline under the 'LastRow' from the table.
                    With wsVarianceAnalysis.Range("A" & LastRow & ":M" & LastRow).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .Color = RGB(0, 0, 0)
                    End With
                    
                ' In columns O:Q make a borderline under the 'LastRow' from the table.
                    With wsVarianceAnalysis.Range("O" & LastRow & ":Q" & LastRow).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .Color = RGB(0, 0, 0)
                    End With
                    
                ' In columns S:U make a borderline under the 'LastRow' from the table.
                    With wsVarianceAnalysis.Range("S" & LastRow & ":U" & LastRow).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .Color = RGB(0, 0, 0)
                    End With
                    
                    
            ' Populate the data in the row beneath the table (Totals)
                ' In columns K:M
                    With wsVarianceAnalysis.Range("K" & (LastRow + 1) & ":M" & (LastRow + 1))
                    ' Add border line around the parameter of each individual cell in the selection.
                        With .Borders
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .Color = RGB(0, 0, 0)
                        End With
                    ' Change format for the cells to 'Accounting'
                        .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    ' Make the font in the cells bold.
                        .Font.Bold = True
                    End With

                ' In columsn O:Q
                    With wsVarianceAnalysis.Range("O" & (LastRow + 1) & ":Q" & (LastRow + 1))
                    ' Add border line around the parameter of each individual cell in the selection.
                        With .Borders
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .Color = RGB(0, 0, 0)
                        End With
                    ' Change format for the cells to 'Accounting'
                        .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    ' Make the font in the cells bold.
                        .Font.Bold = True
                    End With
                    
                ' In columns S:U
                    With wsVarianceAnalysis.Range("S" & (LastRow + 1) & ":U" & (LastRow + 1))
                    ' Add border line around the parameter of each individual cell in the selection.
                        With .Borders
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .Color = RGB(0, 0, 0)
                        End With
                    ' Change format for the cells to 'Accounting'
                        .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    ' Make the font in the cells bold.
                        .Font.Bold = True
                    End With


                ' Populate the formulas (LastRow + 1)
                    ' Column K:
                        wsVarianceAnalysis.Range("K" & (LastRow + 1)).Formula = "=SUM(K3:K" & LastRow & ")"
                    ' Column L:
                        wsVarianceAnalysis.Range("L" & (LastRow + 1)).Formula = "=SUM(L3:L" & LastRow & ")"
                    ' Column M:
                        wsVarianceAnalysis.Range("M" & (LastRow + 1)).Formula = "=SUM(M3:M" & LastRow & ")"
                        
                    ' Column O:
                        wsVarianceAnalysis.Range("O" & (LastRow + 1)).Formula = "=SUM(O3:O" & LastRow & ")"
                    ' Column P:
                        wsVarianceAnalysis.Range("P" & (LastRow + 1)).Formula = "=SUM(P3:P" & LastRow & ")"
                    ' Column Q:
                        wsVarianceAnalysis.Range("Q" & (LastRow + 1)).Formula = "=SUM(Q3:Q" & LastRow & ")"
                    
                    ' Column S:
                        wsVarianceAnalysis.Range("S" & (LastRow + 1)).Formula = "=SUM(S3:S" & LastRow & ")"
                    ' Column T:
                        wsVarianceAnalysis.Range("T" & (LastRow + 1)).Formula = "=SUM(T3:T" & LastRow & ")"
                    ' Column U:
                        wsVarianceAnalysis.Range("U" & (LastRow + 1)).Formula = "=SUM(U3:U" & LastRow & ")"
                        
                        
            ' Populate the data at the bottom of column L and M (LastRow + 3) : (LastRow + 5)
                ' Columns L(LastRow + 3):M(LastRow + 4):
                    With wsVarianceAnalysis.Range("L" & (LastRow + 3) & ":M" & (LastRow + 4))
                    ' Add the border lines around the parameter of the selection
                        With .Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .Color = RGB(0, 0, 0)
                        End With
                        With .Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .Color = RGB(0, 0, 0)
                        End With
                        With .Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .Color = RGB(0, 0, 0)
                        End With
                        With .Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .Color = RGB(0, 0, 0)
                        End With
                    End With
                    
                ' Columns L(LastRow + 5):M(LastRow + 5)
                    With wsVarianceAnalysis.Range("L" & (LastRow + 5) & ":M" & (LastRow + 5))
                    ' Add the border lines around the parameter of the selection
                        With .Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .Color = RGB(0, 0, 0)
                        End With
                        With .Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .Color = RGB(0, 0, 0)
                        End With
                        With .Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .Color = RGB(0, 0, 0)
                        End With
                        With .Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .Color = RGB(0, 0, 0)
                        End With
                    ' Make the font in the cells bold.
                        .Font.Bold = True
                    End With
                    
                ' L(LastRow + 3) = "Starting Bal"
                    wsVarianceAnalysis.Range("L" & (LastRow + 3)).Value = "Starting Bal"
                ' L(LastRow + 4) = "Activity"
                    wsVarianceAnalysis.Range("L" & (LastRow + 4)).Value = "Activity"
                ' L(LastRow + 5) = "Ending Bal"
                    wsVarianceAnalysis.Range("L" & (LastRow + 5)).Value = "Ending Bal"
                    
                ' Cells M(LastRow + 3):M(LastRow + 5)
                    ' Format cells to 'Accounting' format.
                        wsVarianceAnalysis.Range("M" & (LastRow + 3) & ":M" & (LastRow + 5)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        
'***************' M(LastRow + 3) Formula:
                    wsVarianceAnalysis.Range("M" & (LastRow + 3)).Formula = "=0"
                ' M(LastRow + 4) Formula: "=K" & (LastRow + 1) & "-L" & (LastRow + 1) & "-M" & (LastRow + 1)
                    wsVarianceAnalysis.Range("M" & (LastRow + 4)).Formula = "=K" & (LastRow + 1) & "-L" & (LastRow + 1) & "-M" & (LastRow + 1)
                ' M(LastRow + 5) Formula: "=SUM(M" & (LastRow + 3) & ":M" & (LastRow + 4) & ")"
                    wsVarianceAnalysis.Range("M" & (LastRow + 5)).Formula = "=SUM(M" & (LastRow + 3) & ":M" & (LastRow + 4) & ")"
                    
            ' Populate the data at the bottom of columns P:Q
                ' Top Portion
                    ' Columns P(LastRow + 3):Q(LastRow + 6)
                        With wsVarianceAnalysis.Range("P" & (LastRow + 3) & ":Q" & (LastRow + 6))
                        ' Add the border lines around the parameter of the selection
                            With .Borders(xlEdgeTop)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                            With .Borders(xlEdgeBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                            With .Borders(xlEdgeLeft)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                            With .Borders(xlEdgeRight)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                        ' Format the cells to 'Accounting'
                            .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        End With
                        
                    ' Columns P(LastRow + 7):Q(LastRow + 7)
                        With wsVarianceAnalysis.Range("P" & (LastRow + 7) & ":Q" & (LastRow + 7))
                        ' Add the border lines around the parameter of the selection
                            With .Borders(xlEdgeTop)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                            With .Borders(xlEdgeBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                            With .Borders(xlEdgeLeft)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                            With .Borders(xlEdgeRight)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                        ' Make the font in the cells bold.
                            .Font.Bold = True
                        ' Format the cells to 'Accounting'
                            .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        End With
                        
                    ' Columns Q(LastRow + 3):Q(LastRow + 5)
                        ' Change cell color to light blue
                            wsVarianceAnalysis.Range("Q" & (LastRow + 3) & ":Q" & (LastRow + 5)).Interior.Color = RGB(189, 215, 238)
                            
                    ' Values
                        ' P(LastRow + 5): "Reversing Entry"
                            wsVarianceAnalysis.Range("P" & (LastRow + 5)).Value = "Reversing Entry"
                        ' P(LastRow + 6): "Activity"
                            wsVarianceAnalysis.Range("P" & (LastRow + 6)).Value = "Activity"
                        ' P(LastRow + 7): "Ending GL"
                            wsVarianceAnalysis.Range("P" & (LastRow + 7)).Value = "Ending GL"
                            
                    ' Formulas
'***********************' Q(LastRow + 5) Formula:
                            wsVarianceAnalysis.Range("Q" & (LastRow + 5)).Formula = "=0"
                        ' Q(LastRow + 6) Formula: "=O" & (LastRow + 1) & "-P" & (LastRow + 1) & "-Q" & (LastRow + 1)
                            wsVarianceAnalysis.Range("Q" & (LastRow + 6)).Formula = "=O" & (LastRow + 1) & "-P" & (LastRow + 1) & "-Q" & (LastRow + 1)
                        ' Q(LastRow + 7) Formula: "=SUM(Q" & (LastRow + 3) & ":Q" & (LastRow + 6) & ")"
                            wsVarianceAnalysis.Range("Q" & (LastRow + 7)).Formula = "=SUM(Q" & (LastRow + 3) & ":Q" & (LastRow + 6) & ")"
                        
                ' Bottom Portion
                    ' Columns P(LastRow + 9):Q(LastRow + 12)
                        With wsVarianceAnalysis.Range("P" & (LastRow + 9) & ":Q" & (LastRow + 12))
                        ' Add the border lines around the parameter of the selection
                            With .Borders(xlEdgeTop)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                            With .Borders(xlEdgeBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                            With .Borders(xlEdgeLeft)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                            With .Borders(xlEdgeRight)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                        ' Format the cells to 'Accounting'
                            .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        End With
                        
                        
                    ' Columns P(LastRow + 13):Q(LastRow + 13)
                        With wsVarianceAnalysis.Range("P" & (LastRow + 13) & ":Q" & (LastRow + 13))
                        ' Add the border lines around the parameter of the selection
                            With .Borders(xlEdgeTop)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                            With .Borders(xlEdgeBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                            With .Borders(xlEdgeLeft)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                            With .Borders(xlEdgeRight)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(0, 0, 0)
                            End With
                        ' Format the cells to 'Accounting'
                            .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        ' Make the font in the cells bold.
                            .Font.Bold = True
                        End With
                        
                    ' Values
                        ' P(LastRow + 9): "Billing Variance"
                            wsVarianceAnalysis.Range("P" & (LastRow + 9)).Value = "Billing Variance"
                        ' P(LastRow + 10): "In School Var"
                            wsVarianceAnalysis.Range("P" & (LastRow + 10)).Value = "In School Var"
                        ' P(LastRow + 11): "Pmt Variance"
                            wsVarianceAnalysis.Range("P" & (LastRow + 11)).Value = "Pmt Variance"
                    ' Formulas
                        ' Q(LastRow + 9) Formula: "=S" & (LastRow + 1)
                            wsVarianceAnalysis.Range("Q" & (LastRow + 9)).Formula = "=S" & (LastRow + 1)
                        ' Q(LastRow + 10) Formula: "=-T" & (LastRow + 1)
                            wsVarianceAnalysis.Range("Q" & (LastRow + 10)).Formula = "=-T" & (LastRow + 1)
                        ' Q(LastRow + 11) Formula: "=-U" & (LastRow + 1)
                            wsVarianceAnalysis.Range("Q" & (LastRow + 11)).Formula = "=-U" & (LastRow + 1)
                        ' Q(LastRow + 13) Formula: "=SUM(Q" & (LastRow + 9) & ":Q" & (LastRow + 12) & ")"
                            wsVarianceAnalysis.Range("Q" & (LastRow + 13)).Formula = "=SUM(Q" & (LastRow + 9) & ":Q" & (LastRow + 12) & ")"
                            
                            
                ' Last Formula
                    ' Cell Q(LastRow + 15)
                        With wsVarianceAnalysis.Range("Q" & (LastRow + 15))
                            .Formula = "=Q" & (LastRow + 7) & "+Q" & (LastRow + 13)
                        ' Format the cell to be 'Accounting'
                            .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        ' Make the font in the cell bold.
                            .Font.Bold = True
                        End With
'_____________________________________________________________________________________________________________________________________________

' Run the 'A/R Balance' portion of the sub
    BlackbaudARRecon_Part2


End Sub


Sub BlackbaudARRecon_Part2()

' All Variables
    Dim wsVarianceAnalysis As Worksheet
    
    Dim UserResponse2 As VbMsgBoxResult
    
    Dim fd2 As FileDialog
    Dim FolderPath2 As String
    
    Dim School As String
    Dim FY As String
    Dim SchoolSearch As String
    Dim VarianceLastRowA As Long
    
    Dim ARReportFound As Boolean
    
    Dim fso As Object
    Dim folder As Object
    Dim File As Object
    Dim wbTempARReport As Workbook
    Dim wsTempARReport As Worksheet
    Dim ARBalanceDate As Date
    Dim ARReportLastRow As Long
    Dim i As Long
    Dim DataRowValues As Variant
    Dim VariancePasteRow As Long
    Dim RenamedFileName As String
    
    Dim VarianceTableFirstRow As Long
    Dim o As Long
    Dim FirstRow As Long
    Dim VarianceTableLastRow As Long
    Dim c As Long
    Dim FirstBottomRow As Long
    
    Dim ws As Worksheet
    Dim SheetExists As Boolean
    Dim MoreGLsbtn As Button
    Dim MoreGLsbtnTop As Double
    Dim MoreGLsbtnLeft As Double
    Dim MoreGLsbtnHeight As Double
    Dim MoreGLsbtnWidth As Double
    Dim SplitOutbtn As Button
    Dim SplitOutbtnTop As Double
    Dim SplitOutbtnLeft As Double
    Dim SplitOutbtnHeight As Double
    Dim SplitOutbtnWidth As Double
    
    Dim ExtraMessage As String
    Dim btn As Button
    Dim btnTop As Double
    Dim btnLeft As Double
    Dim btnHeight As Double
    Dim btnWidth As Double


' Set the 'wsVarianceAnalysis' worksheet equal to Active Sheet (Variance Analysis Worksheet)
Set wsVarianceAnalysis = ThisWorkbook.ActiveSheet





' Ask user if they are sure they want to go to Part 2 of the converter
    UserResponse2 = MsgBox("Do you have the 'A/R Balance' Reports in a folder together?", vbYesNo + vbQuestion, "Confirmation of 'A/R Balance Reports'")


    ' Check the user's response - If they choose no, add a button and exit the sub.
        If UserResponse2 = vbNo Then
            GoTo NoARBalanceFiles
        End If

    ' If the user says "Yes",  Ask user to select the 'AR Balance' Reports Folder
        Set fd2 = Application.FileDialog(msoFileDialogFolderPicker)
            With fd2
                .Title = "Select 'A/R Balance' Reports Folder"
                If .Show <> -1 Then
                    ExtraMessage = "No Folder selected."
                    GoTo NoARBalanceFiles
                End If
                FolderPath2 = .SelectedItems(1)
            End With
            
     ' Check if there are files in the folder.
        If Dir(FolderPath2 & "\*.*") = "" Then
            ExtraMessage = "No files found in the selected folder. Please choose a folder with your 'A/R Balance' reports."
            GoTo NoARBalanceFiles
        End If
        
     ' Unhide the rows and delete the button
        ' Unhide rows 1:1000
            wsVarianceAnalysis.Rows("1:1000").EntireRow.Hidden = False
        ' Delete the rows with the button
            wsVarianceAnalysis.Range("1002:1035").EntireRow.Delete
            
    ' Check which school is in cell A1 of the Variance Analysis worksheet (wsVarianceAnalysis), store it in a variable called "School"
        School = wsVarianceAnalysis.Range("A1").Value
    
    ' Pass the value of 'School' through a loop to determine which school the reports should show. Store it in a variable called "SeachSchool".
        If School = "101 SAMC" Then
            SearchSchool = "BASIS San Antonio Primary - Medical Center Campus"
        ElseIf School = "102 SANC" Then
            SearchSchool = "BASIS San Antonio Primary - North Central Campus"
        ElseIf School = "103 SASH" Then
            SearchSchool = "BASIS San Antonio - Shavano Campus"
        ElseIf School = "104 SPNE" Then
            SearchSchool = "BASIS San Antonio Primary – Northeast Campus"
        ElseIf School = "105 AUSP" Then
            SearchSchool = "BASIS Austin Primary"
        ElseIf School = "106 SANE" Then
            SearchSchool = "BASIS San Antonio - Northeast Campus"
        ElseIf School = "107 AUS" Then
            SearchSchool = "BASIS Austin"
        ElseIf School = "108 PFLP" Then
            SearchSchool = "BASIS Pflugerville Primary"
        ElseIf School = "109 JLJP" Then
            SearchSchool = "BASIS San Antonio Primary Jack Lewis Jr. Campus"
        ElseIf School = "110 BEN" Then
            SearchSchool = "BASIS Benbrook"
        ElseIf School = "111 PFL" Then
            SearchSchool = "BASIS Pflugerville"
        ElseIf School = "112 JLJ" Then
            SearchSchool = "BASIS San Antonio Jack Lewis Jr. Campus"
        ElseIf School = "113 CPKP" Then
            SearchSchool = "BASIS Cedar Park Primary"
        ElseIf School = "114 CPK" Then
            SearchSchool = "BASIS Cedar Park"
        ElseIf School = "201 BDC" Then
            SearchSchool = "BASIS DC"
        ElseIf School = "401 AH" Then
            SearchSchool = "BASIS AHWATUKEE"
        ElseIf School = "402 CH" Then
            SearchSchool = "BASIS CHANDLER"
        ElseIf School = "403 CPS" Then
            SearchSchool = "BASIS CHANDLER PRIMARY - SOUTH CAMPUS"
        ElseIf School = "404 CPN" Then
            SearchSchool = "BASIS CHANDLER PRIMARY - NORTH CAMPUS"
        ElseIf School = "405 FL" Then
            SearchSchool = "BASIS FLAGSTAFF"
        ElseIf School = "406 GO" Then
            SearchSchool = "BASIS GOODYEAR"
        ElseIf School = "407 ME" Then
            SearchSchool = "BASIS MESA"
        ElseIf School = "408 OV" Then
            SearchSchool = "BASIS ORO VALLEY"
        ElseIf School = "409 PEO" Then
            SearchSchool = "BASIS PEORIA"
        ElseIf School = "410 PHX" Then
            SearchSchool = "BASIS PHOENIX"
        ElseIf School = "411 PC" Then
            SearchSchool = "BASIS PHOENIX CENTRAL"
        ElseIf School = "412 PR" Then
            SearchSchool = "BASIS PRESCOTT"
        ElseIf School = "413 SC" Then
            SearchSchool = "BASIS SCOTTSDALE"
        ElseIf School = "414 SPE" Then
            SearchSchool = "BASIS SCOTTSDALE PRIMARY EAST CAMPUS"
        ElseIf School = "415 TN" Then
            SearchSchool = "BASIS TUCSON NORTH"
        ElseIf School = "416 TP" Then
            SearchSchool = "BASIS TUCSON PRIMARY"
        ElseIf School = "417 GYP" Then
            SearchSchool = "BASIS GOODYEAR PRIMARY"
        ElseIf School = "418 OP" Then
            SearchSchool = "BASIS ORO VALLEY PRIMARY"
        ElseIf School = "419 PP" Then
            SearchSchool = "BASIS Peoria Primary"
        ElseIf School = "420 PS" Then
            SearchSchool = "BASIS Phoenix South"
        ElseIf School = "421 PXP" Then
            SearchSchool = "BASIS PHOENIX PRIMARY"
        ElseIf School = "422 SCW" Then
            SearchSchool = "BASIS SCOTTSDALE PRIMARY WEST CAMPUS"
        ElseIf School = "423 PHXN" Then
            SearchSchool = "BASIS Phoenix North"
        ElseIf School = "701 MA" Then
            SearchSchool = "BASIS Baton Rouge Materra Campus"
        ElseIf School = "702 MC" Then
            SearchSchool = "BASIS Baton Rouge Primary - Mid City Campus"
        End If

    ' Check with Fiscal Year is in the worksheet by using the last 2 characters of cell B4. Store it in a variable called "FY"
        FY = Left(wsVarianceAnalysis.Range("B4"), 2)
        
    ' Set 'ARReportFound' to False (if even 1 file is found) then set it to true in the loop. Otherwise the button hidden rows and button should be re-established.
        ARReportFound = False
        
    ' Find the 'VarianceLastRowA' variable using column A from the "Variance Analysis" worksheet to find the last row.
        VarianceLastRowA = wsVarianceAnalysis.Cells(wsVarianceAnalysis.Rows.Count, "A").End(xlUp).Row
        
    ' Loop through each file in the selected folder using FileSystemObject (This allows us to access file names and paths to process each Excel file individually)
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set folder = fso.GetFolder(FolderPath2)
    
     ' Loop through each file. If they have the correct data (the right FY, and school) then put them into the worksheet by the correct date. Rename the file. Skip over any files that don't meet the criteria.
        ' Open each file in FolderPath2
        For Each File In folder.Files
            ' Check to make sure the file name is an excel file (.xlsx extension)
            If LCase(Right(File.Name, 5)) = ".xlsx" Then
            ' If it is, open the file and set the workbook to a variable called 'wbTempARReport'
                Set wbTempARReport = Workbooks.Open(File.Path)
            ' Set the first worksheet to the be the worksheet we work in and set it to the variable 'wbTempARReport
                Set wsTempARReport = wbTempARReport.Sheets(1)
            ' Unmerge all cells
                wsTempARReport.Cells.UnMerge
            ' Check if School name in B1 matches SearchSchool
                If Trim(wsTempARReport.Range("B1").Value) <> SearchSchool Then
                ' If it does not match, close the file, without saving and jump to the next file
                    wbTempARReport.Close False
                    GoTo NextFile
                End If
            ' Check if FY matches value in A5 (starting at position 6 for 2 characters)
                If Mid(wsTempARReport.Range("A5").Value, 6, 2) <> FY Then
                ' If it does not match, close the file, without saving and jump to the next file
                    wbTempARReport.Close False
                    GoTo NextFile
                End If
            ' Get AR Balance Date from B3 - subtract 1, store it in the variable 'ARBalanceDate'
                ARBalanceDate = DateValue(wsTempARReport.Range("B3").Value) - 1
            ' Find last row in the file, using column I as the column lookup
                ARReportLastRow = wsTempARReport.Cells(wsTempARReport.Rows.Count, "I").End(xlUp).Row
            ' Get Columns D:I values from that last row
                DataRowValues = wsTempARReport.Range("D" & ARReportLastRow & ":I" & ARReportLastRow).Value
            ' Go back into the 'wsVarianceAnalysis' Worksheet and find the matching date in wsVarianceAnalysis, column A to 'ARBalanceDate'. Store the value in 'VariancePasteRow'
                For i = 1 To VarianceLastRowA
                    If IsDate(wsVarianceAnalysis.Cells(i, 1).Value) Then
                        If DateValue(wsVarianceAnalysis.Cells(i, 1).Value) = DateValue(ARBalanceDate) Then
                            VariancePasteRow = i
                            Exit For
                        End If
                    End If
                Next i
            ' Paste values into column C:H
                wsVarianceAnalysis.Range("C" & VariancePasteRow & ":H" & VariancePasteRow).Value = DataRowValues
            ' Update the 'ARReportFound' variable to True
                ARReportFound = True
            ' Save the file based on the the school (School), report date (ARBalanceDate + 1), fiscal year (FY + 1). Example: "401 AH - 2024.01.03 (FY24)"
                ' Create the folder if it doesn't already exist
                    If Dir(FolderPath2 & "\Renamed AR Reports", vbDirectory) = "" Then
                        MkDir FolderPath2 & "\Renamed AR Reports"
                    End If
                ' Create a variable 'RenamedFileName' to save the name of the file into the folder path given by the user
                    RenamedFileName = FolderPath2 & "\Renamed AR Reports\" & School & " - " & Format((ARBalanceDate + 1), "YYYY.MM.DD") & " (FY " & CStr(Val(FY) + 1) & ").xlsx"
                ' Check if the file name already exists.
                    If Dir(RenamedFileName) = "" Then
                    ' If it does not exist, save the file with the new name 'RenamedFileName'
                        wbTempARReport.SaveAs RenamedFileName
                    End If
            ' Close the file and proceed to the next.
                wbTempARReport.Close False
            End If
NextFile:
        ' Go to next file
        Next File
    
' If no files were placed into the 'wsVarianceAnalysis' Worksheet, then re-hide the rows and add a button for the user to go to when they are ready to try this process again.
    If ARReportFound = False Then
        ExtraMessage = "There was an issue with the files in the folder you selected." & vbCrLf & "*They may not be the correct reports." & vbCrLf & "*The school may not match the information in the GL report." & vbCrLf & "*Or the fiscal year may not match the information in the GL report." & vbCrLf
        GoTo NoARBalanceFiles
    End If



' Hide the "COMPLETE RESET" worksheet.
    ThisWorkbook.Worksheets("COMPLETE RESET").Visible = xlHidden

' Group the top unused months together and the bottom unused months together in the "Variance Analysis" worksheet.
    ' The top unused can go from C3:the first non-0 value in column C then subtract 2 rows (meaning if row 10 is the first to have any value, row 8 will be the last to be grouped.)

        For o = 3 To (VarianceLastRowA + 1)
            If ((wsVarianceAnalysis.Cells(o + 1, "C").Value <> 0) And (wsVarianceAnalysis.Cells(o + 1, "C").Value <> "")) Or ((wsVarianceAnalysis.Cells(o + 1, "O").Value <> 0) And (wsVarianceAnalysis.Cells(o + 1, "O").Value <> "")) Or ((wsVarianceAnalysis.Cells(o + 2, "P").Value <> 0) And (wsVarianceAnalysis.Cells(o + 2, "P").Value <> "")) Or ((wsVarianceAnalysis.Cells(o + 2, "Q").Value <> 0) And (wsVarianceAnalysis.Cells(o + 2, "Q").Value <> "")) Then
                FirstRow = o - 1
                Exit For
            End If
        Next o
        
        If (FirstRow > 3) Then
            With wsVarianceAnalysis.Rows(3 & ":" & FirstRow)
            ' Group Rows together
                .Rows.Group
            ' Collapse (Hide) the group
                .EntireRow.Hidden = True
            End With
        End If
        
    ' From last row in column C, go up to the first non-0 row. Then add 1 row and group that row and all the rows to the last row in column C.
        VarianceTableLastRow = VarianceLastRowA + 1
        
        For c = VarianceTableLastRow To 3 Step -1
            If ((wsVarianceAnalysis.Cells(c, "C").Value <> 0) And (wsVarianceAnalysis.Cells(c, "C").Value <> "")) Or ((wsVarianceAnalysis.Cells(c - 1, "O").Value <> 0) And (wsVarianceAnalysis.Cells(c - 1, "O").Value <> "")) Or ((wsVarianceAnalysis.Cells(c, "P").Value <> 0) And (wsVarianceAnalysis.Cells(c, "P").Value <> "")) Or ((wsVarianceAnalysis.Cells(c, "Q").Value <> 0) And (wsVarianceAnalysis.Cells(c, "Q").Value <> "")) Then
                FirstBottomRow = c + 1
                Exit For
            End If
        Next c
         
        If FirstBottomRow <= VarianceTableLastRow Then
            With wsVarianceAnalysis.Rows(FirstBottomRow & ":" & VarianceTableLastRow)
            ' Group Rows Together
                .Rows.Group
            ' Collapse (Hide) the group
                .EntireRow.Hidden = True 'This hides (collapses) the grouped rows
            End With
        End If

' Create a new worksheet to give the user an option to add in more GL report or split out all the worksheets from the macro.
        ' Check if "Add or Split Out Reports" worksheet already exists
            SheetExists = False
            For Each ws In ThisWorkbook.Worksheets
                If ws.Name = "Add or Split Out Reports" Then
                    SheetExists = True
                    Exit For
                End If
            Next ws
        ' If it doesn't exist, create it
            If Not SheetExists Then
                Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                ws.Name = "Add or Split Out Reports"
            ' The 'ADD MORE GL REPORTS' Button
                ' Set up button position and size (e.g., range B1:U13) for the 'ADD MORE GL REPORTS' Button
                    With ws.Range("B1:U13")
                        MoreGLsbtnTop = .Top
                        MoreGLsbtnLeft = .Left
                        MoreGLsbtnHeight = .Height
                        MoreGLsbtnWidth = .Width
                    End With
                ' Create a button for the "BlackbaudARRecon_Part1" Macro. ('ADD MORE GL REPORTS')
                    Set MoreGLsbtn = ws.Buttons.Add(Left:=MoreGLsbtnLeft, Top:=MoreGLsbtnTop, Width:=MoreGLsbtnWidth, Height:=MoreGLsbtnHeight)
                    With MoreGLsbtn
                        .Caption = "CLICK HERE TO ADD MORE GL REPORTS."
                        .OnAction = "BlackbaudARRecon_Part1"
                        .Name = "MoreGLsbtnRunGLImport"
                        .Font.Size = 55
                        .Font.Color = RGB(200, 200, 0)
                        .Font.Bold = True
                    End With
             ' The 'SPLIT OUT REPORTS' Button
                ' Set up button position and size (e.g., range B15:U25) for the 'SPLIT OUT REPORTS' Button
                   With ws.Range("B15:U25")
                       SplitOutbtnTop = .Top
                       SplitOutbtnLeft = .Left
                       SplitOutbtnHeight = .Height
                       SplitOutbtnWidth = .Width
                   End With
                ' Create a button for the "BlackbaudARRecon_SplitReports" Macro. ('SPLIT OUT REPORTS')
                    Set SplitOutbtn = ws.Buttons.Add(Left:=SplitOutbtnLeft, Top:=SplitOutbtnTop, Width:=SplitOutbtnWidth, Height:=SplitOutbtnHeight)
                    With SplitOutbtn
                        .Caption = "CLICK HERE TO SPLIT OUT REPORTS"
                        .OnAction = "BlackbaudARRecon_SplitReports"
                        .Name = "SplitOutReportsbtn"
                        .Font.Size = 55
                        .Font.Color = RGB(173, 216, 230)
                        .Font.Bold = True
                    End With
            End If

' Give the user a message letting them know the macro has completed.
    MsgBox "The macro has now completed. Thank you for your patience. "

Exit Sub


' If the user selects no they are not ready or cannot locate a folder path or no files are in the folder they selected or no files are 'A/R Balance' Report files or no files match the fiscal year.
NoARBalanceFiles:
' Hide the first 1000 Rows
    wsVarianceAnalysis.Rows("1:1000").EntireRow.Hidden = True
' Create a button for when the user is ready to pull the reports in. Put it in Columns: "C:S" and Rows: "1005:1030"
    ' To clear any other buttons (if the user has tried running this macro multiple times) - Delete Rows 1002:1035
        wsVarianceAnalysis.Range("1002:1035").EntireRow.Delete
    ' Get position and size from cell range C1002:S1030
        With wsVarianceAnalysis.Range("C1005:S1030")
            btnTop = .Top
            btnLeft = .Left
            btnHeight = .Height
            btnWidth = .Width
        End With
    ' Add button
        Set btn = wsVarianceAnalysis.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
        With btn
            .Caption = "Click to Load 'A/R Balance' Reports"
            .Name = "BlackbaudARRecon_Part2"
            .OnAction = "BlackbaudARRecon_Part2"
            .Font.Size = 72
            .Font.Bold = True
            .Font.Color = RGB(200, 0, 0)
        End With
' Give the user a message about the situation to let them know to press the button when they are ready.
    MsgBox ExtraMessage & vbCrLf & "Click the button when you have gathered the 'A/R Balance' Reports into one folder and are ready to pull them into this file.", Title:="Prepare to Load A/R Reports"
' Exit Macro
    Exit Sub
    
End Sub

Sub BlackbaudARRecon_SplitReports()

    Dim UserResponse As VbMsgBoxResult
    Dim ws As Worksheet
    Dim wbNew As Workbook
    Dim wbSource As Workbook
    Dim wsCount As Integer

    ' Ask user if they are sure they want to start the converter
        UserResponse = MsgBox("Are you sure you want to split out the reports from the macro?" & vbCrLf & vbCrLf & _
                            "Doing so, will result in the macro being saved with any changes made. The new workbook will be opened and you will need to manually save it.", _
                              vbYesNo + vbQuestion, "Confirmation to split reports out from the macro.")
        
    ' Check the user response. If they don't want to proceed, end the sub.
        If UserResponse = vbNo Then
            Exit Sub
        End If

    ' Create a variable for the macro workbook called 'wbSource'
        Set wbSource = ThisWorkbook

    ' Create a new workbook
        Set wbNew = Workbooks.Add(xlWBATWorksheet)

    ' Copy all worksheets except "Add or Split Out Reports" and "COMPLETE RESET"
        For Each ws In wbSource.Worksheets
            If ws.Name <> "Add or Split Out Reports" And ws.Name <> "COMPLETE RESET" Then
            ' If it's a Variance Analysis worksheet, clear formulas and paste values in O3:Q92 before copying
                If ws.Name Like "*Variance Analysis*" Then
                    With ws.Range("O3:Q92")
                    ' Paste values only (removes formulas)
                        .Value = .Value
                    ' In 'Accounting' Format
                        .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    End With
                End If
                ws.Copy After:=wbNew.Sheets(wbNew.Sheets.Count)
            End If
        Next ws

        
    ' Delete the original worksheet.
        ' Check that there is more than one worksheet. If there is, delete the one that is named "Sheet1".
            If wbNew.Worksheets.Count > 1 Then
            ' Prevent Excel from asking for confirmation
                Application.DisplayAlerts = False
            ' Delete "Sheet1"
                wbNew.Worksheets("Sheet1").Delete
                Application.DisplayAlerts = True
            End If

    ' Activate the new workbook
        wbNew.Activate
    
    ' Let the user know the macro has completed.
        MsgBox "Relevant worksheets have been copied to a new workbook. The original file has been saved and closed." & vbCrLf & vbCrLf & "Make sure to save this new file before closing it.", _
            vbInformation, "Split Complete"
            
    ' Save and close the source workbook (ThisWorkbook)
        wbSource.Save
        wbSource.Close SaveChanges:=True

End Sub


