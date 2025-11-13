Sub Click_and_Pledge_AR_Converter()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' The Purpose of the macro ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The purpose of this macro is to turn the 'Click and Pledge' reports into an 'Intacct' import file. _
    For background: _
        'Click and Pledge' is auto-synced to 'Salesforce'. For this reason, 'Salesforce' will show more details from the "line items" from 'Click and Pledge'. _
        Meaning, one 'Click and Pledge' line item (CNP Order Number) may have 1-50+ 'Salesforce' transactions within the single 'Click and Pledge' transaction. Within 'Salesforce' it shows the split. _
        This makes 'Salesforce' a crucial part of the process. However, 'Salesforce' only records Revenue. The 'Click and Pledge' reports are required to determine the fees and _
        the way the transactions align within a given bank deposit. 'Salesforce' plays a crucial role in helping us to determine the revenue account the transactions should hit, by showing us _
        what each individual transaction is for. _
            Update: The 'Salesforce' revenue data can now be synced to 'Intacct'. _
    The macro: _
        Due to the update, the macro now needs to determine the initial report being used ('Intacct' or 'Salesforce'). The 'Salesforce' revenue data now can sync over into 'Intacct', but all data _
        still needs to be reconciled, fees need to be added in, and the synced data needs to be reclassed out and into a single bank disbursement. _
        The macro will ask the user for a file and determine if it is the correct report format for the initial report ('Salesforce' or 'Intacct'). After that the macro will ask the user to provide _
        a folder path for all the 'Click and Pledge' files. The macro will check the folder says "Click and Pledge" (a way to be intentional about the files placed in a given folder). The macro will _
        go through each file to determine if it is a "Stripe" report or "ProPay" report from 'Click and Pledge'. If it matches either, the macro will manipulate the report to be universal, so the _
        columns will align, regardless of the type. The macro will then pull in the data and consolidate all the reports into one worksheet. While doing this, the macro will also pull in the _
        original report, for easily broken down reporting support. Once all the files have been consolidated, the macro will go back to the initial report and create a new worksheet to align _
        the data the same regardless of the initial report being 'Salesforce' or 'Intacct'. Then the macro will go into the consolidated data and split that in a way that makes the data universal _
        across all donation site platforms BASIS uses. (Making it easier to recycle code or consolidate the data across all platforms) From there, the macro splits out the fees, joins positive _
        transactions and negative transactions with the data from the initial report. The macro determines the accounts the transactions will hit within 'Intacct'. It brings everything back together _
        into an 'Intacct' import file(s). - "CRJ"s for positive bank disbursements and "Adjusting Journals" for negative bank disbursements (if the initial is 'Salesforce') and "Adjusting Journals" _
        every time the initial report is from 'Intacct'. The macro then double checks all the numbers, schools, and account determinations are correct. If any are missing or incorrect, it directs _
        the ones that need the user's attention.
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''' Update Log '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' Last Updated 2025.09.25 '''

''' Update Log:
            ''' 2025.03.14 '''
            ''' Production Date: 2024.04... '''
            ''' Create Time Frame: 2024.04... - 2024.04... '''
            ''
            ''' 2025.09.25 '''
            ''' REDESIGN Production Date: 2025.09.25 '''
            ''' COMPLETE REDESIGN time frame: 2025.09.01 - 2025.09.25 '''
            ''
            ''' 2025.11.06 '''
            ''' Sort the 'data validated' school list alphabetically '''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Create a variable to hold the name of the converter and call it 'ConverterName'. This is used for the button at the bottom.
    Dim ConverterName As String
    ConverterName = "Click_and_Pledge.Click_and_Pledge_AR_Converter"

    
' Turn off alerts and screen updating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False


' Create and assign "Converter" relevant variables to their respective values ('JournalType', 'ShowFormulas', 'Site')
    ' Create the variables
        Dim JournalType As String
        Dim ShowFormulas As Boolean
        Dim Site As String
        Dim ImportType As String ' "CRJ" or "Adjusting"
    
    ' Assign the variables
        JournalType = "CHAR"
        ShowFormulas = True
        Site = "Click And Pledge"


' Create and assign a variable to hold this workbook as a value. Call the variable 'wbMacro'.
    ' Create the variable
        Dim wbMacro As Workbook
    
    ' Assign the variable
        Set wbMacro = ThisWorkbook


' Create the variables relevant to the initial report direction. 'Salesforce' or 'Intacct'.
    ' Create the variables to check if the worksheets have already been added.
        Dim wsCheck As Worksheet
        Dim InitialExists As Boolean
        Dim wsInitialData As Worksheet
        Dim ReportRoute As String

    ' Create the variables for determining the direction, IF the worksheets have not already been added.
        Dim fdInitial As FileDialog
        Dim ExtraMessage As String
        Dim InitialReport As String
        Dim wbInitialTemp As Workbook
        
        Dim InitialReportType As String
        
        Dim InitialCnP() As Variant
        Dim InitialOtherDonations() As Variant
        Dim InitialInSchoolDeposits() As Variant
        Dim InitialIntacct() As Variant
        
        Dim wsInitialTemp As Worksheet
        Dim IHRC As Long ' "IHRC" = Initial Header Row Check
        Dim InitialHeaderRow As Long
        Dim ArrayCheckEnd As String
        Dim ITR_Data As Variant ' "ITR_Data" = Initial Temp Row_Data
        Dim ArrayCol As Long
        Dim ReportMatch As Boolean
        Dim ColumnHeaderRowInitial As Long



' Create the variables for after the initial report direction is determined:
    Dim InitialTempLastRow As Long

        

' Create the variables relevant to this specific converter's consolidated worksheet (for the donation site data).
    ' Create the variables
        Dim wsConsolidated As Worksheet

        Dim fdCnP As FileDialog
        Dim FolderPathCnP As String

        Dim FileCount As Long
        Dim FileName As String
        Dim FileNamesList() As String

        Dim ProPayHeaders() As Variant
        Dim StripeHeaders() As Variant
        
        Dim NonExcelFilesCount As Long
        Dim UsedProPayDailyFilesCount As Long
        Dim UsedProPayMonthlyFilesCount As Long
        Dim UsedStripeDailyFilesCount As Long
        Dim UsedStripeMonthlyFilesCount As Long
        Dim UnusedFilesCount As Long
        
        Dim FileNumber As Long
        
        Dim wbTemp As Workbook
        Dim wsTemp As Worksheet
        
        Dim TempLastRow As Long
        
        Dim RowFound As Boolean
        Dim CurrentRow As Long
        Dim ColumnMatchProPay As Long
        Dim ColumnMatchStripe As Long
        Dim Col As Long
        
        Dim HeaderRow As Long
        Dim ReportType As String
        
        Dim ReportNameRow As Long
        Dim ReportName As String
        
        Dim Underscores As Long
        
        Dim SchoolID As String
        Dim SchoolAbbrev As String
        
        Dim SchoolRow As Long
        
        Dim ReportTypeFull As String
        
        Dim ReportYear As String
        Dim ReportMonth As String
        Dim ReportDay As String
        Dim ReportDisbursement As String
        Dim FileNameStart As String
        
        Dim UsedFolderPathCnP As String
        
        Dim Dup As Long
        
        Dim UsedNonRenamedFolderPathCnP As String
        Dim RenamedFileName As String
        Dim RenamedFullPath As String
        Dim FinalizedFileName As String
        Dim FinalizedFullPath As String
        Dim UsedRenamedFolderPathCnP As String
        
        Dim CurrentFilePath As String
        Dim NewFilePath As String
        
        Dim DotPos As Long
        Dim BaseName As String
        Dim Ext As String
        Dim i As Long
        
        Dim ws As Worksheet
        Dim ConsolidatedLastRow As Long
        Dim DataStartRow As Long
        Dim ConsolidatedLastRowNow As Long
        
        Dim WorksheetName As String
        
        Dim wsNew As Worksheet
        Dim wsSummary As Worksheet
        Dim SummaryLastRow As Long
        
        Dim UnusedFolderPathCnP As String
        
        Dim NonClickandPledgeFilesFound As Long


' Create the variables relevant to the extraction of the data after the donation site data is extracted.
    ' Create the variables
        Dim wsUserChecks As Worksheet
        
        Dim wsStandardSF As Worksheet
        Dim CampaignBreakdown1 As String
        Dim CampaignBreakdown2 As String
        Dim CampaignBreakdown3 As String
        Dim StandardSFLastRow As Long
        
        Dim InitialLastRow As Long
        
        Dim wsStandardDonations As Worksheet
        Dim StandardDonationsLastRow As Long
        
        Dim wsDisbursements As Worksheet
        Dim DisbursementsLastRow As Long
        
        Dim wsPosTransactions As Worksheet
        Dim PosTransactionsLastRow As Long
        
        Dim wsNegTransactions As Worksheet
        Dim NegTransactionsLastRow As Long
        
        Dim wsAllPossibleFees As Worksheet
        Dim AllPossibleFeesLastRow As Long
        
        Dim wsFeesFiltered As Worksheet
        Dim FeesFilteredLastRow As Long
        
        Dim wsBankDisbursementAmounts As Worksheet
        Dim BankDisbursementAmountsLastRow As Long
        
        
' Create the variables to help create the Intacct Journal Import files
    ' CRJ Route
        Dim wsAllDataCombinedPos As Worksheet
        Dim AllDataCombinedPosLastRow As Long
        Dim wsAllDataCombinedNeg As Worksheet
        Dim AllDataCombinedNegLastRow As Long
        Dim wsImportCRJ As Worksheet
        Dim wsImportAdjusting As Worksheet
    
    ' Adjusting Journal Route
        Dim wsAllDataCombined As Worksheet
        Dim AllDataCombinedLastRow As Long
        Dim wsImport As Worksheet
        
        
' Create the variables to help with the user-required checks
    ' User Checks
        Dim UserChecksLastRow As Long
        Dim wsSchoolList As Worksheet
        Dim SchoolNames As Variant
        Dim DataValidationRange_School As Range
        Dim UserChecksNewCheckRow As Long
        Dim DVRow As Long ' "DVRow" means Data Validation Row
        Dim VarianceCount_Gross As Long

' Create the Button-related variables:
    Dim wsButton As Worksheet
    Dim DonationSiteButton As Button
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''--------------------------------------------------------------------------------''''''''''''''''''''
'''''''''''''''''''' Check if the user has already given the Initial Report (Salesforce or Intacct) ''''''''''''''''''''
''''''''''''''''''''--------------------------------------------------------------------------------''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Update the Status Bar
        Application.StatusBar = "Checking For Existing Worksheets"
        
    ' Check if the Salesforce or Intacct Report has already been added into the macro file.
        ' Loop through each worksheet to check if the the Intacct or Salesforce worksheet is present. If it is store "True" in the variable 'InitialExists' _
          Set the 'ReportRoute' and set the 'wsInitialData' to the correct worksheet.
            For Each wsCheck In wbMacro.Worksheets
                wsCheck.Visible = xlSheetVisible
                If wsCheck.Name = "Initial Data - Intacct" Then
                    InitialExists = True
                    ReportRoute = "Intacct"
                    Set wsInitialData = wsCheck
                    Exit For
                    
                ElseIf wsCheck.Name = "Initial Data - SF" Then
                    InitialExists = True
                    ReportRoute = "Salesforce"
                    Set wsInitialData = wsCheck
                    Exit For
                    
                End If
            Next wsCheck
        
        ' Check if 'InitialExists' is True. If it is, jump to consolidating the donation site reports.
            If InitialExists Then
                GoTo Add_ConsolidatedReports
            End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''----------------------------------------------------------------------------------------''''''''''''''''''''
'''''''''''''''''''' Ask user for the initial report and check which path to choose (Salesforce or Intacct) ''''''''''''''''''''
''''''''''''''''''''----------------------------------------------------------------------------------------''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Update the status bar
        Application.StatusBar = "User Must Select the Initial Report (Salesforce or Intacct)"
        
    ' Create the 'Reset' worksheet by running the "Reset.Create_Reset_Worksheet" macro.
        Reset.Create_Reset_Worksheet
    
    ' Ask the user for the initial report
        ' Create a file picker, file dialog message, and store it in a variable called 'fdInitial'.
            Set fdInitial = Application.FileDialog(msoFileDialogFilePicker)
        
        ' Open the file picker, file dialog message. If the user cancels, provide a message for them. If they select a file: _
          Open it and determine the direction the rest of the macro will follow.
            With fdInitial
                .Title = "Select Initial Report (Salesforce or Intacct)"
                .AllowMultiSelect = False
                If .Show <> -1 Then
                    ExtraMessage = "No file selected. Please locate the Initial Report (Salesforce or Intacct) and try again."
                    MsgBox ExtraMessage, vbCritical, "No File Selected"
                    GoTo ResetTheWorkbook
                End If
                
                InitialReport = .SelectedItems(1)
            End With
            
    ' Update the status bar
        Application.StatusBar = "Checking the User Selected File"
            
            ' Open the 'InitialReport' file.
                ' Turn off Calculations
                    Application.Calculation = xlCalculationManual
                    
                ' Make sure the file is an excel type file.
                    If LCase(Right(InitialReport, 5)) = ".xlsx" Or LCase(Right(InitialReport, 5)) = ".xlsm" Or LCase(Right(InitialReport, 4)) = ".xls" Then
                        Set wbInitialTemp = Workbooks.Open(FileName:=InitialReport, ReadOnly:=True, UpdateLinks:=0, Notify:=False)
                        InitialReportType = "Excel"
                        
                    ElseIf LCase(Right(InitialReport, 4)) = ".csv" Then
                        Set wbInitialTemp = Workbooks.Open(InitialReport)
                        InitialReportType = "CSV"
                        
                    Else
                        ExtraMessage = "The selected file, does not appear to be an excel file. Please locate the correct file and try again."
                        MsgBox ExtraMessage, vbCritical, "Unsupported File Type"
                        GoTo ResetTheWorkbook
                        
                    End If
                
                ' Turn back on Calculations
                    Application.Calculation = xlCalculationAutomatic
            
            ' Check if the file is the correct Salesforce Report or the Intacct Report.
                ' Salesforce Report Types
                    ' Click and Pledge (A:AE)
                        InitialCnP = Array("School ID", "C&P Account Number Manual", "Opportunity Record Type", "Application Name", "Primary Contact", "Account Name", "Opportunity Name", _
                                "Primary Campaign Source", "Primary Campaign Source Name", "Description", "Payment from", "Payment: Payment Number", "C&P Order Number", "C&P Payment ReferenceID/GUID", _
                                "Check Number", "Payment Date", "Payment: Created Date", "Payment Amount", "Payment Amount Received", "Stage", "Payment Type", "Payment: ID", "Close Date", _
                                "Deposit Date (Manual)", "Deposit Date", "Last Name", "First Name", "Campaign Type", "Payment: Created By", "Payment: Last Modified By", "C&P Account Name")
                                
                    ' Other Donations (A:P)
                        InitialOtherDonations = Array("School Name", "Created Date", "Close Date", "Deposit Date", "First Name", "Last Name", "Payment Amount", "Payment: Payment Number", _
                                "C&P Payment Type", "Check Number", "Check/Reference Number", "Disbursement ID", "Description", "Company Name", "Opportunity Name", "Primary Campaign Source")
                    
'                    ' Other Donations (A:P)
'                        InitialOtherDonations = Array("School Name", "Created Date", "Close Date", "Deposit Date", "First Name", "Last Name", "Payment Amount", "Payment: Payment Number", _
'                                "C&P Payment Type", "Check Number", "Check/Reference Number", "Disbursement ID", "Company Name", "Opportunity Name", "Primary Campaign Source", "Description")
'
                    ' In-School Deposits (A:Y)
                        InitialInSchoolDeposits = Array("Opportunity Record Type", "C&P Account Name", "C&P Account Number", "Primary Contact", "Account Name", "Transaction Type", _
                                "Opportunity Name", "Primary Campaign Source", "Primary Campaign Source Name", "Description", "Payment from", "C&P Order Number", "Payment: ID", "Stage", _
                                "Payment Type", "Check Number", "Date of check", "Payment: Payment Number", "Payment Amount", "Payment: Created Date", "Payment Date", "Deposit Date (Manual)", _
                                "Deposit Date", "Payment: Last Modified By", "Last Modified By")
                    
                ' Intacct Report Type (A:AH)
                    InitialIntacct = Array("Batch posting date", "Journal entry modified date", "Close Date", "Account no.", "Account title", "Location ID", "Location name", "Memo", "Campaign Source", _
                            "SF Payment Number", "SF Donation Site", "SF Company Name", "C&P Number", "SF Transaction ID", "SF Disbursement ID", "SF Check Number", "SF Payment Method", "Donation Name", _
                            "SF Account Name", "SF Primary Contact", "Record number", "Journal", "Transaction no.", "Batch description", "Division ID", "Division name", "Funding Source", "Customer ID", _
                            "Customer name", "Debt Service Series ID", "Name", "Credit amount", "Debit amount", "Amount")
                            
                ' Loop through the file to see if the column headers fit any of the above Salesforce or Intacct reports. Starting in Row 20 and working our way up. _
                  Store the column header row in a variable called 'ColumnHeaderRowInitial'.
                    ' First set the first worksheet in the file to the variable 'wsInitialTemp'
                        Set wsInitialTemp = wbInitialTemp.Worksheets(1)
                    
                    ' Set the 'ReportRoute' to "Not Found" until it is found.
                        ReportRoute = "Not Found"
                        
                    ' Loop through the rows in the 'wsInitialTemp' worksheet, to find which type of report the user is trying to pass-through. Start in row 20 and go down 1 row at a time.
                        For IHRC = 20 To 1 Step -1 ' "IHRC" = Initial Header Row Check
                        
                        ' ----------Try Click & Pledge (Salesforce) Report ----------
                            ' Create a variable to hold the last column of the 'InitialCnP' array.
                                ArrayCheckEnd = "AE" ' (InitialCnP-Array = Columns A:AE)
                                
                            ' Create a variable called 'ITR_Data' to hold the data for the row to compare the 'InitialCnP' array against.
                                ITR_Data = wsInitialTemp.Range("A" & IHRC & ":" & ArrayCheckEnd & IHRC).Value ' "ITR_Data" = Initial Temp Row_Data
                                
                            ' Use a variable called 'ReportMatch' to determine if the 'InitialCnP' array and the 'ITR_Data' match.
                                ReportMatch = True
                            
                            ' Create a loop to compare between the 'ITR_Data' array and the 'InitialCnP' array. If they do not match, move on.
                                For ArrayCol = 1 To 31 ' (Columns A:AE)
                                    If StrComp(Trim$(CStr(ITR_Data(1, ArrayCol))), Trim$(CStr(InitialCnP(ArrayCol - 1))), vbTextCompare) <> 0 Then
                                        ReportMatch = False
                                        Exit For
                                    End If
                                Next ArrayCol
                                
                            ' If they do match, store the header row in a variable called 'ColumnHeaderRowInitial'. Use "Salesforce" as the 'ReportRoute'.
                                If ReportMatch Then
                                    ColumnHeaderRowInitial = IHRC
                                    ReportRoute = "Salesforce"
                                    Exit For
                                End If
                                
                        '----------Try Other Donations (Salesforce) Report ----------
                            ' Create a variable to hold the last column of the 'InitialOtherDonations' array.
                                ArrayCheckEnd = "P" ' (InitialOtherDonations-Array = Columns A:P)
                                
                            ' Create a variable called 'ITR_Data' to hold the data for the row to compare the 'InitialOtherDonations' array against.
                                ITR_Data = wsInitialTemp.Range("A" & IHRC & ":" & ArrayCheckEnd & IHRC).Value ' "ITR_Data" = Initial Temp Row_Data
                                
                            ' Use a variable called 'ReportMatch' to determine if the 'InitialOtherDonations' array and the 'ITR_Data' match.
                                ReportMatch = True
                            
                            ' Create a loop to compare between the 'ITR_Data' array and the 'InitialCnP' array. If they do not match, move on.
                                For ArrayCol = 1 To 16 ' (Columns A:P)
                                    If StrComp(Trim$(CStr(ITR_Data(1, ArrayCol))), Trim$(CStr(InitialOtherDonations(ArrayCol - 1))), vbTextCompare) <> 0 Then
                                        ReportMatch = False
                                        Exit For
                                    End If
                                Next ArrayCol
                                
                            ' If they do match, let the user know this is the incorrect report.
                                If ReportMatch Then
                                    wbInitialTemp.Close SaveChanges:=False
                                    ExtraMessage = "The selected file is the incorrect report. It appears to be the 'Other Donations' Salesforce Report. Please locate the '" & Site & _
                                            "' Report and try again."
                                    MsgBox ExtraMessage, vbCritical, "Incorrect Report"
                                    GoTo ResetTheWorkbook
                                End If
                            
                            
                        '----------Try In-School Deposits (Salesforce) Report ----------
                            ' Create a variable to hold the last column of the 'InitialInSchoolDeposits' array.
                                ArrayCheckEnd = "Y" ' (InitialInSchoolDeposits-Array = Columns A:Y)
                                
                            ' Create a variable called 'ITR_Data' to hold the data for the row to compare the 'InitialInSchoolDeposits' array against.
                                ITR_Data = wsInitialTemp.Range("A" & IHRC & ":" & ArrayCheckEnd & IHRC).Value ' "ITR_Data" = Initial Temp Row_Data
                                
                            ' Use a variable called 'ReportMatch' to determine if the 'InitialInSchoolDeposits' array and the 'ITR_Data' match.
                                ReportMatch = True
                            
                            ' Create a loop to compare between the 'ITR_Data' array and the 'InitialCnP' array. If they do not match, move on.
                                For ArrayCol = 1 To 25 ' (Columns A:Y)
                                    If StrComp(Trim$(CStr(ITR_Data(1, ArrayCol))), Trim$(CStr(InitialInSchoolDeposits(ArrayCol - 1))), vbTextCompare) <> 0 Then
                                        ReportMatch = False
                                        Exit For
                                    End If
                                Next ArrayCol
                                
                            ' If they do match, let the user know this is the incorrect report.
                                If ReportMatch Then
                                    wbInitialTemp.Close SaveChanges:=False
                                    ExtraMessage = "The selected file is the incorrect report. It appears to be the 'In-School Deposit' Salesforce Report. Please locate the '" & Site & _
                                            "' Report and try again."
                                    MsgBox ExtraMessage, vbCritical, "Incorrect Report"
                                    GoTo ResetTheWorkbook
                                End If
                            
                            
                        '----------Try Intacct Report ----------
                            ' Create a variable to hold the last column of the 'InitialIntacct' array.
                                ArrayCheckEnd = "AH" ' (InitialIntacct-Array = Columns A:AH)
                                
                            ' Create a variable called 'ITR_Data' to hold the data for the row to compare the 'InitialIntacct' array against.
                                ITR_Data = wsInitialTemp.Range("A" & IHRC & ":" & ArrayCheckEnd & IHRC).Value ' "ITR_Data" = Initial Temp Row_Data
                                
                            ' Use a variable called 'ReportMatch' to determine if the 'InitialIntacct' array and the 'ITR_Data' match.
                                ReportMatch = True
                            
                            ' Create a loop to compare between the 'ITR_Data' array and the 'InitialIntacct' array. If they do not match, move on.
                                For ArrayCol = 1 To 34 ' (Columns A:AH)
                                    If StrComp(Trim$(CStr(ITR_Data(1, ArrayCol))), Trim$(CStr(InitialIntacct(ArrayCol - 1))), vbTextCompare) <> 0 Then
                                        ReportMatch = False
                                        Exit For
                                    End If
                                Next ArrayCol
                                
                            ' If they do match, store the header row in a variable called 'ColumnHeaderRowInitial'. Use "Intacct" as the 'ReportRoute'.
                                If ReportMatch Then
                                    ColumnHeaderRowInitial = IHRC
                                    ReportRoute = "Intacct"
                                    Exit For
                                End If
                    ' Move up to the next row.
                        Next IHRC
    
    ' Determine the direction to send the macro.
        If ReportRoute = "Salesforce" Then
            GoTo Add_Salesforce
            
        ElseIf ReportRoute = "Intacct" Then
            GoTo Add_Intacct
            
        Else
            wbInitialTemp.Close SaveChanges:=False
            ExtraMessage = "The selected file, was not a recognized report. Please find the correct file and try again."
            MsgBox ExtraMessage, vbCritical, "Report Not Recognized"
            GoTo ResetTheWorkbook
            
        End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''------------------------------''''''''''''''''''''
'''''''''''''''''''' Add in the Salesforce Report ''''''''''''''''''''
''''''''''''''''''''------------------------------''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Add_Salesforce:
' Update the status bar
    Application.StatusBar = "Adding the Initial 'Salesforce' Report"

' Find the last row of the 'wsInitialTemp' worksheet and store it a variable called 'InitialTempLastRow'. Use the 'PMT-ID (Payment: Payment Number)' Column. (Column L)
    InitialTempLastRow = wsInitialTemp.Cells(wsInitialTemp.Rows.Count, "L").End(xlUp).Row
    
' Check if the last row is the same as the 'ColumnHeaderRowInitial' row.
    If InitialTempLastRow = ColumnHeaderRowInitial Then
        wbInitialTemp.Close SaveChanges:=False
        ExtraMessage = "The selected file is the correct 'Salesforce' Report. However, it has no usable data. Please locate the correct report and try again."
        MsgBox ExtraMessage, vbCritical, "No Data Found"
        GoTo ResetTheWorkbook
    End If

' Create a worksheet called "Initial Data - SF" and store it in a variable called: 'wsInitialData'
    Set wsInitialData = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets("COMPLETE RESET"))
    
    ' Rename the worksheet to "Initial Data - SF"
        wsInitialData.Name = "Initial Data - SF"

' Pull in the data from the 'wsInitialTemp' worksheet.
    wsInitialTemp.Range("A" & ColumnHeaderRowInitial & ":" & ArrayCheckEnd & InitialTempLastRow).Copy Destination:=wsInitialData.Range("A1")
    
    ' Clear 'CutCopy' Mode
        Application.CutCopyMode = False

' Close the 'wbInitialTemp' workbook without saving.
    wbInitialTemp.Close SaveChanges:=False

' Format the 'wsInitialData' worksheet
    ' Unwrap, AutoFilter, and AutoFit the columns
        wsInitialData.Cells.WrapText = False
        wsInitialData.Range("A1:" & ArrayCheckEnd & "1").AutoFilter
        wsInitialData.Columns("A:" & ArrayCheckEnd).AutoFit

GoTo Add_ConsolidatedReports


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''---------------------------''''''''''''''''''''
'''''''''''''''''''' Add in the Intacct Report ''''''''''''''''''''
''''''''''''''''''''---------------------------''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Add_Intacct:
' Update the status bar
    Application.StatusBar = "Adding the Initial 'Intacct' Report"

' Find the last row of the 'wsInitialTemp' worksheet and store it a variable called 'InitialTempLastRow'. Use the 'Journal entry modified date' Column. (Column B)
    InitialTempLastRow = wsInitialTemp.Cells(wsInitialTemp.Rows.Count, "B").End(xlUp).Row
    
' Check if the last row is the same as the 'ColumnHeaderRowInitial' row.
    If InitialTempLastRow = ColumnHeaderRowInitial Then
        wbInitialTemp.Close SaveChanges:=False
        ExtraMessage = "The selected file is the correct 'Intacct' Report. However, it has no usable data. Please locate the correct report and try again."
        MsgBox ExtraMessage, vbCritical, "No Data Found"
        GoTo ResetTheWorkbook
    End If

' Create a worksheet called "Initial Data - Intacct" and store it in a variable called: 'wsInitialData'
    Set wsInitialData = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets("COMPLETE RESET"))
    
    ' Rename the worksheet to "Initial Data - Intacct"
        wsInitialData.Name = "Initial Data - Intacct"

' Pull in the data from the 'wsInitialTemp' worksheet.
    wsInitialTemp.Range("A" & ColumnHeaderRowInitial & ":" & ArrayCheckEnd & InitialTempLastRow).Copy Destination:=wsInitialData.Range("A1")
    
    ' Clear 'CutCopy' Mode
        Application.CutCopyMode = False

' Close the 'wbInitialTemp' workbook without saving.
    wbInitialTemp.Close SaveChanges:=False
    
' Format the 'wsInitialData' worksheet
    ' Unwrap, AutoFilter, and AutoFit the columns
        wsInitialData.Cells.WrapText = False
        wsInitialData.Range("A1:" & ArrayCheckEnd & "1").AutoFilter
        wsInitialData.Columns("A:" & ArrayCheckEnd).AutoFit

GoTo Add_ConsolidatedReports


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''-------------------------------------------------''''''''''''''''''''
'''''''''''''''''''' Consolidate all of the Click and Pledge Reports ''''''''''''''''''''
''''''''''''''''''''-------------------------------------------------''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Add_ConsolidatedReports:
' Turn off Screen Updating again.
    Application.ScreenUpdating = False

' Ask user for a folder path
    Set fdCnP = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fdCnP
        .Title = "Select the '" & Site & "' Reports Folder"
        .AllowMultiSelect = False
        
        If .Show <> -1 Then
            ExtraMessage = "No Folder Selected. Please locate the correct folder and try again."
            GoTo NoFiles
        End If
        
        FolderPathCnP = .SelectedItems(1)
    End With
    
    ' Make sure the 'FolderPathCnP' has the name "Click and Pledge" in it.
        If InStr(1, FolderPathCnP, Site, vbTextCompare) = 0 Then
            ExtraMessage = "The folder path selected, does not contain '" & Site & "' within the folder name. Please rename the folder or " & _
                             "find the correct folder path and try again."
            GoTo NoFiles
        End If
        
    ' Make sure there is at least one file in the folder path.
        FileCount = 0
        
        FileName = Dir(FolderPathCnP & "\*.*", vbNormal Or vbReadOnly Or vbHidden Or vbSystem)
        
        Do While Len(FileName) > 0
            FileCount = FileCount + 1
        ' Create a list for all the files in the 'FolderPathCnP'
            ReDim Preserve FileNamesList(1 To FileCount)
            FileNamesList(FileCount) = FileName
            FileName = Dir()
        Loop
        
        If FileCount = 0 Then
            ExtraMessage = "No files were found in the selected folder. Please locate the correct folder and try again."
            GoTo NoFiles
        End If
        

' Assign values to all necessary variables before jumping into a loop.
    ' Assign the column headers 'ProPayHeaders' and 'StripeHeaders' to the way they will appear in the report from Click and Pledge
        ProPayHeaders = Array("SweepId", "Order Number", "C&P Transaction date", "Fund Date", "First Name", "Last Name", "Gross Amount", "Per Transaction Fees", "Discount Fees", "Net Amount")
        StripeHeaders = Array("Automatic_Payout_Id", "Order Number", "C&P Transaction date", "Fund Date", "First Name", "Last Name", "Gross Amount", "Fee", "Net Amount")
        
    
    ' Create the variables what type and how many of each file type was used in this macro.
        NonExcelFilesCount = 0
        UsedProPayDailyFilesCount = 0
        UsedProPayMonthlyFilesCount = 0
        UsedStripeDailyFilesCount = 0
        UsedStripeMonthlyFilesCount = 0
        UnusedFilesCount = 0
        
    ' Create a new worksheet called "Consolidated Reports" and store it in a variable called 'wsConsolidated'
        Set wsConsolidated = wbMacro.Worksheets.Add(After:=wsInitialData)
        
        ' Rename it
            wsConsolidated.Name = "Consolidated Reports"
            
        ' Add the column headers
            wsConsolidated.Range("A1:V1").Value = Array("SweepId", "Order Number", "C&P Transaction date", "Fund Date", "First Name", _
                "Last Name", "Gross Amount", "Per Transaction Fees", "Discount Fees", "Fee", "Total Fees", "Net Amount", "Original File Name", _
                "Renamed File Name", "Worksheet Name", "Report Type", "Order Number Cleaned", "School ID", _
                "Disbursement ID (School_ID - Sweep ID or Payout ID)", "Transaction Date", "Fund Date", "Name - Last, First")
                

' Start looping through each file, one by one.
    For FileNumber = LBound(FileNamesList) To UBound(FileNamesList)
        ' Update the status bar with the current file number and total file count
            Application.StatusBar = "Processing file " & (FileNumber - LBound(FileNamesList) + 1) & " of " & (UBound(FileNamesList) - LBound(FileNamesList) + 1) & ": " & _
                    FileNamesList(FileNumber)
    
        ' Check to make sure it is an excel supported file
            If Not LCase$(FileNamesList(FileNumber)) Like "*.csv" And Not LCase$(FileNamesList(FileNumber)) Like "*.xls*" Then
                NonExcelFilesCount = NonExcelFilesCount + 1
                GoTo DoNotUseFile
            End If
        
        ' If it is an excel supported file, open the file and assign it to the variable 'wbTemp'
            Set wbTemp = Workbooks.Open(FolderPathCnP & "\" & FileNamesList(FileNumber), ReadOnly:=True)
        
        ' Assign the first sheet to a variable called 'wsTemp'
            Set wsTemp = wbTemp.Worksheets(1)

        ' Find the last row of data in the 'wsTemp' worksheet
            TempLastRow = wsTemp.Cells(wsTemp.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through each row to find the matching column headers. If no row is found, move to the next file.
            RowFound = False
            For CurrentRow = 1 To TempLastRow
            '''''--- ProPay Reports ---'''''
                If StrComp(CStr(wsTemp.Cells(CurrentRow, "A").Value), ProPayHeaders(0), vbTextCompare) = 0 Then
                    
                    ColumnMatchProPay = 0
                    
                    For Col = 0 To 9
                        If StrComp(CStr(wsTemp.Cells(CurrentRow, Col + 1).Value), ProPayHeaders(Col), vbTextCompare) = 0 Then
                            ColumnMatchProPay = ColumnMatchProPay + 1
                        Else
                            Exit For
                        End If
                    Next Col
                    
                    If ColumnMatchProPay = 10 Then
                        HeaderRow = CurrentRow
                        ReportType = "ProPay"
                        GoTo UseFile
                    End If
                    
            '''''--- Stripe Reports ---'''''
                ElseIf StrComp(CStr(wsTemp.Cells(CurrentRow, "A").Value), StripeHeaders(0), vbTextCompare) = 0 Then
                    ColumnMatchStripe = 0
                    
                    For Col = 0 To 8
                        If StrComp(CStr(wsTemp.Cells(CurrentRow, Col + 1).Value), StripeHeaders(Col), vbTextCompare) = 0 Then
                            ColumnMatchStripe = ColumnMatchStripe + 1
                        Else
                            Exit For
                        End If
                    Next Col
                    
                    If ColumnMatchStripe = 9 Then
                        HeaderRow = CurrentRow
                        ReportType = "Stripe"
                        GoTo UseFile
                    End If
                
                End If
            Next CurrentRow
            
        ' If the columns do not match a Click and Pledge Report, do not use the file. Go to the next file.
            UnusedFilesCount = UnusedFilesCount + 1
            GoTo DoNotUseFile

' For usable files:
UseFile:
    ' Find the row where column B = "Report Name" exists and read the value from column C
        For ReportNameRow = 1 To TempLastRow
            If StrComp(Trim$(CStr(wsTemp.Cells(ReportNameRow, "B").Value)), "Report Name", vbTextCompare) = 1 Then
                ReportName = Trim$(CStr(wsTemp.Cells(ReportNameRow, "C").Value))
                Exit For
            End If
        Next ReportNameRow
    
    ' If Not found
        If Len(ReportName) = 0 Then
            UnusedFilesCount = UnusedFilesCount + 1
            GoTo DoNotUseFile
        End If
    
    ' Count underscores in ReportName
        Underscores = Len(ReportName) - Len(Replace(ReportName, "_", ""))
        
    ' SchoolRow
        SchoolID = ""
        SchoolAbbrev = ""
    
    ' Find the 'SchoolID'
        For SchoolRow = (HeaderRow + 1) To TempLastRow
            SchoolID = Mid(Range("B" & SchoolRow).Value, 5, 5)
            If SchoolID <> "Month" Then
                Exit For
            End If
        Next SchoolRow
        
        If SchoolID = "Month" Then
            SchoolAbbrev = "NONE"
        Else
            SchoolAbbrev = ConvertCnPToSchoolAbbrev(SchoolID)
        End If

    ' Determine the exact report type by using the underscore count
        ' ProPay Reports
        If StrComp(ReportType, "ProPay", vbTextCompare) = 0 Then
            Select Case Underscores
            ' If ReportName has 2 underscores => ProPay - Daily
                Case 2:
                    ReportTypeFull = "ProPay - Daily"
                    UsedProPayDailyFilesCount = UsedProPayDailyFilesCount + 1
                    ReportYear = Mid(ReportName, 18, 4)
                    ReportMonth = Mid(ReportName, 22, 2)
                    ReportDay = Mid(ReportName, 24, 2)
                    ReportDisbursement = SchoolID & " - " & Mid(ReportName, 27, Len(ReportName) - 27)
                    FileNameStart = "Click and Pledge_" & ReportYear & "." & ReportMonth & "." & ReportDay & " (" & ReportDisbursement & ")"
                    
            ' If ReportName has 1 underscore  => ProPay - Monthly
                Case 1:
                    ReportTypeFull = "ProPay - Monthly"
                    UsedProPayMonthlyFilesCount = UsedProPayMonthlyFilesCount + 1
                    ReportYear = Mid(ReportName, 18, 4)
                    ReportMonth = Right(ReportName, 2)
                    FileNameStart = "Click and Pledge_" & ReportYear & "." & ReportMonth
                    
                Case Else:
                    UnusedFilesCount = UnusedFilesCount + 1
                    GoTo DoNotUseFile
            End Select
        
        ' Stripe Reports
        ElseIf StrComp(ReportType, "Stripe", vbTextCompare) = 0 Then
            Select Case Underscores
        ' If ReportName has 3 underscores => Stripe - Daily
                Case 3:
                    ReportTypeFull = "Stripe - Daily"
                    UsedStripeDailyFilesCount = UsedStripeDailyFilesCount + 1
                    ReportYear = Mid(ReportName, 14, 4)
                    ReportMonth = Mid(ReportName, 18, 2)
                    ReportDay = Mid(ReportName, 20, 2)
                    ReportDisbursement = Mid(ReportName, 23, Len(ReportName) - 23)
                    FileNameStart = "Click and Pledge_" & ReportYear & "." & ReportMonth & "." & ReportDay & " (" & ReportDisbursement & ")"
                    
        ' If ReportName has 1 underscore => Stripe - Monthly
                Case 1:
                    ReportTypeFull = "Stripe - Monthly"
                    UsedStripeMonthlyFilesCount = UsedStripeMonthlyFilesCount + 1
                    ReportYear = Mid(ReportName, 14, 4)
                    ReportMonth = Mid(ReportName, 19, 2)
                    FileNameStart = "Click and Pledge_" & ReportYear & "." & ReportMonth
                    
                Case Else:
                    UnusedFilesCount = UnusedFilesCount + 1
                    GoTo DoNotUseFile
            End Select
        End If

' Check if there is a folder path made yet
    ' Check if there is a folder called "UsedFiles" in the original folder path, store the name in a variable called 'UsedFolderPathCnP'
        UsedFolderPathCnP = FolderPathCnP & "\Used Files"
        
        ' If it is not created yet, create one
            If Len(Dir(UsedFolderPathCnP, vbDirectory)) = 0 Then
                MkDir UsedFolderPathCnP
            End If

Dup = 0
' Save a renamed file name
    If SchoolAbbrev = "NONE" Then
    ' Check if there is a folder called "Files with no school found" in the 'UsedFolderPathCnP'.
        UsedNonRenamedFolderPathCnP = UsedFolderPathCnP & "\Files with no school found"
        
        If Len(Dir(UsedNonRenamedFolderPathCnP, vbDirectory)) = 0 Then
            MkDir UsedNonRenamedFolderPathCnP
        End If
        
        RenamedFileName = FileNameStart & " (School Not Found) - " & ReportType
        RenamedFullPath = UsedNonRenamedFolderPathCnP & "\" & RenamedFileName
        
        If Len(Dir(RenamedFullPath & ".csv", vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) <> 0 Then
            Dup = 1
            
            Do While Len(Dir(RenamedFullPath & " (" & Dup & ").csv", vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) <> 0
                Dup = Dup + 1
            Loop
            
            FinalizedFileName = RenamedFileName & " (" & Dup & ")"
        Else
            FinalizedFileName = RenamedFileName
        End If
        
        FinalizedFullPath = UsedNonRenamedFolderPathCnP & "\" & FinalizedFileName
        
    Else
    ' Check if there is a folder called "Renamed Files" in the 'UsedFolderPathCnP'.
        UsedRenamedFolderPathCnP = UsedFolderPathCnP & "\Renamed Files"
            
        If Len(Dir(UsedRenamedFolderPathCnP, vbDirectory)) = 0 Then
            MkDir UsedRenamedFolderPathCnP
        End If
            
        RenamedFileName = FileNameStart & " (" & SchoolID & ") - " & SchoolAbbrev & " - " & ReportType
        RenamedFullPath = UsedRenamedFolderPathCnP & "\" & RenamedFileName
        
        If Len(Dir(RenamedFullPath & ".csv", vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) <> 0 Then
            Dup = 1
            
            Do While Len(Dir(RenamedFullPath & " (" & Dup & ").csv", vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) <> 0
                Dup = Dup + 1
            Loop
            
            FinalizedFileName = RenamedFileName & " (" & Dup & ")"
        Else
            FinalizedFileName = RenamedFileName
        End If
        
        FinalizedFullPath = UsedRenamedFolderPathCnP & "\" & FinalizedFileName
        
    End If
    
    wbTemp.SaveAs FileName:=FinalizedFullPath & ".csv", FileFormat:=xlCSV, CreateBackup:=False
    SetAttr FinalizedFullPath & ".csv", vbReadOnly


' Move the original file to the 'Used Files' folder
    ' Make sure there is not an existing file name in there that has the same name
        ' Build the full source file path (original file in the selected folder).
            CurrentFilePath = FolderPathCnP & "\" & FileNamesList(FileNumber)

    ' Build a collision-safe destination path inside "Unused Files".
        NewFilePath = UsedFolderPathCnP & "\" & FileNamesList(FileNumber)
    
    ' If a file with the same name already exists, append " (n)" before the extension.
        If Len(Dir(NewFilePath, vbNormal Or vbHidden Or vbSystem)) > 0 Then
            DotPos = InStrRev(NewFilePath, ".")
            If DotPos > 0 Then
                BaseName = Left$(NewFilePath, DotPos - 1)
                Ext = Mid$(NewFilePath, DotPos) ' includes the dot
            Else
                BaseName = NewFilePath
                Ext = ""
            End If
    
            i = 1
            Do
                i = i + 1
                NewFilePath = BaseName & " (" & i & ")" & Ext
            Loop While Len(Dir(NewFilePath, vbNormal Or vbHidden Or vbSystem)) > 0
        End If
    
    ' Move the file.
        Name CurrentFilePath As NewFilePath

' Find the last row of the 'wsConsolidated' worksheet
    ConsolidatedLastRow = wsConsolidated.Cells(wsConsolidated.Rows.Count, "A").End(xlUp).Row + 1

' Manipulate the file to be uniform with all Click and Pledge files
    DataStartRow = HeaderRow + 1
    
    '''''--- ProPay Reports ---'''''
    If ReportTypeFull = "ProPay - Daily" Or ReportTypeFull = "ProPay - Monthly" Then
    ' Add 2 columns before 'Net Amount'
        wsTemp.Columns(10).Insert Shift:=xlRight
        wsTemp.Columns(10).Insert Shift:=xlRight
    ' Add the column headers
        wsTemp.Range("J" & HeaderRow).Value = "Fee"
        wsTemp.Range("K" & HeaderRow).Value = "Total Fees"
    ' Add in the formula
        wsTemp.Range("K" & DataStartRow).Formula = "=IF(AND(G" & DataStartRow & "=-300,H" & DataStartRow & "=0,I" & DataStartRow & "=0,J" & DataStartRow & "="""",L" & DataStartRow & "=-150),150,IF(AND(G" & DataStartRow & "=0,H" & DataStartRow & "=0,I" & DataStartRow & "=0,J" & DataStartRow & "="""",L" & DataStartRow & "<>0),L" & DataStartRow & ",IF(ISNUMBER(SEARCH(""_"",A" & DataStartRow & ")),J" & DataStartRow & "*-1,H" & DataStartRow & "+I" & DataStartRow & ")))"

        ' Fill Down
            If DataStartRow <> TempLastRow Then
                wsTemp.Range("K" & DataStartRow & ":K" & TempLastRow).FillDown
            End If
            
    '''''--- Stripe Reports ---'''''
    ElseIf ReportTypeFull = "Stripe - Daily" Or ReportTypeFull = "Stripe - Monthly" Then
    ' Add 1 column before 'Net Amount' and 2 columns before 'Fee'
        wsTemp.Columns(9).Insert Shift:=xlRight
        wsTemp.Columns(8).Insert Shift:=xlRight
        wsTemp.Columns(8).Insert Shift:=xlRight
    ' Add the column headers
        wsTemp.Range("H" & HeaderRow).Value = "Per Transaction Fees"
        wsTemp.Range("I" & HeaderRow).Value = "Discount Fees"
        wsTemp.Range("K" & HeaderRow).Value = "Total Fees"
    ' Add in the formula
        wsTemp.Range("K" & DataStartRow).Formula = "=IF(AND(G" & DataStartRow & "=-300,H" & DataStartRow & "=0,I" & DataStartRow & "=0,J" & DataStartRow & "="""",L" & DataStartRow & "=-150),150,IF(AND(G" & DataStartRow & "=0,H" & DataStartRow & "=0,I" & DataStartRow & "=0,J" & DataStartRow & "="""",L" & DataStartRow & "<>0),L" & DataStartRow & ",IF(ISNUMBER(SEARCH(""_"",A" & DataStartRow & ")),J" & DataStartRow & "*-1,H" & DataStartRow & "+I" & DataStartRow & ")))"

        ' Fill Down
            If DataStartRow <> TempLastRow Then
                wsTemp.Range("K" & DataStartRow & ":K" & TempLastRow).FillDown
            End If
    End If


    ' If the worksheet is duplicated, change the name it will have in the file
        If Dup = 0 Then
            WorksheetName = ReportYear & "." & ReportMonth & " - " & SchoolAbbrev & " (" & ReportType & ")"
        Else
            WorksheetName = ReportYear & "." & ReportMonth & " - " & SchoolAbbrev & " (" & ReportType & ") (" & Dup & ")"
        End If
    
    ' Format the 'wsTemp' worksheet.
        wsTemp.Columns("A:L").AutoFit
    
    If DataStartRow = TempLastRow Then
    ' Copy the 'wsTemp' data over to the 'wsConsolidated' worksheet
        wsTemp.Range("A" & DataStartRow & ":L" & DataStartRow).Copy Destination:=wsConsolidated.Range("A" & ConsolidatedLastRow)
    ' Make the vairable 'ConsolidatedLastRowNow' equal to ConsolidatedLastRow (This is used later in the macro)
        ConsolidatedLastRowNow = ConsolidatedLastRow
    ' Add in the original filename and the renamed filename
        wsConsolidated.Range("M" & ConsolidatedLastRow).Value = FileNamesList(FileNumber)
        wsConsolidated.Range("N" & ConsolidatedLastRow).Value = FinalizedFileName
        wsConsolidated.Range("O" & ConsolidatedLastRow).Value = WorksheetName
        wsConsolidated.Range("P" & ConsolidatedLastRow).Value = ReportTypeFull
        
    Else
    ' Copy the 'wsTemp' data over to the 'wsConsolidated' worksheet
        wsTemp.Range("A" & DataStartRow & ":L" & TempLastRow).Copy Destination:=wsConsolidated.Range("A" & ConsolidatedLastRow)
    ' Find the new last row of 'wsConsolidated' and store it in a variable called: 'ConsolidatedLastRowNow'
        ConsolidatedLastRowNow = wsConsolidated.Cells(wsConsolidated.Rows.Count, "A").End(xlUp).Row
    ' Add in the original filename and the renamed filename
        wsConsolidated.Range("M" & ConsolidatedLastRow & ":M" & ConsolidatedLastRowNow).Value = FileNamesList(FileNumber)
        wsConsolidated.Range("N" & ConsolidatedLastRow & ":N" & ConsolidatedLastRowNow).Value = FinalizedFileName
        wsConsolidated.Range("O" & ConsolidatedLastRow & ":O" & ConsolidatedLastRowNow).Value = WorksheetName
        wsConsolidated.Range("P" & ConsolidatedLastRow & ":P" & ConsolidatedLastRowNow).Value = ReportTypeFull
    End If

' Clear the clipboard
    Application.CutCopyMode = False


' Copy the data from the Click and Pledge Report into the 'wsConsolidated' Worksheet
    wsTemp.Copy After:=wbMacro.Sheets(wbMacro.Sheets.Count)
    Set wsNew = wbMacro.Sheets(wbMacro.Sheets.Count)
    wsNew.Name = WorksheetName
    
' Close the 'wbTemp' file.
    wbTemp.Close SaveChanges:=False
    
' Move on to the next file
    GoTo NextFile
    
DoNotUseFile:
    
    ' Close the temporary workbook without saving changes.
        On Error Resume Next
        If Not wbTemp Is Nothing Then
            wbTemp.Close SaveChanges:=False
        End If
        On Error GoTo 0
        
    ' Check if there is a folder called "Unused Files" in the original folder path, store the name in a variable called 'UnusedFolderPathCnP'
        UnusedFolderPathCnP = FolderPathCnP & "\Unused Files"
        
        ' If it is not created yet, create one
            If Len(Dir(UnusedFolderPathCnP, vbDirectory)) = 0 Then
                MkDir UnusedFolderPathCnP
            End If
            
    ' Check if the 'FileNamesList(FileNumber)' is in the 'UnusedFolderPathCnP'
        ' Build the full source file path (original file in the selected folder).
            CurrentFilePath = FolderPathCnP & "\" & FileNamesList(FileNumber)

    ' Build a collision-safe destination path inside "Unused Files".
        NewFilePath = UnusedFolderPathCnP & "\" & FileNamesList(FileNumber)
    
    ' If a file with the same name already exists, append " (n)" before the extension.
        If Len(Dir(NewFilePath, vbNormal Or vbHidden Or vbSystem)) > 0 Then
            DotPos = InStrRev(NewFilePath, ".")
            If DotPos > 0 Then
                BaseName = Left$(NewFilePath, DotPos - 1)
                Ext = Mid$(NewFilePath, DotPos) ' includes the dot
            Else
                BaseName = NewFilePath
                Ext = ""
            End If
    
            i = 1
            Do
                i = i + 1
                NewFilePath = BaseName & " (" & i & ")" & Ext
            Loop While Len(Dir(NewFilePath, vbNormal Or vbHidden Or vbSystem)) > 0
        End If
    
    ' Move the file.
        Name CurrentFilePath As NewFilePath

        ' Move on to the next file
NextFile:
    Next FileNumber

' If no files were used for the consolidated worksheet, create a message for the user and delete the 'wsConsolidated' worksheet.
    If (UsedProPayDailyFilesCount + UsedProPayMonthlyFilesCount + UsedStripeDailyFilesCount + UsedStripeMonthlyFilesCount) = 0 Then
        NonClickandPledgeFilesFound = NonExcelFilesCount + UnusedFilesCount
        ExtraMessage = "There were '" & NonClickandPledgeFilesFound & "' files found in the folder provided. However, no files matched the " & _
                    "ProPay or Stripe reports. Please  find the correct folder and try again."
        
        wsConsolidated.Delete
        
        GoTo NoFiles
    End If


' If files were used, populate the formulas into the 'wsConsolidated' worksheet
    ' Add the formulas
        ' Order Number Cleaned
            wsConsolidated.Range("Q2").Formula = "=TRIM(RIGHT(B2,LEN(B2)-4))"
            
        ' School ID
            wsConsolidated.Range("R2").Formula = "=IF(Q2=""Monthly Fee"",IF(ISNUMBER(SEARCH(""(?????)"",N2)),MID(N2,SEARCH(""(?????)"",N2)+1,5),""""),LEFT(Q2,5))"
            
        ' Disbursement ID
            wsConsolidated.Range("S2").Formula = "=IF(R2="""",TRIM(A2),TRIM(R2)&"" - ""&TRIM(A2))"
            
        ' Transaction Date (YYYY.MM.DD)
            wsConsolidated.Range("T2").Formula = "=TEXT(C2,""YYYY.MM.DD"")"
            
        ' Fund Date (YYYY.MM.DD)
            wsConsolidated.Range("U2").Formula = "=TEXT(D2,""YYYY.MM.DD"")"
            
        ' Donor Name (Last Name, First Name)
            wsConsolidated.Range("V2").Formula = "=F2&"", ""&E2"


        ' Fill Down (if the last row is greater than 2)
            If ConsolidatedLastRowNow > 2 Then
                wsConsolidated.Range("Q2:V" & ConsolidatedLastRowNow).FillDown
            End If
        
        ' Format the 'wsConsolidated' Worksheet.
            ' AutoFilter and AutoFit the columns
                wsConsolidated.Range("A1:V1").AutoFilter
                wsConsolidated.Columns("A:V").AutoFit
        

GoTo StandardizationProcess


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''----------------------------------------------------------''''''''''''''''''''
'''''''''''''''''''' Manipulate all the data to create an Intacct Import File ''''''''''''''''''''
''''''''''''''''''''----------------------------------------------------------''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
StandardizationProcess:

' Create the "User-Required Checks" worksheet. Store it in a variable called 'wsUserChecks'
    Set wsUserChecks = wbMacro.Worksheets.Add(After:=wsConsolidated)
    wsUserChecks.Name = "User-Required Checks"


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''---------------------------------''''''''''''''''''''
'''''''''''''''''''' Standardize the Salesforce data ''''''''''''''''''''
''''''''''''''''''''---------------------------------''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Update the status bar
        Application.StatusBar = "Standardizing Initial Report Data"
    
    ' Salesforce/Intacct --> "Standardized SF Data"
        ' Create a new worksheet called "Standardized SF Data". Store it in a variable called 'wsStandardSF'.
            Set wsStandardSF = wbMacro.Worksheets.Add(After:=wsInitialData)
            
            ' Rename the worksheet to "Standardized SF Data"
                wsStandardSF.Name = "Standardized SF Data"
                
            ' Create the column headers
                wsStandardSF.Range("A1:Z1").Value = Array("SF - Close Date (Transaction Date)", "SF - Deposit Date", "SF - School Name", "SF - Campaign Name", _
                        "SF - Opportunity Name", "SF - Payment Type", "SF - Check Number", "SF - PMT-ID", "SF - Family Name", "SF - Account Holder", _
                        "SF - CNP Order Number", "SF - Transaction ID", "SF - Disbursement ID", "SF - Amount", "SF - Company Name", "SF - Campaign Type", _
                        "SF - Campaign School Name", "Donation Site", "Account ID | Division ID | Funding Source", "Confident or Suggested", "Intacct - Location ID", _
                        "Intacct - Account ID", "Intacct - Division ID", "Intacct - Funding Source", "Intacct - Debt Services Series", "Intacct - Memo")
            
            ' Add in the data
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''''''''''''''''''-----------------------------------------------''''''''''''''''''''
                    '''''''''''''''''''' If the initial report was a Salesforce Report ''''''''''''''''''''
                    ''''''''''''''''''''-----------------------------------------------''''''''''''''''''''
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If ReportRoute = "Salesforce" Then
                        ' Update the status bar
                            Application.StatusBar = "Standardizing Initial Report Data (Salesforce)"
                            
                        ' Find the last row from the 'wsInitialData' worksheet
                            InitialLastRow = wsInitialData.Cells(wsInitialData.Rows.Count, "L").End(xlUp).Row
                            
                        ' Add in the formulas
                            ' For columns A-P: "SF - Close Date (Transaction Date)", "SF - Deposit Date", "SF - School Name", "SF - Campaign Name","SF - Opportunity Name", _
                                    "SF - Payment Type", "SF - Check Number", "SF - PMT-ID", "SF - Family Name", "SF - Account Holder", "SF - CNP Order Number", _
                                    "SF - Transaction ID", "SF - Disbursement ID", "SF - Amount", "SF - Company Name", "SF - Campaign Type",
                                wsStandardSF.Range("A2").Formula2 = "=IF(ISBLANK(CHOOSECOLS('Initial Data - SF'!A2:AF" & InitialLastRow & _
                                        ",23,32,31,8,7,21,32,12,5,6,13,13,32,18,32,28)),""""," & _
                                        "CHOOSECOLS('Initial Data - SF'!A2:AF" & InitialLastRow & ",23,32,31,8,7,21,32,12,5,6,13,13,32,18,32,28))"
                            
                            ' "SF - Campaign School Name"
                                wsStandardSF.Range("Q2").Formula = "=IF(D2="""",""No Campaign Name""," & _
                                        "IF(D2=""Ahwatukee ATF Teacher Talent Show"",""Ahwatukee""," & _
                                        "IF(D2=""Baton Rouge Mid City Father/Daughter Dance 2024"",""Baton Rouge Mid City""," & _
                                        "IF(D2=""BASIS Charter Schools, Inc."",""BCSI""," & _
                                        "IF(OR(D2=""BTCS Gala Table Sponsor"",D2=""BTCS Gala Transaction Fee Donation"",D2=""BTX Growth Fund"",D2=""BASIS Texas Charter Schools, Inc.""),""BASIS Texas Charter Schools, Inc.""," & _
                                        "IF(D2=""BASIS Peoria Holiday Brunch"",""Peoria""," & _
                                        "IF(D2=""BASIS Phoenix Students vs Teacher Basketball Game ATF Event"",""Phoenix""," & _
                                        "IF(ISNUMBER(SEARCH(""AZ Tax Credit"",D2)),LEFT(D2,SEARCH(""AZ Tax Credit"",D2)-2)," & _
                                        "IF(ISNUMBER(SEARCH(""Boosters ATF 20"",D2)),LEFT(D2,SEARCH(""Boosters ATF 20"",D2)-1)," & _
                                        "IF(ISNUMBER(SEARCH("" Capital Campaign"",D2)),LEFT(D2,SEARCH("" Capital Campaign"",D2)-1)," & _
                                        "IF(ISNUMBER(SEARCH(""General Fund"",D2)),LEFT(D2,SEARCH(""General Fund"",D2)-2)," & _
                                        "IF(ISNUMBER(SEARCH(""ATF 201"",D2)),LEFT(D2,SEARCH(""ATF 201"",D2)-1)," & _
                                        "IF(ISNUMBER(SEARCH(""ATF 202"",D2)),LEFT(D2,SEARCH(""ATF 202"",D2)-1)," & _
                                        "IF(ISNUMBER(SEARCH(""ATF 203"",D2)),LEFT(D2,SEARCH(""ATF 203"",D2)-1)," & _
                                        "IF(ISNUMBER(SEARCH("" 20"",D2)),LEFT(D2,SEARCH("" 20"",D2)-1),"""")))))))))))))))"
                                
                            ' "Donation Site"
                                wsStandardSF.Range("R2").Value = Site
                                
                            ' "Account ID | Division ID | Funding Source"
                                ' Breakdown as many campaign names into 3 variables to bring back together into 1 formula.
                                    CampaignBreakdown1 = "=IF(D2="""",""No Campaign Name""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Employer Match"",D2)),ISNUMBER(SEARCH(""Employee Match"",D2))),""73013|2048|7301-ATF Campaign""," & _
                                            "IF(ISNUMBER(SEARCH(""Arkansas 20"",D2)),""No Suggestions: Arkansas""," & _
                                            "IF(D2=""BTX Growth Fund"",""No Suggestions: BTX Growth Fund""," & _
                                            "IF(D2=""Goodyear 2020-21 Boosters General"",""No Suggestions: Goodyear 2020-21 Boosters General""," & _
                                            "IF(D2=""International Student Application"",""No Suggestions: International Student Application""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Founders Circle"",D2)),ISNUMBER(SEARCH(""Founders' Circle"",D2)),ISNUMBER(SEARCH(""BTCS Gala Table Sponsor"",D2)),ISNUMBER(SEARCH(""BTCS Gala Transaction Fee Donation"",D2))),""41005|2060|7306-Local Other (General)""," & _
                                            "IF(ISNUMBER(SEARCH(""Tax Credit"",D2)),""73001|2001|7312-Tax Credit""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Capital Campaign"",D2)),D2=""Firebirds Athletics Gym Banner Program"",D2=""Gym Banner Sponsorship"",D2=""Legacy Bricks Donations"",D2=""Legacy Brick Campaign - Vega, Canopus, and Sirius Level Donations"",D2=""Playground Sponsorships""),""73009|2036|7306-Local Other (General)"","
                                                                                
                                    CampaignBreakdown2 = "IF(ISNUMBER(SEARCH(""General Fund"",D2))," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Lunch"",E2)),ISNUMBER(SEARCH(""day, */*-"",E2))),""73010|2086|7311-Student Reimbursement""," & _
                                            "IF(ISNUMBER(SEARCH(""Founders Circle"",E2)),""Suggested (GF): 41005|2060|7306-Local Other (General)""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Aux Donation"",E2)),ISNUMBER(SEARCH(""Band - $*Donation"",E2)),ISNUMBER(SEARCH(""Band Family Donation"",E2)),ISNUMBER(SEARCH(""Band Individual Donation"",E2)),ISNUMBER(SEARCH(""Band - Suggested Family Donation"",E2)),ISNUMBER(SEARCH(""Band - Suggested Individual Donation"",E2)),ISNUMBER(SEARCH(""Drama Department Donation"",E2)),ISNUMBER(SEARCH(""Orchestra - $*Donation"",E2)),ISNUMBER(SEARCH(""Orchestra Family Donation"",E2)),ISNUMBER(SEARCH(""Orchestra - Suggested Family Donation"",E2)),ISNUMBER(SEARCH(""Orchestra - Suggested Individual Donation"",E2)),ISNUMBER(SEARCH(""Theater Donation"",E2))),""Suggested (GF): 73001|2001|7306-Local Other (General)""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Admission Tickets"",E2)),ISNUMBER(SEARCH(""Admission Tckets"",E2)),ISNUMBER(SEARCH(""Entrance Ticket"",E2)),""73010|2001|7306-Local Other (General)"",ISNUMBER(SEARCH(""General Admission Ticket-"",E2)),ISNUMBER(SEARCH(""Adult Ticket-"",E2)),ISNUMBER(SEARCH(""Student Ticket-"",E2)),ISNUMBER(SEARCH(""Open Adult Seating - "",E2)),ISNUMBER(SEARCH(""Reserve Adult Seating - "",E2)),ISNUMBER(SEARCH(""Open Student Seating - "",E2)),ISNUMBER(SEARCH(""Reserve Student Seating - "",E2)),ISNUMBER(SEARCH(""Volleyball Game Ticket-"",E2))),""Suggested (GF): 73010|2036|7306-Local Other (General)""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""ATF Direct Giving"",E2)),ISNUMBER(SEARCH(""AZ Give Fee"",E2)),ISNUMBER(SEARCH(""Direct ATF Donation"",E2)),ISNUMBER(SEARCH(""First Day Packet Donation"",E2)),ISNUMBER(SEARCH(""Annual Teacher Fund Donation"",E2)),ISNUMBER(SEARCH(""ATF Donation from"",E2)),ISNUMBER(SEARCH(""ATF Contribution"",E2)),ISNUMBER(SEARCH(""ATF Commitment Donations"",E2)),ISNUMBER(SEARCH(""ATF Box Top"",E2)),ISNUMBER(SEARCH(""Employee Giving"",E2))),""Suggested (GF): 73011|2048|7301-ATF Campaign""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""ATF Other Event"",E2)),ISNUMBER(SEARCH(""Indirect ATF Donation"",E2))),""Suggested (GF): 73012|2048|7301-ATF Campaign""," & _
                                            "IF(ISNUMBER(SEARCH(""Employer Matching"",E2)),""Suggested (GF): 73013|2048|7301-ATF Campaign""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""MG-Aetna"",E2)),ISNUMBER(SEARCH(""AmFam Giving"",E2)),ISNUMBER(SEARCH(""Bayer Disburse*on"",E2)),ISNUMBER(SEARCH(""Benevity giving"",E2)),ISNUMBER(SEARCH(""Cadence*on"",E2)),ISNUMBER(SEARCH(""Charles Schwab*on"",E2)),ISNUMBER(SEARCH(""Dell Technologies*on"",E2)),ISNUMBER(SEARCH(""IBM Disburse*on"",E2)),ISNUMBER(SEARCH(""IBM Disburse*on"",E2)),ISNUMBER(SEARCH(""Intel Foundation"",E2)),ISNUMBER(SEARCH(""LOLgives*on"",E2)),ISNUMBER(SEARCH(""MG-Intel Corporation"",E2)),ISNUMBER(SEARCH(""Macy's*on"",E2)),ISNUMBER(SEARCH(""Medtronic*on"",E2)),ISNUMBER(SEARCH(""Microsoft*on"",E2)),ISNUMBER(SEARCH(""MUFG*on"",E2)),ISNUMBER(SEARCH(""Oracle*on"",E2)),ISNUMBER(SEARCH(""Silicon Valley Bank Benevity"",E2)),ISNUMBER(SEARCH(""USAA"",E2)),ISNUMBER(SEARCH(""Wells Fargo*on"",E2))),""Suggested (GF): 73011 or 73013|2048|7301-ATF Campaign""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Camp Only-"",E2)),ISNUMBER(SEARCH(""Extended Care-"",E2)),ISNUMBER(SEARCH(""Junior Chefs"",E2)),ISNUMBER(SEARCH(""Late Bird*Week Package"",E2)),ISNUMBER(SEARCH(""Sports Camp -"",E2)),ISNUMBER(SEARCH(""Summer camp payment"",E2))),""No Suggestions (GF): Camp Related?""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Athletics Program Banner Sponsorship"",E2)),ISNUMBER(SEARCH(""Banner Donation"",E2)),ISNUMBER(SEARCH(""Extracurricular Fund Donation"",E2)),ISNUMBER(SEARCH(""General Athletic Sponsorship"",E2)),ISNUMBER(SEARCH(""Memorial Shade Structure"",E2)),ISNUMBER(SEARCH(""Teacher Technology Fund Donation"",E2))),""No Suggestions (GF): Capital Campaign?""," & _
                                            "IF(ISNUMBER(SEARCH(""Classy"",E2)),""No Suggestions (GF): Classy""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""$*Gift-"",E2)),ISNUMBER(SEARCH(""Additional Donation"",E2)),ISNUMBER(SEARCH(""Additional or Alternative Donation"",E2)),ISNUMBER(SEARCH(""Restricted Donation"",E2)),ISNUMBER(SEARCH(""General Donation"",E2)),ISNUMBER(SEARCH(""Other Donation from"",E2))),""No Suggestions (GF): Donations?""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Baklava Box"",E2)),ISNUMBER(SEARCH(""Balloon Sales"",E2)),ISNUMBER(SEARCH(""Book Fair"",E2)),ISNUMBER(SEARCH(""Fall Candy Gram Goodie Bag"",E2)),ISNUMBER(SEARCH(""Flowers from"",E2)),ISNUMBER(SEARCH(""Guest Prom Ticket"",E2)),ISNUMBER(SEARCH(""Participant Registration"",E2)),ISNUMBER(SEARCH(""Pencil Sales Profit"",E2)),ISNUMBER(SEARCH(""Spell-A-Thon"",E2)),ISNUMBER(SEARCH(""Yearbook Sale"",E2))),""No Suggestions: Event Related?""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Caramel Corn"",E2)),ISNUMBER(SEARCH(""Cinnamon Toast"",E2)),ISNUMBER(SEARCH(""Cone Basket"",E2)),ISNUMBER(SEARCH(""Jalape?o"",E2)),ISNUMBER(SEARCH(""Kettle Corn"",E2)),ISNUMBER(SEARCH(""White Cheddar"",E2)),ISNUMBER(SEARCH(""Zebra ("",E2))),""No Suggestions (GF): Event, Food""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Account Verification"",E2)),ISNUMBER(SEARCH(""Additional Fee"",E2)),ISNUMBER(SEARCH(""Transaction Fee"",E2))),""No Suggestions (GF): Fees""," & _
                                            "IF(ISNUMBER(SEARCH(""Parking Payment"",E2)),""No Suggestions (GF): Parking Payment""," & _
                                            "IF(ISNUMBER(SEARCH(""Partner with Excellence"",E2)),""No Suggestions (GF): Partner with Excellence""," & _
                                            "IF(ISNUMBER(SEARCH(""Red Cross Other"",E2)),""No Suggestions (GF): Red Cross""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Bundle 2: T1 Banner + Medium Logo"",E2)),ISNUMBER(SEARCH(""T-Shirt One Line"",E2)),ISNUMBER(SEARCH(""T-Shirt Main Logo"",E2))),""No Suggestions (GF): T-Shirt""," & _
                                            "IF(ISNUMBER(SEARCH(""General Fund"",E2)),""No Suggestions (GF): General Fund - General Fund""," & _
                                            """No Suggestions: General Fund"")))))))))))))))))))),"
    
                                    CampaignBreakdown3 = "IF(D2=""General Donations"",""Suggested: 73001|2001|7306-Local Other(General)""," & _
                                            "IF(OR(P2=""Direct"",ISNUMBER(SEARCH(""Employee Giving"",D2)),D2=""All school direct giving 2016-17"",D2=""Peoria Primary 2024-25 4th Grade Yearbook Tribute"",D2=""Security Deposit Donation (SMART)""),""73011|2048|7301-ATF Campaign""," & _
                                            "IF(OR(P2=""Indirect"",D2=""All school indirect 2016-17"",D2=""BASIS Peoria Holiday Brunch"",D2=""Baton Rouge Mid City Father/Daughter Dance 2024"",D2=""Goodyear 2018-19 Phoenix Suns Tickets"",D2=""Scottsdale Primary West 2019-20 Spring Week of Giving"",D2=""Popcornopolis"",D2=""Under The Stars Gala""),""73012|2048|7301-ATF Campaign""," & _
                                            "IF(OR(ISNUMBER(SEARCH("" ATF"",D2)),ISNUMBER(SEARCH(""Annual Teacher Fund"",D2))),""Suggested: 73011 or 73012|2048|7301-ATF Campaign""," & _
                                            "IF(OR(AND(TRIM(D2)=""BASIS Charter Schools, Inc."",ISNUMBER(SEARCH(""American Express"",E2))),AND(TRIM(D2)=""BASIS Texas Charter Schools, Inc."",OR(ISNUMBER(SEARCH(""PayPal Giving"",E2)),ISNUMBER(SEARCH(""USAA"",E2))))),""Suggested: 73011 or 73013|2048|7301-ATF Campaign""," & _
                                            """"")))))))))))))))"
                                            
                                ' Stitch the variables together into 1 formula
                                    wsStandardSF.Range("S2").Formula2 = CampaignBreakdown1 & CampaignBreakdown2 & CampaignBreakdown3
                                
                            ' "Confident or Suggested"
                                wsStandardSF.Range("T2").Formula = "=IF(ISNUMBER(SEARCH(""Suggest"",S2)),""Suggested"",""Confident"")"
                                
                            ' "Intacct - Location ID"
                                wsStandardSF.Range("U2").Formula = "=IF(OR(ConvertSalesforceToSchoolLocation(Q2)=""No Campaign Name"",ConvertSalesforceToSchoolLocation(Q2)=""No School Found""),ConvertCnPToIntacctAccount(LEFT(K2,5)),ConvertSalesforceToSchoolLocation(Q2))"
                                
                            ' For columns V-X: "Intacct - Account ID", "Intacct - Division ID", "Intacct - Funding Source"
                                wsStandardSF.Range("V2").Formula2 = "=IF(T2=""Confident"",TEXTSPLIT(S2,""|""),IF(S2=""Suggested: 73011 or 73012|2048|7301-ATF Campaign"",TEXTSPLIT(""73011|2048|7301-ATF Campaign"",""|""),""CHECK""))"
                                
                            ' "Intacct - Debt Services Series"
                                If ShowFormulas = True Then
                                    wsStandardSF.Range("Y2").Formula = "=""000"""
                                Else
                                    wsStandardSF.Range("Y2").Formula = "=""'000"""
                                End If
                                
                            ' "Intacct - Memo"
                                wsStandardSF.Range("Z2").Value = ""
                                
                        ' Fill Down
                            ' Find the last row of the 'wsStandardSF' worksheet.
                                StandardSFLastRow = wsStandardSF.Cells(wsStandardSF.Rows.Count, "H").End(xlUp).Row
                                
                            ' Fill the formulas down
                                If StandardSFLastRow > 2 Then
                                    wsStandardSF.Range("Q2:Z" & StandardSFLastRow).FillDown
                                End If
                            
                        ' Copy and Paste the values only
                            If ShowFormulas = False Then
                                wsStandardSF.Range("A:X").Value = wsStandardSF.Range("A:X").Value
                                wsStandardSF.Range("Z:Z").Value = wsStandardSF.Range("Z:Z").Value
                            End If
                            
                            
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''''''''''''''''''---------------------------------------------''''''''''''''''''''
                    '''''''''''''''''''' If the initial report was an Intacct Report ''''''''''''''''''''
                    ''''''''''''''''''''---------------------------------------------''''''''''''''''''''
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ElseIf ReportRoute = "Intacct" Then
                        ' Update the status bar
                            Application.StatusBar = "Standardizing Initial Report Data (Intacct)"
                        
                        ' Find the last row from the 'wsInitialData' worksheet
                            InitialLastRow = wsInitialData.Cells(wsInitialData.Rows.Count, "A").End(xlUp).Row
                            
                        ' Add in the formulas
                            ' For columns A-O:
                                ' "SF - Close Date (Transaction Date)", "SF - Deposit Date", "SF - School Name", "SF - Campaign Name","SF - Opportunity Name", _
                                    "SF - Payment Type", "SF - Check Number", "SF - PMT-ID", "SF - Family Name", "SF - Account Holder", "SF - CNP Order Number", _
                                    "SF - Transaction ID", "SF - Disbursement ID", "SF - Amount", "SF - Company Name",
                                wsStandardSF.Range("A2").Formula2 = "=IF(ISBLANK(CHOOSECOLS('Initial Data - Intacct'!A2:AI" & InitialLastRow & ",3,35,35,9,18,17,16,10,19,20,13,13,15,34,12))," & _
                                        """"",CHOOSECOLS('Initial Data - Intacct'!A2:AI" & InitialLastRow & ",3,35,35,9,18,17,16,10,19,20,13,13,15,34,12))"
                            
                            ' "SF - Campaign Type",
                                wsStandardSF.Range("P2").Value = ""
                                
                            ' "SF - Campaign School Name"
                                wsStandardSF.Range("Q2").Formula = "=IF(D2="""",""No Campaign Name""," & _
                                        "IF(D2=""Ahwatukee ATF Teacher Talent Show"",""Ahwatukee""," & _
                                        "IF(D2=""Baton Rouge Mid City Father/Daughter Dance 2024"",""Baton Rouge Mid City""," & _
                                        "IF(D2=""BASIS Charter Schools, Inc."",""BCSI""," & _
                                        "IF(OR(D2=""BTCS Gala Table Sponsor"",D2=""BTCS Gala Transaction Fee Donation"",D2=""BTX Growth Fund"",D2=""BASIS Texas Charter Schools, Inc.""),""BASIS Texas Charter Schools, Inc.""," & _
                                        "IF(D2=""BASIS Peoria Holiday Brunch"",""Peoria""," & _
                                        "IF(D2=""BASIS Phoenix Students vs Teacher Basketball Game ATF Event"",""Phoenix""," & _
                                        "IF(ISNUMBER(SEARCH(""AZ Tax Credit"",D2)),LEFT(D2,SEARCH(""AZ Tax Credit"",D2)-2)," & _
                                        "IF(ISNUMBER(SEARCH(""Boosters ATF 20"",D2)),LEFT(D2,SEARCH(""Boosters ATF 20"",D2)-1)," & _
                                        "IF(ISNUMBER(SEARCH("" Capital Campaign"",D2)),LEFT(D2,SEARCH("" Capital Campaign"",D2)-1)," & _
                                        "IF(ISNUMBER(SEARCH(""General Fund"",D2)),LEFT(D2,SEARCH(""General Fund"",D2)-2)," & _
                                        "IF(ISNUMBER(SEARCH(""ATF 201"",D2)),LEFT(D2,SEARCH(""ATF 201"",D2)-1)," & _
                                        "IF(ISNUMBER(SEARCH(""ATF 202"",D2)),LEFT(D2,SEARCH(""ATF 202"",D2)-1)," & _
                                        "IF(ISNUMBER(SEARCH(""ATF 203"",D2)),LEFT(D2,SEARCH(""ATF 203"",D2)-1)," & _
                                        "IF(ISNUMBER(SEARCH("" 20"",D2)),LEFT(D2,SEARCH("" 20"",D2)-1),"""")))))))))))))))"
                                
                            ' "Donation Site"
                                wsStandardSF.Range("R2").Value = Site
                                
                            ' "Account ID | Division ID | Funding Source"
                                ' Breakdown as many campaign names into 3 variables to bring back together into 1 formula.
                                    CampaignBreakdown1 = "=IF(D2="""",""No Campaign Name""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Employer Match"",D2)),ISNUMBER(SEARCH(""Employee Match"",D2))),""73013|2048|7301-ATF Campaign""," & _
                                            "IF(ISNUMBER(SEARCH(""Arkansas 20"",D2)),""No Suggestions: Arkansas""," & _
                                            "IF(D2=""BTX Growth Fund"",""No Suggestions: BTX Growth Fund""," & _
                                            "IF(D2=""Goodyear 2020-21 Boosters General"",""No Suggestions: Goodyear 2020-21 Boosters General""," & _
                                            "IF(D2=""International Student Application"",""No Suggestions: International Student Application""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Founders Circle"",D2)),ISNUMBER(SEARCH(""Founders' Circle"",D2)),ISNUMBER(SEARCH(""BTCS Gala Table Sponsor"",D2)),ISNUMBER(SEARCH(""BTCS Gala Transaction Fee Donation"",D2))),""41005|2060|7306-Local Other (General)""," & _
                                            "IF(ISNUMBER(SEARCH(""Tax Credit"",D2)),""73001|2001|7312-Tax Credit""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Capital Campaign"",D2)),D2=""Firebirds Athletics Gym Banner Program"",D2=""Gym Banner Sponsorship"",D2=""Legacy Bricks Donations"",D2=""Legacy Brick Campaign - Vega, Canopus, and Sirius Level Donations"",D2=""Playground Sponsorships""),""73009|2036|7306-Local Other (General)"","
                                                                                
                                    CampaignBreakdown2 = "IF(ISNUMBER(SEARCH(""General Fund"",D2))," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Lunch"",E2)),ISNUMBER(SEARCH(""day, */*-"",E2))),""73010|2086|7311-Student Reimbursement""," & _
                                            "IF(ISNUMBER(SEARCH(""Founders Circle"",E2)),""Suggested (GF): 41005|2060|7306-Local Other (General)""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Aux Donation"",E2)),ISNUMBER(SEARCH(""Band - $*Donation"",E2)),ISNUMBER(SEARCH(""Band Family Donation"",E2)),ISNUMBER(SEARCH(""Band Individual Donation"",E2)),ISNUMBER(SEARCH(""Band - Suggested Family Donation"",E2)),ISNUMBER(SEARCH(""Band - Suggested Individual Donation"",E2)),ISNUMBER(SEARCH(""Drama Department Donation"",E2)),ISNUMBER(SEARCH(""Orchestra - $*Donation"",E2)),ISNUMBER(SEARCH(""Orchestra Family Donation"",E2)),ISNUMBER(SEARCH(""Orchestra - Suggested Family Donation"",E2)),ISNUMBER(SEARCH(""Orchestra - Suggested Individual Donation"",E2)),ISNUMBER(SEARCH(""Theater Donation"",E2))),""Suggested (GF): 73001|2001|7306-Local Other (General)""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Admission Tickets"",E2)),ISNUMBER(SEARCH(""Admission Tckets"",E2)),ISNUMBER(SEARCH(""Entrance Ticket"",E2)),""73010|2001|7306-Local Other (General)"",ISNUMBER(SEARCH(""General Admission Ticket-"",E2)),ISNUMBER(SEARCH(""Adult Ticket-"",E2)),ISNUMBER(SEARCH(""Student Ticket-"",E2)),ISNUMBER(SEARCH(""Open Adult Seating - "",E2)),ISNUMBER(SEARCH(""Reserve Adult Seating - "",E2)),ISNUMBER(SEARCH(""Open Student Seating - "",E2)),ISNUMBER(SEARCH(""Reserve Student Seating - "",E2)),ISNUMBER(SEARCH(""Volleyball Game Ticket-"",E2))),""Suggested (GF): 73010|2036|7306-Local Other (General)""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""ATF Direct Giving"",E2)),ISNUMBER(SEARCH(""AZ Give Fee"",E2)),ISNUMBER(SEARCH(""Direct ATF Donation"",E2)),ISNUMBER(SEARCH(""First Day Packet Donation"",E2)),ISNUMBER(SEARCH(""Annual Teacher Fund Donation"",E2)),ISNUMBER(SEARCH(""ATF Donation from"",E2)),ISNUMBER(SEARCH(""ATF Contribution"",E2)),ISNUMBER(SEARCH(""ATF Commitment Donations"",E2)),ISNUMBER(SEARCH(""ATF Box Top"",E2)),ISNUMBER(SEARCH(""Employee Giving"",E2))),""Suggested (GF): 73011|2048|7301-ATF Campaign""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""ATF Other Event"",E2)),ISNUMBER(SEARCH(""Indirect ATF Donation"",E2))),""Suggested (GF): 73012|2048|7301-ATF Campaign""," & _
                                            "IF(ISNUMBER(SEARCH(""Employer Matching"",E2)),""Suggested (GF): 73013|2048|7301-ATF Campaign""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""MG-Aetna"",E2)),ISNUMBER(SEARCH(""AmFam Giving"",E2)),ISNUMBER(SEARCH(""Bayer Disburse*on"",E2)),ISNUMBER(SEARCH(""Benevity giving"",E2)),ISNUMBER(SEARCH(""Cadence*on"",E2)),ISNUMBER(SEARCH(""Charles Schwab*on"",E2)),ISNUMBER(SEARCH(""Dell Technologies*on"",E2)),ISNUMBER(SEARCH(""IBM Disburse*on"",E2)),ISNUMBER(SEARCH(""IBM Disburse*on"",E2)),ISNUMBER(SEARCH(""Intel Foundation"",E2)),ISNUMBER(SEARCH(""LOLgives*on"",E2)),ISNUMBER(SEARCH(""MG-Intel Corporation"",E2)),ISNUMBER(SEARCH(""Macy's*on"",E2)),ISNUMBER(SEARCH(""Medtronic*on"",E2)),ISNUMBER(SEARCH(""Microsoft*on"",E2)),ISNUMBER(SEARCH(""MUFG*on"",E2)),ISNUMBER(SEARCH(""Oracle*on"",E2)),ISNUMBER(SEARCH(""Silicon Valley Bank Benevity"",E2)),ISNUMBER(SEARCH(""USAA"",E2)),ISNUMBER(SEARCH(""Wells Fargo*on"",E2))),""Suggested (GF): 73011 or 73013|2048|7301-ATF Campaign""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Camp Only-"",E2)),ISNUMBER(SEARCH(""Extended Care-"",E2)),ISNUMBER(SEARCH(""Junior Chefs"",E2)),ISNUMBER(SEARCH(""Late Bird*Week Package"",E2)),ISNUMBER(SEARCH(""Sports Camp -"",E2)),ISNUMBER(SEARCH(""Summer camp payment"",E2))),""No Suggestions (GF): Camp Related?""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Athletics Program Banner Sponsorship"",E2)),ISNUMBER(SEARCH(""Banner Donation"",E2)),ISNUMBER(SEARCH(""Extracurricular Fund Donation"",E2)),ISNUMBER(SEARCH(""General Athletic Sponsorship"",E2)),ISNUMBER(SEARCH(""Memorial Shade Structure"",E2)),ISNUMBER(SEARCH(""Teacher Technology Fund Donation"",E2))),""No Suggestions (GF): Capital Campaign?""," & _
                                            "IF(ISNUMBER(SEARCH(""Classy"",E2)),""No Suggestions (GF): Classy""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""$*Gift-"",E2)),ISNUMBER(SEARCH(""Additional Donation"",E2)),ISNUMBER(SEARCH(""Additional or Alternative Donation"",E2)),ISNUMBER(SEARCH(""Restricted Donation"",E2)),ISNUMBER(SEARCH(""General Donation"",E2)),ISNUMBER(SEARCH(""Other Donation from"",E2))),""No Suggestions (GF): Donations?""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Baklava Box"",E2)),ISNUMBER(SEARCH(""Balloon Sales"",E2)),ISNUMBER(SEARCH(""Book Fair"",E2)),ISNUMBER(SEARCH(""Fall Candy Gram Goodie Bag"",E2)),ISNUMBER(SEARCH(""Flowers from"",E2)),ISNUMBER(SEARCH(""Guest Prom Ticket"",E2)),ISNUMBER(SEARCH(""Participant Registration"",E2)),ISNUMBER(SEARCH(""Pencil Sales Profit"",E2)),ISNUMBER(SEARCH(""Spell-A-Thon"",E2)),ISNUMBER(SEARCH(""Yearbook Sale"",E2))),""No Suggestions: Event Related?""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Caramel Corn"",E2)),ISNUMBER(SEARCH(""Cinnamon Toast"",E2)),ISNUMBER(SEARCH(""Cone Basket"",E2)),ISNUMBER(SEARCH(""Jalape?o"",E2)),ISNUMBER(SEARCH(""Kettle Corn"",E2)),ISNUMBER(SEARCH(""White Cheddar"",E2)),ISNUMBER(SEARCH(""Zebra ("",E2))),""No Suggestions (GF): Event, Food""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Account Verification"",E2)),ISNUMBER(SEARCH(""Additional Fee"",E2)),ISNUMBER(SEARCH(""Transaction Fee"",E2))),""No Suggestions (GF): Fees""," & _
                                            "IF(ISNUMBER(SEARCH(""Parking Payment"",E2)),""No Suggestions (GF): Parking Payment""," & _
                                            "IF(ISNUMBER(SEARCH(""Partner with Excellence"",E2)),""No Suggestions (GF): Partner with Excellence""," & _
                                            "IF(ISNUMBER(SEARCH(""Red Cross Other"",E2)),""No Suggestions (GF): Red Cross""," & _
                                            "IF(OR(ISNUMBER(SEARCH(""Bundle 2: T1 Banner + Medium Logo"",E2)),ISNUMBER(SEARCH(""T-Shirt One Line"",E2)),ISNUMBER(SEARCH(""T-Shirt Main Logo"",E2))),""No Suggestions (GF): T-Shirt""," & _
                                            "IF(ISNUMBER(SEARCH(""General Fund"",E2)),""No Suggestions (GF): General Fund - General Fund""," & _
                                            """No Suggestions: General Fund"")))))))))))))))))))),"
    
                                    CampaignBreakdown3 = "IF(D2=""General Donations"",""Suggested: 73001|2001|7306-Local Other(General)""," & _
                                            "IF(OR(P2=""Direct"",ISNUMBER(SEARCH(""Employee Giving"",D2)),D2=""All school direct giving 2016-17"",D2=""Peoria Primary 2024-25 4th Grade Yearbook Tribute"",D2=""Security Deposit Donation (SMART)""),""73011|2048|7301-ATF Campaign""," & _
                                            "IF(OR(P2=""Indirect"",D2=""All school indirect 2016-17"",D2=""BASIS Peoria Holiday Brunch"",D2=""Baton Rouge Mid City Father/Daughter Dance 2024"",D2=""Goodyear 2018-19 Phoenix Suns Tickets"",D2=""Scottsdale Primary West 2019-20 Spring Week of Giving"",D2=""Popcornopolis"",D2=""Under The Stars Gala""),""73012|2048|7301-ATF Campaign""," & _
                                            "IF(OR(ISNUMBER(SEARCH("" ATF"",D2)),ISNUMBER(SEARCH(""Annual Teacher Fund"",D2))),""Suggested: 73011 or 73012|2048|7301-ATF Campaign""," & _
                                            "IF(OR(AND(TRIM(D2)=""BASIS Charter Schools, Inc."",ISNUMBER(SEARCH(""American Express"",E2))),AND(TRIM(D2)=""BASIS Texas Charter Schools, Inc."",OR(ISNUMBER(SEARCH(""PayPal Giving"",E2)),ISNUMBER(SEARCH(""USAA"",E2))))),""Suggested: 73011 or 73013|2048|7301-ATF Campaign""," & _
                                            """"")))))))))))))))"
                                            
                                ' Stitch the variables together into 1 formula
                                    wsStandardSF.Range("S2").Formula2 = CampaignBreakdown1 & CampaignBreakdown2 & CampaignBreakdown3
                                
                            ' "Confident or Suggested"
                                wsStandardSF.Range("T2").Formula = "=IF(ISNUMBER(SEARCH(""Suggest"",S2)),""Suggested"",""Confident"")"
                                
                            ' For columns U-Z:
                                ' "Intacct - Location ID", "Intacct - Account ID", "Intacct - Division ID", "Intacct - Funding Source", "Intacct - Debt Services Series", "Intacct - Memo"
                                wsStandardSF.Range("U2").Formula2 = "=IF(ISBLANK(CHOOSECOLS('Initial Data - Intacct'!A2:AI" & InitialLastRow & ",6,4,25,27,30,8))," & _
                                        """"",CHOOSECOLS('Initial Data - Intacct'!A2:AI" & InitialLastRow & ",6,4,25,27,30,8))"
                            
                        ' Fill Down
                            ' Find the last row of the 'wsStandardSF' worksheet.
                                StandardSFLastRow = wsStandardSF.Cells(wsStandardSF.Rows.Count, "H").End(xlUp).Row
                                
                            ' Fill the formulas down
                                If StandardSFLastRow > 2 Then
                                    wsStandardSF.Range("Q2:T" & StandardSFLastRow).FillDown
                                End If
                            
                        ' Copy and Paste the values only
                            If ShowFormulas = False Then
                                wsStandardSF.Range("A:X").Value = wsStandardSF.Range("A:X").Value
                                wsStandardSF.Range("Z:Z").Value = wsStandardSF.Range("Z:Z").Value
                            End If
                    End If

    ' Format the worksheet.
        ' Change the date format.
            wsStandardSF.Range("A:B").NumberFormat = "mm/dd/yyyy"
        
        ' AutoFilter Columns.
            wsStandardSF.Range("A1:Z1").AutoFilter
            
        ' AutoFit Columns.
            wsStandardSF.Columns("A:Z").AutoFit
            
    ' Hide the 'wsInitialData' worksheet.
        wsInitialData.Visible = xlSheetHidden


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''-----------------------------------''''''''''''''''''''
'''''''''''''''''''' Standardize the Consolidated data ''''''''''''''''''''
''''''''''''''''''''-----------------------------------''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Consolidated Donation Site Data --> "Standardized Donation Site Data"
        ' Update the status bar
            Application.StatusBar = "Standardizing Donation Site Data"
        
        ' Create a new worksheet called "Standard Donation Site Report". Store it in a variable called 'wsStandardDonations'.
            Set wsStandardDonations = wbMacro.Worksheets.Add(After:=wsConsolidated)
            
            ' Rename the worksheet to "Standardized Donation Site Data"
                wsStandardDonations.Name = "Standardized Donation Site Data"
                
            ' Create the column headers
                wsStandardDonations.Range("A1:R1").Value = Array("Donation Site", "Transaction Date", "Disbursement Date", "Transaction ID", "Disbursement ID", _
                        "Donor Name (Last Name, First Name)", "Donation Gross Amount", "Donation Total Fees", "Donation Net Amount", "Donation Type", "Donation Method", _
                        "Check Number", "Company", "Site - School ID", "Site - School Abbrev.", "Transaction Number (Amount)", "CHECK REQUIRED", "Corrected - School Abbreviation")
                
            ' Add in the data
                ConsolidatedLastRow = wsConsolidated.Cells(wsConsolidated.Rows.Count, "A").End(xlUp).Row
                
                ' "Donation Site"
                    wsStandardDonations.Range("A2").Value = Site
                
                ' Columns B-N: "Transaction Date", "Disbursement Date", "Transaction ID", "Disbursement ID", "Donor Name (Last Name, First Name)", "Donation Gross Amount", _
                        "Donation Total Fees", "Donation Net Amount", "Donation Type", "Donation Method", "Check Number", "Company", "Site - School ID"
                    wsStandardDonations.Range("B2").Formula2 = "=IF(ISBLANK(CHOOSECOLS('Consolidated Reports'!A2:W" & ConsolidatedLastRow & ",20,21,17,19,22,7,11,12,23,23,23,23,18)),""""," & _
                            "CHOOSECOLS('Consolidated Reports'!A2:W" & ConsolidatedLastRow & ",3,4,17,19,22,7,11,12,23,23,23,23,18))"
                        
                ' "Site - School Abbrev."
                    wsStandardDonations.Range("O2").Formula = "=IF(N2="""",IF(R2="""",""CHECK"",R2),ConvertCnPToSchoolAbbrev(N2))"
                    
                ' "Transaction Number (Amount)"
                    wsStandardDonations.Range("P2").Formula = "=D2&"" (""&G2&"")"""
                    
                ' "CHECK REQUIRED"
                    wsStandardDonations.Range("Q2").Formula = "=IF(N2="""",""CHECK"",ConvertCnPToSchoolAbbrev(N2))"
                    
                ' "Corrected - School Abbreviation"
                    wsStandardDonations.Range("R2").Formula = "=XLOOKUP(E2,'User-Required Checks'!C:C,'User-Required Checks'!F:F,"""")"
                
            ' Fill Down
                ' Find the last row of the 'wsStandardDonations' worksheet and store it in a variable called 'StandardDonationsLastRow'.
                    StandardDonationsLastRow = wsStandardDonations.Cells(wsStandardDonations.Rows.Count, "B").End(xlUp).Row
                
                ' Fill the formulas down
                    If StandardDonationsLastRow > 2 Then
                        wsStandardDonations.Range("A2:A" & StandardDonationsLastRow).FillDown
                        wsStandardDonations.Range("O2:R" & StandardDonationsLastRow).FillDown
                    End If
                        
            ' Copy and Paste the values only
                If ShowFormulas = False Then
                    wsStandardDonations.Range("A:R").Value = wsStandardDonations.Range("A:O").Value
                End If
            
            ' Format the worksheet.
                ' Change the date format.
                    wsStandardDonations.Range("B:C").NumberFormat = "mm/dd/yyyy"
                
                ' AutoFilter Columns.
                    wsStandardDonations.Range("A1:R1").AutoFilter
                    
                ' AutoFit Columns.
                    wsStandardDonations.Columns("A:R").AutoFit

                    
        ' Hide the 'wsConsolidated' worksheet.
            wsConsolidated.Visible = xlSheetHidden
                    
                    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''-------------------------------------''''''''''''''''''''
'''''''''''''''''''' Manipulate and Analyze all the data ''''''''''''''''''''
''''''''''''''''''''-------------------------------------''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Disbursement Breakdown
        ' Update Status Bar
            Application.StatusBar = "Breaking Down All Disbursements"
    
        ' Create a new worksheet called "Disbursements Breakdown". Store it in a variable called 'wsDisbursements'.
            Set wsDisbursements = wbMacro.Worksheets.Add(After:=wsStandardDonations)
            
            ' Rename the worksheet to "Disbursements Breakdown"
                wsDisbursements.Name = "Disbursements Breakdown"
            
            ' Create the column headers
                wsDisbursements.Range("A1:Z1").Value = Array("Site", "Report Type", "Disbursement ID", "Disbursement Date", "Positive Transactions", _
                        "Negative Transactions", "Total Transactions", "Gross Amount", "Transaction Fees", "Monthly Fee", "Total Fees", "Net Amount", "Bank Fees", _
                        "Miscellaneous Fees", "Miscellaneous Fees Notes", "Bank Deposit Amount", "Donation Method", "Check Number", "Company", "Site - School ID", _
                        "Site - School Name", "School Abbrev.", "Intacct - Deposit Description", "Intacct - Bank Account Number", "Intacct - School Location", _
                        "Expected Bank Deposit Name")
                
            ' Add in the data
                ' "Site"
                    wsDisbursements.Range("A2").Formula = Site
                    
                ' "Report Type"
                    wsDisbursements.Range("B2").Formula = "=IF(A2=""Click and Pledge"",IF(ISNUMBER(SEARCH(""po_"",C2)),""Stripe"",""ProPay""),"""")"
                    
                ' "Disbursement ID"
                    wsDisbursements.Range("C2").Formula2 = "=UNIQUE('Standardized Donation Site Data'!E2:E" & StandardDonationsLastRow & ")"
                 
                ' "Disbursement Date"
                    wsDisbursements.Range("D2").Formula = "=TEXT(MAX(FILTER(UNIQUE(FILTER('Standardized Donation Site Data'!C2:C" & StandardDonationsLastRow & _
                            ",'Standardized Donation Site Data'!E2:E" & StandardDonationsLastRow & "=C2)),UNIQUE(FILTER('Standardized Donation Site Data'!C2:C" & _
                            StandardDonationsLastRow & ",'Standardized Donation Site Data'!E2:E" & StandardDonationsLastRow & "=C2))<>"""")),""MM/DD/YYYY"")"
                 
                ' "Positive Transactions"
                    wsDisbursements.Range("E2").Formula = "=COUNTIFS('Standardized Donation Site Data'!E:E,C2,'Standardized Donation Site Data'!I:I,"">0"")"
                 
                ' "Negative Transactions"
                    wsDisbursements.Range("F2").Formula = "=COUNTIFS('Standardized Donation Site Data'!E:E,C2,'Standardized Donation Site Data'!I:I,""<0"")"
                 
                ' "Total Transactions"
                    wsDisbursements.Range("G2").Formula = "=COUNTIFS('Standardized Donation Site Data'!E:E,C2)"
                 
                '"Gross Amount"
                    wsDisbursements.Range("H2").Formula = "=SUMIFS('Standardized Donation Site Data'!G:G,'Standardized Donation Site Data'!E:E,C2)"
                 
                ' "Transaction Fees"
                    wsDisbursements.Range("I2").Formula = "=SUMIFS('Standardized Donation Site Data'!H:H,'Standardized Donation Site Data'!E:E,C2,'Standardized Donation Site Data'!D:D,""<>Monthly Fee"")"
                 
                ' "Monthly Fee"
                    wsDisbursements.Range("J2").Formula = "=SUMIFS('Standardized Donation Site Data'!H:H,'Standardized Donation Site Data'!E:E,C2,'Standardized Donation Site Data'!D:D,""Monthly Fee"")"
                 
                ' "Total Fees"
                    wsDisbursements.Range("K2").Formula = "=SUMIFS('Standardized Donation Site Data'!H:H,'Standardized Donation Site Data'!E:E,C2)"
                 
                ' "Net Amount"
                    wsDisbursements.Range("L2").Formula = "=SUMIFS('Standardized Donation Site Data'!I:I,'Standardized Donation Site Data'!E:E,C2)"
                    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' "Bank Fees"
                    wsDisbursements.Range("M2").Formula = "=0"
                 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' "Miscellaneous Fees"
                    wsDisbursements.Range("N2").Formula = "=0"
                 
                ' "Miscellaneous Fees Notes"
                    wsDisbursements.Range("O2").Formula = ""
                 
                ' "Bank Deposit Amount"
                    wsDisbursements.Range("P2").Formula = "=L2+M2+N2"
                 
                ' "Donation Method"
                    wsDisbursements.Range("Q2").Formula = "=""Credit Card"""
                 
                ' "Check Number"
                    wsDisbursements.Range("R2").Formula = "="""""
                 
                ' "Company"
                    wsDisbursements.Range("S2").Formula = "="""""
                    
                ' "Site - School ID"
                    wsDisbursements.Range("T2").Formula = "=IF(ISBLANK(XLOOKUP(C2,'Standardized Donation Site Data'!E:E,'Standardized Donation Site Data'!N:N)),""""," & _
                            "XLOOKUP(C2,'Standardized Donation Site Data'!E:E,'Standardized Donation Site Data'!N:N))"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' "Site - School Name"
                    wsDisbursements.Range("U2").Formula = "="""""
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' "School Abbrev."
                    wsDisbursements.Range("V2").Formula = "=XLOOKUP(C2,'Standardized Donation Site Data'!E:E,'Standardized Donation Site Data'!O:O)"
                    
                ' "Intacct - Deposit Description"
                    wsDisbursements.Range("W2").Formula = "=""Click and Pledge {""&B2&""} - ""&V2&"" (""&C2&"") [""&TEXT(P2,""$#,##0.00"")&""]"""
                
                ' "Intacct - Bank Account Number"
                    wsDisbursements.Range("X2").Formula2 = "=ConvertSchoolAbbrevToBankAccount(V2)"
                    
                ' "Intacct - School Location"
                    wsDisbursements.Range("Y2").Formula2 = "=ConvertSchoolAbbrevToIntacctAccount(V2)"
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' "Expected Bank Deposit Name"
                    wsDisbursements.Range("Z2").Formula = ""
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ' Fill Down
                ' Find the last row of the 'wsDisbursements' worksheet and store it in a variable called 'DisburementsLastRow'.
                    DisbursementsLastRow = wsDisbursements.Cells(wsDisbursements.Rows.Count, "C").End(xlUp).Row
                
                ' Fill the formulas down
                    If DisbursementsLastRow > 2 Then
                        wsDisbursements.Range("A2:B" & DisbursementsLastRow).FillDown
                        wsDisbursements.Range("D2:Z" & DisbursementsLastRow).FillDown
                    End If
                        
            ' Copy and Paste the values only
                If ShowFormulas = False Then
                    wsDisbursements.Range("A:Z").Value = wsDisbursements.Range("A:Z").Value
                End If
            
            ' Format the worksheet.
                ' AutoFilter Columns.
                    wsDisbursements.Range("A1:Z1").AutoFilter
                
                ' AutoFit Columns.
                    wsDisbursements.Columns("A:Z").AutoFit
            
            ' Hide the worksheet
                wsDisbursements.Visible = xlSheetHidden
            
            

    ' Positive Transactions
        ' Update the status bar
            Application.StatusBar = "Matching All Positive Transactions"
            
        ' Create a worksheet called "Positive Transactions". Store it in a variable called 'wsPosTransactions'.
            Set wsPosTransactions = wbMacro.Worksheets.Add(After:=wsDisbursements)
            
            ' Rename the worksheet to "Positive Transactions"
                wsPosTransactions.Name = "Positive Transactions"
            
            ' Create the column headers
                wsPosTransactions.Range("A1:DD1").Value = Array("SF - Close Date (Transaction Date)", "SF - Deposit Date", "SF - School Name", "SF - Campaign Name", _
                        "SF - Opportunity Name", "SF - Payment Type", "SF - Check Number", "SF - PMT-ID", "SF - Family Name", "SF - Account Holder", "SF - CNP Order Number", _
                        "SF - Transaction ID", "SF - Disbursement ID", "SF - Amount", "SF - Company Name", "SF - Campaign Type", "SF - Campaign School Name", "Donation Site", _
                        "Account ID | Division ID | Funding Source", "Confident or Suggested", "Intacct - Location ID", "Intacct - Account ID", "Intacct - Division ID", _
                        "Intacct - Funding Source", "Intacct - Debt Services Series", "Intacct - Memo", _
                    "<---SF Data............Donation Site Data--->", _
                        "Donation Site", "Transaction Date", "Disbursement Date", "Transaction ID", "Disbursement ID", "Donor Name (Last Name, First Name)", "Donation Type", _
                        "Check Number", "Company", "Site - School ID", "Site - School Abbrev.", _
                    "<---Donation Site Data.......SF/Donation Site Data Combined--->", _
                        "Site - Intacct Bank Account", "Site - Intacct Bank Account Name", "Intacct - Journal Description", "Intacct School ID", "Donation Type", _
                        "Donation Gross Amount (Negative)", "Site - School Abbrev.", "SF - Campaign Name", "SF - Opportunity Name", "Site Name", "Company", "SF - Payment Type", _
                        "Check Number", "SF - PMT-ID", "SF - Transaction ID", "Site - Disbursement ID", "Site - Transaction Date (MM.DD.YYYY)", _
                        "Site - Disbursement Date (MM.DD.YYYY)", "Transaction # ____ of", "Transaction # of ____", "SF - Family Name", "SF - Primary Account Holder", _
                    "<---SF/Donation Site Data Combined........Adjusting Journal for Intacct--->", _
                        "Intacct - Deposit Date", "Intacct - Deposit Description", "Intacct - Line Number", "Intacct - Account Number", "Intacct - Location ID", _
                        "Intacct - Department ID", "Intacct - Memo", "Intacct - Amount", "Intacct - Funding Source", "Intacct - GL Entry Class ID (Debt Services)", _
                    "<---Adjusting Journal for Intacct.............CRJ for Intacct--->", _
                        "RECEIPT_DATE", "PAYMETHOD", "DOCDATE", "DOCNUMBER", "DESCRIPTION", "DEPOSITTO", "BANKACCOUNTID", "DEPOSITDATE", "UNDEPACCTNO", "CURRENCY", "EXCH_RATE_DATE", _
                        "EXCH_RATE_TYPE_ID", "EXCH_RATE_DATE", "LINE_NO", "ACCT_NO", "ACCOUNTLABEL", "TRX_AMOUNT", "AMOUNT", "DEPT_ID", "LOCATION_ID", "ITEM_MEMO", _
                        "OTHERRECEIPTSENTRY_PROJECTID", "OTHERRECEIPTSENTRY_CUSTOMERID", "OTHERRECEIPTSENTRY_ITEMID", "OTHERRECEIPTSENTRY_VENDORID", "OTHERRECEIPTSENTRY_EMPLOYEEID", _
                        "OTHERRECEIPTSENTRY_CLASSID", "PAYER_NAME", "SUPDOCID", "EXCHANGE_RATE", "OR_TRANSACTION_DATE", "GLDIMFUNDING_SOURCE", _
                        "Positive or Negative Disbursement", "Full Disbursement Gross Amount", "Transaction ID (Amount)")
                        
            ' Add in the data
                ' For Columns A-Z: "SF - Close Date (Transaction Date)", "SF - Deposit Date", "SF - School Name", "SF - Campaign Name", "SF - Opportunity Name", "SF - Payment Type", _
                        "SF - Check Number", "SF - PMT-ID", "SF - Family Name", "SF - Account Holder", "SF - CNP Order Number", "SF - Transaction ID", "SF - Disbursement ID", _
                        "SF - Amount", "SF - Company Name", "SF - Campaign Type", "SF - Campaign School Name", "Donation Site", "Account ID | Division ID | Funding Source", _
                        "Confident or Suggested", "Intacct - Location ID", "Intacct - Account ID", "Intacct - Division ID", "Intacct - Funding Source", "Intacct - Debt Services Series", _
                        "Intacct - Memo"
                    wsPosTransactions.Range("A2").Formula2 = "=IFERROR(SORT(IF(ISBLANK(FILTER('Standardized SF Data'!A2:Z" & StandardSFLastRow & ",ISNUMBER(MATCH('Standardized SF Data'!L2:L" & StandardSFLastRow & _
                            ",FILTER('Standardized Donation Site Data'!D:D,'Standardized Donation Site Data'!I:I>0),0)))),"""",FILTER('Standardized SF Data'!A2:Z" & StandardSFLastRow & _
                            ",ISNUMBER(MATCH('Standardized SF Data'!L2:L" & StandardSFLastRow & ",FILTER('Standardized Donation Site Data'!D:D,'Standardized Donation Site Data'!I:I>0),0)))),12),""No Results Found"")"
                ' =IFERROR(SORT(IF(ISBLANK(FILTER('Standardized SF Data'!A2:Z" & StandardSFLastRow & ",ISNUMBER(MATCH('Standardized SF Data'!L2:L" & StandardSFLastRow & _
                        ",FILTER('Standardized Donation Site Data'!D:D,'Standardized Donation Site Data'!I:I>0),0)))),"""",FILTER('Standardized SF Data'!A2:Z" & StandardSFLastRow & _
                        ",ISNUMBER(MATCH('Standardized SF Data'!L2:L" & StandardSFLastRow & ",FILTER('Standardized Donation Site Data'!D:D,'Standardized Donation Site Data'!I:I>0),0)))),12)," & _
                        ""No matches found"")
                
                ' <---SF Data............Donation Site Data--->
                    wsPosTransactions.Range("AA2").Value = "<---SF Data............Donation Site Data--->"
                
                ' Donation Site
                    wsPosTransactions.Range("AB2").Formula = "=XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!A:A)"
                
                ' Transaction Date
                    wsPosTransactions.Range("AC2").Formula = "=XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!B:B)"
                
                ' Disbursement Date
                    wsPosTransactions.Range("AD2").Formula = "=XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!C:C)"
                
                ' Transaction ID
                    wsPosTransactions.Range("AE2").Formula = "=XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!D:D)"
                
                ' Disbursement ID
                 ''''' Below '''''
                    
                ' Donor Name (Last Name, First Name)
                    wsPosTransactions.Range("AG2").Formula = "=XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!F:F)"
                
                ' Donation Type
                    wsPosTransactions.Range("AH2").Formula = "=IF(ISBLANK(XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!J:J,))," & _
                            """"",XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!J:J,))"
                
                ' Check Number
                    wsPosTransactions.Range("AI2").Formula = "=IF(ISBLANK(XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!L:L))," & _
                            """"",XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!L:L))"
                
                ' Company
                    wsPosTransactions.Range("AJ2").Formula = "=IF(ISBLANK(XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!M:M))," & _
                            """"",XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!M:M))"
                
                ' Site - School ID
                    wsPosTransactions.Range("AK2").Formula = "=XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!N:N)"
                
                ' Site - School Abbrev.
                    wsPosTransactions.Range("AL2").Formula = "=XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!O:O)"
                
                ' <---Donation Site Data.......SF/Donation Site Data Combined--->
                    wsPosTransactions.Range("AM2").Value = "<---Donation Site Data.......SF/Donation Site Data Combined--->"
                
                ' Site - Intacct Bank Account
                    wsPosTransactions.Range("AN2").Formula = "=ConvertCnPToBankAccount(AK2)"
                
                ' Site - Intacct Bank Account Name
                    wsPosTransactions.Range("AO2").Formula = "=ConvertBankAccountToBankAccountName(AN2)"
                
                ' Intacct - Journal Description
                    wsPosTransactions.Range("AP2").Formula = "=XLOOKUP(AF2,'Disbursements Breakdown'!C:C,'Disbursements Breakdown'!W:W)"
                
                ' Intacct School ID
                    wsPosTransactions.Range("AQ2").Formula = "=U2"
                
                ' Donation Type
                    wsPosTransactions.Range("AR2").Formula = "=IF(ISNUMBER(SEARCH(""Employee Giving"",D2)),""Employee Giving"",IF(ISNUMBER(SEARCH(""Employer Matching"",D2)),""Employer Matching"",""""))"
                
                ' Donation Gross Amount
                    wsPosTransactions.Range("AS2").Formula = "=N2"
                
                ' Site - School Abbrev.
                    wsPosTransactions.Range("AT2").Formula = "=AL2"
                
                ' SF - Campaign Name
                    wsPosTransactions.Range("AU2").Formula = "=D2"
                
                ' SF - Opportunity Name
                    wsPosTransactions.Range("AV2").Formula = "=E2"
                
                ' Site Name
                    wsPosTransactions.Range("AW2").Value = Site
                
                ' Company
                    wsPosTransactions.Range("AX2").Formula = "=IF(O2="""","""",O2)"
                
                ' SF - Payment Type
                    wsPosTransactions.Range("AY2").Formula = "=IF(F2=""Check"",""Check:"",IF(OR(F2=""VISA"",F2=""MasterCard"",F2=""American Express"",F2=""Discover""),""Credit Card"",F2))"
                
                ' Check Number
                    wsPosTransactions.Range("AZ2").Formula = "=IF(G2="""","""",G2)"
                
                ' SF - PMT-ID
                    wsPosTransactions.Range("BA2").Formula = "=H2"
                
                ' SF - Transaction ID
                    wsPosTransactions.Range("BB2").Formula = "=L2"
                
                ' Site - Disbursement ID
                    wsPosTransactions.Range("BC2").Formula = "=AF2"
                
                ' Site - Transaction Date (MM.DD.YYYY)
                    wsPosTransactions.Range("BD2").Formula = "=TEXT(AC2,""MM.DD.YYYY"")"
                
                ' Site - Disbursement Date (MM.DD.YYYY)
                    wsPosTransactions.Range("BE2").Formula = "=TEXT(AD2,""MM.DD.YYYY"")"
                
                ' Transaction # ____ of
                    wsPosTransactions.Range("BF2").Formula = "=IF(L2=L1,BF1+1,1)"
                
                ' Transaction # of ____
                    wsPosTransactions.Range("BG2").Formula = "=COUNTIFS(L:L,L2)"
                
                ' SF - Family Name
                    wsPosTransactions.Range("BH2").Formula = "=I2"
                
                ' SF - Primary Account Holder
                    wsPosTransactions.Range("BI2").Formula = "=J2"
                
                ' <---SF/Donation Site Data Combined........Adjusting Journal for Intacct--->
                    wsPosTransactions.Range("BJ2").Value = "<---SF/Donation Site Data Combined........Adjusting Journal for Intacct--->"
                
                ' Intacct - Deposit Date
                    wsPosTransactions.Range("BK2").Formula = "=TEXT(AD2,""MM/DD/YYYY"")"
                
                ' Intacct - Deposit Description
                    wsPosTransactions.Range("BL2").Formula = "=AP2"
                
                ' Intacct - Line Number
                    wsPosTransactions.Range("BM2").Formula = "="""""
                
                ' Intacct - Account Number
                    wsPosTransactions.Range("BN2").Formula = "=V2"
                
                ' Intacct - Location ID
                    wsPosTransactions.Range("BO2").Formula = "=AQ2"
                
                ' Intacct - Department ID
                    wsPosTransactions.Range("BP2").Formula = "=W2"
                
                ' Intacct - Memo
                    wsPosTransactions.Range("BQ2").Formula = "=IF(Z2<>"""",""Reclassed Out: ""&Z2,AT2&"" - ""&AU2&"" | ""&AV2&"" | Site: ""&AW2&" & _
                            "IF(AX2="""","" | "","" | Company: ""&AX2)&"" | ""&IF(AZ2<>"""",AY2&"" ""&AZ2,AY2)&"" | ""&BA2&"" | Transaction ID: ""&BB2&" & _
                            """ | Disbursement ID: ""&BC2&"" | ~Transaction Date: ""&BD2&"" | Transaction # ""&BF2&"" of ""&BG2&"" | ^""&BH2&"" | *""&BI2)"
                
                ' Intacct - Amount
                    wsPosTransactions.Range("BR2").Formula = "=AS2"
                
                ' Intacct - Funding Source
                    wsPosTransactions.Range("BS2").Formula = "=X2"
                
                ' Intacct - GL Entry Class ID (Debt Services)
                    wsPosTransactions.Range("BT2").Formula = "=Y2"
                
                ' <---Adjusting Journal for Intacct.............CRJ for Intacct--->
                    wsPosTransactions.Range("BU2").Value = "<---Adjusting Journal for Intacct.............CRJ for Intacct--->"
                
                ' RECEIPT_DATE
                    wsPosTransactions.Range("BV2").Formula = "=BK2"
                
                ' PAYMETHOD
                    wsPosTransactions.Range("BW2").Formula = "=""Credit Card"""
                
                ' DOCDATE
                    wsPosTransactions.Range("BX2").Formula = "=BV2"
                
                ' DOCNUMBER
                    wsPosTransactions.Range("BY2").Formula = "=XLOOKUP(AP2,'Disbursements Breakdown'!W:W,'Disbursements Breakdown'!B:B)"
                
                ' DESCRIPTION
                    wsPosTransactions.Range("BZ2").Formula = "=LEFT(BL2,SEARCH(""["",BL2)-2)"
                
                ' DEPOSITTO
                    wsPosTransactions.Range("CA2").Formula = "=""Bank account"""
                
                ' BANKACCOUNTID
                    wsPosTransactions.Range("CB2").Formula = "=AO2"
                
                ' DEPOSITDATE
                    wsPosTransactions.Range("CC2").Formula = "=BV2"
                
                ' UNDEPACCTNO
                    wsPosTransactions.Range("CD2").Formula = "="""""
                
                ' CURRENCY
                    wsPosTransactions.Range("CE2").Formula = "=""USD"""
                
                ' EXCH_RATE_DATE
                    wsPosTransactions.Range("CF2").Formula = "="""""
                
                ' EXCH_RATE_TYPE_ID
                    wsPosTransactions.Range("CG2").Formula = "="""""
                
                ' EXCH_RATE_DATE
                    wsPosTransactions.Range("CH2").Formula = "="""""
                
                ' LINE_NO
                    wsPosTransactions.Range("CI2").Formula = "="""""
                
                ' ACCT_NO
                    wsPosTransactions.Range("CJ2").Formula = "=BN2"
                
                ' ACCOUNTLABEL
                    wsPosTransactions.Range("CK2").Formula = "="""""
                
                ' TRX_AMOUNT
                    wsPosTransactions.Range("CL2").Formula = "=BR2"
                
                ' AMOUNT
                    wsPosTransactions.Range("CM2").Formula = "=BR2"
                
                ' DEPT_ID
                    wsPosTransactions.Range("CN2").Formula = "=BP2"
                
                ' LOCATION_ID
                    wsPosTransactions.Range("CO2").Formula = "=BO2"
                
                ' ITEM_MEMO
                    wsPosTransactions.Range("CP2").Formula = "=BQ2"
                
                ' OTHERRECEIPTSENTRY_PROJECTID
                    wsPosTransactions.Range("CQ2").Formula = "="""""
                
                ' OTHERRECEIPTSENTRY_CUSTOMERID
                    wsPosTransactions.Range("CR2").Formula = "=IF(BY2=""ProPay"",""19234"",""8179"")"
                
                ' OTHERRECEIPTSENTRY_ITEMID
                    wsPosTransactions.Range("CS2").Formula = "="""""
                
                ' OTHERRECEIPTSENTRY_VENDORID
                    wsPosTransactions.Range("CT2").Formula = "="""""
                
                ' OTHERRECEIPTSENTRY_EMPLOYEEID
                    wsPosTransactions.Range("CU2").Formula = "="""""
                
                ' OTHERRECEIPTSENTRY_CLASSID
                    wsPosTransactions.Range("CV2").Formula = "=BT2"
                
                ' PAYER_NAME
                    wsPosTransactions.Range("CW2").Formula = "=BY2"
                
                ' SUPDOCID
                    wsPosTransactions.Range("CX2").Formula = "="""""
                
                ' EXCHANGE_RATE
                    wsPosTransactions.Range("CY2").Formula = "="""""
                
                ' OR_TRANSACTION_DATE
                    wsPosTransactions.Range("CZ2").Formula = "=BV2"
                
                ' GLDIMFUNDING_SOURCE
                    wsPosTransactions.Range("DA2").Formula = "=BS2"
                    
                ' Positive or Negative Disbursement
                    ''''' Below '''''
                    
                ' Full Disbursement Gross Amount
                    wsPosTransactions.Range("DC2").Formula = "=SUMIFS(AS:AS,AE:AE,AE2)"
                
                ' Transaction ID (Amount)
                    wsPosTransactions.Range("DD2").Formula = "=AE2&"" (""&DC2&"")"""
                    
                ' Disbursement ID
                    wsPosTransactions.Range("AF2").Formula = "=XLOOKUP(DD2,'Standardized Donation Site Data'!P:P,'Standardized Donation Site Data'!E:E)"
                    
                ' Positive or Negative Disbursement
                    wsPosTransactions.Range("DB2").Formula = "=IF(XLOOKUP(AF2,'Disbursements Breakdown'!C:C,'Disbursements Breakdown'!L:L)>0,""Positive"",""Negative"")"
                    
            ' Fill Down
                ' Find the last row of the 'wsPosTransactions' worksheet and store it in a variable called 'PosTransactionLastRow'.
                    PosTransactionsLastRow = wsPosTransactions.Cells(wsPosTransactions.Rows.Count, "C").End(xlUp).Row
                
                ' Fill the formulas down
                    If PosTransactionsLastRow > 2 Then
                        wsPosTransactions.Range("AA2:DD" & PosTransactionsLastRow).FillDown
                    End If
                        
            ' Copy and Paste the values only
                If ShowFormulas = False Then
                    wsPosTransactions.Range("A:DD").Value = wsPosTransactions.Range("A:DD").Value
                Else
                    wsPosTransactions.Range("DC:DD").Value = wsPosTransactions.Range("DC:DD").Value
                End If
            
            ' Format the worksheet.
                ' Change the date format.
                    wsPosTransactions.Range("A:B").NumberFormat = "mm/dd/yyyy"
                    
                ' AutoFilter Columns.
                    wsPosTransactions.Range("A1:DD1").AutoFilter
                
                ' AutoFit Columns.
                    wsPosTransactions.Columns("A:DD").AutoFit


            ' Hide the worksheet
                wsPosTransactions.Visible = xlSheetHidden
                    

    ' Negative Transactions
        ' Update the status bar
            Application.StatusBar = "Matching All Negative Transactions"
            
        ' Create a worksheet called "Negative Transactions". Store it in a variable called 'wsNegTransactions'.
            Set wsNegTransactions = wbMacro.Worksheets.Add(After:=wsPosTransactions)
            
            ' Rename the worksheet to "Negative Transactions"
                wsNegTransactions.Name = "Negative Transactions"
            
            ' Create the column headers
                wsNegTransactions.Range("A1:DD1").Value = Array("SF - Close Date (Transaction Date)", "SF - Deposit Date", "SF - School Name", "SF - Campaign Name", _
                        "SF - Opportunity Name", "SF - Payment Type", "SF - Check Number", "SF - PMT-ID", "SF - Family Name", "SF - Account Holder", "SF - CNP Order Number", _
                        "SF - Transaction ID", "SF - Disbursement ID", "SF - Amount", "SF - Company Name", "SF - Campaign Type", "SF - Campaign School Name", "Donation Site", _
                        "Account ID | Division ID | Funding Source", "Confident or Suggested", "Intacct - Location ID", "Intacct - Account ID", "Intacct - Division ID", _
                        "Intacct - Funding Source", "Intacct - Debt Services Series", "Intacct - Memo", _
                    "<---SF Data............Donation Site Data--->", _
                        "Donation Site", "Transaction Date", "Disbursement Date", "Transaction ID", "Disbursement ID", "Donor Name (Last Name, First Name)", "Donation Type", _
                        "Check Number", "Company", "Site - School ID", "Site - School Abbrev.", _
                    "<---Donation Site Data.......SF/Donation Site Data Combined--->", _
                        "Site - Intacct Bank Account", "Site - Intacct Bank Account Name", "Intacct - Journal Description", "Intacct School ID", "Donation Type", _
                        "Donation Gross Amount (Negative)", "Site - School Abbrev.", "SF - Campaign Name", "SF - Opportunity Name", "Site Name", "Company", "SF - Payment Type", _
                        "Check Number", "SF - PMT-ID", "SF - Transaction ID", "Site - Disbursement ID", "Site - Transaction Date (MM.DD.YYYY)", _
                        "Site - Disbursement Date (MM.DD.YYYY)", "Transaction # ____ of", "Transaction # of ____", "SF - Family Name", "SF - Primary Account Holder", _
                    "<---SF/Donation Site Data Combined........Adjusting Journal for Intacct--->", _
                        "Intacct - Deposit Date", "Intacct - Deposit Description", "Intacct - Line Number", "Intacct - Account Number", "Intacct - Location ID", _
                        "Intacct - Department ID", "Intacct - Memo", "Intacct - Amount", "Intacct - Funding Source", "Intacct - GL Entry Class ID (Debt Services)", _
                    "<---Adjusting Journal for Intacct.............CRJ for Intacct--->", _
                        "RECEIPT_DATE", "PAYMETHOD", "DOCDATE", "DOCNUMBER", "DESCRIPTION", "DEPOSITTO", "BANKACCOUNTID", "DEPOSITDATE", "UNDEPACCTNO", "CURRENCY", "EXCH_RATE_DATE", _
                        "EXCH_RATE_TYPE_ID", "EXCH_RATE_DATE", "LINE_NO", "ACCT_NO", "ACCOUNTLABEL", "TRX_AMOUNT", "AMOUNT", "DEPT_ID", "LOCATION_ID", "ITEM_MEMO", _
                        "OTHERRECEIPTSENTRY_PROJECTID", "OTHERRECEIPTSENTRY_CUSTOMERID", "OTHERRECEIPTSENTRY_ITEMID", "OTHERRECEIPTSENTRY_VENDORID", "OTHERRECEIPTSENTRY_EMPLOYEEID", _
                        "OTHERRECEIPTSENTRY_CLASSID", "PAYER_NAME", "SUPDOCID", "EXCHANGE_RATE", "OR_TRANSACTION_DATE", "GLDIMFUNDING_SOURCE", _
                        "Positive or Negative Disbursement", "Full Disbursement Gross Amount", "Transaction ID (Amount)")
                        
            ' Add in the data
                ' For Columns A-Z: "SF - Close Date (Transaction Date)", "SF - Deposit Date", "SF - School Name", "SF - Campaign Name", "SF - Opportunity Name", "SF - Payment Type", _
                        "SF - Check Number", "SF - PMT-ID", "SF - Family Name", "SF - Account Holder", "SF - CNP Order Number", "SF - Transaction ID", "SF - Disbursement ID", _
                        "SF - Amount", "SF - Company Name", "SF - Campaign Type", "SF - Campaign School Name", "Donation Site", "Account ID | Division ID | Funding Source", _
                        "Confident or Suggested", "Intacct - Location ID", "Intacct - Account ID", "Intacct - Division ID", "Intacct - Funding Source", "Intacct - Debt Services Series", _
                        "Intacct - Memo"
                    wsNegTransactions.Range("A2").Formula2 = "=IFERROR(SORT(IF(ISBLANK(FILTER('Standardized SF Data'!A2:Z" & StandardSFLastRow & ",ISNUMBER(MATCH('Standardized SF Data'!L2:L" & StandardSFLastRow & _
                            ",FILTER('Standardized Donation Site Data'!D:D,'Standardized Donation Site Data'!I:I<0),0)))),"""",FILTER('Standardized SF Data'!A2:Z" & StandardSFLastRow & _
                            ",ISNUMBER(MATCH('Standardized SF Data'!L2:L" & StandardSFLastRow & ",FILTER('Standardized Donation Site Data'!D:D,'Standardized Donation Site Data'!I:I<0),0)))),12),""No Results Found"")"


                ' <---SF Data............Donation Site Data--->
                    wsNegTransactions.Range("AA2").Value = "<---SF Data............Donation Site Data--->"
                
                ' Donation Site
                    wsNegTransactions.Range("AB2").Formula = "=XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!A:A)"
                
                ' Transaction Date
                    wsNegTransactions.Range("AC2").Formula = "=XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!B:B)"
                
                ' Disbursement Date
                    wsNegTransactions.Range("AD2").Formula = "=XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!C:C)"
                
                ' Transaction ID
                    wsNegTransactions.Range("AE2").Formula = "=XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!D:D)"
                
                ' Disbursement ID
                 ''''' Below '''''
                    
                ' Donor Name (Last Name, First Name)
                    wsNegTransactions.Range("AG2").Formula = "=XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!F:F)"
                
                ' Donation Type
                    wsNegTransactions.Range("AH2").Formula = "=IF(ISBLANK(XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!J:J,))," & _
                            """"",XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!J:J,))"
                
                ' Check Number
                    wsNegTransactions.Range("AI2").Formula = "=IF(ISBLANK(XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!L:L))," & _
                            """"",XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!L:L))"
                
                ' Company
                    wsNegTransactions.Range("AJ2").Formula = "=IF(ISBLANK(XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!M:M))," & _
                            """"",XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!M:M))"
                
                ' Site - School ID
                    wsNegTransactions.Range("AK2").Formula = "=XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!N:N)"
                
                ' Site - School Abbrev.
                    wsNegTransactions.Range("AL2").Formula = "=XLOOKUP($L2,'Standardized Donation Site Data'!$D:$D,'Standardized Donation Site Data'!O:O)"
                
                ' <---Donation Site Data.......SF/Donation Site Data Combined--->
                    wsNegTransactions.Range("AM2").Value = "<---Donation Site Data.......SF/Donation Site Data Combined--->"
                
                ' Site - Intacct Bank Account
                    wsNegTransactions.Range("AN2").Formula = "=ConvertCnPToBankAccount(AK2)"
                
                ' Site - Intacct Bank Account Name
                    wsNegTransactions.Range("AO2").Formula = "=ConvertBankAccountToBankAccountName(AN2)"
                
                ' Intacct - Journal Description
                    wsNegTransactions.Range("AP2").Formula = "=XLOOKUP(AF2,'Disbursements Breakdown'!C:C,'Disbursements Breakdown'!W:W)"
                
                ' Intacct School ID
                    wsNegTransactions.Range("AQ2").Formula = "=U2"
                
                ' Donation Type
                    wsNegTransactions.Range("AR2").Formula = "=IF(ISNUMBER(SEARCH(""Employee Giving"",D2)),""Employee Giving"",IF(ISNUMBER(SEARCH(""Employer Matching"",D2)),""Employer Matching"",""""))"
                
                ' Donation Gross Amount (Negative)
                    wsNegTransactions.Range("AS2").Formula = "=N2*-1"
                
                ' Site - School Abbrev.
                    wsNegTransactions.Range("AT2").Formula = "=AL2"
                
                ' SF - Campaign Name
                    wsNegTransactions.Range("AU2").Formula = "=D2"
                
                ' SF - Opportunity Name
                    wsNegTransactions.Range("AV2").Formula = "=E2"
                
                ' Site Name
                    wsNegTransactions.Range("AW2").Value = Site
                
                ' Company
                    wsNegTransactions.Range("AX2").Formula = "=IF(O2="""","""",O2)"
                
                ' SF - Payment Type
                    wsNegTransactions.Range("AY2").Formula = "=IF(F2=""Check"",""Check:"",IF(OR(F2=""VISA"",F2=""MasterCard"",F2=""American Express"",F2=""Discover""),""Credit Card"",F2))"
                
                ' Check Number
                    wsNegTransactions.Range("AZ2").Formula = "=IF(G2="""","""",G2)"
                
                ' SF - PMT-ID
                    wsNegTransactions.Range("BA2").Formula = "=H2"
                
                ' SF - Transaction ID
                    wsNegTransactions.Range("BB2").Formula = "=L2"
                
                ' Site - Disbursement ID
                    wsNegTransactions.Range("BC2").Formula = "=AF2"
                
                ' Site - Transaction Date (MM.DD.YYYY)
                    wsNegTransactions.Range("BD2").Formula = "=TEXT(AC2,""MM.DD.YYYY"")"
                
                ' Site - Disbursement Date (MM.DD.YYYY)
                    wsNegTransactions.Range("BE2").Formula = "=TEXT(AD2,""MM.DD.YYYY"")"
                
                ' Transaction # ____ of
                    wsNegTransactions.Range("BF2").Formula = "=IF(L2=L1,BF1+1,1)"
                
                ' Transaction # of ____
                    wsNegTransactions.Range("BG2").Formula = "=COUNTIFS(L:L,L2)"
                
                ' SF - Family Name
                    wsNegTransactions.Range("BH2").Formula = "=I2"
                
                ' SF - Primary Account Holder
                    wsNegTransactions.Range("BI2").Formula = "=J2"
                
                ' <---SF/Donation Site Data Combined........Adjusting Journal for Intacct--->
                    wsNegTransactions.Range("BJ2").Value = "<---SF/Donation Site Data Combined........Adjusting Journal for Intacct--->"
                
                ' Intacct - Deposit Date
                    wsNegTransactions.Range("BK2").Formula = "=TEXT(AD2,""MM/DD/YYYY"")"
                
                ' Intacct - Deposit Description
                    wsNegTransactions.Range("BL2").Formula = "=AP2"
                
                ' Intacct - Line Number
                    wsNegTransactions.Range("BM2").Formula = "="""""
                
                ' Intacct - Account Number
                    wsNegTransactions.Range("BN2").Formula = "=V2"
                
                ' Intacct - Location ID
                    wsNegTransactions.Range("BO2").Formula = "=AQ2"
                
                ' Intacct - Department ID
                    wsNegTransactions.Range("BP2").Formula = "=W2"
                
                ' Intacct - Memo
                    wsNegTransactions.Range("BQ2").Formula = "=IF(Z2<>"""",""Reclassed Out: Reversed: ""&Z2," & _
                            """Reversed: ""&AT2&"" - ""&AU2&"" | ""&AV2&"" | Site: ""&AW2&" & _
                            "IF(AX2="""","" | "","" | Company: ""&AX2)&"" | ""&IF(AZ2<>"""",AY2&"" ""&AZ2,AY2)&"" | ""&BA2&"" | Transaction ID: ""&BB2&" & _
                            """ | Disbursement ID: ""&BC2&"" | ~Transaction Date: ""&BD2&"" | Transaction # ""&BF2&"" of ""&BG2&"" | ^""&BH2&"" | *""&BI2)"
                
                ' Intacct - Amount
                    wsNegTransactions.Range("BR2").Formula = "=AS2"
                
                ' Intacct - Funding Source
                    wsNegTransactions.Range("BS2").Formula = "=X2"
                
                ' Intacct - GL Entry Class ID (Debt Services)
                    wsNegTransactions.Range("BT2").Formula = "=Y2"
                
                ' <---Adjusting Journal for Intacct.............CRJ for Intacct--->
                    wsNegTransactions.Range("BU2").Value = "<---Adjusting Journal for Intacct.............CRJ for Intacct--->"
                
                ' RECEIPT_DATE
                    wsNegTransactions.Range("BV2").Formula = "=BK2"
                
                ' PAYMETHOD
                    wsNegTransactions.Range("BW2").Formula = "=""Credit Card"""
                
                ' DOCDATE
                    wsNegTransactions.Range("BX2").Formula = "=BV2"
                
                ' DOCNUMBER
                    wsNegTransactions.Range("BY2").Formula = "=XLOOKUP(AP2,'Disbursements Breakdown'!W:W,'Disbursements Breakdown'!B:B)"
                
                ' DESCRIPTION
                    wsNegTransactions.Range("BZ2").Formula = "=LEFT(BL2,SEARCH(""["",BL2)-2)"
                
                ' DEPOSITTO
                    wsNegTransactions.Range("CA2").Formula = "=""Bank account"""
                
                ' BANKACCOUNTID
                    wsNegTransactions.Range("CB2").Formula = "=AO2"
                
                ' DEPOSITDATE
                    wsNegTransactions.Range("CC2").Formula = "=BV2"
                
                ' UNDEPACCTNO
                    wsNegTransactions.Range("CD2").Formula = "="""""
                
                ' CURRENCY
                    wsNegTransactions.Range("CE2").Formula = "=""USD"""
                
                ' EXCH_RATE_DATE
                    wsNegTransactions.Range("CF2").Formula = "="""""
                
                ' EXCH_RATE_TYPE_ID
                    wsNegTransactions.Range("CG2").Formula = "="""""
                
                ' EXCH_RATE_DATE
                    wsNegTransactions.Range("CH2").Formula = "="""""
                
                ' LINE_NO
                    wsNegTransactions.Range("CI2").Formula = "="""""
                
                ' ACCT_NO
                    wsNegTransactions.Range("CJ2").Formula = "=BN2"
                
                ' ACCOUNTLABEL
                    wsNegTransactions.Range("CK2").Formula = "="""""
                
                ' TRX_AMOUNT
                    wsNegTransactions.Range("CL2").Formula = "=BR2"
                
                ' AMOUNT
                    wsNegTransactions.Range("CM2").Formula = "=BR2"
                
                ' DEPT_ID
                    wsNegTransactions.Range("CN2").Formula = "=BP2"
                
                ' LOCATION_ID
                    wsNegTransactions.Range("CO2").Formula = "=BO2"
                
                ' ITEM_MEMO
                    wsNegTransactions.Range("CP2").Formula = "=BQ2"
                
                ' OTHERRECEIPTSENTRY_PROJECTID
                    wsNegTransactions.Range("CQ2").Formula = "="""""
                
                ' OTHERRECEIPTSENTRY_CUSTOMERID
                    wsNegTransactions.Range("CR2").Formula = "=IF(BY2=""ProPay"",""19234"",""8179"")"
                
                ' OTHERRECEIPTSENTRY_ITEMID
                    wsNegTransactions.Range("CS2").Formula = "="""""
                
                ' OTHERRECEIPTSENTRY_VENDORID
                    wsNegTransactions.Range("CT2").Formula = "="""""
                
                ' OTHERRECEIPTSENTRY_EMPLOYEEID
                    wsNegTransactions.Range("CU2").Formula = "="""""
                
                ' OTHERRECEIPTSENTRY_CLASSID
                    wsNegTransactions.Range("CV2").Formula = "=BT2"
                
                ' PAYER_NAME
                    wsNegTransactions.Range("CW2").Formula = "=BY2"
                
                ' SUPDOCID
                    wsNegTransactions.Range("CX2").Formula = "="""""
                
                ' EXCHANGE_RATE
                    wsNegTransactions.Range("CY2").Formula = "="""""
                
                ' OR_TRANSACTION_DATE
                    wsNegTransactions.Range("CZ2").Formula = "=BV2"
                
                ' GLDIMFUNDING_SOURCE
                    wsNegTransactions.Range("DA2").Formula = "=BS2"
                    
                ' Positive or Negative Disbursement
                 ''''' Below '''''
                    
                ' Full Disbursement Gross Amount
                    wsNegTransactions.Range("DC2").Formula = "=SUMIFS(AS:AS,AE:AE,AE2)"
                
                ' Transaction ID (Amount)
                    wsNegTransactions.Range("DD2").Formula = "=AE2&"" (""&DC2&"")"""
                    
                ' Disbursement ID
                    wsNegTransactions.Range("AF2").Formula = "=XLOOKUP(DD2,'Standardized Donation Site Data'!P:P,'Standardized Donation Site Data'!E:E)"
                    
                ' Positive or Negative Disbursement
                    wsNegTransactions.Range("DB2").Formula = "=IF(XLOOKUP(AF2,'Disbursements Breakdown'!C:C,'Disbursements Breakdown'!L:L)>0,""Positive"",""Negative"")"
                    

            ' Fill Down
                ' Find the last row of the 'wsNegTransactions' worksheet and store it in a variable called 'NegTransactionsLastRow'.
                    NegTransactionsLastRow = wsNegTransactions.Cells(wsNegTransactions.Rows.Count, "C").End(xlUp).Row
                
                ' Fill the formulas down
                    If NegTransactionsLastRow > 2 Then
                        wsNegTransactions.Range("AA2:DD" & NegTransactionsLastRow).FillDown
                    Else
                        wsNegTransactions.Range("DC:DD").Value = wsNegTransactions.Range("DC:DD").Value
                    End If
                        
            ' Copy and Paste the values only
                If ShowFormulas = False Then
                    wsNegTransactions.Range("A:DD").Value = wsNegTransactions.Range("A:DA").Value
                End If
                
            ' Format the worksheet.
                ' Change the date format.
                    wsNegTransactions.Range("A:B").NumberFormat = "mm/dd/yyyy"

                ' AutoFilter Columns.
                    wsNegTransactions.Range("A1:DD1").AutoFilter
                    
                ' AutoFit Columns.
                    wsNegTransactions.Columns("A:DD").AutoFit
                    
            ' Hide the worksheet
                wsNegTransactions.Visible = xlSheetHidden
                
                

    ' All Possible Fees
        ' Update the status bar
            Application.StatusBar = "Capturing All Fees"
            
        ' Create a new worksheet called "All Possible Fees". Store it in a variable called 'wsAllPossibleFees'.
            Set wsAllPossibleFees = wbMacro.Worksheets.Add(After:=wsNegTransactions)
            
            ' Rename the worksheet to "All Possible Fees"
                wsAllPossibleFees.Name = "All Possible Fees"
                
            ' Create the column headers
                wsAllPossibleFees.Range("A1:F1").Value = Array("Disbursement ID", "Fee Amount", "School Abbreviation", "Depsoit Description", "Fee Type", "Memo")
                'DisbursementsLastRow
                
            ' Add in the data
                ' For Columns A-D: "Disbursement ID", "Fee Amount", "School Abbreviation", "Depsoit Description"
                    wsAllPossibleFees.Range("A2").Formula2 = "=WRAPROWS(TOROW(CHOOSECOLS('Disbursements Breakdown'!A2:AA" & DisbursementsLastRow & ",3,9,22,23,3,10,22,23,3,13,22,23,3,14,22,23)),4)"
                ' "Fee Type"
                    wsAllPossibleFees.Range("E2").Formula2 = "=IF(OR(E1=""Fee Type"",ISNUMBER(SEARCH(""Additional Fees"",E1))),""Transaction Fees""," & _
                            "IF(E1=""Transaction Fees"",""Monthly Fee"",IF(E1=""Monthly Fee"",""Bank Deposit Fees"",IF(ISBLANK(XLOOKUP(A2,'Disbursements Breakdown'!C:C,'Disbursements Breakdown'!O:O))," & _
                            """Additional Fees"",""Additional Fees (""&XLOOKUP(A2,'Disbursements Breakdown'!C:C,'Disbursements Breakdown'!O:O)&"")""))))"
                ' "Memo"
                    wsAllPossibleFees.Range("F2").Formula = "=C2&"" - ""&E2&"" (""&A2&"")"""
                    
            ' Fill Down
                ' Find the last row of the 'wsAllPossibleFees' worksheet and store it in a variable called 'AllPossibleFeesLastRow'.
                    AllPossibleFeesLastRow = wsAllPossibleFees.Cells(wsAllPossibleFees.Rows.Count, "A").End(xlUp).Row
                
                ' Fill the formulas down
                    If AllPossibleFeesLastRow > 2 Then
                        wsAllPossibleFees.Range("E2:F" & AllPossibleFeesLastRow).FillDown
                    End If
                        
            ' Copy and Paste the values only
                If ShowFormulas = False Then
                    wsAllPossibleFees.Range("A:F").Value = wsAllPossibleFees.Range("A:F").Value
                End If
            
            ' Format the worksheet.
                ' AutoFilter Columns.
                    wsAllPossibleFees.Range("A1:F1").AutoFilter
                    
                ' AutoFit Columns.
                    wsAllPossibleFees.Columns("A:F").AutoFit
                    
            ' Hide the worksheet
                wsAllPossibleFees.Visible = xlSheetHidden
                
                
    ' Fees Filtered
        ' Create a new worksheet called "Fees Filtered". Store it in a variable called 'wsFeesFiltered'.
            Set wsFeesFiltered = wbMacro.Worksheets.Add(After:=wsAllPossibleFees)
            
            ' Rename the worksheet to "Fees Filtered"
                wsFeesFiltered.Name = "Fees Filtered"
            
            ' Create the column headers
                wsFeesFiltered.Range("A1:AV1").Value = Array("Intacct - Deposit Description", "Fee Memo Name", "Fee Amount", _
                    "<---Fees Filtered............Adjusting Journal--->", _
                        "Date", "Deposit Description", "Line Number", "Account Number", "Location ID", "Department ID", "Memo", "Amount", "Funding Source", _
                        "GL Entry Class ID (Debt Services)", _
                    "<---Adjusting Journal............CRJ--->", _
                        "RECEIPT_DATE", "PAYMETHOD", "DOCDATE", "DOCNUMBER", "DESCRIPTION", "DEPOSITTO", "BANKACCOUNTID", "DEPOSITDATE", "UNDEPACCTNO", "CURRENCY", _
                        "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCH_RATE_DATE", "LINE_NO", "ACCT_NO", "ACCOUNTLABEL", "TRX_AMOUNT", "AMOUNT", "DEPT_ID", "LOCATION_ID", _
                        "ITEM_MEMO", "OTHERRECEIPTSENTRY_PROJECTID", "OTHERRECEIPTSENTRY_CUSTOMERID", "OTHERRECEIPTSENTRY_ITEMID", "OTHERRECEIPTSENTRY_VENDORID", _
                        "OTHERRECEIPTSENTRY_EMPLOYEEID", "OTHERRECEIPTSENTRY_CLASSID", "PAYER_NAME", "SUPDOCID", "EXCHANGE_RATE", "OR_TRANSACTION_DATE", "GLDIMFUNDING_SOURCE", _
                        "Positive or Negative Disbursement")
            
            ' Add in the data
                ' For Columns A-D: "Intacct - Deposit Description", "Fee Memo Name", "Fee Amount"
                    wsFeesFiltered.Range("A2").Formula2 = "=IFERROR(CHOOSECOLS(FILTER('All Possible Fees'!A2:F" & AllPossibleFeesLastRow & _
                            ",'All Possible Fees'!B2:B" & AllPossibleFeesLastRow & "<>0),4,6,2),""No Results Found"")"
                
                ' "<---Fees Filtered............Adjusting Journal--->"
                    wsFeesFiltered.Range("D2").Value = "=""<---Fees Filtered............Adjusting Journal--->"""
                    
                ' "Date"
                    wsFeesFiltered.Range("E2").Formula = "=XLOOKUP(A2,'Disbursements Breakdown'!W:W,'Disbursements Breakdown'!D:D)"
                    
                ' "Deposit Description"
                    wsFeesFiltered.Range("F2").Formula = "=A2"
                    
                ' "Line Number"
                    wsFeesFiltered.Range("G2").Formula = "="""""
                    
                ' "Account Number"
                    wsFeesFiltered.Range("H2").Formula = "=""82401"""
                    
                ' "Location ID"
                    wsFeesFiltered.Range("I2").Formula = "=XLOOKUP(F2,'Disbursements Breakdown'!W:W,'Disbursements Breakdown'!Y:Y)"
                    
                ' "Department ID"
                    wsFeesFiltered.Range("J2").Formula = "=""2046"""
                    
                ' "Memo"
                    wsFeesFiltered.Range("K2").Formula = "=B2"
                    
                ' "Amount"
                    wsFeesFiltered.Range("L2").Formula = "=C2"
                    
                ' "Funding Source"
                    wsFeesFiltered.Range("M2").Formula = "=""7301-ATF Campaign"""
                    
                ' "GL Entry Class ID (Debt Services)"
                    If ShowFormulas = True Then
                        wsFeesFiltered.Range("N2").Formula = "=""000"""
                    Else
                        wsFeesFiltered.Range("N2").Formula = "=""'000"""
                    End If
                    
                ' "<---Adjusting Journal............CRJ--->"
                    wsFeesFiltered.Range("O2").Value = "=""<---Adjusting Journal............CRJ--->"""
                    
                ' "RECEIPT_DATE"
                    wsFeesFiltered.Range("P2").Formula = "=E2"
                    
                ' "PAYMETHOD"
                    wsFeesFiltered.Range("Q2").Formula = "=""Credit Card"""
                    
                ' "DOCDATE"
                    wsFeesFiltered.Range("R2").Formula = "=P2"
                    
                ' "DOCNUMBER"
                    wsFeesFiltered.Range("S2").Formula = "=XLOOKUP(A2,'Disbursements Breakdown'!W:W,'Disbursements Breakdown'!B:B)"
                    
                ' "DESCRIPTION"
                    wsFeesFiltered.Range("T2").Formula = "=LEFT(F2,SEARCH(""["",F2)-2)"
                    
                ' "DEPOSITTO"
                    wsFeesFiltered.Range("U2").Formula = "=""Bank account"""
                    
                ' "BANKACCOUNTID"
                    wsFeesFiltered.Range("V2").Formula2 = "=ConvertBankAccountToBankAccountName(ConvertCnPToBankAccount(MID(A2,SEARCH(""("",A2)+1,5)))"
                    
                ' "DEPOSITDATE"
                    wsFeesFiltered.Range("W2").Formula = "=P2"
                    
                ' "UNDEPACCTNO"
                    wsFeesFiltered.Range("X2").Formula = "="""""
                    
                ' "CURRENCY"
                    wsFeesFiltered.Range("Y2").Formula = "=""USD"""
                    
                ' "EXCH_RATE_DATE"
                    wsFeesFiltered.Range("Z2").Formula = "="""""
                    
                ' "EXCH_RATE_TYPE_ID"
                    wsFeesFiltered.Range("AA2").Formula = "="""""
                    
                ' "EXCH_RATE_DATE"
                    wsFeesFiltered.Range("AB2").Formula = "="""""
                    
                ' "LINE_NO"
                    wsFeesFiltered.Range("AC2").Formula = "="""""
                    
                ' "ACCT_NO"
                    wsFeesFiltered.Range("AD2").Formula = "=H2"
                    
                ' "ACCOUNTLABEL"
                    wsFeesFiltered.Range("AE2").Formula = "="""""
                    
                ' "TRX_AMOUNT"
                    wsFeesFiltered.Range("AF2").Formula = "=L2"
                    
                ' "AMOUNT"
                    wsFeesFiltered.Range("AG2").Formula = "=L2"
                    
                ' "DEPT_ID"
                    wsFeesFiltered.Range("AH2").Formula = "=J2"
                    
                ' "LOCATION_ID"
                    wsFeesFiltered.Range("AI2").Formula = "=I2"
                    
                ' "ITEM_MEMO"
                    wsFeesFiltered.Range("AJ2").Formula = "=K2"
                    
                ' "OTHERRECEIPTSENTRY_PROJECTID"
                    wsFeesFiltered.Range("AK2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_CUSTOMERID"
                    wsFeesFiltered.Range("AL2").Formula = "=IF(S2=""ProPay"",""19234"",""8179"")"
                    
                ' "OTHERRECEIPTSENTRY_ITEMID"
                    wsFeesFiltered.Range("AM2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_VENDORID"
                    wsFeesFiltered.Range("AN2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_EMPLOYEEID"
                    wsFeesFiltered.Range("AO2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_CLASSID"
                    wsFeesFiltered.Range("AP2").Formula = "=N2"
                    
                ' "PAYER_NAME"
                    wsFeesFiltered.Range("AQ2").Formula = "=S2"
                    
                ' "SUPDOCID"
                    wsFeesFiltered.Range("AR2").Formula = "="""""
                    
                ' "EXCHANGE_RATE"
                    wsFeesFiltered.Range("AS2").Formula = "="""""
                    
                ' "OR_TRANSACTION_DATE"
                    wsFeesFiltered.Range("AT2").Formula = "=P2"
                    
                ' "GLDIMFUNDING_SOURCE"
                    wsFeesFiltered.Range("AU2").Formula = "=M2"
                
                ' "Positive or Negative Disbursement"
                    wsFeesFiltered.Range("AV2").Formula = "=IF(ISNUMBER(SEARCH(""[-$"",A2)),""Negative"",""Positive"")"
                       
            ' Fill Down
                ' Find the last row of the 'wsFeesFiltered' worksheet and store it in a variable called 'FeesFilteredLastRow'.
                    FeesFilteredLastRow = wsFeesFiltered.Cells(wsFeesFiltered.Rows.Count, "A").End(xlUp).Row
                
                ' Fill the formulas down
                    If FeesFilteredLastRow > 2 Then
                        wsFeesFiltered.Range("D2:AV" & FeesFilteredLastRow).FillDown
                    End If
                        
            ' Copy and Paste the values only
                If ShowFormulas = False Then
                    wsFeesFiltered.Range("A:AV").Value = wsFeesFiltered.Range("A:AU").Value
                End If
            
            ' Format the worksheet.
                ' AutoFilter Columns.
                    wsFeesFiltered.Range("A1:AV1").AutoFilter
                    
                ' AutoFit Columns.
                    wsFeesFiltered.Columns("A:AV").AutoFit
                    
            ' Hide the worksheet
                wsFeesFiltered.Visible = xlSheetHidden

    
    ' Bank Disbursement Amounts
        ' Update the status bar
            Application.StatusBar = "Breaking Out All Bank Disbursements"
            
        ' Create a new worksheet called "Bank Disbursement Amounts". Store it in a variable called 'wsBankDisbursementAmounts'.
            Set wsBankDisbursementAmounts = wbMacro.Worksheets.Add(After:=wsFeesFiltered)
            
            ' Rename the worksheet to "Bank Disbursement Amounts"
                wsBankDisbursementAmounts.Name = "Bank Disbursement Amounts"
                
            ' Create the column headers
                wsBankDisbursementAmounts.Range("A1:K1").Value = Array("Date", "Deposit Description", "Line Number", "Account Number", "Location ID", "Department ID", _
                        "Memo", "Amount", "Funding Source", "GL Entry Class ID (Debt Services)", "Positive or Negative Disbursement")
            
            ' Add in the data
                ' "Date"
                    wsBankDisbursementAmounts.Range("A2").Formula = "=TEXT(XLOOKUP(B2,'Disbursements Breakdown'!W:W,'Disbursements Breakdown'!D:D),""MM/DD/YYYY"")"
                    
                ' "Deposit Description"
                    wsBankDisbursementAmounts.Range("B2").Formula2 = "='Disbursements Breakdown'!W2"
                    
                ' "Line Number"
                    wsBankDisbursementAmounts.Range("C2").Formula = "="""""
                    
                ' "Account Number"
                    wsBankDisbursementAmounts.Range("D2").Formula = "=XLOOKUP(B2,'Disbursements Breakdown'!W:W,'Disbursements Breakdown'!X:X)"
                    
                ' "Location ID"
                    wsBankDisbursementAmounts.Range("E2").Formula = "='Disbursements Breakdown'!Y2"
                    
                ' "Department ID"
                    wsBankDisbursementAmounts.Range("F2").Formula = "=""2048"""
                    
                ' "Memo"
                    wsBankDisbursementAmounts.Range("G2").Formula = "=""Bank Deposit - ""&'Disbursements Breakdown'!V2&"" (""&'Disbursements Breakdown'!C2&"")"""
                    
                ' "Amount"
                    wsBankDisbursementAmounts.Range("H2").Formula = "='Disbursements Breakdown'!P2"
                    
                ' "Funding Source"
                    wsBankDisbursementAmounts.Range("I2").Formula = "=""7301-ATF Campaign"""
                    
                ' "GL Entry Class ID (Debt Services)"
                    If ShowFormulas = True Then
                        wsBankDisbursementAmounts.Range("J2").Formula = "=""000"""
                    Else
                        wsBankDisbursementAmounts.Range("J2").Formula = "=""'000"""
                    End If
                    
                ' "Positive or Negative Disbursement"
                    wsBankDisbursementAmounts.Range("K2").Formula = "=IF(ISNUMBER(SEARCH(""[-$"",B2)),""Negative"",""Positive"")"
                    
            ' Fill Down
                ' Fill the formulas down
                    If DisbursementsLastRow > 2 Then
                        wsBankDisbursementAmounts.Range("A2:K" & DisbursementsLastRow).FillDown
                    End If
                        
            ' Copy and Paste the values only
                If ShowFormulas = False Then
                    wsBankDisbursementAmounts.Range("A:K").Value = wsBankDisbursementAmounts.Range("A:K").Value
                End If
            
            ' Format the worksheet.
                ' AutoFilter Columns.
                    wsBankDisbursementAmounts.Range("A1:K1").AutoFilter
                    
                ' AutoFit Columns.
                    wsBankDisbursementAmounts.Columns("A:K").AutoFit

            ' Find the last row of the 'wsBankDisbursementAmounts' worksheet and store it in a variable called 'BankDisbursementAmountsLastRow'.
                BankDisbursementAmountsLastRow = wsBankDisbursementAmounts.Cells(wsBankDisbursementAmounts.Rows.Count, "B").End(xlUp).Row
                
            ' Hide the worksheet
                wsBankDisbursementAmounts.Visible = xlSheetHidden
                
    
    ' Determine the route for creating the Intacct import files based on the 'ReportRoute' - store the 'ImportType' route.
        If ReportRoute = "Salesforce" Then
            ImportType = "CRJ"
        ElseIf ReportRoute = "Intacct" Then
            ImportType = "Adjusting"
        End If
    
        ' Based on the 'ImportType' determine the route for creating the import file(s).
            If ImportType = "CRJ" Then
                GoTo CRJRoute
            ElseIf ImportType = "Adjusting" Then
                GoTo AdjustingRoute
            End If



'--------------------------------------------
CRJRoute:
'--------------------------------------------
Application.StatusBar = "Creating the Intacct Import Files."
' All Data Combined (CRJ)
    ' Create a new worksheet called "All Data Combined (+)". Store it in a variable called 'wsAllDataCombinedPos'.
        Set wsAllDataCombinedPos = wbMacro.Worksheets.Add(After:=wsBankDisbursementAmounts)
    ' Rename the worksheet
        wsAllDataCombinedPos.Name = "All Data Combined (+)"

    ' Create the column headers
        wsAllDataCombinedPos.Range("A1:AH1").Value = Array("DONOTIMPORT", "DONOTIMPORT", "RECEIPT_DATE", "PAYMETHOD", "DOCDATE", "DOCNUMBER", _
                "DESCRIPTION", "DEPOSITTO", "BANKACCOUNTID", "DEPOSITDATE", "UNDEPACCTNO", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCH_RATE_DATE", "LINE_NO", _
                "ACCT_NO", "ACCOUNTLABEL", "TRX_AMOUNT", "AMOUNT", "DEPT_ID", "LOCATION_ID", "ITEM_MEMO", "OTHERRECEIPTSENTRY_PROJECTID", "OTHERRECEIPTSENTRY_CUSTOMERID", _
                "OTHERRECEIPTSENTRY_ITEMID", "OTHERRECEIPTSENTRY_VENDORID", "OTHERRECEIPTSENTRY_EMPLOYEEID", "OTHERRECEIPTSENTRY_CLASSID", "PAYER_NAME", "SUPDOCID", _
                "EXCHANGE_RATE", "OR_TRANSACTION_DATE", "GLDIMFUNDING_SOURCE")
    
    ' Add in the data
        wsAllDataCombinedPos.Range("C2").Formula2 = "=LET(PosCount,COUNTIF('Positive Transactions'!DB:DB,""Positive"")," & _
                "NegCount,COUNTIF('Negative Transactions'!DB:DB,""Positive"")," & _
                "FeesCount,COUNTIF('Fees Filtered'!AV:AV,""Positive"")," & _
            "PosArray,FILTER('Positive Transactions'!BV2:DA" & PosTransactionsLastRow & ",'Positive Transactions'!DB2:DB" & PosTransactionsLastRow & "=""Positive"")," & _
                "NegArray,FILTER('Negative Transactions'!BV2:DA" & NegTransactionsLastRow & ",'Negative Transactions'!DB2:DB" & NegTransactionsLastRow & "=""Positive"")," & _
                "FeesArray,FILTER('Fees Filtered'!P2:AU" & FeesFilteredLastRow & ",'Fees Filtered'!AV2:AV" & FeesFilteredLastRow & "=""Positive"")," & _
            "IF(AND(PosCount>0,NegCount>0,FeesCount>0),SORT(VSTACK(FeesArray,PosArray,NegArray),5)," & _
                "IF(AND(PosCount=0,NegCount>0,FeesCount>0),SORT(VSTACK(FeesArray,NegArray),5)," & _
                "IF(AND(PosCount>0,NegCount=0,FeesCount>0),SORT(VSTACK(FeesArray,PosArray),5)," & _
                "IF(AND(PosCount>0,NegCount>0,FeesCount=0),SORT(VSTACK(PosArray,NegArray),5)," & _
                "IF(AND(PosCount>0,NegCount=0,FeesCount=0),SORT(PosArray,5)," & _
                "IF(AND(PosCount=0,NegCount>0,FeesCount=0),SORT(NegArray,5)," & _
                "IF(AND(PosCount=0,NegCount=0,FeesCount>0),SORT(FeesArray,5)," & _
            """No Positive Disbursements Found""))))))))"
                
    
    ' Find the last row of the 'wsAllDataCombinedPos' worksheet and store it in a variable called 'AllDataCombinedPosLastRow'.
        AllDataCombinedPosLastRow = wsAllDataCombinedPos.Cells(wsAllDataCombinedPos.Rows.Count, "C").End(xlUp).Row
        
    ' Copy and Paste the values only
        If ShowFormulas = False Then
            wsAllDataCombinedPos.Range("A:AH").Value = wsAllDataCombinedPos.Range("A:AH").Value
        End If
    
    ' Format the worksheet.
        ' AutoFilter Columns.
            wsAllDataCombinedPos.Range("A1:AH1").AutoFilter
        
        ' AutoFit Columns.
            wsAllDataCombinedPos.Columns("A:AH").AutoFit
            
    ' Hide the worksheet
        wsAllDataCombinedPos.Visible = xlSheetHidden
            
            

' All Data Combined (Adjusting)
    ' Create a worksheet called "All Data Combined (-)". Store it in a variable called 'wsAllDataCombinedNeg'.
        Set wsAllDataCombinedNeg = wbMacro.Worksheets.Add(After:=wsAllDataCombinedPos)
    
    ' Rename the worksheet.
        wsAllDataCombinedNeg.Name = "All Data Combined (-)"
        
    ' Create the column headers
        wsAllDataCombinedNeg.Range("A1:K1").Value = Array("Journal", "Date", "Deposit Description", "Line Number", "Account Number", "Location ID", "Department ID", "Memo", _
                "Amount", "Funding Source", "GL Entry Class ID (Debt Services)")
        
    ' Add in the data
        ' Columns B-K: "Date", "Deposit Description", "Line Number", "Account Number", "Location ID", "Department ID", "Memo", "Amount", "Funding Source", "GL Entry Class ID (Debt Services)"
            wsAllDataCombinedNeg.Range("B2").Formula2 = "=LET(PosCount,COUNTIF('Positive Transactions'!DB:DB,""Negative"")," & _
                    "NegCount,COUNTIF('Negative Transactions'!DB:DB,""Negative"")," & _
                    "FeesCount,COUNTIF('Fees Filtered'!AV:AV,""Negative"")," & _
                "PosArray,FILTER('Positive Transactions'!BK2:BT" & PosTransactionsLastRow & ",'Positive Transactions'!DB2:DB" & PosTransactionsLastRow & "=""Negative"")," & _
                    "NegArray,FILTER('Negative Transactions'!BK2:BT" & NegTransactionsLastRow & ",'Negative Transactions'!DB2:DB" & NegTransactionsLastRow & "=""Negative"")," & _
                    "FeesArray,FILTER('Fees Filtered'!E2:N" & FeesFilteredLastRow & ",'Fees Filtered'!AV2:AV" & FeesFilteredLastRow & "=""Negative"")," & _
                    "BankArray,FILTER('Bank Disbursement Amounts'!A2:J" & BankDisbursementAmountsLastRow & ",'Bank Disbursement Amounts'!K2:K" & BankDisbursementAmountsLastRow & "=""Negative"")," & _
                "IF(AND(PosCount>0,NegCount>0,FeesCount>0),SORT(VSTACK(BankArray, FeesArray, PosArray, NegArray),2)," & _
                    "IF(AND(PosCount=0,NegCount>0,FeesCount>0),SORT(VSTACK(BankArray, FeesArray, NegArray),2)," & _
                    "IF(AND(PosCount>0,NegCount=0,FeesCount>0),SORT(VSTACK(BankArray, FeesArray, PosArray),2)," & _
                    "IF(AND(PosCount>0,NegCount>0,FeesCount=0),SORT(VSTACK(BankArray, PosArray, NegArray),2)," & _
                    "IF(AND(PosCount>0,NegCount=0,FeesCount=0),SORT(VSTACK(BankArray, PosArray),2)," & _
                    "IF(AND(PosCount=0,NegCount>0,FeesCount=0),SORT(VSTACK(BankArray, NegArray),2)," & _
                    "IF(AND(PosCount=0,NegCount=0,FeesCount>0),SORT(VSTACK(BankArray, FeesArray),2)," & _
                """No Negative Disbursements Found""))))))))"
            
        ' Column A: "Journal"
            If wsAllDataCombinedNeg.Range("B2").Value <> "No Negative Disbursements Found" Then
                wsAllDataCombinedNeg.Range("A2").Value = JournalType
            End If
            
    ' Find the last row of the 'wsAllDataCombinedNeg' worksheet and store it in a variable called 'AllDataCombinedNegLastRow'.
        AllDataCombinedNegLastRow = wsAllDataCombinedNeg.Cells(wsAllDataCombinedNeg.Rows.Count, "B").End(xlUp).Row
        
    ' Fill Down
        If AllDataCombinedNegLastRow > 2 Then
            wsAllDataCombinedNeg.Range("A2:A" & AllDataCombinedNegLastRow).FillDown
        End If
        
    ' Copy and Paste the values only
        If ShowFormulas = False Then
            wsAllDataCombinedNeg.Range("A:K").Value = wsAllDataCombinedNeg.Range("A:K").Value
        End If
    
    ' Format the worksheet.
        ' AutoFilter Columns.
            wsAllDataCombinedNeg.Range("A1:K1").AutoFilter
            
        ' AutoFit Columns.
            wsAllDataCombinedNeg.Columns("A:K").AutoFit

    ' Hide the worksheet
        wsAllDataCombinedNeg.Visible = xlSheetHidden
        
        
' Import File (CRJ)
    ' Create the worksheet
        Set wsImportCRJ = wbMacro.Worksheets.Add(After:=wsAllDataCombinedNeg)
        wsImportCRJ.Name = "Import File (CRJ)"
        
    
    ' Add the column headers
        wsImportCRJ.Range("A1:AG1").Value = Array("DONOTIMPORT", "RECEIPT_DATE", "PAYMETHOD", "DOCDATE", "DOCNUMBER", "DESCRIPTION", "DEPOSITTO", _
                "BANKACCOUNTID", "DEPOSITDATE", "UNDEPACCTNO", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCH_RATE_DATE", "LINE_NO", _
                "ACCT_NO", "ACCOUNTLABEL", "TRX_AMOUNT", "AMOUNT", "DEPT_ID", "LOCATION_ID", "ITEM_MEMO", "OTHERRECEIPTSENTRY_PROJECTID", "OTHERRECEIPTSENTRY_CUSTOMERID", _
                "OTHERRECEIPTSENTRY_ITEMID", "OTHERRECEIPTSENTRY_VENDORID", "OTHERRECEIPTSENTRY_EMPLOYEEID", "OTHERRECEIPTSENTRY_CLASSID", "PAYER_NAME", "SUPDOCID", _
                "EXCHANGE_RATE", "OR_TRANSACTION_DATE", "GLDIMFUNDING_SOURCE")
    
    ' Add in the data
        If wsAllDataCombinedPos.Range("C2").Value <> "No Positive Disbursements Found" Then
            ' "RECEIPT_DATE"
                wsImportCRJ.Range("B2").Formula = "='All Data Combined (+)'!C2"
                
            ' "PAYMETHOD"
                wsImportCRJ.Range("C2").Formula = "='All Data Combined (+)'!D2"
            
            ' "DOCDATE"
                wsImportCRJ.Range("D2").Formula = "='All Data Combined (+)'!E2"
                
            ' "DOCNUMBER"
                wsImportCRJ.Range("E2").Formula = "='All Data Combined (+)'!F2"
                
            ' "DESCRIPTION"
                wsImportCRJ.Range("F2").Formula = "='All Data Combined (+)'!G2"
                
            ' "DEPOSITTO"
                wsImportCRJ.Range("G2").Formula = "='All Data Combined (+)'!H2"
                
            ' "BANKACCOUNTID"
                wsImportCRJ.Range("H2").Formula = "='All Data Combined (+)'!I2"
                
            ' "DEPOSITDATE"
                wsImportCRJ.Range("I2").Formula = "='All Data Combined (+)'!J2"
                
            ' "UNDEPACCTNO"
                wsImportCRJ.Range("J2").Formula = "='All Data Combined (+)'!K2"
                
            ' "CURRENCY"
                wsImportCRJ.Range("K2").Formula = "='All Data Combined (+)'!L2"
                
            ' "EXCH_RATE_DATE"
                wsImportCRJ.Range("L2").Formula = "='All Data Combined (+)'!M2"
                
            ' "EXCH_RATE_TYPE_ID"
                wsImportCRJ.Range("M2").Formula = "='All Data Combined (+)'!N2"
                
            ' "EXCH_RATE_DATE"
                wsImportCRJ.Range("N2").Formula = "='All Data Combined (+)'!O2"
                
            ' "LINE_NO"
                wsImportCRJ.Range("O2").Formula = "=IF(F2=F1,O1+1,1)"
                
            ' "ACCT_NO"
                wsImportCRJ.Range("P2").Formula = "='All Data Combined (+)'!Q2"
                
            ' "ACCOUNTLABEL"
                wsImportCRJ.Range("Q2").Formula = "='All Data Combined (+)'!R2"
                
            ' "TRX_AMOUNT"
                wsImportCRJ.Range("R2").Formula = "='All Data Combined (+)'!S2"
                
            ' "AMOUNT"
                wsImportCRJ.Range("S2").Formula = "='All Data Combined (+)'!T2"
                
            ' "DEPT_ID"
                wsImportCRJ.Range("T2").Formula = "='All Data Combined (+)'!U2"
                
            ' "LOCATION_ID"
                wsImportCRJ.Range("U2").Formula = "='All Data Combined (+)'!V2"
                
            ' "ITEM_MEMO"
                wsImportCRJ.Range("V2").Formula = "='All Data Combined (+)'!W2"
                
            ' "OTHERRECEIPTSENTRY_PROJECTID"
                wsImportCRJ.Range("W2").Formula = "='All Data Combined (+)'!X2"
                
            ' "OTHERRECEIPTSENTRY_CUSTOMERID"
                wsImportCRJ.Range("X2").Formula = "='All Data Combined (+)'!Y2"
                
            ' "OTHERRECEIPTSENTRY_ITEMID"
                wsImportCRJ.Range("Y2").Formula = "='All Data Combined (+)'!Z2"
                
            '  "OTHERRECEIPTSENTRY_VENDORID"
                wsImportCRJ.Range("Z2").Formula = "='All Data Combined (+)'!AA2"
                
            ' "OTHERRECEIPTSENTRY_EMPLOYEEID"
                wsImportCRJ.Range("AA2").Formula = "='All Data Combined (+)'!AB2"
                
            ' "OTHERRECEIPTSENTRY_CLASSID"
                wsImportCRJ.Range("AB2").Formula = "='All Data Combined (+)'!AC2"
                
            ' "PAYER_NAME"
                wsImportCRJ.Range("AC2").Formula = "='All Data Combined (+)'!AD2"
                
            ' "SUPDOCID"
                wsImportCRJ.Range("AD2").Formula = "='All Data Combined (+)'!AE2"
                
            ' "EXCHANGE_RATE"
                wsImportCRJ.Range("AE2").Formula = "='All Data Combined (+)'!AF2"
                
            '  "OR_TRANSACTION_DATE"
                wsImportCRJ.Range("AF2").Formula = "='All Data Combined (+)'!AG2"
                
            ' "GLDIMFUNDING_SOURCE"
                wsImportCRJ.Range("AG2").Formula = "='All Data Combined (+)'!AH2"
            
            
            ' Fill the data down
                If AllDataCombinedPosLastRow > 2 Then
                    wsImportCRJ.Range("B2:AG" & AllDataCombinedPosLastRow).FillDown
                End If
                
            ' Format the data
                ' AutoFilter Columns.
                    wsImportCRJ.Range("A1:AG1").AutoFilter
                    
                ' AutoFit Columns.
                    wsImportCRJ.Columns("A:AG").AutoFit

        End If
        
'' Import File (Adjusting)
    ' Create the worksheet
        ' wsImport & JournalType
            Set wsImportAdjusting = wbMacro.Worksheets.Add(After:=wsImportCRJ)
            
            wsImportAdjusting.Name = "Import File (Adjusting)"
        
    
    ' Add the column headers
        wsImportAdjusting.Range("A1:AG1").Value = Array("DONOTIMPORT", "JOURNAL", "DATE", "REVERSEDATE", "DESCRIPTION", "REFERENCE_NO", "LINE_NO", "ACCT_NO", "LOCATION_ID", _
                "DEPT_ID", "DOCUMENT", "MEMO", "DEBIT", "CREDIT", "SOURCEENTITY", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCHANGE_RATE", "STATE", "ALLOCATION_ID", _
                "RASSET", "RDEPRECIATION_SCHEDULE", "RASSET_ADJUSTMENT", "RASSET_CLASS", "RASSETOUTOFSERVICE", "GLDIMFUNDING_SOURCE", "GLENTRY_PROJECTID", "GLENTRY_CUSTOMERID", _
                "GLENTRY_VENDORID", "GLENTRY_EMPLOYEEID", "GLENTRY_ITEMID", "GLENTRY_CLASSID")

    ' Add in the data
        If wsAllDataCombinedNeg.Range("B2").Value <> "No Negative Disbursements Found" Then
            ' "JOURNAL"
                wsImportAdjusting.Range("B2").Formula = "=IF(E2="""","""",'All Data Combined (-)'!A2)"
                
            ' "DATE"
                wsImportAdjusting.Range("C2").Formula = "=IF(E2="""","""",'All Data Combined (-)'!B2)"
                
            ' "REVERSEDATE"
                wsImportAdjusting.Range("D2").Formula = "="""""
                
            ' "DESCRIPTION"
                wsImportAdjusting.Range("E2").Formula = "=IF('All Data Combined (-)'!C2='All Data Combined (-)'!C1,"""",'All Data Combined (-)'!C2)"
                
            ' "REFERENCE_NO"
                wsImportAdjusting.Range("F2").Formula = "="""""
                
            ' "LINE_NO"
                wsImportAdjusting.Range("G2").Formula = "=IF('All Data Combined (-)'!C2='All Data Combined (-)'!C1,1+G1,1)"
                
            ' "ACCT_NO"
                wsImportAdjusting.Range("H2").Formula = "='All Data Combined (-)'!E2"
                
            ' "LOCATION_ID"
                wsImportAdjusting.Range("I2").Formula = "='All Data Combined (-)'!F2"
                
            ' "DEPT_ID"
                wsImportAdjusting.Range("J2").Formula = "='All Data Combined (-)'!G2"
                
            ' "DOCUMENT"
                wsImportAdjusting.Range("K2").Formula = "="""""
                
            ' "MEMO"
                wsImportAdjusting.Range("L2").Formula = "='All Data Combined (-)'!H2"
                
            ' "DEBIT"
                wsImportAdjusting.Range("M2").Formula = "=IF(OR(H2=""11100"",H2=""11200"",H2=""11400"",H2=""11700"",G2=1),"""",IF('All Data Combined (-)'!I2<0,ABS('All Data Combined (-)'!I2),""""))"
                
            ' "CREDIT"
                wsImportAdjusting.Range("N2").Formula = "=IF(OR(H2=""11100"",H2=""11200"",H2=""11400"",H2=""11700"",G2=1),ABS('All Data Combined (-)'!I2)," & _
                        "IF('All Data Combined (-)'!I2>0,'All Data Combined (-)'!I2,""""))"
                        
            ' "SOURCEENTITY"
                wsImportAdjusting.Range("O2").Formula = "="""""
                
            ' "CURRENCY"
                wsImportAdjusting.Range("P2").Formula = "="""""
                
            ' "EXCH_RATE_DATE"
                wsImportAdjusting.Range("Q2").Formula = "="""""
                
            ' "EXCH_RATE_TYPE_ID"
                wsImportAdjusting.Range("R2").Formula = "="""""
                
            '  "EXCHANGE_RATE"
                wsImportAdjusting.Range("S2").Formula = "="""""
                
            '  "STATE"
                wsImportAdjusting.Range("T2").Formula = "=""Draft"""
                
            ' "ALLOCATION_ID"
                wsImportAdjusting.Range("U2").Formula = "="""""
                
            ' "RASSET"
                wsImportAdjusting.Range("V2").Formula = "="""""
                
            '  "RDEPRECIATION_SCHEDULE"
                wsImportAdjusting.Range("W2").Formula = "="""""
                
            '  "RASSET_ADJUSTMENT"
                wsImportAdjusting.Range("X2").Formula = "="""""
                
            '  "RASSET_CLASS"
                wsImportAdjusting.Range("Y2").Formula = "="""""
                
            '  "RASSETOUTOFSERVICE"
                wsImportAdjusting.Range("Z2").Formula = "="""""
                
            '  "GLDIMFUNDING_SOURCE"
                wsImportAdjusting.Range("AA2").Formula = "='All Data Combined (-)'!J2"
                
            '  "GLENTRY_PROJECTID"
                wsImportAdjusting.Range("AB2").Formula = "="""""
                
            '  "GLENTRY_CUSTOMERID"
                wsImportAdjusting.Range("AC2").Formula = "="""""
                
            ' "GLENTRY_VENDORID"
                wsImportAdjusting.Range("AD2").Formula = "="""""
                
            ' "GLENTRY_EMPLOYEEID"
                wsImportAdjusting.Range("AE2").Formula = "="""""
                
            ' "GLENTRY_ITEMID"
                wsImportAdjusting.Range("AF2").Formula = "="""""
                
            ' "GLENTRY_CLASSID"
                wsImportAdjusting.Range("AG2").Formula = "='All Data Combined (-)'!K2"

            ' Fill the data down
                If AllDataCombinedNegLastRow > 2 Then
                    wsImportAdjusting.Range("B2:AG" & AllDataCombinedNegLastRow).FillDown
                End If
                
            ' Format the data
                ' AutoFilter Columns.
                    wsImportAdjusting.Range("A1:AG1").AutoFilter
                    
                ' AutoFit Columns.
                    wsImportAdjusting.Columns("A:AG").AutoFit
                
        End If

    ' Finish by jumping over 'AdjustingRoute:' and going to 'CreateChecks:'
        GoTo CreateChecks
        
'--------------------------------------------
AdjustingRoute:
'--------------------------------------------
' Update the status bar
    Application.StatusBar = "Creating the Intacct Import Files."
    
' All Data Combined (Adjusting)
    ' Create a worksheet called "All Data Combined (-)". Store it in a variable called 'wsAllDataCombined'.
        Set wsAllDataCombined = wbMacro.Worksheets.Add(After:=wsBankDisbursementAmounts)
    
    ' Rename the worksheet.
        wsAllDataCombined.Name = "All Data Combined"
        
    ' Create the column headers
        wsAllDataCombined.Range("A1:K1").Value = Array("Journal", "Date", "Deposit Description", "Line Number", "Account Number", "Location ID", "Department ID", "Memo", _
                "Amount", "Funding Source", "GL Entry Class ID (Debt Services)")
        
    ' Add in the data
        ' Columns B-K: "Date", "Deposit Description", "Line Number", "Account Number", "Location ID", "Department ID", "Memo", "Amount", "Funding Source", "GL Entry Class ID (Debt Services)"
            wsAllDataCombined.Range("B2").Formula2 = "=LET(PosCount,COUNTIF('Positive Transactions'!DB:DB,""Negative"") + COUNTIF('Positive Transactions'!DB:DB,""Positive"")," & _
                    "NegCount,COUNTIF('Negative Transactions'!DB:DB,""Negative"") + COUNTIF('Negative Transactions'!DB:DB,""Positive"")," & _
                    "FeesCount,COUNTIF('Fees Filtered'!AV:AV,""Negative"") + COUNTIF('Fees Filtered'!AV:AV,""Positive"")," & _
                "PosArray,FILTER('Positive Transactions'!BK2:BT" & PosTransactionsLastRow & ",'Positive Transactions'!DB2:DB" & PosTransactionsLastRow & "<>"""")," & _
                    "NegArray,FILTER('Negative Transactions'!BK2:BT" & NegTransactionsLastRow & ",'Negative Transactions'!DB2:DB" & NegTransactionsLastRow & "<>"""")," & _
                    "FeesArray,FILTER('Fees Filtered'!E2:N" & FeesFilteredLastRow & ",'Fees Filtered'!AV2:AV" & FeesFilteredLastRow & "<>"""")," & _
                    "BankArray,FILTER('Bank Disbursement Amounts'!A2:J" & BankDisbursementAmountsLastRow & ",'Bank Disbursement Amounts'!K2:K" & BankDisbursementAmountsLastRow & "<>"""")," & _
                "IF(AND(PosCount>0,NegCount>0,FeesCount>0),SORT(VSTACK(BankArray, FeesArray, PosArray, NegArray),2)," & _
                    "IF(AND(PosCount=0,NegCount>0,FeesCount>0),SORT(VSTACK(BankArray, FeesArray, NegArray),2)," & _
                    "IF(AND(PosCount>0,NegCount=0,FeesCount>0),SORT(VSTACK(BankArray, FeesArray, PosArray),2)," & _
                    "IF(AND(PosCount>0,NegCount>0,FeesCount=0),SORT(VSTACK(BankArray, PosArray, NegArray),2)," & _
                    "IF(AND(PosCount>0,NegCount=0,FeesCount=0),SORT(VSTACK(BankArray, PosArray),2)," & _
                    "IF(AND(PosCount=0,NegCount>0,FeesCount=0),SORT(VSTACK(BankArray, NegArray),2)," & _
                    "IF(AND(PosCount=0,NegCount=0,FeesCount>0),SORT(VSTACK(BankArray, FeesArray),2)," & _
                """No Transactions Found""))))))))"
            
        ' Column A: "Journal"
            If wsAllDataCombined.Range("B2").Value <> "No Transactions Found" Then
                wsAllDataCombined.Range("A2").Value = JournalType
            End If
            
    ' Find the last row of the 'wsAllDataCombined' worksheet and store it in a variable called 'AllDataCombinedLastRow'.
        AllDataCombinedLastRow = wsAllDataCombined.Cells(wsAllDataCombined.Rows.Count, "B").End(xlUp).Row
        
    ' Fill Down
        If AllDataCombinedLastRow > 2 Then
            wsAllDataCombined.Range("A2:A" & AllDataCombinedLastRow).FillDown
        End If
        
    ' Copy and Paste the values only
        If ShowFormulas = False Then
            wsAllDataCombined.Range("A:K").Value = wsAllDataCombined.Range("A:K").Value
        End If
    
    ' Format the worksheet.
        ' AutoFilter Columns.
            wsAllDataCombined.Range("A1:K1").AutoFilter
            
        ' AutoFit Columns.
            wsAllDataCombined.Columns("A:K").AutoFit
            
    ' Hide the worksheet
        wsAllDataCombined.Visible = xlSheetHidden

'' Import File (Adjusting)
    ' Create the worksheet
        ' wsImport & JournalType
            Set wsImport = wbMacro.Worksheets.Add(After:=wsAllDataCombined)
        
            wsImport.Name = "Import File"
        
    
    ' Add the column headers
        wsImport.Range("A1:AG1").Value = Array("DONOTIMPORT", "JOURNAL", "DATE", "REVERSEDATE", "DESCRIPTION", "REFERENCE_NO", "LINE_NO", "ACCT_NO", "LOCATION_ID", _
                "DEPT_ID", "DOCUMENT", "MEMO", "DEBIT", "CREDIT", "SOURCEENTITY", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCHANGE_RATE", "STATE", "ALLOCATION_ID", _
                "RASSET", "RDEPRECIATION_SCHEDULE", "RASSET_ADJUSTMENT", "RASSET_CLASS", "RASSETOUTOFSERVICE", "GLDIMFUNDING_SOURCE", "GLENTRY_PROJECTID", "GLENTRY_CUSTOMERID", _
                "GLENTRY_VENDORID", "GLENTRY_EMPLOYEEID", "GLENTRY_ITEMID", "GLENTRY_CLASSID")

    ' Add in the data
        If wsAllDataCombined.Range("B2").Value <> "No Transactions Found" Then
            ' "JOURNAL"
                wsImport.Range("B2").Formula = "=IF(E2="""","""",'All Data Combined'!A2)"
                
            ' "DATE"
                wsImport.Range("C2").Formula = "=IF(E2="""","""",'All Data Combined'!B2)"
                
            ' "REVERSEDATE"
                wsImport.Range("D2").Formula = "="""""
                
            ' "DESCRIPTION"
                wsImport.Range("E2").Formula = "=IF('All Data Combined'!C2='All Data Combined'!C1,"""",'All Data Combined'!C2)"
                
            ' "REFERENCE_NO"
                wsImport.Range("F2").Formula = "="""""
                
            ' "LINE_NO"
                wsImport.Range("G2").Formula = "=IF('All Data Combined'!C2='All Data Combined'!C1,1+G1,1)"
                
            ' "ACCT_NO"
                wsImport.Range("H2").Formula = "='All Data Combined'!E2"
                
            ' "LOCATION_ID"
                wsImport.Range("I2").Formula = "='All Data Combined'!F2"
                
            ' "DEPT_ID"
                wsImport.Range("J2").Formula = "='All Data Combined'!G2"
                
            ' "DOCUMENT"
                wsImport.Range("K2").Formula = "="""""
                
            ' "MEMO"
                wsImport.Range("L2").Formula = "='All Data Combined'!H2"
                
            ' "DEBIT"
                wsImport.Range("M2").Formula = "=IF(OR(H2=""11100"",H2=""11200"",H2=""11400"",H2=""11700"",G2=1),IF('All Data Combined'!I2>0,'All Data Combined'!I2,"""")," & _
                        "IF('All Data Combined'!I2<0,ABS('All Data Combined'!I2),""""))"
                
            ' "CREDIT"
                wsImport.Range("N2").Formula = "=IF(OR(H2=""11100"",H2=""11200"",H2=""11400"",H2=""11700"",G2=1),IF('All Data Combined'!I2<0,ABS('All Data Combined'!I2),"""")," & _
                        "IF('All Data Combined'!I2>0,'All Data Combined'!I2,""""))"
                        
            ' "SOURCEENTITY"
                wsImport.Range("O2").Formula = "="""""
                
            ' "CURRENCY"
                wsImport.Range("P2").Formula = "="""""
                
            ' "EXCH_RATE_DATE"
                wsImport.Range("Q2").Formula = "="""""
                
            ' "EXCH_RATE_TYPE_ID"
                wsImport.Range("R2").Formula = "="""""
                
            '  "EXCHANGE_RATE"
                wsImport.Range("S2").Formula = "="""""
                
            '  "STATE"
                wsImport.Range("T2").Formula = "=""Draft"""
                
            ' "ALLOCATION_ID"
                wsImport.Range("U2").Formula = "="""""
                
            ' "RASSET"
                wsImport.Range("V2").Formula = "="""""
                
            '  "RDEPRECIATION_SCHEDULE"
                wsImport.Range("W2").Formula = "="""""
                
            '  "RASSET_ADJUSTMENT"
                wsImport.Range("X2").Formula = "="""""
                
            '  "RASSET_CLASS"
                wsImport.Range("Y2").Formula = "="""""
                
            '  "RASSETOUTOFSERVICE"
                wsImport.Range("Z2").Formula = "="""""
                
            '  "GLDIMFUNDING_SOURCE"
                wsImport.Range("AA2").Formula = "='All Data Combined'!J2"
                
            '  "GLENTRY_PROJECTID"
                wsImport.Range("AB2").Formula = "="""""
                
            '  "GLENTRY_CUSTOMERID"
                wsImport.Range("AC2").Formula = "="""""
                
            ' "GLENTRY_VENDORID"
                wsImport.Range("AD2").Formula = "="""""
                
            ' "GLENTRY_EMPLOYEEID"
                wsImport.Range("AE2").Formula = "="""""
                
            ' "GLENTRY_ITEMID"
                wsImport.Range("AF2").Formula = "="""""
                
            ' "GLENTRY_CLASSID"
                wsImport.Range("AG2").Formula = "='All Data Combined'!K2"

            ' Fill the data down
                If AllDataCombinedLastRow > 2 Then
                    wsImport.Range("B2:AG" & AllDataCombinedLastRow).FillDown
                End If
                
            ' Format the worksheet.
                ' Format columns H:J as text
                    wsImport.Columns("H:J").NumberFormat = "@"
                    
                ' AutoFilter Columns.
                    wsImport.Range("A1:AG1").AutoFilter
                    
                ' AutoFit Columns.
                    wsImport.Columns("A:AG").AutoFit
                
        End If

    ' Finish by jumping over 'AdjustingRoute:' and going to 'CreateChecks:'
        GoTo CreateChecks


'--------------------------------------------
CreateChecks:
'--------------------------------------------
' Update the status bar
    Application.StatusBar = "Finding All Errors For The User to Manually Check"
    
' Go back into the 'wsPosTransactions' and 'wsNegTransactions' worksheets to help with user-required corrections.
    ' wsPosTransactions
        ' Unhide the worksheet
            wsPosTransactions.Visible = xlSheetVisible
        
        ' Add 3 Columns From AA:AC
            wsPosTransactions.Columns("AA:AC").Insert Shift:=xlToRight
            
        ' Add the column headers
            wsPosTransactions.Range("AA1:AC1").Value = Array("(ADJUSTMENT) Intacct - Account ID", "(ADJUSTMENT) Intacct - Division ID", "(ADJUSTMENT) Intacct - Funding Source")
        
        ' Add in the formulas
            ' Columns AA-AC: "(ADJUSTMENT) Intacct - Account ID", "(ADJUSTMENT) Intacct - Division ID", "(ADJUSTMENT) Intacct - Funding Source"
                wsPosTransactions.Range("AA2").Formula2 = "=IF(V2=""CHECK"",XLOOKUP(H2,'User-Required Checks'!E:E,'User-Required Checks'!K:M),V2:X2)"
                
            ' Make adjustments to previously existing formulas:
                ' "Intacct - Account Number"
                    wsPosTransactions.Range("BQ2").Formula = "=AA2"
                ' "Intacct - Department ID"
                    wsPosTransactions.Range("BS2").Formula = "=AB2"
                ' "Intacct - Funding Source"
                    wsPosTransactions.Range("BV2").Formula = "=AC2"
        
        ' Fill Down
            If PosTransactionsLastRow > 2 Then
                wsPosTransactions.Range("AA2:AA" & PosTransactionsLastRow).FillDown
                wsPosTransactions.Range("BQ2:BQ" & PosTransactionsLastRow).FillDown
                wsPosTransactions.Range("BS2:BS" & PosTransactionsLastRow).FillDown
                wsPosTransactions.Range("BV2:BV" & PosTransactionsLastRow).FillDown
            End If
            
        ' Rehide the worksheet
            wsPosTransactions.Visible = xlSheetHidden
    
    ' wsNegTransactions
        ' Unhide the worksheet
            wsNegTransactions.Visible = xlSheetVisible

        ' Add 3 Columns From AA:AC
            wsNegTransactions.Columns("AA:AC").Insert Shift:=xlToRight

        ' Add the column headers
            wsNegTransactions.Range("AA1:AC1").Value = Array("(ADJUSTMENT) Intacct - Account ID", "(ADJUSTMENT) Intacct - Division ID", "(ADJUSTMENT) Intacct - Funding Source")

        ' Add in the formulas
            ' Columns AA-AC: "(ADJUSTMENT) Intacct - Account ID", "(ADJUSTMENT) Intacct - Division ID", "(ADJUSTMENT) Intacct - Funding Source"
                If ImportType = "CRJ" Then
                    wsNegTransactions.Range("AA2").Formula2 = "=IF(V2=""CHECK"",XLOOKUP(H2,'User-Required Checks'!E:E,'User-Required Checks'!K:M),V2:X2)"
                
                ElseIf ImportType = "Adjusting" Then
                    wsNegTransactions.Range("AA2").Formula2 = "=IF(OR(V2=""CHECK"",T2=""Suggested""),XLOOKUP(H2,'User-Required Checks'!E:E,'User-Required Checks'!K:M),V2:X2)"
                
                End If

            ' Make adjustments to previously existing formulas:
                ' "Intacct - Account Number"
                    wsNegTransactions.Range("BQ2").Formula = "=AA2"

                ' "Intacct - Department ID"
                    wsNegTransactions.Range("BS2").Formula = "=AB2"

                ' "Intacct - Funding Source"
                    wsNegTransactions.Range("BV2").Formula = "=AC2"

        ' Fill Down
            If NegTransactionsLastRow > 2 Then
                wsNegTransactions.Range("AA2:AA" & NegTransactionsLastRow).FillDown
                wsNegTransactions.Range("BQ2:BQ" & NegTransactionsLastRow).FillDown
                wsNegTransactions.Range("BS2:BS" & NegTransactionsLastRow).FillDown
                wsNegTransactions.Range("BV2:BV" & NegTransactionsLastRow).FillDown
            End If

        ' Rehide the worksheet
            wsNegTransactions.Visible = xlSheetHidden


' Start populating the 'wsUserChecks' Worksheet.
    ' Adjustments to School Allocations
        ' Create the heading section
            With wsUserChecks.Range("A1:M1")
                .Merge
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleSingle
                .Value = "ADJUSTMENTS TO SCHOOL ALLOCATIONS"
                .Interior.Color = RGB(200, 100, 100)
            End With
        
        ' Add the column headers
            wsUserChecks.Range("A2:F2").Value = Array("Disbursment Date", "Transaction ID", "Disbursement ID", "Disbursement Amount", "School Name", "Corrected School Abbreviation")

        ' Add in the formula to extract the relevant data to help the user
            wsUserChecks.Range("A3").Formula2 = "=IFERROR(CHOOSECOLS(FILTER('Standardized Donation Site Data'!A:N,'Standardized Donation Site Data'!Q:Q=""CHECK""),3,10,5,8)," & _
                    """No Adjustments Needed"")"
        
        ' If nothing populates, make the cell green
            If wsUserChecks.Range("A3").Value = "No Adjustments Needed" Then
                wsUserChecks.Range("A3").Interior.Color = vbGreen
                wsUserChecks.Range("A1:M1").Interior.Color = vbGreen
            End If
        
        ' Find the last row
            UserChecksLastRow = wsUserChecks.Cells(wsUserChecks.Rows.Count, "A").End(xlUp).Row
        
        ' Add in data validation
            If wsUserChecks.Range("A3").Value <> "No Adjustments Needed" Then
                ' Create a worksheet called "School List" and store it in a variable called 'wsSchoolList'
                    Set wsSchoolList = wbMacro.Worksheets.Add(After:=wsUserChecks)
                    
                    ' Rename the worksheet
                        wsSchoolList.Name = "School List"
                    
                    ' Add in the data
                        ' School Names (BBR -> BCSI -> BDC -> BTCS)
                            SchoolNames = Array("BASIS Baton Rouge Materra|BBRM", "BASIS Baton Rouge Mid City|BRMC", "BASIS Baton Rouge Schools, Inc.|BBR", _
                                        "BASIS Ahwatukee|AHW", "BASIS Chandler|CHD", "BASIS Chandler Primary North|CHPN", "BASIS Chandler Primary South|CHPS", "BASIS Charter Schools, Inc.|BCSI", _
                                        "BASIS Flagstaff|FLG", "BASIS Goodyear|GDY", "BASIS Goodyear Primary|GDYP", "BASIS Mesa|MES", "BASIS Oro Valley|OV", "BASIS Oro Valley Primary|OVP", _
                                        "BASIS Peoria|PEO", "BASIS Peoria Primary|PEOP", "BASIS Phoenix|PHX", "BASIS Phoenix Central|PHXC", "BASIS Phoenix North|PHXN", _
                                        "BASIS Phoenix Primary|PHXP", "BASIS Phoenix South|PHXS", "BASIS Prescott|PRE", "BASIS Scottsdale|SCD", "BASIS Scottsdale Primary East|SCPE", _
                                        "BASIS Scottsdale Primary West|SCPW", "BASIS Tucson North|TUCN", "BASIS Tucson Primary|TUCP", _
                                        "BASIS DC|BDC", "BASIS Washington, DC|DC", _
                                        "BASIS Austin|AUS", "BASIS Austin Primary|AUSP", "BASIS Benbrook|BEN", "BASIS Cedar Park|CPK", "BASIS Cedar Park Primary|CPKP", "BASIS Pflugerville|PFL", _
                                        "BASIS Pflugerville Primary|PFLP", "BASIS Plano|PLN", "BASIS Plano Primary|PLNP", "BASIS Richardson|RCH", "BASIS Richardson Primary|RCHP", _
                                        "BASIS San Antonio Jack Lewis Jr.|JLJ", "BASIS San Antonio Jack Lewis Jr. Primary|JLJP", "BASIS San Antonio Medical Center Primary|SAMC", _
                                        "BASIS San Antonio North Central Primary|SANC", "BASIS San Antonio Northeast|SANE", "BASIS San Antonio Northeast Primary|SPNE", _
                                        "BASIS San Antonio Shavano|SAS", "BASIS Texas Charter Schools, Inc.|BTCS")
                            
                            ' Unpack the array
                                With wsSchoolList.Range("A2").Resize(UBound(SchoolNames) - LBound(SchoolNames) + 1, 1)
                                    .Value = Application.Transpose(SchoolNames)
                                End With
                                
                        ' Split them out
                            wsSchoolList.Range("B2").Formula2 = "=TEXTSPLIT(A2,""|"")"
                            
                        ' FillDown
                            wsSchoolList.Range("B2:B49").FillDown
                        
                        ' Copy and Paste Values Only
                            wsSchoolList.Range("B:C").Value = wsSchoolList.Range("B:C").Value
                        
                        ' Delete column A
                            wsSchoolList.Columns(1).Delete
                        
                        ' Column Headers
                            wsSchoolList.Range("A1:B1").Value = Array("School Name", "School Abbreviation")
                            
                        ' Sort by School Name (Column A)
                            With wsSchoolList.Sort
                                .SortFields.Clear
                                .SortFields.Add key:=wsSchoolList.Range("A2:A" & wsSchoolList.Cells(wsSchoolList.Rows.Count, "A").End(xlUp).Row), _
                                    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                                .SetRange wsSchoolList.Range("A1:B" & wsSchoolList.Cells(wsSchoolList.Rows.Count, "A").End(xlUp).Row)
                                .Header = xlYes
                                .Apply
                            End With
                        
                        ' Format the worksheet
                            wsSchoolList.Columns("A:B").AutoFit
                            
                        ' Hide the worksheet
                            wsSchoolList.Visible = xlSheetHidden
                        
                ' Set up the data validation
                    ' Check if there is more than 1 disbursement that needs a user-required school allocation
                        If UserChecksLastRow > 3 Then
                            Set DataValidationRange_School = wsUserChecks.Range("E3:E" & UserChecksLastRow)
                        Else
                            Set DataValidationRange_School = wsUserChecks.Range("E3")
                        End If
                    
                    ' Create the data validation based off of the range 'DataValidationRange_School'
                        With DataValidationRange_School.Validation
                                .Delete ' Clear existing validation
            
                                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="='School List'!$A$2:$A$49"
                            
                                .IgnoreBlank = True
                                .InCellDropdown = True
                                .InputTitle = ""
                                .ErrorTitle = "Invalid Entry"
                                .InputMessage = ""
                                .ErrorMessage = "Please select a valid school from the list."
                                .ShowInput = True
                                .ShowError = True
                        End With
                    
                    ' Unlock the cells
                        DataValidationRange_School.Locked = False
                
                ' Add the formulas to column F for 'Corrected School Abbreviation' and fill the cell in with yellow
                    ' Check if there is more than 1 disbursement that needs a user-required school allocation
                        If UserChecksLastRow > 3 Then
                            wsUserChecks.Range("F3").Formula2 = "=IF(E3="""","""",XLOOKUP(E3,'School List'!A:A,'School List'!B:B))"
                            wsUserChecks.Range("F3:F" & UserChecksLastRow).FillDown
                            wsUserChecks.Range("E3:E" & UserChecksLastRow).Interior.Color = vbYellow
                        Else
                            wsUserChecks.Range("F3").Formula2 = "=IF(E3="""","""",XLOOKUP(E3,'School List'!A:A,'School List'!B:B))"
                            wsUserChecks.Range("E3").Interior.Color = vbYellow
                        End If
            End If
        

    ' Adjustments to: Account|Division|Funding Source
        ' Use 'UserChecksLastRow' and add 5, for the starting row of this section.
            UserChecksNewCheckRow = UserChecksLastRow + 5
        
        ' Create the heading section
            With wsUserChecks.Range("A" & UserChecksNewCheckRow & ":M" & UserChecksNewCheckRow)
                .Merge
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleSingle
                .Value = "ADJUSTMENTS TO: ACCOUNT|DIVISION|FUNDING SOURCE"
                .Interior.Color = RGB(200, 100, 100)
            End With
        
        ' Add the column headers
            wsUserChecks.Range("A" & UserChecksNewCheckRow + 1 & ":M" & UserChecksNewCheckRow + 1).Value = Array("Transaction Date", "Transaction ID", "Disbursement ID", _
                    "Amount", "PMT-ID", "Campaign Name", "Opportunity Name", "Family Name ", "Intacct School ID", "Suggestions", "Corrected - Account", "Corrected - Division", _
                    "Corrected - Funding Source")

        ' Add in the formula to extract the relevant data to help the user
            If ImportType = "CRJ" Then
                wsUserChecks.Range("A" & UserChecksNewCheckRow + 2).Formula2 = "" & _
                        "=LET(PosTxn," & _
                            "CHOOSECOLS(FILTER('Standardized SF Data'!A2:V" & StandardSFLastRow & ",(ISNUMBER(MATCH('Standardized SF Data'!H2:H" & StandardSFLastRow & _
                            ",'Positive Transactions'!H2:H" & PosTransactionsLastRow & ",0)))*('Standardized SF Data'!V2:V" & StandardSFLastRow & "=""CHECK"")),1,12,13,14,8,4,5,10,21,19)," & _
                        "NegTxn," & _
                            "HSTACK(CHOOSECOLS(FILTER('Standardized SF Data'!A2:V" & StandardSFLastRow & ",(ISNUMBER(MATCH('Standardized SF Data'!H2:H" & StandardSFLastRow & _
                            ",'Negative Transactions'!H2:H" & NegTransactionsLastRow & ",0)))*('Standardized SF Data'!V2:V" & StandardSFLastRow & "=""CHECK"")),1,12,13)," & _
                            "-1 * CHOOSECOLS(FILTER('Standardized SF Data'!A2:V" & StandardSFLastRow & ",(ISNUMBER(MATCH('Standardized SF Data'!H2:H" & StandardSFLastRow & _
                            ",'Negative Transactions'!H2:H" & NegTransactionsLastRow & ",0)))*('Standardized SF Data'!V2:V" & StandardSFLastRow & "=""CHECK"")),14)," & _
                            "CHOOSECOLS(FILTER('Standardized SF Data'!A2:V" & StandardSFLastRow & ",(ISNUMBER(MATCH('Standardized SF Data'!H2:H" & StandardSFLastRow & _
                            ",'Negative Transactions'!H2:H" & NegTransactionsLastRow & ",0)))*('Standardized SF Data'!V2:V" & StandardSFLastRow & "=""CHECK"")),8,4,5,10,21,19))," & _
                        "IF(AND(COUNT(PosTxn)>0,COUNT(NegTxn)>0),VSTACK(PosTxn,NegTxn),IF(COUNT(PosTxn)>0,PosTxn,IF(COUNT(NegTxn)>0,NegTxn,""No Adjustments Needed""))))"
                        
            ElseIf ImportType = "Adjusting" Then
                wsUserChecks.Range("A" & UserChecksNewCheckRow + 2).Formula2 = "" & _
                        "=IFERROR(" & _
                            "HSTACK(" & _
                                "CHOOSECOLS(FILTER('Standardized SF Data'!A2:V" & StandardSFLastRow & ",(ISNUMBER(MATCH('Standardized SF Data'!H2:H" & StandardSFLastRow & _
                                    ",'Negative Transactions'!H2:H" & NegTransactionsLastRow & ",0)))*('Standardized SF Data'!T2:T" & StandardSFLastRow & "=""Suggested"")),1,12,13)," & _
                                "-1 * CHOOSECOLS(FILTER('Standardized SF Data'!A2:V" & StandardSFLastRow & ",(ISNUMBER(MATCH('Standardized SF Data'!H2:H" & StandardSFLastRow & _
                                    ",'Negative Transactions'!H2:H" & NegTransactionsLastRow & ",0)))*('Standardized SF Data'!T2:T" & StandardSFLastRow & "=""Suggested"")),14)," & _
                                "CHOOSECOLS(FILTER('Standardized SF Data'!A2:V" & StandardSFLastRow & ",(ISNUMBER(MATCH('Standardized SF Data'!H2:H" & StandardSFLastRow & _
                                    ",'Negative Transactions'!H2:H" & NegTransactionsLastRow & ",0)))*('Standardized SF Data'!T2:T" & StandardSFLastRow & "=""Suggested"")),8,4,5,10,21,19))," & _
                        """No Adjustments Needed"")"
            End If
                    
        ' Find the new last row
            UserChecksLastRow = wsUserChecks.Cells(wsUserChecks.Rows.Count, "A").End(xlUp).Row
        
        ' If nothing populates, make the cell green
            If wsUserChecks.Range("A" & UserChecksNewCheckRow + 2).Value = "No Adjustments Needed" Then
                wsUserChecks.Range("A" & UserChecksNewCheckRow + 2).Interior.Color = vbGreen
            End If
            
        ' Check if there are any 'Account|Division|Funding Source' Adjustments to be made. Unlock and fill the cells in with yellow if there are.
            If wsUserChecks.Range("A" & UserChecksNewCheckRow + 2).Value <> "No Adjustments Needed" Then
                With wsUserChecks.Range(("K" & (UserChecksNewCheckRow + 2) & ":M" & UserChecksLastRow))
                    .Locked = False
                    .Interior.Color = vbYellow
                End With
            Else
                wsUserChecks.Range("A" & UserChecksNewCheckRow & ":M" & UserChecksNewCheckRow).Interior.Color = vbGreen
            End If
                
' Transaction ID CHECK
        ' Find the last row
            UserChecksLastRow = wsUserChecks.Cells(wsUserChecks.Rows.Count, "A").End(xlUp).Row
        
        ' Use 'UserChecksLastRow' and add 5, for the starting row of this section.
            UserChecksNewCheckRow = UserChecksLastRow + 5
        
        ' Create the heading section
            With wsUserChecks.Range("A" & UserChecksNewCheckRow & ":M" & UserChecksNewCheckRow)
                .Merge
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleSingle
                .Value = "TRANSACTION ID CHECK"
                .Interior.Color = RGB(200, 100, 100)
            End With
        
        ' Add the column headers
            wsUserChecks.Range("A" & UserChecksNewCheckRow + 1 & ":F" & UserChecksNewCheckRow + 1).Value = Array("Transaction Date", "Transaction ID", "Disbursement ID", "SF - Gross Amount", _
            "Click and Pledge Site - Gross Amount", "Total Gross Amount - Variance")
                    
        ' Add the formulas
            ' The new data start row: " & (UserChecksNewCheckRow + 2) & "
            
            ' "Transaction Date"
                wsUserChecks.Range("A" & UserChecksNewCheckRow + 2).Formula2 = "=XLOOKUP(B" & (UserChecksNewCheckRow + 2) & ",'Standardized Donation Site Data'!D:D," & _
                        "'Standardized Donation Site Data'!B:B)"
            
            ' "Transaction ID"
                wsUserChecks.Range("B" & UserChecksNewCheckRow + 2).Formula2 = "=UNIQUE('Standardized Donation Site Data'!D2:D" & StandardDonationsLastRow & ")"
            
            ' "Disbursement ID"
                wsUserChecks.Range("C" & UserChecksNewCheckRow + 2).Formula2 = "=XLOOKUP(B" & (UserChecksNewCheckRow + 2) & ",'Standardized Donation Site Data'!D:D,'Standardized Donation Site Data'!E:E)"
            
            ' "SF - Gross Amount"
                wsUserChecks.Range("D" & UserChecksNewCheckRow + 2).Formula = "=SUMIF('Positive Transactions'!AH:AH,B" & (UserChecksNewCheckRow + 2) & _
                        ",'Positive Transactions'!AV:AV) + SUMIF('Negative Transactions'!AH:AH,B" & (UserChecksNewCheckRow + 2) & ",'Negative Transactions'!AV:AV)"
            
            ' "Click and Pledge Site - Gross Amount"
                wsUserChecks.Range("E" & UserChecksNewCheckRow + 2).Formula = "=SUMIF('Standardized Donation Site Data'!D:D,B" & (UserChecksNewCheckRow + 2) & _
                        ",'Standardized Donation Site Data'!G:G)"
            
            ' "Total Gross Amount - Variance"
                wsUserChecks.Range("F" & UserChecksNewCheckRow + 2).Formula = "=E" & (UserChecksNewCheckRow + 2) & "-D" & (UserChecksNewCheckRow + 2) & ""
        
        ' Find the last row
            UserChecksLastRow = wsUserChecks.Cells(wsUserChecks.Rows.Count, "B").End(xlUp).Row
                
        ' Fill Down
            If UserChecksLastRow > (UserChecksNewCheckRow + 2) Then
                wsUserChecks.Range("A" & (UserChecksNewCheckRow + 2) & ":A" & UserChecksLastRow).FillDown
                wsUserChecks.Range("C" & (UserChecksNewCheckRow + 2) & ":F" & UserChecksLastRow).FillDown
            End If

        ' Highlight the variances
            VarianceCount_Gross = 0
            
            For DVRow = (UserChecksNewCheckRow + 2) To UserChecksLastRow
                If wsUserChecks.Range("F" & DVRow).Value <> 0 Then
                    wsUserChecks.Range("F" & DVRow).Interior.Color = vbYellow
                    VarianceCount_Gross = VarianceCount_Gross + 1
                End If
            Next DVRow
            
            If (VarianceCount_Gross = 0) Then
                wsUserChecks.Range("A" & UserChecksNewCheckRow & ":M" & UserChecksNewCheckRow).Interior.Color = vbGreen
            End If
            
    ' Format the worksheet
        wsUserChecks.Range("A:A").NumberFormat = "mm/dd/yyyy"
        wsUserChecks.Columns("A:M").AutoFit
    
    ' Protect the worksheet
        wsUserChecks.Protect

    ' If both sections have nothing populate, make the tab color green and hide the worksheet.
        If (wsUserChecks.Range("A3").Value = "No Adjustments Needed") And (wsUserChecks.Range("A10").Value = "No Adjustments Needed") _
            And (VarianceCount_Gross = 0) Then
            
            wsUserChecks.Tab.Color = vbGreen
            
            
            If ImportType = "CRJ" Then
                wsImportCRJ.Tab.Color = vbGreen
                If wsImportAdjusting.Visible = xlSheetVisible Then
                    wsImportAdjusting.Tab.Color = vbGreen
                    wsImportCRJ.Activate
                End If
                    
            ElseIf ImportType = "Adjusting" Then
                wsImport.Tab.Color = vbGreen
                wsImport.Activate
            End If
            
        Else
            wsUserChecks.Tab.Color = vbRed
            wsUserChecks.Activate
            
            If ImportType = "CRJ" Then
                wsImportCRJ.Tab.Color = vbYellow
                If wsImportAdjusting.Visible = xlSheetVisible Then
                    wsImportAdjusting.Tab.Color = vbYellow
                End If
                
            ElseIf ImportType = "Adjusting" Then
                wsImport.Tab.Color = vbYellow
                
            End If
            
        End If
    
' Protect the 'wsStandardDonations' and 'wsStandardSF' worksheets
    wsStandardDonations.Protect
    wsStandardSF.Protect
    

' Provide a message to the user to help them know the macro has completed successfully.
    MsgBox "The macro has completed successfully. Thank you for your patience!" & vbNewLine & vbNewLine & _
           "From the folder selected '" & (NonExcelFilesCount + UsedProPayDailyFilesCount + UsedProPayMonthlyFilesCount + _
                UsedStripeDailyFilesCount + UsedStripeMonthlyFilesCount + UnusedFilesCount) & "' files were found. Here is the breakdown:" & vbNewLine & vbNewLine & _
           "Used ProPay 'Daily' Files: " & UsedProPayDailyFilesCount & vbNewLine & _
           "Used ProPay 'Monthly' Files: " & UsedProPayMonthlyFilesCount & vbNewLine & _
           "Used Stripe 'Daily' Files: " & UsedStripeDailyFilesCount & vbNewLine & _
           "Used Stripe 'Monthly' Files: " & UsedStripeMonthlyFilesCount & vbNewLine & _
           "Unused Files: " & UnusedFilesCount & vbNewLine & _
           "Non-Excel Type Files (also unused): " & NonExcelFilesCount, _
           vbInformation, "Macro Completed Successfully"

' End the macro by reseting the workbook
    GoTo ResetTheWorkbook


NoFiles:
' If the user has the first report, but does not have the folder for click and pledge ready, set up a button page.
    MsgBox Title:="Issue Processing Reports", Prompt:=ExtraMessage, Buttons:=vbExclamation

' Check if the "No Donation Site Report" is created yet.
    For Each ws In wbMacro.Worksheets
        If ws.Name = "No Donation Site Report" Then
            wsInitialData.Visible = xlSheetHidden
            GoTo ResetTheWorkbook
        End If
    Next ws
    
' If it was not found, create it.
    ' Create the worksheet.
        Set wsButton = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets("COMPLETE RESET"))
    
    ' Rename the worksheet.
        wsButton.Name = "No Donation Site Report"
    
    ' Format the worksheet.
        wsButton.Cells.Interior.Color = vbBlack
        
    ' Create the button
        Set DonationSiteButton = wsButton.Buttons.Add(150, 50, 825, 275)
        
        With DonationSiteButton
            .Caption = "Click here to add the '" & Site & "' Reports"
            .OnAction = ConverterName
            .Font.Size = 50
            .Font.Bold = True
            .Font.Color = RGB(200, 200, 0)
        End With
        
    ' Hide the other 'Initial Data' worksheet.
        wsInitialData.Visible = xlSheetHidden


ResetTheWorkbook:
' Get rid of the the status bar
    Application.StatusBar = False
    
' Bring back alerts and screen updating
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

' Turn back on Calculations
    Application.Calculation = xlCalculationAutomatic


End Sub




