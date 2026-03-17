Sub YourCause_AR_New()
' ============================================================
' MODULE: YourCause_AR
' AUTHOR: Austin Glawe
' CREATED: 2025.09.30
' LAST UPDATED: 2026.03.17
' CURRENT MAINTAINER: See Module 'A_Global_Constants'
' ============================================================
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' PURPOSE, REQUIREMENTS, FLOW, AND UPDATES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                            PURPOSE
    ' ============================================================
        ' The purpose of this macro is to:
            ' Merge Salesforce data with donation site report data to generate an Intacct import file, including all supporting documentation.
                ' Consolidate all donation site reports into a single worksheet while preserving the original report data for supporting documentation.
            ' Connect transactions to their corresponding deposits to support accurate bank reconciliation and deposit imports into Intacct.
            ' Reconcile records between Salesforce, the donation site, and Intacct to identify missing, incomplete, or incorrectly entered transactions.
            
    ' ============================================================
    '                         REQUIREMENTS
    ' ============================================================
        ' One of the following reports:
            ' Salesforce Report
                ' Found at https://basised.lightning.force.com/lightning/r/Report/00ORj000006bQGDMA2/view
            ' Intacct Report
                ' Found in Intacct under Platform Services >> Custom Reports >> Undeposited Funds Report
                
        ' A folder containing all donation site reports to process
            ' The folder name must contain "Your Cause"
          
    ' ============================================================
    '                             FLOW
    ' ============================================================
        '  1. User selects Salesforce or Intacct report
        '  2. User selects folder containing donation site reports
        '  3. Donation site reports are consolidated and supporting worksheets are created
        '  4. Data is merged with the selected Salesforce or Intacct report
        '  5. Transactions are connected to deposits
        '  6. Reconciliation checks are performed across Salesforce, donation site, and Intacct data
        '  7. Transactions requiring user review are filtered to a designated worksheet
        '  8. Intacct import file is generated
        '  9. User reviews and resolves any flagged transactions
        ' 10. Import file is uploaded to Intacct
    
    ' ============================================================
    '             UPDATE LOG (LAST UPDATED: 2026.03.17)
    ' ============================================================
        ' Original Production Rollout Date: 2025.09.30

        ' Updates:
            ' 2026.03.17 - Initiated the update log.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CONFIGURATIONS AND VARIABLES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '                        CONFIGURATIONS
    ' ============================================================
        ' These values control how the converter behaves.
        ' They are declared first because they act as the setup rules for the rest of the macro.

        ' ---------------------------------------------
        '             DECLARE CONFIGURATIONS
        ' ---------------------------------------------
            Dim wbMacro As Workbook
            
            Dim ConverterName As String
            Dim DonationSite As String
            Dim DonationSite_Salesforce As String
            
            Dim JournalName As String
            Dim JournalType As String
            
            Dim AllowConsolidationOnly As Boolean
            Dim IncludeOriginalReports As Boolean
            Dim AllowRevenueAmountAdjustments As Boolean
            Dim AllowJournalTypeManualOverride As Boolean
            
            Dim RowsToDeleteFromBottomOfDonationSiteReport As Long
            
            Dim AssignedHeaderRow_InitialReport As Long
            Dim AllowHeaderRowSearch_InitialReport As Boolean
            
            Dim AssignedHeaderRow_DonationSiteReports As Long
            Dim AllowHeaderRowSearch_DonationSiteReports As Boolean
            
            Dim ColumnHeaders_Initial_Intacct As Variant
            Dim ColumnHeaders_Initial_Salesforce As Variant
            Dim ColumnHeaders_YourCause As Variant

        ' ---------------------------------------------
        '             ASSIGN CONFIGURATIONS
        ' ---------------------------------------------
            ' Store a reference to this workbook so it is clearly distinguished from
            ' any temporary workbooks opened during the converter process.
            Set wbMacro = ThisWorkbook
            
            ' Store the converter procedure name so it can be assigned to a button
            ' if the user needs to return later and continue the process.
            ConverterName = "YourCause_New.YourCause_AR_New"
            
            ' Store the Donation Site name used throughout this converter.
            DonationSite = "Your Cause"
            
            ' This is the Salesforce-side constant tied to this Donation Site.
            ' See Module 'A_Global_Constants'.
            DonationSite_Salesforce = DonationSiteYourCause

            ' Store the journal name used for the Adjusting import route.
            ' By default this should be "SFREV". ("CHAR" was used prior to 2025.10)
            JournalName = "SFREV"

            ' This switch allows the converter to stop after consolidating / preparing
            ' the Donation Site reports without building the final Intacct import file.
            ' By default this should be False.
            AllowConsolidationOnly = False
            
            ' This switch determines whether the original Donation Site report tabs are
            ' preserved in the workbook for review and supporting documentation.
            ' By default this should be True.
            IncludeOriginalReports = True
            
            ' This setting controls how many non-data rows should be deleted from the
            ' bottom of each Donation Site report.
            ' By default this should be 0 for YourCause.
            RowsToDeleteFromBottomOfDonationSiteReport = 0
            
            ' This switch controls whether mismatching revenue amounts between
            ' Salesforce and the Donation Site are allowed to continue through the
            ' converter as valid adjustment entries.
            ' By default this should be True unless a stricter workflow is required.
            AllowRevenueAmountAdjustments = True
            
            ' ..............................
            '         INITIAL REPORT
            '       HEADER ROW SETTINGS
            ' ..............................
                ' This switch allows the converter to search all rows for the Initial
                ' Report headers instead of assuming a fixed row.
                AllowHeaderRowSearch_InitialReport = True
                
                ' This setting allows a fixed Initial Report header row to be used when known.
                ' Common values:
                '   1 = CSV downloads / standard exports
                '   5 = some Excel exports
                ' By default this should be 0 so searching can determine the row.
                AssignedHeaderRow_InitialReport = 0
                
                ' Logic override:
                ' If no valid assigned row exists and searching was turned off,
                ' force searching back on so the converter can still function.
                If AssignedHeaderRow_InitialReport < 1 And AllowHeaderRowSearch_InitialReport = False Then
                    AllowHeaderRowSearch_InitialReport = True
                End If
            
            ' ..............................
            '      DONATION SITE REPORT
            '       HEADER ROW SETTINGS
            ' ..............................
                ' This switch allows the converter to search each Donation Site report
                ' for its headers instead of using a fixed header row.
                AllowHeaderRowSearch_DonationSiteReports = False
            
                ' This setting defines the expected header row for the Donation Site reports.
                ' For YourCause, the headers are currently expected on row 1.
                AssignedHeaderRow_DonationSiteReports = 1
                
                ' Logic override:
                ' If the assigned row is invalid and searching is disabled, default to row 1.
                If AssignedHeaderRow_DonationSiteReports < 1 And AllowHeaderRowSearch_DonationSiteReports = False Then
                    AssignedHeaderRow_DonationSiteReports = 1
                End If
            
            ' ..............................
            '        JOURNAL SETTINGS
            ' ..............................
                ' This switch allows the user to manually force the final journal route.
                ' By default this should usually be False unless testing or special handling is needed.
                AllowJournalTypeManualOverride = True
                
                ' Valid values when manual override is enabled:
                '   "Adjusting"
                '   "CRJ"
                JournalType = "Adjusting"
                
                ' Logic override:
                ' If manual override is off, clear the JournalType so it can be determined later.
                ' If manual override is on but the value is invalid, clear it as well.
                If AllowJournalTypeManualOverride = False Then
                    JournalType = ""
                ElseIf JournalType <> "Adjusting" And JournalType <> "CRJ" Then
                    JournalType = ""
                End If

            ' ..............................
            '         INITIAL REPORT
            '         COLUMN HEADERS
            ' ..............................
                ' Intacct Report Column Headers (A:AC) - 29 columns
                ColumnHeaders_Initial_Intacct = Array("Journal Entry Modified Date", "Close Date", "Batch Posting Date", "SF Donation Site", "C&P Number", _
                        "SF Transaction ID", "SF Disbursement ID", "SF Payment Method", "SF Check Number", "SF Payment Number", "SF Primary Contact", _
                        "SF Account Name", "SF Company Name", "SF Campaign Source", "SF Opportunity Name", "Memo", "Location Name", "Location ID", "Account Number", _
                        "Division ID", "Funding Source", "Debt Service Series ID", "Journal", "Journal Number", "Journal Description", "Record Number", _
                        "Credit Amount", "Debit Amount", "Amount")
                
                ' Salesforce Report Column Headers (A:T) - 20 columns
                ColumnHeaders_Initial_Salesforce = Array("Payment: Created Date", "Close Date", "Deposit Date", "Donation Site", "C&P Order Number", _
                        "Check/Reference Number", "Disbursement ID", "Payment Type", "Check Number", "Payment: Payment Number", "Primary Contact", "Account Name", _
                        "Company Name", "Primary Campaign Source", "Opportunity Name", "C&P Account Name", "C&P Account Name Correction", _
                        "Payment Amount", "Campaign Type", "Description")
                
            ' ..............................
            '        DONATION SITE REPORT
            '         COLUMN HEADERS
            ' ..............................
                ' YourCause Column Headers (A:AO) - 41 columns
                ' If this array changes, also update the section that sets up
                ' the 'wsConsolidatedData' worksheet environment.
                ColumnHeaders_YourCause = Array("Donation Date", "Company Name", "Transaction Id", "Transaction Type", "Transaction Amount", "Fee Amount", _
                        "Disbursement Fee Amount", "Received Amount", "Is Disbursed?", "Payment ID", "Is Most Recent Payment", "Payment Create Date", _
                        "Payment Status Date", "Payment Status", "Donor Type", "Donor ID", "Donor First Name", "Donor Last Name", "Donor Full Name", _
                        "Donor Email Address", "Donor Address", "Donor Address 2", "Donor City", "Donor State/Province/Region", "Donor Postal Code", _
                        "Donor Country", "Match Donor Id", "Match Donor First Name", "Match Donor Last Name", "Match Donor Email Address", "Dedication Type", _
                        "Dedication", "Designation", "Registration Id", "Designated Charity Name", "Donation Status", "Alternate Recognition Name", "Segment Name", _
                        "Local Currency Receipt Amount", "Local Currency Type", "Fundraising ID")

    ' ============================================================
    '                           VARIABLES
    ' ============================================================
        ' These variables support the working logic of the converter.
        ' They are grouped by purpose so the macro is easier to read,
        ' maintain, and troubleshoot later.

        ' ---------------------------------------------
        '         GENERAL CONTROL / USER RESPONSE
        ' ---------------------------------------------
            Dim UserResponse As VbMsgBoxResult
            
            Dim ExitMessage As String
            Dim ExitMessage_Title As String
            
            Dim ExtraMessage As String
            Dim ExtraMessage_Title As String

        ' ---------------------------------------------
        '          WORKSHEET / WORKBOOK REFERENCES
        ' ---------------------------------------------
            Dim ws As Worksheet
            Dim wsCheck As Worksheet
            Dim wsNew As Worksheet
            Dim wsButton As Worksheet
            Dim wsSchoolValidation As Worksheet
            
            Dim wbTemp_InitialReport As Workbook
            Dim wbTemp_DonationSite As Workbook
            
            Dim wsTemp_InitialReport As Worksheet
            Dim wsTemp_DonationSite As Worksheet
            
            Dim wsInitialData As Worksheet
            Dim wsConsolidatedData As Worksheet
            
            Dim wsStandardizedSF As Worksheet
            Dim wsStandardizedDonationSiteData As Worksheet
            Dim wsDisbursementData As Worksheet
            Dim wsRelevantTransactions As Worksheet
            Dim wsFees As Worksheet
            Dim wsBankDeposits As Worksheet
            Dim wsConnectionAnalysis As Worksheet
            Dim wsUserRequiredAdjustments As Worksheet
            
            ' Adjusting Journal path worksheets
            Dim wsAdjustingUnfiltered As Worksheet
            Dim wsAdjustingFiltered As Worksheet
            Dim wsAdjustingJournal As Worksheet
            
            ' CRJ path worksheets
            Dim wsCRJUnfiltered As Worksheet
            Dim wsCRJFiltered As Worksheet
            Dim wsCRJ As Worksheet

        ' ---------------------------------------------
        '            FILE / FOLDER DIALOGS
        ' ---------------------------------------------
            Dim fdFilePath_InitialReport As FileDialog
            Dim fdFolderPath_DonationSite As FileDialog

        ' ---------------------------------------------
        '            FILE / FOLDER PATHS
        ' ---------------------------------------------
            Dim FilePath_InitialReport As String
            Dim FolderPath_DonationSite As String
            Dim FolderPath_ProcessLater As String
            
            Dim SourceFilePath As String
            Dim DestinationFilePath As String

        ' ---------------------------------------------
        '         INITIAL REPORT STATE / ANALYSIS
        ' ---------------------------------------------
            Dim InitialExists As Boolean
            Dim InitialPath As String
            
            Dim TempLastRow_InitialReport As Long
            Dim HeaderRow_InitialReport As Long
            Dim ColumnCheck_InitialReport As Long
            Dim LastRow_InitialData As Long

        ' ---------------------------------------------
        '      DONATION SITE FILE / REPORT TRACKING
        ' ---------------------------------------------
            Dim FileCount_DonationSite As Long
            Dim FileName_DonationSite As String
            Dim FileNamesList_DonationSite() As String
            
            Dim FileNumber_DonationSite As Long
            Dim FileCount_WrongReport As Long
            Dim FileCount_DonationSite_Unusable As Long
            Dim FileCount_DonationSite_Used As Long
            
            Dim WorksheetName As String
            Dim wsFound As Boolean

        ' ---------------------------------------------
        '       DONATION SITE TEMP WORKSHEET METRICS
        ' ---------------------------------------------
            Dim LastRow_TempDonationSite As Long
            Dim LastRow_TempDonationSite_Adjusted As Long
            
            Dim HeaderRow_DonationSite As Long
            Dim ColumnMatch_DonationSite As Long
            Dim Col_DonationSite As Long
            
            Dim DataStartRow_DonationSite As Long
            Dim DataRows_DonationSite As Long
            Dim DataRows_DonationSite_Total As Long
            Dim CurrentRow_DonationSite As Long
            
            Dim LastRow_ConsolidatedData As Long
            Dim LastRow_ConsolidatedData_AfterInsert As Long

        ' ---------------------------------------------
        '          STANDARDIZED / OUTPUT LAST ROWS
        ' ---------------------------------------------
            Dim LastRow_StandardizedSF As Long
            Dim LastRow_StandardizedDonationSiteData As Long
            Dim LastRow_DisbursementData As Long
            Dim LastRow_RelevantTransactions As Long
            Dim LastRow_Fees As Long
            Dim LastRow_BankDeposits As Long
            Dim LastRow_ConnectionAnalysis As Long
            Dim LastRow_AdjustingUnfiltered As Long
            Dim LastRow_SchoolValidation As Long

        ' ---------------------------------------------
        '      USER-REQUIRED ADJUSTMENTS SECTION ROWS
        ' ---------------------------------------------
            Dim SectionHeaderRow_UserRequiredAdjustments As Long
            Dim HeaderRow_UserRequiredAdjustments As Long
            Dim DataStartRow_UserRequiredAdjustments As Long
            Dim LastRow_UserRequiredAdjustments As Long

            Dim DataStartRow_UserRequiredAdjustments_BankAllocations As Long
            Dim LastRow_UserRequiredAdjustments_BankAllocations As Long
            
            Dim DataStartRow_UserRequiredAdjustments_MissingSchoolNames As Long
            Dim LastRow_UserRequiredAdjustments_MissingSchoolNames As Long
            
            Dim DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments As Long
            Dim LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments As Long
            
            Dim DataStartRow_UserRequiredAdjustments_GrossAmountVariances As Long
            Dim LastRow_UserRequiredAdjustments_GrossAmountVariances As Long
            
            Dim DataStartRow_UserRequiredAdjustments_MissingPaymentIDs As Long
            Dim LastRow_UserRequiredAdjustments_MissingPaymentIDs As Long

        ' ---------------------------------------------
        '     USER-REQUIRED ADJUSTMENTS RANGE STRINGS
        ' ---------------------------------------------
            ' These store workbook-style range addresses that are created later
            ' after each exception section is populated.
            ' They are used in downstream formulas to exclude unresolved records
            ' from the filtered import paths.
            Dim Rng_UserRequiredAdjustments_BankAllocations As String
            Dim Rng_UserRequiredAdjustments_MissingSchoolNames As String
            Dim Rng_UserRequiredAdjustments_AccountDivisionFundingAdjustments As String
            Dim Rng_UserRequiredAdjustments_GrossAmountVariances As String
            Dim Rng_UserRequiredAdjustments_MissingPaymentIDs As String
            
            Dim Rng_SchoolValidation_SchoolNames As String

        ' ---------------------------------------------
        '       FILE-MOVE / PROCESS-LATER VARIABLES
        ' ---------------------------------------------
            ' These variables are used near the end of the converter to build a
            ' unique list of source files that should be moved into a
            ' "Process Later" folder.
            Dim dictFilesToMove As Object
            Dim FilesToMove() As String
            
            Dim FileName_ToMove As String
            Dim UniqueFileName_FromDictionary As Variant
            Dim FileMoveIndex As Long
            Dim UserRequiredAdjustments_Row As Long

        ' ---------------------------------------------
        '          BUTTON / SHAPE OBJECT VARIABLES
        ' ---------------------------------------------
            Dim DonationSiteButton As Button
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CONFIGURE EXCEL ENVIRONMENT ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Temporarily disable Excel interface features to improve performance and prevent prompts while the converter runs.
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
    ' The following options can be enabled if additional performance improvements are needed. They are currently disabled to avoid interfering with other workbook processes.
        'Application.EnableEvents = False
        
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' DETERMINE CONVERTER STARTING POINT ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "Checking For Existing Worksheets"

    ' Set 'InitialExists' to False before checking starting point.
        InitialExists = False
    
    ' Loop through each worksheet to determine whether an Initial Data worksheet already exists in the macro workbook.
    ' If found:
      ' 1. Set InitialExists = True
      ' 2. Identify whether the route is Intacct or Salesforce
      ' 3. Assign wsInitialData to the matching worksheet
      ' 4. Assign JournalType only if it has not already been manually set
        For Each wsCheck In wbMacro.Worksheets
        ' Make the worksheet visible so any existing hidden sheets can be accessed later.
            wsCheck.Visible = xlSheetVisible
    
            If wsCheck.Name = "Initial Data - Intacct" Then
                InitialExists = True
                InitialPath = "Intacct"
    
                If JournalType = "" Then
                    JournalType = "Adjusting"
                End If
    
                Set wsInitialData = wsCheck
                Exit For
    
            ElseIf wsCheck.Name = "Initial Data - SF" Then
                InitialExists = True
                InitialPath = "Salesforce"
    
                If JournalType = "" Then
                    JournalType = "CRJ"
                End If
    
                Set wsInitialData = wsCheck
                Exit For
            End If
        Next wsCheck

    ' If the Initial Data worksheet already exists, skip the Initial Report import process and continue directly to donation site consolidation.
        If InitialExists Then
            Application.DisplayAlerts = False
            Application.ScreenUpdating = False
            GoTo Add_ConsolidatedReports
        End If

    ' Reset InitialPath so it can be assigned later from the selected report.
        InitialPath = ""
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' PRE-RUN CHECKLIST AND CONFIRMATION OF USER READINESS '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "Pre-Run Checklist and Confirmation of User Readiness"
        
    ' Display a pre-run checklist outlining all required information the user must have available before starting the converter.
        UserResponse = MsgBox( _
                "Before starting, please confirm you have the following:" & vbCrLf & vbCrLf & _
                    "1. A report downloaded from either Intacct or Salesforce." & vbCrLf & _
                    "2. All donation site reports downloaded and placed in a folder with '" & DonationSite & "' in the folder's name." & vbCrLf & vbCrLf & _
                "Are you ready to continue?", _
                vbYesNo + vbQuestion, _
                "Your Cause - AR Converter Confirmation")

    ' If the user indicates they are not ready, end the macro immediately.
        If UserResponse = vbNo Then
            Application.StatusBar = False
            Exit Sub
        End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''----------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' RESET WORKBOOK ENVIRONMENT ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''----------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "Resetting Workbook Environment"
        
    ' Clear the workbook using the Reset.Create_Reset_Worksheet procedure to prepare a clean environment for the converter.
        Reset.Create_Reset_Worksheet


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CONFIGURE EXCEL ENVIRONMENT ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
    
    
    ' Temporarily disable Excel interface features to improve performance and prevent prompts while the converter runs.
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        
    ' The following options can be enabled if additional performance improvements are needed. They are currently disabled to avoid interfering with other workbook processes.
        'Application.EnableEvents = False
        'Application.Calculation = xlCalculationManual


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CONSOLIDATION ONLY MODE: CHECK ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' If consolidation-only mode is enabled, skip the remainder of the Converter setup and proceed directly to report consolidation.
        If AllowConsolidationOnly Then
            GoTo ConsolidationOnly
        End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' IMPORT INITIAL REPORT AND DETERMINE REPORT TYPE ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
    
    ' ============================================================
    '                      USER FILE SELECTION
    ' ============================================================
        ' Update the Status Bar
            Application.StatusBar = "Requesting Initial (Intacct or Salesforce) Report from User"
            
        ' Open a file picker so the user can select the Initial Report.
            Set fdFilePath_InitialReport = Application.FileDialog(msoFileDialogFilePicker)
        
            With fdFilePath_InitialReport
                .Title = "Select the Initial (Intacct or Salesforce) Report"
                .AllowMultiSelect = False
                
                .Filters.Clear
                .Filters.Add "Excel Files", "*.xlsx; *.xls; *.csv"
            
            ' If the user cancels the file selection dialog, prepare an exit message and stop the converter.
                If .Show <> -1 Then
                    ExitMessage = "No file selected. Please locate the Intacct or Salesforce Report and try again."
                    ExitMessage_Title = "No File Selected"
                    GoTo CompleteMacro
                End If
            
                FilePath_InitialReport = .SelectedItems(1)
            End With
    
    ' ============================================================
    '              OPEN USER SELECTED 'INITIAL REPORT'
    ' ============================================================
        ' Open the selected Initial Report.
            Set wbTemp_InitialReport = Workbooks.Open(FilePath_InitialReport, ReadOnly:=True)
            
        ' Use the first worksheet in the selected workbook for validation.
            Set wsTemp_InitialReport = wbTemp_InitialReport.Worksheets(1)
            
    ' ============================================================
    '             VALIDATE USER SELECTED INITIAL REPORT
    ' ============================================================
      ' Determine whether the selected report is an Intacct or Salesforce report by comparing column headers against the expected header arrays.
      ' If header search is enabled, each row is tested until a match is found. Otherwise, the assigned header row is used.
          
        ' Update the Status Bar
            Application.StatusBar = "Validating Initial Report"
            
        ' Find the last row of the Initial Report.
            TempLastRow_InitialReport = wsTemp_InitialReport.Cells(wsTemp_InitialReport.Rows.Count, 1).End(xlUp).Row
    
        ' ---------------------------------------------
        '       INITIAL HEADER SEARCH: ON VS. OFF
        ' ---------------------------------------------
            If AllowHeaderRowSearch_InitialReport Then
                
                For HeaderRow_InitialReport = 1 To TempLastRow_InitialReport
                    
                    ' Reset InitialPath so each row is evaluated independently.
                        InitialPath = ""
                    
                    ' ..............................
                    '             INTACCT
                    ' ..............................
                        ' Check whether the row matches the expected Intacct column headers.
                            For ColumnCheck_InitialReport = LBound(ColumnHeaders_Initial_Intacct) To UBound(ColumnHeaders_Initial_Intacct)
                                If wsTemp_InitialReport.Cells(HeaderRow_InitialReport, ColumnCheck_InitialReport + 1).Value <> _
                                   ColumnHeaders_Initial_Intacct(ColumnCheck_InitialReport) Then
                                    Exit For
                                End If
                                
                                If ColumnCheck_InitialReport = UBound(ColumnHeaders_Initial_Intacct) Then
                                    InitialPath = "Intacct"
                                End If
                            Next ColumnCheck_InitialReport
                    
                    ' ..............................
                    '           SALESFORCE
                    ' ..............................
                        ' If the row does not match Intacct headers, check for Salesforce headers.
                            If InitialPath = "" Then
                                For ColumnCheck_InitialReport = LBound(ColumnHeaders_Initial_Salesforce) To UBound(ColumnHeaders_Initial_Salesforce)
                                    If wsTemp_InitialReport.Cells(HeaderRow_InitialReport, ColumnCheck_InitialReport + 1).Value <> _
                                       ColumnHeaders_Initial_Salesforce(ColumnCheck_InitialReport) Then
                                        Exit For
                                    End If
                                    
                                    If ColumnCheck_InitialReport = UBound(ColumnHeaders_Initial_Salesforce) Then
                                        InitialPath = "Salesforce"
                                    End If
                                Next ColumnCheck_InitialReport
                            End If
                    
                    ' If a valid header row was found, stop searching.
                        If InitialPath <> "" Then
                            Exit For
                        End If
                    
                Next HeaderRow_InitialReport
                
            Else
                ' Use the assigned header row instead of searching every row.
                    HeaderRow_InitialReport = AssignedHeaderRow_InitialReport
                    
                    InitialPath = ""
            
                ' ..............................
                '             INTACCT
                ' ..............................
                    ' Check whether the assigned row matches the expected Intacct column headers.
                        For ColumnCheck_InitialReport = LBound(ColumnHeaders_Initial_Intacct) To UBound(ColumnHeaders_Initial_Intacct)
                            If wsTemp_InitialReport.Cells(HeaderRow_InitialReport, ColumnCheck_InitialReport + 1).Value <> _
                               ColumnHeaders_Initial_Intacct(ColumnCheck_InitialReport) Then
                                Exit For
                            End If
                            
                            If ColumnCheck_InitialReport = UBound(ColumnHeaders_Initial_Intacct) Then
                                InitialPath = "Intacct"
                            End If
                        Next ColumnCheck_InitialReport
            
                ' ..............................
                '           SALESFORCE
                ' ..............................
                    ' If the assigned row does not match Intacct headers, check for Salesforce headers.
                        If InitialPath = "" Then
                            For ColumnCheck_InitialReport = LBound(ColumnHeaders_Initial_Salesforce) To UBound(ColumnHeaders_Initial_Salesforce)
                                If wsTemp_InitialReport.Cells(HeaderRow_InitialReport, ColumnCheck_InitialReport + 1).Value <> _
                                   ColumnHeaders_Initial_Salesforce(ColumnCheck_InitialReport) Then
                                    Exit For
                                End If
                                
                                If ColumnCheck_InitialReport = UBound(ColumnHeaders_Initial_Salesforce) Then
                                    InitialPath = "Salesforce"
                                End If
                            Next ColumnCheck_InitialReport
                        End If
            End If
    
    ' ============================================================
    '                  CONFIRM INITIAL REPORT TYPE
    ' ============================================================
        ' If no valid InitialPath was identified, stop the converter.
            If InitialPath = "" Then
                ExitMessage = "The report does not appear to be a valid Intacct or Salesforce report." & vbCrLf & _
                              "If this is an error, please reach out to " & CurrentVBACodeMaintainer & " to further assist in this process."
                ExitMessage_Title = "Invalid Initial Report"
                GoTo CompleteMacro
            ElseIf InitialPath = "Salesforce" Then
                GoTo InitialPath_SF
            ElseIf InitialPath = "Intacct" Then
                ' Continue through the default Intacct path.
            End If
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' PROCESS INITIAL REPORT ROUTE '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "Copying over the Initial Report"
    
    ' ============================================================
    '                    INITIAL DATA - INTACCT
    ' ============================================================
        ' Create a worksheet to hold the Initial 'Intacct Report' Data to be used later in the converter.
            Set wsInitialData = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets(wbMacro.Worksheets.Count))
            wsInitialData.Name = "Initial Data - Intacct"
            
        ' Copy over the Initial 'Intacct' Report into the 'wsInitialData' Worksheet.
            wsTemp_InitialReport.Range("A" & HeaderRow_InitialReport & ":AC" & TempLastRow_InitialReport).Copy wsInitialData.Range("A1")
        
        ' Format the 'wsInitialData' worksheet.
            wsInitialData.Range("A1:AC1").AutoFilter
            wsInitialData.Range("A:AC").WrapText = False
            wsInitialData.Columns("A:AC").AutoFit
            
        ' Close the 'wbTemp_InitialReport' workbook without saving changes.
            wbTemp_InitialReport.Close SaveChanges:=False
        
        ' If the JournalType is not already assigned, assign it to the "Adjusting" path.
            If JournalType = "" Then
                JournalType = "Adjusting"
            End If
        
        ' Jump over the 'INITIAL DATA - SALESFORCE' section into the 'IMPORT DONATION SITE REPORTS' section.
            GoTo Add_ConsolidatedReports

    ' ============================================================
    '                   INITIAL DATA - SALESFORCE
    ' ============================================================
InitialPath_SF:
        ' Create a worksheet to hold the Initial 'Salesforce Report' Data to be used later in the converter.
            Set wsInitialData = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets(wbMacro.Worksheets.Count))
            wsInitialData.Name = "Initial Data - SF"
            
        ' Copy over the Initial 'Salesforce' Report into the 'wsInitialData' Worksheet.
            wsTemp_InitialReport.Range("A" & HeaderRow_InitialReport & ":T" & TempLastRow_InitialReport).Copy wsInitialData.Range("A1")
        
        ' Format the 'wsInitialData' worksheet.
            wsInitialData.Range("A1:T1").AutoFilter
            wsInitialData.Range("A:T").WrapText = False
            wsInitialData.Columns("A:T").AutoFit
            
        ' Close the 'wbTemp_InitialReport' workbook without saving changes.
            wbTemp_InitialReport.Close SaveChanges:=False
        
        ' If the JournalType is not already assigned, assign it to the "CRJ" path.
            If JournalType = "" Then
                JournalType = "CRJ"
            End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' IMPORT DONATION SITE REPORTS '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Add_ConsolidatedReports:
    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = ""
    
    ' ============================================================
    '                     USER FOLDER SELECTION
    ' ============================================================
        ' ---------------------------------------------
        '              USER SELECTS FOLDER
        ' ---------------------------------------------
            ' Open a folder picker so the user can select the folder containing the Donation Site Reports to process.
                Set fdFolderPath_DonationSite = Application.FileDialog(msoFileDialogFolderPicker)
                                
                With fdFolderPath_DonationSite
                    .Title = "Select the Your Cause Reports Folder"
                    .AllowMultiSelect = False
                                    
                ' If the user cancels the folder selection dialog, prepare an exit message and stop the converter.
                    If .Show <> -1 Then
                        ExtraMessage = "No Folder Selected. Please locate the correct folder and try again."
                        ExtraMessage_Title = "No Folder Selected"
                        GoTo CreateButton_Step2
                    End If
                             
                ' Store the selected folder path for use later in the converter.
                    FolderPath_DonationSite = .SelectedItems(1)
                End With
    
        ' ---------------------------------------------
        '       VALIDATE FOLDER NAMING CONVENTION
        ' ---------------------------------------------
            ' This validation helps ensure the user intentionally selected a folder created specifically for the converter to process the Donation Site Reports.
              ' Verify the selected folder path contains "Your Cause" or "YourCause" in the folder name.
                If (InStr(1, FolderPath_DonationSite, "YourCause", vbTextCompare) = 0) And (InStr(1, FolderPath_DonationSite, "Your Cause", vbTextCompare) = 0) Then
                    ExtraMessage = "The selected folder path does not contain '" & DonationSite & "' in the folder name. " & _
                                   "Please rename the folder or locate the correct folder and try again." & vbCrLf & vbCrLf & _
                                   "If this error persists, please contact " & CurrentVBACodeMaintainer & " to further assist in the process."
                    ExtraMessage_Title = "Missing dedicated folder naming convention"
                    GoTo CompleteMacro
                    
                End If
        
        ' ---------------------------------------------
        '          VERIFY FOLDER CONTAINS FILES
        ' ---------------------------------------------
            ' Build a list of files in the selected folder and confirm that at least one file exists.
                FileCount_DonationSite = 0
            
            ' Retrieve the first file name from the selected folder path.
                FileName_DonationSite = Dir(FolderPath_DonationSite & "\*.*", vbNormal Or vbReadOnly Or vbHidden Or vbSystem)
            
            ' Loop through every file returned by Dir and build a list of file names that will later be used for processing the Donation Site Reports.
                Do While Len(FileName_DonationSite) > 0
                    ' Increase the total file count.
                        FileCount_DonationSite = FileCount_DonationSite + 1
                        
                    ' Expand the file list array to store the current file name.
                        ReDim Preserve FileNamesList_DonationSite(1 To FileCount_DonationSite)
                        
                    ' Store the file name in the array.
                        FileNamesList_DonationSite(FileCount_DonationSite) = FileName_DonationSite
                        
                    ' Retrieve the next file in the folder.
                        FileName_DonationSite = Dir()
                Loop
            
            ' If no files were found in the selected folder, stop the converter.
                If FileCount_DonationSite = 0 Then
                    ExtraMessage = "No files were found in the selected folder. Please locate the correct folder and try again."
                    ExtraMessage_Title = "No Files Found"
                    GoTo CompleteMacro
                End If
    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = ""
        
    ' ============================================================
    '       SET UP THE 'wsConsolidatedData' WORKSHEET ENVIRONMENT
    ' ============================================================
        ' ---------------------------------------------
        '    CREATE THE 'wsConsolidatedData' WORKSHEET
        ' ---------------------------------------------
            ' Create a new worksheet to hold all consolidated Donation Site Report data.
                Set wsConsolidatedData = wbMacro.Worksheets.Add(After:=wsInitialData)
                wsConsolidatedData.Name = "Consolidated Reports"

        ' ---------------------------------------------
        '              ADD COLUMN HEADERS
        ' ---------------------------------------------
            ' Add the Donation Site Report columns.
                wsConsolidatedData.Range("A1:AO1").Value = ColumnHeaders_YourCause

            ' Add additional columns used internally by the converter to track the source file and worksheet output.
                wsConsolidatedData.Range("AP1:AQ1").Value = Array("File Name", "Worksheet Name")
            

    ' ============================================================
    '               IMPORT DONATION SITE REPORT DATA
    ' ============================================================
        ' Loop through each file in the selected Donation Site folder and import usable report data into wsConsolidatedData.
            For FileNumber_DonationSite = LBound(FileNamesList_DonationSite) To UBound(FileNamesList_DonationSite)
                ' ---------------------------------------------
                '     UPDATE THE STATUS BAR AND PROGRESS BAR
                ' ---------------------------------------------
                    ' Update the Status Bar to display the current file number, total file count, and file name being processed.
                        Application.StatusBar = "Processing file " & _
                                                (FileNumber_DonationSite - LBound(FileNamesList_DonationSite) + 1) & _
                                                " of " & _
                                                (UBound(FileNamesList_DonationSite) - LBound(FileNamesList_DonationSite) + 1) & _
                                                ": " & _
                                                FileNamesList_DonationSite(FileNumber_DonationSite)
    
                ' ---------------------------------------------
                '            DETERMINE FILE USABILITY
                ' ---------------------------------------------
                    ' ..............................
                    '     FILE TYPE DETERMINATION
                    ' ..............................
                        ' Check whether the file is a supported Excel or CSV file.
                            If Not LCase$(FileNamesList_DonationSite(FileNumber_DonationSite)) Like "*.csv" _
                              And Not LCase$(FileNamesList_DonationSite(FileNumber_DonationSite)) Like "*.xls*" Then
                                GoTo DoNotUseFile
                            End If
                            
                    ' ..............................
                    '   OPEN FILE AND SET VARIABLES
                    ' ..............................
                        ' Open the current file as read-only and assign it to wbTemp_DonationSite.
                            Set wbTemp_DonationSite = Workbooks.Open(FolderPath_DonationSite & "\" & FileNamesList_DonationSite(FileNumber_DonationSite), ReadOnly:=True)
                        
                        ' Assign the first worksheet in the workbook to wsTemp_DonationSite.
                            Set wsTemp_DonationSite = wbTemp_DonationSite.Worksheets(1)
                    
                        ' Determine the last row of data in wsTemp_DonationSite.
                            LastRow_TempDonationSite = wsTemp_DonationSite.Cells(wsTemp_DonationSite.Rows.Count, "A").End(xlUp).Row
                        
                    ' ..............................
                    '      DETERMINE HEADER ROW
                    ' ..............................
                        ' Determine which row contains the Donation Site Report headers based on the header search setting.
                            If AllowHeaderRowSearch_DonationSiteReports Then
                                ' HEADER ROW SEARCH ENABLED
                                  ' Loop through each row and look for a full match against the expected Donation Site Report headers.
                                    For CurrentRow_DonationSite = 1 To LastRow_TempDonationSite
                                      
                                        ' First confirm the first expected header appears in column A before checking the full row.
                                            If StrComp(CStr(wsTemp_DonationSite.Cells(CurrentRow_DonationSite, "A").Value), _
                                              ColumnHeaders_YourCause(0), vbTextCompare) = 0 Then
                                          
                                                ' Reset the header match counter for the current row.
                                                    ColumnMatch_DonationSite = 0
                                          
                                                ' Compare each expected header against the current row.
                                                    For Col_DonationSite = 0 To UBound(ColumnHeaders_YourCause)
                                                        If StrComp(CStr(wsTemp_DonationSite.Cells(CurrentRow_DonationSite, Col_DonationSite + 1).Value), _
                                                          ColumnHeaders_YourCause(Col_DonationSite), vbTextCompare) = 0 Then
                                                          
                                                            ColumnMatch_DonationSite = ColumnMatch_DonationSite + 1
                                                        Else
                                                            Exit For
                                                        End If
                                                    Next Col_DonationSite
                                          
                                                ' If all expected headers match, store the header row and process the file.
                                                    If ColumnMatch_DonationSite = UBound(ColumnHeaders_YourCause) + 1 Then
                                                        HeaderRow_DonationSite = CurrentRow_DonationSite
                                                        GoTo UseFile
                                                    End If
                                            End If
                                      
                                    Next CurrentRow_DonationSite
                                    
                            Else
                                ' HEADER ROW SEARCH DISABLED
                                  ' Validate only the assigned header row instead of searching through every row in the worksheet.
                                    If StrComp(CStr(wsTemp_DonationSite.Cells(AssignedHeaderRow_DonationSiteReports, "A").Value), _
                                      ColumnHeaders_YourCause(0), vbTextCompare) = 0 Then
                                            
                                        ' Reset the header match counter for the assigned row.
                                            ColumnMatch_DonationSite = 0
                                        
                                        ' Compare each expected header against the assigned row.
                                            For Col_DonationSite = 0 To UBound(ColumnHeaders_YourCause)
                                                If StrComp(CStr(wsTemp_DonationSite.Cells(AssignedHeaderRow_DonationSiteReports, Col_DonationSite + 1).Value), _
                                                  ColumnHeaders_YourCause(Col_DonationSite), vbTextCompare) = 0 Then
                                                    ColumnMatch_DonationSite = ColumnMatch_DonationSite + 1
                                                Else
                                                    Exit For
                                                End If
                                            Next Col_DonationSite
                                        
                                        ' If all expected headers match, store the assigned header row and process the file.
                                            If ColumnMatch_DonationSite = UBound(ColumnHeaders_YourCause) + 1 Then
                                            
                                                HeaderRow_DonationSite = AssignedHeaderRow_DonationSiteReports
                                                GoTo UseFile
                                            End If
                                    End If
                                
                            End If
                            
                    ' ..............................
                    '   HEADER ROW: NOT DETERMINED
                    ' ..............................
                        ' If no valid Donation Site Report headers are found, skip the file.
                            FileCount_WrongReport = FileCount_WrongReport + 1
                            GoTo DoNotUseFile
                                    
                    ' Process files with valid Donation Site Report headers.
UseFile:
                    ' ..............................
                    '  DETERMINE NUMBER OF DATA ROWS
                    ' ..............................
                        ' Determine the number of data rows.
                            DataRows_DonationSite = LastRow_TempDonationSite - HeaderRow_DonationSite
    
                        ' Determine the total number of usable data rows after accounting for 'RowsToDeleteFromBottomOfDonationSiteReport'.
                            DataRows_DonationSite_Total = DataRows_DonationSite - RowsToDeleteFromBottomOfDonationSiteReport
                        
                        ' Ensure the report contains usable data after accounting for 'RowsToDeleteFromBottomOfDonationSiteReport'.
                          ' If not, skip the file.
                            If DataRows_DonationSite_Total < 1 Then
                                FileCount_DonationSite_Unusable = FileCount_DonationSite_Unusable + 1
                                GoTo DoNotUseFile
                            End If
                    
                        ' Use HeaderRow_DonationSite to determine the start row of the data.
                            DataStartRow_DonationSite = HeaderRow_DonationSite + 1
                            
                        ' Determine the adjusted last row after accounting for 'RowsToDeleteFromBottomOfDonationSiteReport'.
                            LastRow_TempDonationSite_Adjusted = LastRow_TempDonationSite - RowsToDeleteFromBottomOfDonationSiteReport
    
                ' ---------------------------------------------
                '      COPY DONATION SITE REPORT DATA INTO
                '        THE CONSOLIDATED DATA WORKSHEET
                ' ---------------------------------------------
                    ' Find the next available row in wsConsolidatedData using column A.
                        LastRow_ConsolidatedData = wsConsolidatedData.Cells(wsConsolidatedData.Rows.Count, "A").End(xlUp).Row + 1
                
                    ' Copy the wsTemp_DonationSite data into wsConsolidatedData.
                        wsTemp_DonationSite.Range("A" & DataStartRow_DonationSite & ":AO" & LastRow_TempDonationSite_Adjusted).Copy _
                                Destination:=wsConsolidatedData.Range("A" & LastRow_ConsolidatedData)
                
                    ' Build the worksheet name used for documentation/reference.
                        WorksheetName = Format(wsTemp_DonationSite.Range("M" & DataStartRow_DonationSite).Value, "YYYY.MM.DD") & " (" & _
                                wsTemp_DonationSite.Range("J" & DataStartRow_DonationSite).Value & ")"
                
                    ' Find the new last row so the tracking fields can be written to the correct consolidated data rows.
                        LastRow_ConsolidatedData_AfterInsert = wsConsolidatedData.Cells(wsConsolidatedData.Rows.Count, "A").End(xlUp).Row
                
                    ' Add the original file name and worksheet name to the tracking columns.
                        wsConsolidatedData.Range("AP" & LastRow_ConsolidatedData & ":AP" & LastRow_ConsolidatedData_AfterInsert).Value = "" & _
                                FileNamesList_DonationSite(FileNumber_DonationSite)
                                
                        wsConsolidatedData.Range("AQ" & LastRow_ConsolidatedData & ":AQ" & LastRow_ConsolidatedData_AfterInsert).Value = WorksheetName
                
                    ' Clear the clipboard.
                        Application.CutCopyMode = False
                        
                ' ---------------------------------------------
                '           COPY DONATION SITE REPORT
                '          INTO THE 'wbMacro' WORKBOOK
                ' ---------------------------------------------
                    ' Format the 'wsTemp_DonationSite' worksheet.
                        wsTemp_DonationSite.Columns("A:AO").AutoFit
                        
                    ' Copy the original Donation Site Report worksheet into 'wbMacro' for documentation/reference.
                        If IncludeOriginalReports Then
    
                            ' Copy the original Donation Site worksheet into wbMacro.
                                wsTemp_DonationSite.Copy After:=wbMacro.Sheets(wbMacro.Sheets.Count)
                            
                            ' Reference the newly copied worksheet.
                                Set wsNew = wbMacro.Sheets(wbMacro.Sheets.Count)
                            
                            ' Attempt to rename the worksheet.
                              ' If the name already exists or is invalid, skip renaming.
                                On Error Resume Next
                                    wsNew.Name = WorksheetName
                                On Error GoTo 0
                        
                        End If
                        
                    ' Clear the clipboard.
                        Application.CutCopyMode = False
                        
                ' ---------------------------------------------
                '             UPDATE USED FILE COUNT
                ' ---------------------------------------------
                    ' Increase the used file count.
                        FileCount_DonationSite_Used = FileCount_DonationSite_Used + 1
        
DoNotUseFile:
                ' ---------------------------------------------
                '             CLEAN UP AND CONTINUE
                ' ---------------------------------------------
                    ' Close the temporary workbook without saving changes.
                        On Error Resume Next
                        If Not wbTemp_DonationSite Is Nothing Then
                            wbTemp_DonationSite.Close SaveChanges:=False
                        End If
                        On Error GoTo 0
    
            Next FileNumber_DonationSite

    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = ""
        
    ' ============================================================
    '                      VALIDATE FILE COUNT
    ' ============================================================
        ' If no files were imported into wsConsolidatedData, stop the converter.
            If FileCount_DonationSite_Used = 0 Then
                ExtraMessage = "The selected folder did not contain any usable '" & DonationSite & "' files. " & _
                        "Please find the correct folder and try again."
                ExtraMessage_Title = "No Usable Files Found"
                wsConsolidatedData.Delete
                GoTo CreateButton_Step2
            End If
    
    ' ============================================================
    '             DELETE THE BUTTON WORKSHEET (IF USED)
    ' ============================================================
        ' If the "No Donation Site Report" worksheet exists, delete it.
            For Each wsButton In wbMacro.Worksheets
                If wsButton.Name = "No Donation Site Report" Then
                    wsButton.Delete
                End If
            Next wsButton
    
    ' ============================================================
    '            FORMAT THE CONSOLIDATED DATA WORKSHEET
    ' ============================================================
        ' Format wsConsolidatedData.
          ' Apply AutoFilter and AutoFit to the columns.
            wsConsolidatedData.Range("A1:AQ1").AutoFilter
            wsConsolidatedData.Columns("A:AQ").AutoFit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''----------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CREATE ALL ADDITIONAL WORKSHEETS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''----------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "Creating all additional worksheets"
        
    ' ============================================================
    '            STANDARDIZED SALESFORCE DATA WORKSHEET
    ' ============================================================
        ' This worksheet standardizes the Initial Report data into a Salesforce-based structure.
        ' This allows the converter to work from one consistent layout later, regardless of whether the Initial Report came from Intacct or Salesforce.
            Set wsStandardizedSF = wbMacro.Worksheets.Add(After:=wsInitialData)
            wsStandardizedSF.Name = "Standardized Salesforce"
            
    ' ============================================================
    '           STANDARDIZED DONATION SITE DATA WORKSHEET
    ' ============================================================
        ' This worksheet standardizes the Donation Site Report data into one consistent structure.
        ' This allows later sections of the converter to work from one Donation Site layout instead of relying on source-specific column arrangements.
            Set wsStandardizedDonationSiteData = wbMacro.Worksheets.Add(After:=wsConsolidatedData)
            wsStandardizedDonationSiteData.Name = "Standardized Donation Site Data"
            
    ' ============================================================
    '                  DISBURSEMENT DATA WORKSHEET
    ' ============================================================
        ' This worksheet groups related transactions into their corresponding disbursements.
        ' This is needed so the converter can summarize activity at the disbursement level and use that information later for Fees, Bank Deposits, and import-file creation.
            Set wsDisbursementData = wbMacro.Worksheets.Add(After:=wsStandardizedDonationSiteData)
            wsDisbursementData.Name = "Disbursement Data"
            
    ' ============================================================
    '                RELEVANT TRANSACTIONS WORKSHEET
    ' ============================================================
        ' This worksheet connects Donation Site Report data to Salesforce data.
        ' It exists to keep only the Donation Site transactions that are relevant for the converter based on what is first found in Salesforce.
            Set wsRelevantTransactions = wbMacro.Worksheets.Add(After:=wsDisbursementData)
            wsRelevantTransactions.Name = "Relevant Transactions"
            
    ' ============================================================
    '                        FEES WORKSHEET
    ' ============================================================
        ' This worksheet isolates the fee portion of each disbursement.
        ' This is needed so fees can be separated from donation amounts and later used as their own line items in the import-file process.
            Set wsFees = wbMacro.Worksheets.Add(After:=wsRelevantTransactions)
            wsFees.Name = "Fees"
    
    ' ============================================================
    '            BANK DEPOSITS WORKSHEET (IF APPLICABLE)
    ' ============================================================
        ' This worksheet is used only for the "Adjusting" JournalType.
        ' It exists because Adjusting journals require a bank deposit line item, while CRJs do not.
        ' The worksheet uses the net disbursement amount to create the bank deposit line item needed later in the Adjusting Journal path.
            Set wsBankDeposits = wbMacro.Worksheets.Add(After:=wsFees)
            wsBankDeposits.Name = "Bank Deposits"
            
    ' ============================================================
    '                 CONNECTION ANALYSIS WORKSHEET
    ' ============================================================
        ' This worksheet brings together the fields needed to evaluate whether Salesforce data and Donation Site Report data are connected as expected.
        ' It exists to identify variances before creating the final import file and to feed those variances into the User-Required Adjustments worksheet.
            Set wsConnectionAnalysis = wbMacro.Worksheets.Add(After:=wsBankDeposits)
            wsConnectionAnalysis.Name = "Connection Analysis"
            
    ' ============================================================
    '              USER-REQUIRED ADJUSTMENTS WORKSHEET
    ' ============================================================
        ' This worksheet gives the user one centralized place to review variances and make corrections without needing to re-run the converter.
        ' It exists so unresolved issues can be handled directly in the workbook before the final import file is created.
        ' It provides the following checks:
          ' BANK ALLOCATIONS NOT FOUND
          ' TRANSACTIONS MISSING SCHOOL NAME
          ' ADJUSTMENTS TO: ACCOUNT|DIVISION|FUNDING SOURCE
          ' DONATION SITE VS SALESFORCE: GROSS AMOUNTS MISMATCHES
          ' TRANSACTIONS MISSING PMT-IDS
            Set wsUserRequiredAdjustments = wbMacro.Worksheets.Add(After:=wsConnectionAnalysis)
            wsUserRequiredAdjustments.Name = "User-Required Adjustments"
    
    ' ============================================================
    '              SCHOOL VALIDATION WORKSHEET
    ' ============================================================
        ' This worksheet stores the approved school validation lists used for dropdown selections throughout the converter.
        ' Defining and preparing it here allows the same validation source to be reused across multiple exception sections without rebuilding it each time.
    
        ' Run the macro that creates or refreshes the School Validation worksheet.
            School_Validation.Validation
    
        ' Assign the worksheet to the variable for reuse throughout the macro.
            Set wsSchoolValidation = wbMacro.Worksheets("School Validation")
    
        ' Determine the last populated row using column B, which is expected to always contain values.
        ' Column C holds the school names that will be used for data validation dropdowns.
            LastRow_SchoolValidation = wsSchoolValidation.Cells(wsSchoolValidation.Rows.Count, 2).End(xlUp).Row
    
        ' Store the validation range (school names in column C) so it can be reused later without recalculating the range for each section.
            Rng_SchoolValidation_SchoolNames = "='" & wsSchoolValidation.Name & "'!$C$2:$C$" & LastRow_SchoolValidation
    
        ' Hide the worksheet to keep it out of the user's workflow while still allowing it to serve as a backend validation source.
            wsSchoolValidation.Visible = xlSheetHidden
            
    ' ============================================================
    '                INTACCT IMPORT FILE WORKSHEETS
    ' ============================================================
      ' This section creates the worksheets needed for the final import-file path.
      ' The converter creates a different set of worksheets depending on whether the JournalType is "Adjusting" or "CRJ".
                
        ' ---------------------------------------------
        '         JOURNALTYPE: "ADJUSTING" PATH
        ' ---------------------------------------------
        If JournalType = "Adjusting" Then
            ' ..............................
            '      UNFILTERED LINE ITEMS
            ' ..............................
                ' This worksheet shows all line items that could flow into the Adjusting Journal path.
                ' It includes matched Salesforce and Donation Site data, adjustments, fees, bank deposits, and unresolved transactions that still need user attention.
                    Set wsAdjustingUnfiltered = wbMacro.Worksheets.Add(After:=wsUserRequiredAdjustments)
                    wsAdjustingUnfiltered.Name = "Adjusting Journal - Unfiltered"
                    
            ' ..............................
            '       FILTERED LINE ITEMS
            ' ..............................
                ' This worksheet removes line items tied to unresolved issues found in the User-Required Adjustments worksheet.
                ' It updates as the user makes corrections, allowing the converter to reflect only the currently usable Adjusting Journal lines.
                    Set wsAdjustingFiltered = wbMacro.Worksheets.Add(After:=wsAdjustingUnfiltered)
                    wsAdjustingFiltered.Name = "Adjusting Journal - Filtered"
                    
            ' ..............................
            '      FINALIZED LINE ITEMS
            ' ..............................
                ' This worksheet holds the finalized Adjusting Journal import data.
                ' It exists to present the final set of import-ready lines after filtering and user adjustments have been accounted for.
                    Set wsAdjustingJournal = wbMacro.Worksheets.Add(After:=wsAdjustingFiltered)
                    wsAdjustingJournal.Name = "Adjusting Journal Import"

        ' ---------------------------------------------
        '            JOURNALTYPE: "CRJ" PATH
        ' ---------------------------------------------
        Else
            ' ..............................
            '      UNFILTERED LINE ITEMS
            ' ..............................
                ' This worksheet shows all line items that could flow into the CRJ path.
                ' It includes matched Salesforce and Donation Site data, adjustments, fees, and unresolved transactions that still need user attention.
                    Set wsCRJUnfiltered = wbMacro.Worksheets.Add(After:=wsUserRequiredAdjustments)
                    wsCRJUnfiltered.Name = "CRJ Unfiltered"
                    
            ' ..............................
            '       FILTERED LINE ITEMS
            ' ..............................
                ' This worksheet removes line items tied to unresolved issues found in the User-Required Adjustments worksheet.
                ' It updates as the user makes corrections, allowing the converter to reflect only the currently usable CRJ lines.
                    Set wsCRJFiltered = wbMacro.Worksheets.Add(After:=wsCRJUnfiltered)
                    wsCRJFiltered.Name = "CRJ Filtered"
                    
            ' ..............................
            '      FINALIZED LINE ITEMS
            ' ..............................
                ' This worksheet holds the finalized CRJ import data.
                ' It exists to present the final set of import-ready lines after filtering and user adjustments have been accounted for.
                    Set wsCRJ = wbMacro.Worksheets.Add(After:=wsCRJFiltered)
                    wsCRJ.Name = "CRJ Import"
        End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' STANDARDIZE INITIAL REPORT DATA (SALESFORCE DATA) '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "Standardizing the Initial Report"
        
    ' ============================================================
    '       FIND THE LAST ROW FROM THE INITIAL DATA WORKSHEET
    ' ============================================================
        ' Determine the last used row in wsInitialData so the formulas can reference the full Initial Report data range.
            LastRow_InitialData = wsInitialData.Cells(wsInitialData.Rows.Count, 1).End(xlUp).Row
    
    ' ============================================================
    '      POPULATE THE WORKSHEET BASED ON THE INITIAL REPORT
    ' ============================================================
      ' The Initial Report can come from either Intacct or Salesforce.
      ' This section standardizes both report types into the same Salesforce-based structure so later sections of the converter can use one consistent layout.
      
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Add the standardized column headers that will be used regardless of the Initial Report source.
                wsStandardizedSF.Range("A1:W1").Value = Array("Donation Site", "Transaction ID", "Disbursement ID", "Payment Type", "Check Number", _
                        "SF Payment ID", "Primary Contact", "Family Account Name", "Company Name", "Campaign Name", "Opportunity Name", "Memo", _
                        "Intacct - Location ID", "Intacct - Account", "Intacct - Division", "Intacct - Funding Source", _
                        "Intacct - Debt Services Series", "Amount", "Location Correction", "Account Correction", "Division Correction", _
                        "Funding Source Correction", "Debt Services Correction")
                        
        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' ..............................
            '    INITIAL REPORT: INTACCT
            ' ..............................
              ' If the Initial Report came from Intacct, pull the required fields directly into the standardized structure.
                If InitialPath = "Intacct" Then
                
                ' Columns A:W
                  ' A: "Donation Site".................... = Intacct "SF Donation Site" (Column D)
                  ' B: "Transaction ID"................... = Intacct "SF Transaction ID" (Column F)
                  ' C: "Disbursement ID".................. = Intacct "SF Disbursement ID" (Column G)
                  ' D: "Payment Type"..................... = Intacct "SF Payment Method" (Column H)
                  ' E: "Check Number"..................... = Intacct "SF Check Number" (Column I)
                  ' F: "SF Payment ID".................... = Intacct "SF Payment Number" (Column J)
                  ' G: "Primary Contact".................. = Intacct "SF Primary Contact" (Column K)
                  ' H: "Family Account Name".............. = Intacct "SF Account Name" (Column L)
                  ' I: "Company Name"..................... = Intacct "SF Company Name" (Column M)
                  ' J: "Campaign Name".................... = Intacct "SF Campaign Source" (Column N)
                  ' K: "Opportunity Name"................. = Intacct "SF Opportunity Name" (Column O)
                  ' L: "Memo"............................. = Intacct "Memo" (Column P)
                  ' M: "Intacct - Location ID"............ = Intacct "Location ID" (Column R)
                  ' N: "Intacct - Account"................ = Intacct "Account Number" (Column S)
                  ' O: "Intacct - Division"............... = Intacct "Division ID" (Column T)
                  ' P: "Intacct - Funding Source"......... = Intacct "Funding Source" (Column U)
                  ' Q: "Intacct - Debt Services Series"... = Intacct "Debt Services Series ID" (Column V)
                  ' R: "Amount"........................... = Intacct "Debit Amount" (Column AB)
                  ' S: "Location Correction".............. = Intacct "Location ID" (Column R)
                  ' T: "Account Correction"............... = Intacct "Account Number" (Column S)
                  ' U: "Division Correction".............. = Intacct "Division ID" (Column T)
                  ' V: "Funding Source Correction"........ = Intacct "Funding Source" (Column U)
                  ' W: "Debt Services Correction"......... = Intacct "Debt Services Series ID" (Column V)
                    wsStandardizedSF.Range("A2").Formula2 = "=IF(" & _
                        "ISBLANK(CHOOSECOLS('" & wsInitialData.Name & "'!A2:AC" & LastRow_InitialData & ",4,6,7,8,9,10,11,12,13,14,15,16,18,19,20,21,22,28,18,19,20,21,22)),""""," & _
                        "CHOOSECOLS('" & wsInitialData.Name & "'!A2:AC" & LastRow_InitialData & ",4,6,7,8,9,10,11,12,13,14,15,16,18,19,20,21,22,28,18,19,20,21,22))"
            
            ' ..............................
            '   INITIAL REPORT: SALESFORCE
            ' ..............................
              ' If the Initial Report came from Salesforce, fill the Salesforce-based fields directly and calculate the Intacct-related fields separately.
                ElseIf InitialPath = "Salesforce" Then
                
                ' Columns A:K
                  ' A: "Donation Site".................... = Salesforce "Donation Site" (Column D)
                  ' B: "Transaction ID"................... = Salesforce "Check/Reference Number" (Column F)
                  ' C: "Disbursement ID".................. = Salesforce "Disbursement ID" (Column G)
                  ' D: "Payment Type"..................... = Salesforce "Payment Type" (Column H)
                  ' E: "Check Number"..................... = Salesforce "Check Number" (Column I)
                  ' F: "SF Payment ID".................... = Salesforce "Payment: Payment Number" (Column J)
                  ' G: "Primary Contact".................. = Salesforce "Primary Contact" (Column K)
                  ' H: "Family Account Name".............. = Salesforce "Account Name" (Column L)
                  ' I: "Company Name"..................... = Salesforce "Company Name" (Column M)
                  ' J: "Campaign Name".................... = Salesforce "Primary Campaign Source" (Column N)
                  ' K: "Opportunity Name"................. = Salesforce "Opportunity Name" (Column O)
                    wsStandardizedSF.Range("A2").Formula2 = "=IF(" & _
                            "ISBLANK(CHOOSECOLS('" & wsInitialData.Name & "'!A2:T" & LastRow_InitialData & ",4,6,7,8,9,10,11,12,13,14,15)),""""," & _
                            "CHOOSECOLS('" & wsInitialData.Name & "'!A2:T" & LastRow_InitialData & ",4,6,7,8,9,10,11,12,13,14,15))"
                
                ' Columns L:W
                    ' "Memo"
                        ' This field is created later in the converter and should remain blank in this worksheet.
                        ' wsStandardizedSF.Range("L2").Formula = ""
                    
                    ' To be determined later:
                    ' "Intacct - Location ID"
                        wsStandardizedSF.Range("M2").Formula2 = ""
                    
                    ' "Intacct - Account"
                        wsStandardizedSF.Range("N2").Formula = "=IF(ISNUMBER(SEARCH(""Employer"",J2)),73013," & _
                                                                "IF(ISNUMBER(SEARCH(""Employee"",J2)),73011," & _
                                                                    "IF(ISNUMBER(SEARCH(""Tax"",J2)),73001," & _
                                                                        "IF(ISNUMBER(SEARCH(""Direct Donation"",J2)),73011," & _
                                                                            "IF(ISNUMBER(SEARCH(""General Fund"",J2)),73998," & _
                                                                                "IF(ISNUMBER(SEARCH(""Gala"",J2)),41005," & _
                                                                                    "IF(LEFT(J2,5)=""BASIS"",43026," & _
                                                                                        """CHECK"")))))))"
                                                        
                    ' "Intacct - Division"
                        wsStandardizedSF.Range("O2").Formula = "=IF(ISNUMBER(SEARCH(""Employer"",J2)),2048," & _
                                                                "IF(ISNUMBER(SEARCH(""Employee"",J2)),2048," & _
                                                                    "IF(ISNUMBER(SEARCH(""Tax"",J2)),2001," & _
                                                                        "IF(ISNUMBER(SEARCH(""Direct Donation"",J2)),2048," & _
                                                                            "IF(ISNUMBER(SEARCH(""General Fund"",J2)),2036," & _
                                                                                "IF(ISNUMBER(SEARCH(""Gala"",J2)),2036," & _
                                                                                    "IF(LEFT(J2,5)=""BASIS"",2036," & _
                                                                                        """CHECK"")))))))"
                                                                                    
                    ' "Intacct - Funding Source"
                        wsStandardizedSF.Range("P2").Formula = "=IF(ISNUMBER(SEARCH(""Employer"",J2)),""7301-ATF Campaign""," & _
                                                                "IF(ISNUMBER(SEARCH(""Employee"",J2)),""7301-ATF Campaign""," & _
                                                                    "IF(ISNUMBER(SEARCH(""Tax"",J2)),""7312-Tax Credit""," & _
                                                                        "IF(ISNUMBER(SEARCH(""Direct Donation"",J2)),""7301-ATF Campaign""," & _
                                                                            "IF(ISNUMBER(SEARCH(""General Fund"",J2)),""0000-Not Applicable""," & _
                                                                                "IF(ISNUMBER(SEARCH(""Gala"",J2)),""7306-Local Other (General)""," & _
                                                                                    "IF(LEFT(J2,5)=""BASIS"",""0000-Not Applicable""," & _
                                                                                        """CHECK"")))))))"
                                                                                    
                    ' "Intacct - Debt Services Series"
                        wsStandardizedSF.Range("Q2").Formula = "=IF(ISNUMBER(SEARCH(""Employer"",J2)),""000""," & _
                                                                "IF(ISNUMBER(SEARCH(""Employee"",J2)),""000""," & _
                                                                    "IF(ISNUMBER(SEARCH(""Tax"",J2)),""000""," & _
                                                                        "IF(ISNUMBER(SEARCH(""Direct Donation"",J2)),""000""," & _
                                                                            "IF(ISNUMBER(SEARCH(""General Fund"",J2)),""000""," & _
                                                                                "IF(ISNUMBER(SEARCH(""Gala"",J2)),""000""," & _
                                                                                    "IF(LEFT(J2,5)=""BASIS"",""000""," & _
                                                                                        """CHECK"")))))))"
                    
                    ' "Amount"
                        wsStandardizedSF.Range("R2").Formula2 = "=XLOOKUP(F2,'Initial Data - SF'!J:J,'Initial Data - SF'!R:R)"
                    
                  ' Columns S:W are recalculated later (after the User-Required Adjustments worksheet is populated).
                    ' "Location Correction"
                        wsStandardizedSF.Range("S2").Formula2 = "=M2"
                                
                    ' "Account Correction"
                        wsStandardizedSF.Range("T2").Formula2 = "=N2"
                                
                    ' "Division Correction"
                        wsStandardizedSF.Range("U2").Formula2 = "=O2"
                                
                    ' "Funding Source Correction"
                        wsStandardizedSF.Range("V2").Formula2 = "=P2"
                                
                    ' "Debt Services Correction"
                        wsStandardizedSF.Range("W2").Formula2 = "=Q2"

                End If
    
    ' ============================================================
    ' FIND THE LAST ROW FROM THE STANDARDIZED SALESFORCE WORKSHEET
    ' ============================================================
        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            ' Use column F because it should consistently contain the Salesforce Payment ID.
                LastRow_StandardizedSF = wsStandardizedSF.Cells(wsStandardizedSF.Rows.Count, 6).End(xlUp).Row
        
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            ' If the Initial Report came from Salesforce, fill down the formulas not already populated by the spilled formula in columns A:K.
                If InitialPath = "Salesforce" Then
                    If LastRow_StandardizedSF > 2 Then
                        wsStandardizedSF.Range("M2:W" & LastRow_StandardizedSF).FillDown
                    End If
                End If
    
    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        wsStandardizedSF.Range("A1:W1").AutoFilter
        wsStandardizedSF.Columns("A:W").AutoFit
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''----------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' STANDARDIZE DONATION SITE REPORTS DATA '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''----------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "Standardizing Donation Site Data"
    
    ' ============================================================
    '    FIND THE LAST ROW FROM THE CONSOLIDATED DATA WORKSHEET
    ' ============================================================
        ' Determine the last used row in wsConsolidatedData so the formulas can reference the full Donation Site data range.
            LastRow_ConsolidatedData = wsConsolidatedData.Cells(wsConsolidatedData.Rows.Count, 1).End(xlUp).Row

    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
        ' This worksheet standardizes the consolidated Donation Site Report data into one consistent structure.
        ' This allows later sections of the converter to work from one Donation Site layout instead of relying on the original report columns.
        
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Add the standardized Donation Site column headers.
                wsStandardizedDonationSiteData.Range("A1:P1").Value = Array("Transaction Date", "Disbursement Date", "Donation Site", "Transaction ID", "Disbursement ID", _
                        "Donation Method", "Check Number", "Donor Name (Last Name, First Name)", "Company", "Donation Type", "Donation Gross Amount", _
                        "Donation Total Fees", "Donation Net Amount", "Site - School Name", "Site - School Abbreviation", "Corrected - School Abbreviation")

        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' "Transaction Date"
                wsStandardizedDonationSiteData.Range("A2").Formula = "=TEXT('" & wsConsolidatedData.Name & "'!A2,""MM/DD/YYYY"")"
            
            ' "Disbursement Date"
                wsStandardizedDonationSiteData.Range("B2").Formula = "=TEXT('" & wsConsolidatedData.Name & "'!M2,""MM/DD/YYYY"")"
            
            ' "Donation Site"
                wsStandardizedDonationSiteData.Range("C2").Value = DonationSite
            
            ' "Transaction ID"
                wsStandardizedDonationSiteData.Range("D2").Formula = "=TEXT('" & wsConsolidatedData.Name & "'!C2,""#"")"
            
            ' "Disbursement ID"
                wsStandardizedDonationSiteData.Range("E2").Formula = "=TEXT('" & wsConsolidatedData.Name & "'!J2,""#"")"
            
            ' "Donation Method"
              ' This field is not currently available in the Donation Site Reports.
                wsStandardizedDonationSiteData.Range("F2").Formula = "="""""
            
            ' "Check Number"
              ' This field is not currently available in the Donation Site Reports.
                wsStandardizedDonationSiteData.Range("G2").Formula = "="""""
            
            ' "Donor Name (Last Name, First Name)"
              ' Build the donor name in a consistent format for later matching, review, and analysis.
              ' If the primary donor name is blank, use the matching donor fields when available.
                wsStandardizedDonationSiteData.Range("H2").Formula2 = "" & _
                        "=PROPER(TRIM(IF(AND('" & wsConsolidatedData.Name & "'!Q2="""",'" & wsConsolidatedData.Name & "'!R2="""")," & _
                                        "IF(AND('" & wsConsolidatedData.Name & "'!AC2="""",'" & wsConsolidatedData.Name & "'!AB2=""""),""""," & _
                                            "'" & wsConsolidatedData.Name & "'!AC2&"", ""&'" & wsConsolidatedData.Name & "'!AB2)," & _
                                      "'" & wsConsolidatedData.Name & "'!R2&"", ""&'" & wsConsolidatedData.Name & "'!Q2)))"
                                      
            ' "Company"
                wsStandardizedDonationSiteData.Range("I2").Formula = "='" & wsConsolidatedData.Name & "'!B2"
            
            ' "Donation Type"
                wsStandardizedDonationSiteData.Range("J2").Formula = "=IF('" & wsConsolidatedData.Name & "'!O2=""Individual"",""Employee Giving"",""Employer Matching"")"
            
            ' "Donation Gross Amount"
                wsStandardizedDonationSiteData.Range("K2").Formula = "='" & wsConsolidatedData.Name & "'!E2"
            
            ' "Donation Total Fees"
                wsStandardizedDonationSiteData.Range("L2").Formula = "=('" & wsConsolidatedData.Name & "'!F2+'" & wsConsolidatedData.Name & "'!G2)*-1"
            
            ' "Donation Net Amount"
                wsStandardizedDonationSiteData.Range("M2").Formula = "='" & wsConsolidatedData.Name & "'!H2"
            
            ' "Site - School Name"
                wsStandardizedDonationSiteData.Range("N2").Formula = "='" & wsConsolidatedData.Name & "'!AI2"
                
            ' "Site - School Abbreviation"
                ' Use the custom function to convert the Donation Site school name into the BASIS school abbreviation used later in the converter.
                wsStandardizedDonationSiteData.Range("O2").Formula2 = "=ConvertYourCauseToSchoolAbbrev(N2)"
            
            ' "Corrected - School Abbreviation"
              ' Start by carrying forward the original school abbreviation when one is found.
              ' This value can be recalculated later if a user correction is needed.
                wsStandardizedDonationSiteData.Range("P2").Formula = "=IF(O2=""No School Found"","""",O2)"

    ' ============================================================
    '   FILL FORMULAS DOWN AND FIND THE LAST ROW OF THE WORKSHEET
    ' ============================================================
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            ' Fill the formulas down using the last row from wsConsolidatedData so every consolidated Donation Site record is standardized.
                If LastRow_ConsolidatedData > 2 Then
                    wsStandardizedDonationSiteData.Range("A2:P" & LastRow_ConsolidatedData).FillDown
                End If
        
        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            ' Use column D because it should consistently contain the standardized Transaction ID.
                LastRow_StandardizedDonationSiteData = wsStandardizedDonationSiteData.Cells(wsStandardizedDonationSiteData.Rows.Count, 4).End(xlUp).Row
    
    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        wsStandardizedDonationSiteData.Range("A1:P1").AutoFilter
        wsStandardizedDonationSiteData.Columns("A:P").AutoFit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' POPULATE THE DISBURSEMENT DATA WORKSHEET ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "Populating the Disbursement Data Worksheet"
        
    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
      ' This worksheet summarizes Donation Site transactions at the disbursement level.
      ' It exists so later sections of the converter can work from one row per disbursement when building fees, bank deposits, descriptions, and import data.
        
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Add the Disbursement Data column headers.
                wsDisbursementData.Range("A1:M1").Value = Array("Donation Site", "Disbursement Date", "Disbursement ID", "School Abbreviation", "Account", "Company", "Gross Amount", _
                        "Fees", "Net Amount", "CRJ Description", "Adjusting Journal Description", "Fee Reimbursement", "File Name")

        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' "Donation Site"
                wsDisbursementData.Range("A2").Formula2 = "=XLOOKUP(C2,'" & wsStandardizedDonationSiteData.Name & "'!E:E,'" & wsStandardizedDonationSiteData.Name & "'!C:C)"
            
            ' "Disbursement Date"
                wsDisbursementData.Range("B2").Formula2 = "=XLOOKUP(C2,'" & wsStandardizedDonationSiteData.Name & "'!E:E,'" & wsStandardizedDonationSiteData.Name & "'!B:B)"
            
            ' "Disbursement ID"
              ' This column builds the unique list of Disbursement IDs that drives the rest of the worksheet.
                wsDisbursementData.Range("C2").Formula2 = "=UNIQUE('" & wsStandardizedDonationSiteData.Name & "'!E2:E" & LastRow_StandardizedDonationSiteData & ")"
            
            ' "School Abbreviation"
                wsDisbursementData.Range("D2").Formula2 = "=XLOOKUP(C2,'" & wsStandardizedDonationSiteData.Name & "'!E:E,'" & wsStandardizedDonationSiteData.Name & "'!P:P)"
            
            ' "Account"
              ' Use the custom function to convert the school abbreviation into the Intacct account used later in the converter.
                wsDisbursementData.Range("E2").Formula2 = "=ConvertSchoolAbbrevToIntacctAccount(D2)"
            
            ' "Company"
                wsDisbursementData.Range("F2").Formula2 = "=XLOOKUP(C2,'" & wsStandardizedDonationSiteData.Name & "'!E:E,'" & wsStandardizedDonationSiteData.Name & "'!I:I)"
            
            ' "Gross Amount"
              ' Sum all gross donation amounts tied to the Disbursement ID so the worksheet shows one gross amount per disbursement.
                wsDisbursementData.Range("G2").Formula2 = "=SUMIFS('" & wsStandardizedDonationSiteData.Name & "'!K:K,'" & wsStandardizedDonationSiteData.Name & "'!$E:$E,$C2)"
            
            ' "Fees"
              ' Sum all fees tied to the Disbursement ID so the worksheet shows one fee total per disbursement.
                wsDisbursementData.Range("H2").Formula2 = "=SUMIFS('" & wsStandardizedDonationSiteData.Name & "'!L:L,'" & wsStandardizedDonationSiteData.Name & "'!$E:$E,$C2)"
            
            ' "Net Amount"
              ' Sum all net donation amounts tied to the Disbursement ID so the worksheet shows one net amount per disbursement.
                wsDisbursementData.Range("I2").Formula2 = "=SUMIFS('" & wsStandardizedDonationSiteData.Name & "'!M:M,'" & wsStandardizedDonationSiteData.Name & "'!$E:$E,$C2)"
            
            ' "CRJ Description"
              ' Build the CRJ description from the Donation Site, school abbreviation, Disbursement ID, and company name when available.
                wsDisbursementData.Range("J2").Formula = "=A2&"" - ""&D2&"" (""&C2&"")""&IF(F2<>"""","" {""&F2&""}"","""")"
            
            ' "Adjusting Journal Description"
              ' Build the Adjusting Journal description from the Donation Site, school abbreviation, Disbursement ID, net amount, and company name when available.
                wsDisbursementData.Range("K2").Formula = "=A2&"" - ""&D2&"" (""&C2&"") [""&TEXT(I2,""$#,##0.00"")&""]""&IF(F2<>"""","" {""&F2&""}"","""")"
            
            ' "Fee Reimbursement"
              ' Your Cause does not allow fee reimbursements, so this remains "No" for every disbursement.
                wsDisbursementData.Range("L2").Formula = "=""No"""
            
            ' "File Name"
              ' Pull the source file name tied to the Disbursement ID so the original report file can be traced later if needed.
                wsDisbursementData.Range("M2").Formula = "=XLOOKUP(C2,'" & wsConsolidatedData.Name & "'!J:J,'" & wsConsolidatedData.Name & "'!AP:AP," & _
                        "XLOOKUP(NUMBERVALUE(C2),'" & wsConsolidatedData.Name & "'!J:J,'" & wsConsolidatedData.Name & "'!AP:AP))"

    ' ============================================================
    '    FIND THE LAST ROW FROM THE DISBURSEMENT DATA WORKSHEET
    ' ============================================================
        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            ' Use column C because it holds the unique Disbursement IDs that drive the worksheet.
                LastRow_DisbursementData = wsDisbursementData.Cells(wsDisbursementData.Rows.Count, 3).End(xlUp).Row
        
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            ' Fill down all formulas that are not populated by the spilled UNIQUE formula in column C.
                If LastRow_DisbursementData > 2 Then
                    wsDisbursementData.Range("A2:B" & LastRow_DisbursementData).FillDown
                    wsDisbursementData.Range("D2:M" & LastRow_DisbursementData).FillDown
                End If
    
    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        wsDisbursementData.Range("A1:M1").AutoFilter
        wsDisbursementData.Columns("A:M").AutoFit
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''----------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' POPULATE THE RELEVANT TRANSACTIONS WORKSHEET ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''----------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "Populating the Relevant Transactions Worksheet"
    
    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
      ' This worksheet connects the standardized Salesforce data to the standardized Donation Site data at the transaction level.
      ' It exists to keep only the transactions that are relevant for import-file creation and to stage the fields needed for both the Adjusting and CRJ paths.
        
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Main fields used to connect Salesforce data to Donation Site data and support later import-file creation.
                wsRelevantTransactions.Range("A1:Z1").Value = Array("Transaction ID", "PMT-ID", ".......", "Transaction Date", "Disbursement Date", "Donation Site", _
                        "Transaction ID", "Disbursement ID", "Payment Type", "Check Number", "SF Payment ID", "Primary Contact", "Account Name", "Company Name", _
                        "Campaign Name", "Opportunity Name", "School Abbreviation", "Donation Type", "Memo", "Intacct - Location ID", "Intacct - Account", _
                        "Intacct - Division", "Intacct - Funding Source", "Intacct - Debt Services Series", "Gross Amount", ".......")
            
            ' Fields used later in the Adjusting Journal path.
                wsRelevantTransactions.Range("AA1:BG1").Value = Array("JOURNAL", "DATE", "REVERSEDATE", "DESCRIPTION", "REFERENCE_NO", "LINE_NO", "ACCT_NO", "LOCATION_ID", "DEPT_ID", _
                        "DOCUMENT", "MEMO", "DEBIT", "CREDIT", "SOURCEENTITY", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCHANGE_RATE", "STATE", _
                        "ALLOCATION_ID", "RASSET", "RDEPRECIATION_SCHEDULE", "RASSET_ADJUSTMENT", "RASSET_CLASS", "RASSETOUTOFSERVICE", "GLDIMFUNDING_SOURCE", _
                        "GLENTRY_PROJECTID", "GLENTRY_CUSTOMERID", "GLENTRY_VENDORID", "GLENTRY_EMPLOYEEID", "GLENTRY_ITEMID", "GLENTRY_CLASSID", ".......")
            
            ' Fields used later in the CRJ path.
                wsRelevantTransactions.Range("BH1:CM1").Value = Array("RECEIPT_DATE", "PAYMETHOD", "DOCDATE", "DOCNUMBER", "DESCRIPTION", "DEPOSITTO", "BANKACCOUNTID", _
                        "DEPOSITDATE", "UNDEPACCTNO", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCH_RATE_DATE", "LINE_NO", "ACCT_NO", "ACCOUNTLABEL", "TRX_AMOUNT", _
                        "AMOUNT", "DEPT_ID", "LOCATION_ID", "ITEM_MEMO", "OTHERRECEIPTSENTRY_PROJECTID", "OTHERRECEIPTSENTRY_CUSTOMERID", "OTHERRECEIPTSENTRY_ITEMID", _
                        "OTHERRECEIPTSENTRY_VENDORID", "OTHERRECEIPTSENTRY_EMPLOYEEID", "OTHERRECEIPTSENTRY_CLASSID", "PAYER_NAME", "SUPDOCID", "EXCHANGE_RATE", _
                        "OR_TRANSACTION_DATE", "GLDIMFUNDING_SOURCE")

        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' "Transaction ID" (Salesforce Data)
                wsRelevantTransactions.Range("A2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!B:B)"
            
            ' "PMT-ID"
                ' Build the filtered list of PMT-IDs that exist in Salesforce and also have a matching Transaction ID in the standardized Donation Site data.
                wsRelevantTransactions.Range("B2").Formula2 = "=FILTER('" & wsStandardizedSF.Name & "'!$F2:$F" & LastRow_StandardizedSF & _
                        ",ISNUMBER(MATCH('" & wsStandardizedSF.Name & "'!$B2:$B" & LastRow_StandardizedSF & _
                        ",'" & wsStandardizedDonationSiteData.Name & "'!$D2:$D" & LastRow_StandardizedDonationSiteData & ",0)))"
            
            ' "......."
                wsRelevantTransactions.Range("C2").Value = "......."
            
            ' "Transaction Date" (Donation Site Data)
                wsRelevantTransactions.Range("D2").Formula2 = "=XLOOKUP($A2,'" & wsStandardizedDonationSiteData.Name & "'!$D:$D,'" & wsStandardizedDonationSiteData.Name & "'!A:A)"
            
            ' "Disbursement Date" (Donation Site Data)
                wsRelevantTransactions.Range("E2").Formula2 = "=XLOOKUP($A2,'" & wsStandardizedDonationSiteData.Name & "'!$D:$D,'" & wsStandardizedDonationSiteData.Name & "'!B:B)"
            
            ' "Donation Site" (Donation Site Data)
                wsRelevantTransactions.Range("F2").Formula2 = "=XLOOKUP($A2,'" & wsStandardizedDonationSiteData.Name & "'!$D:$D,'" & wsStandardizedDonationSiteData.Name & "'!C:C)"
            
            ' "Transaction ID" (Donation Site Data)
                wsRelevantTransactions.Range("G2").Formula2 = "=XLOOKUP($A2,'" & wsStandardizedDonationSiteData.Name & "'!$D:$D,'" & wsStandardizedDonationSiteData.Name & "'!D:D)"
            
            ' "Disbursement ID" (Donation Site Data)
                wsRelevantTransactions.Range("H2").Formula2 = "=XLOOKUP($A2,'" & wsStandardizedDonationSiteData.Name & "'!$D:$D,'" & wsStandardizedDonationSiteData.Name & "'!E:E)"
            
            ' "Payment Type" (Salesforce Data)
                wsRelevantTransactions.Range("I2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!D:D)"
            
            ' "Check Number" (Salesforce Data)
                wsRelevantTransactions.Range("J2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!E:E)"
            
            ' "SF Payment ID" (Salesforce Data)
                wsRelevantTransactions.Range("K2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!F:F)"
            
            ' "Primary Contact" (Salesforce Data)
                wsRelevantTransactions.Range("L2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!G:G)"
            
            ' "Account Name" (Salesforce Data)
                wsRelevantTransactions.Range("M2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!H:H)"
            
            ' "Company Name" (Donation Site Data)
                wsRelevantTransactions.Range("N2").Formula2 = "=XLOOKUP($A2,'" & wsStandardizedDonationSiteData.Name & "'!$D:$D,'" & wsStandardizedDonationSiteData.Name & "'!I:I)"
            
            ' "Campaign Name" (Salesforce Data)
                wsRelevantTransactions.Range("O2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!J:J)"
            
            ' "Opportunity Name" (Salesforce Data)
                wsRelevantTransactions.Range("P2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!K:K)"
            
            ' "School Abbreviation" (Donation Site Data)
                wsRelevantTransactions.Range("Q2").Formula2 = "=XLOOKUP($A2,'" & wsStandardizedDonationSiteData.Name & "'!$D:$D,'" & wsStandardizedDonationSiteData.Name & "'!P:P)"
            
            ' "Donation Type" (Donation Site Data)
                wsRelevantTransactions.Range("R2").Formula2 = "=XLOOKUP($A2,'" & wsStandardizedDonationSiteData.Name & "'!$D:$D,'" & wsStandardizedDonationSiteData.Name & "'!J:J)"
            
            ' "Memo"
                ' If the converter started from Intacct, carry forward the existing memo.
                ' Otherwise, build a detailed documentation memo from the combined Salesforce and Donation Site fields.
                If StartingPoint = "Intacct" Then
                    ' (Salesforce Data)
                        wsRelevantTransactions.Range("S2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!L:L)"
                Else
                    ' 1 - 2 | 3 | Transaction Date: 4 | Disbursement Date: 5 | Site: 6 | Transaction ID: 7 | Disbursement ID: 8 | Payment Method: 9 + 10 | 11 | Company: 12 | ^13 | *14
                      '  1 = Donation Site School Abbreviation ............. (Column Q)
                      '  2 = SF Campaign Name .............................. (Column O)
                      '  3 = SF Opportunity Name ........................... (Column P)
                      '  4 = Donation Site Transaction Date (MM.DD.YYYY) ... (Column D)
                      '  5 = Donation Site Disbursement Date (MM.DD.YYYY) .. (Column E)
                      '  6 = Donation Site Name ............................ (Column F)
                      '  7 = Donation Site Transaction ID .................. (Column G)
                      '  8 = Donation Site Disbursement ID ................. (Column H)
                      '  9 = SF Payment Method ............................. (Column I)
                      ' 10 = SF Check Number ............................... (Column J)
                      ' 11 = SF Payment ID ................................. (Column K)
                      ' 12 = Company Name .................................. (Column N)
                      ' 13 = SF Donor Name ................................. (Column L)
                      ' 14 = SF Family Account Name ........................ (Column M)
                        wsRelevantTransactions.Range("S2").Formula2 = "" & _
                                                                    "=$Q2&" & _
                                                                    """ - ""&$O2&" & _
                                                                    """ | ""&$P2&" & _
                                                                    """ | Transaction Date: ""&TEXT($D2,""MM.DD.YYYY"")&" & _
                                                                    """ | Disbursement Date: ""&TEXT($E2,""MM.DD.YYYY"")&" & _
                                                                    """ | Site: ""&$F2&" & _
                                                                    """ | Transaction ID: ""&$G2&" & _
                                                                    """ | Disbursement ID: ""&$H2&" & _
                                                                    """ | Payment Method: ""&IF($I2=""Check"",""Check# ""&$J2,$I2)&" & _
                                                                    """ | ""&$K2&" & _
                                                                    """ | Company: ""&$N2&" & _
                                                                    """ | ^""&$L2&" & _
                                                                    """ | *""&$M2"
                End If
                 
            ' "Intacct - Location ID" (Salesforce Data)
                wsRelevantTransactions.Range("T2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!S:S)"
            
            ' "Intacct - Account" (Salesforce Data)
                wsRelevantTransactions.Range("U2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!T:T)"
        
            ' "Intacct - Division" (Salesforce Data)
                wsRelevantTransactions.Range("V2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!U:U)"
            
            ' "Intacct - Funding Source" (Salesforce Data)
                wsRelevantTransactions.Range("W2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!V:V)"
            
            ' "Intacct - Debt Services Series" (Salesforce Data)
                wsRelevantTransactions.Range("X2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!W:W)"
            
            ' "Gross Amount" (Salesforce Data)
                wsRelevantTransactions.Range("Y2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!R:R)"
            
            ' "......."
                wsRelevantTransactions.Range("Z2").Formula2 = "......."
            
            ' ..............................
            '        ADJUSTING JOURNAL
            '         COLUMN HEADERS
            ' ..............................
                ' "JOURNAL"
                    wsRelevantTransactions.Range("AA2").Formula2 = JournalName
                    
                ' "DATE" = Disbursement Date (Donation Site Data)
                    wsRelevantTransactions.Range("AB2").Formula2 = "=E2"
                    
                ' "REVERSEDATE"
                    wsRelevantTransactions.Range("AC2").Formula = "="""""
                    
                ' "DESCRIPTION" (Disbursement Data)
                    wsRelevantTransactions.Range("AD2").Formula2 = "=XLOOKUP($H2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!K:K)"
                    
                ' "REFERENCE_NO"
                    wsRelevantTransactions.Range("AE2").Formula = "="""""
                    
                ' "LINE_NO"
                    wsRelevantTransactions.Range("AF2").Formula = "="""""
                    
                ' "ACCT_NO"
                    wsRelevantTransactions.Range("AG2").Formula2 = "=U2"
                    
                ' "LOCATION_ID"
                    wsRelevantTransactions.Range("AH2").Formula2 = "=T2"
                    
                ' "DEPT_ID"
                    wsRelevantTransactions.Range("AI2").Formula2 = "=V2"
                    
                ' "DOCUMENT"
                    wsRelevantTransactions.Range("AJ2").Formula = "="""""
                    
                ' "MEMO"
                    wsRelevantTransactions.Range("AK2").Formula2 = "=S2"
                    
                ' "DEBIT"
                    wsRelevantTransactions.Range("AL2").Formula2 = "="""""
                    
                ' "CREDIT"
                    wsRelevantTransactions.Range("AM2").Formula2 = "=Y2"
                    
                ' "SOURCEENTITY"
                    wsRelevantTransactions.Range("AN2").Formula = "="""""
                    
                ' "CURRENCY"
                    wsRelevantTransactions.Range("AO2").Formula = "="""""
                    
                ' "EXCH_RATE_DATE"
                    wsRelevantTransactions.Range("AP2").Formula = "="""""
                    
                ' "EXCH_RATE_TYPE_ID"
                    wsRelevantTransactions.Range("AQ2").Formula = "="""""
                    
                ' "EXCHANGE_RATE"
                    wsRelevantTransactions.Range("AR2").Formula = "="""""
                    
                ' "STATE"
                    wsRelevantTransactions.Range("AS2").Formula = "=""Draft"""
                    
                ' "ALLOCATION_ID"
                    wsRelevantTransactions.Range("AT2").Formula = "="""""
                    
                ' "RASSET"
                    wsRelevantTransactions.Range("AU2").Formula = "="""""
                    
                ' "RDEPRECIATION_SCHEDULE"
                    wsRelevantTransactions.Range("AV2").Formula = "="""""
                    
                ' "RASSET_ADJUSTMENT"
                    wsRelevantTransactions.Range("AW2").Formula = "="""""
                    
                ' "RASSET_CLASS"
                    wsRelevantTransactions.Range("AX2").Formula = "="""""
                    
                ' "RASSETOUTOFSERVICE"
                    wsRelevantTransactions.Range("AY2").Formula = "="""""
                    
                ' "GLDIMFUNDING_SOURCE"
                    wsRelevantTransactions.Range("AZ2").Formula2 = "=W2"
                    
                ' "GLENTRY_PROJECTID"
                    wsRelevantTransactions.Range("BA2").Formula = "="""""
                    
                ' "GLENTRY_CUSTOMERID"
                    wsRelevantTransactions.Range("BB2").Formula = "="""""
                    
                ' "GLENTRY_VENDORID"
                    wsRelevantTransactions.Range("BC2").Formula = "="""""
                    
                ' "GLENTRY_EMPLOYEEID"
                    wsRelevantTransactions.Range("BD2").Formula = "="""""
                    
                ' "GLENTRY_ITEMID"
                    wsRelevantTransactions.Range("BE2").Formula = "="""""
                    
                ' "GLENTRY_CLASSID"
                    wsRelevantTransactions.Range("BF2").Formula = "=X2"
                    
                ' "......."
                    wsRelevantTransactions.Range("BG2").Formula2 = "......."
                    
            ' ..............................
            '       CRJ COLUMN HEADERS
            ' ..............................
                ' "RECEIPT_DATE" = Disbursement Date (Donation Site Data)
                    wsRelevantTransactions.Range("BH2").Formula = "=E2"
                    
                ' "PAYMETHOD" = "Credit Card"
                    wsRelevantTransactions.Range("BI2").Formula = "=""Credit Card"""
                    
                ' "DOCDATE" = Disbursement Date (Donation Site Data)
                    wsRelevantTransactions.Range("BJ2").Formula = "=E2"
                    
                ' "DOCNUMBER" = Donation Site Name
                    wsRelevantTransactions.Range("BK2").Formula2 = DonationSite
                    
                ' "DESCRIPTION" (Disbursement Data)
                    wsRelevantTransactions.Range("BL2").Formula2 = "=XLOOKUP($H2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!J:J)"
                    
                ' "DEPOSITTO" = "Bank account"
                    wsRelevantTransactions.Range("BM2").Formula = "=""Bank account"""
                    
                ' "BANKACCOUNTID" (Disbursement Data)
                    wsRelevantTransactions.Range("BN2").Formula2 = "=ConvertSchoolAbbrevToBankAccountName(Q2)"
                    
                ' "DEPOSITDATE" = Disbursement Date (Donation Site Data)
                    wsRelevantTransactions.Range("BO2").Formula2 = "=E2"
                    
                ' "UNDEPACCTNO"
                    wsRelevantTransactions.Range("BP2").Formula = "="""""
                    
                ' "CURRENCY" = "USD"
                    wsRelevantTransactions.Range("BQ2").Formula2 = "=""USD"""
                    
                ' "EXCH_RATE_DATE"
                    wsRelevantTransactions.Range("BR2").Formula = "="""""
                    
                ' "EXCH_RATE_TYPE_ID"
                    wsRelevantTransactions.Range("BS2").Formula = "="""""
                    
                ' "EXCH_RATE_DATE"
                    wsRelevantTransactions.Range("BT2").Formula = "="""""
                    
                ' "LINE_NO"
                    wsRelevantTransactions.Range("BU2").Formula = "="""""
                    
                ' "ACCT_NO"
                    wsRelevantTransactions.Range("BV2").Formula2 = "=U2"
                    
                ' "ACCOUNTLABEL"
                    wsRelevantTransactions.Range("BW2").Formula = "="""""
                    
                ' "TRX_AMOUNT"
                    wsRelevantTransactions.Range("BX2").Formula2 = "=Y2"
                    
                ' "AMOUNT"
                    wsRelevantTransactions.Range("BY2").Formula2 = "=Y2"
                    
                ' "DEPT_ID"
                    wsRelevantTransactions.Range("BZ2").Formula2 = "=V2"
                    
                ' "LOCATION_ID"
                    wsRelevantTransactions.Range("CA2").Formula2 = "=T2"
                    
                ' "ITEM_MEMO"
                    wsRelevantTransactions.Range("CB2").Formula2 = "=S2"
                    
                ' "OTHERRECEIPTSENTRY_PROJECTID"
                    wsRelevantTransactions.Range("CC2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_CUSTOMERID" = Intacct Donation Site ID
                    wsRelevantTransactions.Range("CD2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_ITEMID"
                    wsRelevantTransactions.Range("CE2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_VENDORID"
                    wsRelevantTransactions.Range("CF2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_EMPLOYEEID"
                    wsRelevantTransactions.Range("CG2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_CLASSID"
                    wsRelevantTransactions.Range("CH2").Formula2 = "=X2"
                    
                ' "PAYER_NAME" = Donation Site Name
                    wsRelevantTransactions.Range("CI2").Formula2 = DonationSite
                    
                ' "SUPDOCID"
                    wsRelevantTransactions.Range("CJ2").Formula = "="""""
                    
                ' "EXCHANGE_RATE"
                    wsRelevantTransactions.Range("CK2").Formula = "="""""
                    
                ' "OR_TRANSACTION_DATE" = Disbursement Date (Donation Site Data)
                    wsRelevantTransactions.Range("CL2").Formula2 = "=E2"
                    
                ' "GLDIMFUNDING_SOURCE"
                    wsRelevantTransactions.Range("CM2").Formula2 = "=W2"
                    
    
    ' ============================================================
    '  FIND THE LAST ROW FROM THE RELEVANT TRANSACTIONS WORKSHEET
    ' ============================================================
        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            ' Use column B because it holds the filtered PMT-IDs that drive the worksheet.
                LastRow_RelevantTransactions = wsRelevantTransactions.Cells(wsRelevantTransactions.Rows.Count, 2).End(xlUp).Row
        
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            ' Fill down all formulas that are not already populated by the spilled FILTER formula in column B.
                If LastRow_RelevantTransactions > 2 Then
                    wsRelevantTransactions.Range("A2:A" & LastRow_RelevantTransactions).FillDown
                    wsRelevantTransactions.Range("C2:CM" & LastRow_RelevantTransactions).FillDown
                End If

    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        wsRelevantTransactions.Range("A1:CM1").AutoFilter
        wsRelevantTransactions.Columns("A:CM").AutoFit
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' POPULATE THE FEES WORKSHEET ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
            Application.StatusBar = "Populating Fees Worksheet"

    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
      ' This worksheet isolates the fee portion of each disbursement so the fees can be imported as separate line items.
      ' It exists to create fee-specific rows for both the Adjusting and CRJ import paths.
        
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Main fields used to identify each fee row and connect it back to the related disbursement.
                wsFees.Range("A1:D1").Value = Array("Disbursement ID", "Fees", "School Abbreviation", ".......")
            
            ' Fields used later in the Adjusting Journal path.
                wsFees.Range("E1:AK1").Value = Array("JOURNAL", "DATE", "REVERSEDATE", "DESCRIPTION", "REFERENCE_NO", "LINE_NO", "ACCT_NO", "LOCATION_ID", "DEPT_ID", _
                        "DOCUMENT", "MEMO", "DEBIT", "CREDIT", "SOURCEENTITY", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCHANGE_RATE", "STATE", _
                        "ALLOCATION_ID", "RASSET", "RDEPRECIATION_SCHEDULE", "RASSET_ADJUSTMENT", "RASSET_CLASS", "RASSETOUTOFSERVICE", "GLDIMFUNDING_SOURCE", _
                        "GLENTRY_PROJECTID", "GLENTRY_CUSTOMERID", "GLENTRY_VENDORID", "GLENTRY_EMPLOYEEID", "GLENTRY_ITEMID", "GLENTRY_CLASSID", ".......")
                        
            ' Fields used later in the CRJ path.
                wsFees.Range("AL1:BQ1").Value = Array("RECEIPT_DATE", "PAYMETHOD", "DOCDATE", "DOCNUMBER", "DESCRIPTION", "DEPOSITTO", "BANKACCOUNTID", _
                        "DEPOSITDATE", "UNDEPACCTNO", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCH_RATE_DATE", "LINE_NO", "ACCT_NO", "ACCOUNTLABEL", "TRX_AMOUNT", _
                        "AMOUNT", "DEPT_ID", "LOCATION_ID", "ITEM_MEMO", "OTHERRECEIPTSENTRY_PROJECTID", "OTHERRECEIPTSENTRY_CUSTOMERID", "OTHERRECEIPTSENTRY_ITEMID", _
                        "OTHERRECEIPTSENTRY_VENDORID", "OTHERRECEIPTSENTRY_EMPLOYEEID", "OTHERRECEIPTSENTRY_CLASSID", "PAYER_NAME", "SUPDOCID", "EXCHANGE_RATE", _
                        "OR_TRANSACTION_DATE", "GLDIMFUNDING_SOURCE")

        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' "Disbursement ID"
              ' Build a filtered list of only the Disbursement IDs that contain non-zero fee amounts.
                wsFees.Range("A2").Formula2 = "=FILTER('" & wsDisbursementData.Name & "'!C:C,('" & _
                        wsDisbursementData.Name & "'!H:H<>0)*('" & wsDisbursementData.Name & "'!H:H<>"""")*('" & wsDisbursementData.Name & "'!H:H<>""Fees""))"

            ' "Fees"
              ' Pull the fee total tied to the Disbursement ID.
                wsFees.Range("B2").Formula2 = "=XLOOKUP($A2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!H:H)"
            
            ' "School Abbreviation"
              ' Pull the school abbreviation tied to the Disbursement ID so the correct school dimensions can be applied later.
                wsFees.Range("C2").Formula2 = "=XLOOKUP($A2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!D:D)"
            
            ' "......."
                wsFees.Range("D2").Value = "......."

            ' ..............................
            '        ADJUSTING JOURNAL
            '         COLUMN HEADERS
            ' ..............................
                ' "JOURNAL"
                    wsFees.Range("E2").Formula2 = JournalName
                    
                ' "DATE" = Disbursement Date (Disbursement Data)
                    wsFees.Range("F2").Formula2 = "=XLOOKUP($A2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!B:B)"
                    
                ' "REVERSEDATE"
                    wsFees.Range("G2").Formula = "="""""
                    
                ' "DESCRIPTION" = Adjusting Journal Description (Disbursement Data)
                    wsFees.Range("H2").Formula2 = "=XLOOKUP($A2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!K:K)"
                    
                ' "REFERENCE_NO"
                    wsFees.Range("I2").Formula = "="""""
                    
                ' "LINE_NO"
                    wsFees.Range("J2").Formula = "="""""
                    
                ' "ACCT_NO"
                    wsFees.Range("K2").Formula = "=""82401"""
                    
                ' "LOCATION_ID"
                    wsFees.Range("L2").Formula2 = "=ConvertSchoolAbbrevToIntacctAccount(C2)"
                    
                ' "DEPT_ID"
                    wsFees.Range("M2").Formula = "=""2036"""
                    
                ' "DOCUMENT"
                    wsFees.Range("N2").Formula = "="""""
                    
                ' "MEMO"
                    wsFees.Range("O2").Formula = "=""Transaction Fees ("" & A2 & "")"""
                    
                ' "DEBIT"
                  ' Fee rows are posted as positive debits.
                    wsFees.Range("P2").Formula = "=ABS(B2)"
                    
                ' "CREDIT"
                    wsFees.Range("Q2").Formula = "="""""
                    
                ' "SOURCEENTITY"
                    wsFees.Range("R2").Formula = "="""""
                    
                ' "CURRENCY"
                    wsFees.Range("S2").Formula = "="""""
                    
                ' "EXCH_RATE_DATE"
                    wsFees.Range("T2").Formula = "="""""
                    
                ' "EXCH_RATE_TYPE_ID"
                    wsFees.Range("U2").Formula = "="""""
                    
                ' "EXCHANGE_RATE"
                    wsFees.Range("V2").Formula = "="""""
                    
                ' "STATE"
                    wsFees.Range("W2").Formula = "=""Draft"""
                    
                ' "ALLOCATION_ID"
                    wsFees.Range("X2").Formula = "="""""
                    
                ' "RASSET"
                    wsFees.Range("Y2").Formula = "="""""
                    
                ' "RDEPRECIATION_SCHEDULE"
                    wsFees.Range("Z2").Formula = "="""""
                    
                ' "RASSET_ADJUSTMENT"
                    wsFees.Range("AA2").Formula = "="""""
                    
                ' "RASSET_CLASS"
                    wsFees.Range("AB2").Formula = "="""""
                    
                ' "RASSETOUTOFSERVICE"
                    wsFees.Range("AC2").Formula = "="""""
                    
                ' "GLDIMFUNDING_SOURCE"
                    wsFees.Range("AD2").Formula = "=""7301-ATF Campaign"""
                    
                ' "GLENTRY_PROJECTID"
                    wsFees.Range("AE2").Formula = "="""""
                    
                ' "GLENTRY_CUSTOMERID"
                    wsFees.Range("AF2").Formula = "="""""
                    
                ' "GLENTRY_VENDORID"
                    wsFees.Range("AG2").Formula = "="""""
                    
                ' "GLENTRY_EMPLOYEEID"
                    wsFees.Range("AH2").Formula = "="""""
                    
                ' "GLENTRY_ITEMID"
                    wsFees.Range("AI2").Formula = "="""""
                    
                ' "GLENTRY_CLASSID"
                    wsFees.Range("AJ2").Formula = "=""000"""
                    
                ' "......."
                    wsFees.Range("AK2").Formula2 = "......."
                    
            ' ..............................
            '       CRJ COLUMN HEADERS
            ' ..............................
                ' "RECEIPT_DATE" = Disbursement Date (Disbursement Data)
                    wsFees.Range("AL2").Formula = "=XLOOKUP($A2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!B:B)"
                    
                ' "PAYMETHOD" = "Credit Card"
                    wsFees.Range("AM2").Formula = "=""Credit Card"""
                    
                ' "DOCDATE" = Disbursement Date (Disbursement Data)
                    wsFees.Range("AN2").Formula = "=AL2"
                    
                ' "DOCNUMBER" = Donation Site Name
                    wsFees.Range("AO2").Formula2 = DonationSite
                    
                ' "DESCRIPTION" = CRJ Description (Disbursement Data)
                    wsFees.Range("AP2").Formula2 = "=XLOOKUP($A2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!J:J)"
                    
                ' "DEPOSITTO" = "Bank account"
                    wsFees.Range("AQ2").Formula = "=""Bank account"""
                    
                ' "BANKACCOUNTID" (Disbursement Data)
                    wsFees.Range("AR2").Formula2 = "=ConvertSchoolAbbrevToBankAccountName(C2)"
                    
                ' "DEPOSITDATE" = Disbursement Date (Disbursement Data)
                    wsFees.Range("AS2").Formula2 = "=AL2"
                    
                ' "UNDEPACCTNO"
                    wsFees.Range("AT2").Formula = "="""""
                    
                ' "CURRENCY" = "USD"
                    wsFees.Range("AU2").Formula2 = "=""USD"""
                    
                ' "EXCH_RATE_DATE"
                    wsFees.Range("AV2").Formula = "="""""
                    
                ' "EXCH_RATE_TYPE_ID"
                    wsFees.Range("AW2").Formula = "="""""
                    
                ' "EXCH_RATE_DATE"
                    wsFees.Range("AX2").Formula = "="""""
                    
                ' "LINE_NO"
                    wsFees.Range("AY2").Formula = "="""""
                    
                ' "ACCT_NO"
                    wsFees.Range("AZ2").Formula2 = "=""82401"""
                    
                ' "ACCOUNTLABEL"
                    wsFees.Range("BA2").Formula = "="""""
                    
                ' "TRX_AMOUNT"
                  ' The CRJ fee row uses the fee amount directly.
                    wsFees.Range("BB2").Formula = "=B2"
                    
                ' "AMOUNT"
                    wsFees.Range("BC2").Formula = "=B2"
                    
                ' "DEPT_ID"
                    wsFees.Range("BD2").Formula2 = "=""2036"""
                    
                ' "LOCATION_ID"
                    wsFees.Range("BE2").Formula2 = "=ConvertSchoolAbbrevToIntacctAccount(C2)"
                    
                ' "ITEM_MEMO"
                    wsFees.Range("BF2").Formula2 = "=""Transaction Fees ("" & A2 & "")"""
                    
                ' "OTHERRECEIPTSENTRY_PROJECTID"
                    wsFees.Range("BG2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_CUSTOMERID"  = Intacct Donation Site ID
                    wsFees.Range("BH2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_ITEMID"
                    wsFees.Range("BI2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_VENDORID"
                    wsFees.Range("BJ2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_EMPLOYEEID"
                    wsFees.Range("BK2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_CLASSID"
                    wsFees.Range("BL2").Formula2 = "=""000"""
                    
                ' "PAYER_NAME" = Donation Site Name
                    wsFees.Range("BM2").Formula2 = DonationSite
                    
                ' "SUPDOCID"
                    wsFees.Range("BN2").Formula = "="""""
                    
                ' "EXCHANGE_RATE"
                    wsFees.Range("BO2").Formula = "="""""
                    
                ' "OR_TRANSACTION_DATE" = Disbursement Date (Disbursement Data)
                    wsFees.Range("BP2").Formula2 = "=AL2"
                    
                ' "GLDIMFUNDING_SOURCE"
                    wsFees.Range("BQ2").Formula2 = "=""7301-ATF Campaign"""
                    
    
    ' ============================================================
    '           FIND THE LAST ROW FROM THE FEES WORKSHEET
    ' ============================================================
        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            ' Use column A because it holds the filtered Disbursement IDs that drive the worksheet.
                LastRow_Fees = wsFees.Cells(wsFees.Rows.Count, 1).End(xlUp).Row
        
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            ' Fill down all formulas that are not already populated by the spilled FILTER formula in column A.
                If LastRow_Fees > 2 Then
                    wsFees.Range("B2:BQ" & LastRow_Fees).FillDown
                End If

    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        wsFees.Range("A1:BQ1").AutoFilter
        wsFees.Columns("A:BQ").AutoFit
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' POPULATE THE BANK DEPOSITS WORKSHEET ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
            Application.StatusBar = "Populating Bank Deposits Worksheet"

    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
      ' This worksheet creates one bank deposit line item per Disbursement ID for the Adjusting Journal path.
      ' It exists so each disbursement's net amount can be recorded as the bank-side entry that offsets the related revenue and fee lines.
        
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Main fields used to stage the unique Disbursement IDs before building the bank deposit journal lines.
                wsBankDeposits.Range("A1:B1").Value = Array("Disbursement ID", "...........................")
            
            ' Adjusting Journal fields used to create the bank deposit import lines.
                wsBankDeposits.Range("C1:AH1").Value = Array("JOURNAL", "DATE", "REVERSEDATE", "DESCRIPTION", "REFERENCE_NO", "LINE_NO", "ACCT_NO", "LOCATION_ID", _
                        "DEPT_ID", "DOCUMENT", "MEMO", "DEBIT", "CREDIT", "SOURCEENTITY", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCHANGE_RATE", _
                        "STATE", "ALLOCATION_ID", "RASSET", "RDEPRECIATION_SCHEDULE", "RASSET_ADJUSTMENT", "RASSET_CLASS", "RASSETOUTOFSERVICE", _
                        "GLDIMFUNDING_SOURCE", "GLENTRY_PROJECTID", "GLENTRY_CUSTOMERID", "GLENTRY_VENDORID", "GLENTRY_EMPLOYEEID", "GLENTRY_ITEMID", "GLENTRY_CLASSID")
                        
        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' ..............................
            '          FILTER FIELD
            ' ..............................
                ' "Disbursement ID"
                  ' Build a unique list of Disbursement IDs so only one bank deposit line is created per disbursement.
                    wsBankDeposits.Range("A2").Formula2 = "=UNIQUE('" & wsDisbursementData.Name & "'!C2:C" & LastRow_DisbursementData & ")"
                    
                ' "..........................."
                    wsBankDeposits.Range("B2").Value = "..........................."
            
            ' ..............................
            '    ADJUSTING JOURNAL FIELDS
            ' ..............................
                ' "JOURNAL"
                    wsBankDeposits.Range("C2").Value = JournalName
                
                ' "DATE"
                  ' Use the disbursement date so the bank deposit line ties to the same posting date as the related disbursement activity.
                    wsBankDeposits.Range("D2").Formula2 = "=XLOOKUP($A2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!B:B)"
                
                ' "REVERSEDATE"
                    wsBankDeposits.Range("E2").Formula = "="""""
                
                ' "DESCRIPTION"
                  ' Use the Adjusting Journal Description from the Disbursement Data worksheet to keep the bank deposit line tied to the same disbursement-level description.
                    wsBankDeposits.Range("F2").Formula2 = "=XLOOKUP($A2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!K:K)"
                
                ' "REFERENCE_NO"
                    wsBankDeposits.Range("G2").Formula = "="""""
                
                ' "LINE_NO"
                  ' Bank deposit rows are created as one summarized line per disbursement.
                    wsBankDeposits.Range("H2").Formula = "=1"
                
                ' "ACCT_NO"
                  ' Convert the school abbreviation tied to the disbursement into the corresponding bank account.
                    wsBankDeposits.Range("I2").Formula2 = "=ConvertSchoolAbbrevToBankAccount(XLOOKUP($A2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!D:D))"
                
                ' "LOCATION_ID"
                  ' Convert the school abbreviation tied to the disbursement into the corresponding Intacct location.
                    wsBankDeposits.Range("J2").Formula2 = "=ConvertSchoolAbbrevToIntacctAccount(XLOOKUP($A2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!D:D))"
                
                ' "DEPT_ID"
                    wsBankDeposits.Range("K2").Formula = "=""2048"""
                
                ' "DOCUMENT"
                    wsBankDeposits.Range("L2").Formula = "="""""
                
                ' "MEMO"
                  ' Build a memo that clearly identifies the line as the bank deposit side of the disbursement.
                    wsBankDeposits.Range("M2").Formula2 = "=""Bank Deposit - ""&XLOOKUP($A2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!J:J)"
                
                ' "DEBIT"
                  ' Use the disbursement net amount as the bank deposit amount.
                    wsBankDeposits.Range("N2").Formula2 = "=XLOOKUP($A2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!I:I)"
                
                ' "CREDIT"
                    wsBankDeposits.Range("O2").Formula = "="""""
                
                ' "SOURCEENTITY"
                  ' Populate SourceEntity only when the Location ID ends in 00, using the last 3 digits of the bank account.
                    wsBankDeposits.Range("P2").Formula = "=IF(RIGHT(J2,2)=""00"",RIGHT(I2,3),"""")"
                
                ' "CURRENCY"
                    wsBankDeposits.Range("Q2").Formula = "="""""
                
                ' "EXCH_RATE_DATE"
                    wsBankDeposits.Range("R2").Formula = "="""""
                
                ' "EXCH_RATE_TYPE_ID"
                    wsBankDeposits.Range("S2").Formula = "="""""
                
                ' "EXCHANGE_RATE"
                    wsBankDeposits.Range("T2").Formula = "="""""
                
                ' "STATE"
                    wsBankDeposits.Range("U2").Formula = "=""Draft"""
                
                ' "ALLOCATION_ID"
                    wsBankDeposits.Range("V2").Formula = "="""""
                
                ' "RASSET"
                    wsBankDeposits.Range("W2").Formula = "="""""
                
                ' "RDEPRECIATION_SCHEDULE"
                    wsBankDeposits.Range("X2").Formula = "="""""
                
                ' "RASSET_ADJUSTMENT"
                    wsBankDeposits.Range("Y2").Formula = "="""""
                
                ' "RASSET_CLASS"
                    wsBankDeposits.Range("Z2").Formula = "="""""
                
                ' "RASSETOUTOFSERVICE"
                    wsBankDeposits.Range("AA2").Formula = "="""""
                
                ' "GLDIMFUNDING_SOURCE"
                    wsBankDeposits.Range("AB2").Formula = "=""7301-ATF Campaign"""
                
                ' "GLENTRY_PROJECTID"
                    wsBankDeposits.Range("AC2").Formula = "="""""
                
                ' "GLENTRY_CUSTOMERID"
                    wsBankDeposits.Range("AD2").Formula = "="""""
                
                ' "GLENTRY_VENDORID"
                    wsBankDeposits.Range("AE2").Formula = "="""""
                
                ' "GLENTRY_EMPLOYEEID"
                    wsBankDeposits.Range("AF2").Formula = "="""""
                
                ' "GLENTRY_ITEMID"
                    wsBankDeposits.Range("AG2").Formula = "="""""
                
                ' "GLENTRY_CLASSID"
                    wsBankDeposits.Range("AH2").Formula = "=""000"""
            
    ' ============================================================
    '  FIND THE LAST ROW FROM THE WORKSHEET
    ' ============================================================
        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            ' Use column A because it holds the unique Disbursement IDs that drive the worksheet.
                LastRow_BankDeposits = wsBankDeposits.Cells(wsBankDeposits.Rows.Count, 1).End(xlUp).Row
        
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            ' Fill down all formulas that are not already populated by the spilled UNIQUE formula in column A.
                If LastRow_BankDeposits > 2 Then
                    wsBankDeposits.Range("B2:AH" & LastRow_BankDeposits).FillDown
                End If

    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
            wsBankDeposits.Range("A1:AH1").AutoFilter
            wsBankDeposits.Columns("A:AH").AutoFit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' POPULATE THE CONNECTION ANALYSIS WORKSHEET '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
            Application.StatusBar = "Populating Connection Analysis Worksheet"
    
    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
        ' This worksheet compares Donation Site data against Salesforce data at the transaction level.
        ' It also stages the journal fields that will later flow into the User-Required Adjustments worksheet
        ' and the final import file paths.
        
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Main fields used to compare Donation Site and Salesforce values and identify transactions requiring review.
                wsConnectionAnalysis.Range("A1:K1").Value = Array("Transaction ID", "Disbursement ID", "Transaction Date", "Disbursement Date", _
                        "Donation Site - Gross Amount", "SF - Gross Amount", "Variance", "PMT-ID", "Donation Type", "Site - School Abbreviation", ".......")
            
            ' Adjusting Journal fields used to stage adjustment lines for the Adjusting Journal path.
                wsConnectionAnalysis.Range("L1:AR1").Value = Array("JOURNAL", "DATE", "REVERSEDATE", "DESCRIPTION", "REFERENCE_NO", "LINE_NO", "ACCT_NO", "LOCATION_ID", "DEPT_ID", _
                        "DOCUMENT", "MEMO", "DEBIT", "CREDIT", "SOURCEENTITY", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCHANGE_RATE", "STATE", _
                        "ALLOCATION_ID", "RASSET", "RDEPRECIATION_SCHEDULE", "RASSET_ADJUSTMENT", "RASSET_CLASS", "RASSETOUTOFSERVICE", "GLDIMFUNDING_SOURCE", _
                        "GLENTRY_PROJECTID", "GLENTRY_CUSTOMERID", "GLENTRY_VENDORID", "GLENTRY_EMPLOYEEID", "GLENTRY_ITEMID", "GLENTRY_CLASSID", "...........................")
            
            ' CRJ fields used to stage adjustment lines for the CRJ path.
                wsConnectionAnalysis.Range("AS1:BX1").Value = Array("RECEIPT_DATE", "PAYMETHOD", "DOCDATE", "DOCNUMBER", "DESCRIPTION", "DEPOSITTO", "BANKACCOUNTID", _
                        "DEPOSITDATE", "UNDEPACCTNO", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCH_RATE_DATE", "LINE_NO", "ACCT_NO", "ACCOUNTLABEL", "TRX_AMOUNT", _
                        "AMOUNT", "DEPT_ID", "LOCATION_ID", "ITEM_MEMO", "OTHERRECEIPTSENTRY_PROJECTID", "OTHERRECEIPTSENTRY_CUSTOMERID", "OTHERRECEIPTSENTRY_ITEMID", _
                        "OTHERRECEIPTSENTRY_VENDORID", "OTHERRECEIPTSENTRY_EMPLOYEEID", "OTHERRECEIPTSENTRY_CLASSID", "PAYER_NAME", "SUPDOCID", "EXCHANGE_RATE", _
                        "OR_TRANSACTION_DATE", "GLDIMFUNDING_SOURCE")
                
        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' "Transaction ID"
              ' Build a unique list of Donation Site Transaction IDs so each transaction is analyzed once.
                wsConnectionAnalysis.Range("A2").Formula2 = "=UNIQUE('" & wsStandardizedDonationSiteData.Name & "'!D2:D" & LastRow_StandardizedDonationSiteData & ")"
            
            ' "Disbursement ID"
              ' Pull the related Disbursement ID for each Transaction ID.
                wsConnectionAnalysis.Range("B2").Formula2 = "=XLOOKUP(A2,'" & wsStandardizedDonationSiteData.Name & "'!D:D,'" & wsStandardizedDonationSiteData.Name & "'!E:E)"
            
            ' "Transaction Date"
              ' Pull the Donation Site transaction date for reference and later review.
                wsConnectionAnalysis.Range("C2").Formula2 = "=XLOOKUP(A2,'" & wsStandardizedDonationSiteData.Name & "'!D:D,'" & wsStandardizedDonationSiteData.Name & "'!A:A)"
            
            ' "Disbursement Date"
              ' Pull the Donation Site disbursement date so the transaction can be tied to its disbursement timing.
                wsConnectionAnalysis.Range("D2").Formula2 = "=XLOOKUP(A2,'" & wsStandardizedDonationSiteData.Name & "'!D:D,'" & wsStandardizedDonationSiteData.Name & "'!B:B)"
            
            ' "Donation Site - Gross Amount"
              ' Sum the Donation Site gross amount by Transaction ID so it can be compared against Salesforce.
                wsConnectionAnalysis.Range("E2").Formula2 = "=SUMIFS('" & wsStandardizedDonationSiteData.Name & "'!K2:K" & LastRow_StandardizedDonationSiteData & ",'" & _
                        wsStandardizedDonationSiteData.Name & "'!D2:D" & LastRow_StandardizedDonationSiteData & ",A2)"
            
            ' "SF - Gross Amount"
              ' Sum the Salesforce gross amount by Transaction ID using the Relevant Transactions worksheet.
                wsConnectionAnalysis.Range("F2").Formula2 = "=SUMIFS('" & wsRelevantTransactions.Name & "'!Y:Y,'" & wsRelevantTransactions.Name & "'!G:G,A2)"
                
            ' "Variance"
              ' Calculate the difference between Donation Site and Salesforce gross amounts.
                wsConnectionAnalysis.Range("G2").Formula = "=ROUND(E2-F2,2)"
            
            ' "PMT-ID"
              ' Pull the matching PMT-ID from the standardized Salesforce data.
              ' If no PMT-ID is found, flag the transaction for later review.
                wsConnectionAnalysis.Range("H2").Formula2 = "=XLOOKUP($A2,'" & wsStandardizedSF.Name & "'!$B:$B,'" & wsStandardizedSF.Name & "'!F:F,""PMT-NOT MATCHED"")"
            
            ' "Donation Type"
              ' Pull the Donation Type from the standardized Donation Site data.
                wsConnectionAnalysis.Range("I2").Formula2 = "=XLOOKUP($A2,'" & wsStandardizedDonationSiteData.Name & "'!$D:$D,'" & wsStandardizedDonationSiteData.Name & "'!J:J)"
            
            ' "Site - School Abbreviation"
              ' Pull the corrected school abbreviation from the standardized Donation Site data.
                wsConnectionAnalysis.Range("J2").Formula2 = "=XLOOKUP($A2,'" & wsStandardizedDonationSiteData.Name & "'!$D:$D,'" & wsStandardizedDonationSiteData.Name & "'!P:P)"
            
            ' "......."
                wsConnectionAnalysis.Range("K2").Value = "......."

            ' ..............................
            '        ADJUSTING JOURNAL
            '         COLUMN HEADERS
            ' ..............................
                ' "JOURNAL"
                    wsConnectionAnalysis.Range("L2").Formula2 = JournalName
                    
                ' "DATE"
                  ' Use the disbursement date so any adjustment posts with the related disbursement activity.
                    wsConnectionAnalysis.Range("M2").Formula2 = "=D2"
                    
                ' "REVERSEDATE"
                    wsConnectionAnalysis.Range("N2").Formula = "="""""
                    
                ' "DESCRIPTION"
                  ' Use the Adjusting Journal Description tied to the Disbursement ID.
                    wsConnectionAnalysis.Range("O2").Formula2 = "=XLOOKUP($B2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!K:K)"
                    
                ' "REFERENCE_NO"
                    wsConnectionAnalysis.Range("P2").Formula = "="""""
                    
                ' "LINE_NO"
                    wsConnectionAnalysis.Range("Q2").Formula = "="""""
                    
                ' "ACCT_NO"
                  ' Use the revenue account based on Donation Type.
                    wsConnectionAnalysis.Range("R2").Formula = "=IF(I2=""Employer Matching"",""73013"",""73011"")"
                    
                ' "LOCATION_ID"
                  ' Convert the school abbreviation into the matching Intacct location.
                    wsConnectionAnalysis.Range("S2").Formula2 = "=ConvertSchoolAbbrevToIntacctAccount(J2)"
                    
                ' "DEPT_ID"
                    wsConnectionAnalysis.Range("T2").Formula = "=""2048"""
                    
                ' "DOCUMENT"
                    wsConnectionAnalysis.Range("U2").Formula = "="""""
                    
                ' "MEMO"
                  ' Build a memo that ties the adjustment back to the PMT-ID, Transaction ID, and Disbursement ID.
                    wsConnectionAnalysis.Range("V2").Formula = "=""Payment Adjustment: ""&XLOOKUP(H2,'" & wsRelevantTransactions.Name & "'!K:K,'" & _
                            wsRelevantTransactions.Name & "'!S:S,""Transaction ID: ""&A2&"" | Disbursement ID: ""&B2&"" | ""&H2)"
                    
                ' "DEBIT"
                  ' Use a debit when the Donation Site gross amount is lower than Salesforce.
                    wsConnectionAnalysis.Range("W2").Formula = "=IF(G2<0,G2*-1,"""")"
                    
                ' "CREDIT"
                  ' Use a credit when the Donation Site gross amount is higher than Salesforce.
                    wsConnectionAnalysis.Range("X2").Formula = "=IF(G2>0,G2,"""")"
                    
                ' "SOURCEENTITY"
                    wsConnectionAnalysis.Range("Y2").Formula = "="""""
                    
                ' "CURRENCY"
                    wsConnectionAnalysis.Range("Z2").Formula = "="""""
                    
                ' "EXCH_RATE_DATE"
                    wsConnectionAnalysis.Range("AA2").Formula = "="""""
                    
                ' "EXCH_RATE_TYPE_ID"
                    wsConnectionAnalysis.Range("AB2").Formula = "="""""
                    
                ' "EXCHANGE_RATE"
                    wsConnectionAnalysis.Range("AC2").Formula = "="""""
                    
                ' "STATE"
                    wsConnectionAnalysis.Range("AD2").Formula = "=""Draft"""
                    
                ' "ALLOCATION_ID"
                    wsConnectionAnalysis.Range("AE2").Formula = "="""""
                    
                ' "RASSET"
                    wsConnectionAnalysis.Range("AF2").Formula = "="""""
                    
                ' "RDEPRECIATION_SCHEDULE"
                    wsConnectionAnalysis.Range("AG2").Formula = "="""""
                    
                ' "RASSET_ADJUSTMENT"
                    wsConnectionAnalysis.Range("AH2").Formula = "="""""
                    
                ' "RASSET_CLASS"
                    wsConnectionAnalysis.Range("AI2").Formula = "="""""
                    
                ' "RASSETOUTOFSERVICE"
                    wsConnectionAnalysis.Range("AJ2").Formula = "="""""
                    
                ' "GLDIMFUNDING_SOURCE"
                    wsConnectionAnalysis.Range("AK2").Formula = "=""7301-ATF Campaign"""
                    
                ' "GLENTRY_PROJECTID"
                    wsConnectionAnalysis.Range("AL2").Formula = "="""""
                    
                ' "GLENTRY_CUSTOMERID"
                    wsConnectionAnalysis.Range("AM2").Formula = "="""""
                    
                ' "GLENTRY_VENDORID"
                    wsConnectionAnalysis.Range("AN2").Formula = "="""""
                    
                ' "GLENTRY_EMPLOYEEID"
                    wsConnectionAnalysis.Range("AO2").Formula = "="""""
                    
                ' "GLENTRY_ITEMID"
                    wsConnectionAnalysis.Range("AP2").Formula = "="""""
                    
                ' "GLENTRY_CLASSID"
                    wsConnectionAnalysis.Range("AQ2").Formula = "=""000"""
                    
                ' "..........................."
                    wsConnectionAnalysis.Range("AR2").Formula2 = "..........................."
                    
            ' ..............................
            '       CRJ COLUMN HEADERS
            ' ..............................
                ' "RECEIPT_DATE"
                  ' Use the disbursement date so the CRJ adjustment ties to the same disbursement activity.
                    wsConnectionAnalysis.Range("AS2").Formula = "=D2"
                    
                ' "PAYMETHOD"
                    wsConnectionAnalysis.Range("AT2").Formula = "=""Credit Card"""
                    
                ' "DOCDATE"
                  ' Use the disbursement date as the CRJ document date.
                    wsConnectionAnalysis.Range("AU2").Formula = "=D2"
                    
                ' "DOCNUMBER"
                  ' Use the Donation Site name as the CRJ document number placeholder.
                    wsConnectionAnalysis.Range("AV2").Formula2 = DonationSite
                    
                ' "DESCRIPTION"
                  ' Use the CRJ Description tied to the Disbursement ID.
                    wsConnectionAnalysis.Range("AW2").Formula2 = "=XLOOKUP($B2,'" & wsDisbursementData.Name & "'!$C:$C,'" & wsDisbursementData.Name & "'!J:J)"
                    
                ' "DEPOSITTO"
                    wsConnectionAnalysis.Range("AX2").Formula = "=""Bank account"""
                    
                ' "BANKACCOUNTID"
                  ' Convert the school abbreviation into the matching bank account name.
                    wsConnectionAnalysis.Range("AY2").Formula2 = "=ConvertSchoolAbbrevToBankAccountName(J2)"
                    
                ' "DEPOSITDATE"
                    wsConnectionAnalysis.Range("AZ2").Formula2 = "=D2"
                    
                ' "UNDEPACCTNO"
                    wsConnectionAnalysis.Range("BA2").Formula = "="""""
                    
                ' "CURRENCY"
                    wsConnectionAnalysis.Range("BB2").Formula2 = "=""USD"""
                    
                ' "EXCH_RATE_DATE"
                    wsConnectionAnalysis.Range("BC2").Formula = "="""""
                    
                ' "EXCH_RATE_TYPE_ID"
                    wsConnectionAnalysis.Range("BD2").Formula = "="""""
                    
                ' "EXCHANGE_RATE"
                    wsConnectionAnalysis.Range("BE2").Formula = "="""""
                    
                ' "LINE_NO"
                    wsConnectionAnalysis.Range("BF2").Formula = "="""""
                    
                ' "ACCT_NO"
                  ' Use the revenue account based on Donation Type.
                    wsConnectionAnalysis.Range("BG2").Formula2 = "=IF(I2=""Employer Matching"",""73013"",""73011"")"
                    
                ' "ACCOUNTLABEL"
                    wsConnectionAnalysis.Range("BH2").Formula = "="""""
                    
                ' "TRX_AMOUNT"
                  ' Use the variance amount as the CRJ transaction amount.
                    wsConnectionAnalysis.Range("BI2").Formula = "=G2"
                    
                ' "AMOUNT"
                  ' Use the variance amount as the CRJ amount.
                    wsConnectionAnalysis.Range("BJ2").Formula = "=G2"
                    
                ' "DEPT_ID"
                    wsConnectionAnalysis.Range("BK2").Formula2 = "=""2048"""
                    
                ' "LOCATION_ID"
                  ' Convert the school abbreviation into the matching Intacct location.
                    wsConnectionAnalysis.Range("BL2").Formula2 = "=ConvertSchoolAbbrevToIntacctAccount(J2)"
                    
                ' "ITEM_MEMO"
                  ' Build a memo that ties the adjustment back to the PMT-ID, Transaction ID, and Disbursement ID.
                    wsConnectionAnalysis.Range("BM2").Formula2 = "=""Payment Adjustment: ""&XLOOKUP(H2,'" & wsRelevantTransactions.Name & "'!K:K,'" & _
                                wsRelevantTransactions.Name & "'!S:S,""Transaction ID: ""&A2&"" | Disbursement ID: ""&B2&"" | ""&H2)"
                
                ' "OTHERRECEIPTSENTRY_PROJECTID"
                    wsConnectionAnalysis.Range("BN2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_CUSTOMERID"
                    wsConnectionAnalysis.Range("BO2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_ITEMID"
                    wsConnectionAnalysis.Range("BP2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_VENDORID"
                    wsConnectionAnalysis.Range("BQ2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_EMPLOYEEID"
                    wsConnectionAnalysis.Range("BR2").Formula = "="""""
                    
                ' "OTHERRECEIPTSENTRY_CLASSID"
                    wsConnectionAnalysis.Range("BS2").Formula2 = "=""000"""
                    
                ' "PAYER_NAME"
                    wsConnectionAnalysis.Range("BT2").Formula2 = DonationSite
                    
                ' "SUPDOCID"
                    wsConnectionAnalysis.Range("BU2").Formula = "="""""
                    
                ' "EXCHANGE_RATE"
                    wsConnectionAnalysis.Range("BV2").Formula = "="""""
                    
                ' "OR_TRANSACTION_DATE"
                    wsConnectionAnalysis.Range("BW2").Formula2 = "=D2"
                    
                ' "GLDIMFUNDING_SOURCE"
                    wsConnectionAnalysis.Range("BX2").Formula2 = "=""7301-ATF Campaign"""
                    
    
    ' ============================================================
    '   FIND THE LAST ROW FROM THE CONNECTION ANALYSIS WORKSHEET
    ' ============================================================
        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            ' Use column A because it holds the unique Transaction IDs that drive the worksheet.
                LastRow_ConnectionAnalysis = wsConnectionAnalysis.Cells(wsConnectionAnalysis.Rows.Count, 1).End(xlUp).Row
        
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            ' Fill down all non-spilled formulas across the worksheet.
                If LastRow_ConnectionAnalysis > 2 Then
                    wsConnectionAnalysis.Range("B2:BX" & LastRow_ConnectionAnalysis).FillDown
                End If

    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        wsConnectionAnalysis.Range("A1:BX1").AutoFilter
        wsConnectionAnalysis.Columns("A:BX").AutoFit
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' POPULATE THE USER-REQUIRED ADJUSTMENTS WORKSHEET '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '                  UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "Populating the User-Required Adjustments Worksheet"

    ' ============================================================
    '        PURPOSE OF THE USER-REQUIRED ADJUSTMENTS WORKSHEET
    ' ============================================================
        ' This worksheet acts as the exception-management area for the converter.
        '
        ' The earlier worksheets standardize and connect data across:
        '   1. the Donation Site Reports
        '   2. the Initial Report (Salesforce or Intacct)
        '   3. the Disbursement-level summary data
        '
        ' By the time the macro reaches this point, most records should already be usable.
        ' Any records that still contain missing, conflicting, or unresolved information are
        ' intentionally redirected here so the user can review them before final import.
        '
        ' This matters because unresolved exceptions can cause:
        '   - revenue to post to the wrong school
        '   - revenue to post to the wrong account, division, or funding source
        '   - disbursements to remain unmatched to expected bank allocation logic
        '   - variances between Donation Site and Salesforce amounts
        '   - transactions to be imported without a valid PMT-ID connection
        '
        ' Each section below isolates one type of issue so the user can work through it in a
        ' controlled, visible way instead of hunting through multiple worksheets.
        '
        ' The journal-name fields that are populated in each section are especially important.
        ' They are later used to identify which journal descriptions should be excluded from the
        ' final import file until the related issue has been resolved.
        '
        ' A section turning green means no exceptions were found for that category.
        ' A red section means at least one transaction or disbursement requires user review.
        '
        ' The grouped rows are used so the worksheet can remain readable even when all sections
        ' are present. The user can expand only the sections that need attention.

    ' ============================================================
    '                  BANK ALLOCATIONS NOT FOUND
    ' ============================================================
        ' This section captures transactions whose Donation Site school assignment could not be
        ' converted into a usable BASIS school abbreviation.
        '
        ' Why this matters:
        ' The school abbreviation is a critical link used to derive downstream values such as:
        '   - Intacct Location ID
        '   - Bank Account / Bank Allocation logic
        '   - School-level posting destination
        '
        ' If the converter cannot identify the school correctly, then it cannot safely determine
        ' where that money belongs. Rather than guessing, those records are sent here so the user
        ' can manually assign the correct school.
        '
        ' Once the school is corrected, the matching journal name can later be released back into
        ' the final import flow.

        ' ---------------------------------------------
        '             INITIATE ROW VARIABLES
        ' ---------------------------------------------
            SectionHeaderRow_UserRequiredAdjustments = 1
            HeaderRow_UserRequiredAdjustments = SectionHeaderRow_UserRequiredAdjustments + 1
            DataStartRow_UserRequiredAdjustments = HeaderRow_UserRequiredAdjustments + 1
            
        ' ---------------------------------------------
        '                 SECTION HEADER
        ' ---------------------------------------------
            With wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments)
                .Merge
                .HorizontalAlignment = xlCenter
                .Value = "BANK ALLOCATIONS NOT FOUND"
                .Interior.Color = vbRed
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleSingle
            End With
            
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            wsUserRequiredAdjustments.Range("A" & HeaderRow_UserRequiredAdjustments & ":N" & HeaderRow_UserRequiredAdjustments).Value = _
                Array("Disbursement ID", "Transaction ID", "Transaction Date", "Disbursement Date", "Donation Type", "Site - School Name", _
                      "Site - School Abbreviation", "Corrected - School", "Corrected - School Abbreviation", "", "", _
                      "Adjusting Journal Name", "CRJ Journal Name", "File Name")

        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' Pull only Donation Site rows where the converted school abbreviation equals
            ' "No School Found".
            '
            ' If nothing is found, return "All Bank Allocations Found" so the section remains
            ' explicit and readable.
            wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IFERROR(IF(ISBLANK(CHOOSECOLS(FILTER('" & wsStandardizedDonationSiteData.Name & "'!A2:O" & LastRow_StandardizedDonationSiteData & _
                ",'" & wsStandardizedDonationSiteData.Name & "'!O2:O" & LastRow_StandardizedDonationSiteData & _
                "=""No School Found""),4,5,1,2,10,14,15)),""""," & _
                "CHOOSECOLS(FILTER('" & wsStandardizedDonationSiteData.Name & "'!A2:O" & LastRow_StandardizedDonationSiteData & ",'" & _
                wsStandardizedDonationSiteData.Name & "'!O2:O" & LastRow_StandardizedDonationSiteData & _
                "=""No School Found""),5,4,1,2,10,14,15)),""All Bank Allocations Found"")"
                
            ' Column H is intentionally left for user input.
            ' The user selects the corrected school name from a validation dropdown.
            
            ' This column stores the school abbreviation derived from the user's corrected school.
            wsUserRequiredAdjustments.Range("I" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(H" & DataStartRow_UserRequiredAdjustments & "="""","""",ConvertSchoolNameToSchoolAbbrev(H" & DataStartRow_UserRequiredAdjustments & "))"
                
            ' Keep the related Adjusting Journal description active until the issue is resolved.
            wsUserRequiredAdjustments.Range("L" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All Bank Allocations Found"",""""," & _
                "IFERROR(IF(I" & DataStartRow_UserRequiredAdjustments & "="""",XLOOKUP(A" & DataStartRow_UserRequiredAdjustments & ",'" & _
                wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!K:K),""CLEARED""),""""))"
            
            ' Same logic as column L, but for the CRJ description.
            wsUserRequiredAdjustments.Range("M" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All Bank Allocations Found"",""""," & _
                "IFERROR(IF(I" & DataStartRow_UserRequiredAdjustments & "="""",XLOOKUP(A" & DataStartRow_UserRequiredAdjustments & ",'" & _
                wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!J:J),""CLEARED""),""""))"

            ' Store the original file name tied to the disbursement for traceability and
            ' possible process-later handling.
            wsUserRequiredAdjustments.Range("N" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All Bank Allocations Found"",""""," & _
                "XLOOKUP(A" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!M:M))"

        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            LastRow_UserRequiredAdjustments = wsUserRequiredAdjustments.Cells(wsUserRequiredAdjustments.Rows.Count, 1).End(xlUp).Row
                
        ' ---------------------------------------------
        '                 GROUP SECTION
        ' ---------------------------------------------
            wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments & ":" & (LastRow_UserRequiredAdjustments + 4)).Group
                
        ' ---------------------------------------------
        '       CREATE THE DATA VALIDATION RULES
        ' ---------------------------------------------
            ' Only create the validation list if exception rows actually exist.
            If wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Value <> "All Bank Allocations Found" Then
                
                With wsUserRequiredAdjustments.Range("H" & DataStartRow_UserRequiredAdjustments & ":H" & LastRow_UserRequiredAdjustments).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Rng_SchoolValidation_SchoolNames
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .InputTitle = "Select Correct School"
                    .ErrorTitle = "Invalid School"
                    .InputMessage = "Select the correct school from the dropdown list."
                    .ErrorMessage = "Please select a school from the approved validation list."
                    .ShowInput = True
                    .ShowError = True
                End With
                
            End If
                
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            ' Column A spills. Columns I:N must be filled down because they depend on the row.
            If LastRow_UserRequiredAdjustments <> DataStartRow_UserRequiredAdjustments Then
                wsUserRequiredAdjustments.Range("I" & DataStartRow_UserRequiredAdjustments & ":N" & LastRow_UserRequiredAdjustments).FillDown
            Else
                wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments).Interior.Color = vbGreen
                wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments).ShowDetail = False
            End If
        
        ' ---------------------------------------------
        '             SAVE SECTION VARIABLES
        ' ---------------------------------------------
            DataStartRow_UserRequiredAdjustments_BankAllocations = DataStartRow_UserRequiredAdjustments
            LastRow_UserRequiredAdjustments_BankAllocations = LastRow_UserRequiredAdjustments
            
            ' Save the unresolved journal-name range for later exclusion from the final import file.
            ' If the section only contains one row, save a single-cell reference instead of a same-cell range.
            If JournalType = "Adjusting" Then
                If DataStartRow_UserRequiredAdjustments_BankAllocations <> LastRow_UserRequiredAdjustments_BankAllocations Then
                    Rng_UserRequiredAdjustments_BankAllocations = "'" & wsUserRequiredAdjustments.Name & "'!L" & _
                        DataStartRow_UserRequiredAdjustments_BankAllocations & ":L" & _
                        LastRow_UserRequiredAdjustments_BankAllocations
                Else
                    Rng_UserRequiredAdjustments_BankAllocations = "'" & wsUserRequiredAdjustments.Name & "'!L" & _
                        DataStartRow_UserRequiredAdjustments_BankAllocations
                End If
            Else
                If DataStartRow_UserRequiredAdjustments_BankAllocations <> LastRow_UserRequiredAdjustments_BankAllocations Then
                    Rng_UserRequiredAdjustments_BankAllocations = "'" & wsUserRequiredAdjustments.Name & "'!M" & _
                        DataStartRow_UserRequiredAdjustments_BankAllocations & ":M" & _
                        LastRow_UserRequiredAdjustments_BankAllocations
                Else
                    Rng_UserRequiredAdjustments_BankAllocations = "'" & wsUserRequiredAdjustments.Name & "'!M" & _
                        DataStartRow_UserRequiredAdjustments_BankAllocations
                End If
            End If
       
    ' ---------------------------------------------
    '           UPDATE ALL RELEVANT FORMULAS
    ' ---------------------------------------------
        ' Now that the Bank Allocations section exists, the converter has a usable lookup range
        ' for corrected school abbreviations.
        '
        ' This update could not be finalized earlier because the lookup range did not exist yet.
        ' At the time the Standardized Donation Site Data worksheet was first built, the macro had
        ' not yet created this exception section or its row boundaries.
        '
        ' Once the section has been populated and its start/end rows are known, the formula in
        ' column P of the Standardized Donation Site Data worksheet can now safely use that range
        ' to look up missing school abbreviations.
        '
        ' This matters because the corrected school abbreviation needs to flow back upstream into
        ' the standardized data so downstream worksheets use the user's correction instead of
        ' continuing to carry "No School Found".
        wsStandardizedDonationSiteData.Range("P2").Formula2 = _
            "=IF(O2=""No School Found"",XLOOKUP(E2,'" & wsUserRequiredAdjustments.Name & "'!$A$" & _
                DataStartRow_UserRequiredAdjustments_BankAllocations & ":$A$" & LastRow_UserRequiredAdjustments_BankAllocations & _
                ",'" & wsUserRequiredAdjustments.Name & "'!$I$" & _
                DataStartRow_UserRequiredAdjustments_BankAllocations & ":$I$" & LastRow_UserRequiredAdjustments_BankAllocations & ",O2),O2)"
                        
        If LastRow_StandardizedDonationSiteData > 2 Then
            wsStandardizedDonationSiteData.Range("P2:P" & LastRow_StandardizedDonationSiteData).FillDown
        End If

    ' ============================================================
    '               TRANSACTIONS MISSING SCHOOL NAME
    ' ============================================================
        ' This section captures Salesforce-side transactions where the school / location assignment
        ' could not be determined.
        '
        ' Why this matters:
        ' Even if the Donation Site data is usable, the Salesforce-side accounting fields still need
        ' a valid school assignment so the correct Intacct location and account logic can be applied.
        ' If the school name is missing here, the transaction cannot be trusted for final posting.

        ' ---------------------------------------------
        '              UPDATE ROW VARIABLES
        ' ---------------------------------------------
            SectionHeaderRow_UserRequiredAdjustments = LastRow_UserRequiredAdjustments + 6
            HeaderRow_UserRequiredAdjustments = SectionHeaderRow_UserRequiredAdjustments + 1
            DataStartRow_UserRequiredAdjustments = HeaderRow_UserRequiredAdjustments + 1
        
        ' ---------------------------------------------
        '                 SECTION HEADER
        ' ---------------------------------------------
            With wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments)
                .Merge
                .HorizontalAlignment = xlCenter
                .Value = "TRANSACTIONS MISSING SCHOOL NAME"
                .Interior.Color = vbRed
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleSingle
            End With

        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            wsUserRequiredAdjustments.Range("A" & HeaderRow_UserRequiredAdjustments & ":N" & HeaderRow_UserRequiredAdjustments).Value = _
                Array("Transaction ID", "Disbursement ID", "SF Payment ID", "Primary Contact", "Account Name", "Company Name", "Campaign Name", "Opportunity Name", _
                      "Corrected School", "Corrected School Account", "", _
                      "Adjusting Journal Name", "CRJ Journal Name", "File Name")

        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IFERROR(CHOOSECOLS(FILTER('" & wsStandardizedSF.Name & "'!B2:K" & LastRow_StandardizedSF & _
                ",'" & wsStandardizedSF.Name & "'!M2:M" & LastRow_StandardizedSF & "=""No School Found""),1,2,5,6,7,8,9,10),""All School Names Found"")"

            ' Column I is user input for the corrected school.
            
            ' Derive the school abbreviation from the corrected school name.
            wsUserRequiredAdjustments.Range("J" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(I" & DataStartRow_UserRequiredAdjustments & "="""","""",ConvertSchoolNameToSchoolAbbrev(I" & DataStartRow_UserRequiredAdjustments & "))"
                
            ' Convert the corrected school abbreviation into the corrected Intacct school account.
            wsUserRequiredAdjustments.Range("K" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(J" & DataStartRow_UserRequiredAdjustments & "="""","""",ConvertSchoolAbbrevToIntacctAccount(J" & DataStartRow_UserRequiredAdjustments & "))"
                
            ' Keep the related Adjusting Journal description active until the issue is resolved.
            wsUserRequiredAdjustments.Range("L" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All School Names Found"",""""," & _
                "IF(K" & DataStartRow_UserRequiredAdjustments & "="""",XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & _
                wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!K:K),""CLEARED""))"

            ' Same logic as column L, but for the CRJ route.
            wsUserRequiredAdjustments.Range("M" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All School Names Found"",""""," & _
                "IF(K" & DataStartRow_UserRequiredAdjustments & "="""",XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & _
                wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!J:J),""CLEARED""))"

            ' Store the source file tied to the disbursement.
            wsUserRequiredAdjustments.Range("N" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All School Names Found"",""""," & _
                "XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!M:M))"
                    
        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            LastRow_UserRequiredAdjustments = wsUserRequiredAdjustments.Cells(wsUserRequiredAdjustments.Rows.Count, 1).End(xlUp).Row
                
        ' ---------------------------------------------
        '                 GROUP SECTION
        ' ---------------------------------------------
            wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments & ":" & (LastRow_UserRequiredAdjustments + 4)).Group
               
        ' ---------------------------------------------
        '       CREATE THE DATA VALIDATION RULES
        ' ---------------------------------------------
            If wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Value <> "All School Names Found" Then
                
                With wsUserRequiredAdjustments.Range("I" & DataStartRow_UserRequiredAdjustments & ":I" & LastRow_UserRequiredAdjustments).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Rng_SchoolValidation_SchoolNames
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .InputTitle = "Select Correct School"
                    .ErrorTitle = "Invalid School"
                    .InputMessage = "Select the correct school from the dropdown list."
                    .ErrorMessage = "Please select a school from the approved validation list."
                    .ShowInput = True
                    .ShowError = True
                End With
                
            End If
        
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            If LastRow_UserRequiredAdjustments <> DataStartRow_UserRequiredAdjustments Then
                wsUserRequiredAdjustments.Range("J" & DataStartRow_UserRequiredAdjustments & ":N" & LastRow_UserRequiredAdjustments).FillDown
            Else
                wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments).Interior.Color = vbGreen
                wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments).ShowDetail = False
            End If
        
        ' ---------------------------------------------
        '             SAVE SECTION VARIABLES
        ' ---------------------------------------------
            DataStartRow_UserRequiredAdjustments_MissingSchoolNames = DataStartRow_UserRequiredAdjustments
            LastRow_UserRequiredAdjustments_MissingSchoolNames = LastRow_UserRequiredAdjustments
            
            If JournalType = "Adjusting" Then
                If DataStartRow_UserRequiredAdjustments_MissingSchoolNames <> LastRow_UserRequiredAdjustments_MissingSchoolNames Then
                    Rng_UserRequiredAdjustments_MissingSchoolNames = "'" & wsUserRequiredAdjustments.Name & "'!L" & _
                        DataStartRow_UserRequiredAdjustments_MissingSchoolNames & ":L" & _
                        LastRow_UserRequiredAdjustments_MissingSchoolNames
                Else
                    Rng_UserRequiredAdjustments_MissingSchoolNames = "'" & wsUserRequiredAdjustments.Name & "'!L" & _
                        DataStartRow_UserRequiredAdjustments_MissingSchoolNames
                End If
            Else
                If DataStartRow_UserRequiredAdjustments_MissingSchoolNames <> LastRow_UserRequiredAdjustments_MissingSchoolNames Then
                    Rng_UserRequiredAdjustments_MissingSchoolNames = "'" & wsUserRequiredAdjustments.Name & "'!M" & _
                        DataStartRow_UserRequiredAdjustments_MissingSchoolNames & ":M" & _
                        LastRow_UserRequiredAdjustments_MissingSchoolNames
                Else
                    Rng_UserRequiredAdjustments_MissingSchoolNames = "'" & wsUserRequiredAdjustments.Name & "'!M" & _
                        DataStartRow_UserRequiredAdjustments_MissingSchoolNames
                End If
            End If
            
    ' ---------------------------------------------
    '           UPDATE ALL RELEVANT FORMULAS
    ' ---------------------------------------------
        ' Now that the Missing School Name section exists, the converter has a usable lookup range
        ' for corrected school accounts.
        '
        ' This is important because when the Standardized Salesforce worksheet was first built,
        ' the user had not yet had an opportunity to correct any missing school assignments.
        ' That means the formulas in the standardized worksheet could not yet look back to a
        ' completed exception section.
        '
        ' Now that this section exists and its row boundaries are known, the macro can update the
        ' Location Correction field so it uses the user-supplied school-account correction where
        ' the original school value was missing.
        If InitialPath = "Salesforce" Then
            
            wsStandardizedSF.Range("S2").Formula2 = _
                "=IF(M2=""No School Found"",XLOOKUP(F2,'" & wsUserRequiredAdjustments.Name & "'!$C$" & _
                    DataStartRow_UserRequiredAdjustments_MissingSchoolNames & ":$C$" & LastRow_UserRequiredAdjustments_MissingSchoolNames & _
                    ",'" & wsUserRequiredAdjustments.Name & "'!$K$" & _
                    DataStartRow_UserRequiredAdjustments_MissingSchoolNames & ":$K$" & LastRow_UserRequiredAdjustments_MissingSchoolNames & ",M2),M2)"
                        
            If LastRow_StandardizedSF > 2 Then
                wsStandardizedSF.Range("S2:S" & LastRow_StandardizedSF).FillDown
            End If
            
        End If

    ' ============================================================
    '        ADJUSTMENTS TO: ACCOUNT|DIVISION|FUNDING SOURCE
    ' ============================================================
        ' This section captures transactions where the converter could not confidently derive the
        ' required accounting dimensions from the standardized Salesforce data.
        '
        ' Why this matters:
        ' Even if the transaction is otherwise valid, it should not be imported if its posting
        ' dimensions are wrong. Incorrect account / division / funding source values can lead to
        ' materially incorrect accounting results.

        ' ---------------------------------------------
        '              UPDATE ROW VARIABLES
        ' ---------------------------------------------
            SectionHeaderRow_UserRequiredAdjustments = LastRow_UserRequiredAdjustments + 6
            HeaderRow_UserRequiredAdjustments = SectionHeaderRow_UserRequiredAdjustments + 1
            DataStartRow_UserRequiredAdjustments = HeaderRow_UserRequiredAdjustments + 1
        
        ' ---------------------------------------------
        '                 SECTION HEADER
        ' ---------------------------------------------
            With wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments)
                .Merge
                .HorizontalAlignment = xlCenter
                .Value = "ADJUSTMENTS TO: ACCOUNT|DIVISION|FUNDING SOURCE"
                .Interior.Color = vbRed
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleSingle
            End With
            
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            wsUserRequiredAdjustments.Range("A" & HeaderRow_UserRequiredAdjustments & ":N" & HeaderRow_UserRequiredAdjustments).Value = _
                Array("Transaction ID", "Disbursement ID", "SF Payment ID", "Primary Contact", "Account Name", "Company Name", "Campaign Name", _
                      "Opportunity Name", "Account Correction", "Division Correction", "Funding Source Correction", _
                      "Adjusting Journal Name", "CRJ Journal Name", "File Name")

        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IFERROR(CHOOSECOLS(FILTER('" & wsStandardizedSF.Name & "'!B2:K" & LastRow_StandardizedSF & ",('" & _
                wsStandardizedSF.Name & "'!N2:N" & LastRow_StandardizedSF & "=73998)+('" & wsStandardizedSF.Name & "'!N2:N" & LastRow_StandardizedSF & _
                "=""CHECK"")),1,2,5,6,7,8,9,10),""All Accounts, Divisions, and Funding Sources Found"")"
                        
            ' Columns I:K are intentionally left for user correction / override.
            
            wsUserRequiredAdjustments.Range("L" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All Accounts, Divisions, and Funding Sources Found"",""""," & _
                "IF(AND(I" & DataStartRow_UserRequiredAdjustments & "<>"""",J" & _
                DataStartRow_UserRequiredAdjustments & "<>"""",K" & DataStartRow_UserRequiredAdjustments & "<>""""),""CLEARED"",XLOOKUP(B" & _
                DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!K:K)))"

            wsUserRequiredAdjustments.Range("M" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All Accounts, Divisions, and Funding Sources Found"",""""," & _
                "IF(AND(I" & DataStartRow_UserRequiredAdjustments & "<>"""",J" & _
                DataStartRow_UserRequiredAdjustments & "<>"""",K" & DataStartRow_UserRequiredAdjustments & "<>""""),""CLEARED"",XLOOKUP(B" & _
                DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!J:J)))"

            wsUserRequiredAdjustments.Range("N" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All Accounts, Divisions, and Funding Sources Found"",""""," & _
                "XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!M:M))"

        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            LastRow_UserRequiredAdjustments = wsUserRequiredAdjustments.Cells(wsUserRequiredAdjustments.Rows.Count, 1).End(xlUp).Row
                
        ' ---------------------------------------------
        '                 GROUP SECTION
        ' ---------------------------------------------
            wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments & ":" & (LastRow_UserRequiredAdjustments + 4)).Group
                
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            If LastRow_UserRequiredAdjustments <> DataStartRow_UserRequiredAdjustments Then
                wsUserRequiredAdjustments.Range("L" & DataStartRow_UserRequiredAdjustments & ":N" & LastRow_UserRequiredAdjustments).FillDown
            Else
                wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments).Interior.Color = vbGreen
                wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments).ShowDetail = False
            End If
        
        ' ---------------------------------------------
        '             SAVE SECTION VARIABLES
        ' ---------------------------------------------
            DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments = DataStartRow_UserRequiredAdjustments
            LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments = LastRow_UserRequiredAdjustments
            
            If JournalType = "Adjusting" Then
                If DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments <> LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments Then
                    Rng_UserRequiredAdjustments_AccountDivisionFundingAdjustments = "'" & wsUserRequiredAdjustments.Name & "'!L" & _
                        DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ":L" & _
                        LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments
                Else
                    Rng_UserRequiredAdjustments_AccountDivisionFundingAdjustments = "'" & wsUserRequiredAdjustments.Name & "'!L" & _
                        DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments
                End If
            Else
                If DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments <> LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments Then
                    Rng_UserRequiredAdjustments_AccountDivisionFundingAdjustments = "'" & wsUserRequiredAdjustments.Name & "'!M" & _
                        DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ":M" & _
                        LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments
                Else
                    Rng_UserRequiredAdjustments_AccountDivisionFundingAdjustments = "'" & wsUserRequiredAdjustments.Name & "'!M" & _
                        DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments
                End If
            End If
    
    ' ---------------------------------------------
    '           UPDATE ALL RELEVANT FORMULAS
    ' ---------------------------------------------
        ' Now that the user has a populated section for account / division / funding-source
        ' exceptions, the converter can finally update the standardized worksheet formulas that
        ' depend on those user-entered corrections.
        '
        ' This must happen here rather than earlier because these lookup ranges did not exist yet
        ' when the standardized worksheet was first built. At that earlier point, the macro did not
        ' yet know which rows would require review, or what row boundaries the exception section
        ' would ultimately occupy.
        '
        ' Once this section exists, the formulas can safely reference the saved range boundaries and
        ' look up corrected values for:
        '   - Account Correction
        '   - Division Correction
        '   - Funding Source Correction
        '
        ' This allows downstream worksheets to use the corrected values instead of continuing to
        ' carry placeholder values such as 73998 or "CHECK".
        If InitialPath = "Salesforce" Then
            
            wsStandardizedSF.Range("T2").Formula2 = _
                "=IF(N2=""CHECK"",XLOOKUP(F2,'" & wsUserRequiredAdjustments.Name & "'!$C$" & _
                    DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ":$C$" & LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & _
                    ",'" & wsUserRequiredAdjustments.Name & "'!$I$" & _
                    DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ":$I$" & LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ",N2),N2)"

            wsStandardizedSF.Range("U2").Formula2 = _
                "=IF(O2=""CHECK"",XLOOKUP(F2,'" & wsUserRequiredAdjustments.Name & "'!$C$" & _
                    DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ":$C$" & LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & _
                    ",'" & wsUserRequiredAdjustments.Name & "'!$J$" & _
                    DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ":$J$" & LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ",O2),O2)"
            
            wsStandardizedSF.Range("V2").Formula2 = _
                "=IF(P2=""CHECK"",XLOOKUP(F2,'" & wsUserRequiredAdjustments.Name & "'!$C$" & _
                    DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ":$C$" & LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & _
                    ",'" & wsUserRequiredAdjustments.Name & "'!$K$" & _
                    DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ":$K$" & LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ",P2),P2)"
                        
            If LastRow_StandardizedSF > 2 Then
                wsStandardizedSF.Range("T2:V" & LastRow_StandardizedSF).FillDown
            End If
            
        End If

    ' ============================================================
    '          DONATION SITE VS SALESFORCE: GROSS AMOUNTS
    ' ============================================================
        ' This section isolates transactions where the Donation Site gross amount does not equal
        ' the Salesforce gross amount.
        '
        ' Why this matters:
        ' If the two systems disagree on the gross amount, then the transaction should not
        ' automatically flow into the final import file without an explicit decision.

        ' ---------------------------------------------
        '              UPDATE ROW VARIABLES
        ' ---------------------------------------------
            SectionHeaderRow_UserRequiredAdjustments = LastRow_UserRequiredAdjustments + 6
            HeaderRow_UserRequiredAdjustments = SectionHeaderRow_UserRequiredAdjustments + 1
            DataStartRow_UserRequiredAdjustments = HeaderRow_UserRequiredAdjustments + 1
            
        ' ---------------------------------------------
        '                 SECTION HEADER
        ' ---------------------------------------------
            With wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments)
                .Merge
                .HorizontalAlignment = xlCenter
                .Value = "DONATION SITE VS SALESFORCE: GROSS AMOUNTS"
                .Interior.Color = vbRed
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleSingle
            End With
            
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            wsUserRequiredAdjustments.Range("A" & HeaderRow_UserRequiredAdjustments & ":N" & HeaderRow_UserRequiredAdjustments).Value = _
                Array("Transaction ID", "Disbursement ID", "Transaction Date", "Disbursement Date", "Donation Site - Gross Amount", "SF - Gross Amount", _
                      "Variance", "PMT-ID", "Donation Type", "Site - School Abbreviation", "Adjustment Allowed?", _
                      "Adjusting Journal Name", "CRJ Journal Name", "File Name")

        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IFERROR(FILTER('" & wsConnectionAnalysis.Name & "'!A2:J" & LastRow_ConnectionAnalysis & _
                ",('" & wsConnectionAnalysis.Name & "'!G2:G" & LastRow_ConnectionAnalysis & "<>0)*" & _
                "('" & wsConnectionAnalysis.Name & "'!H2:H" & LastRow_ConnectionAnalysis & "<>""PMT-NOT MATCHED"")),""No Mismatching Amounts"")"
        
            If AllowRevenueAmountAdjustments Then
                wsUserRequiredAdjustments.Range("K" & DataStartRow_UserRequiredAdjustments).Value = "Yes"
            Else
                wsUserRequiredAdjustments.Range("K" & DataStartRow_UserRequiredAdjustments).Value = "No"
            End If
            
            wsUserRequiredAdjustments.Range("L" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""No Mismatching Amounts"",""""," & _
                "IF(K" & DataStartRow_UserRequiredAdjustments & "=""No"",XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & _
                ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!K:K),""CLEARED""))"
            
            wsUserRequiredAdjustments.Range("M" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""No Mismatching Amounts"",""""," & _
                "IF(K" & DataStartRow_UserRequiredAdjustments & "=""No"",XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & _
                ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!J:J),""CLEARED""))"
 
            wsUserRequiredAdjustments.Range("N" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""No Mismatching Amounts"",""""," & _
                "XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!M:M))"
                    
        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            LastRow_UserRequiredAdjustments = wsUserRequiredAdjustments.Cells(wsUserRequiredAdjustments.Rows.Count, 1).End(xlUp).Row
                
        ' ---------------------------------------------
        '                 GROUP SECTION
        ' ---------------------------------------------
            wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments & ":" & (LastRow_UserRequiredAdjustments + 4)).Group
                    
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            If LastRow_UserRequiredAdjustments <> DataStartRow_UserRequiredAdjustments Then
                wsUserRequiredAdjustments.Range("K" & DataStartRow_UserRequiredAdjustments & ":N" & LastRow_UserRequiredAdjustments).FillDown
            Else
                wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments).Interior.Color = vbGreen
                wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments).ShowDetail = False
            End If
        
        ' ---------------------------------------------
        '             SAVE SECTION VARIABLES
        ' ---------------------------------------------
            DataStartRow_UserRequiredAdjustments_GrossAmountVariances = DataStartRow_UserRequiredAdjustments
            LastRow_UserRequiredAdjustments_GrossAmountVariances = LastRow_UserRequiredAdjustments
            
            If JournalType = "Adjusting" Then
                If DataStartRow_UserRequiredAdjustments_GrossAmountVariances <> LastRow_UserRequiredAdjustments_GrossAmountVariances Then
                    Rng_UserRequiredAdjustments_GrossAmountVariances = "'" & wsUserRequiredAdjustments.Name & "'!L" & _
                        DataStartRow_UserRequiredAdjustments_GrossAmountVariances & ":L" & _
                        LastRow_UserRequiredAdjustments_GrossAmountVariances
                Else
                    Rng_UserRequiredAdjustments_GrossAmountVariances = "'" & wsUserRequiredAdjustments.Name & "'!L" & _
                        DataStartRow_UserRequiredAdjustments_GrossAmountVariances
                End If
            Else
                If DataStartRow_UserRequiredAdjustments_GrossAmountVariances <> LastRow_UserRequiredAdjustments_GrossAmountVariances Then
                    Rng_UserRequiredAdjustments_GrossAmountVariances = "'" & wsUserRequiredAdjustments.Name & "'!M" & _
                        DataStartRow_UserRequiredAdjustments_GrossAmountVariances & ":M" & _
                        LastRow_UserRequiredAdjustments_GrossAmountVariances
                Else
                    Rng_UserRequiredAdjustments_GrossAmountVariances = "'" & wsUserRequiredAdjustments.Name & "'!M" & _
                        DataStartRow_UserRequiredAdjustments_GrossAmountVariances
                End If
            End If

    ' ============================================================
    '                 TRANSACTIONS MISSING PMT-IDS
    ' ============================================================
        ' This section isolates transactions that exist in the Donation Site / comparison logic
        ' but could not be matched back to a Salesforce PMT-ID.

        ' ---------------------------------------------
        '              UPDATE ROW VARIABLES
        ' ---------------------------------------------
            SectionHeaderRow_UserRequiredAdjustments = LastRow_UserRequiredAdjustments + 6
            HeaderRow_UserRequiredAdjustments = SectionHeaderRow_UserRequiredAdjustments + 1
            DataStartRow_UserRequiredAdjustments = HeaderRow_UserRequiredAdjustments + 1
            
        ' ---------------------------------------------
        '                 SECTION HEADER
        ' ---------------------------------------------
            With wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments)
                .Merge
                .HorizontalAlignment = xlCenter
                .Value = "TRANSACTIONS MISSING PMT-IDS"
                .Interior.Color = vbRed
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleSingle
            End With
                
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            wsUserRequiredAdjustments.Range("A" & HeaderRow_UserRequiredAdjustments & ":N" & HeaderRow_UserRequiredAdjustments).Value = _
                Array("Transaction ID", "Disbursement ID", "Transaction Date", "Disbursement Date", "Donation Site - Gross Amount", "SF - Gross Amount", _
                      "Variance", "PMT-ID", "Donation Type", "Site - School Abbreviation", "", _
                      "Adjusting Journal Name", "CRJ Journal Name", "File Name")

        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IFERROR(FILTER('" & wsConnectionAnalysis.Name & "'!A2:J" & LastRow_ConnectionAnalysis & ",'" & _
                wsConnectionAnalysis.Name & "'!H2:H" & LastRow_ConnectionAnalysis & "=""PMT-NOT MATCHED""),""All PMT-IDs Found"")"
                   
            wsUserRequiredAdjustments.Range("L" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All PMT-IDs Found"",""""," & _
                "XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!K:K))"
                   
            wsUserRequiredAdjustments.Range("M" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All PMT-IDs Found"",""""," & _
                "XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!J:J))"

            wsUserRequiredAdjustments.Range("N" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All PMT-IDs Found"",""""," & _
                "XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!M:M))"

        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            LastRow_UserRequiredAdjustments = wsUserRequiredAdjustments.Cells(wsUserRequiredAdjustments.Rows.Count, 1).End(xlUp).Row
                
        ' ---------------------------------------------
        '                 GROUP SECTION
        ' ---------------------------------------------
            wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments & ":" & (LastRow_UserRequiredAdjustments + 4)).Group
                
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            If LastRow_UserRequiredAdjustments <> DataStartRow_UserRequiredAdjustments Then
                wsUserRequiredAdjustments.Range("K" & DataStartRow_UserRequiredAdjustments & ":N" & LastRow_UserRequiredAdjustments).FillDown
            Else
                wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments).Interior.Color = vbGreen
                wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments).ShowDetail = False
            End If
        
        ' ---------------------------------------------
        '             SAVE SECTION VARIABLES
        ' ---------------------------------------------
            DataStartRow_UserRequiredAdjustments_MissingPaymentIDs = DataStartRow_UserRequiredAdjustments
            LastRow_UserRequiredAdjustments_MissingPaymentIDs = LastRow_UserRequiredAdjustments
            
            If JournalType = "Adjusting" Then
                If DataStartRow_UserRequiredAdjustments_MissingPaymentIDs <> LastRow_UserRequiredAdjustments_MissingPaymentIDs Then
                    Rng_UserRequiredAdjustments_MissingPaymentIDs = "'" & wsUserRequiredAdjustments.Name & "'!L" & _
                        DataStartRow_UserRequiredAdjustments_MissingPaymentIDs & ":L" & _
                        LastRow_UserRequiredAdjustments_MissingPaymentIDs
                Else
                    Rng_UserRequiredAdjustments_MissingPaymentIDs = "'" & wsUserRequiredAdjustments.Name & "'!L" & _
                        DataStartRow_UserRequiredAdjustments_MissingPaymentIDs
                End If
            Else
                If DataStartRow_UserRequiredAdjustments_MissingPaymentIDs <> LastRow_UserRequiredAdjustments_MissingPaymentIDs Then
                    Rng_UserRequiredAdjustments_MissingPaymentIDs = "'" & wsUserRequiredAdjustments.Name & "'!M" & _
                        DataStartRow_UserRequiredAdjustments_MissingPaymentIDs & ":M" & _
                        LastRow_UserRequiredAdjustments_MissingPaymentIDs
                Else
                    Rng_UserRequiredAdjustments_MissingPaymentIDs = "'" & wsUserRequiredAdjustments.Name & "'!M" & _
                        DataStartRow_UserRequiredAdjustments_MissingPaymentIDs
                End If
            End If

    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        ' AutoFit at the end so the user can immediately review each section without manual resizing.
        wsUserRequiredAdjustments.Columns("A:N").AutoFit
         
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' DIRECT FINAL JOURNAL PATH ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                ROUTE TO THE FINAL JOURNAL PATH
    ' ============================================================
        ' All shared preparation steps are complete.
        ' From this point forward, the macro follows the final import path required for the selected JournalType.
        
        ' If JournalType = "CRJ", jump to the CRJ-specific section. Otherwise, continue into the Adjusting Journal path below.
            If JournalType = "CRJ" Then
                GoTo JournalPath_CRJ
            End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------------------------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''' (ADJUSTING JOURNAL PATH): POPULATE THE ADJUSTING JOURNAL - UNFILTERED WORKSHEET ''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------------------------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "Creating the Intacct Import File"

    ' ============================================================
    '      PURPOSE OF THE ADJUSTING JOURNAL - UNFILTERED WORKSHEET
    ' ============================================================
        ' This worksheet acts as the full staging area for the Adjusting Journal path
        ' before any unresolved exceptions are removed.
        '
        ' Why this matters:
        ' The converter has already built multiple worksheet-specific outputs that represent
        ' different accounting pieces of the final journal entry structure:
        '   - Bank Deposits
        '   - Fees
        '   - Relevant Transactions
        '   - Connection Analysis adjustment rows
        '
        ' Before the macro can decide what should be excluded, it first needs one combined
        ' worksheet that shows the full picture of what the Adjusting Journal would contain
        ' if everything were allowed through.
        '
        ' In other words:
        '   - "Unfiltered" = everything that is currently eligible to be part of the journal
        '   - later "Filtered" = this same data after unresolved user-required items are removed
        '
        ' This worksheet is also where Salesforce reference fields are added back beside the
        ' accounting fields. That makes later review easier and gives the final output a
        ' traceable connection back to the original Salesforce-side transaction.

    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Columns A:AG hold the Intacct import structure for the Adjusting Journal.
            '
            ' Why this matters:
            ' These are the actual accounting-facing fields that would be exported into the
            ' Adjusting Journal import file.
            wsAdjustingUnfiltered.Range("A1:AG1").Value = Array("DONOTIMPORT", "JOURNAL", "DATE", "REVERSEDATE", "DESCRIPTION", "REFERENCE_NO", "LINE_NO", "ACCT_NO", _
                    "LOCATION_ID", "DEPT_ID", "DOCUMENT", "MEMO", "DEBIT", "CREDIT", "SOURCEENTITY", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", _
                    "EXCHANGE_RATE", "STATE", "ALLOCATION_ID", "RASSET", "RDEPRECIATION_SCHEDULE", "RASSET_ADJUSTMENT", "RASSET_CLASS", "RASSETOUTOFSERVICE", _
                    "GLDIMFUNDING_SOURCE", "GLENTRY_PROJECTID", "GLENTRY_CUSTOMERID", "GLENTRY_VENDORID", "GLENTRY_EMPLOYEEID", "GLENTRY_ITEMID", "GLENTRY_CLASSID")
                    
            ' Columns AH:AT hold Salesforce-side reference fields.
            '
            ' Why this matters:
            ' These columns are not part of the Intacct import structure itself, but they make
            ' the journal lines traceable back to Salesforce data. That makes troubleshooting,
            ' auditing, and exception review much easier later.
            wsAdjustingUnfiltered.Range("AH1:AT1").Value = Array("SF_CLOSE_DATE", "SF_DONATION_SITE", "SF_CP_NUMBER", "SF_TRANSACTION_ID", "SF_DISBURSEMENT_ID", _
                    "SF_PAYMENT_METHOD", "SF_CHECK_NUMBER", "SF_PAYMENT_NUMBER", "SF_PRIMARY_CONTACT", "SF_ACCOUNT_NAME", "SF_COMPANY_NAME", "SF_CAMPAIGN_SOURCE", _
                    "SF_DONATION_NAME")
                    
        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' Build the full unfiltered Adjusting Journal body by vertically stacking all
            ' contributing accounting line-item sources into one worksheet.
            '
            ' Sources included:
            '   1. Bank Deposits worksheet
            '   2. Fees worksheet
            '   3. Relevant Transactions worksheet
            '   4. Connection Analysis worksheet rows where a variance exists
            '
            ' Why this matters:
            ' The final Adjusting Journal is not created from a single source worksheet.
            ' It is made from multiple accounting components that each represent a different
            ' part of the journal logic.
            '
            ' Why Connection Analysis is filtered to variance <> 0:
            ' Only rows with a variance need adjustment-entry treatment from that worksheet.
            ' If variance = 0, then there is no adjusting difference to add from that source.
            '
            ' Why SORT is applied:
            ' Sorting keeps the stacked result in a stable order, which makes downstream review,
            ' debugging, and comparison much easier.
            '
            ' The stacked output begins in column B because column A is reserved for the
            ' DONOTIMPORT field used later in the finalized journal worksheet.
            wsAdjustingUnfiltered.Range("B2").Formula2 = "=SORT(" & _
                                                                "VSTACK('" & wsBankDeposits.Name & "'!C2:AH" & LastRow_BankDeposits & _
                                                                       ",'" & wsFees.Name & "'!E2:AJ" & LastRow_Fees & _
                                                                       ",'" & wsRelevantTransactions.Name & "'!AA2:BF" & LastRow_RelevantTransactions & _
                                                                       ",FILTER('" & wsConnectionAnalysis.Name & "'!L2:AQ" & LastRow_ConnectionAnalysis & _
                                                                       ",'" & wsConnectionAnalysis.Name & "'!G2:G" & LastRow_ConnectionAnalysis & "<>0))" & _
                                                                ",4)"

            ' ------------------------------------------------------------
            '   REBUILD SALESFORCE REFERENCE FIELDS FROM THE MEMO COLUMN
            ' ------------------------------------------------------------
            ' The stacked accounting rows now exist, but many of those rows only carry the PMT-ID
            ' inside the Memo field. These helper columns pull the related Salesforce information
            ' back in by extracting the PMT-ID from column L (Memo) and looking it up against
            ' the Initial Data - Intacct worksheet.
            '
            ' Why this matters:
            ' Once multiple accounting sources are stacked together, the accounting fields alone
            ' do not always provide enough business context for review. Reconstructing the
            ' Salesforce reference data makes the unfiltered journal easier to validate and trace.
            '
            ' Why IFERROR / ISBLANK are used:
            ' Not every row will contain a PMT-ID in the expected way. Rather than showing
            ' formula errors, the worksheet returns blank when no usable match is found.

            ' "SF_CLOSE_DATE"
                ' Pull the Salesforce close date tied to the PMT-ID extracted from the Memo field.
                wsAdjustingUnfiltered.Range("AH2").Formula2 = _
                        "=IFERROR(" & _
                                 "IF(ISBLANK(XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!B:B)),""""," & _
                                            "XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!B:B))," & _
                                 """"")"
            
            ' "SF_DONATION_SITE"
                wsAdjustingUnfiltered.Range("AI2").Formula2 = _
                        "=IFERROR(" & _
                                 "IF(ISBLANK(XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!D:D)),""""," & _
                                            "XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!D:D))," & _
                                 """"")"
            
            ' "SF_CP_NUMBER"
                wsAdjustingUnfiltered.Range("AJ2").Formula2 = _
                        "=IFERROR(" & _
                                 "IF(ISBLANK(XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!E:E)),""""," & _
                                            "XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!E:E))," & _
                                 """"")"
            
            ' "SF_TRANSACTION_ID"
                wsAdjustingUnfiltered.Range("AK2").Formula2 = _
                        "=IFERROR(" & _
                                 "IF(ISBLANK(XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!F:F)),""""," & _
                                            "XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!F:F))," & _
                                 """"")"
            
            ' "SF_DISBURSEMENT_ID"
                wsAdjustingUnfiltered.Range("AL2").Formula2 = _
                        "=IFERROR(" & _
                                 "IF(ISBLANK(XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!G:G)),""""," & _
                                            "XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!G:G))," & _
                                 """"")"
            
            ' "SF_PAYMENT_METHOD"
                wsAdjustingUnfiltered.Range("AM2").Formula2 = _
                        "=IFERROR(" & _
                                 "IF(ISBLANK(XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!H:H)),""""," & _
                                            "XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!H:H))," & _
                                 """"")"
            
            ' "SF_CHECK_NUMBER"
                wsAdjustingUnfiltered.Range("AN2").Formula2 = _
                        "=IFERROR(" & _
                                 "IF(ISBLANK(XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!I:I)),""""," & _
                                            "XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!I:I))," & _
                                 """"")"
            
            ' "SF_PAYMENT_NUMBER"
                wsAdjustingUnfiltered.Range("AO2").Formula2 = _
                        "=IFERROR(" & _
                                 "IF(ISBLANK(XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!J:J)),""""," & _
                                            "XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!J:J))," & _
                                 """"")"
            
            ' "SF_PRIMARY_CONTACT"
                wsAdjustingUnfiltered.Range("AP2").Formula2 = _
                        "=IFERROR(" & _
                                 "IF(ISBLANK(XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!K:K)),""""," & _
                                            "XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!K:K))," & _
                                 """"")"
            
            ' "SF_ACCOUNT_NAME"
                wsAdjustingUnfiltered.Range("AQ2").Formula2 = _
                        "=IFERROR(" & _
                                 "IF(ISBLANK(XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!L:L)),""""," & _
                                            "XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!L:L))," & _
                                 """"")"
            
            ' "SF_COMPANY_NAME"
                wsAdjustingUnfiltered.Range("AR2").Formula2 = _
                        "=IFERROR(" & _
                                 "IF(ISBLANK(XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!M:M)),""""," & _
                                            "XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!M:M))," & _
                                 """"")"
            
            ' "SF_CAMPAIGN_SOURCE"
                wsAdjustingUnfiltered.Range("AS2").Formula2 = _
                        "=IFERROR(" & _
                                 "IF(ISBLANK(XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!N:N)),""""," & _
                                            "XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!N:N))," & _
                                 """"")"
            
            ' "SF_DONATION_NAME"
                wsAdjustingUnfiltered.Range("AT2").Formula2 = _
                        "=IFERROR(" & _
                                 "IF(ISBLANK(XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!O:O)),""""," & _
                                            "XLOOKUP(TRIM(MID($L2,SEARCH(""PMT-"",$L2),11)),'Initial Data - Intacct'!$J:$J,'Initial Data - Intacct'!O:O))," & _
                                 """"")"

    ' ============================================================
    '             FIND THE LAST ROW FROM THE WORKSHEET
    ' ============================================================
        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            ' Use column B because the stacked journal body begins in column B.
            '
            ' Why this matters:
            ' Column A is reserved for DONOTIMPORT and is intentionally not populated here.
            ' Column B is the first populated column of the unfiltered stacked result, so it is
            ' the most reliable column for determining how far the worksheet extends.
            LastRow_AdjustingUnfiltered = wsAdjustingUnfiltered.Cells(wsAdjustingUnfiltered.Rows.Count, 2).End(xlUp).Row
            
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            ' Only the Salesforce reference/helper columns need to be filled down.
            '
            ' Why this matters:
            ' The stacked journal body in columns B:AG is already produced by a spill formula.
            ' The helper columns AH:AT are row-by-row lookups that depend on each row’s memo/PMT-ID,
            ' so those formulas must be filled down to match the full unfiltered row count.
            If LastRow_AdjustingUnfiltered > 2 Then
                wsAdjustingUnfiltered.Range("AH2:AT" & LastRow_AdjustingUnfiltered).FillDown
            End If

    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        ' Add filters and autofit at the end so the worksheet is immediately usable for review,
        ' troubleshooting, and comparison before the filtering phase begins.
        wsAdjustingUnfiltered.Range("A1:AT1").AutoFilter
        wsAdjustingUnfiltered.Columns("A:AT").AutoFit
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------------------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''' (ADJUSTING JOURNAL PATH): POPULATE THE ADJUSTING JOURNAL - FILTERED WORKSHEET '''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------------------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "Filtering Out Any Missing Data from the Intacct Import File"

    ' ============================================================
    '       PURPOSE OF THE ADJUSTING JOURNAL - FILTERED WORKSHEET
    ' ============================================================
        ' The "Adjusting Journal - Unfiltered" worksheet contains the full set of journal lines
        ' that would flow into the Adjusting Journal path if no exceptions still existed.
        '
        ' This worksheet creates the next-stage version of that data by removing any rows that are
        ' still tied to unresolved issues from the User-Required Adjustments worksheet.
        '
        ' Why this matters:
        ' The User-Required Adjustments worksheet is where the user reviews exceptions such as:
        '   - missing bank allocation / school assignments
        '   - missing school names
        '   - missing account / division / funding source determinations
        '   - gross amount variances
        '   - missing PMT-IDs
        '
        ' For each exception section, the macro stored the related journal names in range variables.
        ' If a row is still unresolved, its journal name remains active in those saved ranges.
        ' If the issue is resolved, the journal name changes to "CLEARED".
        '
        ' This worksheet uses those saved ranges to identify which journal descriptions should still
        ' be blocked from the import path.
        '
        ' In other words:
        '   - Unfiltered worksheet = everything currently staged
        '   - Filtered worksheet   = staged data minus unresolved exception rows
        '
        ' This is important because the final import worksheet should only build from rows that are
        ' currently safe to allow forward.

    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Keep the same structure as the Unfiltered worksheet.
            '
            ' Why this matters:
            ' The Filtered worksheet is intended to be the same dataset structure as the Unfiltered
            ' worksheet, just with unresolved exception rows removed. Keeping the same columns allows
            ' the final journal worksheet to reference a consistent layout.
            wsAdjustingFiltered.Range("A1:AG1").Value = Array("DONOTIMPORT", "JOURNAL", "DATE", "REVERSEDATE", "DESCRIPTION", "REFERENCE_NO", "LINE_NO", "ACCT_NO", _
                    "LOCATION_ID", "DEPT_ID", "DOCUMENT", "MEMO", "DEBIT", "CREDIT", "SOURCEENTITY", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", _
                    "EXCHANGE_RATE", "STATE", "ALLOCATION_ID", "RASSET", "RDEPRECIATION_SCHEDULE", "RASSET_ADJUSTMENT", "RASSET_CLASS", "RASSETOUTOFSERVICE", _
                    "GLDIMFUNDING_SOURCE", "GLENTRY_PROJECTID", "GLENTRY_CUSTOMERID", "GLENTRY_VENDORID", "GLENTRY_EMPLOYEEID", "GLENTRY_ITEMID", "GLENTRY_CLASSID")

            wsAdjustingFiltered.Range("AH1:AT1").Value = Array("SF_CLOSE_DATE", "SF_DONATION_SITE", "SF_CP_NUMBER", "SF_TRANSACTION_ID", "SF_DISBURSEMENT_ID", _
                    "SF_PAYMENT_METHOD", "SF_CHECK_NUMBER", "SF_PAYMENT_NUMBER", "SF_PRIMARY_CONTACT", "SF_ACCOUNT_NAME", "SF_COMPANY_NAME", "SF_CAMPAIGN_SOURCE", _
                    "SF_DONATION_NAME")
            
        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' Build one combined list of all unresolved journal names saved from the
            ' User-Required Adjustments worksheet.
            '
            ' The saved ranges come from the exception sections:
            '   - Bank Allocations Not Found
            '   - Transactions Missing School Name
            '   - Adjustments to Account|Division|Funding Source
            '   - Donation Site vs Salesforce Gross Amounts
            '   - Transactions Missing PMT-IDs
            '
            ' Why VSTACK is used:
            ' Each exception section saved its journal-name range separately. VSTACK combines those
            ' separate saved ranges into one vertical array so they can be treated as a single
            ' exclusion list.
            '
            ' Why UNIQUE is used:
            ' The same journal name may appear multiple times across one or more sections. UNIQUE
            ' reduces the stacked list down to distinct journal names so the exclusion logic stays
            ' cleaner and avoids redundant comparisons.
            '
            ' Why FILTER is applied to that stacked list:
            ' The saved ranges may contain:
            '   - "CLEARED" for resolved items
            '   - blank cells
            '
            ' Those should not block anything from the import path, so they are removed from the
            ' exclusion list before the MATCH test is performed.
            '
            ' Why MATCH is used against column E of the Unfiltered worksheet:
            ' Column E in the Unfiltered worksheet is the DESCRIPTION field. Earlier exception
            ' sections saved the relevant journal descriptions so that unresolved rows could later
            ' be identified and excluded here.
            '
            ' Final filtering logic:
            '   - If the Unfiltered DESCRIPTION exists in the unresolved journal-name list,
            '     that row is excluded.
            '   - If the DESCRIPTION does not exist in the unresolved list,
            '     that row is allowed into the Filtered worksheet.
            '
            ' The formula begins in column B because column A is still reserved for DONOTIMPORT,
            ' matching the Unfiltered worksheet structure.
            wsAdjustingFiltered.Range("B2").Formula2 = _
                    "=LET(" & _
                        "UserRequiredAdjustments," & _
                            "UNIQUE(VSTACK(" & _
                                Rng_UserRequiredAdjustments_BankAllocations & "," & _
                                Rng_UserRequiredAdjustments_MissingSchoolNames & "," & _
                                Rng_UserRequiredAdjustments_AccountDivisionFundingAdjustments & "," & _
                                Rng_UserRequiredAdjustments_GrossAmountVariances & "," & _
                                Rng_UserRequiredAdjustments_MissingPaymentIDs & ")), " & _
                        "UserRequiredAdjustmentsFiltered," & _
                            "FILTER(UserRequiredAdjustments,(UserRequiredAdjustments<>""CLEARED"")*(UserRequiredAdjustments<>"""")), " & _
                        "FILTER('" & wsAdjustingUnfiltered.Name & "'!B2:AT" & LastRow_AdjustingUnfiltered & "," & _
                            "NOT(ISNUMBER(MATCH('" & wsAdjustingUnfiltered.Name & "'!E2:E" & LastRow_AdjustingUnfiltered & "," & _
                            "UserRequiredAdjustmentsFiltered,0))))" & _
                    ")"

    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        ' Add filters and autofit once the filtered output is created so the worksheet is ready
        ' for review.
        '
        ' Why this matters:
        ' This worksheet is the bridge between the raw staged journal and the final journal layout.
        ' Making it readable helps when validating whether unresolved exception rows were excluded
        ' correctly.
        wsAdjustingFiltered.Range("A1:AT1").AutoFilter
        wsAdjustingFiltered.Columns("A:AT").AutoFit
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''' (ADJUSTING JOURNAL PATH): POPULATE THE ADJUSTING JOURNAL - FINALIZED WORKSHEET '''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "Finalizing Intacct Import File"

    ' ============================================================
    '      PURPOSE OF THE ADJUSTING JOURNAL - FINALIZED WORKSHEET
    ' ============================================================
        ' The Filtered worksheet already removed any rows tied to unresolved exceptions.
        ' This final worksheet takes that filtered dataset and reshapes it into the final
        ' import-ready layout for Intacct.
        '
        ' Why this matters:
        ' The Filtered worksheet is still acting as a staging area. This worksheet becomes the
        ' actual final journal file structure the user is expected to review and import.
        '
        ' This step is kept separate from the Filtered worksheet because:
        '   - it creates the final line-number behavior
        '   - it keeps the final import output isolated from the working filter logic
        '   - it provides a clean, import-focused worksheet for the end user
        '
        ' In short:
        '   - Adjusting Journal - Unfiltered = everything staged
        '   - Adjusting Journal - Filtered   = unresolved exception rows removed
        '   - Adjusting Journal Import       = final import-ready output

    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Keep the same field structure as the upstream journal worksheets so the final
            ' import tab remains consistent with the staging tabs.
            wsAdjustingJournal.Range("A1:AG1").Value = Array("DONOTIMPORT", "JOURNAL", "DATE", "REVERSEDATE", "DESCRIPTION", "REFERENCE_NO", "LINE_NO", "ACCT_NO", _
                    "LOCATION_ID", "DEPT_ID", "DOCUMENT", "MEMO", "DEBIT", "CREDIT", "SOURCEENTITY", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", _
                    "EXCHANGE_RATE", "STATE", "ALLOCATION_ID", "RASSET", "RDEPRECIATION_SCHEDULE", "RASSET_ADJUSTMENT", "RASSET_CLASS", "RASSETOUTOFSERVICE", _
                    "GLDIMFUNDING_SOURCE", "GLENTRY_PROJECTID", "GLENTRY_CUSTOMERID", "GLENTRY_VENDORID", "GLENTRY_EMPLOYEEID", "GLENTRY_ITEMID", "GLENTRY_CLASSID")

            wsAdjustingJournal.Range("AH1:AT1").Value = Array("SF_CLOSE_DATE", "SF_DONATION_SITE", "SF_CP_NUMBER", "SF_TRANSACTION_ID", "SF_DISBURSEMENT_ID", _
                    "SF_PAYMENT_METHOD", "SF_CHECK_NUMBER", "SF_PAYMENT_NUMBER", "SF_PRIMARY_CONTACT", "SF_ACCOUNT_NAME", "SF_COMPANY_NAME", "SF_CAMPAIGN_SOURCE", _
                    "SF_DONATION_NAME")
                    
        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' Most fields in this worksheet are pulled directly from the Filtered worksheet.
            '
            ' Why this matters:
            ' The Filtered worksheet already represents the currently approved set of rows.
            ' This final worksheet should not rebuild the filtering logic again. Instead, it
            ' should simply carry forward the approved rows into the final layout.
            '
            ' The repeated IF / ISBLANK pattern is used so the final worksheet remains visually
            ' clean and does not show unnecessary zeros or placeholder outputs after the real data ends.

        ' ---------------------------------------------
        '                   COLUMN A
        ' ---------------------------------------------
            ' "DONOTIMPORT"
                ' This column intentionally remains blank.
                '
                ' Why this matters:
                ' It preserves the expected import-file structure and gives the user a place
                ' to manually flag rows later if needed.

        ' ---------------------------------------------
        '                 COLUMNS B:F
        ' ---------------------------------------------
            ' These fields are copied from the Filtered worksheet only when the row is active.
            '
            ' Why the extra IF logic exists:
            ' The finalized worksheet is filled down farther than the true filtered row count so it
            ' can dynamically respond if upstream corrections change the spilled filtered results.
            ' Because of that, each formula must protect against displaying unwanted values below
            ' the active data range.

            ' "JOURNAL"
                wsAdjustingJournal.Range("B2").Formula2 = _
                        "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!B2,""""))),""""," & _
                            "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!B2,"""")))"
            
            ' "DATE"
                wsAdjustingJournal.Range("C2").Formula2 = _
                        "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!C2,""""))),""""," & _
                            "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!C2,"""")))"
                        
            ' "REVERSEDATE"
                wsAdjustingJournal.Range("D2").Formula2 = _
                        "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!D2,""""))),""""," & _
                            "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!D2,"""")))"
            
            ' "DESCRIPTION"
                wsAdjustingJournal.Range("E2").Formula2 = _
                        "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!E2,""""))),""""," & _
                            "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!E2,"""")))"
                            
            ' "REFERENCE_NO"
                wsAdjustingJournal.Range("F2").Formula2 = _
                        "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!F2,""""))),""""," & _
                            "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!F2,"""")))"
                            
        ' ---------------------------------------------
        '                   COLUMN G
        ' ---------------------------------------------
            ' "LINE_NO"
                ' This field is recalculated here instead of being copied straight across.
                '
                ' Why this matters:
                ' The final worksheet must preserve proper journal line sequencing after filtering.
                ' If rows were removed upstream, the original line numbering may no longer be valid
                ' for the final import file.
                '
                ' Logic:
                '   - If there is no active row, return blank.
                '   - If the incoming row is the first line of a journal entry set, keep its line number.
                '   - Otherwise increment from the prior finalized row.
                wsAdjustingJournal.Range("G2").Formula2 = _
                        "=IF('" & wsAdjustingFiltered.Name & "'!$B2="""",""""," & _
                            "IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!G2," & _
                                "1+G1))"

        ' ---------------------------------------------
        '                COLUMNS H:AT
        ' ---------------------------------------------
            ' These columns are passed through from the Filtered worksheet.
            '
            ' Why this matters:
            ' By the time rows reach the finalized worksheet, the row-level values themselves
            ' should already be correct. The main remaining task here is to suppress blank spillover
            ' and preserve a clean final import layout.

            ' "ACCT_NO"
                wsAdjustingJournal.Range("H2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!H2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!H2))"
            
            ' "LOCATION_ID"
                wsAdjustingJournal.Range("I2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!I2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!I2))"
            
            ' "DEPT_ID"
                wsAdjustingJournal.Range("J2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!J2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!J2))"
            
            ' "DOCUMENT"
                wsAdjustingJournal.Range("K2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!K2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!K2))"
            
            ' "MEMO"
                wsAdjustingJournal.Range("L2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!L2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!L2))"
            
            ' "DEBIT"
                wsAdjustingJournal.Range("M2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!M2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!M2))"
            
            ' "CREDIT"
                wsAdjustingJournal.Range("N2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!N2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!N2))"
            
            ' "SOURCEENTITY"
                wsAdjustingJournal.Range("O2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!O2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!O2))"
            
            ' "CURRENCY"
                wsAdjustingJournal.Range("P2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!P2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!P2))"
            
            ' "EXCH_RATE_DATE"
                wsAdjustingJournal.Range("Q2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!Q2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!Q2))"
            
            ' "EXCH_RATE_TYPE_ID"
                wsAdjustingJournal.Range("R2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!R2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!R2))"
            
            ' "EXCHANGE_RATE"
                wsAdjustingJournal.Range("S2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!S2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!S2))"
            
            ' "STATE"
                wsAdjustingJournal.Range("T2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!T2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!T2))"
            
            ' "ALLOCATION_ID"
                wsAdjustingJournal.Range("U2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!U2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!U2))"
            
            ' "RASSET"
                wsAdjustingJournal.Range("V2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!V2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!V2))"
            
            ' "RDEPRECIATION_SCHEDULE"
                wsAdjustingJournal.Range("W2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!W2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!W2))"
            
            ' "RASSET_ADJUSTMENT"
                wsAdjustingJournal.Range("X2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!X2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!X2))"
            
            ' "RASSET_CLASS"
                wsAdjustingJournal.Range("Y2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!Y2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!Y2))"
            
            ' "RASSETOUTOFSERVICE"
                wsAdjustingJournal.Range("Z2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!Z2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!Z2))"
            
            ' "GLDIMFUNDING_SOURCE"
                wsAdjustingJournal.Range("AA2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AA2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AA2))"
            
            ' "GLENTRY_PROJECTID"
                wsAdjustingJournal.Range("AB2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AB2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AB2))"
            
            ' "GLENTRY_CUSTOMERID"
                wsAdjustingJournal.Range("AC2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AC2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AC2))"
            
            ' "GLENTRY_VENDORID"
                wsAdjustingJournal.Range("AD2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AD2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AD2))"
            
            ' "GLENTRY_EMPLOYEEID"
                wsAdjustingJournal.Range("AE2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AE2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AE2))"
            
            ' "GLENTRY_ITEMID"
                wsAdjustingJournal.Range("AF2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AF2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AF2))"
            
            ' "GLENTRY_CLASSID"
                wsAdjustingJournal.Range("AG2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AG2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AG2))"
            
            ' "SF_CLOSE_DATE"
                wsAdjustingJournal.Range("AH2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AH2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AH2))"
            
            ' "SF_DONATION_SITE"
                wsAdjustingJournal.Range("AI2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AI2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AI2))"
            
            ' "SF_CP_NUMBER"
                wsAdjustingJournal.Range("AJ2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AJ2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AJ2))"
            
            ' "SF_TRANSACTION_ID"
                wsAdjustingJournal.Range("AK2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AK2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AK2))"
            
            ' "SF_DISBURSEMENT_ID"
                wsAdjustingJournal.Range("AL2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AL2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AL2))"
            
            ' "SF_PAYMENT_METHOD"
                wsAdjustingJournal.Range("AM2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AM2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AM2))"
            
            ' "SF_CHECK_NUMBER"
                wsAdjustingJournal.Range("AN2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AN2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AN2))"
            
            ' "SF_PAYMENT_NUMBER"
                wsAdjustingJournal.Range("AO2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AO2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AO2))"
            
            ' "SF_PRIMARY_CONTACT"
                wsAdjustingJournal.Range("AP2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AP2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AP2))"
            
            ' "SF_ACCOUNT_NAME"
                wsAdjustingJournal.Range("AQ2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AQ2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AQ2))"
            
            ' "SF_COMPANY_NAME"
                wsAdjustingJournal.Range("AR2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AR2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AR2))"
            
            ' "SF_CAMPAIGN_SOURCE"
                wsAdjustingJournal.Range("AS2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AS2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AS2))"
            
            ' "SF_DONATION_NAME"
                wsAdjustingJournal.Range("AT2").Formula2 = _
                    "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AT2)),""""," & _
                        "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",'" & wsAdjustingFiltered.Name & "'!AT2))"

    ' ---------------------------------------------
    '               FILL FORMULAS DOWN
    ' ---------------------------------------------
        ' Fill down based on the Unfiltered worksheet row count.
        '
        ' Why this matters:
        ' The Filtered worksheet is dynamic and may grow or shrink depending on which exceptions
        ' are currently resolved. Using the Unfiltered row count gives the finalized worksheet
        ' enough space to respond as the filtered results change.
            If LastRow_AdjustingUnfiltered > 2 Then
                wsAdjustingJournal.Range("B2:AT" & LastRow_AdjustingUnfiltered).FillDown
            End If

    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        ' Apply filters and autofit so the final import worksheet is immediately ready for review.
        wsAdjustingJournal.Range("A1:AT1").AutoFilter
        wsAdjustingJournal.Columns("A:AT").AutoFit

    ' ============================================================
    '   DIRECT THE CONVERTER PAST THE "JOURNAL PATH:CRJ" SECTIONS
    ' ============================================================
        ' The Adjusting Journal path is complete at this point.
        '
        ' Why this matters:
        ' The macro contains a separate CRJ build path below. Once the Adjusting path is fully
        ' created, the converter should skip the CRJ-only sections and move directly to the
        ' shared post-processing logic.
        GoTo MoveFiles
        
        
JournalPath_CRJ:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' JOURNAL PATH: CRJ ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Columns AH:AT
' To be determined later:
WorksheetName = "Intacct"
    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
    
    
    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
        
        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
    ' ============================================================
    ' WORKSHEET: CRJ JOURNAL - UNFILTERED
    ' ============================================================
' To be determined later:
' CRJ Route -- Unfiltered

wsCRJUnfiltered.Name = "CRJ Unfiltered"

    ' ============================================================
    ' WORKSHEET: CRJ JOURNAL - FILTERED
    ' ============================================================

' To be determined later:
' CRJ Route -- Filtered
wsCRJFiltered.Name = "CRJ Filtered"

    ' ============================================================
    ' WORKSHEET: CRJ JOURNAL - FINALIZED
    ' ============================================================

' To be determined later:
' CRJ - Finalized
wsCRJ.Name = "CRJ Import"





MoveFiles:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' MOVE MISSING SOURCE FILES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '            PURPOSE OF THIS SECTION
    ' ============================================================
        ' Any source files tied to unresolved exceptions should be moved out of the main processing
        ' folder and into a separate "Process Later" folder.
        '
        ' Why this matters:
        ' Once the macro has identified transactions that still require user review, those source
        ' files should not remain mixed in with files that were fully processed successfully.
        '
        ' Moving those files serves two purposes:
        '   1. It keeps the original source folder cleaner and easier to manage.
        '   2. It separates files tied to unresolved exceptions so they can be reviewed later.
        '
        ' At a minimum, files tied to missing PMT-IDs should be moved.
        '
        ' In addition, if revenue amount adjustments are NOT allowed, then files tied to
        ' gross amount mismatches should also be moved, because those rows are still unresolved
        ' and should not be treated as complete.

    ' ============================================================
    '         BUILD A UNIQUE LIST OF FILES THAT MUST BE MOVED
    ' ============================================================
        ' Use a dictionary so each file name is stored only once.
        '
        ' Why this matters:
        ' A single source file may appear multiple times in the User-Required Adjustments worksheet.
        ' If we tried to move the same file more than once, the macro would fail after the first move.
        ' The dictionary prevents duplicates and gives us one clean list of file names to process.
        Set dictFilesToMove = CreateObject("Scripting.Dictionary")

    ' ============================================================
    '      ADD FILES FROM THE "TRANSACTIONS MISSING PMT-IDS" SECTION
    ' ============================================================
        ' These files should always be moved because missing PMT-ID records are unresolved and
        ' should be reviewed outside the normal completed-processing flow.
        For UserRequiredAdjustmentsRow = DataStartRow_UserRequiredAdjustments_MissingPaymentIDs To LastRow_UserRequiredAdjustments_MissingPaymentIDs
            
            ' Pull the file name from column N of the User-Required Adjustments worksheet.
            AdditionalFileToMove = Trim(CStr(wsUserRequiredAdjustments.Range("N" & UserRequiredAdjustmentsRow).Value))
            
            ' Ignore blanks so we only store real file names.
            If AdditionalFileToMove <> "" Then
                
                ' Add the file name only if it is not already in the dictionary.
                If Not dictFilesToMove.Exists(AdditionalFileToMove) Then
                    dictFilesToMove.Add AdditionalFileToMove, AdditionalFileToMove
                End If
                
            End If
            
        Next UserRequiredAdjustmentsRow

    ' ============================================================
    '   ALSO ADD FILES FROM GROSS AMOUNT VARIANCES IF NOT ALLOWED
    ' ============================================================
        ' If revenue amount adjustments are not allowed, then gross amount mismatches remain
        ' unresolved exceptions and their source files should also be moved to the Process Later folder.
        '
        ' Why this matters:
        ' A mismatch that is not allowed cannot be treated as complete processing. Those files
        ' should stay grouped with the other unresolved exception files for later follow-up.
        If AllowRevenueAmountAdjustments = False Then
            
            For UserRequiredAdjustmentsRow = DataStartRow_UserRequiredAdjustments_GrossAmountVariances To LastRow_UserRequiredAdjustments_GrossAmountVariances
                
                ' Pull the file name from column N of the User-Required Adjustments worksheet.
                AdditionalFileToMove = Trim(CStr(wsUserRequiredAdjustments.Range("N" & UserRequiredAdjustmentsRow).Value))
                
                ' Ignore blanks so only real file names are added.
                If AdditionalFileToMove <> "" Then
                    
                    ' Add the file name only if it is not already in the dictionary.
                    If Not dictFilesToMove.Exists(AdditionalFileToMove) Then
                        dictFilesToMove.Add AdditionalFileToMove, AdditionalFileToMove
                    End If
                    
                End If
                
            Next UserRequiredAdjustmentsRow
            
        End If

    ' ============================================================
    '      CONVERT THE UNIQUE FILE LIST INTO A VBA ARRAY
    ' ============================================================
        ' Convert the dictionary keys into an array so the file-move loop can iterate through
        ' the final unique file list in a simple, controlled way.
        '
        ' Why this matters:
        ' The dictionary is excellent for uniqueness, while the array is convenient for a clean
        ' sequential move process.
        If dictFilesToMove.Count > 0 Then
            ReDim FilesToMove(1 To dictFilesToMove.Count)
            
            FileIndex = 1
            For Each UniqueFileName In dictFilesToMove.Keys
                FilesToMove(FileIndex) = CStr(UniqueFileName)
                FileIndex = FileIndex + 1
            Next UniqueFileName
        Else
            ReDim FilesToMove(1 To 1)
            FilesToMove(1) = ""
        End If

    ' ============================================================
    '             CREATE THE "PROCESS LATER" FOLDER
    ' ============================================================
        ' Create a timestamped folder inside the original Donation Site folder.
        '
        ' Why this matters:
        ' Using a timestamp keeps each run distinct and avoids collisions with prior folders.
        ' It also makes it easier to trace when the files were set aside for later review.
        FolderPath_ProcessLater = FolderPath_DonationSite & "\Process Later - " & Format(Now, "yyyy.mm.dd_hh.mm.ss")
        MkDir FolderPath_ProcessLater

    ' ============================================================
    '              MOVE THE FILES INTO THE NEW FOLDER
    ' ============================================================
        ' Move each unique file from the original Donation Site folder into the new
        ' Process Later folder.
        '
        ' Why this matters:
        ' At this point, the macro has already determined which files are tied to unresolved
        ' exceptions. Moving them now separates incomplete work from completed work.
        If dictFilesToMove.Count > 0 Then
            For FileIndex = LBound(FilesToMove) To UBound(FilesToMove)
                
                SourceFilePath = FolderPath_DonationSite & "\" & FilesToMove(FileIndex)
                DestinationFilePath = FolderPath_ProcessLater & "\" & FilesToMove(FileIndex)
                
                ' Only move the file if it still exists in the source folder.
                '
                ' Why this matters:
                ' This prevents the macro from crashing if the file was already moved, renamed,
                ' deleted, or otherwise no longer exists in the expected location.
                If Dir(SourceFilePath) <> "" Then
                    Name SourceFilePath As DestinationFilePath
                End If
                
            Next FileIndex
        End If

    ' ============================================================
    '                FINAL COMPLETION MESSAGE
    ' ============================================================
        ExtraMessage = "Macro Completed Successfully"
        ExtraMessage_Title = "Macro Completed Successfully"

        GoTo CompleteMacro

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CONSOLIDATION ONLY MODE: TRUE ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ConsolidationOnly:
    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "CONSOLIDATION MODE INITIATED"


    GoTo CompleteMacro
    


CreateButton_Step2:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CREATE BUTTON FOR DONATION SITE REPORTS '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "Creating a worksheet for when the user is ready to import Donation Site Data"

    ' ============================================================
    '          CHECK WHETHER THE BUTTON WORKSHEET EXISTS
    ' ============================================================
        ' This worksheet is used as a holding / instruction page when the required
        ' Donation Site reports have not yet been added to the workbook.
        '
        ' Why this matters:
        ' The converter depends on the Donation Site report data to continue.
        ' If that data is missing, the macro needs a clear stopping point that tells
        ' the user what to do next instead of allowing the process to continue into
        ' incomplete or broken logic.
        '
        ' Before creating the worksheet, first check whether it already exists so the
        ' macro does not create duplicates.
        wsFound = False
        
        For Each ws In wbMacro.Worksheets
            If ws.Name = "No Donation Site Report" Then
                wsFound = True
                Exit For
            End If
        Next ws

    ' ============================================================
    '      CREATE THE BUTTON WORKSHEET IF IT DOES NOT EXIST
    ' ============================================================
        If wsFound = False Then
            
            ' ----------------------------------------------------
            '               CREATE THE WORKSHEET
            ' ----------------------------------------------------
                ' Place this worksheet after "COMPLETE RESET" so it is easy for the
                ' user to find in the workbook flow.
                Set wsButton = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets("COMPLETE RESET"))
                
            ' ----------------------------------------------------
            '               RENAME THE WORKSHEET
            ' ----------------------------------------------------
                ' The worksheet name clearly tells the user why the converter stopped.
                wsButton.Name = "No Donation Site Report"
                
            ' ----------------------------------------------------
            '               FORMAT THE WORKSHEET
            ' ----------------------------------------------------
                ' The dark background helps this page stand out visually from the other
                ' data worksheets and makes it feel more like an instruction screen.
                wsButton.Cells.Interior.Color = vbBlack
                
            ' ----------------------------------------------------
            '                  CREATE THE BUTTON
            ' ----------------------------------------------------
                ' Create a large button so it is obvious what the user should click
                ' once the Donation Site reports have been added.
                '
                ' Why this matters:
                ' This gives the user a simple restart point. Instead of needing to
                ' remember which macro to run, they can just click the button on this page.
                Set DonationSiteButton = wsButton.Buttons.Add(150, 50, 825, 275)
                
                With DonationSiteButton
                    .Caption = "Click here to add the '" & DonationSite & "' Reports"
                    .OnAction = ConverterName
                    .Font.Size = 50
                    .Font.Bold = True
                    .Font.Color = RGB(200, 200, 0)
                End With
                
        End If

    ' ============================================================
    '            HIDE THE INITIAL DATA WORKSHEET
    ' ============================================================
        ' Hide the Initial Data worksheet so the user's attention stays on the
        ' instruction/button worksheet rather than the partially prepared data.
        '
        ' Why this matters:
        ' At this stage, the converter is intentionally pausing for missing input.
        ' Hiding the Initial Data worksheet helps reduce confusion and makes it more
        ' obvious that the next required action is to add the Donation Site reports.
        wsInitialData.Visible = xlSheetHidden

CompleteMacro:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' RESTORE EXCEL ENVIRONMENT ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '            UPDATE THE STATUS BAR AND PROGRESS BAR
    ' ============================================================
        Application.StatusBar = "Completing Converter Process"
    
    ' ============================================================
    '            RESTORE THE EXCEL APPLICATION SETTINGS
    ' ============================================================
        ' Throughout the macro, certain Excel settings may have been adjusted to improve speed,
        ' reduce screen flickering, and prevent unnecessary prompts while the converter runs.
        '
        ' It is important to restore those settings at the end so Excel returns to a normal,
        ' user-friendly state after the macro is finished.
        '
        ' Why this matters:
        '   - DisplayAlerts must be turned back on so Excel shows normal warning / prompt messages.
        '   - ScreenUpdating must be turned back on so the workbook becomes visually responsive again.
        '   - Calculation must be returned to Automatic so formulas calculate normally after the macro.
        '   - StatusBar should be cleared so Excel regains control of its own status display.
        '   - EnableEvents should be turned back on if the macro had disabled them earlier.
        
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
        'Application.EnableEvents = True

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' PROVIDE MESSAGE TO USER ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' ============================================================
    '              DISPLAY THE FINAL USER MESSAGE
    ' ============================================================
        ' Show the completion message only after the Excel environment has been restored.
        '
        ' Why this matters:
        ' If the message is shown before Excel is reset, the user could dismiss the message while
        ' the application is still left in a restricted or non-standard state.
        '
        ' By restoring Excel first, the macro ensures the workbook is ready for normal use as soon
        ' as the user clicks OK.
        MsgBox ExtraMessage, _
               vbOKOnly, _
               ExtraMessage_Title

End Sub
