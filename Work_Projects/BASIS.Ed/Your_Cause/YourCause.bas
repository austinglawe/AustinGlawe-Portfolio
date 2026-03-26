Sub YourCause_AR_Converter()
' ============================================================
' MODULE: YourCause_AR
' AUTHOR: Austin Glawe
' CREATED: 2025.09.30
' LAST UPDATED: 2026.03.26
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
            ' Merge Salesforce/Intacct data with Donation Site Report data to generate an Intacct import file, including all supporting documentation.
                ' Consolidate all donation site reports into a single worksheet while preserving the original report data for supporting documentation.
            ' Connect transactions to their corresponding deposits to support accurate bank reconciliation and deposit imports into Intacct.
            ' Reconcile records between Salesforce, the Donation Site, and Intacct to identify missing, incomplete, or incorrectly entered transactions.
            
    ' ============================================================
    '                         REQUIREMENTS
    ' ============================================================
        ' 1. One of the following reports:
            ' Salesforce Report (SF_Sync_Report)
                ' Found at: https://basised.lightning.force.com/lightning/r/Report/00ORj000006hKo1MAE/view
            ' Intacct Report (SFREV_Undeposited Funds Report)
                ' Found by going to: Intacct >> Platform Services >> Custom Report >> SFREV_Undeposited Funds Report >> Run
                
        ' 2. A folder containing all donation site reports to process
            ' The folder name must contain "Your Cause"
          
    ' ============================================================
    '                             FLOW
    ' ============================================================
        '  1. User selects Salesforce or Intacct report
        '  2. User selects folder containing donation site reports
        '  3. Donation site reports are consolidated and supporting worksheets are created
        '  4. Data is merged with the selected Salesforce or Intacct report
        '  5. Transactions are connected to deposits
        '  6. Reconciliation checks are performed across Salesforce Data, Donation Site Data, and Intacct Data
        '  7. Transactions requiring user review are filtered to a designated worksheet
        '  8. Intacct import file is generated
        '  9. User reviews and resolves any flagged transactions
        ' 10. Intacct import file is updated automatically after flags are resolved
        ' 11. Import file is uploaded to Intacct
        ' 12. Documentation (this file) is attached to all deposits produced by this converter
    
    ' ============================================================
    '             UPDATE LOG (LAST UPDATED: 2026.03.26)
    ' ============================================================
        ' Original Production Rollout Date: 2025.09.30

        ' Updates:
            ' 2026.03.26 - Initiated the update log.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CONFIGURATIONS AND VARIABLES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section provides a spot for all configuration and variable declarations.
    
    ' ============================================================
    '                        CONFIGURATIONS
    ' ============================================================
        ' ---------------------------------------------
        '             DECLARE CONFIGURATIONS
        ' ---------------------------------------------
            Dim wbMacro As Workbook
            
            Dim ConverterName As String
            Dim DonationSite As String
            Dim DonationSite_Salesforce As String
            Dim JournalName As String
            
            Dim AllowConsolidationOnly As Boolean
            Dim IncludeOriginalReports As Boolean
            Dim AllowRevenueAmountAdjustments As Boolean
            
            Dim RowsToDeleteFromBottomOfDonationSiteReport As Long
            
            Dim AssignedHeaderRow_InitialReport As Long
            Dim AllowHeaderRowSearch_InitialReport As Boolean
            
            Dim AssignedHeaderRow_DonationSiteReports As Long
            Dim AllowHeaderRowSearch_DonationSiteReports As Boolean
            
            Dim JournalType As String
            Dim AllowJournalTypeManualOverride As Boolean
            
            Dim ColumnHeaders_Initial_Intacct As Variant
            Dim ColumnHeaders_Initial_Salesforce As Variant
            Dim ColumnHeaders_YourCause As Variant

        ' ---------------------------------------------
        '             ASSIGN CONFIGURATIONS
        ' ---------------------------------------------
            ' Store a reference to this workbook so it is clearly distinguished from any temporary workbooks opened during the converter process.
                Set wbMacro = ThisWorkbook
            
            ' Store the converter procedure name so it can be assigned to a button if the user needs to return later and continue the process.
                ConverterName = "YourCause.YourCause_AR_Converter"
            
            ' Store the Donation Site name used throughout this converter.
                DonationSite = "Your Cause"
            
            ' This is the Salesforce-side constant tied to this Donation Site.
              ' See Module 'A_Global_Constants'.
                DonationSite_Salesforce = DonationSiteYourCause

            ' Store the journal name used for the Adjusting import route.
              ' By default this should be "SFREV". ("CHAR" was used prior to 2025.10)
                JournalName = "SFREV"

            ' This switch allows the converter to stop after consolidating / preparing the Donation Site reports without building the final Intacct import file.
              ' By default this should be False.
                AllowConsolidationOnly = False
            
            ' This switch determines whether the original Donation Site reports are preserved in the workbook for supporting documentation and review.
              ' By default this should be True.
                IncludeOriginalReports = True
            
            ' This switch controls whether mismatching revenue amounts between Salesforce and the Donation Site are allowed to continue through the converter as valid adjustment entries.
              ' By default this should be True, unless a stricter workflow is required.
                AllowRevenueAmountAdjustments = True
            
            ' This setting controls how many non-data rows should be deleted from the bottom of each Donation Site report.
              ' By default this should be 0 because no additional lines are found below the Your Cause data within the reports.
                RowsToDeleteFromBottomOfDonationSiteReport = 0

            ' ..............................
            '         INITIAL REPORT
            '       HEADER ROW SETTINGS
            ' ..............................
                ' This switch allows the converter to search all rows for the Initial Report headers instead of assuming a fixed row.
                 ' By default this should be True, because Intacct Reports can vary in their header rows depending on the file type.
                  ' Salesforce: Row 1
                  ' Intacct: CSV downloads = Row 1
                  ' Intacct: Excel Exports = Row 5
                    AllowHeaderRowSearch_InitialReport = True
                
                ' This setting allows a fixed Initial Report header row to be used when known.
                  ' By default this should be 0.
                    AssignedHeaderRow_InitialReport = 0
                
                ' Logic override:
                  ' If no valid assigned row exists and searching was turned off, force searching back on so the converter can still function.
                    If AssignedHeaderRow_InitialReport < 1 And AllowHeaderRowSearch_InitialReport = False Then
                        AllowHeaderRowSearch_InitialReport = True
                    End If
            
            ' ..............................
            '      DONATION SITE REPORT
            '       HEADER ROW SETTINGS
            ' ..............................
                ' This switch allows the converter to search each Donation Site report for its headers instead of using a fixed header row.
                  ' By default this should be False to allow quicker searches (when the header row always appears in the same row).
                    AllowHeaderRowSearch_DonationSiteReports = False
            
                ' This setting defines the expected header row for the Donation Site reports.
                  ' By Default this should be 1, because YourCause headers are currently expected on row 1 every time.
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
                  ' By default this should be False, to allow the route to be determine based on the Initial Report Type (Intacct or Salesforce), _
                    unless testing or special handling is needed.
                    AllowJournalTypeManualOverride = True
                
                ' Valid values when manual override is enabled:
                  ' "Adjusting"
                  ' "CRJ"
                ' By default this should be "".
                    JournalType = ""
                
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
                  ' Intacct Custom Report Name: "SFREV_Undeposited Funds Report"
                  ' Go To: Intacct >> Platform Services >> Custom Report >> SFREV_Undeposited Funds Report >> Run
                    ColumnHeaders_Initial_Intacct = Array("Journal Entry Modified Date", "Close Date", "Batch Posting Date", "SF Donation Site", "C&P Number", _
                            "SF Transaction ID", "SF Disbursement ID", "SF Payment Method", "SF Check Number", "SF Payment Number", "SF Primary Contact", _
                            "SF Account Name", "SF Company Name", "SF Campaign Source", "SF Opportunity Name", "Memo", "Location Name", "Location ID", "Account Number", _
                            "Division ID", "Funding Source", "Debt Service Series ID", "Journal", "Journal Number", "Journal Description", "Record Number", _
                            "Credit Amount", "Debit Amount", "Amount")
                    
                ' Salesforce Report Column Headers (A:T) - 20 columns
                  ' Salesforce Report Name: "SF_Sync_Report"
                  ' Go To: https://basised.lightning.force.com/lightning/r/Report/00ORj000006hKo1MAE/view
                    ColumnHeaders_Initial_Salesforce = Array("Payment: Created Date", "Close Date", "Deposit Date", "Donation Site", "C&P Order Number", _
                            "Check/Reference Number", "Disbursement ID", "Payment Type", "Check Number", "Payment: Payment Number", "Primary Contact", "Account Name", _
                            "Company Name", "Primary Campaign Source", "Opportunity Name", "C&P Account Name", "C&P Account Name Correction", _
                            "Payment Amount", "Campaign Type", "Description")
                    
            ' ..............................
            '        DONATION SITE REPORT
            '         COLUMN HEADERS
            ' ..............................
                ' YourCause Column Headers (A:AO) - 41 columns
                ' If this array changes, also update the section that sets up the wsConsolidatedData worksheet environment.
                    ' Search for "ESTABLISH_CONSOLIDATED_WORKSHEET_COLUMN_HEADERS" within this converter.
                ' Go To: BASIS SharePoint >> CHARTER AR >> Charter AR Shared >> 1. Donation Sites Reports >> Your Cause >> Ready for AR Team
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
        ' ---------------------------------------------
        '       BUTTON AND SHAPE OBJECT VARIABLES
        ' ---------------------------------------------
            ' These are used for
            Dim DonationSiteButton As Button
        
        ' ---------------------------------------------
        '     DONATION SITE FILE AND REPORT TRACKING
        ' ---------------------------------------------
            Dim FileCount_DonationSite As Long
            Dim FileName_DonationSite As String
            Dim FileNamesList_DonationSite() As String
            
            Dim FileNumber_DonationSite As Long
            
            
            Dim WorksheetName As String
            Dim wsFound As Boolean

        ' ---------------------------------------------
        '      DONATION SITE TEMP WORKSHEET METRICS
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
        '            FILE AND FOLDER DIALOGS
        ' ---------------------------------------------
            Dim fdFilePath_InitialReport As FileDialog
            Dim fdFolderPath_DonationSite As FileDialog

        ' ---------------------------------------------
        '             FILE AND FOLDER PATHS
        ' ---------------------------------------------
            Dim FilePath_InitialReport As String
            Dim FolderPath_DonationSite As String
            Dim FolderPath_ProcessLater As String
            
            Dim SourceFilePath As String
            Dim DestinationFilePath As String
        
        ' ---------------------------------------------
        '     FILE-MOVE AND PROCESS-LATER VARIABLES
        ' ---------------------------------------------
            ' These variables are used near the end of the converter to build a
            ' unique list of source files that should be moved into a
            ' "Process Later" folder.
            Dim dictFilesToMove As Object
            Dim FilesToMove() As String
            
            Dim FileName_ToMove As String
            Dim UniqueFileName_FromDictionary As Variant
            Dim FileMoveIndex As Long
            Dim UserRequiredAdjustmentsRow As Long
            Dim AdditionalFileToMove As String
            Dim FileIndex As Long
            Dim UniqueFileName As Variant
        
        ' ---------------------------------------------
        '       GENERAL CONTROL AND USER RESPONSE
        ' ---------------------------------------------
            Dim UserResponse As VbMsgBoxResult

            Dim ExtraMessage As String
            Dim ExtraMessage_Title As String

        ' ---------------------------------------------
        '       WORKSHEET AND WORKBOOK REFERENCES
        ' ---------------------------------------------
            Dim ws As Worksheet
            Dim wsCheckForInitial As Worksheet
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
            
            ' ..............................
            '     Path: Adjusting Journal
            '           Worksheets
            ' ..............................
                Dim wsAdjustingUnfiltered As Worksheet
                Dim wsAdjustingFiltered As Worksheet
                Dim wsAdjustingJournal As Worksheet
            
            ' ..............................
            '        Path: CRJ Journal
            '           Worksheets
            ' ..............................
                Dim wsCRJUnfiltered As Worksheet
                Dim wsCRJFiltered As Worksheet
                Dim wsCRJ As Worksheet

        

        ' ---------------------------------------------
        '         PATH ROUTES
        ' ---------------------------------------------
            Dim InitialExists As Boolean
            Dim InitialPath As String '....................... "Salesforce" or "Intacct"
            
            
            Dim HeaderRow_InitialReport As Long
            Dim ColumnCheck_InitialReport As Long
            Dim LastRow_InitialData As Long

        

        ' ---------------------------------------------
        '          LAST ROWS
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
Dim TempLastRow_InitialReport As Long
        

        ' ---------------------------------------------
        '    USER-REQUIRED ADJUSTMENTS RANGE STRINGS
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
        '     USER-REQUIRED ADJUSTMENTS SECTION ROWS
        ' ---------------------------------------------
            Dim OK_UserRequiredAdjustments As Long
            Dim PaymentIDs_Missing As Boolean
            
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
' FILE COUNT VARIABLES
' ---------------------------------------------
    Dim FileCount_DonationSite_WrongFileType As Long
    Dim FileCount_DonationSite_WrongReport As Long
    Dim FileCount_DonationSite_Unusable As Long
    Dim FileCount_DonationSite_Used As Long
    
    Dim FilesCount_Total_Message As String
    Dim FileCount_DonationSite_Used_Message As String
    Dim FileCount_DonationSite_WrongFileType_Message As String
    Dim FileCount_DonationSite_WrongReport_Message As String
    Dim FileCount_DonationSite_Unusable_Message As String
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CONFIGURE EXCEL ENVIRONMENT ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section provides a dedicated spot to temporarily disable Excel interface features to improve performance and prevent prompts while the converter runs.
        
    ' ============================================================
    '           DISABLE RELEVANT EXCEL INTERFACE FEATURES
    ' ============================================================
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        'Application.Calculation = xlCalculationManual
  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' DETERMINE CONVERTER STARTING POINT ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section is to establish the starting point for the converter.
            ' If the user has already imported an Intacct or Salesforce report, it allows them to skip to the IMPORT DONATION SITE REPORTS section.
            ' Otherwise this section will take the user through the process to import the Initial (Intacct or Salesforce) Report.
            
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "CHECKING FOR EXISTING INITIAL REPORT WORKSHEET"

    ' ============================================================
    '            CHECK FOR AN EXISTING INITIAL WORKSHEET
    ' ============================================================
            ' Set InitialExists to False before checking starting point.
                InitialExists = False
            
            ' Loop through each worksheet to determine whether an Initial Data worksheet already exists in the converter workbook.
            ' If found:
              ' 1. Set InitialExists = True
              ' 2. Identify whether the InitialPath is "Intacct" or "Salesforce"
              ' 3. Assign wsInitialData to the matching worksheet
              ' 4. Assign JournalType only if it has not already been manually set
            ' If the Initial Data worksheet is not found, these variables will be assigned values later.
              
                For Each wsCheckForInitialForInitial In wbMacro.Worksheets
                ' Make the worksheet visible so any existing hidden sheets can be accessed later.
                    wsCheckForInitialForInitial.Visible = xlSheetVisible
            
                    If wsCheckForInitialForInitial.Name = "Initial Data - Intacct" Then
                        InitialExists = True
                        InitialPath = "Intacct"
            
                        If JournalType = "" Then
                            JournalType = "Adjusting"
                        End If
            
                        Set wsInitialData = wsCheckForInitialForInitial
                        Exit For
            
                    ElseIf wsCheckForInitialForInitial.Name = "Initial Data - SF" Then
                        InitialExists = True
                        InitialPath = "Salesforce"
            
                        If JournalType = "" Then
                            JournalType = "CRJ"
                        End If
            
                        Set wsInitialData = wsCheckForInitialForInitial
                        Exit For
                    End If
                Next wsCheckForInitialForInitial
        
    ' ============================================================
    '              INITIAL WORKSHEET ANALYSIS RESULTS
    ' ============================================================
        ' ---------------------------------------------
        '     INITIAL DATA WORKSHEET ALREADY EXISTS
        ' ---------------------------------------------
            ' If the Initial Data worksheet already exists, skip the Initial Report import process and continue directly to IMPORT DONATION SITE REPORTS section.
                If InitialExists Then
                    GoTo Add_ConsolidatedReports
                End If
        
        ' ---------------------------------------------
        '   INITIAL DATA WORKSHEET DOES NOT YET EXIST
        ' ---------------------------------------------
            ' If the Initial Data worksheet does not already exist, reset InitialPath so it can be assigned later from the user-selected report.
                InitialPath = ""
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' PRE-RUN CHECKLIST AND CONFIRMATION OF USER READINESS '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section is to confirm user readiness to begin the converter.
        ' It provides a comprehensive message for the user to understand the requirements, before starting the converter.
    
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "PROVIDE PRE-RUN CHECKLIST AND CONFIRM USER READINESS"
       
    ' ============================================================
    '     PROVIDE PRE-RUN CHECKLIST AND CONFIRM USER READINESS
    ' ============================================================
        ' ---------------------------------------------
        '           PROVIDE PRE-RUN CHECKLIST
        ' ---------------------------------------------
            ' Display a pre-run checklist outlining all required information the user must have available before starting the converter.
                UserResponse = MsgBox( _
                        "Before starting, please confirm you have the following:" & vbCrLf & _
                            "    1. A report downloaded from either Intacct or Salesforce." & vbCrLf & _
                            "    2. All donation site reports downloaded and placed in a folder" & vbCrLf & "        with '" & DonationSite & "' in the folder's name." & vbCrLf & vbCrLf & _
                        "Are you ready to continue?", _
                        vbYesNo + vbQuestion, _
                        DonationSite & " - AR Converter Confirmation")
        
        ' ---------------------------------------------
        '             CONFIRM USER READINESS
        ' ---------------------------------------------
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
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' After user confirms readiness, this section establishes a clean workbook environment by deleting all worksheets and creating a reset page for the user, _
          by using a pre-established procedure.
        
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "RESETTING WORKBOOK ENVIRONMENT"
        
    ' ============================================================
    '              CREATE A CLEAN WORKBOOK ENVIRONMENT
    ' ============================================================
        ' Clear the workbook using the Reset.Create_Reset_Worksheet procedure to prepare a clean workbook environment for the converter.
            Reset.Create_Reset_Worksheet

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' RECONFIGURE EXCEL ENVIRONMENT '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' After the workbook environment is cleared and cleaned, this section re-establishes the temporarily disabled Excel interface features in case they were turned _
          back on during the RESET WORKBOOK ENVIRONMENT section.
        
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "RE-CONFIGURING EXCEL ENVIRONMENT"
    
    ' ============================================================
    '           DISABLE RELEVANT EXCEL INTERFACE FEATURES
    ' ============================================================
        ' Temporarily disable Excel interface features to improve performance and prevent prompts while the converter runs.
            Application.DisplayAlerts = False
            Application.ScreenUpdating = False
            'Application.Calculation = xlCalculationManual

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CONSOLIDATION ONLY MODE: CHECK ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' If the AllowConsolidationOnly switch is set to True, this section allows the user to jump past the Initial Report import, because the AllowConsolidationOnly switch _
          is intended to allow the user to consolidate donation site reports, without creating the actual Intacct import file. _
          This feature allows the user to easily analyze the Donation Site Data only.
    
    ' ============================================================
    '            DETERMINE CONSOLIDATE ONLY MODE SETTING
    ' ============================================================
        ' If consolidation-only mode is enabled, skip the remainder of the converter setup and proceed directly to the IMPORT DONATION SITE REPORTS section.
            If AllowConsolidationOnly Then
                Application.StatusBar = "CONSOLIDATION MODE INITIATED"
                
                GoTo Add_ConsolidatedReports
            End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' IMPORT INITIAL REPORT AND DETERMINE REPORT TYPE ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section requires user input to get the file path of the Initial Data Report via a file dialog file selector.
        ' This section also checks to make sure the report given by the user matches the column headers of the valid Salesforce or Intacct Reports allowed within this converter.
        ' If a valid report is given, it establishes which report type the user provided and sets the path for where the converter goes next.
        
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        ' Update the Status Bar
            Application.StatusBar = "REQUESTING INITIAL (INTACCT/SALESFORCE) REPORT FROM USER"
            
    ' ============================================================
    '                      USER FILE SELECTION
    ' ============================================================
        ' Open a file picker so the user can select the Initial Report.
            Set fdFilePath_InitialReport = Application.FileDialog(msoFileDialogFilePicker)
        
            With fdFilePath_InitialReport
                .Title = "Select the Initial (Intacct or Salesforce) Report"
                .AllowMultiSelect = False
                
                .Filters.Clear
                .Filters.Add "Excel Files", "*.xlsx; *.xls; *.csv"
            
            ' If the user cancels the file selection dialog, prepare an exit message and stop the converter.
                If .Show <> -1 Then
                    ExtraMessage = "No file selected." & vbCrLf & _
                                    "Please locate the Intacct or Salesforce Report and try again."
                    
                    ExtraMessage_Title = "No File Selected"
                    
                    GoTo CompleteMacro
                End If
            
                FilePath_InitialReport = .SelectedItems(1)
            End With
            
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        ' Update the Status Bar
            Application.StatusBar = "VALIDATING THE INITIAL REPORT"
            
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
        ' Determine whether the selected report is an allowed Intacct or Salesforce report by comparing column headers against _
          the expected header arrays from the CONFIGURATIONS AND VARIABLES section.
            ' If the AllowHeaderRowSearch_InitialReport swtich is enabled, each row is tested until a match is found.
            ' Otherwise, the AssignedHeaderRow_InitialReport is used.
        
        ' ---------------------------------------------
        '    FIND THE LAST ROW OF THE INITIAL REPORT
        ' ---------------------------------------------
            TempLastRow_InitialReport = wsTemp_InitialReport.Cells(wsTemp_InitialReport.Rows.Count, 1).End(xlUp).Row
    
        ' ---------------------------------------------
        '       INITIAL HEADER SEARCH: ON VS. OFF
        ' ---------------------------------------------
          ' INITIAL HEADER SEARCH: ON
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
                    ' If the row does not match Intacct headers, check if it matches the Salesforce headers.
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
                
                ' If a valid header row was found, stop searching. Otherwise search the next row.
                    If InitialPath <> "" Then
                        Exit For
                    End If
                    
                Next HeaderRow_InitialReport
            
          ' INITIAL HEADER SEARCH: OFF
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
                    ' If the row does not match Intacct headers, check if it matches the Salesforce headers.
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
        ' Confirm the established InitialPath and determine the path the converter goes next.
            ' If no valid InitialPath was identified, provide the user with a message and stop the converter.
            ' Otherwise proceed to the appropriate path.
            
            ' ..............................
            '         INVALID REPORT
            ' ..............................
                If InitialPath = "" Then
                    ExtraMessage = "The report does not appear to be a valid Intacct or Salesforce report." & vbCrLf & _
                                  "If this is an error, please reach out to " & CurrentVBACodeMaintainer & " to further assist in this process."
                    
                    ExtraMessage_Title = "Invalid Initial Report"
                    
                    GoTo CompleteMacro
                    
            ' ..............................
            '           SALESFORCE
            ' ..............................
                ElseIf InitialPath = "Salesforce" Then
                    GoTo InitialPath_SF
                    
            ' ..............................
            '             INTACCT
            ' ..............................
                ElseIf InitialPath = "Intacct" Then
                    ' Continue through to the PROCESS INITIAL REPORT ROUTE section
                End If
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' PROCESS INITIAL REPORT ROUTE '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section copies the Initial Report into the Converter workbook, to be worked with later.
        ' If the JournalType is not pre-determined in the CONFIGURATIONS AND VARIABLES section, this section will help determine the JournalType route based on _
          which Initial Report the user provided.
            ' By default: Intacct = "Adjusting", Salesforce = "CRJ"
        
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "COPYING THE INITIAL REPORT TO THE CONVERTER WORKBOOK"
    
    ' ============================================================
    '                    INITIAL DATA - INTACCT
    ' ============================================================
        ' Create a worksheet to hold the Initial Intacct Report Data to be used later in the converter.
            Set wsInitialData = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets(wbMacro.Worksheets.Count))
            wsInitialData.Name = "Initial Data - Intacct"
            
        ' Copy over the Initial Intacct Report into the wsInitialData Worksheet.
            wsTemp_InitialReport.Range("A" & HeaderRow_InitialReport & ":AC" & TempLastRow_InitialReport).Copy wsInitialData.Range("A1")
        
        ' Format the wsInitialData worksheet.
            wsInitialData.Range("A1:AC1").AutoFilter
            wsInitialData.Range("A:AC").WrapText = False
            wsInitialData.Columns("A:AC").AutoFit
            
        ' Close the wbTemp_InitialReport workbook without saving changes.
            wbTemp_InitialReport.Close SaveChanges:=False
        
        ' If the JournalType is not already assigned, assign it to the "Adjusting" path.
            If JournalType = "" Then
                JournalType = "Adjusting"
            End If
        
        ' Jump over the INITIAL DATA - SALESFORCE section into the IMPORT DONATION SITE REPORTS section.
            GoTo Add_ConsolidatedReports

    ' ============================================================
    '                   INITIAL DATA - SALESFORCE
    ' ============================================================
InitialPath_SF:
        ' Create a worksheet to hold the Initial Salesforce Report Data to be used later in the converter.
            Set wsInitialData = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets(wbMacro.Worksheets.Count))
            wsInitialData.Name = "Initial Data - SF"
            
        ' Copy over the Initial Salesforce Report into the wsInitialData Worksheet.
            wsTemp_InitialReport.Range("A" & HeaderRow_InitialReport & ":T" & TempLastRow_InitialReport).Copy wsInitialData.Range("A1")
        
        ' Format the wsInitialData worksheet.
            wsInitialData.Range("A1:T1").AutoFilter
            wsInitialData.Range("A:T").WrapText = False
            wsInitialData.Columns("A:T").AutoFit
            
        ' Close the wbTemp_InitialReport workbook without saving changes.
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
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section requires user input to get the folder path with the Donation Site Reports via a file dialog folder selector.
        ' It validates the user provides a valid folder path name (including "YourCause" or "Your Cause").
        ' It validates there are files in the folder.
        ' It validates the files in the folder are excel type files
        ' It validates the files in the folder are Your Cause report files by comparing column headers to the established ColumnHeaders_YourCause from _
          the CONFIGURATIONS AND VARIABLES section.
            ' If the AllowHeaderRowSearch_DonationSiteReports switch is enabled, it allows the converter to search each row of each file for there column headers.
            ' Otherwise it relies on the AssignedHeaderRow_DonationSiteReports provided in the CONFIGURATIONS AND VARIABLES section.
        ' If all validations check out, the data from the Donation Site Report is moved to a consolidated worksheet with all other valid Your Cause Report files' data.
        ' If the IncludeOriginalReports switch is enabled, the original Donation Site Reports are copied over to provide additional documentation and support.
        ' This section provides comprehensive counts of categories of the files from the user-provided folder path that are used later.
        ' If the user cancels operations, any validations fail, or no usable files are provided within the folder, this section will allow the converter to jump to _
          the CREATE BUTTON FOR DONATION SITE REPORTS section, to create a worksheet with a button that allows the user to jump back to this section when they are ready.
        
        ' If the AllowConsolidationOnly switch is enabled, this section allows the user to jump out of the converter and bypass any other parts of the converter.
    
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "REQUESTING THE DONATION SITE REPORTS FOLDER PATH FROM USER"
        
    
    ' ============================================================
    '                     USER FOLDER SELECTION
    ' ============================================================
        ' ---------------------------------------------
        '              USER SELECTS FOLDER
        ' ---------------------------------------------
            ' Open a folder picker so the user can select the folder containing the Donation Site Reports to process.
                Set fdFolderPath_DonationSite = Application.FileDialog(msoFileDialogFolderPicker)
                                
                With fdFolderPath_DonationSite
                    .Title = "Select the '" & DonationSite & "' Reports Folder"
                    .AllowMultiSelect = False
                                    
                ' If the user cancels the folder selection dialog, prepare an exit message and jump to the CREATE BUTTON FOR DONATION SITE REPORTS section.
                    If .Show <> -1 Then
                        ExtraMessage = "No Folder Selected. Please locate the correct folder and try again."
                        
                        ExtraMessage_Title = "No Folder Selected"
                        
                        GoTo CreateButton_Step2
                    End If
                             
                ' Store the selected folder path for use later in the converter.
                    FolderPath_DonationSite = .SelectedItems(1)
                End With
        
        ' ---------------------------------------------
        '             UPDATE THE STATUS BAR
        ' ---------------------------------------------
            ' Update the status bar
                Application.StatusBar = "VERIFYING FOLDER PATH FROM USER"
    
        ' ---------------------------------------------
        '       VALIDATE FOLDER NAMING CONVENTION
        ' ---------------------------------------------
            ' This validation helps ensure the user intentionally selects a folder created specifically for the converter to process the Donation Site Reports.
              ' Verify the selected folder path contains "Your Cause" or "YourCause" in the folder name.
                If (InStr(1, FolderPath_DonationSite, "YourCause", vbTextCompare) = 0) And (InStr(1, FolderPath_DonationSite, "Your Cause", vbTextCompare) = 0) Then
                    ExtraMessage = "The selected folder path does not contain '" & DonationSite & "' in the folder name. " & _
                                   "Please rename the folder or locate the correct folder and try again." & vbCrLf & vbCrLf & _
                                   "If this error persists, please contact " & CurrentVBACodeMaintainer & " to further assist in the process."
                    
                    ExtraMessage_Title = "Missing dedicated folder naming convention"
                    
                    GoTo CreateButton_Step2
                    
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
                    
                    GoTo CreateButton_Step2
                End If
                
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "CREATING THE CONSOLIDATED DATA WORKSHEET"
        
    ' ============================================================
    '      SET UP THE wsConsolidatedData WORKSHEET ENVIRONMENT
    ' ============================================================
        ' ---------------------------------------------
        '    CREATE THE wsConsolidatedData WORKSHEET
        ' ---------------------------------------------
            ' Create a new worksheet to hold all consolidated Donation Site Report data.
                Set wsConsolidatedData = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets(wbMacro.Worksheets.Count))
                wsConsolidatedData.Name = "Consolidated Reports"

        ' ---------------------------------------------
        '              ADD COLUMN HEADERS
        ' ---------------------------------------------
            ' If updates are needed, this is the "ESTABLISH_CONSOLIDATED_WORKSHEET_COLUMN_HEADERS" section (update column ranges)
            
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
                '             UPDATE THE STATUS BAR
                ' ---------------------------------------------
                    ' Update the Status Bar to display the current file number, total file count, and file name being processed.
                        Application.StatusBar = "PROCESSING FILE " & _
                                                (FileNumber_DonationSite - LBound(FileNamesList_DonationSite) + 1) & _
                                                " OF " & _
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
                        ' If it is not, add to the FileCount_DonationSite_WrongFileType count and move to the next file.
                            If Not LCase$(FileNamesList_DonationSite(FileNumber_DonationSite)) Like "*.csv" _
                              And Not LCase$(FileNamesList_DonationSite(FileNumber_DonationSite)) Like "*.xls*" Then
                              
                                FileCount_DonationSite_WrongFileType = FileCount_DonationSite_WrongFileType + 1
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
                            
                        ' HEADER ROW SEARCH ENABLED
                            If AllowHeaderRowSearch_DonationSiteReports Then
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
                                    
                        ' HEADER ROW SEARCH DISABLED
                            Else
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
                        ' If no valid Donation Site Report headers are found, add to the FileCount_DonationSite_WrongReport count and move to the next file.
                            FileCount_DonationSite_WrongReport = FileCount_DonationSite_WrongReport + 1
                            GoTo DoNotUseFile
                            
                ' ---------------------------------------------
                '    PROCESS FILES WITH VALID COLUMN HEADERS
                ' ---------------------------------------------
UseFile:
                    ' ..............................
                    '  DETERMINE NUMBER OF DATA ROWS
                    ' ..............................
                        ' Determine the number of data rows.
                            DataRows_DonationSite = LastRow_TempDonationSite - HeaderRow_DonationSite
    
                        ' Determine the total number of usable data rows after accounting for RowsToDeleteFromBottomOfDonationSiteReport.
                            DataRows_DonationSite_Total = DataRows_DonationSite - RowsToDeleteFromBottomOfDonationSiteReport
                        
                        ' Ensure the report contains usable data after accounting for RowsToDeleteFromBottomOfDonationSiteReport.
                          ' If not, add to the FileCount_DonationSite_Unusable count and move to the next file.
                            If DataRows_DonationSite_Total < 1 Then
                                FileCount_DonationSite_Unusable = FileCount_DonationSite_Unusable + 1
                                GoTo DoNotUseFile
                            End If
                    
                        ' Use HeaderRow_DonationSite to determine the start row of the data.
                            DataStartRow_DonationSite = HeaderRow_DonationSite + 1
                            
                        ' Determine the adjusted last row after accounting for RowsToDeleteFromBottomOfDonationSiteReport.
                            LastRow_TempDonationSite_Adjusted = LastRow_TempDonationSite - RowsToDeleteFromBottomOfDonationSiteReport
    
                ' ---------------------------------------------
                '      COPY DONATION SITE REPORT DATA INTO
                '        THE CONSOLIDATED DATA WORKSHEET
                ' ---------------------------------------------
                    ' Find the next available row in wsConsolidatedData using Column A.
                        LastRow_ConsolidatedData = wsConsolidatedData.Cells(wsConsolidatedData.Rows.Count, "A").End(xlUp).Row + 1
                
                    ' Copy the wsTemp_DonationSite data into wsConsolidatedData.
                        wsTemp_DonationSite.Range("A" & DataStartRow_DonationSite & ":AO" & LastRow_TempDonationSite_Adjusted).Copy _
                                Destination:=wsConsolidatedData.Range("A" & LastRow_ConsolidatedData)
                
                    ' Build the worksheet name to be used for documentation/reference.
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
                '          INTO THE CONVERTER WORKBOOK
                ' ---------------------------------------------
                    ' Format the wsTemp_DonationSite worksheet.
                        wsTemp_DonationSite.Columns("A:AO").AutoFit
                        
                    ' If the IncludeOriginalReports switch is on, copy the original Donation Site Report worksheet into Converter workbook for documentation/reference.
                        If IncludeOriginalReports Then
    
                            ' Copy the original Donation Site worksheet into the Converter workbook.
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
                    ' Close the temporary workbook without saving changes and proceed to the next file in the folder path.
                        On Error Resume Next
                        If Not wbTemp_DonationSite Is Nothing Then
                            wbTemp_DonationSite.Close SaveChanges:=False
                        End If
                        On Error GoTo 0
    
            Next FileNumber_DonationSite

    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "VALIDATING FILES USED"

    ' ============================================================
    '         CREATE THE CORRECT FORMS OF THE USER MESSAGE
    ' ============================================================
      ' Create a structurally and grammatically correct format, to display to the user, about the files provided.
        ' ---------------------------------------------
        '            FilesCount_Total_Message
        ' ---------------------------------------------
            If UBound(FileNamesList_DonationSite) <> 1 Then
               FilesCount_Total_Message = "From the " & UBound(FileNamesList_DonationSite) & " files found in the selected folder:"
            Else
                FilesCount_Total_Message = "From the " & UBound(FileNamesList_DonationSite) & " file found in the selected folder:"
            End If
            
        ' ---------------------------------------------
        '      FileCount_DonationSite_Used_Message
        ' ---------------------------------------------
            If FileCount_DonationSite_Used <> 1 Then
                FileCount_DonationSite_Used_Message = Right("       " & FileCount_DonationSite_Used, 7) & " were usable files."
            Else
                FileCount_DonationSite_Used_Message = Right("       " & FileCount_DonationSite_Used, 7) & " was a usable file."
            End If
            
        ' ---------------------------------------------
        '        FileCount_DonationSite_WrongFileType_Message
        ' ---------------------------------------------
            If FileCount_DonationSite_WrongFileType <> 1 Then
                FileCount_DonationSite_WrongFileType_Message = Right("       " & FileCount_DonationSite_WrongFileType, 7) & " were the wrong file type."
            Else
                FileCount_DonationSite_WrongFileType_Message = Right("       " & FileCount_DonationSite_WrongFileType, 7) & " was the wrong file type."
            End If
            
        ' ---------------------------------------------
        '         FileCount_DonationSite_WrongReport_Message
        ' ---------------------------------------------
            If FileCount_DonationSite_WrongReport <> 1 Then
                FileCount_DonationSite_WrongReport_Message = Right("       " & FileCount_DonationSite_WrongReport, 7) & " were the correct file types, but not '" & DonationSite & "' Reports."
            Else
                FileCount_DonationSite_WrongReport_Message = Right("       " & FileCount_DonationSite_WrongReport, 7) & " was the correct file type but, not a '" & DonationSite & "' Report."
            End If
            
        ' ---------------------------------------------
        '    FileCount_DonationSite_Unusable_Message
        ' ---------------------------------------------
            If FileCount_DonationSite_Unusable <> 1 Then
                FileCount_DonationSite_Unusable_Message = Right("       " & FileCount_DonationSite_Unusable, 7) & " were the correct file types, were '" & DonationSite & "' Reports," & _
                                                          " but had no usable" & vbCrLf & _
                                                          "         " & "data."
            Else
                FileCount_DonationSite_Unusable_Message = Right("       " & FileCount_DonationSite_Unusable, 7) & " was the correct file type, was a '" & DonationSite & "' Report," & _
                                                          " but had no usable" & vbCrLf & _
                                                          "         " & "data."
            End If
      
    ' ============================================================
    '                      VALIDATE FILE COUNT
    ' ============================================================
        ' If no files were imported into wsConsolidatedData, provide a message to the user and jump to the CREATE BUTTON FOR DONATION SITE REPORTS section.
            If FileCount_DonationSite_Used = 0 Then

                    ExtraMessage = "The selected folder did not contain any usable '" & DonationSite & "' files. " & vbCrLf & _
                                    "Please find the correct folder and try again." & vbCrLf & vbCrLf & _
                                    FilesCount_Total_Message & vbCrLf & _
                                    FileCount_DonationSite_Used_Message & vbCrLf & _
                                    FileCount_DonationSite_WrongFileType_Message & vbCrLf & _
                                    FileCount_DonationSite_WrongReport_Message & vbCrLf & _
                                    FileCount_DonationSite_Unusable_Message & vbCrLf & _
                                    "If this is a mistake, please reach out to " & CurrentVBACodeMaintainer & " to further assist."
    
                    ExtraMessage_Title = "No Usable Files Found"
                
            ' Delete the Consolidated Data worksheet, since no usable files were found.
                wsConsolidatedData.Delete
                
            ' Create a button for the user to jump to the IMPORT DONATION SITE REPORTS section for when they are ready to try again.
                GoTo CreateButton_Step2
            End If
    
    ' ============================================================
    '             DELETE THE BUTTON WORKSHEET (IF USED)
    ' ============================================================
        ' If at least one usable file exists, check if the "No Donation Site Report" worksheet exists, delete it, and proceed.
          ' This is the worksheet created if the CreateButton_Step2 process has been used.
            For Each wsButton In wbMacro.Worksheets
                If wsButton.Name = "No Donation Site Report" Then
                    wsButton.Delete
                End If
            Next wsButton
    
    ' ============================================================
    '            FORMAT THE CONSOLIDATED DATA WORKSHEET
    ' ============================================================
        ' Apply AutoFilter and AutoFit to the wsConsolidatedData worksheet.
            wsConsolidatedData.Range("A1:AQ1").AutoFilter
            wsConsolidatedData.Columns("A:AQ").AutoFit
            
    ' ============================================================
    '        IF CONSOLIDATION MODE: ON -> CONVERTER COMPLETE
    ' ============================================================
        ' If the AllowConsolidationOnly switch is on, break out of the converter.
            If AllowConsolidationOnly Then
                Application.StatusBar = "CONSOLIDATION MODE: COMPLETED"
                
                ' ---------------------------------------------
                '            CREATE THE USER MESSAGE
                ' ---------------------------------------------
                    ExtraMessage = "Thank you for your patience! The consolidation process completed successfully." & vbCrLf & vbCrLf & _
                                    FilesCount_Total_Message & vbCrLf & _
                                    FileCount_DonationSite_Used_Message & vbCrLf & _
                                    FileCount_DonationSite_WrongFileType_Message & vbCrLf & _
                                    FileCount_DonationSite_WrongReport_Message & vbCrLf & _
                                    FileCount_DonationSite_Unusable_Message

                    ExtraMessage_Title = "Consolidation Mode: Successful"
            
            ' Jump to end of the converter.
                GoTo CompleteMacro
            End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''----------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CREATE ALL ADDITIONAL WORKSHEETS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''----------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section is to set up the rest of the workbook environment by creating the rest of the worksheets needed to create the Intacct Import File.
        
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "CREATING ALL ADDITIONAL WORKSHEETS"
        
    ' ============================================================
    '            STANDARDIZED SALESFORCE DATA WORKSHEET
    ' ============================================================
        ' This worksheet standardizes the Initial Report data into a single structure.
        ' This allows the converter to work from one consistent layout later, regardless of whether the Initial Report came from Intacct or Salesforce.
            Set wsStandardizedSF = wbMacro.Worksheets.Add(After:=wsInitialData)
            wsStandardizedSF.Name = "Standardized Salesforce"
            
    ' ============================================================
    '           STANDARDIZED DONATION SITE DATA WORKSHEET
    ' ============================================================
        ' This worksheet standardizes the Donation Site Report data into one consistent structure.
        ' This allows later sections of the converter to work from one Donation Site layout instead of relying on source-specific column arrangements.
          ' This helps to standardize the import file creation process across all Donation Site platforms.
            Set wsStandardizedDonationSiteData = wbMacro.Worksheets.Add(After:=wsConsolidatedData)
            wsStandardizedDonationSiteData.Name = "Standardized Donation Site Data"
            
    ' ============================================================
    '                  DISBURSEMENT DATA WORKSHEET
    ' ============================================================
        ' This worksheet groups related Donation Site transactions into their corresponding disbursements.
        ' This is required so the converter can summarize activity at the disbursement level and use that information later for Fees, Bank Deposits, and import-file creation.
            Set wsDisbursementData = wbMacro.Worksheets.Add(After:=wsStandardizedDonationSiteData)
            wsDisbursementData.Name = "Disbursement Data"
            
    ' ============================================================
    '                RELEVANT TRANSACTIONS WORKSHEET
    ' ============================================================
        ' This worksheet connects Donation Site Report data to Salesforce/Intacct data.
        ' It exists to keep only the Donation Site transactions that are relevant for the converter based on what is found in Salesforce/Intacct.
        ' It uses the Transaction ID given from the Donation Site Reports to match across all platforms.
            Set wsRelevantTransactions = wbMacro.Worksheets.Add(After:=wsDisbursementData)
            wsRelevantTransactions.Name = "Relevant Transactions"
            
    ' ============================================================
    '                        FEES WORKSHEET
    ' ============================================================
        ' This worksheet isolates the fee portion of each disbursement.
        ' This is needed so fees can be separated from donation amounts and later used as their own line items in the import file.
            Set wsFees = wbMacro.Worksheets.Add(After:=wsRelevantTransactions)
            wsFees.Name = "Fees"
    
    ' ============================================================
    '            BANK DEPOSITS WORKSHEET (IF APPLICABLE)
    ' ============================================================
        ' This worksheet is used only for the "Adjusting" JournalType.
        ' It exists because Adjusting journals require a bank deposit line item, while CRJs do not.
        ' The worksheet uses the net disbursement amount to create the bank deposit line item needed later in the "Adjusting" journal path.
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
        ' It exists so unresolved issues can be handled and updated directly in the workbook before the final import file is created.
        ' It provides the following checks:
          ' BANK ALLOCATIONS NOT FOUND
          ' SALESFORCE DATA MISSING SCHOOL NAME
          ' ADJUSTMENTS TO: ACCOUNT|DIVISION|FUNDING SOURCE
          ' DONATION SITE VS SALESFORCE: GROSS AMOUNT MISMATCHES
          ' TRANSACTIONS WITH MISSING PMT-IDS
            Set wsUserRequiredAdjustments = wbMacro.Worksheets.Add(After:=wsConnectionAnalysis)
            wsUserRequiredAdjustments.Name = "User-Required Adjustments"
    
    ' ============================================================
    '              SCHOOL VALIDATION WORKSHEET
    ' ============================================================
        ' This worksheet stores the approved school validation lists used for dropdown selections throughout the converter.
        ' Defining and preparing it here allows the same validation source to be re-used across multiple exception sections within the "User-Required Adjustments" worksheet.
    
        ' Run the pre-established procedure that creates or refreshes the School Validation worksheet.
            School_Validation.Validation
    
        ' Assign the worksheet to a variable for re-use throughout the macro.
            Set wsSchoolValidation = wbMacro.Worksheets("School Validation")
    
        ' Determine the last populated row using column B.
        ' Column C holds the school names that will be used for data validation dropdowns.
            LastRow_SchoolValidation = wsSchoolValidation.Cells(wsSchoolValidation.Rows.Count, 2).End(xlUp).Row
    
        ' Store the validation range (school names in column C) so it can be re-used later without recalculating the range for each section.
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
                    ' This worksheet removes line items and full disbursements tied to unresolved issues found in the User-Required Adjustments worksheet.
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
                    ' This worksheet removes line items and full disbursements tied to unresolved issues found in the User-Required Adjustments worksheet.
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
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section builds the Standardized Salesforce Data worksheet.
        ' The build involves standardizing the Initial Report from either Salesforce or Intacct into one standard format.
        ' Once standardized, it provides consistent column header placement to seamlessly connect with the Standardized Donation Site Data and assist in the creation of _
          the Relevant Transactions worksheet later on.
        ' It establishes the groundwork for the majority of the data points being used in the final Intacct import file.
                
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "STANDARDIZING THE INITIAL REPORT"
        
    ' ============================================================
    '       FIND THE LAST ROW FROM THE INITIAL DATA WORKSHEET
    ' ============================================================
        ' Determine the last used row in wsInitialData so the formulas can reference the full Initial Report data range.
            LastRow_InitialData = wsInitialData.Cells(wsInitialData.Rows.Count, 1).End(xlUp).Row
    
    ' ============================================================
    '      POPULATE THE WORKSHEET BASED ON THE INITIAL REPORT
    ' ============================================================
        ' The Initial Report can come from either Intacct or Salesforce.
        ' This section standardizes both report types into the same structure so later sections of the converter can use one consistent layout.
      
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
                        
                        ' "Intacct - Location ID"
                            wsStandardizedSF.Range("M2").Formula2 = "=ConvertSchoolAbbrevToIntacctAccount(ConvertSFCampaignNameToSchoolAbbrev(" & _
                                    "IFERROR(IFERROR(IFERROR(LEFT(J2,SEARCH(""20"",J2)-1),LEFT(J2,SEARCH(""AZ Tax"",J2)-1)),LEFT(J2,SEARCH(""General Fund"",J2)-1)),J2)))"
                        
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
                        
                      ' Columns S:V formulas are re-created later (after the User-Required Adjustments worksheet is populated).
                        ' Search: "SUBSECTION: SALESFORCE DATA MISSING SCHOOL NAME"
                        
                        ' "Location Correction"
                          ' This formula will be established later, once the User-Required Adjustments worksheet is populated.
                          ' Search "POPULATE STANDARDIZED INITIAL REPORT: LOCATION CORRECTION"
                            wsStandardizedSF.Range("S2").Formula = "=M2"
                                    
                        ' "Account Correction"
                          ' This formula will be established later, once the User-Required Adjustments worksheet is populated.
                          ' Search "POPULATE STANDARDIZED INITIAL REPORT: ACCOUNT CORRECTION"
                            wsStandardizedSF.Range("T2").Formula = "=N2"
                                    
                        ' "Division Correction"
                          ' This formula will be established later, once the User-Required Adjustments worksheet is populated.
                          ' Search "POPULATE STANDARDIZED INITIAL REPORT: DIVISION CORRECTION"
                            wsStandardizedSF.Range("U2").Formula = "=O2"
                                    
                        ' "Funding Source Correction"
                          ' This formula will be established later, once the User-Required Adjustments worksheet is populated.
                          ' Search "POPULATE STANDARDIZED INITIAL REPORT: FUNDING SOURCE CORRECTION"
                            wsStandardizedSF.Range("V2").Formula = "=P2"
                                    
                        ' "Debt Services Correction"
                            wsStandardizedSF.Range("W2").Formula = "=Q2"
    
                    End If
 
    ' ============================================================
    ' FIND THE LAST ROW FROM THE STANDARDIZED SALESFORCE WORKSHEET
    ' ============================================================
        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            ' Use column F because it will consistently contain the Salesforce Payment ID, which is never blank.
                LastRow_StandardizedSF = wsStandardizedSF.Cells(wsStandardizedSF.Rows.Count, 6).End(xlUp).Row
        
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            ' If the Initial Report came from Salesforce, fill down the formulas not already populated by the formula spanning from columns A:K.
            ' Otherwise, filling down is not necessary.
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
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section builds the Standardized Donation Site Data worksheet.
        ' The build involves standardizing the Consolidated Donation Site data into one standard format, consistent across all Donation Site Platforms.
        ' Once standardized, it provides consistent column header placement to seamlessly connect with the Standardized Initial Report Data and assist in the creation of _
          the Relevant Transactions worksheet later on.
        ' It establishes the groundwork for any of the data points, not coming from the Standardized Salesforce data worksheet, to be used in the final Intacct import file.

    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "STANDARDIZING DONATION SITE DATA"
    
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
              ' The Text function is used to standardize the Transaction ID to be a string, to match with how the data comes from the standardized
                wsStandardizedDonationSiteData.Range("D2").Formula = "=TEXT('" & wsConsolidatedData.Name & "'!C2,""#"")"
            
            ' "Disbursement ID"
                wsStandardizedDonationSiteData.Range("E2").Formula = "=TEXT('" & wsConsolidatedData.Name & "'!J2,""#"")"
            
            ' "Donation Method"
              ' This field is not currently available in the Your Cause Reports.
                wsStandardizedDonationSiteData.Range("F2").Formula = "="""""
            
            ' "Check Number"
              ' This field is not currently available in the Your Cause Reports.
                wsStandardizedDonationSiteData.Range("G2").Formula = "="""""
            
            ' "Donor Name (Last Name, First Name)"
              ' Build the donor name in a consistent format for later matching, review, and analysis.
              ' If the primary donor name is blank, use the matching donor fields, when available.
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
              ' If this returns "No School Found" then the school name needs to be added to the function and assigned to the correct school.
                wsStandardizedDonationSiteData.Range("O2").Formula2 = "=ConvertYourCauseToSchoolAbbrev(N2)"
            
            ' Column P updated later:
              ' Search: "SUBSECTION: BANK ALLOCATIONS NOT FOUND"
            ' "Corrected - School Abbreviation"
              ' This formula will be established later, once the User-Required Adjustments worksheet is populated.
              ' Search: "POPULATE STANDARDIZED DONATION SITE: CORRECTED - SCHOOL ABBREVIATION"
                wsStandardizedDonationSiteData.Range("P2").Formula = "=O2"

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
            ' Use column D because it will consistently be poulated, using the Donation Site Report's Transaction ID.
                LastRow_StandardizedDonationSiteData = wsStandardizedDonationSiteData.Cells(wsStandardizedDonationSiteData.Rows.Count, 4).End(xlUp).Row
    
    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        wsStandardizedDonationSiteData.Range("A1:P1").AutoFilter
        wsStandardizedDonationSiteData.Columns("A:P").AutoFit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' POPULATE THE DISBURSEMENT DATA WORKSHEET '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section builds the Disbursement Data worksheet.
        ' All data in this worksheet comes from the Standardized Consolidated Donation Site data worksheet.
        ' The Disbursement Data worksheet lays the foundation for consolidating disbursement fees, creating bank deposit amounts, _
          and standardizing deposit-level journal entry naming conventions.
        
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "POPULATING THE DISBURSEMENT DATA WORKSHEET"
        
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
              ' This is the amount used when creating the bank deposits amount, tied to each disbursement.
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
            ' Fill down all formulas that are not populated by the UNIQUE() function in column C.
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''----------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' POPULATE THE RELEVANT TRANSACTIONS WORKSHEET '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''----------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section drives the rest of the conversion process. It establishes the Relevant Transactions, connecting between all systems.
        ' This section uses Transaction ID to connect data between Salesforce/Intacct and the Donation Site.
            ' It first uses what is in the Standardized Salesforce worksheet and then matches the Transaction IDs present in the Standardized Donation Site worksheet.
        ' This section sets up the data to work for either the "Adjusting" or "CRJ" route.
        
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "POPULATING THE RELEVANT TRANSACTIONS WORKSHEET."
    
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
                
                If InitialPath = "Intacct" Then
                    wsRelevantTransactions.Range("S2").Formula2 = "=XLOOKUP($B2,'" & wsStandardizedSF.Name & "'!$F:$F,'" & wsStandardizedSF.Name & "'!L:L)&"" Reclassed Out"""
                
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
              ' Create a separation from the standardized data and the "Adjusting" Journal route data format
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
                  ' Create a separation from the the "Adjusting" Journal route data format and the "CRJ" route data format
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
            ' Use column B because it holds the filtered PMT-IDs that drives the worksheet.
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' POPULATE THE FEES WORKSHEET '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section builds the Fees worksheet.
        ' It uses the fees from the Disbursement Data worksheet and creates a standardized format to help create the data entering Intacct Import file later.
            ' Formatting for both the "Adjusting" or "CRJ" routes are included in this worksheet.
        
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
            Application.StatusBar = "POPULATING THE FEES WORKSHEET"

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
              ' Create a separation from the standardized data and the "Adjusting" Journal route data format
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
                  ' Create a separation from the the "Adjusting" Journal route data format and the "CRJ" route data format
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' POPULATE THE BANK DEPOSITS WORKSHEET '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section builds the Bank Deposits worksheet for the "Adjusting" Journal route.
        ' It uses the disbursements from the Disbursement Data worksheet to provide line items as the net value.
        ' This worksheet is best viewed as the piece to help with Bank Reconciliations by consolidating the entire deposit amount into 1 line item.
        
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
            Application.StatusBar = "POPULATING THE BANK DEPOSITS WORKSHEET"

    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
      ' This worksheet creates one bank deposit line item per Disbursement ID for the Adjusting Journal path.
      ' It exists so each disbursement's net amount can be recorded as the bank-side entry that offsets the related revenue and fee lines.
        
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Main fields used to stage the unique Disbursement IDs before building the bank deposit journal lines.
                wsBankDeposits.Range("A1:B1").Value = Array("Disbursement ID", ".......")
            
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
                    
                ' "......."
                  ' Create a separation from the standardized data and the "Adjusting" Journal route data format
                    wsBankDeposits.Range("B2").Value = "......."
            
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
    '             FIND THE LAST ROW FROM THE WORKSHEET
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' POPULATE THE CONNECTION ANALYSIS WORKSHEET ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section builds the Connection Analysis worksheet.
        ' The Connection Analysis worksheet uses the Donation Site Report's data to:
            ' Identify which transactions could not be matched within the Relevant Transactions worksheet.
            ' Identify which transactions have mismatching Gross Amounts between the Donation Site and Salesforce/Intacct.
        ' From these transactions being identified, the User-Required Adjustments worksheet analyzes the results to determine which transactions to bring into their respective _
          sections for review.
        
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
            Application.StatusBar = "POPULATING THE CONNECTION ANALYSIS WORKSHEET"
    
    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
      ' This worksheet compares Donation Site data against Salesforce data at the transaction level.
      ' It also stages the journal fields that will later flow into the User-Required Adjustments worksheet and the final import file paths.
        
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
                        "GLENTRY_PROJECTID", "GLENTRY_CUSTOMERID", "GLENTRY_VENDORID", "GLENTRY_EMPLOYEEID", "GLENTRY_ITEMID", "GLENTRY_CLASSID", ".......")
            
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
              ' If the variance does not equal 0, the User-Required Adjustments worksheet, pulls the transaction in.
              ' Search: "SUBSECTION: TRANSACTIONS: GROSS AMOUNT MISMATCHES"
                wsConnectionAnalysis.Range("G2").Formula = "=ROUND(E2-F2,2)"
            
            ' "PMT-ID"
              ' Pull the matching PMT-ID from the standardized Salesforce data.
              ' If no PMT-ID is found, use "PMT-NOT MATCHED" to flag the transaction for later review.
              ' Search: "SUBSECTION: TRANSACTIONS WITH MISSING PMT-IDS"
                wsConnectionAnalysis.Range("H2").Formula2 = "=XLOOKUP($A2,'" & wsStandardizedSF.Name & "'!$B:$B,'" & wsStandardizedSF.Name & "'!F:F,""PMT-NOT MATCHED"")"
            
            ' "Donation Type"
              ' Pull the Donation Type from the standardized Donation Site data.
                wsConnectionAnalysis.Range("I2").Formula2 = "=XLOOKUP($A2,'" & wsStandardizedDonationSiteData.Name & "'!$D:$D,'" & wsStandardizedDonationSiteData.Name & "'!J:J)"
            
            ' "Site - School Abbreviation"
              ' Pull the corrected school abbreviation from the standardized Donation Site data.
                wsConnectionAnalysis.Range("J2").Formula2 = "=XLOOKUP($A2,'" & wsStandardizedDonationSiteData.Name & "'!$D:$D,'" & wsStandardizedDonationSiteData.Name & "'!P:P)"
            
            ' "......."
              ' Create a separation from the standardized data and the "Adjusting" Journal route data format
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
                    If InitialPath = "Intacct" Then
                        wsConnectionAnalysis.Range("V2").Formula2 = "=SUBSTITUTE(" & _
                                "XLOOKUP(H2,'" & wsRelevantTransactions.Name & "'!K:K,'" & _
                                wsRelevantTransactions.Name & "'!S:S&"" Payment Adjustment""," & _
                                """Payment Adjustment -- Site: " & DonationSite & " | Transaction ID: ""&A2&"" | Disbursement ID: ""&B2&"" | ""&H2)," & _
                            """ Reclassed Out"", """")"
                    Else
                        wsConnectionAnalysis.Range("V2").Formula2 = "=""Payment Adjustment -- ""&XLOOKUP(H2,'" & wsRelevantTransactions.Name & "'!K:K,'" & _
                                wsRelevantTransactions.Name & "'!S:S,""Site: " & DonationSite & " | Transaction ID: ""&A2&"" | Disbursement ID: ""&B2&"" | ""&H2)"
                    End If
                    
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
                    
                ' "......."
                  ' Create a separation from the the "Adjusting" Journal route data format and the "CRJ" route data format
                    wsConnectionAnalysis.Range("AR2").Formula2 = "......."
                    
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
                    If InitialPath = "Intacct" Then
                        wsConnectionAnalysis.Range("BM2").Formula2 = "=SUBSTITUTE(" & _
                                "XLOOKUP(H2,'" & wsRelevantTransactions.Name & "'!K:K,'" & _
                                wsRelevantTransactions.Name & "'!S:S&"" Payment Adjustment""," & _
                                """Payment Adjustment -- Site: " & DonationSite & " | Transaction ID: ""&A2&"" | Disbursement ID: ""&B2&"" | ""&H2)," & _
                            """ Reclassed Out"", """")"
                    Else
                        wsConnectionAnalysis.Range("BM2").Formula2 = "=""Payment Adjustment -- ""&XLOOKUP(H2,'" & wsRelevantTransactions.Name & "'!K:K,'" & _
                                wsRelevantTransactions.Name & "'!S:S,""Site: " & DonationSite & " | Transaction ID: ""&A2&"" | Disbursement ID: ""&B2&"" | ""&H2)"
                    End If
                
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section builds the User-Required Adjustments worksheet.
        
        ' This worksheet acts as the exception-management are for the converter.
            ' Any records that still contain missing, conflicting, or unresolved information are intentionally redirected here so the user can review them before final import.
            ' Until resolved, the whole disbursement (using the journal-name field), the specific transaction lives in, is excluded from the final Intacct Import File.
                ' Once resolved, the full disbursement with the updated information is allowed back into the Intacct Import File.
        
        ' This worksheet intends to resolve the following exceptions:
            ' The Donation Site's School name cannot be matched using the dedicated (Donation Site School Name >> BASIS School Abbreviation) Function.
            ' Transactions that cannot be matched using the dedicated (Salesforce Campaign Name >> BASIS School Abbreviation) Function. ---- Only relevant to "Salesforce" Initial Paths
            ' Transactions that cannot be assigned proper Intacct accounts/fields (Revenue Account, Division, or Funding Source). ---- Only relevant to "Salesforce" Initial Paths
            ' Gross Amount mismatches between the Donation Site Reports and Salesforce/Intacct.
        
        ' The last exception it filters out, but cannot be directly resolved in this workbook, is when PMT-IDs from Salesforce are not synced _
          into Intacct or the transaction does not exist in Salesforce. ''' Note: This exception also appears if the the Transaction ID was not imported into Salesforce initially. '''
            ' Donation Site Data and Salesforce/Intacct Data is matched using the Transaction ID from the Donation Site.
            ' This exception requires the disbursement to be processed at a later time when the issue is resolved in either Salesforce or Intacct, or both.
            ' This section of the worksheet serves two purposes, by creating a list of file names that could not be matched:
                ' It allows the user to have a comprehensive list to send to the Salesforce Administrator to find which transactions have not been processed or synced.
                ' It allows the MOVE MISSING SOURCE FILES section later in the converter to have a list of file names to move to a dedicated 'Process Later' folder _
                  to allow easy separation of files processed versus non-processed by this converter.
                  
        ' Each section below isolates one type of issue so the user can work through it in a controlled, visible way instead of hunting through multiple worksheets.
        
        ' Columns from previous sections that the converter had not populated, now get populated with correct formulas using the ranges created in this section's process.
        
        ' A section turning green means no exceptions were found for that category.
        ' A red section means at least one transaction or disbursement requires user review.
        
        ' The grouped rows are used so the worksheet can remain readable even when all sections are present. The user can expand only the sections that need attention.
        
        
    ' ============================================================
    '                  UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "POPULATING THE USER-REQUIRED ADJUSTMENTS WORKSHEET"

    ' ============================================================
    '                     SET WORKSHEET COUNTER
    ' ============================================================
        ' This counter helps determine if the User-Required Adjustments worksheet needs to be reviewed by the user.
          ' If no issues are found, this counter will be 5.
          ' A final count of 5, will allow the worksheet to be hidden, because no user-required adjustments are needed.
          ' If the count is 4 and the only issue is missing PMT-IDs, then the worksheet will remain in view, with a red tab.
          ' Otherwise, the worksheet tab will appear yellow, to remind the user to make the required adjustments.
          ' By default, assign to 0.
            OK_UserRequiredAdjustments = 0
        
    ' ============================================================
    '            SUBSECTION: BANK ALLOCATIONS NOT FOUND
    ' ============================================================
        ' This subsection captures transactions whose Donation Site School Name could not be converted into a usable BASIS School Abbreviation.
        ' This subsection is intended to serve as a manual override when the Donation Site adds new school names or the school name is not present in the Donation Site Report.
        ' The Donation Site School Name is critical to link the disbursement to a specific bank account.
            ' If the converter cannot identify the school correctly, then it cannot safely determine where that money belongs. Rather than guessing, those transactions _
              are sent here so the user can manually assign the correct school.
        ' Once the school is corrected, the matching disbursement can be released back into the final import flow.

        ' ---------------------------------------------
        '       INITIATE SUBSECTION ROW VARIABLES
        ' ---------------------------------------------
            ' The initiation of these rows, assigns the section header row, the column header row, and the data start row.
                SectionHeaderRow_UserRequiredAdjustments = 1
                HeaderRow_UserRequiredAdjustments = SectionHeaderRow_UserRequiredAdjustments + 1
                DataStartRow_UserRequiredAdjustments = HeaderRow_UserRequiredAdjustments + 1
            
        ' ---------------------------------------------
        '               SUBSECTION HEADER
        ' ---------------------------------------------
            ' Provide a section header for the user to easily navigate between section reviews.
                With wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments)
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .Value = "BANK ALLOCATIONS NOT FOUND"
                    .Interior.Color = vbRed
                    .Font.Bold = True
                    .Font.Underline = xlUnderlineStyleSingle
                End With
            
        ' ---------------------------------------------
        '           SUBSECTION COLUMN HEADERS
        ' ---------------------------------------------
            ' Populate the column headers relevant to this subsection to help the user easily determine the Donation Site School Name allocation.
                With wsUserRequiredAdjustments.Range("A" & HeaderRow_UserRequiredAdjustments & ":N" & HeaderRow_UserRequiredAdjustments)
                    .Value = Array("Disbursement ID", "Transaction ID", "Transaction Date", "Disbursement Date", "Donation Type", "Site - School Name", _
                          "Site - School Abbreviation", "Corrected - School Name", "Corrected - School Abbreviation", "", "", _
                          "Adjusting Journal Name", "CRJ Journal Name", "File Name")
                    .Font.Bold = True
                    .Font.Underline = xlUnderlineStyleSingle
                End With

        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' Pull only Donation Site rows where the converted school abbreviation equals "No School Found".
            ' If nothing is found, return "All Bank Allocations Found" so the section remains explicit and readable.
            ' "Disbursement ID", "Transaction ID", "Transaction Date", "Disbursement Date", "Donation Type", "Site - School Name", "Site - School Abbreviation"
                wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                    "=IFERROR(IF(ISBLANK(CHOOSECOLS(FILTER('" & wsStandardizedDonationSiteData.Name & "'!A2:O" & LastRow_StandardizedDonationSiteData & _
                    ",'" & wsStandardizedDonationSiteData.Name & "'!O2:O" & LastRow_StandardizedDonationSiteData & _
                    "=""No School Found""),4,5,1,2,10,14,15)),""""," & _
                    "CHOOSECOLS(FILTER('" & wsStandardizedDonationSiteData.Name & "'!A2:O" & LastRow_StandardizedDonationSiteData & ",'" & _
                    wsStandardizedDonationSiteData.Name & "'!O2:O" & LastRow_StandardizedDonationSiteData & _
                    "=""No School Found""),5,4,1,2,10,14,15)),""All Bank Allocations Found"")"
                
            ' "Corrected - School Name"
              ' Column H is intentionally left blank for user input.
                ' The user selects the corrected school name from a validation dropdown.
            
            ' "Corrected - School Abbreviation"
                wsUserRequiredAdjustments.Range("I" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                    "=IF(H" & DataStartRow_UserRequiredAdjustments & "="""","""",ConvertSFCampaignNameToSchoolAbbrev(H" & DataStartRow_UserRequiredAdjustments & "))"
                    
            ' ""
              ' Intentionally left blank to standardize columns L:N to have the "Adjusting Journal Name", "CRJ Journl Name", "File Name" data.
            
            ' ""
              ' Intentionally left blank to standardize columns L:N to have the "Adjusting Journal Name", "CRJ Journl Name", "File Name" data.
            
            ' "Adjusting Journal Name"
              ' Keep the related Adjusting Journal description active until the issue is resolved.
                wsUserRequiredAdjustments.Range("L" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All Bank Allocations Found"",""""," & _
                        "IFERROR(IF(I" & DataStartRow_UserRequiredAdjustments & "="""",XLOOKUP(A" & DataStartRow_UserRequiredAdjustments & ",'" & _
                        wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!K:K),""CLEARED""),""""))"
               
            ' "CRJ Journal Name"
              ' Keep the related CRJ description active until the issue is resolved.
                wsUserRequiredAdjustments.Range("M" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All Bank Allocations Found"",""""," & _
                        "IFERROR(IF(I" & DataStartRow_UserRequiredAdjustments & "="""",XLOOKUP(A" & DataStartRow_UserRequiredAdjustments & ",'" & _
                        wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!J:J),""CLEARED""),""""))"
                
            ' "File Name"
              ' Store the original file name tied to the disbursement for traceability.
                wsUserRequiredAdjustments.Range("N" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                    "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All Bank Allocations Found"",""""," & _
                    "XLOOKUP(A" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!M:M))"

        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            ' Determine last row to fill relevant data down and to help create last row subsection ranges, used later.
                LastRow_UserRequiredAdjustments = wsUserRequiredAdjustments.Cells(wsUserRequiredAdjustments.Rows.Count, 1).End(xlUp).Row
                
        ' ---------------------------------------------
        '                GROUP SUBSECTION
        ' ---------------------------------------------
            ' Grouping this subsection allows for user to collapse and expand it, helping them focus on one subsection at a time.
            ' Continue to show the subsection header.
                wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments & ":" & (LastRow_UserRequiredAdjustments + 4)).Group
            
        ' ---------------------------------------------
        '          HIGHLIGHT NECESSARY COLUMNS
        ' ---------------------------------------------
            ' Highlight the "Corrected - School" column to indicate to the user, where to make any necessary adjustments.
                With wsUserRequiredAdjustments.Range("H" & HeaderRow_UserRequiredAdjustments & ":H" & LastRow_UserRequiredAdjustments)
                    .Interior.Color = vbYellow
                    .Locked = False
                End With
                
        ' ---------------------------------------------
        '       CREATE THE DATA VALIDATION RULES
        ' ---------------------------------------------
            ' Create a dropdown list for the user to easily assign a BASIS School to each transaction with missing Donation Site School Names.
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
            ' Fill down columns I:N to populate rows with any data appearing in this subsection.
                If LastRow_UserRequiredAdjustments <> DataStartRow_UserRequiredAdjustments Then
                    wsUserRequiredAdjustments.Range("I" & DataStartRow_UserRequiredAdjustments & ":N" & LastRow_UserRequiredAdjustments).FillDown
                End If
        
        ' ---------------------------------------------
        '          DETERMINE SUBSECTION STATUS
        ' ---------------------------------------------
            ' Check subsection status. If all allocations are found:
              ' Highlight the subsection header in green.
              ' Hide subsection details by closing the subsection grouping.
              ' Increment the OK_UserRequiredAdjustments Counter.
                If wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Value = "All Bank Allocations Found" Then
                    wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments).Interior.Color = vbGreen
                    wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments).ShowDetail = False
                    OK_UserRequiredAdjustments = OK_UserRequiredAdjustments + 1
                End If

        ' ---------------------------------------------
        '           SAVE SUBSECTION VARIABLES
        ' ---------------------------------------------
            ' Save the unresolved journal-name range for later exclusion from the final import file.
            ' If the section only contains one row, save a single-cell reference instead of a same-cell range.
                DataStartRow_UserRequiredAdjustments_BankAllocations = DataStartRow_UserRequiredAdjustments
                LastRow_UserRequiredAdjustments_BankAllocations = LastRow_UserRequiredAdjustments

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
            ' Now that the BANK ALLOCATIONS NOT FOUND subsection exists, the converter has a usable lookup range for the corrected school abbreviations.
            ' This update could not be finalized earlier because the lookup range did not exist yet.
            ' Once the section has been populated and its start/end rows are known, the formula in Column P of the Standardized Donation Site Data worksheet can now safely use that range to look up _
              the user-adjusted missing school abbreviations.
    
            ' "POPULATE STANDARDIZED DONATION SITE: CORRECTED - SCHOOL ABBREVIATION"
                wsStandardizedDonationSiteData.Range("P2").Formula2 = _
                    "=IF(O2=""No School Found"",XLOOKUP(E2,'" & wsUserRequiredAdjustments.Name & "'!$A$" & _
                        DataStartRow_UserRequiredAdjustments_BankAllocations & ":$A$" & LastRow_UserRequiredAdjustments_BankAllocations & _
                        ",'" & wsUserRequiredAdjustments.Name & "'!$I$" & _
                        DataStartRow_UserRequiredAdjustments_BankAllocations & ":$I$" & LastRow_UserRequiredAdjustments_BankAllocations & ",O2),O2)"
            
            ' Fill the formula down in Column P of the Standardized Donation Site Data worksheet
                If LastRow_StandardizedDonationSiteData > 2 Then
                    wsStandardizedDonationSiteData.Range("P2:P" & LastRow_StandardizedDonationSiteData).FillDown
                End If

    ' ============================================================
    '        SUBSECTION: SALESFORCE DATA MISSING SCHOOL NAME
    ' ============================================================
      ' Note: This subsection is only relevant when the InitialPath is "Salesforce". Otherwise the Intacct School Locations are assigned during the Salesforce to Intacct Sync Process.
        ' This subsection captures Salesforce-side transactions where the school/location assignment could not be determined.
        ' Even if the Donation Site data is usable, if the School Name cannot be extracted from the Salesforce Campaign Name, the Intacct School allocation cannot be determined.
        ' If new schools are added in Salesforce without the (Salesforce Campaign Name >> BASIS School Abbreviation) Function being updated, this serves as a manual override.

        ' ---------------------------------------------
        '        UPDATE SUBSECTION ROW VARIABLES
        ' ---------------------------------------------
            ' The updated rows, allow adequate spacing between sections within the worksheet.
            ' The updates assign the section header row, the column header row, and the data start row.
                SectionHeaderRow_UserRequiredAdjustments = LastRow_UserRequiredAdjustments + 6
                HeaderRow_UserRequiredAdjustments = SectionHeaderRow_UserRequiredAdjustments + 1
                DataStartRow_UserRequiredAdjustments = HeaderRow_UserRequiredAdjustments + 1
        
        ' ---------------------------------------------
        '               SUBSECTION HEADER
        ' ---------------------------------------------
            ' Provide a section header for the user to easily navigate between section reviews.
                With wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments)
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .Value = "SALESFORCE DATA MISSING SCHOOL NAME"
                    .Interior.Color = vbRed
                    .Font.Bold = True
                    .Font.Underline = xlUnderlineStyleSingle
                End With

        ' ---------------------------------------------
        '           SUBSECTION COLUMN HEADERS
        ' ---------------------------------------------
            ' Populate the column headers relevant to this subsection to help the user easily determine the missing Campaign School Name from Salesforce.
                With wsUserRequiredAdjustments.Range("A" & HeaderRow_UserRequiredAdjustments & ":N" & HeaderRow_UserRequiredAdjustments)
                    .Value = Array("Transaction ID", "Disbursement ID", "SF Payment ID", "Primary Contact", "Account Name", "Company Name", "Campaign Name", "Opportunity Name", _
                          "Corrected - School Name", "Corrected - School Abbreviation", "", _
                          "Adjusting Journal Name", "CRJ Journal Name", "File Name")
                    .Font.Bold = True
                    .Font.Underline = xlUnderlineStyleSingle
                End With

        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' Pull in Standardized Salesforce rows where the converted school abbreviation equals "No School Found".
            ' If nothing is found, return "All School Names Found" so the section remains explicit and readable.
              ' "Transaction ID", "Disbursement ID", "SF Payment ID", "Primary Contact", "Account Name", "Company Name", "Campaign Name", "Opportunity Name"
                wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IFERROR(CHOOSECOLS(FILTER('" & wsStandardizedSF.Name & "'!B2:K" & LastRow_StandardizedSF & "," & _
                        "(ISNUMBER(MATCH('" & wsStandardizedSF.Name & "'!F2:F" & LastRow_StandardizedSF & ", '" & _
                            wsRelevantTransactions.Name & "'!K2:K" & LastRow_RelevantTransactions & ", 0)))*" & _
                        "('" & wsStandardizedSF.Name & "'!M2:M" & LastRow_StandardizedSF & "=""No School Found"")),1,2,5,6,7,8,9,10),""All School Names Found"")"
            
            ' "Corrected - School Name"
              ' Column I is intentionally left blank to allow the user to select from data validation dropdown.
            
            ' "Corrected - School Abbreviation"
              ' Derive the school abbreviation from the "Corrected - School Name".
                wsUserRequiredAdjustments.Range("J" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IF(I" & DataStartRow_UserRequiredAdjustments & "="""","""",ConvertSFCampaignNameToSchoolAbbrev(I" & DataStartRow_UserRequiredAdjustments & "))"
                
            ' ""
              ' Intentionally left blank to standardize columns L:N to have the "Adjusting Journal Name", "CRJ Journl Name", "File Name" data.

            ' "Adjusting Journal Name"
              ' Keep the related Adjusting Journal description active until the issue is resolved.
                wsUserRequiredAdjustments.Range("L" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All School Names Found"",""""," & _
                        "IF(I" & DataStartRow_UserRequiredAdjustments & "=""""," & _
                        "IFERROR(" & _
                        "XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!K:K)," & _
                        "XLOOKUP(TEXT(B" & DataStartRow_UserRequiredAdjustments & ",""#""),'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!K:K))," & _
                        """CLEARED""))"

            ' "CRJ Journal Name"
              ' Keep the related CRJ description active until the issue is resolved.
                wsUserRequiredAdjustments.Range("M" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All School Names Found"",""""," & _
                        "IF(I" & DataStartRow_UserRequiredAdjustments & "=""""," & _
                        "IFERROR(" & _
                        "XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!J:J)," & _
                        "XLOOKUP(TEXT(B" & DataStartRow_UserRequiredAdjustments & ",""#""),'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!J:J))," & _
                        """CLEARED""))"

            ' "File Name"
              ' Store the original file name tied to the disbursement for traceability.
                wsUserRequiredAdjustments.Range("N" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All School Names Found"",""""," & _
                        "IFERROR(" & _
                        "XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!J:J)," & _
                        "XLOOKUP(TEXT(B" & DataStartRow_UserRequiredAdjustments & ",""#""),'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!J:J)))"
                        
        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            ' Determine last row to fill relevant data down and to help create last row subsection ranges, used later.
                LastRow_UserRequiredAdjustments = wsUserRequiredAdjustments.Cells(wsUserRequiredAdjustments.Rows.Count, 1).End(xlUp).Row
                
        ' ---------------------------------------------
        '                GROUP SUBSECTION
        ' ---------------------------------------------
            ' Grouping this subsection allows for user to collapse and expand it, helping them focus on one subsection at a time.
            ' Continue to show the subsection header.
                wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments & ":" & (LastRow_UserRequiredAdjustments + 4)).Group
            
        ' ---------------------------------------------
        '          HIGHLIGHT NECESSARY COLUMNS
        ' ---------------------------------------------
            ' Highlight the "Correct - School Name" column to indicate to the user, where to make any necessary adjustments.
                With wsUserRequiredAdjustments.Range("I" & HeaderRow_UserRequiredAdjustments & ":I" & LastRow_UserRequiredAdjustments)
                    .Interior.Color = vbYellow
                    .Locked = False
                End With
               
        ' ---------------------------------------------
        '       CREATE THE DATA VALIDATION RULES
        ' ---------------------------------------------
            ' Create a dropdown list for the user to easily assign a BASIS School to each transaction with missing Salesforce School Names.
                ' Only create the validation list if exception rows actually exist.
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
            ' Fill down columns J:N to populate rows with any data appearing in this subsection.
                If LastRow_UserRequiredAdjustments <> DataStartRow_UserRequiredAdjustments Then
                    wsUserRequiredAdjustments.Range("J" & DataStartRow_UserRequiredAdjustments & ":N" & LastRow_UserRequiredAdjustments).FillDown
                End If
        
        ' ---------------------------------------------
        '          DETERMINE SUBSECTION STATUS
        ' ---------------------------------------------
            ' Check subsection status. If all allocations are found:
              ' Highlight the subsection header in green.
              ' Hide subsection details by closing the subsection grouping.
              ' Increment the OK_UserRequiredAdjustments Counter.
                If wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Value = "All School Names Found" Then
                    wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments).Interior.Color = vbGreen
                    wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments).ShowDetail = False
                    OK_UserRequiredAdjustments = OK_UserRequiredAdjustments + 1
                End If
 
        ' ---------------------------------------------
        '           SAVE SUBSECTION VARIABLES
        ' ---------------------------------------------
            ' Save the unresolved journal-name range for later exclusion from the final import file.
            ' If the section only contains one row, save a single-cell reference instead of a same-cell range.
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
          ' Reminder: This section is only relevant when the InitialPath is "Salesforce".
            ' Now that the SALESFORCE DATA MISSING SCHOOL NAME subsection exists, the converter has a usable lookup range for the corrected school abbreviations.
            ' This update could not be finalized earlier because the lookup range did not exist yet.
            ' Once the section has been populated and its start/end rows are known, the formula in Column S of the Standardized Salesforce Data worksheet can now safely use that range to look up _
              the user-adjusted missing school abbreviations.
    
            ' "POPULATE STANDARDIZED INITIAL REPORT: LOCATION CORRECTION"
                If InitialPath = "Salesforce" Then
                    
                    wsStandardizedSF.Range("S2").Formula2 = _
                        "=IF(M2=""No School Found"",ConvertSchoolAbbrevToIntacctAccount(XLOOKUP(F2,'" & wsUserRequiredAdjustments.Name & "'!$C$" & _
                            DataStartRow_UserRequiredAdjustments_MissingSchoolNames & ":$C$" & LastRow_UserRequiredAdjustments_MissingSchoolNames & _
                            ",'" & wsUserRequiredAdjustments.Name & "'!$K$" & _
                            DataStartRow_UserRequiredAdjustments_MissingSchoolNames & ":$K$" & LastRow_UserRequiredAdjustments_MissingSchoolNames & ")),M2)"
                
                ' Fill Formulas down in Column S of the Standardized Salesforce Data worksheet
                    If LastRow_StandardizedSF > 2 Then
                        wsStandardizedSF.Range("S2:S" & LastRow_StandardizedSF).FillDown
                    End If
                    
                End If

    ' ============================================================
    '  SUBSECTION: ADJUSTMENTS TO: ACCOUNT|DIVISION|FUNDING SOURCE
    ' ============================================================
      ' Note: This subsection is only relevant when the InitialPath is "Salesforce". Otherwise all dimensions are assigned during the Salesforce to Intacct Sync Process.
        ' This subsection captures transactions from the Standardized Salesforce worksheet, that could not confidently derive the Revenue Account, Division, or Funding Source from _
          the Salesforce Campaign Name or requires user allocation (like 'General Fund').

        ' ---------------------------------------------
        '        UPDATE SUBSECTION ROW VARIABLES
        ' ---------------------------------------------
            ' The updated rows, allow adequate spacing between sections within the worksheet.
            ' The updates assign the section header row, the column header row, and the data start row.
                SectionHeaderRow_UserRequiredAdjustments = LastRow_UserRequiredAdjustments + 6
                HeaderRow_UserRequiredAdjustments = SectionHeaderRow_UserRequiredAdjustments + 1
                DataStartRow_UserRequiredAdjustments = HeaderRow_UserRequiredAdjustments + 1
            
        ' ---------------------------------------------
        '               SUBSECTION HEADER
        ' ---------------------------------------------
            ' Provide a section header for the user to easily navigate between section reviews.
                With wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments)
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .Value = "ADJUSTMENTS TO: ACCOUNT|DIVISION|FUNDING SOURCE"
                    .Interior.Color = vbRed
                    .Font.Bold = True
                    .Font.Underline = xlUnderlineStyleSingle
                End With
            
        ' ---------------------------------------------
        '           SUBSECTION COLUMN HEADERS
        ' ---------------------------------------------
            ' Populate the column headers relevant to this subsection to help the user easily determine the Intacct Revenue Accounts, Divisions, and Funding Sources.
                With wsUserRequiredAdjustments.Range("A" & HeaderRow_UserRequiredAdjustments & ":N" & HeaderRow_UserRequiredAdjustments)
                    .Value = Array("Transaction ID", "Disbursement ID", "SF Payment ID", "Primary Contact", "Account Name", "Company Name", "Campaign Name", _
                          "Opportunity Name", "Account Correction", "Division Correction", "Funding Source Correction", _
                          "Adjusting Journal Name", "CRJ Journal Name", "File Name")
                    .Font.Bold = True
                    .Font.Underline = xlUnderlineStyleSingle
                End With

        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' From the Standardized Salesforce worksheet, pull in the unassigned ("Check") or General Fund (Account: 73998) transactions.
            ' If all revenue accounts, divisions, and funding sources are accounted for, display "All Accounts, Divisions, and Funding Sources Found".
              ' "Transaction ID", "Disbursement ID", "SF Payment ID", "Primary Contact", "Account Name", "Company Name", "Campaign Name", "Opportunity Name"
                wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IFERROR(CHOOSECOLS(FILTER('" & wsStandardizedSF.Name & "'!B2:K" & LastRow_StandardizedSF & "," & _
                        "(ISNUMBER(MATCH('" & wsStandardizedSF.Name & "'!F2:F" & LastRow_StandardizedSF & ", '" & _
                            wsRelevantTransactions.Name & "'!K2:K" & LastRow_RelevantTransactions & ", 0)))*" & _
                        "(('" & wsStandardizedSF.Name & "'!N2:N" & LastRow_StandardizedSF & "=73998) + " & _
                        "('" & wsStandardizedSF.Name & "'!N2:N" & LastRow_StandardizedSF & "=""CHECK"")))," & _
                        "1,2,5,6,7,8,9,10)," & _
                        """All Accounts, Divisions, and Funding Sources Found"")"
            
            ' "Account Correction", "Division Correction", "Funding Source Correction"
              ' Columns I:K are intentionally left for user assignment.

            ' "Adjusting Journal Name"
              ' Keep the related Adjusting Journal description active until the issue is resolved.
                wsUserRequiredAdjustments.Range("L" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All Accounts, Divisions, and Funding Sources Found"",""""," & _
                        "IF(AND(I" & DataStartRow_UserRequiredAdjustments & "<>"""",J" & _
                        DataStartRow_UserRequiredAdjustments & "<>"""",K" & DataStartRow_UserRequiredAdjustments & "<>""""),""CLEARED""," & _
                        "IFERROR(XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!K:K)," & _
                        "XLOOKUP(TEXT(B" & DataStartRow_UserRequiredAdjustments & ",""#""),'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!K:K))))"
                    
            ' "CRJ Journal Name"
              ' Keep the related CRJ description active until the issue is resolved.
                wsUserRequiredAdjustments.Range("M" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All Accounts, Divisions, and Funding Sources Found"",""""," & _
                        "IF(AND(I" & DataStartRow_UserRequiredAdjustments & "<>"""",J" & _
                        DataStartRow_UserRequiredAdjustments & "<>"""",K" & DataStartRow_UserRequiredAdjustments & "<>""""),""CLEARED""," & _
                        "IFERROR(XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!J:J)," & _
                        "XLOOKUP(TEXT(B" & DataStartRow_UserRequiredAdjustments & ",""#""),'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!J:J))))"
                        
            ' "File Name"
              ' Store the original file name tied to the disbursement for traceability.
                wsUserRequiredAdjustments.Range("N" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All Accounts, Divisions, and Funding Sources Found"",""""," & _
                        "IFERROR(XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!M:M)," & _
                        "XLOOKUP(TEXT(B" & DataStartRow_UserRequiredAdjustments & ",""#""),'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!M:M)))"

        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            ' Determine last row to fill relevant data down and to help create last row subsection ranges, used later.
                LastRow_UserRequiredAdjustments = wsUserRequiredAdjustments.Cells(wsUserRequiredAdjustments.Rows.Count, 1).End(xlUp).Row
                
        ' ---------------------------------------------
        '                GROUP SUBSECTION
        ' ---------------------------------------------
            ' Grouping this subsection allows for user to collapse and expand it, helping them focus on one subsection at a time.
            ' Continue to show the subsection header.
                wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments & ":" & (LastRow_UserRequiredAdjustments + 4)).Group
            
        ' ---------------------------------------------
        '          HIGHLIGHT NECESSARY COLUMNS
        ' ---------------------------------------------
            ' Highlight the "Account Correction", "Division Correction", and "Funding Source Correction" columns to indicate to the user, where to make any necessary adjustments.
                With wsUserRequiredAdjustments.Range("I" & HeaderRow_UserRequiredAdjustments & ":K" & LastRow_UserRequiredAdjustments)
                    .Interior.Color = vbYellow
                    .Locked = False
                End With
                
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            ' Fill down columns L:N to populate rows with any data appearing in this subsection.
                If LastRow_UserRequiredAdjustments <> DataStartRow_UserRequiredAdjustments Then
                    wsUserRequiredAdjustments.Range("L" & DataStartRow_UserRequiredAdjustments & ":N" & LastRow_UserRequiredAdjustments).FillDown
                End If
        
        ' ---------------------------------------------
        '          DETERMINE SUBSECTION STATUS
        ' ---------------------------------------------
            ' Check subsection status. If all allocations are found:
              ' Highlight the subsection header in green.
              ' Hide subsection details by closing the subsection grouping.
              ' Increment the OK_UserRequiredAdjustments Counter.
                If wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Value = "All Accounts, Divisions, and Funding Sources Found" Then
                    wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments).Interior.Color = vbGreen
                    wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments).ShowDetail = False
                    OK_UserRequiredAdjustments = OK_UserRequiredAdjustments + 1
                End If

        ' ---------------------------------------------
        '           SAVE SUBSECTION VARIABLES
        ' ---------------------------------------------
            ' Save the unresolved journal-name range for later exclusion from the final import file.
            ' If the section only contains one row, save a single-cell reference instead of a same-cell range.
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
          ' Reminder: This section is only relevant when the InitialPath is "Salesforce".
            ' Now that the ADJUSTMENTS TO: ACCOUNT|DIVISION|FUNDING SOURCE subsection exists, the converter has a usable lookup range for the corrected school abbreviations.
            ' This update could not be finalized earlier because the lookup range did not exist yet.
            ' Once the section has been populated and its start/end rows are known, the formulas in Columns T:V of the Standardized Salesforce Data worksheet can now safely use that range to look up _
              the user-adjusted Revenue Account, Division, and Funding Source.
    
            If InitialPath = "Salesforce" Then
            ' "POPULATE STANDARDIZED INITIAL REPORT: ACCOUNT CORRECTION"
                wsStandardizedSF.Range("T2").Formula2 = _
                        "=IF(OR(N2=73998,N2=""CHECK""),XLOOKUP(F2,'" & wsUserRequiredAdjustments.Name & "'!$C$" & _
                            DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ":$C$" & LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & _
                            ",'" & wsUserRequiredAdjustments.Name & "'!$I$" & _
                            DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ":$I$" & LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ",N2),N2)"
        
            ' "POPULATE STANDARDIZED INITIAL REPORT: DIVISION CORRECTION"
                wsStandardizedSF.Range("U2").Formula2 = _
                        "=IF(OR(N2=73998,O2=""CHECK""),XLOOKUP(F2,'" & wsUserRequiredAdjustments.Name & "'!$C$" & _
                            DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ":$C$" & LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & _
                            ",'" & wsUserRequiredAdjustments.Name & "'!$J$" & _
                            DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ":$J$" & LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ",O2),O2)"
                
            ' "POPULATE STANDARDIZED INITIAL REPORT: FUNDING SOURCE CORRECTION"
                wsStandardizedSF.Range("V2").Formula2 = _
                        "=IF(OR(N2=73998,P2=""CHECK""),XLOOKUP(F2,'" & wsUserRequiredAdjustments.Name & "'!$C$" & _
                            DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ":$C$" & LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & _
                            ",'" & wsUserRequiredAdjustments.Name & "'!$K$" & _
                            DataStartRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ":$K$" & LastRow_UserRequiredAdjustments_AccountDivisionFundingAdjustments & ",P2),P2)"
                
            ' Fill the formulas down in Columns T:V of the Standardized Salesforce Data worksheet
                If LastRow_StandardizedSF > 2 Then
                    wsStandardizedSF.Range("T2:V" & LastRow_StandardizedSF).FillDown
                End If
                
            End If

    ' ============================================================
    '       SUBSECTION: TRANSACTIONS: GROSS AMOUNT MISMATCHES
    ' ============================================================
      ' Note: The automatically genered adjustments defaults to True, unless otherwise specified.
        ' Using the Connection Analysis worksheet, this subsection isolates transactions where the Donation Site gross amount does not equal the Salesforce/Intacct gross amount.
        ' This subsection serves two main purposes:
            ' It can be an on/off switch to whether the user wants automatic adjustments to be generated as separate line items for the Intacct Import Files.
            ' It isolates mismatching revenue amounts, to give an organized and easy to track list to be sent to the Salesforce Administrator for Salesforce amount adjustments.
        ' If these transactions are set to "No", then it will remove the entire disbursement from the Intacct Import file.

        ' ---------------------------------------------
        '        UPDATE SUBSECTION ROW VARIABLES
        ' ---------------------------------------------
            ' The updated rows, allow adequate spacing between sections within the worksheet.
            ' The updates assign the section header row, the column header row, and the data start row.
                SectionHeaderRow_UserRequiredAdjustments = LastRow_UserRequiredAdjustments + 6
                HeaderRow_UserRequiredAdjustments = SectionHeaderRow_UserRequiredAdjustments + 1
                DataStartRow_UserRequiredAdjustments = HeaderRow_UserRequiredAdjustments + 1
            
        ' ---------------------------------------------
        '               SUBSECTION HEADER
        ' ---------------------------------------------
            ' Provide a section header for the user to easily navigate between section reviews.
                With wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments)
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .Value = "DONATION SITE VS SALESFORCE: GROSS AMOUNT MISMATCHES"
                    .Interior.Color = vbRed
                    .Font.Bold = True
                    .Font.Underline = xlUnderlineStyleSingle
                End With
            
        ' ---------------------------------------------
        '           SUBSECTION COLUMN HEADERS
        ' ---------------------------------------------
            ' Populate the column headers relevant to this subsection.
                With wsUserRequiredAdjustments.Range("A" & HeaderRow_UserRequiredAdjustments & ":N" & HeaderRow_UserRequiredAdjustments)
                    .Value = Array("Transaction ID", "Disbursement ID", "Transaction Date", "Disbursement Date", "Donation Site - Gross Amount", "SF - Gross Amount", _
                          "Variance", "PMT-ID", "Donation Type", "Site - School Abbreviation", "Adjustment Allowed?", _
                          "Adjusting Journal Name", "CRJ Journal Name", "File Name")
                    .Font.Bold = True
                    .Font.Underline = xlUnderlineStyleSingle
                End With

        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' From the Connection Analysis worksheet, pull in any mismatching revenue amounts.
            ' If there are no mismatches, display "No Mismatching Amounts".
              ' "Transaction ID", "Disbursement ID", "Transaction Date", "Disbursement Date", "Donation Site - Gross Amount", "SF - Gross Amount", "Variance", "PMT-ID", "Donation Type", "Site - School Abbreviation"
                wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IFERROR(FILTER('" & wsConnectionAnalysis.Name & "'!A2:J" & LastRow_ConnectionAnalysis & _
                        ",('" & wsConnectionAnalysis.Name & "'!G2:G" & LastRow_ConnectionAnalysis & "<>0)*" & _
                        "('" & wsConnectionAnalysis.Name & "'!H2:H" & LastRow_ConnectionAnalysis & "<>""PMT-NOT MATCHED"")),""No Mismatching Amounts"")"
        
            ' "Adjustment Allowed?"
              ' By default this should be set to "Yes" unless otherwise specified.
              ' Setting this switch to yes, allows the converter to automatically generate adjustment line items for any mismatching amounts.
                If AllowRevenueAmountAdjustments Then
                    wsUserRequiredAdjustments.Range("K" & DataStartRow_UserRequiredAdjustments).Value = "Yes"
                Else
                    wsUserRequiredAdjustments.Range("K" & DataStartRow_UserRequiredAdjustments).Value = "No"
                End If

            ' "Adjusting Journal Name"
              ' Keep the related Adjusting Journal description active until the issue is resolved.
                wsUserRequiredAdjustments.Range("L" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""No Mismatching Amounts"",""""," & _
                        "IF(K" & DataStartRow_UserRequiredAdjustments & "=""No"",XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & _
                        ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!K:K),""CLEARED""))"
                        
            ' "CRJ Journal Name"
              ' Keep the related CRJ description active until the issue is resolved.
                wsUserRequiredAdjustments.Range("M" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                    "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""No Mismatching Amounts"",""""," & _
                        "IF(K" & DataStartRow_UserRequiredAdjustments & "=""No"",XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & _
                        ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!J:J),""CLEARED""))"
            
            ' "File Name"
              ' Store the original file name tied to the disbursement for traceability and if the AllowRevenueAmountAdjustments is switched off, _
                to be used to create the list of files to be moved into a dedicated "Process Later" folder.
                wsUserRequiredAdjustments.Range("N" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""No Mismatching Amounts"",""""," & _
                        "XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!M:M))"
                    
        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            ' Determine last row to fill relevant data down and to help create last row subsection ranges, used later.
                LastRow_UserRequiredAdjustments = wsUserRequiredAdjustments.Cells(wsUserRequiredAdjustments.Rows.Count, 1).End(xlUp).Row
                
        ' ---------------------------------------------
        '                GROUP SUBSECTION
        ' ---------------------------------------------
            ' Grouping this subsection allows for user to collapse and expand it, helping them focus on one subsection at a time.
            ' Continue to show the subsection header.
                wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments & ":" & (LastRow_UserRequiredAdjustments + 4)).Group
            
        ' ---------------------------------------------
        '          HIGHLIGHT NECESSARY COLUMNS
        ' ---------------------------------------------
            ' Highlight the "Adjustment Allowed?" column to indicate to the user, where to make any adjustments, if needed.
                wsUserRequiredAdjustments.Range("G" & HeaderRow_UserRequiredAdjustments & ":G" & LastRow_UserRequiredAdjustments).Interior.Color = vbYellow
                    
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            ' Fill down columns K:N to populate rows with any data appearing in this subsection.
                If LastRow_UserRequiredAdjustments <> DataStartRow_UserRequiredAdjustments Then
                    wsUserRequiredAdjustments.Range("K" & DataStartRow_UserRequiredAdjustments & ":N" & LastRow_UserRequiredAdjustments).FillDown
                End If
        
        ' ---------------------------------------------
        '          DETERMINE SUBSECTION STATUS
        ' ---------------------------------------------
            ' Check subsection status. If all allocations are found:
              ' Highlight the subsection header in green.
              ' Hide subsection details by closing the subsection grouping.
              ' Increment the OK_UserRequiredAdjustments Counter.
                If wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Value = "No Mismatching Amounts" Then
                    wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments).Interior.Color = vbGreen
                    wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments).ShowDetail = False
                    OK_UserRequiredAdjustments = OK_UserRequiredAdjustments + 1
                End If

        ' ---------------------------------------------
        '           SAVE SUBSECTION VARIABLES
        ' ---------------------------------------------
            ' Save the unresolved journal-name range for later exclusion from the final import file.
            ' If the section only contains one row, save a single-cell reference instead of a same-cell range.
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
    '         SUBSECTION: TRANSACTIONS WITH MISSING PMT-IDS
    ' ============================================================
        ' This subsection isolates transactions that exist in the Donation Site Reports but did not meet one of the following:
            ' Did not have a Salesforce PMT-ID sync to Intacct.
            ' Is not in Salesforce yet.
            ' Does not contain the transaction id (at least in the correct format).
        ' When PMT-IDs are synced, but unmatched, generally that means converters for the Donation Sites were not used for the transactions imported into Salesforce.
        ' This subsection provides an organized and comprehensive list of Transaction IDs to send to the Salesforce Administrator or look more into.
        ' This subsection is also used to pull out any disbursements with 1 or more missing PMT-IDs.
            ' Since this issue cannot be resolved within the converter file, the converter file uses the list from this section to create a list of files to move into a dedicated "Process Later" folder.
                ' The folder is to help the user to easily decipher between the files that were used and which files could not be used, for the Intacct Import File.

        ' ---------------------------------------------
        '        UPDATE SUBSECTION ROW VARIABLES
        ' ---------------------------------------------
            ' The updated rows, allow adequate spacing between sections within the worksheet.
            ' The updates assign the section header row, the column header row, and the data start row.
                SectionHeaderRow_UserRequiredAdjustments = LastRow_UserRequiredAdjustments + 6
                HeaderRow_UserRequiredAdjustments = SectionHeaderRow_UserRequiredAdjustments + 1
                DataStartRow_UserRequiredAdjustments = HeaderRow_UserRequiredAdjustments + 1
                
        ' ---------------------------------------------
        '               SUBSECTION HEADER
        ' ---------------------------------------------
            ' Provide a section header for the user to easily navigate between section reviews.
                With wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments)
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .Value = "TRANSACTIONS WITH MISSING PMT-IDs"
                    .Interior.Color = vbRed
                    .Font.Bold = True
                    .Font.Underline = xlUnderlineStyleSingle
                End With
                
        ' ---------------------------------------------
        '           SUBSECTION COLUMN HEADERS
        ' ---------------------------------------------
            ' Populate the column headers relevant to this subsection to help the user easily determine the ----------------------------------------------------------------------------------------------.
                With wsUserRequiredAdjustments.Range("A" & HeaderRow_UserRequiredAdjustments & ":N" & HeaderRow_UserRequiredAdjustments)
                    .Value = Array("Transaction ID", "Disbursement ID", "Transaction Date", "Disbursement Date", "Donation Site - Gross Amount", "SF - Gross Amount", _
                          "Variance", "PMT-ID", "Donation Type", "Site - School Abbreviation", "", _
                          "Adjusting Journal Name", "CRJ Journal Name", "File Name")
                    .Font.Bold = True
                    .Font.Underline = xlUnderlineStyleSingle
                End With

        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' From the Connection Analysis worksheet, pull in any unmatched PMT-IDs ("PMT_NOT MATCHED").
            ' If all transactions are matched to a PMT-ID, display "All PMT-IDs Found".
              ' "Transaction ID", "Disbursement ID", "Transaction Date", "Disbursement Date", "Donation Site - Gross Amount", "SF - Gross Amount", "Variance", "PMT-ID", "Donation Type", "Site - School Abbreviation"
                wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IFERROR(FILTER('" & wsConnectionAnalysis.Name & "'!A2:J" & LastRow_ConnectionAnalysis & ",'" & _
                        wsConnectionAnalysis.Name & "'!H2:H" & LastRow_ConnectionAnalysis & "=""PMT-NOT MATCHED""),""All PMT-IDs Found"")"
                     
                 
            ' "Adjusting Journal Name"
              ' Keep the related Adjusting Journal description active until the issue is resolved.
                wsUserRequiredAdjustments.Range("L" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All PMT-IDs Found"",""""," & _
                        "XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!K:K))"
                       
            ' "CRJ Journal Name"
              ' Keep the related CRJ description active until the issue is resolved.
                wsUserRequiredAdjustments.Range("M" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All PMT-IDs Found"",""""," & _
                        "XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!J:J))"
            
            ' "File Name"
              ' Store the original file name tied to the disbursement for traceability and to be used to create the list of files to be moved into a dedicated "Process Later" folder.
                wsUserRequiredAdjustments.Range("N" & DataStartRow_UserRequiredAdjustments).Formula2 = _
                        "=IF(A" & DataStartRow_UserRequiredAdjustments & "=""All PMT-IDs Found"",""""," & _
                        "XLOOKUP(B" & DataStartRow_UserRequiredAdjustments & ",'" & wsDisbursementData.Name & "'!C:C,'" & wsDisbursementData.Name & "'!M:M))"

        ' ---------------------------------------------
        '               FIND THE LAST ROW
        ' ---------------------------------------------
            ' Determine last row to fill relevant data down and to help create last row subsection ranges, used later.
                LastRow_UserRequiredAdjustments = wsUserRequiredAdjustments.Cells(wsUserRequiredAdjustments.Rows.Count, 1).End(xlUp).Row
                
        ' ---------------------------------------------
        '                GROUP SUBSECTION
        ' ---------------------------------------------
            ' Grouping this subsection allows for user to collapse and expand it, helping them focus on one subsection at a time.
            ' Continue to show the subsection header.
                wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments & ":" & (LastRow_UserRequiredAdjustments + 4)).Group
            
        ' ---------------------------------------------
        '          HIGHLIGHT NECESSARY COLUMNS
        ' ---------------------------------------------
            ' Highlight the "Transaction ID" column to show the user which Donation Site transactions could not be matched.
            ' This serves as a list to send to the Salesforce Adminstrator, or for the user to do additional research into these specific unmatched transactions.
                wsUserRequiredAdjustments.Range("A" & HeaderRow_UserRequiredAdjustments & ":A" & LastRow_UserRequiredAdjustments).Interior.Color = vbYellow
            
        ' ---------------------------------------------
        '               FILL FORMULAS DOWN
        ' ---------------------------------------------
            ' Fill down columns I:N to populate rows with any data appearing in this subsection.
                If LastRow_UserRequiredAdjustments <> DataStartRow_UserRequiredAdjustments Then
                    wsUserRequiredAdjustments.Range("K" & DataStartRow_UserRequiredAdjustments & ":N" & LastRow_UserRequiredAdjustments).FillDown
                End If
        
        ' ---------------------------------------------
        '          DETERMINE SUBSECTION STATUS
        ' ---------------------------------------------
            ' Check subsection status. If all allocations are found:
              ' Highlight the subsection header in green.
              ' Hide subsection details by closing the subsection grouping.
              ' Increment the OK_UserRequiredAdjustments Counter.
                If wsUserRequiredAdjustments.Range("A" & DataStartRow_UserRequiredAdjustments).Value = "All PMT-IDs Found" Then
                    wsUserRequiredAdjustments.Range("A" & SectionHeaderRow_UserRequiredAdjustments & ":N" & SectionHeaderRow_UserRequiredAdjustments).Interior.Color = vbGreen
                    wsUserRequiredAdjustments.Rows(HeaderRow_UserRequiredAdjustments).ShowDetail = False
                    OK_UserRequiredAdjustments = OK_UserRequiredAdjustments + 1
                    PaymentIDs_Missing = False
                Else
                    PaymentIDs_Missing = True
                End If
        
        ' ---------------------------------------------
        '           SAVE SUBSECTION VARIABLES
        ' ---------------------------------------------
            ' Save the unresolved journal-name range for later exclusion from the final import file.
            ' If the section only contains one row, save a single-cell reference instead of a same-cell range.
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
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' All Intacct Import File prepatory steps are completed. This section directs the converter to the final journal path route.
            ' The journal path routes include the final three worksheets (Unfiltered Import data, Filtered Import Data, and Filtered Import Data in Import File Formatting)
        
    ' ============================================================
    '                ROUTE TO THE FINAL JOURNAL PATH
    ' ============================================================
        ' If JournalType = "CRJ", jump to the CRJ-specific section. Otherwise, continue into the Adjusting Journal path sections below.
            If JournalType = "CRJ" Then
                GoTo JournalPath_CRJ
            End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------------------------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''' (ADJUSTING JOURNAL PATH): POPULATE THE ADJUSTING JOURNAL - UNFILTERED WORKSHEET '''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------------------------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section builds the initial UNFILTERED adjusting journal worksheet.
        ' This worksheet acts as the full staging area for the journal before any unresolved exceptions are removed.
        ' It serves to show a full picture of what the journal could include, if everything was allowed through.
        ' This worksheet pulls together all the pieces that need to be included in the journal entry:
            ' Bank Deposits
            ' Fees
            ' Relevant Transactions
            ' Connection Analysis Adjustment Rows
        ' If any user-required adjustments are made, this is where that data will be updated, before the finalized journal is updated.
        ' Salesforce Reference fields are added in here, to help create line item lookups that can be used when researching data within Intacct Reports.
        
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "CREATING THE INTACCT IMPORT FILE"

    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Columns A:AG hold the Intacct import structure for the Adjusting Journal.
                wsAdjustingUnfiltered.Range("A1:AG1").Value = Array("DONOTIMPORT", "JOURNAL", "DATE", "REVERSEDATE", "DESCRIPTION", "REFERENCE_NO", "LINE_NO", "ACCT_NO", _
                        "LOCATION_ID", "DEPT_ID", "DOCUMENT", "MEMO", "DEBIT", "CREDIT", "SOURCEENTITY", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", _
                        "EXCHANGE_RATE", "STATE", "ALLOCATION_ID", "RASSET", "RDEPRECIATION_SCHEDULE", "RASSET_ADJUSTMENT", "RASSET_CLASS", "RASSETOUTOFSERVICE", _
                        "GLDIMFUNDING_SOURCE", "GLENTRY_PROJECTID", "GLENTRY_CUSTOMERID", "GLENTRY_VENDORID", "GLENTRY_EMPLOYEEID", "GLENTRY_ITEMID", "GLENTRY_CLASSID")
                        
            ' Columns AH:AT hold Salesforce reference fields.
                wsAdjustingUnfiltered.Range("AH1:AT1").Value = Array("SF_CLOSE_DATE", "SF_DONATION_SITE", "SF_CP_NUMBER", "SF_TRANSACTION_ID", "SF_DISBURSEMENT_ID", _
                        "SF_PAYMENT_METHOD", "SF_CHECK_NUMBER", "SF_PAYMENT_NUMBER", "SF_PRIMARY_CONTACT", "SF_ACCOUNT_NAME", "SF_COMPANY_NAME", "SF_CAMPAIGN_SOURCE", _
                        "SF_DONATION_NAME")
                    
        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' Build an UNFILTERED Adjusting Journal body to support the Bank Deposits worksheet, the Fees worksheet, the Relevant Transactions worksheet, and _
              the Connection Analysis worksheet (where any variances exist).
            ' This formula encapsulates all relevant data for the Intacct Import File and positions each piece in a structured, standardized format.
            ' The stacked output begins in column B because column A is reserved for the DONOTIMPORT field used later in the finalized journal worksheet.
                wsAdjustingUnfiltered.Range("B2").Formula2 = _
                        "=LET(BankDeposits,'" & wsBankDeposits.Name & "'!C2:AH" & LastRow_BankDeposits & "," & _
                        "Fees,'" & wsFees.Name & "'!E2:AJ" & LastRow_Fees & "," & _
                        "RelevantTransactions,'" & wsRelevantTransactions.Name & "'!AA2:BF" & LastRow_RelevantTransactions & "," & _
                        "Mismatches,FILTER('" & wsConnectionAnalysis.Name & "'!L2:AQ" & LastRow_ConnectionAnalysis & _
                                           ",'" & wsConnectionAnalysis.Name & "'!G2:G" & LastRow_ConnectionAnalysis & "<>0)," & _
                        "FeeCheck,ISERROR('" & wsFees.Name & "'!A2)," & _
                        "RelevantTransactionsCheck,ISERROR('" & wsRelevantTransactions.Name & "'!B2)," & _
                        "MismatchesCheck,COUNTA(Mismatches)=1," & _
                        "IF(AND(FeeCheck,RelevantTransactionsCheck,MismatchesCheck),SORT(VSTACK(BankDeposits),4)," & _
                        "IF(AND(FeeCheck,RelevantTransactionsCheck),SORT(VSTACK(BankDeposits,Mismatches),4)," & _
                        "IF(AND(FeeCheck,MismatchesCheck),SORT(VSTACK(BankDeposits,RelevantTransactions),4)," & _
                        "IF(AND(RelevantTransactionsCheck,MismatchesCheck),SORT(VSTACK(BankDeposits,Fees),4)," & _
                        "IF(FeeCheck,SORT(VSTACK(BankDeposits,RelevantTransactions,Mismatches),4)," & _
                        "IF(RelevantTransactionsCheck,SORT(VSTACK(BankDeposits,Fees,Mismatches),4)," & _
                        "IF(MismatchesCheck,SORT(VSTACK(BankDeposits,Fees,RelevantTransactions),4)," & _
                        "SORT(VSTACK(BankDeposits,Fees,RelevantTransactions,Mismatches),4)))))))))"

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
    '                       FIND THE LAST ROW
    ' ============================================================
        ' Use column B because the stacked journal body begins in column B.
            LastRow_AdjustingUnfiltered = wsAdjustingUnfiltered.Cells(wsAdjustingUnfiltered.Rows.Count, 2).End(xlUp).Row
        
    ' ============================================================
    '               FILL FORMULAS DOWN
    ' ============================================================
        ' Fill down the Salesforce Reference Field formulas.
            If LastRow_AdjustingUnfiltered > 2 Then
                wsAdjustingUnfiltered.Range("AH2:AT" & LastRow_AdjustingUnfiltered).FillDown
            End If

    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        wsAdjustingUnfiltered.Range("A1:AT1").AutoFilter
        wsAdjustingUnfiltered.Columns("A:AT").AutoFit
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------------------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''' (ADJUSTING JOURNAL PATH): POPULATE THE ADJUSTING JOURNAL - FILTERED WORKSHEET ''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------------------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section builds the Adjusting Journal FILTERED worksheet.
        ' This worksheet uses the User-Required Adjustments worksheet to pull out any disbursements with at least one transaction requiring user's decision making.
            ' Once the user has made their decisions, the updated transactions/disbursements populate into this worksheet to allow it to pass-through into the _
              finalized Adjusting Journal worksheet.
        
        ' NOTE: This worksheet does not allow disbursements to pass through with at least 1 Missing PMT-ID transaction.
            ' These disbursements are placed in a 'Process Later' folder.
        
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "FILTERING OUT ANY MISSING DATA FROM THE INTACCT IMPORT FILE"

    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Keep the same structure as the Unfiltered worksheet.
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
            ' Build a formula that holds all the data found in the UNFILTERED Adjusting Journal worksheet and filters out any disbursements found in _
              the User-Required Adjustments worksheet.
            ' The formula needs to be able to update based on the adjustments the user makes.
            ' It should never include any disbursements with at least 1 missing PMT-ID, because those disbursements are set aside to be processed later.
            ' The formula begins in column B because column A is still reserved for DONOTIMPORT, matching the UNFILTERED worksheet structure.
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
        wsAdjustingFiltered.Range("A1:AT1").AutoFilter
        wsAdjustingFiltered.Columns("A:AT").AutoFit
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''' (ADJUSTING JOURNAL PATH): POPULATE THE ADJUSTING JOURNAL - FINALIZED WORKSHEET ''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''--------------------------------------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section builds the Finalized Adjusting Journal worksheet.
        ' This worksheet uses the FILTERED Adjusting Journal worksheet to build cell-based formulas rather than using a filter formula to display the data.
            ' This worksheet pulls out any disbursement-based repeating data (Columns "JOURNAL", "DATE", and "DESCRIPTION") that is only required on the first line of the journal.
            ' This worksheet creates the finalized line-number ("LINE_NO") for each line item.
                ' All columns other than the 4 mentioned above, pass through exactly the way they appear in the FILTERED Adjusting Journal worksheet.
                
        ' Additionally, this worksheet allows the user to easily copy or manipulate the data, if any additional changes need to be made.
            
        ' This worksheet uses the last row of the UNFILTERED Adjusting Journal worksheet, to establish the rows it allows to have formulas in.
            ' This allows the worksheet to update seamlessly when the FILTERED Adjusting Journal worksheet is updated.
        
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "FINALIZING THE INTACCT IMPORT FILE"

    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Keep the same field structure as the upstream journal worksheets so the final import tab remains consistent with the staging tabs.
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
            ' COLUMN A
                ' "DONOTIMPORT"
                    ' This column intentionally remains blank.
                    
            ' "JOURNAL"
              ' This item is only required on the first line item of each journal.
                wsAdjustingJournal.Range("B2").Formula2 = _
                        "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!B2,""""))),""""," & _
                            "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!B2,"""")))"
            
            ' "DATE"
              ' This item is only required on the first line item of each journal.
                wsAdjustingJournal.Range("C2").Formula2 = _
                        "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!C2,""""))),""""," & _
                            "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!C2,"""")))"
                        
            ' "REVERSEDATE"
                wsAdjustingJournal.Range("D2").Formula2 = _
                        "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!D2,""""))),""""," & _
                            "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!D2,"""")))"
            
            ' "DESCRIPTION"
              ' This item is only required on the first line item of each journal.
                wsAdjustingJournal.Range("E2").Formula2 = _
                        "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!E2,""""))),""""," & _
                            "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!E2,"""")))"
                            
            ' "REFERENCE_NO"
                wsAdjustingJournal.Range("F2").Formula2 = _
                        "=IF(ISBLANK(IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!F2,""""))),""""," & _
                            "IF('" & wsAdjustingFiltered.Name & "'!$B2="""","""",IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!F2,"""")))"
                            
            ' "LINE_NO"
              ' This field is recalculated here instead of being copied straight across.
                wsAdjustingJournal.Range("G2").Formula2 = _
                        "=IF('" & wsAdjustingFiltered.Name & "'!$B2="""",""""," & _
                            "IF('" & wsAdjustingFiltered.Name & "'!$G2=1,'" & wsAdjustingFiltered.Name & "'!G2," & _
                                "1+G1))"
            
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

    ' ============================================================
    '                      FILL FORMULAS DOWN
    ' ============================================================
        ' Fill down based on the Unfiltered worksheet row count.
            If LastRow_AdjustingUnfiltered > 2 Then
                wsAdjustingJournal.Range("B2:AT" & LastRow_AdjustingUnfiltered).FillDown
            End If

    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        wsAdjustingJournal.Range("A1:AT1").AutoFilter
        wsAdjustingJournal.Columns("A:AT").AutoFit

    ' ============================================================
    '   DIRECT THE CONVERTER PAST THE "JOURNAL PATH:CRJ" SECTIONS
    ' ============================================================
        ' The Adjusting Journal path is complete at this point.
        ' Jump to the final steps of the converter.
            GoTo MoveFiles
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' (CRJ JOURNAL PATH): POPULATE THE CRJ - UNFILTERED WORKSHEET '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
JournalPath_CRJ:
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section builds the initial UNFILTERED CRJ Import worksheet.
        ' This worksheet acts as the full staging area for the journal before any unresolved exceptions are removed.
        ' It serves to show a full picture of what the journal could include, if everything was allowed through.
        ' This worksheet pulls together all the pieces that need to be included in the journal entry:
            ' Fees
            ' Relevant Transactions
            ' Connection Analysis Adjustment Rows
            
            ' Note: Unlike the "Adjusting" Journal route, the Bank Deposits section is not included in CRJ Imports.
            
        ' If any user-required adjustments are made, this is where that data will be updated, before the finalized journal is updated.
        ' Salesforce Reference fields are added in here, to help create line item lookups that can be used when researching data within Intacct Reports.
        
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "CREATING THE INTACCT IMPORT FILE"

    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Columns A:AG holds the Intacct import structure for the CRJ Journal.
                wsCRJUnfiltered.Range("A1:AG1").Value = Array("DONOTIMPORT", "RECEIPT_DATE", "PAYMETHOD", "DOCDATE", "DOCNUMBER", "DESCRIPTION", "DEPOSITTO", "BANKACCOUNTID", _
                            "DEPOSITDATE", "UNDEPACCTNO", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCH_RATE_DATE", "LINE_NO", "ACCT_NO", "ACCOUNTLABEL", "TRX_AMOUNT", _
                            "AMOUNT", "DEPT_ID", "LOCATION_ID", "ITEM_MEMO", "OTHERRECEIPTSENTRY_PROJECTID", "OTHERRECEIPTSENTRY_CUSTOMERID", "OTHERRECEIPTSENTRY_ITEMID", _
                            "OTHERRECEIPTSENTRY_VENDORID", "OTHERRECEIPTSENTRY_EMPLOYEEID", "OTHERRECEIPTSENTRY_CLASSID", "PAYER_NAME", "SUPDOCID", "EXCHANGE_RATE", _
                            "OR_TRANSACTION_DATE", "GLDIMFUNDING_SOURCE")
                    
        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' Build an UNFILTERED CRJ body to support the Fees worksheet, the Relevant Transactions worksheet, and the Connection Analysis worksheet (where any variances exist).
            ' This formula encapsulates all relevant data for the Intacct Import File and positions each piece in a structured, standardized format.
            ' The stacked output begins in column B because column A is reserved for the DONOTIMPORT field used later in the finalized CRJ worksheet.
                wsCRJUnfiltered.Range("B2").Formula2 = _
                            "=LET(Fees,'" & wsFees.Name & "'!AL2:BQ" & LastRow_Fees & "," & _
                            "RelevantTransactions,'" & wsRelevantTransactions.Name & "'!BH2:CM" & LastRow_RelevantTransactions & "," & _
                            "Mismatches,FILTER('" & wsConnectionAnalysis.Name & "'!AS2:BX" & LastRow_ConnectionAnalysis & _
                                               ",'" & wsConnectionAnalysis.Name & "'!G2:G" & LastRow_ConnectionAnalysis & "<>0)," & _
                            "FeeCheck,ISERROR('" & wsFees.Name & "'!A2)," & _
                            "RelevantTransactionsCheck,ISERROR('" & wsRelevantTransactions.Name & "'!B2)," & _
                            "MismatchesCheck,COUNTA(Mismatches)=1," & _
                            "IF(AND(FeeCheck,RelevantTransactionsCheck),SORT(VSTACK(Mismatches),4)," & _
                            "IF(AND(FeeCheck,MismatchesCheck),SORT(VSTACK(RelevantTransactions),4)," & _
                            "IF(AND(RelevantTransactionsCheck,MismatchesCheck),SORT(VSTACK(Fees),4)," & _
                            "IF(FeeCheck,SORT(VSTACK(RelevantTransactions,Mismatches),4)," & _
                            "IF(RelevantTransactionsCheck,SORT(VSTACK(Fees,Mismatches),4)," & _
                            "IF(MismatchesCheck,SORT(VSTACK(Fees,RelevantTransactions),4)," & _
                            "SORT(VSTACK(Fees,RelevantTransactions,Mismatches),4))))))))"
                            
    ' ============================================================
    '                       FIND THE LAST ROW
    ' ============================================================
        ' Use column B because the stacked journal body begins in column B.
            LastRow_CRJUnfiltered = wsCRJUnfiltered.Cells(wsCRJUnfiltered.Rows.Count, 2).End(xlUp).Row
            
    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        wsCRJUnfiltered.Range("A1:AG1").AutoFilter
        wsCRJUnfiltered.Columns("A:AG").AutoFit
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' (CRJ JOURNAL PATH): POPULATE THE CRJ - FILTERED WORKSHEET ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section builds the CRJ FILTERED worksheet.
        ' This worksheet uses the User-Required Adjustments worksheet to pull out any disbursements with at least one transaction requiring user's decision making.
            ' Once the user has made their decisions, the updated transactions/disbursements populate into this worksheet to allow it to pass-through into the _
              finalized CRJ Import worksheet.
        
        ' NOTE: This worksheet does not allow disbursements to pass through with at least 1 Missing PMT-ID transaction.
            ' These disbursements are placed in a 'Process Later' folder.
    
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "FILTERING OUT ANY MISSING DATA FROM THE INTACCT IMPORT FILE"

    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Columns A:AG holds the Intacct import structure for the CRJ Journal.
                wsCRJFiltered.Range("A1:AG1").Value = Array("DONOTIMPORT", "RECEIPT_DATE", "PAYMETHOD", "DOCDATE", "DOCNUMBER", "DESCRIPTION", "DEPOSITTO", "BANKACCOUNTID", _
                            "DEPOSITDATE", "UNDEPACCTNO", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCH_RATE_DATE", "LINE_NO", "ACCT_NO", "ACCOUNTLABEL", "TRX_AMOUNT", _
                            "AMOUNT", "DEPT_ID", "LOCATION_ID", "ITEM_MEMO", "OTHERRECEIPTSENTRY_PROJECTID", "OTHERRECEIPTSENTRY_CUSTOMERID", "OTHERRECEIPTSENTRY_ITEMID", _
                            "OTHERRECEIPTSENTRY_VENDORID", "OTHERRECEIPTSENTRY_EMPLOYEEID", "OTHERRECEIPTSENTRY_CLASSID", "PAYER_NAME", "SUPDOCID", "EXCHANGE_RATE", _
                            "OR_TRANSACTION_DATE", "GLDIMFUNDING_SOURCE")
                    
        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' Build a formula that holds all the data found in the UNFILTERED Adjusting Journal worksheet and filters out any disbursements found in _
              the User-Required Adjustments worksheet.
            ' The formula needs to be able to update based on the adjustments the user makes.
            ' It should never include any disbursements with at least 1 missing PMT-ID, because those disbursements are set aside to be processed later.
            ' The formula begins in column B because column A is still reserved for DONOTIMPORT, matching the UNFILTERED worksheet structure.
                wsCRJFiltered.Range("B2").Formula2 = _
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
                            "FILTER('" & wsCRJUnfiltered.Name & "'!B2:AG" & LastRow_CRJUnfiltered & "," & _
                                "NOT(ISNUMBER(MATCH('" & wsCRJUnfiltered.Name & "'!E2:E" & LastRow_CRJUnfiltered & "," & _
                                "UserRequiredAdjustmentsFiltered,0))))" & _
                        ")"
        
    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        wsCRJFiltered.Range("A1:AG1").AutoFilter
        wsCRJFiltered.Columns("A:AG").AutoFit
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' (CRJ JOURNAL PATH): POPULATE THE CRJ - FINALIZED WORKSHEET ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section builds the Finalized CRJ Import File worksheet.
        ' This worksheet uses the FILTERED CRJ worksheet to build cell-based formulas rather than using a filter formula to display the data.
            ' This worksheet creates the finalized line-number ("LINE_NO") for each line item.
                ' All columns other than the 1 mentioned above, pass through exactly the way they appear in the FILTERED Adjusting Journal worksheet.
                
        ' Additionally, this worksheet allows the user to easily copy or manipulate the data, if any additional changes need to be made.
            
        ' This worksheet uses the last row of the UNFILTERED CRJ worksheet, to establish the rows it allows to have formulas in.
            ' This allows the worksheet to update seamlessly when the FILTERED Adjusting Journal worksheet is updated.
    
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "FINALIZING THE INTACCT IMPORT FILE"
        
    ' ============================================================
    '                  POPULATE THE WORKSHEET DATA
    ' ============================================================
        ' ---------------------------------------------
        '                 COLUMN HEADERS
        ' ---------------------------------------------
            ' Columns A:AG holds the Intacct import structure for the CRJ Journal.
                wsCRJ.Range("A1:AG1").Value = Array("DONOTIMPORT", "RECEIPT_DATE", "PAYMETHOD", "DOCDATE", "DOCNUMBER", "DESCRIPTION", "DEPOSITTO", "BANKACCOUNTID", _
                            "DEPOSITDATE", "UNDEPACCTNO", "CURRENCY", "EXCH_RATE_DATE", "EXCH_RATE_TYPE_ID", "EXCH_RATE_DATE", "LINE_NO", "ACCT_NO", "ACCOUNTLABEL", "TRX_AMOUNT", _
                            "AMOUNT", "DEPT_ID", "LOCATION_ID", "ITEM_MEMO", "OTHERRECEIPTSENTRY_PROJECTID", "OTHERRECEIPTSENTRY_CUSTOMERID", "OTHERRECEIPTSENTRY_ITEMID", _
                            "OTHERRECEIPTSENTRY_VENDORID", "OTHERRECEIPTSENTRY_EMPLOYEEID", "OTHERRECEIPTSENTRY_CLASSID", "PAYER_NAME", "SUPDOCID", "EXCHANGE_RATE", _
                            "OR_TRANSACTION_DATE", "GLDIMFUNDING_SOURCE")
                    
        ' ---------------------------------------------
        '          POPULATE DATA USING FORMULAS
        ' ---------------------------------------------
            ' "DONOTIMPORT"
                ' Column A is intentionally left blank
                
            ' "RECEIPT_DATE"
                wsCRJ.Range("B2").Formula = "=IF('" & wsCRJFiltered.Name & "'!B2="""","""",'" & wsCRJFiltered.Name & "'!B2)"
                
            ' "PAYMETHOD"
                wsCRJ.Range("C2").Formula2 = "=IF('" & wsCRJFiltered.Name & "'!C2="""","""",'" & wsCRJFiltered.Name & "'!C2)"
                
            ' "DOCDATE"
                wsCRJ.Range("D2").Formula = "=IF('" & wsCRJFiltered.Name & "'!D2="""","""",'" & wsCRJFiltered.Name & "'!D2)"
                
            ' "DOCNUMBER"
                wsCRJ.Range("E2").Formula = "=IF('" & wsCRJFiltered.Name & "'!E2="""","""",'" & wsCRJFiltered.Name & "'!E2)"
                
            ' "DESCRIPTION"
                wsCRJ.Range("F2").Formula = "=IF('" & wsCRJFiltered.Name & "'!F2="""","""",'" & wsCRJFiltered.Name & "'!F2)"
                
            ' "DEPOSITTO"
                wsCRJ.Range("G2").Formula = "=IF('" & wsCRJFiltered.Name & "'!G2="""","""",'" & wsCRJFiltered.Name & "'!G2)"
                
            ' "BANKACCOUNTID"
                wsCRJ.Range("H2").Formula = "=IF('" & wsCRJFiltered.Name & "'!H2="""","""",'" & wsCRJFiltered.Name & "'!H2)"
                
            ' "DEPOSITDATE"
                wsCRJ.Range("I2").Formula = "=IF('" & wsCRJFiltered.Name & "'!I2="""","""",'" & wsCRJFiltered.Name & "'!I2)"
                
            ' "UNDEPACCTNO"
                wsCRJ.Range("J2").Formula = "=IF('" & wsCRJFiltered.Name & "'!J2="""","""",'" & wsCRJFiltered.Name & "'!J2)"
                
            ' "CURRENCY"
                wsCRJ.Range("K2").Formula = "=IF('" & wsCRJFiltered.Name & "'!K2="""","""",'" & wsCRJFiltered.Name & "'!K2)"
                
            ' "EXCH_RATE_DATE"
                wsCRJ.Range("L2").Formula = "=IF('" & wsCRJFiltered.Name & "'!L2="""","""",'" & wsCRJFiltered.Name & "'!L2)"
                
            ' "EXCH_RATE_TYPE_ID"
                wsCRJ.Range("M2").Formula = "=IF('" & wsCRJFiltered.Name & "'!M2="""","""",'" & wsCRJFiltered.Name & "'!M2)"
                
            ' "EXCH_RATE_DATE"
                wsCRJ.Range("N2").Formula = "=IF('" & wsCRJFiltered.Name & "'!N2="""","""",'" & wsCRJFiltered.Name & "'!N2)"
                
            ' "LINE_NO"
                wsCRJ.Range("O2").Formula = "=IF('" & wsCRJFiltered.Name & "'!O2="""","""",'" & wsCRJFiltered.Name & "'!O2)"
                
            ' "ACCT_NO"
                wsCRJ.Range("P2").Formula = "=IF('" & wsCRJFiltered.Name & "'!P2="""","""",'" & wsCRJFiltered.Name & "'!P2)"
                
            ' "ACCOUNTLABEL"
                wsCRJ.Range("Q2").Formula = "=IF('" & wsCRJFiltered.Name & "'!Q2="""","""",'" & wsCRJFiltered.Name & "'!Q2)"
                
            ' "TRX_AMOUNT"
                wsCRJ.Range("R2").Formula = "=IF('" & wsCRJFiltered.Name & "'!R2="""","""",'" & wsCRJFiltered.Name & "'!R2)"
                
            ' "AMOUNT"
                wsCRJ.Range("S2").Formula = "=IF('" & wsCRJFiltered.Name & "'!S2="""","""",'" & wsCRJFiltered.Name & "'!S2)"
                
            ' "DEPT_ID"
                wsCRJ.Range("T2").Formula = "=IF('" & wsCRJFiltered.Name & "'!T2="""","""",'" & wsCRJFiltered.Name & "'!T2)"
                
            ' "LOCATION_ID"
                wsCRJ.Range("U2").Formula = "=IF('" & wsCRJFiltered.Name & "'!U2="""","""",'" & wsCRJFiltered.Name & "'!U2)"
                
            ' "ITEM_MEMO"
                wsCRJ.Range("V2").Formula = "=IF('" & wsCRJFiltered.Name & "'!V2="""","""",'" & wsCRJFiltered.Name & "'!V2)"
                
            ' "OTHERRECEIPTSENTRY_PROJECTID"
                wsCRJ.Range("W2").Formula = "=IF('" & wsCRJFiltered.Name & "'!W2="""","""",'" & wsCRJFiltered.Name & "'!W2)"
                
            ' "OTHERRECEIPTSENTRY_CUSTOMERID"
                wsCRJ.Range("X2").Formula = "=IF('" & wsCRJFiltered.Name & "'!X2="""","""",'" & wsCRJFiltered.Name & "'!X2)"
                
            ' "OTHERRECEIPTSENTRY_ITEMID"
                wsCRJ.Range("Y2").Formula = "=IF('" & wsCRJFiltered.Name & "'!Y2="""","""",'" & wsCRJFiltered.Name & "'!Y2)"
                
            ' "OTHERRECEIPTSENTRY_VENDORID"
                wsCRJ.Range("Z2").Formula = "=IF('" & wsCRJFiltered.Name & "'!Z2="""","""",'" & wsCRJFiltered.Name & "'!Z2)"
                
            ' "OTHERRECEIPTSENTRY_EMPLOYEEID"
                wsCRJ.Range("AA2").Formula = "=IF('" & wsCRJFiltered.Name & "'!AA2="""","""",'" & wsCRJFiltered.Name & "'!AA2)"
                
            ' "OTHERRECEIPTSENTRY_CLASSID"
                wsCRJ.Range("AB2").Formula = "=IF('" & wsCRJFiltered.Name & "'!AB2="""","""",'" & wsCRJFiltered.Name & "'!AB2)"
                
            ' "PAYER_NAME"
                wsCRJ.Range("AC2").Formula = "=IF('" & wsCRJFiltered.Name & "'!AC2="""","""",'" & wsCRJFiltered.Name & "'!AC2)"
                
            ' "SUPDOCID"
                wsCRJ.Range("AD2").Formula = "=IF('" & wsCRJFiltered.Name & "'!AD2="""","""",'" & wsCRJFiltered.Name & "'!AD2)"
                
            ' "EXCHANGE_RATE"
                wsCRJ.Range("AE2").Formula = "=IF('" & wsCRJFiltered.Name & "'!AE2="""","""",'" & wsCRJFiltered.Name & "'!AE2)"
                
            ' "OR_TRANSACTION_DATE"
                wsCRJ.Range("AF2").Formula = "=IF('" & wsCRJFiltered.Name & "'!AF2="""","""",'" & wsCRJFiltered.Name & "'!AF2)"
                
            ' "GLDIMFUNDING_SOURCE"
                wsCRJ.Range("AG2").Formula = "=IF('" & wsCRJFiltered.Name & "'!AG2="""","""",'" & wsCRJFiltered.Name & "'!AG2)"
            
    ' ============================================================
    '                      FILL FORMULAS DOWN
    ' ============================================================
        ' Fill down based on the Unfiltered worksheet row count.
            If LastRow_CRJUnfiltered > 2 Then
                wsCRJ.Range("B2:AG" & LastRow_CRJUnfiltered).FillDown
            End If

    ' ============================================================
    '                     FORMAT THE WORKSHEET
    ' ============================================================
        wsCRJ.Range("A1:AG1").AutoFilter
        wsCRJ.Columns("A:AG").AutoFit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' MOVE MISSING SOURCE FILES ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
MoveFiles:
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section is used to move any Donation Site Reports that are missing PMT-IDs into a dedicated 'To be processed later' folder.
            ' If the AllowRevenueAmountAdjustments switch is off, also move those files to the 'To be processed later' folder because that data will be excluded from the Import Files.
        ' The purpose of moving the files, allows the user to know exactly which files still need to be processed later, when the transactions are synced into the system.
        ' It helps the user not need to look through tens, hundreds, or even thousands of files to determine exactly which disbursements were not matched and processed.
    
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "MOVING 'TO BE PROCESSED LATER' FILES TO A NEW FOLDER"
    
    ' ============================================================
    '        BUILD A UNIQUE LIST OF FILES THAT MUST BE MOVED
    ' ============================================================
        ' Use a dictionary so each file name is stored only once.
            ' A single source file may appear multiple times in the User-Required Adjustments worksheet.
                ' If the converter tried to move the file twice, it would fail after the first move.
        ' The dictionary prevents duplicates and gives the converter one clean list of file names to process.
            Set dictFilesToMove = CreateObject("Scripting.Dictionary")

        ' ---------------------------------------------
        '               ADD FILES FROM THE
        '  "TRANSACTIONS WITH MISSING PMT-IDS" SECTION
        ' ---------------------------------------------
            ' These files should always be moved because missing PMT-ID transactions are unresolved and should be reviewed outside the normal processing flow.
                For UserRequiredAdjustmentsRow = DataStartRow_UserRequiredAdjustments_MissingPaymentIDs To LastRow_UserRequiredAdjustments_MissingPaymentIDs
                    
                ' Pull the file name from column N of the User-Required Adjustments worksheet.
                    AdditionalFileToMove = Trim(CStr(wsUserRequiredAdjustments.Range("N" & UserRequiredAdjustmentsRow).Value))
                    
                ' Ignore blanks to store only real file names.
                    If AdditionalFileToMove <> "" Then
                        
                    ' Add the file name only if it is not already in the dictionary.
                        If Not dictFilesToMove.Exists(AdditionalFileToMove) Then
                            dictFilesToMove.Add AdditionalFileToMove, AdditionalFileToMove
                        End If
                        
                    End If
                    
                Next UserRequiredAdjustmentsRow

        ' ---------------------------------------------
        '               ADD FILES FROM THE
        '       "GROSS AMOUNT MISMATCHES" SECTION
        ' ---------------------------------------------
            ' If the AllowRevenueAmountAdjustments switch is Flase, then gross amount mismatches remain unresolved exceptions and _
              their source files should also be moved to the Process Later folder.
                If AllowRevenueAmountAdjustments = False Then
                    
                    For UserRequiredAdjustmentsRow = DataStartRow_UserRequiredAdjustments_GrossAmountVariances To LastRow_UserRequiredAdjustments_GrossAmountVariances
                        
                    ' Pull the file name from column N of the User-Required Adjustments worksheet.
                        AdditionalFileToMove = Trim(CStr(wsUserRequiredAdjustments.Range("N" & UserRequiredAdjustmentsRow).Value))
                        
                    ' Ignore blanks to store only real file names.
                        If AdditionalFileToMove <> "" Then
                            
                        ' Add the file name only if it is not already in the dictionary.
                            If Not dictFilesToMove.Exists(AdditionalFileToMove) Then
                                dictFilesToMove.Add AdditionalFileToMove, AdditionalFileToMove
                            End If
                            
                        End If
                        
                    Next UserRequiredAdjustmentsRow
                    
                End If

    ' ============================================================
    '         CONVERT THE UNIQUE FILE LIST INTO A VBA ARRAY
    ' ============================================================
        ' Convert the dictionary keys into an array so the file-move loop can loop through the final unique file list in a simple, controlled way.
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
    '               CREATE THE "PROCESS LATER" FOLDER
    ' ============================================================
        ' This folder allows the user to easily determine which files still need to be processed later.
        ' Create a timestamped folder inside the original Donation Site folder.
          ' Using a timestamp keeps each run distinct and avoids collisions with prior folders.
            If dictFilesToMove.Count > 0 Then
                FolderPath_ProcessLater = FolderPath_DonationSite & "\Process Later - " & Format(Now, "yyyy.mm.dd_hh.mm.ss")
                MkDir FolderPath_ProcessLater
            End If
            
    ' ============================================================
    '              MOVE THE FILES INTO THE NEW FOLDER
    ' ============================================================
        ' This folder allows the user to easily determine which files still need to be processed later.
          ' Move each unique file from the original Donation Site folder into the new 'Process Later' folder.
            If dictFilesToMove.Count > 0 Then
                For FileIndex = LBound(FilesToMove) To UBound(FilesToMove)
                    
                    SourceFilePath = FolderPath_DonationSite & "\" & FilesToMove(FileIndex)
                    DestinationFilePath = FolderPath_ProcessLater & "\" & FilesToMove(FileIndex)
                    
                ' Only move the file if it still exists in the source folder.
                    If Dir(SourceFilePath) <> "" Then
                        Name SourceFilePath As DestinationFilePath
                    End If
                    
                Next FileIndex
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FINALIZE WORKBOOK FORMATTING '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This section cleans up the workbook to display only to worksheets needed for documentation and centers the focus to the right spot for the user.
    
    ' ============================================================
    '             ACTIVATE THE COMPLETE RESET WORKSHEET
    ' ============================================================
        ' Bring the focus of the workbook back to the beginning and away from the Donation Site Reports.
            wbMacro.Worksheets("COMPLETE RESET").Activate
    
    ' ============================================================
    '                        PROTECT WORKSHEETS
    ' ============================================================
        ' The following worksheets are protected to prevent tampering by the user, of delicate formulas within worksheets that remain visible to the user.
            wsStandardizedSF.Protect
            wsStandardizedDonationSiteData.Protect
            wsUserRequiredAdjustments.Protect
            
    ' ============================================================
    '                        HIDE WORKSHEETS
    ' ============================================================
        ' Hide all the worksheets that are not required to be shown for documentation purposes.
            ' ---------------------------------------------
            '         HIDE THE FOLLOWING WORKSHEETS
            ' ---------------------------------------------
                wsSchoolValidation.Visible = xlSheetHidden
                wsInitialData.Visible = xlSheetHidden
                wsConsolidatedData.Visible = xlSheetHidden
                wsDisbursementData.Visible = xlSheetHidden
                wsRelevantTransactions.Visible = xlSheetHidden
                wsFees.Visible = xlSheetHidden
                wsBankDeposits.Visible = xlSheetHidden
                wsConnectionAnalysis.Visible = xlSheetHidden
                
                
                If JournalType = "Adjusting" Then
                ' Adjusting Journal path worksheets
                    wsAdjustingUnfiltered.Visible = xlSheetHidden
                    wsAdjustingFiltered.Visible = xlSheetHidden
                Else
                ' CRJ path worksheets
                    wsCRJUnfiltered.Visible = xlSheetHidden
                    wsCRJFiltered.Visible = xlSheetHidden
                End If
            
            ' ---------------------------------------------
            '     KEEP THE FOLLOWING WORKSHEETS VISIBLE
            ' ---------------------------------------------
                wsStandardizedSF.Visible = xlSheetVisible
                wsStandardizedDonationSiteData.Visible = xlSheetVisible
                wsUserRequiredAdjustments.Visible = xlSheetVisible

                If JournalType = "Adjusting" Then
                ' Adjusting Journal path worksheets
                    wsAdjustingJournal.Visible = xlSheetVisible
                    wsAdjustingJournal.Activate
                Else
                ' CRJ path worksheets
                    wsCRJ.Visible = xlSheetVisible
                    wsCRJ.Activate
                End If
    
    ' ============================================================
    '                  CHANGE WORKSHEET TAB COLORS
    ' ============================================================
        ' Provide a clean signal to the user for what is left to for them in the process.
            ' ---------------------------------------------
            '        DETERMINE THE VISIBILITY OF THE
            '      wsUserRequiredAdjustments WORKSHEET
            ' ---------------------------------------------
                ' If no adjustments are required in the User Required Adjustments worksheet:
                    If OK_UserRequiredAdjustments = 5 Then
                        wsUserRequiredAdjustments.Tab.Color = vbGreen
                        wsUserRequiredAdjustments.Visible = xlSheetHidden
                        
                        If JournalType = "Adjusting" Then
                            wsAdjustingJournal.Tab.Color = vbGreen
                        Else
                            wsCRJ.Tab.Color = vbGreen
                        End If
                        
                ' If the Missing PMT-IDs section is the only adjustment required in the User Required Adjustments worksheet:
                  ' (The user cannot make any adjustments with this section, in the converter, so the Intacct Import File is ready to be used.)
                    ElseIf OK_UserRequiredAdjustments = 4 And PaymentIDs_Missing = True Then
                        wsUserRequiredAdjustments.Tab.Color = vbYellow
                        wsUserRequiredAdjustments.Activate
                        
                        If JournalType = "Adjusting" Then
                            wsAdjustingJournal.Tab.Color = vbGreen
                        Else
                            wsCRJ.Tab.Color = vbGreen
                        End If
                        
                ' If one or more adjustments (not including Missing PMT-IDs) are required in the User Required Adjustments worksheet:
                    Else
                        wsUserRequiredAdjustments.Tab.Color = vbRed
                        wsUserRequiredAdjustments.Activate
                        
                        If JournalType = "Adjusting" Then
                            wsAdjustingJournal.Tab.Color = vbYellow
                        Else
                            wsCRJ.Tab.Color = vbYellow
                        End If
                    End If
                    
            ' ---------------------------------------------
            '    IF NO RELEVANT TRANSACTION DATA IS FOUND
            ' ---------------------------------------------
                ' If no transactions are connected. Then the Finalized Import File worksheet, will appear red indicating that it is not to be used.
                    If IsError(wsRelevantTransactions.Range("B2").Value) Then
                        If JournalType = "Adjusting" Then
                            wsAdjustingJournal.Tab.Color = vbRed
                        Else
                            wsCRJ.Tab.Color = vbRed
                        End If
                    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CREATE A FINAL MESSAGE TO THE USER ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' Create a final message to help the user to see the converter has completed and give them the statistics on what was processed.
        
    ' ============================================================
    '                FINAL COMPLETION MESSAGE
    ' ============================================================
        ExtraMessage = "Thank you for your patience! The converter has successfully completed." & vbCrLf & vbCrLf & _
                        FilesCount_Total_Message & vbCrLf & _
                        FileCount_DonationSite_Used_Message & vbCrLf & _
                        FileCount_DonationSite_WrongFileType_Message & vbCrLf & _
                        FileCount_DonationSite_WrongReport_Message & vbCrLf & _
                        FileCount_DonationSite_Unusable_Message

                        
        ExtraMessage_Title = "Converter Completed Successfully"

    ' ============================================================
    '                       END THE CONVERTER
    ' ============================================================
        ' Jump to the CompleteMacro ending.
            GoTo CompleteMacro

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' CREATE BUTTON FOR DONATION SITE REPORTS '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-----------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CreateButton_Step2:
    ' ============================================================
    '                    PURPOSE OF THIS SECTION
    ' ============================================================
        ' This worksheet is used as a holding/instruction page when the required Donation Site Reports have not yet been added to the workbook.
        ' The converter depends on the Donation Site Report data to create the Intacct Import file.
        ' If that data is missing, the converter provides a clear place for the user to pick up where they left off, without requiring them to completely restart the process.

    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "CREATING A WORKSHEET FOR WHEN THE USER IS READY TO IMPORT DONATION SITE DATA"

    ' ============================================================
    '           CHECK WHETHER THE BUTTON WORKSHEET EXISTS
    ' ============================================================
        ' Before creating the worksheet, first check whether it already exists so the converter does not try to create a duplicate worksheet.
            wsFound = False
            
            For Each wsButton In wbMacro.Worksheets
                If wsButton.Name = "No Donation Site Report" Then
                    wsFound = True
                    Exit For
                End If
            Next wsButton

    ' ============================================================
    '      CREATE THE BUTTON WORKSHEET (IF IT DOES NOT EXIST)
    ' ============================================================
        If wsFound = False Then
            
            ' ----------------------------------------------------
            '                 CREATE THE WORKSHEET
            ' ----------------------------------------------------
                    Set wsButton = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets("COMPLETE RESET"))
                    wsButton.Name = "No Donation Site Report"
                
            ' ----------------------------------------------------
            '               FORMAT THE WORKSHEET
            ' ----------------------------------------------------
                ' Create a black background to help the button stand out more to the user.
                    wsButton.Cells.Interior.Color = vbBlack
                
            ' ----------------------------------------------------
            '                  CREATE THE BUTTON
            ' ----------------------------------------------------
                ' This gives the user a simple start point to begin where they left off. It is simplified to help the user easily identify what the next steps are.
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
        ' At this stage, the converter is intentionally pausing for missing input.
        ' Hiding the Initial Data worksheet helps reduce confusion and make it easier to understand the next required action is to add the Donation Site reports.
            On Error Resume Next
                wsInitialData.Visible = xlSheetHidden
            On Error GoTo 0

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' RESTORE EXCEL ENVIRONMENT ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''---------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CompleteMacro:
    ' ============================================================
    '                     UPDATE THE STATUS BAR
    ' ============================================================
        Application.StatusBar = "COMPLETING THE CONVERTER PROCESS"
    
    ' ============================================================
    '            RESTORE THE EXCEL APPLICATION SETTINGS
    ' ============================================================
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' PROVIDE MESSAGE TO USER '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    '              DISPLAY THE FINAL USER MESSAGE
    ' ============================================================
        ' Provide a completion message to the user, to alert them the converter has completed successfully or to help them identify what pieces are missing from their input.
            MsgBox ExtraMessage, _
                   vbOKOnly, _
                   ExtraMessage_Title

End Sub
