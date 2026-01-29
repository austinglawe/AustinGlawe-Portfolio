Sub NTXGiving_SF()
' ==============================================================
' MODULE: NTXGiving_SF
' AUTHOR: Austin Glawe
' CREATED: 2025.04.01
' LAST UPDATED: 2026.01.28
' CURRENT MAINTAINER: See configuration section (CurrentVBACodeMaintainer)
' ==============================================================

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' PURPOSE, REQUIREMENTS, AND UPDATES ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ==============================================================
    ' PURPOSE
    ' ==============================================================
        ' The purpose of this macro is to:
            ' Transform the North Texas Giving Day (NTX Giving) report into a Salesforce-ready import file.
            ' Standardize and enhance the original NTX Giving report by removing non-essential columns and adding required fields pulled from the NTX Giving website.
            ' Provide the AR Team with all required data to create the Sage Intacct Accounting System import file.
    
    ' ==============================================================
    ' REQUIREMENTS
    ' ==============================================================
        ' NTX Giving disbursement report exported from: https://www.northtexasgivingday.org/organization/Basistexas/disbursements
        ' The Disbursement ID displayed on each NTX Giving disbursement details page.
        ' Any adjustments or fee reimbursements listed on the NTX Giving details pages.
    
    ' ==============================================================
    ' UPDATE LOG (LAST UPDATED: 2026.01.28)
    ' ==============================================================
        ' Original Production Rollout Date: 2025.04.01

        ' Updates:
            ' 2026.01.28 - Initiated the update log.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' CONFIGURATIONS AND VARIABLES ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
    ' ==============================================================
    ' CONFIGURATIONS
    ' ==============================================================
        ' --------------------------------------------------------------
        ' DECLARE CONFIGURATIONS
        ' --------------------------------------------------------------
            Dim wbMacro As Workbook
            
            Dim DonationSite As String
            
            Dim NTXGiving_Headers_Part1 As Variant
            Dim NTXGiving_Headers_Part1b As Variant
            Dim NTXGiving_Headers_Part2 As Variant
            Dim NTXGiving_Headers_AdvisoryOptional As Variant
            Dim NTXGiving_Headers_Part3 As Variant
            Dim NTXGiving_Headers_Option1 As Variant
            Dim NTXGiving_Headers_Option2 As Variant
            Dim NTXGiving_Headers_Option3 As Variant
            Dim NTXGiving_Headers_Option4 As Variant

            'Dim ExpectedMaxLastColumn As Long
            
            Dim HeaderRow As Long
            Dim AllowHeaderRowSearch As Boolean
            
            Dim ReportTrailingRowsToExclude As Long
            
            Dim AllowReportsPullIn As Boolean
            Dim RunConsolidationOnly As Boolean
            
            Dim CurrentVBACodeMaintainer As String
            
        ' --------------------------------------------------------------
        ' ASSIGN CONFIGURATIONS
        ' --------------------------------------------------------------
            ' Set the converter workbook to a variable
                Set wbMacro = ThisWorkbook
        
            ' The naming convention used when describing where the relevant data is coming from
            ' As of the most recent update, this is how the DonationSite name appears in Salesforce and Intacct.
                DonationSite = "NTXGiving"
                
            ' ...............................
            ' Possible Headers
            ' ...............................
                ' The original report was Part1, Part2, and Part3 combined for the column headers. As of 2025.12, "Entered By" became standard to split Part1 from Part2.
                ' In one case (Disbursement: 935039) the '_AdvisoryOptional' column headers were used to split NTXGiving_Headers_{Part2} from ...{Part3}
                ' These are intentionally broken out to show the variations between reports; however the final report will delete the '_Part1b' and '_AdvisoryOptional' portions _
                    to standardize the process and retain all necessary headers
                
                    ' Columns A through F (Consistent in all reports)
                        NTXGiving_Headers_Part1 = Array("Tracking #", "Type", "Date", "Time (Central)", "Donor First Name", "Donor Last Name")
                    
                    ' As of 2025.12 it became a new standard, to insert "Entered By" between what was originally NTXGiving_Headers_{Part1} and ...{Part2} at Column G
                    ' Currently: Column G
                        NTXGiving_Headers_Part1b = Array("Entered By")
                    
                    ' Originally: Columns G through BF;
                    ' Currently: Columns H through BG
                        NTXGiving_Headers_Part2 = Array("Address", "City", "State", "Zip", "Country", "Email", "Designation", "Dedication Type", "Dedication Name", "Send Dedication Email", _
                                "Dedication Email Address", "Dedication Email Message", "Source", "Page Creator", "Donation Site", "Team Name", "Giving Event Name", "Repeats", _
                                "Payment Method", "Origin", "Amount", "Transaction Fee Rate", "Transaction Fee Cost", "Fees Covered", "Covered Cost", "Net Cost", "Net Amount", "Refund", _
                                "Disbursement Date", "Referral Code", "Fully Anonymous", "Publicly Hidden", "Hide Amount", "Associated Groups", "Fundraiser Tracking ID", _
                                "External Transaction Tracking ID", "Entered At (Central)", "Processing Entity Name", "Phone (ID7302)", "First time giver? (ID7637)", _
                                "Volunteer Interest (ID7638)", "Recognition Name (ID7639)", "Age Demographic (ID7851)", "OLD_Ethnicity/Race Demo (ID7852)", "Phone Number (ID7887)", _
                                "Recognition Name(s) (ID7888)", "First time donor? (ID7889)", "Interested in Volunteering", "Volunteer Hours Pledged", "Notes (ID7921)", _
                                "NEW_Race/Ethnicity Demo (ID8854)")
                    
                    ' Columns BF through BL in the one report it was originally in. If re-downloaded (including the "Entered By" column), it now appears in BG through BM
                        NTXGiving_Headers_AdvisoryOptional = Array("Advisory Institution", "Other Advisory Institution", "Advisor Name", "Advisor Email", "Advisor Phone", _
                        "Advised Payment Details", "Advised Gift Type")
                
                    ' Originally: Column BF; Column BM in the original report including '_AdvisoryOptional'
                    ' Currently: BG; BN if the report with '_AdvisoryOptional' is re-downloaded.
                        NTXGiving_Headers_Part3 = Array("Notes")
                        
                        
                    ' Define all supported NTX Giving header layouts (known report variants), accounting for the presence or absence of the 'Entered By' and 'AdvisoryOptional' sections.
                    ' Header values are intentionally duplicated here to preserve explicit column order and to make each supported report layout fully self-contained and easy to validate.
                        ' Option 1: No Entered By, No Advisory Optional (Part1, Part2, Part3)
                            NTXGiving_Headers_Option1 = Array("Tracking #", "Type", "Date", "Time (Central)", "Donor First Name", "Donor Last Name", _
                                "Address", "City", "State", "Zip", "Country", "Email", "Designation", "Dedication Type", "Dedication Name", "Send Dedication Email", _
                                "Dedication Email Address", "Dedication Email Message", "Source", "Page Creator", "Donation Site", "Team Name", "Giving Event Name", "Repeats", _
                                "Payment Method", "Origin", "Amount", "Transaction Fee Rate", "Transaction Fee Cost", "Fees Covered", "Covered Cost", "Net Cost", "Net Amount", "Refund", _
                                "Disbursement Date", "Referral Code", "Fully Anonymous", "Publicly Hidden", "Hide Amount", "Associated Groups", "Fundraiser Tracking ID", _
                                "External Transaction Tracking ID", "Entered At (Central)", "Processing Entity Name", "Phone (ID7302)", "First time giver? (ID7637)", _
                                "Volunteer Interest (ID7638)", "Recognition Name (ID7639)", "Age Demographic (ID7851)", "OLD_Ethnicity/Race Demo (ID7852)", "Phone Number (ID7887)", _
                                "Recognition Name(s) (ID7888)", "First time donor? (ID7889)", "Interested in Volunteering", "Volunteer Hours Pledged", "Notes (ID7921)", _
                                "NEW_Race/Ethnicity Demo (ID8854)", _
                                "Notes")
                            
                        ' Option 2: Entered By present, No Advisory Optional (Part1, Part1b, Part2, Part3)
                            NTXGiving_Headers_Option2 = Array("Tracking #", "Type", "Date", "Time (Central)", "Donor First Name", "Donor Last Name", _
                                "Entered By", _
                                "Address", "City", "State", "Zip", "Country", "Email", "Designation", "Dedication Type", "Dedication Name", "Send Dedication Email", _
                                "Dedication Email Address", "Dedication Email Message", "Source", "Page Creator", "Donation Site", "Team Name", "Giving Event Name", "Repeats", _
                                "Payment Method", "Origin", "Amount", "Transaction Fee Rate", "Transaction Fee Cost", "Fees Covered", "Covered Cost", "Net Cost", "Net Amount", "Refund", _
                                "Disbursement Date", "Referral Code", "Fully Anonymous", "Publicly Hidden", "Hide Amount", "Associated Groups", "Fundraiser Tracking ID", _
                                "External Transaction Tracking ID", "Entered At (Central)", "Processing Entity Name", "Phone (ID7302)", "First time giver? (ID7637)", _
                                "Volunteer Interest (ID7638)", "Recognition Name (ID7639)", "Age Demographic (ID7851)", "OLD_Ethnicity/Race Demo (ID7852)", "Phone Number (ID7887)", _
                                "Recognition Name(s) (ID7888)", "First time donor? (ID7889)", "Interested in Volunteering", "Volunteer Hours Pledged", "Notes (ID7921)", _
                                "NEW_Race/Ethnicity Demo (ID8854)", _
                                "Notes")
                                
                        ' Option 3: Entered By + Advisory Optional present (Part1, Part1b, Part2, AdvisoryOptional, Part3)
                            NTXGiving_Headers_Option3 = Array("Tracking #", "Type", "Date", "Time (Central)", "Donor First Name", "Donor Last Name", _
                                "Entered By", _
                                "Address", "City", "State", "Zip", "Country", "Email", "Designation", "Dedication Type", "Dedication Name", "Send Dedication Email", _
                                "Dedication Email Address", "Dedication Email Message", "Source", "Page Creator", "Donation Site", "Team Name", "Giving Event Name", "Repeats", _
                                "Payment Method", "Origin", "Amount", "Transaction Fee Rate", "Transaction Fee Cost", "Fees Covered", "Covered Cost", "Net Cost", "Net Amount", "Refund", _
                                "Disbursement Date", "Referral Code", "Fully Anonymous", "Publicly Hidden", "Hide Amount", "Associated Groups", "Fundraiser Tracking ID", _
                                "External Transaction Tracking ID", "Entered At (Central)", "Processing Entity Name", "Phone (ID7302)", "First time giver? (ID7637)", _
                                "Volunteer Interest (ID7638)", "Recognition Name (ID7639)", "Age Demographic (ID7851)", "OLD_Ethnicity/Race Demo (ID7852)", "Phone Number (ID7887)", _
                                "Recognition Name(s) (ID7888)", "First time donor? (ID7889)", "Interested in Volunteering", "Volunteer Hours Pledged", "Notes (ID7921)", _
                                "NEW_Race/Ethnicity Demo (ID8854)", _
                                "Advisory Institution", "Other Advisory Institution", "Advisor Name", "Advisor Email", "Advisor Phone", "Advised Payment Details", "Advised Gift Type", _
                                "Notes")
                                
                        ' Option 4: Advisory Optional present, no Entered By (Part1, Part2, AdvisoryOptional, Part3)
                            NTXGiving_Headers_Option4 = Array("Tracking #", "Type", "Date", "Time (Central)", "Donor First Name", "Donor Last Name", _
                                "Address", "City", "State", "Zip", "Country", "Email", "Designation", "Dedication Type", "Dedication Name", "Send Dedication Email", _
                                "Dedication Email Address", "Dedication Email Message", "Source", "Page Creator", "Donation Site", "Team Name", "Giving Event Name", "Repeats", _
                                "Payment Method", "Origin", "Amount", "Transaction Fee Rate", "Transaction Fee Cost", "Fees Covered", "Covered Cost", "Net Cost", "Net Amount", "Refund", _
                                "Disbursement Date", "Referral Code", "Fully Anonymous", "Publicly Hidden", "Hide Amount", "Associated Groups", "Fundraiser Tracking ID", _
                                "External Transaction Tracking ID", "Entered At (Central)", "Processing Entity Name", "Phone (ID7302)", "First time giver? (ID7637)", _
                                "Volunteer Interest (ID7638)", "Recognition Name (ID7639)", "Age Demographic (ID7851)", "OLD_Ethnicity/Race Demo (ID7852)", "Phone Number (ID7887)", _
                                "Recognition Name(s) (ID7888)", "First time donor? (ID7889)", "Interested in Volunteering", "Volunteer Hours Pledged", "Notes (ID7921)", _
                                "NEW_Race/Ethnicity Demo (ID8854)", _
                                "Advisory Institution", "Other Advisory Institution", "Advisor Name", "Advisor Email", "Advisor Phone", "Advised Payment Details", "Advised Gift Type", _
                                "Notes")
                
                        ' Defines the maximum expected column index for NTX Giving reports.
                            ' As of the most recent update, the furthest-right column is BN (column 66).
                            ' Option 1: 58 columns
                            ' Option 2: 59 columns
                            ' Option 3: 66 columns
                            ' Option 4: 65 columns
                                'ExpectedMaxLastColumn = 66
                                
            
            ' HeaderRow is set to row 1 unless the control is switched on, in which case, it will auto-search for the column header row.
            ' As of the most recent update, the HeaderRow is row 1, but if it changes in the future, this should be changed with it.
                HeaderRow = 1
                AllowHeaderRowSearch = False
            
            
            ' If the report has additional rows, like totals at the bottom of the report, this number will be used to remove those rows when consolidating the data
            ' As of the most recent update, the variable should be set to '0', as all rows should be included when consolidating the reports.
                ReportTrailingRowsToExclude = 0
            
            
            ' This control is used to allow or disallow duplication of the reports as their own worksheet.
            ' Typically this should be set to False because the individual worksheets are not needed for documentation of the Salesforce Import Process.
                AllowReportsPullIn = False
            
            
'            ' This control is used to allow the user to terminate the process after the consolidation of data is completed instead of creating the Salesforce Import File.
'            ' Typically this is set to False and is only used when the user needs to consolidate files.
'                RunConsolidationOnly = False

            
            ' Used in error messages to direct the user to the current code maintainer if unexpected issues occur.
                CurrentVBACodeMaintainer = "Austin Glawe"
                    
    ' ==============================================================
    ' VARIABLES
    ' ==============================================================
        ' --------------------------------------------------------------
        ' DECLARE VARIABLES
        ' --------------------------------------------------------------
            Dim UserResponse As VbMsgBoxResult
            Dim fdFilePath As FileDialog
            Dim ExitMessage As String
            Dim ExitMessage_Title As String
            Dim DonationSiteReportFilePath As String
            
            Dim wbTemp As Workbook
            Dim wsTemp As Worksheet
            Dim TempLastRow As Long
            
            Dim PotentialHeaderRow As Long
            Dim TempHeadersLastCol As Long
            Dim HeaderLayoutOption As Long
            Dim CheckHeaderLayout As Variant
            Dim TempHeaders As Variant
            Dim TempHeadersLookup() As Variant
            Dim TempHeaderIndex As Long
            Dim HeaderIndex As Long
            Dim HeaderLayout As String
            
            Dim DisbursementID As String
            Dim FeesReimbursed As VbMsgBoxResult
            Dim PrizeAwarded As VbMsgBoxResult
            Dim PrizeAmountRaw As Variant
            Dim PrizeAmount As Double
            Dim PrizeDescription As String
            
            Dim wsDonationSiteReport As Worksheet
            Dim wsSFImport As Worksheet
            Dim wsSchoolValidation As Worksheet
            Dim wsDonationSiteReportRaw As Worksheet
            
            Dim FileNameNew As String
            Dim LastSeparatorIndex As Long
            Dim FolderPathOld As String
            Dim CurrentDateTime As String
            Dim FolderPathNew As String
            
            Dim ValidationColumnSearch As Range
            Dim ValidationSFSchoolNameCol As Long
            Dim ValidationNTXGivingSchoolNameCol As Long
            Dim ValidationSchoolAbbrevCol As Long
            Dim ValidationLastRow As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' PRE-RUN CHECKLIST AND CONFIRMATION ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Display a pre-run checklist outlining all required information the user must have available before starting the converter.
        UserResponse = MsgBox( _
                "Before starting, please confirm you have the following:" & vbCrLf & vbCrLf & _
                    "1. A report downloaded from the North Texas Giving donation site." & vbCrLf & _
                    "2. The Disbursement ID from the individual disbursement page." & vbCrLf & _
                    "3. Whether the transaction fees were reimbursed." & vbCrLf & _
                    "4. Whether an additional prize was awarded by North Texas Giving and, if so, the prize description name for it." & vbCrLf & vbCrLf & _
                "Are you ready to continue?", _
                vbYesNo + vbQuestion, _
                "North Texas Giving – SF Converter Confirmation")

    ' If the user responds by not being ready, end the macro immediately.
        If UserResponse = vbNo Then Exit Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' CREATE A RESET PAGE ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Using the 'Reset.Create_Reset_Worksheet' macro, clear the entire workbook to set up a clean work environment.
        Reset.Create_Reset_Worksheet
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-----------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' TURN OFF ALERTS '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-----------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-----------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' OPTIONAL: CONSOLIDATE RAW REPORTS '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-----------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' If enabled, skip the Salesforce import steps and run consolidation only (single combined Raw Reports worksheet).
'        If RunConsolidationOnly = True Then
'            GoTo ConsolidationOnly
'        End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''----------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' FILE SELECTION ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''----------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ==============================================================
    ' USER FILE SELECTION
    ' ==============================================================
        ' Prompt the user to select the NTX Giving disbursement report file for validation and processing.
        ' Using 'FilePicker' without 'MultiSelect' is solely due to a disbursement ID manually needing to be added by the user later on, making it most reasonable to only process one file at a time.
        ' As of the most recent update, files are disbursed as .csv files, but this filter allows all common Excel file types.
        ' If no file is selected, the user is notified and the macro terminates.
        ' If a file is selected, its path is stored for use later in the conversion process.
            Set fdFilePath = Application.FileDialog(msoFileDialogFilePicker)
            
            With fdFilePath
                .Title = "Select the 'North Texas Giving' Report"
                .AllowMultiSelect = False
                
                .Filters.Clear
                .Filters.Add "Excel Files", "*.xlsx; *.xls; *.csv"
    
                If .Show <> -1 Then
                    ExitMessage = "No file selected. Please locate the 'North Texas Giving' Report and try again."
                    ExitMessage_Title = "No File Selected"
                    
                    GoTo CompleteMacro
                End If
                    
                DonationSiteReportFilePath = .SelectedItems(1)
            End With
    
    ' ==============================================================
    ' OPEN SELECTED FILE
    ' ==============================================================
        ' Open the user-selected file and store the workbook in a variable for processing and later closure.
        ' As of the most recent update, all relevant data resides on the first worksheet, which is stored in a variable for ease of use.
            Set wbTemp = Workbooks.Open(DonationSiteReportFilePath, ReadOnly:=True)
            Set wsTemp = wbTemp.Worksheets(1)
        
    ' ==============================================================
    ' EVALUATE FILE FOR CORRECT FORMATTING
    ' ==============================================================
' Update the Status Bar
    Application.StatusBar = "Evaluating File to Determine if it is a NTX Giving Report"
    
        ' Determine the last populated row in the worksheet to establish the full data range.
        ' If 'AllowHeaderRowSearch' is enabled, this serves as the last row to evaluate.
        ' Column 1 (A) is the most reliable column to always have data.
            TempLastRow = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
        
        ' Loop through each column header to see if it matches the expected NTX Giving column headers, to determine if it is the correct report.
            ' If the 'AllowHeaderRowSearch' is enabled, allow it to search through all rows for the column headers row.
            ' Otherwise, make sure the 'HeaderRow' variable is greater than 0 and less than the 'TempLastRow' row. Then evaluate the row.
            
        HeaderLayout = vbNullString
        ' ...............................
        ' 'AllowHeaderRowSearch': ON
        ' ...............................
            ' If 'AllowHeaderRowSearch' is enabled, scan each row (1..TempLastRow) to find where the headers begin.
            ' This supports report variations where the header row is no longer on row 1
                If AllowHeaderRowSearch = True Then
                
                    For PotentialHeaderRow = 1 To TempLastRow
                    
                    ' Reset per row so a previous row's detected option cannot carry forward.
                        HeaderLayoutOption = 0
                    
                    ' Determine the last populated column on the potential header row.
                        TempHeadersLastCol = wsTemp.Cells(PotentialHeaderRow, wsTemp.Columns.Count).End(xlToLeft).Column
                        
                    ' Select the expected layout to validate against based on column count.
                    ' If the column count does not match any supported layout, skip this row immediately.
                        If TempHeadersLastCol > UBound(NTXGiving_Headers_Option3) + 1 Then
                            GoTo NextRow
                        ElseIf TempHeadersLastCol = UBound(NTXGiving_Headers_Option3) + 1 Then
                            HeaderLayoutOption = 3
                        ElseIf TempHeadersLastCol = UBound(NTXGiving_Headers_Option4) + 1 Then
                            HeaderLayoutOption = 4
                        ElseIf TempHeadersLastCol = UBound(NTXGiving_Headers_Option2) + 1 Then
                            HeaderLayoutOption = 2
                        ElseIf TempHeadersLastCol = UBound(NTXGiving_Headers_Option1) + 1 Then
                            HeaderLayoutOption = 1
                        Else
                            GoTo NextRow
                        End If
                    
                    ' Load the expected header array for the selected layout option.
                        Select Case HeaderLayoutOption
                        Case 1
                            CheckHeaderLayout = NTXGiving_Headers_Option1
                        Case 2
                            CheckHeaderLayout = NTXGiving_Headers_Option2
                        Case 3
                            CheckHeaderLayout = NTXGiving_Headers_Option3
                        Case 4
                            CheckHeaderLayout = NTXGiving_Headers_Option4
                        End Select
                        
                    
                    ' Pull the candidate header row values into memory for faster comparisons.
                    ' Range.Value returns a 2-D array (1 row x N columns), which is then flattened into a 0-based 1-D array so the indexes align with our 0-based Option arrays.
                        TempHeaders = wsTemp.Range(wsTemp.Cells(PotentialHeaderRow, 1), wsTemp.Cells(PotentialHeaderRow, TempHeadersLastCol)).Value

                        ReDim TempHeadersLookup(0 To TempHeadersLastCol - 1)

                        For TempHeaderIndex = 1 To TempHeadersLastCol
                            TempHeadersLookup(TempHeaderIndex - 1) = Trim$(CStr(TempHeaders(1, TempHeaderIndex)))
                        Next TempHeaderIndex
                        
                    ' Compare every expected header to the actual header text in the same position.
                    ' Trim + case-insensitive compare reduces false mismatches caused by minor formatting differences.
                        For HeaderIndex = 0 To UBound(CheckHeaderLayout)
                            If StrComp(CheckHeaderLayout(HeaderIndex), TempHeadersLookup(HeaderIndex), vbTextCompare) <> 0 Then
                                GoTo NextRow
                            End If
                        Next HeaderIndex
                        
                    ' If we reach this point, the row is confirmed as the header row and the layout is identified.
                        HeaderRow = PotentialHeaderRow
                        HeaderLayout = "Option " & HeaderLayoutOption
                        Exit For
                    
NextRow:
                    Next PotentialHeaderRow
        ' ...............................
        ' 'AllowHeaderRowSearch': OFF
        ' ...............................
            ' When header search is disabled, validate only the configured HeaderRow (default is 1).
                Else
                ' If the configured HeaderRow is out of bounds, default to row 1.
                    If HeaderRow < 1 Or HeaderRow > TempLastRow Then HeaderRow = 1
                    
                ' Determine the last populated column on the configured header row to identify the layout.
                    TempHeadersLastCol = wsTemp.Cells(HeaderRow, wsTemp.Columns.Count).End(xlToLeft).Column

                ' If the column count does not match any supported layout, exit cleanly before continuing.
                    If TempHeadersLastCol > UBound(NTXGiving_Headers_Option3) + 1 Then
                        wbTemp.Close SaveChanges:=False
                        
                        ExitMessage = "The selected file does not match any supported NTX Giving report layout." & vbCrLf & vbCrLf & _
                                "If you confirmed you selected the correct report and this issue persists, please contact " & CurrentVBACodeMaintainer & " or the current VBA code maintainer."
                        ExitMessage_Title = "Invalid Report Format"
                        
                        GoTo CompleteMacro
                    ElseIf TempHeadersLastCol = UBound(NTXGiving_Headers_Option3) + 1 Then
                        HeaderLayoutOption = 3
                    ElseIf TempHeadersLastCol = UBound(NTXGiving_Headers_Option4) + 1 Then
                        HeaderLayoutOption = 4
                    ElseIf TempHeadersLastCol = UBound(NTXGiving_Headers_Option2) + 1 Then
                        HeaderLayoutOption = 2
                    ElseIf TempHeadersLastCol = UBound(NTXGiving_Headers_Option1) + 1 Then
                        HeaderLayoutOption = 1
                    Else
                        wbTemp.Close SaveChanges:=False
                        
                        ExitMessage = "The selected file does not match any supported NTX Giving report layout." & vbCrLf & vbCrLf & _
                                "If you confirmed you selected the correct report and this issue persists, please contact " & CurrentVBACodeMaintainer & " or the current VBA code maintainer."
                        ExitMessage_Title = "Invalid Report Format"
                        
                        GoTo CompleteMacro
                    End If
                    
                ' Load the expected header array for the identified layout option.
                    Select Case HeaderLayoutOption
                    Case 1
                        CheckHeaderLayout = NTXGiving_Headers_Option1
                    Case 2
                        CheckHeaderLayout = NTXGiving_Headers_Option2
                    Case 3
                        CheckHeaderLayout = NTXGiving_Headers_Option3
                    Case 4
                        CheckHeaderLayout = NTXGiving_Headers_Option4
                    End Select
                    
                ' Pull the configured header row into memory and flatten to a 0-based array for comparison.
                    TempHeaders = wsTemp.Range(wsTemp.Cells(HeaderRow, 1), wsTemp.Cells(HeaderRow, TempHeadersLastCol)).Value
                    
                    
                    ReDim TempHeadersLookup(0 To TempHeadersLastCol - 1)
                    
                    
                    For TempHeaderIndex = 1 To TempHeadersLastCol
                        TempHeadersLookup(TempHeaderIndex - 1) = Trim$(CStr(TempHeaders(1, TempHeaderIndex)))
                    Next TempHeaderIndex
                    
                ' Validate the headers in-place. If any header differs, exit cleanly to prevent mis-mapping.
                    For HeaderIndex = 0 To UBound(CheckHeaderLayout)
                        If StrComp(CheckHeaderLayout(HeaderIndex), TempHeadersLookup(HeaderIndex), vbTextCompare) <> 0 Then
                            wbTemp.Close SaveChanges:=False
                            
                            ExitMessage = "The selected file does not match any supported NTX Giving report layout." & vbCrLf & vbCrLf & _
                                    "If you confirmed you selected the correct report and this issue persists, please contact " & CurrentVBACodeMaintainer & " or the current VBA code maintainer."
                            ExitMessage_Title = "Invalid Report Format"
                            
                            GoTo CompleteMacro
                        End If
                    Next HeaderIndex
                    
                ' Mark the header layout as validated for downstream processing.
                    HeaderLayout = "Option " & HeaderLayoutOption
                End If
                    
                    
        ' If no layout was identified, terminate to prevent processing an unsupported report.
            If HeaderLayout = vbNullString Then
                wbTemp.Close SaveChanges:=False
                
                ExitMessage = "The selected file does not match any supported NTX Giving report layout." & vbCrLf & vbCrLf & _
                        "If you confirmed you selected the correct report and this issue persists, please contact " & CurrentVBACodeMaintainer & " or the current VBA code maintainer."
                ExitMessage_Title = "Invalid Report Format"
                
                GoTo CompleteMacro
            End If
        

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-----------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' GATHER REQUIRED DETAILS FROM USER '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-----------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Update the Status Bar
    Application.StatusBar = "Gathering Required Details From User"
    
    ' ==============================================================
    ' DISBURSEMENT ID
    ' ==============================================================
        ' Prompt for the Disbursement ID shown on the NTX Giving disbursement details page.
        ' This ID is used to standardize tracking across Salesforce and AR reconciliation. Later it will be added to the report.
        ' This detail can only be found on the donation site's Disbursement Details Page otherwise.
            Do
                DisbursementID = InputBox( _
                    "The file you selected was confirmed to be a 'North Texas Giving' File." & vbCrLf & vbCrLf & _
                    "Please enter the Disbursement ID from the 'North Texas Giving' Disbursement Page:", _
                    "Enter Disbursement ID")
    
            ' User canceled the operation
                If DisbursementID = vbNullString Then
                    wbTemp.Close SaveChanges:=False
                    
                    ExitMessage = "Operation Canceled. Disbursement ID not provided."
                    ExitMessage_Title = "Operation Canceled"
                    
                    GoTo CompleteMacro
                End If
                
            ' Normalize input
                DisbursementID = Trim$(DisbursementID)
                
            ' Validate minimum length
                If Len(DisbursementID) >= 5 Then
                    Exit Do
                Else
                    MsgBox _
                        "The Disbursement ID must be at least 5 characters." & vbCrLf & vbCrLf & _
                        "If you confirmed your input is the correct Disbursement ID and this issue persists, please contact " & CurrentVBACodeMaintainer & _
                        " or the current VBA code maintainer.", _
                        vbExclamation, "Invalid Disbursement ID"
                End If
            Loop

    ' ==============================================================
    ' FEE REIMBURSEMENT
    ' ==============================================================
        ' Ask the user if the fees were reimbursed
        ' If the fees were reimbursed, later this fee adjustment will be added to the report, to help the accounting team accurately determine the correct bank deposit amount.
        ' These details can only be found on the donation site's Disbursement Details Page otherwise.
            FeesReimbursed = MsgBox("Were the 'North Texas Giving' fee reimbursed for this disbursement?", _
                                    vbYesNo, _
                                    "Fees Reimbursed?")
            
    ' ==============================================================
    ' PRIZE AWARDED
    ' ==============================================================
        ' Ask the user if there were any prizes awarded
        ' If prizes were awarded, later this prize amount will be added to the report, to help the accounting team accurately determine the correct bank deposit amount.
        ' These details can only be found on the donation site's Disbursement Details Page otherwise.
            PrizeAwarded = MsgBox("Were any prizes awarded by 'North Texas Giving' for this disbursement?", _
                                    vbYesNo, _
                                    "Prizes Awarded?")
                                    
        ' If PrizeAwarded was "Yes" then collect the PrizeAmount and PrizeDescription from the user.
            If PrizeAwarded = vbYes Then
        ' --------------------------------------------------------------
        ' PRIZE AMOUNT
        ' --------------------------------------------------------------
            ' If the user confirms an award was given, get the 'PrizeAmount' and 'PrizeDescription'
                    Do
                        PrizeAmountRaw = InputBox("Please enter the prize amount", _
                                                "Enter Prize Amount")

                    ' If canceled or left blank, exit the sub
                        If PrizeAmountRaw = vbNullString Then
                            wbTemp.Close SaveChanges:=False
                            
                            ExitMessage = "Operation canceled. Prize amount not provided."
                            ExitMessage_Title = "Operation Canceled"
                            
                            GoTo CompleteMacro
                        End If
                
                
                        PrizeAmountRaw = Trim$(PrizeAmountRaw)

                    ' Blank entry (treat as invalid, keep looping)
                        If PrizeAmountRaw = vbNullString Then
                            MsgBox "Prize amount cannot be blank. Please enter a dollar amount or click Cancel.", _
                                   vbExclamation, _
                                   "Invalid Prize Amount"
                        Else
                        ' Remove currency formatting characters ($ and commas) so the value can be converted to a number.
                            PrizeAmountRaw = Replace(PrizeAmountRaw, "$", vbNullString)
                            PrizeAmountRaw = Replace(PrizeAmountRaw, ",", vbNullString)
                            
                            If IsNumeric(PrizeAmountRaw) Then
                                PrizeAmount = CDbl(PrizeAmountRaw)
                                Exit Do
                            Else
                                MsgBox "That does not look like a dollar amount. Please try again or click Cancel.", _
                                       vbExclamation, _
                                       "Invalid Prize Amount"
                            End If
                        End If
                    Loop
        ' --------------------------------------------------------------
        ' PRIZE DESCRIPTION
        ' --------------------------------------------------------------
            ' Get the 'PrizeDescription' from the user.
            ' This prize description will be used by the accounting team for the general ledger to help show the line item and where the money is coming from.
                Do
                    PrizeDescription = InputBox("Please enter the prize description." & vbCrLf & _
                                    "On the Disbursement Details page, this appears under the 'Adjustments' section in the column called 'Memo'.", _
                                    "Enter Prize Description")
            
                ' Validate the user gave some details on the Prize.
                    If PrizeDescription = vbNullString Then
                        wbTemp.Close SaveChanges:=False
                            
                        ExitMessage = "Operation Canceled. Prize Description not provided."
                        ExitMessage_Title = "Operation Canceled"
                        
                        GoTo CompleteMacro
                    End If
                
                ' Trim off any leading or trailing spaces
                    PrizeDescription = Trim$(PrizeDescription)
                
                ' Make sure the user gave at least 1 character or restart the loop.
                    If PrizeDescription <> vbNullString Then
                        Exit Do
                    Else
                        MsgBox "Prize description cannot be blank. Please enter the Prize Description or click Cancel.", _
                                vbExclamation, _
                                "Invalid Prize Description"
                    End If
                Loop
            End If
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''---------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' POPULATE THE DATA: PART 1 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''---------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Update the Status Bar
    Application.StatusBar = "Starting to Populate the Worksheets"
    
    ' ==============================================================
    ' CREATE ALL REQUIRED WORKSHEETS
    ' ==============================================================
        ' --------------------------------------------------------------
        ' DONATION SITE WORKSHEET
        ' --------------------------------------------------------------
            Set wsDonationSiteReport = wbMacro.Worksheets.Add(After:=wbMacro.Worksheets(wbMacro.Worksheets.Count))
            wsDonationSiteReport.Name = "Donation Site Report"
            
        ' --------------------------------------------------------------
        ' SALESFORCE IMPORT FILE WORKSHEET
        ' --------------------------------------------------------------
            Set wsSFImport = wbMacro.Worksheets.Add(After:=wsDonationSiteReport)
            wsSFImport.Name = "Salesforce Import File"
            
        ' --------------------------------------------------------------
        ' DATA VALIDATION WORKSHEET - SCHOOLS
        ' --------------------------------------------------------------
            Set wsSchoolValidation = wbMacro.Worksheets.Add(After:=wsSFImport)
            wsSchoolValidation.Name = "School Validation"
            
        ' --------------------------------------------------------------
        ' OPTIONAL: AllowReportPullIn: ON
        ' --------------------------------------------------------------
            ' CREATE THE RAW DONATION SITE REPORT WORKSHEET
            ' ADD THE ORIGINAL DONATION SITE REPORT (UNALTERED)
                If AllowReportsPullIn = True Then
                    Set wsDonationSiteReportRaw = wbMacro.Worksheets.Add(After:=wsSchoolValidation)
                    wsDonationSiteReportRaw.Name = "NTX Giving (" & DisbursementID & ")"
                    
                    wsTemp.Cells.Copy wsDonationSiteReportRaw.Range("A1")
                    
                    Application.CutCopyMode = False
                End If
        
    ' ==============================================================
    ' ENHANCE THE ORIGINAL DISBURSEMENT REPORT
    ' ==============================================================
        ' --------------------------------------------------------------
        ' DELETE UNNECESSARY COLUMNS
        ' --------------------------------------------------------------
            ' Delete the 'NTXGiving_Headers_AdvisoryOptional' columns:
                ' "Advisory Institution", "Other Advisory Institution", "Advisor Name", "Advisor Email", "Advisor Phone", "Advised Payment Details", "Advised Gift Type"
            ' Option 3: Columns 59:65 (BG:BM)
            ' Option 4: Columns 58:64 (BF:BL)
                If HeaderLayout = "Option 3" Then
                    wsTemp.Columns("BG:BM").Delete
                ElseIf HeaderLayout = "Option 4" Then
                    wsTemp.Columns("BF:BL").Delete
                End If

            ' Delete the 'NTXGiving_Headers_Part1b' column -- "Entered By"
            ' Option 2 or Option 3: Column 7 (Column G)
                If HeaderLayout = "Option 2" Or HeaderLayout = "Option 3" Then
                    wsTemp.Columns(7).Delete
                End If
                
            ' Delete any rows above the HeaderRow
                If HeaderRow > 1 Then
                    wsTemp.Rows("1:" & HeaderRow - 1).Delete
                    TempLastRow = TempLastRow - (HeaderRow - 1)
                    HeaderRow = 1
                End If
            
            
            ' Delete Any rows from the bottom that are required to be deleted
                If ReportTrailingRowsToExclude <> 0 Then
                    If ReportTrailingRowsToExclude < 0 Or ReportTrailingRowsToExclude > TempLastRow Then
                        ReportTrailingRowsToExclude = 0
                    ElseIf ReportTrailingRowsToExclude > 0 Then
                        wsTemp.Rows((TempLastRow - ReportTrailingRowsToExclude + 1) & ":" & TempLastRow).Delete
                        TempLastRow = TempLastRow - ReportTrailingRowsToExclude
                    End If
                End If
                
                
        ' --------------------------------------------------------------
        ' ADD IN ADDITIONAL DETAILS
        ' --------------------------------------------------------------
            ' ...............................
            ' FEE REIMBURSEMENT DETAILS
            ' ...............................
                ' If NTX Giving reimbursed all transaction fees for this disbursement:
                    ' Append a new synthetic line item to the report to represent the total fee reimbursement as a single positive adjustment.
                    ' Amount equals the sum of all prior Net Cost values for this disbursement
                    ' User confirmation is required, as it does not appear on the original donation site report. It only appears in the disbursement details page.
                    ' Used by Accounting to correctly reflect total funds received.
    
                        If FeesReimbursed = vbYes Then
                        ' Update the 'TempLastRow' varaible to add the FeesReimbursed line item
                            TempLastRow = TempLastRow + 1
                        
                        ' Add in all the details for the new line item
                            ' "Tracking #"
                                wsTemp.Range("A" & TempLastRow).Value = DisbursementID & " - North Texas Giving Day Transaction Processing Fee Reimbursement"
                                
                            ' "Date"
                                wsTemp.Range("C" & TempLastRow).Value = wsTemp.Range("AI2").Value
                                
                            ' "Source"
                                wsTemp.Range("S" & TempLastRow).Value = "BASIS Texas Charter Schools"
                                
                            ' "Amount" -- Sum AF through the last original data row (excluding the synthetic row we are adding now) to get the fee reimbursement amount.
                                wsTemp.Range("AA" & TempLastRow).Value = Application.Sum(wsTemp.Range("AF2:AF" & TempLastRow - 1))
                                
                            ' "Transaction Fee Cost"
                                wsTemp.Range("AC" & TempLastRow).Value = 0
                                
                            ' "Covered Cost"
                                wsTemp.Range("AE" & TempLastRow).Value = 0
                                
                            ' "Net Cost"
                                wsTemp.Range("AF" & TempLastRow).Value = 0
                                
                            ' "Net Amount"
                                wsTemp.Range("AG" & TempLastRow).Value = wsTemp.Range("AA" & TempLastRow).Value
                                
                            ' "Disbursement Date"
                                wsTemp.Range("AI" & TempLastRow).Value = wsTemp.Range("AI2").Value
                                
                            ' "Notes"
                                wsTemp.Range("BF" & TempLastRow).Value = "NTX Giving - Disbursement ID: " & DisbursementID & " | Adjustment: Fee Reimbursement"
                        End If
            
            ' ...............................
            ' PRIZE AWARDED DETAILS
            ' ...............................
                ' If NTX Giving awarded a prize for this disbursement:
                    ' Append a new synthetic line item to the report to represent the prize amount paid to BASIS.
                    ' Description and Amount are manually entered by the user, as it does not appear on the original donation site report. It only appears in the disbursement details page.
                    ' Used by Accounting to correctly reflect total funds received.
    
                        If PrizeAwarded = vbYes Then
                        ' Update the 'TempLastRow' varaible to add in the PrizeAwarded Details
                            TempLastRow = TempLastRow + 1
                            
                        ' Add in all the details for the new line item
                            ' "Tracking #"
                                wsTemp.Range("A" & TempLastRow).Value = DisbursementID & " - North Texas Giving Day Prize"
                                    
                            ' "Date"
                                wsTemp.Range("C" & TempLastRow).Value = wsTemp.Range("AI2").Value
                                
                            ' "Source"
                                wsTemp.Range("S" & TempLastRow).Value = "BASIS Texas Charter Schools"
                                
                            ' "Amount"
                                wsTemp.Range("AA" & TempLastRow).Value = PrizeAmount
                                
                            ' "Transaction Fee Cost"
                                wsTemp.Range("AC" & TempLastRow).Value = 0
                                
                            ' "Covered Cost"
                                wsTemp.Range("AE" & TempLastRow).Value = 0
                                
                            ' "Net Cost"
                                wsTemp.Range("AF" & TempLastRow).Value = 0
                                
                            ' "Net Amount"
                                wsTemp.Range("AG" & TempLastRow).Value = wsTemp.Range("AA" & TempLastRow).Value
                                
                            ' "Disbursement Date"
                                wsTemp.Range("AI" & TempLastRow).Value = wsTemp.Range("AI2").Value
                                
                            ' "Notes"
                                wsTemp.Range("BF" & TempLastRow).Value = "NTX Giving - Disbursement ID: " & DisbursementID & " | Adjustment: Prize | Prize Memo: " & PrizeDescription
                        End If

            ' ...............................
            ' DISBURSEMENT ID
            ' ...............................
                ' Add a standardized Disbursement ID column to the donation site report.
                ' Disbursement ID is manually entered by the user, as it does not appear on the original donation site report.
                ' This ID helps accounting group entries together later, to correctly reflect total funds received.
                ' Supports reconciliation between Salesforce, Sage Intacct, and bank statements
                
                    ' Insert a column in front of column A, to add the Disbursement ID to the donation site report.
                        wsTemp.Columns(1).Insert Shift:=xlToRight
                    
                    ' Add the column header to the newly inserted column
                        wsTemp.Range("A1").Value = "Disbursement ID"
                    
                    ' Fill the rest of the line items with the Disbursement ID
                        wsTemp.Range("A2:A" & TempLastRow).Value = DisbursementID
                        


    ' ==============================================================
    ' POPULATE THE DONATION SITE WORKSHEET
    ' ==============================================================
        ' Copy the fully standardized donation site data into the converter workbook.
            wsTemp.Range("A1:BG" & TempLastRow).Copy wsDonationSiteReport.Range("A1")
            
            Application.CutCopyMode = False
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' SAVE THE ADJUSTED REPORT FILE '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Update the Status Bar
    Application.StatusBar = "Creating the Adjusted Report File"
    
    ' The goal of this section is to:
      ' 1. Preserve the original NTX Giving report (do not overwrite it).
      ' 2. Produce an "enhanced" version in a new timestamped folder to be used by the Accounting Team.
      ' 3. Save as CSV and then close the donatioin site report workbook cleanly.
    
    ' Ensure the donor-site workbook's data sheet is active before SaveAs to CSV.
    ' (CSV output is based on the active sheet.)
        wsTemp.Activate
    
    ' Build the enhanced file name using:
      ' Disbursement Date (from the report)
      ' Disbursement ID (manual entry)
      ' Total deposit-related sum for quick identification in the folder
        FileNameNew = "NTX Giving - " & Format(wsTemp.Range("AJ2").Value, "YYYY.MM.DD") & " (" & DisbursementID & ")_" & _
                    Application.Text(Application.Sum(wsTemp.Range("AH2:AH" & TempLastRow)), "$#,##0.00")
        
        
    ' Derive the source folder from the originally selected report path.
    ' The enhanced output folder is created alongside the original file for easy traceability.
        LastSeparatorIndex = InStrRev(DonationSiteReportFilePath, Application.PathSeparator)
           
        If LastSeparatorIndex > 0 Then
            FolderPathOld = Left(DonationSiteReportFilePath, LastSeparatorIndex)
        End If
    
    ' Create a new timestamped folder
        If FolderPathOld <> "" Then
        ' First: Find the Current date and time and put it into an variable
            CurrentDateTime = Format(Now, "YYYY.MM.DD-HH.MM.SS")
        ' Second: Create the New Folder Path variable with the appropriate naming
            FolderPathNew = FolderPathOld & "NTX Giving - for 'AR' Team - " & CurrentDateTime
        ' Third: Create the Folder
            MkDir FolderPathNew
        End If

    ' Save the enhanced version as a CSV into the new folder.
    ' This creates a new file in the destination and avoids overwriting the original download.
        wbTemp.SaveAs _
            FileName:=FolderPathNew & "\" & FileNameNew & ".csv", _
            FileFormat:=xlCSV, _
            CreateBackup:=False
    
    
    ' Close the temp workbook (the donation site report)
        wbTemp.Close SaveChanges:=False
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''---------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' POPULATE THE DATA: PART 2 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''---------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Update the Status Bar
    Application.StatusBar = "Finishing Populating the Data"

    ' ==============================================================
    ' POPULATE THE SCHOOL VALIDATION WORKSHEET
    ' ==============================================================
        ' Populate the School Validation worksheet with all school lookup data.
            School_Validation.Validation
        
        ' Retain only the three columns required for downstream school lookups:
          ' "SF School Name"
          ' "NTX Giving School Name"
          ' "School Abbreviation"
          
        ' The columns are temporarily staged, and all other columns are removed to keep the worksheet minimal and lookup-focused.
        ' The process is to:
          ' 1. Search for the column header
          ' 2. Store it in a variable
          ' 3. Copy that full column in a column after all data
          ' 4. Repeat for the second and third column headers
          ' 5. Delete all columns except for those 3
          ' Note: After the deletion, "SF School Name" should be in Column A, "NTX Giving School Name" in Column B, and "School Abbreviation" in Column C.
            ' This will matter during lookups and data validation later in the converter.
        ' --------------------------------------------------------------
        ' FIND, COPY, AND MOVE ALL 3 COLUMNS
        ' --------------------------------------------------------------
            ' ...............................
            ' "SF School Name"
            ' ...............................
                ' Search for the column header
                    Set ValidationColumnSearch = wsSchoolValidation.Rows(1).Find( _
                                                                        What:="SF School Name", _
                                                                        LookIn:=xlValues, _
                                                                        LookAt:=xlWhole, _
                                                                        MatchCase:=False)
                ' Store the column in a variable
                    ValidationSFSchoolNameCol = ValidationColumnSearch.Column
                
                ' Copy the full column after all data
                    wsSchoolValidation.Columns(ValidationSFSchoolNameCol).Copy wsSchoolValidation.Columns(200)
            
            ' ...............................
            ' "NTX Giving School Name"
            ' ...............................
                ' Search for the column header
                    Set ValidationColumnSearch = wsSchoolValidation.Rows(1).Find( _
                                                                        What:="NTX Giving School Name", _
                                                                        LookIn:=xlValues, _
                                                                        LookAt:=xlWhole, _
                                                                        MatchCase:=False)
                ' Store the column in a variable
                    ValidationNTXGivingSchoolNameCol = ValidationColumnSearch.Column
                
                ' Copy the full column after all data
                    wsSchoolValidation.Columns(ValidationNTXGivingSchoolNameCol).Copy wsSchoolValidation.Columns(201)
                    
            ' ...............................
            ' "School Abbreviation"
            ' ...............................
                ' Search for the column header
                    Set ValidationColumnSearch = wsSchoolValidation.Rows(1).Find( _
                                                                        What:="School Abbreviation", _
                                                                        LookIn:=xlValues, _
                                                                        LookAt:=xlWhole, _
                                                                        MatchCase:=False)
                ' Store the column in a variable
                    ValidationSchoolAbbrevCol = ValidationColumnSearch.Column
                
                ' Copy the full column after all data
                    wsSchoolValidation.Columns(ValidationSchoolAbbrevCol).Copy wsSchoolValidation.Columns(202)
            
        ' --------------------------------------------------------------
        ' DELETE ALL COLUMNS EXCEPT THE 3 JUST COPIED
        ' --------------------------------------------------------------
                wsSchoolValidation.Range(wsSchoolValidation.Columns(1), wsSchoolValidation.Columns(199)).Delete Shift:=xlToLeft
                
        ' --------------------------------------------------------------
        ' FIND THE LAST ROW OF THE 'wsSchoolValidation' WORKSHEET
        ' --------------------------------------------------------------
                ValidationLastRow = wsSchoolValidation.Cells(wsSchoolValidation.Rows.Count, 1).End(xlUp).Row

    ' ==============================================================
    ' POPULATE THE SALESFORCE IMPORT FILE WORKSHEET
    ' ==============================================================
        ' Add the column headers
            wsSFImport.Range("A1:R1").Value = Array("C&P Account Name Correction", "Matching Company Name", "Disbursement ID", "Payment Check/Reference Number", "Donation Date", _
                    "Deposit Date", "Donation Amount", "Contact1 First Name", "Contact1 Last Name", "Contact1 Work Email", "Donation Type", "Donation Stage", "C&P Payment Method", _
                    "Campaign Name", "Donation Name", "Donation Site", "Description", "Notes")
            
        ' "C&P Account Name Correction"
            wsSFImport.Range("A2").Formula2 = "=XLOOKUP(""*""&'Donation Site Report'!T2&""*"",'School Validation'!B:B,'School Validation'!A:A,,2)"
            
        ' "Matching Company Name" (N/A)
            wsSFImport.Range("B2").Formula = "="""""
            
        ' "Disbursement ID"
            wsSFImport.Range("C2").Formula = "='Donation Site Report'!A2"
            
        ' "Payment Check/Reference Number"
            wsSFImport.Range("D2").Formula = "='Donation Site Report'!B2"
            
        ' "Donation Date"
            wsSFImport.Range("E2").Formula = "=TEXT('Donation Site Report'!D2,""MM/DD/YYYY"")"
            
        ' "Deposit Date"
            wsSFImport.Range("F2").Formula = "=TEXT('Donation Site Report'!AJ2,""MM/DD/YYYY"")"
            
        ' "Donation Amount"
            wsSFImport.Range("G2").Formula = "='Donation Site Report'!AB2"
            
        ' "Contact1 First Name"
            wsSFImport.Range("H2").Formula = "=PROPER('Donation Site Report'!F2)"
            
        ' "Contact1 Last Name"
            wsSFImport.Range("I2").Formula = "=PROPER('Donation Site Report'!G2)"
            
        ' "Contact1 Work Email"
            wsSFImport.Range("J2").Formula = "=LOWER('Donation Site Report'!M2)"
            
        ' "Donation Type" (N/A)
            wsSFImport.Range("K2").Formula = "="""""
            
        ' "Donation Stage"
            wsSFImport.Range("L2").Formula = "=""Authorized"""
            
        ' "C&P Payment Method"
            wsSFImport.Range("M2").Formula = "=""EFT"""
            
        ' "Campaign Name"
            wsSFImport.Range("N2").Formula2 = "=A2&"" ""&" & _
                                                "IF(MONTH(E2)>6," & _
                                                    "RIGHT(E2,4)&""-""&(RIGHT(E2,2)+1)," & _
                                                    "IF(MONTH(F2)>8," & _
                                                        "RIGHT(E2,4)&""-""&(RIGHT(E2,2)+1)," & _
                                                        """""" & _
                                                    ")" & _
                                                ")&"" ""&" & _
                                                "IF(LEFT(A2,5)=""BASIS"",""ATF"",""ATF North Texas Giving Day"")"

            
        ' "Donation Name" (Opportunity Name)
            wsSFImport.Range("O2").Formula = "=XLOOKUP(A2,'School Validation'!A:A,'School Validation'!C:C)&"" ATF NTX Giving Day Donation"""
            
        ' "Donation Site"
            wsSFImport.Range("P2").Value = DonationSite
            
        ' "Description"
            wsSFImport.Range("Q2").Formula = "="""""
            
        ' "Notes"
            wsSFImport.Range("R2").Formula = "=IF(ISBLANK('Donation Site Report'!BG2),"""",'Donation Site Report'!BG2)"
    
        ' Fill Down all formulas
            If FeesReimbursed = vbYes And PrizeAwarded = vbYes Then
                TempLastRow = TempLastRow - 2
            ElseIf FeesReimbursed = vbYes Or PrizeAwarded = vbYes Then
                TempLastRow = TempLastRow - 1
            End If
            
            If TempLastRow > 2 Then
                wsSFImport.Range("A2:R" & TempLastRow).FillDown
            End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''---------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' ADD DATA VALIDATION RULES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''---------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Update the Status Bar
    Application.StatusBar = "Adding Data Validations"
    
    ' This section adds "safety rails" to the Salesforce Import File:
        ' 1. Conditional Formatting flags rows that need attention before import (common mapping/anonymous edge cases).
        ' 2. Data Validation restricts School selection to known values to prevent typos and lookup misses.

    ' ==============================================================
    ' CONDITIONAL FORMATTING: VISUAL REVIEW FLAGS
    ' ==============================================================
        ' Applies only to rows with specific BASIS legal entities in Column A (Account Name Correction).
        ' These rules highlight anonymous gifts and missing descriptions so users can review before importing.
        ' Note: These are visual flags only; they do not change data.

        With wsSFImport.Range("A2:R" & TempLastRow)
            .FormatConditions.Delete ' Reset rules to prevent duplicates on re-runs

        ' RED: High attention.
        ' Condition: BASIS entity + Anonymous donor + Description is blank.
        ' Rationale: Anonymous entries without a description are more likely to be questioned or need clarification before import.
            .FormatConditions.Add Type:=xlExpression, _
                Formula1:="=AND(OR($A2=""BASIS Charter Schools, Inc."",$A2=""BASIS Texas Charter Schools, Inc."",$A2=""BASIS Baton Rouge Schools, Inc."",$A2=""BASIS DC, Inc.""), $H2=""Anonymous"", TRIM($Q2)="""")"
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0)

        ' ORANGE: Medium attention.
        ' Condition: BASIS entity + Anonymous donor.
        ' Rationale: Anonymous donations typically require a quick manual spot-check for naming / description conventions.
            .FormatConditions.Add Type:=xlExpression, _
                Formula1:="=AND(OR($A2=""BASIS Charter Schools, Inc."",$A2=""BASIS Texas Charter Schools, Inc."",$A2=""BASIS Baton Rouge Schools, Inc."",$A2=""BASIS DC, Inc.""), $H2=""Anonymous"")"
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 165, 0)

        ' YELLOW: Low attention.
        ' Condition: BASIS entity + donor is not Anonymous.
        ' Rationale: Highlights rows that successfully mapped to a BASIS entity so the user can quickly scan for consistency.
            .FormatConditions.Add Type:=xlExpression, _
                Formula1:="=AND(OR($A2=""BASIS Charter Schools, Inc."",$A2=""BASIS Texas Charter Schools, Inc."",$A2=""BASIS Baton Rouge Schools, Inc."",$A2=""BASIS DC, Inc.""), $H2<>""Anonymous"")"
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 255, 0)
        End With

    ' ==============================================================
    ' DATA VALIDATION: PREVENT INVALID SCHOOL ENTRIES
    ' ==============================================================
        ' Restricts Column A (Account Name Correction) to the official SF School Name list.
        ' This prevents typos / unmatched schools that would break downstream lookups and reconciliation.

            With wsSFImport.Range("A2:A" & TempLastRow).Validation
                .Delete ' Reset validation to prevent stacking rules
    
                .Add Type:=xlValidateList, _
                     AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, _
                     Formula1:="='School Validation'!$A$2:$A$" & ValidationLastRow
    
                .IgnoreBlank = True
                .InCellDropdown = True
                .ErrorTitle = "Invalid Entry"
                .ErrorMessage = "Please select a valid school from the list."
                .ShowError = True
            End With

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' FINAL CLEANUP AND FINISHING TOUCHES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Update the Status Bar
    Application.StatusBar = "Final Cleanup"
    
    ' ==============================================================
    ' FORMAT WORKSHEETS
    ' ==============================================================
        ' --------------------------------------------------------------
        ' DONATION SITE WORKSHEET
        ' --------------------------------------------------------------
            wsDonationSiteReport.Range("A1:BG1").AutoFilter
            wsDonationSiteReport.Columns("A:BG").AutoFit
                
        ' --------------------------------------------------------------
        ' SALESFORCE IMPORT FILE WORKSHEET
        ' --------------------------------------------------------------
            wsSFImport.Range("A1:R1").AutoFilter
            wsSFImport.Columns("A:R").AutoFit
            wsSFImport.Tab.Color = vbYellow
                
        ' --------------------------------------------------------------
        ' DATA VALIDATION WORKSHEET - SCHOOLS
        ' --------------------------------------------------------------
            wsSchoolValidation.Range("A1:B1").AutoFilter
            wsSchoolValidation.Columns("A:B").AutoFit
            wsSchoolValidation.Visible = xlSheetHidden
                
        ' --------------------------------------------------------------
        ' OPTIONAL: AllowReportPullIn: ON
        ' --------------------------------------------------------------
            If AllowReportsPullIn = True Then
                wsDonationSiteReportRaw.Range("A1:BN1").AutoFilter
                wsDonationSiteReportRaw.Columns("A:BN").AutoFit
            End If

    ' ==============================================================
    ' CREATE CLOSING MESSAGES
    ' ==============================================================
        ExitMessage = "The converter completed successfully. Thank you for your patience!" & vbCrLf & vbCrLf & "The Salesforce Import File is now completed for the 'NTX Giving' Report"
        ExitMessage_Title = "Conversion Completed Successfully"
    
    ' ==============================================================
    ' JUMP TO THE 'CompleteMacro' PORTION
    ' ==============================================================
        GoTo CompleteMacro

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''--------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''  Consolidation Only Section ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''--------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ConsolidationOnly:
'


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''----------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' COMPLETING THE MACRO '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''----------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CompleteMacro:
    ' --------------------------------------------------------------
    ' PROVIDE THE USER WITH A MESSAGE
    ' --------------------------------------------------------------
        MsgBox ExitMessage, _
               vbOKOnly, _
               ExitMessage_Title
    
    ' --------------------------------------------------------------
    ' TURN BACK ON ALERTS
    ' --------------------------------------------------------------
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
    
End Sub
