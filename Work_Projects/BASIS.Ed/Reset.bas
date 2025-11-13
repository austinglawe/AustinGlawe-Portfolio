Public Sub Create_Reset_Worksheet()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' Purpose and Updates Log '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    ' PURPOSE
    ' ============================================================
        ' Deletes all existing worksheets and creates a clean "COMPLETE RESET" worksheet that allows the user to return to the Converter Selection Page.
        
    ' ============================================================
    ' Update Log (LAST UPDATED: 2025.10.15)
    ' ============================================================
        ' Initial Setup Time: 2023.04.05 - 2023.04.05
        ' Production Rollout Date: 2023.04.19
        
        ' Updates:
            ' 2025.10.15 - Reformat Code

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-----------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-----------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Declare variables
        Dim wsNew As Worksheet
        Dim ws As Worksheet
        Dim SelectionPageButton As Button
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''--------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' Initialize the Reset Worksheet ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''--------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Prepare Excel environment for reset
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.StatusBar = "Creating reset worksheet..."
    
    ' Create the new worksheet that will serve as the "COMPLETE RESET" worksheet.
        Set wsNew = ThisWorkbook.Worksheets.Add
        
    ' Delete all worksheets that are not the newly created 'wsNew' worksheet.
            For Each ws In ThisWorkbook.Worksheets
                If ws.Name <> wsNew.Name Then
                    ws.Delete
                End If
            Next ws
            
    ' Rename the new worksheet to "COMPLETE RESET"
        wsNew.Name = "COMPLETE RESET"
        
    ' Apply a cautionary background to all the cells in the 'wsNew' worksheet.
        wsNew.Cells.Interior.Color = RGB(255, 235, 80)
        
    ' Create the "Return to Selection Page" button
        ' Create the button (Starting Point- Left to Right, Starting Point- Up and Down, Length of Button, Height of Button)
            Set SelectionPageButton = wsNew.Buttons.Add(200, 0, 800, 400)
            
        ' Format the 'SelectionPageButton' button
            With SelectionPageButton
                .Caption = "CLICK HERE TO RESET THE WORKBOOK AND RETURN TO THE 'Converter Selection Page'"
                .OnAction = "Reset.Selection_Page"
                .Font.Size = 60
                .Font.Bold = True
                .Font.Color = RGB(205, 0, 0)
            End With
    
    ' Format the 'wsNew' worksheet tab and finalize setup
        wsNew.Tab.Color = RGB(255, 235, 80)
        
    ' Restore Excel environment to normal
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.StatusBar = False
    
End Sub
Public Sub Selection_Page()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' Purpose and Updates Log '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    ' PURPOSE
    ' ============================================================
        ' This macro is meant to go back to selection page (re-creating the selection page) - To be able to select a macro to run.
        
    ' ============================================================
    ' Update Log (LAST UPDATED: 2025.11.04)
    ' ============================================================
        ' Initial Setup Time: 2023.04.05 - 2023.04.05
        ' Production Rollout Date: 2023.04.19
        
        ' Updates:
            ' 2025.10.14 - Reformat and Restructure Code
            ' 2025.10.17 - Create Variables for colors
            ' 2025.11.04 - Add ".", "..", "..." progression in worksheet deletion

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-----------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-----------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    ' Declare the key variables for setup and formatting
    ' ============================================================
        Dim UserResponse As VbMsgBoxResult
        
        Dim wsSelectionPage As Worksheet
        Dim ws As Worksheet
        Dim AnimationCount As Long
        Dim AnimationDots As String
        
        Dim HeightofRow As Long
        
        Dim SFReportLink_ManualImports As String
        Dim SFReportLink_ClickandPledge As String
        Dim IntacctLink As String
        
        Dim SelectionPageLastRow As Long
        Dim SelectionRow As Long

    ' ============================================================
    ' Define row number variables for each converter
    ' ============================================================
        ' Donation Sites
            ' AR Team
'                Dim AR_225GivesRow As Long
                Dim AR_AZGivesRow As Long
                Dim AR_BenevityRow As Long
                Dim AR_ClickAndPledgeRow As Long
                Dim AR_CyberGrantsRow As Long
'                Dim AR_FidelityGiftsRow As Long
                Dim AR_FrontStreamRow As Long
                Dim AR_GiveGabRow As Long
                Dim AR_NTXGivingRow As Long
                Dim AR_YourCauseRow As Long
                Dim AR_CraveItRow As Long
                
            ' Salesforce (SF) Team
'                Dim SF_225GivesRow As Long
                Dim SF_AZGivesRow As Long
                Dim SF_BenevityRow As Long
                Dim SF_CyberGrantsRow As Long
'                Dim SF_FidelityGiftsRow As Long
                Dim SF_FrontStreamRow As Long
                Dim SF_GiveGabRow As Long
                Dim SF_NTXGivingRow As Long
                Dim SF_YourCauseRow As Long
            
        ' Other Processes
            Dim BankStatementAnalysisRow As Long
            Dim BlackbaudARReconRow As Long
            Dim BlackbaudCRJsRow As Long
            Dim BlackbaudWebsiteDepositsRow As Long
            Dim OtherRevenueReconRow As Long
            Dim ReportConsolidationRow As Long
            Dim SDReconRow As Long
            
    ' ============================================================
    ' Define button variables for each converter
    ' ============================================================
        ' Visual property values
            Dim ConvertersFontSize As Long
            
            Dim ButtonColor_Accounting As Long
            Dim ButtonColor_AR As Long
            Dim ButtonColor_SF As Long
            Dim ButtonColor_Universal As Long
        
        ' Donation Sites
            ' AR Team
'               Dim AR_225GivesButton As Button
                Dim AR_AZGivesButton As Button
                Dim AR_BenevityButton As Button
                Dim AR_ClickAndPledgeButton As Button
                Dim AR_CraveItButton As Button
                Dim AR_CyberGrantsButton As Button
'               Dim AR_FidelityGiftsButton As Button
                Dim AR_FrontStreamButton As Button
                Dim AR_GiveGabButton As Button
                Dim AR_NTXGivingButton As Button
                Dim AR_YourCauseButton As Button
            
            ' SF Team
'               Dim SF_225GivesButton As Button
                Dim SF_AZGivesButton As Button
                Dim SF_BenevityButton As Button
                Dim SF_CyberGrantsButton As Button
'               Dim SF_FidelityGiftsButton As Button
                Dim SF_FrontStreamButton As Button
                Dim SF_GiveGabButton As Button
                Dim SF_NTXGivingButton As Button
                Dim SF_YourCauseButton As Button
            
        ' Other Processes
            Dim BankStatementAnalysisButton As Button
            Dim BlackbaudARReconButton As Button
            Dim BlackbaudCRJsButton As Button
            Dim BlackbaudWebsiteDepositsButton As Button
            Dim OtherRevenueReconButton As Button
            Dim ReportConsolidationButton As Button
            Dim SDReconButton As Button
            
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' Assign Variables Values '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    ' Assign row numbers to each converter for layout positioning
    ' ============================================================
        ' Donation Sites:
            ' ------------------------------------------------------------
            ' 225 Gives
            ' ------------------------------------------------------------
'                AR_225GivesRow = 2
'                SF_225GivesRow = AR_225GivesRow + 1
                
            ' ------------------------------------------------------------
            ' AZ Gives
            ' ------------------------------------------------------------
                ' AR_AZGivesRow = SF_225GivesRow + 1
                AR_AZGivesRow = 2
                SF_AZGivesRow = AR_AZGivesRow + 1
                
            ' ------------------------------------------------------------
            ' Benevity
            ' ------------------------------------------------------------
                AR_BenevityRow = SF_AZGivesRow + 1
                SF_BenevityRow = AR_BenevityRow + 1
                
            ' ------------------------------------------------------------
            ' Click and Pledge
            ' ------------------------------------------------------------
                AR_ClickAndPledgeRow = SF_BenevityRow + 1
                
            ' ------------------------------------------------------------
            ' Crave It
            ' ------------------------------------------------------------
                AR_CraveItRow = AR_ClickAndPledgeRow + 1
                
            ' ------------------------------------------------------------
            ' Cyber Grants
            ' ------------------------------------------------------------
                AR_CyberGrantsRow = AR_CraveItRow + 1
                SF_CyberGrantsRow = AR_CyberGrantsRow + 1
                
            ' ------------------------------------------------------------
            ' Fidelity
            ' ------------------------------------------------------------
'                AR_FidelityGiftsRow = SF_CyberGrantsRow + 1
'                SF_FidelityGiftsRow = AR_FidelityGiftsRow + 1
                
            ' ------------------------------------------------------------
            ' Front Stream
            ' ------------------------------------------------------------
                ' AR_FrontStreamRow = SF_FidelityGiftsRow + 1
                AR_FrontStreamRow = SF_CyberGrantsRow + 1
                SF_FrontStreamRow = AR_FrontStreamRow + 1
                
            ' ------------------------------------------------------------
            ' Give Gab/Big Gives/Bonterra Tech
            ' ------------------------------------------------------------
                AR_GiveGabRow = SF_FrontStreamRow + 1
                SF_GiveGabRow = AR_GiveGabRow + 1
                
            ' ------------------------------------------------------------
            ' NTX Giving Day
            ' ------------------------------------------------------------
                AR_NTXGivingRow = SF_GiveGabRow + 1
                SF_NTXGivingRow = AR_NTXGivingRow + 1
                
            ' ------------------------------------------------------------
            ' Your Cause
            ' ------------------------------------------------------------
                AR_YourCauseRow = SF_NTXGivingRow + 1
                SF_YourCauseRow = AR_YourCauseRow + 1
        
        ' Other Processes
            ' ------------------------------------------------------------
            ' Bank Statement Analysis
            ' ------------------------------------------------------------
                BankStatementAnalysisRow = SF_YourCauseRow + 1
            ' ------------------------------------------------------------
            ' Blackbaud AR Reconciliations (12010-FY)
            ' ------------------------------------------------------------
                BlackbaudARReconRow = BankStatementAnalysisRow + 1
            ' ------------------------------------------------------------
            ' Blackbaud CRJs
            ' ------------------------------------------------------------
                BlackbaudCRJsRow = BlackbaudARReconRow + 1
        
            ' ------------------------------------------------------------
            ' Blackbaud Website Deposits (School Level - Remittance)
            ' ------------------------------------------------------------
                BlackbaudWebsiteDepositsRow = BlackbaudCRJsRow + 1
        
            ' ------------------------------------------------------------
            ' Other Revenue Reconciliation (73010, 73010-9999)
            ' ------------------------------------------------------------
                OtherRevenueReconRow = BlackbaudWebsiteDepositsRow + 1
            
            ' ------------------------------------------------------------
            ' Report Consolidation (Consolidate any report)
            ' ------------------------------------------------------------
                ReportConsolidationRow = OtherRevenueReconRow + 1
            
            ' ------------------------------------------------------------
            ' Security Deposit (SD) Reconciliation
            ' ------------------------------------------------------------
                SDReconRow = ReportConsolidationRow + 1
        
        ' ============================================================
        ' Assign Report Links
        ' ============================================================
            SFReportLink_ClickandPledge = "https://basised.lightning.force.com/lightning/r/Report/00ORj000005Sw6DMAS/view?queryScope=userFolders"
            
            SFReportLink_ManualImports = "https://basised.lightning.force.com/lightning/r/Report/00OHp000008pMl6MAE/view?queryScope=userFolders"
            
            IntacctLink = "https://www.intacct.com/ia/acct/login.phtml"
           
        ' ============================================================
        ' Assign Visual Property Values for the 'Create Macro Buttons'
        ' Section
        ' ============================================================
            ConvertersFontSize = 15
            
            ButtonColor_Accounting = RGB(255, 165, 0)
            ButtonColor_AR = RGB(0, 175, 0)
            ButtonColor_SF = RGB(0, 130, 225)
            ButtonColor_Universal = RGB(0, 0, 0)
                
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''---------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' Confirm and Initialize Workbook Reset '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''---------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    ' Confirm user wants to return to the Converter Selection Page
    ' ============================================================
        ' Ask user if they would like to reset the workbook and set up the selection page.
            UserResponse = MsgBox(Title:="Confirmation to go to Converter Selection Page", _
                    Prompt:="Are you sure you want to go back to the Converter Selection Page? If you choose yes, all data will be lost. This process is irreversible.", _
                    Buttons:=vbYesNo + vbQuestion)

        ' If the user selects No, exit the sub
            If UserResponse = vbNo Then
                Exit Sub
            End If
    
    ' ============================================================
    ' Prepare Excel environment for setup
    ' ============================================================
            Application.DisplayAlerts = False
            Application.ScreenUpdating = False
            Application.Calculation = xlManual

    ' ============================================================
    ' Create a new worksheet using the unique TimeSheet name
    ' ============================================================
        Set wsSelectionPage = ThisWorkbook.Worksheets.Add
    
    ' ============================================================
    ' Remove all other worksheets to ensure a clean environment
    ' ============================================================
' Update the Status Bar to let the user know all worksheets are being deleted.
    Application.StatusBar = "Deleting all worksheets..."
        
        ' Delete all worksheets except for the newly created 'wsSelectionPage' worksheet.
            AnimationCount = 1
            AnimationDots = "."
            
            For Each ws In ThisWorkbook.Worksheets
                Select Case AnimationCount
                Case 1
                    AnimationDots = "."
                Case 2
                    AnimationDots = ".."
                Case 3
                    AnimationDots = "..."
                Case Else
                    AnimationDots = "."
                    AnimationCount = 1
                End Select
                
                Application.StatusBar = "Deleting all worksheets" & AnimationDots
                If ws.Name <> wsSelectionPage.Name Then
                    ws.Delete
                End If
                
                AnimationCount = AnimationCount + 1
            Next ws

    ' ============================================================
    ' Rename the new worksheet to "Converter Selection Page"
    ' ============================================================
        wsSelectionPage.Name = "Converter Selection Page"
        
    ' ============================================================
    ' Initialize setup and add column headers
    ' ============================================================
' Update the Status Bar to let the user know the 'Converter Selection Page' is being set up.
    Application.StatusBar = "Setting up Converter Selection Page..."
    
    ' Create the headers for the "Converter Selection Page" Worksheet.
        wsSelectionPage.Range("A1:H1").Value = Array("Click Below To Run Macro", "Converter", "Department", "Website Report Link", _
                "Salesforce Report Link", "Intacct Link", "SharePoint Report Link", "Converter Description")
    
    ' ============================================================
    ' Format header row (bold, underline, and enable AutoFilter)
    ' ============================================================
        With wsSelectionPage.Range("A1:H1")
            .Font.Bold = True
            .Font.Underline = True
            .AutoFilter
        End With
        
    ' ============================================================
    ' Set initial column widths, row heights, and base formatting
    ' ============================================================
        HeightofRow = 54
        
        wsSelectionPage.Range("A:A").ColumnWidth = 69#
        wsSelectionPage.Range("B:H").Columns.AutoFit
        wsSelectionPage.Rows("1:1").RowHeight = 18#
        wsSelectionPage.Rows("1:40").RowHeight = HeightofRow
        With wsSelectionPage.Cells
            .Interior.Color = RGB(0, 0, 0)
            .Font.Color = RGB(255, 255, 255)
            .VerticalAlignment = xlCenter
        End With
            

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''---------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' Populate Data Table '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''---------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    ' Populate Columns B: "Name for Converter"
    ' ============================================================
        ' Donation Sites:
            ' ------------------------------------------------------------
            ' 225 Gives
            ' ------------------------------------------------------------
'                wsSelectionPage.Range("B" & AR_225GivesRow).Value = "225 Gives"
'                wsSelectionPage.Range("B" & SF_225GivesRow).Value = "225 Gives"
                
            ' ------------------------------------------------------------
            ' AZ Gives
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & AR_AZGivesRow).Value = "AZ Gives"
                wsSelectionPage.Range("B" & SF_AZGivesRow).Value = "AZ Gives"
                
            ' ------------------------------------------------------------
            ' Benevity
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & AR_BenevityRow).Value = "Benevity"
                wsSelectionPage.Range("B" & SF_BenevityRow).Value = "Benevity"
                
            ' ------------------------------------------------------------
            ' Click and Pledge
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & AR_ClickAndPledgeRow).Value = "Click & Pledge"
                
            ' ------------------------------------------------------------
            ' Crave It
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & AR_CraveItRow).Value = "Crave It"
                
            ' ------------------------------------------------------------
            ' Cyber Grants
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & AR_CyberGrantsRow).Value = "Cyber Grants"
                wsSelectionPage.Range("B" & SF_CyberGrantsRow).Value = "Cyber Grants"
                
            ' ------------------------------------------------------------
            ' Fidelity
            ' ------------------------------------------------------------
'                wsSelectionPage.Range("B" & AR_FidelityGiftsRow).Value = "Fidelity Gifts"
'                wsSelectionPage.Range("B" & SF_FidelityGiftsRow).Value = "Fidelity Gifts"
                
            ' ------------------------------------------------------------
            ' Front Stream
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & AR_FrontStreamRow).Value = "Front Stream"
                wsSelectionPage.Range("B" & SF_FrontStreamRow).Value = "Front Stream"
                
            ' ------------------------------------------------------------
            ' Give Gab/Big Gives/Bonterra Tech
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & AR_GiveGabRow).Value = "Give Gab/Big Gives/Bonterra Tech"
                wsSelectionPage.Range("B" & SF_GiveGabRow).Value = "Give Gab/Big Gives/Bonterra Tech"
                
            ' ------------------------------------------------------------
            ' NTX Giving
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & AR_NTXGivingRow).Value = "NTX Giving Day"
                wsSelectionPage.Range("B" & SF_NTXGivingRow).Value = "NTX Giving Day"
                
            ' ------------------------------------------------------------
            ' Your Cause
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & AR_YourCauseRow).Value = "Your Cause"
                wsSelectionPage.Range("B" & SF_YourCauseRow).Value = "Your Cause"
        
        ' Other Processes
            ' ------------------------------------------------------------
            ' Bank Statement Analysis
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & BankStatementAnalysisRow).Value = "Bank Statement Analysis"
                
            ' ------------------------------------------------------------
            ' Blackbaud AR Reconciliations (12010-FY)
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & BlackbaudARReconRow).Value = "Blackbaud AR Reconciliations (12010-FY)"
                
            ' ------------------------------------------------------------
            ' Blackbaud CRJs
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & BlackbaudCRJsRow).Value = "Blackbaud CRJs"
            
            ' ------------------------------------------------------------
            ' Blackbaud Website Deposits (School Level - Remittance)
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & BlackbaudWebsiteDepositsRow).Value = "Blackbaud Website Deposits (School Level - Remittance Report)"
  
            ' ------------------------------------------------------------
            ' Other Revenue (73010, 73010-9999) Reconciliation
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & OtherRevenueReconRow).Value = "Other Revenue (73010, 73010-9999) Reconciliation"
                
            ' ------------------------------------------------------------
            ' Report Consolidation (Consolidate any report)
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & ReportConsolidationRow).Value = "Report Consolidation (Consolidate Any Report Type)"
                
            ' ------------------------------------------------------------
            ' Security Deposit (SD) Reconciliation
            ' ------------------------------------------------------------
                wsSelectionPage.Range("B" & SDReconRow).Value = "Security Deposit (SD) Reconciliation"
        
    ' ============================================================
    ' Populate Column C: "Department of Converter"
    ' ============================================================
        ' Donation Sites:
            ' ------------------------------------------------------------
            ' 225 Gives
            ' ------------------------------------------------------------
'                wsSelectionPage.Range("C" & AR_225GivesRow).Value = "Accounts Receivable (AR) Team"
'                wsSelectionPage.Range("C" & SF_225GivesRow).Value = "Salesforce Team"
            
            ' ------------------------------------------------------------
            ' AZ Gives
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & AR_AZGivesRow).Value = "Accounts Receivable (AR) Team"
                wsSelectionPage.Range("C" & SF_AZGivesRow).Value = "Salesforce Team"
            
            ' ------------------------------------------------------------
            ' Benevity
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & AR_BenevityRow).Value = "Accounts Receivable (AR) Team"
                wsSelectionPage.Range("C" & SF_BenevityRow).Value = "Salesforce Team"
            
            ' ------------------------------------------------------------
            ' Click and Pledge
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & AR_ClickAndPledgeRow).Value = "Accounts Receivable (AR) Team"
                
            ' ------------------------------------------------------------
            ' Crave It
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & AR_CraveItRow).Value = "Accounts Receivable (AR) Team"
            
            ' ------------------------------------------------------------
            '  Cyber Grants
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & AR_CyberGrantsRow).Value = "Accounts Receivable (AR) Team"
                wsSelectionPage.Range("C" & SF_CyberGrantsRow).Value = "Salesforce Team"
            
            ' ------------------------------------------------------------
            ' Fidelity
            ' ------------------------------------------------------------
'                wsSelectionPage.Range("C" & AR_FidelityGiftsRow).Value = "Accounts Receivable (AR) Team"
'                wsSelectionPage.Range("C" & SF_FidelityGiftsRow).Value = "Salesforce Team"
            
            ' ------------------------------------------------------------
            '  Front Stream
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & AR_FrontStreamRow).Value = "Accounts Receivable (AR) Team"
                wsSelectionPage.Range("C" & SF_FrontStreamRow).Value = "Salesforce Team"
                
            ' ------------------------------------------------------------
            ' Give Gab/Big Gives/Bonterra Tech
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & AR_GiveGabRow).Value = "Accounts Receivable (AR) Team"
                wsSelectionPage.Range("C" & SF_GiveGabRow).Value = "Salesforce Team"
            
            ' ------------------------------------------------------------
            ' NTX Giving Day
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & AR_NTXGivingRow).Value = "Accounts Receivable (AR) Team"
                wsSelectionPage.Range("C" & SF_NTXGivingRow).Value = "Salesforce Team"

            ' ------------------------------------------------------------
            ' Your Cause
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & AR_YourCauseRow).Value = "Accounts Receivable (AR) Team"
                wsSelectionPage.Range("C" & SF_YourCauseRow).Value = "Salesforce Team"
        
        ' Other Processes
            ' ------------------------------------------------------------
            ' Bank Statement Analysis
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & BankStatementAnalysisRow).Value = "Accounting Team"
                
            ' ------------------------------------------------------------
            ' Blackbaud AR Reconciliations (12010-FY)
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & BlackbaudARReconRow).Value = "Accounting Team"
            
            ' ------------------------------------------------------------
            ' Blackbaud CRJs
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & BlackbaudCRJsRow).Value = "Accounts Receivable (AR) Team"
            
            ' ------------------------------------------------------------
            ' Blackbuad Website Deposits
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & BlackbaudWebsiteDepositsRow).Value = "Accounting Team"
            
            ' ------------------------------------------------------------
            ' Other Revenue Reconciliation (73010, 73010-9999)
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & OtherRevenueReconRow).Value = "Accounts Receivable (AR) Team"
                
            ' ------------------------------------------------------------
            ' Report Consolidation (Consolidate any report)
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & ReportConsolidationRow).Value = "Universal"
                
            ' ------------------------------------------------------------
            ' Security Deposit (SD) Reconciliation
            ' ------------------------------------------------------------
                wsSelectionPage.Range("C" & SDReconRow).Value = "Accounting Team"

    ' ============================================================
    ' Populate Column D: "Website Report Link"
    ' ============================================================
        ' Donation Sites:
            ' ------------------------------------------------------------
            ' 225 Gives
            ' ------------------------------------------------------------
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("D" & AR_225GivesRow), _
'                        Address:="https://app.neonsso.com/login?" & _
'                            "client_id=CJtL0DWvnU8LBPz0S6dYu4ydoA3zYsCu" & _
'                            "&response_type=code" & _
'                            "&scope=openid" & _
'                            "&redirect_uri=https%3A%2F%2Fapp.neongivingdays.com%2Findex.php%3Fsection%3Ddashboards" & _
'                            "%26action%3Dcondition%26loginModule%3D0%26ssoLogin%3D1" & _
'                            "&state=6501d7c6cc3f3" & _
'                            "&nonce=6501d7c6cc3f5" & _
'                            "&config_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwczpcL1wvYXBwLm5lb25naXZpbmdkYXlzLmNvbSIsImlhdCI6MTY5NDYxOTU5MCwiZXhw" & _
'                                "IjoxNjk0NjIzMTkwLCJwcm9kdWN0X2JyYW5kX2lkIjoiMTciLCJob3N0cyI6ImFwcC5uZW9uZ2l2aW5nZGF5cy5jb20ifQ.lugmOChzTnNZaY8PRAC2siXtHYKETjlkHXY0KLROblA", _
'                        TextToDisplay:="225 Gives Donation Website Link"

'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("D" & SF_225GivesRow), _
'                        Address:="https://app.neonsso.com/login?" & _
'                            "client_id=CJtL0DWvnU8LBPz0S6dYu4ydoA3zYsCu" & _
'                            "&response_type=code" & _
'                            "&scope=openid" & _
'                            "&redirect_uri=https%3A%2F%2Fapp.neongivingdays.com%2Findex.php%3Fsection%3Ddashboards" & _
'                            "%26action%3Dcondition%26loginModule%3D0%26ssoLogin%3D1" & _
'                            "&state=6501d7c6cc3f3" & _
'                            "&nonce=6501d7c6cc3f5" & _
'                            "&config_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwczpcL1wvYXBwLm5lb25naXZpbmdkYXlzLmNvbSIsImlhdCI6MTY5NDYxOTU5MCwiZXhw" & _
'                                "IjoxNjk0NjIzMTkwLCJwcm9kdWN0X2JyYW5kX2lkIjoiMTciLCJob3N0cyI6ImFwcC5uZW9uZ2l2aW5nZGF5cy5jb20ifQ.lugmOChzTnNZaY8PRAC2siXtHYKETjlkHXY0KLROblA", _
'                        TextToDisplay:="225 Gives Donation Website Link"
                
            ' ------------------------------------------------------------
            ' AZ Gives
            ' ------------------------------------------------------------
                ' Current AZ Gives Link:
                    wsSelectionPage.Hyperlinks.Add _
                            Anchor:=wsSelectionPage.Range("D" & AR_AZGivesRow), _
                            Address:="https://www.azgives.org/login", _
                            TextToDisplay:="AZ Gives Donation Website Link"
                        
'                ' Old AZ Gives Link:
'                    wsSelectionPage.Hyperlinks.Add _
'                            Anchor:=wsSelectionPage.Range("D" & AR_AZGivesRow), _
'                            Address:="https://app.neonsso.com/login?client_id=" & _
'                                "CJtL0DWvnU8LBPz0S6dYu4ydoA3zYsCu" & _
'                                "&response_type=code" & _
'                                "&scope=openid" & _
'                                "&redirect_uri=https%3A%2F%2Fazgives.civicore.com%2Findex.php%3Fsection%3Ddashboards" & _
'                                "%26action%3Dcondition%26loginModule%3D0%26ssoLogin%3D1" & _
'                                "&state=64d55bea564fe" & _
'                                "&nonce=64d55bea564ff" & _
'                                "&config_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwczpcL1wvYXpnaXZlcy5jaXZpY29yZS5jb20iLCJpYXQiOjE2OTE3MDQyOTgsImV4" & _
'                                    "cCI6MTY5MTcwNzg5OCwicHJvZHVjdF9icmFuZF9pZCI6IjEyIiwiaG9zdHMiOiJhemdpdmVzLmNpdmljb3JlLmNvbSJ9.uCQrNKOfZD_AvjvYDotRUrNcF4O9OzBE89JeCbkcG0c", _
'                            TextToDisplay:="AZ Gives Donation Website Link"
                        
                ' Current AZ Gives Link:
                    wsSelectionPage.Hyperlinks.Add _
                            Anchor:=wsSelectionPage.Range("D" & SF_AZGivesRow), _
                            Address:="https://www.azgives.org/login", _
                            TextToDisplay:="AZ Gives Donation Website Link"
                        
'                ' Old AZ Gives Link:
'                    wsSelectionPage.Hyperlinks.Add _
'                            Anchor:=wsSelectionPage.Range("D" & SF_AZGivesRow), _
'                            Address:="https://app.neonsso.com/login?client_id=" & _
'                                "CJtL0DWvnU8LBPz0S6dYu4ydoA3zYsCu" & _
'                                "&response_type=code" & _
'                                "&scope=openid" & _
'                                "&redirect_uri=https%3A%2F%2Fazgives.civicore.com%2Findex.php%3Fsection%3Ddashboards" & _
'                                "%26action%3Dcondition%26loginModule%3D0%26ssoLogin%3D1" & _
'                                "&state=64d55bea564fe" & _
'                                "&nonce=64d55bea564ff" & _
'                                "&config_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwczpcL1wvYXpnaXZlcy5jaXZpY29yZS5jb20iLCJpYXQiOjE2OTE3MDQyOTgsImV4" & _
'                                    "cCI6MTY5MTcwNzg5OCwicHJvZHVjdF9icmFuZF9pZCI6IjEyIiwiaG9zdHMiOiJhemdpdmVzLmNpdmljb3JlLmNvbSJ9.uCQrNKOfZD_AvjvYDotRUrNcF4O9OzBE89JeCbkcG0c", _
'                            TextToDisplay:="AZ Gives Donation Website Link"
                
            ' ------------------------------------------------------------
            ' Benevity
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & AR_BenevityRow), _
                        Address:="https://causes.benevity.org/user?hsCtaTracking=70669402-b761-4e69-b77b-ffe82c79f012%7C2d2019f4-8678-4187-a2bb-d36e5b1095ca", _
                        TextToDisplay:="Benevity Donation Website Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & SF_BenevityRow), _
                        Address:="https://causes.benevity.org/user?hsCtaTracking=70669402-b761-4e69-b77b-ffe82c79f012%7C2d2019f4-8678-4187-a2bb-d36e5b1095ca", _
                        TextToDisplay:="Benevity Donation Website Link"
                
            ' ------------------------------------------------------------
            ' Click and Pledge
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & AR_ClickAndPledgeRow), _
                        Address:="https://login.connect.clickandpledge.com/", _
                        TextToDisplay:="Click and Pledge Donation Website Link"
                
            ' ------------------------------------------------------------
            ' Crave It
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & AR_CraveItRow), _
                        Address:="https://craveit.boonli.com/login", _
                        TextToDisplay:="Crave It Website Link"
                
            ' ------------------------------------------------------------
            ' Cyber Grants
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & AR_CyberGrantsRow), _
                        Address:="https://www.cybergrants.com/pls/cybergrants/ao_login.login?x_gm_id=1&x_style_id=&x_proposal_type_id=9019", _
                        TextToDisplay:="Cyber Grants Donation Website Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & SF_CyberGrantsRow), _
                        Address:="https://www.cybergrants.com/pls/cybergrants/ao_login.login?x_gm_id=1&x_style_id=&x_proposal_type_id=9019", _
                        TextToDisplay:="Cyber Grants Donation Website Link"
                        
                
            ' ------------------------------------------------------------
            ' Fidelity
            ' ------------------------------------------------------------
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("D" & AR_FidelityGiftsRow), _
'                        Address:="https://connect.stripe.com/express_login", _
'                        TextToDisplay:="Fidelity Giving Donation Website Link"
'
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("D" & SF_FidelityGiftsRow), _
'                        Address:="https://connect.stripe.com/express_login", _
'                        TextToDisplay:="Fidelity Giving Donation Website Link"
                
            ' ------------------------------------------------------------
            ' Front Stream
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & AR_FrontStreamRow), _
                        Address:="https://connect.frontstream.com/Login", _
                        TextToDisplay:="Front Stream Donation Website Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & SF_FrontStreamRow), _
                        Address:="https://connect.frontstream.com/Login", _
                        TextToDisplay:="Front Stream Donation Website Link"
                        
            ' ------------------------------------------------------------
            ' Give Gab/Big Gives/Bonterra Tech
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & AR_GiveGabRow), _
                        Address:="https://www.givegab.com/users/sign_in", _
                        TextToDisplay:="Give Gab/Big Gives/Bonterra Tech Donation Website Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & SF_GiveGabRow), _
                        Address:="https://www.givegab.com/users/sign_in", _
                        TextToDisplay:="Give Gab/Big Gives/Bonterra Tech Donation Website Link"
                
            ' ------------------------------------------------------------
            ' NTX Giving
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & AR_NTXGivingRow), _
                        Address:="https://www.northtexasgivingday.org/login?redirectUrl=/giving-events/ntx23", _
                        TextToDisplay:="NTX Giving Donation Website Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & SF_NTXGivingRow), _
                        Address:="https://www.northtexasgivingday.org/login?redirectUrl=/giving-events/ntx23", _
                        TextToDisplay:="NTX Giving Donation Website Link"
                
            ' ------------------------------------------------------------
            ' Your Cause
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & AR_YourCauseRow), _
                        Address:="https://nonprofit.yourcause.com/login", _
                        TextToDisplay:="Your Cause Donation Website Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & SF_YourCauseRow), _
                        Address:="https://nonprofit.yourcause.com/login", _
                        TextToDisplay:="Your Cause Donation Website Link"
        
        ' Other Processes
            ' ------------------------------------------------------------
            ' Bank Statement Analysis
            ' ------------------------------------------------------------
                ' No Website Link
                
            ' ------------------------------------------------------------
            ' Blackbaud AR Reconciliations (12010-FY)
            ' ------------------------------------------------------------
                ' wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & BlackbaudARReconRow), _
                        Address:="", _
                        TextToDisplay:=""
                
            ' ------------------------------------------------------------
            ' Blackbaud CRJs
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & BlackbaudCRJsRow), _
                        Address:="https://tuition.blackbaud.school/#/enterprise/enrollment.html", _
                        TextToDisplay:="Blackbaud Website Link"
            
            ' ------------------------------------------------------------
            ' Blackbaud Website Deposits (School Level - Remittance)
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("D" & BlackbaudWebsiteDepositsRow), _
                        Address:="https://tuition.blackbaud.school/#/main/landing.html", _
                        TextToDisplay:="Blackbaud Website Link"
                
            ' ------------------------------------------------------------
            ' Other Revenue (73010, 73010-9999) Reconciliation
            ' ------------------------------------------------------------
                ' No Website Link
            
            ' ------------------------------------------------------------
            ' Report Consolidation (Consolidate any report)
            ' ------------------------------------------------------------
                ' No Website Link
                
            ' ------------------------------------------------------------
            ' Security Deposit (SD) Reconciliation
            ' ------------------------------------------------------------
                ' No Website Link
        
    ' ============================================================
    ' Populate Column E: "Salesforce Report Link"
    ' ============================================================
        ' Donation Sites:
            ' ------------------------------------------------------------
            ' 225 Gives
            ' ------------------------------------------------------------
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("E" & AR_225GivesRow), _
'                        Address:=SFReportLink_ManualImports, _
'                        TextToDisplay:="SF 'Manual Imports' Report Link"
'
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("E" & SF_225GivesRow), _
'                        Address:=SFReportLink_ManualImports, _
'                        TextToDisplay:="SF 'Manual Imports' Report Link"
                
            ' ------------------------------------------------------------
            ' AZ Gives
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & AR_AZGivesRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & SF_AZGivesRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
                        
            ' ------------------------------------------------------------
            ' Benevity
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & AR_BenevityRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & SF_BenevityRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
        
            ' ------------------------------------------------------------
            ' Click and Pledge
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & AR_ClickAndPledgeRow), _
                        Address:=SFReportLink_ClickandPledge, _
                        TextToDisplay:="SF Click and Pledge Report Link"
                
            ' ------------------------------------------------------------
            ' Crave It
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & AR_CraveItRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
                
            ' ------------------------------------------------------------
            ' Cyber Grants
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & AR_CyberGrantsRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & SF_CyberGrantsRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
                        
            ' ------------------------------------------------------------
            ' Fidelity
            ' ------------------------------------------------------------
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("E" & AR_FidelityGiftsRow), _
'                        Address:=SFReportLink_ManualImports, _
'                        TextToDisplay:="SF 'Manual Imports' Report Link"
'
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("E" & SF_FidelityGiftsRow), _
'                        Address:=SFReportLink_ManualImports, _
'                        TextToDisplay:="SF 'Manual Imports' Report Link"
                
            ' ------------------------------------------------------------
            ' Front Stream
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & AR_FrontStreamRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & SF_FrontStreamRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
                        
            ' ------------------------------------------------------------
            ' Give Gab/Big Gives/Bonterra Tech
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & AR_GiveGabRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & SF_GiveGabRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
                
            ' ------------------------------------------------------------
            ' NTX Giving
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & AR_NTXGivingRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & SF_NTXGivingRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
                
            ' ------------------------------------------------------------
            ' Your Cause
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & AR_YourCauseRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("E" & SF_YourCauseRow), _
                        Address:=SFReportLink_ManualImports, _
                        TextToDisplay:="SF 'Manual Imports' Report Link"
        
        ' Other Processes
            ' ------------------------------------------------------------
            ' Bank Statement Analysis
            ' ------------------------------------------------------------
                ' No Salesforce Link
                
            ' ------------------------------------------------------------
            ' Blackbaud AR Reconciliations (12010-FY)
            ' ------------------------------------------------------------
                ' No Salesforce Link
                
            ' ------------------------------------------------------------
            ' Blackbaud CRJs
            ' ------------------------------------------------------------
                ' No Salesforce Link
            
            ' ------------------------------------------------------------
            ' Blackbaud Website Deposits (School Level - Remittance)
            ' ------------------------------------------------------------
                 ' No Salesforce Link
                
            ' ------------------------------------------------------------
            ' Other Revenue (73010, 73010-9999) Reconciliation
            ' ------------------------------------------------------------
                ' No Salesforce Link
            
            ' ------------------------------------------------------------
            ' Report Consolidation (Consolidate any report)
            ' ------------------------------------------------------------
                ' No Salesforce Link
                
            ' ------------------------------------------------------------
            ' Security Deposit (SD) Reconciliation
            ' ------------------------------------------------------------
                 ' No Salesforce Link
    
    ' ============================================================
    ' Populate Column F: "Intacct Report Link"
    ' ============================================================
        ' Donation Sites:
            ' ------------------------------------------------------------
            ' 225 Gives
            ' ------------------------------------------------------------
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("F" & AR_225GivesRow), _
'                        Address:=IntacctLink, _
'                        TextToDisplay:="Intacct Website Link"
'
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("F" & SF_225GivesRow), _
'                        Address:=IntacctLink, _
'                        TextToDisplay:="Intacct Website Link"

            ' ------------------------------------------------------------
            ' AZ Gives
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & AR_AZGivesRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & SF_AZGivesRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

            ' ------------------------------------------------------------
            ' Benevity
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & AR_BenevityRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & SF_BenevityRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

            ' ------------------------------------------------------------
            ' Click and Pledge
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & AR_ClickAndPledgeRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

            ' ------------------------------------------------------------
            ' Crave It
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & AR_CraveItRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

            ' ------------------------------------------------------------
            ' Cyber Grants
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & AR_CyberGrantsRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & SF_CyberGrantsRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

            ' ------------------------------------------------------------
            ' Fidelity
            ' ------------------------------------------------------------
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("F" & AR_FidelityGiftsRow), _
'                        Address:=IntacctLink, _
'                        TextToDisplay:="Intacct Website Link"
'
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("F" & SF_FidelityGiftsRow), _
'                        Address:=IntacctLink, _
'                        TextToDisplay:="Intacct Website Link"

            ' ------------------------------------------------------------
            ' Front Stream
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & AR_FrontStreamRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & SF_FrontStreamRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"
                        
            ' ------------------------------------------------------------
            ' Give Gab/Big Gives/Bonterra Tech
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & AR_GiveGabRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & SF_GiveGabRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

            ' ------------------------------------------------------------
            ' NTX Giving
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & AR_NTXGivingRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & SF_NTXGivingRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

            ' ------------------------------------------------------------
            ' Your Cause
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & AR_YourCauseRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & SF_YourCauseRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

        ' Other Processes
            ' ------------------------------------------------------------
            ' Bank Statement Analysis
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & BankStatementAnalysisRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

            ' ------------------------------------------------------------
            ' Blackbaud AR Reconciliations (12010-FY)
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & BlackbaudARReconRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

            ' ------------------------------------------------------------
            ' Blackbaud CRJs
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & BlackbaudCRJsRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"
                        
            ' ------------------------------------------------------------
            ' Blackbaud Website Deposits (School Level - Remittance)
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & BlackbaudWebsiteDepositsRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

            ' ------------------------------------------------------------
            ' Other Revenue (73010, 73010-9999) Reconciliation
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & OtherRevenueReconRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"

            ' ------------------------------------------------------------
            ' Report Consolidation (Consolidate any report)
            ' ------------------------------------------------------------
                ' No Links

            ' ------------------------------------------------------------
            ' Security Deposit (SD) Reconciliation
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("F" & SDReconRow), _
                        Address:=IntacctLink, _
                        TextToDisplay:="Intacct Website Link"
                

    ' ============================================================
    ' Populate Column G: "SharePoint Report Link"
    ' ============================================================
        ' Donation Sites:
'            ' ------------------------------------------------------------
'            ' 225 Gives
'            ' ------------------------------------------------------------
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("G" & AR_225GivesRow), _
'                        Address:="", _
'                        TextToDisplay:=""
'
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("G" & SF_225GivesRow), _
'                        Address:="", _
'                        TextToDisplay:=""

            ' ------------------------------------------------------------
            ' AZ Gives
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("G" & AR_AZGivesRow), _
                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "CHARTERAR%2F" & _
                                      "Charter%20AR%20Shared%2F" & _
                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
                                          "AZ%20Gives&" & _
                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
                        TextToDisplay:="AZ Gives SharePoint Folder Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("G" & SF_AZGivesRow), _
                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "CHARTERAR%2F" & _
                                      "Charter%20AR%20Shared%2F" & _
                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
                                          "AZ%20Gives&" & _
                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
                        TextToDisplay:="AZ Gives SharePoint Folder Link"

            ' ------------------------------------------------------------
            ' Benevity
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("G" & AR_BenevityRow), _
                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "CHARTERAR%2F" & _
                                      "Charter%20AR%20Shared%2F" & _
                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
                                          "Benevity%2DAmer%20Online%20Giving&" & _
                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
                        TextToDisplay:="Benevity SharePoint Folder Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("G" & SF_BenevityRow), _
                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "CHARTERAR%2F" & _
                                      "Charter%20AR%20Shared%2F" & _
                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
                                          "Benevity%2DAmer%20Online%20Giving&" & _
                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
                        TextToDisplay:="Benevity SharePoint Folder Link"

            ' ------------------------------------------------------------
            ' Click and Pledge
            ' ------------------------------------------------------------
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("G" & AR_ClickAndPledgeRow), _
'                        Address:="", _
'                        TextToDisplay:=""

            ' ------------------------------------------------------------
            ' Crave It
            ' ------------------------------------------------------------
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("G" & AR_CraveItRow), _
'                        Address:="", _
'                        TextToDisplay:=""

            ' ------------------------------------------------------------
            ' Cyber Grants
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("G" & AR_CyberGrantsRow), _
                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "CHARTERAR%2F" & _
                                      "Charter%20AR%20Shared%2F" & _
                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
                                          "Cyber%20Grants&" & _
                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
                        TextToDisplay:="CyberGrants SharePoint Folder Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("G" & SF_CyberGrantsRow), _
                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "CHARTERAR%2F" & _
                                      "Charter%20AR%20Shared%2F" & _
                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
                                          "Cyber%20Grants&" & _
                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
                        TextToDisplay:="CyberGrants SharePoint Folder Link"

'            ' ------------------------------------------------------------
'            ' Fidelity
'            ' ------------------------------------------------------------
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("G" & AR_FidelityGiftsRow), _
'                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
'                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
'                                "id=%2Fsites%2F" & _
'                                    "CHARTERAR%2F" & _
'                                      "Charter%20AR%20Shared%2F" & _
'                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
'                                          "Fidelity%20Giving%20Marketplace%20%28Stripe%29&" & _
'                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
'                        TextToDisplay:="Fidelity SharePoint Folder Link"
'
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("G" & SF_FidelityGiftsRow), _
'                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
'                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
'                                "id=%2Fsites%2F" & _
'                                    "CHARTERAR%2F" & _
'                                      "Charter%20AR%20Shared%2F" & _
'                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
'                                          "Fidelity%20Giving%20Marketplace%20%28Stripe%29&" & _
'                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
'                        TextToDisplay:="Fidelity SharePoint Folder Link"

            ' ------------------------------------------------------------
            ' Front Stream
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("G" & AR_FrontStreamRow), _
                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "CHARTERAR%2F" & _
                                      "Charter%20AR%20Shared%2F" & _
                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
                                          "FrontStream&" & _
                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
                        TextToDisplay:="FrontStream SharePoint Folder Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("G" & SF_FrontStreamRow), _
                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "CHARTERAR%2F" & _
                                      "Charter%20AR%20Shared%2F" & _
                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
                                          "FrontStream&" & _
                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
                        TextToDisplay:="FrontStream SharePoint Folder Link"
            
            ' ------------------------------------------------------------
            ' Give Gab/Big Gives/Bonterra Tech
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("G" & AR_GiveGabRow), _
                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "CHARTERAR%2F" & _
                                      "Charter%20AR%20Shared%2F" & _
                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
                                          "Big%20Give-Give%20Gab&" & _
                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
                        TextToDisplay:="Give Gab/Big Gives SharePoint Folder Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("G" & SF_GiveGabRow), _
                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "CHARTERAR%2F" & _
                                      "Charter%20AR%20Shared%2F" & _
                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
                                          "Big%20Give-Give%20Gab&" & _
                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
                        TextToDisplay:="Give Gab/Big Gives SharePoint Folder Link"

            ' ------------------------------------------------------------
            ' NTX Giving
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("G" & AR_NTXGivingRow), _
                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "CHARTERAR%2F" & _
                                      "Charter%20AR%20Shared%2F" & _
                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
                                          "NTX%20Giving&" & _
                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
                        TextToDisplay:="NTX Giving SharePoint Folder Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("G" & SF_NTXGivingRow), _
                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "CHARTERAR%2F" & _
                                      "Charter%20AR%20Shared%2F" & _
                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
                                          "NTX%20Giving&" & _
                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
                        TextToDisplay:="NTX Giving SharePoint Folder Link"

            ' ------------------------------------------------------------
            ' Your Cause
            ' ------------------------------------------------------------
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("G" & AR_YourCauseRow), _
                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "CHARTERAR%2F" & _
                                      "Charter%20AR%20Shared%2F" & _
                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
                                          "Your%20Cause&" & _
                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
                        TextToDisplay:="Your Cause SharePoint Folder Link"
                        
                wsSelectionPage.Hyperlinks.Add _
                        Anchor:=wsSelectionPage.Range("G" & SF_YourCauseRow), _
                        Address:="https://basised.sharepoint.com/sites/CHARTERAR/" & _
                                "Charter%20AR%20Shared/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "CHARTERAR%2F" & _
                                      "Charter%20AR%20Shared%2F" & _
                                        "1%2E%20Donation%20Sites%20Reports%2F" & _
                                          "Your%20Cause&" & _
                                "viewid=2bd4d68b%2Ded02%2D4793%2D8829%2D1d0c712f72c4", _
                        TextToDisplay:="Your Cause SharePoint Folder Link"

        ' Other Processes
            ' ------------------------------------------------------------
            ' Bank Statement Analysis
            ' ------------------------------------------------------------
                ' FY26 (School Year: 2025-2026)
                    wsSelectionPage.Hyperlinks.Add _
                            Anchor:=wsSelectionPage.Range("G" & BankStatementAnalysisRow), _
                            Address:="https://basised.sharepoint.com/sites/BASISCHARTERACCOUNTING/" & _
                                "Charter%20Month%20End%20FY%202026/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "BASISCHARTERACCOUNTING%2F" & _
                                      "Charter%20Month%20End%20FY%202026%2F" & _
                                        "Banking%20FY26%2F" & _
                                          "Bank%20Statements%20%26%20Reconciliations%20FY26&" & _
                                "viewid=563786cb%2Daa70%2D4bbb%2Da781%2D9041a3da092f", _
                            TextToDisplay:="BASIS 'FY26' Bank Statements Folder Link"
                    
'                ' FY25 (School Year: 2024-2025)
'                    wsSelectionPage.Hyperlinks.Add _
'                            Anchor:=wsSelectionPage.Range("G" & BankStatementAnalysisRow), _
'                            Address:="https://basised.sharepoint.com/sites/BASISCHARTERACCOUNTING/" & _
'                                "Charter%20Month%20End%20FY%202025/Forms/AllItems.aspx?" & _
'                                "id=%2Fsites%2F" & _
'                                    "BASISCHARTERACCOUNTING%2F" & _
'                                      "Charter%20Month%20End%20FY%202025%2F" & _
'                                        "Banking%20FY25%2F" & _
'                                          "Bank%20Statements%20%26%20Reconciliations%20FY25&" & _
'                                "viewid=55c4d661%2D0533%2D486b%2D93b0%2D5d812a833505", _
'                            TextToDisplay:="BASIS 'FY25' Bank Statements Folder Link"
'
'                ' FY24 (School Year: 2023-2024)
'                    wsSelectionPage.Hyperlinks.Add _
'                            Anchor:=wsSelectionPage.Range("G" & BankStatementAnalysisRow), _
'                            Address:="https://basised.sharepoint.com/sites/BASISCHARTERACCOUNTING/" & _
'                                "Charter%20Month%20End%20FY%202024/Forms/AllItems.aspx?" & _
'                                "id=%2Fsites%2F" & _
'                                "BASISCHARTERACCOUNTING%2F" & _
'                                  "Charter%20Month%20End%20FY%202024%2F" & _
'                                    "Banking%20FY24%2F" & _
'                                      "Bank%20Statements%20and%20Reconciliations%20FY24&" & _
'                                "viewid=66b9eaa3%2D3110%2D4445%2Dba4c%2D91e9f420c459", _
'                            TextToDisplay:="BASIS 'FY24' Bank Statements Folder Link"

            ' ------------------------------------------------------------
            ' Blackbaud AR Reconciliations (12010-FY)
            ' ------------------------------------------------------------
'                ' FY26 (School Year: 2025-2026)
'                    wsSelectionPage.Hyperlinks.Add _
'                            Anchor:=wsSelectionPage.Range("G" & BlackbaudARReconRow), _
'                            Address:="", _
'                            TextToDisplay:="BASIS 'FY26' Inception to Date Reports Link"
                        
                ' FY25 (School Year: 2024-2025)
                    wsSelectionPage.Hyperlinks.Add _
                            Anchor:=wsSelectionPage.Range("G" & BlackbaudARReconRow), _
                            Address:="https://basised.sharepoint.com/sites/BASISCHARTERACCOUNTING/" & _
                                "Charter%20Month%20End%20FY%202025/Forms/AllItems.aspx?" & _
                                "id=%2Fsites%2F" & _
                                    "BASISCHARTERACCOUNTING%2F" & _
                                      "Charter%20Month%20End%20FY%202025%2F" & _
                                        "Blackbaud%20Inception%20To%20Date%20Reports%20%2D%20ALL%20Locations&" & _
                                "viewid=55c4d661%2D0533%2D486b%2D93b0%2D5d812a833505", _
                            TextToDisplay:="BASIS 'FY25' Inception to Date Reports Link"

            ' ------------------------------------------------------------
            ' Blackbaud CRJs
            ' ------------------------------------------------------------
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("G" & BlackbaudCRJsRow), _
'                        Address:="", _
'                        TextToDisplay:=""
                        
            ' ------------------------------------------------------------
            ' Blackbaud Website Deposits (School Level - Remittance)
            ' ------------------------------------------------------------
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("G" & BlackbaudWebsiteDepositsRow), _
'                        Address:="", _
'                        TextToDisplay:=""

            ' ------------------------------------------------------------
            ' Other Revenue (73010, 73010-9999) Reconciliation
            ' ------------------------------------------------------------
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("G" & OtherRevenueReconRow), _
'                        Address:="", _
'                        TextToDisplay:=""

            ' ------------------------------------------------------------
            ' Report Consolidation (Consolidate any report)
            ' ------------------------------------------------------------
                ' No SharePoint Links

            ' ------------------------------------------------------------
            ' Security Deposit (SD) Reconciliation
            ' ------------------------------------------------------------
'                wsSelectionPage.Hyperlinks.Add _
'                        Anchor:=wsSelectionPage.Range("G" & SDReconRow), _
'                        Address:="", _
'                        TextToDisplay:=""

'    ' ============================================================
'    ' Populate Column H: "Description for Converter"
'    ' ============================================================
'        ' Donation Sites:
'            ' ------------------------------------------------------------
'            ' 225 Gives
'            ' ------------------------------------------------------------
''                wsSelectionPage.Range("H" & AR_225GivesRow).Value = ""
''                wsSelectionPage.Range("H" & SF_225GivesRow).Value = ""
'
'
'            ' ------------------------------------------------------------
'            ' AZ Gives
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & AR_AZGivesRow).Value = ""
'
'
'                wsSelectionPage.Range("H" & SF_AZGivesRow).Value = _
'                        "This macro processes AZ Gives donation reports and prepares them for a Salesforce import. " & _
'                        "It automates file selection, Disbursement ID assignment, and prize tracking, while generating " & _
'                        "both a cleaned report and an import-ready file. The process includes school name mapping, " & _
'                        "conditional formatting, and data validation to ensure consistency and accuracy across reporting."
'
'            ' ------------------------------------------------------------
'            ' Benevity
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & AR_BenevityRow).Value = ""
'
'
'                wsSelectionPage.Range("H" & SF_BenevityRow).Value = _
'                        "This macro automates the processing of Benevity donation reports for Salesforce import. " & _
'                        "It standardizes disbursement and transaction data, applies school mappings, and consolidates all reports into a clean, unified format. " & _
'                        "The process includes setup guidance, file validation, renaming, and creation of import-ready worksheets for accurate reporting."
'
'            ' ------------------------------------------------------------
'            ' Click and Pledge
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & AR_ClickAndPledgeRow).Value = ""
'
'            ' ------------------------------------------------------------
'            ' Crave It
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & AR_CraveItRow).Value = _
'                        "This macro automates the consolidation and analysis of Crave-It revenue share data. " & _
'                        "It imports and validates the Unpaid Accounts, Menu Library, and Revenue Share reports, aligns all menu and pricing structures, " & _
'                        "and calculates revenue share allocations for each school. " & _
'                        "The process produces consolidated datasets, exception reports for multi-meal and multi-drink entries, and a dynamic pivot table " & _
'                        "summarizing total revenue, reimbursements, and allocation percentages."
'
'            ' ------------------------------------------------------------
'            ' Cyber Grants
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & AR_CyberGrantsRow).Value = ""
'
'
'                wsSelectionPage.Range("H" & SF_CyberGrantsRow).Value = _
'                        "This two-part macro automates the processing of CyberGrants donation reports for Salesforce import. " & _
'                        "Part 1 creates a setup worksheet for entering disbursement details, while " & _
'                        "Part 2 consolidates multiple CyberGrants files, validates headers, applies school mappings, and assigns disbursement IDs. " & _
'                        "The process generates consolidated datasets, exception reports for duplicate disbursements, and Salesforce import files " & _
'                        "for both electronic and check payments, ensuring accurate and consistent reporting."
'
'            ' ------------------------------------------------------------
'            ' Fidelity
'            ' ------------------------------------------------------------
''                wsSelectionPage.Range("H" & AR_FidelityGiftsRow).Value = ""
''                wsSelectionPage.Range("H" & SF_FidelityGiftsRow).Value = ""
'
'            ' ------------------------------------------------------------
'            ' Front Stream
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & AR_FrontStreamRow).Value = ""
'
'
'                wsSelectionPage.Range("H" & SF_FrontStreamRow).Value = _
'                        "This macro automates the processing of FrontStream donation reports for Salesforce import. " & _
'                        "It consolidates multiple files into a unified dataset, validates column headers, and generates standardized transaction IDs for each record. " & _
'                        "The process applies school mappings, creates a Salesforce import file with donor and payment details, and applies conditional formatting " & _
'                        "and data validation to ensure complete and accurate reporting across all FrontStream disbursements."
'
'            ' ------------------------------------------------------------
'            ' Give Gab/Big Gives/Bonterra Tech
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & AR_GiveGabRow).Value = ""
'
'
'                wsSelectionPage.Range("H" & SF_GiveGabRow).Value = _
'                        "This macro automates the processing of GiveGab, BonterraTech, and Big Give donation reports for Salesforce import. " & _
'                        "It validates report headers, filters data by date, and organizes disbursements into dedicated worksheets for analysis, split payouts, " & _
'                        "checks, pending payments, and school mappings. The process generates an import-ready Salesforce file with standardized disbursement IDs, " & _
'                        "donor details, and campaign information, ensuring accurate and consistent reporting."
'
'            ' ------------------------------------------------------------
'            ' NTX Giving
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & AR_NTXGivingRow).Value = ""
'
'
'                wsSelectionPage.Range("H" & SF_NTXGivingRow).Value = _
'                        "This macro processes North Texas Giving Day donation reports for Salesforce import. " & _
'                        "It validates the selected file, adds disbursement IDs, and accounts for fee reimbursements and prize adjustments. " & _
'                        "The process generates a new report with updated totals, creates a Salesforce import file, and applies school mappings and data validation " & _
'                        "to ensure accuracy and consistency."
'
'            ' ------------------------------------------------------------
'            ' Your Cause
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & AR_YourCauseRow).Value = ""
'
'
'                wsSelectionPage.Range("H" & SF_YourCauseRow).Value = _
'                        "This macro processes YourCause donation reports for Salesforce import. " & _
'                        "It validates and consolidates all files in the selected folder, assigns disbursement IDs, and splits each disbursement into separate files. " & _
'                        "The process creates a Salesforce import file with mapped school names, applies data validation and conditional formatting, and ensures accurate, " & _
'                        "standardized reporting."
'
'        ' Other Processes
'            ' ------------------------------------------------------------
'            ' Bank Statement Analysis
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & BankStatementAnalysisRow).Value = _
'                        "This macro processes monthly bank statement files to extract, standardize, and consolidate account data. " & _
'                        "It identifies each file’s account, month, and year, organizes the information into a unified format, and prepares " & _
'                        "a consolidated dataset for reconciliation and analysis. The process also generates a 'Checks' worksheet and " & _
'                        "a 'Consolidated Bank Data' sheet for comparison, ensuring accurate reporting across all accounts."
'
'            ' ------------------------------------------------------------
'            ' Blackbaud AR Reconciliations (12010-FY)
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & BlackbaudARReconRow).Value = ""
'
'            ' ------------------------------------------------------------
'            ' Blackbaud CRJs
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & BlackbaudCRJsRow).Value = _
'                        "This macro processes Blackbaud Enterprise Remittance Reports to prepare Cash Journal Receipt (CRJ) files for Intacct import. " & _
'                        "It validates report headers, standardizes data formats, and performs automated school code mapping. " & _
'                        "The process generates a detailed remittance analysis, daily breakdown worksheets, and an Intacct import file with accurate account, " & _
'                        "location, and school mappings for streamlined reconciliations."
'
'            ' ------------------------------------------------------------
'            ' Blackbaud Website Deposits (School Level - Remittance)
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & BlackbaudWebsiteDepositsRow).Value = _
'                        "This macro processes School Level Remittance Reports exported from Blackbaud. " & _
'                        "It validates report structure, extracts key details such as school ID and fiscal year, and applies formatting and conditional rules to clean the data. " & _
'                        "The process separates transactions from totals into structured worksheets, applies automated school lookups, and prepares the report " & _
'                        "for reconciliation and further financial analysis."
'
'            ' ------------------------------------------------------------
'            ' Other Revenue (73010, 73010-9999) Reconciliation
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & OtherRevenueReconRow).Value = _
'                        "This macro reconciles Blackbaud Inception-to-Date reports with Intacct General Ledger data for accounts 73010 and 73010-9999. " & _
'                        "It validates each report, extracts and reformats key columns, and applies prefix-based mapping to align fees and GL accounts. " & _
'                        "The process generates pivot comparisons, variance analyses, and formatted reports to identify discrepancies between systems and " & _
'                        "ensure accurate revenue recognition."
'
'            ' ------------------------------------------------------------
'            ' Report Consolidation (Consolidate any report)
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & ReportConsolidationRow).Value = _
'                        "This two-part macro is designed to combine multiple reports that share the same column headers into one consolidated table. " & _
'                        "In Part 1, the exact column headers from the reports are entered so the macro can identify matching fields. " & _
'                        "In Part 2, all files that meet those criteria are located and combined from the selected folder (the folder name must include 'Consolidate'). " & _
'                        "The result is a single, organized worksheet containing all report data in one place."
'
'            ' ------------------------------------------------------------
'            ' Security Deposit (SD) Reconciliation
'            ' ------------------------------------------------------------
'                wsSelectionPage.Range("H" & SDReconRow).Value = _
'                        "This macro reconciles Student Deposit activity between Blackbaud Inception-to-Date reports and Intacct GL account 43020. " & _
'                        "It validates report headers, extracts and reformats key data, and uses a dynamic lookup to identify all Security Deposit–related fees. " & _
'                        "The process builds comparison pivot tables, calculates variances, and generates a detailed summary showing unmatched records for accurate " & _
'                        "tracking and reconciliation."
'

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''----------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' Create Macro Buttons ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''----------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    ' Create the buttons for each of the converters.
    ' ============================================================
        ' ------------------------------------------------------------
        ' 225 Gives
        ' ------------------------------------------------------------
'           ' AR Team's Macro
'                Set AR_225GivesButton = wsSelectionPage.Buttons.Add(0, ((AR_225GivesRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
'
'                With AR_225GivesButton
'                    .Caption = "Click here to run the 'AR' 225 Gives Macro"
'                    .OnAction = "A225Gives.AR_225Gives_Converter"
'                    .Font.Size = ConvertersFontSize
'                    .Font.Bold = True
'                    .Font.Color = ButtonColor_AR
'                End With
'
'            ' SF Team's Macro
'                Set SF_225GivesButton = wsSelectionPage.Buttons.Add(0, ((SF_225GivesRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
'
'                With SF_225GivesButton
'                    .Caption = "Click here to run the 'SF' 225 Gives Macro"
'                    .OnAction = "A225Gives.SF_225Gives_Converter"
'                    .Font.Size = ConvertersFontSize
'                    .Font.Bold = True
'                    .Font.Color = ButtonColor_SF
'                End With
                
        ' ------------------------------------------------------------
        ' AZ Gives
        ' ------------------------------------------------------------
            ' AR Team's Macro
                Set AR_AZGivesButton = wsSelectionPage.Buttons.Add(0, ((AR_AZGivesRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With AR_AZGivesButton
                    .Caption = "Click here to run the 'AR' AZ Gives Macro"
                    .OnAction = "AZGives.AZGives_AR_Converter"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_AR
                End With
                
            ' SF Team's Macro
                Set SF_AZGivesButton = wsSelectionPage.Buttons.Add(0, ((SF_AZGivesRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With SF_AZGivesButton
                    .Caption = "Click here to run the 'SF' AZ Gives Macro"
                    .OnAction = "AZGives.AZGives_SF_Converter"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_SF
                End With
                
                
        ' ------------------------------------------------------------
        ' Benevity
        ' ------------------------------------------------------------
            ' AR Team's Macro
                Set AR_BenevityButton = wsSelectionPage.Buttons.Add(0, ((AR_BenevityRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With AR_BenevityButton
                    .Caption = "Click here to run the 'AR' Benevity Macro"
                    .OnAction = "Benevity.Benevity_AR_Converter"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_AR
                End With
                
            ' SF Team's Macro
                Set SF_BenevityButton = wsSelectionPage.Buttons.Add(0, ((SF_BenevityRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With SF_BenevityButton
                    .Caption = "Click here to run the 'SF' Benevity Macro"
                    .OnAction = "Benevity.Benevity_SF_Converter"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_SF
                End With
        
        ' ------------------------------------------------------------
        ' Click & Pledge
        ' ------------------------------------------------------------
            ' AR Team's Macro
                Set AR_ClickAndPledgeButton = wsSelectionPage.Buttons.Add(0, ((AR_ClickAndPledgeRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With AR_ClickAndPledgeButton
                    .Caption = "Click here to run the 'AR' Click and Pledge Macro"
                    .OnAction = "Click_and_Pledge.Click_and_Pledge_AR_Converter"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_AR
                End With
                
        ' ------------------------------------------------------------
        ' Crave It
        ' ------------------------------------------------------------
            ' AR Team's Macro
                Set AR_CraveItButton = wsSelectionPage.Buttons.Add(0, ((AR_CraveItRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With AR_CraveItButton
                    .Caption = "Click here to run the 'AR' Crave It Macro"
                    .OnAction = "Crave_It.CraveIt_Converter"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_AR
                End With

        ' ------------------------------------------------------------
        ' Cyber Grants
        ' ------------------------------------------------------------
            ' AR Team's Macro
                Set AR_CyberGrantsButton = wsSelectionPage.Buttons.Add(0, ((AR_CyberGrantsRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With AR_CyberGrantsButton
                    .Caption = "Click here to run the 'AR' Cyber Grants Macro"
                    .OnAction = "CyberGrants.CyberGrants_AR_Converter"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_AR
                End With
                
            ' SF Team's Macro
                Set SF_CyberGrantsButton = wsSelectionPage.Buttons.Add(0, ((SF_CyberGrantsRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With SF_CyberGrantsButton
                    .Caption = "Click here to run the 'SF' Cyber Grants Macro"
                    .OnAction = "CyberGrants.CyberGrants_SF_Part1"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_SF
                End With
                
        ' ------------------------------------------------------------
        ' Fidelity Gifts
        ' ------------------------------------------------------------
'            ' AR Team's Macro
'                Set AR_FidelityGiftsButton = wsSelectionPage.Buttons.Add(0, ((AR_FidelityGiftsRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
'
'                With AR_FidelityGiftsButton
'                    .Caption = "Click here to run the 'AR' Fidelity Gifts Macro"
'                    .OnAction = "Fidelity.Fidelity_AR_Converter"
'                    .Font.Size = ConvertersFontSize
'                    .Font.Bold = True
'                    .Font.Color = ButtonColor_AR
'                End With
'
'            ' SF Team's Macro
'                Set SF_FidelityGiftsButton = wsSelectionPage.Buttons.Add(0, ((SF_FidelityGiftsRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
'
'                With SF_FidelityGiftsButton
'                    .Caption = "Click here to run the 'SF' Fidelity Gifts Macro"
'                    .OnAction = "Fidelity.Fidelity_SF_Converter"
'                    .Font.Size = ConvertersFontSize
'                    .Font.Bold = True
'                    .Font.Color = ButtonColor_SF
'                End With
                
        ' ------------------------------------------------------------
        ' Front Stream
        ' ------------------------------------------------------------
            ' AR Team's Macro
                Set AR_FrontStreamButton = wsSelectionPage.Buttons.Add(0, ((AR_FrontStreamRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With AR_FrontStreamButton
                    .Caption = "Click here to run the 'AR' Front Stream Macro"
                    .OnAction = "FrontStream_AR_Converter"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_AR
                End With
                
            ' SF Team's Macro
                Set SF_FrontStreamButton = wsSelectionPage.Buttons.Add(0, ((SF_FrontStreamRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With SF_FrontStreamButton
                    .Caption = "Click here to run the 'SF' Front Stream Macro"
                    .OnAction = "FrontStream.FrontStream_SF_Converter"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_SF
                End With
        
        ' ------------------------------------------------------------
        ' Give Gab/Big Gives/Bonterra Tech
        ' ------------------------------------------------------------
            ' AR Team's Macro
                Set AR_GiveGabButton = wsSelectionPage.Buttons.Add(0, ((AR_GiveGabRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With AR_GiveGabButton
                    .Caption = "Click here to run the 'AR' Give Gab Macro"
                    .OnAction = "GiveGab_BigGives_BonterraTech.GiveGab_AR_Converter"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_AR
                End With
                
            ' SF Team's Macro
                Set SF_GiveGabButton = wsSelectionPage.Buttons.Add(0, ((SF_GiveGabRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With SF_GiveGabButton
                    .Caption = "Click here to run the 'SF' Give Gab Macro"
                    .OnAction = "GiveGab_BigGives_BonterraTech.GiveGab_SF"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_SF
                End With
        
        
        ' ------------------------------------------------------------
        ' NTX Giving Day
        ' ------------------------------------------------------------
            ' AR Team's Macro
                Set AR_NTXGivingButton = wsSelectionPage.Buttons.Add(0, ((AR_NTXGivingRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With AR_NTXGivingButton
                    .Caption = "Click here to run the 'AR' NTX Giving Day Macro"
                    .OnAction = "NTXGiving.NTXGiving_AR_Converter"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_AR
                End With
                
            ' SF Team's Macro
                Set SF_NTXGivingButton = wsSelectionPage.Buttons.Add(0, ((SF_NTXGivingRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With SF_NTXGivingButton
                    .Caption = "Click here to run the 'SF' NTX Giving Day Macro"
                    .OnAction = "NTXGiving.NTXGiving_SF"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_SF
                End With
        
        ' ------------------------------------------------------------
        ' Your Cause
        ' ------------------------------------------------------------
            ' AR Team's Macro
                Set AR_YourCauseButton = wsSelectionPage.Buttons.Add(0, ((AR_YourCauseRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With AR_YourCauseButton
                    .Caption = "Click here to run the 'AR' Your Cause Macro"
                    .OnAction = "YourCause.YourCause_AR_Converter"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_AR
                End With
                
            ' SF Team's Macro
                Set SF_YourCauseButton = wsSelectionPage.Buttons.Add(0, ((SF_YourCauseRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With SF_YourCauseButton
                    .Caption = "Click here to run the 'SF' Your Cause Macro"
                    .OnAction = "YourCause.YourCause_SF"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_SF
                End With
                
                
                
        ' ------------------------------------------------------------
        ' Bank Statement Analysis
        ' ------------------------------------------------------------
            ' Accounting Team's Macro
                Set BankStatementAnalysisButton = wsSelectionPage.Buttons.Add(0, ((BankStatementAnalysisRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With BankStatementAnalysisButton
                    .Caption = "Click here to run the Bank Statement Analysis Macro"
                    .OnAction = "Bank_Statements.Analysis_Part1"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_Accounting
                End With
        
        ' ------------------------------------------------------------
        ' Blackbaud AR Reconciliations (12010-FY)
        ' ------------------------------------------------------------
            ' Accounting Team's Macro
                Set BlackbaudARReconButton = wsSelectionPage.Buttons.Add(0, ((BlackbaudARReconRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With BlackbaudARReconButton
                    .Caption = "Click here to run the Blackbaud AR Reconciliation Macro"
                    .OnAction = "Blackbaud_AR_Recon.BlackbaudARRecon_Part1"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_Accounting
                End With
        
        ' ------------------------------------------------------------
        ' Blackbaud CRJ Macro
        ' ------------------------------------------------------------
            ' AR Team's Macro
                Set BlackbaudCRJsButton = wsSelectionPage.Buttons.Add(0, ((BlackbaudCRJsRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With BlackbaudCRJsButton
                    .Caption = "Click here to run the Blackbaud CRJ Macro"
                    .OnAction = "Blackbaud_Reports_Analysis.Blackbaud_CRJs"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_AR
                End With
        
        
        ' ------------------------------------------------------------
        ' Blackbaud Website Deposits (School Level - Remittance)
        ' ------------------------------------------------------------
            ' Accounting Team's Macro
                Set BlackbaudWebsiteDepositsButton = wsSelectionPage.Buttons.Add(0, ((BlackbaudWebsiteDepositsRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With BlackbaudWebsiteDepositsButton
                    .Caption = "Click here to run the Blackbaud Website Deposits Macro"
                    .OnAction = "Blackbaud_Reports_Analysis.SchoolLevel_RemittanceReport"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_Accounting
                End With
        
        ' ------------------------------------------------------------
        ' Other Revenue (73010, 73010-9999) Reconciliation
        ' ------------------------------------------------------------
            ' AR Team's Macro
                Set OtherRevenueReconButton = wsSelectionPage.Buttons.Add(0, ((OtherRevenueReconRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With OtherRevenueReconButton
                    .Caption = "Click here to run the 'Other Revenue' Reconciliation Macro"
                    .OnAction = "OtherRevenue_Recon.OtherRevenue_BBvsIntacct_Recon"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_AR
                End With
                
        ' ------------------------------------------------------------
        ' Report Consolidation (Consolidate any report)
        ' ------------------------------------------------------------
                Set ReportConsolidationButton = wsSelectionPage.Buttons.Add(0, ((ReportConsolidationRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With ReportConsolidationButton
                    .Caption = "Click here to run the Report Consolidation Macro"
                    .OnAction = "Report_Consolidation.ConsolidationConverter_Part1"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_Universal
                End With
            
        ' ------------------------------------------------------------
        ' Security Deposit (SD) Reconciliation
        ' ------------------------------------------------------------
            ' Accounting Team's Macro
                Set SDReconButton = wsSelectionPage.Buttons.Add(0, ((SDReconRow - 1) * HeightofRow), 375, (HeightofRow - 0.5))
                
                With SDReconButton
                    .Caption = "Click here to run the Security Deposits Reconciliation Macro"
                    .OnAction = "SecurityDeposits_Recon.SD_Recon"
                    .Font.Size = ConvertersFontSize
                    .Font.Bold = True
                    .Font.Color = ButtonColor_Accounting
                End With
               
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''----------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' Final Formatting and Application Reset ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''----------------------------------------''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ============================================================
    ' Make final adjustments to column widths
    ' ============================================================
        wsSelectionPage.Columns("B:I").AutoFit
        wsSelectionPage.Columns("H").ColumnWidth = 150
        wsSelectionPage.Cells.WrapText = True
    
    ' ============================================================
    ' Freeze top row and first column
    ' ============================================================
        ' Freeze both the top row and first column for easy navigation
            With wsSelectionPage
                .Activate
                .Range("B2").Select
            End With
            
            ActiveWindow.FreezePanes = True
    
    ' ============================================================
    ' Find the Last Row
    ' ============================================================
        SelectionPageLastRow = wsSelectionPage.Cells(wsSelectionPage.Rows.Count, "B").End(xlUp).Row
    
    ' ============================================================
    ' Apply alternating row colors
    ' ============================================================
        For SelectionRow = 2 To SelectionPageLastRow
            Select Case SelectionRow Mod 2
                ' Even Rows
                Case 0
                    wsSelectionPage.Range("A" & SelectionRow & ":H" & SelectionRow).Interior.Color = RGB(15, 15, 15)
                    
                ' Odd Rows
                Case 1
                    wsSelectionPage.Range("A" & SelectionRow & ":H" & SelectionRow).Interior.Color = RGB(25, 25, 25)
            End Select
        Next SelectionRow
    
    ' ============================================================
    ' Restore Excel application settings to default
    ' ============================================================
        Application.Calculation = xlAutomatic
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.StatusBar = False
        
    ' ============================================================
    ' Reset view to top-left of sheet
    ' ============================================================
        wsSelectionPage.Range("A1").Select
        
    ' ============================================================
    ' Display confirmation message after reset is complete
    ' ============================================================
'        MsgBox Title:="Reset Successfully Completed", _
'                Prompt:="The Converter Selection Page has been successfully recreated.", _
'                Buttons:=vbInformation

End Sub



