Sub Add_All_Buttons()
' Last Updated: 2025.08.25

Dim ws As Worksheet
Dim SheetExists As Boolean
Dim wsAddReports As Worksheet
Dim btnAddSingleReport As Button
Dim btnAddMultipleReports As Button
Dim btnReset As Button
Dim btnUpdate As Button
'Dim btnSplitOutReports As Button

' Set the variable 'SheetExists' to false: saying that the 'Add Reports' worksheet does not yet exist.
    SheetExists = False

' Check if the "Selection Page" worksheet already exists. If it does, then no action is necessary.
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Selection Page" Then
            SheetExists = True
            Exit Sub
        End If
    Next ws

    ' If the "Add or Split Reports" worksheet does not yet exist, create it.
        ' Add a new worksheet to the beginning of the workbook.
            Set wsAddReports = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        
        ' Change the name of the worksheet to 'Selection Page'
            wsAddReports.Name = "Selection Page"
            
        ' Color the entire background black.
            wsAddReports.Cells.Interior.Color = RGB(0, 0, 0)
            
        ' Create 4 buttons:
            ' 1. 'Add A Single Report' Button
                Set btnAddSingleReport = wsAddReports.Buttons.Add(Left:=100, Top:=25, Width:=900, Height:=100)
                With btnAddSingleReport
                    .Caption = "CLICK HERE TO ADD ONE REPORT"
                    .OnAction = "Add_Single_CraveIt_Report"
                    .Name = "Button_AddSingleReport"
                    .Font.Size = 55
                    .Font.Color = RGB(200, 200, 0)
                    .Font.Bold = True
                End With
                
            ' 2. 'Add Multiple Reports' Button
                Set btnAddMutlipleReports = wsAddReports.Buttons.Add(Left:=100, Top:=150, Width:=900, Height:=100)
                With btnAddMutlipleReports
                    .Caption = "CLICK HERE TO ADD MULTIPLE REPORTS"
                    .OnAction = "Add_Multiple_CraveIt_Reports"
                    .Name = "Button_AddMutlipleReports"
                    .Font.Size = 45
                    .Font.Color = RGB(18, 154, 34)
                    .Font.Bold = True
                End With

            ' 3. 'Reset' Button
                Set btnReset = wsAddReports.Buttons.Add(Left:=100, Top:=275, Width:=900, Height:=100)
                With btnReset
                    .Caption = "CLICK HERE TO RESET WORKBOOK"
                    .OnAction = "Start_Reset"
                    .Name = "Button_Reset"
                    .Font.Size = 55
                    .Font.Color = RGB(255, 0, 0)
                    .Font.Bold = True
                End With
                
            ' 4. 'Update' Button
                Set btnUpdate = wsAddReports.Buttons.Add(Left:=100, Top:=400, Width:=900, Height:=100)
                With btnUpdate
                    .Caption = "CLICK HERE TO UPDATE MENU LIST"
                    .OnAction = "Update_MealsLookup"
                    .Name = "Button_Update"
                    .Font.Size = 55
                    .Font.Color = RGB(75, 75, 75)
                    .Font.Bold = True
                End With
                
            ' 5. 'Split Out Reports' Button
'                Set btnSplitOutReports = wsAddReports.Buttons.Add(Left:=100, Top:=525, Width:=900, Height:=100)
'                With btnSplitOutReports
'                    .Caption = "CLICK HERE TO SPLIT OUT REPORTS"
'                    .OnAction = "SplitOut_Reports"
'                    .Name = "Button_SplitReports"
'                    .Font.Size = 55
'                    .Font.Color = RGB(255, 153, 51)
'                    .Font.Bold = True
'                End With
                
   
        ' Protect the worksheet
            wsAddReports.Protect
        
        ' Add the 'Meals Lookup' worksheet.
            MealsLookup_1

End Sub

Sub Start_Reset()
' Last Updated: 2025.07.02

Dim UserResponse As VbMsgBoxResult
Dim ws As Worksheet
Dim wsReset As Worksheet

' Turn off Alerts
    Application.DisplayAlerts = False

' Give the user a message letting them know all data will be deleted if they proceed
    UserResponse = MsgBox("WARNING: Using this macro, will delete all worksheets and data in this workbook." & vbCrLf & vbCrLf & "Do you want to continue?", vbYesNo + vbExclamation, "Confirm Deletion")

' Check the User's response
    ' If they click 'No' - exit the sub
        If UserResponse = vbNo Then
            Exit Sub
        End If
    
    ' If they proceed, change the 'Selection Page' worksheet's name, and run the 'Add_All_Buttons' Macro.
        ' Check if the "Selection Page" worksheet exists, if it does, rename it to "Reset-" and the current date and time
            For Each ws In ThisWorkbook.Worksheets
                If ws.Name = "Selection Page" Then
                    Set wsReset = ws
                    wsReset.Name = "Reset-" & Format(Now(), "YYYY.MM.DD-HH.MM.SS")
                End If
            Next ws
        
        ' Delete all worksheets except 'wsReset'
            For Each ws In ThisWorkbook.Worksheets
                If ws.Name <> wsReset.Name Then
                    ws.Delete
                End If
            Next ws
            
        ' Run the 'Add_All_Buttons' Macro.
            Add_All_Buttons
        
        ' Delete 'wsReset'
            wsReset.Delete
            
' Turn back on Alerts
    Application.DisplayAlerts = True
    
    ' Provide the user a message, letting them know the reset is complete.
        MsgBox "Reset completed.", vbInformation, "Reset Complete"


End Sub

