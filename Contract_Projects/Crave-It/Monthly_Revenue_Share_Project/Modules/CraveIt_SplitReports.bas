Sub SplitOut_Reports()
' Last Updated: 2025.07.02

Dim wbMacro As Workbook
Dim UserResponse As VbMsgBoxResult
Dim wbNew As Workbook
Dim ws As Worksheet
Dim wsNew As Worksheet
Dim RevenueShareCol As Long
Dim LastRow As Long

' Set this workbook to the 'wbMacro'
    Set wbMacro = ThisWorkbook

' Turn off the screen updating
    Application.ScreenUpdating = False


' Give the user a message letting them know all data will be deleted if they proceed
    UserResponse = MsgBox("NOTE: Using this macro, will split out all of the worksheets from this workbook." & vbCrLf & vbCrLf & "Do you want to continue?", vbYesNo + vbExclamation, "Confirm Split")

' Check the User's response
    ' If they click 'No' - exit the sub
        If UserResponse = vbNo Then
            Exit Sub
        End If

' Unhide all worksheets before splitting
    For Each ws In wbMacro.Worksheets
        ws.Visible = xlSheetVisible
    Next ws

' If they click 'Yes' - Split out all worksheets that are not "Selection Page"
    ' Check to make sure there are at least two worksheets
        If ThisWorkbook.Worksheets.Count = 1 Then
            MsgBox "No worksheets to split out were found.", vbInformation
            Exit Sub
        End If

    ' Create a new workbook
        Set wbNew = Workbooks.Add(xlWBATWorksheet)

    ' Set the first sheet of the 'wbNew' workbook to the variable 'wsNew' and rename it based on the current date and time
        Set wsNew = wbNew.Sheets(1)
        wsNew.Name = Format(Now(), "YYYY.MM.DD-HH.MM.SS")

    ' Loop through each worksheet in the macro workbook
        For Each ws In wbMacro.Worksheets
            If ws.Name <> "Selection Page" Then
            ' Copy the entire worksheet into the new workbook
                ws.Copy After:=wbNew.Sheets(wbNew.Sheets.Count)
            
            ' Set the new worksheet to the variable 'wsNew'
                Set wsNew = ActiveSheet

            ' Find the "Revenue Share" column in row 1
                On Error Resume Next ' Prevents errors if not found
                RevenueShareCol = wsNew.Rows(1).Find(What:="Revenue Share", LookIn:=xlValues, LookAt:=xlWhole).Column
                On Error GoTo 0
                
            ' If the "Revenue Share" column exists, convert it to values only (row 2 to last row)
                If RevenueShareCol > 0 Then
                    LastRow = wsNew.Cells(wsNew.Rows.Count, RevenueShareCol).End(xlUp).Row
                    wsNew.Range(wsNew.Cells(2, RevenueShareCol), wsNew.Cells(LastRow, RevenueShareCol)).Value = _
                    wsNew.Range(wsNew.Cells(2, RevenueShareCol), wsNew.Cells(LastRow, RevenueShareCol)).Value
                End If
            End If
        Next ws

    ' Turn Alerts off
        Application.DisplayAlerts = False

    ' Delete the first worksheet that was automatically created
        wbNew.Sheets(1).Delete

    ' Turn Alerts back on
        Application.DisplayAlerts = True
        
    ' Turn screen updating back on
        Application.ScreenUpdating = True

' Give the user a message letting them know the split was successful
    MsgBox "Worksheets were successfully split into a new workbook. Thank you for your patience!", vbInformation, "Split Successful"

End Sub


