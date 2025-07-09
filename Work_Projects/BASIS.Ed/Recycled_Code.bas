' Ask user a 'yes' or 'no' question. Store it in a variable called 'UserResponse'
  ' Possible responses: vbOK (1), vbCancel (2), vbAbort (3), vbRetry (4), vbIgnore (5), vbYes (6), vbNo (7)
Dim UserResponse As VbMsgBoxResult
' Ask user if they are sure they want to start the [] Converter 
  UserResponse = MsgBox("Are you sure you want to start the [] Converter?", vbYesNo + vbQuestion, "Confirmation to start the [] Converter")
