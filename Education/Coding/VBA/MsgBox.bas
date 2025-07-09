' Message Boxes

    ' MsgBox(prompt, buttons, title, [helpfile], [context]) -> Parentheses can be used or you can just add a space after 'MsgBox'
    
        ' Prompt – Message displayed to the user.

        ' Buttons – Combine multiple options using "+" or "Or"; specify what buttons MsgBox shows (only 1 button set allowed at a time)

            ' Button Sets (choose only one):
                ' vbOKOnly (default if not specified) – Shows "OK" button
                ' vbOKCancel – Shows "OK" and "Cancel" buttons
                ' vbAbortRetryIgnore – Shows "Abort", "Retry", and "Ignore" buttons
                ' vbYesNoCancel – Shows "Yes", "No", and "Cancel" buttons
                ' vbYesNo – Shows "Yes" and "No" buttons
                ' vbRetryCancel – Shows "Retry" and "Cancel" buttons
                
                ' Possible button responses (and their numeric values):
                    ' vbOK (1), vbCancel (2), vbAbort (3), vbRetry (4), vbIgnore (5), vbYes (6), vbNo (7)

            ' Icons (optional, choose only one):
                ' vbCritical – Critical error icon (red X)
                ' vbExclamation – Warning icon (yellow triangle with "!")
                ' vbInformation – Info icon (blue circle with "i")
                ' vbQuestion – Question icon (blue "?")

            ' Default button (optional, choose only one) – Specifies which button has initial focus:
                ' vbDefaultButton1 – 1st button
                ' vbDefaultButton2 – 2nd button
                ' vbDefaultButton3 – 3rd button
                ' vbDefaultButton4 – 4th button

            ' Modality (optional, choose only one):
                ' vbApplicationModal (default) – Blocks input to the application until user responds
                ' vbSystemModal – Blocks input to all windows until user responds

            ' Alignment & UI options (optional):
                ' vbMsgBoxRight – Text aligned right
                ' vbMsgBoxRtlReading – Right-to-left reading order (e.g., Arabic/Hebrew)
                ' vbMsgBoxSetForeground – MsgBox stays on top of all windows
                ' vbMsgBoxHelpButton – Adds Help button (requires helpfile argument)

        ' Title – Text that appears in the title bar of the message box window (optional but recommended)

' Examples:
Sub MsgBox_Examples()

    ' 1. Default: message only (vbOKOnly is default)
        MsgBox "Default MsgBox (vbOKOnly default)"

    ' 2. vbOKOnly + vbCritical
        MsgBox "vbOKOnly + vbCritical", vbOKOnly + vbCritical, "Ok Only + Critical"

    ' 3. vbOKCancel + vbExclamation + vbDefaultButton1 + vbMsgBoxRight
        MsgBox "vbOKCancel + vbExclamation + vbDefaultButton1 + vbMsgBoxRight", vbOKCancel + vbExclamation + vbDefaultButton1 + vbMsgBoxRight, "OkCancel + Exclamation"

    ' 4. vbAbortRetryIgnore + vbInformation + vbDefaultButton2 + vbMsgBoxRtlReading
        MsgBox "vbAbortRetryIgnore + vbInformation + vbDefaultButton2 + vbMsgBoxRtlReading", vbAbortRetryIgnore + vbInformation + vbDefaultButton2 + vbMsgBoxRtlReading, "AbortRetryIgnore + Info"

    ' 5. vbYesNoCancel + vbQuestion + vbDefaultButton3 + vbMsgBoxSetForeground
        MsgBox "vbYesNoCancel + vbQuestion + vbDefaultButton3 + vbMsgBoxSetForeground", vbYesNoCancel + vbQuestion + vbDefaultButton3 + vbMsgBoxSetForeground, "YesNoCancel + Question"

    ' 6. vbYesNo + vbCritical + vbMsgBoxHelpButton
        MsgBox "vbYesNo + vbCritical + vbMsgBoxHelpButton", vbYesNo + vbCritical + vbMsgBoxHelpButton, "YesNo + Critical + Help"

    ' 7. vbRetryCancel + vbSystemModal
        MsgBox "vbRetryCancel + vbSystemModal", vbRetryCancel + vbSystemModal, "RetryCancel + SystemModal"

End Sub
