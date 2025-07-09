' Message Boxes
    ' Syntax:
    '   MsgBox(prompt, buttons, title, [helpfile], [context])

        ' Prompt – Message displayed to the user (string).

        ' Buttons – Specify the type of message box by combining multiple options using "+" or "Or".
        '           Only one button set is allowed, but you can combine it with icons, default buttons, modality, and other UI options.

            ' Button Sets (choose only one):
                ' vbOKOnly (default if not specified) – Shows "OK" button
                ' vbOKCancel – Shows "OK" and "Cancel" buttons
                ' vbAbortRetryIgnore – Shows "Abort", "Retry", and "Ignore" buttons
                ' vbYesNoCancel – Shows "Yes", "No", and "Cancel" buttons
                ' vbYesNo – Shows "Yes" and "No" buttons
                ' vbRetryCancel – Shows "Retry" and "Cancel" buttons

                ' Possible return values (numeric):
                    ' vbOK (1), vbCancel (2), vbAbort (3), vbRetry (4), vbIgnore (5), vbYes (6), vbNo (7)

            ' Icons (optional, choose only one):
                ' vbCritical – Critical error icon (red X)
                ' vbExclamation – Warning icon (yellow triangle with "!")
                ' vbInformation – Info icon (blue circle with "i")
                ' vbQuestion – Question icon (blue "?")

            ' Default Button (optional, choose only one):
                ' vbDefaultButton1 – 1st button (default)
                ' vbDefaultButton2 – 2nd button
                ' vbDefaultButton3 – 3rd button
                ' vbDefaultButton4 – 4th button

            ' Modality (optional, choose only one):
                ' vbApplicationModal (default) – Blocks input to the calling application until user responds
                ' vbSystemModal – Blocks input to all applications until user responds (MsgBox stays on top system-wide)

            ' Alignment & UI Options (optional):
                ' vbMsgBoxRight – Text aligned right
                ' vbMsgBoxRtlReading – Right-to-left reading order (for RTL languages)
                ' vbMsgBoxSetForeground – MsgBox brought to the foreground
                ' vbMsgBoxHelpButton – Adds a Help button (requires helpfile argument)

        ' Title – Text that appears in the title bar of the MsgBox window (optional but recommended for clarity).

        ' Helpfile – Optional path to a legacy Windows Help (.hlp) file (rarely used today; requires WinHlp32.exe support).
        
        ' Context – Numeric context ID for a topic within the help file (only relevant if Helpfile is provided).

    ' Notes:
        ' - Named arguments (Prompt:=, Buttons:=, Title:=, Helpfile:=, Context:=) can be used for clarity or if providing arguments out of order.
        ' - Example:
        '     MsgBox Prompt:="Are you sure?", Buttons:=vbYesNo + vbQuestion, Title:="Confirmation"


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

    ' 6. vbYesNo Or vbCritical Or vbMsgBoxHelpButton - "Or" instead of "+"
        MsgBox "vbYesNo Or vbCritical Or vbMsgBoxHelpButton", vbYesNo Or vbCritical Or vbMsgBoxHelpButton, "YesNo + Critical + Help"

    ' 7. vbRetryCancel + vbSystemModal - naming arguments
        MsgBox Title:="RetryCancel + SystemModal", Buttons:=vbRetryCancel + vbSystemModal, Prompt:="vbRetryCancel + vbSystemModal"

End Sub
