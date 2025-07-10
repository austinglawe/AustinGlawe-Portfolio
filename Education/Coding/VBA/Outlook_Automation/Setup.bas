' -----------------------------------------
' VBA Outlook_Automation:
' Setting up Outlook reference in VBA
' -----------------------------------------
'
' Add Reference:
' - Tools > References > Microsoft Outlook XX.0 Object Library
'
' Declare objects:
'   Dim olApp As Outlook.Application
'   Dim olMail As Outlook.MailItem
'
' Create Outlook instance:
'   Set olApp = New Outlook.Application
'
' Get existing instance if available:
'   On Error Resume Next
'   Set olApp = GetObject(, "Outlook.Application")
'   If olApp Is Nothing Then Set olApp = New Outlook.Application
'   On Error GoTo 0
'
' Best practices:
' - Use early binding during development.
' - Consider late binding for distribution.
' - Release objects with Set ... = Nothing.
' - Handle errors if Outlook not running.
'
' -----------------------------------------
