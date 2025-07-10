' -----------------------------------------
' VBA Outlook_Automation:
' Creating and sending an email
' -----------------------------------------
'
' Basic email:
'   Dim olApp As Outlook.Application
'   Dim olMail As Outlook.MailItem
'
'   Set olApp = New Outlook.Application
'   Set olMail = olApp.CreateItem(olMailItem)
'
'   With olMail
'       .To = "recipient@example.com"
'       .Subject = "Test Email from VBA"
'       .Body = "Hello, this is a test email sent via VBA."
'       .Send  ' Use .Display to review before sending
'   End With
'
' Adding attachments:
'   .Attachments.Add "C:\Path\To\File.xlsx"
'
' Using HTMLBody:
'   .HTMLBody = "<h2>Hello</h2><p>This is an <b>HTML</b> email.</p>"
'
' Best practices:
' - Use .Display for testing.
' - Release objects with Set ... = Nothing.
' - Handle errors gracefully.
' - Use BCC and CC as needed.
'
' -----------------------------------------
