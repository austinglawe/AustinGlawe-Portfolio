' -----------------------------------------
' VBA Outlook_Automation:
' Accessing and reading inbox emails
' -----------------------------------------
'
' Access Outlook:
'   Set olApp = New Outlook.Application
'   Set olNS = olApp.GetNamespace("MAPI")
'   Set olInbox = olNS.GetDefaultFolder(olFolderInbox)
'
' Loop through emails:
'   For i = 1 To 10
'       Set olMail = olInbox.Items(i)
'       If TypeName(olMail) = "MailItem" Then
'           Debug.Print olMail.Subject
'       End If
'   Next i
'
' Filtering unread emails:
'   Set filteredItems = olInbox.Items.Restrict("[UnRead] = True")
'
' Read email properties:
'   Debug.Print olMail.Body
'   Debug.Print olMail.SenderName
'
' Best practices:
' - Check item type before access.
' - Use Restrict or Find for filtering.
' - Manage large inboxes carefully.
' - Release Outlook objects.
'
' -----------------------------------------
