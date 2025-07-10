' -----------------------------------------
' VBA Outlook_Automation:
' Automating tasks and reminders
' -----------------------------------------
'
' Access Tasks folder:
'   Set olTasks = olNS.GetDefaultFolder(olFolderTasks)
'
' Create new task:
'   Set olTask = olTasks.Items.Add(olTaskItem)
'   With olTask
'       .Subject = "Prepare report"
'       .DueDate = Date + 7
'       .Status = olTaskNotStarted
'       .PercentComplete = 0
'       .ReminderSet = True
'       .ReminderTime = Now + TimeValue("00:30:00")
'       .Body = "Complete the sales report for next week."
'       .Save
'   End With
'
' List tasks with reminders:
'   For Each olItem In olTasks.Items
'       If olItem.Class = olTask Then
'           If olItem.ReminderSet Then
'               Debug.Print olItem.Subject & " - Reminder at " & olItem.ReminderTime
'           End If
'       End If
'   Next olItem
'
' Best practices:
' - Use reminders to notify users.
' - Set realistic due dates.
' - Release objects properly.
' - Handle updates carefully.
'
' -----------------------------------------
