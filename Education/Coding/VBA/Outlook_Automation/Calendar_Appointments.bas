' -----------------------------------------
' VBA Outlook_Automation:
' Working with calendar appointments
' -----------------------------------------
'
' Access Calendar folder:
'   Set olCalendar = olNS.GetDefaultFolder(olFolderCalendar)
'
' Create appointment:
'   Set olAppt = olCalendar.Items.Add(olAppointmentItem)
'   With olAppt
'       .Subject = "Team Meeting"
'       .Start = Date + TimeValue("14:00:00")
'       .Duration = 60
'       .Location = "Conference Room"
'       .Body = "Discuss project status."
'       .ReminderSet = True
'       .ReminderMinutesBeforeStart = 15
'       .BusyStatus = olBusy
'       .Save
'   End With
'
' Loop through appointments in date range:
'   olItems.IncludeRecurrences = True
'   sFilter = "[Start] >= 'mm/dd/yyyy' AND [End] <= 'mm/dd/yyyy'"
'   Set olRestrictItems = olItems.Restrict(sFilter)
'   For Each olAppt In olRestrictItems
'       Debug.Print olAppt.Subject & " on " & olAppt.Start
'   Next olAppt
'
' Best practices:
' - Include recurring appointments.
' - Format dates properly in filters.
' - Handle time zones if needed.
' - Release Outlook objects.
'
' -----------------------------------------
