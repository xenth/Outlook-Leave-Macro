Sub CreateALAppt()
 Dim myItem As Object
 Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient

 Set myItem = Application.CreateItem(olAppointmentItem)
 myItem.MeetingStatus = olMeeting
 myItem.Subject = "<user> Annual Leave"
 myItem.Location = "Out of Office"
 myItem.BusyStatus = olOutOfOffice
 myItem.ResponseRequested = False
 'myItem.AllowNewTimeProposal = False
 'myItem.AllDayEvent = True
 myItem.Body = "I am out of the office on my vacation!"
 'myItem.Start = #9/24/2009 1:30:00 PM#
 'myItem.Duration = 90
 Set myRequiredAttendee = myItem.Recipients.Add("<user> <<email>>")
 myRequiredAttendee.Type = olRequired
 'Set myOptionalAttendee = myItem.Recipients.Add("<other_calendar_email>;")
 'myOptionalAttendee.Type = olOptional
 'Set myResourceAttendee = myItem.Recipients.Add("Conf Rm All Stars")
 'myResourceAttendee.Type = olResource
 myItem.Display (True)
 'myItem.Send
 CreateALApptULIT myItem
End Sub

Private Sub CreateALApptULIT(oldItem As Object)
 Dim myItem As Object
 Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient

 Set myItem = Application.CreateItem(olAppointmentItem)
 myItem.MeetingStatus = olMeeting
 myItem.Subject = oldItem.Subject + " - " + CStr(oldItem.Start) + " for " + CStr(oldItem.Duration / 60) + " hours"
 myItem.Location = oldItem.Location
 myItem.BusyStatus = olFree
 myItem.ResponseRequested = False
 'myItem.AllowNewTimeProposal = False
 myItem.Start = oldItem.Start
 myItem.AllDayEvent = True
 myItem.ReminderSet = False
 myItem.Body = "Automatically Generated Information:" + vbNewLine + "Starts at: " + CStr(oldItem.Start) + vbNewLine + "Goes for " + CStr(oldItem.Duration / 60) + " hours"
 'myItem.Duration = 90
 Set myRequiredAttendee = myItem.Recipients.Add("<boss_or_department_email>")
 myRequiredAttendee.Type = olRequired
 'Set myOptionalAttendee = myItem.Recipients.Add("<altemail>; <otheremail>")
 'myOptionalAttendee.Type = olOptional
 'Set myResourceAttendee = myItem.Recipients.Add("Conf Rm All Stars")
 'myResourceAttendee.Type = olResource
 myItem.Display (True)
 'myItem.Send
End Sub

