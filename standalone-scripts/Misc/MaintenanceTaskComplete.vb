' ==============================================================================
' MaintenanceTasks
' ==============================================================================
' This script is trigger by an email being received with "Task" in the body. It
' will parse the body of the email to determine which task was completed.
' ==============================================================================
Sub Main(Param As Object)
  Dim index
  Dim body as string
  index = hs.MailTrigger
  body = hs.MailText(index)
  body = String.Join("",body)
  Dim Message As String
  Dim TaskString As String() = Split(body," ")
  Dim EmailFrom = hs.MailFrom(index)
  
  If TaskString(0) = "Task" And ( TaskString(2) = "complete" Or TaskString(2) = "completed") Then
      Dim TaskId As Integer = Convert.ToInt32(TaskString(1))
      Dim device As Scheduler.Classes.DeviceClass = hs.GetDeviceByRef(TaskId)
      Dim CompletedBy As String
      
      hs.WriteLog("Maintenance Task Complete", "Task " & TaskId & " will be marked as complete. Message was received at " & hs.MailDate(index))

      If (((TaskString.Count > 3) And (TaskString(TaskString.Count-1) = "Derek")) Or ((TaskString.Count < 4) And (EmailFrom = hs.DeviceString(168)))) Then
        CompletedBy = "Derek"
      ElseIf (((TaskString.Count > 3) And (TaskString(TaskString.Count-1) = "Amy")) Or ((TaskString.Count < 4) And (EmailFrom = hs.DeviceString(167)))) Then
        CompletedBy = "Amy"
      Else
        CompletedBy = "Unknown"
      End If

      ' Turn off the device if necessary
      'If (hs.DeviceValue(TaskId) <> 0 And device.Location2(hs) = "Trackers" And device.Location(hs) = "Maintenance") Then
      If (hs.DeviceValue(TaskId) <> 0 ) Then
        hs.SetDeviceValueByRef(TaskId,0,True)

        Select Case CompletedBy
          Case "Derek"
            hs.SetDeviceValueByRef(562, hs.DeviceValue(562) + 1, true)
            hs.SetDeviceValueByRef(564, hs.DeviceValue(564) + 1, true)
          Case "Amy"
            hs.SetDeviceValueByRef(563, hs.DeviceValue(563) + 1, true)
            hs.SetDeviceValueByRef(565, hs.DeviceValue(565) + 1, true)
        End Select

        Message = "Task " & TaskId & " marked complete at " & hs.MailDate(index) & " by " & CompletedBy
        hs.WriteLog("Maintenance Task Complete", Message)
        
        Message = Message & "<br /><br />Weekly Tasks Completed<br />Derek: " & hs.DeviceValue(562) & "<br />Amy: " & hs.DeviceValue(563) & "<br /><br />Monthly Tasks Completed<br />Derek: " & hs.DeviceValue(564) & "<br />Amy: " & hs.DeviceValue(565)
        SendMessage("Maintenance", Message)

        ' Set the last change to the time the message was sent
        hs.SetDeviceLastChange(TaskId,hs.MailDate(index))
      Else
        SendMessage("Maintenance","Task " & TaskId & " already marked complete.")
      End If

      ' Delete the email
      hs.MailDelete(index)
  Else
    hs.WriteLog("Maintenance Task", "Not a valid string: " & TaskString)
  End If
End Sub

' ==============================================================================
' Send Message
' ==============================================================================
Sub SendMessage(SubjectString As String, MessageString As String)

  ' Set up enumerator
  Dim Device As Scheduler.Classes.DeviceClass
  Dim Enumerator As Scheduler.Classes.clsDeviceEnumeration

  ' Start
  Enumerator = hs.GetDeviceEnumerator

  ' Check
  If (Enumerator Is Nothing) Then
    hs.WriteLog("MMS Messaging", "Error getting enumerator")
    Exit Sub
  End If

  ' Loop and send messages
  Do While Not Enumerator.Finished
    Device = Enumerator.GetNext

    If (Device Is Nothing) Then
      Continue Do
    End If

    ' If the device type string is MMSPhoneNumber and the device is on,
    ' send the message. Checking if the device is off allows for easily
    ' disabling sending to a specific number without modifying scripts.
    If (Device.Device_Type_String(hs) = "MMSPhoneNumber") Then
      If (hs.DeviceValue(Device.ref(hs)) = 100) Then
        hs.SendEmail(hs.DeviceString(Device.ref(hs)), hs.DeviceString(174), "", "", SubjectString, MessageString, "")
      Else
        hs.WriteLog("MMS Messaging", "Device is disabled for messaging (ReferenceID: " & Device.ref(hs) & ")")
      End If
    End If

  Loop

End Sub
