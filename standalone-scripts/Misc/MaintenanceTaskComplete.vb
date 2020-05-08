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
  Dim TaskString As String() = Split(body," ")
  Dim TaskId As Integer = Convert.ToInt32(TaskString(1))
  
  hs.WriteLog("Maintenance Task Complete", "Task " & TaskId & " will be marked as complete. Message was received at " & hs.MailDate(index))
  
  ' Turn off the device if necessary
  If (hs.DeviceValue(TaskId) <> 0) Then
    hs.SetDeviceValueByRef(TaskId,0,True)
    SendMessage("Maintenance","Task " & TaskId & " marked complete at " & hs.MailDate(index))
  Else
    SendMessage("Maintenance","Task " & TaskId & " already marked complete.")
  End If
  
  ' Set the last change to the time the message was sent
  hs.SetDeviceLastChange(TaskId,hs.MailDate(index))

  ' Delete the email
  hs.MailDelete(index)
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
