' ==============================================================================
' MaintenanceTasks
' ==============================================================================
' This script is trigger by an email being received with "Task" in the body. It
' will parse the body of the email to determine which task was completed.
' ==============================================================================
Sub Main(Param As Object)
  Dim index = hs.MailTrigger
  ' TMobile started to send unknown leading and trailing characterscharacters.
  ' Remove anything that isn't letters, numbers, or spaces, then trim any
  ' potential leading or trailing spaces.
  Imports System.Text.RegularExpressions
  Dim body as String = Trim(Regex.Replace(hs.MailText(index), "[^a-zA-Z0-9 ]", ""))
  ' Email address now can includes +1 at the start of the address.
  Dim EmailFrom = hs.MailFrom(index)
  EmailFrom = Regex.Replace(EmailFrom, "[^a-zA-Z0-9 @\.]", "")
  EmailFrom = Trim(EmailFrom)
  EmailFrom = Regex.Replace(EmailFrom, "^1", "")
  ' Split the message to parse later
  Dim TaskString As String() = Split(body, " ")
  ' Declare Message for later
  Dim Message As String
  
  If Trim(TaskString(0)) = "Task" Then
    Dim TaskId As Integer = Convert.ToInt32(TaskString(1))
    Dim CompletedBy As String

    If (Trim(TaskString(2)) = "complete" Or Trim(TaskString(2)) = "completed") Then
      hs.WriteLog("Maintenance Task Complete", "Task " & TaskId & " will be marked as complete. Message was received at " & hs.MailDate(index) & " from " & EmailFrom & ")")

      ' Check who completed the task
      If (((TaskString.Count > 3) And (TaskString(TaskString.Count-1)ToLower() = "derek")) Or ((TaskString.Count < 4) And (EmailFrom = hs.DeviceString(168)))) Then
        CompletedBy = "Derek"
      ElseIf (((TaskString.Count > 3) And (TaskString(TaskString.Count-1).ToLower() = "amy")) Or ((TaskString.Count < 4) And (EmailFrom = hs.DeviceString(167)))) Then
        CompletedBy = "Amy"
      Else
        If (TaskString.Count > 3) Then
          CompletedBy = "Unknown (" & TaskString(TaskString.Count-1) & ")"
        Else
          CompletedBy = "Unknown (" & EmailFrom & ")"
        End If
      End If

      Select Case CompletedBy
        Case "Derek"
          hs.SetDeviceValueByRef(562, hs.DeviceValue(562) + 1, true)
          hs.SetDeviceValueByRef(564, hs.DeviceValue(564) + 1, true)
        Case "Amy"
          hs.SetDeviceValueByRef(563, hs.DeviceValue(563) + 1, true)
          hs.SetDeviceValueByRef(565, hs.DeviceValue(565) + 1, true)
      End Select

      Message = "Task " & TaskId & " marked complete at " & hs.MailDate(index) & " by " & CompletedBy

      CompleteTask(TaskId, Message, index)
    Else If Trim(TaskString(0)) = "Task" And ( Trim(TaskString(2)) = "skip" Or Trim(TaskString(2)) = "skipped") Then

      ' Check who skipped the task
      If EmailFrom = hs.DeviceString(168) Then
        CompletedBy = "Derek"
      ElseIf EmailFrom = hs.DeviceString(167) Then
        CompletedBy = "Amy"
      Else
        CompletedBy = "Unknown (" & EmailFrom & ")"
      End If

      Message = "Task " & TaskId & " skipped at " & hs.MailDate(index) & " by " & CompletedBy

      CompleteTask(TaskId, Message, index)
    Else
      hs.WriteLog("Maintenance Task", "Not a valid task string: (" & body & ")")
    End If
  Else
    hs.WriteLog("Maintenance Task", "Not a valid task string: (" & body & ")")
  End If
End Sub

Sub CompleteTask(TaskId As Integer, Message As String, index As Integer)
  ' Turn off the device if necessary
  'Dim device As Scheduler.Classes.DeviceClass = hs.GetDeviceByRef(TaskId)
  'If (device.Location2(hs) = "Trackers" And device.Location(hs) = "Maintenance") Then
  '  hs.SetDeviceValueByRef(TaskId,0,True)
  '  ...
  '  ...
  'Else
  '  hs.WriteLog("Maintenance Task Complete", "This is not a maintenance task ID")
  'End If

  If (hs.DeviceValue(TaskId) <> 0 Then
    hs.SetDeviceValueByRef(TaskId,0,True)
  End If

  hs.WriteLog("Maintenance Task Complete", Message)

  Message = Message & "<br /><br />Weekly Tasks Completed<br />Derek: " & hs.DeviceValue(562) & "<br />Amy: " & hs.DeviceValue(563) & "<br /><br />Monthly Tasks Completed<br />Derek: " & hs.DeviceValue(564) & "<br />Amy: " & hs.DeviceValue(565)
  SendMessage("Maintenance", Message)

  ' Set the last change to the time the message was sent
  hs.SetDeviceLastChange(TaskId,hs.MailDate(index))

  ' Delete the email
  hs.WriteLog("Maintenance Task Complete", "Deleting email " & index)
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
