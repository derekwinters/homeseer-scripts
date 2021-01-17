' ==============================================================================
' MaintenanceTasks
' ==============================================================================
' This script is trigger by an email being received with "Task" in the body. It
' will parse the body of the email to determine which task was completed.
' ==============================================================================

Imports System.Globalization

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
  Dim TaskName as String
  
  If Trim(TaskString(0)) = "Task" Then
    Dim TaskId As Integer = Convert.ToInt32(TaskString(1))
    Dim CompletedBy As String
    Dim Score As Integer = ScoreTask(TaskId)
    TaskName = hs.DeviceName(TaskId).Replace("Trackers Maintenance",String.Empty)

    hs.WriteLog("Maintenance Task", "Task scored as " & Score)
    
    ' If the message was simple, use the email address to determine the user
    If TaskString.Count < 4 Then
        CompletedBy = ConvertEmailToName(EmailFrom)
    Else
        CompletedBy = TaskString(TaskString.Count-1)
    End If

    CompletedBy = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(CompletedBy)

    If (Trim(TaskString(2)) = "complete" Or Trim(TaskString(2)) = "completed") Then
        hs.WriteLog("Maintenance Task Complete", "Task " & TaskId & " will be marked as complete. Message was received at " & hs.MailDate(index) & " from (" & EmailFrom & ")")

        Message = "'" & TaskName & "' marked complete at " & hs.MailDate(index) & " by " & CompletedBy & " for " & Score & " points."
        Select Case CompletedBy
            Case "Derek"
                hs.SetDeviceValueByRef(562, hs.DeviceValue(562) + Score, true)
                hs.SetDeviceValueByRef(564, hs.DeviceValue(564) + Score, true)
            Case "Amy"
                hs.SetDeviceValueByRef(563, hs.DeviceValue(563) + Score, true)
                hs.SetDeviceValueByRef(565, hs.DeviceValue(565) + Score, true)
        End Select
        CompleteTask(TaskId, Message, index)
    Else If Trim(TaskString(0)) = "Task" And ( Trim(TaskString(2)) = "skip" Or Trim(TaskString(2)) = "skipped") Then
        Message = "'" & TaskName & "' skipped at " & hs.MailDate(index) & " by " & CompletedBy
        CompleteTask(TaskId, Message, index)
    Else If Trim(TaskString(0)) = "Task" And ( Trim(TaskString(2)) = "due" ) Then
        MarkTaskDue(TaskId)
        ' Delete the email
        hs.WriteLog("Maintenance Task", "Deleting email " & index)
        hs.MailDelete(index)
    Else
        hs.WriteLog("Maintenance Task", "Not a valid task string: (" & body & ")")
    End If
  Else
    hs.WriteLog("Maintenance Task", "Not a valid task string: (" & body & ")")
  End If
End Sub

Sub MarkTaskDue(TaskId As Integer)
        Dim LastChange As Date
        hs.WriteLog("Maintenance Task", "Task " & TaskId & " will be marked due.")
        ' Save the LastChangeDateTime
        LastChange = hs.DeviceDateTime(TaskId)

        ' Turn the device on
        hs.SetDeviceValueByRef(TaskId,100,True)

        ' Reset the LastChange
        hs.SetDeviceLastChange(TaskId,LastChange)
End Sub

' ==============================================================================
' CompleteTask
' ==============================================================================
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

  If hs.DeviceValue(TaskId) <> 0 Then
    hs.SetDeviceValueByRef(TaskId,0,True)
  End If

  hs.WriteLog("Maintenance Task", Message)

  Message = Message & "<br /><br />Weekly Tasks Completed<br />Derek: " & hs.DeviceValue(562) & "<br />Amy: " & hs.DeviceValue(563) & "<br /><br />Monthly Tasks Completed<br />Derek: " & hs.DeviceValue(564) & "<br />Amy: " & hs.DeviceValue(565)
  SendMessage("Maintenance", Message)

  ' Set the last change to the time the message was sent
  hs.SetDeviceLastChange(TaskId,hs.MailDate(index))

  ' Delete the email
  hs.WriteLog("Maintenance Task", "Deleting email " & index)
  hs.MailDelete(index)
End Sub

' ==============================================================================
' ConvertEmailToName
' ==============================================================================
Private Function ConvertEmailToName(Email As String) As String
  Select Case Email
    Case hs.DeviceString(168)
      Return "Derek"
    Case hs.DeviceString(167)
      Return "Amy"
    Case Else
      Return "Unknown"
  End Select
End Function

' ==============================================================================
' Score Task
' ==============================================================================
Private Function ScoreTask(TaskId As Integer) As Integer
  Select Case TaskId
    Case 223
      Return 2
    Case 224
      Return 1
    Case 317
      Return 1
    Case 318
      Return 3
    Case 321
      Return 5
    Case 356
      Return 2
    Case 429
      Return 3
    Case 435
      Return 2
    Case 436
      Return 3
    Case 521
      Return 3
    Case 524
      Return 2
    Case 545
      Return 2
    Case 546
      Return 2
    Case 547
      Return 5
    Case 548
      Return 5
    Case 549
      Return 2
    Case 550
      Return 5
    Case 551
      Return 5
    Case 552
      Return 3
    Case 553
      Return 5
    Case 554
      Return 1
    Case 555
      Return 5
    Case 556
      Return 2
    Case 559
      Return 3
    Case 560
      Return 5
    Case 561
      Return 3
    Case 566
      Return 1
    Case 567
      Return 1
    Case 588
      Return 3
    Case 589
      Return 3
    Case 620
      Return 5
    Case 625
      Return 3
    Case 626
      Return 2
    Case 627
      Return 3
    Case Else
      Return 0
  End Select
End Function

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
