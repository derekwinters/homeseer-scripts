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
  End If
  
  ' Set the last change to the time the message was sent
  hs.SetDeviceLastChange(TaskId,hs.MailDate(index))
End Sub