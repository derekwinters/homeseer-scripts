' ==============================================================================
' MaintenanceTasks
' ==============================================================================
' This script is trigger by an email being received with "Task" in the body. It
' will parse the body of the email to determine which task was completed.
' ==============================================================================
Sub Main(Param As Object)
  Dim index
  Dim body
  index = hs.MailTrigger
  body = hs.MailText(index)

  hs.WriteLog("MaintenanceTaskComplete",body)
End Sub
