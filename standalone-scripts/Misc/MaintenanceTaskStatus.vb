' ==============================================================================
' MaintenanceTasks
' ==============================================================================
' This function searches for devices with a specific DeviceTypeString. If it
' matches "MaintenanceTask <Type> <Period>", it will check if the conditions are
' met, an alert will be sent.
' If the period is Interval, the device will be turned on and the interval will
' be based of the change time of the device.
' ==============================================================================
Sub Main(Param As Object)
  ' ==========================================================================
  ' Get some basic info
  ' ==========================================================================
  Dim TodayNumber As String = DateTime.Now.ToString("dd")
  Dim TodayDay As String = DateTime.Now.ToString("dddd")

  ' ==========================================================================
  ' Set up enumerator
  ' ==========================================================================
  Dim Device As Scheduler.Classes.DeviceClass
  Dim Enumerator As Scheduler.Classes.clsDeviceEnumeration

  ' Start
  Enumerator = hs.GetDeviceEnumerator

  ' Check
  If (Enumerator Is Nothing) Then
    hs.WriteLog("Maintenance Task", "Error getting enumerator")
    Exit Sub
  End If

  ' ==========================================================================
  ' Loop
  ' ==========================================================================
  Dim TaskString() As String
  Dim TaskType As String
  Dim TaskPeriod As String
  Dim TaskName As String
  Dim TaskId As Integer
  Dim TaskAge As Integer
  Dim Message As String = ""

  Do While Not Enumerator.Finished
    Device = Enumerator.GetNext

    If (Device Is Nothing) Then
      Continue Do
    End If

    If (Device.Device_Type_String(hs).StartsWith("MaintenanceTask")) Then
      TaskId = Device.ref(hs)
      ' Parse DeviceTypeString
      TaskString = Split(Device.Device_Type_String(hs).ToString," ")
      TaskType = TaskString(1)
      TaskPeriod = TaskString(2)

      TaskName = hs.DeviceName(TaskId).Replace("Trackers Maintenance",String.Empty)
      TaskAge = Math.Round((hs.DeviceTime(TaskId)/1440),0,MidpointRounding.AwayFromZero)

      ' Log discovery
      hs.WriteLog("Maintenance Task", "ReferenceID: " & TaskId & " | Type: " & TaskType & " | Period: " & TaskPeriod & " | Name: " & TaskName & " | Age: " & TaskAge)

      ' Add to message
      If Message <> "" Then
        Message += "<br />"
      End If
      Message += "ReferenceID: " & TaskId & " | Type: " & TaskType & " | Period: " & TaskPeriod & " | Name: " & TaskName & " | Age: " & TaskAge
    End If

    SendMessage("Maintenance Task Status",Message)
  Loop
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
