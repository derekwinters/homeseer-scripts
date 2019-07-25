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
  Dim Message As String = ""

  Do While Not Enumerator.Finished
    Device = Enumerator.GetNext

    If (Device Is Nothing) Then
      Continue Do
    End If

    If (Device.Device_Type_String(hs).StartsWith("MaintenanceTask")) Then
      ' Parse DeviceTypeString
      TaskString = Split(Device.Device_Type_String(hs).ToString," ")
      TaskType = TaskString(1)
      TaskPeriod = TaskString(2)
      TaskName = hs.DeviceName(Device.ref(hs)).Replace("Trackers Maintenance",String.Empty)

      ' Log discovery
      hs.WriteLog("Maintenance Task", "ReferenceID: " & Device.ref(hs) & " | Type: " & TaskType & " | Period: " & TaskPeriod & " | Name: " & TaskName)

      ' ==================================================================
      ' 
      ' ==================================================================
      Select Case TaskType
        Case = "DayOfWeek"
          If ( TodayDay = TaskPeriod ) Then
            Message = "Maintenance Task " & Device.ref(hs) & ": " & TaskName & " is due today."
          End If
        Case = "DayOfMonth"
          If ( TodayNumber = TaskPeriod ) Then
            Message = ""
          End If
        Case = "FirstDayOfMonth"
          If ( TodayDay = TaskPeriod And TodayNumber <= 7 ) Then
            Message = ""
          End If
      End Select

      If ( Message <> "" ) Then
        SendMessage("Maintenance Task",Message)
        hs.WriteLog("Maintenance Task", Message)
        Message = ""
      End If

    End If

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
