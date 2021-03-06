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
  Dim IntervalDays As Integer
  Dim IntervalMin As Double
  Dim LastChange As Date
  Dim TaskMinDays As String
  Dim TurnOnDevice As Boolean
  ' This variable is used to limit how often messages are sent
  Dim TaskPeriodMod As Integer

  Do While Not Enumerator.Finished
    Device = Enumerator.GetNext

    If (Device Is Nothing) Then
      Continue Do
    End If

    If (Device.Device_Type_String(hs).StartsWith("MaintenanceTask")) Then
      ' Reset
      TurnOnDevice = false
      TaskPeriodMod = 0

      TaskId = Device.ref(hs)
      ' Parse DeviceTypeString
      TaskString = Split(Device.Device_Type_String(hs).ToString," ")
      TaskType = TaskString(1)
      TaskPeriod = TaskString(2)

      ' DayOfWeek supports a minDays value to allow Every x DayOfWeek
      ' ie: DayOfWeek=Monday MinDays=8 (every other Monday).
      If (TaskString.Count = 4) Then
        TaskMinDays = TaskString(3)
      Else
        TaskMinDays = "0"
      End If

      TaskName = hs.DeviceName(TaskId).Replace("Trackers Maintenance",String.Empty)
      TaskAge = Math.Round((hs.DeviceTime(TaskId)/1440),0,MidpointRounding.AwayFromZero)

      ' Log discovery
      hs.WriteLog("Maintenance Task", "ReferenceID: " & TaskId & " | Type: " & TaskType & " | Period: " & TaskPeriod & " | Name: " & TaskName & " | Age: " & TaskAge)

      ' ==================================================================
      ' 
      ' ==================================================================
      ' If the device is off, check if it should be turned on
      If (hs.DeviceValue(TaskId) = 0) Then
        Select Case TaskType
          Case = "DayOfWeek"
            If ( TodayDay = TaskPeriod ) Then
              If (TaskAge > TaskMinDays) Then
                TurnOnDevice = true
              End If
            End If
          Case = "DayOfMonth"
            If ( TodayNumber = TaskPeriod ) Then
              TurnOnDevice = true
            End If
          Case = "FirstDayOfMonth"
            If ( TodayDay = TaskPeriod And TodayNumber <= 7 ) Then
              TurnOnDevice = true
            End If
          Case = "Interval"
            ' Check the interval against the setting
            If (TaskAge > TaskPeriod) Then
              TurnOnDevice = true

            End If
        End Select
      End If

      If (TurnOnDevice) Then
        ' Save the LastChangeDateTime
        LastChange = hs.DeviceDateTime(TaskId)

        ' Turn the device on
        hs.SetDeviceValueByRef(TaskId,100,True)

        ' Reset the LastChange
        hs.SetDeviceLastChange(TaskId,LastChange)
      End If

      If TaskType = "Interval" Then
        ' Check the TaskPeriod and set a sliding alert window to lower
        ' the number of reminders each day (something that is only done
        ' once a year doesn't need to send a reminder twice a day every
        ' day until it's done).
        If TaskPeriod > 300 Then
          TaskPeriodMod = TaskAge Mod 7

        Else If TaskPeriod > 180 Then
          TaskPeriodMod = TaskAge Mod 4

        Else If TaskPeriod > 60 Then
          TaskPeriodMod = TaskAge Mod 2

        End If
      End If

      ' If a Device is On, the task is due
      If (hs.DeviceValue(TaskId) = 100 And TaskPeriodMod = 0) Then
        Message = "Maintenance Task " & TaskId & ": " & TaskName & " is due today.<br /><br />It was last completed " & hs.DeviceDateTime(TaskId) & "<br /><br />Reply 'Task " & TaskId & " complete' to mark complete, or 'Task " & TaskId & " skipped' to reset the timer."
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
