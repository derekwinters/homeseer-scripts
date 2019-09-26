' ==============================================================================
' How This Script Works
' ==============================================================================
' The BLLAN plugin is used to track devices via their network connection status.
' That device also has a custom "User Note" property which stores the reference
' ID of the virtual tracking device for users that need to be tracked for
' vacation mode.
' Devices are named either "Family <Name>" or "Guest <Name>" to determine how
' that device should be handled. Family members take priority over guests for
' overall occupancy status, but guests are also able to be considered as
' occupants when all family members are gone. Only trusted guests are added to
' the "Guest" list.
' An additional Babysitter device exists for special use cases.
' ==============================================================================

Sub Main(Parms as Object)
  ' ============================================================================
  ' Variables
  ' ============================================================================
  ' Enumerator devices
  Dim Device As Scheduler.Classes.DeviceClass
  Dim Enumerator As Scheduler.Classes.clsDeviceEnumeration

  ' Counting variables
  Dim VacationCount As Integer = 0
  Dim OccupancyCount As Integer = 0
  Dim UserTotal As Integer = 0
  Dim GuestCount As Integer = 0

  ' Static Devices
  Dim TrackerGuest As Integer = 445
  Dim TrackerBabysitter As Integer = 25
  Dim OccupancyMode As Integer = 75
  Dim OccupancyMode_Full As Integer = 100
  Dim OccupancyMode_Occupied As Integer = 75
  Dim OccupancyMode_Away As Integer = 25
  Dim OccupancyMode_Vacation As Integer = 0
  
  ' Device objects
  Dim Custom_DeviceName As String = ""
  Dim Custom_UserNote As Integer = 0
  Dim Custom_Name As String = ""

  ' ============================================================================
  ' Start
  ' ============================================================================
  Enumerator = hs.GetDeviceEnumerator
  If Enumerator Is Nothing Then
    hs.WriteError("Occupancy"," Error getting enumerator")
    Exit Sub
  Else
    ' Loop through all devices to set virtual tracking devices to match
    Do While Not Enumerator.Finished
      Device = Enumerator.GetNext

      If Device Is Nothing Then
        Continue Do
      End If

      If Device.Device_Type_String(hs) = "BLLAN Plug-In Device" Then
        Custom_DeviceName = hs.DeviceName(Device.ref(hs))
        Custom_Name = Custom_DeviceName.Replace("Network Devices ","")
        If Device.UserNote(hs) <> "" Then
          Custom_UserNote = Convert.toInt32(Device.UserNote(hs))
        End If
        
        If Custom_Name.StartsWith("Family") Then
          hs.WriteLog("Occupancy","Found " & Custom_Name & " (" & Device.Ref(hs) & ") Value: " & hs.DeviceValue(Device.Ref(hs)))
          UserTotal += 1
          MatchUserOccupancy(Device.Ref(hs),Custom_UserNote,Custom_Name)
          If hs.DeviceValueEx(Custom_UserNote) = 0 Then
            VacationCount += 1
          ElseIf hs.DeviceValueEx(Custom_UserNote) = 100 Then
            OccupancyCount += 1
          End If
        ElseIf Custom_Name.StartsWith("Guest") Then
          hs.WriteLog("Occupancy","Found " & Custom_Name & " (" & Device.Ref(hs) & ") Value: " & hs.DeviceValue(Device.Ref(hs)))
          If hs.DeviceValueEx(Device.ref(hs)) = 100 Then
            hs.WriteLog("Occupancy",Custom_Name & " is here")
            GuestCount += 1
          End If
        End If
      End If
    Loop

    ' Enable the guest device if any guests are here
    If hs.DeviceValue(TrackerGuest) <> 100 And GuestCount > 0 Then
      hs.SetDeviceValueByRef(TrackerGuest,100,True)
    Else
      hs.SetDeviceValueByRef(TrackerGuest,0,True)
    End If

    ' Determine final occupancy settings
    If OccupancyCount = UserTotal Then
      hs.WriteLog("Occupancy","All family members are home")
      hs.SetDeviceValueByRef(OccupancyMode,OccupancyMode_Full,True)
    ElseIf OccupancyCount > 0 And OccupancyCount + VacationCount = UserTotal Then
      hs.WriteLog("Occupancy","All non-vacation family members are home")
      hs.SetDeviceValueByRef(OccupancyMode,OccupancyMode_Full,True)
    ElseIf OccupancyCount > 0 Then
      hs.WriteLog("Occupancy","Some family members are home")
      hs.SetDeviceValueByRef(OccupancyMode,OccupancyMode_Occupied,True)
    ElseIf hs.DeviceValue(TrackerBabysitter) = 100 Then
      If VacationCount = UserTotal Then
        hs.WriteLog("Occupancy","Babysitter is home")
        hs.SetDeviceValueByRef(OccupancyMode,OccupancyMode_Full,True)
      Else
        hs.WriteLog("Occupancy","Babysitter is home")
        hs.SetDeviceValueByRef(OccupancyMode,OccupancyMode_Occupied,True)
      End If
    ElseIf hs.DeviceValue(TrackerGuest) = 100 Then
      If VacationCount = UserTotal Then
        hs.WriteLog("Occupancy",GuestCount & " guests are home")
        hs.SetDeviceValueByRef(OccupancyMode,OccupancyMode_Full,True)
      Else
        hs.WriteLog("Occupancy",GuestCount & " guests are home")
        hs.SetDeviceValueByRef(OccupancyMode,OccupancyMode_Occupied,True)
      End If
    ElseIf VacationCount = UserTotal Then
      hs.WriteLog("Occupancy","Vacation mode enabled")
      hs.SetDeviceValueByRef(OccupancyMode,OccupancyMode_Vacation,True)
    Else
      hs.WriteLog("Occupancy","No one is home")
      hs.SetDeviceValueByRef(OccupancyMode,OccupancyMode_Away,True)
    End If

  End If
End Sub

Sub MatchUserOccupancy(Device As Integer, User As Integer, Name As String)
  If hs.DeviceValue(Device) = 100 And hs.DeviceValue(User) <> 100 Then
    hs.WriteLog("Occupancy",Name & " arrived")
    hs.SetDeviceValueByRef(User,100,True)
    If hs.DeviceValue(25) = 100 Then
      hs.WriteLog("Occupancy","Babysitter tracking disabled.")
      hs.SetDeviceValueByRef(25,50,True)
    End If
  ElseIf hs.DeviceValue(Device) = 0 And hs.DeviceValue(User) <> 50 Then
    hs.WriteLog("Occupancy",Name & " left")
    hs.SetDeviceValueByRef(User,50,True)
  End If
End Sub
