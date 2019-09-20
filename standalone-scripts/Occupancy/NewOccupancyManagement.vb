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
  Dim UserTotal As Integer = 2
  Dim GuestCount As Integer = 0

  ' Static Devices
  Dim TrackerGuest As Integer = 445
  Dim TrackerBabysitter As Integer = 25

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
        If hs.DeviceName(Device.ref(hs)).Contains("Family") Then
          MatchUserOccupancy(Device.Ref(hs),Device.UserNote(hs),hs.DeviceName(Device.ref(hs)).Replace("Network Devices Family ",""))
          UserTotal += 1
          If hs.DeviceValueByRef(Device.UserNote(hs)) = 0 Then
            VacationCount += 1
          ElseIf hs.DeviceValueByRef(Device.UserNote(hs)) = 100 Then
            OccupancyCount += 1
          End If
        ElseIf hs.DeviceName(Device.ref(hs)).Contains("Guest") Then
          If hs.DeviceValueEx(Device.ref(hs)) = 100 Then
            hs.WriteLog("Occupancy",hs.DeviceName(Device.ref(hs)).Replace("Network Devices ","") & " is here")
            GuestCount += 1
          End If
        End If
      End If
    Loop

    ' Enable the guest device if any guests are here
    If hs.DeviceValueByRef(TrackerGuest) <> 100 And GuestCount > 0 Then
'      hs.SetDeviceValueByRef(TrackerGuest,100,True)
    End If

    ' Determine final occupancy settings
    If OccupancyCount = UserTotal Then
      hs.WriteLog("Occupancy","All family members are home")
    ElseIf OccupancyCount > 0 And OccupancyCount + VacationCount = UserTotal Then
      hs.WriteLog("Occupancy","All non-vacation family members are home")
    ElseIf OccupancyCount > 0 Then
      hs.WriteLog("Occupancy","Some family members are home")
    ElseIf hs.DeviceValueByRef(TrackerBabysitter) = 100 Then
      hs.WriteLog("Occupancy","Babysitter is home")
    ElseIf hs.DeviceValueByRef(TrackerGuest) = 100 Then
      hs.WriteLog("Occupancy",GuestCount & " guests are home")
    ElseIf VacationCount = UserTotal Then
      hs.WriteLog("Occupancy","Vacation mode enabled")
    Else
      hs.WriteLog("Occupancy","No one is home")
    End If

  End If
End Sub

Sub MatchUserOccupancy(Device As Integer, User As Integer, Name As String)
  If hs.DeviceValueByRef(Device) = 100 And hs.DeviceValueByRef(User) <> 100 Then
    hs.WriteLog("Occupancy",Name & " arrived")
'    hs.SetDeviceValueByRef(User,100,True)
    If hs.DeviceValueByRef(25) = 100 Then
      hs.WriteLog("Occupancy","Babysitter tracking disabled.")
'      hs.SetDeviceValueByRef(25,50,True)
    End If
  ElseIf hs.DeviceValueByRef(Device) = 0 And hs.DeviceValueByRef(User) <> 0 Then
    hs.WriteLog("Occupancy",Name & " left")
'    hs.SetDeviceValueByRef(User,0,True)
  End If
End Sub
