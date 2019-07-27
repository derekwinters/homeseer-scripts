' ==============================================================================
' Irrigation Controller
' ==============================================================================
' Use the RainMachine plugin to control the irrigation system. Base the run time
' for each zone on weather from the WeatherXML plugin.
' ==============================================================================
Sub Main(Parm As Object)
  ' Water received in the last four days
  Dim RecentWaterTotal As Double = hs.DeviceValue(418)

  ' Desired water amount
  Dim DesiredWaterInches As Double = hs.DeviceValue(408)

  ' Check if the water received in the last four days is less than the desired
  ' amount of water. If it is, calculate the water requirements.
  If RecentWaterTotal < DesiredWaterInches Then
    CalculateWaterRequirement(RecentWaterTotal,DesiredWaterInches)
  Else
    hs.WriteLog("Irrigation Controller","Enough water has been received in the last four days, irrigation is not needed.")
  End If
End Sub

' ==============================================================================
' Calculate how much water is needed to meet the 4 day requirements
' ==============================================================================
' Determine how to modify the baseline zones and if it meets the minimum
' runtime requirements.
' ==============================================================================
Sub CalculateWaterRequirement(RecentWaterTotal As Double,DesiredWaterInches As Double)
  ' Minimum Run Percentage
  Dim MinimumRuntime As Double = 0.60

  ' Subtract the recent water from what is desired
  Dim WaterNeeded As Double = DesiredWaterInches - RecentWaterTotal

  ' Check if the water needed meeds the requirements
  If WaterNeeded > MinimumRuntime Then
    hs.WriteLog("Irrigation Controller",WaterNeeded & " inches of water are needed, running the irrigation system.")

    ' Determine how to modify the irrigation system based on the needed water
    ' and run the irrigation system.
    Dim RainMultiplier As Double = WaterNeeded / DesiredWaterInches

    IrrigationRun(RainMultiplier)
  Else
    hs.WriteLog("Irrigation Controller","The water needed (" & WaterNeeded & "/10 inches) is below the minimum threshold (" & MinimumRuntime & " inches).")
  End If

  ' Set the day0 device
  hs.SetDeviceValueByRef(414,(hs.DeviceValue(414) + WaterNeeded), True)
End Sub

' ==============================================================================
' Run the irrigation system
' ==============================================================================
' Load the devices for the irrigation zones and set a baseline for how long each
' zone. Accept a parameter to modify those runtimes based on previous water
' calculations from rain water.
' ==============================================================================
Sub IrrigationRun (RainMultiplier As Double)
  hs.WriteLog("Irrigation Controller","Setting irrigation zones with a multiplier of " & RainMultiplier)

  ' RainMachine devices
  Dim Zone1 As Integer = 387
  Dim Zone2 As Integer = 388
  Dim Zone3 As Integer = 389
  Dim Zone4 As Integer = 390
  Dim Zone5 As Integer = 391
  Dim Zone6 As Integer = 392

  ' Zone Default Runtimes
  Dim Zone1Time As Integer = 40
  Dim Zone2Time As Integer = 40
  Dim Zone3Time As Integer = 40
  Dim Zone4Time As Integer = 40
  Dim Zone5Time As Integer = 25
  Dim Zone6Time As Integer = 25

  ' Create message string
  Dim Message As String

  ' ============================================================================
  ' Set the zone times
  ' ============================================================================
  ' Use the multiplier
  Zone1Time = Zone1Time * RainMultiplier
  Zone2Time = Zone2Time * RainMultiplier
  Zone3Time = Zone3Time * RainMultiplier
  Zone4Time = Zone4Time * RainMultiplier
  Zone5Time = Zone5Time * RainMultiplier
  Zone6Time = Zone6Time * RainMultiplier

  ' Log the zone times
  Message = "Beginning irrigation (Zone1: " & Zone1Time & " | Zone2: " & Zone2Time & " | Zone3: " & Zone3Time & " | Zone4: " & Zone4Time & " | Zone5: " & Zone5Time & " | Zone6: " & Zone6Time & ")"
  hs.WriteLog("Irrigation Controller",Message)
  Message = Message.Replace("(","<br />")
  Message = Message.Replace(" | "," minutes<br />")
  Message = Message.Replace(")"," minutes")
  SendMessage("Rain Machine",Message)

  ' Set the program to run. Rain Machine will only allow one zone to run at
  ' a time, so there is no need to wait for zones to complete.
  ZoneController(Zone1,Zone1Time)
  ZoneController(Zone2,Zone2Time)
  ZoneController(Zone3,Zone3Time)
  ZoneController(Zone4,Zone4Time)
  ZoneController(Zone5,Zone5Time)
  ZoneController(Zone6,Zone6Time)

  hs.WriteLog("Irrigation Controller","Irrigation configuration complete.")
End Sub

' ==============================================================================
' Zone Controller
' ==============================================================================
' A simple function to wrap the CAPI calls for Rain Machine
' ==============================================================================
Sub ZoneController (ZoneId As Integer,ZoneRuntime As Integer)
  hs.WriteLog("Irrigation Controller","Setting zone " & hs.DeviceName(ZoneId) & " to run for for " & ZoneRuntime & " minutes.")

  ' Start the zone
  ZoneRuntime = ZoneRuntime * 60
'  MakeTheCapiHappy("Run for(value) Seconds",ZoneId,ZoneRuntime)

  ' Wait 10 seconds just to not run too fast
  Threading.Thread.Sleep(10000)
End Sub

' ==============================================================================
' Make the Capi Happy
' ==============================================================================
' Set a device using CAPI
' ==============================================================================
Sub MakeTheCapiHappy(CapiControlString As String, DeviceReferenceID As Integer, SetValue As Double)
  ' Create device
  Dim CapiControlDevice as HomeSeerAPI.CAPI.CAPIControl = hs.CAPIGetSingleControl(DeviceReferenceID, True, CapiControlString, False, False)

  ' Set value
  CapiControlDevice.ControlValue = SetValue

  ' Persist the value
  hs.CAPIControlHandler(CapiControlDevice)
End Sub

' ==============================================================================
' Send Message
' ==============================================================================
' Find MMSPhoneNumber devices, which are used to store cell phone numbers. If
' the device is enabled, send an email to the address. Using this method makes
' it easy to disable devices during testing.
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

    If (Device.Device_Type_String(hs) = "MMSPhoneNumber") Then
      If (hs.DeviceValue(Device.ref(hs)) = 100) Then
        hs.SendEmail(hs.DeviceString(Device.ref(hs)), hs.DeviceString(174), "", "", SubjectString, MessageString, "")
      Else
        hs.WriteLog("MMS Messaging", "Device is disabled for messaging (ReferenceID: " & Device.ref(hs) & ")")
      End If
    End If
  Loop
End Sub
