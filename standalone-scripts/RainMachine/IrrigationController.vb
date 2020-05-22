' ==============================================================================
' Irrigation Controller
' ==============================================================================
' Use the RainMachine plugin to control the irrigation system. Base the run time
' for each zone on weather from the WeatherXML plugin.
' ==============================================================================
Sub Main(Parm As Object)
  If Parm Is Nothing Then
    ' Water received in the last four days
    Dim RecentWaterTotal As Double = hs.DeviceValueEx(417)
    
    ' Standard desired water amount + carryover
    Dim DesiredWaterInches As Double = hs.DeviceValueEx(408) + hs.DeviceValueEx(421)

    ' Check if the water received in the last four days is less than the desired
    ' amount of water. If it is, calculate the water requirements.
    If RecentWaterTotal < DesiredWaterInches Then
      CalculateWaterRequirement(RecentWaterTotal,DesiredWaterInches)
    Else
      Dim Message As String = "Enough water has been received in the last four days (received " & RecentWaterTotal & " inches, desired "& DesiredWaterInches &" inches), irrigation is not needed."
      hs.WriteLog("Irrigation Controller", Message)
      SendMessage("Rain Machine",Message)
    End If
  Else
    If IsNumeric(Parm) = True Then
      hs.WriteLog("Irrigation Controller", "Manual Irrigation of " & Parm & " inches requested.")

      ' Run Irrigation
      IrrigationRun(cdbl(Parm))

      ' Set the day0 device
      'hs.SetDeviceValueByRef(414,(hs.DeviceValue(414) + cint(Parm)), True)
    Else
      hs.WriteLog("Irrigation Controller", "Non-numeric value provided, cannot process request. (Parm = " & Parm & ")")
    End If
  End If
End Sub

' ==============================================================================
' Calculate how much water is needed to meet the 4 day requirements
' ==============================================================================
' Determine how to modify the baseline zones and if it meets the minimum water
' requirements.
' ==============================================================================
Sub CalculateWaterRequirement(RecentWaterTotal As Double,DesiredWaterInches As Double)
  ' Minimum Water (inches)
  Dim MinimumWater As Double = 0.3

  ' Maximum Water (inches)
  Dim MaximumWater As Double = 0.8

  ' Subtract the recent water from what is desired
  Dim WaterNeeded As Double = DesiredWaterInches - RecentWaterTotal

  If WaterNeeded > MaximumWater Then
    hs.WriteLog("Irrigation Controller","WaterNeeded (" & WaterNeeded & " inches) is greater than the maximum allowed (" & MaximumWater & " inches), resetting WaterNeeded.")
    WaterNeeded = MaximumWater
  End If

  ' Check if the water needed meeds the requirements
  If WaterNeeded >= MinimumWater Then
    hs.WriteLog("Irrigation Controller","A total of " & WaterNeeded & " inches of water are needed, running the irrigation system.")

    IrrigationRun(WaterNeeded)

    ' Set the day0 device
    hs.SetDeviceValueByRef(414,(hs.DeviceValue(414) + WaterNeeded), True)

    ' Clear the carryover
    If hs.DeviceValue(421) <> 0 Then
      hs.SetDeviceValueByRef(421,0,True)
    End If
  ElseIf WaterNeeded > 0 Then
    Dim Message As String = "The water needed (" & WaterNeeded & " inches) is below the minimum threshold (" & MinimumWater & "/10 inches). This will be added to the next watering."
    hs.WriteLog("Irrigation Controller", Message)
    SendMessage("Rain Machine",Message)
    hs.SetDeviceValueByRef(421,(hs.DeviceValue(421) + WaterNeeded),True)
  Else
    hs.WriteLog("Irrigation Controller","No water is needed.")
  End If
End Sub

' ==============================================================================
' Run the irrigation system
' ==============================================================================
' Load the devices for the irrigation zones and set a baseline for how long each
' zone. Accept a parameter to modify those runtimes based on previous water
' calculations from rain water.
' ==============================================================================
Sub IrrigationRun (WaterNeeded As Double)
  hs.WriteLog("Irrigation Controller","Setting irrigation zones to water " & WaterNeeded & " inches.")

  ' RainMachine devices
  Dim Zone1 As Integer = 387
  Dim Zone2 As Integer = 388
  Dim Zone3 As Integer = 389
  Dim Zone4 As Integer = 390
  Dim Zone5 As Integer = 391
  Dim Zone6 As Integer = 392

  ' Zone runtime per 1/10 inch
  Dim Zone1Time As Integer = 10
  Dim Zone2Time As Integer = 10
  Dim Zone3Time As Integer = 8
  Dim Zone4Time As Integer = 8
  Dim Zone5Time As Integer = 5
  Dim Zone6Time As Integer = 5

  ' Create message string
  Dim Message As String

  ' ============================================================================
  ' Set the zone times
  ' ============================================================================
  ' Use the multiplier
  Dim Zone1RunTime As Integer = Zone1Time * WaterNeeded * 10
  Dim Zone2RunTime As Integer = Zone2Time * WaterNeeded * 10
  Dim Zone3RunTime As Integer = Zone3Time * WaterNeeded * 10
  Dim Zone4RunTime As Integer = Zone4Time * WaterNeeded * 10
  Dim Zone5RunTime As Integer = Zone5Time * WaterNeeded * 10
  Dim Zone6RunTime As Integer = Zone6Time * WaterNeeded * 10

  ' Log the zone times
  Message = "Beginning irrigation (Zone1: " & Zone1RunTime & " | Zone2: " & Zone2RunTime & " | Zone3: " & Zone3RunTime & " | Zone4: " & Zone4RunTime & " | Zone5: " & Zone5RunTime & " | Zone6: " & Zone6RunTime & ")"
  hs.WriteLog("Irrigation Controller",Message)
  Message = Message.Replace("(","<br />")
  Message = Message.Replace(" | "," minutes<br />")
  Message = Message.Replace(")"," minutes")
  'SendMessage("Rain Machine",Message)

  ' Set the program to run. Rain Machine will only allow one zone to run at
  ' a time, so there is no need to wait for zones to complete.
  ZoneController(Zone1,ZoneRun1Time)
  ZoneController(Zone2,ZoneRun2Time)
  ZoneController(Zone3,ZoneRun3Time)
  ZoneController(Zone4,ZoneRun4Time)
  ZoneController(Zone5,ZoneRun5Time)
  ZoneController(Zone6,ZoneRun6Time)

'  The API calls (either from the plugin or the API itself) don't let this
'  make the zones do revolving runs. Subsequent calls to set up the addtional
'  cycles do net get created. Holding off on this change until this functionality works
'
'  Dim RunLoops As Double = WaterNeeded * 10
'  If RunLoops = 1 Then
'    hs.WriteLog("Irrigation Controller","Running 1 loop.")
'    Threading.Thread.Sleep(1000)
'    ZoneController(Zone1,Zone1Time)
'    ZoneController(Zone2,Zone2Time)
'    ZoneController(Zone3,Zone3Time)
'    ZoneController(Zone4,Zone4Time)
'    ZoneController(Zone5,Zone5Time)
'    ZoneController(Zone6,Zone6Time)
'  Else
'    Threading.Thread.Sleep(1000)
'    For run As Integer = 0 To (RunLoops-1)
'      hs.WriteLog("Irrigation Controller","Running loop " & run & ".")
'      Threading.Thread.Sleep(1000)
'      ZoneController(Zone1,Zone1Time)
'      ZoneController(Zone2,Zone2Time)
'      ZoneController(Zone3,Zone3Time)
'      ZoneController(Zone4,Zone4Time)
'      ZoneController(Zone5,Zone5Time)
'      ZoneController(Zone6,Zone6Time)
'    Next
'  End If
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
  MakeTheCapiHappy("Run for(value) Seconds",ZoneId,ZoneRuntime)

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
