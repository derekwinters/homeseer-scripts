' ==============================================================================
' Irrigation Controller
' ==============================================================================
' Use the RainMachine plugin to control the irrigation system. Base the run time
' for each zone on weather from the WeatherXML plugin.
' ==============================================================================
Sub Main(Parm As Object)
  ' ============================================================================
  ' Load devices and data
  ' ============================================================================
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

  ' Minimum Run Percentage
  Dim MinimumRuntime As Double = 0.60

  ' Rain total last four days
  ' This is a placeholder until it's possible to determine how much rain
  ' has fallen in the last four days
  Dim RecentWaterTotal As Double = 0.0

  ' Desired water last four days
  Dim DesiredWaterInches As Double = 0.5

  ' Today's high temperature
  Dim TemperatureHigh As Integer = hs.DeviceValue(32)
  
  ' Create message string
  Dim Message As String

  ' ============================================================================
  ' Determine water needed
  ' ============================================================================
  ' Add an extra 0.25 inches of water on very warm days
  If TemperatureHigh > 95 Then
    DesiredWaterInches = DesiredWaterInches + 0.25
  End If

  ' Subtract the recent water from what is desired
  Dim WaterNeeded As Double = DesiredWaterInches - RecentWaterTotal

  ' Turn this into a percentage multiplier
  Dim RainMultiplier As Double = WaterNeeded / DesiredWaterInches

  ' ============================================================================
  ' Set the zone times if rain is still needed
  ' ============================================================================
  If WaterNeeded > 0 Then
    hs.WriteLog("Irrigation Controller","An additional " & WaterNeeded & " inches of water are needed to achieve the desired " & DesiredWaterInches & " total inches.")

    If RainMultiplier > MinimumRuntime Then
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

    Else
      hs.WriteLog("Irrigation Controller","The rain multiplier " & RainMultiplier & " is below the minimum threshold " & MinimumRuntime)
    End If
  End If
End Sub

' ==============================================================================
' Zone Controller
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
'
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