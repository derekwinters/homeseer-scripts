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
  Dim Zone1 As Integer = 358
  Dim Zone2 As Integer = 359
  Dim Zone3 As Integer = 360
  Dim Zone4 As Integer = 361
  Dim Zone5 As Integer = 362
  Dim Zone6 As Integer = 363

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
    hs.WriteLog("IrrigationController","An additional " & WaterNeeded & " inches of water are needed to achieve the desired " & DesiredWaterInches & " total inches.")

    If RainMultiplier > MinimumRuntime Then
      ' Use the multiplier
      Zone1Time = Zone1Time * RainMultiplier
      Zone2Time = Zone2Time * RainMultiplier
      Zone3Time = Zone3Time * RainMultiplier
      Zone4Time = Zone4Time * RainMultiplier
      Zone5Time = Zone5Time * RainMultiplier
      Zone6Time = Zone6Time * RainMultiplier

      ' Log the zone times
      hs.WriteLog("IrrigationController","Beginning irrigation (Zone1: " & Zone1 & " | Zone2: " & Zone2 & " | Zone3: " & Zone3 & " | Zone4: " & Zone4 & " | Zone5: " & Zone5 & " | Zone6: " & Zone6 & ")"

      ' Run the program
      ZoneController(Zone1,Zone1Time)
      ZoneController(Zone2,Zone2Time)
      ZoneController(Zone3,Zone3Time)
      ZoneController(Zone4,Zone4Time)
      ZoneController(Zone5,Zone5Time)
      ZoneController(Zone6,Zone6Time)

      hs.WriteLog("IrrigationController","Irrigation complete.")

    Else
      hs.WriteLog("IrrigationController","The rain multiplier " & RainMulitplier & " is below the minimum threshold " & MinimumRuntime)
    End If
  End If
End Sub

Sub ZoneController (ZoneId As Integer,ZoneRuntime As Integer)
  hs.WriteLog("IrrigationController","Running zone " & ZoneId & " for " & ZoneRuntime & " minutes.")

  ' Start the zone


  ' Sleep for the approximate runtime of the zone (milliseconds)
  Threading.Thread.Sleep(ZoneRuntime * 60000)

  ' Wait for the zone to complete
  While hs.DeviceValue(ZoneId) <> 0
    ' Check every 10 seconds to watch for the final zone completion
    Threading.Thread.Sleep(10000)
  End While

  hs.WriteLog("IrrigationController","Zone " & ZoneId & " has completed.")
End Sub
