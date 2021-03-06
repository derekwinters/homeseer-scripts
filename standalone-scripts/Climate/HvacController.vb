﻿' ==============================================================================
' HVAC Controller
' ==============================================================================
' Controller the set points on the thermostat based on home occupancy and on the
' current and forcasted weather for the day.
'
' Author: Derek Winters
'
' ==============================================================================
' Scheduling
' ==============================================================================
' This script executes any time a temperature devices changes.
'
' ==============================================================================
' Help notes
' ==============================================================================
' CAPI control found at
' https://board.homeseer.com/showthread.php?p=1278809
'
' ==============================================================================
Sub Main(parms As Object)
  If ( hs.DeviceValue("202") <> 0 ) Then
    hs.WriteLog("HvacController", "Thermostat HOLD is enabled. No changes made.")
  Else
    ' ==========================================================================
    ' Set Variables
    ' ==========================================================================
    ' Outside Temperature from WeatherXML
    Dim OutsideTemperature As Integer = hs.DeviceValue("569")
    Dim TemperatureHigh As Integer = hs.DeviceValue("572")
    Dim TemperatureLow As Integer = hs.DeviceValue("573")
    Dim HvacMode As Integer = hs.DeviceValue("46")
    Dim CurrentHumidity As Integer = hs.DeviceValue("571")

    ' Current Home Occupancy Mode
    Dim OccupancyMode As Integer = hs.DeviceValue("75")

    ' Current hour
    Dim CurrentHour As Integer = Hour(Now())

    ' Thermostat devices
    Dim SetpointHeating As Integer = 48
    Dim SetpointCooling As Integer = 49
    Dim FanMode As Integer = 44

    ' CAPI control variable definitions
    Dim SetHeat As Double
    Dim SetCool As Double
    Dim SetMode As Double

    ' Current Average Temperature
    Dim AverageTemperature As Double = hs.DeviceValueEx(118)
    Dim AverageUpstairs As Double = hs.DeviceValueEx(324)
    Dim AverageDownstairs As Double = hs.DeviceValueEx(325)
    Dim AverageBedroom As Double = hs.DeviceValueEx(351)

    ' Current Operating State
    Dim CurrentOperatingState As Integer = hs.DeviceValue("47")

    ' Current Operating Mode
    Dim CurrentOperatingMode As Integer = hs.DeviceValue("46")

    ' Current thermostat Temperature Reading
    Dim CurrentThermostatReading As Integer = hs.DeviceValue("43")

    ' ==========================================================================
    ' If Heading Home is enabled, mock Occupancy
    ' ==========================================================================
    If ( hs.DeviceValue("211") = 100 ) Then
      hs.WriteLog("HvacController", "Heading home is enabled. Mocking OccupancyMode to 100")
      OccupancyMode = 100
    End If

    ' ==========================================================================
    ' Initialize the desired temperature
    ' ==========================================================================
    Dim DesiredWinter = 71
    Dim DesiredSummer = 74

    ' ==========================================================================
    ' Adjust the desired temperature for extremes
    ' ==========================================================================
    ' Modify the set points based on the outside temperature and time of day
    ExtremeTemperatureAdjustments(OutsideTemperature,TemperatureHigh,CurrentHour,CurrentHumidity,DesiredWinter,DesiredSummer)

    hs.WriteLog("HvacController", "Desired Winter: " & DesiredWinter & " | Desired Summer: " & DesiredSummer)

    ' ==========================================================================
    ' Adjust the desired temperature based on time of day
    ' ==========================================================================
    If (OccupancyMode = 100 OrElse OccupancyMode = 75 ) Then
      hs.WriteLog("HvacController","Weather Values: Outside: " & OutsideTemperature & " High: " & TemperatureHigh & " Low: " & TemperatureLow)

      TimeOfDayAdjustments(CurrentHour,SetHeat,SetCool,SetMode,DesiredWinter,DesiredSummer)

      hs.WriteLog("HvacController", "Time adjusted temperatures (Heat: " & SetHeat & " | Cool: " & SetCool & " | Mode: " & SetMode & ")")

      ' ========================================================================
      ' Average Temperature Alterations
      ' ========================================================================
      ' If the system is not active, check if the set points should be adjusted
      ' based on the AverageTemperature.
      If (CurrentOperatingState = 0) Then
        Dim HeatDifference As Double = Math.Abs(AverageTemperature - SetHeat)
        Dim CoolDifference As Double = Math.Abs(AverageTemperature - SetCool)
        Dim FloorDifference As Double = Math.Abs(AverageDownstairs - AverageUpstairs)

        AverageAdjustments(HeatDifference,CoolDifference,FloorDifference,AverageTemperature,SetHeat,SetCool,SetMode)

        hs.WriteLog("HvacController", "Average adjustments (Heat " & SetHeat & " | Cool: " & SetCool & " | Mode: " & SetMode & ")")
      End If

    Else If ( OccupancyMode = 0 ) Then
      ' Vacation
      SetHeat = DesiredWinter - 20
      SetCool = DesiredSummer + 4
      SetMode = 0
    Else
      ' Away
      SetHeat = DesiredWinter - 5
      SetCool = DesiredSummer + 1
      SetMode = 0
    End If

    ' ==========================================================================
    ' Sanity Checks
    ' ==========================================================================
    ' All of the previous adjustments use the same logic to set the winter and
    ' summer temperatures. Now determine what season it is to prevent the wrong
    ' system from running.
    If (TemperatureHigh < 45) Then
      SetCool += 20
    ElseIf (TemperatureHigh > 60 And TemperatureLow > 50) Then
      SetHeat -= 10
    ElseIf (TemperatureHigh > 50 And TemperatureLow > 40) Then
      SetHeat -= 2
    End If

    ' Prevent any of the adjustments from getting too far in either extreme.
    If ( SetHeat < 55 ) Then
      SetHeat = 55
    End If

    If ( SetCool > 80 ) Then
      SetCool = 80
    End If

    ' ==========================================================================
    ' Set temperatures
    ' ==========================================================================
    ' Set the Heat Setpoint
    If (hs.DeviceValue(SetPointHeating) <> SetHeat) Then
      MakeTheCapiHappy("(value) F", SetPointHeating, SetHeat)
    End If

    ' Set the Cool Setpoint
    If (hs.DeviceValue(SetPointCooling) <> SetCool) Then
      MakeTheCapiHappy("(value) F", SetPointCooling, SetCool)
    End If

    ' If the HVAC system is OFF, make sure the fan is set to AUTO
    If (HvacMode = 0) Then
      SetMode = 0
    End If

    ' Set the Fan Mode
    If (hs.DeviceValue(FanMode) <> SetMode) Then
      If (SetMode = 0) Then
        MakeTheCapiHappy("Auto", FanMode, SetMode)
      Else
        MakeTheCapiHappy("On", FanMode, SetMode)
      End If
    End If

    ' ==========================================================================
    ' Output
    ' ==========================================================================
    If (SetMode = 0) Then
      hs.WriteLog("HvacController", "HVAC mode was set to (Cool: " & SetCool & " | Heat: " & SetHeat & " | Fan: Auto | Temp: " & hs.DeviceValue(43) & " | Avg: " & AverageTemperature &")")
    Else
      hs.WriteLog("HvacController", "HVAC mode was set to (Cool: " & SetCool & " | Heat: " & SetHeat & " | Fan: On | Temp: " & hs.DeviceValue(43) & " | Avg: " & AverageTemperature &")")
    End If
  End If
End Sub

' ==============================================================================
' Functions
' ==============================================================================
' TimeOfDayAdjustments
' 
' Adjust the set points for the thermostat based on the time of day.
' ==============================================================================
Sub TimeOfDayAdjustments(CurrentHour As Integer,ByRef SetHeat As Integer,ByRef SetCool As Integer,ByRef SetMode As Integer,DesiredWinter As Integer,DesiredSummer As Integer)
  Select Case CurrentHour
    Case 0 to 4
      SetHeat = DesiredWinter - 3
      SetCool = DesiredSummer - 1
      SetMode = 0
    Case 5 to 8
      SetHeat = DesiredWinter + 1
      SetCool = DesiredSummer - 1
      SetMode = 0
    Case 9
      ' Summer: Cool down a little bit more after everyone is awake and ready,
      ' but before the heat of the day starts.
      SetHeat = DesiredWinter
      SetCool = DesiredSummer - 2
      SetMode = 0
    Case 10 to 13
      SetHeat = DesiredWinter
      SetCool = DesiredSummer
      SetMode = 0
    Case 14 to 18
      ' Don't try to fight the afternoon sun. Let the temperature rise an extra
      ' degree if it needs to. Keep the fan running during this time.
      SetHeat = DesiredWinter - 1
      SetCool = DesiredSummer + 1
      SetMode = 1
    Case 19 to 20
      SetHeat = DesiredWinter
      SetCool = DesiredSummer
      SetMode = 0
    Case 21 to 23
      SetHeat = DesiredWinter - 3
      SetCool = DesiredSummer - 1
      SetMode = 0
  End Select
End Sub

' ==============================================================================
' ExtremeTemperatureAdjustments
'
' Adjust the desired temperatures based on temperature extremes and time of day
' ==============================================================================
Sub ExtremeTemperatureAdjustments(OutsideTemperature As Integer,TemperatureHigh As Integer,CurrentHour As Integer, CurrentHumidity As Integer, ByRef DesiredWinter As Integer,ByRef DesiredSummer As Integer)
  If (CurrentHour > 3 And CurrentHour < 9 And TemperatureHigh > 87) Then
    ' In the morning, if the high is over 90, lower the temperature by 1 degree
    ' to cool the house down before it gets warmer to make it easier to maintain
    ' the desired temperature later.
    DesiredSummer = DesiredSummer - 1
  ElseIf (CurrentHour > 3 And CurrentHour < 16 And TemperatureHigh > 50) Then
    ' During the day, if the high is above 50, lower the winter temperature by
    ' 1 degree to prevent the house from getting too warm from the sun.
    DesiredWinter = DesiredWinter - 1
  End If

  ' After the above section makes adjustments for the time and high for the day,
  ' adjust for the current temperature outside.
  ' The ASHRAE design temperatures (http://cms.ashrae.biz/weatherdata/STATIONS/726580_p.pdf)
  '   1% Dry Bulb
  '     Summer: 87.8
  '     Winter: -19.7
  '   0.4%/99.6%
  '     Summer: 91.0
  '     Winter: -25.7
  If (OutsideTemperature > 91) Then
    DesiredSummer = DesiredSummer + 4
  ElseIf (OutsideTemperature > 87) Then
    DesiredSummer = DesiredSummer + 2
  ElseIf (OutsideTemperature < -20) Then
    DesiredWinter = DesiredWinter - 4
  ElseIf (OutsideTemperature < -10) Then
    DesiredWinter = DesiredWinter - 3
  ElseIf (OutsideTemperature < 0) Then
    DesiredWinter = DesiredWinter - 2
  ElseIf (OutsideTemperature < 10) Then
    DesiredWinter = DesiredWinter - 1
  End If

  ' High Humidity
  ' If the high is over 90 and the current RH is over 55%, add one degree
  ' to ease the load on the A/C
  If TemperatureHigh > 90 And CurrentHumidity > 55 Then
    DesiredSummer = DesiredSummer + 1
  End If
End Sub

' ==============================================================================
' AverageAdjustments
'
' Adjust the set points if there are dramatic differences between upstairs and
' downstairs
' ==============================================================================
Sub AverageAdjustments(HeatDifference As Double,CoolDifference As Double,FloorDifference As Double,AverageTemperature As Integer,ByRef SetHeat As Integer,ByRef SetCool As Integer,ByRef SetMode As Integer)
  ' Adjust Heating
  If (HeatDifference >= 3) Then
    If (AverageTemperature > SetHeat) Then
      SetHeat = SetHeat - 2
    Else
      SetHeat = SetHeat + 2
    End If
  Else If (HeatDifference >= 2) Then
    SetMode = 1
  End If

  ' Adjust Cooling
  If (CoolDifference >= 3) Then
    If (AverageTemperature > SetCool) Then
      SetCool = SetCool - 2
    Else
      SetCool = SetCool + 1
      SetMode = 1
    End If
  Else If (CoolDifference >= 2) Then
    SetMode = 1
  End If

  ' Adjust for differences in floor
  If (FloorDifference > 2) Then
    SetMode = 1
  End If
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
