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
    hs.WriteLog("HVAC Automation", "Thermostat HOLD is enabled. No changes made.")
  Else
    ' ==========================================================================
    ' Set Variables
    ' ==========================================================================
    ' Outside Temperature from WeatherXML
    Dim OutsideTemperature As Integer = hs.DeviceValue("7")
    Dim TemperatureHigh As Integer = hs.DeviceValue("32")
    Dim TemperatureLow As Integer = hs.DeviceValue("33")
    Dim HvacMode As Integer = hs.DeviceValue("46")

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

    ' Current Operating State
    Dim CurrentOperatingState As Integer = hs.DeviceValue("47")

    ' Current Operating Mode
    Dim CurrentOperatingMode As Integer = hs.DeviceValue("46")

    ' ==========================================================================
    ' If Heading Home is enabled, mock Occupancy
    ' ==========================================================================
    If ( hs.DeviceValue("211") = 100 ) Then
      hs.WriteLog("HVAC Automation", "Heading home is enabled. Mocking OccupancyMode to 100")
      OccupancyMode = 100
    End If

    ' ==========================================================================
    ' Initialize the desired temperature
    ' ==========================================================================
    Dim DesiredWinter = 72
    Dim DesiredSummer = 73

    ' ==========================================================================
    ' Adjust the desired temperature for extreme temperatures
    ' ==========================================================================
    ' Modify the set points based on the outside temperature
    If (OutsideTemperature > 100) Then
      DesiredSummer = DesiredSummer + 2
    ElseIf (OutsideTemperature > 80) Then
      DesiredSummer = DesiredSummer + 1
    ElseIf (OutsideTemperature < 0) Then
      DesiredWinter = DesiredWinter - 2
    ElseIf (OutsideTemperature < 10) Then
      DesiredWinter = DesiredWinter - 1
    End If

    ' ==========================================================================
    ' Adjust the desired temperature based on time of day
    ' ==========================================================================
    If (OccupancyMode = 100 OrElse OccupancyMode = 75 ) Then
      ' Start by determining desired temperature based on the current time
      If (CurrentHour >= 23 or CurrentHour < 5) Then
        ' 11PM - 5AM
        SetHeat = DesiredWinter - 2
        SetCool = DesiredSummer - 1
        SetMode = 0
      ElseIf (CurrentHour >= 5 and CurrentHour > 9) Then
        ' 5AM - 9AM
        SetHeat = DesiredWinter + 2
        SetCool = DesiredSummer - 1
        SetMode = 1
      ElseIf (CurrentHour >= 9 and CurrentHour > 19) Then
        ' 9AM - 7PM
        SetHeat = DesiredWinter
        SetCool = DesiredSummer
        SetMode = 0
      Else
        ' 7PM - 11PM
        SetHeat = DesiredWinter -1
        SetCool = DesiredSummer
        SetMode = 0
      End If

      ' If it's earlier than 4PM and the high for the day is higher than
      ' 50, lower the temperature by 1 to help from making the house too
      ' hot during the day.
      If (TemperatureHigh > 50 And CurrentHour < 16) Then
        SetHeat = SetHeat - 1
      End If

    Else If ( OccupancyMode = 0 ) Then
      ' Vacation
      SetHeat = DesiredWinter - 20
      SetCool = DesiredSummer + 8
      SetMode = 0
    Else
      ' Away
      SetHeat = DesiredWinter - 8
      SetCool = DesiredSummer + 3
      SetMode = 0
    End If


    ' ======================================================================
    ' Average Temperature Alterations
    ' ======================================================================
    ' If the system is not active, check if the set points should be
    ' adjusted based on the AverageTemperature.
    If (CurrentOperatingState = 0) Then
      Dim HeatDifference As Double = Math.Abs(AverageTemperature - SetHeat )
      Dim CoolDifference As Double = Math.Abs(AverageTemperature - SetCool )

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
    End If

    ' ======================================================================
    ' Set temperatures
    ' ======================================================================
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

    ' ======================================================================
    ' Output
    ' ======================================================================
    If (SetMode = 0) Then
      hs.WriteLog("HVAC Automation", "HVAC mode was set to (Cool: " & SetCool & " | Heat: " & SetHeat & " | Fan: Auto | Temp: " & hs.DeviceValue(43) & " | Avg: " & AverageTemperature &")")
    Else
      hs.WriteLog("HVAC Automation", "HVAC mode was set to (Cool: " & SetCool & " | Heat: " & SetHeat & " | Fan: On | Temp: " & hs.DeviceValue(43) & " | Avg: " & AverageTemperature &")")
    End If
  End If
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
