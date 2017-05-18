﻿' ==============================================================================
' HVAC Controller
' ==============================================================================
' Controller the set points on the thermostat based on home occupancy and on the
' current and forcasted weather for the day.
'
' Author: Derek Winters
'
' 2017-03-26
'     Original Script
' 2017-03-27
'     Modified set points for COOL
'     Created MakeTheCapiHappy function
' 2017-04-27
'     Included AverageTemperature in calculations for set points
'
' ==============================================================================
' Temperature Reasoning
' ==============================================================================
' Night Temps
'     At night, it is more comfortable to sleep at lower temperatures. It is
'     also cheaper to cool the house down when the outside temperature is lower.
'     For these reasons, set the temperatures lower at night.
' Morning Temps
'     Waking up is more comfortable when it's warmer. Raise the temperature in
'     the winter to warm the house up. Turn the fan to keep circulating air even
'     when the heat/cool cycle has finished.
' Day Temps
'     Keep the fan running during the day to help level the temperature through
'     the house.
'
' ==============================================================================
' Scheduling
' ==============================================================================
' This script is configured to run at the times specified below that trigger a
' change (ie. 7PM, 6AM, etc), and whenever the occupancy state of the house
' changes.
' The changes beginning On 2017-04-27 will add triggers To run the script when
' the average indoor temperature, Or the OutsideTemperature changes. This begins
' a set of changes to attempt to optimize the temperature set points to get a
' more universal temperature throughout the house, by modifying the set points
' when the average temperature is outside of a given range.
' ==============================================================================
' Help notes
' ==============================================================================
'
' CAPI control found at
' https://board.homeseer.com/showthread.php?p=1278809
'
' ==============================================================================

Sub Main(parms As Object)
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
    Dim AverageTemperature As Integer = hs.DeviceValue("118")

    ' Current Operating State
    Dim CurrentOperatingState As Integer = hs.DeviceValue("47")

    ' Current Operating Mode
    Dim CurrentOperatingMode As Integer = hs.DeviceValue("46")

    ' ==========================================================================
    ' Make the decisions based on time and occupancy
    ' ==========================================================================
    If (((OccupancyMode = 100 OrElse OccupancyMode = 75) And parms = "" And (CurrentHour < 5 OrElse CurrentHour >= 21)) OrElse parms = "Night") Then
        ' ======================================================================
        ' Night ( 2100 to 0500 )
        ' ======================================================================
        SetHeat = 68
        SetCool = 72
        SetMode = 0
    ElseIf (((OccupancyMode = 100 OrElse OccupancyMode = 75) And parms = "" And CurrentHour >= 5 And CurrentHour < 9) OrElse parms = "Morning") Then
        ' ======================================================================
        ' Morning ( 0500 to 0900 )
        ' ======================================================================
        If (TemperatureHigh > 50) Then
            SetHeat = 71
        Else
            SetHeat = 72
        End If
        SetCool = 72
        SetMode = 1
    ElseIf (((OccupancyMode = 100 OrElse OccupancyMode = 75) And parms = "" And CurrentHour >= 19 And CurrentHour < 21) OrElse parms = "Evening") Then
        ' ======================================================================
        ' Evening ( 1900 to 2100 )
        ' ======================================================================
        SetHeat = 70
        SetCool = 74
        SetMode = 0
    ElseIf (((OccupancyMode = 100 OrElse OccupancyMode = 75) And parms = "") OrElse (parms = "Day")) Then
        ' ======================================================================
        ' Day ( 0900 to 1900 )
        ' ======================================================================
        If (TemperatureHigh > 50) Then
            SetHeat = 69
        Else
            SetHeat = 70
        End If
        SetCool = 76
        SetMode = 1
    ElseIf (OccupancyMode = 0) Then
        ' ======================================================================
        ' Vacation
        ' ======================================================================
        SetHeat = 50
        SetCool = 85
        SetMode = 0
    Else
        ' ======================================================================
        ' Away
        ' ======================================================================
        SetHeat = 65
        SetCool = 79
        SetMode = 0
    End If

    ' ==========================================================================
    ' Modify the set points based on the outside temperature
    ' ==========================================================================
    If (OutsideTemperature > 100) Then
        SetCool = SetCool + 2
    ElseIf (OutsideTemperature < 0) Then
        SetHeat = SetHeat - 2
    End If

    ' ==========================================================================
    ' Average Temperature Alterations
    ' ==========================================================================
    If (CurrentOperatingState = 0) Then
        ' If the system is not active, check if we should adjust the set points
        ' to get the average temperature to within a reasonable range.
        If (CurrentOperatingMode = 1) Then
            ' If we are heating, compare the heat set point to the average
            If ((AverageTemperature > SetHeat) And ((AverageTemperature - SetHeat) > 2)) Then
                ' If the Average temperature is more than 2 degrees higher than
                ' the desired temperature, lower the set point 2 degrees
                SetHeat = SetHeat - 2
            ElseIf ((SetHeat > AverageTemperature) And ((SetHeat - AverageTemperature) > 2)) Then
                ' If the average temperature is more than 2 degrees below the
                ' desired temperature, raise the set point 2 degrees
                SetHeat = SetHeat + 2
            End If
        ElseIf (CurrentOperatingMode = 2) Then
            ' If we are cooling, compare the cool set point to the average
            If ((AverageTemperature > SetCool) And ((AverageTemperature - SetCool) > 2)) Then
                ' If the average temperature is more than 2 degrees higher than
                ' the desired temperature, lower the set point 2 degrees
                SetCool = SetCool - 2
            ElseIf ((SetCool > AverageTemperature) And ((SetCool - AverageTemperature) > 2)) Then
                ' If the average temperature is more than 2 degrees below the
                ' desired temperature, raise the set point 2 degrees
                SetCool = SetCool + 2
            End If
        End If
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
        hs.WriteLog("HVAC Automation", "HVAC mode was set to (Cool: " & SetCool & ", Heat: " & SetHeat & ", Fan: Auto)")
    Else
        hs.WriteLog("HVAC Automation", "HVAC mode was set to (Cool: " & SetCool & ", Heat: " & SetHeat & ", Fan: On)")
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