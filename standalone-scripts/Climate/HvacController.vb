' ==============================================================================
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
' 2017-05-17
'     Changed set points to virtual devices for easier control in phone app
' 2017-06-12
'     Added fix for automatic heat/cool mode
' 2017-08-26
'     Additional scaling based on outside temperature
' 2017-12-01
'     Refactor time-based determinations
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
    Dim ThermostatHold As Integer = hs.DeviceValue("202")

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

	' ==============================================================================
	' If Heading Home is enabled, mock Occupancy
	' ==============================================================================
	If ( hs.DeviceValue("211") = 100 ) Then
		hs.WriteLog("HVAC Automation", "Heading home is enabled. Mocking OccupancyMode to 100")
		OccupancyMode = 100
	End If
	
    If (ThermostatHold = 0) Then
	
		If (OccupancyMode = 100 OrElse OccupancyMode = 75 ) Then
		
			' Use the value of the tracking devices to set the thermostat based
			' on the time of day.
			Select Case CurrentHour
				Case < 3
					' Start cooling earlier in the morning to help get the temp
					' lower before it starts to warm up outside.
					SetHeat = hs.DeviceValue(186) ' Night Heat
					SetCool = hs.DeviceValue(177) ' Morning Cool
					SetMode = hs.DeviceValue(187) ' Night Mode
				Case < 5
					SetHeat = hs.DeviceValue(186)
					SetCool = hs.DeviceValue(188)
					SetMode = hs.DeviceValue(187)
				Case < 9
					SetHeat = hs.DeviceValue(178)
					SetCool = hs.DeviceValue(177)
					SetMode = hs.DeviceValue(179)
				Case >= 23
					SetHeat = hs.DeviceValue(186)
					SetCool = hs.DeviceValue(188)
					SetMode = hs.DeviceValue(187)
				Case >= 19
					SetHeat = hs.DeviceValue(183)
					SetCool = hs.DeviceValue(185)
					SetMode = hs.DeviceValue(184)
				Case Else
					' General daytime values
					SetHeat = hs.DeviceValue(180)
					SetCool = hs.DeviceValue(182)
					SetMode = hs.DeviceValue(181)
			End Select
		
			' If it's earlier than 4PM and the high for the day is higher than
			' 50, lower the temperature by 1 to help from making the house too
			' hot during the day.
			If (TemperatureHigh > 50 And CurrentHour < 16) Then
				SetHeat = SetHeat - 1
			End If
			
		Else If ( OccupancyMode = 0 ) Then
            ' Vacation
            SetHeat = 50
            SetCool = 85
            SetMode = 0
		Else
            ' Away
            SetHeat = 65
            SetCool = 75
            SetMode = 0
		End If

        ' ==========================================================================
        ' Modify the set points based on the outside temperature
        ' ==========================================================================
        If (OutsideTemperature > 100) Then
            SetCool = SetCool + 2
		ElseIf (OutsideTemperature > 80) Then
			SetCool = SetCool + 1
		ElseIf (OutsideTemperature < 0) Then
			SetHeat = SetHeat - 2
		ElseIf (OutsideTemperature < 10) Then
			SetHeat = SetHeat - 1
        End If

        ' ==========================================================================
        ' Additional weather forecast alterations
        ' ==========================================================================
        ' Setting the thermostat to automatic heat/cool mode causes unneccessary use
        ' during the spring and fall. These alterations should allow the use of auto
        ' mode selection year round without issue.
        ' ==========================================================================
        ' If the high for the day is above 60, drop the heat temp by 10 degrees
        ' If the high for the day is below 50, raise the cool temp by 20 degrees
        If (TemperatureHigh > 60) Then
            SetHeat = SetHeat - 10
        ElseIf (TemperatureHigh < 50) Then
            SetCool = SetCool + 20
        End If

        ' ==========================================================================
        ' Average Temperature Alterations
        ' ==========================================================================
		' If the system is not active, check if the set points should be adjusted
		' based on the AverageTemperature. If the difference is >= 2, alter the set
		' point. If the difference is >= 1, turn the fan on.
        If (CurrentOperatingState = 0) Then
            If (CurrentOperatingMode = 1) Then
				If (Math.Abs(AverageTemperature - SetHeat) >= 2) Then
					If (AverageTemperature > SetHeat) Then
						SetHeat = SetHeat - 2
					Else
						SetHeat = SetHeat + 2
					End If
				ElseIf (Math.Abs(AverageTemperature - SetHeat) >= 1) Then
					SetMode = 1
				End If
            ElseIf (CurrentOperatingMode = 2) Then
				If (Math.Abs(AverageTemperature - SetCool) >= 2) Then
					If (AverageTemperature > SetCool) Then
						SetCool = SetCool - 2
					Else
						SetCool = SetCool + 2
					End If
				ElseIf (Math.Abs(AverageTemperature - SetCool) >= 1) Then
					SetMode = 1
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
            hs.WriteLog("HVAC Automation", "HVAC mode was set to (Cool: " & SetCool & ", Heat: " & SetHeat & ", Fan: Auto, Avg: " & AverageTemperature &")")
        Else
            hs.WriteLog("HVAC Automation", "HVAC mode was set to (Cool: " & SetCool & ", Heat: " & SetHeat & ", Fan: On, Avg: " & AverageTemperature &")")
        End If
    Else
        hs.WriteLog("HVAC Automation", "Thermostat HOLD is enabled. No changes made.")
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