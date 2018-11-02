' ==============================================================================
' Security System State Management
' ==============================================================================
' Manage the state of the security system (Armed, Perimeter, Disarmed) based on
' a variety of criteria.
' ==============================================================================
Sub Main(parms As Object)
	' ==========================================================================
	' Variables
	' ==========================================================================
	' Occupancy
	Dim OccupancyMode         As String = "75"
	Dim OccupancyVacation     As Integer = 0
	Dim OccupancyAway         As Integer = 25
	Dim OccupancyOccupied     As Integer = 75
	Dim OccupancyFull         As Integer = 100
	' Save current hour (24-hour clock, no leading zero)
	Dim CurrentTime           As Integer = Hour(Now())
	' Security System
	Dim SecurityMode          As String = "28"
	Dim SecurityDisarmedStart As Integer = 5   '  5:00 AM
	Dim SecurityDisarmedEnd   As Integer = 22  ' 10:00 PM
	Dim SecurityModeArmed     As Integer = 100
	Dim SecurityModeDisarmed  As Integer = 0
	Dim SecurityModePerimeter As Integer = 50
	
	' ==========================================================================
	' Decision Tree
	' ==========================================================================
	If ( ( hs.DeviceValue(OccupancyMode) = OccupancyVacation ) OrElse ( hs.DeviceValue(OccupancyMode) = OccupancyAway ) ) Then
		' No one is home (vacation mode or away mode), fully arm the security system
		hs.SetDeviceValueByRef(SecurityMode,SecurityModeArmed,True)
		hs.WriteLog("Security System", "Correct state determined to be ARMED")
		
	Else If ( ( CurrentTime >= SecurityDisarmedStart) And ( CurrentTime < SecurityDisarmedEnd ) ) Then
		' We've verified someone is home. If the time is during the disarm period, disarm the system
		hs.SetDeviceValueByRef(SecurityMode,SecurityModeDisarmed,True)
		hs.WriteLog("Security System", "Correct state determined to be DISARMED")
		
		' Clear alert state if intrusion
		If ( hs.DeviceValue("92") = 100 ) Then
            hs.WriteLog("Security System", "Disabling Intrusion")
			hs.SetDeviceValueByRef("92",0,True)
		End If
		
		' If 'Heading Home' is enabled, disable it.
		If ( hs.DeviceValue("211") = 100 ) Then
			hs.WriteLog("Security System", "HVAC Disabling Heading Home")
			hs.SetDeviceValueByRef("211",0,True)
		End If
		
	Else
		' The time is during the perimeter period. Set accordingly
		hs.SetDeviceValueByRef(SecurityMode,SecurityModePerimeter,True)
		hs.WriteLog("Security System", "Correct state determined to be PERIMETER")
		
		' If 'Heading Home' is enabled, disable it.
		If ( hs.DeviceValue("211") = 100 ) Then
			hs.WriteLog("Security System", "HVAC Disabling Heading Home")
			hs.SetDeviceValueByRef("211",0,True)
		End If
		
	End If
	
end Sub