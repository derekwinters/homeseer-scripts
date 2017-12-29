Sub Main(Parms as Object)
	' ==========================================================================
	' Variables
	' ==========================================================================
	' Monitors
	Dim Sensors() As Integer = {112,215}
	
	' Trackers
	Dim Tracker        As String = "117"
	
	' ==========================================================================
	' Check Sensors
	' ==========================================================================
	' Set the default to No Motion (0)
	Dim Value As String = 0
	
	' Loop through sensors
	For Each Sensor As Integer In Sensors
		' If the sensor is On-Open-Motion
		If ( hs.DeviceValue(Sensor) = 255 ) Then
			Value = 100
		End If
	Next

	' Set the Motion Tracker
	hs.SetDeviceValueByRef(Tracker,Value,True)
	
End Sub