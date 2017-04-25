Sub Main(Parms as Object)
	' ==========================================================================
	' Variables
	' ==========================================================================
	' Door Monitors
	Dim DoorsGarageDoor1   As String = "72"
	Dim DoorsPatioDoorDown As String = "68"
	' Window Monitors
	
	' Motion Monitors
	
	' Trackers
	Dim TrackerDoors       As String = "97"
	
	' ==========================================================================
	' Check Doors
	' ==========================================================================
	' Check Garage Door
	If ( hs.DeviceValue(DoorsGarageDoor1) = 255 ) Then
		hs.SetDeviceValueByRef(TrackerDoors,100,True)
	' Check Patio Door
	Else If ( hs.DeviceValue(DoorsPatioDoorDown) = 255 ) Then
		hs.SetDeviceValueByRef(TrackerDoors,100,True)
	' Check Front Door Lock
	Else
		hs.SetDeviceValueByRef(TrackerDoors,0,True)
	End If
	' ==========================================================================
	' Check Windows
	' ==========================================================================
	
	
	' ==========================================================================
	' Check Motion
	' ==========================================================================
	
	
End Sub