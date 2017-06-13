Sub Main(Parms as Object)
	' ==========================================================================
	' Variables
	' ==========================================================================
	' Door Monitors
	Dim DoorsGarageDoor1   As String = "72"
    Dim DoorsPatioDoorDown As String = "68"
    Dim DoorsFrontDoor As String = "201"
    ' Window Monitors

    ' Motion Monitors

    ' Trackers
    Dim TrackerDoors       As String = "97"

    ' ==========================================================================
    ' Check Doors
    ' ==========================================================================

    If (hs.DeviceValue(DoorsGarageDoor1) = 255) Then
        ' Check Garage Door
        hs.SetDeviceValueByRef(TrackerDoors, 100, True)
    ElseIf (hs.DeviceValue(DoorsPatioDoorDown) = 255) Then
        ' Check Patio Door
        hs.SetDeviceValueByRef(TrackerDoors, 100, True)
    ElseIf (hs.DeviceValue(DoorsFrontDoor) <> 255) Then
        ' Check Front Door Lock
        hs.SetDeviceValueByRef(TrackerDoors, 100, True)
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