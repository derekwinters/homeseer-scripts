Sub Main(Parms as Object)
	' ==========================================================================
	' Variables
	' ==========================================================================
	' Monitors
	Dim MotionBasement As String = "112"
	
	' Trackers
	Dim Tracker        As String = "117"
	
	' ==========================================================================
	' Check Doors
	' ==========================================================================
	' Check basement motion sensor
	If ( hs.DeviceValue(MotionBasement) = 255 ) Then
	
		hs.SetDeviceValueByRef(Tracker,100,True)
	
	Else

		hs.SetDeviceValueByRef(Tracker,0,True)

	End If

	
End Sub