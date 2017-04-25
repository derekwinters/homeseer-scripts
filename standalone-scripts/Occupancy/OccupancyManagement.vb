Sub Main(Parms as Object)
	' ==========================================================================
	' Variables
	' ==========================================================================
	' Phone Trackers
	Dim DerekPhone       As String  = "11"
	Dim AmyPhone         As String  = "10"
	' People Trackers
	Dim DerekTrack       As String  = "22"
	Dim AmyTrack         As String  = "24"
	Dim BabysitterTrack  As String  = "25"
	' Home Trackers
	Dim HomeOccupied     As String  = "26"
	Dim HomeOccupiedFull As String  = "27"
	Dim OccupancyMode    As String  = "75"
	' Device variables
	Dim TrackHome        As Integer = 100
	Dim TrackAway        As Integer = 50
	Dim TrackVacation    As Integer = 0
	Dim ModeFull         As Integer = 100
	Dim ModeOccupied     As Integer = 75
	Dim ModeAway         As Integer = 25
	Dim ModeVacation     As Integer = 0
	
	' ==========================================================================
	' Track Derek
	' ==========================================================================
	If ( ( hs.DeviceValue(DerekPhone) = 100 ) And ( hs.DeviceValue(DerekTrack) <> TrackHome ) ) Then
	
		' If device is online and Derek isn't set to home, change it
		hs.WriteLog("Occupancy", "Derek arrived")
		hs.SetDeviceValueByRef(DerekTrack,TrackHome,True)
		Babysitter(BabysitterTrack)
		
	Else If ( ( hs.DeviceValue(DerekPhone) <> 100 ) And ( hs.DeviceValue(DerekTrack) = TrackHome ) ) Then
	
		' If device is offline and Derek is set to home, change it
		hs.SetDeviceValueByRef(DerekTrack,TrackAway,True)
		hs.WriteLog("Occupancy", "Derek left")
		
	End If

	' ==========================================================================
	' Track Amy
	' ==========================================================================
	If ( ( hs.DeviceValue(AmyPhone) = 100 ) And ( hs.DeviceValue(AmyTrack) <> TrackHome ) ) Then
	
		' If device is online and Amy isn't set to home, change it
		hs.SetDeviceValueByRef(AmyTrack,TrackHome,True)
		hs.WriteLog("Occupancy", "Amy arrived")
		Babysitter(BabysitterTrack)
		
	Else If ( ( hs.DeviceValue(AmyPhone) <> 100 ) And ( hs.DeviceValue(AmyTrack) = TrackHome ) ) Then
	
		' If device is offline and Amy is set to home, change it
		hs.SetDeviceValueByRef(AmyTrack,TrackAway,True)
		hs.WriteLog("Occupancy", "Amy left")
		
	End If

	' ==========================================================================
	' Check Occupancy
	' ==========================================================================
	If ( ( hs.DeviceValue(AmyTrack) = TrackHome ) And ( hs.DeviceValue(DerekTrack) = TrackHome ) ) Then
	
		' Everyone is home
		hs.SetDeviceValueByRef(OccupancyMode,ModeFull,True)
		hs.WriteLog("Occupancy", "The house is fully occupied")
		
	Else If ( hs.DeviceValue(AmyTrack) = TrackHome And hs.DeviceValue(DerekTrack) = TrackVacation ) Then
		
		' Derek is on vacation, the house should be considered fully occupied
		hs.SetDeviceValueByRef(OccupancyMode,ModeFull,True)
		hs.WriteLog("Occupancy", "Derek is on vacation, the house is fully occupied")

	Else If	( hs.DeviceValue(AmyTrack) = TrackVacation And hs.DeviceValue(DerekTrack) = TrackHome ) Then
		
		' Amy is on vacation, the house should be considered fully occupied
		hs.SetDeviceValueByRef(OccupancyMode,ModeFull,True)
		hs.WriteLog("Occupancy", "Amy is on vacation, the house is fully occupied")
		
	Else If ( ( hs.DeviceValue(AmyTrack) = TrackHome ) Or ( hs.DeviceValue(DerekTrack) = TrackHome ) ) Then
	
		' Someone is home
		hs.SetDeviceValueByRef(OccupancyMode,ModeOccupied,True)
		hs.WriteLog("Occupancy", "The house is partially occupied")
		
	Else If ( hs.DeviceValue(BabysitterTrack) = TrackHome ) Then
		
		' Babysitter is home
		hs.SetDeviceValueByRef(OccupancyMode,ModeOccupied,True)
		hs.WriteLog("Occupancy", "Babysitter is home")
		
	Else If ( ( hs.DeviceValue(AmyTrack) = TrackVacation ) And ( hs.DeviceValue(DerekTrack) = TrackVacation ) )
	
		' Vacation Mode
		hs.SetDeviceValueByRef(OccupancyMode,ModeVacation,True)
		hs.WriteLog("Occupancy", "Vacation mode enabled")
	
	Else
	
		' No one is home
		hs.SetDeviceValueByRef(OccupancyMode,ModeAway,True)
		hs.WriteLog("Occupancy", "No one is home")
		
	End If
	
End Sub

Sub Babysitter(BabysitterTrack As Integer)

	If ( hs.DeviceValue(BabysitterTrack) = 100 ) Then
		
		hs.SetDeviceValueByRef(BabysitterTrack,0,True)
		hs.WriteLog("Occupancy", "Babysitter tracking stopped")
		
	End If

End Sub