Sub Main(Parms as Object)
	' ==========================================================================
	' Variables
	' ==========================================================================
	' Trackers
	Dim DerekTrack       As String  = "22"
	Dim AmyTrack         As String  = "24"
	Dim VacationDuration As String  = "355"
	Dim OccupancyMode    As String  = "75"
	
	' ==========================================================================
	' Determine Duration
	' ==========================================================================
  ' VacationDuration is a number of days that vacation mode should remain
  ' active. Every morning if vacation mode is enabled, this script should
  ' determine if vacation mode should be disabled or remain active.

  ' First, determine how long vacation mode has been active. This function
  ' returns the number of minutes since a device was last changed.
  Dim TimeInVacation = hs.DeviceTime(hs.DeviceByRef(OccupancyMode))
  hs.WriteLog("Occupancy", "Time in vacation mode: " & TimeInVacation)

  ' Get vacation duration
  Dim VacationDurationDays = hs.DeviceValueByRef(VacationDuration)
  hs.WriteLog("Occupancy", "Vacation duration: " & VacationDurationMinutes)

  ' To make sure the that this triggers on the morning of returning home,
  ' round TimeInVacation to whole days and remove any extra time. Do this
  ' by dividing by the number of minutes in a day and store as an integer
  ' to remove the decimal.
  Dim TimeInVacationDays As Integer = TimeInVacation / 1440
  hs.WriteLog("Occupancy", "Days in vacation mode: " & TimeInVacationDays)

  ' Check if duration has been exceeded and change the trackers from
  ' vacation to away
  If (TimeInVacationDays > VacationDurationDays And VacationDurationdays > 0) Then
    hs.SetDeviceValueByRef(AmyTrack,25,True)
    hs.SetDeviceValueByRef(DerekTrack,25,True)
  End If
End Sub
