Sub Main(Parms As Object)
  ' ==========================================================================
  ' Variables
  ' ==========================================================================
  ' Monitors
  Dim Sensors() As Integer = {341,439,463,458,335,329,213}

  ' Trackers
  Dim Tracker As String = "117"

  ' ==========================================================================
  ' Check Sensors
  ' ==========================================================================
  ' Set the default to No Motion (0)
  Dim Value As String = "0"

  ' Loop through sensors
  For Each Sensor As Integer In Sensors
    ' Some motion sensors only set Home Security to Motion Detected (8), but
    ' others set Binary Motion to (255) and Home Security to (8). Check if
    ' each sensor is Motion Detected (8).
    If hs.DeviceValue(Sensor) = 8 Then
      Value = 100
    End If
  Next

  ' Set the Motion Tracker
  hs.SetDeviceValueByRef(Tracker,Value,True)
End Sub
