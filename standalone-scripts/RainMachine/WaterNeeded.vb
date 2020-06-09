' ==============================================================================
' This runs once a day and keeps track of the water received in the last 7 days
' ==============================================================================
Sub Main (Param As Object)
  ' Load Devices
  Dim DesiredWaterId As Integer = 408
  Dim DesiredWaterValue As Integer = hs.DeviceValue(DesiredWaterId)
  Dim HighTemperature As Integer = hs.DeviceValue(572)
  Dim LowTemperature As Integer = hs.DeviceValue(573)

  ' Adjust based on temperatures
  If LowTemperature > 45 Then
    If HighTemperature >= 85 Then
      hs.SetDeviceValueByRef(DesiredWaterId,DesiredWaterValue+2,True)
    ElseIf HighTemperature >= 60 Then
      hs.SetDeviceValueByRef(DesiredWaterId,DesiredWaterValue+1,True)
    End If
  End If
End Sub
