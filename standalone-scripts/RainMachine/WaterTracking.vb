' ==============================================================================
' This runs once a day and keeps track of the water received in the last 7 days
' ==============================================================================
Sub Main (Param As Object)
  ' Store the existing values
  Dim Day6 As Double
  Dim Day5 As Double = hs.DeviceValueEx(419)
  Dim Day4 As Double = hs.DeviceValueEx(418)
  Dim Day3 As Double = hs.DeviceValueEx(417)
  Dim Day2 As Double = hs.DeviceValueEx(416)
  Dim Day1 As Double = hs.DeviceValueEx(415)
  Dim Day0 As Double = hs.DeviceValueEx(414)

  ' Calculate new values
  Day6 = Day0 + Day5
  Day5 = Day0 + Day4
  Day4 = Day0 + Day3
  Day3 = Day0 + Day2
  Day2 = Day0 + Day1
  Day1 = Day0
  Day0 = 0

  ' Set the new values
  hs.SetDeviceValueByRef((420),Day6,True)
  hs.SetDeviceValueByRef((419),Day5,True)
  hs.SetDeviceValueByRef((418),Day4,True)
  hs.SetDeviceValueByRef((417),Day3,True)
  hs.SetDeviceValueByRef((416),Day2,True)
  hs.SetDeviceValueByRef((415),Day1,True)
  hs.SetDeviceValueByRef((414),Day0,True)
End Sub
