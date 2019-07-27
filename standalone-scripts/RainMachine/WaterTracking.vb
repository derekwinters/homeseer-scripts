' ==============================================================================
' This runs once a day and keeps track of the water received in the last 7 days
' ==============================================================================
Sub Main (Param As Object)
  ' Store the existing values
  Dim Day6 As Integer
  Dim Day5 As Integer = hs.DeviceValue(419)
  Dim Day4 As Integer = hs.DeviceValue(418)
  Dim Day3 As Integer = hs.DeviceValue(417)
  Dim Day2 As Integer = hs.DeviceValue(416)
  Dim Day1 As Integer = hs.DeviceValue(415)
  Dim Day0 As Integer = hs.DeviceValue(414)

  ' Calculate new values
  Day6 = Day0 + Day1 + Day2 + Day3 + Day4 + Day5
  Day5 = Day0 + Day1 + Day2 + Day3 + Day4
  Day4 = Day0 + Day1 + Day2 + Day3
  Day3 = Day0 + Day1 + Day2
  Day2 = Day0 + Day1
  Day1 = Day0
  Day0 = 0

  ' Set the new values
  hs.SetDeviceValueByRef((420),hs.Day6,True)
  hs.SetDeviceValueByRef((419),hs.Day5,True)
  hs.SetDeviceValueByRef((418),hs.Day4,True)
  hs.SetDeviceValueByRef((417),hs.Day3,True)
  hs.SetDeviceValueByRef((416),hs.Day2,True)
  hs.SetDeviceValueByRef((415),hs.Day1,True)
  hs.SetDeviceValueByRef((414),0,True)
End Sub
