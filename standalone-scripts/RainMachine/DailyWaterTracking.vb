' ==============================================================================
' This runs every hour to track the amount of rain received. The weatherXML
' device used as a reference tracks water in whole inches with an accuracy of
' two decimal places.
' ==============================================================================
Sub Main (Param As Object)
  Dim WaterLastHour As Double = hs.DeviceValue(381)
  Dim Day0 As Integer = hs.DeviceValue(414)

  ' Calculate new value
  Day0 = Day0 + Math.Round((WaterLastHour*100),0,MidpointRounding.AwayFromZero)

  ' Set the new value
  hs.SetDeviceValueByRef((414),Day0,True)
End Sub
