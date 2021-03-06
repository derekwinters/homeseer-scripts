' ==============================================================================
' This runs every hour to track the amount of rain received. The weatherXML
' device used as a reference tracks water in whole inches with an accuracy of
' two decimal places.
' ==============================================================================
Sub Main (Param As Object)
  Dim WaterLastHour As Double = hs.DeviceValueEx(381)
  Dim Day0 As Double = hs.DeviceValueEx(414)

  ' Calculate new value
  Day0 = Day0 + WaterLastHour
  hs.WriteLog("Irrigation Controller",WaterLastHour & " inches of rain in the last hour")

  ' Set the new value
  hs.SetDeviceValueByRef((414),Day0,True)
End Sub
