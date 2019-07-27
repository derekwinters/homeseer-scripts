' ==============================================================================
' This runs once a day and keeps track of the water received in the last 7 days
' ==============================================================================
Sub Main (Param As Object)
  hs.SetDeviceValueByRef((420),hs.DeviceValue(419),True)
  hs.SetDeviceValueByRef((419),hs.DeviceValue(418),True)
  hs.SetDeviceValueByRef((418),hs.DeviceValue(417),True)
  hs.SetDeviceValueByRef((417),hs.DeviceValue(416),True)
  hs.SetDeviceValueByRef((416),hs.DeviceValue(415),True)
  hs.SetDeviceValueByRef((415),hs.DeviceValue(414),True)
  hs.SetDeviceValueByRef((414),0,True)
End Sub
