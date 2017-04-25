Sub Main(Param As Object)

	' ==========================================================================
	' Variables
	' ==========================================================================
	Dim Log01  As String = "119"
	Dim Time01 As String = "121"
	Dim Log02  As String = "120"
	Dim Time02 As String = "122" ''
	Dim Log03  As String = "123"
	Dim Time03 As String = "124"
	Dim Log04  As String = "125"
	Dim Time04 As String = "126"
	Dim Log05  As String = "127"
	Dim Time05 As String = "128"
	Dim Time   As String = DateTime.Now.ToString("h:mm tt")
	
	' ==========================================================================
	' Move entries around
	' ==========================================================================
	If ( Param <> "" ) Then

		' Move Log 04 to Log 05
		hs.SetDeviceString(Log05, hs.DeviceString(Log04), True)
		hs.SetDeviceString(Time05,hs.DeviceString(Time04),True)
	
		' Move Log 03 to Log 04
		hs.SetDeviceString(Log04, hs.DeviceString(Log03), True)
		hs.SetDeviceString(Time04,hs.DeviceString(Time03),True)
	
		' Move Log 02 to Log 03
		hs.SetDeviceString(Log03, hs.DeviceString(Log02), True)
		hs.SetDeviceString(Time03,hs.DeviceString(Time02),True)
		
		' Move Log 01 to Log 02
		hs.SetDeviceString(Log02, hs.DeviceString(Log01), True)
		hs.SetDeviceString(Time02,hs.DeviceString(Time01),True)
		
		' Set Log 01 to Parma
		hs.SetDeviceString(Log01, Param,True)
		hs.SetDeviceString(Time01,Time, True)
	
	End If
	

End Sub