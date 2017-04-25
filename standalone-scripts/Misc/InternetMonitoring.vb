Sub Main(Param as Object)
	' ==========================================================================
	' Create Variables
	' ==========================================================================
	' Get the status of the internet tracking status device
	Dim InternetAccess  As Integer = hs.DeviceValue(37)
	' Get the status of the internet monitoring devices from BLLAN
	Dim Checker1        As Integer = hs.devicevalue(29)
	Dim Checker2        As Integer = hs.devicevalue(30)
	Dim Checker3        As Integer = hs.devicevalue(31)
	Dim Checker4        As Integer = hs.devicevalue(34)
	Dim Checker5        As Integer = hs.devicevalue(35)
	Dim Checker6        As Integer = hs.devicevalue(36)
	
	' ==========================================================================
	' Make decision
	' ==========================================================================
	If ( Checker1 = 100 OrElse Checker2 = 100 OrElse Checker3 = 100 OrElse Checker4 = 100 OrElse Checker5 = 100 OrElse Checker6 = 100 ) Then
	
		' If anything is online, the internet is online
		If ( InternetAccess <> 100 ) Then
		
			' Check when the internet went down
			Dim DownAt As String = hs.DeviceLastChange("T7")
			
			 ' Change InternetAccess
			 hs.SetDeviceValueByRef(37,100,True)

			' Build Message
			Dim Message As String = "The internet came online. It went down at " & DownAt
			
			' Log event
			hs.WriteLog("Internet Monitor", Message )
			
			' Send message
			hs.SendEmail(hs.DeviceString(168), hs.DeviceString(174), "", "", "Internet Monitoring", Message, "")
			hs.SendEmail(hs.DeviceString(167), hs.DeviceString(174), "", "", "Internet Monitoring", Message, "")
			 
		End If
		
	Else If ( Checker1 <> 100 And Checker2 <> 100 And Checker3 <> 100 And Checker4 <> 100 And Checker5 <> 100 And Checker6 <> 100 ) Then
	
		' If everything is offline, the internet is down
		If ( InternetAccess = 100 ) Then
		
			' Log event
			hs.WriteLog("Internet Monitor", "The internet went offline" )
			
			 ' Change InternetAccess
			 hs.SetDeviceValueByRef(37,0,True)
			 
		End If
		
	End If
	
end Sub