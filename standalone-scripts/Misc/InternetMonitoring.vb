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
			SendMessage("Internet Monitoring",Message)
			 
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

' ==============================================================================
' Send Message
' ==============================================================================
Sub SendMessage(SubjectString As String, MessageString As String)

    ' Set up enumerator
	Dim Device As Scheduler.Classes.DeviceClass
	Dim Enumerator As Scheduler.Classes.clsDeviceEnumeration

	' Start
	Enumerator = hs.GetDeviceEnumerator

	' Check
	If (Enumerator Is Nothing) Then
		hs.WriteLog("MMS Messaging", "Error getting enumerator")
		Exit Sub
	End If

	' Loop and send messages
	Do While Not Enumerator.Finished
		Device = Enumerator.GetNext

		If (Device Is Nothing) Then
			Continue Do
		End If

		' If the device type string is MMSPhoneNumber and the device is on,
		' send the message. Checking if the device is off allows for easily
		' disabling sending to a specific number without modifying scripts.
		If (Device.Device_Type_String(hs) = "MMSPhoneNumber") Then
			If (hs.DeviceValue(Device.ref(hs)) = 100) Then
				hs.SendEmail(hs.DeviceString(Device.ref(hs)), hs.DeviceString(174), "", "", SubjectString, MessageString, "")
			Else
				hs.WriteLog("MMS Messaging", "Device is disabled for messaging (ReferenceID: " & Device.ref(hs) & ")")
			End If
		End If

	Loop
	
End Sub