' ==============================================================================
' Send Message
' ==============================================================================
Sub Main(Parm As Object)

	Dim Message As String = ""
	Dim SecurityStatus As Integer = hs.DeviceValue(92)
	
	' ==========================================================================
	' Security Status
	' ==========================================================================
	Message = "Security System Status<br />"
	If ( SecurityStatus = 0 ) Then
		Message = Message & "Security State: Disarmed<br />"
	ElseIf ( SecurityStatus = 50 ) Then
		Message = Message & "Security State: Perimeter<br />"
	Else
		Message = Message & "Security State: Armed<br />"
	End If
	
	If ( hs.DeviceValue(28) = 100 ) Then
		Message = Message & "Security Alert: Active<br />"
	Else
		Message = Message & "Security Alert: Inactive<br />"
	End If
	
	' ==========================================================================
	' Environment Status
	' ==========================================================================
	Message = Message & "<br />Environment Status<br />"
	
	Message = Message & "Avg. Temperature: " & hs.DeviceValue(118) & "<br />"
	
	If ( hs.DeviceValue(47) = 0 ) Then
		Message = Message & "Current State: Idle<br />"
	Else If ( hs.DeviceValue(47) = 1 ) Then
		Message = Message & "Current State: Heating<br />"
	Else
		Message = Message & "Current State: Cooling<br />"
	End If
	
	Message = Message & "Cool Set Point: " & hs.DeviceValue(49) & "<br />"
	
	Message = Message & "Cool Set Point: " & hs.DeviceValue(48) & "<br />"
	
	' Send Message
	SendMessage("System Status",Message)
	
End Sub

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