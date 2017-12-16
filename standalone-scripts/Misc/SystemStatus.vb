' ==============================================================================
' Send Message
' ==============================================================================
Sub Main(Parm As Object)

	Dim Message As String = ""
	Dim SecurityStatus As Integer = hs.DeviceValue(28)
	Dim SecurityAlert As Integer = hs.DeviceValue(92)
	
	' ==========================================================================
	' Security Status
	' ==========================================================================
	Message = "Security System Status<br />"
	'Message = Message & "Security System Status<br />"
	'Message = Message & "<br />"
	If ( SecurityStatus = 0 ) Then
		Message = Message & "State: Disarmed<br />"
	ElseIf ( SecurityStatus = 50 ) Then
		Message = Message & "State: Perimeter<br />"
	Else
		Message = Message & "State: Armed<br />"
	End If
	
	If ( SecurityAlert = 100 ) Then
		Message = Message & "Alert: Active<br />"
	Else
		Message = Message & "Alert: Inactive<br />"
	End If
	
	' ==========================================================================
	' Environment Status
	' ==========================================================================
	Message = Message & "<br />Environment Status<br />"
	
	Message = Message & "Avg. Temp: " & hs.DeviceValue(118) & "<br />"
	
	If ( hs.DeviceValue(47) = 0 ) Then
		Message = Message & "State: Idle<br />"
	Else If ( hs.DeviceValue(47) = 1 ) Then
		Message = Message & "State: Heating<br />"
	Else
		Message = Message & "State: Cooling<br />"
	End If
	
	Message = Message & "Cool Set: " & hs.DeviceValue(49) & "<br />"
	
	Message = Message & "Heat Set: " & hs.DeviceValue(48) & "<br />"
	
	' ==========================================================================
	' Occupancy
	' ==========================================================================
	Message = Message & "<br />Occupancy Status<br />"
	
	Select Case hs.DeviceValue(75)
		Case 0
			Message = Message & "Occupancy: Vacation<br />"
		Case 25
			Message = Message & "Occupancy: Away<br />"
		Case 75
			Message = Message & "Occupancy: Occupied<br />"
		Case 100
			Message = Message & "Occupancy: Full<br />"
	End Select
	
	Dim People() As Integer = {22,24,25}
	Dim Value As String
	
	For Each Person As Integer In People
	
		Select Case hs.DeviceValue(Person)
			Case 0
				Value = "Vacation"
			Case 50
				Value = "Away"
			Case 100
				Value = "Home"
		End Select
		
		Message = Message & hs.DeviceName(Person).Replace("Trackers People",String.Empty) & ": " & Value & "<br />"
		
	Next
	
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