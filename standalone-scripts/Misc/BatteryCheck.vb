Sub Main(Parms As Object)
    ' ==========================================================================
    ' Variables
    ' ==========================================================================
    ' Find sensors that are "Z-Wave Battery" and check their value
    Dim Device As Scheduler.Classes.DeviceClass
    Dim Enumerator As Scheduler.Classes.clsDeviceEnumeration
    Dim Body As String = ""
    Dim Total As Integer

    Enumerator = hs.GetDeviceEnumerator

    If (Enumerator Is Nothing) Then
        hs.WriteLog("Battery Check", "Error getting enumerator")
        Exit Sub
    End If

    Do While Not Enumerator.Finished
        Device = Enumerator.GetNext

        If (Device Is Nothing) Then
            Continue Do
        End If

        If (Device.Device_Type_String(hs) = "Z-Wave Battery") Then			
			hs.WriteLog("HVAC Automation", "Found a battery device (Name: " & Device.Location(hs) & " " & Device.Name(hs) & ", ReferenceID: " & Device.ref(hs) & ", Value: " & hs.DeviceValue(Device.ref(hs)) & ")")
            ' 0-100 are battery percentage values. 101-254 are invalid, and 255 is Battery Low Warning
			If ( ( hs.DeviceValue(Device.ref(hs)) < 10 Or hs.DeviceValue(Device.ref(hs)) > 100 ) And Device.ref(hs) <> "87" ) Then
				hs.WriteLog("HVAC Automation", "Alerting on device " & Device.ref(hs))
				Body = Body & Device.Location(hs) & " " & Device.Name(hs) & ": " & hs.DeviceValue(Device.ref(hs))
				Total = Total + 1
            Else
                hs.WriteLog("HVAC Automation", "Not alerting on device " & Device.ref(hs) & " value: " & hs.DeviceValue(Device.ref(hs)))
			End If
			
        End If

    Loop

    If (Total > 0) Then
        'SendMessage("Battery Alert",Body)
    End If

End Sub

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