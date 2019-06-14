' ==============================================================================
' MaintenanceTasks
' ==============================================================================
' Send a reminder to mow the lawn, and provide environmental details.
' ==============================================================================
Sub Main(Param As Object)
    hs.WriteLog("Maintenance Task", "Starting up mowing script")
	' ==========================================================================
	' Get some basic info
	' ==========================================================================
    Dim TemperatureHigh As Integer = hs.DeviceValue("32")
    Dim TemperatureLow As Integer = hs.DeviceValue("33")
    Dim Message as String = ""
    Dim DeviceRefNum As Integer = 321
    
    ' ==========================================================================
	' Craft the message
	' ==========================================================================
    
	' ==========================================================================
    ' Weather check
    ' ==========================================================================
    If ( TemperatureHigh > 50 And TemperatureLow > 40 ) Then
        Message = "The lawn was last mowed on " & hs.DeviceLastChangeRef(DeviceRefNum) & "."
        Message = Message & "<br /><br />Sunset for tonight is at " & hs.sunset & "."
        Message = Message & "<br /><br />Expected temperatures: High of " & TemperatureHigh & "F and Low of " & TemperatureLow & "F."
        Message = Message & "<br /><br />Reply 'Task " & DeviceRefNum & " Complete' to reset the timer."

        ' ======================================================================
        ' Notification
        ' ======================================================================
        
        SendMessage("Maintenance Task",Message)
        hs.WriteLog("Maintenance Task", Message)
    Else
        hs.WriteLog("MaintenanceReminder","Skipping mowing notification due to temperature restrictions.")
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