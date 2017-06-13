' ==============================================================================
' 
' ==============================================================================
' 
' 
' ==============================================================================
Sub Main(parms As Object)
	
	' Wait to allow time if someone just got home
	hs.WaitSecs(10)

	' ======================================================================
	' Variables
	' ======================================================================
	Dim Body  As String = ""
	Dim Alarm as Boolean = False
	
	' ==========================================================================
	' Initial Checks
	' ==========================================================================
	' If Security Current Mode <> Disarmed
	If ( hs.DeviceValue("28") <> 0 ) Then
		
		' ======================================================================
		' Armed and Perimeter Devices
		' ======================================================================
		' These devices will always trigger an alarm if they are changed when
		' the system is not disarmed.
		' ======================================================================
		' Check the garage door
		' ======================================================================
		If ( hs.DeviceValue("72") <> 0 ) Then
			Body = Body & " - Garage door is open - <br />"
			Alarm = True
		Else
			Body = Body & "Garage door is closed<br />"
		End If

        ' ======================================================================
        ' Check the front door
        ' ======================================================================
        If (hs.DeviceValue("201") = 255) Then
            Body = Body & "Front door is locked<br />"
        Else
            Body = Body & " - Front door is unlocked - <br />"
            Alarm = True
        End If

        ' ======================================================================
        ' Check the patio door
        ' ======================================================================
        If (hs.DeviceValue("68") <> 0) Then
            Body = Body & " - Patio door is open - <br />"
            Alarm = True
        Else
            Body = Body & "Patio door is closed<br />"
        End If


        ' ======================================================================
        ' Armed Only Devices
        ' ======================================================================
        ' These devices will be included in the report, but will only trigger an
        ' alarm if the system is fully armed.
        ' ======================================================================
        ' Family Room Motion Sensor
        ' ======================================================================
        If (hs.DeviceValue("112") <> 0) Then
            Body = Body & "- Motion in Family Room -<br />"
            If (hs.DeviceValue("28") = 100) Then
                Alarm = True
            End If
        Else
            Body = Body & "No motion in Family Room<br />"
        End If

    Else
            hs.WriteLog("Security System", "Security system event triggered. System is disarmed. No alert sent.")
	End If
	
	' ======================================================================
	' Alert
	' ======================================================================
	If ( Alarm = True ) Then
		' Set the alarm device to Intrusion
		If ( hs.DeviceValue("92") <> 100 ) Then
			hs.SetDeviceValueByRef("92",100,True)
		End If
	
		' Send email
		Body = "<br />" & Body & "<br />" & DateTime.Now.ToString("MMMM dd, HH:mm:ss")
		hs.WriteLog("Security System", "Security system triggered an alert.")
		
		hs.SendEmail(hs.DeviceString(168), hs.DeviceString(174),"","", "Security System Alert", Body,"")
		hs.SendEmail(hs.DeviceString(167), hs.DeviceString(174),"","", "Security System Alert", Body,"")
	Else
		hs.WriteLog("Security System", "Security system event triggered. All systems secured. No alert sent.")
	End If

end Sub