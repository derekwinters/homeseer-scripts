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
            hs.WriteLog("HVAC Automation", "Found a battery device (ReferenceID: " & Device.ref(hs) & ", Value: " & hs.DeviceValue(Device.ref(hs)) & ")")
            ' If the battery is below 15%, and it's not the power sensor which doesn't have a battery installed, add to alert body
            If (hs.DeviceValue(Device.ref(hs)) < 15 And Device.ref(hs) <> 87) Then
                Body = Body & Device.Location(hs) & " " & Device.Name(hs) & ": " & hs.DeviceValue(Device.ref(hs))
                Total = Total + 1
            End If

        End If

    Loop

    If (Total > 0) Then
        hs.SendEmail(hs.DeviceString(168), hs.DeviceString(174), "", "", "Low Battery Alert", Body, "")
        hs.SendEmail(hs.DeviceString(167), hs.DeviceString(174), "", "", "Low Battery Alert", Body, "")
    End If

End Sub