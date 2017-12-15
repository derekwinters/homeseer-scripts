Sub Main(Parms As Object)
    ' ==========================================================================
    ' Variables
    ' ==========================================================================
    ' Find sensors that are "Z-Wave Temperature" and count the total
    Dim Device As Scheduler.Classes.DeviceClass
    Dim Enumerator As Scheduler.Classes.clsDeviceEnumeration
    ' Math
    Dim Sum As Integer
    Dim Total As Integer
    ' Trackers
    Dim Tracker As String = "118"

    Enumerator = hs.GetDeviceEnumerator

    If (Enumerator Is Nothing) Then
        hs.WriteLog("HVAC Automation", "Error getting enumerator")
        Exit Sub
    End If

    Do While Not Enumerator.Finished
        Device = Enumerator.GetNext

        If (Device Is Nothing) Then
            Continue Do
        End If

        If (Device.Device_Type_String(hs) = "Z-Wave Temperature") Then
            hs.WriteLog("HVAC Automation", "Found a temperature device (ReferenceID: " & Device.ref(hs) & ", Value: " & hs.DeviceValue(Device.ref(hs)) & ")")
            Sum = Sum + hs.DeviceValue(Device.ref(hs))
            Total = Total + 1
        End If

    Loop

    ' ==========================================================================
    ' Math
    ' ==========================================================================
    Dim Average As Double = Sum / Total

    hs.SetDeviceValueByRef(Tracker, Average, True)

    hs.WriteLog("HVAC Automation", "Average home temperature is " & Average & " F")
End Sub