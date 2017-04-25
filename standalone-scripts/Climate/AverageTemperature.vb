Sub Main(Parms as Object)
    ' ==========================================================================
    ' Variables
    ' ==========================================================================
    ' Sensors
    Dim TempBasement As Integer = hs.DeviceValue("113")
    Dim Thermostat As Integer = hs.DeviceValue("43")
    Dim Total As Integer = 2

    ' Trackers
    Dim Tracker As String = "118"

    ' ==========================================================================
    ' Math
    ' ==========================================================================
    Dim Average As Integer = (TempBasement + Thermostat) / Total

    hs.SetDeviceValueByRef(Tracker,Average,True)
	
	hs.WriteLog("HVAC Automation", "Average home temperature is " & Average & " F")
	
End Sub