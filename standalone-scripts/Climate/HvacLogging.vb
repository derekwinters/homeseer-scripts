
Imports System.IO

Sub Main(parms As Object)

	' Use a new file each day
	Dim strFile As String = "C:\Logs\HVACLogs\HvacOperationLog-" & DateTime.Now.ToString("yyyyMMdd") & ".csv"

	' Check if the file exists
	Dim fileExists As Boolean = File.Exists(strFile)

	' Check if the file exists or not
	Using sw As New StreamWriter(File.Open(strFile, FileMode.OpenOrCreate))
		' If the file doesn't exist, add the header
		If Not fileExists Then
			sw.WriteLine( "DateTime,OperatingState,SetPointCooling,SetPointHeating,CurrentTemperature,OutsideTemperature" )
		End If
	End Using
	
	' Create the variables
	Dim CurrentDate              = DateTime.Now.ToString("HH:mm:ss") ' Format the current date and time
	Dim OperatingState As String = ""                                ' Thermostat On/Off
	Dim SetPointCooling          = hs.DeviceValue("49")              ' Thermostat Cooling Setpoint
	Dim SetPointHeating          = hs.DeviceValue("48")              ' Thermostat Heating Setpoint
	Dim CurrentTemperature       = hs.DeviceValue("43")              ' Thermostat Current Temperature
	Dim OutsideTemperature       = hs.DeviceValue("7")               ' WeatherXML Home Temperature
	
	' Determine if the HVAC is currently running
	If hs.DeviceValue("47") <> 0 Then
		OperatingState = "On"
	Else
		OperatingState = "Off"
	End If

	' Write to the file
	Using sw As StreamWriter = File.AppendText(strFile)
	  sw.WriteLine( CurrentDate &"," & OperatingState & "," & SetPointCooling & "," & SetPointHeating & "," & CurrentTemperature & "," & OutsideTemperature )
	End Using
  
End Sub
