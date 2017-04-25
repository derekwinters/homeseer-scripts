' ==============================================================================
' Alert State Change
' ==============================================================================
' Send a formatted message indicating the security system state has changed, as
' well as a list of all monitored users and devices.
' ==============================================================================
Sub Main(Param as Object)
	' ==========================================================================
	' Create Variables
	' ==========================================================================
	Dim SecurityMode    As String
	Dim RefSecurityMode As Integer = 28
	Dim RefDerekStatus  As Integer = 22
	Dim RefAmyStatus    As Integer = 24
	Dim Message As String
	
	' ==========================================================================
	' Setup
	' ==========================================================================
	' Log start of script
	hs.WriteLog("Security System", "State change alert triggered")
	
	' Get the security system state
	SecurityMode = hs.DeviceVSP_GetStatus(RefSecurityMode, hs.devicevalue(RefSecurityMode), 1)
	
	' ==========================================================================
	' Send message
	' ==========================================================================
	' Build message body
	Message = "Security System is set to " & SecurityMode
	
	' Speak
	hs.Speak(Message, True, "*")
	
	' Format for MMS
	Message = "<br />" & Message & "<br /><br />" & DateTime.Now.ToString("MMMM dd, HH:mm:ss")
	
	' Send
	hs.SendEmail(hs.DeviceString(168), hs.DeviceString(174),"", "", "Security System State", Message, "")
	hs.SendEmail(hs.DeviceString(167), hs.DeviceString(174),"", "", "Security System State", Message, "")
	
end Sub