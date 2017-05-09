' ==============================================================================
' Home Maintenance Reminder
' ==============================================================================
' Send a variety of pre-configured home maintenance reminders.
' ==============================================================================
Sub Main(Param as Object)
    ' ==========================================================================
    ' Create Variables
    ' ==========================================================================
    Dim Message As String = ""

    ' ==========================================================================
    ' Build Message
    ' ==========================================================================
    ' A simple string can be passed via the parameter, or a trigger word can be
    ' passed that has a pre-configured rich-message prepared.
    If (Param = "furnace filter") Then
        Message = "<br />Change the furnace filter today."
    ElseIf (Param = "garbage") Then
        Message = "<br />Garbage comes tomorrow."
    ElseIf (Param <> "") Then
        Message = Param
    End If
	
	' ==========================================================================
	' Send message
	' ==========================================================================
	If ( Message = "" ) Then
		hs.WriteLog("Maintenance Reminder","Maintenance reminder failed, invalid parameter passed to function.")
	Else
        hs.SendEmail(hs.DeviceString(168), hs.DeviceString(174), "", "", "Maintenance Reminder", Message, "")
        hs.SendEmail(hs.DeviceString(167), hs.DeviceString(174), "", "", "Maintenance Reminder", Message, "")
    End If
	
end Sub