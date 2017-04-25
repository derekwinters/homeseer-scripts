' ==============================================================================
' Home Maintenance Reminder
' ==============================================================================
' Send a variety of pre-configured home maintenance reminders.
' ==============================================================================
Sub Main(Param as Object)
	' ==========================================================================
	' Create Variables
	' ==========================================================================
	Dim AmyCell   As String = "7637426359@tmomail.net"
	Dim DerekCell As String = "7632184538@mms.att.net"
	Dim SendFrom  As String = "home@derekwinters.com"
	Dim Message   As String = ""
	
	' ==========================================================================
	' Build Message
	' ==========================================================================
	If ( Param = "furnace filter" ) Then
		Message = "<br />Change the furnace filter today."
	End If
	
	' ==========================================================================
	' Send message
	' ==========================================================================
	If ( Message = "" ) Then
		hs.WriteLog("Maintenance Reminder","Maintenance reminder failed, invalid parameter passed to function.")
	Else
		hs.SendEmail(DerekCell,SendFrom ,"", "", "Maintenance Reminder", Message, "")
		hs.SendEmail(AmyCell,  SendFrom ,"", "", "Maintenance Reminder", Message, "")
	End If
	
end Sub