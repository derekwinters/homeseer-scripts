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
    If (Message = "") Then
        hs.WriteLog("Maintenance Reminder", "Maintenance reminder failed, invalid parameter passed to function.")
    Else
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
                    hs.SendEmail(hs.DeviceString(Device.ref(hs)), hs.DeviceString(174), "", "", "Maintenance Reminder", Message, "")
                Else
                    hs.WriteLog("MMS Messaging", "Device is disabled for messaging (ReferenceID: " & Device.ref(hs) & ")")
                End If
            End If

        Loop

    End If
	
end Sub