' ==============================================================================
' ==============================================================================
'
' Foscam Control Script
'
' ==============================================================================
'
' Author:  Derek Winters
'
' API Documentation
'   http://www.foscam.es/descarga/Foscam-IPCamera-CGI-User-Guide-AllPlatforms-2015.11.06.pdf
'
' ==============================================================================
' ==============================================================================
' Requirements
'
' Parameters
' The script is called with a comma-separated set of parameters, which is the
'   camera name and the command to be passed.
'
' Devices
' A device must exist with a string set to the username of the camera
' A device must exist with a string set to the password of the camera
' A device must exist with a string set to the from address to send messages
' A device must exist with a string set to each email address to send to
'   email to MMS address work for this
'
' ==============================================================================
' ==============================================================================
Sub Main(Parm As Object)

    If (Parm Is Nothing) Then

        hs.WriteLog("Camera Control", "No parameters passed")

    Else

        ' ======================================================================
        ' Variables
        ' ======================================================================
        Dim Parms() As String = Split(Parm.ToString, ",")
        Dim Camera As String = Parms(0)
        Dim Action As String = Parms(1)
        Dim Username As String = hs.DeviceString(169)
        Dim Password As String = hs.DeviceString(171)
        Dim Port As Integer = 88
        Dim URI As String
        Dim FullCmd As String
        Dim FilePath As String = "\media\" & Camera & "\" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".png"

        ' ======================================================================
        ' Build the command string
        ' ======================================================================
        If (Action = "Stop") Then
            FullCmd = "cmd=ptzStopRun"
        ElseIf (Action = "Up") Then
            FullCmd = "cmd=ptzMoveUp"
        ElseIf (Action = "Right") Then
            FullCmd = "cmd=ptzMoveRight"
        ElseIf (Action = "Down") Then
            FullCmd = "cmd=ptzMoveDown"
        ElseIf (Action = "Left") Then
            FullCmd = "cmd=ptzMoveLeft"
        ElseIf (Action = "Picture") Then
            FullCmd = "cmd=snapPicture2"
        Else
            FullCmd = "cmd=ptzGotoPresetPoint&name=" & Action
        End If

        ' ======================================================================
        ' Log the execution
        ' ======================================================================
        hs.WriteLog("Camera Control", Camera & " " & Action)

        ' ======================================================================
        ' Build the full URI
        ' ======================================================================
        URI = "/cgi-bin/CGIProxy.fcgi?" & FullCmd & "&usr=" & Username & "&pwd=" & Password

        ' ======================================================================
        ' Run command
        ' ======================================================================
        If (Action = "Picture") Then
            ' Save time and date for message
            Dim TakenAt As String = DateTime.Now.ToString("HH:mm:ss")
            Dim TakenOn As String = DateTime.Now.ToString("MMMM dd, yyyy")

            ' Take picture
            hs.getURLImageEx("http://" & Camera, URI, FilePath, Port)

            ' Send MMS messages
            hs.SendEmail(hs.DeviceString(168), hs.DeviceString(174), "", "", "Camera Picture", "Taken at " & TakenAt & " on " & TakenOn, FilePath)
            hs.SendEmail(hs.DeviceString(167), hs.DeviceString(174), "", "", "Camera Picture", "Taken at " & TakenAt & " on " & TakenOn, FilePath)
        Else
            hs.GetURL(Camera, URI, False, Port)
        End If

    End If

End Sub