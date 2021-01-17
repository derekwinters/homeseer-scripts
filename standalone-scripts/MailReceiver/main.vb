' ==============================================================================
' Mail Receiver
' ==============================================================================
Sub Main(Param as Object)
  ' Imports
  Imports System.Text.RegularExpressions

  ' Declarations
  dim index
  dim body as string
  dim emailFrom

  ' Get the index of the mail that triggered the event
  index = hs.MailTrigger

  ' Keep checking any existing emails
  While index >= 0
    ' Clean up the sender email address
    emailFrom = Trim(Regex.Replace(hs.MailFrom(index), "[^a-zA-Z0-9 @\.], ""))
    emailFrom = Regex.Replace(emailFrom, "^1", "")

    hs.WriteLog("Mail Receiver", "Processing mail (" & index & ") from (" & emailFrom & ")")

    ' Validate Sender
    dim ValideSenders = New List(Of String)({hs.DeviceString(167), hs.DeviceString(168)})

    if People.Contains(emailFrom)
      ' Read the mail body. No need for any special characters.
      body = ""
      body = Trim(Regex.Replace(hs.MailText(index), "[^a-zA-Z0-9 ]", ""))

      ' Decide what to do with the email
      if Regex.IsMatch(body, "[Tt]ask \d+")
        ' This is a maintenance task action
        hs.WriteLog("Mail Receiver", "Maintenance Task message received")
      else if Regex.IsMatch(body, "^[Mm]aintenance status$")
        ' Request status for all maintenance tasks
        hs.WriteLog("Mail Receiver", "Maintenance status message received")
      else if Regex.IsMatch(body, "^[Ss]tatus$")
        ' General system status request
        hs.WriteLog("Mail Receiver", "System status message received")
      else
        'unknown request
        hs.WriteLog("Mail Receiver", "Unknown message received")
      end if
      ' Delete the email
      hs.MailDelete(index)
    else
      ' Invalid Sender
      hs.WriteLog("Mail Receiver", "Invalid email (index: " & index & ") from (" & emailFrom & ")")
    end if

    ' Loop
    index--

  End While

End Sub
