Imports System.IO

Sub Main(Param As Object)
  ' ----------------------------------------------------------------------------
  ' Meal List
  ' ----------------------------------------------------------------------------
  ' File Path
  Dim file As String = "C:\Program Files (x86)\HomeSeer HS3\scripts\github\standalone-scripts\MealPlanner\MealList.txt"
  
  ' Create ArrayList
  Dim MealList As New ArrayList

  ' Read File into Array
  FileOpen(1, file, OpenMode.Input)
  Do While Not EOF(1)
    MealList.Add(LineInput(1))
  Loop

  Dim Message As String = "This Week's Meal Plan<br />"

  ' ----------------------------------------------------------------------------
  ' Random Items
  ' ----------------------------------------------------------------------------
  ' Get Number of meals in array
  Dim Length As Integer = MealList.Count
  hs.WriteLog("Meal Planner", Length)
  'Dim TodayDate  = DateTime.Now.ToString("dd"))
  Dim TotalChoices As Integer = 14
  Dim MealPlan As New ArrayList
  Dim MealChoice As String
  Dim MealDate As Integer
  Dim Randomizer As Random = New Random
  For number As Integer = 0 To (TotalChoices-1)
    ' Choose Meal
    MealChoice = MealList( Randomizer.Next(0,Length-1 ) )

    ' Log
    hs.WriteLog("Meal Planner", "Recipe " & MealChoice )
    
    ' Message
    Message = Message & "<br /><br />Recipe " & MealChoice
  Next

  FileClose(1)
  
  SendMessage("Meal Planner", Message)

End Sub



' ==============================================================================
' Send Message
' ==============================================================================
Sub SendMessage(SubjectString As String, MessageString As String)

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
        hs.SendEmail(hs.DeviceString(Device.ref(hs)), hs.DeviceString(174), "", "", SubjectString, MessageString, "")
      Else
        hs.WriteLog("MMS Messaging", "Device is disabled for messaging (ReferenceID: " & Device.ref(hs) & ")")
      End If
    End If

  Loop

End Sub