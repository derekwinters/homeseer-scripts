Sub Main(Parms As Object)
  ' ==========================================================================
  ' Variables
  ' ==========================================================================
  ' Find sensors that are "Z-Wave Temperature" and count the total
  Dim Device As Scheduler.Classes.DeviceClass
  Dim Enumerator As Scheduler.Classes.clsDeviceEnumeration

  ' Math
  Dim Sum As Double
  Dim Total As Integer
  Dim SumDownstairs As Double
  Dim SumUpstairs As Double
  Dim TotalDownstairs As Integer
  Dim TotalUpstairs As Integer
  Dim SumBedrooms As Double
  Dim TotalBedrooms As Double

  ' Trackers
  Dim Tracker As String = "118"
  Dim Upstairs As String = "324"
  Dim Downstairs As String = "325"
  Dim Bedrooms As String = "351"

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
      If ( hs.DeviceName(Device.ref(hs)).StartsWith("Upstairs") ) Then
        hs.WriteLog("HVAC Automation", "Found a temperature device (Name: " & hs.DeviceName(Device.ref(hs)) & ", Value: " & hs.DeviceValueEx(Device.ref(hs)) & ")")
        SumUpstairs = SumUpstairs + hs.DeviceValueEx(Device.ref(hs))
        TotalUpstairs = TotalUpstairs + 1
        Sum = Sum + hs.DeviceValueEx(Device.ref(hs))
        Total = Total + 1
      ElseIf ( hs.DeviceName(Device.ref(hs)).StartsWith("Downstairs") ) Then
        hs.WriteLog("HVAC Automation", "Found a temperature device (Name: " & hs.DeviceName(Device.ref(hs)) & ", Value: " & hs.DeviceValueEx(Device.ref(hs)) & ")")
        SumDownstairs = SumDownstairs + hs.DeviceValueEx(Device.ref(hs))
        TotalDownstairs = TotalDownstairs + 1
        Sum = Sum + hs.DeviceValueEx(Device.ref(hs))
        Total = Total + 1
      End If

      If (hs.DeviceName(Device.ref(hs)).Contains("Bedroom")) Then
        SumBedrooms = SumBedrooms + hs.DeviceValueEx(Device.ref(hs))
        TotalBedrooms = TotalBedrooms + 1
      End If
    End If

  Loop

  ' ==========================================================================
  ' Math
  ' ==========================================================================
  Dim Average As Double = Sum / Total
  Dim AverageDownstairs As Double = SumDownstairs / TotalDownstairs
  Dim AverageUpstairs As Double = SumUpstairs / TotalUpstairs
  Dim AverageBedrooms As Double = SumBedrooms / TotalBedrooms

  'Log the calculated average before rounding and setting the value
  hs.WriteLog("HVAC Automation", "Average home temperature is " & Average & " F, Downstairs " & AverageDownstairs & " F, Upstairs " & AverageUpstairs & " F, Bedrooms " & AverageBedrooms & " F" )

  Average = Math.Round(Average,1,MidpointRounding.AwayFromZero)
  AverageDownstairs = Math.Round(AverageDownstairs,1,MidpointRounding.AwayFromZero)
  AverageUpstairs = Math.Round(AverageUpstairs,1,MidpointRounding.AwayFromZero)
  AverageBedrooms = Math.Round(AverageBedrooms,1,MidpointRounding.AwayFromZero)

  hs.SetDeviceValueByRef(Tracker, Average, True)
  hs.SetDeviceValueByRef(Upstairs, AverageUpstairs, True)
  hs.SetDeviceValueByRef(Downstairs, AverageDownstairs, True)
  hs.SetDeviceValueByRef(Bedrooms, AverageBedrooms, True)

End Sub
