Sub Main(parms As Object)

	' ==========================================================================
	' Set Variables
	' ==========================================================================
	Dim FanMode As Integer = 44
	
	' ==========================================================================
	' Toggle
	' ==========================================================================
	If ( hs.DeviceValue(FanMode) = 0 ) Then
	
		MakeTheCapiHappy("Auto", FanMode, 1 )
	
	Else
		
		MakeTheCapiHappy("Auto", FanMode, 0 )
		
	End If
	
End Sub

' ==============================================================================
' Make the Capi Happy
' ==============================================================================
' Set a device using CAPI
' ==============================================================================
Sub MakeTheCapiHappy(CapiControlString As String, DeviceReferenceID As Integer, SetValue As Double)

	' Create device
	Dim CapiControlDevice as HomeSeerAPI.CAPI.CAPIControl = hs.CAPIGetSingleControl(DeviceReferenceID, True, CapiControlString, False, False)
	
	' Set value
	CapiControlDevice.ControlValue = SetValue
	
	' Persist the value
	hs.CAPIControlHandler(CapiControlDevice)

End Sub