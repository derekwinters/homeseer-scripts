Sub Main(parms As Object)

	' ==========================================================================
	' Set Variables
	' ==========================================================================
	Dim Mode As Integer = 46

    ' ==========================================================================
    ' Toggle
    ' ==========================================================================
    If (hs.DeviceValue(Mode) = 0) Then
        ' OFF to AUTO
        MakeTheCapiHappy("Auto", Mode, 3)
    Else
        ' AUTO to OFF
        MakeTheCapiHappy("Auto", Mode, 0)
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