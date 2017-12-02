Sub Main(parms As Object)

	' ==========================================================================
	' Set Variables
	' ==========================================================================
	Dim State As Integer = 201

    ' ==========================================================================
    ' Toggle
    ' ==========================================================================
    If (hs.DeviceValue(State) = 255) Then
        ' LOCKED to UNLOCKED
        MakeTheCapiHappy("Unlock", State, 0)
    Else
        ' ALL to LOCKED
        MakeTheCapiHappy("Lock", State, 255)
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