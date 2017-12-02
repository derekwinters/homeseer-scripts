Sub Main(ByVal strDevRef As String)
	Dim objDev As Object
	Dim intDevRef as Integer
	Dim intCount As Integer

	intDevRef = CInt(strDevRef)	
	objDev = hs.GetDevicebyRef(intDevRef)
	
	If objDev Is Nothing Then
		hs.WriteLog("DevInfo","Invalid device ref: " & strDevRef )
	Else
		hs.WriteLog("DevInfo", "Ref=" & strDevRef & "; Name=" & objDev.Name(hs) & "; Location=" & objDev.Location(hs) & "; Location2=" & objDev.Location2(hs) & "; Address=" & objDev.Address(hs))
	
		intCount = 0
		
		For Each objCAPIControl As CAPIControl In hs.CAPIGetControl(intDevRef)
			hs.WriteLog("DevInfo", ".CCIndex=" & CStr(objCAPIControl.CCIndex) & "; Label=" & CStr(objCAPIControl.Label) & "; ControlType=" & CStr(objCAPIControl.ControlType) & "; ControlValue=" & CStr(objCAPIControl.ControlValue) & "; ControlString=" & objCAPIControl.ControlString)
			intCount = intCount + 1
		Next
		
		If intCount = 0 Then
			hs.WriteLog("DevInfo", "No CAPIControl objects defined for this device")
		End If
	End If
End Sub