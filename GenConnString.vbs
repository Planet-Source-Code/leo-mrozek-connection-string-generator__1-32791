Dim oDataLinks, sRetVal
Set oDataLinks = CreateObject("DataLinks")
On Error Resume Next ' Trap Cancel button
sRetVal = oDataLinks.PromptNew
On Error Goto 0
If Not IsEmpty(sRetVal) Then ' Didn't click Cancel
	InputBox "Your Connection String is listed below.", "OLEDB Connection String", sRetVal
End If
Set oDataLinks = Nothing

