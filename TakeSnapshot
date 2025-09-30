' TakeSnapshot.vbs
Dim autECLPSObj
Set autECLPSObj = CreateObject("PCOMM.autECLPS")

' Pick connection name (usually "A", "B", etc from PCOMM session title)
autECLPSObj.SetConnectionByName("A")

' Path will be passed as argument from Java
Dim filePath
filePath = Wscript.Arguments.Item(0)

' Take snapshot
autECLPSObj.PrintScreenToFile filePath
