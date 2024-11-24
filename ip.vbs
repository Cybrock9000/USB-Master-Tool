' File: SaveIPAndHostname.vbs

Dim objWMIService, colAdapters, objAdapter, strIP, objShell, objFSO, objFile, objNetwork, strHostname, strOutput

' Connect to WMI to retrieve network adapter configurations
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
Set colAdapters = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

' Retrieve hostname
Set objNetwork = CreateObject("WScript.Network")
strHostname = objNetwork.ComputerName

strIP = ""

' Loop through adapters and get the first IP address
For Each objAdapter In colAdapters
    If Not IsNull(objAdapter.IPAddress) Then
        strIP = objAdapter.IPAddress(0) ' Get the first IP address
        Exit For
    End If
Next

' Combine hostname and IP address
If strIP <> "" Then
    strOutput = "Hostname: " & strHostname & " | IP Address: " & strIP
Else
    strOutput = "Hostname: " & strHostname & " | IP Address: Not Available"
End If

' Write to a text file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("IP_Hostname.txt", 2, True) ' 2 = ForWriting, True = Create if not exists
objFile.WriteLine strOutput
objFile.Close

' Notify user
MsgBox "Saved to IP_Hostname.txt: " & vbCrLf & strOutput, vbInformation, "Info Saved"

' Clean up
Set objWMIService = Nothing
Set colAdapters = Nothing
Set objShell = Nothing
Set objFSO = Nothing
Set objFile = Nothing
Set objNetwork = Nothing
