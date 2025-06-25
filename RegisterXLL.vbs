On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error: Could not create Excel instance. Please ensure Excel is installed."
    WScript.Quit 1
End If
objExcel.Visible = True
objExcel.Workbooks.Add
Set addIn1 = objExcel.AddIns.Add(WShell.ExpandEnvironmentStrings("%APPDATA%") & "\ThredrDB\ThredrDB_add-in-AddIn64-packed.xll")
If Err.Number <> 0 Then
    WScript.Echo "Error: Failed to register the THREDrDB add-in. Ensure the XLL file exists and you have permissions."
End If
addIn1.Installed = True
If Err.Number <> 0 Then
    WScript.Echo "Error: Failed to enable the THREDrDB add-in. Ensure you have permissions to modify Excel settings."
End If
Set addIn2 = objExcel.AddIns.Add(WShell.ExpandEnvironmentStrings("%APPDATA%") & "\ThredrDB\IntelliSense64.xll")
If Err.Number <> 0 Then
    WScript.Echo "Error: Failed to register the IntelliSense add-in. Ensure the XLL file exists and you have permissions."
End If
addIn2.Installed = True
If Err.Number <> 0 Then
    WScript.Echo "Error: Failed to enable the IntelliSense add-in. Ensure you have permissions to modify Excel settings."
End If
For Each addIn In objExcel.AddIns
    If InStr(LCase(addIn.FullName), LCase("THREDrDB-packed.xll")) > 0 Then
        If Not addIn.Installed Then
            WScript.Echo "Warning: THREDrDB add-in registered but not enabled. Please enable it manually in Excel."
        End If
    End If
    If InStr(LCase(addIn.FullName), LCase("IntelliSense64.xll")) > 0 Then
        If Not addIn.Installed Then
            WScript.Echo "Warning: IntelliSense add-in registered but not enabled. Please enable it manually in Excel."
        End If
    End If
    Exit For
Next
On Error Resume Next
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.DeleteFile WShell.ExpandEnvironmentStrings("%APPDATA%") & "\ThredrDB\Checksum.txt"
objFSO.DeleteFile WShell.ExpandEnvironmentStrings("%APPDATA%") & "\ThredrDB\ComputedChecksum.txt"
If Err.Number <> 0 Then
    WScript.Echo "Error deleting file: " & Err.Description
    Err.Clear
Else
    WScript.Echo "File deleted successfully."
End If
Set objFSO = Nothing
objExcel.ActiveWorkbook.Saved = True
objExcel.ActiveWorkbook.Close(False)
objExcel.Quit
On Error Resume Next

' Create a WMI service object
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

' Query all running processes where the name is "excel.exe"
Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'excel.exe'")

' Check if any Excel processes were found
If colProcesses.Count = 0 Then
    'WScript.Echo "No Excel processes are running."
Else
    ' Loop through each Excel process and terminate it
    For Each objProcess in colProcesses
        objProcess.Terminate()
        If Err.Number = 0 Then
            'WScript.Echo "Terminated Excel process with PID: " & objProcess.ProcessID
        Else
            'WScript.Echo "Failed to terminate Excel process with PID: " & objProcess.ProcessID & ". Error: " & Err.Description
            Err.Clear
        End If
    Next
End If

' Clean up
Set colProcesses = Nothing
Set objWMIService = Nothing