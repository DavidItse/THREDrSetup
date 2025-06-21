On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error: Could not create Excel instance. Please ensure Excel is installed."
    WScript.Quit 1
End If
objExcel.Visible = True
objExcel.Workbooks.Add
Set addIn1 = objExcel.AddIns.Add("C:\Users\david\AppData\Roaming\ThredrDB\ThredrDB_add-in-AddIn64-packed.xll")
If Err.Number <> 0 Then
    WScript.Echo "Error: Failed to register the THREDrDB add-in. Ensure the XLL file exists and you have permissions."
End If
addIn1.Installed = True
If Err.Number <> 0 Then
    WScript.Echo "Error: Failed to enable the THREDrDB add-in. Ensure you have permissions to modify Excel settings."
End If
Set addIn2 = objExcel.AddIns.Add("C:\Users\david\AppData\Roaming\ThredrDB\IntelliSense64.xll")
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
objExcel.Quit