' ============================================================================
'  ExcelToolbox - Optional re-registration helper (VBScript, ASCII only)
' ----------------------------------------------------------------------------
'  Normally NOT needed: the installer already writes the registry entries.
'  Use this only if the "Excel Toolbox" ribbon tab does not appear, e.g. when
'  a group policy disabled Trusted Locations or Excel was running at install.
'
'  Start-menu shortcut:  Excel Toolbox -> Re-register add-in
'  It asks Excel itself to install the add-in via the official COM API:
'      Application.AddIns.Add(path, False).Installed = True
'  This does not change any macro-security setting.
' ============================================================================
Option Explicit

Dim fso, shell, scriptDir, addinPath, xl, ai, found, msg

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
addinPath = fso.BuildPath(scriptDir, "ExcelToolbox.xlam")

If Not fso.FileExists(addinPath) Then
    MsgBox "Add-in file not found:" & vbCrLf & addinPath, vbCritical, "Excel Toolbox"
    WScript.Quit 1
End If

On Error Resume Next
Set xl = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    Err.Clear
    Set xl = CreateObject("Excel.Application")
End If
On Error GoTo 0

If xl Is Nothing Then
    MsgBox "Cannot start Microsoft Excel.", vbCritical, "Excel Toolbox"
    WScript.Quit 1
End If

On Error Resume Next
xl.DisplayAlerts = False
found = False
For Each ai In xl.AddIns
    If LCase(Trim(ai.FullName)) = LCase(Trim(addinPath)) Then
        ai.Installed = True
        found = True
        Exit For
    End If
Next

If Not found Then
    Set ai = xl.AddIns.Add(addinPath, False)
    If Err.Number = 0 Then
        ai.Installed = True
        found = True
    End If
End If

If Err.Number <> 0 Then
    msg = "Registration failed: " & Err.Description & vbCrLf & vbCrLf & _
          "Fallback: in Excel go to File > Options > Add-ins > Manage: Excel Add-ins > Go... " & _
          "and Browse to:" & vbCrLf & addinPath
    MsgBox msg, vbExclamation, "Excel Toolbox"
    WScript.Quit 2
End If

If Not xl.Visible Then xl.Quit
On Error GoTo 0

If found Then
    MsgBox "Add-in registered. Please restart Excel." & vbCrLf & addinPath, _
           vbInformation, "Excel Toolbox"
Else
    MsgBox "Could not register the add-in automatically.", vbExclamation, "Excel Toolbox"
    WScript.Quit 3
End If

WScript.Quit 0
