Option Explicit

' ===== CONFIGURATION =====
Const RUN_AS_ADMIN = False
Const SILENT_MODE  = False

Dim sh, fso, scriptDir, psExe, ps1, cmd, silentExe, windowStyle

Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Get the folder where THIS script lives
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' ===== POWERSHELL PATH =====
' Try PowerShell 7 first, then fall back to Windows PowerShell
psExe = sh.ExpandEnvironmentStrings("%ProgramFiles%") & "\PowerShell\7\pwsh.exe"
If Not fso.FileExists(psExe) Then
    psExe = sh.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
End If

' ===== SCRIPT PATHS =====
ps1       = fso.BuildPath(scriptDir, "src\ps_script.ps1")
silentExe = fso.BuildPath(scriptDir, "src\SilentCMD\SilentCMD.exe")

' ===== VALIDATE PATHS =====
If Not fso.FileExists(psExe) Then
    WScript.Echo "PowerShell not found: " & psExe
    WScript.Quit 1
End If

If Not fso.FileExists(ps1) Then
    WScript.Echo "Script not found: " & ps1
    WScript.Quit 1
End If

If SILENT_MODE And Not fso.FileExists(silentExe) Then
    WScript.Echo "SilentCMD not found: " & silentExe
    WScript.Quit 1
End If

' ===== BUILD COMMAND =====
If SILENT_MODE Then
    ' SilentCMD wraps the entire PowerShell command
    cmd = """" & silentExe & """ """ & psExe & """ -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """"
    windowStyle = 0  ' Hidden
Else
    cmd = """" & psExe & """ -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """"
    windowStyle = 1  ' Normal
End If

' ===== EXECUTE =====
If RUN_AS_ADMIN Then
    Dim shellApp
    Set shellApp = CreateObject("Shell.Application")
    ' Note: runas with SilentCMD may still show UAC prompt
    If SILENT_MODE Then
        shellApp.ShellExecute silentExe, """" & psExe & """ -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """", "", "runas", 0
    Else
        shellApp.ShellExecute psExe, "-NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """", "", "runas", 1
    End If
Else
    sh.Run cmd, windowStyle, False
End If
