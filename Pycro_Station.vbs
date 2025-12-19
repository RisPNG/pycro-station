' Pycro Station launcher (no prompts)
Option Explicit

Dim sh, fso, cwd, ps1, psExe, cmd, sysroot, silentExe

Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Work from the folder where this .vbs lives
cwd = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = cwd
ps1 = cwd & "\ps_script.ps1"
silentExe = cwd & "\src\SilentCMD\SilentCMD.exe"

' Pick the right PowerShell (handles 32-bit wscript on 64-bit Windows)
If fso.FileExists("C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe") Then
  psExe = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
Else
  ' Fallback if running 32-bit WScript on 64-bit OS
  psExe = sh.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
End If

' --- Default: run PowerShell normally (visible window) ---
cmd = """" & psExe & """ -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """"

' --- OPTIONAL: run via SilentCMD (hidden). Uncomment this line to use it. Must have SilentCMD https://github.com/stbrenner/SilentCMD. ---
' cmd = """" & silentExe & """ """ & psExe & """ -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """" & updateArg

' Launch: 1 = show window; use 0 to hide (especially if using SilentCMD)
sh.Run cmd, 1, False