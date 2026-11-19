Option Explicit
Dim sh, fso, dir, logPath, rc
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = sh.ExpandEnvironmentStrings("%APPDATA%") & "\Alifbo\kurulish-log.txt"
On Error Resume Next
If Not fso.FolderExists(sh.ExpandEnvironmentStrings("%APPDATA%") & "\Alifbo") Then fso.CreateFolder sh.ExpandEnvironmentStrings("%APPDATA%") & "\Alifbo"
Err.Clear
dir = fso.GetParentFolderName(WScript.ScriptFullName)
' PowerShell skript orqali ishlatamiz (COM 429 xatosini chetlab o'tish uchun)
Dim psScript
psScript = dir & "\build-ppt.ps1"
If Not fso.FileExists(psScript) Then
  Dim f
  Set f = fso.OpenTextFile(logPath, 8, True)
  f.WriteLine Now & "  [PPT-VBS] build-ppt.ps1 topilmadi: " & psScript
  f.Close
  WScript.Quit 2
End If
rc = sh.Run("powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & psScript & """", 0, True)
WScript.Quit rc
