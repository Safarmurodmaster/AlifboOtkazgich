Option Explicit
Dim sh, fso, ver, key, oldVal, hadOld, w, d, startup, bas, cls, target, code, dir, rc
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
code = 0
On Error Resume Next
ver = sh.RegRead("HKCR\Word.Application\CurVer\")
If Err.Number <> 0 Or ver = "" Then WScript.Quit 2
ver = Mid(ver, InStrRev(ver, ".") + 1) & ".0"
key = "HKCU\Software\Microsoft\Office\" & ver & "\Word\Security\AccessVBOM"
Err.Clear
oldVal = sh.RegRead(key)
hadOld = (Err.Number = 0)
Err.Clear
sh.RegWrite key, 1, "REG_DWORD"
dir = fso.GetParentFolderName(WScript.ScriptFullName)
bas = dir & "\Alifbo.bas"
cls = dir & "\clsAppEvents.cls"
startup = sh.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Word\STARTUP"
If Not fso.FolderExists(startup) Then fso.CreateFolder startup
target = startup & "\AlifboOtkazgich.dotm"
If fso.FileExists(target) Then fso.DeleteFile target, True
Err.Clear
Set w = CreateObject("Word.Application")
If Err.Number <> 0 Then code = 3
If code = 0 Then
  w.Visible = False
  w.DisplayAlerts = 0
  Set d = w.Documents.Add
  d.VBProject.VBComponents.Import cls
  d.VBProject.VBComponents.Import bas
  d.VBProject.VBComponents.Import dir & "\AlifboCore.bas"
  d.VBProject.VBComponents.Import dir & "\AlifboLic.bas"
  d.VBProject.VBComponents.Import dir & "\AlifboKrill.bas"
  If Err.Number <> 0 Then code = 4
End If
If code = 0 Then
  d.SaveAs target, 15
  If Err.Number <> 0 Then code = 5
  d.Close False
End If
If Not w Is Nothing Then w.Quit 0
Err.Clear
If hadOld Then
  sh.RegWrite key, oldVal, "REG_DWORD"
Else
  sh.RegDelete key
End If

If code = 0 Then
  If fso.FileExists(dir & "\patch-ribbon.ps1") And fso.FileExists(dir & "\ribbon-word.xml") Then
    rc = sh.Run("powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & dir & "\patch-ribbon.ps1"" -Path """ & target & """ -XmlPath """ & dir & "\ribbon-word.xml""", 0, True)
  End If
End If

If code = 0 And Not fso.FileExists(target) Then code = 7
WScript.Quit code
