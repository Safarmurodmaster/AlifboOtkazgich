Option Explicit
Dim fso, sh, addinDir, xlstart, ppt, i, p, n, ver
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
On Error Resume Next
addinDir = sh.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\AddIns"
xlstart = sh.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Excel\XLSTART"
n = 0
For Each p In GetObject("winmgmts:").ExecQuery("select * from Win32_Process where Name='EXCEL.EXE'")
  n = n + 1
Next
If n = 0 Then fso.DeleteFile xlstart & "\AlifboOtkazgich.xlam", True
Err.Clear
n = 0
For Each p In GetObject("winmgmts:").ExecQuery("select * from Win32_Process where Name='POWERPNT.EXE'")
  n = n + 1
Next
If n = 0 Then
  ver = sh.RegRead("HKCR\PowerPoint.Application\CurVer\")
  If ver <> "" Then
    ver = Mid(ver, InStrRev(ver, ".") + 1) & ".0"
    sh.Run "reg delete ""HKCU\Software\Microsoft\Office\" & ver & "\PowerPoint\AddIns\AlifboOtkazgich"" /f", 0, True
    fso.DeleteFile addinDir & "\AlifboOtkazgich.ppam", True
  End If
End If
fso.DeleteFile addinDir & "\AlifboOtkazgich.xlam", True
WScript.Quit 0
