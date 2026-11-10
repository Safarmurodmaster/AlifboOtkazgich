Option Explicit
Dim sh, fso, ver, key, oldVal, hadOld, ppt, pres, dir, addinDir, target, code, ad, i, running, p, logPath, tl, rc, rk
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = sh.ExpandEnvironmentStrings("%APPDATA%") & "\Alifbo\kurulish-log.txt"
On Error Resume Next
If Not fso.FolderExists(sh.ExpandEnvironmentStrings("%APPDATA%") & "\Alifbo") Then fso.CreateFolder sh.ExpandEnvironmentStrings("%APPDATA%") & "\Alifbo"
Err.Clear
Sub Lg(msg)
  Dim f
  Set f = fso.OpenTextFile(logPath, 8, True)
  f.WriteLine Now & "  [PowerPoint] " & msg
  f.Close
End Sub
Sub Chk(stepName)
  If Err.Number <> 0 Then
    Lg stepName & ": XATO " & Err.Number & " - " & Err.Description
    Err.Clear
  Else
    Lg stepName & ": ok"
  End If
End Sub
code = 0
ver = sh.RegRead("HKCR\PowerPoint.Application\CurVer\")
If Err.Number <> 0 Or ver = "" Then
  Lg "PowerPoint o'rnatilmagan, o'tkazib yuborildi"
  WScript.Quit 2
End If
Err.Clear
running = 0
For Each p In GetObject("winmgmts:").ExecQuery("select * from Win32_Process where Name='POWERPNT.EXE'")
  running = running + 1
Next
If running > 0 Then
  Lg "PowerPoint ochiq turibdi, o'tkazib yuborildi"
  WScript.Quit 7
End If
ver = Mid(ver, InStrRev(ver, ".") + 1) & ".0"
Lg "Office versiyasi: " & ver
key = "HKCU\Software\Microsoft\Office\" & ver & "\PowerPoint\Security\AccessVBOM"
Err.Clear
oldVal = sh.RegRead(key)
hadOld = (Err.Number = 0)
Err.Clear
sh.RegWrite key, 1, "REG_DWORD"
Chk "AccessVBOM yoqildi"
dir = fso.GetParentFolderName(WScript.ScriptFullName)
addinDir = sh.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\AddIns"
If Not fso.FolderExists(addinDir) Then fso.CreateFolder addinDir
tl = "HKCU\Software\Microsoft\Office\" & ver & "\PowerPoint\Security\Trusted Locations\Location90\"
sh.RegWrite tl & "Path", addinDir & "\", "REG_SZ"
sh.RegWrite tl & "AllowSubfolders", 1, "REG_DWORD"
sh.RegWrite tl & "Description", "AlifboOtkazgich", "REG_SZ"
Chk "Ishonchli joy qo'shildi"
target = addinDir & "\AlifboOtkazgich.ppam"
Set ppt = CreateObject("PowerPoint.Application")
If Err.Number <> 0 Then
  code = 3
  Chk "PowerPoint ishga tushmadi"
End If
If code = 0 Then
  For i = ppt.AddIns.Count To 1 Step -1
    If InStr(LCase(ppt.AddIns(i).Name), "alifbo") > 0 Then
      ppt.AddIns(i).Loaded = False
      ppt.AddIns.Remove i
    End If
  Next
  Err.Clear
  Set pres = ppt.Presentations.Add(0)
  Chk "Yangi taqdimot"
  pres.VBProject.VBComponents.Import dir & "\AlifboCore.bas"
  Chk "AlifboCore import"
  pres.VBProject.VBComponents.Import dir & "\AlifboLic.bas"
  Chk "AlifboLic import"
  pres.VBProject.VBComponents.Import dir & "\AlifboKrill.bas"
  Chk "AlifboKrill import"
  pres.VBProject.VBComponents.Import dir & "\AlifboPpt.bas"
  Chk "AlifboPpt import"
  If pres.VBProject.VBComponents.Count < 4 Then code = 4
End If
If code = 0 Then
  If fso.FileExists(target) Then fso.DeleteFile target, True
  Err.Clear
  pres.SaveAs target, 30
  If Err.Number <> 0 Then code = 5
  Chk "Saqlash (ppam)"
  pres.Close
End If
If code <> 3 Then ppt.Quit
Err.Clear
If hadOld Then
  sh.RegWrite key, oldVal, "REG_DWORD"
Else
  sh.RegDelete key
End If
Err.Clear
If code = 0 Then
  rc = sh.Run("powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & dir & "\patch-ribbon.ps1"" -Path """ & target & """ -XmlPath """ & dir & "\ribbon-ppt.xml""", 0, True)
  Lg "Alifbo lentasi qo'shildi, kod: " & rc
  If rc <> 0 Then code = 6
End If
If code = 0 Then
  rk = "HKCU\Software\Microsoft\Office\" & ver & "\PowerPoint\AddIns\AlifboOtkazgich\"
  sh.RegWrite rk & "Path", target, "REG_SZ"
  sh.RegWrite rk & "AutoLoad", 1, "REG_DWORD"
  Chk "Ro'yxatga olindi (AutoLoad)"
  If Err.Number <> 0 Then code = 7
  Err.Clear
End If
Lg "Yakuniy kod: " & code
WScript.Quit code
