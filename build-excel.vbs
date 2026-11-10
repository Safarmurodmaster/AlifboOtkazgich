Option Explicit
Dim sh, fso, ver, key, oldVal, hadOld, xl, wb, dir, addinDir, xlstart, target, code, w2, ad, running, p, logPath, tl, rc
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = sh.ExpandEnvironmentStrings("%APPDATA%") & "\Alifbo\kurulish-log.txt"
On Error Resume Next
If Not fso.FolderExists(sh.ExpandEnvironmentStrings("%APPDATA%") & "\Alifbo") Then fso.CreateFolder sh.ExpandEnvironmentStrings("%APPDATA%") & "\Alifbo"
Err.Clear
Sub Lg(msg)
  Dim f
  Set f = fso.OpenTextFile(logPath, 8, True)
  f.WriteLine Now & "  [Excel] " & msg
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
ver = sh.RegRead("HKCR\Excel.Application\CurVer\")
If Err.Number <> 0 Or ver = "" Then
  Lg "Excel o'rnatilmagan, o'tkazib yuborildi"
  WScript.Quit 2
End If
Err.Clear
running = 0
For Each p In GetObject("winmgmts:").ExecQuery("select * from Win32_Process where Name='EXCEL.EXE'")
  running = running + 1
Next
If running > 0 Then
  Lg "Excel ochiq turibdi, o'tkazib yuborildi"
  WScript.Quit 7
End If
ver = Mid(ver, InStrRev(ver, ".") + 1) & ".0"
Lg "Office versiyasi: " & ver
key = "HKCU\Software\Microsoft\Office\" & ver & "\Excel\Security\AccessVBOM"
Err.Clear
oldVal = sh.RegRead(key)
hadOld = (Err.Number = 0)
Err.Clear
sh.RegWrite key, 1, "REG_DWORD"
Chk "AccessVBOM yoqildi"
dir = fso.GetParentFolderName(WScript.ScriptFullName)
addinDir = sh.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\AddIns"
xlstart = sh.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Excel\XLSTART"
If Not fso.FolderExists(addinDir) Then fso.CreateFolder addinDir
If Not fso.FolderExists(xlstart) Then fso.CreateFolder xlstart
tl = "HKCU\Software\Microsoft\Office\" & ver & "\Excel\Security\Trusted Locations\Location90\"
sh.RegWrite tl & "Path", addinDir & "\", "REG_SZ"
sh.RegWrite tl & "AllowSubfolders", 1, "REG_DWORD"
sh.RegWrite tl & "Description", "AlifboOtkazgich", "REG_SZ"
Chk "Ishonchli joy qo'shildi"
target = addinDir & "\AlifboOtkazgich.xlam"
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then
  code = 3
  Chk "Excel ishga tushmadi"
End If
If code = 0 Then
  xl.Visible = False
  xl.DisplayAlerts = False
  For Each ad In xl.AddIns
    If InStr(LCase(ad.Name), "alifbo") > 0 Then ad.Installed = False
  Next
  For Each w2 In xl.Workbooks
    If LCase(w2.Name) = "alifbootkazgich.xlam" Then w2.Close False
  Next
  Err.Clear
  Set wb = xl.Workbooks.Add
  Chk "Yangi kitob"
  wb.VBProject.VBComponents.Import dir & "\AlifboCore.bas"
  Chk "AlifboCore import"
  wb.VBProject.VBComponents.Import dir & "\AlifboLic.bas"
  Chk "AlifboLic import"
  wb.VBProject.VBComponents.Import dir & "\AlifboKrill.bas"
  Chk "AlifboKrill import"
  wb.VBProject.VBComponents.Import dir & "\AlifboXl.bas"
  Chk "AlifboXl import"
  If wb.VBProject.VBComponents.Count < 5 Then code = 4
End If
If code = 0 Then
  If fso.FileExists(target) Then fso.DeleteFile target, True
  Err.Clear
  wb.SaveAs target, 55
  If Err.Number <> 0 Then code = 5
  Chk "Saqlash (xlam)"
  wb.Close False
End If
If code <> 3 Then xl.Quit
Err.Clear
If hadOld Then
  sh.RegWrite key, oldVal, "REG_DWORD"
Else
  sh.RegDelete key
End If
Err.Clear
If code = 0 Then
  rc = sh.Run("powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & dir & "\patch-ribbon.ps1"" -Path """ & target & """ -XmlPath """ & dir & "\ribbon-excel.xml""", 0, True)
  Lg "Alifbo lentasi qo'shildi, kod: " & rc
  If rc <> 0 Then code = 6
End If
If code = 0 Then
  Err.Clear
  fso.CopyFile target, xlstart & "\AlifboOtkazgich.xlam", True
  Chk "XLSTART'ga nusxalash"
  If Not fso.FileExists(xlstart & "\AlifboOtkazgich.xlam") Then code = 7
  If code = 0 Then fso.DeleteFile target, True
  Err.Clear
End If
Lg "Yakuniy kod: " & code
WScript.Quit code
