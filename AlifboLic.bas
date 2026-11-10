Attribute VB_Name = "AlifboLic"
Option Explicit

Private Const SALT As String = "SAFAR_2026_"
Private Const FREE_BUILD As Boolean = False

Public Function ProdName() As String
  ProdName = "Alifbo" & ChrW(&HD6) & "tkaz" & ChrW(&H11F) & "ich"
End Function

Public Function AuthorText() As String
  AuthorText = "Muallif: Fozliddinov Safarmurod" & vbCrLf & "Telegram: @Safar_master" & vbCrLf & "Telefon: +998939530473"
End Function

Private Function Hash16(ByVal s As String) As Long
  Dim i As Long, h As Long, c As Long
  h = 17
  For i = 1 To Len(s)
    c = Asc(Mid$(s, i, 1))
    h = ((h * 31 + c) And 32767)
  Next i
  Hash16 = h
End Function

Private Function Hex4(ByVal n As Long) As String
  Dim h As String
  h = Hex(n)
  While Len(h) < 4: h = "0" & h: Wend
  Hex4 = h
End Function

Public Function GetMachineCode() As String
  Dim fso As Object, drv As Object, sn As String, h1 As String, h2 As String
  On Error Resume Next
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set drv = fso.GetDrive("C:")
  sn = CStr(Abs(CLng(drv.SerialNumber)))
  If Err.Number <> 0 Or Len(sn) = 0 Then sn = "12345678"
  h1 = Hex4(Hash16("M1" & sn))
  h2 = Hex4(Hash16("M2" & sn))
  GetMachineCode = "AF-" & h1 & "-" & h2
End Function

Public Function CalcKey(ByVal machCode As String) As String
  Dim clean As String, s As String, k1 As String, k2 As String
  clean = Replace(Replace(UCase$(machCode), "-", ""), " ", "")
  s = SALT & clean
  k1 = Hex4(Hash16("K1" & s))
  k2 = Hex4(Hash16("K2" & s))
  CalcKey = "AK-" & k1 & "-" & k2
End Function

Private Function RegPath() As String
  RegPath = "HKCU\Software\Microsoft\AlifboSetup\lic_key"
End Function

Private Function DatPath() As String
  DatPath = Environ("APPDATA") & "\Alifbo\lic.dat"
End Function

Private Function GetSavedKey() As String
  Dim k As String, ff As Integer, ln As String
  On Error Resume Next
  k = CreateObject("WScript.Shell").RegRead(RegPath())
  If Len(Trim$(k)) = 0 Then
    ff = FreeFile
    Open DatPath() For Input As #ff
    If Err.Number = 0 Then
      Line Input #ff, ln
      k = Trim$(ln)
      Close #ff
    End If
  End If
  GetSavedKey = Trim$(k)
End Function

Private Sub SaveKey(ByVal k As String)
  Dim ff As Integer
  On Error Resume Next
  CreateObject("WScript.Shell").RegWrite RegPath(), k, "REG_SZ"
  MkDir Environ("APPDATA") & "\Alifbo"
  Err.Clear
  ff = FreeFile
  Open DatPath() For Output As #ff
  Print #ff, k
  Close #ff
End Sub

Public Function IsLicensed() As Boolean
  Dim curMc As String, validK As String, savedK As String
  If FREE_BUILD Then
    IsLicensed = True
    Exit Function
  End If
  curMc = GetMachineCode()
  validK = CalcKey(curMc)
  savedK = GetSavedKey()
  IsLicensed = (Len(savedK) > 0 And UCase$(savedK) = UCase$(validK))
End Function

Public Function AskCode() As Boolean
  Dim mc As String, validK As String, userK As String, msg As String
  mc = GetMachineCode()
  validK = CalcKey(mc)
  msg = "Dastur faollashtirilmagan!" & vbCrLf & vbCrLf & _
        "Sizning kompyuter kodingiz: " & mc & vbCrLf & vbCrLf & _
        AuthorText() & vbCrLf & vbCrLf & _
        "Litsenziya kalitini kiriting:"
  userK = InputBox(msg, ProdName() & ": Litsenziya", "")
  userK = Trim$(userK)
  If Len(userK) = 0 Then Exit Function
  If UCase$(userK) = UCase$(validK) Then
    SaveKey validK
    MsgBox "Litsenziya muvaffaqiyatli faollashtirildi! Rahmat.", vbInformation, ProdName()
    AskCode = True
  Else
    MsgBox "Litsenziya kaliti noto'g'ri!" & vbCrLf & "Muallifga murojaat qiling: @Safar_master", vbCritical, ProdName()
    AskCode = False
  End If
End Function

Public Function Gate() As Boolean
  If IsLicensed() Then
    Gate = True
  Else
    Gate = AskCode()
  End If
End Function

Public Function StatusSuffix() As String
  StatusSuffix = ""
End Function

Public Function LeftMinutes() As Long
  LeftMinutes = 0
End Function

Public Sub LicenseInfo()
  Dim mc As String
  mc = GetMachineCode()
  If IsLicensed() Then
    MsgBox "Litsenziya faol!" & vbCrLf & "Kompyuter kodi: " & mc & vbCrLf & vbCrLf & AuthorText(), vbInformation, ProdName()
  Else
    If MsgBox("Dastur faollashtirilmagan." & vbCrLf & "Kompyuter kodi: " & mc & vbCrLf & vbCrLf & "Kalitni kiritasizmi?", vbYesNo + vbQuestion, ProdName()) = vbYes Then
      AskCode
    End If
  End If
End Sub

Public Sub AlifboAbout()
  MsgBox ProdName() & vbCrLf & vbCrLf & AuthorText(), vbInformation, "Muallif haqida"
End Sub

Public Sub OpenToolsApp()
  Dim p As String
  p = Environ("LOCALAPPDATA") & "\AlifboOtkazgich\tools\alifbo-tools.ps1"
  If Dir(p) = "" Then
    MsgBox "Vositalar topilmadi. Dasturni qayta o'rnating.", vbExclamation, ProdName()
    Exit Sub
  End If
  Shell "powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & p & """", vbHide
End Sub
