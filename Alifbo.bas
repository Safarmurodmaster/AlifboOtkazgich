Attribute VB_Name = "Alifbo"
Option Explicit

Private ev As clsAppEvents
Private autoFlag As Boolean
Private busy As Boolean
Private lastPos As Long
Private lastDoc As String
Private excList As Collection
Private allowStamp As Double
Private allowCache As Boolean
Private tickPos As Long
Private evCount As Long
Private tickCount As Long
Private convCount As Long
Private lastErr As String
Private qChanged As Boolean
Private pasteFlag As Boolean
Private pasteCount As Long
Private tickScheduled As Boolean

Private Const SALT As String = "SAFAR_2026_"
Private Const FREE_BUILD As Boolean = False

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

Private Function GetMachineCode() As String
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

Private Function CalcKey(ByVal machCode As String) As String
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

Private Function IsLicensed() As Boolean
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

Private Function AskCode() As Boolean
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

Private Function Gate() As Boolean
  If IsLicensed() Then
    Gate = True
  Else
    Gate = AskCode()
  End If
End Function

Private Function StatusSuffix() As String
  StatusSuffix = ""
End Function

Private Function LeftMinutes() As Long
  LeftMinutes = 0
End Function
Public Sub LicenseInfo()
  Dim m As Long, t As String
  If IsLicensed() Then
    MsgBox "Litsenziya faol. Rahmat!", vbInformation, "Alifbo"
    Exit Sub
  End If
  m = LeftMinutes()
  If m > 0 Then
    t = "Sinov muddati: " & FmtLeft(m) & " qoldi."
  Else
    t = "Sinov muddati tugagan."
  End If
  If MsgBox(t & vbCrLf & "Litsenziya kodini kiritasizmi?" & vbCrLf & vbCrLf & AuthorText(), vbYesNo + vbQuestion, ProdName()) = vbYes Then
    AskCode ""
  End If
End Sub

Private Function Lt() As String
  Lt = "A-Za-z" & ChrW(&HC7) & ChrW(&HD6) & ChrW(&HDC) & ChrW(&HE7) & ChrW(&HF6) & ChrW(&HFC) & ChrW(&H11E) & ChrW(&H11F) & ChrW(&H130) & ChrW(&H131) & ChrW(&H15E) & ChrW(&H15F)
End Function

Private Function UpLt() As String
  UpLt = "A-Z" & ChrW(&HC7) & ChrW(&HD6) & ChrW(&HDC) & ChrW(&H11E) & ChrW(&H130) & ChrW(&H15E)
End Function

Private Function Ap() As String
  Ap = "'" & ChrW(&H2018) & ChrW(&H2019) & ChrW(&H201B) & ChrW(&H2BB) & "`" & ChrW(&HB4) & ChrW(&H2032) & ChrW(&H2B9) & ChrW(&H2C8) & ChrW(&H2CA) & ChrW(-249)
End Function

Private Function ApAll() As String
  ApAll = Ap() & ChrW(&H2BC)
End Function

Private Function uOUp() As String
  uOUp = ChrW(&HD6)
End Function
Private Function uOLo() As String
  uOLo = ChrW(&HF6)
End Function
Private Function uGUp() As String
  uGUp = ChrW(&H11E)
End Function
Private Function uGLo() As String
  uGLo = ChrW(&H11F)
End Function
Private Function uSUp() As String
  uSUp = ChrW(&H15E)
End Function
Private Function uSLo() As String
  uSLo = ChrW(&H15F)
End Function
Private Function uCUp() As String
  uCUp = ChrW(&HC7)
End Function
Private Function uCLo() As String
  uCLo = ChrW(&HE7)
End Function
Private Function uOkina() As String
  uOkina = ChrW(&H2BB)
End Function
Private Function uTutuq() As String
  uTutuq = ChrW(&H2BC)
End Function

Private Function Ph(ByVal n As Long) As String
  Ph = ChrW(167) & ChrW(167) & CStr(n) & ChrW(167) & ChrW(167)
End Function

Private Function ExcPath() As String
  ExcPath = Environ("APPDATA") & "\Alifbo\istisnolar.txt"
End Function

Private Function Fnd() As Find
  If Selection.Start <> Selection.End Then
    Set Fnd = Selection.Find
  Else
    Set Fnd = ActiveDocument.Content.Find
  End If
End Function

Private Sub Rep(ByVal f As String, ByVal r As String, ByVal wild As Boolean, Optional ByVal whole As Boolean = False)
  Dim n As Long, ok As Boolean, fd As Find
  For n = 1 To 8
    Set fd = Fnd()
    With fd
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = f
      .Replacement.Text = r
      .Forward = True
      .Wrap = 0
      .Format = False
      .MatchCase = True
      .MatchWholeWord = whole
      .MatchWildcards = wild
      .MatchSoundsLike = False
      .MatchAllWordForms = False
      ok = .Execute(Replace:=2)
    End With
    If Not ok Then Exit Sub
    If Not wild Then Exit Sub
  Next n
End Sub

Private Function LoadExc() As Collection
  Dim c As New Collection, ln As String, ff As Integer
  If Dir(ExcPath()) <> "" Then
    ff = FreeFile
    Open ExcPath() For Input As #ff
    Do While Not EOF(ff)
      Line Input #ff, ln
      ln = Trim$(ln)
      If Len(ln) > 0 Then
        If Left$(ln, 1) <> "#" Then c.Add ln
      End If
    Loop
    Close #ff
  End If
  Set LoadExc = c
End Function

Private Function Protect(ByVal c As Collection, ByRef orig() As String) As Long
  Dim i As Long, j As Long, cnt As Long, v(1 To 3) As String
  ReDim orig(1 To c.Count * 3 + 1)
  For i = 1 To c.Count
    v(1) = LCase$(c(i))
    v(2) = UCase$(Left$(v(1), 1)) & Mid$(v(1), 2)
    v(3) = UCase$(v(1))
    For j = 1 To 3
      cnt = cnt + 1
      orig(cnt) = v(j)
      Rep v(j), Ph(cnt), False, True
    Next j
  Next i
  Protect = cnt
End Function

Private Sub Unprotect(ByVal cnt As Long, ByRef orig() As String)
  Dim k As Long
  For k = cnt To 1 Step -1
    Rep Ph(k), orig(k), False
  Next k
End Sub

Private Sub SafeUndoStart(ByVal recName As String)
  On Error Resume Next
  Dim ur As Object
  Set ur = Application.UndoRecord
  If Not ur Is Nothing Then ur.StartCustomRecord recName
  On Error GoTo 0
End Sub

Private Sub SafeUndoEnd()
  On Error Resume Next
  Dim ur As Object
  Set ur = Application.UndoRecord
  If Not ur Is Nothing Then ur.EndCustomRecord
  On Error GoTo 0
End Sub

Private Sub DoConvert(ByVal toNewAlpha As Boolean)
  Dim exc As Collection, orig() As String, cnt As Long
  Dim L As String, A As String, AA As String, U As String, app As Object
  If Not Gate() Then Exit Sub
  On Error GoTo Fail
  Set app = Application
  Application.ScreenUpdating = False
  On Error Resume Next
  SafeUndoStart "Alifbo"
  On Error GoTo Fail
  Set exc = LoadExc()
  cnt = Protect(exc, orig)
  L = Lt(): A = Ap(): AA = ApAll(): U = UpLt()
  If toNewAlpha Then
    Rep "o[" & AA & "]([" & L & "])", uOLo() & "\1", True
    Rep "O[" & AA & "]([" & L & "])", uOUp() & "\1", True
    Rep "g[" & AA & "]([" & L & "])", uGLo() & "\1", True
    Rep "G[" & AA & "]([" & L & "])", uGUp() & "\1", True
    Rep "o[" & AA & "]>", uOLo(), True
    Rep "O[" & AA & "]>", uOUp(), True
    Rep "g[" & AA & "]>", uGLo(), True
    Rep "G[" & AA & "]>", uGUp(), True
    Rep "([" & L & "])[" & A & "]([" & L & "])", "\1" & uTutuq() & "\2", True
    Rep "Sh", uSUp(), False
    Rep "SH", uSUp(), False
    Rep "sh", uSLo(), False
    Rep "sH", uSLo(), False
    Rep "Ch", uCUp(), False
    Rep "CH", uCUp(), False
    Rep "ch", uCLo(), False
    Rep "cH", uCLo(), False
  Else
    Rep uSUp() & "([" & U & "])", "SH\1", True
    Rep "([" & U & "])" & uSUp(), "\1SH", True
    Rep uCUp() & "([" & U & "])", "CH\1", True
    Rep "([" & U & "])" & uCUp(), "\1CH", True
    Rep uSUp(), "Sh", False
    Rep uSLo(), "sh", False
    Rep uCUp(), "Ch", False
    Rep uCLo(), "ch", False
    Rep uOUp(), "O" & uOkina(), False
    Rep uOLo(), "o" & uOkina(), False
    Rep uGUp(), "G" & uOkina(), False
    Rep uGLo(), "g" & uOkina(), False
  End If
  Unprotect cnt, orig
  QuotesAll
  Application.ScreenUpdating = True
  On Error Resume Next
  SafeUndoEnd
  On Error GoTo 0
  MsgBox "Tayyor." & StatusSuffix(), vbInformation, "Alifbo"
  Exit Sub
Fail:
  Application.ScreenUpdating = True
  On Error Resume Next
  SafeUndoEnd
  MsgBox "Xatolik: " & Err.Description, vbExclamation, "Alifbo"
End Sub

Public Sub ToNew()
  DoConvert True
End Sub

Public Sub ToOld()
  DoConvert False
End Sub

Public Sub EditExc()
  Dim p As String, ff As Integer
  p = ExcPath()
  On Error Resume Next
  MkDir Environ("APPDATA") & "\Alifbo"
  On Error GoTo 0
  If Dir(p) = "" Then
    ff = FreeFile
    Open p For Output As #ff
    Print #ff, "# Har qatorda bitta so'z. Bu so'zlar o'zgartirilmaydi."
    Print #ff, "# Apostrofni oddiy ' belgisi bilan yozing."
    Close #ff
  End If
  Shell "notepad.exe """ & p & """", vbNormalFocus
End Sub

Private Function IsApos(ByVal c As String) As Boolean
  Dim a As Long
  If Len(c) <> 1 Then Exit Function
  a = AscW(c)
  If a < 0 Then a = a + 65536
  Select Case a
    Case 39, 96, &HB4, &H2018, &H2019, &H201B, &H2B9, &H2BB, &H2BC, &H2C8, &H2CA, &H2032, &HFF07&
      IsApos = True
  End Select
End Function

Private Function IsLetter(ByVal c As String) As Boolean
  Dim a As Long
  If Len(c) <> 1 Then Exit Function
  a = AscW(c)
  If a < 0 Then a = a + 65536
  Select Case a
    Case 65 To 90, 97 To 122, &HC0 To &H24F
      IsLetter = True
  End Select
End Function

Private Function IsWordCh(ByVal c As String) As Boolean
  IsWordCh = IsLetter(c) Or IsApos(c)
End Function

Public Function ConvWord(ByVal w As String) As String
  Dim i As Long, n As Long, c As String, c1 As String, c2 As String, o As String
  n = Len(w)
  i = 1
  Do While i <= n
    c = Mid$(w, i, 1)
    c1 = Mid$(w, i + 1, 1)
    c2 = Mid$(w, i + 2, 1)
    If (c = "o" Or c = "O" Or c = "g" Or c = "G") And IsApos(c1) And (IsLetter(c2) Or c2 = "") Then
      Select Case c
        Case "o": o = o & uOLo()
        Case "O": o = o & uOUp()
        Case "g": o = o & uGLo()
        Case "G": o = o & uGUp()
      End Select
      i = i + 2
    ElseIf IsApos(c) Then
      If i > 1 And i < n Then
        If IsLetter(Mid$(w, i - 1, 1)) And IsLetter(c1) Then
          o = o & uTutuq()
          i = i + 1
        Else
          o = o & c
          i = i + 1
        End If
      Else
        o = o & c
        i = i + 1
      End If
    ElseIf (c = "s" Or c = "S") And (c1 = "h" Or c1 = "H") Then
      If c = "S" Then o = o & uSUp() Else o = o & uSLo()
      i = i + 2
    ElseIf (c = "c" Or c = "C") And (c1 = "h" Or c1 = "H") Then
      If c = "C" Then o = o & uCUp() Else o = o & uCLo()
      i = i + 2
    Else
      o = o & c
      i = i + 1
    End If
  Loop
  ConvWord = o
End Function

Private Function InExc(ByVal w As String) As Boolean
  Dim i As Long
  If excList Is Nothing Then Exit Function
  For i = 1 To excList.Count
    If LCase$(excList(i)) = LCase$(w) Then
      InExc = True
      Exit Function
    End If
  Next i
End Function

Private Function Allowed() As Boolean
  If allowStamp = 0 Or Abs(Timer - allowStamp) > 60 Then
    allowCache = IsLicensed()
    If Not allowCache Then allowCache = (LeftMinutes() > 0)
    allowStamp = Timer + 0.001
    If Not allowCache Then Application.StatusBar = "Alifbo: sinov muddati tugadi. 'Litsenziya' tugmasini bosing."
  End If
  Allowed = allowCache
End Function

Private Function ProdName() As String
  ProdName = "Alifbo" & ChrW(&HD6) & "tkaz" & ChrW(&H11F) & "ich"
End Function

Private Function AuthorText() As String
  AuthorText = "Muallif: Fozliddinov Safarmurod" & vbCrLf & "Telegram: @Safar_master" & vbCrLf & "Telefon: +998939530473"
End Function

Private Function SafeName(ByVal s As String) As String
  Dim i As Long, c As String, o As String
  s = Replace(s, vbCr, " ")
  s = Replace(s, vbLf, " ")
  s = Replace(s, Chr(7), " ")
  s = Replace(s, Chr(11), " ")
  For i = 1 To Len(s)
    c = Mid$(s, i, 1)
    If InStr("\/:*?""<>|", c) > 0 Then c = " "
    o = o & c
  Next i
  o = Trim$(o)
  Do While InStr(o, "  ") > 0
    o = Replace(o, "  ", " ")
  Loop
  If Len(o) > 50 Then o = Trim$(Left$(o, 50))
  SafeName = o
End Function

Private Function FirstLine(ByVal r As Range) As String
  Dim p As Paragraph, t As String
  For Each p In r.Paragraphs
    t = SafeName(p.Range.Text)
    If Len(t) > 0 Then
      FirstLine = t
      Exit Function
    End If
  Next p
End Function

Private Sub AddPiece(ByRef sP() As Long, ByRef eP() As Long, ByRef pa() As Long, ByRef pb() As Long, ByRef tt() As String, ByRef cnt As Long, ByVal s As Long, ByVal e As Long, ByVal a As Long, ByVal b As Long, ByVal t As String)
  cnt = cnt + 1
  ReDim Preserve sP(1 To cnt)
  ReDim Preserve eP(1 To cnt)
  ReDim Preserve pa(1 To cnt)
  ReDim Preserve pb(1 To cnt)
  ReDim Preserve tt(1 To cnt)
  sP(cnt) = s
  eP(cnt) = e
  pa(cnt) = a
  pb(cnt) = b
  tt(cnt) = t
End Sub

Private Sub PagePiece(ByVal doc As Document, ByVal total As Long, ByRef sP() As Long, ByRef eP() As Long, ByRef pa() As Long, ByRef pb() As Long, ByRef tt() As String, ByRef cnt As Long, ByVal a As Long, ByVal b As Long)
  Dim s As Long, e As Long
  If a < 1 Then a = 1
  If b > total Then b = total
  If b < a Then Exit Sub
  s = doc.GoTo(1, 1, a).Start
  If b < total Then e = doc.GoTo(1, 1, b + 1).Start Else e = doc.Content.End
  AddPiece sP, eP, pa, pb, tt, cnt, s, e, a, b, ""
End Sub

Public Sub SplitDoc()
  Dim doc As Document, total As Long, mode As String, n As Long, fmt As Long, useName As Boolean
  Dim folder As String, base As String, dot As Long, doPdf As Boolean, doDocx As Boolean
  Dim sP() As Long, eP() As Long, pa() As Long, pb() As Long, tt() As String, cnt As Long
  Dim i As Long, j As Long, ok As Long, failed As String, fn As String, r As Range, nd As Document
  Dim parts() As String, pr As String, d As Long, x As Long, y As Long, par As Paragraph, lastS As Long, lastT As String
  Dim ans As VbMsgBoxResult, fd As Object, okPdf As Boolean, okDocx As Boolean
  If Not Gate() Then Exit Sub
  If Documents.Count = 0 Then Exit Sub
  Set doc = ActiveDocument
  If Len(doc.Path) = 0 Then
    MsgBox "Avval hujjatni saqlang (fayl nomi va papka kerak).", vbExclamation, ProdName()
    Exit Sub
  End If
  doc.Repaginate
  total = doc.ComputeStatistics(2)
  If total < 2 Then
    MsgBox "Hujjatda bitta varaq bor, bo'lish shart emas.", vbInformation, ProdName()
    Exit Sub
  End If
  mode = InputBox("Hujjatda " & total & " ta varaq bor. Qanday bo'lamiz?" & vbCrLf & vbCrLf & "1 = har bir varaq alohida fayl" & vbCrLf & "2 = har N varaq bitta faylga" & vbCrLf & "3 = sarlavhalar bo'yicha (1-daraja sarlavha)" & vbCrLf & "4 = oraliqlar bo'yicha (masalan: 1-5, 6-12, 13-20)", ProdName() & ": varaqlarni bo'lish", "1")
  mode = Trim$(mode)
  If Len(mode) = 0 Then Exit Sub
  On Error GoTo SplitFail
  Select Case Left$(mode, 1)
    Case "1"
      For i = 1 To total
        PagePiece doc, total, sP, eP, pa, pb, tt, cnt, i, i
      Next i
    Case "2"
      pr = InputBox("Har nechta varaq bitta faylga tushsin?", ProdName(), "5")
      n = CLng(Val(pr))
      If n < 1 Then Exit Sub
      For i = 1 To total Step n
        PagePiece doc, total, sP, eP, pa, pb, tt, cnt, i, i + n - 1
      Next i
    Case "3"
      lastS = -1
      For Each par In doc.Paragraphs
        If par.OutlineLevel = 1 Then
          x = par.Range.Start
          If lastS >= 0 Then AddPiece sP, eP, pa, pb, tt, cnt, lastS, x, 0, 0, lastT
          If lastS < 0 And x > 0 Then AddPiece sP, eP, pa, pb, tt, cnt, 0, x, 0, 0, "Kirish"
          lastS = x
          lastT = SafeName(par.Range.Text)
        End If
      Next par
      If lastS < 0 Then
        MsgBox "Hujjatda 1-daraja sarlavhalar topilmadi (Heading 1 uslubi kerak).", vbExclamation, ProdName()
        Exit Sub
      End If
      AddPiece sP, eP, pa, pb, tt, cnt, lastS, doc.Content.End, 0, 0, lastT
      For i = 1 To cnt
        pa(i) = doc.Range(sP(i), sP(i)).Information(3)
        pb(i) = doc.Range(eP(i) - 1, eP(i) - 1).Information(3)
      Next i
    Case "4"
      pr = InputBox("Oraliqlarni vergul bilan yozing. Masalan: 1-5, 6-12, 13-20", ProdName(), "1-" & total)
      parts = Split(pr, ",")
      For i = LBound(parts) To UBound(parts)
        d = InStr(parts(i), "-")
        If d > 0 Then
          x = CLng(Val(Left$(parts(i), d - 1)))
          y = CLng(Val(Mid$(parts(i), d + 1)))
        Else
          x = CLng(Val(parts(i)))
          y = x
        End If
        If x >= 1 And y >= x Then PagePiece doc, total, sP, eP, pa, pb, tt, cnt, x, y
      Next i
    Case Else
      MsgBox "1, 2, 3 yoki 4 ni tanlang.", vbExclamation, ProdName()
      Exit Sub
  End Select
  If cnt = 0 Then
    MsgBox "Bo'linadigan qism topilmadi.", vbExclamation, ProdName()
    Exit Sub
  End If
  ans = MsgBox("Qaysi formatda saqlansin?" & vbCrLf & vbCrLf & "Ha = PDF" & vbCrLf & "Yo'q = Word (.docx)" & vbCrLf & "Bekor qilish = ikkalasi ham", vbYesNoCancel + vbQuestion, ProdName())
  doPdf = (ans = vbYes Or ans = vbCancel)
  doDocx = (ans = vbNo Or ans = vbCancel)
  useName = (MsgBox("Fayl nomiga birinchi qatordagi matn (sarlavha) qo'shilsinmi?" & vbCrLf & "Masalan: Hujjat_001_Kirish.pdf", vbYesNo + vbQuestion, ProdName()) = vbYes)
  base = doc.Name
  dot = InStrRev(base, ".")
  If dot > 1 Then base = Left$(base, dot - 1)
  folder = doc.Path & "\" & base & "_bolingan"
  ans = MsgBox("Saqlash papkasini o'zingiz tanlaysizmi?" & vbCrLf & "Ha = tanlayman" & vbCrLf & "Yo'q = hujjat yonidagi " & base & "_bolingan papkasi", vbYesNo + vbQuestion, ProdName())
  If ans = vbYes Then
    On Error Resume Next
    Set fd = Application.FileDialog(4)
    fd.Title = "Saqlash papkasini tanlang"
    fd.InitialFileName = doc.Path & "\"
    If fd.Show = -1 Then folder = fd.SelectedItems(1)
    Err.Clear
    On Error GoTo SplitFail
  End If
  If Dir(folder, vbDirectory) = "" Then MkDir folder
  Application.ScreenUpdating = False
  For i = 1 To cnt
    fn = folder & "\" & base & "_" & Format(i, "000")
    If useName Then
      If Len(tt(i)) = 0 Then
        Set r = doc.Range(sP(i), eP(i))
        tt(i) = FirstLine(r)
      End If
      If Len(tt(i)) > 0 Then fn = fn & "_" & tt(i)
    End If
    Application.StatusBar = "Saqlanmoqda: " & i & " / " & cnt
    DoEvents
    okPdf = True
    okDocx = True
    If doPdf Then
      On Error Resume Next
      Err.Clear
      doc.ExportAsFixedFormat OutputFileName:=fn & ".pdf", ExportFormat:=17, OpenAfterExport:=False, Range:=3, From:=pa(i), To:=pb(i)
      If Err.Number <> 0 Then
        okPdf = False
        failed = failed & vbCrLf & i & "-qism (PDF): " & Err.Description
      End If
      Err.Clear
      On Error GoTo SplitFail
    End If
    If doDocx Then
      On Error Resume Next
      Err.Clear
      Set r = doc.Range(sP(i), eP(i))
      Set nd = Documents.Add
      With nd.PageSetup
        .Orientation = doc.PageSetup.Orientation
        .PageWidth = doc.PageSetup.PageWidth
        .PageHeight = doc.PageSetup.PageHeight
        .TopMargin = doc.PageSetup.TopMargin
        .BottomMargin = doc.PageSetup.BottomMargin
        .LeftMargin = doc.PageSetup.LeftMargin
        .RightMargin = doc.PageSetup.RightMargin
      End With
      Err.Clear
      nd.Content.FormattedText = r.FormattedText
      If Err.Number <> 0 Then
        okDocx = False
        failed = failed & vbCrLf & i & "-qism (Word, nusxalash): " & Err.Description
        Err.Clear
      End If
      nd.SaveAs FileName:=fn & ".docx", FileFormat:=16
      If Err.Number <> 0 Then
        okDocx = False
        failed = failed & vbCrLf & i & "-qism (Word, saqlash): " & Err.Description
      End If
      Err.Clear
      nd.Close False
      Err.Clear
      On Error GoTo SplitFail
    End If
    If okPdf And okDocx Then ok = ok + 1
  Next i
  Application.StatusBar = False
  Application.ScreenUpdating = True
  If Len(failed) > 0 Then
    MsgBox ok & " / " & cnt & " qism saqlandi." & vbCrLf & "Papka: " & folder & vbCrLf & vbCrLf & "Xatolar:" & failed, vbExclamation, ProdName()
  Else
    MsgBox cnt & " ta qism saqlandi:" & vbCrLf & folder, vbInformation, ProdName()
  End If
  On Error Resume Next
  Shell "explorer.exe """ & folder & """", vbNormalFocus
  Exit Sub
SplitFail:
  Application.StatusBar = False
  Application.ScreenUpdating = True
  MsgBox "Xatolik: " & Err.Description & vbCrLf & "Saqlangan qismlar: " & ok, vbExclamation, ProdName()
End Sub

Public Sub OpenTools()
  Dim p As String
  p = Environ("LOCALAPPDATA") & "\AlifboOtkazgich\tools\alifbo-tools.ps1"
  If Dir(p) = "" Then
    MsgBox "Vositalar topilmadi. Dasturni qayta o'rnating.", vbExclamation, ProdName()
    Exit Sub
  End If
  Shell "powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & p & """", vbHide
End Sub

Public Sub SplitTool()
  Dim p As String, dp As String
  If Not Gate() Then Exit Sub
  p = Environ("LOCALAPPDATA") & "\AlifboOtkazgich\tools\alifbo-bolish.ps1"
  If Dir(p) = "" Then
    MsgBox "Vositalar topilmadi. Dasturni qayta o'rnating.", vbExclamation, ProdName()
    Exit Sub
  End If
  dp = ""
  If Documents.Count > 0 Then
    If Len(ActiveDocument.Path) = 0 Then
      MsgBox "Avval hujjatni saqlang (fayl nomi kerak), keyin qaytadan bosing.", vbExclamation, ProdName()
      Exit Sub
    End If
    On Error Resume Next
    If Not ActiveDocument.Saved Then ActiveDocument.Save
    On Error GoTo 0
    dp = ActiveDocument.FullName
  End If
  Shell "powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & p & """ -DocPath """ & dp & """", vbHide
End Sub

Public Sub About()
  MsgBox ProdName() & vbCrLf & vbCrLf & AuthorText(), vbInformation, "Muallif haqida"
End Sub

Private Function QOpen() As String
  QOpen = ChrW(&H201C)
End Function

Private Function QClose() As String
  QClose = ChrW(&H201D)
End Function

Private Function IsSp(ByVal c As String) As Boolean
  IsSp = (c = " " Or c = ChrW(160))
End Function

Private Function IsOpenCtx(ByVal p As String) As Boolean
  Select Case p
    Case "", " ", Chr(9), Chr(13), Chr(11), "(", "[", "{", "-", "/", ChrW(160), ChrW(&H2013), ChrW(&H2014), ChrW(&H201C)
      IsOpenCtx = True
  End Select
End Function

Private Function NeedQuotes(ByVal t As String) As Boolean
  If InStr(t, Chr(34)) > 0 Or InStr(t, "<<") > 0 Or InStr(t, ">>") > 0 Then
    NeedQuotes = True
    Exit Function
  End If
  If InStr(t, ChrW(171)) > 0 Or InStr(t, ChrW(187)) > 0 Or InStr(t, ChrW(&H201E)) > 0 Or InStr(t, ChrW(&H201F)) > 0 Then
    NeedQuotes = True
    Exit Function
  End If
  If InStr(t, ChrW(&H201C)) > 0 Or InStr(t, ChrW(&H201D)) > 0 Then
    NeedQuotes = True
    Exit Function
  End If
  If InStr(t, ChrW(&H2018) & ChrW(&H2018)) > 0 Or InStr(t, ChrW(&H2019) & ChrW(&H2019)) > 0 Then NeedQuotes = True
End Function

Private Function QText(ByVal pr As Range) As String
  Dim t As String, n As Long
  t = pr.Text
  n = pr.End - pr.Start
  If Len(t) > n Then t = Left$(t, n)
  QText = t
End Function

Private Function SetCh(ByVal doc As Document, ByVal s0 As Long, ByVal e0 As Long, ByVal oldS As String, ByVal newS As String) As Boolean
  Dim r As Range
  On Error Resume Next
  Set r = doc.Range(s0, e0)
  If Err.Number = 0 Then
    If r.Text = oldS Then
      r.Text = newS
      If Err.Number = 0 Then
        SetCh = True
        qChanged = True
      Else
        lastErr = Err.Description
      End If
    End If
  Else
    lastErr = Err.Description
  End If
  Err.Clear
End Function

Private Function DelRange(ByVal doc As Document, ByVal s0 As Long, ByVal e0 As Long, ByVal oldS As String) As Boolean
  Dim r As Range
  On Error Resume Next
  Set r = doc.Range(s0, e0)
  If Err.Number = 0 Then
    If r.Text = oldS Then
      r.Delete
      If Err.Number = 0 Then
        DelRange = True
        qChanged = True
      Else
        lastErr = Err.Description
      End If
    End If
  End If
  Err.Clear
End Function

Private Sub AdjCp(ByRef cp As Long, ByVal s0 As Long, ByVal e0 As Long)
  If cp < 0 Then Exit Sub
  If e0 <= cp Then
    cp = cp - (e0 - s0)
  ElseIf cp > s0 Then
    cp = s0
  End If
End Sub

Private Sub QPass0(ByVal doc As Document, ByVal base As Long, ByVal t As String)
  Dim i As Long, p As Long, q As Long
  For i = Len(t) To 1 Step -1
    If Mid$(t, i, 1) = ChrW(&H201E) Then
      p = InStr(i + 1, t, QOpen())
      q = InStr(i + 1, t, ChrW(&H201E))
      If p > 0 And (q = 0 Or q > p) Then
        If SetCh(doc, base + p - 1, base + p, QOpen(), QClose()) Then SetCh doc, base + i - 1, base + i, ChrW(&H201E), QOpen()
      End If
    End If
  Next i
End Sub

Private Sub QPass1(ByVal doc As Document, ByVal base As Long, ByVal t As String, ByRef cp As Long)
  Dim i As Long, w As Long, c As String, c2 As String, nw As String, s0 As Long, e0 As Long
  i = Len(t)
  Do While i >= 1
    c = Mid$(t, i, 1)
    If i > 1 Then c2 = Mid$(t, i - 1, 1) Else c2 = ""
    nw = ""
    w = 1
    If c = ">" And c2 = ">" Then
      nw = QClose(): w = 2
    ElseIf c = "<" And c2 = "<" Then
      nw = QOpen(): w = 2
    ElseIf c = ChrW(&H2019) And c2 = ChrW(&H2019) Then
      nw = QClose(): w = 2
    ElseIf c = ChrW(&H2018) And c2 = ChrW(&H2018) Then
      nw = QOpen(): w = 2
    ElseIf c = ChrW(171) Or c = ChrW(&H201E) Or c = ChrW(&H201F) Then
      nw = QOpen()
    ElseIf c = ChrW(187) Then
      nw = QClose()
    ElseIf c = Chr(34) Then
      If IsOpenCtx(c2) Then nw = QOpen() Else nw = QClose()
    End If
    If Len(nw) > 0 Then
      s0 = base + i - w
      e0 = base + i
      If SetCh(doc, s0, e0, Mid$(t, i - w + 1, w), nw) Then
        If cp >= 0 Then
          If e0 <= cp Then
            cp = cp + (1 - w)
          ElseIf cp > s0 Then
            cp = s0 + 1
          End If
        End If
      End If
      i = i - w
    Else
      i = i - 1
    End If
  Loop
End Sub

Private Sub QPass2(ByVal doc As Document, ByVal base As Long, ByVal t As String, ByRef cp As Long)
  Dim i As Long, j As Long, s0 As Long, e0 As Long, c As String
  i = Len(t)
  Do While i >= 1
    c = Mid$(t, i, 1)
    If c = QClose() Then
      j = i - 1
      Do While j >= 1
        If Not IsSp(Mid$(t, j, 1)) Then Exit Do
        j = j - 1
      Loop
      If j < i - 1 Then
        s0 = base + j
        e0 = base + i - 1
        If DelRange(doc, s0, e0, Mid$(t, j + 1, i - 1 - j)) Then AdjCp cp, s0, e0
      End If
    ElseIf c = QOpen() Then
      j = i + 1
      Do While j <= Len(t)
        If Not IsSp(Mid$(t, j, 1)) Then Exit Do
        j = j + 1
      Loop
      If j > i + 1 Then
        s0 = base + i
        e0 = base + j - 1
        If DelRange(doc, s0, e0, Mid$(t, i + 1, j - i - 1)) Then AdjCp cp, s0, e0
      End If
    End If
    i = i - 1
  Loop
End Sub

Private Function QuotesInRange(ByVal doc As Document, ByVal pr As Range, ByRef cp As Long) As Boolean
  Dim t As String, base As Long
  qChanged = False
  base = pr.Start
  t = QText(pr)
  If Not NeedQuotes(t) Then Exit Function
  QPass0 doc, base, t
  Set pr = doc.Range(base, pr.End)
  t = QText(pr)
  QPass1 doc, base, t, cp
  Set pr = doc.Range(base, pr.End)
  t = QText(pr)
  QPass2 doc, base, t, cp
  QuotesInRange = qChanged
End Function

Private Sub QuotesAll()
  Dim rng As Range, p As Paragraph, cp As Long, doc As Document
  Set doc = ActiveDocument
  If Selection.Start <> Selection.End Then
    Set rng = Selection.Range
  Else
    Set rng = doc.Content
  End If
  cp = -1
  For Each p In rng.Paragraphs
    QuotesInRange doc, p.Range, cp
  Next p
End Sub

Private Function GetWin(ByVal sel As Selection, ByVal doc As Document, ByRef pr As Range, ByRef cp As Long) As Boolean
  Dim a As Long, b As Long, de As Long
  cp = sel.Start
  de = doc.Content.End
  a = cp - 2000
  If a < 0 Then a = 0
  b = cp + 2000
  If b > de Then b = de
  If b <= a Then Exit Function
  Set pr = doc.Range(a, b)
  GetWin = True
End Function

Private Function WordPass(ByVal doc As Document, ByVal pr As Range, ByRef cp As Long, ByVal incCaret As Boolean, ByVal edge As Boolean) As Boolean
  Dim t As String, base As Long, i As Long, j As Long, s0 As Long, e0 As Long
  Dim w As String, nw As String, skipit As Boolean, de As Long
  base = pr.Start
  de = doc.Content.End
  t = QText(pr)
  i = Len(t)
  Do While i >= 1
    If IsWordCh(Mid$(t, i, 1)) Then
      j = i
      Do While j > 1
        If Not IsWordCh(Mid$(t, j - 1, 1)) Then Exit Do
        j = j - 1
      Loop
      w = Mid$(t, j, i - j + 1)
      s0 = base + j - 1
      e0 = base + i
      If Len(w) >= 2 Then
        skipit = False
        If cp > s0 And cp < e0 Then skipit = True
        If cp = e0 And Not incCaret Then skipit = True
        If edge Then
          If j = 1 And base > 0 Then skipit = True
          If i = Len(t) And e0 < de Then skipit = True
        End If
        If Not skipit Then
          nw = ConvWord(w)
          If nw <> w Then
            If Not InExc(w) Then
              If SetCh(doc, s0, e0, w, nw) Then
                convCount = convCount + 1
                If e0 <= cp Then cp = cp + (Len(nw) - Len(w))
                WordPass = True
              End If
            End If
          End If
        End If
      End If
      i = j - 1
    Else
      i = i - 1
    End If
  Loop
End Function

Private Sub FixPara(ByVal incCaret As Boolean)
  Dim sel As Selection, doc As Document, pr As Range, cp As Long
  If Documents.Count = 0 Then Exit Sub
  Set sel = Selection
  If sel.Start <> sel.End Then Exit Sub
  Set doc = sel.Document
  If GetWin(sel, doc, pr, cp) Then
    If QuotesInRange(doc, pr, cp) Then sel.SetRange cp, cp
  End If
  If GetWin(sel, doc, pr, cp) Then
    If WordPass(doc, pr, cp, incCaret, True) Then sel.SetRange cp, cp
  End If
End Sub

Private Sub CleanPass(ByVal doc As Document, ByVal pr As Range, ByRef cp As Long)
  Dim t As String, base As Long, i As Long, j As Long, c As String, s0 As Long, e0 As Long
  base = pr.Start
  t = QText(pr)
  i = Len(t)
  Do While i >= 1
    c = Mid$(t, i, 1)
    If c = ChrW(&H200B) Or c = ChrW(-257) Or c = ChrW(&HAD) Or c = ChrW(&H200C) Or c = ChrW(&H200D) Then
      If DelRange(doc, base + i - 1, base + i, c) Then AdjCp cp, base + i - 1, base + i
      i = i - 1
    ElseIf c = " " Then
      j = i
      Do While j > 1
        If Mid$(t, j - 1, 1) <> " " Then Exit Do
        j = j - 1
      Loop
      If i - j >= 1 And j > 1 Then
        If Mid$(t, j - 1, 1) <> Chr(13) Then
          s0 = base + j
          e0 = base + i
          If DelRange(doc, s0, e0, Mid$(t, j + 1, i - j)) Then AdjCp cp, s0, e0
        End If
      End If
      i = j - 1
    Else
      i = i - 1
    End If
  Loop
End Sub

Public Sub PasteClean()
  Dim sel As Selection, doc As Document, p0 As Long, p1 As Long, pr As Range
  Dim cp As Long, app As Object, ok As Boolean
  If Documents.Count = 0 Then Exit Sub
  pasteCount = pasteCount + 1
  Set sel = Selection
  If Not pasteFlag Then
    sel.Paste
    Exit Sub
  End If
  Set app = Application
  p0 = sel.Start
  On Error Resume Next
  SafeUndoStart "Joylash"
  Err.Clear
  sel.Paste
  If Err.Number <> 0 Then
    Err.Clear
    SafeUndoEnd
    Exit Sub
  End If
  ok = Allowed()
  If ok Then
    If excList Is Nothing Then Set excList = LoadExc()
    Set doc = sel.Document
    p1 = sel.End
    If p1 > p0 Then
      busy = True
      cp = p1
      Set pr = doc.Range(p0, cp)
      CleanPass doc, pr, cp
      Set pr = doc.Range(p0, cp)
      QuotesInRange doc, pr, cp
      Set pr = doc.Range(p0, cp)
      WordPass doc, pr, cp, True, False
      sel.SetRange cp, cp
      busy = False
      lastPos = Selection.Start
    End If
  End If
  Err.Clear
  SafeUndoEnd
End Sub

Public Sub EditPaste()
  PasteClean
End Sub

Public Sub Paste()
  PasteClean
End Sub

Public Sub TogglePaste()
  pasteFlag = Not pasteFlag
  SaveSetting "Alifbo", "Cfg" & TAG, "paste", IIf(pasteFlag, "1", "0")
  RefreshAutoBtn
  If pasteFlag Then
    Application.StatusBar = "Alifbo: Ctrl+V bilan joylashda matn avtomatik to'g'rilanadi."
  Else
    Application.StatusBar = "Alifbo: joylashda avto-to'g'rilash o'chirildi."
  End If
End Sub

Public Sub Diag()
  Dim m As String
  If autoFlag Then m = "Avto-rejim: YOQILGAN" Else m = "Avto-rejim: o'chirilgan"
  m = m & vbCrLf & "Hodisalar: " & evCount & ", taymer: " & tickCount & ", almashtirilgan so'z: " & convCount
  If pasteFlag Then m = m & vbCrLf & "Ctrl+V tozalash: YOQILGAN" Else m = m & vbCrLf & "Ctrl+V tozalash: o'chirilgan"
  m = m & " (chaqirilgan: " & pasteCount & ")"
  If Allowed() Then m = m & vbCrLf & "Ruxsat: ha" Else m = m & vbCrLf & "Ruxsat: YO'Q (sinov tugagan)"
  If Len(lastErr) > 0 Then m = m & vbCrLf & "Oxirgi xato: " & lastErr
  m = m & vbCrLf & vbCrLf & "Sinov: o'zim -> " & ConvWord("o'zim") & ", to'g'ri -> " & ConvWord("to'g'ri") & ", shahar -> " & ConvWord("shahar") & ", chiroyli -> " & ConvWord("chiroyli")
  MsgBox m, vbInformation, "Alifbo - diagnostika"
End Sub

Public Sub AutoTyped(ByVal Sel As Selection)
  Exit Sub ' OO, SS xatolarining oldini olish uchun fonda aralashmaydi
  Dim pos As Long, dn As String, typed As Boolean
  If busy Then Exit Sub
  evCount = evCount + 1
  On Error GoTo Done
  pos = Sel.Start
  dn = Sel.Document.Name
  typed = (dn = lastDoc And pos = lastPos + 1 And Sel.Start = Sel.End)
  lastDoc = dn
  lastPos = pos
  If Not typed Then Exit Sub
  If Not autoFlag Then Exit Sub
  If pos < 2 Then Exit Sub
  If IsWordCh(Sel.Document.Range(pos - 1, pos).Text) Then Exit Sub
  If Not Allowed() Then Exit Sub
  busy = True
  FixPara False
  busy = False
  lastPos = Selection.Start
  Exit Sub
Done:
  busy = False
End Sub

Private Sub ScheduleTick()
  If tickScheduled Then Exit Sub
  tickScheduled = True
  Application.OnTime Now + TimeValue("00:00:01"), "Alifbo.Tick"
End Sub

Public Sub Tick()
  Exit Sub ' OO, SS xatolarining oldini olish
  Dim p As Long
  tickScheduled = False
  tickCount = tickCount + 1
  If Not autoFlag Then Exit Sub
  On Error Resume Next
  If Documents.Count > 0 Then
    If Not busy Then
      If Allowed() Then
        If Selection.Start = Selection.End Then
          p = Selection.Start
          If p = tickPos Then
            busy = True
            FixPara True
            busy = False
            lastPos = Selection.Start
          End If
          tickPos = Selection.Start
        End If
      End If
    End If
  End If
  busy = False
  ScheduleTick
End Sub

Private Sub RefreshAutoBtn()
  Dim b As Object
  On Error Resume Next
  Set b = Application.CommandBars.FindControl(, , "alifbo_auto")
  If Not b Is Nothing Then
    If autoFlag Then
      b.Caption = "Avto: YOQILGAN"
      b.State = -1
    Else
      b.Caption = "Avto: o'chirilgan"
      b.State = 0
    End If
  End If
  Set b = Nothing
  Set b = Application.CommandBars.FindControl(, , "alifbo_paste")
  If Not b Is Nothing Then
    If pasteFlag Then
      b.Caption = "Ctrl+V: YOQILGAN"
      b.State = -1
    Else
      b.Caption = "Ctrl+V: o'chirilgan"
      b.State = 0
    End If
  End If
End Sub

Public Sub ToggleAuto()
  If Not autoFlag Then
    If Not Gate() Then Exit Sub
    Set excList = LoadExc()
    allowStamp = 0
    autoFlag = True
    ScheduleTick
    Application.StatusBar = "Alifbo: avto-o'tkazish yoqildi. Har bir so'z tugagach yangi alifboga o'tkaziladi."
  Else
    autoFlag = False
    Application.StatusBar = "Alifbo: avto-o'tkazish o'chirildi."
  End If
  SaveSetting "Alifbo", "Cfg" & TAG, "auto", IIf(autoFlag, "1", "0")
  RefreshAutoBtn
End Sub

Private Sub AddBtn(ByVal cb As Object, ByVal cap As String, ByVal act As String, ByVal tip As String, Optional ByVal tg As String = "")
  Dim b As Object
  Set b = cb.Controls.Add(1, , , , True)
  b.Caption = cap
  b.OnAction = act
  b.Style = 2
  b.TooltipText = tip
  If Len(tg) > 0 Then b.Tag = tg
End Sub

Public Sub AutoExec()
  Dim cb As Object
  On Error Resume Next
  Application.CommandBars("Alifbo").Delete
  On Error GoTo 0
  Dim d0 As Long
  d0 = LeftMinutes()
  Set cb = Application.CommandBars.Add("Alifbo", 1, False, True)
  AddBtn cb, "Yangi alifbo", "Alifbo.ToNew", "Matnni yangi alifboga o'tkazish (belgilangan qism yoki butun hujjat)"
  AddBtn cb, "Eski alifbo", "Alifbo.ToOld", "Yangi alifbodan eskisiga qaytarish"
  AddBtn cb, "Istisnolar", "Alifbo.EditExc", "O'zgartirilmaydigan so'zlar ro'yxati"
  AddBtn cb, "Avto: o'chirilgan", "Alifbo.ToggleAuto", "Eskicha yozing, dastur har bir so'zni yangi alifboga o'zi o'tkazadi", "alifbo_auto"
  AddBtn cb, "Ctrl+V: YOQILGAN", "Alifbo.TogglePaste", "Joylashda (Ctrl+V) matnni avtomatik to'g'rilash", "alifbo_paste"
  AddBtn cb, "Litsenziya", "Alifbo.LicenseInfo", "Sinov muddati va litsenziya kodi"
  AddBtn cb, "Diagnostika", "Alifbo.Diag", "Dastur holatini ko'rsatadi"
  AddBtn cb, "Varaqlarni bo'lish", "Alifbo.SplitTool", "Varaqlarni bo'lish oynasi: tartib, format, papka va nom tanlanadi"
  AddBtn cb, "Bo'lish (oddiy)", "Alifbo.SplitDoc", "Hujjatni varaqlar bo'yicha alohida fayllarga saqlaydi (oddiy usul)"
  AddBtn cb, "Fayllar", "Alifbo.OpenTools", "Fayllarni (Word, Excel, PowerPoint) yangi alifboga o'tkazish oynasi"
  AddBtn cb, "Muallif", "Alifbo.About", "Dastur muallifi va aloqa"
  cb.Visible = True
  Set ev = New clsAppEvents
  Set ev.App = Application
  autoFlag = (GetSetting("Alifbo", "Cfg" & TAG, "auto", "0") = "1")
  pasteFlag = (GetSetting("Alifbo", "Cfg" & TAG, "paste", "1") = "1")
  Set excList = LoadExc()
  If autoFlag Then ScheduleTick
  RefreshAutoBtn
End Sub

Public Sub RbWdNew(control As IRibbonControl)
  ToNew
End Sub

Public Sub RbWdOld(control As IRibbonControl)
  ToOld
End Sub

Public Sub RbWdSplit(control As IRibbonControl)
  SplitTool
End Sub

Public Sub RbWdExc(control As IRibbonControl)
  EditExc
End Sub

Public Sub RbWdTools(control As IRibbonControl)
  OpenTools
End Sub

Public Sub RbWdLic(control As IRibbonControl)
  LicenseInfo
End Sub

Public Sub RbWdAbout(control As IRibbonControl)
  About
End Sub

Public Sub RbWdKrillNew(control As IRibbonControl)
  KrillToNew
End Sub

Public Sub RbWdKrillOld(control As IRibbonControl)
  KrillToOld
End Sub

' ==============================================================
' Kirill -> Yangi / Lotin (Word uchun)
' ==============================================================
Public Sub KrillConvDoc(ByVal toNew As Boolean)
  Dim rng As Object, t As String, newT As String
  Dim changed As Long, wi As Long, wr As Object, words As Object
  Dim excColl As Collection

  On Error GoTo ErrOut
  If Documents.Count = 0 Then Exit Sub

  If Selection.Type <> 1 Then
    Set rng = Selection.Range
  Else
    Set rng = ActiveDocument.Content
  End If

  t = rng.Text
  If Len(t) = 0 Then
    MsgBox "Hujjat bo'sh.", vbInformation, ProdName()
    Exit Sub
  End If

  Set excColl = LoadExc()
  Set words = rng.Words
  Application.ScreenUpdating = False

  For wi = 1 To words.Count
    On Error Resume Next
    Set wr = words(wi)
    If Err.Number <> 0 Then Err.Clear: GoTo NextWord
    On Error GoTo ErrOut

    t = wr.Text
    If Len(Trim$(t)) > 0 Then
      If IsCyrInStr(t) Then
        If Not InExc(Trim$(t)) Then
          newT = KrillConvStr(t, toNew)
          If newT <> t Then
            wr.Text = newT
            changed = changed + 1
          End If
        End If
      End If
    End If
NextWord:
  Next wi

  Application.ScreenUpdating = True
  Dim suf As String
  suf = StatusSuffix()
  Dim modeName As String
  If toNew Then modeName = "Yangi alifbo" Else modeName = "Lotin alifbo"
  MsgBox "Tayyor (" & modeName & "): " & changed & " ta so'z o'zgardi." & suf, vbInformation, ProdName()
  Exit Sub

ErrOut:
  Application.ScreenUpdating = True
  MsgBox "Xatolik: " & Err.Description, vbExclamation, ProdName()
End Sub

Public Sub KrillToNew()
  If Not Gate() Then Exit Sub
  If Documents.Count = 0 Then Exit Sub
  SafeUndoStart "KrillYangi"
  KrillConvDoc True
  SafeUndoEnd
End Sub

Public Sub KrillToOld()
  If Not Gate() Then Exit Sub
  If Documents.Count = 0 Then Exit Sub
  SafeUndoStart "KrillLotin"
  KrillConvDoc False
  SafeUndoEnd
End Sub