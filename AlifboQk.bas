Attribute VB_Name = "AlifboQk"
Option Explicit

' ==============================================================
' AlifboQk.bas  -  Qoraqalpoq tili (Kirill <-> Lotin 2016 reform)
' Word, Excel va PowerPoint uchun to'liq integratsiya
' Muallif: Fozliddinov Safarmurod | @Safar_master | +998939530473
' ==============================================================

Public Function QkIsCyrChar(ByVal c As String) As Boolean
  Dim a As Long
  If Len(c) <> 1 Then Exit Function
  a = AscW(c)
  If a < 0 Then a = a + 65536
  QkIsCyrChar = (a >= &H400 And a <= &H52F)
End Function

Public Function QkIsCyrVowel(ByVal c As String) As Boolean
  Select Case c
    Case ChrW(&H410), ChrW(&H430) ' А а
      QkIsCyrVowel = True
    Case ChrW(&H4D8), ChrW(&H4D9) ' Ә ә
      QkIsCyrVowel = True
    Case ChrW(&H415), ChrW(&H435) ' Е е
      QkIsCyrVowel = True
    Case ChrW(&H401), ChrW(&H451) ' Ё ё
      QkIsCyrVowel = True
    Case ChrW(&H418), ChrW(&H438) ' И и
      QkIsCyrVowel = True
    Case ChrW(&H41E), ChrW(&H43E) ' О о
      QkIsCyrVowel = True
    Case ChrW(&H4E8), ChrW(&H4E9) ' Ө ө
      QkIsCyrVowel = True
    Case ChrW(&H423), ChrW(&H443) ' У у
      QkIsCyrVowel = True
    Case ChrW(&H4AE), ChrW(&H4AF) ' Ү ү
      QkIsCyrVowel = True
    Case ChrW(&H40E), ChrW(&H45E) ' Ў ў
      QkIsCyrVowel = True
    Case ChrW(&H42B), ChrW(&H44B) ' Ы ы
      QkIsCyrVowel = True
    Case ChrW(&H42D), ChrW(&H44D) ' Э э
      QkIsCyrVowel = True
    Case ChrW(&H42E), ChrW(&H44E) ' Ю ю
      QkIsCyrVowel = True
    Case ChrW(&H42F), ChrW(&H44F) ' Я я
      QkIsCyrVowel = True
  End Select
End Function

Public Function QkIsLatVowel(ByVal c As String) As Boolean
  Dim lc As String
  lc = LCase$(c)
  Select Case lc
    Case "a", "e", "i", "o", "u"
      QkIsLatVowel = True
    Case ChrW(&HE1) ' á
      QkIsLatVowel = True
    Case ChrW(&HF3) ' ó
      QkIsLatVowel = True
    Case ChrW(&HFA) ' ú
      QkIsLatVowel = True
    Case ChrW(&H131), ChrW(&HED) ' ı, í
      QkIsLatVowel = True
  End Select
End Function

Public Function QkKrillToLatinWord(ByVal w As String) As String
  Dim i As Long, n As Long, c As String, c1 As String
  Dim o As String, lc As String
  Dim isUp As Boolean, cUp As Boolean
  Dim ePrev As String, eIsWordStart As Boolean, eAfterVowel As Boolean
  Dim piece As String

  n = Len(w)
  If n = 0 Then QkKrillToLatinWord = "": Exit Function

  isUp = (w = UCase$(w) And w <> LCase$(w) And n > 1)
  i = 1

  Do While i <= n
    c = Mid$(w, i, 1)
    c1 = Mid$(w, i + 1, 1)

    If Not QkIsCyrChar(c) Then
      o = o & c
      i = i + 1
    Else
      lc = LCase$(c)
      cUp = (c <> LCase$(c))
      piece = ""

      Select Case AscW(lc)
        Case &H430: piece = "a" ' а
        Case &H4D9: piece = ChrW(&HE1) ' ә -> á
        Case &H431: piece = "b" ' б
        Case &H432: piece = "v" ' в
        Case &H433: piece = "g" ' г
        Case &H493: piece = ChrW(&H1F5) ' ғ -> ǵ
        Case &H434: piece = "d" ' д

        Case &H435 ' е
          If i = 1 Then ePrev = "" Else ePrev = Mid$(w, i - 1, 1)
          eIsWordStart = (i = 1) Or (Not QkIsCyrChar(ePrev))
          eAfterVowel = QkIsCyrVowel(ePrev) Or (AscW(ePrev) = &H44A) Or (AscW(ePrev) = &H44C)
          If eIsWordStart Or eAfterVowel Then
            piece = "ye"
            If isUp Then piece = "YE" Else If cUp Then piece = "Ye"
          Else
            piece = "e"
            If isUp Or cUp Then piece = "E"
          End If
          o = o & piece
          i = i + 1
          GoTo NextChar

        Case &H451 ' ё -> yo
          piece = "yo"
          If isUp Then piece = "YO" Else If cUp Then piece = "Yo"
          o = o & piece
          i = i + 1
          GoTo NextChar

        Case &H436: piece = "j" ' ж
        Case &H437: piece = "z" ' з
        Case &H438: piece = "i" ' и
        Case &H439: piece = "y" ' й
        Case &H43A: piece = "k" ' к
        Case &H49B: piece = "q" ' қ -> q
        Case &H43B: piece = "l" ' л
        Case &H43C: piece = "m" ' м
        Case &H43D: piece = "n" ' н
        Case &H4A3: piece = ChrW(&H144) ' ң -> ń
        Case &H43E: piece = "o" ' о
        Case &H4E9: piece = ChrW(&HF3) ' ө -> ó
        Case &H43F: piece = "p" ' п
        Case &H440: piece = "r" ' р
        Case &H441: piece = "s" ' с
        Case &H442: piece = "t" ' т
        Case &H443: piece = "u" ' у
        Case &H4AF: piece = ChrW(&HFA) ' ү -> ú
        Case &H45E: piece = "w" ' ў -> w
        Case &H444: piece = "f" ' ф
        Case &H445: piece = "x" ' х
        Case &H4B3: piece = "h" ' ҳ -> h
        Case &H446: piece = "c" ' ц -> c (2016 reform)

        Case &H447 ' ч -> ch
          piece = "ch"
          If isUp Then piece = "CH" Else If cUp Then piece = "Ch"
          o = o & piece
          i = i + 1
          GoTo NextChar

        Case &H448, &H449 ' ш, щ -> sh
          piece = "sh"
          If isUp Then piece = "SH" Else If cUp Then piece = "Sh"
          o = o & piece
          i = i + 1
          GoTo NextChar

        Case &H44A ' ъ -> '
          piece = "'"

        Case &H44B ' ы -> ı (dotless i)
          piece = ChrW(&H131)
          If isUp Or cUp Then piece = ChrW(&HCD) ' Í

        Case &H44C ' ь
          Dim sNext As Long
          sNext = 0
          If Len(c1) > 0 Then sNext = AscW(LCase$(c1))
          If sNext = &H43E Then
            piece = "yo": If isUp Then piece = "YO" Else If cUp Then piece = "Yo"
            o = o & piece: i = i + 2: GoTo NextChar
          ElseIf sNext = &H435 Then
            piece = "ye": If isUp Then piece = "YE" Else If cUp Then piece = "Ye"
            o = o & piece: i = i + 2: GoTo NextChar
          ElseIf sNext = &H44E Then
            piece = "yu": If isUp Then piece = "YU" Else If cUp Then piece = "Yu"
            o = o & piece: i = i + 2: GoTo NextChar
          ElseIf sNext = &H44F Then
            piece = "ya": If isUp Then piece = "YA" Else If cUp Then piece = "Ya"
            o = o & piece: i = i + 2: GoTo NextChar
          Else
            piece = ""
          End If

        Case &H44D: piece = "e" ' э -> e

        Case &H44E ' ю -> yu
          piece = "yu"
          If isUp Then piece = "YU" Else If cUp Then piece = "Yu"
          o = o & piece
          i = i + 1
          GoTo NextChar

        Case &H44F ' я -> ya
          piece = "ya"
          If isUp Then piece = "YA" Else If cUp Then piece = "Ya"
          o = o & piece
          i = i + 1
          GoTo NextChar

        Case Else
          piece = c
      End Select

      If Len(piece) = 1 Then
        If isUp Or cUp Then
          Select Case piece
            Case ChrW(&HE1): piece = ChrW(&HC1) ' Á
            Case ChrW(&H1F5): piece = ChrW(&H1F4) ' Ǵ
            Case ChrW(&H144): piece = ChrW(&H143) ' Ń
            Case ChrW(&HF3): piece = ChrW(&HD3) ' Ó
            Case ChrW(&HFA): piece = ChrW(&HDA) ' Ú
            Case ChrW(&H131), "i": piece = ChrW(&HCD) ' Í
            Case Else: piece = UCase$(piece)
          End Select
        End If
      End If

      o = o & piece
      i = i + 1
    End If
NextChar:
  Loop

  QkKrillToLatinWord = o
End Function

Public Function QkKrillToLatinStr(ByVal s As String) As String
  Dim i As Long, n As Long, c As String, o As String
  Dim wSt As Long, inW As Boolean
  n = Len(s)
  i = 1
  Do While i <= n
    c = Mid$(s, i, 1)
    If QkIsCyrChar(c) Then
      If Not inW Then wSt = i: inW = True
      i = i + 1
    Else
      If inW Then
        o = o & QkKrillToLatinWord(Mid$(s, wSt, i - wSt))
        inW = False
      End If
      o = o & c
      i = i + 1
    End If
  Loop
  If inW Then
    o = o & QkKrillToLatinWord(Mid$(s, wSt))
  End If
  QkKrillToLatinStr = o
End Function

Public Function QkLatinToKrillWord(ByVal w As String) As String
  Dim i As Long, n As Long, c As String, c1 As String, c2 As String
  Dim o As String, lc As String, lc2 As String
  Dim isUp As Boolean, cUp As Boolean
  Dim prevC As String, isStart As Boolean

  n = Len(w)
  If n = 0 Then QkLatinToKrillWord = "": Exit Function

  isUp = (w = UCase$(w) And w <> LCase$(w) And n > 1)
  i = 1

  Do While i <= n
    c = Mid$(w, i, 1)
    c1 = ""
    If i < n Then c1 = Mid$(w, i + 1, 1)
    lc = LCase$(c)
    cUp = (c <> LCase$(c))

    isStart = (i = 1)
    If i > 1 Then prevC = Mid$(w, i - 1, 1) Else prevC = ""

    ' 2 ta harfli birikmalar (sh, ch, ye, yo, yu, ya)
    If Len(c1) > 0 Then
      lc2 = lc & LCase$(c1)
      Select Case lc2
        Case "sh"
          If isUp Then o = o & ChrW(&H428) Else If cUp Then o = o & ChrW(&H428) Else o = o & ChrW(&H448)
          i = i + 2: GoTo NextLChar
        Case "ch"
          If isUp Then o = o & ChrW(&H427) Else If cUp Then o = o & ChrW(&H427) Else o = o & ChrW(&H447)
          i = i + 2: GoTo NextLChar
        Case "ye"
          If isUp Then o = o & ChrW(&H415) Else If cUp Then o = o & ChrW(&H415) Else o = o & ChrW(&H435)
          i = i + 2: GoTo NextLChar
        Case "yo"
          If isUp Then o = o & ChrW(&H401) Else If cUp Then o = o & ChrW(&H401) Else o = o & ChrW(&H451)
          i = i + 2: GoTo NextLChar
        Case "yu"
          If isUp Then o = o & ChrW(&H42E) Else If cUp Then o = o & ChrW(&H42E) Else o = o & ChrW(&H44E)
          i = i + 2: GoTo NextLChar
        Case "ya"
          If isUp Then o = o & ChrW(&H42F) Else If cUp Then o = o & ChrW(&H42F) Else o = o & ChrW(&H44F)
          i = i + 2: GoTo NextLChar
      End Select
    End If

    ' Bitta harfli almashtirishlar
    Select Case lc
      Case "a": If isUp Or cUp Then o = o & ChrW(&H410) Else o = o & ChrW(&H430)
      Case ChrW(&HE1): If isUp Or cUp Then o = o & ChrW(&H4D8) Else o = o & ChrW(&H4D9) ' á -> ә
      Case "b": If isUp Or cUp Then o = o & ChrW(&H411) Else o = o & ChrW(&H431)
      Case "v": If isUp Or cUp Then o = o & ChrW(&H412) Else o = o & ChrW(&H432)
      Case "g": If isUp Or cUp Then o = o & ChrW(&H413) Else o = o & ChrW(&H433)
      Case ChrW(&H1F5): If isUp Or cUp Then o = o & ChrW(&H492) Else o = o & ChrW(&H493) ' ǵ -> ғ
      Case "d": If isUp Or cUp Then o = o & ChrW(&H414) Else o = o & ChrW(&H434)
      Case "e"
        If isStart Then
          If isUp Or cUp Then o = o & ChrW(&H42D) Else o = o & ChrW(&H44D) ' э
        Else
          If isUp Or cUp Then o = o & ChrW(&H415) Else o = o & ChrW(&H435) ' е
        End If
      Case "j": If isUp Or cUp Then o = o & ChrW(&H416) Else o = o & ChrW(&H436)
      Case "z": If isUp Or cUp Then o = o & ChrW(&H417) Else o = o & ChrW(&H437)
      Case "i": If isUp Or cUp Then o = o & ChrW(&H418) Else o = o & ChrW(&H438)
      Case ChrW(&H131), ChrW(&HED): If isUp Or cUp Then o = o & ChrW(&H42B) Else o = o & ChrW(&H44B) ' ı / í -> ы
      Case "y": If isUp Or cUp Then o = o & ChrW(&H419) Else o = o & ChrW(&H439)
      Case "k": If isUp Or cUp Then o = o & ChrW(&H41A) Else o = o & ChrW(&H43A)
      Case "q": If isUp Or cUp Then o = o & ChrW(&H49A) Else o = o & ChrW(&H49B) ' q -> қ
      Case "l": If isUp Or cUp Then o = o & ChrW(&H41B) Else o = o & ChrW(&H43B)
      Case "m": If isUp Or cUp Then o = o & ChrW(&H41C) Else o = o & ChrW(&H43C)
      Case "n": If isUp Or cUp Then o = o & ChrW(&H41D) Else o = o & ChrW(&H43D)
      Case ChrW(&H144): If isUp Or cUp Then o = o & ChrW(&H4A2) Else o = o & ChrW(&H4A3) ' ń -> ң
      Case "o": If isUp Or cUp Then o = o & ChrW(&H41E) Else o = o & ChrW(&H43E)
      Case ChrW(&HF3): If isUp Or cUp Then o = o & ChrW(&H4E8) Else o = o & ChrW(&H4E9) ' ó -> ө
      Case "p": If isUp Or cUp Then o = o & ChrW(&H41F) Else o = o & ChrW(&H43F)
      Case "r": If isUp Or cUp Then o = o & ChrW(&H420) Else o = o & ChrW(&H440)
      Case "s": If isUp Or cUp Then o = o & ChrW(&H421) Else o = o & ChrW(&H441)
      Case "t": If isUp Or cUp Then o = o & ChrW(&H422) Else o = o & ChrW(&H442)
      Case "u": If isUp Or cUp Then o = o & ChrW(&H423) Else o = o & ChrW(&H443)
      Case ChrW(&HFA): If isUp Or cUp Then o = o & ChrW(&H4AE) Else o = o & ChrW(&H4AF) ' ú -> ү
      Case "w": If isUp Or cUp Then o = o & ChrW(&H40E) Else o = o & ChrW(&H45E) ' w -> ў
      Case "f": If isUp Or cUp Then o = o & ChrW(&H424) Else o = o & ChrW(&H444)
      Case "x": If isUp Or cUp Then o = o & ChrW(&H425) Else o = o & ChrW(&H445)
      Case "h": If isUp Or cUp Then o = o & ChrW(&H4B2) Else o = o & ChrW(&H4B3) ' h -> ҳ
      Case "c": If isUp Or cUp Then o = o & ChrW(&H426) Else o = o & ChrW(&H446) ' c -> ц
      Case "'", ChrW(&H2BC), ChrW(&H2019): o = o & ChrW(&H44A) ' ' -> ъ
      Case Else: o = o & c
    End Select
    i = i + 1

NextLChar:
  Loop

  QkLatinToKrillWord = o
End Function

Public Function QkLatinToKrillStr(ByVal s As String) As String
  Dim i As Long, n As Long, c As String, o As String
  Dim wSt As Long, inW As Boolean, a As Long
  n = Len(s)
  i = 1
  Do While i <= n
    c = Mid$(s, i, 1)
    a = AscW(c)
    If a < 0 Then a = a + 65536

    ' Harflar: A-Z, a-z, á, ǵ, ń, ó, ú, ı, í, '
    Dim isLat As Boolean
    isLat = ((a >= 65 And a <= 90) Or (a >= 97 And a <= 122) Or _
             a = &HC1 Or a = &HE1 Or a = &H1F4 Or a = &H1F5 Or _
             a = &H143 Or a = &H144 Or a = &HD3 Or a = &HF3 Or _
             a = &HDA Or a = &HFA Or a = &H131 Or a = &HCD Or a = &HED Or _
             a = 39 Or a = &H2BC Or a = &H2019)

    If isLat Then
      If Not inW Then wSt = i: inW = True
      i = i + 1
    Else
      If inW Then
        o = o & QkLatinToKrillWord(Mid$(s, wSt, i - wSt))
        inW = False
      End If
      o = o & c
      i = i + 1
    End If
  Loop
  If inW Then
    o = o & QkLatinToKrillWord(Mid$(s, wSt))
  End If
  QkLatinToKrillStr = o
End Function

' ==============================================================
' Word amallari
' ==============================================================
Public Sub QkWdConvDoc(ByVal toLatin As Boolean)
  Dim rng As Object, t As String, newT As String
  Dim changed As Long, wi As Long, wr As Object, words As Object

  On Error GoTo ErrWord
  If Documents.Count = 0 Then Exit Sub

  If Selection.Type <> 1 Then
    Set rng = Selection.Range
  Else
    Set rng = ActiveDocument.Content
  End If

  t = rng.Text
  If Len(t) = 0 Then
    MsgBox "Hujjat bo'sh.", vbInformation, "AlifboOtkazgich"
    Exit Sub
  End If

  Set words = rng.Words
  Application.ScreenUpdating = False

  For wi = 1 To words.Count
    On Error Resume Next
    Set wr = words(wi)
    If Err.Number <> 0 Then Err.Clear: GoTo WNextWord
    On Error GoTo ErrWord

    t = wr.Text
    If Len(Trim$(t)) > 0 Then
      If toLatin Then
        newT = QkKrillToLatinStr(t)
      Else
        newT = QkLatinToKrillStr(t)
      End If
      If newT <> t Then
        wr.Text = newT
        changed = changed + 1
      End If
    End If
WNextWord:
  Next wi

  Application.ScreenUpdating = True
  Dim modeStr As String
  If toLatin Then modeStr = "Qaraqalpaqsha (Kirill -> Latın)" Else modeStr = "Qaraqalpaqsha (Latın -> Kirill)"
  MsgBox "Tayyor: " & changed & " ta so'z o'zgardi (" & modeStr & ").", vbInformation, "AlifboOtkazgich"
  Exit Sub

ErrWord:
  Application.ScreenUpdating = True
  MsgBox "Xatolik: " & Err.Description, vbExclamation, "AlifboOtkazgich"
End Sub

Public Sub RbWdQkToLat(control As IRibbonControl)
  QkWdConvDoc True
End Sub

Public Sub RbWdQkToKrill(control As IRibbonControl)
  QkWdConvDoc False
End Sub

' ==============================================================
' Excel amallari
' ==============================================================
Private Function QkXlConvArea(ByVal area As Object, ByVal toLatin As Boolean) As Long
  Dim arr As Variant, r As Long, c As Long, nr As Long, nc As Long
  Dim v As String, nv As String, n As Long, changed As Boolean
  On Error Resume Next
  If area Is Nothing Then Exit Function
  If area.Cells.Count = 1 Then
    If Not IsError(area.Value) Then
      If Not area.HasFormula Then
        If VarType(area.Value) = vbString Then
          v = CStr(area.Value)
          If Len(v) > 0 Then
            If toLatin Then nv = QkKrillToLatinStr(v) Else nv = QkLatinToKrillStr(v)
            If Err.Number = 0 And Len(nv) > 0 And nv <> v Then
              area.Value = nv
              QkXlConvArea = 1
            End If
          End If
        End If
      End If
    End If
    Exit Function
  End If
  arr = area.Value
  If Not IsArray(arr) Then Exit Function
  nr = UBound(arr, 1)
  nc = UBound(arr, 2)
  For r = 1 To nr
    For c = 1 To nc
      If VarType(arr(r, c)) = vbString Then
        v = arr(r, c)
        If Len(v) > 0 Then
          If toLatin Then nv = QkKrillToLatinStr(v) Else nv = QkLatinToKrillStr(v)
          If Err.Number = 0 And Len(nv) > 0 And nv <> v Then
            arr(r, c) = nv
            n = n + 1
            changed = True
          End If
        End If
      End If
    Next c
  Next r
  If changed Then area.Value = arr
  QkXlConvArea = n
End Function

Public Sub QkXlRun(ByVal toLatin As Boolean)
  Dim n As Long, rng As Object, txt As Object, ar As Object
  If Workbooks.Count = 0 Then Exit Sub
  Application.ScreenUpdating = False
  On Error GoTo XlFail

  Set rng = ActiveSheet.UsedRange
  If TypeName(Selection) = "Range" Then
    If Selection.Cells.Count > 1 Then Set rng = Selection
  End If

  If rng.Cells.Count = 1 Then
    n = QkXlConvArea(rng, toLatin)
  Else
    Set txt = Nothing
    On Error Resume Next
    Set txt = rng.SpecialCells(2, 2)
    On Error GoTo XlFail
    If txt Is Nothing Then
      n = QkXlConvArea(rng, toLatin)
    Else
      For Each ar In txt.Areas
        n = n + QkXlConvArea(ar, toLatin)
      Next ar
    End If
  End If

  Application.ScreenUpdating = True
  Dim modeStr As String
  If toLatin Then modeStr = "Qaraqalpaqsha (Kirill -> Latın)" Else modeStr = "Qaraqalpaqsha (Latın -> Kirill)"
  MsgBox "Tayyor: " & n & " ta katak o'zgardi (" & modeStr & ").", vbInformation, "AlifboOtkazgich"
  Exit Sub
XlFail:
  Application.ScreenUpdating = True
  MsgBox "Xatolik: " & Err.Description, vbExclamation, "AlifboOtkazgich"
End Sub

Public Sub RbXlQkToLat(control As IRibbonControl)
  QkXlRun True
End Sub

Public Sub RbXlQkToKrill(control As IRibbonControl)
  QkXlRun False
End Sub

' ==============================================================
' PowerPoint amallari
' ==============================================================
Private Sub QkPptConvText(ByVal tr As Object, ByVal toLatin As Boolean, ByRef cnt As Long)
  Dim i As Long, r As Object, t As String, nt As String, rc As Long
  On Error Resume Next
  If tr Is Nothing Then Exit Sub
  rc = tr.Runs.Count
  If rc <= 1 Then
    t = tr.Text
    If Len(t) > 0 Then
      If toLatin Then nt = QkKrillToLatinStr(t) Else nt = QkLatinToKrillStr(t)
      If Err.Number = 0 And nt <> t Then
        tr.Text = nt
        cnt = cnt + 1
      End If
    End If
    Exit Sub
  End If
  For i = 1 To rc
    Set r = tr.Runs(i)
    t = r.Text
    If Len(t) > 0 Then
      If toLatin Then nt = QkKrillToLatinStr(t) Else nt = QkLatinToKrillStr(t)
      If Err.Number = 0 And nt <> t Then
        r.Text = nt
        cnt = cnt + 1
      End If
    End If
  Next i
End Sub

Private Sub QkPptShape(ByVal shp As Object, ByVal toLatin As Boolean, ByRef cnt As Long)
  Dim i As Long, rr As Long, cc As Long
  On Error Resume Next
  If shp.Type = 6 Then ' msoGroup
    For i = 1 To shp.GroupItems.Count
      QkPptShape shp.GroupItems(i), toLatin, cnt
    Next i
    Exit Sub
  End If
  If shp.HasTable Then
    For rr = 1 To shp.Table.Rows.Count
      For cc = 1 To shp.Table.Columns.Count
        QkPptConvText shp.Table.Cell(rr, cc).Shape.TextFrame.TextRange, toLatin, cnt
      Next cc
    Next rr
  End If
  If shp.HasTextFrame Then
    If shp.TextFrame.HasText Then QkPptConvText shp.TextFrame.TextRange, toLatin, cnt
  End If
End Sub

Public Sub QkPptRun(ByVal toLatin As Boolean)
  Dim sld As Object, shp As Object, cnt As Long
  If Presentations.Count = 0 Then Exit Sub
  cnt = 0
  On Error GoTo PptFail
  For Each sld In ActivePresentation.Slides
    For Each shp In sld.Shapes
      QkPptShape shp, toLatin, cnt
    Next shp
    On Error Resume Next
    For Each shp In sld.NotesPage.Shapes
      QkPptShape shp, toLatin, cnt
    Next shp
    On Error GoTo PptFail
  Next sld
  Dim modeStr As String
  If toLatin Then modeStr = "Qaraqalpaqsha (Kirill -> Latın)" Else modeStr = "Qaraqalpaqsha (Latın -> Kirill)"
  MsgBox "Tayyor: " & cnt & " ta matn bo'lagi o'zgardi (" & modeStr & ").", vbInformation, "AlifboOtkazgich"
  Exit Sub
PptFail:
  MsgBox "Xatolik: " & Err.Description, vbExclamation, "AlifboOtkazgich"
End Sub

Public Sub RbPptQkToLat(control As IRibbonControl)
  QkPptRun True
End Sub

Public Sub RbPptQkToKrill(control As IRibbonControl)
  QkPptRun False
End Sub
