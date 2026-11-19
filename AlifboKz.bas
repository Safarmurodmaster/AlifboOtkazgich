Attribute VB_Name = "AlifboKz"
Option Explicit

' ==============================================================
' AlifboKz.bas  -  Qozoq tili (Qazaqşa Kirill <-> Latın 2021)
' Word, Excel va PowerPoint uchun to'liq integratsiya
' Muallif: Fozliddinov Safarmurod | @Safar_master | +998939530473
' ==============================================================

Public Function KzIsCyrChar(ByVal c As String) As Boolean
  Dim a As Long
  If Len(c) <> 1 Then Exit Function
  a = AscW(c)
  If a < 0 Then a = a + 65536
  KzIsCyrChar = (a >= &H400 And a <= &H52F)
End Function

Public Function KzIsCyrVowel(ByVal c As String) As Boolean
  Select Case c
    Case ChrW(&H410), ChrW(&H430) ' А а
      KzIsCyrVowel = True
    Case ChrW(&H4D8), ChrW(&H4D9) ' Ә ә
      KzIsCyrVowel = True
    Case ChrW(&H415), ChrW(&H435) ' Е е
      KzIsCyrVowel = True
    Case ChrW(&H401), ChrW(&H451) ' Ё ё
      KzIsCyrVowel = True
    Case ChrW(&H418), ChrW(&H438) ' И и
      KzIsCyrVowel = True
    Case ChrW(&H406), ChrW(&H456) ' І і
      KzIsCyrVowel = True
    Case ChrW(&H41E), ChrW(&H43E) ' О о
      KzIsCyrVowel = True
    Case ChrW(&H4E8), ChrW(&H4E9) ' Ө ө
      KzIsCyrVowel = True
    Case ChrW(&H4B0), ChrW(&H4B1) ' Ұ ұ
      KzIsCyrVowel = True
    Case ChrW(&H4AE), ChrW(&H4AF) ' Ү ү
      KzIsCyrVowel = True
    Case ChrW(&H423), ChrW(&H443) ' У у
      KzIsCyrVowel = True
    Case ChrW(&H42B), ChrW(&H44B) ' Ы ы
      KzIsCyrVowel = True
    Case ChrW(&H42D), ChrW(&H44D) ' Э э
      KzIsCyrVowel = True
    Case ChrW(&H42E), ChrW(&H44E) ' Ю ю
      KzIsCyrVowel = True
    Case ChrW(&H42F), ChrW(&H44F) ' Я я
      KzIsCyrVowel = True
  End Select
End Function

Public Function KzKrillToLatinWord(ByVal w As String) As String
  Dim i As Long, n As Long, c As String, c1 As String
  Dim o As String, lc As String
  Dim isUp As Boolean, cUp As Boolean
  Dim ePrev As String, eIsWordStart As Boolean, eAfterVowel As Boolean
  Dim piece As String

  n = Len(w)
  If n = 0 Then KzKrillToLatinWord = "": Exit Function

  isUp = (w = UCase$(w) And w <> LCase$(w) And n > 1)
  i = 1

  Do While i <= n
    c = Mid$(w, i, 1)
    c1 = Mid$(w, i + 1, 1)

    If Not KzIsCyrChar(c) Then
      o = o & c
      i = i + 1
    Else
      lc = LCase$(c)
      cUp = (c <> LCase$(c))
      piece = ""

      Select Case AscW(lc)
        Case &H430: piece = "a" ' а
        Case &H4D9: piece = ChrW(&HE4) ' ә -> ä
        Case &H431: piece = "b" ' б
        Case &H432: piece = "v" ' в
        Case &H433: piece = "g" ' г
        Case &H493: piece = ChrW(&H11F) ' ғ -> ğ
        Case &H434: piece = "d" ' д

        Case &H435 ' е
          If i = 1 Then ePrev = "" Else ePrev = Mid$(w, i - 1, 1)
          eIsWordStart = (i = 1) Or (Not KzIsCyrChar(ePrev))
          eAfterVowel = KzIsCyrVowel(ePrev)
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
        Case &H4A3: piece = ChrW(&HF1) ' ң -> ñ
        Case &H43E: piece = "o" ' о
        Case &H4E9: piece = ChrW(&HF6) ' ө -> ö
        Case &H43F: piece = "p" ' п
        Case &H440: piece = "r" ' р
        Case &H441: piece = "s" ' с
        Case &H442: piece = "t" ' т
        Case &H4B1: piece = "u" ' ұ -> u
        Case &H4AF: piece = ChrW(&HFC) ' ү -> ü
        Case &H443: piece = "w" ' у -> w (qozoqcha lotin 2021)
        Case &H444: piece = "f" ' ф
        Case &H445: piece = "x" ' х
        Case &H4BB: piece = "h" ' һ -> h
        Case &H4B3: piece = "h" ' ҳ -> h
        Case &H446: piece = "c" ' ц -> c
        Case &H447: piece = ChrW(&HE7) ' ч -> ç
        Case &H448: piece = ChrW(&H15F) ' ш -> ş
        Case &H449: piece = ChrW(&H15F) & ChrW(&H15F) ' щ -> şş
        Case &H44A: piece = "'" ' ъ
        Case &H44B: piece = "y" ' ы -> y
        Case &H456: piece = "i" ' і -> i
        Case &H44C: piece = "" ' ь
        Case &H44D: piece = "e" ' э -> e
        Case &H44E: piece = "yu" ' ю -> yu
          If isUp Then piece = "YU" Else If cUp Then piece = "Yu"
          o = o & piece: i = i + 1: GoTo NextChar
        Case &H44F: piece = "ya" ' я -> ya
          If isUp Then piece = "YA" Else If cUp Then piece = "Ya"
          o = o & piece: i = i + 1: GoTo NextChar
        Case Else: piece = c
      End Select

      If Len(piece) = 1 Then
        If isUp Or cUp Then
          Select Case piece
            Case ChrW(&HE4): piece = ChrW(&HC4) ' Ä
            Case ChrW(&H11F): piece = ChrW(&H11E) ' Ğ
            Case ChrW(&HF1): piece = ChrW(&HD1) ' Ñ
            Case ChrW(&HF6): piece = ChrW(&HD6) ' Ö
            Case ChrW(&HFC): piece = ChrW(&HDC) ' Ü
            Case ChrW(&HE7): piece = ChrW(&HC7) ' Ç
            Case ChrW(&H15F): piece = ChrW(&H15E) ' Ş
            Case Else: piece = UCase$(piece)
          End Select
        End If
      End If

      o = o & piece
      i = i + 1
    End If
NextChar:
  Loop

  KzKrillToLatinWord = o
End Function

Public Function KzKrillToLatinStr(ByVal s As String) As String
  Dim i As Long, n As Long, c As String, o As String
  Dim wSt As Long, inW As Boolean
  n = Len(s)
  i = 1
  Do While i <= n
    c = Mid$(s, i, 1)
    If KzIsCyrChar(c) Then
      If Not inW Then wSt = i: inW = True
      i = i + 1
    Else
      If inW Then
        o = o & KzKrillToLatinWord(Mid$(s, wSt, i - wSt))
        inW = False
      End If
      o = o & c
      i = i + 1
    End If
  Loop
  If inW Then
    o = o & KzKrillToLatinWord(Mid$(s, wSt))
  End If
  KzKrillToLatinStr = o
End Function

Public Function KzLatinToKrillWord(ByVal w As String) As String
  Dim i As Long, n As Long, c As String, c1 As String
  Dim o As String, lc As String, lc2 As String
  Dim isUp As Boolean, cUp As Boolean
  Dim isStart As Boolean

  n = Len(w)
  If n = 0 Then KzLatinToKrillWord = "": Exit Function

  isUp = (w = UCase$(w) And w <> LCase$(w) And n > 1)
  i = 1

  Do While i <= n
    c = Mid$(w, i, 1)
    c1 = ""
    If i < n Then c1 = Mid$(w, i + 1, 1)
    lc = LCase$(c)
    cUp = (c <> LCase$(c))
    isStart = (i = 1)

    If Len(c1) > 0 Then
      lc2 = lc & LCase$(c1)
      Select Case lc2
        Case "ye"
          If isUp Or cUp Then o = o & ChrW(&H415) Else o = o & ChrW(&H435)
          i = i + 2: GoTo NextLChar
        Case "yo"
          If isUp Or cUp Then o = o & ChrW(&H401) Else o = o & ChrW(&H451)
          i = i + 2: GoTo NextLChar
        Case "yu"
          If isUp Or cUp Then o = o & ChrW(&H42E) Else o = o & ChrW(&H44E)
          i = i + 2: GoTo NextLChar
        Case "ya"
          If isUp Or cUp Then o = o & ChrW(&H42F) Else o = o & ChrW(&H44F)
          i = i + 2: GoTo NextLChar
      End Select
    End If

    Select Case lc
      Case "a": If isUp Or cUp Then o = o & ChrW(&H410) Else o = o & ChrW(&H430)
      Case ChrW(&HE4): If isUp Or cUp Then o = o & ChrW(&H4D8) Else o = o & ChrW(&H4D9) ' ä -> ә
      Case "b": If isUp Or cUp Then o = o & ChrW(&H411) Else o = o & ChrW(&H431)
      Case "v": If isUp Or cUp Then o = o & ChrW(&H412) Else o = o & ChrW(&H432)
      Case "g": If isUp Or cUp Then o = o & ChrW(&H413) Else o = o & ChrW(&H433)
      Case ChrW(&H11F): If isUp Or cUp Then o = o & ChrW(&H492) Else o = o & ChrW(&H493) ' ğ -> ғ
      Case "d": If isUp Or cUp Then o = o & ChrW(&H414) Else o = o & ChrW(&H434)
      Case "e"
        If isStart Then
          If isUp Or cUp Then o = o & ChrW(&H42D) Else o = o & ChrW(&H44D) ' э
        Else
          If isUp Or cUp Then o = o & ChrW(&H415) Else o = o & ChrW(&H435) ' е
        End If
      Case "j": If isUp Or cUp Then o = o & ChrW(&H416) Else o = o & ChrW(&H436)
      Case "z": If isUp Or cUp Then o = o & ChrW(&H417) Else o = o & ChrW(&H437)
      Case "i": If isUp Or cUp Then o = o & ChrW(&H406) Else o = o & ChrW(&H456) ' i -> і
      Case "y": If isUp Or cUp Then o = o & ChrW(&H42B) Else o = o & ChrW(&H44B) ' y -> ы
      Case "k": If isUp Or cUp Then o = o & ChrW(&H41A) Else o = o & ChrW(&H43A)
      Case "q": If isUp Or cUp Then o = o & ChrW(&H49A) Else o = o & ChrW(&H49B) ' q -> қ
      Case "l": If isUp Or cUp Then o = o & ChrW(&H41B) Else o = o & ChrW(&H43B)
      Case "m": If isUp Or cUp Then o = o & ChrW(&H41C) Else o = o & ChrW(&H43C)
      Case "n": If isUp Or cUp Then o = o & ChrW(&H41D) Else o = o & ChrW(&H43D)
      Case ChrW(&HF1): If isUp Or cUp Then o = o & ChrW(&H4A2) Else o = o & ChrW(&H4A3) ' ñ -> ң
      Case "o": If isUp Or cUp Then o = o & ChrW(&H41E) Else o = o & ChrW(&H43E)
      Case ChrW(&HF6): If isUp Or cUp Then o = o & ChrW(&H4E8) Else o = o & ChrW(&H4E9) ' ö -> ө
      Case "p": If isUp Or cUp Then o = o & ChrW(&H41F) Else o = o & ChrW(&H43F)
      Case "r": If isUp Or cUp Then o = o & ChrW(&H420) Else o = o & ChrW(&H440)
      Case "s": If isUp Or cUp Then o = o & ChrW(&H421) Else o = o & ChrW(&H441)
      Case "t": If isUp Or cUp Then o = o & ChrW(&H422) Else o = o & ChrW(&H442)
      Case "u": If isUp Or cUp Then o = o & ChrW(&H4B0) Else o = o & ChrW(&H4B1) ' u -> ұ
      Case ChrW(&HFC): If isUp Or cUp Then o = o & ChrW(&H4AE) Else o = o & ChrW(&H4AF) ' ü -> ү
      Case "w": If isUp Or cUp Then o = o & ChrW(&H423) Else o = o & ChrW(&H443) ' w -> у
      Case "f": If isUp Or cUp Then o = o & ChrW(&H424) Else o = o & ChrW(&H444)
      Case "x": If isUp Or cUp Then o = o & ChrW(&H425) Else o = o & ChrW(&H445)
      Case "h": If isUp Or cUp Then o = o & ChrW(&H4BA) Else o = o & ChrW(&H4BB) ' h -> һ
      Case "c": If isUp Or cUp Then o = o & ChrW(&H426) Else o = o & ChrW(&H446) ' c -> ц
      Case ChrW(&HE7): If isUp Or cUp Then o = o & ChrW(&H427) Else o = o & ChrW(&H447) ' ç -> ч
      Case ChrW(&H15F): If isUp Or cUp Then o = o & ChrW(&H428) Else o = o & ChrW(&H448) ' ş -> ш
      Case "'", ChrW(&H2BC), ChrW(&H2019): o = o & ChrW(&H44A) ' ' -> ъ
      Case Else: o = o & c
    End Select
    i = i + 1

NextLChar:
  Loop

  KzLatinToKrillWord = o
End Function

Public Function KzLatinToKrillStr(ByVal s As String) As String
  Dim i As Long, n As Long, c As String, o As String
  Dim wSt As Long, inW As Boolean, a As Long
  n = Len(s)
  i = 1
  Do While i <= n
    c = Mid$(s, i, 1)
    a = AscW(c)
    If a < 0 Then a = a + 65536

    Dim isLat As Boolean
    isLat = ((a >= 65 And a <= 90) Or (a >= 97 And a <= 122) Or _
             a = &HC4 Or a = &HE4 Or a = &H11E Or a = &H11F Or _
             a = &HD1 Or a = &HF1 Or a = &HD6 Or a = &HF6 Or _
             a = &HDC Or a = &HFC Or a = &HC7 Or a = &HE7 Or _
             a = &H15E Or a = &H15F Or a = 39 Or a = &H2BC Or a = &H2019)

    If isLat Then
      If Not inW Then wSt = i: inW = True
      i = i + 1
    Else
      If inW Then
        o = o & KzLatinToKrillWord(Mid$(s, wSt, i - wSt))
        inW = False
      End If
      o = o & c
      i = i + 1
    End If
  Loop
  If inW Then
    o = o & KzLatinToKrillWord(Mid$(s, wSt))
  End If
  KzLatinToKrillStr = o
End Function

' Word
Public Sub KzWdConvDoc(ByVal toLatin As Boolean)
  Dim rng As Object, t As String, newT As String, changed As Long, wi As Long, wr As Object, words As Object
  On Error GoTo ErrWord
  If Documents.Count = 0 Then Exit Sub
  If Selection.Type <> 1 Then Set rng = Selection.Range Else Set rng = ActiveDocument.Content
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
      If toLatin Then newT = KzKrillToLatinStr(t) Else newT = KzLatinToKrillStr(t)
      If newT <> t Then
        wr.Text = newT
        changed = changed + 1
      End If
    End If
WNextWord:
  Next wi
  Application.ScreenUpdating = True
  Dim modeStr As String
  If toLatin Then modeStr = "Qazaqşa (Kirill -> Latın)" Else modeStr = "Qazaqşa (Latın -> Kirill)"
  MsgBox "Tayyor: " & changed & " ta so'z o'zgardi (" & modeStr & ").", vbInformation, "AlifboOtkazgich"
  Exit Sub
ErrWord:
  Application.ScreenUpdating = True
  MsgBox "Xatolik: " & Err.Description, vbExclamation, "AlifboOtkazgich"
End Sub

Public Sub RbWdKzToLat(control As IRibbonControl)
  KzWdConvDoc True
End Sub

Public Sub RbWdKzToKrill(control As IRibbonControl)
  KzWdConvDoc False
End Sub

' Excel
Private Function KzXlConvArea(ByVal area As Object, ByVal toLatin As Boolean) As Long
  Dim arr As Variant, r As Long, c As Long, nr As Long, nc As Long, v As String, nv As String, n As Long, changed As Boolean
  On Error Resume Next
  If area Is Nothing Then Exit Function
  If area.Cells.Count = 1 Then
    If Not IsError(area.Value) Then
      If Not area.HasFormula Then
        If VarType(area.Value) = vbString Then
          v = CStr(area.Value)
          If Len(v) > 0 Then
            If toLatin Then nv = KzKrillToLatinStr(v) Else nv = KzLatinToKrillStr(v)
            If Err.Number = 0 And Len(nv) > 0 And nv <> v Then
              area.Value = nv: KzXlConvArea = 1
            End If
          End If
        End If
      End If
    End If
    Exit Function
  End If
  arr = area.Value
  If Not IsArray(arr) Then Exit Function
  nr = UBound(arr, 1): nc = UBound(arr, 2)
  For r = 1 To nr
    For c = 1 To nc
      If VarType(arr(r, c)) = vbString Then
        v = arr(r, c)
        If Len(v) > 0 Then
          If toLatin Then nv = KzKrillToLatinStr(v) Else nv = KzLatinToKrillStr(v)
          If Err.Number = 0 And Len(nv) > 0 And nv <> v Then
            arr(r, c) = nv: n = n + 1: changed = True
          End If
        End If
      End If
    Next c
  Next r
  If changed Then area.Value = arr
  KzXlConvArea = n
End Function

Public Sub KzXlRun(ByVal toLatin As Boolean)
  Dim n As Long, rng As Object, txt As Object, ar As Object
  If Workbooks.Count = 0 Then Exit Sub
  Application.ScreenUpdating = False
  On Error GoTo XlFail
  Set rng = ActiveSheet.UsedRange
  If TypeName(Selection) = "Range" Then
    If Selection.Cells.Count > 1 Then Set rng = Selection
  End If
  If rng.Cells.Count = 1 Then
    n = KzXlConvArea(rng, toLatin)
  Else
    Set txt = Nothing
    On Error Resume Next
    Set txt = rng.SpecialCells(2, 2)
    On Error GoTo XlFail
    If txt Is Nothing Then
      n = KzXlConvArea(rng, toLatin)
    Else
      For Each ar In txt.Areas: n = n + KzXlConvArea(ar, toLatin): Next ar
    End If
  End If
  Application.ScreenUpdating = True
  Dim modeStr As String
  If toLatin Then modeStr = "Qazaqşa (Kirill -> Latın)" Else modeStr = "Qazaqşa (Latın -> Kirill)"
  MsgBox "Tayyor: " & n & " ta katak o'zgardi (" & modeStr & ").", vbInformation, "AlifboOtkazgich"
  Exit Sub
XlFail:
  Application.ScreenUpdating = True
  MsgBox "Xatolik: " & Err.Description, vbExclamation, "AlifboOtkazgich"
End Sub

Public Sub RbXlKzToLat(control As IRibbonControl)
  KzXlRun True
End Sub

Public Sub RbXlKzToKrill(control As IRibbonControl)
  KzXlRun False
End Sub
