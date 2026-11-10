Attribute VB_Name = "AlifboKrill"
Option Explicit

' ============================================================
' AlifboKrill.bas  —  Kirill → Yangi / Lotin sof mantiq moduli
' Muallif: Fozliddinov Safarmurod | @Safar_master | +998939530473
' (Office ilovalariga bog'liq bo'lmagan mustaqil funksiyalar)
' ============================================================

Public Function IsCyrChar(ByVal c As String) As Boolean
  Dim a As Long
  If Len(c) <> 1 Then Exit Function
  a = AscW(c)
  If a < 0 Then a = a + 65536
  IsCyrChar = (a >= &H400 And a <= &H52F)
End Function

Public Function IsCyrVowel(ByVal c As String) As Boolean
  Select Case c
    Case ChrW(&H410), ChrW(&H430) ' А а
      IsCyrVowel = True
    Case ChrW(&H415), ChrW(&H435) ' Е е
      IsCyrVowel = True
    Case ChrW(&H401), ChrW(&H451) ' Ё ё
      IsCyrVowel = True
    Case ChrW(&H418), ChrW(&H438) ' И и
      IsCyrVowel = True
    Case ChrW(&H41E), ChrW(&H43E) ' О о
      IsCyrVowel = True
    Case ChrW(&H423), ChrW(&H443) ' У у
      IsCyrVowel = True
    Case ChrW(&H42B), ChrW(&H44B) ' Ы ы
      IsCyrVowel = True
    Case ChrW(&H42D), ChrW(&H44D) ' Э э
      IsCyrVowel = True
    Case ChrW(&H42E), ChrW(&H44E) ' Ю ю
      IsCyrVowel = True
    Case ChrW(&H42F), ChrW(&H44F) ' Я я
      IsCyrVowel = True
    Case ChrW(&H40E), ChrW(&H45E) ' Ў ў
      IsCyrVowel = True
  End Select
End Function

Public Function NormHomoglyphs(ByVal s As String) As String
  Dim i As Long, c As String, o As String
  o = s
  For i = 1 To Len(o)
    c = Mid$(o, i, 1)
    Select Case c
      Case "a": c = ChrW(&H430)  ' a → а
      Case "A": c = ChrW(&H410)  ' A → А
      Case "o": c = ChrW(&H43E)  ' o → о
      Case "O": c = ChrW(&H41E)  ' O → О
      Case "e": c = ChrW(&H435)  ' e → е
      Case "E": c = ChrW(&H415)  ' E → Е
      Case "c": c = ChrW(&H441)  ' c → с
      Case "C": c = ChrW(&H421)  ' C → С
      Case "p": c = ChrW(&H440)  ' p → р
      Case "P": c = ChrW(&H420)  ' P → Р
      Case "x": c = ChrW(&H445)  ' x → х
      Case "X": c = ChrW(&H425)  ' X → Х
      Case "y": c = ChrW(&H443)  ' y → у
      Case "k": c = ChrW(&H43A)  ' k → к
      Case "K": c = ChrW(&H41A)  ' K → К
      Case "H": c = ChrW(&H41D)  ' H → Н
      Case "B": c = ChrW(&H412)  ' B → В
      Case "M": c = ChrW(&H41C)  ' M → М
      Case "T": c = ChrW(&H422)  ' T → Т
    End Select
    Mid$(o, i, 1) = c
  Next i
  NormHomoglyphs = o
End Function

Public Function CyrException(ByVal w As String, ByVal toNew As Boolean) As String
  Dim wl As String
  wl = LCase$(w)
  If toNew Then
    Select Case wl
      Case ChrW(&H441) & ChrW(&H44A) & ChrW(&H435) & ChrW(&H437) & ChrW(&H434) ' съезд
        CyrException = "s" & ChrW(&H2BC) & "ezd"
      Case ChrW(&H43E) & ChrW(&H431) & ChrW(&H44A) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H442) ' объект
        CyrException = "ob" & ChrW(&H2BC) & "ekt"
      Case ChrW(&H441) & ChrW(&H443) & ChrW(&H431) & ChrW(&H44A) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H442) ' субъект
        CyrException = "sub" & ChrW(&H2BC) & "ekt"
      Case ChrW(&H440) & ChrW(&H430) & ChrW(&H437) & ChrW(&H44A) & ChrW(&H435) & ChrW(&H437) & ChrW(&H434) ' разъезд
        CyrException = "raz" & ChrW(&H2BC) & "ezd"
      Case ChrW(&H43F) & ChrW(&H43E) & ChrW(&H434) & ChrW(&H44A) & ChrW(&H435) & ChrW(&H437) & ChrW(&H434) ' подъезд
        CyrException = "pod" & ChrW(&H2BC) & "ezd"
      Case ChrW(&H438) & ChrW(&H43D) & ChrW(&H44A) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F) ' инъекция
        CyrException = "in" & ChrW(&H2BC) & "eks" & ChrW(&H15F) & "ya"
    End Select
  Else
    Select Case wl
      Case ChrW(&H441) & ChrW(&H44A) & ChrW(&H435) & ChrW(&H437) & ChrW(&H434)
        CyrException = "s'ezd"
      Case ChrW(&H43E) & ChrW(&H431) & ChrW(&H44A) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H442)
        CyrException = "ob'ekt"
      Case ChrW(&H441) & ChrW(&H443) & ChrW(&H431) & ChrW(&H44A) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H442)
        CyrException = "sub'ekt"
      Case ChrW(&H440) & ChrW(&H430) & ChrW(&H437) & ChrW(&H44A) & ChrW(&H435) & ChrW(&H437) & ChrW(&H434)
        CyrException = "raz'ezd"
      Case ChrW(&H43F) & ChrW(&H43E) & ChrW(&H434) & ChrW(&H44A) & ChrW(&H435) & ChrW(&H437) & ChrW(&H434)
        CyrException = "pod'ezd"
      Case ChrW(&H438) & ChrW(&H43D) & ChrW(&H44A) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
        CyrException = "in'ektsiya"
    End Select
  End If
End Function

Public Function CyrTypo(ByVal w As String, ByVal toNew As Boolean) As String
  Dim wl As String, corrCyr As String
  wl = w
  corrCyr = ""
  Select Case LCase$(wl)
    Case ChrW(&H445) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H43C) & ChrW(&H430) ' хамма
      corrCyr = ChrW(&H4B3) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H43C) & ChrW(&H430) ' ҳамма
    Case ChrW(&H445) & ChrW(&H43E) & ChrW(&H437) & ChrW(&H438) & ChrW(&H440) ' хозир
      corrCyr = ChrW(&H4B3) & ChrW(&H43E) & ChrW(&H437) & ChrW(&H438) & ChrW(&H440) ' ҳозир
    Case ChrW(&H440) & ChrW(&H430) & ChrW(&H445) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H442) ' рахмат
      corrCyr = ChrW(&H440) & ChrW(&H430) & ChrW(&H4B3) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H442) ' раҳмат
    Case ChrW(&H445) & ChrW(&H430) & ChrW(&H44F) & ChrW(&H442) ' хаят
      corrCyr = ChrW(&H4B3) & ChrW(&H430) & ChrW(&H451) & ChrW(&H442) ' ҳаёт
    Case ChrW(&H445) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H441) & ChrW(&H438) ' хаммаси
      corrCyr = ChrW(&H4B3) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H441) & ChrW(&H438)
    Case ChrW(&H445) & ChrW(&H443) & ChrW(&H440) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H442) ' хурмат
      corrCyr = ChrW(&H4B3) & ChrW(&H443) & ChrW(&H440) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H442) ' ҳурмат
    Case ChrW(&H445) & ChrW(&H430) & ChrW(&H44F) & ChrW(&H43E) & ChrW(&H442) ' хаёт
      corrCyr = ChrW(&H4B3) & ChrW(&H430) & ChrW(&H451) & ChrW(&H442)
  End Select
  If Len(corrCyr) > 0 Then
    If w = UCase$(w) Then
      CyrTypo = UCase$(KrillConvWord(corrCyr, toNew))
    ElseIf w = StrConv(w, vbProperCase) Then
      Dim tmp As String
      tmp = KrillConvWord(corrCyr, toNew)
      CyrTypo = UCase$(Left$(tmp, 1)) & Mid$(tmp, 2)
    Else
      CyrTypo = KrillConvWord(corrCyr, toNew)
    End If
  End If
End Function

Public Function KrillConvWord(ByVal w As String, ByVal toNew As Boolean) As String
  Dim i As Long, n As Long, c As String, c1 As String
  Dim o As String, lc As String
  Dim isUp As Boolean, cUp As Boolean
  Dim ePrev As String, eIsWordStart As Boolean, eAfterVowel As Boolean
  Dim piece As String

  n = Len(w)
  If n = 0 Then KrillConvWord = "": Exit Function

  isUp = (w = UCase$(w) And w <> LCase$(w) And n > 1)
  i = 1

  Do While i <= n
    c = Mid$(w, i, 1)
    c1 = Mid$(w, i + 1, 1)

    If Not IsCyrChar(c) Then
      o = o & c
      i = i + 1
    Else
      lc = LCase$(c)
      cUp = (c <> LCase$(c))
      piece = ""

      Select Case AscW(lc)
        Case &H430: piece = "a"
        Case &H431: piece = "b"
        Case &H432: piece = "v"
        Case &H433: piece = "g"
        Case &H434: piece = "d"

        Case &H435 ' е
          If i = 1 Then ePrev = "" Else ePrev = Mid$(w, i - 1, 1)
          eIsWordStart = (i = 1) Or (Not IsCyrChar(ePrev))
          eAfterVowel = IsCyrVowel(ePrev) And IsCyrChar(ePrev)
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

        Case &H451 ' ё
          piece = "yo"
          If isUp Then piece = "YO" Else If cUp Then piece = "Yo"
          o = o & piece
          i = i + 1
          GoTo NextChar

        Case &H436: piece = "j"
        Case &H437: piece = "z"
        Case &H438: piece = "i"
        Case &H439: piece = "y"
        Case &H43A: piece = "k"
        Case &H43B: piece = "l"
        Case &H43C: piece = "m"
        Case &H43D: piece = "n"
        Case &H43E: piece = "o"
        Case &H43F: piece = "p"
        Case &H440: piece = "r"
        Case &H441: piece = "s"
        Case &H442: piece = "t"
        Case &H443: piece = "u"
        Case &H444: piece = "f"
        Case &H445: piece = "x"

        Case &H446 ' ц
          If i = 1 Or Not IsCyrChar(Mid$(w, i - 1, 1)) Then
            piece = "s"
          Else
            piece = "ts"
            If isUp Then piece = "TS" Else If cUp Then piece = "Ts"
            o = o & piece
            i = i + 1
            GoTo NextChar
          End If

        Case &H447 ' ч
          If toNew Then
            piece = ChrW(&HE7)
            If isUp Or cUp Then piece = ChrW(&HC7)
          Else
            piece = "ch"
            If isUp Then piece = "CH" Else If cUp Then piece = "Ch"
            o = o & piece
            i = i + 1
            GoTo NextChar
          End If

        Case &H448, &H449 ' ш, щ
          If toNew Then
            piece = ChrW(&H15F)
            If isUp Or cUp Then piece = ChrW(&H15E)
          Else
            piece = "sh"
            If isUp Then piece = "SH" Else If cUp Then piece = "Sh"
            o = o & piece
            i = i + 1
            GoTo NextChar
          End If

        Case &H44A ' ъ
          If toNew Then piece = ChrW(&H2BC) Else piece = "'"

        Case &H44B: piece = "i"

        Case &H44C ' ь
          Dim ьNext As Long
          ьNext = AscW(LCase$(c1))
          If ьNext = &H43E Then
            piece = "yo"
            If isUp Then piece = "YO" Else If cUp Then piece = "Yo"
            o = o & piece
            i = i + 2
            GoTo NextChar
          ElseIf ьNext = &H435 Then
            piece = "ye"
            If isUp Then piece = "YE" Else If cUp Then piece = "Ye"
            o = o & piece
            i = i + 2
            GoTo NextChar
          ElseIf ьNext = &H44E Then
            piece = "yu"
            If isUp Then piece = "YU" Else If cUp Then piece = "Yu"
            o = o & piece
            i = i + 2
            GoTo NextChar
          ElseIf ьNext = &H44F Then
            piece = "ya"
            If isUp Then piece = "YA" Else If cUp Then piece = "Ya"
            o = o & piece
            i = i + 2
            GoTo NextChar
          Else
            piece = ""
          End If

        Case &H44D: piece = "e"

        Case &H44E ' ю
          piece = "yu"
          If isUp Then piece = "YU" Else If cUp Then piece = "Yu"
          o = o & piece
          i = i + 1
          GoTo NextChar

        Case &H44F ' я
          piece = "ya"
          If isUp Then piece = "YA" Else If cUp Then piece = "Ya"
          o = o & piece
          i = i + 1
          GoTo NextChar

        Case &H45E ' ў
          If toNew Then
            piece = ChrW(&HF6)
            If isUp Or cUp Then piece = ChrW(&HD6)
          Else
            piece = "o" & ChrW(&H2BC)
            If isUp Or cUp Then piece = "O" & ChrW(&H2BC)
            o = o & piece
            i = i + 1
            GoTo NextChar
          End If

        Case &H493 ' ғ
          If toNew Then
            piece = ChrW(&H11F)
            If isUp Or cUp Then piece = ChrW(&H11E)
          Else
            piece = "g" & ChrW(&H2BC)
            If isUp Or cUp Then piece = "G" & ChrW(&H2BC)
            o = o & piece
            i = i + 1
            GoTo NextChar
          End If

        Case &H49B: piece = "q" ' қ
        Case &H4B3: piece = "h" ' ҳ

        Case &H400 To &H42F
          Select Case AscW(c)
            Case &H40E: ' Ў
              If toNew Then
                piece = ChrW(&HD6)
              Else
                piece = "O" & ChrW(&H2BC)
                o = o & piece
                i = i + 1
                GoTo NextChar
              End If
            Case &H492: ' Ғ
              If toNew Then
                piece = ChrW(&H11E)
              Else
                piece = "G" & ChrW(&H2BC)
                o = o & piece
                i = i + 1
                GoTo NextChar
              End If
            Case &H49A: piece = "Q" ' Қ
            Case &H4B2: piece = "H" ' Ҳ
            Case Else: piece = c
          End Select

        Case Else
          piece = c
      End Select

      If Len(piece) = 1 Then
        If isUp Or cUp Then piece = UCase$(piece)
      End If

      o = o & piece
      i = i + 1
    End If
NextChar:
  Loop

  KrillConvWord = o
End Function

Public Function IsCyrInStr(ByVal s As String) As Boolean
  Dim i As Long
  For i = 1 To Len(s)
    If IsCyrChar(Mid$(s, i, 1)) Then
      IsCyrInStr = True
      Exit Function
    End If
  Next i
End Function

Public Function KrillConvStr(ByVal s As String, ByVal toNew As Boolean) As String
  Dim i As Long, n As Long, c As String, o As String
  Dim wSt As Long, inW As Boolean
  n = Len(s)
  i = 1
  Do While i <= n
    c = Mid$(s, i, 1)
    If IsCyrChar(c) Then
      If Not inW Then wSt = i: inW = True
      i = i + 1
    Else
      If inW Then
        Dim wTxt As String
        wTxt = Mid$(s, wSt, i - wSt)
        wTxt = NormHomoglyphs(wTxt)
        Dim exRes As String
        exRes = CyrException(wTxt, toNew)
        If Len(exRes) > 0 Then
          o = o & exRes
        Else
          Dim tyRes As String
          tyRes = CyrTypo(wTxt, toNew)
          If Len(tyRes) > 0 Then o = o & tyRes Else o = o & KrillConvWord(wTxt, toNew)
        End If
        inW = False
      End If
      o = o & c
      i = i + 1
    End If
  Loop
  If inW Then
    Dim wTxt2 As String
    wTxt2 = Mid$(s, wSt)
    wTxt2 = NormHomoglyphs(wTxt2)
    Dim exRes2 As String
    exRes2 = CyrException(wTxt2, toNew)
    If Len(exRes2) > 0 Then
      o = o & exRes2
    Else
      Dim tyRes2 As String
      tyRes2 = CyrTypo(wTxt2, toNew)
      If Len(tyRes2) > 0 Then o = o & tyRes2 Else o = o & KrillConvWord(wTxt2, toNew)
    End If
  End If
  KrillConvStr = o
End Function
