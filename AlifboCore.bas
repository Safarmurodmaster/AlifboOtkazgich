Attribute VB_Name = "AlifboCore"
Option Explicit

Private Function CAp() As String
  CAp = "'" & ChrW(&H2018) & ChrW(&H2019) & ChrW(&H201B) & ChrW(&H2B9) & ChrW(&H2BB) & ChrW(&H2BC) & ChrW(&H2C8) & ChrW(&H2CA) & ChrW(&H2032) & "`" & ChrW(&HB4) & ChrW(-249)
End Function

Public Function CoreIsApos(ByVal c As String) As Boolean
  If Len(c) <> 1 Then Exit Function
  CoreIsApos = (InStr(CAp(), c) > 0)
End Function

Public Function CoreIsLetter(ByVal c As String) As Boolean
  Dim a As Long
  If Len(c) <> 1 Then Exit Function
  a = AscW(c)
  If a < 0 Then a = a + 65536
  Select Case a
    Case 65 To 90, 97 To 122, &HC0 To &H24F
      CoreIsLetter = True
  End Select
End Function

Public Function CoreIsWordCh(ByVal c As String) As Boolean
  CoreIsWordCh = CoreIsLetter(c) Or CoreIsApos(c)
End Function

Public Function CoreWord(ByVal w As String) As String
  Dim i As Long, n As Long, c As String, c1 As String, c2 As String, o As String
  n = Len(w)
  i = 1
  Do While i <= n
    c = Mid$(w, i, 1)
    c1 = Mid$(w, i + 1, 1)
    c2 = Mid$(w, i + 2, 1)
    If (c = "o" Or c = "O" Or c = "g" Or c = "G") And CoreIsApos(c1) And (CoreIsLetter(c2) Or c2 = "") Then
      Select Case c
        Case "o": o = o & ChrW(&HF6)
        Case "O": o = o & ChrW(&HD6)
        Case "g": o = o & ChrW(&H11F)
        Case "G": o = o & ChrW(&H11E)
      End Select
      i = i + 2
    ElseIf CoreIsApos(c) Then
      If i > 1 And i < n Then
        If CoreIsLetter(Mid$(w, i - 1, 1)) And CoreIsLetter(c1) Then
          o = o & ChrW(&H2BC)
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
      If c = "S" Then o = o & ChrW(&H15E) Else o = o & ChrW(&H15F)
      i = i + 2
    ElseIf (c = "c" Or c = "C") And (c1 = "h" Or c1 = "H") Then
      If c = "C" Then o = o & ChrW(&HC7) Else o = o & ChrW(&HE7)
      i = i + 2
    Else
      o = o & c
      i = i + 1
    End If
  Loop
  CoreWord = o
End Function

Public Function CoreOldWord(ByVal w As String) As String
  Dim i As Long, c As String, o As String, up As Boolean
  up = (Len(w) > 1 And w = UCase$(w) And w <> LCase$(w))
  For i = 1 To Len(w)
    c = Mid$(w, i, 1)
    Select Case AscW(c)
      Case &HF6: o = o & "o" & ChrW(&H2BB)
      Case &HD6: o = o & "O" & ChrW(&H2BB)
      Case &H11F: o = o & "g" & ChrW(&H2BB)
      Case &H11E: o = o & "G" & ChrW(&H2BB)
      Case &H15F: o = o & "sh"
      Case &H15E: If up Then o = o & "SH" Else o = o & "Sh"
      Case &HE7: o = o & "ch"
      Case &HC7: If up Then o = o & "CH" Else o = o & "Ch"
      Case Else: o = o & c
    End Select
  Next i
  CoreOldWord = o
End Function

Private Function IsOpenCtx(ByVal p As String) As Boolean
  Select Case p
    Case "", " ", Chr(9), Chr(13), Chr(10), Chr(11), "(", "[", "{", "-", "/", ChrW(160), ChrW(&H2013), ChrW(&H2014), ChrW(&H201C)
      IsOpenCtx = True
  End Select
End Function

Private Function IsSpc(ByVal c As String) As Boolean
  IsSpc = (c = " " Or c = ChrW(160))
End Function

Public Function CoreClean(ByVal s As String, ByVal doSp As Boolean) As String
  Dim B As String, last As String, lowOpen As Boolean, skipSp As Boolean, runLS As Boolean
  Dim i As Long, c As String, c1 As String, act As Long, extra As Long, handled As Boolean
  runLS = True
  i = 1
  Do While i <= Len(s)
    c = Mid$(s, i, 1)
    c1 = Mid$(s, i + 1, 1)
    act = 0
    extra = 0
    handled = False
    If c = ChrW(&H200B) Or c = ChrW(&H200C) Or c = ChrW(&H200D) Or c = ChrW(-257) Or c = ChrW(&HAD) Then handled = True
    If Not handled Then
      If skipSp And IsSpc(c) Then handled = True
    End If
    If Not handled Then
      skipSp = False
      If c = "<" And c1 = "<" Then
        act = 1: extra = 1
      ElseIf c = ">" And c1 = ">" Then
        act = 2: extra = 1
      ElseIf c = ChrW(171) Or c = ChrW(&H201F) Then
        act = 1
      ElseIf c = ChrW(&H201E) Then
        act = 1: lowOpen = True
      ElseIf c = ChrW(187) Then
        act = 2
      ElseIf c = ChrW(&H201C) And lowOpen Then
        act = 2: lowOpen = False
      ElseIf c = ChrW(&H2018) And c1 = ChrW(&H2018) Then
        act = 1: extra = 1
      ElseIf c = ChrW(&H2019) And c1 = ChrW(&H2019) Then
        act = 2: extra = 1
      ElseIf c = Chr(34) Then
        If IsOpenCtx(last) Then act = 1 Else act = 2
      End If
      If act = 1 Then
        B = B & ChrW(&H201C)
        last = ChrW(&H201C)
        skipSp = True
        i = i + extra
      ElseIf act = 2 Then
        Do While Len(B) > 0
          If Not IsSpc(Right$(B, 1)) Then Exit Do
          B = Left$(B, Len(B) - 1)
        Loop
        B = B & ChrW(&H201D)
        last = ChrW(&H201D)
        i = i + extra
      Else
        If c = " " And doSp Then
          If last = " " Then
            If Not runLS Then handled = True
          Else
            runLS = (last = "" Or last = Chr(10) Or last = Chr(13))
          End If
        End If
        If Not handled Then
          B = B & c
          last = c
        End If
      End If
    End If
    i = i + 1
  Loop
  CoreClean = B
End Function

Public Function CoreLoadExc() As Collection
  Dim c As New Collection, ln As String, ff As Integer, p As String
  p = Environ("APPDATA") & "\Alifbo\istisnolar.txt"
  On Error Resume Next
  If Dir(p) <> "" Then
    ff = FreeFile
    Open p For Input As #ff
    Do While Not EOF(ff)
      Line Input #ff, ln
      ln = Trim$(ln)
      If Len(ln) > 0 Then
        If Left$(ln, 1) <> "#" Then c.Add LCase$(ln)
      End If
    Loop
    Close #ff
  End If
  Set CoreLoadExc = c
End Function

Public Function CoreInExc(ByVal w As String, ByVal exc As Object) As Boolean
  Dim i As Long
  If exc Is Nothing Then Exit Function
  For i = 1 To exc.Count
    If exc(i) = LCase$(w) Then
      CoreInExc = True
      Exit Function
    End If
  Next i
End Function

Public Function CoreConvert(ByVal s As String, ByVal toNewAlpha As Boolean, ByVal doSp As Boolean, ByVal exc As Object) As String
  Dim t As String, i As Long, st As Long, o As String, w As String
  t = CoreClean(s, doSp)
  i = 1
  Do While i <= Len(t)
    If CoreIsWordCh(Mid$(t, i, 1)) Then
      st = i
      Do While i <= Len(t)
        If Not CoreIsWordCh(Mid$(t, i, 1)) Then Exit Do
        i = i + 1
      Loop
      w = Mid$(t, st, i - st)
      If Len(w) >= 2 Then
        If Not CoreInExc(w, exc) Then
          If toNewAlpha Then w = CoreWord(w) Else w = CoreOldWord(w)
        End If
      End If
      o = o & w
    Else
      o = o & Mid$(t, i, 1)
      i = i + 1
    End If
  Loop
  CoreConvert = o
End Function
