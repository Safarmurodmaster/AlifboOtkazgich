Attribute VB_Name = "AlifboXl"
Option Explicit

Private Function XlConvArea(ByVal area As Object, ByVal toNewAlpha As Boolean, ByVal exc As Object) As Long
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
            nv = CoreConvert(v, toNewAlpha, False, exc)
            If Err.Number = 0 And Len(nv) > 0 And nv <> v Then
              area.Value = nv
              XlConvArea = 1
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
          nv = CoreConvert(v, toNewAlpha, False, exc)
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
  XlConvArea = n
End Function

Private Function XlConvRange(ByVal rng As Object, ByVal toNewAlpha As Boolean, ByVal exc As Object) As Long
  Dim txt As Object, ar As Object, n As Long
  On Error Resume Next
  If rng Is Nothing Then Exit Function
  If rng.Cells.Count = 1 Then
    XlConvRange = XlConvArea(rng, toNewAlpha, exc)
    Exit Function
  End If
  Set txt = rng.SpecialCells(2, 2)
  If txt Is Nothing Then Set txt = rng
  For Each ar In txt.Areas
    n = n + XlConvArea(ar, toNewAlpha, exc)
  Next ar
  XlConvRange = n
End Function

Private Sub XlRun(ByVal toNewAlpha As Boolean, ByVal wholeBook As Boolean)
  Dim exc As Collection, n As Long, ws As Object, rng As Object
  If Not Gate() Then Exit Sub
  If Workbooks.Count = 0 Then Exit Sub
  Set exc = CoreLoadExc()
  Application.ScreenUpdating = False
  On Error GoTo Fail
  If wholeBook Then
    For Each ws In ActiveWorkbook.Worksheets
      n = n + XlConvRange(ws.UsedRange, toNewAlpha, exc)
    Next ws
  Else
    Set rng = ActiveSheet.UsedRange
    If TypeName(Selection) = "Range" Then
      If Selection.Cells.Count > 1 Then Set rng = Selection
    End If
    n = XlConvRange(rng, toNewAlpha, exc)
  End If
  Application.ScreenUpdating = True
  MsgBox "Tayyor: " & n & " ta katak o'zgardi." & StatusSuffix(), vbInformation, ProdName()
  Exit Sub
Fail:
  Application.ScreenUpdating = True
  MsgBox "Xatolik: " & Err.Description, vbExclamation, ProdName()
End Sub

Public Sub XlToNew()
  XlRun True, False
End Sub

Public Sub XlToOld()
  XlRun False, False
End Sub

Public Sub XlBookNew()
  XlRun True, True
End Sub

Private Sub XlBtn(ByVal cb As Object, ByVal cap As String, ByVal act As String, ByVal tip As String)
  Dim b As Object
  Set b = cb.Controls.Add(1, , , , True)
  b.Caption = cap
  b.OnAction = act
  b.Style = 2
  b.TooltipText = tip
End Sub

Public Sub Auto_Open()
  Dim cb As Object, d As Long
  On Error Resume Next
  Application.CommandBars("Alifbo").Delete
  Set cb = Application.CommandBars.Add("Alifbo", 1, False, True)
  If Not cb Is Nothing Then
    XlBtn cb, "Yangi alifbo", "XlToNew", "Belgilangan kataklar (yoki butun varaq) yangi alifboga o'tadi"
    XlBtn cb, "Eski alifbo", "XlToOld", "Yangi alifbodan eskisiga qaytaradi"
    XlBtn cb, "Butun kitob", "XlBookNew", "Hamma varaqlarni yangi alifboga o'tkazadi"
    XlBtn cb, "Fayllar", "OpenToolsApp", "Fayllarni (Word, Excel, PowerPoint) yangi alifboga o'tkazish oynasi"
    XlBtn cb, "Litsenziya", "LicenseInfo", "Sinov muddati va litsenziya kodi"
    XlBtn cb, "Muallif", "AlifboAbout", "Dastur muallifi va aloqa"
    cb.Visible = True
  End If
  d = LeftMinutes()
  Err.Clear
End Sub

Public Sub RbXlNew(control As IRibbonControl)
  XlRun True, False
End Sub

Public Sub RbXlOld(control As IRibbonControl)
  XlRun False, False
End Sub

Public Sub RbXlBook(control As IRibbonControl)
  XlRun True, True
End Sub

Public Sub RbTools(control As IRibbonControl)
  OpenToolsApp
End Sub

Public Sub RbLic(control As IRibbonControl)
  LicenseInfo
End Sub

Public Sub RbAbout(control As IRibbonControl)
  AlifboAbout
End Sub

Public Sub RbXlKrillNew(control As IRibbonControl)
  KrillConvXl True
End Sub

Public Sub RbXlKrillOld(control As IRibbonControl)
  KrillConvXl False
End Sub

' ==============================================================
' Excel uchun: kataklar bo'ylab kirill matn o'tkazish
' ==============================================================
Public Sub KrillConvXl(ByVal toNew As Boolean)
  Dim exc As Collection, n As Long, rng As Object
  Dim arr As Variant, r As Long, c As Long, v As String, nv As String
  Dim changed As Boolean

  If Not Gate() Then Exit Sub
  If Workbooks.Count = 0 Then Exit Sub
  Set exc = CoreLoadExc()
  Application.ScreenUpdating = False
  On Error GoTo XlFail

  Set rng = ActiveSheet.UsedRange
  If TypeName(Selection) = "Range" Then
    If Selection.Cells.Count > 1 Then Set rng = Selection
  End If

  If rng.Cells.Count = 1 Then
    If Not rng.HasFormula Then
      v = CStr(rng.Value)
      If IsCyrInStr(v) Then
        nv = KrillConvStr(v, toNew)
        If nv <> v Then rng.Value = nv: n = 1
      End If
    End If
  Else
    arr = rng.Value
    If IsArray(arr) Then
      For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
          If VarType(arr(r, c)) = vbString Then
            v = arr(r, c)
            If IsCyrInStr(v) Then
              nv = KrillConvStr(v, toNew)
              If nv <> v Then arr(r, c) = nv: n = n + 1: changed = True
            End If
          End If
        Next c
      Next r
      If changed Then rng.Value = arr
    End If
  End If

  Application.ScreenUpdating = True
  Dim modeName As String
  If toNew Then modeName = "Yangi alifbo" Else modeName = "Lotin alifbo"
  MsgBox "Tayyor (" & modeName & "): " & n & " ta katak o'zgardi." & StatusSuffix(), vbInformation, ProdName()
  Exit Sub
XlFail:
  Application.ScreenUpdating = True
  MsgBox "Xatolik: " & Err.Description, vbExclamation, ProdName()
End Sub