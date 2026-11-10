Attribute VB_Name = "AlifboPpt"
Option Explicit

Private pCount As Long

Private Sub PptRunText(ByVal tr As Object, ByVal toNewAlpha As Boolean, ByVal exc As Object)
  Dim i As Long, r As Object, t As String, nt As String, rc As Long
  On Error Resume Next
  If tr Is Nothing Then Exit Sub
  rc = tr.Runs.Count
  If rc <= 1 Then
    t = tr.Text
    If Len(t) > 0 Then
      nt = CoreConvert(t, toNewAlpha, False, exc)
      If Err.Number = 0 And Len(nt) > 0 And nt <> t Then
        tr.Text = nt
        pCount = pCount + 1
      End If
    End If
    Exit Sub
  End If
  For i = 1 To rc
    Set r = tr.Runs(i)
    t = r.Text
    If Len(t) > 0 Then
      nt = CoreConvert(t, toNewAlpha, False, exc)
      If Err.Number = 0 And Len(nt) > 0 And nt <> t Then
        r.Text = nt
        pCount = pCount + 1
      End If
    End If
  Next i
End Sub

Private Sub PptShape(ByVal shp As Object, ByVal toNewAlpha As Boolean, ByVal exc As Object)
  Dim i As Long, rr As Long, cc As Long
  On Error Resume Next
  If shp.Type = 6 Then
    For i = 1 To shp.GroupItems.Count
      PptShape shp.GroupItems(i), toNewAlpha, exc
    Next i
    Exit Sub
  End If
  If shp.HasTable Then
    For rr = 1 To shp.Table.Rows.Count
      For cc = 1 To shp.Table.Columns.Count
        PptRunText shp.Table.Cell(rr, cc).Shape.TextFrame.TextRange, toNewAlpha, exc
      Next cc
    Next rr
  End If
  If shp.HasTextFrame Then
    If shp.TextFrame.HasText Then PptRunText shp.TextFrame.TextRange, toNewAlpha, exc
  End If
End Sub

Private Sub PptRun(ByVal toNewAlpha As Boolean)
  Dim exc As Collection, sld As Object, shp As Object
  If Not Gate() Then Exit Sub
  If Presentations.Count = 0 Then Exit Sub
  Set exc = CoreLoadExc()
  pCount = 0
  On Error GoTo Fail
  For Each sld In ActivePresentation.Slides
    For Each shp In sld.Shapes
      PptShape shp, toNewAlpha, exc
    Next shp
    On Error Resume Next
    For Each shp In sld.NotesPage.Shapes
      PptShape shp, toNewAlpha, exc
    Next shp
    On Error GoTo Fail
  Next sld
  MsgBox "Tayyor: " & pCount & " ta matn bo'lagi o'zgardi." & StatusSuffix(), vbInformation, ProdName()
  Exit Sub
Fail:
  MsgBox "Xatolik: " & Err.Description, vbExclamation, ProdName()
End Sub

Public Sub PptToNew()
  PptRun True
End Sub

Public Sub PptToOld()
  PptRun False
End Sub

Private Sub PptBtn(ByVal cb As Object, ByVal cap As String, ByVal act As String, ByVal tip As String)
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
    PptBtn cb, "Yangi alifbo", "PptToNew", "Butun taqdimot yangi alifboga o'tadi"
    PptBtn cb, "Eski alifbo", "PptToOld", "Yangi alifbodan eskisiga qaytaradi"
    PptBtn cb, "Fayllar", "OpenToolsApp", "Fayllarni (Word, Excel, PowerPoint) yangi alifboga o'tkazish oynasi"
    PptBtn cb, "Litsenziya", "LicenseInfo", "Sinov muddati va litsenziya kodi"
    PptBtn cb, "Muallif", "AlifboAbout", "Dastur muallifi va aloqa"
    cb.Visible = True
  End If
  d = LeftMinutes()
  Err.Clear
End Sub

Public Sub RbPptNew(control As IRibbonControl)
  PptRun True
End Sub

Public Sub RbPptOld(control As IRibbonControl)
  PptRun False
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

Public Sub RbPptKrillNew(control As IRibbonControl)
  KrillPptRun True
End Sub

Public Sub RbPptKrillOld(control As IRibbonControl)
  KrillPptRun False
End Sub

' ==============================================================
' PowerPoint uchun: matnlarni kirill -> yangi / lotin
' ==============================================================
Public Sub KrillPptRun(ByVal toNew As Boolean)
  Dim exc As Collection, sld As Object, shp As Object, cnt As Long
  If Not Gate() Then Exit Sub
  If Presentations.Count = 0 Then Exit Sub
  Set exc = CoreLoadExc()
  cnt = 0
  On Error GoTo PptFail
  For Each sld In ActivePresentation.Slides
    For Each shp In sld.Shapes
      KrillPptShape shp, toNew, cnt
    Next shp
    On Error Resume Next
    For Each shp In sld.NotesPage.Shapes
      KrillPptShape shp, toNew, cnt
    Next shp
    On Error GoTo PptFail
  Next sld
  Dim modeName As String
  If toNew Then modeName = "Yangi alifbo" Else modeName = "Lotin alifbo"
  MsgBox "Tayyor (" & modeName & "): " & cnt & " ta matn bo'lagi o'zgardi." & StatusSuffix(), vbInformation, ProdName()
  Exit Sub
PptFail:
  MsgBox "Xatolik: " & Err.Description, vbExclamation, ProdName()
End Sub

Private Sub KrillPptShape(ByVal shp As Object, ByVal toNew As Boolean, ByRef cnt As Long)
  Dim i As Long, rr As Long, cc As Long
  On Error Resume Next
  If shp.Type = 6 Then
    For i = 1 To shp.GroupItems.Count
      KrillPptShape shp.GroupItems(i), toNew, cnt
    Next i
    Exit Sub
  End If
  If shp.HasTable Then
    For rr = 1 To shp.Table.Rows.Count
      For cc = 1 To shp.Table.Columns.Count
        KrillConvPptText shp.Table.Cell(rr, cc).Shape.TextFrame.TextRange, toNew, cnt
      Next cc
    Next rr
  End If
  If shp.HasTextFrame Then
    If shp.TextFrame.HasText Then KrillConvPptText shp.TextFrame.TextRange, toNew, cnt
  End If
End Sub

Private Sub KrillConvPptText(ByVal tr As Object, ByVal toNew As Boolean, ByRef cnt As Long)
  Dim i As Long, r As Object, t As String, nt As String, rc As Long
  On Error Resume Next
  If tr Is Nothing Then Exit Sub
  rc = tr.Runs.Count
  If rc <= 1 Then
    t = tr.Text
    If Len(t) > 0 And IsCyrInStr(t) Then
      nt = KrillConvStr(t, toNew)
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
    If Len(t) > 0 And IsCyrInStr(t) Then
      nt = KrillConvStr(t, toNew)
      If Err.Number = 0 And nt <> t Then
        r.Text = nt
        cnt = cnt + 1
      End If
    End If
  Next i
End Sub