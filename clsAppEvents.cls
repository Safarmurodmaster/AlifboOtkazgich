VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Word.Application

Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
  ' OÖ, SŞ, GĞ, CÇ xatolari kelib chiqmasligi uchun yozish vaqtida aralashilmaydi
End Sub