Attribute VB_Name = "sToolsDocument"
Option Explicit




Sub SwitchView()
  If ActiveDocument Is Nothing Then Exit Sub
  With ActiveDocument.ActiveWindow.ActiveView
    If .Type = cdrEnhancedView Or .Type = cdrPixelView Then
      .Type = cdrSimpleWireframeView
    Else
      If ActiveDocument.Rulers.HUnits = cdrPixel Then .Type = cdrPixelView Else .Type = cdrEnhancedView
    End If
  End With
End Sub




Sub NewWebDocument()
  Dim createopt As StructCreateOptions
  Set createopt = CreateStructCreateOptions
  With createopt
    .Units = cdrPixel
    .PageWidth = 0.048768
    .PageHeight = 0.027432
    .Resolution = 72#
    .ColorContext = CreateColorContext2("sRGB IEC61966-2.1,ISO Coated v2 300% (ECI),Dot Gain 15%", clrRenderPerceptual, clrColorModelRGB)
  End With
  Dim d As Document
  Set d = CreateDocumentEx(createopt)
  With d.ActiveWindow.ActiveView
    .ShowProofColors = False
    .SimulateOverprints = False
    .Type = cdrPixelView
  End With
End Sub



Sub NewCMYKDocument()
  Dim createopt As StructCreateOptions
  Set createopt = CreateStructCreateOptions
  With createopt
    .Units = cdrMillimeter
    .PageWidth = 210#
    .PageHeight = 297#
    .Resolution = 300#
    .ColorContext = CreateColorContext2("sRGB IEC61966-2.1,ISO Coated v2 300% (ECI),Dot Gain 15%", clrRenderPerceptual, clrColorModelCMYK)
  End With
  Dim d As Document
  Set d = CreateDocumentEx(createopt)
  With d.ActiveWindow.ActiveView
    .ShowProofColors = False
    .SimulateOverprints = True
    .Type = cdrEnhancedView
  End With
End Sub






Sub GoToFirstPage()
    If ActiveDocument Is Nothing Then Exit Sub
    If ActiveDocument.Pages.Count > 1 Then ActiveDocument.Pages.First.Activate
End Sub
    
Sub GoToLastPage()
    If ActiveDocument Is Nothing Then Exit Sub
    If ActiveDocument.Pages.Count > 1 Then ActiveDocument.Pages.Last.Activate
End Sub

Sub FitPageToSelect()
    If ActiveDocument Is Nothing Then Exit Sub
    If ActiveShape Is Nothing Then Exit Sub
    Dim s As Shape
    On Error Resume Next
    boostStart "Fit Page To Select"
    Set s = ActiveSelectionRange.Group
    ActivePage.SetSize s.SizeWidth, s.SizeHeight
    s.AlignToPageCenter cdrAlignHCenter + cdrAlignVCenter
    s.Ungroup
    boostFinish endUndoGroup:=True
End Sub

Sub DeleteViewStyles()
    If ActiveDocument Is Nothing Then Exit Sub
    Dim Vws As Views
    Set Vws = ActiveDocument.Views
    While Vws.Count <> 0
        Vws(Vws.Count).Delete
    Wend
    Set Vws = Nothing
End Sub

Sub RemovePagesNames()
    If ActiveDocument Is Nothing Then Exit Sub
    Dim p As Page
    For Each p In ActiveDocument.Pages
        p.Name = " "
    Next
End Sub
