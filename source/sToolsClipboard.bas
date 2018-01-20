Attribute VB_Name = "sToolsClipboard"
Sub Clear()
    If Clipboard.Empty = False Then Clipboard.Clear
End Sub


Sub PasteInCenter()
    If ActiveDocument Is Nothing Then Exit Sub
    Dim Doc As Document, old_refpoint As Long
    On Error GoTo myEnd
    boostStart "Paste in Center"
    Set Doc = ActiveDocument
    old_refpoint = Doc.ReferencePoint
    Doc.ReferencePoint = cdrCenter
    ActiveLayer.Paste
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    sr.SetPosition ActiveWindow.ActiveView.OriginX, ActiveWindow.ActiveView.OriginY
    Doc.ReferencePoint = old_refpoint
    boostFinish endUndoGroup:=True
    Exit Sub
myEnd:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Error"
    boostFinish endUndoGroup:=True
    End Sub



Sub PasteAsBitmap(): pastAsXXX "Bitmap": End Sub
Sub PasteAsMetafile(): pastAsXXX "Metafile": End Sub
Sub PasteAsCorel32(): pastAsXXX "Corel 32-bit Presentation Exchange Date": End Sub
Sub PasteAsText(): pastAsXXX "Text": End Sub
Private Sub pastAsXXX(s$)
    If ActiveDocument Is Nothing Then Exit Sub
    Dim Doc As Document, old_refpoint As Long
    On Error GoTo myEnd
    boostStart "Paste as " & s
    Set Doc = ActiveDocument
    old_refpoint = Doc.ReferencePoint
    Doc.ReferencePoint = cdrCenter
    ActiveLayer.PasteSpecial s, False, False
    ActiveShape.SetPosition ActiveWindow.ActiveView.OriginX, ActiveWindow.ActiveView.OriginY
    Doc.ReferencePoint = old_refpoint
    boostFinish endUndoGroup:=True
    Exit Sub
myEnd:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Error"
    boostFinish endUndoGroup:=True
    End Sub
    
    


Sub PasteOrderBack()
    If ActiveDocument Is Nothing Then Exit Sub
    Dim sr As ShapeRange, s As Shape
    Set sr = ActiveSelectionRange
    If sr.Shapes.Count <> 1 Then Exit Sub
    ActiveDocument.ClearSelection
    Set s = sr(1).Layer.Paste
    'Set s = sr(1).Duplicate(0, 0)
    s.OrderBackOf sr(1)
    s.AddToSelection
End Sub

Sub PasteOrderFront()
    If ActiveDocument Is Nothing Then Exit Sub
    Dim sr As ShapeRange, s As Shape
    Set sr = ActiveSelectionRange
    If sr.Shapes.Count <> 1 Then Exit Sub
    ActiveDocument.ClearSelection
    Set s = sr(1).Layer.Paste
    s.OrderFrontOf sr(1)
    s.AddToSelection
End Sub

