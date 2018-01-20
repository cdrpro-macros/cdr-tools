Attribute VB_Name = "sToolsOutline"
Option Explicit

Sub MakeOutlineSameAsFill()
    Dim sr As ShapeRange, sh As Shape
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then Beep: Exit Sub
    On Error Resume Next
    boostStart "Make outline same as fill"
    For Each sh In sr.Shapes.FindShapes
       If sh.CanHaveFill And sh.CanHaveOutline Then
          If sh.Fill.Type = cdrUniformFill Then
             If sh.Outline.Type = cdrNoOutline Then sh.Outline.Type = cdrOutline
             sh.Outline.Color.CopyAssign sh.Fill.UniformColor
          End If
       End If
    Next sh
    boostFinish True
    sr.CreateSelection
    End Sub
    
    


Sub SetScale()
    If ActiveSelectionRange.Count = 0 Then Exit Sub
    boostStart "Edit Outline Scale"
    OutlineScaleEdit ActiveSelectionRange, True
    boostFinish endUndoGroup:=True
    ActiveDocument.ClearSelection
End Sub
Sub SetNoScale()
    If ActiveSelectionRange.Count = 0 Then Exit Sub
    boostStart "Edit Outline Scale"
    OutlineScaleEdit ActiveSelectionRange, False
    boostFinish endUndoGroup:=True
    ActiveDocument.ClearSelection
End Sub
Private Sub OutlineScaleEdit(sr As ShapeRange, si As Boolean)
    Dim s As Shape
    On Error Resume Next
    For Each s In sr
        If s.Type = cdrGroupShape Then
            OutlineScaleEdit s.Shapes.All, si
        Else
            If s.CanHaveOutline Then
                Select Case s.Outline.Type
                Case cdrNoOutline
                    If s.Outline.ScaleWithShape = True Then
                        s.Outline.ScaleWithShape = False
                        s.Outline.Type = cdrNoOutline
                    End If
                Case cdrOutline
                    If s.Outline.ScaleWithShape <> si Then
                        s.Outline.ScaleWithShape = si
                    End If
                End Select
            End If
        End If
        If Not s.PowerClip Is Nothing Then _
            OutlineScaleEdit s.PowerClip.Shapes.All, si
    Next 's
End Sub





Sub OutlineWidthUp()
    sOutlineWidth 1, GetValWidth
End Sub
Sub OutlineWidthDown()
    sOutlineWidth 2, GetValWidth
End Sub
Private Sub sOutlineWidth(t&, w#)
    If ActiveDocument Is Nothing Then Exit Sub
    If ActiveSelectionRange.Count < 1 Then Exit Sub
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    ActiveDocument.Unit = cdrPoint
    boostStart "Outline Width Edit"
    Call DoOutlineSize(sr, t, w)
    ActiveDocument.ClearSelection
    sr.AddToSelection
    boostFinish True
End Sub
Private Sub DoOutlineSize(sr As ShapeRange, t&, w#)
    Dim s As Shape
    On Error Resume Next
    For Each s In sr
        If s.Type = cdrGroupShape Then
            DoOutlineSize s.Shapes.All, t, w
        Else
            If s.CanHaveOutline Then
                If s.Outline.Width > 0 Then
                    If t = 1 Then s.Outline.Width = s.Outline.Width + w Else s.Outline.Width = s.Outline.Width - w
                End If
            End If
        End If
        If Not s.PowerClip Is Nothing Then _
            DoOutlineSize s.PowerClip.Shapes.All, t, w
    Next
End Sub
Private Function GetValWidth() As Double
    If (GetKeyState(vbKeyShift) And &HFF80) <> 0 = True Then
        GetValWidth = 5
    ElseIf (GetKeyState(vbKeyControl) And &HFF80) <> 0 = True Then
        GetValWidth = 0.1
    Else
        GetValWidth = 1
    End If
End Function
