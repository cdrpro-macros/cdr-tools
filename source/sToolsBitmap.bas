Attribute VB_Name = "sToolsBitmap"
Option Explicit

Sub ToPowerClip()
    Dim sr As ShapeRange, s As Shape, r As Shape
    Set sr = New ShapeRange
    Set sr = ActiveSelectionRange.FindAnyOfType(cdrBitmapShape)
    On Error Resume Next
    boostStart "Bitmaps to PowerClips"
    ActiveDocument.ReferencePoint = cdrTopLeft
    For Each s In sr
        With s
            Set r = ActiveLayer.CreateRectangle2(.PositionX, .PositionY - .SizeHeight, .SizeWidth, .SizeHeight)
            r.Outline.SetProperties , , , , , cdrTrue, cdrTrue, , cdrOutlineRoundLineJoin
            s.AddToPowerClip r, cdrTrue
        End With
    Next s
    boostFinish True
End Sub
