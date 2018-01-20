Attribute VB_Name = "sToolsColor"
Option Explicit
Private sCol As CorelDRAW.cdrColorType, sImg As CorelDRAW.cdrImageType, sCou&, c&, ReColor As Boolean, VectorToo As Boolean

#If VBA7 Then
  Private Declare PtrSafe Function GetKeyState& Lib "user32" (ByVal vKey&)
#Else
  Private Declare Function GetKeyState& Lib "user32" (ByVal vKey&)
#End If


Sub ToRGB()
    sCol = cdrColorRGB
    sImg = cdrRGBColorImage
    VectorToo = True
    convTo "Convert To RGB"
    End Sub
Sub ToCMYK()
    sCol = cdrColorCMYK
    sImg = cdrCMYKColorImage
    ReColor = False
    VectorToo = True
    convTo "Convert To CMYK"
    End Sub
Sub ToGRAY()
    sCol = cdrColorGray
    sImg = cdrGrayscaleImage
    VectorToo = True
    convTo "Convert To GRAY"
    End Sub
Sub ReConvertCMYK()
    sCol = cdrColorCMYK
    sImg = cdrCMYKColorImage
    ReColor = True
    If (GetKeyState(vbKeyControl) And &HFF80) <> 0 = True Then _
        VectorToo = True Else VectorToo = False
    convTo "ReConvert To CMYK"
    End Sub
Private Sub convTo(ts$)
    If CorelDRAW.ActiveDocument Is Nothing Then Exit Sub
    Dim sr As ShapeRange: Set sr = New ShapeRange
    On Error Resume Next
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then MsgBox "No selection", vbInformation, " info": Exit Sub
    
    boostStart ts
        sCou = 0: c = 0
        shCount sr
        Application.Status.BeginProgress "Convert...", False
        convTo2 sr
        Application.Status.EndProgress
        ActiveDocument.ClearSelection
        sr.AddToSelection
    boostFinish endUndoGroup:=True
    End Sub
Private Sub convTo2(sr As ShapeRange)
    Dim s As Shape, fc As FountainColor
    On Error Resume Next
    For Each s In sr
        If s.Type = cdrGroupShape Then
            convTo2 s.Shapes.All
        Else
            Select Case s.Type
            Case cdrBitmapShape
                If s.Bitmap.Mode <> sImg Then
                    s.Bitmap.ConvertTo sImg
                Else
                    If ReColor Then
                        s.Bitmap.ConvertTo cdrRGBColorImage
                        s.Bitmap.ConvertTo sImg
                    End If
                End If
            Case Else
                If VectorToo Then
                    If s.CanHaveFill Then
                        Select Case s.Fill.Type
                        Case cdrUniformFill: fillAndOutlineColor s.Fill.UniformColor
                        Case cdrFountainFill
                            For Each fc In s.Fill.Fountain.Colors
                            fillAndOutlineColor fc.Color
                            Next fc
                        End Select
                    End If
                    If s.CanHaveOutline Then
                        If s.Outline.Type <> cdrNoOutline Then fillAndOutlineColor s.Outline.Color
                    End If
                End If
            End Select
            c = c + 1: Application.Status.Progress = c / sCou * 100
        End If
        If Not s.PowerClip Is Nothing Then convTo2 s.PowerClip.Shapes.All
    Next s
End Sub
    
Private Sub fillAndOutlineColor(myColor2 As Color)
    If myColor2.Type <> sCol Then
        Select Case sCol
            Case cdrColorRGB: myColor2.ConvertToRGB
            Case cdrColorCMYK
                If myColor2.Type = cdrColorPantone Or myColor2.Type = cdrColorSpot Then _
                myColor2.ConvertToRGB: myColor2.ConvertToCMYK Else myColor2.ConvertToCMYK
            Case cdrColorGray: myColor2.ConvertToGRAY
        End Select
    Else
        If ReColor Then
            myColor2.ConvertToRGB
            myColor2.ConvertToCMYK
        End If
    End If
End Sub

Private Sub shCount(sr As ShapeRange)
    Dim s As Shape
    On Error Resume Next
    For Each s In sr
        If Not s.PowerClip Is Nothing Then shCount s.PowerClip.Shapes.All
        If s.Type = cdrGroupShape Then shCount s.Shapes.All Else sCou = sCou + 1
    Next s
End Sub
    
    
    
    
    
    
    
    
    
    
    
    
    

Sub EditStepsNumOfFill()
    If ActiveDocument Is Nothing Then Exit Sub
    Dim sr As ShapeRange, s As Shape, c!
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then MsgBox "No selection": Exit Sub
    On Error Resume Next
    c = InputBox("Fountain Steps", "Long steps...", 256): c = Val(c)
    If c <= 0 Or c > 999 Then Exit Sub
    boostStart "Steps Fill Edit"
    StepsFillEdit2 sr, c
    boostFinish endUndoGroup:=True
    End Sub
Private Sub StepsFillEdit2(sr As ShapeRange, c!)
    Dim s As Shape
    On Error Resume Next
    For Each s In sr
        If s.Type = cdrGroupShape Then
            StepsFillEdit2 s.Shapes.All, c
        Else
            If s.Fill.Type = cdrFountainFill Then s.Fill.Fountain.Steps = c
            If s.Transparency.Type = cdrFountainTransparency Then s.Transparency.Fountain.Steps = c
        End If
        If Not s.PowerClip Is Nothing Then StepsFillEdit2 s.PowerClip.Shapes.All, c
    Next s
    End Sub
    
    
    
    
Sub SelectSame()
    If ActiveDocument Is Nothing Then Exit Sub
    Const sLable$ = "Magic Wand (macros.cdrpro.ru)"
    If ActiveShape Is Nothing Then MsgBox "No selection", vbExclamation, sLable: Exit Sub
    Dim sr As New ShapeRange, srOld As New ShapeRange, _
        b As Boolean, X#, Y#, Shift&, s As Shape, i&, s2 As Shape, _
        c As New Color, srFin As New ShapeRange
    
    Set srOld = ActiveSelectionRange.UngroupAllEx
    If ActiveDocument.GetUserClick(X, Y, Shift, 10, True, cdrCursorEyeDrop) Then Exit Sub
    Set sr = ActivePage.SelectShapesAtPoint(X, Y, False).Shapes.All
    Set s = sr(1)
    If sr.Count > 1 Then
        For i = 1 To sr.Count
            If i > 1 Then
                Set s2 = sr.Item(i)
                If s2.OrderIsInFrontOf(sr.Item(i - 1)) = True Then Set s = s2
            End If
        Next i
    End If
    
    If (GetKeyState(vbKeyControl) And &HFF80) <> 0 = True Then
        If s.Outline.Type = cdrNoOutline Then Beep: Exit Sub
        c.CopyAssign s.Outline.Color

        For Each s2 In srOld
            If s2.Outline.Type = cdrOutline Then
                If s2.Outline.Color.Name(True) = c.Name(True) Then srFin.Add s2
            End If
        Next
    Else
        If s.Fill.Type <> cdrUniformFill Then Beep: Exit Sub
        c.CopyAssign s.Fill.UniformColor
        
        For Each s2 In srOld
            If s2.Fill.Type = cdrUniformFill Then
                If s2.Fill.UniformColor.Name(True) = c.Name(True) Then srFin.Add s2
            End If
        Next
    End If
    ActiveDocument.ClearSelection
    srFin.CreateSelection
End Sub

