Attribute VB_Name = "sToolsShape"
Option Explicit

#If VBA7 Then
  Declare PtrSafe Function GetKeyState& Lib "user32" (ByVal vKey&)
#Else
  Declare Function GetKeyState& Lib "user32" (ByVal vKey&)
#End If







Sub Oculist(): Call cdrOculist.Show(0): End Sub





Sub LineSpacingUp()
    If (GetKeyState(vbKeyControl) And &HFF80) <> 0 = True Then DoLineSpacing -1 Else DoLineSpacing -2
End Sub
Sub LineSpacingDown()
    If (GetKeyState(vbKeyControl) And &HFF80) <> 0 = True Then DoLineSpacing 1 Else DoLineSpacing 2
End Sub
Private Sub DoLineSpacing(t As Single)
    If ActiveDocument Is Nothing Then Exit Sub
    If ActiveShape Is Nothing Then Exit Sub
    If ActiveShape.Type <> cdrTextShape Then Exit Sub
    
    Dim tbLS1 As Variant, tbLS2 As Variant, l As Single
    tbLS1 = GetSetting(REGAPPNAME, REGAPPOPT, "Line Spacing 1", "1")
    tbLS2 = GetSetting(REGAPPNAME, REGAPPOPT, "Line Spacing 2", "5")
    If IsNumeric(tbLS1) = False Then tbLS1 = 1
    If IsNumeric(tbLS2) = False Then tbLS2 = 5
    
    Select Case t
        Case -1: l = tbLS1 * -1
        Case -2: l = tbLS2 * -1
        Case 1: l = tbLS1
        Case 2: l = tbLS2
    End Select
    
    With ActiveShape.Text
        If .IsEditing Then
            If .Selection.LineSpacing > 1 Then _
                .Selection.LineSpacing = .Selection.LineSpacing + l
        Else
            If .Story.LineSpacing > 1 Then _
                .Story.LineSpacing = .Story.LineSpacing + l
        End If
    End With
End Sub



Sub RotateRight()
    If ActiveDocument Is Nothing Then Exit Sub
    On Error Resume Next
    If (GetKeyState(vbKeyShift) And &HFF80) <> 0 = True Then _
    ActiveSelectionRange.Rotate GetRotValue(-2) Else ActiveSelectionRange.Rotate GetRotValue(-1)
End Sub
Sub RotateLeft()
    If ActiveDocument Is Nothing Then Exit Sub
    On Error Resume Next
    If (GetKeyState(vbKeyShift) And &HFF80) <> 0 = True Then _
    ActiveSelectionRange.Rotate GetRotValue(2) Else ActiveSelectionRange.Rotate GetRotValue(1)
End Sub
Private Function GetRotValue(t&) As Double
    Dim tbRot1 As Variant, tbRot2 As Variant
    tbRot1 = Val(Replace(GetSetting(REGAPPNAME, REGAPPOPT, "Rotate 1", "0.1"), ",", "."))
    tbRot2 = Val(Replace(GetSetting(REGAPPNAME, REGAPPOPT, "Rotate 2", "0.5"), ",", "."))
    Select Case t
        Case -1: GetRotValue = tbRot1 * -1
        Case -2: GetRotValue = tbRot2 * -1
        Case 1: GetRotValue = tbRot1
        Case 2: GetRotValue = tbRot2
    End Select
End Function


Sub MoveToDesktop()
    If ActiveDocument Is Nothing Then Exit Sub
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    If sr.Count > 0 Then sr.MoveToLayer ActiveDocument.MasterPage.DesktopLayer
End Sub

Sub MoveToCenter()
    If ActiveDocument Is Nothing Then Exit Sub
    If ActiveSelectionRange.Count < 1 Then Exit Sub
    On Error Resume Next
    boostStart "Move selection objects to center of view"
    Dim sr As ShapeRange, old_refpoint&
    old_refpoint = ActiveDocument.ReferencePoint
    ActiveDocument.ReferencePoint = cdrCenter
    Set sr = ActiveSelectionRange
    sr.SetPosition ActiveWindow.ActiveView.OriginX, ActiveWindow.ActiveView.OriginY
    ActiveDocument.ReferencePoint = old_refpoint
    boostFinish endUndoGroup:=True
End Sub
    
Sub SmartIntersect()
    If ActiveDocument Is Nothing Then Exit Sub
    If ActiveSelectionRange.Count <> 2 Then MsgBox "No selection": Exit Sub
    ActiveSelectionRange.Item(1).Intersect ActiveSelectionRange.Item(2), True, False
End Sub
    
    
Sub DeleteNoCloseCurves()
    If ActiveDocument Is Nothing Then Exit Sub
    If ActiveSelectionRange.Count = 0 Then Exit Sub
    Dim sr As ShapeRange, s As Shape
    Set sr = New ShapeRange
    Set sr = ActiveSelectionRange

    boostStart "Delete Not Closed Curves"
    For Each s In sr
        If s.Type = cdrCurveShape Then If s.Curve.Closed = False Then s.Delete
    Next
    boostFinish endUndoGroup:=True
End Sub

Sub DeArtBrush()
    If ActiveDocument Is Nothing Then Exit Sub
    ActiveDocument.BeginCommandGroup "DeArtBrush"
    On Error Resume Next
    Dim sr As ShapeRange, s As Shape, sr1 As ShapeRange
    Set sr = ActiveSelectionRange
    For Each s In sr
        If s.Type = cdrArtisticMediaGroupShape Then
            Set sr1 = s.GetLinkedShapes(cdrLinkAllConnections)
            s.Separate
            sr1(1).Delete
            sr1.Remove 1
        End If
    Next s
    ActiveDocument.EndCommandGroup
End Sub


Sub SwapTwoShapes()
    If ActiveDocument Is Nothing Then Exit Sub
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    If sr.Count = 2 Then
        On Error Resume Next
        boostStart "Swap Shapes"
        ActiveDocument.ReferencePoint = cdrCenter
         
        Dim b As Boolean
        If (GetKeyState(vbKeyShift) And &HFF80) <> 0 = True Then b = True
         
        Dim s1 As Shape, s2 As Shape, st As Shape
        Set s1 = sr(1): Set s2 = sr(2)
         
        Dim X#, Y#, w#, h#
         
        Set st = s1.Duplicate
        s2.GetPosition X, Y
        If b Then s2.GetSize w, h
        s1.SetPosition X, Y
        If b Then s1.SetSize w, h
        s1.OrderBackOf s2
         
        st.GetPosition X, Y
        If b Then st.GetSize w, h
        s2.SetPosition X, Y
        If b Then s2.SetSize w, h
        s2.OrderBackOf st
         
        st.Delete
        boostFinish True
    End If
End Sub


Sub GetCurveInfo()
    If ActiveDocument Is Nothing Then Beep: Exit Sub
    ActiveDocument.Unit = ActiveDocument.Rulers.HUnits
    On Error Resume Next
    If ActiveSelectionRange.Shapes.Count <> 1 Then Beep: Exit Sub
    Dim s As Shape, c As Curve
    Set s = ActiveSelectionRange(1)
    If s.Type = cdrCurveShape Then Set c = s.Curve Else Set c = s.DisplayCurve
    If c Is Nothing Then Beep: Exit Sub
    Dim sUnit$
    Select Case ActiveDocument.Unit
        Case 3: sUnit = " mm"
        Case 7: sUnit = " m"
        Case 4: sUnit = " cm"
        Case 1: sUnit = " " & Chr(34)
        Case 14: sUnit = " pt"
        Case Else: sUnit = ""
    End Select
    Dim myInfoMsg$
    If c.SubPaths.Count > 1 Then
        Dim sp As SubPath, st As SubPath, l&
        For l = 1 To c.SubPaths.Count
            If l = 1 Then
                Set st = c.SubPaths(l)
            Else
                If c.SubPaths(l).Area > st.Area Then Set st = c.SubPaths(l)
            End If
        Next
        
        Dim sArea#
        sArea = c.Area - st.Area
        
        myInfoMsg = "Total Area: " & Round(c.Area, 5) & sUnit & vbCr & _
        "Total Lenght: " & Round(c.Length, 5) & sUnit & vbCr & _
        "Nodes Count: " & c.Nodes.Count & vbCr & _
        "Paths Count: " & c.SubPaths.Count & vbCr & vbCr & _
        "===================" & vbCr & _
        "Area #1: " & Round(st.Area - sArea, 5) & sUnit & vbCr
    Else
        myInfoMsg = "Area: " & Round(c.Area, 5) & sUnit & vbCr & _
        "Length: " & Round(c.Length, 5) & sUnit & vbCr & _
        "Nodes Count: " & c.Nodes.Count & vbCr
    End If
    'myInfoMsg = myInfoMsg & vbCr & "Copyright " & Chr(169) & " 2010 " & "by Sancho," & Chr(10) & "www.cdrpro.ru"
    If myInfoMsg <> "" Then MsgBox myInfoMsg, vbInformation, "Curve Info (version 2.1)"
End Sub


