VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cdrTestForm 
   Caption         =   "UserForm1"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4305
   OleObjectBlob   =   "cdrTestForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cdrTestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SpeedTestName$ = "CdrSpeedTest"


Private Sub UserForm_Activate()
    Call SpeedTestStart
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = SpeedTestName
    'sProgress.Width = 1
End Sub



Private Sub SpeedTestStart()
    On Error Resume Next
    
    sProgress.Width = 1
    DoEvents
    
    Dim d As Document, l As Layer
    If VersionMajor < 15 Then Set d = CreateDocument Else Set d = CreateDocEx()
    d.Unit = cdrPoint
    If VersionMajor < 15 Then d.MasterPage.SetSize 200#, 200#
    d.DrawingOriginX = -d.ActivePage.SizeWidth / 2
    d.DrawingOriginY = -d.ActivePage.SizeHeight / 2
    d.ActiveWindow.ActiveView.ToFitPage
    Set l = d.ActiveLayer
    
    'With Redrawing
    Dim st As Variant, et As Variant
    st = Timer
    DoTest l
    et = Timer
    Dim p1 As Variant
    p1 = et - st
    
    sProgress.Width = 25 * 2
    DoEvents
    
    'Without Redrawing
    boostStart "Speed Test"
    st = Timer
    DoTest l
    et = Timer
    boostFinish endUndoGroup:=True
    Dim p2 As Variant
    p2 = et - st
    
    sProgress.Width = 30 * 2
    DoEvents
    
    'Export
    st = Timer
    Dim expopt As StructExportOptions, expflt As ExportFilter
    d.ClearSelection
    d.ActiveLayer.Shapes.All.AddToSelection
    Set expopt = CreateStructExportOptions
    With expopt
        .Overwrite = True
        .SizeX = 5000
        .SizeY = 5000
        .ResolutionX = 300
        .ResolutionY = 300
        .ImageType = cdrCMYKColorImage
        .AntiAliasingType = cdrNormalAntiAliasing
        .UseColorProfile = True
        .Compression = cdrCompressionLZW
        .Dithered = False
        .Transparent = False
        .MaintainLayers = False
        '.AlwaysOverprintBlack = True
    End With
    Set expflt = d.ExportEx(Environ$("Temp") & "\CDRPRO_RU_CorelDRAW_Test.tif", cdrTIFF, cdrSelection, expopt)
    expflt.Finish
    et = Timer
    Dim p5 As Variant
    p5 = et - st
    
    sProgress.Width = 45 * 2
    DoEvents
    
    'Import
    st = Timer
    Dim impopt As StructImportOptions, impflt As ImportFilter
    Set impopt = CreateStructImportOptions
    With impopt
        .Mode = cdrImportFull
        .MaintainLayers = True
    End With
    Set impflt = ActiveLayer.ImportEx(Environ$("Temp") & "\CDRPRO_RU_CorelDRAW_Test.tif", cdrTIFF, impopt)
    impflt.Finish
    et = Timer
    Dim p6 As Variant
    p6 = et - st
    
    sProgress.Width = 60 * 2
    DoEvents
    
    boostStart "Speed Test"
    'Bitmap Curves
    Dim c&
    For c = 1 To 6
        ActiveSelectionRange.Duplicate
    Next
    Dim sr As ShapeRange, s As Shape
    Set sr = d.ActiveLayer.Shapes.FindShapes(, cdrBitmapShape)
    st = Timer
    c = 0
    For Each s In sr
        s.Bitmap.ApplyBitmapEffect "Tone Curve", "ToneCurveEffect ToneCurveTable=0|2|5|8|10|13|16|18|21|24|27|29|32|35|37|40|43|45|48|51|54|56|59|62|64|67|70|72|75|78|81|83|86|89|91|94|97|99|102|105|108|110|113|116|118|121|124|126|129|132|135|137|140|143|145|148|151|153|156|159|162|161|160|159|158|157|156|155|154|153|152|151|150|149|148|147|146|145|144|143|142|141|140|139|138|137|136|135|134|132|131|130|129|128|127|126|125|124|123|122|121|120|119|118|117|116|115|114|113|112|111|110|109|108|107|106|105|104|102|101|100|99|98|97|96|95|94|93|92|91|90|89|88|87|86|85|84|83|82|81|80|79|78|77|76|75|74|72|73|75|77|79|81|83|84|86|88|90|92|94|95|97|99|101|103|105|107|108|110|112|114|116|118|119|121|123|125|127|129|130|132|134|136|138|140|142|143|145|147|149|151|153|154|156|158|160|162|164|165|167|169|171|173|175|177|178|180|181|183|184|186|187|189|19" & _
            "0|192|193|195|196|198|199|201|203|204|206|207|209|210|212|213|215|216|218|219|221|222|224|225|227|229|230|232|233|235|236|238|239|241|242|244|245|247|248|250|251|253|255|0|2|5|8|10|13|16|18|21|24|27|29|32|35|37|40|43|45|48|51|54|56|59|62|64|67|70|72|75|78|81|83|86|89|91|94|97|99|102|105|108|110|113|116|118|121|124|126|129|132|135|137|140|143|145|148|151|153|156|159|162|161|160|159|158|157|156|155|154|153|152|151|150|149|148|147|146|145|144|143|142|141|140|139|138|137|136|135|134|132|131|130|129|128|127|126|125|124|123|122|121|120|119|118|117|116|115|114|113|112|111|110|109|108|107|106|105|104|102|101|100|99|98|97|96|95|94|93|92|91|90|89|88|87|86|85|84|83|82|81|80|79|78|77|76|75|74|72|73|75|77|79|81|83|84|86|88|90|92|94|95|97|99|101|103|105|107|108|110|112|114|116|118|119|121|123|125|127|129" & _
            "|130|132|134|136|138|140|142|143|145|147|149|151|153|154|156|158|160|162|164|165|167|169|171|173|175|177|178|180|181|183|184|186|187|189|190|192|193|195|196|198|199|201|203|204|206|207|209|210|212|213|215|216|218|219|221|222|224|225|227|229|230|232|233|235|236|238|239|241|242|244|245|247|248|250|251|253|255|0|2|5|8|10|13|16|18|21|24|27|29|32|35|37|40|43|45|48|51|54|56|59|62|64|67|70|72|75|78|81|83|86|89|91|94|97|99|102|105|108|110|113|116|118|121|124|126|129|132|135|137|140|143|145|148|151|153|156|159|162|161|160|159|158|157|156|155|154|153|152|151|150|149|148|147|146|145|144|143|142|141|140|139|138|137|136|135|134|132|131|130|129|128|127|126|125|124|123|122|121|120|119|118|117|116|115|114|113|112|111|110|109|108|107|106|105|104|102|101|100|99|98|97|96|95|94|93|92|91|90|89|88|87|86|85|84|83" & _
            "|82|81|80|79|78|77|76|75|74|72|73|75|77|79|81|83|84|86|88|90|92|94|95|97|99|101|103|105|107|108|110|112|114|116|118|119|121|123|125|127|129|130|132|134|136|138|140|142|143|145|147|149|151|153|154|156|158|160|162|164|165|167|169|171|173|175|177|178|180|181|183|184|186|187|189|190|192|193|195|196|198|199|201|203|204|206|207|209|210|212|213|215|216|218|219|221|222|224|225|227|229|230|232|233|235|236|238|239|241|242|244|245|247|248|250|251|253|255|0|2|5|8|10|13|16|18|21|24|27|29|32|35|37|40|43|45|48|51|54|56|59|62|64|67|70|72|75|78|81|83|86|89|91|94|97|99|102|105|108|110|113|116|118|121|124|126|129|132|135|137|140|143|145|148|151|153|156|159|162|161|160|159|158|157|156|155|154|153|152|151|150|149|148|147|146|145|144|143|142|141|140|139|138|137|136|135|134|132|131|130|129|128|127|126|125|124|123" & _
            "|122|121|120|119|118|117|116|115|114|113|112|111|110|109|108|107|106|105|104|102|101|100|99|98|97|96|95|94|93|92|91|90|89|88|87|86|85|84|83|82|81|80|79|78|77|76|75|74|72|73|75|77|79|81|83|84|86|88|90|92|94|95|97|99|101|103|105|107|108|110|112|114|116|118|119|121|123|125|127|129|130|132|134|136|138|140|142|143|145|147|149|151|153|154|156|158|160|162|164|165|167|169|171|173|175|177|178|180|181|183|184|186|187|189|190|192|193|195|196|198|199|201|203|204|206|207|209|210|212|213|215|216|218|219|221|222|224|225|227|229|230|232|233|235|236|238|239|241|242|244|245|247|248|250|251|253|255|"
        c = c + 1
        sProgress2.Width = (c * 100 / sr.Count) * 2
        DoEvents
    Next
    et = Timer
    boostFinish True
    Dim p7 As Variant
    p7 = et - st
    
    sProgress.Width = 80 * 2
    DoEvents
    
    'Save
    Dim opt As New StructSaveAsOptions
    opt.Filter = cdrCDR
    opt.IncludeCMXData = False
    opt.Overwrite = True
    opt.Range = cdrAllPages
    opt.Version = cdrCurrentVersion
    
    st = Timer
    d.SaveAs Environ$("Temp") & "\CDRPRO_RU_CorelDRAW_Test.cdr", opt
    et = Timer
    Dim p3 As Variant
    p3 = et - st
    d.Close
    
    sProgress.Width = 90 * 2
    DoEvents
    
    'Open
    st = Timer
    Set d = OpenDocument(Environ$("Temp") & "\CDRPRO_RU_CorelDRAW_Test.cdr")
    et = Timer
    Dim p4 As Variant
    p4 = et - st
    
    d.Close
    FileSystem.Kill Environ$("Temp") & "\CDRPRO_RU_CorelDRAW_Test.cdr"
    FileSystem.Kill Environ$("Temp") & "\CDRPRO_RU_CorelDRAW_Test.tif"
    
    sProgress.Width = 100 * 2
    DoEvents
    
    MsgBox "With redrawing: " & Round(p1, 1) & vbCr & _
           "Without redrawing: " & Round(p2, 1) & vbCr & _
           "Redraw slowdown: " & Round((p1 / p2), 1) & vbCr & _
           "Save: " & Round(p3, 1) & vbCr & _
           "Open: " & Round(p4, 1) & vbCr & _
           "Export: " & Round(p5, 1) & vbCr & _
           "Import: " & Round(p6, 1) & vbCr & _
           "Work with bitmaps: " & Round(p7, 1), vbInformation, SpeedTestName & " Result"
    Unload Me
End Sub

Private Sub DoTest(l As Layer)
    Dim X#, Y#, h#, c&, i&, v#
    Dim s As Shape
    
    X = 0: Y = 0
    h = 20
    
    For c = 0 To 9
        For i = 0 To 9
            Set s = l.CreateRectangle2(X, Y, h, h)
            s.Fill.ApplyUniformFill CreateCMYKColor(c * 10, i * 10, 0, 0)
            s.Outline.Width = 1#
            s.Outline.Color.CMYKAssign 0, i * 10, c * 10, 0
            s.ConvertToBitmapEx cdrCMYKColorImage, False, True, 300, cdrNormalAntiAliasing, True
            X = X + h
            v = (c * 10 + i) * 2
            sProgress2.Width = v
            DoEvents
        Next
        X = 0
        Y = Y + h
    Next
End Sub

Private Function CreateDocEx() As Document
    Dim createopt As StructCreateOptions
    Set createopt = CreateStructCreateOptions
    With createopt
        .Name = "Test"
        .Units = cdrPoint
        .PageWidth = 200#
        .PageHeight = 200#
        .Resolution = 300#
        .ColorContext = CreateColorContext2("sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%", clrRenderPerceptual, clrColorModelCMYK)
    End With
    Set CreateDocEx = CreateDocumentEx(createopt)
End Function




