VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub GlobalMacroStorage_DocumentOpen(ByVal Doc As Document, ByVal FileName As String)
  If GetSetting(REGAPPNAME, REGAPPOPT, "Auto Update Text", "0") = "1" Then
    If Doc.TextFormatter < 1600 Then Doc.TextFormatter = 1600
  End If
End Sub
