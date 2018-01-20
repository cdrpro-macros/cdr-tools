Attribute VB_Name = "sToolsTest"
Option Explicit

Sub SpeedTest()
  If Documents.Count > 0 Then
    MsgBox "Please close all documents and try again", vbExclamation, REGAPPNAME
    Exit Sub
  End If
  cdrTestForm.Show 1
End Sub
