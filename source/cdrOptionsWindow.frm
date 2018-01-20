VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cdrOptionsWindow 
   Caption         =   "UserForm1"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4335
   OleObjectBlob   =   "cdrOptionsWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cdrOptionsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub UserForm_Initialize()
    Me.Caption = REGAPPOPT & " ::: " & REGAPPNAME & " " & macroVersion
    tbLS1.Text = GetSetting(REGAPPNAME, REGAPPOPT, "Line Spacing 1", "1")
    tbLS2.Text = GetSetting(REGAPPNAME, REGAPPOPT, "Line Spacing 2", "5")
    tbRot1.Text = GetSetting(REGAPPNAME, REGAPPOPT, "Rotate 1", "0.1")
    tbRot2.Text = GetSetting(REGAPPNAME, REGAPPOPT, "Rotate 2", "0.5")
    cbAutoUpdateText = (GetSetting(REGAPPNAME, REGAPPOPT, "Auto Update Text", "0") = "1")
End Sub


Private Sub cbSAVE_Click()
    Unload Me
End Sub

Private Sub cbAbout_Click()
    sToolsAbout.ShowAbout
End Sub



Private Sub tbLS1_Change()
    SaveSetting REGAPPNAME, REGAPPOPT, "Line Spacing 1", Trim(tbLS1.Text)
End Sub

Private Sub tbLS2_Change()
    SaveSetting REGAPPNAME, REGAPPOPT, "Line Spacing 2", Trim(tbLS2.Text)
End Sub

Private Sub tbRot1_Change()
    SaveSetting REGAPPNAME, REGAPPOPT, "Rotate 1", Trim(tbRot1.Text)
End Sub

Private Sub tbRot2_Change()
    SaveSetting REGAPPNAME, REGAPPOPT, "Rotate 2", Trim(tbRot2.Text)
End Sub



Private Sub cbAutoUpdateText_Click()
    SaveSetting REGAPPNAME, REGAPPOPT, "Auto Update Text", IIf(cbAutoUpdateText, "1", "0")
End Sub
