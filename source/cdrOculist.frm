VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cdrOculist 
   Caption         =   "Oculist"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3000
   OleObjectBlob   =   "cdrOculist.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cdrOculist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
  'Initialize the random # generator.
  Randomize
End Sub


Private Sub cbDo_Click()
  If ActiveDocument Is Nothing Then Exit Sub
  If ActiveSelectionRange.Count = 0 Then Exit Sub
  
  On Error Resume Next
  Dim s As Shape, c&, d#, c2&, max&, min&, i&
  
  boostStart REGAPPNAME & ": Edit Text"
  
  For Each s In ActiveSelectionRange
    If s.Type = cdrTextShape Then
      
      max = CLng(tbMax.Text)
      min = CLng(tbMin.Text)
      c = s.Text.Story.Characters.Count
      If c > 0 Then
      
        d = c / (max - min)
        c2 = Round(d)
        If d > c2 Then c2 = c2 + 1
        For i = 0 To c Step c2
          If optDown.Value Then
            s.Text.Story.Range(i, (i + c2)).Size = max
            max = max - 1
          End If
          If optUp.Value Then
            s.Text.Story.Range(i, (i + c2)).Size = min
            min = min + 1
          End If
          If optRandom.Value Then
            s.Text.Story.Range(i, (i + c2)).Size = Rand(min, max)
          End If
        Next
        
      End If
    End If
  Next
  
  boostFinish True
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
  Rand = Int((High - Low + 1) * rnd) + Low
End Function
