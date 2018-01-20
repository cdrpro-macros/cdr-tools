Attribute VB_Name = "sToolsBoost"
Option Explicit

Private MSGBADVER$
Public MSGKEYINFO$, MSGOTHERCOMP$

Public Sub boostStart(Optional ByVal unDo$ = "")
    On Error Resume Next
    If unDo <> "" Then ActiveDocument.BeginCommandGroup unDo
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
End Sub

Public Sub boostFinish(Optional ByVal endUndoGroup% = False)
    On Error Resume Next
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    Application.CorelScript.RedrawScreen
    If endUndoGroup Then ActiveDocument.EndCommandGroup
    CorelDRAW.Refresh
End Sub
