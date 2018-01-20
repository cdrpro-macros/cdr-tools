Attribute VB_Name = "sToolsAbout"
Public Const REGAPPNAME$ = "CdrTools"
Public Const REGAPPOPT$ = "Options"
Public Const macroVersion$ = "2.11"
Public Const myWebSite$ = "macros.cdrpro.ru"

Private Const myEmail$ = "sancho@cdrpro.ru"
Private Const macroModifyDate$ = "13.02.2013"

Sub ShowAbout()
    MsgBox "Version " & macroVersion & " (" & macroModifyDate & ")" & Chr(10) & _
    "Copyright " & Chr(169) & " 2013 by Sancho   " & Chr(10) & _
    Chr(10) & _
    "http://" & myWebSite & Chr(10) & _
    "e-mail: " & myEmail & Chr(10), vbInformation, "About CdrTools"
End Sub

Sub OpenOptionsWindow()
    cdrOptionsWindow.Show
End Sub
