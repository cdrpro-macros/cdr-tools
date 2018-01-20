Attribute VB_Name = "sToolsWorkspace"
Option Explicit

Sub SwitchRulersGuidelines()
    If ActiveDocument Is Nothing Then Exit Sub
    Application.FrameWork.Automation.Invoke "4a490617-54c0-4263-99ae-8da808884f50"
    Application.FrameWork.Automation.Invoke "fc6531ef-665a-4548-b357-eda407a8dbd6"
End Sub
