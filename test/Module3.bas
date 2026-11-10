Attribute VB_Name = "Module3"
' Risky patterns for the health check. SaveCopyToC must never run in tests.
Option Explicit

Sub SaveCopyToC()
    ActiveDocument.SaveAs2 Environ$("TEMP") & "\skm-never-run.docx"
End Sub

Sub QuietErrors()
    On Error Resume Next
    ActiveDocument.Range(0, 0).InsertBefore ""
End Sub

Sub AlertsOff()
    Application.DisplayAlerts = wdAlertsNone
End Sub

Sub FixedPath()
    Dim p As String
    p = "C:\Proposals\Library\boilerplate.docx"
End Sub

Sub TwoEdits()
    ActiveDocument.Range(0, 0).InsertBefore "AAA "
    ActiveDocument.Range(0, 0).InsertBefore "BBB "
End Sub
