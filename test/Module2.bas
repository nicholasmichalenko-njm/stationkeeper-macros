Attribute VB_Name = "Module2"
Option Explicit

Sub WriteTestMarker()
    ActiveDocument.Range(0, 0).InsertBefore "SKTEST "
End Sub
