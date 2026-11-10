Attribute VB_Name = "Module1"
' A stand-in for a colleague's macro set: everything in one module.
Option Explicit

' Repeats the header row of every table at the top of each page.
Sub RepeatHeaderRows()
    Dim t As Table
    For Each t In ActiveDocument.Tables
        t.Rows(1).HeadingFormat = True
    Next
End Sub

'-------------------------------------------
' Updates every field and table of contents.
'-------------------------------------------
Public Sub UpdateAllFields()
    ActiveDocument.Fields.Update
End Sub

Public Sub ShowHiddenText()
    ' Turns hidden text display on so you can see RFP text under headings.
    ActiveWindow.View.ShowHiddenText = True
End Sub

Sub ClearDirectFormatting()
    Selection.ClearFormatting
End Sub

Sub InsertCaption_Table()
    Selection.InsertCaption Label:="Table", Position:=wdCaptionPositionAbove
End Sub

Public Sub FitTablesToMargins( _
    )
    Dim t As Table
    For Each t In ActiveDocument.Tables
        t.PreferredWidthType = wdPreferredWidthPercent
        t.PreferredWidth = 100
    Next
End Sub

Sub TOCFix()
    Dim t As TableOfContents
    For Each t In ActiveDocument.TablesOfContents
        t.Update
    Next
End Sub

Private Sub HelperNotAMacro()
End Sub

Sub NeedsArgs(ByVal x As Long)
End Sub

Function ReturnsValue() As Long
    ReturnsValue = 1
End Function

Sub AutoOpen()
End Sub
