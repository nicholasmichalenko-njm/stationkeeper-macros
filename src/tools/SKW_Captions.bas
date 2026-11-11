Attribute VB_Name = "SKW_Captions"
' =====================================================================
' Stationkeeper Tools: captions.
'   CaptionAudit   report: tables and figures without captions, numbers typed
'                  instead of fields, wrong labels, captions on the wrong side
' =====================================================================
Option Explicit

' Checks every table and figure for a caption, captions whose number is typed instead of a field, captions with the wrong label (Figure under a table), captions on the wrong side (setting TableCaptionPosition and FigureCaptionPosition) and repeated caption text. Changes nothing.
Public Sub CaptionAudit()
    SKW_Show SKW_CaptionAudit()
End Sub

' The SEQ label in a caption paragraph ("Table", "Figure"), or "" when there is none.
Private Function SeqLabel(ByVal r As Object) As String
    Dim f As Object, c As String, parts As Variant
    On Error Resume Next
    For Each f In r.Fields
        If f.Type = wdFieldSequence Then
            c = Trim$(f.Code.Text)
            parts = Split(c, " ")
            If UBound(parts) >= 1 Then SeqLabel = parts(1)
            Exit Function
        End If
    Next
End Function

Private Function IsCapPara(ByVal r As Object, ByVal capStyle As String) As Boolean
    If r Is Nothing Then Exit Function
    If LCase$(SKW_ParaStyleName(r)) = LCase$(capStyle) Then IsCapPara = True: Exit Function
    If Len(SeqLabel(r)) > 0 Then IsCapPara = True
End Function

' The paragraph just before or just after a range (Nothing at the ends, or in a table).
Private Function Beside(ByVal rng As Object, ByVal after As Boolean) As Object
    Dim r As Object
    On Error GoTo None
    Set r = rng.Duplicate
    If after Then
        r.Collapse wdCollapseEnd
        If r.Information(wdWithInTable) Then GoTo None
        r.Expand wdParagraph
    Else
        r.Collapse wdCollapseStart
        If r.Start = 0 Then GoTo None
        r.Move wdCharacter, -1
        If r.Information(wdWithInTable) Then GoTo None
        r.Expand wdParagraph
    End If
    Set Beside = r
    Exit Function
None:
    Set Beside = Nothing
End Function

Private Sub CheckOne(ByVal what As String, ByVal n As Long, ByVal rng As Object, ByVal capStyle As String, ByVal wantSide As String, _
                     ByVal label As String, ByVal rows As Collection, ByRef nNo As Long, ByRef nSide As Long, ByRef nLabel As Long)
    Dim above As Object, below As Object, cap As Object, side As String, lab As String, first As String
    Set above = Beside(rng, False)
    Set below = Beside(rng, True)
    ' Look on the team's side first: a caption on the other side may belong to the item next door.
    If LCase$(wantSide) = "below" Then
        If IsCapPara(below, capStyle) Then Set cap = below: side = "below"
        If cap Is Nothing And IsCapPara(above, capStyle) Then Set cap = above: side = "above"
    Else
        If IsCapPara(above, capStyle) Then Set cap = above: side = "above"
        If cap Is Nothing And IsCapPara(below, capStyle) Then Set cap = below: side = "below"
    End If
    On Error Resume Next
    first = SKW_Excerpt(rng.Text, 40)
    On Error GoTo 0
    If cap Is Nothing Then
        nNo = nNo + 1
        rows.Add SKW_Row(what & " " & n, SKW_Page(rng), "No caption", "No caption above or below", first)
        Exit Sub
    End If
    If LCase$(side) <> LCase$(wantSide) And Len(wantSide) > 0 Then
        nSide = nSide + 1
        rows.Add SKW_Row(what & " " & n, SKW_Page(rng), "Position", "Caption is " & side & "; the team puts " & LCase$(what) & " captions " & LCase$(wantSide), cap.Text)
    End If
    lab = SeqLabel(cap)
    If Len(lab) > 0 And LCase$(lab) <> LCase$(label) Then
        nLabel = nLabel + 1
        rows.Add SKW_Row(what & " " & n, SKW_Page(rng), "Label", "Caption numbers this as " & lab & ", not " & label, cap.Text)
    End If
End Sub

Public Function SKW_CaptionAudit() As String
    Dim d As Object, msg As String, rows As New Collection, capStyle As String, t As Object, i As Long, sh As Object
    Dim nNo As Long, nSide As Long, nLabel As Long, nTyped As Long, nDup As Long, p As Object, seen As Object, txt As String
    Dim tSide As String, fSide As String, tLabel As String, fLabel As String, nFig As Long, r As Object
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Caption audit", False)
    If Len(msg) > 0 Then SKW_CaptionAudit = msg: Exit Function
    Set d = ActiveDocument
    capStyle = SKW_Setting("CaptionStyle", "Caption")
    tSide = SKW_Setting("TableCaptionPosition", "Above")
    fSide = SKW_Setting("FigureCaptionPosition", "Below")
    tLabel = SKW_Setting("TableLabel", "Table")
    fLabel = SKW_Setting("FigureLabel", "Figure")
    For Each t In d.Tables
        i = i + 1
        CheckOne "Table", i, t.Range, capStyle, tSide, tLabel, rows, nNo, nSide, nLabel
    Next
    For Each sh In d.InlineShapes
        Select Case sh.Type
            Case 3, 4, 12, 15   ' picture, linked picture, chart, SmartArt
                If Not sh.Range.Information(wdWithInTable) Then
                    nFig = nFig + 1
                    Set r = sh.Range.Paragraphs(1).Range
                    CheckOne "Figure", nFig, r, capStyle, fSide, fLabel, rows, nNo, nSide, nLabel
                End If
        End Select
    Next
    ' Captions with a typed number, and repeated caption text.
    Set seen = CreateObject("Scripting.Dictionary")
    For Each r In SKW_StyleRanges(d.Content, SKW_GetStyle(d, capStyle), 2000)
        For Each p In r.Paragraphs
            txt = SKW_Excerpt(p.Range.Text, 200)
            If Len(txt) > 0 Then
                If Len(SeqLabel(p.Range)) = 0 And txt Like "*[0-9]*" Then
                    nTyped = nTyped + 1
                    rows.Add SKW_Row("Caption", SKW_Page(p.Range), "Typed number", "Number is typed, so it will not renumber or update cross-references. Use References > Insert Caption", txt)
                End If
                If seen.Exists(LCase$(txt)) Then
                    nDup = nDup + 1
                    rows.Add SKW_Row("Caption", SKW_Page(p.Range), "Repeated", "Same caption text as page " & seen(LCase$(txt)), txt)
                Else
                    seen(LCase$(txt)) = SKW_Page(p.Range)
                End If
            End If
        Next
    Next
    SKW_Finish
    SKW_Report "Caption Audit", d.Tables.Count & " table(s) and " & nFig & " figure(s) checked. No caption " & nNo & ", wrong side " & nSide & _
        ", wrong label " & nLabel & ", typed numbers " & nTyped & ", repeated " & nDup & ". Team rules: table captions " & LCase$(tSide) & _
        ", figure captions " & LCase$(fSide) & ", style " & capStyle & ". Floating pictures are not checked. Changes nothing.", _
        Array("Item", "Page", "Rule", "Finding", "Text"), rows
    SKW_CaptionAudit = "No caption " & nNo & ", wrong side " & nSide & ", wrong label " & nLabel & ", typed numbers " & nTyped & ", repeated " & nDup & "."
End Function
