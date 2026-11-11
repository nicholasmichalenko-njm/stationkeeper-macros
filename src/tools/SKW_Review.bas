Attribute VB_Name = "SKW_Review"
' =====================================================================
' Stationkeeper Tools: review before the document goes out.
'   PreDeliveryLeftovers     report: comments, tracked changes, hidden text,
'                            highlights, placeholders and file details
' =====================================================================
Option Explicit

' Lists what should not go out with the document: comments, tracked changes, hidden text, highlighting, placeholder text (TBD, XXX, [insert ...]) and the author and company saved in the file. Changes nothing.
Public Sub PreDeliveryLeftovers()
    SKW_Show SKW_PreDeliveryLeftovers()
End Sub

Public Function SKW_PreDeliveryLeftovers() As String
    Dim d As Object, msg As String, rows As New Collection, story As Object, st As String, c As Object, rv As Object
    Dim r As Object, ph As Variant, nCom As Long, nRev As Long, nHid As Long, nHi As Long, nPh As Long, n As Long
    Dim showHidden As Boolean, win As Object, propName As Variant, v As String, summary As String
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Pre-delivery leftovers", False)
    If Len(msg) > 0 Then SKW_PreDeliveryLeftovers = msg: Exit Function
    Set d = ActiveDocument

    ' Comments
    For Each c In d.Comments
        nCom = nCom + 1
        If nCom <= 300 Then rows.Add SKW_Row("Comment", "Comment by " & c.Author, "Body", SKW_Page(c.Scope), c.Range.Text)
    Next
    ' Tracked changes, every part of the document
    For Each story In SKW_Stories(d)
        If story.StoryType <> wdCommentsStory Then
            st = SKW_StoryName(story.StoryType)
            For Each rv In story.Revisions
                nRev = nRev + 1
                If nRev <= 300 Then rows.Add SKW_Row("Tracked change", RevName(rv.Type) & " by " & rv.Author, st, SKW_Page(rv.Range), rv.Range.Text)
            Next
        End If
    Next
    ' Hidden text and highlighting (hidden text is only found while it is shown)
    On Error Resume Next
    Set win = d.ActiveWindow
    On Error GoTo 0
    If Not win Is Nothing Then
        showHidden = win.View.ShowHiddenText
        win.View.ShowHiddenText = True
    End If
    For Each story In SKW_Stories(d)
        If story.StoryType <> wdCommentsStory Then
            st = SKW_StoryName(story.StoryType)
            For Each r In FormatRanges(story, "hidden")
                nHid = nHid + 1
                If nHid <= 300 Then rows.Add SKW_Row("Hidden text", "Text formatted as hidden", st, SKW_Page(r), r.Text)
            Next
            For Each r In FormatRanges(story, "highlight")
                nHi = nHi + 1
                If nHi <= 300 Then rows.Add SKW_Row("Highlight", "Highlighted text", st, SKW_Page(r), r.Text)
            Next
            For Each ph In SKW_Placeholders()
                For Each r In SKW_TextRanges(story, CStr(ph), False, False)
                    nPh = nPh + 1
                    r.Expand wdSentence
                    If nPh <= 300 Then rows.Add SKW_Row("Placeholder", "Contains " & ph, st, SKW_Page(r), r.Text)
                Next
            Next
        End If
    Next
    If Not win Is Nothing Then win.View.ShowHiddenText = showHidden
    ' File details
    For Each propName In Array("Author", "Last author", "Company", "Manager")
        v = ""
        On Error Resume Next
        v = Trim$(d.BuiltInDocumentProperties(propName).Value)
        On Error GoTo 0
        If Len(v) > 0 Then rows.Add SKW_Row("File details", propName & " is saved in the file", "File > Info", "", v)
    Next
    n = 0
    On Error Resume Next
    n = d.Variables.Count
    On Error GoTo 0
    If n > 0 Then rows.Add SKW_Row("File details", n & " document variable(s) saved in the file", "Hidden", "", "")
    SKW_Finish

    summary = "Comments " & nCom & ", tracked changes " & nRev & ", hidden text " & nHid & ", highlights " & nHi & _
              ", placeholders " & nPh & "."
    SKW_Report "Pre-Delivery Leftovers", summary & " Up to 300 of each are listed. Placeholders: " & _
        JoinList(SKW_Placeholders()) & ". Changes nothing.", Array("Rule", "Finding", "Where", "Page", "Text"), rows
    SKW_PreDeliveryLeftovers = summary
End Function

Private Function RevName(ByVal t As Long) As String
    Select Case t
        Case wdRevisionInsert: RevName = "Insertion"
        Case wdRevisionDelete: RevName = "Deletion"
        Case wdRevisionProperty, wdRevisionParagraphProperty, wdRevisionStyle: RevName = "Formatting change"
        Case wdRevisionMovedFrom: RevName = "Moved from"
        Case wdRevisionMovedTo: RevName = "Moved to"
        Case Else: RevName = "Change"
    End Select
End Function

Private Function JoinList(ByVal c As Collection) As String
    Dim v As Variant
    For Each v In c
        JoinList = JoinList & IIf(Len(JoinList) > 0, ", ", "") & v
    Next
End Function

' Ranges with hidden or highlighted formatting in one story.
Private Function FormatRanges(ByVal story As Object, ByVal what As String) As Collection
    Dim c As New Collection, f As Object, lastStart As Long
    Set f = story.Duplicate
    On Error GoTo Done
    With f.Find
        .ClearFormatting
        .Text = ""
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        If what = "hidden" Then .Font.Hidden = True Else .Highlight = True
    End With
    lastStart = -1
    Do While f.Find.Execute
        If f.Start = lastStart Or f.End <= f.Start Then Exit Do
        lastStart = f.Start
        c.Add f.Duplicate
        If c.Count >= 1000 Or f.End >= story.End Then Exit Do
        f.Collapse wdCollapseEnd
    Loop
Done:
    Set FormatRanges = c
End Function
