Attribute VB_Name = "SKW_Text"
' =====================================================================
' Stationkeeper Tools: text and proofing.
'   TypographySweep          report: double spaces, straight quotes, double hyphens and more
'   LoadTeamDictionaries     adds the team's custom dictionaries to Word
' =====================================================================
Option Explicit

' Lists typing slips: double spaces, a space before punctuation, straight quotes, double hyphens and spaced hyphens used as dashes, three dots instead of an ellipsis, and spaces at the end of a paragraph. Changes nothing.
Public Sub TypographySweep()
    SKW_Show SKW_TypographySweep()
End Sub

' Adds the team's custom dictionaries (Dictionaries sheet of the standards workbook) to Word's list on this PC, so team words are not marked as misspelled. Changes Word's settings, not the document.
Public Sub LoadTeamDictionaries()
    SKW_Say SKW_LoadTeamDictionaries()
End Sub

' ---------------------------------------------------------------------
Public Function SKW_TypographySweep() As String
    Dim d As Object, msg As String, rows As New Collection, rules As Variant, ru As Variant, r As Object, n As Long, x As Object
    Dim summary As String, story As Object
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Typography sweep", False)
    If Len(msg) > 0 Then SKW_TypographySweep = msg: Exit Function
    Set d = ActiveDocument
    ' find text, wildcards, name, finding
    rules = Array( _
        Array("[! ]  [! ]", True, "Double space", "Two spaces between words"), _
        Array(" [,.;:\?\!\)]", True, "Space before punctuation", "A space before punctuation"), _
        Array("""", False, "Straight quote", "Straight double quote (use curly quotes)"), _
        Array("--", False, "Double hyphen", "Two hyphens (use an en or em dash)"), _
        Array(" - ", False, "Spaced hyphen", "A hyphen with spaces (use an en dash)"), _
        Array("...", False, "Three dots", "Three dots (use an ellipsis)"), _
        Array("[ ]{1,}^13", True, "Space at paragraph end", "Space before the paragraph mark"))
    For Each story In SKW_Stories(d)
        If story.StoryType = wdMainTextStory Or story.StoryType = wdFootnotesStory Or story.StoryType = wdEndnotesStory Or story.StoryType = wdTextFrameStory Then
            For Each ru In rules
                n = 0
                For Each r In SKW_TextRanges(story, CStr(ru(0)), False, False, CBool(ru(1)), 500)
                    n = n + 1
                    Set x = r.Duplicate
                    x.MoveStart wdCharacter, -25
                    x.MoveEnd wdCharacter, 25
                    rows.Add SKW_Row(ru(2), ru(3), SKW_StoryName(story.StoryType), SKW_Page(r), Replace(x.Text, " ", ChrW(183)))
                Next
                If n > 0 Then summary = summary & ru(2) & " " & n & IIf(n >= 500, "+", "") & ", "
            Next
        End If
    Next
    SKW_Finish
    If Len(summary) = 0 Then summary = "Nothing found." Else summary = Left$(summary, Len(summary) - 2) & "."
    SKW_Report "Typography Sweep", summary & " In the Text column a middle dot stands for each space. Changes nothing.", _
        Array("Rule", "Finding", "Where", "Page", "Text"), rows
    SKW_TypographySweep = summary
End Function

Public Function SKW_LoadTeamDictionaries() As String
    Dim p As Variant, added As Long, had As Long, missing As String, cd As Object, found As Boolean, list As Collection
    Set list = SKW_Column("Dictionaries", "Path")
    If list.Count = 0 Then
        SKW_LoadTeamDictionaries = "The standards workbook lists no dictionaries. Add their full paths on the Dictionaries sheet."
        Exit Function
    End If
    For Each p In list
        If Len(Dir$(p)) = 0 Then
            missing = missing & vbCr & "  " & p
        Else
            found = False
            For Each cd In Application.CustomDictionaries
                If LCase$(cd.Path & "\" & cd.Name) = LCase$(p) Then found = True: Exit For
            Next
            If found Then
                had = had + 1
            Else
                On Error Resume Next
                Application.CustomDictionaries.Add p
                If Err.Number = 0 Then added = added + 1 Else missing = missing & vbCr & "  " & p & " (" & Err.Description & ")"
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next
    SKW_LoadTeamDictionaries = "Team dictionaries: " & added & " added, " & had & " already in Word." & _
        IIf(Len(missing) > 0, vbCr & "Not found or not added:" & missing, "") & vbCr & _
        "Word checks spelling against them from now on, on this PC. (File > Options > Proofing > Custom Dictionaries lists them.)"
End Function

Private Function SentenceOf(ByVal r As Object) As String
    Dim x As Object
    On Error Resume Next
    Set x = r.Duplicate
    x.Expand wdSentence
    SentenceOf = x.Text
End Function
