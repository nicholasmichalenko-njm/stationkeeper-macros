Attribute VB_Name = "SKW_Fields"
' =====================================================================
' Stationkeeper Tools: fields and cross-references.
'   SafeUpdateAllFields      updates every field, outside-file fields kept as they are
'   FindBrokenReferences     report: references, links and included files that are broken
'   FieldCodeReport          report: every field with its code and result
'   StyleCrossReferences     gives every cross-reference the team's character style, and makes it hold
' Each macro calls a SKW_ function that does the work and returns a summary.
' =====================================================================
Option Explicit

' Updates every field in every part of the document (body, headers, footers, notes, text boxes), then the tables of contents and figures. Fields that pull in outside files (LINK, INCLUDETEXT, INCLUDEPICTURE) are locked first so they keep their content.
Public Sub SafeUpdateAllFields()
    SKW_Show SKW_SafeUpdateAllFields()
End Sub

' Lists cross-references to bookmarks that are gone, fields that show "Error!", links to places that do not exist and included files that are missing. Changes nothing.
Public Sub FindBrokenReferences()
    SKW_Show SKW_FindBrokenReferences()
End Sub

' Lists every field in the document with its type, code, result and page, so you can see what is automatic and what is typed. Changes nothing.
Public Sub FieldCodeReport()
    SKW_Show SKW_FieldCodeReport()
End Sub

' ---------------------------------------------------------------------
Private Function IsOutsideFile(ByVal t As Long) As Boolean
    IsOutsideFile = (t = wdFieldLink Or t = wdFieldIncludeText Or t = wdFieldIncludePicture Or t = wdFieldInclude)
End Function

Public Function SKW_SafeUpdateAllFields() As String
    Dim d As Object, msg As String, story As Object, f As Object, locked As New Collection, i As Long
    Dim nFields As Long, nParts As Long, nErr As Long, track As Boolean, toc As Object, tof As Object
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Update all fields", True)
    If Len(msg) > 0 Then SKW_SafeUpdateAllFields = msg: Exit Function
    Set d = ActiveDocument
    On Error GoTo Fail
    track = d.TrackRevisions
    d.TrackRevisions = False
    ' 1. Lock the fields that would reach outside the document.
    For Each story In SKW_Stories(d)
        SKW_Steps story.Fields.Count, "Field"
        For Each f In story.Fields
            SKW_Step
            If IsOutsideFile(f.Type) And Not f.Locked Then
                f.Locked = True
                locked.Add f
            End If
        Next
    Next
    ' 2. Update each part, then the tables of contents and figures (they need the page numbers).
    For Each story In SKW_Stories(d)
        If story.Fields.Count > 0 Then
            nParts = nParts + 1
            nFields = nFields + story.Fields.Count
            SKW_Status "Updating fields: " & SKW_StoryName(story.StoryType) & "..."
            story.Fields.Update
        End If
    Next
    For Each toc In d.TablesOfContents
        toc.Update
    Next
    For Each tof In d.TablesOfFigures
        tof.Update
    Next
    ' 3. Unlock what this tool locked.
    For i = 1 To locked.Count
        locked(i).Locked = False
    Next
    For Each story In SKW_Stories(d)
        SKW_Steps story.Fields.Count, "Field"
        For Each f In story.Fields
            SKW_Step
            If InStr(1, f.Result.Text, "Error!", vbTextCompare) > 0 Then nErr = nErr + 1
        Next
    Next
    d.TrackRevisions = track
    SKW_Finish
    SKW_SafeUpdateAllFields = "Updated " & nFields & " field(s) in " & nParts & " part(s) of the document, plus " & _
        d.TablesOfContents.Count & " table(s) of contents and " & d.TablesOfFigures.Count & " table(s) of figures." & vbCr & _
        IIf(locked.Count > 0, locked.Count & " field(s) that pull in outside files were kept as they were." & vbCr, "") & _
        IIf(track, "Track changes was paused for the update, so field results are not marked as edits. It is back on." & vbCr, "") & _
        IIf(nErr > 0, nErr & " field(s) now show an error. Run Find Broken References to see them.", "No field shows an error.") & _
        vbCr & "One Ctrl+Z undoes the update."
    Exit Function
Fail:
    msg = Err.Description
    On Error Resume Next
    For i = 1 To locked.Count
        locked(i).Locked = False
    Next
    d.TrackRevisions = track
    SKW_Finish
    SKW_SafeUpdateAllFields = "The update stopped: " & msg & vbCr & "Press Ctrl+Z to undo what was done."
End Function

' The file a LINK, INCLUDETEXT or INCLUDEPICTURE field points to.
Private Function FieldFile(ByVal code As String) As String
    Dim a As Long, b As Long
    a = InStr(code, """")
    If a > 0 Then
        b = InStr(a + 1, code, """")
        If b > a Then FieldFile = Mid$(code, a + 1, b - a - 1)
    Else
        Dim p As Variant: p = Split(Trim$(code), " ")
        If UBound(p) >= 1 Then FieldFile = p(1)
    End If
    FieldFile = Replace(FieldFile, "\\", "\")
End Function

Private Function FileThere(ByVal p As String) As String
    ' "yes", "no", or "not checked" for web addresses.
    If LCase$(Left$(p, 4)) = "http" Then FileThere = "not checked": Exit Function
    On Error Resume Next
    If Len(Dir$(p)) > 0 Then FileThere = "yes" Else FileThere = "no"
    If Err.Number <> 0 Then FileThere = "no"
End Function

Public Function SKW_FindBrokenReferences() As String
    Dim d As Object, msg As String, story As Object, f As Object, rows As New Collection, bm As String
    Dim h As Object, code As String, p As String, st As String, showHidden As Boolean, nRef As Long, res As String
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Find broken references", False)
    If Len(msg) > 0 Then SKW_FindBrokenReferences = msg: Exit Function
    Set d = ActiveDocument
    showHidden = d.Bookmarks.ShowHidden
    d.Bookmarks.ShowHidden = True
    For Each story In SKW_Stories(d)
        st = SKW_StoryName(story.StoryType)
        SKW_Steps story.Fields.Count, "Field"
        For Each f In story.Fields
            SKW_Step
            code = Trim$(f.Code.Text)
            res = ""
            On Error Resume Next
            res = f.Result.Text
            On Error GoTo 0
            Select Case f.Type
                Case wdFieldRef, wdFieldPageRef, wdFieldNoteRef
                    nRef = nRef + 1
                    bm = SKW_FieldBookmark(f)
                    If Len(bm) > 0 And Not d.Bookmarks.Exists(bm) Then
                        rows.Add SKW_Row("Cross-reference", "Points to bookmark " & bm & ", which is not in the document", st, SKW_Page(f.Result), code)
                    ElseIf InStr(1, res, "Error!", vbTextCompare) > 0 Then
                        rows.Add SKW_Row("Cross-reference", "Shows an error", st, SKW_Page(f.Result), res)
                    End If
                Case wdFieldLink, wdFieldIncludeText, wdFieldIncludePicture, wdFieldInclude
                    p = FieldFile(Mid$(code, InStr(code & " ", " ") + 1))
                    If FileThere(p) = "no" Then
                        rows.Add SKW_Row("Outside file", SKW_FieldTypeName(f.Type) & " file not found: " & p, st, SKW_Page(f.Result), code)
                    ElseIf InStr(1, res, "Error!", vbTextCompare) > 0 Then
                        rows.Add SKW_Row("Outside file", "Shows an error", st, SKW_Page(f.Result), res)
                    End If
                Case Else
                    If InStr(1, res, "Error!", vbTextCompare) > 0 Then
                        rows.Add SKW_Row("Field error", SKW_FieldTypeName(f.Type) & " field shows an error", st, SKW_Page(f.Result), code & "  =>  " & res)
                    End If
            End Select
        Next
        SKW_Steps story.Hyperlinks.Count, "Link"
        For Each h In story.Hyperlinks
            SKW_Step
            On Error Resume Next
            bm = ""
            bm = h.SubAddress
            p = ""
            p = h.Address
            On Error GoTo 0
            If Len(p) = 0 And Len(bm) > 0 Then
                If Not d.Bookmarks.Exists(bm) Then
                    rows.Add SKW_Row("Link", "Links to bookmark " & bm & ", which is not in the document", st, SKW_Page(h.Range), h.Range.Text)
                End If
            End If
        Next
    Next
    d.Bookmarks.ShowHidden = showHidden
    SKW_Finish
    SKW_Report "Find Broken References", _
        "Checked " & nRef & " cross-reference(s), every field result, internal links and included files. " & _
        "Field results are as last updated; run Safe Update All Fields first for an up-to-date check. Changes nothing.", _
        Array("Rule", "Finding", "Where", "Page", "Text"), rows
    SKW_FindBrokenReferences = rows.Count & " broken reference(s) found."
End Function

Public Function SKW_FieldCodeReport() As String
    Dim d As Object, msg As String, story As Object, f As Object, rows As New Collection, counts As Object
    Dim nm As String, k As Variant, intro As String, res As String, total As Long
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Field code report", False)
    If Len(msg) > 0 Then SKW_FieldCodeReport = msg: Exit Function
    Set d = ActiveDocument
    Set counts = CreateObject("Scripting.Dictionary")
    For Each story In SKW_Stories(d)
        SKW_Steps story.Fields.Count, "Field"
        For Each f In story.Fields
            SKW_Step
            total = total + 1
            nm = SKW_FieldTypeName(f.Type)
            counts(nm) = counts(nm) + 1
            If rows.Count < 3000 Then
                res = ""
                On Error Resume Next
                res = f.Result.Text
                On Error GoTo 0
                rows.Add SKW_Row(nm, Trim$(f.Code.Text), res, SKW_StoryName(story.StoryType), SKW_Page(f.Code), IIf(f.Locked, "Locked", ""))
            End If
        Next
    Next
    SKW_Finish
    For Each k In counts.Keys
        intro = intro & k & " " & counts(k) & ", "
    Next
    If Len(intro) > 0 Then intro = "By type: " & Left$(intro, Len(intro) - 2) & "." Else intro = "The document has no fields."
    If total > rows.Count Then intro = intro & " The table shows the first " & rows.Count & " of " & total & "."
    SKW_Report "Field Code Report", intro & " Changes nothing.", Array("Type", "Code", "Result", "Where", "Page", "Locked"), rows
    SKW_FieldCodeReport = total & " field(s) listed."
End Function

' ---------------------------------------------------------------------
' Style Cross-References (1.19.0)
' ---------------------------------------------------------------------
' Word has no setting that gives cross-references a style. A cross-reference is a field (REF, PAGEREF or
' NOTEREF). This tool gives every one of them the team's character style, and sets the field so the style
' holds when fields update: the field's own switch \* CHARFORMAT makes its result take the look of the field,
' so the style is put on the whole field and \* MERGEFORMAT, which would hold the old look, is taken off.
' The style's name comes from the Global Standards: the Settings row cross_reference_style.
' Fields inside a table of contents or a table of figures are left alone.

Public Function SKW_XrefStyleName() As String
    Dim r As Object
    For Each r In SKW_WbRows("Settings")
        If r.Exists("name") Then
            If LCase$(Trim$(r("name"))) = "cross_reference_style" Then
                If r.Exists("value") Then SKW_XrefStyleName = Trim$(r("value"))
                Exit Function
            End If
        End If
    Next
End Function

Private Function XrefInToc(ByVal d As Object, ByVal f As Object, ByVal spans As Collection) As Boolean
    Dim v As Variant, p As Long
    On Error Resume Next
    If f.Code.StoryType <> wdMainTextStory Then Exit Function
    p = f.Code.Start
    For Each v In spans
        If p >= v(0) And p <= v(1) Then XrefInToc = True: Exit Function
    Next
End Function

' The cross-reference fields of a document, tables of contents and figures left out.
Private Function XrefFields(ByVal d As Object) As Collection
    Dim c As New Collection, story As Object, f As Object, spans As New Collection, t As Object
    On Error Resume Next
    For Each t In d.TablesOfContents
        spans.Add Array(t.Range.Start, t.Range.End)
    Next
    For Each t In d.TablesOfFigures
        spans.Add Array(t.Range.Start, t.Range.End)
    Next
    For Each story In SKW_Stories(d)
        If story.StoryType <> wdCommentsStory Then
            For Each f In story.Fields
                If f.Type = wdFieldRef Or f.Type = wdFieldPageRef Or f.Type = wdFieldNoteRef Then
                    If Not XrefInToc(d, f, spans) Then c.Add f
                End If
            Next
        End If
    Next
    Set XrefFields = c
End Function

Private Function XrefWhole(ByVal f As Object) As Object
    Dim r As Object
    Set r = f.Code.Duplicate
    r.SetRange f.Code.Start - 1, f.Result.End + 1
    Set XrefWhole = r
End Function

Private Function XrefStyleNow(ByVal f As Object) As String
    Dim st As Object
    On Error Resume Next
    Set st = f.Result.CharacterStyle
    If st Is Nothing Then XrefStyleNow = "(mixed)" Else XrefStyleNow = st.NameLocal
End Function

Private Function XrefHolds(ByVal f As Object) As Boolean
    XrefHolds = (InStr(1, f.Code.Text, "\* CHARFORMAT", vbTextCompare) > 0) And (InStr(1, f.Code.Text, "\* MERGEFORMAT", vbTextCompare) = 0)
End Function

Private Function XrefKind(ByVal f As Object) As String
    Select Case f.Type
        Case wdFieldPageRef: XrefKind = "Page number"
        Case wdFieldNoteRef: XrefKind = "Note number"
        Case Else: XrefKind = "Text or number"
    End Select
End Function

Public Function SKW_XrefStyleCheck(ByVal styleName As String) As String
    Dim d As Object, msg As String, rows As New Collection, f As Object, st As Object, n As Long, total As Long, cur As String, holds As Boolean
    Dim wasSaved As Boolean, pg As String, r As Object
    Set SKW_LastReport = Nothing
    SKW_Fixable = 0
    msg = SKW_Begin("Style cross-references: reading", False)
    If Len(msg) > 0 Then SKW_XrefStyleCheck = msg: Exit Function
    Set d = ActiveDocument
    wasSaved = d.Saved
    If Len(styleName) = 0 Then styleName = SKW_XrefStyleName()
    Set st = SKW_GetStyle(d, styleName)
    If Len(styleName) = 0 Or st Is Nothing Then
        SKW_Finish
        SKW_XrefStyleCheck = IIf(Len(styleName) = 0, "No cross-reference style is named. Put a row named cross_reference_style on the Global Standards' Settings sheet with the style's name as its value.", _
            "The style " & styleName & " is not in this document. Bring it in first (Import Team Styles).") & " Nothing was read."
        Exit Function
    End If
    If st.Type <> wdStyleTypeCharacter And st.Type <> 6 Then
        SKW_Finish
        SKW_XrefStyleCheck = styleName & " is a paragraph style. A cross-reference sits inside a paragraph, so it needs a character style. Nothing was read."
        Exit Function
    End If
    For Each f In XrefFields(d)
        total = total + 1
        cur = XrefStyleNow(f)
        holds = XrefHolds(f)
        If StrComp(cur, st.NameLocal, vbTextCompare) <> 0 Or Not holds Then
            n = n + 1
            If n <= 300 Then
                pg = ""
                On Error Resume Next
                Set r = f.Result
                pg = r.Information(wdActiveEndAdjustedPageNumber)
                On Error GoTo 0
                rows.Add SKW_Row(pg, XrefKind(f), Left$(Replace(f.Result.Text, vbCr, " "), 80), cur, IIf(holds, "Yes", "No"), _
                    IIf(StrComp(cur, st.NameLocal, vbTextCompare) <> 0, "Takes " & st.NameLocal, "Keeps its style") & IIf(holds, "", "; set to hold it when fields update"))
            End If
        End If
    Next
    If wasSaved And Not d.Saved Then d.Saved = True
    SKW_Finish
    SKW_Fixable = n
    msg = total & " cross-reference(s): " & n & " are not in " & st.NameLocal & ", or would lose it when fields update."
    If n > 0 Then
        SKW_Report "Style Cross-References", msg & " The fix puts " & st.NameLocal & " on each whole field and sets the field to keep it (\* CHARFORMAT in place of \* MERGEFORMAT), " & _
            "so the style stays when the document's fields update. Fields inside a table of contents or a table of figures are left alone. Changes nothing." & _
            IIf(n > 300, " The first 300 are listed.", ""), Array("Page", "Kind", "Shows", "Style now", "Holds on update", "What the fix does"), rows
    End If
    SKW_XrefStyleCheck = msg
End Function

Public Function SKW_XrefStyleApply(ByVal styleName As String) As String
    Dim d As Object, msg As String, f As Object, st As Object, n As Long, total As Long, t As String, p As Long, ed As String, fs As Collection
    Set SKW_LastReport = Nothing
    If Documents.Count = 0 Then SKW_XrefStyleApply = "Open a document first.": Exit Function
    Set d = ActiveDocument
    If Len(styleName) = 0 Then styleName = SKW_XrefStyleName()
    Set st = SKW_GetStyle(d, styleName)
    If Len(styleName) = 0 Or st Is Nothing Then SKW_XrefStyleApply = "The cross-reference style is not named or is not in this document. Nothing was changed.": Exit Function
    If st.Type <> wdStyleTypeCharacter And st.Type <> 6 Then SKW_XrefStyleApply = styleName & " is a paragraph style; a cross-reference needs a character style. Nothing was changed.": Exit Function
    msg = SKW_Begin("Style cross-references", True)
    If Len(msg) > 0 Then SKW_XrefStyleApply = msg: Exit Function
    On Error GoTo Fail
    Set fs = XrefFields(d)
    SKW_Steps fs.Count, "Cross-reference"
    For Each f In fs
        SKW_Step
        total = total + 1
        If StrComp(XrefStyleNow(f), st.NameLocal, vbTextCompare) <> 0 Or Not XrefHolds(f) Then
            If Not XrefHolds(f) Then
                t = f.Code.Text
                p = InStr(1, t, "\* MERGEFORMAT", vbTextCompare)
                Do While p > 0
                    t = Left$(t, p - 1) & Mid$(t, p + Len("\* MERGEFORMAT"))
                    p = InStr(1, t, "\* MERGEFORMAT", vbTextCompare)
                Loop
                If InStr(1, t, "\* CHARFORMAT", vbTextCompare) = 0 Then t = RTrim$(t) & " \* CHARFORMAT "
                f.Code.Text = t
            End If
            XrefWhole(f).Style = st
            n = n + 1
        End If
    Next
    SKW_Finish
    SKW_XrefStyleApply = n & " of " & total & " cross-reference(s) set to " & st.NameLocal & ", and set to keep it when fields update." & vbCr & "One Ctrl+Z undoes it."
    Exit Function
Fail:
    ed = Err.Description
    SKW_Finish
    SKW_XrefStyleApply = "Stopped: " & ed & ". Press Ctrl+Z to undo what was done."
End Function

' Gives every cross-reference in the document the team's cross-reference character style, and sets each one to keep it when fields update. The style's name comes from the Global Standards (Settings: cross_reference_style); without one you pick from the document's character styles. Reports first. Tables of contents and figures are left alone. One Ctrl+Z undoes it.
Public Sub StyleCrossReferences()
    Dim nm As String, msg As String, d As Object, st As Object, lines As String, names As New Collection, pick As String, i As Long, dpf As String
    If SKW_IsQuiet() Then Exit Sub
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    Set d = ActiveDocument
    nm = SKW_XrefStyleName()
    If Len(nm) = 0 Then
        ' No standard yet: pick from the character styles this document has, by number.
        dpf = d.Styles(wdStyleDefaultParagraphFont).NameLocal
        For Each st In d.Styles
            If st.Type = wdStyleTypeCharacter And st.InUse And names.Count < 40 Then
                If st.NameLocal <> dpf Then names.Add st.NameLocal: lines = lines & vbCr & names.Count & "  " & st.NameLocal
            End If
        Next
        If names.Count = 0 Then SKW_Say "This document has no character style to give its cross-references. Make or bring in the style first.": Exit Sub
        pick = InputBox("The Global Standards name no cross-reference style (Settings: cross_reference_style)." & vbCr & vbCr & _
            "Which character style do cross-references take? Type its number." & vbCr & lines, "Style Cross-References", "")
        i = Val(pick)
        If i < 1 Or i > names.Count Then Exit Sub
        nm = names(i)
    End If
    SKW_Fixable = 0
    msg = SKW_XrefStyleCheck(nm)
    If SKW_Fixable <= 0 Or SKW_LastReport Is Nothing Then SKW_Show msg: Exit Sub
    If MsgBox(msg & vbCr & vbCr & "Give the " & SKW_Fixable & " cross-reference(s) the style " & nm & " now?" & vbCr & vbCr & _
        "One Ctrl+Z undoes it. The report stays open.", vbYesNo + vbQuestion, SKW_APP) <> vbYes Then Exit Sub
    d.Activate
    SKW_Show SKW_XrefStyleApply(nm)
End Sub
