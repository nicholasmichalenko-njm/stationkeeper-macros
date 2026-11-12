Attribute VB_Name = "SKW_Fields"
' =====================================================================
' Stationkeeper Tools: fields and cross-references.
'   SafeUpdateAllFields      updates every field, outside-file fields kept as they are
'   FindBrokenReferences     report: references, links and included files that are broken
'   FieldCodeReport          report: every field with its code and result
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
        For Each f In story.Fields
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
        For Each f In story.Fields
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
        For Each f In story.Fields
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
        For Each h In story.Hyperlinks
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
        For Each f In story.Fields
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
