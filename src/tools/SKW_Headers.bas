Attribute VB_Name = "SKW_Headers"
' =====================================================================
' Stationkeeper Tools: headers and footers by page size (1.12.0).
'   HeaderFooterMap             report: every section's page size and orientation, what its header and
'                               footer hold, and which follow the section before
'   ReplaceHeadersByPageSize    gives every section of one page size and orientation the same header and
'                               footer, from the section at the cursor or from a Quick Part, in a copy
' Headers and footers belong to sections, and so do page size and orientation. The replace changes what
' the headers and footers hold and nothing else: no section break is added, removed or moved, and no
' section's page number settings are touched. A section of another page size that follows one about to
' change is unlinked first, so it keeps what it shows now.
' =====================================================================
Option Explicit

' A length in points as inches, with a period whatever the PC's settings say.
Private Function InchText(ByVal pts As Single) As String
    InchText = Trim$(Str$(Round(pts / 72, 2)))
End Function

' A section's page size and orientation as words: "Letter portrait", "11 x 17 landscape".
Private Function KindOf(ByVal sec As Object) As String
    Dim w As Single, h As Single, a As Single, b As Single, paper As String
    On Error Resume Next
    w = sec.PageSetup.PageWidth
    h = sec.PageSetup.PageHeight
    If w <= h Then a = w: b = h Else a = h: b = w
    If Abs(a - 612) < 2 And Abs(b - 792) < 2 Then
        paper = "Letter"
    ElseIf Abs(a - 792) < 2 And Abs(b - 1224) < 2 Then
        paper = "11 x 17"
    ElseIf Abs(a - 612) < 2 And Abs(b - 1008) < 2 Then
        paper = "Legal"
    ElseIf Abs(a - 595.3) < 2 And Abs(b - 841.9) < 2 Then
        paper = "A4"
    ElseIf Abs(a - 841.9) < 2 And Abs(b - 1190.5) < 2 Then
        paper = "A3"
    Else
        paper = InchText(a) & " x " & InchText(b) & " in"
    End If
    KindOf = paper & IIf(w > h, " landscape", " portrait")
End Function

' The page kinds a document has, each with how many sections use it: "Letter portrait" -> 3.
Private Function Kinds(ByVal d As Object) As Object
    Dim res As Object, s As Object, k As String
    Set res = CreateObject("Scripting.Dictionary")
    res.CompareMode = 1
    For Each s In d.Sections
        k = KindOf(s)
        res(k) = res(k) + 1
    Next
    Set Kinds = res
End Function

Public Function SKW_PageKinds() As String
    Dim k As Object, v As Variant, s As String
    If Documents.Count = 0 Then Exit Function
    Set k = Kinds(ActiveDocument)
    For Each v In k.Keys
        s = s & IIf(Len(s) > 0, "|", "") & v & "=" & k(v)
    Next
    SKW_PageKinds = s
End Function

Private Function HfText(ByVal hf As Object) As String
    Dim s As String
    On Error Resume Next
    s = hf.Range.Text
    s = Replace(Replace(Replace(s, vbCr, " / "), vbTab, " "), Chr$(7), "")
    Do While Right$(s, 3) = " / "
        s = Left$(s, Len(s) - 3)
    Loop
    If Len(s) = 0 Then
        If hf.Shapes.Count > 0 Then s = "(" & hf.Shapes.Count & " floating object(s))" Else s = "(empty)"
    End If
    HfText = s
End Function

Private Function LinkWord(ByVal hf As Object, ByVal secIndex As Long) As String
    On Error Resume Next
    If secIndex = 1 Then
        LinkWord = "own"
    ElseIf hf.LinkToPrevious Then
        LinkWord = "same as previous"
    Else
        LinkWord = "own"
    End If
End Function

Private Function FirstPageOf(ByVal sec As Object) As Long
    Dim r As Object
    On Error Resume Next
    Set r = sec.Range.Duplicate
    r.Collapse wdCollapseStart
    FirstPageOf = r.Information(wdActiveEndPageNumber)
End Function

Private Function NumberingOf(ByVal sec As Object) As String
    NumberingOf = "continues"
    On Error Resume Next
    If sec.Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection Then
        NumberingOf = "restarts at " & sec.Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber
    End If
End Function

' =====================================================================
' The map
' =====================================================================

' Report of every section's page size and orientation, what its header and footer hold, which follow the section before, and where page numbers restart. Changes nothing.
Public Sub HeaderFooterMap()
    SKW_Show SKW_HeaderMap()
End Sub

Public Function SKW_HeaderMap() As String
    Dim d As Object, msg As String, rows As New Collection, s As Object, i As Long, k As Object, v As Variant, kindList As String, ps As Object, wasSaved As Boolean
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Header and footer map", False)
    If Len(msg) > 0 Then SKW_HeaderMap = msg: Exit Function
    Set d = ActiveDocument
    wasSaved = d.Saved
    SKW_Steps d.Sections.Count, "Section"
    For Each s In d.Sections
        SKW_Step
        i = i + 1
        Set ps = s.PageSetup
        rows.Add SKW_Row(i, FirstPageOf(s), KindOf(s), LinkWord(s.Headers(wdHeaderFooterPrimary), i), SKW_Excerpt(HfText(s.Headers(wdHeaderFooterPrimary)), 60), _
            LinkWord(s.Footers(wdHeaderFooterPrimary), i), SKW_Excerpt(HfText(s.Footers(wdHeaderFooterPrimary)), 60), _
            IIf(ps.DifferentFirstPageHeaderFooter, "Yes", "No") & " / " & IIf(ps.OddAndEvenPagesHeaderFooter, "Yes", "No"), NumberingOf(s))
    Next
    Set k = Kinds(d)
    For Each v In k.Keys
        kindList = kindList & IIf(Len(kindList) > 0, ", ", "") & v & " (" & k(v) & ")"
    Next
    ' Reading a header or footer a section has never used makes Word note it in the file, which marks the
    ' document changed though nothing a person sees has. The map changes nothing, so the mark is put back.
    If wasSaved And Not d.Saved Then d.Saved = True
    SKW_Finish
    msg = d.Sections.Count & " section(s), " & k.Count & " page size(s): " & kindList & "."
    SKW_Report "Header and Footer Map", msg & " ""Same as previous"" means the header or footer follows the section before (Link to Previous). " & _
        "Replace Headers and Footers (by Page Size) changes every section of one page size together. Changes nothing.", _
        Array("Section", "Starts on page", "Page size", "Header", "Header holds", "Footer", "Footer holds", "Different first / odd-even", "Page numbers"), rows
    SKW_HeaderMap = msg
End Function

' =====================================================================
' The replace
' =====================================================================
' src: "section" (the section at the cursor) or "qp:<header Quick Part>|<footer Quick Part>" (either may be blank).
' what: "both", "header" or "footer".

Private Function FindBlock(ByVal nm As String) As Object
    Dim t As Object, i As Long
    If Len(Trim$(nm)) = 0 Then Exit Function
    On Error Resume Next
    Templates.LoadBuildingBlocks
    For Each t In Templates
        For i = 1 To t.BuildingBlockEntries.Count
            If LCase$(t.BuildingBlockEntries(i).Name) = LCase$(Trim$(nm)) Then Set FindBlock = t.BuildingBlockEntries(i): Exit Function
        Next
    Next
End Function

' Is this header or footer one the replace works on, for a source that is a section?
Private Function Wanted(ByVal sec As Object, ByVal srcSec As Object, ByVal idx As Long, ByRef why As String) As Boolean
    why = ""
    Select Case idx
        Case wdHeaderFooterPrimary
            Wanted = True
        Case wdHeaderFooterFirstPage
            If Not sec.PageSetup.DifferentFirstPageHeaderFooter Then Exit Function
            If srcSec Is Nothing Then why = "its first-page one is left as it is (a Quick Part fills the main one)": Exit Function
            If Not srcSec.PageSetup.DifferentFirstPageHeaderFooter Then why = "its first-page one is left as it is (the source has none)": Exit Function
            Wanted = True
        Case wdHeaderFooterEvenPages
            If Not sec.PageSetup.OddAndEvenPagesHeaderFooter Then Exit Function
            If srcSec Is Nothing Then why = "its even-page one is left as it is (a Quick Part fills the main one)": Exit Function
            Wanted = True
    End Select
End Function

' Puts the source's content into a header or footer. Returns "" or what could not be carried.
Private Function Fill(ByVal hf As Object, ByVal srcHf As Object, ByVal block As Object) As String
    Dim s As Object, t As Object, r As Object, n As Long
    On Error GoTo Fail
    If Not block Is Nothing Then
        Set t = hf.Range
        block.Insert t, True
        ' A Quick Part often brings one empty paragraph after it.
        n = hf.Range.Paragraphs.Count
        If n > 1 Then
            If Len(hf.Range.Paragraphs(n).Range.Text) <= 1 Then
                Set r = hf.Range.Paragraphs(n - 1).Range
                r.SetRange r.End - 1, r.End
                r.Delete
            End If
        End If
    Else
        ' Everything but the last paragraph mark, then the last paragraph's own style and layout.
        Set s = srcHf.Range
        s.MoveEnd wdCharacter, -1
        Set t = hf.Range
        t.MoveEnd wdCharacter, -1
        If s.End > s.Start Then
            t.FormattedText = s.FormattedText
        ElseIf t.End > t.Start Then
            t.Delete
        End If
        On Error Resume Next
        hf.Range.Paragraphs.Last.Style = srcHf.Range.Paragraphs.Last.Style
        hf.Range.Paragraphs.Last.Format = srcHf.Range.Paragraphs.Last.Format
        On Error GoTo Fail
        If srcHf.Shapes.Count > 0 And hf.Shapes.Count < srcHf.Shapes.Count Then
            Fill = (srcHf.Shapes.Count - hf.Shapes.Count) & " floating object(s) of the source (a text box, a picture set to float) were not carried"
        End If
    End If
    Exit Function
Fail:
    Fill = "could not be filled (" & Err.Description & ")"
End Function

Private Function SourceWords(ByVal src As String, ByVal srcIndex As Long) As String
    Dim a As Variant
    If LCase$(Left$(src, 3)) = "qp:" Then
        a = Split(Mid$(src, 4) & "|", "|")
        SourceWords = "the Quick Part " & IIf(Len(Trim$(a(0))) > 0, """" & Trim$(a(0)) & """ for the header", "") & _
            IIf(Len(Trim$(a(0))) > 0 And Len(Trim$(a(1))) > 0, " and ", "") & IIf(Len(Trim$(a(1))) > 0, """" & Trim$(a(1)) & """ for the footer", "")
    Else
        SourceWords = "the header and footer of section " & srcIndex
    End If
End Function

' The work, for the check (doIt False: nothing is changed) and for the copy (doIt True).
' Returns rows for the report in rows, and counts: nTarget (sections of the kind), nChange, nUnlink.
Private Function Work(ByVal d As Object, ByVal kind As String, ByVal src As String, ByVal what As String, ByVal srcIndex As Long, ByVal doIt As Boolean, _
                      ByVal rows As Collection, ByRef nTarget As Long, ByRef nChange As Long, ByRef nUnlink As Long, ByRef notes As String) As String
    Dim n As Long, i As Long, isT() As Boolean, srcSec As Object, bH As Object, bF As Object, a As Variant, idx As Variant
    Dim sec As Object, doH As Boolean, doF As Boolean, why As String, s As String, one As String, part As Variant, hf As Object, sHf As Object, blk As Object
    Dim changed As Boolean, isHeader As Boolean, act As String
    n = d.Sections.Count
    ReDim isT(1 To n)
    doH = (LCase$(what) <> "footer")
    doF = (LCase$(what) <> "header")
    If LCase$(Left$(src, 3)) = "qp:" Then
        a = Split(Mid$(src, 4) & "|", "|")
        If doH Then Set bH = FindBlock(CStr(a(0)))
        If doF Then Set bF = FindBlock(CStr(a(1)))
        If doH And bH Is Nothing And Len(Trim$(a(0))) > 0 Then Work = "No Quick Part named """ & Trim$(a(0)) & """ was found in the templates Word has loaded. Nothing was changed.": Exit Function
        If doF And bF Is Nothing And Len(Trim$(a(1))) > 0 Then Work = "No Quick Part named """ & Trim$(a(1)) & """ was found in the templates Word has loaded. Nothing was changed.": Exit Function
        If bH Is Nothing Then doH = False
        If bF Is Nothing Then doF = False
        If Not doH And Not doF Then Work = "Name a Quick Part for the header, the footer, or both. Nothing was changed.": Exit Function
        srcIndex = 0
    Else
        If srcIndex < 1 Or srcIndex > n Then Work = "Put the cursor in the section whose header and footer are the ones to use. Nothing was changed.": Exit Function
        Set srcSec = d.Sections(srcIndex)
    End If
    For i = 1 To n
        isT(i) = (LCase$(KindOf(d.Sections(i))) = LCase$(Trim$(kind))) And i <> srcIndex
        If LCase$(KindOf(d.Sections(i))) = LCase$(Trim$(kind)) Then nTarget = nTarget + 1
    Next
    If nTarget = 0 Then Work = "No section is " & kind & ". Nothing was changed.": Exit Function
    ' 1. Protect the neighbours, and free the sections to change: a header or footer that follows the
    '    section before is unlinked, which keeps what it shows now.
    For i = 2 To n
        If isT(i) Or isT(i - 1) Then
            Set sec = d.Sections(i)
            For Each part In Array("h", "f")
                If (part = "h" And doH) Or (part = "f" And doF) Then
                    For Each idx In Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)
                        If part = "h" Then Set hf = sec.Headers(idx) Else Set hf = sec.Footers(idx)
                        If hf.LinkToPrevious Then
                            If Not isT(i) Then
                                If idx = wdHeaderFooterPrimary Then
                                    nUnlink = nUnlink + 1
                                    rows.Add SKW_Row(i, FirstPageOf(sec), KindOf(sec), IIf(part = "h", "Header", "Footer"), SKW_Excerpt(HfText(hf), 50), _
                                        "Keeps what it shows: it follows a section that changes, so it is unlinked first")
                                End If
                            End If
                            If doIt Then hf.LinkToPrevious = False
                        End If
                    Next
                End If
            Next
        End If
    Next
    ' 2. Replace.
    For i = 1 To n
        If isT(i) Then
            Set sec = d.Sections(i)
            changed = False
            For Each part In Array("h", "f")
                isHeader = (part = "h")
                If (isHeader And doH) Or (Not isHeader And doF) Then
                    For Each idx In Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)
                        If Wanted(sec, srcSec, CLng(idx), why) Then
                            If isHeader Then Set hf = sec.Headers(idx) Else Set hf = sec.Footers(idx)
                            Set sHf = Nothing
                            Set blk = Nothing
                            If srcSec Is Nothing Then
                                If isHeader Then Set blk = bH Else Set blk = bF
                                act = "Takes the Quick Part " & blk.Name
                            Else
                                If isHeader Then Set sHf = srcSec.Headers(idx) Else Set sHf = srcSec.Footers(idx)
                                act = "Takes section " & srcIndex & "'s"
                                If HfText(hf) = HfText(sHf) And hf.Shapes.Count = sHf.Shapes.Count Then act = "Reads the same already; set again so it is its own"
                            End If
                            If idx = wdHeaderFooterPrimary Then
                                rows.Add SKW_Row(i, FirstPageOf(sec), KindOf(sec), IIf(isHeader, "Header", "Footer"), SKW_Excerpt(HfText(hf), 50), act)
                            End If
                            If doIt Then
                                one = Fill(hf, sHf, blk)
                                If Len(one) > 0 Then notes = notes & IIf(Len(notes) > 0, "; ", "") & "Section " & i & " " & IIf(isHeader, "header", "footer") & ": " & one
                            End If
                            changed = True
                        ElseIf Len(why) > 0 And InStr(notes, "Section " & i & ": " & why) = 0 Then
                            notes = notes & IIf(Len(notes) > 0, "; ", "") & "Section " & i & ": " & why
                        End If
                    Next
                End If
            Next
            If changed Then nChange = nChange + 1
        End If
    Next
End Function

' Pages and where each section starts, to prove pagination is as it was.
Private Function Proof(ByVal d As Object) As String
    Dim s As Object, t As String
    On Error Resume Next
    d.Repaginate
    t = d.ComputeStatistics(wdStatisticPages) & " page(s); sections start on"
    For Each s In d.Sections
        t = t & " " & FirstPageOf(s)
    Next
    Proof = t
End Function

' Gives every section of one page size and orientation the same header and footer: the ones of the section at the cursor, or a Quick Part. Reports first, then works on a copy. No section break or page number setting is changed, and the page count is compared before and after.
Public Sub ReplaceHeadersByPageSize()
    Dim k As Object, v As Variant, s As String, i As Long, pick As String, kind As String, src As String, what As String, names() As String
    Dim here As String, dflt As Long, h As String, f As String, msg As String
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    If SKW_IsQuiet() Then Exit Sub
    Set k = Kinds(ActiveDocument)
    here = KindOf(Selection.Sections(1))
    ReDim names(1 To k.Count)
    For Each v In k.Keys
        i = i + 1
        names(i) = CStr(v)
        If LCase$(CStr(v)) = LCase$(here) Then dflt = i
        s = s & i & "  " & v & " (" & k(v) & " section(s))" & vbCr
    Next
    pick = InputBox("Which page size gets the new header and footer?" & vbCr & vbCr & s & vbCr & "Type the number.", "Replace Headers and Footers", CStr(dflt))
    If Len(pick) = 0 Then Exit Sub
    If Not IsNumeric(pick) Then SKW_Say "Type a number from the list.": Exit Sub
    If CLng(pick) < 1 Or CLng(pick) > k.Count Then SKW_Say "Type a number from the list.": Exit Sub
    kind = names(CLng(pick))
    pick = InputBox("Where do the new header and footer come from?" & vbCr & vbCr & _
        "1  The section the cursor is in (section " & Selection.Sections(1).Index & ", " & here & ")" & vbCr & _
        "2  A Quick Part, by name" & vbCr & vbCr & "Type the number.", "Replace Headers and Footers", "1")
    If Len(pick) = 0 Then Exit Sub
    If pick = "2" Then
        h = InputBox("The name of the Quick Part for the header." & vbCr & "Leave it blank to leave headers as they are.", "Replace Headers and Footers")
        f = InputBox("The name of the Quick Part for the footer." & vbCr & "Leave it blank to leave footers as they are.", "Replace Headers and Footers")
        If Len(Trim$(h)) = 0 And Len(Trim$(f)) = 0 Then Exit Sub
        src = "qp:" & Trim$(h) & "|" & Trim$(f)
        what = "both"
    ElseIf pick = "1" Then
        src = "section"
        pick = InputBox("What is replaced?" & vbCr & vbCr & "1  Header and footer" & vbCr & "2  Header only" & vbCr & "3  Footer only" & vbCr & vbCr & "Type the number.", _
            "Replace Headers and Footers", "1")
        Select Case pick
            Case "1": what = "both"
            Case "2": what = "header"
            Case "3": what = "footer"
            Case Else: Exit Sub
        End Select
    Else
        SKW_Say "Type 1 or 2.": Exit Sub
    End If
    SKW_Fixable = 0
    msg = SKW_HeaderReplaceCheck(kind, src, what)
    If SKW_OfferFix(msg, "The fix replaces them in a copy of the document. No section break or page number setting is changed. Your document is not changed.") Then
        SKW_Show SKW_HeaderReplaceCopy("", kind, src, what)
    End If
End Sub

Public Function SKW_HeaderReplaceCheck(ByVal kind As String, ByVal src As String, ByVal what As String) As String
    Dim d As Object, msg As String, rows As New Collection, nT As Long, nC As Long, nU As Long, notes As String, si As Long, e As String, wasSaved As Boolean
    Set SKW_LastReport = Nothing
    SKW_Fixable = 0
    msg = SKW_Begin("Replace headers and footers", False)
    If Len(msg) > 0 Then SKW_HeaderReplaceCheck = msg: Exit Function
    Set d = ActiveDocument
    wasSaved = d.Saved
    si = Selection.Sections(1).Index
    e = Work(d, kind, src, what, si, False, rows, nT, nC, nU, notes)
    If wasSaved And Not d.Saved Then d.Saved = True
    SKW_Finish
    If Len(e) > 0 Then SKW_HeaderReplaceCheck = e: Exit Function
    SKW_Fixable = nC
    msg = nT & " section(s) are " & kind & ": " & nC & " would take " & SourceWords(src, si) & ", " & nU & " neighbouring header(s) or footer(s) are unlinked first so they keep what they show."
    SKW_Report "Replace Headers and Footers", msg & IIf(LCase$(Left$(src, 3)) <> "qp:", " Section " & si & " is the source and is not changed.", "") & _
        IIf(Len(notes) > 0, " Left as they are: " & notes & ".", "") & _
        " Only what headers and footers hold is replaced: no section break is added, removed or moved, and no section's page number settings are touched. " & _
        "A header or footer taller than the one it replaces can still move text to the next page: the fix compares the page count and each section's first page, before and after. Changes nothing.", _
        Array("Section", "Starts on page", "Page size", "Part", "Holds now", "What happens"), rows
    SKW_HeaderReplaceCheck = msg
End Function

Public Function SKW_HeaderReplaceCopy(ByVal target As String, ByVal kind As String, ByVal src As String, ByVal what As String) As String
    Dim d As Object, c As Object, fso As Object, n As Long, base As String, ext As String, rows As New Collection
    Dim nT As Long, nC As Long, nU As Long, notes As String, si As Long, e As String, before As String, after As String, ed As String
    Set SKW_LastReport = Nothing
    If Documents.Count = 0 Then SKW_HeaderReplaceCopy = "Open a document first.": Exit Function
    Set d = ActiveDocument
    If Len(d.Path) = 0 Then SKW_HeaderReplaceCopy = "Save the document first. The copy is made from the saved file.": Exit Function
    If Not d.Saved Then SKW_HeaderReplaceCopy = "Save the document first. The copy is made from the saved file, so unsaved changes would be left out.": Exit Function
    si = Selection.Sections(1).Index
    If Len(target) = 0 Then
        base = d.Path & "\" & SKW_Stem(d.Name) & " (headers replaced)"
        ext = Mid$(d.Name, InStrRev(d.Name, "."))
        target = base & ext
        Do While Len(Dir$(target)) > 0
            n = n + 1
            target = base & " " & (n + 1) & ext
        Loop
    End If
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile d.FullName, target, True
    SKW_ProgressTitle "Headers and footers: replacing in a copy"
    SKW_Status "Opening the copy..."
    Set c = Documents.Open(FileName:=target, AddToRecentFiles:=False, Visible:=True)
    c.TrackRevisions = False
    SKW_Status "Counting the pages..."
    before = Proof(c)
    Application.ScreenUpdating = False
    SKW_Status "Replacing..."
    e = Work(c, kind, src, what, si, True, rows, nT, nC, nU, notes)
    Application.ScreenUpdating = True
    If Len(e) > 0 Then
        c.Close SaveChanges:=0
        fso.DeleteFile target, True
        Application.StatusBar = False
        SKW_ProgressDone
        SKW_HeaderReplaceCopy = e
        Exit Function
    End If
    SKW_Status "Counting the pages again..."
    after = Proof(c)
    SKW_Status "Saving the copy..."
    c.Save
    Application.StatusBar = False
    SKW_ProgressDone
    SKW_HeaderReplaceCopy = nC & " section(s) that are " & kind & " now hold " & SourceWords(src, si) & ", in a copy: " & Mid$(target, InStrRev(target, "\") + 1) & "." & _
        IIf(nU > 0, vbCr & nU & " neighbouring header(s) or footer(s) were unlinked first and keep what they showed.", "") & vbCr & _
        IIf(before = after, "Pagination is as it was: " & after & ".", "Pagination changed. Before: " & before & ". After: " & after & ". A taller header or footer moved text; look at those sections.") & _
        IIf(Len(notes) > 0, vbCr & "Look at: " & notes & ".", "") & vbCr & "Your document is not changed. The copy is open."
    Exit Function
Fail:
    ed = Err.Description
    Application.ScreenUpdating = True
    Application.StatusBar = False
    SKW_ProgressDone
    SKW_HeaderReplaceCopy = "Stopped: " & ed & ". Your document is not changed."
End Function
