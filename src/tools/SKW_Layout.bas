Attribute VB_Name = "SKW_Layout"
' =====================================================================
' Stationkeeper Tools: page layout and flow.
'   SectionMap          report: every section's pages, orientation, margins, headers and numbering
'   KeepWithNextAudit   report: headings and captions that can be stranded at the bottom of a page
'   BreakHygiene        report: manual page breaks, empty paragraphs used as spacing, line breaks
'   TOCCheck            report: headings missing from the contents, entries with no heading, old page numbers
'   PageLimitCounter    report: pages per top-level part against the team's page limits
'   GraphicsAudit       report: pictures and shapes without alt text, too wide, floating, broken links
' =====================================================================
Option Explicit

' Lists every section with its first page, orientation, paper size, margins, columns, how it starts, whether its headers and footers follow the previous section, and where page numbering restarts. Changes nothing.
Public Sub SectionMap()
    SKW_Show SKW_SectionMap()
End Sub

' One report of what moves text between pages: headings that do not keep with the next paragraph, captions that do not keep with their table, long runs of kept paragraphs, manual page and column breaks, empty paragraphs used for spacing, line breaks ending paragraphs and body text set to start a new page. Changes nothing.
Public Sub KeepWithNextAudit()
    SKW_Show SKW_RunSet("Pagination Check", "", Array("SKW_KeepWithNextAudit|Keep with next", "SKW_BreakHygiene|Breaks and spacing"), False)
End Sub

' Compares each table of contents with the headings: headings missing from it, entries whose heading changed or went, and page numbers that are out of date. Changes nothing.
Public Sub TOCCheck()
    SKW_Show SKW_TOCCheck()
End Sub

' ---------------------------------------------------------------------
Private Function Inches(ByVal pts As Single) As String
    Inches = Format$(pts / 72, "0.00") & " in"
End Function

Private Function StartName(ByVal t As Long) As String
    Select Case t
        Case wdSectionContinuous: StartName = "Continuous"
        Case wdSectionNewPage: StartName = "Next page"
        Case wdSectionEvenPage: StartName = "Even page"
        Case wdSectionOddPage: StartName = "Odd page"
        Case wdSectionNewColumn: StartName = "New column"
        Case Else: StartName = "Other"
    End Select
End Function

Private Function PaperName(ByVal w As Single, ByVal h As Single) As String
    Dim a As Single, b As Single
    If w < h Then a = w: b = h Else a = h: b = w
    If Abs(a - 612) < 2 And Abs(b - 792) < 2 Then
        PaperName = "Letter"
    ElseIf Abs(a - 612) < 2 And Abs(b - 1008) < 2 Then
        PaperName = "Legal"
    ElseIf Abs(a - 595.3) < 2 And Abs(b - 841.9) < 2 Then
        PaperName = "A4"
    ElseIf Abs(a - 792) < 2 And Abs(b - 1224) < 2 Then
        PaperName = "Tabloid"
    Else
        PaperName = Format$(w / 72, "0.0") & " x " & Format$(h / 72, "0.0") & " in"
    End If
End Function

Private Function Linked(ByVal hf As Object) As String
    On Error Resume Next
    Linked = IIf(hf.LinkToPrevious, "same as previous", "own")
End Function

Public Function SKW_SectionMap() As String
    Dim d As Object, msg As String, rows As New Collection, s As Object, i As Long, ps As Object
    Dim numbering As String, hdr As String, ftr As String, pg As String, r As Object
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Section map", False)
    If Len(msg) > 0 Then SKW_SectionMap = msg: Exit Function
    Set d = ActiveDocument
    For Each s In d.Sections
        i = i + 1
        Set ps = s.PageSetup
        Set r = s.Range.Duplicate
        r.Collapse wdCollapseStart
        pg = SKW_Page(r)
        numbering = "continues"
        On Error Resume Next
        If s.Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection Then
            numbering = "restarts at " & s.Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber
        End If
        On Error GoTo 0
        If i = 1 Then
            hdr = "-": ftr = "-"
        Else
            hdr = Linked(s.Headers(wdHeaderFooterPrimary))
            ftr = Linked(s.Footers(wdHeaderFooterPrimary))
        End If
        rows.Add SKW_Row(i, pg, IIf(ps.Orientation = wdOrientLandscape, "Landscape", "Portrait"), PaperName(ps.PageWidth, ps.PageHeight), _
            "T " & Inches(ps.TopMargin) & ", B " & Inches(ps.BottomMargin) & ", L " & Inches(ps.LeftMargin) & ", R " & Inches(ps.RightMargin), _
            ps.TextColumns.Count, StartName(ps.SectionStart), _
            IIf(ps.DifferentFirstPageHeaderFooter, "Yes", "No") & " / " & IIf(ps.OddAndEvenPagesHeaderFooter, "Yes", "No"), _
            hdr & " / " & ftr, numbering)
    Next
    SKW_Finish
    SKW_Report "Section Map", d.Sections.Count & " section(s). Headers and footers ""same as previous"" follow the section before (Link to Previous). Changes nothing.", _
        Array("Section", "Starts on page", "Orientation", "Paper", "Margins", "Columns", "Starts", "Different first / odd-even", "Header / footer", "Page numbers"), rows
    SKW_SectionMap = d.Sections.Count & " section(s) mapped."
End Function

Private Function IsHeading(ByVal p As Object) As Boolean
    On Error Resume Next
    IsHeading = (p.Format.OutlineLevel >= 1 And p.Format.OutlineLevel <= 9)
End Function

Private Function NextIsTable(ByVal p As Object) As Boolean
    Dim n As Object
    On Error Resume Next
    Set n = p.Next
    If Not n Is Nothing Then NextIsTable = n.Range.Information(wdWithInTable)
End Function

Public Function SKW_KeepWithNextAudit() As String
    Dim d As Object, msg As String, rows As New Collection, p As Object, txt As String, capStyle As String
    Dim chain As Long, chainStart As Object, maxChain As Long, nH As Long, nC As Long, nL As Long, inTable As Boolean
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Keep with next audit", False)
    If Len(msg) > 0 Then SKW_KeepWithNextAudit = msg: Exit Function
    Set d = ActiveDocument
    capStyle = LCase$(SKW_Setting("CaptionStyle", "Caption"))
    maxChain = Val(SKW_Setting("MaxKeepWithNextChain", "8"))
    If maxChain < 2 Then maxChain = 8
    For Each p In d.Paragraphs
        inTable = p.Range.Information(wdWithInTable)
        txt = SKW_Excerpt(p.Range.Text, 80)
        If Not inTable And Len(txt) > 0 Then
            If IsHeading(p) And Not p.Format.KeepWithNext Then
                nH = nH + 1
                rows.Add SKW_Row("Heading", "Heading does not keep with the next paragraph, so it can end up alone at the bottom of a page", SKW_ParaStyleName(p.Range), SKW_Page(p.Range), txt)
            ElseIf LCase$(SKW_ParaStyleName(p.Range)) = capStyle And NextIsTable(p) And Not p.Format.KeepWithNext Then
                nC = nC + 1
                rows.Add SKW_Row("Caption", "Caption above a table does not keep with the table", capStyle, SKW_Page(p.Range), txt)
            End If
        End If
        If p.Format.KeepWithNext And Not inTable Then
            If chain = 0 Then Set chainStart = p
            chain = chain + 1
        Else
            If chain > maxChain Then
                nL = nL + 1
                rows.Add SKW_Row("Long chain", chain & " paragraphs in a row keep with the next one. Word moves them as a block and can leave a large gap", _
                    SKW_ParaStyleName(chainStart.Range), SKW_Page(chainStart.Range), chainStart.Range.Text)
            End If
            chain = 0
        End If
    Next
    SKW_Finish
    SKW_Report "Keep-With-Next Audit", "Headings " & nH & ", captions " & nC & ", long chains " & nL & " (more than " & maxChain & _
        " paragraphs; setting MaxKeepWithNextChain). Turn on Keep with next in the style (Paragraph > Line and Page Breaks), not paragraph by paragraph. Changes nothing.", _
        Array("Rule", "Finding", "Style", "Page", "Text"), rows
    SKW_KeepWithNextAudit = "Headings " & nH & ", captions " & nC & ", long chains " & nL & "."
End Function

Public Function SKW_BreakHygiene() As String
    Dim d As Object, msg As String, rows As New Collection, r As Object, p As Object, story As Object
    Dim nPage As Long, nCol As Long, nEmpty As Long, nLine As Long, nBefore As Long, run As Long, runStart As Object
    Dim t As String, s As Object, nSec As Long
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Break hygiene", False)
    If Len(msg) > 0 Then SKW_BreakHygiene = msg: Exit Function
    Set d = ActiveDocument
    For Each r In SKW_TextRanges(d.Content, "^m", False, False)
        ' ^m also finds section breaks, which end their section: those belong to Section Map.
        If r.End >= r.Sections(1).Range.End Then GoTo NextBreak
        nPage = nPage + 1
        If nPage <= 300 Then rows.Add SKW_Row("Page break", "Manual page break. Page break before (in the style) moves with the text", "Body", SKW_Page(r), PrevText(r))
NextBreak:
    Next
    For Each r In SKW_TextRanges(d.Content, "^n", False, False)
        nCol = nCol + 1
        If nCol <= 300 Then rows.Add SKW_Row("Column break", "Manual column break", "Body", SKW_Page(r), PrevText(r))
    Next
    For Each p In d.Paragraphs
        t = p.Range.Text
        If (t = vbCr) And Not p.Range.Information(wdWithInTable) Then
            If run = 0 Then Set runStart = p
            run = run + 1
        Else
            If run >= 2 Then
                nEmpty = nEmpty + 1
                rows.Add SKW_Row("Empty paragraphs", run & " empty paragraphs in a row used as spacing. Use space before or after in the style", "Body", SKW_Page(runStart.Range), PrevText(runStart.Range))
            End If
            run = 0
            If Len(t) > 1 Then
                If Right$(t, 2) = Chr$(11) & vbCr Then
                    nLine = nLine + 1
                    If nLine <= 300 Then rows.Add SKW_Row("Line break", "Paragraph ends with a manual line break (Shift+Enter)", "Body", SKW_Page(p.Range), t)
                End If
                If p.Format.PageBreakBefore And Not IsHeading(p) Then
                    nBefore = nBefore + 1
                    rows.Add SKW_Row("Page break before", "Body text set to start on a new page", SKW_ParaStyleName(p.Range), SKW_Page(p.Range), t)
                End If
            End If
        End If
    Next
    For Each s In d.Sections
        If s.Index > 1 Then
            nSec = nSec + 1
        End If
    Next
    SKW_Finish
    SKW_Report "Break Hygiene", "Manual page breaks " & nPage & ", column breaks " & nCol & ", runs of empty paragraphs " & nEmpty & _
        ", line breaks ending a paragraph " & nLine & ", body text starting a new page " & nBefore & ". Section breaks: " & nSec & _
        " (see Section Map). Changes nothing.", Array("Rule", "Finding", "Where", "Page", "Text before"), rows
    SKW_BreakHygiene = "Page breaks " & nPage & ", column breaks " & nCol & ", empty runs " & nEmpty & ", line breaks " & nLine & ", page break before " & nBefore & "."
End Function

' The text just before a break, to find it on the page.
Private Function PrevText(ByVal r As Object) As String
    Dim x As Object
    Set x = r.Duplicate
    x.Collapse wdCollapseStart
    x.MoveStart wdCharacter, -60
    PrevText = x.Text
End Function

' Entry text and page number of a TOC paragraph ("Heading text<tab>12").
Private Sub SplitEntry(ByVal t As String, ByRef txt As String, ByRef pg As String)
    Dim k As Long
    t = Replace(Replace(t, vbCr, ""), Chr$(7), "")
    k = InStrRev(t, vbTab)
    If k > 0 Then
        txt = Trim$(Left$(t, k - 1))
        pg = Trim$(Mid$(t, k + 1))
    Else
        txt = Trim$(t): pg = ""
    End If
    k = InStr(txt, vbTab)
    If k > 0 Then txt = Trim$(Mid$(txt, k + 1))   ' numbered headings: "1.2<tab>Text"
End Sub

Private Function Clean(ByVal t As String) As String
    Dim i As Long, ch As String
    t = Replace(Replace(Replace(t, vbCr, ""), Chr$(7), ""), vbTab, " ")
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If AscW(ch) >= 32 Or AscW(ch) < 0 Then Clean = Clean & ch
    Next
    Clean = LCase$(Trim$(Clean))
End Function

' A range's text as it reads on the page: field results, even while the window shows field codes (Alt+F9).
Private Function Shown(ByVal r As Object) As String
    Dim x As Object
    Set x = r.Duplicate
    x.TextRetrievalMode.IncludeFieldCodes = False
    Shown = x.Text
End Function

Public Function SKW_TOCCheck() As String
    Dim d As Object, msg As String, rows As New Collection, toc As Object, p As Object, ti As Long
    Dim entries As Object, heads As Object, txt As String, pg As String, lvl As Long, k As Variant, hp As String
    Dim nMiss As Long, nGone As Long, nPage As Long, lo As Long, hi As Long, rng As Object, codes As Boolean
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("TOC check", False)
    If Len(msg) > 0 Then SKW_TOCCheck = msg: Exit Function
    Set d = ActiveDocument
    If d.TablesOfContents.Count = 0 Then
        SKW_Finish
        SKW_TOCCheck = "This document has no table of contents."
        Exit Function
    End If
    ' Read the contents as they show on the page: with field codes showing (Alt+F9) a table of contents is only its code.
    On Error Resume Next
    codes = d.ActiveWindow.View.ShowFieldCodes
    If codes Then d.ActiveWindow.View.ShowFieldCodes = False
    On Error GoTo 0
    For Each toc In d.TablesOfContents
        ti = ti + 1
        lo = toc.UpperHeadingLevel: hi = toc.LowerHeadingLevel
        If lo < 1 Then lo = 1
        If hi < lo Or hi > 9 Then hi = 9
        Set entries = CreateObject("Scripting.Dictionary")
        For Each p In toc.Range.Paragraphs
            SplitEntry Shown(p.Range), txt, pg
            If Len(txt) > 0 Then entries(Clean(txt)) = pg
        Next
        Set heads = CreateObject("Scripting.Dictionary")
        For Each p In d.Paragraphs
            If Not p.Range.InRange(toc.Range) Then
                lvl = 0
                On Error Resume Next
                lvl = p.Format.OutlineLevel
                On Error GoTo 0
                If lvl >= lo And lvl <= hi Then
                    txt = Clean(Shown(p.Range))
                    If Len(txt) > 0 Then
                        hp = SKW_Page(p.Range)
                        heads(txt) = hp
                        If Not entries.Exists(txt) Then
                            nMiss = nMiss + 1
                            rows.Add SKW_Row(ti, "Missing", "Heading (level " & lvl & ") is not in the contents", hp, Shown(p.Range))
                        ElseIf IsNumeric(entries(txt)) And IsNumeric(hp) Then
                            If CLng(entries(txt)) <> CLng(hp) Then
                                nPage = nPage + 1
                                rows.Add SKW_Row(ti, "Page number", "Contents says page " & entries(txt) & ", heading is on page " & hp, hp, Shown(p.Range))
                            End If
                        End If
                    End If
                End If
            End If
        Next
        For Each k In entries.Keys
            If Not heads.Exists(k) Then
                nGone = nGone + 1
                rows.Add SKW_Row(ti, "No heading", "Entry has no heading with this text (changed or deleted)", entries(k), k)
            End If
        Next
    Next
    On Error Resume Next
    If codes Then d.ActiveWindow.View.ShowFieldCodes = True
    On Error GoTo 0
    SKW_Finish
    SKW_Report "TOC Check", d.TablesOfContents.Count & " table(s) of contents. Headings missing " & nMiss & ", entries with no heading " & nGone & _
        ", page numbers out of date " & nPage & ". Safe Update All Fields brings them up to date. Changes nothing.", _
        Array("Contents", "Rule", "Finding", "Page", "Text"), rows
    SKW_TOCCheck = "Missing " & nMiss & ", no heading " & nGone & ", old page numbers " & nPage & "."
End Function

' ---------------------------------------------------------------------
' Page limits and graphics
' ---------------------------------------------------------------------

' Counts the pages of each top-level heading's part and of the whole document, and compares them with the limits on the Page limits sheet (Heading, Max pages) and the PageLimitTotal setting. Changes nothing.
Public Sub PageLimitCounter()
    SKW_Show SKW_PageLimitCounter()
End Sub

' Lists every picture, chart and shape: inline or floating, size, wider than the margins, scaled up, missing alt text (screen readers), linked to a file that is gone. Changes nothing.
Public Sub GraphicsAudit()
    SKW_Show SKW_GraphicsAudit()
End Sub

Private Function PageNo(ByVal r As Object) As Long
    On Error Resume Next
    PageNo = r.Information(wdActiveEndAdjustedPageNumber)
End Function

Public Function SKW_PageLimitCounter() As String
    Dim d As Object, msg As String, rows As New Collection, p As Object, limits As Object, r As Object, total As Long, lim As String
    Dim starts As New Collection, titles As New Collection, i As Long, a As Long, b As Long, pages As Long, nOver As Long, t As String
    Dim x As Object, totalLimit As String, k As Variant
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Page limit counter", False)
    If Len(msg) > 0 Then SKW_PageLimitCounter = msg: Exit Function
    Set d = ActiveDocument
    Set limits = CreateObject("Scripting.Dictionary")
    For Each r In SKW_Rows("Page limits")
        If r.Exists("heading") And r.Exists("max pages") Then
            If Len(r("heading")) > 0 Then limits(LCase$(r("heading"))) = r("max pages")
        End If
    Next
    total = d.ComputeStatistics(wdStatisticPages)
    For Each p In d.Paragraphs
        If p.Format.OutlineLevel = 1 And Not p.Range.Information(wdWithInTable) Then
            t = Trim$(Replace(Replace(p.Range.Text, vbCr, ""), Chr$(12), ""))
            If Len(t) > 0 Then starts.Add PageNo(p.Range): titles.Add t
        End If
    Next
    For i = 1 To starts.Count
        a = starts(i)
        If i < starts.Count Then
            b = starts(i + 1)
            pages = b - a + 1
            If b > a Then pages = b - a   ' the next part starts on a later page: this part ends on the page before
        Else
            b = total: pages = b - a + 1
        End If
        If pages < 1 Then pages = 1
        lim = ""
        For Each k In limits.Keys
            If LCase$(titles(i)) Like k & "*" Then lim = limits(k): Exit For
        Next
        If IsNumeric(lim) And Len(lim) > 0 Then
            If pages > CLng(lim) Then nOver = nOver + 1
            rows.Add SKW_Row(titles(i), a, pages, lim, IIf(pages > CLng(lim), "OVER by " & (pages - CLng(lim)), "OK, " & (CLng(lim) - pages) & " to spare"))
        Else
            rows.Add SKW_Row(titles(i), a, pages, "", "")
        End If
    Next
    totalLimit = SKW_Setting("PageLimitTotal", "")
    If IsNumeric(totalLimit) And Len(totalLimit) > 0 Then
        If total > CLng(totalLimit) Then nOver = nOver + 1
        rows.Add SKW_Row("Whole document", 1, total, totalLimit, IIf(total > CLng(totalLimit), "OVER by " & (total - CLng(totalLimit)), "OK, " & (CLng(totalLimit) - total) & " to spare"))
    Else
        rows.Add SKW_Row("Whole document", 1, total, "", "")
    End If
    SKW_Finish
    SKW_Report "Page Limit Counter", total & " page(s), " & starts.Count & " top-level part(s), " & nOver & " over the limit. " & _
        "A part runs from its heading's page to the page before the next top-level heading (a shared page counts once, for the earlier part). " & _
        "Limits: Page limits sheet (Heading, Max pages; a heading matches when it starts with that text) and PageLimitTotal. Check the customer's own counting rules (cover, contents, tabs). Changes nothing.", _
        Array("Part", "Starts on page", "Pages", "Limit", "Result"), rows
    SKW_PageLimitCounter = total & " page(s), " & nOver & " over the limit."
End Function

Public Function SKW_GraphicsAudit() As String
    Dim d As Object, msg As String, rows As New Collection, sh As Object, i As Long, kind As String, notes As String, tw As Single
    Dim nNoAlt As Long, nWide As Long, nLink As Long, nFloat As Long, src As String, alt As String
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Graphics audit", False)
    If Len(msg) > 0 Then SKW_GraphicsAudit = msg: Exit Function
    Set d = ActiveDocument
    For Each sh In d.InlineShapes
        i = i + 1
        notes = ""
        Select Case sh.Type
            Case 3: kind = "Picture"
            Case 4: kind = "Linked picture"
            Case 12: kind = "Chart"
            Case 15: kind = "SmartArt"
            Case 1: kind = "Embedded object"
            Case 2: kind = "Linked object"
            Case Else: kind = "Inline item " & sh.Type
        End Select
        If sh.Type = 6 Or sh.Type = 9 Then GoTo NextInline   ' horizontal lines
        alt = ""
        On Error Resume Next
        alt = Trim$(sh.AlternativeText)
        tw = 0
        With sh.Range.Sections(1).PageSetup
            tw = .PageWidth - .LeftMargin - .RightMargin - .Gutter
        End With
        On Error GoTo 0
        If Len(alt) = 0 Then notes = notes & "; no alt text": nNoAlt = nNoAlt + 1
        If tw > 0 And sh.Width > tw + 2 And Not sh.Range.Information(wdWithInTable) Then notes = notes & "; wider than the margins": nWide = nWide + 1
        On Error Resume Next
        If sh.ScaleWidth > 100.5 Then notes = notes & "; scaled up to " & Format$(sh.ScaleWidth, "0") & "% (may look soft)"
        src = ""
        src = sh.LinkFormat.SourceFullName
        On Error GoTo 0
        If Len(src) > 0 Then
            If Len(Dir$(src)) = 0 Then notes = notes & "; linked file not found: " & src: nLink = nLink + 1
        End If
        rows.Add SKW_Row(i, kind & " (inline)", Format$(sh.Width / 72, "0.00") & " x " & Format$(sh.Height / 72, "0.00") & " in", _
            SKW_Page(sh.Range), Mid$(notes & " ", 3), alt)
NextInline:
    Next
    For Each sh In d.Shapes
        i = i + 1
        nFloat = nFloat + 1
        notes = "; floating (text wraps around it; it can drift from its text)"
        alt = ""
        On Error Resume Next
        alt = Trim$(sh.AlternativeText)
        On Error GoTo 0
        If Len(alt) = 0 And sh.Type <> 17 Then notes = notes & "; no alt text": nNoAlt = nNoAlt + 1   ' 17 = text box
        rows.Add SKW_Row(i, ShapeKind(sh.Type) & " (floating)", Format$(sh.Width / 72, "0.00") & " x " & Format$(sh.Height / 72, "0.00") & " in", _
            SKW_Page(sh.Anchor), Mid$(notes, 3), alt)
    Next
    SKW_Finish
    SKW_Report "Graphics Audit", i & " graphic(s). No alt text " & nNoAlt & ", wider than the margins " & nWide & ", missing linked files " & nLink & _
        ", floating " & nFloat & ". Shapes in headers and footers are not listed. Changes nothing.", Array("#", "Kind", "Size", "Page", "Notes", "Alt text"), rows
    SKW_GraphicsAudit = i & " graphic(s): no alt text " & nNoAlt & ", too wide " & nWide & ", missing links " & nLink & ", floating " & nFloat & "."
End Function

Private Function ShapeKind(ByVal t As Long) As String
    Select Case t
        Case 13: ShapeKind = "Picture"
        Case 11: ShapeKind = "Linked picture"
        Case 3: ShapeKind = "Chart"
        Case 17: ShapeKind = "Text box"
        Case 6: ShapeKind = "Group"
        Case 20: ShapeKind = "Drawing canvas"
        Case 24: ShapeKind = "SmartArt"
        Case Else: ShapeKind = "Shape"
    End Select
End Function
