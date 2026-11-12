Attribute VB_Name = "SKP_Tidy"
' =====================================================================
' Stationkeeper Tools for PowerPoint, batch 3 (1.5.0): acronyms at first
' use, footer and marking, text overflow, placeholders that drifted,
' pictures, colors, proofing language, links and table styles.
' Run from Word on the deck open in PowerPoint. Checks change nothing.
' Fixes never touch the open deck: each saves "<deck> - fixed <date>.pptx"
' beside it, works on that copy, reports what it did and opens the copy.
'   DeckAcronymFirstUseFix, DeckMarkingCheck, DeckMarkingFix,
'   DeckTextOverflowCheck, DeckLayoutDriftCheck, DeckResetPlaceholders,
'   DeckPictureCheck, DeckPictureAspectFix, DeckColorCheck,
'   DeckLanguageCheck, DeckLanguageFix, DeckLinksCheck,
'   DeckTableStyleCheck, DeckTableStyleFix
' Settings (Settings sheet, or the linked effort's rules): DeckFooterText,
' DeckMarking, DeckLanguage, DeckTableStyle, DeckExtraColors.
' =====================================================================
Option Explicit

Private Const TPH_TITLE As Long = 1, TPH_BODY As Long = 2, TPH_CTITLE As Long = 3, TPH_SUBTITLE As Long = 4
Private Const TPH_OBJECT As Long = 7, TPH_NUMBER As Long = 13, TPH_FOOTER As Long = 15, TPH_DATE As Long = 16
Private Const NEAR_PT As Single = 2   ' points: closer than this counts as the same place

' Checks that every slide shows the team footer (DeckFooterText) and the marking (DeckMarking, such as CUI or Proprietary). When some do not, it offers to fix a copy: the team footer goes on every slide. The open deck is not changed.
Public Sub DeckMarkingCheck()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    Dim s As String
    SKW_Fixable = 0
    s = SKP_DeckMarkingCheck()
    If SKW_OfferFix(s, "The fix puts the footer placeholder back on masters and layouts and shows the team footer on every slide.") Then SKW_Show SKP_DeckMarkingFix("")
End Sub

' Lists text that runs past the bottom of its box or off the slide, and tables that run off the slide. Changes nothing.
Public Sub DeckTextOverflowCheck()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Show SKP_DeckTextOverflowCheck()
End Sub

' Lists titles and other placeholders moved, resized or restyled away from their layout. When it finds any, it offers to fix a copy: they go back where the layout has them. The open deck is not changed.
Public Sub DeckLayoutDriftCheck()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    Dim s As String
    SKW_Fixable = 0
    s = SKP_DeckLayoutDriftCheck()
    If SKW_OfferFix(s, "The fix moves and resizes those placeholders back to their layout, and titles back to the layout's font and size.") Then SKW_Show SKP_DeckResetPlaceholders("")
End Sub

' Lists stretched pictures, pictures enlarged past 150 percent, and cropped pictures whose cut-off parts still travel in the file. When some are stretched, it offers to fix a copy. The open deck is not changed.
Public Sub DeckPictureCheck()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    Dim s As String
    SKW_Fixable = 0
    s = SKP_DeckPictureCheck()
    If SKW_OfferFix(s, "The fix gives each stretched picture its own shape back, fitted inside the box it had. Enlarged and cropped pictures are for you to judge.") Then SKW_Show SKP_DeckPictureAspectFix("")
End Sub

' Lists fill, line and text colors that are not the theme's colors (or DeckExtraColors), with the nearest theme color. Changes nothing.
Public Sub DeckColorCheck()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Show SKP_DeckColorCheck()
End Sub

' Lists text marked with a proofing language other than the team's (DeckLanguage, en-US by default). When it finds any, it offers to fix a copy. The open deck is not changed.
Public Sub DeckLanguageCheck()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    Dim s As String
    SKW_Fixable = 0
    s = SKP_DeckLanguageCheck()
    If SKW_OfferFix(s, "The fix marks all text, and the deck's default, with the team's proofing language.") Then SKW_Show SKP_DeckLanguageFix("")
End Sub

' Lists linked pictures, objects and media, hyperlinks and embedded objects: missing files, links into someone's own folders or drive, and embedded files that carry their data along. Changes nothing.
Public Sub DeckLinksCheck()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Show SKP_DeckLinksCheck()
End Sub

' Lists tables not in the team's table style (DeckTableStyle) and tables with no header row. When some are off style, it offers to fix a copy. The open deck is not changed.
Public Sub DeckTableStyleCheck()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    Dim s As String
    SKW_Fixable = 0
    s = SKP_DeckTableStyleCheck()
    If SKW_OfferFix(s, "The fix applies the team table style to every table. Cell colors and borders set by hand are replaced by the style's.") Then SKW_Show SKP_DeckTableStyleFix("")
End Sub

' ---------------------------------------------------------------------
' Shared
' ---------------------------------------------------------------------
Private Function OpenDeck(ByRef p As Object, ByVal title As String) As String
    Set SKW_LastReport = Nothing
    Set p = SKP_Deck.SKP_Deck()
    If p Is Nothing Then OpenDeck = SKP_Why(): Exit Function
    SKW_Status title & ": " & p.Name & "..."
End Function

Private Function PhType(ByVal sh As Object) As Long
    PhType = -1
    On Error Resume Next
    If sh.Type = 14 Then PhType = sh.PlaceholderFormat.Type
End Function

Private Function IsTitlePh(ByVal sh As Object) As Boolean
    Dim t As Long
    t = PhType(sh)
    IsTitlePh = (t = TPH_TITLE Or t = TPH_CTITLE)
End Function

Private Function PhOfType(ByVal shapes As Object, ByVal ty As Long) As Object
    Dim sh As Object
    For Each sh In shapes
        If PhType(sh) = ty Then Set PhOfType = sh: Exit Function
    Next
End Function

Private Function TextOf(ByVal sh As Object) As String
    Dim tr As Object, s As String
    On Error Resume Next
    For Each tr In SKP_TextRanges(sh)
        s = s & tr.Text & " "
    Next
    TextOf = s
End Function

Private Function Pt(ByVal v As Single) As String
    Pt = Format$(Abs(v), "0") & " pt"
End Function

Private Function Hex6(ByVal c As Long) As String
    Hex6 = "#" & Right$("0" & Hex$(c And &HFF&), 2) & Right$("0" & Hex$((c \ &H100&) And &HFF&), 2) & Right$("0" & Hex$((c \ &H10000) And &HFF&), 2)
End Function

Private Function FinishReport(ByVal title As String, ByVal intro As String, ByVal heads As Variant, ByVal rows As Collection, ByVal p As Object) As Boolean
    SKW_Status ""
    SKW_ReportFor title, intro, heads, rows, p.Name
End Function

' ---------------------------------------------------------------------
' Acronyms: spell out at first use
' ---------------------------------------------------------------------
Public Function SKP_DeckAcronymFirstUseFix(ByVal target As String) As String
    Dim p As Object, c As Object, msg As String, rows As New Collection, list As Object, k As Variant, acr As String, def As String
    Dim sld As Object, sh As Object, tr As Object, f As Object, done As Object, later As Object, notes As Object, t As String
    Dim nFixed As Long, nNoDef As Long, nTitleOnly As Long, inTitle As Object
    msg = OpenDeck(p, "Deck acronyms at first use")
    If Len(msg) > 0 Then SKP_DeckAcronymFirstUseFix = msg: Exit Function
    Set list = SKW_AcronymList()
    If list.Count = 0 Then SKP_DeckAcronymFirstUseFix = "No acronyms in the team list (the Acronyms sheet or the AcronymList setting).": Exit Function
    ' Where each acronym is already spelled out, as "(ACR)".
    Set later = CreateObject("Scripting.Dictionary")
    For Each sld In p.Slides
        For Each sh In SKP_Shapes(sld)
            t = TextOf(sh)
            If InStr(t, "(") > 0 Then
                For Each k In list.Keys
                    If Not later.Exists(k) Then
                        If InStr(1, t, "(" & k & ")", vbBinaryCompare) > 0 Then later(k) = sld.SlideIndex
                    End If
                Next
            End If
        Next
    Next
    msg = SKP_FixCopy(p, "fixed", target, c)
    If Len(msg) > 0 Then SKP_DeckAcronymFirstUseFix = msg: Exit Function
    On Error GoTo Fail
    SKW_Status "Spelling out acronyms at first use..."
    Set done = CreateObject("Scripting.Dictionary")
    Set inTitle = CreateObject("Scripting.Dictionary")
    For Each sld In c.Slides
        Set notes = CreateObject("Scripting.Dictionary")
        For Each sh In SKP_Shapes(sld)
            For Each tr In SKP_TextRanges(sh)
                t = tr.Text
                For Each k In list.Keys
                    acr = CStr(k)
                    If Not done.Exists(acr) Then
                        If InStr(1, t, acr, vbBinaryCompare) > 0 Then
                            If IsTitlePh(sh) Then
                                If Not inTitle.Exists(acr) Then inTitle(acr) = sld.SlideIndex   ' titles stay short: spell out in the body
                            ElseIf InStr(1, t, "(" & acr & ")", vbBinaryCompare) > 0 Then
                                done(acr) = True
                            Else
                                Set f = Nothing
                                Set f = tr.Find(FindWhat:=acr, MatchCase:=True, WholeWords:=True)
                                If Not f Is Nothing Then
                                    def = CStr(list(k))
                                    done(acr) = True
                                    If Len(def) = 0 Then
                                        nNoDef = nNoDef + 1
                                        rows.Add SKW_Row(sld.SlideIndex, acr, "(no meaning on the team list)", "Spell it out by hand", "")
                                    Else
                                        f.Text = def & " (" & acr & ")"
                                        nFixed = nFixed + 1
                                        notes("Spelled out " & acr & " at first use") = 1
                                        rows.Add SKW_Row(sld.SlideIndex, acr, def & " (" & acr & ")", sh.Name, _
                                            IIf(later.Exists(k), "Also spelled out on slide " & later(k) & ": take that one out", ""))
                                    End If
                                    t = tr.Text
                                End If
                            End If
                        End If
                    End If
                Next
            Next
        Next
        If notes.Count > 0 Then SKP_AddComment sld, notes
    Next
    For Each k In inTitle.Keys
        If Not done.Exists(k) Then
            nTitleOnly = nTitleOnly + 1
            rows.Add SKW_Row(inTitle(k), k, "(only in titles)", "Spell it out in a title or the notes by hand", "")
        End If
    Next
    SKP_FinishCopy c
    FinishReport "Deck Acronym First Use Fix", nFixed & " acronym(s) spelled out at first use, " & nNoDef & " with no meaning on the team list, " & nTitleOnly & _
        " used only in titles. Titles are left short; the first use in body text, notes aside, gets the meaning. Each changed slide has a comment. " & _
        "Text grows: run Deck Text Overflow Check afterwards. The open deck is not changed. The fixed copy: " & SKP_LastCopy, _
        Array("Slide", "Acronym", "Now reads", "Shape", "Note"), rows, p
    SKP_DeckAcronymFirstUseFix = nFixed & " spelled out at first use, " & nNoDef & " with no meaning, " & nTitleOnly & " only in titles." & vbCr & "Fixed copy: " & SKP_LastCopy
    Exit Function
Fail:
    SKP_DeckAcronymFirstUseFix = "Acronym fix stopped: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Function

' ---------------------------------------------------------------------
' Footer and marking
' ---------------------------------------------------------------------
Private Function MarkRules(ByRef foot As String, ByRef mark As String) As String
    foot = Trim$(SKW_Setting("DeckFooterText", ""))
    mark = Trim$(SKW_Setting("DeckMarking", ""))
    If Len(foot) = 0 And Len(mark) = 0 Then MarkRules = "Set DeckFooterText (the footer every slide shows) or DeckMarking (such as CUI or Proprietary) on the Settings sheet first."
End Function

' Where the marking shows on a slide: "slide", "layout", "master" or "".
Private Function MarkingWhere(ByVal sld As Object, ByVal mark As String) As String
    Dim sh As Object, lay As Object
    For Each sh In SKP_Shapes(sld)
        If InStr(1, TextOf(sh), mark, vbTextCompare) > 0 Then MarkingWhere = "slide": Exit Function
    Next
    On Error Resume Next
    Set lay = sld.CustomLayout
    For Each sh In lay.Shapes
        If sh.Type <> 14 Then
            If InStr(1, TextOf(sh), mark, vbTextCompare) > 0 Then MarkingWhere = "layout": Exit Function
        End If
    Next
    If sld.DisplayMasterShapes And lay.DisplayMasterShapes Then
        For Each sh In lay.Design.SlideMaster.Shapes
            If sh.Type <> 14 Then
                If InStr(1, TextOf(sh), mark, vbTextCompare) > 0 Then MarkingWhere = "master": Exit Function
            End If
        Next
    End If
End Function

Public Function SKP_DeckMarkingCheck() As String
    Dim p As Object, msg As String, rows As New Collection, sld As Object, ph As Object, foot As String, mark As String, t As String
    Dim nNoFoot As Long, nWrong As Long, nNoMark As Long
    msg = OpenDeck(p, "Deck marking check")
    If Len(msg) > 0 Then SKP_DeckMarkingCheck = msg: Exit Function
    msg = MarkRules(foot, mark)
    If Len(msg) > 0 Then SKW_Status "": SKP_DeckMarkingCheck = msg: Exit Function
    For Each sld In p.Slides
        Set ph = PhOfType(sld.Shapes, TPH_FOOTER)
        If Len(foot) > 0 Then
            If ph Is Nothing Then
                nNoFoot = nNoFoot + 1
                rows.Add SKW_Row(sld.SlideIndex, SKP_SlideTitleOf(sld), "No footer shows", IIf(PhOfType(sld.CustomLayout.Shapes, TPH_FOOTER) Is Nothing, "Its layout has no footer placeholder", "Insert > Header & Footer, or Deck Marking Fix"))
            Else
                t = Trim$(Replace(ph.TextFrame.TextRange.Text, vbCr, " "))
                If StrComp(t, foot, vbBinaryCompare) <> 0 Then
                    nWrong = nWrong + 1
                    rows.Add SKW_Row(sld.SlideIndex, SKP_SlideTitleOf(sld), "Footer reads """ & SKW_Excerpt(t, 60) & """", "The team footer: """ & foot & """")
                End If
            End If
        End If
        If Len(mark) > 0 Then
            If Len(MarkingWhere(sld, mark)) = 0 Then
                nNoMark = nNoMark + 1
                rows.Add SKW_Row(sld.SlideIndex, SKP_SlideTitleOf(sld), "No """ & mark & """ marking", IIf(sld.SlideShowTransition.Hidden, "Hidden slide: it still travels in the file", "Not in the slide's text, footer, layout or master"))
            End If
        End If
    Next
    FinishReport "Deck Marking Check", p.Slides.Count & " slide(s): " & nNoFoot & " without a footer, " & nWrong & " with a different footer, " & nNoMark & " without the marking. " & _
        "Hidden slides count: they travel in the file. Changes nothing.", Array("Slide", "Title", "Problem", "Detail"), rows, p
    SKW_Fixable = nNoFoot + nWrong
    SKP_DeckMarkingCheck = nNoFoot & " slide(s) without a footer, " & nWrong & " with a different footer, " & nNoMark & " without the marking."
End Function

Public Function SKP_DeckMarkingFix(ByVal target As String) As String
    Dim p As Object, c As Object, msg As String, rows As New Collection, d As Object, lay As Object, mph As Object, sld As Object, ph As Object
    Dim foot As String, mark As String, want As String, t As String, nAdded As Long, nOn As Long, nText As Long, nNoMark As Long
    msg = OpenDeck(p, "Deck marking fix")
    If Len(msg) > 0 Then SKP_DeckMarkingFix = msg: Exit Function
    msg = MarkRules(foot, mark)
    If Len(msg) > 0 Then SKW_Status "": SKP_DeckMarkingFix = msg: Exit Function
    want = foot: If Len(want) = 0 Then want = mark
    msg = SKP_FixCopy(p, "fixed", target, c)
    If Len(msg) > 0 Then SKP_DeckMarkingFix = msg: Exit Function
    On Error GoTo Fail
    SKW_Status "Putting the footer on every slide..."
    For Each d In c.Designs
        Set mph = PhOfType(d.SlideMaster.Shapes, TPH_FOOTER)
        If mph Is Nothing Then
            On Error Resume Next
            Set mph = d.SlideMaster.Shapes.AddPlaceholder(TPH_FOOTER, c.PageSetup.SlideWidth * 0.33, c.PageSetup.SlideHeight - 40, c.PageSetup.SlideWidth * 0.34, 28)
            If Err.Number = 0 Then nAdded = nAdded + 1: rows.Add SKW_Row("", "Master " & d.Name, "Footer placeholder put back", "")
            Err.Clear
            On Error GoTo Fail
        End If
        For Each lay In d.SlideMaster.CustomLayouts
            If PhOfType(lay.Shapes, TPH_FOOTER) Is Nothing And Not mph Is Nothing Then
                On Error Resume Next
                lay.Shapes.AddPlaceholder TPH_FOOTER, mph.Left, mph.Top, mph.Width, mph.Height
                If Err.Number = 0 Then nAdded = nAdded + 1: rows.Add SKW_Row("", "Layout " & lay.Name, "Footer placeholder put back", "") _
                    Else rows.Add SKW_Row("", "Layout " & lay.Name, "Could not add the footer placeholder", Err.Description)
                Err.Clear
                On Error GoTo Fail
            End If
        Next
    Next
    For Each sld In c.Slides
        Set ph = PhOfType(sld.Shapes, TPH_FOOTER)
        If ph Is Nothing Then
            sld.HeadersFooters.Footer.Visible = 0
            sld.HeadersFooters.Footer.Visible = -1
            Set ph = PhOfType(sld.Shapes, TPH_FOOTER)
            If ph Is Nothing Then
                rows.Add SKW_Row(sld.SlideIndex, SKP_SlideTitleOf(sld), "Could not show a footer", "Its layout still has no footer placeholder")
            Else
                nOn = nOn + 1
                rows.Add SKW_Row(sld.SlideIndex, SKP_SlideTitleOf(sld), "Footer turned on", "")
            End If
        End If
        If Not ph Is Nothing Then
            t = Trim$(Replace(ph.TextFrame.TextRange.Text, vbCr, " "))
            If StrComp(t, want, vbBinaryCompare) <> 0 Then
                sld.HeadersFooters.Footer.Text = want
                If ph.TextFrame.TextRange.Text <> want Then ph.TextFrame.TextRange.Text = want
                nText = nText + 1
                rows.Add SKW_Row(sld.SlideIndex, SKP_SlideTitleOf(sld), "Footer text set", IIf(Len(t) > 0, "Was """ & SKW_Excerpt(t, 50) & """", "Was empty"))
            End If
        End If
        If Len(mark) > 0 Then
            If Len(MarkingWhere(sld, mark)) = 0 Then
                nNoMark = nNoMark + 1
                rows.Add SKW_Row(sld.SlideIndex, SKP_SlideTitleOf(sld), "Still no """ & mark & """ marking", "Put it in DeckFooterText, or on the template's master")
            End If
        End If
    Next
    SKP_FinishCopy c
    FinishReport "Deck Marking Fix", nAdded & " footer placeholder(s) put back, " & nOn & " footer(s) turned on, " & nText & " footer(s) set to """ & want & """, " & nNoMark & _
        " slide(s) still without the marking. The open deck is not changed. The fixed copy: " & SKP_LastCopy, Array("Slide", "Title", "Change", "Detail"), rows, p
    SKP_DeckMarkingFix = nAdded & " placeholder(s) back, " & nOn & " on, " & nText & " set, " & nNoMark & " still unmarked." & vbCr & "Fixed copy: " & SKP_LastCopy
    Exit Function
Fail:
    SKP_DeckMarkingFix = "Marking fix stopped: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Function

' ---------------------------------------------------------------------
' Text that does not fit
' ---------------------------------------------------------------------
Public Function SKP_DeckTextOverflowCheck() As String
    Dim p As Object, msg As String, rows As New Collection, sld As Object, sh As Object, tr As Object
    Dim sw As Single, shh As Single, bot As Single, rgt As Single, nOver As Long, nOff As Long
    msg = OpenDeck(p, "Deck text overflow")
    If Len(msg) > 0 Then SKP_DeckTextOverflowCheck = msg: Exit Function
    sw = p.PageSetup.SlideWidth: shh = p.PageSetup.SlideHeight
    For Each sld In p.Slides
        For Each sh In SKP_Shapes(sld)
            If sh.Type <> 6 Then
                If sh.HasTable Then
                    If sh.Top + sh.Height > shh + NEAR_PT Then
                        nOff = nOff + 1
                        rows.Add SKW_Row(sld.SlideIndex, sh.Name, "Table runs off the slide", Pt(sh.Top + sh.Height - shh) & " below the bottom edge", "")
                    End If
                ElseIf sh.HasTextFrame Then
                    If sh.TextFrame.HasText Then
                        Set tr = sh.TextFrame.TextRange
                        bot = tr.BoundTop + tr.BoundHeight
                        rgt = tr.BoundLeft + tr.BoundWidth
                        If bot > shh + NEAR_PT Or rgt > sw + NEAR_PT Then
                            nOff = nOff + 1
                            rows.Add SKW_Row(sld.SlideIndex, sh.Name, "Text runs off the slide", IIf(bot > shh + NEAR_PT, Pt(bot - shh) & " below the bottom edge", Pt(rgt - sw) & " past the right edge"), SKW_Excerpt(tr.Text, 60))
                        ElseIf bot > sh.Top + sh.Height + NEAR_PT Then
                            nOver = nOver + 1
                            rows.Add SKW_Row(sld.SlideIndex, sh.Name, "Text runs past the bottom of its box", Pt(bot - sh.Top - sh.Height) & " over", SKW_Excerpt(tr.Text, 60))
                        End If
                    End If
                End If
            End If
        Next
    Next
    FinishReport "Deck Text Overflow Check", nOver & " text box(es) overflow, " & nOff & " run off the slide. Text PowerPoint shrinks to fit is not counted. Changes nothing.", _
        Array("Slide", "Shape", "Problem", "By", "Text"), rows, p
    SKP_DeckTextOverflowCheck = nOver & " overflow(s), " & nOff & " off the slide."
End Function

' ---------------------------------------------------------------------
' Placeholders that drifted from their layout
' ---------------------------------------------------------------------
Private Function KindOf(ByVal ty As Long) As String
    Select Case ty
        Case TPH_TITLE, TPH_CTITLE: KindOf = "Title"
        Case TPH_SUBTITLE: KindOf = "Subtitle"
        Case TPH_BODY, TPH_OBJECT, 8, 12, 18, 6: KindOf = "Content"
        Case TPH_NUMBER, TPH_FOOTER, TPH_DATE, -1: KindOf = ""
        Case Else: KindOf = "Placeholder " & ty
    End Select
End Function

' The layout placeholder a slide placeholder comes from: same kind, same order.
Private Function LayoutTwin(ByVal sld As Object, ByVal ph As Object, ByRef label As String) As Object
    Dim kind As String, n As Long, sh As Object, i As Long
    kind = KindOf(PhType(ph))
    If Len(kind) = 0 Then Exit Function
    For Each sh In sld.Shapes
        If KindOf(PhType(sh)) = kind Then n = n + 1
        If sh.Id = ph.Id Then Exit For
    Next
    If n = 0 Then Exit Function
    label = kind & IIf(n > 1, " " & n, "")
    For Each sh In sld.CustomLayout.Shapes
        If KindOf(PhType(sh)) = kind Then
            i = i + 1
            If i = n Then Set LayoutTwin = sh: Exit Function
        End If
    Next
End Function

Private Function FirstRunFont(ByVal sh As Object, ByRef sz As Single) As String
    sz = 0
    On Error Resume Next
    sz = sh.TextFrame.TextRange.Characters(1, 1).Font.Size
    FirstRunFont = sh.TextFrame.TextRange.Characters(1, 1).Font.Name
End Function

' What changed between a placeholder and its layout twin, or "".
Private Function Drift(ByVal ph As Object, ByVal lph As Object, ByVal isTitle As Boolean) As String
    Dim s As String, a As Single, b As Single, fa As String, fb As String
    If Abs(ph.Left - lph.Left) > NEAR_PT Then s = s & ", " & Pt(ph.Left - lph.Left) & IIf(ph.Left > lph.Left, " right", " left")
    If Abs(ph.Top - lph.Top) > NEAR_PT Then s = s & ", " & Pt(ph.Top - lph.Top) & IIf(ph.Top > lph.Top, " down", " up")
    If Abs(ph.Width - lph.Width) > NEAR_PT Then s = s & ", " & Pt(ph.Width - lph.Width) & IIf(ph.Width > lph.Width, " wider", " narrower")
    If Abs(ph.Height - lph.Height) > NEAR_PT Then s = s & ", " & Pt(ph.Height - lph.Height) & IIf(ph.Height > lph.Height, " taller", " shorter")
    If isTitle And ph.TextFrame.HasText Then
        fa = FirstRunFont(ph, a): fb = FirstRunFont(lph, b)
        If a > 0 And b > 0 And Abs(a - b) > 0.5 Then s = s & ", text " & a & " pt (layout " & b & " pt)"
        If Len(fa) > 0 And Len(fb) > 0 And StrComp(fa, fb, vbTextCompare) <> 0 Then s = s & ", font " & fa & " (layout " & fb & ")"
    End If
    If Len(s) > 0 Then Drift = Mid$(s, 3)
End Function

Public Function SKP_DeckLayoutDriftCheck() As String
    Dim p As Object, msg As String, rows As New Collection, sld As Object, sh As Object, lph As Object, label As String, d As String, nDrift As Long, nTitle As Long
    msg = OpenDeck(p, "Deck layout drift")
    If Len(msg) > 0 Then SKP_DeckLayoutDriftCheck = msg: Exit Function
    For Each sld In p.Slides
        For Each sh In sld.Shapes
            If sh.Type = 14 Then
                label = ""
                Set lph = LayoutTwin(sld, sh, label)
                If Not lph Is Nothing Then
                    d = Drift(sh, lph, IsTitlePh(sh))
                    If Len(d) > 0 Then
                        nDrift = nDrift + 1
                        If IsTitlePh(sh) Then nTitle = nTitle + 1
                        rows.Add SKW_Row(sld.SlideIndex, label, d, sld.CustomLayout.Name, SKW_Excerpt(TextOf(sh), 50))
                    End If
                End If
            End If
        Next
    Next
    FinishReport "Deck Layout Drift Check", nDrift & " placeholder(s) moved, resized or restyled away from their layout, " & nTitle & " of them titles. " & _
        "Deck Reset Placeholders puts them back on a copy. Changes nothing.", Array("Slide", "Placeholder", "Drift", "Layout", "Text"), rows, p
    SKW_Fixable = nDrift
    SKP_DeckLayoutDriftCheck = nDrift & " placeholder(s) drifted, " & nTitle & " title(s)."
End Function

Public Function SKP_DeckResetPlaceholders(ByVal target As String) As String
    Dim p As Object, c As Object, msg As String, rows As New Collection, sld As Object, sh As Object, lph As Object, label As String, d As String
    Dim sz As Single, fn As String, nReset As Long
    msg = OpenDeck(p, "Deck reset placeholders")
    If Len(msg) > 0 Then SKP_DeckResetPlaceholders = msg: Exit Function
    msg = SKP_FixCopy(p, "fixed", target, c)
    If Len(msg) > 0 Then SKP_DeckResetPlaceholders = msg: Exit Function
    On Error GoTo Fail
    SKW_Status "Putting placeholders back on their layout..."
    For Each sld In c.Slides
        For Each sh In sld.Shapes
            If sh.Type = 14 Then
                label = ""
                Set lph = LayoutTwin(sld, sh, label)
                If Not lph Is Nothing Then
                    d = Drift(sh, lph, IsTitlePh(sh))
                    If Len(d) > 0 Then
                        sh.Left = lph.Left: sh.Top = lph.Top: sh.Width = lph.Width: sh.Height = lph.Height
                        If IsTitlePh(sh) And sh.TextFrame.HasText Then
                            fn = FirstRunFont(lph, sz)
                            If sz > 0 Then sh.TextFrame.TextRange.Font.Size = sz
                            If Len(fn) > 0 Then sh.TextFrame.TextRange.Font.Name = fn
                        End If
                        nReset = nReset + 1
                        rows.Add SKW_Row(sld.SlideIndex, label, "Reset: was " & d, sld.CustomLayout.Name)
                    End If
                End If
            End If
        Next
    Next
    SKP_FinishCopy c
    FinishReport "Deck Reset Placeholders", nReset & " placeholder(s) put back where their layout has them (titles in the layout's font and size). " & _
        "Check slides where the old place was on purpose. The open deck is not changed. The fixed copy: " & SKP_LastCopy, Array("Slide", "Placeholder", "Change", "Layout"), rows, p
    SKP_DeckResetPlaceholders = nReset & " placeholder(s) reset." & vbCr & "Fixed copy: " & SKP_LastCopy
    Exit Function
Fail:
    SKP_DeckResetPlaceholders = "Reset placeholders stopped: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Function

' ---------------------------------------------------------------------
' Pictures
' ---------------------------------------------------------------------
Private Function IsPicture(ByVal sh As Object) As Boolean
    On Error Resume Next
    If sh.Type = 13 Or sh.Type = 11 Then
        IsPicture = True
    ElseIf sh.Type = 14 Then
        IsPicture = (sh.PlaceholderFormat.ContainedType = 13)
    End If
End Function

' The picture's own width and height (at 100 percent, as cropped), from a throwaway duplicate. False if it cannot be read.
Private Function OwnSize(ByVal sh As Object, ByRef w As Single, ByRef h As Single) As Boolean
    Dim d As Object
    On Error Resume Next
    Set d = sh.Duplicate(1)
    If d Is Nothing Then Exit Function
    d.LockAspectRatio = 0
    d.ScaleHeight 1, -1
    d.ScaleWidth 1, -1
    w = d.Width: h = d.Height
    d.Delete
    OwnSize = (Err.Number = 0 And w > 0 And h > 0)
End Function

' A throwaway copy of the open deck to measure in (pictures are measured by duplicating them).
Private Function ScratchCopy(ByVal p As Object, ByRef c As Object, ByRef tmp As String) As String
    On Error GoTo Fail
    tmp = Environ$("TEMP") & "\Stationkeeper deck check " & Format$(Now, "hhnnss") & ".pptx"
    p.SaveCopyAs tmp
    Set c = p.Application.Presentations.Open(FileName:=tmp, ReadOnly:=-1, Untitled:=0, WithWindow:=0)
    Exit Function
Fail:
    ScratchCopy = "Could not make a scratch copy to measure in: " & Err.Description
End Function

Private Sub DropScratch(ByVal c As Object, ByVal tmp As String)
    On Error Resume Next
    c.Saved = True
    c.Close
    Kill tmp
End Sub

Private Function PictureRows(ByVal c As Object, ByVal rows As Collection, ByRef nStretch As Long, ByRef nBig As Long, ByRef nCrop As Long) As Boolean
    Dim sld As Object, sh As Object, w As Single, h As Single, r As Single, zoom As Single, cl As Single, ct As Single, cr As Single, cb As Single
    For Each sld In c.Slides
        For Each sh In SKP_Shapes(sld)
            If IsPicture(sh) Then
                If sh.Height > 0 And sh.Width > 0 And OwnSize(sh, w, h) Then
                    r = (sh.Width / sh.Height) / (w / h)
                    If Abs(r - 1) > 0.03 Then
                        nStretch = nStretch + 1
                        rows.Add SKW_Row(sld.SlideIndex, sh.Name, "Stretched", IIf(r > 1, Format$(r - 1, "0%") & " too wide", Format$(1 / r - 1, "0%") & " too tall"), "Deck Picture Aspect Fix")
                    End If
                    zoom = sh.Width / w
                    If zoom > 1.5 Then
                        nBig = nBig + 1
                        rows.Add SKW_Row(sld.SlideIndex, sh.Name, "Enlarged", "Shown at " & Format$(zoom, "0%") & " of its own size: may look soft in print", "Use a larger original")
                    End If
                End If
                On Error Resume Next
                cl = 0: ct = 0: cr = 0: cb = 0
                cl = sh.PictureFormat.CropLeft: ct = sh.PictureFormat.CropTop: cr = sh.PictureFormat.CropRight: cb = sh.PictureFormat.CropBottom
                On Error GoTo 0
                If cl + ct + cr + cb > 1 Then
                    nCrop = nCrop + 1
                    rows.Add SKW_Row(sld.SlideIndex, sh.Name, "Cropped", "The cut-off parts still travel in the file (left " & Pt(cl) & ", top " & Pt(ct) & ", right " & Pt(cr) & ", bottom " & Pt(cb) & ")", _
                        "Picture Format > Compress Pictures > Delete cropped areas")
                End If
            End If
        Next
    Next
End Function

Public Function SKP_DeckPictureCheck() As String
    Dim p As Object, c As Object, tmp As String, msg As String, rows As New Collection, nStretch As Long, nBig As Long, nCrop As Long
    msg = OpenDeck(p, "Deck picture check")
    If Len(msg) > 0 Then SKP_DeckPictureCheck = msg: Exit Function
    msg = ScratchCopy(p, c, tmp)
    If Len(msg) > 0 Then SKW_Status "": SKP_DeckPictureCheck = msg: Exit Function
    PictureRows c, rows, nStretch, nBig, nCrop
    DropScratch c, tmp
    FinishReport "Deck Picture Check", nStretch & " stretched, " & nBig & " enlarged past 150 percent, " & nCrop & " cropped with the cut-off parts still in the file. " & _
        "Measured on a scratch copy; the open deck is not changed.", Array("Slide", "Picture", "Problem", "Detail", "Fix"), rows, p
    SKW_Fixable = nStretch
    SKP_DeckPictureCheck = nStretch & " stretched, " & nBig & " enlarged, " & nCrop & " cropped."
End Function

Public Function SKP_DeckPictureAspectFix(ByVal target As String) As String
    Dim p As Object, c As Object, msg As String, rows As New Collection, sld As Object, sh As Object, w As Single, h As Single
    Dim cx As Single, cy As Single, bw As Single, bh As Single, nw As Single, nh As Single, nFixed As Long
    msg = OpenDeck(p, "Deck picture aspect fix")
    If Len(msg) > 0 Then SKP_DeckPictureAspectFix = msg: Exit Function
    msg = SKP_FixCopy(p, "fixed", target, c)
    If Len(msg) > 0 Then SKP_DeckPictureAspectFix = msg: Exit Function
    On Error GoTo Fail
    SKW_Status "Giving stretched pictures their shape back..."
    For Each sld In c.Slides
        For Each sh In SKP_Shapes(sld)
            If IsPicture(sh) And sh.Height > 0 Then
                If OwnSize(sh, w, h) Then
                    If Abs((sh.Width / sh.Height) / (w / h) - 1) > 0.03 Then
                        bw = sh.Width: bh = sh.Height
                        cx = sh.Left + bw / 2: cy = sh.Top + bh / 2
                        If bw / bh > w / h Then
                            nh = bh: nw = bh * w / h
                        Else
                            nw = bw: nh = bw * h / w
                        End If
                        sh.LockAspectRatio = 0
                        sh.Width = nw: sh.Height = nh
                        sh.Left = cx - nw / 2: sh.Top = cy - nh / 2
                        sh.LockAspectRatio = -1
                        nFixed = nFixed + 1
                        rows.Add SKW_Row(sld.SlideIndex, sh.Name, Format$(bw, "0") & " x " & Format$(bh, "0") & " pt", Format$(nw, "0") & " x " & Format$(nh, "0") & " pt, centered in its old box")
                    End If
                End If
            End If
        Next
    Next
    SKP_FinishCopy c
    FinishReport "Deck Picture Aspect Fix", nFixed & " stretched picture(s) given their own shape back, fitted inside the box they had and centered there. " & _
        "The open deck is not changed. The fixed copy: " & SKP_LastCopy, Array("Slide", "Picture", "Was", "Now"), rows, p
    SKP_DeckPictureAspectFix = nFixed & " picture(s) fixed." & vbCr & "Fixed copy: " & SKP_LastCopy
    Exit Function
Fail:
    SKP_DeckPictureAspectFix = "Picture fix stopped: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Function

' ---------------------------------------------------------------------
' Colors not in the theme
' ---------------------------------------------------------------------
Private Function ThemeNames() As Variant
    ThemeNames = Array("", "Dark 1", "Light 1", "Dark 2", "Light 2", "Accent 1", "Accent 2", "Accent 3", "Accent 4", "Accent 5", "Accent 6", "Hyperlink", "Followed hyperlink")
End Function

Private Function ThemeColors(ByVal d As Object) As Variant
    Dim v(1 To 12) As Long, i As Long
    On Error Resume Next
    For i = 1 To 12
        v(i) = -1
        v(i) = d.SlideMaster.Theme.ThemeColorScheme.Colors(i).RGB
    Next
    ThemeColors = v
End Function

Private Function Nearest(ByVal col As Long, ByVal tc As Variant) As String
    Dim i As Long, best As Long, bestD As Double, dd As Double, nm As Variant
    bestD = -1
    For i = 1 To 12
        If tc(i) >= 0 Then
            dd = ((col And &HFF&) - (tc(i) And &HFF&)) ^ 2 + (((col \ &H100&) And &HFF&) - ((tc(i) \ &H100&) And &HFF&)) ^ 2 + (((col \ &H10000) And &HFF&) - ((tc(i) \ &H10000) And &HFF&)) ^ 2
            If bestD < 0 Or dd < bestD Then bestD = dd: best = i
        End If
    Next
    nm = ThemeNames()
    If best > 0 Then Nearest = nm(best) & " " & Hex6(tc(best))
End Function

Private Sub NoteColor(ByVal found As Object, ByVal cf As Object, ByVal what As String, ByVal sld As Object, ByVal tc As Variant, ByVal extra As String)
    Dim col As Long, i As Long, key As String, e As Variant
    On Error Resume Next
    If cf.ObjectThemeColor <> 0 Then Exit Sub   ' a theme color
    col = cf.RGB
    If Err.Number <> 0 Then Exit Sub
    For i = 1 To 12
        If tc(i) = col Then Exit Sub
    Next
    If InStr(1, extra, Hex6(col), vbTextCompare) > 0 Then Exit Sub
    key = Hex6(col) & "|" & what
    If found.Exists(key) Then
        e = found(key): e(0) = e(0) + 1: found(key) = e
    Else
        found(key) = Array(1, sld.SlideIndex, Nearest(col, tc))
    End If
End Sub

Public Function SKP_DeckColorCheck() As String
    Dim p As Object, msg As String, rows As New Collection, found As Object, sld As Object, sh As Object, tr As Object, rn As Object
    Dim tc As Variant, extra As String, k As Variant, e As Variant, parts As Variant, nCol As Object, skip As Boolean
    msg = OpenDeck(p, "Deck color check")
    If Len(msg) > 0 Then SKP_DeckColorCheck = msg: Exit Function
    extra = SKW_Setting("DeckExtraColors", "")
    Set found = CreateObject("Scripting.Dictionary")
    For Each sld In p.Slides
        tc = ThemeColors(sld.Design)
        For Each sh In SKP_Shapes(sld)
            skip = True
            On Error Resume Next
            skip = (sh.Type = 6 Or sh.Type = 13 Or sh.Type = 11 Or sh.Type = 3 Or sh.Type = 24 Or sh.HasChart = -1 Or sh.HasSmartArt = -1)
            On Error GoTo 0
            If Not skip Then
                On Error Resume Next
                If sh.Fill.Visible = -1 And sh.Fill.Type = 1 Then NoteColor found, sh.Fill.ForeColor, "Fill", sld, tc, extra
                If sh.Line.Visible = -1 Then NoteColor found, sh.Line.ForeColor, "Line", sld, tc, extra
                On Error GoTo 0
                For Each tr In SKP_TextRanges(sh)
                    For Each rn In tr.Runs
                        If Len(Trim$(rn.Text)) > 0 Then NoteColor found, rn.Font.Color, "Text", sld, tc, extra
                    Next
                Next
            End If
        Next
    Next
    Set nCol = CreateObject("Scripting.Dictionary")
    For Each k In found.Keys
        e = found(k)
        parts = Split(k, "|")
        nCol(parts(0)) = True
        rows.Add SKW_Row(parts(0), parts(1), e(0), e(1), e(2))
    Next
    FinishReport "Deck Color Check", nCol.Count & " color(s) not in the theme, in " & found.Count & " use(s) by kind. Theme colors and DeckExtraColors are not listed; pictures, charts and SmartArt are not read. Changes nothing.", _
        Array("Color", "Used for", "Times", "First slide", "Nearest theme color"), rows, p
    SKP_DeckColorCheck = nCol.Count & " color(s) not in the theme."
End Function

' ---------------------------------------------------------------------
' Proofing language
' ---------------------------------------------------------------------
Private Function LangId(ByRef label As String) As Long
    Dim v As String
    v = Trim$(SKW_Setting("DeckLanguage", "en-US"))
    label = v
    Select Case LCase$(v)
        Case "en-us", "english (united states)", "": LangId = 1033: label = "en-US"
        Case "en-gb", "english (united kingdom)": LangId = 2057
        Case "en-ca": LangId = 4105
        Case "en-au": LangId = 3081
        Case "fr-fr": LangId = 1036
        Case "fr-ca": LangId = 3084
        Case "de-de": LangId = 1031
        Case "es-es": LangId = 3082
        Case "es-us": LangId = 21514
        Case "es-mx": LangId = 2058
        Case Else: LangId = CLng(Val(v))
    End Select
End Function

Private Function LangName(ByVal id As Long) As String
    Select Case id
        Case 1033: LangName = "en-US"
        Case 2057: LangName = "en-GB"
        Case 4105: LangName = "en-CA"
        Case 3081: LangName = "en-AU"
        Case 1036: LangName = "fr-FR"
        Case 3084: LangName = "fr-CA"
        Case 1031: LangName = "de-DE"
        Case 3082: LangName = "es-ES"
        Case 21514: LangName = "es-US"
        Case 2058: LangName = "es-MX"
        Case 1024: LangName = "(no proofing)"
        Case Else: LangName = "Language " & id
    End Select
End Function

Private Function LangRows(ByVal p As Object, ByVal want As Long, ByVal rows As Collection) As Long
    Dim found As Object, pl As Variant, rn As Object, id As Long, e As Variant, k As Variant, n As Long
    Set found = CreateObject("Scripting.Dictionary")
    For Each pl In SKP_TextPlaces(p)
        For Each rn In pl(1).Runs
            If Len(Trim$(rn.Text)) > 0 Then
                id = 0
                On Error Resume Next
                id = rn.LanguageID
                On Error GoTo 0
                If id <> want And id <> 0 Then
                    n = n + 1
                    If found.Exists(id) Then
                        e = found(id): e(0) = e(0) + 1: found(id) = e
                    Else
                        found(id) = Array(1, pl(0), SKW_Excerpt(rn.Text, 50))
                    End If
                End If
            End If
        Next
    Next
    For Each k In found.Keys
        e = found(k)
        rows.Add SKW_Row(LangName(CLng(k)), e(0), e(1), e(2))
    Next
    LangRows = n
End Function

Public Function SKP_DeckLanguageCheck() As String
    Dim p As Object, msg As String, rows As New Collection, want As Long, label As String, n As Long
    msg = OpenDeck(p, "Deck proofing language")
    If Len(msg) > 0 Then SKP_DeckLanguageCheck = msg: Exit Function
    want = LangId(label)
    If want <= 0 Then SKW_Status "": SKP_DeckLanguageCheck = "DeckLanguage """ & label & """ is not a language this tool knows. Use en-US, en-GB or a number such as 1033.": Exit Function
    n = LangRows(p, want, rows)
    FinishReport "Deck Language Check", n & " text run(s) marked with a language other than " & label & " (slides, notes, masters, layouts, SmartArt). " & _
        "Spelling is checked against the wrong dictionary there, or not at all. Changes nothing.", Array("Language", "Text runs", "First place", "Text"), rows, p
    SKW_Fixable = n
    SKP_DeckLanguageCheck = n & " run(s) not in " & label & "."
End Function

Public Function SKP_DeckLanguageFix(ByVal target As String) As String
    Dim p As Object, c As Object, msg As String, rows As New Collection, want As Long, label As String, n As Long, pl As Variant
    msg = OpenDeck(p, "Deck language fix")
    If Len(msg) > 0 Then SKP_DeckLanguageFix = msg: Exit Function
    want = LangId(label)
    If want <= 0 Then SKW_Status "": SKP_DeckLanguageFix = "DeckLanguage """ & label & """ is not a language this tool knows. Use en-US, en-GB or a number such as 1033.": Exit Function
    msg = SKP_FixCopy(p, "fixed", target, c)
    If Len(msg) > 0 Then SKP_DeckLanguageFix = msg: Exit Function
    On Error GoTo Fail
    SKW_Status "Marking text as " & label & "..."
    n = LangRows(c, want, rows)
    On Error Resume Next
    For Each pl In SKP_TextPlaces(c)
        pl(1).LanguageID = want
    Next
    c.DefaultLanguageID = want
    On Error GoTo Fail
    SKP_FinishCopy c
    FinishReport "Deck Language Fix", n & " text run(s) set to " & label & ", and the deck's default language. The rows show what they were. " & _
        "The open deck is not changed. The fixed copy: " & SKP_LastCopy, Array("Was", "Text runs", "First place", "Text"), rows, p
    SKP_DeckLanguageFix = n & " run(s) set to " & label & "." & vbCr & "Fixed copy: " & SKP_LastCopy
    Exit Function
Fail:
    SKP_DeckLanguageFix = "Language fix stopped: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Function

' ---------------------------------------------------------------------
' Links and embedded objects
' ---------------------------------------------------------------------
Private Function LinkProblem(ByVal addr As String) As String
    Dim a As String, f As String
    a = LCase$(addr)
    If Len(a) = 0 Then Exit Function
    If a Like "http*://*" Then LinkProblem = "Web address: check it opens for the reader": Exit Function
    If a Like "mailto:*" Then LinkProblem = "": Exit Function
    f = addr
    If a Like "file:*" Then
        f = Replace(Mid$(addr, 6), "/", "\")
        Do While Left$(f, 1) = "\"
            f = Mid$(f, 2)
        Loop
        If Not f Like "[A-Za-z]:\*" Then f = "\\" & f   ' file://server/share
    End If
    If InStr(a, "\users\") > 0 Or InStr(a, "/users/") > 0 Or InStr(a, "\appdata\") > 0 Then
        LinkProblem = "Points into someone's own folders: no one else can open it"
    ElseIf f Like "[A-Za-z]:\*" And Not f Like "\\*" Then
        LinkProblem = "Points at a drive on one computer: no one else can open it"
    ElseIf f Like "\\*" Then
        LinkProblem = "Network path: readers outside the network cannot open it"
    End If
    On Error Resume Next
    If (f Like "[A-Za-z]:\*" Or f Like "\\*") Then
        If Len(Dir$(f)) = 0 Then LinkProblem = "File not found. " & LinkProblem
    End If
End Function

Public Function SKP_DeckLinksCheck() As String
    Dim p As Object, msg As String, rows As New Collection, sld As Object, sh As Object, hl As Object, src As String, kind As String, prob As String
    Dim nLinks As Long, nBad As Long, nEmbed As Long, prog As String
    msg = OpenDeck(p, "Deck links check")
    If Len(msg) > 0 Then SKP_DeckLinksCheck = msg: Exit Function
    For Each sld In p.Slides
        For Each sh In SKP_Shapes(sld)
            kind = "": src = ""
            Select Case sh.Type
                Case 11: kind = "Linked picture"
                Case 10: kind = "Linked object"
                Case 16: kind = "Media"
                Case 7: kind = "Embedded object"
            End Select
            If kind = "Embedded object" Then
                prog = ""
                On Error Resume Next
                prog = sh.OLEFormat.ProgID
                On Error GoTo 0
                nEmbed = nEmbed + 1
                rows.Add SKW_Row(sld.SlideIndex, sh.Name, kind & IIf(Len(prog) > 0, " (" & prog & ")", ""), "(inside the deck)", "The whole file travels with the deck: check it holds nothing that should not go out")
            ElseIf Len(kind) > 0 Then
                On Error Resume Next
                src = sh.LinkFormat.SourceFullName
                On Error GoTo 0
                If Len(src) > 0 Then
                    nLinks = nLinks + 1
                    prob = LinkProblem(src)
                    If Len(prob) > 0 And Not LCase$(src) Like "http*" Then nBad = nBad + 1
                    rows.Add SKW_Row(sld.SlideIndex, sh.Name, kind, src, IIf(Len(prob) > 0, prob, "OK"))
                End If
            End If
        Next
        For Each hl In sld.Hyperlinks
            src = ""
            On Error Resume Next
            src = hl.Address
            On Error GoTo 0
            If Len(src) > 0 Then
                nLinks = nLinks + 1
                prob = LinkProblem(src)
                If Len(prob) > 0 And Not LCase$(src) Like "http*" Then nBad = nBad + 1
                rows.Add SKW_Row(sld.SlideIndex, "", "Hyperlink", src, IIf(Len(prob) > 0, prob, "OK"))
            End If
        Next
    Next
    FinishReport "Deck Links Check", nLinks & " link(s), " & nBad & " that readers cannot follow; " & nEmbed & " embedded object(s). Changes nothing.", _
        Array("Slide", "Shape", "Kind", "Points to", "Problem"), rows, p
    SKP_DeckLinksCheck = nLinks & " link(s), " & nBad & " broken or local, " & nEmbed & " embedded."
End Function

' ---------------------------------------------------------------------
' Table style
' ---------------------------------------------------------------------
' The style id to use: a {GUID} as written, a style name another table in the deck already uses, or a known built-in name.
Private Function StyleIdFor(ByVal p As Object, ByVal v As String) As String
    Dim sld As Object, sh As Object
    v = Trim$(v)
    If Left$(v, 1) = "{" Then StyleIdFor = v: Exit Function
    If Len(v) = 0 Then Exit Function
    On Error Resume Next
    For Each sld In p.Slides
        For Each sh In SKP_Shapes(sld)
            If sh.HasTable Then
                If StrComp(sh.Table.Style.Name, v, vbTextCompare) = 0 Then StyleIdFor = sh.Table.Style.Id: Exit Function
            End If
        Next
    Next
    Select Case LCase$(v)
        Case "medium style 2 - accent 1": StyleIdFor = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
        Case "no style, table grid": StyleIdFor = "{5940675A-B579-460E-94D1-54222C63F5DA}"
        Case "no style, no grid": StyleIdFor = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
    End Select
End Function

Public Function SKP_DeckTableStyleCheck() As String
    Dim p As Object, msg As String, rows As New Collection, sld As Object, sh As Object, want As String, id As String, nm As String
    Dim nTables As Long, nOff As Long, nNoHead As Long, used As Object, k As Variant
    msg = OpenDeck(p, "Deck table style")
    If Len(msg) > 0 Then SKP_DeckTableStyleCheck = msg: Exit Function
    want = SKW_Setting("DeckTableStyle", "")
    id = StyleIdFor(p, want)
    Set used = CreateObject("Scripting.Dictionary")
    For Each sld In p.Slides
        For Each sh In SKP_Shapes(sld)
            If sh.HasTable Then
                nTables = nTables + 1
                nm = "": On Error Resume Next: nm = sh.Table.Style.Name: On Error GoTo 0
                used(nm) = used(nm) + 1
                If Len(id) > 0 Then
                    If StrComp(sh.Table.Style.Id, id, vbTextCompare) <> 0 Then
                        nOff = nOff + 1
                        rows.Add SKW_Row(sld.SlideIndex, sh.Name, "Style " & nm, "Not the team style (" & want & ")")
                    End If
                End If
                If Not sh.Table.FirstRow Then
                    nNoHead = nNoHead + 1
                    rows.Add SKW_Row(sld.SlideIndex, sh.Name, "No header row", "Table Design > Header Row: screen readers and the next page need it")
                End If
            End If
        Next
    Next
    If Len(want) = 0 Then
        For Each k In used.Keys
            rows.Add SKW_Row("", "", "Style " & k, used(k) & " table(s): set DeckTableStyle to the team's")
        Next
    ElseIf Len(id) = 0 Then
        rows.Add SKW_Row("", "", "DeckTableStyle """ & want & """", "Not found: use a table style id such as {5C22544A-7EE6-4342-B048-85BDC9FD1C3A}, or a style a table in the deck already uses")
    End If
    FinishReport "Deck Table Style Check", nTables & " table(s): " & nOff & " not in the team style, " & nNoHead & " without a header row. Changes nothing.", _
        Array("Slide", "Table", "Found", "Detail"), rows, p
    SKW_Fixable = nOff
    SKP_DeckTableStyleCheck = nTables & " table(s), " & nOff & " off style, " & nNoHead & " without a header row."
End Function

Public Function SKP_DeckTableStyleFix(ByVal target As String) As String
    Dim p As Object, c As Object, msg As String, rows As New Collection, sld As Object, sh As Object, want As String, id As String, nm As String, nFixed As Long
    msg = OpenDeck(p, "Deck table style fix")
    If Len(msg) > 0 Then SKP_DeckTableStyleFix = msg: Exit Function
    want = SKW_Setting("DeckTableStyle", "")
    id = StyleIdFor(p, want)
    If Len(id) = 0 Then
        SKW_Status ""
        SKP_DeckTableStyleFix = IIf(Len(want) = 0, "Set DeckTableStyle on the Settings sheet first (a table style id, or the name of a style a table in the deck uses).", _
            "DeckTableStyle """ & want & """ was not found. Use a table style id such as {5C22544A-7EE6-4342-B048-85BDC9FD1C3A}.")
        Exit Function
    End If
    msg = SKP_FixCopy(p, "fixed", target, c)
    If Len(msg) > 0 Then SKP_DeckTableStyleFix = msg: Exit Function
    On Error GoTo Fail
    SKW_Status "Applying the team table style..."
    For Each sld In c.Slides
        For Each sh In SKP_Shapes(sld)
            If sh.HasTable Then
                If StrComp(sh.Table.Style.Id, id, vbTextCompare) <> 0 Then
                    nm = sh.Table.Style.Name
                    sh.Table.ApplyStyle id, False
                    nFixed = nFixed + 1
                    rows.Add SKW_Row(sld.SlideIndex, sh.Name, nm, sh.Table.Style.Name)
                End If
            End If
        Next
    Next
    SKP_FinishCopy c
    FinishReport "Deck Table Style Fix", nFixed & " table(s) set to the team style. Cell colors and borders set by hand are replaced by the style's. " & _
        "The open deck is not changed. The fixed copy: " & SKP_LastCopy, Array("Slide", "Table", "Was", "Now"), rows, p
    SKP_DeckTableStyleFix = nFixed & " table(s) restyled." & vbCr & "Fixed copy: " & SKP_LastCopy
    Exit Function
Fail:
    SKP_DeckTableStyleFix = "Table style fix stopped: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Function
