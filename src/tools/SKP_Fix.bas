Attribute VB_Name = "SKP_Fix"
' =====================================================================
' Stationkeeper Tools for PowerPoint, batch 2 (1.4.0): template hygiene,
' fonts, overlaps, slide numbers, banned phrases and the media list.
' Run from Word on the deck open in PowerPoint. Checks change nothing.
' Fixes never touch the open deck: each saves "<deck> - fixed <date>.pptx"
' beside it, works on that copy, reports what it did and opens the copy.
'   DeckMediaList, DeckTemplateCheck, DeckConformToTemplate,
'   DeckRemoveUnusedLayouts, DeckFontFix, DeckOverlapCheck,
'   DeckSlideNumberCheck, DeckSlideNumberFix, DeckBannedPhrasesFix
' (Deck Font Audit in SKP_Deck now reads masters, layouts, notes and bullets too.)
' =====================================================================
Option Explicit

Public SKP_TestTemplate As String   ' tests: the reference template, instead of the DeckTemplate setting
Public SKP_LastCopy As String       ' the copy the last fix (or marked overlap copy) saved

Private Const PH_TITLE As Long = 1, PH_BODY As Long = 2, PH_CTITLE As Long = 3, PH_SUBTITLE As Long = 4
Private Const PH_NUMBER As Long = 13, PH_FOOTER As Long = 15, PH_DATE As Long = 16

' Lists every picture, video, audio, chart, SmartArt, icon and embedded or linked object in the open deck by slide, and lets you rename them (the names show in PowerPoint's Selection Pane) and jump to them.
Public Sub DeckMediaList()
    If SKW_IsQuiet() Then Exit Sub   ' tests use the SKP_ functions and the form's test calls
    Dim p As Object
    Set p = SKP_Deck.SKP_Deck()
    If p Is Nothing Then SKW_Say SKP_Why(): Exit Sub
    frmSKPMedia.Show vbModeless
End Sub

' Lists every master and layout against the team template: copies pasted in from other decks, foreign masters, unused ones, template layouts gone missing. When it finds any, it offers to fix a copy: move slides onto the template's layouts and remove what is unused, or only remove unused layouts and masters. The open deck is not changed.
Public Sub DeckTemplateCheck()
    If SKW_IsQuiet() Then Exit Sub
    Dim s As String
    SKW_Fixable = 0
    s = SKP_DeckTemplateCheck()
    If SKW_Fixable <= 0 Or SKW_LastReport Is Nothing Then SKW_Show s: Exit Sub
    On Error Resume Next
    SKW_LastReport.Activate
    On Error GoTo 0
    Select Case MsgBox(Split(s & vbCr, vbCr)(0) & vbCr & vbCr & "Fix a copy now? The report stays open." & vbCr & vbCr & _
            "Yes: move slides onto the template's layouts, then remove layouts and masters no slide uses." & vbCr & _
            "No: only remove layouts and masters no slide uses." & vbCr & "Cancel: change nothing.", vbYesNoCancel + vbQuestion, SKW_APP)
        Case vbYes: SKW_Show SKP_DeckConformToTemplate("")
        Case vbNo: SKW_Show SKP_DeckRemoveUnusedLayouts("")
    End Select
End Sub

' Lists slide content that covers template elements: the slide number, footer, date, a classification or CUI banner, logos and other layout graphics, and the title area. Saves a copy with each overlap outlined in red. The open deck is not changed.
Public Sub DeckOverlapCheck()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Show SKP_DeckOverlapCheck("")
End Sub

' Lists slides with no number, layouts with no number placeholder, numbers typed as text and numbers moved or restyled. When it finds any, it offers to fix a copy. The open deck is not changed.
Public Sub DeckSlideNumberCheck()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    Dim s As String
    SKW_Fixable = 0
    s = SKP_DeckSlideNumberCheck()
    If SKW_OfferFix(s, "The fix puts the number placeholder back, turns numbers on (off on title slides), removes typed numbers and puts moved numbers back where the layout has them.") Then SKW_Show SKP_DeckSlideNumberFix("")
End Sub

' ---------------------------------------------------------------------
' Shared: the copy a fix works on, the reference template, layouts
' ---------------------------------------------------------------------
' Saves a copy beside the deck (or at target) and opens it without a window. Returns "" or why not.
Private Function FixCopy(ByVal p As Object, ByVal what As String, ByVal target As String, ByRef c As Object) As String
    Set c = Nothing
    SKP_LastCopy = ""
    If Len(p.Path) = 0 Then FixCopy = "Save the deck first, so the copy has a folder to go in.": Exit Function
    If LCase$(Left$(p.FullName, 4)) = "http" Then FixCopy = "The deck is opened from the web. Open it from the synced folder (or save a local copy) first.": Exit Function
    If Len(target) = 0 Then target = p.Path & "\" & SKW_Stem(p.Name) & " - " & what & " " & Format$(Now, "yyyy-mm-dd hhnn") & ".pptx"
    On Error GoTo Fail
    SKW_Status "Saving a copy to work on..."
    p.SaveCopyAs target
    Set c = p.Application.Presentations.Open(FileName:=target, ReadOnly:=0, Untitled:=0, WithWindow:=0)
    SKP_LastCopy = target
    Exit Function
Fail:
    FixCopy = "The copy could not be made: " & Err.Description
End Function

' Saves and closes the copy, then opens it in a window for review (not in tests).
Private Sub FinishCopy(ByVal c As Object)
    Dim app As Object
    On Error Resume Next
    Set app = c.Application
    c.Save
    c.Close
    If Not SKW_IsQuiet() Then app.Presentations.Open FileName:=SKP_LastCopy, ReadOnly:=0, Untitled:=0, WithWindow:=-1
    SKW_Status ""
End Sub

Private Function BaseName(ByVal nm As String) As String
    Dim i As Long
    i = InStr(nm, "_")
    If i > 1 Then
        If IsNumeric(Left$(nm, i - 1)) Then BaseName = Mid$(nm, i + 1): Exit Function
    End If
    BaseName = nm
End Function

' "name|count|sorted placeholder types" of a layout.
Private Function LayoutSig(ByVal lay As Object) As String
    Dim t() As Long, n As Long, sh As Object, i As Long, j As Long, k As Long, s As String
    On Error Resume Next
    n = lay.Shapes.Placeholders.Count
    If n > 0 Then
        ReDim t(1 To n)
        i = 0
        For Each sh In lay.Shapes.Placeholders
            i = i + 1: t(i) = sh.PlaceholderFormat.Type
        Next
        For i = 2 To n
            k = t(i): j = i - 1
            Do While j >= 1
                If t(j) <= k Then Exit Do
                t(j + 1) = t(j): j = j - 1
            Loop
            t(j + 1) = k
        Next
        For i = 1 To n
            s = s & IIf(i > 1, ",", "") & t(i)
        Next
    End If
    LayoutSig = n & "|" & s
End Function

' PowerPoint hands out a new wrapper for the same layout each time it is asked, so layouts and masters
' are compared by their place (design index, layout index), never with Is.
Private Function LayKey(ByVal lay As Object) As String
    LayKey = lay.Design.Index & "|" & lay.Index
End Function

Private Function SameDesign(ByVal a As Object, ByVal b As Object) As Boolean
    If a Is Nothing Or b Is Nothing Then Exit Function
    SameDesign = (a.Index = b.Index)
End Function

Private Function UseCount(ByVal p As Object, ByVal lay As Object) As Long
    Dim sld As Object, k As String
    k = LayKey(lay)
    For Each sld In p.Slides
        If LayKey(sld.CustomLayout) = k Then UseCount = UseCount + 1
    Next
End Function

Private Function DesignUse(ByVal p As Object, ByVal d As Object) As Long
    Dim sld As Object
    For Each sld In p.Slides
        If sld.Design.Index = d.Index Then DesignUse = DesignUse + 1
    Next
End Function

' Moves a slide onto a layout (a property set PowerPoint takes either way, depending on the build).
Private Sub SetLayout(ByVal sld As Object, ByVal lay As Object)
    On Error Resume Next
    sld.CustomLayout = lay
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Set sld.CustomLayout = lay
    End If
End Sub

' The template's layouts: lower-case name -> signature, plus "~master" -> its first master's name.
' From the DeckTemplate setting (a .potx), else from the deck's own main master (the one most slides use).
Private Function TemplateLayouts(ByVal p As Object, ByRef source As String) As Object
    Dim d As Object, path As String, ref As Object, lay As Object, best As Object, n As Long, bestN As Long
    Set d = CreateObject("Scripting.Dictionary")
    path = SKP_TestTemplate
    If Len(path) = 0 Then path = SKW_Setting("DeckTemplate", "")
    If Len(path) > 0 Then
        On Error Resume Next
        Set ref = p.Application.Presentations.Open(FileName:=path, ReadOnly:=-1, Untitled:=0, WithWindow:=0)
        On Error GoTo 0
        If Not ref Is Nothing Then
            For Each lay In ref.Designs(1).SlideMaster.CustomLayouts
                d(LCase$(lay.Name)) = LayoutSig(lay)
            Next
            d("~master") = ref.Designs(1).Name
            ref.Close
            source = "the team template " & Mid$(path, InStrRev(path, "\") + 1)
            Set TemplateLayouts = d
            Exit Function
        End If
        source = "the deck's main master (the DeckTemplate file " & path & " could not be opened)"
    Else
        source = "the deck's main master (no DeckTemplate is set, so the master most slides use stands in)"
    End If
    For Each best In p.Designs
        n = DesignUse(p, best)
        If n > bestN Or d.Count = 0 Then
            bestN = n
            d.RemoveAll
            For Each lay In best.SlideMaster.CustomLayouts
                If BaseName(lay.Name) = lay.Name Then d(LCase$(lay.Name)) = LayoutSig(lay)
            Next
            d("~master") = best.Name
        End If
    Next
    Set TemplateLayouts = d
End Function

' The design in the deck that holds the template's layouts (most name matches; the template master's name breaks ties).
Private Function TemplateDesign(ByVal p As Object, ByVal tl As Object) As Object
    Dim d As Object, lay As Object, n As Long, bestN As Long
    bestN = -1
    For Each d In p.Designs
        n = 0
        For Each lay In d.SlideMaster.CustomLayouts
            If tl.Exists(LCase$(lay.Name)) Then n = n + 1
        Next
        If LCase$(d.Name) = LCase$(tl("~master")) Then n = n + 1
        If n > bestN Then bestN = n: Set TemplateDesign = d
    Next
End Function

Private Function IsTemplateLayout(ByVal lay As Object, ByVal td As Object, ByVal tl As Object) As Boolean
    If Not SameDesign(lay.Design, td) Then Exit Function
    IsTemplateLayout = tl.Exists(LCase$(lay.Name))
End Function

' The template layout a slide's layout should move to, or Nothing.
Private Function MatchLayout(ByVal lay As Object, ByVal td As Object, ByVal tl As Object, ByRef how As String) As Object
    Dim cand As Object, sig As String, bn As String
    bn = LCase$(BaseName(lay.Name))
    For Each cand In td.SlideMaster.CustomLayouts
        If tl.Exists(LCase$(cand.Name)) And LCase$(cand.Name) = LCase$(lay.Name) Then Set MatchLayout = cand: how = "same name": Exit Function
    Next
    For Each cand In td.SlideMaster.CustomLayouts
        If tl.Exists(LCase$(cand.Name)) And LCase$(cand.Name) = bn Then Set MatchLayout = cand: how = "same name without the number": Exit Function
    Next
    sig = Mid$(LayoutSig(lay), 1)
    For Each cand In td.SlideMaster.CustomLayouts
        If tl.Exists(LCase$(cand.Name)) And LayoutSig(cand) = sig Then Set MatchLayout = cand: how = "same placeholders": Exit Function
    Next
End Function

' ---------------------------------------------------------------------
Public Function SKP_DeckTemplateCheck() As String
    Dim p As Object, rows As New Collection, tl As Object, src As String, td As Object, d As Object, lay As Object
    Dim n As Long, st As String, nForeignM As Long, nOrphan As Long, nForeign As Long, nUnused As Long, nMissing As Long
    Dim present As Object, k As Variant
    Set SKW_LastReport = Nothing
    Set p = SKP_Deck.SKP_Deck()
    If p Is Nothing Then SKP_DeckTemplateCheck = SKP_Why(): Exit Function
    SKW_Status "Deck template check: " & p.Name & "..."
    Set tl = TemplateLayouts(p, src)
    Set td = TemplateDesign(p, tl)
    Set present = CreateObject("Scripting.Dictionary")
    For Each d In p.Designs
        If Not SameDesign(d, td) Then nForeignM = nForeignM + 1
        For Each lay In d.SlideMaster.CustomLayouts
            n = UseCount(p, lay)
            If SameDesign(d, td) And tl.Exists(LCase$(lay.Name)) Then
                st = "Template layout": present(LCase$(lay.Name)) = True
            ElseIf tl.Exists(LCase$(BaseName(lay.Name))) And BaseName(lay.Name) <> lay.Name Then
                st = "Copy of " & BaseName(lay.Name) & " (pasted in)": nOrphan = nOrphan + 1
            Else
                st = IIf(SameDesign(d, td), "Not in the template", "Foreign (from another deck)"): nForeign = nForeign + 1
            End If
            If n = 0 And st <> "Template layout" Then nUnused = nUnused + 1
            rows.Add SKW_Row(d.Name & IIf(SameDesign(d, td), " (template)", ""), lay.Name, st, n, IIf(n = 0 And st <> "Template layout", "Unused: Deck Remove Unused Layouts takes it out", _
                IIf(st <> "Template layout" And n > 0, "Deck Conform to Template moves these slides", "")))
        Next
    Next
    For Each k In tl.Keys
        If k <> "~master" And Not present.Exists(k) Then
            nMissing = nMissing + 1
            rows.Add SKW_Row(td.Name, k, "Missing from the deck", 0, "Apply the template again (Design > Browse for Themes) on a copy to bring it back")
        End If
    Next
    SKW_Status ""
    SKW_ReportFor "Deck Template Check", p.Designs.Count & " master(s), " & nForeignM & " not the template's; " & nOrphan & " pasted-in copy layout(s), " & _
        nForeign & " other layout(s) not in the template, " & nUnused & " of them unused; " & nMissing & " template layout(s) missing. Compared with " & src & _
        ". Matching is by layout name and placeholders, the way PowerPoint matches when slides are pasted. Changes nothing.", _
        Array("Master", "Layout", "Status", "Slides", "What to do"), rows, p.Name
    SKW_Fixable = nForeignM + nOrphan + nForeign + nUnused
    SKP_DeckTemplateCheck = p.Designs.Count & " master(s), " & nForeignM & " foreign; " & nOrphan & " copy layout(s), " & nForeign & " not in the template, " & _
        nUnused & " unused; " & nMissing & " missing."
End Function

' ---------------------------------------------------------------------
Public Function SKP_DeckConformToTemplate(ByVal target As String) As String
    Dim p As Object, c As Object, msg As String, rows As New Collection, tl As Object, src As String, td As Object
    Dim sld As Object, lay As Object, nl As Object, how As String, nMoved As Long, nLeft As Long, purge As String, ttl As String
    Set SKW_LastReport = Nothing
    Set p = SKP_Deck.SKP_Deck()
    If p Is Nothing Then SKP_DeckConformToTemplate = SKP_Why(): Exit Function
    msg = FixCopy(p, "fixed", target, c)
    If Len(msg) > 0 Then SKP_DeckConformToTemplate = msg: Exit Function
    On Error GoTo Fail
    SKW_Status "Conforming slides to the template..."
    Set tl = TemplateLayouts(c, src)
    Set td = TemplateDesign(c, tl)
    For Each sld In c.Slides
        Set lay = sld.CustomLayout
        If Not IsTemplateLayout(lay, td, tl) Then
            Set nl = MatchLayout(lay, td, tl, how)
            ttl = SlideTitleOf(sld)
            If nl Is Nothing Then
                nLeft = nLeft + 1
                rows.Add SKW_Row(sld.SlideIndex, lay.Design.Name & " / " & lay.Name, "(left as it is)", "No template layout matches: pick one by hand (Home > Layout)", ttl)
            Else
                SetLayout sld, nl
                nMoved = nMoved + 1
                rows.Add SKW_Row(sld.SlideIndex, lay.Design.Name & " / " & lay.Name, nl.Name, "Moved (" & how & ")" & _
                    IIf(Len(ttl) > 0 And SlideTitleOf(sld) <> ttl, ". Check the title: it moved", ""), ttl)
            End If
        End If
    Next
    purge = PurgeUnused(c, td, tl, rows)
    FinishCopy c
    SKW_ReportFor "Deck Conform to Template", nMoved & " slide(s) moved onto template layouts, " & nLeft & " left for you to place. " & purge & _
        " Compared with " & src & ". Moving a slide keeps its text and places it in the layout's placeholders; positions follow the layout, local formatting stays. " & _
        "The open deck is not changed. The fixed copy: " & SKP_LastCopy, Array("Slide", "From", "To", "What happened", "Title"), rows, p.Name
    SKP_DeckConformToTemplate = nMoved & " slide(s) moved, " & nLeft & " left. " & purge & vbCr & "Fixed copy: " & SKP_LastCopy
    Exit Function
Fail:
    SKP_DeckConformToTemplate = "Conform to template stopped: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Function

' Deletes unused layouts that are not template layouts, then masters no slide uses (not the template's). Returns a sentence.
Private Function PurgeUnused(ByVal c As Object, ByVal td As Object, ByVal tl As Object, ByVal rows As Collection) As String
    Dim i As Long, j As Long, d As Object, lay As Object, nL As Long, nMasters As Long, nm As String
    For i = c.Designs.Count To 1 Step -1
        Set d = c.Designs(i)
        For j = d.SlideMaster.CustomLayouts.Count To 1 Step -1
            Set lay = d.SlideMaster.CustomLayouts(j)
            If Not IsTemplateLayout(lay, td, tl) And UseCount(c, lay) = 0 Then
                nm = d.Name & " / " & lay.Name
                On Error Resume Next
                lay.Delete
                If Err.Number = 0 Then nL = nL + 1: rows.Add SKW_Row("", nm, "(deleted)", "Unused layout, not in the template", "") _
                    Else rows.Add SKW_Row("", nm, "(kept)", "PowerPoint would not delete it: " & Err.Description, "")
                Err.Clear
                On Error GoTo 0
            End If
        Next
    Next
    For i = c.Designs.Count To 1 Step -1
        Set d = c.Designs(i)
        If Not SameDesign(d, td) And DesignUse(c, d) = 0 Then
            nm = d.Name
            On Error Resume Next
            d.Preserved = 0
            d.Delete
            If Err.Number = 0 Then nMasters = nMasters + 1: rows.Add SKW_Row("", nm, "(deleted)", "Master no slide uses", "") _
                Else rows.Add SKW_Row("", nm, "(kept)", "PowerPoint would not delete it: " & Err.Description, "")
            Err.Clear
            On Error GoTo 0
        End If
    Next
    PurgeUnused = nL & " unused layout(s) and " & nMasters & " unused master(s) deleted."
End Function

Private Function SlideTitleOf(ByVal sld As Object) As String
    On Error Resume Next
    If sld.Shapes.HasTitle Then SlideTitleOf = Trim$(Replace(Replace(sld.Shapes.Title.TextFrame.TextRange.Text, vbCr, " "), Chr$(11), " "))
End Function

' ---------------------------------------------------------------------
Public Function SKP_DeckRemoveUnusedLayouts(ByVal target As String) As String
    Dim p As Object, c As Object, msg As String, rows As New Collection, tl As Object, src As String, td As Object, purge As String
    Set SKW_LastReport = Nothing
    Set p = SKP_Deck.SKP_Deck()
    If p Is Nothing Then SKP_DeckRemoveUnusedLayouts = SKP_Why(): Exit Function
    msg = FixCopy(p, "fixed", target, c)
    If Len(msg) > 0 Then SKP_DeckRemoveUnusedLayouts = msg: Exit Function
    On Error GoTo Fail
    Set tl = TemplateLayouts(c, src)
    Set td = TemplateDesign(c, tl)
    purge = PurgeUnused(c, td, tl, rows)
    FinishCopy c
    SKW_ReportFor "Deck Remove Unused Layouts", purge & " Template layouts stay even when unused, so slides pasted in later still find them. Compared with " & src & _
        ". The open deck is not changed. The fixed copy: " & SKP_LastCopy, Array("Slide", "Layout", "Result", "Why", "Title"), rows, p.Name
    SKP_DeckRemoveUnusedLayouts = purge & vbCr & "Fixed copy: " & SKP_LastCopy
    Exit Function
Fail:
    SKP_DeckRemoveUnusedLayouts = "Remove unused layouts stopped: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Function

' ---------------------------------------------------------------------
' Fonts: every text range in slides, notes, layouts and masters.
' ---------------------------------------------------------------------
Public Function SKP_AllowedFonts() As Collection
    Dim allowed As New Collection, r As Object
    For Each r In SKW_Rows("Fonts")
        If r.Exists("font") Then
            If Len(r("font")) > 0 Then
                If Not r.Exists("allowed") Then
                    allowed.Add r("font")
                ElseIf SKW_IsYes(r("allowed")) Or Len(r("allowed")) = 0 Then
                    allowed.Add r("font")
                End If
            End If
        End If
    Next
    Set SKP_AllowedFonts = allowed
End Function

' Every text-bearing place: each item is Array(where, text range, is a title).
Public Function SKP_TextPlaces(ByVal p As Object) As Collection
    Dim c As New Collection, sld As Object, d As Object, lay As Object
    For Each sld In p.Slides
        AddPlaces c, "Slide " & sld.SlideIndex, sld.Shapes
        On Error Resume Next
        AddPlaces c, "Notes " & sld.SlideIndex, sld.NotesPage.Shapes
        On Error GoTo 0
    Next
    For Each d In p.Designs
        AddPlaces c, "Master " & d.Name, d.SlideMaster.Shapes
        For Each lay In d.SlideMaster.CustomLayouts
            AddPlaces c, "Layout " & lay.Name, lay.Shapes
        Next
    Next
    Set SKP_TextPlaces = c
End Function

Private Sub AddPlaces(ByVal c As Collection, ByVal where As String, ByVal shapes As Object)
    Dim sh As Object, g As Object
    For Each sh In shapes
        AddShapePlaces c, where, sh
        If sh.Type = 6 Then
            For Each g In sh.GroupItems
                AddShapePlaces c, where, g
            Next
        End If
    Next
End Sub

Private Sub AddShapePlaces(ByVal c As Collection, ByVal where As String, ByVal sh As Object)
    Dim tr As Variant, isTitle As Boolean, i As Long, n As Long
    On Error Resume Next
    If sh.Type = 14 Then isTitle = (sh.PlaceholderFormat.Type = PH_TITLE Or sh.PlaceholderFormat.Type = PH_CTITLE)
    For Each tr In SKP_TextRanges(sh)
        c.Add Array(where, tr, isTitle, sh.Name)
    Next
    If sh.HasSmartArt Then
        n = 0: n = sh.SmartArt.AllNodes.Count
        For i = 1 To n
            c.Add Array(where, sh.SmartArt.AllNodes(i).TextFrame2.TextRange, False, sh.Name)
        Next
    End If
End Sub

' The font names a run carries: Latin, East Asian, complex script, and the bullet's.
Private Function RunFonts(ByVal rn As Object) As Collection
    Dim c As New Collection, v As String
    On Error Resume Next
    v = rn.Font.Name: If Len(v) > 0 Then c.Add Array("Latin", v)
    v = "": v = rn.Font.NameFarEast: If Len(v) > 0 Then c.Add Array("East Asian", v)
    v = "": v = rn.Font.NameComplexScript: If Len(v) > 0 Then c.Add Array("Complex script", v)
    If rn.ParagraphFormat.Bullet.Visible Then
        If rn.ParagraphFormat.Bullet.Type = 1 Then v = "": v = rn.ParagraphFormat.Bullet.Font.Name: If Len(v) > 0 Then c.Add Array("Bullet", v)
    End If
    Set RunFonts = c
End Function

' The deep font audit behind Deck Font Audit. Rows: Rule, Font, Text runs, First place, Text, Slot.
Public Function SKP_FontRows(ByVal p As Object, ByVal allowed As Collection, ByRef nBad As Long) As Collection
    Dim rows As New Collection, fonts As Object, pl As Variant, rn As Object, f As Variant, key As String, e As Variant, k As Variant
    Dim seen As Object, i As Long, nm As String
    Set fonts = CreateObject("Scripting.Dictionary")
    For Each pl In SKP_TextPlaces(p)
        For Each rn In pl(1).Runs
            If Len(Trim$(rn.Text)) > 0 Then
                For Each f In RunFonts(rn)
                    If f(0) = "Latin" Or f(0) = "Bullet" Or Not SKW_InList(allowed, CStr(f(1))) Or allowed.Count = 0 Then
                        key = f(1) & "|" & f(0)
                        If fonts.Exists(key) Then
                            e = fonts(key): e(0) = e(0) + 1: fonts(key) = e
                        Else
                            fonts(key) = Array(1, pl(0), SKW_Excerpt(rn.Text, 50))
                        End If
                    End If
                Next
            End If
        Next
    Next
    Set seen = CreateObject("Scripting.Dictionary")
    For Each k In fonts.Keys
        e = fonts(k)
        nm = Split(k, "|")(0)
        seen(LCase$(nm)) = True
        If allowed.Count > 0 And Not SKW_InList(allowed, nm) Then
            nBad = nBad + 1
            rows.Add SKW_Row("Not allowed", nm, e(0), e(1), e(2), Split(k, "|")(1))
        Else
            rows.Add SKW_Row(IIf(allowed.Count > 0, "Allowed", "Used"), nm, e(0), e(1), e(2), Split(k, "|")(1))
        End If
    Next
    ' Fonts PowerPoint lists that no text uses (charts, old layouts, embedded fonts).
    On Error Resume Next
    For i = 1 To p.Fonts.Count
        nm = p.Fonts(i).Name
        If Len(nm) > 0 And Not seen.Exists(LCase$(nm)) Then
            rows.Add SKW_Row("Listed only", nm, 0, "", IIf(p.Fonts(i).Embedded, "Embedded in the file. ", "") & "PowerPoint lists it, but no text found uses it (charts, old layouts or spacing)", "")
        End If
    Next
    Set SKP_FontRows = rows
End Function

' ---------------------------------------------------------------------
Public Function SKP_DeckFontFix(ByVal target As String) As String
    Dim p As Object, c As Object, msg As String, rows As New Collection, allowed As Collection, pl As Variant, rn As Object
    Dim nFixed As Long, nSlots As Long, old As String, tgt As String, tally As Object, k As Variant, e As Variant
    Set SKW_LastReport = Nothing
    Set p = SKP_Deck.SKP_Deck()
    If p Is Nothing Then SKP_DeckFontFix = SKP_Why(): Exit Function
    Set allowed = SKP_AllowedFonts()
    If allowed.Count = 0 Then SKP_DeckFontFix = "No Fonts list in the team standards, so there is nothing to fix against. Run Deck Font Audit to see the fonts.": Exit Function
    msg = FixCopy(p, "fixed", target, c)
    If Len(msg) > 0 Then SKP_DeckFontFix = msg: Exit Function
    On Error GoTo Fail
    SKW_Status "Setting fonts back to the theme..."
    Set tally = CreateObject("Scripting.Dictionary")
    For Each pl In SKP_TextPlaces(c)
        For Each rn In pl(1).Runs
            If Len(Trim$(rn.Text)) > 0 Then
                old = rn.Font.Name
                If Len(old) > 0 And Not SKW_InList(allowed, old) Then
                    tgt = IIf(pl(2), "+mj-lt", "+mn-lt")
                    rn.Font.Name = tgt
                    nFixed = nFixed + 1
                    k = pl(0) & "|" & old & "|" & IIf(pl(2), "heading", "body")
                    If tally.Exists(k) Then e = tally(k): e(0) = e(0) + 1: tally(k) = e Else tally(k) = Array(1, rn.Font.Name, SKW_Excerpt(rn.Text, 40))
                End If
                On Error Resume Next
                old = "": old = rn.Font.NameFarEast
                If Len(old) > 0 And Not SKW_InList(allowed, old) And Left$(old, 1) <> "+" Then rn.Font.NameFarEast = "+mn-ea": nSlots = nSlots + 1
                old = "": old = rn.Font.NameComplexScript
                If Len(old) > 0 And Not SKW_InList(allowed, old) And Left$(old, 1) <> "+" Then rn.Font.NameComplexScript = "+mn-cs": nSlots = nSlots + 1
                On Error GoTo Fail
            End If
        Next
    Next
    For Each k In tally.Keys
        e = tally(k)
        rows.Add SKW_Row(Split(k, "|")(0), Split(k, "|")(1), "Theme " & Split(k, "|")(2) & " font (" & e(1) & ")", e(0), e(2))
    Next
    FinishCopy c
    SKW_ReportFor "Deck Font Fix", nFixed & " text run(s) set back to the theme font" & IIf(nSlots > 0, ", plus " & nSlots & " East Asian or complex-script slot(s)", "") & _
        ". Titles take the theme's heading font, the rest its body font, so the text follows the template from now on. If the theme's own font is not on the list, " & _
        "apply the team template. The open deck is not changed. The fixed copy: " & SKP_LastCopy, Array("Where", "Was", "Now", "Runs", "Text"), rows, p.Name
    SKP_DeckFontFix = nFixed & " text run(s) set back to the theme font." & vbCr & "Fixed copy: " & SKP_LastCopy
    Exit Function
Fail:
    SKP_DeckFontFix = "Font fix stopped: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Function

' ---------------------------------------------------------------------
' Overlaps with template elements
' ---------------------------------------------------------------------
' The box a shape covers on the slide, rotation included: Array(left, top, right, bottom).
Private Function Box(ByVal sh As Object) As Variant
    Dim l As Double, t As Double, w As Double, h As Double, a As Double, cx As Double, cy As Double, ew As Double, eh As Double
    l = sh.Left: t = sh.Top: w = sh.Width: h = sh.Height
    On Error Resume Next
    a = sh.Rotation
    On Error GoTo 0
    If a <> 0 Then
        a = a * 3.14159265358979 / 180
        cx = l + w / 2: cy = t + h / 2
        ew = Abs(w * Cos(a)) + Abs(h * Sin(a)): eh = Abs(w * Sin(a)) + Abs(h * Cos(a))
        Box = Array(cx - ew / 2, cy - eh / 2, cx + ew / 2, cy + eh / 2)
    Else
        Box = Array(l, t, l + w, t + h)
    End If
End Function

Private Function Overlap(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim l As Double, t As Double, r As Double, bt As Double
    l = IIf(a(0) > b(0), a(0), b(0)): t = IIf(a(1) > b(1), a(1), b(1))
    r = IIf(a(2) < b(2), a(2), b(2)): bt = IIf(a(3) < b(3), a(3), b(3))
    If r - l > 3 And bt - t > 3 Then Overlap = Array(l, t, r, bt) Else Overlap = Empty
End Function

Private Function IsMarking(ByVal txt As String) As Boolean
    Dim w As Variant
    txt = UCase$(txt)
    For Each w In Array("CUI", "CONTROLLED", "CONFIDENTIAL", "PROPRIETARY", "UNCLASSIFIED", "SECRET", "ITAR", "EXPORT", "FOUO", "DISTRIBUTION")
        If InStr(txt, w) > 0 Then IsMarking = True: Exit Function
    Next
End Function

' The template elements a slide shows: Array(name, kind, box, severity).
Private Function Protected(ByVal sld As Object) As Collection
    Dim c As New Collection, lay As Object, sh As Object, showMaster As Boolean
    Set lay = sld.CustomLayout
    For Each sh In lay.Shapes
        AddProtected c, sh, "layout"
    Next
    On Error Resume Next
    showMaster = True
    showMaster = lay.DisplayMasterShapes
    On Error GoTo 0
    If showMaster Then
        For Each sh In sld.Design.SlideMaster.Shapes
            If sh.Type <> 14 Then AddProtected c, sh, "master"
        Next
    End If
    Set Protected = c
End Function

Private Sub AddProtected(ByVal c As Collection, ByVal sh As Object, ByVal src As String)
    Dim t As Long, txt As String
    On Error Resume Next
    If sh.Type = 14 Then
        t = sh.PlaceholderFormat.Type
        Select Case t
            Case PH_NUMBER: c.Add Array("Slide number (" & src & ")", "number", Box(sh), "High")
            Case PH_FOOTER: c.Add Array("Footer (" & src & ")", "footer", Box(sh), "High")
            Case PH_DATE: c.Add Array("Date (" & src & ")", "date", Box(sh), "Medium")
            Case PH_TITLE, PH_CTITLE: c.Add Array("Title area (" & src & ")", "title", Box(sh), "Medium")
        End Select
    Else
        If sh.HasTextFrame Then txt = sh.TextFrame.TextRange.Text
        If IsMarking(txt) Then
            c.Add Array("Marking """ & SKW_Excerpt(txt, 30) & """ (" & src & ")", "marking", Box(sh), "High")
        ElseIf sh.Visible Then
            c.Add Array(sh.Name & " (" & src & " graphic)", "graphic", Box(sh), "Medium")
        End If
    End If
End Sub

' The overlaps on one slide: each Array(slide shape name, element, severity, overlap box, percent of the element covered).
Private Function SlideOverlaps(ByVal sld As Object, ByVal slideW As Double, ByVal slideH As Double) As Collection
    Dim res As New Collection, prot As Collection, sh As Object, items As Collection, it As Object, pr As Variant, b As Variant, o As Variant
    Dim pt As Long, pct As Long, sev As String, ea As Double
    Set prot = Protected(sld)
    For Each sh In sld.Shapes
        Set items = New Collection
        If sh.Type = 6 Then
            For Each it In sh.GroupItems
                items.Add it
            Next
        Else
            items.Add sh
        End If
        For Each it In items
            pt = 0
            On Error Resume Next
            If it.Type = 14 Then pt = it.PlaceholderFormat.Type
            On Error GoTo 0
            b = Box(it)
            For Each pr In prot
                ' A slide's own placeholder sits where the layout's does: that is not covering it.
                If Not ((pr(1) = "number" And pt = PH_NUMBER) Or (pr(1) = "footer" And pt = PH_FOOTER) Or (pr(1) = "date" And pt = PH_DATE) _
                        Or (pr(1) = "title" And (pt = PH_TITLE Or pt = PH_CTITLE Or pt = PH_SUBTITLE))) Then
                    o = Overlap(b, pr(2))
                    If Not IsEmpty(o) Then
                        ea = (pr(2)(2) - pr(2)(0)) * (pr(2)(3) - pr(2)(1))
                        If ea > 0 Then pct = CLng(100 * (o(2) - o(0)) * (o(3) - o(1)) / ea) Else pct = 100
                        sev = pr(3)
                        If (b(2) - b(0)) >= 0.9 * slideW And (b(3) - b(1)) >= 0.9 * slideH Then sev = "Review (full-slide picture or shape)"
                        res.Add Array(it.Name, pr(0), sev, o, pct)
                    End If
                End If
            Next
        Next
    Next
    Set SlideOverlaps = res
End Function

' Report, plus (target "", or a path) a copy with each overlap outlined in red. target "-" makes no copy.
Public Function SKP_DeckOverlapCheck(ByVal target As String) As String
    Dim p As Object, c As Object, rows As New Collection, sld As Object, ov As Variant, hits As Object, nHigh As Long, n As Long
    Dim w As Double, h As Double, msg As String, mk As Object, i As Long
    Set SKW_LastReport = Nothing
    Set p = SKP_Deck.SKP_Deck()
    If p Is Nothing Then SKP_DeckOverlapCheck = SKP_Why(): Exit Function
    SKW_Status "Deck overlap check: " & p.Name & "..."
    w = p.PageSetup.SlideWidth: h = p.PageSetup.SlideHeight
    Set hits = CreateObject("Scripting.Dictionary")
    For Each sld In p.Slides
        For Each ov In SlideOverlaps(sld, w, h)
            n = n + 1
            If ov(2) = "High" Then nHigh = nHigh + 1
            rows.Add SKW_Row(sld.SlideIndex, ov(0), ov(1), ov(2), ov(4) & "%")
            hits(sld.SlideIndex & "|" & n) = ov(3)
        Next
    Next
    SKP_LastCopy = ""
    If n > 0 And target <> "-" Then
        msg = FixCopy(p, "overlaps marked", target, c)
        If Len(msg) = 0 Then
            On Error Resume Next
            i = 0
            For Each ov In hits.Keys
                i = i + 1
                Set mk = c.Slides(CLng(Split(ov, "|")(0))).Shapes.AddShape(1, hits(ov)(0), hits(ov)(1), hits(ov)(2) - hits(ov)(0), hits(ov)(3) - hits(ov)(1))
                mk.Fill.Visible = 0
                mk.Line.ForeColor.RGB = RGB(230, 0, 0)
                mk.Line.Weight = 2.5
                mk.Name = "SK overlap " & i
                mk.Tags.Add "SKFLAG", "1"
            Next
            On Error GoTo 0
            FinishCopy c
            msg = "Each overlap is outlined in red in " & SKP_LastCopy & "."
        Else
            msg = "No marked copy: " & msg
        End If
    End If
    SKW_Status ""
    SKW_ReportFor "Deck Overlap Check", n & " overlap(s) with template elements on the slides, " & nHigh & " covering the slide number, footer or a marking. " & msg & _
        " Elements come from each slide's layout and master; a slide's own number, footer and title placeholders are not counted. Full-slide pictures are marked Review. " & _
        "The open deck is not changed.", Array("Slide", "Slide shape", "Covers", "Severity", "Of the element"), rows, p.Name
    SKP_DeckOverlapCheck = n & " overlap(s), " & nHigh & " high." & IIf(Len(SKP_LastCopy) > 0, vbCr & "Marked copy: " & SKP_LastCopy, "")
End Function

' ---------------------------------------------------------------------
' Slide numbers
' ---------------------------------------------------------------------
Private Function NumberPh(ByVal shapes As Object) As Object
    Dim sh As Object
    On Error Resume Next
    For Each sh In shapes
        If sh.Type = 14 Then
            If sh.PlaceholderFormat.Type = PH_NUMBER Then Set NumberPh = sh: Exit Function
        End If
    Next
End Function

Private Function IsTitleSlide(ByVal sld As Object) As Boolean
    Dim sh As Object
    On Error Resume Next
    For Each sh In sld.CustomLayout.Shapes.Placeholders
        If sh.PlaceholderFormat.Type = PH_CTITLE Then IsTitleSlide = True: Exit Function
    Next
End Function

' Text boxes holding just the slide's number ("7", "Page 7", "7 of 20"): typed numbers that never update.
Private Function TypedNumbers(ByVal sld As Object) As Collection
    Dim c As New Collection, sh As Object, t As String, n As String
    n = CStr(sld.SlideNumber)
    For Each sh In sld.Shapes
        If sh.Type <> 14 And sh.HasTextFrame Then
            t = LCase$(Trim$(Replace(sh.TextFrame.TextRange.Text, vbCr, " ")))
            If t = n Or t = "page " & n Or t Like n & " of #*" Or t Like "page " & n & " of #*" Or t = "slide " & n Then c.Add sh
        End If
    Next
    Set TypedNumbers = c
End Function

Private Function NumberRules(ByRef first As Long, ByRef onTitle As Boolean) As String
    first = CLng(Val(SKW_Setting("DeckFirstSlideNumber", "1")))
    onTitle = SKW_IsYes(SKW_Setting("DeckNumberTitleSlide", "No"))
End Function

Public Function SKP_DeckSlideNumberCheck() As String
    Dim p As Object, rows As New Collection, sld As Object, lay As Object, ph As Object, lph As Object, d As Object, sh As Object
    Dim first As Long, onTitle As Boolean, want As Boolean, vis As Boolean, nMissing As Long, nTyped As Long, nDrift As Long, nLay As Long, seen As Object, n As Long
    Set SKW_LastReport = Nothing
    Set p = SKP_Deck.SKP_Deck()
    If p Is Nothing Then SKP_DeckSlideNumberCheck = SKP_Why(): Exit Function
    SKW_Status "Deck slide numbers: " & p.Name & "..."
    NumberRules first, onTitle
    If p.PageSetup.FirstSlideNumber <> first Then rows.Add SKW_Row("", "Deck", "Numbering starts at " & p.PageSetup.FirstSlideNumber, "The team starts at " & first, "")
    Set seen = CreateObject("Scripting.Dictionary")
    For Each sld In p.Slides
        Set lay = sld.CustomLayout
        Set lph = NumberPh(lay.Shapes)
        If lph Is Nothing And Not seen.Exists(lay.Design.Name & "|" & lay.Name) Then
            seen(lay.Design.Name & "|" & lay.Name) = True
            nLay = nLay + 1
            rows.Add SKW_Row("", "Layout " & lay.Name, "Has no slide number placeholder", "Slides on it cannot show a number", "")
        End If
        want = Not sld.SlideShowTransition.Hidden And (onTitle Or Not IsTitleSlide(sld))
        Set ph = NumberPh(sld.Shapes)
        vis = Not ph Is Nothing
        If want And Not vis Then
            nMissing = nMissing + 1
            rows.Add SKW_Row(sld.SlideIndex, SlideTitleOf(sld), "No slide number shows", IIf(lph Is Nothing, "Its layout has no number placeholder", "Insert > Slide Number, or Deck Slide Number Fix"), "")
        ElseIf vis And Not want And Not sld.SlideShowTransition.Hidden Then
            rows.Add SKW_Row(sld.SlideIndex, SlideTitleOf(sld), "Number on a title slide", "The team does not number title slides (setting DeckNumberTitleSlide)", "")
        End If
        If vis And Not lph Is Nothing Then
            If Abs(ph.Left - lph.Left) > 2 Or Abs(ph.Top - lph.Top) > 2 Or Abs(ph.Width - lph.Width) > 2 Or Abs(ph.Height - lph.Height) > 2 Then
                nDrift = nDrift + 1
                rows.Add SKW_Row(sld.SlideIndex, SlideTitleOf(sld), "Number moved or resized", "Not where the layout puts it", "")
            ElseIf FontOf(ph) <> FontOf(lph) Then
                nDrift = nDrift + 1
                rows.Add SKW_Row(sld.SlideIndex, SlideTitleOf(sld), "Number restyled", FontOf(ph) & " (layout: " & FontOf(lph) & ")", "")
            End If
        End If
        For Each sh In TypedNumbers(sld)
            nTyped = nTyped + 1
            rows.Add SKW_Row(sld.SlideIndex, SlideTitleOf(sld), "Number typed as text", """" & sh.TextFrame.TextRange.Text & """ in " & sh.Name & " will not change when slides move", "")
        Next
    Next
    SKW_Status ""
    n = nMissing + nTyped + nDrift + nLay
    SKW_ReportFor "Deck Slide Number Check", nMissing & " slide(s) without a number, " & nLay & " layout(s) with no number placeholder, " & nTyped & " typed number(s), " & _
        nDrift & " number(s) moved or restyled. Starts at " & p.PageSetup.FirstSlideNumber & ". Team rules: start at " & first & "; title slides " & IIf(onTitle, "numbered", "not numbered") & _
        "; hidden slides count but need no number. Changes nothing.", Array("Slide", "Title or layout", "Finding", "Detail", ""), rows, p.Name
    SKW_Fixable = nMissing + nLay + nTyped + nDrift
    SKP_DeckSlideNumberCheck = nMissing & " slide(s) without a number, " & nLay & " layout(s) without the placeholder, " & nTyped & " typed, " & nDrift & " moved or restyled."
End Function

Private Function FontOf(ByVal sh As Object) As String
    On Error Resume Next
    With sh.TextFrame.TextRange.Font
        FontOf = .Name & " " & .Size & " pt"
    End With
End Function

Public Function SKP_DeckSlideNumberFix(ByVal target As String) As String
    Dim p As Object, c As Object, msg As String, rows As New Collection, d As Object, lay As Object, mph As Object, lph As Object, sld As Object, ph As Object, sh As Object
    Dim first As Long, onTitle As Boolean, want As Boolean, nAdded As Long, nOn As Long, nOff As Long, nTyped As Long, nReset As Long, i As Long, typed As Collection
    Set SKW_LastReport = Nothing
    Set p = SKP_Deck.SKP_Deck()
    If p Is Nothing Then SKP_DeckSlideNumberFix = SKP_Why(): Exit Function
    msg = FixCopy(p, "fixed", target, c)
    If Len(msg) > 0 Then SKP_DeckSlideNumberFix = msg: Exit Function
    On Error GoTo Fail
    SKW_Status "Fixing slide numbers..."
    NumberRules first, onTitle
    If c.PageSetup.FirstSlideNumber <> first Then c.PageSetup.FirstSlideNumber = first: rows.Add SKW_Row("", "Deck", "Numbering starts at " & first, "")
    ' 1. The placeholder on every master and layout.
    For Each d In c.Designs
        Set mph = NumberPh(d.SlideMaster.Shapes)
        If mph Is Nothing Then
            On Error Resume Next
            Set mph = d.SlideMaster.Shapes.AddPlaceholder(PH_NUMBER, c.PageSetup.SlideWidth - 120, c.PageSetup.SlideHeight - 40, 96, 28)
            If Err.Number = 0 Then nAdded = nAdded + 1: rows.Add SKW_Row("", "Master " & d.Name, "Slide number placeholder put back", "")
            Err.Clear
            On Error GoTo Fail
        End If
        For Each lay In d.SlideMaster.CustomLayouts
            If NumberPh(lay.Shapes) Is Nothing And Not mph Is Nothing Then
                On Error Resume Next
                lay.Shapes.AddPlaceholder PH_NUMBER, mph.Left, mph.Top, mph.Width, mph.Height
                If Err.Number = 0 Then nAdded = nAdded + 1: rows.Add SKW_Row("", "Layout " & lay.Name, "Slide number placeholder put back", "") _
                    Else rows.Add SKW_Row("", "Layout " & lay.Name, "Could not add the placeholder", Err.Description)
                Err.Clear
                On Error GoTo Fail
            End If
        Next
    Next
    ' 2. Each slide: typed numbers out, number on or off, position and font from the layout.
    For Each sld In c.Slides
        Set typed = TypedNumbers(sld)
        For i = typed.Count To 1 Step -1
            rows.Add SKW_Row(sld.SlideIndex, SlideTitleOf(sld), "Typed number removed", """" & typed(i).TextFrame.TextRange.Text & """")
            typed(i).Delete
            nTyped = nTyped + 1
        Next
        want = Not sld.SlideShowTransition.Hidden And (onTitle Or Not IsTitleSlide(sld))
        Set ph = NumberPh(sld.Shapes)
        If want And ph Is Nothing Then
            sld.HeadersFooters.SlideNumber.Visible = 0
            sld.HeadersFooters.SlideNumber.Visible = -1
            Set ph = NumberPh(sld.Shapes)
            If Not ph Is Nothing Then nOn = nOn + 1: rows.Add SKW_Row(sld.SlideIndex, SlideTitleOf(sld), "Number turned on", "") _
                Else rows.Add SKW_Row(sld.SlideIndex, SlideTitleOf(sld), "Could not show a number", "Its layout still has no number placeholder")
        ElseIf Not want And Not ph Is Nothing And Not sld.SlideShowTransition.Hidden Then
            sld.HeadersFooters.SlideNumber.Visible = 0
            nOff = nOff + 1
            rows.Add SKW_Row(sld.SlideIndex, SlideTitleOf(sld), "Number taken off the title slide", "")
            Set ph = Nothing
        End If
        If Not ph Is Nothing Then
            Set lph = NumberPh(sld.CustomLayout.Shapes)
            If Not lph Is Nothing Then
                If Abs(ph.Left - lph.Left) > 2 Or Abs(ph.Top - lph.Top) > 2 Or Abs(ph.Width - lph.Width) > 2 Or Abs(ph.Height - lph.Height) > 2 Or FontOf(ph) <> FontOf(lph) Then
                    ph.Left = lph.Left: ph.Top = lph.Top: ph.Width = lph.Width: ph.Height = lph.Height
                    On Error Resume Next
                    ph.TextFrame.TextRange.Font.Name = lph.TextFrame.TextRange.Font.Name
                    ph.TextFrame.TextRange.Font.Size = lph.TextFrame.TextRange.Font.Size
                    ph.TextFrame.TextRange.Font.Color.RGB = lph.TextFrame.TextRange.Font.Color.RGB
                    On Error GoTo Fail
                    nReset = nReset + 1
                    rows.Add SKW_Row(sld.SlideIndex, SlideTitleOf(sld), "Number put back where the layout has it, in its font", "")
                End If
            End If
        End If
    Next
    FinishCopy c
    SKW_ReportFor "Deck Slide Number Fix", nAdded & " placeholder(s) put back, " & nOn & " number(s) turned on, " & nOff & " taken off title slides, " & nTyped & _
        " typed number(s) removed, " & nReset & " reset to the layout. The numbers are fields now, so they follow the slides when they move. The open deck is not changed. " & _
        "The fixed copy: " & SKP_LastCopy, Array("Slide", "Title or layout", "What was done", "Detail"), rows, p.Name
    SKP_DeckSlideNumberFix = nAdded & " placeholder(s) back, " & nOn & " on, " & nOff & " off, " & nTyped & " typed removed, " & nReset & " reset." & vbCr & "Fixed copy: " & SKP_LastCopy
    Exit Function
Fail:
    SKP_DeckSlideNumberFix = "Slide number fix stopped: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Function

' ---------------------------------------------------------------------
' Banned phrases, replaced on a copy with a comment for each
' ---------------------------------------------------------------------
Private Function Proper1(ByVal s As String) As String
    If Len(s) > 0 Then Proper1 = UCase$(Left$(s, 1)) & Mid$(s, 2)
End Function

' Replaces every whole-word use of ph in tr; returns how many.
Private Function ReplaceAll(ByVal tr As Object, ByVal ph As String, ByVal repl As String, ByVal matchCase As Boolean, ByVal whole As Boolean) As Long
    Dim hit As Object, guard As Long
    Do
        Set hit = Nothing
        On Error Resume Next
        Set hit = tr.Replace(FindWhat:=ph, ReplaceWhat:=repl, MatchCase:=IIf(matchCase, -1, 0), WholeWords:=IIf(whole, -1, 0))
        On Error GoTo 0
        If hit Is Nothing Then Exit Do
        ReplaceAll = ReplaceAll + 1
        guard = guard + 1
        If guard > 500 Then Exit Do
    Loop
End Function

Private Function FlagCount(ByVal tr As Object, ByVal ph As String, ByVal matchCase As Boolean, ByVal whole As Boolean) As Long
    Dim hit As Object, after As Long, guard As Long
    Do
        Set hit = Nothing
        On Error Resume Next
        Set hit = tr.Find(FindWhat:=ph, After:=after, MatchCase:=IIf(matchCase, -1, 0), WholeWords:=IIf(whole, -1, 0))
        On Error GoTo 0
        If hit Is Nothing Then Exit Do
        FlagCount = FlagCount + 1
        after = hit.Start + hit.Length - 1
        guard = guard + 1
        If guard > 500 Then Exit Do
    Loop
End Function

Public Function SKP_DeckBannedPhrasesFix(ByVal target As String) As String
    Dim p As Object, c As Object, msg As String, rows As New Collection, rules As Collection, ru As Object, sld As Object, sh As Object, tr As Object
    Dim ph As String, repl As String, mc As Boolean, whole As Boolean, n As Long, notes As Object, k As Variant, nChanged As Long, nFlagged As Long, rid As String
    Set SKW_LastReport = Nothing
    Set p = SKP_Deck.SKP_Deck()
    If p Is Nothing Then SKP_DeckBannedPhrasesFix = SKP_Why(): Exit Function
    Set rules = SKW_BannedPhrases()
    If rules.Count = 0 Then SKP_DeckBannedPhrasesFix = "No banned phrases in the team standards.": Exit Function
    msg = FixCopy(p, "fixed", target, c)
    If Len(msg) > 0 Then SKP_DeckBannedPhrasesFix = msg: Exit Function
    On Error GoTo Fail
    SKW_Status "Replacing banned phrases..."
    For Each sld In c.Slides
        Set notes = CreateObject("Scripting.Dictionary")
        For Each sh In SKP_Shapes(sld)
            For Each tr In SKP_TextRanges(sh)
                For Each ru In rules
                    ph = ru("phrase")
                    repl = "": If ru.Exists("replacement") Then repl = ru("replacement")
                    mc = False: If ru.Exists("match case") Then mc = SKW_IsYes(ru("match case"))
                    whole = True: If ru.Exists("whole word") Then If Len(ru("whole word")) > 0 Then whole = SKW_IsYes(ru("whole word"))
                    rid = "": If ru.Exists("rule id") Then rid = ru("rule id")
                    If Len(repl) > 0 Then
                        n = 0
                        If Not mc Then n = ReplaceAll(tr, Proper1(ph), Proper1(repl), True, whole)   ' keep a capital at the start of a sentence
                        n = n + ReplaceAll(tr, ph, repl, mc, whole)
                        If n > 0 Then
                            nChanged = nChanged + n
                            k = "Replaced """ & ph & """ with """ & repl & """" & IIf(Len(rid) > 0, " (rule " & rid & ")", "")
                            notes(k) = notes(k) + n
                            rows.Add SKW_Row(sld.SlideIndex, sh.Name, ph, repl, n, rid)
                        End If
                    Else
                        n = FlagCount(tr, ph, mc, whole)
                        If n > 0 Then
                            nFlagged = nFlagged + n
                            k = "Reword """ & ph & """" & IIf(Len(rid) > 0, " (rule " & rid & ")", "")
                            notes(k) = notes(k) + n
                            rows.Add SKW_Row(sld.SlideIndex, sh.Name, ph, "(reword by hand)", n, rid)
                        End If
                    End If
                Next
            Next
        Next
        If notes.Count > 0 Then AddComment sld, notes
    Next
    FinishCopy c
    SKW_ReportFor "Deck Banned Phrases Fix", nChanged & " phrase(s) replaced and " & nFlagged & " flagged to reword. Each slide with a change has a comment listing what " & _
        "changed (PowerPoint has no tracked changes). Text can grow: run Deck Small Text afterwards. Deck Delivery Copy removes the comments. The open deck is not changed. " & _
        "The fixed copy: " & SKP_LastCopy, Array("Slide", "Shape", "Phrase", "Now", "Times", "Rule"), rows, p.Name
    SKP_DeckBannedPhrasesFix = nChanged & " replaced, " & nFlagged & " flagged." & vbCr & "Fixed copy: " & SKP_LastCopy
    Exit Function
Fail:
    SKP_DeckBannedPhrasesFix = "Banned phrases fix stopped: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Function

Private Sub AddComment(ByVal sld As Object, ByVal notes As Object)
    Dim t As String, k As Variant
    t = "Stationkeeper:"
    For Each k In notes.Keys
        t = t & " " & k & IIf(notes(k) > 1, " x" & notes(k), "") & "."
    Next
    On Error Resume Next
    sld.Comments.Add2 Left:=10, Top:=10, Author:="Stationkeeper", AuthorInitials:="SK", Text:=t, ProviderID:="None", UserID:="Stationkeeper"
    If Err.Number <> 0 Then
        Err.Clear
        sld.Comments.Add 10, 10, "Stationkeeper", "SK", t
    End If
End Sub

' ---------------------------------------------------------------------
' The media list (used by frmSKPMedia)
' ---------------------------------------------------------------------
' Each item: slide, id, name, kind, width and height in inches, alt text, link.
Public Function SKP_MediaItems(ByVal p As Object) As Collection
    Dim c As New Collection, sld As Object, sh As Object, g As Object
    For Each sld In p.Slides
        For Each sh In sld.Shapes
            AddMedia c, sld, sh
            If sh.Type = 6 Then
                For Each g In sh.GroupItems
                    AddMedia c, sld, g
                Next
            End If
        Next
    Next
    Set SKP_MediaItems = c
End Function

Public Function SKP_MediaKind(ByVal sh As Object) As String
    Dim t As Long, ct As Long
    On Error Resume Next
    t = sh.Type
    If t = 14 Then ct = sh.PlaceholderFormat.ContainedType: If ct > 0 And ct <> 1 And ct <> 17 Then t = ct
    Select Case t
        Case 13: SKP_MediaKind = "Picture"
        Case 11: SKP_MediaKind = "Linked picture"
        Case 28: SKP_MediaKind = "Icon or SVG"
        Case 29: SKP_MediaKind = "Linked graphic"
        Case 16
            Select Case sh.MediaType
                Case 2: SKP_MediaKind = "Audio"
                Case 3: SKP_MediaKind = "Video"
                Case Else: SKP_MediaKind = "Media"
            End Select
        Case 7: SKP_MediaKind = "Embedded object"
        Case 10: SKP_MediaKind = "Linked object"
        Case 3: SKP_MediaKind = "Chart"
        Case 24: SKP_MediaKind = "SmartArt"
        Case 26: SKP_MediaKind = "Web video"
    End Select
    If Len(SKP_MediaKind) = 0 Then
        If sh.HasChart Then SKP_MediaKind = "Chart"
        If sh.HasSmartArt Then SKP_MediaKind = "SmartArt"
    End If
End Function

Private Sub AddMedia(ByVal c As Collection, ByVal sld As Object, ByVal sh As Object)
    Dim d As Object, k As String
    k = SKP_MediaKind(sh)
    If Len(k) = 0 Then Exit Sub
    Set d = CreateObject("Scripting.Dictionary")
    d("slide") = sld.SlideIndex
    d("slideId") = sld.SlideID
    d("id") = sh.Id
    d("name") = sh.Name
    d("kind") = k
    d("size") = Format$(sh.Width / 72, "0.0") & " x " & Format$(sh.Height / 72, "0.0") & " in"
    d("alt") = ""
    d("link") = ""
    On Error Resume Next
    d("alt") = sh.AlternativeText
    d("link") = sh.LinkFormat.SourceFullName
    c.Add d
End Sub

' The shape by slide ID and shape ID (groups opened up), or Nothing.
Public Function SKP_FindShape(ByVal p As Object, ByVal slideId As Long, ByVal shapeId As Long) As Object
    Dim sld As Object, sh As Object, g As Object
    On Error Resume Next
    Set sld = p.Slides.FindBySlideID(slideId)
    If sld Is Nothing Then Exit Function
    For Each sh In sld.Shapes
        If sh.Id = shapeId Then Set SKP_FindShape = sh: Exit Function
        If sh.Type = 6 Then
            For Each g In sh.GroupItems
                If g.Id = shapeId Then Set SKP_FindShape = g: Exit Function
            Next
        End If
    Next
End Function

' Renames a shape. Returns "" or why not. Names must be unique on their slide.
Public Function SKP_RenameMedia(ByVal p As Object, ByVal slideId As Long, ByVal shapeId As Long, ByVal newName As String) As String
    Dim sh As Object, other As Object, sld As Object
    newName = Trim$(Replace(Replace(newName, vbCr, " "), vbLf, " "))
    If Len(newName) = 0 Then SKP_RenameMedia = "Type a name first.": Exit Function
    Set sh = SKP_FindShape(p, slideId, shapeId)
    If sh Is Nothing Then SKP_RenameMedia = "That item is gone (deleted or moved). Refresh the list.": Exit Function
    Set sld = p.Slides.FindBySlideID(slideId)
    For Each other In SKP_Shapes(sld)
        If other.Id <> shapeId And LCase$(other.Name) = LCase$(newName) Then SKP_RenameMedia = "Slide " & sld.SlideIndex & " already has a shape named " & newName & ".": Exit Function
    Next
    On Error GoTo Fail
    sh.Name = newName
    Exit Function
Fail:
    SKP_RenameMedia = "PowerPoint would not rename it: " & Err.Description
End Function

' A name from the alt text (first words), the kind and the slide: "S3 Picture - Rocket launch".
Public Function SKP_SuggestName(ByVal it As Object) As String
    Dim a As String, w As Variant, s As String, n As Long
    a = Trim$(Replace(Replace(it("alt"), vbCr, " "), vbLf, " "))
    For Each w In Split(a, " ")
        If Len(w) > 0 Then
            s = s & IIf(Len(s) > 0, " ", "") & w
            n = n + 1
            If n = 6 Or Len(s) > 40 Then Exit For
        End If
    Next
    If Len(s) = 0 And Len(it("link")) > 0 Then s = Mid$(it("link"), InStrRev(it("link"), "\") + 1)
    SKP_SuggestName = "S" & it("slide") & " " & it("kind") & IIf(Len(s) > 0, " - " & s, " " & it("id"))
End Function

' Shows the slide and selects the shape in PowerPoint (needs the deck in a window).
Public Function SKP_GoToMedia(ByVal p As Object, ByVal slideId As Long, ByVal shapeId As Long) As String
    Dim sh As Object, sld As Object, w As Object
    Set sh = SKP_FindShape(p, slideId, shapeId)
    If sh Is Nothing Then SKP_GoToMedia = "That item is gone. Refresh the list.": Exit Function
    On Error GoTo Fail
    If p.Windows.Count = 0 Then SKP_GoToMedia = "The deck has no window to show it in.": Exit Function
    Set sld = p.Slides.FindBySlideID(slideId)
    Set w = p.Windows(1)
    w.Activate
    w.View.GotoSlide sld.SlideIndex
    sh.Select
    Exit Function
Fail:
    SKP_GoToMedia = "PowerPoint could not show it: " & Err.Description
End Function

' ---------------------------------------------------------------------
' Shared with SKP_Tidy (1.5.0)
' ---------------------------------------------------------------------
Public Function SKP_FixCopy(ByVal p As Object, ByVal what As String, ByVal target As String, ByRef c As Object) As String
    SKP_FixCopy = FixCopy(p, what, target, c)
End Function

Public Sub SKP_FinishCopy(ByVal c As Object)
    FinishCopy c
End Sub

Public Function SKP_SlideTitleOf(ByVal sld As Object) As String
    SKP_SlideTitleOf = SlideTitleOf(sld)
End Function

Public Sub SKP_AddComment(ByVal sld As Object, ByVal notes As Object)
    AddComment sld, notes
End Sub

Public Function SKP_IsTitleSlide(ByVal sld As Object) As Boolean
    SKP_IsTitleSlide = IsTitleSlide(sld)
End Function
