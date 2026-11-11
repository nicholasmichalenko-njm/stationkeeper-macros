Attribute VB_Name = "SKP_Deck"
' =====================================================================
' Stationkeeper Tools for PowerPoint, run from Word.
' They check the deck open in PowerPoint (they never start PowerPoint or open
' a file on their own), write their reports as Word documents, and change a
' deck only in a copy saved beside it.
'   DeckTextLeftovers, DeckBannedPhrases, DeckAcronymCheck, DeckFontAudit,
'   DeckSmallText, DeckAltTextAudit, DeckOffSlideObjects, DeckTitlesCheck,
'   DeckEmptyPlaceholders, DeckDeliveryCopy
' =====================================================================
Option Explicit

Public SKP_TestDeck As Object      ' tests: the deck to check, so a person's open decks are never touched
Private mWhich As String

' Lists speaker notes, comments, hidden slides and placeholder text (TBD, XXX, [insert) in the deck open in PowerPoint. Changes nothing.
Public Sub DeckTextLeftovers()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    SKW_Show SKP_DeckTextLeftovers()
End Sub

' Lists the team's banned phrases found in the slides and notes of the open deck, with the suggested replacement. Changes nothing.
Public Sub DeckBannedPhrases()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    SKW_Show SKP_DeckBannedPhrases()
End Sub

' Lists the acronyms in the open deck's slide text with their uses, the first slide that spells them out, and the team meaning. Text only: pictures are not read. Changes nothing.
Public Sub DeckAcronymCheck()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    SKW_Show SKP_DeckAcronymCheck()
End Sub

' Lists every font used in the open deck and flags the ones not on the team's Fonts sheet. Changes nothing.
Public Sub DeckFontAudit()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    SKW_Show SKP_DeckFontAudit()
End Sub

' Lists text in the open deck smaller than the team minimum (setting DeckMinFontSize, 10 pt unless set). Changes nothing.
Public Sub DeckSmallText()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    SKW_Show SKP_DeckSmallText()
End Sub

' Lists pictures, charts, SmartArt, groups and media in the open deck that have no alt text for screen readers. Changes nothing.
Public Sub DeckAltTextAudit()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    SKW_Show SKP_DeckAltTextAudit()
End Sub

' Lists shapes that sit partly or wholly off the slide in the open deck. They still print or show in some views. Changes nothing.
Public Sub DeckOffSlideObjects()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    SKW_Show SKP_DeckOffSlideObjects()
End Sub

' Lists slides with no title (screen readers and the outline need one) and slides that share a title. Changes nothing.
Public Sub DeckTitlesCheck()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    SKW_Show SKP_DeckTitlesCheck()
End Sub

' Lists empty placeholders and "Click to add" boxes left on slides of the open deck. Changes nothing.
Public Sub DeckEmptyPlaceholders()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    SKW_Show SKP_DeckEmptyPlaceholders()
End Sub

' Saves a delivery copy of the open deck beside it ("<name> - delivery.pptx"): no speaker notes, no comments, no hidden slides, no personal details. The open deck is not changed.
Public Sub DeckDeliveryCopy()
    If SKW_IsQuiet() Then Exit Sub   ' tests run the SKP_ functions on their own deck
    SKW_Say SKP_DeckDeliveryCopy("")
End Sub

' ---------------------------------------------------------------------
' The deck open in PowerPoint: the active one, or the only one open.
' Returns Nothing (and the reason in mWhich) when there is none to use.
' ---------------------------------------------------------------------
Public Function SKP_Deck() As Object
    Dim pp As Object, p As Object
    mWhich = ""
    If Not SKP_TestDeck Is Nothing Then Set SKP_Deck = SKP_TestDeck: Exit Function
    On Error Resume Next
    Set pp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    If pp Is Nothing Then mWhich = "Open the deck in PowerPoint first. The tools check the deck that is open; they never open one on their own.": Exit Function
    If pp.Presentations.Count = 0 Then mWhich = "Open the deck in PowerPoint first.": Exit Function
    On Error Resume Next
    Set p = pp.ActivePresentation
    On Error GoTo 0
    If p Is Nothing Then
        If pp.Presentations.Count = 1 Then
            Set p = pp.Presentations(1)
        Else
            mWhich = "Several decks are open. Click in the one to check, then run this again."
            Exit Function
        End If
    End If
    Set SKP_Deck = p
End Function

Public Function SKP_Why() As String
    SKP_Why = mWhich
End Function

' Every shape on a slide, groups opened up.
Private Sub AddShapes(ByVal c As Collection, ByVal shapes As Object)
    Dim sh As Object, g As Object
    For Each sh In shapes
        c.Add sh
        If sh.Type = 6 Then   ' group
            For Each g In sh.GroupItems
                c.Add g
            Next
        End If
    Next
End Sub

Public Function SKP_Shapes(ByVal sld As Object) As Collection
    Dim c As New Collection
    AddShapes c, sld.Shapes
    Set SKP_Shapes = c
End Function

' The text pieces of a shape: its text frame, or each cell of its table.
Public Function SKP_TextRanges(ByVal sh As Object) As Collection
    Dim c As New Collection, r As Long, k As Long, nR As Long, nK As Long
    On Error Resume Next
    If sh.HasTable Then
        nR = sh.Table.Rows.Count: nK = sh.Table.Columns.Count   ' read first: a failed loop bound under Resume Next can loop forever
        For r = 1 To nR
            For k = 1 To nK
                c.Add sh.Table.Cell(r, k).Shape.TextFrame.TextRange
            Next
        Next
    ElseIf sh.HasTextFrame Then
        If sh.TextFrame.HasText Then c.Add sh.TextFrame.TextRange
    End If
    Set SKP_TextRanges = c
End Function

Private Function NotesText(ByVal sld As Object) As String
    Dim sh As Object
    On Error Resume Next
    For Each sh In sld.NotesPage.Shapes
        If sh.Type = 14 Then
            If sh.PlaceholderFormat.Type = 2 Then NotesText = sh.TextFrame.TextRange.Text: Exit Function   ' body = notes
        End If
    Next
End Function

Private Function SlideText(ByVal sld As Object) As String
    Dim sh As Object, tr As Object, s As String
    For Each sh In SKP_Shapes(sld)
        For Each tr In SKP_TextRanges(sh)
            s = s & tr.Text & vbCr
        Next
    Next
    SlideText = s
End Function

Private Function SlideTitle(ByVal sld As Object) As String
    On Error Resume Next
    If sld.Shapes.HasTitle Then SlideTitle = Trim$(Replace(Replace(sld.Shapes.Title.TextFrame.TextRange.Text, vbCr, " "), Chr$(11), " "))
End Function

Private Function DeckName(ByVal p As Object) As String
    DeckName = p.Name
End Function

Private Function StartDeck(ByVal title As String, ByRef p As Object) As String
    Set SKW_LastReport = Nothing
    Set p = SKP_Deck()
    If p Is Nothing Then StartDeck = SKP_Why(): Exit Function
    SKW_Status title & ": " & p.Name & "..."
End Function

' ---------------------------------------------------------------------
Public Function SKP_DeckTextLeftovers() As String
    Dim p As Object, msg As String, rows As New Collection, sld As Object, t As String, ph As Variant, c As Object
    Dim nNotes As Long, nCom As Long, nHid As Long, nPh As Long
    msg = StartDeck("Deck leftovers", p)
    If Len(msg) > 0 Then SKP_DeckTextLeftovers = msg: Exit Function
    For Each sld In p.Slides
        If sld.SlideShowTransition.Hidden Then nHid = nHid + 1: rows.Add SKW_Row(sld.SlideIndex, "Hidden slide", "Hidden slides still travel with the file", SlideTitle(sld))
        t = Trim$(NotesText(sld))
        If Len(t) > 0 Then nNotes = nNotes + 1: rows.Add SKW_Row(sld.SlideIndex, "Speaker notes", "Notes travel with the file", t)
        On Error Resume Next
        For Each c In sld.Comments
            nCom = nCom + 1
            rows.Add SKW_Row(sld.SlideIndex, "Comment", "Comment by " & c.Author, c.Text)
        Next
        On Error GoTo 0
        t = SlideText(sld) & vbCr & NotesText(sld)
        For Each ph In SKW_Placeholders()
            If InStr(1, t, CStr(ph), vbTextCompare) > 0 Then nPh = nPh + 1: rows.Add SKW_Row(sld.SlideIndex, "Placeholder", "Contains " & ph, SlideTitle(sld))
        Next
    Next
    SKW_Status ""
    SKW_ReportFor "Deck Leftovers", "Speaker notes on " & nNotes & " slide(s), comments " & nCom & ", hidden slides " & nHid & ", placeholders " & nPh & _
        ". Deck Delivery Copy saves a copy without notes, comments and hidden slides. Changes nothing.", Array("Slide", "Rule", "Finding", "Text"), rows, DeckName(p)
    SKP_DeckTextLeftovers = "Notes " & nNotes & ", comments " & nCom & ", hidden slides " & nHid & ", placeholders " & nPh & "."
End Function

Public Function SKP_DeckBannedPhrases() As String
    Dim p As Object, msg As String, rows As New Collection, sld As Object, ru As Object, t As String, n As Long, where As Variant, cmp As Long
    msg = StartDeck("Deck banned phrases", p)
    If Len(msg) > 0 Then SKP_DeckBannedPhrases = msg: Exit Function
    For Each sld In p.Slides
        For Each where In Array("Slide", "Notes")
            If where = "Slide" Then t = SlideText(sld) Else t = NotesText(sld)
            For Each ru In SKW_BannedPhrases()
                cmp = IIf(SKW_IsYes(IIf(ru.Exists("match case"), ru("match case"), "")), vbBinaryCompare, vbTextCompare)
                If PhraseIn(t, ru("phrase"), cmp, SKW_IsYes(IIf(ru.Exists("whole word"), ru("whole word"), ""))) Then
                    n = n + 1
                    rows.Add SKW_Row(sld.SlideIndex, where, ru("phrase"), IIf(ru.Exists("replacement"), ru("replacement"), ""), _
                        IIf(ru.Exists("rule id"), ru("rule id"), "") & " " & IIf(ru.Exists("reason"), ru("reason"), ""))
                End If
            Next
        Next
    Next
    SKW_Status ""
    SKW_ReportFor "Deck Banned Phrases", n & " finding(s) against " & SKW_BannedPhrases().Count & " banned phrase(s). PowerPoint has no tracked changes, so this lists them to fix by hand. Changes nothing.", _
        Array("Slide", "Where", "Phrase", "Use instead", "Rule"), rows, DeckName(p)
    SKP_DeckBannedPhrases = n & " banned phrase finding(s)."
End Function

Private Function PhraseIn(ByVal t As String, ByVal ph As String, ByVal cmp As Long, ByVal whole As Boolean) As Boolean
    Dim k As Long, a As String, b As String
    k = InStr(1, t, ph, cmp)
    Do While k > 0
        If Not whole Then PhraseIn = True: Exit Function
        If k > 1 Then a = Mid$(t, k - 1, 1) Else a = " "
        b = Mid$(t & " ", k + Len(ph), 1)
        If Not (a Like "[A-Za-z0-9]") And Not (b Like "[A-Za-z0-9]") Then PhraseIn = True: Exit Function
        k = InStr(k + 1, t, ph, cmp)
    Loop
End Function

Public Function SKP_DeckAcronymCheck() As String
    Dim p As Object, msg As String, rows As New Collection, sld As Object, list As Object, info As Object, t As String
    Dim words As Variant, w As Variant, k As Variant, e As Variant, i As Long, nNot As Long, defined As Boolean
    msg = StartDeck("Deck acronym check", p)
    If Len(msg) > 0 Then SKP_DeckAcronymCheck = msg: Exit Function
    Set list = SKW_AcronymList()
    Set info = CreateObject("Scripting.Dictionary")
    For Each sld In p.Slides
        t = SlideText(sld)
        For Each w In Tokens(t)
            If Not info.Exists(w) Then info(w) = Array(0, sld.SlideIndex, "")
            e = info(w)
            e(0) = e(0) + 1
            If Len(e(2)) = 0 Then
                If InStr(t, "(" & w & ")") > 0 Or InStr(t, "(" & w & "s)") > 0 Then e(2) = CStr(sld.SlideIndex)
            End If
            info(w) = e
        Next
    Next
    For Each k In info.Keys
        e = info(k)
        If Len(e(2)) = 0 Then nNot = nNot + 1
        rows.Add SKW_Row(k, e(0), e(1), IIf(Len(e(2)) > 0, "slide " & e(2), "not spelled out"), IIf(list.Exists(k), "Yes", "No"), IIf(list.Exists(k), list(k), ""))
    Next
    SKW_Status ""
    SKW_ReportFor "Deck Acronym Check (text only)", "Text only: pictures and images are not scanned. " & info.Count & " acronym(s); " & nNot & _
        " never spelled out as Words (ACR) on a slide. Changes nothing.", Array("Acronym", "Uses", "First slide", "Spelled out", "On the team list", "Team meaning"), rows, DeckName(p)
    SKP_DeckAcronymCheck = info.Count & " acronym(s), " & nNot & " not spelled out (text only)."
End Function

' Acronym-like words: 2 to 10 characters, capitals and digits, at least two capitals, not Roman numerals.
Private Function Tokens(ByVal t As String) As Collection
    Dim c As New Collection, i As Long, ch As String, cur As String, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    t = t & " "
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch Like "[A-Z0-9&]" Then
            cur = cur & ch
        Else
            If Len(cur) >= 2 And Len(cur) <= 10 And cur Like "[A-Z]*" And cur Like "*[A-Z]*[A-Z]*" And Not (ch Like "[a-z]") Then
                If Not IsRomanP(cur) Then c.Add cur
            End If
            cur = ""
            If ch Like "[a-z]" Then   ' skip the rest of a mixed-case word
                Do While i < Len(t) And Mid$(t, i + 1, 1) Like "[A-Za-z0-9&]"
                    i = i + 1
                Loop
            End If
        End If
    Next
    Set Tokens = c
End Function

Private Function IsRomanP(ByVal t As String) As Boolean
    Dim i As Long
    For i = 1 To Len(t)
        If InStr("IVXLCDM", Mid$(t, i, 1)) = 0 Then Exit Function
    Next
    IsRomanP = True
End Function

Public Function SKP_DeckFontAudit() As String
    Dim p As Object, msg As String, rows As New Collection, sld As Object, sh As Object, tr As Object, rn As Object
    Dim fonts As Object, allowed As New Collection, r As Object, k As Variant, e As Variant, nm As String, nBad As Long
    msg = StartDeck("Deck fonts", p)
    If Len(msg) > 0 Then SKP_DeckFontAudit = msg: Exit Function
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
    Set fonts = CreateObject("Scripting.Dictionary")
    For Each sld In p.Slides
        For Each sh In SKP_Shapes(sld)
            For Each tr In SKP_TextRanges(sh)
                For Each rn In tr.Runs
                    nm = rn.Font.Name
                    If Len(nm) > 0 And Len(Trim$(rn.Text)) > 0 Then
                        If fonts.Exists(nm) Then
                            e = fonts(nm): e(0) = e(0) + 1: fonts(nm) = e
                        Else
                            fonts(nm) = Array(1, sld.SlideIndex, SKW_Excerpt(rn.Text, 50))
                        End If
                    End If
                Next
            Next
        Next
    Next
    For Each k In fonts.Keys
        e = fonts(k)
        If allowed.Count > 0 And Not SKW_InList(allowed, CStr(k)) Then
            nBad = nBad + 1
            rows.Add SKW_Row("Not allowed", k, e(0), e(1), e(2))
        Else
            rows.Add SKW_Row(IIf(allowed.Count > 0, "Allowed", "Used"), k, e(0), e(1), e(2))
        End If
    Next
    SKW_Status ""
    SKW_ReportFor "Deck Font Audit", fonts.Count & " font(s) used, " & nBad & " not on the team list" & IIf(allowed.Count = 0, " (no Fonts list, so none are flagged)", "") & _
        ". Theme fonts are shown by their real names. Changes nothing.", Array("Rule", "Font", "Text runs", "First slide", "Text"), rows, DeckName(p)
    SKP_DeckFontAudit = fonts.Count & " font(s) used, " & nBad & " not allowed."
End Function

Public Function SKP_DeckSmallText() As String
    Dim p As Object, msg As String, rows As New Collection, sld As Object, sh As Object, tr As Object, rn As Object, minSize As Single, n As Long
    msg = StartDeck("Deck small text", p)
    If Len(msg) > 0 Then SKP_DeckSmallText = msg: Exit Function
    minSize = Val(SKW_Setting("DeckMinFontSize", "10"))
    If minSize <= 0 Then minSize = 10
    For Each sld In p.Slides
        For Each sh In SKP_Shapes(sld)
            For Each tr In SKP_TextRanges(sh)
                For Each rn In tr.Runs
                    If Len(Trim$(rn.Text)) > 0 And rn.Font.Size > 0 And rn.Font.Size < minSize Then
                        n = n + 1
                        If n <= 500 Then rows.Add SKW_Row(sld.SlideIndex, sh.Name, rn.Font.Size & " pt", rn.Text)
                    End If
                Next
            Next
        Next
    Next
    SKW_Status ""
    SKW_ReportFor "Deck Small Text", n & " piece(s) of text under " & minSize & " pt (setting DeckMinFontSize). Text that PowerPoint shrinks to fit is measured at its set size. Changes nothing.", _
        Array("Slide", "Shape", "Size", "Text"), rows, DeckName(p)
    SKP_DeckSmallText = n & " piece(s) of text under " & minSize & " pt."
End Function

Public Function SKP_DeckAltTextAudit() As String
    Dim p As Object, msg As String, rows As New Collection, sld As Object, sh As Object, n As Long, nNo As Long, kind As String, alt As String, dec As Boolean
    msg = StartDeck("Deck alt text", p)
    If Len(msg) > 0 Then SKP_DeckAltTextAudit = msg: Exit Function
    For Each sld In p.Slides
        For Each sh In sld.Shapes
            kind = ""
            Select Case sh.Type
                Case 13, 11: kind = "Picture"
                Case 3: kind = "Chart"
                Case 24: kind = "SmartArt"
                Case 6: kind = "Group"
                Case 16: kind = "Media"
                Case 14
                    On Error Resume Next
                    If sh.HasChart Then kind = "Chart"
                    If sh.PlaceholderFormat.ContainedType = 13 Then kind = "Picture"
                    On Error GoTo 0
            End Select
            If Len(kind) > 0 Then
                n = n + 1
                alt = "": dec = False
                On Error Resume Next
                alt = Trim$(sh.AlternativeText)
                dec = sh.Decorative
                On Error GoTo 0
                If Len(alt) = 0 And Not dec Then nNo = nNo + 1: rows.Add SKW_Row(sld.SlideIndex, kind, sh.Name, "No alt text")
            End If
        Next
    Next
    SKW_Status ""
    SKW_ReportFor "Deck Alt Text Audit", n & " picture(s), chart(s) and other graphics; " & nNo & " without alt text (ones marked Decorative are fine). Changes nothing.", _
        Array("Slide", "Kind", "Shape", "Finding"), rows, DeckName(p)
    SKP_DeckAltTextAudit = n & " graphic(s), " & nNo & " without alt text."
End Function

Public Function SKP_DeckOffSlideObjects() As String
    Dim p As Object, msg As String, rows As New Collection, sld As Object, sh As Object, w As Single, h As Single, n As Long, how As String
    msg = StartDeck("Deck off-slide objects", p)
    If Len(msg) > 0 Then SKP_DeckOffSlideObjects = msg: Exit Function
    w = p.PageSetup.SlideWidth: h = p.PageSetup.SlideHeight
    For Each sld In p.Slides
        For Each sh In sld.Shapes
            how = ""
            If sh.Left >= w Or sh.Top >= h Or sh.Left + sh.Width <= 0 Or sh.Top + sh.Height <= 0 Then
                how = "Wholly off the slide"
            ElseIf sh.Left < -1 Or sh.Top < -1 Or sh.Left + sh.Width > w + 1 Or sh.Top + sh.Height > h + 1 Then
                how = "Partly off the slide"
            End If
            If Len(how) > 0 Then
                n = n + 1
                rows.Add SKW_Row(sld.SlideIndex, sh.Name, how, Format$(sh.Left / 72, "0.00") & ", " & Format$(sh.Top / 72, "0.00") & " in")
            End If
        Next
    Next
    SKW_Status ""
    SKW_ReportFor "Deck Off-Slide Objects", n & " shape(s) partly or wholly off the slide (" & Format$(w / 72, "0.00") & " x " & Format$(h / 72, "0.00") & " in). Changes nothing.", _
        Array("Slide", "Shape", "Finding", "Position (left, top)"), rows, DeckName(p)
    SKP_DeckOffSlideObjects = n & " shape(s) off the slide."
End Function

Public Function SKP_DeckTitlesCheck() As String
    Dim p As Object, msg As String, rows As New Collection, sld As Object, t As String, seen As Object, nNo As Long, nDup As Long
    msg = StartDeck("Deck titles", p)
    If Len(msg) > 0 Then SKP_DeckTitlesCheck = msg: Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For Each sld In p.Slides
        t = SlideTitle(sld)
        If Len(t) = 0 Then
            nNo = nNo + 1
            rows.Add SKW_Row(sld.SlideIndex, "No title", "Give the slide a title (it can be hidden off the slide or behind a shape)", "")
        ElseIf seen.Exists(LCase$(t)) Then
            nDup = nDup + 1
            rows.Add SKW_Row(sld.SlideIndex, "Same title", "Same title as slide " & seen(LCase$(t)), t)
        Else
            seen(LCase$(t)) = sld.SlideIndex
        End If
    Next
    SKW_Status ""
    SKW_ReportFor "Deck Titles Check", p.Slides.Count & " slide(s): " & nNo & " without a title, " & nDup & " sharing a title. Screen readers and the outline use slide titles. Changes nothing.", _
        Array("Slide", "Rule", "Finding", "Title"), rows, DeckName(p)
    SKP_DeckTitlesCheck = nNo & " slide(s) without a title, " & nDup & " sharing one."
End Function

Public Function SKP_DeckEmptyPlaceholders() As String
    Dim p As Object, msg As String, rows As New Collection, sld As Object, sh As Object, n As Long, empty1 As Boolean
    msg = StartDeck("Deck empty placeholders", p)
    If Len(msg) > 0 Then SKP_DeckEmptyPlaceholders = msg: Exit Function
    For Each sld In p.Slides
        For Each sh In sld.Shapes
            If sh.Type = 14 Then
                empty1 = False
                On Error Resume Next
                If sh.HasTextFrame Then empty1 = Not sh.TextFrame.HasText
                If sh.PlaceholderFormat.ContainedType <> 1 And sh.PlaceholderFormat.ContainedType <> 0 Then empty1 = False   ' holds a picture, chart or table
                On Error GoTo 0
                If empty1 Then n = n + 1: rows.Add SKW_Row(sld.SlideIndex, sh.Name, "Empty placeholder (shows ""Click to add"" while editing)")
            End If
        Next
    Next
    SKW_Status ""
    SKW_ReportFor "Deck Empty Placeholders", n & " empty placeholder(s). Delete them or fill them. Changes nothing.", Array("Slide", "Shape", "Finding"), rows, DeckName(p)
    SKP_DeckEmptyPlaceholders = n & " empty placeholder(s)."
End Function

Public Function SKP_DeckDeliveryCopy(ByVal target As String) As String
    Dim p As Object, c As Object, msg As String, i As Long, sld As Object, sh As Object, nNotes As Long, nCom As Long, nHid As Long, k As Long, nK As Long
    Set SKW_LastReport = Nothing
    Set p = SKP_Deck()
    If p Is Nothing Then SKP_DeckDeliveryCopy = SKP_Why(): Exit Function
    If Len(p.Path) = 0 Then SKP_DeckDeliveryCopy = "Save the deck first, so the copy has a folder to go in.": Exit Function
    If LCase$(Left$(p.FullName, 4)) = "http" Then SKP_DeckDeliveryCopy = "The deck is opened from the web. Open it from the synced folder (or save a local copy) first.": Exit Function
    If Len(target) = 0 Then target = p.Path & "\" & Left$(p.Name, InStrRev(p.Name & ".", ".") - 1) & " - delivery.pptx"
    On Error GoTo Fail
    SKW_Status "Making the delivery copy..."
    p.SaveCopyAs target
    Set c = p.Application.Presentations.Open(FileName:=target, ReadOnly:=0, Untitled:=0, WithWindow:=0)
    For i = c.Slides.Count To 1 Step -1
        Set sld = c.Slides(i)
        If sld.SlideShowTransition.Hidden Then sld.Delete: nHid = nHid + 1
    Next
    For Each sld In c.Slides
        For Each sh In sld.NotesPage.Shapes
            If sh.Type = 14 Then
                If sh.PlaceholderFormat.Type = 2 Then
                    If sh.TextFrame.HasText Then sh.TextFrame.TextRange.Text = "": nNotes = nNotes + 1
                End If
            End If
        Next
        On Error Resume Next
        nK = 0: nK = sld.Comments.Count
        For k = nK To 1 Step -1
            sld.Comments(k).Delete
            nCom = nCom + 1
        Next
        On Error GoTo Fail
    Next
    On Error Resume Next
    c.RemoveDocumentInformation 99    ' ppRDIAll
    On Error GoTo Fail
    c.Save
    c.Close
    SKW_Status ""
    SKP_DeckDeliveryCopy = "Saved a delivery copy: " & target & vbCr & "Taken out: " & nHid & " hidden slide(s), notes on " & nNotes & " slide(s), " & nCom & _
        " comment(s), and the file's personal details. The open deck is unchanged."
    Exit Function
Fail:
    SKP_DeckDeliveryCopy = "The delivery copy could not be made: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close
End Function
