Attribute VB_Name = "SKW_Quick"
' =====================================================================
' Stationkeeper Tools: quick tools (1.10.0).
' For working in a document before the standards are entered in the
' Stationkeeper app. Each tool reads a few "quick values" saved on this
' PC (Quick Values opens the form). They are not the Global Standards
' and are never written to them. Every report says so.
'   QuickValues                  the form: styles, cell margins, numbering rule
'   QuickTableFormatHere / All   header row style, body row style, cell margins
'   QuickCaptionFormatHere / All caption styles, and the action caption style
'   CaptionNumberingCheck        caption field codes against one rule; fix in a copy
'   FixCaptionNumberHere         the caption at the cursor
' =====================================================================
Option Explicit

Public SKW_QuickTest As Object   ' tests: values held here in place of the registry

Private Const SEC As String = "Quick"

' ---------------------------------------------------------------------
' The quick values
' ---------------------------------------------------------------------
Public Function SKW_QuickKeys() As Variant
    SKW_QuickKeys = Array("head", "body", "keep", "tstyle", "mtop", "mbottom", "mleft", "mright", "headrow", _
        "captable", "capfigure", "captablechar", "capfigurechar", "capaction", "abbrev", "captab", _
        "nummode", "numhead", "numsep", "numlevel", "tlabel", "flabel")
End Function

Public Function SKW_QuickDefault(ByVal key As String) As String
    Select Case LCase$(key)
        Case "head": SKW_QuickDefault = "AA Table Head L1"
        Case "body": SKW_QuickDefault = "AA Table Body"
        Case "keep": SKW_QuickDefault = "AA Table Head L2; AA Table Bullet L1; AA Table Bullet L2; AA Table Bullet L3; AA Table Notes; AA Table Notes Number List"
        Case "headrow": SKW_QuickDefault = "auto"
        Case "captable": SKW_QuickDefault = "AA Caption-Table"
        Case "capfigure": SKW_QuickDefault = "AA Caption-Figure"
        Case "capaction": SKW_QuickDefault = "AA Caption-Action"
        Case "abbrev": SKW_QuickDefault = "U.S.; vs.; No.; Inc.; Co.; Corp.; Ltd.; e.g.; i.e.; etc.; Fig.; Rev.; Vol.; Sec.; Para.; Ref.; approx.; St.; Dr.; Mr.; Ms."
        Case "captab": SKW_QuickDefault = "1"
        Case "nummode": SKW_QuickDefault = "heading"
        Case "numhead": SKW_QuickDefault = "STYLEREF 1 \s"
        Case "numsep": SKW_QuickDefault = "-"
        Case "numlevel": SKW_QuickDefault = "1"
        Case "tlabel": SKW_QuickDefault = "Table"
        Case "flabel": SKW_QuickDefault = "Figure"
    End Select
End Function

Public Function SKW_QuickGet(ByVal key As String) As String
    Dim v As String
    key = LCase$(key)
    If Not SKW_QuickTest Is Nothing Then
        If SKW_QuickTest.Exists(key) Then SKW_QuickGet = SKW_QuickTest(key) Else SKW_QuickGet = SKW_QuickDefault(key)
        Exit Function
    End If
    v = GetSetting(SKW_APP, SEC, key, Chr$(1))
    If v = Chr$(1) Then SKW_QuickGet = SKW_QuickDefault(key) Else SKW_QuickGet = v
End Function

Public Sub SKW_QuickSet(ByVal key As String, ByVal v As String)
    key = LCase$(key)
    If Not SKW_QuickTest Is Nothing Then
        SKW_QuickTest(key) = v
    Else
        SaveSetting SKW_APP, SEC, key, v
    End If
End Sub

' The values the Global Standards workbook can give today: the header and body tag styles and the
' first table type's cell margins. Empty when the workbook has no Table looks sheet.
Public Function SKW_QuickFromStandards() As Object
    Dim res As Object, types As Collection, roles As Object, m As Variant, i As Long, k As Variant
    Set res = CreateObject("Scripting.Dictionary")
    Set SKW_QuickFromStandards = res
    On Error GoTo Done
    If Not SKW_AppStandards() Then Exit Function
    Set types = SKW_TableTypes()
    If types.Count = 0 Then Exit Function
    Set roles = SKW_TableRoles(types(1)("type"))
    If Len(roles("header")("paragraph style")) > 0 Then res("head") = roles("header")("paragraph style")
    If Len(roles("long text")("paragraph style")) > 0 Then res("body") = roles("long text")("paragraph style")
    m = Split(types(1)("margins pt") & ";;;", ";")
    k = Array("mtop", "mbottom", "mleft", "mright")
    For i = 0 To 3
        If Len(Trim$(m(i))) > 0 Then res(k(i)) = Replace(Format$(Val(m(i)) / 72, "0.###"), ",", ".")
    Next
Done:
End Function

' Opens the form. True when the values were saved.
Public Function SKW_QuickAsk() As Boolean
    frmSKWQuick.Show 1
    SKW_QuickAsk = frmSKWQuick.Saved
    Unload frmSKWQuick
End Function

Private Function Ready() As Boolean
    If SKW_IsQuiet() Then Ready = True: Exit Function
    If SKW_QuickGet("saved") = "1" Then Ready = True: Exit Function
    Ready = SKW_QuickAsk()
End Function

' Opens the quick values: the styles, cell margins and caption numbering rule the quick tools use. Saved on this PC. They are not the Global Standards.
Public Sub QuickValues()
    If SKW_IsQuiet() Then Exit Sub
    If SKW_QuickAsk() Then SKW_Say "Quick values saved on this PC. The quick tools use them from the next run."
End Sub

Private Function Inches(ByVal key As String) As Single
    Dim v As String
    v = Trim$(SKW_QuickGet(key))
    If Len(v) = 0 Then Inches = -1 Else Inches = Val(Replace(v, ",", ".")) * 72
End Function

Private Function MarginsText() As String
    Dim k As Variant, s As String
    For Each k In Array("mtop", "mbottom", "mleft", "mright")
        s = s & IIf(Len(s) > 0, ", ", "") & Mid$(k, 2) & " " & IIf(Len(Trim$(SKW_QuickGet(CStr(k)))) > 0, Trim$(SKW_QuickGet(CStr(k))) & " in", "as it is")
    Next
    MarginsText = s
End Function

Private Function QuickLine() As String
    QuickLine = "Quick values saved on this PC, not the Global Standards. "
End Function

Private Function StyleIn(ByVal d As Object, ByVal nm As String) As Object
    If Len(Trim$(nm)) = 0 Then Exit Function
    Set StyleIn = SKW_GetStyle(d, Trim$(nm))
End Function

Private Function Same(ByVal a As String, ByVal b As String) As Boolean
    Same = (LCase$(Trim$(a)) = LCase$(Trim$(b)))
End Function

' ---------------------------------------------------------------------
' A copy of the saved document, opened, for a whole-document fix.
' ---------------------------------------------------------------------
Private Function CopyOpen(ByVal suffix As String, ByRef target As String, ByRef c As Object) As String
    Dim d As Object, fso As Object, n As Long, base As String, ext As String
    If Documents.Count = 0 Then CopyOpen = "Open a document first.": Exit Function
    Set d = ActiveDocument
    If Len(d.Path) = 0 Then CopyOpen = "Save the document first. The copy is made from the saved file.": Exit Function
    If Not d.Saved Then CopyOpen = "Save the document first. The copy is made from the saved file, so unsaved changes would be left out.": Exit Function
    If Len(target) = 0 Then
        base = d.Path & "\" & SKW_Stem(d.Name) & " (" & suffix & ")"
        ext = Mid$(d.Name, InStrRev(d.Name, "."))
        target = base & ext
        Do While Len(Dir$(target)) > 0
            n = n + 1
            target = base & " " & (n + 1) & ext
        Loop
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile d.FullName, target, True
    SKW_Status "Opening the copy..."
    Set c = Documents.Open(FileName:=target, AddToRecentFiles:=False, Visible:=True)
    c.TrackRevisions = False
End Function

Private Function LeafName(ByVal p As String) As String
    LeafName = Mid$(p, InStrRev(p, "\") + 1)
End Function

' =====================================================================
' Quick Table Format
' =====================================================================

' The header row: row 2 when row 1 holds a caption number field (or always row 2, when set), else row 1.
Private Function HeaderRowOf(ByVal d As Object, ByVal t As Object) As Long
    Dim f As Object, r As Object, n As Long
    HeaderRowOf = 1
    On Error Resume Next
    n = t.Rows.Count
    If n < 2 Then Exit Function
    If SKW_QuickGet("headrow") = "2" Then HeaderRowOf = 2: Exit Function
    Set r = t.Cell(1, 1).Range
    If r Is Nothing Then Exit Function
    For Each f In r.Fields
        If f.Type = wdFieldSequence Then HeaderRowOf = 2: Exit Function
    Next
End Function

' Reads one table against the quick values, and sets it when doIt. Returns the findings ("" when it is right).
Private Function QuickTable(ByVal d As Object, ByVal t As Object, ByVal doIt As Boolean, ByRef hdr As Long, ByRef look As String) As String
    Dim c As Object, p As Object, ri As Long, want As String, nm As String, keep As Object, v As Variant, s As String
    Dim stHead As Object, stBody As Object, stTable As Object, nHead As Long, nBody As Long, st As Object, cur As String
    Dim pts As Single, k As Variant, i As Long, mar As Boolean, descr As String, np As Long
    look = ""
    On Error Resume Next
    descr = LCase$(t.Descr)
    On Error GoTo 0
    If InStr(descr, "skip") > 0 Then look = "Marked skip in its Alt Text": hdr = 0: Exit Function
    hdr = HeaderRowOf(d, t)
    Set stHead = StyleIn(d, SKW_QuickGet("head"))
    Set stBody = StyleIn(d, SKW_QuickGet("body"))
    Set stTable = StyleIn(d, SKW_QuickGet("tstyle"))
    If stHead Is Nothing And Len(Trim$(SKW_QuickGet("head"))) > 0 Then look = "Style " & Trim$(SKW_QuickGet("head")) & " is not in this document"
    If stBody Is Nothing And Len(Trim$(SKW_QuickGet("body"))) > 0 Then look = look & IIf(Len(look) > 0, "; ", "") & "Style " & Trim$(SKW_QuickGet("body")) & " is not in this document"
    If stTable Is Nothing And Len(Trim$(SKW_QuickGet("tstyle"))) > 0 Then look = look & IIf(Len(look) > 0, "; ", "") & "Table style " & Trim$(SKW_QuickGet("tstyle")) & " is not in this document"
    Set keep = CreateObject("Scripting.Dictionary")
    keep.CompareMode = 1
    For Each v In Split(SKW_QuickGet("keep"), ";")
        If Len(Trim$(v)) > 0 Then keep(Trim$(v)) = True
    Next
    ' The table style first: a table style can change what the cells show.
    If Not stTable Is Nothing Then
        On Error Resume Next
        cur = ""
        cur = t.Style
        On Error GoTo 0
        If Not Same(cur, stTable.NameLocal) Then
            s = "Table style is " & IIf(Len(cur) > 0, cur, "(none)")
            If doIt Then t.Range.Style = stTable
        End If
    End If
    ' Cell margins. Word keeps left and right margins in steps of its own (0.08 in is held as 5.4 points), so a
    ' difference under 0.4 point is no difference.
    k = Array("mtop", "mbottom", "mleft", "mright")
    For i = 0 To 3
        pts = Inches(CStr(k(i)))
        If pts >= 0 Then
            On Error Resume Next
            Select Case i
                Case 0: If Abs(t.TopPadding - pts) > 0.4 Then mar = True: If doIt Then t.TopPadding = pts
                Case 1: If Abs(t.BottomPadding - pts) > 0.4 Then mar = True: If doIt Then t.BottomPadding = pts
                Case 2: If Abs(t.LeftPadding - pts) > 0.4 Then mar = True: If doIt Then t.LeftPadding = pts
                Case 3: If Abs(t.RightPadding - pts) > 0.4 Then mar = True: If doIt Then t.RightPadding = pts
            End Select
            On Error GoTo 0
        End If
    Next
    If mar Then s = s & IIf(Len(s) > 0, "; ", "") & "Cell margins differ"
    ' Header row and body rows.
    For Each c In t.Range.Cells
        ri = c.RowIndex
        If ri >= hdr Then
            If ri = hdr Then Set st = stHead Else Set st = stBody
            If Not st Is Nothing Then
                want = st.NameLocal
                np = c.Range.Paragraphs.Count
                If np = 1 Then
                    nm = SKW_ParaStyleName(c.Range)
                    If Not Same(nm, want) Then
                        If ri = hdr Or Not keep.Exists(nm) Then
                            If ri = hdr Then nHead = nHead + 1 Else nBody = nBody + 1
                            If doIt Then c.Range.Style = st
                        End If
                    End If
                Else
                    For Each p In c.Range.Paragraphs
                        nm = SKW_ParaStyleName(p.Range)
                        If Not Same(nm, want) Then
                            If ri = hdr Or Not keep.Exists(nm) Then
                                If ri = hdr Then nHead = nHead + 1 Else nBody = nBody + 1
                                If doIt Then p.Style = st
                            End If
                        End If
                    Next
                End If
            End If
        End If
    Next
    If nHead > 0 Then s = s & IIf(Len(s) > 0, "; ", "") & nHead & " header paragraph(s) not in " & stHead.NameLocal
    If nBody > 0 Then s = s & IIf(Len(s) > 0, "; ", "") & nBody & " body paragraph(s) not in " & stBody.NameLocal
    QuickTable = s
End Function

Private Function TableIntro() As String
    TableIntro = QuickLine() & "Header row: " & IIf(Len(Trim$(SKW_QuickGet("head"))) > 0, Trim$(SKW_QuickGet("head")), "as it is") & _
        ". Body rows: " & IIf(Len(Trim$(SKW_QuickGet("body"))) > 0, Trim$(SKW_QuickGet("body")), "as they are") & _
        ". Cell margins: " & MarginsText() & ". Table style: " & IIf(Len(Trim$(SKW_QuickGet("tstyle"))) > 0, Trim$(SKW_QuickGet("tstyle")), "as it is") & _
        ". The header row is " & IIf(SKW_QuickGet("headrow") = "2", "always row 2", "row 2 when row 1 holds a caption number, else row 1") & _
        ". A caption row is left alone. Styles left alone in body rows: " & SKW_QuickGet("keep") & "."
End Function

' Sets the table at the cursor from the quick values: the header row's style, the body rows' style and the cell margins. One Ctrl+Z undoes it.
Public Sub QuickTableFormatHere()
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    If Not Selection.Information(wdWithInTable) Then SKW_Say "Put the cursor in a table first.": Exit Sub
    If Not Ready() Then Exit Sub
    SKW_Show SKW_QuickTableHere()
End Sub

Public Function SKW_QuickTableHere() As String
    Dim d As Object, msg As String, t As Object, s As String, hdr As Long, look As String
    Set SKW_LastReport = Nothing
    If Documents.Count > 0 Then
        If Not Selection.Information(wdWithInTable) Then SKW_QuickTableHere = "Put the cursor in a table first.": Exit Function
    End If
    msg = SKW_Begin("Quick table format", True)
    If Len(msg) > 0 Then SKW_QuickTableHere = msg: Exit Function
    Set d = ActiveDocument
    Set t = Selection.Tables(1)
    s = QuickTable(d, t, True, hdr, look)
    SKW_Finish
    If hdr = 0 Then SKW_QuickTableHere = look & ". Nothing was changed.": Exit Function
    SKW_QuickTableHere = IIf(Len(s) > 0, "Table set from the quick values (header row " & hdr & "). It had: " & s & ".", _
        "The table already matches the quick values (header row " & hdr & "). Nothing was changed.") & _
        IIf(Len(look) > 0, vbCr & "Look at: " & look & ".", "") & IIf(Len(s) > 0, vbCr & "One Ctrl+Z undoes it.", "")
End Function

' Report of every table against the quick values: the header row's style, the body rows' style and the cell margins. Changes nothing, then offers to set every table in a copy.
Public Sub QuickTableFormatAll()
    Dim s As String
    If SKW_IsQuiet() Then Exit Sub
    If Not Ready() Then Exit Sub
    SKW_Fixable = 0
    s = SKW_QuickTableCheck()
    If SKW_OfferFix(s, "The fix sets every table's header row, body rows and cell margins in a copy of the document. Your document is not changed.") Then
        SKW_Show SKW_QuickTableFixCopy("")
    End If
End Sub

Public Function SKW_QuickTableCheck() As String
    Dim d As Object, msg As String, rows As New Collection, t As Object, i As Long, s As String, hdr As Long, look As String
    Dim nBad As Long, nSkip As Long, nLook As Long
    Set SKW_LastReport = Nothing
    SKW_Fixable = 0
    msg = SKW_Begin("Quick table format", False)
    If Len(msg) > 0 Then SKW_QuickTableCheck = msg: Exit Function
    Set d = ActiveDocument
    SKW_Steps d.Tables.Count, "Table"
    For Each t In d.Tables
        SKW_Step
        i = i + 1
        s = QuickTable(d, t, False, hdr, look)
        If hdr = 0 Then
            nSkip = nSkip + 1
            rows.Add SKW_Row("Table " & i, SKW_Page(t.Range), "", "Skipped", look)
        ElseIf Len(s) > 0 Or Len(look) > 0 Then
            If Len(s) > 0 Then nBad = nBad + 1
            If Len(look) > 0 Then nLook = nLook + 1
            rows.Add SKW_Row("Table " & i, SKW_Page(t.Range), hdr, IIf(Len(s) > 0, s, "Matches"), look)
        End If
    Next
    SKW_Fixable = nBad
    SKW_Finish
    msg = d.Tables.Count & " table(s) read, " & nBad & " to set, " & nSkip & " skipped."
    SKW_Report "Quick Table Format", msg & " " & TableIntro() & " Changes nothing.", Array("Table", "Page", "Header row", "Finding", "Look at"), rows
    SKW_QuickTableCheck = msg
End Function

Public Function SKW_QuickTableFixCopy(ByVal target As String) As String
    Dim c As Object, t As Object, s As String, hdr As Long, look As String, n As Long, nSkip As Long, msg As String, ed As String
    Set SKW_LastReport = Nothing
    On Error GoTo Fail
    msg = CopyOpen("tables formatted", target, c)
    If Len(msg) > 0 Then SKW_QuickTableFixCopy = msg: Exit Function
    SKW_ProgressTitle "Quick table format: setting a copy"
    Application.ScreenUpdating = False
    SKW_Steps c.Tables.Count, "Table"
    For Each t In c.Tables
        SKW_Step
        s = QuickTable(c, t, True, hdr, look)
        If hdr = 0 Then nSkip = nSkip + 1 Else If Len(s) > 0 Then n = n + 1
    Next
    SKW_Status "Saving the copy..."
    c.Save
    Application.ScreenUpdating = True
    Application.StatusBar = False
    SKW_ProgressDone
    SKW_QuickTableFixCopy = n & " table(s) set in a copy, " & nSkip & " skipped: " & LeafName(target) & "." & vbCr & "Your document is not changed. The copy is open."
    Exit Function
Fail:
    ed = Err.Description
    Application.ScreenUpdating = True
    Application.StatusBar = False
    SKW_ProgressDone
    SKW_QuickTableFixCopy = "Stopped: " & ed & ". Your document is not changed."
End Function

' =====================================================================
' Captions: shared reading
' =====================================================================
Private Function Norm(ByVal s As String) As String
    s = Replace(Replace(s, vbTab, " "), Chr$(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Norm = Trim$(s)
End Function

' A SEQ field's code as its parts: id, s (restart level), r (restart value), c, h, fmt.
Private Function SeqParts(ByVal code As String) As Object
    Dim q As Object, a As Variant, i As Long, tok As String
    Set q = CreateObject("Scripting.Dictionary")
    q("id") = "": q("s") = "": q("r") = "": q("c") = False: q("h") = False: q("fmt") = ""
    a = Split(Norm(code), " ")
    If UBound(a) >= 1 Then q("id") = Replace(CStr(a(1)), """", "")
    For i = 2 To UBound(a)
        tok = LCase$(CStr(a(i)))
        If tok = "\c" Then
            q("c") = True
        ElseIf tok = "\h" Then
            q("h") = True
        ElseIf tok = "\s" Or tok = "\r" Then
            If i < UBound(a) Then q(Mid$(tok, 2)) = CStr(a(i + 1)): i = i + 1
        ElseIf Left$(tok, 2) = "\s" Or Left$(tok, 2) = "\r" Then
            q(Mid$(tok, 2, 1)) = Mid$(tok, 3)
        ElseIf tok = "\*" Then
            If i < UBound(a) Then
                If LCase$(CStr(a(i + 1))) <> "mergeformat" And Len(q("fmt")) = 0 Then q("fmt") = CStr(a(i + 1))
                i = i + 1
            End If
        End If
    Next
    Set SeqParts = q
End Function

Private Function SeqCode(ByVal id As String, ByVal lvl As Long) As String
    SeqCode = " SEQ " & id & " \* ARABIC" & IIf(lvl > 0, " \s " & lvl, "") & " "
End Function

' A STYLEREF code made plain for comparing: lower case, no quotes, single spaces, no MERGEFORMAT.
Private Function NormRef(ByVal code As String) As String
    Dim s As String
    s = LCase$(Norm(Replace(code, """", "")))
    s = Norm(Replace(s, "\* mergeformat", ""))
    NormRef = s
End Function

' The heading level a STYLEREF names by number (STYLEREF 1 \s), or 0 when it names a style.
Private Function RefLevel(ByVal code As String) As Long
    Dim a As Variant
    a = Split(NormRef(code), " ")
    If UBound(a) >= 1 Then
        If IsNumeric(a(1)) Then RefLevel = CLng(a(1))
    End If
End Function

Private Function FStart(ByVal f As Object) As Long
    FStart = f.Code.Start - 1
End Function

Private Function FEnd(ByVal f As Object) As Long
    On Error Resume Next
    FEnd = f.Code.End + 1
    FEnd = f.Result.End + 1
End Function

' The caption number's fields in a paragraph: the first SEQ that is not hidden, and the STYLEREF before it.
Private Sub NumberFields(ByVal p As Object, ByRef sq As Object, ByRef sr As Object)
    Dim x As Object
    Set sq = Nothing
    Set sr = Nothing
    For Each x In p.Fields
        If sq Is Nothing Then
            If x.Type = wdFieldSequence Then
                If Not SeqParts(x.Code.Text)("h") Then Set sq = x
            ElseIf x.Type = wdFieldStyleRef Then
                Set sr = x
            End If
        End If
    Next
End Sub

Private Function ParaText(ByVal p As Object) As String
    Dim s As String
    s = p.Text
    Do While Len(s) > 0
        Select Case Right$(s, 1)
            Case vbCr, Chr$(7), vbLf: s = Left$(s, Len(s) - 1)
            Case Else: Exit Do
        End Select
    Loop
    ParaText = s
End Function

' The caption's label word: the text before the first space.
Private Function CapLabel(ByVal p As Object) As String
    Dim s As String, i As Long
    s = Replace(Replace(LTrim$(ParaText(p)), Chr$(160), " "), vbTab, " ")
    i = InStr(s, " ")
    If i > 1 Then CapLabel = Left$(s, i - 1) Else CapLabel = s
End Function

' Every caption paragraph of the main text: a paragraph holding a SEQ field that is not hidden. Last first, so a fix never moves one still to come.
Private Function CaptionParas(ByVal d As Object) As Collection
    Dim f As Object, seen As Object, c As New Collection, p As Object, back As New Collection, i As Long
    Set seen = CreateObject("Scripting.Dictionary")
    For Each f In d.Fields
        If f.Type = wdFieldSequence Then
            If Not SeqParts(f.Code.Text)("h") Then
                Set p = f.Code.Paragraphs(1).Range
                ' A range that comes from a field's code reads field codes as text; captions are read as they show.
                p.TextRetrievalMode.IncludeFieldCodes = False
                If Not seen.Exists(p.Start) Then
                    seen(p.Start) = True
                    c.Add p
                End If
            End If
        End If
    Next
    For i = c.Count To 1 Step -1
        back.Add c(i)
    Next
    Set CaptionParas = back
End Function

' "table", "figure" or "" for a caption paragraph: by its label word, then by what its SEQ counts.
Private Function CapKind(ByVal p As Object, ByVal sq As Object) As String
    Dim lab As String, id As String
    lab = CapLabel(p)
    id = SeqParts(sq.Code.Text)("id")
    If Same(lab, SKW_QuickGet("tlabel")) Then
        CapKind = "table"
    ElseIf Same(lab, SKW_QuickGet("flabel")) Then
        CapKind = "figure"
    ElseIf Same(id, SKW_QuickGet("tlabel")) Then
        CapKind = "table"
    ElseIf Same(id, SKW_QuickGet("flabel")) Then
        CapKind = "figure"
    End If
End Function

' =====================================================================
' Quick Caption Format
' =====================================================================
Private Function CharStyleAt(ByVal d As Object, ByVal pos As Long) As String
    On Error Resume Next
    CharStyleAt = d.Range(pos, pos + 1).CharacterStyle
End Function

' True when the period at the end of "before" does not end the caption's title: an abbreviation
' (U.S., vs., a single letter, or one on the list), or the next word starts in lower case.
Private Function IsAbbrev(ByVal before As String, ByVal afterCh As String) As Boolean
    Dim tok As String, i As Long, v As Variant, ok As Boolean
    If Len(afterCh) = 0 Then IsAbbrev = True: Exit Function
    If afterCh Like "[a-z]" Then IsAbbrev = True: Exit Function
    tok = Replace(Replace(before, Chr$(160), " "), vbTab, " ")
    i = InStrRev(tok, " ")
    If i > 0 Then tok = Mid$(tok, i + 1)
    Do While Len(tok) > 0
        If Left$(tok, 1) Like "[A-Za-z0-9]" Then Exit Do
        tok = Mid$(tok, 2)
    Loop
    If Len(tok) < 2 Then Exit Function
    ' U.S. and J. : letters each followed by a period.
    ok = (Len(tok) Mod 2 = 0)
    For i = 1 To Len(tok) Step 2
        If Not (Mid$(tok, i, 1) Like "[A-Za-z]" And Mid$(tok, i + 1, 1) = ".") Then ok = False
    Next
    If ok Then IsAbbrev = True: Exit Function
    For Each v In Split(SKW_QuickGet("abbrev"), ";")
        If Len(Trim$(v)) > 0 Then
            If LCase$(Trim$(v)) = LCase$(tok) Then IsAbbrev = True: Exit Function
        End If
    Next
End Function

' Where the title ends and the action caption starts: the place of the first period that is followed
' by a space and is not an abbreviation. -1 when the caption has no action caption.
Private Function SplitAt(ByVal d As Object, ByVal tr As Object) As Long
    Dim txt As String, i As Long, k As Long, f As Object, a As String
    SplitAt = -1
    If tr.Fields.Count = 0 Then
        txt = tr.Text
        i = 1
        Do
            k = InStr(i, txt, ". ")
            If k = 0 Then Exit Do
            If Not IsAbbrev(Left$(txt, k), Mid$(txt, k + 2, 1)) Then SplitAt = tr.Start + k - 1: Exit Function
            i = k + 1
        Loop
    Else
        Set f = tr.Duplicate
        With f.Find
            .ClearFormatting
            .Text = ". "
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWildcards = False
        End With
        Do While f.Find.Execute
            If f.Start < tr.Start Or f.End > tr.End Then Exit Do
            a = ""
            If f.End < tr.End Then a = d.Range(f.End, f.End + 1).Text
            If Not IsAbbrev(d.Range(tr.Start, f.Start + 1).Text, a) Then SplitAt = f.Start: Exit Function
            f.Collapse wdCollapseEnd
        Loop
    End If
End Function

' What sits between a caption's number and its title, in words.
Private Function SepWords(ByVal s As String) As String
    If Len(s) = 0 Then SepWords = "nothing": Exit Function
    SepWords = """" & Replace(Replace(s, vbTab, "[tab]"), Chr$(160), "[non-breaking space]") & """"
End Function

' A paragraph's layout as text, for comparing with its style: indents and the tab stops a person or a style set.
Private Function LayoutSig(ByVal pf As Object) As String
    Dim t As Object, s As String
    On Error Resume Next
    s = Format$(pf.LeftIndent, "0.0") & "|" & Format$(pf.FirstLineIndent, "0.0") & "|"
    For Each t In pf.TabStops
        If t.CustomTab Then s = s & Format$(t.Position, "0.0") & "/" & t.Alignment & ";"
    Next
    LayoutSig = s
End Function

' Gives a paragraph its style's indents and tab stops. Nothing else on the paragraph is touched.
Private Sub MatchLayout(ByVal p As Object, ByVal st As Object)
    Dim t As Object
    On Error Resume Next
    With p.ParagraphFormat
        .LeftIndent = st.ParagraphFormat.LeftIndent
        .FirstLineIndent = st.ParagraphFormat.FirstLineIndent
        .TabStops.ClearAll
        For Each t In st.ParagraphFormat.TabStops
            If t.CustomTab Then .TabStops.Add t.Position, t.Alignment, t.Leader
        Next
    End With
End Sub

' Reads one caption against the quick values and sets it when doIt. Returns the findings.
Private Function CapOne(ByVal d As Object, ByVal p As Object, ByVal doIt As Boolean, ByRef kind As String, _
                        ByRef title As String, ByRef action As String, ByRef look As String) As String
    Dim sq As Object, sr As Object, pos As Long, pe As Long, sp As Long, s As String, ch As String
    Dim stP As Object, stC As Object, stA As Object, wantP As String, wantC As String, wantA As String, tr As Object, sepTxt As String, r As Object
    kind = "": title = "": action = "": look = ""
    NumberFields p, sq, sr
    If sq Is Nothing Then Exit Function
    ' A number that repeats the one before it (\c) sits in running text: it is not a caption.
    If SeqParts(sq.Code.Text)("c") Then Exit Function
    kind = CapKind(p, sq)
    If Len(kind) = 0 Then Exit Function
    wantP = Trim$(SKW_QuickGet(IIf(kind = "table", "captable", "capfigure")))
    wantC = Trim$(SKW_QuickGet(IIf(kind = "table", "captablechar", "capfigurechar")))
    wantA = Trim$(SKW_QuickGet("capaction"))
    Set stP = StyleIn(d, wantP)
    Set stC = StyleIn(d, wantC)
    Set stA = StyleIn(d, wantA)
    If stP Is Nothing And Len(wantP) > 0 Then look = "Style " & wantP & " is not in this document"
    If stC Is Nothing And Len(wantC) > 0 Then look = look & IIf(Len(look) > 0, "; ", "") & "Style " & wantC & " is not in this document"
    pe = p.End - 1
    pos = FEnd(sq)
    Do While pos < pe
        ch = d.Range(pos, pos + 1).Text
        If Len(ch) = 0 Then Exit Do
        If InStr(" .:-" & vbTab & Chr$(160) & Chr$(30) & ChrW$(8211) & ChrW$(8212), ch) = 0 Then Exit Do
        pos = pos + 1
    Loop
    ' The paragraph style first: the character styles go on top of it.
    If Not stP Is Nothing Then
        If Not Same(SKW_ParaStyleName(p), stP.NameLocal) Then
            s = "Paragraph style is " & SKW_ParaStyleName(p)
            If doIt Then p.Style = stP
        End If
    End If
    If pos >= pe Then look = look & IIf(Len(look) > 0, "; ", "") & "No title after the number": CapOne = s: Exit Function
    ' One tab between the number and the title, so every title starts at the style's tab stop.
    If SKW_QuickGet("captab") <> "0" Then
        sepTxt = d.Range(FEnd(sq), pos).Text
        If sepTxt <> vbTab Then
            s = s & IIf(Len(s) > 0, "; ", "") & "The number is followed by " & SepWords(sepTxt) & ", not one tab"
            If doIt Then
                Set r = d.Range(FEnd(sq), pos)
                If r.End > r.Start Then r.Text = vbTab Else r.InsertBefore vbTab
                pos = FEnd(sq) + 1
                pe = p.End - 1
            End If
        End If
    End If
    ' Tab stops and indents come from the caption style. One set by hand on a caption is how captions drift.
    If Not stP Is Nothing Then
        If Same(SKW_ParaStyleName(p), stP.NameLocal) Then
            If LayoutSig(p.ParagraphFormat) <> LayoutSig(stP.ParagraphFormat) Then
                s = s & IIf(Len(s) > 0, "; ", "") & "Tab stops or indents are set by hand"
                If doIt Then MatchLayout p, stP
            End If
        End If
    End If
    Set tr = d.Range(pos, pe)
    sp = SplitAt(d, tr)
    If sp >= 0 Then
        title = d.Range(pos, sp).Text
        action = Trim$(Mid$(d.Range(sp, pe).Text, 2))
        If stA Is Nothing Then
            If Len(wantA) > 0 Then look = look & IIf(Len(look) > 0, "; ", "") & "Style " & wantA & " is not in this document"
        Else
            ' The period goes with the action caption, so a list built from the title's style shows no period.
            If Not Same(CharStyleAt(d, sp), stA.NameLocal) Or Not Same(CharStyleAt(d, pe - 1), stA.NameLocal) Then
                s = s & IIf(Len(s) > 0, "; ", "") & "The period and action caption are not in " & stA.NameLocal
                If doIt Then d.Range(sp, pe).Style = stA
            End If
            If stC Is Nothing Then
                If Same(CharStyleAt(d, pos), stA.NameLocal) Or Same(CharStyleAt(d, sp - 1), stA.NameLocal) Then
                    s = s & IIf(Len(s) > 0, "; ", "") & "The title carries " & stA.NameLocal
                    If doIt Then d.Range(pos, sp).Style = d.Styles(wdStyleDefaultParagraphFont)
                End If
            End If
        End If
        If Not stC Is Nothing Then
            If Not Same(CharStyleAt(d, pos), stC.NameLocal) Or Not Same(CharStyleAt(d, sp - 1), stC.NameLocal) Then
                s = s & IIf(Len(s) > 0, "; ", "") & "The title is not in " & stC.NameLocal
                If doIt Then d.Range(p.Start, sp).Style = stC
            End If
        End If
    Else
        title = tr.Text
        If Right$(RTrim$(title), 1) = "." Then look = look & IIf(Len(look) > 0, "; ", "") & "Ends in a period with no action caption"
        If Not stC Is Nothing Then
            If Not Same(CharStyleAt(d, pos), stC.NameLocal) Or Not Same(CharStyleAt(d, pe - 1), stC.NameLocal) Then
                s = s & IIf(Len(s) > 0, "; ", "") & "The title is not in " & stC.NameLocal
                If doIt Then d.Range(p.Start, pe).Style = stC
            End If
        ElseIf Not stA Is Nothing Then
            If Same(CharStyleAt(d, pos), stA.NameLocal) Then
                s = s & IIf(Len(s) > 0, "; ", "") & "The title carries " & stA.NameLocal & " and has no action caption"
                If doIt Then d.Range(pos, pe).Style = d.Styles(wdStyleDefaultParagraphFont)
            End If
        End If
    End If
    CapOne = s
End Function

Private Function CaptionIntro() As String
    CaptionIntro = QuickLine() & "Table captions: " & IIf(Len(Trim$(SKW_QuickGet("captable"))) > 0, Trim$(SKW_QuickGet("captable")), "as they are") & _
        ". Figure captions: " & IIf(Len(Trim$(SKW_QuickGet("capfigure"))) > 0, Trim$(SKW_QuickGet("capfigure")), "as they are") & _
        ". Action caption: " & IIf(Len(Trim$(SKW_QuickGet("capaction"))) > 0, Trim$(SKW_QuickGet("capaction")), "as it is") & _
        ". The action caption starts at the first period followed by a space, and the period goes with it. " & _
        "A period after an abbreviation (" & SKW_QuickGet("abbrev") & ") or before a lower-case word does not count. A caption with no period has no action caption. " & _
        IIf(SKW_QuickGet("captab") <> "0", "One tab sits between the number and the title. ", "") & _
        "A caption already in its style takes its tab stops and indents from the style."
End Function

' Sets the caption at the cursor from the quick values: its paragraph style, one tab after the number, the style's tab stops and indents, and the action caption's style from the first period on. One Ctrl+Z undoes it.
Public Sub QuickCaptionFormatHere()
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    If Not Ready() Then Exit Sub
    SKW_Show SKW_QuickCaptionHere()
End Sub

Public Function SKW_QuickCaptionHere() As String
    Dim d As Object, msg As String, p As Object, s As String, kind As String, title As String, action As String, look As String, was As Boolean
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Quick caption format", True)
    If Len(msg) > 0 Then SKW_QuickCaptionHere = msg: Exit Function
    Set d = ActiveDocument
    was = d.ActiveWindow.View.ShowFieldCodes
    d.ActiveWindow.View.ShowFieldCodes = False
    Set p = Selection.Paragraphs(1).Range
    s = CapOne(d, p, True, kind, title, action, look)
    d.ActiveWindow.View.ShowFieldCodes = was
    SKW_Finish
    If Len(kind) = 0 Then
        SKW_QuickCaptionHere = "The paragraph at the cursor is not a table or figure caption: it has no caption number field. Nothing was changed."
        Exit Function
    End If
    SKW_QuickCaptionHere = IIf(Len(s) > 0, "Caption set from the quick values. It had: " & s & ".", "The caption already matches the quick values. Nothing was changed.") & _
        vbCr & "Title: " & title & IIf(Len(action) > 0, vbCr & "Action caption: " & action, vbCr & "No action caption.") & _
        IIf(Len(look) > 0, vbCr & "Look at: " & look & ".", "") & IIf(Len(s) > 0, vbCr & "One Ctrl+Z undoes it.", "")
End Function

' Report of every table and figure caption against the quick values, with each caption shown as title and action caption. Changes nothing, then offers to set every caption in a copy.
Public Sub QuickCaptionFormatAll()
    Dim s As String
    If SKW_IsQuiet() Then Exit Sub
    If Not Ready() Then Exit Sub
    SKW_Fixable = 0
    s = SKW_QuickCaptionCheck()
    If SKW_OfferFix(s, "The fix sets every caption's styles in a copy of the document. Your document is not changed.") Then
        SKW_Show SKW_QuickCaptionFixCopy("")
    End If
End Sub

Public Function SKW_QuickCaptionCheck() As String
    Dim d As Object, msg As String, rows As New Collection, caps As Collection, i As Long, p As Object, s As String
    Dim kind As String, title As String, action As String, look As String, nBad As Long, nT As Long, nF As Long, nAct As Long, nLook As Long, was As Boolean
    Set SKW_LastReport = Nothing
    SKW_Fixable = 0
    msg = SKW_Begin("Quick caption format", False)
    If Len(msg) > 0 Then SKW_QuickCaptionCheck = msg: Exit Function
    Set d = ActiveDocument
    was = d.ActiveWindow.View.ShowFieldCodes
    d.ActiveWindow.View.ShowFieldCodes = False
    Set caps = CaptionParas(d)
    SKW_Steps caps.Count, "Caption"
    ' The list runs last to first; the report reads first to last.
    For i = caps.Count To 1 Step -1
        SKW_Step
        Set p = caps(i)
        s = CapOne(d, p, False, kind, title, action, look)
        If Len(kind) > 0 Then
            If kind = "table" Then nT = nT + 1 Else nF = nF + 1
            If Len(action) > 0 Then nAct = nAct + 1
            If Len(s) > 0 Then nBad = nBad + 1
            If Len(look) > 0 Then nLook = nLook + 1
            rows.Add SKW_Row(IIf(kind = "table", "Table", "Figure"), SKW_Page(p), title, action, IIf(Len(s) > 0, s, "Matches"), look)
        End If
    Next
    d.ActiveWindow.View.ShowFieldCodes = was
    SKW_Fixable = nBad
    SKW_Finish
    msg = nT & " table caption(s) and " & nF & " figure caption(s) read, " & nAct & " with an action caption, " & nBad & " to set, " & nLook & " to look at."
    SKW_Report "Quick Caption Format", msg & " " & CaptionIntro() & " Captions in text boxes, headers and footers are not read. Changes nothing.", _
        Array("Kind", "Page", "Title", "Action caption", "Finding", "Look at"), rows
    SKW_QuickCaptionCheck = msg
End Function

Public Function SKW_QuickCaptionFixCopy(ByVal target As String) As String
    Dim c As Object, caps As Collection, p As Object, s As String, kind As String, title As String, action As String, look As String
    Dim n As Long, msg As String, ed As String
    Set SKW_LastReport = Nothing
    On Error GoTo Fail
    msg = CopyOpen("captions formatted", target, c)
    If Len(msg) > 0 Then SKW_QuickCaptionFixCopy = msg: Exit Function
    SKW_ProgressTitle "Quick caption format: setting a copy"
    c.ActiveWindow.View.ShowFieldCodes = False
    Application.ScreenUpdating = False
    Set caps = CaptionParas(c)
    SKW_Steps caps.Count, "Caption"
    For Each p In caps
        SKW_Step
        s = CapOne(c, p, True, kind, title, action, look)
        If Len(s) > 0 Then n = n + 1
    Next
    SKW_Status "Saving the copy..."
    c.Save
    Application.ScreenUpdating = True
    Application.StatusBar = False
    SKW_ProgressDone
    SKW_QuickCaptionFixCopy = n & " caption(s) set in a copy: " & LeafName(target) & "." & vbCr & "Your document is not changed. The copy is open."
    Exit Function
Fail:
    ed = Err.Description
    Application.ScreenUpdating = True
    Application.StatusBar = False
    SKW_ProgressDone
    SKW_QuickCaptionFixCopy = "Stopped: " & ed & ". Your document is not changed."
End Function

' =====================================================================
' Caption numbering
' =====================================================================
Private Function SepPlain(ByVal s As String) As String
    SepPlain = Replace(Replace(s, Chr$(30), "-"), ChrW$(8209), "-")
End Function

Private Function RuleText() As String
    If LCase$(SKW_QuickGet("nummode")) = "sequential" Then
        RuleText = "sequential all the way through (Table 1, Table 2), with no heading number and no restart"
    Else
        RuleText = "heading number, then sequence (" & Norm(SKW_QuickGet("numhead")) & ", then """ & SKW_QuickGet("numsep") & """, then a count that restarts under heading level " & _
            Val(SKW_QuickGet("numlevel")) & ")"
    End If
End Function

' Reads one caption's number fields against the rule and repairs them when doIt. Returns the findings.
' look: something left alone on purpose. hasHead: the caption has a heading number.
Private Function NumOne(ByVal d As Object, ByVal p As Object, ByVal doIt As Boolean, ByRef look As String, ByRef hasHead As Boolean) As String
    Dim sq As Object, sr As Object, q As Object, seqMode As Boolean, lvl As Long, wantLvl As Long, n As Long
    Dim lab As String, wantId As String, s As String, sep As String, wantSep As String, newCode As String, r As Object, sepBad As Boolean
    look = "": hasHead = False
    NumberFields p, sq, sr
    If sq Is Nothing Then Exit Function
    Set q = SeqParts(sq.Code.Text)
    If q("c") Then look = "Repeats the number before it (\c), so it is not a caption's own number: left alone": Exit Function
    hasHead = Not sr Is Nothing
    seqMode = (LCase$(SKW_QuickGet("nummode")) = "sequential")
    wantLvl = Val(SKW_QuickGet("numlevel"))
    If wantLvl < 1 Or wantLvl > 9 Then wantLvl = 1
    wantSep = SKW_QuickGet("numsep")
    lab = CapLabel(p)
    wantId = q("id")
    If Same(lab, SKW_QuickGet("tlabel")) Then wantId = SKW_QuickGet("tlabel")
    If Same(lab, SKW_QuickGet("flabel")) Then wantId = SKW_QuickGet("flabel")
    If Not Same(q("id"), wantId) Then s = "Counts as " & q("id") & " under the label " & lab
    If Len(q("fmt")) > 0 And UCase$(q("fmt")) <> "ARABIC" Then s = s & IIf(Len(s) > 0, "; ", "") & "Number format is " & q("fmt")
    If seqMode Then
        If Len(q("s")) > 0 Then s = s & IIf(Len(s) > 0, "; ", "") & "Restarts under heading level " & q("s")
        If Len(q("r")) > 0 Then s = s & IIf(Len(s) > 0, "; ", "") & "Restarts at " & q("r")
        If hasHead Then s = s & IIf(Len(s) > 0, "; ", "") & "Has a heading number"
        lvl = 0
    Else
        If sr Is Nothing Then
            s = s & IIf(Len(s) > 0, "; ", "") & "No heading number"
            lvl = wantLvl
        Else
            n = RefLevel(sr.Code.Text)
            If n > 0 Then
                lvl = n
                If n <> wantLvl Then look = "Takes its heading number from level " & n & ", not level " & wantLvl & " (an appendix?): its level is kept"
            Else
                lvl = Val(q("s"))
                If lvl < 1 Then lvl = wantLvl
                If NormRef(sr.Code.Text) <> NormRef(SKW_QuickGet("numhead")) Then look = "Takes its heading number from " & Norm(sr.Code.Text) & ": kept"
            End If
            sep = d.Range(FEnd(sr), FStart(sq)).Text
            If SepPlain(sep) <> SepPlain(wantSep) Then
                sepBad = True
                s = s & IIf(Len(s) > 0, "; ", "") & "Separator is """ & sep & """, not """ & wantSep & """"
            End If
        End If
        If Len(q("s")) = 0 Then
            s = s & IIf(Len(s) > 0, "; ", "") & "Does not restart under the heading"
        ElseIf Val(q("s")) <> lvl Then
            s = s & IIf(Len(s) > 0, "; ", "") & "Restarts under heading level " & q("s") & ", not " & lvl
        End If
        If Len(q("r")) > 0 Then s = s & IIf(Len(s) > 0, "; ", "") & "Restarts at " & q("r")
    End If
    NumOne = s
    If Len(s) = 0 Or Not doIt Then Exit Function
    ' The repair: the smallest change that makes the codes right. The caption's text is not touched.
    newCode = SeqCode(wantId, lvl)
    If Norm(sq.Code.Text) <> Norm(newCode) Then sq.Code.Text = newCode
    If seqMode Then
        If hasHead Then d.Range(FStart(sr), FStart(sq)).Delete
    ElseIf sr Is Nothing Then
        Set r = d.Range(FStart(sq), FStart(sq))
        d.Fields.Add r, wdFieldEmpty, Norm(SKW_QuickGet("numhead")), False
        NumberFields p, sq, sr
        If Not sq Is Nothing And Len(wantSep) > 0 Then
            Set r = d.Range(FStart(sq), FStart(sq))
            r.InsertBefore wantSep
        End If
    ElseIf sepBad Then
        Set r = d.Range(FEnd(sr), FStart(sq))
        If r.End > r.Start Then r.Text = wantSep Else r.InsertBefore wantSep
    End If
    On Error Resume Next
    p.Fields.Update
End Function

' A caption whose number is typed: "Table 3-2" as plain text in a caption style. Returns the typed number, or "".
Private Function TypedNumber(ByVal p As Object, ByRef labLen As Long) As String
    Static re As Object
    Dim m As Object, txt As String
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True
    End If
    re.Pattern = "^((?:" & SKW_QuickGet("tlabel") & "|" & SKW_QuickGet("flabel") & ")[ \u00A0]+)([A-Za-z]?[0-9]+(?:[-.\u2011\u2013][0-9]+)*)"
    txt = ParaText(p)
    If Not re.Test(txt) Then Exit Function
    Set m = re.Execute(txt)(0)
    labLen = Len(m.SubMatches(0))
    TypedNumber = m.SubMatches(1)
End Function

' Paragraphs in the caption styles that hold no caption number field but start with a typed number. Last first.
Private Function TypedCaptions(ByVal d As Object) As Collection
    Dim c As New Collection, back As New Collection, seen As Object, nm As Variant, st As Object, r As Object, p As Object, sq As Object, sr As Object
    Dim names As Object, n As Long, i As Long, starts As Object, keys As Variant, j As Long, tmp As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = 1
    For Each nm In Array(SKW_QuickGet("captable"), SKW_QuickGet("capfigure"), SKW_Setting("CaptionStyle", "Caption"))
        If Len(Trim$(nm)) > 0 Then names(Trim$(nm)) = True
    Next
    Set starts = CreateObject("Scripting.Dictionary")
    For Each nm In names.Keys
        Set st = SKW_GetStyle(d, CStr(nm))
        If Not st Is Nothing Then
            For Each r In SKW_StyleRanges(d.Content, st, 5000)
                For Each p In r.Paragraphs
                    If Not seen.Exists(p.Range.Start) Then
                        seen(p.Range.Start) = True
                        NumberFields p.Range, sq, sr
                        If sq Is Nothing Then
                            If Len(TypedNumber(p.Range, n)) > 0 Then Set starts(p.Range.Start) = p.Range
                        End If
                    End If
                Next
            Next
        End If
    Next
    ' Sorted by place, last first.
    keys = starts.Keys
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(j) > keys(i) Then tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
        Next
    Next
    For i = 0 To UBound(keys)
        back.Add starts(keys(i))
    Next
    Set TypedCaptions = back
End Function

' Puts fields in place of a typed caption number.
Private Function FixTyped(ByVal d As Object, ByVal p As Object) As Boolean
    Dim num As String, labLen As Long, r As Object, lab As String, look As String, hh As Boolean
    num = TypedNumber(p, labLen)
    If Len(num) = 0 Then Exit Function
    If d.Range(p.Start, p.Start + labLen + Len(num)).Fields.Count > 0 Then Exit Function
    lab = CapLabel(p)
    If Same(lab, SKW_QuickGet("tlabel")) Then lab = SKW_QuickGet("tlabel") Else lab = SKW_QuickGet("flabel")
    ' The field takes the typed number's place in one step. Deleting first lets Word's smart cut take the space with it.
    Set r = d.Range(p.Start + labLen, p.Start + labLen + Len(num))
    d.Fields.Add r, wdFieldEmpty, Trim$(SeqCode(lab, 0)), False
    ' The heading number, separator and restart come from the same repair every caption gets.
    NumOne d, p, True, look, hh
    FixTyped = True
End Function

' Updates the numbers and what points at them: caption fields, cross-references, lists of tables and figures.
Private Sub UpdateNumbers(ByVal d As Object)
    Dim f As Object, i As Long
    On Error Resume Next
    SKW_Status "Updating the numbers and cross-references..."
    For Each f In d.Fields
        Select Case f.Type
            Case wdFieldSequence, wdFieldStyleRef, wdFieldRef, wdFieldPageRef
                If Not f.Locked Then f.Update
        End Select
    Next
    For i = 1 To d.TablesOfFigures.Count
        d.TablesOfFigures(i).Update
    Next
End Sub

Private Function NumberIntro() As String
    NumberIntro = QuickLine() & "The rule: " & RuleText() & ". A number counts as right when its field codes follow the rule, whatever number shows now."
End Function

' Report of every caption's number against one rule for the document: heading number then sequence, or sequential all the way through. Finds typed numbers, the wrong sequence name, a missing or wrong heading number and restarts. Changes nothing, then offers to repair the field codes in a copy.
Public Sub CaptionNumberingCheck()
    Dim s As String
    If SKW_IsQuiet() Then Exit Sub
    If Not Ready() Then Exit Sub
    SKW_Fixable = 0
    s = SKW_CaptionNumberCheck()
    If SKW_OfferFix(s, "The fix repairs the caption field codes in a copy of the document and updates the numbers and cross-references there. Caption text is not changed. Your document is not changed.") Then
        SKW_Show SKW_CaptionNumberFixCopy("")
    End If
End Sub

Public Function SKW_CaptionNumberCheck() As String
    Dim d As Object, msg As String, rows As New Collection, caps As Collection, typed As Collection, i As Long, p As Object, s As String
    Dim look As String, hh As Boolean, nBad As Long, nHead As Long, nPlain As Long, nLook As Long, nTyped As Long, changed As Object
    Dim f As Object, bm As String, nRef As Long, n As Long, was As Boolean, st As Long, wasSaved As Boolean
    Set SKW_LastReport = Nothing
    SKW_Fixable = 0
    msg = SKW_Begin("Caption numbering check", False)
    If Len(msg) > 0 Then SKW_CaptionNumberCheck = msg: Exit Function
    Set d = ActiveDocument
    wasSaved = d.Saved
    was = d.ActiveWindow.View.ShowFieldCodes
    d.ActiveWindow.View.ShowFieldCodes = False
    Set changed = CreateObject("Scripting.Dictionary")
    Set caps = CaptionParas(d)
    SKW_Steps caps.Count, "Caption"
    For i = caps.Count To 1 Step -1
        SKW_Step
        Set p = caps(i)
        s = NumOne(d, p, False, look, hh)
        If hh Then nHead = nHead + 1 Else nPlain = nPlain + 1
        If Len(s) > 0 Then nBad = nBad + 1: changed(p.Start) = True
        If Len(look) > 0 Then nLook = nLook + 1
        If Len(s) > 0 Or Len(look) > 0 Then rows.Add SKW_Row(SKW_Page(p), SKW_Excerpt(ParaText(p), 70), IIf(Len(s) > 0, s, ""), look)
    Next
    SKW_Status "Looking for typed numbers..."
    Set typed = TypedCaptions(d)
    For i = typed.Count To 1 Step -1
        Set p = typed(i)
        nTyped = nTyped + 1
        changed(p.Start) = True
        rows.Add SKW_Row(SKW_Page(p), SKW_Excerpt(ParaText(p), 70), "Number " & TypedNumber(p, n) & " is typed, so it does not renumber", "")
    Next
    ' Cross-references that point at a caption whose number the fix would change.
    SKW_Status "Reading the cross-references..."
    If changed.Count > 0 Then
        On Error Resume Next
        For Each f In d.Fields
            If f.Type = wdFieldRef Then
                bm = SKW_FieldBookmark(f)
                If d.Bookmarks.Exists(bm) Then
                    st = -1
                    st = d.Bookmarks(bm).Range.Paragraphs(1).Range.Start
                    If changed.Exists(st) Then nRef = nRef + 1
                End If
            End If
        Next
        On Error GoTo 0
    End If
    d.ActiveWindow.View.ShowFieldCodes = was
    ' Looking for Word's own Caption style adds it to a document that never used it, which marks the
    ' document changed though nothing a person sees has. The check changes nothing, so the mark is put back.
    If wasSaved And Not d.Saved Then d.Saved = True
    SKW_Fixable = nBad + nTyped
    SKW_Finish
    msg = caps.Count & " caption number(s) read: " & nBad & " do not follow the rule, " & nTyped & " typed, " & nLook & " to look at."
    SKW_Report "Caption Numbering Check", msg & " " & NumberIntro() & " In this document " & nHead & " caption(s) have a heading number and " & nPlain & " do not. " & _
        nRef & " cross-reference(s) point at a caption the fix would change; they are updated with it. Captions in text boxes, headers and footers are not read. Changes nothing.", _
        Array("Page", "Caption", "Finding", "Look at"), rows
    SKW_CaptionNumberCheck = msg
End Function

Public Function SKW_CaptionNumberFixCopy(ByVal target As String) As String
    Dim c As Object, caps As Collection, typed As Collection, p As Object, s As String, look As String, hh As Boolean
    Dim n As Long, nT As Long, msg As String, ed As String
    Set SKW_LastReport = Nothing
    On Error GoTo Fail
    msg = CopyOpen("captions renumbered", target, c)
    If Len(msg) > 0 Then SKW_CaptionNumberFixCopy = msg: Exit Function
    SKW_ProgressTitle "Caption numbering: repairing a copy"
    c.ActiveWindow.View.ShowFieldCodes = False
    Application.ScreenUpdating = False
    Set typed = TypedCaptions(c)
    For Each p In typed
        If FixTyped(c, p) Then nT = nT + 1
    Next
    Set caps = CaptionParas(c)
    SKW_Steps caps.Count, "Caption"
    For Each p In caps
        SKW_Step
        s = NumOne(c, p, True, look, hh)
        If Len(s) > 0 Then n = n + 1
    Next
    UpdateNumbers c
    SKW_Status "Saving the copy..."
    c.Save
    Application.ScreenUpdating = True
    Application.StatusBar = False
    SKW_ProgressDone
    SKW_CaptionNumberFixCopy = n & " caption number(s) repaired and " & nT & " typed number(s) made fields in a copy: " & LeafName(target) & "." & vbCr & _
        "Numbers, cross-references and lists of tables and figures are updated there. Your document is not changed. The copy is open."
    Exit Function
Fail:
    ed = Err.Description
    Application.ScreenUpdating = True
    Application.StatusBar = False
    SKW_ProgressDone
    SKW_CaptionNumberFixCopy = "Stopped: " & ed & ". Your document is not changed."
End Function

' Repairs the number of the caption at the cursor to follow the rule: its field codes, heading number, separator and restart. One Ctrl+Z undoes it.
Public Sub FixCaptionNumberHere()
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    If Not Ready() Then Exit Sub
    SKW_Show SKW_CaptionNumberHere()
End Sub

Public Function SKW_CaptionNumberHere() As String
    Dim d As Object, msg As String, p As Object, s As String, look As String, hh As Boolean, sq As Object, sr As Object, n As Long, did As Boolean, num As String
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Fix caption number", True)
    If Len(msg) > 0 Then SKW_CaptionNumberHere = msg: Exit Function
    Set d = ActiveDocument
    Set p = Selection.Paragraphs(1).Range
    NumberFields p, sq, sr
    If sq Is Nothing Then
        num = TypedNumber(p, n)
        If Len(num) > 0 Then did = FixTyped(d, p)
        SKW_Finish
        If did Then
            SKW_CaptionNumberHere = "The typed number " & num & " is now a field that follows the rule: " & RuleText() & "." & vbCr & _
                "Run Safe Update All Fields to renumber the captions after it." & vbCr & "One Ctrl+Z undoes it."
        Else
            SKW_CaptionNumberHere = "The paragraph at the cursor has no caption number: no field, and no typed number after " & SKW_QuickGet("tlabel") & " or " & SKW_QuickGet("flabel") & ". Nothing was changed."
        End If
        Exit Function
    End If
    s = NumOne(d, p, True, look, hh)
    SKW_Finish
    SKW_CaptionNumberHere = IIf(Len(s) > 0, "Caption number repaired. It had: " & s & ".", "The caption number already follows the rule. Nothing was changed.") & _
        vbCr & "The rule: " & RuleText() & "." & IIf(Len(look) > 0, vbCr & "Look at: " & look & ".", "") & _
        IIf(Len(s) > 0, vbCr & "Run Safe Update All Fields to renumber the captions after it." & vbCr & "One Ctrl+Z undoes it.", "")
End Function
