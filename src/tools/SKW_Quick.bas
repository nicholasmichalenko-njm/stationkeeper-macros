Attribute VB_Name = "SKW_Quick"
' =====================================================================
' Stationkeeper Tools: quick tools (1.10.0; one source in 1.11.0).
' Fast fixes for tables and captions. Every value comes from the
' Stationkeeper app's Global Standards workbook, and from nowhere else:
'   Table looks   the header and body rows' Tag styles, the keep list
'   Table types   the cell margins (the first type)
'   Captions      per label and part: Style, After number, Chapter level,
'                 Separator, Identifier, and the added column Action style
'   Settings      caption_abbreviations (optional)
' With no workbook, or a sheet missing, a tool says so and changes nothing.
'   QuickTableFormatHere / All   header row style, body row style, cell margins
'   QuickCaptionFormatHere / All caption style, what follows the number, the action caption style
'   CaptionNumberingCheck        caption field codes against the Captions sheet; fix in a copy
'   FixCaptionNumberHere         the caption at the cursor
' =====================================================================
Option Explicit

Private Const NEED_STD As String = "This tool reads the Global Standards workbook and nothing else. Run Choose Standards Workbook and pick the Stationkeeper app's Global Standards.xlsx."
Private mKey As String, mHead As String, mBody As String, mKeep As String, mMar(0 To 3) As Single
Private Const ABBREV As String = "U.S.; vs.; No.; Inc.; Co.; Corp.; Ltd.; e.g.; i.e.; etc.; Fig.; Rev.; Vol.; Sec.; Para.; Ref.; approx.; St.; Dr.; Mr.; Ms."

' ---------------------------------------------------------------------
' Reading the Global Standards workbook
' ---------------------------------------------------------------------
' A named value of the workbook's Settings sheet (Name, Value).
Private Function WbSetting(ByVal key As String) As String
    Dim r As Object
    For Each r In SKW_WbRows("Settings")
        If r.Exists("name") Then
            If LCase$(Trim$(r("name"))) = LCase$(key) Then
                If r.Exists("value") Then WbSetting = Trim$(r("value"))
                Exit Function
            End If
        End If
    Next
End Function

' A cell of a workbook row, with =name looked up on the Settings sheet.
Private Function WbVal(ByVal r As Object, ByVal key As String) As String
    Dim v As String
    If r.Exists(key) Then v = Trim$(r(key))
    If Left$(v, 1) = "=" Then v = WbSetting(Mid$(v, 2))
    WbVal = v
End Function

Private Function WbApproved(ByVal r As Object) As Boolean
    Dim s As String
    If r.Exists("status") Then s = LCase$(Trim$(r("status")))
    WbApproved = (Len(s) = 0 Or s = "approved")
End Function

' The table values, read once per change of the workbook: the header and body looks' Tag styles and the
' keep list from Table looks, through the first table type, and that type's cell margins in points.
Private Sub LoadTable()
    Dim types As Collection, roles As Object, m As Variant, i As Long, k As String
    k = SKW_StandardsPath() & "|" & SKW_StdChanged()
    If k = mKey Then Exit Sub
    mKey = k
    mHead = "": mBody = "": mKeep = ""
    For i = 0 To 3
        mMar(i) = -1
    Next
    On Error GoTo Done
    If Not SKW_AppStandards() Then Exit Sub
    Set types = SKW_TableTypes()
    If types.Count = 0 Then Exit Sub
    Set roles = SKW_TableRoles(types(1)("type"))
    mHead = roles("header")("paragraph style")
    mBody = roles("long text")("paragraph style")
    mKeep = roles("keep")("recognize")
    m = Split(types(1)("margins pt") & ";;;", ";")
    For i = 0 To 3
        If Len(Trim$(m(i))) > 0 And IsNumeric(Trim$(m(i))) Then mMar(i) = CSng(Trim$(m(i)))
    Next
Done:
End Sub

' The values the tools read by name.
Public Function SKW_QuickGet(ByVal key As String) As String
    Dim v As String
    Select Case LCase$(key)
        Case "head": LoadTable: SKW_QuickGet = mHead
        Case "body": LoadTable: SKW_QuickGet = mBody
        Case "keep": LoadTable: SKW_QuickGet = mKeep
        Case "tlabel": SKW_QuickGet = SKW_Setting("TableLabel", "Table")
        Case "flabel": SKW_QuickGet = SKW_Setting("FigureLabel", "Figure")
        Case "abbrev"
            v = WbSetting("caption_abbreviations")
            If Len(v) > 0 Then SKW_QuickGet = v Else SKW_QuickGet = ABBREV
    End Select
End Function

' The first table type's cell margin for a side (0 top, 1 bottom, 2 left, 3 right), in points. -1 when the type sets none.
Private Function MarginPts(ByVal side As Long) As Single
    LoadTable
    MarginPts = mMar(side)
End Function

Private Function MarginsText() As String
    Dim i As Long, s As String, nm As Variant
    nm = Array("top", "bottom", "left", "right")
    For i = 0 To 3
        s = s & IIf(Len(s) > 0, ", ", "") & nm(i) & " " & IIf(MarginPts(i) >= 0, Trim$(Str$(Round(MarginPts(i), 2))) & " pt", "as it is")
    Next
    MarginsText = s
End Function

' Where the values come from, for a report's first lines.
Private Function SourceLine() As String
    SourceLine = "Every value comes from the Global Standards" & IIf(Len(SKW_StdChanged()) > 0, " (" & SKW_StdChanged() & "). ", ". ")
    If SKW_Effort()("ok") Then SourceLine = SourceLine & "An effort is linked (" & SKW_Effort()("effort") & _
        "): its own standards cannot be read yet, so these are the Global Standards and the effort may differ. "
End Function

' The Captions sheet's Approved rows: label, part, level (0: no heading number), sep, id, after, style, action.
Private Function CapRows() As Collection
    Dim c As New Collection, r As Object, d As Object
    Set CapRows = c
    On Error GoTo Done
    For Each r In SKW_WbRows("Captions")
        If WbApproved(r) And Len(WbVal(r, "label")) > 0 Then
            Set d = CreateObject("Scripting.Dictionary")
            d("label") = WbVal(r, "label")
            d("part") = LCase$(WbVal(r, "part"))
            d("level") = CLng(Val(WbVal(r, "chapter level")))
            d("sep") = WbVal(r, "separator")
            If Len(d("sep")) = 0 Then d("sep") = "-"
            d("id") = WbVal(r, "identifier")
            If Len(d("id")) = 0 Then d("id") = d("label")
            d("after") = LCase$(WbVal(r, "after number"))
            d("style") = WbVal(r, "style")
            d("action") = WbVal(r, "action style")
            If Len(d("action")) = 0 Then d("action") = WbSetting("caption_action_style")
            c.Add d
        End If
    Next
Done:
End Function

' The Captions row for a label: the one whose Chapter level is lvl; with lvl below zero, the body row
' (or the first row for the label). Nothing when the sheet has no such row.
Private Function CapRow(ByVal rows As Collection, ByVal label As String, ByVal lvl As Long) As Object
    Dim d As Object, first As Object
    For Each d In rows
        If LCase$(d("label")) = LCase$(Trim$(label)) Then
            If lvl >= 0 Then
                If d("level") = lvl Then Set CapRow = d: Exit Function
            Else
                If first Is Nothing Then Set first = d
                If d("part") = "body" Then Set CapRow = d: Exit Function
            End If
        End If
    Next
    If lvl < 0 Then Set CapRow = first
End Function

' What the standards put between a caption's number and its title. Chr(0): the standards do not say.
Private Function AfterText(ByVal code As String) As String
    Select Case LCase$(Trim$(code))
        Case "tab": AfterText = vbTab
        Case "period_tab": AfterText = "." & vbTab
        Case "period_space": AfterText = ". "
        Case "colon_space": AfterText = ": "
        Case "space": AfterText = " "
        Case Else: AfterText = Chr$(0)
    End Select
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

' Reads one table against the Global Standards, and sets it when doIt. Returns the findings ("" when it is right).
Private Function QuickTable(ByVal d As Object, ByVal t As Object, ByVal doIt As Boolean, ByRef hdr As Long, ByRef look As String) As String
    Dim c As Object, p As Object, ri As Long, want As String, nm As String, keep As Object, v As Variant, s As String
    Dim stHead As Object, stBody As Object, nHead As Long, nBody As Long, st As Object
    Dim pts As Single, k As Variant, i As Long, mar As Boolean, descr As String, np As Long
    look = ""
    On Error Resume Next
    descr = LCase$(t.Descr)
    On Error GoTo 0
    If InStr(descr, "skip") > 0 Then look = "Marked skip in its Alt Text": hdr = 0: Exit Function
    hdr = HeaderRowOf(d, t)
    Set stHead = StyleIn(d, SKW_QuickGet("head"))
    Set stBody = StyleIn(d, SKW_QuickGet("body"))
    If stHead Is Nothing And Len(Trim$(SKW_QuickGet("head"))) > 0 Then look = "Style " & Trim$(SKW_QuickGet("head")) & " is not in this document (Apply Table Roles makes it from its look)"
    If stBody Is Nothing And Len(Trim$(SKW_QuickGet("body"))) > 0 Then look = look & IIf(Len(look) > 0, "; ", "") & "Style " & Trim$(SKW_QuickGet("body")) & " is not in this document"
    Set keep = CreateObject("Scripting.Dictionary")
    keep.CompareMode = 1
    For Each v In Split(SKW_QuickGet("keep"), ";")
        If Len(Trim$(v)) > 0 Then keep(Trim$(v)) = True
    Next
    ' Cell margins. Word keeps left and right margins in steps of its own (0.08 in is held as 5.4 points), so a
    ' difference under 0.4 point is no difference.
    For i = 0 To 3
        pts = MarginPts(i)
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
    TableIntro = SourceLine() & "Header row: " & IIf(Len(SKW_QuickGet("head")) > 0, SKW_QuickGet("head"), "as it is") & _
        ". Body rows: " & IIf(Len(SKW_QuickGet("body")) > 0, SKW_QuickGet("body"), "as they are") & _
        ". Cell margins: " & MarginsText() & ". The header row is row 2 when row 1 holds a caption number, else row 1" & _
        ". A caption row is left alone. Styles left alone in body rows: " & IIf(Len(SKW_QuickGet("keep")) > 0, SKW_QuickGet("keep"), "none") & _
        ". Only the header row, the body rows and the cell margins are set: Apply Table Roles sets every role, fill and edge."
End Function

' 1.10.0 and 1.10.1 kept "quick values" in the registry. They are not read any more; this removes them.
Private Sub DropOldValues()
    On Error Resume Next
    DeleteSetting SKW_APP, "Quick"
End Sub

' Sets the table at the cursor from the Global Standards: the header row's style, the body rows' style and the cell margins. One Ctrl+Z undoes it.
Public Sub QuickTableFormatHere()
    DropOldValues
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    If Not Selection.Information(wdWithInTable) Then SKW_Say "Put the cursor in a table first.": Exit Sub
    SKW_Show SKW_QuickTableHere()
End Sub

Public Function SKW_QuickTableHere() As String
    Dim d As Object, msg As String, t As Object, s As String, hdr As Long, look As String
    Set SKW_LastReport = Nothing
    If Documents.Count > 0 Then
        If Not Selection.Information(wdWithInTable) Then SKW_QuickTableHere = "Put the cursor in a table first.": Exit Function
    End If
    If Not SKW_AppStandards() Then SKW_QuickTableHere = NEED_STD & " It needs the Table looks sheet.": Exit Function
    msg = SKW_Begin("Quick table format", True)
    If Len(msg) > 0 Then SKW_QuickTableHere = msg: Exit Function
    Set d = ActiveDocument
    Set t = Selection.Tables(1)
    s = QuickTable(d, t, True, hdr, look)
    SKW_Finish
    If hdr = 0 Then SKW_QuickTableHere = look & ". Nothing was changed.": Exit Function
    SKW_QuickTableHere = IIf(Len(s) > 0, "Table set from the Global Standards (header row " & hdr & "). It had: " & s & ".", _
        "The table already matches the Global Standards (header row " & hdr & "). Nothing was changed.") & _
        IIf(Len(look) > 0, vbCr & "Look at: " & look & ".", "") & IIf(Len(s) > 0, vbCr & "One Ctrl+Z undoes it.", "")
End Function

' Report of every table against the Global Standards: the header row's style, the body rows' style and the cell margins. Changes nothing, then offers to set every table in a copy.
Public Sub QuickTableFormatAll()
    DropOldValues
    Dim s As String
    If SKW_IsQuiet() Then Exit Sub
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
    If Not SKW_AppStandards() Then SKW_QuickTableCheck = NEED_STD & " It needs the Table looks sheet.": Exit Function
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
    If Not SKW_AppStandards() Then SKW_QuickTableFixCopy = NEED_STD & " It needs the Table looks sheet.": Exit Function
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

' The Captions row a caption follows: by its label word (or, for a label the sheet lacks, by what its
' number counts), then by the heading level its number takes. note: why the row is a guess.
Private Function RowOf(ByVal p As Object, ByVal sq As Object, ByVal sr As Object, ByVal rows As Collection, ByRef note As String) As Object
    Dim lab As String, d As Object, id As String, n As Long, some As Boolean
    note = ""
    lab = CapLabel(p)
    If CapRow(rows, lab, -1) Is Nothing Then
        id = SeqParts(sq.Code.Text)("id")
        lab = ""
        For Each d In rows
            If LCase$(d("id")) = LCase$(id) And Len(lab) = 0 Then lab = d("label")
        Next
        If Len(lab) = 0 Then Exit Function
    End If
    If Not sr Is Nothing Then
        n = RefLevel(sr.Code.Text)
        If n > 0 Then
            Set RowOf = CapRow(rows, lab, n)
            If RowOf Is Nothing Then
                Set RowOf = CapRow(rows, lab, -1)
                ' A label the sheet numbers by heading, at a level it does not name, is left for a person.
                ' A label the sheet counts straight through has one row: a heading number is then a finding.
                For Each d In rows
                    If LCase$(d("label")) = LCase$(lab) And d("level") > 0 Then some = True
                Next
                If some Then note = "Takes its heading number from level " & n & ", and the Captions sheet has no " & lab & " row at that level"
            End If
            Exit Function
        End If
        note = "Takes its heading number from " & Norm(sr.Code.Text) & ", a style by name"
    End If
    Set RowOf = CapRow(rows, lab, -1)
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

' Reads one caption against its Captions row and sets it when doIt. Returns the findings.
Private Function CapOne(ByVal d As Object, ByVal p As Object, ByVal doIt As Boolean, ByVal rows As Collection, ByRef kind As String, _
                        ByRef title As String, ByRef action As String, ByRef look As String) As String
    Dim sq As Object, sr As Object, pos As Long, pe As Long, sp As Long, s As String, ch As String, row As Object, note As String
    Dim stP As Object, stA As Object, wantP As String, wantA As String, tr As Object, sepTxt As String, r As Object, want As String
    kind = "": title = "": action = "": look = ""
    NumberFields p, sq, sr
    If sq Is Nothing Then Exit Function
    ' A number that repeats the one before it (\c) sits in running text: it is not a caption.
    If SeqParts(sq.Code.Text)("c") Then Exit Function
    Set row = RowOf(p, sq, sr, rows, note)
    If row Is Nothing Then Exit Function
    kind = row("label")
    wantP = row("style")
    wantA = row("action")
    Set stP = StyleIn(d, wantP)
    Set stA = StyleIn(d, wantA)
    If stP Is Nothing And Len(wantP) > 0 Then look = "Style " & wantP & " is not in this document"
    pe = p.End - 1
    pos = FEnd(sq)
    Do While pos < pe
        ch = d.Range(pos, pos + 1).Text
        If Len(ch) = 0 Then Exit Do
        If InStr(" .:-" & vbTab & Chr$(160) & Chr$(30) & ChrW$(8211) & ChrW$(8212), ch) = 0 Then Exit Do
        pos = pos + 1
    Loop
    ' The paragraph style first: the character style goes on top of it.
    If Not stP Is Nothing Then
        If Not Same(SKW_ParaStyleName(p), stP.NameLocal) Then
            s = "Paragraph style is " & SKW_ParaStyleName(p)
            If doIt Then p.Style = stP
        End If
    End If
    If pos >= pe Then look = look & IIf(Len(look) > 0, "; ", "") & "No title after the number": CapOne = s: Exit Function
    ' What follows the number is the Captions row's After number, so every title starts at the same place.
    want = AfterText(row("after"))
    If want <> Chr$(0) Then
        sepTxt = d.Range(FEnd(sq), pos).Text
        If sepTxt <> want Then
            s = s & IIf(Len(s) > 0, "; ", "") & "The number is followed by " & SepWords(sepTxt) & ", not " & SepWords(want)
            If doIt Then
                Set r = d.Range(FEnd(sq), pos)
                If r.End > r.Start Then r.Text = want Else r.InsertBefore want
                pos = FEnd(sq) + Len(want)
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
            look = look & IIf(Len(look) > 0, "; ", "") & IIf(Len(wantA) > 0, "Style " & wantA & " is not in this document", _
                "The Captions sheet names no Action style for " & row("label") & ", so the action caption is not styled")
        Else
            ' The period goes with the action caption, so a list built from the title's style shows no period.
            If Not Same(CharStyleAt(d, sp), stA.NameLocal) Or Not Same(CharStyleAt(d, pe - 1), stA.NameLocal) Then
                s = s & IIf(Len(s) > 0, "; ", "") & "The period and action caption are not in " & stA.NameLocal
                If doIt Then d.Range(sp, pe).Style = stA
            End If
            If Same(CharStyleAt(d, pos), stA.NameLocal) Or Same(CharStyleAt(d, sp - 1), stA.NameLocal) Then
                s = s & IIf(Len(s) > 0, "; ", "") & "The title carries " & stA.NameLocal
                If doIt Then d.Range(pos, sp).Style = d.Styles(wdStyleDefaultParagraphFont)
            End If
        End If
    Else
        title = tr.Text
        If Right$(RTrim$(title), 1) = "." Then look = look & IIf(Len(look) > 0, "; ", "") & "Ends in a period with no action caption"
        If Not stA Is Nothing Then
            If Same(CharStyleAt(d, pos), stA.NameLocal) Then
                s = s & IIf(Len(s) > 0, "; ", "") & "The title carries " & stA.NameLocal & " and has no action caption"
                If doIt Then d.Range(pos, pe).Style = d.Styles(wdStyleDefaultParagraphFont)
            End If
        End If
    End If
    CapOne = s
End Function

Private Function CaptionIntro(ByVal rows As Collection) As String
    Dim d As Object, s As String
    For Each d In rows
        s = s & IIf(Len(s) > 0, "; ", "") & d("label") & IIf(Len(d("part")) > 0, " (" & d("part") & ")", "") & ": style " & _
            IIf(Len(d("style")) > 0, d("style"), "as it is") & ", after the number " & _
            IIf(AfterText(d("after")) = Chr$(0), "as it is", SepWords(AfterText(d("after")))) & ", action caption " & _
            IIf(Len(d("action")) > 0, d("action"), "not styled")
    Next
    CaptionIntro = SourceLine() & "The Captions sheet: " & s & ". The action caption starts at the first period followed by a space, and the period goes with it. " & _
        "A period after an abbreviation (" & SKW_QuickGet("abbrev") & ") or before a lower-case word does not count. A caption with no period has no action caption. " & _
        "A caption already in its style takes its tab stops and indents from the style."
End Function

' Sets the caption at the cursor from the Global Standards: its paragraph style, what follows the number, the style's tab stops and indents, and the action caption's style from the first period on. One Ctrl+Z undoes it.
Public Sub QuickCaptionFormatHere()
    DropOldValues
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    SKW_Show SKW_QuickCaptionHere()
End Sub

Public Function SKW_QuickCaptionHere() As String
    Dim d As Object, msg As String, p As Object, s As String, kind As String, title As String, action As String, look As String, was As Boolean
    Dim rows As Collection
    Set SKW_LastReport = Nothing
    Set rows = CapRows()
    If rows.Count = 0 Then SKW_QuickCaptionHere = NEED_STD & " It needs the Captions sheet.": Exit Function
    msg = SKW_Begin("Quick caption format", True)
    If Len(msg) > 0 Then SKW_QuickCaptionHere = msg: Exit Function
    Set d = ActiveDocument
    was = d.ActiveWindow.View.ShowFieldCodes
    d.ActiveWindow.View.ShowFieldCodes = False
    Set p = Selection.Paragraphs(1).Range
    s = CapOne(d, p, True, rows, kind, title, action, look)
    d.ActiveWindow.View.ShowFieldCodes = was
    SKW_Finish
    If Len(kind) = 0 Then
        SKW_QuickCaptionHere = "The paragraph at the cursor is not a caption the Captions sheet knows: it has no caption number field, or its label has no row. Nothing was changed."
        Exit Function
    End If
    SKW_QuickCaptionHere = IIf(Len(s) > 0, "Caption set from the Global Standards. It had: " & s & ".", "The caption already matches the Global Standards. Nothing was changed.") & _
        vbCr & "Title: " & title & IIf(Len(action) > 0, vbCr & "Action caption: " & action, vbCr & "No action caption.") & _
        IIf(Len(look) > 0, vbCr & "Look at: " & look & ".", "") & IIf(Len(s) > 0, vbCr & "One Ctrl+Z undoes it.", "")
End Function

' Report of every caption against the Global Standards' Captions sheet, with each caption shown as title and action caption. Changes nothing, then offers to set every caption in a copy.
Public Sub QuickCaptionFormatAll()
    DropOldValues
    Dim s As String
    If SKW_IsQuiet() Then Exit Sub
    SKW_Fixable = 0
    s = SKW_QuickCaptionCheck()
    If SKW_OfferFix(s, "The fix sets every caption's styles in a copy of the document. Your document is not changed.") Then
        SKW_Show SKW_QuickCaptionFixCopy("")
    End If
End Sub

Public Function SKW_QuickCaptionCheck() As String
    Dim d As Object, msg As String, rows As New Collection, caps As Collection, i As Long, p As Object, s As String, std As Collection
    Dim kind As String, title As String, action As String, look As String, nBad As Long, n As Long, nAct As Long, nLook As Long, was As Boolean
    Set SKW_LastReport = Nothing
    SKW_Fixable = 0
    Set std = CapRows()
    If std.Count = 0 Then SKW_QuickCaptionCheck = NEED_STD & " It needs the Captions sheet.": Exit Function
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
        s = CapOne(d, p, False, std, kind, title, action, look)
        If Len(kind) > 0 Then
            n = n + 1
            If Len(action) > 0 Then nAct = nAct + 1
            If Len(s) > 0 Then nBad = nBad + 1
            If Len(look) > 0 Then nLook = nLook + 1
            rows.Add SKW_Row(kind, SKW_Page(p), title, action, IIf(Len(s) > 0, s, "Matches"), look)
        End If
    Next
    d.ActiveWindow.View.ShowFieldCodes = was
    SKW_Fixable = nBad
    SKW_Finish
    msg = n & " caption(s) read, " & nAct & " with an action caption, " & nBad & " to set, " & nLook & " to look at."
    SKW_Report "Quick Caption Format", msg & " " & CaptionIntro(std) & " Captions in text boxes, headers and footers are not read. Changes nothing.", _
        Array("Label", "Page", "Title", "Action caption", "Finding", "Look at"), rows
    SKW_QuickCaptionCheck = msg
End Function

Public Function SKW_QuickCaptionFixCopy(ByVal target As String) As String
    Dim c As Object, caps As Collection, p As Object, s As String, kind As String, title As String, action As String, look As String
    Dim n As Long, msg As String, ed As String, std As Collection
    Set SKW_LastReport = Nothing
    Set std = CapRows()
    If std.Count = 0 Then SKW_QuickCaptionFixCopy = NEED_STD & " It needs the Captions sheet.": Exit Function
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
        s = CapOne(c, p, True, std, kind, title, action, look)
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

Private Function RuleText(ByVal rows As Collection) As String
    Dim d As Object, s As String
    For Each d In rows
        s = s & IIf(Len(s) > 0, "; ", "") & d("label") & IIf(Len(d("part")) > 0, " (" & d("part") & ")", "") & ": " & _
            IIf(d("level") > 0, "the number of heading level " & d("level") & ", then """ & d("sep") & """, then a count that restarts under that heading", _
                "counted straight through, with no heading number and no restart") & ", counted as " & d("id")
    Next
    RuleText = s
End Function

' Reads one caption's number fields against its Captions row and repairs them when doIt. Returns the findings.
' look: something left alone on purpose. hasHead: the caption has a heading number.
Private Function NumOne(ByVal d As Object, ByVal p As Object, ByVal doIt As Boolean, ByVal rows As Collection, ByRef look As String, ByRef hasHead As Boolean) As String
    Dim sq As Object, sr As Object, q As Object, lvl As Long, row As Object, note As String
    Dim wantId As String, s As String, sep As String, wantSep As String, newCode As String, r As Object, sepBad As Boolean
    look = "": hasHead = False
    NumberFields p, sq, sr
    If sq Is Nothing Then Exit Function
    Set q = SeqParts(sq.Code.Text)
    If q("c") Then look = "Repeats the number before it (\c), so it is not a caption's own number: left alone": Exit Function
    hasHead = Not sr Is Nothing
    Set row = RowOf(p, sq, sr, rows, note)
    If row Is Nothing Then look = "The Captions sheet has no row for the label " & CapLabel(p) & ": left alone": Exit Function
    If Len(note) > 0 Then look = note & ": left alone": Exit Function
    lvl = row("level")
    wantId = row("id")
    wantSep = row("sep")
    If Not Same(q("id"), wantId) Then s = "Counts as " & q("id") & ", not " & wantId
    If Len(q("fmt")) > 0 And UCase$(q("fmt")) <> "ARABIC" Then s = s & IIf(Len(s) > 0, "; ", "") & "Number format is " & q("fmt")
    If lvl = 0 Then
        If Len(q("s")) > 0 Then s = s & IIf(Len(s) > 0, "; ", "") & "Restarts under heading level " & q("s")
        If Len(q("r")) > 0 Then s = s & IIf(Len(s) > 0, "; ", "") & "Restarts at " & q("r")
        If hasHead Then s = s & IIf(Len(s) > 0, "; ", "") & "Has a heading number"
    Else
        If sr Is Nothing Then
            s = s & IIf(Len(s) > 0, "; ", "") & "No heading number"
        Else
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
    If lvl = 0 Then
        If hasHead Then d.Range(FStart(sr), FStart(sq)).Delete
    ElseIf sr Is Nothing Then
        Set r = d.Range(FStart(sq), FStart(sq))
        d.Fields.Add r, wdFieldEmpty, "STYLEREF " & lvl & " \s", False
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
Private Function TypedNumber(ByVal p As Object, ByVal rows As Collection, ByRef labLen As Long) As String
    Static re As Object
    Dim m As Object, txt As String, d As Object, labs As String, seen As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True
    End If
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    For Each d In rows
        If Not seen.Exists(d("label")) Then seen(d("label")) = True: labs = labs & IIf(Len(labs) > 0, "|", "") & d("label")
    Next
    If Len(labs) = 0 Then Exit Function
    re.Pattern = "^((?:" & labs & ")[ \u00A0]+)([A-Za-z]?[0-9]+(?:[-.\u2011\u2013][0-9]+)*)"
    txt = ParaText(p)
    If Not re.Test(txt) Then Exit Function
    Set m = re.Execute(txt)(0)
    labLen = Len(m.SubMatches(0))
    TypedNumber = m.SubMatches(1)
End Function

' Paragraphs in the caption styles that hold no caption number field but start with a typed number. Last first.
Private Function TypedCaptions(ByVal d As Object, ByVal rows As Collection) As Collection
    Dim c As New Collection, back As New Collection, seen As Object, nm As Variant, st As Object, r As Object, p As Object, sq As Object, sr As Object
    Dim names As Object, n As Long, i As Long, starts As Object, keys As Variant, j As Long, tmp As Variant, row As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = 1
    ' The caption styles the Captions sheet names, and Word's own Caption style.
    names(SKW_Setting("CaptionStyle", "Caption")) = True
    For Each row In rows
        If Len(row("style")) > 0 Then names(row("style")) = True
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
                            If Len(TypedNumber(p.Range, rows, n)) > 0 Then Set starts(p.Range.Start) = p.Range
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
Private Function FixTyped(ByVal d As Object, ByVal p As Object, ByVal rows As Collection) As Boolean
    Dim num As String, labLen As Long, r As Object, row As Object, look As String, hh As Boolean
    num = TypedNumber(p, rows, labLen)
    If Len(num) = 0 Then Exit Function
    If d.Range(p.Start, p.Start + labLen + Len(num)).Fields.Count > 0 Then Exit Function
    Set row = CapRow(rows, CapLabel(p), -1)
    If row Is Nothing Then Exit Function
    ' The field takes the typed number's place in one step. Deleting first lets Word's smart cut take the space with it.
    Set r = d.Range(p.Start + labLen, p.Start + labLen + Len(num))
    d.Fields.Add r, wdFieldEmpty, Trim$(SeqCode(row("id"), 0)), False
    ' The heading number, separator and restart come from the same repair every caption gets.
    NumOne d, p, True, rows, look, hh
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

Private Function NumberIntro(ByVal rows As Collection) As String
    NumberIntro = SourceLine() & "The Captions sheet: " & RuleText(rows) & ". A number counts as right when its field codes follow its row, whatever number shows now. " & _
        "A caption with no heading number is read against the body row for its label."
End Function

' Report of every caption's number against the Global Standards' Captions sheet: heading number then sequence, or counted straight through. Finds typed numbers, the wrong sequence name, a missing or wrong heading number and restarts. Changes nothing, then offers to repair the field codes in a copy.
Public Sub CaptionNumberingCheck()
    DropOldValues
    Dim s As String
    If SKW_IsQuiet() Then Exit Sub
    SKW_Fixable = 0
    s = SKW_CaptionNumberCheck()
    If SKW_OfferFix(s, "The fix repairs the caption field codes in a copy of the document and updates the numbers and cross-references there. Caption text is not changed. Your document is not changed.") Then
        SKW_Show SKW_CaptionNumberFixCopy("")
    End If
End Sub

Public Function SKW_CaptionNumberCheck() As String
    Dim d As Object, msg As String, rows As New Collection, caps As Collection, typed As Collection, i As Long, p As Object, s As String
    Dim look As String, hh As Boolean, nBad As Long, nHead As Long, nPlain As Long, nLook As Long, nTyped As Long, changed As Object
    Dim f As Object, bm As String, nRef As Long, n As Long, was As Boolean, st As Long, wasSaved As Boolean, std As Collection
    Set SKW_LastReport = Nothing
    SKW_Fixable = 0
    Set std = CapRows()
    If std.Count = 0 Then SKW_CaptionNumberCheck = NEED_STD & " It needs the Captions sheet.": Exit Function
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
        s = NumOne(d, p, False, std, look, hh)
        If hh Then nHead = nHead + 1 Else nPlain = nPlain + 1
        If Len(s) > 0 Then nBad = nBad + 1: changed(p.Start) = True
        If Len(look) > 0 Then nLook = nLook + 1
        If Len(s) > 0 Or Len(look) > 0 Then rows.Add SKW_Row(SKW_Page(p), SKW_Excerpt(ParaText(p), 70), IIf(Len(s) > 0, s, ""), look)
    Next
    SKW_Status "Looking for typed numbers..."
    Set typed = TypedCaptions(d, std)
    For i = typed.Count To 1 Step -1
        Set p = typed(i)
        nTyped = nTyped + 1
        changed(p.Start) = True
        rows.Add SKW_Row(SKW_Page(p), SKW_Excerpt(ParaText(p), 70), "Number " & TypedNumber(p, std, n) & " is typed, so it does not renumber", "")
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
    SKW_Report "Caption Numbering Check", msg & " " & NumberIntro(std) & " In this document " & nHead & " caption(s) have a heading number and " & nPlain & " do not. " & _
        nRef & " cross-reference(s) point at a caption the fix would change; they are updated with it. Captions in text boxes, headers and footers are not read. Changes nothing.", _
        Array("Page", "Caption", "Finding", "Look at"), rows
    SKW_CaptionNumberCheck = msg
End Function

Public Function SKW_CaptionNumberFixCopy(ByVal target As String) As String
    Dim c As Object, caps As Collection, typed As Collection, p As Object, s As String, look As String, hh As Boolean
    Dim n As Long, nT As Long, msg As String, ed As String, std As Collection
    Set SKW_LastReport = Nothing
    Set std = CapRows()
    If std.Count = 0 Then SKW_CaptionNumberFixCopy = NEED_STD & " It needs the Captions sheet.": Exit Function
    On Error GoTo Fail
    msg = CopyOpen("captions renumbered", target, c)
    If Len(msg) > 0 Then SKW_CaptionNumberFixCopy = msg: Exit Function
    SKW_ProgressTitle "Caption numbering: repairing a copy"
    c.ActiveWindow.View.ShowFieldCodes = False
    Application.ScreenUpdating = False
    Set typed = TypedCaptions(c, std)
    For Each p In typed
        If FixTyped(c, p, std) Then nT = nT + 1
    Next
    Set caps = CaptionParas(c)
    SKW_Steps caps.Count, "Caption"
    For Each p In caps
        SKW_Step
        s = NumOne(c, p, True, std, look, hh)
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

' Repairs the number of the caption at the cursor to follow its row on the Global Standards' Captions sheet: its field codes, heading number, separator and restart. One Ctrl+Z undoes it.
Public Sub FixCaptionNumberHere()
    DropOldValues
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    SKW_Show SKW_CaptionNumberHere()
End Sub

Public Function SKW_CaptionNumberHere() As String
    Dim d As Object, msg As String, p As Object, s As String, look As String, hh As Boolean, sq As Object, sr As Object, n As Long, did As Boolean, num As String
    Dim std As Collection
    Set SKW_LastReport = Nothing
    Set std = CapRows()
    If std.Count = 0 Then SKW_CaptionNumberHere = NEED_STD & " It needs the Captions sheet.": Exit Function
    msg = SKW_Begin("Fix caption number", True)
    If Len(msg) > 0 Then SKW_CaptionNumberHere = msg: Exit Function
    Set d = ActiveDocument
    Set p = Selection.Paragraphs(1).Range
    NumberFields p, sq, sr
    If sq Is Nothing Then
        num = TypedNumber(p, std, n)
        If Len(num) > 0 Then did = FixTyped(d, p, std)
        SKW_Finish
        If did Then
            SKW_CaptionNumberHere = "The typed number " & num & " is now a field that follows the Captions sheet." & vbCr & _
                "Run Safe Update All Fields to renumber the captions after it." & vbCr & "One Ctrl+Z undoes it."
        Else
            SKW_CaptionNumberHere = "The paragraph at the cursor has no caption number: no field, and no typed number after a label the Captions sheet names. Nothing was changed."
        End If
        Exit Function
    End If
    s = NumOne(d, p, True, std, look, hh)
    SKW_Finish
    SKW_CaptionNumberHere = IIf(Len(s) > 0, "Caption number repaired. It had: " & s & ".", "The caption number already follows the Captions sheet. Nothing was changed.") & _
        IIf(Len(look) > 0, vbCr & "Look at: " & look & ".", "") & _
        IIf(Len(s) > 0, vbCr & "Run Safe Update All Fields to renumber the captions after it." & vbCr & "One Ctrl+Z undoes it.", "")
End Function
