Attribute VB_Name = "SKW_Proof"
' =====================================================================
' Stationkeeper Tools: work PC proofs (1.21.0).
' Small tools that find out what Word, Excel and a synced folder allow on a PC,
' before the workbook layout is built on it. They change no document and no
' workbook of yours. Each report holds setting names and counts only: no style
' name, no value, no file name and no term of the team's.
'   ProofThisPC               proof 1: this PC, and what Word lets a macro do to a style (in a scratch document)
'   ProofThisDocumentStyles   proof 2: the styles of the document in front, read, rebuilt in a scratch document and compared
'   ProofAWorkbook            proof 3: how long a workbook takes to read, and what is still there after a save (on a copy)
'   ProofSharedWorkbook       proof 4: writing a scratch workbook in a synced folder while someone else has it open
' The proofs are P1 to P7 of the build order page. They leave the kit once they are answered.
' =====================================================================
Option Explicit

Private Const N_SET As Long = 49
Private Const SET_NAMES As String = "Type|Based on|Next style|Automatically update|Linked style|Locked|Show in the Styles gallery|Semi-hidden|Unhide when used|Order in the gallery|" & _
    "Font|Theme font|Size (pt)|Bold|Italic|Underline|All caps|Small caps|Hidden text|Strikethrough|Double strikethrough|Superscript or subscript|Color (hex)|" & _
    "Character spacing (pt)|Kerning from (pt)|Character shading (hex, or none)|Do not check spelling|Alignment|Outline level|Keep with next|Keep lines together|" & _
    "Page break before|Widow control|No hyphenation|Space before (pt)|Space after (pt)|Line spacing (lines, or pt if exact)|Line rule|" & _
    "No space between paragraphs of the same style|Left indent (pt)|Right indent (pt)|First line (pt)|Hanging (pt)|Tab stops|Border above|Border below|Border left|Border right|Shading (hex, or none)"
Private Const KNOWN_TABS As String = "|read me|settings|styles|fonts allowed|table looks|table types|page setups|heading numbering|captions|list numbering|page numbering|templates|theme|phrases|" & _
    "placeholders|ignore|style map|page limits|quick parts|pages|stationkeeper admins|fields|authors|compliance words|piece types|document types|suggestion rules|columns|about|" & _
    "allowed styles|fonts|acronyms|dictionaries|banned phrases|history|"
Private Const SCRATCH_BOOK As String = "Stationkeeper proof scratch.xlsx"
Private Const NOTE_END As String = "This report holds setting names and counts only. Copy all of it into the Stationkeeper Macros thread."

Private mRows As Collection
Private mBad As Long
Private mRe As Object

' ---------------------------------------------------------------------
' The report: its own, so no document or workbook name goes into it.
' ---------------------------------------------------------------------
Private Sub ProofReport(ByVal title As String, ByVal intro As String, ByVal cols As Variant, ByVal rows As Collection)
    Dim src As Object, d As Object, h As String, r As Variant, i As Long, tmp As String, v As Variant, parts() As String, n As Long, one As String
    On Error Resume Next
    If Documents.Count > 0 Then Set src = ActiveDocument
    On Error GoTo 0
    h = "<html><head><meta charset=""utf-8""><style>body{font-family:Calibri,sans-serif;font-size:10pt}" & _
        "h1{font-size:16pt}table{border-collapse:collapse}th,td{border:1px solid #999;padding:2pt 4pt;vertical-align:top;font-size:9pt}" & _
        "th{background:#DDEBF7;text-align:left}</style></head><body>" & _
        "<h1>" & SKW_Html(title) & "</h1><p>Stationkeeper Macros " & SKW_VERSION & " &nbsp; " & Format$(Now, "yyyy-mm-dd hh:nn") & "</p><p>" & SKW_Html(intro) & "</p>"
    If rows.Count > 0 Then
        h = h & "<table><thead><tr>"
        For Each v In cols
            h = h & "<th>" & SKW_Html(CStr(v)) & "</th>"
        Next
        h = h & "</tr></thead><tbody>"
        ReDim parts(1 To rows.Count)
        For Each r In rows
            n = n + 1
            one = "<tr>"
            For i = LBound(r) To UBound(r)
                one = one & "<td>" & SKW_Html(CStr(r(i))) & "</td>"
            Next
            parts(n) = one & "</tr>"
        Next
        h = h & Join(parts, "") & "</tbody></table>"
    End If
    h = h & "<p>" & SKW_Html(NOTE_END) & "</p></body></html>"
    tmp = Environ$("TEMP") & "\Stationkeeper proof " & Format$(Now, "hhnnss") & Int(Rnd() * 10000) & ".htm"
    SKW_WriteUtf8 tmp, h
    On Error Resume Next
    Set d = Documents.Add(Template:=tmp, Visible:=Not SKW_IsQuiet())
    If d Is Nothing Then
        Err.Clear
        Set d = Documents.Open(FileName:=tmp, AddToRecentFiles:=False, Visible:=Not SKW_IsQuiet())
    End If
    Kill tmp
    On Error GoTo 0
    SKW_KeepReport d, title, "this PC"
    SKW_LastRows = rows.Count
    SKW_FindCol = -1
    Set SKW_FindRows = rows
    SKW_FindDoc = ""
    SKW_FindTitle = title
    SKW_FindPos = 0
    Set SKW_LastReport = d
    On Error Resume Next
    If SKW_IsQuiet() And Not src Is Nothing Then src.Activate
End Sub

' A line in a small file in the temporary folder as each step starts, so that if Word closes
' in the middle of a proof the file says where. It holds step and setting names only.
Private Sub Trace(ByVal s As String, Optional ByVal fresh As Boolean = False)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    If fresh Then
        Open Environ$("TEMP") & "\Stationkeeper proof trace.txt" For Output As #f
    Else
        Open Environ$("TEMP") & "\Stationkeeper proof trace.txt" For Append As #f
    End If
    Print #f, Format$(Now, "hh:nn:ss") & " " & s
    Close #f
End Sub

Private Sub Note(ByVal sect As String, ByVal what As String, ByVal result As String, ByVal detail As String)
    Trace "noted: " & sect & " / " & what & " / " & result
    mRows.Add SKW_Row(sect, what, result, detail)
    If result = "Failed" Then mBad = mBad + 1
End Sub

Private Function SetName(ByVal i As Long) As String
    SetName = Split(SET_NAMES, "|")(i - 1)
End Function

Private Function Num(ByVal v As Double) As String
    Dim s As String
    s = Trim$(Str$(Round(v, 2)))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    Num = s
End Function

Private Function Secs(ByVal t0 As Single) As String
    Dim t As Single
    t = Timer - t0
    If t < 0 Then t = t + 86400
    Secs = Num(t)
End Function

Private Function RxHas(ByVal text As String, ByVal pat As String) As Boolean
    If mRe Is Nothing Then Set mRe = CreateObject("VBScript.RegExp")
    mRe.Global = False
    mRe.IgnoreCase = False
    mRe.Pattern = pat
    RxHas = mRe.Test(text)
End Function

' ---------------------------------------------------------------------
' One setting of a style, read and set through Word. i is 1 to 49, in the order of SET_NAMES.
' A value is text: a number, True or False, a name, or parts joined by commas.
' ---------------------------------------------------------------------
Private Function BorderAt(ByVal i As Long) As Long
    BorderAt = Choose(i - 44, -1, -3, -2, -4)     ' above, below, left, right
End Function

Private Function BorderGap(ByVal st As Object, ByVal i As Long) As Double
    Select Case i
        Case 45: BorderGap = st.ParagraphFormat.Borders.DistanceFromTop
        Case 46: BorderGap = st.ParagraphFormat.Borders.DistanceFromBottom
        Case 47: BorderGap = st.ParagraphFormat.Borders.DistanceFromLeft
        Case 48: BorderGap = st.ParagraphFormat.Borders.DistanceFromRight
    End Select
End Function

Private Function StGet(ByVal st As Object, ByVal i As Long) As String
    Dim s As String, ts As Object, bd As Object
    Select Case i
        Case 1: StGet = CStr(st.Type)
        Case 2: s = st.BaseStyle: StGet = s
        Case 3: s = st.NextParagraphStyle: StGet = s
        Case 4: StGet = CStr(CBool(st.AutomaticallyUpdate))
        Case 5: s = st.LinkStyle: StGet = s
        Case 6: StGet = CStr(CBool(st.Locked))
        Case 7: StGet = CStr(CBool(st.QuickStyle))
        Case 8: StGet = CStr(CBool(st.Visibility))
        Case 9: StGet = CStr(CBool(st.UnhideWhenUsed))
        Case 10: StGet = CStr(st.Priority)
        Case 11: StGet = st.Font.Name
        Case 12
            s = st.Font.Name
            If Left$(s, 1) = "+" Then StGet = s
        Case 13: StGet = Num(st.Font.Size)
        Case 14: StGet = CStr(CBool(st.Font.Bold))
        Case 15: StGet = CStr(CBool(st.Font.Italic))
        Case 16: StGet = CStr(st.Font.Underline)
        Case 17: StGet = CStr(CBool(st.Font.AllCaps))
        Case 18: StGet = CStr(CBool(st.Font.SmallCaps))
        Case 19: StGet = CStr(CBool(st.Font.Hidden))
        Case 20: StGet = CStr(CBool(st.Font.StrikeThrough))
        Case 21: StGet = CStr(CBool(st.Font.DoubleStrikeThrough))
        Case 22
            If CBool(st.Font.Superscript) Then
                StGet = "superscript"
            ElseIf CBool(st.Font.Subscript) Then
                StGet = "subscript"
            End If
        Case 23: StGet = CStr(st.Font.Color)
        Case 24: StGet = Num(st.Font.Spacing)
        Case 25: StGet = Num(st.Font.Kerning)
        Case 26: StGet = CStr(st.Font.Shading.BackgroundPatternColor)
        Case 27: StGet = CStr(CBool(st.NoProofing))
        Case 28: StGet = CStr(st.ParagraphFormat.Alignment)
        Case 29: StGet = CStr(st.ParagraphFormat.OutlineLevel)
        Case 30: StGet = CStr(CBool(st.ParagraphFormat.KeepWithNext))
        Case 31: StGet = CStr(CBool(st.ParagraphFormat.KeepTogether))
        Case 32: StGet = CStr(CBool(st.ParagraphFormat.PageBreakBefore))
        Case 33: StGet = CStr(CBool(st.ParagraphFormat.WidowControl))
        Case 34: StGet = CStr(Not CBool(st.ParagraphFormat.Hyphenation))
        Case 35: StGet = Num(st.ParagraphFormat.SpaceBefore)
        Case 36: StGet = Num(st.ParagraphFormat.SpaceAfter)
        Case 37: StGet = Num(st.ParagraphFormat.LineSpacing)
        Case 38: StGet = CStr(st.ParagraphFormat.LineSpacingRule)
        Case 39: StGet = CStr(CBool(st.NoSpaceBetweenParagraphsOfSameStyle))
        Case 40: StGet = Num(st.ParagraphFormat.LeftIndent)
        Case 41: StGet = Num(st.ParagraphFormat.RightIndent)
        Case 42
            If st.ParagraphFormat.FirstLineIndent > 0 Then StGet = Num(st.ParagraphFormat.FirstLineIndent) Else StGet = "0"
        Case 43
            If st.ParagraphFormat.FirstLineIndent < 0 Then StGet = Num(-st.ParagraphFormat.FirstLineIndent) Else StGet = "0"
        Case 44
            For Each ts In st.ParagraphFormat.TabStops
                If ts.CustomTab Then s = s & IIf(Len(s) > 0, ";", "") & Num(ts.Position) & "," & ts.Alignment & "," & ts.Leader
            Next
            StGet = s
        Case 45 To 48
            Set bd = st.ParagraphFormat.Borders(BorderAt(i))
            If bd.LineStyle <> 0 Then StGet = bd.LineStyle & "," & bd.LineWidth & "," & bd.Color & "," & Num(BorderGap(st, i))
        Case 49: StGet = CStr(st.ParagraphFormat.Shading.BackgroundPatternColor)
    End Select
End Function

Private Sub StSet(ByVal st As Object, ByVal i As Long, ByVal v As String)
    Dim a As Variant, b As Variant, k As Long, bd As Object, pf As Object, rule As Long
    Select Case i
        Case 2: st.BaseStyle = v
        Case 3: st.NextParagraphStyle = v
        Case 4: st.AutomaticallyUpdate = CBool(v)
        ' 5, Linked style, is never set: setting it from a macro closed Word on the build PC.
        Case 6: st.Locked = CBool(v)
        Case 7: st.QuickStyle = CBool(v)
        Case 8: st.Visibility = CBool(v)
        Case 9: st.UnhideWhenUsed = CBool(v)
        Case 10: st.Priority = CLng(v)
        Case 11, 12
            If Len(v) > 0 Then st.Font.Name = v
        Case 13: st.Font.Size = Val(v)
        Case 14: st.Font.Bold = CBool(v)
        Case 15: st.Font.Italic = CBool(v)
        Case 16: st.Font.Underline = CLng(v)
        Case 17: st.Font.AllCaps = CBool(v)
        Case 18: st.Font.SmallCaps = CBool(v)
        Case 19: st.Font.Hidden = CBool(v)
        Case 20: st.Font.StrikeThrough = CBool(v)
        Case 21: st.Font.DoubleStrikeThrough = CBool(v)
        Case 22
            st.Font.Superscript = False
            st.Font.Subscript = False
            If v = "superscript" Then st.Font.Superscript = True
            If v = "subscript" Then st.Font.Subscript = True
        Case 23: st.Font.Color = CLng(v)
        Case 24: st.Font.Spacing = Val(v)
        Case 25: st.Font.Kerning = Val(v)
        Case 26: st.Font.Shading.BackgroundPatternColor = CLng(v)
        Case 27: st.NoProofing = CBool(v)
        Case 28: st.ParagraphFormat.Alignment = CLng(v)
        Case 29: st.ParagraphFormat.OutlineLevel = CLng(v)
        Case 30: st.ParagraphFormat.KeepWithNext = CBool(v)
        Case 31: st.ParagraphFormat.KeepTogether = CBool(v)
        Case 32: st.ParagraphFormat.PageBreakBefore = CBool(v)
        Case 33: st.ParagraphFormat.WidowControl = CBool(v)
        Case 34: st.ParagraphFormat.Hyphenation = Not CBool(v)
        Case 35: st.ParagraphFormat.SpaceBefore = Val(v)
        Case 36: st.ParagraphFormat.SpaceAfter = Val(v)
        Case 37
            ' Word takes a line spacing only under at least, exactly or multiple.
            rule = st.ParagraphFormat.LineSpacingRule
            If rule = 3 Or rule = 4 Or rule = 5 Then st.ParagraphFormat.LineSpacing = Val(v)
        Case 38: st.ParagraphFormat.LineSpacingRule = CLng(v)
        Case 39: st.NoSpaceBetweenParagraphsOfSameStyle = CBool(v)
        Case 40: st.ParagraphFormat.LeftIndent = Val(v)
        Case 41: st.ParagraphFormat.RightIndent = Val(v)
        Case 42
            If Val(v) > 0 Then
                st.ParagraphFormat.FirstLineIndent = Val(v)
            ElseIf st.ParagraphFormat.FirstLineIndent > 0 Then
                st.ParagraphFormat.FirstLineIndent = 0
            End If
        Case 43
            If Val(v) > 0 Then
                st.ParagraphFormat.FirstLineIndent = -Val(v)
            ElseIf st.ParagraphFormat.FirstLineIndent < 0 Then
                st.ParagraphFormat.FirstLineIndent = 0
            End If
        Case 44
            Set pf = st.ParagraphFormat
            pf.TabStops.ClearAll
            If Len(v) > 0 Then
                a = Split(v, ";")
                For k = LBound(a) To UBound(a)
                    b = Split(a(k), ",")
                    pf.TabStops.Add Position:=Val(b(0)), Alignment:=CLng(b(1)), Leader:=CLng(b(2))
                Next
            End If
        Case 45 To 48
            Set bd = st.ParagraphFormat.Borders(BorderAt(i))
            If Len(v) = 0 Then
                bd.LineStyle = 0
            Else
                b = Split(v, ",")
                bd.LineStyle = CLng(b(0))
                bd.LineWidth = CLng(b(1))
                bd.Color = CLng(b(2))
                Select Case i
                    Case 45: st.ParagraphFormat.Borders.DistanceFromTop = Val(b(3))
                    Case 46: st.ParagraphFormat.Borders.DistanceFromBottom = Val(b(3))
                    Case 47: st.ParagraphFormat.Borders.DistanceFromLeft = Val(b(3))
                    Case 48: st.ParagraphFormat.Borders.DistanceFromRight = Val(b(3))
                End Select
            End If
        Case 49: st.ParagraphFormat.Shading.BackgroundPatternColor = CLng(v)
    End Select
End Sub

Private Function TryGet(ByVal st As Object, ByVal i As Long, ByRef v As String) As Boolean
    On Error GoTo No
    v = StGet(st, i)
    TryGet = True
    Exit Function
No:
    v = ""
End Function

Private Function TrySet(ByVal st As Object, ByVal i As Long, ByVal v As String) As String
    On Error GoTo No
    StSet st, i, v
    Exit Function
No:
    TrySet = "Word refused (error " & Err.Number & "). "
End Function

' A paragraph style carries every setting; a character style the first 27, without Next style and Automatically update.
Private Function Applies(ByVal st As Object, ByVal i As Long) As Boolean
    If st.Type = 1 Then
        Applies = True
    Else
        Applies = (i <= 27 And i <> 3 And i <> 4)
    End If
End Function

' The proof's own test values: 1 for the base, 2 for a follower's own value, 3 for the base after it changes.
Private Function TV(ByVal i As Long, ByVal k As Long) As String
    Select Case i
        Case 4, 6, 7, 8, 9, 14, 15, 17, 18, 19, 20, 21, 27, 30, 31, 32, 34, 39: TV = IIf(k = 1, "True", "False")
        Case 33: TV = IIf(k = 1, "False", "True")
        Case 10: TV = CStr(30 + k)
        Case 11: TV = Choose(k, "Courier New", "Arial", "Times New Roman")
        Case 12: TV = Choose(k, "+Headings", "+Body", "+Headings")
        Case 13: TV = Choose(k, "13.5", "17", "21")
        Case 16: TV = Choose(k, "1", "3", "4")
        Case 22: TV = Choose(k, "superscript", "subscript", "")
        Case 23: TV = Choose(k, "5649426", "1193046", "11259375")
        Case 24: TV = Choose(k, "1.5", "2", "-1")
        Case 25: TV = Choose(k, "14", "16", "18")
        Case 26, 49: TV = Choose(k, "13421772", "10079487", "16764057")
        Case 28: TV = Choose(k, "1", "2", "3")
        Case 29: TV = Choose(k, "3", "4", "5")
        Case 35: TV = Choose(k, "7", "11", "15")
        Case 36: TV = Choose(k, "9", "13", "17")
        Case 37: TV = Choose(k, "15", "18", "21")
        Case 38: TV = Choose(k, "4", "3", "5")
        Case 40: TV = Choose(k, "18", "27", "36")
        Case 41: TV = Choose(k, "9", "18", "27")
        Case 42, 43: TV = Choose(k, "12", "18", "24")
        Case 44: TV = Choose(k, "144,2,1", "216,1,0", "288,0,2")
        Case 45 To 48: TV = Choose(k, "1,12,255,1", "1,8,65280,2", "1,18,16711680,3")
    End Select
End Function

' ---------------------------------------------------------------------
' A style's own XML, as Word hands it over. A value written there is set by the style itself.
' ---------------------------------------------------------------------
Private Function XmlText(ByVal s As String) As String
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&gt;", ">")
    s = Replace(s, "&quot;", """")
    s = Replace(s, "&apos;", "'")
    XmlText = Replace(s, "&amp;", "&")
End Function

' Every style in the XML: its name in lower case -> its piece of the XML.
Private Function StyleChunks(ByVal xml As String) As Object
    Dim d As Object, p As Long, q As Long, c As String, a As Long, b As Long, nm As String
    Set d = CreateObject("Scripting.Dictionary")
    Set StyleChunks = d
    p = InStr(1, xml, "<w:style ")
    Do While p > 0
        q = InStr(p, xml, "</w:style>")
        If q = 0 Then Exit Do
        c = Mid$(xml, p, q - p)
        a = InStr(1, c, "<w:name w:val=""")
        If a > 0 Then
            a = a + 15
            b = InStr(a, c, """")
            If b > a Then
                nm = LCase$(XmlText(Mid$(c, a, b - a)))
                If Not d.Exists(nm) Then d(nm) = c
            End If
        End If
        p = InStr(q, xml, "<w:style ")
    Loop
End Function

Private Function Between(ByVal s As String, ByVal a As String, ByVal b As String) As String
    Dim p As Long, q As Long
    p = InStr(1, s, a)
    If p = 0 Then Exit Function
    q = InStr(p, s, b)
    If q = 0 Then Exit Function
    Between = Mid$(s, p, q - p)
End Function

' Is setting i written in this style's own XML?
Private Function Written(ByVal chunk As String, ByVal i As Long) As Boolean
    Dim r As String, p As String
    r = Between(chunk, "<w:rPr>", "</w:rPr>")
    p = Between(chunk, "<w:pPr>", "</w:pPr>")
    Select Case i
        Case 1: Written = True
        Case 2: Written = RxHas(chunk, "<w:basedOn ")
        Case 3: Written = RxHas(chunk, "<w:next ")
        Case 4: Written = RxHas(chunk, "<w:autoRedefine")
        Case 5: Written = RxHas(chunk, "<w:link ")
        Case 6: Written = RxHas(chunk, "<w:locked")
        Case 7: Written = RxHas(chunk, "<w:qFormat")
        Case 8: Written = RxHas(chunk, "<w:semiHidden")
        Case 9: Written = RxHas(chunk, "<w:unhideWhenUsed")
        Case 10: Written = RxHas(chunk, "<w:uiPriority ")
        Case 11: Written = RxHas(r, "<w:rFonts ")
        Case 12: Written = RxHas(r, "w:(ascii|hAnsi)Theme=")
        Case 13: Written = RxHas(r, "<w:sz ")
        Case 14: Written = RxHas(r, "<w:b[ />]")
        Case 15: Written = RxHas(r, "<w:i[ />]")
        Case 16: Written = RxHas(r, "<w:u ")
        Case 17: Written = RxHas(r, "<w:caps[ />]")
        Case 18: Written = RxHas(r, "<w:smallCaps[ />]")
        Case 19: Written = RxHas(r, "<w:vanish[ />]")
        Case 20: Written = RxHas(r, "<w:strike[ />]")
        Case 21: Written = RxHas(r, "<w:dstrike[ />]")
        Case 22: Written = RxHas(r, "<w:vertAlign ")
        Case 23: Written = RxHas(r, "<w:color ")
        Case 24: Written = RxHas(r, "<w:spacing ")
        Case 25: Written = RxHas(r, "<w:kern ")
        Case 26: Written = RxHas(r, "<w:shd ")
        Case 27: Written = RxHas(r, "<w:noProof[ />]")
        Case 28: Written = RxHas(p, "<w:jc ")
        Case 29: Written = RxHas(p, "<w:outlineLvl ")
        Case 30: Written = RxHas(p, "<w:keepNext[ />]")
        Case 31: Written = RxHas(p, "<w:keepLines[ />]")
        Case 32: Written = RxHas(p, "<w:pageBreakBefore[ />]")
        Case 33: Written = RxHas(p, "<w:widowControl[ />]")
        Case 34: Written = RxHas(p, "<w:suppressAutoHyphens[ />]")
        Case 35: Written = RxHas(p, "<w:spacing [^>]*w:before(Lines)?=")
        Case 36: Written = RxHas(p, "<w:spacing [^>]*w:after(Lines)?=")
        Case 37: Written = RxHas(p, "<w:spacing [^>]*w:line=")
        Case 38: Written = RxHas(p, "<w:spacing [^>]*w:lineRule=")
        Case 39: Written = RxHas(p, "<w:contextualSpacing[ />]")
        Case 40: Written = RxHas(p, "<w:ind [^>]*w:(left|start)(Chars)?=")
        Case 41: Written = RxHas(p, "<w:ind [^>]*w:(right|end)(Chars)?=")
        Case 42: Written = RxHas(p, "<w:ind [^>]*w:firstLine(Chars)?=")
        Case 43: Written = RxHas(p, "<w:ind [^>]*w:hanging(Chars)?=")
        Case 44: Written = RxHas(p, "<w:tabs>")
        Case 45: Written = RxHas(Between(p, "<w:pBdr>", "</w:pBdr>"), "<w:top ")
        Case 46: Written = RxHas(Between(p, "<w:pBdr>", "</w:pBdr>"), "<w:bottom ")
        Case 47: Written = RxHas(Between(p, "<w:pBdr>", "</w:pBdr>"), "<w:left ")
        Case 48: Written = RxHas(Between(p, "<w:pBdr>", "</w:pBdr>"), "<w:right ")
        Case 49: Written = RxHas(p, "<w:shd ")
    End Select
End Function

Private Function DocXml(ByVal d As Object) As String
    On Error Resume Next
    DocXml = d.Range(0, 0).WordOpenXML
End Function

Private Function CountOf(ByVal s As String, ByVal what As String) As Long
    Dim p As Long
    p = InStr(1, s, what)
    Do While p > 0
        CountOf = CountOf + 1
        p = InStr(p + Len(what), s, what)
    Loop
End Function

' =====================================================================
' Proof 1: this PC, and what Word lets a macro do to a style.
' =====================================================================
Private Function TryObject(ByVal progId As String) As String
    Dim o As Object
    On Error GoTo No
    Set o = CreateObject(progId)
    Set o = Nothing
    Exit Function
No:
    TryObject = "error " & Err.Number
End Function

Private Sub ProbePC()
    Dim v As Variant, e As String, t0 As Single, xl As Object, s As String, p As String, f As Integer, n As Long, ok As Boolean
    ' Word
    s = "Word " & Application.Version & ", build " & Application.Build
    #If Win64 Then
        s = s & ", 64-bit"
    #Else
        s = s & ", 32-bit"
    #End If
    Note "P1 This PC", "Word", "Seen", s
    ' The helpers the tools are made of
    For Each v In Array("Scripting.Dictionary", "Scripting.FileSystemObject", "ADODB.Stream", "VBScript.RegExp")
        e = TryObject(CStr(v))
        Note "P1 This PC", "Windows helper " & v, IIf(Len(e) = 0, "Works", "Failed"), e
    Next
    ' Excel driven from Word
    t0 = Timer
    On Error Resume Next
    Err.Clear
    Set xl = CreateObject("Excel.Application")
    If Err.Number <> 0 Or xl Is Nothing Then
        Note "P1 This PC", "Excel started from Word", "Failed", "error " & Err.Number
    Else
        s = "Excel " & xl.Version & ", started in " & Secs(t0) & " s"
        xl.Quit
        Set xl = Nothing
        Note "P1 This PC", "Excel started from Word", "Works", s
    End If
    ' Settings kept per PC
    Err.Clear
    SaveSetting SKW_APP, "Proof", "probe", "1"
    ok = (GetSetting(SKW_APP, "Proof", "probe", "") = "1") And Err.Number = 0
    DeleteSetting SKW_APP, "Proof"
    Note "P1 This PC", "Settings kept on this PC (the registry)", IIf(ok, "Works", "Failed"), ""
    ' Files the tools write
    Err.Clear
    p = Environ$("TEMP") & "\Stationkeeper proof probe.txt"
    SKW_WriteUtf8 p, "proof"
    ok = (Err.Number = 0) And Len(Dir$(p)) > 0
    Kill p
    Note "P1 This PC", "A file written to the temporary folder", IIf(ok, "Works", "Failed"), ""
    Err.Clear
    p = Environ$("LOCALAPPDATA") & "\Stationkeeper Macros"
    If Len(Dir$(p, vbDirectory)) = 0 Then MkDir p
    p = p & "\proof probe.txt"
    Err.Clear
    f = FreeFile
    Open p For Output As #f
    Print #f, "proof"
    Close #f
    ok = (Err.Number = 0) And Len(Dir$(p)) > 0
    Kill p
    Note "P1 This PC", "A file written to this PC's own app folder (workbook backups go there)", IIf(ok, "Works", "Failed"), ""
    ' A window
    Err.Clear
    n = -1
    Load frmSKWBrand
    n = frmSKWBrand.Controls.Count
    Unload frmSKWBrand
    Note "P1 This PC", "A Stationkeeper window can be made", IIf(Err.Number = 0 And n >= 0, "Works", "Failed"), IIf(Err.Number <> 0, "error " & Err.Number, "")
    ' The VBA project (the admin reads macros through it; the menu and the tools do not need it)
    Err.Clear
    n = -1
    n = ThisDocument.VBProject.VBComponents.Count
    Note "P1 This PC", "Trust access to the VBA project object model", IIf(Err.Number = 0 And n >= 0, "On", "Off"), "Off is expected on a work PC. The menu and the tools run without it."
    ' Synced folders
    Err.Clear
    n = SKW_SyncRootCount()
    Note "P1 This PC", "Synced OneDrive or SharePoint folders this PC knows", "Seen", n & " found"
    s = LCase$(ThisDocument.FullName)
    Note "P1 This PC", "Where the template with the tools sits", "Seen", IIf(Left$(s, 4) = "http", "Opened from a web address", IIf(SKW_InSyncRoot(ThisDocument.FullName), "In a synced folder", "In a folder on this PC")) & _
        ", a " & Mid$(s, InStrRev(s, ".")) & " file"
    ' The person at the PC, for the Stationkeeper Admins list (yes or no only)
    Note "P1 This PC", "The PC has a Windows user name and a Word user name", IIf(Len(Environ$("USERNAME")) > 0 And Len(Application.UserName) > 0, "Yes", "No"), _
        IIf(StrComp(Environ$("USERNAME"), Application.UserName, vbTextCompare) = 0, "The two are the same.", "The two differ: either may be on the admin list.")
    ' Spelling
    Err.Clear
    ok = Options.IgnoreUppercase
    Note "P1 This PC", "Word skips words in capitals when it checks spelling", IIf(Err.Number <> 0, "Failed", IIf(ok, "Yes", "No")), "Acronyms then need no dictionary."
    Err.Clear
    n = CustomDictionaries.Count
    Note "P1 This PC", "Custom dictionaries in Word's list", IIf(Err.Number <> 0, "Failed", "Seen"), n & " listed"
    Note "P1 This PC", "A standards workbook is chosen for this template", IIf(Len(SKW_StandardsPath()) > 0, "Yes", "No"), ""
    On Error GoTo 0
End Sub

' P4: can each setting be set and read back, on a paragraph style and a character style?
' The link between a paragraph style and its character half is read and never set: on the build PC,
' setting it from a macro closed Word. Word links a new paragraph style itself.
Private Sub ProbeSetEach(ByVal d As Object)
    Dim sp As Object, sc As Object, i As Long, k As Long, st As Object, want As String, got As String, e As String
    Dim res As String, bad As Long, kindName As String, chunk As String, linkNote As String, isLinked As Boolean
    Trace "set each: making the scratch styles"
    d.Styles.Add "ZZ Proof Base", 1
    d.Styles.Add "ZZ Proof Char Base", 2
    Set sp = d.Styles.Add("ZZ Proof Para", 1)
    Set sc = d.Styles.Add("ZZ Proof Char", 2)
    Trace "set each: reading the link of a new paragraph style"
    On Error Resume Next
    isLinked = CBool(sp.Linked)
    got = ""
    got = sp.LinkStyle
    On Error GoTo 0
    linkNote = "Read only; the proof never sets it (setting it from a macro closed Word on the build PC). A new paragraph style is " & _
        IIf(isLinked, "linked by Word itself", "not linked by Word") & IIf(Len(got) > 0, ", and its link could be read.", ", and its link could not be read.")
    For i = 2 To N_SET
        res = ""
        bad = 0
        For k = 1 To 2
            Set st = Choose(k, sp, sc)
            kindName = Choose(k, "paragraph", "character")
            If Applies(st, i) And Not (i = 5) Then
                want = TV(i, 1)
                If i = 2 Then want = IIf(k = 2, "ZZ Proof Char Base", "ZZ Proof Base")
                If i = 3 Then want = "ZZ Proof Base"
                If i = 37 Then TrySet st, 38, "4"
                Trace "set each: " & SetName(i) & " on a " & kindName & " style"
                e = TrySet(st, i, want)
                got = ""
                If Len(e) = 0 Then
                    If Not TryGet(st, i, got) Then e = "Set, but could not be read back. "
                End If
                If Len(e) = 0 And i = 12 Then
                    ' A theme font: Word may hand back the font's own name. The style's XML says which it is.
                    chunk = ""
                    If StyleChunks(DocXml(d)).Exists(LCase$(st.NameLocal)) Then chunk = StyleChunks(DocXml(d))(LCase$(st.NameLocal))
                    res = res & kindName & ": " & IIf(Left$(got, 1) = "+", "reads back as a theme font", "reads back as a font name") & _
                        IIf(Written(chunk, 12), ", written as a theme font in the XML", ", not written as a theme font in the XML") & ". "
                ElseIf Len(e) = 0 And got = want Then
                    res = res & kindName & ": set and read back. "
                Else
                    bad = bad + 1
                    res = res & kindName & ": " & IIf(Len(e) > 0, e, "set " & want & ", read back " & got & ". ")
                End If
            End If
        Next
        If i = 5 Then
            Note "P4 Set and read each value", SetName(i), "Not set", linkNote
        Else
            Note "P4 Set and read each value", SetName(i), IIf(bad = 0, "Works", "See detail"), Trim$(res)
        End If
    Next
End Sub

' P2 and P3: a follower is given its own value, then set to the same value as its base. Does it follow the base again
' when the base changes? And is the value still written in the follower's own XML?
Private Sub ProbeFollow(ByVal d As Object, ByRef nFollow As Long, ByRef nStay As Long, ByRef nOther As Long, ByRef nWrit As Long)
    Dim a As Object, b As Object, c As Object, i As Long, e As String, vb As String, vc As String, res As String, det As String, chunks As Object, wr As Boolean
    Set a = d.Styles("ZZ Proof Base")
    Set b = d.Styles.Add("ZZ Proof Follower", 1)
    b.BaseStyle = "ZZ Proof Base"
    Set c = d.Styles.Add("ZZ Proof Control", 1)
    c.BaseStyle = "ZZ Proof Base"
    For i = 11 To N_SET
        If i <> 12 Then
            If i = 37 Then TrySet a, 38, "4"
            Trace "back to base: " & SetName(i)
            e = TrySet(a, i, TV(i, 1))
            e = e & TrySet(b, i, TV(i, 2))
            e = e & TrySet(b, i, TV(i, 1))
            Set chunks = StyleChunks(DocXml(d))
            wr = False
            If chunks.Exists("zz proof follower") Then wr = Written(chunks("zz proof follower"), i)
            e = e & TrySet(a, i, TV(i, 3))
            vb = "": vc = ""
            TryGet b, i, vb
            TryGet c, i, vc
            det = IIf(wr, "Still written in the follower's XML.", "No longer written in the follower's XML.")
            If wr Then nWrit = nWrit + 1
            If Len(e) > 0 Then
                res = "Could not be tried": det = e
                nOther = nOther + 1
            ElseIf vc <> TV(i, 3) Then
                res = "Not passed down": det = "A follower that was never touched did not take the base's new value (read " & vc & ", base set to " & TV(i, 3) & ")."
                nOther = nOther + 1
            ElseIf vb = TV(i, 3) Then
                res = "Follows the base again": nFollow = nFollow + 1
            ElseIf vb = TV(i, 1) Then
                res = "Stays as set": nStay = nStay + 1
            Else
                res = "Other": det = det & " Read " & vb & " after the base went from " & TV(i, 1) & " to " & TV(i, 3) & "."
                nOther = nOther + 1
            End If
            Note "P2 Set a value back to its base", SetName(i), res, det
        End If
    Next
End Sub

' Other ways to clear a value, on five settings: Reset, and taking the base away and giving it back.
Private Sub ProbeClearWays(ByVal d As Object)
    Dim a As Object, b As Object, v As Variant, i As Long, m As Long, e As String, vb As String, res As String, way As String
    Set a = d.Styles("ZZ Proof Base")
    Set b = d.Styles("ZZ Proof Follower")
    For m = 1 To 2
        way = Choose(m, "Reset on the style's font or paragraph format", "Base style taken away and given back")
        For Each v In Array(13, 14, 28, 36, 40)
            i = CLng(v)
            Trace "clear way " & m & ": " & SetName(i)
            e = TrySet(a, i, TV(i, 1)) & TrySet(b, i, TV(i, 2))
            On Error Resume Next
            Err.Clear
            If m = 1 Then
                If i <= 27 Then b.Font.Reset Else b.ParagraphFormat.Reset
            Else
                b.BaseStyle = "Normal"
                b.BaseStyle = "ZZ Proof Base"
            End If
            If Err.Number <> 0 Then e = e & "Word refused (error " & Err.Number & "). "
            On Error GoTo 0
            e = e & TrySet(a, i, TV(i, 3))
            vb = ""
            TryGet b, i, vb
            If Len(e) > 0 Then
                res = "Word refuses it"
            ElseIf vb = TV(i, 3) Then
                res = "Clears it: the follower takes the base's value"
            ElseIf vb = TV(i, 2) Then
                res = "Does not clear it"
            Else
                res = "Other"
                e = "Read " & vb & "; own value " & TV(i, 2) & ", base now " & TV(i, 3) & "."
            End If
            Note "P2 Other ways to clear", way & ": " & SetName(i), res, e
        Next
    Next
End Sub

Private Function PkgXml(ByVal followerExtra As String) As String
    Const W As String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    Const REL As String = "http://schemas.openxmlformats.org/package/2006/relationships"
    Const OD As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    Dim s As String
    s = "<pkg:package xmlns:pkg=""http://schemas.microsoft.com/office/2006/xmlPackage"">"
    s = s & "<pkg:part pkg:name=""/_rels/.rels"" pkg:contentType=""application/vnd.openxmlformats-package.relationships+xml""><pkg:xmlData>" & _
        "<Relationships xmlns=""" & REL & """><Relationship Id=""rId1"" Type=""" & OD & "/officeDocument"" Target=""word/document.xml""/></Relationships></pkg:xmlData></pkg:part>"
    s = s & "<pkg:part pkg:name=""/word/_rels/document.xml.rels"" pkg:contentType=""application/vnd.openxmlformats-package.relationships+xml""><pkg:xmlData>" & _
        "<Relationships xmlns=""" & REL & """><Relationship Id=""rId1"" Type=""" & OD & "/styles"" Target=""styles.xml""/></Relationships></pkg:xmlData></pkg:part>"
    s = s & "<pkg:part pkg:name=""/word/document.xml"" pkg:contentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml""><pkg:xmlData>" & _
        "<w:document xmlns:w=""" & W & """><w:body><w:p><w:pPr><w:pStyle w:val=""ZZXmlBase""/></w:pPr><w:r><w:t>Proof</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val=""ZZXmlFollower""/></w:pPr><w:r><w:t>Proof</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part>"
    s = s & "<pkg:part pkg:name=""/word/styles.xml"" pkg:contentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml""><pkg:xmlData>" & _
        "<w:styles xmlns:w=""" & W & """>" & _
        "<w:style w:type=""paragraph"" w:default=""1"" w:styleId=""Normal""><w:name w:val=""Normal""/></w:style>" & _
        "<w:style w:type=""paragraph"" w:customStyle=""1"" w:styleId=""ZZXmlBase""><w:name w:val=""ZZ Xml Base""/><w:basedOn w:val=""Normal""/><w:rPr><w:b/><w:sz w:val=""28""/></w:rPr></w:style>" & _
        "<w:style w:type=""paragraph"" w:customStyle=""1"" w:styleId=""ZZXmlFollower""><w:name w:val=""ZZ Xml Follower""/><w:basedOn w:val=""ZZXmlBase""/><w:rPr><w:i/>" & followerExtra & "</w:rPr></w:style>" & _
        "</w:styles></pkg:xmlData></pkg:part></pkg:package>"
    PkgXml = s
End Function

Private Function PutXml(ByVal d As Object, ByVal xml As String) As String
    Dim r As Object
    On Error GoTo No
    Set r = d.Content
    r.Collapse 0
    r.InsertXML xml
    Exit Function
No:
    PutXml = "Word refused the XML (error " & Err.Number & "). "
End Function

' The other road for Build: hand Word a style as XML, holding only what the style itself sets.
Private Sub ProbeXml(ByVal d As Object)
    Dim xml As String, xmlAll As String, nA As Long, nB As Long, nUse As Long, st As Object, e As String, xf As Object, xb As Object, s As String, sz As String
    Const XPART As String = "P2 A style handed over as XML"
    Trace "xml: reading"
    xml = DocXml(d)
    On Error Resume Next
    xmlAll = d.Content.WordOpenXML
    For Each st In d.Styles
        If st.InUse Then nUse = nUse + 1
    Next
    On Error GoTo 0
    nA = CountOf(xml, "<w:style ")
    nB = CountOf(xmlAll, "<w:style ")
    Note "P3 Read what a style itself sets", "Word hands over the styles' own XML", IIf(nA > 0, "Works", "Failed"), _
        nA & " styles in the XML of an empty spot, " & nB & " in the XML of the whole text, " & nUse & " styles in use in the document."
    ' A new style from XML
    Trace "xml: a new style"
    e = PutXml(d, PkgXml(""))
    On Error Resume Next
    Set xb = d.Styles("ZZ Xml Base")
    Set xf = d.Styles("ZZ Xml Follower")
    On Error GoTo 0
    If Len(e) > 0 Or xb Is Nothing Or xf Is Nothing Then
        s = ""
        On Error Resume Next
        If Not xf Is Nothing Then s = xf.BaseStyle
        On Error GoTo 0
        Note XPART, "A new style arrives from XML", "Did not arrive whole", e & "The base style " & IIf(xb Is Nothing, "is not", "is") & " in the document afterwards; the follower " & _
            IIf(xf Is Nothing, "is not.", "is" & IIf(Len(s) > 0, ", based on a style.", ", with no base style."))
        Exit Sub
    End If
    s = "Bold " & CBool(xf.Font.Bold) & ", size " & Num(xf.Font.Size) & ", italic " & CBool(xf.Font.Italic) & " (the XML gave the follower italic only; bold and 14 come from its base)."
    Note XPART, "A new style arrives from XML", IIf(CBool(xf.Font.Bold) And xf.Font.Size = 14 And CBool(xf.Font.Italic), "Works", "Other"), s
    Trace "xml: the base changes"
    xb.Font.Size = 20
    Note XPART, "A style that came from XML follows its base", IIf(xf.Font.Size = 20, "Follows the base", "Does not follow"), "Base set to 20, follower reads " & Num(xf.Font.Size) & "."
    ' The same style handed over again with a value of its own
    Trace "xml: handed over again"
    e = PutXml(d, PkgXml("<w:sz w:val=""44""/>"))
    sz = Num(xf.Font.Size)
    Note XPART, "A style already in the document, handed over again with a change", IIf(Len(e) > 0, "Word refused it", IIf(sz = "22", "Takes the change", "Left as it was")), e & "Follower reads " & sz & " (the XML said 22)."
    ' Deleted first, then handed over
    Trace "xml: deleted, then handed over"
    On Error Resume Next
    Err.Clear
    xf.Delete
    e = ""
    If Err.Number <> 0 Then e = "The style could not be deleted (error " & Err.Number & "). "
    On Error GoTo 0
    e = e & PutXml(d, PkgXml("<w:sz w:val=""44""/>"))
    Set xf = Nothing
    On Error Resume Next
    Set xf = d.Styles("ZZ Xml Follower")
    sz = ""
    sz = Num(xf.Font.Size)
    On Error GoTo 0
    Note XPART, "A style deleted, then handed over again with a change", IIf(Len(e) > 0 Or xf Is Nothing, "Did not arrive", IIf(sz = "22", "Takes the change", "Other")), e & "Follower reads " & sz & " (the XML said 22)."
End Sub

Public Function SKW_ProofPC() As String
    Dim d As Object, nFollow As Long, nStay As Long, nOther As Long, nWrit As Long, t0 As Single, msg As String, stage As String, src As Object, wasScreen As Boolean
    Set SKW_LastReport = Nothing
    wasScreen = True
    Set mRows = New Collection
    mBad = 0
    t0 = Timer
    On Error Resume Next
    If Documents.Count > 0 Then Set src = ActiveDocument
    On Error GoTo Bad
    Trace "proof 1 starts", True
    SKW_ProgressTitle "Work PC proof 1"
    stage = "this PC"
    SKW_Status "Looking at this PC..."
    ProbePC
    Trace "the scratch document"
    stage = "the scratch document"
    SKW_Status "Trying each style setting in a scratch document..."
    ' The scratch document is in front while it is worked on: Word looks up a base style or a next style by name in the document in front.
    wasScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set d = Documents.Add
    stage = "setting each value"
    ProbeSetEach d
    stage = "setting a value back to its base"
    SKW_Status "Setting values back to the base style..."
    ProbeFollow d, nFollow, nStay, nOther, nWrit
    stage = "other ways to clear"
    ProbeClearWays d
    stage = "styles as XML"
    SKW_Status "Handing styles over as XML..."
    ProbeXml d
    d.Close SaveChanges:=0
    Set d = Nothing
    If Not src Is Nothing Then src.Activate
    Application.ScreenUpdating = wasScreen
    SKW_ProgressDone
    msg = "Proof 1 ran in " & Secs(t0) & " s. " & mBad & " check(s) of this PC failed. Set equal to its base, a value follows the base again for " & nFollow & _
        " setting(s), stays as set for " & nStay & ", and " & nOther & " did something else; " & nWrit & " are still written in the follower's XML."
    ProofReport "Work PC Proof 1: This PC and Word's Styles", msg & " Everything was tried in a scratch document that was not saved. " & _
        "The values named below are the proof's own test values, not yours. Changes nothing.", Array("Proof", "What was tried", "Result", "Detail"), mRows
    SKW_ProofPC = msg
    Exit Function
Bad:
    msg = "Proof 1 stopped at " & stage & ": error " & Err.Number & " " & Err.Description
    On Error Resume Next
    If Not d Is Nothing Then d.Close SaveChanges:=0
    If Not src Is Nothing Then src.Activate
    Application.ScreenUpdating = wasScreen
    SKW_ProgressDone
    Note "Stopped", stage, "Failed", msg
    ProofReport "Work PC Proof 1: This PC and Word's Styles", msg, Array("Proof", "What was tried", "Result", "Detail"), mRows
    SKW_ProofPC = msg
End Function

' Finds out what this PC allows: Excel started from Word, settings and files kept on the PC, a Stationkeeper window, synced folders. Then, in a scratch document that is not saved, tries each of the 49 style settings and whether a value can be given back to the base style. Changes nothing of yours.
Public Sub ProofThisPC()
    SKW_Show SKW_ProofPC()
End Sub

' =====================================================================
' Proof 2: the styles of the document in front.
' =====================================================================
Public Function SKW_ProofDocStyles() As String
    Dim d As Object, t As Object, st As Object, bs As Object, ts As Object, wasSaved As Boolean, chunks As Object, chunk As String, hasXml As Boolean
    Dim i As Long, v As String, vbv As String, nm As String, baseNm As String, rows As New Collection, msg As String, stage As String, e As String
    Dim nAp(1 To 49) As Long, nFail(1 To 49) As Long, nDiff(1 To 49) As Long, nWr(1 To 49) As Long, nWrSame(1 To 49) As Long, nDiffNot(1 To 49) As Long
    Dim nBack(1 To 49) As Long, nSetFail(1 To 49) As Long
    Dim stList As Collection, done As Object, inList As Object, pass As Long, restList As Collection, itm As Variant
    Dim nPara As Long, nChar As Long, nSkip As Long, nOtherKind As Long, nNoXml As Long, nNoBase As Long, nMake As Long, tRead As String, tBuild As String, tComp As String, t0 As Single
    Dim nSetsMade As Long, anyBack As Long, wr As Boolean, hasBase As Boolean
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Work PC proof 2", False)
    If Len(msg) > 0 Then SKW_ProofDocStyles = msg: Exit Function
    On Error GoTo Bad
    Set d = ActiveDocument
    wasSaved = d.Saved
    Trace "proof 2 starts", True
    stage = "reading the styles"
    Set chunks = StyleChunks(DocXml(d))
    Set stList = New Collection
    Set inList = CreateObject("Scripting.Dictionary")
    ' The styles the file itself holds: Word lists every built-in style, and hands over the XML of the ones kept in the file.
    For Each st In d.Styles
        If st.Type = 1 Or st.Type = 2 Then
            If chunks.Exists(LCase$(st.NameLocal)) Then
                stList.Add st
                inList(LCase$(st.NameLocal)) = True
                If st.Type = 1 Then nPara = nPara + 1 Else nChar = nChar + 1
            ElseIf st.BuiltIn Then
                nSkip = nSkip + 1
            Else
                nNoXml = nNoXml + 1
            End If
        Else
            nOtherKind = nOtherKind + 1
        End If
    Next
    t0 = Timer
    SKW_Steps stList.Count, "Reading style"
    For Each st In stList
        SKW_Step
        nm = LCase$(st.NameLocal)
        hasXml = True
        chunk = chunks(nm)
        baseNm = ""
        TryGet st, 2, baseNm
        Set bs = Nothing
        On Error Resume Next
        If Len(baseNm) > 0 Then Set bs = d.Styles(baseNm)
        On Error GoTo Bad
        hasBase = Not bs Is Nothing
        If Not hasBase Then nNoBase = nNoBase + 1
        For i = 1 To N_SET
            If Applies(st, i) Then
                nAp(i) = nAp(i) + 1
                If Not TryGet(st, i, v) Then
                    nFail(i) = nFail(i) + 1
                Else
                    wr = False
                    If hasXml Then wr = Written(chunk, i)
                    If wr Then nWr(i) = nWr(i) + 1
                    If hasBase And i > 10 Then
                        vbv = ""
                        If TryGet(bs, i, vbv) Then
                            If v <> vbv Then
                                nDiff(i) = nDiff(i) + 1
                                If hasXml And Not wr Then nDiffNot(i) = nDiffNot(i) + 1
                            ElseIf wr Then
                                nWrSame(i) = nWrSame(i) + 1
                            End If
                        End If
                    End If
                End If
            End If
        Next
    Next
    tRead = Secs(t0)
    ' The round trip: each style made again in a scratch document from its values, then compared.
    Trace "proof 2: reading done, building"
    stage = "building the styles again"
    t0 = Timer
    ' In front while it is built: Word looks up a base style or a next style by name in the document in front.
    Set t = Documents.Add
    Set done = CreateObject("Scripting.Dictionary")
    ' A base style before the styles that follow it.
    For pass = 1 To 25
        Set restList = New Collection
        For Each st In stList
            baseNm = ""
            TryGet st, 2, baseNm
            If Len(baseNm) = 0 Or done.Exists(LCase$(baseNm)) Or Not inList.Exists(LCase$(baseNm)) Or pass = 25 Then
                done(LCase$(st.NameLocal)) = st.NameLocal
                SKW_Pulse
                Set ts = Nothing
                On Error Resume Next
                Set ts = t.Styles(st.NameLocal)
                If ts Is Nothing Then
                    Err.Clear
                    Set ts = t.Styles.Add(st.NameLocal, st.Type)
                End If
                On Error GoTo Bad
                If ts Is Nothing Then
                    nMake = nMake + 1
                Else
                    If Len(baseNm) > 0 Then
                        v = ""
                        TryGet ts, 2, v
                        If v <> baseNm Then
                            If Len(TrySet(ts, 2, baseNm)) > 0 Then nSetFail(2) = nSetFail(2) + 1
                        End If
                    End If
                    For Each itm In Array(4, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 38, 37, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49)
                        i = CLng(itm)
                        If Applies(st, i) Then
                            If TryGet(st, i, v) Then
                                vbv = ""
                                TryGet ts, i, vbv
                                If v <> vbv Then
                                    nSetsMade = nSetsMade + 1
                                    If Len(TrySet(ts, i, v)) > 0 Then nSetFail(i) = nSetFail(i) + 1
                                End If
                            End If
                        End If
                    Next
                End If
            Else
                restList.Add st
            End If
        Next
        If restList.Count = 0 Then Exit For
        Set stList = restList
    Next
    tBuild = Secs(t0)
    ' Next style and Linked style need every style to be there first.
    Trace "proof 2: building done, comparing"
    stage = "comparing"
    t0 = Timer
    For Each itm In done.Items
        Set st = Nothing: Set ts = Nothing
        On Error Resume Next
        Set st = d.Styles(CStr(itm))
        Set ts = t.Styles(CStr(itm))
        On Error GoTo Bad
        If Not st Is Nothing And Not ts Is Nothing Then
            If st.Type = 1 Then
                If TryGet(st, 3, v) Then
                    vbv = ""
                    TryGet ts, 3, vbv
                    If v <> vbv Then
                        If Len(TrySet(ts, 3, v)) > 0 Then nSetFail(3) = nSetFail(3) + 1
                    End If
                End If
            End If
        End If
    Next
    For Each itm In done.Items
        Set st = Nothing: Set ts = Nothing
        On Error Resume Next
        Set st = d.Styles(CStr(itm))
        Set ts = t.Styles(CStr(itm))
        On Error GoTo Bad
        If Not st Is Nothing And Not ts Is Nothing Then
            SKW_Pulse
            For i = 1 To N_SET
                If Applies(st, i) And i <> 5 And i <> 12 Then
                    If TryGet(st, i, v) Then
                        vbv = ""
                        TryGet ts, i, vbv
                        If v <> vbv Then nBack(i) = nBack(i) + 1: anyBack = anyBack + 1
                    End If
                End If
            Next
        End If
    Next
    tComp = Secs(t0)
    t.Close SaveChanges:=0
    Set t = Nothing
    d.Activate
    If wasSaved And Not d.Saved Then d.Saved = True
    SKW_Finish
    For i = 1 To N_SET
        rows.Add SKW_Row(SetName(i), nAp(i), nFail(i), IIf(i > 10, CStr(nDiff(i)), "-"), nWr(i), IIf(i > 10, CStr(nWrSame(i)), "-"), IIf(i > 10, CStr(nDiffNot(i)), "-"), _
            IIf(i = 5 Or i = 12, "-", CStr(nBack(i))), nSetFail(i))
    Next
    msg = (nPara + nChar) & " style(s) kept in the file were read: " & nPara & " paragraph, " & nChar & " character. Left out: " & nSkip & _
        " built-in styles the file does not hold, " & nOtherKind & " table and list styles, and " & nNoXml & " custom style(s) Word listed but did not hand over as XML. " & nNoBase & " have no base style. " & _
        "Reading took " & tRead & " s, building them again " & tBuild & " s (" & nSetsMade & " values set, " & nMake & " style(s) could not be made), comparing " & tComp & " s. " & _
        anyBack & " value(s) came back different after the round trip."
    ProofReport "Work PC Proof 2: This Document's Styles", msg & " The styles were built again in a scratch document that was not saved. Your document was not changed.", _
        Array("Setting", "Styles it applies to", "Could not be read", "Differ from the base style", "Written in the style's own XML", _
              "Written, but the same as the base", "Differ, but not written", "Came back different", "Could not be set"), rows
    SKW_ProofDocStyles = msg
    Exit Function
Bad:
    msg = "Proof 2 stopped at " & stage & ": error " & Err.Number & " " & Err.Description & ". Your document was not changed."
    On Error Resume Next
    If Not t Is Nothing Then t.Close SaveChanges:=0
    d.Activate
    If wasSaved And Not d.Saved Then d.Saved = True
    SKW_Finish
    SKW_ProofDocStyles = msg
End Function

' Reads every paragraph and character style of the document in front: which of the 49 settings Word hands over, which differ from the base style, and which the style's own XML holds. Then makes the styles again in a scratch document and counts what came back different. The report holds setting names and counts only. Your document is not changed.
Public Sub ProofThisDocumentStyles()
    SKW_Show SKW_ProofDocStyles()
End Sub

' =====================================================================
' Proof 3: a workbook. How long it takes to read, and what is still there after a save.
' =====================================================================
Private Function TabLabel(ByVal nm As String, ByRef nOwn As Long) As String
    If InStr(1, KNOWN_TABS, "|" & LCase$(Trim$(nm)) & "|") > 0 Then
        TabLabel = Trim$(nm)
    Else
        nOwn = nOwn + 1
        TabLabel = "A tab of your own (" & nOwn & ")"
    End If
End Function

' One line of numbers per tab: what a save must not lose. firstHdr: each tab's first column name.
Private Function BookPrint(ByVal path As String, ByRef firstHdr As Object, ByRef why As String) As Object
    Dim xl As Object, wb As Object, ws As Object, res As Object, c As Long, nC As Long, nR As Long, hid As Long, lvl As Long, nVal As Long, hdrs As Long, grp As Long, nOwn As Long, lab As String
    Set res = CreateObject("Scripting.Dictionary")
    Set firstHdr = CreateObject("Scripting.Dictionary")
    On Error GoTo No
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(FileName:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    For Each ws In wb.Worksheets
        nR = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
        nC = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
        hid = 0: lvl = 1: hdrs = 0: grp = 0
        For c = 1 To nC
            If ws.Columns(c).Hidden Then hid = hid + 1
            If ws.Columns(c).OutlineLevel > 1 Then grp = grp + 1
            If ws.Columns(c).OutlineLevel > lvl Then lvl = ws.Columns(c).OutlineLevel
            If Len(Trim$(CStr(ws.Cells(1, c).Text))) > 0 Then hdrs = hdrs + 1
        Next
        nVal = 0
        On Error Resume Next
        nVal = ws.Cells.SpecialCells(-4174).Count
        On Error GoTo No
        lab = TabLabel(ws.Name, nOwn)
        res(lab) = "rows " & nR & ", columns " & nC & ", column names " & hdrs & ", hidden columns " & hid & ", folded columns " & grp & _
            ", cells with a drop-down " & nVal & ", tab " & IIf(ws.Visible = -1, "shown", "hidden")
        firstHdr(lab) = ws.Name & vbTab & Trim$(CStr(ws.Cells(1, 1).Text))
    Next
    res("(the workbook)") = "tabs " & wb.Worksheets.Count & ", names " & wb.Names.Count
    wb.Close False
    xl.Quit
    Set BookPrint = res
    Exit Function
No:
    why = "error " & Err.Number & " " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Set BookPrint = res
End Function

' The whole workbook read as the tools read it: every tab's used cells in one go.
Private Function TimeRead(ByVal path As String, ByRef nCells As Long, ByRef why As String) As String
    Dim xl As Object, wb As Object, ws As Object, v As Variant, t0 As Single
    On Error GoTo No
    t0 = Timer
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(FileName:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    For Each ws In wb.Worksheets
        v = ws.UsedRange.Value
        If IsArray(v) Then nCells = nCells + (UBound(v, 1) - LBound(v, 1) + 1) * (UBound(v, 2) - LBound(v, 2) + 1)
    Next
    wb.Close False
    xl.Quit
    TimeRead = Secs(t0)
    Exit Function
No:
    why = "error " & Err.Number & " " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
End Function

Public Function SKW_ProofWorkbook(ByVal path As String) As String
    Dim rows As New Collection, why As String, nCells As Long, secs1 As String, copyPath As String, before As Object, after As Object, hdr As Object, hdr2 As Object
    Dim k As Variant, sheetNm As String, keyHdr As String, parts As Variant, row As Object, wrows As New Collection, a As Long, c As Long, s As Long, e As String
    Dim nSame As Long, nDiff As Long, msg As String, dirp As String, kept As String
    Set SKW_LastReport = Nothing
    path = SKW_LocalPath(path)
    If LCase$(Left$(path, 4)) = "http" Then SKW_ProofWorkbook = "That workbook is kept in OneDrive or SharePoint, and its copy on this PC was not found in a synced folder. Nothing was done.": Exit Function
    If Len(Dir$(path)) = 0 Then SKW_ProofWorkbook = "The workbook was not found. Nothing was done.": Exit Function
    SKW_ProgressTitle "Work PC proof 3"
    SKW_Status "Reading the workbook..."
    secs1 = TimeRead(path, nCells, why)
    If Len(why) > 0 Then
        SKW_ProgressDone
        SKW_ProofWorkbook = "The workbook could not be read (" & why & "). Nothing was done."
        Exit Function
    End If
    ' A copy in the temporary folder: the workbook itself is only read.
    SKW_Status "Saving a copy and looking at what is still there..."
    dirp = Environ$("TEMP") & "\Stationkeeper proof"
    On Error Resume Next
    If Len(Dir$(dirp, vbDirectory)) = 0 Then MkDir dirp
    copyPath = dirp & "\proof copy " & Format$(Now, "hhnnss") & Mid$(path, InStrRev(path, "."))
    Err.Clear
    FileCopy path, copyPath
    If Err.Number <> 0 Then
        SKW_ProgressDone
        SKW_ProofWorkbook = "A copy of the workbook could not be made (error " & Err.Number & "). Nothing was done."
        Exit Function
    End If
    On Error GoTo 0
    Set before = BookPrint(copyPath, hdr, why)
    ' Write one row through the tools' own writer: the Styles tab if it has a Name column, else the first tab with a column name.
    For Each k In hdr.Keys
        parts = Split(hdr(k), vbTab)
        If LCase$(parts(0)) = "styles" And Len(parts(1)) > 0 Then sheetNm = parts(0): keyHdr = "Name"
    Next
    If Len(sheetNm) = 0 Then
        For Each k In hdr.Keys
            parts = Split(hdr(k), vbTab)
            If Len(sheetNm) = 0 And Len(parts(1)) > 0 Then sheetNm = parts(0): keyHdr = parts(1)
        Next
    End If
    If Len(sheetNm) > 0 Then
        Set row = CreateObject("Scripting.Dictionary")
        row(keyHdr) = "ZZ Stationkeeper proof row"
        If keyHdr = "Name" Then row("Status") = "Draft"
        wrows.Add row
        e = SKW_BookWriteTo(copyPath, sheetNm, keyHdr, wrows, False, a, c, s)
        kept = SKW_LastBookBackup
    Else
        e = "The workbook has no tab with a column name to write under."
    End If
    Application.StatusBar = False
    Set after = BookPrint(copyPath, hdr2, why)
    For Each k In before.Keys
        If Not after.Exists(k) Then
            nDiff = nDiff + 1
            rows.Add SKW_Row(k, "Gone after the save", before(k), "")
        ElseIf after(k) = before(k) Then
            nSame = nSame + 1
            rows.Add SKW_Row(k, "The same", before(k), "")
        ElseIf k = TabLabelOf(hdr, sheetNm) And a = 1 Then
            ' The tab the row went onto has one row more, and nothing else may differ.
            If OnlyRowsDiffer(before(k), after(k)) Then
                nSame = nSame + 1
                rows.Add SKW_Row(k, "The same, plus the proof's one row", before(k), after(k))
            Else
                nDiff = nDiff + 1
                rows.Add SKW_Row(k, "Different", before(k), after(k))
            End If
        Else
            nDiff = nDiff + 1
            rows.Add SKW_Row(k, "Different", before(k), after(k))
        End If
    Next
    On Error Resume Next
    Kill copyPath
    If Len(kept) > 0 Then Kill kept
    RmDir dirp
    On Error GoTo 0
    SKW_ProgressDone
    msg = "The workbook was read in " & secs1 & " s (" & nCells & " cells; that time includes starting Excel). " & _
        IIf(Len(e) > 0, "The proof's row could not be written to the copy: " & Replace(e, "Global Standards workbook", "workbook") & " ", "One row was written to a copy through the tools' own writer and the copy saved. ") & _
        nSame & " of " & (nSame + nDiff) & " tab(s) kept their columns, hidden columns, folded columns and drop-downs; " & nDiff & " did not."
    ProofReport "Work PC Proof 3: A Workbook", msg & " The workbook you chose was only read. The copy, and the backup the writer keeps, were deleted afterwards.", _
        Array("Tab", "After the save", "Before", "After, where it differs"), rows
    SKW_ProofWorkbook = msg
End Function

Private Function TabLabelOf(ByVal hdr As Object, ByVal sheetNm As String) As String
    Dim k As Variant
    For Each k In hdr.Keys
        If Split(hdr(k), vbTab)(0) = sheetNm Then TabLabelOf = CStr(k)
    Next
End Function

Private Function OnlyRowsDiffer(ByVal a As String, ByVal b As String) As Boolean
    Dim p As Long, q As Long
    p = InStr(1, a, ", columns ")
    q = InStr(1, b, ", columns ")
    If p = 0 Or q = 0 Then Exit Function
    OnlyRowsDiffer = (Mid$(a, p) = Mid$(b, q)) And (Val(Mid$(b, 6)) = Val(Mid$(a, 6)) + 1)
End Function

' Asks for a workbook. Times how long it takes to read whole, then saves one row to a copy through the tools' own writer and counts, tab by tab, whether the columns, hidden columns, folded columns and drop-downs are still there. The workbook you choose is only read; the copy is deleted.
Public Sub ProofAWorkbook()
    Dim p As String
    If SKW_IsQuiet() Then Exit Sub
    With Application.FileDialog(3)
        .Title = "Choose a workbook to prove (it is only read)"
        .Filters.Clear
        .Filters.Add "Excel workbooks", "*.xlsx;*.xlsm"
        .AllowMultiSelect = False
        If Len(SKW_StandardsPath()) > 0 Then .InitialFileName = SKW_StandardsPath()
        If .Show <> -1 Then Exit Sub
        p = .SelectedItems(1)
    End With
    SKW_Show SKW_ProofWorkbook(p)
End Sub

' =====================================================================
' Proof 4: a scratch workbook in a synced folder, written while someone else has it open.
' =====================================================================
Private Function MakeScratch(ByVal p As String) As String
    Dim xl As Object, wb As Object, ws As Object
    On Error GoTo No
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Styles"
    ws.Range("A1:C1").Value = Array("Name", "Status", "Notes")
    ws.Range("A2:C2").Value = Array("Stationkeeper proof scratch", "Draft", "A scratch workbook made by a Stationkeeper proof. It holds nothing of yours and can be deleted.")
    wb.SaveAs p, 51
    wb.Close False
    xl.Quit
    Exit Function
No:
    MakeScratch = "error " & Err.Number & " " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
End Function

Private Function SiblingCount(ByVal folder As String) As Long
    Dim f As String
    f = Dir$(folder & "\Stationkeeper proof scratch*.xlsx")
    Do While Len(f) > 0
        If Left$(f, 2) <> "~$" Then SiblingCount = SiblingCount + 1
        f = Dir$
    Loop
End Function

Private Function LockFileThere(ByVal folder As String) As Boolean
    On Error Resume Next
    LockFileThere = Len(Dir$(folder & "\~$" & SCRATCH_BOOK, vbHidden)) > 0
End Function

' held: "yes" or "no", the person's answer to whether someone else has the scratch workbook open now.
Public Function SKW_ProofShared(ByVal folder As String, ByVal held As String) As String
    Dim p As String, e As String, rows As New Collection, wrows As New Collection, row As Object, a As Long, c As Long, s As Long, msg As String
    Dim nBefore As Long, nAfter As Long, isFree As Boolean, f As Integer, nRows As Long, lockSeen As Boolean
    Set SKW_LastReport = Nothing
    If Right$(folder, 1) = "\" Then folder = Left$(folder, Len(folder) - 1)
    folder = SKW_LocalPath(folder)
    If Len(folder) = 0 Or LCase$(Left$(folder, 4)) = "http" Or Len(Dir$(folder, vbDirectory)) = 0 Then SKW_ProofShared = "That folder was not found on this PC. Nothing was done.": Exit Function
    p = folder & "\" & SCRATCH_BOOK
    SKW_ProgressTitle "Work PC proof 4"
    If Len(Dir$(p)) = 0 Then
        SKW_Status "Making the scratch workbook..."
        e = MakeScratch(p)
        SKW_ProgressDone
        If Len(e) > 0 Then SKW_ProofShared = "The scratch workbook could not be made in that folder (" & e & "). Nothing else was done.": Exit Function
        SKW_ProofShared = "Made " & SCRATCH_BOOK & " in that folder. It holds nothing of yours." & vbCr & vbCr & _
            "Next: wait until it has synced. Ask a colleague to open it on another PC, or open it yourself in the browser, and leave it open. Then run this proof again and choose the same folder."
        Exit Function
    End If
    SKW_Status "Writing to the scratch workbook..."
    nBefore = SiblingCount(folder)
    lockSeen = LockFileThere(folder)
    On Error Resume Next
    f = FreeFile
    Err.Clear
    Open p For Binary Access Read Write Lock Read Write As #f
    isFree = (Err.Number = 0)
    Close #f
    On Error GoTo 0
    Set row = CreateObject("Scripting.Dictionary")
    row("Name") = "Proof row " & Format$(Now, "yyyy-mm-dd hhnnss")
    row("Status") = "Draft"
    row("Notes") = "Written by proof 4. Someone else had it open: " & held
    wrows.Add row
    e = SKW_BookWriteTo(p, "Styles", "Name", wrows, False, a, c, s)
    Application.StatusBar = False
    On Error Resume Next
    If Len(SKW_LastBookBackup) > 0 Then Kill SKW_LastBookBackup
    On Error GoTo 0
    nAfter = SiblingCount(folder)
    nRows = ScratchRows(p)
    SKW_ProgressDone
    rows.Add SKW_Row("The folder is one this PC syncs", IIf(SKW_InSyncRoot(folder), "Yes", "No"), "A proof of syncing needs a synced folder.")
    rows.Add SKW_Row("Someone else had the scratch workbook open (your answer)", held, "")
    rows.Add SKW_Row("Excel's lock file was beside the workbook on this PC", IIf(lockSeen, "Yes", "No"), "Excel on this PC makes one. Another PC's does not always sync here.")
    rows.Add SKW_Row("This PC saw the workbook as free to write", IIf(isFree, "Yes", "No"), IIf(isFree And LCase$(held) = "yes", "The tools cannot tell it is open elsewhere.", ""))
    rows.Add SKW_Row("The row was written", IIf(Len(e) = 0 And a = 1, "Yes", "No"), Replace(e, "Global Standards workbook", "scratch workbook"))
    rows.Add SKW_Row("Proof rows in the scratch workbook on this PC now", CStr(nRows), "Ask the other person how many rows they see, and whether Excel showed them a message.")
    rows.Add SKW_Row("Scratch workbooks in the folder before the write", CStr(nBefore), "More than one means a conflict copy.")
    rows.Add SKW_Row("Scratch workbooks in the folder after the write", CStr(nAfter), "Run this proof again in a few minutes: a conflict copy can take a while to appear.")
    msg = "Someone else had it open: " & held & ". This PC saw the workbook as " & IIf(isFree, "free", "held") & ". The row was " & IIf(Len(e) = 0 And a = 1, "written", "not written") & _
        ". Scratch workbooks in the folder: " & nBefore & " before, " & nAfter & " after; proof rows on this PC: " & nRows & "."
    ProofReport "Work PC Proof 4: A Shared Workbook", msg & " Only the scratch workbook was written. Delete it, and any copy of it, when the proof is done.", _
        Array("What was looked at", "Result", "Note"), rows
    SKW_ProofShared = msg
End Function

Private Function ScratchRows(ByVal p As String) As Long
    Dim xl As Object, wb As Object
    On Error GoTo No
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(FileName:=p, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    ScratchRows = wb.Worksheets("Styles").UsedRange.Rows.Count - 2
    wb.Close False
    xl.Quit
    Exit Function
No:
    ScratchRows = -1
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
End Function

' Asks for a synced folder. The first run makes a scratch workbook there. Run again while a colleague has that scratch workbook open: it writes one row and reports whether this PC saw the file as held, whether the row landed, and whether a conflict copy appeared. Only the scratch workbook is written.
Public Sub ProofSharedWorkbook()
    Dim p As String, a As VbMsgBoxResult, held As String
    If SKW_IsQuiet() Then Exit Sub
    With Application.FileDialog(4)
        .Title = "Choose a synced folder for the scratch workbook"
        If .Show <> -1 Then Exit Sub
        p = .SelectedItems(1)
    End With
    If Len(Dir$(p & "\" & SCRATCH_BOOK)) > 0 Then
        a = MsgBox("Does someone else have " & SCRATCH_BOOK & " open right now, on another PC or in the browser?" & vbCr & vbCr & _
            "Yes: this proof writes one row to it and reports what happened." & vbCr & "No: it writes one row with nobody in it, to compare.", vbYesNoCancel + vbQuestion, SKW_APP)
        If a = vbCancel Then Exit Sub
        held = IIf(a = vbYes, "yes", "no")
    End If
    SKW_Show SKW_ProofShared(p, held)
End Sub
