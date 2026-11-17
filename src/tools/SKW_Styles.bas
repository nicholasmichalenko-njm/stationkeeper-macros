Attribute VB_Name = "SKW_Styles"
' =====================================================================
' Stationkeeper Tools: styles and direct formatting.
'   StyleInspector                   what styles the cursor is in, and what is direct formatting
'   NonAllowedStyleReport            report: styles in use that are not on the team list
'   DirectFormattingReport           report: text formatted by hand instead of by style
'   ClearDirectFormattingSelection   clears hand formatting in the selection, keeps emphasis
'   ImportTeamStyles                 brings in the team's style definitions
'   RemapStyle                       a window: changes a style to another by what each paragraph is (list, level, table), keeping formatting by hand as chosen
'   FontNotAllowed                   report: fonts used that are not on the team list
'   RemoveUnusedCustomStyles         deletes custom styles nobody uses
' =====================================================================
Option Explicit

Private Const UNDEF As Long = 9999999
Private mRmKeys() As String          ' Remap Style: the row keys of the last check, 1-based: "From|kind|level"
Private mRmCount As Long

' Shows the paragraph and character style at the cursor and any formatting applied by hand on top of them. With several paragraphs selected, makes a report. Changes nothing.
Public Sub StyleInspector()
    Dim s As String
    s = SKW_StyleInspector()
    If SKW_LastReport Is Nothing And Not SKW_IsQuiet() Then
        MsgBox s, vbInformation, "Style Inspector"
    Else
        SKW_Show s
    End If
End Sub

' One report of what is off the team's lists: styles used that are not allowed, required styles not used, and every font used with the ones not on the Fonts sheet flagged. Without a standards workbook, lists custom styles and the fonts in use. Changes nothing.
Public Sub NonAllowedStyleReport()
    SKW_Show SKW_RunSet("Styles and Fonts Check", "", Array("SKW_NonAllowedStyleReport|Styles", "SKW_FontNotAllowed|Fonts"), False)
End Sub

' Lists paragraphs whose font or spacing was set by hand instead of by their style. Bold, italic, superscript and subscript are treated as normal emphasis (change this in the standards workbook). Changes nothing.
Public Sub DirectFormattingReport()
    SKW_Show SKW_DirectFormattingReport()
End Sub

' Clears formatting applied by hand in the selected text, so it looks like its style again. Keeps bold, italic, superscript and subscript (set in the standards workbook). One Ctrl+Z undoes it.
Public Sub ClearDirectFormattingSelection()
    SKW_Show SKW_ClearDirectFormattingSelection()
End Sub

' ---------------------------------------------------------------------
' What to keep as emphasis (not reported, not cleared).
' ---------------------------------------------------------------------
Public Function SKW_KeepList() As String
    Dim v As String
    v = LCase$(SKW_Setting("KeepOnClear", "Bold, Italic, Superscript, Subscript"))
    If v = "none" Then v = ""
    SKW_KeepList = "," & Replace(Replace(v, ";", ","), " ", "") & ","
End Function

Private Function Kept(ByVal keep As String, ByVal prop As String) As Boolean
    Kept = (InStr(keep, "," & LCase$(prop) & ",") > 0)
End Function

Private Function OnOff(ByVal v As Long) As String
    If v = UNDEF Then OnOff = "mixed" Else OnOff = IIf(v <> 0, "on", "off")
End Function

Private Function FontNameOf(ByVal d As Object, ByVal n As String) As String
    If Left$(n, 1) = "+" Then
        On Error Resume Next
        If InStr(LCase$(n), "head") > 0 Then n = d.DocumentTheme.ThemeFontScheme.MajorFont(1).Name _
        Else n = d.DocumentTheme.ThemeFontScheme.MinorFont(1).Name
    End If
    FontNameOf = n
End Function

' Hand-applied character formatting of rng compared with style st.
' Returns "" when none; "?" when some property is mixed (look closer).
Public Function SKW_FontDiffs(ByVal d As Object, ByVal rng As Object, ByVal st As Object, ByVal keep As String) As String
    Dim f As Object, sf As Object, out As String, mixed As Boolean, a As String, b As String
    Set f = rng.Font
    Set sf = st.Font
    a = FontNameOf(d, f.Name): b = FontNameOf(d, sf.Name)
    If Len(a) = 0 Then
        mixed = True
    ElseIf LCase$(a) <> LCase$(b) And Len(b) > 0 Then
        out = out & "; font " & a & " (style: " & b & ")"
    End If
    If f.Size = UNDEF Then
        mixed = True
    ElseIf f.Size <> sf.Size Then
        out = out & "; size " & f.Size & " (style: " & sf.Size & ")"
    End If
    If Not Kept(keep, "bold") Then CheckOnOff "bold", f.Bold, sf.Bold, out, mixed
    If Not Kept(keep, "italic") Then CheckOnOff "italic", f.Italic, sf.Italic, out, mixed
    If Not Kept(keep, "underline") Then
        If f.Underline = UNDEF Then
            mixed = True
        ElseIf f.Underline <> sf.Underline Then
            out = out & "; underline " & IIf(f.Underline = 0, "off", "on") & " (style: " & IIf(sf.Underline = 0, "off", "on") & ")"
        End If
    End If
    If Not Kept(keep, "superscript") Then CheckOnOff "superscript", f.Superscript, sf.Superscript, out, mixed
    If Not Kept(keep, "subscript") Then CheckOnOff "subscript", f.Subscript, sf.Subscript, out, mixed
    If Not Kept(keep, "color") Then
        If f.Color = UNDEF Then
            mixed = True
        ElseIf f.Color <> sf.Color Then
            out = out & "; colour changed"
        End If
    End If
    If mixed And Len(out) = 0 Then out = "?"
    If Left$(out, 2) = "; " Then out = Mid$(out, 3)
    SKW_FontDiffs = out
End Function

Private Sub CheckOnOff(ByVal nm As String, ByVal v As Long, ByVal sv As Long, ByRef out As String, ByRef mixed As Boolean)
    If v = UNDEF Then
        mixed = True
    ElseIf (v <> 0) <> (sv <> 0) Then
        out = out & "; " & nm & " " & OnOff(v) & " (style: " & OnOff(sv) & ")"
    End If
End Sub

' Hand-applied paragraph formatting of paragraph p compared with its style.
Public Function SKW_ParaDiffs(ByVal p As Object, ByVal st As Object) As String
    Dim pf As Object, sp As Object, out As String, inList As Boolean
    Set pf = p.Format
    Set sp = st.ParagraphFormat
    On Error Resume Next
    inList = (p.Range.ListFormat.ListType <> wdListNoNumbering)
    If pf.SpaceBefore <> sp.SpaceBefore Then out = out & "; space before " & pf.SpaceBefore & " pt (style: " & sp.SpaceBefore & ")"
    If pf.SpaceAfter <> sp.SpaceAfter Then out = out & "; space after " & pf.SpaceAfter & " pt (style: " & sp.SpaceAfter & ")"
    If pf.LineSpacingRule <> sp.LineSpacingRule Or Abs(pf.LineSpacing - sp.LineSpacing) > 0.05 Then out = out & "; line spacing changed"
    If pf.Alignment <> sp.Alignment Then out = out & "; alignment changed"
    If Not inList Then
        If Abs(pf.LeftIndent - sp.LeftIndent) > 0.05 Then out = out & "; left indent " & pf.LeftIndent & " pt (style: " & sp.LeftIndent & ")"
        If Abs(pf.RightIndent - sp.RightIndent) > 0.05 Then out = out & "; right indent changed"
        If Abs(pf.FirstLineIndent - sp.FirstLineIndent) > 0.05 Then out = out & "; first line indent changed"
    End If
    If Left$(out, 2) = "; " Then out = Mid$(out, 3)
    SKW_ParaDiffs = out
End Function

Private Function KeepText(ByVal keep As String) As String
    Dim t As String
    t = Mid$(keep, 2)
    If Len(t) > 0 Then t = Left$(t, Len(t) - 1)
    If Len(t) = 0 Then KeepText = "nothing" Else KeepText = Replace(t, ",", ", ")
End Function

Private Function TextOnly(ByVal p As Object) As Object
    Dim r As Object
    Set r = p.Range.Duplicate
    If r.End - r.Start > 1 Then r.MoveEnd wdCharacter, -1
    Set TextOnly = r
End Function

Private Function CharStyleName(ByVal r As Object) As String
    Dim cs As Object
    On Error Resume Next
    Set cs = r.CharacterStyle
    If Not cs Is Nothing Then CharStyleName = cs.NameLocal
End Function

Private Function IsPlainChar(ByVal nm As String) As Boolean
    IsPlainChar = (Len(nm) = 0 Or LCase$(nm) = "default paragraph font")
End Function

' Character findings for one paragraph: whole paragraph first, then word by word when mixed.
Private Function CharFindings(ByVal d As Object, ByVal p As Object, ByVal st As Object, ByVal keep As String) As String
    Dim r As Object, s As String, w As Object, seen As Object, one As String, n As Long, k As Variant
    Set r = TextOnly(p)
    If Len(r.Text) = 0 Then Exit Function
    If Not IsPlainChar(CharStyleName(r)) And CharStyleName(r) <> "" Then Exit Function
    s = SKW_FontDiffs(d, r, st, keep)
    If s <> "?" Then CharFindings = s: Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For Each w In r.Words
        n = n + 1
        If n > 400 Then Exit For
        If IsPlainChar(CharStyleName(w)) And Len(Trim$(w.Text)) > 0 Then
            one = SKW_FontDiffs(d, w, st, keep)
            If Len(one) > 0 And one <> "?" Then
                If Not seen.Exists(one) Then seen(one) = Trim$(w.Text)
            End If
        End If
    Next
    s = ""
    For Each k In seen.Keys
        s = s & "; " & k & " at """ & SKW_Excerpt(seen(k), 30) & """"
    Next
    If Len(s) > 0 Then s = Mid$(s, 3)
    CharFindings = s
End Function

' ---------------------------------------------------------------------
Public Function SKW_StyleInspector() As String
    Dim d As Object, msg As String, p As Object, st As Object, s As String, rows As New Collection, keep As String
    Dim cf As String, pf As String, n As Long, csn As String
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Style inspector", False)
    If Len(msg) > 0 Then SKW_StyleInspector = msg: Exit Function
    Set d = ActiveDocument
    keep = SKW_KeepList()
    If Selection.Paragraphs.Count <= 1 Then
        Set p = Selection.Paragraphs(1)
        Set st = SKW_ParaStyle(p.Range)
        s = "Paragraph style: " & st.NameLocal
        On Error Resume Next
        If Len(st.BaseStyle.NameLocal) > 0 Then s = s & " (based on " & st.BaseStyle.NameLocal & ")"
        On Error GoTo 0
        csn = CharStyleName(Selection.Range)
        s = s & vbCr & "Character style at the cursor: " & IIf(IsPlainChar(csn), "none", csn)
        s = s & vbCr & "Font from the style: " & FontNameOf(d, st.Font.Name) & " " & st.Font.Size & " pt"
        cf = CharFindings(d, p, st, keep)
        pf = SKW_ParaDiffs(p, st)
        s = s & vbCr & vbCr & "Formatting applied by hand:" & vbCr & _
            "  Characters: " & IIf(Len(cf) > 0, cf, "none") & vbCr & _
            "  Paragraph: " & IIf(Len(pf) > 0, pf, "none")
        If Len(cf) + Len(pf) > 0 Then s = s & vbCr & vbCr & "Select the text and run Clear Direct Formatting (Selection) to go back to the style."
        SKW_Finish
        SKW_StyleInspector = s
        Exit Function
    End If
    For Each p In Selection.Paragraphs
        n = n + 1
        If n > 500 Then Exit For
        Set st = SKW_ParaStyle(p.Range)
        cf = CharFindings(d, p, st, keep)
        pf = SKW_ParaDiffs(p, st)
        rows.Add SKW_Row(n, st.NameLocal, IIf(Len(cf) > 0, cf, "none"), IIf(Len(pf) > 0, pf, "none"), SKW_Page(p.Range), p.Range.Text)
    Next
    SKW_Finish
    SKW_Report "Style Inspector", "Styles of the " & rows.Count & " selected paragraph(s), and formatting applied by hand on top of them. Changes nothing.", _
        Array("#", "Style", "Characters by hand", "Paragraph by hand", "Page", "Text"), rows
    SKW_StyleInspector = rows.Count & " paragraph(s) inspected."
End Function

' Styles that are always fine (Word uses them on its own).
Private Function SystemStyle(ByVal nm As String) As Boolean
    Select Case LCase$(nm)
        Case "default paragraph font", "normal table", "table normal", "no list", "hyperlink", "followedhyperlink", "comment reference", _
             "comment text", "comment subject", "balloon text", "placeholder text", "footnote reference", "endnote reference"
            SystemStyle = True
    End Select
End Function

Public Function SKW_StyleAllowed(ByVal st As Object, ByVal allowed As Collection, ByVal prefix As String) As Boolean
    Dim nm As String
    nm = st.NameLocal
    If SystemStyle(nm) Then SKW_StyleAllowed = True: Exit Function
    If Len(prefix) > 0 Then
        If LCase$(Left$(nm, Len(prefix))) = LCase$(prefix) Then SKW_StyleAllowed = True: Exit Function
    End If
    If allowed.Count = 0 Then
        SKW_StyleAllowed = st.BuiltIn
    Else
        SKW_StyleAllowed = SKW_InList(allowed, nm)
    End If
End Function

Public Function SKW_NonAllowedStyleReport() As String
    Dim d As Object, msg As String, allowed As Collection, required As New Collection, prefix As String, r As Object
    Dim st As Object, story As Object, hits As Collection, n As Long, first As Object, rows As New Collection
    Dim t As Object, tn As String, i As Long, intro As String, nBad As Long, nReq As Long, v As Variant, kind As String
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Non-allowed style report", False)
    If Len(msg) > 0 Then SKW_NonAllowedStyleReport = msg: Exit Function
    Set d = ActiveDocument
    Set allowed = SKW_Column("Allowed styles", "Style")
    For Each r In SKW_Rows("Allowed styles")
        If r.Exists("required") Then
            If SKW_IsYes(r("required")) Then required.Add r("style")
        End If
    Next
    prefix = SKW_Setting("StylePrefix", "")
    SKW_Steps d.Styles.Count, "Style"
    For Each st In d.Styles
        SKW_Step
        If st.InUse And st.Type <> wdStyleTypeTable And st.Type <> wdStyleTypeList Then
            If Not SKW_StyleAllowed(st, allowed, prefix) Then
                n = 0
                Set first = Nothing
                For Each story In SKW_Stories(d)
                    If story.StoryType <> wdCommentsStory Then
                        Set hits = SKW_StyleRanges(story, st, 500)
                        n = n + hits.Count
                        If first Is Nothing And hits.Count > 0 Then Set first = hits(1)
                    End If
                Next
                If n > 0 Then
                    nBad = nBad + 1
                    kind = IIf(st.Type = wdStyleTypeCharacter, "Character", "Paragraph")
                    rows.Add SKW_Row(IIf(allowed.Count = 0, "Custom style", "Not allowed"), _
                        kind & " style " & st.NameLocal & " used " & n & " time(s)" & IIf(n >= 500, " or more", ""), _
                        SKW_StoryName(first.StoryType), SKW_Page(first), first.Text)
                End If
            End If
        End If
    Next
    SKW_Steps d.Tables.Count, "Table"
    For Each t In d.Tables
        SKW_Step
        i = i + 1
        tn = ""
        On Error Resume Next
        tn = t.Style.NameLocal
        On Error GoTo 0
        If Len(tn) > 0 And allowed.Count > 0 Then
            If Not SKW_InList(allowed, tn) And Not SystemStyle(tn) And Not (Len(prefix) > 0 And LCase$(Left$(tn, Len(prefix))) = LCase$(prefix)) Then
                nBad = nBad + 1
                rows.Add SKW_Row("Not allowed", "Table " & i & " uses table style " & tn, "Body", SKW_Page(t.Range), t.Range.Cells(1).Range.Text)
            End If
        End If
    Next
    For Each v In required
        Set st = SKW_GetStyle(d, CStr(v))
        n = 0
        If Not st Is Nothing Then n = SKW_StyleRanges(d.Content, st, 1).Count
        If n = 0 Then
            nReq = nReq + 1
            rows.Add SKW_Row("Required", "Required style " & v & " is " & IIf(st Is Nothing, "not in the document", "not used"), "Body", "", "")
        End If
    Next
    SKW_Finish
    If allowed.Count = 0 Then
        intro = "No allowed-styles list, so built-in Word styles count as allowed and custom styles are listed for you to check."
    Else
        intro = allowed.Count & " allowed style(s) on the team list" & IIf(Len(prefix) > 0, ", plus any style starting with " & prefix, "") & "."
    End If
    SKW_Report "Non-Allowed Style Report", intro & " Changes nothing.", Array("Rule", "Finding", "Where", "Page", "First use"), rows
    SKW_NonAllowedStyleReport = nBad & " style(s) not allowed, " & nReq & " required style(s) missing."
End Function

Public Function SKW_DirectFormattingReport() As String
    Dim d As Object, msg As String, story As Object, p As Object, st As Object, rows As New Collection, keep As String
    Dim cf As String, pf As String, n As Long, nPar As Long, inTables As Boolean, total As Long, stn As String, inCell As Boolean
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Direct formatting report", False)
    If Len(msg) > 0 Then SKW_DirectFormattingReport = msg: Exit Function
    Set d = ActiveDocument
    keep = SKW_KeepList()
    inTables = SKW_IsYes(SKW_Setting("CheckTablesForDirectFormatting", "Yes"))
    For Each story In SKW_Stories(d)
        If story.StoryType <> wdCommentsStory Then
            stn = SKW_StoryName(story.StoryType)
            SKW_Steps story.Paragraphs.Count, "Paragraph"
            For Each p In story.Paragraphs
                SKW_Step
                nPar = nPar + 1
                If nPar Mod 100 = 0 Then SKW_Status "Direct formatting: paragraph " & nPar & "..."
                inCell = p.Range.Information(wdWithInTable)
                If inTables Or Not inCell Then
                    Set st = SKW_ParaStyle(p.Range)
                    If st Is Nothing Then
                        cf = "": pf = ""
                    ElseIf inCell Then
                        ' Table styles set spacing, emphasis and colour without hand formatting: check font and size only.
                        cf = CharFindings(d, p, st, keep & "bold,italic,underline,color,")
                        pf = ""
                    Else
                        cf = CharFindings(d, p, st, keep)
                        pf = SKW_ParaDiffs(p, st)
                    End If
                    If Len(cf) + Len(pf) > 0 Then
                        total = total + 1
                        If rows.Count < 2000 Then
                            rows.Add SKW_Row(IIf(Len(cf) > 0 And Len(pf) > 0, "Characters and paragraph", IIf(Len(cf) > 0, "Characters", "Paragraph")), _
                                st.NameLocal & ": " & cf & IIf(Len(cf) > 0 And Len(pf) > 0, "; ", "") & pf, stn, SKW_Page(p.Range), p.Range.Text)
                        End If
                    End If
                End If
            Next
        End If
    Next
    SKW_Finish
    SKW_Report "Direct Formatting Report", total & " of " & nPar & " paragraph(s) have formatting applied by hand." & _
        IIf(total > rows.Count, " The first " & rows.Count & " are listed.", "") & _
        " Not counted as hand formatting: " & KeepText(keep) & _
        ", and character styles. " & IIf(inTables, "In tables only font and size are checked, since table styles set the rest. ", "Tables are skipped. ") & _
        "To fix, select the text and run Clear Direct Formatting (Selection). Changes nothing.", _
        Array("Kind", "Finding", "Where", "Page", "Text"), rows
    SKW_DirectFormattingReport = total & " paragraph(s) with hand formatting."
End Function

Public Function SKW_ClearDirectFormattingSelection() As String
    Dim d As Object, msg As String, sel As Object, keep As String, props As Variant, pr As Variant
    Dim saved As New Collection, f As Object, item As Variant, r As Object, nKept As Long, lastStart As Long
    Dim p As Object, nList As Long
    Set SKW_LastReport = Nothing
    If Documents.Count > 0 Then
        If Selection.Type = wdSelectionIP Or Len(Selection.Range.Text) = 0 Then
            SKW_ClearDirectFormattingSelection = "Select the text to clear first."
            Exit Function
        End If
    End If
    msg = SKW_Begin("Clear direct formatting", True)
    If Len(msg) > 0 Then SKW_ClearDirectFormattingSelection = msg: Exit Function
    Set d = ActiveDocument
    Set sel = Selection.Range.Duplicate
    keep = SKW_KeepList()
    On Error GoTo Fail
    ' 1. Remember the emphasis to keep.
    props = Array("bold", "italic", "underline", "superscript", "subscript")
    For Each pr In props
        If Kept(keep, CStr(pr)) Then
            Set f = sel.Duplicate
            With f.Find
                .ClearFormatting
                .Text = ""
                .Format = True
                .Forward = True
                .Wrap = wdFindStop
                Select Case pr
                    Case "bold": .Font.Bold = True
                    Case "italic": .Font.Italic = True
                    Case "underline": .Font.Underline = wdUnderlineSingle
                    Case "superscript": .Font.Superscript = True
                    Case "subscript": .Font.Subscript = True
                End Select
            End With
            lastStart = -1
            Do While f.Find.Execute
                If Not f.InRange(sel) Or f.Start = lastStart Then Exit Do
                lastStart = f.Start
                saved.Add Array(pr, f.Start, f.End)
                If f.End >= sel.End Then Exit Do
                f.Collapse wdCollapseEnd
            Loop
        End If
    Next
    ' 2. Clear, then put the kept emphasis back where the style does not give it.
    sel.Select
    Selection.ClearCharacterDirectFormatting
    For Each p In sel.Paragraphs
        If p.Range.ListFormat.ListType = wdListNoNumbering Then p.Reset Else nList = nList + 1
    Next
    For Each item In saved
        Set r = d.Range(item(1), item(2))
        Select Case item(0)
            Case "bold": If r.Font.Bold <> True Then r.Font.Bold = True: nKept = nKept + 1
            Case "italic": If r.Font.Italic <> True Then r.Font.Italic = True: nKept = nKept + 1
            Case "underline": If r.Font.Underline <> wdUnderlineSingle Then r.Font.Underline = wdUnderlineSingle: nKept = nKept + 1
            Case "superscript": If r.Font.Superscript <> True Then r.Font.Superscript = True: nKept = nKept + 1
            Case "subscript": If r.Font.Subscript <> True Then r.Font.Subscript = True: nKept = nKept + 1
        End Select
    Next
    sel.Select
    SKW_Finish
    SKW_ClearDirectFormattingSelection = "Cleared formatting applied by hand in the selection (" & sel.Paragraphs.Count & " paragraph(s))." & _
        IIf(nKept > 0, " Kept " & nKept & " piece(s) of emphasis.", "") & _
        IIf(nList > 0, " Paragraph spacing and indents in " & nList & " list paragraph(s) were left alone so the numbering stays.", "") & _
        " One Ctrl+Z undoes it."
    Exit Function
Fail:
    msg = Err.Description
    SKW_Finish
    SKW_ClearDirectFormattingSelection = "Clearing stopped: " & msg & ". Press Ctrl+Z to undo what was done."
End Function

' ---------------------------------------------------------------------
' Import Team Styles and Remap Style
' ---------------------------------------------------------------------

' Brings the team's style definitions into this document from the team template (setting StyleSource, or the template the document is attached to), and lists which styles changed. One Ctrl+Z undoes it.
Public Sub ImportTeamStyles()
    If SKW_IsQuiet() Then
        SKW_Status SKW_ImportTeamStyles("")
        Exit Sub
    End If
    If MsgBox("Bring the team's style definitions into " & ActiveDocument.Name & "?" & vbCr & vbCr & _
              "Styles with the same name are replaced by the team's version. One Ctrl+Z puts them back.", _
              vbOKCancel + vbQuestion, SKW_APP) <> vbOK Then Exit Sub
    SKW_Show SKW_ImportTeamStyles("")
End Sub

' Changes a style to another without losing what the paragraphs are. A window lists the styles in the document to pick from, shows each kind of paragraph in the style (not in a list, bulleted or numbered by hand with its level, in a table) and lets you pick what each becomes. Formatting applied by hand is kept, kept in part or cleared, as chosen. One Ctrl+Z undoes each run.
Public Sub RemapStyle()
    If SKW_IsQuiet() Then Exit Sub
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    frmSKWRemap.Show vbModeless
End Sub

Private Function StyleSig(ByVal st As Object) As String
    On Error Resume Next
    With st.Font
        StyleSig = .Name & "|" & .Size & "|" & .Bold & "|" & .Italic & "|" & .Color
    End With
    If st.Type = wdStyleTypeParagraph Or st.Type = 4 Then
        With st.ParagraphFormat
            StyleSig = StyleSig & "|" & .SpaceBefore & "|" & .SpaceAfter & "|" & .LineSpacing & "|" & .LeftIndent & "|" & .Alignment & "|" & .KeepWithNext
        End With
    End If
End Function

Public Function SKW_ImportTeamStyles(ByVal source As String) As String
    Dim d As Object, msg As String, before As Object, st As Object, changed As String, n As Long, added As Long
    Set SKW_LastReport = Nothing
    If Documents.Count = 0 Then SKW_ImportTeamStyles = "Open a document first.": Exit Function
    Set d = ActiveDocument
    If Len(source) = 0 Then source = SKW_Setting("StyleSource", "")
    If Len(source) = 0 Then
        On Error Resume Next
        source = d.AttachedTemplate.FullName
        On Error GoTo 0
        If LCase$(Mid$(source, InStrRev(source, "\") + 1)) = "normal.dotm" Then
            SKW_ImportTeamStyles = "This document is attached to Normal.dotm. Put the team template's path in the StyleSource setting of the standards workbook."
            Exit Function
        End If
    End If
    If Len(Dir$(source)) = 0 Then SKW_ImportTeamStyles = "Team style source not found: " & source: Exit Function
    If LCase$(source) = LCase$(d.FullName) Then SKW_ImportTeamStyles = "The document is its own style source; nothing to import.": Exit Function
    msg = SKW_Begin("Import team styles", True)
    If Len(msg) > 0 Then SKW_ImportTeamStyles = msg: Exit Function
    Set before = CreateObject("Scripting.Dictionary")
    SKW_Steps d.Styles.Count, "Style"
    For Each st In d.Styles
        SKW_Step
        If st.InUse Then before(st.NameLocal) = StyleSig(st)
    Next
    On Error GoTo Fail
    d.CopyStylesFromTemplate source
    On Error GoTo 0
    SKW_Steps d.Styles.Count, "Style"
    For Each st In d.Styles
        SKW_Step
        If st.InUse Then
            If before.Exists(st.NameLocal) Then
                If before(st.NameLocal) <> StyleSig(st) Then
                    n = n + 1
                    If n <= 25 Then changed = changed & ", " & st.NameLocal
                End If
            Else
                added = added + 1
            End If
        End If
    Next
    SKW_Finish
    SKW_ImportTeamStyles = "Styles brought in from " & Mid$(source, InStrRev(source, "\") + 1) & "." & vbCr & _
        n & " style(s) changed" & IIf(n > 0, ": " & Mid$(changed, 3) & IIf(n > 25, ", ...", ""), "") & "." & vbCr & _
        added & " style(s) added." & vbCr & "Text formatted by hand keeps that formatting: run Direct Formatting Report next." & vbCr & "One Ctrl+Z undoes it."
    Exit Function
Fail:
    msg = Err.Description
    SKW_Finish
    SKW_ImportTeamStyles = "Styles could not be brought in: " & msg
End Function

' =====================================================================
' Remap Style (1.14.0): by what each paragraph is
' =====================================================================
' Every paragraph in a From style is one of: plain, bulleted or numbered by hand (with its level),
' in a table, or led by a typed bullet. A row of the report is one From style and one of those.
'   plain            becomes the To style given for the From style (blank: left alone)
'   bullet, number   becomes the Style of the approved List numbering row of that kind and level in the
'                    Global Standards workbook. With no such row or style it is left alone: a bullet is never lost
'   table            left alone unless that one row is applied
'   typed bullet     always left alone
' A list that comes from the From style itself is not "by hand": those paragraphs are plain.

Private Function RmStyleHasList(ByVal st As Object) As Boolean
    Dim lt As Object
    On Error Resume Next
    Set lt = st.ListTemplate
    RmStyleHasList = Not lt Is Nothing
End Function

Private Function RmTyped(ByVal t As String) As Boolean
    Dim c As Long, nx As String
    If Len(t) < 3 Then Exit Function
    c = AscW(Left$(t, 1))
    nx = Mid$(t, 2, 1)
    If nx <> vbTab And nx <> " " Then Exit Function
    Select Case c
        Case 8226, 183, 9642, 9679, -3913, 8211, 45, 42: RmTyped = True
    End Select
End Function

' "plain|0", "bullet|2", "number|1", "table|0" or "typed|0"
Private Function RmKind(ByVal p As Object, ByVal styleList As Boolean) As String
    Dim lt As Long, lvl As Long, ns As Long
    On Error Resume Next
    If p.Range.Information(wdWithInTable) Then RmKind = "table|0": Exit Function
    lt = p.Range.ListFormat.ListType
    If lt = wdListNoNumbering Or lt = wdListListNumOnly Or styleList Then
        RmKind = IIf(RmTyped(p.Range.Text), "typed|0", "plain|0")
        Exit Function
    End If
    lvl = p.Range.ListFormat.ListLevelNumber
    If lvl < 1 Then lvl = 1
    ns = -1
    ns = p.Range.ListFormat.ListTemplate.ListLevels(lvl).NumberStyle
    If ns = wdListNumberStyleBullet Or ns = wdListNumberStylePictureBullet Or (ns = -1 And (lt = wdListBullet Or lt = wdListPictureBullet)) Then
        RmKind = "bullet|" & lvl
    Else
        RmKind = "number|" & lvl
    End If
End Function

' The list style the Global Standards give a kind and level, or "" with the reason in why.
Private Function RmListStyle(ByVal knd As String, ByVal lvl As Long, ByRef why As String) As String
    Dim r As Object, fmt As String, st As String, isB As Boolean, n As Long
    On Error Resume Next
    For Each r In SKW_WbRows("List numbering")
        n = n + 1
        st = ""
        If r.Exists("status") Then st = LCase$(Trim$(r("status")))
        If Len(st) = 0 Or st = "approved" Then
            fmt = ""
            If r.Exists("fmt") Then fmt = LCase$(Trim$(r("fmt")))
            isB = (fmt = "bullet")
            If isB = (knd = "bullet") And Val(r("level")) = lvl Then
                If r.Exists("style") Then RmListStyle = Trim$(r("style"))
                If Len(RmListStyle) > 0 Then Exit Function
            End If
        End If
    Next
    If n = 0 Then
        why = "no Global Standards workbook with a List numbering sheet is chosen"
    Else
        why = "the List numbering sheet has no approved " & knd & " row for level " & lvl & " with a Style"
    End If
End Function

Private Function RmKindWords(ByVal knd As String) As String
    Select Case knd
        Case "plain": RmKindWords = "Not in a list"
        Case "bullet": RmKindWords = "Bulleted by hand"
        Case "number": RmKindWords = "Numbered by hand"
        Case "table": RmKindWords = "In a table"
        Case "typed": RmKindWords = "Typed bullet"
    End Select
End Function

' Does the paragraph carry formatting applied by hand? A quick reading: mixed runs, or the whole
' paragraph differing from its style in emphasis, font, size or indents.
Private Function RmHasDirect(ByVal p As Object, ByVal sa As Object) As Boolean
    Dim f As Object, sf As Object
    On Error Resume Next
    Set f = p.Range.Font
    Set sf = sa.Font
    If f.Bold = UNDEF Or f.Italic = UNDEF Or f.Underline = UNDEF Or f.Size = UNDEF Or f.Color = UNDEF Or Len(f.Name) = 0 Then RmHasDirect = True: Exit Function
    If f.Bold <> sf.Bold Or f.Italic <> sf.Italic Or f.Underline <> sf.Underline Or f.Size <> sf.Size Or f.Name <> sf.Name Or f.Color <> sf.Color Then RmHasDirect = True: Exit Function
    If f.Superscript <> sf.Superscript Or f.Subscript <> sf.Subscript Or f.StrikeThrough <> sf.StrikeThrough Then RmHasDirect = True: Exit Function
    If p.Range.HighlightColorIndex <> wdNoHighlight Then RmHasDirect = True: Exit Function
    If p.Range.ListFormat.ListType = wdListNoNumbering Then
        If Abs(p.LeftIndent - sa.ParagraphFormat.LeftIndent) > 0.05 Or Abs(p.FirstLineIndent - sa.ParagraphFormat.FirstLineIndent) > 0.05 Then RmHasDirect = True: Exit Function
    End If
    If Abs(p.SpaceBefore - sa.ParagraphFormat.SpaceBefore) > 0.05 Or Abs(p.SpaceAfter - sa.ParagraphFormat.SpaceAfter) > 0.05 Or p.Alignment <> sa.ParagraphFormat.Alignment Then RmHasDirect = True
End Function

' One run's formatting, or "" when the range is not one run (something differs inside it).
Private Function RmSig(ByVal r As Object) As String
    Dim f As Object, h As Long
    On Error Resume Next
    Set f = r.Font
    h = r.HighlightColorIndex
    If f.Bold = UNDEF Or f.Italic = UNDEF Or f.Underline = UNDEF Or f.StrikeThrough = UNDEF Or f.Superscript = UNDEF Or f.Subscript = UNDEF _
        Or f.Color = UNDEF Or f.Size = UNDEF Or Len(f.Name) = 0 Or h = UNDEF Then Exit Function
    RmSig = f.Bold & "|" & f.Italic & "|" & f.Underline & "|" & f.StrikeThrough & "|" & f.Superscript & "|" & f.Subscript & "|" & f.Color & "|" & h & "|" & f.Size & "|" & f.Name
End Function

' The runs of a range: each "start|end|signature". Halves the range until each piece is one run.
Private Function RmRange(ByVal story As Object, ByVal a As Long, ByVal b As Long) As Object
    Set RmRange = story.Duplicate
    RmRange.SetRange a, b
End Function

Private Sub RmRuns(ByVal d As Object, ByVal a As Long, ByVal b As Long, ByVal runs As Collection)
    Dim sg As String, m As Long
    If b <= a Then Exit Sub
    sg = RmSig(RmRange(d, a, b))
    If Len(sg) > 0 Or b - a = 1 Then
        If Len(sg) > 0 Then runs.Add a & "|" & b & "|" & sg
        Exit Sub
    End If
    m = a + (b - a) \ 2
    RmRuns d, a, m, runs
    RmRuns d, m, b, runs
End Sub

Private Function RmStyleSig(ByVal st As Object) As String
    Dim f As Object
    On Error Resume Next
    Set f = st.Font
    RmStyleSig = f.Bold & "|" & f.Italic & "|" & f.Underline & "|" & f.StrikeThrough & "|" & f.Superscript & "|" & f.Subscript & "|" & f.Color & "|" & wdNoHighlight & "|" & f.Size & "|" & f.Name
End Function

' Puts back, on each run, what differed from the old style. upTo 8: emphasis only. upTo 10: font size and name too.
Private Sub RmRestore(ByVal d As Object, ByVal runs As Collection, ByVal base As String, ByVal upTo As Long)
    Dim v As Variant, a() As String, b() As String, r As Object, i As Long
    On Error Resume Next
    b = Split(base, "|")
    For Each v In runs
        a = Split(v, "|")
        Set r = RmRange(d, CLng(a(0)), CLng(a(1)))
        For i = 1 To upTo
            If a(i + 1) <> b(i - 1) Then
                Select Case i
                    Case 1: If r.Font.Bold <> CLng(a(2)) Then r.Font.Bold = CLng(a(2))
                    Case 2: If r.Font.Italic <> CLng(a(3)) Then r.Font.Italic = CLng(a(3))
                    Case 3: If r.Font.Underline <> CLng(a(4)) Then r.Font.Underline = CLng(a(4))
                    Case 4: If r.Font.StrikeThrough <> CLng(a(5)) Then r.Font.StrikeThrough = CLng(a(5))
                    Case 5: If r.Font.Superscript <> CLng(a(6)) Then r.Font.Superscript = CLng(a(6))
                    Case 6: If r.Font.Subscript <> CLng(a(7)) Then r.Font.Subscript = CLng(a(7))
                    Case 7: If r.Font.Color <> CLng(a(8)) Then r.Font.Color = CLng(a(8))
                    Case 8: If r.HighlightColorIndex <> CLng(a(9)) Then r.HighlightColorIndex = CLng(a(9))
                    Case 9: If r.Font.Size <> CSng(a(10)) Then r.Font.Size = CSng(a(10))
                    Case 10: If r.Font.Name <> a(11) Then r.Font.Name = a(11)
                End Select
            End If
        Next
    Next
End Sub

' Sets one paragraph. d is the paragraph's story. keep: "all", "text" or "none". Returns "" or a note.
Private Function RmSet(ByVal d As Object, ByVal p As Object, ByVal sa As Object, ByVal sb As Object, ByVal knd As String, ByVal keep As String) As String
    Dim runs As New Collection, base As String, isList As Boolean, lt As Object, lvl As Long, targetList As Boolean
    Dim pf(1 To 6) As Single, al As Long, hadPf As Boolean, a As Long, b As Long
    On Error Resume Next
    isList = (knd = "bullet" Or knd = "number")
    targetList = RmStyleHasList(sb)
    a = p.Range.Start: b = p.Range.End
    If keep <> "none" Then
        base = RmStyleSig(sa)
        RmRuns d, a, b, runs
    End If
    If keep = "all" Then
        hadPf = True
        pf(1) = p.LeftIndent: pf(2) = p.FirstLineIndent: pf(3) = p.RightIndent
        pf(4) = p.SpaceBefore: pf(5) = p.SpaceAfter: pf(6) = p.LineSpacing
        al = p.Alignment
    End If
    If isList Then
        Set lt = p.Range.ListFormat.ListTemplate
        lvl = p.Range.ListFormat.ListLevelNumber
        ' The list style brings its own bullet or number: the one applied by hand comes off first, so there is one.
        If targetList Then p.Range.ListFormat.RemoveNumbers
    End If
    p.Style = sb
    Set p = RmRange(d, a, a).Paragraphs(1)
    If isList Then
        If targetList Then
            ' The list's own indents were set by hand with the bullet: the list style's take over.
            p.LeftIndent = sb.ParagraphFormat.LeftIndent
            p.FirstLineIndent = sb.ParagraphFormat.FirstLineIndent
        End If
        If p.Range.ListFormat.ListType = wdListNoNumbering And Not lt Is Nothing Then
            ' The style has no bullet of its own: the one applied by hand goes back on.
            p.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:=lt, ContinuePreviousList:=True, ApplyLevel:=lvl
            RmSet = "kept the bullet or number applied by hand: " & sb.NameLocal & " has none of its own"
        End If
    Else
        If keep <> "all" Then p.Range.ParagraphFormat.Reset
    End If
    If keep <> "all" Then p.Range.Font.Reset
    If keep = "all" Then
        RmRestore d, runs, base, 10
        If hadPf Then
            If Not isList Then
                If Abs(pf(1) - sa.ParagraphFormat.LeftIndent) > 0.05 Then p.LeftIndent = pf(1)
                If Abs(pf(2) - sa.ParagraphFormat.FirstLineIndent) > 0.05 Then p.FirstLineIndent = pf(2)
            End If
            If Abs(pf(3) - sa.ParagraphFormat.RightIndent) > 0.05 Then p.RightIndent = pf(3)
            If Abs(pf(4) - sa.ParagraphFormat.SpaceBefore) > 0.05 Then p.SpaceBefore = pf(4)
            If Abs(pf(5) - sa.ParagraphFormat.SpaceAfter) > 0.05 Then p.SpaceAfter = pf(5)
            If al <> sa.ParagraphFormat.Alignment Then p.Alignment = al
        End If
    ElseIf keep = "text" Then
        RmRestore d, runs, base, 8
    End If
End Function

' The text of a character style changed to another, as before 1.14.0.
Private Function RmCharStyle(ByVal d As Object, ByVal sa As Object, ByVal sb As Object) As Long
    Dim story As Object, f As Object, n As Long
    For Each story In SKW_Stories(d)
        If story.StoryType <> wdCommentsStory Then
            n = n + SKW_StyleRanges(story, sa).Count
            Set f = story.Duplicate
            With f.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = ""
                .Replacement.Text = ""
                .Format = True
                .Style = sa
                .Replacement.Style = sb
                .Forward = True
                .Wrap = wdFindStop
                .Execute Replace:=wdReplaceAll
            End With
        End If
    Next
    RmCharStyle = n
End Function

' The work behind the check and the apply.
'   doIt False: fills groups (key -> Array(count, withDirect, becomes, note, changes)) and changes nothing.
'   doIt True: sets the paragraphs. onlyKey "" for every row, or one row's key. selA, selB: the selection, or -1.
Private Function RmWalk(ByVal d As Object, ByVal pairs As String, ByVal doIt As Boolean, ByVal keep As String, ByVal onlyKey As String, _
                        ByVal selA As Long, ByVal selB As Long, ByVal selStory As Long, ByVal groups As Object, ByRef skipped As String, ByRef notes As String, _
                        Optional ByVal ov As Object = Nothing) As Long
    Dim pr As Variant, a As String, b As String, sa As Object, sb As Object, st As Object, story As Object, rg As Object, p As Object
    Dim kd As String, knd As String, lvl As Long, key As String, target As String, why As String, will As Boolean, g As Variant
    Dim styleList As Boolean, n As Long, paras As Collection, v As Variant, nt As String, seen As Object, tname As Object
    Set tname = CreateObject("Scripting.Dictionary")
    For Each pr In Split(pairs, "|")
        If InStr(pr, ">") > 0 Then
            a = Trim$(Split(pr, ">")(0)): b = Trim$(Split(pr, ">")(1))
            Set sa = SKW_GetStyle(d, a)
            Set sb = Nothing
            If Len(b) > 0 Then Set sb = SKW_GetStyle(d, b)
            If sa Is Nothing Then
                skipped = skipped & "; " & a & " (" & a & " is not in this document)"
            ElseIf Len(b) > 0 And sb Is Nothing Then
                skipped = skipped & "; " & a & " to " & b & " (" & b & " is not in this document)"
            ElseIf sa.Type = wdStyleTypeCharacter Then
                If sb Is Nothing Then
                    skipped = skipped & "; " & a & " (a character style needs a style to change to)"
                Else
                    key = a & "|char|0"
                    If doIt Then
                        If Len(onlyKey) = 0 Or onlyKey = key Then n = n + RmCharStyle(d, sa, sb)
                    Else
                        groups(key) = Array(0, 0, b, "Character style: the text takes " & b, True)
                        For Each story In SKW_Stories(d)
                            If story.StoryType <> wdCommentsStory Then
                                g = groups(key): g(0) = g(0) + SKW_StyleRanges(story, sa).Count: groups(key) = g
                            End If
                        Next
                    End If
                End If
            Else
                styleList = RmStyleHasList(sa)
                For Each story In SKW_Stories(d)
                    If story.StoryType <> wdCommentsStory Then
                        ' The paragraphs first: a style change does not move them.
                        Set paras = New Collection
                        Set seen = CreateObject("Scripting.Dictionary")
                        For Each rg In SKW_StyleRanges(story, sa)
                            For Each p In rg.Paragraphs
                                If Not seen.Exists(p.Range.Start) Then
                                    seen(p.Range.Start) = 1
                                    paras.Add p.Range.Start
                                End If
                            Next
                        Next
                        SKW_Steps paras.Count, "Paragraph"
                        For Each v In paras
                            SKW_Step
                            Set p = RmParaAt(story, CLng(v))
                            If Not p Is Nothing Then
                                kd = RmKind(p, styleList)
                                knd = Split(kd, "|")(0): lvl = Val(Split(kd, "|")(1))
                                key = a & "|" & kd
                                target = "": why = "": will = False
                                Select Case knd
                                    Case "plain"
                                        If sb Is Nothing Then why = "no style was given for paragraphs that are not in a list" Else target = b: will = True
                                    Case "bullet", "number"
                                        If Not tname.Exists(kd) Then tname(kd) = RmListStyle(knd, lvl, why): tname(kd & "?") = why
                                        target = tname(kd): why = tname(kd & "?")
                                        If Len(target) > 0 Then
                                            If SKW_GetStyle(d, target) Is Nothing Then
                                                why = target & " is not in this document (Import Team Styles brings it in)": target = ""
                                            Else
                                                will = True
                                            End If
                                        End If
                                    Case "table"
                                        If sb Is Nothing Then
                                            why = "no style was given"
                                        Else
                                            target = b
                                            why = "tables are left alone unless this row is applied by its number"
                                            will = (onlyKey = key)
                                        End If
                                    Case "typed"
                                        why = "the bullet is typed text, not a list"
                                End Select
                                ' Picks made in the window stand in for the defaults: a row with no pick is left alone.
                                If Not ov Is Nothing Then
                                    target = "": will = False
                                    If ov.Exists(key) And knd <> "typed" Then
                                        target = ov(key)
                                        will = (Len(target) > 0)
                                        If will Then will = Not SKW_GetStyle(d, target) Is Nothing
                                    End If
                                End If
                                If Not doIt Then
                                    If Not groups.Exists(key) Then groups(key) = Array(0, 0, target, why, will Or (knd = "table" And Len(target) > 0))
                                    g = groups(key)
                                    g(0) = g(0) + 1
                                    If RmHasDirect(p, sa) Then g(1) = g(1) + 1
                                    groups(key) = g
                                    If will Then n = n + 1
                                ElseIf will And (Len(onlyKey) = 0 Or onlyKey = key) Then
                                    If selA < 0 Or (story.StoryType = selStory And p.Range.Start < selB And p.Range.End > selA) Then
                                        Set st = SKW_GetStyle(d, target)
                                        nt = RmSet(story, p, sa, st, knd, keep)
                                        If Len(nt) > 0 And InStr(notes, nt) = 0 Then notes = notes & "; " & nt
                                        n = n + 1
                                    End If
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        End If
    Next
    RmWalk = n
End Function

' A paragraph by its start in a story that is not the main text.
Private Function RmParaAt(ByVal story As Object, ByVal pos As Long) As Object
    Dim r As Object
    On Error Resume Next
    Set r = story.Duplicate
    r.SetRange pos, pos
    Set RmParaAt = r.Paragraphs(1)
End Function

' Word moves text from Normal to List Paragraph when someone clicks a bullet or number on it. When Normal is
' being changed and List Paragraph is not, say how many list paragraphs sit there.
Private Function RmListParaHint(ByVal d As Object, ByVal pairs As String) As String
    Dim pr As Variant, nm As String, lp As String, hasNormal As Boolean, hasLp As Boolean, st As Object, n As Long, story As Object
    On Error Resume Next
    nm = d.Styles(wdStyleNormal).NameLocal
    lp = d.Styles(wdStyleListParagraph).NameLocal
    For Each pr In Split(pairs, "|")
        If LCase$(Trim$(Split(pr & ">", ">")(0))) = LCase$(nm) Then hasNormal = True
        If LCase$(Trim$(Split(pr & ">", ">")(0))) = LCase$(lp) Then hasLp = True
    Next
    If Not hasNormal Or hasLp Or Len(lp) = 0 Then Exit Function
    Set st = d.Styles(lp)
    For Each story In SKW_Stories(d)
        If story.StoryType <> wdCommentsStory Then n = n + SKW_StyleRanges(story, st).Count
    Next
    If n > 0 Then RmListParaHint = "This document also has text in " & lp & " in " & n & " place(s): the style Word gives text that is bulleted or numbered from " & nm & _
        ". Name " & lp & " as another style to change, and its bullets go to the list styles too."
End Function

' The key of a numbered row of the last check, or "".
Public Function SKW_RemapRowKey(ByVal n As Long) As String
    If n >= 1 And n <= mRmCount Then SKW_RemapRowKey = mRmKeys(n)
End Function

' pairs: "From>To|From>To". To may be blank: paragraphs not in a list are then left alone.
Public Function SKW_RemapCheck(ByVal pairs As String) As String
    Dim d As Object, msg As String, rows As New Collection, groups As Object, skipped As String, notes As String, n As Long
    Dim k As Variant, a() As String, g As Variant, i As Long, wasSaved As Boolean, alone As Long, hint As String
    Set SKW_LastReport = Nothing
    SKW_Fixable = 0
    mRmCount = 0
    msg = SKW_Begin("Remap style: reading", False)
    If Len(msg) > 0 Then SKW_RemapCheck = msg: Exit Function
    Set d = ActiveDocument
    wasSaved = d.Saved
    Set groups = CreateObject("Scripting.Dictionary")
    n = RmWalk(d, pairs, False, "all", "", -1, -1, 0, groups, skipped, notes)
    hint = RmListParaHint(d, pairs)
    If wasSaved And Not d.Saved Then d.Saved = True
    SKW_Finish
    If groups.Count > 0 Then ReDim mRmKeys(1 To groups.Count)
    For Each k In groups.Keys
        i = i + 1
        mRmKeys(i) = k
        a = Split(k, "|")
        g = groups(k)
        If a(1) = "char" Then n = n + g(0)
        If Len(g(2)) = 0 Or (a(1) = "table") Then alone = alone + g(0)
        rows.Add SKW_Row(i, a(0), IIf(a(1) = "char", "Character style", RmKindWords(a(1))), IIf(Val(a(2)) > 0, a(2), ""), g(0), g(1), _
            IIf(Len(g(2)) > 0 And a(1) <> "table", g(2), "Left alone"), IIf(a(1) = "table" And Len(g(2)) > 0, "Would become " & g(2) & ": ", "") & g(3))
    Next
    mRmCount = i
    SKW_Fixable = n
    msg = n & " paragraph(s) would change style and " & alone & " would be left alone, in " & i & " row(s)."
    If i = 0 Then msg = "Nothing in this document is in the style(s) named."
    If Len(skipped) > 0 Then msg = msg & " Skipped: " & Mid$(skipped, 3) & "."
    If Len(hint) > 0 Then msg = msg & " " & hint
    If i > 0 Then
        SKW_Report "Remap Style", msg & " A paragraph bulleted or numbered by hand goes to the list style the Global Standards' List numbering sheet gives its level, " & _
            "and is left alone when there is none: a bullet is never lost. Apply every row, one row by its number, or the selection. Changes nothing.", _
            Array("Row", "Style now", "What the paragraph is", "List level", "Paragraphs", "With formatting by hand", "Becomes", "Note"), rows
    End If
    SKW_RemapCheck = msg
End Function

' keep: "all", "text" (emphasis only) or "none". scope: "all" or "selection". onlyKey: "" or one row's key.
' targets: "" for the defaults, or the picks made in the window, one per line: a row's key, a tab, the style it becomes.
Public Function SKW_RemapApply(ByVal pairs As String, ByVal keep As String, ByVal scope As String, ByVal onlyKey As String, Optional ByVal targets As String = "") As String
    Dim d As Object, msg As String, skipped As String, notes As String, n As Long, selA As Long, selB As Long, selStory As Long
    Dim ov As Object, ln As Variant
    If Len(targets) > 0 Then
        Set ov = CreateObject("Scripting.Dictionary")
        For Each ln In Split(targets, vbLf)
            If InStr(ln, vbTab) > 0 Then ov(Split(ln, vbTab)(0)) = Trim$(Split(ln, vbTab)(1))
        Next
    End If
    Set SKW_LastReport = Nothing
    If Documents.Count = 0 Then SKW_RemapApply = "Open a document first.": Exit Function
    Set d = ActiveDocument
    selA = -1: selB = -1
    If LCase$(scope) = "selection" Then
        selA = Selection.Range.Start: selB = Selection.Range.End: selStory = Selection.Range.StoryType
        If selB = selA Then selB = selA + 1
    End If
    keep = LCase$(keep)
    If keep <> "text" And keep <> "none" Then keep = "all"
    msg = SKW_Begin("Remap style", True)
    If Len(msg) > 0 Then SKW_RemapApply = msg: Exit Function
    On Error GoTo Fail
    n = RmWalk(d, pairs, True, keep, onlyKey, selA, selB, selStory, Nothing, skipped, notes, ov)
    SKW_Finish
    SKW_RemapApply = "Remapped " & n & " paragraph(s)" & IIf(Len(onlyKey) > 0, " in the row " & Replace(onlyKey, "|", ", "), "") & _
        IIf(selA >= 0, " in the selection", "") & ". Formatting by hand: " & _
        IIf(keep = "all", "kept", IIf(keep = "text", "emphasis kept; fonts, sizes, indents and spacing take the style's", "cleared")) & "." & _
        IIf(Len(notes) > 0, vbCr & "Note: " & Mid$(notes, 3) & ".", "") & _
        IIf(Len(skipped) > 0, vbCr & "Skipped: " & Mid$(skipped, 3) & ".", "") & vbCr & "One Ctrl+Z undoes it."
    Exit Function
Fail:
    msg = Err.Description
    SKW_Finish
    SKW_RemapApply = "Stopped: " & msg & ". Press Ctrl+Z to undo what was done."
End Function

' For the Remap Style window: each kind of paragraph in one style.
' Each item: Array(key, what it is, list level, paragraphs, with formatting by hand, the Global Standards' style for it or "", why not).
Public Function SKW_RemapRows(ByVal fromStyle As String) As Collection
    Dim d As Object, rows As New Collection, groups As Object, skipped As String, notes As String, k As Variant, a() As String, g As Variant
    Dim wasSaved As Boolean, msg As String
    Set SKW_RemapRows = rows
    If Documents.Count = 0 Or Len(Trim$(fromStyle)) = 0 Then Exit Function
    msg = SKW_Begin("Remap style: reading", False)
    If Len(msg) > 0 Then Exit Function
    Set d = ActiveDocument
    wasSaved = d.Saved
    Set groups = CreateObject("Scripting.Dictionary")
    RmWalk d, Trim$(fromStyle) & ">", False, "all", "", -1, -1, 0, groups, skipped, notes
    If wasSaved And Not d.Saved Then d.Saved = True
    SKW_Finish
    For Each k In groups.Keys
        a = Split(k, "|")
        g = groups(k)
        If a(1) <> "char" Then rows.Add Array(CStr(k), RmKindWords(a(1)), Val(a(2)), g(0), g(1), IIf(a(1) = "bullet" Or a(1) = "number", g(2), ""), g(3))
    Next
End Function

' For the Remap Style window: paragraph style names in the document, sorted. usedOnly: only those some text is in.
' With usedOnly False the list styles the Global Standards name are added, so they can be picked before they are used,
' and allBuiltIn adds Word's built-in styles the document has not used yet (Body Text, List Bullet 2 and the like).
Public Function SKW_RemapStyleNames(ByVal usedOnly As Boolean, Optional ByVal allBuiltIn As Boolean = False) As Collection
    Dim d As Object, st As Object, c As New Collection, story As Object, hit As Boolean, r As Object, nm As String, wasSaved As Boolean
    Set SKW_RemapStyleNames = c
    If Documents.Count = 0 Then Exit Function
    Set d = ActiveDocument
    wasSaved = d.Saved
    On Error Resume Next
    For Each st In d.Styles
        If (st.InUse Or (allBuiltIn And Not usedOnly)) And (st.Type = wdStyleTypeParagraph Or st.Type = 6) Then
            hit = True
            If usedOnly Then
                hit = False
                For Each story In SKW_Stories(d)
                    If story.StoryType <> wdCommentsStory Then
                        If SKW_StyleRanges(story, st, 1).Count > 0 Then hit = True: Exit For
                    End If
                Next
            End If
            If hit Then RmSorted c, st.NameLocal
        End If
    Next
    If Not usedOnly Then
        For Each r In SKW_WbRows("List numbering")
            nm = ""
            If r.Exists("style") Then nm = Trim$(r("style"))
            If Len(nm) > 0 Then
                If Not SKW_GetStyle(d, nm) Is Nothing Then RmSorted c, nm
            End If
        Next
    End If
    If wasSaved And Not d.Saved Then d.Saved = True
End Function

Private Sub RmSorted(ByVal c As Collection, ByVal nm As String)
    Dim i As Long, k As Long
    For i = 1 To c.Count
        k = StrComp(c(i), nm, vbTextCompare)
        If k = 0 Then Exit Sub
        If k > 0 Then c.Add nm, , i: Exit Sub
    Next
    c.Add nm
End Sub

' pairs: "From>To|From>To". Every row, formatting by hand kept. Kept for callers from before 1.14.0.
Public Function SKW_RemapStyle(ByVal pairs As String) As String
    SKW_RemapStyle = SKW_RemapApply(pairs, "all", "all", "")
End Function

' ---------------------------------------------------------------------
' Fonts and unused styles
' ---------------------------------------------------------------------

' Deletes custom styles nobody uses: not applied anywhere, not on the team's allowed list, and not the base or next style of another style. Lists what it deleted. One Ctrl+Z undoes it.
Public Sub RemoveUnusedCustomStyles()
    If Not SKW_IsQuiet() Then
        If MsgBox("Delete the custom styles that are not used in " & ActiveDocument.Name & "?" & vbCr & vbCr & _
                  "Styles on the team's allowed list, built-in styles and styles other styles depend on are kept. One Ctrl+Z undoes it.", _
                  vbOKCancel + vbQuestion, SKW_APP) <> vbOK Then Exit Sub
    End If
    SKW_Show SKW_RemoveUnusedCustomStyles()
End Sub

Public Function SKW_FontNotAllowed() As String
    Dim d As Object, msg As String, rows As New Collection, allowed As New Collection, r As Object, story As Object, p As Object
    Dim fonts As Object, nm As String, w As Object, k As Variant, e As Variant, nBad As Long, n As Long
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Fonts not allowed", False)
    If Len(msg) > 0 Then SKW_FontNotAllowed = msg: Exit Function
    Set d = ActiveDocument
    For Each r In SKW_Rows("Fonts")
        If r.Exists("font") Then
            If Len(r("font")) > 0 And (Not r.Exists("allowed") Or SKW_IsYes(IIf(r.Exists("allowed"), r("allowed"), "Yes"))) Then allowed.Add r("font")
        End If
    Next
    Set fonts = CreateObject("Scripting.Dictionary")
    For Each story In SKW_Stories(d)
        If story.StoryType <> wdCommentsStory Then
            SKW_Steps story.Paragraphs.Count, "Paragraph"
            For Each p In story.Paragraphs
                SKW_Step
                If Len(p.Range.Text) > 1 Then
                    nm = FontNameOf(d, p.Range.Font.Name)
                    If Len(nm) > 0 Then
                        NoteFont fonts, nm, story, p.Range
                    Else
                        n = 0
                        For Each w In p.Range.Words
                            n = n + 1
                            If n > 300 Then Exit For
                            nm = FontNameOf(d, w.Font.Name)
                            If Len(nm) > 0 Then NoteFont fonts, nm, story, w
                        Next
                    End If
                End If
            Next
        End If
    Next
    For Each k In fonts.Keys
        e = fonts(k)
        If allowed.Count > 0 And Not SKW_InList(allowed, CStr(k)) Then
            nBad = nBad + 1
            rows.Add SKW_Row("Not allowed", k, e(0), e(1), e(2), e(3))
        Else
            rows.Add SKW_Row(IIf(allowed.Count > 0, "Allowed", "Used"), k, e(0), e(1), e(2), e(3))
        End If
    Next
    SKW_Finish
    SKW_Report "Font Not Allowed", fonts.Count & " font(s) used, " & nBad & " not on the team list" & _
        IIf(allowed.Count = 0, " (no Fonts list in the standards workbook, so none are flagged)", " (" & allowed.Count & " allowed)") & _
        ". Theme fonts are shown by their real names. Changes nothing.", Array("Rule", "Font", "Paragraphs", "Where first", "Page", "Text"), rows
    SKW_FontNotAllowed = fonts.Count & " font(s) used, " & nBad & " not allowed."
End Function

Private Sub NoteFont(ByVal fonts As Object, ByVal nm As String, ByVal story As Object, ByVal r As Object)
    Dim e As Variant
    If fonts.Exists(nm) Then
        e = fonts(nm)
        e(0) = e(0) + 1
        fonts(nm) = e
    Else
        fonts(nm) = Array(1, SKW_StoryName(story.StoryType), SKW_Page(r), SKW_Excerpt(r.Text, 60))
    End If
End Sub

Private Sub KeepRef(ByVal keep As Object, ByVal o As Object, ByVal ref As Variant)
    Dim nm As String
    On Error Resume Next
    nm = ref.NameLocal
    If Err.Number <> 0 Then nm = CStr(ref)
    If Len(nm) > 0 And LCase$(nm) <> LCase$(o.NameLocal) Then keep(nm) = True
End Sub

Public Function SKW_RemoveUnusedCustomStyles() As String
    Dim d As Object, msg As String, st As Object, allowed As Collection, keep As Object, used As Boolean, story As Object
    Dim t As Object, o As Object, del As New Collection, names As String, n As Long, v As Variant, prefix As String, tn As String
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Remove unused custom styles", True)
    If Len(msg) > 0 Then SKW_RemoveUnusedCustomStyles = msg: Exit Function
    Set d = ActiveDocument
    Set allowed = SKW_Column("Allowed styles", "Style")
    prefix = SKW_Setting("StylePrefix", "")
    ' Styles other styles depend on, and table styles in use, are kept.
    Set keep = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    SKW_Steps d.Styles.Count, "Style"
    For Each o In d.Styles
        SKW_Step
        ' A style naming itself (its own next style, say) does not count as depending on it.
        KeepRef keep, o, o.BaseStyle
        KeepRef keep, o, o.NextParagraphStyle
        ' LinkStyle answers with an unrelated style when a style is not linked, so ask only linked ones.
        If o.Linked Then KeepRef keep, o, o.LinkStyle
    Next
    SKW_Steps d.Tables.Count, "Table"
    For Each t In d.Tables
        SKW_Step
        tn = "": tn = t.Style.NameLocal
        keep(tn) = True
    Next
    Err.Clear
    On Error GoTo Fail
    SKW_Steps d.Styles.Count, "Style"
    For Each st In d.Styles
        SKW_Step
        If Not st.BuiltIn And st.Type <> wdStyleTypeList Then
            If Not keep.Exists(st.NameLocal) And Not SKW_InList(allowed, st.NameLocal) And _
               Not (Len(prefix) > 0 And LCase$(Left$(st.NameLocal, Len(prefix))) = LCase$(prefix)) Then
                used = False
                If st.Type <> wdStyleTypeTable Then
                    For Each story In SKW_Stories(d)
                        If SKW_StyleRanges(story, st, 1).Count > 0 Then used = True: Exit For
                    Next
                End If
                If Not used Then del.Add st
            End If
        End If
    Next
    For Each v In del
        names = names & ", " & v.NameLocal
        v.Delete
        n = n + 1
    Next
    SKW_Finish
    SKW_RemoveUnusedCustomStyles = "Deleted " & n & " unused custom style(s)" & IIf(n > 0, ": " & Mid$(names, 3), "") & "." & vbCr & _
        "Kept built-in styles, styles on the team list, styles in use and styles others depend on. One Ctrl+Z undoes it."
    Exit Function
Fail:
    msg = Err.Description
    SKW_Finish
    SKW_RemoveUnusedCustomStyles = "Stopped after " & n & " style(s): " & msg & ". Press Ctrl+Z to undo what was done."
End Function

' =====================================================================
' Delete Unused Styles (1.8.1)
' Clears the Styles pane down to what the document uses. A custom style
' nobody uses is deleted. Word does not allow a built-in style to be
' deleted, so an unused built-in style that shows in the pane is hidden
' until it is used. Kept: styles in use, styles a kept style depends on,
' the team's styles (list or prefix), Normal and Heading 1 to 9.
' =====================================================================

' Lists the styles the document does not use: custom ones to delete and built-in ones to hide from the Styles pane (Word does not let a built-in style be deleted). Changes nothing, then asks before doing it. One Ctrl+Z undoes it.
Public Sub DeleteUnusedStyles()
    Dim d As Object, s As String
    If SKW_IsQuiet() Then Exit Sub
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    Set d = ActiveDocument
    SKW_Fixable = 0
    s = SKW_DeleteUnusedStyles("")
    If SKW_Fixable <= 0 Or SKW_LastReport Is Nothing Then SKW_Show s: Exit Sub
    On Error Resume Next
    SKW_LastReport.Activate
    On Error GoTo 0
    If MsgBox(Split(s & vbCr, vbCr)(0) & vbCr & vbCr & "Do it now in " & d.Name & "?" & vbCr & _
              "One Ctrl+Z undoes it. The report stays open.", vbYesNo + vbQuestion, SKW_APP) <> vbYes Then Exit Sub
    d.Activate
    SKW_Say SKW_DeleteUnusedStyles("fix")
End Sub

' A style that cannot be looked for counts as in use, so it is left alone.
Private Function StyleInUse(ByVal d As Object, ByVal st As Object) As Boolean
    Dim story As Object, n As Long
    On Error Resume Next
    For Each story In SKW_Stories(d)
        n = -1
        n = SKW_StyleRanges(story, st, 1).Count
        If Err.Number <> 0 Or n <> 0 Then
            Err.Clear
            StyleInUse = True
            Exit Function
        End If
    Next
End Function

Private Function StyleTypeName(ByVal t As Long) As String
    Select Case t
        Case wdStyleTypeParagraph: StyleTypeName = "Paragraph"
        Case wdStyleTypeCharacter: StyleTypeName = "Character"
        Case wdStyleTypeTable: StyleTypeName = "Table"
        Case wdStyleTypeList: StyleTypeName = "List"
        Case Else: StyleTypeName = "Paragraph and character"
    End Select
End Function

' The styles other styles lean on (base, next paragraph, linked) and the table styles in use.
Private Function StyleDependencies(ByVal d As Object) As Object
    Dim keep As Object, o As Object, t As Object, tn As String
    Set keep = CreateObject("Scripting.Dictionary")
    keep.CompareMode = 1
    On Error Resume Next
    For Each o In d.Styles
        KeepRef keep, o, o.BaseStyle
        KeepRef keep, o, o.NextParagraphStyle
        If o.Linked Then KeepRef keep, o, o.LinkStyle
    Next
    For Each t In d.Tables
        tn = "": tn = t.Style.NameLocal
        If Len(tn) > 0 Then keep(tn) = True
    Next
    Set StyleDependencies = keep
End Function

' how = "": report only. how = "fix": delete and hide.
Public Function SKW_DeleteUnusedStyles(ByVal how As String) As String
    Dim d As Object, msg As String, st As Object, allowed As Collection, keep As Object, prefix As String, nm As String
    Dim rows As New Collection, del As Collection, toHide As New Collection, v As Variant, pass As Long, isTeam As Boolean
    Dim nDel As Long, nHide As Long, nTeam As Long, nLean As Long, doIt As Boolean, bi As Boolean, tp As Long, vis As Boolean, total As Long, i As Long
    Dim failed As Object, nGone As Long
    Set failed = CreateObject("Scripting.Dictionary")
    failed.CompareMode = 1
    Set SKW_LastReport = Nothing
    SKW_Fixable = 0
    doIt = (LCase$(how) = "fix")
    msg = SKW_Begin("Delete unused styles", doIt)
    If Len(msg) > 0 Then SKW_DeleteUnusedStyles = msg: Exit Function
    Set d = ActiveDocument
    Set allowed = SKW_Column("Allowed styles", "Style")
    prefix = SKW_Setting("StylePrefix", "")
    On Error GoTo Fail
    ' Custom styles: pass after pass, because deleting a style frees the style it was based on.
    For pass = 1 To 8
        Set keep = StyleDependencies(d)
        Set del = New Collection
        total = d.Styles.Count
        i = 0
        For Each st In d.Styles
            i = i + 1
            SKW_Status "Reading style " & i & " of " & total & IIf(pass > 1, " (pass " & pass & ")", "") & "..."
            bi = True: tp = 0: nm = ""
            On Error Resume Next
            bi = st.BuiltIn: tp = st.Type: nm = st.NameLocal
            On Error GoTo Fail
            If Not bi And tp <> wdStyleTypeList And Len(nm) > 0 Then
                isTeam = SKW_InList(allowed, nm) Or (Len(prefix) > 0 And LCase$(Left$(nm, Len(prefix))) = LCase$(prefix))
                If keep.Exists(nm) Then
                    If pass = 1 And Not doIt Then
                        If tp = wdStyleTypeTable Then
                            ' a table style in use: nothing to say
                        ElseIf Not StyleInUse(d, st) Then
                            nLean = nLean + 1
                            rows.Add SKW_Row(nm, "Custom", StyleTypeName(tp), "Kept for now: another style is based on it or follows it. It goes once that style is gone")
                        End If
                    End If
                ElseIf isTeam Then
                    If pass = 1 Then
                        If tp = wdStyleTypeTable Then
                            nTeam = nTeam + 1
                            If Not doIt Then rows.Add SKW_Row(nm, "Custom", StyleTypeName(tp), "Kept: a team style, not used")
                        ElseIf Not StyleInUse(d, st) Then
                            nTeam = nTeam + 1
                            If Not doIt Then rows.Add SKW_Row(nm, "Custom", StyleTypeName(tp), "Kept: a team style, not used")
                        End If
                    End If
                ElseIf failed.Exists(nm) Then
                    ' Word would not delete it on an earlier pass
                ElseIf tp = wdStyleTypeTable Then
                    del.Add st
                ElseIf Not StyleInUse(d, st) Then
                    del.Add st
                End If
            End If
        Next
        If del.Count = 0 Then Exit For
        nGone = 0
        For Each v In del
            If doIt Then
                nm = v.NameLocal
                On Error Resume Next
                v.Delete
                If Err.Number <> 0 Then
                    failed(nm) = True
                    Err.Clear
                Else
                    nDel = nDel + 1: nGone = nGone + 1
                End If
                On Error GoTo Fail
            Else
                nDel = nDel + 1
                rows.Add SKW_Row(v.NameLocal, "Custom", StyleTypeName(v.Type), "Delete")
            End If
        Next
        If Not doIt Or nGone = 0 Then Exit For
    Next
    ' Built-in styles that show in the pane and are not used: hidden until used.
    total = d.Styles.Count
    i = 0
    For Each st In d.Styles
        i = i + 1
        SKW_Status "Reading built-in style " & i & " of " & total & "..."
        bi = False: tp = 0: nm = "": vis = False
        On Error Resume Next
        bi = st.BuiltIn: tp = st.Type: nm = st.NameLocal
        vis = Not st.Visibility
        On Error GoTo Fail
        If bi And vis And tp <> wdStyleTypeList And tp <> wdStyleTypeTable And Len(nm) > 0 Then
            If LCase$(nm) <> "normal" And Not (LCase$(nm) Like "heading [1-9]") And Not SKW_InList(allowed, nm) And Not keep.Exists(nm) Then
                If Not StyleInUse(d, st) Then
                    nHide = nHide + 1
                    If doIt Then
                        toHide.Add st
                    Else
                        rows.Add SKW_Row(nm, "Built-in", StyleTypeName(tp), "Hide from the Styles pane until it is used (Word does not delete built-in styles)")
                    End If
                End If
            End If
        End If
    Next
    If doIt Then
        On Error Resume Next
        For Each v In toHide
            v.QuickStyle = False
            v.UnhideWhenUsed = True
            v.Visibility = True
        Next
        On Error GoTo Fail
    End If
    SKW_Finish
    If doIt Then
        SKW_DeleteUnusedStyles = "Deleted " & nDel & " unused custom style(s) and hid " & nHide & " unused built-in style(s) from the Styles pane." & vbCr & _
            "Kept: styles in use, styles they depend on, " & nTeam & " unused team style(s), Normal and Heading 1 to 9." & _
            IIf(failed.Count > 0, " Word would not delete " & failed.Count & ": " & Join(failed.Keys, ", ") & ".", "") & " One Ctrl+Z undoes it."
    Else
        SKW_Report "Delete Unused Styles", nDel & " custom style(s) to delete, " & nHide & " built-in style(s) to hide, " & nTeam & _
            " unused team style(s) kept" & IIf(nLean > 0, ", " & nLean & " that go once the style leaning on them is gone", "") & _
            ". Word does not let a built-in style be deleted, so an unused one is hidden from the Styles pane until it is used. " & _
            "Kept: styles in use, styles they depend on, the team's styles (the Allowed styles list and the StylePrefix setting), Normal and Heading 1 to 9. Changes nothing.", _
            Array("Style", "Kind", "Type", "What happens"), rows
        SKW_Fixable = nDel + nHide
        SKW_DeleteUnusedStyles = nDel & " custom style(s) to delete, " & nHide & " built-in style(s) to hide, " & nTeam & " unused team style(s) kept."
    End If
    Exit Function
Fail:
    msg = Err.Description
    SKW_Finish
    SKW_DeleteUnusedStyles = "Stopped: " & msg & "." & IIf(doIt, " Press Ctrl+Z to undo what was done.", "")
End Function
