Attribute VB_Name = "SKW_Styles"
' =====================================================================
' Stationkeeper Tools: styles and direct formatting.
'   StyleInspector                   what styles the cursor is in, and what is direct formatting
'   NonAllowedStyleReport            report: styles in use that are not on the team list
'   DirectFormattingReport           report: text formatted by hand instead of by style
'   ClearDirectFormattingSelection   clears hand formatting in the selection, keeps emphasis
'   ImportTeamStyles                 brings in the team's style definitions
'   RemapStyle                       changes every use of one style to another
'   FontNotAllowed                   report: fonts used that are not on the team list
'   RemoveUnusedCustomStyles         deletes custom styles nobody uses
' =====================================================================
Option Explicit

Private Const UNDEF As Long = 9999999

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
    For Each st In d.Styles
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
    For Each t In d.Tables
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
            For Each p In story.Paragraphs
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

' Changes every paragraph (or text) in one style to another style, or applies every pair on the Style map sheet of the standards workbook. One Ctrl+Z undoes it.
Public Sub RemapStyle()
    Dim pairs As String, r As Object, n As Long, a As String, b As String, ans As VbMsgBoxResult
    If SKW_IsQuiet() Or Documents.Count = 0 Then Exit Sub
    For Each r In SKW_Rows("Style map")
        If r.Exists("from") And r.Exists("to") Then
            If Len(r("from")) > 0 And Len(r("to")) > 0 Then pairs = pairs & "|" & r("from") & ">" & r("to"): n = n + 1
        End If
    Next
    If n > 0 Then
        ans = MsgBox("The standards workbook maps " & n & " style(s) to others." & vbCr & vbCr & "Yes: apply all of them" & vbCr & _
                     "No: change one style now", vbYesNoCancel + vbQuestion, "Remap Style")
        If ans = vbCancel Then Exit Sub
        If ans = vbYes Then SKW_Show SKW_RemapStyle(Mid$(pairs, 2)): Exit Sub
    End If
    a = InputBox("Change every use of which style?", "Remap Style", SKW_ParaStyleName(Selection.Range))
    If Len(a) = 0 Then Exit Sub
    b = InputBox("Change " & a & " to which style?", "Remap Style", "Normal")
    If Len(b) = 0 Then Exit Sub
    SKW_Show SKW_RemapStyle(a & ">" & b)
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
    For Each st In d.Styles
        If st.InUse Then before(st.NameLocal) = StyleSig(st)
    Next
    On Error GoTo Fail
    d.CopyStylesFromTemplate source
    On Error GoTo 0
    For Each st In d.Styles
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

' pairs: "From>To|From>To"
Public Function SKW_RemapStyle(ByVal pairs As String) As String
    Dim d As Object, msg As String, pr As Variant, a As String, b As String, sa As Object, sb As Object
    Dim story As Object, f As Object, n As Long, total As Long, out As String, skipped As String
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Remap style", True)
    If Len(msg) > 0 Then SKW_RemapStyle = msg: Exit Function
    Set d = ActiveDocument
    On Error GoTo Fail
    For Each pr In Split(pairs, "|")
        If InStr(pr, ">") > 0 Then
            a = Trim$(Split(pr, ">")(0)): b = Trim$(Split(pr, ">")(1))
            Set sa = SKW_GetStyle(d, a): Set sb = SKW_GetStyle(d, b)
            If sa Is Nothing Or sb Is Nothing Then
                skipped = skipped & "; " & a & " to " & b & " (" & IIf(sa Is Nothing, a, b) & " is not in this document)"
            Else
                n = 0
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
                total = total + n
                out = out & "; " & a & " to " & b & ": " & n & " place(s)"
            End If
        End If
    Next
    SKW_Finish
    SKW_RemapStyle = "Remapped " & total & " place(s)" & IIf(Len(out) > 0, ": " & Mid$(out, 3), "") & "." & _
        IIf(Len(skipped) > 0, vbCr & "Skipped: " & Mid$(skipped, 3) & ".", "") & vbCr & "One Ctrl+Z undoes it."
    Exit Function
Fail:
    msg = Err.Description
    SKW_Finish
    SKW_RemapStyle = "Stopped: " & msg & ". Press Ctrl+Z to undo what was done."
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
            For Each p In story.Paragraphs
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
    For Each o In d.Styles
        ' A style naming itself (its own next style, say) does not count as depending on it.
        KeepRef keep, o, o.BaseStyle
        KeepRef keep, o, o.NextParagraphStyle
        ' LinkStyle answers with an unrelated style when a style is not linked, so ask only linked ones.
        If o.Linked Then KeepRef keep, o, o.LinkStyle
    Next
    For Each t In d.Tables
        tn = "": tn = t.Style.NameLocal
        keep(tn) = True
    Next
    Err.Clear
    On Error GoTo Fail
    For Each st In d.Styles
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
