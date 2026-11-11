Attribute VB_Name = "SKW_Tables"
' =====================================================================
' Stationkeeper Tools: tables.
'   TableStandardsAudit   report: tables that do not follow the team table types
'   ApplyTeamTableType    sets the table at the cursor to a team table type
' Table types come from the "Table types" sheet of the standards workbook.
' =====================================================================
Option Explicit

' Checks every table against the team table types: table style, header row repeating on each page, rows breaking across pages, width, floating tables, tables inside tables, header and body styles, and a caption above or below. Changes nothing.
Public Sub TableStandardsAudit()
    SKW_Show SKW_TableStandardsAudit()
End Sub

' Sets the table at the cursor to a team table type: table style, header row that repeats, rows kept whole, width, cell styles and padding. Asks which type when there are several. One Ctrl+Z undoes it.
Public Sub ApplyTeamTableType()
    Dim types As Collection, s As String, i As Long, pick As String, nm As String
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    If Not Selection.Information(wdWithInTable) Then SKW_Say "Put the cursor in a table first.": Exit Sub
    Set types = SKW_TableTypes()
    If types.Count = 1 Or SKW_IsQuiet() Then
        nm = types(1)("type")
    Else
        For i = 1 To types.Count
            s = s & i & "  " & types(i)("type") & IIf(Len(types(i)("table style")) > 0, "  (" & types(i)("table style") & ")", "") & vbCr
        Next
        pick = InputBox("Which table type?" & vbCr & vbCr & s & vbCr & "Type the number.", "Apply Team Table Type", "1")
        If Len(pick) = 0 Then Exit Sub
        If Not IsNumeric(pick) Then SKW_Say "Type a number from the list.": Exit Sub
        If CLng(pick) < 1 Or CLng(pick) > types.Count Then SKW_Say "Type a number from the list.": Exit Sub
        nm = types(CLng(pick))("type")
    End If
    SKW_Show SKW_ApplyTeamTableType(nm)
End Sub

' ---------------------------------------------------------------------
' Table types: the workbook's sheet, or one built-in "Standard" type.
' Each is a dictionary with lower-case keys.
' ---------------------------------------------------------------------
Public Function SKW_TableTypes() As Collection
    Dim c As New Collection, r As Object, d As Object, k As Variant
    Dim keys As Variant
    keys = Array("type", "table style", "header row repeats", "rows may break across pages", "width percent", _
                 "header row style", "body style", "notes row style", "has notes row", "cell padding")
    For Each r In SKW_Rows("Table types")
        If r.Exists("type") Then
            If Len(r("type")) > 0 Then
                Set d = CreateObject("Scripting.Dictionary")
                For Each k In keys
                    If r.Exists(k) Then d(k) = r(k) Else d(k) = ""
                Next
                c.Add d
            End If
        End If
    Next
    If c.Count = 0 Then
        Set d = CreateObject("Scripting.Dictionary")
        For Each k In keys
            d(k) = ""
        Next
        d("type") = "Standard"
        d("table style") = "Table Grid"
        d("header row repeats") = "Yes"
        d("rows may break across pages") = "No"
        d("width percent") = SKW_Setting("TableWidthPercent", "100")
        c.Add d
    End If
    Set SKW_TableTypes = c
End Function

Private Function TypeFor(ByVal types As Collection, ByVal styleName As String) As Object
    Dim t As Object
    For Each t In types
        If LCase$(t("table style")) = LCase$(styleName) And Len(styleName) > 0 Then Set TypeFor = t: Exit Function
    Next
End Function

Private Function TableStyleName(ByVal t As Object) As String
    On Error Resume Next
    TableStyleName = t.Style.NameLocal
End Function

Private Function FirstCellText(ByVal t As Object) As String
    On Error Resume Next
    FirstCellText = t.Range.Cells(1).Range.Text
End Function

' Width of the first row in points (0 if it cannot be measured).
Private Function RowOneWidth(ByVal t As Object) As Single
    Dim c As Object, w As Single
    On Error GoTo Done
    For Each c In t.Range.Cells
        If c.RowIndex > 1 Then Exit For
        w = w + c.Width
    Next
    RowOneWidth = w
Done:
End Function

Private Function TextWidth(ByVal t As Object) As Single
    On Error Resume Next
    With t.Range.Sections(1).PageSetup
        TextWidth = .PageWidth - .LeftMargin - .RightMargin - .Gutter
    End With
End Function

Private Function HasCaption(ByVal t As Object, ByVal capStyle As String, ByVal inFirstRow As Boolean) As Boolean
    Dim r As Object
    On Error Resume Next
    Set r = t.Range.Duplicate
    r.Collapse wdCollapseStart
    r.Move wdParagraph, -1
    r.Expand wdParagraph
    If IsCaption(r, capStyle) Then HasCaption = True: Exit Function
    Set r = t.Range.Duplicate
    r.Collapse wdCollapseEnd
    r.Expand wdParagraph
    If IsCaption(r, capStyle) Then HasCaption = True: Exit Function
    If inFirstRow Then
        If IsCaption(t.Range.Cells(1).Range, capStyle) Then HasCaption = True
    End If
End Function

Private Function IsCaption(ByVal r As Object, ByVal capStyle As String) As Boolean
    Dim f As Object
    On Error Resume Next
    If LCase$(SKW_ParaStyleName(r.Paragraphs(1).Range)) = LCase$(capStyle) Then IsCaption = True: Exit Function
    For Each f In r.Fields
        If f.Type = wdFieldSequence Then IsCaption = True: Exit Function
    Next
End Function

Public Function SKW_TableStandardsAudit() As String
    Dim d As Object, msg As String, types As Collection, rows As New Collection, t As Object, i As Long, ty As Object
    Dim stn As String, pg As String, ft As String, names As String, k As Object, capStyle As String, inFirst As Boolean
    Dim hf As Long, br As Long, w As Single, tw As Single, pct As String, c As Object, n As Long, bad As Long, nBadTables As Long
    Dim onePage As Boolean, before As Long
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Table standards audit", False)
    If Len(msg) > 0 Then SKW_TableStandardsAudit = msg: Exit Function
    Set d = ActiveDocument
    Set types = SKW_TableTypes()
    For Each k In types
        names = names & IIf(Len(names) > 0, ", ", "") & k("type") & IIf(Len(k("table style")) > 0, " (" & k("table style") & ")", "")
    Next
    capStyle = SKW_Setting("CaptionStyle", "Caption")
    inFirst = SKW_IsYes(SKW_Setting("CaptionInFirstRow", "No"))
    For Each t In d.Tables
        i = i + 1
        before = rows.Count
        SKW_Status "Checking table " & i & " of " & d.Tables.Count & "..."
        stn = TableStyleName(t)
        pg = SKW_Page(t.Range)
        ft = FirstCellText(t)
        Set ty = TypeFor(types, stn)
        If ty Is Nothing Then
            rows.Add SKW_Row(i, pg, "Table style", "Style " & IIf(Len(stn) > 0, stn, "(none)") & " is not a team table type. Types: " & names, ft)
            Set ty = types(1)
        End If
        On Error Resume Next
        onePage = (t.Range.Characters.First.Information(wdActiveEndAdjustedPageNumber) = t.Range.Characters.Last.Information(wdActiveEndAdjustedPageNumber))
        hf = 0: hf = t.Rows(1).HeadingFormat
        br = 0: br = t.Rows.AllowBreakAcrossPages
        On Error GoTo 0
        If SKW_IsYes(ty("header row repeats")) And hf = 0 Then
            rows.Add SKW_Row(i, pg, "Header row", "Header row does not repeat at the top of each page" & IIf(onePage, " (the table fits on one page now)", ""), ft)
        End If
        If LCase$(ty("rows may break across pages")) = "no" And br <> 0 Then
            rows.Add SKW_Row(i, pg, "Rows break", "Rows can break across pages", ft)
        End If
        On Error Resume Next
        If t.Rows.WrapAroundText Then rows.Add SKW_Row(i, pg, "Floating", "Text wraps around the table (floating table). It can move away from its text", ft)
        If t.Tables.Count > 0 Then rows.Add SKW_Row(i, pg, "Nested", "Has " & t.Tables.Count & " table(s) inside it", ft)
        On Error GoTo 0
        w = RowOneWidth(t): tw = TextWidth(t)
        If w > 0 And tw > 0 And w > tw + 3 Then
            rows.Add SKW_Row(i, pg, "Width", "Wider than the margins (" & Format$(w / 72, "0.00") & " in, text area " & Format$(tw / 72, "0.00") & " in)", ft)
        End If
        pct = ty("width percent")
        If IsNumeric(pct) And Len(pct) > 0 Then
            If t.PreferredWidthType <> wdPreferredWidthPercent Or Abs(t.PreferredWidth - CSng(pct)) > 1 Then
                If Not (w > 0 And tw > 0 And Abs(w - tw * CSng(pct) / 100) <= 3) Then
                    rows.Add SKW_Row(i, pg, "Width", "Not set to " & pct & "% of the text width", ft)
                End If
            End If
        End If
        If Len(ty("header row style")) + Len(ty("body style")) > 0 Then
            n = 0: bad = 0
            For Each c In t.Range.Cells
                n = n + 1
                If n > 60 Then Exit For
                If c.RowIndex = 1 And Len(ty("header row style")) > 0 Then
                    If LCase$(SKW_ParaStyleName(c.Range.Paragraphs(1).Range)) <> LCase$(ty("header row style")) Then bad = bad + 1
                ElseIf c.RowIndex > 1 And Len(ty("body style")) > 0 Then
                    If Not (SKW_IsYes(ty("has notes row")) And c.RowIndex = t.Rows.Count) Then
                        If LCase$(SKW_ParaStyleName(c.Range.Paragraphs(1).Range)) <> LCase$(ty("body style")) Then bad = bad + 1
                    End If
                End If
            Next
            If bad > 0 Then rows.Add SKW_Row(i, pg, "Cell styles", bad & " cell(s) not in the type's header or body style", ft)
        End If
        If Not HasCaption(t, capStyle, inFirst) Then rows.Add SKW_Row(i, pg, "Caption", "No caption above or below (style " & capStyle & " or a SEQ field)", ft)
        If rows.Count > before Then nBadTables = nBadTables + 1
    Next
    SKW_Finish
    SKW_Report "Table Standards Audit", d.Tables.Count & " table(s) checked, " & nBadTables & " with findings. Team table types: " & names & _
        ". To fix one, put the cursor in it and run Apply Team Table Type. Changes nothing.", _
        Array("Table", "Page", "Rule", "Finding", "First cell"), rows
    SKW_TableStandardsAudit = d.Tables.Count & " table(s) checked, " & nBadTables & " with findings."
End Function

Public Function SKW_ApplyTeamTableType(ByVal typeName As String) As String
    Dim d As Object, msg As String, ty As Object, k As Object, t As Object, done As String, skipped As String
    Dim c As Object, nm As String, pct As String, pad As String, lastRow As Long, notes As Boolean, nm2 As Variant
    Set SKW_LastReport = Nothing
    If Documents.Count > 0 Then
        If Not Selection.Information(wdWithInTable) Then SKW_ApplyTeamTableType = "Put the cursor in a table first.": Exit Function
    End If
    For Each k In SKW_TableTypes()
        If LCase$(k("type")) = LCase$(typeName) Then Set ty = k
    Next
    If ty Is Nothing Then SKW_ApplyTeamTableType = "There is no table type called " & typeName & ".": Exit Function
    msg = SKW_Begin("Apply table type", True)
    If Len(msg) > 0 Then SKW_ApplyTeamTableType = msg: Exit Function
    Set d = ActiveDocument
    Set t = Selection.Tables(1)
    On Error GoTo Fail
    nm = ty("table style")
    If Len(nm) > 0 Then
        If SKW_GetStyle(d, nm) Is Nothing Then
            skipped = skipped & "; table style " & nm & " is not in this document (import the team styles first)"
        Else
            ' Through the range: Table.Style escapes Word's undo grouping, Range.Style does not.
            t.Range.Style = nm
            done = done & "; style " & nm
        End If
    End If
    t.Rows.WrapAroundText = False
    If Len(ty("header row repeats")) > 0 Then
        On Error Resume Next
        t.Rows(1).HeadingFormat = SKW_IsYes(ty("header row repeats"))
        If Err.Number <> 0 Then skipped = skipped & "; header row (merged cells)" Else done = done & "; header row " & IIf(SKW_IsYes(ty("header row repeats")), "repeats", "does not repeat")
        Err.Clear
        On Error GoTo Fail
    End If
    If Len(ty("rows may break across pages")) > 0 Then
        t.Rows.AllowBreakAcrossPages = SKW_IsYes(ty("rows may break across pages"))
        done = done & "; rows " & IIf(SKW_IsYes(ty("rows may break across pages")), "may", "do not") & " break across pages"
    End If
    pct = ty("width percent")
    If IsNumeric(pct) And Len(pct) > 0 Then
        If CSng(pct) = 100 Then
            t.AutoFitBehavior wdAutoFitWindow
        Else
            t.PreferredWidthType = wdPreferredWidthPercent
            t.PreferredWidth = CSng(pct)
        End If
        done = done & "; width " & pct & "%"
    End If
    pad = ty("cell padding")
    If IsNumeric(pad) And Len(pad) > 0 Then
        t.LeftPadding = CSng(pad): t.RightPadding = CSng(pad)
        t.TopPadding = CSng(pad): t.BottomPadding = CSng(pad)
        done = done & "; padding " & pad & " pt"
    End If
    notes = SKW_IsYes(ty("has notes row"))
    lastRow = t.Rows.Count
    If Len(ty("header row style")) + Len(ty("body style")) + Len(ty("notes row style")) > 0 Then
        For Each nm2 In Array("header row style", "body style", "notes row style")
            If Len(ty(nm2)) > 0 Then
                If SKW_GetStyle(d, ty(nm2)) Is Nothing Then skipped = skipped & "; " & nm2 & " " & ty(nm2) & " is not in this document"
            End If
        Next
        For Each c In t.Range.Cells
            If c.RowIndex = 1 Then
                nm = ty("header row style")
            ElseIf notes And c.RowIndex = lastRow Then
                nm = ty("notes row style")
            Else
                nm = ty("body style")
            End If
            If Len(nm) > 0 Then
                If Not SKW_GetStyle(d, nm) Is Nothing Then c.Range.Style = nm
            End If
        Next
        done = done & "; cell styles"
    End If
    SKW_Finish
    If Left$(done, 2) = "; " Then done = Mid$(done, 3)
    If Left$(skipped, 2) = "; " Then skipped = Mid$(skipped, 3)
    SKW_ApplyTeamTableType = "Table set to " & ty("type") & ": " & done & "." & _
        IIf(Len(skipped) > 0, vbCr & "Not set: " & skipped & ".", "") & vbCr & "One Ctrl+Z undoes it."
    Exit Function
Fail:
    msg = Err.Description
    SKW_Finish
    SKW_ApplyTeamTableType = "Stopped: " & msg & ". Press Ctrl+Z to undo what was done."
End Function
