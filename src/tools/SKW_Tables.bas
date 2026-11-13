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
                 "header row style", "body style", "notes row style", "has notes row", "cell padding", _
                 "cell margins", "border style", "border width", "border color")
    If SKW_AppStandards() Then
        ' The Stationkeeper app's workbook: one row per table type, named by its Name column.
        For Each r In SKW_Rows("Table types")
            If AppApproved(r) And Len(AppVal(r, "name")) > 0 Then
                Set d = CreateObject("Scripting.Dictionary")
                For Each k In keys
                    d(k) = ""
                Next
                d("type") = AppVal(r, "name")
                d("label") = AppVal(r, "name (value)")
                d("width percent") = AppVal(r, "width pct")
                d("table align") = AppVal(r, "align")
                d("margins pt") = AppVal(r, "cell margins: top") & ";" & AppVal(r, "cell margins: bottom") & ";" & _
                    AppVal(r, "cell margins: left") & ";" & AppVal(r, "cell margins: right")
                d("roles") = AppVal(r, "roles")
                d("override") = AppVal(r, "override")
                d("caption position") = AppVal(r, "caption position")
                d("app") = "yes"
                c.Add d
            End If
        Next
        If c.Count > 0 Then Set SKW_TableTypes = c: Exit Function
    End If
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

' =====================================================================
' Table roles (1.8.0)
'   TableRolesCheck       report: every table's rows and cells by role; offers to style a copy
'   ApplyTableRolesHere   styles the table at the cursor by role
' A row is a Caption, Header, Divider, Total or Body row. A cell in a body or total row is a
' Number, Short text or Long text cell. Each role's format comes from the "Table roles" sheet
' of the standards workbook (or the effort's table_roles), with built-in values when there is none.
' A table's Alt Text can hold "skip" (leave the table alone) or "type=Name" (use that table type).
' =====================================================================

' Reads every table and lists its rows by role (caption, header, divider, total, body) and its cells by kind (number, short text, long text), with the rows to look at. Changes nothing. Then offers to style a copy of the document.
Public Sub TableRolesCheck()
    Dim s As String
    If SKW_IsQuiet() Then Exit Sub
    SKW_Fixable = 0
    s = SKW_TableRolesCheck()
    If SKW_OfferFix(s, "The fix styles every table by its rows and cells in a copy of the document. Your document is not changed.") Then
        SKW_Show SKW_TableRolesFixCopy("")
    End If
End Sub

' Styles the table at the cursor by role: caption row, header row, divider rows, total rows, number cells, short text and long text. One Ctrl+Z undoes it.
Public Sub ApplyTableRolesHere()
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    If Not Selection.Information(wdWithInTable) Then SKW_Say "Put the cursor in a table first.": Exit Sub
    SKW_Show SKW_ApplyTableRolesHere()
End Sub

' The roles for one table type: built-in values, then the sheet's rows for every type, then the rows for this type.
Public Function SKW_TableRoles(ByVal typeName As String) As Object
    Dim res As Object, r As Object, pass As Long, ty As String, role As String, k As Variant, d As Object
    If SKW_AppStandards() Then Set SKW_TableRoles = AppRoles(typeName): Exit Function
    Set res = CreateObject("Scripting.Dictionary")
    res.CompareMode = 1
    AddRole res, "caption", "", "", "", "", "", ""
    AddRole res, "header", "", "center", "center", "", "yes", "1"
    AddRole res, "divider", "", "left", "center", "", "yes", "merged"
    AddRole res, "total", "", "", "center", "", "yes", "Total;Subtotal;Grand Total"
    AddRole res, "number", "", "right", "center", "", "", "-"
    AddRole res, "short text", "", "center", "center", "", "", "3"
    AddRole res, "long text", "", "left", "top", "", "", ""
    AddRole res, "notes", "", "left", "top", "", "", "Note;Notes;Source;Sources"
    AddRole res, "keep", "", "", "", "", "", ""
    AddRole res, "figure caption", "", "", "", "", "", ""
    AddRole res, "header2", "", "center", "center", "", "yes", ""
    AddRole res, "subtotal", "", "", "center", "", "yes", ""
    res("header2")("absent") = "yes"
    For pass = 1 To 2
        For Each r In SKW_Rows("Table roles")
            If r.Exists("role") Then
                ty = ""
                If r.Exists("type") Then ty = LCase$(Trim$(r("type")))
                role = LCase$(Trim$(r("role")))
                If res.Exists(role) Then
                    If (pass = 1 And (Len(ty) = 0 Or ty = "all")) Or (pass = 2 And Len(ty) > 0 And ty = LCase$(typeName)) Then
                        Set d = res(role)
                        For Each k In Array("paragraph style", "align", "vertical", "shading", "bold", "recognize", "left indent", "tab stops")
                            If r.Exists(k) Then
                                If Len(Trim$(r(k))) > 0 Then d(k) = Trim$(r(k))
                            End If
                        Next
                    End If
                End If
            End If
        Next
    Next
    Set SKW_TableRoles = res
End Function

Private Sub AddRole(ByVal res As Object, ByVal role As String, ByVal sty As String, ByVal al As String, ByVal vert As String, _
                    ByVal shade As String, ByVal bld As String, ByVal recog As String)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("paragraph style") = sty: d("align") = al: d("vertical") = vert
    d("shading") = shade: d("bold") = bld: d("recognize") = recog
    d("left indent") = "": d("tab stops") = ""
    d("absent") = "": d("app") = "": d("repeat") = ""
    Set res(role) = d
End Sub

' True when the text is a figure: digits with the usual signs, or one of the listed stand-ins (a dash).
Private Function IsNumberText(ByVal s As String, ByVal extra As String) As Boolean
    Static re As Object
    Dim v As Variant
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If s = "-" Or s = ChrW$(8211) Or s = ChrW$(8212) Then IsNumberText = True: Exit Function
    For Each v In Split(extra, ";")
        If Len(Trim$(v)) > 0 And LCase$(Trim$(v)) = LCase$(s) Then IsNumberText = True: Exit Function
    Next
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "^[\(\-\+\u2212]?\s?[\$\u20AC\u00A3]?\s?\(?(\d[\d,]*(\.\d+)?|\.\d+)\s?(%|[KkMmBb])?\)?$"
    End If
    IsNumberText = re.Test(s)
End Function

Private Function WordCount(ByVal s As String) As Long
    Dim v As Variant, n As Long
    For Each v In Split(Trim$(s), " ")
        If Len(v) > 0 Then n = n + 1
    Next
    WordCount = n
End Function

Private Function StartsWithWord(ByVal s As String, ByVal words As String) As Boolean
    Dim v As Variant, w As String, nx As String
    s = LCase$(Trim$(s))
    For Each v In Split(words, ";")
        w = LCase$(Trim$(v))
        If Len(w) > 0 And Left$(s, Len(w)) = w Then
            nx = Mid$(s, Len(w) + 1, 1)
            If Len(nx) = 0 Or Not (nx Like "[a-z]") Then StartsWithWord = True: Exit Function
        End If
    Next
End Function

' The table type for a table: named in its Alt Text, or the type with its table style, or the first type.
Private Function TypeOfTable(ByVal t As Object, ByVal types As Collection, ByVal descr As String) As Object
    Dim p As Long, nm As String, k As Object
    p = InStr(descr, "type=")
    If p > 0 Then
        nm = Mid$(descr, p + 5)
        If InStr(nm, ";") > 0 Then nm = Left$(nm, InStr(nm, ";") - 1)
        If InStr(nm, vbCr) > 0 Then nm = Left$(nm, InStr(nm, vbCr) - 1)
        nm = Trim$(nm)
        For Each k In types
            If LCase$(k("type")) = nm Then Set TypeOfTable = k: Exit Function
            If k.Exists("label") Then
                If LCase$(k("label")) = nm And Len(nm) > 0 Then Set TypeOfTable = k: Exit Function
            End If
        Next
    End If
    Set TypeOfTable = TypeFor(types, TableStyleName(t))
    If TypeOfTable Is Nothing Then Set TypeOfTable = types(1)
End Function

' Reads one table: each row's role and each cell's kind. Changes nothing.
' ---------------------------------------------------------------------
' Assigned tables. A row's role is the paragraph style on it (the role's tag style). A table whose
' roles have been set carries a hidden bookmark (_SKA...): from then on its rows are read from their
' styles and never guessed again. The mark is not in the Alt Text, so accessibility text stays clean.
' ---------------------------------------------------------------------
Private Function TableAssigned(ByVal d As Object, ByVal t As Object) As Boolean
    Dim b As Object, sh As Boolean, nm As String
    On Error Resume Next
    sh = d.Bookmarks.ShowHidden
    d.Bookmarks.ShowHidden = True
    For Each b In t.Range.Bookmarks
        nm = ""
        nm = b.Name
        If Left$(nm, 4) = "_SKA" Then TableAssigned = True: Exit For
    Next
    d.Bookmarks.ShowHidden = sh
End Function

Private Sub MarkAssigned(ByVal d As Object, ByVal t As Object)
    Dim nm As String
    If TableAssigned(d, t) Then Exit Sub
    On Error Resume Next
    Randomize
    nm = "_SKA" & Hex$(CLng(Rnd() * 16777215)) & Hex$(CLng(Rnd() * 16777215))
    d.Bookmarks.Add nm, t.Range
End Sub

Private Sub ClearAssigned(ByVal d As Object, ByVal t As Object)
    Dim b As Object, sh As Boolean, names As New Collection, v As Variant, nm As String
    On Error Resume Next
    sh = d.Bookmarks.ShowHidden
    d.Bookmarks.ShowHidden = True
    For Each b In t.Range.Bookmarks
        nm = ""
        nm = b.Name
        If Left$(nm, 4) = "_SKA" Then names.Add nm
    Next
    For Each v In names
        d.Bookmarks(CStr(v)).Delete
    Next
    d.Bookmarks.ShowHidden = sh
End Sub

' Tag style (lower case) to role. Body, number and short text styles all read as a body row.
Private Function RolesByStyle(ByVal roles As Object) As Object
    Dim res As Object, k As Variant, nm As String
    Set res = CreateObject("Scripting.Dictionary")
    For Each k In Array("caption", "header", "header2", "divider", "subtotal", "total", "notes")
        nm = LCase$(Trim$(roles(k)("paragraph style")))
        If Len(nm) > 0 And roles(k)("absent") <> "yes" Then res(nm) = CStr(k)
    Next
    For Each k In Array("number", "short text", "long text")
        nm = LCase$(Trim$(roles(k)("paragraph style")))
        If Len(nm) > 0 And roles(k)("absent") <> "yes" Then res(nm) = "body"
    Next
    Set RolesByStyle = res
End Function

' forced: row number to role, for rows a person is setting now. Nothing for a plain read.
' fresh: read the table as if nobody had set a style on it (Suggest again).
Private Function ScanTable(ByVal d As Object, ByVal t As Object, ByVal types As Collection, Optional ByVal forced As Object = Nothing, _
                           Optional ByVal fresh As Boolean = False) As Object
    Dim res As Object, n As Long, c As Object, rg As Object, i As Long, r As Long, maxRow As Long, maxCols As Long
    Dim ri() As Long, cs() As Long, ce() As Long, kind() As Long, txt() As String, tx As String, descr As String
    Dim rFirst() As Long, rLast() As Long, rRole() As String, rCells() As Long, rFilled() As Long, rText() As String
    Dim roles As Object, ty As Object, capRows As Long, hdrN As Long, f As Object, look As String, capStyle As String
    Dim divHow As String, totWords As String, numExtra As String, maxWords As Long
    Dim nCap As Long, nHdr As Long, nDiv As Long, nTot As Long, nBody As Long, nNum As Long, nShort As Long, nLong As Long, ft As Long, nNotes As Long
    Dim byStyle As Object, assigned As Boolean, pre() As String, sn As String, nNoRole As Long, canMark As Boolean
    Set res = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    descr = LCase$(t.Descr & ";" & t.Title)
    On Error GoTo 0
    Set ty = TypeOfTable(t, types, descr)
    Set res("type") = ty
    res("skip") = False
    If descr Like "*skip*" Then res("skip") = True: res("look") = "Skipped: its Alt Text says skip": Set ScanTable = res: Exit Function
    Set roles = SKW_TableRoles(ty("type"))
    Set res("roles") = roles
    n = 0
    On Error Resume Next
    n = t.Range.Cells.Count
    On Error GoTo 0
    If n = 0 Then res("skip") = True: res("look") = "Skipped: its cells could not be read": Set ScanTable = res: Exit Function
    ReDim ri(1 To n): ReDim cs(1 To n): ReDim ce(1 To n): ReDim kind(1 To n): ReDim txt(1 To n)
    For Each c In t.Range.Cells
        i = i + 1
        If i > n Then Exit For
        If i Mod 200 = 0 Then SKW_Pulse
        ri(i) = c.RowIndex
        Set rg = c.Range
        cs(i) = rg.Start: ce(i) = rg.End
        tx = rg.Text
        If Len(tx) >= 2 Then tx = Left$(tx, Len(tx) - 2) Else tx = ""
        txt(i) = tx
        If ri(i) > maxRow Then maxRow = ri(i)
    Next
    n = i
    ReDim rFirst(1 To maxRow): ReDim rLast(1 To maxRow): ReDim rRole(1 To maxRow)
    ReDim rCells(1 To maxRow): ReDim rFilled(1 To maxRow): ReDim rText(1 To maxRow)
    For i = 1 To n
        r = ri(i)
        If rFirst(r) = 0 Then rFirst(r) = i
        rLast(r) = i
        rCells(r) = rCells(r) + 1
        If Len(Trim$(Replace(txt(i), vbCr, ""))) > 0 Then
            rFilled(r) = rFilled(r) + 1
            If Len(rText(r)) = 0 And i = rFirst(r) Then rText(r) = txt(i)
        End If
    Next
    For r = 1 To maxRow
        If rCells(r) > maxCols Then maxCols = rCells(r)
    Next
    ' Caption row: row 1 holds a caption number field, or is in the caption style.
    capStyle = SKW_Setting("CaptionStyle", "Caption")
    If rFirst(1) > 0 Then
        Set rg = d.Range(cs(rFirst(1)), ce(rLast(1)))
        On Error Resume Next
        For Each f In rg.Fields
            ft = 0
            ft = f.Type
            If ft = wdFieldSequence Then capRows = 1
        Next
        If capRows = 0 And rCells(1) = 1 And maxCols > 1 Then
            If LCase$(SKW_ParaStyleName(rg.Paragraphs(1).Range)) = LCase$(capStyle) Then
                capRows = 1
                look = look & "; row 1 is in the caption style but has no caption number field"
            End If
        End If
        On Error GoTo 0
    End If
    hdrN = 1
    If IsNumeric(roles("header")("recognize")) Then hdrN = CLng(roles("header")("recognize"))
    If hdrN < 0 Then hdrN = 0
    divHow = LCase$(roles("divider")("recognize"))
    totWords = roles("total")("recognize")
    numExtra = roles("number")("recognize")
    maxWords = 3
    If IsNumeric(roles("short text")("recognize")) Then maxWords = CLng(roles("short text")("recognize"))
    ' Roles already on the rows. In an assigned table the styles are the roles. In any other table a
    ' style someone set (header, subhead, total...) still wins over a guess; a body style does not.
    Set byStyle = RolesByStyle(roles)
    canMark = (Len(roles("header")("paragraph style")) > 0 And Len(Present(roles, "long text")("paragraph style")) > 0)
    assigned = False
    If canMark Then assigned = TableAssigned(d, t)
    ReDim pre(1 To maxRow)
    For r = 1 To maxRow
        If rFirst(r) > 0 Then
            sn = ""
            If byStyle.Count > 0 Then
                On Error Resume Next
                sn = LCase$(SKW_ParaStyleName(d.Range(cs(rFirst(r)), cs(rFirst(r)))))
                On Error GoTo 0
            End If
            If Not forced Is Nothing Then
                If forced.Exists(r) Then pre(r) = forced(r)
            End If
            If Len(pre(r)) = 0 Then
                If assigned Then
                    If byStyle.Exists(sn) Then
                        pre(r) = byStyle(sn)
                    Else
                        pre(r) = "body"
                        nNoRole = nNoRole + 1
                        look = look & "; row " & r & " has no role (its style is " & IIf(Len(sn) > 0, sn, "not read") & "), read as body"
                    End If
                ElseIf byStyle.Exists(sn) And Not fresh Then
                    If byStyle(sn) <> "body" Then pre(r) = byStyle(sn)
                End If
            End If
        End If
    Next
    If assigned Then
        capRows = 0
        For r = 1 To maxRow
            If pre(r) <> "caption" Then Exit For
            capRows = r
        Next
    End If
    For r = 1 To maxRow
        If rFirst(r) = 0 Then
            rRole(r) = ""
        ElseIf Len(pre(r)) > 0 Then
            rRole(r) = pre(r)
            Select Case pre(r)
                Case "caption": nCap = nCap + 1
                Case "header", "header2": nHdr = nHdr + 1
                Case "divider": nDiv = nDiv + 1
                Case "total", "subtotal": nTot = nTot + 1
                Case "notes": nNotes = nNotes + 1
                Case Else: rRole(r) = "body": nBody = nBody + 1
            End Select
        ElseIf r <= capRows Then
            rRole(r) = "caption": nCap = nCap + 1
        ElseIf r <= capRows + hdrN Then
            If r > capRows + 1 And roles("header2")("absent") <> "yes" Then rRole(r) = "header2" Else rRole(r) = "header"
            nHdr = nHdr + 1
        ElseIf rCells(r) = 1 And maxCols > 1 And StartsWithWord(rText(r), roles("notes")("recognize")) Then
            rRole(r) = "notes": nNotes = nNotes + 1
        ElseIf rCells(r) = 1 And maxCols > 1 And InStr(divHow, "merged") > 0 Then
            rRole(r) = "divider": nDiv = nDiv + 1
        ElseIf rCells(r) > 1 And rFilled(r) = 1 And Len(rText(r)) > 0 And InStr(divHow, "first cell") > 0 And Not StartsWithWord(rText(r), totWords) Then
            rRole(r) = "divider": nDiv = nDiv + 1
            look = look & "; row " & r & " has text in its first cell only, read as a divider"
        ElseIf StartsWithWord(rText(r), roles("subtotal")("recognize")) Then
            rRole(r) = "subtotal": nTot = nTot + 1
        ElseIf StartsWithWord(rText(r), totWords) Then
            rRole(r) = "total": nTot = nTot + 1
        Else
            rRole(r) = "body": nBody = nBody + 1
        End If
    Next
    If maxRow <= capRows Then look = look & "; no rows under the caption row"
    If nHdr = 0 And hdrN > 0 And Not assigned Then look = look & "; no header row found"
    For i = 1 To n
        If rRole(ri(i)) = "body" Or rRole(ri(i)) = "total" Or rRole(ri(i)) = "subtotal" Then
            tx = Trim$(txt(i))
            If Len(Replace(tx, vbCr, "")) = 0 Then
                kind(i) = 0
            ElseIf IsNumberText(tx, numExtra) Then
                kind(i) = 1: nNum = nNum + 1
            ElseIf InStr(tx, vbCr) = 0 And InStr(tx, Chr$(11)) = 0 And WordCount(tx) <= maxWords Then
                kind(i) = 2: nShort = nShort + 1
            Else
                kind(i) = 3: nLong = nLong + 1
            End If
        End If
    Next
    On Error Resume Next
    If t.Tables.Count > 0 Then look = look & "; has " & t.Tables.Count & " table(s) inside it, left as they are"
    On Error GoTo 0
    If Left$(look, 2) = "; " Then look = Mid$(look, 3)
    res("n") = n: res("maxRow") = maxRow: res("capRows") = capRows: res("hdrN") = hdrN
    res("ri") = ri: res("cs") = cs: res("ce") = ce: res("kind") = kind
    res("rFirst") = rFirst: res("rLast") = rLast: res("rRole") = rRole
    res("rowsText") = IIf(nCap > 0, "Caption " & nCap & ", ", "") & "Header " & nHdr & ", Divider " & nDiv & ", Total " & nTot & ", Body " & nBody & IIf(nNotes > 0, ", Notes " & nNotes, "")
    res("cellsText") = "Number " & nNum & ", Short text " & nShort & ", Long text " & nLong
    res("look") = look
    res("assigned") = assigned: res("canMark") = canMark: res("noRole") = nNoRole
    Set ScanTable = res
End Function

Private Function AlignValue(ByVal s As String) As Long
    Select Case LCase$(Trim$(s))
        Case "left": AlignValue = 0
        Case "center", "centre", "centered": AlignValue = 1
        Case "right": AlignValue = 2
        Case "justify", "justified": AlignValue = 3
        Case Else: AlignValue = -1
    End Select
End Function

Private Function VerticalValue(ByVal s As String) As Long
    Select Case LCase$(Trim$(s))
        Case "top": VerticalValue = 0
        Case "center", "centre", "middle", "centered": VerticalValue = 1
        Case "bottom": VerticalValue = 3
        Case Else: VerticalValue = -1
    End Select
End Function

Private Function BorderStyleValue(ByVal s As String) As Long
    Select Case LCase$(Trim$(s))
        Case "none": BorderStyleValue = 0
        Case "single", "solid": BorderStyleValue = 1
        Case "dotted", "dot": BorderStyleValue = 2
        Case "dashed", "dash": BorderStyleValue = 4
        Case "double": BorderStyleValue = 7
        Case Else: BorderStyleValue = -1
    End Select
End Function

' Word's line widths: a quarter, half, three quarters, 1, 1.5, 2.25, 3, 4.5 and 6 points.
Private Function BorderWidthValue(ByVal s As String) As Long
    Dim v As Double
    s = Trim$(Replace(LCase$(s), "pt", ""))
    If Not IsNumeric(s) Or Len(s) = 0 Then Exit Function
    v = CDbl(s)
    Select Case v
        Case 0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6: BorderWidthValue = CLng(v * 8)
    End Select
End Function

' A shading value: RRGGBB (with or without #), or None. -2 means leave it.
Private Function ShadeValue(ByVal s As String) As Long
    s = Replace(Trim$(s), "#", "")
    ShadeValue = -2
    If Len(s) = 0 Then Exit Function
    If LCase$(s) = "none" Then ShadeValue = -16777216: Exit Function
    If Len(s) = 6 Then
        On Error Resume Next
        ShadeValue = RGB(CLng("&H" & Mid$(s, 1, 2)), CLng("&H" & Mid$(s, 3, 2)), CLng("&H" & Mid$(s, 5, 2)))
        If Err.Number <> 0 Then ShadeValue = -2
    End If
End Function

' Applies one role's format to a range of whole cells. what: "all", or "place" for alignment only.
' Sets a paragraph style on whole cells, leaving alone paragraphs already in it and paragraphs in a kept style
' (bullets and numbered notes inside a cell).
Private Sub StyleRange(ByVal rg As Object, ByVal so As Object, ByVal keep As String)
    Dim cur As String, p As Object, target As String, n As Long
    On Error Resume Next
    target = LCase$(so.NameLocal)
    n = rg.Paragraphs.Count
    If n = 1 Or keep = "|" Then
        If n = 1 Then
            cur = ""
            cur = LCase$(rg.Paragraphs(1).Style.NameLocal)
            If cur = target Then Exit Sub
            If InStr(keep, "|" & cur & "|") > 0 Then Exit Sub
        End If
        rg.Style = so
        Exit Sub
    End If
    ' Several paragraphs and a keep list: one paragraph at a time.
    For Each p In rg.Paragraphs
        cur = ""
        cur = LCase$(p.Style.NameLocal)
        If cur <> target And InStr(keep, "|" & cur & "|") = 0 Then p.Style = so
    Next
End Sub

Private Function KeepList(ByVal roles As Object) As String
    Dim v As Variant, s As String
    s = "|"
    For Each v In Split(roles("keep")("recognize"), ";")
        If Len(Trim$(v)) > 0 Then s = s & LCase$(Trim$(v)) & "|"
    Next
    KeepList = s
End Function

Private Sub ApplyRole(ByVal d As Object, ByVal rg As Object, ByVal st As Object, ByVal what As String, ByVal missing As Object, ByVal keep As String)
    Dim nm As String, v As Long, so As Object
    On Error Resume Next
    If what = "all" Then
        nm = st("paragraph style")
        If Len(nm) > 0 Then
            Set so = SKW_GetStyle(d, nm)
            ' A look from the Stationkeeper standards carries the style's values, so a missing style is made.
            If so Is Nothing And st("app") = "yes" Then Set so = MakeLookStyle(d, nm, st, missing)
            If so Is Nothing Then
                missing(nm) = True
            Else
                StyleRange rg, so, keep
            End If
        End If
        If st("app") = "yes" Then
            ' With no tag style the look's font and spacing go on the text itself.
            If so Is Nothing Then LookFormat rg.Font, rg.ParagraphFormat, st
            v = ShadeValue(st("shading"))
            If v = -2 Then v = -16777216
            rg.Cells.Shading.BackgroundPatternColor = v
        Else
            If LCase$(st("bold")) = "yes" Then rg.Font.Bold = True
            If LCase$(st("bold")) = "no" Then rg.Font.Bold = False
            v = ShadeValue(st("shading"))
            If v <> -2 Then rg.Cells.Shading.BackgroundPatternColor = v
        End If
    End If
    v = AlignValue(st("align"))
    If v >= 0 Then rg.ParagraphFormat.Alignment = v
    v = VerticalValue(st("vertical"))
    If v >= 0 Then rg.Cells.VerticalAlignment = v
    If what = "all" Then IndentAndTabs rg, st
End Sub

' Left indent and left tab stops, in inches. Several tab stops are separated by ;. A blank value leaves them as they are.
Private Sub IndentAndTabs(ByVal rg As Object, ByVal st As Object)
    Dim t As Variant, s As String
    On Error Resume Next
    s = Trim$(st("left indent"))
    If Len(s) > 0 And IsNumeric(s) Then rg.ParagraphFormat.LeftIndent = CSng(s) * 72
    s = Trim$(st("tab stops"))
    If Len(s) > 0 Then
        rg.ParagraphFormat.TabStops.ClearAll
        For Each t In Split(s, ";")
            If IsNumeric(Trim$(t)) Then rg.ParagraphFormat.TabStops.Add CSng(Trim$(t)) * 72, 0, 0
        Next
    End If
End Sub

' Captions outside tables (figure captions): every paragraph in a style named on the Figure caption row
' gets that row's alignment, left indent and tab stops. Returns how many.
Private Function StyleFigureCaptions(ByVal d As Object, ByVal types As Collection) As Long
    Dim roles As Object, st As Object, nm As Variant, so As Object, story As Object, r As Object, v As Long, n As Long
    Set roles = SKW_TableRoles(types(1)("type"))
    Set st = roles("figure caption")
    If Len(Trim$(st("recognize"))) = 0 Then Exit Function
    On Error Resume Next
    For Each nm In Split(st("recognize"), ";")
        Set so = Nothing
        If Len(Trim$(nm)) > 0 Then Set so = SKW_GetStyle(d, Trim$(nm))
        If Not so Is Nothing Then
            For Each story In SKW_Stories(d)
                For Each r In SKW_StyleRanges(story, so)
                    If Not r.Information(wdWithInTable) Then
                        v = AlignValue(st("align"))
                        If v >= 0 Then r.ParagraphFormat.Alignment = v
                        IndentAndTabs r, st
                        n = n + 1
                        SKW_Pulse
                    End If
                Next
            Next
        End If
    Next
    StyleFigureCaptions = n
End Function

' Styles one scanned table. Returns what could not be set.
Private Function StyleTable(ByVal d As Object, ByVal t As Object, ByVal sc As Object, ByVal missing As Object) As String
    Dim ty As Object, roles As Object, nm As String, pct As String, pad As String, skipped As String
    Dim ri As Variant, cs As Variant, ce As Variant, kind As Variant, rFirst As Variant, rLast As Variant, rRole As Variant
    Dim r As Long, i As Long, j As Long, k As Long, role As String, rg As Object, lastHead As Long, kd As Long
    Dim keep As String, mg() As String, bs As Long, bw As Long, bc As Long, b As Variant, hasB As Boolean
    Set ty = sc("type"): Set roles = sc("roles")
    keep = KeepList(roles)
    ri = sc("ri"): cs = sc("cs"): ce = sc("ce"): kind = sc("kind")
    rFirst = sc("rFirst"): rLast = sc("rLast"): rRole = sc("rRole")
    On Error Resume Next
    nm = ty("table style")
    If Len(nm) > 0 Then
        If SKW_GetStyle(d, nm) Is Nothing Then
            missing(nm) = True
        ElseIf LCase$(TableStyleName(t)) <> LCase$(nm) Then
            t.Range.Style = nm
        End If
    End If
    If Len(ty("rows may break across pages")) > 0 Then t.Rows.AllowBreakAcrossPages = SKW_IsYes(ty("rows may break across pages"))
    pct = ty("width percent")
    If IsNumeric(pct) And Len(pct) > 0 Then
        t.PreferredWidthType = wdPreferredWidthPercent
        t.PreferredWidth = CSng(pct)
    End If
    pad = ty("cell padding")
    If IsNumeric(pad) And Len(pad) > 0 Then
        t.LeftPadding = CSng(pad): t.RightPadding = CSng(pad)
        t.TopPadding = CSng(pad): t.BottomPadding = CSng(pad)
    End If
    ' Cell margins in inches: one value for all four sides, or top; bottom; left; right.
    If Len(ty("cell margins")) > 0 Then
        mg = Split(Replace(ty("cell margins"), ",", ";"), ";")
        If UBound(mg) = 0 And IsNumeric(Trim$(mg(0))) Then
            t.TopPadding = CSng(Trim$(mg(0))) * 72: t.BottomPadding = t.TopPadding
            t.LeftPadding = t.TopPadding: t.RightPadding = t.TopPadding
        ElseIf UBound(mg) = 3 Then
            If IsNumeric(Trim$(mg(0))) Then t.TopPadding = CSng(Trim$(mg(0))) * 72
            If IsNumeric(Trim$(mg(1))) Then t.BottomPadding = CSng(Trim$(mg(1))) * 72
            If IsNumeric(Trim$(mg(2))) Then t.LeftPadding = CSng(Trim$(mg(2))) * 72
            If IsNumeric(Trim$(mg(3))) Then t.RightPadding = CSng(Trim$(mg(3))) * 72
        Else
            skipped = skipped & "; cell margins (give one value, or top; bottom; left; right, in inches)"
        End If
    End If
    ' The same border on every edge and between every row and column.
    bs = BorderStyleValue(ty("border style")): bw = BorderWidthValue(ty("border width")): bc = ShadeValue(ty("border color"))
    hasB = (bs >= 0 Or bw > 0 Or bc <> -2)
    If hasB Then
        For Each b In Array(wdBorderTop, wdBorderLeft, wdBorderBottom, wdBorderRight, wdBorderHorizontal, wdBorderVertical)
            Err.Clear
            With t.Borders(b)
                If bs >= 0 Then .LineStyle = bs
                If bs <> 0 Then
                    If bs < 0 And .LineStyle = 0 Then .LineStyle = 1
                    If bw > 0 Then .LineWidth = bw
                    If bc <> -2 Then .Color = bc
                End If
            End With
        Next
    End If
    If ty("app") = "yes" Then
        ' Stationkeeper standards: margins in points, the table's place on the page, and the rows whose look says Repeat.
        mg = Split(ty("margins pt"), ";")
        If UBound(mg) = 3 Then
            If IsNumeric(mg(0)) And Len(mg(0)) > 0 Then t.TopPadding = CSng(mg(0))
            If IsNumeric(mg(1)) And Len(mg(1)) > 0 Then t.BottomPadding = CSng(mg(1))
            If IsNumeric(mg(2)) And Len(mg(2)) > 0 Then t.LeftPadding = CSng(mg(2))
            If IsNumeric(mg(3)) And Len(mg(3)) > 0 Then t.RightPadding = CSng(mg(3))
        End If
        If AlignValue(ty("table align")) >= 0 Then t.Rows.Alignment = AlignValue(ty("table align"))
        lastHead = 0
        For r = 1 To sc("maxRow")
            If Len(rRole(r)) = 0 Then Exit For
            If Present(roles, rRole(r))("repeat") <> "yes" Then Exit For
            lastHead = r
        Next
        If lastHead >= 1 Then
            Err.Clear
            d.Range(cs(rFirst(1)), ce(rLast(lastHead))).Rows.HeadingFormat = True
            If Err.Number <> 0 Then skipped = skipped & "; repeating header rows (merged cells)"
            Err.Clear
        End If
    End If
    ' Word repeats heading rows only from row 1 down, so the caption row repeats with the header row.
    If Len(ty("header row repeats")) > 0 Then
        lastHead = sc("capRows") + sc("hdrN")
        If lastHead > sc("maxRow") Then lastHead = sc("maxRow")
        If lastHead >= 1 And SKW_IsYes(ty("header row repeats")) Then
            Err.Clear
            d.Range(cs(rFirst(1)), ce(rLast(lastHead))).Rows.HeadingFormat = True
            If Err.Number <> 0 Then skipped = skipped & "; repeating header rows (merged cells)"
            Err.Clear
        End If
    End If
    For r = 1 To sc("maxRow")
        role = rRole(r)
        If r Mod 20 = 0 Then SKW_Pulse
        If Len(role) > 0 Then
            Set rg = d.Range(cs(rFirst(r)), ce(rLast(r)))
            Select Case role
                Case "caption", "header", "header2", "divider", "notes"
                    ApplyRole d, rg, Present(roles, role), "all", missing, keep
                    ApplyEdges rg, Present(roles, role)
                Case "total", "subtotal"
                    ApplyRole d, rg, Present(roles, role), "all", missing, keep
                    ApplyEdges rg, Present(roles, role)
                    For i = rFirst(r) To rLast(r)
                        If kind(i) = 1 And roles("number")("absent") <> "yes" Then ApplyRole d, d.Range(cs(i), ce(i)), roles("number"), "place", missing, keep
                    Next
                Case "body"
                    i = rFirst(r)
                    Do While i <= rLast(r)
                        kd = kind(i): If kd = 0 Then kd = 3
                        j = i
                        Do While j < rLast(r)
                            k = kind(j + 1): If k = 0 Then k = 3
                            If k <> kd Then Exit Do
                            j = j + 1
                        Loop
                        ApplyRole d, d.Range(cs(i), ce(j)), Present(roles, Choose(kd, "number", "short text", "long text")), "all", missing, keep
                        i = j + 1
                    Loop
                    ApplyEdges rg, roles("long text")
            End Select
        End If
    Next
    If Left$(skipped, 2) = "; " Then skipped = Mid$(skipped, 3)
    StyleTable = skipped
End Function

Private Function MissingText(ByVal missing As Object) As String
    Dim k As Variant, s As String, made As String
    For Each k In missing.Keys
        If Left$(k, 5) = "made:" Then
            made = made & IIf(Len(made) > 0, ", ", "") & Mid$(k, 6)
        Else
            s = s & IIf(Len(s) > 0, ", ", "") & k
        End If
    Next
    If Len(made) > 0 Then MissingText = "Styles the document lacks, made from the standards: " & made & "."
    If Len(s) > 0 Then MissingText = MissingText & IIf(Len(MissingText) > 0, " ", "") & "Not in this document, so not applied: " & s & ". Run Import Team Styles first."
End Function

' =====================================================================
' The Stationkeeper app's standards (Company Standards.xlsx, format 2).
' When the standards workbook the tools read has a Table looks sheet, the
' table tools take every value from it: one look per row role (caption,
' header, header2, subhead, body, subtotal, total, note) and the table
' types with their cell margins, roles and overrides. The tools' own
' Table roles sheet is not read then. Only Approved rows count. A value
' that starts with = names a row of the Settings sheet.
' Read by the macros only, and kept by the app: a Recognize column on
' Table looks, and looks named number, short and keep.
' =====================================================================
Public Function SKW_AppStandards() As Boolean
    SKW_AppStandards = (SKW_Rows("Table looks").Count > 0)
End Function

Private Function AppApproved(ByVal r As Object) As Boolean
    Dim s As String
    If r.Exists("status") Then s = LCase$(Trim$(r("status")))
    AppApproved = (Len(s) = 0 Or s = "approved")
End Function

' A cell of an app sheet, with =name looked up on the Settings sheet.
Private Function AppVal(ByVal r As Object, ByVal key As String) As String
    Dim v As String
    If r.Exists(key) Then v = Trim$(r(key))
    If Left$(v, 1) = "=" Then v = SKW_Setting(Mid$(v, 2), "")
    AppVal = v
End Function

Private Function YesNo(ByVal v As String) As String
    Select Case LCase$(Trim$(v))
        Case "true", "yes", "1": YesNo = "yes"
        Case "false", "no", "0": YesNo = "no"
    End Select
End Function

' The role to use: the one asked for, or the body look when the standards have no such look.
Private Function Present(ByVal roles As Object, ByVal role As String) As Object
    If roles.Exists(role) Then
        If roles(role)("absent") <> "yes" Then Set Present = roles(role): Exit Function
    End If
    Set Present = roles("long text")
End Function

' One look as a role. Keys the older sheet used are filled too, so the same code styles from either.
Private Sub AppLook(ByVal res As Object, ByVal role As String, ByVal r As Object, ByVal recog As String)
    Dim d As Object, e As Variant
    Set d = CreateObject("Scripting.Dictionary")
    d("app") = "yes": d("absent") = ""
    d("paragraph style") = AppVal(r, "tag style")
    d("align") = AppVal(r, "align")
    d("vertical") = AppVal(r, "valign")
    d("shading") = AppVal(r, "fill")
    d("bold") = YesNo(AppVal(r, "bold"))
    d("italic") = YesNo(AppVal(r, "italic"))
    d("font") = AppVal(r, "font")
    d("size") = AppVal(r, "size (pt)")
    d("color") = AppVal(r, "color (hex)")
    d("before") = AppVal(r, "space before (pt)")
    d("after") = AppVal(r, "space after (pt)")
    d("left pt") = AppVal(r, "left indent (pt)")
    d("right pt") = AppVal(r, "right indent (pt)")
    d("repeat") = YesNo(AppVal(r, "repeat"))
    For Each e In Array("top", "bottom", "left", "right", "inside v")
        d("edge " & e) = AppVal(r, "edges: " & e)
    Next
    d("left indent") = "": d("tab stops") = ""
    If r.Exists("tab stops") Then d("tab stops") = Trim$(r("tab stops"))
    d("recognize") = recog
    If r.Exists("recognize") Then
        If Len(Trim$(r("recognize"))) > 0 Then d("recognize") = Trim$(r("recognize"))
    End If
    Set res(role) = d
End Sub

Private Sub AppAbsent(ByVal res As Object, ByVal role As String, ByVal recog As String)
    Dim d As Object, k As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each k In Array("paragraph style", "align", "vertical", "shading", "bold", "italic", "font", "size", "color", "before", "after", _
                        "left pt", "right pt", "repeat", "edge top", "edge bottom", "edge left", "edge right", "edge inside v", "left indent", "tab stops")
        d(k) = ""
    Next
    d("app") = "yes": d("absent") = "yes": d("recognize") = recog
    Set res(role) = d
End Sub

' The roles for one table type from the app's Table looks, with the type's overrides laid over them.
Private Function AppRoles(ByVal tyName As String) As Object
    Dim res As Object, r As Object, nm As String, ty As Object, k As Object, ov As Object, used As Object, v As Variant
    Dim role As Variant, key As Variant, o As Object, ed As Variant, names As Object
    Set res = CreateObject("Scripting.Dictionary")
    res.CompareMode = 1
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = 1
    names("caption") = "caption": names("header") = "header": names("header2") = "header2": names("subhead") = "divider"
    names("body") = "long text": names("subtotal") = "subtotal": names("total") = "total": names("note") = "notes"
    names("number") = "number": names("short") = "short text": names("keep") = "keep": names("figure caption") = "figure caption"
    AppAbsent res, "caption", ""
    AppAbsent res, "header", "1"
    AppAbsent res, "header2", ""
    AppAbsent res, "divider", "merged"
    AppAbsent res, "subtotal", "Subtotal"
    AppAbsent res, "total", "Total;Grand Total"
    AppAbsent res, "notes", "Note;Notes;Source;Sources"
    AppAbsent res, "number", "-"
    AppAbsent res, "short text", "3"
    AppAbsent res, "long text", ""
    AppAbsent res, "keep", SKW_Setting("table_keep_styles", "")
    AppAbsent res, "figure caption", ""
    For Each k In SKW_TableTypes()
        If LCase$(k("type")) = LCase$(tyName) Then Set ty = k
    Next
    ' Which looks this type uses. No list means all of them.
    If Not ty Is Nothing Then
        Set o = SKW_Json(ty("roles"))
        If Not o Is Nothing Then
            If TypeName(o) = "Collection" Then
                Set used = CreateObject("Scripting.Dictionary")
                used.CompareMode = 1
                For Each v In o
                    used(CStr(v)) = True
                Next
            End If
        End If
    End If
    For Each r In SKW_Rows("Table looks")
        nm = ""
        If r.Exists("name") Then nm = LCase$(Trim$(r("name")))
        If AppApproved(r) And names.Exists(nm) Then
            ' number, short and keep are the macros' own looks: a type's list does not have to name them.
            If used Is Nothing Then
                AppLook res, names(nm), r, res(names(nm))("recognize")
            ElseIf used.Exists(nm) Or nm = "number" Or nm = "short" Or nm = "keep" Or nm = "figure caption" Then
                AppLook res, names(nm), r, res(names(nm))("recognize")
            End If
        End If
    Next
    ' The type's overrides: {"header": {"fill": "D9E2F3", "edges": {"top": null}}}
    If Not ty Is Nothing Then
        Set ov = SKW_Json(ty("override"))
        If Not ov Is Nothing Then
            If TypeName(ov) = "Dictionary" Then
                For Each role In ov.Keys
                    If names.Exists(CStr(role)) And IsObject(ov(role)) Then
                        If res(names(CStr(role)))("absent") <> "yes" Then
                            Set o = ov(role)
                            For Each key In o.Keys
                                Select Case LCase$(CStr(key))
                                    Case "tag_style": res(names(CStr(role)))("paragraph style") = JText(o(key))
                                    Case "fill": res(names(CStr(role)))("shading") = JText(o(key))
                                    Case "valign": res(names(CStr(role)))("vertical") = JText(o(key))
                                    Case "align", "font", "size", "color", "before", "after": res(names(CStr(role)))(LCase$(CStr(key))) = JText(o(key))
                                    Case "bold", "italic", "repeat": res(names(CStr(role)))(LCase$(CStr(key))) = YesNo(JText(o(key)))
                                    Case "left": res(names(CStr(role)))("left pt") = JText(o(key))
                                    Case "right": res(names(CStr(role)))("right pt") = JText(o(key))
                                    Case "edges"
                                        If IsObject(o(key)) Then
                                            For Each ed In o(key).Keys
                                                res(names(CStr(role)))("edge " & Replace(LCase$(CStr(ed)), "_", " ")) = EdgeText(o(key)(ed))
                                            Next
                                        End If
                                End Select
                            Next
                        End If
                    End If
                Next
            End If
        End If
    End If
    Set AppRoles = res
End Function

Private Function JText(ByVal v As Variant) As String
    If IsObject(v) Then Exit Function
    If IsNull(v) Then Exit Function
    JText = CStr(v)
End Function

' An edge from an override: null means no line, a list is [width, "color"].
Private Function EdgeText(ByVal v As Variant) As String
    If IsObject(v) Then
        If TypeName(v) = "Collection" Then
            If v.Count >= 2 Then EdgeText = "[" & CStr(v(1)) & ", """ & CStr(v(2)) & """]"
        End If
    ElseIf IsNull(v) Then
        EdgeText = ""
    Else
        EdgeText = CStr(v)
    End If
End Function

' The font and spacing of a look, on a style or on text.
Private Sub LookFormat(ByVal fnt As Object, ByVal pf As Object, ByVal st As Object)
    Dim v As Long
    On Error Resume Next
    If Len(st("font")) > 0 Then fnt.Name = st("font")
    If IsNumeric(st("size")) And Len(st("size")) > 0 Then fnt.Size = CSng(st("size"))
    If st("bold") = "yes" Then fnt.Bold = True
    If st("bold") = "no" Then fnt.Bold = False
    If st("italic") = "yes" Then fnt.Italic = True
    If st("italic") = "no" Then fnt.Italic = False
    v = ShadeValue(st("color"))
    If v <> -2 Then fnt.Color = v
    If IsNumeric(st("before")) And Len(st("before")) > 0 Then pf.SpaceBefore = CSng(st("before"))
    If IsNumeric(st("after")) And Len(st("after")) > 0 Then pf.SpaceAfter = CSng(st("after"))
    If IsNumeric(st("left pt")) And Len(st("left pt")) > 0 Then pf.LeftIndent = CSng(st("left pt"))
    If IsNumeric(st("right pt")) And Len(st("right pt")) > 0 Then pf.RightIndent = CSng(st("right pt"))
    v = AlignValue(st("align"))
    If v >= 0 Then pf.Alignment = v
End Sub

' Makes a look's tag style in the document from the look's values. Returns it, or Nothing.
Private Function MakeLookStyle(ByVal d As Object, ByVal nm As String, ByVal st As Object, ByVal missing As Object) As Object
    Dim so As Object
    On Error Resume Next
    Set so = d.Styles.Add(nm, wdStyleTypeParagraph)
    If so Is Nothing Then Exit Function
    so.BaseStyle = d.Styles(wdStyleNormal)
    so.NextParagraphStyle = so
    so.QuickStyle = False
    LookFormat so.Font, so.ParagraphFormat, st
    missing("made:" & nm) = True
    Set MakeLookStyle = so
End Function

' One edge: "[1.0, ""1F3864""]" is a single line of that width and color; "" is no line.
Private Sub OneEdge(ByVal b As Object, ByVal spec As String)
    Dim o As Object, w As Long, c As Long
    On Error Resume Next
    If Len(Trim$(spec)) = 0 Then b.LineStyle = 0: Exit Sub
    Set o = SKW_Json(spec)
    If o Is Nothing Then Exit Sub
    If TypeName(o) <> "Collection" Then Exit Sub
    If o.Count < 2 Then Exit Sub
    w = BorderWidthValue(CStr(o(1)))
    c = ShadeValue(CStr(o(2)))
    b.LineStyle = 1
    If w > 0 Then b.LineWidth = w
    If c <> -2 Then b.Color = c
End Sub

' The edges of a row from its look. Only looks from the Stationkeeper standards carry edges.
Private Sub ApplyEdges(ByVal rg As Object, ByVal st As Object)
    If st("app") <> "yes" Then Exit Sub
    On Error Resume Next
    With rg.Cells.Borders
        OneEdge .Item(wdBorderTop), st("edge top")
        OneEdge .Item(wdBorderBottom), st("edge bottom")
        OneEdge .Item(wdBorderLeft), st("edge left")
        OneEdge .Item(wdBorderRight), st("edge right")
        If rg.Cells.Count > 1 Then OneEdge .Item(wdBorderVertical), st("edge inside v")
    End With
End Sub

Private Function RolesSource() As String
    If SKW_AppStandards() Then
        RolesSource = "The looks come from the Table looks sheet of the Stationkeeper standards" & IIf(Len(SKW_StdChanged()) > 0, " (" & SKW_StdChanged() & ").", ".") & _
            " A tag style the document lacks is made from its look."
    ElseIf SKW_Rows("Table roles").Count > 0 Then
        RolesSource = "The roles come from the Table roles sheet" & IIf(SKW_FromEffort("Table roles"), " of the linked effort.", " of the standards workbook.")
    Else
        RolesSource = "There is no Table roles sheet, so the built-in roles are used."
    End If
End Function

Public Function SKW_TableRolesCheck() As String
    Dim d As Object, msg As String, types As Collection, rows As New Collection, t As Object, i As Long, sc As Object
    Dim nLook As Long, nSkip As Long, nOk As Long, missing As Object, roles As Object, k As Variant, nm As String, role As Variant
    Dim nAssigned As Long
    Set SKW_LastReport = Nothing
    SKW_Fixable = 0
    msg = SKW_Begin("Table roles check", False)
    If Len(msg) > 0 Then SKW_TableRolesCheck = msg: Exit Function
    Set d = ActiveDocument
    Set types = SKW_TableTypes()
    Set missing = CreateObject("Scripting.Dictionary")
    For Each t In d.Tables
        i = i + 1
        SKW_Status "Reading table " & i & " of " & d.Tables.Count & "..."
        Set sc = ScanTable(d, t, types)
        If sc("skip") Then
            nSkip = nSkip + 1
            rows.Add SKW_Row(i, SKW_Page(t.Range), sc("type")("type"), "", "", sc("look"), FirstCellText(t))
        Else
            nOk = nOk + 1
            If Len(sc("look")) > 0 Then nLook = nLook + 1
            rows.Add SKW_Row(i, SKW_Page(t.Range), sc("type")("type") & IIf(sc("assigned"), ", assigned", ""), sc("rowsText"), sc("cellsText"), sc("look"), FirstCellText(t))
            If sc("assigned") Then nAssigned = nAssigned + 1
            Set roles = sc("roles")
            For Each role In roles.Keys
                nm = roles(role)("paragraph style")
                If Len(nm) > 0 Then
                    If SKW_GetStyle(d, nm) Is Nothing Then
                        If roles(role)("app") = "yes" Then missing("made:" & nm) = True Else missing(nm) = True
                    End If
                End If
            Next
            nm = sc("type")("table style")
            If Len(nm) > 0 Then
                If SKW_GetStyle(d, nm) Is Nothing Then missing(nm) = True
            End If
        End If
    Next
    SKW_Finish
    SKW_Report "Table Roles Check", d.Tables.Count & " table(s) read, " & nLook & " with rows to look at, " & nSkip & " skipped. " & _
        RolesSource() & " " & nAssigned & " table(s) are assigned: their rows are read from their styles, not guessed. " & MissingText(missing) & " To leave a table alone, put skip in its Alt Text (Table Properties); " & _
        "to pick its table type, put type=Name there. Changes nothing.", _
        Array("Table", "Page", "Type", "Rows", "Cells", "Look at", "First cell"), rows
    SKW_Fixable = nOk
    SKW_TableRolesCheck = d.Tables.Count & " table(s) read, " & nLook & " with rows to look at, " & nSkip & " skipped."
End Function

' Styles every table of a document. Returns "styled|skipped|notes".
Private Function StyleAllTables(ByVal d As Object, ByVal clearUndo As Boolean, ByVal missing As Object) As String
    Dim types As Collection, t As Object, i As Long, sc As Object, nDone As Long, nSkip As Long, notes As String, s As String, total As Long
    Set types = SKW_TableTypes()
    total = d.Tables.Count
    For Each t In d.Tables
        i = i + 1
        SKW_Status "Styling table " & i & " of " & total & "..."
        Set sc = ScanTable(d, t, types)
        If sc("skip") Then
            nSkip = nSkip + 1
        Else
            s = StyleTable(d, t, sc, missing)
            If sc("canMark") Then MarkAssigned d, t
            nDone = nDone + 1
            If Len(s) > 0 Then notes = notes & IIf(Len(notes) > 0, "; ", "") & "table " & i & ": " & s
        End If
        If clearUndo Then
            On Error Resume Next
            d.UndoClear
            On Error GoTo 0
        End If
    Next
    SKW_Status "Setting captions outside tables..."
    StyleAllTables = nDone & "|" & nSkip & "|" & notes & "|" & StyleFigureCaptions(d, types)
End Function

' Styles every table in a copy of the active document. target: the copy's full path, or "" for one beside the document.
Public Function SKW_TableRolesFixCopy(ByVal target As String) As String
    Dim d As Object, c As Object, fso As Object, n As Long, parts() As String, missing As Object, ed As String, base As String, ext As String
    Set SKW_LastReport = Nothing
    If Documents.Count = 0 Then SKW_TableRolesFixCopy = "Open a document first.": Exit Function
    Set d = ActiveDocument
    If Len(d.Path) = 0 Then SKW_TableRolesFixCopy = "Save the document first. The copy is made from the saved file.": Exit Function
    If Not d.Saved Then SKW_TableRolesFixCopy = "Save the document first. The copy is made from the saved file, so unsaved changes would be left out.": Exit Function
    If Len(target) = 0 Then
        base = d.Path & "\" & SKW_Stem(d.Name) & " (tables styled)"
        ext = Mid$(d.Name, InStrRev(d.Name, "."))
        target = base & ext
        Do While Len(Dir$(target)) > 0
            n = n + 1
            target = base & " " & (n + 1) & ext
        Loop
    End If
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile d.FullName, target, True
    SKW_ProgressTitle "Table roles: styling a copy"
    SKW_Status "Opening the copy..."
    Set c = Documents.Open(FileName:=target, AddToRecentFiles:=False, Visible:=True)
    Application.ScreenUpdating = False
    Set missing = CreateObject("Scripting.Dictionary")
    c.TrackRevisions = False
    parts = Split(StyleAllTables(c, True, missing), "|")
    SKW_Status "Saving the copy..."
    c.Save
    Application.ScreenUpdating = True
    Application.StatusBar = False
    SKW_ProgressDone
    SKW_TableRolesFixCopy = parts(0) & " table(s) styled in a copy, " & parts(1) & " skipped: " & Mid$(target, InStrRev(target, "\") + 1) & "." & _
        IIf(Val(parts(3)) > 0, " " & parts(3) & " caption(s) outside tables set.", "") & _
        IIf(Len(parts(2)) > 0, vbCr & "Not set: " & parts(2) & ".", "") & _
        IIf(Len(MissingText(missing)) > 0, vbCr & MissingText(missing), "") & vbCr & "Your document is not changed. The copy is open."
    Exit Function
Fail:
    ed = Err.Description
    Application.ScreenUpdating = True
    Application.StatusBar = False
    SKW_ProgressDone
    SKW_TableRolesFixCopy = "Stopped: " & ed & ". Your document is not changed."
End Function

Public Function SKW_ApplyTableRolesHere() As String
    Dim d As Object, msg As String, t As Object, sc As Object, missing As Object, s As String
    Set SKW_LastReport = Nothing
    If Documents.Count > 0 Then
        If Not Selection.Information(wdWithInTable) Then SKW_ApplyTableRolesHere = "Put the cursor in a table first.": Exit Function
    End If
    msg = SKW_Begin("Apply table roles", True)
    If Len(msg) > 0 Then SKW_ApplyTableRolesHere = msg: Exit Function
    Set d = ActiveDocument
    Set t = Selection.Tables(1)
    Set missing = CreateObject("Scripting.Dictionary")
    Set sc = ScanTable(d, t, SKW_TableTypes())
    If sc("skip") Then
        SKW_Finish
        SKW_ApplyTableRolesHere = sc("look") & ". Nothing was changed."
        Exit Function
    End If
    s = StyleTable(d, t, sc, missing)
    If sc("canMark") Then MarkAssigned d, t
    SKW_Finish
    SKW_ApplyTableRolesHere = "Table styled as " & sc("type")("type") & IIf(sc("assigned"), " (its rows were read from their styles)", IIf(sc("canMark"), " (roles suggested; the table is now assigned)", "")) & ". Rows: " & sc("rowsText") & ". Cells: " & sc("cellsText") & "." & _
        IIf(Len(sc("look")) > 0, vbCr & "Look at: " & sc("look") & ".", "") & _
        IIf(Len(s) > 0, vbCr & "Not set: " & s & ".", "") & _
        IIf(Len(MissingText(missing)) > 0, vbCr & MissingText(missing), "") & vbCr & "One Ctrl+Z undoes it."
End Function

' =====================================================================
' Assigning by hand, and seeing the roles (1.9.0)
' =====================================================================

' Sets the rows in the selection to a role you pick (header, subhead, body, subtotal, total, note, caption) and styles them. The table is then assigned: its rows keep their roles on every later fix. One Ctrl+Z undoes it.
Public Sub AssignTableRole()
    Dim pick As String, keys As Variant, labels As Variant, s As String, i As Long
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    If Not Selection.Information(wdWithInTable) Then SKW_Say "Put the cursor in a table row first, or select the rows.": Exit Sub
    If SKW_IsQuiet() Then Exit Sub
    keys = Array("header", "header2", "divider", "body", "subtotal", "total", "notes", "caption", "suggest")
    labels = Array("Header", "Header 2 (a second header row)", "Subhead (a row that divides the table)", "Body", "Subtotal", "Total", "Note", "Caption", _
                   "Suggest again (forget this table's roles and read it afresh)")
    For i = 0 To UBound(labels)
        s = s & (i + 1) & "  " & labels(i) & vbCr
    Next
    pick = InputBox("Which role for the selected row(s)?" & vbCr & vbCr & s & vbCr & "Type the number.", "Assign Table Role", "4")
    If Len(pick) = 0 Then Exit Sub
    If Not IsNumeric(pick) Then SKW_Say "Type a number from the list.": Exit Sub
    If CLng(pick) < 1 Or CLng(pick) > UBound(keys) + 1 Then SKW_Say "Type a number from the list.": Exit Sub
    SKW_Show SKW_AssignTableRole(CStr(keys(CLng(pick) - 1)))
End Sub

' role: header, header2, divider, body, subtotal, total, notes, caption, or suggest.
Public Function SKW_AssignTableRole(ByVal role As String) As String
    Dim d As Object, msg As String, t As Object, sc As Object, missing As Object, s As String, forced As Object, c As Object, n As Long
    Set SKW_LastReport = Nothing
    If Documents.Count > 0 Then
        If Not Selection.Information(wdWithInTable) Then SKW_AssignTableRole = "Put the cursor in a table row first, or select the rows.": Exit Function
    End If
    role = LCase$(Trim$(role))
    msg = SKW_Begin("Assign table role", True)
    If Len(msg) > 0 Then SKW_AssignTableRole = msg: Exit Function
    Set d = ActiveDocument
    Set t = Selection.Tables(1)
    Set missing = CreateObject("Scripting.Dictionary")
    Set forced = CreateObject("Scripting.Dictionary")
    If role = "suggest" Then
        ClearAssigned d, t
    Else
        On Error Resume Next
        For Each c In Selection.Cells
            forced(c.RowIndex) = role
        Next
        On Error GoTo 0
        n = forced.Count
        If n = 0 Then
            SKW_Finish
            SKW_AssignTableRole = "No row was selected. Nothing was changed."
            Exit Function
        End If
    End If
    Set sc = ScanTable(d, t, SKW_TableTypes(), forced, (role = "suggest"))
    If sc("skip") Then
        SKW_Finish
        SKW_AssignTableRole = sc("look") & ". Nothing was changed."
        Exit Function
    End If
    If Not sc("canMark") Then
        SKW_Finish
        SKW_AssignTableRole = "The standards give the roles no styles, so a role cannot be kept on a row. " & _
            "Point the tools at the Stationkeeper standards workbook, or give the Header and Long text roles a paragraph style. Nothing was changed."
        Exit Function
    End If
    s = StyleTable(d, t, sc, missing)
    MarkAssigned d, t
    SKW_Finish
    If role = "suggest" Then
        SKW_AssignTableRole = "The table was read afresh and is assigned again. Rows: " & sc("rowsText") & "."
    Else
        SKW_AssignTableRole = n & " row(s) set to " & IIf(role = "divider", "subhead", IIf(role = "notes", "note", role)) & ". Rows: " & sc("rowsText") & "." & _
            IIf(sc("assigned"), "", vbCr & "The other rows were suggested. The table is now assigned: its rows keep their roles.")
    End If
    SKW_AssignTableRole = SKW_AssignTableRole & IIf(Len(sc("look")) > 0, vbCr & "Look at: " & sc("look") & ".", "") & _
        IIf(Len(s) > 0, vbCr & "Not set: " & s & ".", "") & IIf(Len(MissingText(missing)) > 0, vbCr & MissingText(missing), "") & vbCr & "One Ctrl+Z undoes it."
End Function

' Shows each row's role beside it: Word's Draft view with the style area open, so every paragraph's style name sits at the left. Run it again to go back.
Public Sub ShowTableRoles()
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    SKW_Say SKW_ShowTableRoles()
End Sub

Public Function SKW_ShowTableRoles() As String
    Static wasView As Long, wasWidth As Single
    Dim w As Object
    On Error GoTo Fail
    Set w = ActiveWindow
    ' The style area shows only in Draft and Outline view, so "on" means both: that view and a width.
    If w.StyleAreaWidth > 0 And w.View.Type = wdNormalView Then
        w.StyleAreaWidth = wasWidth
        If wasView > 0 Then w.View.Type = wasView Else w.View.Type = wdPrintView
        SKW_ShowTableRoles = "Roles are hidden again."
    Else
        wasView = w.View.Type
        wasWidth = w.StyleAreaWidth
        w.View.Type = wdNormalView
        If w.StyleAreaWidth < 100 Then w.StyleAreaWidth = 110
        SKW_ShowTableRoles = "Each row's style shows at the left: that is its role. Select rows and run Assign Table Role to change one. Run Show Table Roles again to go back."
    End If
    Exit Function
Fail:
    SKW_ShowTableRoles = "The view could not be changed: " & Err.Description
End Function
