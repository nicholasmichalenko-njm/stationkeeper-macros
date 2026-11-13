Attribute VB_Name = "SKW_Setup"
' =====================================================================
' Stationkeeper Tools: setup.
'   MakeStandardsWorkbook     makes a team standards workbook in Excel to fill in
'   ChooseStandardsWorkbook   points the tools at the team standards workbook
' SKW_ToolList gives the Stationkeeper Macros menu the title, category and
' description of every tool, so they show up described in the menu.
' =====================================================================
Option Explicit

' Makes a new standards workbook (allowed styles, table types, placeholders, settings and more) with sensible starting values, opens it in Excel to fill in, and points the tools at it.
Public Sub MakeStandardsWorkbook()
    Dim p As String, sug As String
    If SKW_IsQuiet() Then Exit Sub   ' needs a person to choose where
    sug = ThisDocument.Path & "\" & SKW_Stem(ThisDocument.Name) & " standards.xlsx"
    With Application.FileDialog(2)
        .Title = "Where should the standards workbook go?"
        .InitialFileName = sug
        If .Show <> -1 Then Exit Sub
        p = .SelectedItems(1)
    End With
    If LCase$(Right$(p, 5)) <> ".xlsx" Then p = SKW_Stem(p) & ".xlsx"
    If Len(Dir$(p)) > 0 Then
        If MsgBox(Mid$(p, InStrRev(p, "\") + 1) & " already exists. Replace it?", vbYesNo + vbQuestion, SKW_APP) <> vbYes Then Exit Sub
    End If
    SKW_Say SKW_MakeStandardsWorkbook(p, True)
End Sub

' Points the tools at your team's standards workbook (allowed styles, table types, placeholders and settings). Each person chooses it once per template.
Public Sub ChooseStandardsWorkbook()
    Dim p As String
    If SKW_IsQuiet() Then Exit Sub   ' needs a person to choose the file
    With Application.FileDialog(3)
        .Title = "Choose the team standards workbook"
        .Filters.Clear
        .Filters.Add "Excel workbooks", "*.xlsx;*.xlsm;*.xls"
        .AllowMultiSelect = False
        If Len(SKW_StandardsPath()) > 0 Then .InitialFileName = SKW_StandardsPath()
        If .Show <> -1 Then Exit Sub
        p = .SelectedItems(1)
    End With
    SKW_Say SKW_ChooseStandards(p)
End Sub

' Chooses the folder where reports are saved, on this PC. Each report is saved there as it is made, named for the report, the file it checked and the time. With no folder chosen, reports open and are not saved.
Public Sub ChooseReportsFolder()
    Dim cur As String, a As VbMsgBoxResult, p As String
    If SKW_IsQuiet() Then Exit Sub   ' needs a person to choose the folder
    cur = SKW_ReportsFolder()
    If Len(cur) > 0 Then
        a = MsgBox("Reports are saved in:" & vbCr & cur & vbCr & vbCr & "Choose another folder?" & vbCr & _
                   "No stops saving them: reports then open and are yours to save or close.", vbYesNoCancel + vbQuestion, SKW_APP)
        If a = vbCancel Then Exit Sub
        If a = vbNo Then
            SKW_SetReportsFolder ""
            SKW_Say "Reports are no longer saved. They open, and you save the ones you want."
            Exit Sub
        End If
    End If
    With Application.FileDialog(4)
        .Title = "Choose where Stationkeeper reports are saved"
        If Len(cur) > 0 Then .InitialFileName = cur & "\"
        If .Show <> -1 Then Exit Sub
        p = .SelectedItems(1)
    End With
    SKW_SetReportsFolder p
    SKW_Say "Reports are saved in " & p & "."
End Sub

' Adds the Table roles sheet to your standards workbook: one row per role (caption, header, divider, total, number, short text, long text) with its style, alignment and shading, for Table Roles Check.
Public Sub AddTableRolesSheet()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Say SKW_AddTableRolesSheet()
End Sub

' ---------------------------------------------------------------------
Public Function SKW_ChooseStandards(ByVal p As String) As String
    If Len(Dir$(p)) = 0 Then SKW_ChooseStandards = "Not found: " & p: Exit Function
    SKW_SetStandardsPath p
    If SKW_Workbook() Is Nothing Then
        SKW_SetStandardsPath ""
        SKW_ChooseStandards = "That workbook could not be read, so the tools keep their built-in defaults."
        Exit Function
    End If
    SKW_ChooseStandards = "The tools now use " & Mid$(p, InStrRev(p, "\") + 1) & " for " & ThisDocument.Name & "." & vbCr & vbCr & _
        SKW_StandardsSummary()
End Function

Public Function SKW_StandardsSummary() As String
    Dim s As String, v As Variant
    For Each v In Array("Allowed styles", "Table types", "Placeholders", "Banned phrases", "Fonts", "Acronyms", "Style map", "Page limits", "Settings")
        s = s & v & ": " & SKW_Rows(CStr(v)).Count & IIf(v = "Settings", "", ", ")
    Next
    SKW_StandardsSummary = s
End Function

' Writes the starter workbook. openAfter leaves it open in Excel to fill in.
Public Function SKW_MakeStandardsWorkbook(ByVal p As String, ByVal openAfter As Boolean) As String
    Dim xl As Object, wb As Object, ed As String
    On Error GoTo Fail
    SKW_Status "Making the standards workbook..."
    Set xl = CreateObject("Excel.Application")
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    Sheet wb, 1, "Read me", Array( _
        Array("Stationkeeper Tools standards workbook"), _
        Array("Each sheet is one kind of team rule. Row 1 holds the column names: keep them as they are."), _
        Array("Add, change or remove rows below row 1. Blank rows are ignored."), _
        Array("Yes/No columns accept Yes or No."), _
        Array("In Word, run ChooseStandardsWorkbook once to point the tools at this file. The tools read it again whenever it changes."), _
        Array("Made " & Format$(Now, "yyyy-mm-dd") & " for " & ThisDocument.Name & " by Stationkeeper Tools " & SKW_VERSION & "."))
    Dim moreSettings As Variant   ' a second list: one statement allows only 24 line continuations
    moreSettings = Array( _
        Array("DeckFooterText", "", "The footer every slide shows, such as Proprietary - Company Name (Deck Marking Check and Fix)."), _
        Array("DeckMarking", "", "A marking every slide must show somewhere, such as CUI or Proprietary (Deck Marking Check and Fix)."), _
        Array("DeckLanguage", "en-US", "The proofing language for slide text: en-US, en-GB or a number such as 1033 (Deck Language Check and Fix)."), _
        Array("DeckTableStyle", "", "The team's table style: an id such as {5C22544A-7EE6-4342-B048-85BDC9FD1C3A}, or the name of a style a table in the deck uses (Deck Table Style Check and Fix)."), _
        Array("DeckExtraColors", "", "Colors allowed besides the theme's, such as #003366; #C00000 (Deck Color Check)."), _
        Array("DeliveryCheckSkip", "", "Checks the Delivery Checks leave out, by their name in the report, such as Colors; Typography; Merged cells."))
    Sheet wb, 2, "Settings", JoinRows(Array( _
        Array("Setting", "Value", "What it does"), _
        Array("StylePrefix", "", "Styles whose names start with this (for example AA) count as team styles and are allowed."), _
        Array("CaptionStyle", "Caption", "The style captions use."), _
        Array("CaptionInFirstRow", "No", "Yes if your tables carry their caption in the first row instead of above or below."), _
        Array("KeepOnClear", "Bold, Italic, Superscript, Subscript", "Emphasis that is not hand formatting: not reported and kept when clearing. Use None to clear everything."), _
        Array("CheckTablesForDirectFormatting", "Yes", "No skips table text in the Direct Formatting Report."), _
        Array("TableWidthPercent", "100", "Width of tables when no table types are listed."), _
        Array("TableCaptionPosition", "Above", "Where table captions go: Above or Below."), _
        Array("FigureCaptionPosition", "Below", "Where figure captions go: Above or Below."), _
        Array("TableLabel", "Table", "The caption label for tables (the SEQ name)."), _
        Array("FigureLabel", "Figure", "The caption label for figures (the SEQ name)."), _
        Array("MaxKeepWithNextChain", "8", "More paragraphs than this in a row set to Keep with next are reported."), _
        Array("StyleSource", "", "Full path of the team template to bring styles in from. Blank: the document's attached template."), _
        Array("AcronymList", "", "Full path of an acronym list exported from the Stationkeeper app (Excel or CSV), read with the Acronyms sheet."), _
        Array("PageLimitTotal", "", "Page limit for the whole document. Blank: no limit."), _
        Array("BulletStyle", "List Bullet", "The team's bullet list style (Convert to Team List)."), _
        Array("NumberStyle", "List Number", "The team's numbered list style (Convert to Team List)."), _
        Array("AcronymListUsedOnly", "Yes", "Insert Acronym List: Yes lists only acronyms used in the document; No lists every team meaning."), _
        Array("DeckMinFontSize", "10", "Smallest text size allowed on slides (Deck Small Text)."), _
        Array("DeckTemplate", "", "Full path of the team's PowerPoint template (.potx) that decks are checked and conformed against. Blank: the master most slides use."), _
        Array("DeckFirstSlideNumber", "1", "The number the first slide carries (Deck Slide Number Check and Fix)."), _
        Array("DeckNumberTitleSlide", "No", "Yes: title slides show a number too. No: they do not.")), moreSettings)
    Sheet wb, 3, "Allowed styles", Array( _
        Array("Style", "Required", "Notes"), _
        Array("Normal", "", ""), Array("Heading 1", "Yes", ""), Array("Heading 2", "", ""), Array("Heading 3", "", ""), _
        Array("Heading 4", "", ""), Array("Title", "", ""), Array("Subtitle", "", ""), Array("Caption", "", ""), _
        Array("List Bullet", "", ""), Array("List Bullet 2", "", ""), Array("List Number", "", ""), Array("List Number 2", "", ""), _
        Array("List Paragraph", "", ""), Array("Quote", "", ""), Array("No Spacing", "", ""), Array("Table Grid", "", "Table style"), _
        Array("TOC Heading", "", ""), Array("TOC 1", "", ""), Array("TOC 2", "", ""), Array("TOC 3", "", ""), _
        Array("Table of Figures", "", ""), Array("Header", "", ""), Array("Footer", "", ""), Array("Footnote Text", "", ""), _
        Array("Endnote Text", "", ""))
    Sheet wb, 4, "Table types", Array( _
        Array("Type", "Table style", "Header row repeats", "Rows may break across pages", "Width percent", _
              "Header row style", "Body style", "Notes row style", "Has notes row", "Cell padding", _
              "Cell margins", "Border style", "Border width", "Border color"), _
        Array("Standard", "Table Grid", "Yes", "No", "100", "", "", "", "No", "", "", "", "", ""), _
        Array("Compact", "Table Grid", "Yes", "No", "100", "", "", "", "No", "2", "", "", "", ""))
    Sheet wb, 5, "Placeholders", Array( _
        Array("Text", "Notes"), Array("TBD", ""), Array("TBR", ""), Array("TBS", ""), Array("XXX", ""), _
        Array("[insert", "Matches [insert ...] and [Insert ...]"), Array("Lorem ipsum", ""), Array("<<", ""), Array(">>", ""), _
        Array("???", ""), Array("Click or tap here to enter text", "Empty content controls"))
    Sheet wb, 6, "Banned phrases", Array( _
        Array("Phrase", "Replacement", "Rule ID", "Reason", "Match case", "Whole word"), _
        Array("utilize", "use", "PL-01", "Plain language", "No", "Yes"), _
        Array("in order to", "to", "PL-02", "Plain language", "No", "Yes"), _
        Array("and/or", "", "PL-03", "Say which you mean", "No", "No"))
    Sheet wb, 7, "Fonts", Array(Array("Font", "Allowed", "Notes"), Array("Calibri", "Yes", ""), Array("Calibri Light", "Yes", ""), _
        Array("Cambria Math", "Yes", "Equations"), Array("Symbol", "Yes", ""))
    Sheet wb, 8, "Acronyms", Array(Array("Acronym", "Definition"), Array("RFP", "Request for Proposal"))
    Sheet wb, 9, "Ignore", Array(Array("Text", "Why"), Array("", ""))
    Sheet wb, 10, "Dictionaries", Array(Array("Path", "Notes"), Array("", "Full path to a team custom dictionary (.dic)"))
    Sheet wb, 11, "Style map", Array(Array("From", "To", "Notes"), Array("", "", "Remap Style changes every use of From to To (for example an old style to the team style)"))
    Sheet wb, 12, "Page limits", Array(Array("Heading", "Max pages", "Notes"), Array("", "", "A top-level heading (or how it starts, for example Volume I) and its page limit"))
    Sheet wb, 13, "Table roles", TableRolesRows()
    wb.Worksheets(1).Activate
    wb.SaveAs p, 51
    If openAfter And Not SKW_IsQuiet() Then
        xl.Visible = True
        xl.DisplayAlerts = True
        xl.UserControl = True
    Else
        wb.Close False
        xl.Quit
    End If
    SKW_SetStandardsPath p
    SKW_Status ""
    SKW_MakeStandardsWorkbook = "Made " & Mid$(p, InStrRev(p, "\") + 1) & " and pointed the tools at it." & vbCr & vbCr & _
        "Fill in your team's styles, table types and placeholders, save it, and the tools use it from the next run." & vbCr & _
        "Share the file with your team: each person runs ChooseStandardsWorkbook once."
    Exit Function
Fail:
    ed = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    SKW_MakeStandardsWorkbook = "The workbook could not be made: " & ed
End Function

' The starting rows of the Table roles sheet: one row per role. A blank Type means every table type.
Private Function TableRolesRows() As Variant
    TableRolesRows = Array( _
        Array("Type", "Role", "Paragraph style", "Align", "Vertical", "Shading", "Bold", "Recognize", "Left indent", "Tab stops", "Notes"), _
        Array("", "Caption", "", "", "", "", "", "", "", "", "Row 1 when it holds a caption number field. Its text and field codes are never changed."), _
        Array("", "Header", "", "Center", "Center", "", "Yes", "1", "", "", "Recognize: how many header rows follow the caption row."), _
        Array("", "Divider", "", "Left", "Center", "", "Yes", "merged", "", "", "Recognize: merged (one cell across the table), or merged; first cell only."), _
        Array("", "Total", "", "", "Center", "", "Yes", "Total;Subtotal;Grand Total", "", "", "Recognize: the words a total row's first cell starts with."), _
        Array("", "Number", "", "Right", "Center", "", "", "-", "", "", "Cells holding figures. Recognize: other texts that count as a figure, such as N/A."), _
        Array("", "Short text", "", "Center", "Center", "", "", "3", "", "", "Recognize: the most words a short cell holds."), _
        Array("", "Long text", "", "Left", "Top", "", "", "", "", "", "Every other cell."), _
        Array("", "Notes", "", "Left", "Top", "", "", "Note;Notes;Source;Sources", "", "", "A merged row whose text starts with one of these words."), _
        Array("", "Keep", "", "", "", "", "", "", "", "", "Recognize: paragraph styles left as they are inside cells, such as bullet and numbered note styles, separated by ;"), _
        Array("", "Figure caption", "", "", "", "", "", "", "", "", "Captions outside tables. Recognize: their paragraph style. Only Align, Left indent and Tab stops are set on them."), _
        Array("", "", "", "", "", "", "", "", "", "", "Align: Left, Center, Right. Vertical: Top, Center, Bottom. Shading: RRGGBB or None. Bold: Yes or No. A blank cell leaves that setting as it is."), _
        Array("", "", "", "", "", "", "", "", "", "", "Left indent and Tab stops: inches. Tab stops are left tabs; give several separated by ; and they replace the paragraph's own."), _
        Array("", "", "", "", "", "", "", "", "", "", "To give one table type its own value, add a row with the type's name in Type."), _
        Array("", "", "", "", "", "", "", "", "", "", "Borders and cell margins are on the Table types sheet. Cell margins: inches, one value or top; bottom; left; right. Border style: Single, Dotted, Dashed, Double or None. Border width: 0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5 or 6 points. Border color: RRGGBB."))
End Function

' Adds the Table roles sheet to the standards workbook the tools already use.
Public Function SKW_AddTableRolesSheet() As String
    Dim xl As Object, wb As Object, ed As String, p As String, ws As Object, has As Boolean
    p = SKW_StandardsPath()
    If Len(p) = 0 Then SKW_AddTableRolesSheet = "No standards workbook is chosen. Run Make Standards Workbook or Choose Standards Workbook first.": Exit Function
    If Len(Dir$(p)) = 0 Then SKW_AddTableRolesSheet = "Not found: " & p: Exit Function
    On Error GoTo Fail
    Set xl = CreateObject("Excel.Application")
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(p)
    If wb.ReadOnly Then
        wb.Close False
        xl.Quit
        SKW_AddTableRolesSheet = "The workbook is open somewhere else. Close it in Excel and run this again."
        Exit Function
    End If
    For Each ws In wb.Worksheets
        If LCase$(ws.Name) = "table roles" Then has = True
    Next
    If has Then
        wb.Close False
        xl.Quit
        SKW_AddTableRolesSheet = "The workbook already has a Table roles sheet. A sheet made in 1.8.0 lacks two columns: add the headings Left indent and Tab stops to row 1 to use them, and a row with Figure caption in Role for captions outside tables."
        Exit Function
    End If
    Sheet wb, wb.Worksheets.Count + 1, "Table roles", TableRolesRows()
    wb.Worksheets(1).Activate
    wb.Save
    wb.Close False
    xl.Quit
    SKW_AddTableRolesSheet = "Added the Table roles sheet to " & Mid$(p, InStrRev(p, "\") + 1) & ". Fill in the style for each role and save it."
    Exit Function
Fail:
    ed = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    SKW_AddTableRolesSheet = "The sheet could not be added: " & ed
End Function

Private Function JoinRows(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim v() As Variant, i As Long, n As Long
    ReDim v(0 To UBound(a) + UBound(b) + 1)
    For i = 0 To UBound(a): v(i) = a(i): Next
    n = UBound(a) + 1
    For i = 0 To UBound(b): v(n + i) = b(i): Next
    JoinRows = v
End Function

Private Sub Sheet(ByVal wb As Object, ByVal idx As Long, ByVal nm As String, ByVal rows As Variant)
    Dim ws As Object, r As Long, c As Long, nC As Long, v() As Variant
    If wb.Worksheets.Count < idx Then wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
    Set ws = wb.Worksheets(idx)
    ws.Name = nm
    For r = 0 To UBound(rows)
        If UBound(rows(r)) + 1 > nC Then nC = UBound(rows(r)) + 1
    Next
    ReDim v(1 To UBound(rows) + 1, 1 To nC)
    For r = 0 To UBound(rows)
        For c = 0 To UBound(rows(r))
            v(r + 1, c + 1) = rows(r)(c)
        Next
    Next
    ws.Range(ws.Cells(1, 1), ws.Cells(UBound(rows) + 1, nC)).NumberFormat = "@"
    ws.Range(ws.Cells(1, 1), ws.Cells(UBound(rows) + 1, nC)).Value = v
    If nm <> "Read me" Then
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(221, 235, 247)
        ws.Activate
        ws.Parent.Windows(1).FreezePanes = False
        ws.Range("A2").Select
        ws.Parent.Windows(1).FreezePanes = True
    Else
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 14
    End If
    ws.Columns.AutoFit
    For c = 1 To nC
        If ws.Columns(c).ColumnWidth > 60 Then ws.Columns(c).ColumnWidth = 60
    Next
End Sub

' ---------------------------------------------------------------------
' The menu entries for every tool (JSON, read by Stationkeeper Macros).
' ---------------------------------------------------------------------
Public Function SKW_ToolList() As String
    Dim s As String
    s = "{""macros"": ["
    s = s & T("SKW_Setup.MakeStandardsWorkbook", "Make Standards Workbook", "Tools setup", _
        "Makes a new team standards workbook in Excel (allowed styles, table types, placeholders, settings) with starting values, opens it to fill in, and points the tools at it.", _
        "standards, excel, setup, rules", "Template", "Template", "No, it only reports")
    s = s & "," & T("SKW_Setup.ChooseStandardsWorkbook", "Choose Standards Workbook", "Tools setup", _
        "Points the tools at your team's standards workbook. Each person does this once per template.", _
        "standards, excel, setup, rules", "Template", "Template", "No, it only reports")
    s = s & "," & T("SKW_Setup.ChooseReportsFolder", "Choose Reports Folder", "Tools setup", _
        "Chooses the folder on this PC where every report is saved as it is made, named for the report, the file it checked and the time. With no folder chosen, reports open and are not saved.", _
        "reports, folder, save, where, setup", "Template", "Template", "No, it only reports")
    s = s & "," & T("SKW_Fields.SafeUpdateAllFields", "Safe Update All Fields", "Fields", _
        "Updates every field in every part of the document (body, headers, footers, notes, text boxes), then the tables of contents and figures. Fields that pull in outside files are kept as they are. One Ctrl+Z undoes it.", _
        "update, fields, f9, toc, cross-references, refresh", "Fields, Cross-references, Headers and footers", "Whole document", "Yes")
    s = s & "," & T("SKW_Fields.FindBrokenReferences", "Find Broken References", "Fields", _
        "Report of cross-references to bookmarks that are gone, fields showing Error!, internal links that go nowhere and included files that are missing. Changes nothing.", _
        "broken, error, reference not found, bookmark, links, check", "Cross-references, Fields", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Fields.FieldCodeReport", "Field Code Report", "Fields", _
        "Report of every field with its type, code, result and page. Changes nothing.", _
        "fields, codes, inventory, list", "Fields", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Review.PreDeliveryLeftovers", "Pre-Delivery Leftovers", "Before delivery", _
        "Report of what should not go out: comments, tracked changes, hidden text, highlighting, placeholders (TBD, XXX, [insert) and author details in the file. Changes nothing.", _
        "final, delivery, comments, track changes, hidden, highlight, tbd, placeholder, metadata", "Comments and tracking, Hidden text, Text", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Styles.StyleInspector", "Style Inspector", "Styles", _
        "Shows the paragraph and character style at the cursor and any formatting applied by hand on top. With several paragraphs selected, makes a report. Changes nothing.", _
        "style, inspect, formatting, what style", "Styles", "Cursor position", "No, it only reports")
    s = s & "," & T("SKW_Styles.NonAllowedStyleReport", "Styles and Fonts Check", "Styles", _
        "One report of what is off the team's lists: styles used that are not allowed, required styles not used, and every font used with the ones not on the Fonts sheet flagged. Changes nothing.", _
        "styles, fonts, allowed, approved, typeface, audit, required", "Styles", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Styles.DirectFormattingReport", "Direct Formatting Report", "Styles", _
        "Report of paragraphs whose font or spacing was set by hand instead of by style. Bold, italic, superscript and subscript count as normal emphasis. Changes nothing.", _
        "direct formatting, manual formatting, font, spacing", "Styles, Text", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Styles.ClearDirectFormattingSelection", "Clear Direct Formatting (Selection)", "Styles", _
        "Clears formatting applied by hand in the selected text so it matches its style again. Keeps bold, italic, superscript and subscript, and list numbering. One Ctrl+Z undoes it.", _
        "clear, reset, direct formatting, ctrl+space, ctrl+q", "Styles, Text", "Selection", "Yes")
    s = s & "," & T("SKW_Tables.TableStandardsAudit", "Table Standards Audit", "Tables", _
        "Report of tables that do not follow the team table types: style, repeating header row, rows breaking across pages, width, floating or nested tables, cell styles and captions. Changes nothing.", _
        "tables, audit, header row, width, caption, check", "Tables, Captions", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Tables.ApplyTeamTableType", "Apply Team Table Type", "Tables", _
        "Sets the table at the cursor to a team table type: style, repeating header row, rows kept whole, width, cell styles and padding. One Ctrl+Z undoes it.", _
        "tables, format, table style, header row, fix", "Tables", "Table at the cursor", "Yes")
    s = s & "," & T("SKW_Layout.SectionMap", "Section Map", "Layout", _
        "Report of every section: first page, orientation, paper, margins, columns, how it starts, whether headers and footers follow the previous section, and page number restarts. Changes nothing.", _
        "sections, landscape, margins, headers, footers, page numbers, link to previous", "Pages and sections, Headers and footers", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Review.BannedPhrasesReview", "Banned Phrases Review", "Before delivery", _
        "Finds the team's banned phrases in the selection or the whole body. Ones with a replacement are changed as tracked changes; every finding gets a comment naming the rule. One Ctrl+Z undoes it.", _
        "banned, phrases, plain language, words, wording, style guide, track", "Text, Comments and tracking", "Selection or whole document", "Yes")
    s = s & "," & T("SKW_Tables.TableRolesCheck", "Table Roles Check", "Tables", _
        "Report of every table's rows by role (caption, header, divider, total, body) and its cells by kind (number, short text, long text), with the rows to look at. Changes nothing, then offers to style every table in a copy.", _
        "tables, roles, total row, divider row, numbers, alignment, format all tables, fix", "Tables", "Whole document", "No, it reports; the fix works on a copy")
    s = s & "," & T("SKW_Tables.ApplyTableRolesHere", "Apply Table Roles (Table at Cursor)", "Tables", _
        "Styles the table at the cursor by role: caption row, header row, divider rows, total rows, number cells, short text and long text. One Ctrl+Z undoes it.", _
        "tables, roles, total row, divider row, numbers, alignment, format, fix", "Tables", "Table at the cursor", "Yes")
    s = s & "," & T("SKW_Setup.AddTableRolesSheet", "Add the Table Roles Sheet", "Tools setup", _
        "Adds the Table roles sheet to your standards workbook: one row per role with its style, alignment, shading and how it is recognized.", _
        "tables, roles, standards workbook, sheet, setup", "Template", "Template", "No, it only reports")
    s = s & "," & T("SKW_Captions.CaptionAudit", "Caption Audit", "Tables", _
        "Report of tables and figures without captions, typed caption numbers, wrong labels, captions on the wrong side and repeated captions. Changes nothing.", _
        "captions, figures, tables, numbering, seq", "Captions, Tables, Figures", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Styles.ImportTeamStyles", "Import Team Styles", "Styles", _
        "Brings the team's style definitions in from the team template (or the attached template) and lists which styles changed. One Ctrl+Z undoes it.", _
        "styles, import, update styles, template, organizer", "Styles", "Whole document", "Yes")
    s = s & "," & T("SKW_Styles.RemapStyle", "Remap Style", "Styles", _
        "Changes every use of one style to another, or applies every pair on the Style map sheet. One Ctrl+Z undoes it.", _
        "styles, replace style, remap, convert, style map", "Styles", "Whole document", "Yes")
    s = s & "," & T("SKW_Layout.KeepWithNextAudit", "Pagination Check", "Layout", _
        "One report of what moves text between pages: headings and captions that do not keep with what follows, long kept runs, manual page and column breaks, empty paragraphs used as spacing and line breaks ending paragraphs. Changes nothing.", _
        "pagination, keep with next, breaks, page break, empty paragraphs, widow, orphan, spacing", "Pages and sections, Styles", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Lists.TypedListsFinder", "Typed Lists Finder", "Lists", _
        "Report of paragraphs that start with a typed bullet or number instead of a real Word list. Changes nothing.", _
        "lists, bullets, numbering, typed, manual numbers", "Lists", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Layout.TOCCheck", "TOC Check", "Fields", _
        "Report comparing the table of contents with the headings: headings missing, entries with no heading, page numbers out of date. Changes nothing.", _
        "toc, contents, headings, page numbers", "Fields, Pages and sections", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Acronyms.AcronymQuickCheck", "Acronym Quick Check (text only)", "Acronyms", _
        "Report of the acronyms in the typed text: uses, first use in running text, spelled out there or not, and the team meaning (Acronyms sheet or the app's export). Text only: pictures are not scanned; the Stationkeeper app does the full check. Changes nothing.", _
        "acronyms, abbreviations, first use, define, spelled out", "Acronyms, Text", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Text.TypographySweep", "Typography Sweep", "Text", _
        "Report of typing slips: double spaces, spaces before punctuation, straight quotes, double or spaced hyphens, three dots and spaces at paragraph ends. Changes nothing.", _
        "typography, double spaces, quotes, dashes, ellipsis, typos", "Text", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Text.LoadTeamDictionaries", "Load Team Dictionaries", "Tools setup", _
        "Adds the team's custom dictionaries (Dictionaries sheet) to Word on this PC, so team words are not marked as misspelled.", _
        "dictionary, spelling, custom dictionary, proofing", "Template", "Template", "No, it changes Word's dictionary list")
    s = s & "," & T("SKW_Layout.PageLimitCounter", "Page Limit Counter", "Layout", _
        "Report of the pages in each top-level part and the whole document against the team's page limits. Changes nothing.", _
        "page limit, page count, volumes, pages", "Pages and sections", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Layout.GraphicsAudit", "Graphics Audit", "Layout", _
        "Report of every picture, chart and shape: inline or floating, size, too wide, scaled up, no alt text, linked file missing. Changes nothing.", _
        "graphics, pictures, images, alt text, accessibility, figures", "Figures", "Whole document", "No, it only reports")
    s = s & "," & T("SKW_Output.MetadataCleanCopy", "Metadata Clean Copy", "Before delivery", _
        "Saves a copy beside the document without author, company, template path, document variables and other personal details. The open document is not changed.", _
        "metadata, clean, properties, author, personal information, delivery", "Whole document", "Whole document", "No, it saves a clean copy")
    s = s & "," & T("SKW_Output.PDFWithBookmarks", "PDF with Bookmarks", "Before delivery", _
        "Saves a PDF beside the document with bookmarks from the headings and tags for screen readers, without the file's personal details.", _
        "pdf, export, bookmarks, accessible, tagged", "Whole document", "Whole document", "No, it saves a PDF")
    s = s & "," & T("SKW_Styles.DeleteUnusedStyles", "Delete Unused Styles", "Styles", _
        "Clears the Styles pane down to what the document uses: lists the unused custom styles to delete and the unused built-in styles to hide (Word does not delete built-in styles), then asks before doing it. One Ctrl+Z undoes it.", _
        "styles, delete, unused, clean, styles pane, hide, purge", "Styles", "Whole document", "Yes, after it asks")
    s = s & "," & T("SKW_Styles.RemoveUnusedCustomStyles", "Remove Unused Custom Styles", "Styles", _
        "Deletes custom styles nobody uses (not applied, not on the team list, not the base or next style of another). Lists what it deleted. One Ctrl+Z undoes it.", _
        "styles, cleanup, unused, delete styles", "Styles", "Whole document", "Yes")
    s = s & "," & T("SKW_Utils.UnlinkFieldsSelection", "Unlink Fields (Selection)", "Fields", _
        "Turns the fields in the selected text into plain text, so they never update again. Tables of contents and indexes stay fields. One Ctrl+Z undoes it.", _
        "unlink, fields, plain text, freeze", "Fields", "Selection", "Yes")
    s = s & "," & T("SKW_Utils.LockFieldsSelection", "Lock or Unlock Fields (Selection)", "Fields", _
        "Locks the fields in the selected text so updating fields leaves them as they are. When every field there is already locked, it unlocks them. One Ctrl+Z undoes it.", _
        "lock, fields, freeze, ctrl+f11", "Fields", "Selection", "Yes")
    s = s & "," & T("SKW_Utils.ConvertToTeamList", "Convert to Team List", "Lists", _
        "Turns selected paragraphs that start with a typed bullet or number into the team list styles and takes the typed marker out. One Ctrl+Z undoes it.", _
        "lists, bullets, numbering, convert, typed", "Lists", "Selection", "Yes")
    s = s & "," & T("SKW_Utils.FitAllTables", "Fit All Tables", "Tables", _
        "Fits every table wider than the margins to the page width (or only the tables in the selection). One Ctrl+Z undoes it.", _
        "tables, fit, width, margins, autofit", "Tables", "Selection or whole document", "Yes")
    s = s & "," & T("SKW_Utils.GoToNextFinding", "Go to Next Finding", "Tools setup", _
        "Selects the next finding of the last report you ran, in the document it checked. Run it again for the one after.", _
        "next, finding, report, go to, jump", "Text", "Cursor position", "No, it only changes the view")
    s = s & "," & T("SKW_Utils.InstallToolShortcuts", "Install Tool Shortcuts", "Tools setup", _
        "Sets Ctrl+Alt+Shift shortcuts for the most used tools in this template, leaving keys already in use alone.", _
        "shortcuts, keyboard, keys", "Template", "Template", "It changes the template")
    s = s & "," & T("SKW_Utils.InsertAcronymList", "Insert Acronym List", "Acronyms", _
        "Inserts an Acronym and Definition table at the cursor: the acronyms used in this document that have a team meaning, A to Z. One Ctrl+Z undoes it.", _
        "acronyms, list, table, insert, glossary", "Acronyms", "Cursor position", "Yes")
    s = s & "," & T("SKW_Utils.ResetParagraphToStyle", "Reset Paragraph to Style", "Styles", _
        "Puts the paragraph at the cursor or the selected paragraphs back to their style, removing all formatting set by hand. One Ctrl+Z undoes it.", _
        "reset, paragraph, style, ctrl+q, ctrl+space", "Styles", "Cursor position", "Yes")
    s = s & "," & T("SKP_Deck.DeckTextLeftovers", "Deck Leftovers", "PowerPoint", _
        "Report of speaker notes, comments, hidden slides and placeholder text in the deck open in PowerPoint. Changes nothing.", _
        "powerpoint, deck, notes, comments, hidden slides, placeholders, delivery", "Comments and tracking, Text", "Whole document", "No, it only reports")
    s = s & "," & T("SKP_Deck.DeckBannedPhrases", "Deck Banned Phrases", "PowerPoint", _
        "Report of the team's banned phrases in the deck. When it finds any, it offers to fix a copy: phrases with a replacement are replaced and each changed slide gets a comment. The open deck is not changed.", _
        "powerpoint, deck, banned, phrases, plain language", "Text", "Whole document", "No, it reports; a fix goes to a copy")
    s = s & "," & T("SKP_Deck.DeckAcronymCheck", "Deck Acronym Check (text only)", "PowerPoint", _
        "Report of the acronyms in the deck's text and whether each is spelled out. When some are not, it offers to fix a copy: each is spelled out at its first use in body text, with a slide comment. The open deck is not changed.", _
        "powerpoint, deck, acronyms", "Acronyms, Text", "Whole document", "No, it reports; a fix goes to a copy")
    s = s & "," & T("SKP_Deck.DeckFontAudit", "Deck Font Audit", "PowerPoint", _
        "Report of every font in the deck (slides, notes, masters, layouts, SmartArt), flagging ones not on the Fonts sheet. When it finds any, it offers to fix a copy: that text goes back to the theme fonts. The open deck is not changed.", _
        "powerpoint, deck, fonts", "Text", "Whole document", "No, it reports; a fix goes to a copy")
    s = s & "," & T("SKP_Deck.DeckSmallText", "Deck Small Text", "PowerPoint", _
        "Report of slide text smaller than the team minimum (DeckMinFontSize). Changes nothing.", _
        "powerpoint, deck, font size, readability", "Text", "Whole document", "No, it only reports")
    s = s & "," & T("SKP_Deck.DeckAltTextAudit", "Deck Alt Text Audit", "PowerPoint", _
        "Report of pictures, charts, SmartArt, groups and media in the open deck without alt text. Changes nothing.", _
        "powerpoint, deck, alt text, accessibility, pictures", "Figures", "Whole document", "No, it only reports")
    s = s & "," & T("SKP_Deck.DeckOffSlideObjects", "Deck Off-Slide Objects", "PowerPoint", _
        "Report of shapes partly or wholly off the slide in the open deck. Changes nothing.", _
        "powerpoint, deck, off slide, shapes, layout", "Figures", "Whole document", "No, it only reports")
    s = s & "," & T("SKP_Deck.DeckTitlesCheck", "Deck Titles Check", "PowerPoint", _
        "Report of slides without a title and slides that share one. Changes nothing.", _
        "powerpoint, deck, titles, accessibility, outline", "Text", "Whole document", "No, it only reports")
    s = s & "," & T("SKP_Deck.DeckEmptyPlaceholders", "Deck Empty Placeholders", "PowerPoint", _
        "Report of empty placeholders left on slides of the open deck. Changes nothing.", _
        "powerpoint, deck, placeholders, click to add", "Text", "Whole document", "No, it only reports")
    s = s & "," & T("SKP_Deck.DeckDeliveryCopy", "Deck Delivery Copy", "PowerPoint", _
        "Saves a copy of the open deck beside it with no speaker notes, comments, hidden slides or personal details. The open deck is not changed.", _
        "powerpoint, deck, delivery, clean, notes, metadata", "Whole document", "Whole document", "No, it saves a clean copy")
    s = s & "," & T("SKP_Fix.DeckMediaList", "Deck Media List", "PowerPoint", _
        "Lists every picture, video, audio, chart, SmartArt, icon and embedded or linked object in the open deck by slide. Rename them (the Selection Pane shows the names), name them from their alt text, and jump to them.", _
        "powerpoint, deck, media, pictures, video, audio, rename, selection pane, images", "Figures", "Whole document", "Only shape names, when you rename")
    s = s & "," & T("SKP_Fix.DeckTemplateCheck", "Deck Template Check", "PowerPoint", _
        "Report of every master and layout against the team template: pasted-in copies, foreign masters, unused ones, missing template layouts. When it finds any, it offers to fix a copy: move slides onto the template layouts and remove what is unused, or only remove unused layouts. The open deck is not changed.", _
        "powerpoint, deck, master, layouts, template, bloat, pasted slides", "Whole document", "Whole document", "No, it reports; a fix goes to a copy")
    s = s & "," & T("SKP_Fix.DeckOverlapCheck", "Deck Overlap Check", "PowerPoint", _
        "Report of slide content covering template elements (slide number, footer, date, marking banner, logos, title area), plus a copy with each overlap outlined in red. The open deck is not changed.", _
        "powerpoint, deck, overlap, covers, logo, footer, marking, banner", "Figures", "Whole document", "No, it reports and marks a copy")
    s = s & "," & T("SKP_Fix.DeckSlideNumberCheck", "Deck Slide Number Check", "PowerPoint", _
        "Report of slides with no number, layouts with no number placeholder, numbers typed as text and numbers moved or restyled. When it finds any, it offers to fix a copy. The open deck is not changed.", _
        "powerpoint, deck, slide numbers, page numbers, footer", "Page numbers", "Whole document", "No, it reports; a fix goes to a copy")
    s = s & "," & T("SKP_Tidy.DeckMarkingCheck", "Deck Marking Check", "PowerPoint", _
        "Report of slides without the team footer (DeckFooterText) or the marking (DeckMarking, such as CUI or Proprietary). When some lack the footer, it offers to fix a copy: the team footer goes on every slide. The open deck is not changed.", _
        "powerpoint, deck, marking, cui, proprietary, footer, banner, classification", "Headers and footers", "Whole document", "No, it reports; a fix goes to a copy")
    s = s & "," & T("SKP_Tidy.DeckTextOverflowCheck", "Deck Text Overflow Check", "PowerPoint", _
        "Report of text running past the bottom of its box or off the slide, and tables running off the slide. Changes nothing.", _
        "powerpoint, deck, overflow, text, fit, off slide", "Text", "Whole document", "No, it only reports")
    s = s & "," & T("SKP_Tidy.DeckLayoutDriftCheck", "Deck Layout Drift Check", "PowerPoint", _
        "Report of titles and other placeholders moved, resized or restyled away from their layout. When it finds any, it offers to fix a copy: they go back where the layout has them. The open deck is not changed.", _
        "powerpoint, deck, layout, title, drift, moved, placeholder, reset", "Layout", "Whole document", "No, it reports; a fix goes to a copy")
    s = s & "," & T("SKP_Tidy.DeckPictureCheck", "Deck Picture Check", "PowerPoint", _
        "Report of stretched pictures, pictures enlarged past 150 percent, and cropped pictures whose cut-off parts still travel in the file. When some are stretched, it offers to fix a copy. The open deck is not changed.", _
        "powerpoint, deck, pictures, images, stretched, distorted, cropped, resolution", "Figures", "Whole document", "No, it reports; a fix goes to a copy")
    s = s & "," & T("SKP_Tidy.DeckColorCheck", "Deck Color Check", "PowerPoint", _
        "Report of fill, line and text colors that are not the theme colors (or DeckExtraColors), with the nearest theme color. Changes nothing.", _
        "powerpoint, deck, colors, theme, brand, palette", "Whole document", "Whole document", "No, it only reports")
    s = s & "," & T("SKP_Tidy.DeckLanguageCheck", "Deck Language Check", "PowerPoint", _
        "Report of text marked with a proofing language other than the team language (DeckLanguage, en-US by default). When it finds any, it offers to fix a copy. The open deck is not changed.", _
        "powerpoint, deck, language, proofing, spelling, dictionary", "Text", "Whole document", "No, it reports; a fix goes to a copy")
    s = s & "," & T("SKP_Tidy.DeckLinksCheck", "Deck Links Check", "PowerPoint", _
        "Report of linked pictures, objects and media, hyperlinks and embedded objects: missing files, links into one computer or someone's own folders, embedded files carrying their data. Changes nothing.", _
        "powerpoint, deck, links, hyperlinks, embedded, ole, broken, delivery", "Whole document", "Whole document", "No, it only reports")
    s = s & "," & T("SKP_Tidy.DeckTableStyleCheck", "Deck Table Style Check", "PowerPoint", _
        "Report of tables not in the team table style (DeckTableStyle) and tables with no header row. When some are off style, it offers to fix a copy. The open deck is not changed.", _
        "powerpoint, deck, tables, table style, header row", "Tables", "Whole document", "No, it reports; a fix goes to a copy")
    s = s & "," & T("SKW_Review.DeliveryCheck", "Delivery Check", "Before delivery", _
        "Runs the document's checks in one go and makes one report: leftovers, broken references, tables of contents, styles, fonts, direct formatting, tables, captions, pagination, page limits, graphics, typed lists, acronyms and typography. Leave checks out with the DeliveryCheckSkip setting. Changes nothing.", _
        "delivery, final, check, all, preflight, review, before delivery", "Whole document", "Whole document", "No, it only reports")
    s = s & "," & T("SKP_Deck.DeckDeliveryCheck", "Deck Delivery Check", "PowerPoint", _
        "Runs every deck check in one go and makes one report, from leftovers and banned phrases to template, slide numbers, marking, pictures and table style. Leave checks out with the DeliveryCheckSkip setting. Changes nothing.", _
        "powerpoint, deck, delivery, final, check, all, preflight", "Whole document", "Whole document", "No, it only reports")
    s = s & "," & T("SKX_Book.BookCheck", "Book Check", "Excel", _
        "One report on the workbook open in Excel: formula errors, broken names, links to other files, hidden sheets, rows and columns, leftovers, banned phrases, acronyms, fonts and merged cells. Changes nothing.", _
        "excel, workbook, check, delivery, errors, names, links, hidden, comments, banned, acronyms, fonts, merged", "Whole document", "Whole document", "No, it only reports")
    s = s & "," & T("SKX_Book.BookDeliveryCopy", "Book Delivery Copy", "Excel", _
        "Saves a copy of the open workbook beside it with notes and comments out, links to other files turned into values, and personal details out. The open workbook is not changed.", _
        "excel, workbook, delivery, clean, comments, links, metadata", "Whole document", "Whole document", "No, it saves a clean copy")
    s = s & "]}"
    SKW_ToolList = s
End Function

Private Function T(ByVal run As String, ByVal title As String, ByVal cat As String, ByVal desc As String, _
                   ByVal kw As String, ByVal affects As String, ByVal works As String, ByVal changes As String) As String
    T = "{""Macro"": """ & run & """, ""Title"": """ & title & """, ""Category"": """ & cat & """, ""Description"": """ & desc & _
        """, ""Keywords"": """ & kw & """, ""Affects"": """ & affects & """, ""Works on"": """ & works & _
        """, ""Changes"": """ & changes & """, ""Show in menu"": ""Yes"", ""Stationkeeper"": """ & SkKind(run) & """}"
End Function

' Tools that read team rules (the standards workbook now, the effort's data once linked) are
' "Built for Stationkeeper": they run on their own. The rest are general tools with no tie.
Private Function SkKind(ByVal run As String) As String
    Const BUILT As String = "|MakeStandardsWorkbook|ChooseStandardsWorkbook|PreDeliveryLeftovers|StyleInspector|" & _
        "NonAllowedStyleReport|DirectFormattingReport|ClearDirectFormattingSelection|TableStandardsAudit|" & _
        "ApplyTeamTableType|BannedPhrasesReview|CaptionAudit|ImportTeamStyles|RemapStyle|KeepWithNextAudit|" & _
        "AcronymQuickCheck|LoadTeamDictionaries|PageLimitCounter|RemoveUnusedCustomStyles|DeleteUnusedStyles|" & _
        "ConvertToTeamList|InsertAcronymList|DeliveryCheck|DeckTextLeftovers|DeckBannedPhrases|DeckAcronymCheck|DeckFontAudit|" & _
        "DeckSmallText|DeckTemplateCheck|DeckSlideNumberCheck|DeckMarkingCheck|DeckLanguageCheck|DeckTableStyleCheck|DeckColorCheck|" & _
        "DeckDeliveryCheck|BookCheck|TableRolesCheck|ApplyTableRolesHere|AddTableRolesSheet|"
    If InStr(1, BUILT, "|" & Mid$(run, InStr(run, ".") + 1) & "|", vbTextCompare) > 0 Then
        SkKind = "Built for Stationkeeper"
    Else
        SkKind = "Not related"
    End If
End Function
