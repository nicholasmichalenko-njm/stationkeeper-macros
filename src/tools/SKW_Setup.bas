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
    sug = ThisDocument.Path & "\" & Left$(ThisDocument.Name, InStrRev(ThisDocument.Name & ".", ".") - 1) & " standards.xlsx"
    With Application.FileDialog(2)
        .Title = "Where should the standards workbook go?"
        .InitialFileName = sug
        If .Show <> -1 Then Exit Sub
        p = .SelectedItems(1)
    End With
    If LCase$(Right$(p, 5)) <> ".xlsx" Then p = Left$(p, InStrRev(p & ".", ".") - 1) & ".xlsx"
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

' ---------------------------------------------------------------------
Public Function SKW_ChooseStandards(ByVal p As String) As String
    If Len(Dir$(p)) = 0 Then SKW_ChooseStandards = "Not found: " & p: Exit Function
    SKW_SetStandardsPath p
    If Not SKW_HasStandards() Then
        SKW_SetStandardsPath ""
        SKW_ChooseStandards = "That workbook could not be read, so the tools keep their built-in defaults."
        Exit Function
    End If
    SKW_ChooseStandards = "The tools now use " & Mid$(p, InStrRev(p, "\") + 1) & " for " & ThisDocument.Name & "." & vbCr & vbCr & _
        SKW_StandardsSummary()
End Function

Public Function SKW_StandardsSummary() As String
    Dim s As String, v As Variant
    For Each v In Array("Allowed styles", "Table types", "Placeholders", "Banned phrases", "Fonts", "Acronyms", "Settings")
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
    Sheet wb, 2, "Settings", Array( _
        Array("Setting", "Value", "What it does"), _
        Array("StylePrefix", "", "Styles whose names start with this (for example AA) count as team styles and are allowed."), _
        Array("CaptionStyle", "Caption", "The style captions use."), _
        Array("CaptionInFirstRow", "No", "Yes if your tables carry their caption in the first row instead of above or below."), _
        Array("KeepOnClear", "Bold, Italic, Superscript, Subscript", "Emphasis that is not hand formatting: not reported and kept when clearing. Use None to clear everything."), _
        Array("CheckTablesForDirectFormatting", "Yes", "No skips table text in the Direct Formatting Report."), _
        Array("TableWidthPercent", "100", "Width of tables when no table types are listed."))
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
              "Header row style", "Body style", "Notes row style", "Has notes row", "Cell padding"), _
        Array("Standard", "Table Grid", "Yes", "No", "100", "", "", "", "No", ""), _
        Array("Compact", "Table Grid", "Yes", "No", "100", "", "", "", "No", "2"))
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
    s = s & "," & T("SKW_Styles.NonAllowedStyleReport", "Non-Allowed Style Report", "Styles", _
        "Report of styles used that are not on the team's allowed list, and required styles that are missing. Without a standards workbook it lists custom styles. Changes nothing.", _
        "styles, allowed, compliance, custom styles", "Styles", "Whole document", "No, it only reports")
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
    s = s & "]}"
    SKW_ToolList = s
End Function

Private Function T(ByVal run As String, ByVal title As String, ByVal cat As String, ByVal desc As String, _
                   ByVal kw As String, ByVal affects As String, ByVal works As String, ByVal changes As String) As String
    T = "{""Macro"": """ & run & """, ""Title"": """ & title & """, ""Category"": """ & cat & """, ""Description"": """ & desc & _
        """, ""Keywords"": """ & kw & """, ""Affects"": """ & affects & """, ""Works on"": """ & works & _
        """, ""Changes"": """ & changes & """, ""Show in menu"": ""Yes""}"
End Function
