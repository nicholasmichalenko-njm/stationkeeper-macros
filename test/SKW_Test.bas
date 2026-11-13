Attribute VB_Name = "SKW_Test"
' Test harness for Stationkeeper Tools. Imported into the sample template by test.ps1
' after the tools are installed. Every function returns text for test.ps1 to check.
Option Explicit
Private mPP As Object
Private mPPMine As Boolean
Private mXL As Object

Private Function ReportText() As String
    If SKW_LastReport Is Nothing Then ReportText = "(no report)": Exit Function
    ReportText = Replace(Replace(SKW_LastReport.Content.Text, Chr(13) & Chr(7), " | "), vbCr, " / ")
End Function

Private Sub CloseReport()
    On Error Resume Next
    If Not SKW_LastReport Is Nothing Then SKW_LastReport.Close SaveChanges:=0
    Set SKW_LastReport = Nothing
End Sub

Private Function AddPara(ByVal d As Object, ByVal txt As String, Optional ByVal styleName As String = "Normal") As Object
    Dim r As Object
    Set r = d.Content
    r.Collapse 0
    r.InsertAfter txt & vbCr
    Set r = d.Paragraphs(d.Paragraphs.Count - 1).Range
    r.Style = d.Styles(styleName)
    Set AddPara = r
End Function

Private Function AtEnd(ByVal d As Object) As Object
    Dim r As Object
    Set r = d.Content
    r.Collapse 0
    r.MoveEnd wdCharacter, -1
    Set AtEnd = r
End Function

Private Function FindText(ByVal d As Object, ByVal txt As String) As Object
    Dim r As Object
    Set r = d.Content
    With r.Find
        .ClearFormatting
        .Text = txt
        .MatchCase = True
        .Forward = True
        .Wrap = wdFindStop
    End With
    If r.Find.Execute Then Set FindText = r
End Function

' Steps on Word's Undo list for the active document (the Undo button's list).
Public Function TW_UndoList() As String
    Dim c As Object, i As Long, s As String
    Set c = Application.CommandBars.FindControl(, 128)
    For i = 1 To c.ListCount
        If i <= 3 Then s = s & " | " & c.List(i)
    Next
    TW_UndoList = c.ListCount & s
End Function

' Runs every report tool. They must leave the Undo list as it was.
Public Function TW_AllReports() As String
    Dim s As String, v As Variant
    SKW_Quiet = True
    For Each v In Array("SKW_FindBrokenReferences", "SKW_FieldCodeReport", "SKW_PreDeliveryLeftovers", "SKW_NonAllowedStyleReport", _
                        "SKW_DirectFormattingReport", "SKW_TableStandardsAudit", "SKW_StyleInspector", "SKW_SectionMap", "SKW_CaptionAudit", _
                        "SKW_KeepWithNextAudit", "SKW_BreakHygiene", "SKW_TypedListsFinder", "SKW_TOCCheck", "SKW_AcronymQuickCheck", _
                        "SKW_TypographySweep", "SKW_FontNotAllowed", "SKW_PageLimitCounter", "SKW_GraphicsAudit")
        s = s & v & ": " & Application.Run(CStr(v)) & " / "
        CloseReport
    Next
    TW_AllReports = s
End Function

Public Function TW_ToolCount() As String
    Dim m As Object, e As Object, n As Long, s As String
    Set m = SK_LoadModel()
    For Each e In m("entries")
        If e("run") Like "SK[WPX]_*" Then
            n = n + 1
            s = s & " " & e("run") & "[" & SK_EntryStatus(e) & "/" & e("category") & "/" & e("workson") & "/" & e("shown") & "]"
        End If
    Next
    TW_ToolCount = n & " |" & s
End Function

Public Function TW_MakeStandards(ByVal path As String) As String
    SKW_Quiet = True
    TW_MakeStandards = SKW_MakeStandardsWorkbook(path, False) & " ## " & SKW_StandardsSummary() & " ## " & SKW_StdNote()
End Function

Public Function TW_NoStandards() As String
    Dim p As String
    p = SKW_StandardsPath()
    SKW_SetStandardsPath ""
    TW_NoStandards = p
End Function

Public Function TW_UseStandards(ByVal p As String) As String
    SKW_Quiet = True
    TW_UseStandards = SKW_ChooseStandards(p)
End Function

' Builds a document with one of each problem the tools look for.
Public Function TW_BuildDefects() As String
    Dim d As Object, r As Object, st As Object, t As Object, i As Long, f As Object
    SKW_Quiet = True
    Set d = ActiveDocument
    d.TrackRevisions = False
    d.Content.Delete
    d.Content.InsertAfter "Contents" & vbCr
    d.Paragraphs(1).Style = d.Styles(wdStyleNormal)
    Set r = AddPara(d, "Introduction", "Heading 1")
    d.Bookmarks.Add "Intro", d.Paragraphs(2).Range
    AddPara d, "Good paragraph with a bold word and H2O written with a subscript."
    Set r = FindText(d, "bold"): r.Font.Bold = True
    Set r = FindText(d, "H2O"): r.Characters(2).Font.Subscript = True
    Set r = AddPara(d, "Paragraph typed in Arial by hand with kept bold.")
    r.MoveEnd wdCharacter, -1
    r.Font.Name = "Arial": r.Font.Size = 14
    Set r = FindText(d, "kept"): r.Font.Bold = True
    Set r = AddPara(d, "Paragraph with extra space after.")
    r.ParagraphFormat.SpaceAfter = 30
    On Error Resume Next
    Set st = d.Styles("Fancy Body")
    On Error GoTo 0
    If st Is Nothing Then
        Set st = d.Styles.Add("Fancy Body", wdStyleTypeParagraph)
        st.BaseStyle = "Normal"
        st.Font.Name = "Courier New"
    End If
    AddPara d, "Custom style paragraph.", "Fancy Body"
    Set r = AddPara(d, "Gone target text.")
    r.MoveEnd wdCharacter, -1
    d.Bookmarks.Add "Gone", r
    AddPara d, "Good ref: X. Broken ref: Y."
    Set r = FindText(d, "X.")
    r.MoveEnd wdCharacter, -1
    d.Fields.Add r, wdFieldEmpty, "REF Intro \h", False
    Set r = FindText(d, "Y.")
    r.MoveEnd wdCharacter, -1
    d.Fields.Add r, wdFieldEmpty, "REF Gone \h", False
    d.Bookmarks("Gone").Delete
    AddPara d, "Placeholder: TBD and [insert name]."
    Set r = AddPara(d, "Secret hidden note.")
    r.MoveEnd wdCharacter, -1
    r.Font.Hidden = True
    Set r = AddPara(d, "Highlighted text.")
    r.MoveEnd wdCharacter, -1
    r.HighlightColorIndex = wdYellow
    Set r = AddPara(d, "Commented text.")
    d.Comments.Add r, "Please check this."
    d.TrackRevisions = True
    AtEnd(d).InsertAfter "Inserted with tracking." & vbCr
    d.TrackRevisions = False
    AddPara d, "Included: Z."
    Set r = FindText(d, "Z.")
    r.MoveEnd wdCharacter, -1
    ' Added as a QUOTE field and recoded, so Word never tries the outside file (that asks a security question).
    Set f = d.Fields.Add(r, wdFieldEmpty, "QUOTE ""Old included text""", False)
    f.Code.Text = " INCLUDETEXT ""C:\\NoSuchFolder\\missing.docx"" "
    AddPara d, "Link to nowhere."
    Set r = FindText(d, "nowhere")
    d.Hyperlinks.Add r, "", "NoSuchPlace"
    ' Good table with a caption above.
    AddPara d, "Table : Good table", "Caption"
    Set r = FindText(d, "Table :")
    r.Collapse 0
    r.Move wdCharacter, -1
    d.Fields.Add r, wdFieldEmpty, "SEQ Table \* ARABIC", False
    Set r = AddPara(d, "")
    Set t = d.Tables.Add(r, 3, 3)
    t.Style = "Table Grid"
    t.Rows(1).HeadingFormat = True
    t.Rows.AllowBreakAcrossPages = False
    t.AutoFitBehavior wdAutoFitWindow
    t.Cell(1, 1).Range.Text = "Good head"
    ' Bad table: no caption, long, header does not repeat, rows break, wider than the margins.
    AddPara d, "Text between tables."
    Set r = AddPara(d, "")
    Set t = d.Tables.Add(r, 60, 3)
    t.Style = "Table Normal"
    t.Rows(1).HeadingFormat = False
    t.Rows.AllowBreakAcrossPages = True
    t.PreferredWidthType = wdPreferredWidthPoints
    t.PreferredWidth = 648
    For i = 1 To 3
        t.Columns(i).Width = 216
    Next
    t.Cell(1, 1).Range.Text = "Bad head"
    ' Floating table.
    AddPara d, "More text."
    Set r = AddPara(d, "")
    Set t = d.Tables.Add(r, 2, 2)
    t.Style = "Table Grid"
    t.Rows.WrapAroundText = True
    t.Cell(1, 1).Range.Text = "Float head"
    AddPara d, "The end."
    ' Contents at the top, page number in the footer.
    Set r = d.Paragraphs(1).Range
    r.MoveEnd wdCharacter, -1
    d.TablesOfContents.Add r, True, 1, 3
    d.Sections(1).Footers(wdHeaderFooterPrimary).Range.Fields.Add d.Sections(1).Footers(wdHeaderFooterPrimary).Range, wdFieldPage, "", False
    d.Range(0, 0).Select
    TW_BuildDefects = "paras=" & d.Paragraphs.Count & " tables=" & d.Tables.Count & " fields=" & d.Fields.Count & _
        " pages=" & d.ComputeStatistics(wdStatisticPages)
End Function

Public Function TW_Broken() As String
    Dim s As String
    SKW_Quiet = True
    s = SKW_FindBrokenReferences() & " ## " & ReportText()
    CloseReport
    TW_Broken = s
End Function

Public Function TW_FieldReport() As String
    Dim s As String
    SKW_Quiet = True
    s = SKW_FieldCodeReport() & " ## " & ReportText()
    CloseReport
    TW_FieldReport = s
End Function

Public Function TW_Leftovers() As String
    Dim s As String
    SKW_Quiet = True
    s = SKW_PreDeliveryLeftovers() & " ## " & ReportText()
    CloseReport
    TW_Leftovers = s
End Function

Public Function TW_Inspector() As String
    SKW_Quiet = True
    FindText(ActiveDocument, "typed in Arial").Select
    Selection.Collapse wdCollapseStart
    TW_Inspector = Replace(SKW_StyleInspector(), vbCr, " / ")
End Function

Public Function TW_NonAllowed() As String
    Dim s As String
    SKW_Quiet = True
    s = SKW_NonAllowedStyleReport() & " ## " & ReportText()
    CloseReport
    TW_NonAllowed = s
End Function

Public Function TW_Direct() As String
    Dim s As String
    SKW_Quiet = True
    s = SKW_DirectFormattingReport() & " ## " & ReportText()
    CloseReport
    TW_Direct = s
End Function

Public Function TW_Clear() As String
    Dim d As Object, r As Object, s As String, p As Object
    SKW_Quiet = True
    Set d = ActiveDocument
    Set r = FindText(d, "Paragraph typed in Arial")
    r.Expand wdParagraph
    Set p = FindText(d, "Paragraph with extra space after.")
    r.End = p.End
    r.Select
    s = SKW_ClearDirectFormattingSelection()
    s = s & " ## font=" & FindText(d, "typed in Arial").Font.Name & " size=" & FindText(d, "typed in Arial").Font.Size & _
        " kept=" & FindText(d, "kept").Font.Bold & " after=" & FindText(d, "extra space after").Paragraphs(1).SpaceAfter
    d.Undo 1
    s = s & " ## undo: font=" & FindText(d, "typed in Arial").Font.Name & " after=" & FindText(d, "extra space after").Paragraphs(1).SpaceAfter
    TW_Clear = s
End Function

Public Function TW_Coauthor() As String
    Dim s As String
    SKW_Quiet = True
    SKW_ForceCoauthored = True
    FindText(ActiveDocument, "Paragraph typed in Arial").Select
    s = "clear: " & SKW_ClearDirectFormattingSelection() & " ## report: " & SKW_FindBrokenReferences()
    CloseReport
    SKW_ForceCoauthored = False
    TW_Coauthor = Replace(s, vbCr, " / ")
End Function

Public Function TW_Tables() As String
    Dim s As String
    SKW_Quiet = True
    s = SKW_TableStandardsAudit() & " ## " & ReportText()
    CloseReport
    TW_Tables = s
End Function

Public Function TW_ApplyType() As String
    Dim d As Object, t As Object, s As String
    SKW_Quiet = True
    Set d = ActiveDocument
    FindText(d, "Bad head").Select
    ' Only the run here: reading the table back inside the same macro puts a step on Word's Undo list.
    TW_ApplyType = Replace(SKW_ApplyTeamTableType("Standard"), vbCr, " / ")
End Function

' Read after test.ps1 presses Undo once. Undo inside the same macro run is not reliable in Word
' (its undo list settles after the macro ends), and a person presses Ctrl+Z after the run anyway.
Public Function TW_TableState() As String
    With ActiveDocument.Tables(2)
        TW_TableState = "style=" & .Style.NameLocal & " head=" & .Rows(1).HeadingFormat & " break=" & .Rows.AllowBreakAcrossPages & _
            " widthType=" & .PreferredWidthType & " width=" & .PreferredWidth
    End With
End Function

Public Function TW_SafeUpdate() As String
    Dim d As Object, f As Object, inc As Object, s As String, before As String, ref As String
    SKW_Quiet = True
    Set d = ActiveDocument
    For Each f In d.Fields
        If f.Type = wdFieldIncludeText Then Set inc = f
    Next
    before = inc.Result.Text
    d.TrackRevisions = True
    s = Replace(SKW_SafeUpdateAllFields(), vbCr, " / ")
    For Each f In d.Fields
        If f.Type = wdFieldRef And InStr(f.Code.Text, "Gone") > 0 Then ref = f.Result.Text
    Next
    s = s & " ## include same=" & (inc.Result.Text = before) & " locked=" & inc.Locked & " brokenref=" & Left$(ref, 6) & _
        " track=" & d.TrackRevisions & " toc=" & d.TablesOfContents.Count
    d.TrackRevisions = False
    TW_SafeUpdate = s
End Function

Public Function TW_MenuRun() As String
    Dim n As Long, ok As Boolean
    SKW_Quiet = True
    n = Documents.Count
    ok = SK_RunMacro("SKW_Fields.FieldCodeReport", "Field Code Report")
    TW_MenuRun = "ran=" & ok & " newdocs=" & (Documents.Count - n) & " state=" & SK_TestState("SKW_Fields.FieldCodeReport")
    CloseReport
End Function

Public Function TW_TestAll() As String
    Dim m As Object, n As Long, s As String
    SKW_Quiet = False      ' the tester itself asks the tools to stay quiet
    n = Documents.Count
    Set m = SK_LoadModel()
    SK_Health m
    s = Replace(SK_TestAll(m, False), vbCr, " / ")
    TW_TestAll = s & " | docs before=" & n & " after=" & Documents.Count & " quiet flag=" & GetSetting("Stationkeeper Tools", "Run", "quiet", "")
End Function

' ---------------------------------------------------------------------
' 0.5.0: batch 2 defects and checks
' ---------------------------------------------------------------------
Private Function TinyBmp() As String
    Dim p As String, f As Integer, b() As Byte, i As Long, hdr As Variant
    p = Environ$("TEMP") & "\skw_test.bmp"
    hdr = Array(66, 77, 58, 0, 0, 0, 0, 0, 0, 0, 54, 0, 0, 0, 40, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 24, 0, 0, 0, 0, 0, 4, 0, 0, 0, _
                19, 11, 0, 0, 19, 11, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 255, 0)
    ReDim b(0 To UBound(hdr))
    For i = 0 To UBound(hdr)
        b(i) = hdr(i)
    Next
    f = FreeFile
    Open p For Binary Access Write As #f
    Put #f, 1, b
    Close #f
    TinyBmp = p
End Function

Public Function TW_BuildDefects2() As String
    Dim d As Object, r As Object, bmp As String, n As Long
    SKW_Quiet = True
    Set d = ActiveDocument
    d.TrackRevisions = False
    AddPara d, "We utilize tools in order to win and/or lose."
    AddPara d, ChrW(8226) & " Typed bullet one"
    AddPara d, "1. Typed number one"
    AddPara d, "a) Typed letter one"
    Set r = AddPara(d, "Real list item")
    r.ListFormat.ApplyBulletDefault
    Set r = AddPara(d, "After the list.")
    r.ListFormat.RemoveNumbers
    AddPara d, ""
    AddPara d, ""
    AddPara d, ""
    AddPara d, "After the gap."
    Set r = AddPara(d, "Ends with a line break")
    r.MoveEnd wdCharacter, -1
    r.InsertAfter Chr(11)
    Set r = AddPara(d, "Before break")
    r.MoveEnd wdCharacter, -1
    r.Collapse 0
    r.InsertBreak wdPageBreak
    AddPara d, "After break"
    Set r = AddPara(d, "Loose heading", "Heading 2")
    r.ParagraphFormat.KeepWithNext = False
    AddPara d, "Text under the loose heading."
    AddPara d, "Heading added later", "Heading 1"
    AddPara d, "Text under the later heading."
    bmp = TinyBmp()
    Set r = AddPara(d, "")
    r.Collapse wdCollapseStart
    d.InlineShapes.AddPicture bmp, False, True, r
    AddPara d, "Figure : Good figure", "Caption"
    Set r = FindText(d, "Figure :")
    r.Collapse 0
    r.Move wdCharacter, -1
    d.Fields.Add r, wdFieldEmpty, "SEQ Figure \* ARABIC", False
    Set r = AddPara(d, "")
    r.Collapse wdCollapseStart
    d.InlineShapes.AddPicture bmp, False, True, r
    AddPara d, "Figure 7: Typed number", "Caption"
    Set r = AtEnd(d)
    r.InsertBreak wdSectionBreakNextPage
    n = d.Sections.Count
    d.Sections(n).PageSetup.Orientation = wdOrientLandscape
    d.Sections(n).Footers(wdHeaderFooterPrimary).LinkToPrevious = False
    AddPara d, "Landscape section."
    d.Range(0, 0).Select
    TW_BuildDefects2 = "sections=" & d.Sections.Count & " pictures=" & d.InlineShapes.Count & " pages=" & d.ComputeStatistics(wdStatisticPages)
End Function

Public Function TW_Report(ByVal fn As String) As String
    Dim s As String
    SKW_Quiet = True
    s = Application.Run(fn) & " ## " & ReportText()
    CloseReport
    TW_Report = s
End Function

Public Function TW_Banned() As String
    Dim d As Object, s As String, n0 As Long
    SKW_Quiet = True
    Set d = ActiveDocument
    d.Range(0, 0).Select
    n0 = d.Comments.Count
    s = Replace(SKW_BannedPhrasesReview(), vbCr, " / ")
    s = s & " ## comments+" & (d.Comments.Count - n0) & " text=" & Trim$(FindText(d, "tools").Paragraphs(1).Range.Text) & " track=" & d.TrackRevisions
    TW_Banned = s
End Function

Public Function TW_StyleCount(ByVal nm As String) As String
    TW_StyleCount = SKW_StyleRanges(ActiveDocument.Content, ActiveDocument.Styles(nm)).Count
End Function

Public Function TW_Remap() As String
    SKW_Quiet = True
    TW_Remap = Replace(SKW_RemapStyle("Fancy Body>Normal|No Such Style>Normal"), vbCr, " / ")
End Function

Public Function TW_SetNormalSize(ByVal sz As String) As String
    ActiveDocument.Styles("Normal").Font.Size = CSng(sz)
    TW_SetNormalSize = ActiveDocument.Styles("Normal").Font.Size
End Function

Public Function TW_NormalSize() As String
    TW_NormalSize = ActiveDocument.Styles("Normal").Font.Size
End Function

Public Function TW_Import() As String
    SKW_Quiet = True
    TW_Import = Replace(SKW_ImportTeamStyles(""), vbCr, " / ")
End Function

' ---------------------------------------------------------------------
' 0.6.0: batch 3
' ---------------------------------------------------------------------
Public Function TW_BuildDefects3(ByVal savePath As String) As String
    Dim d As Object, st As Object
    SKW_Quiet = True
    Set d = ActiveDocument
    d.TrackRevisions = False
    AddPara d, "Request for Proposal (RFP) text."
    AddPara d, "The SOW covers the work."
    AddPara d, "Statement of Work (SOW) is the scope."
    AddPara d, "Typos  here ,and ""straight"" quotes -- dash ... end. "
    AddPara d, "We recieve teh data."
    On Error Resume Next
    Set st = d.Styles("Unused Custom")
    On Error GoTo 0
    If st Is Nothing Then
        d.Styles.Add "Unused Custom", wdStyleTypeParagraph
        d.Styles.Add "Unused Char", wdStyleTypeCharacter
    End If
    If d.InlineShapes.Count > 0 Then d.InlineShapes(1).AlternativeText = "Red dot"
    d.BuiltInDocumentProperties("Author") = "Test Author"
    d.BuiltInDocumentProperties("Company") = "Test Company"
    d.SaveAs2 FileName:=savePath, FileFormat:=wdFormatXMLDocument
    d.Range(0, 0).Select
    TW_BuildDefects3 = "saved=" & d.FullName
End Function

Public Function TW_Dictionaries() As String
    TW_Dictionaries = Replace(SKW_LoadTeamDictionaries(), vbCr, " / ")
End Function

Public Function TW_Clean(ByVal target As String) As String
    Dim s As String, c As Object, a As String
    SKW_Quiet = True
    s = Replace(SKW_MetadataCleanCopy(target), vbCr, " / ")
    Set c = Documents.Open(FileName:=target, ReadOnly:=True, AddToRecentFiles:=False, Visible:=False)
    a = c.BuiltInDocumentProperties("Author") & "|" & c.BuiltInDocumentProperties("Company")
    c.Close 0
    TW_Clean = s & " ## props=[" & a & "] active=" & ActiveDocument.Name
End Function

Public Function TW_Pdf(ByVal target As String) As String
    SKW_Quiet = True
    TW_Pdf = Replace(SKW_PDFWithBookmarks(target), vbCr, " / ") & " ## bytes=" & FileLen(target)
End Function

Public Function TW_RemoveUnused() As String
    SKW_Quiet = True
    TW_RemoveUnused = Replace(SKW_RemoveUnusedCustomStyles(), vbCr, " / ")
End Function

Public Function TW_StyleExists(ByVal nm As String) As String
    Dim st As Object
    On Error Resume Next
    Set st = ActiveDocument.Styles(nm)
    TW_StyleExists = Not st Is Nothing
End Function

' ---------------------------------------------------------------------
' 0.7.0: utilities
' ---------------------------------------------------------------------
Public Function TW_BuildLists() As String
    Dim d As Object, r As Object, a As Object, b As Object
    SKW_Quiet = True
    Set d = ActiveDocument
    Set a = AddPara(d, "Item A1")
    AddPara d, "Item A2"
    Set r = d.Range(FindText(d, "Item A1").Start, FindText(d, "Item A2").End)
    r.ListFormat.ApplyNumberDefault
    Set r = AddPara(d, "Between the lists.")
    r.ListFormat.RemoveNumbers
    AddPara d, "Item B1"
    AddPara d, "Item B2"
    Set r = d.Range(FindText(d, "Item B1").Start, FindText(d, "Item B2").End)
    r.ListFormat.ApplyNumberDefault
    Set r = AddPara(d, "After the lists.")
    r.ListFormat.RemoveNumbers
    ' A paragraph with two plain fields for the field tools.
    AddPara d, "Fields here: A and B end."
    Set r = FindText(d, "A and")
    r.MoveEnd wdCharacter, -4
    d.Fields.Add r, wdFieldEmpty, "QUOTE ""Alpha""", False
    Set r = FindText(d, "B end")
    r.MoveEnd wdCharacter, -4
    d.Fields.Add r, wdFieldEmpty, "QUOTE ""Bravo""", False
    TW_BuildLists = "B1=" & FindText(d, "Item B1").ListFormat.ListString
End Function

Public Function TW_Fields(ByVal what As String) As String
    Dim d As Object, r As Object, n0 As Long
    SKW_Quiet = True
    Set d = ActiveDocument
    Set r = FindText(d, "Fields here:").Paragraphs(1).Range
    r.Select
    n0 = r.Fields.Count
    TW_Fields = SKW_FieldsSelection(what) & " ## before=" & n0 & " after=" & r.Fields.Count & " text=" & SKW_Excerpt(r.Text, 60)
End Function

' Read in its own run: inside the run that changed them, Word still reports the old lock state.
Public Function TW_FieldLocks() As String
    Dim f As Object
    For Each f In FindText(ActiveDocument, "Fields here:").Paragraphs(1).Range.Fields
        TW_FieldLocks = TW_FieldLocks & IIf(f.Locked, "L", "u")
    Next
End Function

Public Function TW_Convert() As String
    Dim d As Object, r As Object
    SKW_Quiet = True
    Set d = ActiveDocument
    Set r = d.Range(FindText(d, "Typed bullet one").Paragraphs(1).Range.Start, FindText(d, "Typed letter one").Paragraphs(1).Range.End)
    r.Select
    TW_Convert = Replace(SKW_ConvertToTeamList(), vbCr, " / ") & " ## " & ParaInfo(d, "Typed bullet one") & " | " & ParaInfo(d, "Typed number one") & " | " & ParaInfo(d, "Typed letter one")
End Function

Private Function ParaInfo(ByVal d As Object, ByVal t As String) As String
    Dim p As Object
    Set p = FindText(d, t).Paragraphs(1)
    ParaInfo = SKW_ParaStyleName(p.Range) & ":" & Trim$(Replace(p.Range.Text, vbCr, "")) & ":" & p.Range.ListFormat.ListString
End Function

Public Function TW_Fit() As String
    SKW_Quiet = True
    ActiveDocument.Range(0, 0).Select
    TW_Fit = Replace(SKW_FitAllTables(), vbCr, " / ")
End Function

Public Function TW_TableWidth(ByVal i As String) As String
    TW_TableWidth = TableWidth(ActiveDocument.Tables(CLng(i)))
End Function

Private Function TableWidth(ByVal t As Object) As String
    Dim c As Object, w As Single
    For Each c In t.Range.Cells
        If c.RowIndex > 1 Then Exit For
        w = w + c.Width
    Next
    TableWidth = Format$(w / 72, "0.00")
End Function

Public Function TW_NextFinding() As String
    Dim s As String
    SKW_Quiet = True
    s = SKW_TableStandardsAudit()
    CloseReport
    s = s & " ## 1: " & SKW_GoToNextFinding() & " [sel=" & SKW_Excerpt(Selection.Text, 30) & "]"
    s = s & " ## 2: " & SKW_GoToNextFinding() & " [sel=" & SKW_Excerpt(Selection.Text, 30) & "]"
    TW_NextFinding = s
End Function

Public Function TW_Shortcuts() As String
    Dim s As String, kb As Object
    s = Replace(SKW_InstallToolShortcuts(), vbCr, " / ")
    CustomizationContext = MacroContainer
    Set kb = FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyShift, wdKeyN))
    TW_Shortcuts = s & " ## N=" & kb.Command
End Function

Public Function TW_InsertAcronyms() As String
    Dim d As Object, n As Long, s As String
    SKW_Quiet = True
    Set d = ActiveDocument
    AddPara d, "Acronym list goes here."
    FindText(d, "Acronym list goes here.").Select
    Selection.Collapse wdCollapseStart
    n = d.Tables.Count
    s = Replace(SKW_InsertAcronymList(True), vbCr, " / ")
    TW_InsertAcronyms = s & " ## tables+" & (d.Tables.Count - n) & IIf(d.Tables.Count > n, " cell=" & SKW_Excerpt(d.Tables(d.Tables.Count).Range.Text, 80), "")
End Function

Public Function TW_TableCount() As String
    TW_TableCount = ActiveDocument.Tables.Count
End Function

Public Function TW_Reset() As String
    Dim d As Object, s As String
    SKW_Quiet = True
    Set d = ActiveDocument
    FindText(d, "typed in Arial").Select
    Selection.Collapse wdCollapseStart
    s = Replace(SKW_ResetParagraphToStyle(), vbCr, " / ")
    TW_Reset = s & " ## font=" & FindText(d, "typed in Arial").Font.Name & " kept=" & FindText(d, "kept").Font.Bold
End Function

Public Function TW_FontName(ByVal t As String) As String
    TW_FontName = FindText(ActiveDocument, t).Font.Name
End Function

' ---------------------------------------------------------------------
' 0.8.0: PowerPoint. The test makes its own windowless deck and hands it to the
' tools directly (SKP_TestDeck), so a person's open decks are never touched.
' ---------------------------------------------------------------------

Public Function TW_MakeDeck(ByVal path As String) As String
    Dim p As Object, s As Object, tb As Object, w As Single
    Set mPP = CreateObject("PowerPoint.Application")
    mPPMine = (mPP.Presentations.Count = 0)
    Set p = mPP.Presentations.Add(0)
    ' 1: title and text, with notes
    Set s = p.Slides.Add(1, 2)
    s.Shapes(1).TextFrame.TextRange.Text = "Proposal overview"
    s.Shapes(2).TextFrame.TextRange.Text = "We utilize the Request for Proposal (RFP) process. The RFP date is TBD."
    s.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = "Speaker note here."
    ' 2: same title, empty body
    Set s = p.Slides.Add(2, 2)
    s.Shapes(1).TextFrame.TextRange.Text = "Proposal overview"
    ' 3: blank: off-slide box, small Arial text, picture without alt text
    Set s = p.Slides.Add(3, 12)
    w = p.PageSetup.SlideWidth
    Set tb = s.Shapes.AddTextbox(1, w + 20, 50, 100, 40)
    tb.TextFrame.TextRange.Text = "Parked off the slide"
    Set tb = s.Shapes.AddTextbox(1, 50, 50, 300, 40)
    tb.TextFrame.TextRange.Text = "Tiny Arial footnote"
    tb.TextFrame.TextRange.Font.Name = "Arial"
    tb.TextFrame.TextRange.Font.Size = 8
    s.Shapes.AddPicture TinyBmp(), 0, -1, 100, 150, 50, 50
    ' 4: hidden slide with a title
    Set s = p.Slides.Add(4, 11)
    s.Shapes(1).TextFrame.TextRange.Text = "Hidden one"
    s.SlideShowTransition.Hidden = -1
    p.SaveAs path
    Set SKP_TestDeck = p
    TW_MakeDeck = "slides=" & p.Slides.Count & " mine=" & mPPMine
End Function

Public Function TW_Deck(ByVal fn As String) As String
    Dim s As String
    SKW_Quiet = True
    s = Application.Run(fn) & " ## " & ReportText()
    CloseReport
    TW_Deck = s
End Function

Public Function TW_DeckCopy(ByVal target As String) As String
    Dim s As String, c As Object, n As Long
    s = Replace(SKP_DeckDeliveryCopy(target), vbCr, " / ")
    Set c = mPP.Presentations.Open(target, -1, 0, 0)
    n = c.Slides.Count
    c.Close
    TW_DeckCopy = s & " ## copySlides=" & n & " openSlides=" & SKP_TestDeck.Slides.Count
End Function

Public Function TW_CloseDeck() As String
    On Error Resume Next
    SKP_TestDeck.Close
    Set SKP_TestDeck = Nothing
    If mPPMine And mPP.Presentations.Count = 0 Then mPP.Quit
    Set mPP = Nothing
    TW_CloseDeck = "closed"
End Function

Public Function TW_NoDeck() As String
    Dim keep As Object
    Set keep = SKP_TestDeck
    Set SKP_TestDeck = Nothing
    TW_NoDeck = "quiet entry: " & IIf(GetSetting("Stationkeeper Tools", "Run", "quiet", "") = "1", "flag", "noflag")
    Set SKP_TestDeck = keep
End Function

' ---------------------------------------------------------------------
' 0.9.0: Excel. The test makes its own hidden workbook and hands it to the
' tools directly (SKX_TestBook), so a person's open workbooks are never touched.
' ---------------------------------------------------------------------

Public Function TW_MakeBook(ByVal path As String) As String
    Dim wb As Object, ws As Object
    Set mXL = CreateObject("Excel.Application")
    mXL.Visible = False
    mXL.DisplayAlerts = False
    Set wb = mXL.Workbooks.Add
    Do While wb.Worksheets.Count < 3
        wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
    Loop
    Set ws = wb.Worksheets(1)
    ws.Name = "Pricing"
    ws.Range("A1").Value = "Item"
    ws.Range("B1").Formula = "=1/0"
    ws.Range("B2").Formula = "=NA()"
    ws.Range("A2").Value = "We utilize the Request for Proposal (RFP) rates."
    ws.Range("A3").Value = "Rate TBD"
    ws.Range("A4").Value = "Arial cell"
    ws.Range("A4").Font.Name = "Arial"
    ws.Range("A1").AddComment "Check this"
    ws.Range("A10:B10").Merge
    ws.Range("A10").Value = "Merged heading"
    ws.Rows(6).Hidden = True
    wb.Names.Add Name:="BadName", RefersTo:="=#REF!"
    wb.Worksheets(2).Name = "Notes"
    wb.Worksheets(3).Name = "Secret"
    wb.Worksheets(3).Visible = 0
    wb.SaveAs path, 51
    Set SKX_TestBook = wb
    TW_MakeBook = "sheets=" & wb.Worksheets.Count
End Function

Public Function TW_Book(ByVal fn As String) As String
    Dim s As String
    SKW_Quiet = True
    s = Application.Run(fn) & " ## " & ReportText()
    CloseReport
    TW_Book = s
End Function

Public Function TW_BookCopy(ByVal target As String) As String
    Dim s As String, c As Object, n As Long
    s = Replace(SKX_BookDeliveryCopy(target), vbCr, " / ")
    Set c = mXL.Workbooks.Open(target, 0, True)
    n = c.Worksheets("Pricing").Comments.Count
    c.Close False
    TW_BookCopy = s & " ## copyComments=" & n & " openComments=" & SKX_TestBook.Worksheets("Pricing").Comments.Count
End Function

Public Function TW_CloseBook() As String
    On Error Resume Next
    SKX_TestBook.Close False
    Set SKX_TestBook = Nothing
    mXL.Quit
    Set mXL = Nothing
    TW_CloseBook = "closed"
End Function

' ---------------------------------------------------------------------
' 1.0.0: the rules the tools use once an effort is linked.
' ---------------------------------------------------------------------
Public Function TW_Rules() As String
    Dim s As String, v As Variant, b As String, c As Collection
    SKW_Quiet = True
    SKW_ForgetEffort
    For Each v In SKW_Column("Banned phrases", "Phrase")
        b = b & IIf(Len(b) > 0, ",", "") & v
    Next
    Set c = SKW_Column("Dictionaries", "Path")
    s = "note=" & SKW_StdNote() & " | caption=" & SKW_Setting("CaptionStyle", "") & " | table=" & SKW_Setting("TableLabel", "") & _
        " | firstrow=" & SKW_Setting("CaptionInFirstRow", "") & " | banned=" & b & " | placeholders=" & SKW_Placeholders().Count & _
        " | acronyms=" & SKW_AcronymList().Count & " | source=" & SKW_Setting("StyleSource", "")
    If c.Count > 0 Then s = s & " | dic=" & c(1) Else s = s & " | dic="
    TW_Rules = s
End Function

' The local copy: which revision folders exist.
Public Function TW_Cache() As String
    Dim fso As Object, f As Object, s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(SKW_CacheDir()) Then TW_Cache = "no cache": Exit Function
    For Each f In fso.GetFolder(SKW_CacheDir()).SubFolders
        s = s & " " & f.Name
    Next
    TW_Cache = "folders:" & s & " | root=" & Mid$(SKW_CacheRoot(), InStrRev(SKW_CacheRoot(), "\") + 1)
End Function

' Banned Phrases Review on a fresh document: the effort's phrase is caught, the workbook's is not.
Public Function TW_EffortBanned() As String
    Dim d As Object, s As String
    SKW_Quiet = True
    SKW_ForgetEffort
    Set d = Documents.Add
    d.Content.Text = "We leverage tools and utilize them."
    d.Range(0, 0).Select
    s = Replace(SKW_BannedPhrasesReview(), vbCr, " / ")
    s = s & " ## text=" & Trim$(Replace(d.Content.Text, vbCr, ""))
    If d.Comments.Count > 0 Then s = s & " ## comment=" & d.Comments(1).Range.Text
    d.Range(0, 0).Select
    SKW_SectionMap
    s = s & " ## report=" & Left$(ReportText(), 300)
    CloseReport
    d.Close SaveChanges:=0
    TW_EffortBanned = s
End Function

Public Function TW_RenameFolder(ByVal arg As String) As String
    Dim a As Variant
    a = Split(arg, "|")
    Name a(0) As a(1)
    SKW_ForgetEffort
    TW_RenameFolder = "renamed"
End Function

' ---------------------------------------------------------------------
' 1.4.0: PowerPoint batch 2. A deck with a pasted-in foreign master, a
' "1_" copy layout, a banned phrase, a font not allowed, a typed slide
' number, a slide covering the logo and the number, and pictures.
' arg = "<deck path>|<template path>"
' ---------------------------------------------------------------------
Public Function TW_MakeDeck2(ByVal arg As String) As String
    Dim a As Variant, t As Object, p As Object, s As Object, d2 As Object, lay As Object, sh As Object, num As Object, w As Single, h As Single
    Dim extra As String
    a = Split(arg, "|")
    Set mPP = CreateObject("PowerPoint.Application")
    mPPMine = (mPP.Presentations.Count = 0)
    ' The team template: a plain deck saved as .potx.
    Set t = mPP.Presentations.Add(0)
    t.Slides.Add 1, 1
    t.SaveAs CStr(a(1)), 26
    t.Close
    Set p = mPP.Presentations.Add(0)
    w = p.PageSetup.SlideWidth: h = p.PageSetup.SlideHeight
    ' 1: title slide
    Set s = p.Slides.Add(1, 1)
    s.Shapes(1).TextFrame.TextRange.Text = "Proposal kickoff"
    ' 2: title and content: banned phrase twice, a font not allowed, a typed number, two pictures
    Set s = p.Slides.AddSlide(2, p.Designs(1).SlideMaster.CustomLayouts(2))
    s.Shapes(1).TextFrame.TextRange.Text = "Approach"
    s.Shapes(2).TextFrame.TextRange.Text = "We utilize proven tools. Utilize them well."
    s.Shapes(2).TextFrame.TextRange.Paragraphs(1).Font.Name = "Comic Sans MS"
    Set sh = s.Shapes.AddTextbox(1, w - 60, h - 40, 40, 24)
    sh.TextFrame.TextRange.Text = "2"
    Set sh = s.Shapes.AddPicture(TinyBmp(), 0, -1, 40, 300, 60, 60)
    sh.AlternativeText = "Rocket on the pad"
    s.Shapes.AddPicture TinyBmp(), 0, -1, 120, 300, 60, 60
    ' 3: a slide pasted in with its own master (Keep Source Formatting)
    Set d2 = p.Designs.Add("Foreign Theme")
    Set s = p.Slides.AddSlide(3, d2.SlideMaster.CustomLayouts(2))
    s.Shapes(1).TextFrame.TextRange.Text = "Foreign slide"
    ' 4: a slide on a "1_" copy of a template layout
    Set lay = p.Designs(1).SlideMaster.CustomLayouts(2).Duplicate
    Set s = p.Slides.AddSlide(4, lay)
    s.Shapes(1).TextFrame.TextRange.Text = "Copied layout slide"
    ' 5: title only, with a box over the master's logo and the slide number place
    Set sh = p.Designs(1).SlideMaster.Shapes.AddShape(1, 10, h - 50, 80, 40)
    sh.Name = "Team Logo"
    Set s = p.Slides.Add(5, 11)
    s.Shapes(1).TextFrame.TextRange.Text = "Covered"
    Set num = Nothing
    For Each sh In s.CustomLayout.Shapes
        If sh.Type = 14 Then
            If sh.PlaceholderFormat.Type = 13 Then Set num = sh
        End If
    Next
    Set sh = s.Shapes.AddShape(1, 0, h - 60, w, 60)
    sh.Name = "Wide band"
    ' A linked picture (kept in the file too).
    extra = "none"
    On Error Resume Next
    s.Shapes.AddPicture TinyBmp(), -1, -1, 600, 200, 60, 60
    If Err.Number = 0 Then extra = "Linked picture"
    On Error GoTo 0
    p.SaveAs CStr(a(0))
    Set SKP_TestDeck = p
    SKP_TestTemplate = CStr(a(1))
    TW_MakeDeck2 = "slides=" & p.Slides.Count & " designs=" & p.Designs.Count & " numberOnLayout=" & (Not num Is Nothing) & " media5=" & extra
End Function

' Runs a fixer with a target path. arg = "<function>|<target>"
Public Function TW_Fix(ByVal arg As String) As String
    Dim a As Variant, s As String
    a = Split(arg, "|")
    SKW_Quiet = True
    s = Replace(Application.Run(CStr(a(0)), CStr(a(1))), vbCr, " / ") & " ## " & ReportText()
    CloseReport
    TW_Fix = s
End Function

' What a saved copy holds: masters, layouts in use, each slide's layout, number, comments; fonts and text on slide 2.
Public Function TW_Inspect(ByVal path As String) As String
    Dim c As Object
    Set c = mPP.Presentations.Open(path, -1, 0, 0)
    TW_Inspect = InspectDeck(c)
    c.Close
End Function

' The same, for the open test deck (left open).
Public Function TW_InspectOpen() As String
    TW_InspectOpen = InspectDeck(SKP_TestDeck)
End Function

Private Function InspectDeck(ByVal c As Object) As String
    Dim s As String, sld As Object, sh As Object, f As Object, rn As Object, hasNum As Boolean, nMark As Long, d As Object, lay As Object, n As Long
    s = "designs=" & c.Designs.Count & " layouts:"
    For Each d In c.Designs
        For Each lay In d.SlideMaster.CustomLayouts
            n = n + 1
            If Left$(lay.Name, 2) = "1_" Then s = s & " [" & lay.Name & "]"
        Next
    Next
    s = s & " total=" & n & " | slides:"
    Set f = CreateObject("Scripting.Dictionary")
    For Each sld In c.Slides
        hasNum = False
        For Each sh In sld.Shapes
            If sh.Type = 14 Then
                If sh.PlaceholderFormat.Type = 13 Then hasNum = True
            End If
            If Left$(sh.Name, 11) = "SK overlap " Then nMark = nMark + 1
        Next
        s = s & " " & sld.SlideIndex & ":" & sld.CustomLayout.Name & ":num=" & IIf(hasNum, 1, 0) & ":com=" & sld.Comments.Count
        If sld.SlideIndex = 2 Then
            For Each rn In sld.Shapes(2).TextFrame.TextRange.Runs
                f(rn.Font.Name) = True
            Next
        End If
    Next
    s = s & " | marks=" & nMark & " | fonts2=" & Join(f.Keys, ",") & " | text2=" & Replace(c.Slides(2).Shapes(2).TextFrame.TextRange.Text, vbCr, " / ")
    s = s & " | typed2=" & TypedTwo(c.Slides(2))
    If c.Slides(2).Comments.Count > 0 Then s = s & " | comment2=" & c.Slides(2).Comments(1).Text
    InspectDeck = s
End Function

Private Function TypedTwo(ByVal sld As Object) As Long
    Dim sh As Object
    For Each sh In sld.Shapes
        If sh.Type = 17 Then
            If Trim$(sh.TextFrame.TextRange.Text) = "2" Then TypedTwo = TypedTwo + 1
        End If
    Next
End Function

' The media list form on the test deck: the list, a rename, naming from alt text, and the deck's names after.
Public Function TW_Media() As String
    Dim s As String, sh As Object
    SKW_Quiet = True
    Load frmSKPMedia
    s = "list: " & frmSKPMedia.TestList("All media")
    s = s & " || noalt: " & frmSKPMedia.TestList("No alt text")
    s = s & " || nameall: " & frmSKPMedia.TestNameAll()
    s = s & " || rename: " & frmSKPMedia.TestRename(1, "Second photo")
    s = s & " || dup: " & frmSKPMedia.TestRename(0, "Second photo")
    Unload frmSKPMedia
    s = s & " || deck2:"
    For Each sh In SKP_TestDeck.Slides(2).Shapes
        s = s & " [" & sh.Name & "]"
    Next
    TW_Media = s
End Function

' File names without and with their extension.
Public Function TW_Names() As String
    TW_Names = SKW_Stem("Deck 1.4.pptx") & "|" & SKW_Ext("Book.xlsx") & "|" & SKW_Stem("C:\a.b\Deck") & "|" & SKW_Stem(".hidden") & "|" & SKW_Ext("Deck")
End Function

' ---------------------------------------------------------------------
' 1.5.0: PowerPoint batch 3
' ---------------------------------------------------------------------
' A 24-bit bitmap w x h pixels at 72 dpi (one point a pixel), in TEMP.
Private Function BigBmp(ByVal w As Long, ByVal h As Long) As String
    Dim p As String, f As Integer, b() As Byte, rowLen As Long, size As Long, i As Long
    p = Environ$("TEMP") & "\skw_test_" & w & "x" & h & ".bmp"
    rowLen = ((w * 3 + 3) \ 4) * 4
    size = 54 + rowLen * h
    ReDim b(0 To size - 1)
    b(0) = 66: b(1) = 77
    PutLong b, 2, size: PutLong b, 10, 54: PutLong b, 14, 40: PutLong b, 18, w: PutLong b, 22, h
    b(26) = 1: b(28) = 24
    PutLong b, 34, rowLen * h: PutLong b, 38, 2835: PutLong b, 42, 2835
    For i = 54 To size - 1
        b(i) = (i * 7) Mod 256
    Next
    f = FreeFile
    Open p For Binary Access Write As #f
    Put #f, 1, b
    Close #f
    BigBmp = p
End Function

Private Sub PutLong(ByRef b() As Byte, ByVal at As Long, ByVal v As Long)
    b(at) = v And &HFF&: b(at + 1) = (v \ &H100&) And &HFF&: b(at + 2) = (v \ &H10000) And &HFF&: b(at + 3) = (v \ &H1000000) And &HFF&
End Sub

' Sets a team setting in the standards read into memory (not the workbook). Returns what SKW_Setting now reads.
Public Function TW_SetSetting(ByVal arg As String) As String
    Dim a As Variant, r As Object, rs As Collection
    a = Split(arg, "|")
    Set rs = SKW_Rows("Settings")
    For Each r In rs
        If LCase$(r("setting")) = LCase$(CStr(a(0))) Then
            r("value") = CStr(a(1))
            TW_SetSetting = SKW_Setting(CStr(a(0)), "?")
            Exit Function
        End If
    Next
    Set r = CreateObject("Scripting.Dictionary")
    r("setting") = CStr(a(0)): r("value") = CStr(a(1))
    rs.Add r
    TW_SetSetting = SKW_Setting(CStr(a(0)), "?")
End Function

' Deck 1.5: an acronym in a title and in the body, a moved and shrunk title, a French paragraph, an off-theme
' color, an overflowing text box, a wrong footer, a stretched and a cropped picture, a local hyperlink and a
' web one, a table in another style with no header row, and a CUI marking on slide 4 only.
Public Function TW_MakeDeck3(ByVal path As String) As String
    Dim p As Object, s As Object, sh As Object, tb As Object, bmp As String, w As Single, h As Single
    Set mPP = CreateObject("PowerPoint.Application")
    mPPMine = (mPP.Presentations.Count = 0)
    Set p = mPP.Presentations.Add(0)
    w = p.PageSetup.SlideWidth: h = p.PageSetup.SlideHeight
    ' 1: title slide; RFP only in the title
    Set s = p.Slides.AddSlide(1, p.Designs(1).SlideMaster.CustomLayouts(1))
    s.Shapes(1).TextFrame.TextRange.Text = "Approach for the RFP"
    ' 2: title moved and shrunk; body with RFP, a French paragraph and an off-theme color; overflow; wrong footer
    Set s = p.Slides.AddSlide(2, p.Designs(1).SlideMaster.CustomLayouts(2))
    s.Shapes(1).TextFrame.TextRange.Text = "Overview"
    s.Shapes(1).Top = s.Shapes(1).Top + 30
    s.Shapes(1).TextFrame.TextRange.Font.Size = 20
    s.Shapes(2).TextFrame.TextRange.Text = "We answer the RFP in full." & vbCr & "Bonjour tout le monde."
    s.Shapes(2).TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = RGB(&H12, &H34, &H56)
    s.Shapes(2).TextFrame.TextRange.Paragraphs(2).LanguageID = 1036
    Set tb = s.Shapes.AddTextbox(1, 40, 420, 220, 30)
    tb.Name = "Tight box"
    tb.TextFrame.AutoSize = 0
    tb.TextFrame.WordWrap = -1
    tb.TextFrame.TextRange.Text = "This text is far too long for its small box." & vbCr & "A second line." & vbCr & "A third line." & vbCr & "A fourth line."
    tb.Height = 30
    s.HeadersFooters.Footer.Visible = -1
    s.HeadersFooters.Footer.Text = "Draft"
    ' 3: pictures, links, tables
    Set s = p.Slides.AddSlide(3, p.Designs(1).SlideMaster.CustomLayouts(6))
    s.Shapes(1).TextFrame.TextRange.Text = "Evidence"
    bmp = BigBmp(100, 50)
    Set sh = s.Shapes.AddPicture(bmp, 0, -1, 40, 150)
    sh.Name = "Stretched photo"
    sh.LockAspectRatio = 0
    sh.Width = 200: sh.Height = 50
    Set sh = s.Shapes.AddPicture(bmp, 0, -1, 300, 150)
    sh.Name = "Cropped photo"
    sh.PictureFormat.CropRight = 20
    Set tb = s.Shapes.AddTextbox(1, 40, 260, 300, 30)
    tb.Name = "Local link"
    tb.TextFrame.TextRange.Text = "See the plan"
    tb.TextFrame.TextRange.ActionSettings(1).Hyperlink.Address = "C:\Users\nobody\Secret plan.docx"
    Set tb = s.Shapes.AddTextbox(1, 40, 300, 300, 30)
    tb.Name = "Web link"
    tb.TextFrame.TextRange.Text = "Our site"
    tb.TextFrame.TextRange.ActionSettings(1).Hyperlink.Address = "https://example.com/"
    Set sh = s.Shapes.AddTable(2, 2, 450, 150, 200, 60)
    sh.Name = "Team table"
    Set sh = s.Shapes.AddTable(2, 2, 450, 260, 200, 60)
    sh.Name = "Odd table"
    sh.Table.ApplyStyle "{2D5ABB26-0587-4C30-8999-92F81FD0307C}", False
    sh.Table.FirstRow = False
    ' 4: the marking
    Set s = p.Slides.AddSlide(4, p.Designs(1).SlideMaster.CustomLayouts(2))
    s.Shapes(1).TextFrame.TextRange.Text = "Close"
    s.Shapes(2).TextFrame.TextRange.Text = "Thank you."
    Set tb = s.Shapes.AddTextbox(1, 10, 5, 100, 20)
    tb.TextFrame.TextRange.Text = "CUI"
    p.SaveAs path
    Set SKP_TestDeck = p
    SKP_TestTemplate = ""
    TW_MakeDeck3 = "slides=" & p.Slides.Count & " shapes3=" & p.Slides(3).Shapes.Count & " bmp=" & Len(Dir$(bmp))
End Function

' What a 1.5.0 fixed copy holds.
Public Function TW_Inspect3(ByVal path As String) As String
    Dim c As Object, s As String, sld As Object, sh As Object, lt As Object, ft As String, ph As Object
    Set c = mPP.Presentations.Open(path, -1, 0, 0)
    Set sld = c.Slides(2)
    s = "body2=" & Replace(sld.Shapes(2).TextFrame.TextRange.Text, vbCr, " / ")
    s = s & " | lang2=" & sld.Shapes(2).TextFrame.TextRange.Paragraphs(2).LanguageID
    For Each lt In sld.CustomLayout.Shapes
        If lt.Type = 14 Then
            If lt.PlaceholderFormat.Type = 1 Then s = s & " | titleDy=" & Format$(sld.Shapes(1).Top - lt.Top, "0") & " titleSize=" & sld.Shapes(1).TextFrame.TextRange.Characters(1, 1).Font.Size
        End If
    Next
    s = s & " | com2=" & sld.Comments.Count
    s = s & " | footers="
    For Each sld In c.Slides
        ft = "-"
        For Each ph In sld.Shapes
            If ph.Type = 14 Then
                If ph.PlaceholderFormat.Type = 15 Then ft = ph.TextFrame.TextRange.Text
            End If
        Next
        s = s & "[" & ft & "]"
    Next
    Set sld = c.Slides(3)
    s = s & " | pic=" & Format$(sld.Shapes("Stretched photo").Width, "0") & "x" & Format$(sld.Shapes("Stretched photo").Height, "0")
    s = s & " | tables=" & sld.Shapes("Team table").Table.Style.Id & "," & sld.Shapes("Odd table").Table.Style.Id
    On Error Resume Next
    s = s & " | deflang=" & c.DefaultLanguageID
    c.Close
    TW_Inspect3 = s
End Function

' 1.6.0: the Stationkeeper look on the Deck Media List (on the open test deck).
Public Function TW_MediaBrand() As String
    SKW_Quiet = True
    Load frmSKPMedia
    TW_MediaBrand = frmSKPMedia.TestBrand()
    Unload frmSKPMedia
End Function

' ---------------------------------------------------------------------
' 1.7.0: sets, one tool for a check and its fix
' ---------------------------------------------------------------------
' Runs a tool the way the menu does (its macro, not its function) and returns the findings and the report.
Public Function TW_Entry(ByVal macro As String) As String
    Dim s As String
    SKW_Quiet = True
    Application.Run macro
    s = "rows=" & SKW_LastRows & " ## " & ReportText()
    CloseReport
    TW_Entry = s
End Function

' Runs a check and returns what its fix could put right.
Public Function TW_Fixable(ByVal fn As String) As String
    SKW_Quiet = True
    SKW_Fixable = -1
    Application.Run fn
    CloseReport
    TW_Fixable = fn & "=" & SKW_Fixable
End Function

' The offer after a check is never made while tests run, and not when nothing can be fixed.
Public Function TW_Offer() As String
    SKW_Quiet = True
    SKW_Fixable = 3
    TW_Offer = "quiet=" & SKW_OfferFix("3 found.", "The fix does it.")
End Function

' ---------------------------------------------------------------------
' 1.8.0: table roles
' ---------------------------------------------------------------------
Private Sub FillRow(ByVal t As Object, ByVal r As Long, ByVal a As String, ByVal b As String, ByVal c As String, ByVal e As String)
    t.Cell(r, 1).Range.Text = a
    t.Cell(r, 2).Range.Text = b
    t.Cell(r, 3).Range.Text = c
    t.Cell(r, 4).Range.Text = e
End Sub

' A saved document with three fictional tables: one with a caption row, dividers and a total row; one marked skip; one plain.
Public Function TW_BuildRoles(ByVal savePath As String) As String
    Dim d As Object, t As Object, rg As Object
    SKW_Quiet = True
    ' Based on this template, so its macros stay in reach while the new document is the active one.
    Set d = Documents.Add(Template:=ThisDocument.FullName)
    d.TrackRevisions = False
    d.Content.Delete
    d.ActiveWindow.View.ShowFieldCodes = False
    d.Content.InsertAfter "Before the first table." & vbCr
    Set t = d.Tables.Add(AtEnd(d), 8, 4)
    FillRow t, 2, "Item", "Description", "FY26", "FY27"
    FillRow t, 4, "Engineer", "Designs the ground segment and writes the interface documents for the station", "$1,200", "(300)"
    FillRow t, 5, "Tech", "Two words", "45%", "-"
    FillRow t, 7, "Antenna", "Long text of more than three words here", "2,500.50", "N/A"
    FillRow t, 8, "Total", "", "$3,745", "1,000"
    t.Rows(6).Cells.Merge
    t.Cell(6, 1).Range.Text = "Materials"
    t.Rows(3).Cells.Merge
    t.Cell(3, 1).Range.Text = "Labor"
    t.Rows(1).Cells.Merge
    t.Cell(1, 1).Range.Text = "Table "
    Set rg = t.Cell(1, 1).Range
    rg.MoveEnd wdCharacter, -1
    rg.Collapse 0
    d.Fields.Add rg, wdFieldEmpty, "SEQ Table \* ARABIC", False
    Set rg = t.Cell(1, 1).Range
    rg.MoveEnd wdCharacter, -1
    rg.InsertAfter ". Costs by year"
    d.Content.InsertParagraphAfter
    d.Content.InsertAfter "Between." & vbCr
    Set t = d.Tables.Add(AtEnd(d), 2, 2)
    t.Cell(1, 1).Range.Text = "Leave me"
    t.Cell(2, 2).Range.Text = "500"
    t.Descr = "skip"
    d.Content.InsertParagraphAfter
    d.Content.InsertAfter "Between again." & vbCr
    Set t = d.Tables.Add(AtEnd(d), 3, 2)
    t.Cell(1, 1).Range.Text = "Name": t.Cell(1, 2).Range.Text = "Count"
    t.Cell(2, 1).Range.Text = "Alpha": t.Cell(2, 2).Range.Text = "12"
    t.Cell(3, 1).Range.Text = "Subtotal": t.Cell(3, 2).Range.Text = "12"
    d.Content.InsertParagraphAfter
    d.SaveAs2 FileName:=savePath, FileFormat:=wdFormatXMLDocument
    d.Range(0, 0).Select
    TW_BuildRoles = "saved=" & d.Name & " tables=" & d.Tables.Count & " fields=" & d.Tables(1).Range.Fields.Count & " cap=" & _
        Replace(Replace(d.Tables(1).Cell(1, 1).Range.Text, Chr(13) & Chr(7), ""), vbCr, "")
End Function

Private Function RolesDoc() As Object
    Dim d As Object
    For Each d In Documents
        If d.Name = "Roles 1.8.docx" Then Set RolesDoc = d
    Next
    RolesDoc.Activate
End Function

Public Function TW_RolesCheck() As String
    Dim s As String
    SKW_Quiet = True
    RolesDoc
    SKW_Fixable = -1
    s = SKW_TableRolesCheck() & " ## fixable=" & SKW_Fixable & " ## " & ReportText()
    CloseReport
    TW_RolesCheck = s
End Function

Private Function CellState(ByVal t As Object, ByVal r As Long, ByVal c As Long) As String
    With t.Cell(r, c)
        CellState = r & "." & c & "=" & .Range.ParagraphFormat.Alignment & "/" & .VerticalAlignment & "/" & IIf(.Range.Font.Bold = True, "B", "n")
    End With
End Function

Private Function RolesState(ByVal d As Object) As String
    Dim t As Object, s As String, h As String
    Set t = d.Tables(1)
    On Error Resume Next
    h = "?"
    h = IIf(d.Range(t.Cell(1, 1).Range.Start, t.Cell(2, 4).Range.End).Rows.HeadingFormat = True, "Y", "N")
    On Error GoTo 0
    s = CellState(t, 2, 1) & " " & CellState(t, 3, 1) & " " & CellState(t, 4, 1) & " " & CellState(t, 4, 2) & " " & CellState(t, 4, 3) & " " & _
        CellState(t, 5, 4) & " " & CellState(t, 7, 4) & " " & CellState(t, 8, 1) & " " & CellState(t, 8, 3) & " head=" & h & _
        " fields=" & t.Range.Fields.Count & " cap=" & Replace(Replace(t.Cell(1, 1).Range.Text, Chr(13) & Chr(7), ""), vbCr, "")
    s = s & " ## skip " & CellState(d.Tables(2), 2, 2) & " ## plain " & CellState(d.Tables(3), 1, 1) & " " & CellState(d.Tables(3), 2, 2) & " " & _
        CellState(d.Tables(3), 3, 1) & " " & CellState(d.Tables(3), 3, 2)
    RolesState = s
End Function

' Styles a copy, then reads the copy and the document it came from.
Public Function TW_RolesFix(ByVal target As String) As String
    Dim src As Object, s As String, c As Object
    SKW_Quiet = True
    Set src = RolesDoc()
    s = Replace(SKW_TableRolesFixCopy(target), vbCr, " / ")
    Set c = ActiveDocument
    s = s & " ## same=" & (c.FullName = src.FullName) & " ## COPY " & RolesState(c) & " ## SOURCE " & RolesState(src)
    If c.FullName <> src.FullName Then c.Close SaveChanges:=0
    src.Activate
    TW_RolesFix = s
End Function

' Styles the plain table where it is.
Public Function TW_RolesHere() As String
    Dim d As Object
    SKW_Quiet = True
    Set d = RolesDoc()
    d.Tables(3).Cell(2, 1).Range.Select
    Selection.Collapse
    TW_RolesHere = Replace(SKW_ApplyTableRolesHere(), vbCr, " / ")
End Function

Public Function TW_RolesHereState() As String
    Dim d As Object
    Set d = RolesDoc()
    TW_RolesHereState = CellState(d.Tables(3), 1, 1) & " " & CellState(d.Tables(3), 2, 2) & " " & CellState(d.Tables(3), 3, 1) & " " & CellState(d.Tables(3), 3, 2)
End Function

Public Function TW_RolesClose() As String
    RolesDoc().Close SaveChanges:=0
    TW_RolesClose = "closed"
End Function

' The roles as the tools read them now: one line per role.
Public Function TW_RolesRead() As String
    Dim r As Object, k As Variant, s As String
    Set r = SKW_TableRoles("Standard")
    For Each k In r.Keys
        s = s & k & "=" & r(k)("paragraph style") & "," & r(k)("align") & "," & r(k)("vertical") & "," & r(k)("bold") & "," & r(k)("recognize") & " | "
    Next
    TW_RolesRead = "sheet rows=" & SKW_Rows("Table roles").Count & " ## " & s
End Function

' The tab stops a person set (not Word's default ones), as positions in points.
Private Function CustomTabs(ByVal ts As Object) As String
    Dim t As Object, s As String, n As Long
    For Each t In ts
        If t.CustomTab Then n = n + 1: s = s & "," & Format$(t.Position, "0")
    Next
    CustomTabs = n & "@" & Mid$(s, 2)
End Function

' Makes a standards workbook, sets a header style and fill, a long text style, a kept style, borders and cell margins in it,
' then styles a copy of the roles document from it.
Public Function TW_RolesWorkbook(ByVal path As String) As String
    Dim prev As String, xl As Object, wb As Object, ws As Object, r As Long, d As Object, c As Object, t As Object, s As String, rg As Object
    SKW_Quiet = True
    prev = SKW_StandardsPath()
    s = Left$(SKW_MakeStandardsWorkbook(path, False), 4)
    Set xl = CreateObject("Excel.Application")
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(path)
    Set ws = wb.Worksheets("Table types")
    ws.Cells(2, 11).Value = "0.05": ws.Cells(2, 12).Value = "Single": ws.Cells(2, 13).Value = "1": ws.Cells(2, 14).Value = "FF0000"
    Set ws = wb.Worksheets("Table roles")
    For r = 2 To 14
        Select Case LCase$(ws.Cells(r, 2).Value)
            Case "header": ws.Cells(r, 3).Value = "Heading 2": ws.Cells(r, 6).Value = "D9E2F3"
            Case "long text": ws.Cells(r, 3).Value = "Body Text"
            Case "caption": ws.Cells(r, 9).Value = "0.1": ws.Cells(r, 10).Value = "1"
            Case "figure caption": ws.Cells(r, 8).Value = "Caption": ws.Cells(r, 4).Value = "Center": ws.Cells(r, 10).Value = "0.5; 1.25"
            Case "keep": ws.Cells(r, 8).Value = "List Bullet"
        End Select
    Next
    wb.Save
    wb.Close False
    xl.Quit
    SKW_SetStandardsPath path
    Set d = RolesDoc()
    Set rg = d.Tables(1).Cell(4, 2).Range
    rg.MoveEnd wdCharacter, -1
    rg.InsertAfter vbCr & "A bullet inside the cell"
    d.Tables(1).Cell(4, 2).Range.Paragraphs(2).Style = d.Styles("List Bullet")
    d.Tables(1).Rows.Add
    d.Tables(1).Rows(9).Cells.Merge
    d.Tables(1).Cell(9, 1).Range.Text = "Note: fictional figures."
    d.Paragraphs(d.Paragraphs.Count).Range.InsertBefore "Figure 1" & vbTab & "A caption outside a table"
    d.Paragraphs(d.Paragraphs.Count).Style = d.Styles("Caption")
    d.Save
    s = s & " ## sheet rows=" & SKW_Rows("Table roles").Count & " ## " & Replace(SKW_TableRolesFixCopy(""), vbCr, " / ")
    Set c = ActiveDocument
    Set t = c.Tables(1)
    s = s & " ## same=" & (c.FullName = d.FullName) & " head=" & t.Cell(2, 1).Range.Paragraphs(1).Style.NameLocal & _
        " fill=" & Hex$(t.Cell(2, 1).Shading.BackgroundPatternColor) & " p1=" & t.Cell(4, 2).Range.Paragraphs(1).Style.NameLocal & _
        " p2=" & t.Cell(4, 2).Range.Paragraphs(2).Style.NameLocal & " border=" & t.Borders(wdBorderTop).LineStyle & "/" & _
        t.Borders(wdBorderTop).LineWidth & "/" & Hex$(t.Borders(wdBorderHorizontal).Color) & " pad=" & Format$(t.TopPadding, "0.0") & "/" & Format$(t.LeftPadding, "0.0") & _
        " " & CellState(t, 9, 1)
    With t.Cell(1, 1).Range.ParagraphFormat
        s = s & " cap indent=" & Format$(.LeftIndent, "0.0") & " tabs=" & CustomTabs(.TabStops)
    End With
    With c.Paragraphs(c.Paragraphs.Count).Range.ParagraphFormat
        s = s & " ## fig align=" & .Alignment & " tabs=" & CustomTabs(.TabStops)
    End With
    If c.FullName <> d.FullName Then c.Close SaveChanges:=0
    SKW_SetStandardsPath prev
    d.Activate
    TW_RolesWorkbook = s
End Function

' ---------------------------------------------------------------------
' 1.8.1: the progress window
' ---------------------------------------------------------------------
Private Sub WaitSecs(ByVal secs As Single)
    Dim t As Single
    t = Timer
    Do While Timer - t < secs And Timer >= t
    Loop
End Sub

' Runs a counted loop the way a tool does and reads the window at three points.
Public Function TW_Progress() As String
    Dim i As Long, s As String, msg As String
    SKW_Quiet = True
    SKW_ProgressForce = True
    RolesDoc
    msg = SKW_Begin("progress test", False)
    SKW_Steps 200, "Paragraph"
    For i = 1 To 99
        SKW_Step
    Next
    s = "first: " & SKW_ProgressProbe()
    WaitSecs 0.4
    SKW_Step
    s = s & " ## half: " & SKW_ProgressProbe()
    WaitSecs 0.4
    SKW_Status "Reading table 3 of 4..."
    s = s & " ## status: " & SKW_ProgressProbe() & " ## line: " & SKW_ProgressLine
    SKW_Status "Saving"
    s = s & " ## phase: " & SKW_ProgressProbe()
    SKW_Finish
    s = s & " ## after: " & SKW_ProgressProbe()
    SKW_ProgressForce = False
    TW_Progress = s
End Function

' A set shows how far the whole set is, not only the check running.
Public Function TW_ProgressSet() As String
    Dim s As String
    SKW_Quiet = True
    RolesDoc
    s = SKW_RunSet("Mini set", "", Array("SKW_TableRolesCheck|Tables one", "SKW_TableRolesCheck|Tables two"), False)
    CloseReport
    TW_ProgressSet = s & " ## " & SKW_ProgressLine & " ## " & SKW_ProgressProbe()
End Function

' ---------------------------------------------------------------------
' 1.8.1: the reports folder
' ---------------------------------------------------------------------
' Sets the folder the way the menu and admin do (the kit's setting), runs a report, then clears it and runs another.
Public Function TW_ReportsFolder(ByVal folder As String) As String
    Dim s As String, nm As String, pth As String
    SKW_Quiet = True
    RolesDoc
    SK_SetReportsDir folder
    s = "tools read=" & SKW_ReportsFolder()
    SKW_TableRolesCheck
    nm = SKW_LastReport.Name: pth = SKW_LastReport.Path
    s = s & " ## saved=" & nm & " in=" & pth & " exists=" & (Len(Dir$(pth & "\" & nm)) > 0) & " clean=" & SKW_LastReport.Saved
    CloseReport
    SKW_RunSet "Mini set", "", Array("SKW_TableRolesCheck|Tables one"), False
    s = s & " ## set=" & SKW_LastReport.Name & " in=" & SKW_LastReport.Path
    CloseReport
    SK_SetReportsDir ""
    SKW_TableRolesCheck
    s = s & " ## cleared: folder=[" & SKW_ReportsFolder() & "] path=[" & SKW_LastReport.Path & "]"
    CloseReport
    TW_ReportsFolder = s
End Function

' ---------------------------------------------------------------------
' 1.8.1: Delete Unused Styles
' ---------------------------------------------------------------------
Private Function HasStyle(ByVal d As Object, ByVal nm As String) As String
    Dim st As Object
    On Error Resume Next
    Set st = d.Styles(nm)
    If st Is Nothing Then HasStyle = nm & "=gone" Else HasStyle = nm & "=" & IIf(st.Visibility, "hidden", "shown")
End Function

' On the roles document: two unused custom styles (one based on the other), one custom style in use,
' an unused built-in style that shows in the pane, and a built-in style in use.
Public Function TW_DeleteUnused() As String
    Dim d As Object, s As String, a As Object
    SKW_Quiet = True
    Set d = RolesDoc()
    Set a = d.Styles.Add("ZZ Unused A", wdStyleTypeParagraph)
    d.Styles.Add "ZZ Unused B", wdStyleTypeParagraph
    d.Styles("ZZ Unused B").BaseStyle = "ZZ Unused A"
    d.Styles.Add "ZZ In Use", wdStyleTypeParagraph
    d.Paragraphs(1).Style = d.Styles("ZZ In Use")
    d.Styles("Intense Quote").Visibility = False
    d.Styles("Title").Visibility = False
    d.Paragraphs(d.Paragraphs.Count).Range.InsertBefore "A title line"
    d.Paragraphs(d.Paragraphs.Count).Style = d.Styles("Title")
    s = SKW_DeleteUnusedStyles("") & " ## fixable>0=" & (SKW_Fixable > 0) & " ## " & ReportText()
    CloseReport
    s = s & " ## before: " & HasStyle(d, "ZZ Unused A") & " " & HasStyle(d, "ZZ Unused B")
    d.Activate
    s = s & " ## " & Replace(SKW_DeleteUnusedStyles("fix"), vbCr, " / ")
    s = s & " ## after: " & HasStyle(d, "ZZ Unused A") & " " & HasStyle(d, "ZZ Unused B") & " " & HasStyle(d, "ZZ In Use") & " " & _
        HasStyle(d, "Intense Quote") & " " & HasStyle(d, "Title") & " " & HasStyle(d, "Normal") & " " & HasStyle(d, "Heading 1")
    TW_DeleteUnused = s
End Function

' ---------------------------------------------------------------------
' 1.9.0: the Stationkeeper app's standards workbook (Company Standards.xlsx, format 2)
' ---------------------------------------------------------------------
Private Function EdgeState(ByVal c As Object, ByVal which As Long) As String
    On Error Resume Next
    EdgeState = "?"
    With c.Borders(which)
        EdgeState = .LineStyle & "/" & .LineWidth & "/" & Hex$(.Color)
    End With
End Function

Private Function ParaStyleOf(ByVal c As Object) As String
    On Error Resume Next
    ParaStyleOf = c.Range.Paragraphs(1).Style.NameLocal
End Function

' Points the tools at the app's own workbook and styles a copy of the roles document from its Table looks and Table types.
Public Function TW_AppStandards(ByVal path As String) As String
    Dim prev As String, d As Object, c As Object, t As Object, s As String, st As Object, t3 As Object
    SKW_Quiet = True
    prev = SKW_StandardsPath()
    SKW_SetStandardsPath path
    Set d = RolesDoc()
    s = "app=" & SKW_AppStandards() & " types=" & SKW_TableTypes().Count & " first=" & SKW_TableTypes()(1)("type") & _
        " font=" & SKW_Setting("table_font", "?") & " changed=" & (Len(SKW_StdChanged()) > 0)
    s = s & " ## " & SKW_TableRolesCheck() & " ## " & ReportText()
    CloseReport
    d.Activate
    d.Save
    s = s & " ## " & Replace(SKW_TableRolesFixCopy(""), vbCr, " / ")
    Set c = ActiveDocument
    Set t = c.Tables(1)
    Set t3 = c.Tables(3)
    s = s & " ## same=" & (c.FullName = d.FullName)
    s = s & " cap=" & ParaStyleOf(t.Cell(1, 1)) & "|" & Replace(Replace(t.Cell(1, 1).Range.Text, Chr(13) & Chr(7), ""), vbCr, "") & "|fields=" & t.Cell(1, 1).Range.Fields.Count
    s = s & " head=" & ParaStyleOf(t.Cell(2, 1)) & "|" & Hex$(t.Cell(2, 1).Shading.BackgroundPatternColor) & "|" & Hex$(t.Cell(2, 1).Range.Font.Color)
    s = s & " div=" & ParaStyleOf(t.Cell(3, 1)) & "|" & Hex$(t.Cell(3, 1).Shading.BackgroundPatternColor)
    s = s & " body=" & ParaStyleOf(t.Cell(5, 1)) & "|" & t.Cell(5, 1).Range.Font.Name & "|" & t.Cell(5, 1).Range.Font.Size & "|" & t.Cell(5, 1).VerticalAlignment
    s = s & " total=" & ParaStyleOf(t.Cell(8, 1)) & "|" & Hex$(t.Cell(8, 1).Shading.BackgroundPatternColor) & "|top=" & EdgeState(t.Cell(8, 1), wdBorderTop)
    s = s & " note=" & ParaStyleOf(t.Cell(9, 1))
    s = s & " sub=" & ParaStyleOf(t3.Cell(3, 1)) & " pad=" & Format$(t.TopPadding, "0") & "/" & Format$(t.LeftPadding, "0") & " tblalign=" & t.Rows.Alignment
    s = s & " bodyedge=" & EdgeState(t.Cell(5, 2), wdBorderLeft) & " rep2=" & t.Rows(2).HeadingFormat
    Set st = Nothing
    On Error Resume Next
    Set st = c.Styles("PE Table Header")
    On Error GoTo 0
    If Not st Is Nothing Then s = s & " style=" & st.Font.Name & "|" & st.Font.Size & "|" & st.Font.Bold & "|" & st.ParagraphFormat.SpaceBefore
    s = s & " ## source head=" & ParaStyleOf(d.Tables(1).Cell(2, 1))
    If c.FullName <> d.FullName Then c.Close SaveChanges:=0
    SKW_SetStandardsPath prev
    d.Activate
    TW_AppStandards = s
End Function

' Assigning: the plain table of the roles document, read from the app's standards.
' Suggest and mark; set a row by hand; fix again (the hand-set row stays); change a word a guess would
' use (the role stays); add a row with no role (the check names it); suggest again (the guesses return).
Public Function TW_Assign(ByVal path As String) As String
    Dim prev As String, d As Object, t As Object, s As String, w0 As Single
    SKW_Quiet = True
    prev = SKW_StandardsPath()
    SKW_SetStandardsPath path
    Set d = RolesDoc()
    Set t = d.Tables(3)
    t.Cell(2, 1).Range.Select
    Selection.Collapse
    s = "here: " & Replace(SKW_ApplyTableRolesHere(), vbCr, " / ")
    s = s & " ## r=" & ParaStyleOf(t.Cell(1, 1)) & "," & ParaStyleOf(t.Cell(2, 1)) & "," & ParaStyleOf(t.Cell(3, 1))
    t.Cell(2, 1).Range.Select
    Selection.Collapse
    s = s & " ## set: " & Replace(SKW_AssignTableRole("total"), vbCr, " / ")
    t.Cell(3, 1).Range.Text = "Zeta"
    t.Cell(1, 1).Range.Select
    Selection.Collapse
    s = s & " ## again: " & Replace(SKW_ApplyTableRolesHere(), vbCr, " / ")
    s = s & " ## kept=" & ParaStyleOf(t.Cell(2, 1)) & "," & ParaStyleOf(t.Cell(3, 1))
    t.Rows.Add
    t.Cell(4, 1).Range.Text = "New row"
    t.Cell(4, 1).Range.Style = d.Styles("Normal")
    s = s & " ## check: " & SKW_TableRolesCheck() & " ## " & ReportText()
    CloseReport
    d.Activate
    t.Cell(2, 1).Range.Select
    Selection.Collapse
    s = s & " ## suggest: " & Replace(SKW_AssignTableRole("suggest"), vbCr, " / ")
    s = s & " ## fresh=" & ParaStyleOf(t.Cell(2, 1)) & "," & ParaStyleOf(t.Cell(3, 1)) & "," & ParaStyleOf(t.Cell(4, 1))
    On Error Resume Next
    w0 = ActiveWindow.StyleAreaWidth
    ActiveWindow.View.Type = wdPrintView
    ActiveWindow.StyleAreaWidth = 0
    s = s & " ## show: " & Left$(SKW_ShowTableRoles(), 22) & " width>0=" & (ActiveWindow.StyleAreaWidth > 0)
    s = s & " hide: " & SKW_ShowTableRoles() & " width=" & ActiveWindow.StyleAreaWidth
    ActiveWindow.StyleAreaWidth = w0
    On Error GoTo 0
    SKW_SetStandardsPath prev
    TW_Assign = s
End Function
