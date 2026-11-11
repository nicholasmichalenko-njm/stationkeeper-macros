Attribute VB_Name = "SKW_Test"
' Test harness for Stationkeeper Tools. Imported into the sample template by test.ps1
' after the tools are installed. Every function returns text for test.ps1 to check.
Option Explicit

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
                        "SKW_DirectFormattingReport", "SKW_TableStandardsAudit", "SKW_StyleInspector")
        s = s & v & ": " & Application.Run(CStr(v)) & " / "
        CloseReport
    Next
    TW_AllReports = s
End Function

Public Function TW_ToolCount() As String
    Dim m As Object, e As Object, n As Long, s As String
    Set m = SK_LoadModel()
    For Each e In m("entries")
        If Left$(e("run"), 4) = "SKW_" Then
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
