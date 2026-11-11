Attribute VB_Name = "SK_Test"
' Self-tests, run by the build script. Only in test templates, never in the kit.
Option Explicit

Public Function T_Scan() As String
    Dim c As Collection, d As Object, s As String
    Set c = SK_Scan()
    s = SK_LastScanMethod & " | " & c.Count & " | " & SK_LastModuleCount & " |"
    For Each d In c
        s = s & " " & d("run") & "{" & d("comment") & "}"
    Next
    T_Scan = s
End Function

Public Function T_WordBasic() As String
    Dim s As String, ctx As Long, n As Long, i As Long
    On Error Resume Next
    For ctx = 0 To 1
        n = -1
        n = WordBasic.CountMacros(ctx, 0)
        s = s & "ctx" & ctx & "=" & n & ":"
        For i = 1 To n
            s = s & " " & WordBasic.[MacroName$](i, ctx, 0)
        Next
        s = s & " || "
        Err.Clear
    Next
    T_WordBasic = s
End Function

Public Function T_ModelSummary() As String
    Dim m As Object, e As Object, s As String
    Set m = SK_LoadModel()
    s = "found=" & m("foundCount") & " mods=" & m("moduleCount") & " by " & m("scanMethod") & " entries=" & m("entries").Count & " cats=" & m("cats").Count & " hist=" & m("history").Count & " |"
    For Each e In m("entries")
        s = s & " " & e("run") & "[" & SK_EntryStatus(e) & "/" & e("category") & "/" & e("relinked") & "]"
    Next
    T_ModelSummary = s
End Function

Public Function T_EditAndSave() As String
    Dim m As Object, e As Object, n As Long
    Set m = SK_LoadModel()
    For Each e In m("entries")
        n = n + 1
        If n <= 3 Then
            e("description") = "Test description " & n & " with <angle> & ""quotes"""
            e("category") = IIf(n = 3, "Fields", "Tables")
            e("keywords") = "alpha, beta"
            e("affects") = "Tables, Fields"
            e("shown") = "1"
        End If
    Next
    SK_AddUnique m("cats"), "Tables"
    SK_AddUnique m("cats"), "Fields"
    T_EditAndSave = SK_SaveModel(m, "self-test edit")
End Function

Public Function T_CatalogPresent() As String
    Dim x As String: x = SK_CatalogXml()
    T_CatalogPresent = Len(x) & " chars; parts=" & ThisDocument.CustomXMLParts.Count & "; " & Left$(x, 160)
End Function

Public Function T_MenuSearch(ByVal q As String) As String
    Dim s As String
    Load frmSKMenu
    s = frmSKMenu.TestSearch(q)
    Unload frmSKMenu
    T_MenuSearch = s
End Function

Public Function T_Admin() As String
    Dim s As String
    Load frmSKAdmin
    s = frmSKAdmin.TestState()
    s = s & " || labels: " & frmSKAdmin.TestLabels()
    s = s & " || rows: " & frmSKAdmin.TestRow(0) & " ## " & frmSKAdmin.TestRow(3) & " ## " & frmSKAdmin.TestRow(7)
    s = s & " || gate0: " & frmSKAdmin.TestSaveGate()
    s = s & " || " & frmSKAdmin.TestEditFirst("Edited by the admin self-test")
    s = s & " || gate1: " & frmSKAdmin.TestSaveGate()
    Unload frmSKAdmin
    T_Admin = s
End Function

Public Function T_Csv() As String
    Dim m As Object, p As String, msg As String, e As Object, s As String, txt As String, n As Long
    Set m = SK_LoadModel()
    p = SK_ExportCsv(m)
    txt = SK_ReadUtf8(p)
    txt = Replace(txt, "Test description 1", "Changed in Excel, line one" & vbCrLf & "line two, with comma")
    SK_WriteUtf8 p, txt
    msg = SK_UpdateFromRows(m, SK_ReadListFile(p), True, n)
    For Each e In m("entries")
        If InStr(e("description"), "Changed in Excel") > 0 Then s = e("run") & ": " & Replace(e("description"), vbCrLf, "\n")
    Next
    T_Csv = Replace(msg, vbCr, " / ") & " | " & s & " | " & p
End Function

Public Function T_Shortcut() As String
    T_Shortcut = SK_SetShortcut() & " || " & SK_AddButton()
End Function

Public Function T_Run(ByVal run As String) As String
    T_Run = SK_RunMacro(run) & " uses=" & SK_Uses(run)
End Function

Public Function T_Friendly() As String
    T_Friendly = SK_Friendly("FindBrokenRefs") & "|" & SK_Friendly("InsertCaption_Table") & "|" & SK_Friendly("TOCFix") & "|" & SK_Friendly("UpdateTOC")
End Function

Public Function T_AdminQuiet() As String
    Dim s As String
    SK_Quiet = True
    Load frmSKAdmin
    s = frmSKAdmin.TestState()
    Unload frmSKAdmin
    T_AdminQuiet = s
End Function

Public Function T_NoAccess() As String
    Dim s As String, c As New Collection, found As Collection, d As Object, root As String
    SK_ForceNoAccess = True
    SK_Quiet = True
    s = "model: " & T_ModelSummary() & " ##"
    s = s & " menu: " & T_MenuSearch("") & " ##"
    s = s & " search: " & T_MenuSearch("tab") & " ##"
    s = s & " admin: " & T_AdminQuiet() & " ##"
    root = Left$(ThisDocument.Path, InStrRev(ThisDocument.Path, "\"))
    c.Add root & "test\Module1.bas"
    c.Add root & "test\Module2.bas"
    Set found = SK_ReadBasFiles(c)
    s = s & " bas: " & found.Count & " |"
    For Each d In found
        s = s & " " & d("run") & "{" & d("comment") & "}"
    Next
    s = s & " ## msg: " & Replace(SK_NoAccessMessage(), vbCr, " / ")
    SK_ForceNoAccess = False
    SK_Quiet = False
    T_NoAccess = s
End Function

Public Function T_SetBackup(ByVal d As String) As String
    SK_SetBackupDir d
    T_SetBackup = SK_BackupDir()
End Function

Public Function T_Health() As String
    Dim m As Object, e As Object, s As String
    Set m = SK_LoadModel()
    s = "libs: " & SK_Health(m) & " ##"
    For Each e In m("entries")
        If Len(e("issues")) > 0 Then s = s & " " & e("run") & "{" & Replace(e("issues"), vbLf, "; ") & "}"
    Next
    T_Health = s
End Function

Public Function T_TestOne(ByVal run As String) As String
    T_TestOne = SK_TestMacro(run, "") & " | state=" & SK_TestState(run) & " | " & SK_TestDetail(run) & " | docs=" & Documents.Count
End Function

Public Function T_Marker() As String
    Dim m As Object
    SaveSetting SK_APP, SK_UserKey(), "mark|Module1.TOCFix", "test|2026-11-10 10:00"
    Set m = SK_LoadModel()
    T_Marker = SK_TestState("Module1.TOCFix") & " " & SK_TestDetail("Module1.TOCFix")
End Function

Public Function T_TestAll() As String
    Dim m As Object
    Set m = SK_LoadModel()
    SK_Health m
    T_TestAll = Replace(SK_TestAll(m, False), vbCr, " / ") & " | docs=" & Documents.Count
End Function

Public Function T_Undo() As String
    Dim d As Object, a As String, b As String
    Set d = Documents.Add(Template:=ThisDocument.FullName)
    d.Content.Text = "base"
    SK_RunMacro "Module3.TwoEdits", "Two Edits"
    a = d.Content.Text
    d.Undo 1
    b = d.Content.Text
    d.Close SaveChanges:=0
    T_Undo = "after run [" & Replace(a, vbCr, "") & "] after one undo [" & Replace(b, vbCr, "") & "]"
End Function

Public Function T_Needs() As String
    Dim d As Object, why As String, s As String, t As Object
    Set d = Documents.Add
    s = "table:" & SK_NeedsMet("Table at the cursor", why) & "(" & why & ")"
    s = s & " sel:" & SK_NeedsMet("Selection", why) & "(" & why & ")"
    Set t = d.Tables.Add(d.Range(0, 0), 2, 2)
    t.Cell(1, 1).Select
    s = s & " table after:" & SK_NeedsMet("Table at the cursor", why)
    d.Range.Select
    s = s & " sel after:" & SK_NeedsMet("Selection", why)
    s = s & " whole:" & SK_NeedsMet("Whole document", why)
    d.Close SaveChanges:=0
    T_Needs = s
End Function

Public Function T_ExportXlsx() As String
    Dim m As Object, p As String, rows As Collection
    Set m = SK_LoadModel()
    p = SK_ExportXlsx(m)
    Set rows = SK_ReadListFile(p)
    T_ExportXlsx = p & " | rows=" & rows.Count & " | first=" & rows(1)("run")
End Function

Public Function T_MakeDocxList(ByVal path As String) As String
    Dim d As Object, t As Object
    Set d = Documents.Add
    d.Content.Text = "Macro list for the team template"
    d.Content.InsertParagraphAfter
    Set t = d.Tables.Add(d.Paragraphs(2).Range, 3, 3)
    t.Cell(1, 1).Range.Text = "Macro": t.Cell(1, 2).Range.Text = "Category": t.Cell(1, 3).Range.Text = "Description"
    t.Cell(2, 1).Range.Text = "Module1.FitTablesToMargins": t.Cell(2, 2).Range.Text = "Tables"
    t.Cell(2, 3).Range.Text = "Sets every table to the full text width."
    t.Cell(3, 1).Range.Text = "NoSuch.Macro": t.Cell(3, 2).Range.Text = "X": t.Cell(3, 3).Range.Text = "Y"
    d.SaveAs2 FileName:=path, FileFormat:=16
    d.Close SaveChanges:=0
    T_MakeDocxList = path
End Function

Private Function T_Show(ByVal m As Object, ByVal key As String) As String
    Dim e As Object
    Set e = SK_FindEntry(m, key)
    If e Is Nothing Then T_Show = key & ": none": Exit Function
    T_Show = SK_NamePart(e("run")) & ": title=" & e("title") & "|cat=" & e("category") & "|desc=" & _
        Replace(e("description"), vbCrLf, "\n") & "|kw=" & e("keywords") & "|works=" & e("workson") & "|shown=" & e("shown")
End Function

Public Function T_UpdateFrom(ByVal path As String) As String
    Dim m As Object, rows As Collection, a As String, b As String, n1 As Long, n2 As Long, s As String, k As Variant
    Set m = SK_LoadModel()
    Set rows = SK_ReadListFile(path)
    a = SK_UpdateFromRows(m, rows, False, n1)
    b = SK_UpdateFromRows(m, rows, True, n2)
    s = "PREVIEW " & Replace(a, vbCr, " / ") & " || APPLIED " & n2 & " ||"
    For Each k In Array("ShowHiddenText", "ClearDirectFormatting", "TOCFix", "UpdateAllFields", "FitTablesToMargins")
        s = s & " " & T_Show(m, CStr(k)) & " ##"
    Next
    T_UpdateFrom = s
End Function

Public Function T_PrimaryLabel(ByVal path As String) As String
    SK_SetPrimaryList path
    Load frmSKAdmin
    T_PrimaryLabel = frmSKAdmin.TestPrimary()
    Unload frmSKAdmin
    SK_SetPrimaryList ""
End Function

' ---------------------------------------------------------------------
' 0.10.0: three kinds of macro, the Stationkeeper link
' ---------------------------------------------------------------------
Private Function T_SkMark(ByVal m As Object, ByVal key As String, ByVal kind As String) As String
    Dim e As Object
    Set e = SK_FindEntry(m, key)
    If e Is Nothing Then T_SkMark = key & "=missing": Exit Function
    e("stationkeeper") = kind
    e("shown") = "1"
    If Len(e("description")) = 0 Then e("description") = "Self-test macro."
    T_SkMark = e("run")
End Function

' Marks four menu macros: built, needs rules, needs the template, not related. Saves.
Public Function T_SkSetup() As String
    Dim m As Object, s As String
    Set m = SK_LoadModel()
    s = "built=" & T_SkMark(m, "ShowHiddenText", "Built for Stationkeeper")
    s = s & " needs=" & T_SkMark(m, "WriteTestMarker", "Needs Stationkeeper data: rules")
    s = s & " needsT=" & T_SkMark(m, "UpdateAllFields", "Needs Stationkeeper data: template")
    s = s & " none=" & T_SkMark(m, "FitTablesToMargins", "Not related")
    s = s & " | " & Split(SK_SaveModel(m, "self-test: Stationkeeper kinds"), vbCr)(0)
    T_SkSetup = s
End Function

' After a reload: the kinds kept in the catalog, and the tools' prefilled kinds.
Public Function T_SkKinds() As String
    Dim m As Object, s As String, k As Variant, e As Object
    Set m = SK_LoadModel()
    For Each k In Array("ShowHiddenText", "WriteTestMarker", "UpdateAllFields", "FitTablesToMargins", _
                        "SKW_Styles.StyleInspector", "SKW_Fields.ToggleFieldCodes", "SKX_Book.BookFontAudit", "SKP_Deck.DeckDeliveryCopy")
        Set e = SK_FindEntry(m, CStr(k))
        If e Is Nothing Then s = s & " " & k & "=missing" Else s = s & " " & SK_NamePart(e("run")) & "=" & e("stationkeeper") & ";"
    Next
    T_SkKinds = Trim$(s) & " | xml has stationkeeper=" & (InStr(SK_CatalogXml(), "<stationkeeper>Needs Stationkeeper data: rules</stationkeeper>") > 0)
End Function

' Links to a folder ("" unlinks) and reads the state: state | parts | note.
Public Function T_SkLink(ByVal folder As String) As String
    Dim d As Object
    SK_SetLinkFolder folder
    Set d = SK_LinkInfo()
    T_SkLink = d("state") & " | parts=" & d("parts") & " | rev=" & d("revision") & " | effort=" & d("effort") & " | " & d("note")
End Function

' Writes a stand-in for what Stationkeeper publishes. arg = "<effort folder>|current|stale|incomplete|bad".
Public Function T_SkMakeData(ByVal arg As String) As String
    Dim f As String, kind As String, dd As String, fso As Object, utc As String, rev As Long, rulesRev As Long
    Dim rules As String, man As String, bytes As Long, nowUtc As Date
    f = Split(arg, "|")(0): kind = Split(arg, "|")(1)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(f) Then fso.CreateFolder f
    dd = f & "\Stationkeeper Data"
    If Not fso.FolderExists(dd) Then fso.CreateFolder dd
    If Not fso.FolderExists(dd & "\rules") Then fso.CreateFolder dd & "\rules"
    nowUtc = Now + (Now - SK_FromUtcText(Format$(Now, "yyyy-mm-dd\Thh:nn:ss\Z")))
    If kind = "stale" Then nowUtc = nowUtc - 5
    utc = Format$(nowUtc, "yyyy-mm-dd\Thh:nn:ss\Z")
    rev = 15: rulesRev = 15
    If kind = "incomplete" Then rulesRev = 14
    rules = "{""schema_version"": ""1.0"", ""manifest_revision"": " & rulesRev & ", ""acronyms"": [{""acronym"": ""RFP"", ""definition"": ""Request for Proposal""}]}"
    SK_WriteUtf8 dd & "\rules\rules.json", rules
    bytes = FileLen(dd & "\rules\rules.json")
    man = "{""contract"": """ & IIf(kind = "bad", "something.else", "stationkeeper.effort-data") & """, ""schema_version"": ""1.0"", " & _
        """min_reader_version"": ""1.0"", ""effort"": {""id"": ""E-1"", ""name"": ""Effort A"", ""cui"": true}, " & _
        """revision"": " & rev & ", ""published_utc"": """ & utc & """, ""published_by"": ""tester"", " & _
        """freshness"": {""warn_after_hours"": 24, ""stale_after_hours"": 72}, " & _
        """parts"": [{""key"": ""rules"", ""path"": ""rules/rules.json"", ""format"": ""json"", ""part_revision"": 15, ""bytes"": " & bytes & "}]}"
    SK_WriteUtf8 dd & "\manifest.json", man
    T_SkMakeData = dd & " rules=" & bytes & " bytes, utc=" & utc
End Function

' The menu under each filter, the link strip, and what Run does for the four marked macros.
Public Function T_SkMenu() As String
    Dim s As String, k As Variant
    Load frmSKMenu
    s = "LINK " & frmSKMenu.TestLink()
    For Each k In Array("Needs effort data", "Stationkeeper", "Other", "Ready to run now")
        s = s & " || " & UCase$(k) & " " & frmSKMenu.TestShow(CStr(k))
    Next
    s = s & " || GATE built: " & frmSKMenu.TestGate("Module1.ShowHiddenText")
    s = s & " || GATE needs: " & frmSKMenu.TestGate("Module2.WriteTestMarker")
    s = s & " || GATE needsT: " & frmSKMenu.TestGate("Module1.UpdateAllFields")
    s = s & " || GATE none: " & frmSKMenu.TestGate("Module1.FitTablesToMargins")
    frmSKMenu.TestShow "All macros"
    Unload frmSKMenu
    T_SkMenu = s
End Function

Public Function T_SkAskPref(ByVal v As String) As String
    SK_SetPref "askBuilt", v
    T_SkAskPref = SK_GetPref("askBuilt", "1")
End Function

Public Function T_SkTestOne(ByVal run As String) As String
    Dim e As Object
    Set e = SK_FindEntry(SK_LoadModel(), run)
    T_SkTestOne = SK_TestMacro(e("run"), e("workson"), e("stationkeeper"))
End Function

' 1.0.0: a stand-in for Stationkeeper's publish with rules, template and dictionary.
' arg = "<effort folder>|<revision>|<revision written into rules.json>|<CaptionStyle>"
Public Function T_SkPublish(ByVal arg As String) As String
    Dim a As Variant, f As String, dd As String, fso As Object, rules As String, man As String, utc As String, p As Variant
    Dim parts As String, nowUtc As Date
    a = Split(arg, "|")
    f = a(0)
    Set fso = CreateObject("Scripting.FileSystemObject")
    dd = f & "\Stationkeeper Data"
    For Each p In Array(f, dd, dd & "\rules", dd & "\templates", dd & "\dictionaries")
        If Not fso.FolderExists(p) Then fso.CreateFolder p
    Next
    rules = "{""schema_version"": ""1.0"", ""manifest_revision"": " & a(2) & ", " & _
        """settings"": {""CaptionStyle"": """ & a(3) & """, ""TableLabel"": null, ""CaptionInFirstRow"": true}, " & _
        """banned_phrases"": [{""phrase"": ""leverage"", ""replacement"": ""use"", ""rule_id"": ""EF-1"", ""reason"": ""Effort rule"", ""match_case"": false, ""whole_word"": true}], " & _
        """placeholders"": []}"
    SK_WriteUtf8 dd & "\rules\rules.json", rules
    SK_WriteUtf8 dd & "\templates\effort.dotx", "stand-in template, revision " & a(1)
    SK_WriteUtf8 dd & "\dictionaries\effort.dic", "Effortword" & vbCrLf & "Stationkeeper"
    For Each p In Array(Array("rules", "rules/rules.json", "json"), Array("template", "templates/effort.dotx", "dotx"), _
                        Array("dictionary", "dictionaries/effort.dic", "dic-utf16"))
        parts = parts & IIf(Len(parts) > 0, ", ", "") & "{""key"": """ & p(0) & """, ""path"": """ & p(1) & """, ""format"": """ & p(2) & _
            """, ""part_revision"": " & a(1) & ", ""bytes"": " & FileLen(dd & "\" & Replace(p(1), "/", "\")) & "}"
    Next
    nowUtc = Now + (Now - SK_FromUtcText(Format$(Now, "yyyy-mm-dd\Thh:nn:ss\Z")))
    utc = Format$(nowUtc, "yyyy-mm-dd\Thh:nn:ss\Z")
    man = "{""contract"": ""stationkeeper.effort-data"", ""schema_version"": ""1.0"", ""min_reader_version"": ""1.0"", " & _
        """effort"": {""id"": ""E-2"", ""name"": ""Effort B"", ""cui"": true}, ""revision"": " & a(1) & ", ""published_utc"": """ & utc & _
        """, ""published_by"": ""tester"", ""freshness"": {""warn_after_hours"": 24, ""stale_after_hours"": 72}, ""parts"": [" & parts & "]}"
    SK_WriteUtf8 dd & "\manifest.json", man
    T_SkPublish = "published rev " & a(1)
End Function

Public Function T_SkClearCache() As String
    SK_ClearCache
    T_SkClearCache = "cleared"
End Function

' The link as the menu sees it, plus the file a team macro would read for each part.
Public Function T_SkLinkFiles() As String
    Dim d As Object
    Set d = SK_LinkInfo()
    T_SkLinkFiles = d("state") & " | cached=" & d("cached") & " | parts=" & d("parts") & " | rules file=" & SK_EffortFile("rules") & _
        " | " & d("note")
End Function

' 1.2.0: the template a team macro gets is a local copy, never Stationkeeper's own file.
Public Function T_SkAttachFile() As String
    Dim f As String, live As String, d As Object
    Set d = SK_LinkInfo()
    f = SK_EffortFile("template")
    live = d("dataDir") & "\templates\effort.dotx"
    T_SkAttachFile = "file=" & f & " | same=" & (SK_ReadUtf8(f) = SK_ReadUtf8(live)) & " | live=" & (LCase$(f) = LCase$(live))
End Function

' 1.1.0: sync warnings with stand-in file attributes ("" reads the real ones).
Public Function T_SkWarn(ByVal attrs As String) As String
    Dim d As Object
    SK_TestAttrs = attrs
    Set d = SK_LinkInfo()
    SK_TestAttrs = ""
    T_SkWarn = d("state") & " | warn=" & d("warn")
End Function

Public Function T_SkConflictNames() As String
    Dim v As Variant, s As String
    For Each v In Array("rules-DESKTOP1.json", "manifest (1).json", "effort-PC.dic", "effort-LAPTOP 7.dotx", "rules.json", _
                        "rules.json.tmp", "effort.dotx", "notes-x.json", "manifest.json")
        s = s & v & "=" & SK_IsConflictCopy(CStr(v)) & ";"
    Next
    T_SkConflictNames = s
End Function

' The strip's offer for a document in a folder. arg = "<document folder>|1" to click Use this folder.
Public Function T_SkSuggest(ByVal arg As String) As String
    Dim a As Variant
    a = Split(arg, "|")
    Load frmSKMenu
    T_SkSuggest = frmSKMenu.TestSuggest(CStr(a(0)), a(1) = "1")
    Unload frmSKMenu
End Function

' ---------------------------------------------------------------------
' 1.3.0: kit macros in the kit's menu, Add from a template, the filters
' ---------------------------------------------------------------------
Public Function T_KitEntries() As String
    Dim m As Object, e As Object, s As String
    Set m = SK_LoadModel()
    s = "installer=" & SK_HasInstaller() & " |"
    For Each e In m("entries")
        Select Case LCase$(SK_ModulePart(e("run")))
            Case "sk_install", "sk_macros"
                s = s & " " & e("run") & "[" & e("title") & "/" & e("category") & "/" & e("shown") & "/" & e("found") & "]"
        End Select
    Next
    T_KitEntries = s
End Function

' arg = "<Show>|<Category>"
Public Function T_MenuCat(ByVal arg As String) As String
    Dim a As Variant
    a = Split(arg, "|")
    Load frmSKMenu
    T_MenuCat = frmSKMenu.TestCat(CStr(a(0)), CStr(a(1)))
    Unload frmSKMenu
End Function

Public Function T_HideForTest(ByVal key As String) As String
    Dim m As Object, e As Object
    Set m = SK_LoadModel()
    Set e = SK_FindEntry(m, key)
    If e Is Nothing Then T_HideForTest = "missing": Exit Function
    e("shown") = "0"
    T_HideForTest = Split(SK_SaveModel(m, "self-test: hide " & key), vbCr)(0)
End Function

Public Function T_AdminFilters() As String
    Dim s As String
    SK_Quiet = True
    Load frmSKAdmin
    s = "choices: " & frmSKAdmin.TestFilterChoices()
    s = s & " || ALL " & frmSKAdmin.TestShowFilter("All visible", "")
    s = s & " || SK " & frmSKAdmin.TestShowFilter("Stationkeeper", "")
    s = s & " || OTHER " & frmSKAdmin.TestShowFilter("Other", "")
    s = s & " || HIDDEN " & frmSKAdmin.TestShowFilter("Hidden", "")
    s = s & " || NODESC " & frmSKAdmin.TestShowFilter("Description missing", "")
    s = s & " || SKFIELDS " & frmSKAdmin.TestShowFilter("Stationkeeper", "Fields")
    Unload frmSKAdmin
    T_AdminFilters = s
End Function

' arg = "<template path>|<modules to tick, A;B>|<1 to replace>"
Public Function T_AddFrom(ByVal arg As String) As String
    Dim a As Variant, s As String, d As Object
    a = Split(arg, "|")
    SK_Quiet = True
    Load frmSKAdmin
    s = frmSKAdmin.TestAddFrom(CStr(a(0)), Replace(CStr(a(1)), ";", "|"), a(2) = "1")
    s = s & " || IMPORTED " & frmSKAdmin.TestShowFilter("All visible", "Imported")
    Unload frmSKAdmin
    s = s & " || scan:"
    For Each d In SK_Scan()
        If LCase$(d("module")) = "extras" Or LCase$(d("name")) = "fromsource" Then s = s & " " & d("run")
    Next
    T_AddFrom = s
End Function
