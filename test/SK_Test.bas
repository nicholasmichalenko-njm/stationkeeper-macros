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
