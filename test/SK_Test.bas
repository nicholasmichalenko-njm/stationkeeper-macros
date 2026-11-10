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
    s = s & " || rows: " & frmSKAdmin.TestRow(0) & " ## " & frmSKAdmin.TestRow(3) & " ## " & frmSKAdmin.TestRow(7)
    s = s & " || gate0: " & frmSKAdmin.TestSaveGate()
    s = s & " || " & frmSKAdmin.TestEditFirst("Edited by the admin self-test")
    s = s & " || gate1: " & frmSKAdmin.TestSaveGate()
    Unload frmSKAdmin
    T_Admin = s
End Function

Public Function T_Csv() As String
    Dim m As Object, p As String, msg As String, e As Object, s As String, txt As String
    Set m = SK_LoadModel()
    p = SK_ExportCsv(m)
    txt = SK_ReadUtf8(p)
    txt = Replace(txt, "Test description 1", "Changed in Excel, line one" & vbCrLf & "line two, with comma")
    SK_WriteUtf8 p, txt
    msg = SK_ImportCsv(m, p)
    For Each e In m("entries")
        If InStr(e("description"), "Changed in Excel") > 0 Then s = e("run") & ": " & Replace(e("description"), vbCrLf, "\n")
    Next
    T_Csv = msg & " | " & s & " | " & p
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

' Runs inside a normally started Word (not automation), so Word's real
' "Trust access to the VBA project object model" setting applies.
Public Sub T_OffRun()
    Dim s As String, c As New Collection, found As Collection, d As Object, vbp As Object, n As Long
    SK_Quiet = True
    On Error Resume Next
    Set vbp = ThisDocument.VBProject
    n = -1: n = vbp.VBComponents.Count
    s = "VBA project readable: " & IIf(Err.Number = 0 And n >= 0, "yes", "no (" & Err.Number & ")") & vbCrLf
    Err.Clear
    s = s & "scan: " & T_Scan() & " err=" & Err.Number & vbCrLf: Err.Clear
    s = s & "model: " & T_ModelSummary() & " err=" & Err.Number & vbCrLf: Err.Clear
    s = s & "menu: " & T_MenuSearch("") & " err=" & Err.Number & vbCrLf: Err.Clear
    s = s & "menu search: " & T_MenuSearch("alp") & " err=" & Err.Number & vbCrLf: Err.Clear
    s = s & "admin: " & T_AdminQuiet() & " err=" & Err.Number & vbCrLf: Err.Clear
    c.Add Left$(ThisDocument.Path, InStrRev(ThisDocument.Path, "\")) & "test\Module1.bas"
    c.Add Left$(ThisDocument.Path, InStrRev(ThisDocument.Path, "\")) & "test\Module2.bas"
    Set found = SK_ReadBasFiles(c)
    s = s & "bas files: " & found.Count & " |"
    For Each d In found
        s = s & " " & d("run") & "{" & d("comment") & "}"
    Next
    s = s & " err=" & Err.Number & vbCrLf: Err.Clear
    s = s & "run: " & T_Run("Module1.ShowHiddenText") & " err=" & Err.Number & vbCrLf: Err.Clear
    s = s & "check: " & Replace(Left$(SK_CheckReport(SK_LoadModel()), 400), vbCrLf, " / ") & vbCrLf
    SK_WriteUtf8 ThisDocument.Path & "\off-results.txt", s
    Application.Quit SaveChanges:=0
End Sub

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
    s = s & " ## check: " & Replace(Left$(SK_CheckReport(SK_LoadModel()), 1500), vbCrLf, " / ")
    SK_ForceNoAccess = False
    SK_Quiet = False
    T_NoAccess = s
End Function
