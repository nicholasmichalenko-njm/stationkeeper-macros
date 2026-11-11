Attribute VB_Name = "SKW_Lib"
' =====================================================================
' Stationkeeper Tools: shared plumbing for the Word tools.
' Option Private Module keeps everything here out of Word's Macros list.
' Nothing here depends on the Stationkeeper Macros menu; the tools work
' with or without it.
' =====================================================================
Option Explicit
Option Private Module

Public Const SKW_APP As String = "Stationkeeper Tools"
Public Const SKW_VERSION As String = "1.2.0"

Public SKW_Quiet As Boolean              ' tests: no message boxes, reports made hidden
Public SKW_ForceCoauthored As Boolean    ' tests: behave as if the file is co-authored
Public SKW_LastReport As Object          ' the last report document made
Public SKW_LastRows As Long              ' findings in the last report
Public SKW_FindRows As Collection         ' the last report's rows, for Go to Next Finding
Public SKW_FindCol As Long                ' which column of those rows holds the document's text
Public SKW_FindDoc As String              ' the document that report checked
Public SKW_FindTitle As String
Public SKW_FindPos As Long

Private mUndo As Boolean
Private mScreen As Boolean
Private mStd As Object
Private mStdPath As String
Private mStdStamp As Date
Private mStdRead As Boolean
Private mMerged As Object
Private mMergedKey As String
Private mRules As Object
Private mRulesKey As String
Private mEff As Object
Private mEffAt As Date

' ---------------------------------------------------------------------
' Start and finish. changes = True for tools that change the document:
' they are blocked on co-authored files and grouped into one undo step.
' ---------------------------------------------------------------------
Public Function SKW_Begin(ByVal title As String, ByVal changes As Boolean) As String
    If Documents.Count = 0 Then
        SKW_Begin = "Open a document first."
        Exit Function
    End If
    If changes And SKW_IsCoauthored(ActiveDocument) Then
        SKW_Begin = "This document is on SharePoint or OneDrive with AutoSave on. To protect your co-authors' edits, " & _
            "tools that change the document do not run on it. Turn AutoSave off for this run, or use a copy. " & _
            "Report tools still run."
        Exit Function
    End If
    mScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    mUndo = False
    If changes Then
        On Error Resume Next
        If Not Application.UndoRecord.IsRecordingCustomRecord Then
            Application.UndoRecord.StartCustomRecord Left$("Stationkeeper: " & title, 64)
            mUndo = (Err.Number = 0)
        End If
        Err.Clear
    End If
End Function

Public Sub SKW_Finish()
    On Error Resume Next
    If mUndo Then Application.UndoRecord.EndCustomRecord
    mUndo = False
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Public Function SKW_IsCoauthored(ByVal d As Object) As Boolean
    Dim a As Boolean
    If SKW_ForceCoauthored Then SKW_IsCoauthored = True: Exit Function
    On Error Resume Next
    If LCase$(Left$(d.FullName, 8)) = "https://" Then
        a = True
        a = d.AutoSaveOn
        SKW_IsCoauthored = a
    End If
End Function

' Quiet while tests run: set here, or by the menu's tester through the registry.
Public Function SKW_IsQuiet() As Boolean
    SKW_IsQuiet = SKW_Quiet Or (GetSetting(SKW_APP, "Run", "quiet", "") = "1")
End Function

Public Sub SKW_Say(ByVal msg As String)
    If SKW_IsQuiet() Then Exit Sub
    MsgBox msg, vbInformation, SKW_APP
End Sub

' Shows a tool's result: the report if one was made, else a message.
Public Sub SKW_Show(ByVal summary As String)
    If SKW_IsQuiet() Then Exit Sub
    If Not SKW_LastReport Is Nothing Then
        On Error Resume Next
        SKW_LastReport.Activate
        Application.StatusBar = Left$(Replace(summary, vbCr, " "), 250)
    Else
        MsgBox summary, vbInformation, SKW_APP
    End If
End Sub

Public Sub SKW_Status(ByVal s As String)
    Application.StatusBar = s
End Sub

' ---------------------------------------------------------------------
' Stories: every part of the document (body, headers, footers, notes,
' text boxes), following the linked headers and footers too.
' ---------------------------------------------------------------------
Public Function SKW_Stories(ByVal d As Object) As Collection
    Dim c As New Collection, r As Object
    On Error Resume Next
    For Each r In d.StoryRanges
        Do While Not r Is Nothing
            c.Add r
            Set r = r.NextStoryRange
        Loop
    Next
    Set SKW_Stories = c
End Function

Public Function SKW_StoryName(ByVal t As Long) As String
    Select Case t
        Case wdMainTextStory: SKW_StoryName = "Body"
        Case wdFootnotesStory: SKW_StoryName = "Footnotes"
        Case wdEndnotesStory: SKW_StoryName = "Endnotes"
        Case wdCommentsStory: SKW_StoryName = "Comments"
        Case wdTextFrameStory: SKW_StoryName = "Text box"
        Case wdEvenPagesHeaderStory, wdPrimaryHeaderStory, wdFirstPageHeaderStory: SKW_StoryName = "Header"
        Case wdEvenPagesFooterStory, wdPrimaryFooterStory, wdFirstPageFooterStory: SKW_StoryName = "Footer"
        Case Else: SKW_StoryName = "Other"
    End Select
End Function

Public Function SKW_Page(ByVal r As Object) As String
    On Error Resume Next
    If r.StoryType = wdMainTextStory Or r.StoryType = wdFootnotesStory Or r.StoryType = wdEndnotesStory Then
        SKW_Page = CStr(r.Information(wdActiveEndAdjustedPageNumber))
    Else
        SKW_Page = "Section " & r.Information(wdActiveEndSectionNumber)
    End If
End Function

Public Function SKW_Excerpt(ByVal s As String, Optional ByVal n As Long = 90) As String
    Dim i As Long, ch As String, out As String
    s = Replace(Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), vbTab, " "), Chr$(7), " ")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If AscW(ch) >= 32 Or AscW(ch) < 0 Then out = out & ch
        If Len(out) >= n Then out = out & "...": Exit For
    Next
    SKW_Excerpt = Trim$(out)
End Function

' ---------------------------------------------------------------------
' Reports: a new Word document with the findings in one table.
' rows holds Variant arrays, one value per column.
' ---------------------------------------------------------------------
Public Function SKW_Report(ByVal title As String, ByVal intro As String, ByVal cols As Variant, _
                           ByVal rows As Collection) As Object
    Set SKW_Report = SKW_ReportFor(title, intro, cols, rows, "")
End Function

' sourceName: what was checked when it is not the active Word document (a deck, a workbook).
Public Function SKW_ReportFor(ByVal title As String, ByVal intro As String, ByVal cols As Variant, _
                              ByVal rows As Collection, ByVal sourceName As String) As Object
    ' Written as an HTML file and opened as a new document, so nothing is edited in Word while it is
    ' made: the checked document's Undo list cannot pick up a step (Word sometimes files edits made to
    ' a hidden new document under the active one).
    Dim src As Object, d As Object, h As String, r As Variant, i As Long, tmp As String, v As Variant
    Dim parts() As String, one As String, n As Long, srcName As String
    If Len(sourceName) = 0 Then
        Set src = ActiveDocument
        srcName = src.Name
    Else
        srcName = sourceName
    End If
    h = "<html><head><meta charset=""utf-8""><style>body{font-family:Calibri,sans-serif;font-size:10pt}" & _
        "h1{font-size:16pt}table{border-collapse:collapse}th,td{border:1px solid #999;padding:2pt 4pt;vertical-align:top;font-size:9pt}" & _
        "th{background:#DDEBF7;text-align:left}</style></head><body>" & _
        "<h1>" & SKW_Html(title) & "</h1><p>" & IIf(Len(sourceName) > 0, "File: ", "Document: ") & SKW_Html(srcName) & " &nbsp; " & Format$(Now, "yyyy-mm-dd hh:nn") & _
        " &nbsp; " & SKW_Html(SKW_StdNote()) & "</p><p>" & SKW_Html(intro) & "</p>"
    If rows.Count > 0 Then
        h = h & "<table><thead><tr>"
        For Each v In cols
            h = h & "<th>" & SKW_Html(CStr(v)) & "</th>"
        Next
        h = h & "</tr></thead><tbody>"
        ReDim parts(1 To rows.Count)
        n = 0
        For Each r In rows
            n = n + 1
            one = "<tr>"
            For i = LBound(r) To UBound(r)
                one = one & "<td>" & SKW_Html(SKW_Excerpt(CStr(r(i)), 400)) & "</td>"
            Next
            parts(n) = one & "</tr>"
        Next
        h = h & Join(parts, "") & "</tbody></table>"
    Else
        h = h & "<p><b>No findings.</b></p>"
    End If
    h = h & "</body></html>"
    tmp = Environ$("TEMP") & "\Stationkeeper report " & Format$(Now, "hhnnss") & Int(Rnd() * 10000) & ".htm"
    SKW_WriteUtf8 tmp, h
    On Error Resume Next
    Set d = Documents.Add(Template:=tmp, Visible:=Not SKW_IsQuiet())
    If d Is Nothing Then
        Err.Clear
        Set d = Documents.Open(FileName:=tmp, AddToRecentFiles:=False, Visible:=Not SKW_IsQuiet())
    End If
    Kill tmp
    On Error GoTo 0
    If Not d Is Nothing Then d.Saved = True
    SKW_LastRows = rows.Count
    ' Remember the findings so Go to Next Finding can walk through them in the document.
    SKW_FindCol = -1
    i = 0
    For Each v In cols
        Select Case LCase$(CStr(v))
            Case "text", "excerpt", "text before", "first use", "sentence", "first cell"
                If SKW_FindCol < 0 Then SKW_FindCol = i
        End Select
        i = i + 1
    Next
    Set SKW_FindRows = rows
    If src Is Nothing Then SKW_FindDoc = "" Else SKW_FindDoc = src.FullName
    SKW_FindTitle = title
    SKW_FindPos = 0
    Set SKW_LastReport = d
    Set SKW_ReportFor = d
    If SKW_IsQuiet() And Not src Is Nothing Then src.Activate
End Function

Public Function SKW_Html(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    SKW_Html = Replace(Replace(s, vbCrLf, "<br>"), vbCr, "<br>")
End Function

Public Sub SKW_WriteUtf8(ByVal path As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.SaveToFile path, 2
    st.Close
End Sub

Public Function SKW_Row(ParamArray v() As Variant) As Variant
    SKW_Row = v
End Function

' ---------------------------------------------------------------------
' The rules: the linked Stationkeeper effort first, then the standards
' workbook (one per template, chosen per person), then built-in defaults.
' The workbook is read once and kept until the file changes.
' ---------------------------------------------------------------------
Public Function SKW_StandardsPath() As String
    SKW_StandardsPath = GetSetting(SKW_APP, "Standards", ThisDocument.Name, "")
End Function

Public Sub SKW_SetStandardsPath(ByVal p As String)
    SaveSetting SKW_APP, "Standards", ThisDocument.Name, p
    Set mStd = Nothing
    mStdRead = False
End Sub

Public Function SKW_HasStandards() As Boolean
    SKW_HasStandards = Not (SKW_Std() Is Nothing)
End Function

' Where the rules came from, for the top of every report.
Public Function SKW_StdNote() As String
    Dim e As Object, wb As String
    Set e = SKW_Effort()
    If Not (SKW_Workbook() Is Nothing) Then wb = Mid$(mStdPath, InStrRev(mStdPath, "\") + 1)
    If e("ok") Then
        SKW_StdNote = "Rules: effort " & e("effort") & " rev " & e("revision") & _
            IIf(e("fromCache"), " (last complete copy: " & LCase$(e("state")) & ")", "") & _
            IIf(Len(wb) > 0, ", then " & wb, "")
    ElseIf Len(wb) > 0 Then
        SKW_StdNote = "Standards: " & wb & IIf(e("state") <> "Not linked", " (effort data: " & LCase$(e("state")) & ")", "")
    Else
        SKW_StdNote = "Standards: built-in defaults (no standards workbook chosen)" & _
            IIf(e("state") <> "Not linked", " (effort data: " & LCase$(e("state")) & ")", "")
    End If
End Function

' The rules the tools use: the linked effort's rules first, then the standards workbook for
' anything the effort does not set, then (in each tool) the built-in defaults.
Public Function SKW_Std() As Object
    Dim wbk As Object, e As Object, key As String
    Set wbk = SKW_Workbook()
    Set e = SKW_Effort()
    If Not e("ok") Then Set SKW_Std = wbk: Exit Function
    key = mStdPath & "|" & mStdStamp & "|" & IIf(wbk Is Nothing, "none", "wb") & "|" & e("key")
    If key = mMergedKey And Not mMerged Is Nothing Then Set SKW_Std = mMerged: Exit Function
    Set mMerged = SKW_Merge(wbk, e)
    mMergedKey = key
    Set SKW_Std = mMerged
End Function

' The standards workbook alone (read again whenever the file changes), or Nothing.
Public Function SKW_Workbook() As Object
    Dim p As String, stamp As Date
    p = SKW_StandardsPath()
    If Len(p) = 0 Then Set SKW_Workbook = Nothing: Exit Function
    On Error Resume Next
    stamp = FileDateTime(p)
    If Err.Number <> 0 Then Set SKW_Workbook = Nothing: Exit Function
    On Error GoTo 0
    If mStdRead And p = mStdPath And stamp = mStdStamp Then Set SKW_Workbook = mStd: Exit Function
    Set mStd = SKW_ReadWorkbook(p)
    mStdPath = p
    mStdStamp = stamp
    mStdRead = True
    Set SKW_Workbook = mStd
End Function

' Workbook sheets with the effort's sections laid over them. A section the effort sets replaces the
' sheet (even when its list is empty); settings are laid over one by one; the effort's template and
' dictionary become StyleSource and the Dictionaries sheet.
Private Function SKW_Merge(ByVal wbk As Object, ByVal e As Object) As Object
    Dim res As Object, k As Variant, rules As Object, sheet As String, rows As Collection, item As Variant
    Dim d As Object, f As Variant, setRows As Collection, r As Object, sk As Variant, done As Object, st As Object
    Set res = CreateObject("Scripting.Dictionary")
    If Not wbk Is Nothing Then
        For Each k In wbk.Keys
            Set res(k) = wbk(k)
        Next
    End If
    Set rules = e("rules")
    For Each k In rules.Keys
        If k <> "settings" And IsObject(rules(k)) Then
            If TypeName(rules(k)) = "Collection" Then
                sheet = Replace(LCase$(k), "_", " ")
                Set rows = New Collection
                For Each item In rules(k)
                    If IsObject(item) Then
                        Set d = CreateObject("Scripting.Dictionary")
                        For Each f In item.Keys
                            If Not IsObject(item(f)) Then d(Replace(LCase$(f), "_", " ")) = SKW_JText(item(f))
                        Next
                        rows.Add d
                    End If
                Next
                Set res(sheet) = rows
            End If
        End If
    Next
    ' Settings, one by one.
    Set setRows = New Collection
    Set done = CreateObject("Scripting.Dictionary")
    If rules.Exists("settings") Then
        If TypeName(rules("settings")) = "Dictionary" Then
            Set st = rules("settings")
            For Each sk In st.Keys
                If Not IsObject(st(sk)) Then
                    If Not IsNull(st(sk)) Then done(LCase$(sk)) = SKW_JText(st(sk))
                End If
            Next
        End If
    End If
    If Len(e("template")) > 0 Then done("stylesource") = e("template")
    If res.Exists("settings") Then
        For Each r In res("settings")
            If done.Exists(LCase$(r("setting"))) Then
                Set d = CreateObject("Scripting.Dictionary")
                d("setting") = r("setting"): d("value") = done(LCase$(r("setting"))): d("from") = "effort"
                setRows.Add d
                done.Remove LCase$(r("setting"))
            Else
                setRows.Add r
            End If
        Next
    End If
    For Each sk In done.Keys
        Set d = CreateObject("Scripting.Dictionary")
        d("setting") = sk: d("value") = done(sk): d("from") = "effort"
        setRows.Add d
    Next
    Set res("settings") = setRows
    If Len(e("dictionary")) > 0 Then
        Set rows = New Collection
        Set d = CreateObject("Scripting.Dictionary")
        d("path") = e("dictionary"): d("notes") = "Effort dictionary from Stationkeeper"
        rows.Add d
        Set res("dictionaries") = rows
    End If
    Set SKW_Merge = res
End Function

Private Function SKW_JText(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then SKW_JText = IIf(v, "Yes", "No") Else SKW_JText = CStr(v)
End Function

' ---------------------------------------------------------------------
' The linked effort's data (written by the Stationkeeper app; read only).
' The link is one folder per person per PC (set in the menu). Each run
' reads the manifest. A complete new revision is copied to a local cache
' (this person's profile, one effort, removed on Unlink); while the folder
' is mid-save, mid-sync or out of reach, the last complete copy is used.
' ---------------------------------------------------------------------
Public Function SKW_LinkFolder() As String
    SKW_LinkFolder = GetSetting(SKW_APP, "Link", "folder", "")
End Function

Public Function SKW_CacheDir() As String
    SKW_CacheDir = Environ$("LOCALAPPDATA") & "\Stationkeeper Tools\Effort cache"
End Function

Private Function SKW_DataDir(ByVal f As String) As String
    If Right$(f, 1) = "\" Then f = Left$(f, Len(f) - 1)
    If LCase$(Mid$(f, InStrRev(f, "\") + 1)) = "stationkeeper data" Then SKW_DataDir = f Else SKW_DataDir = f & "\Stationkeeper Data"
End Function

' Keys: ok, state, effort, revision, fromCache, rules (parsed rules.json), template, dictionary, key.
' Read fresh on each run; inside one run (3 seconds) the answer is reused.
Public Function SKW_Effort() As Object
    If Not mEff Is Nothing Then
        If (Now - mEffAt) * 86400# < 3 And (Now - mEffAt) >= 0 Then Set SKW_Effort = mEff: Exit Function
    End If
    Set mEff = SKW_ReadEffort()
    mEffAt = Now
    Set SKW_Effort = mEff
End Function

' Forget the last answer (after a new link, and in tests).
Public Sub SKW_ForgetEffort()
    Set mEff = Nothing
End Sub

' True when the linked effort sets this sheet (its list wins, even when empty).
Public Function SKW_FromEffort(ByVal sheet As String) As Boolean
    Dim e As Object
    Set e = SKW_Effort()
    If e("ok") Then SKW_FromEffort = e("rules").Exists(Replace(LCase$(sheet), " ", "_"))
End Function

Private Function SKW_ReadEffort() As Object
    Dim d As Object, f As String, dd As String, fso As Object, man As Object, why As String, cache As String
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False: d("state") = "Not linked": d("effort") = "": d("revision") = "": d("fromCache") = False
    d("template") = "": d("dictionary") = "": d("key") = ""
    Set d("rules") = Nothing
    Set SKW_ReadEffort = d
    f = SKW_LinkFolder()
    If Len(f) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    cache = SKW_CacheDir()
    dd = SKW_DataDir(f)
    If Not fso.FolderExists(f) Then
        d("state") = "Folder missing"
    ElseIf Not fso.FileExists(dd & "\manifest.json") Then
        d("state") = "No data yet"
    Else
        Set man = SKW_CheckData(dd, fso, why)
        If man Is Nothing Then
            d("state") = why
        Else
            d("state") = "Current"
            If SKW_CachedRevision(dd, fso) <> SKW_J(man, "revision") Then
                If Not SKW_CopyToCache(dd, man, fso) Then d("state") = "Incomplete"
            End If
        End If
    End If
    ' Load from the cache (it now holds this revision, or the last complete one).
    If SKW_CacheSource(fso) <> LCase$(dd) Then Exit Function
    If d("state") = "Current" Or d("state") = "Incomplete" Or d("state") = "Unreadable" Or d("state") = "Folder missing" Then
        If d("state") <> "Current" Then d("fromCache") = True
        SKW_LoadCache d, fso
    End If
End Function

' Returns the manifest when every part is there at full size and rules.json carries the same revision; else why.
Private Function SKW_CheckData(ByVal dd As String, ByVal fso As Object, ByRef why As String) As Object
    Dim man As Object, p As Variant, fp As String, rules As Object
    On Error GoTo Bad
    Set man = SKW_JParse(SKW_ReadText(dd & "\manifest.json"))
    If SKW_J(man, "contract") <> "stationkeeper.effort-data" Then GoTo Bad
    If Val(SKW_J(man, "min_reader_version")) > 1 Then why = "Needs a newer kit": Exit Function
    For Each p In man("parts")
        fp = dd & "\" & Replace(SKW_J(p, "path"), "/", "\")
        If Not fso.FileExists(fp) Then why = "Incomplete": Exit Function
        If Len(SKW_J(p, "bytes")) > 0 Then
            If CStr(FileLen(fp)) <> SKW_J(p, "bytes") Then why = "Incomplete": Exit Function
        End If
        If LCase$(SKW_J(p, "key")) = "rules" Then
            Set rules = SKW_JParse(SKW_ReadText(fp))
            If SKW_J(rules, "manifest_revision") <> SKW_J(man, "revision") Then why = "Incomplete": Exit Function
        End If
    Next
    Set SKW_CheckData = man
    Exit Function
Bad:
    why = "Unreadable"
End Function

' The cache: source.txt (which Stationkeeper Data folder), current.txt (which revision folder),
' and one folder per revision (r15\manifest.json, r15\rules\rules.json, ...). A new revision goes into
' its own folder, so a dictionary or template Word still holds open is never overwritten.
Private Function SKW_CacheSource(ByVal fso As Object) As String
    On Error Resume Next
    If fso.FileExists(SKW_CacheDir() & "\source.txt") Then SKW_CacheSource = LCase$(SKW_Line(SKW_ReadText(SKW_CacheDir() & "\source.txt")))
End Function

Private Function SKW_Line(ByVal s As String) As String
    SKW_Line = Trim$(Replace(Replace(s, vbCr, ""), vbLf, ""))
End Function

' The folder of the revision in use, or "".
Public Function SKW_CacheRoot() As String
    Dim fso As Object, cur As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(SKW_CacheDir() & "\current.txt") Then Exit Function
    cur = SKW_Line(SKW_ReadText(SKW_CacheDir() & "\current.txt"))
    If Len(cur) > 0 And fso.FileExists(SKW_CacheDir() & "\" & cur & "\manifest.json") Then SKW_CacheRoot = SKW_CacheDir() & "\" & cur
End Function

Private Function SKW_CachedRevision(ByVal dd As String, ByVal fso As Object) As String
    On Error Resume Next
    If SKW_CacheSource(fso) <> LCase$(dd) Then Exit Function
    If Len(SKW_CacheRoot()) = 0 Then Exit Function
    SKW_CachedRevision = SKW_J(SKW_JParse(SKW_ReadText(SKW_CacheRoot() & "\manifest.json")), "revision")
End Function

' Copies each part into a new revision folder, checks the copies (size, revision), writes the manifest
' there, then points current.txt at it. Older revision folders are removed when nothing holds them.
Private Function SKW_CopyToCache(ByVal dd As String, ByVal man As Object, ByVal fso As Object) As Boolean
    Dim c As String, rv As String, p As Variant, rel As String, dst As String, sub1 As Object
    On Error GoTo Fail
    c = SKW_CacheDir()
    SKW_MakeDirs c, fso
    If SKW_CacheSource(fso) <> LCase$(dd) Then
        On Error Resume Next
        If fso.FileExists(c & "\current.txt") Then fso.DeleteFile c & "\current.txt", True
        On Error GoTo Fail
        SKW_WriteUtf8 c & "\source.txt", dd
    End If
    rv = "r" & SKW_J(man, "revision")
    If fso.FolderExists(c & "\" & rv) Then fso.DeleteFolder c & "\" & rv, True
    For Each p In man("parts")
        rel = Replace(SKW_J(p, "path"), "/", "\")
        dst = c & "\" & rv & "\" & rel
        SKW_MakeDirs Left$(dst, InStrRev(dst, "\") - 1), fso
        fso.CopyFile dd & "\" & rel, dst, True
        If Len(SKW_J(p, "bytes")) > 0 Then
            If CStr(FileLen(dst)) <> SKW_J(p, "bytes") Then GoTo Fail
        End If
        If LCase$(SKW_J(p, "key")) = "rules" Then
            If SKW_J(SKW_JParse(SKW_ReadText(dst)), "manifest_revision") <> SKW_J(man, "revision") Then GoTo Fail
        End If
    Next
    fso.CopyFile dd & "\manifest.json", c & "\" & rv & "\manifest.json", True
    SKW_WriteUtf8 c & "\current.txt", rv
    On Error Resume Next
    For Each sub1 In fso.GetFolder(c).SubFolders
        If LCase$(sub1.Name) <> LCase$(rv) Then sub1.Delete True
    Next
    SKW_CopyToCache = True
    Exit Function
Fail:
    SKW_CopyToCache = False
End Function

Private Sub SKW_MakeDirs(ByVal p As String, ByVal fso As Object)
    If fso.FolderExists(p) Then Exit Sub
    SKW_MakeDirs fso.GetParentFolderName(p), fso
    fso.CreateFolder p
End Sub

Private Sub SKW_LoadCache(ByVal d As Object, ByVal fso As Object)
    Dim c As String, man As Object, p As Variant, k As String, fp As String, eff As Variant
    On Error GoTo Fail
    c = SKW_CacheRoot()
    If Len(c) = 0 Then Exit Sub
    Set man = SKW_JParse(SKW_ReadText(c & "\manifest.json"))
    d("revision") = SKW_J(man, "revision")
    If man.Exists("effort") Then
        If IsObject(man("effort")) Then d("effort") = SKW_J(man("effort"), "name")
    End If
    If Len(d("effort")) = 0 Then d("effort") = "(unnamed)"
    For Each p In man("parts")
        k = LCase$(SKW_J(p, "key"))
        fp = c & "\" & Replace(SKW_J(p, "path"), "/", "\")
        If fso.FileExists(fp) Then
            Select Case k
                Case "rules"
                    If mRulesKey <> fp & "|" & d("revision") Then
                        Set mRules = SKW_JParse(SKW_ReadText(fp))
                        mRulesKey = fp & "|" & d("revision")
                    End If
                    Set d("rules") = mRules
                Case "template": d("template") = fp
                Case "dictionary": d("dictionary") = fp
            End Select
        End If
    Next
    If d("rules") Is Nothing Then Set d("rules") = CreateObject("Scripting.Dictionary")
    d("ok") = True
    d("key") = SKW_CacheSource(fso) & "|" & d("revision")
    Exit Sub
Fail:
    d("ok") = False
End Sub

Private Function SKW_ReadText(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    SKW_ReadText = st.ReadText
    st.Close
End Function

' --- A small JSON reader (objects become Dictionaries, arrays Collections, numbers text) ---
Private Function SKW_J(ByVal o As Variant, ByVal k As String) As String
    If Not IsObject(o) Then Exit Function
    If TypeName(o) <> "Dictionary" Then Exit Function
    If Not o.Exists(k) Then Exit Function
    If IsObject(o(k)) Then Exit Function
    If IsNull(o(k)) Then Exit Function
    SKW_J = CStr(o(k))
End Function

Private Function SKW_JParse(ByVal s As String) As Object
    Dim pos As Long, v As Variant
    pos = 1
    SKW_JValue v, s, pos
    Set SKW_JParse = v
End Function

Private Sub SKW_JSkip(ByVal s As String, pos As Long)
    Do While pos <= Len(s)
        Select Case Mid$(s, pos, 1)
            Case " ", vbTab, vbCr, vbLf, ChrW$(&HFEFF): pos = pos + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Sub SKW_JValue(ByRef v As Variant, ByVal s As String, pos As Long)
    Dim num As String, d As Object, c As Collection, k As String, x As Variant
    SKW_JSkip s, pos
    If pos > Len(s) Then Err.Raise vbObjectError + 7, SKW_APP, "JSON ends too early."
    Select Case Mid$(s, pos, 1)
        Case "{"
            Set d = CreateObject("Scripting.Dictionary")
            pos = pos + 1
            Do
                SKW_JSkip s, pos
                If pos > Len(s) Then Err.Raise vbObjectError + 7, SKW_APP, "JSON ends too early."
                If Mid$(s, pos, 1) = "}" Then pos = pos + 1: Exit Do
                If Mid$(s, pos, 1) = "," Then pos = pos + 1: SKW_JSkip s, pos
                k = SKW_JString(s, pos)
                SKW_JSkip s, pos
                If Mid$(s, pos, 1) <> ":" Then Err.Raise vbObjectError + 7, SKW_APP, "JSON is missing a colon."
                pos = pos + 1
                SKW_JValue x, s, pos
                If IsObject(x) Then Set d(k) = x Else d(k) = x
            Loop
            Set v = d
        Case "["
            Set c = New Collection
            pos = pos + 1
            Do
                SKW_JSkip s, pos
                If pos > Len(s) Then Err.Raise vbObjectError + 7, SKW_APP, "JSON ends too early."
                If Mid$(s, pos, 1) = "]" Then pos = pos + 1: Exit Do
                If Mid$(s, pos, 1) = "," Then pos = pos + 1: SKW_JSkip s, pos
                SKW_JValue x, s, pos
                c.Add x
            Loop
            Set v = c
        Case """": v = SKW_JString(s, pos)
        Case "t": v = True: pos = pos + 4
        Case "f": v = False: pos = pos + 5
        Case "n": v = Null: pos = pos + 4
        Case Else
            Do While pos <= Len(s) And InStr("+-0123456789.eE", Mid$(s, pos, 1)) > 0
                num = num & Mid$(s, pos, 1)
                pos = pos + 1
            Loop
            If Len(num) = 0 Then Err.Raise vbObjectError + 7, SKW_APP, "JSON has something unexpected at character " & pos & "."
            v = num
    End Select
End Sub

Private Function SKW_JString(ByVal s As String, pos As Long) As String
    Dim out As String, ch As String
    If Mid$(s, pos, 1) <> """" Then Err.Raise vbObjectError + 7, SKW_APP, "JSON expected text at character " & pos & "."
    pos = pos + 1
    Do While pos <= Len(s)
        ch = Mid$(s, pos, 1)
        If ch = """" Then
            pos = pos + 1
            SKW_JString = out
            Exit Function
        ElseIf ch = "\" Then
            pos = pos + 1
            ch = Mid$(s, pos, 1)
            Select Case ch
                Case "n": out = out & vbCrLf
                Case "r", "b", "f"
                Case "t": out = out & vbTab
                Case "u"
                    out = out & ChrW$(CLng("&H" & Mid$(s, pos + 1, 4)))
                    pos = pos + 4
                Case Else: out = out & ch
            End Select
        Else
            out = out & ch
        End If
        pos = pos + 1
    Loop
    Err.Raise vbObjectError + 7, SKW_APP, "JSON has text that never ends."
End Function

' Every sheet: lower-case sheet name -> Collection of rows (lower-case header -> text).
Private Function SKW_ReadWorkbook(ByVal path As String) As Object
    Dim xl As Object, wb As Object, ws As Object, v As Variant, r As Long, c As Long, res As Object
    Dim rows As Collection, d As Object, h As String, hasAny As Boolean, ed As String
    Set res = CreateObject("Scripting.Dictionary")
    On Error GoTo Fail
    SKW_Status "Reading the standards workbook..."
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(FileName:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    For Each ws In wb.Worksheets
        v = ws.UsedRange.Value
        Set rows = New Collection
        If IsArray(v) Then
            For r = LBound(v, 1) + 1 To UBound(v, 1)
                Set d = CreateObject("Scripting.Dictionary")
                hasAny = False
                For c = LBound(v, 2) To UBound(v, 2)
                    h = LCase$(Trim$(SKW_Cell(v(LBound(v, 1), c))))
                    If Len(h) > 0 Then
                        d(h) = Trim$(SKW_Cell(v(r, c)))
                        If Len(d(h)) > 0 Then hasAny = True
                    End If
                Next
                If hasAny Then rows.Add d
            Next
        End If
        Set res(LCase$(Trim$(ws.Name))) = rows
    Next
    wb.Close False
    xl.Quit
    Set SKW_ReadWorkbook = res
    Exit Function
Fail:
    ed = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Set SKW_ReadWorkbook = Nothing
    SKW_Say "Could not read the standards workbook (" & ed & "). The tools use built-in defaults."
End Function

Private Function SKW_Cell(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    SKW_Cell = CStr(v)
End Function

Public Function SKW_Rows(ByVal sheet As String) As Collection
    Dim s As Object
    Set s = SKW_Std()
    If s Is Nothing Then Set SKW_Rows = New Collection: Exit Function
    If s.Exists(LCase$(sheet)) Then Set SKW_Rows = s(LCase$(sheet)) Else Set SKW_Rows = New Collection
End Function

Public Function SKW_Setting(ByVal key As String, ByVal dflt As String) As String
    Dim r As Object
    SKW_Setting = dflt
    For Each r In SKW_Rows("Settings")
        If LCase$(r("setting")) = LCase$(key) Then
            If Len(r("value")) > 0 Then SKW_Setting = r("value")
            Exit Function
        End If
    Next
End Function

Public Function SKW_IsYes(ByVal v As String) As Boolean
    Select Case LCase$(Trim$(v))
        Case "yes", "y", "true", "1": SKW_IsYes = True
    End Select
End Function

' One column of a sheet as a list of non-blank values.
Public Function SKW_Column(ByVal sheet As String, ByVal col As String) As Collection
    Dim c As New Collection, r As Object
    For Each r In SKW_Rows(sheet)
        If r.Exists(LCase$(col)) Then
            If Len(r(LCase$(col))) > 0 Then c.Add r(LCase$(col))
        End If
    Next
    Set SKW_Column = c
End Function

Public Function SKW_InList(ByVal c As Collection, ByVal s As String) As Boolean
    Dim v As Variant
    For Each v In c
        If LCase$(v) = LCase$(s) Then SKW_InList = True: Exit Function
    Next
End Function

Public Function SKW_SplitList(ByVal s As String) As Collection
    Dim c As New Collection, v As Variant
    For Each v In Split(Replace(s, ";", ","), ",")
        If Len(Trim$(v)) > 0 Then c.Add Trim$(v)
    Next
    Set SKW_SplitList = c
End Function

' Placeholders: the effort's or the workbook's list, else the defaults.
Public Function SKW_Placeholders() As Collection
    Dim c As Collection, v As Variant
    Set c = SKW_Column("Placeholders", "Text")
    If c.Count = 0 And Not SKW_FromEffort("Placeholders") Then
        For Each v In Array("TBD", "TBR", "TBS", "XXX", "[insert", "Lorem ipsum", "<<", ">>", "???", _
                            "Click or tap here to enter text")
            c.Add v
        Next
    End If
    Set SKW_Placeholders = c
End Function

' A style's font name, with theme fonts resolved (+Body, +Headings).
Public Function SKW_StyleFont(ByVal d As Object, ByVal st As Object) As String
    Dim n As String
    On Error Resume Next
    n = st.Font.Name
    If Left$(n, 1) = "+" Then
        If InStr(LCase$(n), "head") > 0 Then
            n = d.DocumentTheme.ThemeFontScheme.MajorFont(1).Name
        Else
            n = d.DocumentTheme.ThemeFontScheme.MinorFont(1).Name
        End If
    End If
    SKW_StyleFont = n
End Function

' The paragraph style of a range, read without touching Word's undo list.
' (Reading Paragraph.Style from VBA adds a "VBA-Paragraph.Style" step to Undo; Range.ParagraphStyle does not.)
Public Function SKW_ParaStyle(ByVal r As Object) As Object
    On Error Resume Next
    Set SKW_ParaStyle = r.ParagraphStyle
    ' A range holding field codes (a table of contents, say) can answer Nothing: ask its paragraph mark.
    If SKW_ParaStyle Is Nothing Then Set SKW_ParaStyle = r.Paragraphs(1).Range.Characters.Last.ParagraphStyle
    If SKW_ParaStyle Is Nothing Then Set SKW_ParaStyle = r.Paragraphs(1).Style
End Function

Public Function SKW_ParaStyleName(ByVal r As Object) As String
    Dim st As Object
    Set st = SKW_ParaStyle(r)
    If Not st Is Nothing Then SKW_ParaStyleName = st.NameLocal
End Function

' A style by name, or Nothing.
Public Function SKW_GetStyle(ByVal d As Object, ByVal nm As String) As Object
    On Error Resume Next
    Set SKW_GetStyle = d.Styles(nm)
End Function

' Ranges of a paragraph or character style in one story (Find, not a paragraph loop).
Public Function SKW_StyleRanges(ByVal story As Object, ByVal st As Object, Optional ByVal limit As Long = 100000) As Collection
    Dim c As New Collection, f As Object, lastStart As Long
    Set f = story.Duplicate
    On Error GoTo Done
    With f.Find
        .ClearFormatting
        .Text = ""
        .Format = True
        .Style = st
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
    End With
    lastStart = -1
    Do While f.Find.Execute
        If f.Start = lastStart Then Exit Do
        lastStart = f.Start
        c.Add f.Duplicate
        If c.Count >= limit Then Exit Do
        If f.End >= story.End Then Exit Do
        f.Collapse wdCollapseEnd
    Loop
Done:
    Set SKW_StyleRanges = c
End Function

' Ranges of a piece of text in one story.
Public Function SKW_TextRanges(ByVal story As Object, ByVal txt As String, ByVal wholeWord As Boolean, _
                               ByVal matchCase As Boolean, Optional ByVal wildcards As Boolean = False, _
                               Optional ByVal limit As Long = 5000) As Collection
    Dim c As New Collection, f As Object, lastStart As Long
    Set f = story.Duplicate
    On Error GoTo Done
    With f.Find
        .ClearFormatting
        .Text = txt
        .Format = False
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = wildcards
        If Not wildcards Then
            .MatchWholeWord = wholeWord
            .MatchCase = matchCase
        End If
    End With
    lastStart = -1
    Do While f.Find.Execute
        If f.Start = lastStart Then Exit Do
        lastStart = f.Start
        c.Add f.Duplicate
        If c.Count >= limit Then Exit Do
        If f.End >= story.End Then Exit Do
        f.Collapse wdCollapseEnd
    Loop
Done:
    Set SKW_TextRanges = c
End Function

Public Function SKW_FieldTypeName(ByVal t As Long) As String
    Select Case t
        Case wdFieldRef: SKW_FieldTypeName = "REF"
        Case wdFieldPageRef: SKW_FieldTypeName = "PAGEREF"
        Case wdFieldNoteRef: SKW_FieldTypeName = "NOTEREF"
        Case wdFieldSequence: SKW_FieldTypeName = "SEQ"
        Case wdFieldStyleRef: SKW_FieldTypeName = "STYLEREF"
        Case wdFieldTOC: SKW_FieldTypeName = "TOC"
        Case wdFieldTOCEntry: SKW_FieldTypeName = "TC"
        Case wdFieldPage: SKW_FieldTypeName = "PAGE"
        Case wdFieldNumPages: SKW_FieldTypeName = "NUMPAGES"
        Case wdFieldSectionPages: SKW_FieldTypeName = "SECTIONPAGES"
        Case wdFieldHyperlink: SKW_FieldTypeName = "HYPERLINK"
        Case wdFieldLink: SKW_FieldTypeName = "LINK"
        Case wdFieldIncludeText: SKW_FieldTypeName = "INCLUDETEXT"
        Case wdFieldIncludePicture: SKW_FieldTypeName = "INCLUDEPICTURE"
        Case wdFieldDocProperty: SKW_FieldTypeName = "DOCPROPERTY"
        Case wdFieldDate: SKW_FieldTypeName = "DATE"
        Case wdFieldMergeField: SKW_FieldTypeName = "MERGEFIELD"
        Case wdFieldAutoTextList: SKW_FieldTypeName = "AUTOTEXTLIST"
        Case wdFieldEmpty: SKW_FieldTypeName = "(empty)"
        Case Else: SKW_FieldTypeName = "Type " & t
    End Select
End Function

' The bookmark a REF, PAGEREF or NOTEREF field points to.
Public Function SKW_FieldBookmark(ByVal f As Object) As String
    Dim c As String, parts As Variant, i As Long
    c = Trim$(f.Code.Text)
    parts = Split(c, " ")
    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 Then SKW_FieldBookmark = Replace(parts(i), """", ""): Exit Function
    Next
End Function
