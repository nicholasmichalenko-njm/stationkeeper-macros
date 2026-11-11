Attribute VB_Name = "SK_Macros"
' =====================================================================
' Stationkeeper Macros 0.4.0
' A menu and an admin for the macros already in this template.
'
'   StationkeeperMacros        opens the menu (search, sort, read, run)
'   StationkeeperMacrosAdmin   opens the admin (describe, sort, show/hide)
'
' Nothing here changes your macros. The admin keeps a catalog (titles,
' descriptions, categories, keywords) inside this template as a hidden
' XML part, plus a dated copy beside the backups.
' Everything else in this module is a Function or Private, so Word's
' Macros list shows only the two entry macros above.
' =====================================================================
Option Explicit

Public Const SK_APP As String = "Stationkeeper Macros"
Public Const SK_VERSION As String = "0.4.0"
Public Const SK_NS As String = "urn:stationkeeper:macros:1"

Public SK_LastScanMethod As String
Public SK_LastModuleCount As Long
Public SK_ScanAvailable As Boolean
Public SK_Quiet As Boolean
Public SK_ForceNoAccess As Boolean   ' test switch: behave as if Word blocks the VBA project

' ---------------------------------------------------------------------
' Entry macros
' ---------------------------------------------------------------------
Public Sub StationkeeperMacros()
    On Error GoTo Fail
    frmSKMenu.Show vbModeless
    Exit Sub
Fail:
    MsgBox "The menu could not open." & vbCr & Err.Description, vbExclamation, SK_APP
End Sub

Public Sub StationkeeperMacrosAdmin()
    On Error GoTo Fail
    frmSKAdmin.Show vbModal
    Exit Sub
Fail:
    MsgBox "The admin could not open." & vbCr & Err.Description, vbExclamation, SK_APP
End Sub

' ---------------------------------------------------------------------
' Finding the macros in this template
' A macro is a Public Sub with no arguments in a standard module.
' Private subs, Functions, subs with arguments, Auto macros, modules
' marked Option Private Module and this kit itself are left out.
' Reading the template's own code needs "Trust access to the VBA project
' object model". Without it the admin can read exported .bas files, and
' the menu simply uses the saved catalog.
' ---------------------------------------------------------------------
Public Function SK_Scan() As Collection
    Dim res As New Collection
    Dim vbp As Object, comp As Object, cm As Object
    Dim ok As Boolean, n As Long

    On Error Resume Next
    Set vbp = ThisDocument.VBProject
    ok = (Err.Number = 0)
    If ok Then
        n = vbp.VBComponents.Count
        ok = (Err.Number = 0)
    End If
    Err.Clear
    On Error GoTo 0

    If SK_ForceNoAccess Then ok = False
    SK_ScanAvailable = ok
    If ok Then
        SK_LastScanMethod = "the VBA project"
        For Each comp In vbp.VBComponents
            If comp.Type = 1 Then
                If Not SK_IsKitName(comp.Name) Then
                    Set cm = comp.CodeModule
                    If cm.CountOfLines > 0 Then
                        SK_ScanLines comp.Name, Split(cm.Lines(1, cm.CountOfLines), vbCrLf), res
                    End If
                End If
            End If
        Next
    Else
        SK_LastScanMethod = "the saved catalog (VBA project access is off)"
    End If
    SK_LastModuleCount = SK_CountModules(res)
    Set SK_Scan = res
End Function

Public Function SK_CountModules(ByVal res As Collection) As Long
    Dim mods As Object, d As Object
    Set mods = CreateObject("Scripting.Dictionary")
    For Each d In res
        mods(LCase$(d("module"))) = True
    Next
    SK_CountModules = mods.Count
End Function

' Reads exported .bas files (VBA editor > File > Export File). One file or many.
Public Function SK_ReadBasFiles(ByVal paths As Collection) As Collection
    Dim res As New Collection, p As Variant, txt As String, modName As String, lines As Variant, i As Long
    For Each p In paths
        txt = Replace(SK_ReadUtf8(CStr(p)), vbCrLf, vbLf)
        lines = Split(txt, vbLf)
        modName = ""
        For i = 0 To UBound(lines)
            If LCase$(Left$(Trim$(lines(i)), 21)) = "attribute vb_name = """ Then
                modName = Mid$(Trim$(lines(i)), 22)
                modName = Left$(modName, Len(modName) - 1)
                Exit For
            End If
        Next
        If Len(modName) = 0 Then
            modName = Mid$(p, InStrRev(p, "\") + 1)
            If InStrRev(modName, ".") > 0 Then modName = Left$(modName, InStrRev(modName, ".") - 1)
        End If
        If Not SK_IsKitName(modName) Then SK_ScanLines modName, lines, res
    Next
    Set SK_ReadBasFiles = res
End Function

' The parser shared by the live scan and .bas files. lines is a 0-based array.
Private Sub SK_ScanLines(ByVal modName As String, ByVal lines As Variant, ByVal res As Collection)
    Dim i As Long, total As Long, startLine As Long, ln As String, nm As String, d As Object, t As String
    total = UBound(lines)
    For i = 0 To total
        t = LCase$(Trim$(lines(i)))
        If t Like "option private module*" Then Exit Sub
        If t Like "*sub *(*" Or t Like "*function *(*" Then Exit For
    Next
    i = 0
    Do While i <= total
        startLine = i
        ln = lines(i)
        Do While Right$(RTrim$(ln), 2) = " _" And i < total
            i = i + 1
            ln = Left$(RTrim$(ln), Len(RTrim$(ln)) - 1) & " " & Trim$(lines(i))
        Loop
        nm = SK_MacroNameFromLine(ln)
        If Len(nm) > 0 Then
            Set d = CreateObject("Scripting.Dictionary")
            d("run") = modName & "." & nm
            d("module") = modName
            d("name") = nm
            d("comment") = SK_CommentFor(lines, startLine, i, total)
            res.Add d
        End If
        i = i + 1
    Loop
End Sub

Private Function SK_MacroNameFromLine(ByVal ln As String) As String
    Dim s As String, low As String, nm As String
    Dim p As Long, q As Long
    s = Trim$(ln)
    low = LCase$(s)
    If Left$(low, 8) = "private " Or Left$(low, 7) = "friend " Then Exit Function
    If Left$(low, 7) = "public " Then
        s = Trim$(Mid$(s, 8)): low = LCase$(s)
    End If
    If Left$(low, 7) = "static " Then
        s = Trim$(Mid$(s, 8)): low = LCase$(s)
    End If
    If Left$(low, 4) <> "sub " Then Exit Function
    s = Trim$(Mid$(s, 5))
    p = InStr(s, "(")
    If p = 0 Then Exit Function
    nm = Trim$(Left$(s, p - 1))
    q = InStr(p, s, ")")
    If q = 0 Then Exit Function
    If Len(Trim$(Mid$(s, p + 1, q - p - 1))) > 0 Then Exit Function
    Select Case LCase$(nm)
        Case "autoexec", "autonew", "autoopen", "autoclose", "autoexit"
            Exit Function
    End Select
    If SK_IsKitName(nm) Then Exit Function
    SK_MacroNameFromLine = nm
End Function

' The comment lines right above a macro, or right below its Sub line.
Private Function SK_CommentFor(ByVal lines As Variant, ByVal startLine As Long, ByVal sigEnd As Long, _
                               ByVal total As Long) As String
    Dim j As Long, t As String, out As String, c As String
    j = startLine - 1
    Do While j >= 0
        t = Trim$(lines(j))
        If Left$(t, 1) <> "'" Then Exit Do
        c = SK_StripComment(t)
        If Len(c) > 0 Then out = c & IIf(Len(out) > 0, " ", "") & out
        j = j - 1
    Loop
    If Len(out) = 0 Then
        j = sigEnd + 1
        Do While j <= total
            t = Trim$(lines(j))
            If LCase$(Left$(t, 10)) <> "attribute " Then
                If Left$(t, 1) <> "'" Then Exit Do
                c = SK_StripComment(t)
                If Len(c) > 0 Then out = out & IIf(Len(out) > 0, " ", "") & c
            End If
            j = j + 1
        Loop
    End If
    If Len(out) > 600 Then out = Left$(out, 600)
    SK_CommentFor = Trim$(out)
End Function

Private Function SK_StripComment(ByVal t As String) As String
    Dim bare As String
    Do While Left$(t, 1) = "'"
        t = Mid$(t, 2)
    Loop
    t = Trim$(t)
    bare = Replace(Replace(Replace(Replace(Replace(t, "-", ""), "=", ""), "*", ""), "_", ""), " ", "")
    If Len(bare) = 0 Then t = ""
    SK_StripComment = t
End Function

' Adds macros read from .bas files to the catalog. Returns how many were new.
Public Function SK_MergeFound(ByVal m As Object, ByVal found As Collection) As Long
    Dim f As Object, e As Object, hit As Object, n As Long
    For Each f In found
        Set hit = Nothing
        For Each e In m("entries")
            If LCase$(e("run")) = LCase$(f("run")) Then Set hit = e: Exit For
        Next
        If hit Is Nothing Then
            Set hit = SK_NewEntry(f("run"))
            m("entries").Add hit
            n = n + 1
        End If
        hit("found") = True
        If Len(f("comment")) > 0 Then hit("comment") = f("comment")
    Next
    SK_MergeFound = n
End Function

Public Function SK_IsKitName(ByVal nm As String) As Boolean
    Dim l As String: l = LCase$(nm)
    SK_IsKitName = (l Like "sk_*" Or l Like "frmsk*" Or l Like "stationkeepermacros*" _
                    Or l Like "*.stationkeepermacros*")
End Function

Public Function SK_NamePart(ByVal run As String) As String
    If InStr(run, ".") > 0 Then
        SK_NamePart = Mid$(run, InStrRev(run, ".") + 1)
    Else
        SK_NamePart = run
    End If
End Function

Public Function SK_ModulePart(ByVal run As String) As String
    If InStr(run, ".") > 0 Then SK_ModulePart = Left$(run, InStrRev(run, ".") - 1)
End Function

' "FindBrokenRefs" -> "Find Broken Refs", "Update_TOC" -> "Update TOC"
Public Function SK_Friendly(ByVal nm As String) As String
    Dim i As Long, ch As String, prev As String, nxt As String, out As String
    nm = Replace(nm, "_", " ")
    For i = 1 To Len(nm)
        ch = Mid$(nm, i, 1)
        If i > 1 Then prev = Mid$(nm, i - 1, 1) Else prev = ""
        If i < Len(nm) Then nxt = Mid$(nm, i + 1, 1) Else nxt = ""
        If ch Like "[A-Z]" And Len(prev) > 0 Then
            If prev Like "[a-z0-9]" Then
                out = out & " "
            ElseIf prev Like "[A-Z]" And nxt Like "[a-z]" Then
                out = out & " "
            End If
        End If
        out = out & ch
    Next
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    SK_Friendly = Trim$(out)
End Function

' ---------------------------------------------------------------------
' The catalog: read, merge with what is really in the template, save
' ---------------------------------------------------------------------
Public Function SK_AffectsVocabulary() As Variant
    SK_AffectsVocabulary = Array("Tables", "Captions", "Cross-references", "Styles", "Headers and footers", _
        "Pages and sections", "Lists", "Quick Parts", "Fields", "Acronyms", "Text", "Hidden text", _
        "Figures", "Comments and tracking", "Whole document", "Template")
End Function

Public Function SK_NewEntry(ByVal run As String) As Object
    Dim e As Object: Set e = CreateObject("Scripting.Dictionary")
    e("run") = run
    e("title") = SK_Friendly(SK_NamePart(run))
    e("category") = ""
    e("description") = ""
    e("keywords") = ""
    e("affects") = ""
    e("workson") = ""
    e("changes") = ""
    e("also") = ""
    e("shown") = ""          ' "" = not decided yet, "1" = in the menu, "0" = hidden
    e("added") = Format$(Now, "yyyy-mm-dd")
    e("found") = False
    e("isNew") = True
    e("relinked") = ""
    e("comment") = ""
    e("issues") = ""
    Set SK_NewEntry = e
End Function

Public Function SK_CatalogXml() As String
    Dim p As Object, f As String
    On Error Resume Next
    For Each p In ThisDocument.CustomXMLParts
        If p.NamespaceURI = SK_NS Then
            SK_CatalogXml = p.XML
            Exit Function
        End If
    Next
    Err.Clear
    f = SK_SidecarPath()
    If Len(f) > 0 Then
        If Len(Dir$(f)) > 0 Then SK_CatalogXml = SK_ReadUtf8(f)
    End If
End Function

Public Function SK_LoadModel() As Object
    Dim m As Object, look As Object, e As Object, f As Object, hit As Object
    Dim cats As New Collection, entries As New Collection, hist As New Collection
    Dim xml As String, dom As Object, root As Object, nd As Object, k As Variant
    Dim found As Collection, cnt As Long

    Set m = CreateObject("Scripting.Dictionary")
    Set look = CreateObject("Scripting.Dictionary")
    look("menuTitle") = SK_APP
    look("defaultSort") = "Category"
    look("showUndescribed") = "0"
    look("keepOpen") = "1"
    look("confirmChanges") = "0"

    xml = SK_CatalogXml()
    If Len(xml) > 0 Then
        Set dom = CreateObject("MSXML2.DOMDocument.6.0")
        dom.async = False
        dom.setProperty "SelectionNamespaces", "xmlns:s='" & SK_NS & "'"
        If dom.LoadXML(xml) Then
            Set root = dom.DocumentElement
            For Each k In look.Keys
                If Not IsNull(root.getAttribute(k)) Then look(k) = CStr(root.getAttribute(k))
            Next
            For Each nd In dom.SelectNodes("/s:macros/s:categories/s:category")
                SK_AddUnique cats, nd.Text
            Next
            For Each nd In dom.SelectNodes("/s:macros/s:macro")
                entries.Add SK_EntryFromNode(nd)
            Next
            For Each nd In dom.SelectNodes("/s:macros/s:history/s:entry")
                hist.Add nd.Text
            Next
        End If
    End If

    ' Match the catalog with the macros really in the template.
    SK_SettleMarkers
    Set found = SK_Scan()
    If Not SK_ScanAvailable Then
        ' Cannot look inside the template: trust the catalog as saved.
        For Each e In entries
            e("found") = True
        Next
    End If
    For Each e In entries
        If Not SK_ScanAvailable Then GoTo NextEntry
        Set hit = Nothing
        For Each f In found
            If LCase$(f("run")) = LCase$(e("run")) Then
                Set hit = f
                Exit For
            End If
        Next
        If hit Is Nothing Then
            ' Same macro name in exactly one place: it moved modules. Follow it.
            cnt = 0
            For Each f In found
                If LCase$(f("name")) = LCase$(SK_NamePart(e("run"))) Then
                    cnt = cnt + 1
                    Set hit = f
                End If
            Next
            If cnt = 1 Then
                e("relinked") = e("run")
                e("run") = hit("run")
            Else
                Set hit = Nothing
            End If
        End If
        If Not hit Is Nothing Then
            e("found") = True
            e("comment") = hit("comment")
            hit("matched") = True
        End If
NextEntry:
    Next
    For Each f In found
        If Not f.Exists("matched") Then
            Set e = SK_NewEntry(f("run"))
            e("found") = True
            e("comment") = f("comment")
            entries.Add e
        End If
    Next
    SK_PrefillTools entries
    For Each e In entries
        If Len(e("category")) > 0 Then SK_AddUnique cats, e("category")
    Next

    Set m("look") = look
    Set m("cats") = cats
    Set m("entries") = entries
    Set m("history") = hist
    m("scanMethod") = SK_LastScanMethod
    m("foundCount") = found.Count
    m("moduleCount") = SK_LastModuleCount
    m("scanAvailable") = SK_ScanAvailable
    If Not SK_ScanAvailable Then
        m("foundCount") = entries.Count
        m("moduleCount") = SK_ModuleCountOfEntries(entries)
    End If
    Set SK_LoadModel = m
End Function

' Stationkeeper Tools describe themselves (SKW_Setup.SKW_ToolList). A tool new to the
' catalog arrives with its title, category and description, and in the menu.
' Without VBA project access, tools that answer are added as found.
Private Sub SK_PrefillTools(ByVal entries As Collection)
    Dim js As String, rows As Collection, r As Object, e As Object, hit As Object, f As Variant
    On Error Resume Next
    js = Application.Run("SKW_Setup.SKW_ToolList")
    If Err.Number <> 0 Or Len(js) = 0 Then Exit Sub
    Set rows = SK_RowsFromJsonText(js)
    If Err.Number <> 0 Or rows Is Nothing Then Exit Sub
    On Error GoTo 0
    For Each r In rows
        Set hit = Nothing
        For Each e In entries
            If LCase$(e("run")) = LCase$(r("run")) Then Set hit = e: Exit For
        Next
        If hit Is Nothing And Not SK_ScanAvailable Then
            Set hit = SK_NewEntry(r("run"))
            hit("found") = True
            entries.Add hit
        End If
        If Not hit Is Nothing Then
            If hit("isNew") Then
                For Each f In Array("title", "category", "description", "keywords", "affects", "workson", "changes")
                    If r.Exists(f) Then hit(f) = r(f)
                Next
                If r.Exists("shown") Then hit("shown") = SK_ShownFromText(r("shown"))
            End If
        End If
    Next
End Sub

Private Function SK_ModuleCountOfEntries(ByVal entries As Collection) As Long
    Dim mods As Object, e As Object
    Set mods = CreateObject("Scripting.Dictionary")
    For Each e In entries
        mods(LCase$(SK_ModulePart(e("run")))) = True
    Next
    SK_ModuleCountOfEntries = mods.Count
End Function

Private Function SK_EntryFromNode(ByVal nd As Object) As Object
    Dim e As Object
    Set e = SK_NewEntry(SK_AttrText(nd, "run"))
    e("isNew") = False
    e("shown") = SK_AttrText(nd, "shown")
    e("title") = SK_ChildText(nd, "title")
    If Len(e("title")) = 0 Then e("title") = SK_Friendly(SK_NamePart(e("run")))
    e("category") = SK_ChildText(nd, "category")
    e("description") = SK_ChildText(nd, "description")
    e("keywords") = SK_ChildText(nd, "keywords")
    e("affects") = SK_ChildText(nd, "affects")
    e("workson") = SK_ChildText(nd, "worksOn")
    e("changes") = SK_ChildText(nd, "changes")
    e("also") = SK_ChildText(nd, "also")
    e("added") = SK_ChildText(nd, "added")
    Set SK_EntryFromNode = e
End Function

Private Function SK_AttrText(ByVal nd As Object, ByVal nm As String) As String
    Dim v As Variant: v = nd.getAttribute(nm)
    If Not IsNull(v) Then SK_AttrText = CStr(v)
End Function

Private Function SK_ChildText(ByVal nd As Object, ByVal nm As String) As String
    Dim c As Object, t As String
    Set c = nd.selectSingleNode("s:" & nm)
    If Not c Is Nothing Then
        t = Replace(c.Text, vbCrLf, vbLf)
        SK_ChildText = Replace(t, vbLf, vbCrLf)
    End If
End Function

Public Sub SK_AddUnique(ByVal col As Collection, ByVal s As String)
    Dim v As Variant
    s = Trim$(s)
    If Len(s) = 0 Then Exit Sub
    For Each v In col
        If LCase$(v) = LCase$(s) Then Exit Sub
    Next
    col.Add s
End Sub

Public Function SK_EntryStatus(ByVal e As Object) As String
    If Not e("found") Then
        SK_EntryStatus = "Not found"
    ElseIf e("shown") = "0" Then
        SK_EntryStatus = "Hidden"
    ElseIf Len(Trim$(e("description"))) = 0 Then
        If e("shown") = "1" Then SK_EntryStatus = "In menu, no description" Else SK_EntryStatus = "Not described"
    ElseIf e("shown") = "1" Then
        SK_EntryStatus = "In the menu"
    Else
        SK_EntryStatus = "Described, not shown"
    End If
End Function

Public Function SK_BuildXml(ByVal m As Object) As String
    Dim s As String, k As Variant, v As Variant, e As Object, look As Object
    Set look = m("look")
    s = "<macros xmlns=""" & SK_NS & """ version=""1"" kit=""" & SK_VERSION & """"
    For Each k In look.Keys
        s = s & " " & k & "=""" & SK_Esc(look(k)) & """"
    Next
    s = s & "><categories>"
    For Each v In m("cats")
        s = s & "<category>" & SK_Esc(v) & "</category>"
    Next
    s = s & "</categories>"
    For Each e In m("entries")
        s = s & "<macro run=""" & SK_Esc(e("run")) & """ shown=""" & SK_Esc(e("shown")) & """>" & _
            SK_El("title", e("title")) & SK_El("category", e("category")) & _
            SK_El("description", e("description")) & SK_El("keywords", e("keywords")) & _
            SK_El("affects", e("affects")) & SK_El("worksOn", e("workson")) & _
            SK_El("changes", e("changes")) & SK_El("also", e("also")) & _
            SK_El("added", e("added")) & "</macro>"
    Next
    s = s & "<history>"
    For Each v In m("history")
        s = s & "<entry>" & SK_Esc(v) & "</entry>"
    Next
    s = s & "</history></macros>"
    SK_BuildXml = s
End Function

Private Function SK_El(ByVal nm As String, ByVal v As String) As String
    SK_El = "<" & nm & ">" & SK_Esc(v) & "</" & nm & ">"
End Function

Public Function SK_Esc(ByVal v As Variant) As String
    Dim s As String, i As Long, ch As String, code As Long, out As String
    s = CStr(v)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case ch
            Case "&": out = out & "&amp;"
            Case "<": out = out & "&lt;"
            Case ">": out = out & "&gt;"
            Case """": out = out & "&quot;"
            Case Else
                If code >= 32 Or code = 9 Or code = 10 Or code = 13 Or code < 0 Then out = out & ch
        End Select
    Next
    SK_Esc = out
End Function

' Saves the catalog into this template. Returns a message for the user.
Public Function SK_SaveModel(ByVal m As Object, ByVal summary As String) As String
    Dim msg As String, bdir As String, dst As String, i As Long, p As Object
    Dim xml As String

    ' 1. Backup of the template as it is on disk.
    bdir = SK_BackupDir()
    If Len(bdir) > 0 And Len(ThisDocument.Path) > 0 Then
        dst = bdir & "\" & SK_BaseName() & " - backup " & Format$(Now, "yyyy-mm-dd hhnnss") & SK_Ext()
        On Error Resume Next
        CreateObject("Scripting.FileSystemObject").CopyFile ThisDocument.FullName, dst, True
        If Err.Number <> 0 Then
            msg = msg & "No backup made (" & Err.Description & ")." & vbCr
            dst = ""
        End If
        Err.Clear
        On Error GoTo 0
    End If

    ' 2. History line.
    m("history").Add Format$(Now, "yyyy-mm-dd hh:nn") & " | " & Application.UserName & " | " & summary

    ' 3. Replace the catalog part.
    xml = SK_BuildXml(m)
    For i = ThisDocument.CustomXMLParts.Count To 1 Step -1
        Set p = ThisDocument.CustomXMLParts(i)
        If p.NamespaceURI = SK_NS Then p.Delete
    Next
    ThisDocument.CustomXMLParts.Add xml

    ' 4. Save the template.
    On Error GoTo SaveFail
    SK_SaveTemplate
    On Error GoTo 0

    ' 5. Readable copies beside the backups.
    If Len(bdir) > 0 Then
        On Error Resume Next
        SK_WriteUtf8 SK_SidecarPath(), xml
        Err.Clear
        On Error GoTo 0
    End If

    If Len(bdir) = 0 Then msg = msg & "No backup made: no backup folder is chosen (Settings tab)." & vbCr
    msg = msg & "Saved to " & ThisDocument.Name & "."
    If Len(dst) > 0 Then msg = msg & vbCr & "Backup: " & dst
    SK_SaveModel = msg
    Exit Function
SaveFail:
    SK_SaveModel = msg & "The catalog changed but the template did not save: " & Err.Description & _
        vbCr & "Open the template itself (File > Open) and save it."
End Function

Public Sub SK_SaveTemplate()
    Dim d As Object, isOpen As Boolean
    For Each d In Documents
        If LCase$(d.FullName) = LCase$(ThisDocument.FullName) Then isOpen = True
    Next
    If isOpen Then
        ThisDocument.Save
    Else
        MacroContainer.Saved = False
        MacroContainer.Save
    End If
End Sub

Public Function SK_BaseName() As String
    Dim n As String: n = ThisDocument.Name
    If InStrRev(n, ".") > 0 Then SK_BaseName = Left$(n, InStrRev(n, ".") - 1) Else SK_BaseName = n
End Function

Private Function SK_Ext() As String
    Dim n As String: n = ThisDocument.Name
    If InStrRev(n, ".") > 0 Then SK_Ext = Mid$(n, InStrRev(n, ".")) Else SK_Ext = ".dotm"
End Function

' ---------------------------------------------------------------------
' Backups folder. Each person chooses it once (the admin asks), and it is
' kept in their Windows profile for every template. "" = not chosen yet.
' ---------------------------------------------------------------------
Public Function SK_BackupDir() As String
    Dim d As String
    d = GetSetting(SK_APP, "Settings", "backupDir", "")
    If Len(d) > 0 Then
        On Error Resume Next
        If Len(Dir$(d, vbDirectory)) = 0 Then d = ""
        If Err.Number <> 0 Then d = ""
        Err.Clear
    End If
    SK_BackupDir = d
End Function

Public Function SK_SetBackupDir(ByVal d As String) As Boolean
    On Error Resume Next
    If Len(d) > 0 Then
        If Len(Dir$(d, vbDirectory)) = 0 Then MkDir d
    End If
    SaveSetting SK_APP, "Settings", "backupDir", d
    SK_SetBackupDir = (Err.Number = 0)
End Function

' The folder suggested the first time: beside the template, else Documents.
Public Function SK_SuggestedBackupDir() As String
    Dim p As String
    p = ThisDocument.Path
    If Len(p) = 0 Or LCase$(Left$(p, 4)) = "http" Then p = Environ$("USERPROFILE") & "\Documents"
    SK_SuggestedBackupDir = p & "\Stationkeeper Macros backups"
End Function

Public Function SK_PickFolder(ByVal title As String, ByVal initial As String) As String
    With Application.FileDialog(4)
        .Title = title
        If Len(initial) > 0 Then .InitialFileName = initial & "\"
        If .Show = -1 Then SK_PickFolder = .SelectedItems(1)
    End With
End Function

Public Function SK_SidecarPath() As String
    Dim d As String
    On Error Resume Next
    d = SK_BackupDir()
    If Len(d) > 0 Then SK_SidecarPath = d & "\" & SK_BaseName() & " - macro catalog.xml"
End Function

Public Sub SK_WriteUtf8(ByVal path As String, ByVal text As String)
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.SaveToFile path, 2
    st.Close
End Sub

Public Function SK_ReadUtf8(ByVal path As String) As String
    Dim st As Object, s As String
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    s = st.ReadText(-1)
    st.Close
    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = &HFEFF Then s = Mid$(s, 2)
    End If
    SK_ReadUtf8 = s
End Function

' ---------------------------------------------------------------------
' Check, export and import
' ---------------------------------------------------------------------
Public Function SK_NoAccessMessage() As String
    SK_NoAccessMessage = "Word is not letting the admin look inside this template, so the list shows only what was saved before." & _
        vbCr & vbCr & "To list every macro, do one of these:" & vbCr & _
        "1. Turn on File > Options > Trust Center > Trust Center Settings > Macro Settings > " & _
        """Trust access to the VBA project object model"", then open the admin again. " & _
        "Only the person running the admin needs this. People using the menu do not." & vbCr & _
        "2. Export your module(s) from the VBA editor (right-click > Export File), then use " & _
        "Bulk tools > Read macros from .bas files."
End Function

Private Function SK_CsvQ(ByVal v As Variant) As String
    SK_CsvQ = """" & Replace(CStr(v), """", """""") & """"
End Function

Private Function SK_YesNo(ByVal shown As String) As String
    Select Case shown
        Case "1": SK_YesNo = "Yes"
        Case "0": SK_YesNo = "No"
        Case Else: SK_YesNo = ""
    End Select
End Function

Public Function SK_ExportCsv(ByVal m As Object) As String
    Dim s As String, e As Object, path As String
    s = "Macro,Title,Category,Description,Keywords,What it affects,Works on,Does it change the document,Show in menu,Status" & vbCrLf
    For Each e In m("entries")
        s = s & SK_CsvQ(e("run")) & "," & SK_CsvQ(e("title")) & "," & SK_CsvQ(e("category")) & "," & _
            SK_CsvQ(e("description")) & "," & SK_CsvQ(e("keywords")) & "," & SK_CsvQ(e("affects")) & "," & _
            SK_CsvQ(e("workson")) & "," & SK_CsvQ(e("changes")) & "," & _
            SK_CsvQ(SK_YesNo(e("shown"))) & "," & SK_CsvQ(SK_EntryStatus(e)) & vbCrLf
    Next
    If Len(SK_BackupDir()) = 0 Then Err.Raise vbObjectError + 1, SK_APP, "Choose a backup folder first (Settings tab)."
    path = SK_BackupDir() & "\" & SK_BaseName() & " - macro catalog " & Format$(Now, "yyyy-mm-dd hhnn") & ".csv"
    SK_WriteUtf8 path, s
    SK_ExportCsv = path
End Function

Public Function SK_ParseCsv(ByVal text As String) As Collection
    Dim rows As New Collection, row As Collection, fld As String
    Dim i As Long, n As Long, ch As String, inQ As Boolean
    Set row = New Collection
    n = Len(text)
    i = 1
    Do While i <= n
        ch = Mid$(text, i, 1)
        If inQ Then
            If ch = """" Then
                If Mid$(text, i + 1, 1) = """" Then
                    fld = fld & """"
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                fld = fld & ch
            End If
        Else
            Select Case ch
                Case """"
                    inQ = True
                Case ","
                    row.Add fld: fld = ""
                Case vbCr
                Case vbLf
                    row.Add fld: fld = ""
                    rows.Add row
                    Set row = New Collection
                Case Else
                    fld = fld & ch
            End Select
        End If
        i = i + 1
    Loop
    If Len(fld) > 0 Or row.Count > 0 Then
        row.Add fld
        rows.Add row
    End If
    Set SK_ParseCsv = rows
End Function

' ---------------------------------------------------------------------
' Lists: update titles, categories, descriptions and the rest from a file.
' Excel (.xlsx/.xlsm/.xls), CSV, a table in a Word file, or JSON.
' Each row needs a Macro column (Module.Macro, or just the macro name).
' Blank cells leave values as they are. Rows for macros that are not in
' the template are skipped and listed. Code is never changed.
' ---------------------------------------------------------------------
Public Function SK_FieldKey(ByVal h As String) As String
    h = LCase$(Trim$(Replace(Replace(Replace(h, "?", ""), ":", ""), "_", " ")))
    Select Case h
        Case "macro", "macro name", "run": SK_FieldKey = "run"
        Case "title", "title in the menu", "name": SK_FieldKey = "title"
        Case "category": SK_FieldKey = "category"
        Case "description": SK_FieldKey = "description"
        Case "keywords", "search keywords": SK_FieldKey = "keywords"
        Case "what it affects", "what this macro affects", "affects": SK_FieldKey = "affects"
        Case "works on", "workson": SK_FieldKey = "workson"
        Case "does it change the document", "does this change the document", "changes the document", "changes"
            SK_FieldKey = "changes"
        Case "show in menu", "show in the menu", "showinmenu", "in menu", "shown": SK_FieldKey = "shown"
    End Select
End Function

Public Function SK_ReadListFile(ByVal path As String) As Collection
    Select Case LCase$(Mid$(path, InStrRev(path, ".") + 1))
        Case "csv", "txt": Set SK_ReadListFile = SK_RowsFromCells(SK_ParseCsv(SK_ReadUtf8(path)))
        Case "xlsx", "xlsm", "xls": Set SK_ReadListFile = SK_RowsFromCells(SK_CellsFromExcel(path))
        Case "docx", "docm", "doc": Set SK_ReadListFile = SK_RowsFromCells(SK_CellsFromWord(path))
        Case "json": Set SK_ReadListFile = SK_RowsFromJson(path)
        Case Else: Err.Raise vbObjectError + 2, SK_APP, "Use an Excel, CSV, Word or JSON file."
    End Select
End Function

' Rows of cells (a Collection of Collections) to rows of fields. The header
' row is the first row with a Macro column; rows above it are ignored.
Private Function SK_RowsFromCells(ByVal cells As Collection) As Collection
    Dim res As New Collection, row As Collection, cols As Object, d As Object, i As Long, r As Long, hdr As Long
    Dim k As Variant, v As String
    For r = 1 To cells.Count
        Set row = cells(r)
        For i = 1 To row.Count
            If SK_FieldKey(row(i)) = "run" Then hdr = r: Exit For
        Next
        If hdr > 0 Then Exit For
    Next
    If hdr = 0 Then Err.Raise vbObjectError + 3, SK_APP, "No Macro column found. The first row of the list needs a column named Macro."
    Set cols = CreateObject("Scripting.Dictionary")
    Set row = cells(hdr)
    For i = 1 To row.Count
        If Len(SK_FieldKey(row(i))) > 0 Then
            If Not cols.Exists(SK_FieldKey(row(i))) Then cols(SK_FieldKey(row(i))) = i
        End If
    Next
    For r = hdr + 1 To cells.Count
        Set row = cells(r)
        Set d = CreateObject("Scripting.Dictionary")
        For Each k In cols.Keys
            If cols(k) <= row.Count Then
                v = Replace(Replace(row(cols(k)), vbCrLf, vbLf), vbCr, vbLf)
                d(k) = Trim$(Replace(v, vbLf, vbCrLf))
            End If
        Next
        If d.Exists("run") Then
            If Len(d("run")) > 0 Then res.Add d
        End If
    Next
    Set SK_RowsFromCells = res
End Function

Private Function SK_CellsFromExcel(ByVal path As String) As Collection
    Dim xl As Object, wb As Object, ws As Object, v As Variant, r As Long, c As Long
    Dim res As New Collection, row As Collection, en As Long, ed As String
    On Error GoTo Fail
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(FileName:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    On Error Resume Next
    Set ws = wb.Worksheets("Macros")
    On Error GoTo Fail
    If ws Is Nothing Then Set ws = wb.Worksheets(1)
    v = ws.UsedRange.Value
    wb.Close False
    xl.Quit
    Set xl = Nothing
    If IsArray(v) Then
        For r = LBound(v, 1) To UBound(v, 1)
            Set row = New Collection
            For c = LBound(v, 2) To UBound(v, 2)
                If IsError(v(r, c)) Or IsEmpty(v(r, c)) Then row.Add "" Else row.Add CStr(v(r, c))
            Next
            res.Add row
        Next
    End If
    Set SK_CellsFromExcel = res
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Err.Raise vbObjectError + 4, SK_APP, "Could not read the Excel file: " & ed
End Function

Private Function SK_CellsFromWord(ByVal path As String) As Collection
    Dim d As Object, t As Object, tbl As Object, res As New Collection, row As Collection
    Dim r As Long, c As Long, txt As String, ed As String
    On Error GoTo Fail
    Set d = Documents.Open(FileName:=path, ReadOnly:=True, AddToRecentFiles:=False, Visible:=False)
    For Each t In d.Tables
        For c = 1 To t.Rows(1).Cells.Count
            If SK_FieldKey(SK_CellText(t.Rows(1).Cells(c))) = "run" Then Set tbl = t: Exit For
        Next
        If Not tbl Is Nothing Then Exit For
    Next
    If Not tbl Is Nothing Then
        For r = 1 To tbl.Rows.Count
            Set row = New Collection
            For c = 1 To tbl.Rows(r).Cells.Count
                row.Add SK_CellText(tbl.Rows(r).Cells(c))
            Next
            res.Add row
        Next
    End If
    d.Close SaveChanges:=0
    Set SK_CellsFromWord = res
    Exit Function
Fail:
    ed = Err.Description
    On Error Resume Next
    If Not d Is Nothing Then d.Close SaveChanges:=0
    Err.Raise vbObjectError + 5, SK_APP, "Could not read the table in the Word file (use a simple table, no merged cells): " & ed
End Function

Private Function SK_CellText(ByVal cel As Object) As String
    Dim t As String
    t = cel.Range.Text
    If Len(t) >= 2 Then t = Left$(t, Len(t) - 2)
    SK_CellText = Replace(t, vbCr, vbCrLf)
End Function

' JSON: an array of objects, or an object with a "macros" array.
Private Function SK_RowsFromJson(ByVal path As String) As Collection
    Set SK_RowsFromJson = SK_RowsFromJsonText(SK_ReadUtf8(path))
End Function

Public Function SK_RowsFromJsonText(ByVal s As String) As Collection
    Dim pos As Long, v As Variant, arr As Object, item As Variant, d As Object, k As Variant, key As String
    Dim res As New Collection
    pos = 1
    SK_JsonAssign v, s, pos
    If Not IsObject(v) Then Err.Raise vbObjectError + 6, SK_APP, "The JSON file needs a list of macros."
    If TypeName(v) = "Dictionary" Then
        If v.Exists("macros") Then Set arr = v("macros") Else Err.Raise vbObjectError + 6, SK_APP, "The JSON file needs a list of macros."
    Else
        Set arr = v
    End If
    For Each item In arr
        If IsObject(item) Then
            Set d = CreateObject("Scripting.Dictionary")
            For Each k In item.Keys
                key = SK_FieldKey(k)
                If Len(key) > 0 And Not IsObject(item(k)) Then
                    If IsNull(item(k)) Then
                        d(key) = ""
                    ElseIf VarType(item(k)) = vbBoolean Then
                        d(key) = IIf(item(k), "Yes", "No")
                    Else
                        d(key) = Trim$(CStr(item(k)))
                    End If
                End If
            Next
            If d.Exists("run") Then
                If Len(d("run")) > 0 Then res.Add d
            End If
        End If
    Next
    Set SK_RowsFromJsonText = res
End Function

Private Sub SK_JsonSkip(ByVal s As String, pos As Long)
    Do While pos <= Len(s)
        Select Case Mid$(s, pos, 1)
            Case " ", vbTab, vbCr, vbLf, ChrW$(&HFEFF): pos = pos + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Sub SK_JsonAssign(ByRef v As Variant, ByVal s As String, pos As Long)
    Dim num As String
    SK_JsonSkip s, pos
    If pos > Len(s) Then Err.Raise vbObjectError + 7, SK_APP, "The JSON file ends too early."
    Select Case Mid$(s, pos, 1)
        Case "{": Set v = SK_JsonObject(s, pos)
        Case "[": Set v = SK_JsonArray(s, pos)
        Case """": v = SK_JsonString(s, pos)
        Case "t": v = True: pos = pos + 4
        Case "f": v = False: pos = pos + 5
        Case "n": v = Null: pos = pos + 4
        Case Else
            Do While pos <= Len(s) And InStr("+-0123456789.eE", Mid$(s, pos, 1)) > 0
                num = num & Mid$(s, pos, 1)
                pos = pos + 1
            Loop
            If Len(num) = 0 Then Err.Raise vbObjectError + 7, SK_APP, "The JSON file has something unexpected at character " & pos & "."
            v = num
    End Select
End Sub

Private Function SK_JsonObject(ByVal s As String, pos As Long) As Object
    Dim d As Object, k As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    pos = pos + 1
    Do
        SK_JsonSkip s, pos
        If pos > Len(s) Then Err.Raise vbObjectError + 7, SK_APP, "The JSON file ends too early."
        If Mid$(s, pos, 1) = "}" Then pos = pos + 1: Exit Do
        If Mid$(s, pos, 1) = "," Then pos = pos + 1: SK_JsonSkip s, pos
        k = SK_JsonString(s, pos)
        SK_JsonSkip s, pos
        If Mid$(s, pos, 1) <> ":" Then Err.Raise vbObjectError + 7, SK_APP, "The JSON file is missing a colon at character " & pos & "."
        pos = pos + 1
        SK_JsonAssign v, s, pos
        If IsObject(v) Then Set d(k) = v Else d(k) = v
    Loop
    Set SK_JsonObject = d
End Function

Private Function SK_JsonArray(ByVal s As String, pos As Long) As Collection
    Dim c As New Collection, v As Variant
    pos = pos + 1
    Do
        SK_JsonSkip s, pos
        If pos > Len(s) Then Err.Raise vbObjectError + 7, SK_APP, "The JSON file ends too early."
        If Mid$(s, pos, 1) = "]" Then pos = pos + 1: Exit Do
        If Mid$(s, pos, 1) = "," Then pos = pos + 1: SK_JsonSkip s, pos
        SK_JsonAssign v, s, pos
        If IsObject(v) Then c.Add v Else c.Add v
    Loop
    Set SK_JsonArray = c
End Function

Private Function SK_JsonString(ByVal s As String, pos As Long) As String
    Dim out As String, ch As String
    If Mid$(s, pos, 1) <> """" Then Err.Raise vbObjectError + 7, SK_APP, "The JSON file expected text at character " & pos & "."
    pos = pos + 1
    Do While pos <= Len(s)
        ch = Mid$(s, pos, 1)
        If ch = """" Then
            pos = pos + 1
            SK_JsonString = out
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
    Err.Raise vbObjectError + 7, SK_APP, "The JSON file has text that never ends."
End Function

Public Function SK_FindEntry(ByVal m As Object, ByVal key As String) As Object
    Dim e As Object, hit As Object, n As Long
    For Each e In m("entries")
        If LCase$(e("run")) = LCase$(key) Then Set SK_FindEntry = e: Exit Function
    Next
    For Each e In m("entries")
        If LCase$(SK_NamePart(e("run"))) = LCase$(SK_NamePart(key)) Then n = n + 1: Set hit = e
    Next
    If n = 1 Then Set SK_FindEntry = hit
End Function

Private Function SK_ShownFromText(ByVal v As String) As String
    Select Case LCase$(Trim$(v))
        Case "yes", "y", "1", "true", "shown", "show": SK_ShownFromText = "1"
        Case "no", "n", "0", "false", "hidden", "hide": SK_ShownFromText = "0"
    End Select
End Function

' Compares a list with the catalog. With apply=False it changes nothing and
' only describes what would change. nChanged returns how many macros change.
Public Function SK_UpdateFromRows(ByVal m As Object, ByVal rows As Collection, ByVal apply As Boolean, _
                                  ByRef nChanged As Long) As String
    Dim r As Object, e As Object, f As Variant, v As String, changed As Boolean, i As Long
    Dim nSame As Long, nMissing As Long, missing As String, notOnList As Long
    Dim per As Object, onList As Object, s As String, names As Variant, fields As Variant
    fields = Array("title", "category", "description", "keywords", "affects", "workson", "changes", "shown")
    names = Array("titles", "categories", "descriptions", "keywords", "what it affects", "works on", _
                  "does it change the document", "show in menu")
    Set per = CreateObject("Scripting.Dictionary")
    Set onList = CreateObject("Scripting.Dictionary")
    nChanged = 0
    For Each r In rows
        Set e = SK_FindEntry(m, r("run"))
        If e Is Nothing Then
            nMissing = nMissing + 1
            If nMissing <= 8 Then missing = missing & IIf(Len(missing) > 0, ", ", "") & r("run")
        Else
            onList(LCase$(e("run"))) = True
            changed = False
            For Each f In fields
                If r.Exists(f) Then
                    v = r(f)
                    If f = "shown" Then v = SK_ShownFromText(v)
                    If Len(v) > 0 And v <> e(f) Then
                        per(f) = per(f) + 1
                        changed = True
                        If apply Then
                            e(f) = v
                            If f = "category" Then SK_AddUnique m("cats"), v
                        End If
                    End If
                End If
            Next
            If changed Then nChanged = nChanged + 1 Else nSame = nSame + 1
        End If
    Next
    For Each e In m("entries")
        If e("found") And Not onList.Exists(LCase$(e("run"))) Then notOnList = notOnList + 1
    Next
    s = rows.Count & " row(s) read." & vbCr
    If nChanged > 0 Then
        s = s & nChanged & " macro(s) " & IIf(apply, "changed", "will change") & ":"
        For i = 0 To UBound(fields)
            If per.Exists(fields(i)) Then s = s & " " & names(i) & " " & per(fields(i)) & ","
        Next
        s = Left$(s, Len(s) - 1) & "." & vbCr
    End If
    s = s & nSame & " macro(s) already match." & vbCr
    If nMissing > 0 Then s = s & nMissing & " row(s) are for macros not in this template, skipped: " & missing & _
        IIf(nMissing > 8, ", ...", "") & "." & vbCr
    If notOnList > 0 Then s = s & notOnList & " macro(s) in this template are not on the list, left as they are." & vbCr
    s = s & "Blank cells leave values as they are."
    SK_UpdateFromRows = s
End Function

' Primary list for this template (per person, kept in their Windows profile).
Public Function SK_GetPrimaryList() As String
    SK_GetPrimaryList = GetSetting(SK_APP, SK_UserKey(), "pref|primaryList", "")
End Function

Public Function SK_SetPrimaryList(ByVal path As String) As Boolean
    SaveSetting SK_APP, SK_UserKey(), "pref|primaryList", path
    SK_SetPrimaryList = True
End Function

' Excel export: one row per macro, sheet "Macros", the same columns a list uses.
Public Function SK_ExportXlsx(ByVal m As Object) As String
    Dim xl As Object, wb As Object, ws As Object, e As Object, path As String, hdr As Variant
    Dim data() As Variant, r As Long, c As Long, n As Long, ed As String, widths As Variant
    If Len(SK_BackupDir()) = 0 Then Err.Raise vbObjectError + 1, SK_APP, "Choose a backup folder first (Settings tab)."
    path = SK_BackupDir() & "\" & SK_BaseName() & " - macro list " & Format$(Now, "yyyy-mm-dd hhnn") & ".xlsx"
    hdr = Array("Macro", "Title", "Category", "Description", "Keywords", "What it affects", "Works on", _
                "Does it change the document", "Show in menu", "Status", "Test")
    widths = Array(30, 28, 18, 60, 24, 24, 18, 22, 12, 22, 12)
    n = m("entries").Count
    ReDim data(1 To n + 1, 1 To 11)
    For c = 0 To 10: data(1, c + 1) = hdr(c): Next
    r = 1
    For Each e In m("entries")
        r = r + 1
        data(r, 1) = e("run"): data(r, 2) = e("title"): data(r, 3) = e("category")
        data(r, 4) = Replace(e("description"), vbCrLf, vbLf): data(r, 5) = e("keywords"): data(r, 6) = e("affects")
        data(r, 7) = e("workson"): data(r, 8) = e("changes"): data(r, 9) = SK_YesNo(e("shown"))
        data(r, 10) = SK_EntryStatus(e): data(r, 11) = SK_TestState(e("run"))
    Next
    On Error GoTo Fail
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Macros"
    ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, 11)).NumberFormat = "@"
    ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, 11)).Value = data
    ws.Rows(1).Font.Bold = True
    For c = 0 To 10: ws.Columns(c + 1).ColumnWidth = widths(c): Next
    ws.Columns(4).WrapText = True
    ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, 11)).AutoFilter
    wb.SaveAs FileName:=path, FileFormat:=51
    wb.Close False
    xl.Quit
    SK_ExportXlsx = path
    Exit Function
Fail:
    ed = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Err.Raise vbObjectError + 8, SK_APP, "Could not make the Excel file: " & ed
End Function

' ---------------------------------------------------------------------
' Running a macro, and what each person sees (kept in their Windows
' profile, never in the template)
' ---------------------------------------------------------------------
' Runs a macro as one undo step ("Stationkeeper: <title>" on the Undo list).
' A marker is set before and cleared after. If Word stops the macro with an
' error dialog and the person clicks End, the marker stays, and the next time
' the menu or admin opens, the macro is recorded as Failed on this PC.
Public Function SK_RunMacro(ByVal run As String, Optional ByVal title As String = "") As Boolean
    Dim rec As Boolean, t As Single
    On Error GoTo Fail
    SaveSetting SK_APP, SK_UserKey(), "uses|" & run, CStr(SK_Uses(run) + 1)
    SaveSetting SK_APP, SK_UserKey(), "last|" & run, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    SaveSetting SK_APP, SK_UserKey(), "mark|" & run, "menu|" & Format$(Now, "yyyy-mm-dd hh:nn")
    If Len(title) = 0 Then title = SK_Friendly(SK_NamePart(run))
    rec = SK_StartUndo(title)
    t = Timer
    Application.Run run
    If rec Then Application.UndoRecord.EndCustomRecord
    SaveSetting SK_APP, SK_UserKey(), "mark|" & run, ""
    SK_SetResult run, "Passed", Format$(Timer - t, "0.00") & " s, run from the menu"
    SK_RunMacro = True
    Exit Function
Fail:
    If rec Then
        On Error Resume Next
        Application.UndoRecord.EndCustomRecord
    End If
    SaveSetting SK_APP, SK_UserKey(), "mark|" & run, ""
    If Err.Number = -2147352573 Or InStr(Err.Description, "run the specified macro") > 0 Then
        SK_SetResult run, "Failed", "the macro is not in the template"
    End If
    If Not SK_Quiet Then MsgBox run & " could not run:" & vbCr & Err.Description, vbExclamation, SK_APP
End Function

Private Function SK_StartUndo(ByVal title As String) As Boolean
    On Error Resume Next
    If Application.UndoRecord.IsRecordingCustomRecord Then Exit Function
    Application.UndoRecord.StartCustomRecord Left$("Stationkeeper: " & title, 64)
    SK_StartUndo = (Err.Number = 0)
End Function

' ---------------------------------------------------------------------
' Test results (per PC): Passed or Failed, with date and detail
' ---------------------------------------------------------------------
Public Function SK_SetResult(ByVal run As String, ByVal state As String, ByVal detail As String) As Boolean
    SaveSetting SK_APP, SK_UserKey(), "result|" & run, state & "|" & Format$(Now, "yyyy-mm-dd hh:nn") & "|" & detail
    SK_SetResult = True
End Function

Public Function SK_TestState(ByVal run As String) As String
    Dim v As String
    v = GetSetting(SK_APP, SK_UserKey(), "result|" & run, "")
    If Len(v) = 0 Then SK_TestState = "Not tested" Else SK_TestState = Split(v, "|")(0)
End Function

Public Function SK_TestDetail(ByVal run As String) As String
    Dim v As Variant
    v = Split(GetSetting(SK_APP, SK_UserKey(), "result|" & run, ""), "|")
    If UBound(v) >= 2 Then SK_TestDetail = v(0) & " " & v(1) & ": " & v(2)
End Function

' Turns markers left by runs that never finished into Failed results.
Public Function SK_SettleMarkers() As Long
    Dim a As Variant, i As Long, k As String, v As String
    On Error Resume Next
    a = GetAllSettings(SK_APP, SK_UserKey())
    If IsEmpty(a) Then Exit Function
    For i = LBound(a, 1) To UBound(a, 1)
        k = a(i, 0): v = a(i, 1)
        If Left$(k, 5) = "mark|" And Len(v) > 0 Then
            SK_SetResult Mid$(k, 6), "Failed", IIf(Left$(v, 4) = "test", "stopped with an error during a test", _
                                                   "stopped with an error when run from the menu")
            SaveSetting SK_APP, SK_UserKey(), k, ""
            SK_SettleMarkers = SK_SettleMarkers + 1
        End If
    Next
End Function

' ---------------------------------------------------------------------
' What a macro needs before it runs (from its "Works on" answer)
' ---------------------------------------------------------------------
Public Function SK_NeedsMet(ByVal worksOn As String, ByRef why As String) As Boolean
    why = ""
    On Error Resume Next
    If Documents.Count = 0 Then
        why = "Open a document first."
        Exit Function
    End If
    Select Case LCase$(worksOn)
        Case "selection"
            If Selection.Type = wdSelectionIP Or Len(Selection.Range.Text) = 0 Then why = "Select some text first."
        Case "table at the cursor"
            If Not Selection.Information(wdWithInTable) Then why = "Put the cursor in a table first."
        Case "list at the cursor"
            If Selection.Range.ListFormat.ListType = wdListNoNumbering Then why = "Put the cursor in a list first."
    End Select
    SK_NeedsMet = (Len(why) = 0)
End Function

' ---------------------------------------------------------------------
' Test runs: each test uses a fresh scratch document, never a real file
' ---------------------------------------------------------------------
Public Function SK_GetTestDoc() As String
    Dim d As String
    d = GetSetting(SK_APP, SK_UserKey(), "pref|testDoc", "")
    If Len(d) > 0 Then
        On Error Resume Next
        If Len(Dir$(d)) = 0 Then d = ""
    End If
    SK_GetTestDoc = d
End Function

Public Function SK_SetTestDoc(ByVal path As String) As Boolean
    SaveSetting SK_APP, SK_UserKey(), "pref|testDoc", path
    SK_SetTestDoc = True
End Function

' Returns "Passed ...", "Failed ..." or "Skipped ...".
Public Function SK_TestMacro(ByVal run As String, ByVal worksOn As String) As String
    Dim prev As Object, scratch As Object, t As Single, why As String, src As String, rec As Boolean
    Dim before As New Collection, dd As Object, extra As New Collection, x As Variant, known As Boolean
    On Error Resume Next
    For Each dd In Documents
        before.Add dd.FullName
    Next
    Set prev = ActiveDocument
    Err.Clear
    On Error GoTo MakeFail
    src = SK_GetTestDoc()
    If Len(src) > 0 Then
        Set scratch = Documents.Add(Template:=src)
        scratch.AttachedTemplate = ThisDocument.FullName
    Else
        Set scratch = Documents.Add(Template:=ThisDocument.FullName)
    End If
    scratch.Activate
    scratch.Range(0, 0).Select
    On Error GoTo 0

    If Not SK_NeedsMet(worksOn, why) Then
        SK_TestMacro = "Skipped: " & why & " Pick a test document that starts there (Bulk tools)."
        GoTo Tidy
    End If

    SaveSetting SK_APP, SK_UserKey(), "mark|" & run, "test|" & Format$(Now, "yyyy-mm-dd hh:nn")
    SaveSetting "Stationkeeper Tools", "Run", "quiet", "1"   ' tools skip their messages and dialogs
    t = Timer
    On Error GoTo RunFail
    Application.Run run
    On Error GoTo 0
    SaveSetting SK_APP, SK_UserKey(), "mark|" & run, ""
    SK_SetResult run, "Passed", Format$(Timer - t, "0.00") & " s, test on a scratch document"
    SK_TestMacro = "Passed in " & Format$(Timer - t, "0.00") & " s"
    GoTo Tidy
RunFail:
    SaveSetting SK_APP, SK_UserKey(), "mark|" & run, ""
    SK_SetResult run, "Failed", Err.Description
    SK_TestMacro = "Failed: " & Err.Description
    Resume Tidy
MakeFail:
    SK_TestMacro = "Not tested: could not make a scratch document (" & Err.Description & ")"
    Resume Tidy
Tidy:
    On Error Resume Next
    SaveSetting "Stationkeeper Tools", "Run", "quiet", ""
    ' Close the scratch document and any document the macro opened (a report, say).
    For Each dd In Documents
        known = False
        For Each x In before
            If x = dd.FullName Then known = True: Exit For
        Next
        If Not known Then extra.Add dd
    Next
    For Each dd In extra
        dd.Close SaveChanges:=0
    Next
    If Not prev Is Nothing Then prev.Activate
End Function

' Tests every macro found. Skips the ones whose code saves, closes or deletes
' files or ends Word, unless includeRisky. Returns a summary.
Public Function SK_TestAll(ByVal m As Object, ByVal includeRisky As Boolean) As String
    Dim e As Object, r As String, nP As Long, nF As Long, nS As Long, nR As Long, failed As String
    For Each e In m("entries")
        If e("found") Then
            If Not includeRisky And SK_IsRisky(e) Then
                nR = nR + 1
            Else
                r = SK_TestMacro(e("run"), e("workson"))
                If Left$(r, 6) = "Passed" Then
                    nP = nP + 1
                ElseIf Left$(r, 6) = "Failed" Then
                    nF = nF + 1
                    failed = failed & vbCr & "  " & e("title") & ": " & Mid$(r, 9)
                Else
                    nS = nS + 1
                End If
            End If
        End If
    Next
    SK_TestAll = "Passed: " & nP & "    Failed: " & nF & "    Skipped (needs a selection, table or list): " & nS & _
        "    Not run (saves, closes or deletes files, or ends Word): " & nR & failed
End Function

Public Function SK_IsRisky(ByVal e As Object) As Boolean
    SK_IsRisky = (InStr(e("issues"), "Saves, closes or deletes files") > 0 Or InStr(e("issues"), "Ends Word") > 0)
End Function

' ---------------------------------------------------------------------
' Health check of the code (needs VBA project access). Never runs anything.
' Sets e("issues") on each macro; returns notes for the whole template.
' ---------------------------------------------------------------------
Public Function SK_Health(ByVal m As Object) As String
    Dim e As Object, r As Object, s As String, nm As String
    If Not m("scanAvailable") Then
        SK_Health = "The health check needs ""Trust access to the VBA project object model""."
        Exit Function
    End If
    On Error Resume Next
    For Each r In ThisDocument.VBProject.References
        If r.IsBroken Then
            nm = r.Description
            If Len(nm) = 0 Then nm = r.FullPath
            If Len(nm) = 0 Then nm = r.GUID
            s = s & "Missing library: " & nm & ". Macros that use it fail on this PC." & vbCr
        End If
    Next
    Err.Clear
    On Error GoTo 0
    For Each e In m("entries")
        If e("found") Then e("issues") = SK_IssuesFor(e("run"))
    Next
    If Len(s) = 0 Then s = "Every library the code uses is present."
    SK_Health = s & vbCr & "To check the code compiles, open the VBA editor (Alt+F11) and choose Debug > Compile."
End Function

Private Function SK_IssuesFor(ByVal run As String) As String
    Dim cm As Object, st As Long, n As Long, lines As Variant, i As Long, t As String, low As String
    Dim body As String, b As String, out As String
    On Error GoTo NoCode
    Set cm = ThisDocument.VBProject.VBComponents(SK_ModulePart(run)).CodeModule
    st = cm.ProcStartLine(SK_NamePart(run), 0)
    n = cm.ProcCountLines(SK_NamePart(run), 0)
    lines = Split(cm.Lines(st, n), vbCrLf)
    On Error GoTo 0
    For i = 0 To UBound(lines)
        t = Trim$(lines(i))
        low = LCase$(t)
        If Left$(t, 1) <> "'" And Left$(low, 4) <> "rem " And Len(t) > 0 Then
            body = body & low & vbLf
            If low = "stop" Then out = SK_AddIssue(out, "Has a Stop line, so it pauses in the VBA editor")
            If t Like "*""[A-Za-z]:\*" Or t Like "*""\\*" Then out = SK_AddIssue(out, "Uses a fixed file path that may not exist on other PCs")
            If Left$(low, 5) = "kill " Or InStr(low, " kill ") > 0 Or Left$(low, 9) = "filecopy " Then out = SK_AddIssue(out, "Saves, closes or deletes files")
        End If
    Next
    b = Replace(body, " ", "")
    If InStr(b, ".saveas") > 0 Or (SK_Count(b, ".save") > SK_Count(b, ".saved")) Or InStr(b, ".close") > 0 Then _
        out = SK_AddIssue(out, "Saves, closes or deletes files")
    If InStr(b, "application.quit") > 0 Then out = SK_AddIssue(out, "Ends Word")
    If (InStr(b, "displayalerts=false") > 0 Or InStr(b, "displayalerts=wdalertsnone") > 0 Or InStr(b, "displayalerts=0") > 0) And _
       Not (InStr(b, "displayalerts=true") > 0 Or InStr(b, "displayalerts=wdalertsall") > 0 Or InStr(b, "displayalerts=-1") > 0) Then _
        out = SK_AddIssue(out, "Turns Word alerts off and never back on")
    If InStr(b, "screenupdating=false") > 0 And InStr(b, "screenupdating=true") = 0 Then _
        out = SK_AddIssue(out, "Turns screen updating off and never back on")
    If InStr(b, "onerrorresumenext") > 0 And InStr(b, "onerrorgoto") = 0 Then _
        out = SK_AddIssue(out, "Hides its own errors (On Error Resume Next), so a test cannot see them")
    SK_IssuesFor = out
    Exit Function
NoCode:
    SK_IssuesFor = ""
End Function

Private Function SK_AddIssue(ByVal out As String, ByVal issue As String) As String
    If InStr(vbLf & out & vbLf, vbLf & issue & vbLf) > 0 Then
        SK_AddIssue = out
    Else
        SK_AddIssue = out & IIf(Len(out) > 0, vbLf, "") & issue
    End If
End Function

Private Function SK_Count(ByVal s As String, ByVal f As String) As Long
    SK_Count = (Len(s) - Len(Replace(s, f, ""))) / Len(f)
End Function


Public Function SK_UserKey() As String
    SK_UserKey = ThisDocument.Name
End Function

Public Function SK_Uses(ByVal run As String) As Long
    SK_Uses = Val(GetSetting(SK_APP, SK_UserKey(), "uses|" & run, "0"))
End Function

Public Function SK_LastRun(ByVal run As String) As String
    SK_LastRun = GetSetting(SK_APP, SK_UserKey(), "last|" & run, "")
End Function

Public Function SK_IsPinned(ByVal run As String) As Boolean
    SK_IsPinned = (GetSetting(SK_APP, SK_UserKey(), "pin|" & run, "0") = "1")
End Function

Public Function SK_TogglePin(ByVal run As String) As Boolean
    If SK_IsPinned(run) Then
        SaveSetting SK_APP, SK_UserKey(), "pin|" & run, "0"
    Else
        SaveSetting SK_APP, SK_UserKey(), "pin|" & run, "1"
        SK_TogglePin = True
    End If
End Function

Public Function SK_GetPref(ByVal nm As String, ByVal dflt As String) As String
    SK_GetPref = GetSetting(SK_APP, SK_UserKey(), "pref|" & nm, dflt)
End Function

Public Function SK_SetPref(ByVal nm As String, ByVal v As String) As Boolean
    SaveSetting SK_APP, SK_UserKey(), "pref|" & nm, v
    SK_SetPref = True
End Function

' ---------------------------------------------------------------------
' Ways to open the menu: a key and a button (both saved in the template)
' ---------------------------------------------------------------------
Public Function SK_SetShortcut() As String
    On Error GoTo Fail
    CustomizationContext = MacroContainer
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:="StationkeeperMacros", _
        KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyM)
    SK_SetShortcut = "Alt+Shift+M now opens the menu in documents that use this template."
    Exit Function
Fail:
    SK_SetShortcut = "The shortcut was not set: " & Err.Description
End Function

Public Function SK_AddButton() As String
    Dim cb As Object, btn As Object
    On Error Resume Next
    CustomizationContext = MacroContainer
    CommandBars(SK_APP).Delete
    Err.Clear
    On Error GoTo Fail
    Set cb = CommandBars.Add(Name:=SK_APP, Position:=1, Temporary:=False)
    Set btn = cb.Controls.Add(Type:=1)
    btn.Caption = SK_APP
    btn.Style = 2
    btn.OnAction = "StationkeeperMacros"
    btn.TooltipText = "Open the Stationkeeper Macros menu"
    cb.Visible = True
    SK_AddButton = "The Add-ins tab now has a Stationkeeper Macros button in documents that use this template."
    Exit Function
Fail:
    SK_AddButton = "The button was not added: " & Err.Description
End Function
