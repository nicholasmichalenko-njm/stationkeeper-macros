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
Public Const SKW_VERSION As String = "0.5.0"

Public SKW_Quiet As Boolean              ' tests: no message boxes, reports made hidden
Public SKW_ForceCoauthored As Boolean    ' tests: behave as if the file is co-authored
Public SKW_LastReport As Object          ' the last report document made
Public SKW_LastRows As Long              ' findings in the last report

Private mUndo As Boolean
Private mScreen As Boolean
Private mStd As Object
Private mStdPath As String
Private mStdStamp As Date
Private mStdRead As Boolean

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
    ' Written as an HTML file and opened as a new document, so nothing is edited in Word while it is
    ' made: the checked document's Undo list cannot pick up a step (Word sometimes files edits made to
    ' a hidden new document under the active one).
    Dim src As Object, d As Object, h As String, r As Variant, i As Long, tmp As String, v As Variant
    Dim parts() As String, one As String, n As Long
    Set src = ActiveDocument
    h = "<html><head><meta charset=""utf-8""><style>body{font-family:Calibri,sans-serif;font-size:10pt}" & _
        "h1{font-size:16pt}table{border-collapse:collapse}th,td{border:1px solid #999;padding:2pt 4pt;vertical-align:top;font-size:9pt}" & _
        "th{background:#DDEBF7;text-align:left}</style></head><body>" & _
        "<h1>" & SKW_Html(title) & "</h1><p>Document: " & SKW_Html(src.Name) & " &nbsp; " & Format$(Now, "yyyy-mm-dd hh:nn") & _
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
    Set SKW_LastReport = d
    Set SKW_Report = d
    If SKW_IsQuiet() Then src.Activate
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
' The standards workbook (one per template, chosen per person).
' Sheets are read once and kept until the file changes.
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

Public Function SKW_StdNote() As String
    If SKW_HasStandards() Then
        SKW_StdNote = "Standards: " & Mid$(mStdPath, InStrRev(mStdPath, "\") + 1)
    Else
        SKW_StdNote = "Standards: built-in defaults (no standards workbook chosen)"
    End If
End Function

Public Function SKW_Std() As Object
    Dim p As String, stamp As Date
    p = SKW_StandardsPath()
    If Len(p) = 0 Then Set SKW_Std = Nothing: Exit Function
    On Error Resume Next
    stamp = FileDateTime(p)
    If Err.Number <> 0 Then Set SKW_Std = Nothing: Exit Function
    On Error GoTo 0
    If mStdRead And p = mStdPath And stamp = mStdStamp Then Set SKW_Std = mStd: Exit Function
    Set mStd = SKW_ReadWorkbook(p)
    mStdPath = p
    mStdStamp = stamp
    mStdRead = True
    Set SKW_Std = mStd
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

' Placeholders: the workbook's list, else the defaults.
Public Function SKW_Placeholders() As Collection
    Dim c As Collection, v As Variant
    Set c = SKW_Column("Placeholders", "Text")
    If c.Count = 0 Then
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
