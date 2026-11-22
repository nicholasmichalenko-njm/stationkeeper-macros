Attribute VB_Name = "SKW_Capture"
' =====================================================================
' Stationkeeper Tools: capture into the Global Standards workbook (1.17.0).
' The Stationkeeper app reads its Global Standards workbook again whenever the
' file changes, so a row written here in the app's own shape reaches the app.
'   CaptureStyles              the document's styles that the Styles sheet lacks, as Draft rows
'   CaptureBannedPhrase        a phrase onto the Phrases sheet
'   CaptureHeadingNumbering    the document's heading numbering onto the Heading numbering sheet
' Rules the writer keeps, from the app's code (standardsbook.py):
'   - a row is found by its Name; the app gives a row with no ID one of its own
'   - a Draft row is held by the app and reaches no template until it is approved there
'   - rows are only added or changed, never removed: the app reads a missing row as taken out
'   - only columns the workbook already has are filled; a style's other values go in Other Word settings
'   - nothing is written while Excel has the workbook open, and a dated copy is kept before every write
' Each tool reports first and writes only on a yes.
' =====================================================================
Option Explicit

Public SKW_LastBookBackup As String      ' the copy kept before the last write

Private Function Q(ByVal s As String) As String
    Q = """" & Replace(Replace(s, "\", "\\"), """", "\""") & """"
End Function

Private Function JVal(ByVal v As Variant) As String
    Select Case VarType(v)
        Case vbBoolean: JVal = IIf(v, "true", "false")
        Case vbString: JVal = Q(CStr(v))
        Case Else: JVal = Trim$(Str$(v))
    End Select
End Function

Private Function CellText(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then CellText = IIf(v, "True", "False"): Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbSingle Then
        CellText = Trim$(Str$(Round(v, 4)))
        If Left$(CellText, 1) = "." Then CellText = "0" & CellText
        If Left$(CellText, 2) = "-." Then CellText = "-0" & Mid$(CellText, 2)
        Exit Function
    End If
    CellText = Trim$(CStr(v))
End Function

' Is the file free to write? Excel holding it open locks it.
Private Function FileFree(ByVal p As String) As Boolean
    Dim f As Integer
    On Error GoTo Held
    f = FreeFile
    Open p For Binary Access Read Write Lock Read Write As #f
    Close #f
    FileFree = True
    Exit Function
Held:
    On Error Resume Next
    Close #f
End Function

Private Function BackupOf(ByVal p As String) As String
    Dim d As String, nm As String
    On Error Resume Next
    d = Environ$("LOCALAPPDATA") & "\Stationkeeper Macros"
    If Len(Dir$(d, vbDirectory)) = 0 Then MkDir d
    d = d & "\Workbook backups"
    If Len(Dir$(d, vbDirectory)) = 0 Then MkDir d
    nm = Mid$(p, InStrRev(p, "\") + 1)
    BackupOf = d & "\" & Left$(nm, InStrRev(nm, ".") - 1) & " " & Format$(Now, "yyyy-mm-dd hhnnss") & Mid$(nm, InStrRev(nm, "."))
    Err.Clear
    FileCopy p, BackupOf
    If Err.Number <> 0 Then BackupOf = ""
End Function

' Writes rows to one sheet of the Global Standards workbook.
'   rows: each a Dictionary. Its entries are column header -> value. Two special entries:
'         "*other" -> a Dictionary of app key -> value for the Other Word settings column
'         "*keys"  -> a Dictionary of column header (lower case) -> app key, for values whose column the sheet lacks
'   keyHeader: the column that names a row.   update: change the cells of a row that is already there.
' Returns "" when it worked, else why not. added, changed and same count rows.
Public Function SKW_BookWrite(ByVal sheet As String, ByVal keyHeader As String, ByVal rows As Collection, ByVal update As Boolean, _
                              ByRef added As Long, ByRef changed As Long, ByRef same As Long) As String
    SKW_BookWrite = SKW_BookWriteTo(SKW_StandardsPath(), sheet, keyHeader, rows, update, added, changed, same)
End Function

' The same writer for a workbook named by its path (1.21.0): the work PC proofs write a scratch workbook and a copy with it.
Public Function SKW_BookWriteTo(ByVal p As String, ByVal sheet As String, ByVal keyHeader As String, ByVal rows As Collection, ByVal update As Boolean, _
                                ByRef added As Long, ByRef changed As Long, ByRef same As Long) As String
    Dim xl As Object, wb As Object, ws As Object, hdr As Object, c As Long, lastCol As Long, lastRow As Long, r As Long
    Dim have As Object, row As Object, k As Variant, key As String, at As Long, h As String, other As String, o As Object, keys As Object
    Dim did As Boolean, ed As String, isNew As Boolean
    added = 0: changed = 0: same = 0
    SKW_LastBookBackup = ""
    If Len(p) = 0 Then SKW_BookWriteTo = "No Global Standards workbook is chosen. Choose the Stationkeeper app's Global Standards.xlsx first. Nothing was written.": Exit Function
    If Len(Dir$(p)) = 0 Then SKW_BookWriteTo = "The Global Standards workbook was not found: " & p & ". Nothing was written.": Exit Function
    If Not FileFree(p) Then
        SKW_BookWriteTo = "The Global Standards workbook is open in Excel or held by another program. Close it there and run this again. Nothing was written."
        Exit Function
    End If
    SKW_LastBookBackup = BackupOf(p)
    If Len(SKW_LastBookBackup) = 0 Then SKW_BookWriteTo = "A copy of the workbook could not be kept first, so nothing was written.": Exit Function
    On Error GoTo Fail
    SKW_Status "Writing to the Global Standards workbook..."
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(FileName:=p, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
    On Error Resume Next
    Set ws = wb.Worksheets(sheet)
    On Error GoTo Fail
    If ws Is Nothing Then
        wb.Close False: xl.Quit
        SKW_BookWriteTo = "The workbook has no " & sheet & " sheet. The Stationkeeper app writes that sheet; open the app once so it does, then run this again. Nothing was written."
        Exit Function
    End If
    Set hdr = CreateObject("Scripting.Dictionary")
    lastCol = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
    For c = 1 To lastCol
        h = LCase$(Trim$(CellText(ws.Cells(1, c).Value)))
        If Len(h) > 0 And Not hdr.Exists(h) Then hdr(h) = c
    Next
    ' 1.24.0: a column the workbook names the new, plain way is written to under the name the tools use.
    If Len(SKW_NewNames(sheet)) > 0 Then
        For Each k In Split(SKW_NewNames(sheet), "|")
            h = Left$(k, InStr(k, "=") - 1)
            If hdr.Exists(h) And Not hdr.Exists(Mid$(k, InStr(k, "=") + 1)) Then hdr(Mid$(k, InStr(k, "=") + 1)) = hdr(h)
        Next
    End If
    If Not hdr.Exists(LCase$(keyHeader)) Then
        wb.Close False: xl.Quit
        SKW_BookWriteTo = "The " & sheet & " sheet has no " & keyHeader & " column. Nothing was written."
        Exit Function
    End If
    lastRow = ws.Cells(ws.Rows.Count, hdr(LCase$(keyHeader))).End(-4162).Row
    If lastRow < 1 Then lastRow = 1
    Set have = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRow
        key = LCase$(CellText(ws.Cells(r, hdr(LCase$(keyHeader))).Value))
        If Len(key) > 0 And Not have.Exists(key) Then have(key) = r
    Next
    For Each row In rows
        key = LCase$(Trim$(CStr(row(keyHeader))))
        If Len(key) > 0 Then
            isNew = Not have.Exists(key)
            If isNew Or update Then
                If isNew Then
                    lastRow = lastRow + 1
                    at = lastRow
                    have(key) = at
                Else
                    at = have(key)
                End If
                did = False
                Set keys = Nothing
                If row.Exists("*keys") Then Set keys = row("*keys")
                Set o = CreateObject("Scripting.Dictionary")
                If row.Exists("*other") Then
                    For Each k In row("*other").Keys
                        o(k) = row("*other")(k)
                    Next
                End If
                For Each k In row.Keys
                    If Left$(k, 1) <> "*" Then
                        h = LCase$(CStr(k))
                        If hdr.Exists(h) Then
                            ' 1.26.0: a choice in an older word and the same choice in Word's words are the same: left as it is.
                            If CellText(ws.Cells(at, hdr(h)).Value) <> CellText(row(k)) And Not SKW_SameChoice(sheet, h, ws.Cells(at, hdr(h)).Value, row(k)) Then
                                ws.Cells(at, hdr(h)).Value = row(k)
                                ' 1.25.0: Excel turns some text into something else as it lands (%1 becomes 0.01, a
                                ' leading = becomes a formula). Such a cell is made a text cell and written again.
                                If VarType(row(k)) = vbString Then
                                    If CellText(ws.Cells(at, hdr(h)).Value) <> CellText(row(k)) Then
                                        ws.Cells(at, hdr(h)).NumberFormat = "@"
                                        ws.Cells(at, hdr(h)).Value = row(k)
                                    End If
                                End If
                                did = True
                            End If
                        ElseIf Not keys Is Nothing Then
                            ' A value whose column this workbook does not have yet: it rides in Other Word settings.
                            ' 1.26.0: there a choice stays in the app's older word, which is what the app reads in that cell.
                            If keys.Exists(h) Then
                                If Len(SKW_ChoiceMeaning(h, row(k))) > 0 Then
                                    If Len(SKW_ChoiceOld(h, row(k))) > 0 Then o(keys(h)) = SKW_ChoiceOld(h, row(k))
                                Else
                                    o(keys(h)) = row(k)
                                End If
                            End If
                        End If
                    End If
                Next
                If isNew And o.Count > 0 And hdr.Exists("other word settings") Then
                    other = ""
                    For Each k In o.Keys
                        other = other & IIf(Len(other) > 0, ", ", "") & Q(CStr(k)) & ": " & JVal(o(k))
                    Next
                    ws.Cells(at, hdr("other word settings")).Value = "{" & other & "}"
                End If
                If isNew Then
                    added = added + 1
                ElseIf did Then
                    changed = changed + 1
                Else
                    same = same + 1
                End If
            Else
                same = same + 1
            End If
        End If
    Next
    If added + changed > 0 Then wb.Save
    wb.Close False
    xl.Quit
    ' The tools read the workbook again.
    If LCase$(p) = LCase$(SKW_StandardsPath()) Then SKW_SetStandardsPath p
    Exit Function
Fail:
    ed = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    SKW_BookWriteTo = "Stopped: " & ed & ". The workbook was closed without saving. The copy kept first is " & SKW_LastBookBackup & "."
End Function

Private Function AfterWrite(ByVal added As Long, ByVal changed As Long) As String
    If added + changed = 0 Then AfterWrite = "The workbook already held it. Nothing was written.": Exit Function
    AfterWrite = "The Stationkeeper app takes it in the next time it reads the workbook. A copy of the workbook as it was is kept at " & SKW_LastBookBackup & "."
End Function

' ---------------------------------------------------------------------
' Styles
' ---------------------------------------------------------------------
Private Function HexOf(ByVal c As Long) As String
    If c < 0 Then Exit Function               ' automatic
    HexOf = Right$("0" & Hex$(c And &HFF), 2) & Right$("0" & Hex$((c \ 256) And &HFF), 2) & Right$("0" & Hex$((c \ 65536) And &HFF), 2)
End Function

Private Sub Prop(ByVal row As Object, ByVal keys As Object, ByVal label As String, ByVal key As String, ByVal v As Variant)
    If VarType(v) = vbString Then
        If Len(v) = 0 Then Exit Sub
    End If
    row(label) = v
    keys(LCase$(label)) = key
End Sub

' One style as a row for the Styles sheet: the column headers are the app's own (libstyle.py PROPS).
Private Function StyleRow(ByVal d As Object, ByVal st As Object) As Object
    Dim row As Object, keys As Object, pf As Object, f As Object, nm As String, isPara As Boolean
    Set row = CreateObject("Scripting.Dictionary")
    Set keys = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    isPara = (st.Type = wdStyleTypeParagraph Or st.Type = 6)
    row("Name") = st.NameLocal
    row("Status") = "Draft"
    row("Notes") = "Captured from " & d.Name & " by Stationkeeper Macros, " & Format$(Now, "yyyy-mm-dd")
    Prop row, keys, "Type", "type", IIf(isPara, "paragraph", "character")
    nm = "": nm = st.BaseStyle
    Prop row, keys, "Based on", "based_on", nm
    If isPara Then
        nm = "": nm = st.NextParagraphStyle
        If nm <> st.NameLocal Then Prop row, keys, "Next style", "next", nm
    End If
    Set f = st.Font
    Prop row, keys, "Font", "font", f.Name
    Prop row, keys, "Size (pt)", "size", CDbl(f.Size)
    Prop row, keys, "Bold", "bold", CBool(f.Bold)
    Prop row, keys, "Italic", "italic", CBool(f.Italic)
    If f.Underline <> wdUnderlineNone Then Prop row, keys, "Underline", "underline", True
    If f.AllCaps Then Prop row, keys, "All caps", "caps", True
    If f.SmallCaps Then Prop row, keys, "Small caps", "small_caps", True
    If f.Hidden Then Prop row, keys, "Hidden text", "hidden", True
    Prop row, keys, "Color (hex)", "color", HexOf(f.Color)
    If isPara Then
        Set pf = st.ParagraphFormat
        Select Case pf.Alignment
            Case wdAlignParagraphLeft: Prop row, keys, "Alignment", "jc", "Left"
            Case wdAlignParagraphCenter: Prop row, keys, "Alignment", "jc", "Centered"
            Case wdAlignParagraphRight: Prop row, keys, "Alignment", "jc", "Right"
            Case wdAlignParagraphJustify: Prop row, keys, "Alignment", "jc", "Justified"
        End Select
        Prop row, keys, "Space before (pt)", "before", Round(CDbl(pf.SpaceBefore), 2)
        Prop row, keys, "Space after (pt)", "after", Round(CDbl(pf.SpaceAfter), 2)
        Select Case pf.LineSpacingRule
            Case wdLineSpaceExactly
                Prop row, keys, "Line spacing (lines, or pt if exact)", "line", Round(CDbl(pf.LineSpacing), 2)
                Prop row, keys, "Line rule", "line_rule", "Exactly"
            Case wdLineSpaceAtLeast
                Prop row, keys, "Line spacing (lines, or pt if exact)", "line", Round(CDbl(pf.LineSpacing), 2)
                Prop row, keys, "Line rule", "line_rule", "At least"
            Case Else
                Prop row, keys, "Line spacing (lines, or pt if exact)", "line", Round(CDbl(pf.LineSpacing) / 12, 2)
                Prop row, keys, "Line rule", "line_rule", "Multiple"
        End Select
        Prop row, keys, "Left indent (pt)", "left", Round(CDbl(pf.LeftIndent), 2)
        Prop row, keys, "Right indent (pt)", "right", Round(CDbl(pf.RightIndent), 2)
        If pf.FirstLineIndent > 0.05 Then Prop row, keys, "First line (pt)", "first", Round(CDbl(pf.FirstLineIndent), 2)
        If pf.FirstLineIndent < -0.05 Then Prop row, keys, "Hanging (pt)", "hanging", Round(-CDbl(pf.FirstLineIndent), 2)
        If pf.KeepWithNext Then Prop row, keys, "Keep with next", "keep_next", True
        If pf.KeepTogether Then Prop row, keys, "Keep lines together", "keep_lines", True
        If pf.PageBreakBefore Then Prop row, keys, "Page break before", "page_break_before", True
        Prop row, keys, "Widow control", "widow", CBool(pf.WidowControl)
    End If
    Set row("*keys") = keys
    Set StyleRow = row
End Function

Private Function RowWords(ByVal row As Object) As String
    Dim k As Variant, s As String
    For Each k In row.Keys
        If Left$(k, 1) <> "*" And k <> "Name" And k <> "Status" And k <> "Notes" Then s = s & IIf(Len(s) > 0, "; ", "") & k & ": " & CellText(row(k))
    Next
    RowWords = s
End Function

' The styles to offer: the paragraph styles the document's text is in, and every custom paragraph or character style.
' only: "" for all of them, or names joined with "|".
Private Function StylesToCapture(ByVal d As Object, ByVal only As String) As Collection
    Dim c As New Collection, v As Variant, st As Object, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    If Len(only) > 0 Then
        For Each v In Split(only, "|")
            Set st = Nothing
            Set st = d.Styles(Trim$(v))
            If Not st Is Nothing Then c.Add st
        Next
    Else
        For Each v In SKW_RemapStyleNames(True)
            Set st = Nothing
            Set st = d.Styles(CStr(v))
            If Not st Is Nothing Then seen(LCase$(st.NameLocal)) = True: c.Add st
        Next
        For Each st In d.Styles
            If Not st.BuiltIn And st.InUse And (st.Type = wdStyleTypeParagraph Or st.Type = 6 Or st.Type = wdStyleTypeCharacter) Then
                If Not seen.Exists(LCase$(st.NameLocal)) Then seen(LCase$(st.NameLocal)) = True: c.Add st
            End If
        Next
    End If
    Set StylesToCapture = c
End Function

Private Function OnSheet(ByVal sheet As String, ByVal keyHeader As String) As Object
    Dim d As Object, r As Object
    Set d = CreateObject("Scripting.Dictionary")
    For Each r In SKW_WbRows(sheet)
        If r.Exists(LCase$(keyHeader)) Then
            If Len(r(LCase$(keyHeader))) > 0 Then Set d(LCase$(r(LCase$(keyHeader)))) = r
        End If
    Next
    Set OnSheet = d
End Function

Public Function SKW_CaptureStylesCheck(ByVal only As String) As String
    Dim d As Object, msg As String, rows As New Collection, st As Object, have As Object, n As Long, row As Object, wasSaved As Boolean, there As Long
    Set SKW_LastReport = Nothing
    SKW_Fixable = 0
    msg = SKW_Begin("Capture styles: reading", False)
    If Len(msg) > 0 Then SKW_CaptureStylesCheck = msg: Exit Function
    Set d = ActiveDocument
    wasSaved = d.Saved
    If SKW_WbRows("Styles").Count = 0 And SKW_WbRows("Captions").Count = 0 And SKW_WbRows("Table looks").Count = 0 Then
        SKW_Finish
        SKW_CaptureStylesCheck = "No Global Standards workbook from the Stationkeeper app is chosen. Choose it first (the menu's Global Standards line). Nothing was read."
        Exit Function
    End If
    Set have = OnSheet("Styles", "Name")
    For Each st In StylesToCapture(d, only)
        Set row = StyleRow(d, st)
        If have.Exists(LCase$(st.NameLocal)) Then
            there = there + 1
            rows.Add SKW_Row(st.NameLocal, row("Type"), "Already there (" & have(LCase$(st.NameLocal))("status") & ")", "Left as it is")
        Else
            n = n + 1
            rows.Add SKW_Row(st.NameLocal, row("Type"), "Not there", "Added as a Draft row: " & RowWords(row))
        End If
    Next
    If wasSaved And Not d.Saved Then d.Saved = True
    SKW_Finish
    SKW_Fixable = n
    msg = n & " style(s) of this document are not on the Global Standards' Styles sheet; " & there & " already are."
    SKW_Report "Capture Styles", msg & " A captured style is added as a Draft row: the Stationkeeper app holds a Draft and it reaches no template until it is approved in the app. " & _
        "A style already on the sheet is never changed. The values are the style's own as Word shows them. Changes nothing.", _
        Array("Style", "Type", "On the Styles sheet", "What capture does"), rows
    SKW_CaptureStylesCheck = msg
End Function

Public Function SKW_CaptureStylesWrite(ByVal only As String) As String
    Dim d As Object, rows As New Collection, st As Object, have As Object, e As String, a As Long, c As Long, s As Long
    Set SKW_LastReport = Nothing
    If Not SKW_IsAdmin() Then SKW_CaptureStylesWrite = SKW_NotAdminMsg(): Exit Function
    If Documents.Count = 0 Then SKW_CaptureStylesWrite = "Open a document first.": Exit Function
    Set d = ActiveDocument
    Set have = OnSheet("Styles", "Name")
    For Each st In StylesToCapture(d, only)
        If Not have.Exists(LCase$(st.NameLocal)) Then rows.Add StyleRow(d, st)
    Next
    If rows.Count = 0 Then SKW_CaptureStylesWrite = "Every one of those styles is already on the Styles sheet. Nothing was written.": Exit Function
    e = SKW_BookWrite("Styles", "Name", rows, False, a, c, s)
    Application.StatusBar = False
    If Len(e) > 0 Then SKW_CaptureStylesWrite = e: Exit Function
    SKW_CaptureStylesWrite = a & " style(s) added to the Global Standards' Styles sheet as Draft rows. " & AfterWrite(a, c)
End Function

' Adds the document's styles that the Global Standards' Styles sheet lacks, as Draft rows the Stationkeeper app holds until they are approved there. Reports first; writes only on a yes. A style already on the sheet is never changed. A dated copy of the workbook is kept first.
Public Sub CaptureStyles()
    Dim msg As String
    If SKW_IsQuiet() Then Exit Sub
    If Not SKW_AdminGate() Then Exit Sub
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    SKW_Fixable = 0
    msg = SKW_CaptureStylesCheck("")
    If SKW_Fixable <= 0 Or SKW_LastReport Is Nothing Then SKW_Show msg: Exit Sub
    If MsgBox(msg & vbCr & vbCr & "Add the " & SKW_Fixable & " missing style(s) to the Global Standards workbook as Draft rows?" & vbCr & vbCr & _
        "A copy of the workbook is kept first. Close the workbook in Excel before you say yes. The report stays open.", vbYesNo + vbQuestion, SKW_APP) <> vbYes Then Exit Sub
    SKW_Say SKW_CaptureStylesWrite("")
End Sub

' ---------------------------------------------------------------------
' Banned phrases (the Phrases sheet: Phrase, Level, Use instead, Why)
' ---------------------------------------------------------------------
Public Function SKW_CapturePhrase(ByVal phrase As String, ByVal level As String, ByVal useInstead As String, ByVal why As String) As String
    Dim rows As New Collection, row As Object, e As String, a As Long, c As Long, s As Long
    If Not SKW_IsAdmin() Then SKW_CapturePhrase = SKW_NotAdminMsg(): Exit Function
    phrase = Trim$(Replace(Replace(phrase, vbCr, " "), vbTab, " "))
    If Len(phrase) = 0 Then SKW_CapturePhrase = "No phrase was given. Nothing was written.": Exit Function
    If Len(phrase) > 120 Then SKW_CapturePhrase = "That is longer than a phrase (120 characters). Nothing was written.": Exit Function
    level = IIf(LCase$(Left$(Trim$(level), 1)) = "a", "Avoid", "Banned")
    Set row = CreateObject("Scripting.Dictionary")
    row("Phrase") = phrase
    row("Level") = level
    If Len(Trim$(useInstead)) > 0 Then row("Use instead") = Trim$(useInstead)
    If Len(Trim$(why)) > 0 Then row("Why") = Trim$(why)
    rows.Add row
    e = SKW_BookWrite("Phrases", "Phrase", rows, False, a, c, s)
    Application.StatusBar = False
    If Len(e) > 0 Then SKW_CapturePhrase = e: Exit Function
    If a = 0 Then SKW_CapturePhrase = """" & phrase & """ is already on the Phrases sheet. Nothing was written.": Exit Function
    SKW_CapturePhrase = """" & phrase & """ added to the Global Standards' Phrases sheet as " & level & ". " & AfterWrite(a, c)
End Function

' Adds the selected words (or a phrase you type) to the Global Standards' Phrases sheet as Banned or Avoid, with what to use instead and why. A phrase already on the sheet is left as it is. A dated copy of the workbook is kept first.
Public Sub CaptureBannedPhrase()
    Dim ph As String, lv As String, ui As String, wy As String
    If SKW_IsQuiet() Then Exit Sub
    If Not SKW_AdminGate() Then Exit Sub
    On Error Resume Next
    If Documents.Count > 0 Then
        If Selection.Type = wdSelectionNormal And Len(Selection.Range.Text) <= 120 Then ph = Trim$(Replace(Selection.Range.Text, vbCr, " "))
    End If
    On Error GoTo 0
    ph = InputBox("The phrase to add to the Global Standards' Phrases sheet:", "Capture Banned Phrase", ph)
    If Len(Trim$(ph)) = 0 Then Exit Sub
    lv = InputBox("1  Banned" & vbCr & "2  Avoid" & vbCr & vbCr & "Type the number.", "Capture Banned Phrase", "1")
    If Len(lv) = 0 Then Exit Sub
    ui = InputBox("What to use instead (leave it blank for none):", "Capture Banned Phrase", "")
    wy = InputBox("Why (leave it blank for none):", "Capture Banned Phrase", "")
    SKW_Say SKW_CapturePhrase(ph, IIf(lv = "2", "Avoid", "Banned"), ui, wy)
End Sub

' ---------------------------------------------------------------------
' Heading numbering (the Heading numbering sheet: one row per heading style)
' ---------------------------------------------------------------------
' 1.26.0: in the words Word itself shows (column list rule 6). The older words are still read (SKW_SameChoice).
Private Function FmtOf(ByVal ns As Long) As String
    Select Case ns
        Case wdListNumberStyleArabic: FmtOf = "1, 2, 3"
        Case wdListNumberStyleUppercaseRoman: FmtOf = "I, II, III"
        Case wdListNumberStyleLowercaseRoman: FmtOf = "i, ii, iii"
        Case wdListNumberStyleUppercaseLetter: FmtOf = "A, B, C"
        Case wdListNumberStyleLowercaseLetter: FmtOf = "a, b, c"
        Case wdListNumberStyleBullet: FmtOf = "Bullet"
    End Select
End Function

' The document's numbering for Heading 1 to 9: rows for the sheet. A heading that is not numbered, or numbered
' in a format the sheet has no word for, is left out and named in skipped.
Private Function HeadingRows(ByVal d As Object, ByRef skipped As String) As Collection
    Dim c As New Collection, i As Long, st As Object, lt As Object, lv As Object, row As Object, f As String, n As Long
    On Error Resume Next
    For i = 1 To 9
        Set st = Nothing: Set lt = Nothing: Set lv = Nothing
        Set st = d.Styles(-1 - i)
        Set lt = st.ListTemplate
        If lt Is Nothing Then
            skipped = skipped & "; " & st.NameLocal & " (not numbered in this document)"
        Else
            n = st.ListLevelNumber
            If n < 1 Then n = i
            Set lv = lt.ListLevels(n)
            f = FmtOf(lv.NumberStyle)
            If Len(f) = 0 Then
                skipped = skipped & "; " & st.NameLocal & " (a number format the sheet has no word for)"
            Else
                Set row = CreateObject("Scripting.Dictionary")
                row("Name") = st.NameLocal
                row("Style") = st.NameLocal
                row("Fmt") = f
                row("Text") = lv.NumberFormat
                row("Number at") = Round(CDbl(lv.NumberPosition) / 72, 2)
                row("Text at") = Round(CDbl(lv.TextPosition) / 72, 2)
                c.Add row
            End If
        End If
    Next
    Set HeadingRows = c
End Function

Public Function SKW_CaptureHeadingsCheck() As String
    Dim d As Object, msg As String, rows As New Collection, have As Object, row As Object, r As Object, skipped As String, n As Long, k As Variant
    Dim diff As String, wasSaved As Boolean, nNew As Long, same As Boolean
    Set SKW_LastReport = Nothing
    SKW_Fixable = 0
    msg = SKW_Begin("Capture heading numbering: reading", False)
    If Len(msg) > 0 Then SKW_CaptureHeadingsCheck = msg: Exit Function
    Set d = ActiveDocument
    wasSaved = d.Saved
    Set have = OnSheet("Heading numbering", "Style")
    For Each row In HeadingRows(d, skipped)
        If have.Exists(LCase$(row("Style"))) Then
            Set r = have(LCase$(row("Style")))
            diff = ""
            For Each k In Array("Fmt", "Text", "Number at", "Text at")
                If k = "Fmt" Then
                    same = SKW_SameChoice("Heading numbering", "Fmt", row(k), r(LCase$(k)))
                Else
                    same = (CellText(row(k)) = CellText(IIf(IsNumeric(r(LCase$(k))) And k <> "Text", Val(r(LCase$(k))), r(LCase$(k)))))
                End If
                If Not same Then
                    diff = diff & IIf(Len(diff) > 0, "; ", "") & k & ": " & r(LCase$(k)) & " becomes " & CellText(row(k))
                End If
            Next
            If Len(diff) > 0 Then
                n = n + 1
                rows.Add SKW_Row(row("Style"), "Differs", diff)
            Else
                rows.Add SKW_Row(row("Style"), "The same", "")
            End If
        Else
            nNew = nNew + 1
            rows.Add SKW_Row(row("Style"), "Not on the sheet", "Added: " & RowWords(row))
        End If
    Next
    If wasSaved And Not d.Saved Then d.Saved = True
    SKW_Finish
    SKW_Fixable = n + nNew
    msg = n & " heading level(s) are numbered differently in this document than on the Global Standards' Heading numbering sheet, and " & nNew & " are not on the sheet." & _
        IIf(Len(skipped) > 0, " Left out: " & Mid$(skipped, 3) & ".", "")
    SKW_Report "Capture Heading Numbering", msg & " Writing changes the rows that differ. A changed row keeps its status, so an Approved row changes the standard " & _
        "for everyone the next time the Stationkeeper app reads the workbook. Changes nothing.", Array("Heading style", "Against the sheet", "What capture does"), rows
    SKW_CaptureHeadingsCheck = msg
End Function

Public Function SKW_CaptureHeadingsWrite() As String
    Dim d As Object, rows As Collection, row As Object, e As String, a As Long, c As Long, s As Long, skipped As String
    Set SKW_LastReport = Nothing
    If Not SKW_IsAdmin() Then SKW_CaptureHeadingsWrite = SKW_NotAdminMsg(): Exit Function
    If Documents.Count = 0 Then SKW_CaptureHeadingsWrite = "Open a document first.": Exit Function
    Set d = ActiveDocument
    Set rows = HeadingRows(d, skipped)
    If rows.Count = 0 Then SKW_CaptureHeadingsWrite = "No heading in this document is numbered in a way the sheet can hold. Nothing was written.": Exit Function
    For Each row In rows
        row("Status") = "Draft"       ' used only for a row that is new; a row already there keeps its own
    Next
    e = SKW_BookWriteHeadings(rows, a, c, s)
    Application.StatusBar = False
    If Len(e) > 0 Then SKW_CaptureHeadingsWrite = e: Exit Function
    SKW_CaptureHeadingsWrite = c & " heading level(s) changed and " & a & " added (as Draft) on the Global Standards' Heading numbering sheet; " & s & " were the same. " & AfterWrite(a, c)
End Function

' A row already on the sheet keeps its Status: only a new row is written as Draft.
Private Function SKW_BookWriteHeadings(ByVal rows As Collection, ByRef a As Long, ByRef c As Long, ByRef s As Long) As String
    Dim have As Object, row As Object
    Set have = OnSheet("Heading numbering", "Style")
    For Each row In rows
        If have.Exists(LCase$(row("Style"))) Then
            row.Remove "Status"
            row.Remove "Name"
        End If
    Next
    SKW_BookWriteHeadings = SKW_BookWrite("Heading numbering", "Style", rows, True, a, c, s)
End Function

' Compares the document's heading numbering (Heading 1 to 9: format, text, positions) with the Global Standards' Heading numbering sheet, then writes the document's numbering to the sheet on a yes. Changing an Approved row changes the standard for everyone. A dated copy of the workbook is kept first.
Public Sub CaptureHeadingNumbering()
    Dim msg As String
    If SKW_IsQuiet() Then Exit Sub
    If Not SKW_AdminGate() Then Exit Sub
    If Documents.Count = 0 Then SKW_Say "Open a document first.": Exit Sub
    SKW_Fixable = 0
    msg = SKW_CaptureHeadingsCheck()
    If SKW_Fixable <= 0 Or SKW_LastReport Is Nothing Then SKW_Show msg: Exit Sub
    If MsgBox(msg & vbCr & vbCr & "Write this document's heading numbering to the Global Standards workbook?" & vbCr & vbCr & _
        "This changes the standard for everyone the next time the Stationkeeper app reads the workbook. A copy of the workbook is kept first. " & _
        "Close the workbook in Excel before you say yes. The report stays open.", vbYesNo + vbExclamation + vbDefaultButton2, SKW_APP) <> vbYes Then Exit Sub
    SKW_Say SKW_CaptureHeadingsWrite()
End Sub
