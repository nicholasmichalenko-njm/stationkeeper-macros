Attribute VB_Name = "SKW_Workbooks"
' =====================================================================
' Stationkeeper Tools: the standards workbooks (1.22.0).
'   CheckThisWorkbook   says what is wrong with a standards workbook, in plain words. It only reads.
'   NewWorkbook         makes a new, blank standards workbook with every tab and column (1.25.0).
'   AddReadMeTab        adds the Read me tab that says what a workbook is (1.25.0).
' The workbooks are the record both tools read and write (UX decisions 128 to 163). A workbook says
' what it is on its Read me tab; nothing is hidden. Nothing fails silently: a tab, a column or a row
' that cannot be used is named with its place and the reason (decision 134).
' =====================================================================
Option Explicit

' Tabs whose rows are known by a Name.
Private Const TABS_NAME As String = "|styles|table looks|table types|heading numbering|list numbering|page setups|captions|fonts allowed|fields|page numbering|templates|" & _
    "quick parts|pages|stationkeeper admins|authors|piece types|document types|theme|"
' Tabs whose rows are known by another column: tab=column.
Private Const TABS_OTHER As String = "|phrases=Phrase|compliance words=Word|suggestion rules=Phrase|page limits=Heading|placeholders=Text|ignore=Text|style map=From|"
' Tabs that are not lists of rows.
Private Const TABS_PLAIN As String = "|read me|about|columns|history|"
' Tabs only the macros' own older workbook has, and where the one layout keeps that list.
Private Const TABS_OLD As String = "|allowed styles=the Styles tab|fonts=the Fonts allowed tab|banned phrases=the Phrases tab|acronyms=an acronym workbook of its own|" & _
    "dictionaries=the dictionary files beside the workbook|"

' 1.25.0: what a new workbook holds.
Private Const ORG12 As String = "ID|Name|Category|Status|Changed|Used on|Set|Menu group|Show in menu|Notes|Tags|Owner"
' Columns folded away on a new workbook (decision 146: every column is there, the everyday ones open). The last
' column of a tab is never folded: Excel would mark the empty column after it as part of the sheet.
Private Const FOLDED As String = "|Category|Changed|Used on|Menu group|Show in menu|Tags|Owner|"
' The Styles columns left open.
Private Const STYLES_OPEN As String = "|Name|Volume|Status|Set|Notes|Font|Size (pt)|Bold|Italic|Color (hex)|Alignment|Space before (pt)|Space after (pt)|Other Word settings|"
' Columns that hold TRUE, FALSE or nothing.
Private Const SWITCHES As String = "|Show in menu|Required|Bold|Italic|All caps|Small caps|Hidden text|Strikethrough|Double strikethrough|Do not check spelling|" & _
    "Keep with next|Keep lines together|Page break before|Widow control|No hyphenation|No space between paragraphs of the same style|Automatically update|Locked|" & _
    "Show in the Styles gallery|Semi-hidden|Unhide when used|Repeat as header row|One cell across the table|Rows may break across pages|Has caption row|Has header row|" & _
    "Has second header row|Has subhead rows|Has subtotal rows|Has total row|Has notes row|Rows may be left as they are|Page numbers restart|" & _
    "Header and footer continue from the section before|Continues from the part before|Include chapter number|Also in the Quick Parts gallery|"

Private mFound As Collection      ' where Found adds its rows
Private mProblems As Long
Private mNotes As Long

Private Sub Found(ByVal tabName As String, ByVal level As String, ByVal what As String, ByVal toDo As String)
    mFound.Add SKW_Row(tabName, level, what, toDo)
    If level = "Problem" Then mProblems = mProblems + 1
    If level = "Note" Then mNotes = mNotes + 1
End Sub

Private Function After(ByVal pairs As String, ByVal key As String) As String
    Dim p As Long, q As Long
    p = InStr(1, pairs, "|" & key & "=", vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len(key) + 2
    q = InStr(p, pairs, "|")
    After = Mid$(pairs, p, q - p)
End Function

Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    CellText = Trim$(CStr(v))
End Function

' The column whose name in the first row is nm, or 0.
Private Function ColOf(ByVal v As Variant, ByVal nm As String) As Long
    Dim c As Long
    For c = LBound(v, 2) To UBound(v, 2)
        If StrComp(CellText(v(LBound(v, 1), c)), nm, vbTextCompare) = 0 Then ColOf = c: Exit Function
    Next
End Function

' One list tab: its first row, its key column, blank and repeated keys, its Status values.
Private Sub CheckList(ByVal tabName As String, ByVal v As Variant, ByVal firstRow As Long, ByVal keyName As String, ByVal altKey As String)
    Dim c As Long, r As Long, nHdr As Long, kc As Long, sc As Long, seen As Object, k As String, s As String, hasAny As Boolean
    Dim blanks As String, nBlank As Long, dups As String, nDup As Long, badStatus As String, nBad As Long, nRows As Long, rowNo As Long
    Dim uc As Long, vc As Long, also As String
    For c = LBound(v, 2) To UBound(v, 2)
        If Len(CellText(v(LBound(v, 1), c))) > 0 Then nHdr = nHdr + 1
    Next
    If nHdr = 0 Then
        Found tabName, "Problem", "The first row has no column names, so the tab cannot be read.", "Put the column names back in row 1."
        Exit Sub
    End If
    kc = ColOf(v, keyName)
    If kc = 0 And Len(altKey) > 0 Then
        kc = ColOf(v, altKey)
        If kc > 0 Then keyName = altKey
    End If
    If kc = 0 Then
        Found tabName, "Problem", "There is no column named " & keyName & " in row 1, so its rows cannot be told apart.", "Name a column " & keyName & " in row 1. Column names must stay as they are."
        Exit Sub
    End If
    sc = ColOf(v, "Status")
    ' 1.24.0: on Fonts allowed a row is told apart by Name and Used for together; in an effort's workbook a row
    ' set for one volume is its own row.
    If LCase$(tabName) = "fonts allowed" Then uc = ColOf(v, "Used for")
    vc = ColOf(v, "Volume")
    If uc > 0 Then also = " with the same Used for"
    If vc > 0 Then also = also & IIf(Len(also) > 0, " and", " with the same") & " Volume"
    Set seen = CreateObject("Scripting.Dictionary")
    For r = LBound(v, 1) + 1 To UBound(v, 1)
        hasAny = False
        For c = LBound(v, 2) To UBound(v, 2)
            If Len(CellText(v(r, c))) > 0 Then hasAny = True: Exit For
        Next
        If hasAny Then
            nRows = nRows + 1
            rowNo = firstRow + r - LBound(v, 1)
            k = LCase$(CellText(v(r, kc)))
            If Len(k) > 0 And uc > 0 Then k = k & "|" & LCase$(CellText(v(r, uc)))
            If Len(k) > 0 And vc > 0 Then k = k & "|" & LCase$(CellText(v(r, vc)))
            If Len(k) = 0 Then
                nBlank = nBlank + 1
                If nBlank <= 12 Then blanks = blanks & IIf(Len(blanks) > 0, ", ", "") & rowNo
            ElseIf seen.Exists(k) Then
                nDup = nDup + 1
                If nDup <= 8 Then dups = dups & IIf(Len(dups) > 0, "; ", "") & CellText(v(r, kc)) & " (rows " & seen(k) & " and " & rowNo & ")"
            Else
                seen(k) = rowNo
            End If
            If sc > 0 Then
                s = CellText(v(r, sc))
                If Len(s) > 0 And InStr(1, "|approved|draft|retired|", "|" & LCase$(s) & "|") = 0 Then
                    nBad = nBad + 1
                    If nBad <= 8 Then badStatus = badStatus & IIf(Len(badStatus) > 0, "; ", "") & "row " & rowNo & " says " & s
                End If
            End If
        End If
    Next
    If nBlank > 0 Then Found tabName, "Problem", nBlank & " row(s) have values and no " & keyName & ": row " & blanks & IIf(nBlank > 12, " and more", "") & ".", _
        "Give each row a " & keyName & ", or clear the row. A row with no " & keyName & " is not used."
    If nDup > 0 Then Found tabName, "Problem", nDup & " " & keyName & "(s) are there more than once" & also & ": " & dups & IIf(nDup > 8, " and more", "") & ".", _
        IIf(Len(also) > 0, "On this tab a row is told apart by its " & keyName & IIf(uc > 0, " and its Used for", "") & IIf(vc > 0, " and its Volume", "") & " together. Change or remove the second row; only one of them is used.", _
        "A " & keyName & " is used once on a tab. Rename or remove the second row; only one of them is used.")
    If nBad > 0 Then Found tabName, "Problem", nBad & " row(s) have a Status that is not Approved, Draft or Retired: " & badStatus & IIf(nBad > 8, " and more", "") & ".", _
        "Use Approved, Draft or Retired. Only Approved rows are used in documents and templates."
    If nBlank = 0 And nDup = 0 And nBad = 0 Then Found tabName, "Fine", nRows & " row(s), each with its " & keyName & IIf(sc > 0, " and a Status the tools know", "") & ".", ""
End Sub

Public Function SKW_CheckWorkbook(ByVal path As String) As String
    Dim xl As Object, wb As Object, ws As Object, v As Variant, nm As String, low As String, kind As String, owner As String, hasReadMe As Boolean, hasAbout As Boolean
    Dim readMeFirst As Boolean, nOwn As Long, nOld As Long, fromName As String, merged As Variant, f As String, folder As String, nBeside As Long, msg As String, ed As String
    Dim firstRow As Long, one As Variant, i As Long, sayKind As String, tabRows As Collection, topRows As Collection, r As Variant
    Set SKW_LastReport = Nothing
    path = SKW_LocalPath(path)
    If LCase$(Left$(path, 4)) = "http" Then
        SKW_CheckWorkbook = "That workbook is kept in OneDrive or SharePoint, and its copy on this PC was not found in a synced folder. Nothing was checked."
        Exit Function
    End If
    If Len(Dir$(path)) = 0 Then SKW_CheckWorkbook = "The workbook was not found. Nothing was checked.": Exit Function
    Set tabRows = New Collection
    Set topRows = New Collection
    Set mFound = tabRows
    mProblems = 0
    mNotes = 0
    SKW_ProgressTitle "Check This Workbook"
    SKW_Status "Reading the workbook..."
    On Error GoTo Fail
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(FileName:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    i = 0
    For Each ws In wb.Worksheets
        i = i + 1
        nm = Trim$(ws.Name)
        low = LCase$(nm)
        If low = "read me" Then
            hasReadMe = True
            readMeFirst = (i = 1)
            SKW_ReadMeLines ws.Range("A1:C40").Value, kind, owner
        ElseIf low = "about" Then
            hasAbout = True
        ElseIf InStr(1, TABS_PLAIN, "|" & low & "|") > 0 Then
            ' Not a list: nothing to check.
        ElseIf Len(After(TABS_OLD, low)) > 0 Then
            nOld = nOld + 1
            Found nm, "Note", "A tab of the macros' own older workbook. Some older tools still read it.", "In the one layout this list lives in " & After(TABS_OLD, low) & ". Nothing to do yet: the tools will move to it."
        ElseIf InStr(1, TABS_NAME, "|" & low & "|") > 0 Or Len(After(TABS_OTHER, low)) > 0 Or low = "settings" Then
            v = ws.UsedRange.Value
            firstRow = ws.UsedRange.Row
            If Not IsArray(v) Then
                ' One cell or none: a single cell comes back as a value, not a table.
                ReDim one(1 To 1, 1 To 1)
                one(1, 1) = v
                v = one
            End If
            If UBound(v, 1) = LBound(v, 1) And UBound(v, 2) = LBound(v, 2) And Len(CellText(v(LBound(v, 1), LBound(v, 2)))) = 0 Then
                Found nm, "Note", "The tab is empty: no column names and no rows. Nothing is read from it.", "Nothing to do. Put the column names in row 1 before adding rows."
            ElseIf firstRow <> 1 Then
                Found nm, "Problem", "Row 1 is empty: the column names start on row " & firstRow & ".", "Move the column names to row 1."
            Else
                merged = ws.UsedRange.MergeCells
                If IsNull(merged) Then
                    Found nm, "Problem", "Some cells are merged. A list needs one cell per value.", "Unmerge the cells on this tab (Home > Merge & Center > Unmerge Cells)."
                ElseIf CBool(merged) Then
                    Found nm, "Problem", "Some cells are merged. A list needs one cell per value.", "Unmerge the cells on this tab (Home > Merge & Center > Unmerge Cells)."
                End If
                If low = "settings" Then
                    CheckList nm, v, firstRow, "Name", "Setting"
                ElseIf Len(After(TABS_OTHER, low)) > 0 Then
                    CheckList nm, v, firstRow, After(TABS_OTHER, low), ""
                Else
                    CheckList nm, v, firstRow, "Name", ""
                End If
            End If
        Else
            nOwn = nOwn + 1
            Found nm, "Note", "A tab of your own. The tools keep it and never read it.", "Nothing to do. If it was meant to be one of the tools' tabs, its name has changed: give it its name back."
        End If
    Next
    wb.Close False
    xl.Quit
    Set wb = Nothing
    Set xl = Nothing
    On Error GoTo 0
    ' What the workbook says it is: these rows lead the report.
    Set mFound = topRows
    fromName = SKW_KindFromName(path)
    If Len(kind) > 0 Then
        sayKind = kind
        Found "Read me", "Fine", "It says: This workbook is " & kind & "." & IIf(Len(owner) > 0, " It belongs to " & owner & ".", ""), ""
        If Not readMeFirst Then Found "Read me", "Note", "The Read me tab is not the first tab.", "Nothing breaks. It is easier to find as the first tab."
        If kind <> "Global Standards" And Len(owner) = 0 Then Found "Read me", "Note", "The line ""It belongs to"" does not name an effort or a program.", "Type the name beside ""It belongs to"". It is a label for people."
    ElseIf hasReadMe Then
        sayKind = fromName
        Found "Read me", "Problem", "The Read me tab has no line ""This workbook is"" that says Global Standards, Effort Standards or Program Standards." & _
            IIf(Len(fromName) > 0, " Going by the file's name it is read as " & fromName & ".", " The file's name does not say either."), _
            "On the Read me tab, type This workbook is in one cell and Global Standards, Effort Standards or Program Standards in the cell to its right. Until then the tools go by the workbook you point to."
    ElseIf hasAbout Then
        sayKind = fromName
        Found "Read me", "Note", "An older workbook from the Stationkeeper app: it has an About tab and no Read me tab. It is still read." & _
            IIf(Len(fromName) > 0, " Going by the file's name it is " & fromName & ".", " The file's name does not say what it is."), _
            "Nothing breaks without one. Add a Read Me Tab (Tools setup) adds the tab that says what the workbook is."
    Else
        sayKind = fromName
        Found "Read me", "Problem", "There is no Read me tab, so the workbook does not say what it is." & _
            IIf(Len(fromName) > 0, " Going by the file's name it is read as " & fromName & ".", ""), _
            "Run Add a Read Me Tab (Tools setup). Or add a tab named Read me yourself: type This workbook is in one cell and Global Standards, Effort Standards or Program Standards in the cell to its right."
    End If
    ' Other workbooks beside it that could be taken for it.
    folder = Left$(path, InStrRev(path, "\") - 1)
    On Error Resume Next
    f = Dir$(folder & "\*.xls*")
    Do While Len(f) > 0
        If Left$(f, 2) <> "~$" And StrComp(folder & "\" & f, path, vbTextCompare) <> 0 Then
            If Len(sayKind) > 0 And SKW_KindFromName(f) = sayKind Then nBeside = nBeside + 1
        End If
        f = Dir$
    Loop
    On Error GoTo 0
    If nBeside > 0 Then Found "(the folder)", "Note", nBeside & " other workbook(s) in this folder are named like " & sayKind & " workbooks. One may be a copy or a sync conflict copy.", _
        "Keep one workbook of a kind in a folder. The tools use the one you point to."
    For Each r In tabRows
        topRows.Add r
    Next
    SKW_ProgressDone
    msg = Mid$(path, InStrRev(path, "\") + 1) & IIf(Len(sayKind) > 0, " (" & sayKind & ")", "") & ": " & mProblems & " problem(s), " & mNotes & " note(s). " & _
        IIf(mProblems = 0, "The tools can read every tab they know.", "Each problem names its tab, its rows and what to do.") & _
        " Tabs of your own: " & nOwn & ". Changes nothing."
    SKW_ReportFor "Check This Workbook", msg, Array("Tab", "Level", "What was found", "What to do"), topRows, Mid$(path, InStrRev(path, "\") + 1)
    SKW_CheckWorkbook = msg
    Exit Function
Fail:
    ed = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    SKW_ProgressDone
    SKW_CheckWorkbook = "The workbook could not be read (" & ed & "). Nothing was checked." & vbCr & "If it is open in Excel with a cell being typed in, press Enter there and try again."
End Function

' Asks for a standards workbook and says what is wrong with it in plain words: what it says it is on its Read me tab, each tab the tools know and each tab of your own, a first row without column names, a missing Name column, blank and repeated Names, a Status that is not Approved, Draft or Retired, and merged cells in a list. The workbook is only read.
Public Sub CheckThisWorkbook()
    Dim p As String
    If SKW_IsQuiet() Then Exit Sub   ' needs a person to choose the file
    With Application.FileDialog(3)
        .Title = "Choose the workbook to check (it is only read)"
        .Filters.Clear
        .Filters.Add "Excel workbooks", "*.xlsx;*.xlsm"
        .AllowMultiSelect = False
        If Len(SKW_StandardsPath()) > 0 Then .InitialFileName = SKW_StandardsPath()
        If .Show <> -1 Then Exit Sub
        p = .SelectedItems(1)
    End With
    SKW_Show SKW_CheckWorkbook(p)
End Sub

' =====================================================================
' 1.25.0: New Workbook and Add a Read Me Tab.
' A new workbook carries every tab and column of the one layout (project page "the column list of every tab,
' for both tools", 2026-11-19), under the plain names, with no rows. The Stationkeeper app reads those names
' since 4.57.0; an older app misreads them. No format number is stated yet (agreed with the app's build thread).
' =====================================================================

' The three kinds, as the Read me tab writes them; "" for anything else.
Public Function SKW_KindWords(ByVal kind As String) As String
    Select Case LCase$(Trim$(kind))
        Case "global standards", "global", "1": SKW_KindWords = "Global Standards"
        Case "effort standards", "effort", "2": SKW_KindWords = "Effort Standards"
        Case "program standards", "program", "3": SKW_KindWords = "Program Standards"
    End Select
End Function

' The tabs of a new workbook of a kind, after Read me, in order, as a|b|c.
Public Function SKW_BookTabs(ByVal kind As String) As String
    Dim s As String
    s = "Settings|Styles|Fonts allowed|Table looks|Table types|Heading numbering|List numbering|Captions|Page setups|Page numbering|Templates|Theme|" & _
        "Phrases|Placeholders|Ignore|Style map|"
    If SKW_KindWords(kind) <> "Global Standards" Then s = s & "Page limits|"
    s = s & "Quick Parts|Pages"
    If SKW_KindWords(kind) = "Global Standards" Then s = s & "|Stationkeeper Admins"
    SKW_BookTabs = s
End Function

' The columns of a tab, in order, as a|b|c. In an effort's or a program's workbook Volume follows Name on the
' standards tabs and on Settings.
Public Function SKW_BookColumns(ByVal tabName As String, ByVal kind As String) As String
    Dim s As String, std As Boolean
    std = True
    Select Case LCase$(Trim$(tabName))
        Case "styles"
            s = ORG12 & "|Required|" & _
                "Font|Theme font|Size (pt)|Bold|Italic|Underline|All caps|Small caps|Hidden text|Strikethrough|Double strikethrough|Superscript or subscript|" & _
                "Color (hex)|Character spacing (pt)|Kerning from (pt)|Character shading (hex, or none)|Do not check spelling|" & _
                "Type|Based on|Next style|Alignment|Outline level|Keep with next|Keep lines together|Page break before|Widow control|No hyphenation|" & _
                "Space before (pt)|Space after (pt)|Line spacing (lines, or pt if exact)|Line rule|No space between paragraphs of the same style|" & _
                "Left indent (pt)|Right indent (pt)|First line (pt)|Hanging (pt)|Tab stops|" & _
                "Border above|Border below|Border left|Border right|Shading (hex, or none)|" & _
                "Automatically update|Linked style|Locked|Show in the Styles gallery|Semi-hidden|Unhide when used|Order in the gallery|Other Word settings"
        Case "fonts allowed"
            s = ORG12 & "|Used for|Min size (pt)|Max size (pt)|RFP reference|Other Word settings"
        Case "table looks"
            s = ORG12 & "|Row style|Font|Size (pt)|Bold|Italic|Color (hex)|Alignment|Vertical alignment|Repeat as header row|One cell across the table|" & _
                "Space before (pt)|Space after (pt)|Left indent (pt)|Right indent (pt)|Edge above|Edge below|Edge left|Edge right|Edge between columns|" & _
                "Shading (hex, or none)|Recognized by|Tab stops|Other Word settings"
        Case "table types"
            s = ORG12 & "|Alignment|Width (%)|Cell margin top (pt)|Cell margin bottom (pt)|Cell margin left (pt)|Cell margin right (pt)|Caption position|" & _
                "Rows may break across pages|Has caption row|Has header row|Has second header row|Has subhead rows|Has subtotal rows|Has total row|Has notes row|" & _
                "Rows may be left as they are|Caption look|Header look|Second header look|Subhead look|Body look|Subtotal look|Total look|Notes look|Other Word settings"
        Case "heading numbering"
            s = ORG12 & "|Number format|Number text|Number at (in)|Text at (in)|Restart|Style|Other Word settings"
        Case "list numbering"
            s = ORG12 & "|List|Level|Number format|Number text|Font|Left indent (in)|Hanging (in)|Style|Other Word settings"
        Case "captions"
            s = ORG12 & "|Label|Part|Position|Chapter level|Separator|Identifier|After number|Style|Action style|Other Word settings"
        Case "page setups"
            s = ORG12 & "|Paper width (in)|Paper height (in)|Orientation|Top margin (in)|Bottom margin (in)|Left margin (in)|Right margin (in)|" & _
                "Header distance (in)|Footer distance (in)|Page numbers restart|Header and footer continue from the section before|Other Word settings"
        Case "page numbering"
            s = ORG12 & "|Number format|Starts at|Continues from the part before|Include chapter number|Chapter level|Separator|Other Word settings"
        Case "templates"
            s = Replace(ORG12, "|Set|", "|") & "|Style set|Setup style set|Quick Part set|Page set|Page layout|Numbering"
        Case "settings": s = "Name|Value|Notes"
        Case "theme": s = "Name|Value|Notes": std = False
        Case "phrases": s = "Phrase|Level|Use instead|Why|Status": std = False
        Case "placeholders", "ignore": s = "Text|Notes": std = False
        Case "style map": s = "From|To|Notes": std = False
        Case "page limits": s = "Heading|Max pages|Does not count|Notes": std = False
        Case "quick parts": s = "ID|Name|Gallery|Category|Source file|Styles used|Layout needed|Set|Status|Changed|Notes": std = False
        Case "pages": s = "ID|Name|Kind|Gallery|Category|Source file|Styles used|Layout needed|Set|Also in the Quick Parts gallery|Status|Changed|Notes": std = False
        Case "stationkeeper admins": s = "Name|Notes": std = False
    End Select
    If std And Len(s) > 0 And SKW_KindWords(kind) <> "Global Standards" And Len(SKW_KindWords(kind)) > 0 Then
        If Left$(s, 5) = "Name|" Then
            s = "Name|Volume|" & Mid$(s, 6)
        Else
            s = Replace(s, "|Name|", "|Name|Volume|", 1, 1)
        End If
    End If
    SKW_BookColumns = s
End Function

' The choices a column offers in its drop-down, as a,b,c; "" for none. Only words both tools read are offered.
Private Function ChoicesFor(ByVal tabName As String, ByVal col As String) As String
    If InStr(1, SWITCHES, "|" & col & "|", vbTextCompare) > 0 Then ChoicesFor = "TRUE,FALSE": Exit Function
    Select Case LCase$(col)
        Case "status": ChoicesFor = "Approved,Draft,Retired"
        Case "alignment"
            Select Case LCase$(tabName)
                Case "table types": ChoicesFor = "Left,Centered,Right"
                Case Else: ChoicesFor = "Left,Centered,Right,Justified"
            End Select
        Case "line rule": ChoicesFor = "Multiple,Exactly,At least"
        Case "orientation": ChoicesFor = "Portrait,Landscape"
        Case "used for": ChoicesFor = "body,heading,table,caption,header and footer,footnote,graphic,any"
    End Select
End Function

' The Read me tab: the two lines in the agreed cells, then the text both tools use word for word.
Private Sub WriteReadMe(ByVal ws As Object, ByVal kind As String, ByVal owner As String, ByVal withText As Boolean)
    Dim r As Long, v As Variant
    ws.Range("A1").Value = "This workbook is"
    ws.Range("B1").Value = kind
    ws.Range("A2").Value = "It belongs to"
    ws.Range("B2").NumberFormat = "@"
    ws.Range("B2").Value = owner
    On Error Resume Next
    With ws.Range("B1").Validation
        .Delete
        .Add 3, 1, 1, "Global Standards,Effort Standards,Program Standards"
        .IgnoreBlank = True
    End With
    ws.Range("A1:A2").Font.Bold = True
    On Error GoTo 0
    If Not withText Then Exit Sub
    r = 4
    For Each v In Array( _
        "Stationkeeper and the Stationkeeper Macros both read and write this workbook. You can edit it in Excel.", _
        "Free to change: the file's name; any row; the order of tabs and columns; colours and widths. You can add tabs and columns of your own; they are kept and never reach Word.", _
        "Please do not change: the name of this tab; the name of any other tab; the first row of any tab (the column names).", _
        "A blank cell means ""not set"". In an effort's workbook it means ""use the Global Standards value"".", _
        "Only rows marked Approved are used in documents and templates. Draft rows are waiting for a Stationkeeper Admin.", _
        "If something stops working: run ""Check this workbook"" from the Stationkeeper Macros menu or from Stationkeeper. It says what is wrong in plain words.", _
        "Close the workbook when you are done. While it is open in Excel, neither tool can save to it; they keep your change and save it when the file is free.")
        ws.Cells(r, 1).Value = CStr(v)
        r = r + 1
    Next
    r = r + 1
    ws.Cells(r, 1).Value = "Made by"
    ws.Cells(r, 2).Value = "Stationkeeper Macros " & SKW_VERSION
    On Error Resume Next
    ws.Columns(1).ColumnWidth = 18
    ws.Columns(2).ColumnWidth = 40
    On Error GoTo 0
End Sub

' One list tab of a new workbook: its column names, the ID out of sight, the less used columns folded, drop-downs.
Private Sub WriteListTab(ByVal xl As Object, ByVal ws As Object, ByVal tabName As String, ByVal kind As String)
    Dim cols() As String, i As Long, n As Long, a As Long, fold As Boolean, ch As String, isStyles As Boolean
    cols = Split(SKW_BookColumns(tabName, kind), "|")
    n = UBound(cols) + 1
    ws.Range(ws.Cells(1, 1), ws.Cells(1, n)).NumberFormat = "@"
    For i = 1 To n
        ws.Cells(1, i).Value = cols(i - 1)
    Next
    On Error Resume Next
    ws.Range(ws.Cells(1, 1), ws.Cells(1, n)).Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(1, n)).EntireColumn.AutoFit
    For i = 1 To n
        ' Number text holds %1 and the like, which Excel would turn into a number: the column is a text column.
        If cols(i - 1) = "Number text" Then ws.Columns(i).NumberFormat = "@"
        ch = ChoicesFor(tabName, cols(i - 1))
        If Len(ch) > 0 Then
            With ws.Range(ws.Cells(2, i), ws.Cells(2000, i)).Validation
                .Delete
                .Add 3, 1, 1, ch
                .IgnoreBlank = True
                .ShowError = False
            End With
        End If
    Next
    ' Fold the runs of less used columns.
    isStyles = (LCase$(tabName) = "styles")
    a = 0
    For i = 1 To n + 1
        fold = False
        If i <= n Then
            If cols(i - 1) = "ID" Then
                fold = False
            ElseIf isStyles Then
                fold = (InStr(1, STYLES_OPEN, "|" & cols(i - 1) & "|", vbTextCompare) = 0)
            Else
                fold = (InStr(1, FOLDED, "|" & cols(i - 1) & "|", vbTextCompare) > 0)
            End If
        End If
        If fold Then
            If a = 0 Then a = i
        ElseIf a > 0 Then
            ws.Range(ws.Columns(a), ws.Columns(i - 1)).Group
            a = 0
        End If
    Next
    ws.Outline.ShowLevels 0, 1
    If cols(0) = "ID" Then ws.Columns(1).Hidden = True
    ws.Activate
    xl.ActiveWindow.SplitRow = 1
    xl.ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

' Makes a new, blank standards workbook of a kind at path. Never writes over a file. Returns what happened, in words.
Public Function SKW_NewWorkbook(ByVal path As String, ByVal kind As String, ByVal owner As String) As String
    Dim xl As Object, wb As Object, ws As Object, t As Variant, ed As String, nTabs As Long, folder As String, i As Long
    kind = SKW_KindWords(kind)
    owner = Trim$(owner)
    If Len(kind) = 0 Then SKW_NewWorkbook = "Say which kind: Global Standards, Effort Standards or Program Standards. Nothing was made.": Exit Function
    If kind <> "Global Standards" And Len(owner) = 0 Then
        SKW_NewWorkbook = IIf(kind = "Effort Standards", "An effort's workbook needs the effort's name.", "A program's workbook needs the program's name.") & " Nothing was made."
        Exit Function
    End If
    path = SKW_LocalPath(path)
    If LCase$(Left$(path, 4)) = "http" Then SKW_NewWorkbook = "That folder is kept in OneDrive or SharePoint, and its copy on this PC was not found in a synced folder. Nothing was made.": Exit Function
    If LCase$(Right$(path, 5)) <> ".xlsx" Then path = path & ".xlsx"
    If InStrRev(path, "\") < 2 Then SKW_NewWorkbook = "The workbook needs a folder. Nothing was made.": Exit Function
    folder = Left$(path, InStrRev(path, "\") - 1)
    If Len(Dir$(folder, vbDirectory)) = 0 Then SKW_NewWorkbook = "The folder was not found. Nothing was made.": Exit Function
    If Len(Dir$(path)) > 0 Then
        SKW_NewWorkbook = "A file named " & Mid$(path, InStrRev(path, "\") + 1) & " is already in that folder. Nothing was made and nothing was changed. Choose another name or folder."
        Exit Function
    End If
    SKW_ProgressTitle "New Workbook"
    SKW_Status "Making the workbook..."
    On Error GoTo Fail
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    wb.Worksheets(1).Name = "Read me"
    WriteReadMe wb.Worksheets(1), kind, owner, True
    For Each t In Split(SKW_BookTabs(kind), "|")
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = CStr(t)
        WriteListTab xl, ws, CStr(t), kind
        nTabs = nTabs + 1
    Next
    wb.Worksheets(1).Activate
    wb.SaveAs FileName:=path, FileFormat:=51
    wb.Close False
    xl.Quit
    Set wb = Nothing
    Set xl = Nothing
    SKW_ProgressDone
    SKW_NewWorkbook = "Made " & Mid$(path, InStrRev(path, "\") + 1) & ": " & IIf(kind = "Effort Standards", "an ", "a ") & kind & " workbook" & IIf(Len(owner) > 0, " for " & owner, "") & _
        ", with a Read me tab and " & nTabs & " tabs, each with all its columns and no rows." & vbCr & _
        "To use it, choose it on the menu's standards line. Check This Workbook reads it back." & vbCr & _
        "It uses the plain column names. Stationkeeper 4.57.0 or later reads them; an older Stationkeeper misreads them, so update the app on a PC before it opens this workbook."
    Exit Function
Fail:
    ed = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    SKW_ProgressDone
    SKW_NewWorkbook = "The workbook could not be made (" & ed & "). Nothing was changed."
End Function

' Adds a Read me tab that says what a workbook is, as its first tab. Changes nothing else. A workbook that
' already says what it is is left alone. Returns what happened, in words.
Public Function SKW_AddReadMe(ByVal path As String, ByVal kind As String, ByVal owner As String) As String
    Dim xl As Object, wb As Object, ws As Object, rm As Object, has As String, hasOwner As String, ed As String, nm As String, filled As Boolean, c As Variant, added As Boolean
    path = SKW_LocalPath(path)
    If LCase$(Left$(path, 4)) = "http" Then SKW_AddReadMe = "That workbook is kept in OneDrive or SharePoint, and its copy on this PC was not found in a synced folder. Nothing was changed.": Exit Function
    If Len(Dir$(path)) = 0 Then SKW_AddReadMe = "The workbook was not found. Nothing was changed.": Exit Function
    nm = Mid$(path, InStrRev(path, "\") + 1)
    If Len(Trim$(kind)) = 0 Then kind = SKW_KindFromName(path)
    kind = SKW_KindWords(kind)
    owner = Trim$(owner)
    If Len(kind) = 0 Then SKW_AddReadMe = "Say which kind " & nm & " is: Global Standards, Effort Standards or Program Standards. Its name does not say. Nothing was changed.": Exit Function
    SKW_ProgressTitle "Add a Read Me Tab"
    SKW_Status "Opening the workbook..."
    On Error GoTo Fail
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(FileName:=path, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
    If wb.ReadOnly Then
        wb.Close False
        xl.Quit
        SKW_ProgressDone
        SKW_AddReadMe = nm & " is open somewhere else or cannot be written to. Nothing was changed. Close it and run this again."
        Exit Function
    End If
    For Each ws In wb.Worksheets
        If LCase$(Trim$(ws.Name)) = "read me" Then Set rm = ws
    Next
    If Not rm Is Nothing Then
        SKW_ReadMeLines rm.Range("A1:C40").Value, has, hasOwner
        If Len(has) > 0 Then
            wb.Close False
            xl.Quit
            SKW_ProgressDone
            SKW_AddReadMe = nm & " already says what it is: " & has & IIf(Len(hasOwner) > 0, ", belonging to " & hasOwner, "") & ". Nothing was changed."
            Exit Function
        End If
        For Each c In rm.Range("A1:B2").Cells
            If Len(Trim$(CellText(c.Value))) > 0 Then filled = True
        Next
        If filled Then
            wb.Close False
            xl.Quit
            SKW_ProgressDone
            SKW_AddReadMe = nm & " has a Read me tab with other words in its first two rows. Nothing was changed. Type This workbook is in one cell there and " & kind & " in the cell to its right."
            Exit Function
        End If
        WriteReadMe rm, kind, owner, False
    Else
        Set rm = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        rm.Name = "Read me"
        WriteReadMe rm, kind, owner, True
        added = True
    End If
    wb.Save
    wb.Close False
    xl.Quit
    Set wb = Nothing
    Set xl = Nothing
    SKW_ProgressDone
    SKW_AddReadMe = nm & " now says what it is: " & kind & IIf(Len(owner) > 0, ", belonging to " & owner, "") & ". " & _
        IIf(added, "A Read me tab was added as its first tab.", "The two lines were written at the top of its Read me tab.") & " No other tab was changed."
    Exit Function
Fail:
    ed = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    SKW_ProgressDone
    SKW_AddReadMe = "The workbook could not be changed (" & ed & "). Nothing was saved."
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim v As Variant
    For Each v In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, CStr(v), " ")
    Next
    SafeFileName = Trim$(s)
End Function

' Makes a new, blank standards workbook: the team's Global Standards, one effort's Effort Standards or one program's Program Standards. It has a Read me tab that says what it is, and every tab with all its columns and no rows, under the plain column names. It asks for the kind, the name and the folder, and never writes over a file. Stationkeeper 4.57.0 or later reads the plain names; an older Stationkeeper misreads them.
Public Sub NewWorkbook()
    Dim kind As String, owner As String, folder As String, p As String, f As String
    If SKW_IsQuiet() Then Exit Sub   ' needs a person to choose
    kind = SKW_KindWords(InputBox("Which kind of workbook? Type its number." & vbCr & vbCr & "1  Global Standards: the team's" & vbCr & _
        "2  Effort Standards: one effort's own values, laid over the team's" & vbCr & "3  Program Standards: one program's own values", "New Workbook", "2"))
    If Len(kind) = 0 Then Exit Sub
    If kind = "Global Standards" Then
        owner = Trim$(InputBox("The team's name, for the line ""It belongs to"". You can leave it blank.", "New Workbook"))
    Else
        owner = Trim$(InputBox("The " & IIf(kind = "Effort Standards", "effort's", "program's") & " name. It goes on the line ""It belongs to"" and starts the file's name.", "New Workbook"))
        If Len(owner) = 0 Then Exit Sub
    End If
    With Application.FileDialog(4)
        .Title = "Choose the folder for the new workbook"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        folder = .SelectedItems(1)
    End With
    If Right$(folder, 1) = "\" Then folder = Left$(folder, Len(folder) - 1)
    f = IIf(kind = "Global Standards", "Global Standards", SafeFileName(owner) & " " & kind) & ".xlsx"
    p = folder & "\" & f
    If MsgBox("Make " & f & " in this folder?" & vbCr & folder & vbCr & vbCr & "It will have every tab with all its columns, and no rows." & vbCr & vbCr & _
        "It uses the plain column names. Stationkeeper 4.57.0 or later reads them; an older Stationkeeper misreads them. With no Stationkeeper app on a PC it is fine.", _
        vbOKCancel + vbQuestion, "New Workbook") <> vbOK Then Exit Sub
    SKW_Show SKW_NewWorkbook(p, kind, owner)
End Sub

' Adds a Read me tab to a standards workbook that has none, as its first tab, so the workbook says what it is: Global Standards, Effort Standards or Program Standards, and who it belongs to. Nothing else in the workbook is changed. A workbook that already says what it is is left alone. The workbook must be closed in Excel.
Public Sub AddReadMeTab()
    Dim p As String, kind As String, owner As String, nm As String, v As Variant
    If SKW_IsQuiet() Then Exit Sub   ' needs a person to choose the file
    With Application.FileDialog(3)
        .Title = "Choose the workbook to add a Read me tab to"
        .Filters.Clear
        .Filters.Add "Excel workbooks", "*.xlsx;*.xlsm"
        .AllowMultiSelect = False
        If Len(SKW_StandardsPath()) > 0 Then .InitialFileName = SKW_StandardsPath()
        If .Show <> -1 Then Exit Sub
        p = .SelectedItems(1)
    End With
    nm = Mid$(p, InStrRev(p, "\") + 1)
    kind = SKW_KindFromName(p)
    If Len(kind) = 0 Then
        kind = SKW_KindWords(InputBox("What is " & nm & "? Type its number." & vbCr & vbCr & "1  Global Standards" & vbCr & "2  Effort Standards" & vbCr & "3  Program Standards", "Add a Read Me Tab"))
        If Len(kind) = 0 Then Exit Sub
    End If
    owner = nm
    If InStrRev(owner, ".") > 1 Then owner = Left$(owner, InStrRev(owner, ".") - 1)
    For Each v In Array("Effort Standards", "Program Standards", "Effort Formatting", "Global Standards")
        owner = Trim$(Replace(owner, CStr(v), "", 1, -1, vbTextCompare))
    Next
    owner = Trim$(InputBox("Who does it belong to? The " & IIf(kind = "Global Standards", "team's", IIf(kind = "Effort Standards", "effort's", "program's")) & _
        " name, for the line ""It belongs to""." & IIf(kind = "Global Standards", " You can leave it blank.", ""), "Add a Read Me Tab", owner))
    If MsgBox("Add a Read me tab to " & nm & " that says it is " & kind & IIf(Len(owner) > 0, ", belonging to " & owner, "") & "?" & vbCr & vbCr & _
        "Nothing else in the workbook is changed. Close the workbook in Excel first.", vbOKCancel + vbQuestion, "Add a Read Me Tab") <> vbOK Then Exit Sub
    SKW_Show SKW_AddReadMe(p, kind, owner)
End Sub
