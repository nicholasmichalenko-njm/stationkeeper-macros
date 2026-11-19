Attribute VB_Name = "SKW_Workbooks"
' =====================================================================
' Stationkeeper Tools: the standards workbooks (1.22.0).
'   CheckThisWorkbook   says what is wrong with a standards workbook, in plain words. It only reads.
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
    If nDup > 0 Then Found tabName, "Problem", nDup & " " & keyName & "(s) are there more than once: " & dups & IIf(nDup > 8, " and more", "") & ".", _
        "A " & keyName & " is used once on a tab. Rename or remove the second row; only one of them is used."
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
            "Nothing to do now. A Read me tab that says what the workbook is will be added by the tools in a later release."
    Else
        sayKind = fromName
        Found "Read me", "Problem", "There is no Read me tab, so the workbook does not say what it is." & _
            IIf(Len(fromName) > 0, " Going by the file's name it is read as " & fromName & ".", ""), _
            "Add a tab named Read me. Type This workbook is in one cell and Global Standards, Effort Standards or Program Standards in the cell to its right."
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
