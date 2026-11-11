Attribute VB_Name = "SKX_Book"
' =====================================================================
' Stationkeeper Tools for Excel, run from Word.
' They check the workbook open in Excel (they never start Excel on their own
' for this or open a file), write their reports as Word documents, and change
' a workbook only in a copy saved beside it.
'   BookFormulaErrors, BookBrokenNames, BookExternalLinks, BookHiddenContent,
'   BookLeftovers, BookBannedPhrases, BookAcronymCheck, BookFontAudit,
'   BookMergedCells, BookDeliveryCopy
' =====================================================================
Option Explicit

Public SKX_TestBook As Object      ' tests: the workbook to check, so a person's open workbooks are never touched
Private mWhy As String
Private Const MAXCELLS As Long = 20000

' Lists cells showing an error (#REF!, #N/A, #DIV/0!, #VALUE! and the rest) in the workbook open in Excel. Changes nothing.
Public Sub BookFormulaErrors()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Show SKX_BookFormulaErrors()
End Sub

' Lists defined names that point at #REF! or at another file. Changes nothing.
Public Sub BookBrokenNames()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Show SKX_BookBrokenNames()
End Sub

' Lists links to other workbooks and the cells whose formulas use them. Changes nothing.
Public Sub BookExternalLinks()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Show SKX_BookExternalLinks()
End Sub

' Lists hidden and very hidden sheets, hidden rows and columns, and hidden names. They travel with the file. Changes nothing.
Public Sub BookHiddenContent()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Show SKX_BookHiddenContent()
End Sub

' Lists comments, notes and placeholder text (TBD, XXX, [insert) in the open workbook. Changes nothing.
Public Sub BookLeftovers()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Show SKX_BookLeftovers()
End Sub

' Lists the team's banned phrases in the open workbook's cell text, with the replacement. Changes nothing.
Public Sub BookBannedPhrases()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Show SKX_BookBannedPhrases()
End Sub

' Lists the acronyms in the open workbook's cell text with the team meaning. Text only. Changes nothing.
Public Sub BookAcronymCheck()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Show SKX_BookAcronymCheck()
End Sub

' Lists the fonts used in the open workbook's filled cells and flags ones not on the team's Fonts sheet. Changes nothing.
Public Sub BookFontAudit()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Show SKX_BookFontAudit()
End Sub

' Lists merged cells. They break sorting, filtering, copying and screen readers. Changes nothing.
Public Sub BookMergedCells()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Show SKX_BookMergedCells()
End Sub

' Saves a delivery copy of the open workbook beside it ("<name> - delivery.xlsx"): comments and notes out, links to other files turned into values, personal details out. The open workbook is not changed.
Public Sub BookDeliveryCopy()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Say SKX_BookDeliveryCopy("")
End Sub

' ---------------------------------------------------------------------
Public Function SKX_Book() As Object
    Dim xl As Object, wb As Object
    mWhy = ""
    If Not SKX_TestBook Is Nothing Then Set SKX_Book = SKX_TestBook: Exit Function
    On Error Resume Next
    Set xl = GetObject(, "Excel.Application")
    On Error GoTo 0
    If xl Is Nothing Then mWhy = "Open the workbook in Excel first. The tools check the workbook that is open; they never open one on their own.": Exit Function
    On Error Resume Next
    Set wb = xl.ActiveWorkbook
    On Error GoTo 0
    If wb Is Nothing Then
        If xl.Workbooks.Count = 1 Then
            Set wb = xl.Workbooks(1)
        Else
            mWhy = IIf(xl.Workbooks.Count = 0, "Open the workbook in Excel first.", "Several workbooks are open. Click in the one to check, then run this again.")
            Exit Function
        End If
    End If
    Set SKX_Book = wb
End Function

Private Function StartBook(ByVal title As String, ByRef wb As Object) As String
    Set SKW_LastReport = Nothing
    Set wb = SKX_Book()
    If wb Is Nothing Then StartBook = mWhy: Exit Function
    SKW_Status title & ": " & wb.Name & "..."
End Function

' Cells holding typed text (not formulas, not numbers), capped.
Private Function TextCells(ByVal ws As Object) As Collection
    Dim c As New Collection, r As Object, cel As Object
    On Error Resume Next
    Set r = ws.UsedRange.SpecialCells(2, 2)   ' xlCellTypeConstants, xlTextValues
    On Error GoTo 0
    If r Is Nothing Then Set TextCells = c: Exit Function
    For Each cel In r.Cells
        c.Add cel
        If c.Count >= MAXCELLS Then Exit For
    Next
    Set TextCells = c
End Function

Private Function Addr(ByVal cel As Object) As String
    Addr = cel.Address(False, False)
End Function

' ---------------------------------------------------------------------
Public Function SKX_BookFormulaErrors() As String
    Dim wb As Object, msg As String, rows As New Collection, ws As Object, r As Object, cel As Object, n As Long, t As Variant
    msg = StartBook("Formula errors", wb)
    If Len(msg) > 0 Then SKX_BookFormulaErrors = msg: Exit Function
    For Each ws In wb.Worksheets
        For Each t In Array(-4123, 2)   ' formulas, constants
            Set r = Nothing
            On Error Resume Next
            Set r = ws.UsedRange.SpecialCells(t, 16)   ' xlErrors
            On Error GoTo 0
            If Not r Is Nothing Then
                For Each cel In r.Cells
                    n = n + 1
                    If n <= 1000 Then rows.Add SKW_Row(ws.Name, Addr(cel), cel.Text, IIf(cel.HasFormula, cel.Formula, "(typed)"))
                Next
            End If
        Next
    Next
    SKW_Status ""
    SKW_ReportFor "Book Formula Errors", n & " cell(s) show an error. Changes nothing.", Array("Sheet", "Cell", "Shows", "Formula"), rows, wb.Name
    SKX_BookFormulaErrors = n & " cell(s) with an error."
End Function

Public Function SKX_BookBrokenNames() As String
    Dim wb As Object, msg As String, rows As New Collection, nm As Object, rt As String, n As Long
    msg = StartBook("Broken names", wb)
    If Len(msg) > 0 Then SKX_BookBrokenNames = msg: Exit Function
    For Each nm In wb.Names
        rt = ""
        On Error Resume Next
        rt = nm.RefersTo
        On Error GoTo 0
        If InStr(rt, "#REF!") > 0 Then
            n = n + 1: rows.Add SKW_Row(nm.Name, "Points to #REF! (the cells it named were deleted)", rt)
        ElseIf InStr(rt, "[") > 0 Or InStr(rt, ":\") > 0 Or InStr(rt, "\\") > 0 Then
            n = n + 1: rows.Add SKW_Row(nm.Name, "Points to another file", rt)
        End If
    Next
    SKW_Status ""
    SKW_ReportFor "Book Broken Names", wb.Names.Count & " name(s); " & n & " broken or pointing at another file. Changes nothing.", Array("Name", "Finding", "Refers to"), rows, wb.Name
    SKX_BookBrokenNames = n & " broken name(s)."
End Function

Public Function SKX_BookExternalLinks() As String
    Dim wb As Object, msg As String, rows As New Collection, links As Variant, i As Long, ws As Object, r As Object, cel As Object
    Dim nLinks As Long, nCells As Long
    msg = StartBook("External links", wb)
    If Len(msg) > 0 Then SKX_BookExternalLinks = msg: Exit Function
    On Error Resume Next
    links = wb.LinkSources(1)   ' xlExcelLinks
    On Error GoTo 0
    If IsArray(links) Then
        For i = LBound(links) To UBound(links)
            nLinks = nLinks + 1
            rows.Add SKW_Row("Link", links(i), IIf(Len(Dir$(CStr(links(i)))) > 0, "file found", "file not found"), "")
        Next
    End If
    For Each ws In wb.Worksheets
        Set r = Nothing
        On Error Resume Next
        Set r = ws.UsedRange.SpecialCells(-4123)   ' formulas
        On Error GoTo 0
        If Not r Is Nothing Then
            For Each cel In r.Cells
                If InStr(cel.Formula, "[") > 0 Then
                    nCells = nCells + 1
                    If nCells <= 500 Then rows.Add SKW_Row("Cell", ws.Name & "!" & Addr(cel), "", cel.Formula)
                End If
            Next
        End If
    Next
    SKW_Status ""
    SKW_ReportFor "Book External Links", nLinks & " linked workbook(s), " & nCells & " cell(s) using them. Book Delivery Copy turns them into values. Changes nothing.", _
        Array("Kind", "Where", "Status", "Formula"), rows, wb.Name
    SKX_BookExternalLinks = nLinks & " linked workbook(s), " & nCells & " cell(s)."
End Function

Public Function SKX_BookHiddenContent() As String
    Dim wb As Object, msg As String, rows As New Collection, ws As Object, nm As Object, hr As Long, hc As Long, r As Object, nS As Long, nRC As Long, nN As Long
    msg = StartBook("Hidden content", wb)
    If Len(msg) > 0 Then SKX_BookHiddenContent = msg: Exit Function
    For Each ws In wb.Worksheets
        If ws.Visible = 0 Then nS = nS + 1: rows.Add SKW_Row(ws.Name, "Hidden sheet", "Right-click a sheet tab > Unhide shows it")
        If ws.Visible = 2 Then nS = nS + 1: rows.Add SKW_Row(ws.Name, "Very hidden sheet", "Only VBA can show it; it still travels with the file")
        hr = 0: hc = 0
        For Each r In ws.UsedRange.Rows
            If r.Hidden Then hr = hr + 1
        Next
        For Each r In ws.UsedRange.Columns
            If r.Hidden Then hc = hc + 1
        Next
        If hr + hc > 0 Then nRC = nRC + 1: rows.Add SKW_Row(ws.Name, "Hidden rows and columns", hr & " row(s), " & hc & " column(s) in the used range")
    Next
    For Each nm In wb.Names
        If Not nm.Visible Then nN = nN + 1: rows.Add SKW_Row(nm.Name, "Hidden name", nm.RefersTo)
    Next
    SKW_Status ""
    SKW_ReportFor "Book Hidden Content", "Hidden sheets " & nS & ", sheets with hidden rows or columns " & nRC & ", hidden names " & nN & ". Changes nothing.", _
        Array("Sheet or name", "Finding", "Detail"), rows, wb.Name
    SKX_BookHiddenContent = "Hidden sheets " & nS & ", sheets with hidden rows or columns " & nRC & ", hidden names " & nN & "."
End Function

Public Function SKX_BookLeftovers() As String
    Dim wb As Object, msg As String, rows As New Collection, ws As Object, c As Object, cel As Object, ph As Variant, nC As Long, nP As Long, k As Long, n As Long
    msg = StartBook("Workbook leftovers", wb)
    If Len(msg) > 0 Then SKX_BookLeftovers = msg: Exit Function
    For Each ws In wb.Worksheets
        For Each c In ws.Comments
            nC = nC + 1
            rows.Add SKW_Row(ws.Name, Addr(c.Parent), "Note", c.Text)
        Next
        On Error Resume Next
        n = 0: n = ws.CommentsThreaded.Count   ' read first: a failed loop bound under Resume Next can loop forever
        For k = 1 To n
            nC = nC + 1
            rows.Add SKW_Row(ws.Name, Addr(ws.CommentsThreaded(k).Parent), "Comment", ws.CommentsThreaded(k).Text)
        Next
        On Error GoTo 0
        For Each cel In TextCells(ws)
            For Each ph In SKW_Placeholders()
                If InStr(1, cel.Value, CStr(ph), vbTextCompare) > 0 Then
                    nP = nP + 1
                    rows.Add SKW_Row(ws.Name, Addr(cel), "Placeholder " & ph, cel.Value)
                    Exit For
                End If
            Next
        Next
    Next
    SKW_Status ""
    SKW_ReportFor "Book Leftovers", "Notes and comments " & nC & ", placeholders " & nP & ". Book Delivery Copy saves a copy without notes and comments. Changes nothing.", _
        Array("Sheet", "Cell", "Finding", "Text"), rows, wb.Name
    SKX_BookLeftovers = "Notes and comments " & nC & ", placeholders " & nP & "."
End Function

Public Function SKX_BookBannedPhrases() As String
    Dim wb As Object, msg As String, rows As New Collection, ws As Object, cel As Object, ru As Object, n As Long, t As String
    msg = StartBook("Workbook banned phrases", wb)
    If Len(msg) > 0 Then SKX_BookBannedPhrases = msg: Exit Function
    For Each ws In wb.Worksheets
        For Each cel In TextCells(ws)
            t = cel.Value
            For Each ru In SKW_BannedPhrases()
                If InStr(1, t, ru("phrase"), vbTextCompare) > 0 Then
                    n = n + 1
                    rows.Add SKW_Row(ws.Name, Addr(cel), ru("phrase"), IIf(ru.Exists("replacement"), ru("replacement"), ""), t)
                End If
            Next
        Next
    Next
    SKW_Status ""
    SKW_ReportFor "Book Banned Phrases", n & " finding(s) against " & SKW_BannedPhrases().Count & " banned phrase(s), in typed cell text. Changes nothing.", _
        Array("Sheet", "Cell", "Phrase", "Use instead", "Text"), rows, wb.Name
    SKX_BookBannedPhrases = n & " banned phrase finding(s)."
End Function

Public Function SKX_BookAcronymCheck() As String
    Dim wb As Object, msg As String, rows As New Collection, ws As Object, cel As Object, list As Object, info As Object, w As Variant, k As Variant, e As Variant, t As String
    msg = StartBook("Workbook acronyms", wb)
    If Len(msg) > 0 Then SKX_BookAcronymCheck = msg: Exit Function
    Set list = SKW_AcronymList()
    Set info = CreateObject("Scripting.Dictionary")
    For Each ws In wb.Worksheets
        For Each cel In TextCells(ws)
            t = cel.Value
            For Each w In SKX_Tokens(t)
                If Not info.Exists(w) Then info(w) = Array(0, ws.Name & "!" & Addr(cel), False)
                e = info(w)
                e(0) = e(0) + 1
                If InStr(t, "(" & w & ")") > 0 Then e(2) = True
                info(w) = e
            Next
        Next
    Next
    For Each k In info.Keys
        e = info(k)
        rows.Add SKW_Row(k, e(0), e(1), IIf(e(2), "Yes", "No"), IIf(list.Exists(k), "Yes", "No"), IIf(list.Exists(k), list(k), ""))
    Next
    SKW_Status ""
    SKW_ReportFor "Book Acronym Check (text only)", info.Count & " acronym(s) in typed cell text. Charts and pictures are not read. Changes nothing.", _
        Array("Acronym", "Uses", "First cell", "Spelled out somewhere", "On the team list", "Team meaning"), rows, wb.Name
    SKX_BookAcronymCheck = info.Count & " acronym(s) (text only)."
End Function

Public Function SKX_Tokens(ByVal t As String) As Collection
    Dim c As New Collection, i As Long, ch As String, cur As String, j As Long, roman As Boolean
    t = t & " "
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch Like "[A-Z0-9&]" Then
            cur = cur & ch
        Else
            If ch Like "[a-z]" Then
                cur = ""
                Do While i < Len(t) And Mid$(t, i + 1, 1) Like "[A-Za-z0-9&]"
                    i = i + 1
                Loop
            Else
                If Len(cur) >= 2 And Len(cur) <= 10 And cur Like "[A-Z]*" And cur Like "*[A-Z]*[A-Z]*" Then
                    roman = True
                    For j = 1 To Len(cur)
                        If InStr("IVXLCDM", Mid$(cur, j, 1)) = 0 Then roman = False: Exit For
                    Next
                    If Not roman Then c.Add cur
                End If
                cur = ""
            End If
        End If
    Next
    Set SKX_Tokens = c
End Function

Public Function SKX_BookFontAudit() As String
    Dim wb As Object, msg As String, rows As New Collection, ws As Object, cel As Object, r As Object, fonts As Object, allowed As New Collection
    Dim nm As String, k As Variant, e As Variant, nBad As Long, n As Long, t As Variant
    msg = StartBook("Workbook fonts", wb)
    If Len(msg) > 0 Then SKX_BookFontAudit = msg: Exit Function
    For Each r In SKW_Rows("Fonts")
        If r.Exists("font") Then
            If Len(r("font")) > 0 Then
                If Not r.Exists("allowed") Then
                    allowed.Add r("font")
                ElseIf SKW_IsYes(r("allowed")) Or Len(r("allowed")) = 0 Then
                    allowed.Add r("font")
                End If
            End If
        End If
    Next
    Set fonts = CreateObject("Scripting.Dictionary")
    For Each ws In wb.Worksheets
        For Each t In Array(2, -4123)   ' constants, formulas
            Set r = Nothing
            On Error Resume Next
            Set r = ws.UsedRange.SpecialCells(t)
            On Error GoTo 0
            If Not r Is Nothing Then
                For Each cel In r.Cells
                    n = n + 1
                    If n > MAXCELLS Then Exit For
                    nm = cel.Font.Name
                    If Len(nm) > 0 Then
                        If fonts.Exists(nm) Then
                            e = fonts(nm): e(0) = e(0) + 1: fonts(nm) = e
                        Else
                            fonts(nm) = Array(1, ws.Name & "!" & Addr(cel))
                        End If
                    End If
                Next
            End If
        Next
    Next
    For Each k In fonts.Keys
        e = fonts(k)
        If allowed.Count > 0 And Not SKW_InList(allowed, CStr(k)) Then
            nBad = nBad + 1
            rows.Add SKW_Row("Not allowed", k, e(0), e(1))
        Else
            rows.Add SKW_Row(IIf(allowed.Count > 0, "Allowed", "Used"), k, e(0), e(1))
        End If
    Next
    SKW_Status ""
    SKW_ReportFor "Book Font Audit", fonts.Count & " font(s) in filled cells, " & nBad & " not on the team list" & IIf(allowed.Count = 0, " (no Fonts list, so none are flagged)", "") & _
        IIf(n > MAXCELLS, ". The first " & MAXCELLS & " cells were checked", "") & ". Changes nothing.", Array("Rule", "Font", "Cells", "First cell"), rows, wb.Name
    SKX_BookFontAudit = fonts.Count & " font(s) used, " & nBad & " not allowed."
End Function

Public Function SKX_BookMergedCells() As String
    Dim wb As Object, msg As String, rows As New Collection, ws As Object, cel As Object, seen As Object, n As Long, a As String, k As Long
    msg = StartBook("Merged cells", wb)
    If Len(msg) > 0 Then SKX_BookMergedCells = msg: Exit Function
    For Each ws In wb.Worksheets
        Set seen = CreateObject("Scripting.Dictionary")
        k = 0
        For Each cel In ws.UsedRange.Cells
            k = k + 1
            If k > MAXCELLS * 5 Then Exit For
            If cel.MergeCells Then
                a = cel.MergeArea.Address(False, False)
                If Not seen.Exists(a) Then
                    seen(a) = True
                    n = n + 1
                    If n <= 500 Then rows.Add SKW_Row(ws.Name, a, SKW_Excerpt(CStr(cel.MergeArea.Cells(1, 1).Text), 60))
                End If
            End If
        Next
    Next
    SKW_Status ""
    SKW_ReportFor "Book Merged Cells", n & " merged area(s). Center Across Selection gives the same look without merging. Changes nothing.", _
        Array("Sheet", "Merged area", "Text"), rows, wb.Name
    SKX_BookMergedCells = n & " merged area(s)."
End Function

Public Function SKX_BookDeliveryCopy(ByVal target As String) As String
    Dim wb As Object, c As Object, xl As Object, ws As Object, links As Variant, i As Long, n As Long, nC As Long, nL As Long, ext As String, alerts As Boolean
    Set SKW_LastReport = Nothing
    Set wb = SKX_Book()
    If wb Is Nothing Then SKX_BookDeliveryCopy = mWhy: Exit Function
    If Len(wb.Path) = 0 Then SKX_BookDeliveryCopy = "Save the workbook first, so the copy has a folder to go in.": Exit Function
    If LCase$(Left$(wb.FullName, 4)) = "http" Then SKX_BookDeliveryCopy = "The workbook is opened from the web. Open it from the synced folder (or save a local copy) first.": Exit Function
    ext = Mid$(wb.Name, InStrRev(wb.Name & ".", "."))
    If Len(target) = 0 Then target = wb.Path & "\" & Left$(wb.Name, InStrRev(wb.Name & ".", ".") - 1) & " - delivery" & ext
    Set xl = wb.Application
    alerts = xl.DisplayAlerts
    On Error GoTo Fail
    SKW_Status "Making the delivery copy..."
    wb.SaveCopyAs target
    xl.DisplayAlerts = False
    Set c = xl.Workbooks.Open(FileName:=target, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
    For Each ws In c.Worksheets
        nC = nC + ws.Comments.Count
        For i = ws.Comments.Count To 1 Step -1
            ws.Comments(i).Delete
        Next
        On Error Resume Next
        n = 0: n = ws.CommentsThreaded.Count   ' read first: a failed loop bound under Resume Next can loop forever
        For i = n To 1 Step -1
            ws.CommentsThreaded(i).Delete
            nC = nC + 1
        Next
        On Error GoTo Fail
    Next
    On Error Resume Next
    links = c.LinkSources(1)
    If IsArray(links) Then
        For i = LBound(links) To UBound(links)
            c.BreakLink links(i), 1
            nL = nL + 1
        Next
    End If
    c.RemoveDocumentInformation 99    ' xlRDIAll
    c.RemovePersonalInformation = True
    On Error GoTo Fail
    c.Save
    c.Close False
    xl.DisplayAlerts = alerts
    SKW_Status ""
    SKX_BookDeliveryCopy = "Saved a delivery copy: " & target & vbCr & "Taken out: " & nC & " note(s) and comment(s), " & nL & _
        " link(s) to other workbooks (now values), and the file's personal details. Hidden sheets stay: Book Hidden Content lists them. The open workbook is unchanged."
    Exit Function
Fail:
    SKX_BookDeliveryCopy = "The delivery copy could not be made: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close False
    xl.DisplayAlerts = alerts
End Function
