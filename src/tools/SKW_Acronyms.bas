Attribute VB_Name = "SKW_Acronyms"
' =====================================================================
' Stationkeeper Tools: acronyms (a quick check, not the Stationkeeper app).
'   AcronymQuickCheck   report: acronyms in the typed text, first use in running
'                       text, defined or not, on the team list or not
' Text only: pictures and images are not read. For a full check (pictures,
' customer lists, the master database), use the Stationkeeper app.
' =====================================================================
Option Explicit

' Lists every acronym in the typed text with its count, where it is first used in running text (not tables or headings), whether it is spelled out there as "Words (ACR)", and its meaning from the team list (the Acronyms sheet, or an acronym list exported from the Stationkeeper app: setting AcronymList). Text only: pictures are not read. Changes nothing.
Public Sub AcronymQuickCheck()
    SKW_Show SKW_AcronymQuickCheck()
End Sub

' Team meanings: the Acronyms sheet, then the app's export (setting AcronymList: an Excel or CSV file
' whose first row names a Term or Acronym column and a Definition or Meaning column).
Public Function SKW_AcronymList() As Object
    Dim m As Object, r As Object, p As String
    Set m = CreateObject("Scripting.Dictionary")
    m.CompareMode = 0
    For Each r In SKW_Rows("Acronyms")
        If r.Exists("acronym") Then
            If Len(r("acronym")) > 0 And Not m.Exists(r("acronym")) Then m(r("acronym")) = IIf(r.Exists("definition"), r("definition"), "")
        End If
    Next
    p = SKW_Setting("AcronymList", "")
    If Len(p) > 0 Then ReadExport p, m
    Set SKW_AcronymList = m
End Function

Private Sub ReadExport(ByVal p As String, ByVal m As Object)
    Dim xl As Object, wb As Object, v As Variant, r As Long, c As Long, ct As Long, cd As Long, h As String, t As String
    If Len(Dir$(p)) = 0 Then Exit Sub
    On Error GoTo Done
    Set xl = CreateObject("Excel.Application")
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(FileName:=p, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    v = wb.Worksheets(1).UsedRange.Value
    If IsArray(v) Then
        For c = LBound(v, 2) To UBound(v, 2)
            h = LCase$(Trim$(CStr(v(LBound(v, 1), c))))
            If ct = 0 And (h = "term" Or h = "acronym" Or h = "acronym or abbreviation" Or h = "abbreviation") Then ct = c
            If cd = 0 And (h = "definition" Or h = "selected definition" Or h = "meaning" Or h = "long form" Or h = "stands for") Then cd = c
        Next
        If ct > 0 Then
            For r = LBound(v, 1) + 1 To UBound(v, 1)
                t = Trim$(CStr(v(r, ct)))
                If Len(t) > 0 And Not m.Exists(t) Then
                    If cd > 0 Then m(t) = Trim$(CStr(v(r, cd))) Else m(t) = ""
                End If
            Next
        End If
    End If
Done:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
End Sub

Private Function IsRoman(ByVal t As String) As Boolean
    Dim i As Long
    For i = 1 To Len(t)
        If InStr("IVXLCDM", Mid$(t, i, 1)) = 0 Then Exit Function
    Next
    IsRoman = True
End Function

' True when the acronym at r is written "Words (ACR)".
Private Function DefinedHere(ByVal d As Object, ByVal r As Object) As Boolean
    Dim b As Object, a As Object
    If r.Start < 2 Then Exit Function
    Set b = d.Range(r.Start - 1, r.Start)
    Set a = d.Range(r.End, r.End + 1)
    If b.Text = "(" And (a.Text = ")" Or (a.Text = "s" And d.Range(r.End + 1, r.End + 2).Text = ")")) Then DefinedHere = True
End Function

Public Function SKW_AcronymQuickCheck() As String
    Dim d As Object, msg As String, list As Object, ign As Collection, f As Object, t As String, info As Object, k As Variant
    Dim rows As New Collection, lastStart As Long, inTable As Boolean, heading As Boolean, e As Variant, n As Long
    Dim nTerms As Long, nNotDef As Long, nUnknown As Long, firstPg As String, defFirst As String, onList As String
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Acronym quick check", False)
    If Len(msg) > 0 Then SKW_AcronymQuickCheck = msg: Exit Function
    Set d = ActiveDocument
    Set list = SKW_AcronymList()
    Set ign = SKW_Column("Ignore", "Text")
    Set info = CreateObject("Scripting.Dictionary")
    ' term -> Array(count, first page (running text), defined at that first use, defined anywhere, first page anywhere)
    Set f = d.Content.Duplicate
    With f.Find
        .ClearFormatting
        .Text = "<[A-Z][A-Z0-9&]{1,9}>"
        .MatchWildcards = True
        .Format = False
        .Forward = True
        .Wrap = wdFindStop
    End With
    lastStart = -1
    Do While f.Find.Execute
        If f.Start = lastStart Then Exit Do
        lastStart = f.Start
        t = f.Text
        If t Like "*[A-Z]*[A-Z]*" And Not IsRoman(t) And Not SKW_InList(ign, t) Then
            If Not info.Exists(t) Then
                If info.Count >= 2000 Then Exit Do
                info(t) = Array(0, "", "", False, SKW_Page(f))
            End If
            e = info(t)
            e(0) = e(0) + 1
            If DefinedHere(d, f) Then e(3) = True
            If Len(e(1)) = 0 Then
                inTable = f.Information(wdWithInTable)
                heading = False
                On Error Resume Next
                heading = (f.Paragraphs(1).Format.OutlineLevel <= 9)
                On Error GoTo 0
                If Not inTable And Not heading Then
                    e(1) = SKW_Page(f)
                    e(2) = IIf(DefinedHere(d, f), "Yes", "No")
                End If
            End If
            info(t) = e
        End If
        f.Collapse wdCollapseEnd
        n = n + 1
        If n Mod 200 = 0 Then SKW_Status "Acronym quick check: " & info.Count & " term(s)..."
    Loop
    For Each k In info.Keys
        e = info(k)
        nTerms = nTerms + 1
        firstPg = e(1): defFirst = e(2)
        If Len(firstPg) = 0 Then firstPg = "only in tables or headings (first seen page " & e(4) & ")": defFirst = "-"
        If list.Exists(k) Then onList = list(k) Else onList = "": nUnknown = nUnknown + 1
        If defFirst = "No" Then nNotDef = nNotDef + 1
        rows.Add SKW_Row(k, e(0), firstPg, defFirst & IIf(defFirst = "No" And e(3), " (spelled out later)", ""), _
            IIf(list.Exists(k), "Yes", "No"), onList)
    Next
    SKW_Finish
    SKW_Report "Acronym Quick Check (text only)", "Text only: pictures and images are not scanned, so acronyms inside them are not listed. " & _
        "For pictures, customer lists and the master database, use the Stationkeeper app. " & nTerms & " acronym(s); " & nNotDef & _
        " not spelled out at their first use in running text (tables and headings do not count); " & nUnknown & " not on the team list (" & list.Count & _
        " meanings from the Acronyms sheet" & IIf(Len(SKW_Setting("AcronymList", "")) > 0, " and the app's export", "") & "). Changes nothing.", _
        Array("Acronym", "Uses", "First use in running text (page)", "Spelled out there", "On the team list", "Team meaning"), rows
    SKW_AcronymQuickCheck = nTerms & " acronym(s), " & nNotDef & " not spelled out at first use, " & nUnknown & " not on the team list (text only)."
End Function
