Attribute VB_Name = "SKW_Utils"
' =====================================================================
' Stationkeeper Tools: everyday utilities.
'   UnlinkFieldsSelection, LockFieldsSelection, UnlockFieldsSelection
'   RestartNumbering, ContinueNumbering, ConvertToTeamList
'   FitAllTables, GoToNextFinding, InstallToolShortcuts
'   InsertAcronymList, ResetParagraphToStyle
' =====================================================================
Option Explicit

' Turns the fields in the selected text into plain text (their current results), so they never update again. One Ctrl+Z undoes it.
Public Sub UnlinkFieldsSelection()
    SKW_Show SKW_FieldsSelection("unlink")
End Sub

' Locks the fields in the selected text so updating fields leaves their results as they are. One Ctrl+Z undoes it.
Public Sub LockFieldsSelection()
    SKW_Show SKW_FieldsSelection("lock")
End Sub

' Unlocks the fields in the selected text so they update again. One Ctrl+Z undoes it.
Public Sub UnlockFieldsSelection()
    SKW_Show SKW_FieldsSelection("unlock")
End Sub

' Starts the numbered list at the cursor again from 1. One Ctrl+Z undoes it.
Public Sub RestartNumbering()
    SKW_Show SKW_Numbering(False)
End Sub

' Continues the numbered list at the cursor from the list before it. One Ctrl+Z undoes it.
Public Sub ContinueNumbering()
    SKW_Show SKW_Numbering(True)
End Sub

' Turns the selected paragraphs that start with a typed bullet or number into the team's list styles (BulletStyle and NumberStyle settings), taking the typed marker out. One Ctrl+Z undoes it.
Public Sub ConvertToTeamList()
    SKW_Show SKW_ConvertToTeamList()
End Sub

' Fits every table wider than the margins to the page width (only the tables in the selection, if it holds any). One Ctrl+Z undoes it.
Public Sub FitAllTables()
    SKW_Show SKW_FitAllTables()
End Sub

' Selects the next finding of the last report you ran, in the document it checked. Run it again for the one after.
Public Sub GoToNextFinding()
    SKW_Status SKW_GoToNextFinding()
End Sub

' Sets keyboard shortcuts for the tools used most (Ctrl+Alt+Shift with N, U, C, R, L) in this template. Keys Word or someone else already uses are left alone and listed.
Public Sub InstallToolShortcuts()
    If SKW_IsQuiet() Then Exit Sub
    SKW_Say SKW_InstallToolShortcuts()
End Sub

' Inserts an acronym table (Acronym, Definition) at the cursor: the acronyms used in this document that have a team meaning, A to Z. One Ctrl+Z undoes it.
Public Sub InsertAcronymList()
    SKW_Show SKW_InsertAcronymList(SKW_IsYes(SKW_Setting("AcronymListUsedOnly", "Yes")))
End Sub

' Puts the paragraph at the cursor (or the selected paragraphs) back to its style: all formatting set by hand goes, bold and italic included. Character styles stay. One Ctrl+Z undoes it.
Public Sub ResetParagraphToStyle()
    SKW_Show SKW_ResetParagraphToStyle()
End Sub

' ---------------------------------------------------------------------
Private Function HasSelection() As Boolean
    If Documents.Count = 0 Then Exit Function
    HasSelection = Not (Selection.Type = wdSelectionIP Or Len(Selection.Range.Text) = 0)
End Function

Public Function SKW_FieldsSelection(ByVal what As String) As String
    Dim msg As String, r As Object, f As Object, n As Long, i As Long
    Set SKW_LastReport = Nothing
    If Not HasSelection() Then SKW_FieldsSelection = "Select the text with the fields first.": Exit Function
    msg = SKW_Begin(IIf(what = "unlink", "Unlink fields", IIf(what = "lock", "Lock fields", "Unlock fields")), True)
    If Len(msg) > 0 Then SKW_FieldsSelection = msg: Exit Function
    Set r = Selection.Range.Duplicate
    On Error GoTo Fail
    n = r.Fields.Count
    Select Case what
        Case "unlink"
            For i = r.Fields.Count To 1 Step -1
                Set f = r.Fields(i)
                If f.Type <> wdFieldTOC And f.Type <> wdFieldTOA And f.Type <> wdFieldIndex Then f.Unlink Else n = n - 1
            Next
        Case "lock", "unlock"
            ' One by one: setting Fields.Locked on the whole collection is not reliable.
            For Each f In r.Fields
                f.Locked = (what = "lock")
            Next
    End Select
    SKW_Finish
    Select Case what
        Case "unlink": SKW_FieldsSelection = n & " field(s) turned into plain text (tables of contents and indexes left as fields). One Ctrl+Z undoes it."
        Case "lock": SKW_FieldsSelection = n & " field(s) locked: updating fields leaves them as they are. One Ctrl+Z undoes it."
        Case Else: SKW_FieldsSelection = n & " field(s) unlocked. One Ctrl+Z undoes it."
    End Select
    Exit Function
Fail:
    msg = Err.Description
    SKW_Finish
    SKW_FieldsSelection = "Stopped: " & msg
End Function

Public Function SKW_Numbering(ByVal continuePrevious As Boolean) As String
    Dim msg As String, lf As Object, lt As Object, prev As Object
    Set SKW_LastReport = Nothing
    If Documents.Count > 0 Then
        If Selection.Range.ListFormat.ListType <> wdListOutlineNumbering And Selection.Range.ListFormat.ListType <> wdListSimpleNumbering _
           And Selection.Range.ListFormat.ListType <> wdListMixedNumbering Then
            SKW_Numbering = "Put the cursor in a numbered list first.": Exit Function
        End If
    End If
    msg = SKW_Begin(IIf(continuePrevious, "Continue numbering", "Restart numbering"), True)
    If Len(msg) > 0 Then SKW_Numbering = msg: Exit Function
    On Error GoTo Fail
    Set lf = Selection.Range.ListFormat
    If continuePrevious Then
        ' Continue: join the numbered list before this one (Word continues only a list of the same template).
        Set prev = PreviousList(Selection.Paragraphs(1))
        If prev Is Nothing Then
            SKW_Finish
            SKW_Numbering = "There is no numbered list before this one to continue."
            Exit Function
        End If
        lf.ApplyListTemplateWithLevel ListTemplate:=prev.Range.ListFormat.ListTemplate, ContinuePreviousList:=True, _
            ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:=wdWord10ListBehavior
    Else
        Set lt = lf.ListTemplate
        lf.ApplyListTemplateWithLevel ListTemplate:=lt, ContinuePreviousList:=False, _
            ApplyTo:=wdListApplyToThisPointForward, DefaultListBehavior:=wdWord10ListBehavior, ApplyLevel:=lf.ListLevelNumber
    End If
    SKW_Finish
    SKW_Numbering = IIf(continuePrevious, "The list continues from the one before: this item is now " & Selection.Range.ListFormat.ListString, _
        "The list starts again here: " & Selection.Range.ListFormat.ListString) & " One Ctrl+Z undoes it."
    Exit Function
Fail:
    msg = Err.Description
    SKW_Finish
    SKW_Numbering = "Stopped: " & msg
End Function

' The nearest numbered paragraph before p that belongs to another list.
Private Function PreviousList(ByVal p As Object) As Object
    Dim q As Object, mine As Long
    On Error Resume Next
    mine = p.Range.ListFormat.List.Range.Start
    Set q = p.Previous
    Do While Not q Is Nothing
        If q.Range.ListFormat.ListType = wdListSimpleNumbering Or q.Range.ListFormat.ListType = wdListOutlineNumbering Then
            If q.Range.ListFormat.List.Range.Start <> mine Then Set PreviousList = q: Exit Function
        End If
        Set q = q.Previous
    Loop
End Function

Public Function SKW_ConvertToTeamList() As String
    Dim d As Object, msg As String, p As Object, m As String, cut As Long, t As String, bs As String, ns As String
    Dim nB As Long, nN As Long, r As Object, sty As String, paras As Collection, pr As Object
    Set SKW_LastReport = Nothing
    If Not HasSelection() Then SKW_ConvertToTeamList = "Select the typed list paragraphs first (Typed Lists Finder lists them).": Exit Function
    msg = SKW_Begin("Convert to team list", True)
    If Len(msg) > 0 Then SKW_ConvertToTeamList = msg: Exit Function
    Set d = ActiveDocument
    bs = SKW_Setting("BulletStyle", "List Bullet")
    ns = SKW_Setting("NumberStyle", "List Number")
    If SKW_GetStyle(d, bs) Is Nothing Or SKW_GetStyle(d, ns) Is Nothing Then
        SKW_Finish
        SKW_ConvertToTeamList = "The list styles " & bs & " and " & ns & " must be in the document (settings BulletStyle and NumberStyle)."
        Exit Function
    End If
    On Error GoTo Fail
    ' Take the paragraphs first: changing them moves the selection.
    Set paras = New Collection
    For Each p In Selection.Range.Duplicate.Paragraphs
        paras.Add p.Range.Duplicate
    Next
    For Each pr In paras
        Set p = pr.Paragraphs(1)
        If p.Range.ListFormat.ListType = wdListNoNumbering Then
            t = p.Range.Text
            m = SKW_TypedMarker(t)
            If Len(m) > 0 Then
                ' The typed marker and the spaces or tab after it.
                cut = InStr(Replace(t, vbTab, " "), " ")
                Do While cut < Len(t) And (Mid$(t, cut + 1, 1) = " " Or Mid$(t, cut + 1, 1) = vbTab)
                    cut = cut + 1
                Loop
                Set r = d.Range(p.Range.Start, p.Range.Start + cut)
                r.Delete
                sty = IIf(Left$(m, 6) = "bullet", bs, ns)
                p.Range.Style = d.Styles(sty)
                If sty = bs Then nB = nB + 1 Else nN = nN + 1
            End If
        End If
    Next
    SKW_Finish
    SKW_ConvertToTeamList = "Converted " & nB & " bullet(s) to " & bs & " and " & nN & " numbered item(s) to " & ns & ". One Ctrl+Z undoes it."
    Exit Function
Fail:
    msg = Err.Description
    SKW_Finish
    SKW_ConvertToTeamList = "Stopped: " & msg & ". Press Ctrl+Z to undo what was done."
End Function

Public Function SKW_FitAllTables() As String
    Dim d As Object, msg As String, t As Object, tw As Single, w As Single, c As Object, n As Long, scope As Object, total As Long
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Fit all tables", True)
    If Len(msg) > 0 Then SKW_FitAllTables = msg: Exit Function
    Set d = ActiveDocument
    If Selection.Tables.Count > 0 And HasSelection() Then Set scope = Selection.Range Else Set scope = d.Content
    On Error GoTo Fail
    For Each t In scope.Tables
        total = total + 1
        tw = 0: w = 0
        On Error Resume Next
        With t.Range.Sections(1).PageSetup
            tw = .PageWidth - .LeftMargin - .RightMargin - .Gutter
        End With
        For Each c In t.Range.Cells
            If c.RowIndex > 1 Then Exit For
            w = w + c.Width
        Next
        On Error GoTo Fail
        If tw > 0 And w > tw + 2 Then
            ' Scale every cell by the same factor (AutoFit to window does not always resize fixed columns).
            t.Rows.WrapAroundText = False
            t.AllowAutoFit = False
            For Each c In t.Range.Cells
                c.SetWidth ColumnWidth:=c.Width * tw / w, RulerStyle:=wdAdjustNone
            Next
            t.PreferredWidthType = wdPreferredWidthPercent
            t.PreferredWidth = 100
            n = n + 1
        End If
    Next
    SKW_Finish
    SKW_FitAllTables = "Fitted " & n & " of " & total & " table(s) to the page width; the others already fit. One Ctrl+Z undoes it."
    Exit Function
Fail:
    msg = Err.Description
    SKW_Finish
    SKW_FitAllTables = "Stopped: " & msg & ". Press Ctrl+Z to undo what was done."
End Function

Public Function SKW_GoToNextFinding() As String
    Dim d As Object, i As Long, n As Long, r As Variant, t As String, f As Object, tried As Long, src As Object
    If SKW_FindRows Is Nothing Then SKW_GoToNextFinding = "Run a report first; then this walks through its findings.": Exit Function
    If SKW_FindRows.Count = 0 Or SKW_FindCol < 0 Then SKW_GoToNextFinding = "The last report (" & SKW_FindTitle & ") has nothing to go to.": Exit Function
    On Error Resume Next
    For Each d In Documents
        If d.FullName = SKW_FindDoc Then Set src = d
    Next
    On Error GoTo 0
    If src Is Nothing Then SKW_GoToNextFinding = "The document the last report checked is closed.": Exit Function
    src.Activate
    n = SKW_FindRows.Count
    Do While tried < n
        SKW_FindPos = SKW_FindPos + 1
        If SKW_FindPos > n Then SKW_FindPos = 1
        tried = tried + 1
        r = SKW_FindRows(SKW_FindPos)
        t = SearchText(CStr(r(SKW_FindCol)))
        If Len(t) >= 3 Then
            Set f = src.Content
            With f.Find
                .ClearFormatting
                .Text = t
                .MatchCase = True
                .MatchWildcards = False
                .Forward = True
                .Wrap = wdFindStop
            End With
            If f.Find.Execute Then
                f.Select
                SKW_GoToNextFinding = "Finding " & SKW_FindPos & " of " & n & " (" & SKW_FindTitle & "): " & SKW_Excerpt(CStr(r(LBound(r))), 40) & _
                    " " & SKW_Excerpt(CStr(r(LBound(r) + 1)), 80)
                Exit Function
            End If
        End If
    Loop
    SKW_GoToNextFinding = "None of the findings could be found in the text any more (fixed or moved)."
End Function

' The part of a report's text worth searching for: up to 60 characters, without the "..." it was cut with.
Private Function SearchText(ByVal s As String) As String
    s = SKW_Excerpt(Replace(s, ChrW(183), " "), 200)   ' raw cell text: drop cell marks and line breaks
    If Right$(s, 3) = "..." Then s = Left$(s, Len(s) - 3)
    s = Trim$(s)
    If Len(s) > 60 Then s = Left$(s, 60)
    If InStr(s, "^") > 0 Then s = Replace(s, "^", "^^")
    SearchText = Trim$(s)
End Function

Public Function SKW_InstallToolShortcuts() As String
    Dim keys As Variant, k As Variant, code As Long, kb As Object, done As String, skipped As String, existing As String
    keys = Array(Array(wdKeyN, "GoToNextFinding", "next finding"), Array(wdKeyU, "SafeUpdateAllFields", "safe update all fields"), _
                 Array(wdKeyC, "ClearDirectFormattingSelection", "clear direct formatting"), Array(wdKeyR, "ResetParagraphToStyle", "reset paragraph to style"), _
                 Array(wdKeyL, "ConvertToTeamList", "convert to team list"))
    On Error GoTo Fail
    CustomizationContext = MacroContainer
    For Each k In keys
        code = BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyShift, k(0))
        existing = ""
        On Error Resume Next
        Set kb = Nothing
        Set kb = FindKey(code)
        existing = kb.Command
        On Error GoTo Fail
        If Len(existing) > 0 And InStr(1, existing, CStr(k(1)), vbTextCompare) = 0 Then
            skipped = skipped & vbCr & "  Ctrl+Alt+Shift+" & Chr$(k(0)) & " is already " & existing
        Else
            KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:=CStr(k(1)), KeyCode:=code
            done = done & vbCr & "  Ctrl+Alt+Shift+" & Chr$(k(0)) & ": " & k(2)
        End If
    Next
    On Error Resume Next
    MacroContainer.Saved = False
    SKW_InstallToolShortcuts = "Shortcuts in this template:" & done & IIf(Len(skipped) > 0, vbCr & "Left alone:" & skipped, "") & vbCr & _
        "Save the template to keep them."
    Exit Function
Fail:
    SKW_InstallToolShortcuts = "The shortcuts were not set: " & Err.Description
End Function

' The acronyms used in the active document's typed text (the same search as Acronym Quick Check).
Public Function SKW_UsedAcronyms() As Object
    Dim used As Object, f As Object, lastStart As Long
    Set used = CreateObject("Scripting.Dictionary")
    Set f = ActiveDocument.Content.Duplicate
    With f.Find
        .ClearFormatting
        .Text = "<[A-Z][A-Z0-9&]{1,9}>"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
    End With
    lastStart = -1
    Do While f.Find.Execute
        If f.Start = lastStart Then Exit Do
        lastStart = f.Start
        used(f.Text) = True
        f.Collapse wdCollapseEnd
    Loop
    Set SKW_UsedAcronyms = used
End Function

Public Function SKW_InsertAcronymList(ByVal usedOnly As Boolean) As String
    Dim d As Object, msg As String, list As Object, used As Object, k As Variant, names() As String, n As Long, i As Long, j As Long
    Dim tmp As String, s As String, r As Object, t As Object
    Set SKW_LastReport = Nothing
    If Documents.Count > 0 Then
        If Selection.Information(wdWithInTable) Then SKW_InsertAcronymList = "Put the cursor outside a table first.": Exit Function
    End If
    msg = SKW_Begin("Insert acronym list", True)
    If Len(msg) > 0 Then SKW_InsertAcronymList = msg: Exit Function
    Set d = ActiveDocument
    Set list = SKW_AcronymList()
    If usedOnly Then Set used = SKW_UsedAcronyms()
    ReDim names(1 To list.Count + 1)
    For Each k In list.Keys
        If Not usedOnly Then
            n = n + 1: names(n) = k
        ElseIf used.Exists(k) Then
            n = n + 1: names(n) = k
        End If
    Next
    If n = 0 Then
        SKW_Finish
        SKW_InsertAcronymList = IIf(list.Count = 0, "There are no team meanings yet (Acronyms sheet, or the AcronymList setting for the app's export).", _
            "None of the " & list.Count & " acronyms with a team meaning is used in this document.")
        Exit Function
    End If
    For i = 2 To n
        tmp = names(i): j = i - 1
        Do While j >= 1
            If StrComp(names(j), tmp, vbTextCompare) <= 0 Then Exit Do
            names(j + 1) = names(j): j = j - 1
        Loop
        names(j + 1) = tmp
    Next
    On Error GoTo Fail
    s = "Acronym" & vbTab & "Definition"
    For i = 1 To n
        s = s & vbCr & names(i) & vbTab & Replace(Replace(list(names(i)), vbTab, " "), vbCr, " ")
    Next
    Set r = Selection.Range
    r.Collapse wdCollapseStart
    r.InsertParagraphBefore
    r.Collapse wdCollapseStart
    r.InsertBefore s
    Set t = r.ConvertToTable(Separator:=wdSeparateByTabs, NumColumns:=2)
    If Not SKW_GetStyle(d, "Table Grid") Is Nothing Then t.Range.Style = "Table Grid"
    t.Rows(1).HeadingFormat = True
    t.Rows(1).Range.Font.Bold = True
    SKW_Finish
    SKW_InsertAcronymList = "Inserted " & n & " acronym(s), A to Z" & IIf(usedOnly, ", the ones used in this document", "") & _
        ". Text only: acronyms inside pictures are not counted. One Ctrl+Z undoes it."
    Exit Function
Fail:
    msg = Err.Description
    SKW_Finish
    SKW_InsertAcronymList = "Stopped: " & msg & ". Press Ctrl+Z to undo what was done."
End Function

Public Function SKW_ResetParagraphToStyle() As String
    Dim msg As String, p As Object, n As Long, nList As Long, r As Object
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Reset paragraph to style", True)
    If Len(msg) > 0 Then SKW_ResetParagraphToStyle = msg: Exit Function
    On Error GoTo Fail
    For Each p In Selection.Paragraphs
        Set r = p.Range.Duplicate
        If r.End - r.Start > 1 Then r.MoveEnd wdCharacter, -1
        r.Font.Reset
        If p.Range.ListFormat.ListType = wdListNoNumbering Then p.Reset Else nList = nList + 1
        n = n + 1
    Next
    SKW_Finish
    SKW_ResetParagraphToStyle = n & " paragraph(s) back to their style." & _
        IIf(nList > 0, " Spacing and indents in " & nList & " list paragraph(s) were left so the numbering stays.", "") & " One Ctrl+Z undoes it."
    Exit Function
Fail:
    msg = Err.Description
    SKW_Finish
    SKW_ResetParagraphToStyle = "Stopped: " & msg & ". Press Ctrl+Z to undo what was done."
End Function
