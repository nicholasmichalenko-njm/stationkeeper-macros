' =====================================================================
' Stationkeeper Macros: the menu
' Search as you type, sort, filter by what a macro affects, read, run.
' The controls are built in code, so this form needs no designer work.
' =====================================================================
Option Explicit

Private WithEvents txtSearch As MSForms.TextBox
Private WithEvents lstMacros As MSForms.ListBox
Private WithEvents cboSort As MSForms.ComboBox
Private WithEvents cboAffects As MSForms.ComboBox
Private WithEvents btnRun As MSForms.CommandButton
Private WithEvents btnPin As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton
Private WithEvents chkKeep As MSForms.CheckBox
Private lblTitle As MSForms.Label
Private lblRun As MSForms.Label
Private txtDesc As MSForms.TextBox
Private lblInfo As MSForms.Label
Private lblStatus As MSForms.Label
Private lblCount As MSForms.Label

Private mModel As Object
Private mItems As Collection
Private mView() As Long
Private mBusy As Boolean

Private Function AddCtl(ByVal progId As String, ByVal nm As String, ByVal l As Single, ByVal t As Single, _
                        ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = Me.Controls.Add(progId, nm, True)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    On Error Resume Next
    c.Font.Name = "Segoe UI"
    c.Font.Size = 9
    Set AddCtl = c
End Function

Private Function AddLabel(ByVal nm As String, ByVal cap As String, ByVal l As Single, ByVal t As Single, _
                          ByVal w As Single, ByVal h As Single) As MSForms.Label
    Dim c As MSForms.Label
    Set c = AddCtl("Forms.Label.1", nm, l, t, w, h)
    c.Caption = cap
    Set AddLabel = c
End Function

Private Sub UserForm_Initialize()
    Dim e As Object, showUnd As Boolean, v As Variant, a As Variant, affs As New Collection
    Dim i As Long, j As Long, tmp As String, arr() As String

    mBusy = True
    Set mModel = SK_LoadModel()
    Me.Caption = mModel("look")("menuTitle") & "   |   " & ThisDocument.Name
    Me.Width = 672
    Me.Height = 448
    Me.BackColor = &HFFFFFF

    AddLabel "lblSearch", "Search", 10, 8, 60, 14
    Set lblCount = AddLabel("lblCount", "", 160, 8, 150, 14)
    lblCount.TextAlign = fmTextAlignRight
    lblCount.ForeColor = &H808080
    Set txtSearch = AddCtl("Forms.TextBox.1", "txtSearch", 10, 22, 300, 20)
    txtSearch.Font.Size = 10

    AddLabel "lblSort", "Sort by", 322, 8, 100, 14
    Set cboSort = AddCtl("Forms.ComboBox.1", "cboSort", 322, 22, 110, 20)
    cboSort.Style = fmStyleDropDownList
    cboSort.AddItem "Category"
    cboSort.AddItem "A to Z"
    cboSort.AddItem "Most used"
    cboSort.AddItem "Recent"

    AddLabel "lblAff", "What it affects", 444, 8, 160, 14
    Set cboAffects = AddCtl("Forms.ComboBox.1", "cboAffects", 444, 22, 206, 20)
    cboAffects.Style = fmStyleDropDownList

    Set lstMacros = AddCtl("Forms.ListBox.1", "lstMacros", 10, 50, 300, 318)
    lstMacros.ColumnCount = 2
    lstMacros.ColumnWidths = "190 pt;95 pt"
    lstMacros.IntegralHeight = False

    Set lblTitle = AddLabel("lblTitle", "", 322, 50, 328, 20)
    lblTitle.Font.Size = 12
    lblTitle.Font.Bold = True
    Set lblRun = AddLabel("lblRun", "", 322, 72, 328, 14)
    lblRun.ForeColor = &H808080
    Set txtDesc = AddCtl("Forms.TextBox.1", "txtDesc", 322, 90, 328, 120)
    txtDesc.MultiLine = True
    txtDesc.WordWrap = True
    txtDesc.Locked = True
    txtDesc.ScrollBars = fmScrollBarsVertical
    txtDesc.SpecialEffect = fmSpecialEffectFlat
    txtDesc.BorderStyle = fmBorderStyleSingle
    txtDesc.BorderColor = &HD0D0D0
    txtDesc.Font.Size = 10
    Set lblInfo = AddLabel("lblInfo", "", 322, 218, 328, 112)
    lblInfo.WordWrap = True

    Set btnRun = AddCtl("Forms.CommandButton.1", "btnRun", 322, 340, 90, 26)
    btnRun.Caption = "Run"
    btnRun.Default = True
    Set btnPin = AddCtl("Forms.CommandButton.1", "btnPin", 418, 340, 90, 26)
    btnPin.Caption = "Pin to top"
    Set chkKeep = AddCtl("Forms.CheckBox.1", "chkKeep", 516, 344, 134, 18)
    chkKeep.Caption = "Keep this menu open"
    chkKeep.BackStyle = fmBackStyleTransparent

    Set lblStatus = AddLabel("lblStatus", "Type to search. Up and Down move, Enter runs, Esc closes.", 10, 380, 450, 14)
    lblStatus.ForeColor = &H808080
    Set btnClose = AddCtl("Forms.CommandButton.1", "btnClose", 570, 374, 80, 24)
    btnClose.Caption = "Close"
    btnClose.Cancel = True

    ' What the menu shows: macros in the template that the admin put in the menu.
    Set mItems = New Collection
    showUnd = (mModel("look")("showUndescribed") = "1")
    For Each e In mModel("entries")
        If e("found") Then
            If e("shown") = "1" Then
                mItems.Add e
            ElseIf e("shown") = "" And showUnd Then
                mItems.Add e
            End If
        End If
    Next

    ' Affects filter: every value in use, A to Z.
    For Each e In mItems
        For Each a In Split(e("affects"), ",")
            If Len(Trim$(a)) > 0 Then SK_AddUnique affs, Trim$(a)
        Next
    Next
    If affs.Count > 0 Then
        ReDim arr(1 To affs.Count)
        For i = 1 To affs.Count: arr(i) = affs(i): Next
        For i = 2 To affs.Count
            tmp = arr(i): j = i - 1
            Do While j >= 1
                If LCase$(arr(j)) <= LCase$(tmp) Then Exit Do
                arr(j + 1) = arr(j): j = j - 1
            Loop
            arr(j + 1) = tmp
        Next
    End If
    cboAffects.AddItem "Anything"
    For i = 1 To affs.Count: cboAffects.AddItem arr(i): Next
    cboAffects.ListIndex = 0

    v = SK_GetPref("sort", mModel("look")("defaultSort"))
    cboSort.ListIndex = 0
    For i = 0 To cboSort.ListCount - 1
        If LCase$(cboSort.List(i)) = LCase$(v) Then cboSort.ListIndex = i
    Next
    chkKeep.Value = (SK_GetPref("keepOpen", mModel("look")("keepOpen")) = "1")

    mBusy = False
    RefreshList
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

' ---------------------------------------------------------------------
' Search and sort
' ---------------------------------------------------------------------
Private Function Norm(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[a-z0-9]" Or AscW(ch) > 127 Then out = out & ch Else out = out & " "
    Next
    Norm = " " & out & " "
End Function

Private Function WordStarts(ByVal text As String, ByVal tok As String) As Boolean
    WordStarts = (InStr(1, Norm(text), " " & tok, vbBinaryCompare) > 0)
End Function

Private Function ScoreEntry(ByVal e As Object, ByVal q As String, ByRef why As String) As Double
    Dim toks As Variant, t As Variant, tot As Double, b As Double, w As String
    why = ""
    toks = Split(Trim$(Norm(q)), " ")
    For Each t In toks
        If Len(t) > 0 Then
            b = 0: w = ""
            If Left$(Trim$(Norm(e("title"))), Len(t)) = t Then
                b = 100: w = "title"
            ElseIf WordStarts(e("title"), t) Then
                b = 70: w = "title"
            ElseIf WordStarts(e("keywords"), t) Then
                b = 55: w = "keywords"
            ElseIf WordStarts(e("affects"), t) Then
                b = 45: w = "affects"
            ElseIf InStr(1, Norm(e("category")), t, vbBinaryCompare) > 0 Then
                b = 30: w = "category"
            ElseIf InStr(1, LCase$(e("run")), t, vbBinaryCompare) > 0 Then
                b = 25: w = "macro name"
            ElseIf InStr(1, Norm(e("description")), t, vbBinaryCompare) > 0 Then
                b = 12: w = "description"
            End If
            If b = 0 Then
                ScoreEntry = 0
                why = ""
                Exit Function
            End If
            tot = tot + b
            If InStr(", " & why & ",", ", " & w & ",") = 0 Then why = why & IIf(Len(why) > 0, ", ", "") & w
        End If
    Next
    ScoreEntry = tot
End Function

Private Function HasAffect(ByVal e As Object, ByVal aff As String) As Boolean
    Dim a As Variant
    For Each a In Split(e("affects"), ",")
        If LCase$(Trim$(a)) = LCase$(aff) Then HasAffect = True: Exit Function
    Next
End Function

Private Function CatIndex(ByVal c As String) As Long
    Dim i As Long
    For i = 1 To mModel("cats").Count
        If LCase$(mModel("cats")(i)) = LCase$(c) Then CatIndex = i: Exit Function
    Next
    CatIndex = 999
End Function

Private Sub RefreshList()
    Dim q As String, aff As String, mode As String, n As Long, cnt As Long, i As Long, j As Long
    Dim e As Object, s As Double, w As String, keyTmp As String, idxTmp As Long, whyTmp As String
    Dim keys() As String, idx() As Long, why() As String, col2 As String, pinned As Boolean, lastRun As String

    If mBusy Then Exit Sub
    q = Trim$(txtSearch.Text)
    If cboAffects.ListIndex > 0 Then aff = cboAffects.Text
    mode = cboSort.Text
    n = mItems.Count
    ReDim keys(1 To n + 1): ReDim idx(1 To n + 1): ReDim why(1 To n + 1)

    For i = 1 To n
        Set e = mItems(i)
        If Len(aff) = 0 Or HasAffect(e, aff) Then
            s = 0: w = ""
            If Len(q) > 0 Then s = ScoreEntry(e, q, w)
            If Len(q) = 0 Or s > 0 Then
                cnt = cnt + 1
                idx(cnt) = i
                why(cnt) = w
                pinned = SK_IsPinned(e("run"))
                If Len(q) > 0 Then
                    keys(cnt) = Format$(100000 - s, "000000") & "|" & LCase$(e("title"))
                Else
                    Select Case mode
                        Case "A to Z"
                            keys(cnt) = LCase$(e("title"))
                        Case "Most used"
                            keys(cnt) = Format$(999999 - SK_Uses(e("run")), "000000") & "|" & LCase$(e("title"))
                        Case "Recent"
                            lastRun = SK_LastRun(e("run"))
                            If Len(lastRun) = 0 Then
                                keys(cnt) = "1|" & LCase$(e("title"))
                            Else
                                keys(cnt) = "0|" & Format$(1000000 - CDbl(CDate(lastRun)), "0000000.000000")
                            End If
                        Case Else
                            keys(cnt) = Format$(CatIndex(e("category")), "000") & "|" & LCase$(e("title"))
                    End Select
                    keys(cnt) = IIf(pinned, "0", "1") & keys(cnt)
                End If
            End If
        End If
    Next

    ' Insertion sort (menus hold a few hundred macros at most).
    For i = 2 To cnt
        keyTmp = keys(i): idxTmp = idx(i): whyTmp = why(i)
        j = i - 1
        Do While j >= 1
            If keys(j) <= keyTmp Then Exit Do
            keys(j + 1) = keys(j): idx(j + 1) = idx(j): why(j + 1) = why(j)
            j = j - 1
        Loop
        keys(j + 1) = keyTmp: idx(j + 1) = idxTmp: why(j + 1) = whyTmp
    Next

    lstMacros.Clear
    If cnt > 0 Then ReDim mView(0 To cnt - 1) Else ReDim mView(0 To 0)
    For i = 1 To cnt
        Set e = mItems(idx(i))
        If Len(q) > 0 Then
            col2 = "in " & why(i)
        ElseIf SK_IsPinned(e("run")) Then
            col2 = "Pinned"
        Else
            Select Case mode
                Case "Most used": col2 = SK_Uses(e("run")) & " runs"
                Case "Recent"
                    lastRun = SK_LastRun(e("run"))
                    If Len(lastRun) > 0 Then col2 = Left$(lastRun, 16) Else col2 = "not run yet"
                Case Else
                    col2 = e("category")
                    If Len(col2) = 0 Then col2 = "Uncategorized"
            End Select
        End If
        lstMacros.AddItem e("title")
        lstMacros.List(i - 1, 1) = col2
        mView(i - 1) = idx(i)
    Next

    lblCount.Caption = cnt & " of " & n & " macros"
    If cnt > 0 Then lstMacros.ListIndex = 0
    ShowDetail
End Sub

Private Function Dash(ByVal s As String, ByVal alt As String) As String
    If Len(Trim$(s)) > 0 Then Dash = s Else Dash = alt
End Function

Private Sub ShowDetail()
    Dim e As Object, info As String
    If lstMacros.ListIndex < 0 Then
        lblTitle.Caption = ""
        lblRun.Caption = ""
        If mItems.Count = 0 Then
            txtDesc.Text = "No macros are in the menu yet." & vbCrLf & vbCrLf & _
                "Run StationkeeperMacrosAdmin to describe this template's macros and choose which ones show here."
        Else
            txtDesc.Text = "No macro matches. Try fewer letters, another word for what it does, or set Affects to Anything."
        End If
        lblInfo.Caption = ""
        btnRun.Enabled = False
        btnPin.Enabled = False
        Exit Sub
    End If
    Set e = mItems(mView(lstMacros.ListIndex))
    lblTitle.Caption = e("title")
    lblRun.Caption = e("run")
    txtDesc.Text = Dash(e("description"), "No description yet.")
    info = "Category:  " & Dash(e("category"), "Uncategorized") & vbCrLf & _
           "What it affects:  " & Dash(e("affects"), "not set") & vbCrLf & _
           "Works on:  " & Dash(e("workson"), "not set") & vbCrLf & _
           "Does it change the document?  " & Dash(e("changes"), "not set")
    lblInfo.Caption = info
    btnPin.Caption = IIf(SK_IsPinned(e("run")), "Unpin", "Pin to top")
    btnRun.Enabled = True
    btnPin.Enabled = True
End Sub

Private Sub SelectRun(ByVal run As String)
    Dim i As Long
    For i = 0 To lstMacros.ListCount - 1
        If mItems(mView(i))("run") = run Then
            lstMacros.ListIndex = i
            Exit Sub
        End If
    Next
End Sub

Private Sub RunSelected()
    Dim e As Object, keep As Boolean, ok As Boolean, run As String
    If lstMacros.ListIndex < 0 Then Exit Sub
    Set e = mItems(mView(lstMacros.ListIndex))
    If mModel("look")("confirmChanges") = "1" And LCase$(Left$(e("changes"), 3)) = "yes" Then
        If MsgBox(e("title") & " changes the document. Run it?", vbYesNo + vbQuestion, SK_APP) <> vbYes Then Exit Sub
    End If
    keep = chkKeep.Value
    run = e("run")
    If Not keep Then Me.Hide
    ok = SK_RunMacro(run)
    If keep Then
        lblStatus.Caption = IIf(ok, "Ran " & e("title") & " at " & Format$(Now, "h:nn am/pm") & ".", _
                                    e("title") & " stopped with an error.")
        RefreshList
        SelectRun run
    Else
        Unload Me
    End If
End Sub

' ---------------------------------------------------------------------
' Events
' ---------------------------------------------------------------------
Private Sub txtSearch_Change()
    RefreshList
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 40 Then
        If lstMacros.ListIndex < lstMacros.ListCount - 1 Then lstMacros.ListIndex = lstMacros.ListIndex + 1
        KeyCode = 0
    ElseIf KeyCode = 38 Then
        If lstMacros.ListIndex > 0 Then lstMacros.ListIndex = lstMacros.ListIndex - 1
        KeyCode = 0
    End If
End Sub

Private Sub lstMacros_Click()
    If Not mBusy Then ShowDetail
End Sub

Private Sub lstMacros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    RunSelected
End Sub

Private Sub cboSort_Change()
    If mBusy Then Exit Sub
    SK_SetPref "sort", cboSort.Text
    RefreshList
End Sub

Private Sub cboAffects_Change()
    RefreshList
End Sub

Private Sub chkKeep_Click()
    If mBusy Then Exit Sub
    SK_SetPref "keepOpen", IIf(chkKeep.Value, "1", "0")
End Sub

Private Sub btnRun_Click()
    RunSelected
End Sub

Private Sub btnPin_Click()
    Dim run As String
    If lstMacros.ListIndex < 0 Then Exit Sub
    run = mItems(mView(lstMacros.ListIndex))("run")
    SK_TogglePin run
    RefreshList
    SelectRun run
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' ---------------------------------------------------------------------
' Used by the self-test only
' ---------------------------------------------------------------------
Public Function TestSearch(ByVal q As String) As String
    Dim s As String, i As Long
    txtSearch.Text = q
    s = lstMacros.ListCount & ":"
    For i = 0 To lstMacros.ListCount - 1
        s = s & " [" & lstMacros.List(i, 0) & " / " & lstMacros.List(i, 1) & "]"
    Next
    TestSearch = s
End Function
