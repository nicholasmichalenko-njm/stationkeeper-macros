' =====================================================================
' Stationkeeper Macros: the admin
' Describe, sort and show or hide the macros already in this template.
' Nothing here edits your code. Saving writes the catalog into the
' template after a dated backup, and only after you tick Ready to save.
' The controls are built in code, so this form needs no designer work.
' =====================================================================
Option Explicit

Private mp As MSForms.MultiPage
Private lblHeader As MSForms.Label

' Macros page
Private WithEvents txtFind As MSForms.TextBox
Private WithEvents cboFilter As MSForms.ComboBox
Private WithEvents lstEntries As MSForms.ListBox
Private lblMacro As MSForms.Label
Private WithEvents txtTitle As MSForms.TextBox
Private WithEvents cboCategory As MSForms.ComboBox
Private WithEvents cboWorksOn As MSForms.ComboBox
Private WithEvents txtDescription As MSForms.TextBox
Private WithEvents btnUseComment As MSForms.CommandButton
Private WithEvents txtKeywords As MSForms.TextBox
Private WithEvents lstAffects As MSForms.ListBox
Private WithEvents cboChanges As MSForms.ComboBox
Private WithEvents chkShown As MSForms.CheckBox
Private lblComment As MSForms.Label

' Categories page
Private WithEvents lstCats As MSForms.ListBox
Private WithEvents txtCatName As MSForms.TextBox
Private WithEvents btnCatAdd As MSForms.CommandButton
Private WithEvents btnCatRename As MSForms.CommandButton
Private WithEvents btnCatUp As MSForms.CommandButton
Private WithEvents btnCatDown As MSForms.CommandButton
Private WithEvents btnCatDelete As MSForms.CommandButton

' Look page
Private WithEvents txtMenuTitle As MSForms.TextBox
Private WithEvents cboDefaultSort As MSForms.ComboBox
Private WithEvents chkShowUnd As MSForms.CheckBox
Private WithEvents chkKeepOpen As MSForms.CheckBox
Private WithEvents chkConfirm As MSForms.CheckBox
Private WithEvents btnShortcut As MSForms.CommandButton
Private WithEvents btnButton As MSForms.CommandButton

' Check page
Private WithEvents btnFillComments As MSForms.CommandButton
Private WithEvents btnShowDescribed As MSForms.CommandButton
Private WithEvents btnDropMissing As MSForms.CommandButton
Private WithEvents btnExportCsv As MSForms.CommandButton
Private WithEvents btnImportCsv As MSForms.CommandButton
Private WithEvents btnReadBas As MSForms.CommandButton

' History page
Private txtHistory As MSForms.TextBox

' Bottom bar
Private lblPending As MSForms.Label
Private WithEvents chkReady As MSForms.CheckBox
Private WithEvents btnSave As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton

Private mModel As Object
Private mView() As Long
Private mLoading As Boolean
Private mPending As Collection
Private mEdited As Object
Private mCur As Object

' ---------------------------------------------------------------------
' Building the form
' ---------------------------------------------------------------------
Private Function AddTo(ByVal parent As Object, ByVal progId As String, ByVal nm As String, _
                       ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = parent.Controls.Add(progId, nm, True)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    On Error Resume Next
    c.Font.Name = "Segoe UI"
    c.Font.Size = 9
    Set AddTo = c
End Function

Private Function Lbl(ByVal parent As Object, ByVal nm As String, ByVal cap As String, ByVal l As Single, _
                     ByVal t As Single, ByVal w As Single, ByVal h As Single) As MSForms.Label
    Dim c As MSForms.Label
    Set c = AddTo(parent, "Forms.Label.1", nm, l, t, w, h)
    c.Caption = cap
    c.BackStyle = fmBackStyleTransparent
    Set Lbl = c
End Function

Private Function Btn(ByVal parent As Object, ByVal nm As String, ByVal cap As String, ByVal l As Single, _
                     ByVal t As Single, ByVal w As Single, ByVal h As Single) As MSForms.CommandButton
    Dim c As MSForms.CommandButton
    Set c = AddTo(parent, "Forms.CommandButton.1", nm, l, t, w, h)
    c.Caption = cap
    Set Btn = c
End Function

Private Function Chk(ByVal parent As Object, ByVal nm As String, ByVal cap As String, ByVal l As Single, _
                     ByVal t As Single, ByVal w As Single, ByVal h As Single) As MSForms.CheckBox
    Dim c As MSForms.CheckBox
    Set c = AddTo(parent, "Forms.CheckBox.1", nm, l, t, w, h)
    c.Caption = cap
    c.BackStyle = fmBackStyleTransparent
    Set Chk = c
End Function

Private Sub UserForm_Initialize()
    Dim pg As Object, v As Variant

    mLoading = True
    Set mPending = New Collection
    Set mEdited = CreateObject("Scripting.Dictionary")
    Set mModel = SK_LoadModel()

    Me.Caption = "Stationkeeper Macros Admin   |   " & ThisDocument.Name
    Me.Width = 792
    Me.Height = 552

    Set lblHeader = Lbl(Me, "lblHeader", "", 10, 6, 764, 16)

    Set mp = AddTo(Me, "Forms.MultiPage.1", "mp", 6, 24, 772, 444)
    mp.Pages(0).Caption = "Macros"
    mp.Pages(1).Caption = "Categories"
    mp.Pages.Add "pLook", "Look"
    mp.Pages.Add "pBulk", "Bulk tools"
    mp.Pages.Add "pHistory", "History"

    ' ----- Macros -----
    Set pg = mp.Pages(0)
    Lbl pg, "lblFind", "Find", 8, 6, 60, 14
    Set txtFind = AddTo(pg, "Forms.TextBox.1", "txtFind", 8, 20, 228, 20)
    Lbl pg, "lblFilter", "Show", 244, 6, 60, 14
    Set cboFilter = AddTo(pg, "Forms.ComboBox.1", "cboFilter", 244, 20, 118, 20)
    cboFilter.Style = fmStyleDropDownList
    For Each v In Array("All", "Description missing", "Shown in menu", "Not shown in menu", "Hidden", "Macro gone")
        cboFilter.AddItem v
    Next
    ' Column headings (a ListBox filled in code cannot show its own).
    Dim hd As MSForms.Label
    Set hd = Lbl(pg, "hdMacro", "Macro", 12, 46, 180, 12): hd.Font.Bold = True
    Set hd = Lbl(pg, "hdMenu", "In menu", 202, 46, 70, 12): hd.Font.Bold = True
    Set hd = Lbl(pg, "hdDesc", "Description", 274, 46, 84, 12): hd.Font.Bold = True
    Set lstEntries = AddTo(pg, "Forms.ListBox.1", "lstEntries", 8, 60, 354, 354)
    lstEntries.ColumnCount = 3
    lstEntries.ColumnWidths = "190 pt;72 pt;72 pt"
    lstEntries.IntegralHeight = False

    Set lblMacro = Lbl(pg, "lblMacro", "", 372, 6, 386, 14)
    lblMacro.ForeColor = &H808080
    Lbl pg, "lblT", "Title in the menu", 372, 22, 200, 14
    Set txtTitle = AddTo(pg, "Forms.TextBox.1", "txtTitle", 372, 36, 386, 20)
    Lbl pg, "lblC", "Category", 372, 60, 180, 14
    Set cboCategory = AddTo(pg, "Forms.ComboBox.1", "cboCategory", 372, 74, 186, 20)
    cboCategory.Style = fmStyleDropDownCombo
    Lbl pg, "lblW", "Works on", 568, 60, 180, 14
    Set cboWorksOn = AddTo(pg, "Forms.ComboBox.1", "cboWorksOn", 568, 74, 190, 20)
    cboWorksOn.Style = fmStyleDropDownCombo
    For Each v In Array("Selection", "Selection or whole document", "Whole document", "Cursor position", _
                        "Table at the cursor", "Section at the cursor", "Template")
        cboWorksOn.AddItem v
    Next
    Lbl pg, "lblD", "Description (what it does, in plain words)", 372, 98, 260, 14
    Set btnUseComment = Btn(pg, "btnUseComment", "Use code comment", 650, 96, 108, 18)
    Set txtDescription = AddTo(pg, "Forms.TextBox.1", "txtDescription", 372, 116, 386, 78)
    txtDescription.MultiLine = True
    txtDescription.WordWrap = True
    txtDescription.EnterKeyBehavior = True
    txtDescription.ScrollBars = fmScrollBarsVertical
    Lbl pg, "lblK", "Search keywords (commas between)", 372, 198, 260, 14
    Set txtKeywords = AddTo(pg, "Forms.TextBox.1", "txtKeywords", 372, 212, 386, 20)
    Lbl pg, "lblA", "What this macro affects", 372, 236, 186, 14
    Set lstAffects = AddTo(pg, "Forms.ListBox.1", "lstAffects", 372, 250, 186, 124)
    lstAffects.ListStyle = fmListStyleOption
    lstAffects.MultiSelect = fmMultiSelectMulti
    lstAffects.IntegralHeight = False
    Lbl pg, "lblCh", "Does this change the document?", 568, 236, 190, 14
    Set cboChanges = AddTo(pg, "Forms.ComboBox.1", "cboChanges", 568, 250, 190, 20)
    cboChanges.Style = fmStyleDropDownCombo
    For Each v In Array("Yes", "No, it only reports", "No, it only changes the view", "It changes the template")
        cboChanges.AddItem v
    Next
    Set chkShown = Chk(pg, "chkShown", "Show in the menu", 568, 282, 190, 18)
    chkShown.Font.Bold = True
    Set lblComment = Lbl(pg, "lblComment", "", 372, 378, 386, 36)
    lblComment.ForeColor = &H808080
    lblComment.WordWrap = True

    ' ----- Categories -----
    Set pg = mp.Pages(1)
    Lbl pg, "lblCatInfo", "Categories in the order the menu shows them. Renaming moves every macro with it. " & _
        "Deleting moves its macros to Uncategorized.", 8, 6, 740, 28
    Set lstCats = AddTo(pg, "Forms.ListBox.1", "lstCats", 8, 38, 300, 370)
    lstCats.IntegralHeight = False
    Lbl pg, "lblCatName", "Name", 320, 38, 200, 14
    Set txtCatName = AddTo(pg, "Forms.TextBox.1", "txtCatName", 320, 52, 230, 20)
    Set btnCatAdd = Btn(pg, "btnCatAdd", "Add", 320, 80, 112, 24)
    Set btnCatRename = Btn(pg, "btnCatRename", "Rename", 438, 80, 112, 24)
    Set btnCatUp = Btn(pg, "btnCatUp", "Move up", 320, 110, 112, 24)
    Set btnCatDown = Btn(pg, "btnCatDown", "Move down", 438, 110, 112, 24)
    Set btnCatDelete = Btn(pg, "btnCatDelete", "Delete", 320, 140, 230, 24)

    ' ----- Look -----
    Set pg = mp.Pages(2)
    Lbl pg, "lblMT", "Menu title", 8, 8, 200, 14
    Set txtMenuTitle = AddTo(pg, "Forms.TextBox.1", "txtMenuTitle", 8, 22, 300, 20)
    Lbl pg, "lblDS", "The menu opens sorted by", 8, 50, 200, 14
    Set cboDefaultSort = AddTo(pg, "Forms.ComboBox.1", "cboDefaultSort", 8, 64, 160, 20)
    cboDefaultSort.Style = fmStyleDropDownList
    For Each v In Array("Category", "A to Z", "Most used", "Recent")
        cboDefaultSort.AddItem v
    Next
    Set chkShowUnd = Chk(pg, "chkShowUnd", "List macros nobody has described or hidden yet (useful while you set up)", 8, 96, 520, 18)
    Set chkKeepOpen = Chk(pg, "chkKeepOpen", "Keep the menu open after Run (each person can change this)", 8, 118, 520, 18)
    Set chkConfirm = Chk(pg, "chkConfirm", "Ask before running a macro marked as changing the document", 8, 140, 520, 18)
    Lbl pg, "lblOpen", "Ways to open the menu (saved in this template with the next save):", 8, 176, 500, 14
    Set btnShortcut = Btn(pg, "btnShortcut", "Set Alt+Shift+M", 8, 194, 170, 24)
    Set btnButton = Btn(pg, "btnButton", "Add a button on the Add-ins tab", 186, 194, 220, 24)
    Lbl pg, "lblOpen2", "Anyone can also run StationkeeperMacros from View > Macros, or add it to the Quick Access Toolbar.", _
        8, 226, 600, 14

    ' ----- Bulk tools: each button does one thing to many macros at once -----
    Set pg = mp.Pages(3)
    Lbl pg, "lblBulkInfo", "Each button works on the whole list at once. Nothing is saved until you tick Ready to save and save.", _
        8, 6, 750, 14
    Set btnFillComments = Btn(pg, "btnFillComments", "Fill blank descriptions from code comments", 8, 28, 260, 24)
    Lbl pg, "lblB1", "Copies the comment written above each macro in the code into any description that is still blank.", _
        278, 32, 480, 28
    Set btnShowDescribed = Btn(pg, "btnShowDescribed", "Show every described macro in the menu", 8, 64, 260, 24)
    Lbl pg, "lblB2", "Puts every macro that has a description into the menu. Macros you hid on purpose stay hidden.", _
        278, 68, 480, 28
    Set btnDropMissing = Btn(pg, "btnDropMissing", "Remove macros that are gone", 8, 100, 260, 24)
    Lbl pg, "lblB3", "Clears out list entries whose macro was deleted from the template (In menu shows Macro gone).", _
        278, 104, 480, 28
    Set btnExportCsv = Btn(pg, "btnExportCsv", "Export the list to Excel (CSV)", 8, 148, 260, 24)
    Lbl pg, "lblB4", "Saves every macro and its details as a CSV file, to fill in many descriptions quickly in Excel.", _
        278, 152, 480, 28
    Set btnImportCsv = Btn(pg, "btnImportCsv", "Import the list from Excel (CSV)", 8, 184, 260, 24)
    Lbl pg, "lblB5", "Reads that file back in after you save it in Excel as CSV UTF-8. Rows match on the Macro column.", _
        278, 188, 480, 28
    Set btnReadBas = Btn(pg, "btnReadBas", "Read macros from .bas files", 8, 232, 260, 24)
    Lbl pg, "lblB6", "Only needed when Word will not let the admin look inside the template. In the VBA editor, " & _
        "right-click each module > Export File, then pick the .bas file(s) here. One file or many.", 278, 236, 480, 42

    ' ----- History -----
    Set pg = mp.Pages(4)
    Set txtHistory = AddTo(pg, "Forms.TextBox.1", "txtHistory", 8, 8, 750, 400)
    txtHistory.MultiLine = True
    txtHistory.WordWrap = True
    txtHistory.ScrollBars = fmScrollBarsVertical
    txtHistory.Locked = True

    ' ----- Bottom bar -----
    Set lblPending = Lbl(Me, "lblPending", "", 10, 478, 360, 16)
    Set chkReady = Chk(Me, "chkReady", "Ready to save (a backup is made first)", 316, 476, 246, 18)
    Set btnSave = Btn(Me, "btnSave", "Save to template", 566, 472, 104, 24)
    Set btnClose = Btn(Me, "btnClose", "Close", 676, 472, 100, 24)
    btnClose.Cancel = True

    ' ----- Fill -----
    Dim lk As Object: Set lk = mModel("look")
    txtMenuTitle.Text = lk("menuTitle")
    cboDefaultSort.ListIndex = 0
    For Each v In Array(0, 1, 2, 3)
        If LCase$(cboDefaultSort.List(v)) = LCase$(lk("defaultSort")) Then cboDefaultSort.ListIndex = v
    Next
    chkShowUnd.Value = (lk("showUndescribed") = "1")
    chkKeepOpen.Value = (lk("keepOpen") = "1")
    chkConfirm.Value = (lk("confirmChanges") = "1")
    For Each v In SK_AffectsVocabulary()
        lstAffects.AddItem v
    Next
    AddExtraAffects
    cboFilter.ListIndex = 0
    FillCats
    FillCategoryCombo
    FillHistory
    mLoading = False
    FillList
    UpdateHeader
    UpdatePending
    If Not mModel("scanAvailable") Then
        mp.Value = 3
        If Not SK_Quiet Then
            MsgBox SK_NoAccessMessage(), vbInformation, SK_APP
        End If
    End If
End Sub

' Affects values typed elsewhere (or imported) that are not in the standard list.
Private Sub AddExtraAffects()
    Dim e As Object, a As Variant, i As Long, there As Boolean
    For Each e In mModel("entries")
        For Each a In Split(e("affects"), ",")
            a = Trim$(a)
            If Len(a) > 0 Then
                there = False
                For i = 0 To lstAffects.ListCount - 1
                    If LCase$(lstAffects.List(i)) = LCase$(a) Then there = True
                Next
                If Not there Then lstAffects.AddItem a
            End If
        Next
    Next
End Sub

' ---------------------------------------------------------------------
' Lists
' ---------------------------------------------------------------------
Private Sub FillList(Optional ByVal keepRun As String = "")
    Dim e As Object, i As Long, n As Long, cnt As Long, f As String, q As String, st As String, okRow As Boolean
    Dim keys() As String, idx() As Long, j As Long, kt As String, it As Long

    q = LCase$(Trim$(txtFind.Text))
    f = cboFilter.Text
    n = mModel("entries").Count
    ReDim keys(1 To n + 1): ReDim idx(1 To n + 1)
    For i = 1 To n
        Set e = mModel("entries")(i)
        st = SK_EntryStatus(e)
        Select Case f
            Case "Description missing": okRow = (e("found") And Len(Trim$(e("description"))) = 0)
            Case "Shown in menu": okRow = (e("found") And e("shown") = "1")
            Case "Not shown in menu": okRow = (e("found") And e("shown") <> "1")
            Case "Hidden": okRow = (e("shown") = "0")
            Case "Macro gone": okRow = (Not e("found"))
            Case Else: okRow = True
        End Select
        If okRow And Len(q) > 0 Then
            okRow = (InStr(LCase$(e("title") & " " & e("run") & " " & e("category") & " " & e("description") & _
                     " " & e("keywords")), q) > 0)
        End If
        If okRow Then
            cnt = cnt + 1
            idx(cnt) = i
            keys(cnt) = LCase$(e("title"))
        End If
    Next
    For i = 2 To cnt
        kt = keys(i): it = idx(i): j = i - 1
        Do While j >= 1
            If keys(j) <= kt Then Exit Do
            keys(j + 1) = keys(j): idx(j + 1) = idx(j): j = j - 1
        Loop
        keys(j + 1) = kt: idx(j + 1) = it
    Next

    mLoading = True
    lstEntries.Clear
    If cnt > 0 Then ReDim mView(0 To cnt - 1) Else ReDim mView(0 To 0)
    For i = 1 To cnt
        Set e = mModel("entries")(idx(i))
        lstEntries.AddItem e("title")
        lstEntries.List(i - 1, 1) = MenuCol(e)
        lstEntries.List(i - 1, 2) = DescCol(e)
        mView(i - 1) = idx(i)
    Next
    mLoading = False

    If cnt = 0 Then
        Set mCur = Nothing
        LoadFields
        Exit Sub
    End If
    lstEntries.ListIndex = 0
    If Len(keepRun) > 0 Then
        For i = 0 To cnt - 1
            If mModel("entries")(mView(i))("run") = keepRun Then lstEntries.ListIndex = i
        Next
    End If
    PickCurrent
End Sub

Private Sub KeepCategory()
    Dim before As Long
    If mCur Is Nothing Then Exit Sub
    If Len(mCur("category")) = 0 Then Exit Sub
    before = mModel("cats").Count
    SK_AddUnique mModel("cats"), mCur("category")
    If mModel("cats").Count > before Then
        FillCats
        FillCategoryCombo
    End If
End Sub

Private Sub PickCurrent()
    KeepCategory
    If lstEntries.ListIndex < 0 Then
        Set mCur = Nothing
    Else
        Set mCur = mModel("entries")(mView(lstEntries.ListIndex))
    End If
    LoadFields
End Sub

Private Sub LoadFields()
    Dim i As Long, a As Variant, en As Boolean
    mLoading = True
    en = Not (mCur Is Nothing)
    txtTitle.Enabled = en: cboCategory.Enabled = en: cboWorksOn.Enabled = en
    txtDescription.Enabled = en: txtKeywords.Enabled = en: lstAffects.Enabled = en
    cboChanges.Enabled = en: chkShown.Enabled = en: btnUseComment.Enabled = en
    For i = 0 To lstAffects.ListCount - 1
        lstAffects.Selected(i) = False
    Next
    If Not en Then
        lblMacro.Caption = "No macro matches this filter."
        txtTitle.Text = "": cboCategory.Text = "": cboWorksOn.Text = "": txtDescription.Text = ""
        txtKeywords.Text = "": cboChanges.Text = "": chkShown.Value = False
        lblComment.Caption = ""
        mLoading = False
        Exit Sub
    End If
    lblMacro.Caption = "Macro: " & mCur("run") & "     Status: " & SK_EntryStatus(mCur)
    txtTitle.Text = mCur("title")
    cboCategory.Text = mCur("category")
    cboWorksOn.Text = mCur("workson")
    txtDescription.Text = mCur("description")
    txtKeywords.Text = mCur("keywords")
    cboChanges.Text = mCur("changes")
    chkShown.Value = (mCur("shown") = "1")
    chkShown.Enabled = mCur("found")
    For Each a In Split(mCur("affects"), ",")
        For i = 0 To lstAffects.ListCount - 1
            If LCase$(lstAffects.List(i)) = LCase$(Trim$(a)) Then lstAffects.Selected(i) = True
        Next
    Next
    If Not mCur("found") Then
        lblComment.Caption = "This macro is not in the template anymore. Remove it with Bulk tools > Remove macros that are gone, " & _
                             "or put the macro back."
    ElseIf Len(mCur("comment")) > 0 Then
        lblComment.Caption = "Code comment: " & mCur("comment")
    Else
        lblComment.Caption = ""
    End If
    btnUseComment.Enabled = (Len(mCur("comment")) > 0)
    mLoading = False
End Sub

Private Sub FillCats()
    Dim v As Variant, e As Object, n As Long
    lstCats.Clear
    For Each v In mModel("cats")
        n = 0
        For Each e In mModel("entries")
            If LCase$(e("category")) = LCase$(v) Then n = n + 1
        Next
        lstCats.AddItem v & "   (" & n & ")"
    Next
End Sub

Private Sub FillCategoryCombo()
    Dim v As Variant, keep As String, wasLoading As Boolean
    wasLoading = mLoading
    mLoading = True
    keep = cboCategory.Text
    cboCategory.Clear
    For Each v In mModel("cats")
        cboCategory.AddItem v
    Next
    cboCategory.Text = keep
    mLoading = wasLoading
End Sub

Private Sub FillHistory()
    Dim i As Long, s As String
    For i = mModel("history").Count To 1 Step -1
        s = s & mModel("history")(i) & vbCrLf
    Next
    If Len(s) = 0 Then s = "No saves yet."
    txtHistory.Text = s
End Sub

Private Sub UpdateHeader()
    Dim e As Object, nShown As Long, nDesc As Long
    For Each e In mModel("entries")
        If e("found") Then
            If e("shown") = "1" Then nShown = nShown + 1
            If Len(Trim$(e("description"))) > 0 Then nDesc = nDesc + 1
        End If
    Next
    lblHeader.Caption = mModel("foundCount") & " macros found in " & mModel("moduleCount") & " module(s), read from " & _
        mModel("scanMethod") & ".   Described: " & nDesc & ".   In the menu: " & nShown & "."
End Sub

Private Sub UpdateRow()
    Dim i As Long
    If mCur Is Nothing Then Exit Sub
    i = lstEntries.ListIndex
    If i < 0 Then Exit Sub
    mLoading = True
    lstEntries.List(i, 0) = mCur("title")
    lstEntries.List(i, 1) = MenuCol(mCur)
    lstEntries.List(i, 2) = DescCol(mCur)
    mLoading = False
    lblMacro.Caption = "Macro: " & mCur("run") & "     Status: " & SK_EntryStatus(mCur)
End Sub

' ---------------------------------------------------------------------
' Pending changes
' ---------------------------------------------------------------------
Private Sub Mark(ByVal what As String)
    mPending.Add what
    UpdatePending
End Sub

Private Sub MarkEdit()
    If mCur Is Nothing Then Exit Sub
    If Not mEdited.Exists(mCur("run")) Then
        mEdited(mCur("run")) = True
        Mark "Edited " & mCur("title")
    End If
    UpdateRow
    UpdateHeader
End Sub

Private Sub UpdatePending()
    If mPending.Count = 0 Then
        lblPending.Caption = "No changes waiting."
    Else
        lblPending.Caption = mPending.Count & " change(s) waiting. Nothing is saved until you save."
    End If
    UpdateSaveButton
End Sub

' Save works only when there is something to save and Ready to save is ticked.
Private Sub UpdateSaveButton()
    If btnSave Is Nothing Or chkReady Is Nothing Then Exit Sub
    btnSave.Enabled = (mPending.Count > 0 And chkReady.Value = True)
End Sub

Private Sub chkReady_Click()
    UpdateSaveButton
End Sub

' The two list columns.
Private Function MenuCol(ByVal e As Object) As String
    If Not e("found") Then
        MenuCol = "Macro gone"
    ElseIf e("shown") = "1" Then
        MenuCol = "Shown"
    ElseIf e("shown") = "0" Then
        MenuCol = "Hidden"
    Else
        MenuCol = "Not set"
    End If
End Function

Private Function DescCol(ByVal e As Object) As String
    If Len(Trim$(e("description"))) > 0 Then DescCol = "Written" Else DescCol = "Missing"
End Function

Private Function Summary() As String
    Dim i As Long, s As String
    For i = 1 To mPending.Count
        If i > 12 Then
            s = s & "; and " & (mPending.Count - 12) & " more"
            Exit For
        End If
        s = s & IIf(Len(s) > 0, "; ", "") & mPending(i)
    Next
    Summary = s
End Function

' ---------------------------------------------------------------------
' Macros page events
' ---------------------------------------------------------------------
Private Sub txtFind_Change()
    If Not mLoading Then FillList
End Sub

Private Sub cboFilter_Change()
    If Not mLoading Then FillList
End Sub

Private Sub lstEntries_Click()
    If Not mLoading Then PickCurrent
End Sub

Private Sub txtTitle_Change()
    If mLoading Or mCur Is Nothing Then Exit Sub
    mCur("title") = txtTitle.Text
    MarkEdit
End Sub

Private Sub cboCategory_Change()
    If mLoading Or mCur Is Nothing Then Exit Sub
    mCur("category") = Trim$(cboCategory.Text)
    MarkEdit
End Sub

Private Sub cboWorksOn_Change()
    If mLoading Or mCur Is Nothing Then Exit Sub
    mCur("workson") = cboWorksOn.Text
    MarkEdit
End Sub

Private Sub txtDescription_Change()
    If mLoading Or mCur Is Nothing Then Exit Sub
    mCur("description") = txtDescription.Text
    MarkEdit
End Sub

Private Sub btnUseComment_Click()
    If mCur Is Nothing Then Exit Sub
    If Len(mCur("comment")) = 0 Then Exit Sub
    txtDescription.Text = mCur("comment")
End Sub

Private Sub txtKeywords_Change()
    If mLoading Or mCur Is Nothing Then Exit Sub
    mCur("keywords") = txtKeywords.Text
    MarkEdit
End Sub

Private Sub lstAffects_Change()
    Dim i As Long, s As String
    If mLoading Or mCur Is Nothing Then Exit Sub
    For i = 0 To lstAffects.ListCount - 1
        If lstAffects.Selected(i) Then s = s & IIf(Len(s) > 0, ", ", "") & lstAffects.List(i)
    Next
    mCur("affects") = s
    MarkEdit
End Sub

Private Sub cboChanges_Change()
    If mLoading Or mCur Is Nothing Then Exit Sub
    mCur("changes") = cboChanges.Text
    MarkEdit
End Sub

Private Sub chkShown_Click()
    If mLoading Or mCur Is Nothing Then Exit Sub
    mCur("shown") = IIf(chkShown.Value, "1", "0")
    Mark IIf(chkShown.Value, "Showed ", "Hid ") & mCur("title")
    UpdateRow
    UpdateHeader
End Sub

' ---------------------------------------------------------------------
' Categories page events
' ---------------------------------------------------------------------
Private Function CatAt(ByVal i As Long) As String
    If i >= 0 And i < mModel("cats").Count Then CatAt = mModel("cats")(i + 1)
End Function

Private Sub lstCats_Click()
    txtCatName.Text = CatAt(lstCats.ListIndex)
End Sub

Private Sub btnCatAdd_Click()
    Dim nm As String: nm = Trim$(txtCatName.Text)
    If Len(nm) = 0 Then Exit Sub
    Dim before As Long: before = mModel("cats").Count
    SK_AddUnique mModel("cats"), nm
    If mModel("cats").Count > before Then
        Mark "Added category " & nm
        FillCats
        FillCategoryCombo
        lstCats.ListIndex = mModel("cats").Count - 1
    End If
End Sub

Private Sub btnCatRename_Click()
    Dim i As Long, old As String, nm As String, e As Object, c As New Collection, v As Variant
    i = lstCats.ListIndex
    If i < 0 Then Exit Sub
    old = CatAt(i)
    nm = Trim$(txtCatName.Text)
    If Len(nm) = 0 Or nm = old Then Exit Sub
    For Each v In mModel("cats")
        If v = old Then c.Add nm Else c.Add v
    Next
    Set mModel("cats") = c
    For Each e In mModel("entries")
        If LCase$(e("category")) = LCase$(old) Then e("category") = nm
    Next
    Mark "Renamed category " & old & " to " & nm
    FillCats
    FillCategoryCombo
    lstCats.ListIndex = i
    FillList KeepRunNow()
End Sub

Private Sub MoveCat(ByVal delta As Long)
    Dim i As Long, arr() As String, k As Long, tmp As String, c As New Collection
    i = lstCats.ListIndex
    If i < 0 Then Exit Sub
    If i + delta < 0 Or i + delta >= mModel("cats").Count Then Exit Sub
    ReDim arr(0 To mModel("cats").Count - 1)
    For k = 0 To mModel("cats").Count - 1: arr(k) = mModel("cats")(k + 1): Next
    tmp = arr(i): arr(i) = arr(i + delta): arr(i + delta) = tmp
    For k = 0 To UBound(arr): c.Add arr(k): Next
    Set mModel("cats") = c
    Mark "Moved category " & tmp & IIf(delta < 0, " up", " down")
    FillCats
    FillCategoryCombo
    lstCats.ListIndex = i + delta
End Sub

Private Sub btnCatUp_Click()
    MoveCat -1
End Sub

Private Sub btnCatDown_Click()
    MoveCat 1
End Sub

Private Sub btnCatDelete_Click()
    Dim i As Long, old As String, e As Object, c As New Collection, v As Variant, n As Long
    i = lstCats.ListIndex
    If i < 0 Then Exit Sub
    old = CatAt(i)
    For Each e In mModel("entries")
        If LCase$(e("category")) = LCase$(old) Then n = n + 1
    Next
    If n > 0 Then
        If MsgBox("Delete the category " & old & "? Its " & n & " macro(s) move to Uncategorized.", _
                  vbYesNo + vbQuestion, SK_APP) <> vbYes Then Exit Sub
    End If
    For Each v In mModel("cats")
        If v <> old Then c.Add v
    Next
    Set mModel("cats") = c
    For Each e In mModel("entries")
        If LCase$(e("category")) = LCase$(old) Then e("category") = ""
    Next
    Mark "Deleted category " & old
    FillCats
    FillCategoryCombo
    txtCatName.Text = ""
    FillList KeepRunNow()
End Sub

Private Function KeepRunNow() As String
    If Not mCur Is Nothing Then KeepRunNow = mCur("run")
End Function

' ---------------------------------------------------------------------
' Look page events
' ---------------------------------------------------------------------
Private Sub txtMenuTitle_Change()
    If mLoading Then Exit Sub
    mModel("look")("menuTitle") = txtMenuTitle.Text
    If Not mEdited.Exists("|look|title") Then mEdited("|look|title") = True: Mark "Changed the menu title"
End Sub

Private Sub cboDefaultSort_Change()
    If mLoading Then Exit Sub
    mModel("look")("defaultSort") = cboDefaultSort.Text
    Mark "Menu opens sorted by " & cboDefaultSort.Text
End Sub

Private Sub chkShowUnd_Click()
    If mLoading Then Exit Sub
    mModel("look")("showUndescribed") = IIf(chkShowUnd.Value, "1", "0")
    Mark IIf(chkShowUnd.Value, "Menu lists undescribed macros", "Menu hides undescribed macros")
End Sub

Private Sub chkKeepOpen_Click()
    If mLoading Then Exit Sub
    mModel("look")("keepOpen") = IIf(chkKeepOpen.Value, "1", "0")
    Mark "Keep open default: " & IIf(chkKeepOpen.Value, "on", "off")
End Sub

Private Sub chkConfirm_Click()
    If mLoading Then Exit Sub
    mModel("look")("confirmChanges") = IIf(chkConfirm.Value, "1", "0")
    Mark "Ask before changing a document: " & IIf(chkConfirm.Value, "on", "off")
End Sub

Private Sub btnShortcut_Click()
    Dim msg As String: msg = SK_SetShortcut()
    If Left$(msg, 3) = "Alt" Then Mark "Set Alt+Shift+M"
    MsgBox msg & vbCr & "Save to keep it.", vbInformation, SK_APP
End Sub

Private Sub btnButton_Click()
    Dim msg As String: msg = SK_AddButton()
    If Left$(msg, 3) = "The" And InStr(msg, "now has") > 0 Then Mark "Added the Add-ins tab button"
    MsgBox msg & vbCr & "Save to keep it.", vbInformation, SK_APP
End Sub

' ---------------------------------------------------------------------
' Bulk tools events
' ---------------------------------------------------------------------

Private Sub btnFillComments_Click()
    Dim e As Object, n As Long
    For Each e In mModel("entries")
        If e("found") And Len(Trim$(e("description"))) = 0 And Len(e("comment")) > 0 Then
            e("description") = e("comment")
            n = n + 1
        End If
    Next
    If n > 0 Then Mark "Described " & n & " macro(s) from code comments"
    FillList KeepRunNow()
    UpdateHeader
    MsgBox IIf(n > 0, n & " macro(s) now use their code comment as the description. Review them on the Macros tab.", _
                    "No empty descriptions had a code comment to use."), vbInformation, SK_APP
End Sub

Private Sub btnShowDescribed_Click()
    Dim e As Object, n As Long
    For Each e In mModel("entries")
        If e("found") And Len(Trim$(e("description"))) > 0 And e("shown") <> "1" And e("shown") <> "0" Then
            e("shown") = "1"
            n = n + 1
        End If
    Next
    If n > 0 Then Mark "Showed " & n & " described macro(s)"
    FillList KeepRunNow()
    UpdateHeader
    MsgBox n & " described macro(s) added to the menu. Macros you hid on purpose stay hidden.", vbInformation, SK_APP
End Sub

Private Sub btnDropMissing_Click()
    Dim e As Object, c As New Collection, n As Long
    For Each e In mModel("entries")
        If e("found") Then c.Add e Else n = n + 1
    Next
    If n = 0 Then
        MsgBox "Every entry has its macro.", vbInformation, SK_APP
        Exit Sub
    End If
    Set mModel("entries") = c
    Mark "Removed " & n & " entry(ies) whose macro is gone"
    FillList
    UpdateHeader
End Sub

Private Sub btnExportCsv_Click()
    Dim p As String
    On Error GoTo Fail
    p = SK_ExportCsv(mModel)
    MsgBox "Saved:" & vbCr & p & vbCr & vbCr & "Edit it in Excel, save as CSV UTF-8, then use Import.", vbInformation, SK_APP
    Exit Sub
Fail:
    MsgBox "Export failed: " & Err.Description, vbExclamation, SK_APP
End Sub

Private Sub btnImportCsv_Click()
    Dim p As String, msg As String
    On Error GoTo Fail
    With Application.FileDialog(3)
        .Title = "Pick the catalog CSV to import"
        .Filters.Clear
        .Filters.Add "CSV files", "*.csv"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        p = .SelectedItems(1)
    End With
    msg = SK_ImportCsv(mModel, p)
    If Left$(msg, 8) = "Updated " Then Mark "Imported " & Mid$(p, InStrRev(p, "\") + 1)
    FillCats
    FillCategoryCombo
    AddExtraAffects
    FillList KeepRunNow()
    UpdateHeader
    MsgBox msg, vbInformation, SK_APP
    Exit Sub
Fail:
    MsgBox "Import failed: " & Err.Description, vbExclamation, SK_APP
End Sub

Private Sub btnReadBas_Click()
    Dim paths As New Collection, i As Long, found As Collection, n As Long
    On Error GoTo Fail
    With Application.FileDialog(3)
        .Title = "Pick the exported module file(s) to read"
        .Filters.Clear
        .Filters.Add "VBA modules", "*.bas"
        .AllowMultiSelect = True
        If .Show <> -1 Then Exit Sub
        For i = 1 To .SelectedItems.Count
            paths.Add .SelectedItems(i)
        Next
    End With
    Set found = SK_ReadBasFiles(paths)
    n = SK_MergeFound(mModel, found)
    If found.Count > 0 Then Mark "Read " & found.Count & " macro(s) from " & paths.Count & " .bas file(s)"
    mModel("foundCount") = mModel("entries").Count
    FillList KeepRunNow()
    UpdateHeader
    MsgBox found.Count & " macro(s) found in " & paths.Count & " file(s); " & n & " new to the catalog." & vbCr & _
        "They stay listed after you save. If a macro is not really in this template, the menu says so when it is run.", _
        vbInformation, SK_APP
    Exit Sub
Fail:
    MsgBox "Could not read the files: " & Err.Description, vbExclamation, SK_APP
End Sub

' ---------------------------------------------------------------------
' Save and close
' ---------------------------------------------------------------------
Private Sub btnSave_Click()
    Dim msg As String
    If mPending.Count = 0 Then
        MsgBox "Nothing to save.", vbInformation, SK_APP
        Exit Sub
    End If
    If Not chkReady.Value Then
        MsgBox "Tick Ready to save first.", vbInformation, SK_APP
        Exit Sub
    End If
    KeepCategory
    Dim e As Object
    For Each e In mModel("entries")
        If Len(e("category")) > 0 Then SK_AddUnique mModel("cats"), e("category")
    Next
    msg = SK_SaveModel(mModel, Summary())
    FillCats
    Set mPending = New Collection
    mEdited.RemoveAll
    chkReady.Value = False
    UpdatePending
    FillHistory
    MsgBox msg, vbInformation, SK_APP
End Sub

Private Function OkToClose() As Boolean
    OkToClose = True
    If mPending Is Nothing Then Exit Function
    If mPending.Count > 0 Then
        OkToClose = (MsgBox(mPending.Count & " change(s) are not saved. Close anyway?", _
                            vbYesNo + vbQuestion, SK_APP) = vbYes)
    End If
End Function

Private Sub btnClose_Click()
    If OkToClose() Then Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Only the window's X asks here. Close asks in btnClose_Click.
    If CloseMode = vbFormControlMenu Then
        If Not OkToClose() Then Cancel = True
    End If
End Sub

' ---------------------------------------------------------------------
' Used by the self-test only
' ---------------------------------------------------------------------
Public Function TestState() As String
    TestState = lstEntries.ListCount & " rows; " & lblHeader.Caption
End Function

Public Function TestLabels() As String
    Dim s As String, i As Long, c As Object
    For i = 0 To mp.Pages.Count - 1
        s = s & mp.Pages(i).Caption & "|"
    Next
    s = s & " A=" & mp.Pages(0).Controls("lblA").Caption & " | Ch=" & mp.Pages(0).Controls("lblCh").Caption
    On Error Resume Next
    Set c = Nothing
    Set c = mp.Pages(0).Controls("txtAlso")
    s = s & " | also box=" & IIf(c Is Nothing, "none", "present")
    TestLabels = s
End Function

Public Function TestRow(ByVal i As Long) As String
    If i >= lstEntries.ListCount Then Exit Function
    TestRow = lstEntries.List(i, 0) & " | " & lstEntries.List(i, 1) & " | " & lstEntries.List(i, 2)
End Function

Public Function TestSaveGate() As String
    Dim a As Boolean, b As Boolean
    chkReady.Value = False
    a = btnSave.Enabled
    chkReady.Value = True
    b = btnSave.Enabled
    chkReady.Value = False
    TestSaveGate = "pending=" & mPending.Count & " unticked=" & a & " ticked=" & b
End Function

Public Function TestEditFirst(ByVal desc As String) As String
    If lstEntries.ListCount = 0 Then Exit Function
    lstEntries.ListIndex = 0
    txtDescription.Text = desc
    chkShown.Value = True
    TestEditFirst = mCur("run") & " | " & SK_EntryStatus(mCur) & " | pending " & mPending.Count
End Function
