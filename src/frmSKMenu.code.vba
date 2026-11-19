' =====================================================================
' Stationkeeper Macros: the menu
' Search as you type, sort, filter by what a macro affects or by its
' tie to Stationkeeper, read, run. The link strip shows the effort folder.
' The controls are built in code, so this form needs no designer work.
' =====================================================================
Option Explicit

Private WithEvents txtSearch As MSForms.TextBox
Private WithEvents lstMacros As MSForms.ListBox
Private WithEvents cboSort As MSForms.ComboBox
Private WithEvents btnRun As MSForms.CommandButton
Private WithEvents btnPin As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton
Private WithEvents btnReports As MSForms.CommandButton
Private WithEvents btnStd As MSForms.CommandButton
Private WithEvents btnEff As MSForms.CommandButton
Private WithEvents btnEffClear As MSForms.CommandButton
Private lblStdG As MSForms.Label
Private lblStdE As MSForms.Label
Private mHide As String          ' the admin tools left out for this person: |Module.Macro|...|, or ""
Private mKind As String          ' tests: list only the macros made for Stationkeeper, or only the others
Private WithEvents chkKeep As MSForms.CheckBox
Private WithEvents cboShow As MSForms.ComboBox
Private WithEvents cboCat As MSForms.ComboBox
Private WithEvents btnLink As MSForms.CommandButton
Private WithEvents btnRefresh As MSForms.CommandButton
Private WithEvents btnUnlink As MSForms.CommandButton
Private WithEvents btnSuggest As MSForms.CommandButton
Private lblSuggest As MSForms.Label
Private WithEvents btnAskRun As MSForms.CommandButton
Private WithEvents btnAskLink As MSForms.CommandButton
Private WithEvents btnAskCancel As MSForms.CommandButton
Private chkAskNever As MSForms.CheckBox
Private fraAsk As MSForms.Frame
Private lblAskHead As MSForms.Label
Private lblAskText As MSForms.Label
Private lblLink As MSForms.Label
Private lblLegend As MSForms.Label
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
Private mInfo As Object          ' the Stationkeeper link, read fresh (SK_LinkInfo)
Private mAskRun As String        ' the macro waiting on the question panel
Private mAskKind As String       ' "ask" or "needs"
Private mDocPath As String       ' the folder of the document the menu was opened from
Private mSuggest As String       ' an effort folder to offer (the document's own), or ""
Private mReady As Boolean        ' the controls exist and the form is dressed: Resize may lay them out
Private mLaying As Boolean
Private mNoSave As Boolean       ' tests: do not keep the size

' The menu opens compact and the user drags its border to resize it; the size is kept per user.
Private Const DEF_W As Single = 500
Private Const DEF_H As Single = 600
Private Const MIN_W As Single = 460
Private Const MIN_H As Single = 540
Private Const MAX_W As Single = 1400
Private Const MAX_H As Single = 1000

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

Private Function AddIn(ByVal parent As Object, ByVal progId As String, ByVal nm As String, ByVal l As Single, _
                       ByVal t As Single, ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = parent.Controls.Add(progId, nm, True)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    On Error Resume Next
    c.Font.Name = "Segoe UI"
    c.Font.Size = 9
    Set AddIn = c
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
    Me.Width = DEF_W
    Me.Height = DEF_H - 50     ' the Stationkeeper band adds its height when the form is dressed
    Me.BackColor = &HFFFFFF

    AddLabel "lblSearch", "Search", 10, 8, 60, 14
    Set lblCount = AddLabel("lblCount", "", 50, 8, 120, 14)
    lblCount.TextAlign = fmTextAlignRight
    lblCount.ForeColor = &H808080
    Set txtSearch = AddCtl("Forms.TextBox.1", "txtSearch", 10, 22, 160, 20)
    txtSearch.Font.Size = 10

    AddLabel "lblShow", "Application", 176, 8, 100, 14
    Set cboShow = AddCtl("Forms.ComboBox.1", "cboShow", 176, 22, 100, 20)
    cboShow.Style = fmStyleDropDownList
    cboShow.AddItem "All"
    cboShow.AddItem "Word"
    cboShow.AddItem "Excel"
    cboShow.AddItem "PowerPoint"
    cboShow.ControlTipText = "The application a macro works in"

    AddLabel "lblCat", "Category", 282, 8, 100, 14
    Set cboCat = AddCtl("Forms.ComboBox.1", "cboCat", 282, 22, 110, 20)
    cboCat.Style = fmStyleDropDownList

    AddLabel "lblSort", "Sort by", 398, 8, 70, 14
    Set cboSort = AddCtl("Forms.ComboBox.1", "cboSort", 398, 22, 72, 20)
    cboSort.Style = fmStyleDropDownList
    cboSort.AddItem "Category"
    cboSort.AddItem "Name"
    cboSort.AddItem "Most used"
    cboSort.AddItem "Recent"

    Set lstMacros = AddCtl("Forms.ListBox.1", "lstMacros", 10, 50, 300, 298)
    lstMacros.ColumnCount = 3
    lstMacros.ColumnWidths = "168 pt;34 pt;84 pt"
    lstMacros.IntegralHeight = False
    Set lblLegend = AddLabel("lblLegend", ChrW$(&H25C7) & " SK  built for Stationkeeper, works alone     " & _
                             ChrW$(&H25C6) & " SK  needs Stationkeeper data", 10, 352, 300, 14)
    lblLegend.ForeColor = &H808080
    lblLegend.Font.Size = 8

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
    lblTitle.WordWrap = True

    Set btnRun = AddCtl("Forms.CommandButton.1", "btnRun", 322, 340, 90, 26)
    btnRun.Caption = "Run"
    btnRun.Default = True
    Set btnPin = AddCtl("Forms.CommandButton.1", "btnPin", 418, 340, 90, 26)
    btnPin.Caption = "Pin to top"
    Set chkKeep = AddCtl("Forms.CheckBox.1", "chkKeep", 516, 344, 134, 18)
    chkKeep.Caption = "Keep this menu open"
    chkKeep.BackStyle = fmBackStyleTransparent

    ' The link strip: where the Stationkeeper data comes from.
    Set lblLink = AddLabel("lblLink", "", 10, 374, 428, 42)
    lblLink.WordWrap = True
    Set btnLink = AddCtl("Forms.CommandButton.1", "btnLink", 444, 376, 82, 22)
    Set btnRefresh = AddCtl("Forms.CommandButton.1", "btnRefresh", 530, 376, 58, 22)
    btnRefresh.Caption = "Refresh"
    Set btnUnlink = AddCtl("Forms.CommandButton.1", "btnUnlink", 592, 376, 58, 22)
    btnUnlink.Caption = "Unlink"

    Set lblSuggest = AddLabel("lblSuggest", "", 10, 422, 428, 16)
    lblSuggest.ForeColor = &H8B4513
    Set btnSuggest = AddCtl("Forms.CommandButton.1", "btnSuggest", 444, 418, 206, 22)
    btnSuggest.Caption = "Use this folder"

    Set lblStatus = AddLabel("lblStatus", "Type to search. Up and Down move, Enter runs, Esc closes.", 10, 448, 332, 14)
    lblStatus.ForeColor = &H808080
    Set btnClose = AddCtl("Forms.CommandButton.1", "btnClose", 570, 442, 80, 24)
    Set btnReports = AddCtl("Forms.CommandButton.1", "btnReports", 466, 442, 98, 24)
    btnReports.Caption = "Reports folder..."
    btnReports.ControlTipText = IIf(Len(SK_ReportsDir()) > 0, "Reports are saved in " & SK_ReportsDir(), "Reports open and are not saved. Choose a folder to keep them.")
    btnClose.Caption = "Close"
    btnClose.Cancel = True
    ' The Global Standards workbook the tools read. It is not the effort's folder: that is the link strip above.
    Set btnStd = AddCtl("Forms.CommandButton.1", "btnStd", 346, 442, 116, 22)
    btnStd.Caption = "Choose..."
    btnStd.ControlTipText = StdTip()
    Set lblStdG = AddLabel("lblStdG", "", 10, 400, 300, 14)
    Set lblStdE = AddLabel("lblStdE", "", 10, 420, 300, 14)
    Set btnEff = AddCtl("Forms.CommandButton.1", "btnEff", 346, 420, 84, 22)
    btnEff.Caption = "Choose..."
    btnEff.ControlTipText = "The effort's or the program's own standards workbook. Where it sets a value it wins; a blank keeps the Global Standards' value."
    Set btnEffClear = AddCtl("Forms.CommandButton.1", "btnEffClear", 432, 420, 56, 22)
    btnEffClear.Caption = "Clear"
    btnEffClear.ControlTipText = "Stop laying an effort's standards over the Global Standards"
    StdLines

    ' The question panel (over the description) for macros tied to Stationkeeper.
    Set fraAsk = AddCtl("Forms.Frame.1", "fraAsk", 322, 88, 328, 244)
    fraAsk.Caption = ""
    fraAsk.BackColor = &HF7F3EE
    fraAsk.SpecialEffect = fmSpecialEffectFlat
    fraAsk.BorderStyle = fmBorderStyleSingle
    fraAsk.BorderColor = &HC0B4A0
    Set lblAskHead = AddIn(fraAsk, "Forms.Label.1", "lblAskHead", 10, 8, 306, 18)
    lblAskHead.Font.Bold = True
    lblAskHead.Font.Size = 10
    lblAskHead.BackStyle = fmBackStyleTransparent
    Set lblAskText = AddIn(fraAsk, "Forms.Label.1", "lblAskText", 10, 30, 306, 146)
    lblAskText.WordWrap = True
    lblAskText.BackStyle = fmBackStyleTransparent
    lblAskText.Font.Size = 9.5
    Set chkAskNever = AddIn(fraAsk, "Forms.CheckBox.1", "chkAskNever", 10, 180, 300, 16)
    chkAskNever.Caption = "Don't ask again"
    chkAskNever.BackStyle = fmBackStyleTransparent
    Set btnAskRun = AddIn(fraAsk, "Forms.CommandButton.1", "btnAskRun", 10, 206, 80, 24)
    btnAskRun.Caption = "Run it"
    Set btnAskLink = AddIn(fraAsk, "Forms.CommandButton.1", "btnAskLink", 96, 206, 140, 24)
    Set btnAskCancel = AddIn(fraAsk, "Forms.CommandButton.1", "btnAskCancel", 242, 206, 74, 24)
    btnAskCancel.Caption = "Cancel"
    fraAsk.Visible = False

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

    v = SK_GetPref("sort", mModel("look")("defaultSort"))
    If v = "A to Z" Then v = "Name"                        ' its name before 1.16.0
    cboSort.ListIndex = 0
    For i = 0 To cboSort.ListCount - 1
        If LCase$(cboSort.List(i)) = LCase$(v) Then cboSort.ListIndex = i
    Next
    chkKeep.Value = (SK_GetPref("keepOpen", mModel("look")("keepOpen")) = "1")
    v = SK_GetPref("app", "All")
    cboShow.ListIndex = 0
    For i = 0 To cboShow.ListCount - 1
        If LCase$(cboShow.List(i)) = LCase$(v) Then cboShow.ListIndex = i
    Next
    On Error Resume Next
    If Documents.Count > 0 Then mDocPath = ActiveDocument.Path
    On Error GoTo 0
    Set mInfo = SK_LinkInfo()
    ShowLink
    FillCats

    mBusy = False
    RefreshList
    frmSKBrand.Dress Me, mModel("look")("menuTitle"), "Team macros and Stationkeeper Tools   |   " & ThisDocument.Name & "   |   " & SK_VERSION, "btnRun|btnAskRun"
    ' The size the user left the menu at.
    Me.Width = Clamp(Val(SK_GetPref("menuW", CStr(DEF_W))), MIN_W, MAX_W)
    Me.Height = Clamp(Val(SK_GetPref("menuH", CStr(DEF_H))), MIN_H, MAX_H)
    mReady = True
    Layout
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    frmSKBrand.SetIcon Me
    frmSKBrand.Sizable Me
    Layout
    txtSearch.SetFocus
End Sub

Private Function Clamp(ByVal v As Single, ByVal lo As Single, ByVal hi As Single) As Single
    If v < lo Then v = lo
    If v > hi Then v = hi
    Clamp = v
End Function

Private Sub UserForm_Resize()
    If Not mReady Or mLaying Then Exit Sub
    Layout
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    If mNoSave Or Not mReady Then Exit Sub
    If Me.Width <> Val(SK_GetPref("menuW", CStr(DEF_W))) Then SK_SetPref "menuW", CStr(CLng(Me.Width))
    If Me.Height <> Val(SK_GetPref("menuH", CStr(DEF_H))) Then SK_SetPref "menuH", CStr(CLng(Me.Height))
End Sub

' By name: Left, Top, Width and Height belong to the control as the form holds it.
Private Sub Place(ByVal nm As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single)
    Dim c As MSForms.Control
    On Error Resume Next
    Set c = Me.Controls(nm)
    If c Is Nothing Then Exit Sub
    If w < 4 Then w = 4
    If h < 4 Then h = 4
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
End Sub

' Places every control for the form's size now. Two rows of filters, the list beside the
' description, then the link strip, the status line and the bottom buttons.
Private Sub Layout()
    Dim w As Single, h As Single, o As Single, u As Single, lw As Single, rx As Single, rw As Single
    Dim y As Single, pb As Single, half As Single, ph As Single, fh As Single, fw As Single
    If mLaying Then Exit Sub
    mLaying = True
    On Error Resume Next
    If Me.Width < MIN_W Then Me.Width = MIN_W
    If Me.Height < MIN_H Then Me.Height = MIN_H
    w = Me.InsideWidth
    h = Me.InsideHeight
    o = Me.Controls("lblSKLine").Top + Me.Controls("lblSKLine").Height
    frmSKBrand.Fit Me
    u = w - 20

    ' One row of filters: Search, Application, Category, Sort by.
    Place "lblSearch", 10, o + 8, 60, 14
    Place "txtSearch", 10, o + 22, u - 306, 20
    Place "lblCount", 10 + (u - 306) - 90, o + 8, 90, 14
    Place "lblShow", 10 + u - 300, o + 8, 92, 14
    Place "cboShow", 10 + u - 300, o + 22, 92, 20
    Place "lblCat", 10 + u - 202, o + 8, 116, 14
    Place "cboCat", 10 + u - 202, o + 22, 116, 20
    Place "lblSort", 10 + u - 80, o + 8, 80, 14
    Place "cboSort", 10 + u - 80, o + 22, 80, 20

    ' From the bottom up: buttons, status, the folder offer, the link strip, the legend.
    y = h - 8 - 24
    Place "btnReports", 10, y, 96, 24
    Place "chkKeep", 114, y + 3, 130, 18
    Place "btnClose", w - 10 - 70, y, 70, 24
    y = y - 18
    Place "lblStatus", 10, y, u, 14
    y = y - 26
    Place "lblSuggest", 10, y + 4, u - 146, 16
    Place "btnSuggest", w - 10 - 140, y, 140, 22
    y = y - 46
    Place "lblLink", 10, y, u - 210, 42
    Place "btnLink", w - 10 - 204, y + 2, 84, 22
    Place "btnRefresh", w - 10 - 116, y + 2, 56, 22
    Place "btnUnlink", w - 10 - 56, y + 2, 56, 22
    ' Above the effort data link: the two standards workbooks, each with its own Choose.
    y = y - 26
    Place "lblStdE", 10, y + 4, u - 156, 14
    Place "btnEff", w - 10 - 146, y, 84, 22
    Place "btnEffClear", w - 10 - 56, y, 56, 22
    y = y - 24
    Place "lblStdG", 10, y + 4, u - 156, 14
    Place "btnStd", w - 10 - 146, y, 84, 22
    y = y - 18
    Place "lblLegend", 10, y, u, 14
    pb = y - 4

    ' The list beside the description.
    y = o + 50
    ph = pb - y
    lw = Int(u * 0.44)
    If lw < 190 Then lw = 190
    rx = 10 + lw + 10
    rw = w - 10 - rx
    Place "lstMacros", 10, y, lw, ph
    ' A list box made in Initialize can keep its first width until it is set a second time.
    If Abs(lstMacros.Width - lw) > 1 Then lstMacros.Width = lw + 1: lstMacros.Width = lw
    If lw >= 280 Then
        lstMacros.ColumnWidths = CLng(lw - 124) & " pt;34 pt;84 pt"
    Else
        lstMacros.ColumnWidths = CLng(lw - 40) & " pt;34 pt;0 pt"
    End If
    Place "lblTitle", rx, y, rw, 34
    Place "lblRun", rx, y + 36, rw, 14
    Place "txtDesc", rx, y + 54, rw, ph - 54 - 6 - 84 - 6 - 26
    Place "lblInfo", rx, y + ph - 26 - 6 - 84, rw, 84
    Place "btnRun", rx, y + ph - 26, 84, 26
    Place "btnPin", rx + 90, y + ph - 26, 84, 26

    ' The question panel covers the description.
    fw = rw
    fh = ph - 52 - 32
    Place "fraAsk", rx, y + 52, fw, fh
    Place "lblAskHead", 10, 8, fw - 22, 18
    Place "lblAskText", 10, 30, fw - 22, fh - 98
    Place "chkAskNever", 10, fh - 64, fw - 22, 16
    Place "btnAskRun", 10, fh - 40, 56, 24
    Place "btnAskLink", 72, fh - 40, fw - 22 - 62 - 64, 24
    Place "btnAskCancel", fw - 12 - 58, fh - 40, 58, 24
    mLaying = False
End Sub

' Tests: the size, whether the border can be dragged, whether every control is inside the form,
' and the same after the form is made wider and taller. The size is not kept.
Public Function TestSize() As String
    Dim s As String, lw As Single
    mNoSave = True
    Layout      ' as Activate does when the menu is shown
    s = "w=" & Me.Width & " h=" & Me.Height & " list=" & CLng(lstMacros.Width) & " inside=" & Inside() & " cols=" & lstMacros.ColumnWidths
    frmSKBrand.Sizable Me
    s = s & " sizable=" & frmSKBrand.IsSizable(Me)
    lw = lstMacros.Width
    Me.Width = 800
    Me.Height = 620
    s = s & " || wide: w=" & Me.Width & " h=" & Me.Height & " grew=" & (lstMacros.Width > lw + 100) & " list=" & CLng(lstMacros.Width) & " inside=" & Inside() & _
        " band=" & (Abs(Me.Controls("lblSKBand").Width - Me.InsideWidth) < 1) & " cols=" & lstMacros.ColumnWidths
    Me.Width = 200
    Me.Height = 200
    s = s & " || small: w=" & Me.Width & " h=" & Me.Height & " inside=" & Inside()
    Me.Width = DEF_W
    Me.Height = DEF_H
    TestSize = s
End Function

' Every visible control lies inside the form, and no two on the form's bottom row overlap.
Private Function Inside() As String
    Dim c As Object, bad As String
    For Each c In Me.Controls
        If c.Parent.Name = Me.Name Then
            If c.Left < -0.5 Or c.Top < -0.5 Or c.Left + c.Width > Me.InsideWidth + 0.5 Or c.Top + c.Height > Me.InsideHeight + 0.5 Then bad = bad & " " & c.Name
        End If
    Next
    If chkKeep.Left + chkKeep.Width > btnClose.Left Then bad = bad & " keep-over-close"
    If lstMacros.Left + lstMacros.Width > txtDesc.Left Then bad = bad & " list-over-desc"
    If Abs(txtDesc.Left + txtDesc.Width - (Me.InsideWidth - 10)) > 1 Then bad = bad & " desc-not-at-edge"
    If lblLink.Left + lblLink.Width > btnLink.Left Then bad = bad & " link-over-button"
    If txtDesc.Height < 40 Then bad = bad & " desc-short"
    Inside = IIf(Len(bad) = 0, "all", "out:" & bad)
End Function

' For the tests: what the Stationkeeper look put on this form.
Public Function TestBrand() As String
    TestBrand = frmSKBrand.Probe(Me) & " height=" & Me.Height & " searchTop=" & txtSearch.Top
End Function

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
    Dim flt As String

    If mBusy Then Exit Sub
    q = Trim$(txtSearch.Text)
    flt = cboShow.Text
    mode = cboSort.Text
    n = mItems.Count
    mHide = AdminHide()
    ReDim keys(1 To n + 1): ReDim idx(1 To n + 1): ReDim why(1 To n + 1)

    For i = 1 To n
        Set e = mItems(i)
        If (Len(aff) = 0 Or HasAffect(e, aff)) And ShowPasses(e, flt) And CatPasses(e) And Not AdminHidden(e) Then
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
                        Case "Name"
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
        lstMacros.List(i - 1, 1) = SK_SkBadge(e("stationkeeper"))
        lstMacros.List(i - 1, 2) = col2
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
    If Not fraAsk Is Nothing Then fraAsk.Visible = False
    If lstMacros.ListIndex < 0 Then
        lblTitle.Caption = ""
        lblRun.Caption = ""
        If mItems.Count = 0 Then
            txtDesc.Text = "No macros are in the menu yet." & vbCrLf & vbCrLf & _
                "Run StationkeeperMacrosAdmin to describe this template's macros and choose which ones show here."
        Else
            txtDesc.Text = "No macro matches. Try fewer letters, another word for what it does, or set Show to All macros and Affects to Anything."
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
           "Works on:  " & Dash(e("workson"), "not set") & NeedsNote(e("workson")) & vbCrLf & _
           "Does it change the document?  " & Dash(e("changes"), "not set") & vbCrLf & _
           "Stationkeeper:  " & SK_SkLabel(e("stationkeeper"))
    lblInfo.Caption = info
    btnPin.Caption = IIf(SK_IsPinned(e("run")), "Unpin", "Pin to top")
    btnRun.Enabled = True
    btnPin.Enabled = True
End Sub

Private Function NeedsNote(ByVal w As String) As String
    Select Case LCase$(w)
        Case "selection", "table at the cursor", "list at the cursor"
            NeedsNote = " (checked before it runs)"
    End Select
End Function

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
    Dim e As Object, why As String, gate As String
    If lstMacros.ListIndex < 0 Then Exit Sub
    Set e = mItems(mView(lstMacros.ListIndex))
    ' Check what the macro needs before it can fail.
    If Not SK_NeedsMet(e("workson"), why) Then
        lblStatus.Caption = why & " " & e("title") & " works on: " & LCase$(e("workson")) & "."
        Beep
        Exit Sub
    End If
    ' Its tie to Stationkeeper: run, ask gently, or say what data it needs.
    Set mInfo = SK_LinkInfo()
    ShowLink
    gate = SK_Gate(e("stationkeeper"), mInfo, why)
    If Len(gate) > 0 Then
        ShowAsk e, gate, why
        Exit Sub
    End If
    DoRun e
End Sub

Private Sub DoRun(ByVal e As Object)
    Dim keep As Boolean, ok As Boolean, run As String
    If mModel("look")("confirmChanges") = "1" And LCase$(Left$(e("changes"), 3)) = "yes" Then
        If MsgBox(e("title") & " changes the document. Run it?", vbYesNo + vbQuestion, SK_APP) <> vbYes Then Exit Sub
    End If
    keep = chkKeep.Value
    run = e("run")
    If Not keep Then Me.Hide
    ok = SK_RunMacro(run, e("title"))
    If keep Then
        lblStatus.Caption = IIf(ok, "Ran " & e("title") & " at " & Format$(Now, "h:nn am/pm") & _
                                    ". One Ctrl+Z undoes all of it.", e("title") & " could not run.")
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


Private Sub cboShow_Change()
    If mBusy Then Exit Sub
    SK_SetPref "app", cboShow.Text
    RefreshList
End Sub

' ---------------------------------------------------------------------
' Stationkeeper: the filter, the link strip, the question panel
' ---------------------------------------------------------------------
Private Function IsSkTool(ByVal e As Object) As Boolean
    IsSkTool = (SK_ModulePart(e("run")) Like "SK[WPX]_*") Or (LCase$(SK_ModulePart(e("run"))) = "sk_install") Or (LCase$(SK_ModulePart(e("run"))) = "sk_macros")
End Function

' The Category list: the admin's categories in order, then any others macros use, then (No category).
Private Sub FillCats()
    Dim v As Variant, e As Object, seen As Object, blank As Boolean, keep As String, i As Long
    keep = SK_GetPref("category", "All categories")
    cboCat.Clear
    cboCat.AddItem "All categories"
    Set seen = CreateObject("Scripting.Dictionary")
    For Each v In mModel("cats")
        If Not seen.Exists(LCase$(v)) Then cboCat.AddItem v: seen(LCase$(v)) = True
    Next
    For Each e In mItems
        If Len(Trim$(e("category"))) = 0 Then
            blank = True
        ElseIf Not seen.Exists(LCase$(e("category"))) Then
            cboCat.AddItem e("category"): seen(LCase$(e("category"))) = True
        End If
    Next
    If blank Then cboCat.AddItem "(No category)"
    cboCat.ListIndex = 0
    For i = 0 To cboCat.ListCount - 1
        If LCase$(cboCat.List(i)) = LCase$(keep) Then cboCat.ListIndex = i
    Next
End Sub

Private Function CatPasses(ByVal e As Object) As Boolean
    If cboCat.ListIndex <= 0 Then CatPasses = True: Exit Function
    If cboCat.Text = "(No category)" Then
        CatPasses = (Len(Trim$(e("category"))) = 0)
    Else
        CatPasses = (LCase$(e("category")) = LCase$(cboCat.Text))
    End If
End Function

Private Sub cboCat_Change()
    If mBusy Then Exit Sub
    SK_SetPref "category", cboCat.Text
    RefreshList
End Sub

' The application a macro works in: the tools say it by their module's name; a team macro says it in its
' category or in what it affects; otherwise it is a Word macro (this is a Word template).
Private Function AppOf(ByVal e As Object) As String
    Dim m As String, t As String
    On Error Resume Next
    m = UCase$(SK_ModulePart(e("run")))
    t = LCase$(e("category") & " " & e("affects"))
    If Left$(m, 4) = "SKP_" Or InStr(t, "powerpoint") > 0 Then
        AppOf = "PowerPoint"
    ElseIf Left$(m, 4) = "SKX_" Or InStr(t, "excel") > 0 Then
        AppOf = "Excel"
    Else
        AppOf = "Word"
    End If
End Function

' The tools that change the Global Standards are for Stationkeeper Admins (the list is a sheet of the
' workbook). The tools say which of them this person does not see; with no tools installed, none are hidden.
Private Function AdminHide() As String
    On Error Resume Next
    AdminHide = Application.Run("SKW_Lib.SKW_AdminHiddenList")
End Function

Private Function AdminHidden(ByVal e As Object) As Boolean
    If Len(mHide) = 0 Then Exit Function
    On Error Resume Next
    AdminHidden = (InStr(1, mHide, "|" & e("run") & "|", vbTextCompare) > 0)
End Function

Private Function AdminWords(ByVal fn As String) As String
    On Error Resume Next
    AdminWords = Application.Run("SKW_Lib." & fn)
End Function

Private Function ShowPasses(ByVal e As Object, ByVal flt As String) As Boolean
    Dim k As String
    ShowPasses = True
    If flt = "Word" Or flt = "Excel" Or flt = "PowerPoint" Then ShowPasses = (AppOf(e) = flt)
    If Not ShowPasses Or Len(mKind) = 0 Then Exit Function
    k = SK_SkKind(e("stationkeeper"))
    If mKind = "Other" Then ShowPasses = (k = "none" And Not IsSkTool(e))
    If mKind = "Stationkeeper" Then ShowPasses = (k <> "none" Or IsSkTool(e))
End Function

Private Sub ShowLink()
    Dim st As String
    st = mInfo("state")
    lblLink.Caption = SK_LinkLine(mInfo) & vbCrLf & mInfo("note")
    Select Case st
        Case "Current": lblLink.ForeColor = &H2E7D32
        Case "Not linked": lblLink.ForeColor = &H606060
        Case "Aging", "Stale", "Incomplete": lblLink.ForeColor = &H1A6FB8
        Case Else: lblLink.ForeColor = &H2828C0
    End Select
    btnLink.Caption = IIf(st = "Not linked", "Link a folder...", "Change...")
    btnUnlink.Enabled = (st <> "Not linked")
    If Len(mInfo("warn")) > 0 And (st = "Current" Or st = "Aging" Or st = "Stale") Then lblLink.ForeColor = &H1A6FB8
    ' Offer the effort folder this document sits in, when the menu reads another one (or none).
    Dim msg As String
    mSuggest = SK_SuggestLink(mDocPath, mInfo, msg)
    lblSuggest.Caption = msg
    lblSuggest.Visible = (Len(mSuggest) > 0)
    btnSuggest.Visible = (Len(mSuggest) > 0)
End Sub

Private Sub btnSuggest_Click()
    If Len(mSuggest) = 0 Then Exit Sub
    SK_SetLinkFolder mSuggest
    ReadLink
    lblStatus.Caption = "Linked: " & mInfo("folder")
End Sub

Private Sub ReadLink()
    Dim run As String
    Set mInfo = SK_LinkInfo()
    ShowLink
    If lstMacros.ListIndex >= 0 Then run = mItems(mView(lstMacros.ListIndex))("run")
    RefreshList
    If Len(run) > 0 Then SelectRun run
End Sub

Private Sub ShowAsk(ByVal e As Object, ByVal kind As String, ByVal why As String)
    Dim parts As String, files As String, p As Variant
    mAskRun = e("run")
    mAskKind = kind
    chkAskNever.Value = False
    If kind = "ask" Then
        lblAskHead.Caption = "Built for Stationkeeper"
        lblAskText.Caption = "This is built for Stationkeeper but is still functional. Would you like to run it?" & vbCrLf & vbCrLf & _
            "Without a linked effort it uses the team standards workbook, or its own defaults if there is none. " & _
            "Link a folder to point it at an effort's data."
        btnAskRun.Visible = True
        chkAskNever.Visible = True
        btnAskLink.Caption = "Link a folder..."
        btnAskRun.Default = True
    Else
        parts = SK_SkParts(e("stationkeeper"))
        For Each p In Split(parts, ",")
            files = files & IIf(Len(files) > 0, ", ", "") & SK_PartFile(p)
        Next
        lblAskHead.Caption = "Needs Stationkeeper data"
        lblAskText.Caption = "This macro is used with the Stationkeeper application. Please point to the Stationkeeper " & _
            "data folder for your effort." & vbCrLf & vbCrLf & _
            "It needs " & SK_PartNames(parts) & ". Stationkeeper writes " & IIf(InStr(parts, ",") > 0, "them", "it") & _
            " when you save the effort, to Stationkeeper Data\" & files & " inside the effort folder. " & _
            "Pick the effort folder (or the Stationkeeper Data folder in it)." & vbCrLf & vbCrLf & "Right now: " & why
        btnAskRun.Visible = False
        chkAskNever.Visible = False
        btnAskLink.Caption = "Point to the folder..."
        btnAskLink.Default = True
    End If
    fraAsk.Visible = True
    lblStatus.Caption = e("title") & ": " & IIf(kind = "ask", "run it, link a folder, or cancel.", "point to the effort folder, or cancel.")
End Sub

Private Sub HideAsk()
    fraAsk.Visible = False
    btnAskRun.Default = False
    btnAskLink.Default = False
    btnRun.Default = True
    mAskRun = ""
End Sub

Private Function AskEntry() As Object
    Dim e As Object
    For Each e In mItems
        If e("run") = mAskRun Then Set AskEntry = e: Exit Function
    Next
End Function

Private Sub btnAskRun_Click()
    Dim e As Object
    Set e = AskEntry()
    If chkAskNever.Value Then SK_SetPref "askBuilt", "0"
    HideAsk
    If Not e Is Nothing Then DoRun e
End Sub

Private Sub btnAskLink_Click()
    Dim e As Object, why As String
    Set e = AskEntry()
    If Len(SK_PickLinkFolder()) = 0 Then Exit Sub     ' cancelled: the question stays
    Set mInfo = SK_LinkInfo()
    ShowLink
    If e Is Nothing Then HideAsk: Exit Sub
    If Len(SK_Gate(e("stationkeeper"), mInfo, why)) > 0 And mAskKind = "needs" Then
        ShowAsk e, "needs", why                          ' still not there: say what is wrong now
        Exit Sub
    End If
    HideAsk
    DoRun e
End Sub

Private Sub btnAskCancel_Click()
    HideAsk
    ShowDetail
    lblStatus.Caption = "Type to search. Up and Down move, Enter runs, Esc closes."
End Sub

Private Sub btnLink_Click()
    If Len(SK_PickLinkFolder()) = 0 Then Exit Sub
    ReadLink
    lblStatus.Caption = "Linked: " & mInfo("folder")
End Sub

Private Sub btnRefresh_Click()
    ReadLink
    lblStatus.Caption = "Read the link again at " & Format$(Now, "h:nn am/pm") & "."
End Sub

Private Sub btnUnlink_Click()
    If MsgBox("Stop using " & mInfo("folder") & "?" & vbCr & vbCr & _
              "Nothing is deleted. Stationkeeper tools go back to the team standards workbook.", _
              vbYesNo + vbQuestion, SK_APP) <> vbYes Then Exit Sub
    SK_SetLinkFolder ""
    ReadLink
    lblStatus.Caption = "Unlinked."
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

Private Sub btnReports_Click()
    Dim s As String
    s = SK_ChooseReportsDir()
    If Len(s) = 0 Then Exit Sub
    lblStatus.Caption = s
    btnReports.ControlTipText = IIf(Len(SK_ReportsDir()) > 0, "Reports are saved in " & SK_ReportsDir(), "Reports open and are not saved. Choose a folder to keep them.")
End Sub

' The workbook the Stationkeeper Tools read, as they keep it for this template.
Private Function StdPath() As String
    StdPath = GetSetting("Stationkeeper Tools", "Standards", ThisDocument.Name, "")
End Function

Private Function StdTip() As String
    If Len(StdPath()) > 0 Then
        StdTip = "The tools read " & StdPath()
    Else
        StdTip = "No Global Standards workbook is chosen. Choose the Stationkeeper app's Global Standards.xlsx."
    End If
End Function

' Runs the tools' Choose Standards Workbook. The link strip above is for an effort's folder; this is the workbook.
Private Function EffPath() As String
    EffPath = GetSetting("Stationkeeper Tools", "EffortStandards", ThisDocument.Name, "")
End Function

' The two lines that say which standards the tools read.
Private Sub StdLines()
    Dim g As String, e As String, tagWords As String, noteWords As String, foundWords As String, fw() As String
    g = StdPath(): e = EffPath()
    tagWords = AdminWords("SKW_AdminTag"): noteWords = AdminWords("SKW_AdminNote")
    lblStdG.Caption = "Global Standards: " & IIf(Len(g) > 0, Mid$(g, InStrRev(g, "\") + 1), "not chosen") & IIf(Len(tagWords) > 0, " (" & tagWords & ")", "")
    lblStdG.ControlTipText = StdTip() & IIf(Len(noteWords) > 0, " " & noteWords, "")
    ' A program's workbook says so on its Read me tab (1.22.0); the tools keep the word when it is chosen.
    lblStdE.Caption = "This " & IIf(Len(e) > 0 And GetSetting("Stationkeeper Tools", "EffortKind", ThisDocument.Name, "") = "program", "program", "effort") & "'s standards: " & _
        IIf(Len(e) > 0, Mid$(e, InStrRev(e, "\") + 1), "none (the Global Standards stand alone)")
    lblStdE.ControlTipText = IIf(Len(e) > 0, "Laid over the Global Standards: " & e, "Choose the effort's Effort Standards workbook (or a program's Program Standards) to lay its values over the Global Standards.")
    ' With none pointed to, the tools use the one found by the folder of the document in front (1.23.0).
    If Len(e) = 0 Then
        foundWords = AdminWords("SKW_EffortFoundLine")
        If InStr(foundWords, "|") > 0 Then
            fw = Split(foundWords, "|")
            lblStdE.Caption = "This " & fw(0) & "'s standards: " & fw(1) & " (found by the document's folder)"
            lblStdE.ControlTipText = "Nobody pointed to a workbook, so the tools use the one found in the folder of the document in front or up to three folders above it. Pointing to one wins."
        End If
    End If
    btnStd.ControlTipText = StdTip()
    btnEffClear.Enabled = (Len(e) > 0)
End Sub

Private Sub btnStd_Click()
    On Error Resume Next
    Application.Run "SKW_Setup.ChooseStandardsWorkbook"
    If Err.Number <> 0 Then
        Err.Clear
        lblStatus.Caption = "Stationkeeper Tools are not in this template. Install them from the admin first."
        Exit Sub
    End If
    On Error GoTo 0
    StdLines
    RefreshList
End Sub

Private Sub btnEff_Click()
    On Error Resume Next
    Application.Run "SKW_Setup.ChooseEffortStandardsWorkbook"
    If Err.Number <> 0 Then
        Err.Clear
        lblStatus.Caption = "Stationkeeper Tools are not in this template. Install them from the admin first."
        Exit Sub
    End If
    On Error GoTo 0
    StdLines
End Sub

Private Sub btnEffClear_Click()
    Dim msg As String
    On Error Resume Next
    msg = Application.Run("SKW_Setup.SKW_ChooseEffortStandards", "")
    On Error GoTo 0
    StdLines
    lblStatus.Caption = "No effort's workbook is pointed to. The tools look beside the document in front."
End Sub

' Tests: the Global Standards button.
Public Function TestStd() As String
    StdLines
    TestStd = lblStdG.Caption & " | " & lblStdE.Caption & " | " & btnStd.Caption & "/" & btnEff.Caption & "/" & btnEffClear.Caption & " clear=" & btnEffClear.Enabled & _
        " | order=" & (lblStdG.Top < lblStdE.Top And lblStdE.Top < lblLink.Top) & " filters=" & (cboShow.Top = cboCat.Top And cboCat.Top = cboSort.Top)
End Function

' Tests: what the list holds for this person: every row, the Capture tools among them, and the standards line.
Public Function TestAdmin() As String
    Dim i As Long, n As Long
    txtSearch.Text = ""
    SetShowChoice "All"
    StdLines
    RefreshList
    For i = 0 To lstMacros.ListCount - 1
        If LCase$(lstMacros.List(i, 0)) Like "capture *" Then n = n + 1
    Next
    TestAdmin = "rows=" & lstMacros.ListCount & " capture=" & n & " line=" & lblStdG.Caption
End Function

' Tests: the list under an Application choice: count, then the first titles.
Public Function TestApp(ByVal appName As String) As String
    Dim s As String, i As Long
    txtSearch.Text = ""
    SetShowChoice appName
    s = cboShow.Text & " " & lstMacros.ListCount & ":"
    For i = 0 To lstMacros.ListCount - 1
        If i < 3 Then s = s & " [" & lstMacros.List(i, 0) & "]"
    Next
    TestApp = s
End Function

' Tests: the first titles under a Sort choice.
Public Function TestSort(ByVal sortName As String) As String
    Dim s As String, i As Long
    txtSearch.Text = ""
    SetShowChoice "All"
    For i = 0 To cboSort.ListCount - 1
        s = s & IIf(i > 0, "|", "") & cboSort.List(i)
        If cboSort.List(i) = sortName Then cboSort.ListIndex = i
    Next
    s = s & " -> " & cboSort.Text & ":"
    For i = 0 To lstMacros.ListCount - 1
        If i < 3 Then s = s & " [" & lstMacros.List(i, 0) & "]"
    Next
    TestSort = s
End Function

' Tests: the reports button.
Public Function TestReports() As String
    TestReports = btnReports.Caption & " | " & btnReports.ControlTipText & " | left=" & btnReports.Left
End Function

Private Sub btnClose_Click()
    If fraAsk.Visible Then btnAskCancel_Click: Exit Sub   ' Esc closes the question first
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
        s = s & " [" & lstMacros.List(i, 0) & " / " & lstMacros.List(i, 2) & "]"
    Next
    TestSearch = s
End Function

' Picks an Application choice by name. For the tests, "Stationkeeper" and "Other" list only the macros made
' for Stationkeeper or only the others, across every application; "All macros" lists everything.
Private Sub SetShowChoice(ByVal flt As String)
    Dim i As Long, pick As Long
    mKind = ""
    If flt = "Stationkeeper" Or flt = "Other" Then mKind = flt
    For i = 0 To cboShow.ListCount - 1
        If LCase$(cboShow.List(i)) = LCase$(flt) Then pick = i
    Next
    If cboShow.ListIndex = pick Then RefreshList Else cboShow.ListIndex = pick
End Sub

' The Show choices, joined with "|".
Public Function TestShowChoices() As String
    Dim s As String, i As Long
    For i = 0 To cboShow.ListCount - 1
        s = s & IIf(i > 0, "|", "") & cboShow.List(i)
    Next
    TestShowChoices = s
End Function

' The list under a Show filter: count, then title and badge (SK- for built, SK+ for needs data).
Public Function TestShow(ByVal flt As String) As String
    Dim s As String, i As Long, b As String
    txtSearch.Text = ""
    SetShowChoice flt
    s = lstMacros.ListCount & ":"
    For i = 0 To lstMacros.ListCount - 1
        b = lstMacros.List(i, 1)
        b = Replace(Replace(b, ChrW$(&H25C7) & " SK", "SK-"), ChrW$(&H25C6) & " SK", "SK+")
        s = s & " [" & lstMacros.List(i, 0) & IIf(Len(b) > 0, " " & b, "") & "]"
    Next
    TestShow = s
End Function

' The list under a Show choice and a Category choice; the category choices first.
Public Function TestCat(ByVal flt As String, ByVal cat As String) As String
    Dim s As String, i As Long
    For i = 0 To cboCat.ListCount - 1
        s = s & cboCat.List(i) & "|"
    Next
    txtSearch.Text = ""
    SetShowChoice flt
    cboCat.Text = cat
    s = s & " -> " & lstMacros.ListCount & ":"
    For i = 0 To lstMacros.ListCount - 1
        s = s & " [" & lstMacros.List(i, 0) & "]"
    Next
    cboCat.ListIndex = 0
    cboShow.ListIndex = 0
    TestCat = s
End Function

' What Run does for a macro, without running it: the panel it shows, or "run".
Public Function TestGate(ByVal run As String) As String
    Dim e As Object, why As String, g As String
    For Each e In mItems
        If LCase$(e("run")) = LCase$(run) Then Exit For
    Next
    If e Is Nothing Then TestGate = "not in the menu": Exit Function
    Set mInfo = SK_LinkInfo()
    ShowLink
    g = SK_Gate(e("stationkeeper"), mInfo, why)
    If Len(g) = 0 Then TestGate = "run": Exit Function
    ShowAsk e, g, why
    TestGate = g & " | " & lblAskHead.Caption & " | " & Replace(lblAskText.Caption, vbCrLf, " / ") & " | buttons: " & _
        IIf(btnAskRun.Visible, btnAskRun.Caption & ", ", "") & btnAskLink.Caption & ", " & btnAskCancel.Caption & _
        IIf(chkAskNever.Visible, " | " & chkAskNever.Caption, "")
    HideAsk
End Function

' Tests: what the strip offers for a document in docFolder, then (use=True) clicks Use this folder.
Public Function TestSuggest(ByVal docFolder As String, ByVal useIt As Boolean) As String
    mDocPath = docFolder
    Set mInfo = SK_LinkInfo()
    ShowLink
    TestSuggest = "suggest=" & IIf(btnSuggest.Visible, lblSuggest.Caption & " [" & btnSuggest.Caption & "]", "none")
    If useIt And btnSuggest.Visible Then
        btnSuggest_Click
        TestSuggest = TestSuggest & " || after: " & Replace(lblLink.Caption, vbCrLf, " / ") & " | suggest now=" & IIf(btnSuggest.Visible, "shown", "none")
    End If
End Function

Public Function TestLink() As String
    TestLink = Replace(lblLink.Caption, vbCrLf, " / ") & " | button=" & btnLink.Caption & " unlink=" & btnUnlink.Enabled & _
        " | legend=" & Replace(Replace(lblLegend.Caption, ChrW$(&H25C7), "<o>"), ChrW$(&H25C6), "<f>")
End Function
