' =====================================================================
' Stationkeeper Tools: Remap Style
' Pick a style from the ones in the document. The window lists each
' kind of paragraph in that style (not in a list; bulleted or numbered
' by hand, with its level; in a table) and, beside each, a list of the
' document's styles to pick what it becomes. A bulleted paragraph starts
' on the list style the Global Standards give its level. Nothing picked
' means left alone. The lists show the styles the document uses; a tick box
' adds Word's built-in styles it has not used yet. Each Apply is one
' Ctrl+Z. Nothing is stored.
' The controls are built in code, so this form needs no designer work.
' =====================================================================
Option Explicit

Private Const MAXR As Long = 12
Private Const LEAVE As String = "(leave alone)"
Private Const ROW_TOP As Single = 76
Private Const ROW_H As Single = 22

Private WithEvents cboFrom As MSForms.ComboBox
Private WithEvents btnApply As MSForms.CommandButton
Private WithEvents btnReport As MSForms.CommandButton
Private WithEvents btnRead As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton
Private WithEvents chkBuilt As MSForms.CheckBox
Private optAll As MSForms.OptionButton
Private optText As MSForms.OptionButton
Private optNone As MSForms.OptionButton
Private optDoc As MSForms.OptionButton
Private optSel As MSForms.OptionButton
Private lblStatus As MSForms.Label
Private lblNone As MSForms.Label
Private mKeys(1 To 12) As String
Private mN As Long
Private mBusy As Boolean
Private mBand As Single

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
    c.BackStyle = 0
    Set AddLabel = c
End Function

Private Function AddBtn(ByVal nm As String, ByVal cap As String, ByVal l As Single, ByVal t As Single, _
                        ByVal w As Single, ByVal h As Single) As MSForms.CommandButton
    Dim c As MSForms.CommandButton
    Set c = AddCtl("Forms.CommandButton.1", nm, l, t, w, h)
    c.Caption = cap
    Set AddBtn = c
End Function

Private Function AddOpt(ByVal nm As String, ByVal cap As String, ByVal grp As String, ByVal l As Single, ByVal t As Single, _
                        ByVal w As Single) As MSForms.OptionButton
    Dim c As MSForms.OptionButton
    Set c = AddCtl("Forms.OptionButton.1", nm, l, t, w, 16)
    c.Caption = cap
    c.GroupName = grp
    c.BackStyle = 0
    Set AddOpt = c
End Function

Private Sub UserForm_Initialize()
    Dim i As Long, c As MSForms.ComboBox, lb As MSForms.Label
    mBusy = True
    Me.Caption = "Remap Style"
    Me.Width = 480
    Me.Height = 330
    Me.BackColor = &HFFFFFF
    AddLabel "lblFrom", "Style to change", 10, 8, 200, 14
    Set cboFrom = AddCtl("Forms.ComboBox.1", "cboFrom", 10, 22, 250, 20)
    cboFrom.Style = fmStyleDropDownList
    Set btnRead = AddBtn("btnRead", "Read again", 268, 21, 76, 22)
    btnRead.ControlTipText = "Reads the document again, after you have changed it"
    Set chkBuilt = AddCtl("Forms.CheckBox.1", "chkBuilt", 352, 23, 110, 18)
    chkBuilt.Caption = "All Word styles"
    chkBuilt.BackStyle = 0
    chkBuilt.ControlTipText = "The Becomes lists show the styles this document uses. Tick this to add Word's built-in styles it has not used yet, such as Body Text."
    Set lb = AddLabel("hdWhat", "What the paragraph is", 10, 56, 170, 14)
    lb.Font.Bold = True
    Set lb = AddLabel("hdCount", "Paragraphs", 184, 56, 90, 14)
    lb.Font.Bold = True
    Set lb = AddLabel("hdTo", "Becomes", 278, 56, 180, 14)
    lb.Font.Bold = True
    For i = 1 To MAXR
        AddLabel "lblRow" & i, "", 10, ROW_TOP + (i - 1) * ROW_H + 3, 170, 14
        AddLabel "lblCnt" & i, "", 184, ROW_TOP + (i - 1) * ROW_H + 3, 90, 14
        Set c = AddCtl("Forms.ComboBox.1", "cboTo" & i, 278, ROW_TOP + (i - 1) * ROW_H, 180, 20)
        c.Style = fmStyleDropDownList
    Next
    Set lblNone = AddLabel("lblNone", "", 10, ROW_TOP + 3, 448, 28)
    lblNone.WordWrap = True
    lblNone.ForeColor = &H808080
    Set lb = AddLabel("hdKeep", "Formatting applied by hand", 10, 0, 220, 14)
    lb.Font.Bold = True
    Set optAll = AddOpt("optAll", "Keep all of it", "keep", 10, 0, 220)
    Set optText = AddOpt("optText", "Keep emphasis only (bold, italic, color...)", "keep", 10, 0, 240)
    Set optNone = AddOpt("optNone", "Clear it: the style's look only", "keep", 10, 0, 220)
    optText.ControlTipText = "Bold, italic, underline, strikethrough, superscript, subscript, color and highlight stay. Fonts, sizes, indents and spacing take the style's."
    optAll.Value = True
    Set lb = AddLabel("hdScope", "Apply to", 278, 0, 180, 14)
    lb.Font.Bold = True
    Set optDoc = AddOpt("optDoc", "The whole document", "scope", 278, 0, 180)
    Set optSel = AddOpt("optSel", "The selected paragraphs only", "scope", 278, 0, 180)
    optDoc.Value = True
    Set lblStatus = AddLabel("lblStatus", "", 10, 0, 448, 40)
    lblStatus.WordWrap = True
    Set btnApply = AddBtn("btnApply", "Apply", 10, 0, 84, 26)
    btnApply.Default = True
    Set btnReport = AddBtn("btnReport", "Report...", 100, 0, 84, 26)
    btnReport.ControlTipText = "A report of every kind of paragraph in this style. Changes nothing."
    Set btnClose = AddBtn("btnClose", "Close", 388, 0, 70, 26)
    btnClose.Cancel = True
    frmSKWBrand.Dress Me, "Remap Style", "Stationkeeper Tools   |   pick a style, then what each kind of paragraph becomes   |   " & SKW_VERSION, "btnApply"
    mBand = Me.Controls("lblFrom").Top - 8
    FillFrom ""
    mBusy = False
    ReadRows
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    frmSKWBrand.SetIcon Me
End Sub

' The styles some text in the document is in. Keeps the pick when it is still there, else the style at the cursor.
Private Sub FillFrom(ByVal want As String)
    Dim v As Variant, i As Long, at As String
    mBusy = True
    cboFrom.Clear
    For Each v In SKW_RemapStyleNames(True)
        cboFrom.AddItem v
    Next
    On Error Resume Next
    If Len(want) = 0 Then want = SKW_ParaStyleName(Selection.Range)
    On Error GoTo 0
    For i = 0 To cboFrom.ListCount - 1
        If StrComp(cboFrom.List(i), want, vbTextCompare) = 0 Then cboFrom.ListIndex = i: Exit For
    Next
    If cboFrom.ListIndex < 0 And cboFrom.ListCount > 0 Then cboFrom.ListIndex = 0
    mBusy = False
End Sub

Private Sub PutAt(ByVal nm As String, ByVal t As Single)
    On Error Resume Next
    Me.Controls(nm).Top = t
End Sub

' Reads the document for the style picked and fills the rows.
Private Sub ReadRows()
    Dim rows As Collection, v As Variant, i As Long, c As MSForms.ComboBox, names As Collection, n As Variant, y As Single, k As Long, knd As String
    mBusy = True
    mN = 0
    Set rows = New Collection
    If cboFrom.ListIndex >= 0 Then Set rows = SKW_RemapRows(cboFrom.Text)
    Set names = SKW_RemapStyleNames(False, chkBuilt.Value)
    For i = 1 To MAXR
        Set c = Me.Controls("cboTo" & i)
        c.Clear
        If i <= rows.Count Then
            v = rows(i)
            mN = i
            mKeys(i) = v(0)
            knd = Split(v(0), "|")(1)
            Me.Controls("lblRow" & i).Caption = v(1) & IIf(v(2) > 0, ", level " & v(2), "")
            Me.Controls("lblCnt" & i).Caption = v(3) & IIf(v(4) > 0, "  (" & v(4) & " by hand)", "")
            Me.Controls("lblCnt" & i).ControlTipText = v(4) & " of them carry formatting applied by hand"
            c.AddItem LEAVE
            If knd <> "typed" Then
                For Each n In names
                    If StrComp(n, cboFrom.Text, vbTextCompare) <> 0 Then c.AddItem n
                Next
            End If
            c.ListIndex = 0
            For k = 1 To c.ListCount - 1
                If Len(v(5)) > 0 And StrComp(c.List(k), v(5), vbTextCompare) = 0 Then c.ListIndex = k: Exit For
            Next
            c.Enabled = (knd <> "typed")
            c.ControlTipText = IIf(knd = "typed", "The bullet is typed text, not a list: it is left alone", _
                IIf(c.ListIndex > 0, "From the Global Standards' List numbering sheet. Pick another to change it for this run.", IIf(Len(v(6)) > 0, v(6), "")))
        End If
        Me.Controls("lblRow" & i).Visible = (i <= rows.Count)
        Me.Controls("lblCnt" & i).Visible = (i <= rows.Count)
        c.Visible = (i <= rows.Count)
    Next
    lblNone.Visible = (mN = 0)
    lblNone.Caption = IIf(cboFrom.ListCount = 0, "This document has no text to change.", "Nothing in this document is in " & cboFrom.Text & " now.")
    ' The rest of the window follows the rows.
    k = IIf(mN = 0, 2, mN)
    y = mBand + ROW_TOP + k * ROW_H + 8
    PutAt "hdKeep", y: PutAt "hdScope", y
    PutAt "optAll", y + 16: PutAt "optDoc", y + 16
    PutAt "optText", y + 34: PutAt "optSel", y + 34
    PutAt "optNone", y + 52
    PutAt "lblStatus", y + 74
    PutAt "btnApply", y + 118: PutAt "btnReport", y + 118: PutAt "btnClose", y + 118
    Me.Height = y + 118 + 26 + 34
    btnApply.Enabled = (mN > 0)
    btnReport.Enabled = (mN > 0)
    mBusy = False
End Sub

Private Sub cboFrom_Change()
    If mBusy Then Exit Sub
    lblStatus.Caption = ""
    ReadRows
End Sub

Private Sub chkBuilt_Click()
    If mBusy Then Exit Sub
    ReadRows
End Sub

Private Sub btnRead_Click()
    Dim want As String
    want = cboFrom.Text
    FillFrom want
    lblStatus.Caption = ""
    ReadRows
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnReport_Click()
    If cboFrom.ListIndex < 0 Then Exit Sub
    SKW_Show SKW_RemapCheck(cboFrom.Text & ">")
End Sub

Private Function Picks() As String
    Dim i As Long, c As MSForms.ComboBox, s As String
    For i = 1 To mN
        Set c = Me.Controls("cboTo" & i)
        If c.ListIndex > 0 Then s = s & vbLf & mKeys(i) & vbTab & c.Text
    Next
    Picks = Mid$(s, 2)
End Function

Private Sub btnApply_Click()
    DoApply
End Sub

Private Sub DoApply()
    Dim p As String, keep As String, msg As String, want As String
    If cboFrom.ListIndex < 0 Then Exit Sub
    p = Picks()
    If Len(p) = 0 Then
        lblStatus.Caption = "Pick a style in the Becomes column first. A row left on " & LEAVE & " is not changed."
        Exit Sub
    End If
    keep = IIf(optText.Value, "text", IIf(optNone.Value, "none", "all"))
    want = cboFrom.Text
    msg = SKW_RemapApply(want & ">", keep, IIf(optSel.Value, "selection", "all"), "", p)
    FillFrom want
    ReadRows
    lblStatus.Caption = Replace(msg, vbCr, "  ")
End Sub

' ---- for the tests ----
Public Function TestBrand() As String
    TestBrand = frmSKWBrand.Probe(Me) & " width=" & Me.Width
End Function

' Picks the style to change and lists the rows: [what | count | picked].
Public Function TestRows(ByVal fromStyle As String) As String
    Dim i As Long, s As String, c As MSForms.ComboBox, hasN As Boolean
    FillFrom fromStyle
    ReadRows
    For i = 0 To cboFrom.ListCount - 1
        s = s & "," & cboFrom.List(i)
    Next
    s = "from=" & cboFrom.Text & " of " & Mid$(s, 2) & " rows=" & mN
    For i = 1 To mN
        Set c = Me.Controls("cboTo" & i)
        s = s & " [" & Me.Controls("lblRow" & i).Caption & " | " & Me.Controls("lblCnt" & i).Caption & " | " & c.Text & IIf(c.Enabled, "", " off") & " | choices=" & (c.ListCount > 3) & "]"
    Next
    TestRows = s & " apply=" & btnApply.Enabled
End Function

' Picks what one row becomes, by the row's words.
Public Function TestPick(ByVal rowWords As String, ByVal toStyle As String) As String
    Dim i As Long, k As Long, c As MSForms.ComboBox
    TestPick = "no such row or style"
    For i = 1 To mN
        If Me.Controls("lblRow" & i).Caption = rowWords Then
            Set c = Me.Controls("cboTo" & i)
            For k = 0 To c.ListCount - 1
                If c.List(k) = toStyle Then c.ListIndex = k: TestPick = rowWords & " -> " & c.Text
            Next
        End If
    Next
End Function

Public Function TestAllStyles(ByVal turnOn As Boolean) As String
    Dim c As MSForms.ComboBox
    chkBuilt.Value = turnOn
    ReadRows
    Set c = Me.Controls("cboTo1")
    TestAllStyles = "choices=" & c.ListCount
End Function

Public Function TestApply(ByVal keep As String, ByVal sel As Boolean) As String
    optAll.Value = (keep = "all"): optText.Value = (keep = "text"): optNone.Value = (keep = "none")
    optSel.Value = sel: optDoc.Value = Not sel
    DoApply
    TestApply = lblStatus.Caption & " | rows now=" & mN
End Function
