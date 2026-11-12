' =====================================================================
' Stationkeeper Tools: Deck Media List
' Every picture, video, audio, chart, SmartArt, icon and embedded or
' linked object in the deck open in PowerPoint, by slide. Rename them
' (the names show in PowerPoint's Selection Pane and Animation Pane),
' name them from their alt text, and jump to them. Renaming changes the
' open deck's shape names only; save the deck in PowerPoint to keep them.
' The controls are built in code, so this form needs no designer work.
' =====================================================================
Option Explicit

Private WithEvents cboKind As MSForms.ComboBox
Private WithEvents txtFind As MSForms.TextBox
Private WithEvents lstMedia As MSForms.ListBox
Private WithEvents txtName As MSForms.TextBox
Private WithEvents btnRename As MSForms.CommandButton
Private WithEvents btnSuggest As MSForms.CommandButton
Private WithEvents btnNameAll As MSForms.CommandButton
Private WithEvents btnGoTo As MSForms.CommandButton
Private WithEvents btnRefresh As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton
Private lblAlt As MSForms.Label
Private lblStatus As MSForms.Label
Private lblCount As MSForms.Label

Private mDeck As Object
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

Private Function AddBtn(ByVal nm As String, ByVal cap As String, ByVal l As Single, ByVal t As Single, _
                        ByVal w As Single, ByVal h As Single) As MSForms.CommandButton
    Dim c As MSForms.CommandButton
    Set c = AddCtl("Forms.CommandButton.1", nm, l, t, w, h)
    c.Caption = cap
    Set AddBtn = c
End Function

Private Sub UserForm_Initialize()
    Dim v As Variant
    mBusy = True
    Me.Width = 700
    Me.Height = 470
    Me.BackColor = &HFFFFFF
    AddLabel "lblKind", "Show", 10, 8, 80, 14
    Set cboKind = AddCtl("Forms.ComboBox.1", "cboKind", 10, 22, 150, 20)
    cboKind.Style = fmStyleDropDownList
    For Each v In Array("All media", "Pictures and icons", "Video and audio", "Charts and SmartArt", "Embedded and linked", "No alt text", "Default names")
        cboKind.AddItem v
    Next
    cboKind.ListIndex = 0
    AddLabel "lblFind", "Find (name or alt text)", 168, 8, 200, 14
    Set txtFind = AddCtl("Forms.TextBox.1", "txtFind", 168, 22, 220, 20)
    Set lblCount = AddLabel("lblCount", "", 396, 26, 280, 14)
    lblCount.ForeColor = &H808080
    AddLabel "hd1", "Slide", 14, 50, 40, 12
    AddLabel "hd2", "Name (as in the Selection Pane)", 52, 50, 200, 12
    AddLabel "hd3", "Kind", 262, 50, 100, 12
    AddLabel "hd4", "Size", 362, 50, 80, 12
    AddLabel "hd5", "Alt text", 446, 50, 200, 12
    Set lstMedia = AddCtl("Forms.ListBox.1", "lstMedia", 10, 64, 670, 250)
    lstMedia.ColumnCount = 5
    lstMedia.ColumnWidths = "38 pt;208 pt;98 pt;82 pt;230 pt"
    lstMedia.IntegralHeight = False
    Set lblAlt = AddLabel("lblAlt", "", 10, 320, 670, 28)
    lblAlt.WordWrap = True
    lblAlt.ForeColor = &H606060
    AddLabel "lblName", "Name", 10, 352, 60, 14
    Set txtName = AddCtl("Forms.TextBox.1", "txtName", 10, 366, 320, 20)
    Set btnRename = AddBtn("btnRename", "Rename", 336, 364, 70, 24)
    btnRename.Default = True
    Set btnSuggest = AddBtn("btnSuggest", "Suggest a name", 412, 364, 100, 24)
    Set btnGoTo = AddBtn("btnGoTo", "Go to it", 518, 364, 70, 24)
    Set btnNameAll = AddBtn("btnNameAll", "Name every default-named item from its alt text", 10, 396, 300, 24)
    Set btnRefresh = AddBtn("btnRefresh", "Refresh", 518, 396, 70, 24)
    Set btnClose = AddBtn("btnClose", "Close", 600, 396, 80, 24)
    btnClose.Cancel = True
    Set lblStatus = AddLabel("lblStatus", "Pick an item, type a name, Rename. Double-click goes to it. Save the deck in PowerPoint to keep names.", 10, 428, 670, 14)
    lblStatus.ForeColor = &H808080
    Set mDeck = SKP_Deck.SKP_Deck()
    mBusy = False
    ReadItems
    frmSKWBrand.Dress Me, "Deck Media List", "Stationkeeper Tools   |   pictures, video, charts and objects by slide   |   " & SKW_VERSION, "btnRename"
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    frmSKWBrand.SetIcon Me
End Sub

' For the tests: what the Stationkeeper look put on this form.
Public Function TestBrand() As String
    TestBrand = frmSKWBrand.Probe(Me) & " height=" & Me.Height & " listTop=" & lstMedia.Top
End Function

Private Sub ReadItems()
    If mDeck Is Nothing Then
        Set mItems = New Collection
        Me.Caption = "Deck Media List"
        lblStatus.Caption = SKP_Why()
    Else
        Set mItems = SKP_MediaItems(mDeck)
        Me.Caption = "Deck Media List   |   " & mDeck.Name
    End If
    FillList
End Sub

Private Function IsDefaultName(ByVal it As Object) As Boolean
    Dim nm As String, w As Variant
    nm = LCase$(it("name"))
    For Each w In Array("picture ", "graphic ", "content placeholder ", "chart ", "diagram ", "video ", "audio ", "object ", "media ", "picture placeholder ", "online media ")
        If Left$(nm, Len(w)) = w And IsNumeric(Mid$(nm, Len(w) + 1)) Then IsDefaultName = True: Exit Function
    Next
End Function

Private Function Passes(ByVal it As Object) As Boolean
    Dim k As String, q As String
    k = it("kind")
    Select Case cboKind.Text
        Case "Pictures and icons": Passes = (k = "Picture" Or k = "Linked picture" Or k = "Icon or SVG" Or k = "Linked graphic")
        Case "Video and audio": Passes = (k = "Video" Or k = "Audio" Or k = "Media" Or k = "Web video")
        Case "Charts and SmartArt": Passes = (k = "Chart" Or k = "SmartArt")
        Case "Embedded and linked": Passes = (k = "Embedded object" Or k = "Linked object" Or Len(it("link")) > 0)
        Case "No alt text": Passes = (Len(Trim$(it("alt"))) = 0)
        Case "Default names": Passes = IsDefaultName(it)
        Case Else: Passes = True
    End Select
    q = LCase$(Trim$(txtFind.Text))
    If Passes And Len(q) > 0 Then Passes = (InStr(LCase$(it("name") & " " & it("alt")), q) > 0)
End Function

Private Sub FillList(Optional ByVal keepId As Long = 0)
    Dim i As Long, n As Long, it As Object, sel As Long
    mBusy = True
    lstMedia.Clear
    ReDim mView(0 To 0)
    sel = -1
    For i = 1 To mItems.Count
        Set it = mItems(i)
        If Passes(it) Then
            lstMedia.AddItem it("slide")
            lstMedia.List(n, 1) = it("name")
            lstMedia.List(n, 2) = it("kind")
            lstMedia.List(n, 3) = it("size")
            lstMedia.List(n, 4) = IIf(Len(Trim$(it("alt"))) > 0, Replace(Replace(it("alt"), vbCr, " "), vbLf, " "), "(none)")
            ReDim Preserve mView(0 To n)
            mView(n) = i
            If it("id") = keepId Then sel = n
            n = n + 1
        End If
    Next
    lblCount.Caption = n & " of " & mItems.Count & " media item(s)"
    mBusy = False
    If n > 0 Then lstMedia.ListIndex = IIf(sel >= 0, sel, 0) Else ShowItem
End Sub

Private Function Cur() As Object
    If lstMedia.ListIndex < 0 Or mItems.Count = 0 Then Exit Function
    Set Cur = mItems(mView(lstMedia.ListIndex))
End Function

Private Sub ShowItem()
    Dim it As Object
    Set it = Cur()
    If it Is Nothing Then
        txtName.Text = "": lblAlt.Caption = ""
        Exit Sub
    End If
    txtName.Text = it("name")
    lblAlt.Caption = "Slide " & it("slide") & ", " & LCase$(it("kind")) & ". Alt text: " & IIf(Len(Trim$(it("alt"))) > 0, it("alt"), "(none: screen readers skip it)") & _
        IIf(Len(it("link")) > 0, "   Linked to: " & it("link"), "")
End Sub

Private Sub lstMedia_Change()
    If Not mBusy Then ShowItem
End Sub

Private Sub lstMedia_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnGoTo_Click
End Sub

Private Sub cboKind_Change()
    If Not mBusy Then FillList
End Sub

Private Sub txtFind_Change()
    If Not mBusy Then FillList
End Sub

Private Sub btnRename_Click()
    Dim it As Object, msg As String
    Set it = Cur()
    If it Is Nothing Then Exit Sub
    msg = SKP_RenameMedia(mDeck, it("slideId"), it("id"), txtName.Text)
    If Len(msg) > 0 Then lblStatus.Caption = msg: Beep: Exit Sub
    it("name") = Trim$(txtName.Text)
    lblStatus.Caption = "Renamed. Save the deck in PowerPoint to keep it."
    FillList it("id")
End Sub

Private Sub btnSuggest_Click()
    Dim it As Object
    Set it = Cur()
    If it Is Nothing Then Exit Sub
    txtName.Text = SKP_SuggestName(it)
    lblStatus.Caption = "Edit the name if you like, then Rename."
End Sub

Private Sub btnNameAll_Click()
    Dim it As Object, n As Long, nSkip As Long, msg As String
    For Each it In mItems
        If IsDefaultName(it) And Len(Trim$(it("alt"))) > 0 Then n = n + 1
    Next
    If n = 0 Then lblStatus.Caption = "No item has both a default name (like Picture 3) and alt text.": Exit Sub
    If Not SKW_IsQuiet() Then
        If MsgBox("Rename " & n & " item(s) from their alt text (for example ""S3 Picture - Rocket on the pad"")?", vbYesNo + vbQuestion, "Deck Media List") <> vbYes Then Exit Sub
    End If
    n = 0
    For Each it In mItems
        If IsDefaultName(it) And Len(Trim$(it("alt"))) > 0 Then
            msg = SKP_RenameMedia(mDeck, it("slideId"), it("id"), SKP_SuggestName(it))
            If Len(msg) = 0 Then it("name") = SKP_SuggestName(it): n = n + 1 Else nSkip = nSkip + 1
        End If
    Next
    lblStatus.Caption = n & " renamed" & IIf(nSkip > 0, ", " & nSkip & " skipped (a name already on that slide)", "") & ". Save the deck in PowerPoint to keep them."
    FillList
End Sub

Private Sub btnGoTo_Click()
    Dim it As Object, msg As String
    Set it = Cur()
    If it Is Nothing Then Exit Sub
    msg = SKP_GoToMedia(mDeck, it("slideId"), it("id"))
    lblStatus.Caption = IIf(Len(msg) > 0, msg, "Selected on slide " & it("slide") & " in PowerPoint.")
End Sub

Private Sub btnRefresh_Click()
    Set mDeck = SKP_Deck.SKP_Deck()
    ReadItems
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' ---------------------------------------------------------------------
' Tests
' ---------------------------------------------------------------------
Public Function TestList(ByVal kind As String) As String
    Dim i As Long, s As String
    cboKind.Text = kind
    s = lstMedia.ListCount & ":"
    For i = 0 To lstMedia.ListCount - 1
        s = s & " [" & lstMedia.List(i, 0) & "/" & lstMedia.List(i, 1) & "/" & lstMedia.List(i, 2) & "/" & lstMedia.List(i, 4) & "]"
    Next
    cboKind.ListIndex = 0
    TestList = s
End Function

Public Function TestRename(ByVal i As Long, ByVal nm As String) As String
    lstMedia.ListIndex = i
    txtName.Text = nm
    btnRename_Click
    TestRename = lblStatus.Caption & " | now " & lstMedia.List(lstMedia.ListIndex, 1)
End Function

Public Function TestNameAll() As String
    btnNameAll_Click
    TestNameAll = lblStatus.Caption
End Function
