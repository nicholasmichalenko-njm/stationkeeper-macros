' =====================================================================
' Stationkeeper Tools: the quick values form (1.10.0).
' The few values the quick tools use: table styles and cell margins,
' caption styles, and the caption numbering rule. Saved on this PC.
' They are not the Global Standards and are never written to them.
' The controls are built in code, so this form needs no designer work.
' =====================================================================
Option Explicit

Public Saved As Boolean

Private WithEvents btnSave As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private WithEvents btnFill As MSForms.CommandButton
Private cboHead As MSForms.ComboBox
Private cboMode As MSForms.ComboBox
Private lblFrom As MSForms.Label

Private Const C_NAVY As Long = &H603416
Private Const C_PAGE As Long = &HFCF8F5
Private Const C_INK As Long = &H2E2017
Private Const X_LABEL As Single = 12
Private Const X_FIELD As Single = 214
Private Const W_FIELD As Single = 330
Private Const ROW_H As Single = 21

Private mTop As Single

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
    c.ForeColor = C_INK
    Set AddLabel = c
End Function

Private Sub Heading(ByVal cap As String)
    Dim c As MSForms.Label
    mTop = mTop + 6
    Set c = AddLabel("lblH" & CStr(CLng(mTop)), cap, X_LABEL, mTop, 520, 16)
    c.Font.Bold = True
    c.Font.Size = 10
    c.ForeColor = C_NAVY
    mTop = mTop + 19
End Sub

' One row: a label and a text box named txt_<key>.
Private Sub Row(ByVal key As String, ByVal cap As String)
    Dim c As Object
    AddLabel "lbl_" & key, cap, X_LABEL, mTop + 3, X_FIELD - X_LABEL - 6, 14
    Set c = AddCtl("Forms.TextBox.1", "txt_" & key, X_FIELD, mTop, W_FIELD, 18)
    mTop = mTop + ROW_H
End Sub

Private Sub UserForm_Initialize()
    Dim k As Variant, i As Long, c As Object, x As Single
    Me.Caption = "Stationkeeper Tools: Quick Values"
    Me.Width = 572
    Me.BackColor = C_PAGE
    mTop = 8
    Set c = AddLabel("lblNote", "The quick tools use these values. They are saved on this PC. They are not the Global Standards and are never " & _
        "written to them. A blank box leaves that setting as it is in the document.", X_LABEL, mTop, 532, 26)
    mTop = mTop + 30

    Heading "Tables"
    Row "head", "Header row paragraph style"
    Row "body", "Body rows paragraph style"
    Row "keep", "Styles left alone in body rows ( ; )"
    Row "tstyle", "Table style (optional)"
    AddLabel "lbl_mar", "Cell margins, inches", X_LABEL, mTop + 3, X_FIELD - X_LABEL - 6, 14
    x = X_FIELD
    For Each k In Array("mtop|Top", "mbottom|Bottom", "mleft|Left", "mright|Right")
        AddLabel "lbl_" & Split(k, "|")(0), Split(k, "|")(1), x, mTop + 3, 36, 14
        Set c = AddCtl("Forms.TextBox.1", "txt_" & Split(k, "|")(0), x + 36, mTop, 40, 18)
        x = x + 84
    Next
    mTop = mTop + ROW_H
    AddLabel "lbl_headrow", "The header row is", X_LABEL, mTop + 3, X_FIELD - X_LABEL - 6, 14
    Set cboHead = AddCtl("Forms.ComboBox.1", "cboHead", X_FIELD, mTop, W_FIELD, 18)
    cboHead.Style = 2
    cboHead.AddItem "Row 2 when row 1 holds a caption number, else row 1"
    cboHead.AddItem "Always row 2"
    mTop = mTop + ROW_H

    Heading "Captions"
    Row "captable", "Table caption paragraph style"
    Row "capfigure", "Figure caption paragraph style"
    Row "captablechar", "Table caption character style (optional)"
    Row "capfigurechar", "Figure caption character style (optional)"
    Row "capaction", "Action caption character style"
    Row "abbrev", "Periods that do not end a title ( ; )"

    Heading "Caption numbering"
    AddLabel "lbl_nummode", "The rule for this work", X_LABEL, mTop + 3, X_FIELD - X_LABEL - 6, 14
    Set cboMode = AddCtl("Forms.ComboBox.1", "cboMode", X_FIELD, mTop, W_FIELD, 18)
    cboMode.Style = 2
    cboMode.AddItem "Heading number, then sequence (Table 3-2)"
    cboMode.AddItem "Sequential all the way through (Table 1, Table 2)"
    mTop = mTop + ROW_H
    Row "numhead", "Heading number field code"
    Row "numsep", "Separator after the heading number"
    Row "numlevel", "Count restarts under heading level"
    Row "tlabel", "Table label"
    Row "flabel", "Figure label"

    mTop = mTop + 8
    Set lblFrom = AddLabel("lblFrom", "", X_LABEL, mTop + 4, 250, 28)
    Set btnFill = AddCtl("Forms.CommandButton.1", "btnFill", 268, mTop, 130, 24)
    btnFill.Caption = "Fill from the standards"
    Set btnSave = AddCtl("Forms.CommandButton.1", "btnSave", 404, mTop, 68, 24)
    btnSave.Caption = "Save"
    btnSave.Default = True
    Set btnCancel = AddCtl("Forms.CommandButton.1", "btnCancel", 476, mTop, 68, 24)
    btnCancel.Caption = "Cancel"
    btnCancel.Cancel = True
    Me.Height = mTop + 24 + 40
    LoadValues
    btnFill.Enabled = (SKW_QuickFromStandards().Count > 0)
    If btnFill.Enabled Then
        lblFrom.Caption = "The Global Standards workbook can fill the table styles and cell margins."
    Else
        lblFrom.Caption = "No Global Standards table looks are in reach, so these values are typed."
    End If
End Sub

Private Function Box(ByVal key As String) As Object
    On Error Resume Next
    Set Box = Me.Controls("txt_" & key)
End Function

Private Sub LoadValues()
    Dim k As Variant, b As Object
    For Each k In SKW_QuickKeys()
        Set b = Box(CStr(k))
        If Not b Is Nothing Then b.Text = SKW_QuickGet(CStr(k))
    Next
    cboHead.ListIndex = IIf(SKW_QuickGet("headrow") = "2", 1, 0)
    cboMode.ListIndex = IIf(LCase$(SKW_QuickGet("nummode")) = "sequential", 1, 0)
End Sub

Private Sub SaveValues()
    Dim k As Variant, b As Object
    For Each k In SKW_QuickKeys()
        Set b = Box(CStr(k))
        If Not b Is Nothing Then SKW_QuickSet CStr(k), Trim$(b.Text)
    Next
    SKW_QuickSet "headrow", IIf(cboHead.ListIndex = 1, "2", "auto")
    SKW_QuickSet "nummode", IIf(cboMode.ListIndex = 1, "sequential", "heading")
    SKW_QuickSet "saved", "1"
    Saved = True
End Sub

Private Function FillFromStandards() As Long
    Dim v As Object, k As Variant, b As Object
    Set v = SKW_QuickFromStandards()
    For Each k In v.Keys
        Set b = Box(CStr(k))
        If Not b Is Nothing Then b.Text = v(k): FillFromStandards = FillFromStandards + 1
    Next
End Function

Private Sub btnFill_Click()
    Dim n As Long
    n = FillFromStandards()
    lblFrom.Caption = n & " value(s) filled from the Global Standards. Change any of them, then Save."
End Sub

Private Sub btnSave_Click()
    Dim lv As String
    lv = Trim$(Box("numlevel").Text)
    If Len(lv) > 0 Then
        If Not IsNumeric(lv) Then MsgBox "The heading level is a number from 1 to 9.", vbExclamation, SKW_APP: Exit Sub
        If CLng(lv) < 1 Or CLng(lv) > 9 Then MsgBox "The heading level is a number from 1 to 9.", vbExclamation, SKW_APP: Exit Sub
    End If
    SaveValues
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Saved = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: Saved = False: Me.Hide
End Sub

' ---------------------------------------------------------------------
' Tests
' ---------------------------------------------------------------------
Public Function Probe() As String
    Dim k As Variant, b As Object, s As String, n As Long
    For Each k In SKW_QuickKeys()
        Set b = Box(CStr(k))
        If Not b Is Nothing Then n = n + 1: s = s & k & "=" & b.Text & " | "
    Next
    Probe = "boxes=" & n & " head=" & cboHead.ListIndex & " mode=" & cboMode.ListIndex & " fill=" & btnFill.Enabled & " ## " & s
End Function

Public Sub TestSet(ByVal key As String, ByVal v As String)
    If key = "headrow" Then
        cboHead.ListIndex = IIf(v = "2", 1, 0)
    ElseIf key = "nummode" Then
        cboMode.ListIndex = IIf(v = "sequential", 1, 0)
    Else
        Box(key).Text = v
    End If
End Sub

Public Sub TestSave()
    SaveValues
End Sub

Public Function TestFill() As Long
    TestFill = FillFromStandards()
End Function
