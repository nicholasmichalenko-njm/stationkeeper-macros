' =====================================================================
' Stationkeeper Tools: the progress window.
' Shown by SKW_Lib while a tool runs longer than a moment: the tool's
' name, what it is on, a bar with the percent done, the time it has run
' and the time left. The controls are built in code, so this form needs
' no designer work. It has no buttons: it closes when the tool ends.
' =====================================================================
Option Explicit

Private lblTitle As MSForms.Label
Private lblLine As MSForms.Label
Private lblTrack As MSForms.Label
Private lblFill As MSForms.Label
Private lblTime As MSForms.Label

Private Const C_NAVY As Long = &H603416
Private Const C_AMBER As Long = &H1CB8FF
Private Const C_PAGE As Long = &HFCF8F5
Private Const C_INK As Long = &H2E2017
Private Const TRACK_W As Single = 318

Private Function AddLabel(ByVal nm As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As MSForms.Label
    Dim c As MSForms.Label
    Set c = Me.Controls.Add("Forms.Label.1", nm, True)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Caption = ""
    On Error Resume Next
    c.Font.Name = "Segoe UI"
    c.Font.Size = 9
    Set AddLabel = c
End Function

Private Sub UserForm_Initialize()
    Me.Caption = "Stationkeeper Tools"
    Me.Width = 354
    Me.Height = 122
    Me.BackColor = C_PAGE
    Set lblTitle = AddLabel("lblTitle", 12, 8, TRACK_W, 16)
    lblTitle.Font.Size = 11: lblTitle.Font.Bold = True: lblTitle.ForeColor = C_NAVY: lblTitle.BackStyle = 0
    Set lblLine = AddLabel("lblLine", 12, 28, TRACK_W, 14)
    lblLine.ForeColor = C_INK: lblLine.BackStyle = 0
    Set lblTrack = AddLabel("lblTrack", 12, 48, TRACK_W, 10)
    lblTrack.BackStyle = 1: lblTrack.BackColor = &HE6DCD2
    Set lblFill = AddLabel("lblFill", 12, 48, 0, 10)
    lblFill.BackStyle = 1: lblFill.BackColor = C_NAVY
    Set lblTime = AddLabel("lblTime", 12, 64, TRACK_W, 14)
    lblTime.ForeColor = C_INK: lblTime.BackStyle = 0
End Sub

' pct below zero: the tool cannot say how far along it is, so the bar shows amber, full width.
Public Sub SetState(ByVal title As String, ByVal line As String, ByVal pct As Double, ByVal timeText As String)
    On Error Resume Next
    lblTitle.Caption = title
    lblLine.Caption = line
    If pct < 0 Then
        lblFill.BackColor = C_AMBER
        lblFill.Width = TRACK_W
    Else
        If pct > 100 Then pct = 100
        lblFill.BackColor = C_NAVY
        lblFill.Width = TRACK_W * pct / 100
    End If
    lblTime.Caption = timeText
    Me.Repaint
End Sub

' Tests read what the window shows.
Public Function Probe() As String
    Probe = lblTitle.Caption & " | " & lblLine.Caption & " | fill=" & Format$(lblFill.Width / TRACK_W * 100, "0") & " | " & lblTime.Caption
End Function

' Closing it by hand hides it and does not stop the tool. It shows again with the next tool.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: Me.Hide
End Sub
