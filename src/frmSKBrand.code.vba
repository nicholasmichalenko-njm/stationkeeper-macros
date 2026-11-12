' =====================================================================
' The Stationkeeper look (1.6.0). This form is never shown: it holds the
' satellite pictures (imgLogo, imgIcon, put in when the kit is built, so no
' file is needed on the PC) and dresses the other forms:
'   Dress frm, title, subtitle, primaryButtons
'     a navy band with the satellite and the name, an amber line under it,
'     a light page, navy main buttons; the icon in the title bar.
'   SetIcon frm      the satellite in the title bar (call from Activate too)
' The kit's copy is frmSKBrand; the tools carry their own copy, frmSKWBrand,
' so either can be installed or removed alone.
' =====================================================================
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SendMessageA Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Private Const WM_SETICON As Long = &H80
Private Const BAND_H As Single = 50
' Stationkeeper colors (Windows writes them blue-green-red): navy #163460, amber #FFB81C,
' page #F5F8FC, subtitle #C9D4E3, ink #142030.
Private Const C_NAVY As Long = &H603416
Private Const C_AMBER As Long = &H1CB8FF
Private Const C_PAGE As Long = &HFCF8F5
Private Const C_SUBTLE As Long = &HE3D4C9
Private Const C_INK As Long = &H302014

Public Property Get Navy() As Long
    Navy = C_NAVY
End Property

Public Property Get Amber() As Long
    Amber = C_AMBER
End Property

Public Property Get PageColor() As Long
    PageColor = C_PAGE
End Property

' Dresses a form once all its controls are made. primary: names of the main buttons, "btnRun|btnSave".
Public Sub Dress(ByVal frm As Object, ByVal title As String, ByVal subtitle As String, Optional ByVal primary As String = "")
    Dim c As Object, oldBack As Long, band As Object, im As Object, lb As Object, w As Single
    On Error Resume Next
    If Not frm.Controls("lblSKBand") Is Nothing Then
        If Err.Number = 0 Then Exit Sub
    End If
    Err.Clear
    oldBack = frm.BackColor
    ' Room for the band: everything on the form itself moves down.
    For Each c In frm.Controls
        If c.Parent.Name = frm.Name Then c.Top = c.Top + BAND_H
    Next
    frm.Height = frm.Height + BAND_H
    w = frm.InsideWidth
    frm.BackColor = C_PAGE
    Set band = frm.Controls.Add("Forms.Label.1", "lblSKBand")
    band.Left = 0: band.Top = 0: band.Width = w: band.Height = BAND_H - 3
    band.BackStyle = 1: band.BackColor = C_NAVY: band.Caption = ""
    Set im = frm.Controls.Add("Forms.Image.1", "imgSKLogo")
    im.Left = 8: im.Top = 5: im.Width = 37.5: im.Height = 37.5
    im.BorderStyle = 0: im.BackStyle = 0: im.PictureSizeMode = 1
    Set im.Picture = Me.Controls("imgLogo").Picture
    Set lb = frm.Controls.Add("Forms.Label.1", "lblSKTitle")
    lb.Left = 54: lb.Top = 6: lb.Width = w - 64: lb.Height = 20
    lb.BackStyle = 0: lb.ForeColor = &HFFFFFF: lb.Caption = title
    lb.Font.Name = "Segoe UI": lb.Font.Size = 13: lb.Font.Bold = True
    Set lb = frm.Controls.Add("Forms.Label.1", "lblSKSubtitle")
    lb.Left = 54: lb.Top = 27: lb.Width = w - 64: lb.Height = 14
    lb.BackStyle = 0: lb.ForeColor = C_SUBTLE: lb.Caption = subtitle
    lb.Font.Name = "Segoe UI": lb.Font.Size = 8.5
    Set lb = frm.Controls.Add("Forms.Label.1", "lblSKLine")
    lb.Left = 0: lb.Top = BAND_H - 3: lb.Width = w: lb.Height = 3
    lb.BackStyle = 1: lb.BackColor = C_AMBER: lb.Caption = ""
    ' The rest of the form: the page color, navy main buttons.
    For Each c In frm.Controls
        If Left$(c.Name, 5) <> "lblSK" And c.Name <> "imgSKLogo" Then
            Select Case TypeName(c)
                Case "Label", "CheckBox", "OptionButton"
                    If c.BackStyle = 1 And (c.BackColor = oldBack Or c.BackColor = &HFFFFFF Or c.BackColor = &H8000000F) Then c.BackStyle = 0
                    If c.ForeColor = &H80000012 Or c.ForeColor = 0 Then c.ForeColor = C_INK
                Case "CommandButton"
                    If InStr(1, "|" & primary & "|", "|" & c.Name & "|", vbTextCompare) > 0 Then
                        c.BackColor = C_NAVY: c.ForeColor = &HFFFFFF: c.Font.Bold = True
                    Else
                        c.BackColor = &HFFFFFF: c.ForeColor = C_INK
                    End If
                Case "MultiPage"
                    c.BackColor = C_PAGE
                Case "Frame"
                    If c.BackColor = oldBack Or c.BackColor = &H8000000F Or c.BackColor = &HFFFFFF Then c.BackColor = C_PAGE
            End Select
        End If
    Next
    SetIcon frm
End Sub

' The satellite in the form's title bar. The window has to exist: call it from Activate as well.
Public Sub SetIcon(ByVal frm As Object)
    Dim h As LongPtr, hIcon As LongPtr
    On Error Resume Next
    hIcon = Me.Controls("imgIcon").Picture.Handle
    If hIcon = 0 Then Exit Sub
    h = FindWindowA("ThunderDFrame", frm.Caption)
    If h = 0 Then h = FindWindowA("ThunderXFrame", frm.Caption)
    If h = 0 Then Exit Sub
    SendMessageA h, WM_SETICON, 0, hIcon
    SendMessageA h, WM_SETICON, 1, hIcon
End Sub

' For the tests: what the dressing put on a form.
Public Function Probe(ByVal frm As Object) As String
    Dim s As String
    On Error Resume Next
    s = "band=" & Hex$(frm.Controls("lblSKBand").BackColor)
    s = s & " line=" & Hex$(frm.Controls("lblSKLine").BackColor)
    s = s & " title=" & frm.Controls("lblSKTitle").Caption
    s = s & " logo=" & (frm.Controls("imgSKLogo").Picture.Handle <> 0)
    s = s & " icon=" & Me.Controls("imgIcon").Picture.Type
    s = s & " page=" & Hex$(frm.BackColor)
    Probe = s
End Function
