Attribute VB_Name = "SKR_Ribbon"
' =====================================================================
' Stationkeeper Ribbon: the Stationkeeper tab on Word's ribbon (since 1.6.0).
' A small add-in of its own ("Stationkeeper Ribbon.dotm" in Word's Startup
' folder). Its buttons open the menu and the admin of whichever template has
' Stationkeeper Macros. It holds no menu of its own and changes no document.
' Each template carries its own copy of the menu, at the version last
' installed into it. When the copy that opens is older than the release this
' tab came with, the tab says so once, names the file, and says how to update it.
' Turn it off: File > Options > Add-ins > Manage: Templates > Go, untick it.
' Remove it: run StationkeeperRibbonRemove from the kit, or close Word and
' delete the file from the Startup folder.
' =====================================================================
Option Explicit

Private Const APP As String = "Stationkeeper"
Private Const VER As String = "1.17.0"
Private mTold As String

Public Sub SKR_Menu(ByVal control As Object)
    RunIt "StationkeeperMacros", "the macro menu"
End Sub

Public Sub SKR_Admin(ByVal control As Object)
    RunIt "StationkeeperMacrosAdmin", "the admin"
End Sub

Public Sub SKR_About(ByVal control As Object)
    MsgBox "Stationkeeper Ribbon " & VER & vbCr & vbCr & _
        "This tab opens Stationkeeper Macros in the template your document is based on. It is a small add-in, loaded from:" & vbCr & _
        ThisDocument.FullName & vbCr & vbCr & _
        "Turn it off for now: File > Options > Add-ins > Manage: Templates > Go, and untick Stationkeeper Ribbon.dotm." & vbCr & _
        "Take it off for good: run StationkeeperRibbonRemove from the kit, or close Word and delete the file above." & vbCr & _
        "Without the tab, the menu opens from View > Macros > StationkeeperMacros.", vbInformation, APP
End Sub

' For the tests and the kit: the add-in is loaded and answers.
Public Function SKR_Probe() As String
    SKR_Probe = "Stationkeeper Ribbon " & VER
End Function

Private Function VerNum(ByVal v As String) As Double
    Dim a() As String, i As Long, n As Double
    On Error Resume Next
    a = Split(Trim$(v) & ".0.0", ".")
    For i = 0 To 2
        n = n * 1000 + Val(a(i))
    Next
    VerNum = n
End Function

' What to say when the menu that opens is version ver, in the file where. "" when it is this release or newer.
Public Function SKR_BehindFor(ByVal menuVer As String, ByVal where As String) As String
    If Len(menuVer) > 0 And VerNum(menuVer) >= VerNum(VER) Then Exit Function
    SKR_BehindFor = "The Stationkeeper menu that opens here is " & IIf(Len(menuVer) > 0, "version " & menuVer, "an older version") & ", in:" & vbCr & where & vbCr & vbCr & _
        "This tab came with version " & VER & ". Each template keeps its own copy of the menu, at the version last installed into it." & vbCr & vbCr & _
        "To bring that template up to date: open Stationkeeper Macros Kit.dotm from the " & VER & " release, run StationkeeperMacrosInstall, " & _
        "then StationkeeperToolsInstall, and pick that template each time. Your descriptions, categories and settings are kept."
End Function

' Asks the copy of Stationkeeper Macros that Word would run here for its version and file.
Public Function SKR_Behind() As String
    Dim a As String, ver As String, where As String
    On Error Resume Next
    a = Application.Run("SK_About")
    If InStr(a, "|") > 0 Then
        ver = Split(a, "|")(0)
        where = Split(a, "|")(1)
    Else
        ' A copy from before 1.15.0 cannot say where it is.
        where = "the template this document is based on (" & ActiveDocument.AttachedTemplate.Name & "), or another template Word has loaded"
        If Documents.Count = 0 Then where = "a template Word has loaded"
    End If
    SKR_Behind = SKR_BehindFor(ver, where)
End Function

Private Sub RunIt(ByVal macro As String, ByVal what As String)
    Dim msg As String
    On Error GoTo NotThere
    Application.Run macro
    On Error Resume Next
    msg = SKR_Behind()
    If Len(msg) > 0 And InStr(mTold, "|" & msg & "|") = 0 Then
        mTold = mTold & "|" & msg & "|"
        MsgBox msg, vbInformation, APP
    End If
    Exit Sub
NotThere:
    MsgBox "Could not open " & what & ": " & Err.Description & vbCr & vbCr & _
        "Open a document based on your team's template (the one Stationkeeper Macros is installed in), then try again.", vbInformation, APP
End Sub
