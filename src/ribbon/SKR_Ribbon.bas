Attribute VB_Name = "SKR_Ribbon"
' =====================================================================
' Stationkeeper Ribbon 1.6.0: the Stationkeeper tab on Word's ribbon.
' A small add-in of its own ("Stationkeeper Ribbon.dotm" in Word's Startup
' folder). Its buttons open the menu and the admin of whichever template has
' Stationkeeper Macros. It holds no other code and changes no document.
' Turn it off: File > Options > Add-ins > Manage: Templates > Go, untick it.
' Remove it: run StationkeeperRibbonRemove from the kit, or close Word and
' delete the file from the Startup folder.
' =====================================================================
Option Explicit

Private Const APP As String = "Stationkeeper"
Private Const VER As String = "1.6.0"

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

Private Sub RunIt(ByVal macro As String, ByVal what As String)
    On Error GoTo NotThere
    Application.Run macro
    Exit Sub
NotThere:
    MsgBox "Could not open " & what & ": " & Err.Description & vbCr & vbCr & _
        "Open a document based on your team's template (the one Stationkeeper Macros is installed in), then try again.", vbInformation, APP
End Sub
