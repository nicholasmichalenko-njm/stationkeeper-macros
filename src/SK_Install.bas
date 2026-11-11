Attribute VB_Name = "SK_Install"
' =====================================================================
' Stationkeeper Macros kit installer 1.1.0
' Lives only in "Stationkeeper Macros Kit.dotm". It copies the menu and
' the admin into a template you pick, with Word's Organizer. No VBA
' editor needed. Your own macros are not touched.
'
'   StationkeeperMacrosInstall   add (or update) the kit in a template
'   StationkeeperMacrosRemove    take the kit back out of a template
'   StationkeeperToolsInstall    add (or update) Stationkeeper Tools in a template
'   StationkeeperToolsRemove     take Stationkeeper Tools back out
' The tools work on their own; with the menu they arrive described.
' =====================================================================
Option Explicit

Private Const KIT_ITEMS As String = "SK_Macros|frmSKMenu|frmSKAdmin"
Private Const TOOL_ITEMS As String = "SKW_Lib|SKW_Setup|SKW_Fields|SKW_Review|SKW_Styles|SKW_Tables|SKW_Layout|SKW_Captions|SKW_Lists|SKW_Acronyms|SKW_Text|SKW_Output|SKW_Utils|SKP_Deck|SKX_Book"

Public Sub StationkeeperMacrosInstall()
    Dim p As String
    Dim d As String
    p = SK_PickTemplate("Pick the template to add Stationkeeper Macros to (try it on a copy first)")
    If Len(p) = 0 Then Exit Sub
    d = SK_AskBackupDir(Left$(p, InStrRev(p, "\") - 1))
    If d = "?" Then Exit Sub
    MsgBox SK_InstallInto(p, d), vbInformation, "Stationkeeper Macros"
End Sub

Public Sub StationkeeperMacrosRemove()
    Dim p As String
    p = SK_PickTemplate("Pick the template to take Stationkeeper Macros out of")
    If Len(p) = 0 Then Exit Sub
    If MsgBox("Take the menu and admin out of " & Mid$(p, InStrRev(p, "\") + 1) & "?" & vbCr & _
              "Your macros and the catalog of descriptions stay in the template.", _
              vbYesNo + vbQuestion, "Stationkeeper Macros") <> vbYes Then Exit Sub
    Dim d As String
    d = SK_AskBackupDir(Left$(p, InStrRev(p, "\") - 1))
    If d = "?" Then Exit Sub
    MsgBox SK_RemoveFrom(p, d), vbInformation, "Stationkeeper Macros"
End Sub

Public Sub StationkeeperToolsInstall()
    Dim p As String, d As String
    p = SK_PickTemplate("Pick the template to add Stationkeeper Tools to (try it on a copy first)")
    If Len(p) = 0 Then Exit Sub
    d = SK_AskBackupDir(Left$(p, InStrRev(p, "\") - 1))
    If d = "?" Then Exit Sub
    MsgBox SK_InstallToolsInto(p, d), vbInformation, "Stationkeeper Tools"
End Sub

Public Sub StationkeeperToolsRemove()
    Dim p As String, d As String
    p = SK_PickTemplate("Pick the template to take Stationkeeper Tools out of")
    If Len(p) = 0 Then Exit Sub
    If MsgBox("Take Stationkeeper Tools out of " & Mid$(p, InStrRev(p, "\") + 1) & "?" & vbCr & _
              "Your own macros, the menu and the catalog stay.", vbYesNo + vbQuestion, "Stationkeeper Tools") <> vbYes Then Exit Sub
    d = SK_AskBackupDir(Left$(p, InStrRev(p, "\") - 1))
    If d = "?" Then Exit Sub
    MsgBox SK_RemoveToolsFrom(p, d), vbInformation, "Stationkeeper Tools"
End Sub

Private Function SK_PickTemplate(ByVal title As String) As String
    With Application.FileDialog(3)
        .Title = title
        .Filters.Clear
        .Filters.Add "Macro-enabled templates and documents", "*.dotm;*.docm"
        .AllowMultiSelect = False
        If .Show = -1 Then SK_PickTemplate = .SelectedItems(1)
    End With
End Function

' Where backups go. Uses the folder chosen before; otherwise asks.
' Returns the folder, "-" for no backup, or "?" if the person cancelled.
Private Function SK_AskBackupDir(ByVal besideFolder As String) As String
    Dim d As String, sug As String, r As VbMsgBoxResult
    d = SK_BackupDir()
    If Len(d) > 0 Then SK_AskBackupDir = d: Exit Function
    sug = besideFolder & "\Stationkeeper Macros backups"
    r = MsgBox("Before changing a template, Stationkeeper Macros saves a dated backup of it." & vbCr & vbCr & _
               "Where should backups go?" & vbCr & vbCr & "Yes: " & sug & vbCr & "No: choose another folder", _
               vbYesNoCancel + vbQuestion, "Stationkeeper Macros")
    If r = vbYes Then
        d = sug
    ElseIf r = vbNo Then
        d = SK_PickFolder("Choose where Stationkeeper Macros keeps backups", besideFolder)
    Else
        SK_AskBackupDir = "?"
        Exit Function
    End If
    If Len(d) = 0 Then
        If MsgBox("Go ahead without a backup?", vbYesNo + vbQuestion, "Stationkeeper Macros") = vbYes Then
            SK_AskBackupDir = "-"
        Else
            SK_AskBackupDir = "?"
        End If
        Exit Function
    End If
    SK_SetBackupDir d
    SK_AskBackupDir = d
End Function

Private Function SK_OpenDoc(ByVal path As String) As Object
    Dim d As Object
    For Each d In Documents
        If LCase$(d.FullName) = LCase$(path) Then Set SK_OpenDoc = d: Exit Function
    Next
End Function

' bdir: a folder, "-" for none, or "" to use the chosen folder (or beside the template).
Private Function SK_Backup(ByVal path As String, ByVal tag As String, ByVal bdir As String) As String
    Dim folder As String, base As String, ext As String, dst As String
    If bdir = "-" Then Exit Function
    On Error GoTo Fail
    folder = bdir
    If Len(folder) = 0 Then folder = SK_BackupDir()
    If Len(folder) = 0 Then folder = Left$(path, InStrRev(path, "\") - 1) & "\Stationkeeper Macros backups"
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    base = Mid$(path, InStrRev(path, "\") + 1)
    ext = Mid$(base, InStrRev(base, "."))
    base = Left$(base, InStrRev(base, ".") - 1)
    dst = folder & "\" & base & " - " & tag & " " & Format$(Now, "yyyy-mm-dd hhnnss") & ext
    CreateObject("Scripting.FileSystemObject").CopyFile path, dst, True
    SK_Backup = dst
    Exit Function
Fail:
    SK_Backup = ""
End Function

' Copies the three kit items into the target. Replaces older copies of them.
Public Function SK_InstallInto(ByVal target As String, Optional ByVal backupDir As String = "") As String
    Dim r As String, bak As String
    r = SK_CopyItems(target, KIT_ITEMS, "before kit", backupDir, bak)
    If Len(r) > 0 Then SK_InstallInto = r: Exit Function
    SK_InstallInto = "Stationkeeper Macros is in " & Mid$(target, InStrRev(target, "\") + 1) & "." & vbCr & vbCr & _
        "Next: open that template, run StationkeeperMacrosAdmin (View > Macros) to describe its macros, " & _
        "then StationkeeperMacros to see the menu." & vbCr & _
        "To add the checking tools too, run StationkeeperToolsInstall." & _
        IIf(Len(bak) > 0, vbCr & vbCr & "Backup of the template before the change:" & vbCr & bak, "")
End Function

' Copies the Stationkeeper Tools modules into the target. Replaces older copies of them.
Public Function SK_InstallToolsInto(ByVal target As String, Optional ByVal backupDir As String = "") As String
    Dim r As String, bak As String
    r = SK_CopyItems(target, TOOL_ITEMS, "before tools", backupDir, bak)
    If Len(r) > 0 Then SK_InstallToolsInto = r: Exit Function
    SK_InstallToolsInto = "Stationkeeper Tools are in " & Mid$(target, InStrRev(target, "\") + 1) & " (" & _
        (UBound(Split(TOOL_ITEMS, "|")) + 1) & " modules)." & vbCr & vbCr & _
        "Open that template: with Stationkeeper Macros installed, the tools appear in the menu already described " & _
        "(save once in the admin to keep them). Without it, run them from View > Macros." & vbCr & _
        "Then run MakeStandardsWorkbook or ChooseStandardsWorkbook to set your team's rules." & _
        IIf(Len(bak) > 0, vbCr & vbCr & "Backup of the template before the change:" & vbCr & bak, "")
End Function

Public Function SK_RemoveToolsFrom(ByVal target As String, Optional ByVal backupDir As String = "") As String
    SK_RemoveToolsFrom = Replace(SK_DeleteItems(target, TOOL_ITEMS, "before removing tools", backupDir), "kit item(s)", "tool module(s)")
End Function

Private Function SK_CopyItems(ByVal target As String, ByVal items As String, ByVal tag As String, _
                              ByVal backupDir As String, ByRef bak As String) As String
    Dim it As Variant, d As Object

    If LCase$(target) = LCase$(ThisDocument.FullName) Then
        SK_CopyItems = "Pick a template other than the kit itself."
        Exit Function
    End If
    If Len(Dir$(target)) = 0 Then
        SK_CopyItems = "Not found: " & target
        Exit Function
    End If
    Set d = SK_OpenDoc(target)
    If Not d Is Nothing Then
        If Not d.Saved Then
            SK_CopyItems = "Save or close " & d.Name & " first. It has unsaved changes."
            Exit Function
        End If
    End If

    bak = SK_Backup(target, tag, backupDir)

    On Error GoTo Fail
    For Each it In Split(items, "|")
        On Error Resume Next
        Application.OrganizerDelete Source:=target, Name:=it, Object:=wdOrganizerObjectProjectItems
        Err.Clear
        On Error GoTo Fail
        Application.OrganizerCopy Source:=ThisDocument.FullName, Destination:=target, Name:=it, _
                                  Object:=wdOrganizerObjectProjectItems
    Next
    If Not d Is Nothing Then d.Save
    Exit Function
Fail:
    SK_CopyItems = "Could not copy " & it & ": " & Err.Description & vbCr & _
        IIf(Len(bak) > 0, "Your template is unchanged or can be restored from:" & vbCr & bak, "")
End Function

Public Function SK_RemoveFrom(ByVal target As String, Optional ByVal backupDir As String = "") As String
    SK_RemoveFrom = SK_DeleteItems(target, KIT_ITEMS, "before removing kit", backupDir)
End Function

Private Function SK_DeleteItems(ByVal target As String, ByVal items As String, ByVal tag As String, _
                                ByVal backupDir As String) As String
    Dim it As Variant, d As Object, bak As String, n As Long
    Set d = SK_OpenDoc(target)
    If Not d Is Nothing Then
        If Not d.Saved Then
            SK_DeleteItems = "Save or close " & d.Name & " first. It has unsaved changes."
            Exit Function
        End If
    End If
    bak = SK_Backup(target, tag, backupDir)
    For Each it In Split(items, "|")
        On Error Resume Next
        Application.OrganizerDelete Source:=target, Name:=it, Object:=wdOrganizerObjectProjectItems
        If Err.Number = 0 Then n = n + 1
        Err.Clear
        On Error GoTo 0
    Next
    If Not d Is Nothing Then d.Save
    SK_DeleteItems = "Removed " & n & " kit item(s) from " & Mid$(target, InStrRev(target, "\") + 1) & "." & _
        IIf(Len(bak) > 0, vbCr & "Backup: " & bak, "")
End Function
