Attribute VB_Name = "SK_Install"
' =====================================================================
' Stationkeeper Macros kit installer 0.3.0
' Lives only in "Stationkeeper Macros Kit.dotm". It copies the menu and
' the admin into a template you pick, with Word's Organizer. No VBA
' editor needed. Your own macros are not touched.
'
'   StationkeeperMacrosInstall   add (or update) the kit in a template
'   StationkeeperMacrosRemove    take the kit back out of a template
' =====================================================================
Option Explicit

Private Const KIT_ITEMS As String = "SK_Macros|frmSKMenu|frmSKAdmin"

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
    Dim it As Variant, d As Object, bak As String, wasDirty As Boolean

    If LCase$(target) = LCase$(ThisDocument.FullName) Then
        SK_InstallInto = "Pick a template other than the kit itself."
        Exit Function
    End If
    If Len(Dir$(target)) = 0 Then
        SK_InstallInto = "Not found: " & target
        Exit Function
    End If
    Set d = SK_OpenDoc(target)
    If Not d Is Nothing Then
        wasDirty = Not d.Saved
        If wasDirty Then
            SK_InstallInto = "Save or close " & d.Name & " first. It has unsaved changes."
            Exit Function
        End If
    End If

    bak = SK_Backup(target, "before kit", backupDir)

    On Error GoTo Fail
    For Each it In Split(KIT_ITEMS, "|")
        On Error Resume Next
        Application.OrganizerDelete Source:=target, Name:=it, Object:=wdOrganizerObjectProjectItems
        Err.Clear
        On Error GoTo Fail
        Application.OrganizerCopy Source:=ThisDocument.FullName, Destination:=target, Name:=it, _
                                  Object:=wdOrganizerObjectProjectItems
    Next
    If Not d Is Nothing Then d.Save

    SK_InstallInto = "Stationkeeper Macros is in " & Mid$(target, InStrRev(target, "\") + 1) & "." & vbCr & vbCr & _
        "Next: open that template, run StationkeeperMacrosAdmin (View > Macros) to describe its macros, " & _
        "then StationkeeperMacros to see the menu." & _
        IIf(Len(bak) > 0, vbCr & vbCr & "Backup of the template before the change:" & vbCr & bak, "")
    Exit Function
Fail:
    SK_InstallInto = "The kit could not be copied: " & Err.Description & vbCr & _
        IIf(Len(bak) > 0, "Your template is unchanged or can be restored from:" & vbCr & bak, "")
End Function

Public Function SK_RemoveFrom(ByVal target As String, Optional ByVal backupDir As String = "") As String
    Dim it As Variant, d As Object, bak As String, n As Long
    Set d = SK_OpenDoc(target)
    If Not d Is Nothing Then
        If Not d.Saved Then
            SK_RemoveFrom = "Save or close " & d.Name & " first. It has unsaved changes."
            Exit Function
        End If
    End If
    bak = SK_Backup(target, "before removing kit", backupDir)
    For Each it In Split(KIT_ITEMS, "|")
        On Error Resume Next
        Application.OrganizerDelete Source:=target, Name:=it, Object:=wdOrganizerObjectProjectItems
        If Err.Number = 0 Then n = n + 1
        Err.Clear
        On Error GoTo 0
    Next
    If Not d Is Nothing Then d.Save
    SK_RemoveFrom = "Removed " & n & " kit item(s) from " & Mid$(target, InStrRev(target, "\") + 1) & "." & _
        IIf(Len(bak) > 0, vbCr & "Backup: " & bak, "")
End Function
