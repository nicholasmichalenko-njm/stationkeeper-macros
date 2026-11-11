Attribute VB_Name = "SKW_Output"
' =====================================================================
' Stationkeeper Tools: files to send.
'   MetadataCleanCopy   saves a copy without author, company, template path and other file details
'   PDFWithBookmarks    saves a PDF with bookmarks from the headings and tags for screen readers
' Both leave the open document as it is.
' =====================================================================
Option Explicit

' Saves a copy beside the document ("<name> - clean.docx") with the author, company, last-saved-by, template path, document variables and other personal details taken out. The open document is not changed. Comments and tracked changes stay: run Pre-Delivery Leftovers first.
Public Sub MetadataCleanCopy()
    SKW_Say SKW_MetadataCleanCopy("")
End Sub

' Saves the document as a PDF beside it, with bookmarks from the headings and tags for screen readers, and without the file's personal details. The open document is not changed.
Public Sub PDFWithBookmarks()
    Dim p As String
    p = SKW_DefaultOut(".pdf")
    If Len(p) = 0 Then SKW_Say "Save the document first, so the PDF has a folder to go in.": Exit Sub
    If Not SKW_IsQuiet() Then
        With Application.FileDialog(2)
            .Title = "Save the PDF as"
            .InitialFileName = p
            If .Show <> -1 Then Exit Sub
            p = .SelectedItems(1)
        End With
        If LCase$(Right$(p, 4)) <> ".pdf" Then p = SKW_Stem(p) & ".pdf"
    End If
    SKW_Say SKW_PDFWithBookmarks(p)
End Sub

' "<folder>\<name><suffix>" beside the active document, or "" if it was never saved.
Public Function SKW_DefaultOut(ByVal suffix As String) As String
    Dim d As Object, base As String
    If Documents.Count = 0 Then Exit Function
    Set d = ActiveDocument
    If Len(d.Path) = 0 Or LCase$(Left$(d.FullName, 4)) = "http" Then Exit Function
    base = d.Name
    If InStrRev(base, ".") > 0 Then base = Left$(base, InStrRev(base, ".") - 1)
    SKW_DefaultOut = d.Path & "\" & base & suffix
End Function

Public Function SKW_MetadataCleanCopy(ByVal target As String) As String
    Dim d As Object, c As Object, tmp As String, v As Variant, n As Long, warn As String, i As Long
    If Documents.Count = 0 Then SKW_MetadataCleanCopy = "Open a document first.": Exit Function
    Set d = ActiveDocument
    If Len(target) = 0 Then target = SKW_DefaultOut(" - clean.docx")
    If Len(target) = 0 Then SKW_MetadataCleanCopy = "Save the document first, so the copy has a folder to go in.": Exit Function
    On Error GoTo Fail
    SKW_Status "Making a clean copy..."
    ' A new document from the saved file, so the open one is never touched.
    If Not d.Saved Then warn = "The document had unsaved changes; the copy is made from the last saved version." & vbCr
    Set c = Documents.Add(Template:=d.FullName, Visible:=False)
    For Each v In Array(4, 8, 9, 14, 15, 16)   ' personal info, properties, template, server properties, policy, content type
        On Error Resume Next
        c.RemoveDocumentInformation v
        On Error GoTo Fail
    Next
    c.RemovePersonalInformation = True
    For i = c.Variables.Count To 1 Step -1
        c.Variables(i).Delete
    Next
    On Error Resume Next
    c.AttachedTemplate = ""
    On Error GoTo Fail
    If c.Comments.Count > 0 Then warn = warn & c.Comments.Count & " comment(s) are still in the copy." & vbCr
    If c.Revisions.Count > 0 Then warn = warn & c.Revisions.Count & " tracked change(s) are still in the copy." & vbCr
    c.SaveAs2 FileName:=target, FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
    c.Close SaveChanges:=0
    d.Activate
    SKW_Status ""
    SKW_MetadataCleanCopy = "Saved a clean copy: " & target & vbCr & _
        "Taken out: author, last saved by, company, manager and other properties, the template path, document variables and personal details." & vbCr & _
        IIf(Len(warn) > 0, vbCr & warn & "Run Pre-Delivery Leftovers to see them.", "No comments or tracked changes were in it.")
    Exit Function
Fail:
    SKW_MetadataCleanCopy = "The clean copy could not be made: " & Err.Description
    On Error Resume Next
    If Not c Is Nothing Then c.Close SaveChanges:=0
    d.Activate
End Function

Public Function SKW_PDFWithBookmarks(ByVal target As String) As String
    Dim d As Object
    If Documents.Count = 0 Then SKW_PDFWithBookmarks = "Open a document first.": Exit Function
    Set d = ActiveDocument
    If Len(target) = 0 Then target = SKW_DefaultOut(".pdf")
    If Len(target) = 0 Then SKW_PDFWithBookmarks = "Save the document first, so the PDF has a folder to go in.": Exit Function
    On Error GoTo Fail
    SKW_Status "Saving the PDF..."
    d.ExportAsFixedFormat OutputFileName:=target, ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, Item:=wdExportDocumentContent, _
        IncludeDocProps:=False, KeepIRM:=True, CreateBookmarks:=wdExportCreateHeadingBookmarks, _
        DocStructureTags:=True, BitmapMissingFonts:=True, UseISO19005_1:=False
    SKW_Status ""
    SKW_PDFWithBookmarks = "Saved the PDF: " & target & vbCr & "Bookmarks come from the headings; tagged for screen readers; file details left out."
    Exit Function
Fail:
    SKW_PDFWithBookmarks = "The PDF could not be saved: " & Err.Description
End Function
