Attribute VB_Name = "SKW_Lists"
' =====================================================================
' Stationkeeper Tools: lists.
'   TypedListsFinder   report: bullets and numbers typed by hand instead of real lists
' =====================================================================
Option Explicit

' Lists paragraphs that start with a typed bullet or number (a dash, a bullet character, "1.", "a)", "(iv)") instead of real Word list formatting, which does not renumber or line up. Changes nothing.
Public Sub TypedListsFinder()
    SKW_Show SKW_TypedListsFinder()
End Sub

' What a paragraph starts with, if it looks like a typed list marker; "" otherwise.
Public Function SKW_TypedMarker(ByVal t As String) As String
    Dim first As String, k As Long, head As String, sep As String
    t = Replace(t, vbTab, " ")
    If Len(t) < 3 Then Exit Function
    first = Left$(t, 1)
    Select Case AscW(first)
        Case &H2022, &H25CF, &H25AA, &H25A0, &H2013, &H2014, &HB7, &H25E6, &H27A2, &H2192, &HF0B7, &HF0A7
            If Mid$(t, 2, 1) = " " Then SKW_TypedMarker = "bullet " & first
            Exit Function
    End Select
    If (first = "-" Or first = "*" Or first = ">") And Mid$(t, 2, 1) = " " Then SKW_TypedMarker = "bullet " & first: Exit Function
    If first = "o" And Mid$(t, 2, 1) = " " And Mid$(t, 3, 1) Like "[A-Z]" Then SKW_TypedMarker = "bullet o": Exit Function
    k = InStr(t, " ")
    If k < 2 Or k > 7 Then Exit Function
    head = Left$(t, k - 1)
    If head Like "(*)" Then head = Mid$(head, 2, Len(head) - 2): sep = ")"
    If Len(sep) = 0 Then
        sep = Right$(head, 1)
        If sep <> "." And sep <> ")" Then Exit Function
        head = Left$(head, Len(head) - 1)
        If Len(head) = 0 Then Exit Function
    End If
    If head Like "#" Or head Like "##" Or head Like "#.#" Or head Like "#.##" Or head Like "##.#" Then
        SKW_TypedMarker = "number " & Left$(t, k - 1)
    ElseIf head Like "[a-z]" Or head Like "[A-Z]" Then
        SKW_TypedMarker = "letter " & Left$(t, k - 1)
    ElseIf LCase$(head) Like "[ivx]" Or LCase$(head) Like "[ivx][ivx]" Or LCase$(head) Like "[ivx][ivx][ivx]" Or LCase$(head) Like "[ivx][ivx][ivx][ivx]" Then
        SKW_TypedMarker = "roman " & Left$(t, k - 1)
    End If
End Function

Public Function SKW_TypedListsFinder() As String
    Dim d As Object, msg As String, rows As New Collection, story As Object, p As Object, m As String, n As Long
    Set SKW_LastReport = Nothing
    msg = SKW_Begin("Typed lists finder", False)
    If Len(msg) > 0 Then SKW_TypedListsFinder = msg: Exit Function
    Set d = ActiveDocument
    For Each story In SKW_Stories(d)
        If story.StoryType = wdMainTextStory Or story.StoryType = wdFootnotesStory Or story.StoryType = wdEndnotesStory Or story.StoryType = wdTextFrameStory Then
            SKW_Steps story.Paragraphs.Count, "Paragraph"
            For Each p In story.Paragraphs
                SKW_Step
                If p.Range.ListFormat.ListType = wdListNoNumbering Then
                    m = SKW_TypedMarker(p.Range.Text)
                    If Len(m) > 0 Then
                        n = n + 1
                        If n <= 1000 Then rows.Add SKW_Row("Typed " & Split(m, " ")(0), "Starts with a typed " & m & ". Use a list style (or Home > Bullets/Numbering)", _
                            SKW_StoryName(story.StoryType), SKW_Page(p.Range), p.Range.Text)
                    End If
                End If
            Next
        End If
    Next
    SKW_Finish
    SKW_Report "Typed Lists Finder", n & " paragraph(s) start with a typed bullet or number. Real Word lists are not listed. Changes nothing.", _
        Array("Rule", "Finding", "Where", "Page", "Text"), rows
    SKW_TypedListsFinder = n & " typed list paragraph(s)."
End Function
