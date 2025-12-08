Option Explicit

Sub FormatMarkdown()
    Dim marks As Variant
    Dim rngSel As Range
    Dim para As Paragraph
    Dim txt As String
    Dim i As Long
    Dim mark As String
    Dim scope As String
    Dim name As String
    Dim markLen As Long
    Dim searchPos As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim innerStart As Long
    Dim innerEnd As Long
    Dim r As Range

    ' 1) mark, scope, name
    marks = Array( _
        Array("### ", "para", "heading3"), _
        Array("## ", "para", "heading2"), _
        Array("# ", "para", "heading1"), _
        Array("- ", "para", "bullet"), _
        Array("* ", "para", "bullet"), _
        Array("**", "char", "bold"), _
        Array("*", "char", "italic") _
    )

    ' 2) selected range
    Set rngSel = Selection.Range
    If Selection.Type = wdNoSelection Or rngSel.Paragraphs.Count = 0 Then Exit Sub

    ' 3) Loop 1: paragraphs
    For Each para In rngSel.Paragraphs

        ' prepare one reusable range per paragraph
        Set r = para.Range.Duplicate

        ' Loop 2: marks
        For i = LBound(marks) To UBound(marks)
            mark = marks(i)(0)
            scope = marks(i)(1)
            name = marks(i)(2)
            markLen = Len(mark)

            ' re-sync txt for this paragraph (in case of prior edits)
            txt = para.Range.Text
            txt = Left$(txt, Len(txt) - 1)

            If scope = "para" Then
                ' Paragraph-level: only check at start
                If Len(txt) >= markLen And Left$(txt, markLen) = mark Then

                    ' Apply paragraph formatting
                    Select Case name
                        Case "heading1": para.Style = "Heading 1"
                        Case "heading2": para.Style = "Heading 2"
                        Case "heading3": para.Style = "Heading 3"
                        Case "bullet":   para.Range.ListFormat.ApplyBulletDefault
                    End Select

                    ' Delete the mark at the start (length = markLen)
                    DeleteMark r, para, 1, markLen

                End If

            ElseIf scope = "char" Then
                ' Loop 3: character-level scanning for this mark
                searchPos = 1

                Do
                    startPos = InStr(searchPos, txt, mark)
                    If startPos = 0 Then Exit Do

                    endPos = InStr(startPos + markLen, txt, mark)
                    If endPos = 0 Then Exit Do

                    innerStart = startPos + markLen
                    innerEnd = endPos - 1

                    If innerStart <= innerEnd Then

                        ' Apply character formatting between the markers
                        r.SetRange _
                            para.Range.Characters(innerStart).Start, _
                            para.Range.Characters(innerEnd).End

                        Select Case name
                            Case "bold":   r.Font.Bold = True
                            Case "italic": r.Font.Italic = True
                        End Select

                        ' Delete end mark then start mark (each markLen long)
                        DeleteMark r, para, endPos, markLen
                        DeleteMark r, para, startPos, markLen

                        ' Refresh txt after edits
                        txt = para.Range.Text
                        txt = Left$(txt, Len(txt) - 1)

                        ' Continue searching this paragraph for this mark, after the formatted span
                        searchPos = endPos - markLen

                    Else
                        ' Empty span: just advance past the closing marker
                        searchPos = endPos + markLen
                    End If

                Loop
            End If
        Next i
    Next para
    
End Sub

Private Sub DeleteMark(ByVal base As Range, ByVal para As Paragraph, ByVal pos As Long, ByVal markLen As Long)

    base.SetRange _
        para.Range.Characters(pos).Start, _
        para.Range.Characters(pos + markLen - 1).End

    base.Delete
    
End Sub


