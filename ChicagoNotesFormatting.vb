Sub ChicagoNotesFormatting()

    ' If endnotes found in active document, adjust formatting of endnotes and footnotes
    If ActiveDocument.Endnotes.Count >= 1 Then

        ' Locate endnotes at end of section and format endnotes marks using Arabic numerals
        With Selection.EndnoteOptions
            .Location = wdEndOfSection
            .NumberingRule = wdRestartContinuous
            .StartingNumber = 1
            .NumberStyle = wdNoteNumberStyleArabic
        End With

        ' In the endnotes range, endnote marks are followed by a period and are not superscript
        Set NotesRange = ActiveDocument.StoryRanges(wdEndnotesStory)
        Call MarkFormatting(NotesRange, "\1. ")


        ' Footnote count restarts with each page; footnote marks using symbols instead of Arabic numerals
        If ActiveDocument.Footnotes.Count >= 1 Then

            With Selection.FootnoteOptions
                .Location = wdBottomOfPage
                .StartingNumber = 1
                .NumberStyle = wdNoteNumberStyleSymbol
                .NumberingRule = wdRestartPage
            End With

            ' In the footnotes range, footnote marks are not superscript
            Set NotesRange = ActiveDocument.StoryRanges(wdFootnotesStory)
            Call MarkFormatting(NotesRange, "\1 ")

        End If

    ' If no endnotes found in active document, adjust formatting of footnotes only
    Else

        ' If footnotes found in active document, footnote count is continuous with footnote marks using Arabic numerals
        If ActiveDocument.Footnotes.Count >= 1 Then

            With Selection.FootnoteOptions
                .Location = wdBottomOfPage
                .NumberingRule = wdRestartContinuous
                .StartingNumber = 1
                .NumberStyle = wdNoteNumberStyleArabic
            End With

            ' In the footnotes range, footnote marks are followed by a period and are not superscript
            Set NotesRange = ActiveDocument.StoryRanges(wdFootnotesStory)
            Call MarkFormatting(NotesRange, "\1. ")

        End If

    End If

End Sub


Sub MarkFormatting(ByRef NotesRange, MyReplaceText As String)

    ' Finds footnote/endnote marks, removing superscript formatting while adjusting format of marks based on specified string pattern
    With NotesRange.Find
        .ClearFormatting
        .Text = "(^2)([. ]{1,})"
        .Replacement.ClearFormatting
        .Replacement.Text = MyReplaceText
        .Replacement.Font.Superscript = False
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

End Sub

