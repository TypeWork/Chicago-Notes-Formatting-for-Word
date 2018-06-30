Sub ChicagoNotesFormatting()

    ' If endnotes found in active document, adjust formatting of footnotes and endnotes
    If ActiveDocument.Endnotes.Count >= 1 Then

        ' Locate endnotes at end of section and format endnote marks using Arabic numerals
        With Selection.EndnoteOptions
            .Location = wdEndOfSection
            .NumberingRule = wdRestartContinuous
            .StartingNumber = 1
            .NumberStyle = wdNoteNumberStyleArabic
        End With

        ' In the endnotes, endnote marks are followed by a period and are not superscript
        Set NotesRange = ActiveDocument.StoryRanges(wdEndnotesStory)
        Call MarkFormatting(NotesRange, "\1. ")


        If ActiveDocument.Footnotes.Count >= 1 Then

            ' Footnote count restarts with each page with footnote marks using symbols
            With Selection.FootnoteOptions
                .Location = wdBottomOfPage
                .StartingNumber = 1
                .NumberStyle = wdNoteNumberStyleSymbol
                .NumberingRule = wdRestartPage
            End With

            ' In the footnotes, footnote marks are not superscript nor are they followed by a period
            Set NotesRange = ActiveDocument.StoryRanges(wdFootnotesStory)
            Call MarkFormatting(NotesRange, "\1 ")

        End If

    ' If no endnotes found in active document, adjust formatting of footnotes
    Else

        If ActiveDocument.Footnotes.Count >= 1 Then

            ' Footnote count is continuous with footnote marks using Arabic numerals
            With Selection.FootnoteOptions
                .Location = wdBottomOfPage
                .NumberingRule = wdRestartContinuous
                .StartingNumber = 1
                .NumberStyle = wdNoteNumberStyleArabic
            End With

            ' In the footnotes, footnote marks are followed by a period and are not superscript
            Set NotesRange = ActiveDocument.StoryRanges(wdFootnotesStory)
            Call MarkFormatting(NotesRange, "\1. ")

        End If

    End If

End Sub


Sub MarkFormatting(ByRef NotesRange, MyReplaceText As String)

    ' Find footnote/endnote marks in specified range (NotesRange).
    ' With find, superscript formatting removed and note marks typeset based on a specified string pattern (MyReplaceText).
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

