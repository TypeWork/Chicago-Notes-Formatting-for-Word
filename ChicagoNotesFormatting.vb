Sub ChicagoNotesFormatting()

    ' If endnotes found in active document...
    If ActiveDocument.Endnotes.Count >= 1 Then

        ' Endnotes are placed at the end of a section and labelled using Arabic numerals
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

        ' Footnotes are restarted on each and labelled using symbols (without a following period)
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

    ' If no endnotes found in active document...
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

