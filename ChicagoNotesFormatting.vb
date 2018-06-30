Sub ChicagoNotesFormatting()

    ' If endnotes found in active document...
    If ActiveDocument.Endnotes.Count >= 1 Then

        Set NotesOptions = Selection.EndnoteOptions
        Set NotesRange = ActiveDocument.StoryRanges(wdEndnotesStory)
        
        ' If endnotes use symbols, a period does not immediately follow endnote marks in the endnotes
        If NotesOptions.NumberStyle = wdNoteNumberStyleSymbol Then
            Call MarkFormatting(NotesRange, "\1 ")
        Else
            Call MarkFormatting(NotesRange, "\1. ")
        End If

    End If

    ' If footnotes found in active document...
    If ActiveDocument.Footnotes.Count >= 1 Then

        Set NotesOptions = Selection.FootnoteOptions
        Set NotesRange = ActiveDocument.StoryRanges(wdFootnotesStory)
        
        ' If footnotes use symbols, a period does not immediately follow footnote marks in the footnotes
        If NotesOptions.NumberStyle = wdNoteNumberStyleSymbol Then
            Call MarkFormatting(NotesRange, "\1 ")
        Else
            Call MarkFormatting(NotesRange, "\1. ")
        End If

    End If

End Sub


Sub MarkFormatting(ByRef NotesRange, NoteMarkText As String)

    ' Find footnote/endnote marks in specified range (NotesRange).
    ' With find, superscript formatting removed and note marks typeset based on a specified string pattern (NoteMarkText).
    With NotesRange.Find
        .ClearFormatting
        .Text = "(^2)([. ]{1,})"
        .Replacement.ClearFormatting
        .Replacement.Text = NoteMarkText
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

