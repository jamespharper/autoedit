Sub autoedit()
'To work on:
'Capitalization (e.g., fig. to Fig.)

    Dim rngStory As word.range
    Dim lngJunk As Long
    Dim i As Integer
    Dim temp

    Dim editlist As Dictionary   'Keys = original text, values = new text

    lngJunk = ActiveDocument.Sections(1).Headers(1).range.StoryType   'Fix the skipped blank Header/Footer problem

    'Read list of edits from editlist.csv
    For Each Line In IO.file.ReadAllLines("~/Users/jamesharper/GitHub/autoedit/editlist_test.txt")
        MsgBox (Line)
        temp = Line.Split(",")
        editlist.Add temp(1), temp(2)
    Next

    For Each rngStory In ActiveDocument.StoryRanges
        Do
            For i = 0 To editlist.Count
                With rngStory.Find
                    .text = editlist(i, 0)
                    .Replacement.text = editlist(i, 1)
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End With
            Next i
            Set rngStory = rngStory.NextStoryRange  'Get next linked story, if any
        Loop Until rngStory Is Nothing
    Next

End Sub
