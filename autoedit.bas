Attribute VB_Name = "Module2"
Sub autoedit()
'An automated editing script for Microsoft Word
'Created by James Harper, james@noblepursuits.us

    'Read edits from text file called "editlist_test"
    Open "Mac OS X:Users:jamesharper:GitHub:autoedit:editlist_test" For Input As #1
    Do Until EOF(1)
        Line Input #1, alltext
    Loop
    Close #1

    'Split each line into an element of the array "text"
    text = Split(alltext, vbNewLine)

    'Initialize variables to store original and new edits and counter variables
    Dim edit_orig() As String
    Dim edit_new() As String
    ReDim edit_orig(0)
    ReDim edit_new(0)
    Dim n_orig As Integer      'number of elements in array "edit_orig"
    Dim n_new As Integer      'number of elements in array "edit_new"
    n_orig = 0
    n_new = 0

    'Parse array "text" for edits
    Dim word As Variant
    For Each word In text
        'MsgBox (word)
        If word = "" Then
            'MsgBox ("Blank line reached")
            Exit For
        ElseIf Mid(word, 1, 3) = "###" Then
            'MsgBox ("Skipping line starting with ###")
        ElseIf Mid(word, 1, 3) <> "---" Then
            If n_orig = n_new Then     'Original text found
                ReDim Preserve edit_orig(UBound(edit_orig) + 1)
                edit_orig(n_orig) = Mid(word, 1, Len(word) - 3)   'Remove last three characters in word and save in edit_orig
                'MsgBox (edit_orig(n_orig))
                n_orig = n_orig + 1
                'MsgBox ("n_orig = " & n_orig)
            Else                             'New text found
                ReDim Preserve edit_new(UBound(edit_new) + 1)
                edit_new(n_new) = Mid(word, 1, Len(word) - 3)   'Remove last three characters in word and save in edit_new
                'MsgBox (edit_new(n_new))
                n_new = n_new + 1
                'MsgBox ("n_new = " & n_new)
            End If
        End If
    Next

    'Shrink edit_orig and edit_new to just before first blank edit_orig element
    ReDim Preserve edit_orig(UBound(edit_orig) - 1)
    ReDim Preserve edit_new(UBound(edit_new) - 1)

    'Fix the skipped blank header/footer problem
    Dim lngJunk As Long
    lngJunk = ActiveDocument.Sections(1).Headers(1).range.StoryType

    'Replace all text found in document that equals the elements in "edit_orig" with the corresponding elements in "edit_new"
    Dim rngStory As word.range
    Dim i As Integer
    For Each rngStory In ActiveDocument.StoryRanges
        Do
            For i = 0 To UBound(edit_orig)
                With rngStory.Find
                    .text = edit_orig(i)
                    .Replacement.text = edit_new(i)
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End With
            Next i
            Set rngStory = rngStory.NextStoryRange  'Get next linked story, if any
        Loop Until rngStory Is Nothing
    Next

End Sub
