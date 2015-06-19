Attribute VB_Name = "Module2"
Sub autoedit()
'An automated editing script for Microsoft Word that works with Track Changes
'Created by James Harper, james@noblepursuits.us, www.noblepursuits.us
'
'INSTRUCTIONS TO USER:
'1) Change the path in the first line of code to point to where you placed the text file "editlist."
'2) Modify the text file "editlist" as you need, following the instructions in that file.
'3) Run the script and enjoy!
'
'For questions, comments or problems, please email james@noblepursuits.us.
    
    'Read edits from text file called "editlist"
    Open "Mac OS X:Users:jamesharper:GitHub:autoedit:editlist" For Input As #1
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
        Else
            editpair = Split(word, "--->")
            ReDim Preserve edit_orig(UBound(edit_orig) + 1)
            ReDim Preserve edit_new(UBound(edit_new) + 1)
            edit_orig(n) = editpair(0)
            'MsgBox (edit_orig(n))
            edit_new(n) = Mid(editpair(1), 1, Len(editpair(1)) - 3)   'Remove last three characters in editpair(1) and save in edit_new
            'MsgBox (edit_new(n))
            n = n + 1
            'MsgBox ("n = " & n)
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
