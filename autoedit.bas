Attribute VB_Name = “autoedit”
Sub autoedit()
'To work on:
'Capitalization (e.g., fig. to Fig.)

    Dim rngStory As word.range
    Dim lngJunk As Long
    Dim text(150, 2) As String   'First column is original text, second column is new text
    Dim i As Integer
    
    lngJunk = ActiveDocument.Sections(1).Headers(1).range.StoryType   'Fix the skipped blank Header/Footer problem
    
    text(0, 0) = "on the other hand,"
    text(0, 1) = "conversely,"
    text(3, 0) = "besides,"
    text(3, 1) = "however,"
    text(4, 0) = "close to"
    text(4, 1) = "near"
    text(5, 0) = "yet,"
    text(5, 1) = "however,"
    text(6, 0) = "likewise,"
    text(6, 1) = "similarly,"
    text(7, 0) = "nevertheless,"
    text(7, 1) = "however,"
    text(8, 0) = " mainly"
    text(8, 1) = " primarily"
    text(9, 0) = "carried out"
    text(9, 1) = "performed"
    text(10, 0) = "by means of"
    text(10, 1) = "using"
    text(12, 0) = "but,"
    text(12, 1) = "however,"
    text(13, 0) = "at the same time,"
    text(13, 1) = "concurrently, "
    text(14, 0) = " gives"
    text(14, 1) = " yields"
    text(16, 0) = " enough"
    text(16, 1) = " sufficient"
    text(17, 0) = "our research"
    text(17, 1) = "this study"
    text(18, 0) = " main "
    text(18, 1) = " primary "
    text(19, 0) = "almost the same"
    text(19, 1) = "nearly identical"
    text(20, 0) = "obvious"
    text(20, 1) = "significant"
    text(24, 0) = "complicated"
    text(24, 1) = "complex"
    text(25, 0) = "largely "
    text(25, 1) = "significantly"
    text(26, 0) = "especially"
    text(26, 1) = "particularly"
    text(27, 0) = "the Eq."
    text(27, 1) = "Eq. "
    text(28, 0) = "takes into account"
    text(28, 1) = "accounts for"
    text(29, 0) = "on the other side,"
    text(29, 1) = "conversely,"
    text(30, 0) = "moreover"
    text(30, 1) = ""
    text(31, 0) = "in order "
    text(31, 1) = ""
    text(32, 0) = "where, "
    text(32, 1) = "where "
    text(33, 0) = " as a whole, "
    text(33, 1) = " "
    text(34, 0) = " to be able "
    text(34, 1) = " "
    text(35, 0) = " so, "
    text(35, 1) = " "
    text(36, 0) = " meanwhile, "
    text(36, 1) = " "
    text(37, 0) = " since "
    text(37, 1) = " because "
    text(38, 0) = " impact "
    text(38, 1) = " influence "
    text(40, 0) = "we can see"
    text(40, 1) = "it is shown"
    text(41, 0) = " happen "
    text(41, 1) = " occur "
    text(42, 0) = " takes place "
    text(42, 1) = " occurs "
    text(43, 0) = " in other words, "
    text(43, 1) = " "
    'text(44, 0) = "we "
    'text(44, 1) = "the authors "
    text(45, 0) = "a little"
    text(45, 1) = "slightly"
    text(46, 0) = " means "
    text(46, 1) = " indicates "
    text(47, 0) = " In fact, "
    text(47, 1) = " "
    text(48, 0) = "a lot"
    text(48, 1) = "significantly"
    text(49, 0) = " bad "
    text(49, 1) = " poor "
    text(50, 0) = " it's "
    text(50, 1) = " it is "
    text(51, 0) = " took place "
    text(51, 1) = " occured "
    text(52, 0) = " indexes "
    text(52, 1) = " indices "
    text(53, 0) = " too much "
    text(53, 1) = " significantly "
    text(54, 0) = " owing to "
    text(54, 1) = " due to "
    text(55, 0) = " greatly "
    text(55, 1) = " significantly "
    text(56, 0) = " namely, "
    text(56, 1) = " "
    text(58, 0) = "doesn't"
    text(58, 1) = "does not"
    text(59, 0) = " a bit "
    text(59, 1) = " some "
    text(60, 0) = "the present study"
    text(60, 1) = "this study"
    text(61, 0) = "taken into account"
    text(61, 1) = "considered"
    text(62, 0) = "in the whole"
    text(62, 1) = "throughout the"
    text(63, 0) = " our "
    text(63, 1) = " this "
    text(64, 0) = "the sake of "
    text(64, 1) = " "
    text(65, 0) = "take into account"
    text(65, 1) = "consider"
    text(66, 0) = "in comparison with"
    text(66, 1) = "compared to"
    text(67, 0) = "on the basis of"
    text(67, 1) = "based on"
    text(68, 0) = "in comparison to"
    text(68, 1) = "compared to"
    text(69, 0) = "the point of view"
    text(69, 1) = "the perspective"
    text(70, 0) = "), ("
    text(70, 1) = "; "
    text(71, 0) = " handle "
    text(71, 1) = " manage "
    text(72, 0) = "mostly"
    text(72, 1) = "primarily"
    text(73, 0) = " all of the "
    text(73, 1) = " all "
    text(74, 0) = " all the "
    text(74, 1) = " all "
    text(75, 0) = " need to be "
    text(75, 1) = " must be "
    text(76, 0) = " needs to be "
    text(76, 1) = " must be "
    text(77, 0) = "takes the value"
    text(77, 1) = "equals"
    text(78, 0) = "takes value"
    text(78, 1) = "equals"
    text(79, 0) = "in accordance to"
    text(79, 1) = "in accordance with"
    text(80, 0) = " associated to "
    text(80, 1) = " associated with "
    text(81, 0) = " some of the "
    text(81, 1) = " some "
    text(82, 0) = " most of the "
    text(82, 1) = " most "
    text(83, 0) = " most of "
    text(83, 1) = " most "
    'text(84, 0) = " us "
    'text(84, 1) = " the authors "
    text(85, 0) = " et al "
    text(85, 1) = " et al. "
    text(86, 0) = " multi "
    text(86, 1) = " multiple "
    text(87, 0) = " put forward"
    text(87, 1) = " proposed"
    text(88, 0) = " so as to "
    text(88, 1) = " to "
    text(89, 0) = " needs to "
    text(89, 1) = " must "
    text(90, 0) = " besides "
    text(90, 1) = " in addition to "
    text(91, 0) = " has to be "
    text(91, 1) = " must be "
    text(92, 0) = " have to be "
    text(92, 1) = " must be "
    text(93, 0) = " during the whole "
    text(93, 1) = " throughout the "
    text(94, 0) = "employed"
    text(94, 1) = "used"
    text(95, 0) = " cope with "
    text(95, 1) = " manage "
    text(96, 0) = " starts out "
    text(96, 1) = " begins "
    text(97, 0) = " happens "
    text(97, 1) = " occurs "
    text(98, 0) = " both of the "
    text(98, 1) = " both "
    text(99, 0) = " both of "
    text(99, 1) = " both "
    text(100, 0) = " taking into account "
    text(100, 1) = " considering "
    text(101, 0) = ".("
    text(101, 1) = ". ("
    text(102, 0) = "formulas"
    text(102, 1) = "formulae"
    text(103, 0) = "In the present work, "
    text(103, 1) = "In this study, "
    text(104, 0) = " sharply "
    text(104, 1) = " significantly "
    text(105, 0) = " crucial "
    text(105, 1) = " critical "
    text(106, 0) = "needed"
    text(106, 1) = "required"
    text(107, 0) = "kind"
    text(107, 1) = "type"
    text(108, 0) = " actual "
    text(108, 1) = " real "
    text(109, 0) = " minimal "
    text(109, 1) = " minimum "
    text(110, 0) = " maximal "
    text(110, 1) = " maximum "
    text(111, 0) = "What's more"
    text(111, 1) = "Additionally)"
    text(112, 0) = "the present research"
    text(112, 1) = "this study"
    text(113, 0) = " geometrical "
    text(113, 1) = " geometric "
    text(114, 0) = "nowadays"
    text(114, 1) = "currently"
    text(115, 0) = "and so on"
    text(115, 1) = "etc."
    text(116, 0) = "slightly"
    text(116, 1) = "marginally"
    text(117, 0) = "a part of"
    text(117, 1) = "a portion of"
    text(118, 0) = "figured out"
    text(118, 1) = "determined"
    text(119, 0) = "figure out"
    text(119, 1) = "determine"
    text(120, 0) = " can't "
    text(120, 1) = " cannot "
    text(121, 0) = " isn't"
    text(121, 1) = " is not"
    text(122, 0) = "point of view"
    text(122, 1) = "perspective"
    text(123, 0) = "as long as"
    text(123, 1) = "if"
    text(124, 0) = "taken as"
    text(124, 1) = "considered to be"
    text(125, 0) = "&"
    text(125, 1) = "and"
    text(126, 0) = "starts"
    text(126, 1) = "begins"
    text(127, 0) = "finishes"
    text(127, 1) = "ends"
    text(128, 0) = "for the purpose of"
    text(128, 1) = "to"
    text(129, 0) = "take place"
    text(129, 1) = "occur"
    text(130, 0) = "whereas"
    text(130, 1) = "while"
    text(131, 0) = "almost"
    text(131, 1) = "nearly"
    text(132, 0) = "furthermore"
    text(132, 1) = "also"
    text(133, 0) = "what is more"
    text(133, 1) = "additionally"
    text(134, 0) = "it should be mentioned that"
    text(134, 1) = ""
    text(135, 0) = "in the present paper"
    text(135, 1) = "in this study"
    text(136, 0) = "as can be seen"
    text(136, 1) = "as shown"
    text(137, 0) = "moreover, "
    text(137, 1) = ""
    text(138, 0) = "in the current study"
    text(138, 1) = "in this study"
    text(141, 0) = "on the contrary"
    text(141, 1) = "conversely"
    text(142, 0) = "the present work"
    text(142, 1) = "this study"
    text(143, 0) = "can be seen"
    text(143, 1) = "is shown"
    text(144, 0) = "raise"
    text(144, 1) = "increase"
    text(145, 0) = "whole"
    text(145, 1) = "entire"
    text(146, 0) = " get "
    text(146, 1) = " obtain "
    text(147, 0) = " got "
    text(147, 1) = " obtained "
    text(148, 0) = " et.al"
    text(148, 1) = " et al."
    text(149, 0) = "it is seen"
    text(149, 1) = "it is shown"
    text(150, 0) = "according to"
    text(150, 1) = "based on"
    
    'text(21, 0) = " fig. "
    'text(21, 1) = " Fig. "
    'text(22, 0) = " eqs. "
    'text(22, 1) = " Eqs. "
    'text(15, 0) = " Where "
    'text(15, 1) = " where "
    'text(10, 0) = " eq. "
    'text(10, 1) = " Eq. "
    'text(11, 0) = " the same "
    'text(11, 1) = " identical "
    'text(1, 0) = " around "
    'text(1, 1) = " approximately "
    'text(2, 0) = " about "
    'text(2, 1) = " near "
    
    For Each rngStory In ActiveDocument.StoryRanges
        Do
            For i = 0 To UBound(text)
                With rngStory.Find
                    .text = text(i, 0)
                    .Replacement.text = text(i, 1)
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End With
            Next i
            Set rngStory = rngStory.NextStoryRange  'Get next linked story, if any
        Loop Until rngStory Is Nothing
    Next
End Sub
