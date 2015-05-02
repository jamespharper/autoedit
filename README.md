#autoedit
An automated editing script for Microsoft Word

Written in Visual Basic, version 14.0 (2010)

Created by James Harper, PE, james@noblepursuits.us

###Description
This script makes automated edits of the Microsoft Word document in which it is run. As an editor of academic journal articles, I have found that many authors make the same errors regularly, whether grammatical or stylistic; as a result, this script and the edits that it contains have been developed over many hundreds of hours of editing. I have found this script to be extremely useful when beginning to edit a new paper and typically run it with Track Changes turned on for every paper I edit. I have hypothesized that this script saves me an average of about 5 minutes per paper and certainly a ton of irritation; however, its development has undoubtedly taken far longer than the cumulative time I have saved. :-) I hope that it saves you time and toil as well!

###License
Please share and modify as you please; however, please always acknowledge my contributions!

###Algorithm
This script reads pairs of edits to be made from the text file "editlist.txt." After parsing the text from the file into usable arrays (edit_orig and edit_new), the script searches the Word document for each piece of original text found in the elements of the array "edit_orig;" when one is located, it is replaced with the corresponding piece of "new text" (i.e., the edit to be made) in the corresponding element of the array "edit_new."

###Instructions
1. Download the files "autoedit.bas" and "editlist.txt."
2. Open the Visual Basic Editor in Microsoft Word.
3. Import "autoedit.bas."
4. Modify the path in the first line of code to point to the folder where you placed "editlist.txt."
5. Open the Word document you'd like to autoedit and turn on Track Changes (high recommended but not required).
6. Run the script from within the Visual Basic Editor and wait for it to complete
  * The length of time required for the script to complete varies significantly based on the length and complexity (e.g., tables, formatting) of the document being edited and on the number of edit pairs present in "editlist.txt."
  * Word is not usable while the script runs; go watch the Sun set while your document is edited. :-)
7. Check for errors after the script finishes!
  * While a lot of testing has been done on this script, it isn't perfect; this is why I highly recommend Track Changes be turned on when running this script! You have been warned.
