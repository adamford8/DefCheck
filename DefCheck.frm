VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DefCheck 
   Caption         =   "DefCheck"
   ClientHeight    =   9324
   ClientLeft      =   0
   ClientTop       =   -4140
   ClientWidth     =   6732
   OleObjectBlob   =   "DefCheck.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DefCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
On Error GoTo ErrorSub
ColourCapitalLetters
MsgBox "Process complete"
Application.ScreenUpdating = True

ProcExit:
    Exit Sub

ErrorSub:
    MsgBox "Application error: please ensure a word document is open to DefCheck"
    Resume ProcExit

End Sub

Private Sub CommandButton10_Click()
MsgBox "Use the button above to specify the word document which contains the defined terms list. This list can only reliably be created manually (it takes 5minutes for a medium length document)." & vbCrLf & vbCrLf & "The required input format is: one defined term per line, Case matching the Case of the defined term. For example:" & vbCrLf & vbCrLf & "Agreement" & vbCrLf & "Parties" & vbCrLf & "Herbert Smith Freehills" & vbCrLf & vbCrLf & "DefCheck will then search for exact matches with the defined terms list, highlighting matches and close matches where relevant. When creating the input file, the 'pluraling' element of each defined term should be removed, e.g.:" & vbCrLf & vbCrLf & "Obligations to Obligation" & vbCrLf & "Equity Documents to Equity Document" & vbCrLf & "Facilities to Facilit (changing to Facility would exclude the plural (therefore also amend odd ending singular nouns))" & vbCrLf & vbCrLf & "Final note: remember to include definitions not in the definitions clause."

End Sub

Private Sub CommandButton11_Click()
MsgBox "The DefCheck process output is self-explanatory. In summary:" & vbCrLf & vbCrLf & "Green highlighting: the word or phrase is defined (green can generally be ignored)." & vbCrLf & vbCrLf & "Yellow highlighting: the word or phrase would be defined if capitalised, or vice-versa. (E.g. this highlights instances of 'day' and 'year' when it should be 'Day' and 'Year'.)" & vbCrLf & vbCrLf & "Red highlighting: all red words are capitalised but not defined. You should consider whether a definition should be added for these words. This process will also identify false positives (e.g. days, first words in sentences, acronyms, etc.). If there are many, repetitive false positives, an exclusions list of non-propblematic capitalised words can be created at the bottom of the main user form."

End Sub

Private Sub CommandButton12_Click()
On Error GoTo ErrorSub
'select input file
Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)
'get the number of the button chosen
Dim FileChosen As Integer
FileChosen = fd.Show
If FileChosen <> -1 Then
'didn't choose anything (clicked on CANCEL)
'MsgBox "You chose cancel"
Else
'display name and path of file chosen
TextBox1.Text = fd.SelectedItems(1)
End If

ProcExit:
    Exit Sub

ErrorSub:
    MsgBox "Application error: please ensure a word document is open to DefCheck"
    Resume ProcExit

End Sub

Private Sub CommandButton13_Click()
On Error GoTo ErrorSub
'remove highlighting on terms specified in exlusion file
Dim RefDocExclusionLocation As String

RefDocExclusionLocation = DefCheck.TextBox2.Text
  
If RefDocExclusionLocation = "[Select exclusions file]" Then MsgBox ("Error: please select an exclusions file using the 'Select File' button to the right") Else RemoveHighlight (RefDocExclusionLocation): MsgBox "Process complete"

ProcExit:
    Exit Sub

ErrorSub:
    MsgBox "Application error: please ensure a word document is open to DefCheck"
    Resume ProcExit

End Sub

Private Sub CommandButton14_Click()
On Error GoTo ErrorSub
'select exclusion files
'select input file
Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)
'get the number of the button chosen
Dim FileChosen As Integer
FileChosen = fd.Show
If FileChosen <> -1 Then
'didn't choose anything (clicked on CANCEL)
'MsgBox "You chose cancel"
Else
'display name and path of file chosen
TextBox2.Text = fd.SelectedItems(1)
End If

ProcExit:
    Exit Sub

ErrorSub:
    MsgBox "Application error: please ensure a word document is open to DefCheck"
    Resume ProcExit

End Sub

Private Sub CommandButton15_Click()

MsgBox ("For basic proofreading, the following steps should be followed:" & vbCrLf & vbCrLf & "1. Extract the defined terms you wish to check for into a new, locally saved document." & vbCrLf & vbCrLf & "2. Open the document you want to DefCheck, and reopen the DefCheck window." & vbCrLf & vbCrLf & "3. Using the 'Select file' button, select the document containing definitions which you have just created." & vbCrLf & vbCrLf & "4. Click the 'Full process - DefCheck your document' button." & vbCrLf & vbCrLf & "5. Wait for the process to complete and then print and analyse the output." & vbCrLf & vbCrLf & "6. Correct identified errors." & vbCrLf & vbCrLf & "7. Run additional steps, as required (e.g. run Function A to identify unused definitons, Function B to identify square brackets, Function C to to highlight all references to 'Clauses' and 'Paragraphs' etc. for cross-reference checking, and Function D to highlight all numbers, currencies and  dates for a final check of values).")

End Sub

Private Sub CommandButton16_Click()

'Function D - highlight numbers currencies and dates pink.
On Error GoTo ErrorSub
ColourNumbersetc
MsgBox "Process complete"


ProcExit:
    Exit Sub
ErrorSub:
    MsgBox "Application error: please ensure a word document is open to DefCheck"
    Resume ProcExit

End Sub

Private Sub CommandButton17_Click()

MsgBox ("Function E can be used to remove highlighting on any specified words or characters." & vbCrLf & vbCrLf & "Simply write the terms you wish to remove highlighting on (one to a line) in a new document, save this document locally, select the document from this Function E frame, and then run 'Remove highlighting'." & vbCrLf & vbCrLf & "This can be useful where there are lots of false positives being flagged,e.g. in a document with a lot of names or months.")

End Sub

Private Sub CommandButton18_Click()

DefCheck.Hide
DocCom.Show

End Sub

Private Sub CommandButton2_Click()

On Error GoTo ErrorSub

Dim RefDocLocation As String

RefDocLocation = DefCheck.TextBox1.Text
  
If RefDocLocation = "[Select definitions file]" Then MsgBox ("Error: please select a file using the 'Select File' button above") Else Highlightdefinedterms (RefDocLocation): MsgBox "Process complete"

ProcExit:
    Exit Sub

ErrorSub:
    MsgBox "Application error: please ensure a word document is open to DefCheck"
    Resume ProcExit

End Sub

Private Sub CommandButton3_Click()

On Error GoTo ErrorSub

Dim RefDocLocation As String

RefDocLocation = DefCheck.TextBox1.Text
  
If RefDocLocation = "[Select definitions file]" Then MsgBox ("Error: please select a definitions file using the 'Select File' button above") Else MakeDefinedTermsTableAsAnnex (RefDocLocation): MsgBox "Process complete"

ProcExit:
    Exit Sub

ErrorSub:
    MsgBox "Application error: please ensure a word document is open to DefCheck"
    Resume ProcExit

End Sub

Private Sub CommandButton4_Click()

End

End Sub



Private Sub CommandButton6_Click()

On Error GoTo ErrorSub

Dim RefDocLocation As String

RefDocLocation = DefCheck.TextBox1.Text
  
If RefDocLocation = "[Select definitions file]" Then MsgBox ("Error: please select a definitions file using the 'Select File' button above") Else HighlightdefinedtermsYellow (RefDocLocation): MsgBox "Process complete"

ProcExit:
    Exit Sub

ErrorSub:
    MsgBox "Application error: please ensure a word document is open to DefCheck"
    Resume ProcExit

End Sub

Private Sub CommandButton7_Click()

On Error GoTo ErrorSub

'main def check process - calls all subs

Dim RefDocLocation As String

RefDocLocation = DefCheck.TextBox1.Text
  
If RefDocLocation = "[Select definitions file]" Then MsgBox ("Error: please select a definitions file using the 'Select File' button above") Else MsgBox ("Note: this process may take 5 minutes + for large documents."): ColourCapitalLetters: HighlightdefinedtermsYellow (RefDocLocation): Highlightdefinedterms (RefDocLocation): MsgBox "Process complete"

ProcExit:
    Exit Sub

ErrorSub:
    MsgBox "Application error: please ensure a word document is open to DefCheck"
    Resume ProcExit

End Sub


Private Sub CommandButton8_Click()
On Error GoTo ErrorSub
ColourSQB
MsgBox "Process complete"


ProcExit:
    Exit Sub
ErrorSub:
    MsgBox "Application error: please ensure a word document is open to DefCheck"
    Resume ProcExit


End Sub

Private Sub CommandButton9_Click()
On Error GoTo ErrorSub
ColourPCA
MsgBox "Process complete"


ProcExit:
    Exit Sub

ErrorSub:
    MsgBox "Application error: please ensure a word document is open to DefCheck"
    Resume ProcExit


End Sub


Private Sub Label4_Click()

End Sub

Private Sub UserForm_Initialize()
'TextBox1.Text = Replace(Environ("APPDATA"), "AppData\Roaming", "Desktop\definitions.docx")
End Sub

Private Sub Highlightdefinedterms(RefDocLocation As String)

Application.ScreenUpdating = False
Dim Doc As Document, RefDoc As Document, Rng As Range
Dim StrTerms As String, strFnd As String, StrPages As String
Dim i As Long, j As Long, StrOut As String, StrBreak As String


On Error GoTo ErrMsg
Set Doc = ActiveDocument
Set RefDoc = Documents.Open(RefDocLocation, AddtorecentFiles:=False)
StrTerms = RefDoc.Range.Text
RefDoc.Close False
Set RefDoc = Nothing

For i = 0 To UBound(Split(StrTerms, vbCr))
  strFnd = Trim(Split(StrTerms, vbCr)(i))
  If strFnd = "" Then GoTo NullString
  StrPages = ""
  
  'strFnd is defined.
  
  ChangeColourToGreen (strFnd)
  
 
NullString:
Next i

Application.ScreenUpdating = True
Exit Sub

ErrMsg:
MsgBox "An error has occured - check the definitions file name/path is correct?"
  
End Sub

Private Sub ChangeColourToGreen(DefinedTerm As String)

Options.DefaultHighlightColorIndex = wdBrightGreen
'Options.DefaultHighlightColorIndex = wdRed

With ActiveDocument.Range.Find
    .Text = DefinedTerm
    '.Replacement.Font.Italic = False
    .Replacement.Font.Color = wdColorWhite
    .Replacement.Highlight = True
    ' .Replacement.Frame
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute Replace:=wdReplaceAll
End With

With ActiveDocument.Range.Find
    .Text = DefinedTerm & "'s"
    '.Replacement.Font.Italic = False
    .Replacement.Font.Color = wdColorWhite
    .Replacement.Highlight = True
    ' .Replacement.Frame
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute Replace:=wdReplaceAll
End With

With ActiveDocument.Range.Find
    .Text = DefinedTerm & "s"
    '.Replacement.Font.Italic = False
    .Replacement.Font.Color = wdColorWhite
    .Replacement.Highlight = True
    ' .Replacement.Frame
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = True
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute Replace:=wdReplaceAll
End With
   
End Sub


Private Sub MakeDefinedTermsTableAsAnnex(DefinitionsFileLocation As String)
'create an annex to the main document including a list of where all the defined terms are used
Application.ScreenUpdating = False
'define terms
Dim Doc As Document
Dim DefinitionsFile As Document
Dim Definitions As String
Dim PageVar As Long
Dim LinetoWrite As String
Dim Rng As Range
Dim NewLine As String
Dim Definitions2 As String
Dim DefPage As String
Dim i As Integer

'error handing
On Error GoTo ErrMsg
'set up defined terms
LinetoWrite = "Term" & vbTab & "Found on pages" & vbCr
Set Doc = ActiveDocument
Set DefinitionsFile = Documents.Open(DefinitionsFileLocation, AddtorecentFiles:=False)

'read in definitions
Definitions = DefinitionsFile.Range.Text
DefinitionsFile.Close False
Set DefinitionsFile = Nothing

'analyse the definitions content and cycle through for each term
For i = 0 To UBound(Split(Definitions, vbCr))
  Definitions2 = Trim(Split(Definitions, vbCr)(i))
  If Definitions2 = "" Then GoTo NullString
  DefPage = ""
  With Doc.Content
 With .Find
.ClearFormatting
.Replacement.ClearFormatting
.Text = Definitions2
.Wrap = wdFindStop
.Format = False
.MatchCase = True
'.MatchWildcards = True
.MatchWildcards = False
.MatchWholeWord = False
.Execute
End With
'PageVar = 0
   
'read page numbers

DefPage = "0 "

'loop
Do While .Find.Found
If PageVar <> .Duplicate.Information(wdActiveEndPageNumber) Then
PageVar = .Duplicate.Information(wdActiveEndPageNumber)
DefPage = DefPage & PageVar & " "
End If
.Find.Execute
    Loop
'end loop

DefPage = Replace(Trim(DefPage), " ", ",")


If DefPage <> "" Then
If i Mod 2 = 0 Then NewLine = vbCr Else NewLine = vbCr
'call the function

LinetoWrite = LinetoWrite & Definitions2 & vbTab & ProcessPages(DefPage) & NewLine
End If
End With
NullString:
Next i
Set Rng = Doc.Range.Characters.Last
With Rng
.InsertAfter vbCr & Chr(12) & LinetoWrite
.Start = .Start + 2
.ParagraphFormat.Alignment = wdAlignParagraphLeft
.ConvertToTable Separator:=vbTab, NumColumns:=2, AutoFit:=False
'format first line of table
With .Tables(1).Rows.First.Range
.Font.Bold = True
.ParagraphFormat.Alignment = wdAlignParagraphCenter
End With
End With

Application.ScreenUpdating = True

Exit Sub

ErrMsg:
MsgBox "An error has occured - check the definitions file name/path is correct?"

End Sub
 
Function ProcessPages(DefPage As String)

'define terms
Dim Array1()
Dim PageVar As Integer
Dim i As Integer
Dim k As Integer
ReDim Array1(UBound(Split(DefPage, ",")))

'run process
For i = 0 To UBound(Split(DefPage, ","))
    Array1(i) = Split(DefPage, ",")(i)
Next
For i = 0 To UBound(Array1) - 1
  If IsNumeric(Array1(i)) Then
    k = 2
    For PageVar = i + 2 To UBound(Array1)
      If CInt(Array1(i) + k) <> CInt(Array1(PageVar)) Then Exit For
      Array1(PageVar - 1) = ""
      k = k + 1
    Next
    i = PageVar - 1
  End If
Next
ProcessPages = Replace(Replace(Replace(Replace(Join(Array1, ","), ",,", " "), " ,", " "), "  ", " "), " ", "-")

If Mid(ProcessPages, 1, 2) = "0," Then
ProcessPages = Right(ProcessPages, Len(ProcessPages) - 2)
ElseIf Len(ProcessPages) = 1 Then
'ProcessPages = Right(ProcessPages, Len(ProcessPages) - 1)
    If ProcessPages = "0" Then ProcessPages = "[not found]"
    End If
'display not found message when defined term not found at all
'ProcessPages = "[not found]"
'ProcessPages = Right(ProcessPages, Len(ProcessPages) - 1)


'If Len(ProcessPages) > 1 Then
'ProcessPages = Right(ProcessPages, Len(ProcessPages) - 2)
'Else
'display not found message when defined term not found at all
'ProcessPages = "[not found]"
'ProcessPages = Right(ProcessPages, Len(ProcessPages) - 1)
'End If

End Function
    


Private Sub colourcaps()
    

    
End Sub

Private Sub ColourCapitalLetters()

Application.ScreenUpdating = False

    Options.DefaultHighlightColorIndex = wdRed
    
    With ActiveDocument.Content.Find
        .Text = "<[A-Z]*[a-z]>"
        .MatchWildcards = True
        .MatchCase = True
        .Replacement.Highlight = True
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Text = "[A-Z]"
        .MatchWildcards = True
        .MatchCase = True
        .Replacement.Highlight = True
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    Options.DefaultHighlightColorIndex = wdNoHighlight
    

    With ActiveDocument.Content.Find
        .Text = " "
        .MatchWildcards = True
        .MatchCase = True
        .Replacement.Highlight = True
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With ActiveDocument.Content.Find
        .Text = "<[a-z]*[a-z]>"
        .MatchWildcards = True
        .MatchCase = True
        .Replacement.Highlight = True
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
Application.ScreenUpdating = True

End Sub


Private Sub HighlightdefinedtermsYellow(RefDocLocation As String)

Application.ScreenUpdating = False
Dim Doc As Document, RefDoc As Document, Rng As Range
Dim StrTerms As String, strFnd As String, StrPages As String
Dim i As Long, j As Long, StrOut As String, StrBreak As String

On Error GoTo ErrMsg
Set Doc = ActiveDocument
Set RefDoc = Documents.Open(RefDocLocation, AddtorecentFiles:=False)
StrTerms = RefDoc.Range.Text
RefDoc.Close False
Set RefDoc = Nothing

For i = 0 To UBound(Split(StrTerms, vbCr))
  strFnd = Trim(Split(StrTerms, vbCr)(i))
  If strFnd = "" Then GoTo NullString
  StrPages = ""
  
  'strFnd is defined.
  
  ChangeColourToYellow (strFnd)
  
 
NullString:
Next i

Application.ScreenUpdating = True

Exit Sub

ErrMsg:
MsgBox "An error has occured - check the definitions file name/path is correct?"
  
End Sub

Private Sub ChangeColourToYellow(DefinedTerm As String)

Options.DefaultHighlightColorIndex = wdYellow
'Options.DefaultHighlightColorIndex = wdRed

With ActiveDocument.Range.Find
    .Text = DefinedTerm
    '.Replacement.Text = Code
    '.Replacement.Font.Italic = False
    '.Replacement.Font.Color = wdColorBlue
    .Replacement.Highlight = True
    ' .Replacement.Frame
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute Replace:=wdReplaceAll
End With

   
End Sub



Private Sub ColourSQB()

Application.ScreenUpdating = False

Dim n, i As Integer

    'number of numbers is:
n = 2

Dim Codes, Out As String
ReDim Codes(n)
    'insert code numbers here:

Codes(1) = "["
Codes(2) = "]"

    'workings

For i = 1 To n
    ColourYellowSB (Codes(i))
Next i
    
Application.ScreenUpdating = True

End Sub
Private Sub ColourYellowSB(Code As String)

'Options.DefaultHighlightColorIndex = wdBrightGreen
Options.DefaultHighlightColorIndex = wdYellow

With ActiveDocument.Range.Find
    .Text = Code
    '.Replacement.Text = Code
    '.Replacement.Font.Italic = False
    '.Replacement.Font.Color = wdColorBlue
    .Replacement.Highlight = True
    ' .Replacement.Frame
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute Replace:=wdReplaceAll
End With
   
End Sub



Private Sub ColourPCA()

Application.ScreenUpdating = False

Dim n, i As Integer

    'number of numbers is:
n = 7

Dim Codes, Out As String
ReDim Codes(n)
    'insert code numbers here:

Codes(1) = "paragraph"
Codes(2) = "clause"
Codes(3) = "appendix"
Codes(4) = "schedule"
Codes(5) = "section"
Codes(6) = "annex"
Codes(6) = "exhibit"

    'workings

For i = 1 To n
    ColourYellowPCA (Codes(i))
Next i
    
Application.ScreenUpdating = True

End Sub
Private Sub ColourYellowPCA(Code As String)

'Options.DefaultHighlightColorIndex = wdBrightGreen
Options.DefaultHighlightColorIndex = wdYellow

With ActiveDocument.Range.Find
    .Text = Code
    '.Replacement.Text = Code
    '.Replacement.Font.Italic = False
    '.Replacement.Font.Color = wdColorBlue
    .Replacement.Highlight = True
    ' .Replacement.Frame
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute Replace:=wdReplaceAll
End With
   
End Sub

Private Sub RemoveHighlight(RefDocExclusionLocation As String)

Application.ScreenUpdating = False
Dim Doc As Document, RefDoc As Document, Rng As Range
Dim StrTerms As String, strFnd As String, StrPages As String
Dim i As Long, j As Long, StrOut As String, StrBreak As String


On Error GoTo ErrMsg
Set Doc = ActiveDocument
Set RefDoc = Documents.Open(RefDocExclusionLocation, AddtorecentFiles:=False)
StrTerms = RefDoc.Range.Text
RefDoc.Close False
Set RefDoc = Nothing

For i = 0 To UBound(Split(StrTerms, vbCr))
  strFnd = Trim(Split(StrTerms, vbCr)(i))
  If strFnd = "" Then GoTo NullString
  StrPages = ""
  
  'strFnd is defined.
  
  ChangeColourToClear (strFnd)
  
 
NullString:
Next i

Application.ScreenUpdating = True
Exit Sub

ErrMsg:
MsgBox "An error has occured - check the exclusions document file name/path is correct?"
  
End Sub

Private Sub ChangeColourToClear(DefinedTerm As String)

Options.DefaultHighlightColorIndex = wdNoHighlight

With ActiveDocument.Range.Find
    .Text = DefinedTerm
    '.Replacement.Text = Code
    '.Replacement.Font.Italic = False
    .Replacement.Font.Color = wdColorBlack
    .Replacement.Highlight = True
    ' .Replacement.Frame
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute Replace:=wdReplaceAll
End With

   
End Sub

Private Sub ColourNumbersetc()


Application.ScreenUpdating = False

Dim n, i As Integer
Dim Codes, Out As String

    'number of numbers is:
n = 16
ReDim Codes(n)
    'insert code numbers here:
Codes(1) = "1"
Codes(2) = "2"
Codes(3) = "3"
Codes(4) = "4"
Codes(5) = "5"
Codes(6) = "6"
Codes(7) = "7"
Codes(8) = "8"
Codes(9) = "9"
Codes(10) = "0"
Codes(11) = "£"
Codes(12) = "$"
Codes(13) = "¥"
Codes(14) = "€"
Codes(15) = "%"
Codes(16) = "@"
    'workings
For i = 1 To n
    ColourPinkNos1 (Codes(i))
Next i
    

n = 37
ReDim Codes(n)
    'insert code numbers here:
Codes(1) = "January"
Codes(2) = "February"
Codes(3) = "March"
Codes(4) = "April"
Codes(5) = "May"
Codes(6) = "June"
Codes(7) = "July"
Codes(8) = "August"
Codes(9) = "September"
Codes(10) = "October"
Codes(11) = "November"
Codes(12) = "December"
Codes(13) = "Jan"
Codes(14) = "Feb"
Codes(15) = "Mar"
Codes(16) = "Apr"
Codes(17) = "May"
Codes(18) = "Jun"
Codes(19) = "Jul"
Codes(20) = "Aug"
Codes(21) = "Sep"
Codes(22) = "Oct"
Codes(23) = "Nov"
Codes(24) = "Dec"
Codes(25) = "GBP"
Codes(26) = "AUD"
Codes(27) = "YEN"
Codes(28) = "USD"
Codes(29) = "AM"
Codes(30) = "PM"
Codes(31) = "Monday"
Codes(32) = "Tuesday"
Codes(33) = "Wednesday"
Codes(34) = "Thursday"
Codes(35) = "Friday"
Codes(36) = "Saturday"
Codes(37) = "Sunday"
    
    'workings
For i = 1 To n
    ColourPinkNos2 (Codes(i))
Next i
    
n = 63
ReDim Codes(n)
    'insert code numbers here:
Codes(1) = "daily"
Codes(2) = "weekly"
Codes(3) = "monthly"
Codes(4) = "yearly"
Codes(5) = "fortnightly"
Codes(6) = "bi-weekly"
Codes(7) = "midday"
Codes(8) = "midnight"
Codes(9) = "zero"
Codes(10) = "one"
Codes(11) = "two"
Codes(12) = "three"
Codes(13) = "four"
Codes(14) = "five"
Codes(15) = "six"
Codes(16) = "seven"
Codes(17) = "eight"
Codes(18) = "nine"
Codes(19) = "ten"
Codes(20) = "eleven"
Codes(21) = "twelve"
Codes(22) = "thirteen"
Codes(23) = "fourteen"
Codes(24) = "fifteen"
Codes(25) = "sixteen"
Codes(26) = "seventeen"
Codes(27) = "eighteen"
Codes(28) = "nineteen"
Codes(29) = "twenty"
Codes(30) = "thirty"
Codes(31) = "fouty"
Codes(32) = "fifty"
Codes(33) = "sixty"
Codes(34) = "seventy"
Codes(35) = "eighty"
Codes(36) = "ninety"
Codes(37) = "hundred"
Codes(38) = "thousand"
Codes(39) = "million"
Codes(40) = "billion"
Codes(41) = "trillion"
Codes(42) = "half"
Codes(43) = "third"
Codes(44) = "quarter"
Codes(45) = "fifth"
Codes(46) = "sixth"
Codes(47) = "seventh"
Codes(48) = "eighth"
Codes(49) = "ninth"
Codes(50) = "tenth"
Codes(51) = "first"
Codes(52) = "second"
Codes(53) = "third"
Codes(54) = "fourth"
Codes(55) = "third"
Codes(56) = "fourth"
Codes(57) = "annually"
Codes(58) = "month"
Codes(59) = "year"
Codes(60) = "day"
Codes(61) = "months"
Codes(62) = "years"
Codes(63) = "days"
    'workings
For i = 1 To n
    ColourPinkNos3 (Codes(i))
Next i
 
    
Application.ScreenUpdating = True

End Sub
Private Sub ColourPinkNos1(Code As String)

'Options.DefaultHighlightColorIndex = wdBrightGreen
Options.DefaultHighlightColorIndex = wdPink

With ActiveDocument.Range.Find
    .Text = Code
    '.Replacement.Text = Code
    '.Replacement.Font.Italic = False
    '.Replacement.Font.Color = wdColorBlue
    .Replacement.Highlight = True
    ' .Replacement.Frame
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute Replace:=wdReplaceAll
End With
   
End Sub

Private Sub ColourPinkNos2(Code As String)

'Options.DefaultHighlightColorIndex = wdBrightGreen
Options.DefaultHighlightColorIndex = wdPink

With ActiveDocument.Range.Find
    .Text = Code
    '.Replacement.Text = Code
    '.Replacement.Font.Italic = False
    '.Replacement.Font.Color = wdColorBlue
    .Replacement.Highlight = True
    ' .Replacement.Frame
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = True
    .MatchWholeWord = True
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute Replace:=wdReplaceAll
End With
   
End Sub

Private Sub ColourPinkNos3(Code As String)

'Options.DefaultHighlightColorIndex = wdBrightGreen
Options.DefaultHighlightColorIndex = wdPink

With ActiveDocument.Range.Find
    .Text = Code
    '.Replacement.Text = Code
    '.Replacement.Font.Italic = False
    '.Replacement.Font.Color = wdColorBlue
    .Replacement.Highlight = True
    ' .Replacement.Frame
    .Forward = True
    .Wrap = wdFindContinue
    .MatchCase = False
    .MatchWholeWord = True
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute Replace:=wdReplaceAll
End With
   
End Sub
