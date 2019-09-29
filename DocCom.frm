VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DocCom 
   Caption         =   "DocCom"
   ClientHeight    =   6408
   ClientLeft      =   0
   ClientTop       =   -4140
   ClientWidth     =   7560
   OleObjectBlob   =   "DocCom.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DocCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





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
    MsgBox "Application error: please ensure a word document is selected"
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
    MsgBox "Application error: please ensure a word document is selected"
    Resume ProcExit

End Sub

Private Sub CommandButton15_Click()

MsgBox ("Use this tool to compare a folder of files against a single input word document. Either Word's track changes or Workshare's DeltaView function can be used. To use, select the relevant files and folders as per the below, ensuring that the 'modified documents' folder contains only the documents you wish to compare against.")

End Sub






Private Sub CommandButton16_Click()

'select folder

On Error GoTo ErrorSub

Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFolderPicker)
'get the number of the button chosen
Dim FolderChosen As Integer
FolderChosen = fd.Show
If FolderChosen <> -1 Then
'didn't choose anything (clicked on CANCEL)
'MsgBox "You chose cancel"
Else
'display name and path of file chosen
TextBox3.Text = fd.SelectedItems(1)
End If

ProcExit:
    Exit Sub

ErrorSub:
    MsgBox "Application error: please ensure a word document is selected"
    Resume ProcExit



End Sub

Private Sub CommandButton17_Click()


'select folder

On Error GoTo ErrorSub

Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFolderPicker)
'get the number of the button chosen
Dim FolderChosen As Integer
FolderChosen = fd.Show
If FolderChosen <> -1 Then
'didn't choose anything (clicked on CANCEL)
'MsgBox "You chose cancel"
Else

TextBox4.Text = fd.SelectedItems(1)
End If

ProcExit:
    Exit Sub

ErrorSub:
    MsgBox "Application error: please ensure a word document is selected"
    Resume ProcExit

End Sub





Private Sub CommandButton6_Click()

On Error GoTo ErrorSub

Dim RefDocLocation As String

RefDocLocation = DefCheck.TextBox1.Text
  
If RefDocLocation = "[Select definitions file]" Then MsgBox ("Error: please select a definitions file using the 'Select File' button above") Else HighlightdefinedtermsYellow (RefDocLocation): MsgBox "Process complete"

ProcExit:
    Exit Sub

ErrorSub:
    MsgBox "Application error: please ensure a word document is selected"
    Resume ProcExit

End Sub

Private Sub CommandButton18_Click()
'workshare process

On Error GoTo ErrorSub

Application.ScreenUpdating = False

'read file locations
Dim OriginalFile As String
Dim AmendedFolder As String


OriginalFile = TextBox1.Text
AmendedFolder = TextBox3.Text
OutputFolder = TextBox4.Text

Dim OrigFile As String
OrigFile = Dir(OriginalFile)

'close all open word docs to prevent issues
    With Application
        
         'Loop Through open documents
        Do Until .Documents.Count = 0
             'Close no save
            .Documents(1).Close SaveChanges:=wdDoNotSaveChanges
        Loop
    End With
    


'import list of files

    Dim MyFile As String
    Dim Counter As Long

    'Create a dynamic array variable, and then declare its initial size
    Dim AmendedFile() As String
    ReDim AmendedFile(1000)
    
    'create file path for amended files

    Dim AmendedFolderFiles As String
    
    
    AmendedFolderFiles = AmendedFolder + "/*.*"
    
    
    'Loop through all the files in the directory by using Dir$ function
    MyFile = Dir$(AmendedFolderFiles)
    Do While MyFile <> ""
        AmendedFile(Counter) = MyFile
        MyFile = Dir$
        Counter = Counter + 1
    Loop
    
Dim nooffiles As Integer
nooffiles = Counter - 1
    
'Reset the size of the array without losing its values by using Redim Preserve
    ReDim Preserve AmendedFile(nooffiles)


'compare for each file in amended docs file

Dim amendfile As String
Dim outputfile As String
Dim original As String

'prep calling shell
Dim wsh As Object
Set wsh = VBA.CreateObject("WScript.Shell")
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1
Dim command As String
Dim a1 As String
Dim a2 As String
Dim a3 As String
Dim a4 As String

a0 = Year(Now()) & "-" & Right("0" & Month(Now()), 2) & "-" & Day(Now())
a1 = "DeltaVw.exe /original="""
a2 = """ /modified="""
a3 = """ /outfile="""
a4 = ".RTF"" /RTF /V"

For i = 0 To nooffiles

    amendfile = AmendedFolder + "\" + AmendedFile(i)
    
    output = OutputFolder + "\" + a0 + " - Comparison of " & OrigFile & " vs " + AmendedFile(i)


    command = a1 + OriginalFile + a2 + amendfile + a3 + output + a4
    
    wsh.Run command, windowStyle, waitOnReturn
       

Next i



Application.ScreenUpdating = True

MsgBox ("Process complete. Comparison files are in the following directory:" & vbCrLf & vbCrLf & OutputFolder)
End



'error handing
ProcExit:
        Exit Sub
    
ErrorSub:
        MsgBox "Application error: please ensure you correctly selected file and folder locations and that the folder containing the modified documents contains only word documents (i.e. no PDFs or spreadsheets)."
        Resume ProcExit




End Sub

Private Sub CommandButton19_Click()
DocCom.Hide
DefCheck.Show
End Sub

Private Sub CommandButton4_Click()

'close button

End

End Sub

Private Sub CommandButton7_Click()

'track changes compare

On Error GoTo ErrorSub

Application.ScreenUpdating = False

'read file locations
Dim OriginalFile As String
Dim AmendedFolder As String
Dim OutputFolder As String

OriginalFile = TextBox1.Text
AmendedFolder = TextBox3.Text
OutputFolder = TextBox4.Text




Dim OrigFile As String
OrigFile = Dir(OriginalFile)


'close all open word docs to prevent issues
    With Application
        
         'Loop Through open documents
        Do Until .Documents.Count = 0
             'Close no save
            .Documents(1).Close SaveChanges:=wdDoNotSaveChanges
        Loop
    End With
    
    
'import list of files

    Dim MyFile As String
    Dim Counter As Long

    'Create a dynamic array variable, and then declare its initial size
    Dim AmendedFile() As String
    ReDim AmendedFile(1000)
    
    'create file path for amended files

    Dim AmendedFolderFiles As String
    
    
    AmendedFolderFiles = AmendedFolder + "/*.*"
    
    
    'Loop through all the files in the directory by using Dir$ function
    MyFile = Dir$(AmendedFolderFiles)
    Do While MyFile <> ""
        AmendedFile(Counter) = MyFile
        MyFile = Dir$
        Counter = Counter + 1
    Loop
    
Dim nooffiles As Integer
nooffiles = Counter - 1
    
'Reset the size of the array without losing its values by using Redim Preserve
    ReDim Preserve AmendedFile(nooffiles)


'compare for each file in amended docs file
Dim amend As String
Dim amendfile As Word.Document
Dim outputfile As Word.Document
Dim original As Word.Document



For i = 0 To nooffiles

    Set original = Documents.Open(OriginalFile)
    amend = AmendedFolder + "\" + AmendedFile(i)
    Set amendfile = Documents.Open(amend)
    

    Application.CompareDocuments _
    OriginalDocument:=original, _
    RevisedDocument:=amendfile, _
    Destination:=wdCompareDestinationNew, _
        Granularity:=wdGranularityWordLevel, _
        CompareFormatting:=True, CompareCaseChanges:=True, CompareWhitespace:= _
        True, CompareTables:=True, CompareHeaders:=True, CompareFootnotes:=True, _
        CompareTextboxes:=True, CompareFields:=True, CompareComments:=True, _
        CompareMoves:=True, RevisedAuthor:="HSF", IgnoreAllComparisonWarnings:= _
        True

    original.Close
    amendfile.Close

Dim a0 As String
a0 = Year(Now()) & "-" & Right("0" & Month(Now()), 2) & "-" & Day(Now())
    output = OutputFolder & "\" & a0 & " - Comparison of " & OrigFile & " vs " & AmendedFile(i)
    Set outputfile = ActiveDocument
    
    outputfile.SaveAs FileName:=output
    outputfile.Close SaveChanges:=False
    

Next i

Set original = Nothing
Set amendfile = Nothing
Set outputfile = Nothing

Application.ScreenUpdating = True

MsgBox ("Process complete. Comparison files are in the following directory:" & vbCrLf & vbCrLf & OutputFolder)




'error handing
ProcExit:
        Exit Sub
    
ErrorSub:
        MsgBox "Application error: please ensure you correctly selected file and folder locations and that the folder containing the modified documents contains only word documents (i.e. no PDFs or spreadsheets)."
        Resume ProcExit

End Sub








