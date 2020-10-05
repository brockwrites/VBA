Sub DeleteAllTabStops()
'
' DeleteAllTabStops Macro
' Brock created to delete all tab stops from selection
' Assign to ALT+T
'
    Application.Selection.Paragraphs.TabStops.ClearAll
    
End Sub

Sub AddPeriods()
'
' AddPeriods Macro
' Brock created to add periods at end of all selected paragraphs/lists
' Assign to ALT+.
'
  Dim oPara As Word.Paragraph
  Dim Rng As Range
  Dim text As String

  For Each oPara In Application.Selection.Paragraphs
      If Len(oPara.Range.text) > 1 Then
         Set Rng = ActiveDocument.Range(oPara.Range.Start, oPara.Range.End - 1)
         Rng.InsertAfter "."
      End If
   Next

End Sub
Sub DeleteAllTabStopsAndCreate1()
'
' DeleteAllTabStopsAndCreate1 Macro
' Brock created to add Issues Addressed line to release notes (with release number)
' Assign to ALT+S
'
    With Application.Selection.Paragraphs.TabStops
        .ClearAll
        .Add Position:=InchesToPoints(6.5), _
        Alignment:=wdAlignTabRight
    End With
    
    'I commented out the stuff below so I could just create the tab stop
'    Selection.TypeText text:="Issues Addressed: No Associated Issues" & vbTab _
'        & "19.2.1"
 '   Selection.HomeKey unit:=wdLine, Extend:=wdExtend
 '   Selection.Font.Italic = wdToggle
           
End Sub
Sub IssuesAddressed()
'
' IssuesAddressed Macro
' Brock Created to type only Issues Addressed: No Associated Issues
' Assign to ALT+I
'
    Selection.TypeText text:="Issues Addressed: No Associated Issues"
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    Selection.Font.Italic = True

End Sub
Sub TestSteps()
'
' TestSteps Macro
' Brock Created to type bold Test Steps: line
' Assign to ALT+T
'
    Selection.TypeText text:="Test Steps:"
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    Selection.Font.Bold = True
    Selection.EndKey Unit:=wdLine
    Selection.Font.Bold = wdToggle

End Sub
Sub ConvertNumbersToText()
'    Dim oPara As Word.Paragraph
'    Dim Rng As Range

'This section finds all numbered lists, converts them to text, and inserts 3 NBSP's before them
'For Each oPara In ActiveDocument.Paragraphs
'    If oPara.Range.ListFormat.ListType = WdListType.wdListBullet Then
'        Set Rng = ActiveDocument.Range(oPara.Range.Start, oPara.Range.End)
'            Rng.ListFormat.ConvertNumbersToText
'            Rng.InsertBefore Chr$(160) & Chr$(160) & Chr$(160)
'      End If
'   Next

'This section finds all Symbol font and converts it to Calibri
'With ActiveDocument.Range.Find
'    .Font.Name = "Symbol"
'    .Replacement.Font.Name = "Calibri"
'    .Execute Replace:=wdReplaceAll
'End With

'This converts all numbered/bulleted lists to text
ActiveDocument.Range.ListFormat.ConvertNumbersToText

'This replaces all bullets and adds 3 non-breaking spaces in front of them
Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = ChrW(61623)
        .Replacement.text = "^s•^s "
        .Replacement.Font.Name = "Calibri"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
'This finds all numbered list steps and puts 2 non-breaking spaces in front of them
With ActiveDocument.Range.Find
    .text = "^13([1-99].^t)"
    .Replacement.text = "^13^s^s\1"
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
End With

'This finds all lettered list steps and puts 15 non-breaking spaces in front of them
With ActiveDocument.Range.Find
    .text = "^13([a-z].^t)"
    .Replacement.text = "^13^s^s^s^s^s\1"
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
End With

End Sub
Sub CheckLists()
    Dim oL As List
    Dim sMsg As String
    Dim J As Integer
    Dim K As Integer

    J = ActiveDocument.Lists.Count
    For Each oL In ActiveDocument.Lists
        K = K + 1
        oL.Range.Select

        sMsg = "This is list " & K & " of " & J
        sMsg = sMsg & " lists in the document." & vbCrLf & vbCrLf
        sMsg = sMsg & "This list is this type: "
        Select Case oL.Range.ListFormat.ListType
            Case wdListBullet
                sMsg = sMsg & "wdListBullet"
            Case wdListListNumOnly
                sMsg = sMsg & "wdListListNumOnly"
            Case wdListMixedNumbering
                sMsg = sMsg & "wdListMixedNumbering"
            Case wdListNoNumbering
                sMsg = sMsg & "wdListNoNumbering"
            Case wdListOutlineNumbering
                sMsg = sMsg & "wdListOutlineNumbering"
            Case wdListPictureBullet
                sMsg = sMsg & "wdListPictureBullet"
            Case wdListSimpleNumbering
                sMsg = sMsg & "wdListSimpleNumbering"
        End Select
        MsgBox sMsg
    Next oL
End Sub
Sub LoadDropdown()
    Dim bk As Bookmark
    Dim par As Paragraph
    Dim rg As Range
    Dim CC As ContentControl
    
    If ActiveDocument.Bookmarks.Count = 0 Then
        MsgBox "Insert a bookmark around the sentences" & vbCr & _
            "to add to each dropdown and then rerun this macro."
        Exit Sub
    End If
    
    For Each bk In ActiveDocument.Bookmarks
        Set rg = ActiveDocument.Range
        rg.Collapse wdCollapseEnd
        rg.InsertParagraphAfter
        
        Set CC = ActiveDocument.ContentControls.Add( _
            Type:=wdContentControlComboBox, Range:=rg)
        
        For Each par In bk.Range.Paragraphs
            Set rg = par.Range
            rg.MoveEnd Unit:=wdCharacter, Count:=-1 ' don't include paragraph mark
            CC.DropdownListEntries.Add rg.text
        Next par
    Next bk
End Sub
' ReplaceArrows Macro
' Brock Created to replace selected arrows
' Assign to ALT+SHIFT+.
Sub ReplaceArrows()
    With Selection.Find
        .text = "->"
        .Replacement.text = ">"
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub RemoveHeaderNos()
' Remove the header nos
'From https://stackoverflow.com/questions/36405534/removing-heading-numbers-from-docx

    Debug.Print "Removing header numbers and formatting..."
   For Each s In ActiveDocument.Styles
        s.LinkToListTemplate ListTemplate:=Nothing
    Next
End Sub

Sub TestAMacro()
  Dim aHL As
  For Each aHL In ActiveDocument.Hyperlinks
    ActiveDocument.Range.InsertAfter aHL.TextToDisplay & ": " & aHL.Address & vbCr
  Next aHL
End Sub

Sub TestAMacro1()
'
' TestAMacro1 Macro
' another test
'
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="_AO-7552:_CAISO:_Improve", ScreenTip:="", TextToDisplay:= _
        "AO-7552: CAISO: Improve OMS Query Tasks to Support Not Sending Market Participant"
End Sub
  'From https://stackoverflow.com/questions/49328477/converting-headings-to-bookmarks-in-word
  Sub HeadingsToBookmarks()
    Dim heading As Range
    Set heading = ActiveDocument.Range(Start:=0, End:=0)
    Do
        Dim current As Long
        current = heading.Start
        Set heading = heading.GoTo(What:=wdGoToHeading, Which:=wdGoToNext)
        If heading.Start = current Then
            Exit Do
        End If
    'This is the part I changed: ListFormat.ListString
        ActiveDocument.Bookmarks.Add MakeValidBMName(heading.Paragraphs(1).Range.ListFormat.ListString), Range:=heading.Paragraphs(1).Range
        Loop
    End Sub
    
    Function MakeValidBMName(strIn As String)
    Dim pFirstChr As String
    Dim i As Long
    Dim tempStr As String
    strIn = Trim(strIn)
    pFirstChr = Left(strIn, 1)
    If Not pFirstChr Like "[A-Za-z]" Then
        strIn = "Section_" & strIn
    End If
    For i = 1 To Len(strIn)
        Select Case Asc(Mid$(strIn, i, 1))
            Case 49 To 58, 65 To 90, 97 To 122
                tempStr = tempStr & Mid$(strIn, i, 1)
            Case Else
                tempStr = tempStr & "_"
        End Select
    Next i
    tempStr = Replace(tempStr, " ", " ")
    tempStr = Replace(tempStr, ":", "")

    If Right(tempStr, 1) = "_" Then
        tempStr = Left(tempStr, Len(tempStr) - 1)
    End If

    MakeValidBMName = tempStr
 End Function
Sub FindStyleStory2()
'
' FindStyleStory2 Macro
'
'
    Selection.Find.ClearFormatting
    '***THIS IS WHAT I NEED*** Selection.Find.Style = ActiveDocument.Styles("story 2")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Find.Execute
End Sub
Sub FindAndCreate()
'
' FindAndCreate Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("story 2")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="RS791"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("story 2")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="RS962"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With
End Sub

Sub DemoteAllHeadings()
'This VBA actually demotes the headings
'Use this with DoVBRoutineNow to demote headings of all documents in a folder.
'Found at https://www.datanumen.com/blogs/3-ways-batch-promote-demote-heading-levels-word-document/
    Dim p As Paragraph
    Dim sParStyle As String
    Dim iHeadLevel As Integer

    For Each p In ActiveDocument.Paragraphs
        sParStyle = p.Style
        If Left(sParStyle, 7) = "Heading" Then
            iHeadLevel = Val(Mid(sParStyle, 8)) + 1
            If iHeadLevel > 9 Then iHeadLevel = 9
            p.Style = "Heading " & iHeadLevel
        End If
    Next p
End Sub

Sub DoVBRoutineNow()
'This allows you to loop through all the documents in the path specified below.
'Found at https://stackoverflow.com/questions/11526577/loop-through-all-word-files-in-directory
'It calls DemoteAllHeadings to demote all the headings.
Dim file
Dim path As String

path = "C:\tempdocs\"

'This next line mentions DOC but it still picked up DOCX... cool
file = Dir(path & "*.doc")
Do While file <> ""
Documents.Open FileName:=path & file

Call DemoteAllHeadings

ActiveDocument.Save
ActiveDocument.Close

file = Dir()
Loop
End Sub

Sub DoVBRoutineNow2()
'Again: This allows you to loop through all the documents in the path specified below.
'Found at https://stackoverflow.com/questions/11526577/loop-through-all-word-files-in-directory
'It calls EnterFilename to enter the filename as a heading.
Dim file
Dim path As String


path = "C:\tempdocs\"

file = Dir(path & "*.doc")
Do While file <> ""
Documents.Open FileName:=path & file

Call InsertFileNameOnly

ActiveDocument.Save
ActiveDocument.Close

file = Dir()
Loop
End Sub
Sub InsertFileNameOnly()
'https://www.extendoffice.com/documents/word/5412-insert-filename-without-extension-in-word.html
    Dim xPathName As String
    Dim xDotPos As Integer
    With Application.ActiveDocument
        If Len(.path) = 0 Then .Save
        xDotPos = VBA.InStrRev(.Name, ".")
        xPathName = VBA.Left(.Name, xDotPos - 1)
    End With
    Application.Selection.Style = ActiveDocument.Styles("Heading 1")
    Application.Selection.TypeText xPathName
    'Selection.TypeParagraph
End Sub

Sub MergeMultiDocsIntoOne()
'From https://www.datanumen.com/blogs/2-ways-quickly-merge-multiple-word-documents-one-via-vba/
  Dim dlgFile As FileDialog
  Dim nTotalFiles As Integer
  Dim nEachSelectedFile As Integer
    
  Set dlgFile = Application.FileDialog(msoFileDialogFilePicker)
 
  With dlgFile
    .AllowMultiSelect = True
    If .Show <> -1 Then
      Exit Sub
    Else
      nTotalFiles = .SelectedItems.Count
    End If
  End With
 
  For nEachSelectedFile = 1 To nTotalFiles
    Selection.InsertFile dlgFile.SelectedItems.Item(nEachSelectedFile)
    If nEachSelectedFile < nTotalFiles Then
      Selection.InsertBreak Type:=wdPageBreak
    Else
      If nEachSelectedFile = nTotalFiles Then
        Exit Sub
      End If
    End If
  Next nEachSelectedFile
End Sub
Sub EscapeSpecialChar()
'This code will escape these characters: % | { @ # *

    ActiveDocument.Select
    
'    With Selection.Find
'        .ClearFormatting
'        .text = "]"
'        With .Replacement
'            .ClearFormatting
'            .text = "U+0005D"
'        End With
'        .Execute Replace:=wdReplaceAll
'    End With
    
'    With Selection.Find
'        .ClearFormatting
'        .text = "["
'        With .Replacement
'            .ClearFormatting
'            .text = "U+0005B"
'        End With
'        .Execute Replace:=wdReplaceAll
'    End With
    
    With Selection.Find
        .ClearFormatting
        .text = "%"
        With .Replacement
            .ClearFormatting
            .text = "\%"
        End With
        .Execute Replace:=wdReplaceAll
    End With
    
    ActiveDocument.Select
    
    With Selection.Find
        .ClearFormatting
        .text = "|"
        With .Replacement
            .ClearFormatting
            .text = "\|"
        End With
        .Execute Replace:=wdReplaceAll
    End With
    
 '   With Selection.Find
 '       .ClearFormatting
 '       .text = ":"
 '       With .Replacement
 '           .ClearFormatting
 '           .text = "&#58;"
 '       End With
 '       .Execute Replace:=wdReplaceAll
 '   End With
    
    With Selection.Find
        .ClearFormatting
        .text = "{"
        With .Replacement
            .ClearFormatting
            .text = "\{"
        End With
        .Execute Replace:=wdReplaceAll
    End With
    
    With Selection.Find
        .ClearFormatting
        .text = "@"
        With .Replacement
            .ClearFormatting
            .text = "\@"
        End With
        .Execute Replace:=wdReplaceAll
    End With
    
    With Selection.Find
        .ClearFormatting
        .text = "#"
        With .Replacement
            .ClearFormatting
            .text = "&#35;"
        End With
        .Execute Replace:=wdReplaceAll
    End With
    
    With Selection.Find
        .ClearFormatting
        .text = "*"
        With .Replacement
            .ClearFormatting
            .text = "&#42;"
        End With
        .Execute Replace:=wdReplaceAll
    End With

End Sub
Sub FixCaps()
'This macro makes all Level 1 Headings use title case.
'I found this macro here: http://editorium.com/archive/title-case-macro/
ActiveDocument.Select
Selection.Find.ClearFormatting
Selection.Find.Style = ActiveDocument.Styles("Heading 1")
With Selection.Find
.text = ""
.Replacement.text = ""
.Forward = True
.Wrap = wdFindContinue
.Format = True
.MatchCase = False
.MatchWholeWord = False
.MatchWildcards = False
.MatchSoundsLike = False
.MatchAllWordForms = False
End With
Selection.Find.Execute
While Selection.Find.Found = True
Selection.Range.Case = wdTitleWord
Selection.MoveRight Unit:=wdCharacter, Count:=1
Selection.Find.Execute
Wend
End Sub

Sub RunMacroOnAllFilesInFolder()
'I found this at https://www.quora.com/How-do-I-automatically-run-a-macro-on-all-Word-files-docx-in-a-folder#
'It works along with the MyMacro sub below to process all documents in a folder.
    Dim flpath As String, fl As String
    flpath = InputBox("Please enter the path to the folder you want to run the macro on.")
    If flpath = "" Then Exit Sub
    If Right(flpath, 1) <> Application.PathSeparator Then flpath = flpath & Application.PathSeparator
    fl = Dir(flpath & "*.docx")
    Application.ScreenUpdating = False
    Do Until fl = ""
        MyMacro flpath, fl
        fl = Dir
    Loop
    
    'This part is to open a blank Word document. Found here: https://word.tips.net/T000822_Creating_a_New_Document_in_VBA.html
    Documents.Add
    
    Call DoVBRoutineNow
    Call DoVBRoutineNow2
    Call MergeMultiDocsIntoOne
    Call FixCaps
End Sub
Sub MyMacro(flpath As String, fl As String)
    Dim doc As Document
    Set doc = Documents.Open(flpath & fl)
    'Do stuff
    Call RemoveHeaderNos
    Call ConvertNumbersToText
    Call EscapeSpecialChar
    'Call FixCaps -This needs to be run later, after the docs are merged into one
    doc.Save
    doc.Close SaveChanges:=False
End Sub

