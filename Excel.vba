Sub Prepare4KnowBase()
'Sub to run that combines all functions and processes Column A

Call Delete1stThreeLastAndImage
Call NewDeleteIssueLinks  'Changed to add IssueType column
Call ReplaceHashNumbers  'Changed to add IssueType column
Call LabelDescriptionConfigurationAndTestSteps  'Changed to add IssueType column
Call MergeColumns  'Changed to add IssueType column

Columns("B").Select  'Changed to add IssueType column
Selection.Hyperlinks.Delete
Selection.VerticalAlignment = xlTop

End Sub
Sub FinishPrepare4PDF()
    'Delete first row
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp

    'Merge all columns
    Dim lr As Long, i As Long
    lr = Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 1 To lr
    If Range("A" & i) <> "" Then
        Range("A" & i) = Range("A" & i).Text & "  " & Range("B" & i).Text & Chr(10) & Range("C" & i).Text & Chr(10) & Range("D" & i).Text
    End If
    Next i
    
    Columns("B:D").EntireColumn.Delete

End Sub

Function MergeColumns()
'Adapted from https://www.mrexcel.com/board/threads/vba-merge-two-columns.896609/
    Dim lr As Long, i As Long
    lr = Range("D" & Rows.Count).End(xlUp).Row
    
    For i = 1 To lr
    If Range("A" & i) <> "" Then
        Range("E" & i) = Range("E" & i).Text & Range("F" & i).Text & Range("G" & i).Text
    End If
    Next i
    
    Columns("F:G").EntireColumn.Delete
    
    Range("E1").Value = "Summary . Configuration Steps . Test Steps"
    
End Function
Function DeleteIssueLinks()
'
' It removes all linked issues that aren't "PCI-XXXXX".
' TO DO: Code to ONLY include "PCI-XXXXX" instead of deleting everything else seperately
'
    Columns("D").Select
    
    Selection.Replace What:="PCIBPA-*,", Replacement:=""
    
    Selection.Replace What:="CE-*,", Replacement:=""

    Selection.Replace What:="EA-*,", Replacement:=""
    Selection.Replace What:="EA-????", Replacement:=""
        
    Selection.Replace What:="RS-*,", Replacement:=""
        
    Selection.Replace What:="ET-*,", Replacement:=""
    
    Selection.Replace What:="GP-*,", Replacement:=""
    Selection.Replace What:="GP-?????", Replacement:=""
    
    Selection.Replace What:="GTD-????", Replacement:=""
    
    Selection.Replace What:="RM-*,", Replacement:=""
    Selection.Replace What:="RM-???", Replacement:=""
        
    Selection.Replace What:="EI-*,", Replacement:=""
    
    Selection.Replace What:="ER-*,", Replacement:=""
        
    Selection.Replace What:="FL-*,", Replacement:=""
        
    Selection.Replace What:="MX-*,", Replacement:=""
    
    Selection.Replace What:="AO-*,", Replacement:=""
    Selection.Replace What:="AO-????", Replacement:=""

    Selection.Replace What:="EO-*,", Replacement:=""
    Selection.Replace What:="EO-????", Replacement:=""
    
    Selection.Replace What:="NE-*,", Replacement:=""
    Selection.Replace What:="NE-????,", Replacement:=""
    
    Selection.Replace What:="SP-*,", Replacement:=""

    Selection.Replace What:="MI-*,", Replacement:=""

    Selection.Replace What:="LG-*,", Replacement:=""
    
    Selection.Replace What:="DA-*,", Replacement:=""
    Selection.Replace What:="DA-???,", Replacement:=""
    Selection.Replace What:="DA-????,", Replacement:=""
    
    Selection.Replace What:="EP-*,", Replacement:=""
    
    Selection.Replace What:="IA-*,", Replacement:=""
    
    Selection.Replace What:="TS-*,", Replacement:=""
    
End Function

Function DeleteBrackets()
'
' This macro relies on the user highlighting the "Summary" column first.
' It removes open and closing brackets to avoid Confluence creating dead-end hyperlinks from the enclosed text.
' TO DO: Research alternatives. This is a clumsy hack because it could affect text that should be in brackets. What about just converting to greater/less than?
'
    Selection.Replace What:="[", Replacement:=""
    Selection.Replace What:="]", Replacement:=":"
    
End Function

Function InsertMarkup()
'
'This macro converts the spreadsheet to markup so it can be Markup imported into Confluence.
'It uses pipes that will be converted into tables and code, encasing content, to be included in expandable sections.
'
    Range("A1").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C1").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G1").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("I1").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("I1").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    
'Get the last row number
    Dim lastRow As Long
    lastRow = Cells.Find(What:="*", _
        After:=Range("A1"), _
        LookAt:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False).Row

'Use the lastRow variable to create columns
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "|"
    Range("A1:A" & lastRow).Select
    Selection.FillDown

Application.Wait Now + TimeValue("00:00:005")

    Range("C1").Select
    ActiveCell.FormulaR1C1 = "|"
    Range("C1:C" & lastRow).Select
    Selection.FillDown
    
Application.Wait Now + TimeValue("00:00:005")
    
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "|"
    Range("E1:E" & lastRow).Select
    Selection.FillDown
    
Application.Wait Now + TimeValue("00:00:005")
    
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "|"
    Range("G1:G" & lastRow).Select
    Selection.FillDown

Application.Wait Now + TimeValue("00:00:005")
    
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "|"
    Range("I1:I" & lastRow).Select
    Selection.FillDown
    
Application.Wait Now + TimeValue("00:00:005")
    
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "{expand:title=Click to Expand}"
    Range("J2:J" & lastRow).Select
    Selection.FillDown
    
Application.Wait Now + TimeValue("00:00:005")
    
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "{expand}"
    Range("N2:N" & lastRow).Select
    Selection.FillDown

Application.Wait Now + TimeValue("00:00:005")
    
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "|"
    Range("O1:O" & lastRow).Select
    Selection.FillDown
    
End Function

Function ReplaceHashNumbers()

'This macro looks for hash marks at the beginning of test steps (in Test Steps column) and converts them to incrementing numbers.
'It is necessary because the Jira Excel export turns auto-numbering into hash marks.
'-----
'https://www.mrexcel.com/board/threads/vba-code-for-a-regex-find-and-replace.986704/
'Helpul: https://stackoverflow.com/questions/16084909/vba-multiple-matches-within-one-string-using-regular-expressions-execute-method
'Had to set Tools > References > Microsoft VBScript Regular Expressions 5.5

'Set the variables
Dim RegEx1a As Object, RegEx1b As Object, RegEx2 As Object
Dim r As Range, rC As Range
Dim NumberCounter As Integer
    
'Create first search pattern with global (for count)
Set RegEx1a = CreateObject("VBScript.RegExp")
RegEx1a.Pattern = "\n# "
RegEx1a.Global = True '<-- set flag to true to replace all occurences of match
    
'Duplicate first search pattern without global (for actual replace)
Set RegEx1b = CreateObject("VBScript.RegExp")
RegEx1b.Pattern = "\n# "

'Create second search pattern
Set RegEx2 = CreateObject("VBScript.RegExp")
RegEx2.Pattern = "^# "
RegEx2.Global = True '<-- set flag to true to replace all occurences of match

' Cells in column F
Set r = Range("G1", Cells(Rows.Count, "G").End(xlUp))

' Loop through the cells in column G and execute regex replace
    For Each rC In r
        NumberCounter = 2
        If rC.Value <> "" Then rC.Value = RegEx2.Replace(rC.Value, "1. ")
            Set MyMatches = RegEx1a.Execute(rC.Value)
            Debug.Print "myMatchCt: " & MyMatches.Count
            For i = 1 To MyMatches.Count
                If rC.Value <> "" Then rC.Value = RegEx1b.Replace(rC.Value, Chr(10) & NumberCounter & ". ")
                NumberCounter = NumberCounter + 1
            Next i
    Next rC

End Function
Function Delete1stThreeLastAndImage()
'Delete1stThreeLastAndImage Macro
    
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Selection.Delete
    Rows("1:3").Select
    Selection.Delete Shift:=xlUp
    
    Dim lastRow As Long
        lastRow = Cells.Find(What:="*", _
            After:=Range("A1"), _
            LookAt:=xlPart, _
            LookIn:=xlFormulas, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False).Row
       
        Rows(lastRow).Select
        Selection.Delete Shift:=xlUp

End Function

Function LabelDescriptionConfigurationAndTestSteps()

Dim c As Range
Dim lastRow As Long
    lastRow = Cells.Find(What:="*", _
        After:=Range("A1"), _
        LookAt:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False).Row

    Range("E2:E" & lastRow).Select
    For Each c In Selection
        If c.Value <> "" Then c.Value = "DESCRIPTION: " & c.Value
    Next
    
    Range("F2:F" & lastRow).Select
    For Each c In Selection
        If c.Value <> "" Then c.Value = Chr(10) & "---" & Chr(10) & "CONFIGURATION STEPS: " & Chr(10) & c.Value
    Next
    
    Range("G2:G" & lastRow).Select
    For Each c In Selection
        If c.Value <> "" Then c.Value = Chr(10) & "---" & Chr(10) & "TEST STEPS: " & Chr(10) & c.Value
    Next

End Function
                                                                                                                
Function NewDeleteIssueLinks()

'This macro looks for values that begin with "PCI-" and keeps those.
'It deletes everything else.  Need to maybe RegEx available in Excel.
'-----
'Based on the ReplaceHashNumbers Function

'Set the variables
Dim RegEx2 As Object
Dim r As Range, rC As Range
Dim NumberCounter As Integer
    
'Create first search pattern with global (for count)
'Help was here https://stackoverflow.com/questions/7124778/how-to-match-anything-up-until-this-sequence-of-characters-in-a-regular-expres
'Use this Regex to find all the ticket numbers and count them

'Create second search pattern
Set RegEx2 = CreateObject("VBScript.RegExp")
'Takes time to run, maybe the following pattern can be reduced to lessen run time, try testing at https://regex101.com/
RegEx2.Pattern = "(?= PCI-) | (.+?-.+?,|.+?-.+?$)|(?!^PCI-)(^.+?-.+?,)|(?!^PCI-)(^.+?-.+?$)"
RegEx2.Global = True '<-- set flag to true to replace all occurences of match

' Cells in column F
Set r = Range("D1", Cells(Rows.Count, "D").End(xlUp))

' Loop through the cells in column G and execute regex replace
    For Each rC In r
        NumberCounter = 2
        If rC.Value <> "" Then rC.Value = RegEx2.Replace(rC.Value, "")
    Next rC

End Function
