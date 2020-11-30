Sub DeleteIssueLinks()
'
' DeleteIssueLinks Macro
'
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
        
    Selection.Replace What:="ER-*,", Replacement:=""
        
    Selection.Replace What:="FL-*,", Replacement:=""
        
    Selection.Replace What:="MX-*,", Replacement:=""
    
    Selection.Replace What:="AO-*,", Replacement:=""
    Selection.Replace What:="AO-????", Replacement:=""

    Selection.Replace What:="EO-*,", Replacement:=""
    Selection.Replace What:="EO-????", Replacement:=""
    
    Selection.Replace What:="NE-*,", Replacement:=""
    
    Selection.Replace What:="SP-*,", Replacement:=""

    Selection.Replace What:="MI-*,", Replacement:=""

    Selection.Replace What:="LG-*,", Replacement:=""
    
    Selection.Replace What:="DA-*,", Replacement:=""
    
    Selection.Replace What:="EP-*,", Replacement:=""
    
    Selection.Replace What:="IA-*,", Replacement:=""
    
    Selection.Replace What:="TS-*,", Replacement:=""
    
End Sub

Sub DeleteBrackets()
'
' DeleteBrackets Macro
'
    Selection.Replace What:="[", Replacement:=""
    Selection.Replace What:="]", Replacement:=":"
    
End Sub

Sub InsertMarkup()
'Insert blank columns
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
    
'Get the last row number from the user
    Dim LastRowNumber As Integer
    LastRowNumber = InputBox("What is the Last Row Number?", "Row Number")

'Use the LastRowNumber variable to create columns
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "|"
    Range("A1:A" & LastRowNumber).Select
    Selection.FillDown

Application.Wait Now + TimeValue("00:00:01")

    Range("C1").Select
    ActiveCell.FormulaR1C1 = "|"
    Range("C1:C" & LastRowNumber).Select
    Selection.FillDown
    
Application.Wait Now + TimeValue("00:00:01")
    
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "|"
    Range("E1:E" & LastRowNumber).Select
    Selection.FillDown
    
Application.Wait Now + TimeValue("00:00:01")
    
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "|"
    Range("G1:G" & LastRowNumber).Select
    Selection.FillDown

Application.Wait Now + TimeValue("00:00:01")
    
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "|"
    Range("I1:I" & LastRowNumber).Select
    Selection.FillDown
    
Application.Wait Now + TimeValue("00:00:01")
    
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "{expand:title=Click to Expand}"
    Range("J2:J" & LastRowNumber).Select
    Selection.FillDown
    
Application.Wait Now + TimeValue("00:00:01")
    
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "{expand}"
    Range("N2:N" & LastRowNumber).Select
    Selection.FillDown

Application.Wait Now + TimeValue("00:00:01")
    
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "|"
    Range("O1:O" & LastRowNumber).Select
    Selection.FillDown
    
End Sub

Sub ReplaceHashNumbers()

    'https://www.mrexcel.com/board/threads/vba-code-for-a-regex-find-and-replace.986704/
    'Helpul: https://stackoverflow.com/questions/16084909/vba-multiple-matches-within-one-string-using-regular-expressions-execute-method

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
    Set r = Range("F1", Cells(Rows.Count, "F").End(xlUp))

    ' Loop through the cells in column F and execute regex replace
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

End Sub
