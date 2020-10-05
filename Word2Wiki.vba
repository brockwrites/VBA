
Sub Word2Wiki()
    
    Application.ScreenUpdating = False
    
    ConvertH1
    ConvertH2
    ConvertH3
    
    ConvertItalic
    ConvertBold
    ConvertUnderline
    
    ConvertLists
    ConvertTables
    
    EscapeSpecialChar
    
    ' Copy to clipboard
    ActiveDocument.Content.Copy
    
    Application.ScreenUpdating = True
End Sub

Private Sub ConvertH1()
    Dim normalStyle As Style
    Set normalStyle = ActiveDocument.Styles(wdStyleNormal)
    
    ActiveDocument.Select
    
    With Selection.Find
    
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleHeading1)
        .text = ""
        
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Forward = True
        .Wrap = wdFindContinue
        
        Do While .Execute
            With Selection
                If InStr(1, .text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .text = vbCr Then
                    .InsertBefore "!"
                Else
                    .MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                End If
                .Style = normalStyle
            End With
        Loop
    End With
End Sub

Private Sub ConvertH2()
    Dim normalStyle As Style
    Set normalStyle = ActiveDocument.Styles(wdStyleNormal)
    
    ActiveDocument.Select
    
    With Selection.Find
    
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleHeading2)
        .text = ""
        
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Forward = True
        .Wrap = wdFindContinue
        
        Do While .Execute
            With Selection
                If InStr(1, .text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .text = vbCr Then
                    .InsertBefore "!!"
                Else
                    .MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                End If
                .Style = normalStyle
            End With
        Loop
    End With
End Sub

Private Sub ConvertH3()
    Dim normalStyle As Style
    Set normalStyle = ActiveDocument.Styles(wdStyleNormal)
    
    ActiveDocument.Select
    
    With Selection.Find
    
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleHeading3)
        .text = ""
        
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Forward = True
        .Wrap = wdFindContinue
        
        Do While .Execute
            With Selection
                If InStr(1, .text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .text = vbCr Then
                    .InsertBefore "!!!"
                Else
                    .MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                End If
                .Style = normalStyle
            End With
        Loop
    End With
End Sub

Private Sub ConvertBold()
    ActiveDocument.Select
    
    With Selection.Find
    
        .ClearFormatting
        .Font.Bold = True
        .text = ""
        
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Forward = True
        .Wrap = wdFindContinue
        
        Do While .Execute
            With Selection
                If InStr(1, .text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .text = vbCr Then
                    .InsertBefore "__"
                    .InsertAfter "__"
                Else
                    .MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                End If
                
                .Font.Bold = False
            End With
        Loop
    End With
End Sub

Private Sub ConvertItalic()
    ActiveDocument.Select
    
    With Selection.Find
    
        .ClearFormatting
        .Font.Italic = True
        .text = ""
        
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Forward = True
        .Wrap = wdFindContinue
        
        Do While .Execute
            With Selection
                If InStr(1, .text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .text = vbCr Then
                    .InsertBefore "''"
                    .InsertAfter "''"
                Else
                    .MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                End If
                .Font.Italic = False
            End With
        Loop
    End With
End Sub

Private Sub ConvertUnderline()
    ActiveDocument.Select
    
    With Selection.Find
    
        .ClearFormatting
        .Font.Underline = True
        .text = ""
        
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Forward = True
        .Wrap = wdFindContinue
        
        Do While .Execute
            With Selection
                If InStr(1, .text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .text = vbCr Then
                    .InsertBefore "==="
                    .InsertAfter "==="
                Else
                    .MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
                End If
                .Font.Underline = False
            End With
        Loop
    End With
End Sub

Private Sub ConvertLists()
    Dim para As Paragraph
    For Each para In ActiveDocument.ListParagraphs
        With para.Range
            If .ListFormat.ListType = wdListBullet Then
                .InsertBefore String(.ListFormat.ListLevelNumber, "*")
            Else
                .InsertBefore String(.ListFormat.ListLevelNumber, "#")
            End If
            
            .ListFormat.RemoveNumbers
        End With
    Next para
End Sub

Private Sub ConvertTables()
    Dim thisTable As Table
    Dim thisRow As Row
    Dim ElRango As Object
    For Each thisTable In ActiveDocument.Tables
        With thisTable
            For Each thisRow In .Rows
                thisRow.Range.InsertBefore "||"
                If thisRow.Index = .Rows.Count Then
                    'Cerramos la tabla al final
                    thisRow.Range.InsertAfter "||"
                End If
            Next thisRow
            Set ElRango = .ConvertToText(Separator:="|")
            With ElRango.Find
                .ClearFormatting
                .text = "^p"
                With .Replacement
                    .ClearFormatting
                    .text = ""
                End With
                .Execute Replace:=wdReplaceAll
            End With
        End With
    Next thisTable
End Sub

Private Sub EscapeSpecialChar()

    ActiveDocument.Select
    
    With Selection.Find
        .ClearFormatting
        .text = "%"
        With .Replacement
            .ClearFormatting
            .text = "~np~%~/np~"
        End With
        .Execute Replace:=wdReplaceAll
    End With

End Sub
