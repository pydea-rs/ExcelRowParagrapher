Attribute VB_Name = "Proceedings470"

Sub SaveToDocx(Text As String)
    Dim objWord
    Dim objDoc

    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Add
    objWord.Visible = True
    
    With objWord.Selection
        .TypeText (Text)
        .WholeStory
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Collapse
        .Find.ClearFormatting
        .Find.Font.Color = wdColorYellow
        .Find.Replacement.ClearFormatting
        .Find.Replacement.Font.Color = wdColorAuto
        .Find.Replacement.Font.Size = 8

        ' all this running, just for making font change for persian text as well!
        With .Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchKashida = False
            .MatchDiacritics = False
            .MatchAlefHamza = False
            .MatchControl = False
            .MatchByte = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        
        .Find.Execute Replace:=wdReplaceAll
            
         .WholeStory
        .Font.Color = wdColorAutomatic
        .Collapse
    End With
    
    ' make text right to left
    Dim Paragraph
    Dim AllParagraphs
    Set AllParagraphs = objDoc.Paragraphs
    
    For i = 1 To AllParagraphs.Count '''Iterate over all paragraphs
        Set Paragraph = AllParagraphs(i)
        Paragraph.ReadingOrder = xlRTwdReadingOrderRtl '''set text direction(aka Reading Order)
        Paragraph.Range.ParagraphFormat.Alignment = 3
    Next
    
    MsgBox "ﬁ«·» »« „Ê›ﬁÌ  «⁄„«· ‘œ!", vbOKOnly + vbInformation
    
    objDoc.Save
    objDoc.Close
    
End Sub

Sub Proceedings470()
    Application.ScreenUpdating = False
    Dim TemplateText As Object
    Set TemplateText = CreateObject("Scripting.Dictionary")
        TemplateText.Add "Female Name", "- ’Ê— Ã·”Â 470:  œ—ŒÊ«”  Œ«‰„ "
        TemplateText.Add "Male Name", "- ’Ê— Ã·”Â 470:  œ—ŒÊ«”  ¬ﬁ«Ì "
        TemplateText.Add "Course", " œ«‰‘ÃÊÌ œÊ—Â "
        TemplateText.Add "Grade", " „ﬁÿ⁄ "
        TemplateText.Add "Field", " —‘ Â "
        TemplateText.Add "Entrance", " Ê—ÊœÌ "
        TemplateText.Add "Price Tag", " „ÿ—Õ ‘œ Ê »«  ÊÃÂ »Â œ·«Ì· –ò— ‘œÂ „Ê«›ﬁ  ê—œÌœ »— «”«” „«œÂ 2 ‘ÌÊÂ ‰«„Â «›“«Ì‘ ”‰Ê«  „«“«œ »— „œ  „Ã«“  Õ’Ì· œ«‰‘ÃÊÌ«‰ ò«—‘‰«”Ì «—‘œ Ê œò —Ì „’Ê» 19/10/1389 ÂÌ√  —∆Ì”Â œ«‰‘ê«Â° ›ﬁÿ „»·€ "
        TemplateText.Add "No More Commision", " —Ì«·  Ê”ÿ ‰«„»—œÂ Å—œ«Œ  ‘Êœ. »œÌÂÌ «”  Å—Ê‰œÂ œ«‰‘ÃÊÌ „–òÊ— ﬁ«»· ÿ—Õ „Ãœœ œ— ò„Ì”ÌÊ‰ ‰„Ì »«‘œ."
    
    Dim Cell As Range
    Dim i As Integer
    Dim Paragraph As String
    Dim FullText As String
    FullText = ""
    ' Loop through each row column to create the text for each person
    Dim FirstRow As Integer
    Dim LastRow As Integer
    FirstRow = 3
    LastRow = 20
    Dim Values() As String
    Dim Cost As Currency
    For i = FirstRow To LastRow
        If Range("C" & i).Value = 0 Then
            Paragraph = TemplateText("Male Name")
        Else
            Paragraph = TemplateText("Female Name")
        End If
        
        Paragraph = Paragraph & Range("B" & i).Value   ' Name
        Values = Split(Range("G" & i).Value, "/")
        Paragraph = Paragraph & TemplateText("Course") _
                & Values(1) & TemplateText("Grade") & Values(0) _
                & TemplateText("Entrance") & Left(Range("D" & i).Value, 2) _
                & TemplateText("Field") & Range("F" & i).Value _
                & TemplateText("Price Tag")
                Cost = Range("J" & i).Value
                Paragraph = Paragraph & Cost & TemplateText("No More Commision")
        ' type the result
        FullText = FullText & vbNewLine & Paragraph
    Next i
    
    Call SaveToDocx(FullText)
    Application.ScreenUpdating = True
End Sub



