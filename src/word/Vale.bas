Attribute VB_Name = "Vale"
'Option Explicit
Option Private Module

Sub finish()

    If Not Library.propertyExists("NumeroCliente") Or _
    Not Library.propertyExists("Titulo1") Or _
    Not Library.propertyExists("Titulo2") Or _
    Not Library.propertyExists("Titulo3") Or _
    Not Library.propertyExists("Titulo4") Or _
    Not Library.propertyExists("Titulo5") Or _
    Not Library.propertyExists("NumeroNosso") Or _
    Not Library.propertyExists("NumeroCliente") Or _
    Not Library.propertyExists("Revisao") Or _
    Not Library.propertyExists("Projeto") Then
        MsgBox "Um ou todos os campos do documento não estão definidos (Titulo1,Titulo2, Cliente, Projetom, Revisão, etc.)"
        Exit Sub
    End If
                       
    Call setupPage
    Library.Sleep 500
    Call setupStyles
    Library.Sleep 500
    Call addHeader
    Library.Sleep 500
    Call addRevisionTable
    Library.Sleep 500
    Call updateClientText
    Library.Sleep 500
    Call addTOC
End Sub

Sub updateClientText()
    ' Update the style to client's
    
    ' Set the document object to the active document
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim DePara As Variant
    DePara = Array( _
        Array("Normal", "Parágrafo Normal_VALE_"), _
        Array("Heading 1", "Título I_VALE_"), _
        Array("Heading 2", "Título I.I_VALE_"), _
        Array("Heading 3", "Título I.I.I_VALE_"), _
        Array("Heading 4", "Título I.I.I.I_VALE_") _
    )
    
    ' Loop through all paragraphs in the document
    Dim rng As Range
    Set rng = doc.Content
    For Each style In DePara
        ' Find and replace style
        With rng.Find
            .ClearFormatting
            .style = style(0)
            .Replacement.ClearFormatting
            .Replacement.style = style(1)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .Execute Replace:=wdReplaceAll
        End With
    Next style
    
    ' Loop through all tables starting from the second table (index 2)
    Dim tbl As Table
    Dim cel As cell
    Dim i As Integer
    For i = 2 To doc.Tables.Count
        Set tbl = doc.Tables(i)
        For Each cel In tbl.Range.Cells
            Set rng = cel.Range
            ' Remove the end of cell marker from the range
            rng.End = rng.End - 1
            rng.style = "Tabela Normal_VALE_" ' Replace with your desired style name
        Next cel
    Next i
    
    ' Loop through all lists in the document
    Dim para As Paragraph
    Dim ListLevel As Integer
    For Each para In doc.Paragraphs
        ' Check if the paragraph is part of a list
        If para.Range.ListFormat.listType = wdListBullet Then
            ListLevel = para.Range.ListFormat.ListLevelNumber
            ' Apply the custom list template to the bullet list
            para.Range.style = "Lista Texto"
            para.Range.ListFormat.ListLevelNumber = ListLevel
        End If
    Next para
End Sub

Sub addTOC()
    Dim tocRange As Range
    
    ' Check if there is at least one TOC in the document
    If ActiveDocument.TablesOfContents.Count > 0 Then
        Dim startPos As Long
        startPos = ActiveDocument.TablesOfContents(1).Range.Start
        
        Dim endPos As Long
        endPos = ActiveDocument.TablesOfContents(1).Range.End
        
        ActiveDocument.TablesOfContents(1).Delete
        
        ' Find the paragraph before the TOC start position
        Dim paraBeforeTOC As Paragraph
        Set paraBeforeTOC = ActiveDocument.Paragraphs(ActiveDocument.Range(0, startPos).Paragraphs.Count)
        
        ' Find the paragraph before the previous paragraph
        If Not paraBeforeTOC.Previous Is Nothing Then
            Set paraBeforeTOC = paraBeforeTOC.Previous.Previous.Previous
            ActiveDocument.Range(paraBeforeTOC.Range.Start, startPos).Delete
        Else
            MsgBox "There is no content before the TOC to delete."
            Exit Sub
        End If
        
        
        ' Define the range where the TOC will be inserted
        Set tocRange = ActiveDocument.Range(Start:=startPos, End:=endPos)
        
    Else
        ' Define the range where the TOC will be inserted
        Set tocRange = ActiveDocument.Range(Start:=ActiveDocument.Tables(1).Range.End, End:=ActiveDocument.Tables(1).Range.End)
    End If
    
    ' Add the Table of Contents to the defined range
    Dim TOC As TableOfContents
    Set TOC = ActiveDocument.TablesOfContents.Add( _
        Range:=tocRange, _
        UseHeadingStyles:=False, _
        UpperHeadingLevel:=1, _
        LowerHeadingLevel:=2, _
        UseFields:=False, _
        IncludePageNumbers:=True, _
        RightAlignPageNumbers:=True, _
        UseHyperlinks:=True, _
        HidePageNumbersInWeb:=False _
    )
    
    ' Customize TOC entries
    With TOC
        .TabLeader = wdTabLeaderSpaces
        tocRange.InsertBreak Type:=wdPageBreak
        tocRange.InsertBefore vbCrLf & "ÍNDICE" & vbCrLf & vbCrLf
        tocRange.Paragraphs(2).style = "12ptCenterBoldUnderline"
        tocRange.Paragraphs(1).style = "12ptCenterBoldUnderline"
        tocRange.InsertAfter "ITEM" & vbTab & "DESCRIÇÃO" & vbTab & "PÁGINA"
        .Update
    End With
End Sub

Sub addRevisionTable()
    Dim MyRange As Range
    Set MyRange = ActiveDocument.Content
    MyRange.Collapse Direction:=wdCollapseStart
    
    If ActiveDocument.Tables.Count > 0 Then
        ActiveDocument.Tables(1).Delete
    End If
    
    Dim revisionTable As Table
    Set revisionTable = MyRange.Tables.Add( _
        Range:=MyRange, _
        NumRows:=22, _
        NumColumns:=8, _
        DefaultTableBehavior:=wdWord9TableBehavior, _
        AutoFitBehavior:=wdAutoFitFixed _
    )
    
    With revisionTable
        .Borders.Enable = True
        
        Dim i As Integer
        Dim cell As cell
        
        For i = 1 To .Rows.Count
            For Each cell In .Rows(i).Cells
                cell.VerticalAlignment = wdAlignVerticalCenter
                cell.Range.style = "12ptCenter"
            Next cell
        Next i
        
        .Rows.SetHeight RowHeight:=CentimetersToPoints(1.05), HeightRule:=wdRowHeightExactly
        .Columns(1).width = Application.CentimetersToPoints(1.5)
        .Columns(2).width = Application.CentimetersToPoints(1.5)
        .Columns(3).width = Application.CentimetersToPoints(6.5)
        .Columns(4).width = Application.CentimetersToPoints(1.5)
        .Columns(5).width = Application.CentimetersToPoints(1.5)
        .Columns(6).width = Application.CentimetersToPoints(1.5)
        .Columns(7).width = Application.CentimetersToPoints(1.5)
        .Columns(8).width = Application.CentimetersToPoints(2)
        
        .Rows(1).Height = Application.CentimetersToPoints(0.65)
        .Rows(2).Height = Application.CentimetersToPoints(0.5)
        .Rows(3).Height = Application.CentimetersToPoints(0.5)
        .Rows(4).Height = Application.CentimetersToPoints(0.65)
        
        .cell(1, 1).Merge MergeTo:=.cell(1, 8)
        .cell(2, 1).Merge MergeTo:=.cell(2, 8)
        .cell(3, 1).Merge MergeTo:=.cell(3, 8)
        
        .cell(2, 1).Split NumRows:=1, NumColumns:=5
        .cell(3, 1).Split NumRows:=1, NumColumns:=5
        
        ' Add text and styles to cells
        insertCellText .cell(1, 1), "REVISÕES", "10ptCenterBold"
        insertCellText .cell(2, 1), "TE: TIPO", "7ptLeft", 2.5, wdBorderRight, wdBorderBottom
        insertCellText .cell(3, 1), "EMISSÃO", "7ptLeft", 2.5, wdBorderRight, wdBorderTop
        insertCellText .cell(2, 2), "A- PRELIMINAR", "7ptLeft", 3.25, wdBorderRight, wdBorderBottom
        insertCellText .cell(3, 2), "B - PARA APROVAÇÃO", "7ptLeft", 3.25, wdBorderRight, wdBorderTop
        insertCellText .cell(2, 3), "C - PARA CONHECIMENTO", "7ptLeft", 3.75, wdBorderRight, wdBorderBottom
        insertCellText .cell(3, 3), "D - PARA COTAÇÃO", "7ptLeft", 3.75, wdBorderRight, wdBorderTop
        insertCellText .cell(2, 4), "E - PARA CONSTRUÇÃO", "7ptLeft", 3.75, wdBorderRight, wdBorderBottom
        insertCellText .cell(3, 4), "F - CONFORME COMPRADO", "7ptLeft", 3.75, wdBorderRight, wdBorderTop
        insertCellText .cell(2, 5), "G - CONFORME CONSTRUÍDO", "7ptLeft", 4.25, wdBorderBottom
        insertCellText .cell(3, 5), "H - CANCELADO", "7ptLeft", 4.25, wdBorderTop
        insertCellText .cell(4, 1), "Rev.", "12ptCenter"
        insertCellText .cell(4, 2), "TE ", "12ptCenter"
        insertCellText .cell(4, 3), "Descrição", "12ptCenter"
        insertCellText .cell(4, 4), "Por", "12ptCenter"
        insertCellText .cell(4, 5), "Ver.", "12ptCenter"
        insertCellText .cell(4, 6), "Apr.", "12ptCenter"
        insertCellText .cell(4, 7), "Aut.", "12ptCenter"
        insertCellText .cell(4, 8), "Data", "12ptCenter"
        For i = 5 To .Rows.Count
            insertCellText .cell(i, 3), "", "Parágrafo Normal_VALE_"
        Next i
    End With
    
    Set MyRange = revisionTable.Range
    MyRange.Collapse Direction:=wdCollapseEnd
    MyRange.InsertBreak Type:=wdPageBreak
End Sub

Sub addHeader()
    ' Check if the document has a primary header
    If ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Exists Then
        ' Set the header range
        Dim headerRange As Range
        Set headerRange = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range
        
        ' Clear existing content in the header
        headerRange.text = ""
        
        ' Set the header range
        Dim footerRange As Range
        Set footerRange = ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range
        
        ' Clear existing content in the footer
        footerRange.text = ""
        
        Dim headerTable As Table
        Set headerTable = headerRange.Tables.Add( _
            Range:=headerRange, _
            NumRows:=6, _
            NumColumns:=5, _
            DefaultTableBehavior:=wdWord9TableBehavior, _
            AutoFitBehavior:=wdAutoFitFixed _
        )
        
        With headerTable
            .Borders.Enable = True
            
            For i = 1 To .Rows.Count
                With .Rows(i)
                    .Cells.VerticalAlignment = wdAlignVerticalCenter
                End With
            Next i
            
            .Columns(1).width = Application.CentimetersToPoints(4)
            .Columns(2).width = Application.CentimetersToPoints(3)
            .Columns(3).width = Application.CentimetersToPoints(3)
            .Columns(4).width = Application.CentimetersToPoints(5.5)
            .Columns(5).width = Application.CentimetersToPoints(2)
            
            .cell(Row:=1, Column:=4).Merge MergeTo:=.cell(Row:=2, Column:=5)
            .cell(Row:=1, Column:=1).Merge MergeTo:=.cell(Row:=2, Column:=1)
            .cell(Row:=1, Column:=2).Merge MergeTo:=.cell(Row:=2, Column:=2)
            .cell(Row:=3, Column:=1).Merge MergeTo:=.cell(Row:=6, Column:=3)
            
            Call Library.InsertImage(getValeLogoBase64(), .cell(Row:=1, Column:=1).Range)
            .cell(Row:=1, Column:=1).Range.style = "8ptCenter"
            
            MsgBox "Indique o logo da empresa para ser incluído no cabeçalho."
            
            Dim FileInputPath As String
            FileInputPath = Library.UseFileDialog(msoFileDialogFilePicker)
            
            Call Library.InsertImage(EncodeFile(FileInputPath), .cell(Row:=1, Column:=2).Range)
            .cell(Row:=1, Column:=2).Range.style = "8ptCenter"
            
            ' Insert text and styles
            Dim CombinedTitle As String
            CombinedTitle = ActiveDocument.CustomDocumentProperties("Titulo1") & vbCrLf & _
                       ActiveDocument.CustomDocumentProperties("Titulo2") & vbCrLf & _
                       ActiveDocument.CustomDocumentProperties("Titulo3") & vbCrLf & _
                       ActiveDocument.CustomDocumentProperties("Titulo4") & vbCrLf & _
                       ActiveDocument.CustomDocumentProperties("Titulo5")
            insertCellText .cell(1, 3), "CLASSIFICAÇÃO", "8ptCenter", 0, wdBorderBottom
            insertCellText .cell(2, 3), "RESTRITA", "10ptCenterBold", 0, wdBorderTop
            insertCellText .cell(1, 4), ActiveDocument.CustomDocumentProperties("Projeto"), "11ptCenterBold", 0
            insertCellText .cell(3, 1), CombinedTitle, "10ptLeftBold", 0
            insertCellText .cell(3, 2), "Nº VALE", "8ptLeft", 0, wdBorderBottom
            insertCellText .cell(4, 2), ActiveDocument.CustomDocumentProperties("NumeroCliente"), "10ptCenterBold", 0, wdBorderTop
            insertCellText .cell(5, 2), "Nº BRASS", "8ptLeft", 0, wdBorderBottom
            insertCellText .cell(6, 2), ActiveDocument.CustomDocumentProperties("NumeroNosso"), "10ptCenterBold", 0, wdBorderTop
            insertCellText .cell(3, 3), "PÁGINA", "8ptCenter", 0, wdBorderBottom
            insertCellText .cell(5, 3), "REV.", "8ptCenter", 0, wdBorderBottom
            insertCellText .cell(6, 3), ActiveDocument.CustomDocumentProperties("Revisao"), "10ptCenterBold", 0, wdBorderTop
            
            ' Add page number fields
            With .cell(Row:=4, Column:=3).Range
                .Collapse wdCollapseStart
                .Fields.Add Range:=.Duplicate, Type:=wdFieldNumPages
                .InsertBefore "/"
                .Collapse wdCollapseStart
                .Fields.Add Range:=.Duplicate, Type:=wdFieldPage
                .style = "10ptCenterBold"
                .Borders(wdBorderTop).LineStyle = wdLineStyleNone
            End With
        End With
    Else
        MsgBox "No primary header found in the document.", vbExclamation
    End If
    
End Sub

Sub insertCellText(cell As cell, text As String, style As String, Optional width As Single = 0, Optional border1 As WdBorderType = 0, Optional border2 As WdBorderType = 0)
    With cell
        .Range.InsertAfter text
        .Range.style = style
        If width > 0 Then .SetWidth ColumnWidth:=CentimetersToPoints(width), RulerStyle:=wdAdjustNone
        If border1 < 0 Then .Range.Borders.Item(border1).LineStyle = wdLineStyleNone
        If border2 < 0 Then .Range.Borders.Item(border2).LineStyle = wdLineStyleNone
    End With
End Sub

Sub insertStyle(styleName As String, baseStyleName As String, fontName As String, fontSize As Single, fontBold As Boolean, fontItalic As Boolean, fontUnderline As WdUnderline, fontColor As Long, fontAllCaps As Boolean, paragraphAlignment As WdParagraphAlignment, paragraphLeftIndent As Single, paragraphSpaceBefore As Single, paragraphSpaceAfter As Single, tabStops As Variant, lineSpacingRule As WdLineSpacing, FirstLineIndent As Integer)

    Dim myStyle As style
    Dim ts As tabStop
    
    On Error Resume Next
    Set myStyle = ActiveDocument.Styles(styleName)
    If myStyle Is Nothing Then
        Set myStyle = ActiveDocument.Styles.Add(name:=styleName, Type:=wdStyleTypeParagraph)
    End If
    On Error GoTo 0
    
    With myStyle
        .AutomaticallyUpdate = False
        .BaseStyle = baseStyleName
        .NextParagraphStyle = "Normal"
        .NoSpaceBetweenParagraphsOfSameStyle = False
        
        With .Font
            .name = fontName
            .Size = fontSize
            .Bold = fontBold
            .Italic = fontItalic
            .Underline = fontUnderline
            .Color = fontColor
            .AllCaps = fontAllCaps
        End With
        
        With .ParagraphFormat
            .LeftIndent = CentimetersToPoints(paragraphLeftIndent)
            .Alignment = paragraphAlignment
            .SpaceBefore = paragraphSpaceBefore
            .SpaceAfter = paragraphSpaceAfter
            .lineSpacingRule = lineSpacingRule
            .FirstLineIndent = FirstLineIndent
            
            .tabStops.ClearAll
            
            ' Adiciona os tab stops
            Dim tabStop As Variant
            For Each tabStop In tabStops
                .tabStops.Add Position:=CentimetersToPoints(tabStop(0)), Alignment:=tabStop(1), Leader:=tabStop(2)
            Next tabStop
        End With
    End With
End Sub

Sub setupStyles()
    ' *insertStyle*
    ' styleName (0)
    ' baseStyleName (1)
    ' fontName (2)
    ' fontSize (3)
    ' fontBold (4)
    ' fontItalic (5)
    ' fontUnderline (6)
    ' fontColor (7)
    ' fontAllCaps (8)
    ' paragraphAlignment (9)
    ' paragraphLeftIndent (10)
    ' paragraphSpaceBefore (11)
    ' paragraphSpaceAfter (12)
    ' lineSpacingRule (13)
    ' FirstLineIndent (14)
    insertStyle "7ptLeft", "", "Arial", 7, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphLeft, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "8ptLeft", "", "Arial", 8, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphLeft, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "8ptCenter", "", "Arial", 8, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphCenter, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "10ptLeftBold", "", "Arial", 10, True, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphLeft, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "10ptCenterBold", "", "Arial", 10, True, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphCenter, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "11ptCenterBold", "", "Arial", 11, True, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphCenter, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "12ptCenter", "", "Arial", 12, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphCenter, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "12ptLeftUnderline", "", "Arial", 12, False, False, wdUnderlineSingle, RGB(0, 0, 0), False, wdAlignParagraphLeft, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "12ptLeftBoldUnderline", "", "Arial", 12, True, False, wdUnderlineSingle, RGB(0, 0, 0), False, wdAlignParagraphLeft, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "12ptCenterBoldUnderline", "", "Arial", 12, True, False, wdUnderlineSingle, RGB(0, 0, 0), False, wdAlignParagraphCenter, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "Parágrafo Normal_VALE_", "", "Arial", 12, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphJustify, 0, 0, 14, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "Tabela Normal_VALE_", "", "Arial", 12, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphCenter, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "Título I_VALE_", "Heading 1", "Arial", 12, True, False, wdUnderlineNone, RGB(0, 0, 0), True, wdAlignParagraphJustify, 0, 14, 14, Array(Array(2, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "Título I.I_VALE_", "Heading 2", "Arial", 12, False, False, wdUnderlineNone, RGB(0, 0, 0), True, wdAlignParagraphJustify, 0, 14, 14, Array(Array(2, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "Título I.I.I_VALE_", "Heading 3", "Arial", 12, False, False, wdUnderlineSingle, RGB(0, 0, 0), True, wdAlignParagraphJustify, 0, 14, 14, Array(Array(2, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "Título I.I.I.I_VALE_", "Heading 4", "Arial", 12, False, False, wdUnderlineNone, RGB(0, 0, 0), True, wdAlignParagraphJustify, 0, 14, 14, Array(Array(2, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    
    insertStyle "Lista Texto", "", "Arial", 12, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphJustify, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "Lista Titulo", "", "Arial", 12, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphJustify, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    
    insertStyle "TOC 1", "", "Arial", 12, True, False, wdUnderlineNone, RGB(0, 0, 0), True, wdAlignParagraphJustify, 0, 0, 0, Array(Array(1.5, wdAlignTabLeft, wdTabLeaderSpaces), Array(17.5, wdAlignTabRight, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    insertStyle "TOC 2", "", "Arial", 12, False, False, wdUnderlineNone, RGB(0, 0, 0), True, wdAlignParagraphJustify, 0, 0, 0, Array(Array(1.5, wdAlignTabLeft, wdTabLeaderSpaces), Array(17.5, wdAlignTabRight, wdTabLeaderSpaces)), wdLineSpaceSingle, 0
    
    ' Configure the list levels
    'setupList 1, "%1.", 0, 2, 2, "Título I_VALE_"
    'setupList 2, "%1.%2.", 0, 2, 2, "Título I.I_VALE_"
    'setupList 3, "%1.%2.%3.", 0, 2, 2, "Título I.I.I_VALE_"
    'setupList 4, "%1.%2.%3.%4.", 0, 2, 2, "Título I.I.I.I_VALE_"

    Dim doc As Document
    Set doc = ActiveDocument
    
    'Dim Teste As listTemplate
    Set Teste = insertStyleBulletList(doc, "Listas")
    Set Teste = insertStyleNumbertList(doc, "Lista Titulo", "Título I_VALE_", "Título I.I_VALE_", "Título I.I.I_VALE_", "Título I.I.I.I_VALE_")
    
    
    ' Create a new style for the list
    'On Error Resume Next
    'Dim listStyle As style
    'Set listStyle = doc.Styles(styleName)
    'If listStyle Is Nothing Then
    '    Set listStyle = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeParagraph)
    'End If
    'On Error GoTo 0
    
End Sub
    
Function insertStyleNumbertList(doc As Document, styleName As String, firstLevelStyle As String, secondLevelStyle As String, thirdLevelStyle As String, fourthLevelStyle As String) As listTemplate
    ' Create a new list template
    Dim listTemplate As listTemplate
    Set listTemplate = doc.ListTemplates.Add(OutlineNumbered:=True)
    
    ' Configure the list levels
    With listTemplate.ListLevels(1)
        .NumberFormat = "%1." ' Bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignJustify
        .TextPosition = CentimetersToPoints(2)
        .TabPosition = CentimetersToPoints(2)
        .ResetOnHigher = 0
        .StartAt = 1
        .LinkedStyle = firstLevelStyle
    End With
    
    With listTemplate.ListLevels(2)
        .NumberFormat = "%1.%2." ' Hollow bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2)
        .TabPosition = CentimetersToPoints(2)
        .ResetOnHigher = 1
        .StartAt = 1
        .LinkedStyle = secondLevelStyle
    End With
    
    With listTemplate.ListLevels(3)
        .NumberFormat = "%1.%2.%3." ' Hollow bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2)
        .TabPosition = CentimetersToPoints(2)
        .ResetOnHigher = 1
        .StartAt = 1
        .LinkedStyle = thirdLevelStyle
    End With
    
    With listTemplate.ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4." ' Hollow bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2)
        .TabPosition = CentimetersToPoints(2)
        .ResetOnHigher = 1
        .StartAt = 1
        .LinkedStyle = fourthLevelStyle
    End With
    
    Set insertStyleNumbertList = listTemplate

End Function

Function insertStyleBulletList(doc As Document, styleName As String) As listTemplate
    ' Create a new list template
    Dim listTemplate As listTemplate
    Set listTemplate = doc.ListTemplates.Add(OutlineNumbered:=True)
    
    ' Configure the list levels
    With listTemplate.ListLevels(1)
        .NumberFormat = ChrW(&H2022) ' Bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = CentimetersToPoints(1)
        .Alignment = wdListLevelAlignJustify
        .TextPosition = CentimetersToPoints(2)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        .LinkedStyle = "Lista Texto"
    End With
    
    With listTemplate.ListLevels(2)
        .NumberFormat = ChrW(&H25E6) ' Hollow bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = CentimetersToPoints(2)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(3)
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
        .LinkedStyle = "Lista Texto"
    End With
    
    With listTemplate.ListLevels(3)
        .NumberFormat = ChrW(&H25AA) ' Filled square bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = CentimetersToPoints(3)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(4)
        .TabPosition = wdUndefined
        .ResetOnHigher = 2
        .StartAt = 1
        .LinkedStyle = "Lista Texto"
    End With
    
    With listTemplate.ListLevels(4)
        .NumberFormat = ChrW(&H2022) ' Bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = CentimetersToPoints(4)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(5)
        .TabPosition = wdUndefined
        .ResetOnHigher = 3
        .StartAt = 1
        .LinkedStyle = "Lista Texto"
    End With
    
    With listTemplate.ListLevels(5)
        .NumberFormat = ChrW(&H25E6) ' Hollow bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = CentimetersToPoints(5)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(6)
        .TabPosition = wdUndefined
        .ResetOnHigher = 4
        .StartAt = 1
        .LinkedStyle = "Lista Texto"
    End With
    
    With listTemplate.ListLevels(6)
        .NumberFormat = ChrW(&H25AA) ' Filled square bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = CentimetersToPoints(6)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(7)
        .TabPosition = wdUndefined
        .ResetOnHigher = 4
        .StartAt = 1
        .LinkedStyle = "Lista Texto"
    End With
    
    With listTemplate.ListLevels(7)
        .NumberFormat = ChrW(&H2022) ' Bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = CentimetersToPoints(7)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(8)
        .TabPosition = wdUndefined
        .ResetOnHigher = 4
        .StartAt = 1
        .LinkedStyle = "Lista Texto"
    End With
    
    With listTemplate.ListLevels(8)
        .NumberFormat = ChrW(&H25E6) ' Hollow bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = CentimetersToPoints(8)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(9)
        .TabPosition = wdUndefined
        .ResetOnHigher = 4
        .StartAt = 1
        .LinkedStyle = "Lista Texto"
    End With
    
    With listTemplate.ListLevels(9)
        .NumberFormat = ChrW(&H25AA) ' Filled square bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = CentimetersToPoints(9)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(10)
        .TabPosition = wdUndefined
        .ResetOnHigher = 4
        .StartAt = 1
        .LinkedStyle = "Lista Texto"
    End With

    Set insertStyleBulletList = insertStyleBulletList
End Function

Sub setupPage()
    With ActiveDocument.PageSetup
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1)
        .LeftMargin = Application.CentimetersToPoints(2.5)
        .RightMargin = Application.CentimetersToPoints(1)
        .HeaderDistance = Application.CentimetersToPoints(1.6)
        .FooterDistance = Application.CentimetersToPoints(0)
    End With
End Sub

Private Function getValeLogoBase64() As String
    
    Dim ValeLogoBase64 As String
    ValeLogoBase64 = ""
    ValeLogoBase64 = ValeLogoBase64 & "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGB"
    ValeLogoBase64 = ValeLogoBase64 & "YUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCAAi"
    ValeLogoBase64 = ValeLogoBase64 & "AFcDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhBy"
    ValeLogoBase64 = ValeLogoBase64 & "JxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKT"
    ValeLogoBase64 = ValeLogoBase64 & "lJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAA"
    ValeLogoBase64 = ValeLogoBase64 & "AAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRom"
    ValeLogoBase64 = ValeLogoBase64 & "JygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExc"
    ValeLogoBase64 = ValeLogoBase64 & "bHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDQ/aq8bfFG2+JWreH/ABfrVzHaxuXtLWxZobOW3J+R1UH5uOu7"
    ValeLogoBase64 = ValeLogoBase64 & "JByK8f8AC/jbxB4J1FL/AEDWb3SLtTu8y0mKZ9mHRh7Hiv01/aq/Z/g+OHgVjZRpH4o0xWm06c8eZ3aFj/dbH4HBr8tbyzn0+7ntbqF7e6gdop"
    ValeLogoBase64 = ValeLogoBase64 & "YZBhkYHBUj1BqgP0B/Zm/bYt/HNxbeGPHbwadrz4jtdTXCQXZ7Kw6I5/I+3SvpTxR8QND8I2zS399GJMfLbxndI59Aor8ZP0Neo+C/jtqOhQLa"
    ValeLogoBase64 = ValeLogoBase64 & "6xE+r26jCTl/3yj0JP3h9eaLAfaniz49a/rVwy6W/wDY9mD8ojAaVh6sx6fhXpHwO1LxRrmmXF/rV4bjT2O2285B5jEdWz/d7c18sfAnU7v48+"
    ValeLogoBase64 = ValeLogoBase64 & "M103SdLmtNGs8TalqVwwxGmeI0A6u3Qegye1feVnZw6faw21vGsUEKhEjUYCgDAFDsImoooqRhRRRQAUUUUAFfEH7dv7OfEvxK8O2vIwNatol6"
    ValeLogoBase64 = ValeLogoBase64 & "joLgAfk34H1r7fqG8s4NQtJrW5iSe3mQxyRSDKupGCCO4IoA/Eetnwf4R1Tx54n07w/otu11qV/KIoox0HqzeigZJPoK9U/ao/Z/n+B/jomxie"
    ValeLogoBase64 = ValeLogoBase64 & "TwtqjNLp0uCfKOctAT6rnj1GK+tv2Lf2cx8L/DI8Va7bBfFOrxArHIvzWdueQnszcFvwHrVAeu/BH4Q6X8FPANl4e04LJMo828u9uGuZyPmc+3"
    ValeLogoBase64 = ValeLogoBase64 & "YDsAKZ8doddHw31G/wDDl7PZatphW/j8hiPNWM7njb1Urnj2r0GmyRrNGyOodGG1lYZBB7VIHiXjH4jXfjzTfhxpnhe+msbrxRMl5PPbtiSC1i"
    ValeLogoBase64 = ValeLogoBase64 & "AaYZHT5vkP41Q/aK8beJYdWsfD/g29e1v9Ns5dd1Bo3wTBHwsXuWOeO/FbHwe+A8/w28aa3qd1fJfWCqbfRYskm1geRpHUg9DkgZHXmm2f7PVn"
    ValeLogoBase64 = ValeLogoBase64 & "4o8TeJfEHjdRf3+oXWLVLW5kRYLVVCohIxk9c1Wgix8VPiFPqX7NupeLdBvJLG4n0+K5hnt2w0TF1DAH1B3D8K9JW4l/4RMT+Y3nfYd/mZ53eX"
    ValeLogoBase64 = ValeLogoBase64 & "nP1zXj0fwN8QWvwZ8Y+AYry1e1urlzo0kjt+7gaQPsfjjGD09a6/wpZ/Edlaw8RReH00v7G0CvYPKZd+3audwxj1pAeP8AwN1y18cLpMWrfEDx"
    ValeLogoBase64 = ValeLogoBase64 & "ZNr9wJvM0+OSRLf5Qx+/sxwoz97rRXa/Dfwb8VPh54dsdBgj8MT6fa79s0kk3mnJJ5wMdTRTA90oooqRnDfFzS7LVtF0dL2zt7xI9Ys3RbiJXC"
    ValeLogoBase64 = ValeLogoBase64 & "sJRggEcH3ruaKKACiiigAooooAKKKKACiiigD/2Q=="
    
    getValeLogoBase64 = ValeLogoBase64
    
End Function

