Attribute VB_Name = "libstyles"
Sub setupStyles(doc As Document, Optional fontFace As String = "Arial")
    Call createStyle(doc, "1 - Parágrafo Normal", "", fontFace, 12, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphJustify, 0, 0, 12, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0)
    Call createStyle(doc, "1 - Tabela", "", fontFace, 10, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphCenter, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0)
    Call createStyle(doc, "1 - Legenda", "", fontFace, 10, True, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphCenter, 0, 12, 12, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0, False)
    Call createStyle(doc, "1 - Figura", "", fontFace, 12, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphCenter, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0, True)
    Call createStyle(doc, "1 - Equacao", "", fontFace, 12, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphJustify, 0, 0, 12, Array(Array(17, wdAlignTabRight, wdTabLeaderHeavy)), wdLineSpaceSingle, 0)
    Call createStyle(doc, "1 - Título I", "Heading 1", fontFace, 12, True, False, wdUnderlineNone, RGB(0, 0, 0), True, wdAlignParagraphJustify, 0, 24, 12, Array(Array(2, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0, True)
    Call createStyle(doc, "1 - Título I.I", "Heading 2", fontFace, 12, False, False, wdUnderlineNone, RGB(0, 0, 0), True, wdAlignParagraphJustify, 0, 12, 12, Array(Array(2, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0, True)
    Call createStyle(doc, "1 - Título I.I.I", "Heading 3", fontFace, 12, False, False, wdUnderlineSingle, RGB(0, 0, 0), False, wdAlignParagraphJustify, 0, 12, 12, Array(Array(2, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0, True)
    Call createStyle(doc, "1 - Título I.I.I.I", "Heading 4", fontFace, 12, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphJustify, 0, 12, 12, Array(Array(2, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0, True)
    Call createStyle(doc, "1 - Título I.I.I.I.I", "Heading 5", fontFace, 12, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphJustify, 0, 12, 12, Array(Array(2, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0, True)
    Call createStyle(doc, "TOC 1", "", fontFace, 12, True, False, wdUnderlineNone, RGB(0, 0, 0), True, wdAlignParagraphJustify, 0, 0, 12, Array(Array(1.5, wdAlignTabLeft, wdTabLeaderSpaces), Array(17.5, wdAlignTabRight, wdTabLeaderSpaces)), wdLineSpaceSingle, 0)
    Call createStyle(doc, "TOC 2", "", fontFace, 12, False, False, wdUnderlineNone, RGB(0, 0, 0), True, wdAlignParagraphJustify, 0, 0, 12, Array(Array(1.5, wdAlignTabLeft, wdTabLeaderSpaces), Array(17.5, wdAlignTabRight, wdTabLeaderSpaces)), wdLineSpaceSingle, 0)
    
    Call createStyle(doc, "1 - Bullet List", "", fontFace, 12, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphJustify, 10, 0, 12, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0, False, True, True)
    Call createStyleBulletList(doc, "1 - Bullet List")
    
    Call createStyle(doc, "1 - Title List", "", fontFace, 12, False, False, wdUnderlineNone, RGB(0, 0, 0), False, wdAlignParagraphJustify, 0, 0, 0, Array(Array(0, wdAlignTabLeft, wdTabLeaderSpaces)), wdLineSpaceSingle, 0)
    Call createStyleNumbertList(doc, "1 - Title List", "1 - Título I", "1 - Título I.I", "1 - Título I.I.I", "1 - Título I.I.I.I", "1 - Título I.I.I.I.I")
    
    ' Convert original formatting to client's formatting.
    Dim DePara As Variant
    DePara = Array( _
        Array("Normal", "1 - Parágrafo Normal"), _
        Array("Heading 1", "1 - Título I"), _
        Array("Heading 2", "1 - Título I.I"), _
        Array("Heading 3", "1 - Título I.I.I"), _
        Array("Heading 4", "1 - Título I.I.I.I"), _
        Array("Heading 5", "1 - Título I.I.I.I.I"), _
        Array("Caption", "1 - Legenda"), _
        Array("List Paragraph", "1 - Bullet List") _
    )
    Call updateClientText(doc:=doc, DePara:=DePara, tableStyle:="1 - Tabela", figureStyle:="1 - Figura", bulletStyle:="1 - Bullet List", equationStyle:="1 - Equacao")
End Sub

Sub createStyle( _
    doc As Document, _
    styleName As String, _
    baseStyleName As String, _
    fontName As String, _
    fontSize As Single, _
    fontBold As Boolean, _
    fontItalic As Boolean, _
    fontUnderline As WdUnderline, _
    fontColor As Long, _
    fontAllCaps As Boolean, _
    paragraphAlignment As WdParagraphAlignment, _
    paragraphLeftIndent As Single, _
    paragraphSpaceBefore As Single, _
    paragraphSpaceAfter As Single, _
    tabStops As Variant, _
    lineSpacingRule As WdLineSpacing, _
    FirstLineIndent As Integer, _
    Optional KeepWithNext As Boolean = False, _
    Optional SpaceAfterAuto As Boolean = False, _
    Optional NoSpaceBetweenParagraphsOfSameStyle As Boolean = False)

    Dim myStyle As Style
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
        .NextParagraphStyle = "1 - Parágrafo Normal"
        .NoSpaceBetweenParagraphsOfSameStyle = NoSpaceBetweenParagraphsOfSameStyle
        
        With .Font
            .name = fontName
            .size = fontSize
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
            .KeepWithNext = KeepWithNext
            .SpaceAfterAuto = SpaceAfterAuto
            
            ' Adiciona os tab stops
            .tabStops.ClearAll
            Dim tabStop As Variant
            For Each tabStop In tabStops
                .tabStops.Add Position:=CentimetersToPoints(tabStop(0)), Alignment:=tabStop(1), Leader:=tabStop(2)
            Next tabStop
        End With
    End With
End Sub

Sub createStyleBulletList(doc As Document, styleName As String)
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
        .LinkedStyle = doc.Styles(styleName)
    End With
    
    With listTemplate.ListLevels(2)
        .NumberFormat = ChrW(&H2013) ' Dash character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = CentimetersToPoints(2)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(3)
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
        '.LinkedStyle = styleName
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
        '.LinkedStyle = styleName
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
        '.LinkedStyle = styleName
    End With
    
    With listTemplate.ListLevels(5)
        .NumberFormat = ChrW(&H2013) ' Dash character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = CentimetersToPoints(5)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(6)
        .TabPosition = wdUndefined
        .ResetOnHigher = 4
        .StartAt = 1
        '.LinkedStyle = styleName
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
        '.LinkedStyle = styleName
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
        '.LinkedStyle = styleName
    End With
    
    With listTemplate.ListLevels(8)
        .NumberFormat = ChrW(&H2013) ' Dash character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = CentimetersToPoints(8)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(9)
        .TabPosition = wdUndefined
        .ResetOnHigher = 4
        .StartAt = 1
        '.LinkedStyle = styleName
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
        '.LinkedStyle = styleName
    End With
End Sub

Sub createStyleNumbertList(doc As Document, styleName As String, level1 As String, level2 As String, level3 As String, level4 As String, level5 As String)
    ' Create a new list template
    Dim listTemplate As listTemplate
    Set listTemplate = doc.ListTemplates.Add(OutlineNumbered:=True)
    
    ' Configure the list levels
    With listTemplate.ListLevels(1)
        .NumberFormat = "%1.0" ' Bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignJustify
        .TextPosition = CentimetersToPoints(2)
        .TabPosition = CentimetersToPoints(2)
        .ResetOnHigher = 0
        .StartAt = 1
        .LinkedStyle = doc.Styles(level1)
    End With
    
    With listTemplate.ListLevels(2)
        .NumberFormat = "%1.%2" ' Hollow bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2)
        .TabPosition = CentimetersToPoints(2)
        .ResetOnHigher = 1
        .StartAt = 1
        .LinkedStyle = doc.Styles(level2)
    End With
    
    With listTemplate.ListLevels(3)
        .NumberFormat = "%1.%2.%3" ' Hollow bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2)
        .TabPosition = CentimetersToPoints(2)
        .ResetOnHigher = 1
        .StartAt = 1
        .LinkedStyle = doc.Styles(level3)
    End With
    
    With listTemplate.ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4" ' Hollow bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2)
        .TabPosition = CentimetersToPoints(2)
        .ResetOnHigher = 1
        .StartAt = 1
        .LinkedStyle = doc.Styles(level4)
    End With
    
    With listTemplate.ListLevels(5)
        .NumberFormat = "%1.%2.%3.%4.%5" ' Hollow bullet character
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2)
        .TabPosition = CentimetersToPoints(2)
        .ResetOnHigher = 1
        .StartAt = 1
        .LinkedStyle = doc.Styles(level5)
    End With
End Sub

Sub deleteCustomStyles(doc As Document)
    Dim oStyle As Style
    For Each oStyle In doc.Styles
        If Not oStyle.BuiltIn Then
            oStyle.Delete
        End If
    Next oStyle
End Sub

Function ExistsStyle(doc As Document, styleName As String) As Boolean
    On Error Resume Next
    ExistsStyle = Not doc.Styles(styleName) Is Nothing
    On Error GoTo 0
End Function

Sub updateClientText(doc As Document, DePara As Variant, tableStyle As String, figureStyle As String, bulletStyle As String, equationStyle As String)
    ' Atualiza os estilos do documento conforme padrão do client.

    '--------------------------------------
    ' Atualiza estilos de parágrafos em listas
    '--------------------------------------
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        If para.Range.ListFormat.ListType = wdListBullet Then
            If ExistsStyle(doc, "Bullet List") Then
                para.Range.Style = bulletStyle
            End If
        End If
    Next para

    '--------------------------------------
    ' Atualiza estilos com base no mapeamento DePara
    '--------------------------------------
    Dim i As Integer
    For i = LBound(DePara) To UBound(DePara)
        Dim originalStyle As String
        Dim targetStyle As String
        originalStyle = DePara(i)(0)
        targetStyle = DePara(i)(1)
        If ExistsStyle(doc, originalStyle) And ExistsStyle(doc, targetStyle) Then
            Dim p As Paragraph
            For Each p In doc.Paragraphs
                If p.Style = originalStyle Then
                    p.Style = targetStyle
                End If
            Next p
        End If
    Next i

    '--------------------------------------
    ' Aplica estilo às células das tabelas (a partir da segunda)
    '--------------------------------------
    Dim tbl As Table
    Dim cel As cell
    For i = 2 To doc.Tables.count
        Set tbl = doc.Tables(i)
        For Each cel In tbl.Range.Cells
            If ExistsStyle(doc, tableStyle) Then
                cel.Range.Style = tableStyle
            End If
        Next cel
    Next i

    '--------------------------------------
    ' Aplica estilo a figuras inline
    '--------------------------------------
    Dim fig As InlineShape
    If ExistsStyle(doc, figureStyle) Then
        For Each fig In doc.InlineShapes
            fig.Range.Style = figureStyle
        Next fig
    End If
    
    '--------------------------------------
    ' Aplica estilo apenas a parágrafos que:
    ' - Contenham uma equação (OMath)
    ' - Contenham também a referência cruzada "Equation"
    '--------------------------------------
    If ExistsStyle(doc, equationStyle) Then
        For Each para In doc.Paragraphs
            Dim hasOMath As Boolean
            Dim hasEquationRef As Boolean
            Dim field As field
    
            hasOMath = (para.Range.OMaths.count > 0)
            hasEquationRef = False
    
            If para.Range.Fields.count > 0 Then
                For Each field In para.Range.Fields
                    If InStr(1, field.Code.text, "SEQ Equation", vbTextCompare) > 0 Then
                        hasEquationRef = True
                        Exit For
                    End If
                Next field
            End If
    
            If hasOMath And hasEquationRef Then
                para.Style = equationStyle
            End If
        Next para
    End If
End Sub

Sub RenameHeadingStylesToDefault(doc As Document)
    Dim i As Integer
    Dim builtinName As String

    On Error Resume Next

    For i = 1 To 9
        builtinName = "Heading " & i
        With doc.Styles(builtinName)
            If Not .NameLocal = builtinName Then
                .NameLocal = builtinName
            End If
        End With
    Next i
End Sub

