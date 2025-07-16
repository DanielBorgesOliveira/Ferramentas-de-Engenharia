Attribute VB_Name = "LayoutConverter"
Sub LayoutConverter()
    'Dim fd As FileDialog
    
    UserForm2.Show
    
    ' Client name
    Dim client As String
    client = UserForm2.TextBox3.value
    
    ' Template Document Path
    Dim templatePath As String
    templatePath = UserForm2.TextBox5.value
    
    ' Source Document Path
    Dim sourcePath As String
    sourcePath = UserForm2.TextBox7.value
    
    ' Destination Document Path
    Dim destinationPath As String
    destinationPath = UserForm2.TextBox6.value

    ' Open the template
    Dim templateDoc As Document
    Set templateDoc = Documents.Open(fileName:=templatePath, ReadOnly:=True)

    ' Save as a new document
    templateDoc.SaveAs2 fileName:=destinationPath

    ' Close the template document
    templateDoc.Close SaveChanges:=False
    
    ' === Save As Dialog ===
    'MsgBox "Indique o local de destino para salvar convertido.", vbInformation
    'Set fd = Application.FileDialog(msoFileDialogSaveAs)
    'With fd
    '    .Title = "Save As - Choose location and name for the copy"
    '    .InitialFileName = "ConvertedLayout.docx"

        'If .Show = -1 Then
            'destinationPath = .SelectedItems(1)

            ' Open the template
            'Set templateDoc = Documents.Open(fileName:=templatePath, ReadOnly:=True)

            ' Save as a new document
            'templateDoc.SaveAs2 fileName:=destinationPath

            ' Close the template document
            'templateDoc.Close SaveChanges:=False
        'Else
        '    MsgBox "Save canceled by user.", vbInformation
        '    Exit Sub
        'End If
    'End With

    ' === Reopen the destination document ===
    Dim destinationDoc As Document
    Set destinationDoc = Documents.Open(fileName:=destinationPath, ReadOnly:=False)
    
    ' === Open source document ===
    'MsgBox "Indique o arquivo de origem com o documento no formato neutro.", vbInformation
    'Set fd = Application.FileDialog(msoFileDialogFilePicker)
    'With fd
    '    .Title = "Select Source Document to Copy Content From"
    '    .Filters.Clear
    '    .Filters.Add "Word Documents", "*.doc; *.docx"
    '    If .Show = -1 Then
    '        Set sourceDoc = Documents.Open(fileName:=.SelectedItems(1), ReadOnly:=False)
    '    Else
    '        MsgBox "No source document selected.", vbExclamation
    '        Exit Sub
    '    End If
    'End With
    Dim sourceDoc As Document
    Set sourceDoc = Documents.Open(fileName:=sourcePath, ReadOnly:=False)
    
    'Call libstyles.deleteCustomStyles(sourceDoc)
    
    Call libUtils.DeleteUnusedStyles(destinationDoc)
    
    Call libUtils.CopyCustomDocProperties(sourceDoc, destinationDoc)
    
    Call libUtils.CopyBuiltinDocProperties(sourceDoc, destinationDoc)
    
    Call libstyles.RenameHeadingStylesToDefault(destinationDoc)
    
    ' Setup all styles
    Dim fontFace As String
    Select Case client
    Case "Samarco"
        fontFace = "Times New Roman"
    Case "AngloAmerican"
        fontFace = "Aptos"
    Case Else
        fontFace = "Arial"
    End Select
    
    Call libstyles.setupStyles(sourceDoc, fontFace)

    ' === Copy and Paste content ===
    sourceDoc.Content.Copy

    ' Get a Range at the end of the destination document
    Dim endRange As Range
    Set endRange = destinationDoc.Content
    endRange.Collapse Direction:=wdCollapseEnd
    
    ' Add table of contents
    'Call addTOC(destinationDoc, endRange)
    
    ' Paste at the end
    'endRange.PasteAndFormat wdFormatSurroundingFormattingWithEmphasis
    endRange.Paste
    
    ' Restrict the allowed styles.
    'Call libUtils.RestrictToSpecificStyles(destinationDoc)

    ' Optional: Close source document
    sourceDoc.Close SaveChanges:=False

    ' Bring destination document to front
    destinationDoc.Activate
End Sub

Sub addTOC(doc As Document, tocRange As Range)
    ' Add the Table of Contents to the defined range
    Dim TOC As TableOfContents
    Set TOC = doc.TablesOfContents.Add( _
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
        
        tocRange.InsertBefore "ÍNDICE" & vbCrLf
        tocRange.Paragraphs(1).Style = "12ptCenterBoldUnderline"
        
        tocRange.InsertAfter "ITEM" & vbTab & "DESCRIÇÃO" & vbTab & "PÁGINA" & vbCrLf
        tocRange.Paragraphs(2).Style = "12ptCenterBoldUnderline"
        
        .Update
    End With
End Sub


