Attribute VB_Name = "Teste"
Sub Teste()
    Dim destinationDoc As Document
    Dim sourceDoc As Document
    
    Set sourceDoc = Documents("BdB201460-0000-V-ET0001.docx")
    Set destinationDoc = Documents("ConvertedLayout.docx")
    
    
    Call libUtils.CopyFieldsBetweenDocs(sourceDoc, destinationDoc)
End Sub

Sub setupStylesTest()
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Setup all styles
    Dim fontFace As String
    fontFace = "Arial"
    
    Call libUtils.NegritoCamposTabelaFigura(doc)
    Call libstyles.setupStyles(doc, fontFace)
End Sub

Sub deleteCustomStyles()
    Dim doc As Document
    Set doc = ActiveDocument
    'Call libUtils.DeleteUnusedStyles(doc)
    'Call libUtils.RestrictToSpecificStyles(doc)
    Call libstyles.RenameHeadingStylesToDefault(doc)
End Sub
