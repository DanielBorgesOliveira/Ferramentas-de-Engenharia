Attribute VB_Name = "Teste"
Sub teste()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Call libUtils.ItalicizeNonPortugueseText(doc)
End Sub
