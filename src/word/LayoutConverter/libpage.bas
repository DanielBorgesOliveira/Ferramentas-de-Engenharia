Attribute VB_Name = "libpage"
Sub setupPage(doc As Document, pageMargins As Variant)
    With doc.PageSetup
        .TopMargin = pageMargins(0)
        .BottomMargin = pageMargins(1)
        .LeftMargin = pageMargins(2)
        .RightMargin = pageMargins(3)
        .HeaderDistance = pageMargins(4)
        .FooterDistance = pageMargins(5)
    End With
End Sub
