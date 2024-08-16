Attribute VB_Name = "AngloAmerican"
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
    
    Call Library.DeleteCustomStyles
    Library.Sleep 500
    Call setupPage
    Library.Sleep 500
    Call setupStyles
    Library.Sleep 500
    'Call updateClientText
    'Library.Sleep 500
    Call addHeader
    Library.Sleep 500
    Call addFooter
    Library.Sleep 500
    Call addRevisionTable
    'Library.Sleep 500
    'Call addTOC
End Sub

Sub addFooter()
    ' Check if the document has a primary header
    If ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Exists Then
        ' Set the header range
        Dim footerRange As Range
        Set footerRange = ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range
        
        ' Clear existing content in the footer
        footerRange.text = ""
    
        Dim footerTable As Table
        Set footerTable = footerRange.Tables.Add( _
            Range:=footerRange, _
            NumRows:=6, _
            NumColumns:=3, _
            DefaultTableBehavior:=wdWord9TableBehavior, _
            AutoFitBehavior:=wdAutoFitFixed _
        )
        
        With footerTable
            .Borders.Enable = True
            
            For i = 1 To .Rows.Count
                With .Rows(i)
                    .Cells.VerticalAlignment = wdAlignVerticalCenter
                End With
            Next i
            
            .Columns(1).width = Application.CentimetersToPoints(5.83)
            .Columns(2).width = Application.CentimetersToPoints(5.83)
            .Columns(3).width = Application.CentimetersToPoints(5.83)
            
            .Rows.SetHeight RowHeight:=CentimetersToPoints(0.7), HeightRule:=wdRowHeightExactly
            .Rows(1).Height = Application.CentimetersToPoints(2)
            
            .cell(Row:=1, Column:=1).Merge MergeTo:=.cell(Row:=1, Column:=3)
            .cell(Row:=2, Column:=1).Merge MergeTo:=.cell(Row:=2, Column:=3)

            .cell(Row:=1, Column:=1).Range.InsertAfter vbCrLf & "Esta é a folha-rosto deste documento. Uma breve descrição de cada revisão do documento deverá constar nesta folha-rosto. O número da última revisão do documento constará do cabeçalho desta e das demais folhas deste documento." & vbCrLf & vbCrLf
            .cell(Row:=1, Column:=1).Range.style = "10ptLeft"
            
            .cell(Row:=2, Column:=1).Range.InsertAfter "TE - TIPO DE EMISSÃO"
            .cell(Row:=2, Column:=1).Range.style = "8ptCenterBold"
            
            .cell(Row:=3, Column:=1).Range.InsertAfter "(A) PRELIMINAR"
            .cell(Row:=3, Column:=1).Range.style = "9ptLeftBold"
            
            .cell(Row:=4, Column:=1).Range.InsertAfter "(B) PARA APROVAÇÃO"
            .cell(Row:=4, Column:=1).Range.style = "9ptLeftBold"
            
            .cell(Row:=5, Column:=1).Range.InsertAfter "(C) PARA CONHECIMENTO"
            .cell(Row:=5, Column:=1).Range.style = "9ptLeftBold"
            
            .cell(Row:=6, Column:=1).Range.InsertAfter "(D) PARA COTAÇÃO"
            .cell(Row:=6, Column:=1).Range.style = "9ptLeftBold"
            
            .cell(Row:=3, Column:=2).Range.InsertAfter "(E) PARA CONSTRUÇÃO"
            .cell(Row:=3, Column:=2).Range.style = "9ptLeftBold"
            
            .cell(Row:=4, Column:=2).Range.InsertAfter "(F) CONFORME COMPRADO"
            .cell(Row:=4, Column:=2).Range.style = "9ptLeftBold"
            
            .cell(Row:=5, Column:=2).Range.InsertAfter "(G) CONFORME CONSTRUÍDO"
            .cell(Row:=5, Column:=2).Range.style = "9ptLeftBold"
            
            .cell(Row:=6, Column:=2).Range.InsertAfter "(H) CANCELADO"
            .cell(Row:=6, Column:=2).Range.style = "9ptLeftBold"
            
            
        End With
    Else
        MsgBox "No primary footer found in the document.", vbExclamation
    End If
End Sub

Sub addRevisionTable()
    Set MyRange = ActiveDocument.Content
    MyRange.Collapse Direction:=wdCollapseStart
    
    If ActiveDocument.Tables.Count > 0 Then
        ActiveDocument.Tables(1).Delete
    End If
    
    Dim revisionTable As Table
    Set revisionTable = MyRange.Tables.Add( _
        Range:=MyRange, _
        NumRows:=25, _
        NumColumns:=8, _
        DefaultTableBehavior:=wdWord9TableBehavior, _
        AutoFitBehavior:=wdAutoFitFixed _
    )
    
    With revisionTable
        .Borders.Enable = True
        
        For i = 1 To .Rows.Count
            With .Rows(i)
                .Cells.VerticalAlignment = wdAlignVerticalCenter
            End With
        Next i
        
        ' Loop through each cell in the table
        For Each cell In .Range.Cells
            cell.Range.style = "8ptCenter"
        Next cell
        
        .Rows.SetHeight RowHeight:=CentimetersToPoints(0.7), HeightRule:=wdRowHeightExactly
        
        .Columns(1).width = Application.CentimetersToPoints(1.5)
        .Columns(2).width = Application.CentimetersToPoints(1.5)
        .Columns(3).width = Application.CentimetersToPoints(1.5)
        .Columns(4).width = Application.CentimetersToPoints(1.5)
        .Columns(5).width = Application.CentimetersToPoints(1.5)
        .Columns(6).width = Application.CentimetersToPoints(1.5)
        .Columns(7).width = Application.CentimetersToPoints(1.5)
        .Columns(8).width = Application.CentimetersToPoints(7)
        
        .Rows(1).Height = Application.CentimetersToPoints(0.55)
                
        .cell(Row:=1, Column:=1).Range.InsertAfter "REV."
        .cell(Row:=1, Column:=1).Range.style = "8ptCenterBold"
        
        .cell(Row:=1, Column:=2).Range.InsertAfter "DATA"
        .cell(Row:=1, Column:=2).Range.style = "8ptCenterBold"
        
        .cell(Row:=1, Column:=3).Range.InsertAfter "POR"
        .cell(Row:=1, Column:=3).Range.style = "8ptCenterBold"
        
        .cell(Row:=1, Column:=4).Range.InsertAfter "VER."
        .cell(Row:=1, Column:=4).Range.style = "8ptCenterBold"
        
        .cell(Row:=1, Column:=5).Range.InsertAfter "APR."
        .cell(Row:=1, Column:=5).Range.style = "8ptCenterBold"
        
        .cell(Row:=1, Column:=6).Range.InsertAfter "AUT."
        .cell(Row:=1, Column:=6).Range.style = "8ptCenterBold"
        
        .cell(Row:=1, Column:=7).Range.InsertAfter "EMISSÃO"
        .cell(Row:=1, Column:=7).Range.style = "8ptCenterBold"
        
        .cell(Row:=1, Column:=8).Range.InsertAfter "DESCRIÇÃO DE REVISÕES"
        .cell(Row:=1, Column:=8).Range.style = "8ptCenterBold"
        
    End With
    
End Sub

Sub addHeader()
    
    ' Check if the document has a primary header
    If ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Exists And ActiveDocument.Sections(2).Headers(wdHeaderFooterPrimary).Exists Then
        
        Dim Section As Integer
        For Section = 1 To 2
        
            ' Set the header range
            Dim headerRange As Range
            Set headerRange = ActiveDocument.Sections(Section).Headers(wdHeaderFooterPrimary).Range
            
            ' Clear existing content in the header
            headerRange.text = ""
        
            Dim headerTable As Table
            Set headerTable = headerRange.Tables.Add( _
                Range:=headerRange, _
                NumRows:=5, _
                NumColumns:=4, _
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
                
                .Columns(1).width = Application.CentimetersToPoints(6)
                .Columns(2).width = Application.CentimetersToPoints(4.5)
                .Columns(3).width = Application.CentimetersToPoints(5.5)
                .Columns(4).width = Application.CentimetersToPoints(1.5)
                
                .cell(Row:=1, Column:=3).Merge MergeTo:=.cell(Row:=1, Column:=4)
                .cell(Row:=2, Column:=1).Merge MergeTo:=.cell(Row:=5, Column:=2)
                
                Call Library_Image.InsertImage(getAngloLogoBase64(), .cell(Row:=1, Column:=1).Range)
                .cell(Row:=1, Column:=1).Range.style = "8ptCenter"
                
                Dim FileInputPath As String
                FileInputPath = Library.UseFileDialog(msoFileDialogFilePicker)
                
                Call Library_Image.InsertImage(EncodeFile(FileInputPath), .cell(Row:=1, Column:=2).Range)
                .cell(Row:=1, Column:=2).Range.style = "8ptCenter"
                
                .cell(Row:=1, Column:=3).Range.InsertAfter ActiveDocument.CustomDocumentProperties("Projeto")
                .cell(Row:=1, Column:=3).Range.style = "14ptCenterBold"
                
                .cell(Row:=2, Column:=1).Range.InsertAfter ActiveDocument.CustomDocumentProperties("Titulo1") & vbCrLf & ActiveDocument.CustomDocumentProperties("Titulo2") & vbCrLf & ActiveDocument.CustomDocumentProperties("Titulo3") & vbCrLf & ActiveDocument.CustomDocumentProperties("Titulo4") & vbCrLf & ActiveDocument.CustomDocumentProperties("Titulo5")
                .cell(Row:=2, Column:=1).Range.style = "10ptLeft"
                
                .cell(Row:=2, Column:=2).Range.InsertAfter "Nº. Anglo American:"
                .cell(Row:=2, Column:=2).Range.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                .cell(Row:=2, Column:=2).Range.style = "8ptLeft"
                
                .cell(Row:=3, Column:=2).Range.InsertAfter ActiveDocument.CustomDocumentProperties("NumeroCliente")
                .cell(Row:=3, Column:=2).Range.Borders(wdBorderTop).LineStyle = wdLineStyleNone
                .cell(Row:=3, Column:=2).Range.style = "8ptCenter"
                
                .cell(Row:=4, Column:=2).Range.InsertAfter "Nº. Brass:"
                .cell(Row:=4, Column:=2).Range.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                .cell(Row:=4, Column:=2).Range.style = "8ptLeft"
                
                .cell(Row:=5, Column:=2).Range.InsertAfter ActiveDocument.CustomDocumentProperties("NumeroNosso")
                .cell(Row:=5, Column:=2).Range.Borders(wdBorderTop).LineStyle = wdLineStyleNone
                .cell(Row:=5, Column:=2).Range.style = "8ptCenter"
                
                .cell(Row:=2, Column:=3).Range.InsertAfter "FOLHA"
                .cell(Row:=2, Column:=3).Range.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                .cell(Row:=2, Column:=3).Range.style = "8ptLeft"
    
                ' Add the page number field to the specified cell
                With .cell(Row:=3, Column:=3).Range
                    .Collapse wdCollapseStart ' Move to the begining of the text
                    .Fields.Add Range:=.Duplicate, Type:=wdFieldNumPages
                End With
                .cell(Row:=3, Column:=3).Range.InsertBefore "/"
                With .cell(Row:=3, Column:=3).Range
                    .Collapse wdCollapseStart ' Move to the begining of the text
                    .Fields.Add Range:=.Duplicate, Type:=wdFieldPage
                End With
                .cell(Row:=3, Column:=3).Range.Borders(wdBorderTop).LineStyle = wdLineStyleNone
                .cell(Row:=3, Column:=3).Range.style = "8ptCenter"
                
                .cell(Row:=4, Column:=3).Range.InsertAfter "REV."
                .cell(Row:=4, Column:=3).Range.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                .cell(Row:=4, Column:=3).Range.style = "8ptLeft"
                
                .cell(Row:=5, Column:=3).Range.InsertAfter ActiveDocument.CustomDocumentProperties("Revisao")
                .cell(Row:=5, Column:=3).Range.Borders(wdBorderTop).LineStyle = wdLineStyleNone
                .cell(Row:=5, Column:=3).Range.style = "8ptCenter"
                
            End With
        Next Section
    Else
        MsgBox "No primary or second section found in the document.", vbExclamation
    End If
    
End Sub

Sub setupStyles()
    On Error Resume Next
        Dim myStyle As style
        Set myStyle = ActiveDocument.Styles.Add(name:="8ptLeft", Type:=wdStyleTypeParagraph)
        With myStyle
            .AutomaticallyUpdate = False
            .BaseStyle = ""
            .NextParagraphStyle = "Normal"
            .NoSpaceBetweenParagraphsOfSameStyle = False
            .Font.name = "Arial"
            .Font.Size = 8
            .Font.Bold = False
            .Font.Italic = False
            .Font.Underline = False
            .Font.TextColor.RGB = RGB(0, 0, 0)
            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
        End With
        Set myStyle = ActiveDocument.Styles.Add(name:="8ptCenter", Type:=wdStyleTypeParagraph)
        With myStyle
            .AutomaticallyUpdate = False
            .BaseStyle = ""
            .NextParagraphStyle = "Normal"
            .NoSpaceBetweenParagraphsOfSameStyle = False
            .Font.name = "Arial"
            .Font.Size = 8
            .Font.Bold = False
            .Font.Italic = False
            .Font.Underline = False
            .Font.TextColor.RGB = RGB(0, 0, 0)
            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
        End With
        Set myStyle = ActiveDocument.Styles.Add(name:="10ptLeft", Type:=wdStyleTypeParagraph)
        With myStyle
            .AutomaticallyUpdate = False
            .BaseStyle = ""
            .NextParagraphStyle = "Normal"
            .NoSpaceBetweenParagraphsOfSameStyle = False
            .Font.name = "Arial"
            .Font.Size = 10
            .Font.Bold = False
            .Font.Italic = False
            .Font.Underline = False
            .Font.TextColor.RGB = RGB(0, 0, 0)
            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
        End With
        Set myStyle = ActiveDocument.Styles.Add(name:="14ptCenterBold", Type:=wdStyleTypeParagraph)
        With myStyle
            .AutomaticallyUpdate = False
            .BaseStyle = ""
            .NextParagraphStyle = "Normal"
            .NoSpaceBetweenParagraphsOfSameStyle = False
            .Font.name = "Arial"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Italic = False
            .Font.Underline = False
            .Font.TextColor.RGB = RGB(0, 0, 0)
            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
        End With
        Set myStyle = ActiveDocument.Styles.Add(name:="8ptCenterBold", Type:=wdStyleTypeParagraph)
        With myStyle
            .AutomaticallyUpdate = False
            .BaseStyle = ""
            .NextParagraphStyle = "Normal"
            .NoSpaceBetweenParagraphsOfSameStyle = False
            .Font.name = "Arial"
            .Font.Size = 8
            .Font.Bold = True
            .Font.Italic = False
            .Font.Underline = False
            .Font.TextColor.RGB = RGB(0, 0, 0)
            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
        End With
        Set myStyle = ActiveDocument.Styles.Add(name:="9ptLeftBold", Type:=wdStyleTypeParagraph)
        With myStyle
            .AutomaticallyUpdate = False
            .BaseStyle = ""
            .NextParagraphStyle = "Normal"
            .NoSpaceBetweenParagraphsOfSameStyle = False
            .Font.name = "Arial"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Italic = False
            .Font.Underline = False
            .Font.TextColor.RGB = RGB(0, 0, 0)
            .ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
        End With
    On Error GoTo 0
End Sub

Sub setupPage()
    With ActiveDocument.PageSetup
        .TopMargin = Application.CentimetersToPoints(0.7)
        .BottomMargin = Application.CentimetersToPoints(0.6)
        .LeftMargin = Application.CentimetersToPoints(2.5)
        .RightMargin = Application.CentimetersToPoints(1.3)
        .HeaderDistance = Application.CentimetersToPoints(0.6)
        .FooterDistance = Application.CentimetersToPoints(0.8)
    End With
End Sub

Private Function getAngloLogoBase64() As String
    
    Dim AngloAmericanLogoBase64 As String
    AngloAmericanLogoBase64 = ""
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "iVBORw0KGgoAAAANSUhEUgAAAnoAAAD8CAYAAADg3AngAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAALEsAACxLAaU9lq"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "kAAFG1SURBVHhe7Z0JmBxF3f97umcTEDAmO109k0SCiIr4vl7g/Spe4Pl6I/rH4/VEfF9RQVE8WDwBIcdOV/fs5uBSFKKA4oEHinKKioJyKocG"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "QW4ICVfO//dXXbOZnanumd3sJtn4/TzP75ndmarq6qrqqm/X6RFCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhB"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "BCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEII"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "IYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhB"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "BCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEII"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "IYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhB"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "BCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCtl/2GpjmVYZqXrWxlxc1nuuFgy/0"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "KuneXn99T2/O4FxvRjrT8wbK1jUhhBBCCNnm6T9ulyCMXxVEjfl+lFwRhPrBQOl1QRRvwOfaQCWr/Cj+m6/0D/H3Mb5K3uWp9PleZUHNhkAIIY"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "QQQrY1ypX6iyHwToOAWxtUl24MqkMb8b/D8H11GLYYfy/dCLG3GqLw5341HQgq+rXejIWPs0ESQgghhJCty/llv6oPgYi7JaidZMVc0qOlMLg3"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "whB+Qw3R1/hBWcWf8FS6u70AIYQQQgjZ4uy1fFopTI6GYHsoqC5pE3HjMOkFrJ6Iv9ONfpRcXorir0wPB59or0YIIYQQQrYMG0tBFB8dqHRNNg"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "zrEG7jNtvLJ59KX1uO9Ge92XG/vTAhhBBCCJlMfBV/MFDJg0E00SKv1UTwSU/h0FoIvt8F1fpbvHkDO9goEEIIIYSQicYuvFiR9bq5BNpEm8z7"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "E0GpH4XgO9GrJk+1USGEEEIIIRNGdVnoq/g3QXUZhJcspnAJs3azPXOy4MLYiS1/S4/d2MLxo+Ravxq/19tjcLqNFSGEEEII2TwOCGSBhF8d3h"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "BUpZfNJcZaDQLOrKhtyDYqVwXV+Iwg1AsDlR4bhOn8QCXfgeF7WYQhwrGXMGFmTmC6Bn61V9XzbOQIIYQQQsh4KVeTfSGw7uxp8UXmZr0fNb7n"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "h/rdXlh/Rsf8uj0Gp/ep5Ol+NX0vxN/pgYpXZ/5E8HXr5ZN9+paIgLzA609fbkMkhBBCCCFjZu6SWX6kzwxq0vPmEl5NE5EGsabSq/qUfqf4sy"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "EUIxslV9IXB9W0jnDuC6JhhCNirkjw2R5Dlf7TV/pgGxIhhBBCCBkLfrX+nqA6tLZ4eFV+G9pYUskPzbm240Hm3VWH9oGIWxyoZGW2CEMEn+t6"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "1mSenxp6CO6/4e16zEwbEiGEEEII6cqsxhxfxX/sOmQrvyu93At11frcDIb7gqj+8rJKfhyo9JHuizZknt+QLNQ4y+tPn2IDIYQQQgghRfhRfG"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "R2dm2B0KrJfLn4Yi9q7Ga9TRAn7eCH8ccgIK/LxF5R7142bOyr9HdeWH+RDYAQQgghhDgJB58IAbciqMqcOZe4EpPf0jthk7coQiVPh9g7HddY"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "X9yzKPP2ZCi3cVNQi99sfRNCCCGEkHaCMF6Q9ZTlzc0TYbVMhmzrMtxqvU0O4cDOEJ2HQcTd0XUo1/zeuCtbpLGxZEMghBBCCCFCX6SfFqj4n4"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "U9aBBUfpTcUp499ALrbdKB+HyVH6VXbNqKxREvMSP2hlchfp/x9lo+zXonhBBCCCFBpBuBbI5c2JsHMRUmyzxveWC99UZ0/E6eGnpBOdSHlCHE"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "IMY+0DdbP9Pzejztojq4V0kl52Zz9grm7RkxOPRwoNKvm2sSQgghhPzbAxEWKL0i6xVzCCixbJXtHRB8Y5ibt3yap5I3Qdj9BgLyXvh/FGGtDV"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "QCMda4G9+f56nBN3pz5+9oPeQT6ir8nxxUG+uzxSKOOIqZ+YWN9aVIJ968gcdZ34QQQggh/44sD4IwHjKbEefOg5PevMUb/TA+23rqzuxT+iG2"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "FhpxJ+IrE2DWINTMd0Y8QvilJ3nV5KnWZz67D89AfBbA/9osPFdcxUQIyr3oJd7M4RnWNyGEEELIvxlh8kJfNYp782SlbahX4+/XWV/FVJeFEH"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "DfycRYQe+bEWMQfhCZfrUh5+MeYEPIJ9Q7l6rJNxDumu5iT0RlfYk3d35vJ3YQQgghhGw37D3cVwqTBUH1xBaB5DAZKlXxbzxvoPsih/6lu5RC"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "PZz12OXN93OY9O5VGyv9KDnCnJhRRHT8TkGk5yNe6wrFnsTbDPOK2OvxeDZCCCGEkO2Bvkr87EDpW7O98RxCyZiINb3Br+p3W2+F+GH8mSBa3K"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "W3LceMn8Y6XO/4fghGG6SbWYOPDVS9DnG4oXjOnunZ2wChGosItb4JIYQQQrZj9j2/HITx18y+eC6B1DQIJb+a/NmrDobWZy5BmOwP4Xh/8TBw"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "FzM9gYvXI5yk6/w6EXvVZJkVc51hNc38PrS2pPTX5fQN65sQQgghZPtkWrWxV6CS27rOoZNFGDKc2o1QV/0w+XlX4diTSZwWbzBiLxzY2V7Bze"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "y4H4L121mvZJHYk9/TNX5Y734vhBBCCCFTGb8aH2XmxeWutBUzIvC2aT2siIXI+5icPVscnghHhGl62IrciYkbiLOKPtbzBsr2Mm7649m+Sn7Y"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "/X4QXpg8LPv4WZ+EEEIIIdsZVT3PV/Hfi3vzYM3jzuZ1Ge6sLHqSH+m/Fi+MWJr1qim9GvagGd4t3NJFLBOE5TD9tL1SPird3Vf6/K5hZj17dw"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "cqfYP1SQghhBCy/eCHyRHZMGfRqljz+6pAxftZb3mUSmHylfwFHdKLJ3vwJX8uRfFX/DB+rxfG70O4x0GY/THrhSsSnPhNpas9pQ+y18ulb+ai"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "/4DgvDITe66wxEQEIj4q+bsX6edZr4QQQgghU5+daydUILKuKux9E6uduNGP4rO9GQsLT5eYHg4+EYLtFrdYE1E1JHP8fuqp+D+tl03U6nuabV"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "KidFUm+Nr9N02Emb6lHDZeZH3mo+IXQKDeWLwgRMTnEgnzMumNtD4JIYQQQqY2fqT/D+JqXXEvWgNCaGidr+ofst5y8av1gUzQOXoHMzF1jRc1"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "drPOHQz4iNP/BCq9LV+cZcIM8b54h6qeZz3mElSS1wVh445i8Sj3iDBVfLoH8Wu9EkIIIYRMUfqP2wWi6pLioU1YdZn05v3Jm6v3sD7dVBbUIJ"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "T+6hZUIiQbj/hheoh1XUigBt8IQXdrvjizwkyONeu2OAP4Kv4wxOND3fcIXLwuCOOjPW9jyXolhBBCCJl6+EoflA2TFomfkd6z4623XHyVHAb3"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "a929eRBsVX2pF526k3XeFV/V3wm/92erd9vCM2bivc4P4/+1XgoY8GXfPHM/rviNmISpH/YryYHWIyGEEELIFGOPwekQSOd0P+7MiKx/lav6Jd"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "anG+kdDPWl7n3zRFjpdRCC77KuewZi9JOBStfkijMjQoduLKv4BdZLPjOOmVlSyXczP46wmoZ7Rlxv8EL9TOuTEEIIIWTqEFTqr/Gj9M78oVEx"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "6c1bKqLnh5630bdencjqWYixle7ewWEZ+r1UxKB13jtz5+9YCnWjUJyZoWX9g16ONJseDj0xUPqK7itxZT5h/CNv5rHFp3EQQgghhGxT7LV8Wh"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "DWIZ5OahM47SZbmSSP+EoXL8II9c4Qct939w5mQ6V+tT7m3rwRKgtqiMMFQZQnzkSQmqPZDpMhWusrl2BW+krc223FIreBMBsbyyo5ynojhBBC"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "CNn2KVcXPSdQ8T/z575ZEyGk9NVyrJj16iSo1F8HQXifc6FDVXrz9O+83VNlnY+Lclh/ka+Gbs7t2ct6Em8v98f7WC8FbCz5Uf0IP2qsyU7lcI"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "QnJmEqfZ/0flqPhBBCCCHbNuVq/Yvdz6CVnjixZIH15mbfgXIQydCqq3dQetoWb/Sr+uAJWcWq4k/4stgjT5xBBErPYteTOwSz4jg+s/vJGRCW"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Sl/0mMqCmvVJCCGEELKNEjV2881pEd1680RM6ZXe7PRZ1qebWvxsXyX/cIZnet/0ldMmahPiPQanB2F8eibOXIszzHdrfJV+0PoopE/F/wn31x"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "fO/zNhDm8IKvEx3HKFEEIIIds0fjV5TybKCnqxjA1vLKn4XBFX1qsDGQLVn83tHcwWcnzeOp4Q5OQNhH15bk+c9CBG+mqvtmhX66WQcnXwEISz"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "tnDDaEkv1bgzqOrXWm+EEEIIIdsYswYfCxH0i+J988REQDXkPNr3WJ9uHqfnZb2Drrl5Moya3OBVBot7BMdBEMav8pW+xzkn0CyiGBKROthTD9"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "y8k3ZAWOd0Fb+yslfFP/PmLpllfRJCCCGEbENEyctkPzv3sGerDW0MQn3jDl16xWRfvGxBh6tnbSksja3TiSeMj8x64Rz3konM270ofql1XUx/"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "vA/u5c7CXj1zj8Mb/Ur8iV5W9hJCyDZFdXCvshp6gRcOPtF+QwjZvthYgsg7pXCVadMg0kqqcYxsw2I9dyK9gyr5pbs3TzYc1nd6Ufpy63riqQ"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "0/xg/1We75dRBlMpys9HKJp/VRRKkU6S917ek0Q7j6Zq8/fYr1Rwgh2zZ4kQ1UvMyPGn+S+suvpr9FW7DQi+pPsC4IIdsD0+YsehIe7tu7i5nG"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "xiBMV3v9gy+zXp2gwngFwnMcdyYia6n0qJ3leQO5QnGj5/Vt8LwD13veUti38fcCfPcm2M7WSXfC+jP8KC04W3foQT8afLt1XUy0cDfE+c/usF"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "rMiL0k9rzhPutzm6APaVFS+rhSFA/hUzvN/BYf1xct+g/rjfTKjHSmr+JP4IVgSWfaxkkpTBb3qeTp1vWUIVCL3mzib6z9vpoWL/VV/XBv3sLH"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "WW9kqqDS5/uRviKonWTqZVN/yUswDPXBed6Mxm7WJSFkqlNW+gsQKhu6DtuKSFPxj72dhyvWqwM5NzY+NZsj1zZsa4RkusqP4ndYxx1AzD0Gwu"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "4EfN4P29g0fHcPRN8l6zyvN3EGENcPZ3Fw9FSiMitHyXm9bo3ih+khuUPRI2au80A51P9lvW0TBOHgVzZV5MgDp8m9DSN/0y9Yb6RXasO74kXg"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "kqB2sjtdZXuhsP4W63qq4Gdzdl331GK1E6XM3O2peD/rj0wFagOPwQv5KUbkddT7+F/qi0h/1dsr/4WcEDJV6F+6i6+SPxVvIyImvXFLNvjV5K"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "PWp5vqwqeiArnTNALtYZj5cfpSiMHcvewg6N4B29Aq8loNvz0I+zz+7t5rFh2/UxDF38oqrba4mPsZ2uCH9f+xrouZ1ZiDdPqNeePtCKtpEqas"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "7I2/73nnl63PrYvsCajqP832MhSRWmCo9NFon+N5BcPypJPZyeMh9C7MykZ7utpGUy16s3U9JShXZUhPr3DfU6vJ/YnY01+0XskUwORvlNyYW+"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "+j/kae/sXrj2dbL4SQqQoaqAPxYD/gFGatZhorfX2fGvpP69UJhNXX0AA4egelt0tv8FX8ceu0Awi4XdZ73hnt4q7d4G497Iv4u6sg6ass2TtQ"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "yU1OsSeVWagvFxFnnRfiV7X06j3aeW8tZtIxXYW/X2e9bVUCNfQKxOdfXfNXzIjU5Bavol9ivZNeyIQeXgLyXijQmKr6m6zrKQFeVo5EOe9hcR"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "YMYhAvcD/z1KLIeifbOEGoX436HPV+3rxsfB/quz2V7m69EEKmJsuDQMVnuBuoNpOzakM9XLiqdPYp/RByVznfEkVUVZObvF2HcodKId72gdC7"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "p13YuQzu1sH956zXQsqyn180DIHWXqlJj8QQ3lyTj1inxTw+no233Iu69+otkbfhn8miEOtzq4F7O0qG19xxbTeJu+mR/Iz1TnphexN6c+fviO"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "f4R9mwXvv9OAxiwVeN+yZ1gRWZWCL9PJTZv2dTNhx5aqZx6Gu9WfN7egkmhGyr1OJnQ+jd3FXoZb1BK71Qv836dFKuJh9Fw/ZQ51uiCKrhjaVq"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "epx12gHEmw/h9vF2QVdkcP8Q7P02iHxqwxU0XBBobgEaqOQqTy3pqTeirPTnIA7XmIUp7WE1LQtztV+pH2i9bR1mx/1llfw4qBUJ0zaDKIQw/o"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "E37yROru+V7U3ohYv+C8/Liq7nXbcaXgTLUXIktxeaImQ7I3zXzMHsyE8ps6gzVHpsT0dGEkK2XSBaBtAIrc+EWPvD3mIyNKP0+d7M4RnWayeh"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "3hmN3bnu3q6GDAmu9GbrZ1rXHUCwzVrveRe5BF2Rwc9d8Lu/DSaXIIwPCKp6decQpgi2FPeXfsI6LSYceqIf6Wu6i2MZzqr3drbuJBHMkb0R05"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "U9Db+NGES6iu8Net1nkGx3Qg/l+9OI84au9UKrSU+w1BFRqmwwZBunHCYv9KOhv5nRGlNHiOH5l7orTC71osXcYoWQKQ0qZFTMv8weckfFPWKm"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "AlhfiuKjrU83Ufp6hHevcyjANHT6OzIkZF13ALH2fIi2NS4x14P9Af6faoPKYaCMBuwH7tXAEmf9Z++x83s62aIU1RtZOAUCygyBJvd5Ufx662"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "2LU47iz3TPX4dBsPiq/kkbDOnG9iT0Zix8nB/G3+952HbE8Cyo5JFyNPQ8GxKZCkT15wbVxndQ/z2AsroedfjdpbDR8Crxk60LQshUJVCNNwYq"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "fbBQrIiZRipeIYf8W68OZMPlZLF7LpgRVRuCSlooeCDWjm4Tb8Zkzt46zzsCf38Yf/+2/fem4bcTIfZyhaSAuLwcghP33D60LGmQri0r3ZO4kd"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "4uNGp3589vsSbpoeJlnndAYL1uOaqDIYTtL0z+ueKWCfi276yJ0IvicyUMGxopYnsSetLLo/Qd+ZP0zbPS9p2Y3Ke83OjP4nHsfrwg2XboP24X"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "GanwoqX/YTZK5nAtIdsDy4NSlCzsPklfKm8z3+wM69FJX2VwbzQAN7p786R3SF9QtEwfAm0nCLU/tos3fC+raz9mnYm7ObBT292Jwf8G/FY8X2"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "+v5dNKKj3dOffIxDO5zJtxzEzruoDhPgihC7sKPfN7ektZxS+wHrcY5XDBiyBqc1YII1+VfjgTva5G2zTmq8pR47k2uMlFGpa582fJXEpv1xTp"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Pwl7d0lvslxDxOvsuF+mGthfNp9JF3oDZXOKi8S7mUYFveObQQnPwMcLe4FDvTorV47fIA5LUXzxFts8uX/pLiZN5IzpsS58mjewg0nHZnrK/5"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "OCzbvaCRVT9qQMzj1jMvKuGBFzkk5FJxpNLCXzjGV5Y8vsJN235H2z/jDpOynPxtjYY3D6SJxMncYtq8gWZlpteE+It+u6ChXp+QqThwIVv8F6"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "dTDg+2FyZNY4tIuGrJGD0Cmc/waB9kIItQfaxRu+PxO2k3VmwP8yl8+5BQt++xvsadapE9zLfrh3NFRtcTU9GI2HfBX3tAK3rOqfhL/1biHVNL"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "l/md8omxAv34K9esiTKDnCCs3OOCHOJRWfDgH+I/ecyizf/Jo+VMKygXZnVmNOEA6+Ouhf+Iagsuj1o0zhO3OiCho+S1/YeEafSj8I4ZAgvt9H"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "fH6Cz+/h+vP9avruzT6lA41r0J+8rBwNfQwCpeFH+mxc42cQMz9EGTjJpFEt2X/UtiAzF+0a9Os3dsTf3MPgG/sq8bM7Tj6ZDKEHcSBnj/qh/h"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "+kz7EI61uI+w+zNNLfQ/zTcjU+LAjTV03YFhiPnT8L1zvHxLfjPvB8KH2XpFsQxX/rnOvadJM84vUP7WND7I1IPy2oxEhjsdY0x//9g/8NwbBp"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "bi8aTdkapIxnSvbIlDKcTcnQS/xo6NDspSr/ZBoZmfAr+v2IZwx/Z6Es/NiUOZXGCOvgcnXRc6zT8SPCTk6dCOP3ItzjkDbfRr3yY1zrp36Uno"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "06dQh5+wl87i/lzfrqCYk/8iGnfC58g1dNnoqq0PaoLp8W9DdeWq4OfypQjVPxcnomfh+5PzkxZ1NZr78us2ZYg2+cHs7fwzrtkQ/34aX/WX5U"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "PxB59CXE82TJm6zMJmeZ+67qw4Nq/TWbfcRaZUEtmBXv51cTPAPJMO4N9Ud8Lq7zfTwTDTzzhwb98UuNwLX0FZUzpJ3E3TrdBMQ50uU1Ug5H+4"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "EhjcqVhS9p3VRa6izE4X3I97rktX1ez0T5XORHg+/3CkfGCJlAIMzeE0R5Q3othsocBfWPsvGw9drJ4/Q8cePuzZMGLr5BKhTr2gmE23EQaOvb"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "RNsq2L7WySjw+6747ZJW901DWMvwmS+qzGqzGI2ZK76ygCI+t5ceiWn96VPw8N7rbvBaDNdBBf8nSSfrdfJB/JEnF7vjJj20+h4PAicI46/l9t"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "6YvK+fL3O2bKhdsaJkbRampG+LyYbNSq/w9j2/bIaVq/qLqAxvMEJTBJLkx4jJ/zCVXou8Osw0nGNBtgep1t+FCvcHuOYjJj5SFkddQ/7H90qv"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "gxD/sTR44hXpdnBQO60z/uYeTt0o89e8GW0rkidS6MkG11X9bjNvSiX3uNNHTOKfiXQ/alyKdDqqeHpFD/TH+6Bxus+5otzkR3Khp5KnI43Od5"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "ebbFi3FCVftiH2hBy/l9VHYri3EZP/T9xYCpNvi7uyrAaOGngRwAuWSZdmWohJekh+ypSKOIY4HC1SasMV80Kq5FjEvDJn0vO2ktLHoczNtT57"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "RxakzY7fAVH1HTzzdxXmnZkDKfuKpr82R8jNTB5vQykE8T8xv3yeJC9wg8ZhVc+DEB5Ceqw2IzfN/IKoMb8DCKRv5Yd1Ml7Q9YB12oUBH9d6Pf"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Jervf3kbJpPiWsptnvVYoym16CNPr8tP76njaQ3th9eAbK+oeRbj+TZ9d9HUlzUxYeRlk+rVxLXyxe8X8cRPjNVc6Qdgj3THONFoz4VVK2Jf1a"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "/cCkPlDxdaYHcZe4X1ad4xm4Lrt+e75n38mxnOUoPtL09BEyaZjTIpJvm4ejo6JuNam0ZbWszLnJpy+URilH7JiHDZXu3vlv2BBsO0Kc/c4h2L"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "6Jz9zhmDWe9xK4uavdH8J7BFa4WbFf0weZh8/0brXGGfdRTR/AG2f346r2GJyOB/snWU9gezitJuk4tB6VxZaboyUTrEXgmGu3xQcVIyrYy4wY"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "qui3o1Fa6Z6PBb9Kr/bUop7PaEW474JfpF8zbVtMyojSV0vvDRrcs0Yq547rtpgpo3Aj+zeigreXKaay6Em4zim4phWcjjQYZfjdlNP0fjQeH/"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "RrjQ9ljXBb/M09mPK8vGP1+UQJPRW/AELybPhZn8XBlS/tJuHbhiVK/+xF8f+O79iqAR/59/EszeV+265j7l0n4rKk6jp75l3uzL6Uv0d4PQ+F"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "QpygAZZ8EpMwm2a/k17Eqn4t/v5Hli7yfdt1RwzxghuI0V951cZe5gJR/QmI++lZOnWp92yjjAb8XMlX478XosZzcU3pjbblrpe8w33YOMHvBZ"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "70dHUB4nDYHLU3Kp2sIe5Iy2MgInZFePY0nGY88CnTNSr6tTYoGd04KT+spRB6yeet03zCwSdmQj1dafImr1yMsix8I9Ki9HJ5QbShFRM2noHn"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "7Kws3cRvUTkQwz1n5WWF6WWM9CJ8bzfzb7lXEx7yXSXfsVcaAS/Dr8L3D29axNdiUpdGySXe7PRZcHNGVna6lS/8buo9fTJeDKr2MoRMLH0z0/"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "9Aobyvp4dE6bu8/vQp1msnaPDwhvKzrOC2+TcNe3oPKpOCYd9s2BZ2e6tYg4B7aK3nFfoT4O+wVn8t/i/AZ/6cHbztQ8D+2R1vVAwqXtrLnB00"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "aB/ZVGm0hdNq5uGunyjnS1qvkwrexL+4qUJqjYv8j8YgSoaNw12HakiHP2SVZqu7plu8eYf608ZtD1ihB+EolaIjPKXvQANq5zZ2K3/WJCwTXp"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "dV30CGZqQXMrsfVxwKTPJI6btRcV+U20hLuCo+YzKEnohuNM439SZOXSbXEYHSeFR6csa8D6L0dEfJL5yiP5vWsEryV5zKvaChut+dz/hOJStl"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "6w4Tbg+UlK43y1tneHo9hP7VCPPv7vR1mTTcko7Jt6ardHfk6ykm75wvNC6TtIT7MFnied3n/kl6IO2y3umxljtjWXwRxu1+v363DdYJ3A9l4s"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "UVztDGUhgvQHqenMWjtRyZ+rxN6OkTc8NCmS2rtHBDehkmR5m9LMuX8dy3TWeVroZA/UZRnVvGyyvqItTZksZjfD6yZ/sO2BX4f21nOZN4mPib"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "nuNWrNB7yF0nGD//QBr8amx1GsLKyvLXWqeyEDJhZIfzSwF1VaothoKLCmNp4eRdpV8Bt3hwHAUcD6ScsdptBRdE2dEQbKOGbfHdL/Bd1/244H"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Ym3J7b6lcMftfCDrbOHAz4pUgf71yMggceD+6t3qweFiJUZMPpxN1zNsrM7//aIscJzZ2/Iyq03zsbYflOxavQoIxs5IyG9JtZpeMoDyYMfbGZ"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "WNwDxULPmmlsu5S9djPx0yvKkc7fukMtiZBvds7hGBuCEZN4Ffg1jdLEC72gGh9gGiIT93b/YzVJe8nn5CQzpNQrKnl61qC5nmUIyDD+y0j53f"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "W0mUjra819tbs1aZhuKKnkG8ZtD3QRehuyz4Iy5TQpZxCjKrkMabsmV7znGtzLopP+5GU2mk7g9r/h7p7cZ6hnk3KCMGS4PozfaoPvAG4LhJ4M"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "iULMKO3IR3M/Eyb0+irp3hBo12RldnPuG9YUSWGywHlGuFqwO8ob6rTNSONNdZItT60maW9+H6PQa5qk9VjjhXtW+r5yWH+RvRQhE0Rt+DGoVH"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "/ZvdJDoVXpukA19rM+HSwPgqpe6haNppJZ220vNoiyAELtp20ibYNsp2KddAV+Xgk/HQs5EO7v8X3uvnioEN+OeOeIVOkZ0YcjmOJtImTej/QA"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "FQkbY5I+DQnzIOtz0pC3bFQg7t4WVN6472ta54fg//9BRbba3ZODMFAZoVLvnKTsoCehN6qsSNr3Ukni9+yt/Ov2Um2gLIbJ17L5N93KtliWH8"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "68L4rLJAi97HB5mavo8ttuvaaXpH9Dene+LC819lKF+GHyqSzc9rBt/O08uSalCkRCXlzk+QmT3/b6glAs9MRav5e/8/Ku3cRtlhajv+slDZt+"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "pVcvp9clE8c3mfRxhiGGcETISP5K+TGfBc9HJpyuzFuIBDcFQq9prnubQKFnXqrii3sTeZLW8kx2yy9Jk3Q9yuzRo3v2zi/LQqpsXly3a4nJdZ"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "rW/luef8kjkyfjFHqt4Tav3UNcJf3C+ld7fUYJ6Q2pmKJ4VdfGEA+5H8XnFe6jFuo90OjdlScoUIlcJ8Mm1rUTCLFnwW5oE2j/WON5PQ/7wP90"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "WNIahhi+exSCMX+1rwzfKn25s4E1968vKtoSxrAvKiGIj+yNtC2MUSYPvcxdir9btCJwIggq8TG4nmM1sFRmku/6NOvUsMPsEx6P73KEhqmwZO"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "jsS9Z5Ib0JPTHEwzQSCF/p9VnjV9RYwmRoS+ZNma0KRiNvxaiM/9k1HyReZkgPnzKR26SJbXx7rZgnUuhFp+7kh/qsrg2m+B9xI/Fu/t+exy1m"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "ynBDepFfaa9WwPJpeBYuyOqF9niYMvOwrxqjesiDavpWpOGjzroky4e7zYrmHugu9Kw1BZP5H/feLd3azaSj8Y/nQ8pgl3yX8qL0X72Zx47Ob8"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "PANPy2PLvXnDBMeTPPwq14Ni6HSe/iX0xPoynvef5kAUp8glzDXmwE/N5F6OXdj7mXR2VxgQ1qnEJvY0l6a7P7Kip/cu/NaQi2R9UsCjHp0Wkm"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "b8zcyh95c1oWwlT0S1D/PGhFd76ZsiHXM+UVz4iUD7mepLPD/ShDmmXxGqfQE5PyNPJMok6T+HS5Ntygnbl0XAt/CMnDD+PPoNB23xIEb0/lqi"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "4Y+pSw9KcD5RIU8v/QBjyceBMuBkLsUBFkbQLtbHz2PJFbgJ83wh5qDUcMovHX+D6nV2+4DxXCadmE5db4N+8hQeVUPGwjmJ7Bosq+aVKRqHjF"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "pB4RNe+kHaTicAsOuSf9KCrSd1rXI8hK0qyia78HqQBNZXSBq9FppyehZ9IheRBuv+tX9BcRn8ODqq7j87pifyYef/Wq8eitO/YdKJdC/dWskn"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "f4a5qpdPW9EGTfL4XJ1yBcPonrfw55sgxxuSr7vei5gElFPoFCD/d8EJ6VVfn33fQrvUbpiRJfxPUwND4n4PM3WX4VNEBmvmlyqvTk20u6yV4A"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "c3uBIQhunt5+HFaoq/h+hTvuEu9hWS0LsdKdnoSeXEcltyKei/ES+hm/mhyF9PtVV38jJqvfpYxpjfL+KSkz+P/yLP1y/Jt70/dnW5aMBvXbq3"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "H/j7jv3/qtpqtLIsykHpkh+8jN33GHmVJW4nf4YXJeVt4c18YLma/Sf3nhwo4jI/F7Dz16YghbyqO4rZ2Cv+2iC9muxjIuoScrs1X8D3dZtybP"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "qqx+VunpcHuUH6WHliL9JTwjZyEP78vqy6Z7iac8u+kt+P+E6eGSUaulgzBBHV2QR/J99ozcIc+m9GIjXz8JG8D/pyIe2SrgXP9iEobJx/EJPb"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "kfqdOi+GxbLg+HvwX47trR99puZrHjA7LIxF6OkM0EDSIeugtzK6ammYemfv3IijUXRlAklzgLcVY5rkSFXHhWKoRYGULslFZhBlG2DtZ9pVcb"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "8PNshPXX1rBseKvweYB11gEexq/lvtUjHVBJf8U5Z6QVmSRsxE03kSAVRfzIZB6Jhuu8DGl/uztfzPVXeJWhmnU+Qp+qfwi/rcnctPszYd1eDv"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "V/Wee5oFItFnpZhf0AKrf/6xiuiPTzUEH+0RkHY+Z7NBLp6JWJUWM3hHd5YYWaNQS/R8X/ZrO9SzuyN1kYL8vKQUGDMJFCT+ZSRvG33C8a1nBP"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "CPe88mzHwga8MJSi9EtoTB/OFaiSD7KgqpJtL5GHrK5Eucl6QRxhQIidY522sNEvSfxzry09sMkFuM+uxwp2FXomb/WNSMPRvZPR8Qr3952sIX"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "f4a5r4D/XvOubdVhY9CQLkovyygzJnNhVPOlbxQ3R/LyvnjjibMjz0EMrzF6zzTvqXzkZ9nFNmJAwIrSj+WPtzgt+6C70s32Wl/29EaEJwfxkv"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "5p9FWfqYFy3czQY1LqGXvVQVCSeTZtd6KtuuaDQbfb9SPxDxutqEIfcu2zFJHkayx+ZodsBzhTAhKvPyV+IgolhfLMLbehtNpf5ilK+fZOU0r4"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "5GOFmdNXahJ3ldHVqJsv759t0l+sxq3Pi3xfGXOPFccTJByGaQeKDuLGwQxVBBo/JbiAomV+CUq8m+qHjv6QxLCq5ZzHBRty0eHva8PWQeXZsw"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "u32t5/Uw1DSaRz3v6fB7VWtYLWHmVrYQCDI/7SG3uDDf/cnbuct+R7VFspXBxV3T1aRNuqEU1RdZnxMOKrv5SH/3sK1UiFHiHjqWlYlGILoEmo"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Q1tA7pdIx1nUux0JPKFGnaNterFeTH2yHkMrcd/iUeGmJUv806N8imqHDvWElnDfniq/Sq7qd8yCIWiL2ifJxIoVeLn437/WtuI4C0Qnpe7c3q"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "7E3ahB1GM/uCOcIQq0FwVdPDrQcHcgZ0fIE7HkhzlazNhHknKPfvzNK9vbyJmXrgHi9KRnqQ8igUeiOiKXVPw6hIOjZucz/DYnIPjdxN33F/b4"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Ll7IeJMFXyiJRr69ywg2yqHebNzZN7wL1Hybndpmn4UfoOXAPPiyPuUm6ryQ/byxp+K1yMkT3n6V99NXjQDrXhws2Yxyz0INpRFn6Wf6KSKQf3"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Im32tz6cmN7QKL3FV40/+Kr+zrx9WpGGByK8VbnDtpJGkb5StnixXtzg5RZ5eH5uOJJu4xZ6Zj739/Lmo+Je35blsesZyfILQm+rnYdOtjP8MD"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "kChcu9+KBpUtil0utSOeP3haiccoZtUyn4Xc+MXed5b4fQGzVsK8IPwix/c+Yc4Gc/2MrWsJqG74+3zjqJ0pcjrjlzu+Re9CPSkFjXbmrDj0Ga"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "nZTNz2gPo82M6NC/tj4nlhnpzKyHwFUJZ/mE3/+fdd3G8gC//TirdFyNrewXlZ7nPb1g42xQLPSk0dQQ1Y2RyeAdzI6fjIr7FnejbfJjHX5/u3"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Vt6Avj92UCyxFvW4Hj3j5unRcju9qbYdK8Sn3ihB7KAUSSPEN58UZZCdNDrPN8Kic8GXHOH/bOyuWpXs6Rb+Vo6Hn4/V9O0WLqA32X9/ghd0M6"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "o/4EuLs9Kzdtfo1/XLuSt4BmE8VCz8wXvsF7fM58WTnNQ046cYquEf8XyFCz9TGaXU7pR/m42p13KHMi9Nr2eYOQeQt+LxKHD8uejJlrOec6x6"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "qNvVAO/uKOu3leVsiLZBZOBn4rEHomvg97Xba0ajJWoVeekz4fouTv+WVNtm7SDeu8gI2lvlq6t9fl5I1SlB6PMrHBXS6QPtV0NeqDnvbgk4WF"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "COdud9zleTPfj03oSRyUxCH5gHXeCUQo8tiuGG7zb64rYfSWX4R0YWMJhe1n7sLWYjLcEulfeLO+mX8Sgdlvq7niqj0MFNpQ3+m1z+dxAAH2lX"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "ZRBqFXeKauC/jbAWF1LMZoGsLM3+ahkj4L6XK9u6LNGh40AoXHtwmlKDm6cAiuaabRSW7w5jbmWK8TRlBJX4MK6Y7cxieKbytaXOLLcU1mK4tM"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "FI6yTAjfHtSKVmF3EXryXahvL7x3lBuE8Sf3PbiE3oAvu/e7hTrMfK9v6KssKBbrIwxMQxwbuc/JhAq99PBss9p292KmN+9Or8uJMoa9liPO8Z"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "D7eYRlaXCxnJRgfYyiFKLsRo3cXmCIsB9ap05Q5k525xfMDN/q82W/RuvcSbHQk5ej+FfWaScQsLg/xCEvz/DiY06LyDmC0ExDiVHmXP6bQi99"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "r3VtQF4fKWXR+ayYe4jxW3wF3F2I8nqR2+ILEa/LEM5Kdzj4LtSPtk+hwW/5Qg+ioYQw3YtHOhmr0Ms2Hc6b4y33beLsHkIdByg7Z2b54i4XiP"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "/1Zu5jT8jLbJIzjIrwszI8RqEnz7W+udzfNm+4leqyEOXg5+7nU65rwu3a601IV6ZVB/dCgbzBXcibJg/T8Ho83MXH3sibkdnzy9G4yncyNNdl"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "77xVaNIhztq3VZG99LpuitsK/MyCzbd+Rwm8pq0rmvM3O30KHsI/5zbQ8nCr+PSiYWwBb/yHZA+yo0JqNalMlL4Xb4ET/gYnPSe5cUBlgkrzD3"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "irfWNZVrHJ/octVq6kL0Zj+mFUamvyK3GUnbD4aKuuQq+qby48zkylu49R6E3DbwucZVEM6QH3P5ftIKyHrsgQobtXNAtvooQexM3X8xtZI24u"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "yxNn7fhyfmhuWGYI9TrnnFuIREkf9/1KnqcQDnK2bfpyfPeyUeXGzAdtvBTpEbvzCyblLkrvhxgrbPzzhZ6kHZ7BapL/ApgJvVNyhR7SBaLqSO"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "u6k7ln7Ii8G5PQK0Xpoqzhb49vq8GvKfNFVhRG9hyiDLzEXtaA7wqEHsJU+uReN2aH27EJPdkgPk94ZWm1KjsLemKAsPxp/jMlw+P186zTnkBZ"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "PdWdb1l4+HuMQk/ipv+8Y9HuDEbo6V9Q6JFJBxX9hyAuHrSFym1ZY3lLWQ0+33pzkvVeyUPhelhQsVc2bcabxxrPe8F6z7unVZDh/9UQbD3vM/"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "eo5z0Nfs5sDaPd8Pt9CDN/qHBO40l4CK/MbSTMfeoruq1aREOBN13XUHa7mcpwbWHDMx6iVPlhfF6uQBkx5L9pOB3WLe6md0a2Nzkm9w26q9CL"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "kuu9/qUjh4x3MFaht8fgdKTnYKHQkx6pnt/6cQ/V+CP5jd/ECb0giiFQ8wSl2d7nF66FMy5wjwfn9g4iLfH7zc7D1CHWcJ3bctPPmDznjvJiyo"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "xYez3QZqbcyJmpMlzpplDoZdc4xTrtpKvQOxnXz5nfJ4xL6OnhLM263PtmWfY8epXRC0HwXaHQM/Oriza5b2HMQi/U/5cJFtd9i6jXt3tqaMIO"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "7cf18gWSXE/mxo0BSZvMb3v88X9W54xZ6EnZkQ3ErfNOKPTIlgKN07KuIsAUWjQurlWJTdCwQqSc6a4cpMDq2+Swf+s6F4ivD7SLMnx3N2xf6y"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "QXuC3B3TtgV7aH0W5w8z1YvrCQsxPNEvhCoXebpxYV9wjVkv2RLvc7K4NRllUopTBean1OCEHVTG6+31YaBdas5PLM5cdaFvY9RROtt7jQ8wbK"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "KNvH5Qs9KdPxxa0rDbuBxuML+QJsAoWeSo7Jv46ZZH6JN3P0/Kw8IMAPK2r8Edb1XnWwo0cvG/aWXv6ivG8tH3nm8mdNRKuKL/Rmzc8dsu8u9J"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "JTrdNOehJ6BXOGxyP0VL2ePQ9d7n2zbNxC7/iis8VbGV+PXpHQix/oOqd5DCC8nOlGcn0p18nPrdOeKMnZviZdHeUsq3PGJ/TmLslfWU6hR7YI"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "sou5zAdxFrSmmQptXVklR1lfTqbVhvdEpXpF1ji0hWEaM326t0fB/D4A4bUjbIlDlN0OK5yThN+fCkthq9v9t9t6z7sM7goP5Ed8X4H7yd9o1z"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "z8+t4+lRQf7K/S56Mx+3txz4i1rIejcN7TWDFn23YT8hNhKEPlSH/WXraDLS/0vFJZhi1zxQoqUZXcH8wa3N+6L8YsaEl+7CzfYhMo9Pyw8ami"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "OXoom3f1uXrhOlguG/fq3Oc7u5dLZBsa6yEju9ef5Db0E2ZStzQektXR9sodTDWhB/dHIs3XORv/pkkZljwZtyFf8CkLxuxlDQi7i9CLT5g0oV"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "eJ34E0yVnQJ+W8IfunTtwcvUjm6OU821LX4iW9sD4ZxQEBysHF7mdb4m7qHAo9MkUJB/dHYXTvrdY089Doe2QFnvXlRBoruFvdWejlQZRhsuQj"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "1mkuEGG7QYTd1C7MIMrugjlPxFjteTX89jn4u7rdX7vB3cOwQVjxknvgy/5x5hxMV8UFyx7ClbDCjZPt1jVdNse0hgceD/5FciqC9b55hLqKCu"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "xX7gpsPOaoVJuGig1C72c7Rcc7N31G/m9poYcyGb85214kJ95SwYbJcC/nvqJCfxOug0o9rzwgrAlbjCGbJaNhdMZb/JgV2h+yzvPJVvVd404z"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "WNbAnNae7kG/zK9L7srfcmKsVlxu0GAeLXt52suPYqoJPQiat+H7+9xpjjjL3nBRsiSopm+VLVRw/XfmGdwcIFsGIQ4Hjv6tYT69x44WEQh3qw"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "k92c8Rab0it54z7Ugy2HXvUTlasja852Pm5ayEtpRk/q0pA67nEeVCJavN9iw9YLYEy5tbLtfI8pJCj0xNUMg+nVVi8sC0FzRr5gHVf4Q+Kjyb"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "0oTlrBiyhw5/dz1JAmLspe3iTAzCbM06zzsSf/eJ4f+dYPtA3H0ddg3+zl1wIYbfN8DdJfh8Lf7vbY5KqBcW9oQ1hZ4aeoX14qQvrD8D6ZdtAu"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "oKp9WkclDxVTI/0HrfLBCmbJIMsVqQvz1bXqXaNNPwPVSuDI2aIN5kawi9bOU0BEuWVw6TODceLpuziwv2dszOCM4XTGITKPTQ8DwHjcjNuY2m"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "GZrSV3hRvWAFu6w6jo8qjLPsoxclHWdHywa6+fuhjdW6lBvkDRrky/MWxUw1oTcdeRJE8Y3u513ye2gDysnp1nkhO9dOqBQuUGoD19hqQk+OxE"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "RaFcwFNul1j1frsrl6Vb8E4VyH8n1ptoG8Wxiinvx/iKP7HG4xeXbC5Lc7zCzeL9CbsfBxfihb8AwXvFiZZ4hCj0xFlgcoREvyh4jEpOA3UEEk"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "i62nHDaWUHnFzrCyxuryaT2IFwixg9tFWtMg1G6HnQw338TnlfhcKQLO5bbV4OY2uD8an/ln87YjlVaof+l+AK2Zh1/f1zdbdxxF1MrYhJ7MWT"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "KT419gvW8GG0vlSvz53EbOWJa/xYbKRsKQ4/FU8ltUTLdm3zvCg7tyqD8t17aRGGGrCL1dzRDk94qHrnF/KlkDO8kL6y8y28zMjvvNcXT9C58i"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "QgBle0XX3q0JFHrmZAyVnJG/LY99LkN9Tja/rq0xlKHXqH4oyt2DzkbImIjF9G5ZdGF9ZUjZV/G5hWV/pGwUmOSTOWZNr0AaXoS0yXnhEPd6vZ"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "xSYGMwiqkm9ASUw7Oycuq6X1PeHvVV+oXCnmQ5bzvSP0P8rguqi97abZN5AeFvPaEHgmr8jSytXPctZp6hP0LAvbTj3mVj5DDZH8/rX0wYRqih"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "/KrGMucijqixG+L4j/yXISlXstgoOVc2IO9M6+XTvEr8ZLj9dlaOxL0rHNxLVudQ6JEpSFXPQ0G8MLcSNGYq4XV+NXmP9eVGNiYN9dnOAps1st"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "/K2+G8CURZADF2fLtQG6+JEITAW47PLqceOFDxG/yoAVHieICbZh5+/U9pGK0vJ+X+oX1QGfw1v0JqMbjBg3+bDKlb7+OnNlxBWJfaSsphzYZS"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "8jjP4EbphwM1dHNQ0XWZjwgRcE6uCJD4K32BK01Q4W55oQf8Sv0DiNe6/MZHTPwjLZRei3J8La7zS9hv8f/9I785/bXYRAo94IfJewKVPpovMO"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "V+5Dc5W7RxTJ9K/x8axreVI30o7uEnmZui8ovnUunlXrh89GbJELu4LvJJ7tvhbyQ95DPPjNsH4O5qpONH7NBYTg9lVg4h6I7F1UcffQemotDD"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "7/9tnpvcvJNeTL0GL9BLTA+XzJGUF4vKgpoX6adJfYt0u0LO4M3Kh34EVi/uwTXX3apCr5z1fN9q4uzyZwxpovR9+Hsxntf3yxB2n+z8UJWteq"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "TMtKaZ5K/5/1/lqP7Z6eaUi01D/GXzEie/t5eNpuH7rE69Hc/JEPLyfRBnb/XD+nvwu5w3e1NWJvP8i0kYJg4UemTqUa7VX4yG6Z5CMSMFWMWr"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "vFp9T+vNjZkLlLPhJCoLOczZuswFgmz6Ws871SXaxmoQeL9e53kH4e+OnqWuyGkWkf5m16ErpFtZNnzNOd6miZnvFMr8j9YKLMfEjUruxN+b/3"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Bnc2ZyJkdLvqKyDVHZVvUXzWHbHaYHIOo+j8+DIR5G5meWIv2loJb3coDKKUweLs/q3IYH5WOrCL0d5zbmoAz/urhXT0wqe4Qjz4NcwzwXEm5R"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "I9BiEyz0vBnHzIQY+FHmLy8O8r3E1wqCEZM0Koi3eU418qI++mxgb8A3++4VCHl8XguB8pXCchMmn4HYfHfWMAMRTCo5Nz8PjPD5i2sz36ko9L"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "IdCJLv517XGOKPfEC6PIw4XCZ5DT8/xycEsZS/1nIu5VK21Ul/H1TjA/L2IoXbrSr0BJTlOq7l3lx9xOR+2suslC2XH8lnSQ+0I3Ku8ryTHmcv"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "ZYVWjJchUw4KLO964q+9XLWb5JPJCwo9MvWQiri4IhIzFfB1Xfcam62fiQJ74+jKSUweXOk+1wdbl7lAlO0gPXCtgm0c9gcIxkNhPW+C2840lb"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "4B9403yy4VFQz39UVctlBMBuHQW+C2+Hi5pmXpB/Fdf4v1Pk4G/CBMv5Rd01GRocFFA3N+r3tqjSLUr8Z950zUzyplafAlDtaHYWsJPSGoxK8P"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "1NDdpoLv8Nuryb050rJpEy30hLDxItzX7b3Fuxm/gjiKmcZjaENJxce151F2ZJhMWXClMUwa/mr3c41dINwBlI21zp7CrEF7pGMYGUxJoSeE+p"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "koDz2stkc4o0SI3FNOHpp4DD2KF8cvdfTEArjZ6kLPlPcw/kNWZruUxZF87eIuK7PrkJdfHf1iPdyHeObnb4f1eL1RBrfZ80ChR6YcpWyfrKJe"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "DnkYIPQifXa3TYHLNf1fWXd8e4E3QvERr+JugFuRHj0ItZPbhFtPBr9XwQ6HPdkGNy52qC3aFfdhK6nW+2g33Gco95Xubb3m4supEqYi6qFykQ"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "pFpXejgXiz9T4+Zg0+NojiK21F0WYSD1T6oZ5vXY8NNJ4oEz/P7fE019SXt08i35pCL+upSg6Du4e7N7wukzTTG1CxP+j+HTYZQg9ASLwb7h5w"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "3/dYDXkj96/i5c4GqD+WaQYPZ+nZ5tdcX6YzDL/Vuh4T0suLNLzVnf6SvukG6Q2yzkeYskIPoLy81VeNO7N7bo//OMzUS41VKOeHStzsZUaAm6"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "0v9IT+xkt9la4wQ8+bfd+Sx2LxYmf9IHMZVZqTR72alHdHmTeG+GfPHoUemWJIBaiSU4sFTVaRmk02uxzzFahFcvRZ5/mO8oBI708P+ydBsMlm"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "x7Ky1inm2g1u16/3vD/h8+OwrtuldAXCpBQm38kqp5Z7cBkqboiOX/ayIg5C76iezroVM42wviNo2wh1rJRrwyK8cyogc417gv7xVyClqj7WNv"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "yd4cs1VfygzNexzg1bV+gBM5xWPwLlHnGQci9p060Rkt/leks2mtNFIj0/t0GZJKEn+OGi9yLed2Rl05GnXU2uI/ch9xOfmrPJd8kPkyNy0wX3"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "h7L8u/GfxTzch/y5ID/9JE/kpJlzRr1UTmWhJ0iZDKrpTb2XOZfBn4lDequcOy0L6Wzwo4DbbUPoCeHQq5F212fzDMdTZpG+2UvBupKKY2/Gpi"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "HbdmQfRrwg283tXXVSniEv5LlQ8f0oB3/BPTsWDFk3FHpkyjGrMQeN5oX2QcoxKfDSoxd/zPrKJagOvjrz0/aQZeH/E9/nbojaCgTc3hBtd7aL"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "uqbhtzXyOwTeOfh8G6z3lbRFoEFAJTiUxb9bRZE9gH62QKV4DqDM95PzQJ0PssOQXqhwVnjRYNetaIpAxTWYCRQxqSxaTOa6hPoPcs/W+ZhBGr"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "wcleJd2X21hS/XNBXj6B7DbHFB8mBWGbf5Md/pGwuFnswDDVEZu/zbngOU1XdY1znISvP661ABXyLbqpjyKQ2wxFcqVVOxyqf8L9+LCGs87Kv0"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "HDmJAnnzofzGT4Re8p0O8V+Dv0hfnIn9ZnybJtc5URqa7j24aMwQj1/D3yNZ3OSepay2N0xNsw2UuDXX0rfivo/I3aMRaS9CbpP7FjPhIJ6hHr"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "aux4UZRTDxkri3XwP5WjXbMI1q2CD0EndZlv8lrslp1mknRugl33KnPaz2TXneDreuOxGhJwLA6R/3oJL1KNfvs67zqcb7QBCekx012cw7CaMg"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "7yRvm+VTtgAK9S/Ksiq8APhdEtROhd/2uMKQvuYItN6F3in5YZ0oQu8L1mk+snhL5ipGDdy33Ivcd69lVu4/uQn17Ee6zYM2VJPnIC/wMja0Nk"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "tjKR9F15H4SF0S34ty/0HkZZzFzVHOxJ3qPFNZOjDw/dqsHLb6gUk9K/NOuwk9s8+p7FbR5t9cV/J+4s8+J/8uVJOnooBBgEmBcj0IYvKQ4HeV"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "FPY2CPlCTx6a5FZYz8IF4u0tsAthN8H+KZ8QdlfCzlzneZ+E4NsT1lNl1RNqUWQPZ0dci9LDGh56NNwXiVi2IeQT1Z8A8fGnrIJzhNVueLAR9v"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Vy/JoNYeygcUPeosEeki1DHmmzR43AMRuXbgZz50sDeD4aO4TXcQ0IEVxbjulq2dIA7t9pxGHU6PRjvkOlWCT0JC1VchnyyHHNVPyjMdFvs66L"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "kb2zouQDiNOPkO7XIwxZALMKnw/BViOse/wovQHC8ueyatfzBszkd1zn2FzRnk2DGOrYBmPO4rm41nnIW0d+pEirxfg+7q0ynzk8A3E6BOlwUV"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Bt3J5dVwQShIhcX+JmPsVMQyf3dD0apIZXbXQcczYKiBFp9ODHkT8p4qjvkjy0rsdFWa6hpD5ouNMiStaWwniBdW4wG+PKb6bstrlXptycaJ12"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "IkKvKkc8utIeVl22Fs/bodZ1J0boocy58+5RxHclfn+Xdd2F5dP6VP2dUqaQnrfg+VhjGvKRvLMmLxJGgMgISXqbPGcoP+/zagOF02cECCMNv2"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "s74wrDcyM98T0LvSheLOnjDAtlthwln7FOi5l30g52Y+hfIt+R93pDbpnN6t+VyNNr4G6RN2fRmPYTnWH2xNOfRp5cgbS7f+QFpfU6Jn3lOo27"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Eadfe2H6KvGLuuM43JuUsc5yJvWT9IS3ged2P8Tz/qwctvqBocxI2SkUerXTZGeEn8CtI52b19280R3yb0w5qj8XBajLQfsifBqyL1rxBpcgUP"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "VXonB2niKBBw0F+X5vjN3PEHJlCDw50uyFsL1g+ftNbQblMHkhKtFfbHqLa4m706TiGF6FCrWnxRJIw5fDeluIIYbKqGvl0I3ZcT/S/P1+tf5J"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "6Y0dZSr5ON6QP9rb8VnFBLX0Vb6Sa9QP7biOubZ+f+swppxzjIr1YLw9f6LDvXwX6ncXNkIzj50hO91n1+zw/3Fc7/+8ytgaBmFHCHZ5M/er6X"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "uRPh8RAWiGUtv3fZxnThlBo58j2tGAIB6dq8v7j9vFzNWq6sM74i1pV40PG1md2jMD06RsyfXQ2JyKeJ8rjRY+fyPl2a/os6RXF2G/V+7Peiqk"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "T7b1QBpmadkWT8lPhLVzl62EunN+2aSzq2ya68SfQj0y6sWyXB16CeJ1aGe85Lvk42hk97NOO0F5Cmrxfu60F6sf0Td78FnWtQPEV0SKy788S9"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Hg/06btfCp1nHPSP0r+V6K6ktQ7n+IsGQ7n98gD38F+5GIV9zfZ8tV/ZJu02ZaQTl8Ge7p0x1xlXKG5yary9sW4OSAOLwCcXCEBavqT8qm3tZp"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "jwz4cn3c5+cR9in4bCuz9e+JyPdr+qDHhMWnYnRlrllU9M6SPAOhPjtLX32B+Yz02eblQaautKxeLtfSF+eXs/gTKEdGEI6ievw8lOWPIlzHM4"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "MyIydz5KyQNuBFWMq7lPsO//a63bbUISSXoNbYL3uryevaFst+kyO8rLd8wsEX4kG6s1MsicAZ6m14Y0sy08ybOhIP2T/Nm96oOOeYrBbEG7hU"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Ht0Wp2Qsl/MTv2je2l3huQxvnHi4f2IDINsQqIw/ElTlmXA9M1LOh9f6qnGQdb7lkaF4M2zcW48N2YaYN7CDybuOTX23c0bKbLdj0TYTSV/z0p"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "n1zBPyb4H0MNihnZbGqt3QeCkRevGzrbdc+mbqp+Et7Wpnb4d0lytZ0FG8DckWQXokVPwRvIFdlHXr9zikKmkhIi/SPzGbmvbCrGPnBiq+tree"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "QjEjFmQ1bMOGQCYDlAGU60GI8IOlt81+W4jsXZa9FOTkpQzno/zLZrfWCyGEELL18GUXfSNyuvXopdKj13ULETk2ygyBunqv8B2E1V+8yuLN2v"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "pks+iv71kK46MhOC8zc2tML54Iq7a4Ok3cyT3oX8lpIjbErviV+IOZ36I0bjEREUo/jLT6iA2CTDgD5ZJMupb8V+kjIvjx/zGyek4WTZg3f8uM"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "GelM5N9LgzAZglu7RYYj3yR/TXhxOpZhNkIIIWTSgGg5qDehN7RRji+y3gr4cF8Q6iXZ6jRHWNXFGwI1+I0tM6y0sWSGVs1Zken7/Uj/AJ//GJ"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "kE3HMPG8ymUUnF3x3TthJqSQSBcFXWa+oI12UikpVeMW1WMuY5P6Q3yjLXRibBmzIAM4If5Vwld6Gc3AjhdzU+r8RLy5+RJ9cjP+7I3BX0/MqL"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "TJTeIoex28sQQgghW5cgjA/oPnSbCT3p/bPeCjF7cMkwrTOshjSma9GQHmXOdJxIZLKrDMnONgdUvw4C68u+in+DBv1eXPPRbIhWehp77cETk8"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "bdCK97S1H8lV72y2sliOKvZdcbwzUhJiAwzrNBkAkmiAb/G2XaccoE8qhZRkZMVgXKZ7dnJBOMpRrKCCGEELKtkC3G6K1HD0LvCz3Nr4tkny/p"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Acnr/ZBGsSGrcC8qR+mhssmlNzt9ihFpRSuTBNlHacYxM82WJrX6ntnh2cmb5AB3iLHED/UFCHflJnElQk0aafm76B7bTRp908uzIROLY1/abv"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "dWuquwF6jDJJ4Jrpl83AZDJhApL36UXp316LrSfzwmZWWJzKns3DuPEEII2Zr0zU6fFYSyo7cRGDmWCT2IljN6Wg0WnboTxNF383v1xBCm6T1B"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "g6u0iL5rIW5+iN8W4//j/Gp8lB/FR/pR8hmE9flSlByN638jiPSwhA2358PP9fj/0U29LxBUI6JObCzCrmkSL+nFkQ2LG3+AYD3c27XL+b4uZi"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "18KuJ5zZgFRVUEtV4hq4FtSGQC8cP4M0Ek+Ssi3pH+YzVThkXIx9/OOWWCEEII2YrUhneFaPp9JpAcDdmIidBL/yb7ElmfxVTT15ghUxFgzvBa"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "rSn6RKyhARaBKCZnqMqmlvLZ/M78Lg11czhtvIKuaeJXemTk+ggf/yM9LoBIO8yrtO2f1iuV+MkIw542Mpa4idvhjaWK/jpC2fork7dHKgtqEH"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "sfC1Tjj2YXe1M+x1qGmmUGZVGld5Ui/aWenwtCCCFkiyIbNUayeKJbDwcaNpnnFib7W59d2OgHKj4GAnF9NlTrCnNLmzTQtpFuDumKGFN6va/0"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "5UG18Q0ZBh7Litp2ppmTRuLfZAJiLOIBJr2Ikb5BjviywZFJYlp/fU8/TP8XgvyXyP9HTTkw5UHKao7Jb80yE+n7IPKWIt8264g6QgghZNLxQ3"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "1IEElvlkN8jJiIliHpbco/ZqidWYOPLSn9nU2NpSvcybRmz4s0zjATByO+1qKhvhXi7hxfNT5pDt2PFu6Wd0B4r8iO6hANEIxj7ckTE/HZ2OhX"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "kg/Y4MiWoLKgVlaDz/erjY8i706B0L40COO/m95oFT8I4S/Hca3C3/+CgP8Tvj/Nr+qD+8L6M7zoePdZsYQQQsg2hRzDo5JbMjHkEiHWjIBJbv"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "HCod57nKrLwqBqzo/NDpgeswAaizVFnQzJyfCuuR+IuuT+IIr/7kfJT/H3l3G/L/dmzZ8zYRPn55rzXj8EUXDbuEWexFnFy3o7aYNMCpL2s+N+"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "rzJUM3vpifhX6e5e1NgNZWUufq9Q3BFCCJl6oPEKQv3N3L3vRizrdTJHf40FCKFylBwBMfmPERG2WYJP/LaKOjt3z3yvHwhU+ncIr98HSp+O6x"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "6Jz1eYBnwy9u4L688oRcnJuO66bJXxOERezZxr+0sjKgghhBBCJhq/khwYqMbqrkOsZi5TcgfE1Cut196ppXuXVLLMj4b+ZkSRCEvTA9YcUm03"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "EXP4zVxTBB3EXO1kfNpFE1F8LwTStb5KfwWh9U3Z5w4C76C+sPGMyd6QeVrVrKr9fNYTirgZEexIryKTM3NrS802M32hfqYNmhBCCCFkgpmx8H"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "EQST/KRJRDlLRa1ZzneZkX1Z9gfY+NWY3nliv6c36UnhlAqAVKP5gJOrvthenxy4aRIYIeCJQZdr0cf/8U7pdB1H0VIu9jnorf7IlAkgOxtxSV"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "wWfh3r+AuF2RCbwuwjjXpDdSFl/Uf+DxBAxCCCGETDYQVG8IVOPerJfNJU6allqRos/eaXP2DotO3amvkj4riOLX+6r+Lr9S/5BfiT/sq/iDfp"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "i8JwjTA7xK+hqzWKIKMSQbKm8NICQhRt+Ie1/sq/SarHdRRJ4rbbqZpJ34TVeWQv1VOSLNXoUQQgghZDJZHgShbmSLGLoNRdphVRUv9+YMzrUB"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "bD/MHJ4B0btfKUoW+FFySWBOuGj2NsqwsitNikzSU9J1CAI5/bmn4v08b8C3VyOEEEII2QL0x7ODMPmt2UzWKVhaTcSLnOSQnO/1D+1jQ5iCQH"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "CFemezB16UvC+I4m8Fkb7WiDsjaEXgZUPJYzfxLz2kjQ0I7zo/TA/xaidU7IUJIYQQQrYs2fmx+uas98olXlpNxB6EjNIr/CpEzNYaXh0TA2UT"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "T5kbV4vfXFLxN3ylf4f7WQl7NBNnMjzbS89mnkkYkn74VMnfSmH8ZW/O4u2v55MQQgghU4+gUn9NoNJbexN7MpSZiSLZJiQIk7d4oa7aoLY+80"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "7aweyFBgHbpxoH4b6OhbD7VRDFq7K4iyiTxSCysEKE3XiGZpsm4WSrgoNIX1mK4q9NH++iFUIIIYSQyQJi73WBGrpp0x51LmHTanBjF3L4qvFT"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "v5p8VPaZs8FtKUoy/Axh91zZMsaP5BD7ZImdZ7c6O9C+tbdO7mtzhJ01s6cfBJ5K1orYLYfJEZtzlBohhBBCyKRTrtRf7Kv09yPDkC6R02Ei+E"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "RQLYbgS6+SEx/KYXpIXyV+trfH4HQb9OYj26rMXvgUOX/XV/rgoJIcE8gxVhBafqRvNPGoyV59MsfOzJNri+fmmIhD6QlEuLhXCMnbIfJO9tXg"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "QVNj+JoQQgghRKgselIQ6dMyoSSCqdceMBF8SzKxFSUbIMKugyD6uRwGD2F2OITRm8rV5DnTZVh1l7jfCDc5hkpMjpqasfBxchzVtFmNvcrVwX"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "2DanoA/H2ypJIYn2fK8CsE3ZWI220Ia73ZgHnShJ2YFXdmeBbhK70ecbjIV/Fhnhp6gRkiJoQQQgiZcswafCxEzYcgeP6aCR0Z+uxV8IlJ75cM"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "mcpxZWbe31qIs7th/4Bd76vkLxBtf4TJxsiwWP6+Ete8BkLuRri/Fe7ug7hak/WiiaCzYU2KqBNr9trJsGxzqBfXV/rqoJocU64M7uvNasyxKU"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "QIIYQQMsV5/OATS1GyEGLnthHxNmaR1SqgRLSJABTBlmfyu7gTsSULJsYiMMdqCNvES64rAlKvg8C8y1fxH/D7CeVqsm82NDtQtilCCCGEELKd"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "oRY9HWIvgfi5HrYh610TYTSZImwyrEXY1eQeICSznsNrIO7O8ZX+nFeN9/G8CZxbSAghhBAyFZge6j38KD7Sl3Noo+S2rPdNhlUnazh1c0xEqJ"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "1fJ/P4aiebvxH3e/0ouQKfZ5WUPlbmAu44azs87YMQQgghZFzsMTg9UPF+fkV/zlfpd32V3GSEVcfiiC3R4yfXaB0almFYxEO2VlHJQ4HSf4Ww"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "Oxd/x+Uw+ZSnGm/2KvGT7Z0QQgghhJBcZMVspJ/nV+oH+kp/AcJquR/FV/qRvn/T/LeWeXfGZO5d00QQ5lmLuxG/Eo5Yc67f8EZc9yHYzbjuhd"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "mK4eRoiND3QwS+vq+S7u1Vl3ELFEIIIYSQzWbGMTNlixYRf0E4tH9fFL8PwusrEIAnQozJPLiLYVdDkK0IVHInfrsHvz3QZvfj93uCUN8Bt/+A"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "XeNHyaW+Ss6F+28HKo7LUXIkvj9I9tXzwuSFnor/06st2lVWDNuYEEIIIYSQyWfAN/vlyV55Uaq8WfPnmJMkwsEnGlFYq+85ymRoVb5X6e6ZeG"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "vMMf7mzp/l9R+3y4RuxEwIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQggh"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "hBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEE"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "IIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQggh"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "hBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEE"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "IIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQggh"
    AngloAmericanLogoBase64 = AngloAmericanLogoBase64 & "hBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQrY/PO//A7jLn+oF5KoXAAAAAElFTkSuQmCC"

    getAngloLogoBase64 = AngloAmericanLogoBase64
   
End Function
