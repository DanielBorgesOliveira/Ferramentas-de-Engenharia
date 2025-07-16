Attribute VB_Name = "Tools"
Sub CopyHeadings()
    Dim paraTexto As String
    Dim p As Paragraph
    Dim tituloComNumero As String
    Dim nivel As Integer
    Dim recuo As String

    For Each p In ActiveDocument.Paragraphs
        Select Case p.Style
            Case "Heading 1": nivel = 1
            Case "Heading 2": nivel = 2
            Case "Heading 3": nivel = 3
            Case "Heading 4": nivel = 4
            Case Else: nivel = 0
        End Select

        If nivel > 0 Then
            ' Define a indentação com tabs
            recuo = String(nivel - 1, vbTab)
            
            ' Monta o texto com numeração e indentação
            If p.Range.ListFormat.ListString <> "" Then
                tituloComNumero = recuo & p.Range.ListFormat.ListString & " " & Trim(Replace(p.Range.text, vbCr, ""))
            Else
                tituloComNumero = recuo & Trim(Replace(p.Range.text, vbCr, ""))
            End If
            
            ' Adiciona ao texto final com quebra apenas ao fim de cada item
            paraTexto = paraTexto & tituloComNumero & vbCrLf
        End If
    Next p

    ' Cria novo documento com os Headings organizados
    Documents.Add
    Selection.TypeText paraTexto
    MsgBox "Headings com tabulação e sem linhas extras copiados para novo documento."
End Sub


