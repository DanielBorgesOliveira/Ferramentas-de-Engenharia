Attribute VB_Name = "FinishHim"
'/*******************************************************************
'* Copyright         : 2024 Daniel Oliveira
'* Email             : danielbo17@hotmail.com
'* File Name         : FinishHim
'* Description       : Finaliza documentos WORD em formato neutro e converte para o padrão do cliente..
'*
'* Revision History  :
'* 16/12/2024      Daniel Borges de Oliveira          Desenvolvimento Inicial
'/******************************************************************/

Option Explicit

Private Type ClientStatus
    CorrectionNeeded As Boolean
    Words As String
End Type

Sub FinishHim()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim Cliente As String
    Cliente = UserForm1.Cliente

    If libUtils.propertyExists("NumeroCliente") Then
        UserForm1.NumeroCliente = doc.CustomDocumentProperties("NumeroCliente").value
    End If
    If libUtils.propertyExists("NumeroNosso") Then
        UserForm1.NumeroNosso = doc.CustomDocumentProperties("NumeroNosso").value
    End If
    If libUtils.propertyExists("Revisao") Then
        UserForm1.Revisao = doc.CustomDocumentProperties("Revisao").value
    End If
    If libUtils.propertyExists("Titulo1") Then
        UserForm1.Titulo1 = doc.CustomDocumentProperties("Titulo1").value
    End If
    If libUtils.propertyExists("Titulo2") Then
        UserForm1.Titulo2 = doc.CustomDocumentProperties("Titulo2").value
    End If
    If libUtils.propertyExists("Titulo3") Then
        UserForm1.Titulo3 = doc.CustomDocumentProperties("Titulo3").value
    End If
    If libUtils.propertyExists("Titulo4") Then
        UserForm1.Titulo4 = doc.CustomDocumentProperties("Titulo4").value
    End If
    If libUtils.propertyExists("Titulo5") Then
        UserForm1.Titulo5 = doc.CustomDocumentProperties("Titulo5").value
    End If
    If libUtils.propertyExists("Cliente") Then
        UserForm1.Cliente = doc.CustomDocumentProperties("Cliente").value
    End If
    If libUtils.propertyExists("Projeto") Then
        UserForm1.Projeto = doc.CustomDocumentProperties("Projeto").value
    End If
    
    UserForm1.Show
    
    Dim fileName As String
    fileName = doc.name
    doc.BuiltInDocumentProperties("Title").value = fileName
    doc.BuiltInDocumentProperties("Author").value = Application.UserName
    doc.BuiltInDocumentProperties("Company").value = "Brass do Brasil"
    Call libUtils.UpdateProperty(doc, "NumeroCliente", UserForm1.NumeroCliente)
    Call libUtils.UpdateProperty(doc, "NumeroNosso", UserForm1.NumeroNosso)
    Call libUtils.UpdateProperty(doc, "Revisao", UserForm1.Revisao)
    Call libUtils.UpdateProperty(doc, "Titulo1", UserForm1.Titulo1)
    Call libUtils.UpdateProperty(doc, "Titulo2", UserForm1.Titulo2)
    Call libUtils.UpdateProperty(doc, "Titulo3", UserForm1.Titulo3)
    Call libUtils.UpdateProperty(doc, "Titulo4", UserForm1.Titulo4)
    Call libUtils.UpdateProperty(doc, "Titulo5", UserForm1.Titulo5)
    Call libUtils.UpdateProperty(doc, "Cliente", UserForm1.Cliente)
    Call libUtils.UpdateProperty(doc, "Projeto", UserForm1.Projeto)
    
    Call libUtils.UpdateAllFields
    
    ' Update the table of contents
    Dim TOC As TableOfContents
    For Each TOC In doc.TablesOfContents
        TOC.Update
    Next
    
    Call libUtils.AcceptAllChangesAndStopTracking(doc)
    Call libUtils.RemoveAllComments(doc)
    Call libUtils.NegritoCamposTabelaFigura(doc)
    
    ' Verifica se temos erros de referência cruzada
    Dim Warning As String
    If libUtils.findErrorErrorKeyword(doc) Then
        Warning = Warning & "Foram encontrados erros de referência cruzada. Verificar as referências cruzadas no documento." & vbNewLine & vbNewLine
    End If
    
    ' Verifica se temos anexos no documento.
    If libUtils.HasAttachment(doc) Then
        Warning = Warning & "Foram encontrados anexos no documento. Lembre-se de anexar os anexos no PDF resultante." & vbNewLine & vbNewLine
    End If
    
    ' Verifica se temos citações à outros clientes no documento.
    Dim ClientVerification As ClientStatus
    ClientVerification = VerifyClientContentForCorrections(doc, Cliente)
    If ClientVerification.CorrectionNeeded Then
        Warning = Warning & "Foram encontrados referências às outros clientes no documento. Palavras encontradas: " & ClientVerification.Words & "." & vbNewLine & vbNewLine
    End If
    
    If Not libUtils.isStringEmpty(Warning) Then MsgBox Warning, vbCritical
    
    'If UserForm1.ExportarPDF Then
        Call ExportPDF(doc)
    'End If
    
End Sub

Private Sub ExportPDF(doc As Document)
    Dim OutputFile As String
    OutputFile = Replace(doc.FullName, ".docx", ".pdf")
    OutputFile = Replace(OutputFile, ".doc", ".pdf")
    
    Call doc.ExportAsFixedFormat( _
        OutputFileName:=OutputFile, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=True, _
        OptimizeFor:= _
        wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument, _
        From:=1, _
        To:=1, _
        item:=wdExportDocumentContent, _
        IncludeDocProps:=True, _
        KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, _
        DocStructureTags:=True, _
        BitmapMissingFonts:=True, _
        UseISO19005_1:=False _
    )
    
    ' Save the documment
    'If doc.Saved = False Then doc.Save
    
    ' Attach PDF: there is problems with anti virus when calling python.
    'Do While True
    '    Dim Answer As Integer
    '    Answer = MsgBox("Gostaria de adicionar algum anexo ao PDF?", vbQuestion + vbYesNo + vbDefaultButton2, "Arquivo para Anexar no PDF")
    '    If Answer = vbYes Then
    '        MsgBox "Indique o arquivo com os resultados para anexar ao PDF.", vbOKOnly + vbQuestion, "Arquivo para Anexar no PDF"
    '        Dim FileName As String
    '        FileName = libUtils.UseFileDialog
    '        Call AttachToPDF(OutputFile, FileName)
    '    Else
    '        Exit Do
    '    End If
    'Loop
End Sub

Private Sub AttachToPDF(InputFile As String, AttachedFile As String)
    
    Dim PythonExe As String
    PythonExe = """C:\Users\dboliveira\AppData\Local\Programs\Python\Python312\python.exe"""
    
    Dim PythonScript As String
    PythonScript = """C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\0-GERENCIAMENTO\Script\AttachPDF.py"""
    
    Dim Arguments As String
    Arguments = Chr(34) & InputFile & Chr(34) & " " & Chr(34) & AttachedFile & Chr(34)
    
    'Dim objShell As Object
    'Set objShell = VBA.CreateObject("Wscript.Shell")
    'objShell.Run PythonExe & PythonScript
    'RetVal = Shell("C:\Users\dboliveira\AppData\Local\Programs\Python\Python312\python.exe " & "C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\0-GERENCIAMENTO\Script\AttachPDF.py " & "InputFile " & "AttachedFile ")
    
    MsgBox (PythonExe & " " & PythonScript & " " & Arguments)
    
End Sub

Private Function VerifyClientContentForCorrections(doc As Document, client As String) As ClientStatus
    '---------------------------------------------------------------------------------------
    ' Function: VerifyClientContentForCorrections
    '
    ' Description:
    '   Checks the document content for words that are considered incorrect or undesired
    '   for the specified client. The function relies on a predefined list of words stored
    '   in the Variaveis module. Each client is associated with a specific combination of
    '   word lists, and the document is scanned accordingly.
    '
    ' Parameters:
    '   doc    - The Word Document to be analyzed.
    '   client - The client name ("Vale", "Anglo American", or "CBMM").
    '
    ' Returns:
    '   A ClientStatus type with:
    '       .CorrectionNeeded = True if any undesired words are found.
    '       .Words = A comma-separated list of the found words (if any).
    '---------------------------------------------------------------------------------------

    ' Initialize word sets and correction flag
    Call Variaveis.Main
    
    Dim sets As ClientStatus
    sets.CorrectionNeeded = False
    sets.Words = ""

    ' Determine word list indexes based on client
    Dim selectArray As Variant
    Select Case client
        Case "Vale"
            selectArray = Array(1, 2)
        Case "Anglo American"
            selectArray = Array(0, 2)
        Case "CBMM"
            selectArray = Array(0, 1)
        Case Else
            VerifyClientContentForCorrections = sets
            Exit Function
    End Select

    ' Scan document for each word in the selected lists
    Dim item As Variant, word As Variant
    Dim counter As Integer: counter = 0

    For Each item In selectArray
        For Each word In Variaveis.Words(item)
            If doc.Content.Find.Execute( _
                FindText:=word, _
                MatchCase:=False, _
                MatchWholeWord:=True, _
                Forward:=True) Then

                sets.Words = sets.Words & ", " & word
                counter = counter + 1
            End If
        Next word
    Next item

    ' Set correction flag if any words were found
    sets.CorrectionNeeded = (counter > 0)

    ' Return result
    VerifyClientContentForCorrections = sets
End Function



































