Attribute VB_Name = "FinishHim"
Option Explicit

Private Type ClientStatus
    CorrectionNeeded As Boolean
    Words As String
End Type

Sub FinishHim()

    If Library.propertyExists("NumeroCliente") Then
        UserForm1.NumeroCliente = ActiveDocument.CustomDocumentProperties("NumeroCliente").value
    End If
    If Library.propertyExists("NumeroNosso") Then
        UserForm1.NumeroNosso = ActiveDocument.CustomDocumentProperties("NumeroNosso").value
    End If
    If Library.propertyExists("Revisao") Then
        UserForm1.Revisao = ActiveDocument.CustomDocumentProperties("Revisao").value
    End If
    If Library.propertyExists("Titulo1") Then
        UserForm1.Titulo1 = ActiveDocument.CustomDocumentProperties("Titulo1").value
    End If
    If Library.propertyExists("Titulo2") Then
        UserForm1.Titulo2 = ActiveDocument.CustomDocumentProperties("Titulo2").value
    End If
    If Library.propertyExists("Titulo3") Then
        UserForm1.Titulo3 = ActiveDocument.CustomDocumentProperties("Titulo3").value
    End If
    If Library.propertyExists("Titulo4") Then
        UserForm1.Titulo4 = ActiveDocument.CustomDocumentProperties("Titulo4").value
    End If
    If Library.propertyExists("Titulo5") Then
        UserForm1.Titulo5 = ActiveDocument.CustomDocumentProperties("Titulo5").value
    End If
    If Library.propertyExists("Cliente") Then
        UserForm1.Cliente = ActiveDocument.CustomDocumentProperties("Cliente").value
    End If
    If Library.propertyExists("Projeto") Then
        UserForm1.Projeto = ActiveDocument.CustomDocumentProperties("Projeto").value
    End If
    
    UserForm1.Show
    
    Dim FileName As String
    FileName = ActiveDocument.name
    ActiveDocument.BuiltInDocumentProperties("Title").value = FileName
    ActiveDocument.BuiltInDocumentProperties("Author").value = Application.UserName
    ActiveDocument.BuiltInDocumentProperties("Company").value = "Brass do Brasil"
    Call UpdateProperty("NumeroCliente", UserForm1.NumeroCliente)
    Call UpdateProperty("NumeroNosso", UserForm1.NumeroNosso)
    Call UpdateProperty("Revisao", UserForm1.Revisao)
    Call UpdateProperty("Titulo1", UserForm1.Titulo1)
    Call UpdateProperty("Titulo2", UserForm1.Titulo2)
    Call UpdateProperty("Titulo3", UserForm1.Titulo3)
    Call UpdateProperty("Titulo4", UserForm1.Titulo4)
    Call UpdateProperty("Titulo5", UserForm1.Titulo5)
    Call UpdateProperty("Cliente", UserForm1.Cliente)
    Call UpdateProperty("Projeto", UserForm1.Projeto)
    
    ' Update the table of contents
    Dim TOC As TableOfContents
    For Each TOC In ActiveDocument.TablesOfContents
        TOC.Update
    Next
    
    If UserForm1.ResolverComentarios Then
        Call RemoveAllComments
    End If
    
    ' ************************************ Adiciona Cabeçalho e Rodapé ************************************ '
    If UserForm1.convertVale Then
        If UserForm1.ResolverComentarios Then
            If UserForm1.Cliente = "Vale" Then
                Call Vale.finish
            ElseIf UserForm1.Cliente = "Anglo American" Then
                Call AngloAmerican.finish
            End If
        Else
            MsgBox "O documento não será convertido para o padrão do cliente, pois os comentários não foram resolvido."
        End If
    End If
    ' ************************************ ******** ********* * ****** ************************************ '
    
    Dim Warning As String
    
    ' Verifica se temos erros de referência cruzada
    If ReferenceError() Then
        Warning = Warning & "Foram encontrados erros de referência cruzada. Verificar as referências cruzadas no documento." & vbNewLine & vbNewLine
    End If
    
    ' Verifica se temos anexos no documento.
    If hasAttachement Then
        Warning = Warning & "Foram encontrados anexos no documento. Lembre-se de anexar os anexos no PDF resutlante." & vbNewLine & vbNewLine
    End If
    
    ' Verifica se temos citações à outros clientes no documento.
    Dim ClientVerification As ClientStatus
    ClientVerification = verifyClient(UserForm1.Cliente)
    If ClientVerification.CorrectionNeeded Then
        Warning = Warning & "Foram encontrados referências às outros clientes no documento. Palavras encontradas: " & ClientVerification.Words & "." & vbNewLine & vbNewLine
    End If
    
    If Not isStringEmpty(Warning) Then MsgBox Warning, vbCritical
    
    If UserForm1.ExportarPDF Then
        Call ExportPDF(ActiveDocument.FullName)
    End If
    
End Sub

Private Function isStringEmpty(text As String) As Boolean
    If Trim(text & vbNullString) = vbNullString Then
        isStringEmpty = True
    Else
        isStringEmpty = False
    End If
End Function

Private Sub UpdateProperty(name As String, value As String)
    If Library.propertyExists(name) Then
        ActiveDocument.CustomDocumentProperties(name).value = value
    Else
        ActiveDocument.CustomDocumentProperties.Add name:=name, LinkToContent:=False, Type:=msoPropertyTypeString, value:=value
    End If
End Sub

Private Sub ExportPDF(OutputFile As String)
    ' Export PDF
    OutputFile = Replace(OutputFile, ".docx", ".pdf")
    OutputFile = Replace(OutputFile, ".doc", ".pdf")
    ActiveDocument.ExportAsFixedFormat OutputFileName:=OutputFile, _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
    
    ' Save the documment
    If ActiveDocument.Saved = False Then ActiveDocument.Save
    
    ' Attach PDF: there is problems with anti virus when calling python.
    'Do While True
    '    Dim Answer As Integer
    '    Answer = MsgBox("Gostaria de adicionar algum anexo ao PDF?", vbQuestion + vbYesNo + vbDefaultButton2, "Arquivo para Anexar no PDF")
    '    If Answer = vbYes Then
    '        MsgBox "Indique o arquivo com os resultados para anexar ao PDF.", vbOKOnly + vbQuestion, "Arquivo para Anexar no PDF"
    '        Dim FileName As String
    '        FileName = UseFileDialog
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

Private Sub RemoveAllComments()
    ' Delete all comments
    Dim Comments As Comments
    Set Comments = ActiveDocument.Comments
    
    Dim n As Long
    For n = Comments.Count To 1 Step -1
        Comments(n).Delete
    Next
    
    Set Comments = Nothing
    
    'Accept all revisions and disable track revisions
    ActiveDocument.Revisions.AcceptAll
    ActiveDocument.TrackRevisions = False
End Sub

Function UseFileDialog() As String
    Dim lngCount As Long
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Show
        UseFileDialog = .SelectedItems(1)
    End With
End Function

Private Function ReferenceError() As Boolean
    
    If ActiveDocument.Content.Find.Execute( _
        FindText:="Error", _
        MatchCase:=False, _
        MatchWholeWord:=True, _
        Forward:=True) Then
        
        ReferenceError = True
    ElseIf ActiveDocument.Content.Find.Execute( _
        FindText:="Erro", _
        MatchCase:=False, _
        MatchWholeWord:=True, _
        Forward:=True) Then
        
        ReferenceError = True
    Else
        ReferenceError = False
    End If
    
End Function

Private Function hasAttachement() As Boolean
    
    If ActiveDocument.Content.Find.Execute( _
        FindText:="Anexo", _
        MatchCase:=False, _
        MatchWholeWord:=True, _
        Forward:=True) Then
        
        hasAttachement = True
    Else
        hasAttachement = False
    End If
    
End Function

Private Function verifyClient(Client As String) As ClientStatus
    
    Call Variaveis.Main
    
    Dim sets As ClientStatus
    With sets
        .CorrectionNeeded = False
        .Words = ""
    End With
    
    Dim SelectArray As Variant
    If Client = "Vale" Then SelectArray = Array(1, 2)
    If Client = "Anglo American" Then SelectArray = Array(0, 2)
    If Client = "CBMM" Then SelectArray = Array(0, 1)
    
    Dim Counter As Integer
    Counter = 0
    
    Dim Item As Variant
    Dim Word As Variant
    For Each Item In SelectArray
        For Each Word In Variaveis.Words(Item)
            If ActiveDocument.Content.Find.Execute( _
                FindText:=Word, _
                MatchCase:=False, _
                MatchWholeWord:=True, _
                Forward:=True) Then
                
                sets.Words = sets.Words & ", " & Word
                Counter = Counter + 1
            End If
        Next
    Next
    
    If Counter Then
        sets.CorrectionNeeded = True
    Else
        sets.CorrectionNeeded = False
    End If
    
    verifyClient = sets
    
End Function


































