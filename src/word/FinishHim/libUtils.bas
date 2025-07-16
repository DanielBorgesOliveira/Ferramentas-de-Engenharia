Attribute VB_Name = "libUtils"
'Option Explicit
Option Private Module

#If VBA7 And Win64 Then
    ' For 64bit version of Excel
    Public Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As LongPtr)
#Else
    ' For 32bit version of Excel
    Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
#End If

Sub CreateZipFile(folderToZipPath As Variant, zippedFileFullName As Variant)

    Dim ShellApp As Object
    
    'Create an empty zip file
    Open zippedFileFullName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    
    'Copy the files & folders into the zip file
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(zippedFileFullName).CopyHere ShellApp.Namespace(folderToZipPath).items
    
    'Zipping the files may take a while, create loop to pause the macro until zipping has finished.
    On Error Resume Next
    Do Until ShellApp.Namespace(zippedFileFullName).items.count = ShellApp.Namespace(folderToZipPath).items.count
        Sleep 1000
    Loop
    On Error GoTo 0

End Sub

Function UseFolderDialog() As String
    Dim lngCount As Long
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show
        UseFolderDialog = .SelectedItems(1)
    End With
End Function

Function UseFileDialog( _
    Optional DialogType As Integer = msoFileDialogSaveAs, _
    Optional AllowMultiSelect As Boolean = False) As Variant
    Dim lngCount As Long
    ' Open the file dialog
    With Application.FileDialog(DialogType)
        .AllowMultiSelect = AllowMultiSelect
        .Show
        
        If Not AllowMultiSelect Then
            UseFileDialog = CStr(.SelectedItems(1))
        Else
            Dim items() As Variant
            Dim count As Integer
            
            For Each SelectedFile In .SelectedItems
                    count = count + 1
                    ReDim Preserve items(1 To count)
                    items(count) = SelectedFile
            Next SelectedFile
            
            UseFileDialog = items
        End If
    End With
End Function

Function propertyExists(propName) As Boolean
    Dim tempObj
    On Error Resume Next
    Set tempObj = ActiveDocument.CustomDocumentProperties.item(propName)
    propertyExists = (Err = 0)
    On Error GoTo 0
End Function

Public Function EncodeFile(strPicPath As String) As String
    Const adTypeBinary = 1          ' Binary file is encoded

    ' Variables for encoding
    Dim objXML
    Dim objDocElem

    ' Variable for reading binary picture
    Dim objStream

    ' Open data stream from picture
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = adTypeBinary
    objStream.Open
    objStream.LoadFromFile (strPicPath)

    ' Create XML Document object and root node
    ' that will contain the data
    Set objXML = CreateObject("MSXml2.DOMDocument")
    Set objDocElem = objXML.createElement("Base64Data")
    objDocElem.dataType = "bin.base64"

    ' Set binary value
    objDocElem.nodeTypedValue = objStream.Read()

    ' Get base64 value
    EncodeFile = objDocElem.text

    ' Clean all
    Set objXML = Nothing
    Set objDocElem = Nothing
    Set objStream = Nothing

End Function

Function InsertImage(ImageNameBase64 As String, rng As Range)
    
    Dim ImagePath As String
    ImagePath = Environ("TMP") & "\temp.png"
    
    
    
    ' Save byte array to temp file
    Open ImagePath For Binary As #1
       Put #1, 1, DecodeBase64(ImageNameBase64)
    Close #1
    
    ' Insert image from temp file
    rng.InlineShapes.AddPicture fileName:=ImagePath, Range:=rng
    
    Kill ImagePath
    
End Function
    
Private Function DecodeBase64(ByVal strData As String) As Byte()

    Dim objXML As Object 'MSXML2.DOMDocument
    Dim objNode As Object 'MSXML2.IXMLDOMElement
    
    'get dom document
    Set objXML = CreateObject("MSXML2.DOMDocument")
    
    'create node with type of base 64 and decode
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.text = strData

    DecodeBase64 = objNode.nodeTypedValue
    
    'clean up
    Set objNode = Nothing
    Set objXML = Nothing
    
End Function

Sub ResizeImages()
    Dim i As Long
    With ActiveDocument
        For i = 1 To .InlineShapes.count - 1
            With .InlineShapes(i)
                Aspect = .width / .Height
                .width = CentimetersToPoints(16)
                .Height = .width / Aspect
                Set rng = .Range
                rng.Style = ActiveDocument.styles("VALE_IMAGEM")
            End With
        Next i
    End With
End Sub

Public Sub UpdateAllFields()
  
  Dim lngJunk As Long
  lngJunk = ActiveDocument.Sections(1).Headers(1).Range.StoryType
  
  Dim rngStory As word.Range
  Dim oShp As Shape
  For Each rngStory In ActiveDocument.StoryRanges
    'Iterate through all linked stories
    Do
      On Error Resume Next
      rngStory.Fields.Update
      Select Case rngStory.StoryType
        Case 6, 7, 8, 9, 10, 11
          If rngStory.ShapeRange.count > 0 Then
            For Each oShp In rngStory.ShapeRange
              If oShp.TextFrame.HasText Then
                oShp.TextFrame.TextRange.Fields.Update
              End If
            Next
          End If
        Case Else
          'Do Nothing
        End Select
        On Error GoTo 0
        'Get next linked story (if any)
        Set rngStory = rngStory.NextStoryRange
      Loop Until rngStory Is Nothing
    Next
End Sub

Function GUID$(Optional lowercase As Boolean, Optional parens As Boolean)
    Dim k&, h$
    GUID = Space(36)
    For k = 1 To Len(GUID)
        Randomize
        Select Case k
            Case 9, 14, 19, 24: h = "-"
            Case 15:            h = "4"
            Case 20:            h = Hex(Rnd * 3 + 8)
            Case Else:          h = Hex(Rnd * 15)
        End Select
        Mid$(GUID, k, 1) = h
    Next
    If lowercase Then GUID = LCase$(GUID)
    If parens Then GUID = "{" & GUID & "}"
End Function

Function TrustVBAAccess() As Boolean
    On Error Resume Next
    Dim testProj As Object
    Set testProj = ActiveDocument.VBProject
    TrustVBAAccess = Not testProj Is Nothing
    On Error GoTo 0
End Function

Sub BackupModules(fileName As String)
    ' Verifica se o acesso ao projeto VBA está permitido
    If Not TrustVBAAccess() Then
        MsgBox "O acesso ao projeto VBA está restrito. Habilite o acesso programático.", vbCritical
        Exit Sub
    End If

    ' Define a pasta temporária para exportação
    Dim FolderInputPath As String
    FolderInputPath = Environ("temp") & "\" & GUID(lowercase:=True)

    If Dir(FolderInputPath, vbDirectory) = "" Then MkDir FolderInputPath
    If Right(FolderInputPath, 1) <> "\" Then FolderInputPath = FolderInputPath & "\"

    ' Exporta os módulos do projeto VBA
    Dim vbComp As Object
    Dim strModuleName As String
    For Each vbComp In ActiveDocument.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: strModuleName = vbComp.name & ".bas"
            Case 2: strModuleName = vbComp.name & ".cls"
            Case 3: strModuleName = vbComp.name & ".frm"
            Case Else: strModuleName = vbComp.name & ".txt"
        End Select
        vbComp.Export FolderInputPath & "\" & strModuleName
    Next vbComp

    ' Permite seleção de arquivos relacionados ao projeto para backup
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim sourceFilePaths As Variant
    MsgBox "Indique os arquivos que compõem o projeto."
    sourceFilePaths = UseFileDialog(DialogType:=msoFileDialogFilePicker, AllowMultiSelect:=True)
    
    Dim sourceFilePath As Variant
    For Each sourceFilePath In sourceFilePaths
        Dim destFilePath As String
        destFilePath = FolderInputPath & fso.GetFileName(sourceFilePath)
        fso.CopyFile sourceFilePath, destFilePath
    Next sourceFilePath

    Set fso = Nothing

    ' Define local para salvar o arquivo zip
    Dim FolderOutputPath As String
    MsgBox "Indique a pasta onde o backup será salvo."
    FolderOutputPath = UseFolderDialog() & "\" & fileName & "-" & Format(Now(), "DD-MMM-YYYY-hh-mm-ss") & ".zip"

    ' Cria o arquivo zip
    Call libUtils.CreateZipFile(CVar(FolderInputPath), CVar(FolderOutputPath))

    MsgBox "Backup concluído com sucesso!", vbInformation
End Sub

Sub UpdateProperty(doc As Document, name As String, value As String)
    If libUtils.propertyExists(name) Then
        doc.CustomDocumentProperties(name).value = value
    Else
        doc.CustomDocumentProperties.Add name:=name, LinkToContent:=False, Type:=msoPropertyTypeString, value:=value
    End If
End Sub

Sub RemoveAllComments(doc As Document)
    ' Delete all comments
    Dim Comments As Comments
    Set Comments = doc.Comments
    
    Dim n As Long
    For n = Comments.count To 1 Step -1
        Comments(n).Delete
    Next
    
    Set Comments = Nothing
End Sub

Function isStringEmpty(text As String) As Boolean
    If Trim(text & vbNullString) = vbNullString Then
        isStringEmpty = True
    Else
        isStringEmpty = False
    End If
End Function

Sub AcceptAllChangesAndStopTracking(doc As Document)
    ' Accept all changes in the document
    doc.AcceptAllRevisions
    
    ' Loop through each section to accept changes in headers
    Dim sec As Section
    Dim rev As Revision
    For Each sec In doc.Sections
        ' Check for revisions in the primary header
        With sec.Headers(wdHeaderFooterPrimary)
            For Each rev In .Range.Revisions
                rev.Accept
            Next rev
        End With
    Next sec
    
    ' Stop tracking changes
    doc.TrackRevisions = False
End Sub

Function findErrorErrorKeyword(doc As Document) As Boolean
    '---------------------------------------------------------------------------------------
    ' Function: DocumentContainsErrorKeyword
    '
    ' Description:
    '   Checks if the specified Word document contains the keyword "Error" or "Erro"
    '   (case-insensitive, whole word match). This can be useful for identifying the
    '   presence of reference or processing errors in the document.
    '
    ' Parameters:
    '   doc - The Word Document object to search through.
    '
    ' Returns:
    '   True if either "Error" or "Erro" is found as a whole word (case-insensitive).
    '   False otherwise.
    '---------------------------------------------------------------------------------------

    Dim keywords As Variant
    Dim keyword As Variant
    Dim found As Boolean

    ' List of keywords to search for
    keywords = Array("Error", "Erro")
    found = False

    ' Loop through all keywords and search in the document content
    For Each keyword In keywords
        With doc.Content.Find
            .text = keyword
            .MatchCase = False
            .MatchWholeWord = True
            .Forward = True

            If .Execute Then
                found = True
                Exit For ' Exit loop as soon as one keyword is found
            End If
        End With
    Next keyword

    DocumentContainsErrorKeyword = found

End Function

Function HasAttachment(doc As Document) As Boolean
    '---------------------------------------------------------------------------------------
    ' Function: HasAttachment
    '
    ' Description:
    '   Checks if the active Word document contains the keyword "Anexo"
    '   (case-insensitive, whole word match). Useful to verify if attachments are
    '   referenced in the document content.
    '
    ' Parameters:
    '   None (uses ActiveDocument).
    '
    ' Returns:
    '   True if "Anexo" is found as a whole word (case-insensitive).
    '   False otherwise.
    '---------------------------------------------------------------------------------------

    Dim found As Boolean
    found = False

    With doc.Content.Find
        .text = "Anexo"
        .MatchCase = False
        .MatchWholeWord = True
        .Forward = True

        If .Execute Then
            found = True
        End If
    End With

    HasAttachment = found
End Function

Sub CopyCustomDocProperties(docSource As Document, docDestination As Document)
    Dim prop As DocumentProperty
    
    If docSource.CustomDocumentProperties.count = 0 Then
        Exit Sub
    End If
    
    For Each prop In docSource.CustomDocumentProperties
        On Error Resume Next
        Call UpdateProperty(docDestination, prop.name, prop.value)
        On Error GoTo 0
    Next prop
End Sub

Sub CopyBuiltinDocProperties(docSource As Document, docDestination As Document)
    Dim prop As DocumentProperty
    Dim propName As String

    For Each prop In docSource.BuiltInDocumentProperties
        propName = prop.name
        On Error Resume Next
        If Not prop.ReadOnly Then
            docDestination.BuiltInDocumentProperties(propName).value = prop.value
        End If
        On Error GoTo 0
    Next prop
End Sub


Sub AplicarEstiloEmTodasTabelas()
    Dim tbl As Table
    Dim cel As Cell
    Dim rng As Range

    For Each tbl In ActiveDocument.Tables
        For Each cel In tbl.Range.Cells
            Set rng = cel.Range
            rng.End = rng.End - 1 ' Remove o caractere de final de célula
            rng.Style = "1 - Texto de Tabela"
        Next cel
    Next tbl
End Sub

Sub AjustarTabelasParaJanela()
    Dim tbl As Table
    Dim doc As Document

    Set doc = ActiveDocument

    For Each tbl In doc.Tables
        tbl.AutoFitBehavior (wdAutoFitWindow)
    Next tbl

    MsgBox "Todas as tabelas foram ajustadas para a largura da janela.", vbInformation
End Sub

Sub NegritoCamposTabelaFigura(doc As Document)
    Dim fld As Field
    Dim resultadoTexto As String
    
    Set doc = ActiveDocument
    
    For Each fld In doc.Fields
        ' Somente campos de referência cruzada (FieldRef)
        If fld.Type = wdFieldRef Then
            resultadoTexto = fld.Result.text
            
            ' Verificar se o resultado contém a palavra "Tabela" ou "Figura"
            If InStr(1, resultadoTexto, "Tabela", vbTextCompare) > 0 Or _
               InStr(1, resultadoTexto, "Figura", vbTextCompare) > 0 Then
               
                ' Aplicar negrito ao resultado
                fld.Result.Font.Bold = True
            End If
        End If
    Next fld
End Sub

Sub ItalicizeNonPortugueseText(doc As Document)
    Dim rng As Range
    Dim word As Range
    
    Set rng = doc.Content
    
    For Each word In rng.Words
        ' Skip punctuation and whitespace
        If Trim(word.text) <> "" And word.text Like "*[A-Za-z]*" Then
            If word.LanguageID <> wdPortuguese And word.LanguageID <> wdPortugueseBrazil Then
                word.Font.Italic = True
            End If
        End If
    Next word
End Sub

