Attribute VB_Name = "Tools"
'/*******************************************************************
'* Copyright         : 2024 Daniel Oliveira
'* Email             : danielbo17@hotmail.com
'* File Name         : Tools
'* Description       : Conjunto de ferramentas diversas utilizadas na converção de formatos neutros para formato dos clientes.
'*
'* Revision History  :
'* 16/12/2024      Daniel Borges de Oliveira          Desenvolvimento Inicial
'/******************************************************************/

Option Explicit

Sub CopyContentFromPage3()
    Dim wdApp As Word.Application
    Dim srcDoc As Word.Document, destDoc As Word.Document
    Dim rngStart As Word.Range, rngCopy As Word.Range
    Dim pasteRange As Word.Range
    
    ' Update these paths to the actual locations of your documents
    Dim srcFile As String
    srcFile = "C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\Desktop\1\BdB211813-0000-V-ET0007.docx"
    
    Dim destFile As String
    destFile = "C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\Desktop\1\Samarco-Template - Copy.docx"
    
    ' Create a new instance of Word (or use Application if running from Word)
    Set wdApp = New Word.Application
    wdApp.Visible = True
    
    ' Open the source and destination documents
    Set srcDoc = wdApp.Documents.Open(srcFile)
    Set destDoc = wdApp.Documents.Open(destFile)
    
    ' Get the beginning of page 3 in the source document.
    ' Note: This assumes the source document has at least three pages.
    Set rngStart = srcDoc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, count:=3)
    Set rngCopy = srcDoc.Range(Start:=rngStart.Start, End:=srcDoc.Content.End)
    
    ' Copy the content from page 3 to the end
    rngCopy.Copy
    
    ' Paste at the end of the destination document
    Set pasteRange = destDoc.Content
    pasteRange.Collapse Direction:=wdCollapseEnd
    pasteRange.Paste
    
    ' Save the destination document
    destDoc.Save
    
    ' Optionally close the source document without saving changes
    srcDoc.Close SaveChanges:=False
    
    ' Optional: quit Word if no longer needed
    ' wdApp.Quit
    
    ' Clean up object variables
    Set rngStart = Nothing
    Set rngCopy = Nothing
    Set pasteRange = Nothing
    Set srcDoc = Nothing
    Set destDoc = Nothing
    Set wdApp = Nothing
    
    MsgBox "Content copied from page 3 of the source document to the template."
End Sub

