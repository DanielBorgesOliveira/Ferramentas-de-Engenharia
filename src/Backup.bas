Attribute VB_Name = "Backup"
Option Explicit

Sub BackupModules()

    ' Set the export path (you can change this to your desired backup location)
    Dim FolderInputPath As String
    FolderInputPath = Library.UseFolderDialog

    ' Check if the export path exists, if not, create it
    If Dir(FolderInputPath, vbDirectory) = "" Then
        MkDir FolderInputPath
    End If

     ' Check if the export path exists, if not, create it
    If Dir(FolderInputPath, vbDirectory) = "" Then
        MkDir FolderInputPath
    End If

    ' Loop through each component in the VBA project
    Dim vbComp As Object
    Dim strModuleName As String
    For Each vbComp In ThisDocument.VBProject.VBComponents
        'MsgBox vbComp.name & " " & vbComp.Type
        ' Get the name of the module and determine the file extension
        Select Case vbComp.Type
            Case 1
                strModuleName = vbComp.name & ".bas"
            Case 3
                strModuleName = vbComp.name & ".frm"
            Case Else
                strModuleName = vbComp.name & ".bas"
        End Select
        
        ' Export the module to the specified path
        vbComp.Export FolderInputPath & "\" & strModuleName
    Next vbComp

    Dim FolderOutputPath As String
    FolderOutputPath = Library.UseFolderDialog & "\FinishHim-" & Format(Now(), "DD-MMM-YYYY-hh-mm-ss") & ".zip"
    
    Call Library.CreateZipFile(CVar(FolderInputPath), CVar(FolderOutputPath))
    
End Sub
