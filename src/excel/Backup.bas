Attribute VB_Name = "Backup"
Sub BackupModules()
    ' Check if programmatic access to the VBA project is allowed
    If Not TrustVBAAccess() Then
        MsgBox "Access to the VBA project is restricted. Please enable programmatic access.", vbCritical
        Exit Sub
    End If

    ' Set the export path (you can change this to your desired backup location)
    Dim FolderInputPath As String
    FolderInputPath = Library.UseFolderDialog
    If FolderInputPath = "" Then
        MsgBox "No folder selected. Exiting.", vbExclamation
        Exit Sub
    End If

    ' Check if the export path exists, if not, create it
    If Dir(FolderInputPath, vbDirectory) = "" Then
        MkDir FolderInputPath
    End If

    ' Loop through each component in the VBA project
    Dim vbComp As Object
    Dim strModuleName As String
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' Get the name of the module and determine the file extension
        Select Case vbComp.Type
            Case 1
                strModuleName = vbComp.Name & ".bas"
            Case 2
                strModuleName = vbComp.Name & ".cls"
            Case 3
                strModuleName = vbComp.Name & ".frm"
            Case Else
                strModuleName = vbComp.Name & ".bas"
        End Select
        
        ' Export the module to the specified path
        vbComp.Export FolderInputPath & "\" & strModuleName
    Next vbComp
    
    ' Copy the current workbook to the backup folder
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim currentWB As Workbook
    
    Dim destFilePath As String

    ' Get the active workbook
    Set currentWB = ActiveWorkbook
    
    Dim sourceFilePath As String
    'sourceFilePath = currentWB.FullName
    sourceFilePath = "C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\0-GERENCIAMENTO\Script\Excel\FinishHim.xltm"

    ' Ensure the destination folder path ends with a backslash
    If Right(FolderInputPath, 1) <> "\" Then
        FolderInputPath = FolderInputPath & "\"
    End If

    ' Create the destination file path
    destFilePath = FolderInputPath & fso.GetFileName(sourceFilePath)
    
    ' Copy the file
    fso.CopyFile sourceFilePath, destFilePath

    ' Clean up
    Set fso = Nothing
    
    Dim FolderOutputPath As String
    FolderOutputPath = Library.UseFolderDialog & "\FinishHim-" & Format(Now(), "DD-MMM-YYYY-hh-mm-ss") & ".zip"
    
    ' Create a zip file from the backup folder
    Call Library.CreateZipFile(CVar(FolderInputPath), CVar(FolderOutputPath))
    
    MsgBox "Backup completed successfully!", vbInformation
End Sub

Function TrustVBAAccess() As Boolean
    On Error Resume Next
    Dim testProj As Object
    Set testProj = ActiveWorkbook.VBProject
    TrustVBAAccess = Not testProj Is Nothing
    On Error GoTo 0
End Function

