Attribute VB_Name = "Library"
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
    Do Until ShellApp.Namespace(zippedFileFullName).items.Count = ShellApp.Namespace(folderToZipPath).items.Count
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

Function UseFileDialog(Optional DialogType As Integer = msoFileDialogSaveAs) As String
    Dim lngCount As Long
    ' Open the file dialog
    With Application.FileDialog(DialogType)
        .AllowMultiSelect = False
        .Show
        UseFileDialog = .SelectedItems(1)
    End With
End Function

Function propertyExists(propName) As Boolean
    Dim tempObj
    On Error Resume Next
    Set tempObj = ActiveDocument.CustomDocumentProperties.Item(propName)
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

Sub DeleteCustomStyles()
    Dim oStyle As style
    For Each oStyle In ActiveDocument.Styles
        If Not oStyle.BuiltIn Then
            oStyle.Delete
        End If
    Next oStyle
End Sub
