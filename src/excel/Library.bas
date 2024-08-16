Attribute VB_Name = "Library"
Option Private Module

#If VBA7 And Win64 Then
    ' For 64bit version of Excel
    Public Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As LongPtr)
#Else
    ' For 32bit version of Excel
    Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
#End If

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

Sub CreateZipFile(folderToZipPath As Variant, zippedFileFullName As Variant)

    Dim ShellApp As Object
    Dim zipNamespace As Object
    Dim folderNamespace As Object
    Dim fso As Object

    ' Ensure the zip file name ends with .zip
    If Right(zippedFileFullName, 4) <> ".zip" Then
        zippedFileFullName = zippedFileFullName & ".zip"
    End If

    ' Create an empty file if it doesn't exist
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(zippedFileFullName) Then
        ' Create an empty file
        fso.CreateTextFile zippedFileFullName
    End If

    ' Create a Shell Application object
    Set ShellApp = CreateObject("Shell.Application")

    ' Get the namespace for the zip file and the folder to zip
    Set zipNamespace = ShellApp.Namespace(zippedFileFullName)
    Set folderNamespace = ShellApp.Namespace(folderToZipPath)

    ' Check if the namespaces were created successfully
    If zipNamespace Is Nothing Then
        MsgBox "Could not create namespace for the zip file."
        Exit Sub
    End If

    If folderNamespace Is Nothing Then
        MsgBox "Could not create namespace for the folder to zip."
        Exit Sub
    End If

    ' Copy the files & folders into the zip file
    zipNamespace.CopyHere folderNamespace.items

    ' Zipping the files may take a while, create loop to pause the macro until zipping has finished.
    On Error Resume Next
    Do Until zipNamespace.items.Count = folderNamespace.items.Count
        Sleep 1000
    Loop
    On Error GoTo 0

End Sub

Function DecodeBase64(ByVal strData As String, extension As String) As String
    Dim byteData() As Byte
    Dim binaryStream As Object
    Dim imagePath As String

    On Error GoTo ErrorHandler
    
    ' Decode the Base64 string into a byte array
    byteData = DecodeBase64ToByteArray(strData)
    
    ' Initialize an ADODB stream
    Set binaryStream = CreateObject("ADODB.Stream")
    With binaryStream
        .Type = 1 ' Binary data
        .Open
        .Write byteData
        .SaveToFile Environ("TMP") & "\temp." & extension, 2 ' Save to temporary path
        .Close
    End With
    
    ' Return the path to the image file
    imagePath = Environ("TMP") & "\temp." & extension
    DecodeBase64 = imagePath

    Exit Function

ErrorHandler:
    ' Clean up on error
    If Not binaryStream Is Nothing Then
        If binaryStream.State = 1 Then binaryStream.Close
    End If
    Set binaryStream = Nothing
    DecodeBase64 = ""
End Function

Function DecodeBase64ToByteArray(ByVal strData As String) As Byte()
    Dim objXML As Object
    Dim objNode As Object

    ' Create an XML document object
    Set objXML = CreateObject("MSXML2.DOMDocument")
    
    ' Create a node with Base64 data type
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = strData
    
    ' Get the decoded byte array
    DecodeBase64ToByteArray = objNode.nodeTypedValue
    
    ' Clean up
    Set objNode = Nothing
    Set objXML = Nothing
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
    objDocElem.DataType = "bin.base64"

    ' Set binary value
    objDocElem.nodeTypedValue = objStream.Read()

    ' Get base64 value
    EncodeFile = objDocElem.Text

    ' Clean all
    Set objXML = Nothing
    Set objDocElem = Nothing
    Set objStream = Nothing

End Function
