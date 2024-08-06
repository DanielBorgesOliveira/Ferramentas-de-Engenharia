Attribute VB_Name = "Library_Image"
'Option Explicit
Option Private Module

Function InsertImage(ImageNameBase64 As String, rng As Range)
    
    Dim ImagePath As String
    ImagePath = Environ("TMP") & "\temp.png"
    
    
    
    ' Save byte array to temp file
    Open ImagePath For Binary As #1
       Put #1, 1, DecodeBase64(ImageNameBase64)
    Close #1
    
    ' Insert image from temp file
    rng.InlineShapes.AddPicture FileName:=ImagePath, Range:=rng
    
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
    
    
    
