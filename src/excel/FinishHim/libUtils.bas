Attribute VB_Name = "libUtils"
Option Private Module

#If VBA7 And Win64 Then
    ' For 64bit version of Excel
    Public Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As LongPtr)
#Else
    ' For 32bit version of Excel
    Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
#End If

Sub ApplyFilter(Name As Variant, rng As Range, Field As Integer)
    ' Ensure the range has AutoFilter applied
    If Not rng.Parent.AutoFilterMode Then rng.AutoFilter

    ' Apply the filter
    rng.AutoFilter _
        Field:=Field, _
        Criteria1:=Name, _
        Operator:=xlFilterValues
End Sub

Sub ClearFilter(ws As Worksheet)
    If ws.AutoFilterMode = True Then
      ws.AutoFilterMode = False
    End If
    
    ' Loop through each table (ListObject) in the worksheet
    If ws.ListObjects.count > 0 Then
        Dim tbl As ListObject
        For Each tbl In ws.ListObjects
            ' Check if a filter is applied, and if so, clear it
            If tbl.AutoFilter.FilterMode Then
                tbl.AutoFilter.ShowAllData
            End If
        Next tbl
    End If
End Sub

Function ShowCriteria(Range As Range, DataSheet As Worksheet, Optional Full As Boolean) As Collection
    'https://stackoverflow.com/questions/29773987/get-all-list-of-possible-filter-criteria
    
    Dim OnlyVisible As Collection
    Set OnlyVisible = New Collection
    Dim All As Collection
    Set All = New Collection
    
    ' OnlyVisible criteria
    On Error Resume Next
    For Each R In Range
        v = R.Value
        All.Add v, CStr(v)
        If R.EntireRow.Hidden = False Then
            OnlyVisible.Add v, CStr(v)
        End If
    Next
    On Error GoTo 0
    
    'msg = "Full criteria"
    'For i = 1 To All.Count
    '    msg = msg & vbCrLf & All.Item(i)
    'Next i

    'msg = msg & vbCrLf & vbCrLf & "All criteria"
    'For i = 1 To OnlyVisible.Count
    '    msg = msg & vbCrLf & OnlyVisible.Item(i)
    'Next i
    
    'MsgBox msg
    
    If Full Then
        Set ShowCriteria = All
    Else
        Set ShowCriteria = OnlyVisible
    End If
    
End Function

Function UniqueItems(ByVal R As Range, _
    Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, _
    Optional ByRef count) As Variant
  'Return an array with all unique values in R
  '  and the number of occurrences in Count
  Dim area As Range, Data
  Dim i As Long, j As Long
  Dim Dict As Object 'Scripting.Dictionary
  Set R = Intersect(R.Parent.UsedRange, R)
  If R Is Nothing Then
    UniqueItems = Array()
    Exit Function
  End If
  Set Dict = CreateObject("Scripting.Dictionary")
  Dict.CompareMode = Compare
  For Each area In R.Areas
    Data = area
    If IsArray(Data) Then
      For i = 1 To UBound(Data)
        For j = 1 To UBound(Data, 2)
          If Not Dict.Exists(Data(i, j)) Then
            Dict.Add Data(i, j), 1
          Else
            Dict(Data(i, j)) = Dict(Data(i, j)) + 1
          End If
        Next
      Next
    Else
      If Not Dict.Exists(Data) Then
        Dict.Add Data, 1
      Else
        Dict(Data) = Dict(Data) + 1
      End If
    End If
  Next
  UniqueItems = Dict.Keys
  count = Dict.items
End Function

Function UseFolderDialog(Optional AllowMultiSelect As Boolean = False) As String
    Dim lngCount As Long
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = AllowMultiSelect
        .Show
        UseFolderDialog = .SelectedItems(1)
    End With
End Function

Function UseFileDialog(Optional AllowMultiSelect As Boolean = False) As Variant
    Dim lngCount As Long
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogFilePicker)
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

Function OpenWorkbook(File As String) As Workbook
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.FullName = File Then
            Set OpenWorkbook = wb
            Exit Function
        End If
    Next wb
    Set OpenWorkbook = Workbooks.Open(File)
End Function

Sub CopyToFile(DirName As String, fileName As String, DataSheet As Worksheet)
    DataSheet.UsedRange.SpecialCells(xlCellTypeVisible).Copy
    
    Dim CsvWorkbook As Workbook
    Set CsvWorkbook = Workbooks.Add
    CsvWorkbook.Sheets(1).Paste
    
    Call SetWidth(CsvWorkbook.Sheets(1).Range("BK:BK"), 65)
    Call SetHeight(CsvWorkbook.Sheets(1).Range("A2", Range("A2").End(xlDown)), 65)
    
    Call Hide(CsvWorkbook.Sheets(1), CsvWorkbook.Sheets(1).Range("A:C,H:J,L:BG"))
    
    CsvWorkbook.SaveAs fileName:=DirName & "\" & fileName & ".xlsx", FileFormat:=51, local:=True
    CsvWorkbook.Close SaveChanges:=True
End Sub

Sub Hide(DataSheet As Worksheet, R As Range)
    R.EntireColumn.Hidden = True
End Sub

Sub UnhideAll(DataSheet As Worksheet)
    DataSheet.Columns.EntireColumn.Hidden = False
    DataSheet.rows.EntireRow.Hidden = False
    ' Oculta as linhas que não são importantes para nós.
    DataSheet.rows("1:3").EntireRow.Hidden = True
End Sub

Private Sub CloseWorkbook(File As String)
  Workbooks(File).Close SaveChanges:=True
  'Close method has additional parameters
  'Workbooks.Close(SaveChanges, Filename, RouteWorkbook)
  'Help page: https://docs.microsoft.com/en-us/office/vba/api/excel.workbook.close
End Sub

Private Sub PressEnter()
    Dim Range00 As Range
    Set Range00 = Worksheets("LD").Range("BJ5", Range("BJ5").End(xlDown))
    For Each c In Range00.Cells
        ActiveCell.FormulaR1C1 = _
        "=CONCAT(RC[-23],CHAR(10),RC[-21],CHAR(10),RC[-20],CHAR(10),RC[-19],CHAR(10))"
        Worksheets("LD").Range("BJ" & c.row).Select
    Next
End Sub

Function IsInCollection(oCollection As Collection, sItem As String) As Boolean
    Dim vItem As Variant
    For Each vItem In oCollection
        If vItem = sItem Then
            IsInCollection = True
            Exit Function
        End If
    Next vItem
    IsInCollection = False
End Function

Public Function CollectionToArray(myCol As Collection) As Variant
 
    Dim result  As Variant
    Dim cnt     As Long
 
    ReDim result(myCol.count - 1)
 
    For cnt = 0 To myCol.count - 1
        result(cnt) = myCol(cnt + 1)
    Next cnt
 
    CollectionToArray = result
    
End Function

Function CheckWorkbook(WorkbookName As String) As Boolean
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name = WorkbookName Then
            wb.Activate
            CheckWorkbook = True
            Exit Function
        End If
    Next wb
    CheckWorkbook = False
End Function

Function Duplicated(rng As Range) As Boolean
    
    Dim myArray As Variant
    myArray = UniqueItems(rng)

    Dim x As Integer
    For x = LBound(myArray) To UBound(myArray)
      If Application.WorksheetFunction.CountIfs(rng, myArray(x)) > 1 Then
        Duplicated = True
        Exit Function
      End If
    Next x
End Function

Function sheetExists(sheetToFind As String, Optional InWorkbook As Workbook) As Boolean
    If InWorkbook Is Nothing Then Set InWorkbook = ThisWorkbook
    Dim Sheet As Object
    For Each Sheet In InWorkbook.Sheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
    sheetExists = False
End Function

Function GetArrLength(a As Variant) As Long
   If IsEmpty(a) Then
      GetArrLength = 0
   Else
      GetArrLength = UBound(a) - LBound(a) + 1
   End If
End Function

Sub CreateZipFile(folderToZipPath As Variant, zippedFileFullName As Variant)

    Dim ShellApp As Object
    
    'Create an empty zip file
    Open zippedFileFullName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    
    'Copy the files & folders into the zip file
    Set ShellApp = CreateObject("Shell.Application")
    
    'MsgBox zippedFileFullName
    'MsgBox folderToZipPath
    
    ShellApp.Namespace(zippedFileFullName).CopyHere ShellApp.Namespace(folderToZipPath).items
    
    'Zipping the files may take a while, create loop to pause the macro until zipping has finished.
    On Error Resume Next
    Do Until ShellApp.Namespace(zippedFileFullName).items.count = ShellApp.Namespace(folderToZipPath).items.count
        Sleep 1000
    Loop
    On Error GoTo 0

End Sub

Sub RefreshPivotTables(wb As Workbook)
    ' Refresh all connections and Pivot Tables in the workbook
    wb.RefreshAll
    
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    ' Loop through each worksheet
    For Each ws In wb.Worksheets
        ' Loop through each Pivot Table in the worksheet
        For Each pt In ws.PivotTables
            ' Refresh the Pivot Table
            pt.RefreshTable
            pt.Update
            
            ' Show the name of the Pivot Table in a message box or immediate window (Debug.Print)
            ' MsgBox "Pivot Table refreshed: " & pt.name
            ' Alternatively, use the following line to print the name in the Immediate Window (Ctrl + G)
            ' Debug.Print "Pivot Table refreshed: " & pt.Name
        Next pt
    Next ws
End Sub

Sub ScreenUpdateOff()
    ' Turn off screen updating and calculation to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
End Sub

Sub ScreenUpdateOn()
' Turn on settings after completion
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Function GetSelectedItems(lBox As MSForms.ListBox) As Variant
    'returns an array of selected items in a ListBox
    Dim tmpArray() As Variant
    Dim i As Integer
    Dim selCount As Integer
    selCount = -1
    '## Iterate over each item in the ListBox control:
    For i = 0 To lBox.ListCount - 1
        '## Check to see if this item is selected:
        If lBox.Selected(i) = True Then
            '## If this item is selected, then add it to the array
            selCount = selCount + 1
            ReDim Preserve tmpArray(selCount)
            tmpArray(selCount) = lBox.List(i)
        End If
    Next

    If selCount = -1 Then
        '## If no items were selected, return an empty array
        GetSelectedItems = Array() ' Empty array
    Else:
        '## Otherwise, return the array of selected items
        GetSelectedItems = tmpArray
    End If
End Function

Function GUID$(Optional lowercase As Boolean, Optional parens As Boolean)
    Dim k&, h$
    GUID = Space(36)
    For k = 1 To Len(GUID)
        Randomize
        Select Case k
            Case 9, 14, 19, 24: h = "-"
            Case 15:            h = "4"
            Case 20:            h = Hex(rnd * 3 + 8)
            Case Else:          h = Hex(rnd * 15)
        End Select
        Mid$(GUID, k, 1) = h
    Next
    If lowercase Then GUID = LCase$(GUID)
    If parens Then GUID = "{" & GUID & "}"
End Function

Function TrustVBAAccess() As Boolean
    On Error Resume Next
    Dim testProj As Object
    Set testProj = ActiveWorkbook.VBProject
    TrustVBAAccess = Not testProj Is Nothing
    On Error GoTo 0
End Function

Sub BackupModules(fileName As String)
    ' Check if programmatic access to the VBA project is allowed
    If Not TrustVBAAccess() Then
        MsgBox "Access to the VBA project is restricted. Please enable programmatic access.", vbCritical
        Exit Sub
    End If

    ' Set the export path (you can change this to your desired backup location)
    Dim FolderInputPath As String
    'MsgBox "Indique a pasta para exportar os módulos."
    'FolderInputPath = libUtils.UseFolderDialog
    'If FolderInputPath = "" Then
    '    MsgBox "No folder selected. Exiting.", vbExclamation
    '    Exit Sub
    'End If
    FolderInputPath = Environ("temp") & "\" & GUID(lowercase:=True)

    ' Check if the export path exists, if not, create it
    If Dir(FolderInputPath, vbDirectory) = "" Then
        MkDir FolderInputPath
    End If
    
    ' Ensure the destination folder path ends with a backslash
    If Right(FolderInputPath, 1) <> "\" Then
        FolderInputPath = FolderInputPath & "\"
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
    
    ' Get the active workbook
    Dim currentWB As Workbook
    Set currentWB = ActiveWorkbook
    
    Dim sourceFilePaths As Variant
    MsgBox "Indique os arquivos que compõe a pasta de trabalho."
    sourceFilePaths = UseFileDialog(AllowMultiSelect:=True)
    
    For Each sourceFilePath In sourceFilePaths
    ' Create the destination file path
        Dim destFilePath As String
        destFilePath = FolderInputPath & fso.GetFileName(sourceFilePath)
        
        ' Copy the file
        fso.CopyFile sourceFilePath, destFilePath
    Next sourceFilePath
    
    'fso.DeleteFolder FolderInputPath

    ' Clean up
    Set fso = Nothing
    
    Dim FolderOutputPath As String
    MsgBox "Indique a pasta onde o backup será salvo."
    FolderOutputPath = UseFolderDialog(AllowMultiSelect:=False) & "\" & fileName & "-" & format(Now(), "DD-MMM-YYYY-hh-mm-ss") & ".zip"
    
    ' Create a zip file from the backup folder
    Call libUtils.CreateZipFile(CVar(FolderInputPath), CVar(FolderOutputPath))
    
    'MsgBox "Backup completed successfully!", vbInformation
End Sub

Sub ExportSheetToNewWorkbook(wsSource As Worksheet, fileName As String, filePath As String)
    ' Create a new workbook for exporting the data
    Dim wbExport As Workbook
    Set wbExport = Workbooks.Add
    
    ' Remove default sheet in the new workbook if necessary
    Application.DisplayAlerts = False
    'wbExport.Sheets(1).Delete
    Application.DisplayAlerts = True
    
    ' Copy the entire source worksheet to the new workbook
    wsSource.Copy before:=wbExport.Sheets(1) ' Copies the entire worksheet
    
    ' Rename the new sheet if needed (optional)
    'wbExport.Sheets(1).name = fileName
    
    ' Save the new workbook as an .xlsx file
    wbExport.SaveAs fileName:=filePath & "\" & fileName, FileFormat:=xlOpenXMLWorkbook, local:=True
    
    ' Close the new workbook after saving
    wbExport.Close SaveChanges:=True
End Sub

Function ListSheets(wb As Workbook, Optional ignoreList As Variant = Empty) As Variant
    Dim Worksheets() As String
    Dim ws As Worksheet
    Dim size As Long: size = 0
    Dim i As Long
    Dim skip As Boolean

    ReDim Worksheets(0 To wb.Worksheets.count - 1)

    For Each ws In wb.Worksheets
        skip = False
        
        ' Only check ignoreList if it was provided
        If Not IsEmpty(ignoreList) Then
            For i = LBound(ignoreList) To UBound(ignoreList)
                If ws.Name = ignoreList(i) Then
                    skip = True
                    Exit For
                End If
            Next i
        End If
        
        If Not skip Then
            Worksheets(size) = ws.Name
            size = size + 1
        End If
    Next ws

    If size > 0 Then
        ReDim Preserve Worksheets(0 To size - 1)
        ListSheets = Worksheets
    Else
        ListSheets = Array() ' Return empty array
    End If
End Function

Function ISO8601TimeStamp(Data As Date) As String
    ' Esse é o formato aceito pelo Planner.
    ISO8601TimeStamp = format(Data, "yyyy-mm-ddTHH:MM:SSZ")
End Function

Function DecodeBase64(ByVal strData As String, extension As String) As String
    Dim byteData() As Byte
    Dim binaryStream As Object
    Dim fso As Object
    Dim tempFileName As String
    Dim imagePath As String

    On Error GoTo ErrorHandler
    
    ' Decode the Base64 string into a byte array
    byteData = DecodeBase64ToByteArray(strData)
    
    ' Generate a random temporary file name
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempFileName = fso.GetTempName
    tempFileName = Replace(tempFileName, ".tmp", "." & extension)
    imagePath = Environ("TMP") & "\" & tempFileName
    
    ' Initialize an ADODB stream
    Set binaryStream = CreateObject("ADODB.Stream")
    With binaryStream
        .Type = 1 ' Binary data
        .Open
        .Write byteData
        .SaveToFile imagePath, 2 ' Save to temporary path
        .Close
    End With
    
    ' Return the path to the image file
    DecodeBase64 = imagePath

    Exit Function

ErrorHandler:
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
    objNode.text = strData
    
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
    EncodeFile = objDocElem.text

    ' Clean all
    Set objXML = Nothing
    Set objDocElem = Nothing
    Set objStream = Nothing

End Function

Private Function isStringEmpty(text As String) As Boolean
    If Trim(text & vbNullString) = vbNullString Then
        isStringEmpty = True
    Else
        isStringEmpty = False
    End If
End Function

Function propertyExists(wb As Workbook, propName As String) As Boolean
    Dim tempObj As Object
    On Error Resume Next
    Set tempObj = wb.CustomDocumentProperties(propName)
    propertyExists = (Not tempObj Is Nothing)
    On Error GoTo 0
End Function

Sub UpdateProperty(wb As Workbook, propName As String, propValue As String)
    If propertyExists(wb, propName) Then
        wb.CustomDocumentProperties(propName).Value = propValue
    Else
        wb.CustomDocumentProperties.Add _
            Name:=propName, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeString, _
            Value:=propValue
    End If
End Sub

Sub ExportAllWorksheetsAsPDF(wb As Workbook, OutputFile As String)
    ' Ensure the output file is a PDF
    OutputFile = Replace(OutputFile, ".xlsx", ".pdf")
    OutputFile = Replace(OutputFile, ".xls", ".pdf")
    
    ' Export all worksheets to a single PDF
    wb.ExportAsFixedFormat Type:=xlTypePDF, fileName:=OutputFile, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
       
    ' Save the workbook if there are unsaved changes
    If wb.Saved = False Then wb.Save
End Sub

Sub DeleteAllShapes(ws As Worksheet, KeepShapeName As String)
    Dim shp As Shape ' Variable to hold each shape during the loop

    On Error Resume Next ' Prevent the code from stopping due to unexpected errors

    ' Loop through all shapes in the specified worksheet
    For Each shp In ws.Shapes
        ' Check if the shape's name does not contain the specified name (KeepShapeName)
        ' If the name does not match, delete the shape
        If InStr(1, shp.Name, KeepShapeName, vbTextCompare) = 0 Then
            shp.Delete ' Delete the shape
        End If
    Next shp

    On Error GoTo 0 ' Turn off error handling after the loop
End Sub

Sub DeleteAllNotes(ws As Worksheet)
    ' Atualizado em 11/07/2025
    Dim cell As Range
    Dim commentRange As Range
    Dim commentCount As Long

    ' Verifica se há comentários (notas antigas)
    commentCount = ws.Comments.count
    If commentCount > 0 Then
        On Error Resume Next
        Set commentRange = ws.Cells.SpecialCells(xlCellTypeComments)
        On Error GoTo 0 ' Desativa o desvio de erro após a tentativa

        If Not commentRange Is Nothing Then
            For Each cell In commentRange
                If Not cell.Comment Is Nothing Then
                    cell.Comment.Delete
                End If
            Next cell
        End If
    End If

    ' Também verifica comentários modernos (Threaded Comments)
    If ws.CommentsThreaded.count > 0 Then
        Dim cmt As CommentThreaded
        For Each cmt In ws.CommentsThreaded
            cmt.Delete
        Next cmt
    End If
End Sub

Sub DeleteAllThreadedComments(ws As Worksheet)
    Dim tc As CommentThreaded
    On Error Resume Next
    For Each tc In ws.CommentsThreaded
        tc.Delete
    Next tc
End Sub

Sub setupPage(wb As Workbook, ws As Worksheet, pagesetupConfiguration As Object)
    On Error GoTo CleanUp
    
    ' Remove the AutoFilter and configure page settings
    ws.AutoFilterMode = False
    
    ' Remove all page breaks.
    ws.ResetAllPageBreaks
    
    With ws.PageSetup
        .printArea = "" ' Clear the print area
        .PrintTitleRows = "'" & ws.Name & "'!" & pagesetupConfiguration("PrintTitleRows") ' Rows you want to repeat at the top

        
        .Zoom = False ' Turn off zoom to use FitToPages settings
        .FitToPagesWide = pagesetupConfiguration("FitToPagesWide") ' Fit to one page wide
        .FitToPagesTall = False ' Let Excel determine the number of pages tall
        
        ' Set margins in points
        .TopMargin = pagesetupConfiguration("TopMargin")
        .BottomMargin = pagesetupConfiguration("BottomMargin")
        .LeftMargin = pagesetupConfiguration("LeftMargin")
        .RightMargin = pagesetupConfiguration("RightMargin")
        .HeaderMargin = pagesetupConfiguration("HeaderMargin")
        .FooterMargin = pagesetupConfiguration("FooterMargin")
        
        ' Set the header and footer text
        .RightHeader = pagesetupConfiguration("RightHeader")
        .LeftFooter = pagesetupConfiguration("LeftFooter")
        .CenterFooter = pagesetupConfiguration("CenterFooter")
        
        .Orientation = pagesetupConfiguration("Orientation")
        .Order = pagesetupConfiguration("Order")
    End With
    
    ' Unfreeze panes if frozen, without activating the sheet
    If ws.Parent.Windows(1).FreezePanes Then
        ws.Parent.Windows(1).FreezePanes = False
    End If
    
    ' Set the worksheet to Page Break Preview and adjust zoom
    With ws.Parent.Windows(1)
        .View = xlPageBreakPreview
        .Zoom = 100
    End With
    
    Exit Sub

CleanUp:
    If Err.Number <> 0 Then
        Debug.Print "setupPage[ERROR]: " & Err.Description, vbExclamation
    End If
End Sub

Sub addStyles(wb As Workbook, Optional fontName As String = "Arial")
    'Call libUtils.createStyle(wb, "Arial", "6ptLeft", 6, False, False, False, True, RGB(0, 0, 0), xlLeft, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "6ptLeftBold", 6, True, False, False, True, RGB(0, 0, 0), xlLeft, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "6ptCenterBold", 6, True, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "8ptLeft", 8, False, False, False, True, RGB(0, 0, 0), xlLeft, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "8ptLeftBold", 8, True, False, False, True, RGB(0, 0, 0), xlLeft, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "8ptCenter", 8, False, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "8ptCenterBold", 8, True, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "9ptCenter", 9, True, False, False, True, RGB(0, 0, 0), xlLeft, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "9ptLeftBold", 9, True, False, False, True, RGB(0, 0, 0), xlLeft, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "9ptCenterBold", 9, True, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "10ptCenterBold", 10, True, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "10ptLeftBold", 10, True, False, False, True, RGB(0, 0, 0), xlLeft, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "11ptCenterBold", 11, True, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "12ptCenterBold", 12, True, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "12ptCenter", 12, False, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "26ptCenter", 26, False, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    'Call libUtils.createStyle(wb, "Arial", "26ptCenterBold", 26, True, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous)))
    
    Dim size As Single
    Dim isBold As Variant
    Dim alignment As Variant
    Dim styleSuffix As String
    Dim styleName As String
    
    Dim borderDef As Variant
    borderDef = Array( _
        Array(xlEdgeLeft, xlContinuous), _
        Array(xlEdgeRight, xlContinuous), _
        Array(xlEdgeTop, xlContinuous), _
        Array(xlEdgeBottom, xlContinuous) _
    )
    
    For size = 6 To 26
        For Each isBold In Array(False, True)
            For Each alignment In Array(xlHAlignLeft, xlHAlignCenter)
                
                ' Construct style suffix
                styleSuffix = ""
                If isBold Then styleSuffix = styleSuffix & "_Bold" Else styleSuffix = styleSuffix & "_Regular"
                If alignment = xlHAlignLeft Then styleSuffix = styleSuffix & "_Left" Else styleSuffix = styleSuffix & "_Center"
                
                ' Final style name
                styleName = size & "pt" & styleSuffix
                
                ' Create the style
                Call createStyle( _
                    wb:=wb, _
                    fontName:=fontName, _
                    styleName:=styleName, _
                    fontSize:=size, _
                    fontBold:=CBool(isBold), _
                    HorizontalAlignment:=CLng(alignment), _
                    Border:=borderDef)
                
            Next alignment
        Next isBold
    Next size
End Sub

Function createStyle( _
    wb As Workbook, _
    fontName As String, _
    styleName As String, _
    fontSize As Single, _
    Optional fontBold As Boolean = False, _
    Optional fontItalic As Boolean = False, _
    Optional fontUnderline As Boolean = False, _
    Optional wrapText As Boolean = True, _
    Optional fontColor As Long = -1, _
    Optional HorizontalAlignment As XlHAlign = xlCenter, _
    Optional VerticalAlignment As XlVAlign = xlCenter, _
    Optional Border As Variant = Nothing) As Style
    
    Dim customStyle As Style
    Dim i As Integer
    
    ' Set default color if not provided
    If fontColor = -1 Then fontColor = RGB(0, 0, 0)
    
    ' Check if the style exists, and create it if it doesn't
    On Error Resume Next
    Set customStyle = wb.Styles(styleName)
    If customStyle Is Nothing Then
        Set customStyle = wb.Styles.Add(Name:=styleName)
    End If
    On Error GoTo 0
    
    ' Update the style properties
    With customStyle
        .Font.Name = fontName
        .Font.size = fontSize
        .Font.Bold = fontBold
        .Font.Italic = fontItalic
        .Font.Underline = IIf(fontUnderline, xlUnderlineStyleSingle, xlUnderlineStyleNone)
        .Font.Color = fontColor
        .HorizontalAlignment = HorizontalAlignment
        .VerticalAlignment = VerticalAlignment
        .wrapText = wrapText
        
        ' Apply borders if provided
        If Not IsMissing(Border) Then
            For i = 0 To UBound(Border)
                If IsArray(Border(i)) And UBound(Border(i)) >= 1 Then
                    With .Borders(Border(i)(0))
                        .LineStyle = Border(i)(1)
                        .Color = RGB(0, 0, 0)
                        .Weight = xlThin
                    End With
                End If
            Next i
        End If
    End With
    
    ' Return the created or updated style
    Set createStyle = customStyle
End Function

Function getCellsHeight(ws As Worksheet) As Variant
    Dim reference() As Double
    Dim lastRow As Long
    Dim iRow As Range
    
    ' Get the last used row to initialize the array with the correct size
    lastRow = ws.UsedRange.rows(ws.UsedRange.rows.count).row
    ReDim reference(1 To lastRow)
    
    ' Loop through each row in the used range and store the row height
    For Each iRow In ws.UsedRange.rows
        reference(iRow.row) = iRow.RowHeight
    Next iRow
    
    ' Return the array containing the row heights
    getCellsHeight = reference
End Function

Sub insertLogo(wb As Workbook, ws As Worksheet, imagePath As String, targetCell As Range, Optional scaleFactor As Double = 0.8, Optional moveAndSize As Boolean = False)
    Dim img As Shape
    
    ' Error handling
    On Error GoTo ErrHandler
    
    ' Insert the image
    Set img = ws.Shapes.AddPicture( _
        fileName:=imagePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoCTrue, _
        Left:=0, _
        Top:=0, _
        width:=-1, _
        Height:=-1 _
    )
    
    ' Calculate the position to center the image
    With img
        .LockAspectRatio = msoTrue ' Maintain aspect ratio of the image
        
        ' Scale the image within the cell, with scaling factor adjustable
        If .width > targetCell.width Or .Height > targetCell.Height Then
            If .width / targetCell.width > .Height / targetCell.Height Then
                .width = targetCell.width * scaleFactor
            Else
                .Height = targetCell.Height * scaleFactor
            End If
        Else
            .width = .width * scaleFactor
            .Height = .Height * scaleFactor
        End If
        
        ' Center the image within the target cell
        .Left = targetCell.Left + (targetCell.width - .width) / 2
        .Top = targetCell.Top + (targetCell.Height - .Height) / 2
        
        ' Set the image to move and size with cells or remain fixed
        If moveAndSize Then
            .Placement = xlMoveAndSize
        Else
            .Placement = xlMove
        End If
    End With

    Exit Sub
    
ErrHandler:
    Debug.Print "insertLogo[ERROR]: " & Err.Description, vbExclamation, "Error"
End Sub

Sub insertRows(wb As Workbook, ws As Worksheet, rows As String)
    On Error GoTo CleanUp
    
    ' Validate rows parameter (optional)
    If Not IsNumeric(Replace(rows, ":", "")) Then
        Debug.Print "insertRows[ERROR]: Invalid row reference: " & rows, vbExclamation
        Exit Sub
    End If

    ' Insert rows
    With ws
        .rows(rows).Insert Shift:=xlDown
    End With

CleanUp:
    If Err.Number <> 0 Then
        Debug.Print "insertRows[ERROR]: " & Err.Description, vbExclamation
    End If
End Sub

Sub ApplyFormatting(rng As Range, config As Object)
    With rng
        If config.Exists("Borders.LineStyle") Then .Borders.LineStyle = config("Borders.LineStyle")
        If config.Exists("Borders.ColorIndex") Then .Borders.ColorIndex = config("Borders.ColorIndex")
        If config.Exists("Borders.TintAndShade") Then .Borders.TintAndShade = config("Borders.TintAndShade")
        If config.Exists("Borders.Weight") Then .Borders.Weight = config("Borders.Weight")
        
        If config.Exists("Font.Name") Then .Font.Name = config("Font.Name")
        If config.Exists("Font.Size") Then .Font.size = config("Font.Size")
        If config.Exists("Font.Bold") Then .Font.Bold = config("Font.Bold")
        If config.Exists("Font.Color") Then .Font.Color = config("Font.Color")
        
        If config.Exists("HorizontalAlignment") Then .HorizontalAlignment = config("HorizontalAlignment")
        If config.Exists("VerticalAlignment") Then .VerticalAlignment = config("VerticalAlignment")
        If config.Exists("wrapText") Then .wrapText = config("wrapText")
    End With
End Sub

Sub AppendToArray(ByRef arr() As Variant, ByVal newElement As Variant)
    Dim currentLength As Long

    ' Handle uninitialized array case
    On Error Resume Next
    currentLength = UBound(arr) - LBound(arr) + 1
    If Err.Number <> 0 Then
        currentLength = 0
        ReDim arr(0 To 0) ' Initialize as an empty array
        Err.Clear
    End If
    On Error GoTo 0

    ' Resize the array and append the new element
    ReDim Preserve arr(LBound(arr) To LBound(arr) + currentLength)
    arr(LBound(arr) + currentLength) = newElement
End Sub

Sub hideRow(wb As Workbook, ws As Worksheet, rows As String)
    ws.rows(rows).EntireRow.Hidden = True
End Sub

Sub SetColumnWidth(ws As Worksheet, items As Variant)
    With ws
        Dim item As Variant
        For Each item In items
            .Columns(item(0)).ColumnWidth = item(1)
        Next
    End With
End Sub

Sub insertColumns(wb As Workbook, ws As Worksheet, items As Variant)
    With ws
        Dim item As Variant
        For Each item In items
            .Columns(item(0)).Insert Shift:=item(1)
        Next
    End With
End Sub

Sub pasteColumns(wb As Workbook, ws As Worksheet, referenceColumn As String, destinationColumn As String)
    ws.Range(referenceColumn).Copy
    ws.Range(destinationColumn).Insert
End Sub

Function getNextLetter(letter As String) As String
    getNextLetter = Chr(Asc(letter) + 1)
End Function

Function columnToLetter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Sub SetHeight(wb As Workbook, ws As Worksheet, items As Variant)
    With ws
        Dim item As Variant
        For Each item In items
            .rows(item(0)).RowHeight = item(1)
        Next
    End With
End Sub

Sub insertText(wb As Workbook, ws As Worksheet, rng As Range)
    Dim borderStyle As Object
    Set borderStyle = CreateObject("Scripting.Dictionary")
    borderStyle.Add "xlContinuous", 1
    borderStyle.Add "xlDash", -4115
    borderStyle.Add "xlDashDot", 4
    borderStyle.Add "xlDashDotDot", 5
    borderStyle.Add "xlDot", -4118
    borderStyle.Add "xlDouble", -4119
    borderStyle.Add "xlLineStyleNone", -4142
    borderStyle.Add "xlSlantDashDot", 13
    
    Dim fontColor As String
    
    Dim row As Range
    For Each row In rng.rows
        Debug.Print "Número da linha: " & row.row & vbCrLf & "Range: " & row.Cells(1, 1).Value & vbCrLf & "Texto: " & row.Cells(1, 2).Value & vbCrLf & "Estilo de Texto: " & row.Cells(1, 3).Value & vbCrLf _
        & "Borda Top: " & row.Cells(1, 4).Value & vbCrLf & "Borda Right: " & row.Cells(1, 5).Value & vbCrLf & "Borda Bottom: " & row.Cells(1, 6).Value & vbCrLf & "Borda Left: " & row.Cells(1, 7).Value & vbCrLf
        
        fontColor = ParseColorString(row.Cells(1, 4).Value)
        
        With ws.Range(row.Cells(1, 1).Value)
            .Value = row.Cells(1, 2).Value
            .Style = row.Cells(1, 3).Value
            .Font.Color = fontColor
            .Merge
            .Borders(xlEdgeTop).LineStyle = borderStyle(row.Cells(1, 5).Value)
            .Borders(xlEdgeRight).LineStyle = borderStyle(row.Cells(1, 6).Value)
            .Borders(xlEdgeBottom).LineStyle = borderStyle(row.Cells(1, 7).Value)
            .Borders(xlEdgeLeft).LineStyle = borderStyle(row.Cells(1, 8).Value)
        End With
    Next
End Sub

Sub ProtectWorksheet(wb As Workbook, Optional ws As Worksheet = Nothing, Optional lockedRange As Range = Nothing, Optional password As String = "Abcd1234")
    Dim targetSheet As Worksheet
    Dim sheetList As Collection
    Set sheetList = New Collection

    ' Determine which sheets to protect
    If ws Is Nothing Then
        ' Protect all worksheets
        For Each targetSheet In wb.Worksheets
            sheetList.Add targetSheet
        Next targetSheet
    Else
        sheetList.Add ws
    End If

    ' Apply protection
    For Each targetSheet In sheetList
        With targetSheet
            .Unprotect password ' In case it's already protected
            .Protect _
                password:=password, _
                DrawingObjects:=True, _
                Contents:=True, _
                UserInterfaceOnly:=False, _
                AllowFormattingCells:=False, _
                AllowFormattingColumns:=False, _
                AllowFormattingRows:=False, _
                AllowInsertingColumns:=False, _
                AllowInsertingRows:=False, _
                AllowInsertingHyperlinks:=False, _
                AllowDeletingColumns:=False, _
                AllowDeletingRows:=False, _
                AllowSorting:=False, _
                AllowFiltering:=False, _
                AllowUsingPivotTables:=False
            
            ' Lock the provided range if specified (only on single sheet mode)
            If Not lockedRange Is Nothing And sheetList.count = 1 Then
                lockedRange.Locked = True
            End If
        End With
    Next targetSheet
End Sub


Sub UnprotectWorksheet(ws As Worksheet, Optional password As String = "Abcd1234")
    On Error Resume Next
    Call ws.Unprotect(password)
    On Error GoTo 0
    
    ws.Cells.Locked = False
End Sub

Function getLogoBase64(Cliente As String) As String
    ' https://base64.guru/converter/encode/file
    ' https://www.dcode.fr/text-splitter
    
    Dim LogoBase64 As String
    LogoBase64 = ""
    
    If Cliente = "Samarco" Then
        LogoBase64 = LogoBase64 & "PHN2ZyB2ZXJzaW9uPSIxLjEiIGlkPSJDYW1hZGFfMSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIiB4bWxuczp4bGluaz0iaH"
        LogoBase64 = LogoBase64 & "R0cDovL3d3dy53My5vcmcvMTk5OS94bGluayIgeD0iMCIgeT0iMCIgdmlld0JveD0iMCAwIDEwMDAgMjA1LjMiIHN0eWxlPSJlbmFibGUtYmFj"
        LogoBase64 = LogoBase64 & "a2dyb3VuZDpuZXcgMCAwIDEwMDAgMjA1LjMiIHhtbDpzcGFjZT0icHJlc2VydmUiPjxzdHlsZT4uc3Qxe2ZpbGw6IzAwNDA3MX08L3N0eWxlPj"
        LogoBase64 = LogoBase64 & "xkZWZzPjxwYXRoIGlkPSJTVkdJRF8xXyIgZD0iTTAgMGgxMDAwdjIwNS4zSDB6Ii8+PC9kZWZzPjxjbGlwUGF0aCBpZD0iU1ZHSURfMl8iPjx1"
        LogoBase64 = LogoBase64 & "c2UgeGxpbms6aHJlZj0iI1NWR0lEXzFfIiBzdHlsZT0ib3ZlcmZsb3c6dmlzaWJsZSIvPjwvY2xpcFBhdGg+PGcgc3R5bGU9ImNsaXAtcGF0aD"
        LogoBase64 = LogoBase64 & "p1cmwoI1NWR0lEXzJfKSI+PHBhdGggY2xhc3M9InN0MSIgZD0iTTYxLjIgMTAyLjNjLS4xLTIxLjEgMTYuOS0zOC4yIDM4LTM4LjMgMTEuMSAw"
        LogoBase64 = LogoBase64 & "IDIxLjcgNC43IDI5IDEzLjFsNTUuMS00OWMtNDEuOC0zNy41LTEwOS44LTM3LjUtMTUyIDBzLTQxLjggOTcuOSAwIDEzNS40bDQwLTM1LjNjLT"
        LogoBase64 = LogoBase64 & "YuNi03LTEwLjItMTYuMy0xMC4xLTI1LjlNMjU2IDEyMS40czEyLjYgMTEuMiAyNi42IDExLjJjNS40IDAgMTEuNS0yLjIgMTEuNS04LjYgMC0x"
        LogoBase64 = LogoBase64 & "My4zLTUwLjEtMTMtNTAuMS00Ni44IDAtMjAuNSAxNy4zLTM0LjIgMzkuNi0zNC4yczM1LjYgMTIuNiAzNS42IDEyLjZsLTExLjUgMjIuM2MtNi"
        LogoBase64 = LogoBase64 & "45LTUuOC0xNS41LTkuMi0yNC41LTkuNy01LjggMC0xMS45IDIuMi0xMS45IDguNiAwIDE0IDUwLjEgMTEuNSA1MC4xIDQ2LjUgMCAxOC40LTE0"
        LogoBase64 = LogoBase64 & "IDM0LjYtMzkuMyAzNC42LTE1LjEuMi0yOS43LTUuNS00MC43LTE1LjhsMTQuNi0yMC43ek0zNzYuMyA3MC4ycy0yLjUgMTEuOS00LjcgMTkuMW"
        LogoBase64 = LogoBase64 & "wtNi41IDIxLjZoMjJMMzgxIDg5LjNjLTIuMi03LjItNC43LTE5LjEtNC43LTE5LjF6bTE3LjMgNjIuM0gzNTlsLTYuOCAyMy40aC0yOC4xbDM3"
        LogoBase64 = LogoBase64 & "LjgtMTExLjNoMjguOGwzNy41IDExMS4zaC0yNy43bC02LjktMjMuNHpNNDQ2LjUgNDQuN0g0NzZsMTUuOCA0Ny4yYzIuNSA2LjggNS44IDE4ID"
        LogoBase64 = LogoBase64 & "UuOCAxOGguNHMyLjktMTEuMiA1LjQtMThsMTYuMi00Ny4yaDI5LjVsOSAxMTEuM2gtMjdsLTMuNi01MC4xYy0uNC04LjMgMC0xOC43IDAtMTgu"
        LogoBase64 = LogoBase64 & "N2gtLjRzLTMuNiAxMS41LTYuMSAxOC43bC0xMS41IDMyaC0yMy40bC0xMS41LTMyLTYuNS0xOC43cy40IDEwLjQgMCAxOC43bC0zLjYgNTAuMW"
        LogoBase64 = LogoBase64 & "gtMjcuNGw5LjQtMTExLjN6TTYxOSA3MC4ycy0yLjIgMTEuOS00LjMgMTkuMWwtNi41IDIxLjZoMjJsLTYuMS0yMS42Yy0yLjItNy4yLTQuNy0x"
        LogoBase64 = LogoBase64 & "OS4xLTQuNy0xOS4xaC0uNHptMTcuNyA2Mi4zaC0zNC42bC02LjggMjMuNGgtMjguMUw2MDUgNDQuN2gyOC44TDY3MS4zIDE1NmgtMjcuN2wtNi"
        LogoBase64 = LogoBase64 & "45LTIzLjV6TTcyMC45IDk1LjFjOC42IDAgMTQtNC43IDE0LTEzLjdzLTIuOS0xMy4zLTE2LjYtMTMuM0g3MDl2MjdoMTEuOXpNNjgyIDQ0Ljdo"
        LogoBase64 = LogoBase64 & "MzguNWMxMS41IDAgMTYuOS43IDIxLjYgMi45IDEyLjYgNC43IDIwLjUgMTUuOCAyMC41IDMycy01LjggMjQuOC0xNi42IDMwLjJsNCA2LjggMj"
        LogoBase64 = LogoBase64 & "EuNiAzOC45aC0zMC4yTDcyMS42IDExOEg3MDl2MzcuOGgtMjdWNDQuN3pNODM0LjcgNDIuOWMyNy40IDAgNDEuNCAxNS44IDQxLjQgMTUuOGwt"
        LogoBase64 = LogoBase64 & "MTIuNiAyMC41cy0xMi4yLTExLjUtMjcuNy0xMS41Yy0yMS4yIDAtMzEuMyAxNS44LTMxLjMgMzEuN3MxMC44IDMzLjUgMzEuMyAzMy41IDI5Lj"
        LogoBase64 = LogoBase64 & "UtMTMuNyAyOS41LTEzLjdsMTQgMTkuOGMtMTEuNyAxMi4xLTI3LjggMTguOS00NC43IDE4LjctMzQuOSAwLTU4LjMtMjQuOC01OC4zLTU3LjZz"
        LogoBase64 = LogoBase64 & "MjQuOS01Ny4yIDU4LjQtNTcuMk05NDIgMTMyLjljMTYuNiAwIDI5LjktMTQuNCAyOS45LTMzLjFzLTEzLjMtMzItMjkuOS0zMi0yOS45IDEzLj"
        LogoBase64 = LogoBase64 & "ctMjkuOSAzMiAxMy40IDMzLjEgMjkuOSAzMy4xbTAtOTBjMzMuNSAwIDU4IDI0LjggNTggNTYuOSAwIDMyLTI1LjkgNTgtNTcuOSA1OHMtNTgt"
        LogoBase64 = LogoBase64 & "MjUuOS01OC01Ny45di0uMWMtLjEtMzIuMSAyNC40LTU2LjkgNTcuOS01Ni45Ii8+PHBhdGggZD0ibTE2OC4yIDQxLjgtNDAgMzUuM2M2LjMgNi"
        LogoBase64 = LogoBase64 & "45IDkuNyAxNS45IDkuNyAyNS4yIDAgMjEuMy0xNy4zIDM4LjUtMzguNSAzOC41LTEwLjctLjEtMjAuOS00LjYtMjguMS0xMi42bC01NS4xIDQ5"
        LogoBase64 = LogoBase64 & "YzQyLjEgMzcuNSAxMDkuOCAzNy41IDE1MiAwczQxLjctOTggMC0xMzUuNCIgc3R5bGU9ImZpbGw6I2E3YjZiZiIvPjwvZz48L3N2Zz4="
    ElseIf Cliente = "AngloAmerican" Or Cliente = "AngloAmerican1" Then
        LogoBase64 = LogoBase64 & "PHN2ZyB2ZXJzaW9uPSIxLjIiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgdmlld0JveD0iMCAwIDE1ODUgMzQ5IiB3aWR0aD"
        LogoBase64 = LogoBase64 & "0iMTU4NSIgaGVpZ2h0PSIzNDkiPgoJPHRpdGxlPkMxNEI3MkYzQkJCMTQwOUVBNjk3MzFFNDIxNDU0RUUzLXN2ZzwvdGl0bGU+Cgk8c3R5bGU+"
        LogoBase64 = LogoBase64 & "CgkJLnMwIHsgZmlsbDogIzAzMTc5NSB9IAoJCS5zMSB7IGZpbGw6ICNmZjAwMDAgfSAKCTwvc3R5bGU+Cgk8ZyBpZD0iTGF5ZXIiPgoJCTxwYX"
        LogoBase64 = LogoBase64 & "RoIGlkPSJMYXllciIgZmlsbC1ydWxlPSJldmVub2RkIiBjbGFzcz0iczAiIGQ9Im0zMDQuOCAzMDNjLTc3IDUwLjEtMjg3IDgxLjYtMzAzLjkt"
        LogoBase64 = LogoBase64 & "MjguNC03LjgtNTEuNCA1OS44LTE5NC44IDgwLjgtMjMyLjMgMTIuNy0yMi40IDI2LjYtNDQuMSA1My4zLTQyIDM1IDIuNyA2MiAyMS4zIDg5Lj"
        LogoBase64 = LogoBase64 & "MgNDQgMTA4LjIgOTAuNSAxNDAuMiAyMTkuOSA4MC41IDI1OC43em0tMzAuNS0xNTYuMmMtMTQuOS0zMC4xLTM1LjMtNTIuNy01OC4yLTc2LjUt"
        LogoBase64 = LogoBase64 & "MTYuNi0xNy4yLTM2LjktMzEuNS01OS4yLTM5LjgtMTkuMy03LjMtMzkuOCA1LTQ3LjQgMjAuOS0yMi43IDQ3LjEtMzguNCA5My42LTU3LjcgMT"
        LogoBase64 = LogoBase64 & "M5LjYtMTQuOSAzNS40LTM3LjUgOTkuOSAxNy44IDEyNS4zIDIzLjggMTAuOSA0Ny44IDExLjcgNzEuNyA4LjQgMjguOC0zLjggNTcuMi0xMC4x"
        LogoBase64 = LogoBase64 & "IDg1LjItMTggMjkuNS04LjMgNTkuMy0yNS40IDY2LjItNTguOXExLjMtNi43IDItMTMuNmMyLjYtMjguMy04LjQtNjMuMS0yMC40LTg3LjR6bS"
        LogoBase64 = LogoBase64 & "04LjcgODYuNWMtNC42IDE4LjEtMTMuMSAzMy45LTQyLjYgNDcuMS0xMi41IDUuNi04NS4zIDIzLTExMi4xIDIxLTE4LjQtMS4zLTM4LjYtMTIu"
        LogoBase64 = LogoBase64 & "Mi00My40LTMyLjktNi45LTMwLjIgMy43LTYyLjUgMTUuMy04OC45IDExLjktMjcuMSAyMy4xLTU0LjUgMzUuNC04MS4zIDctMTUuMyAyNC43LT"
        LogoBase64 = LogoBase64 & "UzIDUyLjgtMzcuMiA0Mi4yIDIzLjggNzkuMSAxMDQuOSA4Ni4xIDEyMi43IDYgMTUuMSAxMi40IDM0LjUgOC41IDQ5LjV6bS02OS4yLTExMS4z"
        LogoBase64 = LogoBase64 & "Yy0xMi4xLTE3LjEtMzEtNDQuNy00OC0yOC42LTYuOSA2LjQtNDMuNCA5Mi45LTUxLjMgMTE1LjctNS45IDE2LjgtMTcuMyA1Ni44IDUuNiA2My"
        LogoBase64 = LogoBase64 & "44IDI3LjIgOC4zIDcwLjQtMS40IDg2LjQtNi43IDIzLjgtOCA5MC44LTM2LjkgNy4zLTE0NC4yem0yOTguNyAxMzEuM2gtMjcuOGwtOC4zLTI1"
        LogoBase64 = LogoBase64 & "LjJoLTAuNGwtMTUuOCA1LjZoLTI0LjRsLTYuMyAxOS42aC0yNi42bDM4LjQtMTE0LjVoMzIuN3ptLTY4LjctNDQuN2gyNi44bC0xMy4zLTQxLj"
        LogoBase64 = LogoBase64 & "loLTAuMnptMTYxIDQ0LjdoLTI3Ljl2LTQ3LjljMC04LjMtNS4yLTEzLjItMTMuMS0xMy4yLTEwLjUgMC0xNC44IDcuNC0xNC44IDE5Ljh2NDEu"
        LogoBase64 = LogoBase64 & "M2gtMjUuMXYtODMuOGgyNS4xdjExLjJoMC4yYzUuOS03LjUgMTQuNi0xMS42IDI1LjEtMTEuNiAyMC4xIDAgMzAuNSAxMS42IDMwLjUgMzIuOC"
        LogoBase64 = LogoBase64 & "AwIDAgMCA1MS40IDAgNTEuNHptNzguMi04My44aDI1LjJ2NzUuOWMwIDI0LjktMTYuNSA0My4yLTQ0IDQzLjItMTcuMyAwLTI5LjEtNS0zNy43"
        LogoBase64 = LogoBase64 & "LTExLjZsMTMuMy0xNy41YzQuNCAyLjkgMTEuOCA2LjcgMjIuOCA2LjcgMTMuNyAwIDIwLjQtOC41IDIwLjQtMjAuOHYtMy4zaC0wLjFjLTQuNi"
        LogoBase64 = LogoBase64 & "A3LjUtMTMuMiAxMS42LTI0LjIgMTEuNi0yMy42IDAtMzcuNy0xOC4zLTM3LjctNDIuMyAwLTI0LjEgMTQuMS00Mi40IDM3LjctNDIuNCAxMSAw"
        LogoBase64 = LogoBase64 & "IDE5LjYgNC4yIDI0LjIgMTEuNmgwLjFjMCAwIDAtMTEuMSAwLTExLjF6bS0wLjggNDEuNWMwLTEwLjUtOC41LTE5LTE5LTE5LTEwLjUgMC0xOC"
        LogoBase64 = LogoBase64 & "45IDguNS0xOC45IDE5IDAgMTAuNSA4LjQgMTkgMTguOSAxOSAxMC41IDAgMTktOC41IDE5LTE5em00NS41IDQyLjN2LTExNC41aDI1LjFjMCAw"
        LogoBase64 = LogoBase64 & "IDAgMTE0LjUgMCAxMTQuNXptODIuNCAxLjRjLTI0LjYgMC00NC41LTE5LjItNDQuNS00Mi45IDAtMjMuNyAxOS45LTQyLjkgNDQuNS00Mi45ID"
        LogoBase64 = LogoBase64 & "I0LjcgMCA0NC42IDE5LjIgNDQuNiA0Mi45IDAgMjMuNy0xOS45IDQyLjktNDQuNiA0Mi45em0xNy40LTQyLjljMC0xMS03LjgtMTkuOC0xNy40"
        LogoBase64 = LogoBase64 & "LTE5LjgtOS42IDAtMTcuMyA4LjgtMTcuMyAxOS44IDAgMTAuOSA3LjcgMTkuOCAxNy4zIDE5LjggOS42IDAgMTcuNC04LjkgMTcuNC0xOS44em"
        LogoBase64 = LogoBase64 & "0xNDIuMyA0MS41aC0yNy44bC04LjMtMjUuMmgtMC40bC0xNS45IDUuNmgtMjQuM2wtNi4zIDE5LjZoLTI2LjZsMzguNC0xMTQuNWgzMi43em0t"
        LogoBase64 = LogoBase64 & "NjguNy00NC43aDI2LjdsLTEzLjMtNDEuOWgtMC4xem0yMDkuMy01LjZ2NTAuM2gtMjUuMXYtNDcuOWMwLTguMy01LjEtMTMuMi0xMC45LTEzLj"
        LogoBase64 = LogoBase64 & "ItMTAuMSAwLTE0LjMgNy40LTE0LjMgMTkuOHY0MS4zaC0yNy45di00Ny45YzAtOC4zLTUtMTMuMi0xMC45LTEzLjItMTAgMC0xNC4yIDcuNC0x"
        LogoBase64 = LogoBase64 & "NC4yIDE5Ljh2NDEuM2gtMjUuMXYtODMuOGgyNS4xdjExLjFoMS40YzQuOC03LjYgMTMtMTEuNyAyMi45LTExLjcgMTEuOCAwIDI0LjUgMTQuNS"
        LogoBase64 = LogoBase64 & "AyNC41IDE0LjVoMC4zYzUtNy40IDEyLjQtMTQgMjUuNi0xNCAxOS44IDAgMjguOSAxMS42IDI4LjkgMzIuOXptNDEuOSAxNi43YzIuNSA3Ljkg"
        LogoBase64 = LogoBase64 & "OC43IDEyLjQgMTcuOCAxMi40IDguMyAwIDEzLjUtMy44IDE2LjUtOC4zbDE5LjggMTMuMmMtNy40IDkuMS0xOSAxNi42LTM3LjEgMTYuNi0yNy"
        LogoBase64 = LogoBase64 & "4zIDAtNDUuNC0xOC4yLTQ1LjQtNDMgMC0yNC43IDE4LjEtNDIuOSA0NC42LTQyLjkgMjQuNyAwIDQyLjEgMTYuNSA0Mi4xIDQ0LjIgMCAyLjUg"
        LogoBase64 = LogoBase64 & "MCA1LjItMC40IDcuOHptMzEuOS0xNi43Yy0wLjgtOC4zLTUtMTQuMS0xNS43LTE0LjEtOC42IDAtMTQuMyA1LjEtMTYuNSAxNC4xem05MS03Lj"
        LogoBase64 = LogoBase64 & "RjLTIyLjYgMC0yNS4xIDUuMy0yNS4xIDIxdjM2LjZoLTI1LjJ2LTgzLjhoMjUuMnYxNGgwLjFjNC45LTEwLjUgMTEuNi0xNS43IDI1LTE1Ljcg"
        LogoBase64 = LogoBase64 & "MCAwIDAgMjcuOSAwIDI3Ljl6bTIyLjktMzMuNGMtOC4yIDAtMTQuOC02LjYtMTQuOC0xNC44IDAtOC4zIDYuNi0xNC45IDE0LjgtMTQuOSA4Lj"
        LogoBase64 = LogoBase64 & "IgMCAxNC45IDYuNiAxNC45IDE0LjkgMCA4LjItNi43IDE0LjgtMTQuOSAxNC44em0tMTEuNyA3LjJoMjUuMXY4My44aC0yNS4xem05OS4yIDUy"
        LogoBase64 = LogoBase64 & "LjJsMjEuNSAxMi42Yy04LjEgMTIuNy0yMS41IDIwLjUtMzguOCAyMC41LTI1LjYgMC00My44LTE4LjItNDMuOC00MyAwLTI0LjcgMTguMi00Mi"
        LogoBase64 = LogoBase64 & "45IDQ0LjYtNDIuOSAxNi41IDAgMjkuOSA3LjggMzggMjAuNWwtMjEuNSAxMi41Yy00LjEtNi45LTkuNi05LjktMTYuNS05LjktMTEuMSAwLTE3"
        LogoBase64 = LogoBase64 & "LjMgOC40LTE3LjMgMTkuOCAwIDExLjQgNi4zIDE5LjggMTcuMyAxOS44IDYuOSAwIDEyLjQtMi45IDE2LjUtOS45em05My41LTQxdi0xMS4yaD"
        LogoBase64 = LogoBase64 & "I1LjF2ODMuOGgtMjUuMXYtMTEuMmgtMC4yYy0zLjggNi0xMC41IDkuOC0xOSAxMS0yMy42IDMuMi00My45LTE1LjktNDQuNy0zOS44LTAuOC0y"
        LogoBase64 = LogoBase64 & "NS4yIDEzLjktNDMuOSAzOC44LTQzLjkgMTEuMyAwIDIwLjIgNCAyNC45IDExLjN6bS0wLjIgMzEuMWMwLTEwLjktOC41LTE5LjgtMTktMTkuOC"
        LogoBase64 = LogoBase64 & "0xMC41IDAtMTguOSA4LjktMTguOSAxOS44IDAgMTEgOC40IDE5LjggMTguOSAxOS44IDEwLjUgMCAxOS04LjggMTktMTkuOHptMTIzIDQxLjVo"
        LogoBase64 = LogoBase64 & "LTI3Ljl2LTQ3LjljMC04LjItNC43LTEzLjItMTEuOC0xMy4yLTkuNCAwLTEzLjMgNy40LTEzLjMgMTkuOHY0MS4zaC0yNS4xdi04My44aDI1Lj"
        LogoBase64 = LogoBase64 & "F2MTEuMmgwLjJjNS42LTcuNCAxMy45LTExLjYgMjMuOC0xMS42IDE5LjEgMCAyOS4xIDExLjYgMjkuMSAzMi45IDAgMCAwIDUxLjMtMC4xIDUx"
        LogoBase64 = LogoBase64 & "LjN6Ii8+CgkJPHBhdGggaWQ9IkxheWVyIiBmaWxsLXJ1bGU9ImV2ZW5vZGQiIGNsYXNzPSJzMSIgZD0ibTIwNS41IDIyMC4yYy0xNC43IDkuNC"
        LogoBase64 = LogoBase64 & "02Mi4zIDQwLjUtNzkuMSAzMC03LjMtNC42LTQuMi0yMS44LTIuMi0yOC43IDctMjQuOSAxNC40LTUwLjcgMjMuOS03NS42IDUuMi0xNC4zIDIy"
        LogoBase64 = LogoBase64 & "LjMtOCAyOC4zLTIuNyAxNS41IDEzLjkgMjcuNCAzNC4xIDM0LjUgNDkuNCA0IDguOCAzIDIyLjMtNS40IDI3LjZ6bS0xNC43LTI5Yy0xLjEtMy"
        LogoBase64 = LogoBase64 & "4yLTE5LjEtMzQuMi0yNy4zLTI3LjUtMTAuMyA4LjQtMjcuOSA2My4yLTE1LjMgNjMuNyAxMi4yIDAuNSAyOS43LTExLjEgNDAuMy0yMC41IDMu"
        LogoBase64 = LogoBase64 & "OC0zLjQgNC0xMC4zIDIuMy0xNS43eiIvPgoJPC9nPgo8L3N2Zz4="
    ElseIf Cliente = "Vale" Then
        LogoBase64 = LogoBase64 & "PHN2ZyBoZWlnaHQ9IjEwMjMiIHZpZXdCb3g9IjIuMzkgLTkuMjQyIDEwNC43MDUgNTIuMzg0IiB3aWR0aD0iMjUwMCIgeG1sbnM9Imh0dHA6Ly"
        LogoBase64 = LogoBase64 & "93d3cudzMub3JnLzIwMDAvc3ZnIj48ZyBmaWxsLXJ1bGU9ImV2ZW5vZGQiPjxwYXRoIGQ9Im01MC45NyAxNC42MmMtNy4xNTggNS4zNzMtMTMu"
        LogoBase64 = LogoBase64 & "NzgyIDQuNzk0LTE5Ljg3Ni0xLjczNSAxMi4zNTMtNC43MzUgMTguODYyLTEyLjA1MSAyNS4yMDMtNS43NzhsLTQuNzAzIDYuNzI4aC0uMDA4di"
        LogoBase64 = LogoBase64 & "4wMWgtLjAwOHYuMDAxaC0uMDAxdi4wMWgtLjAxdi4wMDNoLS4wMDh2LjAwMWMtLjAwNSAwLS4wMDMuMDA1LS4wMDYuMDA1di4wMDNoLS4wMDF2"
        LogoBase64 = LogoBase64 & "LjAwM2gtLjAwMmwtLjAwMS4wMDZoLS4wMDJsLS4wMDIuMDA0aC0uMDAxYzAgLjAwMy0uMDAzLjAwMy0uMDAzLjAwNWgtLjAwMnYuMDA1Yy0uMD"
        LogoBase64 = LogoBase64 & "A5IDAtLjAxNS4wMTMtLjAyLjAxM2gtLjAwMWwtLjAwMS4wMDVoLjAwMnYuMDAxaC0uMDAybC0uMDA0LjAwNy0uMDA5LjAwNnYuMDAzaC0uMDAy"
        LogoBase64 = LogoBase64 & "YzAgLjAwNi0uMDEyLjAyMS0uMDE3LjAyNHYuMDAyaC0uMDAydi4wMTVoLS4wMDN2LjAwNGgtLjAwMnYuMDA3bC0uMDA1LjAwMmgtLjAwMmwtLj"
        LogoBase64 = LogoBase64 & "AwNS4wMDR2LjAwMmMtLjAwMy4wMDEtLjAwNS4wMDMtLjAwNS4wMDhsLS4wMDMuMDA0LS4wMTIuMDA1di4wMDJsLS4wMDMuMDAyYS43MjIuNzIy"
        LogoBase64 = LogoBase64 & "IDAgMCAwIC0uMDYzLjA4bC0uMDAyLjAwNGgtLjAwMmMwIC4wMDYtLjAwNi4wMTMtLjAwNi4wMjJoLS4wMDF2LjAwM2MtLjAwOS4wMDQtLjAyNi"
        LogoBase64 = LogoBase64 & "4wMjQtLjAzNi4wMjR2LjAwM2MtLjAxOC4wMDktLjAzMy4wMjUtLjA1My4wMzN2LjAwM2gtLjAwN3YuMDAzaC0uMDAybC0uMDAzLjAwNmMtLjAw"
        LogoBase64 = LogoBase64 & "NCAwLS4wMDUuMDAzLS4wMDkuMDAzdi4wMDJsLS4wMTUuMDF2LjAwNWgtLjAwM3YuMDAyYS4xNjMuMTYzIDAgMCAxIC0uMDMuMDJsLS4wMTIuMD"
        LogoBase64 = LogoBase64 & "E1aC0uMDAzdi4wMDNsLS4wMDQuMDA4di4wMDNoLS4wMDZjMCAuMTIyLS4xNzIuMTY3LS4xNzIuMjQ2aC0uMDAzYzAgLjAwOS0uMDIuMDI3LS4w"
        LogoBase64 = LogoBase64 & "Mi4wMzloLS4wMDNjLS4wMDMuMDA0LS4wMTMuMDE0LS4wMTMuMDE4aC0uMDAzYy0uMDA4LjAwOS0uMDAzLjA0Ny0uMDAzLjA1OCIgZmlsbD0iI2"
        LogoBase64 = LogoBase64 & "VjYjgzMyIvPjxwYXRoIGQ9Im01MS42MDQgMTMuODIyYy0xOC42NjQgMTYuODkyLTI1Ljk4Mi0yMy4wNjQtNDkuMjE0LTcuOTQybDI2LjE3OCAz"
        LogoBase64 = LogoBase64 & "Ny4yNjIiIGZpbGw9IiMwMDkzOWEiLz48ZyBjbGlwLXJ1bGU9ImV2ZW5vZGQiPjxwYXRoIGQ9Im01NS45MDQgMjQuMzUgNi4zOSAxMy44NTQgMS"
        LogoBase64 = LogoBase64 & "40NjgtLjAxMSA1LjYxNC0xMy45My0yLjUwNC0uMDg1LTMuOCAxMC4yOTQtNC41NzgtMTAuMjA4bTExLjY1OSAxMy44NDIgMi43NjMuMDg3IDEu"
        LogoBase64 = LogoBase64 & "MTIzLTMuNDZoNS44NzRsMS4xMiAzLjM3MyAyLjY3Ny4wODctNS42MTMtMTMuOTNoLTEuOTg2IiBmaWxsPSIjNzc3ODdiIi8+PHBhdGggZD0ibT"
        LogoBase64 = LogoBase64 & "c1LjIwNSAzMi41N2gzLjU0MWwtMS41NTUtNC44ODgiIGZpbGw9IiNmZmYiLz48L2c+PC9nPjxwYXRoIGQ9Im04Ny41OTkgMzguMDJ2LTEzLjIx"
        LogoBase64 = LogoBase64 & "NWgyLjQ0NXYxMC45N2g2LjA4djIuMjQ1bTMuMDU3LS4wODZ2LTEzLjQxaDcuOTE0djEuMDM2Yy0uNTc3LjM2LTEuMTI0LjgxNi0xLjY1OSAxLj"
        LogoBase64 = LogoBase64 & "IzM2gtNC4xNTd2Mi45NzJoNS4yMTZ2Mi4yNmgtNS4yMTZ2My42NDhoNS44MDN2Mi4yNjEiIGZpbGw9IiM3Nzc4N2IiLz48L3N2Zz4="
    ElseIf Cliente = "CBMM" Then
        LogoBase64 = LogoBase64 & "/9j/4AAQSkZJRgABAQEA3ADcAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoM"
        LogoBase64 = LogoBase64 & "DAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGB YUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQs"
        LogoBase64 = LogoBase64 & "NFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAA"
        LogoBase64 = LogoBase64 & "RCAB7 APUDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtR"
        LogoBase64 = LogoBase64 & "AAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhBy JxFDKBkaEII0KxwRVS0fAkM2Jyg"
        LogoBase64 = LogoBase64 & "gkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDh"
        LogoBase64 = LogoBase64 & "IWGh4iJipKT lJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi"
        LogoBase64 = LogoBase64 & "4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAA AAAAECAwQFBgcICQoL/8Q"
        LogoBase64 = LogoBase64 & "AtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnL"
        LogoBase64 = LogoBase64 & "RChYkNOEl8RcYGRom JygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eH"
        LogoBase64 = LogoBase64 & "l6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExc bHyMnK0tPU1dbX2"
        LogoBase64 = LogoBase64 & "Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9U6KKKACiiigAooooAKKKKACii"
        LogoBase64 = LogoBase64 & "igAooooAKKKKACiiigAoooo AKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoooo"
        LogoBase64 = LogoBase64 & "AKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAopKWgAooooAKKKKACi iigAooooA"
        LogoBase64 = LogoBase64 & "KKKKACiiigCC4njtYZZpnWOGNS7yOcBVAySTX59ar/wAFqvhLY6pd29r4R8WajawzPHFeRx2"
        LogoBase64 = LogoBase64 & "6LOgJAcK8oYAjnDAEZ5Fejf8ABUz4 8n4O/syaro9miyav4y36FEGAYRwSI32hyNwP+r3KGG"
        LogoBase64 = LogoBase64 & "cMy5r8HT1oA/er4U/8FHvCnxj+FvxL8ceH/BHime38CQWtzfad5cDXNxHM0mXiCyEYjW J3b"
        LogoBase64 = LogoBase64 & "cRhVJ5ryf8A4fZ/C3/oRPGH5Wv/AMerz3/giPp1rrGh/HKwvreO7srpdJgnt5lDJLGy3ysrA"
        LogoBase64 = LogoBase64 & "9QQSCPevgj9qr4K3X7P3x58WeC7hZDb2d00 llNJGEM9s53RSBQSACD09qAP05/4fZ/C3/oR"
        LogoBase64 = LogoBase64 & "PF/5Wv8A8er6o/ZS/ao8P/taeCdS8TeHdI1LR7SxvTYvDqnl+YzBFbI2Mwxhu9fzi1+z3/BF"
        LogoBase64 = LogoBase64 & "n/ k3fxX/ANjC3/oiOgD9CW+7Xh/x5/bM+E37OLRweMvFEMOqSxtJFpdijXN0wG4AlEztBZS"
        LogoBase64 = LogoBase64 & "uWwMivFv+Clf7ak37NngOHwt4TvDB8QfEERaC5VA4 sLbOHl54DnkLnODzggV+HOoX9zql7c"
        LogoBase64 = LogoBase64 & "Xt7cy3d5cyNNPcTyF5JZGOWdmJySSSSTySaAP2K/4fY/CzPHgXxgfQ4tf/AI9X0v4C/bD8N/"
        LogoBase64 = LogoBase64 & "ED9l/Wfj hZ6LqttoGl295cyadP5X2p1tiQ4XDFcnbx81fzrV+uv7NP/KHXx5/2C9e/9CkoA"
        LogoBase64 = LogoBase64 & "3f+H2nwt/6EXxh+Vr/8eo/4fafC3/oRfGH5Wv8A8er8baKA P6HP2cP28fhR+01NBp3h7WH0"
        LogoBase64 = LogoBase64 & "zxLIhc6BqyiG5OBkhOSsmBydhNfRa9K/lm0PXL7w3rFlqumXU1jqNlMs9vcW7lHjkU5VlYEE"
        LogoBase64 = LogoBase64 & "EEdq/ov/AGPP2g I/2lv2f/DPjV/ITV5o2tdWtrdgRDeRMUkBA+5vAWQKeQsi0Ael+OPHnhz"
        LogoBase64 = LogoBase64 & "4a+HbjX/FWt2Ph7Rrcqst9qEyxRIWIVRuPckgAe9fEPiz/gs18HNB 12ey0rQvE3iSyjC7NS"
        LogoBase64 = LogoBase64 & "tIIoY5CRkgLK6vweORzjivlH/grp+0bqHj/wCNA+Gdjf7vC/hTy5J7XyChbUirB2Zj94LG4C"
        LogoBase64 = LogoBase64 & "4wPnbr2+A6AP6Hf2Y/26 /hf+1O7af4bv59M8TRwmeXQNWQRXAQOy5QglZeAGOwnaHXOK9P+"
        LogoBase64 = LogoBase64 & "L3xu8D/AAI8NNr3jrxFZ6BYEOIvtD/vLhlUsUijHzSPgHCqCTX803hnxNqv gzxBYa5oeoT6"
        LogoBase64 = LogoBase64 & "Vq9hMs9teWzlJInHQg/5yMg16V+0n+094z/ak8WafrvjCeHdp9otna2tsCsUagDe+Cfvuw3M"
        LogoBase64 = LogoBase64 & "RgdOBigD9O9Q/wCC1Xwmt764ht /Bvi29gjkZY7hEtkWVQeGCtKCAeuCM19W/s5/tXfDr9qL"
        LogoBase64 = LogoBase64 & "Q7m/8Eax9ourMR/btLukMN1as6BvmQ9VySu9cqSjYJxX83Vdv8GPi94h+BfxI0Xxl 4ZvZrT"
        LogoBase64 = LogoBase64 & "UNPnV3SGTYLmHcDJA/ByjgYOQex6gUAf04GvB/2lv20Phl+yzbwx+L9Ukm1u5gae10TTo/Nu"
        LogoBase64 = LogoBase64 & "plHQkdIwTkBnIBweeK3fGH7QWlaD+zLq vxf0+Wx1Gxg8PSazax/adsNzKISyQCTH8UgEfTO"
        LogoBase64 = LogoBase64 & "T07V/Oj448c698SvFepeJfE+q3Gta5qMpmub26bc7sf0AAwAoAAAAAAGKAP1/0b/gtN8I9Q "
        LogoBase64 = LogoBase64 & "1S1t7zwn4s0u1lcJJeSxW8iQqf4iqSFiB6AE19ufDP4q+EfjJ4Xi8ReDNes/EGkSNs+0Wcgb"
        LogoBase64 = LogoBase64 & "Y4AJRx1VhkZU4NfzDV7N+zf+1X40/ZhuPFMnhO 42jX9Km0+RXkYC3lZT5N1GvK+ZG2CNynI"
        LogoBase64 = LogoBase64 & "3LxnNAH7QftLf8FD/AIUfsy61LoGrXV14h8Uw+W02i6MiySQq2eZHYhEYAA7CwbDA45ry7wD"
        LogoBase64 = LogoBase64 & "/AMFi vgz4w8SW+l6tpfiDwjbT4UalqcMTwK5YAB/KdmUck7iMAA5NfilqOo3Wr6hc319czX"
        LogoBase64 = LogoBase64 & "t7cyNNPc3EhkklkY5Z2Y8sxJJJPJJqvQB/Uj4T8WaL44 8PWWueHtUtdZ0e9TzLe+spRJFKv"
        LogoBase64 = LogoBase64 & "qrDrWvX5Af8Ebv2idQ0j4g6j8IdSvJJdF1aCa/0q38st5N1Gu+XDZwqtGrk8csq1+vq0AOoo"
        LogoBase64 = LogoBase64 & "ooAKZ3p9cL8 bfFFv4N+EnizVbrxJa+D0jsJIotevSRFYzSDyopWwD0kdOMcnFAH4n/8FQfj"
        LogoBase64 = LogoBase64 & "1/wuj9pjUtOs5/N0LwmraRabWR1eQHM0isvVWYDGT/DXyCetfW V1+yH8Ob65luJ/2pvh7LP"
        LogoBase64 = LogoBase64 & "M5kkdornLMTkk8etereAf+CQuq/FLwraeJfCfxh8K67oN2XWC/tLS4aOQo5R8H2ZWH4UAeof"
        LogoBase64 = LogoBase64 & "8ENf+PP4z/wDXTRv5 Xtb/APwWZ+AJ1vwToPxZ0+ONZtEddN1TARC8MrgROT95yJCqgc4DE1"
        LogoBase64 = LogoBase64 & "6r+xv+y+n/AATt8H/EnxF8QvHOjzaFqh055NQjikhitBE00Y3lv7zXKA Yrtvid+1f+zF8Wv"
        LogoBase64 = LogoBase64 & "h/r3g/xB8UvDtxpGs2klpOouDvUMpG9CVO1geQ2OCAaAP5/TX7Pf8EWf+Td/Ff/AGMLf+iI6"
        LogoBase64 = LogoBase64 & "+Dv+GOvhp/0dH8PP+/Vz/hX 39+x74HtfgD+xF8Yb/wd8RNJ8bvBb6rqFrrego6x29zFYblX"
        LogoBase64 = LogoBase64 & "5x95SEb05FAH5X/tXfGK6+O37QHjLxhcNceRdXzxWcN1t3wWyErFEdvHyqMf1r hvh98PfEP"
        LogoBase64 = LogoBase64 & "xU8X6d4Y8LaVcaxrd/II4LW3XJPqxP8KjqWPAFYM00lxNJNKxeWRi7MepJOSa/S3/AIIi+HN"
        LogoBase64 = LogoBase64 & "MvviD8TdauLSOXVdO0+ygtbps7oUm ebzVHbDeVHn/AHRQBlWf/BEz4jT2cEk/j7w1bTPGrS"
        LogoBase64 = LogoBase64 & "QmGdjGxAJXIXBweMjrivq2T9n7VP2Y/wDgmf8AEvwFrGqWmsX1poesXDXVirrGwlDuAA wBy"
        LogoBase64 = LogoBase64 & "Aea+368V/bW/wCTSfi9/wBixff+iWoA/nBr9L/Af7Afwi179gN/jBqkmrW3iweFr/VVk/tFU"
        LogoBase64 = LogoBase64 & "tjdRLN5S7CvRmRF25yc4HJr80K+xP2r9Yvr X9jr9lSwhvbiGxutF1Vri1jlZYpit3EVLqDh"
        LogoBase64 = LogoBase64 & "iDyM9KAPjuv2k/4Ip/8AJsPiz/scbj/0isq/Fziv3p/4Jb+OfAvij9lPQdI8GRSWl3oMj2ut"
        LogoBase64 = LogoBase64 & "W1 3s+0G9c+a8xK/ejff8jEcKuzJKGgD8kv2+Ls337YnxVmK7CdYZdoOekaD+lO/YX+AOift"
        LogoBase64 = LogoBase64 & "KftFaL4L8R3V1a6LJbz3lx9jIWSVYk3eWGP3c/wB7 nFVv26f+Tu/in3/4nMn/AKCtes/8Ej"
        LogoBase64 = LogoBase64 & "f+TztG/wCwTf8A/oqgDkP+CiH7M/hr9lf472fhjwldXs+iajosGrRRX7h5LdmlmiZN4xuGYN"
        LogoBase64 = LogoBase64 & "2SB9/HbN fL4r7+/wCC1X/J0vhf/sTrX/0tva+AaAP0U/b1/wCCdngz9m34AaF4z8IahfXF7"
        LogoBase64 = LogoBase64 & "bXUNpq0mpT7jdGRMK8aBcJ8wYkZ6Gvzrr9yv+CuH/Jls/8A 2GNP/m1fhrQB+ufxa/5Qq6H/"
        LogoBase64 = LogoBase64 & "ANgzSv8A05RV+Rlfrn8Wv+UKuh/9gzSf/TlFX5GUAfoF+yj+wn4E+Mf7F/jf4m+JG1NfE9h/"
        LogoBase64 = LogoBase64 & "acmmtZ3gSAJb2q SRh02nJMm/PPQivz9r9mf+Cf8A/wAozfHH/Xtrv/pLX4zUAfY//BNf9kX"
        LogoBase64 = LogoBase64 & "wh+1d428XweNLm/XS9CsoZVtdPkETTPK7qCXwcBdnTHOfavE/2svg /p3wE/aI8beA9Iu573"
        LogoBase64 = LogoBase64 & "TNHukW2muQBJ5ckMcoVsdSvmbc99ueM19u/wDBD3/kdPix/wBg+w/9GTV8wf8ABSn/AJPe+K"
        LogoBase64 = LogoBase64 & "P/AF9Wv/pFb0AUv+CePi a98Lftj/DO4sWVZLrUhYSblyDFOpikH12ua/obWv5z/wBhT/k77"
        LogoBase64 = LogoBase64 & "4Uf9h+2/wDQxX9GNABRRRQAV8v/APBTL/kx74of9cLL/wBL7avqCvl//gpl /wAmPfFD/rhZ"
        LogoBase64 = LogoBase64 & "f+l9tQB/PvX6t/8ABF349PcWXij4S6neqRbn+19Hikdy5VuLiNB91UUhX7EmVutflJXof7Pv"
        LogoBase64 = LogoBase64 & "xauvgb8ZPCfji1TzjpF8k0sLFg skR4dTtIJ+UnjPUCgD9wv+Cm//ACYz8T/+uVh/6cLav5+"
        LogoBase64 = LogoBase64 & "q/fP/AIKJeILPxb/wT28c65p7M1hqen6Xe25kXaxjkvbR1yOxww4r8DKALGn2Nxql /bWVpE"
        LogoBase64 = LogoBase64 & "091cSLDFEvV3YgKo9ySBX9Ev7Lv7NmkfCL9l3SPhtqNgkg1HTpP7fjZNjXE1zHidX2k8hSI8"
        LogoBase64 = LogoBase64 & "g9EFfkr/wS3+AMXxq/aTtNS1Sya78PeE 4hqlzlUaIz7sW8bq3JDEO3A6x9q/eLn0oA/l28c"
        LogoBase64 = LogoBase64 & "eHLzwf4y1vQ7+yl0680+9ltpbSdSrxFXI2kHuK+gv8Agn3+1Mn7LvxygvtWnkj8Ha3GtjrK "
        LogoBase64 = LogoBase64 & "pkhFyTHPtH3ijE/g7V7b/wAFbv2Wb3wH8TH+LWjWkZ8M+JJFS/W3ibNtfY+Z5DyMSEZzx83F"
        LogoBase64 = LogoBase64 & "fnpQB/TD4c/aK+F3izRbTVtL+IXhu4sLpS8Mj6 nFEWAJHKOwYcg9QK4z9rnxBpfif9jn4uX"
        LogoBase64 = LogoBase64 & "+j6lZ6tYt4a1BVurGdJoiREwIDKSMg1/OjX66/s0/8odfHn/YL17/ANCkoA/Iqv3C/Zn/AGW"
        LogoBase64 = LogoBase64 & "fhv8A tKfsS/BuPx7oR1WTT9IuIrO4jneKSASXDl9pU9TsX8q/D2vvH46fG7x78J/2O/2YbH"
        LogoBase64 = LogoBase64 & "wd4t1Tw3aavoOrJfxabcGIXAW5QDdjrgO4H+8aAPjz4u +Crb4b/FLxZ4VstTh1m00fU7ixh"
        LogoBase64 = LogoBase64 & "1CAgpcJHIVDjBIwQPU19nf8EY/Euoaf+0xruiRXssWl6l4emmuLRfuTSwzReUzD1USy4P8At"
        LogoBase64 = LogoBase64 & "n1r4EZi7FmO 4k5JPev0M/4Is+B9T1L4+eLPFkIiGk6RoZsblmYhzLcyo0YUY5GLeTPPHy+t"
        LogoBase64 = LogoBase64 & "AHzR+3V/yd58U/8AsMyf+grXrP8AwSN/5PO0b/sE3/8A6KryX9 un/k7v4p/9hmT/ANBWvWv"
        LogoBase64 = LogoBase64 & "+CRv/ACedo3/YJv8A/wBFUAdf/wAFqv8Ak6Xwv/2J1r/6W3tfANff3/Ban/k6Twv/ANida/8"
        LogoBase64 = LogoBase64 & "Apbe18A0AfuV/wVw/ 5Mtn/wCwxp/82r8Na/cr/grh/wAmWz/9hjT/AObV+GtAH65/Fr/lCr"
        LogoBase64 = LogoBase64 & "of/YM0n/05RV+Rlfrn8Wv+UKuh/wDYM0n/ANOUVfkZQB+zP/BP/wD5Rm +OP+vbXf8A0lr8Z"
        LogoBase64 = LogoBase64 & "q/Zn/gn/wD8ozfHH/Xtrv8A6S1+M1AH6cf8EPf+R0+LH/YPsP8A0ZNXzB/wUp/5Pe+KP/X1a"
        LogoBase64 = LogoBase64 & "/8ApFb19P8A/BD3/kdPix/2 D7D/ANGTV8wf8FKf+T3vij/19Wv/AKRW9AHO/sKf8nffCj/s"
        LogoBase64 = LogoBase64 & "P23/AKGK/oxr+c79hP8A5O++FH/Yetv/AEMV/RjQAUUUUAFeA/t3fDzxF8Vv2T /H/hXwppc"
        LogoBase64 = LogoBase64 & "ms+INRitVtbGJ0RpSl5BI2CxCjCox5PavfqKAP56/+Hb/AO0j/wBEs1D/AMDLT/49R/w7g/a"
        LogoBase64 = LogoBase64 & "R/wCiWaj/AOBlp/8AHq/oUooA/Nbw v8H/AI3+IP8Agmd45+D3iTwLqUHjTT5raDRbWSW3xe"
        LogoBase64 = LogoBase64 & "WZvYJ9qyCYgtHsnzu24XYBnoPhT/h3D+0iP+aWaj/4GWn/AMer+hSigD5V/wCCcX7Mep /sy"
        LogoBase64 = LogoBase64 & "/ACKw8SQfZvF2uXb6pqdtuST7KSAkcIdRziNELDJAdnwSK+qqKKAMnxX4V0jxx4dv8AQdf02"
        LogoBase64 = LogoBase64 & "31fR7+JoLmyuow8cqEYIINfkZ+01/wR98Xe GtU1PWvhDcxeJNBklaWHQLqXy762Q7cRpIx2"
        LogoBase64 = LogoBase64 & "zAEtgsVIVRkscmv2JooA/nr/AOHb/wC0j/0SzUP/AAMtP/j1fpH8C/2fviD4V/4Jl+MPhrqv"
        LogoBase64 = LogoBase64 & "hq 4s/HF5p+sQwaO0sRkkeZn8oBgxT5sj+LvX3jRQB/PX/wAO3/2kf+iWah/4GWn/AMer6Q/"
        LogoBase64 = LogoBase64 & "aO/Yv+NPjT9nH9nHw7ovgO8v9a8M6TqVvq9olzbq1 pJLcRtGrFpADlVJ+UnpX7CUUAfgR4L"
        LogoBase64 = LogoBase64 & "/4Je/tD+LPEFvpt54LXwzbyAl9S1a9h8iMAZwfLZ2JPYbevcda/Zb9lf8AZv0P9lv4Q6X4L0"
        LogoBase64 = LogoBase64 & "hkvbqMtc alq3krFJf3LHLSMB2A2ooJJCooJJyT7DRQB+I37W37CXx4+If7SXxB8SeHvh1fa"
        LogoBase64 = LogoBase64 & "loupao89pdpdWyrKhVcMA0oPbuK9G/4Jv8A7G/xl+C37UGm eJvGvga80HQotNvIXvJrm3dQ"
        LogoBase64 = LogoBase64 & "7x4UYSRjyfav1zooA/LH/gqZ+yX8W/jx+0BoGv8AgPwXd+ItIt/DNvYy3UNxBGFmW6unZMPI"
        LogoBase64 = LogoBase64 & "pyFkQ9MfNXxz/w AO3/2kf+iWah/4GWn/AMer+hSigD5R/wCCkHwi8X/Gj9luXwx4L0SbXte"
        LogoBase64 = LogoBase64 & "bU7KcWUMiI2xC25suyjjPrX5M/wDDt/8AaR/6JZqH/gZaf/Hq/oUo oA+DPiN+z78Qta/4JY"
        LogoBase64 = LogoBase64 & "6T8L7Lw1cT+PIbDTopNFWWISK0d9HI43FtnCAn73avze/4dv8A7SP/AESzUP8AwMtP/j1f0K"
        LogoBase64 = LogoBase64 & "UUAfFH7HHwN8d/Df8AYP 8AFvgbxJ4duNL8WXkOrrb6ZJLEzyGW32xAMrlfmbjkj3r8wv8Ah"
        LogoBase64 = LogoBase64 & "2/+0j/0SzUP/Ay0/wDj1f0KUUAfm/8A8En/ANmX4nfADxR8RLn4geErnw3B qVlZxWjzzQye"
        LogoBase64 = LogoBase64 & "aySSlgPLdsYDDr614F+3R+xD8cPip+1d8QPFXhX4f3useH9RuLd7W+jubdFlC2sKMQGkB4ZW"
        LogoBase64 = LogoBase64 & "HI7V+ztFAH4j/sk/sJ/Hj4eftK fDrxJ4h+HV9puiabrNvcXd291bMsMauCWIWUkgewNftut"
        LogoBase64 = LogoBase64 & "LRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRR RQAUUU"
        LogoBase64 = LogoBase64 & "UAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFAB"
        LogoBase64 = LogoBase64 & "RRRQAUUUUAFFFFABRRRQAUUUUAFFFFAB RRRQB//9k="
    ElseIf Cliente = "Itaminas" Then
        LogoBase64 = LogoBase64 & "iVBORw0KGgoAAAANSUhEUgAAAEMAAAAxCAYAAACBIBS5AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAACxMAAAsTAQCanB"
        LogoBase64 = LogoBase64 & "gAAAuISURBVGhD7Vp7kBxlEV9IIsgjyJG9+Xr2wgm5QHJ30z27CyGguDyKcLmd7pm73ApKJKW8FOXhq6wgGgl/QAhQKsVLLaISUCwiUJIq3o9C"
        LogoBase64 = LogoBase64 & "kBjRAkTCSx4WQR4hhAASLmSt/mb3cpm75OZS3JFEu6r/2J1v+uvvN/311193ZzIfA+19UDjRIfmuIX4YiN8BkqohfsOQ3NZIcsLexcpeyXd2OC"
        LogoBase64 = LogoBase64 & "qVSmNdn+cY5JfBD6sKQpJdP6oC8XKTl88l399hqFg8dZxB+ZYhXg80OBB9rEAhv+x6Ikk5OwQZX3qAZO3mLCLJ1kKQlxs/ODgpa7umRq+8v0G5"
        LogoBase64 = LogoBase64 & "yc13DVj0ltgtzKqCx+c1N5d2Tcrcbslp48Agv5fWKvrAsNYhy1yUfFLmdksOyddzhZ4Bix2S/bBqUN417WFnUua2Q8XiODV9aA+muJ4cuFnG6A"
        LogoBase64 = LogoBase64 & "DT1jnV+LLAzXcPXGwKVuswyF9ronJugPwEm7ZoKhSDCZlMZuekyiNG2baOSYZkje5pXeRmWfc8RU8b5NvjI3PgYodk3V7EK8DvWu3mtzxfrlip"
        LogoBase64 = LogoBase64 & "uhgtaGjpGJ/UecRIv/gApQfh+p4H4itha8CIt8kzhuRug/L6UEeyOmjH44WjCkajF3q5FGYfm7j8EYjPNiTPDBcQ/doGw8UG+RqDvDoNGIb4J6"
        LogoBase64 = LogoBase64 & "MawTqt0bSc3QIDFdpEuTiafNih8AiHeOHwjlZdeFh1UY4Hj8+Jw/ahwQCUq/b1ynsndR4xMu1ByS2kswzdJtlCNAm8oADEj8UAbXlR9giOHefi"
        LogoBase64 = LogoBase64 & "Zgo/ZfJyriF5f6j33Ly+E17T1NrTkNR5xChH4RHDAcO0Squ+BxTNAOQngORD+8xPglC3JlkHKEtz7Yz2PeSLADWMHzhHcj6Dcp17cLRPUucRI/"
        LogoBase64 = LogoBase64 & "DKh8cnxUCFksoByqM5j6dvfDcoAMrvDMmLhvg91w9rAIQbAGUNED/lYnjpvsj7bXxHHfCWraLffDfUjtfRIfDKxWH4jGfdfMibvF8MdnOQj3KJ"
        LogoBase64 = LogoBase64 & "LwA/uMGQ/AFQrgXkuY3eTK//WCXHl2vT+JtaTHKdO2UULcPkO6fmUigX3z7lNZMPT0vKGA4Z5N+nAkOPVgqv3Gs0HWi2LZoExGuHcmj2OfJ6F2"
        LogoBase64 = LogoBase64 & "VBUkZaUpM3KHekByOYn22t7JGUM2LktoUTgfjxVPtYF+EFS5Iy0pKD5XaDvCxNBGsjUYy+0tpa+URSzoiRfi11gqkUjJ3a33KFjU50OOSQfBlI"
        LogoBase64 = LogoBase64 & "XkkHfHcVvPDwpIwRpabpPZ80yN+2kw+i1KYcVjVGcH25MClnKGqgcg6Ql6YBvbYl39arQlLOiJPb1nnYQIUG59rR+ayDwXFJOZuj5oM7jUNyuc"
        LogoBase64 = LogoBase64 & "Hw3aF9U1+McfuEQhckZY04TcyLC8j3pPtq9e0SPm288PQtKazHLmD0WRd5sUvh2tTy85oVk7PUapMyR56KRU3wnpjGy/cpbKPO6E3w5EaH+BtZ"
        LogoBase64 = LogoBase64 & "CmZk/TJpXqTR4+ng8ReAgosNBis0GEvjJ2K2R/galzr9pJqjRg6Fnwbiv0I+3derK259jR/9B/QmS3y/Qb4TMFwOyKv1fjE8eVKNA8DwEr3HJH"
        LogoBase64 = LogoBase64 & "UcPbLWEXzJfsVBlBySbSjeZY/f2MK2Qk7sK1aadjkoqd6oU2M7Owbl+lyKi9tIsLUkLzz74/EVgxAUykXw5fHh+I+PgvUDGD9cnG3tNEmdPj4q"
        LogoBase64 = LogoBase64 & "lcYCBRFQ+LKafVLpkeDYT0R3GT9OD2xT1Fyas6uhaDaQvJQuGNtaDmMgUO51qDwtk8nslNRlmyALCHI3+NFfbPJnax3rZliP5jjPGdwcF5bmjV"
        LogoBase64 = LogoBase64 & "5JYKuoeOo4h6JpQPxrIFkfH6MDFzYs1gyYlROuMsQX6JGenHabJuN3ZA3xbKDgbvDDDWopLqXIf/ZjmwWzWy7qBeRb3LzIqF7PP1IqFsc5yPsZ"
        LogoBase64 = LogoBase64 & "DGYZlKsN8XPqYHXPx4WfKDb9OtfiDQtcfFSvBF8WmTx3a0UtU6mMSU6xHVJljF77mwqhl/Ojma4ffMf4/DMgXmqQHzLIjwDyMiC52yW5HpDPd4"
        LogoBase64 = LogoBase64 & "k/D4Wo6GBXo76flLjDUHPznF1b/I4sFIN9tYyg1+7c1O7JmgSeUDgW4q2wrTvH/9P/ENnI0edw/PSehqxN2JS74miyP5e7jF/u3HPK0TYF39LS"
        LogoBase64 = LogoBase64 & "sYvrh4cZ5I7m5tImN0Xt3WrSvKVf7jH5YJaDckj/5/qu1k7sc79c0bH6/wRtX/ClomUI9RGaZzV+MDPnlyvqO/rfPfYvVvZy88HRqttgpUQ92n"
        LogoBase64 = LogoBase64 & "XuzddPSmM10QQ+n6/+SfVQvTOaygeUlZpPcChYAhi+DShv2aq3pvlJ3qz991wTRRrxZbJ5bgHkVwD5eeMF3+w/jV6jgWS+NplojsEQ3zbBL0+u"
        LogoBase64 = LogoBase64 & "P7dA2Cu7rDEoeo2fH/8vZxmUXvDkCk3w5DypAIZPqC6A8oJDcmRdRs4vk0F50JC8nuzqgeDU3YDkIdvkgvwDldX/uSaWgML7DckqIP4QUHp1jQ"
        LogoBase64 = LogoBase64 & "7JjZk4MRL2apXMUPB9h2SJQ3KzbTtCWe8gL9eBBmWRO2XmAZp5Nn4t/xkXbR7MZKp9IXFT64wGwPBSW3GL8xKvgien1J8bj+fUq2hxVS68RP8H"
        LogoBase64 = LogoBase64 & "LzwnV6jY+mq2tbSHJo7su/bI1RhFrldriceWi4bCJ4GiqlpBXbZ9huFJxg9X2blR/p3Mqrme/ELLFwqGVuwNyc8N8b8ava7pGRvgoPS67cGhiq"
        LogoBase64 = LogoBase64 & "KW8zURa0j+aQEhOSWDs3ePy/zVnfS8B5KnbKLX53WauTaeVOqTqdlaMOJ0nVrdB1ot02fWalCuMqQWwSvjyxbHYCDPta1OKNc6eMzumicB4lcB"
        LogoBase64 = LogoBase64 & "uVc/iob2OV9OtGPVukiejHXYtBtQk0Uasaru9tLoByf3Lx2Az8/HhW0+J67szd59nwPFtXecOhj966F6tNkmkbgp7Yt9M5VKY43fVbGToPwdKJ"
        LogoBase64 = LogoBase64 & "hf64m4pT6kDkYt6bLMkDxgUB51ceYBuvcB+WntDDYkN2unzZbAMMhvAcl9mmsFsl/zMZOPpjZhV/tgYNjCtloTyT8ckiuA5B3NxPUvRKul65od"
        LogoBase64 = LogoBase64 & "5Begnc9omtyTqz8bGgySE/rGanWLwjtrW+hMLfDYBhTk59Wh6pg+y1CQtMDsy7zYLIPTnLwcr5GlZr2BwvOaDjpui2DYQjTKfergwZd79bnBYF"
        LogoBase64 = LogoBase64 & "6O5Mj4YyTB4N/a7YfBD5vy3GKIX1I9sm3BjHrckiX+jFqzzaaTvK/+wiDfaq1jOGA05oNDc9Yq+DVX9z52H1drMdig4baO2QhGt07200afj7FJ"
        LogoBase64 = LogoBase64 & "H9Rcp1ynjjn2B/yjdGDwQxqYNfrByYZ4lUF+Gyj8nvZ69AejobWjNS5yd6uMM7WYZKNa3RKeLNnnQNkzXsW8nffHrkYrT8uWvnyoVgfEd6YGo2"
        LogoBase64 = LogoBase64 & "Fax3gH5eqaL9gAxOsMWZ+xIc6B8mMm39msk/aBgXKZvZto0rjWiaOnS84vT9bbZzowgj85xG12DMkSm+bTfU/yjlGLq4FhUBbYE4zsF18X+yXV"
        LogoBase64 = LogoBase64 & "Uy99YdWhWVZGH1UqY1o6OnYB5JNqvR+9qcGYWOiYpE0kOomDvBQwWAQoi2pp/zdsmxHyXD2v62DovtV3DfGP4/xGVHU8Waj/AcnCtGBoD5mOca"
        LogoBase64 = LogoBase64 & "h8hCFeUStKKdi94EVFDe+B5NHa1rwDPP6Vsp4cmiyubeVLNcaZOJXbDPItDnYe0uL3ZAHDi/SIVYsdGgx1oJXKGPW+NXN/wt4eledldh4/fUaD"
        LogoBase64 = LogoBase64 & "9nnam6gvDzRRj542F9csw24drYfUfr+oyZ8aGBcPF4xYliy0LU1xJ2CvWg144claZjAer9YajB1YKo21432eG39UXlvzIyvibc0fxJYddwQZCs"
        LogoBase64 = LogoBase64 & "5VJe4F5LvqrUZKtpbq82Jr0th1lCrnoPxSxwGGZ9TH1clt7fQB+c/qwByMAvACVW6Z1mP1uSpsvPA3LgaXaTZd/zMoX1VgjRecrr+ByicY4kcU"
        LogoBase64 = LogoBase64 & "dM2YAQXa8nSrg3y5NuDW51JHpyAD8oNW73xns245baRTQJOdfk10VM4g36S+BzA8ttETMcj3aNcQoLWIJ512XVNp7H8BhDlwnlb1bwQAAAAASU"
        LogoBase64 = LogoBase64 & "VORK5CYII="
    ElseIf Cliente = "Vallourec" Then
        LogoBase64 = LogoBase64 & "PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9Im5vIj8+CjwhLS0gQ3JlYXRlZCB3aXRoIElua3NjYXBlIC"
        LogoBase64 = LogoBase64 & "hodHRwOi8vd3d3Lmlua3NjYXBlLm9yZy8pIC0tPgoKPHN2ZwogICB4bWxuczpkYz0iaHR0cDovL3B1cmwub3JnL2RjL2VsZW1lbnRzLzEuMS8i"
        LogoBase64 = LogoBase64 & "CiAgIHhtbG5zOmNjPSJodHRwOi8vY3JlYXRpdmVjb21tb25zLm9yZy9ucyMiCiAgIHhtbG5zOnJkZj0iaHR0cDovL3d3dy53My5vcmcvMTk5OS"
        LogoBase64 = LogoBase64 & "8wMi8yMi1yZGYtc3ludGF4LW5zIyIKICAgeG1sbnM6c3ZnPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIKICAgeG1sbnM9Imh0dHA6Ly93"
        LogoBase64 = LogoBase64 & "d3cudzMub3JnLzIwMDAvc3ZnIgogICB4bWxuczpzb2RpcG9kaT0iaHR0cDovL3NvZGlwb2RpLnNvdXJjZWZvcmdlLm5ldC9EVEQvc29kaXBvZG"
        LogoBase64 = LogoBase64 & "ktMC5kdGQiCiAgIHhtbG5zOmlua3NjYXBlPSJodHRwOi8vd3d3Lmlua3NjYXBlLm9yZy9uYW1lc3BhY2VzL2lua3NjYXBlIgogICB3aWR0aD0i"
        LogoBase64 = LogoBase64 & "MTUyLjExNTc4IgogICBoZWlnaHQ9IjM1LjkwMTMyOSIKICAgdmlld0JveD0iMCAwIDQwLjI0NzMwMiA5LjQ5ODg5MzQiCiAgIHZlcnNpb249Ij"
        LogoBase64 = LogoBase64 & "EuMSIKICAgaWQ9InN2ZzQ3NTMiCiAgIGlua3NjYXBlOnZlcnNpb249IjAuOTIuMCByMTUyOTkiCiAgIHNvZGlwb2RpOmRvY25hbWU9InZhbGxv"
        LogoBase64 = LogoBase64 & "dXJlYy5zdmciPgogIDxkZWZzCiAgICAgaWQ9ImRlZnM0NzQ3Ij4KICAgIDxjbGlwUGF0aAogICAgICAgaWQ9ImNsaXBQYXRoNDE3MSIKICAgIC"
        LogoBase64 = LogoBase64 & "AgIGNsaXBQYXRoVW5pdHM9InVzZXJTcGFjZU9uVXNlIj4KICAgICAgPHBhdGgKICAgICAgICAgaW5rc2NhcGU6Y29ubmVjdG9yLWN1cnZhdHVy"
        LogoBase64 = LogoBase64 & "ZT0iMCIKICAgICAgICAgaWQ9InBhdGg0MTY5IgogICAgICAgICBkPSJtIDMyNi4zMDksOTcuNDQ0IGggNi4wNzMgdiAtMTAuNzUgaCAtNi4wNz"
        LogoBase64 = LogoBase64 & "MgeiIgLz4KICAgIDwvY2xpcFBhdGg+CiAgICA8Y2xpcFBhdGgKICAgICAgIGlkPSJjbGlwUGF0aDQxNzkiCiAgICAgICBjbGlwUGF0aFVuaXRz"
        LogoBase64 = LogoBase64 & "PSJ1c2VyU3BhY2VPblVzZSI+CiAgICAgIDxwYXRoCiAgICAgICAgIGlua3NjYXBlOmNvbm5lY3Rvci1jdXJ2YXR1cmU9IjAiCiAgICAgICAgIG"
        LogoBase64 = LogoBase64 & "lkPSJwYXRoNDE3NyIKICAgICAgICAgZD0ibSAzMjkuMDgzLDk1LjQ3NiBoIC0wLjA0IHYgMS42ODcgaCAtMi43MzQgViA4Ni42OTQgaCAyLjg3"
        LogoBase64 = LogoBase64 & "NSB2IDQuNzE4IGMgMCwwLjQ3MSAwLjA0OCwwLjkxMSAwLjE0MywxLjMxNSAwLjA5NCwwLjQwNiAwLjI1MywwLjc2MSAwLjQ3NSwxLjA2NSAwLj"
        LogoBase64 = LogoBase64 & "IyNCwwLjMwMiAwLjUxNywwLjU0MiAwLjg4LDAuNzE4IDAuMzY2LDAuMTc1IDAuODExLDAuMjY0IDEuMzM4LDAuMjY0IDAuMTE2LDAgMC4yNDEs"
        LogoBase64 = LogoBase64 & "LTAuMDEyIDAuMzYzLC0wLjAyMSB2IDIuNjkxIGMgLTAuMzQ5LC0wLjAwNSAtMi40NzMsLTAuMjIxIC0zLjMsLTEuOTY4IiAvPgogICAgPC9jbG"
        LogoBase64 = LogoBase64 & "lwUGF0aD4KICAgIDxjbGlwUGF0aAogICAgICAgaWQ9ImNsaXBQYXRoNDE4OSIKICAgICAgIGNsaXBQYXRoVW5pdHM9InVzZXJTcGFjZU9uVXNl"
        LogoBase64 = LogoBase64 & "Ij4KICAgICAgPHBhdGgKICAgICAgICAgaW5rc2NhcGU6Y29ubmVjdG9yLWN1cnZhdHVyZT0iMCIKICAgICAgICAgaWQ9InBhdGg0MTg3IgogIC"
        LogoBase64 = LogoBase64 & "AgICAgICBkPSJNIDMzMi4zODI2OCw5Ny40NDQxIFYgODYuNjk0MTQ0IEggMzI2LjMwOTMgViA5Ny40NDQxIFoiIC8+CiAgICA8L2NsaXBQYXRo"
        LogoBase64 = LogoBase64 & "PgogICAgPGxpbmVhckdyYWRpZW50CiAgICAgICBpZD0ibGluZWFyR3JhZGllbnQ0MTk1IgogICAgICAgc3ByZWFkTWV0aG9kPSJwYWQiCiAgIC"
        LogoBase64 = LogoBase64 & "AgICBncmFkaWVudFRyYW5zZm9ybT0ibWF0cml4KDAsLTMyLjM5OTAwMiwtMzIuMzk5MDAyLDAsMzI5LjM0NTk5LDEwNi44OTMwMSkiCiAgICAg"
        LogoBase64 = LogoBase64 & "ICBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSIKICAgICAgIHkyPSIwIgogICAgICAgeDI9IjEiCiAgICAgICB5MT0iMCIKICAgICAgIH"
        LogoBase64 = LogoBase64 & "gxPSIwIj4KICAgICAgPHN0b3AKICAgICAgICAgaWQ9InN0b3A0MTkxIgogICAgICAgICBvZmZzZXQ9IjAiCiAgICAgICAgIHN0eWxlPSJzdG9w"
        LogoBase64 = LogoBase64 & "LW9wYWNpdHk6MTtzdG9wLWNvbG9yOiMxNzlkZDkiIC8+CiAgICAgIDxzdG9wCiAgICAgICAgIGlkPSJzdG9wNDE5MyIKICAgICAgICAgb2Zmc2"
        LogoBase64 = LogoBase64 & "V0PSIxIgogICAgICAgICBzdHlsZT0ic3RvcC1vcGFjaXR5OjE7c3RvcC1jb2xvcjojMmIzMzdkIiAvPgogICAgPC9saW5lYXJHcmFkaWVudD4K"
        LogoBase64 = LogoBase64 & "ICAgIDxjbGlwUGF0aAogICAgICAgaWQ9ImNsaXBQYXRoNDEzNyIKICAgICAgIGNsaXBQYXRoVW5pdHM9InVzZXJTcGFjZU9uVXNlIj4KICAgIC"
        LogoBase64 = LogoBase64 & "AgPHBhdGgKICAgICAgICAgaW5rc2NhcGU6Y29ubmVjdG9yLWN1cnZhdHVyZT0iMCIKICAgICAgICAgaWQ9InBhdGg0MTM1IgogICAgICAgICBk"
        LogoBase64 = LogoBase64 & "PSJtIDM0My43MDUsOTcuNDQ2IGggMTAuMzI4IFYgODYuNDMgaCAtMTAuMzI4IHoiIC8+CiAgICA8L2NsaXBQYXRoPgogICAgPGNsaXBQYXRoCi"
        LogoBase64 = LogoBase64 & "AgICAgICBpZD0iY2xpcFBhdGg0MTQ1IgogICAgICAgY2xpcFBhdGhVbml0cz0idXNlclNwYWNlT25Vc2UiPgogICAgICA8cGF0aAogICAgICAg"
        LogoBase64 = LogoBase64 & "ICBpbmtzY2FwZTpjb25uZWN0b3ItY3VydmF0dXJlPSIwIgogICAgICAgICBpZD0icGF0aDQxNDMiCiAgICAgICAgIGQ9Im0gMzQ2Ljc4Myw5Ny"
        LogoBase64 = LogoBase64 & "4wMjEgYyAtMC42NjEsLTAuMjg0IC0xLjIyMiwtMC42NzkgLTEuNjgsLTEuMTg0IC0wLjQ2LC0wLjUwNyAtMC44MDcsLTEuMTA4IC0xLjA0Mywt"
        LogoBase64 = LogoBase64 & "MS44MDMgLTAuMjM3LC0wLjY5NSAtMC4zNTUsLTEuNDQ4IC0wLjM1NSwtMi4yNTggMCwtMC43ODIgMC4xMjksLTEuNTAyIDAuMzg1LC0yLjE1Ny"
        LogoBase64 = LogoBase64 & "AwLjI1NiwtMC42NTUgMC42MTMsLTEuMjE5IDEuMDczLC0xLjY5MSAwLjQ1OSwtMC40NzMgMS4wMTYsLTAuODQgMS42NywtMS4xMDQgMC42NTUs"
        LogoBase64 = LogoBase64 & "LTAuMjYyIDEuMzc0LC0wLjM5NCAyLjE1OCwtMC4zOTQgMS4zOSwwIDIuNTMxLDAuMzY0IDMuNDIyLDEuMDkzIDAuODkyLDAuNzI5IDEuNDMsMS"
        LogoBase64 = LogoBase64 & "43ODkgMS42MiwzLjE3OSBoIC0yLjc3NCBjIC0wLjA5NCwtMC42NDggLTAuMzI4LC0xLjE2NCAtMC42OTksLTEuNTQ4IC0wLjM3MSwtMC4zODUg"
        LogoBase64 = LogoBase64 & "LTAuOTAxLC0wLjU3OCAtMS41ODksLTAuNTc4IC0wLjQ0NiwwIC0wLjgyNCwwLjEwMSAtMS4xMzUsMC4zMDUgLTAuMzEsMC4yMDIgLTAuNTU4LD"
        LogoBase64 = LogoBase64 & "AuNDYxIC0wLjczOSwwLjc3OCAtMC4xODMsMC4zMTggLTAuMzE0LDAuNjcyIC0wLjM5NSwxLjA2MyAtMC4wOCwwLjM5MiAtMC4xMjEsMC43Nzgg"
        LogoBase64 = LogoBase64 & "LTAuMTIxLDEuMTU1IDAsMC4zOTEgMC4wNDEsMC43ODYgMC4xMjEsMS4xODUgMC4wODEsMC4zOTkgMC4yMTgsMC43NjMgMC40MTUsMS4wOTMgMC"
        LogoBase64 = LogoBase64 & "4xOTYsMC4zMzEgMC40NDksMC42MDEgMC43NTksMC44MSAwLjMxMSwwLjIxIDAuNjk2LDAuMzE1IDEuMTU1LDAuMzE1IDEuMjI4LDAgMS45Mzcs"
        LogoBase64 = LogoBase64 & "LTAuNjAxIDIuMTI3LC0xLjgwMiBoIDIuODE0IGMgLTAuMDQxLDAuNjc0IC0wLjIwMiwxLjI1OCAtMC40ODUsMS43NTEgLTAuMjgzLDAuNDkzIC"
        LogoBase64 = LogoBase64 & "0wLjY1MSwwLjkwNSAtMS4xMDUsMS4yMzYgLTAuNDUyLDAuMzI5IC0wLjk2NCwwLjU3NiAtMS41MzksMC43MzkgLTAuNTczLDAuMTYxIC0xLjE3"
        LogoBase64 = LogoBase64 & "MSwwLjI0MyAtMS43OTIsMC4yNDMgLTAuODUsMCAtMS42MDcsLTAuMTQyIC0yLjI2OCwtMC40MjYiIC8+CiAgICA8L2NsaXBQYXRoPgogICAgPG"
        LogoBase64 = LogoBase64 & "NsaXBQYXRoCiAgICAgICBpZD0iY2xpcFBhdGg0MTU1IgogICAgICAgY2xpcFBhdGhVbml0cz0idXNlclNwYWNlT25Vc2UiPgogICAgICA8cGF0"
        LogoBase64 = LogoBase64 & "aAogICAgICAgICBpbmtzY2FwZTpjb25uZWN0b3ItY3VydmF0dXJlPSIwIgogICAgICAgICBpZD0icGF0aDQxNTMiCiAgICAgICAgIGQ9Ik0gMz"
        LogoBase64 = LogoBase64 & "U0LjAzMzEsOTcuNDQ2NTMxIFYgODYuNDI5ODY2IGggLTEwLjMyNzg4IHYgMTEuMDE2NjY1IHoiIC8+CiAgICA8L2NsaXBQYXRoPgogICAgPGxp"
        LogoBase64 = LogoBase64 & "bmVhckdyYWRpZW50CiAgICAgICBpZD0ibGluZWFyR3JhZGllbnQ0MTYxIgogICAgICAgc3ByZWFkTWV0aG9kPSJwYWQiCiAgICAgICBncmFkaW"
        LogoBase64 = LogoBase64 & "VudFRyYW5zZm9ybT0ibWF0cml4KDAsLTMyLjQwMTAwMSwtMzIuNDAxMDAxLDAsMzQ4Ljg2OSwxMDYuODkzMDEpIgogICAgICAgZ3JhZGllbnRV"
        LogoBase64 = LogoBase64 & "bml0cz0idXNlclNwYWNlT25Vc2UiCiAgICAgICB5Mj0iMCIKICAgICAgIHgyPSIxIgogICAgICAgeTE9IjAiCiAgICAgICB4MT0iMCI+CiAgIC"
        LogoBase64 = LogoBase64 & "AgIDxzdG9wCiAgICAgICAgIGlkPSJzdG9wNDE1NyIKICAgICAgICAgb2Zmc2V0PSIwIgogICAgICAgICBzdHlsZT0ic3RvcC1vcGFjaXR5OjE7"
        LogoBase64 = LogoBase64 & "c3RvcC1jb2xvcjojMTc5ZGQ5IiAvPgogICAgICA8c3RvcAogICAgICAgICBpZD0ic3RvcDQxNTkiCiAgICAgICAgIG9mZnNldD0iMSIKICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgc3R5bGU9InN0b3Atb3BhY2l0eToxO3N0b3AtY29sb3I6IzJiMzM3ZCIgLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8Y2xpcFBh"
        LogoBase64 = LogoBase64 & "dGgKICAgICAgIGlkPSJjbGlwUGF0aDQxMDMiCiAgICAgICBjbGlwUGF0aFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+CiAgICAgIDxwYXRoCiAgIC"
        LogoBase64 = LogoBase64 & "AgICAgIGlua3NjYXBlOmNvbm5lY3Rvci1jdXJ2YXR1cmU9IjAiCiAgICAgICAgIGlkPSJwYXRoNDEwMSIKICAgICAgICAgZD0ibSAzMzIuNjA4"
        LogoBase64 = LogoBase64 & "LDk3LjQ0NiBoIDEwLjQ4NCBWIDg2LjQzIGggLTEwLjQ4NCB6IiAvPgogICAgPC9jbGlwUGF0aD4KICAgIDxjbGlwUGF0aAogICAgICAgaWQ9Im"
        LogoBase64 = LogoBase64 & "NsaXBQYXRoNDExMSIKICAgICAgIGNsaXBQYXRoVW5pdHM9InVzZXJTcGFjZU9uVXNlIj4KICAgICAgPHBhdGgKICAgICAgICAgaW5rc2NhcGU6"
        LogoBase64 = LogoBase64 & "Y29ubmVjdG9yLWN1cnZhdHVyZT0iMCIKICAgICAgICAgaWQ9InBhdGg0MTA5IgogICAgICAgICBkPSJtIDMzNS43NzcsOTcuMDIxIGMgLTAuNj"
        LogoBase64 = LogoBase64 & "U1LC0wLjI4NCAtMS4yMTksLTAuNjcyIC0xLjY5MSwtMS4xNjUgLTAuNDczLC0wLjQ5MiAtMC44MzgsLTEuMDc2IC0xLjA5NSwtMS43NTIgLTAu"
        LogoBase64 = LogoBase64 & "MjU1LC0wLjY3NSAtMC4zODMsLTEuNDAzIC0wLjM4MywtMi4xODcgMCwtMC44MDkgMC4xMjUsLTEuNTUyIDAuMzc0LC0yLjIyNyAwLjI1LC0wLj"
        LogoBase64 = LogoBase64 & "Y3NSAwLjYwNCwtMS4yNTUgMS4wNjMsLTEuNzQyIDAuNDU5LC0wLjQ4NSAxLjAyLC0wLjg2IDEuNjgxLC0xLjEyNCAwLjY2MSwtMC4yNjIgMS40"
        LogoBase64 = LogoBase64 & "MDUsLTAuMzk0IDIuMjI4LC0wLjM5NCAxLjE4OCwwIDIuMjAxLDAuMjcgMy4wMzcsMC44MDkgMC44MzgsMC41NCAxLjQ1OCwxLjQzOCAxLjg2My"
        LogoBase64 = LogoBase64 & "wyLjY5NCBoIC0yLjUzMSBjIC0wLjA5NCwtMC4zMjUgLTAuMzUyLC0wLjYzMSAtMC43NjgsLTAuOTIgLTAuNDIsLTAuMjkyIC0wLjkxOSwtMC40"
        LogoBase64 = LogoBase64 & "MzcgLTEuNTAxLC0wLjQzNyAtMC44MDksMCAtMS40MywwLjIxIC0xLjg2MiwwLjYyOCAtMC40MzMsMC40MTggLTAuNjY5LDEuMDkzIC0wLjcwOS"
        LogoBase64 = LogoBase64 & "wyLjAyNiBoIDcuNTUzIGMgMC4wNTUsMC44MDkgLTAuMDEyLDEuNTg2IC0wLjIwMiwyLjMyOSAtMC4xODksMC43NDEgLTAuNDk2LDEuNDAzIC0w"
        LogoBase64 = LogoBase64 & "LjkyLDEuOTg0IC0wLjQyNiwwLjU4MSAtMC45NywxLjA0MiAtMS42MzEsMS4zODcgLTAuNjYyLDAuMzQ1IC0xLjQzOSwwLjUxNyAtMi4zMjksMC"
        LogoBase64 = LogoBase64 & "41MTcgLTAuNzk2LDAgLTEuNTIzLC0wLjE0MiAtMi4xNzcsLTAuNDI2IG0gLTAuMTYzLC0zLjI4IGMgMC4wNzQsMC4yNTUgMC4yMDMsMC40OTkg"
        LogoBase64 = LogoBase64 & "MC4zODUsMC43MjggMC4xODMsMC4yMyAwLjQyNiwwLjQyMiAwLjcyOSwwLjU3OCAwLjMwNSwwLjE1NSAwLjY4NiwwLjIzMyAxLjE0NSwwLjIzMy"
        LogoBase64 = LogoBase64 & "AwLjcwMSwwIDEuMjI1LC0wLjE5IDEuNTcsLTAuNTY3IDAuMzQ1LC0wLjM3OSAwLjU4MywtMC45MzIgMC43MTgsLTEuNjYgaCAtNC42NzggYyAw"
        LogoBase64 = LogoBase64 & "LjAxMiwwLjIwMiAwLjA1OCwwLjQzMiAwLjEzMSwwLjY4OCIgLz4KICAgIDwvY2xpcFBhdGg+CiAgICA8Y2xpcFBhdGgKICAgICAgIGlkPSJjbG"
        LogoBase64 = LogoBase64 & "lwUGF0aDQxMjEiCiAgICAgICBjbGlwUGF0aFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+CiAgICAgIDxwYXRoCiAgICAgICAgIGlua3NjYXBlOmNv"
        LogoBase64 = LogoBase64 & "bm5lY3Rvci1jdXJ2YXR1cmU9IjAiCiAgICAgICAgIGlkPSJwYXRoNDExOSIKICAgICAgICAgZD0iTSAzNDMuMDkwNzEsOTcuNDQ2NTMxIFYgOD"
        LogoBase64 = LogoBase64 & "YuNDI5ODY2IGggLTEwLjQ4MzM4IHYgMTEuMDE2NjY1IHoiIC8+CiAgICA8L2NsaXBQYXRoPgogICAgPGxpbmVhckdyYWRpZW50CiAgICAgICBp"
        LogoBase64 = LogoBase64 & "ZD0ibGluZWFyR3JhZGllbnQ0MTI3IgogICAgICAgc3ByZWFkTWV0aG9kPSJwYWQiCiAgICAgICBncmFkaWVudFRyYW5zZm9ybT0ibWF0cml4KD"
        LogoBase64 = LogoBase64 & "AsLTMyLjQwMTAwMSwtMzIuNDAxMDAxLDAsMzM3LjgzMDk5LDEwNi44OTMwMSkiCiAgICAgICBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVz"
        LogoBase64 = LogoBase64 & "ZSIKICAgICAgIHkyPSIwIgogICAgICAgeDI9IjEiCiAgICAgICB5MT0iMCIKICAgICAgIHgxPSIwIj4KICAgICAgPHN0b3AKICAgICAgICAgaW"
        LogoBase64 = LogoBase64 & "Q9InN0b3A0MTIzIgogICAgICAgICBvZmZzZXQ9IjAiCiAgICAgICAgIHN0eWxlPSJzdG9wLW9wYWNpdHk6MTtzdG9wLWNvbG9yOiMxNzlkZDki"
        LogoBase64 = LogoBase64 & "IC8+CiAgICAgIDxzdG9wCiAgICAgICAgIGlkPSJzdG9wNDEyNSIKICAgICAgICAgb2Zmc2V0PSIxIgogICAgICAgICBzdHlsZT0ic3RvcC1vcG"
        LogoBase64 = LogoBase64 & "FjaXR5OjE7c3RvcC1jb2xvcjojMmIzMzdkIiAvPgogICAgPC9saW5lYXJHcmFkaWVudD4KICAgIDxjbGlwUGF0aAogICAgICAgaWQ9ImNsaXBQ"
        LogoBase64 = LogoBase64 & "YXRoNDA2OSIKICAgICAgIGNsaXBQYXRoVW5pdHM9InVzZXJTcGFjZU9uVXNlIj4KICAgICAgPHBhdGgKICAgICAgICAgaW5rc2NhcGU6Y29ubm"
        LogoBase64 = LogoBase64 & "VjdG9yLWN1cnZhdHVyZT0iMCIKICAgICAgICAgaWQ9InBhdGg0MDY3IgogICAgICAgICBkPSJtIDMxNC44NDcsOTcuMTYzIGggOS44MjIgViA4"
        LogoBase64 = LogoBase64 & "Ni40MyBoIC05LjgyMiB6IiAvPgogICAgPC9jbGlwUGF0aD4KICAgIDxjbGlwUGF0aAogICAgICAgaWQ9ImNsaXBQYXRoNDA3NyIKICAgICAgIG"
        LogoBase64 = LogoBase64 & "NsaXBQYXRoVW5pdHM9InVzZXJTcGFjZU9uVXNlIj4KICAgICAgPHBhdGgKICAgICAgICAgaW5rc2NhcGU6Y29ubmVjdG9yLWN1cnZhdHVyZT0i"
        LogoBase64 = LogoBase64 & "MCIKICAgICAgICAgaWQ9InBhdGg0MDc1IgogICAgICAgICBkPSJtIDMyMS43OTMsOTcuMTYzIHYgLTUuNDg5IGMgMCwtMS4wNjYgLTAuMTc2LC"
        LogoBase64 = LogoBase64 & "0xLjgzMyAtMC41MjYsLTIuMjk4IC0wLjM1MSwtMC40NjUgLTAuOTE5LC0wLjY5OSAtMS43MDEsLTAuNjk5IC0wLjY5LDAgLTEuMTY4LDAuMjEz"
        LogoBase64 = LogoBase64 & "IC0xLjQzOCwwLjYzOCAtMC4yNzEsMC40MjUgLTAuNDA1LDEuMDcgLTAuNDA1LDEuOTM1IHYgNS45MTMgaCAtMi44NzYgdiAtNi40NDEgYyAwLC"
        LogoBase64 = LogoBase64 & "0wLjY0OCAwLjA1NywtMS4yMzkgMC4xNzIsLTEuNzcxIDAuMTE0LC0wLjUzNCAwLjMxNSwtMC45ODUgMC41OTcsLTEuMzU3IDAuMjg0LC0wLjM3"
        LogoBase64 = LogoBase64 & "MSAwLjY3MiwtMC42NTggMS4xNjQsLTAuODYgMC40OTMsLTAuMjAyIDEuMTI1LC0wLjMwNCAxLjg5NSwtMC4zMDQgMC42MDYsMCAxLjIwMSwwLj"
        LogoBase64 = LogoBase64 & "EzNSAxLjc4MiwwLjQwNSAwLjU4MSwwLjI2OSAxLjA1MywwLjcwOCAxLjQxNywxLjMxNiBoIDAuMDYgdiAtMS40NTcgaCAyLjczNSB2IDEwLjQ2"
        LogoBase64 = LogoBase64 & "OSIgLz4KICAgIDwvY2xpcFBhdGg+CiAgICA8Y2xpcFBhdGgKICAgICAgIGlkPSJjbGlwUGF0aDQwODciCiAgICAgICBjbGlwUGF0aFVuaXRzPS"
        LogoBase64 = LogoBase64 & "J1c2VyU3BhY2VPblVzZSI+CiAgICAgIDxwYXRoCiAgICAgICAgIGlua3NjYXBlOmNvbm5lY3Rvci1jdXJ2YXR1cmU9IjAiCiAgICAgICAgIGlk"
        LogoBase64 = LogoBase64 & "PSJwYXRoNDA4NSIKICAgICAgICAgZD0ibSAzMjQuNjY4ODksOTcuMTYzMjE2IHYgLTEwLjczMzM1IGggLTkuODIxNzUgdiAxMC43MzMzNSB6Ii"
        LogoBase64 = LogoBase64 & "AvPgogICAgPC9jbGlwUGF0aD4KICAgIDxsaW5lYXJHcmFkaWVudAogICAgICAgaWQ9ImxpbmVhckdyYWRpZW50NDA5MyIKICAgICAgIHNwcmVh"
        LogoBase64 = LogoBase64 & "ZE1ldGhvZD0icGFkIgogICAgICAgZ3JhZGllbnRUcmFuc2Zvcm09Im1hdHJpeCgwLC0zMi40MDEwMDEsLTMyLjQwMTAwMSwwLDMxOS43NTgsMT"
        LogoBase64 = LogoBase64 & "A2Ljg5MzAxKSIKICAgICAgIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIgogICAgICAgeTI9IjAiCiAgICAgICB4Mj0iMSIKICAgICAg"
        LogoBase64 = LogoBase64 & "IHkxPSIwIgogICAgICAgeDE9IjAiPgogICAgICA8c3RvcAogICAgICAgICBpZD0ic3RvcDQwODkiCiAgICAgICAgIG9mZnNldD0iMCIKICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgc3R5bGU9InN0b3Atb3BhY2l0eToxO3N0b3AtY29sb3I6IzE3OWRkOSIgLz4KICAgICAgPHN0b3AKICAgICAgICAgaWQ9InN0b3A0MDkx"
        LogoBase64 = LogoBase64 & "IgogICAgICAgICBvZmZzZXQ9IjEiCiAgICAgICAgIHN0eWxlPSJzdG9wLW9wYWNpdHk6MTtzdG9wLWNvbG9yOiMyYjMzN2QiIC8+CiAgICA8L2"
        LogoBase64 = LogoBase64 & "xpbmVhckdyYWRpZW50PgogICAgPGNsaXBQYXRoCiAgICAgICBpZD0iY2xpcFBhdGg0MDM1IgogICAgICAgY2xpcFBhdGhVbml0cz0idXNlclNw"
        LogoBase64 = LogoBase64 & "YWNlT25Vc2UiPgogICAgICA8cGF0aAogICAgICAgICBpbmtzY2FwZTpjb25uZWN0b3ItY3VydmF0dXJlPSIwIgogICAgICAgICBpZD0icGF0aD"
        LogoBase64 = LogoBase64 & "QwMzMiCiAgICAgICAgIGQ9Im0gMzAyLjk0NSw5Ny40NDYgaCAxMC44MzMgViA4Ni40MyBoIC0xMC44MzMgeiIgLz4KICAgIDwvY2xpcFBhdGg+"
        LogoBase64 = LogoBase64 & "CiAgICA8Y2xpcFBhdGgKICAgICAgIGlkPSJjbGlwUGF0aDQwNDMiCiAgICAgICBjbGlwUGF0aFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+CiAgIC"
        LogoBase64 = LogoBase64 & "AgIDxwYXRoCiAgICAgICAgIGlua3NjYXBlOmNvbm5lY3Rvci1jdXJ2YXR1cmU9IjAiCiAgICAgICAgIGlkPSJwYXRoNDA0MSIKICAgICAgICAg"
        LogoBase64 = LogoBase64 & "ZD0ibSAzMDYuMTI1LDk3LjA1MSBjIC0wLjY2MiwtMC4yNjIgLTEuMjMxLC0wLjYzNSAtMS43MDEsLTEuMTEzIC0wLjQ3NCwtMC40OCAtMC44Mz"
        LogoBase64 = LogoBase64 & "gsLTEuMDYgLTEuMDk1LC0xLjc0MiAtMC4yNTYsLTAuNjgxIC0wLjM4NCwtMS40NDEgLTAuMzg0LC0yLjI3OSAwLC0wLjgzNyAwLjEyOCwtMS41"
        LogoBase64 = LogoBase64 & "OTIgMC4zODQsLTIuMjY3IDAuMjU3LC0wLjY3NiAwLjYyMSwtMS4yNTMgMS4wOTUsLTEuNzMyIDAuNDcsLTAuNDc5IDEuMDM5LC0wLjg0OCAxLj"
        LogoBase64 = LogoBase64 & "cwMSwtMS4xMDQgMC42NjEsLTAuMjU2IDEuNDAyLC0wLjM4NCAyLjIyNiwtMC4zODQgMC44MjQsMCAxLjU3LDAuMTI4IDIuMjM4LDAuMzg0IDAu"
        LogoBase64 = LogoBase64 & "NjY5LDAuMjU2IDEuMjM5LDAuNjI1IDEuNzExLDEuMTA0IDAuNDczLDAuNDc5IDAuODM4LDEuMDU2IDEuMDk0LDEuNzMyIDAuMjU3LDAuNjc1ID"
        LogoBase64 = LogoBase64 & "AuMzg1LDEuNDMgMC4zODUsMi4yNjcgMCwwLjgzOCAtMC4xMjgsMS41OTggLTAuMzg1LDIuMjc5IC0wLjI1NiwwLjY4MiAtMC42MjEsMS4yNjIg"
        LogoBase64 = LogoBase64 & "LTEuMDk0LDEuNzQyIC0wLjQ3MiwwLjQ3OCAtMS4wNDIsMC44NTEgLTEuNzExLDEuMTEzIC0wLjY2OCwwLjI2NCAtMS40MTQsMC4zOTYgLTIuMj"
        LogoBase64 = LogoBase64 & "M4LDAuMzk2IC0wLjgyNCwwIC0xLjU2NSwtMC4xMzIgLTIuMjI2LC0wLjM5NiBtIDEuMDEyLC04LjE5MSBjIC0wLjMyNCwwLjE4OSAtMC41ODQs"
        LogoBase64 = LogoBase64 & "MC40NDIgLTAuNzgsMC43NTkgLTAuMTk2LDAuMzE3IC0wLjMzMywwLjY3NSAtMC40MTUsMS4wNzUgLTAuMDgxLDAuMzk3IC0wLjEyMSwwLjgwNi"
        LogoBase64 = LogoBase64 & "AtMC4xMjEsMS4yMjMgMCwwLjQxOSAwLjA0LDAuODMxIDAuMTIxLDEuMjM2IDAuMDgyLDAuNDA2IDAuMjE5LDAuNzY0IDAuNDE1LDEuMDczIDAu"
        LogoBase64 = LogoBase64 & "MTk2LDAuMzExIDAuNDU2LDAuNTY0IDAuNzgsMC43NiAwLjMyNCwwLjE5NiAwLjcyOCwwLjI5NCAxLjIxNCwwLjI5NCAwLjQ4NywwIDAuODk1LC"
        LogoBase64 = LogoBase64 & "0wLjA5OCAxLjIyNSwtMC4yOTQgMC4zMzIsLTAuMTk2IDAuNTk1LC0wLjQ0OSAwLjc5LC0wLjc2IDAuMTk3LC0wLjMwOSAwLjMzNSwtMC42Njcg"
        LogoBase64 = LogoBase64 & "MC40MTYsLTEuMDczIDAuMDgxLC0wLjQwNSAwLjEyMiwtMC44MTcgMC4xMjIsLTEuMjM2IDAsLTAuNDE3IC0wLjA0MSwtMC44MjYgLTAuMTIyLC"
        LogoBase64 = LogoBase64 & "0xLjIyMyAtMC4wODEsLTAuNCAtMC4yMTksLTAuNzU4IC0wLjQxNiwtMS4wNzUgLTAuMTk1LC0wLjMxNyAtMC40NTgsLTAuNTcgLTAuNzksLTAu"
        LogoBase64 = LogoBase64 & "NzU5IC0wLjMzLC0wLjE4OSAtMC43MzgsLTAuMjg0IC0xLjIyNSwtMC4yODQgLTAuNDg2LDAgLTAuODksMC4wOTUgLTEuMjE0LDAuMjg0IiAvPg"
        LogoBase64 = LogoBase64 & "ogICAgPC9jbGlwUGF0aD4KICAgIDxjbGlwUGF0aAogICAgICAgaWQ9ImNsaXBQYXRoNDA1MyIKICAgICAgIGNsaXBQYXRoVW5pdHM9InVzZXJT"
        LogoBase64 = LogoBase64 & "cGFjZU9uVXNlIj4KICAgICAgPHBhdGgKICAgICAgICAgaW5rc2NhcGU6Y29ubmVjdG9yLWN1cnZhdHVyZT0iMCIKICAgICAgICAgaWQ9InBhdG"
        LogoBase64 = LogoBase64 & "g0MDUxIgogICAgICAgICBkPSJNIDMxMy43Nzg2Nyw5Ny40NDY1MzEgViA4Ni40Mjk4NjYgaCAtMTAuODMzMzEgdiAxMS4wMTY2NjUgeiIgLz4K"
        LogoBase64 = LogoBase64 & "ICAgIDwvY2xpcFBhdGg+CiAgICA8bGluZWFyR3JhZGllbnQKICAgICAgIGlkPSJsaW5lYXJHcmFkaWVudDQwNTkiCiAgICAgICBzcHJlYWRNZX"
        LogoBase64 = LogoBase64 & "Rob2Q9InBhZCIKICAgICAgIGdyYWRpZW50VHJhbnNmb3JtPSJtYXRyaXgoMCwtMzIuNDAxMDAxLC0zMi40MDEwMDEsMCwzMDguMzYyLDEwNi44"
        LogoBase64 = LogoBase64 & "OTMwMSkiCiAgICAgICBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSIKICAgICAgIHkyPSIwIgogICAgICAgeDI9IjEiCiAgICAgICB5MT"
        LogoBase64 = LogoBase64 & "0iMCIKICAgICAgIHgxPSIwIj4KICAgICAgPHN0b3AKICAgICAgICAgaWQ9InN0b3A0MDU1IgogICAgICAgICBvZmZzZXQ9IjAiCiAgICAgICAg"
        LogoBase64 = LogoBase64 & "IHN0eWxlPSJzdG9wLW9wYWNpdHk6MTtzdG9wLWNvbG9yOiMxNzlkZDkiIC8+CiAgICAgIDxzdG9wCiAgICAgICAgIGlkPSJzdG9wNDA1NyIKIC"
        LogoBase64 = LogoBase64 & "AgICAgICAgb2Zmc2V0PSIxIgogICAgICAgICBzdHlsZT0ic3RvcC1vcGFjaXR5OjE7c3RvcC1jb2xvcjojMmIzMzdkIiAvPgogICAgPC9saW5l"
        LogoBase64 = LogoBase64 & "YXJHcmFkaWVudD4KICAgIDxjbGlwUGF0aAogICAgICAgaWQ9ImNsaXBQYXRoNDAwOSIKICAgICAgIGNsaXBQYXRoVW5pdHM9InVzZXJTcGFjZU"
        LogoBase64 = LogoBase64 & "9uVXNlIj4KICAgICAgPHBhdGgKICAgICAgICAgaW5rc2NhcGU6Y29ubmVjdG9yLWN1cnZhdHVyZT0iMCIKICAgICAgICAgaWQ9InBhdGg0MDA3"
        LogoBase64 = LogoBase64 & "IgogICAgICAgICBkPSJtIDI5OC44MDUsMTAxLjE1MiBoIDIuODc3IFYgODYuNjkzIGggLTIuODc3IHoiIC8+CiAgICA8L2NsaXBQYXRoPgogIC"
        LogoBase64 = LogoBase64 & "AgPGNsaXBQYXRoCiAgICAgICBpZD0iY2xpcFBhdGg0MDE5IgogICAgICAgY2xpcFBhdGhVbml0cz0idXNlclNwYWNlT25Vc2UiPgogICAgICA8"
        LogoBase64 = LogoBase64 & "cGF0aAogICAgICAgICBpbmtzY2FwZTpjb25uZWN0b3ItY3VydmF0dXJlPSIwIgogICAgICAgICBpZD0icGF0aDQwMTciCiAgICAgICAgIGQ9Ik"
        LogoBase64 = LogoBase64 & "0gMzAxLjY4MTM3LDEwMS4xNTI0OCBWIDg2LjY5MzcyIGggLTIuODc2MzUgdiAxNC40NTg3NiB6IiAvPgogICAgPC9jbGlwUGF0aD4KICAgIDxs"
        LogoBase64 = LogoBase64 & "aW5lYXJHcmFkaWVudAogICAgICAgaWQ9ImxpbmVhckdyYWRpZW50NDAyNSIKICAgICAgIHNwcmVhZE1ldGhvZD0icGFkIgogICAgICAgZ3JhZG"
        LogoBase64 = LogoBase64 & "llbnRUcmFuc2Zvcm09Im1hdHJpeCgwLC0zMi4zOTk5OTQsLTMyLjM5OTk5NCwwLDMwMC4yNDMsMTA2Ljg5MzAxKSIKICAgICAgIGdyYWRpZW50"
        LogoBase64 = LogoBase64 & "VW5pdHM9InVzZXJTcGFjZU9uVXNlIgogICAgICAgeTI9IjAiCiAgICAgICB4Mj0iMSIKICAgICAgIHkxPSIwIgogICAgICAgeDE9IjAiPgogIC"
        LogoBase64 = LogoBase64 & "AgICA8c3RvcAogICAgICAgICBpZD0ic3RvcDQwMjEiCiAgICAgICAgIG9mZnNldD0iMCIKICAgICAgICAgc3R5bGU9InN0b3Atb3BhY2l0eTox"
        LogoBase64 = LogoBase64 & "O3N0b3AtY29sb3I6IzE3OWRkOSIgLz4KICAgICAgPHN0b3AKICAgICAgICAgaWQ9InN0b3A0MDIzIgogICAgICAgICBvZmZzZXQ9IjEiCiAgIC"
        LogoBase64 = LogoBase64 & "AgICAgIHN0eWxlPSJzdG9wLW9wYWNpdHk6MTtzdG9wLWNvbG9yOiMyYjMzN2QiIC8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPGNsaXBQ"
        LogoBase64 = LogoBase64 & "YXRoCiAgICAgICBpZD0iY2xpcFBhdGgzOTgzIgogICAgICAgY2xpcFBhdGhVbml0cz0idXNlclNwYWNlT25Vc2UiPgogICAgICA8cGF0aAogIC"
        LogoBase64 = LogoBase64 & "AgICAgICBpbmtzY2FwZTpjb25uZWN0b3ItY3VydmF0dXJlPSIwIgogICAgICAgICBpZD0icGF0aDM5ODEiCiAgICAgICAgIGQ9Im0gMjk0LjM4"
        LogoBase64 = LogoBase64 & "NCwxMDEuMTUyIGggMi44NzUgViA4Ni42OTMgaCAtMi44NzUgeiIgLz4KICAgIDwvY2xpcFBhdGg+CiAgICA8Y2xpcFBhdGgKICAgICAgIGlkPS"
        LogoBase64 = LogoBase64 & "JjbGlwUGF0aDM5OTMiCiAgICAgICBjbGlwUGF0aFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+CiAgICAgIDxwYXRoCiAgICAgICAgIGlua3NjYXBl"
        LogoBase64 = LogoBase64 & "OmNvbm5lY3Rvci1jdXJ2YXR1cmU9IjAiCiAgICAgICAgIGlkPSJwYXRoMzk5MSIKICAgICAgICAgZD0iTSAyOTcuMjU4NDMsMTAxLjE1MjQ4IF"
        LogoBase64 = LogoBase64 & "YgODYuNjkzNzIgaCAtMi44NzQ3OSB2IDE0LjQ1ODc2IHoiIC8+CiAgICA8L2NsaXBQYXRoPgogICAgPGxpbmVhckdyYWRpZW50CiAgICAgICBp"
        LogoBase64 = LogoBase64 & "ZD0ibGluZWFyR3JhZGllbnQzOTk5IgogICAgICAgc3ByZWFkTWV0aG9kPSJwYWQiCiAgICAgICBncmFkaWVudFRyYW5zZm9ybT0ibWF0cml4KD"
        LogoBase64 = LogoBase64 & "AsLTMyLjM5OTk5NCwtMzIuMzk5OTk0LDAsMjk1LjgyMSwxMDYuODkzMDEpIgogICAgICAgZ3JhZGllbnRVbml0cz0idXNlclNwYWNlT25Vc2Ui"
        LogoBase64 = LogoBase64 & "CiAgICAgICB5Mj0iMCIKICAgICAgIHgyPSIxIgogICAgICAgeTE9IjAiCiAgICAgICB4MT0iMCI+CiAgICAgIDxzdG9wCiAgICAgICAgIGlkPS"
        LogoBase64 = LogoBase64 & "JzdG9wMzk5NSIKICAgICAgICAgb2Zmc2V0PSIwIgogICAgICAgICBzdHlsZT0ic3RvcC1vcGFjaXR5OjE7c3RvcC1jb2xvcjojMTc5ZGQ5IiAv"
        LogoBase64 = LogoBase64 & "PgogICAgICA8c3RvcAogICAgICAgICBpZD0ic3RvcDM5OTciCiAgICAgICAgIG9mZnNldD0iMSIKICAgICAgICAgc3R5bGU9InN0b3Atb3BhY2"
        LogoBase64 = LogoBase64 & "l0eToxO3N0b3AtY29sb3I6IzJiMzM3ZCIgLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8Y2xpcFBhdGgKICAgICAgIGlkPSJjbGlwUGF0"
        LogoBase64 = LogoBase64 & "aDM5NDkiCiAgICAgICBjbGlwUGF0aFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+CiAgICAgIDxwYXRoCiAgICAgICAgIGlua3NjYXBlOmNvbm5lY3"
        LogoBase64 = LogoBase64 & "Rvci1jdXJ2YXR1cmU9IjAiCiAgICAgICAgIGlkPSJwYXRoMzk0NyIKICAgICAgICAgZD0ibSAyODMuMTE2LDk3LjQ0NiBoIDkuOTI2IFYgODYu"
        LogoBase64 = LogoBase64 & "NDMgaCAtOS45MjYgeiIgLz4KICAgIDwvY2xpcFBhdGg+CiAgICA8Y2xpcFBhdGgKICAgICAgIGlkPSJjbGlwUGF0aDM5NTciCiAgICAgICBjbG"
        LogoBase64 = LogoBase64 & "lwUGF0aFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+CiAgICAgIDxwYXRoCiAgICAgICAgIGlua3NjYXBlOmNvbm5lY3Rvci1jdXJ2YXR1cmU9IjAi"
        LogoBase64 = LogoBase64 & "CiAgICAgICAgIGlkPSJwYXRoMzk1NSIKICAgICAgICAgZD0ibSAyODYuNjMsOTcuMjc1IGMgLTAuNTc0LC0wLjExNiAtMS4wOTEsLTAuMzA4IC"
        LogoBase64 = LogoBase64 & "0xLjU0OSwtMC41NzkgLTAuNDYsLTAuMjcgLTAuODM5LC0wLjYyNiAtMS4xMzUsLTEuMDcyIC0wLjI5NiwtMC40NDUgLTAuNDY2LC0xLjAwNiAt"
        LogoBase64 = LogoBase64 & "MC41MDUsLTEuNjgxIGggMi44NzUgYyAwLjA1MywwLjU2NyAwLjI0NCwwLjk3MyAwLjU2NywxLjIxNiAwLjMyMywwLjI0MiAwLjc2OSwwLjM2NC"
        LogoBase64 = LogoBase64 & "AxLjMzNywwLjM2NCAwLjI1NSwwIDAuNDk1LC0wLjAxOCAwLjcxOCwtMC4wNTEgMC4yMjQsLTAuMDM0IDAuNDE5LC0wLjEwMSAwLjU4NywtMC4y"
        LogoBase64 = LogoBase64 & "MDIgMC4xNywtMC4xMDIgMC4zMDUsLTAuMjQzIDAuNDA2LC0wLjQyNSAwLjEwMSwtMC4xODMgMC4xNTIsLTAuNDMgMC4xNTIsLTAuNzQxIDAuMD"
        LogoBase64 = LogoBase64 & "EzLC0wLjI5NiAtMC4wNzUsLTAuNTIzIC0wLjI2NCwtMC42NzcgLTAuMTg5LC0wLjE1NSAtMC40NDUsLTAuMjc0IC0wLjc3LC0wLjM1NSAtMC4z"
        LogoBase64 = LogoBase64 & "MjQsLTAuMDggLTAuNjk1LC0wLjE0MiAtMS4xMTMsLTAuMTgyIC0wLjQxOCwtMC4wNDEgLTAuODQ0LC0wLjA5NSAtMS4yNzcsLTAuMTYzIC0wLj"
        LogoBase64 = LogoBase64 & "QzMiwtMC4wNjcgLTAuODU5LC0wLjE1NyAtMS4yODUsLTAuMjcyIC0wLjQyNSwtMC4xMTUgLTAuODA0LC0wLjI4OCAtMS4xMzQsLTAuNTE3IC0w"
        LogoBase64 = LogoBase64 & "LjMzMSwtMC4yMyAtMC42MDEsLTAuNTM2IC0wLjgxLC0wLjkyIC0wLjIxLC0wLjM4NiAtMC4zMTQsLTAuODc2IC0wLjMxNCwtMS40NyAwLC0wLj"
        LogoBase64 = LogoBase64 & "U0IDAuMDkxLC0xLjAwNiAwLjI3MywtMS4zOTcgMC4xODMsLTAuMzkyIDAuNDM2LC0wLjcxNiAwLjc2LC0wLjk3MiAwLjMyNCwtMC4yNTcgMC43"
        LogoBase64 = LogoBase64 & "MDIsLTAuNDQ1IDEuMTM0LC0wLjU2NyAwLjQzMiwtMC4xMjEgMC44OTgsLTAuMTgyIDEuMzk3LC0wLjE4MiAwLjY0OCwwIDEuMjgzLDAuMDk0ID"
        LogoBase64 = LogoBase64 & "EuOTAzLDAuMjgzIDAuNjIyLDAuMTg5IDEuMTYyLDAuNTIgMS42MjEsMC45OTIgMC4wMTYsLTAuMjQ4IDAuMTM1LC0wLjgyIDAuMjAzLC0xLjAx"
        LogoBase64 = LogoBase64 & "MSBoIDIuNjM1IGMgMCwwIC0wLjA4MywxLjIyNSAtMC4wODMsMi4zMjggdiA1LjQ0NyBjIDAsMC42MzYgLTAuMTQxLDEuMTQ0IC0wLjQyNSwxLj"
        LogoBase64 = LogoBase64 & "UyOSAtMC4yODQsMC4zODUgLTAuNjQ5LDAuNjg1IC0xLjA5NSwwLjkwMSAtMC40NDUsMC4yMTcgLTAuOTM4LDAuMzYyIC0xLjQ3OCwwLjQzNiAt"
        LogoBase64 = LogoBase64 & "MC41NCwwLjA3NCAtMS4wNzQsMC4xMTIgLTEuNTk5LDAuMTEyIC0wLjU4MSwwIC0xLjE1OSwtMC4wNTggLTEuNzMyLC0wLjE3MiBtIDAuNDU2LC"
        LogoBase64 = LogoBase64 & "04Ljg4MSBjIC0wLjIxOCwwLjA0MSAtMC40MDYsMC4xMTEgLTAuNTY4LDAuMjEyIC0wLjE2MiwwLjEwMyAtMC4yOTEsMC4yNDEgLTAuMzg0LDAu"
        LogoBase64 = LogoBase64 & "NDE2IC0wLjA5NSwwLjE3NiAtMC4xNDMsMC4zOTEgLTAuMTQzLDAuNjQ2IDAsMC4yNzIgMC4wNDgsMC40OTQgMC4xNDMsMC42NyAwLjA5MywwLj"
        LogoBase64 = LogoBase64 & "E3NSAwLjIxOSwwLjMyIDAuMzczLDAuNDM2IDAuMTU1LDAuMTE1IDAuMzM5LDAuMjA1IDAuNTQ4LDAuMjc0IDAuMjEsMC4wNjYgMC40MjMsMC4x"
        LogoBase64 = LogoBase64 & "MjEgMC42MzgsMC4xNjEgMC4yMjksMC4wNDEgMC40NTksMC4wNzUgMC42ODgsMC4xMDIgMC4yMywwLjAyNyAwLjQ0OCwwLjA2MSAwLjY1OCwwLj"
        LogoBase64 = LogoBase64 & "EwMSAwLjIxLDAuMDQgMC40MDYsMC4wOTIgMC41ODgsMC4xNTIgMC4xODIsMC4wNiAwLjMzNCwwLjE0NCAwLjQ1NiwwLjI1MiB2IC0xLjA3MiBj"
        LogoBase64 = LogoBase64 & "IDAsLTAuMTYzIC0wLjAxOCwtMC4zNzkgLTAuMDUxLC0wLjY0OCAtMC4wMzQsLTAuMjcxIC0wLjEyNSwtMC41MzggLTAuMjczLC0wLjgwMiAtMC"
        LogoBase64 = LogoBase64 & "4xNDksLTAuMjYyIC0wLjM3OCwtMC40ODggLTAuNjksLTAuNjc3IC0wLjMxLC0wLjE5IC0wLjc0NywtMC4yODQgLTEuMzE1LC0wLjI4NCAtMC4y"
        LogoBase64 = LogoBase64 & "MywwIC0wLjQ1MiwwLjAyIC0wLjY2OCwwLjA2MSIgLz4KICAgIDwvY2xpcFBhdGg+CiAgICA8Y2xpcFBhdGgKICAgICAgIGlkPSJjbGlwUGF0aD"
        LogoBase64 = LogoBase64 & "M5NjciCiAgICAgICBjbGlwUGF0aFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+CiAgICAgIDxwYXRoCiAgICAgICAgIGlua3NjYXBlOmNvbm5lY3Rv"
        LogoBase64 = LogoBase64 & "ci1jdXJ2YXR1cmU9IjAiCiAgICAgICAgIGlkPSJwYXRoMzk2NSIKICAgICAgICAgZD0iTSAyOTMuMDQxNTksOTcuNDQ2NTMxIFYgODYuNDI5OD"
        LogoBase64 = LogoBase64 & "Y2IGggLTkuOTI1MzMgdiAxMS4wMTY2NjUgeiIgLz4KICAgIDwvY2xpcFBhdGg+CiAgICA8bGluZWFyR3JhZGllbnQKICAgICAgIGlkPSJsaW5l"
        LogoBase64 = LogoBase64 & "YXJHcmFkaWVudDM5NzMiCiAgICAgICBzcHJlYWRNZXRob2Q9InBhZCIKICAgICAgIGdyYWRpZW50VHJhbnNmb3JtPSJtYXRyaXgoMCwtMzIuND"
        LogoBase64 = LogoBase64 & "AxMDAxLC0zMi40MDEwMDEsMCwyODguMDc4OTksMTA2Ljg5MzAxKSIKICAgICAgIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIgogICAg"
        LogoBase64 = LogoBase64 & "ICAgeTI9IjAiCiAgICAgICB4Mj0iMSIKICAgICAgIHkxPSIwIgogICAgICAgeDE9IjAiPgogICAgICA8c3RvcAogICAgICAgICBpZD0ic3RvcD"
        LogoBase64 = LogoBase64 & "M5NjkiCiAgICAgICAgIG9mZnNldD0iMCIKICAgICAgICAgc3R5bGU9InN0b3Atb3BhY2l0eToxO3N0b3AtY29sb3I6IzE3OWRkOSIgLz4KICAg"
        LogoBase64 = LogoBase64 & "ICAgPHN0b3AKICAgICAgICAgaWQ9InN0b3AzOTcxIgogICAgICAgICBvZmZzZXQ9IjEiCiAgICAgICAgIHN0eWxlPSJzdG9wLW9wYWNpdHk6MT"
        LogoBase64 = LogoBase64 & "tzdG9wLWNvbG9yOiMyYjMzN2QiIC8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPGNsaXBQYXRoCiAgICAgICBpZD0iY2xpcFBhdGgzOTE1"
        LogoBase64 = LogoBase64 & "IgogICAgICAgY2xpcFBhdGhVbml0cz0idXNlclNwYWNlT25Vc2UiPgogICAgICA8cGF0aAogICAgICAgICBpbmtzY2FwZTpjb25uZWN0b3ItY3"
        LogoBase64 = LogoBase64 & "VydmF0dXJlPSIwIgogICAgICAgICBpZD0icGF0aDM5MTMiCiAgICAgICAgIGQ9Im0gMjcyLjc0Nyw5Ny4xNjMgaCAxMC4zMjkgdiAtMTAuNDcg"
        LogoBase64 = LogoBase64 & "aCAtMTAuMzI5IHoiIC8+CiAgICA8L2NsaXBQYXRoPgogICAgPGNsaXBQYXRoCiAgICAgICBpZD0iY2xpcFBhdGgzOTIzIgogICAgICAgY2xpcF"
        LogoBase64 = LogoBase64 & "BhdGhVbml0cz0idXNlclNwYWNlT25Vc2UiPgogICAgICA8cGF0aAogICAgICAgICBpbmtzY2FwZTpjb25uZWN0b3ItY3VydmF0dXJlPSIwIgog"
        LogoBase64 = LogoBase64 & "ICAgICAgICBpZD0icGF0aDM5MjEiCiAgICAgICAgIGQ9Im0gMjgwLjIyLDk3LjE2MyAtMi4yMDcsLTcuMTQ5IGggLTAuMDQgbCAtMi4yMDgsNy"
        LogoBase64 = LogoBase64 & "4xNDkgaCAtMy4wMTggbCAzLjU4NSwtMTAuNDY5IGggMy4xOTkgbCAzLjU0NSwxMC40NjkiIC8+CiAgICA8L2NsaXBQYXRoPgogICAgPGNsaXBQ"
        LogoBase64 = LogoBase64 & "YXRoCiAgICAgICBpZD0iY2xpcFBhdGgzOTMzIgogICAgICAgY2xpcFBhdGhVbml0cz0idXNlclNwYWNlT25Vc2UiPgogICAgICA8cGF0aAogIC"
        LogoBase64 = LogoBase64 & "AgICAgICBpbmtzY2FwZTpjb25uZWN0b3ItY3VydmF0dXJlPSIwIgogICAgICAgICBpZD0icGF0aDM5MzEiCiAgICAgICAgIGQ9Ik0gMjgzLjA3"
        LogoBase64 = LogoBase64 & "NTc1LDk3LjE2MzA5OCBWIDg2LjY5MzcxNyBIIDI3Mi43NDc2IHYgMTAuNDY5MzgxIHoiIC8+CiAgICA8L2NsaXBQYXRoPgogICAgPGxpbmVhck"
        LogoBase64 = LogoBase64 & "dyYWRpZW50CiAgICAgICBpZD0ibGluZWFyR3JhZGllbnQzOTM5IgogICAgICAgc3ByZWFkTWV0aG9kPSJwYWQiCiAgICAgICBncmFkaWVudFRy"
        LogoBase64 = LogoBase64 & "YW5zZm9ybT0ibWF0cml4KDAsLTMyLjM5OTAwMiwtMzIuMzk5MDAyLDAsMjc3LjkxMiwxMDYuODkyKSIKICAgICAgIGdyYWRpZW50VW5pdHM9In"
        LogoBase64 = LogoBase64 & "VzZXJTcGFjZU9uVXNlIgogICAgICAgeTI9IjAiCiAgICAgICB4Mj0iMSIKICAgICAgIHkxPSIwIgogICAgICAgeDE9IjAiPgogICAgICA8c3Rv"
        LogoBase64 = LogoBase64 & "cAogICAgICAgICBpZD0ic3RvcDM5MzUiCiAgICAgICAgIG9mZnNldD0iMCIKICAgICAgICAgc3R5bGU9InN0b3Atb3BhY2l0eToxO3N0b3AtY2"
        LogoBase64 = LogoBase64 & "9sb3I6IzE3OWRkOSIgLz4KICAgICAgPHN0b3AKICAgICAgICAgaWQ9InN0b3AzOTM3IgogICAgICAgICBvZmZzZXQ9IjEiCiAgICAgICAgIHN0"
        LogoBase64 = LogoBase64 & "eWxlPSJzdG9wLW9wYWNpdHk6MTtzdG9wLWNvbG9yOiMyYjMzN2QiIC8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPGNsaXBQYXRoCiAgIC"
        LogoBase64 = LogoBase64 & "AgICBpZD0iY2xpcFBhdGgzODgxIgogICAgICAgY2xpcFBhdGhVbml0cz0idXNlclNwYWNlT25Vc2UiPgogICAgICA8cGF0aAogICAgICAgICBp"
        LogoBase64 = LogoBase64 & "bmtzY2FwZTpjb25uZWN0b3ItY3VydmF0dXJlPSIwIgogICAgICAgICBpZD0icGF0aDM4NzkiCiAgICAgICAgIGQ9Im0gMjU2LjY2MSwxMDMuMz"
        LogoBase64 = LogoBase64 & "IgaCAxMi4xODMgViA5MS4xMzIgaCAtMTIuMTgzIHoiIC8+CiAgICA8L2NsaXBQYXRoPgogICAgPGNsaXBQYXRoCiAgICAgICBpZD0iY2xpcFBh"
        LogoBase64 = LogoBase64 & "dGgzODg5IgogICAgICAgY2xpcFBhdGhVbml0cz0idXNlclNwYWNlT25Vc2UiPgogICAgICA8cGF0aAogICAgICAgICBpbmtzY2FwZTpjb25uZW"
        LogoBase64 = LogoBase64 & "N0b3ItY3VydmF0dXJlPSIwIgogICAgICAgICBpZD0icGF0aDM4ODciCiAgICAgICAgIGQ9Im0gMjU2LjY2MSw5Ny4yMjcgYyAwLC0zLjM2NiAy"
        LogoBase64 = LogoBase64 & "LjczNCwtNi4wOTQgNi4wOTYsLTYuMDk0IDMuMzY0LDAgNi4wODcsMi43MjggNi4wODcsNi4wOTQgMCwzLjM2NSAtMi43MjMsNi4wOTMgLTYuMD"
        LogoBase64 = LogoBase64 & "g3LDYuMDkzIC0zLjM2MiwwIC02LjA5NiwtMi43MjggLTYuMDk2LC02LjA5MyIgLz4KICAgIDwvY2xpcFBhdGg+CiAgICA8Y2xpcFBhdGgKICAg"
        LogoBase64 = LogoBase64 & "ICAgIGlkPSJjbGlwUGF0aDM4OTkiCiAgICAgICBjbGlwUGF0aFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+CiAgICAgIDxwYXRoCiAgICAgICAgIG"
        LogoBase64 = LogoBase64 & "lua3NjYXBlOmNvbm5lY3Rvci1jdXJ2YXR1cmU9IjAiCiAgICAgICAgIGlkPSJwYXRoMzg5NyIKICAgICAgICAgZD0iTSAyNjguODQzMDMsMTAz"
        LogoBase64 = LogoBase64 & "LjMyMDMgViA5MS4xMzI5NDYgSCAyNTYuNjYwOSBWIDEwMy4zMjAzIFoiIC8+CiAgICA8L2NsaXBQYXRoPgogICAgPGxpbmVhckdyYWRpZW50Ci"
        LogoBase64 = LogoBase64 & "AgICAgICBpZD0ibGluZWFyR3JhZGllbnQzOTA1IgogICAgICAgc3ByZWFkTWV0aG9kPSJwYWQiCiAgICAgICBncmFkaWVudFRyYW5zZm9ybT0i"
        LogoBase64 = LogoBase64 & "bWF0cml4KDAsLTI1LjM4OTAwOCwtMjUuMzg5MDA4LDAsMjYyLjc1MiwxMDIuNDY1KSIKICAgICAgIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU"
        LogoBase64 = LogoBase64 & "9uVXNlIgogICAgICAgeTI9IjAiCiAgICAgICB4Mj0iMSIKICAgICAgIHkxPSIwIgogICAgICAgeDE9IjAiPgogICAgICA8c3RvcAogICAgICAg"
        LogoBase64 = LogoBase64 & "ICBpZD0ic3RvcDM5MDEiCiAgICAgICAgIG9mZnNldD0iMCIKICAgICAgICAgc3R5bGU9InN0b3Atb3BhY2l0eToxO3N0b3AtY29sb3I6IzE3OW"
        LogoBase64 = LogoBase64 & "RkOSIgLz4KICAgICAgPHN0b3AKICAgICAgICAgaWQ9InN0b3AzOTAzIgogICAgICAgICBvZmZzZXQ9IjEiCiAgICAgICAgIHN0eWxlPSJzdG9w"
        LogoBase64 = LogoBase64 & "LW9wYWNpdHk6MTtzdG9wLWNvbG9yOiMyYjMzN2QiIC8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPGNsaXBQYXRoCiAgICAgICBpZD0iY2"
        LogoBase64 = LogoBase64 & "xpcFBhdGgzODQ3IgogICAgICAgY2xpcFBhdGhVbml0cz0idXNlclNwYWNlT25Vc2UiPgogICAgICA8cGF0aAogICAgICAgICBpbmtzY2FwZTpj"
        LogoBase64 = LogoBase64 & "b25uZWN0b3ItY3VydmF0dXJlPSIwIgogICAgICAgICBpZD0icGF0aDM4NDUiCiAgICAgICAgIGQ9Im0gMjQwLjY5NiwxMDIuMjI4IGggMTkuMT"
        LogoBase64 = LogoBase64 & "c2IFYgNzcuMTQ0IGggLTE5LjE3NiB6IiAvPgogICAgPC9jbGlwUGF0aD4KICAgIDxjbGlwUGF0aAogICAgICAgaWQ9ImNsaXBQYXRoMzg1NSIK"
        LogoBase64 = LogoBase64 & "ICAgICAgIGNsaXBQYXRoVW5pdHM9InVzZXJTcGFjZU9uVXNlIj4KICAgICAgPHBhdGgKICAgICAgICAgaW5rc2NhcGU6Y29ubmVjdG9yLWN1cn"
        LogoBase64 = LogoBase64 & "ZhdHVyZT0iMCIKICAgICAgICAgaWQ9InBhdGgzODUzIgogICAgICAgICBkPSJtIDI0MC42OTYsMTAyLjIyOSAxNC4wODMsLTI1LjA4NSA1LjA5"
        LogoBase64 = LogoBase64 & "Myw5LjQ5NSAtOC42MzgsMTUuNTkiIC8+CiAgICA8L2NsaXBQYXRoPgogICAgPGNsaXBQYXRoCiAgICAgICBpZD0iY2xpcFBhdGgzODY1IgogIC"
        LogoBase64 = LogoBase64 & "AgICAgY2xpcFBhdGhVbml0cz0idXNlclNwYWNlT25Vc2UiPgogICAgICA8cGF0aAogICAgICAgICBpbmtzY2FwZTpjb25uZWN0b3ItY3VydmF0"
        LogoBase64 = LogoBase64 & "dXJlPSIwIgogICAgICAgICBpZD0icGF0aDM4NjMiCiAgICAgICAgIGQ9Ik0gMjU5Ljg3MTg1LDEwMi4yMjgyOSBWIDc3LjE0Mzk2IGggLTE5Lj"
        LogoBase64 = LogoBase64 & "E3NTcgdiAyNS4wODQzMyB6IiAvPgogICAgPC9jbGlwUGF0aD4KICAgIDxsaW5lYXJHcmFkaWVudAogICAgICAgaWQ9ImxpbmVhckdyYWRpZW50"
        LogoBase64 = LogoBase64 & "Mzg3MSIKICAgICAgIHNwcmVhZE1ldGhvZD0icGFkIgogICAgICAgZ3JhZGllbnRUcmFuc2Zvcm09Im1hdHJpeCgwLC0yNS4zODk5OTksLTI1Lj"
        LogoBase64 = LogoBase64 & "M4OTk5OSwwLDI1MC4yODQsMTAyLjQ2NSkiCiAgICAgICBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSIKICAgICAgIHkyPSIwIgogICAg"
        LogoBase64 = LogoBase64 & "ICAgeDI9IjEiCiAgICAgICB5MT0iMCIKICAgICAgIHgxPSIwIj4KICAgICAgPHN0b3AKICAgICAgICAgaWQ9InN0b3AzODY3IgogICAgICAgIC"
        LogoBase64 = LogoBase64 & "BvZmZzZXQ9IjAiCiAgICAgICAgIHN0eWxlPSJzdG9wLW9wYWNpdHk6MTtzdG9wLWNvbG9yOiMxNzlkZDkiIC8+CiAgICAgIDxzdG9wCiAgICAg"
        LogoBase64 = LogoBase64 & "ICAgIGlkPSJzdG9wMzg2OSIKICAgICAgICAgb2Zmc2V0PSIxIgogICAgICAgICBzdHlsZT0ic3RvcC1vcGFjaXR5OjE7c3RvcC1jb2xvcjojMm"
        LogoBase64 = LogoBase64 & "IzMzdkIiAvPgogICAgPC9saW5lYXJHcmFkaWVudD4KICA8L2RlZnM+CiAgPHNvZGlwb2RpOm5hbWVkdmlldwogICAgIGlkPSJiYXNlIgogICAg"
        LogoBase64 = LogoBase64 & "IHBhZ2Vjb2xvcj0iI2ZmZmZmZiIKICAgICBib3JkZXJjb2xvcj0iIzY2NjY2NiIKICAgICBib3JkZXJvcGFjaXR5PSIxLjAiCiAgICAgaW5rc2"
        LogoBase64 = LogoBase64 & "NhcGU6cGFnZW9wYWNpdHk9IjAuMCIKICAgICBpbmtzY2FwZTpwYWdlc2hhZG93PSIyIgogICAgIGlua3NjYXBlOnpvb209IjUuMzA3MTg4NCIK"
        LogoBase64 = LogoBase64 & "ICAgICBpbmtzY2FwZTpjeD0iNzUuNTU3ODk4IgogICAgIGlua3NjYXBlOmN5PSIxMi4zMDUzMzQiCiAgICAgaW5rc2NhcGU6ZG9jdW1lbnQtdW"
        LogoBase64 = LogoBase64 & "5pdHM9Im1tIgogICAgIGlua3NjYXBlOmN1cnJlbnQtbGF5ZXI9ImxheWVyMSIKICAgICBzaG93Z3JpZD0iZmFsc2UiCiAgICAgZml0LW1hcmdp"
        LogoBase64 = LogoBase64 & "bi10b3A9IjAiCiAgICAgZml0LW1hcmdpbi1sZWZ0PSIwIgogICAgIGZpdC1tYXJnaW4tcmlnaHQ9IjAiCiAgICAgZml0LW1hcmdpbi1ib3R0b2"
        LogoBase64 = LogoBase64 & "09IjAiCiAgICAgdW5pdHM9InB4IgogICAgIGlua3NjYXBlOndpbmRvdy13aWR0aD0iMTI4MCIKICAgICBpbmtzY2FwZTp3aW5kb3ctaGVpZ2h0"
        LogoBase64 = LogoBase64 & "PSI3NDQiCiAgICAgaW5rc2NhcGU6d2luZG93LXg9Ii00IgogICAgIGlua3NjYXBlOndpbmRvdy15PSItNCIKICAgICBpbmtzY2FwZTp3aW5kb3"
        LogoBase64 = LogoBase64 & "ctbWF4aW1pemVkPSIxIiAvPgogIDxtZXRhZGF0YQogICAgIGlkPSJtZXRhZGF0YTQ3NTAiPgogICAgPHJkZjpSREY+CiAgICAgIDxjYzpXb3Jr"
        LogoBase64 = LogoBase64 & "CiAgICAgICAgIHJkZjphYm91dD0iIj4KICAgICAgICA8ZGM6Zm9ybWF0PmltYWdlL3N2Zyt4bWw8L2RjOmZvcm1hdD4KICAgICAgICA8ZGM6dH"
        LogoBase64 = LogoBase64 & "lwZQogICAgICAgICAgIHJkZjpyZXNvdXJjZT0iaHR0cDovL3B1cmwub3JnL2RjL2RjbWl0eXBlL1N0aWxsSW1hZ2UiIC8+CiAgICAgICAgPGRj"
        LogoBase64 = LogoBase64 & "OnRpdGxlPjwvZGM6dGl0bGU+CiAgICAgIDwvY2M6V29yaz4KICAgIDwvcmRmOlJERj4KICA8L21ldGFkYXRhPgogIDxnCiAgICAgaW5rc2NhcG"
        LogoBase64 = LogoBase64 & "U6bGFiZWw9IkxheWVyIDEiCiAgICAgaW5rc2NhcGU6Z3JvdXBtb2RlPSJsYXllciIKICAgICBpZD0ibGF5ZXIxIgogICAgIHRyYW5zZm9ybT0i"
        LogoBase64 = LogoBase64 & "dHJhbnNsYXRlKDQzLjQyNTg4OSwtMTQyLjU5MDIzKSI+CiAgICA8ZwogICAgICAgaWQ9Imc1MjA3IgogICAgICAgdHJhbnNmb3JtPSJ0cmFuc2"
        LogoBase64 = LogoBase64 & "xhdGUoMC4xMzIyODk5MSwtMC4xMzIyOTAxMSkiPgogICAgICA8ZwogICAgICAgICB0cmFuc2Zvcm09Im1hdHJpeCgwLjM1Mjc3Nzc3LDAsMCwt"
        LogoBase64 = LogoBase64 & "MC4zNTI3Nzc3NywtMTI4LjMzODE0LDE3OS4zMDM4MSkiCiAgICAgICAgIGlkPSJnMzg0MSI+CiAgICAgICAgPGcKICAgICAgICAgICBjbGlwLX"
        LogoBase64 = LogoBase64 & "BhdGg9InVybCgjY2xpcFBhdGgzODQ3KSIKICAgICAgICAgICBpZD0iZzM4NDMiPgogICAgICAgICAgPGcKICAgICAgICAgICAgIGlkPSJnMzg0"
        LogoBase64 = LogoBase64 & "OSI+CiAgICAgICAgICAgIDxnCiAgICAgICAgICAgICAgIGNsaXAtcGF0aD0idXJsKCNjbGlwUGF0aDM4NTUpIgogICAgICAgICAgICAgICBpZD"
        LogoBase64 = LogoBase64 & "0iZzM4NTEiPgogICAgICAgICAgICAgIDxnCiAgICAgICAgICAgICAgICAgaWQ9ImczODU3Ij4KICAgICAgICAgICAgICAgIDxnCiAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgICAgICBpZD0iZzM4NTkiPgogICAgICAgICAgICAgICAgICA8ZwogICAgICAgICAgICAgICAgICAgICBjbGlwLXBhdGg9InVybCgjY2"
        LogoBase64 = LogoBase64 & "xpcFBhdGgzODY1KSIKICAgICAgICAgICAgICAgICAgICAgaWQ9ImczODYxIj4KICAgICAgICAgICAgICAgICAgICA8cGF0aAogICAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgICAgICAgIGlua3NjYXBlOmNvbm5lY3Rvci1jdXJ2YXR1cmU9IjAiCiAgICAgICAgICAgICAgICAgICAgICAgaWQ9InBhdGgzODczIg"
        LogoBase64 = LogoBase64 & "ogICAgICAgICAgICAgICAgICAgICAgIHN0eWxlPSJmaWxsOnVybCgjbGluZWFyR3JhZGllbnQzODcxKTtzdHJva2U6bm9uZSIKICAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgICAgICAgICBkPSJtIDI0MC42OTYsMTAyLjIyOSAxNC4wODMsLTI1LjA4NSA1LjA5Myw5LjQ5NSAtOC42MzgsMTUuNTkiIC8+CiAgIC"
        LogoBase64 = LogoBase64 & "AgICAgICAgICAgICAgIDwvZz4KICAgICAgICAgICAgICAgIDwvZz4KICAgICAgICAgICAgICA8L2c+CiAgICAgICAgICAgIDwvZz4KICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgIDwvZz4KICAgICAgICA8L2c+CiAgICAgIDwvZz4KICAgICAgPGcKICAgICAgICAgdHJhbnNmb3JtPSJtYXRyaXgoMC4zNTI3Nzc3NywwLD"
        LogoBase64 = LogoBase64 & "AsLTAuMzUyNzc3NzcsLTEyOC4zMzgxNCwxNzkuMzAzODEpIgogICAgICAgICBpZD0iZzM4NzUiPgogICAgICAgIDxnCiAgICAgICAgICAgY2xp"
        LogoBase64 = LogoBase64 & "cC1wYXRoPSJ1cmwoI2NsaXBQYXRoMzg4MSkiCiAgICAgICAgICAgaWQ9ImczODc3Ij4KICAgICAgICAgIDxnCiAgICAgICAgICAgICBpZD0iZz"
        LogoBase64 = LogoBase64 & "M4ODMiPgogICAgICAgICAgICA8ZwogICAgICAgICAgICAgICBjbGlwLXBhdGg9InVybCgjY2xpcFBhdGgzODg5KSIKICAgICAgICAgICAgICAg"
        LogoBase64 = LogoBase64 & "aWQ9ImczODg1Ij4KICAgICAgICAgICAgICA8ZwogICAgICAgICAgICAgICAgIGlkPSJnMzg5MSI+CiAgICAgICAgICAgICAgICA8ZwogICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgICAgICAgICAgaWQ9ImczODkzIj4KICAgICAgICAgICAgICAgICAgPGcKICAgICAgICAgICAgICAgICAgICAgY2xpcC1wYXRoPSJ1cmwo"
        LogoBase64 = LogoBase64 & "I2NsaXBQYXRoMzg5OSkiCiAgICAgICAgICAgICAgICAgICAgIGlkPSJnMzg5NSI+CiAgICAgICAgICAgICAgICAgICAgPHBhdGgKICAgICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgICAgICAgICAgICBpbmtzY2FwZTpjb25uZWN0b3ItY3VydmF0dXJlPSIwIgogICAgICAgICAgICAgICAgICAgICAgIGlkPSJwYXRoMzkw"
        LogoBase64 = LogoBase64 & "NyIKICAgICAgICAgICAgICAgICAgICAgICBzdHlsZT0iZmlsbDp1cmwoI2xpbmVhckdyYWRpZW50MzkwNSk7c3Ryb2tlOm5vbmUiCiAgICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgICAgICAgICAgICAgZD0ibSAyNTYuNjYxLDk3LjIyNyBjIDAsLTMuMzY2IDIuNzM0LC02LjA5NCA2LjA5NiwtNi4wOTQgMy4zNjQsMCA2"
        LogoBase64 = LogoBase64 & "LjA4NywyLjcyOCA2LjA4Nyw2LjA5NCAwLDMuMzY1IC0yLjcyMyw2LjA5MyAtNi4wODcsNi4wOTMgLTMuMzYyLDAgLTYuMDk2LC0yLjcyOCAtNi"
        LogoBase64 = LogoBase64 & "4wOTYsLTYuMDkzIiAvPgogICAgICAgICAgICAgICAgICA8L2c+CiAgICAgICAgICAgICAgICA8L2c+CiAgICAgICAgICAgICAgPC9nPgogICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgICA8L2c+CiAgICAgICAgICA8L2c+CiAgICAgICAgPC9nPgogICAgICA8L2c+CiAgICAgIDxnCiAgICAgICAgIHRyYW5zZm9ybT0ibW"
        LogoBase64 = LogoBase64 & "F0cml4KDAuMzUyNzc3NzcsMCwwLC0wLjM1Mjc3Nzc3LC0xMjguMzM4MTQsMTc5LjMwMzgxKSIKICAgICAgICAgaWQ9ImczOTA5Ij4KICAgICAg"
        LogoBase64 = LogoBase64 & "ICA8ZwogICAgICAgICAgIGNsaXAtcGF0aD0idXJsKCNjbGlwUGF0aDM5MTUpIgogICAgICAgICAgIGlkPSJnMzkxMSI+CiAgICAgICAgICA8Zw"
        LogoBase64 = LogoBase64 & "ogICAgICAgICAgICAgaWQ9ImczOTE3Ij4KICAgICAgICAgICAgPGcKICAgICAgICAgICAgICAgY2xpcC1wYXRoPSJ1cmwoI2NsaXBQYXRoMzky"
        LogoBase64 = LogoBase64 & "MykiCiAgICAgICAgICAgICAgIGlkPSJnMzkxOSI+CiAgICAgICAgICAgICAgPGcKICAgICAgICAgICAgICAgICBpZD0iZzM5MjUiPgogICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgICAgICAgPGcKICAgICAgICAgICAgICAgICAgIGlkPSJnMzkyNyI+CiAgICAgICAgICAgICAgICAgIDxnCiAgICAgICAgICAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgIGNsaXAtcGF0aD0idXJsKCNjbGlwUGF0aDM5MzMpIgogICAgICAgICAgICAgICAgICAgICBpZD0iZzM5MjkiPgogICAgICAgICAgICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgIDxwYXRoCiAgICAgICAgICAgICAgICAgICAgICAgaW5rc2NhcGU6Y29ubmVjdG9yLWN1cnZhdHVyZT0iMCIKICAgICAgICAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgICBpZD0icGF0aDM5NDEiCiAgICAgICAgICAgICAgICAgICAgICAgc3R5bGU9ImZpbGw6dXJsKCNsaW5lYXJHcmFkaWVudDM5MzkpO3"
        LogoBase64 = LogoBase64 & "N0cm9rZTpub25lIgogICAgICAgICAgICAgICAgICAgICAgIGQ9Im0gMjgwLjIyLDk3LjE2MyAtMi4yMDcsLTcuMTQ5IGggLTAuMDQgbCAtMi4y"
        LogoBase64 = LogoBase64 & "MDgsNy4xNDkgaCAtMy4wMTggbCAzLjU4NSwtMTAuNDY5IGggMy4xOTkgbCAzLjU0NSwxMC40NjkiIC8+CiAgICAgICAgICAgICAgICAgIDwvZz"
        LogoBase64 = LogoBase64 & "4KICAgICAgICAgICAgICAgIDwvZz4KICAgICAgICAgICAgICA8L2c+CiAgICAgICAgICAgIDwvZz4KICAgICAgICAgIDwvZz4KICAgICAgICA8"
        LogoBase64 = LogoBase64 & "L2c+CiAgICAgIDwvZz4KICAgICAgPGcKICAgICAgICAgdHJhbnNmb3JtPSJtYXRyaXgoMC4zNTI3Nzc3NywwLDAsLTAuMzUyNzc3NzcsLTEyOC"
        LogoBase64 = LogoBase64 & "4zMzgxNCwxNzkuMzAzODEpIgogICAgICAgICBpZD0iZzM5NDMiPgogICAgICAgIDxnCiAgICAgICAgICAgY2xpcC1wYXRoPSJ1cmwoI2NsaXBQ"
        LogoBase64 = LogoBase64 & "YXRoMzk0OSkiCiAgICAgICAgICAgaWQ9ImczOTQ1Ij4KICAgICAgICAgIDxnCiAgICAgICAgICAgICBpZD0iZzM5NTEiPgogICAgICAgICAgIC"
        LogoBase64 = LogoBase64 & "A8ZwogICAgICAgICAgICAgICBjbGlwLXBhdGg9InVybCgjY2xpcFBhdGgzOTU3KSIKICAgICAgICAgICAgICAgaWQ9ImczOTUzIj4KICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgICA8ZwogICAgICAgICAgICAgICAgIGlkPSJnMzk1OSI+CiAgICAgICAgICAgICAgICA8ZwogICAgICAgICAgICAgICAgICAgaWQ9Im"
        LogoBase64 = LogoBase64 & "czOTYxIj4KICAgICAgICAgICAgICAgICAgPGcKICAgICAgICAgICAgICAgICAgICAgY2xpcC1wYXRoPSJ1cmwoI2NsaXBQYXRoMzk2NykiCiAg"
        LogoBase64 = LogoBase64 & "ICAgICAgICAgICAgICAgICAgIGlkPSJnMzk2MyI+CiAgICAgICAgICAgICAgICAgICAgPHBhdGgKICAgICAgICAgICAgICAgICAgICAgICBpbm"
        LogoBase64 = LogoBase64 & "tzY2FwZTpjb25uZWN0b3ItY3VydmF0dXJlPSIwIgogICAgICAgICAgICAgICAgICAgICAgIGlkPSJwYXRoMzk3NSIKICAgICAgICAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgICBzdHlsZT0iZmlsbDp1cmwoI2xpbmVhckdyYWRpZW50Mzk3Myk7c3Ryb2tlOm5vbmUiCiAgICAgICAgICAgICAgICAgICAgICAgZD"
        LogoBase64 = LogoBase64 & "0ibSAyODYuNjMsOTcuMjc1IGMgLTAuNTc0LC0wLjExNiAtMS4wOTEsLTAuMzA4IC0xLjU0OSwtMC41NzkgLTAuNDYsLTAuMjcgLTAuODM5LC0w"
        LogoBase64 = LogoBase64 & "LjYyNiAtMS4xMzUsLTEuMDcyIC0wLjI5NiwtMC40NDUgLTAuNDY2LC0xLjAwNiAtMC41MDUsLTEuNjgxIGggMi44NzUgYyAwLjA1MywwLjU2Ny"
        LogoBase64 = LogoBase64 & "AwLjI0NCwwLjk3MyAwLjU2NywxLjIxNiAwLjMyMywwLjI0MiAwLjc2OSwwLjM2NCAxLjMzNywwLjM2NCAwLjI1NSwwIDAuNDk1LC0wLjAxOCAw"
        LogoBase64 = LogoBase64 & "LjcxOCwtMC4wNTEgMC4yMjQsLTAuMDM0IDAuNDE5LC0wLjEwMSAwLjU4NywtMC4yMDIgMC4xNywtMC4xMDIgMC4zMDUsLTAuMjQzIDAuNDA2LC"
        LogoBase64 = LogoBase64 & "0wLjQyNSAwLjEwMSwtMC4xODMgMC4xNTIsLTAuNDMgMC4xNTIsLTAuNzQxIDAuMDEzLC0wLjI5NiAtMC4wNzUsLTAuNTIzIC0wLjI2NCwtMC42"
        LogoBase64 = LogoBase64 & "NzcgLTAuMTg5LC0wLjE1NSAtMC40NDUsLTAuMjc0IC0wLjc3LC0wLjM1NSAtMC4zMjQsLTAuMDggLTAuNjk1LC0wLjE0MiAtMS4xMTMsLTAuMT"
        LogoBase64 = LogoBase64 & "gyIC0wLjQxOCwtMC4wNDEgLTAuODQ0LC0wLjA5NSAtMS4yNzcsLTAuMTYzIC0wLjQzMiwtMC4wNjcgLTAuODU5LC0wLjE1NyAtMS4yODUsLTAu"
        LogoBase64 = LogoBase64 & "MjcyIC0wLjQyNSwtMC4xMTUgLTAuODA0LC0wLjI4OCAtMS4xMzQsLTAuNTE3IC0wLjMzMSwtMC4yMyAtMC42MDEsLTAuNTM2IC0wLjgxLC0wLj"
        LogoBase64 = LogoBase64 & "kyIC0wLjIxLC0wLjM4NiAtMC4zMTQsLTAuODc2IC0wLjMxNCwtMS40NyAwLC0wLjU0IDAuMDkxLC0xLjAwNiAwLjI3MywtMS4zOTcgMC4xODMs"
        LogoBase64 = LogoBase64 & "LTAuMzkyIDAuNDM2LC0wLjcxNiAwLjc2LC0wLjk3MiAwLjMyNCwtMC4yNTcgMC43MDIsLTAuNDQ1IDEuMTM0LC0wLjU2NyAwLjQzMiwtMC4xMj"
        LogoBase64 = LogoBase64 & "EgMC44OTgsLTAuMTgyIDEuMzk3LC0wLjE4MiAwLjY0OCwwIDEuMjgzLDAuMDk0IDEuOTAzLDAuMjgzIDAuNjIyLDAuMTg5IDEuMTYyLDAuNTIg"
        LogoBase64 = LogoBase64 & "MS42MjEsMC45OTIgMC4wMTYsLTAuMjQ4IDAuMTM1LC0wLjgyIDAuMjAzLC0xLjAxMSBoIDIuNjM1IGMgMCwwIC0wLjA4MywxLjIyNSAtMC4wOD"
        LogoBase64 = LogoBase64 & "MsMi4zMjggdiA1LjQ0NyBjIDAsMC42MzYgLTAuMTQxLDEuMTQ0IC0wLjQyNSwxLjUyOSAtMC4yODQsMC4zODUgLTAuNjQ5LDAuNjg1IC0xLjA5"
        LogoBase64 = LogoBase64 & "NSwwLjkwMSAtMC40NDUsMC4yMTcgLTAuOTM4LDAuMzYyIC0xLjQ3OCwwLjQzNiAtMC41NCwwLjA3NCAtMS4wNzQsMC4xMTIgLTEuNTk5LDAuMT"
        LogoBase64 = LogoBase64 & "EyIC0wLjU4MSwwIC0xLjE1OSwtMC4wNTggLTEuNzMyLC0wLjE3MiBtIDAuNDU2LC04Ljg4MSBjIC0wLjIxOCwwLjA0MSAtMC40MDYsMC4xMTEg"
        LogoBase64 = LogoBase64 & "LTAuNTY4LDAuMjEyIC0wLjE2MiwwLjEwMyAtMC4yOTEsMC4yNDEgLTAuMzg0LDAuNDE2IC0wLjA5NSwwLjE3NiAtMC4xNDMsMC4zOTEgLTAuMT"
        LogoBase64 = LogoBase64 & "QzLDAuNjQ2IDAsMC4yNzIgMC4wNDgsMC40OTQgMC4xNDMsMC42NyAwLjA5MywwLjE3NSAwLjIxOSwwLjMyIDAuMzczLDAuNDM2IDAuMTU1LDAu"
        LogoBase64 = LogoBase64 & "MTE1IDAuMzM5LDAuMjA1IDAuNTQ4LDAuMjc0IDAuMjEsMC4wNjYgMC40MjMsMC4xMjEgMC42MzgsMC4xNjEgMC4yMjksMC4wNDEgMC40NTksMC"
        LogoBase64 = LogoBase64 & "4wNzUgMC42ODgsMC4xMDIgMC4yMywwLjAyNyAwLjQ0OCwwLjA2MSAwLjY1OCwwLjEwMSAwLjIxLDAuMDQgMC40MDYsMC4wOTIgMC41ODgsMC4x"
        LogoBase64 = LogoBase64 & "NTIgMC4xODIsMC4wNiAwLjMzNCwwLjE0NCAwLjQ1NiwwLjI1MiB2IC0xLjA3MiBjIDAsLTAuMTYzIC0wLjAxOCwtMC4zNzkgLTAuMDUxLC0wLj"
        LogoBase64 = LogoBase64 & "Y0OCAtMC4wMzQsLTAuMjcxIC0wLjEyNSwtMC41MzggLTAuMjczLC0wLjgwMiAtMC4xNDksLTAuMjYyIC0wLjM3OCwtMC40ODggLTAuNjksLTAu"
        LogoBase64 = LogoBase64 & "Njc3IC0wLjMxLC0wLjE5IC0wLjc0NywtMC4yODQgLTEuMzE1LC0wLjI4NCAtMC4yMywwIC0wLjQ1MiwwLjAyIC0wLjY2OCwwLjA2MSIgLz4KIC"
        LogoBase64 = LogoBase64 & "AgICAgICAgICAgICAgICAgPC9nPgogICAgICAgICAgICAgICAgPC9nPgogICAgICAgICAgICAgIDwvZz4KICAgICAgICAgICAgPC9nPgogICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgPC9nPgogICAgICAgIDwvZz4KICAgICAgPC9nPgogICAgICA8ZwogICAgICAgICB0cmFuc2Zvcm09Im1hdHJpeCgwLjM1Mjc3Nzc3LD"
        LogoBase64 = LogoBase64 & "AsMCwtMC4zNTI3Nzc3NywtMTI4LjMzODE0LDE3OS4zMDM4MSkiCiAgICAgICAgIGlkPSJnMzk3NyI+CiAgICAgICAgPGcKICAgICAgICAgICBj"
        LogoBase64 = LogoBase64 & "bGlwLXBhdGg9InVybCgjY2xpcFBhdGgzOTgzKSIKICAgICAgICAgICBpZD0iZzM5NzkiPgogICAgICAgICAgPGcKICAgICAgICAgICAgIGlkPS"
        LogoBase64 = LogoBase64 & "JnMzk4NSI+CiAgICAgICAgICAgIDxnCiAgICAgICAgICAgICAgIGlkPSJnMzk4NyI+CiAgICAgICAgICAgICAgPGcKICAgICAgICAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICBjbGlwLXBhdGg9InVybCgjY2xpcFBhdGgzOTkzKSIKICAgICAgICAgICAgICAgICBpZD0iZzM5ODkiPgogICAgICAgICAgICAgICAgPHBhdG"
        LogoBase64 = LogoBase64 & "gKICAgICAgICAgICAgICAgICAgIGlua3NjYXBlOmNvbm5lY3Rvci1jdXJ2YXR1cmU9IjAiCiAgICAgICAgICAgICAgICAgICBpZD0icGF0aDQw"
        LogoBase64 = LogoBase64 & "MDEiCiAgICAgICAgICAgICAgICAgICBzdHlsZT0iZmlsbDp1cmwoI2xpbmVhckdyYWRpZW50Mzk5OSk7c3Ryb2tlOm5vbmUiCiAgICAgICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgICAgICBkPSJtIDI5NC4zODQsMTAxLjE1MiBoIDIuODc1IFYgODYuNjkzIGggLTIuODc1IHoiIC8+CiAgICAgICAgICAgICAgPC9nPgog"
        LogoBase64 = LogoBase64 & "ICAgICAgICAgICA8L2c+CiAgICAgICAgICA8L2c+CiAgICAgICAgPC9nPgogICAgICA8L2c+CiAgICAgIDxnCiAgICAgICAgIHRyYW5zZm9ybT"
        LogoBase64 = LogoBase64 & "0ibWF0cml4KDAuMzUyNzc3NzcsMCwwLC0wLjM1Mjc3Nzc3LC0xMjguMzM4MTQsMTc5LjMwMzgxKSIKICAgICAgICAgaWQ9Imc0MDAzIj4KICAg"
        LogoBase64 = LogoBase64 & "ICAgICA8ZwogICAgICAgICAgIGNsaXAtcGF0aD0idXJsKCNjbGlwUGF0aDQwMDkpIgogICAgICAgICAgIGlkPSJnNDAwNSI+CiAgICAgICAgIC"
        LogoBase64 = LogoBase64 & "A8ZwogICAgICAgICAgICAgaWQ9Imc0MDExIj4KICAgICAgICAgICAgPGcKICAgICAgICAgICAgICAgaWQ9Imc0MDEzIj4KICAgICAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICA8ZwogICAgICAgICAgICAgICAgIGNsaXAtcGF0aD0idXJsKCNjbGlwUGF0aDQwMTkpIgogICAgICAgICAgICAgICAgIGlkPSJnNDAxNSI+Ci"
        LogoBase64 = LogoBase64 & "AgICAgICAgICAgICAgICA8cGF0aAogICAgICAgICAgICAgICAgICAgaW5rc2NhcGU6Y29ubmVjdG9yLWN1cnZhdHVyZT0iMCIKICAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgICAgIGlkPSJwYXRoNDAyNyIKICAgICAgICAgICAgICAgICAgIHN0eWxlPSJmaWxsOnVybCgjbGluZWFyR3JhZGllbnQ0MDI1KTtzdH"
        LogoBase64 = LogoBase64 & "Jva2U6bm9uZSIKICAgICAgICAgICAgICAgICAgIGQ9Im0gMjk4LjgwNSwxMDEuMTUyIGggMi44NzcgViA4Ni42OTMgaCAtMi44NzcgeiIgLz4K"
        LogoBase64 = LogoBase64 & "ICAgICAgICAgICAgICA8L2c+CiAgICAgICAgICAgIDwvZz4KICAgICAgICAgIDwvZz4KICAgICAgICA8L2c+CiAgICAgIDwvZz4KICAgICAgPG"
        LogoBase64 = LogoBase64 & "cKICAgICAgICAgdHJhbnNmb3JtPSJtYXRyaXgoMC4zNTI3Nzc3NywwLDAsLTAuMzUyNzc3NzcsLTEyOC4zMzgxNCwxNzkuMzAzODEpIgogICAg"
        LogoBase64 = LogoBase64 & "ICAgICBpZD0iZzQwMjkiPgogICAgICAgIDxnCiAgICAgICAgICAgY2xpcC1wYXRoPSJ1cmwoI2NsaXBQYXRoNDAzNSkiCiAgICAgICAgICAgaW"
        LogoBase64 = LogoBase64 & "Q9Imc0MDMxIj4KICAgICAgICAgIDxnCiAgICAgICAgICAgICBpZD0iZzQwMzciPgogICAgICAgICAgICA8ZwogICAgICAgICAgICAgICBjbGlw"
        LogoBase64 = LogoBase64 & "LXBhdGg9InVybCgjY2xpcFBhdGg0MDQzKSIKICAgICAgICAgICAgICAgaWQ9Imc0MDM5Ij4KICAgICAgICAgICAgICA8ZwogICAgICAgICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgIGlkPSJnNDA0NSI+CiAgICAgICAgICAgICAgICA8ZwogICAgICAgICAgICAgICAgICAgaWQ9Imc0MDQ3Ij4KICAgICAgICAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgPGcKICAgICAgICAgICAgICAgICAgICAgY2xpcC1wYXRoPSJ1cmwoI2NsaXBQYXRoNDA1MykiCiAgICAgICAgICAgICAgICAgICAgIGlkPS"
        LogoBase64 = LogoBase64 & "JnNDA0OSI+CiAgICAgICAgICAgICAgICAgICAgPHBhdGgKICAgICAgICAgICAgICAgICAgICAgICBpbmtzY2FwZTpjb25uZWN0b3ItY3VydmF0"
        LogoBase64 = LogoBase64 & "dXJlPSIwIgogICAgICAgICAgICAgICAgICAgICAgIGlkPSJwYXRoNDA2MSIKICAgICAgICAgICAgICAgICAgICAgICBzdHlsZT0iZmlsbDp1cm"
        LogoBase64 = LogoBase64 & "woI2xpbmVhckdyYWRpZW50NDA1OSk7c3Ryb2tlOm5vbmUiCiAgICAgICAgICAgICAgICAgICAgICAgZD0ibSAzMDYuMTI1LDk3LjA1MSBjIC0w"
        LogoBase64 = LogoBase64 & "LjY2MiwtMC4yNjIgLTEuMjMxLC0wLjYzNSAtMS43MDEsLTEuMTEzIC0wLjQ3NCwtMC40OCAtMC44MzgsLTEuMDYgLTEuMDk1LC0xLjc0MiAtMC"
        LogoBase64 = LogoBase64 & "4yNTYsLTAuNjgxIC0wLjM4NCwtMS40NDEgLTAuMzg0LC0yLjI3OSAwLC0wLjgzNyAwLjEyOCwtMS41OTIgMC4zODQsLTIuMjY3IDAuMjU3LC0w"
        LogoBase64 = LogoBase64 & "LjY3NiAwLjYyMSwtMS4yNTMgMS4wOTUsLTEuNzMyIDAuNDcsLTAuNDc5IDEuMDM5LC0wLjg0OCAxLjcwMSwtMS4xMDQgMC42NjEsLTAuMjU2ID"
        LogoBase64 = LogoBase64 & "EuNDAyLC0wLjM4NCAyLjIyNiwtMC4zODQgMC44MjQsMCAxLjU3LDAuMTI4IDIuMjM4LDAuMzg0IDAuNjY5LDAuMjU2IDEuMjM5LDAuNjI1IDEu"
        LogoBase64 = LogoBase64 & "NzExLDEuMTA0IDAuNDczLDAuNDc5IDAuODM4LDEuMDU2IDEuMDk0LDEuNzMyIDAuMjU3LDAuNjc1IDAuMzg1LDEuNDMgMC4zODUsMi4yNjcgMC"
        LogoBase64 = LogoBase64 & "wwLjgzOCAtMC4xMjgsMS41OTggLTAuMzg1LDIuMjc5IC0wLjI1NiwwLjY4MiAtMC42MjEsMS4yNjIgLTEuMDk0LDEuNzQyIC0wLjQ3MiwwLjQ3"
        LogoBase64 = LogoBase64 & "OCAtMS4wNDIsMC44NTEgLTEuNzExLDEuMTEzIC0wLjY2OCwwLjI2NCAtMS40MTQsMC4zOTYgLTIuMjM4LDAuMzk2IC0wLjgyNCwwIC0xLjU2NS"
        LogoBase64 = LogoBase64 & "wtMC4xMzIgLTIuMjI2LC0wLjM5NiBtIDEuMDEyLC04LjE5MSBjIC0wLjMyNCwwLjE4OSAtMC41ODQsMC40NDIgLTAuNzgsMC43NTkgLTAuMTk2"
        LogoBase64 = LogoBase64 & "LDAuMzE3IC0wLjMzMywwLjY3NSAtMC40MTUsMS4wNzUgLTAuMDgxLDAuMzk3IC0wLjEyMSwwLjgwNiAtMC4xMjEsMS4yMjMgMCwwLjQxOSAwLj"
        LogoBase64 = LogoBase64 & "A0LDAuODMxIDAuMTIxLDEuMjM2IDAuMDgyLDAuNDA2IDAuMjE5LDAuNzY0IDAuNDE1LDEuMDczIDAuMTk2LDAuMzExIDAuNDU2LDAuNTY0IDAu"
        LogoBase64 = LogoBase64 & "NzgsMC43NiAwLjMyNCwwLjE5NiAwLjcyOCwwLjI5NCAxLjIxNCwwLjI5NCAwLjQ4NywwIDAuODk1LC0wLjA5OCAxLjIyNSwtMC4yOTQgMC4zMz"
        LogoBase64 = LogoBase64 & "IsLTAuMTk2IDAuNTk1LC0wLjQ0OSAwLjc5LC0wLjc2IDAuMTk3LC0wLjMwOSAwLjMzNSwtMC42NjcgMC40MTYsLTEuMDczIDAuMDgxLC0wLjQw"
        LogoBase64 = LogoBase64 & "NSAwLjEyMiwtMC44MTcgMC4xMjIsLTEuMjM2IDAsLTAuNDE3IC0wLjA0MSwtMC44MjYgLTAuMTIyLC0xLjIyMyAtMC4wODEsLTAuNCAtMC4yMT"
        LogoBase64 = LogoBase64 & "ksLTAuNzU4IC0wLjQxNiwtMS4wNzUgLTAuMTk1LC0wLjMxNyAtMC40NTgsLTAuNTcgLTAuNzksLTAuNzU5IC0wLjMzLC0wLjE4OSAtMC43Mzgs"
        LogoBase64 = LogoBase64 & "LTAuMjg0IC0xLjIyNSwtMC4yODQgLTAuNDg2LDAgLTAuODksMC4wOTUgLTEuMjE0LDAuMjg0IiAvPgogICAgICAgICAgICAgICAgICA8L2c+Ci"
        LogoBase64 = LogoBase64 & "AgICAgICAgICAgICAgICA8L2c+CiAgICAgICAgICAgICAgPC9nPgogICAgICAgICAgICA8L2c+CiAgICAgICAgICA8L2c+CiAgICAgICAgPC9n"
        LogoBase64 = LogoBase64 & "PgogICAgICA8L2c+CiAgICAgIDxnCiAgICAgICAgIHRyYW5zZm9ybT0ibWF0cml4KDAuMzUyNzc3NzcsMCwwLC0wLjM1Mjc3Nzc3LC0xMjguMz"
        LogoBase64 = LogoBase64 & "M4MTQsMTc5LjMwMzgxKSIKICAgICAgICAgaWQ9Imc0MDYzIj4KICAgICAgICA8ZwogICAgICAgICAgIGNsaXAtcGF0aD0idXJsKCNjbGlwUGF0"
        LogoBase64 = LogoBase64 & "aDQwNjkpIgogICAgICAgICAgIGlkPSJnNDA2NSI+CiAgICAgICAgICA8ZwogICAgICAgICAgICAgaWQ9Imc0MDcxIj4KICAgICAgICAgICAgPG"
        LogoBase64 = LogoBase64 & "cKICAgICAgICAgICAgICAgY2xpcC1wYXRoPSJ1cmwoI2NsaXBQYXRoNDA3NykiCiAgICAgICAgICAgICAgIGlkPSJnNDA3MyI+CiAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgPGcKICAgICAgICAgICAgICAgICBpZD0iZzQwNzkiPgogICAgICAgICAgICAgICAgPGcKICAgICAgICAgICAgICAgICAgIGlkPSJnND"
        LogoBase64 = LogoBase64 & "A4MSI+CiAgICAgICAgICAgICAgICAgIDxnCiAgICAgICAgICAgICAgICAgICAgIGNsaXAtcGF0aD0idXJsKCNjbGlwUGF0aDQwODcpIgogICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgICAgICAgICAgICBpZD0iZzQwODMiPgogICAgICAgICAgICAgICAgICAgIDxwYXRoCiAgICAgICAgICAgICAgICAgICAgICAgaW5rc2"
        LogoBase64 = LogoBase64 & "NhcGU6Y29ubmVjdG9yLWN1cnZhdHVyZT0iMCIKICAgICAgICAgICAgICAgICAgICAgICBpZD0icGF0aDQwOTUiCiAgICAgICAgICAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgc3R5bGU9ImZpbGw6dXJsKCNsaW5lYXJHcmFkaWVudDQwOTMpO3N0cm9rZTpub25lIgogICAgICAgICAgICAgICAgICAgICAgIGQ9Im"
        LogoBase64 = LogoBase64 & "0gMzIxLjc5Myw5Ny4xNjMgdiAtNS40ODkgYyAwLC0xLjA2NiAtMC4xNzYsLTEuODMzIC0wLjUyNiwtMi4yOTggLTAuMzUxLC0wLjQ2NSAtMC45"
        LogoBase64 = LogoBase64 & "MTksLTAuNjk5IC0xLjcwMSwtMC42OTkgLTAuNjksMCAtMS4xNjgsMC4yMTMgLTEuNDM4LDAuNjM4IC0wLjI3MSwwLjQyNSAtMC40MDUsMS4wNy"
        LogoBase64 = LogoBase64 & "AtMC40MDUsMS45MzUgdiA1LjkxMyBoIC0yLjg3NiB2IC02LjQ0MSBjIDAsLTAuNjQ4IDAuMDU3LC0xLjIzOSAwLjE3MiwtMS43NzEgMC4xMTQs"
        LogoBase64 = LogoBase64 & "LTAuNTM0IDAuMzE1LC0wLjk4NSAwLjU5NywtMS4zNTcgMC4yODQsLTAuMzcxIDAuNjcyLC0wLjY1OCAxLjE2NCwtMC44NiAwLjQ5MywtMC4yMD"
        LogoBase64 = LogoBase64 & "IgMS4xMjUsLTAuMzA0IDEuODk1LC0wLjMwNCAwLjYwNiwwIDEuMjAxLDAuMTM1IDEuNzgyLDAuNDA1IDAuNTgxLDAuMjY5IDEuMDUzLDAuNzA4"
        LogoBase64 = LogoBase64 & "IDEuNDE3LDEuMzE2IGggMC4wNiB2IC0xLjQ1NyBoIDIuNzM1IHYgMTAuNDY5IiAvPgogICAgICAgICAgICAgICAgICA8L2c+CiAgICAgICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgICA8L2c+CiAgICAgICAgICAgICAgPC9nPgogICAgICAgICAgICA8L2c+CiAgICAgICAgICA8L2c+CiAgICAgICAgPC9nPgogICAgICA8"
        LogoBase64 = LogoBase64 & "L2c+CiAgICAgIDxnCiAgICAgICAgIHRyYW5zZm9ybT0ibWF0cml4KDAuMzUyNzc3NzcsMCwwLC0wLjM1Mjc3Nzc3LC0xMjguMzM4MTQsMTc5Lj"
        LogoBase64 = LogoBase64 & "MwMzgxKSIKICAgICAgICAgaWQ9Imc0MDk3Ij4KICAgICAgICA8ZwogICAgICAgICAgIGNsaXAtcGF0aD0idXJsKCNjbGlwUGF0aDQxMDMpIgog"
        LogoBase64 = LogoBase64 & "ICAgICAgICAgIGlkPSJnNDA5OSI+CiAgICAgICAgICA8ZwogICAgICAgICAgICAgaWQ9Imc0MTA1Ij4KICAgICAgICAgICAgPGcKICAgICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgICAgY2xpcC1wYXRoPSJ1cmwoI2NsaXBQYXRoNDExMSkiCiAgICAgICAgICAgICAgIGlkPSJnNDEwNyI+CiAgICAgICAgICAgICAgPGcK"
        LogoBase64 = LogoBase64 & "ICAgICAgICAgICAgICAgICBpZD0iZzQxMTMiPgogICAgICAgICAgICAgICAgPGcKICAgICAgICAgICAgICAgICAgIGlkPSJnNDExNSI+CiAgIC"
        LogoBase64 = LogoBase64 & "AgICAgICAgICAgICAgIDxnCiAgICAgICAgICAgICAgICAgICAgIGNsaXAtcGF0aD0idXJsKCNjbGlwUGF0aDQxMjEpIgogICAgICAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgICBpZD0iZzQxMTciPgogICAgICAgICAgICAgICAgICAgIDxwYXRoCiAgICAgICAgICAgICAgICAgICAgICAgaW5rc2NhcGU6Y29ubm"
        LogoBase64 = LogoBase64 & "VjdG9yLWN1cnZhdHVyZT0iMCIKICAgICAgICAgICAgICAgICAgICAgICBpZD0icGF0aDQxMjkiCiAgICAgICAgICAgICAgICAgICAgICAgc3R5"
        LogoBase64 = LogoBase64 & "bGU9ImZpbGw6dXJsKCNsaW5lYXJHcmFkaWVudDQxMjcpO3N0cm9rZTpub25lIgogICAgICAgICAgICAgICAgICAgICAgIGQ9Im0gMzM1Ljc3Ny"
        LogoBase64 = LogoBase64 & "w5Ny4wMjEgYyAtMC42NTUsLTAuMjg0IC0xLjIxOSwtMC42NzIgLTEuNjkxLC0xLjE2NSAtMC40NzMsLTAuNDkyIC0wLjgzOCwtMS4wNzYgLTEu"
        LogoBase64 = LogoBase64 & "MDk1LC0xLjc1MiAtMC4yNTUsLTAuNjc1IC0wLjM4MywtMS40MDMgLTAuMzgzLC0yLjE4NyAwLC0wLjgwOSAwLjEyNSwtMS41NTIgMC4zNzQsLT"
        LogoBase64 = LogoBase64 & "IuMjI3IDAuMjUsLTAuNjc1IDAuNjA0LC0xLjI1NSAxLjA2MywtMS43NDIgMC40NTksLTAuNDg1IDEuMDIsLTAuODYgMS42ODEsLTEuMTI0IDAu"
        LogoBase64 = LogoBase64 & "NjYxLC0wLjI2MiAxLjQwNSwtMC4zOTQgMi4yMjgsLTAuMzk0IDEuMTg4LDAgMi4yMDEsMC4yNyAzLjAzNywwLjgwOSAwLjgzOCwwLjU0IDEuND"
        LogoBase64 = LogoBase64 & "U4LDEuNDM4IDEuODYzLDIuNjk0IGggLTIuNTMxIGMgLTAuMDk0LC0wLjMyNSAtMC4zNTIsLTAuNjMxIC0wLjc2OCwtMC45MiAtMC40MiwtMC4y"
        LogoBase64 = LogoBase64 & "OTIgLTAuOTE5LC0wLjQzNyAtMS41MDEsLTAuNDM3IC0wLjgwOSwwIC0xLjQzLDAuMjEgLTEuODYyLDAuNjI4IC0wLjQzMywwLjQxOCAtMC42Nj"
        LogoBase64 = LogoBase64 & "ksMS4wOTMgLTAuNzA5LDIuMDI2IGggNy41NTMgYyAwLjA1NSwwLjgwOSAtMC4wMTIsMS41ODYgLTAuMjAyLDIuMzI5IC0wLjE4OSwwLjc0MSAt"
        LogoBase64 = LogoBase64 & "MC40OTYsMS40MDMgLTAuOTIsMS45ODQgLTAuNDI2LDAuNTgxIC0wLjk3LDEuMDQyIC0xLjYzMSwxLjM4NyAtMC42NjIsMC4zNDUgLTEuNDM5LD"
        LogoBase64 = LogoBase64 & "AuNTE3IC0yLjMyOSwwLjUxNyAtMC43OTYsMCAtMS41MjMsLTAuMTQyIC0yLjE3NywtMC40MjYgbSAtMC4xNjMsLTMuMjggYyAwLjA3NCwwLjI1"
        LogoBase64 = LogoBase64 & "NSAwLjIwMywwLjQ5OSAwLjM4NSwwLjcyOCAwLjE4MywwLjIzIDAuNDI2LDAuNDIyIDAuNzI5LDAuNTc4IDAuMzA1LDAuMTU1IDAuNjg2LDAuMj"
        LogoBase64 = LogoBase64 & "MzIDEuMTQ1LDAuMjMzIDAuNzAxLDAgMS4yMjUsLTAuMTkgMS41NywtMC41NjcgMC4zNDUsLTAuMzc5IDAuNTgzLC0wLjkzMiAwLjcxOCwtMS42"
        LogoBase64 = LogoBase64 & "NiBoIC00LjY3OCBjIDAuMDEyLDAuMjAyIDAuMDU4LDAuNDMyIDAuMTMxLDAuNjg4IiAvPgogICAgICAgICAgICAgICAgICA8L2c+CiAgICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgICAgICA8L2c+CiAgICAgICAgICAgICAgPC9nPgogICAgICAgICAgICA8L2c+CiAgICAgICAgICA8L2c+CiAgICAgICAgPC9nPgogICAg"
        LogoBase64 = LogoBase64 & "ICA8L2c+CiAgICAgIDxnCiAgICAgICAgIHRyYW5zZm9ybT0ibWF0cml4KDAuMzUyNzc3NzcsMCwwLC0wLjM1Mjc3Nzc3LC0xMjguMzM4MTQsMT"
        LogoBase64 = LogoBase64 & "c5LjMwMzgxKSIKICAgICAgICAgaWQ9Imc0MTMxIj4KICAgICAgICA8ZwogICAgICAgICAgIGNsaXAtcGF0aD0idXJsKCNjbGlwUGF0aDQxMzcp"
        LogoBase64 = LogoBase64 & "IgogICAgICAgICAgIGlkPSJnNDEzMyI+CiAgICAgICAgICA8ZwogICAgICAgICAgICAgaWQ9Imc0MTM5Ij4KICAgICAgICAgICAgPGcKICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgICAgICAgY2xpcC1wYXRoPSJ1cmwoI2NsaXBQYXRoNDE0NSkiCiAgICAgICAgICAgICAgIGlkPSJnNDE0MSI+CiAgICAgICAgICAgICAg"
        LogoBase64 = LogoBase64 & "PGcKICAgICAgICAgICAgICAgICBpZD0iZzQxNDciPgogICAgICAgICAgICAgICAgPGcKICAgICAgICAgICAgICAgICAgIGlkPSJnNDE0OSI+Ci"
        LogoBase64 = LogoBase64 & "AgICAgICAgICAgICAgICAgIDxnCiAgICAgICAgICAgICAgICAgICAgIGNsaXAtcGF0aD0idXJsKCNjbGlwUGF0aDQxNTUpIgogICAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgICAgICBpZD0iZzQxNTEiPgogICAgICAgICAgICAgICAgICAgIDxwYXRoCiAgICAgICAgICAgICAgICAgICAgICAgaW5rc2NhcGU6Y2"
        LogoBase64 = LogoBase64 & "9ubmVjdG9yLWN1cnZhdHVyZT0iMCIKICAgICAgICAgICAgICAgICAgICAgICBpZD0icGF0aDQxNjMiCiAgICAgICAgICAgICAgICAgICAgICAg"
        LogoBase64 = LogoBase64 & "c3R5bGU9ImZpbGw6dXJsKCNsaW5lYXJHcmFkaWVudDQxNjEpO3N0cm9rZTpub25lIgogICAgICAgICAgICAgICAgICAgICAgIGQ9Im0gMzQ2Lj"
        LogoBase64 = LogoBase64 & "c4Myw5Ny4wMjEgYyAtMC42NjEsLTAuMjg0IC0xLjIyMiwtMC42NzkgLTEuNjgsLTEuMTg0IC0wLjQ2LC0wLjUwNyAtMC44MDcsLTEuMTA4IC0x"
        LogoBase64 = LogoBase64 & "LjA0MywtMS44MDMgLTAuMjM3LC0wLjY5NSAtMC4zNTUsLTEuNDQ4IC0wLjM1NSwtMi4yNTggMCwtMC43ODIgMC4xMjksLTEuNTAyIDAuMzg1LC"
        LogoBase64 = LogoBase64 & "0yLjE1NyAwLjI1NiwtMC42NTUgMC42MTMsLTEuMjE5IDEuMDczLC0xLjY5MSAwLjQ1OSwtMC40NzMgMS4wMTYsLTAuODQgMS42NywtMS4xMDQg"
        LogoBase64 = LogoBase64 & "MC42NTUsLTAuMjYyIDEuMzc0LC0wLjM5NCAyLjE1OCwtMC4zOTQgMS4zOSwwIDIuNTMxLDAuMzY0IDMuNDIyLDEuMDkzIDAuODkyLDAuNzI5ID"
        LogoBase64 = LogoBase64 & "EuNDMsMS43ODkgMS42MiwzLjE3OSBoIC0yLjc3NCBjIC0wLjA5NCwtMC42NDggLTAuMzI4LC0xLjE2NCAtMC42OTksLTEuNTQ4IC0wLjM3MSwt"
        LogoBase64 = LogoBase64 & "MC4zODUgLTAuOTAxLC0wLjU3OCAtMS41ODksLTAuNTc4IC0wLjQ0NiwwIC0wLjgyNCwwLjEwMSAtMS4xMzUsMC4zMDUgLTAuMzEsMC4yMDIgLT"
        LogoBase64 = LogoBase64 & "AuNTU4LDAuNDYxIC0wLjczOSwwLjc3OCAtMC4xODMsMC4zMTggLTAuMzE0LDAuNjcyIC0wLjM5NSwxLjA2MyAtMC4wOCwwLjM5MiAtMC4xMjEs"
        LogoBase64 = LogoBase64 & "MC43NzggLTAuMTIxLDEuMTU1IDAsMC4zOTEgMC4wNDEsMC43ODYgMC4xMjEsMS4xODUgMC4wODEsMC4zOTkgMC4yMTgsMC43NjMgMC40MTUsMS"
        LogoBase64 = LogoBase64 & "4wOTMgMC4xOTYsMC4zMzEgMC40NDksMC42MDEgMC43NTksMC44MSAwLjMxMSwwLjIxIDAuNjk2LDAuMzE1IDEuMTU1LDAuMzE1IDEuMjI4LDAg"
        LogoBase64 = LogoBase64 & "MS45MzcsLTAuNjAxIDIuMTI3LC0xLjgwMiBoIDIuODE0IGMgLTAuMDQxLDAuNjc0IC0wLjIwMiwxLjI1OCAtMC40ODUsMS43NTEgLTAuMjgzLD"
        LogoBase64 = LogoBase64 & "AuNDkzIC0wLjY1MSwwLjkwNSAtMS4xMDUsMS4yMzYgLTAuNDUyLDAuMzI5IC0wLjk2NCwwLjU3NiAtMS41MzksMC43MzkgLTAuNTczLDAuMTYx"
        LogoBase64 = LogoBase64 & "IC0xLjE3MSwwLjI0MyAtMS43OTIsMC4yNDMgLTAuODUsMCAtMS42MDcsLTAuMTQyIC0yLjI2OCwtMC40MjYiIC8+CiAgICAgICAgICAgICAgIC"
        LogoBase64 = LogoBase64 & "AgIDwvZz4KICAgICAgICAgICAgICAgIDwvZz4KICAgICAgICAgICAgICA8L2c+CiAgICAgICAgICAgIDwvZz4KICAgICAgICAgIDwvZz4KICAg"
        LogoBase64 = LogoBase64 & "ICAgICA8L2c+CiAgICAgIDwvZz4KICAgICAgPGcKICAgICAgICAgdHJhbnNmb3JtPSJtYXRyaXgoMC4zNTI3Nzc3NywwLDAsLTAuMzUyNzc3Nz"
        LogoBase64 = LogoBase64 & "csLTEyOC4zMzgxNCwxNzkuMzAzODEpIgogICAgICAgICBpZD0iZzQxNjUiPgogICAgICAgIDxnCiAgICAgICAgICAgY2xpcC1wYXRoPSJ1cmwo"
        LogoBase64 = LogoBase64 & "I2NsaXBQYXRoNDE3MSkiCiAgICAgICAgICAgaWQ9Imc0MTY3Ij4KICAgICAgICAgIDxnCiAgICAgICAgICAgICBpZD0iZzQxNzMiPgogICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgICA8ZwogICAgICAgICAgICAgICBjbGlwLXBhdGg9InVybCgjY2xpcFBhdGg0MTc5KSIKICAgICAgICAgICAgICAgaWQ9Imc0MTc1Ij4K"
        LogoBase64 = LogoBase64 & "ICAgICAgICAgICAgICA8ZwogICAgICAgICAgICAgICAgIGlkPSJnNDE4MSI+CiAgICAgICAgICAgICAgICA8ZwogICAgICAgICAgICAgICAgIC"
        LogoBase64 = LogoBase64 & "AgaWQ9Imc0MTgzIj4KICAgICAgICAgICAgICAgICAgPGcKICAgICAgICAgICAgICAgICAgICAgY2xpcC1wYXRoPSJ1cmwoI2NsaXBQYXRoNDE4"
        LogoBase64 = LogoBase64 & "OSkiCiAgICAgICAgICAgICAgICAgICAgIGlkPSJnNDE4NSI+CiAgICAgICAgICAgICAgICAgICAgPHBhdGgKICAgICAgICAgICAgICAgICAgIC"
        LogoBase64 = LogoBase64 & "AgICBpbmtzY2FwZTpjb25uZWN0b3ItY3VydmF0dXJlPSIwIgogICAgICAgICAgICAgICAgICAgICAgIGlkPSJwYXRoNDE5NyIKICAgICAgICAg"
        LogoBase64 = LogoBase64 & "ICAgICAgICAgICAgICBzdHlsZT0iZmlsbDp1cmwoI2xpbmVhckdyYWRpZW50NDE5NSk7c3Ryb2tlOm5vbmUiCiAgICAgICAgICAgICAgICAgIC"
        LogoBase64 = LogoBase64 & "AgICAgZD0ibSAzMjkuMDgzLDk1LjQ3NiBoIC0wLjA0IHYgMS42ODcgaCAtMi43MzQgViA4Ni42OTQgaCAyLjg3NSB2IDQuNzE4IGMgMCwwLjQ3"
        LogoBase64 = LogoBase64 & "MSAwLjA0OCwwLjkxMSAwLjE0MywxLjMxNSAwLjA5NCwwLjQwNiAwLjI1MywwLjc2MSAwLjQ3NSwxLjA2NSAwLjIyNCwwLjMwMiAwLjUxNywwLj"
        LogoBase64 = LogoBase64 & "U0MiAwLjg4LDAuNzE4IDAuMzY2LDAuMTc1IDAuODExLDAuMjY0IDEuMzM4LDAuMjY0IDAuMTE2LDAgMC4yNDEsLTAuMDEyIDAuMzYzLC0wLjAy"
        LogoBase64 = LogoBase64 & "MSB2IDIuNjkxIGMgLTAuMzQ5LC0wLjAwNSAtMi40NzMsLTAuMjIxIC0zLjMsLTEuOTY4IiAvPgogICAgICAgICAgICAgICAgICA8L2c+CiAgIC"
        LogoBase64 = LogoBase64 & "AgICAgICAgICAgICA8L2c+CiAgICAgICAgICAgICAgPC9nPgogICAgICAgICAgICA8L2c+CiAgICAgICAgICA8L2c+CiAgICAgICAgPC9nPgog"
        LogoBase64 = LogoBase64 & "ICAgICA8L2c+CiAgICA8L2c+CiAgPC9nPgo8L3N2Zz4K"
        ElseIf Cliente = "Rumo" Then
        LogoBase64 = LogoBase64 & "iVBORw0KGgoAAAANSUhEUgAAAHcAAAA/CAYAAADaMleLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDh"
        LogoBase64 = LogoBase64 & "sAAAzWSURBVHhe7VoJVBTHFp2oMV+T/HyTH2PyY062v8QFGEDRaGI0wSUmMfG4McwAoqgoBBfEFYioqCQaF0xQ3HBFXD6KYATjd8MVdw3igkgU"
        LogoBase64 = LogoBase64 & "F1xQNtm8/7zq6Zme7gZmxKDp0++ce8Tp9169qltd/epVaR4+fAgVyoRG/IMK5UAlV8FQyVUwVHIVDJVcBUMlV8FQyVUwVHIVDJVcBUMlV8FQyV"
        LogoBase64 = LogoBase64 & "UwVHIVDJVcBUMlV8FQyVUwVHIVDJVcBUMlV8H4Q8gFAfSP9NnTAFtj40T6u62oXKS6jwM2kVuZWPO8KhG3U50/Emt1q3vOiW1tk4j1q0NZWTnS"
        LogoBase64 = LogoBase64 & "L15B/M4jWLppF1YkpmJr6glcysmtkd+qYDW5JHfuFWDjjsNYnrgXK5NSsSxhD45lXEZ5eQV7nnv3HsKiN+GL4bPgOnAKZq5IQn5RMXtW9KAEa3"
        LogoBase64 = LogoBase64 & "7ZjxVJqcyWfKzffghXbtyR7RRJQWExth86zdohG8KKxL3IvHJTopt+KQcrt+5jz0lvVVIqjp3NMg4ZcDYrB6ELNuLrwNlwHTod3QNmwm/GCmzZ"
        LogoBase64 = LogoBase64 & "c9ykw/siOXUhG5MWxaPn6LlwHToDXfx/wMDJi7F+x2GUlpWb9KuDUC7n5CIu5SD0E6PQ2NUfb3YLgLM+BH3Hz8fa5IMmPbGPR4VN5J7JvApHfQ"
        LogoBase64 = LogoBase64 & "iauPqhSddv8VonX4QvS0BZeTlOns9G+4GT8cLHg6Cx10Hzfk/0HTcft/Pyme3Nu/fwzlej0KSLP7Nt8tkwtOw3AdsPnpbtEMnv128zH9QOsyF0"
        LogoBase64 = LogoBase64 & "9sOGXw9LdBdt2ok3Pg9Ak86c/9e7BWBGTAJ7Nm9tClr0HofnPxoEjYM7NC3doLHToZ5Lf7z1xQgMnR6DAuMkLH5QiunLtqBZ77F44SMfTt+O09"
        LogoBase64 = LogoBase64 & "e08sSbnf3hM3UpcnLvysYtjIn14cZtxCTuRfDPGxA0dy0ilidh2/6TbFI5G0Kh+XcvaBz1eKPbt+gZNBdnLl6p0q8tsInc4+ez8UbXAI48RwM0"
        LogoBase64 = LogoBase64 & "zfth3Pw4ZF29CScKVKuHxsmDDYKmRT/og6NM5F6/nYf6bQdwOmRrp8PfPxuGLXuOyXaGhJasTr7TWTvMhuDgzpY0se7s2G3QOHtA42D07+yBoL"
        LogoBase64 = LogoBase64 & "mxmLVyK17uMJgjiY+NB+lr9ajv0h8jf1yNsrIKTIqOx18/HGDUN0j1HdxR38mAsfPWIr+QmxBysecXFCNsYTya9xmHxq5+bNITGnf2g70uGBN/"
        LogoBase64 = LogoBase64 & "3oC03zJhCF7AjQmhlSea9RmHJONqIvZrK2wi98T5bDTtPoLNNNZZOx1bujxCF3IdFw6EDLkN2/uY9bR6vNbFH4l75TtCkpWTC9dhEaa3hsHJgJ"
        LogoBase64 = LogoBase64 & "VJ+yS6c9cmo46Ll4nAZ1p7obVnKP7WcYh5QtGkJNLEsWr1ePurUfCZsoRNOI3W3axvL6//UkdfpB7LkMTO+nrrHjxDo/FCW29z2xQXgf7W6vF8"
        LogoBase64 = LogoBase64 & "G2/oQxYgLf0Svho5m2vTONne+3o0tu07KfFtK2pGroMefcZG4uVPh3KBGYNjA/if3tBN+NlM7q3aJZdQr01/LlYnA+zcJsB32jIWL/t00CALCK"
        LogoBase64 = LogoBase64 & "vr0h/P0Rvr7MlIbd53PHymLoZ7SBQ3QRxFBDfvi+lLElBaVmaKn4TeZlp+6/MrjdBGCEcDnnU0IGDmSqQcOIV3e4zixo2NjTucDCG48PsN2bGx"
        LogoBase64 = LogoBase64 & "FjUj19kDr3b2Nw6qAfXaeKP78JkYHL4UXf2+R9CcNci9c58jt5bfXD6+Z9t4Y1jEctb+vfxC3LlfiJQDp9m3XzL4ZGuvg+/0GGRfv4V7BUW4m1"
        LogoBase64 = LogoBase64 & "+Io+mXuH4Lfdu5odeYeSyJFJKbfOA0GrbxtvRNyzutAPStb9EPmpb9uD7Z69ikSjl4GoGz16AuxW98Seo4GTAofCnzKTc+1qBm5NLyZxzEpt1H"
        LogoBase64 = LogoBase64 & "4sCpC8jLL2Sz9+79AtwvKDJl0k+EXCcDHA2hpmSJl4qKhzCELECdVjSYAnLt3dDOezJ+y7xqqV9egWlLErhlmtfV6lmekZVzy6SXd78QvcdGmv"
        LogoBase64 = LogoBase64 & "Wor/Y62OsmYtrSBOw+ehYXsq/j5PnfsWrrfvQKmofgnzawJG7fyfN41dXP/J130OPdrwOZDYl4fKxBjcnlO7Fl9zFwWpbC29Y6uc4eqN/WG+FL"
        LogoBase64 = LogoBase64 & "uIxZor8mGc+1G2iOh0hu6YYJP61HhUwsu9LOQtPCzRyHox7v9AjEucvX2HOSjKxreJa+s+RLa0CjTr6IXJuC2/fyUVJaxiYVL7TvpZeAiKWiT3"
        LogoBase64 = LogoBase64 & "lFBf7ZM8j8xtPb66jH+Mg4pi8eH2tQc3KdPFhKX1RcwoIU2/G2T4Lclz4ZgtTj5yT+SejNaUhbHT4eR46M1b/sl9U/ffEKt5yayDXg9W7fsv01"
        LogoBase64 = LogoBase64 & "SUVFBWK27OGWX60eTb8YgYTdx9g2kRP5ceF/J2E7A2Hy1rIfeoyajbz8IklM1qDm5LLZtQ7l5eUSG6FtrSdUTh545dNhuHWX++aL9eOSD3L7Xj"
        LogoBase64 = LogoBase64 & "4eB3e8981o7ExLl9VPv3SV+2YKye1qJpfezBGzVrNYX+o4BIs375KsAFWBhIg0+SfY6+Di+R3SMx9t71tzch3cEbttv0RfbPskyCX/lfmWI5e+"
        LogoBase64 = LogoBase64 & "jcczLktsSKojt7ikFL3GRLJYaYvD24nbrgwkHQdPs3xzHdxZoefwmUybfPF4LOTuPird74lta31ZdvJg1afKfEvItdexBIlPpsT61ZFL5VVXv+"
        LogoBase64 = LogoBase64 & "9ZdYzyj8KiB4xw2irRki2OQeyflm9KoCyybHsdtO7BOCEz4axBzcm11+HkuewqGyeRkOtoJLeSagzJuezraD9wimWWagO5b385slLfcuQ6ezw6"
        LogoBase64 = LogoBase64 & "ucUPStAtYBY+9pmK5AOn0OiTIaw65Td9OfYczWDVr6pykl8PncErVEARVsXs3NjkzrlZdamzMtScXDsdMrK4DopthLZy5FIpjpIOOVuS4xlZ+F"
        LogoBase64 = LogoBase64 & "evMebN/VNM7oOSUvhFxMB78iIs2bSbFXGYbzs3fDZsBm7nFUj8mgBgQNgi07aSb+MZOx38I5Yzu0ptq8DjITeb2w6IbYS2UnINePGTwexkSM6W"
        LogoBase64 = LogoBase64 & "hN6A+lRlEs7mp5Rc2tpQthw4Zw1CozayChbTc/bAX9oNkPXL+y4pLUWHQeFcCZL3r9XjrS9HsjGQs7MGtUTuQ1bcaNheuK+kOqseURt+ldiSUP"
        LogoBase64 = LogoBase64 & "FjyuLN0DTrY27rKSaX5OKVG5gcHc8OBVglSkBUwA8r2dst9s37pxXsVVdalrn6M+3RqYzJi9jGGtQKuQTaFjTu7G+ZDTbrw0p9fH1WKHSE2KzX"
        LogoBase64 = LogoBase64 & "WDYwfxZyKYmiitKsVVuh+cByUlJsdFBP4yD2z7dBtW+K+8UOg+A9KRo3buWhtLSMlCX61qDWyKW3lw69LcjV6vHm58OxYUca7hcUs0PwwuISVu"
        LogoBase64 = LogoBase64 & "kxhERZVoSecnJJjyeOSKRDAYu+Onug0ae+SNxzjPXRJII26JClk+8MVlVLO5PJtpisgiXTB2tQi+QCM2lGi0uXDu74x+fDMTYyjiUiUxdvhtY9"
        LogoBase64 = LogoBase64 & "hKsG0eAIa8VPObm8LpUk23qHWeobbRq0G4DwJZvZzRCqvQuFtkNZOTfZ5OgZOAe7jkgLKragVsnNupbLLc3i0xj6P32j6Pvaoi+3FDsaWPmQ9q"
        LogoBase64 = LogoBase64 & "oW+n8Cclk8sclowPdN2FdKDu3cWB154k/r8d8dadh5JJ2dDEXGbWdbHzpijFovzUVshc3k0ltmOlgmtHSzjlzjvz+u+gUNPvS23N5YdJ677UD7"
        LogoBase64 = LogoBase64 & "ROo83ZBgpPPtOeqxIlFK7pzYZGhaGweT6RmqLGLQnaUGlL3TYJO+nRs7Q62SXJqAfBzGIoyYXF6fdgdek6JRl/yL8wYC/UYZ9Qe9uc8PbZ3e78"
        LogoBase64 = LogoBase64 & "l8zlmzTeLzUWATuZTk0D2ouq082EE4oY5Wj3NWkMv7oIxxzLxYNO1G13WMZ5zsxoOO/U33mhx0wayDtEzRKQ0t0Xx7dVt7YvVWKbmRcSmo/6E3"
        LogoBase64 = LogoBase64 & "6rb2Mup54b0egbJxkaxLOchWhnounD6dwLh4TWIkim1I6PdnHNzNcbTyRNPuw3FWhlzehips/cOi0Yiu+dA9LP7qDk1gY1bMfmvRjx1idPAJx8"
        LogoBase64 = LogoBase64 & "b/HZb19yiwidxruXmYEbOF3R8aP38dAzuQlynOVwZOHmLd9kPwDF2ITkOmoY1XGNoNmIIvR/zIjriOpF8y+aPrn/T28u2NmxeH42ezJD73nTiH"
        LogoBase64 = LogoBase64 & "CfQ8Mo7Ti4xDREyibFwkVFX7bsFGkz7dt5oftx037pgP34X69PvoObGmOMZGrmVntDdl9IV2hUXF7IJeV/8f8P43QdyEatOfXSKgqzq0PHcZFs"
        LogoBase64 = LogoBase64 & "FyjWvVXLqzFVaTS6hKxLpVQSjUIbrxdz77umkfKPRXmVTlUyhiver0ScS6j6IvtqNrwUmpJ9i3eNLCeIRFxzPSafLSM2t82QqbyH3ckBOxjlJQ"
        LogoBase64 = LogoBase64 & "nYj1HweeKLkq/lio5CoYKrkKhkqugqGSq2Co5CoYKrkKhkqugqGSq2Co5CoYKrkKhkqugqGSq2Co5CoYKrkKhkqugqGSq2D8HxexiJXc7hgxAA"
        LogoBase64 = LogoBase64 & "AAAElFTkSuQmCC"
    End If

    getLogoBase64 = LogoBase64
End Function

Function CountUsedColumns(ws As Worksheet) As Integer
    Dim col As Long
    Dim count As Long
    
    count = 0
    col = 1 ' Start at column A
    
    Do While ws.Cells(2, col).Value <> ""
        count = count + 1
        col = col + 1
    Loop
    
    CountUsedColumns = count
End Function

Function GetSelectedWorksheets() As Variant
    Dim selectedSheets As Sheets
    Dim nameArray() As String
    Dim i As Long

    ' Validate ActiveWindow and selected sheets
    If ActiveWindow Is Nothing Then
        Debug.Print "GetSelectedWorksheets[ERROR]: No active window.", vbExclamation
        Exit Function
    End If

    Set selectedSheets = ActiveWindow.selectedSheets
    If selectedSheets.count = 0 Then
        Debug.Print "GetSelectedWorksheets[ERROR]: No sheets selected.", vbExclamation
        Exit Function
    End If

    ReDim nameArray(1 To selectedSheets.count)
    For i = 1 To selectedSheets.count
        nameArray(i) = selectedSheets(i).Name
    Next i

    GetSelectedWorksheets = nameArray
End Function

Sub AutoFitAllRows(ws As Worksheet)
    ws.rows.AutoFit
End Sub

Sub AutoWrapTextUsedCells(ws As Worksheet)
    Dim cell As Range

    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) Then
            cell.wrapText = True
        End If
    Next cell
End Sub

'-----------------------------------------------------------------------------
' Function: GetUsedColumnRangeString
' Purpose : Returns the range of used columns on a worksheet in "A:Z" format.
' Input   : ws - Worksheet object from which to get the used column range.
' Output  : String representing the used column range in letter format (e.g., "B:G").
'-----------------------------------------------------------------------------
Function GetUsedColumnRangeString(ws As Worksheet) As String
    Dim firstCol As Long, lastCol As Long
    Dim colRangeStr As String

    ' Determine first and last used column numbers within the UsedRange
    With ws.UsedRange
        firstCol = .Columns(1).Column
        lastCol = .Columns(.Columns.count).Column
    End With

    ' Convert column numbers to letter references (e.g., 1 ? A, 26 ? Z)
    colRangeStr = Split(Cells(1, firstCol).Address(True, False), "$")(0) & ":" & _
                  Split(Cells(1, lastCol).Address(True, False), "$")(0)

    ' Return the column letter range
    GetUsedColumnRangeString = colRangeStr
End Function

'-----------------------------------------------------------------------------
' Function: GetUsedColumnCount
' Purpose : Returns the number of used columns in the worksheet's UsedRange.
' Input   : ws - Worksheet object from which to calculate the column count.
' Output  : Long - Number of columns used.
'-----------------------------------------------------------------------------
Function GetUsedColumnCount(ws As Worksheet) As Long
    With ws.UsedRange
        ' Calculate the number of used columns based on their positions
        GetUsedColumnCount = .Columns(.Columns.count).Column - .Columns(1).Column + 1
    End With
End Function

Function ValueExistsInArray(val As Variant, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If element = val Then
            ValueExistsInArray = True
            Exit Function
        End If
    Next element
    ValueExistsInArray = False
End Function

Function RemoveItemsByValues(sourceArr As Variant, itemsToRemove As Variant) As Variant
    Dim result() As Variant
    Dim i As Long, j As Long
    Dim skip As Boolean
    Dim count As Long: count = -1

    For i = LBound(sourceArr) To UBound(sourceArr)
        skip = False
        For j = LBound(itemsToRemove) To UBound(itemsToRemove)
            If sourceArr(i) = itemsToRemove(j) Then
                skip = True
                Exit For
            End If
        Next j
        If Not skip Then
            count = count + 1
            ReDim Preserve result(count)
            result(count) = sourceArr(i)
        End If
    Next i

    RemoveItemsByValues = result
End Function

Function ParseColorString(colorInput As String) As Long
    Dim inputStr As String
    inputStr = Trim(LCase(colorInput))
    
    ' RGB format: "r,g,b"
    If InStr(inputStr, ",") > 0 Then
        Dim rgbParts() As String
        rgbParts = Split(inputStr, ",")
        
        If UBound(rgbParts) = 2 Then
            On Error GoTo InvalidRGB
            ParseColorString = RGB( _
                CLng(Trim(rgbParts(0))), _
                CLng(Trim(rgbParts(1))), _
                CLng(Trim(rgbParts(2))) _
            )
            Exit Function
        End If
    End If
    
    ' HEX format: "#rrggbb" or "rrggbb"
    If Left(inputStr, 1) = "#" Then inputStr = Mid(inputStr, 2)
    If Len(inputStr) = 6 Then
        On Error GoTo InvalidHEX
        ParseColorString = RGB( _
            CLng("&H" & Mid(inputStr, 1, 2)), _
            CLng("&H" & Mid(inputStr, 3, 2)), _
            CLng("&H" & Mid(inputStr, 5, 2)) _
        )
        Exit Function
    End If

InvalidRGB:
InvalidHEX:
    Err.Clear
    ParseColorString = -1 ' Return -1 to indicate invalid input
End Function

Sub BoldTextBetwwenPattern(ws As Worksheet, pattern As String)
    Dim cell As Range
    Dim matches As Object, match As Object
    Dim regex As Object
    Dim startPos As Long
    Dim txt As String

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = pattern

    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) And Not IsError(cell.Value) Then
            txt = CStr(cell.Value)

            If regex.test(txt) Then
                ' Remove o negrito antes de aplicar as regras.
                cell.Font.Bold = False
                
                If cell.HasFormula Then
                    Dim formulaText As String
                    formulaText = cell.Formula
                    
                    ' Detecta fórmulas do tipo CONCAT, TEXTJOIN, etc.
                    If InStr(1, formulaText, "CONCAT", vbTextCompare) > 0 Or _
                       InStr(1, formulaText, "TEXTJOIN", vbTextCompare) > 0 Then
                        cell.Font.Bold = True
                    Else
                        ' Fórmulas genéricas — aplicar negrito parcial se possível
                        Set matches = regex.Execute(txt)
                        For Each match In matches
                            startPos = InStr(txt, match.Value)
                            If startPos > 0 Then
                                cell.Characters(Start:=startPos, Length:=Len(match.Value)).Font.Bold = True
                            End If
                        Next match
                    End If
                Else
                    ' Célula sem fórmula — aplicar negrito ao trecho correspondente
                    Set matches = regex.Execute(txt)
                    For Each match In matches
                        startPos = InStr(txt, match.Value)
                        If startPos > 0 Then
                            cell.Characters(Start:=startPos, Length:=Len(match.Value)).Font.Bold = True
                        End If
                    Next match
                End If
            End If
        End If
    Next cell
End Sub

Sub pasteDataMarker(targetWS As Worksheet, marker As String, Optional after As Boolean = True, Optional startBlock As Boolean = True)
    ' Encontra a última linha preenchida na planilha de destino
    Dim totalRowsTarget As Long
    totalRowsTarget = targetWS.Cells(targetWS.rows.count, 1).End(xlUp).row
    
    Dim markerRowTarget As Long
    markerRowTarget = 0
    
    Dim pattern As String
    
    If startBlock Then
        pattern = "{% block " & marker & " %}"
    Else
        pattern = "{% endblock " & marker & " %}"
    End If

    ' Localiza o marcador na coluna A da planilha de destino
    For i = 1 To totalRowsTarget
        If Trim(LCase(targetWS.Cells(i, 1).Value)) = LCase(pattern) Then
            markerRowTarget = i
            Exit For
        End If
    Next i

    If markerRowTarget = 0 Then
        MsgBox "Marker not found in target sheet.", vbExclamation
        Exit Sub
    End If
    
    Dim insertRow As Long
    ' After
    If after Then
        insertRow = markerRowTarget + 1
    ' Before
    Else
        insertRow = markerRowTarget - 1
    End If

    ' Insere o conteúdo após o marcador na planilha de destino
    targetWS.Cells(insertRow, 1).EntireRow.Insert Shift:=xlDown
    
    Application.CutCopyMode = False
End Sub

Sub InsertTextMarker(targetWS As Worksheet, marker As String, textToInsert As String, Optional after As Boolean = True)
    '------------------------------------------------------------------------------
    ' Function : InsertTextMarker
    '
    ' Purpose  : Localiza um marcador "{% block <marker> %}" na coluna A da planilha
    '            e insere uma nova linha contendo o texto especificado na linha seguinte.
    '
    ' Parameters:
    ' - targetWS     : Worksheet onde o conteúdo será inserido.
    ' - marker       : Nome do marcador (sem "block"), ex: "NotasExplicativas".
    ' - textToInsert : Texto a ser inserido na nova linha após o marcador.
    '
    ' Notes:
    ' - O texto é inserido na coluna A, uma linha abaixo do marcador.
    ' - Caso o marcador não seja encontrado, uma mensagem de erro é exibida.
    '------------------------------------------------------------------------------

    ' Determina a última linha preenchida na coluna A
    Dim totalRowsTarget As Long
    totalRowsTarget = targetWS.Cells(targetWS.rows.count, 1).End(xlUp).row
    
    Dim markerRowTarget As Long
    markerRowTarget = 0
    
    Dim i As Long
    ' Procura pelo marcador na coluna A
    For i = 1 To totalRowsTarget
        If Trim(LCase(targetWS.Cells(i, 1).Value)) = LCase("{% block " & marker & " %}") Then
            markerRowTarget = i
            Exit For
        End If
    Next i

    ' Se o marcador não for encontrado, aborta
    If markerRowTarget = 0 Then
        Debug.Print "Marcador '{% block " & marker & " %}' não encontrado na planilha.", vbExclamation
        Exit Sub
    End If

    ' Insere uma nova linha após o marcador e insere o texto
    Dim insertRow As Long
    insertRow = markerRowTarget + 1
    targetWS.rows(insertRow).Insert Shift:=xlDown
    targetWS.Cells(insertRow, 1).Value = textToInsert
End Sub

Function copyDataBetweenMarker(ws As Worksheet, marker As String, Optional cut As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    ' Function : copyDataBetweenMarker
    '
    ' Purpose  : Localiza e copia todas as linhas entre os marcadores
    '            personalizados "{% block <marker> %}" e "{% endblock <marker> %}"
    '            na coluna A da planilha especificada.
    '
    ' Parameters:
    ' - ws     : Planilha onde será feita a busca pelos marcadores e cópia das linhas.
    ' - marker : Identificador do bloco (exemplo: "NotasExplicativas", "DetalhesTecnicos").
    '
    ' Usage    : copyDataBetweenMarker ActiveSheet, "NotasExplicativas"
    '
    ' Notes    :
    ' - Os marcadores devem estar em células da coluna A (A1:A1000).
    ' - Apenas as linhas entre os marcadores são copiadas, excluindo os próprios marcadores.
    ' - Os dados copiados são colocados na área de transferência (clipboard).
    ' - Uma mensagem será exibida caso os marcadores não sejam encontrados.
    '------------------------------------------------------------------------------

    Dim startRow As Long, endRow As Long
    Dim cell As Range

    ' Buscar o início do bloco na coluna A (valor exato do marcador de abertura)
    For Each cell In ws.Range("A1:A1000")
        If Trim(cell.Value) = "{% block " & marker & " %}" Then
            startRow = cell.row
            Exit For
        End If
    Next cell

    ' Buscar o fim do bloco a partir da linha seguinte
    For Each cell In ws.Range("A" & startRow + 1 & ":A1000")
        If Trim(cell.Value) = "{% endblock " & marker & " %}" Then
            endRow = cell.row
            Exit For
        End If
    Next cell

    ' Verifica se os dois marcadores foram localizados corretamente
    If startRow = 0 Or endRow = 0 Then
        Debug.Print "Marcadores não encontrados.", vbExclamation
        copyDataBetweenMarker = False
        Exit Function
    End If
    
    If cut Then
        ' Recortar as linhas entre os dois marcadores (exclui os marcadores em si)
        ws.rows((startRow + 1) & ":" & (endRow - 1)).cut
    Else
        ' Copia as linhas entre os dois marcadores (exclui os marcadores em si)
        ws.rows((startRow + 1) & ":" & (endRow - 1)).Copy
    End If
    
    copyDataBetweenMarker = True

    ' Mensagem de confirmação ao usuário
    'MsgBox "Linhas copiadas com sucesso: " & (endRow - startRow - 1) & " linha(s).", vbInformation
End Function

Sub DeleteAllBlockMarkers(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim targetCol As Range
    Dim cell As Range

    ' Colunas onde os marcadores podem estar
    Set targetCol = Union(ws.Range("A1:A1000"), ws.Range("B1:B1000"))

    ' Verifica de baixo para cima para evitar problemas ao excluir
    For i = targetCol.rows.count To 1 Step -1
        For Each cell In targetCol.rows(i).Cells
            cellValue = Trim(cell.Value)
            If LCase(cellValue) Like "*{% block *%}*" Or _
               LCase(cellValue) Like "*{% endblock *%}*" Then
                ws.rows(cell.row).Delete
                Exit For ' Sai do segundo loop após deletar a linha
            End If
        Next cell
    Next i
End Sub

Sub ZoomAllSheetsTo100(wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        On Error Resume Next
        ws.Activate
        ActiveWindow.View = xlPageBreakPreview
        ActiveWindow.Zoom = 100
        On Error GoTo 0
    Next ws
    wb.Worksheets(1).Activate ' Return to first sheet if needed
End Sub

Sub GoToA1InAllSheets(wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        On Error Resume Next
        ws.Activate
        ws.Range("A1").Select
        On Error GoTo 0
    Next ws
    wb.Worksheets(1).Activate ' Optional: return to the first worksheet
End Sub

Sub changeFontFace(ws As Worksheet, fontFace As String)
    With ws.UsedRange
        .Font.Name = fontFace
    End With
End Sub

Public Sub FormatarCelulasComTraco(ws As Worksheet)
    Debug.Print "Formatando a planilha: " & ws.Name
    Dim cell As Range
    Dim rng As Range
    Dim cellText As String

    Set rng = ws.UsedRange

    Application.ScreenUpdating = False

    For Each cell In rng
        ' Evita erro ao acessar a propriedade Text
        If Not IsError(cell.Value) Then
            ' Captura o texto visível na célula (o que o usuário vê na planilha)
            cellText = CStr(cell.text)

            ' Verifica se o texto visível é exatamente "-"
            If Trim(cellText) = "-" Then
                With cell.Interior
                    .pattern = xlDown ' Preenchimento diagonal
                    .PatternColorIndex = xlAutomatic
                    .PatternColor = RGB(230, 230, 230)
                    '.Color = xlNone
                    .PatternTintAndShade = 0
                End With
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
End Sub

Sub ColorirCelulasComFormulas(ws As Worksheet)
    Dim cell As Range
    Dim rng As Range

    Application.ScreenUpdating = False
    
    Set rng = ws.UsedRange

    For Each cell In rng
        With cell.Interior
            If cell.HasFormula Then
                .Color = RGB(204, 255, 229) ' Verde
                '.pattern = xlGrid ' Padrão de grade
                '.PatternColorIndex = xlAutomatic ' Cor do padrão como automático
            ElseIf Not IsEmpty(cell) Then
                .Color = xlNone
                '.Color = RGB(0, 255, 0) ' Fundo verde
                '.pattern = xlNone ' Sem padrão
            Else
                .ColorIndex = xlNone ' Sem cor de fundo
                '.pattern = xlNone ' Sem padrão
            End If
        End With
    Next cell

    Application.ScreenUpdating = True
End Sub


