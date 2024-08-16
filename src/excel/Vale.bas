Attribute VB_Name = "Vale"
Sub finish()
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Dim ws As Worksheet
    Set ws = wb.ActiveSheet
    
    setWidthOriginal wb, ws
    
    '  Get height reference. This will be used after merging to avoid incorrect line break.
    Dim reference As Variant
    reference = getCellsHeight(ws)
    
    AddStyles wb, ws
    
    insertColumns wb, ws
    
    mergeDataCells wb, ws, reference
    
    insertLines wb, ws
    
    mergeHeaderCells wb, ws
    
    mergeHeaderCells wb, ws
    
    setWidth wb, ws
    
    setHeight wb, ws
    
    insertText wb, ws
    
    DeleteAllShapes wb, ws
    
    DeleteAllNotes wb, ws
    
    setupPage wb
    
    ' Insert Vale's logo
    insertLogo wb, ws, Library.DecodeBase64(getValeLogoBase64(), "svg"), ws.Range("A1:F4")
    
    ' Get the original path of our logo
    Dim logoPath As String
    logoPath = Library.UseFileDialog(msoFileDialogFilePicker)
    
    ' Insert our company's logo
    insertLogo wb, ws, logoPath, ws.Range("G1:L4")
End Sub

Function GetFileExtension(ByVal filePath As String) As String
    Dim dotPosition As Long
    Dim fileExtension As String
    
    ' Find the position of the last dot in the file name
    dotPosition = InStrRev(filePath, ".")
    
    ' If there is no dot, return an empty string
    If dotPosition > 0 Then
        fileExtension = Right(filePath, Len(filePath) - dotPosition)
    Else
        fileExtension = ""
    End If
    
    GetFileExtension = fileExtension
End Function

Sub DeleteAllShapes(wb As Workbook, ws As Worksheet)
    Dim i As Long
    
    ' Loop through all shapes in the worksheet in reverse order
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

Sub DeleteAllNotes(wb As Workbook, ws As Worksheet)
    Dim cell As Range
    
    ' Loop through all cells in the used range of the worksheet
    For Each cell In ws.UsedRange
        ' Check if the cell has a comment (note in modern Excel)
        If Not cell.Comment Is Nothing Then
            cell.Comment.Delete ' This should completely delete the note
        End If
    Next cell
End Sub

Sub setWidthOriginal(wb As Workbook, ws As Worksheet)
    ws.Columns("A:A").ColumnWidth = 5
    ws.Columns("B:B").ColumnWidth = 32.5
    ws.Columns("C:C").ColumnWidth = 5
    ws.Columns("D:D").ColumnWidth = 15
    
    With ws
        .Cells.Font.Name = "Arial"
        .Cells.Font.Size = 8
    End With
End Sub

Sub createStyle( _
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
    Optional Border As Variant = Nothing)
    
    Dim customStyle As Style
    
    ' Set default color if not provided
    If fontColor = -1 Then fontColor = RGB(0, 0, 0)
    
    ' Check if the style exists
    On Error Resume Next
        Set customStyle = wb.Styles(styleName)
    On Error GoTo 0
    
    ' If the style doesn't exist, create it
    If customStyle Is Nothing Then
        Set customStyle = wb.Styles.Add(Name:=styleName)
    End If
    
    ' Update the style properties
    With customStyle
        .Font.Name = fontName
        .Font.Size = fontSize
        .Font.Bold = fontBold
        .Font.Italic = fontItalic
        .Font.Underline = IIf(fontUnderline, xlUnderlineStyleSingle, xlUnderlineStyleNone)
        .Font.Color = fontColor
        '.Interior.Color = RGB(200, 200, 200) ' Light gray background color
        
        .HorizontalAlignment = HorizontalAlignment
        .VerticalAlignment = VerticalAlignment
        .wrapText = wrapText
        
        ' Apply borders if provided
        If Not IsMissing(Border) Then
            ' Top Border
            .Borders(Border(0)(0)).LineStyle = Border(0)(1)
            .Borders(Border(0)(0)).Color = RGB(0, 0, 0)
            ' Right Border
            .Borders(Border(1)(0)).LineStyle = Border(1)(1)
            .Borders(Border(1)(0)).Color = RGB(0, 0, 0)
            ' Bottom Border
            .Borders(Border(2)(0)).LineStyle = Border(2)(1)
            .Borders(Border(2)(0)).Color = RGB(0, 0, 0)
            ' Left Border
            .Borders(Border(3)(0)).LineStyle = Border(3)(1)
            .Borders(Border(3)(0)).Color = RGB(0, 0, 0)
        End If
    End With
End Sub

Sub AddStyles(wb As Workbook, ws As Worksheet)
    createStyle wb, "Arial", "6ptLeft", 6, False, False, False, True, RGB(0, 0, 0), xlLeft, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous))
    createStyle wb, "Arial", "6ptLeftBold", 6, True, False, False, True, RGB(0, 0, 0), xlLeft, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous))
    createStyle wb, "Arial", "6ptCenterBold", 6, True, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous))
    createStyle wb, "Arial", "8ptLeft", 8, False, False, False, True, RGB(0, 0, 0), xlLeft, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous))
    createStyle wb, "Arial", "8ptLeftBold", 8, True, False, False, True, RGB(0, 0, 0), xlLeft, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous))
    createStyle wb, "Arial", "8ptCenter", 8, False, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous))
    createStyle wb, "Arial", "9ptCenterBold", 9, True, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous))
    createStyle wb, "Arial", "10ptCenterBold", 10, True, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous))
    createStyle wb, "Arial", "10ptLeftBold", 10, True, False, False, True, RGB(0, 0, 0), xlLeft, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous))
    createStyle wb, "Arial", "11ptCenterBold", 11, True, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous))
    createStyle wb, "Arial", "12ptCenterBold", 12, True, False, False, True, RGB(0, 0, 0), xlCenter, xlCenter, Array(Array(xlEdgeTop, xlContinuous), Array(xlEdgeRight, xlContinuous), Array(xlEdgeBottom, xlContinuous), Array(xlEdgeLeft, xlContinuous))
End Sub

Function getCellsHeight(ws As Worksheet) As Variant
    Dim reference() As Double
    Dim lastRow As Long
    Dim iRow As Range
    Dim i As Long
    
    ' Get the last used row to initialize the array with the correct size
    lastRow = ws.UsedRange.Rows.Count
    ReDim reference(1 To lastRow)
    
    ' Loop through each row in the used range and store the row height
    i = 1
    For Each iRow In ws.UsedRange.Rows
        reference(i) = iRow.rowHeight
        i = i + 1
    Next iRow
    
    ' Return the array containing the row heights
    getCellsHeight = reference
End Function

Sub insertColumns(wb As Workbook, ws As Worksheet)
    ws.Columns("B").Insert Shift:=xlToRight
    ws.Columns("D:O").Insert Shift:=xlToRight
    ws.Columns("Q").Insert Shift:=xlToRight
    ws.Columns("S:W").Insert Shift:=xlToRight
End Sub

Sub mergeDataCells(wb As Workbook, ws As Worksheet, reference As Variant)
    Dim iRow As Range
    Dim rangeToMerge As Variant
    Dim mergeRanges As Variant
    Dim leftCell As Range
    
    For Each iRow In ws.UsedRange.Rows
        mergeRanges = Array( _
            "A" & iRow.Row & ":" & "B" & iRow.Row, _
            "C" & iRow.Row & ":" & "O" & iRow.Row, _
            "P" & iRow.Row & ":" & "Q" & iRow.Row, _
            "R" & iRow.Row & ":" & "W" & iRow.Row, _
            "X" & iRow.Row & ":" & "AD" & iRow.Row _
        )
        
        For Each rangeToMerge In mergeRanges
            With ws.Range(rangeToMerge)
                ' Merge the cells
                .Merge
                .Borders.LineStyle = xlContinuous
                .Borders.ColorIndex = 0
                .Borders.TintAndShade = 0
                .Borders.Weight = xlThin
                
                .Font.Name = "Arial"
                .Font.Size = 8
                .Font.Color = RGB(0, 0, 0)
                
                ' Check if the left cell is within valid bounds
                If .Cells(1, 1).Column > 1 Then
                    ' Check the cell to the left of the merge range
                    Set leftCell = .Cells(1, 1).Offset(0, -1)
                    
                    ' If the left cell's fill color is RGB(217, 217, 217), apply the same color to the merged range
                    If leftCell.Interior.Color = RGB(217, 217, 217) Then
                        .Interior.Color = RGB(217, 217, 217)
                    Else
                        .Interior.ColorIndex = xlNone
                    End If
                Else
                    
                    If Not .Cells(1, 1).Interior.Color = RGB(217, 217, 217) Then
                        ' Remove fill color if there is no valid left cell
                        .Interior.ColorIndex = xlNone
                    End If
                End If
            End With
        Next rangeToMerge
        
        ' Set the row height of the current row using the reference array
        iRow.rowHeight = reference(iRow.Row)
    Next iRow
End Sub

Sub insertLines(wb As Workbook, ws As Worksheet)
    ws.Rows("1:26").Insert Shift:=xlDown
End Sub

Sub mergeHeaderCells(wb As Workbook, ws As Worksheet)
    ' Define ranges to merge
    Dim rangesToMerge As Variant
    rangesToMerge = Array( _
        "A1:F4", "G1:L4", "AA6:AD7", "A11:AD11", "A15:B15", _
        "A16:B16", "A17:B17", "A18:B18", "A19:B19", "C15:D15", _
        "C16:D16", "C17:D17", "C18:D18", "C19:D19", "E15:R15", _
        "E16:R16", "E17:R17", "E18:R18", "E19:R19", "S15:T15", _
        "S16:T16", "S17:T17", "S18:T18", "S19:T19", "U15:V15", _
        "U16:V16", "U17:V17", "U18:V18", "U19:V19", "W15:X15", _
        "W16:X16", "W17:X17", "W18:X18", "W19:X19", "Y15:Z15", _
        "Y16:Z16", "Y17:Z17", "Y18:Z18", "Y19:Z19", "AA15:AD15", _
        "AA16:AD16", "AA17:AD17", "AA18:AD18", "AA19:AD19", "H24:O24", _
        "T24:AD24", "H25:O25", "T25:AD25", "A26:AD26" _
    )
    
    ' Find the cell containing text "Notas Explicativas"
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="Notas Explicativas", LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Save the last position + 1 in the rangesToMerge array. This will be used to merge the notes header.
    Dim LastPosition As Integer
    LastPosition = UBound(rangesToMerge) + 1
    
    ' Check if the cell was found
    Dim nextCell As Integer
    If Not foundCell Is Nothing Then
        ' Get the row number of the found cell
        nextCell = foundCell.Row
        
        ' Add space to rangesToMerge to merge the notes header.
        ReDim Preserve rangesToMerge(LBound(rangesToMerge) To LastPosition)
        rangesToMerge(LastPosition) = "A" & foundCell.Row & ":AD" & foundCell.Row
    End If
    
    ' Count the number of notes in the document
    Dim NumberNotes As Integer
    NumberNotes = ws.UsedRange.Rows.Count - nextCell
    
    ' Add space to rangesToMerge array.
    ReDim Preserve rangesToMerge(LBound(rangesToMerge) To (UBound(rangesToMerge) + NumberNotes))
    
    ' Append the ranges to merge
    For j = 1 To NumberNotes
        rangesToMerge(UBound(rangesToMerge) - NumberNotes + j) = "C" & (nextCell + j) & ":AD" & (nextCell + j)
    Next j
    
    'Dim i As Integer
    'For i = LBound(rangesToMerge) To UBound(rangesToMerge)
    Dim rng As Variant
    For Each rng In rangesToMerge
        With ws.Range(rng)
            .Merge
            .Borders.LineStyle = xlContinuous
            .Borders.ColorIndex = 0
            .Borders.TintAndShade = 0
            .Borders.Weight = xlThin
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    Next rng
    'Next i
End Sub

Sub setWidth(wb As Workbook, ws As Worksheet)
    For Each col In ws.UsedRange.Columns
        ws.Columns(col.Column).ColumnWidth = 2.5
    Next col
End Sub

Sub setHeight(wb As Workbook, ws As Worksheet)
    ws.Rows("1:10").rowHeight = 12
    ws.Rows("11:11").rowHeight = 5
    ws.Rows("12:20").rowHeight = 12
    ws.Rows("21:23").rowHeight = 15
    ws.Rows("24:25").rowHeight = 12
    ws.Rows("26:26").rowHeight = 5
    ws.Rows("27:27").rowHeight = 12
End Sub

Sub insertText(wb As Workbook, ws As Worksheet)
    Dim TextArray(33) As Variant
    TextArray(0) = Array( _
        "S8:Z8", _
        "Nº BRASS", _
        "8ptLeft", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlNone), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(1) = Array( _
        "S5:Z5", _
        "Nº VALE", _
        "8ptLeft", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlNone), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(2) = Array( _
        "AA5:AD5", _
        "PÁGINA", _
        "8ptCenter", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlNone), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(3) = Array( _
        "AA8:AD8", _
        "REV.", _
        "8ptCenter", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlNone), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(4) = Array( _
        "M1:R1", _
        "CLASSIFICAÇÃO", _
        "8ptCenter", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlNone), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(5) = Array( _
        "S9:Z10", _
        "BdB201452-0000-V-FD0005", _
        "10ptCenterBold", _
        Array( _
            Array(xlEdgeTop, xlNone), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(6) = Array( _
        "M2:R4", _
        "RESTRITA", _
        "9ptCenterBold", _
        Array( _
            Array(xlEdgeTop, xlNone), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(7) = Array( _
        "S6:Z7", _
        "FD-1880HH-T-31903", _
        "10ptCenterBold", _
        Array( _
            Array(xlEdgeTop, xlNone), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(8) = Array( _
        "AA9:AD10", _
        "A", _
        "10ptCenterBold", _
        Array( _
            Array(xlEdgeTop, xlNone), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(9) = Array( _
        "A5:R10", _
        "PROJETO DETALHADO" & vbCrLf & "SISTEMA DE REJEITO" & vbCrLf & "COMPRESSORES DE AR" & vbCrLf & "FOLHA DE DADOS", _
        "10ptLeftBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(10) = Array( _
        "S1:AD4", _
        "BOMBEAMENTO DE ULTRAFINOS DE FABRICA - AREA 08" & vbCrLf & "S-1955-08", _
        "12ptCenterBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(11) = Array( _
        "A12:D13", _
        "TE: TIPO EMISSÃO", _
        "6ptLeftBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlNone), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(12) = Array( _
        "E12:J12", _
        "A - PRELIMINAR", _
        "6ptLeft", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlNone), _
            Array(xlEdgeBottom, xlNone), _
            Array(xlEdgeLeft, xlNone) _
        ) _
    )
    TextArray(13) = Array( _
        "E13:J13", _
        "B - PARA APROVAÇÃO", _
        "6ptLeft", _
        Array( _
            Array(xlEdgeTop, xlNone), _
            Array(xlEdgeRight, xlNone), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlNone) _
        ) _
    )
    TextArray(14) = Array( _
        "K12:P12", _
        "C - PARA CONHECIMENTO", _
        "6ptLeft", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlNone), _
            Array(xlEdgeBottom, xlNone), _
            Array(xlEdgeLeft, xlNone) _
        ) _
    )
    TextArray(15) = Array( _
        "K13:P13", _
        "D - PARA COTAÇÃO", _
        "6ptLeft", _
        Array( _
            Array(xlEdgeTop, xlNone), _
            Array(xlEdgeRight, xlNone), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlNone) _
        ) _
    )
    TextArray(16) = Array( _
        "Q12:W12", _
        "E - PARA CONSTRUÇÃO", _
        "6ptLeft", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlNone), _
            Array(xlEdgeBottom, xlNone), _
            Array(xlEdgeLeft, xlNone) _
        ) _
    )
    TextArray(17) = Array( _
        "Q13:W13", _
        "F - CONFORME COMPRADO", _
        "6ptLeft", _
        Array( _
            Array(xlEdgeTop, xlNone), _
            Array(xlEdgeRight, xlNone), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlNone) _
        ) _
    )
    TextArray(18) = Array( _
        "X12:AD12", _
        "G - CONFORME CONSTRUÍDO", _
        "6ptLeft", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlNone), _
            Array(xlEdgeLeft, xlNone) _
        ) _
    )
    TextArray(19) = Array( _
        "X13:AD13", _
        "H -  CANCELADO", _
        "6ptLeft", _
        Array( _
            Array(xlEdgeTop, xlNone), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlNone) _
        ) _
    )
    TextArray(20) = Array( _
        "A14:B14", _
        "Rev.", _
        "6ptCenterBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(21) = Array( _
        "C14:D14", _
        "TE", _
        "6ptCenterBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(22) = Array( _
        "E14:R14", _
        "Descrição", _
        "6ptCenterBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(23) = Array( _
        "S14:T14", _
        "Por", _
        "6ptCenterBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(24) = Array( _
        "U14:V14", _
        "Ver.", _
        "6ptCenterBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(25) = Array( _
        "W14:X14", _
        "Apr.", _
        "6ptCenterBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(26) = Array( _
        "Y14:Z14", _
        "Aut.", _
        "6ptCenterBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(27) = Array( _
        "AA14:AD14", _
        "Data", _
        "6ptCenterBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(28) = Array( _
        "A20:AD20", _
        "Instruções de Preenchimento pelo Fornecedor", _
        "6ptLeftBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlNone), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(29) = Array( _
        "A21:AD23", _
        "I - O Fornecedor deve preencher a primeira coluna do campo “Proposto” com uma das opções a seguir: 'A' (atendido) ou 'D' (desvio)." & vbCrLf & _
        "II - Os itens assinalados como 'D', assim como os esclarecimentos a estes pertinentes, devem obrigatoriamente ser informados pelo Fornecedor através da 'Lista de Desvios', conforme definido na Seção 1 da Requisição Técnica. Para a apresentação de informações adicionais às contidas nesta Folha de Dados, o Fornecedor deve proceder da mesma forma." & vbCrLf & _
        "III - As Notas Explicativas ao final da Folha de Dados são de preenchimento exclusivo do Emitente e não devem ser preenchidas pelo Fornecedor.", _
        "6ptLeft", _
        Array( _
            Array(xlEdgeTop, xlNone), _
            Array(xlEdgeRight, xlContinuous), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(30) = Array( _
        "A24:G24", _
        "Fornecedor:", _
        "8ptLeftBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlNone), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(31) = Array( _
        "A25:G25", _
        "Identificação (TAG):", _
        "8ptLeftBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlNone), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(32) = Array( _
        "P24:S24", _
        "Proposta:", _
        "8ptLeftBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlNone), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    TextArray(33) = Array( _
        "P25:S25", _
        "Quantidade:", _
        "8ptLeftBold", _
        Array( _
            Array(xlEdgeTop, xlContinuous), _
            Array(xlEdgeRight, xlNone), _
            Array(xlEdgeBottom, xlContinuous), _
            Array(xlEdgeLeft, xlContinuous) _
        ) _
    )
    
    ' Loop through the array and apply the values, styles, and merges
    Dim i As Integer
    Dim Text As Variant
    For Each Text In TextArray
        With ws.Range(Text(0))
            .Merge
            .Value = Text(1)
            .Style = Text(2)
            
            ' Apply Top Border
            Dim j As Integer
            For j = 0 To 3
                With .Borders(Text(3)(j)(0))
                    .LineStyle = Text(3)(j)(1)
                    '.Color = RGB(0, 0, 0)
                    '.ColorIndex = 0
                    '.TintAndShade = 0
                    '.Weight = xlThin
                End With
            Next j
        End With
    Next Text
End Sub

Sub insertLogo(wb As Workbook, ws As Worksheet, imagePath As String, targetCell As Range)
    ' Insert the image
    Dim img As Shape
    Set img = ws.Shapes.AddPicture( _
        Filename:=imagePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoCTrue, _
        Left:=0, _
        Top:=0, _
        Width:=-1, _
        Height:=-1 _
    )
    
    ' Calculate the position to center the image
    With img
        .LockAspectRatio = msoTrue ' Maintain aspect ratio of the image
        .Height = targetCell.Height * 0.8 ' Scale the image to 80% of the cell's height
        .Left = targetCell.Left + (targetCell.Width - .Width) / 2
        .Top = targetCell.Top + (targetCell.Height - .Height) / 2
        
        ' Set the image to move and size with cells
        .Placement = xlMove
    End With
End Sub

Private Function getValeLogoBase64() As String
    Dim ValeLogoBase64 As String
    ValeLogoBase64 = ""
    ValeLogoBase64 = ValeLogoBase64 & "PHN2ZyBoZWlnaHQ9IjEwMjMiIHZpZXdCb3g9IjIuMzkgLTkuMjQyIDEwNC43MDUgNTIuMzg0IiB3aWR0aD0iMjUwMCIgeG1sbnM9Imh0dHA6Ly"
    ValeLogoBase64 = ValeLogoBase64 & "93d3cudzMub3JnLzIwMDAvc3ZnIj48ZyBmaWxsLXJ1bGU9ImV2ZW5vZGQiPjxwYXRoIGQ9Im01MC45NyAxNC42MmMtNy4xNTggNS4zNzMtMTMu"
    ValeLogoBase64 = ValeLogoBase64 & "NzgyIDQuNzk0LTE5Ljg3Ni0xLjczNSAxMi4zNTMtNC43MzUgMTguODYyLTEyLjA1MSAyNS4yMDMtNS43NzhsLTQuNzAzIDYuNzI4aC0uMDA4di"
    ValeLogoBase64 = ValeLogoBase64 & "4wMWgtLjAwOHYuMDAxaC0uMDAxdi4wMWgtLjAxdi4wMDNoLS4wMDh2LjAwMWMtLjAwNSAwLS4wMDMuMDA1LS4wMDYuMDA1di4wMDNoLS4wMDF2"
    ValeLogoBase64 = ValeLogoBase64 & "LjAwM2gtLjAwMmwtLjAwMS4wMDZoLS4wMDJsLS4wMDIuMDA0aC0uMDAxYzAgLjAwMy0uMDAzLjAwMy0uMDAzLjAwNWgtLjAwMnYuMDA1Yy0uMD"
    ValeLogoBase64 = ValeLogoBase64 & "A5IDAtLjAxNS4wMTMtLjAyLjAxM2gtLjAwMWwtLjAwMS4wMDVoLjAwMnYuMDAxaC0uMDAybC0uMDA0LjAwNy0uMDA5LjAwNnYuMDAzaC0uMDAy"
    ValeLogoBase64 = ValeLogoBase64 & "YzAgLjAwNi0uMDEyLjAyMS0uMDE3LjAyNHYuMDAyaC0uMDAydi4wMTVoLS4wMDN2LjAwNGgtLjAwMnYuMDA3bC0uMDA1LjAwMmgtLjAwMmwtLj"
    ValeLogoBase64 = ValeLogoBase64 & "AwNS4wMDR2LjAwMmMtLjAwMy4wMDEtLjAwNS4wMDMtLjAwNS4wMDhsLS4wMDMuMDA0LS4wMTIuMDA1di4wMDJsLS4wMDMuMDAyYS43MjIuNzIy"
    ValeLogoBase64 = ValeLogoBase64 & "IDAgMCAwIC0uMDYzLjA4bC0uMDAyLjAwNGgtLjAwMmMwIC4wMDYtLjAwNi4wMTMtLjAwNi4wMjJoLS4wMDF2LjAwM2MtLjAwOS4wMDQtLjAyNi"
    ValeLogoBase64 = ValeLogoBase64 & "4wMjQtLjAzNi4wMjR2LjAwM2MtLjAxOC4wMDktLjAzMy4wMjUtLjA1My4wMzN2LjAwM2gtLjAwN3YuMDAzaC0uMDAybC0uMDAzLjAwNmMtLjAw"
    ValeLogoBase64 = ValeLogoBase64 & "NCAwLS4wMDUuMDAzLS4wMDkuMDAzdi4wMDJsLS4wMTUuMDF2LjAwNWgtLjAwM3YuMDAyYS4xNjMuMTYzIDAgMCAxIC0uMDMuMDJsLS4wMTIuMD"
    ValeLogoBase64 = ValeLogoBase64 & "E1aC0uMDAzdi4wMDNsLS4wMDQuMDA4di4wMDNoLS4wMDZjMCAuMTIyLS4xNzIuMTY3LS4xNzIuMjQ2aC0uMDAzYzAgLjAwOS0uMDIuMDI3LS4w"
    ValeLogoBase64 = ValeLogoBase64 & "Mi4wMzloLS4wMDNjLS4wMDMuMDA0LS4wMTMuMDE0LS4wMTMuMDE4aC0uMDAzYy0uMDA4LjAwOS0uMDAzLjA0Ny0uMDAzLjA1OCIgZmlsbD0iI2"
    ValeLogoBase64 = ValeLogoBase64 & "VjYjgzMyIvPjxwYXRoIGQ9Im01MS42MDQgMTMuODIyYy0xOC42NjQgMTYuODkyLTI1Ljk4Mi0yMy4wNjQtNDkuMjE0LTcuOTQybDI2LjE3OCAz"
    ValeLogoBase64 = ValeLogoBase64 & "Ny4yNjIiIGZpbGw9IiMwMDkzOWEiLz48ZyBjbGlwLXJ1bGU9ImV2ZW5vZGQiPjxwYXRoIGQ9Im01NS45MDQgMjQuMzUgNi4zOSAxMy44NTQgMS"
    ValeLogoBase64 = ValeLogoBase64 & "40NjgtLjAxMSA1LjYxNC0xMy45My0yLjUwNC0uMDg1LTMuOCAxMC4yOTQtNC41NzgtMTAuMjA4bTExLjY1OSAxMy44NDIgMi43NjMuMDg3IDEu"
    ValeLogoBase64 = ValeLogoBase64 & "MTIzLTMuNDZoNS44NzRsMS4xMiAzLjM3MyAyLjY3Ny4wODctNS42MTMtMTMuOTNoLTEuOTg2IiBmaWxsPSIjNzc3ODdiIi8+PHBhdGggZD0ibT"
    ValeLogoBase64 = ValeLogoBase64 & "c1LjIwNSAzMi41N2gzLjU0MWwtMS41NTUtNC44ODgiIGZpbGw9IiNmZmYiLz48L2c+PC9nPjxwYXRoIGQ9Im04Ny41OTkgMzguMDJ2LTEzLjIx"
    ValeLogoBase64 = ValeLogoBase64 & "NWgyLjQ0NXYxMC45N2g2LjA4djIuMjQ1bTMuMDU3LS4wODZ2LTEzLjQxaDcuOTE0djEuMDM2Yy0uNTc3LjM2LTEuMTI0LjgxNi0xLjY1OSAxLj"
    ValeLogoBase64 = ValeLogoBase64 & "IzM2gtNC4xNTd2Mi45NzJoNS4yMTZ2Mi4yNmgtNS4yMTZ2My42NDhoNS44MDN2Mi4yNjEiIGZpbGw9IiM3Nzc4N2IiLz48L3N2Zz4="
    
    getValeLogoBase64 = ValeLogoBase64
End Function

Sub setupPage(wb As Workbook)
    ' Loop through each worksheet in the workbook
    For Each ws In wb.Worksheets
        With ws
            .Activate ' Activate the sheet to set zoom
            
            ' Remove the AutoFilter
            .AutoFilterMode = False
            
            With .pageSetup
                .printArea = "" ' Clear the print area
                .PrintTitleRows = "$1:$10" ' Change "$1:$10" to the rows you want to repeat at the top
                
                .Zoom = False ' Turn off zoom to use FitToPages settings
                .FitToPagesWide = 1 ' Fit to one page wide
                .FitToPagesTall = False ' Let Excel determine the number of pages tall
                
                .TopMargin = Application.CentimetersToPoints(1.5)
                .BottomMargin = Application.CentimetersToPoints(1)
                .LeftMargin = Application.CentimetersToPoints(2.5)
                .RightMargin = Application.CentimetersToPoints(1)
                .HeaderMargin = Application.CentimetersToPoints(4.5)
                .FooterMargin = Application.CentimetersToPoints(0)
                
                ' Set the header text
                .RightHeader = "&""Arial,Bold""&10 &P/&N_____"
                
                ' Set the footer text
                .LeftFooter = "&""Arial,Regular""&6 Anexo A PP-G-617_Rev_01"
            End With
            
            With ActiveWindow
                ' Unfrozen the panel
                .FreezePanes = False
                
                ' Set the worksheet to Page Break Preview
                .View = xlPageBreakPreview
                
                ' Set the zoom level to 100%
                .Zoom = 100
            End With
            
            .Range("A1").Select
        End With
    Next ws
    
    wb.Worksheets(1).Activate
End Sub













Sub setHeaderHeight(wb As Workbook, ws As Worksheet)
    ' Store a reference to the first worksheet
    Dim firstSheet As Worksheet
    Set firstSheet = wb.Worksheets(1)
    
    ' Offset between the first worksheet and the followings
    Dim Offset As Integer
    Offset = 26

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Index <> firstSheet.Index Then
            ' Copy row heights from the first sheet to the current sheet
            For i = 1 To Application.WorksheetFunction.Min(1000, firstSheet.Rows.Count)
                ws.Rows(i + Offset).rowHeight = firstSheet.Rows(i).rowHeight
            Next i
            
        End If
    Next ws
End Sub
