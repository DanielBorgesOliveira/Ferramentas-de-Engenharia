Attribute VB_Name = "teste"
Sub format()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        ' Format cells with only "-" text
        Call libUtils.FormatarCelulasComTraco(ws)
    Next
End Sub

Sub Teste()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Call libUtils.ZoomAllSheetsTo100(wb)
End Sub
