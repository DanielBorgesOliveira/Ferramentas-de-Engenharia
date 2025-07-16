Attribute VB_Name = "FinishHim"
Option Explicit

Private Type ClientStatus
    CorrectionNeeded As Boolean
    Words As String
End Type

Sub FinishHim()
        
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    If libUtils.propertyExists(wb, "NumeroCliente") Then UserForm1.NumeroCliente = wb.CustomDocumentProperties("NumeroCliente").Value
    If libUtils.propertyExists(wb, "NumeroNosso") Then UserForm1.NumeroNosso = wb.CustomDocumentProperties("NumeroNosso").Value
    If libUtils.propertyExists(wb, "Revisao") Then UserForm1.Revisao = wb.CustomDocumentProperties("Revisao").Value
    If libUtils.propertyExists(wb, "Titulo1") Then UserForm1.Titulo1 = wb.CustomDocumentProperties("Titulo1").Value
    If libUtils.propertyExists(wb, "Titulo2") Then UserForm1.Titulo2 = wb.CustomDocumentProperties("Titulo2").Value
    If libUtils.propertyExists(wb, "Titulo3") Then UserForm1.Titulo3 = wb.CustomDocumentProperties("Titulo3").Value
    If libUtils.propertyExists(wb, "Titulo4") Then UserForm1.Titulo4 = wb.CustomDocumentProperties("Titulo4").Value
    If libUtils.propertyExists(wb, "Titulo5") Then UserForm1.Titulo5 = wb.CustomDocumentProperties("Titulo5").Value
    If libUtils.propertyExists(wb, "Cliente") Then UserForm1.Cliente = wb.CustomDocumentProperties("Cliente").Value
    If libUtils.propertyExists(wb, "Projeto") Then UserForm1.Projeto = wb.CustomDocumentProperties("Projeto").Value
    If libUtils.propertyExists(wb, "Fase") Then UserForm1.Fase = wb.CustomDocumentProperties("Fase").Value
    If libUtils.propertyExists(wb, "NumeroProjeto") Then UserForm1.Fase = wb.CustomDocumentProperties("NumeroProjeto").Value
    
    UserForm1.Show
    
    wb.BuiltinDocumentProperties("Title").Value = wb.Name
    wb.BuiltinDocumentProperties("Author").Value = Application.UserName
    wb.BuiltinDocumentProperties("Company").Value = "Brass do Brasil"
    Call libUtils.UpdateProperty(wb, "NumeroCliente", UserForm1.NumeroCliente)
    Call libUtils.UpdateProperty(wb, "NumeroNosso", UserForm1.NumeroNosso)
    Call libUtils.UpdateProperty(wb, "Revisao", UserForm1.Revisao)
    Call libUtils.UpdateProperty(wb, "Titulo1", UserForm1.Titulo1)
    Call libUtils.UpdateProperty(wb, "Titulo2", UserForm1.Titulo2)
    Call libUtils.UpdateProperty(wb, "Titulo3", UserForm1.Titulo3)
    Call libUtils.UpdateProperty(wb, "Titulo4", UserForm1.Titulo4)
    Call libUtils.UpdateProperty(wb, "Titulo5", UserForm1.Titulo5)
    Call libUtils.UpdateProperty(wb, "Cliente", UserForm1.Cliente)
    Call libUtils.UpdateProperty(wb, "Projeto", UserForm1.Projeto)
    Call libUtils.UpdateProperty(wb, "Fase", UserForm1.Fase)
    Call libUtils.UpdateProperty(wb, "NumeroProjeto", UserForm1.NumeroProjeto)
    
    Dim client As String
    client = UserForm1.Cliente.Value
    
    Dim fontFace As String
    Select Case client
    Case "Samarco"
        fontFace = "Times New Roman"
    Case "AngloAmerican"
        fontFace = "Aptos"
    Case Else
        fontFace = "Arial"
    End Select
    
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        ' Delete all shapes, except images
        'Call libUtils.DeleteAllShapes(ws, "Image")
        
        ' Delete all notes.
        Call libUtils.DeleteAllNotes(ws)
        Call libUtils.DeleteAllThreadedComments(ws)
        
        ' Change to client's font face.
        Call libUtils.changeFontFace(ws, fontFace)
        
        ' Make the text between bracketsbold.
        Call libUtils.BoldTextBetwwenPattern(ws, "\[[^\]]+\]")
        Call libUtils.BoldTextBetwwenPattern(ws, "\(Notas? [^\)]+\)")
        
        If InStr(1, LCase(ws.Name), "capa") = 0 Then
            ' Format cells with only "-" text
            Call libUtils.FormatarCelulasComTraco(ws)
        End If
    Next
    
    Call libUtils.ZoomAllSheetsTo100(wb)
    
    Call libUtils.GoToA1InAllSheets(wb)
    
    ' Export PDF
    Call libUtils.ExportAllWorksheetsAsPDF(wb, wb.FullName)
End Sub

















































