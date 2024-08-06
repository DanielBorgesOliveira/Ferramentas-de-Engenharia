Attribute VB_Name = "Módulo1"
'Option Explicit

Sub ResizeImages()
    Dim i As Long
    With ActiveDocument
        For i = 1 To .InlineShapes.Count - 1
            With .InlineShapes(i)
                Aspect = .width / .Height
                .width = CentimetersToPoints(16)
                .Height = .width / Aspect
                Set rng = .Range
                rng.style = ActiveDocument.Styles("VALE_IMAGEM")
            End With
        Next i
    End With
End Sub
