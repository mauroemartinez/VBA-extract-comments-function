Attribute VB_Name = "PruebasExtractor"
' Pruebas del extractor de comentarios, sobre el libro de ejemplo
' (example/comments-example.xlsm). Ejecutar la macro ProbarExtractor:
' devuelve un informe y lo muestra en la ventana Inmediato (Ctrl+G).
Option Explicit

Private informe As String
Private fallas As Long

Public Function ProbarExtractor() As String
    Dim t As Variant, v As Worksheet, ws As Worksheet, n As Long, i As Long
    informe = "": fallas = 0
    Set v = ThisWorkbook.Worksheets("Ventas")

    ' 1. Funcion de una celda
    Comprobar "Hilo con dos respuestas: 3 renglones", _
        UBound(Split(ExtraerComentarios(v.Range("D2")), vbLf)) = 2
    Comprobar "Hilo: incluye el texto del comentario", _
        InStr(ExtraerComentarios(v.Range("D2")), "Revisar este precio") > 0
    Comprobar "Celda sin comentarios: texto vacio", ExtraerComentarios(v.Range("D3")) = ""
    Comprobar "Nota: incluye el texto", InStr(ExtraerComentarios(v.Range("A7")), "Producto de temporada.") > 0
    Comprobar "Formato propio con {n}, {author}, {date} y {text}", _
        Left$(ExtraerComentarios(v.Range("D2"), "{n}. {text}", " | "), 22) = "1. Revisar este precio"
    Comprobar "Separador propio", _
        UBound(Split(ExtraerComentarios(v.Range("D2"), "{text}", " | "), " | ")) = 2
    Comprobar "Alias en ingles da lo mismo", _
        ExtractComments(v.Range("D2")) = ExtraerComentarios(v.Range("D2"))

    ' 2. Tabla de una hoja (lo mismo que devuelve =TablaComentarios)
    t = TablaComentarios(v.Range("A1:G20"))
    Comprobar "Tabla de Ventas: encabezado + 7 filas", UBound(t, 1) = 8
    Comprobar "Encabezado en castellano", t(1, 1) = "Hoja" And t(1, 10) = "Texto"
    Comprobar "Orden por fila: primero D2", t(2, 4) = "D2" And t(2, 5) = 1
    Comprobar "Respuestas numeradas 2 y 3", t(3, 5) = 2 And t(4, 5) = 3
    Comprobar "Columna y fila separadas", t(2, 2) = "D" And t(2, 3) = 2
    Comprobar "Hilo resuelto = VERDADERO", t(2, 9) = True
    Comprobar "Tipo de la nota = Nota", t(8, 6) = "Nota" And t(8, 4) = "A7"
    Comprobar "La nota no tiene fecha", t(8, 8) = ""
    Comprobar "La nota de un hilo no se cuenta dos veces", Contar(t, "D2") = 3
    Comprobar "Rango que no toca comentarios: solo encabezado", _
        UBound(TablaComentarios(v.Range("E1:E20")), 1) = 1
    t = CommentsTable(v.Range("A1:G20"))
    Comprobar "Alias en ingles: encabezados en ingles", t(1, 1) = "Sheet" And t(1, 8) = "Timestamp" And t(2, 6) = "Threaded"

    ' 3. Macro de todo el libro
    ExportarComentarios
    Set ws = ThisWorkbook.Worksheets("Comentarios")
    n = ws.ListObjects(1).ListRows.Count
    Comprobar "Macro: 9 filas (7 de Ventas + 2 de Compras)", n = 9
    Comprobar "Macro: queda como tabla de Excel", ws.ListObjects(1).Name = "tblComentarios"
    Comprobar "Macro: la fecha es una fecha de verdad", IsDate(ws.Range("H2").Value)
    ExportarComentarios
    Comprobar "Macro: correrla dos veces no duplica", ThisWorkbook.Worksheets("Comentarios").ListObjects(1).ListRows.Count = 9

    informe = informe & vbLf & IIf(fallas = 0, "TODO OK", fallas & " FALLA(S)")
    Debug.Print informe
    ProbarExtractor = informe
End Function

Private Sub Comprobar(nombre As String, ok As Boolean)
    If ok Then
        informe = informe & "OK     " & nombre & vbLf
    Else
        informe = informe & "FALLA  " & nombre & vbLf
        fallas = fallas + 1
    End If
End Sub

Private Function Contar(t As Variant, celda As String) As Long
    Dim i As Long
    For i = 2 To UBound(t, 1)
        If t(i, 4) = celda Then Contar = Contar + 1
    Next i
End Function
