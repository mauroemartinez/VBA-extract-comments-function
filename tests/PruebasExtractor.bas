Attribute VB_Name = "PruebasExtractor"
' Pruebas del extractor de comentarios, sobre el libro de ejemplo
' (example/comments-example.xlsm). Ejecutar la macro ProbarExtractor:
' devuelve un informe y lo muestra en la ventana Inmediato (Ctrl+G).
Option Explicit

Private informe As String
Private fallas As Long
Private paso As String

Public Function ProbarExtractor() As String
    Dim t As Variant, v As Worksheet, ws As Worksheet, n As Long, i As Long
    informe = "": fallas = 0
    On Error GoTo Fallo
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

    paso = "4 filtros"
    ' 4. Filtros por hoja y por fecha
    ExportarComentariosDe "Compras"
    Comprobar "Filtro por hoja: solo Compras (2 filas)", FilasDeLaTabla() = 2
    ExportarComentariosDe "Ventas; compras"
    Comprobar "Filtro con dos hojas, sin importar mayusculas (9 filas)", FilasDeLaTabla() = 9
    ExportarComentariosDe , Date - 1
    Comprobar "Filtro desde ayer: los 7 hilos, sin las notas", FilasDeLaTabla() = 7
    ExportarComentariosDe , Date + 1
    Comprobar "Filtro desde manana: tabla vacia, sin error", FilasDeLaTabla() = 0
    ExportarComentariosDe , , Date
    Comprobar "Filtro hasta hoy incluye todo el dia", FilasDeLaTabla() = 7

    paso = "5 dinamica"
    ' 5. La tabla dinamica sigue a la tabla
    ExportarComentarios
    Comprobar "La tabla sigue llamandose tblComentarios", _
        ThisWorkbook.Worksheets("Comentarios").ListObjects(1).Name = "tblComentarios"
    If HojaExiste("Resumen") Then
        Comprobar "Tabla dinamica actualizada (9 registros)", _
            ThisWorkbook.Worksheets("Resumen").PivotTables(1).PivotCache.RecordCount = 9
    End If

    paso = "6 crear hoja"
    ' 6. Importar: de la tabla a las celdas, en una hoja de prueba
    Dim p As Worksheet, lo As ListObject, r As ListRow, cambios As Long
    Application.DisplayAlerts = False
    If HojaExiste("_Prueba") Then ThisWorkbook.Worksheets("_Prueba").Delete
    Application.DisplayAlerts = True
    Set p = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    p.Name = "_Prueba"
    p.Range("B2").AddCommentThreaded("Hola").AddReply "Respuesta"
    p.Range("B3").AddComment "Nota vieja"
    p.Range("B4").AddComment "Para borrar"
    paso = "6 exportar prueba"
    ExportarComentariosDe "_Prueba"
    Set lo = ThisWorkbook.Worksheets("Comentarios").ListObjects(1)
    Comprobar "Importar: la hoja de prueba exporta 4 filas", lo.ListRows.Count = 4
    Comprobar "Importar sin cambios no toca nada", ImportarComentariosContar() = 0
    For Each r In lo.ListRows
        If r.Range(1, 4).Value = "B3" Then r.Range(1, 10).Value = "Nota nueva"
        If r.Range(1, 4).Value = "B4" Then r.Range(1, 10).Value = ""
        If r.Range(1, 4).Value = "B2" And r.Range(1, 5).Value = 1 Then r.Range(1, 9).Value = True
    Next r
    Set r = lo.ListRows.Add
    r.Range(1, 1).Value = "_Prueba"
    r.Range(1, 4).Value = "C5"
    r.Range(1, 5).Value = 1
    r.Range(1, 6).Value = "Hilo"
    r.Range(1, 10).Value = "Nuevo hilo"
    paso = "6 importar"
    cambios = ImportarComentariosContar()
    Comprobar "Importar: 4 celdas cambiadas", cambios = 4
    paso = "6 chequeos"
    Comprobar "Importar: texto editado", ExtraerComentarios(p.Range("B3"), "{text}") = "Nota nueva"
    Comprobar "Importar: comentario borrado", p.Range("B4").Comment Is Nothing
    paso = "6 hilo nuevo"
    Comprobar "Importar: hilo nuevo creado", p.Range("C5").CommentThreaded.Text = "Nuevo hilo"
    Comprobar "Importar: hilo marcado como resuelto", p.Range("B2").CommentThreaded.Resolved = True
    Comprobar "Importar: la respuesta del hilo sigue", p.Range("B2").CommentThreaded.Replies.Count = 1
    Comprobar "Importar dos veces: la segunda no cambia nada", ImportarComentariosContar() = 0
    paso = "6 ida y vuelta"
    ExportarComentarios
    Comprobar "Ida y vuelta de todo el libro: no cambia nada", ImportarComentariosContar() = 0
    Comprobar "Ida y vuelta: el hilo de D2 conserva sus 3 comentarios", _
        UBound(Split(ExtraerComentarios(v.Range("D2")), vbLf)) = 2

    Application.DisplayAlerts = False
    p.Delete
    Application.DisplayAlerts = True
    ExportarComentarios
    Comprobar "Sin la hoja de prueba vuelven a ser 9 filas", FilasDeLaTabla() = 9

    informe = informe & vbLf & IIf(fallas = 0, "TODO OK", fallas & " FALLA(S)")
    Debug.Print informe
    ProbarExtractor = informe
    Exit Function
Fallo:
    ProbarExtractor = informe & vbLf & "ERROR en el paso [" & paso & "]: " & Err.Number & " " & Err.Description
End Function

Private Sub Comprobar(nombre As String, ok As Boolean)
    If ok Then
        informe = informe & "OK     " & nombre & vbLf
    Else
        informe = informe & "FALLA  " & nombre & vbLf
        fallas = fallas + 1
    End If
End Sub

' Filas con datos en la tabla de la hoja Comentarios (0 si quedo vacia).
Private Function FilasDeLaTabla() As Long
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("Comentarios").ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Function
    If lo.ListRows.Count = 1 And lo.DataBodyRange.Cells(1, 1).Value = "" Then Exit Function
    FilasDeLaTabla = lo.ListRows.Count
End Function

Private Function HojaExiste(nombre As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nombre Then HojaExiste = True
    Next ws
End Function

Private Function Contar(t As Variant, celda As String) As Long
    Dim i As Long
    For i = 2 To UBound(t, 1)
        If t(i, 4) = celda Then Contar = Contar + 1
    Next i
End Function
