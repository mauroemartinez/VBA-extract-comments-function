Attribute VB_Name = "ExtractorComentarios"
' =============================================================================
' Extractor de comentarios de Excel
' https://github.com/mauroemartinez/VBA-extract-comments-function
' Licencia MIT
'
' Funciones de hoja:
'   =ExtraerComentarios(celda, [formato], [separador])   texto de una celda
'   =TablaComentarios([rango])                            tabla que se derrama (Excel 365)
'   Alias en ingles: =ExtractComments(...)  y  =CommentsTable(...)
'
' Macros (cualquier version de Excel):
'   ExportarComentarios            todos los comentarios del libro en la hoja
'                                  "Comentarios", como tabla de Excel
'   ExportarComentariosConFiltro   lo mismo, preguntando hojas y fechas
'   ImportarComentarios            el camino de vuelta: lo que se edito en la
'                                  tabla se escribe en las celdas
'   Para llamar desde otra macro: ExportarComentariosDe [hojas], [desde], [hasta]
'   Alias en ingles: ExportComments, ExportCommentsWithFilter, ImportComments,
'   ExportCommentsFrom (hoja "Comments")
'
' Actualizar: Excel no recalcula una formula cuando cambia un comentario.
'   Despues de agregar, editar o borrar comentarios, apreta Ctrl+Alt+F9.
'   (Las funciones no son volatiles a proposito: una funcion volatil que lee
'   comentarios mientras Excel abre el libro puede hacer caer a Excel.)
'
' Columnas de la tabla:
'   Hoja | Columna | Fila | Celda | N comentario | Tipo | Autor | Fecha | Resuelto | Texto
'   (N comentario: 1 es el inicio del hilo, 2 en adelante son las respuestas)
' =============================================================================
Option Explicit

Private Const COLUMNAS As Long = 10

' -----------------------------------------------------------------------------
' Texto de los comentarios de una celda, en una sola celda.
'   formato: plantilla opcional para cada comentario, con {author}, {date},
'            {text} y {n}. Por ejemplo: "{author} ({date}): {text}"
'            Sin formato, se mantiene la salida original:
'            "Autor el fecha: texto" (hilos) y "Autor: texto" (notas).
'   separador: lo que va entre un comentario y el siguiente. Por defecto, un
'              salto de linea.
' Si la celda no tiene comentarios, devuelve "" (celda vacia, sin error).
' -----------------------------------------------------------------------------
Public Function ExtraerComentarios(celda As Range, Optional formato As String = "", _
                                   Optional separador As String = vbLf) As String
    ExtraerComentarios = TextoDeCelda(celda.Cells(1, 1), formato, separador)
End Function

Public Function ExtractComments(cell As Range, Optional format As String = "", _
                                Optional separator As String = vbLf) As String
    ExtractComments = TextoDeCelda(cell.Cells(1, 1), format, separator)
End Function

' -----------------------------------------------------------------------------
' Tabla de comentarios que se derrama (matriz dinamica, Excel 365).
'   rango: opcional. Sin rango, toma toda la hoja de la celda donde esta la
'          formula. Con un rango, solo los comentarios que caen adentro.
' Una fila por comentario o respuesta, con encabezados.
' -----------------------------------------------------------------------------
Public Function TablaComentarios(Optional rango As Range) As Variant
    TablaComentarios = TablaParaFuncion(rango, False)
End Function

Public Function CommentsTable(Optional target As Range) As Variant
    CommentsTable = TablaParaFuncion(target, True)
End Function

' -----------------------------------------------------------------------------
' Macro: todos los comentarios de todas las hojas del libro activo, en la hoja
' "Comentarios", como tabla de Excel (tblComentarios) lista para una tabla
' dinamica. Si la hoja ya existe, la tabla se reescribe (nunca se duplica) y
' las tablas dinamicas que salen de ella se actualizan solas.
' -----------------------------------------------------------------------------
Public Sub ExportarComentarios()
    ExportarALaHoja False, "", Empty, Empty
End Sub

Public Sub ExportComments()
    ExportarALaHoja True, "", Empty, Empty
End Sub

' -----------------------------------------------------------------------------
' Igual, pero solo algunas hojas o algunas fechas. Para llamar desde otra macro
' o desde la ventana Inmediato:
'   ExportarComentariosDe "Ventas;Compras"
'   ExportarComentariosDe , #11/1/2026#, #11/30/2026#
'   hojas: nombres separados por ; (vacio = todas)
'   desde / hasta: fechas incluidas. Las notas no tienen fecha en Excel, asi
'   que con un filtro de fechas quedan afuera.
' -----------------------------------------------------------------------------
Public Sub ExportarComentariosDe(Optional hojas As String = "", Optional desde As Variant, _
                                 Optional hasta As Variant)
    ExportarALaHoja False, hojas, FechaOVacio(desde), FechaOVacio(hasta)
End Sub

Public Sub ExportCommentsFrom(Optional sheets As String = "", Optional fromDate As Variant, _
                              Optional toDate As Variant)
    ExportarALaHoja True, sheets, FechaOVacio(fromDate), FechaOVacio(toDate)
End Sub

' Macro con preguntas, para correr desde Alt + F8.
Public Sub ExportarComentariosConFiltro()
    PreguntarYExportar False
End Sub

Public Sub ExportCommentsWithFilter()
    PreguntarYExportar True
End Sub

' -----------------------------------------------------------------------------
' Macro: el camino de vuelta. Lee la tabla de la hoja "Comentarios" y deja los
' comentarios de las celdas como dice la tabla:
'   - se cambia el texto de una fila: se reescribe ese comentario;
'   - se agrega una fila (Hoja, Celda, N comentario, Tipo y Texto): se crea;
'   - se borra el texto de todas las filas de una celda: se borra el comentario;
'   - se cambia Resuelto: se marca o desmarca el hilo.
' Las celdas que no cambiaron no se tocan (conservan autor y fecha). Las celdas
' que no figuran en la tabla tampoco se tocan.
' Lo que Excel no deja cambiar: el autor y la fecha de un comentario en hilo
' (los nuevos quedan a nombre de quien corre la macro, con la fecha de hoy).
' -----------------------------------------------------------------------------
Public Sub ImportarComentarios()
    Dim n As Long
    n = ImportarDesdeLaHoja(False)
    Application.StatusBar = "ImportarComentarios: " & n & " celda(s) actualizada(s)."
End Sub

Public Sub ImportComments()
    Dim n As Long
    n = ImportarDesdeLaHoja(True)
    Application.StatusBar = "ImportComments: " & n & " cell(s) updated."
End Sub

' Para pruebas y para otras macros: devuelve cuantas celdas cambio.
Public Function ImportarComentariosContar() As Long
    ImportarComentariosContar = ImportarDesdeLaHoja(False)
End Function

' =============================================================================
' Lo de adentro
' =============================================================================

Private Function TablaParaFuncion(rango As Range, ingles As Boolean) As Variant
    Dim ws As Worksheet, filas As Collection
    On Error GoTo Fallo
    If rango Is Nothing Then
        Set ws = Application.Caller.Worksheet
        Set filas = FilasDeHoja(ws, Nothing, ingles)
    Else
        Set ws = rango.Worksheet
        Set filas = FilasDeHoja(ws, rango, ingles)
    End If
    TablaParaFuncion = Matriz(filas, ingles)
    Exit Function
Fallo:
    TablaParaFuncion = CVErr(xlErrValue)
End Function

Private Sub ExportarALaHoja(ingles As Boolean, hojas As String, desde As Variant, hasta As Variant)
    Dim wb As Workbook, ws As Worksheet, destino As Worksheet, lo As ListObject
    Dim todas As New Collection, filas As Collection, f As Variant
    Dim nombre As String, nombreTabla As String, datos As Variant, n As Long
    Dim calculo As XlCalculation, pc As PivotCache

    Set wb = ActiveWorkbook
    nombre = IIf(ingles, "Comments", "Comentarios")
    nombreTabla = IIf(ingles, "tblComments", "tblComentarios")

    For Each ws In wb.Worksheets
        If ws.Name <> nombre And HojaElegida(ws.Name, hojas) Then
            Set filas = FilasDeHoja(ws, Nothing, ingles)
            For Each f In filas
                If FechaElegida(f(8), desde, hasta) Then todas.Add f
            Next f
        End If
    Next ws

    calculo = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    Set destino = wb.Worksheets(nombre)
    On Error GoTo 0
    If destino Is Nothing Then
        Set destino = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        destino.Name = nombre
    End If

    ' Si la tabla ya existe, se reescribe en el mismo lugar y con el mismo nombre:
    ' asi las tablas dinamicas y las formulas que la usan no se rompen.
    On Error Resume Next
    Set lo = destino.ListObjects(nombreTabla)
    On Error GoTo 0
    Dim filasViejas As Long
    If Not lo Is Nothing Then
        ' Se vacia sin borrar filas: borrar filas deja la referencia a la tabla invalida.
        filasViejas = lo.Range.Rows.Count
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    Else
        destino.Cells.Clear
    End If

    datos = Matriz(todas, ingles)
    n = UBound(datos, 1)
    ' Todo en una sola escritura, no celda por celda: asi es rapido con libros grandes.
    destino.Range("A1").Resize(n, COLUMNAS).Value = datos

    ' Una tabla de Excel necesita al menos una fila de datos: si no hay
    ' comentarios, queda una fila vacia.
    If n < 2 Then n = 2
    ' Escribir sobre los encabezados invalida la referencia a la tabla: se vuelve a buscar.
    If Not lo Is Nothing Then
        Set lo = Nothing
        On Error Resume Next
        Set lo = destino.ListObjects(nombreTabla)
        On Error GoTo 0
    End If
    If lo Is Nothing Then
        Set lo = destino.ListObjects.Add(xlSrcRange, destino.Range("A1").Resize(n, COLUMNAS), , xlYes)
        lo.Name = nombreTabla
        lo.TableStyle = "TableStyleMedium2"
    Else
        lo.Resize destino.Range("A1").Resize(n, COLUMNAS)
        ' Si la tabla se achico, se limpia lo que quedo debajo.
        If filasViejas > n Then destino.Range("A1").Offset(n).Resize(filasViejas - n, COLUMNAS).Clear
    End If
    destino.Columns(8).NumberFormat = "yyyy-mm-dd hh:mm"
    destino.Columns(10).ColumnWidth = 60
    destino.Columns(10).WrapText = True
    destino.Range("A:I").EntireColumn.AutoFit
    destino.Rows.AutoFit

    ' Las tablas dinamicas del libro se actualizan con la tabla nueva.
    On Error Resume Next
    For Each pc In wb.PivotCaches
        pc.MissingItemsLimit = xlMissingItemsNone
        pc.Refresh
    Next pc
    On Error GoTo 0

    destino.Activate
    destino.Range("A1").Select

    Application.Calculation = calculo
    Application.ScreenUpdating = True
End Sub

Private Sub PreguntarYExportar(ingles As Boolean)
    Dim hojas As Variant, desde As Variant, hasta As Variant
    hojas = Application.InputBox(IIf(ingles, "Sheets separated by ; (empty = all):", _
                                     "Hojas separadas por ; (vacio = todas):"), _
                                 IIf(ingles, "Export comments", "Exportar comentarios"), "", Type:=2)
    If VarType(hojas) = vbBoolean Then Exit Sub
    desde = Application.InputBox(IIf(ingles, "From date (empty = no limit):", _
                                     "Desde la fecha (vacio = sin limite):"), , "", Type:=2)
    If VarType(desde) = vbBoolean Then Exit Sub
    hasta = Application.InputBox(IIf(ingles, "To date (empty = no limit):", _
                                     "Hasta la fecha (vacio = sin limite):"), , "", Type:=2)
    If VarType(hasta) = vbBoolean Then Exit Sub
    ExportarALaHoja ingles, CStr(hojas), FechaOVacio(desde), FechaOVacio(hasta)
End Sub

Private Function FechaOVacio(v As Variant) As Variant
    FechaOVacio = Empty
    If IsMissing(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim$(v) = "" Then Exit Function
    End If
    If IsDate(v) Then FechaOVacio = CDate(v)
End Function

Private Function HojaElegida(nombre As String, hojas As String) As Boolean
    Dim h As Variant
    If Trim$(hojas) = "" Then
        HojaElegida = True
        Exit Function
    End If
    For Each h In Split(hojas, ";")
        If StrComp(Trim$(h), nombre, vbTextCompare) = 0 Then
            HojaElegida = True
            Exit Function
        End If
    Next h
End Function

' "hasta" incluye todo ese dia: una fecha sin hora vale hasta las 23:59:59.
Private Function FechaElegida(fecha As Variant, desde As Variant, hasta As Variant) As Boolean
    If IsEmpty(desde) And IsEmpty(hasta) Then
        FechaElegida = True
    ElseIf IsEmpty(fecha) Then
        FechaElegida = False
    Else
        FechaElegida = True
        If Not IsEmpty(desde) Then If fecha < desde Then FechaElegida = False
        If Not IsEmpty(hasta) Then
            If hasta = Int(hasta) Then
                If fecha >= hasta + 1 Then FechaElegida = False
            Else
                If fecha > hasta Then FechaElegida = False
            End If
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' Importar: de la tabla a las celdas
' -----------------------------------------------------------------------------
Private Function ImportarDesdeLaHoja(ingles As Boolean) As Long
    Dim wb As Workbook, origen As Worksheet, lo As ListObject, datos As Variant
    Dim celdas As New Collection, claves As New Collection, clave As String
    Dim i As Long, k As Variant, grupo As Collection, cambios As Long
    Dim nombre As String

    Set wb = ActiveWorkbook
    nombre = IIf(ingles, "Comments", "Comentarios")
    On Error Resume Next
    Set origen = wb.Worksheets(nombre)
    On Error GoTo 0
    If origen Is Nothing Then Err.Raise vbObjectError + 1, , "No existe la hoja " & nombre & ". Corre primero la macro de exportar."

    If origen.ListObjects.Count > 0 Then
        Set lo = origen.ListObjects(1)
        If lo.DataBodyRange Is Nothing Then Exit Function
        datos = lo.DataBodyRange.Value
    Else
        datos = origen.Range("A1").CurrentRegion.Offset(1).Resize(origen.Range("A1").CurrentRegion.Rows.Count - 1).Value
    End If

    ' Agrupar las filas por Hoja + Celda.
    For i = 1 To UBound(datos, 1)
        If Trim$(CStr(datos(i, 1))) <> "" And Trim$(CStr(datos(i, 4))) <> "" Then
            clave = UCase$(Trim$(CStr(datos(i, 1)))) & "!" & UCase$(Trim$(CStr(datos(i, 4))))
            Set grupo = Nothing
            On Error Resume Next
            Set grupo = celdas(clave)
            On Error GoTo 0
            If grupo Is Nothing Then
                Set grupo = New Collection
                celdas.Add grupo, clave
                claves.Add clave
            End If
            grupo.Add i
        End If
    Next i

    Application.ScreenUpdating = False
    For Each k In claves
        If AplicarCelda(wb, datos, celdas(k)) Then cambios = cambios + 1
    Next k
    Application.ScreenUpdating = True
    ImportarDesdeLaHoja = cambios
End Function

' Deja una celda como dicen sus filas. Devuelve True si cambio algo.
Private Function AplicarCelda(wb As Workbook, datos As Variant, grupo As Collection) As Boolean
    Dim ws As Worksheet, c As Range, idx() As Long, i As Long, j As Long, t As Long
    Dim textos As New Collection, esHilo As Boolean, resuelto As Variant
    Dim actual As String, deseado As String, tipo As String
    Dim hilo As Object, existeHilo As Boolean, existeNota As Boolean

    ' Filas de la celda, ordenadas por N comentario.
    ReDim idx(1 To grupo.Count)
    For i = 1 To grupo.Count
        idx(i) = grupo(i)
    Next i
    For i = 2 To grupo.Count
        t = idx(i): j = i - 1
        Do While j >= 1
            If Val(datos(idx(j), 5)) <= Val(datos(t, 5)) Then Exit Do
            idx(j + 1) = idx(j): j = j - 1
        Loop
        idx(j + 1) = t
    Next i

    On Error Resume Next
    Set ws = wb.Worksheets(Trim$(CStr(datos(idx(1), 1))))
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Set c = ws.Range(Trim$(CStr(datos(idx(1), 4))))

    tipo = LCase$(Trim$(CStr(datos(idx(1), 6))))
    esHilo = (tipo = "hilo" Or tipo = "threaded")
    resuelto = datos(idx(1), 9)
    For i = 1 To grupo.Count
        If Trim$(CStr(datos(idx(i), 10))) <> "" Then textos.Add CStr(datos(idx(i), 10))
    Next i

    ' Como esta hoy la celda.
    On Error Resume Next
    Set hilo = c.CommentThreaded
    On Error GoTo 0
    existeHilo = Not (hilo Is Nothing)
    existeNota = (Not existeHilo) And Not (c.Comment Is Nothing)
    actual = FirmaActual(c, existeHilo, existeNota)
    deseado = FirmaDeseada(textos, esHilo)

    If actual = deseado Then
        ' Mismo contenido: solo puede cambiar Resuelto.
        If existeHilo And VarType(resuelto) = vbBoolean Then
            On Error Resume Next
            If hilo.Resolved <> resuelto Then
                hilo.Resolved = resuelto
                AplicarCelda = True
            End If
            On Error GoTo 0
        End If
        Exit Function
    End If

    ' Distinto: se borra lo que hay y se escribe lo de la tabla.
    If existeHilo Then hilo.Delete
    If Not c.Comment Is Nothing Then c.Comment.Delete
    If textos.Count > 0 Then
        If esHilo Then
            Set hilo = c.AddCommentThreaded(textos(1))
            For i = 2 To textos.Count
                hilo.AddReply textos(i)
            Next i
            If VarType(resuelto) = vbBoolean Then
                On Error Resume Next
                hilo.Resolved = resuelto
                On Error GoTo 0
            End If
        Else
            deseado = textos(1)
            For i = 2 To textos.Count
                deseado = deseado & vbLf & textos(i)
            Next i
            c.AddComment deseado
        End If
    End If
    AplicarCelda = True
End Function

Private Function FirmaActual(c As Range, existeHilo As Boolean, existeNota As Boolean) As String
    Dim s As String, r As Object
    If existeHilo Then
        s = "H|" & c.CommentThreaded.Text
        For Each r In c.CommentThreaded.Replies
            s = s & "|" & r.Text
        Next r
    ElseIf existeNota Then
        s = "N|" & TextoDeNota(c.Comment)
    End If
    FirmaActual = s
End Function

Private Function FirmaDeseada(textos As Collection, esHilo As Boolean) As String
    Dim s As String, i As Long
    If textos.Count = 0 Then Exit Function
    If esHilo Then
        s = "H"
        For i = 1 To textos.Count
            s = s & "|" & textos(i)
        Next i
    Else
        s = "N|" & textos(1)
        For i = 2 To textos.Count
            s = s & vbLf & textos(i)
        Next i
    End If
    FirmaDeseada = s
End Function

' Una fila (matriz de 10 valores) por comentario o respuesta, ordenadas por
' fila, columna y numero de comentario.
Private Function FilasDeHoja(ws As Worksheet, rango As Range, ingles As Boolean) As Collection
    Dim filas As New Collection
    Dim hilo As Object, resp As Object, nota As Comment, c As Range
    Dim k As Long, tieneHilo As Boolean

    ' Comentarios modernos (hilos). En versiones sin hilos, se saltea.
    On Error Resume Next
    Dim hilos As Object
    Set hilos = ws.CommentsThreaded
    On Error GoTo 0
    If Not hilos Is Nothing Then
        For Each hilo In hilos
            Set c = hilo.Parent
            If Dentro(c, rango) Then
                filas.Add Fila(c, 1, IIf(ingles, "Threaded", "Hilo"), hilo.Author.Name, _
                               hilo.Date, Resuelto(hilo), hilo.Text)
                k = 1
                For Each resp In hilo.Replies
                    k = k + 1
                    filas.Add Fila(c, k, IIf(ingles, "Threaded", "Hilo"), resp.Author.Name, _
                                   resp.Date, Empty, resp.Text)
                Next resp
            End If
        Next hilo
    End If

    ' Notas (comentarios clasicos). Excel tambien crea una nota "fantasma" para
    ' cada hilo: esas se saltean para no contar dos veces.
    For Each nota In ws.Comments
        Set c = nota.Parent
        If Dentro(c, rango) Then
            tieneHilo = False
            On Error Resume Next
            tieneHilo = Not (c.CommentThreaded Is Nothing)
            On Error GoTo 0
            If Not tieneHilo Then
                filas.Add Fila(c, 1, IIf(ingles, "Note", "Nota"), nota.Author, _
                               Empty, Empty, TextoDeNota(nota))
            End If
        End If
    Next nota

    Set FilasDeHoja = Ordenadas(filas)
End Function

Private Function Fila(c As Range, numero As Long, tipo As String, autor As String, _
                      fecha As Variant, resuelto As Variant, texto As String) As Variant
    Dim v(1 To COLUMNAS) As Variant
    v(1) = c.Worksheet.Name
    v(2) = Split(c.Address(True, False), "$")(0)
    v(3) = c.Row
    v(4) = c.Address(False, False)
    v(5) = numero
    v(6) = tipo
    v(7) = autor
    v(8) = fecha
    v(9) = resuelto
    v(10) = texto
    Fila = v
End Function

Private Function Matriz(filas As Collection, ingles As Boolean) As Variant
    Dim m() As Variant, i As Long, j As Long, f As Variant
    ReDim m(1 To filas.Count + 1, 1 To COLUMNAS)
    Dim titulos As Variant
    If ingles Then
        titulos = Array("Sheet", "Column", "Row", "Cell", "Comment #", "Type", "Author", "Timestamp", "Resolved", "Text")
    Else
        titulos = Array("Hoja", "Columna", "Fila", "Celda", "N comentario", "Tipo", "Autor", "Fecha", "Resuelto", "Texto")
    End If
    For j = 1 To COLUMNAS
        m(1, j) = titulos(j - 1)
    Next j
    i = 1
    For Each f In filas
        i = i + 1
        For j = 1 To COLUMNAS
            If IsEmpty(f(j)) Then m(i, j) = "" Else m(i, j) = f(j)
        Next j
    Next f
    Matriz = m
End Function

' Orden estable por hoja, fila, columna y numero de comentario.
Private Function Ordenadas(filas As Collection) As Collection
    Dim a() As Variant, i As Long, j As Long, t As Variant, n As Long
    Dim res As New Collection
    n = filas.Count
    If n = 0 Then
        Set Ordenadas = res
        Exit Function
    End If
    ReDim a(1 To n)
    For i = 1 To n
        a(i) = filas(i)
    Next i
    For i = 2 To n
        t = a(i)
        j = i - 1
        Do While j >= 1
            If Not Mayor(a(j), t) Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = t
    Next i
    For i = 1 To n
        res.Add a(i)
    Next i
    Set Ordenadas = res
End Function

Private Function Mayor(x As Variant, y As Variant) As Boolean
    If x(3) <> y(3) Then
        Mayor = x(3) > y(3)
    ElseIf ColumnaNumero(x(2)) <> ColumnaNumero(y(2)) Then
        Mayor = ColumnaNumero(x(2)) > ColumnaNumero(y(2))
    Else
        Mayor = x(5) > y(5)
    End If
End Function

Private Function ColumnaNumero(letra As Variant) As Long
    Dim i As Long, n As Long
    For i = 1 To Len(letra)
        n = n * 26 + (Asc(UCase(Mid(letra, i, 1))) - 64)
    Next i
    ColumnaNumero = n
End Function

Private Function Dentro(c As Range, rango As Range) As Boolean
    If rango Is Nothing Then
        Dentro = True
    ElseIf c.Worksheet.Name <> rango.Worksheet.Name Then
        Dentro = False
    Else
        Dentro = Not (Application.Intersect(c, rango) Is Nothing)
    End If
End Function

Private Function Resuelto(hilo As Object) As Variant
    Resuelto = Empty
    On Error Resume Next
    Resuelto = CBool(hilo.Resolved)
    On Error GoTo 0
End Function

' Las notas suelen empezar con "Autor:" y un salto de linea. Se saca para
' dejar solo el texto (el autor ya va en su columna).
Private Function TextoDeNota(nota As Comment) As String
    Dim t As String, a As String
    t = nota.Text
    a = nota.Author
    If Len(a) > 0 And Left$(t, Len(a) + 1) = a & ":" Then
        t = Mid$(t, Len(a) + 2)
        Do While Left$(t, 1) = vbLf Or Left$(t, 1) = vbCr
            t = Mid$(t, 2)
        Loop
    End If
    TextoDeNota = t
End Function

Private Function TextoDeCelda(c As Range, formato As String, separador As String) As String
    Dim hilo As Object, resp As Object, k As Long, salida As String
    Dim tieneHilo As Boolean

    On Error Resume Next
    Set hilo = c.CommentThreaded
    On Error GoTo 0
    tieneHilo = Not (hilo Is Nothing)

    If tieneHilo Then
        k = 1
        salida = Renglon(formato, hilo.Author.Name, hilo.Date, hilo.Text, k, True)
        For Each resp In hilo.Replies
            k = k + 1
            salida = salida & separador & Renglon(formato, resp.Author.Name, resp.Date, resp.Text, k, True)
        Next resp
        TextoDeCelda = salida
    ElseIf Not c.Comment Is Nothing Then
        If formato = "" Then
            ' Comportamiento original: si la nota ya empieza con el autor, se deja como esta.
            If InStr(Left(c.Comment.Text, Len(c.Comment.Author)), c.Comment.Author) > 0 Then
                TextoDeCelda = c.Comment.Text
            Else
                TextoDeCelda = c.Comment.Author & ": " & c.Comment.Text
            End If
        Else
            TextoDeCelda = Renglon(formato, c.Comment.Author, Empty, TextoDeNota(c.Comment), 1, False)
        End If
    Else
        TextoDeCelda = ""
    End If
End Function

Private Function Renglon(formato As String, autor As String, fecha As Variant, _
                         texto As String, n As Long, esHilo As Boolean) As String
    Dim s As String
    If formato = "" Then
        Renglon = autor & " el " & fecha & ": " & texto
        Exit Function
    End If
    s = Replace(formato, "{author}", autor)
    If IsEmpty(fecha) Then
        s = Replace(s, "{date}", "")
    Else
        s = Replace(s, "{date}", Format$(fecha, "yyyy-mm-dd hh:mm"))
    End If
    s = Replace(s, "{n}", CStr(n))
    s = Replace(s, "{text}", texto)
    Renglon = s
End Function
