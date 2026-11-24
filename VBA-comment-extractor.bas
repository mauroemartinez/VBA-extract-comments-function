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
' Macro:
'   ExportarComentarios   escribe todos los comentarios del libro en la hoja
'                         "Comentarios", como tabla de Excel (cualquier version)
'   Alias en ingles: ExportComments (hoja "Comments")
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
' Macro: todos los comentarios de todas las hojas del libro activo, en una hoja
' nueva llamada "Comentarios" (si existe, se reemplaza), como tabla de Excel
' lista para una tabla dinamica. Funciona en cualquier version de Excel.
' -----------------------------------------------------------------------------
Public Sub ExportarComentarios()
    ExportarALaHoja False
End Sub

Public Sub ExportComments()
    ExportarALaHoja True
End Sub

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

Private Sub ExportarALaHoja(ingles As Boolean)
    Dim wb As Workbook, ws As Worksheet, destino As Worksheet
    Dim todas As New Collection, filas As Collection, f As Variant
    Dim nombre As String, datos As Variant, n As Long
    Dim calculo As XlCalculation

    Set wb = ActiveWorkbook
    nombre = IIf(ingles, "Comments", "Comentarios")

    For Each ws In wb.Worksheets
        If ws.Name <> nombre Then
            Set filas = FilasDeHoja(ws, Nothing, ingles)
            For Each f In filas
                todas.Add f
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
    Else
        destino.Cells.Clear
        Do While destino.ListObjects.Count > 0
            destino.ListObjects(1).Unlist
        Loop
        destino.Cells.Clear
    End If

    datos = Matriz(todas, ingles)
    n = UBound(datos, 1)
    ' Todo en una sola escritura, no celda por celda: asi es rapido con libros grandes.
    destino.Range("A1").Resize(n, COLUMNAS).Value = datos

    With destino.ListObjects.Add(xlSrcRange, destino.Range("A1").Resize(n, COLUMNAS), , xlYes)
        .Name = IIf(ingles, "tblComments", "tblComentarios")
        .TableStyle = "TableStyleMedium2"
    End With
    destino.Columns(8).NumberFormat = "yyyy-mm-dd hh:mm"
    destino.Columns(10).ColumnWidth = 60
    destino.Columns(10).WrapText = True
    destino.Range("A:I").EntireColumn.AutoFit
    destino.Rows.AutoFit
    destino.Activate
    destino.Range("A1").Select

    Application.Calculation = calculo
    Application.ScreenUpdating = True
End Sub

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
