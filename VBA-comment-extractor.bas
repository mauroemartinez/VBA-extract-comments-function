Attribute VB_Name = "Módulo1"
Function ExtraerComentarios(celda As Range) As String
    Dim c As CommentThreaded

    ExtraerComentarios = ""  ' fallback en caso de que no haya comentarios
    If Not celda.Cells(1, 1).CommentThreaded Is Nothing Then
        ' comentario en el nuevo formato
        With celda.Cells(1, 1).CommentThreaded
            ExtraerComentarios = .Author.Name & " el " & .Date & ": " & .Text
            For Each c In .Replies
                ExtraerComentarios = ExtraerComentarios & vbCrLf & _
                    c.Author.Name & " el " & c.Date & ": " & c.Text
            Next c
        End With
    ElseIf Not celda.Cells(1, 1).Comment Is Nothing Then
        ' comentario en el formato antiguo
        With celda.Cells(1, 1).Comment
            ' adivina inteligentemente si el comentario comienza con el nombre del autor, si no, agrégalo
            If InStr(Left(.Text, Len(.Author)), .Author) > 0 Then
                ExtraerComentarios = .Text
            Else
                ExtraerComentarios = .Author & ": " & .Text
            End If
        End With
    End If
End Function


