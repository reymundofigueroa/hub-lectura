Attribute VB_Name = "ModuloCargarDatos"
Public Sub CargarDatos(idLibro As Integer, idUsuario As Integer)
    libroId = idLibro
    usuarioId = idUsuario

    ' Obtener info del libro
    If conn Is Nothing Then Call ConectarBase
    If conn.State = adStateClosed Then conn.Open

    Dim rs As New ADODB.Recordset
    rs.Open "SELECT Titulo, Autor FROM Libros WHERE Id = " & libroId, conn, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        txtTitulo.Text = rs("Titulo")
        txtAutor.Text = rs("Autor")
    End If
    rs.Close

    ' Obtener estados actuales
    Dim sqlEstados As String
    sqlEstados = "SELECT Estado FROM ListasDeLecturaEstados E " & _
                 "INNER JOIN ListasDeLectura LL ON E.ListaLecturaId = LL.Id " & _
                 "WHERE LL.LibroId = " & libroId & " AND LL.UsuarioId = " & usuarioId

    rs.Open sqlEstados, conn, adOpenStatic, adLockReadOnly
    chkLeido.Value = vbUnchecked
    chkFavorito.Value = vbUnchecked
    chkNoGusto.Value = vbUnchecked

    Do Until rs.EOF
        Select Case rs("Estado")
            Case "Leido": chkLeido.Value = vbChecked
            Case "Favorito": chkFavorito.Value = vbChecked
            Case "NoGusto": chkNoGusto.Value = vbChecked
        End Select
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub


