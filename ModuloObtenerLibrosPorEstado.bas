Attribute VB_Name = "ModuloObtenerLibrosPorEstado"
Public Function ObtenerLibrosPorEstado(idUsuario As Integer, estado As String) As ADODB.Recordset
    If conn Is Nothing Then Call ConectarBase
    If conn.State = adStateClosed Then conn.Open

    Dim sql As String
    sql = "SELECT L.Titulo, L.Autor, L.UrlMega " & _
          "FROM Libros L " & _
          "INNER JOIN ListasDeLectura LL ON L.Id = LL.LibroId " & _
          "INNER JOIN ListasDeLecturaEstados E ON E.ListaLecturaId = LL.Id " & _
          "WHERE LL.UsuarioId = " & idUsuario & " AND E.Estado = '" & estado & "'"

    Dim rs As New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    Set ObtenerLibrosPorEstado = rs
End Function

Public Function ObtenerRecomendados(idUsuario As Integer) As ADODB.Recordset
    If conn Is Nothing Then Call ConectarBase
    If conn.State = adStateClosed Then conn.Open

    Dim sql As String
    sql = "SELECT DISTINCT L.Titulo, L.Autor, L.UrlMega " & _
          "FROM Libros L " & _
          "INNER JOIN Preferencias P ON L.GeneroId = P.GeneroId " & _
          "WHERE P.UsuarioId = " & idUsuario & " " & _
          "AND L.Id NOT IN ( " & _
              "SELECT LibroId FROM ListasDeLectura WHERE UsuarioId = " & idUsuario & _
          ")"

    Dim rs As New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    Set ObtenerRecomendados = rs
End Function
