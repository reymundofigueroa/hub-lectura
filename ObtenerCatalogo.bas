Attribute VB_Name = "ModuloObtenerCatalogo"
Public Function ObtenerCatalogo() As ADODB.Recordset
    If conn Is Nothing Then Call ConectarBase
    If conn.State = adStateClosed Then conn.Open

    Dim rs As New ADODB.Recordset
    rs.Open "SELECT Id, Titulo, Autor FROM Libros", conn, adOpenStatic, adLockReadOnly
    Set ObtenerCatalogo = rs
End Function
