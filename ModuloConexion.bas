Attribute VB_Name = "ModuloConexion"
Public conn As ADODB.Connection

Public Sub ConectarBase()
     Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=PC_Reymundo;Initial Catalog=HubLectura;Integrated Security=SSPI;"
    conn.Open ConnectionString

End Sub
