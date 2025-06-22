VERSION 5.00
Begin VB.Form frmEditarLibro 
   Caption         =   "Form2"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7740
   LinkTopic       =   "Form2"
   ScaleHeight     =   8760
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   4320
      TabIndex        =   4
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton btnGuardar 
      Caption         =   "Guardar cambios"
      Height          =   735
      Left            =   1440
      TabIndex        =   3
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Editar libro"
      Height          =   5535
      Left            =   720
      TabIndex        =   0
      Top             =   1680
      Width           =   6375
      Begin VB.TextBox txtAutor 
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtTitulo 
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   960
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Height          =   2895
         Left            =   720
         TabIndex        =   5
         Top             =   1920
         Width           =   4815
         Begin VB.CheckBox chkNoGusto 
            Caption         =   "No me gustó"
            Height          =   255
            Left            =   600
            TabIndex        =   8
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CheckBox chkFavorito 
            Caption         =   "Favorito"
            Height          =   255
            Left            =   600
            TabIndex        =   7
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CheckBox chkLeido 
            Caption         =   "Leido"
            Height          =   375
            Left            =   600
            TabIndex        =   6
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Label lblAutor 
         Caption         =   "Autor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Título"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   1695
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Edita este libro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   11
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "frmEditarLibro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private libroId As Integer
Private usuarioId As Integer

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnGuardarCambios_Click()

End Sub

Private Sub Check1_Click()

End Sub

Private Sub Check2_Click()

End Sub

Private Sub btnGuardar_Click()
    If conn Is Nothing Then Call ConectarBase
    If conn.State = adStateClosed Then conn.Open

    Dim lecturaId As Integer
    Dim rs As New ADODB.Recordset

    ' Obtener o insertar ListasDeLectura
    rs.Open "SELECT Id FROM ListasDeLectura WHERE LibroId = " & libroId & " AND UsuarioId = " & usuarioId, conn, adOpenStatic, adLockOptimistic
    If rs.EOF Then
        ' Insertar
        conn.Execute "INSERT INTO ListasDeLectura (UsuarioId, LibroId) VALUES (" & usuarioId & ", " & libroId & ")"
        rs.Requery
    End If
    lecturaId = rs("Id")
    rs.Close

    ' Borrar estados anteriores
    conn.Execute "DELETE FROM ListasDeLecturaEstados WHERE ListaLecturaId = " & lecturaId

    ' Insertar nuevos estados
    If chkLeido.Value = vbChecked Then
        conn.Execute "INSERT INTO ListasDeLecturaEstados (ListaLecturaId, Estado) VALUES (" & lecturaId & ", 'Leido')"
    End If
    If chkFavorito.Value = vbChecked Then
        conn.Execute "INSERT INTO ListasDeLecturaEstados (ListaLecturaId, Estado) VALUES (" & lecturaId & ", 'Favorito')"
    End If
    If chkNoGusto.Value = vbChecked Then
        conn.Execute "INSERT INTO ListasDeLecturaEstados (ListaLecturaId, Estado) VALUES (" & lecturaId & ", 'NoGusto')"
    End If

    ' Si no marcó nada, agregar PorLeer
    If chkLeido.Value = vbUnchecked And chkFavorito.Value = vbUnchecked And chkNoGusto.Value = vbUnchecked Then
        conn.Execute "INSERT INTO ListasDeLecturaEstados (ListaLecturaId, Estado) VALUES (" & lecturaId & ", 'PorLeer')"
    End If

    MsgBox "Cambios guardados correctamente.", vbInformation
    Unload Me
End Sub

Public Sub CargarDatos(idLibro As Integer, idUsuario As Integer)
    libroId = idLibro
    usuarioId = idUsuario

    If conn Is Nothing Then Call ConectarBase
    If conn.State = adStateClosed Then conn.Open

    Dim rs As New ADODB.Recordset
    rs.Open "SELECT Titulo, Autor FROM Libros WHERE Id = " & libroId, conn, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        txtTitulo.Text = rs("Titulo")
        txtAutor.Text = rs("Autor")
    End If
    rs.Close

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


