VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15870
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   15870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Modificar 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   10
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton Eliminar 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   9
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton Agregar 
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   8
      Top             =   8280
      Width           =   2535
   End
   Begin MSComctlLib.ListView Books_list 
      Height          =   6615
      Left            =   3720
      TabIndex        =   7
      Top             =   1440
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11668
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Categories 
      Caption         =   "Categorías"
      Height          =   8175
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   2895
      Begin VB.CommandButton Favorites 
         Caption         =   "Favoritos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   6
         Top             =   5520
         Width           =   1935
      End
      Begin VB.CommandButton Want_to_read 
         Caption         =   "Quiero leer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   5
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CommandButton Read 
         Caption         =   "Leído"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton dislike 
         Caption         =   "No te gustaron"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   3
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton Recommended 
         Caption         =   "Recomendados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   2
         Top             =   6840
         Width           =   1935
      End
      Begin VB.CommandButton Mega_catalog 
         Caption         =   "Catálogo Mega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      Caption         =   "MEGA Libros"
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
      Left            =   5880
      TabIndex        =   11
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Agregar_Click()
    frmAgregarLibro.Show vbModal
End Sub

Private Sub dislike_Click()
     Dim rs As ADODB.Recordset
    Set rs = ObtenerLibrosPorEstado(1, "Nogusto")

    Books_list.ListItems.Clear

    Do Until rs.EOF
        With Books_list.ListItems.Add
            .Text = rs("Titulo")
            .SubItems(1) = rs("Autor")
            .Tag = rs("Id") ' ? Aquí se guarda el Id del libro
        End With
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Eliminar_Click()
Dim libroId As Integer

    ' Validar que haya un ítem seleccionado
    If Books_list.SelectedItem Is Nothing Then
        MsgBox "Por favor selecciona un libro para eliminar.", vbExclamation
        Exit Sub
    End If

    libroId = CInt(Books_list.SelectedItem.Tag)

    ' Confirmación
    If MsgBox("¿Estás seguro de que deseas eliminar este libro?", vbYesNo + vbQuestion, "Confirmar eliminación") = vbNo Then
        Exit Sub
    End If

    ' Conexión
    If conn Is Nothing Then Call ConectarBase
    If conn.State = adStateClosed Then conn.Open

    ' Eliminar en orden correcto (dependencias primero)
    conn.Execute "DELETE FROM ListasDeLecturaEstados WHERE ListaLecturaId IN (SELECT Id FROM ListasDeLectura WHERE LibroId = " & libroId & ")"
    conn.Execute "DELETE FROM ListasDeLectura WHERE LibroId = " & libroId
    conn.Execute "DELETE FROM Libros WHERE Id = " & libroId

    ' Quitar del ListView
    Books_list.ListItems.Remove Books_list.SelectedItem.Index

    MsgBox "Libro eliminado correctamente.", vbInformation
End Sub

Private Sub Favorites_Click()
     Dim rs As ADODB.Recordset
    Set rs = ObtenerLibrosPorEstado(1, "Favorito")

    Books_list.ListItems.Clear

    Do Until rs.EOF
        With Books_list.ListItems.Add
            .Text = rs("Titulo")
            .SubItems(1) = rs("Autor")
            .Tag = rs("Id") ' ? Aquí se guarda el Id del libro
        End With
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Form_Load()
    Call ConectarBase

    With Books_list
        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Título", 4000
        .ColumnHeaders.Add , , "Autor", 3000
    End With
End Sub

Private Sub Mega_catalog_Click()
    Dim rs As ADODB.Recordset
    Set rs = ObtenerCatalogo '

    Books_list.ListItems.Clear

   Do Until rs.EOF
    With Books_list.ListItems.Add
        .Text = rs("Titulo")
        If Not IsNull(rs("Autor")) Then
            .SubItems(1) = rs("Autor")
        Else
            .SubItems(1) = "Desconocido"
        End If
        .Tag = rs("Id")
    End With
    rs.MoveNext
Loop

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Modificar_Click()
    If Books_list.SelectedItem Is Nothing Then
        MsgBox "Selecciona un libro de la lista.", vbExclamation
        Exit Sub
    End If

    Dim idLibro As Integer
    idLibro = Books_list.SelectedItem.Tag ' Asegúrate que Tag guarde el Id

    frmEditarLibro.CargarDatos idLibro, 1 ' ? Ajusta con el ID real del usuario
    frmEditarLibro.Show vbModal
End Sub

Private Sub Read_Click()
     Dim rs As ADODB.Recordset
    Set rs = ObtenerLibrosPorEstado(1, "Leido")

    Books_list.ListItems.Clear

    Do Until rs.EOF
        With Books_list.ListItems.Add
            .Text = rs("Titulo")
            .SubItems(1) = rs("Autor")
            .Tag = rs("Id") ' ? Aquí se guarda el Id del libro
        End With
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Recommended_Click()
    Dim rs As ADODB.Recordset
    Set rs = ObtenerRecomendados(1)

    Books_list.ListItems.Clear

    Do Until rs.EOF
        With Books_list.ListItems.Add
            .Text = rs("Titulo")
            .SubItems(1) = rs("Autor")
            .Tag = rs("Id")
        End With
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Want_to_read_Click()
     Dim rs As ADODB.Recordset
    Set rs = ObtenerLibrosPorEstado(1, "PorLeer")

    Books_list.ListItems.Clear

    Do Until rs.EOF
        With Books_list.ListItems.Add
            .Text = rs("Titulo")
            .SubItems(1) = rs("Autor")
            .Tag = rs("Id") ' ? Aquí se guarda el Id del libro
        End With
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub



Private Sub MostrarLibrosPorEstado(estado As String)
    Dim rs As ADODB.Recordset
    Set rs = ObtenerLibrosPorEstado(1, estado)

    Books_list.ListItems.Clear

    Do Until rs.EOF
        With Books_list.ListItems.Add
            .Text = rs("Titulo")
            If Not IsNull(rs("Autor")) Then .SubItems(1) = rs("Autor")
        End With
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub
