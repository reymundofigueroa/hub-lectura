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
      Height          =   615
      Left            =   7200
      TabIndex        =   10
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton Eliminar 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   10200
      TabIndex        =   9
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton Agregar 
      Caption         =   "Agregar"
      Height          =   615
      Left            =   3960
      TabIndex        =   8
      Top             =   7800
      Width           =   2535
   End
   Begin MSComctlLib.ListView Books_list 
      Height          =   6615
      Left            =   4200
      TabIndex        =   7
      Top             =   960
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
      Height          =   9015
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   2895
      Begin VB.CommandButton Favorites 
         Caption         =   "Favoritos"
         Height          =   975
         Left            =   480
         TabIndex        =   6
         Top             =   6360
         Width           =   1935
      End
      Begin VB.CommandButton Want_to_read 
         Caption         =   "Quiero leer"
         Height          =   975
         Left            =   480
         TabIndex        =   5
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Read 
         Caption         =   "Leído"
         Height          =   975
         Left            =   480
         TabIndex        =   4
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton dislike 
         Caption         =   "No te gustaron"
         Height          =   975
         Left            =   480
         TabIndex        =   3
         Top             =   5040
         Width           =   1935
      End
      Begin VB.CommandButton Recommended 
         Caption         =   "Recomendados"
         Height          =   975
         Left            =   480
         TabIndex        =   2
         Top             =   7800
         Width           =   1935
      End
      Begin VB.CommandButton Mega_catalog 
         Caption         =   "Catálogo Mega"
         Height          =   975
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dislike_Click()
    MostrarLibrosPorEstado "NoGusto"
End Sub

Private Sub Favorites_Click()
    MostrarLibrosPorEstado "Favorito"
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
    End With
    rs.MoveNext
Loop

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Read_Click()
    MostrarLibrosPorEstado "Leido"
End Sub

Private Sub Recommended_Click()
    Dim rs As ADODB.Recordset
    Set rs = ObtenerRecomendados(1)

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

Private Sub Want_to_read_Click()
    MostrarLibrosPorEstado "PorLeer"
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
