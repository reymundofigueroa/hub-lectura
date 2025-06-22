VERSION 5.00
Begin VB.Form frmAgregarLibro 
   Caption         =   "Agregar libro"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7650
   LinkTopic       =   "Form2"
   ScaleHeight     =   6780
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   10
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton btnGuardarLibro 
      Caption         =   "Guardar libro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   9
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Frame caption 
      Caption         =   "Agregar un libro"
      Height          =   4575
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   6495
      Begin VB.TextBox txtUrlMega 
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   3120
         Width           =   3375
      End
      Begin VB.ComboBox cmbGenero 
         Height          =   315
         Left            =   2400
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtAutor 
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtTitulo 
         Height          =   405
         Left            =   2400
         TabIndex        =   2
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Enlace_label 
         Caption         =   "Enlace libreria MEGA"
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
         Left            =   480
         TabIndex        =   7
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Genero_label 
         Caption         =   "Genero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Autor_label 
         Caption         =   "Autor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Titulo_label 
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   960
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAgregarLibro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub btnCancelar_Click()
    Unload Me

End Sub

Private Sub btnGuardarLibro_Click()
If Trim(txtTitulo.Text) = "" Or _
       cmbGenero.ListIndex = -1 Or _
       Trim(txtUrlMega.Text) = "" Then
        MsgBox "Por favor completa todos los campos obligatorios.", vbExclamation
        Exit Sub
    End If

    ' Obtener datos
    Dim titulo As String, autor As String, url As String
    Dim generoId As Integer

    titulo = Trim(txtTitulo.Text)
    autor = Trim(txtAutor.Text)
    url = Trim(txtUrlMega.Text)
    generoId = cmbGenero.ItemData(cmbGenero.ListIndex)

    ' Ejecutar inserción
    Dim sql As String
    sql = "INSERT INTO Libros (Titulo, Autor, GeneroId, UrlMega) VALUES (?, ?, ?, ?)"

    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText

    cmd.Parameters.Append cmd.CreateParameter("Titulo", adVarWChar, adParamInput, 200, titulo)
    cmd.Parameters.Append cmd.CreateParameter("Autor", adVarWChar, adParamInput, 100, autor)
    cmd.Parameters.Append cmd.CreateParameter("GeneroId", adInteger, adParamInput, , generoId)
    cmd.Parameters.Append cmd.CreateParameter("UrlMega", adLongVarWChar, adParamInput, -1, url)

    On Error GoTo ErrorHandler
    cmd.Execute

    MsgBox "Libro agregado correctamente.", vbInformation
    Unload Me
    Exit Sub

ErrorHandler:
    MsgBox "Error al guardar el libro: " & Err.Description, vbCritical
End Sub

Private Sub Form_Load()
   If conn Is Nothing Then Call ConectarBase
    If conn.State = adStateClosed Then conn.Open

    Dim rs As New ADODB.Recordset
    cmbGenero.Clear

    rs.Open "SELECT Id, Nombre FROM Generos", conn, adOpenStatic, adLockReadOnly
    Do Until rs.EOF
        cmbGenero.AddItem rs("Nombre")
        cmbGenero.ItemData(cmbGenero.NewIndex) = rs("Id")
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
