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
   Begin VB.CommandButton Command3 
      Caption         =   "Command1"
      Height          =   615
      Left            =   7080
      TabIndex        =   10
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   615
      Left            =   10200
      TabIndex        =   9
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3960
      TabIndex        =   8
      Top             =   7800
      Width           =   2535
   End
   Begin MSComctlLib.ListView Books_list 
      Height          =   6615
      Left            =   3840
      TabIndex        =   7
      Top             =   840
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

Private Sub Command4_Click()

End Sub
