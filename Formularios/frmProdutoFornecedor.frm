VERSION 5.00
Begin VB.Form frmProdutoFornecedor 
   Caption         =   "frmProdutoFornecedor"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbProduto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3240
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.ComboBox cmbFornecedor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Produtos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Fornecedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "frmProdutoFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT ", db, 3, 3
   
     
   
   rs.Close
   
   FechaDB
End Sub
