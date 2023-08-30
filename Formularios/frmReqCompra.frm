VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReqCompra 
   Caption         =   "frmReqCompra     "
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   14220
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRequisitante 
      Height          =   375
      Left            =   8040
      TabIndex        =   18
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtRequisicao 
      Height          =   375
      Left            =   6960
      TabIndex        =   16
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtQtdComprar 
      Height          =   375
      Left            =   9720
      TabIndex        =   14
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtEstMax 
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtReqPend 
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtEstoque 
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtProduto 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   13335
      Begin MSFlexGridLib.MSFlexGrid tblProdutos 
         Height          =   5535
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   9763
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         FormatString    =   $"frmReqCompra.frx":0000
      End
   End
   Begin VB.TextBox txtHoje 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Requisitante"
      Height          =   375
      Left            =   8040
      TabIndex        =   17
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Requisição"
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Qtd. Comprar"
      Height          =   375
      Left            =   9720
      TabIndex        =   13
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Est. Max"
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Req.Pend"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Estoque"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Produto"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Hoje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Requisição de Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmReqCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub tblProdutos_Click()
   If tblProdutos.Col = 1 Then
      Call Rotina_AbrirBanco
      rs.Open "SELECT especificacaoTecnica,descricao FROM supproduto WHERE nomeProd = ('" & tblProdutos.TextMatrix(tblProdutos.Row, 0) & "')", db, 3, 3
         frmEspecTec.txtEspecificacaoTecnica = rs!especificacaoTecnica
         frmEspecTec.txtDescricao = rs!Descricao
      rs.Close
      FechaDB
      frmEspecTec.Show
   
   Else
   
      If tblProdutos.TextMatrix(tblProdutos.Row, 9) = "OK" Then
         tblProdutos.TextMatrix(tblProdutos.Row, 9) = Empty
      Else
         tblProdutos.TextMatrix(tblProdutos.Row, 9) = "OK"
      End If
   
   End If
End Sub
