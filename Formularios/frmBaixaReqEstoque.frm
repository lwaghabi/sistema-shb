VERSION 5.00
Begin VB.Form frmBaixaReqEstoque 
   Caption         =   "frmBaixaReqEstoque"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txtCodReq 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   8535
   End
   Begin VB.Label Label2 
      Caption         =   "Código da Requisição"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Baixa de Requisição ao Estoque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   5895
   End
End
Attribute VB_Name = "frmBaixaReqEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub txtCodReq_LostFocus()
   
      
   If txtCodReq = Empty Then
      MsgBox ("Codigo de retirada não informado"), vbInformation
      cmdSair.SetFocus
      Exit Sub
   End If
   
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM suprequisicaodetalhe inner join supproduto on supproduto.grupo = suprequisicaodetalhe.grupo and supproduto.classe = suprequisicaodetalhe.classe and supproduto.codProd = suprequisicaodetalhe.codProd WHERE codigo = ('" & txtCodReq & "')", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Código inválido!"), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   rs.MoveFirst
   
   If rs!statusEntrega = 1 Then
      frmEquipamentosRequisitados.cmdEntrega.Enabled = False
   End If
   
   frmEquipamentosRequisitados.txtCodBaixa = rs!codigo
   frmEquipamentosRequisitados.txtNumReq = rs!Id
   
   frmEquipamentosRequisitados.tblProdutos.Rows = 1
   
   Do While Not rs.EOF
   
      frmEquipamentosRequisitados.tblProdutos.AddItem rs!nomeProd & vbTab & rs!quantidade & vbTab & rs!quantidadeAtendida
      rs.MoveNext
   
   Loop
    
   frmEquipamentosRequisitados.Show
      
   Unload Me
   
   FechaDB
End Sub
