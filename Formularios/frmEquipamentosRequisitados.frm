VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEquipamentosRequisitados 
   Caption         =   "frmEquipamentosRequisitados"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEntrega 
      Caption         =   "Confirma Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5760
      TabIndex        =   5
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox txtCodBaixa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtNumReq 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid tblProdutos 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      FormatString    =   "Produto                                  |Qtd. Pedida|Qtd. Atendida"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Codigo de Baixa"
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
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Número Requisição"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmEquipamentosRequisitados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEntrega_Click()
   Call Rotina_AbrirBanco
   Dim i As Integer
   
   i = 1
   
   db.BeginTrans
   
   Do While i < tblProdutos.Rows
      Prod.Open "SELECT grupo,classe,codProd FROM supProduto WHERE nomeProd = ('" & tblProdutos.TextMatrix(i, 0) & "')", db, 3, 3
      rs.Open "SELECT * FROM supRequisicaoDetalhe WHERE id = ('" & txtNumReq & "') and codigo = ('" & txtCodBaixa & "') and grupo = ('" & Prod!grupo & "') and classe = ('" & Prod!classe & "') and codProd = ('" & Prod!codProd & "')", db, 3, 3
      
         If rs!quantidade = rs!QtdEntregue Then
         
            rs!Status = 1
            rs!dataProcessamento = Date
            rs.Update
         
         End If
         
         
      rs.Close
      rs.Open "SELECT * FROM supEstoque WHERE grupo = ('" & Prod!grupo & "') and classe = ('" & Prod!classe & "') and codProd = ('" & Prod!codProd & "')", db, 3, 3
      Prod.Close
         
         rs!qtdEmEstoque = rs!qtdEmEstoque - tblProdutos.TextMatrix(i, 2)
         rs!qtdReservado = rs!qtdReservado - tblProdutos.TextMatrix(i, 2)
         rs!dataUltimaAtualizacao = Date
         rs!dataUltimaRequisicao = Date
         If rs!qtdReservado = 0 And rs!qtdEmEstoque <= rs!estoqueMinimo Then
         
            Call geraRequisicaoCompra(tblProdutos.TextMatrix(i, 0), tblProdutos.TextMatrix(i, 1), rs!qtdEmEstoque, tblProdutos.TextMatrix(i, 2), rs!estoqueMaximo)
         
         End If
         rs.Update
      rs.Close
      i = i + 1
   
   Loop
   
   db.CommitTrans
   
   MsgBox ("Entrega processada"), vbInformation
   Unload Me
   
   FechaDB
End Sub

Public Sub geraRequisicaoCompra(Produto As String, qtdRequisitada As Integer, qtdEstoque As Integer, qtdAtendida As Integer, estoqueMaximo As Integer)
   
   pes.Open "SELECT * FROM supRequisicaoCompra WHERE nomeProd = ('" & Produto & "') and idRequisicao = ('" & txtNumReq & "')", db, 3, 3
   
   If pes.EOF Then
   
      pes.AddNew
      
   End If
   
   pes!nomeProd = Produto
   pes!idRequisicao = txtNumReq
   pes!qtdRequisitada = qtdRequisitada
   pes!qtdEmEstoque = qtdEstoque
   pes!qtdPendente = qtdRequisitada - qtdAtendida
   pes!estoqueMaximo = estoqueMaximo
   pes!qtdComprar = estoqueMaximo + pes!qtdPendente
   pes!Status = 0
   pes.Update
   
   pes.Close

End Sub

