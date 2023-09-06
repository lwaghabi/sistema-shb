VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReqCompra 
   Caption         =   "frmReqCompra     "
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17145
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   17145
   StartUpPosition =   2  'CenterScreen
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
      Height          =   735
      Left            =   15840
      TabIndex        =   23
      Top             =   7200
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid tblAcordo 
      Height          =   1095
      Left            =   11880
      TabIndex        =   22
      Top             =   1200
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1931
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "Fornecedor                  |Preço             "
   End
   Begin VB.CommandButton cmdGeraPO 
      Caption         =   "Gera P.O."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15840
      TabIndex        =   20
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdAtualizaLista 
      Caption         =   "Atualiza Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15720
      TabIndex        =   19
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtRequisitante 
      Height          =   375
      Left            =   8520
      TabIndex        =   18
      Top             =   1800
      Width           =   2280
   End
   Begin VB.TextBox txtRequisicao 
      Height          =   375
      Left            =   7440
      TabIndex        =   16
      Top             =   1800
      Width           =   1035
   End
   Begin VB.TextBox txtQtdComprar 
      Height          =   375
      Left            =   10800
      TabIndex        =   14
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtEstMax 
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtReqPend 
      Height          =   375
      Left            =   5600
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtEstoque 
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtProduto 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   4575
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
      Width           =   15375
      Begin MSFlexGridLib.MSFlexGrid tblProdutos 
         Height          =   5535
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   9763
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         FormatString    =   $"frmReqCompra.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Left            =   12960
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Acordo/Fornecedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11880
      TabIndex        =   21
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Requisitante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   17
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Requisição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   15
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Qtd Comprar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   13
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Est. Max"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Req.Pend"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Estoque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   3735
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
      Left            =   13080
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

Private Sub cmbFornecedor_Change()

End Sub

Private Sub cmdGeraPO_Click()
   Call Rotina_AbrirBanco
   Dim i As Integer
   Dim id As Integer
   
   db.BeginTrans
   
   rs.Open "SELECT * FROM supPedidoDeCompra WHERE id=('" & -1 & "')", db, 3, 3
   
   
   If rs.EOF Then
   
      rs.AddNew
   
   End If
   
   rs!DataPedido = Date
   rs!Status = 0
   rs!formaDePagamento = Empty
   rs!metodoPagamento = Empty
   rs!moeda = Empty
   rs.Update
   
   rs.Close
      
   Prod.Open "SELECT MAX(id) as idNovo FROM supPedidoDeCompra", db, 3, 3
   id = Prod!idNovo
   Prod.Close
      
   i = 1
   
   Do While i < tblProdutos.Rows
      If tblProdutos.TextMatrix(i, 7) <> Empty And tblProdutos.TextMatrix(i, 10) = "OK" Then
         Prod.Open "SELECT grupo,classe,codProd,AbreviaturaUnidadeEmbalagem as unid FROM supProduto INNER JOIN unidadeEmbalagem ON indice=unidadeProd WHERE nomeProd = ('" & tblProdutos.TextMatrix(i, 0) & "')", db, 3, 3
         rs.Open "SELECT * FROM supPedidoDetalhe WHERE id = ('" & id & "') AND grupo = ('" & Prod!grupo & "') AND classe = ('" & Prod!classe & "') and codProd = ('" & Prod!codProd & "')", db, 3, 3
         
         If rs.EOF Then
         
            rs.AddNew
         
         End If
         
         rs!id = id
         rs!grupo = Prod!grupo
         rs!classe = Prod!classe
         rs!codProd = Prod!codProd
         rs!qtdPedida = tblProdutos.TextMatrix(i, 7)
         rs!Status = 0
         rs!Unidade = Prod!Unid
         rs.Update
         Prod.Close
         rs.Close
         
         rs.Open "SELECT * FROM supRequisicaoCompra WHERE nomeProd = ('" & tblProdutos.TextMatrix(i, 0) & "') and idRequisicao = ('" & tblProdutos.TextMatrix(i, 5) & "')", db, 3, 3
         rs!Status = 1
         rs.Update
         rs.Close
         
      ElseIf tblProdutos.TextMatrix(i, 7) = Empty And tblProdutos.TextMatrix(i, 10) = "OK" Then
         
         rs.Open "SELECT * FROM supRequisicaoCompra WHERE nomeProd = ('" & tblProdutos.TextMatrix(i, 0) & "') and idRequisicao = ('" & tblProdutos.TextMatrix(i, 5) & "')", db, 3, 3
         rs!Status = 1
         rs.Update
         rs.Close
         
      End If
      i = i + 1
   Loop
   
   db.CommitTrans
   
   FechaDB
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
   Call Rotina_AbrirBanco
   Dim agregado As Integer
   Dim nomeAnterior As String
   
   rs.Open "SELECT * FROM supRequisicaoCompra WHERE status = 0 ORDER BY nomeProd", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Não há requisição de compra"), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   rs.MoveFirst
   tblProdutos.Rows = 1
   Do While Not rs.EOF
      
      Prod.Open "SELECT chPessoa FROM suprequisicao WHERE id = ('" & rs!idRequisicao & "')", db, 3, 3
      
      If rs!nomeProd = nomeAnterior Then
         agregado = agregado + rs!qtdPendente
         tblProdutos.AddItem rs!nomeProd & vbTab & rs!nomeProd & vbTab & rs!qtdEmEstoque & vbTab & rs!qtdPendente & vbTab & rs!estoqueMaximo & vbTab & rs!idRequisicao & vbTab & Prod!chPessoa & vbTab & ""
      
      Else
         nomeAnterior = rs!nomeProd
         If tblProdutos.Rows > 1 Then
            tblProdutos.TextMatrix(tblProdutos.Rows - 1, 7) = agregado
         End If
         agregado = rs!estoqueMaximo + rs!qtdPendente
         tblProdutos.AddItem rs!nomeProd & vbTab & rs!nomeProd & vbTab & rs!qtdEmEstoque & vbTab & rs!qtdPendente & vbTab & rs!estoqueMaximo & vbTab & rs!idRequisicao & vbTab & Prod!chPessoa & vbTab & ""
      
      End If
      
      Prod.Close
      
      rs.MoveNext
   
   Loop
   
   tblProdutos.TextMatrix(tblProdutos.Rows - 1, 7) = agregado
   
   rs.Close
   FechaDB
End Sub

Private Sub tblAcordo_Click()
   tblProdutos.TextMatrix(tblProdutos.Row, 8) = tblAcordo.TextMatrix(tblAcordo.Row, 0)
   tblProdutos.TextMatrix(tblProdutos.Row, 9) = tblAcordo.TextMatrix(tblAcordo.Row, 1)
End Sub

Private Sub tblProdutos_Click()
   Call Rotina_AbrirBanco
   If tblProdutos.Col = 1 Then
      
      rs.Open "SELECT * FROM supproduto WHERE nomeProd = ('" & tblProdutos.TextMatrix(tblProdutos.Row, 1) & "')", db, 3, 3
      If Not rs.EOF Then
         frmEspecTec.txtEspecificacaoTecnica = rs!especificacaoTecnica
         frmEspecTec.txtDescricao = rs!Descricao
      End If
      rs.Close
      frmEspecTec.Show
   
   Else
   
      txtProduto = tblProdutos.TextMatrix(tblProdutos.Row, 1)
      txtEstoque = tblProdutos.TextMatrix(tblProdutos.Row, 2)
      txtReqPend = tblProdutos.TextMatrix(tblProdutos.Row, 3)
      txtEstMax = tblProdutos.TextMatrix(tblProdutos.Row, 4)
      txtRequisicao = tblProdutos.TextMatrix(tblProdutos.Row, 5)
      txtRequisitante = tblProdutos.TextMatrix(tblProdutos.Row, 6)
      txtQtdComprar = tblProdutos.TextMatrix(tblProdutos.Row, 7)
      Prod.Open "SELECT grupo,classe,codProd FROM supProduto WHERE nomeProd = ('" & tblProdutos.TextMatrix(tblProdutos.Row, 0) & "')", db, 3, 3
      rs.Open "SELECT * FROM supAcordoComercial INNER JOIN supAcordoComercialDetalhe ON supAcordoComercialDetalhe.id = supAcordoComercial.id WHERE supAcordoComercial.grupo=('" & Prod!grupo & "') AND supAcordoComercial.classe=('" & Prod!classe & "') AND supAcordoComercialDetalhe.codProd=('" & Prod!codProd & "')", db, 3, 3
      
      tblAcordo.Rows = 1
      
         If rs.EOF Then
         
            tblAcordo.AddItem "S/Acordo"
            FechaDB
            Exit Sub
         
         End If
         
         tblAcordo.AddItem "S/Acordo"
         
         rs.MoveFirst
         
         Do While Not rs.EOF
         
            tblAcordo.AddItem rs!Fornecedor & vbTab & Format$(rs!precoUnit, "##,##0.00")
            rs.MoveNext
         
         Loop
         
      Prod.Close
      rs.Close
      
   
   End If
   FechaDB
End Sub

Private Sub tblProdutos_DblClick()

   If tblProdutos.TextMatrix(tblProdutos.Row, 10) = "OK" Then
      tblProdutos.TextMatrix(tblProdutos.Row, 10) = Empty
   Else
      tblProdutos.TextMatrix(tblProdutos.Row, 10) = "OK"
   End If

End Sub
