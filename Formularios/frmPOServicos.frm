VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPOServicos 
   Caption         =   " frmPOServicos"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18105
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   18105
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbTipoFaturamento 
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
      Left            =   15480
      TabIndex        =   34
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtFaturamento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   30
      Top             =   2040
      Width           =   5055
   End
   Begin MSComCtl2.DTPicker dtDataPedido 
      Height          =   495
      Left            =   7800
      TabIndex        =   24
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   243466241
      CurrentDate     =   45264
   End
   Begin VB.ComboBox cmbClasse 
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
      Left            =   4800
      TabIndex        =   10
      Top             =   3600
      Width           =   3615
   End
   Begin VB.ComboBox cmbLocalEntrega 
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
      Left            =   5160
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
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
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
   End
   Begin VB.ComboBox cmbIdPO 
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
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.Frame frServicos 
      Caption         =   "Serviços"
      Height          =   5535
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   17895
      Begin VB.TextBox txtQtdAdicionar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8520
         TabIndex        =   37
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdAtualizaConsumo 
         Caption         =   "Atualiza Consumo"
         Height          =   615
         Left            =   16080
         TabIndex        =   36
         Top             =   2400
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtDataServico 
         Height          =   495
         Left            =   11760
         TabIndex        =   33
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   243466241
         CurrentDate     =   45275
      End
      Begin VB.TextBox txtValorTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   13440
         TabIndex        =   28
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox txtQuantidade 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8280
         TabIndex        =   26
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtUnidade 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         TabIndex        =   25
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   615
         Left            =   16080
         TabIndex        =   22
         Top             =   4560
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelaPO 
         Caption         =   "Cancela PO"
         Height          =   615
         Left            =   16080
         TabIndex        =   21
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdEmitePO 
         Caption         =   "Emite PO"
         Height          =   615
         Left            =   16080
         TabIndex        =   20
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Salva PO"
         Height          =   615
         Left            =   16080
         TabIndex        =   19
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdAtualizarNaLista 
         Caption         =   "Atualiza Item na Lista"
         Height          =   615
         Left            =   16080
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtPreco 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10200
         TabIndex        =   16
         Top             =   1560
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid tblServicos 
         Height          =   2415
         Left            =   360
         TabIndex        =   14
         Top             =   2280
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   4260
         _Version        =   393216
         Rows            =   1
         Cols            =   10
         FixedCols       =   0
         FormatString    =   $"frmPOServicos.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cmbCodServ 
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
         Left            =   360
         TabIndex        =   9
         Top             =   1560
         Width           =   7935
      End
      Begin VB.ComboBox cmbGrupo 
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
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label16 
         Caption         =   "Qtd a adicionar"
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
         Left            =   8520
         TabIndex        =   38
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label14 
         Caption         =   "Validade"
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
         Left            =   11760
         TabIndex        =   32
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Valor Total"
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
         Left            =   11880
         TabIndex        =   29
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Qtd."
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
         Left            =   8280
         TabIndex        =   27
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Preço"
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
         Left            =   10200
         TabIndex        =   17
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Unidade"
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
         Left            =   8880
         TabIndex        =   15
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Servico"
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
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Classe"
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
         Left            =   4680
         TabIndex        =   12
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Grupo"
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
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Label Label15 
      Caption         =   "Tipo Faturamento"
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
      Left            =   15480
      TabIndex        =   35
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label13 
      Caption         =   "Faturamento"
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
      Left            =   10200
      TabIndex        =   31
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Label Label10 
      Caption         =   "Data do Pedido"
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
      Left            =   7800
      TabIndex        =   23
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Local de Entrega"
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
      Left            =   5160
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Fornecedor"
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
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "PO"
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
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Pedido de Compra de Serviços"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8055
   End
End
Attribute VB_Name = "frmPOServicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim Rel As Object
Dim Relatorio As String
Dim colunaGrupo As Integer
Dim colunaClasse As Integer
Dim colunaCodServ As Integer
Dim colunaUnidade As Integer
Dim colunaPreco As Integer
Dim colunaNome As Integer
Dim colunaQuantidade As Integer
Dim colunaPrecoTotal As Integer
Dim colunaDataServ As Integer
Dim colunaAcordo As Integer

Private Sub cmbClasse_LostFocus()
Call carregaServico
End Sub

Private Sub cmbGrupo_LostFocus()
Call carregaClasse
End Sub

Private Sub cmbIdPO_LostFocus()
Call carregaInfo
End Sub
Private Sub cmdAtualizaConsumo_Click()
Call atualizaConsumo
End Sub

Private Sub cmdAtualizarNaLista_Click()
Call atualizaLista
End Sub

Private Sub cmdCancelaPO_Click()
Call cancelaPO
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()
If txtFaturamento <> Empty Then
   Call salvaPO
Else
   MsgBox ("Faturamento não foi informado"), vbInformation
End If

End Sub

Private Sub cmdEmitePO_Click()

On Error GoTo Erro:

Call Rotina_AbrirBanco

db.BeginTrans

If cmbIdPO <> Empty Then
   rs.Open "SELECT * FROM servpo WHERE id = ('" & cmbIdPO & "')", db, 3, 3
   If Not rs.EOF Then
      If rs!Status = 0 Then
         db.Execute ("UPDATE servpo SET status = 1 WHERE id='" & cmbIdPO & "'")
      End If
   End If
   rs.Close
End If

db.CommitTrans
Relatorio = "drOrdemDeCompraServico"

Set Rel = drOrdemDeCompraServico
sql = "Select emp.empEmpresa, emp.empEndereco, emp.empCidade, emp.empBairro, emp.empUF, emp.empCEP, emp.empCNPJ, emp.empInscEst, emp.empEMAIL, pes.chPessoa, "
sql = sql & " pes.pesRazaoSocial, pes.pesEndereco, pes.pesBairro, pes.pesCidade, pes.chUF, pes.pesCEP, pes.chCNPJ_CPF, pes.pesInscEst_Ident, pes.pesTelContato, "
sql = sql & " po.id, po.fornecedor, po.dataPedido, po.localEntrega, po.faturamento, prd.descricao, userv.abrevunidserv, "
sql = sql & " det.grupo, det.classe, det.codServ, (det.quantidade - det.quantidadeAtendida) as saldo, det.valorServ, (det.valorServ * det.quantidade) as total, det.dataServico, "
sql = sql & " ender.rua, ender.numero, ender.complemento, ender.bairro, ender.cidade, ender.uf, ender.cep From empresa emp, supendereco ender, servpo po, servpodetalhe det, pessoa pes, servservico prd, unidadedeservicos userv "
sql = sql & " WHERE po.id = ('" & cmbIdPO & "') and det.id = po.id and (det.quantidade - det.quantidadeAtendida)>0 and ender.apelido = ('" & cmbLocalEntrega & "') and pes.chPessoa = ('" & cmbFornecedor & "') and det.grupo = prd.grupo and det.classe = prd.classe and det.codServ = prd.codServ and prd.unidade = userv.indice "

AbrirRelatorio sql, Rel

Call FechaDB
Exit Sub
Erro:  MsgBox ("Erro ao imprimir ordem de compra de Serviço: " & Err.Description), vbInformation
db.RollbackTrans
End Sub



Private Sub Form_Load()
Call carregaPO
Call carregaFornecedor
Call carregaGrupo
Call carregaData
Call carregaLocalEntrega
Call carregaTipoLancamento
colunaNome = 0
colunaQuantidade = 1
colunaUnidade = 2
colunaPreco = 3
colunaDataServ = 4
colunaPrecoTotal = 5
colunaGrupo = 6
colunaClasse = 7
colunaCodServ = 8
colunaAcordo = 9
End Sub

Public Sub carregaPO()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   rs.Open "SELECT * FROM servpo WHERE status = 0", db, 3, 3
   
   If rs.EOF Then
      MsgBox ("Não existem POs cadastradas"), vbInformation
      rs.Close
      Exit Sub
   End If
   
   Do While Not rs.EOF
   
      cmbIdPO.AddItem rs!Id
      rs.MoveNext
      
   Loop
   
   rs.Close
Exit Sub
Erro: MsgBox ("Erro ao carregar PO: " & Err.Description), vbInformation
rs.Close
End Sub

Public Sub carregaFornecedor()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   rs.Open "Select chPessoa from pessoa where pesTipoPessoa=1", db, 3, 3
   
   If rs.EOF Then
      MsgBox ("Não existem fonecedores cadastradas"), vbInformation
      rs.Close
      Exit Sub
   End If
   
   Do While Not rs.EOF
   
      cmbFornecedor.AddItem rs!chPessoa
      rs.MoveNext
      
   Loop
   
   rs.Close
Exit Sub
Erro: MsgBox ("Erro ao carregar fonecedores: " & Err.Description), vbInformation
rs.Close
End Sub

Public Sub carregaGrupo()
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   Prod.Open "Select * from servgrupoclasse where classe = '000'", db, 3, 3
      If Prod.EOF Then
         MsgBox ("ERRO: Arquivo vazio."), vbCritical
         Call FechaDB
         Exit Sub
      End If
      
      Prod.MoveFirst
      
      Do While Not Prod.EOF
         cmbGrupo.AddItem Prod!Descricao
         Prod.MoveNext
      Loop
   Prod.Close
   
   FechaDB
   
Exit Sub
Erro: MsgBox ("Erro ao carregar grupos" & Err.Description), vbInformation
FechaDB
End Sub

Public Sub carregaData()
   dtDataPedido = Date
   dtDataServico = Date
End Sub

Public Sub carregaLocalEntrega()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   rs.Open "Select apelido from supendereco", db, 3, 3

   If rs.EOF Then

      MsgBox ("Não existem endereços")
      FechaDB
      Exit Sub
   
   End If

   rs.MoveFirst

   Do While Not rs.EOF

      cmbLocalEntrega.AddItem rs!apelido
      rs.MoveNext

   Loop

   rs.Close

Exit Sub
Erro: MsgBox ("Erro ao carregar locais de entrega: " & Err.Description), vbInformation
rs.Close
End Sub

Public Sub carregaClasse()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   neg.Open "Select * from servgrupoclasse where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe != '000' ", db, 3, 3
   If neg.EOF Then
      MsgBox "Erro: Não existem classes nesse grupo", vbCritical
      FechaDB
      Exit Sub
   End If
   neg.MoveFirst
   
   cmbClasse.Clear
   
   Do While Not neg.EOF
      cmbClasse.AddItem neg!Descricao
      neg.MoveNext
   Loop
   
   neg.Close
   
Exit Sub
Erro: MsgBox ("Erro ao carregar classes" & Err.Description), vbInformation
FechaDB
End Sub

Public Sub carregaServico()
   cmbCodServ.Clear
   
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   Prod.Open "Select descricao from servservico where grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') order by codServ", db, 3, 3
   
   If Prod.EOF Then
   
      MsgBox ("Não há serviços cadastrados nessa categoria"), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   Prod.MoveFirst
   
   Do While Not Prod.EOF
   
      cmbCodServ.AddItem Prod!Descricao
      Prod.MoveNext
   
   Loop
   
   FechaDB
   
Exit Sub
Erro: MsgBox ("Erro ao carregar serviços: " & Err.Description), vbInformation
FechaDB
End Sub

Public Function verificaLista(Grupo As String, Classe As String, codServ As String) As Integer
   Dim result As Integer
   Dim i As Integer
   i = 1
   result = 0
   Do While i < tblServicos.Rows
      If tblServicos.TextMatrix(i, colunaGrupo) = Grupo And tblServicos.TextMatrix(i, colunaClasse) = Classe And tblServicos.TextMatrix(i, colunaCodServ) = codServ Then
         result = i
      End If
      i = i + 1
   Loop
   verificaLista = result
End Function

Public Sub atualizaLista()
   Dim Index As Integer
   Index = verificaLista(Format$(cmbGrupo.ListIndex + 1, "00"), Format$(cmbClasse.ListIndex + 1, "000"), Format$(cmbCodServ.ListIndex + 1, "00000"))
   If Index > 0 Then
      Call salvaMudanca(Index)
      txtValorTotal = Format(calculaTotal, "##,##0.00")
   Else
      MsgBox ("Não podem ser incluidos mais serviços na PO!"), vbInformation
   End If
End Sub

Public Sub salvaMudanca(Index As Integer)
   tblServicos.TextMatrix(Index, colunaPreco) = Format$(txtPreco, "##,##0.00")
   tblServicos.TextMatrix(Index, colunaQuantidade) = txtQuantidade
   tblServicos.TextMatrix(Index, colunaPrecoTotal) = Format(txtPreco * txtQuantidade, "##,##0.00")
   tblServicos.TextMatrix(Index, colunaDataServ) = dtDataServico
   
End Sub

Public Sub salvaPO()
   On Error GoTo Erro
   Dim i As Integer
   
   Call Rotina_AbrirBanco
   
   
   rs.Open "SELECT * FROM servpo WHERE id = '" & cmbIdPO & "'", db, 3, 3
   
   If rs!Status > 0 Then
      MsgBox ("Não foi possível fazer alterações na PO, pois a PO já foi emitida ou concluida ou cancelada!"), vbInformation
      FechaDB
      Exit Sub
   End If
   
   db.BeginTrans
   
   rs!Id = cmbIdPO
   rs!fornecedor = cmbFornecedor
   rs!DataPedido = dtDataPedido
   rs!localEntrega = cmbLocalEntrega
   rs!faturamento = txtFaturamento
   rs!tipoFaturamento = cmbTipoFaturamento.ListIndex
   
   
   rs.Update
   rs.Close
   
   i = 1
   
   Do While i < tblServicos.Rows
      db.Execute ("UPDATE servpodetalhe SET valorServ= '" & Replace(tblServicos.TextMatrix(i, colunaPreco), ",", ".") & "',quantidade='" & tblServicos.TextMatrix(i, colunaQuantidade) & "',dataServico='" & Format(tblServicos.TextMatrix(i, colunaDataServ), "yyyy-MM-dd") & "' WHERE id = '" & cmbIdPO & "' AND grupo = '" & tblServicos.TextMatrix(i, colunaGrupo) & "' AND classe = '" & tblServicos.TextMatrix(i, colunaClasse) & "' AND codServ = '" & tblServicos.TextMatrix(i, colunaCodServ) & "'")
      If Not tblServicos.TextMatrix(i, colunaAcordo) = Empty And tblServicos.TextMatrix(i, colunaAcordo) > 0 Then
         Call subtraiDoAcordo(tblServicos.TextMatrix(i, colunaGrupo), tblServicos.TextMatrix(i, colunaClasse), tblServicos.TextMatrix(i, colunaCodServ), CInt(tblServicos.TextMatrix(i, colunaAcordo)), CInt(tblServicos.TextMatrix(i, colunaQuantidade)))
      End If
      i = i + 1
   Loop
   
   db.CommitTrans
   FechaDB
   MsgBox ("PO salva com sucesso!"), vbInformation
   cmdSalvar.Enabled = False
   cmdEmitePO.Enabled = True
Exit Sub
Erro: MsgBox ("Erro ao salvar PO: " & Err.Description), vbInformation
db.RollbackTrans
FechaDB
End Sub

Public Function calculaTotal() As Currency
   On Error GoTo Erro
   Dim i As Integer
   Dim total As Currency
   i = 1
   total = 0
   Do While i < tblServicos.Rows
      total = total + CCur(tblServicos.TextMatrix(i, colunaPrecoTotal))
      i = i + 1
   Loop
   calculaTotal = total
Exit Function
Erro: MsgBox ("Erro ao calcular total: " & Err.Description)
End Function

Public Sub cancelaPO()
   On Error GoTo Erro
   Dim i As Integer
   
   Call Rotina_AbrirBanco
   
   Prod.Open "SELECT * FROM servpo WHERE id = '" & cmbIdPO & "'", db, 3, 3
   If Prod!Status < 2 Then
   
      db.BeginTrans
      db.Execute ("UPDATE servpo SET status = 3 WHERE id='" & cmbIdPO & "' ")
      i = 1
      Do While i < tblServicos.Rows
         rs.Open "SELECT * FROM servpodetalhe WHERE id = '" & cmbIdPO & "' AND grupo = ('" & tblServicos.TextMatrix(i, colunaGrupo) & "') AND classe = ('" & tblServicos.TextMatrix(i, colunaClasse) & "') AND codServ = ('" & tblServicos.TextMatrix(i, colunaCodServ) & "')", db, 3, 3
         If rs!Status <> 2 Then
            Call retornaSaldoAcordo(i, rs!quantidadeAtendida)
         End If
         i = i + 1
         rs.Close
      Loop
      
      db.CommitTrans
      MsgBox ("Cancelado com sucesso!"), vbInformation
   Else
      MsgBox ("PO não pode ser cancelada pois já foi cancelada previamente ou atendida!"), vbInformation
   End If
   Prod.Close
   FechaDB
   cmdCancelaPO.Enabled = False
Exit Sub
Erro: MsgBox ("Erro ao cancelar a PO: " & Err.Description), vbInformation
db.RollbackTrans
FechaDB
End Sub

Public Sub carregaInfo()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   Dim valor As Currency
   
   If cmbIdPO = Empty Then
      MsgBox ("Insira valor válida para o campo de PO!"), vbInformation
      Exit Sub
   End If
   
   rs.Open "SELECT sp.fornecedor,sp.status,spd.status as statusServ,spd.quantidadeAtendida,spd.acordo,sp.tipoFaturamento,sp.dataPedido,sp.faturamento,spd.dataServico,sp.localEntrega,spd.grupo,spd.classe,spd.codServ,ss.descricao,us.abrevUnidServ,spd.valorServ,spd.quantidade FROM servpo sp INNER JOIN servpodetalhe spd ON  sp.id=spd.id INNER JOIN servservico ss ON ss.grupo=spd.grupo AND ss.classe = spd.classe AND ss.codServ=spd.codServ INNER JOIN unidadedeservicos us ON us.indice=ss.unidade WHERE sp.id = ('" & cmbIdPO & "')", db, 3, 3
   If Not IsNull(rs!fornecedor) Then
      cmbFornecedor = rs!fornecedor
   End If
   If Not IsNull(rs!DataPedido) Then
      dtDataPedido = rs!DataPedido
   End If
   If Not IsNull(rs!localEntrega) Then
      cmbLocalEntrega = rs!localEntrega
   End If
   
   txtFaturamento = rs!faturamento
   
   cmbTipoFaturamento.ListIndex = rs!tipoFaturamento
   
   If rs!Status = 1 Or rs!Status = 2 Then
      txtQtdAdicionar.Visible = True
      Label16.Visible = True
   End If
   
   If rs!Status > 0 Then
      cmdSalvar.Enabled = False
      cmdEmitePO.Enabled = True
   Else
      cmdSalvar.Enabled = True
      cmdEmitePO.Enabled = False
   End If
   
   tblServicos.Rows = 1
   
   Do While Not rs.EOF
      If Not IsNull(rs!valorServ) Then
         valor = rs!valorServ
      End If
      tblServicos.AddItem rs!Descricao & vbTab & rs!quantidade & vbTab & rs!abrevunidserv & vbTab & Format(valor, "##,##0.00") & vbTab & rs!dataServico & vbTab & Format(rs!quantidade * valor, "##,##0.00") & vbTab & rs!Grupo & vbTab & rs!Classe & vbTab & rs!codServ & vbTab & rs!acordo
      If rs!dataServico < Date And rs!statusServ <> 2 Then
         Call retornaSaldoAcordo(tblServicos.Rows - 1, rs!quantidadeAtendida)
         Call atualizaQuantidade(tblServicos.Rows - 1)
         MsgBox ("Data expirada no item: " & rs!Descricao), vbInformation
      End If
      rs.MoveNext
   Loop
   
   txtValorTotal = Format(calculaTotal, "##,##0.00")
   
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao carregar informações da PO: " & Err.Description), vbInformation
rs.Close
End Sub

Private Sub tblServicos_Click()
   cmbGrupo.ListIndex = CInt(tblServicos.TextMatrix(tblServicos.Row, colunaGrupo)) - 1
   Call carregaClasse
   cmbClasse.ListIndex = CInt(tblServicos.TextMatrix(tblServicos.Row, colunaClasse)) - 1
   Call carregaServico
   cmbCodServ.ListIndex = CInt(tblServicos.TextMatrix(tblServicos.Row, colunaCodServ)) - 1
   txtUnidade = tblServicos.TextMatrix(tblServicos.Row, colunaUnidade)
   txtPreco = tblServicos.TextMatrix(tblServicos.Row, colunaPreco)
   txtQuantidade = tblServicos.TextMatrix(tblServicos.Row, colunaQuantidade)
   cmdAtualizaConsumo.Enabled = True
End Sub

Public Sub subtraiDoAcordo(Grupo As String, Classe As String, codServ As String, acordo As Integer, qtdAtendida As Integer)
   On Error GoTo Erro
   db.Execute ("UPDATE servacordocomercialdetalhe SET qtdEntregue = qtdEntregue + " & qtdAtendida & " WHERE id = '" & acordo & "' AND codServ = ('" & codServ & "') ")
Exit Sub
Erro: MsgBox ("Erro ao atualizar serviços restantes em contrato: " & Err.Description), vbInformation
End Sub

Public Sub carregaTipoLancamento()
   
   On Error GoTo Erro
   
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT * FROM tipolancamento", db, 3, 3
   
   If rs.EOF Then
      MsgBox ("Não foi possível localizar os tipo de lançamento"), vbCritical
      FechaDB
      Exit Sub
   End If
   
   Do While Not rs.EOF
      cmbTipoFaturamento.AddItem rs!chTipoDocumento
      rs.MoveNext
   Loop
   
   rs.Close
   
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao carregar os tipos de lançamento: " & Err.Description), vbInformation
rs.Close
End Sub

Public Sub retornaSaldoAcordo(i As Integer, qtdAtendida As Integer)
   On Error GoTo Erro
   db.Execute ("UPDATE servacordocomercialdetalhe SET qtdEntregue = qtdEntregue - " & tblServicos.TextMatrix(i, colunaQuantidade) - qtdAtendida & " WHERE id = '" & tblServicos.TextMatrix(i, colunaAcordo) & "' AND codServ = ('" & tblServicos.TextMatrix(i, colunaCodServ) & "') ")
Exit Sub
Erro: MsgBox ("Erro ao atualizar serviços restantes em contrato: " & Err.Description), vbInformation
End Sub

Private Sub txtFaturamento_LostFocus()
Verifica = Empty
Verifica = Mid$(txtFaturamento, 31, 5)
If Not Verifica = Empty Then
   MsgBox ("Faturamento informado ultrapassa 30 caracteres.")
   cmdSair.SetFocus
   Exit Sub
End If
End Sub

Private Sub txtQuantidade_LostFocus()
   On Error GoTo Erro
   Call Rotina_AbrirBanco
   
   If tblServicos.TextMatrix(tblServicos.Row, colunaAcordo) <> Empty And CInt(tblServicos.TextMatrix(tblServicos.Row, colunaAcordo)) > 0 Then
      rs.Open "SELECT (qtdTotal-qtdEntregue) as saldo FROM servacordocomercialdetalhe sacd INNER JOIN servacordocomercial sac ON sacd.id=sac.id WHERE sacd.id=('" & tblServicos.TextMatrix(tblServicos.Row, colunaAcordo) & "') and sac.grupo = ('" & tblServicos.TextMatrix(tblServicos.Row, colunaGrupo) & "') and sac.classe = ('" & tblServicos.TextMatrix(tblServicos.Row, colunaClasse) & "') and sacd.codServ = ('" & tblServicos.TextMatrix(tblServicos.Row, colunaCodServ) & "')", db, 3, 3
      
      If Not rs.EOF Then
         If rs!saldo < CInt(txtQuantidade) Then
            MsgBox ("Saldo contido no acordo insuficiente! Saldo acordo: " & rs!saldo), vbInformation
         End If
      rs.Close
      End If
   End If
   FechaDB
Exit Sub
Erro: MsgBox ("Erro ao verificar saldo restante do acordo" & Err.Description), vbInformation
FechaDB
End Sub

Public Sub atualizaConsumo()
   On Error GoTo Erro
   Dim i As Integer
   i = 1
      Call Rotina_AbrirBanco
      db.BeginTrans
      rs.Open "SELECT * FROM servacordocomercialdetalhe sacd inner join servacordocomercial sac on sac.id=sacd.id WHERE sacd.id = '" & tblServicos.TextMatrix(tblServicos.Row, colunaAcordo) & "' AND grupo = '" & tblServicos.TextMatrix(tblServicos.Row, colunaGrupo) & "' AND classe = '" & tblServicos.TextMatrix(tblServicos.Row, colunaClasse) & "' AND codServ = '" & tblServicos.TextMatrix(tblServicos.Row, colunaCodServ) & "'", db, 3, 3
      If tblServicos.TextMatrix(tblServicos.Row, colunaAcordo) <> Empty And CInt(tblServicos.TextMatrix(tblServicos.Row, colunaAcordo)) > 0 And txtQtdAdicionar <> Empty And (rs!qtdTotal - rs!QtdEntregue) >= CInt(txtQtdAdicionar) Then
         db.Execute ("UPDATE servpodetalhe SET quantidade = quantidade + " & txtQtdAdicionar & ",dataServico='" & Format(dtDataServico, "yyyy-MM-dd") & "',status=1 WHERE id = '" & cmbIdPO & "' AND grupo = '" & tblServicos.TextMatrix(tblServicos.Row, colunaGrupo) & "' AND classe='" & tblServicos.TextMatrix(tblServicos.Row, colunaClasse) & "'  AND codServ ='" & tblServicos.TextMatrix(tblServicos.Row, colunaCodServ) & "'")
         Call subtraiDoAcordo(tblServicos.TextMatrix(tblServicos.Row, colunaGrupo), tblServicos.TextMatrix(tblServicos.Row, colunaClasse), tblServicos.TextMatrix(tblServicos.Row, colunaCodServ), CInt(tblServicos.TextMatrix(tblServicos.Row, colunaAcordo)), CInt(txtQtdAdicionar))
      Else
         MsgBox ("Valor ao adicionar inválido!"), vbInformation
      End If
      rs.Close
      db.CommitTrans
      FechaDB
      Call carregaInfo
      txtValorTotal = Format(calculaTotal, "##,##0.00")
      cmdAtualizaConsumo.Enabled = False
Exit Sub
Erro: MsgBox ("Erro ao atualizar consumo: " & Err.Description), vbInformation
db.RollbackTrans
End Sub

Public Sub atualizaQuantidade(i As Integer)
On Error GoTo Erro
   db.Execute ("UPDATE servpodetalhe SET quantidade = quantidadeAtendida,status=2 WHERE id = '" & cmbIdPO & "' AND grupo = ('" & tblServicos.TextMatrix(i, colunaGrupo) & "') AND classe = ('" & tblServicos.TextMatrix(i, colunaClasse) & "') AND codServ = ('" & tblServicos.TextMatrix(i, colunaCodServ) & "') ")
Exit Sub
Erro: MsgBox ("Erro ao atualizar serviços restantes em contrato: " & Err.Description), vbInformation
End Sub
