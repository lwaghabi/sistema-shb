VERSION 5.00
Begin VB.Form frmCalculadora 
   Caption         =   "SHB - Calculadora"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14340
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   14340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command25 
      Caption         =   "VERIFICA DETPROD"
      Height          =   975
      Left            =   11400
      TabIndex        =   51
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Ajustar Centro de Custo de Int Para Varchar"
      Height          =   975
      Left            =   11040
      TabIndex        =   50
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Descricao Centro de Custo em Produto Fornecedor"
      Height          =   1215
      Left            =   11040
      TabIndex        =   49
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Lançar Data Pagamento em Desdobramento"
      Height          =   1215
      Left            =   10920
      TabIndex        =   48
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Gerar UnidOperFunc"
      Height          =   855
      Left            =   8880
      TabIndex        =   47
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Criar Data Pagamento em Historico DetProd"
      Height          =   975
      Left            =   6720
      TabIndex        =   46
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Cria Centro de Custo em NFDetProd"
      Height          =   975
      Left            =   5280
      TabIndex        =   45
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Espandir Centro de Custo em Produto Entrada"
      Height          =   975
      Left            =   3720
      TabIndex        =   44
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Expandir Centro de Custo em ProdFornec"
      Height          =   975
      Left            =   1800
      TabIndex        =   43
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Gerar Classificacao na Tab Produto"
      Height          =   975
      Left            =   120
      TabIndex        =   42
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Gerar Faturamento"
      Height          =   1215
      Left            =   6720
      TabIndex        =   41
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Zera Custo Contabil"
      Height          =   1215
      Left            =   5280
      TabIndex        =   40
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Gera Custo Anual"
      Height          =   1215
      Left            =   8880
      TabIndex        =   39
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Ajusta Fornecedores e Despesas"
      Height          =   1215
      Left            =   3600
      TabIndex        =   38
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Criar Centro de Custo Contabilidade"
      Height          =   1215
      Left            =   1800
      TabIndex        =   37
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Corrigir Negociação e Detalhe"
      Height          =   1215
      Left            =   120
      TabIndex        =   36
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Corrige Not Det Prod"
      Height          =   1095
      Left            =   480
      TabIndex        =   35
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   255
      Left            =   8520
      TabIndex        =   34
      Top             =   4560
      Width           =   90
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   195
      Left            =   8280
      TabIndex        =   33
      Top             =   4560
      Width           =   135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Corrige Unidade Operacional no Historico"
      Height          =   495
      Left            =   5520
      TabIndex        =   32
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton cmdAjustaHist 
      Caption         =   "Ajusta Hist"
      Height          =   855
      Left            =   3120
      TabIndex        =   31
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmddrMedicao 
      Caption         =   "Medição"
      Height          =   360
      Left            =   4200
      TabIndex        =   30
      Top             =   360
      Width           =   990
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Limpa tabe produtos de entrada"
      Height          =   615
      Left            =   8400
      TabIndex        =   29
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Limpa Promotora"
      Height          =   495
      Left            =   8400
      TabIndex        =   28
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Limpa Representante"
      Height          =   615
      Left            =   8400
      TabIndex        =   27
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmd 
      Caption         =   "LIMPA PROD TERCEIROS"
      Height          =   735
      Left            =   6240
      TabIndex        =   26
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LIMPA NEGOCIAÇÃO E DETALHE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   25
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LIMPA PESSOA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   24
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdLimpaPreco 
      Caption         =   "Limpa Preco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   23
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton cmdLimpaHistorico 
      Caption         =   "Limpa Historico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      TabIndex        =   22
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdLimpaCtaPagar 
      Caption         =   "Limpa Contas a Pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   21
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdLimpaCtaReceber 
      Caption         =   "Limpa Contas a Receber"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   20
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Digits 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   555
      TabIndex        =   18
      Top             =   2820
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   4
      Left            =   0
      TabIndex        =   17
      Top             =   1725
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   5
      Left            =   555
      TabIndex        =   16
      Top             =   1725
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   6
      Left            =   1125
      TabIndex        =   15
      Top             =   1725
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   7
      Left            =   0
      TabIndex        =   14
      Top             =   2280
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   8
      Left            =   555
      TabIndex        =   13
      Top             =   2280
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   9
      Left            =   1125
      TabIndex        =   12
      Top             =   2280
      Width           =   465
   End
   Begin VB.CommandButton DotBttn 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1320
      TabIndex        =   11
      Top             =   2820
      Width           =   465
   End
   Begin VB.CommandButton ClearBttn 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      TabIndex        =   10
      Top             =   2820
      Width           =   465
   End
   Begin VB.CommandButton Plus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1785
      TabIndex        =   9
      Top             =   1725
      Width           =   465
   End
   Begin VB.CommandButton Minus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2385
      TabIndex        =   8
      Top             =   1725
      Width           =   465
   End
   Begin VB.CommandButton Times 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1785
      TabIndex        =   7
      Top             =   2280
      Width           =   465
   End
   Begin VB.CommandButton Div 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2385
      TabIndex        =   6
      Top             =   2280
      Width           =   465
   End
   Begin VB.CommandButton Equals 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1785
      TabIndex        =   5
      Top             =   2820
      Width           =   1095
   End
   Begin VB.CommandButton Digits 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   1125
      TabIndex        =   4
      Top             =   1185
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   555
      TabIndex        =   3
      Top             =   1185
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   1185
      Width           =   465
   End
   Begin VB.CommandButton PlusMinus 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1785
      TabIndex        =   1
      Top             =   1185
      Width           =   465
   End
   Begin VB.CommandButton Over 
      Caption         =   "1/X"
      Height          =   450
      Left            =   2385
      TabIndex        =   0
      Top             =   1185
      Width           =   465
   End
   Begin VB.Label Display 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   0
      TabIndex        =   19
      Top             =   600
      Width           =   2820
   End
End
Attribute VB_Name = "frmCalculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ******************************
'  ******************************
'  ** MASTERING VB6            **
'  ** by Evangelos Petroutos   **
'  ** SYBEX, 1998              **
'  ******************************
'  ******************************
Option Explicit
Dim Operand1 As Double, Operand2 As Double
Dim Operator As String
Dim ClearDisplay As Boolean
Dim ValorTotal As Currency
Dim ProdutoAnterior As String
Dim Nome As String

Private Sub ClearBttn_Click()
    Display.Caption = ""
End Sub

Private Sub cmdAjustaHist_Click()
Dim Pessoa As String
Dim NotaFiscal As String
Dim Fatura As String


End Sub

Private Sub cmddrMedicao_Click()
Dim sql As String
Dim Rel As Object
Dim txtNome As String
Dim txtNumPedido As String
Dim txtPedidoComp As String
txtNome = "HEFTOS"
txtPedidoComp = "1"
txtNumPedido = "HFT4.1"
Set Rel = drMedicao
sql = "Select neg.chPessoa, neg.chUnidadeOperacional, neg.chNumPedido, neg.chNumPedidoComp, pes.pesRazaoSocial, "
sql = sql & " det.chDataInicio, det.chDataFim, det.chProduto, det.pedValorDaOperacao, det.pedUnidade, det.pedQuantidadePedida, "
sql = sql & " det.pedPrecoUnidadePedida, det.pedValorDaDiaria, det.pedQtdDias, det.pedValorDaOperacao, prd.prdDescCompleta "
sql = sql & " from negociacao neg, detalhenegociacao det, pessoa pes, produto prd "
sql = sql & " WHERE neg.chNumPedido = ('" & txtNumPedido & "') and neg.chNumPedidoComp = ('" & txtPedidoComp & "') "
sql = sql & " and det.chNumpedido = neg.chNumpedido and det.chNumpedidoComp = neg.chNumPedidoComp "
sql = sql & " and neg.chPessoa = pes.chPessoa and det.chProduto = prd.chProduto"

AbrirRelatorio sql, Rel

End Sub

Private Sub Command10_Click()
Dim ContaReg As Integer
Call Rotina_AbrirBanco
ContaReg = 0

hneg.Open "Select * from negociacao", db, 3, 3
If hneg.EOF Then
   MsgBox ("negociacao não encontrado"), vbCritical
   Call FechaDB
   Exit Sub
End If

neg.Open "Select * from negociacao where chNumPedido = ('" & hneg!chNumPedido & "') and chNumPedidoComp = ('" & hneg!chNumPedidoComp & "')", db, 3, 3
If neg.EOF Then
   neg.AddNew
   
   neg!chNumPedido = hneg!chNumPedido
   neg!chNumPedidoComp = hneg!chNumPedidoComp
   neg!negContrato = hneg!negContrato
   neg!negContratoComp = hneg!negContratoComp
   neg!negNumFatura = hneg!negNumFatura
   neg!negSerieFatura = hneg!negSerieFatura
   neg!negDataEmissaoFatura = hneg!negDataEmissaoFatura
   neg!negTipoProduto = hneg!negTipoProduto
   neg!negNotaFiscal = hneg!negNotaFiscal
   neg!negEmissorNF = hneg!negEmissorNF
   neg!chCodBcoLart = hneg!chCodBcoLart
   neg!chOrdemDeCarga = hneg!chOrdemDeCarga
   neg!negTransporte = hneg!negTransporte
   neg!negPlaca = hneg!negPlaca
   neg!negStatus = hneg!negStatus
   neg!negDataEnvioAprovMedicao = hneg!negDataEnvioAprovMedicao
   neg!negDataPedido = hneg!negDataPedido
   neg!negdatanegociação = hneg!negdatanegociação
   neg!negInicioMedicao = hneg!negInicioMedicao
   neg!negFinalMedicao = hneg!negFinalMedicao
   neg!negvalornegociacao = hneg!negvalornegociacao
   neg!chPessoa = hneg!chPessoa
   neg!chUnidadeOperacional = hneg!chUnidadeOperacional
   neg!chrepresentante = hneg!chrepresentante
   neg!chPromotor = hneg!chPromotor
   neg!negIntervaloFatura = hneg!negIntervaloFatura
   neg!negAPartirDe = hneg!negAPartirDe
   neg!negFreteColeta = hneg!negFreteColeta
   neg!negCobrancaFrete = hneg!negCobrancaFrete
   neg!negboletafrete = hneg!negboletafrete
   neg!negValorFixoFrete = hneg!negValorFixoFrete
   neg!negCondProcess = hneg!negCondProcess
   neg!negdesccomissao = hneg!negdesccomissao
   neg!negDescComisPromot = hneg!negDescComisPromot
   neg!negPrazoAdicional = hneg!negPrazoAdicional
   neg!negLançamento = hneg!negLançamento
   neg!negUltimaAtualizacao = hneg!negUltimaAtualizacao
   neg!negCntrlFaturamento = hneg!negCntrlFaturamento
   neg!negICMS = hneg!negICMS
   neg!negAliquota = hneg!negAliquota
   neg!negFretePedido = hneg!negFretePedido
   neg!negValorDoProduto = hneg!negValorDoProduto
   neg!negIPI = hneg!negIPI
   neg!negDescontoTotalPedido = hneg!negDescontoTotalPedido
   neg!negComisRepPedido = hneg!negComisRepPedido
   neg!negComisPromotPedido = hneg!negComisPromotPedido
   neg!negMotivacao = hneg!negMotivacao
   neg!negDataLancamento = hneg!negDataLancamento
   neg!negCEFOP = hneg!negCEFOP

   neg.Update
   
End If

dneg.Open "Select * from detalhenegociacao", db, 3, 3

If dneg.EOF Then
   MsgBox ("Detalhe não encontrado"), vbCritical
   Call FechaDB
   Exit Sub
End If

dneg.MoveFirst

Do While Not dneg.EOF
   If hneg.State = 1 Then
      hneg.Close: Set hneg = Nothing
   End If
   hneg.Open "Select *from detalhenegociacao where chNumPedido = ('" & dneg!chNumPedido & "') and chNumPedidoComp = ('" & dneg!chNumPedidoComp & "') and chDataInicio = ('" & dneg!chDataInicio & "')and chDataFim = ('" & dneg!chDataFim & "')", db, 3, 3
   If hneg.EOF Then
      hneg.AddNew
      hneg!chNumPedido = dneg!chNumPedido
      hneg!chNumPedidoComp = dneg!chNumPedidoComp
      hneg!chProduto = dneg!chProduto
      hneg!chDataInicio = dneg!chDataInicio
      hneg!chDataFim = dneg!chDataFim
      hneg!pedAtividade = dneg!pedAtividade
      hneg!pedunidade = dneg!pedunidade
      hneg!pedquantidadePedida = dneg!pedquantidadePedida
      hneg!pedPUCheio = dneg!pedPUCheio
      hneg!pedPrecoUnidadePedida = dneg!pedPrecoUnidadePedida
      hneg!pedValorDaDiaria = dneg!pedValorDaDiaria
      hneg!pedqtddias = dneg!pedqtddias
      hneg!pedValorDaOperacao = dneg!pedValorDaOperacao
      hneg!pedDesconto = dneg!pedDesconto
      hneg!pedValorDesconto = dneg!pedValorDesconto
      hneg!pedcomissaorep = dneg!pedcomissaorep
      hneg!pedcomissaopromot = dneg!pedcomissaopromot
      hneg!pedStatus = dneg!pedStatus
      
      hneg.Update
      
      ContaReg = ContaReg + 1
      
            
   End If
     
   dneg.MoveNext
   
Loop

MsgBox ("Total registroos = "), ContaReg
   
Call FechaDB

End Sub

Private Sub Command11_Click()

Call Rotina_AbrirBanco

hnfd.Open "Select * from notafiscaldetprod", db, 3, 3
If hnfd.EOF Then
   MsgBox ("Det Prod Vazio")
   Exit Sub
End If

hnfd.MoveFirst

Do While Not hnfd.EOF
  
   If pes.State = 1 Then
      pes.Close: Set pes = Nothing
   End If
      
   pes.Open "Select * from pessoa where chPessoa =  ('" & hnfd!chPessoa & "')", db, 3, 3
   If pes.EOF Then
      If (hnfd!chPessoa = "IMPOSTOS") Or (hnfd!chPessoa = "DESPESAS") Then
         Call GeraCentroCustoTratarFornecedor
      Else
         Call GeraCentroCustoTratarDespesa
      End If
   Else
      Call GeraCentroCustoTratarFornecedor
   End If

hnfd.MoveNext
   
Loop

MsgBox ("Fim de Serviço")

End Sub


Public Sub GeraCentroCustoTratarDespesa()

   If Not hnfd!chPessoa = ProdutoAnterior Then
      
      If ccc.State = 1 Then
         ccc.Close: Set ccc = Nothing
      End If
      
      ccc.Open "Select * from centrodecustoconatbil where chProduto = ('" & hnfd!chPessoa & "')", db, 3, 3
      If ccc.EOF Then
         ccc.AddNew
         ccc!cccClassificacaoCusto = "9.9"
         ccc!chProduto = hnfd!chPessoa
         ProdutoAnterior = hnfd!chPessoa
         ccc.Update
      End If
   End If
End Sub


Public Sub GeraCentroCustoTratarFornecedor()

If ccc.State = 1 Then
   ccc.Close: Set ccc = Nothing
End If

ccc.Open "Select * from centrodecustoconatbil where chProduto = ('" & hnfd!chCodProduto & "')", db, 3, 3
If ccc.EOF Then
   ccc.AddNew
   ccc!cccClassificacaoCusto = "9.9"
   ccc!chProduto = hnfd!chCodProduto
   'ProdutoAnterior = hnfd!chPessoa
   ccc.Update
End If

End Sub

Private Sub Command12_Click()
Dim fornecedorAnterior As String


Call Rotina_AbrirBanco

ProdEntrada.Open "Select * from produtoentrada", db, 3, 3
If ProdEntrada.EOF Then
   MsgBox ("Produto Entrada, fornecedores DoEvents Produtos e Seviços vazio"), vbInformation
   Call FechaDB
   Exit Sub
End If

ProdEntrada.MoveFirst

Do While Not ProdEntrada.EOF
   If pes.State = 1 Then
      pes.Close: Set pes = Nothing
   End If
   pes.Open "Select * from pessoa where chPessoa = ('" & ProdEntrada!chPessoa & "')", db, 3, 3
   If Not pes.EOF Then
      Call TratarFornecedor
   Else
      Call TratarDespesa
   End If
   
   ProdEntrada.MoveNext
Loop

End Sub

Public Sub TratarFornecedor()

'Ele não pode estar em ProdFornec

If ProdFornec.State = 1 Then
   ProdFornec.Close: Set ProdFornec = Nothing
End If

ProdFornec.Open "Select * from produtofornecedor where chTipoProduto = ('" & ProdEntrada!chPessoa & "') and chProdutoFabrica = ('" & ProdEntrada!chTipoProduto & "')", db, 3, 3
If Not ProdFornec.EOF Then
   ProdFornec.Delete
End If

End Sub
Public Sub TratarDespesa()

'Não pode estar em Produto Entrada
If ProdFornec.State = 1 Then
   ProdFornec.Close: Set ProdFornec = Nothing
End If

ProdFornec.Open "Select * from produtofornecedor where chTipoProduto = ('" & ProdEntrada!chPessoa & "') and chProdutoFabrica = ('" & ProdEntrada!chTipoProduto & "')", db, 3, 3
If Not ProdFornec.EOF Then
   ProdEntrada.Delete
Else
   ProdFornec.AddNew
   ProdFornec!chTipoProduto = ProdEntrada!chPessoa
   ProdFornec!chProdutoFabrica = ProdEntrada!chTipoProduto
   ProdFornec!chCentroDeCusto = ProdEntrada!chProdutoFabrica
   ProdFornec.Update
   ProdEntrada.Delete
End If

End Sub

Private Sub Command13_Click()

Call Rotina_AbrirBanco

hnfd.Open "Select * from notafiscaldetprod", db, 3, 3
If hnfd.EOF Then
   MsgBox ("Historico Det Prod Vazio"), vbInformation
   Call FechaDB
   Exit Sub
End If

hnfd.MoveFirst

Do While Not hnfd.EOF

'      If hnfd!chPessoa = "SISTEMA" Then
'         MsgBox ("Sistema - ") & hnfd!nfdValorParcela
'      End If
'
'      If hnfd!chCodProduto = "SISTEMA" Then
'         MsgBox ("Sistema 2 - ") & hnfd!nfdValorParcela
'      End If

      If ccc.State = 1 Then
         ccc.Close: Set ccc = Nothing
      End If

      ccc.Open "Select * from centrodecustoconatbil where chProduto = ('" & hnfd!chPessoa & "')", db, 3, 3
      If ccc.EOF Then
         If ccc.State = 1 Then
            ccc.Close: Set ccc = Nothing
         End If

         ccc.Open "Select * from centrodecustoconatbil where chProduto = ('" & hnfd!chCodProduto & "')", db, 3, 3
         If Not ccc.EOF Then
            ccc!cccValorProduto = ccc!cccValorProduto + hnfd!nfdValorParcela
            ccc.Update
         Else
            MsgBox ("Valor não encontrado - ") & hnfd!chPessoa & " - " & hnfd!chCodProduto
         End If
      Else
         ccc!cccValorPessoa = ccc!cccValorPessoa + hnfd!nfdValorParcela
         ccc.Update
      End If
      
      hnfd.MoveNext
      
 Loop
 
 MsgBox ("Fim do serviço"), vbInformation
   
End Sub

Private Sub Command14_Click()


Call Rotina_AbrirBanco

ccc.Open "Select * from centrodecustoconatbil", db, 3, 3
If ccc.EOF Then
   MsgBox ("Centro de Custo Contabil Vazio"), vbInformation
   Call FechaDB
   Exit Sub
End If

ccc.MoveFirst

Do While Not ccc.EOF
   ccc!cccValorPago = 0 'ccc!cccValorPessoa + ccc!cccValorProduto
   ccc!cccValorProduto = 0
   ccc!cccValorPessoa = 0
   ccc.Update
   ccc.MoveNext
Loop

MsgBox ("Fim de serviço"), vbInformation

Call FechaDB
         
         
End Sub

Private Sub Command15_Click()
Dim dataInicio As String
Dim AcumulaValor As Currency
Dim CodClassifCusto As String
Dim Status As String

CodClassifCusto = "1.1"
dataInicio = "2022-12-31"
AcumulaValor = 0
Status = 1


Call Rotina_AbrirBanco

ctr.Open "Select * from  contas_a_receber where ctrStatus = ('" & Status & "')", db, 3, 3
If ctr.EOF Then
   MsgBox ("Contas a Receber Vazio")
   Call FechaDB
   Exit Sub
End If

ctr.MoveFirst

Do While Not ctr.EOF
   AcumulaValor = AcumulaValor + ctr!ctrValorDaBoleta
   ctr.MoveNext
Loop

If ccc.State = 1 Then
   ccc.Close: Set ccc = Nothing
End If
      
ccc.Open "Select * from centrodecustoconatbil where cccClassificacaoCusto = ('" & CodClassifCusto & "')", db, 3, 3
If ccc.EOF Then
   ccc.AddNew
End If

   ccc!cccValorPago = AcumulaValor
   ccc.Update
   
   Call FechaDB
End Sub

Private Sub Command16_Click()
Call Rotina_AbrirBanco

'Produto Entrada

Prod.Open "Select * from produtoentrada", db, 3, 3
If Prod.EOF Then
   MsgBox ("Produto Entrada vazio"), vbInformation
   Call FechaDB
   Exit Sub
End If

Prod.MoveFirst

Do While Not Prod.EOF

   If ccc.State = 1 Then
      ccc.Close: Set ccc = Nothing
   End If
   
   ccc.Open "Select * from centrodecustoconatbil where chProduto = ('" & Prod!chTipoProduto & "')", db, 3, 3
   If Not ccc.EOF Then
      Prod!pinClassificacao = ccc!cccClassificacaoCusto
      Prod.Update
   End If
   
   Prod.MoveNext
   
Loop

MsgBox ("Fim do Serviço."), vbInformation

Call FechaDB

End Sub

Private Sub Command17_Click()
Dim ChaveAnterior As String
Dim centrodecusto As String
Dim Grupo As String
Dim SubGrupo As String

Call Rotina_AbrirBanco

ProdFornec.Open "Select * from produtofornecedor", db, 3, 3
If Not ProdFornec.EOF Then
   ProdFornec.MoveFirst
   ChaveAnterior = ProdFornec!chTipoProduto
   Do While Not ProdFornec.EOF
      If Not IsNull(ProdFornec!pinCentroDeCusto) Then
         ChaveAnterior = ProdFornec!chTipoProduto
         centrodecusto = ProdFornec!pinCentroDeCusto
         Grupo = ProdFornec!pinGrupoCentroDeCusto
         SubGrupo = ProdFornec!pinSubGrupoCentroDeCusto
      End If
      If ChaveAnterior = ProdFornec!chTipoProduto Then
         ProdFornec!pinCentroDeCusto = centrodecusto
         ProdFornec!pinGrupoCentroDeCusto = Grupo
         ProdFornec!pinSubGrupoCentroDeCusto = SubGrupo
         ProdFornec.Update
       End If
   ProdFornec.MoveNext
    
   Loop
   
 End If
 MsgBox ("Fim de serviço"), vbInformation
      
         
End Sub

Private Sub Command18_Click()
Dim ChaveAnterior As String
Dim centrodecusto As String
Dim Grupo As String
Dim SubGrupo As String

Call Rotina_AbrirBanco

Prod.Open "Select * from produtoentrada", db, 3, 3
If Not Prod.EOF Then
   Prod.MoveFirst
   ChaveAnterior = Empty
   Do While Not Prod.EOF
      If Not IsNull(Prod!pinCentroDeCusto) Then
         ChaveAnterior = Prod!chPessoa
         centrodecusto = Prod!pinCentroDeCusto
         Grupo = Prod!pinGrupoCentroDeCusto
         SubGrupo = Prod!pinSubGrupoCentroDeCusto
      End If
      If (ChaveAnterior = Prod!chPessoa) Then
         Prod!pinCentroDeCusto = centrodecusto
         Prod!pinGrupoCentroDeCusto = Grupo
         Prod!pinSubGrupoCentroDeCusto = SubGrupo
         Prod.Update
       End If
   Prod.MoveNext
    
   Loop
   
 End If
 MsgBox ("Fim de serviço"), vbInformation
      
         
End Sub

Private Sub Command19_Click()

Call Rotina_AbrirBanco

nfd.Open "Select * from historiconotafiscaldetprod", db, 3, 3
If nfd.EOF Then
   MsgBox ("Nota Fiscal DetProd vazio"), vbInformation
   Call FechaDB
   Exit Sub
End If

nfd.MoveFirst

Do While Not nfd.EOF
   If Prod.State = 1 Then
      Prod.Close: Set Prod = Nothing
   End If
   
   Prod.Open "Select * from produtoentrada where chTipoProduto = ('" & nfd!chCodProduto & "')", db, 3, 3
   If Not Prod.EOF Then
      nfd!nfdCentroDeCusto = Prod!pinCentroDeCusto
      nfd!nfdGrupoCentroDeCusto = Prod!pinGrupoCentroDeCusto
      nfd!nfdSubGrupoCentroDeCusto = Prod!pinSubGrupoCentroDeCusto
      nfd.Update
   End If
   
   nfd.MoveNext
   
Loop

MsgBox ("Fim da carga de produtoentrada"), vbInformation

nfd.MoveFirst

Do While Not nfd.EOF
   If Prod.State = 1 Then
      Prod.Close: Set Prod = Nothing
   End If
   
   Prod.Open "Select * from produtofornecedor where chTipoProduto = ('" & nfd!chPessoa & "') and chProdutoFabrica = ('" & nfd!chCodProduto & "')", db, 3, 3
   If Not Prod.EOF Then
      nfd!nfdCentroDeCusto = Prod!pinCentroDeCusto
      nfd!nfdGrupoCentroDeCusto = Prod!pinGrupoCentroDeCusto
      nfd!nfdSubGrupoCentroDeCusto = Prod!pinSubGrupoCentroDeCusto
      nfd.Update
   End If
   
   nfd.MoveNext
   
Loop

MsgBox ("Fim de serviço"), vbInformation

Call FechaDB
   
End Sub

'--------------------------------------------------------------------------------
' Project    :       SHB
' Procedure  :       Command20_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LAPTOP-S2VJ78M7
' Date-Time  :       7/2/2023-18:05:01
'
' Parameters :
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
' Project    :       SHB
' Procedure  :       Command20_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LAPTOP-S2VJ78M7
' Date-Time  :       7/2/2023-18:08:02
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Command20_Click()

Dim DataProcura As String
Dim AcumulaReg As Integer

DataProcura = "99"
AcumulaReg = 0

Call Rotina_AbrirBanco

nfd.Open "Select * from historiconotafiscaldetprod", db, 3, 3
If nfd.EOF Then
   MsgBox ("ERRO: Historico nfdDetProd vazio."), vbInformation
   Call FechaDB
   Exit Sub
End If

nfd.MoveFirst

Do While Not nfd.EOF

   If IsNull(nfd!nfdDataPagamento) Then
      If ctp.State = 1 Then
         ctp.Close: Set ctp = Nothing
      End If
      
      ctp.Open "Select * from historicocontaspagar where chPessoa = ('" & nfd!chPessoa & "') and chNotaFiscal = ('" & nfd!chNotaFiscalEntrada & "')", db, 3, 3
      If Not ctp.EOF Then
         If Not IsNull(ctp!ctpDataPagamento) Then
            nfd!nfdDataPagamento = ctp!ctpDataPagamento
            nfd.Update
         End If
      Else
         If ctp.State = 1 Then
            ctp.Close: Set ctp = Nothing
         End If
         ctp.Open "Select * from contas_a_pagar where chPessoa = ('" & nfd!chPessoa & "') and chNotaFiscal = ('" & nfd!chNotaFiscalEntrada & "')", db, 3, 3
         If Not ctp.EOF Then
            If Not IsNull(ctp!ctpDataPagamento) Then
               nfd!nfdDataPagamento = ctp!ctpDataPagamento
               nfd.Update
            Else
               AcumulaReg = AcumulaReg + 1
            End If
          'Else
          '   MsgBox ("Registro checado - ") & nfd!chPessoa & " - " & nfd!chNotaFiscalEntrada & " - " & nfd!nfdValorParcela
         End If
      End If
   End If
   nfd.MoveNext
Loop

MsgBox ("Fim do serviço. ") & AcumulaReg, vbInformation

End Sub
 
Private Sub Command21_Click()
Dim Chave As Integer

Chave = 6

Call Rotina_AbrirBanco

pes.Open "Select * from detalhenegociacao", db, 3, 3
If pes.EOF Then
   MsgBox ("Erro no acesso a pessoa"), vbInformation
   Call FechaDB
   Exit Sub
End If

pes.MoveFirst

Do While Not pes.EOF
 
   Nome = pes!chPessoaFunc
   
   If neg.State = 1 Then
      neg.Close: Set neg = Nothing
   End If
   
   neg.Open "Select negContrato, chPessoa from negociacao where chNumPedido IN (Select chNumPedido from detalhenegociacao where chProduto = ('" & Nome & "'))", db, 3, 3
   If Not neg.EOF Then
      neg.MoveFirst
      Do While Not neg.EOF
         pes!chContrato = neg!negContrato
         'MsgBox ("Contrato - ") & neg!negContrato & " - " & neg!chPessoa
         pes.Update
         neg.MoveNext
      Loop

   End If
   pes.MoveNext

Loop

MsgBox ("Fim do serviço")


End Sub

Private Sub Command22_Click()
Call Rotina_AbrirBanco
Dim Inicio As String
Dim Fim As String

Inicio = "2023-06-31"
Fim = "2023-08-01"

ctp.Open "Select * from historicocontaspagar where ctpDataPagamento > ('" & Inicio & "') and ctpDataPagamento < ('" & Fim & "')", db, 3, 3
If ctp.EOF Then
   MsgBox ("ERRO: Historico de Contas a Pagar vazio."), vbCritical
   Call FechaDB
   Exit Sub
End If

ctp.MoveFirst

Do While Not ctp.EOF
   If nfd.State = 1 Then
      nfd.Close: Set nfd = Nothing
   End If
   
   nfd.Open "Select * from historiconotafiscaldesdobramento where chPessoa = ('" & ctp!chPessoa & "') and chNotaFiscalEntrada = ('" & ctp!chNotafiscal & "') and nfdFaturaNumero = ('" & ctp!chFatura & "')", db, 3, 3
   If nfd.EOF Then
      MsgBox ("Nota fiscal não encontrada no Detalhe de Nota Fiscal - ") & ctp!chPessoa & "-" & ctp!chNotafiscal
   Else
      'If IsNull(nfd!nfddatapagamento) Then
         nfd!nfdDataPagamento = Format$(ctp!ctpDataPagamento, "yyyy-mm-dd")
         nfd.Update
      'End If
   End If
   
   ctp.MoveNext
Loop

MsgBox ("Fim do Serviço."), vbInformation

End Sub

Private Sub Command24_Click()
Call Rotina_AbrirBanco

ctr.Open "Select * from contas_a_receber", db, 3, 3

ctr.MoveFirst
Do While Not ctr.EOF
   If ctr!ctrGrupoCentroDeCusto = "1" Then
      ctr!ctrGrupoCentroDeCusto = "01"
   End If
   If ctr!ctrGrupoCentroDeCusto = "2" Then
      ctr!ctrGrupoCentroDeCusto = "02"
   End If
   If ctr!ctrSubGrupoCentroDeCusto = "1" Then
      ctr!ctrSubGrupoCentroDeCusto = "01"
   End If
   If ctr!ctrSubGrupoCentroDeCusto = "2" Then
      ctr!ctrSubGrupoCentroDeCusto = "02"
   End If
   If ctr!ctrSubGrupoCentroDeCusto = "3" Then
      ctr!ctrSubGrupoCentroDeCusto = "03"
   End If
   
   ctr.Update
   
   ctr.MoveNext
Loop
MsgBox ("Fim de Serviço")
End Sub

Private Sub Command25_Click()
Dim FimDet As Integer
Dim NotaFiscalAnterior As String
Dim pessoaAnterior As String
Dim Encontrei As Integer

Call Rotina_AbrirBanco

dnfe.Open "select * from notafiscaldetprod", db, 3, 3
If dnfe.EOF Then
   'MsgBox ("Nota Fiscal sem Detalhe."), vbInformation
   FimDet = 1
Else
   FimDet = 0
   NotaFiscalAnterior = Empty
   pessoaAnterior = Empty
   dnfe.MoveFirst
   Do While FimDet = 0
'      If dnfe!chPessoa = "PROTCAP" Then
'         MsgBox "PROTCAP"
'      End If
      If Not (dnfe!chPessoa = pessoaAnterior And dnfe!chNotaFiscalEntrada = NotaFiscalAnterior) Then
         
         pessoaAnterior = dnfe!chPessoa
         NotaFiscalAnterior = dnfe!chNotaFiscalEntrada
         If ctp.State = 1 Then
            ctp.Close: Set ctp = Nothing
         End If
         ctp.Open "Select * from contas_a_pagar where chPessoa = ('" & dnfe!chPessoa & "') and chNotaFiscal = ('" & dnfe!chNotaFiscalEntrada & "') and ctpStatus = 0", db, 3, 3
         If ctp.EOF Then
            Encontrei = 0
         Else
            Encontrei = 1
         End If
      End If
      
      If hdnfe.State = 1 Then
         hdnfe.Close: Set hdnfe = Nothing
      End If
      hdnfe.Open "Select * from historiconotafiscaldetprod where chPessoa = ('" & dnfe!chPessoa & "') and chNotaFiscalEntrada = ('" & dnfe!chNotaFiscalEntrada & "') and chCodProduto = ('" & dnfe!chCodProduto & "')", db, 3, 3
      If hdnfe.EOF Then
         hdnfe.AddNew
      End If
      hdnfe!chPessoa = dnfe!chPessoa
      hdnfe!chNotaFiscalEntrada = dnfe!chNotaFiscalEntrada
      hdnfe!chCodProduto = dnfe!chCodProduto
      hdnfe!chProdutoFabrica = dnfe!chProdutoFabrica
      hdnfe!nfdQtd = dnfe!nfdQtd
      hdnfe!nfdPU = dnfe!nfdPU
      hdnfe!nfdValorDaCompra = dnfe!nfdValorDaCompra
      hdnfe!nfdQtdParcelas = dnfe!nfdQtdParcelas
      hdnfe!nfdValorParcela = dnfe!nfdValorParcela
      hdnfe!nfdCentroDeCusto = dnfe!nfdCentroDeCusto
      hdnfe!nfdGrupoCentroDeCusto = dnfe!nfdGrupoCentroDeCusto
      hdnfe!nfdSubGrupoCentroDeCusto = dnfe!nfdSubGrupoCentroDeCusto
      
      If Encontrei = 1 Then
         hdnfe!nfdDataPagamento = ctp!ctpDataPagamento
      End If

      ultimoRegistro = "Rotina Grava Det Produto. - " & dnfe!chPessoa & " - " & dnfe!chNotaFiscalEntrada & " - " & dnfe!chCodProduto & " - " & dnfe!chProdutoFabrica

      hdnfe.Update
      
      If Encontrei = 0 Then
         dnfe.Delete
      End If

'O encontrei = 0 significa que não há mais contas a pagar para este cliente nota fiscal.
'O registro em DetProd tem que ficar tanto no mes quanto no historico.
'Será retirado do mes qdo não hover mais o financeiro do mes.

      dnfe.MoveNext
      If dnfe.EOF Then
         FimDet = 1
      End If
   Loop
End If
         
'db.CommitTrans

Call FechaDB

Exit Sub

End Sub

Private Sub Command9_Click()
Dim AcumulaValor As Currency
ValorTotal = 0
AcumulaValor = 0

Call Rotina_AbrirBanco

ctp.Open "Select * from contas_a_pagar", db, 3, 3

ctp.MoveFirst

Do While Not ctp.EOF
   
   If nfd.State = 1 Then
      nfd.Close: Set nfd = Nothing
   End If
   
   nfd.Open "Select * from notafiscaldetprod where chPessoa = ('" & ctp!chPessoa & "') and chNotaFiscalEntrada = ('" & ctp!chNotafiscal & "')", db, 3, 3
   If nfd.EOF Then
      If hnfd.State = 1 Then
         hnfd.Close: Set hnfd = Nothing
      End If
      
      'hnfd.Open "Select * from historiconotafiscaldetprod where chPessoa = ('" & ctp!chPessoa & "') and chNotaFiscalEntrada = ('" & ctp!chNotaFiscal & "')", db, 3, 3
      hnfd.Open "Select * from historiconotafiscaldetprod where chPessoa = ('" & ctp!chPessoa & "') and chNotaFiscalEntrada = ('" & ctp!chNotafiscal & "')", db, 3, 3
      If hnfd.EOF Then
         AcumulaValor = AcumulaValor + ctp!ctpValorDaBoleta
         MsgBox ("Não encontrei o detalhe em histroico ") & ctp!chPessoa & " - " & ctp!chNotafiscal
      Else
         hnfd.MoveFirst
         Do While Not hnfd.EOF
         
            If nfd.State = 1 Then
               nfd.Close: Set nfd = Nothing
            End If
            
            nfd.Open "Select * from notafiscaldetprod where chPessoa = ('" & hnfd!chPessoa & "') and chNotaFiscalEntrada = ('" & hnfd!chNotaFiscalEntrada & "') and chCodProduto = ('" & hnfd!chCodProduto & "')", db, 3, 3
            If nfd.EOF Then
               nfd.AddNew
            End If
            
            nfd!chPessoa = hnfd!chPessoa
            nfd!chNotaFiscalEntrada = hnfd!chNotaFiscalEntrada
            nfd!chCodProduto = hnfd!chCodProduto
            nfd!chProdutoFabrica = hnfd!chProdutoFabrica
            nfd!chProdutoFabrica = hnfd!chProdutoFabrica
            nfd!nfdQtd = hnfd!nfdQtd
            nfd!nfdPU = hnfd!nfdPU
            nfd!nfdValorDaCompra = hnfd!nfdValorDaCompra
            nfd!nfdQtdParcelas = hnfd!nfdQtdParcelas
            nfd!nfdValorParcela = hnfd!nfdValorParcela
            ValorTotal = ValorTotal + nfd!nfdValorParcela
            nfd.Update
            
            'hnfd.Delete
            
            hnfd.MoveNext
         Loop
      End If
      
      
   End If
   
   ctp.MoveNext
   
Loop

MsgBox "Valor Total Criado = " & ValorTotal
MsgBox "Valor Total Sem Det Prod = " & AcumulaValor

End Sub



'Private Sub cmd_Click()
'Dim fim As Integer
'fim = 0
'TabProdutoTerceiros.MoveFirst
'Do While fim = 0
'   If TabProdutoTerceiros("campo1") = 1 Then
'      TabProdutoTerceiros.Delete
'   End If
'   TabProdutoTerceiros.MoveNext
'   If TabProdutoTerceiros.EOF Then
'      fim = 1
'   Else
'      fim = 0
'   End If
'Loop

'MsgBox "Fim da limpeza"

'End Sub

'Private Sub cmdLimpaCtaPagar_Click()
'Dim fim As Integer
'fim = 0
'If TabCtaPagar.RecordCount > 0 Then
'   TabCtaPagar.MoveFirst
'   If TabCtaPagar.EOF Then
'      fim = 1
'   End If
'Else
'   fim = 1
'End If
'Do While fim = 0
'   TabCtaPagar.Delete
'   TabCtaPagar.MoveNext
'   If TabCtaPagar.EOF Then
'      fim = 1
'   Else
'      fim = 0
 '  End If
'Loop
'End Sub

'Private Sub cmdLimpaCtaReceber_Click()
'Dim fim As Integer
'fim = 0
'If TabCtaReceber.RecordCount > 0 Then
'   TabCtaReceber.MoveFirst
'   If TabCtaReceber.EOF Then
'      fim = 1
'   End If
'Else
'   fim = 1
'End If
'Do While fim = 0
'   TabCtaReceber.Delete
 '  TabCtaReceber.MoveNext
 '  If TabCtaReceber.EOF Then
 '     fim = 1
 '  Else
 '     fim = 0
 '  End If
'L'oop

'End Sub
'Private Sub cmdLimpaHistorico_Click()
'Dim fim As Integer
'fim = 0
'If TabTelefone.RecordCount > 0 Then
'   TabTelefone.MoveFirst
'   If TabTelefone.EOF Then
'      fim = 1
'   End If
'Else
'   fim = 1
'End If
'Do While fim = 0
'   TabTelefone.Delete
'   TabTelefone.MoveNext
'   If TabTelefone.EOF Then
'      fim = 1
'   Else
'      fim = 0
'   End If
'Loop
'fim = 0
'If TabNfDesdobrComp.RecordCount > 0 Then
'   TabNfDesdobrComp.MoveFirst
'   If TabNfDesdobrComp.EOF Then
'      fim = 1
'   End If
'Else
'   fim = 1
'End If
'Do While fim = 0
'   TabNfDesdobrComp.Delete
'   TabNfDesdobrComp.MoveNext
'   If TabNfDesdobrComp.EOF Then
'      fim = 1
'   Else
'      fim = 0
'   End If
'Loop
'fim = 0
'If TabNotaFiscalEntrada.RecordCount > 0 Then
'   TabNotaFiscalEntrada.MoveFirst
'   If TabNotaFiscalEntrada.EOF Then
'      fim = 1
'   End If
'Else
'   fim = 1
'End If
'Do While fim = 0
 '  TabNotaFiscalEntrada.Delete
 '  TabNotaFiscalEntrada.MoveNext
 '  If TabNotaFiscalEntrada.EOF Then
 '     fim = 1
 '  Else
 '     fim = 0
 '  End If
'L 'oop
'If TabHistCtaReceber.RecordCount > 0 Then
'   TabHistCtaReceber.MoveFirst
'   If TabHistCtaReceber.EOF Then
'      Fim = 1
'   End If
'Else
'   Fim = 1
'End If
'Do While Fim = 0
'   TabHistCtaReceber.Delete
'   TabHistCtaReceber.MoveNext
'   If TabHistCtaReceber.EOF Then
'      Fim = 1
'   Else
'      Fim = 0
'   End If
'Loop
'Fim = 0 '
'If TabHistoricoDetNeg.RecordCount > 0 Then
'   TabHistoricoDetNeg.MoveFirst
'   If TabHistoricoDetNeg.EOF Then
'      Fim = 1
'   End If
'Else
'   Fim = 1
'End If
'Do While Fim = 0
'   TabHistoricoDetNeg.Delete
'   TabHistoricoDetNeg.MoveNext
'   If TabHistoricoDetNeg.EOF Then
'      Fim = 1
'   Else
'      Fim = 0
'   End If
'Loop
'Fim = 0
'If TabHistoricoNegociacao.RecordCount > 0 Then
'   TabHistoricoNegociacao.MoveFirst
'   If TabHistoricoNegociacao.EOF Then
'      Fim = 1
'   End If
'Else
'   Fim = 1
'End If'

'Do While Fim = 0
'   TabHistoricoNegociacao.Delete
'   TabHistoricoNegociacao.MoveNext
'   If TabHistoricoNegociacao.EOF Then
'      Fim = 1
'   Else
'      Fim = 0
'   End If
'Loop

'Fim = 0
'If TabHistNfDesdobrComp.RecordCount > 0 Then
'   TabHistNfDesdobrComp.MoveFirst
'   If TabHistNfDesdobrComp.EOF Then
'      Fim = 1
'   End If
'Else
'   Fim = 1
'End If
'Do While Fim = 0
'   TabHistNfDesdobrComp.Delete
'   TabHistNfDesdobrComp.MoveNext
'   If TabHistNfDesdobrComp.EOF Then'
'      Fim = 1
'   Else
'      Fim = 0
'   End If
'Loop
'Fim = 0
'If TabHistNotaFiscalDetProd.RecordCount > 0 Then
'   TabHistNotaFiscalDetProd.MoveFirst
'   If TabHistNotaFiscalDetProd.EOF Then
'      Fim = 1
'   End If
'Else
'   Fim = 1
'End If
'Do While Fim = 0
'   TabHistNotaFiscalDetProd.Delete
'   TabHistNotaFiscalDetProd.MoveNext
'   If TabHistNotaFiscalDetProd.EOF Then
'      Fim = 1
'   Else
'      Fim = 0
'   End If
'Loop
'End Sub

'Private Sub cmdLimpaPreco_Click()
'Dim fim As Integer
'tabProdutoPreco.MoveFirst
'Do While fim = 0
'   tabProdutoPreco.Delete
'   tabProdutoPreco.MoveNext
'''   If tabProdutoPreco.EOF Then
'      fim = 1
'   Else
'      fim = 0
'   End If
'Loop
'End Sub

'Private Sub Command1_Click()
'Dim fim As Integer
'fim = 0
'If Tabpessoa.RecordCount > 0 Then
'   Tabpessoa.MoveFirst
'   If Tabpessoa.EOF Then
'      fim = 1
'   End If
'Else
'   fim = 1
'End If
'Do While fim = 0
'   If Tabpessoa("chpessoa") = "SUED" Or Tabpessoa("chpessoa") = "TECHOCEAN" Then
 '     Tabpessoa.MoveNext
 '     If Tabpessoa.EOF Then
 '        fim = 1
  '    Else
'         fim = 0
'      End If
'   Else
 '     Tabpessoa.Delete
 '     Tabpessoa.MoveNext
 '     If Tabpessoa.EOF Then
 '        fim = 1
 '     Else
  '       fim = 0
  '    End If
'   End If
'Loop
'E'nd Sub

'Private Sub Command2_Click()
'Dim fim As Integer
'F 'im = 0
'If TabDetNeg.RecordCount > 0 Then
'   TabDetNeg.MoveFirst
'   If TabDetNeg.EOF Then
'      fim = 1
'   End If
'Else
'   fim = 1
'End If
'Do While fim = 0
'   TabDetNeg.Delete
'   TabDetNeg.MoveNext
'   If TabDetNeg.EOF Then
'      fim = 1
'   Else
'      fim = 0
'   End If
'Loop
'fim = 0
'If TabNegociacao.RecordCount > 0 Then
'   TabNegociacao.MoveFirst
'   If TabNegociacao.EOF Then
'      fim = 1
'   End If
'Else
'   fim = 1
'End If
'Do While fim = 0
'  TabNegociacao.Delete
'   TabNegociacao.MoveNext
 '  If TabNegociacao.EOF Then
  '    fim = 1
'   Else
'      fim = 0
'   End If
'Loop
'End Sub

'Private Sub Command3_Click()
'Dim fim As Integer
'F 'im = 0
'TabCarteira_Rep.MoveFirst
'Do While fim = 0
'   If Not (TabCarteira_Rep("repregiao") = "CLIENTE FÁBRICA") Then
'      TabCarteira_Rep.Delete
'   End If
'   TabCarteira_Rep.MoveNext
'   If TabCarteira_Rep.EOF Then
'      fim = 1
'   Else
'      fim = 0
'   End If
'Loop''

'MsgBox "Fim da limpeza"

'End Sub

'Private Sub Command4_Click()
'Dim fim As Integer
'fim = 0
'TabCarteira_Promot.MoveFirst
'Do While fim = 0
'   If Not (TabCarteira_Promot("CHPESSOA") = "NENHUM") Then
'      TabCarteira_Promot.Delete
'   End If
'   TabCarteira_Promot.MoveNext
'   If TabCarteira_Promot.EOF Then
'      fim = 1
'   Else
'      fim = 0
'   End If
'L'oop'

'MsgBox "Fim da limpeza"

'End Sub

'Private Sub Command5_Click()
'Dim fim As Integer
'fim = 0
'TabProdutoEntrada.MoveFirst
'Do While fim = 0
'   If Not (TabProdutoEntrada("pinclassificacao") = "N") Then
 '     TabProdutoEntrada.Delete
 '  End If
 '  TabProdutoEntrada.MoveNext
 '  If TabProdutoEntrada.EOF Then
 '     fim = 1
 '  Else
 '     fim = 0
 '  End If
'L 'oop

'MsgBox "Fim da limpeza"

'End Sub

'Private Sub Digits_Click(Index As Integer)
'    If ClearDisplay Then
'        Display.Caption = ""
 '       ClearDisplay = False
 '   End If
 '   Display.Caption = Display.Caption + Digits(Index).Caption
'E 'nd Sub

'Private Sub Div_Click()
'    Operand1 = Val(Display.Caption)
 '   Operator = "/"
 '   Display.Caption = ""
'E'nd Sub

Private Sub DotBttn_Click()
    If ClearDisplay Then
        Display.Caption = ""
        ClearDisplay = False
    End If
    If InStr(Display.Caption, ".") Then
        Exit Sub
    Else
        Display.Caption = Display.Caption + "."
    End If
End Sub

Private Sub Equals_Click()
Dim result As Double

On Error GoTo ErrorHandler
    Operand2 = Val(Display.Caption)
    If Operator = "+" Then result = Operand1 + Operand2
    If Operator = "-" Then result = Operand1 - Operand2
    If Operator = "*" Then result = Operand1 * Operand2
    If Operator = "/" And Operand2 <> "0" Then result = Operand1 / Operand2
    Display.Caption = result
    ClearDisplay = True
    Exit Sub
ErrorHandler:
    MsgBox "The operation resulted in the following error" & vbCrLf & Err.Description
    Display.Caption = "ERROR"
    ClearDisplay = True
End Sub

Private Sub Minus_Click()
    Operand1 = Val(Display.Caption)
    Operator = "-"
    Display.Caption = ""
End Sub

Private Sub Over_Click()
    If Val(Display.Caption) <> 0 Then Display.Caption = 1 / Val(Display.Caption)
End Sub

Private Sub Plus_Click()
    Operand1 = Val(Display.Caption)
    Operator = "+"
    Display.Caption = ""
End Sub

Private Sub PlusMinus_Click()
    Display.Caption = -Val(Display.Caption)
End Sub

Private Sub Times_Click()
    Operand1 = Val(Display.Caption)
    Operator = "*"
    Display.Caption = ""
End Sub
Public Sub GeraSql()
Call Rotina_AbrirBanco

If neg.State = 1 Then
   neg.Close: Set neg = Nothing
End If
neg.Open "Select negContrato from negociacao where chNumPedido IN (Select chNumPedido from detalhenegociacao where chProduto = ('" & Nome & "'))", db, 3, 3
If neg.EOF Then
   MsgBox ("Não achei"), vbInformation
End If

End Sub
