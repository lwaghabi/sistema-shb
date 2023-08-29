VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFaturaLocacao 
   Caption         =   "frmFaturaDeLocação"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   13635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Referência"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      TabIndex        =   28
      Top             =   3240
      Width           =   7335
      Begin VB.TextBox txtContratoComp 
         Alignment       =   2  'Center
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
         Left            =   1920
         TabIndex        =   32
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox txtContrato 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   30
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Complemento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComCtl2.DTPicker dtEmis 
      Height          =   495
      Left            =   6120
      TabIndex        =   24
      Top             =   5160
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   241827841
      CurrentDate     =   44328
   End
   Begin VB.TextBox dtHoje 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   19
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H008080FF&
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
      Height          =   855
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdFatura 
      BackColor       =   &H00FFFF80&
      Caption         =   "Fatura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Frame Frame 
      Caption         =   " Fatura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   6000
      TabIndex        =   18
      Top             =   5760
      Width           =   3735
      Begin VB.TextBox txtSerieFatura 
         Alignment       =   2  'Center
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
         Left            =   2280
         TabIndex        =   27
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtNumFatura 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
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
         Height          =   495
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Série"
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
         Left            =   2280
         TabIndex        =   26
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Número"
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
         Left            =   960
         TabIndex        =   25
         Top             =   150
         Width           =   975
      End
   End
   Begin VB.TextBox txtPessoa 
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
      Left            =   6120
      TabIndex        =   11
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtRazaoSocial 
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
      Left            =   6120
      TabIndex        =   10
      Top             =   2760
      Width           =   7335
   End
   Begin VB.TextBox txtUnidadeOperacional 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.TextBox txtMedicao 
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
      Left            =   9000
      TabIndex        =   8
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtComplemento 
      Alignment       =   2  'Center
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
      Index           =   0
      Left            =   6120
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid GridFatura 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   10610
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColor       =   16777152
      ForeColor       =   0
      BackColorFixed  =   16776960
      ForeColorFixed  =   0
      ForeColorSel    =   0
      BackColorBkg    =   16777152
      FormatString    =   " Medição||Cliente                     ||||||"
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
   Begin MSComCtl2.DTPicker dtFim 
      Height          =   495
      Left            =   11160
      TabIndex        =   6
      Top             =   5160
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   241827841
      CurrentDate     =   44298
   End
   Begin MSComCtl2.DTPicker dtInicio 
      Height          =   495
      Left            =   8880
      TabIndex        =   5
      Top             =   5160
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   241827841
      CurrentDate     =   44298
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Data Emissão"
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
      Index           =   10
      Left            =   6120
      TabIndex        =   23
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label 
      Caption         =   "Até"
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
      Index           =   9
      Left            =   11280
      TabIndex        =   22
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "De"
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
      Index           =   8
      Left            =   8880
      TabIndex        =   21
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "Hoje"
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
      Index           =   1
      Left            =   11040
      TabIndex        =   20
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "Medição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9000
      TabIndex        =   17
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "Complemento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6120
      TabIndex        =   15
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label 
      Caption         =   "Unidade Operacional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   6120
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label 
      Caption         =   "Razão Soicial Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6120
      TabIndex        =   13
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Período"
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
      Index           =   7
      Left            =   8880
      TabIndex        =   12
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Label Label 
      Caption         =   "Fatura de Locação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmFaturaLocacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sql As String
Dim rel As Object
Dim Relatorio As String
Dim txtNome As String
Dim txtNumPedido As String
Dim txtPedidoComp As String
Dim UnidadeOperaional As String
Dim TipoLocacao As Integer
Dim dataInicio As Date
Dim dataFim As Date
Dim DataVenc As Date
Dim TipoProduto As Byte
Dim MedicaoAnter As String
Dim Item As Integer
Dim Resp As String
Dim NumPedido As Integer
Dim AtualizaEmpresa As Integer


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Ind As Integer

dtInicio = Date
dtFim = Date
dtEmis = Date
Relatorio = "drMedicao"
dtHoje = Date
MedicaoAnter = Empty

GridFatura.Rows = 2
GridFatura.TextMatrix(1, 0) = Empty
GridFatura.TextMatrix(1, 1) = Empty
GridFatura.TextMatrix(1, 2) = Empty
GridFatura.TextMatrix(1, 3) = Empty
GridFatura.TextMatrix(1, 4) = Empty
GridFatura.TextMatrix(1, 5) = Empty
GridFatura.TextMatrix(1, 6) = Empty
GridFatura.TextMatrix(1, 7) = Empty
GridFatura.TextMatrix(1, 8) = Empty

Call Rotina_AbrirBanco

neg.Open "Select * from Negociacao where negStatus = ('" & 1 & "') and negTipoProduto = ('" & 0 & "')", db, 3, 3
If neg.EOF Then
   MsgBox ("Não há Medição para Fatura até a presente data"), vbInformation
   Call FechaDB
   Exit Sub
End If
TipoProduto = 0
Ind = 0
neg.MoveFirst
Do While Not neg.EOF
   If Not neg!chNumPedido = MedicaoAnter Then
      MedicaoAnter = neg!chNumPedido
      If ctr.State = 1 Then
         ctr.Close: Set ctr = Nothing
      End If
      ctr.Open "Select * from Contas_A_Receber where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
      If ctr.EOF Then
         ctr.Close: Set ctr = Nothing
         Exit Sub
      End If
      If dneg.State = 1 Then
         dneg.Close: Set dneg = Nothing
      End If
      dneg.Open "Select * from DetalheNegociacao where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
      If dneg.EOF Then
         dneg.Close: Set dneg = Nothing
         Exit Sub
      End If
         
      Ind = Ind + 1
      GridFatura.Rows = Ind + 1
      GridFatura.TextMatrix(Ind, 0) = neg!chNumPedido
      GridFatura.TextMatrix(Ind, 1) = neg!chNumPedidoComp
      GridFatura.TextMatrix(Ind, 2) = neg!chPessoa
      GridFatura.TextMatrix(Ind, 3) = neg!chUnidadeOperacional
      GridFatura.TextMatrix(Ind, 4) = ctr!ctrDataVencito
      GridFatura.TextMatrix(Ind, 5) = dneg!chDataInicio
      GridFatura.TextMatrix(Ind, 6) = dneg!chDataFim
      If IsNull(neg!negSerieFatura) Then
         GridFatura.TextMatrix(Ind, 7) = Empty
      Else
         GridFatura.TextMatrix(Ind, 7) = neg!negSerieFatura
      End If
      
      If IsNull(neg!negDataEmissaoFatura) Then
         GridFatura.TextMatrix(Ind, 8) = Empty
      Else
         GridFatura.TextMatrix(Ind, 8) = neg!negDataEmissaoFatura
      End If
   End If
      
   neg.MoveNext

Loop

Call FechaDB
   
End Sub
Private Sub cmdFatura_Click()
If dtInicio = dtFim Then
   MsgBox ("Ajustar o período de Medição"), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If
If txtNumFatura = Empty Then
   MsgBox ("Número da fatura não informado."), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If


Call Rotina_AbrirBanco

neg.Open "Select * from Negociacao where chNumPedido = ('" & txtMedicao & "') and negTipoProduto = ('" & TipoProduto & "')", db, 3, 3
If neg.EOF Then
   MsgBox ("Número do Pedido Inexistente. Comunicar ao analista responsável"), vbCritical
   Call FechaDB
   Exit Sub
End If

neg.MoveFirst

Do While Not neg.EOF
   If neg!negNumFatura = 0 Then
      AtualizaEmpresa = 1
   Else
      AtualizaEmpresa = 0
   End If
   neg!negNumFatura = Format$(txtNumFatura, "000")
   neg!negSerieFatura = txtSerieFatura
   neg!negDataEmissaoFatura = dtEmis
   neg!negContrato = txtContrato
   neg!negContratoComp = txtContratoComp
   neg.Update
   neg.MoveNext
Loop


Relatorio = "drFatura"
'db.begintrans
gge.Open "Select * from GeradorGeral where chAlfaNumerica = ('" & Relatorio & "')", db, 3, 3
If gge.EOF Then
   gge.AddNew
End If

gge!chAlfaNumerica = "drFatura"
gge!ggeDataHoje = dtEmis
gge!ggeDataIni = dtInicio
gge!chNumerica = Format$(dtInicio, "yyyymmdd")
gge!ggeDataFim = dtFim

gge!num2 = txtNumFatura
gge!Alfa2 = txtUnidadeOperacional
gge!Alfa3 = txtSerieFatura
gge!Data2 = DataVenc

gge.Update

'db.CommitTrans

Item = 0
Set rel = drFatura
Sql = "Select gge.ggeDatahoje, gge.ggeDataIni, gge.ggeDataFim, gge.Alfa2, gge.Alfa3, gge.chNumerica, gge.num2, Unid.AbreviaturaUnidadeMedida, "
Sql = Sql & " neg.chPessoa, neg.chUnidadeOperacional, neg.negContrato, neg.negContratoComp, neg.chNumPedido, neg.chNumPedidoComp, gge.Data2, "
Sql = Sql & " pes.pesRazaoSocial, pes.pesEndereco, pes.pesBairro, pes.pesCidade, pes.chUF, pes.pesCEP, pes.chCNPJ_CPF, pes.pesInscEst_Ident, pes.pesTelContato, "
Sql = Sql & " det.chDataInicio, det.chDataFim, det.chProduto, det.pedValorDaOperacao, det.pedQuantidadePedida, "
Sql = Sql & " det.pedPrecoUnidadePedida, det.pedValorDaDiaria, det.pedQtdDias, det.pedValorDaOperacao, det.pedAtividade, prd.prdDescCompleta, prd.chProduto, prd.prdNomeProd, prdNomeComercial, "
Sql = Sql & " ender.rua, ender.bairro, ender.cidade, ender.uf, ender.cep From supendereco ender, GeradorGeral gge, UnidadeDeMedida Unid, Negociacao neg, DetalheNegociacao det, Pessoa pes, Produto prd "
Sql = Sql & " WHERE neg.chNumpedido = ('" & txtMedicao & "') and neg.negTipoProduto = ('" & TipoProduto & "') and gge.chAlfaNumerica = ('" & Relatorio & "') and det.chProduto = prd.chProduto "
Sql = Sql & " and det.chNumpedido = neg.chNumpedido and det.chNumpedidoComp = neg.chNumPedidoComp and ender.apelido='BASE'"
Sql = Sql & " and neg.chPessoa = pes.chPessoa and Unid.chUnidadeDeMedida = det.pedUnidade "
Sql = Sql & " order by neg.chUnidadeOperacional, prd.chProduto "
AbrirRelatorio Sql, rel

'Resp = MsgBox("O número da Fatura e a emissão da mesma estão corretos???", vbExclamation + vbYesNo)
'   If Resp = vbNo Then
'      Call FechaDB
'      Exit Sub
'   End If
   
Emp.Open "Select * from Empresa", db, 3, 3
If Emp.EOF Then
   MsgBox ("Erro no acesso a Empresa. Reportar ao analista responsável"), vbCritical
   Call FechaDB
   Exit Sub
Else
   If neg.State = 1 Then
      neg.Close: Set neg = Nothing
   End If
   neg.Open "Select * from Negociacao where chNumPedido = ('" & txtMedicao & "') and negTipoProduto = ('" & TipoProduto & "')", db, 3, 3
   If neg.EOF Then
      MsgBox ("Número do Pedido Inexistente. Comunicar ao analista responsável"), vbCritical
      Call FechaDB
      Exit Sub
   End If
   If AtualizaEmpresa = 1 Then
      NumPedido = txtNumFatura
      If (txtSerieFatura = "LR") Or (txtSerieFatura = "A") Then
         Emp!empNumFatura = Format$(txtNumFatura, "000")
      Else
         Emp!empNumFaturaE = Format$(txtNumFatura, "000")
      End If
      Emp.Update
   End If
End If

Call FechaDB

End Sub

Private Sub GridFatura_Click()
Dim Limite As Integer
Dim IndLinha As Integer

Limite = GridFatura.Rows

IndLinha = GridFatura.Row

If GridFatura.TextMatrix(IndLinha, 0) = "" Then
   MsgBox "Clicar em linha com conteúdo."
   Exit Sub
End If

txtPessoa = GridFatura.TextMatrix(IndLinha, 2)

Call Rotina_AbrirBanco

pes.Open "Select * from Pessoa where chPessoa = ('" & GridFatura.TextMatrix(IndLinha, 2) & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Cliente não encontrado. Comuniicar ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
End If


txtRazaoSocial = pes!pesRazaoSocial
txtUnidadeOperacional = GridFatura.TextMatrix(IndLinha, 3)
txtNome = txtPessoa
txtPedidoComp = GridFatura.TextMatrix(IndLinha, 1)
txtNumPedido = GridFatura.TextMatrix(IndLinha, 0)
txtMedicao = GridFatura.TextMatrix(IndLinha, 0)
DataVenc = GridFatura.TextMatrix(IndLinha, 4)
txtSerieFatura = GridFatura.TextMatrix(IndLinha, 7)
If GridFatura.TextMatrix(IndLinha, 8) = Empty Then
   dtEmis = Date
Else
   dtEmis = GridFatura.TextMatrix(IndLinha, 8)
End If

neg.Open "Select * From Negociacao where chNumPedido = ('" & txtNumPedido & "') and negTipoProduto = ('" & 0 & "')", db, 3, 3
If neg.EOF Then
   MsgBox ("Pedido inexistente. Reportar ao analista responsável"), vbCritical
   Call FechaDB
   Exit Sub
Else
   txtContrato = neg!negContrato
   dtInicio = neg!negInicioMedicao
   dtFim = neg!negFinalMedicao
   If IsNull(neg!negContratoComp) Then
      txtContratoComp = Empty
   Else
      txtContratoComp = neg!negContratoComp
   End If
   
   If neg!negNumFatura > 0 Then
      MsgBox ("Esta Fatura já foi Impressa. A reimpressão pode ser efetuada???"), vbExclamation
      txtNumFatura = neg!negNumFatura
   Else
      Emp.Open "Select * from Empresa", db, 3, 3
      If Emp.EOF Then
         MsgBox ("Banco de dados sem registro da empresa. Reportar ao analista responsável"), vbCritical
         Call FechaDB
         Exit Sub
      Else
         If pes!pesClassFiscal = "Lucro Real" Then
            txtNumFatura = Format$((Emp!empNumFatura + 1), "000")
         Else
            txtNumFatura = Format$((Emp!empNumFaturaE + 1), "000")
         End If
      End If
   End If
End If

Dia = Day(Date)
Mes = Month(Date)
ano = Year(Date)

Call FechaDB

End Sub


Private Sub txtNumFatura_LostFocus()
If Not IsNumeric(txtNumFatura) Then
   MsgBox ("Esta informação só pode conter números."), vbCritical
   cmdSair.SetFocus
   Exit Sub
End If
   
End Sub
