VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMedicao 
   Caption         =   "frmMedicao"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16050
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   16050
   StartUpPosition =   2  'CenterScreen
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
      Left            =   13800
      TabIndex        =   13
      Top             =   480
      Width           =   1935
   End
   Begin VB.Frame Frame 
      Height          =   5175
      Index           =   0
      Left            =   8160
      TabIndex        =   9
      Top             =   1440
      Width           =   7695
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Envio Para Aprovação"
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Frame Frame 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Objeto da Locação"
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
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   3360
         Width           =   4455
         Begin VB.OptionButton optPessoal 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Pessoal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2520
            TabIndex        =   4
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton optEquipamento 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Equipamento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H0000FFFF&
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4320
         Width           =   2655
      End
      Begin VB.CommandButton cmdConsolidado 
         BackColor       =   &H0080FF80&
         Caption         =   "IMP MEDIÇÃO"
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2520
         Width           =   2655
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
         Left            =   2640
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2055
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
         Left            =   2640
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.CommandButton cmdGerarMedicao 
         BackColor       =   &H000000FF&
         Caption         =   "Gerar Analítico de Medição"
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
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4320
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtFim 
         Height          =   495
         Left            =   2760
         TabIndex        =   2
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   391380993
         CurrentDate     =   44298
      End
      Begin MSComCtl2.DTPicker dtInicio 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   2640
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
         Format          =   391380993
         CurrentDate     =   44298
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
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   7335
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
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "Até"
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
         Index           =   9
         Left            =   2760
         TabIndex        =   24
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "De"
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
         Index           =   8
         Left            =   240
         TabIndex        =   23
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "Período"
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
         Index           =   7
         Left            =   240
         TabIndex        =   22
         Top             =   2160
         Width           =   4335
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
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   2535
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
         Left            =   2640
         TabIndex        =   20
         Top             =   840
         Visible         =   0   'False
         Width           =   3255
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
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1575
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
         Left            =   2640
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridMedicao 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9128
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16777152
      ForeColor       =   0
      BackColorFixed  =   16776960
      ForeColorFixed  =   0
      ForeColorSel    =   0
      BackColorBkg    =   16777152
      FormatString    =   " Medição|Comp|Cliente                     |Unid. Operac."
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
      Left            =   13800
      TabIndex        =   14
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "Controle e Emissão de Mapa de Medição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmMedicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim Rel As Object
Dim Relatorio As String
Dim txtNome As String
Dim txtNumPedido As String
Dim txtPedidoComp As String
Dim UnidadeOperaional As String
Dim TipoLocacao As Integer
Dim dataInicio As Date
Dim dataFim As Date
Dim TipoProduto As Byte
Dim Resp As String
Dim Ind As Integer
Dim PedidoAnterior As String


Private Sub cmdConsolidado_Click()
 
If dtInicio = dtFim Then
   MsgBox ("Ajustar o período de Medição"), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If

If optEquipamento = False And optPessoal = False Then
   MsgBox ("Para emissão de Mapa Consolidado faz-se necessário uma opção"), vbCritical
   cmdSair.SetFocus
   Exit Sub
End If

Call Rotina_AbrirBanco

Relatorio = "drMedicaoEqptoPes"
db.BeginTrans
gge.Open "Select * from geradorgeral where chAlfaNumerica = ('" & Relatorio & "')", db, 3, 3
If gge.EOF Then
   gge.AddNew
End If

gge!chAlfaNumerica = Relatorio
gge!ggeDataHoje = Date
gge!ggeDataIni = dtInicio
gge!chNumerica = Format$(dtInicio, "yyyymmdd")
gge!ggeDataFim = dtFim

If optEquipamento = True Then
   gge!Alfa2 = "EQUIPAMENTOS"
   optEquipamento = False
   TipoLocacao = 0
Else
   gge!Alfa2 = "PESSOAL"
   optPessoal = False
   TipoLocacao = 1
End If

gge.Update

db.CommitTrans

Call FechaDB

Set Rel = drMedicaoEqptoPes
sql = "Select gge.ggeDatahoje, gge.ggeDataIni, gge.ggeDataFim, gge.Alfa2, gge.chNumerica, Unid.AbreviaturaUnidadeMedida, "
sql = sql & " neg.chPessoa, neg.chUnidadeOperacional, neg.chNumPedido, neg.negContrato, neg.chNumPedidoComp, pes.pesRazaoSocial, "
sql = sql & " det.chDataInicio, det.chDataFim, det.chProduto, det.pedValorDaOperacao, det.pedQuantidadePedida, "
sql = sql & " det.pedPrecoUnidadePedida, det.pedValorDaDiaria, det.pedQtdDias, det.pedValorDaOperacao, det.pedAtividade, prd.prdDescCompleta "
sql = sql & " From geradorgeral gge, unidadedemedida Unid, negociacao neg, detalhenegociacao det, pessoa pes, produto prd "
sql = sql & " WHERE neg.negStatus = ('" & 0 & "') and neg.chNumpedido = ('" & txtMedicao & "') and gge.chAlfaNumerica = ('" & Relatorio & "') and prd.prdOrdemApresentacao = ('" & TipoLocacao & "') and det.chProduto = prd.chProduto "
sql = sql & " and det.chNumpedido = neg.chNumpedido and det.chNumpedidoComp = neg.chNumPedidoComp "
sql = sql & " and neg.chPessoa = pes.chPessoa and det.chProduto = prd.chProduto and Unid.chUnidadeDeMedida = det.pedUnidade "
sql = sql & " order by neg.chUnidadeOperacional, det.chDataInicio, prd.chProduto"
AbrirRelatorio sql, Rel

End Sub

Private Sub cmdGerarMedicao_Click()
If dtInicio = dtFim Then
   MsgBox ("Ajustar o período de Medição"), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If
Call Rotina_AbrirBanco
Relatorio = "drMedicao"

db.BeginTrans

gge.Open "Select * from geradorgeral where chAlfaNumerica = ('" & Relatorio & "')", db, 3, 3
If gge.EOF Then
   gge.AddNew
End If

gge!chAlfaNumerica = "drMedicao"
gge!ggeDataHoje = Date
gge!ggeDataIni = dtInicio
gge!chNumerica = Format$(dtInicio, "yyyymmdd")
gge!ggeDataFim = dtFim
gge!Alfa2 = txtUnidadeOperacional
gge.Update

db.CommitTrans

Call FechaDB

Set Rel = drMedicao
sql = "Select gge.ggeDatahoje, gge.ggeDataIni, gge.ggeDataFim, gge.Alfa2, gge.chNumerica, Unid.AbreviaturaUnidadeMedida, "
sql = sql & " neg.chPessoa, neg.chUnidadeOperacional, neg.chNumPedido, neg.chNumPedidoComp, pes.pesRazaoSocial, "
sql = sql & " det.chDataInicio, det.chDataFim, det.chProduto, det.pedValorDaOperacao, det.pedQuantidadePedida, "
sql = sql & " det.pedPrecoUnidadePedida, det.pedValorDaDiaria, det.pedQtdDias, det.pedValorDaOperacao, det.pedAtividade, prd.prdDescCompleta "
sql = sql & " From geradorgeral gge, unidadedemedida Unid, Negociacao neg, detalhenegociacao det, pessoa pes, produto prd "
sql = sql & " WHERE neg.negStatus = ('" & 0 & "') and neg.chNumpedido = ('" & txtMedicao & "') and neg.chNumpedidoComp = ('" & txtComplemento & "') and gge.chAlfaNumerica = ('" & Relatorio & "') and det.chProduto = prd.chProduto "
sql = sql & " and det.chNumpedido = neg.chNumpedido and det.chNumpedidoComp = neg.chNumPedidoComp "
sql = sql & " and neg.chPessoa = pes.chPessoa and Unid.chUnidadeDeMedida = det.pedUnidade "
sql = sql & " order by det.chDataInicio, prd.chProduto"
AbrirRelatorio sql, Rel
End Sub


Private Sub cmdSair_Click()

Unload Me

End Sub

Private Sub Command1_Click()

Call Rotina_AbrirBanco

If txtMedicao = Empty Then
   MsgBox ("Solicitação de envio de medição Inválida"), vbInformation
   Exit Sub
End If

Resp = MsgBox(("Envio de Medição - ") & txtMedicao & (" para aprovação. Confirma???"), vbYesNo)

If Resp = vbYes Then
   neg.Open "Select * from negociacao where chNumPedido = ('" & txtMedicao & "')", db, 3, 3
   If neg.EOF Then
      MsgBox ("Medição inexistente. Comunicar ao analista reponsável."), vbCritical
      Call FechaDB
      Exit Sub
   End If
Else
   Call FechaDB
   Exit Sub
End If

neg.MoveFirst

Do While Not neg.EOF
   neg!negStatus = 2
   neg!negDataEnvioAprovMedicao = Date
   neg.Update
   neg.MoveNext
Loop

Call FechaDB

Call CarregaGridMedicao

End Sub

Private Sub Form_Load()
Dim Ind As Integer

dtInicio = Date
dtFim = Date
Relatorio = "drMedicao"
dtHoje = Date
gridMedicao.Rows = 2
gridMedicao.TextMatrix(1, 0) = Empty
gridMedicao.TextMatrix(1, 1) = Empty
gridMedicao.TextMatrix(1, 2) = Empty
gridMedicao.TextMatrix(1, 3) = Empty

optEquipamento = False
optPessoal = False

Call CarregaGridMedicao

End Sub

Private Sub GridMedicao_Click()

Dim Limite As Integer
Dim IndLinha As Integer

Limite = gridMedicao.Rows

IndLinha = gridMedicao.Row

If gridMedicao.TextMatrix(IndLinha, 0) = "" Then
   MsgBox "Clicar em linha com conteúdo."
   Exit Sub
End If

txtPessoa = gridMedicao.TextMatrix(IndLinha, 2)

Call Rotina_AbrirBanco

pes.Open "Select * from pessoa where chPessoa = ('" & gridMedicao.TextMatrix(IndLinha, 2) & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Cliente não encontrado. Comuniicar ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
End If


txtRazaoSocial = pes!pesRazaoSocial
txtUnidadeOperacional = gridMedicao.TextMatrix(IndLinha, 3)
txtNome = txtPessoa
txtPedidoComp = gridMedicao.TextMatrix(IndLinha, 1)
txtComplemento = gridMedicao.TextMatrix(IndLinha, 1)
txtNumPedido = gridMedicao.TextMatrix(IndLinha, 0)
txtMedicao = gridMedicao.TextMatrix(IndLinha, 0)

neg.Open "Select * from negociacao where chNumPedido = ('" & txtNumPedido & "') AND chNumPedidoComp = ('" & txtPedidoComp & "')", db, 3, 3
If neg.EOF Then
   MsgBox ("Negociação não encontrada. Comuniicar ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
End If

dtInicio = neg!negInicioMedicao
dtFim = neg!negFinalMedicao

Dia = Day(Date)
mes = Month(Date)
ano = Year(Date)

Call FechaDB

End Sub

Public Sub CarregaGridMedicao()

Call Rotina_AbrirBanco

PedidoAnterior = Empty

gridMedicao.Rows = 2
gridMedicao.TextMatrix(1, 0) = Empty
gridMedicao.TextMatrix(1, 1) = Empty
gridMedicao.TextMatrix(1, 2) = Empty
gridMedicao.TextMatrix(1, 3) = Empty

neg.Open "Select * from negociacao where negStatus = ('" & 0 & "')", db, 3, 3
If neg.EOF Then
   MsgBox ("Não há Medição até a presente data"), vbInformation
   Call FechaDB
   Exit Sub
End If
TipoProduto = 0
Ind = 0
neg.MoveFirst
Do While Not neg.EOF
   If Not neg!chNumPedido = PedidoAnterior Then
      Ind = Ind + 1
      gridMedicao.Rows = Ind + 1
      gridMedicao.TextMatrix(Ind, 0) = neg!chNumPedido
      gridMedicao.TextMatrix(Ind, 1) = neg!chNumPedidoComp
      gridMedicao.TextMatrix(Ind, 2) = neg!chPessoa
      gridMedicao.TextMatrix(Ind, 3) = neg!chUnidadeOperacional
      PedidoAnterior = neg!chNumPedido
   End If
   neg.MoveNext
Loop

Call FechaDB
End Sub
