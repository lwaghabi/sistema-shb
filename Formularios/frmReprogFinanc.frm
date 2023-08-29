VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReprogFinanc 
   Caption         =   "(frmReprogFinanc)"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   LinkTopic       =   "Form4"
   ScaleHeight     =   7875
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   10455
      Begin MSMask.MaskEdBox txtHoje 
         Height          =   375
         Left            =   8760
         TabIndex        =   23
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtNovoValor 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7920
         TabIndex        =   6
         Top             =   3600
         Width           =   1935
      End
      Begin VB.TextBox txtIntervalo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9240
         TabIndex        =   4
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox cmbNVezes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7920
         TabIndex        =   3
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Posição Após Reprogramação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   17
         Top             =   4680
         Width           =   7215
         Begin MSFlexGridLib.MSFlexGrid GridPos 
            Height          =   1695
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   2990
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            BackColorBkg    =   16777152
            FormatString    =   "Nota Fiscal    |Fatura         |Data Vencito |Valo           "
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
      End
      Begin VB.Frame Frame2 
         Caption         =   "Posição Atual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   7215
         Begin MSFlexGridLib.MSFlexGrid GridAtual 
            Height          =   1815
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   3201
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            BackColorBkg    =   16777152
            FormatString    =   "Nota Fiscal    |Fatura         |Data Vencito |Valo           "
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
      End
      Begin VB.Frame Frame4 
         Height          =   2775
         Left            =   7560
         TabIndex        =   15
         Top             =   4080
         Width           =   2775
         Begin VB.CommandButton cmdNovaReprogramacao 
            BackColor       =   &H0000C000&
            Caption         =   "Nova Reprogramação"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1560
            Width           =   2535
         End
         Begin VB.CommandButton cmdSai 
            BackColor       =   &H00FFFF00&
            Caption         =   "Sair"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2280
            Width           =   2535
         End
         Begin VB.CommandButton cmdConfirmar 
            BackColor       =   &H008080FF&
            Caption         =   "Confirmar Reprogramação"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton cmdCancelar 
            BackColor       =   &H0000C0C0&
            Caption         =   "Cancelar Reprogramação"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            MaskColor       =   &H0000C0C0&
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.ComboBox cmbCliFornec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4440
         TabIndex        =   1
         Text            =   "cmbCliFornec"
         Top             =   720
         Width           =   2895
      End
      Begin VB.ComboBox cmbTipoConta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker txtNovoVencimento 
         Height          =   375
         Left            =   8040
         TabIndex        =   5
         Top             =   2760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   240975873
         CurrentDate     =   38135
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Atenção: Clicar na Linha da N.F. que se deseja Reprogramar"
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
         Left            =   240
         TabIndex        =   25
         Top             =   1560
         Width           =   7215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Hoje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9120
         TabIndex        =   24
         Top             =   120
         Width           =   585
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Val. Reprogramado"
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
         Left            =   7560
         TabIndex        =   22
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "A Partir De"
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
         Left            =   8160
         TabIndex        =   21
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Intervalo"
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
         Left            =   9120
         TabIndex        =   20
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "N Vezes"
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
         Left            =   7920
         TabIndex        =   19
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Novo Parcelamento"
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
         Left            =   7920
         TabIndex        =   18
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente/Fornecedor"
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
         Left            =   4440
         TabIndex        =   14
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reprogramação a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2205
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reprogramação Financeira"
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
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmReprogFinanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim fim As Byte
Dim ClienteAnterior As String
Dim Contador As Integer
Dim NotaFiscalAnterior As String
Dim Resp As String
Dim IndLinha As Integer
Dim DataUtil As Date
Dim DataInvertida As Double
Dim Cliente As String
Dim NotaFiscal As String
Dim Fatura As String
Dim DataVencito As Date
Dim ValorAtual As Currency
Dim ValorParcela As Currency
Dim ValorCorrecaoParc As Currency

Dim DataEmissao As Date
Dim datavencoriginal As Date
Dim primeirovencito As Date
Dim descoperacao As String
Dim ValorAnterior As Currency
Dim TipoLancDespAnterior As String
Dim ano As String
Dim Mes As String
Dim Dia  As String
Dim NumPedido As String
Dim NumPedidoComp As String
Dim SemDesd As Byte
Dim banco As String
Dim DiaUtil As Date


Private Sub cmbTipoConta_lostfocus()

If cmbTipoConta = Empty Then
   cmdSai.SetFocus
   Exit Sub
End If

cmbCliFornec.Clear

If cmbTipoConta.ListIndex = 0 Then
   Call Rotina_010_Carga_Credito
Else
   Call Rotina_060_Carga_Debito
End If

End Sub
Private Sub cmbCliFornec_lostfocus()

If cmbCliFornec = Empty Then
   cmdSai.SetFocus
   Exit Sub
End If

GridAtual.Rows = 2
GridAtual.TextMatrix(1, 0) = Empty
GridAtual.TextMatrix(1, 1) = Empty
GridAtual.TextMatrix(1, 2) = Empty
GridAtual.TextMatrix(1, 3) = Empty


If cmbTipoConta.ListIndex = 0 Then
   Rotina_020_NotaFiscal_Credito
Else
   Rotina_070_NotaFiscal_Debito
End If

End Sub


Private Sub cmdCancelar_Click()

'cmbTipoConta = Empty
cmbCliFornec.Clear
cmbNVezes = Empty
txtNovoValor = Format$(0, "#0.00")
txtNovoVencimento = Date
txtIntervalo = Empty
cmbTipoConta.SetFocus
NotaFiscal = Empty

GridAtual.Rows = 2
GridAtual.TextMatrix(1, 0) = Empty
GridAtual.TextMatrix(1, 1) = Empty
GridAtual.TextMatrix(1, 2) = Empty
GridAtual.TextMatrix(1, 3) = Empty
GridPos.Rows = 2
GridPos.TextMatrix(1, 0) = Empty
GridPos.TextMatrix(1, 1) = Empty
GridPos.TextMatrix(1, 2) = Empty
GridPos.TextMatrix(1, 3) = Empty
End Sub

Private Sub cmdConfirmar_Click()
If NotaFiscal = Empty Then
   MsgBox ("Clicar na linha da Nota Fiscal que deseja reprogramar.")
   cmbNVezes.SetFocus
   Exit Sub
End If

If cmbNVezes = 0 Then
   MsgBox ("Não informado o numero de vezes.")
   cmbNVezes.SetFocus
   Exit Sub
End If

If cmbNVezes > 1 Then
   If txtIntervalo = 0 Then
      MsgBox ("Intervalo de dias entre as parcelas não informado")
      txtIntervalo.SetFocus
      Exit Sub
   End If
End If

If txtNovoVencimento < DataVencito Then
   Resp = MsgBox("Novo Vencimento menor do que o atual. Confirma???", vbYesNo)
   If Resp = vbNo Then
      txtNovoVencimento.SetFocus
      Exit Sub
   End If
End If

If txtNovoValor = 0 Then
   Resp = MsgBox("Não foi informado o novo valor. Manter o valor anterior???", vbYesNo)
   If Resp = vbYes Then
      txtNovoValor = Format$(ctr!ctrValorDaBoleta, "#0.00")
   Else
      txtNovoValor.SetFocus
      Exit Sub
   End If
End If

ValorParcela = txtNovoValor / cmbNVezes

If cmbTipoConta.ListIndex = 0 Then
   Call Rotina_030_Processa_Credito
Else
   Call Rotina_080_Processa_Debito
End If

End Sub

Private Sub cmdNovaReprogramacao_Click()

'cmbTipoConta = Empty
cmbCliFornec.Clear
cmbNVezes = Empty
txtNovoValor = Format$(0, "#0.00")
txtNovoVencimento = Date
txtIntervalo = Empty

GridAtual.Rows = 2
GridAtual.TextMatrix(1, 0) = Empty
GridAtual.TextMatrix(1, 1) = Empty
GridAtual.TextMatrix(1, 2) = Empty
GridAtual.TextMatrix(1, 3) = Empty
GridPos.Rows = 2
GridPos.TextMatrix(1, 0) = Empty
GridPos.TextMatrix(1, 1) = Empty
GridPos.TextMatrix(1, 2) = Empty
GridPos.TextMatrix(1, 3) = Empty

NotaFiscal = Empty
cmdConfirmar.Enabled = True
cmdCancelar.Enabled = True

cmbTipoConta.SetFocus
End Sub

Private Sub cmdSai_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtHoje = Date
txtNovoVencimento = Date
cmbTipoConta.AddItem "Receber"
cmbTipoConta.AddItem "Pagar"

For IndLinha = 1 To 12
    cmbNVezes.AddItem IndLinha
Next

txtIntervalo = Empty
txtNovoValor = Empty

End Sub

Public Sub Rotina_010_Carga_Credito()
fim = 0

Call Rotina_AbrirBanco

ctr.Open "Select * from Contas_A_Receber", db, 3, 3
If ctr.EOF Then
   MsgBox ("Não há Contas a Receber para reprogramação"), vbInformation
   Call FechaDB
   Exit Sub
End If

ctr.MoveFirst
Do While fim = 0
   If Not (ctr!chPessoa = ClienteAnterior) Then
      cmbCliFornec.AddItem ctr!chPessoa
      ClienteAnterior = ctr!chPessoa
   End If
   ctr.MoveNext
   If ctr.EOF Then
      fim = 1
   End If
Loop

Call FechaDB

End Sub

Public Sub Rotina_020_NotaFiscal_Credito()

fim = 0
Contador = 0

Call Rotina_AbrirBanco

ctr.Open "Select * from Contas_A_Receber where chPessoa = ('" & cmbCliFornec & "') and ctrStatus = ('" & 0 & "')", db, 3, 3
If ctr.EOF Then
   MsgBox ("Não há financeiro parra reprogaramação para este cliente"), vbInformation
   Call FechaDB
   Exit Sub
End If

ctr.MoveFirst

Do While fim = 0
   Contador = Contador + 1
   GridAtual.Rows = Contador + 1
   GridAtual.TextMatrix(Contador, 0) = ctr!chNotaFiscal
   GridAtual.TextMatrix(Contador, 1) = ctr!chFatura
   GridAtual.TextMatrix(Contador, 2) = ctr!ctrDataVencito
   GridAtual.TextMatrix(Contador, 3) = Format$(ctr!ctrValorDaBoleta, "#,##0.00")
   NotaFiscalAnterior = ctr!chNotaFiscal

   ctr.MoveNext
   If ctr.EOF Then
      fim = 1
   End If
Loop
If Contador = 0 Then
   MsgBox ("Não há Recebimentos pendentes para Reprogramação")
   cmdSai.SetFocus
End If

Call FechaDB

End Sub

Public Sub Rotina_030_Processa_Credito()
If cmbCliFornec = "" Then
   MsgBox "Informe o Cliente"
   Exit Sub
End If

Call Rotina_AbrirBanco

ctr.Open "Select * from Contas_A_Receber where chPessoa = ('" & cmbCliFornec & "') and chNotaFiscal = ('" & NotaFiscal & "') and chFatura = ('" & Fatura & "')", db, 3, 3
If ctr.EOF Then
   MsgBox ("Erro no acesso a Contas a Receber em Reprogramação Financeira"), vbCritical
   Call FechaDB
   Exit Sub
End If

primeirovencito = txtNovoVencimento
DataEmissao = ctr!ctrDataEmissao
datavencoriginal = ctr!ctrDataVencitoOriginal
descoperacao = ctr!ctrDescricaoOperacao
ValorAnterior = ctr!ctrValorDaBoleta
ano = ctr!chAno
Mes = ctr!chMes
Dia = ctr!chDia
NumPedido = ctr!chNumPedido
NumPedidoComp = ctr!chNumPedidoComp
banco = ctr!chCodBcoLart

ValorCorrecaoParc = ((ValorParcela * cmbNVezes) - ValorAnterior) / cmbNVezes

ctr.Delete

db.BeginTrans

For IndLinha = 1 To cmbNVezes

    ctr.AddNew
    ctr!chFabricante = 0
    ctr!chPessoa = cmbCliFornec
    ctr!chNotaFiscal = NotaFiscal
    ctr!chFatura = Fatura & "-" & IndLinha
    ctr!ctrDataEmissao = DataEmissao
    ctr!ctrDataVencito = txtNovoVencimento
    
    'Calcula data banco
    
    DataUtil = txtNovoVencimento
   
    DataInformada = DataUtil
    NDias = 0
    'DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)
    'DataUtil = DataRetorno.DiaUtil
    ctr!ctrDataBanco = DataUtil
    
    'Fim calcula data banco
    
    ctr!ctrDataVencitoOriginal = datavencoriginal
    ctr!ctrDescricaoOperacao = "RF.NF-" & NotaFiscal & "-" & Fatura & "(" & IndLinha & "/" & cmbNVezes & ")"
    ctr!ctrValorMerco = 0
    ctr!ctrPercentCorrecao = 0
    'If ValorAnterior = ValorParcela Then
    '   ctr!ctrvalorlart = TabCtaReceberNF("ctrValorLart")
    '   ctr!ctrPercentCorrecao = TabCtaReceberNF("ctrPercentCorrecao")
    '   ctr!ctrValorCorrecao = TabCtaReceberNF("ctrValorCorrecao")
    '   ctr!ctrPercentlogistica = TabCtaReceberNF("ctrPercentlogistica")
    '   ctr!ctrValorlogistica = TabCtaReceberNF("ctrValorlogistica")
    '   ctr!ctrvalordaboleta = ValorParcela
    'Else
       ctr!ctrvalorcorrecao = ValorCorrecaoParc
       ctr!ctrValorLart = ValorParcela - ValorCorrecaoParc
       ctr!ctrValorDaBoleta = ValorParcela
    'End If
    ctr!chAno = ano
    ctr!chMes = Mes
    ctr!chDia = Dia
    ctr!chNumPedido = NumPedido
    ctr!chNumPedidoComp = NumPedidoComp
    ctr!chCodBcoLart = banco
    ctr!ctrStatus = 0
    txtNovoVencimento = txtNovoVencimento + txtIntervalo
'    If ValorAnterior = ValorParcela Then
'       ctr.Delete
'    End If
    ctr.Update
Next

neg.Open "SELECT * FROM Negociacao WHERE chNumPedido = ('" & NumPedido & "') AND chNumPedidoComp = ('" & NumPedidoComp & "')", db, 3, 3
If neg.EOF Then
   MsgBox ("Erro no acesso a Negociação na rotina de Reprogramação Financeira."), vbCritical
   Call FechaDB
   Exit Sub
End If

neg!negFaturamento = cmbNVezes
neg!negIntervaloFatura = txtIntervalo
neg!negAPartirDe = primeirovencito - DataEmissao
neg!negUltimaAtualizacao = Date
neg!negCntrlFaturamento = 0
neg.Update

db.CommitTrans

fim = 0
Contador = 0

If ctr.State = 1 Then
   ctr.Close: Set ctr = Nothing
End If

ctr.Open "Select * from Contas_A_Receber where chPessoa = ('" & cmbCliFornec & "') and chNotaFiscal = ('" & NotaFiscal & "')", db, 3, 3
If ctr.EOF Then
   MsgBox ("Erro no acesso a Contas a Receber em Reprogramação Financeira grid da reprogramação"), vbCritical
   Call FechaDB
   Exit Sub
End If

ctr.MoveFirst

Do While fim = 0
   Contador = Contador + 1
   GridPos.Rows = Contador + 1
   GridPos.TextMatrix(Contador, 0) = ctr!chNotaFiscal
   GridPos.TextMatrix(Contador, 1) = ctr!chFatura
   GridPos.TextMatrix(Contador, 2) = ctr!ctrDataVencito
   GridPos.TextMatrix(Contador, 3) = Format$(ctr!ctrValorDaBoleta, "#,##0.00")
   NotaFiscalAnterior = ctr!chNotaFiscal

   ctr.MoveNext
   If ctr.EOF Then
      fim = 1
   End If
Loop

NotaFiscal = Empty

Call FechaDB

cmdCancelar.Enabled = False
cmdConfirmar.Enabled = False
End Sub
Public Sub Rotina_060_Carga_Debito()
fim = 0

Call Rotina_AbrirBanco

ctp.Open "Select * from Contas_A_Pagar", db, 3, 3
If ctp.EOF Then
   MsgBox ("Não há contas a pagar para reprogramação."), vbInformation
   Call FechaDB
   Exit Sub
End If

ClienteAnterior = Empty

ctp.MoveFirst

Do While fim = 0
   If ctp!ctpStatus = 0 Then
      If Not ctp!chPessoa = ClienteAnterior Then
         cmbCliFornec.AddItem ctp!chPessoa
         ClienteAnterior = ctp!chPessoa
      End If
   End If
   ctp.MoveNext
   If ctp.EOF Then
      fim = 1
   End If
Loop

Call FechaDB

End Sub

Public Sub Rotina_070_NotaFiscal_Debito()

fim = 0
Contador = 0
NotaFiscalAnterior = Empty

Call Rotina_AbrirBanco

ctp.Open "Select * from Contas_A_Pagar", db, 3, 3
If ctp.EOF Then
   MsgBox ("Não há Nota fiscal para reprogramação"), vbInformation
   Call FechaDB
   Exit Sub
End If

ctp.MoveFirst
Do While fim = 0
   If ctp!chPessoa = cmbCliFornec And ctp!ctpStatus = 0 Then
      Contador = Contador + 1
      GridAtual.Rows = Contador + 1
      GridAtual.TextMatrix(Contador, 0) = ctp!chNotaFiscal
      GridAtual.TextMatrix(Contador, 1) = ctp!chFatura
      GridAtual.TextMatrix(Contador, 2) = ctp!chDataVencito
      GridAtual.TextMatrix(Contador, 3) = Format$(ctp!ctpValorDaBoleta, "#,##0.00")
      NotaFiscalAnterior = ctp!chNotaFiscal
   End If
   ctp.MoveNext
   If ctp.EOF Then
      fim = 1
   End If
Loop
If Contador = 0 Then
   MsgBox ("Não há Pagamentos pendentes para Reprogramação"), vbInformation
   cmdSai.SetFocus
End If

Call FechaDB

End Sub

Public Sub Rotina_080_Processa_Debito()

SemDesd = 0

Call Rotina_AbrirBanco

DataInvertida = Year(GridAtual.TextMatrix(IndLinha, 2)) & Format$(Month(GridAtual.TextMatrix(IndLinha, 2)), "00") & Format$(Day(GridAtual.TextMatrix(IndLinha, 2)), "00")

ctp.Open "Select * from Contas_A_Pagar where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbCliFornec & "') and chNotaFiscal = ('" & NotaFiscal & "') and chFatura = ('" & Fatura & "') and chDataVencito = ('" & DataInvertida & "')", db, 3, 3
If ctp.EOF Then
   MsgBox ("Fatura não encontrada. Verifique o numero e tente denovo"), vbCritical
   Call FechaDB
   cmdSai.SetFocus
   Exit Sub
End If


DataInvertida = Year(GridAtual.TextMatrix(IndLinha, 2)) & Format$(Month(GridAtual.TextMatrix(IndLinha, 2)), "00") & Format$(Day(GridAtual.TextMatrix(IndLinha, 2)), "00")

nfd.Open "Select * from NotaFiscalDesdobramento where chPessoa = ('" & cmbCliFornec & "') and chNotaFiscalEntrada = ('" & NotaFiscal & "') and chDataVencimento = ('" & DataInvertida & "')", db, 3, 3
If nfd.EOF Then
   MsgBox ("Desdobramento da Nota não encontrada"), vbInformation
   SemDesd = 1
End If


primeirovencito = txtNovoVencimento
DataEmissao = ctp!ctpdataemissao
datavencoriginal = ctp!ctpdatavencOriginal
descoperacao = ctp!ctpdescricaooperacao
ValorAnterior = ctp!ctpValorDaBoleta
TipoLancDespAnterior = ctp!ctpTipoLancamentoDesc

ano = ctp!chAno
Mes = ctp!chMes
Dia = ctp!chDia
banco = ctp!chCodBcoLart

db.BeginTrans


ctp.Delete

If SemDesd = 0 Then
   nfd.Delete
End If

For IndLinha = 1 To cmbNVezes
    If ValorParcela > 0 Then
        ctp.AddNew
        ctp!chFabricante = 0
        ctp!chPessoa = cmbCliFornec
        ctp!chNotaFiscal = NotaFiscal
        If IndLinha > 1 Then
           ctp!chFatura = IndLinha
        Else
           ctp!chFatura = Fatura
        End If
        ctp!ctpdataemissao = DataEmissao
        ctp!ctpdatalanc = DataEmissao
        ctp!chDataVencito = txtNovoVencimento
        ctp!ctpdatabanco = txtNovoVencimento
        ctp!ctpValorDaBoleta = txtNovoValor
        ctp!ctpValorLart = txtNovoValor
        ctp!ctpdatavencOriginal = datavencoriginal
        ctp!chCodBcoLart = banco
        ctp!ctpdataproc = Date
        ctp!ctpdescricaooperacao = descoperacao
        ctp!ctpTipoLancamentoDesc = TipoLancDespAnterior
        
        'Calcula data banco
    
        DataUtil = txtNovoVencimento
   
        'DataInformada = DataUtil
        'NDias = 0
        'DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)
        'DataUtil = DataRetorno.DiaUtil
        'TabCtaPagar("ctpDataBanco") = DataUtil
    
        'Fim calcula data banco
        
        
        ctp.Update
        
        If SemDesd = 0 Then
            nfd.AddNew
            nfd!chPessoa = cmbCliFornec
            nfd!chNotaFiscalEntrada = NotaFiscal
            nfd!chDataVencimento = txtNovoVencimento
            nfd!nfdDataVencoriginal = datavencoriginal
            nfd!nfdFaturaNumero = "RF.NF-" & NotaFiscal & "-" & IndLinha & "/" & cmbNVezes
            nfd!nfdValorDaFatura = ValorParcela
            nfd!nfdStatus = 0
            nfd!nfdstatuspagto = 0
            nfd.Update
        End If
        txtNovoVencimento = txtNovoVencimento + txtIntervalo
   End If
Next

db.CommitTrans
txtNovoVencimento = primeirovencito
fim = 0
Contador = 0

If ctp.State = 1 Then
   ctp.Close: Set ctp = Nothing
End If

ctp.Open "Select * from Contas_A_Pagar", db, 3, 3
If ctp.EOF Then
   MsgBox ("Erro no acesso a Contas a Pagar em Reprogramação Financeira."), vbCritical
   Call FechaDB
   Exit Sub
End If



ctp.MoveFirst
Do While fim = 0
   If ctp!chPessoa = cmbCliFornec And ctp!chNotaFiscal = NotaFiscal Then
      Contador = Contador + 1
      GridPos.Rows = Contador + 1
      GridPos.TextMatrix(Contador, 0) = ctp!chNotaFiscal
      GridPos.TextMatrix(Contador, 1) = ctp!chFatura
      GridPos.TextMatrix(Contador, 2) = ctp!chDataVencito
      GridPos.TextMatrix(Contador, 3) = Format$(ctp!ctpValorDaBoleta, "#,##0.00")
      NotaFiscalAnterior = ctp!chNotaFiscal
   End If
   ctp.MoveNext
   If ctp.EOF Then
      fim = 1
   End If
Loop

NotaFiscal = Empty

Call FechaDB

'cmdCancelar.Enabled = False
'cmdConfirmar.Enabled = False
End Sub
Private Sub GridAtual_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
IndLinha = GridAtual.Row
Resp = MsgBox("Deseja Reprogramar a Nota Fiscal - " & GridAtual.TextMatrix(IndLinha, 0) & " Fatura - " & GridAtual.TextMatrix(IndLinha, 1), vbYesNo)
If Resp = vbYes Then
   Cliente = cmbCliFornec
   NotaFiscal = GridAtual.TextMatrix(IndLinha, 0)
   Fatura = GridAtual.TextMatrix(IndLinha, 1)
   DataVencito = GridAtual.TextMatrix(IndLinha, 2)
   ValorAtual = GridAtual.TextMatrix(IndLinha, 3)
   txtNovoValor = Format$(0, "0.00")
   txtNovoVencimento = Date
   txtIntervalo = 0
   cmbNVezes = 1
   cmbNVezes.SetFocus
Else
   MsgBox ("Caso queira reprogramar, clicar com o mouse sobre a linha da nota fiscal desejada.")
   cmdSai.SetFocus
End If
End Sub


