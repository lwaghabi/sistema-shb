VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAjustaCobranca 
   Caption         =   "Ajusta Cobrança"
   ClientHeight    =   4515
   ClientLeft      =   2715
   ClientTop       =   2385
   ClientWidth     =   7350
   LinkTopic       =   "Form3"
   ScaleHeight     =   4515
   ScaleWidth      =   7350
   Begin MSMask.MaskEdBox txtHoje 
      Height          =   495
      Left            =   5280
      TabIndex        =   17
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   7095
      Begin MSComCtl2.DTPicker txtDataReceb 
         Height          =   255
         Left            =   5520
         TabIndex        =   0
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   242745345
         CurrentDate     =   44649
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   29
         Top             =   2640
         Width           =   3975
         Begin VB.CommandButton cmdCancela 
            BackColor       =   &H000000FF&
            Caption         =   "Cancela"
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
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Confirma 
            BackColor       =   &H0000C0C0&
            Caption         =   "Confirma"
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
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSMask.MaskEdBox txtDataVencito 
         Height          =   255
         Left            =   3240
         TabIndex        =   18
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   -2147483644
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbFormaDeCobranca 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtAjuste 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         TabIndex        =   2
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Recebido em"
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
         Left            =   5520
         TabIndex        =   30
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label txtValorFaturaCorrigido 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5520
         TabIndex        =   28
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label txtValorDaCorrecao 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5520
         TabIndex        =   27
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Dias de   Atraso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   26
         Top             =   960
         Width           =   735
      End
      Begin VB.Label txtValorDaFatura 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label txtDiasDeAtraso 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4680
         TabIndex        =   24
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label txtDescOperacao 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label txtFatura 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5640
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.Label txtNotaFiscal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.Label txtCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Operação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nota Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4080
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fatura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5880
         TabIndex        =   12
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Data Vencimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   11
         Top             =   1200
         Width           =   1470
      End
      Begin VB.Label Label7 
         Caption         =   "Forma de Cobrança"
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
         Left            =   1800
         TabIndex        =   10
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Valor da Fatura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   1320
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "% ou Valor dia/Fixo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3720
         TabIndex        =   8
         Top             =   1920
         Width           =   1665
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Valor da Correção"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5520
         TabIndex        =   7
         Top             =   1920
         Width           =   1545
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Valor da Fatura Corrigido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   6
         Top             =   3120
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ajuste de Cobrança"
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
      TabIndex        =   16
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hoje"
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
      Left            =   4320
      TabIndex        =   15
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frmAjustaCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DiasAtraso As Integer
Dim DataVencito As Date
Dim ValorDaFatura As Currency
Dim ValorDaCorrecao As Currency
Dim ValorParcela As Currency
Dim Cliente As String
Dim Pessoa As String
Dim NotaFiscal As String
Dim Fatura As String
Dim DataEmissao As String

Dim DataBanco As String
Dim DataVencitoOriginal As String
Dim DescricaoOperacao As String
Dim ValorLart As Currency
Dim ValorMerco As Currency
Dim PercentCorrecao As Integer
Dim ValorCorrecao As Currency
Dim ValorDaBoleta As Currency
Dim ano As String
Dim Mes As String
Dim Dia As String
Dim NumPedido As String
Dim NumPedidoComp As String
Dim CodBcoLart As String
Dim Status As Integer
Dim ControleParcela As Integer
Dim FormaDeCobranca As Integer



Private Sub cmbFormaDeCobranca_LostFocus()

DataVencito = txtDataVencito

DiasAtraso = Date - DataVencito

txtDiasDeAtraso = DiasAtraso

If cmbFormaDeCobranca.ListIndex = 0 Then
   txtAjuste = txtValorDaFatura
   ValorDaCorrecao = txtValorDaFatura
End If


End Sub

Private Sub cmdCancela_Click()
frmControleFinanceiro.txtStatusCancela = 1
Unload Me
End Sub

Private Sub Confirma_Click()

Call Rotina_AbrirBanco

ctr.Open "Select * from Contas_A_Receber where chFabricante = ('" & 0 & "') and chPessoa = ('" & txtCliente & "') and chNotaFiscal = ('" & txtNotaFiscal & "') and chFatura = ('" & txtFatura & "')", db, 3, 3
If ctr.EOF Then
   MsgBox ("Erro na abertura do contas a receber"), vbCritical
   Call FechaDB
   Exit Sub
End If
   
If txtAjuste = Empty Then
   MsgBox ("Não Informado o Valor do pagamento"), vbInformation
   Call FechaDB
   txtAjuste.SetFocus
   Exit Sub
End If

db.BeginTrans

ctr!ctrvalorcorrecao = ValorDaCorrecao
ctr!ctrDataRecebimento = txtDataReceb
ctr!ctrValorDaBoleta = Format$(txtValorFaturaCorrigido, "#,##0.00")
ctr.Update
db.CommitTrans
frmControleFinanceiro.txtStatusCancela = 0
Unload Me

Call FechaDB

If FormaDeCobranca = 1 Then
   FormaDeCobranca = 0
   If ValorParcela > 0 Then
      Call Rotina_070_Gera_Parcelamento
   End If
End If


End Sub

Private Sub Form_Load()
txtHoje = Date
txtDataReceb = Date

cmbFormaDeCobranca.AddItem "Valor Sem Acréscimo"
cmbFormaDeCobranca.AddItem "Valor Pago"
cmbFormaDeCobranca.AddItem "Dias atraso X Valor"
cmbFormaDeCobranca.AddItem "% sobre valor"
cmbFormaDeCobranca.AddItem "%/Valor X N Dias"
cmbFormaDeCobranca.AddItem "Parcelamento"

cmbFormaDeCobranca = Empty
txtAjuste = Empty
txtValorDaCorrecao = Empty
txtValorFaturaCorrigido = Empty

ValorDaFatura = Format$(0, "#0.00")
ValorDaCorrecao = Format$(0, "#0.00")

End Sub


Private Sub txtAjuste_LostFocus()

If cmbFormaDeCobranca = Empty Then
   MsgBox ("Nao informada a forma de cobrança")
   cmbFormaDeCobranca.SetFocus
   Exit Sub
End If

If txtAjuste = Empty Then
   MsgBox ("Informar o valor do ajuste")
   txtAjuste.SetFocus
   Exit Sub
End If

Select Case cmbFormaDeCobranca.ListIndex
   Case 0
        Call Rotina_05_Valor_Sem_Acerscimo
   Case 1
        Call Rotina_010_Valor_Fixo
   Case 2
        Call Rotina_020_NDias_Valor
   Case 3
        Call Rotina_030_Perc_Valor
   Case 4
        Call Rotina_040_Perc_Valor_NDias
   Case 5
        Call Rotina_060_Parcelamento
        
End Select

'cmdCancela.SetFocus

End Sub

Public Sub Rotina_05_Valor_Sem_Acerscimo()

txtAjuste = txtValorDaFatura
txtValorDaCorrecao = 0
ValorDaFatura = txtAjuste
ValorDaCorrecao = 0
txtValorFaturaCorrigido = Format$(txtAjuste, "#,##0.00")
ValorParcela = txtValorDaFatura - txtValorDaCorrecao
   
End Sub
Public Sub Rotina_010_Valor_Fixo()
txtValorDaCorrecao = Format$((txtAjuste - txtValorDaFatura), "#,##0.00")
Call Rotina_050_Calcula_Nova_Fatura
End Sub

Public Sub Rotina_020_NDias_Valor()
txtValorDaCorrecao = Format$(txtAjuste * txtDiasDeAtraso, "#,##0.00")
Call Rotina_050_Calcula_Nova_Fatura
End Sub

Public Sub Rotina_030_Perc_Valor()
txtValorDaCorrecao = Format$((txtAjuste * txtValorDaFatura) / 100, "#,##0.00")
Call Rotina_050_Calcula_Nova_Fatura
End Sub

Public Sub Rotina_040_Perc_Valor_NDias()
txtValorDaCorrecao = Format$(((txtAjuste * txtValorDaFatura) / 100) * txtDiasDeAtraso, "#,##0.00")
Call Rotina_050_Calcula_Nova_Fatura
End Sub

Public Sub Rotina_050_Calcula_Nova_Fatura()
ValorDaFatura = txtValorDaFatura
ValorDaCorrecao = txtValorDaCorrecao
txtValorFaturaCorrigido = Format$(ValorDaFatura + ValorDaCorrecao, "#,##0.00")
End Sub

Public Sub Rotina_060_Parcelamento()

FormaDeCobranca = 1
Cliente = txtCliente
NotaFiscal = txtNotaFiscal
Fatura = txtFatura

txtValorDaCorrecao = txtAjuste
ValorDaFatura = txtAjuste
ValorDaCorrecao = txtValorDaCorrecao
txtValorFaturaCorrigido = Format$(txtAjuste, "#,##0.00")
ValorParcela = txtValorDaFatura - txtValorDaCorrecao

End Sub

Public Sub Rotina_070_Gera_Parcelamento()

Call Rotina_AbrirBanco

ctr.Open "Select * from Contas_A_Receber where chFabricante = ('" & 0 & "') and chPessoa = ('" & Cliente & "') and chNotaFiscal = ('" & NotaFiscal & "') and chFatura = ('" & Fatura & "')", db, 3, 3
If ctr.EOF Then
   MsgBox ("Erro na abertura do contas a receber"), vbCritical
   Call FechaDB
   Exit Sub
Else
   Call SalvaRegistro
End If

ctr.AddNew

ctr!chPessoa = Pessoa
ctr!chNotafiscal = NotaFiscal
ctr!ctrControleParcela = ControleParcela + 1
ctr!chFatura = "Parc - " & ctr!ctrControleParcela

ctr!ctrDataEmissao = DataEmissao
ctr!ctrDataVencito = DataVencito
ctr!ctrDataBanco = DataBanco
ctr!ctrDataVencitoOriginal = DataVencitoOriginal
ctr!ctrDescricaoOperacao = DescricaoOperacao
ctr!ctrValorLart = ValorParcela
ctr!ctrPercentCorrecao = PercentCorrecao
ctr!ctrvalorcorrecao = 0
ctr!ctrValorMerco = ValorMerco

ctr!chAno = ano
ctr!chMes = Mes
ctr!chDia = Dia
ctr!chNumPedido = NumPedido
ctr!chNumPedidoComp = NumPedidoComp
ctr!chCodBcoLart = CodBcoLart
ctr!ctrStatus = Status
ctr!ctrDataRecebimento = txtDataReceb

ctr!ctrValorDaBoleta = Format$(ValorParcela, "#,##0.00")

ctr.Update

Call FechaDB

End Sub

Public Sub SalvaRegistro()

Pessoa = ctr!chPessoa
NotaFiscal = ctr!chNotafiscal
Fatura = ctr!chFatura
DataEmissao = ctr!ctrDataEmissao
DataVencito = ctr!ctrDataVencito
DataBanco = ctr!ctrDataBanco
DataVencitoOriginal = ctr!ctrDataVencitoOriginal
DescricaoOperacao = ctr!ctrDescricaoOperacao
ValorLart = ctr!ctrValorLart
PercentCorrecao = ctr!ctrPercentCorrecao
ValorCorrecao = ctr!ctrvalorcorrecao
ValorMerco = ctr!ctrValorMerco
ValorDaBoleta = ctr!ctrValorDaBoleta
ano = ctr!chAno
Mes = ctr!chMes
Dia = ctr!chDia
NumPedido = ctr!chNumPedido
NumPedidoComp = ctr!chNumPedidoComp
CodBcoLart = ctr!chCodBcoLart
Status = ctr!ctrStatus

ControleParcela = ctr!ctrControleParcela

End Sub

