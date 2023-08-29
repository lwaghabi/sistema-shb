VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGeraCredito 
   Caption         =   "(frmGeraCredito)    Lançamento de Creditos Financeiros"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   1560
      TabIndex        =   10
      Top             =   0
      Width           =   3615
      Begin VB.ComboBox cmbBanco 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H0000FFFF&
         Caption         =   "Sair"
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdSalvar 
         BackColor       =   &H0000FF00&
         Caption         =   "Salvar"
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdExcluir 
         BackColor       =   &H000000FF&
         Caption         =   "Excluir"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4920
         Width           =   975
      End
      Begin VB.ComboBox cmbLancamento 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   1985
      End
      Begin VB.TextBox txtDescricaoCredito 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtValorCredito 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   4320
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtDataCredito 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   118292481
         CurrentDate     =   39307
      End
      Begin VB.ComboBox cmbTipoCredito 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   2175
      End
      Begin VB.ComboBox cmbColaborador 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label label01 
         Caption         =   "Banco Credito"
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
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Produto/Receita"
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
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Descrição do Crédito"
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
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Valor a Creditar"
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
         Left            =   240
         TabIndex        =   14
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data do Crédito"
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
         TabIndex        =   13
         Top             =   2640
         Width           =   1350
      End
      Begin VB.Label Label2 
         Caption         =   "Motivo do Crédito"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Colaborador"
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
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmGeraCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Resp As String

Private Sub cmbColaborador_LostFocus()
If cmbColaborador = "" Then
   MsgBox "Informar Colaborador"
   cmdSair.SetFocus
   Exit Sub
End If
cmbLancamento.Clear

Call Rotina_AbrirBanco

ctr.Open "Select * from Contas_A_Receber", db, 3, 3
If ctr.EOF Then
   MsgBox ("Não há Contas a Receber disponível até presente momento."), vbInformation
   Call FechaDB
   Exit Sub
End If

ctr.MoveFirst
  Do While Not ctr.EOF
     If ctr!chPessoa = cmbColaborador Then
        cmbLancamento.AddItem ctr!chNotaFiscal
     End If
     ctr.MoveNext
  Loop

If cmbLancamento.ListCount = 0 Then
   cmbLancamento = ""
Else
   cmbLancamento.ListIndex = 0
End If

Call FechaDB

End Sub

Private Sub cmbLancamento_LostFocus()
If cmbLancamento = "" Then
   Exit Sub
End If

Call Rotina_AbrirBanco

ctr.Open "Select * from Contas_A_Receber where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbColaborador & "') and chNotaFiscal = ('" & cmbLancamento & "') and chFatura = ('" & "CREDITO" & "')", db, 3, 3
If Not (ctr.EOF) Then
   txtDescricaoCredito = ctr!ctrDescricaoOperacao
   cmbTipoCredito.ListIndex = ctr!chNumPedido
   dtDataCredito = ctr!ctrDataVencito
   txtValorCredito = Format$(ctr!ctrValorDaBoleta, "##,##0.00")
   cmbBanco = ctr!chCodBcoLart
   cmdExcluir.Enabled = True
   cmdSalvar.Enabled = True
Else
   cmdExcluir.Enabled = False
   cmdSalvar.Enabled = True
End If

Call FechaDB

End Sub

Private Sub cmbTipoCredito_LostFocus()
txtDescricaoCredito = cmbTipoCredito & " " & cmbLancamento
dtDataCredito.SetFocus
End Sub

Private Sub cmdExcluir_Click()

Call Rotina_AbrirBanco

ctr.Open "Select * from Contas_A_Receber where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbColaborador & "') and chNotaFiscal = ('" & cmbLancamento & "') and chFatura = ('" & "CREDITO" & "')", db, 3, 3
If Not (ctr.EOF) Then
   Resp = MsgBox("Solicitação de exclusão de Registro. Confirma???", vbYesNo)
   If vbYes Then
      ctr.Delete
      cmdExcluir.Enabled = False
      cmdSalvar.Enabled = False
   End If

End If
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()

Call Rotina_AbrirBanco

ctr.Open "Select * from Contas_A_Receber where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbColaborador & "') and chNotaFiscal = ('" & cmbLancamento & "') and chFatura = ('" & "CREDITO" & "')", db, 3, 3
If ctr.EOF Then
    ctr.AddNew
    ctr!chFabricante = 0
    ctr!chPessoa = cmbColaborador
    ctr!chNotaFiscal = cmbLancamento
    ctr!chFatura = "CREDITO"
    ctr!ctrDescricaoOperacao = "APORTE "  'cmbTipoCredito & " " & cmbLancamento
    ctr!ctrDescricaoOperacao = cmbLancamento & " " & cmbTipoCredito
    ctr!ctrDataEmissao = dtDataCredito
    ctr!ctrDataVencito = dtDataCredito
    ctr!ctrDataBanco = dtDataCredito
    ctr!ctrDataVencitoOriginal = dtDataCredito
    ctr!ctrValorLart = txtValorCredito
    ctr!ctrValorMerco = 0
    ctr!ctrPercentCorrecao = 0
    ctr!ctrvalorcorrecao = 0
    ctr!ctrPercentlogistica = 0
    ctr!ctrValorlogistica = 0
    ctr!ctrValorDaBoleta = txtValorCredito
    ctr!chAno = Year(Date)
    ctr!chMes = Month(Date)
    ctr!chDia = Day(Date)
    ctr!chNumPedido = cmbTipoCredito.ListIndex
    ctr!chNumPedidoComp = 0
    ctr!chCodBcoLart = cmbBanco
    ctr!ctrStatus = 0
    
    ctr.Update
Else
   MsgBox ("Este crédito já foi gerado."), vbInformation
End If

cmdExcluir.Enabled = False
cmdSalvar.Enabled = False

Call FechaDB

End Sub

Private Sub Form_Load()
cmbColaborador.Clear
cmbLancamento.Clear
cmbTipoCredito.Clear
dtDataCredito = Date

Call Rotina_AbrirBanco

pes.Open "Select * from Pessoa", db, 3, 3
If pes.EOF Then
   MsgBox ("Erro. Não há cadastro de Pessoa."), vbCritical
   Call FechaDB
   Exit Sub
End If
   
pes.MoveFirst
Do While Not pes.EOF
   If Not (pes!pestipopessoa = 0) Then
      cmbColaborador.AddItem pes!chPessoa
   End If
   pes.MoveNext
Loop

cmbColaborador.ListIndex = 0


Bco.Open "Select * from Banco", db, 3, 3

Bco.MoveFirst
Do While Not Bco.EOF
   cmbBanco.AddItem Bco!bcosiglabco
   Bco.MoveNext
Loop

cmbBanco.ListIndex = 0

cmbTipoCredito.AddItem "Devolução"
cmbTipoCredito.AddItem "APORTE"
cmbTipoCredito.AddItem "Ressarcimento"
cmbTipoCredito.AddItem "VENDA"

cmbTipoCredito.ListIndex = 0

Call FechaDB

End Sub


