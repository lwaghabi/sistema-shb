VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConsultaFinanceiro 
   Caption         =   "Consulta a Informações Financeiras(frmConsultaFinanceiro)"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H008080FF&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdConsultaMesAtu 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consulta no Mes Atual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdConsultaMesAnt 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consulta em Meses Anteriores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin MSMask.MaskEdBox txtHoje 
         Height          =   315
         Left            =   6120
         TabIndex        =   7
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtValorAConsultar 
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
      Begin VB.ComboBox cmbTipoConsulta 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Hoje"
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
         Left            =   6120
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Valor Para Consulta"
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
         Left            =   2760
         TabIndex        =   8
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Consultar por"
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
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmConsultaFinanceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbTipoConsulta_LostFocus()
If cmbTipoConsulta.ListIndex = 2 Then
   txtValorAConsultar.Alignment = 1
Else
   txtValorAConsultar.Alignment = 0
End If
txtValorAConsultar = ""
End Sub

Private Sub cmdConsultaMesAnt_Click()
If cmbTipoConsulta = "" Then
   MsgBox "Tipo de Consulta Não Informado"
   cmdSair.SetFocus
   Exit Sub
End If
If txtValorAConsultar = "" Then
   MsgBox "Valor da Consulta Não Informado"
   cmdSair.SetFocus
   Exit Sub
End If

Sql = "Select ctp.chDataVencito AS Vencimento, ctp.ctpDataPagamento as Pago_Em, ctp.chNotaFiscal as Documento,"
Sql = Sql & " ctp.ctpDescricaoOperacao as Descrição, ctp.chPessoa as Colaborador, ctp.ctpValorDaBoleta as Valor"
Sql = Sql & " from historicoContasPagar ctp where "

If cmbTipoConsulta.ListIndex = 0 Then
   Sql = Sql & " ctp.chpessoa like '%" & txtValorAConsultar & "%'"
Else
   If cmbTipoConsulta.ListIndex = 1 Then
      Sql = Sql & " ctp.chNotaFiscal like '%" & txtValorAConsultar & "%'"
   Else
      If cmbTipoConsulta.ListIndex = 2 Then
         Sql = Sql & " ctp.ctpValorDaBoleta like '%" & txtValorAConsultar & "%'"
      Else
         Sql = Sql & " ctp.ctpDescricaoOperacao like '%" & txtValorAConsultar & "%'"
      End If
   End If
End If
Sql = Sql & " order by ctp.chDataVencito desc"
'MsgBox Sql
deBusFinanceiro.Commands.Item("cmdbusFinanceiro").CommandText = Sql
frmResultPesqFinanc.Show vbModal
deBusFinanceiro.rscmdBusFinanceiro.Close
End Sub

Private Sub cmdConsultaMesAtu_Click()

If cmbTipoConsulta = "" Then
   MsgBox "Tipo de Consulta Não Informado"
   cmdSair.SetFocus
   Exit Sub
End If
If txtValorAConsultar = "" Then
   MsgBox "Valor da Consulta Não Informado"
   cmdSair.SetFocus
   Exit Sub
End If

Sql = "Select ctp.chDataVencito AS Vencimento, ctp.ctpDataPagamento as Pago_Em, ctp.chNotaFiscal as Documento,"
Sql = Sql & " ctp.ctpDescricaoOperacao as Descrição, ctp.chPessoa as Colaborador, ctp.ctpValorDaBoleta as Valor"
Sql = Sql & " from Contas_A_Pagar ctp where "

If cmbTipoConsulta.ListIndex = 0 Then
   Sql = Sql & " ctp.chpessoa like '%" & txtValorAConsultar & "%'"
Else
   If cmbTipoConsulta.ListIndex = 1 Then
      Sql = Sql & " ctp.chNotaFiscal like '%" & txtValorAConsultar & "%'"
   Else
      If cmbTipoConsulta.ListIndex = 2 Then
         Sql = Sql & " ctp.ctpValorDaBoleta like '%" & txtValorAConsultar & "%'"
      Else
         Sql = Sql & " ctp.ctpDescricaoOperacao like '%" & txtValorAConsultar & "%'"
      End If
   End If
End If
Sql = Sql & " order by ctp.chDataVencito"
MsgBox Sql
deBusFinanceiro.Commands.Item("cmdbusFinanceiro").CommandText = Sql
frmResultPesqFinanc.Show vbModal
deBusFinanceiro.rscmdBusFinanceiro.Close
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtHoje = Date
cmbTipoConsulta.AddItem "Colaborador"
cmbTipoConsulta.AddItem "Número do Documento"
cmbTipoConsulta.AddItem "Valor do Documento"
cmbTipoConsulta.AddItem "Descrição da Operação"
cmbTipoConsulta.ListIndex = 0

End Sub


