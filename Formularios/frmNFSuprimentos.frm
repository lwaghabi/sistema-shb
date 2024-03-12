VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNFSuprimentos 
   Caption         =   "frmNFSuprimentos"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   855
      Left            =   8880
      TabIndex        =   21
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdGeraFinanceiro 
      Caption         =   "Gera Financeiro"
      Height          =   975
      Left            =   8880
      TabIndex        =   20
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   735
      Left            =   7320
      TabIndex        =   19
      Top             =   5760
      Width           =   1095
   End
   Begin VB.ComboBox cmbNotaFiscal 
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
      Left            =   2280
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Faturamento"
      Height          =   3855
      Left            =   240
      TabIndex        =   15
      Top             =   2880
      Width           =   8535
      Begin VB.CommandButton cmdJogaNaLista 
         Caption         =   "Incluir na Lista"
         Height          =   615
         Left            =   6840
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtQtdFaturas 
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
         Left            =   6960
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtValorFatura 
         Alignment       =   1  'Right Justify
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
         Left            =   4560
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtDataVencitoFatura 
         Height          =   495
         Left            =   2640
         TabIndex        =   5
         Top             =   600
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
         Format          =   391708673
         CurrentDate     =   45149
      End
      Begin VB.CommandButton cmdAlterarNumFaturas 
         Caption         =   "Alterar qtd. faturas"
         Height          =   975
         Left            =   7080
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid tblFaturas 
         Height          =   2535
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         FormatString    =   "Número da Fatura |Data                |Valor                 "
      End
      Begin VB.Label Label8 
         Caption         =   "Qtd Faturas"
         Height          =   255
         Left            =   6840
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Valor da Fatura"
         Height          =   255
         Left            =   4560
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Data"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox txtValorTotalNota 
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
      Left            =   5880
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtNumPO 
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
      Left            =   4080
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
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
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Valor Total da Nota"
      Height          =   615
      Left            =   5880
      TabIndex        =   14
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "PO"
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Nota Fiscal"
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Fornecedor"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Nota Fiscal de Suprimentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   4335
   End
End
Attribute VB_Name = "frmNFSuprimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flagVerifica As Boolean

Private Sub cmbFornecedor_LostFocus()
   Call Rotina_AbrirBanco
   
   cmbNotaFiscal.Clear
   
   rs.Open "SELECT notaFiscal FROM suppedidodecompra INNER JOIN supFaturareceb ON id = idPO WHERE fornecedor = ('" & cmbFornecedor & "')", db, 3, 3
   
   Do While Not rs.EOF
      
      cmbNotaFiscal.AddItem rs!NotaFiscal
      rs.MoveNext
   
   Loop
   
   rs.Close
   
   FechaDB
End Sub
Private Sub cmdSair_Click()
   Unload Me
End Sub
Private Sub Form_Load()
   
   Label8.Visible = False
   txtQtdFaturas.Visible = False
   
   Call Rotina_AbrirBanco
   
   rs.Open "SELECT fornecedor FROM suppedidodecompra INNER JOIN supfaturareceb ON id = idPO", db, 3, 3
   
   If rs.EOF Then
   
      MsgBox ("Não existem fornecedores com faturas pendentes."), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   Do While Not rs.EOF
      
      cmbFornecedor.AddItem rs!fornecedor
      rs.MoveNext
   
   Loop
   
   rs.Close
   
   FechaDB
End Sub
Private Sub tblFaturas_Click()

   If tblFaturas.TextMatrix(tblFaturas.Row, 1) <> Empty And tblFaturas.TextMatrix(tblFaturas.Row, 2) <> Empty Then
   
      dtDataVencitoFatura = tblFaturas.TextMatrix(tblFaturas.Row, 1)
      txtValorFatura = tblFaturas.TextMatrix(tblFaturas.Row, 2)
   
   Else
   
      dtDataVencitoFatura.SetFocus
   
   End If
   
End Sub

Private Sub txtQtdFaturas_LostFocus()
   
   Dim i As Integer
   tblFaturas.Rows = 1

   i = 1
   Do While i <= txtQtdFaturas
      tblFaturas.AddItem cmbNotaFiscal & "-" & i & "/" & txtQtdFaturas
      i = i + 1
   Loop

   txtQtdFaturas.Visible = False
   Label8.Visible = False
   
End Sub

Public Sub verificaValorTotal()
   On Error GoTo Erro
   Dim i As Integer
   Dim total As Currency
   i = 1
   
   Do While i < tblFaturas.Rows
      total = total + tblFaturas.TextMatrix(i, 2)
      i = i + 1
   Loop
   
   If total = txtValorTotalNota Then
   
      flagVerifica = True
   
   Else
   
      flagVerifica = False
   
   End If
Exit Sub
Erro: MsgBox ("Erro ao verificar soma de valores"), vbInformation
End Sub
