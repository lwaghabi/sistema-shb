VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmProdutoAtividadePreco 
   Caption         =   "frmProdutoAtividadePreco"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   12960
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCliente 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3600
      TabIndex        =   7
      Top             =   2040
      Width           =   4455
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   1230
   End
   Begin MSFlexGridLib.MSFlexGrid grdAtividadePreco 
      Height          =   3015
      Left            =   3600
      TabIndex        =   5
      Top             =   2520
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   16777152
      BackColorFixed  =   12648447
      BackColorBkg    =   16777152
      FocusRect       =   2
      FormatString    =   "Atividade                      |Atualização  |Valor                           "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbProduto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox txtHoje 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblLabel5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contrato"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1260
   End
   Begin VB.Label lblLabel3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      TabIndex        =   6
      Top             =   1680
      Width           =   1005
   End
   Begin VB.Label lblLabel2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hoje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta de Atividades e Valores por Contrato"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   7200
   End
End
Attribute VB_Name = "frmProdutoAtividadePreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ProdutoAnterior As String
Dim Ind As Integer
Dim Contrato As String


Private Sub cmbProduto_LostFocus()

If cmbProduto = Empty Then
   cmdSair.SetFocus
   Exit Sub
End If

Call Rotina_AbrirBanco

ProdPco.Open "Select * from produtopreco where chProduto = ('" & cmbProduto & "') and pdpStatus = ('" & 0 & "')", db, 3, 3

If ProdPco.EOF Then
   MsgBox ("Atividade não encontrada para o Contrato."), vbInformation
   Call FechaDB
   grdAtividadePreco.Rows = 1
   txtCliente = Empty
   Exit Sub
End If

'If Prod.State = 1 Then
'   Prod.Close: Set Prod = Nothing
'End If
Prod.Open "Select * from produto where chProduto = ('" & ProdPco!chProduto & "') and prdLocadora = ('" & ProdPco!chPessoa & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Cliente não encontrado em Produto. Comunicar ao analista responsável."), vbCritical
   Call FechaDB
   
   Exit Sub
End If
   
txtCliente = Prod!prdLocadora
'txtUnidadeOperacional = Prod!prdUnidadeOperacional
   

grdAtividadePreco.Rows = 2

Ind = 1

Do While Not ProdPco.EOF
   grdAtividadePreco.Rows = Ind + 1
   grdAtividadePreco.TextMatrix(Ind, 0) = ProdPco!chAtividade
   grdAtividadePreco.TextMatrix(Ind, 1) = ProdPco!pdpDataInicio
   grdAtividadePreco.TextMatrix(Ind, 2) = Format$(ProdPco!pdpPrecoDoProduto, "#,##0.00")
   Ind = Ind + 1
   ProdPco.MoveNext
Loop

   
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtHoje = Date

Contrato = "CONTRATO"

ProdutoAnterior = Empty

Call Rotina_AbrirBanco

Prod.Open "Select * from produto where prdUnidadeOperacional  = ('" & Contrato & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Erro na carga de Produtos. Comunicar ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
End If


Prod.MoveFirst

Do While Not Prod.EOF
   If Not Prod!chProduto = ProdutoAnterior Then
      cmbProduto.AddItem Prod!chProduto
      ProdutoAnterior = Prod!chProduto
   End If
   Prod.MoveNext
Loop
   
End Sub

