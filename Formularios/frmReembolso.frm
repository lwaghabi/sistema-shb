VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReembolso 
   Caption         =   "frmReembolso"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10680
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
   ScaleHeight     =   9165
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtValorTotalDaNotaFiscal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8040
      TabIndex        =   23
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Frame Frame 
      Height          =   7095
      Left            =   720
      TabIndex        =   7
      Top             =   720
      Width           =   9375
      Begin MSFlexGridLib.MSFlexGrid gridDetalheNota 
         Height          =   2175
         Left            =   240
         TabIndex        =   27
         Top             =   2280
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         FormatString    =   "Fornec/Despesa |N.Fiscal/Doc|Produto      |Qtd. | P.U      | Valor Total"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdEmissaoRecibo 
         BackColor       =   &H0000FF00&
         Caption         =   "Salvar e Emissão de Recibo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "3750"
         Top             =   6270
         Width           =   5895
      End
      Begin VB.CommandButton cmbExluir 
         BackColor       =   &H000000FF&
         Caption         =   "Excluir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   6240
         Width           =   1455
      End
      Begin VB.ComboBox cmbMeioPagto 
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
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   5565
         Width           =   2655
      End
      Begin VB.ComboBox cmbReembolso 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   5565
         Width           =   2415
      End
      Begin VB.ComboBox cmbColaborador 
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
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtNomePessoa 
         Enabled         =   0   'False
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
         TabIndex        =   12
         Top             =   480
         Width           =   5655
      End
      Begin VB.TextBox txtBanco 
         Enabled         =   0   'False
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
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtAgencia 
         Enabled         =   0   'False
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
         Left            =   1080
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtContaCorrente 
         Enabled         =   0   'False
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
         Left            =   2040
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtNumComprovante 
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
         Left            =   5640
         TabIndex        =   3
         Top             =   5565
         Width           =   1935
      End
      Begin VB.TextBox txtCPF 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         Left            =   3960
         TabIndex        =   8
         Top             =   1440
         Width           =   3135
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H0000FFFF&
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "3750"
         Top             =   6270
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtDataDeposito 
         Height          =   465
         Left            =   7080
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   820
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   242941953
         CurrentDate     =   44370
      End
      Begin VB.Label Label1 
         Caption         =   "Total do Reembolso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   28
         Top             =   4440
         Width           =   2655
      End
      Begin VB.Label lblColaborador 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colaborador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   120
         Width           =   1500
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   21
         Top             =   120
         Width           =   705
      End
      Begin VB.Label lblBco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label lblAg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ag."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label lblContaCorrente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta Corrente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         Top             =   1080
         Width           =   1875
      End
      Begin VB.Label lblDataDo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data do Depósito"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7080
         TabIndex        =   17
         Top             =   1080
         Width           =   2115
      End
      Begin VB.Label lblComprovante 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Comprovante"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5520
         TabIndex        =   16
         Top             =   5280
         Width           =   1710
      End
      Begin VB.Label lblCNPJCPF 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ/CPF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   15
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label lblReembolso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reembolso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   5280
         Width           =   1350
      End
      Begin VB.Label lblMeioDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Meio de Pagto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   13
         Top             =   5280
         Width           =   1755
      End
   End
   Begin VB.TextBox txtHoje 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   29
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lblTotalPara 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total para Reembolso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   25
      Top             =   5280
      Width           =   2700
   End
   Begin VB.Label lblControleE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Controle e Lançamento de Reembolso"
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
      Left            =   840
      TabIndex        =   24
      Top             =   0
      Width           =   5910
   End
End
Attribute VB_Name = "frmReembolso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pessoaAnterior As String
Dim NotaFiscalAnterior As String
Dim TipoPessoa As String
Dim Ind As Integer
Dim AcumulaValor As Currency
Dim Salvar As Integer
Dim Resp As Boolean
Dim Fim As Boolean
Dim Produto As String
Dim IndSalvo As Integer
Dim Inclusao As Boolean
Dim Relatorio As String
Dim Rel As Object
Dim sql As String
Dim ValorPorExtenso As String
Dim ChavePessoa As String
Dim ChaveNotaFiscal As String
Dim AcumulaValorGrid As Currency
Dim ContaReembolso As Integer

Private Sub cmbColaborador_LostFocus()

gridDetalheNota.Rows = 2
gridDetalheNota.TextMatrix(1, 0) = Empty
gridDetalheNota.TextMatrix(1, 1) = Empty
gridDetalheNota.TextMatrix(1, 2) = Empty
gridDetalheNota.TextMatrix(1, 3) = Empty
gridDetalheNota.TextMatrix(1, 4) = Empty
gridDetalheNota.TextMatrix(1, 5) = Empty

'cmbColaborador.ListIndex = 0

If cmbColaborador = "NENHUM" Then
   Call FechaDB
   Exit Sub
   cmdSair.SetFocus
End If

Call Rotina_AbrirBanco

pes.Open "Select * from pessoa where chPessoa = ('" & cmbColaborador & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Erro no acesso ao pessoa. "), vbCritical
   Call FechaDB
   Exit Sub
End If

txtNomePessoa = pes!pesRazaoSocial
txtBanco = pes!pesBanco
txtAgencia = pes!pesAgencia
txtContaCorrente = pes!pesConta
txtCPF = pes!chCNPJ_CPF

ctp.Open "Select * from contas_a_pagar where ctpPessoaReembolso = ('" & cmbColaborador & "') and ctpStatus = ('" & "2" & "')", db, 3, 3
If ctp.EOF Then
   MsgBox ("Colaborador sem reembolso a processar"), vbInformation
   Call FechaDB
   txtNomePessoa = Empty
   txtBanco = Empty
   txtAgencia = Empty
   txtContaCorrente = Empty
   txtCPF = Empty
   Exit Sub
End If

ChavePessoa = ctp!chPessoa
ChaveNotaFiscal = ctp!chNotafiscal

AcumulaValor = 0
Ind = 1

ctp.MoveFirst

Do While Not ctp.EOF

   ChavePessoa = ctp!chPessoa
   ChaveNotaFiscal = ctp!chNotafiscal
   
   If nfd.State = 1 Then
      nfd.Close: Set nfd = Nothing
   End If
     
   nfd.Open "Select * from notafiscaldetprod where chPessoa = ('" & ChavePessoa & "') and chNotaFiscalEntrada = ('" & ChaveNotaFiscal & "')", db, 3, 3
   If nfd.EOF Then
      nfd.Close
      nfd.Open "Select * from historiconotafiscaldetprod where chPessoa = ('" & ChavePessoa & "') and chNotaFiscalEntrada = ('" & ChaveNotaFiscal & "')", db, 3, 3
      If nfd.EOF Then
         MsgBox ("Nota Fiscal não encontrada."), vbCritical
         Call FechaDB
         Exit Sub
      End If
   End If
   
   nfd.MoveFirst
   
   Do While Not nfd.EOF
      
      gridDetalheNota.Rows = Ind + 1
      gridDetalheNota.TextMatrix(Ind, 0) = nfd!chPessoa
      gridDetalheNota.TextMatrix(Ind, 1) = nfd!chNotaFiscalEntrada
      gridDetalheNota.TextMatrix(Ind, 2) = nfd!chCodProduto
      gridDetalheNota.TextMatrix(Ind, 3) = nfd!nfdQtd
      gridDetalheNota.TextMatrix(Ind, 4) = Format$(nfd!nfdPU, "##0.00")
      gridDetalheNota.TextMatrix(Ind, 5) = Format$(nfd!nfdValorParcela, "##0.00")
      
      AcumulaValor = AcumulaValor + nfd!nfdValorParcela
      
      Ind = Ind + 1
      
      nfd.MoveNext
   Loop
   
   ctp.MoveNext
   
Loop

Rmb.Open "Select * from reembolso where chPessoa = ('" & ChavePessoa & "') and chNotaFiscal = ('" & ChaveNotaFiscal & "')", db, 3, 3
If Rmb.EOF Then
   txtNumComprovante = Empty
Else
   If IsNull(Rmb!rmbTiporeembolso) Then
      cmbReembolso.ListIndex = 0
      cmbMeioPagto.ListIndex = 0
      txtNumComprovante = Empty
   Else
      cmbReembolso.ListIndex = Rmb!rmbTiporeembolso
      cmbMeioPagto.ListIndex = Rmb!rmbMeioPagto
      txtNumComprovante = Rmb!rmbNumComprovanteReembolso
   End If
End If

IndSalvo = Ind

txtValorTotalDaNotaFiscal = Format$(AcumulaValor, "##0.00")

cmbReembolso.SetFocus

End Sub

'Private Sub cmbPessoa_LostFocus()

'If cmbPessoa = Empty Or cmbPessoa = "" Then
'   cmdSair.SetFocus
'   Exit Sub
'End If

'cmbNotaFiscal.Clear

'Call Rotina_AbrirBanco

'ctp.Open "Select * from contas_a_pagar where chPessoa = ('" & cmbPessoa & "') and ctpStatus = ('" & 2 & "')", db, 3, 3
'If ctp.EOF Then
'   MsgBox ("Fornecedor/Despesa sem Nota Fiscal para reembolso."), vbInformation
'   Call FechaDB
'   Exit Sub
'End If

'ctp.MoveFirst

'Do While Not ctp.EOF
''   cmbNotaFiscal.AddItem ctp!chNotaFiscal
'   ctp.MoveNext
'Loop

'Call FechaDB

'End Sub

Private Sub cmdEmissaoRecibo_Click()
Call CriticaEntrada

If Salvar = 0 Then
   Exit Sub
End If
Ind = 0
pessoaAnterior = Empty
AcumulaValorGrid = 0

Do While (Ind + 1) < IndSalvo
   
   Ind = Ind + 1
   If gridDetalheNota.TextMatrix(Ind, 0) = Empty Then
      Ind = IndSalvo + 1
   Else
      If Not (gridDetalheNota.TextMatrix(Ind, 0) = pessoaAnterior And NotaFiscalAnterior = gridDetalheNota.TextMatrix(Ind, 1)) Then
         If Not pessoaAnterior = Empty Then
            
            Call GravarReembolso
            
         End If
            
         pessoaAnterior = gridDetalheNota.TextMatrix(Ind, 0)
         NotaFiscalAnterior = gridDetalheNota.TextMatrix(Ind, 1)
         AcumulaValorGrid = gridDetalheNota.TextMatrix(Ind, 5)
      Else
         AcumulaValorGrid = AcumulaValorGrid + gridDetalheNota.TextMatrix(Ind, 5)
      End If
   End If
Loop

'Gravar Último Grid

Call GravarReembolso

Resp = MsgBox("Reembolso Salvo com Sucesso. Deseja Imprimir o Recibo???", vbExclamation + vbYesNo)
If Not Resp = vbYes Then
   
   Call Rotina_AbrirBanco

   Rmb.Open "Select * from reembolso where rmbStatusRecibo = ('" & 0 & "') and rmbColaborador = ('" & cmbColaborador & "')", db, 3, 3
   If Rmb.EOF Then
      MsgBox ("Erro no acessoa reembolso na finalização da emissão de Recibo"), vbCritical
      Call FechaDB
      Exit Sub
   End If
   
   
   Rmb.MoveFirst
   
   Do While Not Rmb.EOF
      Rmb!rmbStatusReembolso = 1
      Rmb.Update
      Rmb.MoveNext
   Loop
   
   ctp.Open "Select * from contas_a_pagar where chPessoa = ('" & pessoaAnterior & "') and chNotaFiscal = ('" & NotaFiscalAnterior & "')", db, 3, 3
     If ctp.EOF Then
        MsgBox ("Contas a Pagar não encontrada. Erro. Comunicar ao analista responsável"), vbCritical
        Call FechaDB
        Exit Sub
     End If
     
     ctp!ctpDataPagamento = Date
     ctp!ctpStatus = 1
     ctp.Update

Else

   AcumulaValor = 0
   
   If txtNumComprovante = "" Then
      MsgBox ("Número do Comprovante de Depósito não informado."), vbInformation
      Exit Sub
   End If
   
   If cmbReembolso = Empty Then
      MsgBox ("Tipo de reembolso não informado"), vbInformation
      Call FechaDB
      Exit Sub
   End If
   
   If cmbMeioPagto = Empty Then
      MsgBox ("Forma de reembolso não informado"), vbInformation
      Call FechaDB
      Exit Sub
   End If
   
   Call Rotina_AbrirBanco
   
   Rmb.Open "Select * from reembolso where rmbStatusRecibo = ('" & 0 & "') and rmbColaborador = ('" & cmbColaborador & "')", db, 3, 3
   If Rmb.EOF Then
      MsgBox ("Erro no acessoa reembolso na finalização da emissão de Recibo"), vbCritical
      Call FechaDB
      Exit Sub
   End If
   
   
   Rmb.MoveFirst
   
   Do While Not Rmb.EOF
      Rmb!rmbTiporeembolso = cmbReembolso.ListIndex
      Rmb!rmbTipoReembolsoTexto = cmbReembolso
      Rmb!rmbMeioPagto = cmbMeioPagto.ListIndex
      Rmb!rmbMeioPagtoTexto = cmbMeioPagto
      Rmb!rmbNumComprovanteReembolso = txtNumComprovante
      Rmb!rmbStatusReembolso = 1
      AcumulaValor = AcumulaValor + Rmb!rmbValorReembolso
      Rmb.Update
      Rmb.MoveNext
   Loop
      
   'Rotina de chamada Extenso
   
   Call Extenso
   '
   'ValorPorExtenso = ValorExtenso(AcumulaValor)
   
   'Call ValorExtenso(AcumulaValor)
   
   'MsgBox ("Valor por extenso = ") & ValorPorExtenso
   
   
   'Rotina para limpar Extenso
   
   '    Maskvalor.Text = ""
   '    Text2.Text = ""
   '    Maskvalor.SetFocus
   '
   
   Call Rotina_AbrirBanco
   
   Relatorio = "drReciboReembolso"
   
   gge.Open "Select * from geradorgeral where chAlfaNumerica = ('" & Relatorio & "')", db, 3, 3
   If gge.EOF Then
      gge.AddNew
   End If
   
   gge!chAlfaNumerica = "drReciboReembolso"
   gge!ggeDataHoje = Date
   gge!ggeDataIni = dtDataDeposito
   gge!Alfa2 = txtNomePessoa
   gge!Alfa3 = cmbColaborador
   gge!Alfa4 = "(" & ValorPorExtenso & ")"
   gge!num2 = Format$(AcumulaValor, "##0,##0.00")
   gge!Num3 = txtNumComprovante
   gge.Update
   
   Set Rel = drReciboReembolso
      sql = "Select gge.ggeDataHoje, gge.ggeDataIni, gge.Alfa3, gge.chAlfaNumerica, rmb.rmbDataNotaFiscal, rmb.chNotaFiscal,"
      sql = sql & " rmb.rmbBanco, rmb.rmbAgencia, rmb.rmbContaCorrente, rmb.chPessoa, rmb.rmbNomeColaborador, rmb.rmbValorReembolso,"
      sql = sql & " rmb.rmbTipoReembolsoTexto, rmb.rmbMeioPagtoTexto, gge.Alfa4, gge.Num2, gge.Num3"
      sql = sql & " From geradorgeral gge, reembolso rmb"
      sql = sql & " WHERE gge.chAlfaNumerica = ('" & Relatorio & "')"
      sql = sql & " AND rmb.rmbStatusReembolso = ('" & 1 & "') AND rmb.rmbStatusRecibo = ('" & 0 & "')"
      sql = sql & " ORDER BY rmb.chPessoa"
   
   AbrirRelatorio sql, Rel
   
   Call Rotina_AbrirBanco
   
   Rmb.Open "Select * from reembolso where rmbStatusRecibo = ('" & 0 & "') and rmbColaborador = ('" & cmbColaborador & "')", db, 3, 3
   If Rmb.EOF Then
      MsgBox ("Erro no acessoa reembolso na finalização da emissão de Recibo"), vbCritical
      Call FechaDB
      Exit Sub
   End If
   
   Rmb.MoveFirst
   
   Do While Not Rmb.EOF
      Rmb!rmbTiporeembolso = cmbReembolso.ListIndex
      Rmb!rmbTipoReembolsoTexto = cmbReembolso
      Rmb!rmbMeioPagto = cmbMeioPagto.ListIndex
      Rmb!rmbMeioPagtoTexto = cmbMeioPagto
      Rmb!RmbDataLancReembolso = Date
      Rmb!rmbStatusRecibo = 1
      pessoaAnterior = Rmb!chPessoa
      NotaFiscalAnterior = Rmb!chNotafiscal
      Rmb.Update
      Rmb.MoveNext
      
      If ctp.State = 1 Then
         ctp.Close: Set ctp = Nothing
      End If
      
      ctp.Open "Select * from contas_a_pagar where chPessoa = ('" & pessoaAnterior & "') and chNotaFiscal = ('" & NotaFiscalAnterior & "')", db, 3, 3
      If ctp.EOF Then
         MsgBox ("Contas a Pagar não encontrada. Erro. Comunicar ao analista responsável"), vbCritical
         Call FechaDB
         Exit Sub
      End If
      
      ctp!ctpDataPagamento = Date
      ctp!ctpStatus = 1
      ctp.Update
   Loop
   
   Call LimpaReembolso
End If
End Sub

'Private Sub cmdProcessaPagto_Click()

'If txtNumComprovante = "" Then
'   MsgBox ("Número do Comprovante de Depósito não informado."), vbInformation
'   Exit Sub
'End If

'Call Rotina_AbrirBanco

'db.BeginTrans

'Rmb.Open "Select * from reembolso where chPessoa = ('" & cmbPessoa & "') and chNotaFiscal = ('" & cmbNotaFiscal & "')", db, 3, 3
'If Rmb.EOF Then
'   MsgBox ("Erro no acesso a reembolso no Processamento de Pagamento."), vbInformation
'   Call FechaDB
'   Exit Sub
'End If

'If Rmb!rmbStatusReembolso = 1 Then
'   MsgBox ("Esse reembolso já foi Processado"), vbInformation
'   Call FechaDB
'   Exit Sub
'End If

'Rmb!rmbStatusReembolso = 1
'Rmb!rmbDataReembolso = dtDataDeposito

'Rmb.Update

'db.CommitTrans

'Call FechaDB

'MsgBox ("Processamento realizado com sucesso."), vbInformation

'End Sub

Private Sub cmdSair_Click()

Unload Me

End Sub

Public Sub GravarReembolso()
      

Call Rotina_AbrirBanco
      
Rmb.Open "Select * from reembolso where chPessoa = ('" & pessoaAnterior & "') and chNotaFiscal = ('" & NotaFiscalAnterior & "')", db, 3, 3

If Rmb.EOF Then
   Rmb.AddNew
   Inclusao = 1
End If
      
db.BeginTrans
      
      Rmb!chPessoa = pessoaAnterior
      Rmb!chNotafiscal = NotaFiscalAnterior
      Rmb!rmbFatura = Empty
      Rmb!RmbColaborador = cmbColaborador
      Rmb!RmbNomeColaborador = txtNomePessoa
      Rmb!rmbBanco = txtBanco
      Rmb!rmbAgencia = txtAgencia
      Rmb!rmbContaCorrente = txtContaCorrente
      Rmb!rmbCNPJ_CPF = txtCPF
      
      Rmb!RmbDataLancReembolso = Date
      
      Rmb!rmbDataReembolso = dtDataDeposito
      Rmb!rmbTiporeembolso = cmbReembolso.ListIndex
      Rmb!rmbTipoReembolsoTexto = cmbReembolso
      Rmb!rmbMeioPagto = cmbMeioPagto.ListIndex
      Rmb!rmbMeioPagtoTexto = cmbMeioPagto
      Rmb!rmbValorReembolso = AcumulaValorGrid
      Rmb!rmbNumComprovanteReembolso = txtNumComprovante
      
      nfe.Open "Select * from notafiscalentrada where chPessoa = ('" & pessoaAnterior & "') and chnotafiscalentrada = ('" & NotaFiscalAnterior & "')", db, 3, 3
      If nfe.EOF Then
         If nfe.State = 1 Then
            nfe.Close: Set nfe = Nothing
         End If
         nfe.Open "Select * from historiconotafiscalentrada where chPessoa = ('" & pessoaAnterior & "') and chnotafiscalentrada = ('" & NotaFiscalAnterior & "')", db, 3, 3
         If nfe.EOF Then
            MsgBox ("Nota Fiscal não encontrada. Data de hoje como data de emissão"), vbInformation
            Rmb!rmbDataNotaFiscal = Date
         Else
            Rmb!rmbDataNotaFiscal = nfe!nfeDataEmissao
         End If
      Else
         Rmb!rmbDataNotaFiscal = nfe!nfeDataEmissao
      End If
            
      Rmb!rmbStatusReembolso = 1
      Rmb!rmbStatusRecibo = 0

      Rmb.Update
      
db.CommitTrans

Call FechaDB

End Sub

Private Sub Form_Load()
Dim TipoPessoa2 As Integer
txtHoje = Date

Call Rotina_AbrirBanco

ContaReembolso = 0
dtDataDeposito = Date

TipoPessoa = 8
TipoPessoa2 = 5
pes.Open "Select * from pessoa where pesTipoPessoa > ('" & TipoPessoa2 & "') and pesTipoPessoa < ('" & TipoPessoa & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Pessoa sem cadastro de Colaboradores."), vbInformation
   Call FechaDB
   Exit Sub
End If

pes.MoveFirst

'Alteração: Modifiquei para colocar todos que tem status 2 em contas a pagar

Do While Not pes.EOF
   If ctp.State = 1 Then
      ctp.Close: Set ctp = Nothing
   End If
   ctp.Open "Select * from contas_a_pagar where ctpStatus = ('" & 2 & "') and ctpPessoaReembolso = ('" & pes!chPessoa & "')", db, 3, 3
   If Not ctp.EOF Then
      cmbColaborador.AddItem pes!chPessoa
      ContaReembolso = ContaReembolso + 1
   End If
   pes.MoveNext
Loop

'Versão antiga considerava o que estava cadastrado no reembolso

'Do While Not pes.EOF
'   If Rmb.State = 1 Then
'      Rmb.Close: Set Rmb = Nothing
'   End If
'   Rmb.Open "Select * from reembolso where rmbStatusReembolso = ('" & 0 & "') and rmbColaborador = ('" & pes!chPessoa & "')", db, 3, 3
'   If Not Rmb.EOF Then
'      cmbColaborador.AddItem pes!chPessoa
'      ContaReembolso = ContaReembolso + 1
'   End If
'   pes.MoveNext
'Loop

cmbReembolso.AddItem "Cash"
cmbReembolso.AddItem "Depósito"
cmbReembolso.AddItem "Transferência"

cmbMeioPagto.AddItem "EM ESPÉCIE"
cmbMeioPagto.AddItem "DOC"
cmbMeioPagto.AddItem "PIX"
cmbMeioPagto.AddItem "TED"
cmbMeioPagto.AddItem "TRANSF CONTA"

If ContaReembolso = 0 Then
   cmbColaborador.AddItem "NENHUM"
   MsgBox ("Não há reembolso até o presente momento."), vbInformation
   Call FechaDB
End If

End Sub

Public Sub CriticaEntrada()

Salvar = 0

If cmbColaborador = Empty Then
   'MsgBox ("Não Informado o Colaborador."), vbCritical
   Salvar = 0
   cmdSair.SetFocus
   Exit Sub
End If

If txtValorTotalDaNotaFiscal = Empty Then
   MsgBox ("Não Informado Produtos da Nota Fiscal ou o valor dos produtos é igual a zero."), vbCritical
   Salvar = 0
   Exit Sub
End If

If txtNumComprovante = Empty Then
   MsgBox ("Não Informado o Número do Comprovante do Documento de reembolso."), vbCritical
   Salvar = 0
   Exit Sub
End If

Salvar = 1

End Sub

Public Sub CriticaAlteracao()
Salvar = 1
If Rmb!rmbStatusRecibo = 1 Then
   Resp = MsgBox("Alteração inválida. O reembolso já foi efetuado e o Recibo já foi emitido. Continuar a alteração???", vbExclamation + vbYesNo)
   If Resp = vbYes Then
      Salvar = 1
   Else
      Salvar = 0
   End If
End If
If Rmb!rmbStatusReembolso = 1 Then
   Resp = MsgBox("Alteração inválida. O reembolso já foi efetuado mas o Recibo não foi emitido. Continuar a alteração???", vbExclamation + vbYesNo)
   If Resp = vbYes Then
      Salvar = 0
   Else
      Salvar = 1
   End If
End If

End Sub

Public Sub LimpaReembolso()


cmbColaborador.ListIndex = 0
txtNomePessoa = Empty
txtBanco = Empty
txtAgencia = Empty
txtContaCorrente = Empty
txtCPF = Empty
dtDataDeposito = Date
'cmbReembolso.ListIndex = 0
'cmbMeioPagto.ListIndex = 0
txtNumComprovante = Empty
gridDetalheNota.TextMatrix(1, 0) = Empty
gridDetalheNota.TextMatrix(1, 1) = Empty
gridDetalheNota.TextMatrix(1, 2) = Empty
gridDetalheNota.TextMatrix(1, 3) = Empty
gridDetalheNota.TextMatrix(1, 0) = Empty
gridDetalheNota.TextMatrix(1, 1) = Empty
gridDetalheNota.TextMatrix(1, 2) = Empty
gridDetalheNota.Rows = 1
Call LimpaCombo
'cmbPessoa.SetFocus

End Sub

'Public Sub CarregaReembolso()

'txtFatura = Rmb!rmbFatura
'dtDataNotaFiscal = Rmb!rmbDataNotaFiscal
'cmbColaborador = Rmb!rmbColaborador
'txtNomePessoa = Rmb!rmbNomeColaborador
'txtBanco = Rmb!rmbBanco
'txtAgencia = Rmb!rmbAgencia
'txtContaCorrente = Rmb!rmbContaCorrente
'txtCPF = Rmb!rmbCNPJ_CPF

'If Not IsNull(Rmb!rmbTipoReembolso) Then
'   cmbReembolso.ListIndex = Rmb!rmbTipoReembolso
'End If

'If Not IsNull(Rmb!rmbMeioPagto) Then
'   cmbMeioPagto.ListIndex = Rmb!rmbMeioPagto
'End If

'If Not IsNull(Rmb!rmbNumComprovanteReembolso) Then
'   txtNumComprovante = Rmb!rmbNumComprovanteReembolso
'End If

'If Not IsNull(Rmb!rmbDataReembolso) Then
'   dtDataDeposito = Rmb!rmbDataReembolso
'End If

'If Rmb!rmbStatusReembolso = 0 Then
''   txtStatusPagto = "PENDENTE"
'   txtStatusPagto.BackColor = vbRed
'Else
'   txtStatusPagto = "PROCESSADO"
'   txtStatusPagto.BackColor = vbCyan
'End If

'If Rmb!rmbStatusRecibo = 0 Then
'   txtStatusRecibo = "PENDENTE"
'   txtStatusRecibo.BackColor = vbRed
'Else
'   txtStatusRecibo = "PROCESSADO"
'   txtStatusRecibo.BackColor = vbCyan
'End If

'nfd.Open "Select * from notafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & cmbNotaFiscal & "')", db, 3, 3
'If nfd.EOF Then
'   MsgBox ("Nota Fiscal não encontrada."), vbCritical
'   Call FechaDB
'   Exit Sub
'End If'

'nfd.MoveFirst

'Ind = 1

'Do While Not nfd.EOF
'   gridDetalheNota.Rows = Ind + 1
'   gridDetalheNota.TextMatrix(Ind, 0) = nfd!chCodProduto
'   gridDetalheNota.TextMatrix(Ind, 1) = nfd!nfdQtd
'   gridDetalheNota.TextMatrix(Ind, 2) = Format$(nfd!nfdPU, "##0.00")
'   gridDetalheNota.TextMatrix(Ind, 3) = Format$(nfd!nfdValorParcela, "##0.00")
'
'   AcumulaValor = AcumulaValor + nfd!nfdValorParcela
'
'   Ind = Ind + 1
'
'   nfd.MoveNext
'Loop'

'IndSalvo = Ind

'txtValorTotalDaNotaFiscal = Format$(AcumulaValor, "##0.00")

'cmbColaborador.SetFocus


'End Sub

Public Sub LimpaCombo()

cmbReembolso.Clear
cmbMeioPagto.Clear

cmbReembolso.AddItem "Cash"
cmbReembolso.AddItem "Depósito"
cmbReembolso.AddItem "Transferência"

cmbMeioPagto.AddItem "EM ESPÉCIE"
cmbMeioPagto.AddItem "DOC"
cmbMeioPagto.AddItem "PIX"
cmbMeioPagto.AddItem "TED"
cmbMeioPagto.AddItem "TRANSF CONTA"
End Sub

Public Function Extenso()
'Public Function Extenso(AcumulaValor As Currency) As ValorPorExtenso
Dim nValor As Currency
nValor = AcumulaValor
'Valida Argumento
If IsNull(nValor) Or nValor <= 0 Or nValor > 9999999.99 Then
Exit Function
End If

'Variáveis
Dim nContador, nTamanho As Integer
Dim cValor, cParte, cFinal As String
ReDim aGrupo(4), aTexto(4) As String

'Matrizes de extensos (Parciais)
ReDim aUnid(19) As String
aUnid(1) = "um ": aUnid(2) = "dois ": aUnid(3) = "tres "
aUnid(4) = "quatro ": aUnid(5) = "cinco ": aUnid(6) = "seis "
aUnid(7) = "sete ": aUnid(8) = "oito ": aUnid(9) = "nove "
aUnid(10) = "dez ": aUnid(11) = "onze ": aUnid(12) = "doze "
aUnid(13) = "treze ": aUnid(14) = "quatorze ": aUnid(15) = "quinze "
aUnid(16) = "dezesseis ": aUnid(17) = "dezessete ": aUnid(18) = "dezoito "
aUnid(19) = "dezenove "

ReDim aDezena(9) As String
aDezena(1) = "dez ": aDezena(2) = "vinte ": aDezena(3) = "trinta "
aDezena(4) = "quarenta ": aDezena(5) = "cinquenta "
aDezena(6) = "sessenta ": aDezena(7) = "setenta ": aDezena(8) = "oitenta "
aDezena(9) = "noventa "

ReDim aCentena(9) As String
aCentena(1) = "cento ": aCentena(2) = "duzentos "
aCentena(3) = "trezentos ": aCentena(4) = "quatrocentos "
aCentena(5) = "quinhentos ": aCentena(6) = "seiscentos "
aCentena(7) = "setecentos ": aCentena(8) = "oitocentos "
aCentena(9) = "novecentos "

'Separa valor em grupos
cValor = Format$(nValor, "0000000000.00")
aGrupo(1) = Mid$(cValor, 2, 3)
aGrupo(2) = Mid$(cValor, 5, 3)
aGrupo(3) = Mid$(cValor, 8, 3)
aGrupo(4) = "0" + Mid$(cValor, 12, 2)

'Calcula cada grupo
For nContador = 1 To 4
  cParte = aGrupo(nContador)
  nTamanho = Switch(Val(cParte) < 10, 1, Val(cParte) < 100, 2, Val(cParte) < 1000, 3)
  If nTamanho = 3 Then
    If Right$(cParte, 2) <> "00" Then
      aTexto(nContador) = aTexto(nContador) + aCentena(Left(cParte, 1)) + "e "
      nTamanho = 2
    Else
      aTexto(nContador) = aTexto(nContador) + IIf(Left$(cParte, 1) = "1", "cem ", aCentena(Left(cParte, 1)))
    End If
  End If
  If nTamanho = 2 Then
    If Val(Right(cParte, 2)) < 20 Then
      aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 2))
    Else
      aTexto(nContador) = aTexto(nContador) + aDezena(Mid(cParte, 2, 1))
      If Right$(cParte, 1) <> "0" Then
        aTexto(nContador) = aTexto(nContador) + "e "
        nTamanho = 1
      End If
    End If
  End If
  If nTamanho = 1 Then
    aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 1))
  End If
Next

'Final
If Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And Val(aGrupo(4)) <> 0 Then
  cFinal = aTexto(4) + IIf(Val(aGrupo(4)) = 1, "centavo", "centavos")
Else
  cFinal = ""
  cFinal = cFinal + IIf(Val(aGrupo(1)) <> 0, aTexto(1) + IIf(Val(aGrupo(1)) > 1, "milhões ", "milhão "), "")
  If Val(aGrupo(2) + aGrupo(3)) = 0 Then
    cFinal = cFinal + "de "
  Else
    cFinal = cFinal + IIf(Val(aGrupo(2)) <> 0, aTexto(2) + "mil ", "")
  End If
  cFinal = cFinal + aTexto(3) + IIf(Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "real ", "reais ")
  cFinal = cFinal + IIf(Val(aGrupo(4)) <> 0, "E " + aTexto(4) + IIf(Val(aGrupo(4)) = 1, "centavo", "centavos"), "")
End If
ValorPorExtenso = UCase$(cFinal)

End Function





