VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReciboPagamento 
   Caption         =   "Recibo de Pagamentos"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid grdCtaPag 
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   7
      FormatString    =   "Nota fiscal     |Fatura      |Emissão     |Vencimento|Valor               | |Status"
   End
   Begin VB.Data RandCtaPagar 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Meus Documentos\dbSHB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Contas_A_Pagar"
      Top             =   5880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSMask.MaskEdBox txtHoje 
      Height          =   315
      Left            =   5520
      TabIndex        =   9
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker txtDataVenc 
      Height          =   315
      Left            =   5040
      TabIndex        =   5
      Top             =   5400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   260112385
      CurrentDate     =   39261
   End
   Begin VB.CommandButton cmbConfirma 
      BackColor       =   &H008080FF&
      Caption         =   "Confirma"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1335
   End
   Begin VB.ComboBox cmbBanco 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox txtNumeroDocumento 
      Height          =   315
      Left            =   3600
      TabIndex        =   4
      Top             =   5400
      Width           =   1095
   End
   Begin VB.ComboBox cmbTipoPagamento 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H0080FFFF&
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdCarregaCtaPagar 
      BackColor       =   &H0080FF80&
      Caption         =   "Contas a Pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox cmbColaborador 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label txtValor 
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
      Height          =   345
      Left            =   4680
      TabIndex        =   8
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmReciboPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Ind As Single
Dim Ano As Integer
Dim Mes As Integer
Dim Dia As Integer
Dim ValorTotal As Currency
Dim Resp As String
Dim Erro As Integer
'Global TabNfDesdobrComp As Recordset
'Global TabNfDesdobrParc As Recordset

Private Sub cmbConfirma_Click()
If ValorTotal = 0 Then
   MsgBox "Para executar este procedimento selecionar uma ou mais notas fiscais"
   Exit Sub
End If
Resp = MsgBox("Execução de Recibo de Pagamento solicitada. Confirma???", vbOKCancel)
If Not (Resp = vbOK) Then
   MsgBox "Rotina cancelada"
   Exit Sub
Else
   If txtNumeroDocumento = Empty Then
      MsgBox "Informar o número do Documento"
      txtNumeroDocumento.SetFocus
      Exit Sub
   End If
End If

BeginTrans

Erro = 0

For Ind = 1 To grdCtaPag.Rows - 1
    If grdCtaPag.TextMatrix(Ind, 5) = "Ok" Then
       Call AtualizaCtaPagar
    End If
Next
    
CommitTrans

If Erro = 0 Then

   TabGeradorGeral.AddNew
   TabGeradorGeral("chTipoGerador") = 112
   TabGeradorGeral("chTipoDoRelatorio") = 0
   TabGeradorGeral("chMaquina") = glbMaquina
   TabGeradorGeral("chAlfaNumerica") = cmbColaborador
   TabGeradorGeral("chNumerica") = cmbBanco.ListIndex
   TabGeradorGeral("chChaveData") = txtDataVenc
   TabGeradorGeral("Data2") = Date
   TabGeradorGeral("alfa2") = txtNumeroDocumento
   TabGeradorGeral("alfa3") = cmbBanco
   TabGeradorGeral.Update

   MsgBox "Impressão de Recibo será Iniciada. Preparar impressora e dar Ok"
   'impReciboDePagamento.Show vbModal
   'deReciboDePagamentos.rscmdReciboDePagamentos.Close
   TabGeradorGeral.MoveFirst
   Do While Not TabGeradorGeral.EOF
      If TabGeradorGeral("chTipoGerador") = 112 And TabGeradorGeral("chMaquina") = glbMaquina Then
         TabGeradorGeral.Delete
      End If
      TabGeradorGeral.MoveNext
   Loop
Else
   MsgBox "Erro na atualização de contas a pagar. Não haverá impressão"
   Exit Sub
End If
Call CargaGrid
End Sub

Private Sub cmdCarregaCtaPagar_Click()
Call CargaGrid
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Tabpessoa.MoveFirst
Do While Not Tabpessoa.EOF
   If Tabpessoa("pestipopessoa") > 0 Then
      cmbColaborador.AddItem Tabpessoa("chpessoa")
   End If
   Tabpessoa.MoveNext
Loop
cmbTipoPagamento.AddItem "Cheque"
cmbTipoPagamento.AddItem "Especie"

TabBanco.MoveFirst
Do While Not TabBanco.EOF
   cmbBanco.AddItem TabBanco("bcosiglabco")
   TabBanco.MoveNext
Loop
   
cmbTipoPagamento.ListIndex = 0
cmbBanco.ListIndex = 0
txtDataVenc = Date
txtHoje = Date
End Sub


Private Sub grdCtaPag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ValorTotal = 0
Ind = grdCtaPag.RowSel
If grdCtaPag.TextMatrix(Ind, 5) = "Ok" Then
   grdCtaPag.TextMatrix(Ind, 5) = " "
Else
   grdCtaPag.TextMatrix(Ind, 5) = "Ok"
End If

For Ind = 1 To grdCtaPag.Rows - 1
    If grdCtaPag.TextMatrix(Ind, 5) = "Ok" Then
       ValorTotal = ValorTotal + grdCtaPag.TextMatrix(Ind, 4)
    End If
Next

txtValor = Format$(ValorTotal, "###,###,##0.00")
End Sub

Public Sub LimpaGrid()
grdCtaPag.Rows = 2
grdCtaPag.TextMatrix(1, 0) = Empty
grdCtaPag.TextMatrix(1, 1) = Empty
grdCtaPag.TextMatrix(1, 2) = Empty
grdCtaPag.TextMatrix(1, 3) = Empty
grdCtaPag.TextMatrix(1, 4) = Empty
grdCtaPag.TextMatrix(1, 5) = Empty
grdCtaPag.TextMatrix(1, 6) = Empty
End Sub

Public Sub AtualizaCtaPagar()

TabCtaPagar.Seek "=", 0, cmbColaborador, grdCtaPag.TextMatrix(Ind, 0), grdCtaPag.TextMatrix(Ind, 1), grdCtaPag.TextMatrix(Ind, 3)
If TabCtaPagar.NoMatch Then
   MsgBox "Erro no acesso a Contas a Pagar"
   Ind = grdCtaPag.Rows - 1
   Erro = Erro + 1
   Exit Sub
End If

TabNfDesdobrComp.Seek "=", cmbColaborador, grdCtaPag.TextMatrix(Ind, 0), grdCtaPag.TextMatrix(Ind, 3)
If TabNfDesdobrComp.NoMatch Then
   MsgBox "Erro no acesso a desdobramento de Contas a pagar"
   Ind = grdCtaPag.Rows - 1
   Erro = Erro + 1
   Exit Sub
End If

TabCtaPagarNF.AddNew
TabCtaPagarNF("chFabricante") = TabCtaPagar("chFabricante")
TabCtaPagarNF("chPessoa") = TabCtaPagar("chPessoa")
TabCtaPagarNF("chNotaFiscal") = TabCtaPagar("chNotaFiscal")
TabCtaPagarNF("chFatura") = txtNumeroDocumento
TabCtaPagarNF("chDataVencito") = txtDataVenc
TabCtaPagarNF("ctpDataEmissao") = TabCtaPagar("ctpDataEmissao")
TabCtaPagarNF("ctpDataBanco") = txtDataVenc
TabCtaPagarNF("ctpDataLanc") = TabCtaPagar("ctpDataLanc")
TabCtaPagarNF("ctpDataVencOriginal") = TabCtaPagar("ctpDataVencOriginal")
TabCtaPagarNF("ctpDescricaoOperacao") = TabCtaPagar("ctpDescricaoOperacao")
TabCtaPagarNF("ctpValorLart") = TabCtaPagar("ctpValorLart")
TabCtaPagarNF("ctpValorMerco") = TabCtaPagar("ctpValorMerco")
TabCtaPagarNF("ctpValorDaBoleta") = TabCtaPagar("ctpValorDaBoleta")
TabCtaPagarNF("chAno") = TabCtaPagar("chAno")
TabCtaPagarNF("chMes") = TabCtaPagar("chMes")
TabCtaPagarNF("chDia") = TabCtaPagar("chDia")
TabCtaPagarNF("chNumPedido") = TabCtaPagar("chNumPedido")
TabCtaPagarNF("chNumPedidoComp") = TabCtaPagar("chNumPedidoComp")
TabCtaPagarNF("chCodBcoLart") = TabCtaPagar("chCodBcoLart")
TabCtaPagarNF("ctpStatus") = TabCtaPagar("ctpStatus")
TabCtaPagarNF("ctpDataProc") = TabCtaPagar("ctpDataProc")
TabCtaPagarNF("ctpDataPagamento") = TabCtaPagar("ctpDataPagamento")

If cmbTipoPagamento.ListIndex = 0 Then
   TabCtaPagarNF("ctpTipoLancamento") = 1
Else
   TabCtaPagarNF("ctpTipoLancamento") = 8
End If

TabCtaPagarNF.Update

TabCtaPagar.Delete

TabNfDesdobrParc.AddNew
TabNfDesdobrParc("chPessoa") = TabNfDesdobrComp("chPessoa")
TabNfDesdobrParc("chNotaFiscalEntrada") = TabNfDesdobrComp("chNotaFiscalEntrada")
TabNfDesdobrParc("chDataVencimento") = txtDataVenc
TabNfDesdobrParc("nfdDataVencOriginal") = TabNfDesdobrComp("nfdDataVencOriginal")
TabNfDesdobrParc("nfdDataPagamento") = TabNfDesdobrComp("nfdDataPagamento")
TabNfDesdobrParc("nfdFaturaNumero") = txtNumeroDocumento
TabNfDesdobrParc("nfdValorDaFatura") = TabNfDesdobrComp("nfdValorDaFatura")
TabNfDesdobrParc("nfdStatus") = TabNfDesdobrComp("nfdStatus")
TabNfDesdobrParc("nfdStatusPagto") = TabNfDesdobrComp("nfdStatusPagto")
TabNfDesdobrParc("nfdOrdemBoleto") = TabNfDesdobrComp("nfdOrdemBoleto")
TabNfDesdobrComp.Delete
TabNfDesdobrParc.Update
End Sub

Public Sub CargaGrid()
Call LimpaGrid
Ind = 0
ValorTotal = Format$(0, "0.00")
txtValor = Format$(0, "0.00")
txtNumeroDocumento = Empty
If TabCtaPagar.RecordCount = 0 Then
   Exit Sub
End If

TabCtaPagar.MoveFirst
Do While Not TabCtaPagar.EOF
   If TabCtaPagar("chpessoa") = cmbColaborador Then
      If TabCtaPagar("ctpstatus") = 0 And Not (TabCtaPagar("ctptipolancamento") = 1 Or TabCtaPagar("ctptipolancamento") = 8) Then
         Ind = Ind + 1
         grdCtaPag.Rows = Ind + 1
         grdCtaPag.TextMatrix(Ind, 0) = TabCtaPagar("chNotaFiscal")
         grdCtaPag.TextMatrix(Ind, 1) = TabCtaPagar("chFatura")
         grdCtaPag.TextMatrix(Ind, 2) = TabCtaPagar("ctpDataEmissao")
         grdCtaPag.TextMatrix(Ind, 3) = TabCtaPagar("chDataVencito")
         grdCtaPag.TextMatrix(Ind, 4) = Format$(TabCtaPagar("ctpValorDaBoleta"), "###,###,##0.00")
         Ano = Year(TabCtaPagar("chDataVencito"))
         Mes = Month(TabCtaPagar("chDataVencito"))
         Dia = Day(TabCtaPagar("chDataVencito"))
         grdCtaPag.TextMatrix(Ind, 6) = Ano & Format$(Mes, "00") & Format$(Dia, "00")
      End If
   End If
   TabCtaPagar.MoveNext
Loop
grdCtaPag.Row = 1
grdCtaPag.RowSel = grdCtaPag.Rows - 1
grdCtaPag.Col = 6
grdCtaPag.ColSel = 6
grdCtaPag.Sort = 1
grdCtaPag.Col = 0
grdCtaPag.Row = 0
grdCtaPag.ColSel = 0
grdCtaPag.RowSel = 0
End Sub

Private Sub txtNumeroDocumento_LostFocus()

RandCtaPagar.Refresh
RandCtaPagar.Recordset.FindFirst "chFatura = '" & txtNumeroDocumento & "'"

If Not (RandCtaPagar.Recordset.NoMatch) Then
   MsgBox "Este Documento já foi utilizado. Favor Verificar."
   
   cmdSair.SetFocus
End If
'tabctapagar.Seek =

End Sub

