VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOS 
   Caption         =   "frmOS"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16485
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   16485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdValidaOS 
      Caption         =   "Validar O.S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   14880
      TabIndex        =   28
      Top             =   3240
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtPrevisaoDeEntrega 
      Height          =   495
      Left            =   13080
      TabIndex        =   6
      Top             =   1320
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
      Format          =   242941953
      CurrentDate     =   45101
   End
   Begin VB.ComboBox cmbNumOs 
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
      Left            =   5520
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ComboBox cmbColaborador 
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
      Left            =   12480
      TabIndex        =   11
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox txtCompContrato 
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
      Left            =   5280
      TabIndex        =   9
      Top             =   2400
      Width           =   2415
   End
   Begin VB.ComboBox cmbPlataforma 
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
      Left            =   9120
      TabIndex        =   5
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdGeraOS 
      BackColor       =   &H0000FF00&
      Caption         =   "Gerar O.S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtRespComercial 
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
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox txtContrato 
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
      Left            =   3240
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtLocalEntrega 
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
      Left            =   7800
      TabIndex        =   10
      Top             =   2400
      Width           =   4575
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
      Left            =   7800
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox cmbProposta 
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
      Left            =   3720
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ComboBox cmbCliente 
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
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid grdDetProp 
      Height          =   3015
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      FormatString    =   "Qtd  |Equipto/Operador                         |Unid.|P.U         |Qtd.Unid|Valor diária       "
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
   Begin VB.Label txtDataEntrega 
      Alignment       =   2  'Center
      Caption         =   "Label13"
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
      Left            =   13560
      TabIndex        =   27
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Data Prevista para Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   26
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label10 
      Caption         =   "Email Para"
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
      Left            =   12480
      TabIndex        =   25
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Complemento Contrato"
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
      Left            =   5280
      TabIndex        =   24
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "Responsável Comercial"
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
      Left            =   240
      TabIndex        =   22
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   13440
      TabIndex        =   21
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Contrato"
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
      Left            =   3240
      TabIndex        =   20
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Local de Entrega"
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
      Left            =   7800
      TabIndex        =   19
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Plataforma"
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
      Left            =   9120
      TabIndex        =   18
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Número PO"
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
      Left            =   7800
      TabIndex        =   17
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Número OS"
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
      Left            =   5520
      TabIndex        =   16
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Proposta"
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
      Left            =   3720
      TabIndex        =   15
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
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
      TabIndex        =   14
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label frmOS 
      Caption         =   "Registro e Emissão de O.S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim Rel As Object
Dim Relatorio As String
Dim NumOrdemSevico As Integer

Private Sub cmbCliente_LostFocus()
   cmbProposta.Clear
   Call Rotina_AbrirBanco
      rs.Open "Select numProposta from proposta where cliente = ('" & cmbCliente & "') and status=1", db, 3, 3
      If rs.EOF Then
         MsgBox ("Erro: Não existem propostas para este cliente")
         FechaDB
         Exit Sub
      End If
      
      rs.MoveFirst
   
   Do While Not rs.EOF
      cmbProposta.AddItem rs!numProposta
      rs.MoveNext
   Loop
   
   rs.Close
   
   pes.Open "Select chUnidadeOperacional from unidadeoperacional where chPessoa = ('" & cmbCliente & "')", db, 3, 3
   If pes.EOF Then
      MsgBox ("Erro: Unidade Operacional não cadastrada para este cliente, favor cadastra-la"), vbInformation
      FechaDB
      Exit Sub
   End If
   
   pes.MoveFirst
   
   Do While Not pes.EOF
      cmbPlataforma.AddItem pes!chUnidadeOperacional
      pes.MoveNext
   Loop
   
   Call FechaDB
End Sub

Private Sub cmbNumOs_LostFocus()
   If cmbNumOS <> "Nova OS" Then
      Call Rotina_AbrirBanco
      rs.Open "Select * from ordemservico where numOS=('" & cmbNumOS & "') and numProposta = ('" & cmbProposta & "')", db, 3, 3
      cmbPlataforma = rs!plataforma
      txtRespComercial = rs!responsavelCliente
      txtContrato = rs!Contrato
      txtCompContrato = rs!complementoContrato
      txtLocalEntrega = rs!localEntrega
      dtPrevisaoDeEntrega = rs!dataPrevistaEntrega
      txtNumPO = rs!numPO
      Call FechaDB
   Else
      cmbPlataforma = Empty
      txtRespComercial = Empty
      txtContrato = Empty
      txtCompContrato = Empty
      txtLocalEntrega = Empty
   End If
End Sub

Private Sub cmbProposta_LostFocus()
   Call Rotina_AbrirBanco
   rs.Open "Select empNumOrdemDeServico from empresa ", db, 3, 3
   If rs.EOF Then
         MsgBox ("Erro: Falha ao gerar número da O.S")
         FechaDB
         Exit Sub
   End If

   Prod.Open "Select numOS from ordemservico where numProposta = ('" & cmbProposta & "')", db, 3, 3
   
   cmbNumOS.Clear
   cmbNumOS.AddItem "Nova OS"
   
   
   If Not Prod.EOF Then
   Prod.MoveFirst
   
      Do While Not Prod.EOF
      
         cmbNumOS.AddItem Prod!NumOS
         Prod.MoveNext
         
      Loop
   
   End If
   
   Call carga_grid

   rs.Close
   Call FechaDB
End Sub



Private Sub cmdGeraOS_Click()

On Error GoTo Erro

Call Rotina_AbrirBanco

db.BeginTrans

rs.Open "Select * from ordemservico where numProposta=('" & cmbProposta & "') and numOS=('" & cmbNumOS & "')", db, 3, 3
If rs.EOF Then
   rs.AddNew
End If

Prod.Open "Select emailResp,revisao,anoProposta from proposta where numProposta=('" & cmbProposta & "') and status = 1", db, 3, 3
If Not Prod.EOF Then
   If cmbNumOS = "Nova OS" Then
      Emp.Open "Select * from empresa where chPessoa = 'SHB Brasil'", db, 3, 3
      If Emp.EOF Then
         MsgBox ("ERRO: Erro sistema"), vbCritical
         Call FechaDB
         Exit Sub
      End If
      NumOrdemSevico = Emp!empNumOrdemDeServico + 1
      Emp!empNumOrdemDeServico = NumOrdemSevico
      Emp.Update
      cmbNumOS = NumOrdemSevico
   End If
   rs!NumOS = cmbNumOS
   rs!numProposta = cmbProposta
   rs!dataOS = txtDataEntrega
   rs!dataPrevistaEntrega = dtPrevisaoDeEntrega
   rs!Cliente = cmbCliente
   rs!Contrato = txtContrato
   rs!numPO = txtNumPO
   rs!responsavelCliente = txtRespComercial
   rs!plataforma = cmbPlataforma
   rs!localEntrega = txtLocalEntrega
   rs!Contato = Prod!emailResp
   rs!revisaoProposta = Prod!revisao
   rs!ano = Prod!anoProposta
   rs!complementoContrato = txtCompContrato
   rs.Update
End If

db.CommitTrans

Call FechaDB

'Call mandaEmail

Set Rel = drOrdemDeServicoNew
sql = "Select os.numOS, os.cliente, os.numProposta, os.numPo, os.revisaoProposta, os.dataPrevistaEntrega, os.dataOS, os.contrato, os.complementoContrato, os.responsavelCliente, os.plataforma, os.contato, os.localEntrega, pd.quantidade, pd.areaTotal, pd.equipamento, pd.dimensoes from ordemservico os, propostadetalhe pd where os.numOS = ('" & cmbNumOS & "') and pd.numProposta = os.numProposta"
AbrirRelatorio sql, Rel
Exit Sub
Erro: MsgBox ("Erro ao gerar OS: " & Err.Description), vbInformation
db.RollbackTrans
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdValidaOS_Click()
   Call Rotina_AbrirBanco
   
   rs.Open "Select * from ordemservico where numOS = ('" & cmbNumOS & "')", db, 3, 3
   If rs.EOF Then
   
      MsgBox ("OS inexistente, gerar OS antes de validar"), vbInformation
      FechaDB
      Exit Sub
   
   End If
   
   rs!Status = 1
   rs!validadaPor = glbUsuario
   rs!validadaNoDispositivo = glbMaquina
   rs.Update
   
   rs.Close
   FechaDB
   
End Sub

Private Sub Form_Load()
   Dim ano As Integer
   ano = Format$(Date, "yy")
   
   txtDataEntrega = Date
   dtPrevisaoDeEntrega = Date
   
   Call Rotina_AbrirBanco
      rs.Open "Select distinct cliente from proposta where status=1", db, 3, 3
      
      If rs.EOF Then
         MsgBox ("Erro: Clientes não possuem propostas aprovadas")
         FechaDB
         Exit Sub
      End If
   
   rs.MoveFirst
   
   Do While Not rs.EOF
      cmbCliente.AddItem rs!Cliente
      rs.MoveNext
   Loop
   
   rs.Close
   
   pes.Open "Select chPessoa from pessoa where pesTipoPessoa=7 and pesStatusPessoa=0", db, 3, 3
      
      If pes.EOF Then
         MsgBox ("Erro: Clientes não Registrados")
         FechaDB
         Exit Sub
      End If
   
   pes.MoveFirst
   
   Do While Not pes.EOF
      cmbColaborador.AddItem pes!chPessoa
      pes.MoveNext
   Loop
   
   pes.Close
   
   Prod.Open "Select * from empresa WHERE chPessoa = 'SHB Brasil'", db, 3, 3
   If Not Prod.EOF Then
      If Prod!empAnoOrdemDeServico < ano Then
         Prod!empNumOrdemDeServico = 0
         Prod!empAnoOrdemDeServico = ano
         Prod.Update
      End If
   End If
   
   Prod.Close
   Call FechaDB
End Sub
Public Sub carga_grid()
Dim Linha As Integer
   Call Rotina_AbrirBanco
   rs.Open "Select * from propostadetalhe where numProposta=('" & cmbProposta & "')", db, 3, 3
   If Not rs.EOF Then
      rs.MoveFirst
      Linha = 1
      Do While Not rs.EOF
         grdDetProp.Rows = Linha + 1
         grdDetProp.TextMatrix(Linha, 0) = rs!quantidade
         grdDetProp.TextMatrix(Linha, 1) = rs!equipamento
         grdDetProp.TextMatrix(Linha, 2) = rs!unidade
         grdDetProp.TextMatrix(Linha, 3) = Format$(rs!precoUnit, "##,##0.00")
         grdDetProp.TextMatrix(Linha, 4) = rs!areaTotal
         grdDetProp.TextMatrix(Linha, 5) = Format$(rs!diaria, "##,#0.00")
         Linha = Linha + 1
         rs.MoveNext
      Loop
      
   End If
End Sub


Public Sub mandaEmail()
   Dim outlookApp As Object
   Dim outlookMail As Object
   
   Set outlookApp = CreateObject("Outlook.Application")
   Set outlookMail = outlookApp.CreateItem(0)
   
   Call Rotina_AbrirBanco
   
   'rs.Open "Select numOS,numProposta from ordemservico where numOS=(Select MAX(numOS) from ordemservico)", db, 3, 3
   rs.Open "Select numOS,numProposta from ordemservico where numOS = ('" & cmbNumOS & "')", db, 3, 3
   pes.Open "Select pesEmail from pessoa where chPessoa = ('" & cmbColaborador & "')", db, 3, 3
   
   If pes.EOF Then
      MsgBox ("Registro não encontrado")
      Call FechaDB
      Exit Sub
   ElseIf IsNull(pes!pesEmail) Then
      MsgBox ("Cadastrar email no pessoa e retornar a esta função")
      Call FechaDB
      Exit Sub
   End If
   
   With outlookMail
      .To = pes!pesEmail
      .CC = ""
      .BCC = ""
      .Subject = "OS número: " & rs!NumOS & " relacionado a proposta número : " & rs!numProposta
      .Body = "Nova O.S foi gerada no sistema aguardando atendimento"
      .Send
    End With
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
   MsgBox ("Email enviado")
    
   FechaDB
End Sub
