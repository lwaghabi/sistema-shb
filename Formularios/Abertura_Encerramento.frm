VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAbre_Fecha 
   Caption         =   "Abertura e Encerramento do Sistema"
   ClientHeight    =   6510
   ClientLeft      =   2115
   ClientTop       =   2310
   ClientWidth     =   11205
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   11205
   WindowState     =   2  'Maximized
   Begin ComctlLib.ProgressBar pbCopiaArquivos 
      Height          =   495
      Left            =   4320
      TabIndex        =   27
      Top             =   3960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker txtHoje 
      Height          =   375
      Left            =   1680
      TabIndex        =   25
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   243793921
      CurrentDate     =   43883
   End
   Begin VB.Frame Frame3 
      Caption         =   "Encerramento de mes"
      Height          =   4575
      Left            =   8280
      TabIndex        =   15
      Top             =   1680
      Width           =   2655
      Begin VB.TextBox txtSuprimento 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox txtNotaFiscal 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtProcessaCtaReceber 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtProcessaCtaPagar 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox txtProcessaPedido 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtProcessaEstoque 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label txtRelApos 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label txtRelAnt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5640
      TabIndex        =   11
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtData 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5640
      TabIndex        =   12
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtHora 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5640
      TabIndex        =   13
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Status"
      Height          =   1695
      Left            =   5040
      TabIndex        =   14
      Top             =   1680
      Width           =   2535
   End
   Begin VB.OptionButton optFecha 
      Caption         =   "Encerramento do Sistema"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   2760
      Width           =   1935
   End
   Begin VB.OptionButton optAbre 
      Caption         =   "Abertura do Sistema"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   3975
      Begin MSComCtl2.DTPicker txtDataEvento 
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   243859457
         CurrentDate     =   43883
      End
      Begin VB.TextBox txtHoraEvento 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   4
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data do Evento"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hora do Evento"
         Height          =   195
         Left            =   2520
         TabIndex        =   9
         Top             =   2040
         Width           =   1125
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
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
      Height          =   360
      Left            =   1920
      TabIndex        =   21
      Top             =   5040
      Width           =   630
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Abertura e Encerramento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3975
      TabIndex        =   6
      Top             =   1200
      Width           =   3225
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   " Sistema Integrado SHB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3645
      TabIndex        =   4
      Top             =   720
      Width           =   3840
   End
End
Attribute VB_Name = "frmAbre_Fecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DataLogin As Double
Dim DataInvertida As String

Dim ano As Integer
Dim mes As Integer
Dim Dia As Integer

Dim FSys As New FileSystemObject
Dim Servidor As String
Dim DriveDestino As String
Dim TabIndex As Index
Dim Incluir As Byte
Dim baseaberta As Byte
Dim msg As String
Dim rotinicial As Byte
Dim Data_Hoje As Date
Dim Data_Anterior As Date
Dim StatusSistema As Byte
Dim Resp As String
Dim Fim As Byte
Dim Fim_Detalhe As Byte
Dim Chave_Gira_mes As Byte
Dim DataProdCli As Integer
Dim AnoProdCli As Integer
Dim MesProdCli As Integer
Dim DiaProdCli As Integer
Dim SistemaEncerrado As Byte
Dim Hora As Double
Dim AcumQtdAnoCli As Currency
Dim AcumQtdAnoCliProd As Currency
Dim ClienteAnterior As String
Dim GuardaPessoa As String
Dim QtdAnterior As Currency
Dim FaturaAnterior As Currency
Dim pessoaAnterior As String
Dim NotaFiscalAnterior As String
Dim Encontrei As Byte

Dim qtd As Currency
Dim Fatura As Currency

Dim Origem As String
Dim destino As String
Dim i As Integer

Dim TesteProgressao As Long

Dim KeyPedido As String
Dim KeyCompPedido As String
Dim KeyProduto As String

Dim Erro As Byte


Private Sub cmdFechar_Click()
     Unload Me
End Sub


Private Sub cmdOk_Click()
Erro = 0
SistemaEncerrado = 0
     On Error Resume Next
   
     If optAbre = False And optFecha = False Then
        MsgBox ("Informe o Status do Sistema"), vbInformation
     End If
     
     If txtDataEvento = Empty Then
        MsgBox ("Informe a Data "), vbInformation
     End If
     
     If txtDataEvento <> txtHoje Then
        Resp = MsgBox("Data Informada difere de Data Hoje. Continua??? s/n", vbYesNo)
        If Resp = vbNo Then
           txtDataEvento = Empty
           MsgBox ("Informe novamente a Data do Evento"), vbInformation
           txtDataEvento.SetFocus
           Exit Sub
        End If
     End If

If optAbre = True Then
   Call AberturaDoSistema
Else
   Call EncerramentoDoSistema
End If
Call CarregaCampos
End Sub

Public Sub AberturaDoSistema()

   
Call Rotina_AbrirBanco
     
glb.Open "select * from global where chDataAbertura = ('" & DataHojeInvertida & "')", db, 3, 3
acGlb = acGlb + 1
If glb.EOF Then
   glb.AddNew
   'glb!chDataAbertura = txtDataEvento
   glb!chDataAbertura = DataHojeInvertida
   glb!glostatussistema = 1

   glb!gloHoraAbertura = Format$(Time, "hh:mm:ss")
   glb!glodataAberturaanterior = Data_Anterior
   glb!gloStatusAtuMensal = 0
   glb.Update
   StatusSistema = 1
   Call CarregaCampos
   cmdFechar.SetFocus
   If Month(Data_Hoje) <> Month(Data_Anterior) Then
      Resp = MsgBox("Execução de Rotina Mensal. Confirma???", vbYesNo)
      If Resp = vbYes Then
         Call Rotina_Execucao_Mensal
         'glb!gloStatusAtuMensal = 1
         'glb.Update
      Else
         Resp = MsgBox("Você deseja abrir o sistema sem a execução mensal????. Confirma???", vbYesNo)
         If Resp = vbYes Then
            glb!gloStatusAtuMensal = 0
            glb.Update
            Call CarregaCampos
         Else
            Exit Sub
         End If
      End If
   End If
Else
   If glb!glostatussistema = 1 And optAbre = True Then
      MsgBox ("O Sistema já está aberto"), vbInformation
      optAbre = Empty
   Else
      Resp = MsgBox("O sistema foi encerrado. Voce deseja reabri-lo?", vbYesNo)
      If Resp = vbYes Then
         StatusSistema = 1
         glb!glostatussistema = 1
         glb.Update
      End If
   End If
End If

Call FechaDB
      
End Sub

Public Sub EncerramentoDoSistema()

SistemaEncerrado = 0

Call Rotina_AbrirBanco
     
glb.Open "select * from global where chDataAbertura = ('" & DataHojeInvertida & "')", db, 3, 3
acGlb = acGlb + 1
If glb.EOF Then
   MsgBox ("Não encontrado o global na rotina de encerramento"), vbCritical
Else
   
   db.BeginTrans
   
      glb!glostatussistema = 0
      glb!gloDataEncerramento = Date
      glb!gloHoraEncerramento = Format$(txtHoraEvento, "hh:mm:ss")
      MsgBox ("Sistema Encerrado"), vbInformation
      SistemaEncerrado = 1
      glb.Update
     
  db.CommitTrans

End If

Call CarregaCampos

Call FechaDB

If SistemaEncerrado = 1 Then
   Resp = MsgBox("Deseja Fazer o Backup Agora?????", vbYesNo)

   If Resp = vbYes Then
      Hora = Format$(Time, "hhmmss")

      DriveDestino = InputBox("Informe o Drive de Destino", "Drive de Destino")
           
      FSys.CopyFolder "C:\Meus Documentos\SISTEMA SHB\Projeto SHB MySql", DriveDestino & ":\Meus Documentos\Projeto SHB Backups\Projeto SHB MySql" & (Year(Date) & Format$(Month(Date), "00") & Format$(Day(Date), "00") & "_" & Hora), True

      'Origem = "C:\Meus Documentos\SISTEMA SHB\dbSHB.mdb" 'Novo
      'destino = DriveDestino & ":\Meus Documentos\Projeto SHB Backups\dbSHB" & (Year(Date) & Format$(Month(Date), "00") & Format$(Day(Date), "00") & "_" & Hora & ".mdb") 'Novo
      'pbCopiaArquivos.Value = CopiarArquivos(Origem, destino)
      ' Call CopiarArquivos(Origem, destino)
       MsgBox ("Backup realizado com sucesso"), vbInformation
    End If
      MsgBox ("Backuo não realizado"), vbCritical
      SistemaEncerrado = 0
End If
End Sub

Private Sub Form_Load()

Data_Hoje = Date
txtDataEvento = Date
txtHoje = Data_Hoje
               
txtHoraEvento = Time()

Call Rotina_AbrirBanco
   
 
glb.Open "select * from global where chDataAbertura = ('" & DataHojeInvertida & "')", db, 3, 3
acGlb = acGlb + 1
If glb.EOF Then
   StatusSistema = 0
Else
   StatusSistema = glb!glostatussistema
   DataLogin = glb!chDataAbertura
End If
If StatusSistema = 1 Then
   optFecha.Enabled = True
   optAbre.Enabled = False
Else
   optAbre.Enabled = True
   optFecha.Enabled = False
End If

Call CarregaCampos

Call FechaDB

End Sub

Private Sub optAbre_Click()

Call Rotina_AbrirBanco

glb.Open "select * from global", db, 3, 3
acGlb = acGlb + 1
If glb.EOF Then
   MsgBox ("Global sem registro"), vbCritical
   End
Else
   glb.MoveLast
   Data_Anterior = glb!chDataAbertura
   'If glb!glostatussistema = 1 Then
   '   MsgBox ("Sistema encontra-se aberto. Para abertura ou reabertura é necessário o seu encerramento"), vbInformation
   ' End If
End If
      
Call FechaDB

End Sub

Public Sub CarregaCampos()

 If StatusSistema = 1 Then
     If glb!glostatussistema = 1 Then
          txtStatus = "Aberto"
          txtStatus.BackColor = vbCyan
          txtData = glb!chDataAbertura
          txtData.BackColor = vbCyan
          txtHora = Format$(glb!gloHoraAbertura, "hh:mm:ss")
          txtHora.BackColor = vbCyan
          
     Else
          txtStatus = "Fechado"
          txtStatus.BackColor = vbRed
          txtData = glb!gloDataEncerramento
          txtData.BackColor = vbRed
          txtHora = glb!gloHoraEncerramento
          txtHora.BackColor = vbRed
     End If
Else
     txtStatus = "Fechado"
     txtStatus.BackColor = vbRed
     txtData = "__/__/____"
     txtData.BackColor = vbRed
     txtHora = "__:__:__"
     txtHora.BackColor = vbRed
End If
End Sub

Public Sub Rotina_Execucao_Mensal()

txtRelAnt = "Inicio Rotina Mensal"
'txtRelAnt = "Emissão de Relatórios do Mês"
'txtRelAnt.BackColor = vbRed
'frmAbre_Fecha.Refresh
'impNegMes.Show vbModal
'impNegMesConsig.Show vbModal
'impCtaPagas.Show vbModal
'impCtaRecebidas.Show vbModal
'txtRelAnt = "Fim da Emissão de Relatórios do Mês"
txtRelAnt.BackColor = vbCyan
'Chave_Gira_mes = 0
'txtProcessaEstoque = "Processando Estoque"
'txtProcessaEstoque.BackColor = vbRed
frmAbre_Fecha.Refresh
'Call Rotina_Processa_Estoque
'If Erro = 0 Then
   txtProcessaEstoque = "Estoque Processado"
   txtProcessaEstoque.BackColor = vbCyan
   txtProcessaPedido = "Processando Pedido"
   txtProcessaPedido.BackColor = vbRed
   frmAbre_Fecha.Refresh
   Call Rotina_Processa_Pedido
   If Erro = 0 Then
      txtProcessaPedido = "Pedido Processado"
      txtProcessaPedido.BackColor = vbCyan
      txtProcessaCtaReceber = "Processando Contas a Receber"
      txtProcessaCtaReceber.BackColor = vbRed
      frmAbre_Fecha.Refresh
      Call Rotina_Processa_Contas_Receber
      If Erro = 0 Then
         txtProcessaCtaReceber = "Contas a Receber Processado"
         txtProcessaCtaReceber.BackColor = vbCyan
         frmAbre_Fecha.Refresh
         txtNotaFiscal = "Processando Nota Fiscal"
         txtNotaFiscal.BackColor = vbRed
         frmAbre_Fecha.Refresh
         Call Rotina_Nota_Fiscal
         If Erro = 0 Then
            txtNotaFiscal = "Nota Fiscal Processado"
            txtNotaFiscal.BackColor = vbCyan
            frmAbre_Fecha.Refresh
            If Erro = 0 Then
               txtProcessaCtaPagar = "Processando Contas a Pagar"
               txtProcessaCtaPagar.BackColor = vbRed
               frmAbre_Fecha.Refresh
               Call Rotina_Processa_Contas_Pagar
               txtProcessaCtaPagar = "Contas a Pagar Processado"
               txtProcessaCtaPagar.BackColor = vbCyan
               If Erro = 0 Then
                  txtProcessaCtaPagar = "Processando Suprimento Mensal"
                  txtSuprimento.BackColor = vbRed
                  frmAbre_Fecha.Refresh
                  db.Execute ("INSERT INTO supestoquehist SELECT grupo,classe,codProd,qtdEmEstoque,qtdReservado,estoqueMinimo,estoqueMaximo,CURDATE() FROM supestoque")
                  txtSuprimento = "Suprimento Mensal Processado"
                  txtSuprimento.BackColor = vbCyan
            
                  If Erro = 0 Then
                     txtSuprimento = "Fim da Rotina Mensal"
                     txtSuprimento.BackColor = vbCyan
                     frmAbre_Fecha.Refresh
                  End If
               End If
            End If
         End If
      End If
   End If

End Sub

'Public Sub Rotina_Processa_Estoque()
'On Error GoTo ErroProcessaEstoque
'fim = 0
'If TabEstoqueAnterior.RecordCount > 0 Then
'   TabEstoqueAnterior.MoveLast
'   If TabEstoqueAnterior.NoMatch Then
'      fim = 1
'   End If
'Else
'   fim = 1
'End If
'Do While fim = 0
'   TabEstoqueProdAcabado.AddNew
'   TabEstoqueProdAcabado("chAno") = Year(Data_Hoje)
'   TabEstoqueProdAcabado("chmes") = Month(Data_Hoje)
'   TabEstoqueProdAcabado("chProduto") = TabEstoqueAnterior("chProduto")
'   TabEstoqueProdAcabado("estqtdtracos") = 0
'   TabEstoqueProdAcabado("estqtdinicial") = TabEstoqueAnterior("estsaldoatual")
'   TabEstoqueProdAcabado("estSaldoAtual") = TabEstoqueAnterior("estsaldoatual")
'   TabEstoqueProdAcabado("estentradaproduto") = 0
'   TabEstoqueProdAcabado("estentradaprodutodevolucao") = 0
'   TabEstoqueProdAcabado("estsaidaprodutovenda") = 0
'   TabEstoqueProdAcabado("estsaidaprodutoassisttec") = 0
'   TabEstoqueProdAcabado("estsaidaprodutopromocao") = 0
'   TabEstoqueProdAcabado("estpedidomesanter") = TabEstoqueAnterior("estpedidomesanter") + TabEstoqueAnterior("estpedidomesatual")
'   TabEstoqueProdAcabado("estPedidoEmCarteira") = TabEstoqueAnterior("estpedidomesanter") + TabEstoqueAnterior("estpedidomesatual")
'   TabEstoqueProdAcabado("estpedidomesatual") = 0
'   UltimoRegistro = Year(Data_Hoje) & " - " & Month(Data_Hoje) & " - " & TabEstoqueAnterior("chProduto")
'   TabEstoqueProdAcabado.Update
'   TabEstoqueAnterior.MovePrevious
'   If TabEstoqueAnterior.BOF Then
'      fim = 1
'   Else
'      If TabEstoqueAnterior("chmes") = Month(TabGlobal("gloDataAberturaAnterior")) Then
'         fim = 0
'      Else
'         fim = 1
'      End If
'   End If
'Loop'
'
'Exit Sub

'ErroProcessaEstoque:

'    MsgBox UltimoRegistro
'    Erro = Erro + 1
'    MsgBox Err.Description
    
'End Sub

Public Sub Rotina_Processa_Pedido()
On Error GoTo ErroPedido

Fim_Detalhe = 0
Fim = 0

Call Rotina_AbrirBanco
     
neg.Open "select * from negociacao where negStatus = ('" & 1 & "')", db, 3, 3
acNeg = acNeg + 1
If neg.EOF Then
   MsgBox ("Não há movimento em Negociação"), vbInformation
   Call FechaDB
   Exit Sub
End If

dneg.Open "select * from detalhenegociacao where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
acdNeg = acdNeg + 1
If dneg.EOF Then
   MsgBox ("Negociação sem Detalhe"), vbCritical
   Call FechaDB
   Exit Sub
End If

pes.Open "select * from pessoa where chPessoa = ('" & neg!chPessoa & "')", db, 3, 3
acPes = acPes + 1
If pes.EOF Then
   MsgBox ("Erro no acesso a pessoa"), vbCritical
   Call FechaDB
   End
End If

hneg.Open "select * from historiconegociacao", db, 3, 3
achNeg = achNeg + 1
'If hneg.EOF Then
'   MsgBox ("Historico Negociação Vazio"), vbInformation
'End If


hdneg.Open "select * from historicodetalhenegociacao", db, 3, 3
achdNeg = achdNeg + 1
'If hdneg.EOF Then
'   MsgBox ("Historico Detalhe Negociação Vazio"), vbInformation
'End If
   

db.BeginTrans

Do While Fim = 0

   hdneg.AddNew
   hdneg!chPessoa = pes!chPessoa
   hdneg!chNumPedido = dneg!chNumPedido
   hdneg!chNumPedidoComp = dneg!chNumPedidoComp
   hdneg!chProduto = dneg!chProduto
   hdneg!chDataInicio = dneg!chDataInicio
   hdneg!chDataFim = dneg!chDataFim
   hdneg!hdnAtividade = dneg!pedAtividade
   hdneg!hdnUnidade = dneg!pedunidade
   hdneg!hdnquantidadePedida = dneg!pedquantidadePedida
   hdneg!hdnPrecoUnidadePedida = dneg!pedPrecoUnidadePedida
   hdneg!hdnValorDaDiaria = dneg!pedValorDaDiaria
   hdneg!hdnQtdDias = dneg!pedqtddias
   hdneg!hdnValorDaOperacao = dneg!pedValorDaOperacao
   hdneg!hdncomissaorep = dneg!pedcomissaorep
   hdneg!hdncomissaopromot = dneg!pedcomissaopromot
   hdneg!hdnStatus = dneg!pedStatus
        
   hdneg.Update
   
   ultimoRegistro = "Grava Hist Neg - " & pes!chPessoa & " - " & dneg!chNumPedido & " - " & dneg!chNumPedidoComp & " - " & dneg!chDataInicio & " - " & dneg!chDataFim
   
  dneg.Delete
   
   dneg.MoveNext
   
   If dneg.EOF Then
      Call MovimentaNegHistNeg
      neg.Delete
      neg.MoveNext
      acdNeg = acdNeg + 1
      If neg.EOF Then
         Fim = 1
      Else
         If dneg.State = 1 Then
            dneg.Close: Set dneg = Nothing
            acdNeg = 0
            dneg.Open "select * from detalhenegociacao where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
            acdNeg = acdNeg + 1
            If dneg.EOF Then
               MsgBox ("Negociação sem Detalhe"), vbCritical
               Call FechaDB
               Exit Sub
            End If
            If pes.State = 1 Then
               pes.Close: Set pes = Nothing
               acPes = 0
               pes.Open "select * from pessoa where chPessoa = ('" & neg!chPessoa & "')", db, 3, 3
               acPes = acPes + 1
               If pes.EOF Then
                  MsgBox ("Erro no acesso a pessoa"), vbCritical
                  Call FechaDB
                  End
               End If
            End If
          End If
       End If
   End If
Loop

db.CommitTrans

 'UltimoRegistro = "Grava Hist Det Neg - " & pes!chPessoa & " - " & dneg!chNumPedido & " - " & dneg!chNumPedidoComp
                  
Call FechaDB

Exit Sub

ErroPedido:

    MsgBox ultimoRegistro
    Erro = Erro + 1
    MsgBox Err.Description

End Sub

Public Sub MovimentaNegHistNeg()

hneg.AddNew

hneg!chPessoa = neg!chPessoa
hneg!chNumPedido = neg!chNumPedido
hneg!negContrato = neg!negContrato
hneg!negContratoComp = neg!negContratoComp
hneg!chNumPedidoComp = neg!chNumPedidoComp
hneg!negNumFatura = neg!negNumFatura
hneg!chUnidadeOperacional = neg!chUnidadeOperacional
hneg!negSerieFatura = neg!negSerieFatura
hneg!negDataEmissaoFatura = neg!negDataEmissaoFatura
hneg!hngnotafiscal = neg!negNotaFiscal
hneg!hngEmissorNF = neg!negEmissorNF
hneg!chCodBcoLart = neg!chCodBcoLart
'hneg!chOrdemDeCarga = neg!chOrdemDeCarga
'hneg!hngTransporte = neg!negTransporte
'hneg!hngPlaca = neg!negPlaca
hneg!hngStatus = neg!negStatus
hneg!hngDataPedido = neg!negDataPedido
hneg!chdatanegociacao = neg!negdatanegociação
hneg!hngValornegociacao = neg!negvalornegociacao
hneg!hnegTipoProduto = neg!negTipoProduto
hneg!chrepresentante = neg!chrepresentante
hneg!chPromotor = neg!chPromotor
hneg!hngFaturamento = neg!negFaturamento
hneg!hngintervalofatura = neg!negIntervaloFatura
hneg!hngAPartirDe = neg!negAPartirDe
hneg!hngFreteColeta = neg!negFreteColeta
hneg!hngBoletaFrete = neg!negboletafrete
hneg!hngValorFixoFrete = neg!negValorFixoFrete
hneg!hngCondProcess = neg!negCondProcess
hneg!hngdesccomissao = neg!negdesccomissao
'hneg!hngDescComissaoPromot = neg!negDescComissaoPromot
hneg!hngPrazoAdicional = neg!negPrazoAdicional
'hneg!hngLançamento = neg!negLançamento
hneg!hngUltimaAtualizacao = neg!negUltimaAtualizacao
'hneg!negCntrlFaturamento = neg!negCntrlFaturamento
hneg!negICMS = neg!negICMS
hneg!negAliquota = neg!negAliquota
hneg!negFretePedido = neg!negFretePedido
hneg!negValorDoProduto = neg!negValorDoProduto
hneg!negIPI = neg!negIPI
hneg!negDescontoTotalPedido = neg!negDescontoTotalPedido
hneg!negComisRepPedido = neg!negComisRepPedido
hneg!negComisPromotPedido = neg!negComisPromotPedido
hneg!negMotivacao = neg!negMotivacao
hneg!negDataLancamento = neg!negDataLancamento
hneg!negCEFOP = neg!negCEFOP
hneg!negInicioMedicao = neg!negInicioMedicao
hneg!negFinalMedicao = neg!negFinalMedicao

hneg.Update

End Sub
Public Sub Rotina_Processa_Contas_Pagar()
On Error GoTo ErroContasPagar

Call Rotina_AbrirBanco

ctp.Open "select * from contas_a_pagar where ctpstatus = 1", db, 3, 3

If ctp.EOF Then
   Fim = 1
Else
   Fim = 0
End If

Do While Fim = 0

   hctp.Open "select * from historicocontaspagar where chPessoa = ('" & ctp!chPessoa & "') and chNotaFiscal = ('" & ctp!chNotafiscal & "') and chFatura = ('" & ctp!chFatura & "')", db, 3, 3
   If hctp.EOF Then
      hctp.AddNew
   End If
   hctp!chFabricante = ctp!chFabricante
   hctp!chPessoa = ctp!chPessoa
   hctp!chNotafiscal = ctp!chNotafiscal
   hctp!chFatura = ctp!chFatura
   hctp!chDataVencito = ctp!chDataVencito
   'hctp!ctpDataBanco =  ctp!ctpDataBanco
   hctp!ctpDataVencOriginal = ctp!ctpDataVencOriginal
   hctp!ctpDataEmissao = ctp!ctpDataEmissao
   hctp!ctpDataLanc = ctp!ctpDataLanc
   hctp!ctpdescricaooperacao = ctp!ctpdescricaooperacao
   hctp!ctpValorLart = ctp!ctpValorLart
   hctp!ctpValorMerco = ctp!ctpValorMerco
   hctp!ctpValorDaBoleta = ctp!ctpValorDaBoleta
   hctp!chAno = ctp!chAno
   hctp!chMes = ctp!chMes
   hctp!chDia = ctp!chDia
      'hctp!chNumPedido =  ctp!chNumPedido
      'hctp!chNumPedidoComp =  ctp!chNumPedidoComp
   hctp!ctpStatus = ctp!ctpStatus
   hctp!ctpDataProc = ctp!ctpDataProc
   hctp!ctpDataPagamento = ctp!ctpDataPagamento
   hctp!chCodBcoLart = ctp!chCodBcoLart
   hctp!ctpTipoLancamento = ctp!ctpTipoLancamento
   hctp!ctpTipoLancamentoDesc = ctp!ctpTipoLancamentoDesc
   ultimoRegistro = ctp!chPessoa & " - " & ctp!chNotafiscal & " - " & ctp!chFatura
   hctp.Update
   
   hctp.Close
   
   ctp.Delete

   ctp.MoveNext
   If ctp.EOF Then
      Fim = 1
   Else
      Fim = 0
   End If
Loop

Exit Sub

ErroContasPagar:

    MsgBox ultimoRegistro
    Erro = Erro + 1
    MsgBox Err.Description
    
End Sub

Public Sub Rotina_Processa_Contas_Receber()
On Error GoTo ErroContasReceber

Call Rotina_AbrirBanco

hctr.Open "select * from historicocontasreceber", db, 3, 3

ctr.Open "select * from contas_a_receber where ctrstatus = 1", db, 3, 3
If ctr.EOF Then
   Fim = 1
Else
   Fim = 0
End If

Do While Fim = 0

      hctr.AddNew
      hctr!chFabricante = ctr!chFabricante
      hctr!chPessoa = ctr!chPessoa
      hctr!chNotafiscal = ctr!chNotafiscal
      hctr!chFatura = ctr!chFatura
      hctr!ctrDataEmissao = ctr!ctrDataEmissao
      hctr!ctrDataVencito = ctr!ctrDataVencito
      hctr!ctrDataBanco = ctr!ctrDataBanco
      hctr!ctrDataVencOriginal = ctr!ctrDataVencitoOriginal
      hctr!ctrDescricaoOperacao = ctr!ctrDescricaoOperacao
      hctr!ctrValorLart = ctr!ctrValorLart
      hctr!ctrValorMerco = ctr!ctrValorMerco
      hctr!ctrPercentCorrecao = ctr!ctrPercentCorrecao
      hctr!ctrPercentlogistica = ctr!ctrPercentlogistica
      hctr!ctrValorlogistica = ctr!ctrValorlogistica
      hctr!ctrvalorcorrecao = ctr!ctrvalorcorrecao
      hctr!ctrValorDaBoleta = ctr!ctrValorDaBoleta
      hctr!chAno = ctr!chAno
      hctr!chMes = ctr!chMes
      hctr!chDia = ctr!chDia
      hctr!chNumPedido = ctr!chNumPedido
      hctr!chNumPedidoComp = ctr!chNumPedidoComp
      hctr!ctrStatus = ctr!ctrStatus
      hctr!ctrDataRecebimento = ctr!ctrDataRecebimento
      hctr!chCodBcoLart = ctr!chCodBcoLart
      
      hctr!ctrCentroDeCusto = ctr!ctrCentroDeCusto
      hctr!ctrGrupoCentroDeCusto = ctr!ctrGrupoCentroDeCusto
      hctr!ctrSubGrupoCentroDeCusto = ctr!ctrSubGrupoCentroDeCusto
      
      ultimoRegistro = ctr!chPessoa & " - " & ctr!chNotafiscal & " - " & ctr!chFatura
      
      hctr.Update
     
      ctr.Delete
      
      ctr.MoveNext
      If ctr.EOF Then
         Fim = 1
      Else
         Fim = 0
      End If
Loop


Exit Sub

ErroContasReceber:

    MsgBox ultimoRegistro
    Erro = Erro + 1
    MsgBox Err.Description

End Sub


Public Sub Rotina_Nota_Fiscal()

On Error GoTo ErroNotaFiscal

Dim Encontrei As Byte
Dim Fim As Byte
Dim FimNfe As Byte
Dim FimNfd As Byte
Dim FimDet As Byte
Dim ContaDelDesdob As Long
Dim ContaDelDet As Long
Dim ContaDelNF As Long
Dim ContaPendente As Long
Dim WsStatus As Integer

ContaDelDesdob = 0
ContaDelDet = 0
ContaDelNF = 0

'db.begintrans

Call Rotina_AbrirBanco

   
nfe.Open "select * from notafiscalentrada where nfeStatus = 1", db, 3, 3
If nfe.EOF Then
   Fim = 1
   Exit Sub
End If
   
Fim = 0
FimNfe = 0
nfe.MoveFirst

' nfd = Nota Fiscal Desdobramento.

Do While FimNfe = 0
   nfd.Open "select * from notafiscaldesdobramento where chPessoa = ('" & nfe!chPessoa & "') and chNotaFiscalEntrada = ('" & nfe!chNotaFiscalEntrada & "')", db, 3, 3
   If nfd.EOF Then
      MsgBox ("Não há Notas Fiscais de Entrada - Desdobramento") & nfe!chPessoa & " - " & nfe!chNotaFiscalEntrada, vbInformation
      FimNfd = 1
   Else
      FimNfd = 0
      nfd.MoveFirst
      Do While FimNfd = 0
'         If nfd!chPessoa = "VIVO TEL" Then
'            MsgBox ("Chegou"), vbInformation
'         End If
         
         hnfd.Open "Select * from historiconotafiscaldesdobramento where chPessoa = ('" & nfd!chPessoa & "') and chNotaFiscalEntrada = ('" & nfd!chNotaFiscalEntrada & "') and chDataVencimento = ('" & nfd!chDataVencimento & "')", db, 3, 3
         If hnfd.EOF Then
            hnfd.AddNew
         End If
         hnfd!chPessoa = nfd!chPessoa
         hnfd!chNotaFiscalEntrada = nfd!chNotaFiscalEntrada
         hnfd!chDataVencimento = nfd!chDataVencimento
         hnfd!nfdDataVencOriginal = nfd!nfdDataVencOriginal
         hnfd!nfdDataPagamento = nfd!nfdDataPagamento
         hnfd!nfdFaturaNumero = nfd!nfdFaturaNumero
         hnfd!nfdValorDaFatura = nfd!nfdValorDaFatura
         hnfd!nfdStatus = nfd!nfdStatus
         hnfd!nfdStatusPagto = nfd!nfdStatusPagto
         hnfd!nfdOrdemBoleto = nfd!nfdOrdemBoleto
         hnfd!nfdIPTE = nfd!nfdIPTE
          
         ultimoRegistro = "Rotina Grava NF Desdobram. - " & nfd!chPessoa & " - " & nfd!chNotaFiscalEntrada & " - " & nfd!chDataVencimento
          
         hnfd.Update
          
         nfd.Delete
         ContaDelDesdob = ContaDelDesdob + 1
      
         nfd.MoveNext
         
         If nfd.EOF Then
            FimNfd = 1
         End If
         'nfd.Close: Set nfd = Nothing
         hnfd.Close: Set hnfd = Nothing
      Loop
   End If
'   If nfe!chPessoa = "VIVO TEL" Then
'      MsgBox ("Chegou"), vbInformation
'   End If
   
   hnfe.Open "Select * from historiconotafiscalentrada where chPessoa = ('" & nfe!chPessoa & "') and chNotaFiscalEntrada = ('" & nfe!chNotaFiscalEntrada & "')", db, 3, 3
   If hnfe.EOF Then
      hnfe.AddNew
   End If
'   If nfe!chPessoa = "VIVO TEL" Then
'      MsgBox ("Chegou"), vbInformation
'   End If
   hnfe!chPessoa = nfe!chPessoa
   hnfe!chNotaFiscalEntrada = nfe!chNotaFiscalEntrada
   hnfe!nfelartmerco = nfe!nfelartmerco
   hnfe!chCodBcoLart = nfe!chCodBcoLart
   hnfe!nfeDataEmissao = nfe!nfeDataEmissao
   hnfe!nfedataLanc = nfe!nfedataLanc
   hnfe!nfeValorDaNota = nfe!nfeValorDaNota
   hnfe!nfeValorFrete = nfe!nfeValorFrete
   hnfe!nfePagtoFrete = nfe!nfePagtoFrete
   hnfe!nfeValorICMS = nfe!nfeValorICMS
   hnfe!nfeValorIPI = nfe!nfeValorIPI
   hnfe!nfeNF_Boleto = nfe!nfeNF_Boleto
   hnfe!nfeDesdobramento = nfe!nfeDesdobramento
   hnfe!nfeTipoLancamento = nfe!nfeTipoLancamento
   hnfe!nfeStatus = nfe!nfeStatus
        
   ultimoRegistro = "Rotina Grava Nota Fiscal Entrada. - " & nfe!chPessoa & " - " & nfe!chNotaFiscalEntrada & " - " & nfe!nfelartmerco
         
   hnfe.Update
   
   nfe.Delete
   
   ContaDelNF = ContaDelNF + 1

   nfe.MoveNext
   If nfe.EOF Then
      FimNfe = 1
   End If
   If hnfe.State = 1 Then
      hnfe.Close: Set hnfe = Nothing
   End If
   If nfd.State = 1 Then
      nfd.Close: Set nfd = Nothing
   End If
   If dnfe.State = 1 Then
      dnfe.Close: Set dnfe = Nothing
   End If
Loop

'MsgBox ("Total Deletado do Desdobramento = "), , ContaDelDesdob
'MsgBox ("Total Deletado do Detalhe da NF = "), , ContaDelDet
'MsgBox ("Total Deletado da Nota Fiscal = "), , ContaDelNF



 'dnfe = Detalhe Nota Fiscal de Entrada - Produtos.
      

dnfe.Open "select * from notafiscaldetprod", db, 3, 3
If dnfe.EOF Then
   'MsgBox ("Nota Fiscal sem Detalhe."), vbInformation
   FimDet = 1
Else
   FimDet = 0
   NotaFiscalAnterior = Empty
   pessoaAnterior = Empty
   dnfe.MoveFirst
   Do While FimDet = 0
      If Not (dnfe!chPessoa = pessoaAnterior And dnfe!chNotaFiscalEntrada = NotaFiscalAnterior) Then
         pessoaAnterior = dnfe!chPessoa
         NotaFiscalAnterior = dnfe!chNotaFiscalEntrada
         If ctp.State = 1 Then
            ctp.Close: Set ctp = Nothing
         End If
         ctp.Open "Select * from contas_a_pagar where chPessoa = ('" & dnfe!chPessoa & "') and chNotaFiscal = ('" & dnfe!chNotaFiscalEntrada & "')", db, 3, 3
         If ctp.EOF Then
            WsStatus = 0
            Encontrei = 0
         Else
            Encontrei = 1
            WsStatus = ctp!ctpStatus
         End If
      End If
      
      If Encontrei = 1 And WsStatus = 1 Then
     
         If hdnfe.State = 1 Then
            hdnfe.Close: Set hdnfe = Nothing
         End If
   
         hdnfe.Open "Select * from historiconotafiscaldetprod where chPessoa = ('" & dnfe!chPessoa & "') and chNotaFiscalEntrada = ('" & dnfe!chNotaFiscalEntrada & "') and chCodProduto = ('" & dnfe!chCodProduto & "')", db, 3, 3
         If hdnfe.EOF Then
            hdnfe.AddNew
         End If
'         If dnfe!chPessoa = "VIVO TEL" Then
'            MsgBox ("Chegou"), vbInformation
'         End If
         hdnfe!chPessoa = dnfe!chPessoa
         hdnfe!chNotaFiscalEntrada = dnfe!chNotaFiscalEntrada
         hdnfe!chCodProduto = dnfe!chCodProduto
         hdnfe!chProdutoFabrica = dnfe!chProdutoFabrica
         hdnfe!nfdQtd = dnfe!nfdQtd
         hdnfe!nfdPU = dnfe!nfdPU
         hdnfe!nfdValorDaCompra = dnfe!nfdValorDaCompra
         hdnfe!nfdQtdParcelas = dnfe!nfdQtdParcelas
         hdnfe!nfdValorParcela = dnfe!nfdValorParcela
         hdnfe!nfdCentroDeCusto = dnfe!nfdCentroDeCusto
         hdnfe!nfdGrupoCentroDeCusto = dnfe!nfdGrupoCentroDeCusto
         hdnfe!nfdSubGrupoCentroDeCusto = dnfe!nfdSubGrupoCentroDeCusto
         
         If Encontrei = 1 And WsStatus = 1 Then
            hdnfe!nfdDataPagamento = ctp!ctpDataPagamento
         End If
   
         ultimoRegistro = "Rotina Grava Det Produto. - " & dnfe!chPessoa & " - " & dnfe!chNotaFiscalEntrada & " - " & dnfe!chCodProduto & " - " & dnfe!chProdutoFabrica
   
         hdnfe.Update
         
         dnfe.Delete
                  
      End If
      
      dnfe.MoveNext
      If dnfe.EOF Then
         FimDet = 1
      End If
   Loop
End If
         
'db.CommitTrans

Call FechaDB

Exit Sub

ErroNotaFiscal:

    MsgBox ultimoRegistro
    Erro = Erro + 1
    MsgBox Err.Description
End Sub

'Public Sub Rotina_Atualiza_Calendario()
'On Error GoTo ErroCalendario
'tabCalendario.MoveFirst
'tabCalendario.Edit
'tabCalendario("mes1") = tabCalendario("mes2")
'tabCalendario("mes2") = tabCalendario("mes3")
'tabCalendario("mes3") = tabCalendario("mes4")
'tabCalendario("mes4") = tabCalendario("mes5")
'tabCalendario("mes5") = tabCalendario("mes6")
'tabCalendario("mes6") = tabCalendario("mes7")
'tabCalendario("mes7") = tabCalendario("mes8")
'tabCalendario("mes8") = tabCalendario("mes9")
'tabCalendario("mes9") = tabCalendario("mes10")
'tabCalendario("mes10") = tabCalendario("mes11")
'tabCalendario("mes11") = tabCalendario("mes12")
'tabCalendario("mes12") = 1 & "/" & Month(Date) & "/" & Year(Date)
'tabCalendario("mes12") = 1 & "/" & Month(Data_Anterior) & "/" & Year(Data_Anterior)

'UltimoRegistro = "Rotina calendario - " & tabCalendario("mes2")

'tabCalendario.Update

'Exit Sub

'ErroCalendario:

'    MsgBox UltimoRegistro
'    Erro = Erro + 1
'    MsgBox Err.Description
    
'End Sub

'Public Sub Rotina_Gera_Estatistica_regiao()

'On Error GoTo ErroEstatisticaRegiao

'If TabEstatisticaUF.RecordCount > 0 Then
'   TabEstatisticaUF.MoveFirst
'   Do While Not TabEstatisticaUF.EOF
 '     TabEstatisticaUF.Delete
'      TabEstatisticaUF.MoveNext
'   Loop
'End If
'If TabAnualCliente.RecordCount > 0 Then
'   TabAnualCliente.MoveFirst''

'   Do While Not TabAnualCliente.EOF
'      If Not TabAnualCliente("chproduto") = "TOTAL" Then
'         Tabpessoa.Seek "=", TabAnualCliente("chpessoa")
'         If Tabpessoa.NoMatch Then
'            MsgBox ("Cliente não encontrado - "), TabAnualCliente("chpessoa")
'         Else
'            If TabEstatisticaUF.RecordCount > 0 Then
'               TabEstatisticaUF.MoveFirst
'            End If
'            TabEstatisticaUF.Seek "=", Tabpessoa("chuf"), Tabpessoa("pesregiao"), TabAnualCliente("chproduto")
 '           If TabEstatisticaUF.NoMatch Then
 '              TabEstatisticaUF.AddNew
 '              TabEstatisticaUF("uf") = Tabpessoa("chuf")
 '              TabEstatisticaUF("regiao") = Tabpessoa("pesregiao")
 '              TabEstatisticaUF("produto") = TabAnualCliente("chproduto")
 '              TabEstatisticaUF("quantidadenoperiodo") = TabAnualCliente("apsqtdtotal")
 '              TabEstatisticaUF("valornoperiodo") = TabAnualCliente("apsfaturatotal")
 '           Else
 '              TabEstatisticaUF.Edit
 '              TabEstatisticaUF("quantidadenoperiodo") = TabEstatisticaUF("quantidadenoperiodo") + TabAnualCliente("apsqtdtotal")
 '              TabEstatisticaUF("valornoperiodo") = TabEstatisticaUF("valornoperiodo") + TabAnualCliente("apsfaturatotal")
 '           End If
 '
 '           UltimoRegistro = "Rotina de Estatistica regiao - " & Tabpessoa("chuf") & " - " & Tabpessoa("pesregiao") & " - " & TabAnualCliente("chproduto")
 '
 '           TabEstatisticaUF.Update
 '        End If
 '     End If
 '     TabAnualCliente.MoveNext
 ''  Loop
'End If
'Exit Sub

'ErroEstatisticaRegiao:

'    MsgBox UltimoRegistro
'    Erro = Erro + 1
'    MsgBox Err.Description
'MsgBox "Fim da geração de Estatistica UF"
'End Sub
Function CopiarArquivos(Origem As String, destino As String) As Single

Static Buf As String
Dim BTest As Double
Dim FSize As Double
Dim Chunk As Integer
Dim F1 As String
Dim F2 As String
Dim pbCopiaArquivos As Integer

Const bufsize = 1024

pbCopiaArquivos = 0

F1 = FreeFile
Open Origem For Binary As F1
F2 = FreeFile
Open destino For Binary As F2

FSize = LOF(F1)
BTest = FSize - LOF(F2)
i = 1

Do
If BTest < bufsize Then
   Chunk = BTest
Else
   Chunk = bufsize
End If

Buf = String(Chunk, " ")
Get F1, , Buf
Put F2, , Buf
BTest = FSize - LOF(F2)

pbCopiaArquivos = (100 - Int(BTest * 100 / FSize))
'pbCopiaArquivos = (100 - Int(BTest * 100 / FSize))

Loop Until BTest = 0

Close F1
Close F2
CopiarArquivos = FSize

MsgBox "Backups realizados com Sucesso", vbInformation, "Sistema e banco de Dados"

'pbCopiaArquivos.Value = 0
pbCopiaArquivos = 0
Exit Function
End Function

