VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGeraExcelCredito 
   Caption         =   "frmGeraExcelCredito"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   14550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGerarExcel 
      BackColor       =   &H00C0C000&
      Caption         =   "Gerar Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   7575
      Left            =   480
      TabIndex        =   9
      Top             =   840
      Width           =   13815
      Begin MSFlexGridLib.MSFlexGrid grdCtaReceb 
         Height          =   5055
         Left            =   360
         TabIndex        =   14
         Top             =   2400
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   8916
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         FormatString    =   "Cliente                          |Doc.        |Emissão        |Vencimento  |Pago em         |Valor            |Status           |"
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
      Begin VB.ComboBox cmbTipoProcess 
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
         TabIndex        =   0
         Top             =   600
         Width           =   4815
      End
      Begin VB.ComboBox cmbAnoFim 
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
         Left            =   11040
         TabIndex        =   5
         Text            =   "Combo2"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ComboBox cmbMesFim 
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
         Left            =   9240
         TabIndex        =   4
         Text            =   "cmbMesFim"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ComboBox cmbAno 
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
         Left            =   7080
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox cmbMes 
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
         Left            =   5400
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox cmbPessoa 
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
         Text            =   " Todos"
         Top             =   1560
         Width           =   4815
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tipo de Processamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Período Final"
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
         Left            =   9240
         TabIndex        =   12
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Período de Início"
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
         Left            =   5400
         TabIndex        =   11
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   2415
      End
   End
   Begin VB.Label lblTotalEmAtraso 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7800
      TabIndex        =   20
      Top             =   8880
      Width           =   2655
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Total em atraso"
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
      Left            =   7800
      TabIndex        =   19
      Top             =   8520
      Width           =   2535
   End
   Begin VB.Label lblTotalPendente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4080
      TabIndex        =   18
      Top             =   8880
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Total Pendente"
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
      Left            =   4080
      TabIndex        =   17
      Top             =   8520
      Width           =   2775
   End
   Begin VB.Label lblTotalPago 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   840
      TabIndex        =   16
      Top             =   8880
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      TabIndex        =   15
      Top             =   8520
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Gera Planilha de Excel de Contas a Receber"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   8775
   End
End
Attribute VB_Name = "frmGeraExcelCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ind As Integer
Dim ano As Integer
Dim mes As Integer
Dim Dia As Integer
Dim AnoInicioOperacao
Dim DataHoje As Date
Dim AnoHoje As Integer
Dim MesHoje As Integer
Dim DataHojeInvertida As String
Dim DataInicioInvertida As String
Dim DataFimInvertida As String
Dim DataInicioOperacao As Date
Dim DataFimOperacao As Date
Dim DataFinalInvertida As String
Dim MesProximo As Integer
Dim Periodo As Integer
Dim ValorAcumulado As Currency
Dim TotalRecebido As Currency
Dim TotalReceber As Currency
Dim TotalEmAtraso As Currency
Dim Resp As String
Dim Historico As Integer
Dim Status As Integer

Dim TipoDeLancamento As String

Dim dataInicio As String
Dim dataFim As String






Private Sub cmbTipoProcess_LostFocus()

If cmbTipoProcess.ListIndex = 1 Then
   Status = 1
Else
   If cmbTipoProcess.ListIndex = 2 Then
      Status = 0
   End If
End If

End Sub

Private Sub cmdGerarExcel_Click()


'Dim contasreceber As String
Call DeletaTabExcel

ValorAcumulado = 0

If AnoHoje = cmbAno Then
   If cmbMes > MesHoje Then
      MsgBox ("Mês para consulta inválido. Maior que o mês da data atual."), vbInformation
      Exit Sub
   End If
End If

If AnoHoje = cmbAnoFim Then
   If cmbMesFim > MesHoje Then
      MsgBox ("Mês final para consulta inválido. Maior que o mês da data atual."), vbInformation
      Exit Sub
   End If
End If

If cmbAno = cmbAnoFim Then
   If cmbMes > cmbMesFim Then
      MsgBox ("Mês final para consulta inválido. Menor que o mês de início da pesquisa."), vbInformation
      Exit Sub
   End If
End If
 
If cmbAnoFim < cmbAno Then
   MsgBox ("Ano inicio não pode ser menor que o ano fim da pesquisa."), vbInformation
   Exit Sub
End If

Call CriaDatasPesquisa

If Year(DataInicioOperacao) < Year(Date) Then
   Historico = 1
Else
   If Month(DataInicioOperacao) < Month(Date) Then
      Historico = 1
   Else
      Historico = 0
   End If
End If

If Year(DataFimOperacao) < Year(Date) Then
   Periodo = 1
Else
   If Month(DataFimOperacao) < Month(Date) Then
      Periodo = 1
   Else
      Periodo = 0
   End If
End If
Call Rotina_AbrirBanco

If cmbTipoProcess.ListIndex = 3 Then
   If Periodo = 0 Then
      If cmbPessoa = " Todos" Then
         ctr.Open "Select * from contas_a_receber where ctrStatus = ('" & 0 & "') and ctrDataVencito < ('" & DataHojeInvertida & "')", db, 3, 3
         If ctr.EOF Then
            MsgBox ("Sem Contas a Receber para o periodo atual"), vbInformation
         Else
            Call RotinaGaravaReceber
         End If
      Else
         ctr.Open "Select * from contas_a_receber where chPessoa = ('" & cmbPessoa & "') and ctrStatus = ('" & 0 & "') and ctrDataVencito < ('" & DataHojeInvertida & "')", db, 3, 3
         If ctr.EOF Then
            MsgBox ("Sem Contas a Receber para o periodo atual"), vbInformation
         Else
            Call RotinaGaravaReceber
         End If
      End If
   End If
Else
   If cmbTipoProcess.ListIndex = 0 Then
      If Periodo = 0 Then
         If cmbPessoa = " Todos" Then
            ctr.Open "Select * from contas_a_receber", db, 3, 3
         
            If ctr.EOF Then
               MsgBox ("Sem Contas a Receber para o periodo atual"), vbInformation
            Else
               Call RotinaGaravaReceber
            End If
         Else
            ctr.Open "Select * from contas_a_receber where chPessoa = ('" & cmbPessoa & "')", db, 3, 3
            If ctr.EOF Then
               MsgBox ("Sem Contas a Receber para o periodo atual"), vbInformation
            Else
               Call RotinaGaravaReceber
            End If
         End If
      End If
   Else
      If Periodo = 0 Then
         If cmbPessoa = " Todos" Then
            ctr.Open "Select * from contas_a_receber where ctrStatus = ('" & Status & "')", db, 3, 3
         
            If ctr.EOF Then
               MsgBox ("Sem Contas a Receber para o periodo atual"), vbInformation
            Else
               Call RotinaGaravaReceber
            End If
         Else
            ctr.Open "Select * from contas_a_receber where chPessoa = ('" & cmbPessoa & "') and ctrStatus = ('" & Status & "')", db, 3, 3
      
            If ctr.EOF Then
               MsgBox ("Sem Contas a Receber para o periodo atual"), vbInformation
            Else
               Call RotinaGaravaReceber
            End If
         End If
      End If
   End If

   If cmbTipoProcess.ListIndex = 0 Then
      If cmbPessoa = " Todos" Then
         If Not Historico = 0 Then
            If ctr.State = 1 Then
               ctr.Close: Set ctr = Nothing
            End If
         
            ctr.Open "Select * from historicocontasreceber where ctrDataRecebimento > ('" & DataInicioInvertida & "') and ctrDataRecebimento < ('" & DataFinalInvertida & "')", db, 3, 3
            If ctr.EOF Then
               MsgBox ("Sem Contas a Receber para o periodo anterior solicitado"), vbInformation
            Else
               Call RotinaGaravaReceber
            End If
         End If
      Else
         If Not Historico = 0 Then
            If ctr.State = 1 Then
               ctr.Close: Set ctr = Nothing
            End If
            ctr.Open "Select * from historicocontasreceber where chPessoa = ('" & cmbPessoa & "') and ctrDataRecebimento > ('" & DataInicioInvertida & "') and ctrDataRecebimento < ('" & DataFinalInvertida & "')", db, 3, 3
         
            If ctr.EOF Then
               MsgBox ("Sem Contas a Receber para o periodo anterior solicitado"), vbInformation
            Else
               Call RotinaGaravaReceber
            End If
         End If
      End If
   Else
      If cmbPessoa = " Todos" Then
         If Not Historico = 0 Then
            If ctr.State = 1 Then
               ctr.Close: Set ctr = Nothing
            End If
         
            ctr.Open "Select * from historicocontasreceber where ctrStatus = ('" & 1 & "') and ctrDataRecebimento > ('" & DataInicioInvertida & "') and ctrDataRecebimento < ('" & DataFinalInvertida & "')", db, 3, 3
         
            If ctr.EOF Then
               MsgBox ("Sem Contas a Receber para o periodo anterior solicitado"), vbInformation
            Else
               Call RotinaGaravaReceber
            End If
         End If
      Else
         If Not Periodo = 0 Then
            If ctr.State = 1 Then
               ctr.Close: Set ctr = Nothing
            End If
            ctr.Open "Select * from historicocontasreceber where chPessoa = ('" & cmbPessoa & "') and ctrStatus = ('" & 1 & "') and ctrDataRecebimento > ('" & DataInicioInvertida & "') and ctrDataRecebimento < ('" & DataFinalInvertida & "')", db, 3, 3
         
            If ctr.EOF Then
               MsgBox ("Sem Contas a Receber para o periodo anterior solicitado"), vbInformation
            Else
               Call RotinaGaravaReceber
            End If
         End If
      End If
   End If
End If
'MsgBox ("Valor Total Gerado = ") & Format$(ValorAcumulado, "##,##0.00")

'Call FechaDB

If ctr.State = 1 Then
   ctr.Close: Set ctr = Nothing
End If

ctr.Open "Select * from contasreceber", db, 3, 3
If Not ctr.EOF Then

   TotalRecebido = 0
   TotalReceber = 0
   TotalEmAtraso = 0
      
   Ind = 1
   grdCtaReceb.Rows = 1
   
   ctr.MoveFirst
   
   Do While Not ctr.EOF
   
      grdCtaReceb.Rows = grdCtaReceb.Rows + 1
      grdCtaReceb.TextMatrix(Ind, 0) = ctr!chPessoa
      grdCtaReceb.TextMatrix(Ind, 1) = ctr!chNotafiscal
      grdCtaReceb.TextMatrix(Ind, 2) = ctr!ctrDataEmissao
      grdCtaReceb.TextMatrix(Ind, 3) = ctr!ctrDataVencito
      
      If Not IsNull(ctr!ctrDataRecebimento) Then
         grdCtaReceb.TextMatrix(Ind, 4) = ctr!ctrDataRecebimento
      Else
         grdCtaReceb.TextMatrix(Ind, 4) = Empty
      End If
      grdCtaReceb.TextMatrix(Ind, 5) = Format$(ctr!ctrValorDaBoleta, "##,##0.00")

      If ctr!ctrStatus = 1 Then
         grdCtaReceb.TextMatrix(Ind, 6) = "PAGO"
         TotalRecebido = TotalRecebido + ctr!ctrValorDaBoleta
      Else
         If ctr!ctrDataVencito < Date Then
            grdCtaReceb.TextMatrix(Ind, 6) = "EM ATRASO"
            TotalEmAtraso = TotalEmAtraso + ctr!ctrValorDaBoleta
         Else
            grdCtaReceb.TextMatrix(Ind, 6) = "PENDENTE"
            TotalReceber = TotalReceber + ctr!ctrValorDaBoleta
         End If
      End If
      
      Dia = Day(ctr!ctrDataVencito)
      mes = Month(ctr!ctrDataVencito)
      ano = Year(ctr!ctrDataVencito)

      grdCtaReceb.TextMatrix(Ind, 7) = Format$(ano & "-" & mes & "-" & Dia, "yyyy-mm-dd")

      Ind = Ind + 1
      ctr.MoveNext
   Loop
       
   grdCtaReceb.Row = 1
   grdCtaReceb.RowSel = Ind - 1
   grdCtaReceb.Col = 7
   grdCtaReceb.ColSel = 7
   grdCtaReceb.Sort = 6
   
   lblTotalPago = Format$(TotalRecebido, "##,##0.00")
   lblTotalPendente = Format$(TotalReceber, "##,##0.00")
   lblTotalEmAtraso = Format$(TotalEmAtraso, "##,##0.00")

   Resp = MsgBox("Deseja gerar Planilha de Excel com essas Informações???", vbExclamation + vbYesNo)

   If Resp = vbYes Then
      Call ExportarCtaReceber
   End If
End If

End Sub

Public Sub RotinaGaravaReceber()

ctr.MoveFirst

Do While Not ctr.EOF

   If contab.State = 1 Then
      contab.Close: Set contab = Nothing
      acContab = 0
   End If
   
   contab.Open "Select * from contasreceber where chPessoa = ('" & ctr!chPessoa & "') and chNotaFiscal = ('" & ctr!chNotafiscal & "') and chFatura = ('" & ctr!chFatura & "')", db, 3, 3
   If contab.EOF Then

      contab.AddNew
   
      contab!chPessoa = ctr!chPessoa
      contab!chNotafiscal = ctr!chNotafiscal
      contab!chFatura = ctr!chFatura
      contab!ctrDataEmissao = ctr!ctrDataEmissao
      contab!ctrDataVencito = ctr!ctrDataVencito
      contab!ctrDescricaoOperacao = ctr!ctrDescricaoOperacao
      contab!ctrValorDaBoleta = ctr!ctrValorDaBoleta
      contab!chNumPedido = ctr!chNumPedido
      contab!chNumPedidoComp = ctr!chNumPedidoComp
      contab!ctrStatus = ctr!ctrStatus
      If Not IsNull(ctr!ctrDataRecebimento) Then
         contab!ctrDataRecebimento = ctr!ctrDataRecebimento
      End If
      contab!ctrValorOriginal = ctr!ctrValorLart
      contab.Update
      
      ValorAcumulado = ValorAcumulado + ctr!ctrValorDaBoleta
      
   End If
   
   ctr.MoveNext
Loop

End Sub


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

Periodo = 0
'lblHoje = Date

DataHojeInvertida = Format$(Date, "yyyy-mm-dd")

TipoDeLancamento = "RECEITA"

AnoHoje = Year(Date)
MesHoje = Month(Date)
'lblHoje = Date

AnoInicioOperacao = 2019

ano = Year(Date)

For Ind = 1 To 12
    cmbMes.AddItem Format$(Ind, "00")
    cmbMesFim.AddItem Format$(Ind, "00")
Next

cmbMes.ListIndex = 0
cmbMesFim.ListIndex = 0

Do While (ano + 1) > AnoInicioOperacao
   cmbAno.AddItem ano
   cmbAnoFim.AddItem ano
   ano = ano - 1
Loop

cmbAno.ListIndex = 0
cmbAnoFim.ListIndex = 0

Call Rotina_AbrirBanco

pes.Open "Select * from pessoa where pesTipoPessoa = ('" & 0 & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("ERRO; Acesso a pessoa sem Clientes"), vbInformation
   Call FechaDB
   Exit Sub
End If

pes.MoveFirst

cmbPessoa.AddItem " Todos"
Do While Not pes.EOF
   cmbPessoa.AddItem pes!chPessoa
   pes.MoveNext
Loop

cmbTipoProcess.AddItem "Contas Recebidas e a Receber"
cmbTipoProcess.AddItem "Contas Recebidas"
cmbTipoProcess.AddItem "Contas a Receber"
cmbTipoProcess.AddItem "Contas em Atraso"

End Sub

Public Sub CriaDatasPesquisa()

mes = Format$(cmbMes, "00")
ano = cmbAno
Dia = Format$(1, "00")

DataHoje = Format$(Dia, "00") & "/" & Format$(mes, "00") & "/" & ano
DataInicioOperacao = DataHoje
DataHoje = DataHoje - 1
Dia = Day(DataHoje)
mes = Month(DataHoje)
ano = Year(DataHoje)

DataInicioInvertida = Format$(ano & "-" & mes & "-" & Dia, "yyyy-mm-dd")

DataHoje = DataHoje + 1
MesProximo = Month(DataHoje)

Do While Month(DataHoje) = MesProximo
   DataHoje = DataHoje + 1
   'MesProximo = Format$(Month(DataHoje), "00")
Loop

mes = Format$(cmbMesFim, "00")
ano = cmbAnoFim
Dia = Format$(1, "00")

DataHoje = Format$(Dia, "00") & "/" & Format$(mes, "00") & "/" & ano
DataFimOperacao = DataHoje
DataHoje = DataHoje - 1
Dia = Day(DataHoje)
mes = Month(DataHoje)
ano = Year(DataHoje)

DataFimInvertida = Format$(ano & "-" & mes & "-" & Dia, "yyyy-mm-dd")

DataHoje = DataHoje + 1
MesProximo = Month(DataHoje)

Do While Month(DataHoje) = MesProximo
   DataHoje = DataHoje + 1
   'MesProximo = Format$(Month(DataHoje), "00")
Loop

DataFinalInvertida = Format$(DataHoje, "yyyy-mm-dd")
DataHoje = Date

End Sub

Public Sub DeletaTabExcel()

Call Rotina_AbrirBanco

contab.Open "Select * from contasreceber", db, 3, 3
If Not contab.EOF Then

   contab.MoveFirst
   
   Do While Not contab.EOF
      contab.Delete
      contab.MoveNext
   Loop
End If

Call FechaDB

End Sub




