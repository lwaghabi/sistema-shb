VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSerieHistoricaMedicao 
   Caption         =   "frmSerieHistroricaMedicao"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17865
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
   ScaleHeight     =   9840
   ScaleWidth      =   17865
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGeralFaturado 
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
      Height          =   480
      Left            =   13320
      TabIndex        =   15
      Top             =   1710
      Width           =   1935
   End
   Begin VB.TextBox txtHoje 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15360
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Height          =   7695
      Left            =   0
      TabIndex        =   11
      Top             =   2160
      Width           =   17775
      Begin MSFlexGridLib.MSFlexGrid GridConsulta 
         Height          =   7215
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   17535
         _ExtentX        =   30930
         _ExtentY        =   12726
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         FormatString    =   $"frmSerieHistoricaMedicao.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   12375
      Begin VB.CommandButton cmdConsulta 
         BackColor       =   &H0000FF00&
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
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
         Height          =   615
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1335
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
         Left            =   240
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   600
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker dtInicioPesquisa 
         Height          =   495
         Left            =   3840
         TabIndex        =   1
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
         Format          =   117571585
         CurrentDate     =   44750
      End
      Begin MSComCtl2.DTPicker dtFinalPesquisa 
         Height          =   495
         Left            =   6840
         TabIndex        =   2
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
         Format          =   117309441
         CurrentDate     =   44750
      End
      Begin VB.Label Label5 
         Caption         =   "Até"
         Height          =   375
         Left            =   7440
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "De"
         Height          =   375
         Left            =   4320
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Período"
         Height          =   375
         Left            =   5760
         TabIndex        =   8
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Geral Faturado"
      Height          =   375
      Left            =   13320
      TabIndex        =   16
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label6 
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
      Left            =   15360
      TabIndex        =   13
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Consulta de Faturamento Médio por Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "frmSerieHistoricaMedicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IndCol As Integer
Dim IndLin As Integer
Dim DataInicioInvertida As String
Dim DataFinalInvertida As String
Dim InicioPesquisa As Date
Dim FinalPesquisa As Date
Dim ano As Integer
Dim mes As Integer
Dim Dia As Integer
Dim ClienteAnterior As String
Dim MedicaoAnterior As String
Dim Verifica As String
Dim TipoProduto As String
Dim AcumulaQtdMedicao As Integer
Dim DataInicioOperacao As Date
Dim DataFinalOperacao As Date


'Acumuladores de Valores e Medias

Dim AcumulaHBT As Currency
Dim AcumulaJBX As Currency
Dim AcumulaServico As Currency
Dim ValorTotalPeriodo As Currency
Dim MediaPeriodo As Currency
Dim ValorTotalMedio As Currency
Dim TotalGeral As Currency
Dim MediaGeral As Currency


Private Sub cmbCliente_LostFocus()
ClienteAnterior = Empty
End Sub

Private Sub cmdConsulta_Click()

InicioPesquisa = dtInicioPesquisa - 1

ClienteAnterior = Empty

ano = Year(InicioPesquisa)
mes = Month(InicioPesquisa)
Dia = Day(InicioPesquisa)
DataInicioInvertida = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")

FinalPesquisa = dtFinalPesquisa + 1

ano = Year(FinalPesquisa)
mes = Month(FinalPesquisa)
Dia = Day(FinalPesquisa)
DataFinalInvertida = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")

Call Rotina_AbrirBanco

If cmbCliente = " TODOS" Then
   neg.Open "Select * from historiconegociacao where negInicioMedicao > ('" & DataInicioInvertida & "') and negFinalMedicao < ('" & DataFinalInvertida & "')", db, 3, 3
   If neg.EOF Then
      MsgBox ("Parâmetros inválidos para consulta. Verificar datas."), vbInformation
      Call FechaDB
      Exit Sub
   End If
Else
   neg.Open "Select * from historiconegociacao where chPessoa = ('" & cmbCliente & "') and negInicioMedicao > ('" & DataInicioInvertida & "') and negFinalMedicao < ('" & DataFinalInvertida & "')", db, 3, 3
   If neg.EOF Then
      MsgBox ("Parâmetros inválidos para consulta. Verificar datas informadas."), vbInformation
      Call FechaDB
      Exit Sub
   End If
End If

Call CargaGridConsulta

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtHoje = Date
dtInicioPesquisa = Date
dtFinalPesquisa = Date

Call Rotina_AbrirBanco

pes.Open "Select * from pessoa where pesTipoPessoa = ('" & 0 & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Cadastro pessoa vazio. Comunicar ao analista responsável"), vbCritical
   Call FechaDB
   Exit Sub
End If

pes.MoveFirst

cmbCliente.Clear

cmbCliente.AddItem " TODOS"

Do While Not pes.EOF
   cmbCliente.AddItem pes!chPessoa
   pes.MoveNext
Loop

End Sub

Public Sub CargaGridConsulta()

neg.MoveFirst

TotalGeral = 0
MediaGeral = 0

DataInicioOperacao = neg!negInicioMedicao
DataFinalOperacao = neg!negFinalMedicao
   
IndLin = 0
IndCol = 0

GridConsulta.Rows = 2

Do While Not neg.EOF

   If dneg.State = 1 Then
      dneg.Close: Set dneg = Nothing
   End If
   
   dneg.Open "Select * from historicodetalhenegociacao where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
   If dneg.EOF Then
      MsgBox ("Erro no acesso a Detalhe de Negociacao. Comunicar ao analista responsável."), vbInformation
      Call FechaDB
      Exit Sub
   End If

   dneg.MoveFirst

   Do While Not dneg.EOF
      
      If ClienteAnterior = Empty Then
         ClienteAnterior = dneg!chPessoa
         MedicaoAnterior = dneg!chNumPedido
         AcumulaHBT = 0
         AcumulaJBX = 0
         AcumulaServico = 0
         AcumulaQtdMedicao = 1
         GridConsulta.Rows = 2
      End If
      
      Call VerificaTipoProduto
      
      If TipoProduto = "HBT" Then
         AcumulaHBT = AcumulaHBT + dneg!hdnValorDaOperacao
      Else
         If TipoProduto = "JBX" Then
            AcumulaJBX = AcumulaJBX + dneg!hdnValorDaOperacao
         Else
            AcumulaServico = AcumulaServico + dneg!hdnValorDaOperacao
         End If
      End If
      
      If Not dneg!chNumPedido = MedicaoAnterior Then
         MedicaoAnterior = dneg!chNumPedido
         AcumulaQtdMedicao = AcumulaQtdMedicao + 1
      End If
      
      
      dneg.MoveNext
   
   Loop
   
   neg.MoveNext
   
   If Not neg.EOF Then
      If DataInicioOperacao > neg!negInicioMedicao Then
         DataInicioOperacao = neg!negInicioMedicao
      End If
   
      If DataFinalOperacao < neg!negFinalMedicao Then
         DataFinalOperacao = neg!negFinalMedicao
      End If
      
      If Not neg!chPessoa = ClienteAnterior Then
         Call QuebraCliente
         ClienteAnterior = neg!chPessoa
         AcumulaHBT = 0
         AcumulaJBX = 0
         AcumulaServico = 0
         AcumulaQtdMedicao = 0
         DataInicioOperacao = neg!negInicioMedicao
         DataFinalOperacao = neg!negFinalMedicao
      End If
      
      
   End If
Loop

Call QuebraCliente

txtGeralFaturado = Format$(TotalGeral, "##,#00.00")

End Sub

Public Sub QuebraCliente()

ValorTotalPeriodo = AcumulaJBX + AcumulaHBT + AcumulaServico
ValorTotalMedio = ValorTotalPeriodo / AcumulaQtdMedicao
         
IndLin = IndLin + 1

      GridConsulta.Rows = IndLin + 1
      GridConsulta.TextMatrix(IndLin, 1) = ClienteAnterior
      GridConsulta.TextMatrix(IndLin, 2) = DataInicioOperacao
      GridConsulta.TextMatrix(IndLin, 3) = DataFinalOperacao
      GridConsulta.TextMatrix(IndLin, 4) = AcumulaQtdMedicao
      GridConsulta.TextMatrix(IndLin, 5) = "JBX"
      GridConsulta.TextMatrix(IndLin, 6) = Format$(AcumulaJBX, "##,##0.00")
      MediaPeriodo = AcumulaJBX / AcumulaQtdMedicao
      GridConsulta.TextMatrix(IndLin, 7) = Format$(MediaPeriodo, "##,##0.00")
      GridConsulta.TextMatrix(IndLin, 8) = Empty
      GridConsulta.TextMatrix(IndLin, 9) = Empty
      
      IndLin = IndLin + 1

      GridConsulta.Rows = IndLin + 1
      GridConsulta.TextMatrix(IndLin, 1) = Empty
      GridConsulta.TextMatrix(IndLin, 2) = Empty
      GridConsulta.TextMatrix(IndLin, 3) = Empty
      GridConsulta.TextMatrix(IndLin, 4) = Empty
      GridConsulta.TextMatrix(IndLin, 5) = "HBT"
      GridConsulta.TextMatrix(IndLin, 6) = Format$(AcumulaHBT, "##,##0.00")
      MediaPeriodo = AcumulaHBT / AcumulaQtdMedicao
      GridConsulta.TextMatrix(IndLin, 7) = Format$(MediaPeriodo, "##,##0.00")
      GridConsulta.TextMatrix(IndLin, 8) = Empty
      GridConsulta.TextMatrix(IndLin, 9) = Empty
      
      IndLin = IndLin + 1

      GridConsulta.Rows = IndLin + 1
      GridConsulta.TextMatrix(IndLin, 1) = Empty
      GridConsulta.TextMatrix(IndLin, 2) = Empty
      GridConsulta.TextMatrix(IndLin, 3) = Empty
      GridConsulta.TextMatrix(IndLin, 4) = Empty
      GridConsulta.TextMatrix(IndLin, 5) = "SERVIÇO"
      GridConsulta.TextMatrix(IndLin, 6) = Format$(AcumulaServico, "##,##0.00")
      MediaPeriodo = AcumulaServico / AcumulaQtdMedicao
      GridConsulta.TextMatrix(IndLin, 7) = Format$(MediaPeriodo, "##,##0.00")
      GridConsulta.TextMatrix(IndLin, 8) = Format$(ValorTotalPeriodo, "##,##0.00")
      ValorTotalMedio = ValorTotalPeriodo / AcumulaQtdMedicao
      GridConsulta.TextMatrix(IndLin, 9) = Format$(ValorTotalMedio, "##,##0.00")
                  
      TotalGeral = TotalGeral + ValorTotalPeriodo
      
      IndLin = IndLin + 1
      
End Sub

Public Sub VerificaTipoProduto()

Verifica = Mid$(dneg!chProduto, 7, 2)
      
If Verifica = "JB" Then
   TipoProduto = "JBX"
Else
   Verifica = Mid$(dneg!chProduto, 1, 3)
   If Verifica = "HBT" Then
      TipoProduto = "HBT"
   Else
      If Verifica = "JBX" Then
         TipoProduto = "JBX"
      Else
         TipoProduto = "SERVICO"
      End If
   End If
End If
End Sub
