VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLocacaoeServicoNaoFaturado 
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16275
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   16275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sair"
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
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid grdNaoProc 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   9551
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      FormatString    =   "Cliente                     |Medição    |Comp|Início Medição    |Final Medição    |Valor            |Status Processamento"
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   3
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lblAcumulaGeral 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   10560
      TabIndex        =   2
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Locações e Serviços prestados não faturados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "frmLocacaoeServicoNaoFaturado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ind As Integer
Dim AcumulaValor As Currency
Dim AcumulaGeral As Currency
Dim DataHoje As String


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

AcumulaValor = 0
AcumulaGeral = 0
DataHoje = Format$(Date, "yyyy-mm-dd")

Call Rotina_AbrirBanco

neg.Open "Select * from negociacao where not negStatus = ('" & 1 & "') and negFinalMedicao < ('" & DataHoje & "')", db, 3, 3
If neg.EOF Then
   MsgBox ("Não há Locações e Serviços sem processamento."), vbInformation
   Call FechaDB
   Exit Sub
End If

grdNaoProc.Rows = 1
Ind = 1

neg.MoveFirst

Do While Not neg.EOF

   If dneg.State = 1 Then
      dneg.Close: Set dneg = Nothing
   End If

   dneg.Open "Select * from detalhenegociacao where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
   If dneg.EOF Then
      MsgBox ("ERRO: Comunicar ao analista responsável."), vbCritical
      Call FechaDB
      Exit Sub
   End If
   
   dneg.MoveFirst
   
   Do While Not dneg.EOF
      AcumulaValor = AcumulaValor + dneg!pedValorDaOperacao
      dneg.MoveNext
   Loop
   
   grdNaoProc.Rows = Ind + 1
   grdNaoProc.TextMatrix(Ind, 0) = neg!chPessoa
   grdNaoProc.TextMatrix(Ind, 1) = neg!chNumPedido
   grdNaoProc.TextMatrix(Ind, 2) = neg!chNumPedidoComp
   grdNaoProc.TextMatrix(Ind, 3) = neg!negInicioMedicao
   grdNaoProc.TextMatrix(Ind, 4) = neg!negFinalMedicao
   grdNaoProc.TextMatrix(Ind, 5) = Format$(AcumulaValor, "##,###,#00.00")
   If neg!negStatus = 0 Then
      grdNaoProc.TextMatrix(Ind, 6) = "NÃO PROCESSADO"
   Else
      grdNaoProc.TextMatrix(Ind, 6) = "EM APROVAÇÃO"
   End If
      
   AcumulaGeral = AcumulaGeral + AcumulaValor
   
   AcumulaValor = 0
   
   Ind = Ind + 1
   neg.MoveNext
Loop

lblAcumulaGeral = Format$(AcumulaGeral, "##,##0.00")

End Sub


