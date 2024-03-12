VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFaturaLocacaoEsp 
   Caption         =   "Fatura Locação Especial"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   15285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Referência"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7560
      TabIndex        =   30
      Top             =   3240
      Width           =   7335
      Begin VB.TextBox txtContratoComp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   34
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox txtContrato 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   32
         Top             =   120
         Width           =   3735
      End
      Begin VB.Label Label5 
         Caption         =   "Complemento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox txtTaxa 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   5
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13080
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtComplemento 
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
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtMedicao 
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
      Left            =   10440
      TabIndex        =   8
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtUnidadeOperacional 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.TextBox txtRazaoSocial 
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
      Left            =   7560
      TabIndex        =   6
      Top             =   2760
      Width           =   7335
   End
   Begin VB.TextBox txtPessoa 
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
      Left            =   7560
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Frame Frame 
      Caption         =   " Fatura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   7440
      TabIndex        =   3
      Top             =   6000
      Width           =   3615
      Begin VB.TextBox txtNumSerie 
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
         Left            =   2280
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtNumFatura 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
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
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Série"
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
         Left            =   2280
         TabIndex        =   28
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Número"
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
         Left            =   960
         TabIndex        =   27
         Top             =   150
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdFatura 
      BackColor       =   &H00FFFF80&
      Caption         =   "Fatura"
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H008080FF&
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox dtHoje 
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
      Left            =   12480
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtEmis 
      Height          =   495
      Left            =   7680
      TabIndex        =   0
      Top             =   5400
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   392822785
      CurrentDate     =   44328
   End
   Begin MSFlexGridLib.MSFlexGrid GridFatura 
      Height          =   6135
      Left            =   0
      TabIndex        =   12
      Top             =   840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   10821
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColor       =   16777152
      ForeColor       =   0
      BackColorFixed  =   16776960
      ForeColorFixed  =   0
      ForeColorSel    =   0
      BackColorBkg    =   16777152
      FormatString    =   "Nº Fatura|Medição||Cliente                     ||||||"
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
   Begin MSComCtl2.DTPicker dtFim 
      Height          =   495
      Left            =   12720
      TabIndex        =   13
      Top             =   5400
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   392822785
      CurrentDate     =   44298
   End
   Begin MSComCtl2.DTPicker dtInicio 
      Height          =   495
      Left            =   10440
      TabIndex        =   14
      Top             =   5400
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   392822785
      CurrentDate     =   44298
   End
   Begin VB.Label Label1 
      Caption         =   "Taxa p/envio"
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
      Left            =   13080
      TabIndex        =   26
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "Fatura de Locação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   25
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Período"
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
      Index           =   7
      Left            =   10440
      TabIndex        =   24
      Top             =   4560
      Width           =   4335
   End
   Begin VB.Label Label 
      Caption         =   "Razão Soicial Cliente"
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
      Index           =   6
      Left            =   7560
      TabIndex        =   23
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label 
      Caption         =   "Unidade Operacional"
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
      Index           =   5
      Left            =   7560
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label 
      Caption         =   "Cliente"
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
      Index           =   4
      Left            =   7560
      TabIndex        =   21
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label 
      Caption         =   "Complemento"
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
      Index           =   3
      Left            =   7560
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "Medição"
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
      Index           =   2
      Left            =   10440
      TabIndex        =   19
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label 
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
      Height          =   495
      Index           =   1
      Left            =   12480
      TabIndex        =   18
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "De"
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
      Index           =   8
      Left            =   10440
      TabIndex        =   17
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "Até"
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
      Index           =   9
      Left            =   12840
      TabIndex        =   16
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Data Emissão"
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
      Index           =   10
      Left            =   7680
      TabIndex        =   15
      Top             =   5040
      Width           =   2055
   End
End
Attribute VB_Name = "frmFaturaLocacaoEsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim Rel As Object
Dim Relatorio As String
Dim txtNome As String
Dim txtNumPedido As String
Dim txtPedidoComp As String
Dim UnidadeOperaional As String
Dim TipoLocacao As Integer
Dim dataInicio As Date
Dim dataFim As Date
Dim DataVenc As Date
Dim TipoProduto As Byte
Dim MedicaoAnter As String
Dim Item As Integer
Dim Resp As String
Dim NumPedido As Integer
Dim FaturaNoMes As Byte


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Ind As Integer

FaturaNoMes = 0

dtInicio = Date
dtFim = Date
dtEmis = Date
Relatorio = "drMedicao"
dtHoje = Date
MedicaoAnter = Empty

GridFatura.Rows = 2
GridFatura.TextMatrix(1, 0) = Empty
GridFatura.TextMatrix(1, 1) = Empty
GridFatura.TextMatrix(1, 2) = Empty
GridFatura.TextMatrix(1, 3) = Empty
GridFatura.TextMatrix(1, 4) = Empty
GridFatura.TextMatrix(1, 5) = Empty
GridFatura.TextMatrix(1, 6) = Empty
GridFatura.TextMatrix(1, 7) = Empty
GridFatura.TextMatrix(1, 8) = Empty
GridFatura.TextMatrix(1, 9) = Empty

Call Rotina_AbrirBanco

'Carrega Negociação

neg.Open "Select * from negociacao where negStatus = ('" & 1 & "')", db, 3, 3
If neg.EOF Then
   'MsgBox ("Não há Medição para Fatura até a presente data"), vbInformation
   'Call FechaDB
   FaturaNoMes = 0
Else
   FaturaNoMes = 1
End If
TipoProduto = 0
Ind = 0

If FaturaNoMes = 1 Then
   neg.MoveFirst
   Do While Not neg.EOF
      If Not neg!chNumPedido = MedicaoAnter Then
         MedicaoAnter = neg!chNumPedido
         If ctr.State = 1 Then
            ctr.Close: Set ctr = Nothing
         End If
         ctr.Open "Select * from contas_a_receber where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
         If ctr.EOF Then
            ctr.Close: Set ctr = Nothing
            Exit Sub
         End If
         If dneg.State = 1 Then
            dneg.Close: Set dneg = Nothing
         End If
         dneg.Open "Select * from detalhenegociacao where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
         If dneg.EOF Then
            dneg.Close: Set dneg = Nothing
            Exit Sub
         End If
            
         Ind = Ind + 1
         GridFatura.Rows = Ind + 1
         GridFatura.TextMatrix(Ind, 0) = neg!negNumFatura
         GridFatura.TextMatrix(Ind, 1) = neg!chNumPedido
         GridFatura.TextMatrix(Ind, 2) = neg!chNumPedidoComp
         GridFatura.TextMatrix(Ind, 3) = neg!chPessoa
         GridFatura.TextMatrix(Ind, 4) = neg!chUnidadeOperacional
         GridFatura.TextMatrix(Ind, 5) = ctr!ctrDataVencito
         GridFatura.TextMatrix(Ind, 6) = dneg!chDataInicio
         GridFatura.TextMatrix(Ind, 7) = dneg!chDataFim
         GridFatura.TextMatrix(Ind, 8) = neg!negSerieFatura
         GridFatura.TextMatrix(Ind, 9) = neg!negDataEmissaoFatura
      End If
         
      neg.MoveNext
   
   Loop
End If

'Carrega Historico de Negociação

If neg.State = 1 Then
   neg.Close: Set neg = Nothing
End If

neg.Open "Select * from historiconegociacao where hnegTipoProduto = ('" & 0 & "') and not (negSerieFatura = ('" & "E" & "'))", db, 3, 3
If neg.EOF Then
   MsgBox ("Não há Medição para Fatura até a presente data"), vbInformation
   Call FechaDB
   Exit Sub
End If
TipoProduto = 0
'Ind = 0
neg.MoveFirst
Do While Not neg.EOF
   If Not neg!chNumPedido = MedicaoAnter Then
      MedicaoAnter = neg!chNumPedido
      If ctr.State = 1 Then
         ctr.Close: Set ctr = Nothing
      End If
      ctr.Open "Select * from contas_a_receber where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
      If ctr.EOF Then
         If ctr.State = 1 Then
            ctr.Close: Set ctr = Nothing
         End If
         ctr.Open "Select * from historicocontasreceber where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
         If ctr.EOF Then
            ctr.Close: Set ctr = Nothing
            Exit Sub
         End If
      End If
      If dneg.State = 1 Then
         dneg.Close: Set dneg = Nothing
      End If
      dneg.Open "Select * from historicodetalhenegociacao where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
      If dneg.EOF Then
         dneg.Close: Set dneg = Nothing
         Exit Sub
      End If
         
      Ind = Ind + 1
      GridFatura.Rows = Ind + 1
      GridFatura.TextMatrix(Ind, 0) = neg!negNumFatura
      GridFatura.TextMatrix(Ind, 1) = neg!chNumPedido
      GridFatura.TextMatrix(Ind, 2) = neg!chNumPedidoComp
      GridFatura.TextMatrix(Ind, 3) = neg!chPessoa
      GridFatura.TextMatrix(Ind, 4) = neg!chUnidadeOperacional
      GridFatura.TextMatrix(Ind, 5) = ctr!ctrDataVencito
      GridFatura.TextMatrix(Ind, 6) = dneg!chDataInicio
      GridFatura.TextMatrix(Ind, 7) = dneg!chDataFim
      If Not IsNull(neg!negSerieFatura) Then
         GridFatura.TextMatrix(Ind, 8) = neg!negSerieFatura
      Else
         GridFatura.TextMatrix(Ind, 8) = Empty
      End If
      If Not IsNull(neg!negDataEmissaoFatura) Then
         GridFatura.TextMatrix(Ind, 9) = neg!negDataEmissaoFatura
      Else
         GridFatura.TextMatrix(Ind, 9) = Date
      End If
      
   End If
      
   neg.MoveNext

Loop

GridFatura.ColSel = 0
GridFatura.Sort = 1

Call FechaDB
   
End Sub
Private Sub cmdFatura_Click()
If dtInicio = dtFim Then
   MsgBox ("Ajustar o período de Medição"), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If
If txtNumFatura = Empty Then
   MsgBox ("Número da fatura não informado."), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If


Call Rotina_AbrirBanco
Relatorio = "drFatura"
'db.begintrans
gge.Open "Select * from geradorgeral where chAlfaNumerica = ('" & Relatorio & "')", db, 3, 3
If gge.EOF Then
   gge.AddNew
End If
If txtTaxa = "" Then
   MsgBox ("Para impressão é preciso informar a taxa"), vbInformation
   Exit Sub
End If

gge!chAlfaNumerica = "drFatura"
gge!ggeDataHoje = dtEmis
gge!ggeDataIni = dtInicio
gge!chNumerica = Format$(dtInicio, "yyyymmdd")
gge!ggeDataFim = dtFim
gge!Num3 = txtTaxa
gge!num2 = txtNumFatura
gge!Alfa2 = txtUnidadeOperacional
gge!Alfa3 = txtNumSerie
gge!data2 = DataVenc
gge.Update

'db.CommitTrans

Item = 0

Set Rel = drFaturaEsp
'txtTaxa = txtTaxa
neg.Open "Select * from negociacao where chNumPedido = ('" & txtMedicao & "') and negTipoProduto = ('" & TipoProduto & "')", db, 3, 3
If neg.EOF Then
   If neg.State = 1 Then
      neg.Close: Set neg = Nothing
   End If
   neg.Open "Select * from historiconegociacao where chNumPedido = ('" & txtMedicao & "') and hnegTipoProduto = ('" & TipoProduto & "')", db, 3, 3
   If neg.EOF Then
      MsgBox ("Número do Pedido Inexistente. Comunicar ao analista responsável"), vbCritical
      Call FechaDB
      Exit Sub
   Else
      sql = "Select gge.ggeDatahoje, gge.ggeDataIni, gge.ggeDataFim, gge.Alfa2, Alfa3, gge.chNumerica, gge.num2, Unid.AbreviaturaUnidadeMedida, "
      sql = sql & " neg.chPessoa, neg.chUnidadeOperacional, neg.negContrato, neg.negContratoComp, neg.chNumPedido, neg.chNumPedidoComp, gge.Data2, "
      sql = sql & " pes.pesRazaoSocial, pes.pesEndereco, pes.pesBairro, pes.pesCidade, pes.chUF, pes.pesCEP, pes.chCNPJ_CPF, pes.pesInscEst_Ident, pes.pesTelContato, "
      sql = sql & " det.chDataInicio, det.chDataFim, det.chProduto, det.hdnQuantidadePedida / 1 as QtdPedida, "
      sql = sql & " (det.hdnPrecoUnidadePedida - (det.hdnPrecoUnidadePedida * num3) /100) as PrecoUnit, "
      sql = sql & " (det.hdnValorDaDiaria - (det.hdnValorDaDiaria * num3) /100) as PrecoDiaria, det.hdnQtdDias / 1 as QtdDias, det.hdnAtividade, prd.prdDescCompleta, prd.chProduto, prd.prdNomeProd, prdNomeComercial, "
      sql = sql & " ROUND(((det.hdnValorDaDiaria - (det.hdnValorDaDiaria * num3) /100) * det.hdnQtdDias), 1) as PrecoOperacao "
      sql = sql & " From geradorgeral gge, unidadedemedida Unid, historiconegociacao neg, historicodetalhenegociacao det, pessoa pes, produto prd "
      sql = sql & " WHERE neg.chNumpedido = ('" & txtMedicao & "') and neg.hnegTipoProduto = ('" & TipoProduto & "') and gge.chAlfaNumerica = ('" & Relatorio & "') and det.chProduto = prd.chProduto "
      sql = sql & " and det.chNumpedido = neg.chNumpedido and det.chNumpedidoComp = neg.chNumPedidoComp "
      sql = sql & " and neg.chPessoa = pes.chPessoa and Unid.chUnidadeDeMedida = det.hdnUnidade "
      sql = sql & " order by neg.chUnidadeOperacional, prd.chProduto "
   End If
Else
   sql = "Select gge.ggeDatahoje, gge.ggeDataIni, gge.ggeDataFim, gge.Alfa2, Alfa3, gge.chNumerica, gge.num2, Unid.AbreviaturaUnidadeMedida, "
   sql = sql & " neg.chPessoa, neg.chUnidadeOperacional, neg.negContrato, neg.negContratoComp, neg.chNumPedido, neg.chNumPedidoComp, gge.Data2, "
   sql = sql & " pes.pesRazaoSocial, pes.pesEndereco, pes.pesBairro, pes.pesCidade, pes.chUF, pes.pesCEP, pes.chCNPJ_CPF, pes.pesInscEst_Ident, pes.pesTelContato, "
   sql = sql & " det.chDataInicio, det.chDataFim, det.chProduto, det.pedQuantidadePedida / 1 as QtdPedida, "
   sql = sql & " (det.pedPrecoUnidadePedida - (det.pedPrecoUnidadePedida * num3) /100) as PrecoUnit, "
   sql = sql & " (det.pedValorDaDiaria - (det.pedValorDaDiaria * num3) /100) as PrecoDiaria, det.pedQtdDias / 1 as QtdDias, det.pedAtividade, prd.prdDescCompleta, prd.chProduto, prd.prdNomeProd, prdNomeComercial, "
   sql = sql & " ROUND(((det.pedValorDaDiaria - (det.pedValorDaDiaria * num3) /100) * det.pedQtdDias), 1) as PrecoOperacao "
   sql = sql & " From geradorgeral gge, unidadedemedida Unid, negociacao neg, detalhenegociacao det, pessoa pes, produto prd "
   sql = sql & " WHERE neg.chNumpedido = ('" & txtMedicao & "') and neg.negTipoProduto = ('" & TipoProduto & "') and gge.chAlfaNumerica = ('" & Relatorio & "') and det.chProduto = prd.chProduto "
   sql = sql & " and det.chNumpedido = neg.chNumpedido and det.chNumpedidoComp = neg.chNumPedidoComp "
   sql = sql & " and neg.chPessoa = pes.chPessoa and Unid.chUnidadeDeMedida = det.pedUnidade "
   sql = sql & " order by neg.chUnidadeOperacional, prd.chProduto "
End If

AbrirRelatorio sql, Rel

'Resp = MsgBox("O número da Fatura e a emissão da mesma estão corretos???", vbExclamation + vbYesNo)
'   If Resp = vbNo Then
'      Call FechaDB
'      Exit Sub
'   End If
   

Call FechaDB

End Sub

Private Sub GridFatura_Click()
Dim Limite As Integer
Dim IndLinha As Integer

Limite = GridFatura.Rows

IndLinha = GridFatura.Row

If GridFatura.TextMatrix(IndLinha, 1) = "" Then
   MsgBox "Clicar em linha com conteúdo."
   Exit Sub
End If

txtPessoa = GridFatura.TextMatrix(IndLinha, 3)

Call Rotina_AbrirBanco

pes.Open "Select * from pessoa where chPessoa = ('" & GridFatura.TextMatrix(IndLinha, 3) & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Cliente não encontrado. Comuniicar ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
End If


txtRazaoSocial = pes!pesRazaoSocial
txtUnidadeOperacional = GridFatura.TextMatrix(IndLinha, 4)
txtNome = txtPessoa
txtPedidoComp = GridFatura.TextMatrix(IndLinha, 2)
txtComplemento = GridFatura.TextMatrix(IndLinha, 2)
txtNumPedido = GridFatura.TextMatrix(IndLinha, 1)
txtMedicao = GridFatura.TextMatrix(IndLinha, 1)
DataVenc = GridFatura.TextMatrix(IndLinha, 5)
dtInicio = GridFatura.TextMatrix(IndLinha, 6)
dtFim = GridFatura.TextMatrix(IndLinha, 7)
txtNumSerie = GridFatura.TextMatrix(IndLinha, 8)
If GridFatura.TextMatrix(IndLinha, 9) = Empty Then
   dtEmis = Date
Else
   dtEmis = GridFatura.TextMatrix(IndLinha, 9)
End If

neg.Open "Select * from negociacao where chNumPedido = ('" & txtNumPedido & "') and negTipoProduto = ('" & 0 & "')", db, 3, 3
If neg.EOF Then
   If neg.State = 1 Then
      neg.Close: Set neg = Nothing
   End If
   neg.Open "Select * From historiconegociacao where chNumPedido = ('" & txtNumPedido & "') and hnegTipoProduto = ('" & 0 & "')", db, 3, 3
   If neg.EOF Then
      MsgBox ("Pedido inexistente. Reportar ao analista responsável"), vbCritical
      Call FechaDB
      Exit Sub
   End If
   txtNumFatura = neg!negNumFatura
   txtContrato = neg!negContrato
   If IsNull(neg!negContratoComp) Then
      txtContratoComp = Empty
   Else
      txtContratoComp = neg!negContratoComp
   End If
Else
   txtContrato = neg!negContrato
   If IsNull(neg!negContratoComp) Then
      txtContratoComp = Empty
   Else
      txtContratoComp = neg!negContratoComp
   End If
   txtNumFatura = neg!negNumFatura
   If Not IsNull(neg!negNumFatura) Then
      MsgBox ("Esta Fatura já foi Impressa. A reimpressão pode ser efetuada"), vbExclamation
      txtNumFatura = neg!negNumFatura
   Else
      Emp.Open "Select * from empresa", db, 3, 3
      If Emp.EOF Then
         MsgBox ("Banco de dados sem registro da empresa. Reportar ao analista responsável"), vbCritical
         Call FechaDB
         Exit Sub
      Else
         Emp.MoveFirst
         If pes!pesClassFiscal = "Lucro Real" Then
            txtNumFatura = Format$((Emp!empNumFatura + 1), "000")
         Else
            txtNumFatura = Format$((Emp!empNumFaturaE + 1), "000")
         End If
      End If
   End If
End If

Dia = Day(Date)
mes = Month(Date)
ano = Year(Date)

txtTaxa.SetFocus

Call FechaDB

End Sub


Private Sub txtNumFatura_LostFocus()
If Not IsNumeric(txtNumFatura) Then
   MsgBox ("Esta informação só pode conter números."), vbCritical
   cmdSair.SetFocus
   Exit Sub
End If
   
End Sub

