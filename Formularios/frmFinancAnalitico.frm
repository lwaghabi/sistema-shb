VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFinancAnalitico 
   Caption         =   "frmFinancAnalitico"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17670
   LinkTopic       =   "Form3"
   ScaleHeight     =   10155
   ScaleWidth      =   17670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Atrasados a Receber"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3495
      Left            =   0
      TabIndex        =   29
      Top             =   5280
      Width           =   8775
      Begin MSFlexGridLib.MSFlexGrid GridAtrasados 
         Height          =   3015
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5318
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColor       =   14737632
         BackColorFixed  =   14737632
         BackColorSel    =   14737632
         BackColorBkg    =   14737632
         FormatString    =   "Vencimento|Cliente                  |Descrição                 |Valor           |"
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pagamentos na Semana"
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
      Height          =   3615
      Left            =   8760
      TabIndex        =   25
      Top             =   720
      Width           =   8775
      Begin MSFlexGridLib.MSFlexGrid GridPagar 
         Height          =   3135
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColor       =   14737632
         BackColorFixed  =   14737632
         BackColorSel    =   14737632
         BackColorBkg    =   14737632
         FormatString    =   "Vencimento|Cliente                  |Descrição                 |Valor           |"
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
   End
   Begin VB.Frame Frame9 
      Caption         =   "Recebimentos na Semana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3615
      Left            =   0
      TabIndex        =   20
      Top             =   720
      Width           =   8655
      Begin MSFlexGridLib.MSFlexGrid GridReceber 
         Height          =   3135
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColor       =   14737632
         BackColorFixed  =   14737632
         BackColorSel    =   14737632
         BackColorBkg    =   14737632
         FormatString    =   "Vencimento|Cliente                  |Descrição                 |Valor           |"
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
   End
   Begin VB.Frame Frame6 
      Caption         =   "Atrasados a Pagar"
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
      Height          =   3495
      Left            =   8760
      TabIndex        =   15
      Top             =   5280
      Width           =   8775
      Begin MSFlexGridLib.MSFlexGrid GridPagtosAtrasados 
         Height          =   3015
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5318
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColor       =   14737632
         BackColorFixed  =   14737632
         BackColorSel    =   14737632
         BackColorBkg    =   14737632
         FormatString    =   "Vencimento|Cliente                  |Descrição                   |Valor           |"
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
   End
   Begin VB.Frame Frame8 
      Caption         =   "Navegação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11640
      TabIndex        =   12
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton cmdNavega 
         BackColor       =   &H0080FFFF&
         Caption         =   "Anterior"
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
         Index           =   2
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdNavega 
         BackColor       =   &H0080FF80&
         Caption         =   "Próximo"
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
         Index           =   1
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H008080FF&
         Caption         =   "Sair"
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdNavega 
         BackColor       =   &H00FFFF80&
         Caption         =   "Início"
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
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   11
      Top             =   0
      Width           =   1695
      Begin VB.ComboBox cmbFiltro 
         BackColor       =   &H00FFFFE0&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Hoje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   9
      Top             =   0
      Width           =   1575
      Begin MSMask.MaskEdBox txtDataHoje 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16777184
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Semana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      TabIndex        =   6
      Top             =   0
      Width           =   1815
      Begin MSComCtl2.DTPicker txtDataFim 
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   420
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CalendarBackColor=   16777184
         Format          =   244383745
         CurrentDate     =   38548
      End
      Begin MSComCtl2.DTPicker txtDataInicio 
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   180
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CalendarBackColor=   16777184
         Format          =   244383745
         CurrentDate     =   38548
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "De"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Até"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   300
      End
   End
   Begin VB.Label lblTotalAtrasado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   420
      Left            =   15555
      TabIndex        =   31
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Atrasados a Pagar - Total"
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
      Left            =   11760
      TabIndex        =   30
      Top             =   8880
      Width           =   3855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Atrasados a Receber - Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   1920
      TabIndex        =   28
      Top             =   8880
      Width           =   4155
   End
   Begin VB.Label lblTotalAtraso 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   420
      Left            =   6480
      TabIndex        =   27
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Label lblTotalPagarNew 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   420
      Left            =   15240
      TabIndex        =   26
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "Pagamentos na Semana - Total"
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
      Height          =   360
      Left            =   10560
      TabIndex        =   24
      Top             =   4440
      Width           =   4380
   End
   Begin VB.Label lblTotalPagar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1.01140e5
      TabIndex        =   23
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Recebimentos na Semana - Total "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   1200
      TabIndex        =   22
      Top             =   4440
      Width           =   4815
   End
   Begin VB.Label lblTotalReceber 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   420
      Left            =   6360
      TabIndex        =   21
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Saldo Em Atraso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   13920
      TabIndex        =   19
      Top             =   9480
      Width           =   1575
   End
   Begin VB.Label lblSaldoEmAtraso 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   420
      Left            =   15555
      TabIndex        =   18
      Top             =   9480
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Saldo na Semana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12480
      TabIndex        =   17
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label lblSaldoSemana 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   420
      Left            =   15240
      TabIndex        =   16
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Financeiro Analítico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2670
   End
End
Attribute VB_Name = "frmFinancAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim NomeGrid As String
Dim Ind As Double
Dim Linha As Double
Dim Coluna As Integer
Dim IndReceber As Byte
Dim IndAtrasado As Byte
Dim IndPagar As Byte
Dim IndPagarAtrasados As Byte

Dim IndConf As Byte
Dim AcumulaCtaReceber As Currency
Dim AcumulaCtaPagar As Currency
Dim AcumulaCtaAtraso As Currency
Dim AcumulaPagarAtrasados As Currency

Dim DiaUtilAnterior As Date
Dim DataInicio As Date
Dim DataFim As Date
Dim DataReceber As Date
Dim DataPagos As Date
Dim DataAtrasados As Date
Dim DataInformada As Date
Dim DiadaSemana As Integer

Dim Dia As String
Dim Mes As String
Dim Ano As String

Dim DataInvertida As String

Dim Indice As Byte
Dim DataAcesso As Date

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtDataHoje = Date
DataInformada = Date

NDias = 1

'DataRetorno = ObterDiaUtilAnterior(DataInformada, NDias)
'DiaUtilAnterior = DataRetorno.DiaUtil

Call Rotina_070_Ajusta_Data


'DataPagos = DiaUtilAnterior
DataPagos = Date - 1
DataRetorno = Date - 1
'NDias = 2

'DataRetorno = ObterDiaUtilAnterior(DataInformada, NDias)

'DataAtrasados = DataRetorno.DiaUtil
DataAtrasados = Date - 1
AcumulaCtaReceber = 0
AcumulaCtaPagar = 0
AcumulaCtaAtraso = 0
AcumulaPagarAtrasados = 0

cmbFiltro.Clear
cmbFiltro.AddItem "Geral"

Call Rotina_AbrirBanco

Bco.Open "Select * from Banco", db, 3, 3
If Bco.EOF Then
   MsgBox ("Tabela de Banco vazia. "), vbCritical
   Call FechaDB
   Exit Sub
End If


Bco.MoveFirst

Do While Not Bco.EOF
   cmbFiltro.AddItem Bco!bcosiglabco
   Bco.MoveNext
Loop

cmbFiltro.ListIndex = 0

Call Rotina_010_Limpa_Cta_Pagar

Call Rotina_011_Limpa_Atrasados

Call Rotina_012_Limpa_Cta_Receber

Call Rotina_013_Limpa_Atrasados_Pagar

Call Rotina_020_Gerencia_Grid

Call Rotina_021_Gerencia_Grid_Pagar

Call FechaDB

End Sub

Public Sub Rotina_010_Limpa_Cta_Pagar()
GridPagar.Rows = 2
IndReceber = 1
GridPagar.TextMatrix(IndReceber, 0) = Empty
GridPagar.TextMatrix(IndReceber, 1) = Empty
GridPagar.TextMatrix(IndReceber, 2) = Empty
GridPagar.TextMatrix(IndReceber, 3) = Empty

End Sub
Public Sub Rotina_011_Limpa_Atrasados()
GridAtrasados.Rows = 2
IndAtrasado = 1
GridAtrasados.TextMatrix(IndAtrasado, 0) = Empty
GridAtrasados.TextMatrix(IndAtrasado, 1) = Empty
GridAtrasados.TextMatrix(IndAtrasado, 2) = Empty
GridAtrasados.TextMatrix(IndAtrasado, 3) = Empty
GridAtrasados.TextMatrix(IndAtrasado, 4) = Empty

End Sub
Public Sub Rotina_012_Limpa_Cta_Receber()
GridReceber.Rows = 2
IndConf = 1
GridReceber.TextMatrix(IndConf, 0) = Empty
GridReceber.TextMatrix(IndConf, 1) = Empty
GridReceber.TextMatrix(IndConf, 2) = Empty
GridReceber.TextMatrix(IndConf, 3) = Empty

End Sub
Public Sub Rotina_013_Limpa_Atrasados_Pagar()
GridAtrasados.Rows = 2
IndPagarAtrasados = 1
GridPagtosAtrasados.TextMatrix(IndPagarAtrasados, 0) = Empty
GridPagtosAtrasados.TextMatrix(IndPagarAtrasados, 1) = Empty
GridPagtosAtrasados.TextMatrix(IndPagarAtrasados, 2) = Empty
GridPagtosAtrasados.TextMatrix(IndPagarAtrasados, 3) = Empty
GridPagtosAtrasados.TextMatrix(IndPagarAtrasados, 4) = Empty

End Sub
Public Sub Rotina_020_Gerencia_Grid()

IndReceber = 0
IndAtrasado = 0
IndPagarAtrasados = 0

Call Rotina_AbrirBanco

ctr.Open "Select * from Contas_A_Receber", db, 3, 3
If Not (ctr.EOF) Then
   
   ctr.MoveFirst
   Do While Not ctr.EOF
   
      Indice = cmbFiltro.ListIndex
   
      If (cmbFiltro = "Geral") Or (cmbFiltro = ctr!chCodBcoLart) Then
         If ctr!ctrDataBanco > DataInicio - 1 And ctr!ctrDataBanco < DataFim Then
            If ctr!ctrStatus = 0 And ctr!ctrDataBanco > (Date - 1) Then
               Call Rotina_051_Carga_Cta_Receber
            End If
         End If
  
           
         If ctr!ctrStatus = 0 Then
            If ctr!ctrDataBanco < Date Then
               Call Rotina_052_Carga_Atrasados
            End If
         End If
      End If
   
      ctr.MoveNext

   Loop
End If
'txtDataInicio = Date
'txtDataFim = Date + 7

lblTotalReceber = Format$(AcumulaCtaReceber, "##,##0.00")
lblTotalAtraso = Format$(AcumulaCtaAtraso, "##,##0.00")

GridReceber.Col = 4
GridReceber.ColSel = 4
     
GridReceber.Row = 1
GridReceber.RowSel = IndReceber
        
If IndReceber > 1 Then
   GridReceber.Sort = 5
End If

GridReceber.Col = 0
GridReceber.ColSel = 0
GridReceber.Row = 0
GridReceber.RowSel = 0

GridAtrasados.Col = 4
GridAtrasados.ColSel = 4
     
GridAtrasados.Row = 1
GridAtrasados.RowSel = IndAtrasado
        
If IndAtrasado > 1 Then
   GridAtrasados.Sort = 5
End If

GridAtrasados.Col = 0
GridAtrasados.ColSel = 0
GridAtrasados.Row = 0
GridAtrasados.RowSel = 0


Call FechaDB

End Sub
Public Sub Rotina_021_Gerencia_Grid_Pagar()

IndPagar = 0

Call Rotina_AbrirBanco

ctp.Open "Select * from Contas_A_Pagar", db, 3, 3
If ctp.EOF Then
   Call FechaDB
   Exit Sub
End If

ctp.MoveFirst

Do While Not ctp.EOF
   
   Indice = cmbFiltro.ListIndex
   
   If (cmbFiltro = "Geral") Or (cmbFiltro = ctp!chCodBcoLart) Then
      If ctp!chDataVencito > DataInicio And ctp!chDataVencito < DataFim Then
         If ctp!ctpStatus = 0 Then 'DataAtrasados + 2 Then
            Call Rotina_050_Carga_Cta_Pagar
         End If
      End If
      If ctp!ctpStatus = 0 Then
         If ctp!chDataVencito < Date Then 'DataAtrasados + 3 Then
            Call Rotina_050_Pagar_Atrasados
         End If
      End If
   End If
   
   ctp.MoveNext

Loop
      

lblTotalPagarNew = Format$(AcumulaCtaPagar, "##,##0.00")
lblTotalAtrasado = Format$(AcumulaPagarAtrasados, "##,##0.00")

lblSaldoSemana = Format$((AcumulaCtaReceber - AcumulaCtaPagar), "##,##0.00")
lblSaldoEmAtraso = Format$((AcumulaCtaAtraso - AcumulaPagarAtrasados), "##,##0.00")

GridPagar.Col = 4
GridPagar.ColSel = 4
     
GridPagar.Row = 1
GridPagar.RowSel = IndPagar
        
If IndPagar > 1 Then
   GridPagar.Sort = 5
End If

GridPagar.Col = 0
GridPagar.ColSel = 0
GridPagar.Row = 0
GridPagar.RowSel = 0

GridAtrasados.Col = 4
GridAtrasados.ColSel = 4

Call FechaDB

End Sub

Public Sub Rotina_050_Carga_Cta_Pagar()

IndPagar = IndPagar + 1
GridPagar.Rows = IndPagar + 1
GridPagar.TextMatrix(IndPagar, 0) = ctp!chDataVencito
GridPagar.TextMatrix(IndPagar, 1) = ctp!chPessoa
GridPagar.TextMatrix(IndPagar, 2) = ctp!ctpdescricaooperacao
GridPagar.TextMatrix(IndPagar, 3) = Format$(ctp!ctpValorDaBoleta, "##,##0.00")

Dia = Format$(Day(ctp!chDataVencito), "00")
Mes = Format$(Month(ctp!chDataVencito), "00")
Ano = Year(ctp!chDataVencito)

DataInvertida = Ano & Mes & Dia

GridPagar.TextMatrix(IndPagar, 4) = DataInvertida & ctp!ctpdescricaooperacao & ctp!chPessoa
AcumulaCtaPagar = AcumulaCtaPagar + ctp!ctpValorDaBoleta

End Sub
Public Sub Rotina_050_Pagar_Atrasados()

IndPagarAtrasados = IndPagarAtrasados + 1
GridPagtosAtrasados.Rows = IndPagarAtrasados + 1
GridPagtosAtrasados.TextMatrix(IndPagarAtrasados, 0) = ctp!chDataVencito
GridPagtosAtrasados.TextMatrix(IndPagarAtrasados, 1) = ctp!chPessoa
GridPagtosAtrasados.TextMatrix(IndPagarAtrasados, 2) = ctp!ctpdescricaooperacao
GridPagtosAtrasados.TextMatrix(IndPagarAtrasados, 3) = Format$(ctp!ctpValorDaBoleta, "##,##0.00")

Dia = Format$(Day(ctp!chDataVencito), "00")
Mes = Format$(Month(ctp!chDataVencito), "00")
Ano = Year(ctp!chDataVencito)

DataInvertida = Ano & Mes & Dia

GridPagtosAtrasados.TextMatrix(IndPagarAtrasados, 4) = DataInvertida & ctp!chPessoa
AcumulaPagarAtrasados = AcumulaPagarAtrasados + ctp!ctpValorDaBoleta

End Sub
Public Sub Rotina_051_Carga_Cta_Receber()

IndReceber = IndReceber + 1
GridReceber.Rows = IndReceber + 1
GridReceber.TextMatrix(IndReceber, 0) = ctr!ctrDataVencito
GridReceber.TextMatrix(IndReceber, 1) = ctr!chPessoa
GridReceber.TextMatrix(IndReceber, 2) = ctr!ctrDescricaoOperacao
GridReceber.TextMatrix(IndReceber, 3) = Format$(ctr!ctrValorDaBoleta, "##,##0.00")

Dia = Format$(Day(ctr!ctrDataVencito), "00")
Mes = Format$(Month(ctr!ctrDataVencito), "00")
Ano = Year(ctr!ctrDataVencito)

DataInvertida = Ano & Mes & Dia

GridReceber.TextMatrix(IndReceber, 4) = DataInvertida & ctr!chPessoa

AcumulaCtaReceber = AcumulaCtaReceber + ctr!ctrValorDaBoleta
End Sub
Public Sub Rotina_052_Carga_Atrasados()

IndAtrasado = IndAtrasado + 1
GridAtrasados.Rows = IndAtrasado + 1
GridAtrasados.TextMatrix(IndAtrasado, 0) = ctr!ctrDataVencito
GridAtrasados.TextMatrix(IndAtrasado, 1) = ctr!chPessoa
GridAtrasados.TextMatrix(IndAtrasado, 2) = ctr!ctrDescricaoOperacao
GridAtrasados.TextMatrix(IndAtrasado, 3) = Format(ctr!ctrValorDaBoleta, "##,##0.00")

Dia = Format$(Day(ctr!ctrDataVencito), "00")
Mes = Format$(Month(ctr!ctrDataVencito), "00")
Ano = Year(ctr!ctrDataVencito)

DataInvertida = Ano & Mes & Dia

GridAtrasados.TextMatrix(IndAtrasado, 4) = ctr!chPessoa

AcumulaCtaAtraso = AcumulaCtaAtraso + ctr!ctrValorDaBoleta

End Sub

Private Sub cmdNavega_Click(Index As Integer)
   
Select Case Index

Case 0
     DataInformada = Date
     Call Rotina_070_Ajusta_Data
     cmdNavega(0).SetFocus
Case 1
     
     If DataFim = Empty Then
        DataInformada = Date
        Call Rotina_070_Ajusta_Data
     Else
        DataInformada = txtDataFim
        DataInformada = DataInformada + 2
        Call Rotina_070_Ajusta_Data
     End If
     cmdNavega(1).SetFocus
Case 2
     If DataInicio = Empty Then
        DataInformada = Date
        Call Rotina_070_Ajusta_Data
     Else
        DataInformada = txtDataInicio
        DataInformada = DataInformada - 1
        Call Rotina_070_Ajusta_Data
     End If
     cmdNavega(2).SetFocus
End Select

AcumulaCtaReceber = 0
AcumulaCtaPagar = 0
AcumulaCtaAtraso = 0
AcumulaPagarAtrasados = 0

Call Rotina_010_Limpa_Cta_Pagar

Call Rotina_012_Limpa_Cta_Receber

Call Rotina_013_Limpa_Atrasados_Pagar

Call Rotina_020_Gerencia_Grid

Call Rotina_021_Gerencia_Grid_Pagar

End Sub

Public Sub Rotina_070_Ajusta_Data()

DiadaSemana = Weekday(DataInformada)

'Calcular Range de datas

DataInicio = DataInformada - (DiadaSemana)
DataFim = DataInicio + 7
txtDataInicio = DataInicio
txtDataFim = DataFim

End Sub

