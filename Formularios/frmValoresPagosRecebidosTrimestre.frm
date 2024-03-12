VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmValoresPagosRecebidosTrimestre 
   Caption         =   "frmValoresPagosRecebidosTrimestre"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   14850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H0000FFFF&
      Cancel          =   -1  'True
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7800
      Width           =   2175
   End
   Begin VB.ComboBox cmbAnoRef 
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtPercentSaldoMedio 
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
      Height          =   615
      Left            =   9840
      TabIndex        =   2
      Top             =   7080
      Width           =   735
   End
   Begin VB.ComboBox cmbTrimestre 
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
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtFim 
      Height          =   495
      Left            =   9000
      TabIndex        =   9
      Top             =   1920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   240189441
      CurrentDate     =   44869
   End
   Begin MSComCtl2.DTPicker dtInicio 
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   240189441
      CurrentDate     =   44869
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   2760
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   873
      _Version        =   393216
      Format          =   240189441
      CurrentDate     =   44869
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   14535
      Begin VB.TextBox txtHoje 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   615
         Left            =   12000
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
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
         Left            =   12120
         TabIndex        =   5
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Valores Recebidos e Pagos por trimestre"
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
         TabIndex        =   4
         Top             =   480
         Width           =   11535
      End
   End
   Begin VB.Label Label14 
      Caption         =   "Até"
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
      Left            =   2040
      TabIndex        =   43
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "De"
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
      Left            =   120
      TabIndex        =   42
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Trimestre Pesquisado"
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
      Left            =   6840
      TabIndex        =   41
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label11 
      Caption         =   "Trimestre"
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
      Left            =   3480
      TabIndex        =   40
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblCalcSobreMedia 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   11400
      TabIndex        =   39
      Top             =   7080
      Width           =   3015
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   375
      Left            =   11400
      TabIndex        =   38
      Top             =   7080
      Width           =   15
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   " % / Sdo Total"
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
      Left            =   7560
      TabIndex        =   37
      Top             =   7200
      Width           =   2085
   End
   Begin VB.Label lblMediaT3 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   11400
      TabIndex        =   36
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label lblMediaT2 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   7560
      TabIndex        =   35
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label lblMediaT1 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   4080
      TabIndex        =   34
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Média no Trimestre"
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
      Left            =   0
      TabIndex        =   33
      Top             =   6480
      Width           =   3855
   End
   Begin VB.Label lblTotalT3 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   11400
      TabIndex        =   32
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label lblTotalT2 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   7560
      TabIndex        =   31
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label lblTotalT1 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   4080
      TabIndex        =   30
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Totais no Trimestre"
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
      Left            =   120
      TabIndex        =   29
      Top             =   5880
      Width           =   3735
   End
   Begin VB.Label Label6 
      Caption         =   "Ano Ref.:"
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
      Left            =   1080
      TabIndex        =   28
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Saldo no Mês"
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
      Left            =   11400
      TabIndex        =   27
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Totais Pagos"
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
      Left            =   7560
      TabIndex        =   26
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label lblFimT3 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   2040
      TabIndex        =   25
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label lblInicioT3 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label lblFimT2 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   2040
      TabIndex        =   23
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblInicioT2 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblFimT1 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   2040
      TabIndex        =   21
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblInicioT1 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblSaldo3T 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   11400
      TabIndex        =   19
      Top             =   4920
      Width           =   3015
   End
   Begin VB.Label lblTotalPago3T 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   7560
      TabIndex        =   18
      Top             =   4920
      Width           =   3015
   End
   Begin VB.Label lblTotalRecebido3T 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   4080
      TabIndex        =   17
      Top             =   4920
      Width           =   3015
   End
   Begin VB.Label lblSaldo2T 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   11400
      TabIndex        =   16
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label lblTotalPago2T 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   7560
      TabIndex        =   15
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label lblTotalRecebido2T 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label lblSaldo1T 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   11400
      TabIndex        =   13
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label lblTotalPago1T 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   7560
      TabIndex        =   12
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Totais Recebidos"
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
      Left            =   4080
      TabIndex        =   11
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label lblTotalRecebido1T 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   4080
      TabIndex        =   10
      Top             =   3240
      Width           =   3015
   End
End
Attribute VB_Name = "frmValoresPagosRecebidosTrimestre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim InicioPrimeiroTrimestre As Date
Dim FimPrimeiroTrimestre As Date
Dim InicioSegundoTrimestre As Date
Dim FimSegundoTrimestre As Date
Dim InicioTerceiroTrimestre As Date
Dim FimTerceiroTrimestre As Date
Dim InicioQuartoTrimestre As Date
Dim FimQuartoTrimestre As Date
Dim ano As Integer
Dim mes As Integer
Dim Dia As Integer
Dim DataMontada As String
Dim DataProcessada As Date
Dim Ind As Integer
Dim DataAuxiliarInicio As Date
Dim DataAuxiliarFim As Date
Dim DataInvertidaInicio As String
Dim DataInvertidaFim As String
Dim Status As Integer
Dim Encontrei As Integer

Dim AcumulaReceberT1 As Currency
Dim AcumulaReceberT2 As Currency
Dim AcumulaReceberT3 As Currency
Dim AcumulaPagarT1 As Currency
Dim AcumulaPagarT2 As Currency
Dim AcumulaPagarT3 As Currency
Dim SaldoT1 As Currency
Dim SaldoT2 As Currency
Dim SaldoT3 As Currency

Dim AcumulaValor As Currency

Private Sub cmbAnoRef_LostFocus()
Dim AnoReferencia As Integer

If Not Year(dtInicio) = cmbAnoRef Then
   dtInicio = Day(dtInicio) & "/" & Month(dtInicio) & "/" & cmbAnoRef
   dtFim = Day(dtInicio) & "/" & Month(dtInicio) & "/" & cmbAnoRef
End If

ano = Year(dtInicio)
Dia = Format$(1, "00")
mes = Format$(1, "00")

InicioPrimeiroTrimestre = Format$(Dia & "-" & mes & "-" & ano, "dd/mm/yyyy")

mes = mes + 3

InicioSegundoTrimestre = Format$(Dia & "-" & mes & "-" & ano, "dd/mm/yyyy")
FimPrimeiroTrimestre = InicioSegundoTrimestre - 1

mes = mes + 3

InicioTerceiroTrimestre = Format$(Dia & "-" & mes & "-" & ano, "dd/mm/yyyy")
FimSegundoTrimestre = InicioTerceiroTrimestre - 1

mes = mes + 3

InicioQuartoTrimestre = Format$(Dia & "-" & mes & "-" & ano, "dd/mm/yyyy")
FimTerceiroTrimestre = InicioQuartoTrimestre - 1

FimQuartoTrimestre = Format$(31 & "-" & 12 & "-" & ano, "dd/mm/yyyy")

End Sub

Private Sub cmbTrimestre_LostFocus()

Call LimparConsulta

If cmbTrimestre = 1 Then
   dtInicio = InicioPrimeiroTrimestre
   dtFim = FimPrimeiroTrimestre
Else
   If cmbTrimestre = 2 Then
      dtInicio = InicioSegundoTrimestre
      dtFim = FimSegundoTrimestre
   Else
      If cmbTrimestre = 3 Then
         dtInicio = InicioTerceiroTrimestre
         dtFim = FimTerceiroTrimestre
      Else
         If cmbTrimestre = 4 Then
            dtInicio = InicioQuartoTrimestre
            dtFim = FimQuartoTrimestre
         End If
      End If
   End If
End If

lblInicioT1 = dtInicio
mes = Month(dtInicio)
mes = mes + 1
Dia = Format$(1, "00")
ano = Year(dtInicio)
lblInicioT2 = Format$(Dia, "00") & "/" & Format$(mes, "00") & "/" & ano

DataProcessada = lblInicioT2
DataProcessada = DataProcessada - 1
lblFimT1 = DataProcessada

Dia = Day(dtFim)
mes = Month(dtFim)
ano = Year(dtFim)

lblFimT3 = dtFim
lblInicioT3 = Format$(1, "00") & "/" & Format$(mes, "00") & "/" & ano

'Mes = Mes - 1
mes = Month(dtFim)
Dia = Format$(1, "00")
ano = Year(dtInicio)

DataProcessada = Format$(1, "00") & "/" & Format$(mes, "00") & "/" & ano
DataProcessada = DataProcessada - 1

lblFimT2 = DataProcessada

Call ProcessaValoresReceber

Call ProcessaValoresPagos

Call FinalizarConsulta

End Sub


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Limite As Integer

Limite = 2020

txtHoje = Date

dtInicio = Date
dtFim = Date

For Ind = 1 To 4
    cmbTrimestre.AddItem Ind
Next

Limite = Year(Date) - 2020
ano = Year(Date)
For Ind = 1 To Limite
    cmbAnoRef.AddItem ano
    ano = ano - 1
Next

cmbAnoRef.ListIndex = 0
cmbTrimestre.ListIndex = 0

ano = Year(dtInicio)
Dia = Format$(1, "00")
mes = Format$(1, "00")

End Sub

Public Sub ProcessaValoresReceber()

AcumulaReceberT1 = 0
AcumulaReceberT2 = 0
AcumulaReceberT3 = 0
Encontrei = 0

'Trata Trimestre 1 a Receber

If Year(dtInicio) = Year(Date) Then
   If Month(lblInicioT1) > Month(dtInicio) Then
      MsgBox ("Período Inválido."), vbCritical
      Exit Sub
   End If
End If

If Month(lblInicioT1) = Month(Date) And Year(dtInicio) = Year(Date) Then
   Call ValoresDoMes
Else
   DataAuxiliarInicio = lblInicioT1
   DataAuxiliarFim = lblFimT1
   Call InverterData
   Call ValoresDoHistorico
End If

If Encontrei = 1 Then
   Encontrei = 0
   If Not ctr.EOF Then
      Call ProcessarTrimestre
      AcumulaReceberT1 = AcumulaValor
      lblTotalRecebido1T = Format$(AcumulaValor, "###,##0.00")
      AcumulaValor = 0
   End If
End If

lblTotalRecebido1T.ForeColor = &HFF0000

'Fim Trata Trimestre 1 a Receber

'Trata Trimestre 2 a Receber

If Month(lblInicioT2) > Month(Date) And Year(dtInicio) = Year(Date) Then
   AcumulaReceberT2 = 0
   lblTotalRecebido2T = Format$(0, "###,##0.00")
   AcumulaValor = 0
   Exit Sub
End If

If Month(lblInicioT2) = Month(Date) And Year(dtInicio) = Year(Date) Then
   Call ValoresDoMes
Else
   DataAuxiliarInicio = lblInicioT2
   DataAuxiliarFim = lblFimT2
   Call InverterData
   Call ValoresDoHistorico
End If

If Encontrei = 1 Then
   Encontrei = 0
   If Not ctr.EOF Then
      Call ProcessarTrimestre
      AcumulaReceberT2 = AcumulaValor
      lblTotalRecebido2T = Format$(AcumulaValor, "###,##0.00")
      AcumulaValor = 0
   End If
End If

lblTotalRecebido2T.ForeColor = &HFF0000

'Fim Trata Trimestre 2 a Receber

'Trata Trimestre 3 a Receber

If Month(lblInicioT3) > Month(Date) And Year(dtInicio) = Year(Date) Then
   AcumulaReceberT3 = 0
   lblTotalRecebido3T = Format$(0, "###,##0.00")
   AcumulaValor = 0
   Exit Sub
End If

If Month(lblInicioT3) = Month(Date) And Year(dtInicio) = Year(Date) Then
   Call ValoresDoMes
Else
   DataAuxiliarInicio = lblInicioT3
   DataAuxiliarFim = lblFimT3
   Call InverterData
   Call ValoresDoHistorico
End If

If Encontrei = 1 Then
   Encontrei = 0
   If Not ctr.EOF Then
      Call ProcessarTrimestre
      AcumulaReceberT3 = AcumulaValor
      lblTotalRecebido3T = Format$(AcumulaValor, "###,##0.00")
      AcumulaValor = 0
   End If
End If

lblTotalRecebido3T.ForeColor = &HFF0000

'Fim Trata Trimestre 3 a Receber

End Sub

Public Sub ValoresDoMes()

Call Rotina_AbrirBanco

ctr.Open "Select * from contas_a_receber where ctrStatus = ('" & 1 & "')", db, 3, 3
If ctr.EOF Then
   MsgBox ("Sem financeiro para tratar neste mes do trimestre."), vbInformation
   Call FechaDB
Else
   Encontrei = 1
End If

End Sub

Public Sub ValoresDoHistorico()
Call Rotina_AbrirBanco

Status = 1
ctr.Open "Select * from historicocontasreceber where ctrStatus = ('" & Status & "') and ctrDataRecebimento > ('" & DataInvertidaInicio & "') and ctrDataRecebimento < ('" & DataInvertidaFim & "')", db, 3, 3
If ctr.EOF Then
   MsgBox ("Sem financeiro para tratar neste mês do trimestre."), vbInformation
   Call FechaDB
Else
   Encontrei = 1
End If
End Sub

Public Sub ProcessarTrimestre()

AcumulaValor = 0

ctr.MoveFirst

Do While Not ctr.EOF
   If ctr!ctrStatus = 1 Then
      AcumulaValor = AcumulaValor + ctr!ctrValorDaBoleta
   End If
   
   ctr.MoveNext
Loop

End Sub

Public Sub InverterData()

DataAuxiliarInicio = DataAuxiliarInicio - 1

Dia = Day(DataAuxiliarInicio)
mes = Month(DataAuxiliarInicio)
ano = Year(DataAuxiliarInicio)

DataInvertidaInicio = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")

DataAuxiliarFim = DataAuxiliarFim + 1

Dia = Day(DataAuxiliarFim)
mes = Month(DataAuxiliarFim)
ano = Year(DataAuxiliarFim)

DataInvertidaFim = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")

End Sub

Public Sub ProcessaValoresPagos()
AcumulaPagarT1 = 0
AcumulaPagarT2 = 0
AcumulaPagarT3 = 0

'Trata Trimestre 1 Pagas
If Year(lblInicioT1) = Year(Date) Then
   If Month(lblInicioT1) > Month(Date) Then
      MsgBox ("Período Inválido."), vbCritical
      Exit Sub
   End If
End If

If Month(lblInicioT1) = Month(Date) And Year(dtInicio) = Year(Date) Then
   Call ValoresDoMesPagos
Else
   DataAuxiliarInicio = lblInicioT1
   DataAuxiliarFim = lblFimT1
   Call InverterData
   Call ValoresDoHistoricoPagos
End If

If Not ctp.EOF Then
   Call ProcessarTrimestrePagos
   AcumulaPagarT1 = AcumulaValor
   lblTotalPago1T = Format$(AcumulaValor, "###,##0.00")
   AcumulaValor = 0
End If

lblTotalPago1T.ForeColor = &HFF&

'Fim Trata Trimestre 1 a Pagar

'Trata Trimestre 2 a Receber

If Month(lblInicioT2) > Month(Date) And Year(dtInicio) = Year(Date) Then
   AcumulaPagarT2 = 0
   lblTotalPago2T = Format$(0, "###,##0.00")
   AcumulaValor = 0
   Exit Sub
End If

If Month(lblInicioT2) = Month(Date) And Year(dtInicio) = Year(Date) Then
   Call ValoresDoMesPagos
Else
   DataAuxiliarInicio = lblInicioT2
   DataAuxiliarFim = lblFimT2
   Call InverterData
   Call ValoresDoHistoricoPagos
End If

If Not ctp.EOF Then
   Call ProcessarTrimestrePagos
   AcumulaPagarT2 = AcumulaValor
   lblTotalPago2T = Format$(AcumulaValor, "###,##0.00")
   AcumulaValor = 0
End If

lblTotalPago2T.ForeColor = &HFF&

'Fim Trata Trimestre 2 a Receber

'Trata Trimestre 3 a Receber

If Month(lblInicioT3) > Month(Date) And Year(dtInicio) = Year(Date) Then
   AcumulaPagarT3 = 0
   lblTotalPago3T = Format$(0, "###,##0.00")
   AcumulaValor = 0
   Exit Sub
End If

If Month(lblInicioT3) = Month(Date) And Year(dtInicio) = Year(Date) Then
   Call ValoresDoMesPagos
Else
   DataAuxiliarInicio = lblInicioT3
   DataAuxiliarFim = lblFimT3
   Call InverterData
   Call ValoresDoHistoricoPagos
End If
If Not ctp.EOF Then
   Call ProcessarTrimestrePagos
   AcumulaPagarT3 = AcumulaValor
   lblTotalPago3T = Format$(AcumulaValor, "###,##0.00")
   AcumulaValor = 0
End If

lblTotalPago3T.ForeColor = &HFF&
'Fim Trata Trimestre 3 a Receber
End Sub

Public Sub ValoresDoMesPagos()

Call Rotina_AbrirBanco

ctp.Open "Select * from contas_a_pagar where ctpStatus = ('" & Status & "')", db, 3, 3
If ctp.EOF Then
   MsgBox ("Sem Pagamentos para tratar neste mes do trimestre."), vbInformation
   Call FechaDB
End If

End Sub

Public Sub ValoresDoHistoricoPagos()

Call Rotina_AbrirBanco

ctp.Open "Select * from historicocontaspagar where ctpDataPagamento > ('" & DataInvertidaInicio & "') and ctpDataPagamento < ('" & DataInvertidaFim & "')", db, 3, 3
If ctp.EOF Then
   MsgBox ("Trimestre no Historico sem valores pagos."), vbInformation
End If

End Sub

Public Sub ProcessarTrimestrePagos()

AcumulaValor = 0

ctp.MoveFirst

Do While Not ctp.EOF
   If ctp!ctpStatus = 1 Then
      AcumulaValor = AcumulaValor + ctp!ctpValorDaBoleta
   End If
   
   ctp.MoveNext
Loop

End Sub

Public Sub FinalizarConsulta()
 
 
SaldoT1 = AcumulaReceberT1 - AcumulaPagarT1
lblSaldo1T = Format$(AcumulaReceberT1 - AcumulaPagarT1, "###,###,##0.00")
If lblSaldo1T < 0 Then
   lblSaldo1T.ForeColor = &HFF&
Else
   lblSaldo1T.ForeColor = &HFF0000
End If
SaldoT2 = AcumulaReceberT2 - AcumulaPagarT2
lblSaldo2T = Format$(AcumulaReceberT2 - AcumulaPagarT2, "###,###,##0.00")
If lblSaldo2T < 0 Then
   lblSaldo2T.ForeColor = &HFF&
Else
   lblSaldo2T.ForeColor = &HFF0000
End If
SaldoT3 = AcumulaReceberT3 - AcumulaPagarT3
lblSaldo3T = Format$(AcumulaReceberT3 - AcumulaPagarT3, "###,###,##0.00")
If lblSaldo3T < 0 Then
   lblSaldo3T.ForeColor = &HFF&
Else
   lblSaldo3T.ForeColor = &HFF0000
End If

lblTotalT1 = Format$(AcumulaReceberT1 + AcumulaReceberT2 + AcumulaReceberT3, "###,###,#00.00")
lblTotalT1.ForeColor = &HFF0000

lblTotalT2 = Format$(AcumulaPagarT1 + AcumulaPagarT2 + AcumulaPagarT3, "###,###,#00.00")
lblTotalT2.ForeColor = &HFF&

lblTotalT3 = Format$(SaldoT1 + SaldoT2 + SaldoT3, "###,###,#00.00")
If lblTotalT3 < 0 Then
   lblTotalT3.ForeColor = &HFF&
Else
   lblTotalT3.ForeColor = &HFF0000
End If

lblMediaT1 = Format$((AcumulaReceberT1 + AcumulaReceberT2 + AcumulaReceberT3) / 3, "###,###,#00.00")
lblMediaT1.ForeColor = &HFF0000

lblMediaT2 = Format$((AcumulaPagarT1 + AcumulaPagarT2 + AcumulaPagarT3) / 3, "###,###,#00.00")
lblMediaT2.ForeColor = &HFF&

lblMediaT3 = Format$((SaldoT1 + SaldoT2 + SaldoT3) / 3, "###,###,#00.00")
If lblMediaT3 < 0 Then
   lblMediaT3.ForeColor = &HFF&
Else
   lblMediaT3.ForeColor = &HFF0000
End If

End Sub


Private Sub txtPercentSaldoMedio_LostFocus()

If txtPercentSaldoMedio = Empty Then
  ' MsgBox ("Não informado o percentual para Cálculo."), vbInformation
   Exit Sub
Else
   AcumulaValor = ((SaldoT1 + SaldoT2 + SaldoT3) * txtPercentSaldoMedio) / 100
   
   lblCalcSobreMedia = Format$(AcumulaValor, "###,###,##0.00")
End If

If lblCalcSobreMedia < 0 Then
   lblCalcSobreMedia.ForeColor = &HFF&
Else
   lblCalcSobreMedia.ForeColor = &HFF0000
End If


End Sub

Public Sub LimparConsulta()
lblInicioT1 = Empty
lblFimT1 = Empty
lblInicioT2 = Empty
lblFimT2 = Empty
lblInicioT3 = Empty
lblFimT3 = Empty
lblTotalRecebido1T = Empty
lblTotalPago1T = Empty
lblSaldo1T = Empty
lblTotalRecebido2T = Empty
lblTotalPago2T = Empty
lblSaldo2T = Empty
lblTotalRecebido3T = Empty
lblTotalPago3T = Empty
lblSaldo3T = Empty
lblTotalT1 = Empty
lblTotalT2 = Empty
lblTotalT3 = Empty
lblMediaT1 = Empty
lblMediaT2 = Empty
lblMediaT3 = Empty
txtPercentSaldoMedio = Empty
lblCalcSobreMedia = Empty
End Sub
