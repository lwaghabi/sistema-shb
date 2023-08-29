VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLancFinanc 
   Caption         =   "Lançamentos Financeiros"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox GridPagar 
      Height          =   4095
      Left            =   5880
      ScaleHeight     =   4035
      ScaleWidth      =   5715
      TabIndex        =   27
      Top             =   1200
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Pesquisada"
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
      Left            =   8040
      TabIndex        =   25
      Top             =   0
      Width           =   1815
      Begin MSMask.MaskEdBox txtDataConsulta 
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
   End
   Begin VB.Frame Frame11 
      Caption         =   "Hoje -"
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
      Left            =   5880
      TabIndex        =   22
      Top             =   0
      Width           =   2175
      Begin VB.TextBox txtDMaisMenos 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Dias Úteis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   24
         Top             =   120
         Width           =   615
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
      Left            =   9840
      TabIndex        =   20
      Top             =   0
      Width           =   1815
      Begin MSMask.MaskEdBox txtDataHoje 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   720
      Width           =   5775
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Contas a Receber"
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
         Left            =   1560
         TabIndex        =   16
         Top             =   120
         Width           =   2520
      End
   End
   Begin VB.Frame Frame4 
      Height          =   495
      Left            =   5880
      TabIndex        =   13
      Top             =   720
      Width           =   5775
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Contas a Pagar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   360
         Left            =   1800
         TabIndex        =   14
         Top             =   120
         Width           =   2130
      End
   End
   Begin VB.Frame Frame5 
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   5280
      Width           =   5775
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Detalhe da Operação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   1200
         TabIndex        =   12
         Top             =   120
         Width           =   2985
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Resumo Financeiro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5880
      TabIndex        =   4
      Top             =   5280
      Width           =   5775
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total de Lançamentos a RECEBER..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   3180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Total de Lançamentos a PAGAR......."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   3180
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "SALDO dos Lançamentos na data....."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   3180
      End
      Begin VB.Label lblSaldo 
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
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         Top             =   1200
         Width           =   2055
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
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblTotalReceber 
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
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   240
         Width           =   2055
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
      Left            =   4080
      TabIndex        =   2
      Top             =   0
      Width           =   1695
      Begin VB.ComboBox cmbFiltro 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   0
      Top             =   7080
      Width           =   5775
      Begin VB.CommandButton cmdSair 
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
         Left            =   4320
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.PictureBox GridReceber 
      Height          =   4095
      Left            =   0
      ScaleHeight     =   4035
      ScaleWidth      =   5715
      TabIndex        =   17
      Top             =   1200
      Width           =   5775
   End
   Begin VB.PictureBox GridDetalhe 
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   5715
      TabIndex        =   18
      Top             =   5760
      Width           =   5775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lançamentos Financeiros"
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
      Left            =   240
      TabIndex        =   19
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmLancFinanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim IndReceber As Byte
Dim IndDetalhe As Byte
Dim IndPagar As Byte

Dim IndConf As Byte
Dim AcumulaCtaReceber As Currency
Dim AcumulaCtaPagar As Currency
Dim AcumulaCtaAtraso As Currency

Dim DiaUtilAnterior As Date
Dim DataInicio As Date
Dim DataFim As Date
Dim DataReceber As Date
Dim DataPagos As Date
Dim DataAtrasados As Date
Dim DiaDaSemana As Integer

Dim Dia As String
Dim Mes As String
Dim Ano As String

Dim DataInvertida As String

Dim indice As Byte
Dim DataAcesso As Date

Private Sub cmbFiltro_lostfocus()

txtDataHoje = Date
DataInformada = Date

AcumulaCtaReceber = 0
AcumulaCtaPagar = 0
AcumulaCtaAtraso = 0

Call Rotina_010_Limpa_Cta_Pagar

Call Rotina_011_Limpa_Detalhe

Call Rotina_012_Limpa_Cta_Receber

Call Rotina_020_Gerencia_Grid

Call Rotina_021_Gerencia_Grid_Pagar

End Sub


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

txtDataHoje = Date
txtDataConsulta = Date

AcumulaCtaReceber = 0
AcumulaCtaPagar = 0

cmbFiltro.AddItem "Geral"

tabbanco.MoveFirst

Do While Not tabbanco.EOF
   cmbFiltro.AddItem tabbanco("bcosiglabco")
   tabbanco.MoveNext
Loop

cmbFiltro.ListIndex = 0

Call Rotina_010_Limpa_Cta_Pagar

Call Rotina_011_Limpa_Detalhe

Call Rotina_012_Limpa_Cta_Receber

Call Rotina_020_Gerencia_Grid

Call Rotina_021_Gerencia_Grid_Pagar

End Sub

Public Sub Rotina_010_Limpa_Cta_Pagar()
For IndReceber = 1 To 49
    GridPagar.TextMatrix(IndPagar, 0) = Empty
    GridPagar.TextMatrix(IndPagar, 1) = Empty
    GridPagar.TextMatrix(IndPagar, 2) = Empty
    GridPagar.TextMatrix(IndPagar, 3) = Empty
    
Next
End Sub
Public Sub Rotina_011_Limpa_Detalhe()
For IndDetalhe = 1 To 49
    GridDetalhe.TextMatrix(IndDetalhe, 0) = Empty
    GridDetalhe.TextMatrix(IndDetalhe, 1) = Empty
    GridDetalhe.TextMatrix(IndDetalhe, 2) = Empty
    GridDetalhe.TextMatrix(IndDetalhe, 3) = Empty
    GridDetalhe.TextMatrix(IndDetalhe, 3) = Empty
Next
End Sub
Public Sub Rotina_012_Limpa_Cta_Receber()
For IndConf = 1 To 49
    GridReceber.TextMatrix(IndConf, 0) = Empty
    GridReceber.TextMatrix(IndConf, 1) = Empty
    GridReceber.TextMatrix(IndConf, 2) = Empty
    GridReceber.TextMatrix(IndConf, 3) = Empty
    GridReceber.TextMatrix(IndConf, 4) = Empty
    GridReceber.TextMatrix(IndConf, 5) = Empty
Next
End Sub

Public Sub Rotina_020_Gerencia_Grid()

IndReceber = 0
IndDetalhe = 0

TabNegociacao.MoveFirst

Do While Not TabNegociacao.EOF
   
   indice = cmbFiltro.ListIndex
   
   If (cmbFiltro = "Geral") Or ((indice - 1) = TabNegociacao("chcodbcolart")) Then
      If TabNegociacao("negdatanegociação") = txtDataConsulta Then
         If TabNegociacao("negstatus") = 1 Then
            Call Rotina_051_Carga_Cta_Receber
         End If
      End If
  
  End If
   
   TabNegociacao.MoveNext

Loop

lblTotalReceber = Format$(AcumulaCtaReceber, "##,##0.00")
'lblTotalAtraso = Format$(AcumulaCtaAtraso, "##,##0.00")

'GridReceber.Col = 4
'GridReceber.ColSel = 4
     
'GridReceber.Row = 1
'GridReceber.RowSel = IndReceber
        
'If IndReceber > 1 Then
'   GridReceber.Sort = 1
'End If

'GridReceber.Col = 0
'GridReceber.ColSel = 0
'GridReceber.Row = 0
'GridReceber.RowSel = 0

'GridDetalhe.Col = 4
'GridDetalhe.ColSel = 4
     
'GridDetalhe.Row = 1
'GridDetalhe.RowSel = IndDetalhe
        
'If IndDetalhe > 1 Then
'   GridDetalhe.Sort = 1
'End If

'GridDetalhe.Col = 0
'GridDetalhe.ColSel = 0
'GridDetalhe.Row = 0
'GridDetalhe.RowSel = 0

End Sub
Public Sub Rotina_021_Gerencia_Grid_Pagar()

IndPagar = 0

TabNotaFiscalEntrada.MoveFirst

Do While Not TabNotaFiscalEntrada.EOF
   
   indice = cmbFiltro.ListIndex
   
   If (cmbFiltro = "Geral") Or ((indice - 1) = TabNotaFiscalEntrada("chcodbcolart")) Then
      If TabNotaFiscalEntrada("nfeDataLanc") = txtDataConsulta Then
          Call Rotina_050_Carga_Cta_Pagar
      End If
   End If
   
   TabNotaFiscalEntrada.MoveNext

Loop

'lblTotalPagar = Format$(AcumulaCtaPagar, "##,##0.00")

'lblSaldo = Format$(AcumulaCtaReceber - AcumulaCtaPagar, "##,##0.00")
           
'If lblSaldo < 0 Then
'   lblSaldo.ForeColor = vbRed
'Else
'   lblSaldo.ForeColor = vbBlue
'End If

'GridPagar.Col = 4
'GridPagar.ColSel = 4
     
'GridPagar.Row = 1
'GridPagar.RowSel = IndPagar
        
'If IndPagar > 1 Then
'   GridPagar.Sort = 1
'End If

'GridPagar.Col = 0
'GridPagar.ColSel = 0
'GridPagar.Row = 0
'GridPagar.RowSel = 0

End Sub

Public Sub Rotina_050_Carga_Cta_Pagar()

IndPagar = IndPagar + 1
GridPagar.TextMatrix(IndPagar, 0) = TabCtaPagar("chdatavencito")
GridPagar.TextMatrix(IndPagar, 1) = TabCtaPagar("chpessoa")
GridPagar.TextMatrix(IndPagar, 2) = TabCtaPagar("ctpdescricaooperacao")
GridPagar.TextMatrix(IndPagar, 3) = Format$(TabCtaPagar("ctpValordaboleta"), "##,##0.00")

Dia = Format$(Day(TabCtaPagar("chdatavencito")), "00")
Mes = Format$(Month(TabCtaPagar("chdatavencito")), "00")
Ano = Year(TabCtaPagar("chdatavencito"))

DataInvertida = Ano & Mes & Dia

GridPagar.TextMatrix(IndPagar, 4) = DataInvertida
AcumulaCtaPagar = AcumulaCtaPagar + TabCtaPagar("ctpValordaboleta")

End Sub
Public Sub Rotina_051_Carga_Cta_Receber()

IndReceber = IndReceber + 1
GridReceber.TextMatrix(IndReceber, 0) = TabCtaReceber("ctrdatavencito")
GridReceber.TextMatrix(IndReceber, 1) = TabCtaReceber("chpessoa")
GridReceber.TextMatrix(IndReceber, 2) = TabCtaReceber("ctrdescricaooperacao")
GridReceber.TextMatrix(IndReceber, 3) = Format$(TabCtaReceber("ctrValordaboleta"), "##,##0.00")

Dia = Format$(Day(TabCtaReceber("ctrdatavencito")), "00")
Mes = Format$(Month(TabCtaReceber("ctrdatavencito")), "00")
Ano = Year(TabCtaReceber("ctrdatavencito"))

DataInvertida = Ano & Mes & Dia

GridReceber.TextMatrix(IndReceber, 4) = DataInvertida

AcumulaCtaReceber = AcumulaCtaReceber + TabCtaReceber("ctrValordaboleta")
End Sub
Public Sub Rotina_052_Carga_Atrasados()

IndDetalhe = IndDetalhe + 1
GridDetalhe.TextMatrix(IndDetalhe, 0) = TabCtaReceber("ctrdatavencito")
GridDetalhe.TextMatrix(IndDetalhe, 1) = TabCtaReceber("chpessoa")
GridDetalhe.TextMatrix(IndDetalhe, 2) = TabCtaReceber("ctrdescricaooperacao")
GridDetalhe.TextMatrix(IndDetalhe, 3) = Format(TabCtaReceber("ctrValordaboleta"), "##,##0.00")

Dia = Format$(Day(TabCtaReceber("ctrdatavencito")), "00")
Mes = Format$(Month(TabCtaReceber("ctrdatavencito")), "00")
Ano = Year(TabCtaReceber("ctrdatavencito"))

DataInvertida = Ano & Mes & Dia

GridDetalhe.TextMatrix(IndDetalhe, 4) = DataInvertida

AcumulaCtaAtraso = AcumulaCtaAtraso + TabCtaReceber("ctrValordaboleta")

End Sub


Private Sub optDMais_Click()

NDias = txtDMaisMenos
DataInformada = Date

DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)

DataInformada = DataRetorno.DiaUtil

txtDataHoje = Date

Call Rotina_070_Ajusta_Data

AcumulaCtaReceber = 0
AcumulaCtaPagar = 0
AcumulaCtaAtraso = 0

Call Rotina_010_Limpa_Cta_Pagar

Call Rotina_012_Limpa_Cta_Receber

Call Rotina_020_Gerencia_Grid

Call Rotina_021_Gerencia_Grid_Pagar

txtDMaisMenos.SetFocus

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

Call Rotina_010_Limpa_Cta_Pagar

Call Rotina_012_Limpa_Cta_Receber

Call Rotina_020_Gerencia_Grid

Call Rotina_021_Gerencia_Grid_Pagar

End Sub

Public Sub Rotina_070_Ajusta_Data()

DiaDaSemana = Weekday(DataInformada)

'Calcular Range de datas

DataInicio = DataInformada - (DiaDaSemana + 1)
DataFim = DataInicio + 7

End Sub


