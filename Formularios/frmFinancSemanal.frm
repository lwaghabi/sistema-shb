VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConsolidadoFinanc 
   Caption         =   "Consolidado Financeiro Semanal"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Caption         =   "Posição Financeira Consolidada Anual"
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
      Left            =   6720
      TabIndex        =   39
      Top             =   6360
      Width           =   4575
      Begin VB.TextBox txtSaldoMensal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   49
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtTotalDebitoMensal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   47
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtTotalCreditoMensal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   45
         Top             =   360
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtDataDeMensal 
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDataAteMensal 
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo No Período....."
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
         Left            =   1200
         TabIndex        =   48
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Total a Débito.."
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
         Left            =   1200
         TabIndex        =   46
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Total a Crédito."
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
         Left            =   1200
         TabIndex        =   44
         Top             =   360
         Width           =   1890
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "De"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   210
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Até"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Posição Financeira Mensal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   6720
      TabIndex        =   37
      Top             =   1200
      Width           =   5055
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridMensal 
         Height          =   4815
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   8493
         _Version        =   393216
         Rows            =   100
         Cols            =   4
         FixedCols       =   0
         FormatString    =   "<Mes/Ano    |>Crédito            |>Débito             |>Saldo               "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
   End
   Begin VB.Frame Frame6 
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
      Left            =   6600
      TabIndex        =   34
      Top             =   360
      Width           =   2775
      Begin VB.ComboBox cmbFiltro 
         Height          =   315
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   32
      Top             =   360
      Width           =   6375
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Consolidado Financeiro"
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
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1815
      Left            =   11280
      TabIndex        =   30
      Top             =   6360
      Width           =   495
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H000000FF&
         Caption         =   "Sair"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   0
         MaskColor       =   &H000000FF&
         TabIndex        =   31
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
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
      Left            =   9480
      TabIndex        =   28
      Top             =   360
      Width           =   2295
      Begin MSMask.MaskEdBox txtDataHoje 
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
   Begin VB.Frame Frame2 
      Caption         =   "Posição Financeira Consolidada até a data"
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
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Width           =   6615
      Begin MSMask.MaskEdBox txtDataAte 
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDataDE 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "V  a  l  o  r"
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
         Left            =   2520
         TabIndex        =   36
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label txtSaldoEmAtraso 
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
         Left            =   4920
         TabIndex        =   27
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label txtPagtosEmAtraso 
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
         Left            =   3720
         TabIndex        =   26
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label txtRecebEmAtraso 
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
         Left            =   2520
         TabIndex        =   25
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Atrasados........."
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
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label txtSaldoTotal 
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
         Left            =   4920
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label txtTotalAPagar 
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
         Left            =   3720
         TabIndex        =   22
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label txtTotalAReceber 
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
         Left            =   2520
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total................"
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
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label txtSaldoProces 
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
         Left            =   4920
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label txtProcesAPagar 
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
         Left            =   3720
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label txtProcesAReceber 
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
         Left            =   2520
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proces. no mes."
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
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
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
         Left            =   4920
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Débito"
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
         Left            =   3720
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito"
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
         Left            =   2520
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label txtSaldoPendente 
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
         Left            =   4920
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label txtPendenteAPagar 
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
         Left            =   3720
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label txtPendenteAReceber 
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
         Left            =   2520
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pendentes........"
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
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSemana 
      Height          =   4815
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8493
      _Version        =   393216
      Rows            =   100
      Cols            =   5
      FixedCols       =   0
      FormatString    =   "De              |Até              |>Crédito        |>Débito          |>Saldo           "
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Frame Frame1 
      Caption         =   "Posição Financeira Semanal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6615
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "V a l o r"
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
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   120
      Picture         =   "frmFinancSemanal.frx":0000
      Top             =   0
      Width           =   2865
   End
End
Attribute VB_Name = "frmConsolidadoFinanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IndReceber As Byte
Dim IndPagar As Byte
Dim IndMes As Byte
Dim Indice As Byte

Dim fimsemana As Byte
Dim fimmensal As Byte
Dim UltimoMes As Date

Dim AcumProcesPagar As Currency
Dim AcumProcesPagarMensal As Currency
Dim AcumPendPagar As Currency
Dim AcumPendPagarMensal As Currency
Dim AcumPagarAtrasado As Currency
Dim AcumPagarAtrasadoMensal As Currency

Dim AcumPendReceber As Currency
Dim AcumPendReceberMensal As Currency
Dim AcumReceberAtrasado As Currency
Dim AcumReceberAtrasadoMensal As Currency
Dim AcumProcesReceber As Currency
Dim AcumProcesReceberMensal As Currency

Dim DiaUtilAnterior As Date
Dim DataInicio As Date
Dim DataFim As Date
Dim DataParaCalculo As Date

Dim DataMensalInicio As Date
Dim DataMensalFim As Date

Dim dia As Integer
Dim mes As Integer
Dim ano As Integer

Dim DiaDaSemana As Integer

Dim tabDataIni(100) As Date
Dim tabDataFim(100) As Date
Dim TabMesAnoIni(100) As Date
Dim TabMesAnoFim(100) As Date
Dim tabValor(100, 3) As Currency
Dim tabValorMensal(100, 3) As Currency

Private Sub cmbFiltro_lostfocus()

Indice = cmbFiltro.ListIndex

Call Rotina_00_Principal

cmdSair.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

cmbFiltro.Clear
cmbFiltro.AddItem "Geral"

TabEmpresa.MoveLast

Do While Not TabEmpresa.BOF
   cmbFiltro.AddItem TabEmpresa("chPessoa")
   TabEmpresa.MovePrevious
Loop
cmbFiltro.ListIndex = 0
Call Rotina_00_Principal

End Sub

Public Sub Rotina_00_Principal()

txtDataHoje = Date

DataInformada = Date

NDias = 2
DataInformada = Date
DataRetorno = ObterDiaUtilAnterior(DataInformada, NDias)

DiaUtilAnterior = DataRetorno.DiaUtil

Call Rotina_070_Ajusta_Data

txtDataDE = DataInicio
txtDataAte = DataFim

Call Rotina_012_Limpa_Cta_Receber

tabDataIni(1) = DataInicio
tabDataFim(1) = DataFim
TabMesAnoIni(1) = DataMensalInicio
TabMesAnoFim(1) = DataMensalFim

AcumProcesPagar = 0
AcumProcesPagarMensal = 0
AcumPendPagar = 0
AcumPendPagarMensal = 0
AcumPagarAtrasado = 0
AcumPagarAtrasadoMensal = 0

AcumPendReceber = 0
AcumPendReceberMensal = 0
AcumReceberAtrasado = 0
AcumReceberAtrasadoMensal = 0
AcumProcesReceber = 0
AcumProcesReceberMensal = 0

TabCtaReceber.MoveFirst

Do While Not TabCtaReceber.EOF
   
   IndReceber = 1
   IndMes = 1
   
   If (cmbFiltro = "Geral") Or ((Indice - 1) = TabCtaReceber("chfabricante")) Then
   
      Call Rotina_030_TabSemana
      
      Call Rotina_040_TabMensal
   End If
      
   TabCtaReceber.MoveNext
      
Loop


TabCtaPagar.MoveFirst

Do While Not TabCtaPagar.EOF
   
   IndPagar = 1
   IndMes = 1
      
   If (cmbFiltro = "Geral") Or ((Indice - 1) = TabCtaPagar("chfabricante")) Then
      
      Call Rotina_050_CtaPagar_Semana
     
      Call Rotina_060_CtaPagar_Mes

   End If

   TabCtaPagar.MoveNext
   
Loop

IndReceber = 1

Do While Not (tabDataIni(IndReceber) = Empty)
   GridSemana.TextMatrix(IndReceber, 0) = tabDataIni(IndReceber)
   GridSemana.TextMatrix(IndReceber, 1) = tabDataFim(IndReceber)
   GridSemana.TextMatrix(IndReceber, 2) = Format$(tabValor(IndReceber, 0), "##,##0.00")
   GridSemana.TextMatrix(IndReceber, 3) = Format$(tabValor(IndReceber, 1), "##,##0.00")
   GridSemana.TextMatrix(IndReceber, 4) = Format$(tabValor(IndReceber, 2), "##,##0.00")
   
   GridSemana.Col = 0
   GridSemana.ColSel = 0
   GridSemana.Row = IndReceber
   GridSemana.RowSel = IndReceber
   GridSemana.CellForeColor = RGB(256, 51, 0)
   GridSemana.CellFontBold = True
     
   GridSemana.Col = 1
   GridSemana.ColSel = 1
   GridSemana.Row = IndReceber
   GridSemana.RowSel = IndReceber
   GridSemana.CellForeColor = RGB(256, 51, 0)
   GridSemana.CellFontBold = True
   
   GridSemana.Col = 2
   GridSemana.ColSel = 2
   GridSemana.Row = IndReceber
   GridSemana.RowSel = IndReceber
   GridSemana.CellForeColor = RGB(51, 0, 256)
   GridSemana.CellFontBold = True
     
   GridSemana.Col = 3
   GridSemana.ColSel = 3
   GridSemana.Row = IndReceber
   GridSemana.RowSel = IndReceber
   GridSemana.CellForeColor = RGB(256, 121, 0)
   GridSemana.CellFontBold = True
   
   If tabValor(IndReceber, 2) < 0 Then
      GridSemana.Col = 4
      GridSemana.ColSel = 4
      GridSemana.Row = IndReceber
      GridSemana.RowSel = IndReceber
      GridSemana.CellForeColor = vbRed
      GridSemana.CellFontBold = True
   Else
      GridSemana.Col = 4
      GridSemana.ColSel = 4
      GridSemana.Row = IndReceber
      GridSemana.RowSel = IndReceber
      GridSemana.CellForeColor = vbBlue
   End If
   IndReceber = IndReceber + 1
Loop

txtPendenteAReceber = Format$(AcumPendReceber, "##,##0.00")
txtPendenteAReceber.ForeColor = RGB(51, 0, 256)
txtPendenteAPagar = Format$(AcumPendPagar, "##,##0.00")
txtPendenteAPagar.ForeColor = RGB(256, 121, 0)
txtSaldoPendente = Format$(AcumPendReceber - AcumPendPagar, "##,##0.00")
If AcumPendPagar > AcumPendReceber Then
   txtSaldoPendente.ForeColor = vbRed
Else
   txtSaldoPendente.ForeColor = vbBlue
End If

txtProcesAReceber = Format$(AcumProcesReceber, "##,##0.00")
txtProcesAReceber.ForeColor = RGB(51, 0, 256)
txtProcesAPagar = Format$(AcumProcesPagar, "##,##0.00")
txtProcesAPagar.ForeColor = RGB(256, 121, 0)
txtSaldoProces = Format$(AcumProcesReceber - AcumProcesPagar, "##,##0.00")
If AcumProcesPagar > AcumProcesReceber Then
   txtSaldoProces.ForeColor = vbRed
Else
   txtSaldoProces.ForeColor = vbBlue
End If

txtTotalAReceber = Format$(AcumPendReceber + AcumProcesReceber, "##,##0.00")
txtTotalAReceber.ForeColor = RGB(51, 0, 256)
txtTotalAPagar = Format$(AcumPendPagar + AcumProcesPagar, "##,##0.00")
txtTotalAPagar.ForeColor = RGB(256, 121, 0)
txtSaldoTotal = Format$((AcumPendReceber + AcumProcesReceber) - (AcumPendPagar + AcumProcesPagar), "##,##0.00")
If (AcumProcesPagar + AcumPendPagar) > (AcumProcesReceber + AcumPendReceber) Then
   txtSaldoTotal.ForeColor = vbRed
Else
   txtSaldoTotal.ForeColor = vbBlue
End If

txtRecebEmAtraso = Format$(AcumReceberAtrasado, "##,##0.00")
txtPagtosEmAtraso = Format$(AcumPagarAtrasado, "##,##0.00")
txtSaldoEmAtraso = Format$(AcumReceberAtrasado - AcumPagarAtrasado, "##,##0.00")


'MENSAL
IndMes = 1

Do While Not (TabMesAnoIni(IndMes) = Empty)
   GridMensal.TextMatrix(IndMes, 0) = Format$(TabMesAnoIni(IndMes), "mmm/yyyy")
   GridMensal.TextMatrix(IndMes, 1) = Format$(tabValorMensal(IndMes, 0), "##,##0.00")
   GridMensal.TextMatrix(IndMes, 2) = Format$(tabValorMensal(IndMes, 1), "##,##0.00")
   GridMensal.TextMatrix(IndMes, 3) = Format$(tabValorMensal(IndMes, 2), "##,##0.00")
   
   GridMensal.Col = 0
   GridMensal.ColSel = 0
   GridMensal.Row = IndMes
   GridMensal.RowSel = IndMes
   GridMensal.CellForeColor = RGB(256, 51, 0)
   GridMensal.CellFontBold = True
   
   GridMensal.Col = 1
   GridMensal.ColSel = 1
   GridMensal.Row = IndMes
   GridMensal.RowSel = IndMes
   GridMensal.CellForeColor = RGB(51, 0, 256)
   GridMensal.CellFontBold = True
     
   GridMensal.Col = 2
   GridMensal.ColSel = 2
   GridMensal.Row = IndMes
   GridMensal.RowSel = IndMes
   GridMensal.CellForeColor = RGB(256, 121, 0)
   GridMensal.CellFontBold = True
   
   If tabValor(IndMes, 2) < 0 Then
      GridMensal.Col = 3
      GridMensal.ColSel = 3
      GridMensal.Row = IndMes
      GridMensal.RowSel = IndMes
      GridMensal.CellForeColor = vbRed
      GridMensal.CellFontBold = True
   Else
      GridMensal.Col = 3
      GridMensal.ColSel = 3
      GridMensal.Row = IndMes
      GridMensal.RowSel = IndMes
      GridMensal.CellForeColor = vbBlue
      GridMensal.CellFontBold = True
   End If
   IndMes = IndMes + 1
Loop

txtDataDeMensal = Format$(TabMesAnoIni(1), "mmm/yyyy")
txtDataAteMensal = Format$(UltimoMes, "mmm/yyyy")

txtTotalCreditoMensal = Format$(AcumPendReceberMensal, "##,##0.00")
txtTotalCreditoMensal.ForeColor = RGB(51, 0, 256)
txtTotalDebitoMensal = Format$(AcumPendPagarMensal, "##,##0.00")
txtTotalDebitoMensal.ForeColor = RGB(256, 121, 0)
txtSaldoMensal = Format$(AcumPendReceberMensal - AcumPendPagarMensal, "##,##0.00")
If AcumPendPagarMensal > AcumPendReceberMensal Then
   txtSaldoMensal.ForeColor = vbRed
Else
   txtSaldoMensal.ForeColor = vbBlue
End If

End Sub

Public Sub Rotina_070_Ajusta_Data()

DiaDaSemana = Weekday(DataInformada)

'Calcular Range de datas

DataInicio = DataInformada - (DiaDaSemana)
DataFim = DataInicio + 6

dia = 1
mes = Month(DataInicio)
ano = Year(DataInicio)

DataMensalInicio = dia & "/" & mes & "/" & ano

mes = mes + 1

DataMensalFim = (dia & "/" & mes & "/" & ano)

DataMensalFim = DataMensalFim - 1

End Sub

Public Sub Rotina_012_Limpa_Cta_Receber()
For IndReceber = 1 To 99
    
    tabDataIni(IndReceber) = Empty
    tabDataFim(IndReceber) = Empty
    TabMesAnoIni(IndReceber) = Empty
    TabMesAnoFim(IndReceber) = Empty
    tabValor(IndReceber, 0) = 0
    tabValor(IndReceber, 1) = 0
    tabValor(IndReceber, 2) = 0
    tabValorMensal(IndReceber, 0) = 0
    tabValorMensal(IndReceber, 1) = 0
    tabValorMensal(IndReceber, 2) = 0
    GridSemana.TextMatrix(IndReceber, 0) = Empty
    GridSemana.TextMatrix(IndReceber, 1) = Empty
    GridSemana.TextMatrix(IndReceber, 2) = Empty
    GridSemana.TextMatrix(IndReceber, 3) = Empty
    GridSemana.TextMatrix(IndReceber, 4) = Empty
    GridMensal.TextMatrix(IndReceber, 0) = Empty
    GridMensal.TextMatrix(IndReceber, 1) = Empty
    GridMensal.TextMatrix(IndReceber, 2) = Empty
    GridMensal.TextMatrix(IndReceber, 3) = Empty
    
Next
End Sub
Public Sub Rotina_020_Acumula_Receber()

If TabCtaReceber("ctrstatus") = 0 Then
   If TabCtaReceber("ctrdatavencito") < DiaUtilAnterior + 1 Then
      AcumReceberAtrasado = AcumReceberAtrasado + TabCtaReceber("ctrvalordaboleta")
   Else
      If TabCtaReceber("ctrdatavencito") < tabDataFim(1) + 1 Then
         AcumPendReceber = AcumPendReceber + TabCtaReceber("ctrvalordaboleta")
      End If
   End If
Else
   AcumProcesReceber = AcumProcesReceber + TabCtaReceber("ctrvalordaboleta")
End If

End Sub

Public Sub Rotina_025_Acumula_Pagar()
If TabCtaPagar("ctpstatus") = 0 Then
   If TabCtaPagar("chdatavencito") < DiaUtilAnterior + 1 Then
      AcumPagarAtrasado = AcumPagarAtrasado + TabCtaPagar("ctpvalordaboleta")
   Else
      If TabCtaPagar("chdatavencito") < tabDataFim(1) + 1 Then
         AcumPendPagar = AcumPendPagar + TabCtaPagar("ctpvalordaboleta")
      End If
   End If
Else
   AcumProcesPagar = AcumProcesPagar + TabCtaPagar("ctpvalordaboleta")
End If

End Sub

Public Sub Rotina_026_Acumula_Mensal()

   If TabCtaReceber("ctrdatavencito") < DiaUtilAnterior + 1 Then
      AcumReceberAtrasadoMensal = AcumReceberAtrasadoMensal + TabCtaReceber("ctrvalordaboleta")
   Else
      AcumPendReceberMensal = AcumPendReceberMensal + TabCtaReceber("ctrvalordaboleta")
   End If

End Sub
Public Sub Rotina_027_Acumula_Mensal_Pagar()

   If TabCtaPagar("chdatavencito") < DiaUtilAnterior + 1 Then
      AcumPagarAtrasadoMensal = AcumPagarAtrasadoMensal + TabCtaPagar("ctpvalordaboleta")
   Else
      AcumPendPagarMensal = AcumPendPagarMensal + TabCtaPagar("ctpvalordaboleta")
   End If

End Sub

Public Sub Rotina_030_TabSemana()

fimsemana = 0

Do While fimsemana = 0

If tabDataFim(IndReceber) = Empty Then
   tabDataIni(IndReceber) = tabDataFim(IndReceber - 1) + 1
   tabDataFim(IndReceber) = tabDataIni(IndReceber) + 6
   tabValor(IndReceber, 0) = Format$(0#, "#,##0.00")
   tabValor(IndReceber, 1) = Format$(0#, "#,##0.00")
   tabValor(IndReceber, 2) = Format$(0#, "#,##0.00")
End If

If TabCtaReceber("ctrdatavencito") > DataInicio - 1 Then
      If TabCtaReceber("ctrdatavencito") > tabDataFim(IndReceber) Then
         IndReceber = IndReceber + 1
      Else
         If TabCtaReceber("ctrstatus") = 0 And TabCtaReceber("ctrdatavencito") > DiaUtilAnterior Then
            tabValor(IndReceber, 0) = Format$(tabValor(IndReceber, 0) + TabCtaReceber("ctrvalordaboleta"), "#,##0.00")
         End If
         Call Rotina_020_Acumula_Receber
         tabValor(IndReceber, 2) = tabValor(IndReceber, 0) - tabValor(IndReceber, 1)
         IndReceber = 1
         fimsemana = 1
      End If
   Else
      Call Rotina_020_Acumula_Receber
      
      fimsemana = 1
      IndReceber = 1
   End If
Loop
End Sub

Public Sub Rotina_040_TabMensal()

fimmensal = 0

Do While fimmensal = 0

If TabCtaReceber("ctrdatavencito") > DataMensalInicio - 1 Then

   If TabMesAnoFim(IndMes) = Empty Then
         TabMesAnoIni(IndMes) = TabMesAnoFim(IndMes - 1) + 1
         ano = Year(TabMesAnoIni(IndMes))
         mes = Month(TabMesAnoIni(IndMes))
         mes = mes + 1
         If mes = 13 Then
            mes = 1
            ano = ano + 1
         End If
         dia = 1
         DataParaCalculo = (dia & "/" & mes & "/" & ano)
         UltimoMes = DataParaCalculo - 1
         TabMesAnoFim(IndMes) = DataParaCalculo - 1
         tabValorMensal(IndMes, 0) = Format$(0#, "#,##0.00")
         tabValorMensal(IndMes, 1) = Format$(0#, "#,##0.00")
         tabValorMensal(IndMes, 2) = Format$(0#, "#,##0.00")
    End If
End If

If TabCtaReceber("ctrdatavencito") > DataMensalInicio - 1 Then
   If TabCtaReceber("ctrdatavencito") > TabMesAnoFim(IndMes) Then
      IndMes = IndMes + 1
   Else
      If TabCtaReceber("ctrstatus") = 0 And TabCtaReceber("ctrdatavencito") > DiaUtilAnterior Then
         tabValorMensal(IndMes, 0) = Format$(tabValorMensal(IndMes, 0) + TabCtaReceber("ctrvalordaboleta"), "#,##0.00")
      End If
      Call Rotina_026_Acumula_Mensal
      tabValorMensal(IndMes, 2) = tabValorMensal(IndMes, 0) - tabValorMensal(IndMes, 1)
      IndMes = 1
      fimmensal = 1
   End If
Else
   Call Rotina_026_Acumula_Mensal
   fimmensal = 1
   IndMes = 1
End If

Loop

End Sub

Public Sub Rotina_050_CtaPagar_Semana()

fimsemana = 0

Do While fimsemana = 0

If tabDataFim(IndPagar) = Empty Then
   tabDataIni(IndPagar) = tabDataFim(IndPagar - 1) + 1
   tabDataFim(IndPagar) = tabDataIni(IndPagar) + 6
   tabValor(IndPagar, 0) = Format$(0#, "#,##0.00")
   tabValor(IndPagar, 1) = Format$(0#, "#,##0.00")
   tabValor(IndPagar, 2) = Format$(0#, "#,##0.00")
End If

If TabCtaPagar("chdatavencito") > DataInicio - 1 Then
      If TabCtaPagar("chdatavencito") > tabDataFim(IndPagar) Then
         IndPagar = IndPagar + 1
      Else
         If TabCtaPagar("ctpstatus") = 0 And TabCtaPagar("chdatavencito") > DiaUtilAnterior - 1 Then
            tabValor(IndPagar, 1) = Format$(tabValor(IndPagar, 1) + TabCtaPagar("ctpvalordaboleta"), "#,##0.00")
         End If
         Call Rotina_025_Acumula_Pagar
         tabValor(IndPagar, 2) = tabValor(IndPagar, 0) - tabValor(IndPagar, 1)
         IndPagar = 1
         fimsemana = 1
      End If
   Else
      Call Rotina_025_Acumula_Pagar
      
      fimsemana = 1
      IndPagar = 1
   End If
Loop
End Sub

Public Sub Rotina_060_CtaPagar_Mes()
fimmensal = 0

Do While fimmensal = 0

If TabCtaPagar("chdatavencito") > DataMensalInicio - 1 Then

   If TabMesAnoFim(IndMes) = Empty Then
         TabMesAnoIni(IndMes) = TabMesAnoFim(IndMes - 1) + 1
         ano = Year(TabMesAnoIni(IndMes))
         mes = Month(TabMesAnoIni(IndMes))
         mes = mes + 1
         If mes = 13 Then
            mes = 1
            ano = ano + 1
         End If
         dia = 1
         DataParaCalculo = (dia & "/" & mes & "/" & ano)
         TabMesAnoFim(IndMes) = DataParaCalculo - 1
         tabValorMensal(IndMes, 0) = Format$(0#, "#,##0.00")
         tabValorMensal(IndMes, 1) = Format$(0#, "#,##0.00")
         tabValorMensal(IndMes, 2) = Format$(0#, "#,##0.00")
    End If
End If

If TabCtaPagar("chdatavencito") > DataMensalInicio - 1 Then
   If TabCtaPagar("chdatavencito") > TabMesAnoFim(IndMes) Then
      IndMes = IndMes + 1
   Else
      If TabCtaPagar("ctpstatus") = 0 And TabCtaPagar("chdatavencito") > DiaUtilAnterior Then
         tabValorMensal(IndMes, 1) = Format$(tabValorMensal(IndMes, 1) + TabCtaPagar("ctpvalordaboleta"), "#,##0.00")
      End If
      Call Rotina_027_Acumula_Mensal_Pagar
      tabValorMensal(IndMes, 2) = tabValorMensal(IndMes, 0) - tabValorMensal(IndMes, 1)
      IndMes = 1
      fimmensal = 1
   End If
Else
   Call Rotina_027_Acumula_Mensal_Pagar
   fimmensal = 1
   IndMes = 1
End If

Loop

End Sub


