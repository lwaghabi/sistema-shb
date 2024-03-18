VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPedido 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Registro e Processamento de Medições"
   ClientHeight    =   10710
   ClientLeft      =   13170
   ClientTop       =   -345
   ClientWidth     =   20370
   LinkTopic       =   "Form3"
   ScaleHeight     =   10710
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Status Medição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   102
      Top             =   8160
      Width           =   5655
      Begin VB.TextBox txtTotalMedicao 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3560
         TabIndex        =   104
         Top             =   1800
         Width           =   1480
      End
      Begin MSFlexGridLib.MSFlexGrid gridMedicao 
         Height          =   1575
         Left            =   240
         TabIndex        =   103
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColor       =   16777152
         BackColorFixed  =   16776960
         BackColorBkg    =   16777152
         FormatString    =   " Localização  |Medição |Comp| Valor Total   "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblLabel11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   2760
         TabIndex        =   105
         Top             =   1800
         Width           =   675
      End
   End
   Begin VB.ComboBox CFOPAux 
      Height          =   315
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   94
      Top             =   10320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid GridPedido 
      Height          =   5895
      Left            =   15840
      TabIndex        =   51
      Top             =   2240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   10398
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   16777152
      BackColorFixed  =   16776960
      BackColorBkg    =   16777152
      FormatString    =   "Med.    |C|Data      |Cliente          |"
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
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   0
      TabIndex        =   82
      Top             =   0
      Width           =   15855
      Begin MSComCtl2.DTPicker dtFimMedicao 
         Height          =   420
         Left            =   14040
         TabIndex        =   6
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   741
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         Format          =   242089985
         CurrentDate     =   44656
      End
      Begin MSComCtl2.DTPicker dtInicioMedicao 
         Height          =   420
         Left            =   11760
         TabIndex        =   5
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   741
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   242089985
         CurrentDate     =   44656
      End
      Begin VB.ComboBox cmbContrato 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7800
         TabIndex        =   4
         Top             =   360
         Width           =   3255
      End
      Begin VB.ComboBox cmbLocal 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7800
         TabIndex        =   7
         Top             =   1200
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker txtDataProces 
         Height          =   375
         Left            =   4200
         TabIndex        =   44
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   242089985
         CurrentDate     =   43969
      End
      Begin MSComCtl2.DTPicker txtDataProc 
         Height          =   375
         Left            =   4320
         TabIndex        =   92
         Top             =   6000
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   242155521
         CurrentDate     =   43967
      End
      Begin MSComCtl2.DTPicker txtDataPedido 
         Height          =   435
         Left            =   2040
         TabIndex        =   2
         Top             =   330
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   767
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   12648447
         CalendarForeColor=   0
         Format          =   242155521
         CurrentDate     =   43902
      End
      Begin MSMask.MaskEdBox txtCNPJCPF 
         Height          =   390
         Left            =   4200
         TabIndex        =   41
         Tag             =   "4"
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   688
         _Version        =   393216
         BackColor       =   16777088
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtNumPedido 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbPessoa 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   4200
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtComplementoPedido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtRepresentante 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   11280
         TabIndex        =   42
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtPromotora 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   13560
         TabIndex        =   43
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "De"
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
         Left            =   11280
         TabIndex        =   114
         Top             =   405
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Label17"
         Height          =   495
         Left            =   8640
         TabIndex        =   113
         Top             =   360
         Width           =   15
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "até"
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
         Left            =   13560
         TabIndex        =   112
         Top             =   405
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   11760
         TabIndex        =   111
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label lblLabel19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   7800
         TabIndex        =   109
         Top             =   120
         Width           =   885
      End
      Begin VB.Label lblLabel13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Promotor"
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
         Left            =   13440
         TabIndex        =   107
         Top             =   960
         Width           =   765
      End
      Begin VB.Label lblLabel10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   7800
         TabIndex        =   96
         Top             =   960
         Width           =   585
      End
      Begin VB.Label lblLabel10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Representante"
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
         Index           =   2
         Left            =   11280
         TabIndex        =   95
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label txtStatusPedido 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   510
         Left            =   240
         TabIndex        =   90
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   120
         TabIndex        =   88
         Top             =   120
         Width           =   900
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Comp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   87
         Top             =   90
         Width           =   615
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Reg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2160
         TabIndex        =   86
         Top             =   60
         Width           =   1005
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4200
         TabIndex        =   85
         Top             =   90
         Width           =   705
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1680
         TabIndex        =   84
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ/CPF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4200
         TabIndex        =   83
         Top             =   960
         Width           =   1110
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Controles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   15840
      TabIndex        =   80
      Top             =   8160
      Width           =   4575
      Begin VB.CommandButton cmdProcessar 
         BackColor       =   &H000000FF&
         Caption         =   "&Processar Medição"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   1800
      End
      Begin VB.CommandButton cmdExcluiPedido 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1800
      End
      Begin VB.CommandButton cmdAlteraPedido 
         BackColor       =   &H0000FF00&
         Caption         =   "&Alterar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H00C0C000&
         Caption         =   "&Sair"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1440
         Width           =   1935
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Endereço"
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
      TabIndex        =   71
      Top             =   2160
      Width           =   15855
      Begin VB.TextBox txtEndereco 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   6495
      End
      Begin VB.TextBox txtBairro 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6840
         TabIndex        =   46
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtCidade 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10800
         TabIndex        =   47
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox txtUF 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   15120
         TabIndex        =   48
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Logradouro"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   75
         Top             =   120
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6840
         TabIndex        =   74
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10800
         TabIndex        =   73
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   15120
         TabIndex        =   72
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.Frame frm01 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pesquisa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   15840
      TabIndex        =   63
      Top             =   720
      Width           =   4575
      Begin VB.ComboBox cmbPesqPedido 
         BackColor       =   &H00FFFFE0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   960
         Sorted          =   -1  'True
         TabIndex        =   38
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton cmbConsultaPedidos 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cmbStatusPedido 
         BackColor       =   &H00FFFFE0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   960
         TabIndex        =   39
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   65
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
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
         Height          =   300
         Left            =   0
         TabIndex        =   64
         Top             =   480
         Width           =   1005
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Processamento de Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   10320
      TabIndex        =   62
      Top             =   8160
      Width           =   4935
      Begin VB.TextBox txtNotaFiscal 
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
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   1680
         TabIndex        =   27
         Top             =   960
         Width           =   3015
      End
      Begin VB.ComboBox cmbCFOP 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1800
         Width           =   4575
      End
      Begin VB.ComboBox cmbEmissor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cmbBanco 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label46 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CFOP-Natureza da Operação"
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
         Left            =   120
         TabIndex        =   81
         Top             =   1560
         Width           =   4575
      End
      Begin VB.Label Label38 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco"
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
         TabIndex        =   69
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota Fiscal"
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
         Left            =   1680
         TabIndex        =   68
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblPlaca 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   50
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Resumo Financeiro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6240
      TabIndex        =   59
      Top             =   8160
      Width           =   3615
      Begin VB.Label txtValorComDesconto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1680
         TabIndex        =   89
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
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
         Left            =   120
         TabIndex        =   70
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label txtAcumula_Desconto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1680
         TabIndex        =   67
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label txtAcumula_Produto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1680
         TabIndex        =   66
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Produto....."
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
         Left            =   120
         TabIndex        =   61
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desconto.."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   60
         Top             =   1080
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalhe da Medição"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   0
      TabIndex        =   52
      Top             =   3360
      Width           =   15855
      Begin VB.OptionButton optImpNao 
         Caption         =   "Não"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   14
         Top             =   1230
         Width           =   680
      End
      Begin VB.OptionButton optImpSim 
         Caption         =   "Sim"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5550
         TabIndex        =   13
         Top             =   1230
         Width           =   735
      End
      Begin VB.TextBox txtPUCheio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8640
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtDesconto 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9840
         TabIndex        =   18
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox cmbAtividade 
         BackColor       =   &H00FFFFE0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   4395
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   750
         Width           =   2670
      End
      Begin VB.CommandButton cmdCalculaDias 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Calcula N/Dias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   12600
         MaskColor       =   &H00C0FFFF&
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtfim 
         Height          =   375
         Left            =   10560
         TabIndex        =   20
         Top             =   430
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   344653825
         CurrentDate     =   44268
      End
      Begin MSComCtl2.DTPicker dtInicio 
         Height          =   375
         Left            =   7680
         TabIndex        =   19
         Top             =   430
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   344653825
         CurrentDate     =   44268
      End
      Begin VB.ComboBox txtUnidade 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7080
         TabIndex        =   15
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtQtdDias 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   13200
         TabIndex        =   36
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtValorDiaria 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   11760
         TabIndex        =   35
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   0
         TabIndex        =   76
         Top             =   255
         Width           =   1215
         Begin VB.TextBox txtQtdFat 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   600
            TabIndex        =   8
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox txtIntervalo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   600
            TabIndex        =   9
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox TxtAPartirDe 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   600
            TabIndex        =   10
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fat."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   79
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Int."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   78
            Top             =   600
            Width           =   420
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Iníc"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   77
            Top             =   1080
            Width           =   435
         End
      End
      Begin VB.CommandButton cmdNovoProduto 
         BackColor       =   &H0000FFFF&
         Caption         =   "Novo Prod."
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   14760
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdSairProduto 
         BackColor       =   &H000080FF&
         Caption         =   "Nova Medição"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   14760
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CommandButton cmdExcluiDetalhe 
         BackColor       =   &H008080FF&
         Caption         =   "Exc"
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
         Left            =   14040
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1260
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdAlteraDetalhe 
         BackColor       =   &H00FFFF80&
         Caption         =   "Alt"
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
         Left            =   14040
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   750
         Width           =   735
      End
      Begin VB.CommandButton cmdIncluiDetalhe 
         BackColor       =   &H0000FF00&
         Caption         =   "Inc"
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
         Left            =   14040
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtQtd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7800
         TabIndex        =   16
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtPreçoUnit 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10560
         TabIndex        =   34
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cmbProduto 
         BackColor       =   &H00FFFFE0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1200
         TabIndex        =   11
         Top             =   750
         Width           =   3135
      End
      Begin VB.TextBox txtNomeProduto 
         BackColor       =   &H00FFFFE0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1200
         TabIndex        =   33
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Importar"
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
         Left            =   4400
         TabIndex        =   110
         Top             =   1230
         Width           =   1100
      End
      Begin VB.Label lblLabel14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PU Desc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10560
         TabIndex        =   108
         Top             =   960
         Width           =   990
      End
      Begin VB.Label lblLabel12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9840
         TabIndex        =   106
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblLabel9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9720
         TabIndex        =   101
         Top             =   480
         Width           =   420
      End
      Begin VB.Label lblLabel8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7200
         TabIndex        =   100
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblLocalizacao 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         TabIndex        =   99
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblLabel10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local"
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
         Index           =   4
         Left            =   3960
         TabIndex        =   98
         Top             =   165
         Width           =   480
      End
      Begin VB.Label lbllocal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   8475
         TabIndex        =   97
         Top             =   -480
         Width           =   135
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N/Dias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   13200
         TabIndex        =   93
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Diária"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11760
         TabIndex        =   57
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "PU Cheio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8640
         TabIndex        =   56
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Unid."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7080
         TabIndex        =   55
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7800
         TabIndex        =   54
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Produto"
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
         Left            =   2040
         TabIndex        =   53
         Top             =   540
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Relação dos Produtos Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3015
      Left            =   -120
      TabIndex        =   58
      Top             =   5160
      Width           =   15975
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   3015
         Left            =   -120
         TabIndex        =   91
         Top             =   0
         Width           =   16095
         _ExtentX        =   28390
         _ExtentY        =   5318
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         BackColor       =   16777152
         BackColorFixed  =   16776960
         BackColorBkg    =   16777152
         FocusRect       =   2
         FormatString    =   $"Pedido.frx":0000
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
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   8040
   End
   Begin VB.Line Line1 
      X1              =   11880
      X2              =   11880
      Y1              =   0
      Y2              =   7920
   End
End
Attribute VB_Name = "frmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DataInicioReal As Date
Dim DataFinalReal As Date
Dim precoUnit As Currency
Dim Produto As String
Dim Pedido As String
Dim PedidoComp As String
Dim Verifica As String
Dim ContaComponente As Integer
Dim GrupoAnterior As String
Dim Fim As Byte
Dim Linha As Integer
Dim coluna As Integer
Dim SalvaLocal As String
Dim DataInvertida As String
Dim dataInicio As String
Dim dataFim As String
Dim DataInicioAnter As Date
Dim DataFimAnter As Date
Dim DataInicioEventoStr As String
Dim dtInicioEventoMenosUm As Date
Dim DataInicioEvento As Date
Dim DataFinalEvento As Date
Dim VerificaData As Byte
Dim ValorHora As Currency
Dim ValorMinuto As Currency
Dim QtdHoras As Integer
Dim QtdMinutos As Integer
Dim MinutosParaCalculo As Integer
Dim TipoProduto As Byte
Dim OrdemApresentacao As Byte

Dim Altera As Byte
Dim Resp As String
Dim Fechado As Byte
Dim Encontrei As Byte
Dim Inclui_Pedido As Integer
Dim Inclui_Detalhe As Integer
Dim inclui_Entrega As Integer
Dim IndCobranca As Integer
Dim PrecoProdutoComposto As Currency
Dim Valor_Operacao As Currency
Dim AjusteComissao As Currency
Dim Acumula_Operacao As Currency
Dim DataProc As Date
Dim DataPedido As Date
Dim lstPessoa As Integer
Dim DeletaCheque As Byte
Dim ValorFrete As Currency
Dim UltimaLinha As Integer
Dim RecalculaComissao As Byte
Dim RedutorDeComissao As Integer
Dim IndiceComissao As Byte
Dim Ind As Byte
Dim AcumDel As Integer
Dim Pessoa As String
Dim NaoAchei As Byte
Dim ComisPromot As Integer
Dim ComisRep As Integer
Dim AjustaComis As Integer

Dim SQL2 As String

Dim ProdutoConsig(3) As String

Dim Erro_Critica As Byte

Dim Acumula_Comis_Rep As Currency
Dim Acumula_Comis_Promot As Currency

Dim QtdAnterior As Currency  'Para Alteracao ou retirada da qtd de produto em um pedido

'Dim Valor_IPI_Qtd As Currency
'Dim Acumula_IPI As Currency

Dim Valor_Com_Desconto As Currency
Dim Valor_Com_Desconto_Qtd As Currency
Dim Acumula_Desconto As Currency

Dim Valor_Frete As Currency
Dim Acumula_Frete As Currency

Dim Acumula_Metro As Currency
Dim Acumula_Produto As Currency

Dim Acumula_Caixa As Currency
Dim Acumula_Peso As Currency

'Dim fim As Byte

Dim Data_Pedido As Date
Dim Dia_Pedido As Integer
Dim Mes_Pedido As Integer
Dim Ano_Pedido As Integer

Dim Data_Comis As Date
Dim Dia_Comis As Integer
Dim Mes_Comis As Integer
Dim Ano_Comis As Integer

Dim ChavePedido As String
Dim ChaveCompPedido As String

Dim DataConv As Date

Dim DataGridInvertida As String
Dim OC As Byte 'Flag indicativa de que haverá cadastramento de ordem de carga
Dim Contrato As String

Private Sub cmbAtividade_LostFocus()
If cmbAtividade = "HORA EXTRA" Then
   txtUnidade.ListIndex = 2
Else
   If cmbAtividade = "HORA EXTRA NOTU" Then
      txtUnidade.ListIndex = 3
   End If
End If
   
End Sub



Private Sub cmbLocal_LostFocus()

Call Rotina_AbrirBanco

Call CargaProduto

Call FechaDB

End Sub

Private Sub cmdCalculaDias_Click()

If VerificaData = 1 Then
   If Not (dtInicio = DataInicioAnter) Then
      MsgBox ("Data inicio e final de período não pode ser alterada. Primeiro exclua o lançamento com período incorreto."), vbInformation
      Call FechaDB
      VerificaData = 0
      cmdIncluiDetalhe.Enabled = False
      cmdAlteraDetalhe.Enabled = False
      cmdExcluiDetalhe.Enabled = True
      dtInicio = DataInicioAnter
      Exit Sub
   Else
      If Not (dtFim = DataFimAnter) Then
         MsgBox ("Data inicio e final de período não podem ser alteradas. Primeiro exclua o lançamento com período incorreto."), vbInformation
         Call FechaDB
         VerificaData = 0
         cmdIncluiDetalhe.Enabled = False
         cmdAlteraDetalhe.Enabled = False
         cmdExcluiDetalhe.Enabled = True
         dtFim = DataFimAnter
         Exit Sub
      End If
   End If
End If



If Not ((dtFim + 1) > dtInicio) Then
   MsgBox ("Data fim tem que ser posterior a data início."), vbCritical
   Exit Sub
End If

If Not (txtUnidade.ListIndex = 2 Or txtUnidade.ListIndex = 3) Then
   txtQtdDias = (dtFim - dtInicio) + 1
Else
   txtQtdDias = 1
End If

Call Rotina_AbrirBanco

Prod.Open "Select * from produto where chproduto = ('" & cmbContrato & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Informar QQ Produto da lista"), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If

ano = Year(dtInicio)
mes = Month(dtInicio)
Dia = Day(dtInicio)

DataInvertida = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")

dataInicio = DataInvertida

ano = Year(dtFim)
mes = Month(dtFim)
Dia = Day(dtFim)

DataInvertida = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")

dataFim = DataInvertida

dneg.Open "Select * from detalhenegociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtComplementoPedido & "') and chProduto = ('" & cmbProduto & "') and chDatainicio = ('" & dataInicio & "') and chDataFim = ('" & dataFim & "')", db, 3, 3
If dneg.EOF Then
      'txtNomeProduto = Prod!prdNomeProd
      Inclui_Detalhe = 1
      cmdIncluiDetalhe.Enabled = True
      cmdAlteraDetalhe.Enabled = False
      cmdExcluiDetalhe.Enabled = False
      If Inclui_Pedido = 1 Then
         cmdAlteraPedido.Enabled = False
         cmdExcluiPedido.Enabled = False
      Else
         cmdAlteraPedido.Enabled = True
         cmdExcluiPedido.Enabled = True
      End If
      'txtUnidade.SetFocus
Else
      Inclui_Detalhe = 0
      cmdIncluiDetalhe.Enabled = False
      cmdAlteraDetalhe.Enabled = False
      cmdExcluiDetalhe.Enabled = True
      txtNomeProduto = Prod!prdNomeProd
      txtQtd = Format$(dneg!pedquantidadePedida, "0.00")
      txtPreçoUnit = Format$(dneg!pedPrecoUnidadePedida, "#0.00")
      'txtFrete = Format$(TabDetalheNegociacao("pedFreteUnidadePedida"), "#0.00")
      'txtDesc = Format$(TabDetalheNegociacao("pedDesc"), "#0.00")
      'If dneg!pedunidade = 1 Then
      '   txtUnidade = "M"
      'Else
      '   txtUnidade = "Un"
      '   txtPreçoUnit.SetFocus
      'End If
      txtUnidade.ListIndex = dneg!pedunidade
      txtDesconto = Format$(dneg!pedDesconto, "##0.00")
      
      'txtUnidade.SetFocus
      
      cmdIncluiDetalhe.Enabled = False
      cmdAlteraDetalhe.Enabled = True
      cmdExcluiDetalhe.Enabled = True
      cmdNovoProduto.Enabled = True
      cmdSairProduto.Enabled = True
End If

Call FechaDB

End Sub



Private Sub dtFimMedicao_LostFocus()

If dtInicioMedicao = dtFimMedicao Then
   Resp = MsgBox("Data de Inicio da Mediçao é igual a Data Final. Caso esteja correto informe que SIM???", vbYesNo)

   If Resp = vbNo Then
      dtInicioMedicao.SetFocus
   End If
End If

End Sub

'Private Sub cmdDiasPeriodo_Click()
'txtDiasPeriodo = dtFim - dtInicio - 1
'End Sub

Private Sub Form_Initialize()
Fechado = 0
Contrato = "CONTRATO"
End Sub
'Private Sub cmbCobrancaFrete_LostFocus()

'IndCobranca = cmbCobrancaFrete.ListIndex

'If IndCobranca = 0 Then
'   txtPrzBoletaFrete.Enabled = False
'   txtPreçoFixo.Enabled = False
'   txtPreçoFixo.SetFocus
'Else
'   If IndCobranca = 2 Then
'      txtPrzBoletaFrete.Enabled = False
'      txtPreçoFixo = Format$(0, "#0.00")
'      txtPreçoFixo.Enabled = False
'      cmbCondProcessamento.SetFocus
'      lblTransporte = "Cliente"
'      lblPlaca = "Cliente"
'      cmbOrdemDeCarga = "Cliente"
 '   Else
'      If IndCobranca = 4 Or IndCobranca = 5 Then
'         txtPreçoFixo.Enabled = True
'         lblTransporte = "Cliente"
'         lblPlaca = "Cliente"
 '        cmbOrdemDeCarga = "Cliente"
 '     Else
 '        If IndCobranca = 7 Then
 '           txtPreçoFixo.Enabled = True
 '        Else
 '           txtPreçoFixo.Enabled = False
 '        End If
 '     End If
 '   End If
'End If
'If IndCobranca = 1 Then
'   txtPrzBoletaFrete.Enabled = True
'End If
'End Sub

Private Sub cmbConsultaPedidos_Click()

Dim linha_pedido As Integer
GridPedido.ColAlignment(0) = 2
GridPedido.ColAlignment(1) = 1
GridPedido.ColAlignment(2) = 3
   
If cmbPesqPedido = Empty Then
   MsgBox ("Cliente para Pesquisa não Informado"), vbInformation
   cmbPesqPedido.SetFocus
   Exit Sub
End If

If cmbStatusPedido = Empty Then
   cmbStatusPedido.ListIndex = 0
End If

GridPedido.Rows = 2
linha_pedido = 1
    GridPedido.TextMatrix(linha_pedido, 0) = Empty
    GridPedido.TextMatrix(linha_pedido, 1) = Empty
    GridPedido.TextMatrix(linha_pedido, 2) = Empty
    GridPedido.TextMatrix(linha_pedido, 3) = Empty
    GridPedido.TextMatrix(linha_pedido, 4) = Empty


linha_pedido = 0

Call Rotina_AbrirBanco

neg.Open "select * from negociacao", db, 3, 3
If neg.EOF Then
   MsgBox ("Não há negociação até o presente momento."), vbInformation
   Call FechaDB
   Exit Sub
End If
    
neg.MoveFirst

Do While Not neg.EOF
   If (cmbPesqPedido = " Geral" And neg!negStatus = cmbStatusPedido.ListIndex) Or (cmbPesqPedido = neg!chPessoa And neg!negStatus = cmbStatusPedido.ListIndex) Then
      linha_pedido = linha_pedido + 1
      GridPedido.Rows = linha_pedido + 1
      GridPedido.TextMatrix(linha_pedido, 0) = neg!chNumPedido
      GridPedido.TextMatrix(linha_pedido, 1) = neg!chNumPedidoComp
      GridPedido.TextMatrix(linha_pedido, 2) = Format$(neg!negDataPedido, "dd/mm/yy")
      GridPedido.TextMatrix(linha_pedido, 3) = neg!chPessoa
      
      ano = Year(neg!negDataPedido)
      mes = Month(neg!negDataPedido)
      Dia = Day(neg!negDataPedido)

      DataGridInvertida = ano & Format$(mes, "00") & Format$(Dia, "00")

      GridPedido.TextMatrix(linha_pedido, 4) = (DataGridInvertida & GridPedido.TextMatrix(linha_pedido, 1) & GridPedido.TextMatrix(linha_pedido, 0))
      neg.MoveNext
   Else
      neg.MoveNext
   End If
Loop
'GridPedido.Col = 4
'GridPedido.ColSel = 4
     
GridPedido.Row = 1
GridPedido.RowSel = linha_pedido
        
If linha_pedido > 1 Then
   GridPedido.Sort = 7
End If
     
GridPedido.Col = 0
GridPedido.ColSel = 0
GridPedido.Row = 0
GridPedido.RowSel = 0

cmbPesqPedido.SetFocus

Call FechaDB

End Sub

'Private Sub cmbEmissor_LostFocus()

'TabPagtosEmCheque.MoveFirst

'Do While Not TabPagtosEmCheque.EOF
'   If TabPagtosEmCheque("ocgdatadacarga") = Data_Hoje Then
'      If TabPagtosEmCheque("chemissor") = cmbEmissor Then
'         cmbOrdemDeCarga.AddItem TabPagtosEmCheque("chordemdecarga")
'      End If
'   End If
'   TabPagtosEmCheque.MoveNext
'Loop
'End Sub

'Private Sub cmbOrdemDeCarga_LostFocus()
'OC = 0
'If cmbOrdemDeCarga = "Cliente" Then
'   lblPlaca = "Cliente"
'   lblTransporte = "Cliente"
'   Exit Sub
'End If

'If cmbOrdemDeCarga = Empty Or cmbOrdemDeCarga = "Cliente" Then
'   Exit Sub
'End If

'Verifica = Empty
'Verifica = Mid$(cmbOrdemDeCarga, 11, 5)
'If Not Verifica = Empty Then
'   MsgBox ("Número da Ordem de Carga não pode ter mais que 10 caracteres")
'   cmdSair.SetFocus
'   Exit Sub
'End If'

'TabPagtosEmCheque.Seek "=", cmbOrdemDeCarga, cmbEmissor

'If TabPagtosEmCheque.NoMatch Then
'   Resp = MsgBox("Ordem de Carga não cadastrada. Deseja cadastra-la agora???", vbYesNo)

'   If Resp = vbYes Then
'      EmissorOrdemDeCarga = cmbEmissor
'      OrdemDeCarga = cmbOrdemDeCarga
'      OC = 1
'      frmOrdemDeCarga.Show vbModal
'      cmbOrdemDeCarga.SetFocus
'   Else
'      cmbOrdemDeCarga.SetFocus
'   End If
'Else
'   lblTransporte = TabPagtosEmCheque("ocgmotorista")
'   lblPlaca = TabPagtosEmCheque("ocgplaca")
'End If
'End Sub

Private Sub cmbPessoa_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub cmbPessoa_LostFocus()
Dim Promotora As String
Dim Representante As String

On Error Resume Next


Call Rotina_AbrirBanco

Contrato = "CONTRATO"

If cmbPessoa = Empty Then
   MsgBox "Informar um codigo da lista para pessoa"
   txtNumPedido.SetFocus
   Exit Sub
End If

cmbLocal.Clear

If Inclui_Pedido = 1 Then
   cmbContrato.Clear
   
   Prod.Open "Select * from produto where prdLocadora = ('" & cmbPessoa & "') AND prdUnidadeOperacional = ('" & Contrato & "')", db, 3, 3
   If Prod.EOF Then
      MsgBox ("Cliente sem contrato cadastrado"), vbInformation
      Call FechaDB
      Exit Sub
   End If
   
   Prod.MoveFirst
   
   cmbContrato.Clear
   
   Do While Not Prod.EOF
      cmbContrato.AddItem Prod!chProduto
      Prod.MoveNext
   Loop
End If

uoper.Open "Select * from unidadeoperacional where chpessoa = ('" & cmbPessoa & "')", db, 3, 3
If uoper.EOF Then
   MsgBox ("Cadastrar a Unidade Operacional deste cliente. Somente após o cadastramentoserá possível esta operação"), vbCritical
   Call FechaDB
   Exit Sub
End If

cmbLocal.Clear

Do While Not uoper.EOF
   cmbLocal.AddItem uoper!chUnidadeOperacional
   uoper.MoveNext
Loop

cmbLocal = SalvaLocal
cmbContrato = neg!negContrato

pes.Open "Select * from pessoa where chpessoa = ('" & cmbPessoa & "')", db, 3, 3

If pes.EOF Then
   Resp = MsgBox("Cliente não cadastrado. Deseja cadastrar agora???", vbYesNo)

   If Resp = vbYes Then
      cmbPessoa.SetFocus
      Call FechaDB
      frmPessoa.Show vbModal
      Call Rotina_AbrirBanco
      pes.Open "Select * from pessoa where chPessoa = ('" & cmbPessoa & "')", db, 3, 3
      If pes.EOF Then
         Call FechaDB
         MsgBox ("Execução de cadastramento de pessoa inválido"), vbCritical
         Exit Sub
      Else
         Promotora = pes!chCarteiraPromot
         Representante = pes!chCarteiraRepresentante
      End If
   Else
      cmbPessoa.SetFocus
   End If
Else
   Promotora = pes!chCarteiraPromot
   Representante = pes!chcarteirarep
   If Not pes!pesStatusPessoa = 0 Then
         Resp = MsgBox("Cliente Inativo. Retornar a condição de ATIVO", vbYesNo)
         If Resp = vbYes Then
            pes!pesStatusPessoa = 0
            pes.Update
            
            pes.Close: Set pes = Nothing
            Call Rotina_AbrirBanco
            pes.Open "Select * from pessoa where chpessoa = ('" & cmbPessoa & "')", db, 3, 3
         End If
   End If
End If

hneg.Open "Select * from historiconegociacao where chPessoa = ('" & cmbPessoa & "') and chnumpedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtComplementoPedido & "')", db, 3, 3
If hneg.EOF Then
   
 '  CartPromot.Open "Select * from carteira_promot where chCarteiraPromot = ('" & Promotora & "')", db, 3, 3'

 '  If CartPromot.EOF Then
 '     MsgBox ("Promotor não informado. Verificar cadastro deste cliente"), vbInformation
 '     frmPessoa.Show vbModal
 '     Call FechaDB
 '''     Exit Sub
 '  End If
 '  CartRep.Open "Select * from carteira_rep where chCarteiraRep = ('" & Representante & "')", db, 3, 3
     
 '  If CartRep.EOF Then
 '     MsgBox ("Representante não cadastrado. Cadastrar Representante e retornar para confecção do pedido"), vbInformation
 '     Unload Me
 '     Call FechaDB
 '     Exit Sub
 '  End If
   If pes!pesPessoa = 0 Then
      txtCNPJCPF.Mask = "###.###.###-##"
      txtCNPJCPF = pes!chCNPJ_CPF
   Else
      txtCNPJCPF.Mask = "##.###.###/####-##"
      txtCNPJCPF = pes!chCNPJ_CPF
   End If
   txtRepresentante = CartRep!chPessoa
   ' Verifica = Mid$(txtRepresentante, 1, 7)
   ' If Verifica = "Fabrica" Or Verifica = "FABRICA" Then
   '    cmbCondProcessamento.ListIndex = 2
   ' End If
   txtPromotora = CartPromot!chPessoa
   txtEndereco = pes!pesEndereco
   txtBairro = pes!pesBairro
   txtCidade = pes!pesCidade
   txtUF = pes!huf
   ' txtCEP = pes!pesCEP
   'Contato.Open "Select * from telefone where codPessoa = ('" & cmbPessoa & "')", db, 3, 3
   'If Contato.EOF Then
   '   txtTel = "N/INFORMADO"
   'Else
   '   txtTel = Contato!codigocontato
   'End If
   txtRepresentante = Representante
   txtPromotora = Promotora
Else
   MsgBox ("Número de pedido e complemento ja utilizado para este Cliente."), vbInformation
   cmbPessoa = Empty
   txtNumPedido.SetFocus
End If

'If Prod.State = 1 Then
'   Prod.Close: Set Prod = Nothing
'End If

'cmbLocal.SetFocus

Call FechaDB

End Sub
Private Sub cmbCondProcessamento_LostFocus()
If txtQtdFat = Empty Then
   txtQtdFat = 0
End If
If txtIntervalo = Empty Then
   txtIntervalo = 0
End If
'If txtPrzBoletaFrete = Empty Then
'   txtPrzBoletaFrete = 0
'End If

'If cmbCondProcessamento.ListIndex = 1 Then
'   txtDescComissao.Enabled = True
'Else
'   txtDescComissao = 0
'   txtDescComissao.Enabled = False
'End If

End Sub

Private Sub cmbProduto_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub cmbProduto_LostFocus()

On Error Resume Next

Call Rotina_AbrirBanco

Prod.Open "Select * from produto where chproduto = ('" & cmbProduto & "')", db, 3, 3
If Prod.EOF Then
   'MsgBox ("Informar QQ Produto da lista"), vbInformation
   'cmdSair.SetFocus
   Exit Sub
Else
   txtUnidade.ListIndex = Prod!prdunidade
   txtNomeProduto = Prod!prdNomeProd
   TipoProduto = Prod!prdOrdemApresentacao
   OrdemApresentacao = Prod!prdOrdemApresentacao
End If

optImpSim = False
optImpNao = True

Call FechaDB

End Sub

Private Sub cmdAlteraDetalhe_Click()

VerificaData = 0

QtdAnterior = 0

Call Rotina_AbrirBanco

ano = Year(dtInicio)
mes = Month(dtInicio)
Dia = Day(dtInicio)

DataInvertida = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")
 
dataInicio = DataInvertida

ano = Year(dtFim)
mes = Month(dtFim)
Dia = Day(dtFim)

DataInvertida = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")

dataFim = DataInvertida

dneg.Open "Select * from detalhenegociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtComplementoPedido & "') and chProduto = ('" & cmbProduto & "') and chDatainicio = ('" & dataInicio & "') and chDataFim = ('" & dataFim & "')", db, 3, 3
If dneg.EOF Then
   MsgBox ("Erro no acesso a Produto em cmdAlteraDetalhe."), vbCritical
   Call FechaDB
   Exit Sub
End If


Call Rotina_Mover_Detalhe

db.BeginTrans
   

   QtdAnterior = dneg!pedquantidadePedida
   dneg.Update


db.CommitTrans

'Atualizar pedidos em carteira de pedidos
   
Funcao = 2
ano = Year(Date)
mes = Month(Date)
Produto = cmbProduto
  
Sai = QtdAnterior
Entra = dneg!pedquantidadePedida
 
TracoIn = 0
TracoOut = 0
Mes_Pedido = Month(txtDataPedido) 'Month(TabNegociacao("negDataPedido"))
   
'Call Rotina_Atualiza_Estoque(Funcao, Ano, Mes, Produto, Entra, Sai, TracoIn, TracoOut, Mes_Pedido)
 
Call Rotina_Limpa_Detalhe
'TabDetalheNegociacao.MoveFirst
Call Rotina_Carga_Grid
'cmdNovoProduto.SetFocus
Call ResumoMedicao

Call FechaDB

End Sub
Private Sub cmdAlteraPedido_Click()

Altera = 1

Call Rotina_AbrirBanco


neg.Open "Select * from negociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtComplementoPedido & "')", db, 3, 3
If neg.EOF Then
   MsgBox ("Erro no acesso a Negociacao em Altera Negociacao"), vbCritical
   Call FechaDB
   End
Else
   TipoProduto = neg!negTipoProduto
End If

pes.Open "SElect * from pessoa where chPessoa = ('" & neg!chPessoa & "')", db, 3, 3
If neg.EOF Then
   MsgBox "Caquinha na leitura de pessoa - cmdAlteraPedido"
   End
End If
If neg!negStatus = 1 Then
   MsgBox ("Pedido Processado. Alteração não permitida"), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If
Inclui_Pedido = 0

' If Tabpessoa("pesEndereco") <> txtEndereco Then
'   TabEntrega.Seek "=", TabNegociacao("chNumPedido"), TabNegociacao("chNumPedidoComp")''

'      If TabEntrega.NoMatch Then
'         TabEntrega.AddNew
'         Call Rotina_Carrega_Entrega
'         TabEntrega.Update
'      Else
'         TabEntrega.Edit
'         Call Rotina_Carrega_Entrega
'         TabEntrega.Update
'      End If
'End If

Erro_Critica = 0

Call Rotina_Criticar_Campos

If Erro_Critica = 0 Then

    db.BeginTrans
    
    Call Rotina_Mover_Negociacao
    neg.Update
   db.CommitTrans
'Aqui Pedido
    
    Call Rotina_Carrega_Pedido
    Call Rotina_Limpa_Detalhe
    
    dneg.Open "Select * from detalhenegociacao where chNumPedido = ('" & txtNumPedido & "') and chnumpedidocomp = ('" & txtComplementoPedido & "')", db, 3, 3
    If dneg.EOF Then
       MsgBox ("Detalhe de Negociação não encontrado."), vbInformation
    Else
       dneg.MoveFirst
    End If
    
    Call Rotina_Carga_Grid
    cmbPessoa.Enabled = False
End If

Call FechaDB
    
End Sub
Private Sub cmdDesmembrar_Click()
MsgBox ("Função não disponível "), vbInformation
'Fechado = 0

'If cmbPessoa = Empty Then
'   MsgBox ("Solicitação de desmembramento sem informação.")
'   cmbPessoa.SetFocus
'   Exit Sub
'End If

'Resp = MsgBox("Desmembramento solicitado. Confirma???", vbYesNo)

'TabNegociacao.Seek "=", txtNumPedido, txtComplementoPedido
'If TabNegociacao.NoMatch Then
'   MsgBox "Erro no acesso ao pedido para desmembramento"
'   Exit Sub
'End If

'If Resp = vbYes Then
'   If Fechado = 1 Then
'      deNegociacao.rscmdNegociacao.Close
'      deHistNeg.rscmdHistNeg.Close
'   Else
'      Fechado = 1
''   End If
'   If Not (TabNegociacao("chnumpedido") = txtNumPedido And TabNegociacao("chnumpedidocomp") = txtComplementoPedido) Then
'      TabNegociacao.Seek "=", txtNumPedido, txtComplementoPedido
'      If TabNegociacao.NoMatch Then
'         MsgBox "Erro no acesso a pedido para demembramento"
'         Exit Sub
'      End If
''   End If
''   glbNumPedido = TabNegociacao("chNumPedido")
'   glbCompPedido = TabNegociacao("chNumPedidoComp")
   
'   pessoa = TabNegociacao("chpessoa")
'   Pedido = TabNegociacao("chnumpedido")

'   SQL2 = "Select chnumpedido as Numero_do_Pedido, chnumpedidocomp as Compl from negociacao"
'   SQL2 = SQL2 & " where chPessoa like '" & pessoa & "'"
'   SQL2 = SQL2 & " and chnumpedido like '" & Pedido & "'"
'   SQL2 = SQL2 & " order by chnumpedido, chnumpedidocomp"
'   'MsgBox SQL2
   
'   deNegociacao.Commands.Item("cmdnegociacao").CommandText = SQL2

'   SQL2 = "Select hng.chnumpedido as Numero_do_Pedido, hng.chnumpedidocomp as Compl from historiconegociacao hng"
'   SQL2 = SQL2 & " where hng.chPessoa like '" & pessoa & "'"
'   SQL2 = SQL2 & " and hng.chnumpedido like '" & Pedido & "'"
'   SQL2 = SQL2 & " order by hng.chnumpedido, hng.chnumpedidocomp"
'   'MsgBox SQL2
   
'   deHistNeg.Commands.Item("cmdhistneg").CommandText = SQL2

'   frmDesmembra_Pedido.Show
'Else
'   txtNumPedido.Enabled = True
'   txtComplementoPedido.Enabled = True
'   txtNumPedido.SetFocus
'End If

End Sub
Private Sub cmdExcluiDetalhe_Click()

Resp = MsgBox("Exclusão de Produto. Confirma???", vbYesNo)
If Resp = vbYes Then
   'QtdAnterior = TabDetalheNegociacao("pedQuantidadeMetro")
   
   Call Rotina_AbrirBanco
   
   ano = Year(dtInicio)
   mes = Month(dtInicio)
   Dia = Day(dtInicio)
   
   DataInvertida = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")
    
   dataInicio = DataInvertida
  
   ano = Year(dtFim)
   mes = Month(dtFim)
   Dia = Day(dtFim)
   
   DataInvertida = ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00")
   
   dataFim = DataInvertida
   
   dneg.Open "Select * from detalhenegociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtComplementoPedido & "') and chProduto = ('" & cmbProduto & "') and chDatainicio = ('" & dataInicio & "') and chDataFim = ('" & dataFim & "')", db, 3, 3
   If dneg.EOF Then
      MsgBox ("Erro no acesso a Detalhe de Negociacao em cmdExcluiDetalhe"), vbCritical
      Exit Sub
   End If
   
   db.BeginTrans
        
   Call Rotina_Mover_Detalhe
   dneg.Delete
   
  db.CommitTrans
   
   dneg.Close: Set dneg = Nothing
   
   'Atualizar peidos em carteira de pedidos
   
   Funcao = 2
   ano = Year(Date)
   mes = Month(Date)
   Produto = cmbProduto
  
   Sai = QtdAnterior
   Entra = 0
 
   TracoIn = 0
   TracoOut = 0
   Mes_Pedido = Month(txtDataPedido) 'Month(TabNegociacao("negDataPedido"))
   
'   Call Rotina_Atualiza_Estoque(Funcao, Ano, Mes, Produto, Entra, Sai, TracoIn, TracoOut, Mes_Pedido)
 
   Call Rotina_Limpa_Detalhe
   Call Rotina_Limpa_Grid
   
   dneg.Open "Select * from detalhenegociacao", db, 3, 3
   If dneg.EOF Then
      MsgBox ("Erro no acesso a Detalhe de Negociacao em cmdExcluiDetalhe Movefirst"), vbCritical
      Exit Sub
   End If
   
   
   dneg.MoveFirst
   Call Rotina_Carga_Grid
   cmbProduto.SetFocus
Else
   cmbProduto.SetFocus
End If

Call ResumoMedicao

Call FechaDB

End Sub
Private Sub cmdExcluiPedido_Click()

Call Rotina_AbrirBanco

neg.Open "Select * from negociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtComplementoPedido & "')", db, 3, 3
If neg.EOF Then
   MsgBox ("Registro para exclusão não encontrado em cmdExcluiPedido."), vbCritical
   Call FechaDB
   Exit Sub
End If

If neg!negStatus = 0 Then
   Resp = MsgBox("Exclusão de PEDIDO. Confirma???", vbYesNo)
Else
   Resp = MsgBox("Cancelamento de Operação Processada. Confirma???", vbYesNo)
End If
   
   If Resp = vbYes Then
      If neg!negStatus = 1 Then
         Call Rotina_Exclui_Comissao_Rep
         Call Rotina_Exclui_Comissao_Promot
         Call Rotina_Exclui_Cta_Receber
         Call Rotina_Exclui_Cta_Pagar
               
         If Not IsNull(neg!negNumFatura) Or neg!negNumFatura > 0 Then
            MsgBox ("Atenção: Fatura ") & neg!negNumFatura & (" já havia sido impressa. Deve ser cancelada junto a Contabilidade. Este Número de Série NÃO mais será utilizado."), vbCritical
         End If
         neg!negStatus = 0
        'neg!negDataVencimento = Empty
         neg!negNotaFiscal = Empty
         neg!negCEFOP = 0
         neg!negICMS = Format$(0, "0.00")
         neg!negAliquota = Format$(0, "0.00")
         neg!negFretePedido = Format$(0, "0.00")
         neg!negValorDoProduto = Format$(0, "0.00")
         'neg!negIPI = Format$(0, "0.00")
         neg!chOrdemDeCarga = " "
         neg!negTransporte = " "
         neg!negPlaca = " "
         neg!negDescontoTotalPedido = Format$(0, "0.00")
         neg!negComisRepPedido = Format$(0, "0.00")
         neg!negComisPromotPedido = Format$(0, "0.00")
         neg!negCntrlFaturamento = 0

         neg.Update
         txtNumPedido.Enabled = True
         txtComplementoPedido.Enabled = True
         txtNumPedido.SetFocus
      Else
         'Call Rotina_Exclui_Entrega
         Call Rotina_Exclui_Detalhe
      
         neg.Delete
         
         MsgBox ("Exclusão de Registro de Negociação efetuada com sucesso"), vbInformation
      
         Call Limpa_Pedido
         Call Rotina_Carrega_Pedidos
         Call cmbConsultaPedidos_Click
         txtNumPedido.Enabled = True
         txtComplementoPedido.Enabled = True
         txtNumPedido.SetFocus
      End If
   Else
      cmdAlteraPedido.Enabled = True
      cmdExcluiPedido.Enabled = True
   End If
'Else
'   TabNegociacao.Seek "=", txtNumPedido, txtComplementoPedido
'   If TabNegociacao.NoMatch Then
'      MsgBox "Erro: Leitura de Negociacao inexistente"
'      End
'   Else
'      Resp = MsgBox("ExclusãO de Produto. Confirma???", vbYesNo)
'      If Resp = vbYes Then
'         Call Rotina_Exclui_Entrega
'         Call Rotina_Exclui_Detalhe
'         db.begintrans
'         TabNegociacao.Delete
'        db.CommitTrans
'         Call Limpa_Pedido
'         txtNumPedido.SetFocus
'      Else
'         cmdAlteraPedido.Enabled = True
'         cmdExcluiPedido.Enabled = True
'      End If
'   End If

'End If

Call FechaDB

End Sub
Private Sub cmdIncluiDetalhe_Click()


Call Rotina_AbrirBanco

If txtNumPedido = "" And txtComplementoPedido = "" Then
   Resp = MsgBox("Numero do Pedido Gerado pelo Sistema. Confirma???", vbYesNo)
   If Resp = vbNo Then
      MsgBox ("Numero do Pedido não Informado"), vbInformation
      cmdSair.SetFocus
      Exit Sub
   End If
   ChavePedido = 0
   ChaveCompPedido = 1

   Emp.Open "Select * from empresa where chPessoa = ('" & "SHB BRASIL" & "')", db, 3, 3
   If Emp.EOF Then
      MsgBox ("ERRO: Tabempresa não encontrado"), vbCritical
      Call FechaDB
      Unload Me
      Exit Sub
   End If

   ChavePedido = Emp!empNumPedido
   Fim = 0
   Do While Fim = 0

      neg.Open "Select * from negociacao where chNumPedido = ('" & ChavePedido & "') and chNumPedidoComp = ('" & ChaveCompPedido & "')", db, 3, 3
      If neg.EOF Then
         Do While Fim = 0
            hdneg.Open "Select * from historicodetalhenegociacao where chPessoa = ('" & cmbPessoa & "') and chNumPedido = ('" & ChavePedido & "') and chNumPedidoComp - ('" & ChaveCompPedido & "')", db, 3, 3
            If hdneg.EOF Then
               Fim = 2
            Else
               ChavePedido = ChavePedido + 1
               hdneg.Close: Set hdneg = Nothing
               Fim = 0
            End If
         Loop
      Else
         ChavePedido = ChavePedido + 1
         neg.Close: Set neg = Nothing
      End If
   Loop

   neg.Close: Set neg = Nothing
   hdneg.Close: Set hdneg = Nothing

   Emp!empNumPedido = ChavePedido + 1
   Emp.Update
   txtNumPedido = ChavePedido
   txtComplementoPedido = ChaveCompPedido
End If

Call Rotina_Criticar_Campos

If Erro_Critica = 1 Then
   Erro_Critica = 0
   Exit Sub
End If

If Inclui_Pedido = 1 Then
   If Not (txtRepresentante = "NENHUM") Then
      CartRep.Open "Select * from carteira_rep where chCarteiraRep = ('" & txtRepresentante & "')", db, 3, 3
      If CartRep.EOF Then
         MsgBox ("Erro no acesso a Carteira de representantes"), vbCritical
         Call FechaDB
         Exit Sub
      Else
         Prod.Open "Select * from produto where chproduto = ('" & cmbProduto & "')", db, 3, 3
         If pes.EOF Then
            MsgBox ("Erro na carga de pessoa em cmdIncluiDetalhe_Click."), vbCritical
            Call FechaDB
            Exit Sub
         Else
            ComisRep = Prod!prdComissao
            AjustaComis = CartRep!repajustecomissao
            Prod.Close: Set Prod = Nothing
         End If
      End If
   End If
End If


If Not (txtPromotora = "NENHUM") Then
   CartPromot.Open "Select * from carteira_promot where chCarteiraPromot = ('" & txtPromotora & "')", db, 3, 3
   If CartPromot.EOF Then
      MsgBox ("Promotores não informado. Verificar cadstro deste cliente"), vbCritical
      Call FechaDB
      Exit Sub
   Else
      txtPromotora = CartPromot!chPessoa
      ComisPromot = CartPromot!prdcomissaopromot
   End If
End If

neg.Open "Select * from negociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtComplementoPedido & "')", db, 3, 3
If neg.EOF Then
   neg.AddNew
   db.BeginTrans
       Call Rotina_Mover_Negociacao
       neg!negDataLancamento = Date
       neg!negContrato = cmbContrato
       neg.Update
       cmdProcessar.Enabled = True
       cmdExcluiPedido.Enabled = True
       txtDataPedido.Enabled = False
       cmdSair.Enabled = True
   db.CommitTrans

    Inclui_Pedido = 0
End If
'Grava Detalhe da Negociação

dneg.Open "Select * from detalhenegociacao where chNumPedido = ('" & ChavePedido & "') and chNumPedidoComp = ('" & ChaveCompPedido & "') and chProduto = ('" & cmbProduto & "')", db, 3, 3
   If dneg.EOF Then
      dneg.AddNew
   End If

db.BeginTrans

   Inclui_Detalhe = 0

   Call Rotina_Mover_Detalhe
   
   'Atualizar pedidos em carteira de pedidos
   
   Funcao = 2
   ano = Year(Date)
   mes = Month(Date)
   Produto = cmbProduto
  
   Entra = dneg!pedquantidadePedida
   Sai = 0

   TracoIn = 0
   TracoOut = 0
   Mes_Pedido = Month(txtDataPedido) 'Month(neg!negDataPedido)
   
'   Call Rotina_Atualiza_Estoque(Funcao, Ano, Mes, Produto, Entra, Sai, TracoIn, TracoOut, Mes_Pedido)
 
   dneg.Update
   
   mes = Month(dtInicioMedicao)
   dataInicio = Year(dtInicioMedicao) & Format$(mes, "00")
   
   lgt.Open "Select * from logistica where chAnoMesRef = ('" & dataInicio & "') and chPessoa = ('" & cmbPessoa & "') and chUnidadeOperacional = ('" & cmbLocal & "') and chColaborador = ('" & cmbProduto & "') and chEvento = ('" & cmbAtividade & "') and lgtStatusImport = ('" & 0 & "')", db, 3, 3
   If lgt.EOF Then
      DataInicioReal = Empty
      DataFinalReal = Empty
   Else
      DataInicioReal = lgt!chInicioEvento
      DataFinalReal = lgt!lgtFimEventoReal
   End If
   
   If Not (DataInicioReal = Empty) Then
      lgt!lgtStatusImport = 1
      lgt.Update
   End If
   
   If (DataInicioReal = Empty) Or (Month(DataInicioReal) < Month(DataFinalReal)) Then
      If lgt.State = 1 Then
         lgt.Close: Set lgt = Nothing
      End If
      mes = Month(dtFimMedicao)
      dataInicio = Year(dtFimMedicao) & Format$(mes, "00")
      lgt.Open "Select * from logistica where chAnoMesRef = ('" & dataInicio & "') and chPessoa = ('" & cmbPessoa & "') and chUnidadeOperacional = ('" & cmbLocal & "') and chColaborador = ('" & cmbProduto & "') and chEvento = ('" & cmbAtividade & "') and lgtStatusImport = ('" & 0 & "')", db, 3, 3
      If lgt.EOF Then
         DataInicioReal = Empty
         DataFinalReal = Empty
      Else
         DataInicioReal = lgt!chInicioEvento
         DataFinalReal = lgt!lgtFimEventoReal
      End If
      
      If Not (DataInicioReal = Empty) Then
         lgt!lgtStatusImport = 1
         lgt.Update
      End If
   End If
db.CommitTrans

Call FechaDB

Call ResumoMedicao
'AQUI
 
'TabDetalheNegociacao.MoveFirst

Call Rotina_Carga_Grid

'Limpar Campos para entrada de novo produto para o pedido

Call Rotina_Limpa_Detalhe
cmdNovoProduto.Enabled = True
cmdNovoProduto.SetFocus
cmdIncluiDetalhe.Enabled = False
cmdAlteraDetalhe.Enabled = False
cmdExcluiDetalhe.Enabled = False
cmdSairProduto.Enabled = True
cmdNovoProduto.Enabled = True

End Sub
'Private Sub cmdNavega_Click(Index As Integer)
   
'   If neg.BOF Then
'      neg.MoveFirst
'   End If
   
'   If neg.EOF Then
'      neg.MoveLast
'   End If
 '
'   Select Case Index

'   Case 0
'        Call Rotina_Ler_Primeiro
'   Case 1
'        Call Rotina_Ler_Proximo
'   Case 2
'        Call Rotina_Ler_Anterior
'   Case 3
'        Call Rotina_Ler_Ultimo
        
'End Select

 '  If neg.BOF = True Then
'      neg.MoveFirst
'   End If
   
'   If neg.EOF = True Then
'      neg.MoveLast
'   End If

'Call Limpa_Pedido
'Aqui Pedido
'Call Rotina_Carrega_Pedido
'dneg.MoveFirst
'Call Rotina_Carga_Grid

'cmdAlteraPedido.Enabled = True
'cmdExcluiPedido.Enabled = True
'cmdIncluiDetalhe.Enabled = False
'cmdAlteraDetalhe.Enabled = False
'cmdExcluiDetalhe.Enabled = False
'cmdSairProduto.Enabled = True
'txtNumPedido.Enabled = False
'txtComplementoPedido.Enabled = False
''cmbProduto.SetFocus
'End Sub

Private Sub cmdNovoProduto_Click()
If cmbProduto = Empty Then
   cmbProduto = Empty
   cmbProduto.SetFocus
Else
   Resp = MsgBox("Você irá ignorar este lançamento. Confirma? Sim ou Não?", vbYesNo)
   If Resp = vbYes Then
      Call Rotina_Limpa_Detalhe
      cmbProduto.SetFocus
   Else
      cmbProduto.SetFocus
   End If
End If
End Sub
Private Sub cmdProcessar_Click()

'glbCFOP = cmbCFOP

If txtNumPedido = Empty Then
   MsgBox ("Processamento inválido. Pedido não informado")
   txtNumPedido.SetFocus
   Exit Sub
End If

Call Rotina_AbrirBanco

Bco.Open "Select * from banco where bcoempresa = ('" & 0 & "') and bcoCodBcoLart = ('" & cmbBanco.ListIndex & "')", db, 3, 3
If Bco.EOF Then
   MsgBox ("Numero do banco inválido. Posicione o cursor no banco desejado."), vbCritical
   Call FechaDB
   'cmbOrdemDeCarga.SetFocus
   Exit Sub
End If

ctr.Open "Select * from contas_a_receber where chNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3
If Not ctr.EOF Then
   Call FechaDB
   MsgBox ("Nota Fiscal existente em Negociação"), vbCritical
   Exit Sub
End If
   Resp = MsgBox("Processamento solicitado com Data Hoje. Confirma???", vbYesNo)
   If Resp = vbYes Then
      DataProc = Date
      If txtDataPedido > DataProc Then
            MsgBox ("Data de Processamento ANTERIOR a data do Pedido"), vbCritical
            Call FechaDB
            cmdSair.SetFocus
            Exit Sub
      End If
      If txtNotaFiscal = Empty Or txtNotaFiscal = " " Then
         MsgBox ("Informe o número da Nota Fiscal"), vbCritical
         Call FechaDB
         txtNotaFiscal.SetFocus
         Exit Sub
      End If
   Else
      Resp = MsgBox("Deseja Reprocessar ou Processar com data dif. de hoje????", vbYesNo)
      If Resp = vbYes Then
         DataProc = InputBox("Data Proc.", "Data de Processamento/Reprocessamento no mês no Formato DD/MM/AAAA")
         If txtDataPedido > DataProc Then
            MsgBox ("Data de Processamento MENOR que data do Pedido"), vbCritical
            Call FechaDB
            cmdSair.SetFocus
            Exit Sub
         End If
         If txtNotaFiscal = Empty Or txtNotaFiscal = " " Then
            MsgBox ("Informe o número da Nota Fiscal"), vbCritical
            Call FechaDB
            txtNotaFiscal.SetFocus
            Exit Sub
         Else
            If (Month(DataProc) <> Month(Date)) Or (Year(DataProc) <> Year(Date)) Then
               MsgBox ("Processamento/Reprocessamento permitido somente dentro do mes atual"), vbCritical
               Call FechaDB
               txtNotaFiscal.SetFocus
               Exit Sub
            Else
               If DataProc > Date Then
                  MsgBox ("Proc/Reproc com data posterior a hoje"), vbCritical
                  Call FechaDB
                  txtNotaFiscal.SetFocus
                  Exit Sub
               End If
            End If
         End If
      Else
         MsgBox ("Processamento abortado"), vbCritical
         Call FechaDB
         cmdSair.SetFocus
         Exit Sub
      End If
   End If
'Else
'   MsgBox ("Nota Fiscal ja processada"), vbCritical
'   Call FechaDB
'   txtNotaFiscal.Enabled = True
'   txtNotaFiscal.SetFocus
'   Exit Sub
'End If

txtDataProc = DataProc
'cmbEmissor.ListIndex = 1
If cmbEmissor = Empty Then
   MsgBox ("Não Informado o Emissor da Nota Fiscal"), vbCritical
   Call FechaDB
   cmbEmissor.SetFocus
   Exit Sub
End If

If cmbBanco = Empty Then
   MsgBox ("Não Informado o banco para faturamento"), vbCritical
   Call FechaDB
   cmdSair.SetFocus
   Exit Sub
End If

'If lblTransporte = Empty And lblPlaca = Empty Then
''   MsgBox ("Não informado o Transporte")
'   cmdSair.SetFocus
'   Exit Sub
'End If

neg.Open "Select * from negociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtComplementoPedido & "')", db, 3, 3
If neg.EOF Then
   MsgBox ("Registro não cadastrado corretamente. Verificar..."), vbCritical
   Call FechaDB
   cmdSair.SetFocus
   Exit Sub
End If

If neg!negStatus = 1 Then
   MsgBox ("Pedido já Processado"), vbCritical
   Call FechaDB
   txtNumPedido.SetFocus
   Exit Sub
Else
   If neg!negStatus = 2 Then
      neg!negStatus = 0
      neg.Update
   End If
End If

txtDataProc = DataProc

If Resp = vbYes Then
   glbFuncao = "frmPedido"
   glbNumPedido = neg!chNumPedido
   Call FechaDB
   frmProcessaPedido.Show vbModal
   'Call Carrega_Ordem_De_Carga
   cmdSair.SetFocus
Else
   Call FechaDB
   txtNumPedido.SetFocus
End If
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub cmdSairProduto_Click()

Call Rotina_Limpa_Detalhe
Call Limpa_Pedido
Call LimpaGridMedicao
Call Rotina_Desbloqueia_Campos

txtNumPedido.SetFocus
txtStatusPedido = Empty

cmdIncluiDetalhe = False
cmdAlteraDetalhe = False
cmdExcluiDetalhe = False
cmdNovoProduto = False
cmdAlteraPedido = False
cmdExcluiPedido = False
End Sub

Private Sub txtDataInicio_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
MsgBox "Função data a partir de não disponível " '
'If TxtAPartirDe = Empty Then
'   TxtAPartirDe = txtDataInicio - Date
'End If
End Sub

Private Sub Form_Load()

'ProdutoConsig(1) = "CSGR"
'ProdutoConsig(2) = "CSGP"

Dim SalvaCobFrete As String
txtDataPedido = Date
cmbEmissor.AddItem "SHE B"
cmbEmissor.ListIndex = 0

'optGeral = True
'optGrupo = False

VerificaData = 0

txtStatusPedido = Empty

'txtDataInicio = Date

dtInicio = Date
dtFim = Date

txtUnidade.AddItem "M2"
txtUnidade.AddItem "Un"
txtUnidade.AddItem "Hr"
txtUnidade.AddItem "HrN"
txtUnidade.ListIndex = 1

cmbStatusPedido.AddItem "PENDENTE"
cmbStatusPedido.AddItem "PROCESSADO"
cmbStatusPedido.AddItem "EM APROVAÇÃO"

cmbStatusPedido.ListIndex = 0

'cmbMotivacao.AddItem "Representante"
'cmbMotivacao.AddItem "Compra Direta"
'cmbMotivacao.ListIndex = 0

Call Rotina_AbrirBanco

Ativ.Open "Select * from atividade", db, 3, 3

Ativ.MoveFirst

Do While Not Ativ.EOF
   cmbAtividade.AddItem Ativ!atvAtividade
   Ativ.MoveNext
Loop

'cmbAtividade = Empty


pes.Open "Select * from pessoa", db, 3, 3

pes.MoveFirst
Do While Not pes.EOF
   If pes!pestipopessoa = 0 Then
      cmbPessoa.AddItem pes!chPessoa
      pes.MoveNext
   Else
      pes.MoveNext
   End If
Loop

'Carga de Tipo de Cobranca de Frete

'FreteCobranca.Open "Select * from cobrancafrete", db, 3, 3
'
'FreteCobranca.MoveFirst

'Do While Not FreteCobranca.EOF
'   cmbCobrancaFrete.AddItem FreteCobranca!parDescCobrancaFrete
'   FreteCobranca.MoveNext
'Loop

'cmbCobrancaFrete.ListIndex = 2

'FreteCobranca.Close: Set FreteCobranca = Nothing

'Carga de Condições de Processamento

'CondProc.Open "Select * from condprocessamento", db, 3, 3

'CondProc.MoveFirst

'Do While Not CondProc.EOF
'   cmbCondProcessamento.AddItem CondProc!cprDescCondProcess
'   CondProc.MoveNext
'Loop

'cmbCondProcessamento.ListIndex = 2

'CondProc.Close: Set CondProc = Nothing

'Carrega bancos

Bco.Open "Select * from banco", db, 3, 3

Bco.MoveFirst

Do While Not Bco.EOF
   cmbBanco.AddItem Bco!bcosiglabco
   Bco.MoveNext
Loop

Bco.Close: Set Bco = Nothing

cmbBanco.ListIndex = 0

'Carrega Natureza da Operação

NatuOper.Open "Select * from naturezaoperacao", db, 3, 3

NatuOper.MoveFirst

Do While Not NatuOper.EOF
   If NatuOper!Status = 1 Then
      cmbCFOP.AddItem NatuOper!cfop & "-" & NatuOper!natoperacaoabrev
      CFOPAux.AddItem NatuOper!cfop
   End If
   NatuOper.MoveNext
Loop

NatuOper.Close: Set NatuOper = Nothing

cmbCFOP.ListIndex = 0

Call Limpa_Pedido
'txtPreçoFixo = Empty
'txtFrete = Empty
cmdIncluiDetalhe.Enabled = False
cmdAlteraDetalhe.Enabled = False
cmdExcluiDetalhe.Enabled = False
cmdNovoProduto.Enabled = False
cmdSairProduto.Enabled = True
cmdAlteraPedido.Enabled = False
cmdExcluiPedido.Enabled = False
cmdSair.Enabled = True
'cmdNavega(0).Enabled = True
'cmdNavega(1).Enabled = True
'cmdNavega(2).Enabled = True
'cmdNavega(3).Enabled = True
'txtUnidade = "Dia"
'cmbCondProcessamento.ListIndex = 0
'If glbFuncao = "frmDesmembra_Pedido" Then
'   glbFuncao = Empty
'   txtNumPedido = glbNumPedido
'   txtComplementoPedido = glbCompPedido
'   TabNegociacao.Seek "=", txtNumPedido, txtComplementoPedido
'   Call Rotina_Carrega_Pedido
'   TabDetalheNegociacao.MoveFirst
'   Call Rotina_Carga_Grid
'   cmbProduto.Enabled = True
'   cmdProcessar.Enabled = True
'   cmdDesmembrar.Enabled = True
'End If

'Monte combo de pedidos enviados para a fábrica

For indPedido = 1 To 500
    Tabela_Pedido(indPedido) = Empty
Next
'Aqui Pedidos


neg.Open "Select * from negociacao", db, 3, 3
If Not (neg.EOF) Then
   Call Rotina_Carrega_Pedidos
   cmbPesqPedido.ListIndex = 0
End If
If (GlbStatus = "PROCESSADO" Or GlbStatus = "PENDENTE") And Not (txtNumPedido = Empty) Then
   txtStatusPedido = GlbStatus
   GlbStatus = Empty
   'txtNumPedido = frmProcessaPedido.txtNumPedido
   'txtComplementoPedido = frmProcessaPedido.txtCompPedido
   cmbPessoa.Enabled = True

   txtNumPedido.Enabled = True
   txtComplementoPedido.Enabled = True

   If txtNumPedido = Empty Then
      If Not (txtStatusPedido = "PROCESSADO") Then
         MsgBox "Informar o numero do pedido"
         txtNumPedido.SetFocus
         Exit Sub
      End If
   End If

   If txtComplementoPedido = Empty Then
      txtComplementoPedido = 0
   End If

   'neg.Seek "=", txtNumPedido, txtComplementoPedido

   If neg.EOF Then
      Inclui_Pedido = 1
      Inclui_Detalhe = 1
      cmdAlteraPedido.Enabled = False
      cmdExcluiPedido.Enabled = False
   
   Else
      cmdAlteraPedido.Enabled = True
      cmdExcluiPedido.Enabled = True
   
      Inclui_Pedido = 0
      
      dneg.Open "Select * from detalhenegociacao", db, 3, 3
'Aqui Pedido
      Call Rotina_Carrega_Pedido
      dneg.MoveFirst
      Call Rotina_Carga_Grid
   End If

End If

If txtStatusPedido = "PROCESSADO" Then
   txtDataProces = neg!negdatanegociação
   Call Rotina_Bloqueia_Campos
Else
   Call Rotina_Desbloqueia_Campos
End If

'Restaurar lgtTipo de logistica
If lgt.State = 1 Then
   lgt.Close: Set lgt = Nothing
End If

lgt.Open "Select * from logistica where lgtTipo = ('" & 2 & "')", db, 3, 3
If Not lgt.EOF Then
   lgt.MoveFirst
   Do While Not lgt.EOF
      lgt!lgtTipo = 0
      lgt.Update
      lgt.MoveNext
   Loop
End If

dtInicioMedicao = Date
dtFimMedicao = Date

Call FechaDB

End Sub




Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

coluna = Grid.Col
Linha = Grid.Row

If Linha > Grid.Rows Then
   MsgBox "Clicar somente em Linha com conteúdo."
   cmdSair.SetFocus
   Exit Sub
End If

If Grid.TextMatrix(Linha, 1) = Empty Then
   MsgBox "Clicar somente em Linha com conteúdo."
   cmdSair.SetFocus
   Exit Sub
End If

If txtStatusPedido = "PROCESSADO" Then
   MsgBox ("Função Válida somente para pedidos pendentes"), vbInformation
   Exit Sub
End If

cmbProduto = Grid.TextMatrix(Linha, 1)
txtNomeProduto = Grid.TextMatrix(Linha, 2)
cmbAtividade = Grid.TextMatrix(Linha, 3)
txtUnidade = Grid.TextMatrix(Linha, 4)
txtQtd = Format$(Grid.TextMatrix(Linha, 5), "0.00")
txtPUCheio = Format$(Grid.TextMatrix(Linha, 6), "##0.00")
txtDesconto = Format$(Grid.TextMatrix(Linha, 7), "##0.00")
txtPreçoUnit = Grid.TextMatrix(Linha, 8)
txtValorDiaria = Grid.TextMatrix(Linha, 9)
txtQtdDias = Grid.TextMatrix(Linha, 10)
dtInicio = Grid.TextMatrix(Linha, 12)
dtFim = Grid.TextMatrix(Linha, 13)
DataInicioAnter = dtInicio
DataFimAnter = dtFim
VerificaData = 1
cmdAlteraDetalhe.Enabled = True
cmdExcluiDetalhe.Enabled = True

cmbProduto.SetFocus


End Sub

Private Sub GridPedido_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim ind_Pedido As Integer
Dim ind_Linha As Integer
Dim ind_Coluna As Integer

Call Limpa_Pedido


ind_Linha = GridPedido.Row

txtNumPedido = GridPedido.TextMatrix(ind_Linha, 0)
txtComplementoPedido = GridPedido.TextMatrix(ind_Linha, 1)

If txtNumPedido = Empty Then
   MsgBox ("Clicar somente em linha com conteúdo"), vbCritical
   Exit Sub
End If
   
cmbPessoa.Enabled = True

txtNumPedido.Enabled = True
txtComplementoPedido.Enabled = True

If txtNumPedido = Empty Then
   MsgBox ("Informar o numero do pedido"), vbInformation
   txtNumPedido.SetFocus
End If

If txtComplementoPedido = Empty Then
   txtComplementoPedido = 0
End If

Call Rotina_AbrirBanco

neg.Open "Select * from negociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtComplementoPedido & "')", db, 3, 3
If neg.EOF Then
   Inclui_Pedido = 1
   Inclui_Detalhe = 1
   cmdAlteraPedido.Enabled = False
   cmdExcluiPedido.Enabled = False
Else
   cmdAlteraPedido.Enabled = True
   cmdExcluiPedido.Enabled = True
   pes.Open "Select * from pessoa where chPessoa = ('" & neg!chPessoa & "')", db, 3, 3
   If pes.EOF Then
      MsgBox ("Erro no acesso a pessoa em Grid_Pedido Click."), vbCritical
      Call FechaDB
      Exit Sub
   End If
End If
   
Inclui_Pedido = 0

Call Rotina_Carrega_Pedido

Call CargaProduto

Call Rotina_Carga_Grid

Call ResumoMedicao

txtDataPedido.Enabled = False
cmdSair.SetFocus

Call FechaDB

End Sub


Private Sub optImpSim_LostFocus()

Call Rotina_AbrirBanco

If optImpSim = True Then
   
   Call ConsultaMesInicioMedicao
   
   If DataFinalEvento = Empty Then
      Call ConsultaMesFinalEvento
   End If
   
   If DataInicioEvento = Empty Then
      dtInicio = Date
   Else
      dtInicio = DataInicioEvento
   End If
   
   dtFim = DataFinalEvento
   
End If

   
End Sub
Public Sub ConsultaMesInicioMedicao()

Call Rotina_AbrirBanco

mes = Month(dtInicioMedicao)
dataInicio = Year(dtInicioMedicao) & Format$(mes, "00")

dtInicioEventoMenosUm = dtInicioMedicao - 1

Dia = Day(dtInicioEventoMenosUm)
mes = Month(dtInicioEventoMenosUm)
ano = Year(dtInicioEventoMenosUm)

DataInicioEventoStr = Format$(ano & "-" & mes & "-" & Dia, "yyyy-mm-dd")


lgt.Open "Select * from logistica where chAnoMesRef = ('" & dataInicio & "') and lgtFimEventoReal > ('" & DataInicioEventoStr & "') and lgtTipo = ('" & 0 & "') and chPessoa = ('" & cmbPessoa & "') and chUnidadeOperacional = ('" & cmbLocal & "') and chColaborador = ('" & cmbProduto & "') and chEvento = ('" & cmbAtividade & "')", db, 3, 3
If lgt.EOF Then
   If lgt.State = 1 Then
      lgt.Close: Set lgt = Nothing
   End If
   mes = Month(dtFimMedicao)
   dataInicio = Year(dtFimMedicao) & Format$(mes, "00")
   lgt.Open "Select * from logistica where chAnoMesRef = ('" & dataInicio & "') and lgtTipo = ('" & 0 & "') and chPessoa = ('" & cmbPessoa & "') and chUnidadeOperacional = ('" & cmbLocal & "') and chColaborador = ('" & cmbProduto & "') and chEvento = ('" & cmbAtividade & "')", db, 3, 3
   If lgt.EOF Then
      DataInicioEvento = Empty
      txtQtd = Empty
      txtPUCheio = Empty
      txtDesconto = Empty
      cmbAtividade = Empty
      Call FechaDB
      Exit Sub
   End If
End If

If lgt!chInicioEvento > dtInicioMedicao Then
   DataInicioEvento = lgt!chInicioEvento
Else
   DataInicioEvento = dtInicioMedicao
End If

DataFinalEvento = lgt!lgtFimEventoReal

If DataFinalEvento > dtFimMedicao Then
   DataFinalEvento = dtFimMedicao
Else
   DataFinalEvento = lgt!lgtFimEventoReal
End If

If lgt!lgtTipo = 0 Then
   lgt!lgtTipo = 2
   lgt.Update
End If

Call FechaDB

End Sub

Public Sub ConsultaMesFinalEvento()

Call Rotina_AbrirBanco

mes = Month(dtFimMedicao)
dataFim = Year(dtFimMedicao) & Format$(mes, "00")

lgt.Open "Select * from logistica where chAnoMesRef = ('" & dataFim & "') and chPessoa = ('" & cmbPessoa & "') and chUnidadeOperacional = ('" & cmbLocal & "') and chColaborador = ('" & cmbProduto & "') and chEvento = ('" & cmbAtividade & "')", db, 3, 3
If lgt.EOF Then
   DataFinalEvento = dtFimMedicao
   Call FechaDB
   Exit Sub
End If

If DataInicioEvento = Empty Then
   DataInicioEvento = lgt!chInicioEvento
End If
DataFinalEvento = lgt!lgtFimEvento

End Sub

Private Sub txtComplementoPedido_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtComplementoPedido_LostFocus()
On Error Resume Next

cmbPessoa.Enabled = True

txtNumPedido.Enabled = True
txtComplementoPedido.Enabled = True

If Not txtNumPedido = Empty Then
   If txtComplementoPedido = Empty Then
      txtComplementoPedido = 0
   End If
End If

Call Rotina_AbrirBanco

neg.Open "Select * from negociacao where chnumpedido = ('" & txtNumPedido & "') and chnumpedidocomp = ('" & txtComplementoPedido & "')", db, 3.3

If neg.EOF Then
   Inclui_Pedido = 1
   Inclui_Detalhe = 1
   txtDataPedido.Enabled = True
   txtDataPedido = Date
   cmdAlteraPedido.Enabled = False
   cmdExcluiPedido.Enabled = False
Else
   pes.Open "Select * from pessoa where chPessoa = ('" & neg!chPessoa & "')", db, 3, 3
   If pes.EOF Then
      MsgBox ("Erro na carga de pessoa em Negociação"), vbCritical
      Call FechaDB
      Exit Sub
   End If
   
   cmdAlteraPedido.Enabled = True
   cmdExcluiPedido.Enabled = True
   
   Inclui_Pedido = 0

   Call Rotina_Carrega_Pedido
   
   dneg.Open "Select * from detalhenegociacao", db, 3, 3
     
   dneg.MoveFirst
   Call Rotina_Carga_Grid
   txtDataPedido.Enabled = False
   
   Call ResumoMedicao
   
   If cmbProduto = Empty Then
      Call CargaProduto
   End If
   
   If Inclui_Pedido = 0 Then
      cmbProduto.SetFocus
   End If
   
   Call FechaDB
   
End If
End Sub
Public Sub Limpa_Pedido()

Acumula_Produto = 0
Acumula_Frete = 0
Acumula_Desconto = 0
Valor_Operacao = 0

txtNumPedido = Empty
txtComplementoPedido = Empty
cmbPessoa = Empty
txtEndereco = Empty
txtBairro = Empty
txtCidade = Empty
txtUF = Empty
'txtCEP = Empty
'txtTel = Empty
txtQtdFat = Empty
txtIntervalo = Empty
TxtAPartirDe = Empty
'txtPrzBoletaFrete = Empty
'cmbCobrancaFrete.ListIndex = 2
'txtPreçoFixo = Empty

'cmbCondProcessamento.ListIndex = 0
'txtDescComissao = Empty
cmbProduto = Empty
txtNomeProduto = Empty
txtUnidade.ListIndex = 1
txtQtd = Empty
txtPreçoUnit = Empty
txtPUCheio = Empty

'txtDesc = Empty

'txtTotalMetros = Empty
'txtTotalCaixa = Empty
'txtTotalPeso = Empty
txtAcumula_Produto = Empty
txtAcumula_Desconto = Empty
txtValorComDesconto = Empty
'txtAcumula_Frete = Empty
'txtValor_Operacao = Empty

'txtPreçoFixo = Empty

cmbBanco.ListIndex = 0
'cmbOrdemDeCarga = Empty
'lblTransporte = Empty
'lblPlaca = Empty
txtNotaFiscal = Empty
'txtDataInicio = Date

cmbBanco.ListIndex = 0
'cmbMotivacao.Enabled = True
'cmbMotivacao.ListIndex = 0
txtRepresentante = Empty
txtPromotora = Empty

txtCNPJCPF.Mask = "##.###.###/####-##"
txtCNPJCPF = "__.___.___/____-__"

Grid.Rows = 2
Linha = 1
Grid.TextMatrix(Linha, 0) = Empty
Grid.TextMatrix(Linha, 1) = Empty
Grid.TextMatrix(Linha, 2) = Empty
Grid.TextMatrix(Linha, 3) = Empty
Grid.TextMatrix(Linha, 4) = Empty
Grid.TextMatrix(Linha, 5) = Empty
Grid.TextMatrix(Linha, 6) = Empty
Grid.TextMatrix(Linha, 7) = Empty
Grid.TextMatrix(Linha, 8) = Empty
Grid.TextMatrix(Linha, 9) = Empty
Grid.TextMatrix(Linha, 10) = Empty
Grid.TextMatrix(Linha, 11) = Empty
Grid.TextMatrix(Linha, 12) = Empty
Grid.TextMatrix(Linha, 13) = Empty

'cmbOrdemDeCarga.Clear

'cmbOrdemDeCarga.AddItem "Cliente"

'cmbOrdemDeCarga.ListIndex = 0

End Sub
Public Sub Rotina_Mover_Negociacao()

Dim indCondProcess As Long


If txtDataPedido = "__/__/____" Then
   MsgBox ("Data do pedido invalida")
   txtDataPedido.SetFocus
   Exit Sub
End If

neg!negDataPedido = txtDataPedido
neg!negdatanegociação = 0
neg!chPessoa = cmbPessoa
neg!negContrato = cmbContrato
neg!chUnidadeOperacional = cmbLocal
neg!chNumPedido = txtNumPedido
neg!chNumPedidoComp = txtComplementoPedido
neg!negFaturamento = txtQtdFat
neg!negIntervaloFatura = txtIntervalo
neg!negAPartirDe = TxtAPartirDe
'neg!negCobrancaFrete = cmbCobrancaFrete.ListIndex
'neg!negBoletaFrete = txtPrzBoletaFrete
'If cmbCobrancaFrete.ListIndex = 4 Or cmbCobrancaFrete.ListIndex = 5 Or cmbCobrancaFrete.ListIndex = 7 Then
'   neg!negValorFixoFrete = txtPreçoFixo
'Else
'   neg!negValorFixoFrete = 0
'   txtPreçoFixo.Enabled = False
'End If
neg!chrepresentante = txtRepresentante
neg!chPromotor = txtPromotora

neg!negTipoProduto = TipoProduto

'neg!negCondProcess = cmbCondProcessamento.ListIndex

'If txtDescComissao = Empty Then
'   txtDescComissao = 0
'End If
'neg!negDescComissao = txtDescComissao
neg!negStatus = 0
'neg!negDataVencimento = Empty
neg!negLançamento = Empty
neg!negUltimaAtualizacao = Empty
'neg!negMotivacao = cmbMotivacao.ListIndex
'neg!negCEFOP = Mid$(cmbCFOP, 1, 4)
neg!negCEFOP = cmbCFOP.ListIndex
If pes.State = 1 Then
   pes.Close: Set pes = Nothing
End If
pes.Open "Select * from pessoa where chPessoa = ('" & cmbPessoa & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Erro em acesso a pessoa em: txtUnidade.setfocus"), vbCritical
   Call FechaDB
   Exit Sub
End If
If pes!pesClassFiscal = "Lucro Real" Then
   neg!negSerieFatura = "A"
Else
   neg!negSerieFatura = "E"
End If

neg!negInicioMedicao = dtInicioMedicao
neg!negFinalMedicao = dtFimMedicao

End Sub
Public Sub Rotina_Mover_Detalhe()
Dim Comissao As Currency

dneg!chNumPedido = txtNumPedido
dneg!chNumPedidoComp = txtComplementoPedido
dneg!chProduto = cmbProduto
dneg!pedAtividade = cmbAtividade
dneg!chDataInicio = dtInicio
dneg!chDataFim = dtFim

dneg!pedPrecoUnidadePedida = txtPUCheio - ((txtPUCheio * txtDesconto) / 100)
dneg!pedquantidadePedida = txtQtd

If txtUnidade = "M2" Then
   txtUnidade.ListIndex = 0
Else
   If txtUnidade = "Un" Then
      txtUnidade.ListIndex = 1
   Else
      If txtUnidade = "Hr" Then
         txtUnidade.ListIndex = 2
      Else
         txtUnidade.ListIndex = 3
      End If
   End If
End If

dneg!pedunidade = txtUnidade.ListIndex
dneg!pedPUCheio = txtPUCheio

If (txtUnidade.ListIndex = 2) Or (txtUnidade.ListIndex = 3) Then
   dneg!pedValorDaDiaria = txtValorDiaria
Else
   dneg!pedValorDaDiaria = precoUnit * txtQtd
End If

dneg!pedqtddias = txtQtdDias
dneg!pedValorDaOperacao = Format$(txtQtdDias * dneg!pedValorDaDiaria, "##,##0.00")
If txtRepresentante = "NENHUM" Then
   dneg!pedcomissaorep = Format$(0, "#.00")
Else
      dneg!pedcomissaorep = Format$((Comissao * (ComisRep + AjustaComis) / 100), "#.00")
     ' - txtDescComissao) / 100), "#.00")
End If

If txtPromotora = "NENHUM" Then
   dneg!pedcomissaopromot = Format$(0, "#.00")
Else
   dneg!pedcomissaopromot = Format$((Comissao * ComisPromot / 100), "#.00")
End If

dneg!pedquantidadePedida = txtQtd
'dneg!pedPrecoUnidadePedida = txtPreçoUnit
dneg!pedDesconto = Format$(txtDesconto, "#00.00")
dneg!pedValorDesconto = ((txtDesconto * ((txtPUCheio * txtQtd) * txtQtdDias)) / 100)

End Sub
Public Sub Rotina_Carrega_Pedido()

Dim indCondProcess As Integer

If neg!negStatus = 1 Then
   txtStatusPedido = "PROCESSADO"
   txtDataProces = neg!negdatanegociação
   cmdProcessar.Enabled = False
   cmdAlteraPedido.Enabled = False
   'cmdDesmembrar.Enabled = False
Else
   If neg!negStatus = 2 Then
      txtStatusPedido = "EM APROVAÇÃO"
      txtDataProces = neg!negDataEnvioAprovMedicao
      cmdProcessar.Enabled = True
      cmdAlteraPedido.Enabled = True
   Else
      cmdProcessar.Enabled = True
      'cmdDesmembrar.Enabled = True
      cmdAlteraPedido.Enabled = True
      txtStatusPedido = "PENDENTE"
      'txtDataProces = Date
   End If
End If

If Not IsNull(neg!negInicioMedicao) Then
   dtInicioMedicao = neg!negInicioMedicao
Else
   dtInicioMedicao = Date
End If
If Not IsNull(neg!negFinalMedicao) Then
   dtFimMedicao = neg!negFinalMedicao
Else
   dtFimMedicao = Date
End If
   
txtDataPedido = neg!negDataPedido

cmbPessoa = neg!chPessoa
cmbContrato = neg!negContrato

txtNumPedido = neg!chNumPedido
txtComplementoPedido = neg!chNumPedidoComp
txtQtdFat = neg!negFaturamento
TxtAPartirDe = neg!negAPartirDe
txtIntervalo = neg!negIntervaloFatura
If Not (neg!chPessoa = Empty) Then
  cmbPessoa = neg!chPessoa
Else
  cmbPessoa = Empty
End If
If Not (neg!chUnidadeOperacional = Empty) Then
   lblLocalizacao = neg!chUnidadeOperacional
   cmbLocal = neg!chUnidadeOperacional
   SalvaLocal = neg!chUnidadeOperacional
Else
   lblLocalizacao = Empty
   cmbLocal = Empty
   SalvaLocal = Empty
End If
'cmbCobrancaFrete.ListIndex = neg!negCobrancaFrete
'txtPrzBoletaFrete = neg!negBoletaFrete

'cmbCondProcessamento.ListIndex = neg!negCondProcess

'txtDescComissao = neg!negDescComissao

cmbEmissor.ListIndex = 0

'If neg!negCobrancaFrete = 2 Then
'      lblTransporte = "Cliente"
'      lblPlaca = "Cliente"
'Else
'   If IsNull(neg!negTransporte) Then
'      lblTransporte = Empty
'   Else
'      lblTransporte = neg!negTransporte
'   End If
'End If

If IsNull(neg!negNotaFiscal) Then
   txtNotaFiscal = Empty
Else
   txtNotaFiscal = neg!negNotaFiscal
End If

If IsNull(neg!negPlaca) Then
   lblPlaca = Empty
Else
   lblPlaca = neg!negPlaca
End If

'If IsNull(neg!chordemdecarga) Then
'   cmbOrdemDeCarga = Empty
'Else
'   cmbOrdemDeCarga = neg!chordemdecarga
'End If

'cmbMotivacao.ListIndex = neg!negMotivacao

If neg!negStatus = 1 Or neg!negStatus = 2 Then
   
CFOPAux.ListIndex = neg!negCEFOP
   
NatuOper.Open "Select * from naturezaoperacao where CFOP = ('" & CFOPAux & "')", db, 3, 3
If NatuOper.EOF Then
      MsgBox ("Atenção: Negociação Processada sem informar a sua Natureza"), vbInformation
      cmbCFOP.ListIndex = 0
   Else
      cmbCFOP = NatuOper!cfop & "-" & NatuOper!natoperacaoabrev
   End If
Else
   cmbCFOP.ListIndex = neg!negCEFOP
End If

Bco.Open "Select * from banco where bcoSiglaBco = ('" & cmbBanco & "')", db, 3, 3
If Bco.EOF Then
   MsgBox ("Banco informado Inv."), vbInformation
Else
   cmbBanco = Bco!bcosiglabco
End If

'TabEntrega.Seek "=", txtNumPedido, txtComplementoPedido



txtEndereco = pes!pesEndereco
txtBairro = pes!pesBairro
txtCidade = pes!pesCidade
txtUF = pes!chUF
'txtCEP = pes!pesCEP
'TabTelefone.Seek "=", cmbPessoa, "Tel-1"
'      If TabTelefone.NoMatch Then
'         txtTel = "N/INFORMADO"
'      Else
 '        txtTel = TabTelefone("codigocontato")
'      End If
'   End If
'Else
'   txtEndereco = TabEntrega("entEndereço")
'   txtBairro = TabEntrega("entBairro")
'   txtCidade = TabEntrega("entCidade")
'   txtUF = TabEntrega("ENTUF")
'   txtCEP = TabEntrega("entCEP")
'   txtTel = TabEntrega("entTel")
'End If

If txtStatusPedido = "PROCESSADO" Then
   Call Rotina_Bloqueia_Campos
Else
   Call Rotina_Desbloqueia_Campos
End If

'TabCarteira_Rep.Seek "=", Tabpessoa("chcarteirarep")
'If TabCarteira_Rep.NoMatch Then
'   MsgBox "Representante não cadastrado. Cadastrar Representante e retornar para confecção do pedido"
'   Unload Me
'   Exit Sub
'End If
'TabCarteira_Promot.Seek "=", Tabpessoa("chcarteiraPromot")
'If TabCarteira_Promot.NoMatch Then
'   MsgBox "Promotora não cadastrado. Cadastrar Promotora e retornar para confecção do pedido"
'   Unload Me
 '  Exit Sub
'End If
 If pes!pesPessoa = 0 Then
    txtCNPJCPF.Mask = "###.###.###-##"
    txtCNPJCPF = pes!chCNPJ_CPF
 Else
   txtCNPJCPF.Mask = "##.###.###/####-##"
   txtCNPJCPF = pes!chCNPJ_CPF
End If
txtRepresentante = neg!chrepresentante
txtPromotora = neg!chPromotor

End Sub
Public Sub Rotina_Carga_Grid()

Dim Fim_Carga As Byte
Linha = 0
Acumula_Produto = 0
Acumula_Desconto = 0
Acumula_Frete = 0
Valor_Operacao = 0
Valor_Frete = 0
Acumula_Operacao = 0
Acumula_Metro = 0
'Acumula_IPI = 0
Acumula_Caixa = 0
Acumula_Peso = 0
Fim_Carga = 0


Call Rotina_AbrirBanco


Grid.ColAlignment(1) = 1
Grid.ColAlignment(2) = 1

Encontrei = 0

dneg.Open "Select * from detalhenegociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtComplementoPedido & "')", db, 3, 3


If dneg.EOF Then
   MsgBox ("Registro de negociação sem detalhe"), vbCritical
   Encontrei = 2
   Fim_Carga = 1
Else
   Fim_Carga = 0
   Pedido = dneg!chNumPedido
   PedidoComp = dneg!chNumPedidoComp
   Produto = dneg!chProduto
End If

Do While Fim_Carga = 0
   
   If dneg!chNumPedido <> txtNumPedido Then
      Fim_Carga = 1
   Else
      If dneg!chNumPedidoComp <> txtComplementoPedido Then
         Fim_Carga = 1
      Else
         Linha = Linha + 1
         Grid.Rows = Linha + 1
         Grid.TextMatrix(Linha, 1) = dneg!chProduto
         If pes.State = 1 Then
            pes.Close: Set pes = Nothing
         End If
         pes.Open "Select * from pessoa where chPessoa = ('" & dneg!chProduto & "')", db, 3, 3
         If pes.EOF Then
            Prod.Open "Select * from produto where chProduto = ('" & dneg!chProduto & "')", db, 3, 3
            If Prod.EOF Then
               MsgBox ("Produto não encontrado Rotina Carga Grid"), vbCritical
               Call FechaDB
               Exit Sub
            Else
               Grid.TextMatrix(Linha, 2) = Prod!prdNomeProd
               Grid.TextMatrix(Linha, 0) = (0 & dneg!chProduto)
               Prod.Close: Set Prod = Nothing
            End If
         Else
            Grid.TextMatrix(Linha, 2) = pes!chPessoa
            Grid.TextMatrix(Linha, 0) = (0 & dneg!chProduto)
            'Prod.Close: Set Prod = Nothing
         End If
         txtUnidade.ListIndex = dneg!pedunidade
         Grid.TextMatrix(Linha, 3) = dneg!pedAtividade
         Grid.TextMatrix(Linha, 4) = txtUnidade
         If (txtUnidade.ListIndex = 2) Or (txtUnidade.ListIndex = 3) Then
            Grid.TextMatrix(Linha, 5) = Format$((dneg!pedquantidadePedida), "#0.00")
         Else
            Grid.TextMatrix(Linha, 5) = Format$((dneg!pedquantidadePedida), "##0")
         End If
         Grid.TextMatrix(Linha, 6) = Format$((dneg!pedPUCheio), "##0.00")
         Grid.TextMatrix(Linha, 7) = Format$(dneg!pedDesconto, "##0.00")
         Grid.TextMatrix(Linha, 8) = Format$((dneg!pedPrecoUnidadePedida), "##0.00")
         Grid.TextMatrix(Linha, 9) = Format$((dneg!pedValorDaDiaria), "#0.00")
         Grid.TextMatrix(Linha, 10) = Format$((dneg!pedqtddias), "##0")
         Grid.TextMatrix(Linha, 11) = Format$((dneg!pedValorDaDiaria * dneg!pedqtddias), "#,##0.00")
         If Not dneg!chDataInicio = Empty Then
            Grid.TextMatrix(Linha, 12) = Format$(dneg!chDataInicio, "dd/mm/yy")
         Else
            Grid.TextMatrix(Linha, 12) = Format$(Date, "dd/mm/yy")
         End If
         If Not dneg!chDataFim = Empty Then
            Grid.TextMatrix(Linha, 13) = Format(dneg!chDataFim, "dd/mm/yy")
         Else
            Grid.TextMatrix(Linha, 13) = Format$(Date, "dd/mm/yy")
         End If

         UltimaLinha = Linha
         'Teste aqui
         'Alteracao em 19/01/05 formato em acumula Produto
         If Not (dneg!pedunidade = 2) Then
            Acumula_Produto = Format$(Acumula_Produto + ((dneg!pedPUCheio * dneg!pedquantidadePedida * dneg!pedqtddias)), "###,###,##0.00")
            Acumula_Operacao = Format$((Acumula_Operacao + (((dneg!pedPrecoUnidadePedida * dneg!pedquantidadePedida) * dneg!pedqtddias))), "###,###,##0.00")
            Acumula_Desconto = Format$(Acumula_Desconto + dneg!pedValorDesconto, "###,###,##0.00")
         Else
            Acumula_Produto = Format$(Acumula_Produto + dneg!pedValorDaOperacao, "###,###,##0.00")
            Acumula_Operacao = Format$(Acumula_Operacao + dneg!pedValorDaOperacao, "###,###,##0.00")
            Acumula_Desconto = Format$(Acumula_Desconto + dneg!pedValorDesconto, "###,###,##0.00")
         End If
         
          
         'Grid.TextMatrix(Linha, 9) = Format$(Valor_Operacao, "###,##0.00")
                
         dneg.MoveNext
         
         If dneg.EOF Then
            Fim_Carga = 1
         End If
      End If
End If
   
Loop
'txtTotalMetros = Format$(Acumula_Metro, "#,##0.00")
Grid.Col = 0
Grid.ColSel = 0
     
Grid.Row = 1
Grid.RowSel = Linha
        
If Linha > 1 Then
   Grid.Sort = 5
End If

txtAcumula_Produto = Format$(Acumula_Produto, "###,##0.00")

txtAcumula_Desconto = Format$(Acumula_Desconto, "###,##0.00")

txtValorComDesconto = Format$(Acumula_Produto - Acumula_Desconto, "###,##0.00")

'txtValor_Operacao = Format$((Acumula_Operacao), "###,##0.00")

Call FechaDB

End Sub

'Private Sub txtDataInicio_LostFocus()
'If TxtAPartirDe = Empty Then
'   TxtAPartirDe = txtDataInicio - Date
'End If
'End Sub

Private Sub txtDataPedido_LostFocus()

DataPedido = txtDataPedido

If DataPedido > Date Then
   MsgBox "Atenção: Pedido com data posterior a HOJE. Não Permitido"
   cmdSair.SetFocus
End If

End Sub

Private Sub txtDesconto_LostFocus()

If txtDesconto = "" Then
   txtDesconto = 0
   End If

If txtDesconto > 0 Then
   precoUnit = txtPUCheio - ((txtPUCheio * txtDesconto) / 100)
   txtPreçoUnit = Format(precoUnit, "##0.00")
   txtValorDiaria = Format((precoUnit * txtQtd), "##0.00")
Else
   If Not (txtUnidade = "Hr") Then
      If Not (txtUnidade = "HrN") Then
         txtPreçoUnit = Format(txtPUCheio, "##0.00")
         txtValorDiaria = Format((txtPreçoUnit * txtQtd), "##0.00")
      End If
   End If
End If
End Sub

'Private Sub txtDescComissao_GotFocus()
'If txtDescComissao = Empty Then
'   RecalculaComissao = 99
'Else
'   RecalculaComissao = txtDescComissao
'End If
'End Sub

'Private Sub txtDescComissao_LostFocus()
'If RecalculaComissao = 99 Then
'   RedutorDeComissao = 0
'Else
'   If Not RecalculaComissao = txtDescComissao Then
'      RedutorDeComissao = 1
'   End If
'End If
'End Sub

Private Sub txtNumPedido_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtQtd_LostFocus()

If txtUnidade.ListIndex = 2 Or txtUnidade.ListIndex = 3 Then
   Call CalculaMinutos
   'ValorHora = ValorHora * QtdHoras-
   'ValorMinuto = ValorMinuto * QtdMinutos
   'txtValorDiaria = ValorHora + ValorMinuto
   txtValorDiaria = Format$((ValorHora * QtdHoras) + (QtdMinutos * ValorMinuto), "##0.00")
Else
   If Not txtQtd = "" Then ' Null Or txtQtd = 0 Or txtQtd = "" Then
      If Not txtPreçoUnit = "" Then
         txtValorDiaria = Format(txtQtd * txtPreçoUnit, "##0.00")
        ' txtQtdDias.SetFocus
      'Else
      '   txtPreçoUnit.SetFocus
      End If
   End If
End If
'cmdCalculaDias.SetFocus
End Sub

Private Sub txtQtdDias_LostFocus()
If txtQtdDias = Empty Then
  'MsgBox "Quantidade de dias tem que ser informado com um valor diferente de zero"
   cmdSair.SetFocus
Else
   If txtQtdDias = 0 Then
      MsgBox "Quantidade de dias tem que ser informado com um valor diferente de zero"
      txtQtdDias.SetFocus
   End If
End If
End Sub

Private Sub txtUnidade_LostFocus()
   
   Call Rotina_AbrirBanco
   
   Fim = 0
   If cmbProduto = Empty Then
      Exit Sub
   Else
       pes.Open "Select * from pessoa where chPessoa = ('" & cmbProduto & "')", db, 3, 3
       If pes.EOF Then
          Prod.Open "Select * from produto where chProduto = ('" & cmbProduto & "')", db, 3, 3
       Else
          Prod.Open "Select * from produto where chProduto = ('" & cmbContrato & "')", db, 3, 3
       End If
       If Prod.EOF Then
          MsgBox ("Erro na carga do produto. Comunicar ao analista responsável."), vbCritical
          Call FechaDB
          Unload Me
          Exit Sub
       End If
       
       If Prod!prdOrdemApresentacao = 1 Then
          If Not (txtUnidade.ListIndex = 0 Or txtUnidade.ListIndex = 1) Then
             MsgBox ("Atividade incompatível com o Produto informado"), vbCritical
             Call FechaDB
             cmbProduto.SetFocus
             Exit Sub
          End If
       End If
       If pes.State = 1 Then
          pes.Close: Set pes = Nothing
       End If
       pes.Open "Select * from pessoa where chPessoa = ('" & cmbPessoa & "')", db, 3, 3
       If pes.EOF Then
          MsgBox ("Erro em acesso a pessoa em: txtUnidade.setfocus"), vbCritical
          Call FechaDB
          Exit Sub
       End If
         
       If Prod!prdUnidadeOperacional = "CONTRATO" Then
          ProdPco.Open "Select * from produtopreco where chPessoa = ('" & pes!chPessoa & "') and chProduto = ('" & cmbContrato & "') AND chAtividade = ('" & cmbAtividade & "') and pdpStatus = 0", db, 3, 3
       Else
         ProdPco.Open "Select * from produtopreco where chPessoa = ('" & pes!chPessoa & "') and chProduto = ('" & cmbProduto & "') AND chAtividade = ('" & cmbAtividade & "') and pdpStatus = 0", db, 3, 3
       End If
       If ProdPco.EOF Then
          If Not (cmbAtividade = Empty) Then
             MsgBox ("Preço não cadastrado. Efetue primeiramente o cadastro do preço e retorne a esta função"), vbInformation
             Call FechaDB
             Unload Me
             Exit Sub
          Else
             Call FechaDB
             Exit Sub
          End If
          'Resp = MsgBox("Preço não cadastrado. Deseja cadastrar agora???", vbYesNo)
          'If Resp = vbYes Then
          '   fim = 0
          '   frmAtualizaPrecoProd.cmbTabPreco = pes!chtabprecoproduto
          '   frmAtualizaPrecoProd.cmbProduto = cmbProduto
          '   frmAtualizaPrecoProd.txtAjuste = "999"
          '   frmAtualizaPrecoProd.Show vbModal
          '   cmbProduto.SetFocus 'txtPreçoUnit = Format(tabProdPrecoAtivo("pdpprecodoproduto"), "##0.00")
          'Else
          '   txtQtd.SetFocus
          'End If
       Else
          If (txtUnidade.ListIndex = 2) Or (txtUnidade.ListIndex = 3) Then
              ValorHora = Format$(((ProdPco!pdpPrecoDoProduto / 12) * 2), "#,##0.00")
              If (txtUnidade.ListIndex = 3) Then
                  ValorHora = Format$(ValorHora + ((ValorHora * 20 / 100)), "##0.00")
              End If
              ValorMinuto = ValorHora / 60
              txtPUCheio = Format$(ValorHora, "##0.00")
              txtDesconto = Format$(ProdPco!pdpDesconto, "##0.00")

              'txtPreçoUnit = Format$((ProdPco!pdpPrecoDoProduto) - (txtDesconto), "###,##0.00")
           Else
              txtPUCheio = Format(ProdPco!pdpPrecoDoProduto, "##0.00")
              txtPreçoUnit = Format(ProdPco!pdpPrecoDoProduto - ((ProdPco!pdpPrecoDoProduto * ProdPco!pdpDesconto) / 100), "###,##0.00")
              precoUnit = txtPreçoUnit
           End If
        End If

   End If
   
   txtQtd.SetFocus

   Call FechaDB
   
End Sub
'Public Sub Rotina_Carrega_Entrega()

'      TabEntrega("chNumPedido") = txtNumPedido
'      TabEntrega("chNumPedidoComp") = txtComplementoPedido
'      TabEntrega("entCliente") = cmbPessoa
'      TabEntrega("entEndereço") = txtEndereco
'      TabEntrega("entBairro") = txtBairro
'      TabEntrega("entCidade") = txtCidade
'      TabEntrega("entUF") = txtUF
'      TabEntrega("entCEP") = txtCEP
'      TabEntrega("entTel") = txtTel
'End Sub
Public Sub Rotina_Limpa_Detalhe()
cmbProduto = Empty
cmbAtividade.ListIndex = 0
txtNomeProduto = Empty
txtUnidade.ListIndex = 0
txtQtd = Empty
txtPreçoUnit = Empty
txtPUCheio = Empty
txtQtdDias = Empty
txtValorDiaria = Empty
txtDesconto = Format$(0, "#0.00")
txtValorDiaria = Empty

End Sub
Public Sub Rotina_Limpa_Grid()
Grid.Rows = 2
Linha = 1
    Grid.TextMatrix(Linha, 1) = Empty
    Grid.TextMatrix(Linha, 2) = Empty
    Grid.TextMatrix(Linha, 3) = Empty
    Grid.TextMatrix(Linha, 4) = Empty
    Grid.TextMatrix(Linha, 5) = Empty
    Grid.TextMatrix(Linha, 6) = Empty
    Grid.TextMatrix(Linha, 7) = Empty
    Grid.TextMatrix(Linha, 8) = Empty
    Grid.TextMatrix(Linha, 9) = Empty
    Grid.TextMatrix(Linha, 10) = Empty
    Grid.TextMatrix(Linha, 11) = Empty
    Grid.TextMatrix(Linha, 12) = Empty
    Grid.TextMatrix(Linha, 13) = Empty
End Sub
'Public Sub Rotina_Exclui_Entrega()

'TabEntrega.Seek "=", TabNegociacao("chNumPedido"), TabNegociacao("chNumPedidoComp")
'If TabEntrega.NoMatch Then
'   Exit Sub
'Else
'   db.begintrans
'   TabEntrega.Delete
'  db.CommitTrans
'End If
'End Sub
Public Sub Rotina_Exclui_Detalhe()

Dim ArrayProduto(99) As String
Dim IndProd As Integer
Acumula_Comis_Rep = 0
Acumula_Comis_Promot = 0

IndProd = 0
VerificaData = 0
'Call Rotina_AbrirBanco

dneg.Open "Select * from detalhenegociacao where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
If dneg.EOF Then
   Exit Sub
End If

dneg.MoveFirst

Do While Not dneg.EOF
   dneg.Delete
   dneg.MoveNext
Loop
'If TabDetalheNegociacao("chNumPedido") = txtNumPedido And TabDetalheNegociacao("chNumPedidoComp") = txtComplementoPedido Then
'   IndProd = IndProd + 1
'   ArrayProduto(IndProd) = TabDetalheNegociacao("chProduto")
'   Acumula_Comis_Rep = Acumula_Comis_Rep + TabDetalheNegociacao("pedcomissaorep")
'   Acumula_Comis_Promot = Acumula_Comis_Promot + TabDetalheNegociacao("pedcomissaopromot")
'   TabDetalheNegociacao.MoveNext
'Else
'   TabDetalheNegociacao.MoveNext
'End If
'Loop
'IndProd = IndProd + 1
'ArrayProduto(IndProd) = "Fim"'

'Descarrega tabela
'For IndProd = 1 To 99

'If ArrayProduto(IndProd) = "Fim" Then
'   ArrayProduto(IndProd) = Empty
'   IndProd = 99
'Else
'   TabDetalheNegociacao.Seek "=", TabNegociacao("chNumPedido"), TabNegociacao("chNumPedidoComp"), ArrayProduto(IndProd)
'   If TabDetalheNegociacao.NoMatch Then
'      MsgBox "Deu caquinha"
'      End
'   Else
'
'''      'Atualizar pedidos em carteira de pedidos
'      If TabNegociacao("negstatus") = 1 Or TabNegociacao("negstatus") = 2 Then
''         Funcao = 6
'         Ano = Year(Data_Hoje)
'         Mes = Month(Data_Hoje)
'         Produto = TabDetalheNegociacao("chproduto")
  
'         Sai = 0
'         Entra = TabDetalheNegociacao("pedQuantidadepedida")
      
'         TracoIn = 0
'         TracoOut = 0
'         Mes_Pedido = Month(TabNegociacao("negDataPedido"))
   
'         Call Rotina_Atualiza_Estoque(Funcao, Ano, Mes, Produto, Entra, Sai, TracoIn, TracoOut, Mes_Pedido)
 
         'TabDetalheNegociacao.Delete
'         ArrayProduto(IndProd) = Empty
'     Else
'         Funcao = 2
'         Ano = Year(Data_Hoje)
'         Mes = Month(Data_Hoje)
'         Produto = TabDetalheNegociacao("chproduto")
  
'         Sai = TabDetalheNegociacao("pedQuantidadepedida")
'         Entra = 0
      
 '        TracoIn = 0
 '        TracoOut = 0
 '        Mes_Pedido = Month(TabNegociacao("negDataPedido"))
   
 '        Call Rotina_Atualiza_Estoque(Funcao, Ano, Mes, Produto, Entra, Sai, TracoIn, TracoOut, Mes_Pedido)
   
 '        TabDetalheNegociacao.Delete
 '        ArrayProduto(IndProd) = Empty
  '    End If
  ' End If
'End If
'Next


'Call FechaDB

End Sub

Public Sub Rotina_Carrega_Pedidos()

'neg.Open "Select * from negociacao", db, 3, 3

Tabela_Pedido(1) = " Geral"
cmbPesqPedido.AddItem " Geral"

If neg.BOF Then
   Exit Sub
Else
   neg.MoveFirst

   Do While Not neg.EOF
   
   indPedido = 1
   
      Do While Not indPedido = 500
         If Tabela_Pedido(indPedido) = neg!chPessoa Then
            neg.MoveNext
            indPedido = 500
         Else
            If Tabela_Pedido(indPedido) = Empty Then
               Tabela_Pedido(indPedido) = neg!chPessoa
               cmbPesqPedido.AddItem neg!chPessoa
               indPedido = 500
               neg.MoveNext
            Else
               indPedido = indPedido + 1
            End If
         End If
      Loop
  
Loop
frmPedido.Refresh
End If
End Sub

'Public Sub Rotina_Recalcula_CondProc()
'Dim A As Integer
'Dim Comissao As Currency
'fim = 0
'A = 1

'TabCondProcessamento.Seek "=", cmbCondProcessamento.ListIndex
'If TabCondProcessamento.NoMatch Then
'   MsgBox ("Condicao de processamento inválido")
'   cmbCondProcessamento.SetFocus
'   Exit Sub
'End If


'If cmbCondProcessamento.ListIndex > 1 Then
'   IndiceComissao = 0
'Else
'   IndiceComissao = 1
'End If

'For A = 1 To UltimaLinha
'    TabDetalheNegociacao.Seek "=", txtNumPedido, txtComplementoPedido, Grid.TextMatrix(A, 1)
'    If TabDetalheNegociacao.NoMatch Then
'       MsgBox ("Detalhe de negociacao nao encontrado"), txtNumPedido
'       A = 1 / 0
'    End If
'
'    TabDetalheNegociacao.Edit
'
'    tabproduto.Seek "=", TabDetalheNegociacao("chProduto")
'
'    If tabproduto.NoMatch Then
'       MsgBox ("Produto nao encontrado.")
'       A = 1 / 0
'    End If
'
'    TabDetalheNegociacao("pedPrecoUnidadePedida") = TabDetalheNegociacao("pedPrecoUnidadePedida") * TabCondProcessamento("cprIncideValor")
'
'    If txtUnidade = "Dia" Then
'       TabDetalheNegociacao("pedPrecoMetro") = TabDetalheNegociacao("pedPrecoUnidadePedida")
'    Else
'       TabDetalheNegociacao("pedPrecoMetro") = TabDetalheNegociacao("pedPrecoUnidadePedida") / tabproduto("prdMetroCx")
'    End If
'
'    fim = 0
'    TabCarteira_Rep.MoveFirst
'    Do While fim = 0
'       If TabCarteira_Rep("chpessoa") = TabNegociacao("chrepresentante") Then
'          AjusteComissao = TabCarteira_Rep("repajustecomissao")
'          fim = 1
'       Else
'          TabCarteira_Rep.MoveNext
'          If TabCarteira_Rep.EOF Then
'             AjusteComissao = 0
'             fim = 1
'          End If
'       End If
'    Loop
'
'    TabDetalheNegociacao("pedValorDesconto") = Format$(TabDetalheNegociacao("pedPrecoUnidadePedida") * (TabDetalheNegociacao("pedDesc") / 100), "#.00")
'    'TabDetalheNegociacao("pedIPI") = (TabDetalheNegociacao("pedPrecoUnidadePedida") - TabDetalheNegociacao("pedValorDesconto")) * (tabproduto("prdIPI") / 100)
'    Comissao = (TabDetalheNegociacao("pedPrecoUnidadePedida") - TabDetalheNegociacao("pedValorDesconto")) * TabDetalheNegociacao("pedQuantidadePedida")
'
'    TabDetalheNegociacao("pedComissaoRep") = Format$(((Comissao * IndiceComissao) * ((tabproduto("prdcomissao") + AjusteComissao) - txtDescComissao) / 100), "#.00")
'
'    TabDetalheNegociacao.Update
'
'Next'

'End Sub


Public Sub Rotina_Exclui_Comissao_Rep()

'Call Rotina_AbrirBanco

'neg.Open "Select * from negociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtComplementoPedido & "')", db, 3, 3
'If neg.EOF Then
'   MsgBox ("Erro no acesso a Negociacao na rotina Exclui Comissão."), vbCritical
'   Call FechaDB
'   Exit Sub
'End If

If neg!chrepresentante = "NENHUM" Then
   'Call FechaDB
   Exit Sub
End If

Dia_Comis = 25
Mes_Comis = Month(neg!negdatanegociação)
Mes_Comis = Mes_Comis + 1
If Mes_Comis = 13 Then
   Mes_Comis = 1
   Ano_Comis = Year(Date) + 1
Else
   Ano_Comis = Year(neg!negdatanegociação)
End If
Data_Comis = Dia_Comis & "/" & Mes_Comis & "/" & Ano_Comis

ctp.Open "Select * from contas_a_pagar where chFabricante = ('" & 0 & "') and chPessoa = ('" & neg!chrepresentante & "') and chNotaFiscal = ('" & "Representante" & "') and chFatura = ('" & "Comissao" & "') and chDataVencito = ('" & Data_Comis & "')", db, 3, 3
If ctp.EOF Then
   MsgBox ("negEmissorNF = "), , neg!negEmissorNF
   MsgBox ("negEmissorNF = "), , neg!chrepresentante
   MsgBox ("negEmissorNF = "), , neg!negNotaFiscal
   MsgBox ("Representante sem registro no contas a pagar")
   Exit Sub
Else
   If ctp!chFabricante = 0 Then
      ctp!ctpValorLart = ctp!ctpValorLart - Acumula_Comis_Rep
   Else
      ctp!ctpValorMerco = ctp!ctpValorMerco - Acumula_Comis_Rep
   End If
   
   ctp!ctpValorDaBoleta = ctp!ctpValorDaBoleta - Acumula_Comis_Rep
   ctp.Update
End If

'Call FechaDB

End Sub

Public Sub Rotina_Exclui_Comissao_Promot()

'Call Rotina_AbrirBanco

'neg.Open "Select * from negociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtComplementoPedido & "')", db, 3, 3
'If neg.EOF Then
'   MsgBox ("Erro no acesso a Negociacao na rotina Exclui Comissão Promot."), vbCritical
'   Call FechaDB
'   Exit Sub
'End If

If neg!chPromotor = "NENHUM" Then
   'Call FechaDB
   Exit Sub
End If

Dia_Comis = 5
Mes_Comis = Month(neg!negdatanegociação)
Mes_Comis = Mes_Comis + 1
If Mes_Comis = 13 Then
   Mes_Comis = 1
   Ano_Comis = Year(Date) + 1
Else
   Ano_Comis = Year(neg!negdatanegociação)
End If
Data_Comis = Dia_Comis & "/" & Mes_Comis & "/" & Ano_Comis

ctp.Open "Select * from contas_a_pagar where chFabricante = ('" & 0 & "') and chPessoa = ('" & neg!chPromotor & "') and chNotaFiscal = ('" & "Promotora & " ') and chDataVencito = ('" & data_Comis & "')",db,3,3
If ctp.EOF Then
   MsgBox ("Promotora sem registro no contas a pagar"), vbInformation
   Call FechaDB
   Exit Sub
Else
   If ctp!chFabricante = 0 Then
      ctp!ctpValorLart = ctp!ctpValorLart - Acumula_Comis_Promot
   Else
      ctp!ctpValorMerco = ctp!ctpValorMerco - Acumula_Comis_Promot
   End If

   ctp!ctpValorDaBoleta = ctp!ctpValorDaBoleta - Acumula_Comis_Promot
   ctp.Update
End If
 
'Call FechaDB
 
End Sub

Public Sub Rotina_Exclui_Cta_Receber()

'Call Rotina_AbrirBanco

'neg.Open "Select * from negociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ( '" & txtComplementoPedido & "')", db, 3, 3
'If neg.EOF Then
'   MsgBox ("Erro no acesso a negociação em Exclui Contas a Receber."), vbCritical
'   Call FechaDB
'   Exit Sub
'End If

ctr.Open "Select * from contas_a_receber where chNotaFiscal = ('" & neg!negNotaFiscal & "')", db, 3, 3
If ctr.EOF Then
   'MsgBox ("Contas a receber não contém registros."), vbInformation
   'Call FechaDB
   Exit Sub
End If


'ctr.MoveFirst

Do While Not ctr.EOF

If ctr!chNotafiscal = neg!negNotaFiscal Then
   ctr.Delete
   ctr.MoveNext
Else
   ctr.MoveNext
End If
Loop
End Sub

Public Sub Rotina_Exclui_Cta_Pagar()

'Call Rotina_AbrirBanco

'neg.Open "Select * from negociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ( '" & txtComplementoPedido & "')", db, 3, 3
'If neg.EOF Then
'   MsgBox ("Erro no acesso a negociação em Exclui Contas a Receber."), vbCritical
'   Call FechaDB
'   Exit Sub
'End If

ctp.Open "Select * from contas_a_pagar where chNotaFiscal = ('" & neg!negNotaFiscal & "')", db, 3, 3
If ctp.EOF Then
   'MsgBox ("Contas a Pagar não contém registros."), vbInformation
   'Call FechaDB
   Exit Sub
End If


ctp.MoveFirst

Do While Not ctp.EOF

If ctp!chNotafiscal = neg!negNotaFiscal Then
   ctp.Delete
   ctp.MoveNext
Else
   ctp.MoveNext
End If
Loop

'Call FechaDB

End Sub

Public Sub Rotina_Bloqueia_Campos()
txtNumPedido.Enabled = False
txtComplementoPedido.Enabled = False
txtDataPedido.Enabled = False
cmbPessoa.Enabled = False
txtEndereco.Enabled = False
txtBairro.Enabled = False
txtCidade.Enabled = False
txtUF.Enabled = False
'txtCEP.Enabled = False
'txtTel.Enabled = False
txtQtdFat.Enabled = False
txtIntervalo.Enabled = False
TxtAPartirDe.Enabled = False
'txtPrzBoletaFrete.Enabled = False

'txtPreçoFixo.Enabled = False

'cmbCobrancaFrete.Enabled = False
'cmbCondProcessamento.Enabled = False
txtUnidade.Enabled = False

'txtDescComissao.Enabled = False
cmbProduto.Enabled = False
txtNomeProduto.Enabled = False

txtQtd.Enabled = False
txtPreçoUnit.Enabled = False

'txtDesc.Enabled = False

'txtTotalMetros.Enabled = False
'txtTotalCaixa.Enabled = False
'txtTotalPeso.Enabled = False
txtAcumula_Produto.Enabled = False
txtAcumula_Desconto.Enabled = False
txtValorComDesconto.Enabled = False
'txtAcumula_Frete.Enabled = False
'txtValor_Operacao.Enabled = False
'txtAcumula_IPI.Enabled = False
'txtPreçoFixo.Enabled = False
'cmbEmissor.Enabled = False
cmbCFOP.Enabled = False
'cmbOrdemDeCarga.Enabled = False
'lblTransporte.Enabled = False
txtNotaFiscal.Enabled = False

cmbBanco.Enabled = False
'cmbMotivacao.Enabled = False
txtRepresentante.Enabled = False
txtPromotora.Enabled = False

End Sub


Public Sub Rotina_Desbloqueia_Campos()
txtNumPedido.Enabled = True
txtComplementoPedido.Enabled = True
txtDataPedido.Enabled = True
'txtDataPedido = "__/__/____"
cmbPessoa.Enabled = True
txtEndereco.Enabled = True
txtBairro.Enabled = True
txtCidade.Enabled = True
txtUF.Enabled = True
'txtCEP.Enabled = True
'txtTel.Enabled = True
txtQtdFat.Enabled = True
txtIntervalo.Enabled = True
TxtAPartirDe.Enabled = True
'txtPrzBoletaFrete.Enabled = True

'txtPreçoFixo.Enabled = True

'txtDescComissao.Enabled = True
cmbProduto.Enabled = True
txtNomeProduto.Enabled = True

txtQtd.Enabled = True
txtPreçoUnit.Enabled = True
'cmbCobrancaFrete.Enabled = True
'cmbCondProcessamento.Enabled = True
txtUnidade.Enabled = True
'txtDesc.Enabled = True
'
'txtTotalMetros.Enabled = True
''txtTotalCaixa.Enabled = True
'txtTotalPeso.Enabled = True
txtAcumula_Produto.Enabled = True
txtAcumula_Desconto.Enabled = True
txtValorComDesconto.Enabled = True
'txtAcumula_Frete.Enabled = True
'txtValor_Operacao.Enabled = True
'txtAcumula_IPI.Enabled = True
'txtPreçoFixo.Enabled = True
'cmbEmissor.Enabled = True
cmbCFOP.Enabled = True
'cmbOrdemDeCarga.Enabled = True
'lblTransporte.Enabled = True
txtNotaFiscal.Enabled = True

cmbBanco.Enabled = True
'cmbMotivacao.Enabled = True
txtRepresentante.Enabled = True
txtPromotora.Enabled = True

End Sub

Public Sub Rotina_Criticar_Campos()
If txtDataPedido = Empty Then
   MsgBox ("Data do pedido não informada")
   txtDataPedido.SetFocus
   Erro_Critica = 1
   Exit Sub
End If

If txtDataPedido = "__/__/____" Then
   MsgBox ("Data do pedido não infornada")
   txtDataPedido.SetFocus
   Erro_Critica = 1
   Exit Sub
End If

If Not (IsDate(txtDataPedido)) Then
   MsgBox ("Data do pedido inválida")
   txtDataPedido.SetFocus
   Erro_Critica = 1
   Exit Sub
End If

Dia_Pedido = Day(txtDataPedido)
Mes_Pedido = Month(txtDataPedido)
Ano_Pedido = Year(txtDataPedido)

If IsNull(Dia_Pedido) Then
   MsgBox ("Data (Dia) do pedido inválida")
   txtDataPedido.SetFocus
   Erro_Critica = 1
   Exit Sub
End If
If Dia_Pedido < 1 Or Dia_Pedido > 31 Then
   MsgBox ("Data do pedido inválida")
   txtDataPedido.SetFocus
   Erro_Critica = 1
   Exit Sub
End If
If IsNull(Mes_Pedido) Then
   MsgBox ("Data (Mes) do pedido inválida")
   txtDataPedido.SetFocus
   Erro_Critica = 1
   Exit Sub
End If
If Mes_Pedido < 1 Or Mes_Pedido > 12 Then
   MsgBox ("Data do pedido inválida")
   txtDataPedido.SetFocus
   Erro_Critica = 1
   Exit Sub
End If
If IsNull(Ano_Pedido) Then
   MsgBox ("Data (ano) do pedido inválida")
   txtDataPedido.SetFocus
   Erro_Critica = 1
   Exit Sub
End If
If Ano_Pedido < Year(Date) Then
   Resp = MsgBox("Data do pedido anterior a data de hoje. Aceitar? Sim ou Não???", vbYesNo)
   If Resp = vbNo Then
      Erro_Critica = 1
      txtDataPedido.SetFocus
      Erro_Critica = 1
      Exit Sub
   End If
End If
If DataConv = txtDataPedido Then
   If DataConv > Date Then
      Resp = MsgBox("Data do pedido posterior a data de hoje. Aceitar? Sim ou Não???", vbYesNo)
      If Resp = vbNo Then
         Erro_Critica = 1
         txtDataPedido.SetFocus
         Erro_Critica = 1
         Exit Sub
      End If
   End If
End If

If txtQtdFat = "" Then
   MsgBox "Quantidade de faturas não informada"
   txtDataPedido.SetFocus
   Erro_Critica = 1
   Exit Sub
End If
If txtIntervalo = "" Then
   MsgBox "Intervalo da fatura não informada"
   txtDataPedido.SetFocus
   Erro_Critica = 1
   Exit Sub
End If
If TxtAPartirDe = "" Then
   MsgBox "Data da fatura a partir de não informada"
   TxtAPartirDe.SetFocus
   Erro_Critica = 1
   Exit Sub
End If
If cmbCFOP = "" Then
   MsgBox "CFOP informado"
   cmbCFOP.SetFocus
   Erro_Critica = 1
   Exit Sub
End If
If Altera = 1 Then
   Altera = 0
Else
   If cmbProduto = Empty Then
      MsgBox "Produto não Informado."
      cmbProduto.SetFocus
      Erro_Critica = 1
      Exit Sub
   End If

   If txtUnidade = Empty Then
      MsgBox ("Unidade não Informada")
      txtUnidade.SetFocus
      Erro_Critica = 1
      Exit Sub
   Else
      If txtQtd = Empty Then
         MsgBox ("Quantidade não Informada")
         txtQtd.SetFocus
         Erro_Critica = 1
         Exit Sub
      Else
         If txtPUCheio = Empty Then
            MsgBox ("Preço não Informado")
            txtPUCheio.SetFocus
            Erro_Critica = 1
            Exit Sub
         Else
            'If txtDesc = Empty Then
            '   txtDesc = Format$(0, "#0.00")
            'Else
            '   txtDesc = Format$(txtDesc, "#0.00")
            'End If
            cmdIncluiDetalhe.SetFocus
         End If
      End If
   End If
End If
End Sub

'Public Sub Rotina_Exclui_Frete()
'If IsNull(TabNegociacao("chordemdecarga")) Or TabNegociacao("chordemdecarga") = "Cliente" Then
'   Exit Sub
'End If

'TabPagtosEmCheque.Seek "=", cmbOrdemDeCarga, "Maxalbido"

'If TabPagtosEmCheque.NoMatch Then
'   Exit Sub
'End If

'MsgBox ("Os cheques cadastrados para esta ordem de carga serão deletados. Devem ser recadastrados")'
'
'If TabDetPgCheque.RecordCount = 0 Then
'   Exit Sub
'End If
'
'TabDetPgCheque.MoveFirst
'
'Do While Not TabDetPgCheque.EOF'
'
'   If TabDetPgCheque("chordemdecarga") = cmbOrdemDeCarga And TabDetPgCheque("chemissor") = "Maxalbido" Then
'      TabCtaPagar.Seek "=", 0, TabPagtosEmCheque("ocgmotorista"), TabDetPgCheque("chnumdoc"), TabDetPgCheque("chnumdoc"), TabDetPgCheque("docgDataCompensacao")
'      If TabCtaPagar.NoMatch Then
'         TabDetPgCheque.Delete
'      Else
'         If TabCtaPagar("ctpstatus") = 0 Then
'            TabCtaPagar.Delete
'            TabDetPgCheque.Delete
'            TabPagtosEmCheque.Seek "=", cmbOrdemDeCarga, "Maxalbido"
'            If TabPagtosEmCheque.NoMatch Then
'               MsgBox ("Não encontrei. O que fazer???")
'            End If
'            TabPagtosEmCheque.Edit
 '           TabPagtosEmCheque("ocgstatus") = 0
'            TabPagtosEmCheque.Update
 '        Else
 '           MsgBox ("Antes de cancelar esta ordem de carga, voce devera cancelar o pagamento do cheque")
'            Exit Sub
'         End If
'      End If
'   End If
'   TabDetPgCheque.MoveNext

'Loop

'TabDetalhePessoaFrete.Seek "=", cmbOrdemDeCarga, txtNotaFiscal
'If TabDetalhePessoaFrete.NoMatch Then
'   MsgBox ("Não há nota fiscal para exclusão")
'Else
'   ValorFrete = TabDetalhePessoaFrete("dfpvalor")
'   TabDetalhePessoaFrete.Delete
'End If'

'If TabPagtosEmCheque("ocgvalortotal") < ValorFrete Then
'   MsgBox ("Frete da tela maior que o frete do arquivo. Vou assumir igual")
'   txtAcumula_Frete = TabPagtosEmCheque("ocgvalortotal")
'End If
'
'TabPagtosEmCheque.Edit

'TabPagtosEmCheque("ocgvalorpedagio") = 0
'TabPagtosEmCheque("ocgvalorfrete") = TabPagtosEmCheque("ocgvalortotal") - ValorFrete
'TabPagtosEmCheque("ocgvalortotal") = TabPagtosEmCheque("ocgvalortotal") - ValorFrete
'If TabPagtosEmCheque("ocgvalortotal") = 0 Then
 '  TabPagtosEmCheque.Delete
'El'se
'   TabPagtosEmCheque.Update
'End If

'End Sub

'Public Sub Carrega_Ordem_De_Carga()

'cmbOrdemDeCarga.Clear

'cmbOrdemDeCarga.AddItem "Cliente"

'If TabPagtosEmCheque.RecordCount > 0 Then
'   TabPagtosEmCheque.MoveFirst
'End If
'Do While Not TabPagtosEmCheque.EOF
'   If TabPagtosEmCheque("ocgdatadacarga") = Data_Hoje Then
'      If TabPagtosEmCheque("chemissor") = cmbEmissor Then
'         cmbOrdemDeCarga.AddItem TabPagtosEmCheque("chordemdecarga")
'      End If
'   End If
'   TabPagtosEmCheque.MoveNext
'
'Loop
'End Sub

'Private Sub Rotina_Calcula_Preco_composto()
'Acesso para recuperar a chave de preco dos componentes

'ContaComponente = 0
'fim = 0
'PrecoProdutoComposto = 0

'TabComposicaoProdutoFinal.Seek "=", tabproduto("prdgrupocomppreco"), 0
'If TabComposicaoProdutoFinal.NoMatch Then
 '  MsgBox ("Componente não cadastrado")
'   cmdSair.SetFocus
'Else
'   Do While fim = 0
'      TabComposicaoProdutoFinal.MoveNext
'      If TabComposicaoProdutoFinal.EOF Then
'         fim = 1
'         If ContaComponente = 0 Then
'            MsgBox ("Grupo cadastrado sem componentes ")
'         End If
'      Else
'         If TabComposicaoProdutoFinal("chchavegrupo") = tabproduto("prdgrupocomppreco") Then
'            ContaComponente = ContaComponente + 1
'            tabProdutoPreco.Seek "=", Tabpessoa("chpessoa"), TabComposicaoProdutoFinal("chcomponente"), 0
'            If tabProdutoPreco.NoMatch Then
'               tabProdutoPreco.Seek "=", "GERAL", TabComposicaoProdutoFinal("CHCOMPONENTE"), 0
'               If tabProdutoPreco.NoMatch Then
'                  MsgBox ("Preco de componente " & TabComposicaoProdutoFinal("chcomponente") & " não cadastrado")
 '                 fim = 1
 '                 Exit Sub
  '             Else
  '                Call Calcula_Preco_Composto
  '             End If
  '          Else
  '             Calcula_Preco_Composto
  '          End If
  '       End If
'      End If
         
'   Loop
'End If
   
'End Sub

'Private Sub Calcula_Preco_Composto()
'tabComponentesProdutos.Seek "=", tabproduto("chproduto"), TabComposicaoProdutoFinal("chcomponente")
'If tabComponentesProdutos.NoMatch Then
'   MsgBox ("Produto " & TabComposicaoProdutoFinal("chcomponente") & " sem componentes.")
'   Exit Sub
'Else
'   PrecoProdutoComposto = PrecoProdutoComposto + (tabComponentesProdutos("qtdcomponente") * tabProdutoPreco("pdpprecodoproduto"))
'   'PrecoProdutoComposto = PrecoProdutoComposto + (tabComponentesProdutos("qtdcomponente") * tabProdutoPreco("pdpprecodoproduto"))
'End If
'End Sub

Private Sub CargaProduto()
Dim SalvaIndice As Integer
Dim FimCargaProduto As Byte
SalvaIndice = 0
FimCargaProduto = 0

cmbProduto.Clear

Call Rotina_AbrirBanco

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If

'O primeiro acesso é para pegar os equipamentos

Prod.Open "Select * from produto where prdLocadora = ('" & cmbPessoa & "') and prdUnidadeOperacional = ('" & cmbLocal & "') and prdOrdemApresentacao = ('" & 0 & "')", db, 3, 3
If Not Prod.EOF Then
   Prod.MoveFirst
   Do While Not Prod.EOF
      cmbProduto.AddItem Prod!chProduto
      Prod.MoveNext
   Loop
   cmbProduto.AddItem "MOB/DESMOB"
End If

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If

'Esse segundo acesso é para pegar pessoal

'Prod.Open "Select * from produto where prdLocadora = ('" & cmbPessoa & "') and prdUnidadeOperacional = ('" & Contrato & "') and prdOrdemApresentacao = ('" & 0 & "')", db, 3, 3
'If Not Prod.EOF Then
'   If pes.State = 1 Then
'      pes.Close: Set pes = Nothing
'   End If
      
   pes.Open "Select * from pessoa where pesClienteLocador = ('" & cmbPessoa & "') and pesUnidadeOperacional = ('" & cmbLocal & "')", db, 3, 3
   If pes.EOF Then
      MsgBox ("Cliente sem mão de obra cadastrada"), vbInformation
   Else
      Do While Not pes.EOF
         If pes!pesStatusPessoa = 0 Then
            cmbProduto.AddItem pes!chPessoa
         End If
         pes.MoveNext
      Loop
   End If
'End If

lblLocalizacao = cmbLocal

Call FechaDB

End Sub

Public Sub CalculaValorHora()

ValorHora = ProdPco!pdpPrecoDoProduto / 8
ValorMinuto = ProdPco!pdpPrecoDoProduto / 480

End Sub

Public Sub CalculaMinutos()
If Not txtQtd = Empty Then
    QtdHoras = Int(txtQtd)
    QtdMinutos = (txtQtd - QtdHoras) * 100
    MinutosParaCalculo = (QtdHoras * 60) + QtdMinutos
End If
End Sub

Public Sub ResumoMedicao()

Dim TotalMedicao As Currency
Dim TotalGeral As Currency
Dim IndMedicao As Integer

'If txtStatusPedido = "PROCESSADO" Then
   gridMedicao.Rows = 2
   Linha = 1
   gridMedicao.TextMatrix(Linha, 0) = Empty
   gridMedicao.TextMatrix(Linha, 1) = Empty
   gridMedicao.TextMatrix(Linha, 2) = Empty
   gridMedicao.TextMatrix(Linha, 3) = Empty
   txtTotalMedicao = Empty
'   Exit Sub
'End If
   
Call Rotina_AbrirBanco
TotalMedicao = 0
TotalGeral = 0

neg.Open "Select * from negociacao where chNumPedido = ('" & txtNumPedido & "')", db, 3, 3

If neg.EOF Then
   Call FechaDB
   Exit Sub
End If

IndMedicao = 0


neg.MoveFirst

Do While Not neg.EOF
   
   If dneg.State = 1 Then
      dneg.Close: Set dneg = Nothing
      acdNeg = 0
   End If

   dneg.Open "Select * from detalhenegociacao where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
   If dneg.EOF Then
      Call FechaDB
      Exit Sub
   End If
   
   dneg.MoveFirst
   TotalMedicao = 0
   Do While Not dneg.EOF
      TotalMedicao = TotalMedicao + dneg!pedValorDaOperacao
      dneg.MoveNext
   Loop
   
   IndMedicao = IndMedicao + 1
   gridMedicao.Rows = IndMedicao + 1
   If neg!chUnidadeOperacional = Empty Then
      gridMedicao.TextMatrix(IndMedicao, 0) = Empty
   Else
      gridMedicao.TextMatrix(IndMedicao, 0) = neg!chUnidadeOperacional
   End If
   gridMedicao.TextMatrix(IndMedicao, 1) = neg!chNumPedido
   gridMedicao.TextMatrix(IndMedicao, 2) = neg!chNumPedidoComp
   gridMedicao.TextMatrix(IndMedicao, 3) = Format$(TotalMedicao, "#,##0.00")
   TotalGeral = TotalGeral + TotalMedicao
   TotalMedicao = 0
   
   neg.MoveNext
Loop

txtTotalMedicao = Format(TotalGeral, "#,##0.00")

End Sub


Public Sub LimpaGridMedicao()
gridMedicao.Rows = 2
Linha = 1
    gridMedicao.TextMatrix(Linha, 0) = Empty
    gridMedicao.TextMatrix(Linha, 1) = Empty
    gridMedicao.TextMatrix(Linha, 2) = Empty
    gridMedicao.TextMatrix(Linha, 3) = Empty
    txtTotalMedicao = Empty
End Sub

