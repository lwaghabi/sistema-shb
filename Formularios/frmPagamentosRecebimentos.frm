VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPagamentosRecebimentos 
   Caption         =   "frmPagamentosRecebimentos"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Parâmetros de Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   11175
      Begin VB.ComboBox cmbClienteColaborador 
         Height          =   315
         Left            =   5880
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   1920
         Width           =   5175
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3855
         Begin VB.OptionButton Option1 
            Caption         =   "Fornecedor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            TabIndex        =   13
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Despesas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2040
            TabIndex        =   12
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   1335
         Left            =   3960
         TabIndex        =   4
         Top             =   240
         Width           =   7095
         Begin VB.OptionButton Option3 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Caption         =   "De/Até"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1440
            TabIndex        =   7
            Top             =   480
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtAte 
            Height          =   375
            Left            =   5160
            TabIndex        =   5
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   118095873
            CurrentDate     =   44527
         End
         Begin MSComCtl2.DTPicker dtDe 
            Height          =   375
            Left            =   3120
            TabIndex        =   6
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   118095873
            CurrentDate     =   44527
         End
         Begin VB.Label lblDe 
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
            Height          =   255
            Left            =   3120
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblAte 
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
            Height          =   255
            Left            =   5160
            TabIndex        =   9
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Label lblCliFornec 
         Caption         =   "lblCliFornec"
         Height          =   255
         Left            =   5880
         TabIndex        =   15
         Top             =   1680
         Width           =   2415
      End
   End
   Begin VB.TextBox dtHoje 
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
      Height          =   375
      Left            =   9360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   1695
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
      Left            =   9360
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Consulta a Valores Pagos e Recebidos"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmPagamentosRecebimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
