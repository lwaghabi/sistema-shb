VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmExtratoNotaFiscal 
   Caption         =   "Extrato de Nota Fiscal- (frmExtratoNotaFiscal)"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   16770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   12960
      TabIndex        =   77
      Top             =   7920
      Width           =   3735
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H000000FF&
         Caption         =   "Sair"
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
         Left            =   1080
         MaskColor       =   &H00FFFF00&
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   3135
      Left            =   13080
      TabIndex        =   64
      Top             =   2880
      Width           =   3615
      Begin VB.TextBox txtResTotalIPI 
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
         Height          =   405
         Left            =   1920
         TabIndex        =   70
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtResTotalFrete 
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
         Height          =   405
         Left            =   1920
         TabIndex        =   69
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtResTotalProduto 
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
         Height          =   405
         Left            =   1920
         TabIndex        =   68
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtResTotalICMS 
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
         Height          =   405
         Left            =   1920
         TabIndex        =   67
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtResTotalDesconto 
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
         Height          =   405
         Left            =   1920
         TabIndex        =   66
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtResValorTotalNF 
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
         Height          =   405
         Left            =   1920
         TabIndex        =   65
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "ICMS"
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
         TabIndex        =   76
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "IPI"
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
         TabIndex        =   75
         Top             =   1800
         Width           =   285
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Valor do Produto"
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
         TabIndex        =   74
         Top             =   1320
         Width           =   1770
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Desconto"
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
         TabIndex        =   73
         Top             =   2280
         Width           =   1005
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Frete"
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
         TabIndex        =   72
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Valor da Nota"
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
         TabIndex        =   71
         Top             =   2760
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Negociação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   56
      Top             =   0
      Width           =   7575
      Begin VB.ComboBox txtNotaFiscal 
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
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   435
         Left            =   240
         TabIndex        =   5
         Top             =   2250
         Width           =   7095
      End
      Begin VB.TextBox txtCompPedido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4320
         TabIndex        =   3
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtNumPedido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   3855
      End
      Begin VB.TextBox txtEmissorNF 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4800
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin MSMask.MaskEdBox txtDataPedido 
         Height          =   405
         Left            =   5760
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   714
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDataEmissao 
         Height          =   405
         Left            =   2040
         TabIndex        =   0
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   714
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Left            =   360
         TabIndex        =   63
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data Medição"
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
         Left            =   5760
         TabIndex        =   62
         Top             =   1200
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Comp"
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
         Left            =   4320
         TabIndex        =   61
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Medição"
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
         Left            =   240
         TabIndex        =   60
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Emissor"
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
         Left            =   4800
         TabIndex        =   59
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Data Emissão"
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
         Left            =   2280
         TabIndex        =   58
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nota Fiscal"
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
         Left            =   240
         TabIndex        =   57
         Top             =   360
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalhes da Negociação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   7920
      TabIndex        =   39
      Top             =   0
      Width           =   5055
      Begin VB.Frame Frame6 
         Caption         =   "Faturamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   4815
         Begin VB.TextBox txtIntervalo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2160
            TabIndex        =   48
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtFaturamento 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1320
            TabIndex        =   47
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtBancoFat 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   45
            Top             =   600
            Width           =   975
         End
         Begin MSMask.MaskEdBox txtDataPrimParc 
            Height          =   405
            Left            =   3000
            TabIndex        =   46
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Interv."
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
            Left            =   2160
            TabIndex        =   52
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Fatur."
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
            Left            =   1320
            TabIndex        =   51
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Prim. parcela"
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
            Left            =   3000
            TabIndex        =   50
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
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
            Height          =   300
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   780
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frete - Condição e Data"
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
         Left            =   120
         TabIndex        =   41
         Top             =   1920
         Width           =   4815
         Begin VB.TextBox txtCondFrete 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   2415
         End
         Begin MSMask.MaskEdBox txtDataFrete 
            Height          =   405
            Left            =   2880
            TabIndex        =   42
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
      Begin VB.TextBox txtAliquotaICMS 
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
         Height          =   405
         Left            =   3720
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtCondProcess 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtBancoFatN 
         Height          =   285
         Left            =   4440
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtPercDescComis 
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
         Height          =   405
         Left            =   2760
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cond. de Proces."
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
         TabIndex        =   55
         Top             =   240
         Width           =   2070
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "  ICMS"
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
         Left            =   3600
         TabIndex        =   54
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "%Desc"
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
         Left            =   2760
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame8 
      Height          =   2895
      Left            =   13080
      TabIndex        =   24
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtMotorista 
         Height          =   405
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtTransportadora 
         Height          =   405
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtIdentidade 
         Height          =   405
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtMunicipioPlaca 
         Height          =   405
         Left            =   120
         TabIndex        =   28
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox txtUFPlaca 
         Height          =   405
         Left            =   2760
         TabIndex        =   27
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtPlaca 
         Height          =   405
         Left            =   2520
         TabIndex        =   26
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtOrdemDeCarga 
         Height          =   405
         Left            =   2280
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ/CPF"
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
         TabIndex        =   38
         Top             =   1320
         Width           =   1260
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Motorista"
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
         TabIndex        =   37
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Transportadora"
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
         TabIndex        =   36
         Top             =   120
         Width           =   1860
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Placa"
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
         Left            =   2520
         TabIndex        =   35
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Município"
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
         TabIndex        =   34
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "UF"
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
         Left            =   2760
         TabIndex        =   33
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "O. de Carga"
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
         Left            =   2160
         TabIndex        =   32
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comissão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   13080
      TabIndex        =   17
      Top             =   6000
      Width           =   3615
      Begin VB.TextBox txtRepresentante 
         Height          =   405
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtComissaoRep 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   1680
         TabIndex        =   20
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtPromotora 
         Height          =   405
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtComissaoPromot 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   1680
         TabIndex        =   18
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Left            =   840
         TabIndex        =   79
         Top             =   1560
         Width           =   645
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
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
         Left            =   2100
         TabIndex        =   23
         Top             =   120
         Width           =   1260
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Promotor"
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
         TabIndex        =   22
         Top             =   840
         Width           =   1110
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Distribuição do Faturamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   12
      Top             =   6000
      Width           =   12735
      Begin MSFlexGridLib.MSFlexGrid GridFatura 
         Height          =   2175
         Left            =   120
         TabIndex        =   80
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColor       =   16777152
         BackColorFixed  =   16776960
         BackColorBkg    =   16777152
         FormatString    =   "Operação                   |Nota Fiscal |Ordem|Data Vencito|Valor            |Frete       | Valor da Boleta    "
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
      Begin VB.TextBox txtTotalFaturaLart 
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
         Height          =   360
         Left            =   6960
         TabIndex        =   15
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtTotalFaturaMerco 
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
         Height          =   360
         Left            =   8640
         TabIndex        =   14
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtTotalFatura 
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
         Height          =   360
         Left            =   9960
         TabIndex        =   13
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Totais"
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
         Left            =   5880
         TabIndex        =   16
         Top             =   2400
         Width           =   750
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Negociação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   12735
      Begin MSFlexGridLib.MSFlexGrid GridNegocio 
         Height          =   2535
         Left            =   120
         TabIndex        =   81
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColor       =   16777152
         BackColorFixed  =   16776960
         BackColorBkg    =   16777152
         FormatString    =   "Nome Produto                                   |Unid |Qtd. |P.U.           |Valor Diária |Qtd Dias|ValorOperação"
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
      Begin VB.TextBox txtValorTotal 
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
         Height          =   405
         Left            =   10320
         TabIndex        =   10
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
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
         Height          =   300
         Left            =   9360
         TabIndex        =   11
         Top             =   2880
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmExtratoNotaFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Guarda_Cliente As String
Dim Guarda_Representante As String
Dim Guarda_Promotor As String
Dim Guarda_Motorista As String
Dim Fazer As Byte
Dim EmissorNF As Byte
Dim DataString As Date
Dim Encontrei As Byte
Dim Base As Byte
Dim NumPedidoAux As String
Dim NumPedidoCompAux As String

Dim IndFatSalvo As Integer
Dim IndNegSalvo As Integer

Dim Resp As String

Dim A As Integer
Dim fim As Byte
Dim Fim_Carga As Byte
Dim Linha As Integer
Dim Unidade(3) As String
Dim IndFabricante As Byte
Dim Tipo_Comissao As Byte '1=Representante; 2=Promotor

Dim ResTotalICMS As Currency

Dim Valor_Total As Currency
Dim Valor_Produto As Currency
Dim Valor_Frete As Currency
Dim Valor_IPI As Currency
Dim Valor_Desconto As Currency

Dim Acumula_Valor_Total As Currency
Dim Acumula_Valor_Produto As Currency
Dim Acumula_Valor_Frete As Currency
Dim Acumula_Valor_IPI As Currency
Dim Acumula_Valor_Desconto As Currency
Dim Acumula_Quantidade As Integer
Dim Acumula_Comis_Rep(2) As Currency
Dim Acumula_Comis_Promot(2) As Currency

Dim Acumula_Total_Fabricante(2) As Currency
Dim Acumula_Produto_Fabricante(2) As Currency
Dim Acumula_Frete_Fabricante(2) As Currency
Dim Acumula_IPI_Fabricante(2) As Currency
Dim Acumula_Desconto_Fabricante(2) As Currency
Dim Acumula_Quantidade_Fabricante(2) As Integer

Dim Acumula_Comissao_Rep As Currency
Dim Acumula_Comissao_Promot As Currency

Dim AcumValorConsig(3) As Currency
Dim ProdutoConsig(3) As String

Dim Valor_Fatura(10) As Currency
Dim Frete_Fatura(10) As Currency
Dim Fatura_Lart(10) As Currency
Dim Fatura_Merco(10) As Currency
Dim Frete_Lart(10) As Currency
Dim Frete_Merco(10) As Currency
Dim Mneu_Produto(100) As String

Dim Lart_Frete As Currency
Dim Merco_Frete As Currency

Dim Base_Fatura As Currency
Dim Base_Frete As Currency

Dim Ind_Inicial_Fatura As Byte
Dim Ind_Inicial_Frete As Byte
Dim Ind_Aux As Byte

Dim Base_Fatura_Lart As Currency
Dim Base_Fatura_Merco As Currency

Dim Parcelas_Frete As Byte
Dim Intervalo As Byte
Dim Data_Cobranca As Date

Dim Acumula_Fatura_Lart As Currency
Dim Acumula_Fatura_Merco As Currency
Dim Acumula_Fatura_Geral As Currency

Dim Data_Comissao As Date

Dim Ano_Comissao As Integer
Dim Mes_Comissao As Integer
Dim Dia_Comis_Promot As Integer
Dim Dia_Comis_Repres As Integer

Dim Data_Proc As Date

Private Sub cmdSair_Click()
Unload Me
End Sub



Private Sub Form_Load()

fim = 0

Call Rotina_AbrirBanco

neg.Open "Select * from Negociacao where negStatus = ('" & 1 & "')", db, 3, 3
If neg.EOF Then
   fim = 1
Else
   neg.MoveFirst
End If

Do While fim = 0
   txtNotaFiscal.AddItem neg!negNotaFiscal
   neg.MoveNext
   If neg.EOF Then
      fim = 1
   End If
Loop

fim = 0

hneg.Open "Select * from HistoricoNegociacao where hngStatus = ('" & 1 & "')", db, 3, 3
   If hneg.EOF Then
      fim = 1
   Else
      hneg.MoveFirst
   End If
   
Do While fim = 0
   txtNotaFiscal.AddItem hneg!hngnotafiscal
   hneg.MoveNext
   If hneg.EOF Then
      fim = 1
   End If
Loop


End Sub

Private Sub txtNotaFiscal_LostFocus()

If txtNotaFiscal = "" Then
   MsgBox "Nota Fiscal Não Informada. Rotina Será Descontinuada."
   cmdSair.SetFocus
   Exit Sub
End If

Unidade(0) = Empty
Unidade(1) = "M"
Unidade(2) = "Cx"

ProdutoConsig(1) = "CSGR"
ProdutoConsig(2) = "CSGP"

Call Rotina_AbrirBanco

neg.Open "Select * from Negociacao where negNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3
If neg.EOF Then
   hneg.Open "Select * from HistoricoNegociacao where hngNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3
   If hneg.EOF Then
      MsgBox ("Nota Fiscal Não Encontrada"), vbCritical
      Call FechaDB
      Exit Sub
   Else
      NumPedidoAux = hneg!chNumpedido
      NumPedidoCompAux = hneg!chNumPedidoComp
      Rotina_Nota_Fiscal_Historico
   End If
Else
   NumPedidoAux = neg!chNumpedido
   NumPedidoCompAux = neg!chNumPedidoComp
   Rotina_Nota_Fiscal_Atual
End If

Call FechaDB

End Sub

Public Sub Rotina_Nota_Fiscal_Atual()

'DataString = nfe!negdatanegociação

Data_Proc = DataString

Ano = Year(Data_Proc)
Mes = Month(Data_Proc)
Dia = Day(Data_Proc)

Ano_Comissao = Ano
Mes_Comissao = Mes + 1

Dia_Comis_Repres = 25
Dia_Comis_Promot = 5

If Mes_Comissao > 12 Then
   Mes_Comissao = 1
   Ano_Comissao = Ano_Comissao + 1
End If

Data_Comissao = 1 & "/" & Mes_Comissao & "/" & Ano_Comissao

Call Rotina_Limpa_Form

For A = 0 To 9
    Valor_Fatura(A) = 0
    Fatura_Lart(A) = 0
    Fatura_Merco(A) = 0
    Frete_Lart(A) = 0
    Frete_Merco(A) = 0
    Frete_Fatura(A) = 0
Next

For A = 0 To 99
    Mneu_Produto(A) = Empty
Next

Acumula_Fatura_Lart = 0
Acumula_Fatura_Merco = 0
Acumula_Fatura_Geral = 0
Acumula_Comis_Rep(0) = 0
Acumula_Comis_Rep(1) = 0
Acumula_Comis_Promot(0) = 0
Acumula_Comis_Promot(1) = 0
Acumula_Comissao_Rep = 0
Acumula_Comissao_Promot = 0

Call Rotina_AbrirBanco

neg.Open "Select * from Negociacao where chNumPedido = ('" & NumPedidoAux & "') and chNumPedidoComp = ('" & NumPedidoCompAux & "')", db, 3, 3
If neg.EOF Then
   MsgBox ("Número da Medição Inválido"), vbCritical
   Call FechaDB
   Exit Sub
End If


txtDataEmissao = neg!negdatanegociação
txtEmissorNF = "SHB Brasil"

Bco.Open "Select * from Banco where bcoCodBcoLart = ('" & neg!chcodbcolart & "')", db, 3, 3
If Bco.EOF Then
   MsgBox ("Numero do Banco inválido. "), vbCritical
   Call FechaDB
   Exit Sub
End If

txtCliente = neg!chPessoa
txtDataPedido = neg!negDataPedido
txtNumPedido = neg!chNumpedido
txtCompPedido = neg!chNumPedidoComp
txtNotaFiscal = neg!negNotaFiscal
txtBancoFat = Bco!bcosiglabco
txtBancoFatN = Bco!bcoBodBcoFEBRABAN
txtOrdemDeCarga = neg!chordemdecarga
txtEmissorNF = "SHB Brasil"

'TabCobrancaFrete.Seek "=", neg!negCobrancaFrete

'If TabCobrancaFrete.NoMatch Then
'   MsgBox ("Erro na Leitura do Parametro Cobranca Frete")
''   A = 1 / 0
'Else
'   txtCondFrete = TabCobrancaFrete("parDescCobrancaFrete")
'   If TabNegociacao("negCobrancaFrete") = 1 Then
'      txtDataFrete = Data_Proc + TabNegociacao("negBoletaFrete")
''   Else
'      txtDataFrete = "__/__/____"
'   End If
'End If
CondProc.Open "Select * from CondProcessamento where chCondicaoProcessamento = ('" & neg!negCondProcess & "')", db, 3, 3
If CondProc.EOF Then
   MsgBox ("Erro na Leitura do Parâmetro Condição de Processamento"), vbCritical
   Call FechaDB
   Exit Sub
Else
   txtCondProcess = CondProc!cprDescCondProcess
End If

txtPercDescComis = Format$(neg!negdesccomissao, "#0.00") & "%"
 
pes.Open "Select * from Pessoa where chPessoa = ('" & neg!chPessoa & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Erro no acesso a Pessoa - Extrato Nota Fiscal"), vbCritical
   Call FechaDB
   Exit Sub
End If

'ICM.Open "Select * from ICMS where chUF = ('" & pes!chUF & "')", db, 3, 3
'If ICM.EOF Then
'   MsgBox ("Estado não cadastrado na Tabela de ICMS"), vbCritical
'   Call FechaDB
'   Exit Sub
'End If
    
'txtAliquotaICMS = ICM!icmAliquota & "%"
   
If Not ((pes!chcarteirarep = Empty Or pes!chcarteirarep = "NENHUM")) Then
   CartRep.Open "Select * from Carteira_Rep where chCarteiraRep = ('" & pes!chcarteirarep & "')", db, 3, 3
   If CartRep.EOF Then
      MsgBox ("Representante não cadastrado. Verificar"), vbInformation
      Call FechaDB
      Exit Sub
   End If
End If

If Not (pes!chCarteiraPromot = Empty Or pes!chCarteiraPromot = "NENHUM") Then
   CartPromot.Open "Select * from Carteira_Promot where chCarteiraPromot = ('" & pes!chCarteiraPromot & "')", db, 3, 3
   If CartPromot.EOF Then
      MsgBox ("Promotor não Cadastrado. Verificar."), vbInformation
      Call FechaDB
      Exit Sub
   End If
End If

txtFaturamento = neg!negFaturamento
txtIntervalo = neg!negIntervaloFatura
txtDataPrimParc = Data_Proc + neg!negAPartirDe
Data_Cobranca = Data_Proc + neg!negAPartirDe

Fim_Carga = 0
Linha = 0
Intervalo = 0

AcumValorConsig(0) = 0
AcumValorConsig(1) = 0
AcumValorConsig(2) = 0

'TabDetNeg.MoveFirst
Encontrei = 0

dneg.Open "Select * from DetalheNegociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtCompPedido & "')", db, 3, 3
If dneg.EOF Then
   MsgBox ("Medição sem Detalhe de Negociação. Verificar."), vbCritical
   Call FechaDB
   Exit Sub
End If


Linha = 0
Base_Fatura_Lart = 0
Acumula_Total_Fabricante(0) = 0
Base_Fatura_Merco = 0
Acumula_Total_Fabricante(1) = 0

Acumula_Comis_Rep(0) = 0
Acumula_Comis_Rep(1) = 0
AcumValorConsig(0) = 0
AcumValorConsig(1) = 0
AcumValorConsig(2) = 0

GridNegocio.ColAlignment(0) = 1
Linha = 0

dneg.MoveFirst

Do While Fim_Carga = 0
    
    If Prod.State = 1 Then
       Prod.Close: Set Prod = Nothing
    End If
    
    Prod.Open "Select * from Produto where chproduto = ('" & dneg!chProduto & "')", db, 3, 3
    If Prod.EOF Then
       MsgBox ("Produto não cadastrado. Verificar"), vbInformation
       Call FechaDB
       Exit Sub
    End If
    Linha = Linha + 1
    GridNegocio.Rows = Linha + 1
    GridNegocio.TextMatrix(Linha, 0) = Prod!prdNomeProd
    Mneu_Produto(Linha) = Prod!chProduto
    If dneg!pedunidade = 0 Then
       GridNegocio.TextMatrix(Linha, 1) = "Un"
    Else
       GridNegocio.TextMatrix(Linha, 1) = "M"
    End If
    GridNegocio.TextMatrix(Linha, 3) = Format$(dneg!pedPrecoUnidadePedida, "0.00")
    GridNegocio.TextMatrix(Linha, 2) = dneg!pedquantidadePedida
    GridNegocio.TextMatrix(Linha, 4) = Format$((dneg!pedquantidadePedida * GridNegocio.TextMatrix(Linha, 3)), "0.00")
    Valor_Produto = Format$(GridNegocio.TextMatrix(Linha, 4), "#,##0.00")
    GridNegocio.TextMatrix(Linha, 5) = dneg!pedqtddias
    GridNegocio.TextMatrix(Linha, 6) = Format$((dneg!pedqtddias * Valor_Produto), "#,##0.00")
    'Valor_Frete = Format$(GridNegocio.TextMatrix(Linha, 5), "##0.00")
    
    '   GridNegocio.TextMatrix(Linha, 6) = Format$((dneg!pedPrecoUnidadePedida ), "0,00
    
    '   Valor_IPI = Format$(GridNegocio.TextMatrix(Linha, 6), "0.00")
    '   GridNegocio.TextMatrix(Linha, 7) = Format$((dneg!pedquantidadePedida ), "0.00")
    '   Valor_Desconto = Format$(GridNegocio.TextMatrix(Linha, 7), "#0.00")
    Valor_Total = Format$((Valor_Produto) * dneg!pedqtddias, "#,##0.00")
    '  GridNegocio.TextMatrix(Linha, 8) = Format$(Valor_Total, "##,##0.00")
    
    IndNegSalvo = Linha
    
    'Acumula por tipo de produto, para calculo de comissao de consignacao.(Piso ou Revestimento)
    
    '  AcumValorConsig(tabproduto("prdtipo ) = AcumValorConsig(tabproduto("prdtipo ) + Valor_Total
             
    Acumula_Quantidade = Acumula_Quantidade + dneg!pedquantidadePedida
    Acumula_Valor_Total = Acumula_Valor_Total + Valor_Total
    Acumula_Valor_Produto = Acumula_Valor_Produto + Valor_Total
    Acumula_Valor_Frete = Acumula_Valor_Frete + Valor_Frete
    '  Acumula_Valor_IPI = Acumula_Valor_IPI + Valor_IPI
    '  Acumula_Valor_Desconto = Acumula_Valor_Desconto + Valor_Desconto
    
    'IndFabricante = tabproduto("prdFabricante
    IndFabricante = 0
    
    Acumula_Total_Fabricante(IndFabricante) = Acumula_Total_Fabricante(IndFabricante) + Valor_Total
    Acumula_Produto_Fabricante(IndFabricante) = Acumula_Produto_Fabricante(IndFabricante) + Valor_Produto
    Acumula_Frete_Fabricante(IndFabricante) = Acumula_Frete_Fabricante(IndFabricante) + Valor_Frete
    Acumula_IPI_Fabricante(IndFabricante) = Acumula_IPI_Fabricante(IndFabricante) + Valor_IPI
    Acumula_Desconto_Fabricante(IndFabricante) = Acumula_Desconto_Fabricante(IndFabricante) + Valor_Desconto
    Acumula_Quantidade_Fabricante(IndFabricante) = Acumula_Quantidade_Fabricante(IndFabricante) + dneg!pedquantidadePedida
    
    Acumula_Comissao_Rep = Acumula_Comissao_Rep + dneg!pedcomissaorep
    Acumula_Comissao_Promot = Acumula_Comissao_Promot + dneg!pedcomissaopromot
    
    'Na segunda fase usamos o acumulado por fabricante para gerar o valor da comissao isoladamente
    Acumula_Comis_Rep(IndFabricante) = Acumula_Comis_Rep(IndFabricante) + dneg!pedcomissaorep
    Acumula_Comis_Promot(IndFabricante) = Acumula_Comis_Promot(IndFabricante) + dneg!pedcomissaopromot
    
    Valor_Total = 0
    Valor_Produto = 0
    Valor_Frete = 0
    Valor_IPI = 0
    Valor_Desconto = 0
             
    dneg.MoveNext
    If dneg.EOF Then
       Fim_Carga = 1
    End If
    
    'Acumula_Comissao_Promot = Acumula_Comis_Promot(0) + Acumula_Comis_Promot(1)
    
    
    'If neg!NEGCOBRANCAFRETE = 4 Or neg!NEGCOBRANCAFRETE = 5 Then
    '   Acumula_Valor_Frete = neg!negvalorfixofrete
    '   txtResTotalFrete = neg!negvalorfixofrete
    'End If
    
    txtValorTotal = Format$(Acumula_Valor_Total, "#,##0.00")
    txtResValorTotalNF = Format$(Acumula_Valor_Total + Acumula_Valor_Frete, "#,##0.00")
    'txtTotalPUQtd = Format$(Acumula_Valor_Produto, "#,##0.00")
    txtResTotalProduto = Format$(Acumula_Valor_Produto, "#,##0.00")
    'txtTotalFrete = Format$(Acumula_Valor_Frete, "##0.00")
    txtResTotalFrete = Format$(Acumula_Valor_Frete, "##0.00")
    'txtTotalIPI = Format$(Acumula_Valor_IPI, "##0.00")
    txtResTotalIPI = Format$(Acumula_Valor_IPI, "##0.00")
    'txtTotalDesc = Format$(Acumula_Valor_Desconto, "##0.00")
    txtResTotalDesconto = Format$(Acumula_Valor_Desconto, "##0.00")
    'txtTotalQtd = Acumula_Quantidade
    txtResTotalICMS = Format$(((Acumula_Valor_Produto + Acumula_Valor_Frete) - Acumula_Valor_Desconto) * (ICM!icmAliquota / 100), "#0.00")

Loop

Guarda_Cliente = pes!chPessoa

'If neg!negtransporte = "Cliente" Then
'   txtTransportadora = TabNegociacao("negtransporte")
'   txtMotorista = "Cliente"
'   txtIdentidade = "Cliente"
'   txtPlaca = "Cliente"
'   txtUFPlaca = "Cliente"
'   txtMunicipioPlaca = "Cliente"
'Else
'   Tabpessoa.Seek "=", TabNegociacao("negtransporte")

'   TabCompTransporte.Seek "=", Tabpessoa("chPessoa"), TabNegociacao("negplaca")
'   If TabCompTransporte.NoMatch Then
'      MsgBox ("nao encontrei")
'      Resp = MsgBox("Vou assumir a placa informada. Confirma: S/N???, vbyesno")
'      If Resp = vbNo Then
'         A = 1 / 0
'      Else
'         txtMunicipioPlaca = Tabpessoa("pescidade")
'         txtUFPlaca = Tabpessoa("chuf")
'      End If
'   Else
'      txtPlaca = TabCompTransporte("chTransPlaca")
'      txtUFPlaca = TabCompTransporte("comtransuf")
'      txtMunicipioPlaca = TabCompTransporte("comtransmunicipio")
'   End If
'   txtTransportadora = TabNegociacao("negtransporte")
'   txtMotorista = Tabpessoa("pesRazaoSocial")
'   txtIdentidade = Tabpessoa("chCNPJ_CPF")
   
'End If

txtRepresentante = pes!chcarteirarep
txtPromotora = pes!chCarteiraPromot
'Tabpessoa.Seek "=", Guarda_Cliente

'Alteração para implementação da rotina de Consignação da C&C

'Rotina de Calculo da distribuicao do faturamento

If txtFaturamento = 0 Then
   Base_Fatura = 0
Else
   Base_Fatura = Format$(Acumula_Valor_Total / txtFaturamento, "##,##0.00")
End If

Base_Frete = Acumula_Valor_Frete

If (Acumula_Total_Fabricante(0)) > 0 Then
   Base_Fatura_Lart = Format$(Acumula_Total_Fabricante(0) / txtFaturamento, "##,##0.00")
End If

If (Acumula_Total_Fabricante(1)) > 0 Then
   Base_Fatura_Merco = Format$(Acumula_Total_Fabricante(1) / txtFaturamento, "##,##0.00")
End If

ctr.Open "Select * from Contas_A_Receber where chNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3
If ctr.EOF Then
   Call FechaDB
   MsgBox ("Nota Fiscal inexistente em Negociação"), vbCritical
   Exit Sub
End If


'TabCtaReceber.Seek "=", 0, txtCliente, txtNotaFiscal, "ICMS-ST NF-"
'If TabCtaReceber.NoMatch Then
'   Base = 0
'Else
'   Base = 1
'End If

'If TabCobrancaFrete("chcodcobrancafrete") = 1 Then 'Indica que a boleta de frete é separada e portanto o frete ficara na primeira ocorrencia da tabela
'   Ind_Aux = 1 + Base
'   Ind_Inicial_Frete = 1
'   Ind_Inicial_Fatura = 2
'   Parcelas_Frete = 1
'   Data_Cobranca = txtDataFrete
'Else
'   Ind_Aux = 0 + Base
'   Ind_Inicial_Fatura = 1
'   Ind_Inicial_Frete = 1
'   If TabCobrancaFrete("chcodcobrancafrete") = 3 Then
'      Parcelas_Frete = txtFaturamento
'   Else
'      Parcelas_Frete = 1
'   End If
'End If
      
'Carga de tabela com valores para fatura

For A = 0 To 9
    Valor_Fatura(A) = 0
    Fatura_Lart(A) = 0
    Fatura_Merco(A) = 0
    Frete_Lart(A) = 0
    Frete_Merco(A) = 0
    Frete_Fatura(A) = 0
Next

Ind_Inicial_Fatura = Ind_Inicial_Fatura + Base

For A = Ind_Inicial_Fatura To (txtFaturamento + Ind_Aux)
   Valor_Fatura(A) = Base_Fatura
   Fatura_Lart(A) = Format$(Base_Fatura_Lart, "##,##0.00")
   Fatura_Merco(A) = Format$(Base_Fatura_Merco, "##,##0.00")
Next

If Fatura_Lart(A - 1) > 0 Then
   Fatura_Lart(A - 1) = Fatura_Lart(A - 1) + (Acumula_Total_Fabricante(0) - (Base_Fatura_Lart * txtFaturamento))
   Valor_Fatura(A - 1) = Fatura_Lart(A - 1)
End If

If Fatura_Merco(A - 1) > 0 Then
   Fatura_Merco(A - 1) = Fatura_Merco(A - 1) + (Acumula_Total_Fabricante(1) - (Base_Fatura_Merco * txtFaturamento))
   Valor_Fatura(A - 1) = Fatura_Merco(A - 1)
End If

'Inicio do tratamento para carga dos valores para frete

'If txtEmissorNF = "Mercopiso" Then
'   Merco_Frete = Base_Frete
'   Lart_Frete = 0
'Else
''   Lart_Frete = Base_Frete
'   Merco_Frete = 0
'End If

'For A = Ind_Inicial_Frete To (Parcelas_Frete + Base)
'       Frete_Fatura(A) = Base_Frete / Parcelas_Frete
'       Frete_Lart(A) = Lart_Frete / Parcelas_Frete
'       Frete_Merco(A) = Merco_Frete / Parcelas_Frete
'Next

'Carga do gridFatura





If Base = 1 Then
   A = 1
   GridFatura.Rows = A + 1
   GridFatura.TextMatrix(A, 0) = ctr!ctrDescricaoOperacao
   GridFatura.TextMatrix(A, 1) = txtNotaFiscal
   GridFatura.TextMatrix(A, 2) = "ICMS-ST"
   'gridFatura.TextMatrix(A, 3) = Data_Cobranca + Intervalo
   GridFatura.TextMatrix(A, 3) = ctr!ctrDataVencito
   GridFatura.TextMatrix(A, 4) = Format$(ctr!ctrvalordaboleta, "#,##0.00")
   GridFatura.TextMatrix(A, 5) = Format$(Fatura_Merco(A), "#,##0.00")
   GridFatura.TextMatrix(A, 6) = Format$(ctr!ctrvalordaboleta, "#,##0.00")
End If

For A = (1 + Base) To (txtFaturamento + Ind_Aux)
   
'Alteração: "Fatura" pelo número da Nota Fiscal em descrição
   GridFatura.Rows = A + 1
   If (Fatura_Lart(A) + Fatura_Merco(A)) > 0 And Frete_Fatura(A) > 0 Then
      GridFatura.TextMatrix(A, 0) = "NF-" & txtNotaFiscal & "-" & (A - Ind_Aux) & "/" & txtFaturamento & (" + Frete")
   Else
      If (Fatura_Lart(A) + Fatura_Merco(A)) = 0 Then
         GridFatura.TextMatrix(A, 0) = "Frete da NF " & txtNotaFiscal
      Else
         GridFatura.TextMatrix(A, 0) = "NF-" & txtNotaFiscal & "-" & (A - Ind_Aux) & "/" & txtFaturamento
      End If
   End If

   GridFatura.TextMatrix(A, 1) = txtNotaFiscal
   GridFatura.TextMatrix(A, 2) = (A - Ind_Aux)
   'gridFatura.TextMatrix(A, 3) = Data_Cobranca + Intervalo
   GridFatura.TextMatrix(A, 3) = ctr!ctrDataVencito
   GridFatura.TextMatrix(A, 4) = Format$(Fatura_Lart(A) + Frete_Lart(A), "#,##0.00")
   GridFatura.TextMatrix(A, 5) = Format$(Fatura_Merco(A) + Frete_Merco(A), "#,##0.00")
   GridFatura.TextMatrix(A, 6) = Format$((Fatura_Lart(A) + Fatura_Merco(A)) + (Frete_Lart(A) + Frete_Merco(A)), "#,##0.00")
   'Alterado em 18/09/2003
  ' If TabCobrancaFrete("chcodcobrancafrete") = 1 And (A - Base) = 1 Then
  '    Data_Cobranca = txtDataPrimParc
  ' Else
  '    Intervalo = TabNegociacao("negintervalofatura")
  ' End If
   IndFatSalvo = A
   Data_Cobranca = Data_Cobranca + Intervalo
   'Intervalo = TabNegociacao("negintervalofatura")
   Acumula_Fatura_Lart = Format$(Acumula_Fatura_Lart + Fatura_Lart(A) + Frete_Lart(A), "#,##0.00")
   Acumula_Fatura_Merco = Format$(Acumula_Fatura_Merco + Fatura_Merco(A) + Frete_Merco(A), "#,##0.00")
   Acumula_Fatura_Geral = Format$(Acumula_Fatura_Geral + Valor_Fatura(A) + Frete_Fatura(A), "#,##0.00")
Next
txtTotalFaturaLart = Format$(Acumula_Fatura_Lart, "#,##0.00")
txtTotalFaturaMerco = Format$(Acumula_Fatura_Merco, "#,##0.00")
txtTotalFatura = Format$((Acumula_Fatura_Lart + Acumula_Fatura_Merco), "#,##0.00")

Acumula_Valor_Total = 0
Acumula_Valor_Produto = 0
Acumula_Valor_Frete = 0
Acumula_Valor_IPI = 0
Acumula_Valor_Desconto = 0
Acumula_Quantidade = 0

Call FechaDB

End Sub

Public Sub Rotina_Limpa_Form()

txtCliente = Empty
txtNumPedido = Empty
txtCompPedido = Empty
txtDataPedido = "__/__/____"
'txtNotaFiscal = Empty
'txtDataEmissao = "__/__/____"
txtEmissorNF = Empty
txtCondProcess = Empty
txtAliquotaICMS = Empty
txtBancoFat = Empty
txtFaturamento = Empty
txtIntervalo = Empty
txtDataPrimParc = "__/__/____"
txtCondFrete = Empty
txtDataFrete = "__/__/____"
'txtTotalQtd = Empty
'txtTotalFrete = Empty
'txtTotalIPI = Empty
txtValorTotal = Empty
txtTotalFaturaLart = Empty
txtTotalFaturaMerco = Empty
txtTotalFatura = Empty
txtTransportadora = Empty
txtOrdemDeCarga = Empty
txtMotorista = Empty
txtIdentidade = Empty
txtMunicipioPlaca = Empty
txtUFPlaca = Empty
txtResTotalProduto = Empty
txtResTotalFrete = Empty
txtResTotalIPI = Empty
txtResTotalICMS = Empty
txtResTotalDesconto = Empty
txtResValorTotalNF = Empty
txtRepresentante = Empty
txtComissaoRep = Empty
txtPromotora = Empty
txtComissaoPromot = Empty

Acumula_Valor_Total = 0
Acumula_Valor_Produto = 0
Acumula_Valor_Frete = 0
Acumula_Valor_IPI = 0
Acumula_Valor_Desconto = 0
Acumula_Quantidade = 0

Call Rotina_Limpa_GridNegocio

Call Rotina_Limpa_GridFatura

End Sub


Public Sub Rotina_Limpa_GridNegocio()
Dim Ind As Integer

GridNegocio.Rows = 2

Ind = 1
    GridNegocio.TextMatrix(Ind, 0) = Empty
    GridNegocio.TextMatrix(Ind, 1) = Empty
    GridNegocio.TextMatrix(Ind, 2) = Empty
    GridNegocio.TextMatrix(Ind, 3) = Empty
    GridNegocio.TextMatrix(Ind, 4) = Empty
    GridNegocio.TextMatrix(Ind, 5) = Empty
    GridNegocio.TextMatrix(Ind, 6) = Empty
    'GridNegocio.TextMatrix(Ind, 7) = Empty
    'GridNegocio.TextMatrix(Ind, 8) = Empty

End Sub

Public Sub Rotina_Limpa_GridFatura()
Dim Ind As Integer

GridFatura.Rows = 2

Ind = 1
    GridFatura.TextMatrix(Ind, 0) = Empty
    GridFatura.TextMatrix(Ind, 1) = Empty
    GridFatura.TextMatrix(Ind, 2) = Empty
    GridFatura.TextMatrix(Ind, 3) = Empty
    GridFatura.TextMatrix(Ind, 4) = Empty

End Sub


'Public Sub Rotina_MoverDados_CtaPagar()
'TabCtaPagar("ctpdataemissao") = txtDataEmissao
'TabCtaPagar("ctpdatalanc") = Date

'If Tipo_Comissao = 1 Then
'   TabCtaPagar("ctpDescricaoOperacao") = "Comissão de Representante"
''Else
'   TabCtaPagar("ctpDescricaoOperacao") = "Comissão de Promotor"
      
'End If

'TabCtaPagar("chano") = Ano_Comissao

'TabCtaPagar("chmes") = Mes_Comissao

'TabCtaPagar("chdia") = 1
'TabBanco.Seek "=", 0, TabNegociacao("chcodbcolart")
'If TabBanco.NoMatch Then
'   MsgBox ("Banco inválido"), , frmPedido.cmbBanco
'   Mes_Comissao = 1 / 0
'End If
'txtBancoFat = TabNegociacao("bcosiglabco")
'TabCtaPagar("chcodbcolart") = TabBanco("bcosiglabco")
'End Sub


Public Sub Rotina_Nota_Fiscal_Historico()

DataString = hneg!chdatanegociacao

Data_Proc = DataString

Ano = Year(Data_Proc)
Mes = Month(Data_Proc)
Dia = Day(Data_Proc)

Ano_Comissao = Ano
Mes_Comissao = Mes + 1

Dia_Comis_Repres = 25
Dia_Comis_Promot = 5

If Mes_Comissao > 12 Then
   Mes_Comissao = 1
   Ano_Comissao = Ano_Comissao + 1
End If

Data_Comissao = 1 & "/" & Mes_Comissao & "/" & Ano_Comissao

Call Rotina_Limpa_Form

For A = 0 To 9
    Valor_Fatura(A) = 0
    Fatura_Lart(A) = 0
    Fatura_Merco(A) = 0
    Frete_Lart(A) = 0
    Frete_Merco(A) = 0
    Frete_Fatura(A) = 0
Next

For A = 0 To 99
    Mneu_Produto(A) = Empty
Next

Acumula_Fatura_Lart = 0
Acumula_Fatura_Merco = 0
Acumula_Fatura_Geral = 0
Acumula_Comis_Rep(0) = 0
Acumula_Comis_Rep(1) = 0
Acumula_Comis_Promot(0) = 0
Acumula_Comis_Promot(1) = 0
Acumula_Comissao_Rep = 0
Acumula_Comissao_Promot = 0

'TabHistoricoNegociacao.Seek "=", TabHistNegNF("chpessoa"), TabHistNegNF("chnumpedido"), TabHistNegNF("chnumpedidocomp")
'If TabHistoricoNegociacao.NoMatch Then
'   MsgBox "Numero de pedido inválido"
'   Unload Me
'End If

txtDataEmissao = hneg!chdatanegociacao
txtEmissorNF = "SHB Brasil"
'Call Rotina_AbrirBanco

Bco.Open "Select * from Banco where bcoCodBcoLart = ('" & hneg!chcodbcolart & "')", db, 3, 3
If Bco.EOF Then
   MsgBox ("Banco não encontrado em carga de historico de negociação"), vbCritical
   Call FechaDB
   Exit Sub
End If



'TabBanco.Seek "=", 0, TabHistoricoNegociacao("chcodbcolart")
'If TabBanco.NoMatch Then
'   MsgBox "Numero do Banco inválido"
'   Unload Me
'End If

txtCliente = hneg!chPessoa
txtDataPedido = hneg!hngDataPedido
txtNumPedido = hneg!chNumpedido
txtCompPedido = hneg!chNumPedidoComp
txtNotaFiscal = hneg!hngnotafiscal
txtBancoFat = Bco!bcosiglabco
txtBancoFatN = Bco!bcoBodBcoFEBRABAN
'If IsNull(hneg!chordemdecarga) Then
'   txtOrdemDeCarga = "Não Informado"
'Else
'   txtOrdemDeCarga = hneg!chordemdecarga
'End If
txtEmissorNF = "SHB Brasil"

'TabCobrancaFrete.Seek "=", hneg!hngCOBRANCAFRETE

'If TabCobrancaFrete.NoMatch Then
'   MsgBox ("Erro na Leitura do Parametro Cobranca Frete
'   A = 1 / 0
'Else
'   txtCondFrete = TabCobrancaFrete("parDescCobrancaFrete
'   If hneg!hngCOBRANCAFRETE = 1 Then
'      txtDataFrete = Data_Proc + hneg!hngBoletaFrete
'   Else
'      txtDataFrete = "__/__/____"
'   End If
'End If

'TabCondProcessamento.Seek "=", hneg!hngCondProcess

'If TabCondProcessamento.NoMatch Then
'   MsgBox ("Erro na Leitura do Parametro Condição de Processamento
'   A = 1 / 0
'Else

CondProc.Open "Select * from CondProcessamento where chCondicaoProcessamento = ('" & hneg!hngCondProcess & "')", db, 3, 3
If CondProc.EOF Then
   MsgBox ("Erro na Leitura do Parâmetro Condição de Processamento"), vbCritical
   Call FechaDB
   Exit Sub
Else
   txtCondProcess = CondProc!cprDescCondProcess
End If

txtPercDescComis = Format$(hneg!hngdesccomissao, "#0.00") & "%"
 
pes.Open "Select * from Pessoa where chPessoa = ('" & hneg!chPessoa & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Erro no acesso a Pessoa - Extrato Nota Fiscal"), vbCritical
   Call FechaDB
   Exit Sub
End If

'ICM.Open "Select * from ICMS where chUF = ('" & pes!chUF & "')", db, 3, 3
'If ICM.EOF Then
'   MsgBox ("Estado não cadastrado na Tabela de ICMS"), vbCritical
'   Call FechaDB
'   Exit Sub
'End If
    
'txtAliquotaICMS = ICM!icmAliquota & "%"
   
If Not ((pes!chcarteirarep = Empty Or pes!chcarteirarep = "NENHUM")) Then
   CartRep.Open "Select * from Carteira_Rep where chCarteiraRep = ('" & pes!chcarteirarep & "')", db, 3, 3
   If CartRep.EOF Then
      MsgBox ("Representante não cadastrado. Verificar"), vbInformation
      Call FechaDB
      Exit Sub
   End If
End If

If Not (pes!chCarteiraPromot = Empty Or pes!chCarteiraPromot = "NENHUM") Then
   CartPromot.Open "Select * from Carteira_Promot where chCarteiraPromot = ('" & pes!chCarteiraPromot & "')", db, 3, 3
   If CartPromot.EOF Then
      MsgBox ("Promotor não Cadastrado. Verificar."), vbInformation
      Call FechaDB
      Exit Sub
   End If
End If

txtFaturamento = hneg!hngFaturamento
txtIntervalo = hneg!hngintervalofatura
txtDataPrimParc = Data_Proc + hneg!hngAPartirDe
Data_Cobranca = Data_Proc + hneg!hngAPartirDe

Fim_Carga = 0
Linha = 0
Intervalo = 0

AcumValorConsig(0) = 0
AcumValorConsig(1) = 0
AcumValorConsig(2) = 0

'tabhistoricodetneg.MoveFirst
Encontrei = 0
Encontrei = 0

hdneg.Open "Select * from HistoricoDetalheNegociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtCompPedido & "')", db, 3, 3
If hdneg.EOF Then
   MsgBox ("Medição sem Detalhe no Histórico de Negociação. Verificar."), vbCritical
   Call FechaDB
   Exit Sub
End If

Linha = 0
Base_Fatura_Lart = 0
Acumula_Total_Fabricante(0) = 0
Base_Fatura_Merco = 0
Acumula_Total_Fabricante(1) = 0

Acumula_Comis_Rep(0) = 0
Acumula_Comis_Rep(1) = 0
AcumValorConsig(0) = 0
AcumValorConsig(1) = 0
AcumValorConsig(2) = 0

GridNegocio.ColAlignment(0) = 1

hdneg.MoveFirst

Do While Fim_Carga = 0
   
   If Prod.State = 1 Then
      Prod.Close: Set Prod = Nothing
   End If
    
    Prod.Open "Select * from Produto where chproduto = ('" & hdneg!chProduto & "')", db, 3, 3
    If Prod.EOF Then
       MsgBox ("Produto não cadastrado. Verificar"), vbInformation
       Call FechaDB
       Exit Sub
    End If
      GridNegocio.Rows = Linha + 1
      GridNegocio.TextMatrix(Linha, 0) = Prod!prdNomeProd
      Mneu_Produto(Linha) = Prod!chProduto
      GridNegocio.TextMatrix(Linha, 1) = Unidade(hdneg!hdnUnidade)
      GridNegocio.TextMatrix(Linha, 2) = Format$(hdneg!hdnPrecoUnidadePedida, "0.00")
      GridNegocio.TextMatrix(Linha, 3) = hdneg!hdnquantidadePedida
      GridNegocio.TextMatrix(Linha, 4) = Format$((hdneg!hdnquantidadePedida * GridNegocio.TextMatrix(Linha, 2)), "0.00")
      Valor_Produto = Format$(GridNegocio.TextMatrix(Linha, 4), "#,##0.00")
      'gridNegocio.TextMatrix(Linha, 5) = Format$((hdneg!hdnquantidadePedida * hdneg!hngFreteUnidadePedida), "0.00")
      'Valor_Frete = Format$(gridNegocio.TextMatrix(Linha, 5), "##0.00")
    
      GridNegocio.TextMatrix(Linha, 5) = Format$(((hdneg!hdnPrecoUnidadePedida) * hdneg!hdnquantidadePedida) * Prod!prdIPI / 100, "#,##0.00")
      
      'Valor_IPI = Format$(GridNegocio.TextMatrix(Linha, 6), "0.00")
      GridNegocio.TextMatrix(Linha, 6) = Format$((hdneg!hdnquantidadePedida), "0.00")
      Valor_Desconto = Format$(GridNegocio.TextMatrix(Linha, 6), "#0.00")
      Valor_Total = Format$((Valor_Produto + Valor_IPI) - (Valor_Desconto), "#,##0.00")
      GridNegocio.TextMatrix(Linha, 6) = Format$(Valor_Total, "##,##0.00")
      
      IndNegSalvo = Linha
      
      'Acumula por tipo de produto, para calculo de comissao de consignacao.(Piso ou Revestimento)
      
      AcumValorConsig(Prod!prdtipo) = AcumValorConsig(Prod!prdtipo) + Valor_Total
               
      Acumula_Quantidade = Acumula_Quantidade + hdneg!hdnquantidadePedida
      Acumula_Valor_Total = Acumula_Valor_Total + Valor_Total
      Acumula_Valor_Produto = Acumula_Valor_Produto + Valor_Produto
      Acumula_Valor_Frete = Acumula_Valor_Frete + Valor_Frete
      Acumula_Valor_IPI = Acumula_Valor_IPI + Valor_IPI
      Acumula_Valor_Desconto = Acumula_Valor_Desconto + Valor_Desconto
      
      'IndFabricante = tabproduto("prdFabricante
      IndFabricante = 0
      
      Acumula_Total_Fabricante(IndFabricante) = Acumula_Total_Fabricante(IndFabricante) + Valor_Total
      Acumula_Produto_Fabricante(IndFabricante) = Acumula_Produto_Fabricante(IndFabricante) + Valor_Produto
      Acumula_Frete_Fabricante(IndFabricante) = Acumula_Frete_Fabricante(IndFabricante) + Valor_Frete
      Acumula_IPI_Fabricante(IndFabricante) = Acumula_IPI_Fabricante(IndFabricante) + Valor_IPI
      Acumula_Desconto_Fabricante(IndFabricante) = Acumula_Desconto_Fabricante(IndFabricante) + Valor_Desconto
      Acumula_Quantidade_Fabricante(IndFabricante) = Acumula_Quantidade_Fabricante(IndFabricante) + hdneg!hdnquantidadePedida
      
      Acumula_Comissao_Rep = Acumula_Comissao_Rep + hdneg!hdncomissaorep
      Acumula_Comissao_Promot = Acumula_Comissao_Promot + hdneg!hdncomissaopromot
      
      'Na segunda fase usamos o acumulado por fabricante para gerar o valor da comissao isoladamente
      Acumula_Comis_Rep(IndFabricante) = Acumula_Comis_Rep(IndFabricante) + hdneg!hdncomissaorep
      Acumula_Comis_Promot(IndFabricante) = Acumula_Comis_Promot(IndFabricante) + hdneg!hdncomissaopromot
      
      Valor_Total = 0
      Valor_Produto = 0
      Valor_Frete = 0
      Valor_IPI = 0
      Valor_Desconto = 0
               
      hdneg.MoveNext
      If hdneg.EOF Then
         Fim_Carga = 1
      Else
         Linha = Linha + 1
      End If


'Acumula_Comissao_Promot = Acumula_Comis_Promot(0) + Acumula_Comis_Promot(1)


'If hneg!hngCOBRANCAFRETE = 4 Or hneg!hngCOBRANCAFRETE = 5 Then
'   Acumula_Valor_Frete = hneg!hngvalorfixofrete
'   txtResTotalFrete = hneg!hngvalorfixofrete
'End If

txtValorTotal = Format$(Acumula_Valor_Total, "#,##0.00")
txtResValorTotalNF = Format$(Acumula_Valor_Total + Acumula_Valor_Frete, "#,##0.00")
'txtTotalPUQtd = Format$(Acumula_Valor_Produto, "#,##0.00")
txtResTotalProduto = Format$(Acumula_Valor_Produto, "#,##0.00")
'txtTotalFrete = Format$(Acumula_Valor_Frete, "##0.00")
txtResTotalFrete = Format$(Acumula_Valor_Frete, "##0.00")
'txtTotalIPI = Format$(Acumula_Valor_IPI, "##0.00")
txtResTotalIPI = Format$(Acumula_Valor_IPI, "##0.00")
'txtTotalDesc = Format$(Acumula_Valor_Desconto, "##0.00")
txtResTotalDesconto = Format$(Acumula_Valor_Desconto, "##0.00")
'txtTotalQtd = Acumula_Quantidade
txtResTotalICMS = Format$((Acumula_Valor_Produto + Acumula_Valor_Frete) - Acumula_Valor_Desconto, "#0.00")

Loop
'If IsNull(hneg!chordemdecarga) Then
'   txtTransportadora = "Não Informado"
'   txtMotorista = "Não Informado"
'   txtIdentidade = "Não Informado"
'   txtPlaca = "Não Informado"
'   txtUFPlaca = "Não Informado"
'   txtMunicipioPlaca = "Não Informado"
'Else
'    Guarda_Cliente = pes!chPessoa
'
'    If hneg!chordemdecarga = "Cliente" Then
 '      txtTransportadora = "Cliente"
  '     txtMotorista = "Cliente"
  '     txtIdentidade = "Cliente"
   ''    txtPlaca = "Cliente"
'       txtUFPlaca = "Cliente"
'       txtMunicipioPlaca = "Cliente"
'    Else
'       TabPagtosEmCheque.Seek "=", hneg!chordemdecarga, "Maxalbido"
 '      If TabPagtosEmCheque.NoMatch Then
'          MsgBox "Ordem de carga não encontrada"
'       Else
'           Tabpessoa.Seek "=", TabPagtosEmCheque("ocgmotorista
'           If Tabpessoa.NoMatch Then
'              MsgBox "Motorista não cadastrado em Pessoa", , TabPagtosEmCheque("ocgmotorista
'              Exit Sub
'           Else
'              TabCompTransporte.Seek "=", Tabpessoa("chPessoa , TabPagtosEmCheque("ocgplaca
'              If TabCompTransporte.NoMatch Then
'                 MsgBox ("nao encontrei
'                 Resp = MsgBox("Vou assumir a placa informada. Confirma: S/N???, vbyesno
'                 If Resp = vbNo Then
'                    A = 1 / 0
'                 Else
'                    txtMunicipioPlaca = Tabpessoa("pescidade
'                    txtUFPlaca = Tabpessoa("chuf
'                 End If
'              Else
'                 txtPlaca = TabCompTransporte("chTransPlaca
'                 txtUFPlaca = TabCompTransporte("comtransuf
'                 txtMunicipioPlaca = TabCompTransporte("comtransmunicipio
'              End If
'           End If
'           txtTransportadora = Tabpessoa("pesrazaosocial
'           txtMotorista = Tabpessoa("chPessoa
''           txtIdentidade = Tabpessoa("chCNPJ_CPF
'       End If
'    End If
'End If
'txtRepresentante = pes!chRepresntante

'txtPromotora = pes!chPromotora
'Tabpessoa.Seek "=", Guarda_Cliente

'Alteração para implementação da rotina de Consignação da C&C

'If txtCondProcess = "CONSIGNAÇÃO" Then
'   txtComissaoRep = Format(0, "#,##0.00")
'   txtComissaoPromot = Format(0, "#,##0.00")
'Else
'   txtComissaoRep = Format(Acumula_Comissao_Rep, "#,##0.00")
'   txtComissaoPromot = Format$(Acumula_Comissao_Promot, "#,##0.00")
'End If
'Rotina de Calculo da distribuicao do faturamento
'Call Rotina_AbrirBanco
hctr.Open "Select * from HistoricoContasReceber where chNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3
If hctr.EOF Then
   Call FechaDB
   MsgBox ("Nota Fiscal inexistente em Historico de Contas a Receber"), vbCritical
   Exit Sub
End If


If txtFaturamento = 0 Then
   Base_Fatura = 0
Else
   Base_Fatura = Format$(Acumula_Valor_Total / txtFaturamento, "##,##0.00")
End If

Base_Frete = Acumula_Valor_Frete

If (Acumula_Total_Fabricante(0)) > 0 And txtFaturamento > 0 Then
   Base_Fatura_Lart = Format$(Acumula_Total_Fabricante(0) / txtFaturamento, "##,##0.00")
Else
   Base_Fatura_Lart = 1
End If

If (Acumula_Total_Fabricante(1)) > 0 Then
   Base_Fatura_Merco = Format$(Acumula_Total_Fabricante(1) / txtFaturamento, "##,##0.00")
End If

'If TabCobrancaFrete("chcodcobrancafrete  = 1 Then 'Indica que a boleta de frete é separada e portanto o frete ficara na primeira ocorrencia da tabela
'   Ind_Aux = 1 + Base
'   Ind_Inicial_Frete = 1
'   Ind_Inicial_Fatura = 2
'   Parcelas_Frete = 1
'   Data_Cobranca = txtDataFrete
'Else
'   Ind_Aux = 0 + Base
'   Ind_Inicial_Fatura = 1
'   Ind_Inicial_Frete = 1
'   If TabCobrancaFrete("chcodcobrancafrete  = 3 Then
'      Parcelas_Frete = txtFaturamento
'   Else
'      Parcelas_Frete = 1
'   End If
'End If
      
'Carga de tabela com valores para fatura

For A = 0 To 9
    Valor_Fatura(A) = 0
    Fatura_Lart(A) = 0
    Fatura_Merco(A) = 0
    Frete_Lart(A) = 0
    Frete_Merco(A) = 0
    Frete_Fatura(A) = 0
  
Next

Ind_Inicial_Fatura = Ind_Inicial_Fatura + Base

For A = Ind_Inicial_Fatura To (txtFaturamento + Ind_Aux)
   Valor_Fatura(A) = Base_Fatura
   Fatura_Lart(A) = Format$(Base_Fatura_Lart, "##,##0.00")
   Fatura_Merco(A) = Format$(Base_Fatura_Merco, "##,##0.00")
Next

If Fatura_Lart(A - 1) > 0 Then
   Fatura_Lart(A - 1) = Fatura_Lart(A - 1) + (Acumula_Total_Fabricante(0) - (Base_Fatura_Lart * txtFaturamento))
   Valor_Fatura(A - 1) = Fatura_Lart(A - 1)
End If

If Fatura_Merco(A - 1) > 0 Then
   Fatura_Merco(A - 1) = Fatura_Merco(A - 1) + (Acumula_Total_Fabricante(1) - (Base_Fatura_Merco * txtFaturamento))
   Valor_Fatura(A - 1) = Fatura_Merco(A - 1)
End If

'Inicio do tratamento para carga dos valores para frete

If txtEmissorNF = "Mercopiso" Then
   Merco_Frete = Base_Frete
   Lart_Frete = 0
Else
   Lart_Frete = Base_Frete
   Merco_Frete = 0
End If

'For A = Ind_Inicial_Frete To Parcelas_Frete
'       Frete_Fatura(A) = Base_Frete / Parcelas_Frete
'       Frete_Lart(A) = Lart_Frete / Parcelas_Frete
'       Frete_Merco(A) = Merco_Frete / Parcelas_Frete
'Next

'Carga do gridFatura

If Base = 1 Then
   A = 1
   GridFatura.Rows = A + 1
   GridFatura.TextMatrix(A, 0) = hctr!ctrDescricaoOperacao
   GridFatura.TextMatrix(A, 1) = txtNotaFiscal
   GridFatura.TextMatrix(A, 2) = (A - Ind_Aux)
   'gridFatura.TextMatrix(A, 3) = Data_Cobranca + Intervalo
   GridFatura.TextMatrix(A, 3) = hctr!ctrDataVencito
   GridFatura.TextMatrix(A, 4) = Format$(hctr!ctrvalordaboleta, "#,##0.00")
   GridFatura.TextMatrix(A, 5) = Format$(Fatura_Merco(A) + Frete_Merco(A), "#,##0.00")
   GridFatura.TextMatrix(A, 6) = Format$(hctr!ctrvalordaboleta, "#,##0.00")
End If

For A = 1 To (txtFaturamento + Ind_Aux)
   
'Alteração: "Fatura" pelo número da Nota Fiscal em descrição
   GridFatura.Rows = A + 1
   If (Fatura_Lart(A) + Fatura_Merco(A)) > 0 And Frete_Fatura(A) > 0 Then
      GridFatura.TextMatrix(A, 0) = "NF-" & txtNotaFiscal & "-" & (A - Ind_Aux) & "/" & txtFaturamento & (" + Frete ")
   Else
      If (Fatura_Lart(A) + Fatura_Merco(A)) = 0 Then
         GridFatura.TextMatrix(A, 0) = "Frete da NF " & txtNotaFiscal
      Else
         GridFatura.TextMatrix(A, 0) = "NF-" & txtNotaFiscal & "-" & (A - Ind_Aux) & "/" & txtFaturamento
      End If
   End If

   GridFatura.TextMatrix(A, 1) = txtNotaFiscal
   GridFatura.TextMatrix(A, 2) = (A - Ind_Aux)
   'gridFatura.TextMatrix(A, 3) = Data_Cobranca + Intervalo
   GridFatura.TextMatrix(A, 3) = Data_Cobranca
   GridFatura.TextMatrix(A, 4) = Format$(Fatura_Lart(A) + Frete_Lart(A), "#,##0.00")
   GridFatura.TextMatrix(A, 5) = Format$(Fatura_Merco(A) + Frete_Merco(A), "#,##0.00")
   GridFatura.TextMatrix(A, 6) = Format$((Fatura_Lart(A) + Fatura_Merco(A)) + (Frete_Lart(A) + Frete_Merco(A)), "#,##0.00")
   'Alterado em 18/09/2003
   'If TabCobrancaFrete("chcodcobrancafrete  = 1 And A = 1 Then
   '   Data_Cobranca = txtDataPrimParc
   'Else
   '   Intervalo = hneg!hngintervalofatura
   'End If
   IndFatSalvo = A
   Data_Cobranca = Data_Cobranca + Intervalo
   'Intervalo = hneg!hngintervalofatura
   Acumula_Fatura_Lart = Format$(Acumula_Fatura_Lart + Fatura_Lart(A) + Frete_Lart(A), "#,##0.00")
   Acumula_Fatura_Merco = Format$(Acumula_Fatura_Merco + Fatura_Merco(A) + Frete_Merco(A), "#,##0.00")
   Acumula_Fatura_Geral = Format$(Acumula_Fatura_Geral + Valor_Fatura(A) + Frete_Fatura(A), "#,##0.00")
Next
txtTotalFaturaLart = Format$(Acumula_Fatura_Lart, "#,##0.00")
txtTotalFaturaMerco = Format$(Acumula_Fatura_Merco, "#,##0.00")
txtTotalFatura = Format$((Acumula_Fatura_Lart + Acumula_Fatura_Merco), "#,##0.00")

'Acumula_Valor_Total = 0
'Acumula_Valor_Produto = 0
'Acumula_Valor_Frete = 0
'Acumula_Valor_IPI = 0
'Acumula_Valor_Desconto = 0
'Acumula_Quantidade = 0

'Call Rotina_Limpa_GridNegocio

'Call Rotina_Limpa_GridFatura

Call FechaDB

End Sub

'Public Sub Rotina_MoverDados_CtaPagar_Hist()
'TabCtaPagar("ctpdataemissao") = txtDataEmissao
'TabCtaPagar("ctpdatalanc") = Date

'If Tipo_Comissao = 1 Then
'   TabCtaPagar("ctpDescricaoOperacao") = "Comissão de Representante"
'Else
'   TabCtaPagar("ctpDescricaoOperacao") = "Comissão de Promotor"
'
'End If

'TabCtaPagar("chano") = Ano_Comissao

'TabCtaPagar("chmes") = Mes_Comissao

'TabCtaPagar("chdia") = 1
'TabBanco.Seek "=", 0, TabNegociacao("chcodbcolart")
''If TabBanco.NoMatch Then
'   MsgBox ("Banco inválido"), , frmPedido.cmbBanco
'   Mes_Comissao = 1 / 0
'End If
''txtBancoFat = TabNegociacao("bcosiglabco")
'TabCtaPagar("chcodbcolart") = TabBanco("bcosiglabco")
'End Sub
