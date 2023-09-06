VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPO 
   Caption         =   "frmPO"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   2910
   ClientWidth     =   20370
   HelpContextID   =   -2147483646
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   20370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ComboBox cmbNumPO 
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
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ComboBox cmbTipoPO 
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
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtDataPrevista 
      Height          =   495
      Left            =   11040
      TabIndex        =   4
      Top             =   1440
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
      Format          =   382795777
      CurrentDate     =   45125
   End
   Begin VB.ComboBox cmbEndEntrega 
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
      Left            =   7560
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.ComboBox cmbFornecedor 
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
      Left            =   4080
      TabIndex        =   2
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Frame frmEquipamento 
      Height          =   5460
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   20235
      Begin MSComCtl2.DTPicker dtDataEntregaProd 
         Height          =   495
         Left            =   16800
         TabIndex        =   59
         Top             =   720
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
         Format          =   382795777
         CurrentDate     =   45155
      End
      Begin VB.CommandButton cmdEmitePO 
         Caption         =   "Imprimir Ordem de Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   19080
         TabIndex        =   58
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtDesc 
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
         Height          =   495
         Left            =   14880
         TabIndex        =   55
         Top             =   4200
         Width           =   1905
      End
      Begin VB.ComboBox cmbAcordo 
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
         Left            =   6480
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtValorTotal 
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
         Height          =   495
         Left            =   14760
         TabIndex        =   52
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtSaldo 
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
         Left            =   8760
         TabIndex        =   51
         Top             =   4680
         Width           =   2025
      End
      Begin VB.TextBox txtPago 
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
         Left            =   11760
         TabIndex        =   49
         Top             =   4680
         Width           =   2025
      End
      Begin VB.TextBox txtTotal 
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
         Left            =   14880
         TabIndex        =   41
         Top             =   4680
         Width           =   1905
      End
      Begin VB.Frame Frame1 
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
         Height          =   3975
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   7455
         Begin VB.ComboBox cmbFrete 
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
            Left            =   4320
            TabIndex        =   63
            Top             =   3000
            Width           =   3015
         End
         Begin VB.TextBox txtPrazosParcelas 
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
            Left            =   5280
            TabIndex        =   62
            Top             =   2080
            Width           =   1815
         End
         Begin VB.TextBox txtPercPagMoeda 
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
            Height          =   420
            Left            =   1800
            TabIndex        =   47
            Top             =   2080
            Width           =   1095
         End
         Begin VB.TextBox txtValorPgo 
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
            Left            =   4440
            TabIndex        =   44
            Top             =   1320
            Width           =   1935
         End
         Begin VB.ComboBox cmbMoeda 
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
            Left            =   240
            TabIndex        =   43
            Top             =   1320
            Width           =   2535
         End
         Begin VB.ComboBox cmbMetodoPagto 
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
            TabIndex        =   39
            Top             =   3000
            Width           =   3855
         End
         Begin VB.TextBox txtNumParcelas 
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
            Height          =   420
            Left            =   4200
            TabIndex        =   37
            Top             =   2080
            Width           =   615
         End
         Begin VB.TextBox txtDesconto 
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
            Height          =   510
            Left            =   4800
            TabIndex        =   35
            Top             =   600
            Width           =   1575
         End
         Begin VB.ComboBox cmbFormaPagto 
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
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   4335
         End
         Begin VB.Label Label27 
            Caption         =   "Frete"
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
            Left            =   4320
            TabIndex        =   64
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label Label26 
            Caption         =   "Faturamento"
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
            Left            =   5280
            TabIndex        =   61
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "Perc Pgo na moeda"
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
            Left            =   1320
            TabIndex        =   46
            Top             =   1800
            Width           =   2535
         End
         Begin VB.Label Label17 
            Caption         =   "Valor Pago BRL"
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
            Left            =   4440
            TabIndex        =   45
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label16 
            Caption         =   "Moeda"
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
            Left            =   240
            TabIndex        =   42
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label Label14 
            Caption         =   "Método de Pagto"
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
            Left            =   240
            TabIndex        =   38
            Top             =   2640
            Width           =   4095
         End
         Begin VB.Label Label13 
            Caption         =   "Parcelas"
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
            Left            =   3960
            TabIndex        =   36
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Desconto"
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
            Left            =   4800
            TabIndex        =   34
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Forma de Pagamento"
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
            TabIndex        =   32
            Top             =   300
            Width           =   3135
         End
      End
      Begin VB.CommandButton cmdSair 
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
         Height          =   615
         Left            =   19080
         TabIndex        =   30
         Top             =   4560
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid tblEquipamentos 
         Height          =   2655
         Left            =   7680
         TabIndex        =   29
         Top             =   1440
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         FormatString    =   "Descrição                              |Qtd.     |Unid. |Valor Unitario| Valor Total       |Data Entrega||||"
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
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Salvar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   19080
         TabIndex        =   14
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdExcluiDaLista 
         Caption         =   "Exclui da Lista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   19080
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdJogaNaLista 
         Caption         =   "Joga na Lista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   19080
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbDescricao 
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
         Left            =   7680
         TabIndex        =   8
         Top             =   720
         Width           =   3735
      End
      Begin VB.ComboBox cmbClasse 
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
         Left            =   3000
         TabIndex        =   6
         Top             =   720
         Width           =   3375
      End
      Begin VB.ComboBox cmbGrupo 
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
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtValorUnid 
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
         Left            =   13050
         TabIndex        =   11
         Top             =   735
         Width           =   1770
      End
      Begin VB.TextBox txtUnid 
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
         Left            =   12240
         TabIndex        =   10
         Top             =   735
         Width           =   780
      End
      Begin VB.TextBox txtQtd 
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
         Left            =   11400
         TabIndex        =   9
         Top             =   735
         Width           =   780
      End
      Begin VB.Label Label25 
         Caption         =   "Data de Entrega"
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
         Left            =   16800
         TabIndex        =   60
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label22 
         Caption         =   "Desc."
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
         Left            =   13920
         TabIndex        =   54
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Acordo"
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
         Left            =   6480
         TabIndex        =   53
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Saldo"
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
         Left            =   7680
         TabIndex        =   50
         Top             =   4725
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Pago"
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
         Left            =   10920
         TabIndex        =   48
         Top             =   4725
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
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
         Left            =   13920
         TabIndex        =   40
         Top             =   4725
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Classe"
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
         Left            =   3000
         TabIndex        =   26
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Grupo"
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
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total"
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
         Left            =   14880
         TabIndex        =   23
         Top             =   360
         Width           =   2130
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Unitário"
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
         Left            =   13080
         TabIndex        =   22
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unid"
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
         Left            =   12255
         TabIndex        =   21
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd"
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
         Left            =   11400
         TabIndex        =   20
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         Left            =   7680
         TabIndex        =   19
         Top             =   360
         Width           =   3480
      End
   End
   Begin VB.Label lblHoje 
      Caption         =   "07/08/2023"
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
      Left            =   18600
      TabIndex        =   57
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label23 
      Caption         =   "Hoje"
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
      Left            =   19080
      TabIndex        =   56
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Tipo PO"
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
      Left            =   120
      TabIndex        =   28
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Data prevista p/entrega"
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
      Left            =   10440
      TabIndex        =   27
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Endereço de Entrega"
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
      Left            =   7560
      TabIndex        =   24
      Top             =   1125
      Width           =   2655
   End
   Begin VB.Label lblLabel3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   300
      Left            =   4020
      TabIndex        =   17
      Top             =   1125
      Width           =   1395
   End
   Begin VB.Label lblLabel2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número da PO"
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
      Left            =   1860
      TabIndex        =   16
      Top             =   1125
      Width           =   1740
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registro e Emissão de Purchase Order"
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
      TabIndex        =   15
      Top             =   360
      Width           =   5595
   End
End
Attribute VB_Name = "frmPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Linha As Integer
Dim Resp As String
Dim Hoje As Date
Dim Relatorio As String
Dim rel As Object
Dim Sql As String

Private Sub cmbAcordo_LostFocus()
   
   cmbDescricao.Clear
   
   If cmbAcordo = "NÃO" Then
      
      Call Rotina_AbrirBanco
      
         pes.Open "Select nomeProd from supproduto where grupo = ('" & Format$((cmbGrupo.ListIndex + 1), "00") & "') and classe = ('" & Format$((cmbClasse.ListIndex + 1), "000") & "') order by codProd", db, 3, 3

         If pes.EOF Then
      
            MsgBox ("Não existem produtos para essa classe")
            FechaDB
            cmdSair.SetFocus
            Exit Sub
         
         End If
      
         pes.MoveFirst
         cmbDescricao.Clear
      
         Do While Not pes.EOF
      
            cmbDescricao.AddItem pes!nomeProd
            pes.MoveNext
      
         Loop
      
         pes.Close
            
      FechaDB
   
   Else
   
      Call Rotina_AbrirBanco
      
         rs.Open "SELECT nomeProd FROM supacordocomercialdetalhe INNER JOIN supproduto ON supacordocomercialdetalhe.codProd = supproduto.codProd WHERE grupo = ('" & Format$((cmbGrupo.ListIndex + 1), "00") & "') AND classe = ('" & Format$((cmbClasse.ListIndex + 1), "000") & "')", db, 3, 3
         
         If Not rs.EOF Then
         
            rs.MoveFirst
            
            Do While Not rs.EOF
            
               cmbDescricao.AddItem rs!nomeProd
               rs.MoveNext
            
            Loop
         
         End If
         
         rs.Close
      
      FechaDB
   
   End If
End Sub

Private Sub cmbClasse_LostFocus()
   Call Rotina_AbrirBanco
   
      cmbAcordo.Clear
      cmbAcordo.AddItem "NÃO"
      
      rs.Open "SELECT id FROM supacordocomercial WHERE fornecedor = ('" & cmbFornecedor & "') and grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "')", db, 3, 3
      
      If Not rs.EOF Then
      
         cmbAcordo.AddItem rs!id
         
      End If
      
      rs.Close
      
   FechaDB
End Sub

Private Sub cmbDescricao_LostFocus()
   If cmbAcordo = "NÃO" Then
      Dim Resp As String
      Call Rotina_AbrirBanco
      rs.Open "Select * from supproduto where nomeProd = ('" & cmbDescricao & "') and grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "')", db, 3, 3
      If rs.EOF Then
   
         Resp = MsgBox("Produto não cadastrado. Confirma???", vbExclamation + vbYesNo)
   
         If Resp = vbYes Then
            
            frmSupProduto.Show
            frmSupProduto.cmbGrupo = cmbGrupo
            frmSupProduto.txtProduto = cmbDescricao
            frmSupProduto.txtFlag = 1
         Else
            
            cmbDescricao = Empty
            
         End If
      
      Else
      
         Prod.Open "SELECT codProd FROM supProduto WHERE nomeProd = ('" & cmbDescricao & "')", db, 3, 3
         pes.Open "SELECT supAcordoComercial.id,fornecedor FROM supAcordoComercial INNER JOIN supAcordoComercialDetalhe ON supAcordoComercialDetalhe.id = supAcordoComercial.id WHERE grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') AND classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "') AND codProd = ('" & Prod!codProd & "')", db, 3, 3
         Prod.Close
         If Not pes.EOF Then
            MsgBox ("Atenção: Produto existente no acordo " & pes!id & " do fornecedor " & pes!Fornecedor), vbInformation
         End If
         pes.Close
      End If
      FechaDB
   Else
      Call Rotina_AbrirBanco
         pes.Open "SELECT codProd from supProduto where nomeProd=('" & cmbDescricao & "') and grupo = ('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe = ('" & Format$(cmbClasse.ListIndex + 1, "000") & "')", db, 3, 3
         rs.Open "SELECT precoUnit FROM supacordocomercialdetalhe where id = ('" & cmbAcordo & "') and codProd = ('" & pes!codProd & "')", db, 3, 3
         pes.Close
         txtValorUnid = rs!precoUnit
         rs.Close
      FechaDB
   End If
End Sub

Private Sub cmbGrupo_LostFocus()
   Call Rotina_AbrirBanco
   
   pes.Open "Select descricao from supgrupoclasse where grupo = ('" & Format$((cmbGrupo.ListIndex + 1), "00") & "') and classe != 0", db, 3, 3

   If pes.EOF Then

      MsgBox ("Não existem classes para esse grupo")
      FechaDB
      Exit Sub
   
   End If

   pes.MoveFirst
   cmbClasse.Clear

   Do While Not pes.EOF

      cmbClasse.AddItem pes!Descricao
      pes.MoveNext

   Loop

   pes.Close
   
   FechaDB
End Sub

Private Sub cmbNumPO_LostFocus()
   Call Rotina_AbrirBanco
   If cmbTipoPO <> "NOVA" Then
      rs.Open "Select * from suppedidodecompra where id = ('" & cmbNumPO & "')", db, 3, 3
      
         cmbFornecedor = rs!Fornecedor
         cmbEndEntrega = rs!localEntrega
         dtDataPrevista = Date
         If Not IsNull(rs!formaDePagamento) Then
            cmbFormaPagto = rs!formaDePagamento
         End If
         If IsNull(rs!desconto) Then
            txtDesconto = 0
         Else
            txtDesconto = rs!desconto
         End If
         
         If IsNull(rs!numParcelas) Then
            txtNumParcelas = 0
         Else
            txtNumParcelas = rs!numParcelas
         End If
         
         If IsNull(rs!metodoPagamento) Then
            cmbMetodoPagto = Empty
         Else
            cmbMetodoPagto = rs!metodoPagamento
         End If
         
         If Not IsNull(rs!moeda) Then
            cmbMoeda = rs!moeda
         End If
         
         If IsNull(rs!valorPagoBrl) Then
            txtValorPgo = 0
         Else
            txtValorPgo = rs!valorPagoBrl
         End If
         
         If IsNull(rs!valorDesconto) Then
            txtDesc = 0
         Else
            txtDesc = Format$(rs!valorDesconto, "##,##0.00")
         End If
         
         If IsNull(rs!percntPago) Then
            txtPercPagMoeda = 0
         Else
            txtPercPagMoeda = rs!percntPago
         End If
         
         If IsNull(rs!frete) Then
            cmbFrete = Empty
         Else
            cmbFrete = rs!frete
         End If
         
         If IsNull(rs!faturamento) Then
            txtPrazosParcelas = Empty
         Else
            txtPrazosParcelas = rs!faturamento
         End If
         
         txtTotal = Format$(rs!total, "##,##0.00")
         txtPago = Format$(rs!pago, "##,##0.00")
         txtSaldo = Format$(rs!saldo, "##,##0.00")

         
      
      rs.Close
      rs.Open "Select * from suppedidodetalhe where id = ('" & cmbNumPO & "')", db, 3, 3
         If rs.EOF Then
            MsgBox ("Esse pedido de compra não possui produtos")
            tblEquipamentos.Rows = 1
            FechaDB
            Exit Sub
         End If
   
         Linha = 0
         tblEquipamentos.Rows = 2
         
         Do While Not rs.EOF
            
            Prod.Open "Select nomeProd from supproduto where grupo=('" & rs!grupo & "') and classe = ('" & rs!classe & "') and codProd=('" & rs!codProd & "')", db, 3, 3
            
            If Linha = tblEquipamentos.Rows Then
               tblEquipamentos.Rows = tblEquipamentos.Rows + 1
               Linha = tblEquipamentos.Rows - 1
            Else
               Linha = tblEquipamentos.Rows - 1
            End If
               
            tblEquipamentos.TextMatrix(Linha, 0) = Prod!nomeProd
            tblEquipamentos.TextMatrix(Linha, 1) = rs!qtdPedida
            tblEquipamentos.TextMatrix(Linha, 2) = rs!Unidade
            tblEquipamentos.TextMatrix(Linha, 3) = Format$(rs!valorUnitario, "##,##0.00")
            tblEquipamentos.TextMatrix(Linha, 4) = Format$(rs!ValorTotal, "##,##0.00")
            tblEquipamentos.TextMatrix(Linha, 5) = Date
            pes.Open "SELECT descricao FROM supgrupoclasse WHERE grupo = ('" & rs!grupo & "') and classe = '000'", db, 3, 3
            tblEquipamentos.TextMatrix(Linha, 6) = pes!Descricao
            pes.Close
            pes.Open "SELECT descricao FROM supgrupoclasse WHERE grupo = ('" & rs!grupo & "') and classe = ('" & rs!classe & "')", db, 3, 3
            tblEquipamentos.TextMatrix(Linha, 7) = pes!Descricao
            pes.Close
            tblEquipamentos.TextMatrix(Linha, 8) = rs!codProd
            tblEquipamentos.TextMatrix(Linha, 9) = rs!acordo
            Linha = Linha + 1
            rs.MoveNext
            Prod.Close
            
         Loop
         Call calculaTotal
      rs.Close
      FechaDB
   End If
End Sub

Private Sub cmbTipoPO_LostFocus()
   Call Rotina_AbrirBanco
   If cmbTipoPO = "NOVA" Then
      
      cmbNumPO.Enabled = False
   
   Else
         
      cmbNumPO.Enabled = True
      
      Prod.Open "Select id from suppedidodecompra where status = ('" & cmbTipoPO.ListIndex - 1 & "')", db, 3, 3
   
      If Prod.EOF Then
   
         MsgBox ("Não existem pedidos de compra registrados")
         FechaDB
         Exit Sub
      
      End If
   
      Prod.MoveFirst
      cmbNumPO.Clear
   
      Do While Not Prod.EOF
   
         cmbNumPO.AddItem Prod!id
         Prod.MoveNext
   
      Loop
   
      Prod.Close
      
      FechaDB
   
   End If
End Sub

Private Sub cmdEmitePO_Click()

Relatorio = "drOrdemDeCompra"

Set rel = drOrdemDeCompra
Sql = "Select emp.empEmpresa, emp.empEndereco, emp.empCidade, emp.empBairro, emp.empUF, emp.empCEP, emp.empCNPJ, emp.empInscEst, emp.empEMAIL, pes.chPessoa, "
Sql = Sql & " pes.pesRazaoSocial, pes.pesEndereco, pes.pesBairro, pes.pesCidade, pes.chUF, pes.pesCEP, pes.chCNPJ_CPF, pes.pesInscEst_Ident, pes.pesTelContato, "
Sql = Sql & " po.id, po.fornecedor, po.dataPedido, po.dataPrevistaDeEntrega, po.localEntrega,  po.frete,  po.faturamento, prd.descricao, "
Sql = Sql & " det.grupo, det.classe, det.codProd, det.qtdPedida, det.unidade, det.valorUnitario, det.valorTotal, det.dataEntregaProd, "
Sql = Sql & " ender.rua, ender.numero, ender.complemento, ender.bairro, ender.cidade, ender.uf, ender.cep From Empresa emp, supendereco ender, supPedidoDeCompra po, suppedidodetalhe det, Pessoa pes, supproduto prd "
Sql = Sql & " WHERE po.id = ('" & cmbNumPO & "') and det.id = po.id and ender.apelido = ('" & cmbEndEntrega & "') and pes.chPessoa = ('" & cmbFornecedor & "') and det.grupo = prd.grupo and det.classe = prd.classe and det.codProd = prd.codProd "

AbrirRelatorio Sql, rel

Call FechaDB

End Sub

Private Sub cmdExcluiDaLista_Click()
   If tblEquipamentos.Rows = 2 Then
      tblEquipamentos.Rows = 1
   Else
      tblEquipamentos.RemoveItem (tblEquipamentos.Row)
   End If
   Call limparCamposDetalhe
   Call calculaTotal
End Sub
Private Sub cmdJogaNaLista_Click()
   Dim i As Integer
   Dim cod As String
   i = 0
   If tblEquipamentos.Rows > 1 Then
      Do While i < tblEquipamentos.Rows
         If cmbDescricao = tblEquipamentos.TextMatrix(i, 0) Then
            tblEquipamentos.TextMatrix(i, 1) = txtQtd
            tblEquipamentos.TextMatrix(i, 2) = txtUnid
            tblEquipamentos.TextMatrix(i, 3) = Format$(txtValorUnid, "##,##0.00")
            tblEquipamentos.TextMatrix(i, 4) = Format$(txtValorTotal, "##,##0.00")
            tblEquipamentos.TextMatrix(i, 5) = dtDataEntregaProd
            tblEquipamentos.TextMatrix(i, 9) = cmbAcordo
            FechaDB
            Call limparCamposDetalhe
            Call calculaTotal
            Exit Sub
         End If
         i = i + 1
      Loop
      
   End If
   
   If cmbGrupo <> Empty And cmbClasse <> Empty And cmbDescricao <> Empty And txtQtd <> Empty And txtUnid <> Empty And txtValorUnid <> Empty And txtValorTotal <> Empty Then
      If Linha = tblEquipamentos.Rows Then
         tblEquipamentos.Rows = tblEquipamentos.Rows + 1
         Linha = tblEquipamentos.Rows - 1
      Else
         Linha = tblEquipamentos.Rows - 1
      End If
      tblEquipamentos.TextMatrix(Linha, 0) = cmbDescricao
      tblEquipamentos.TextMatrix(Linha, 1) = txtQtd
      tblEquipamentos.TextMatrix(Linha, 2) = txtUnid
      tblEquipamentos.TextMatrix(Linha, 3) = Format$(txtValorUnid, "##,##0.00")
      tblEquipamentos.TextMatrix(Linha, 4) = Format$(txtValorTotal, "##,##0.00")
      tblEquipamentos.TextMatrix(Linha, 5) = dtDataEntregaProd
      tblEquipamentos.TextMatrix(Linha, 6) = cmbGrupo
      tblEquipamentos.TextMatrix(Linha, 7) = cmbClasse
      Call Rotina_AbrirBanco
      rs.Open "Select codProd from supproduto where grupo=('" & Format$(cmbGrupo.ListIndex + 1, "00") & "') and classe=('" & Format$(cmbClasse.ListIndex + 1, "000") & "') and nomeProd=('" & cmbDescricao & "')", db, 3, 3
         cod = rs!codProd
         tblEquipamentos.TextMatrix(Linha, 8) = Format$(cod, "00000")
      rs.Close
      tblEquipamentos.TextMatrix(Linha, 9) = cmbAcordo
      FechaDB
      Linha = Linha + 1
      Call limparCamposDetalhe
      Call calculaTotal
      
   Else
   
      MsgBox ("Informação Inválida! Verificar"), vbCritical

   End If
   
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdSalvar_Click()
cmbNumPO.Enabled = True
Dim ValorUnit As Currency
Dim ValorTotal As Currency
Dim id As Integer
 
   Call Rotina_AbrirBanco
   Dim i As Integer
   If cmbFrete <> Empty And txtPrazosParcelas <> Empty Then
      db.BeginTrans
      If cmbTipoPO = "NOVA" Then
      
         rs.Open "Select * from suppedidodecompra where id = -1", db, 3, 3
         rs.AddNew
         id = -1
         rs!Fornecedor = cmbFornecedor
         rs!DataPedido = Date
         rs!dataPrevistaDeEntrega = dtDataPrevista
         rs!localEntrega = cmbEndEntrega
         rs!Status = 0
         rs!formaDePagamento = cmbFormaPagto
         rs!desconto = txtDesconto
         rs!numParcelas = txtNumParcelas
         rs!metodoPagamento = cmbMetodoPagto
         rs!moeda = cmbMoeda
         rs!valorDesconto = txtDesc
         If txtValorPgo <> Empty Then
            rs!valorPagoBrl = txtValorPgo
         Else
            rs!valorPagoBrl = 0
         End If
         If txtPercPagMoeda <> Empty Then
            rs!percntPago = txtPercPagMoeda
         Else
            rs!percntPago = 0
         End If
         If txtTotal <> Empty Then
            rs!total = txtTotal
         Else
            rs!total = 0
         End If
         If txtPago <> Empty Then
            rs!pago = txtPago
         Else
            rs!pago = 0
         End If
         If txtSaldo <> Empty Then
            rs!saldo = Format$(txtSaldo, "##,##0.00")
         Else
            rs!saldo = Format$(0, "##,##0.00")
         End If
         
         rs!frete = cmbFrete
        
         rs!faturamento = txtPrazosParcelas
      
         rs.Update
         MsgBox ("Salvo com Sucesso!"), vbInformation
         
      Else
          
         rs.Open "Select * from suppedidodecompra where id = ('" & cmbNumPO & "')", db, 3, 3
         id = rs!id
         rs!id = id
         rs!Fornecedor = cmbFornecedor
         rs!DataPedido = Date
         rs!dataPrevistaDeEntrega = dtDataPrevista
         rs!localEntrega = cmbEndEntrega
         rs!formaDePagamento = cmbFormaPagto
         rs!desconto = txtDesconto
         rs!numParcelas = txtNumParcelas
         rs!metodoPagamento = cmbMetodoPagto
         rs!moeda = cmbMoeda
         If txtValorPgo <> Empty Then
            rs!valorPagoBrl = txtValorPgo
         Else
            rs!valorPagoBrl = 0
         End If
         If txtPercPagMoeda <> Empty Then
            rs!percntPago = txtPercPagMoeda
         Else
            rs!percntPago = 0
         End If
         If txtTotal <> Empty Then
            rs!total = Format$(txtTotal, "##,##0.00")
         Else
            rs!total = Format$(0, "##,##0.00")
         End If
         If txtPago <> Empty Then
            rs!pago = Format$(txtPago, "##,##0.00")
         Else
            rs!pago = Format$(0, "##,##0.00")
         End If
         If txtSaldo <> Empty Then
            rs!saldo = Format$(txtSaldo, "##,##0.00")
         Else
            rs!saldo = Format$(0, "##,##0.00")
         End If
         rs.Update
         MsgBox ("Atualizado com Sucesso!"), vbInformation
      
      End If
      rs.Close
      If id = -1 Then
         rs.Open "SELECT id FROM suppedidodecompra ORDER BY id DESC LIMIT 1", db, 3, 3
         id = rs!id
         rs.Close
      End If
      i = 1
      
      
      db.Execute ("DELETE FROM suppedidodetalhe where id=('" & cmbNumPO & "')")
   
         
      Do While i < tblEquipamentos.Rows
          ValorUnit = Format(tblEquipamentos.TextMatrix(i, 3), "#,##0.00")
          ValorTotal = Format$(tblEquipamentos.TextMatrix(i, 4), "##,##0.00")
          rs.Open "Select grupo,classe from supproduto where nomeProd=('" & tblEquipamentos.TextMatrix(i, 0) & "')", db, 3, 3
   '      db.Execute ("INSERT INTO suppedidodetalhe (id,grupo,classe,codProd,unidade,qtdPedida,status,valorUnitario,valorTotal,acordo) VALUES ('" & id & "','" & rs!Grupo & "','" & rs!Classe & "','" & tblEquipamentos.TextMatrix(i, 8) & "','" & tblEquipamentos.TextMatrix(i, 2) & "','" & tblEquipamentos.TextMatrix(i, 1) & "',0,'" & ValorUnit & "','" & ValorTotal & "','" & tblEquipamentos.TextMatrix(i, 9) & "')")
         pes.Open "SELECT * FROM suppedidodetalhe where id = ('" & id & "') and grupo = ('" & rs!grupo & "') and classe = ('" & rs!classe & "') and codProd = ('" & tblEquipamentos.TextMatrix(i, 8) & "')", db, 3, 3
   
         If pes.EOF Then
   
            pes.AddNew
   
         End If
         
         pes!id = id
         pes!grupo = rs!grupo
         pes!classe = rs!classe
         pes!codProd = tblEquipamentos.TextMatrix(i, 8)
         pes!Unidade = tblEquipamentos.TextMatrix(i, 2)
         pes!qtdPedida = tblEquipamentos.TextMatrix(i, 1)
         pes!Status = 0
         pes!valorUnitario = ValorUnit
         pes!ValorTotal = ValorTotal
         pes!dataEntregaProd = tblEquipamentos.TextMatrix(i, 5)
         pes!acordo = tblEquipamentos.TextMatrix(i, 9)
         pes.Update
         
         pes.Close
         rs.Close
         i = i + 1
      Loop
      db.CommitTrans
   
      If cmbFormaPagto = "Antecipado" And txtValorPgo <> Empty Then
         If txtValorPgo > 0 Then
            Resp = MsgBox("Pagamento antecipado. Gerar financeiro ?", vbExclamation + vbYesNo)
            If Resp = vbYes Then
            
               Call gerarfinanceiro
            
            End If
         
         End If
      
      End If
      
      If cmbTipoPO = "NOVA" Then
      
         cmbTipoPO = "GERADA"
         cmbNumPO = id
         
      End If
      
      FechaDB
   
   Else
      
      MsgBox ("Campos necessários não foram preenchidos"), vbInformation
      
   End If

End Sub



Private Sub Form_Load()
   
   lblHoje = Date
   dtDataEntregaProd = Date
   
   cmbFrete.AddItem "CIF"
   cmbFrete.AddItem "FOB"
   cmbFrete.AddItem "AEREO"
   cmbFrete.AddItem "MARITIMO"
   cmbFrete.AddItem "RODOVIARIO"
   cmbFrete.AddItem "RETIRADA NO FORNECEDOR"
   
   lblHoje = Date

   dtDataPrevista = Date

   cmbTipoPO.AddItem "NOVA"
   cmbTipoPO.AddItem "GERADA"
   cmbTipoPO.AddItem "EMITIDA"

   cmbTipoPO.ListIndex = 0

   Call Rotina_AbrirBanco

   rs.Open "Select descricao from supgrupoclasse where classe = 0", db, 3, 3

   If rs.EOF Then

      MsgBox ("Não existem grupo registrados")
      FechaDB
      Exit Sub
   
   End If

   rs.MoveFirst

   Do While Not rs.EOF

      cmbGrupo.AddItem rs!Descricao
      rs.MoveNext

   Loop

   rs.Close

   pes.Open "Select chPessoa from Pessoa where pesTipoPessoa=2", db, 3, 3

   If pes.EOF Then

      MsgBox ("Não existem fornecedores registrados")
      FechaDB
      Exit Sub
   
   End If

   pes.MoveFirst

   Do While Not pes.EOF

      cmbFornecedor.AddItem pes!chPessoa
      pes.MoveNext

   Loop
   
   pes.Close
   
   rs.Open "Select apelido from supendereco", db, 3, 3

   If rs.EOF Then

      MsgBox ("Não existem endereços")
      FechaDB
      Exit Sub
   
   End If

   rs.MoveFirst

   Do While Not rs.EOF

      cmbEndEntrega.AddItem rs!apelido
      rs.MoveNext

   Loop

   rs.Close

   cmbFormaPagto.AddItem "À Vista"
   cmbFormaPagto.AddItem "Parcelado"
   cmbFormaPagto.AddItem "Antecipado"
   
   pes.Open "SELECT * FROM TipoLancamento", db, 3, 3
   If pes.EOF Then
      MsgBox ("Tipo de Lançamento vazio"), vbCritical
      FechaDB
      Exit Sub
   End If
   
   pes.MoveFirst
   
   Do While Not pes.EOF
      cmbMetodoPagto.AddItem pes!chTipoDocumento
      pes.MoveNext
   Loop
   
   pes.Close
   
   txtDesconto = 0
   txtNumParcelas = 1
   
   cmbMoeda.AddItem "BRL"
   cmbMoeda.AddItem "USD"
   cmbMoeda.AddItem "EUR"
   cmbMoeda.AddItem "CNY"
   cmbMoeda.ListIndex = 0

   FechaDB

End Sub

Private Sub tblEquipamentos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     
      Linha = tblEquipamentos.Row
      cmbDescricao = tblEquipamentos.TextMatrix(Linha, 0)
      txtQtd = tblEquipamentos.TextMatrix(Linha, 1)
      txtUnid = tblEquipamentos.TextMatrix(Linha, 2)
      txtValorUnid = Format$(tblEquipamentos.TextMatrix(Linha, 3), "##,#0.00")
      txtValorTotal = Format$(tblEquipamentos.TextMatrix(Linha, 4), "##,#0.00")
      'Call Rotina_AbrirBanco
      'rs.Open "SELECT descricao FROM supgrupoclasse WHERE grupo = ('" & tblEquipamentos.TextMatrix(Linha, 5) & "') and classe = '000'", db, 3, 3
      dtDataEntregaProd = tblEquipamentos.TextMatrix(Linha, 5)
      cmbGrupo = tblEquipamentos.TextMatrix(Linha, 6)
      'rs.Close
      'rs.Open "SELECT descricao FROM supgrupoclasse WHERE grupo = ('" & tblEquipamentos.TextMatrix(Linha, 5) & "') and classe = ('" & tblEquipamentos.TextMatrix(Linha, 6) & "')", db, 3, 3
      cmbClasse = tblEquipamentos.TextMatrix(Linha, 7)
      'rs.Close
      cmbAcordo = tblEquipamentos.TextMatrix(Linha, 9)

      'FechaDB
End Sub


Public Sub limparCamposDetalhe()
   cmbDescricao = Empty
   txtQtd = Empty
   txtUnid = Empty
   txtValorUnid = Empty
   txtValorTotal = Empty
   
End Sub


Private Sub txtPercPagMoeda_LostFocus()
   If txtPercPagMoeda <> Empty Then
      
      txtPago = Format$(txtTotal * (txtPercPagMoeda / 100), "##,##0.00")
      txtSaldo = Format$(txtTotal - txtPago, "##,##0.00")
   
   Else
   
      txtPago = Format$(0, "##,##0.00")
      txtSaldo = Format$(txtTotal, "##,##0.00")
   
   End If
End Sub

Private Sub txtQtd_LostFocus()
   Dim qtd As Integer
      
   If cmbDescricao <> Empty Then
      
      qtd = verificaEstoque(cmbDescricao)
      
      If txtQtd > qtd Then
      
         MsgBox ("Valor informado maior que estoque máximo: " & qtd), vbInformation
      
      End If
   
   End If
End Sub

Private Sub txtValorUnid_LostFocus()
On Error GoTo Erro:
   
   txtValorTotal = Format$(txtValorUnid * txtQtd, "##,##0.00")

Exit Sub

Erro:    MsgBox ("Valor Inválido!")

End Sub

Public Sub gerarfinanceiro()
   Call Rotina_AbrirBanco
   Dim id As String
   Dim i As Integer
   Dim codigo As String
   Dim grupo As String
   Dim classe As String
   
   Prod.Open "SELECT id FROM suppedidodecompra ORDER BY id DESC LIMIT 1", db, 3, 3
   id = "PO-" & Prod!id
   Prod.Close
   
   rs.Open "Select * from NotaFiscalEntrada where chPessoa=('" & cmbFornecedor & "') and chNotaFiscalEntrada=('" & id & "')", db, 3, 3
   
   If rs.EOF Then
   
      rs.AddNew
   
   End If
   
   rs!chPessoa = cmbFornecedor
   rs!chNotaFiscalEntrada = id
   rs!nfeFinalidadePagto = 2
   rs!nfeDataEmissao = Date
   rs!nfedataLanc = Date
   rs!nfeValorDaNota = txtTotal
   rs!nfeValorFrete = 0
   rs!nfePagtoFrete = 0
   rs!nfeValorICMS = 0
   rs!nfeValorIPI = 0
   rs!nfeNF_Boleto = 3
   'rs!nfeDesdobramento
   rs!nfeTipoLancamento = 11
   rs!nfeStatus = 1
   rs.Update
   
   rs.Close
   
   i = 1
   
   Do While i < tblEquipamentos.Rows
      Prod.Open "Select grupo,classe from supgrupoclasse where descricao = ('" & tblEquipamentos.TextMatrix(i, 7) & "')", db, 3, 3
         grupo = Prod!grupo
         classe = Prod!classe
         codigo = grupo & classe & tblEquipamentos.TextMatrix(i, 8)
      Prod.Close
      rs.Open "Select * from notaFiscalDetProd where chPessoa=('" & cmbFornecedor & "') and chNotaFiscalEntrada=('" & id & "') and chCodProduto=('" & codigo & "')", db, 3, 3
      
      If rs.EOF Then
      
         rs.AddNew
      
      End If
      rs!chPessoa = cmbFornecedor
      rs!chNotaFiscalEntrada = id
      rs!chCodProduto = codigo
      'rs!chFatura = 1
      rs!nfdCentroDeCusto = "2"
      
      Prod.Open "Select GrupoCentroDeCusto,SubGrupoCentroDeCusto from supProduto where grupo = ('" & grupo & "') and classe = ('" & classe & "') and codProd=('" & tblEquipamentos.TextMatrix(i, 8) & "')", db, 3, 3
         rs!nfdGrupoCentroDeCusto = Prod!GrupoCentroDeCusto
         rs!nfdSubGrupoCentroDeCusto = Prod!SubGrupoCentroDeCusto
      Prod.Close
      
      rs!nfdQtd = tblEquipamentos.TextMatrix(i, 1)
      rs!nfdPU = tblEquipamentos.TextMatrix(i, 3)
      rs!nfdValorDaCompra = tblEquipamentos.TextMatrix(i, 4)
      'rs!nfdQtdParcelas
      'rs!nfdValorDaParcela
      rs!nfdStatusPagto = 1
      rs.Update
      i = i + 1
      rs.Close
   Loop
   
   rs.Open "Select * from Contas_A_Pagar where chPessoa=('" & cmbFornecedor & "') and chNotaFiscal=('" & id & "')", db, 3, 3
   
   If rs.EOF Then
   
      rs.AddNew
   
   End If
   
   rs!chFabricante = 0
   rs!chPessoa = cmbFornecedor
   rs!chNotafiscal = id
   rs!chFatura = id
   rs!chDataVencito = Date
   rs!ctpDataEmissao = Date
   rs!ctpdatabanco = Date
   rs!ctpDataLanc = Date
   rs!ctpDataVencOriginal = Date
   rs!ctpdescricaooperacao = "Pagamento Antecipado"
   rs!ctpValorLart = 0
   rs!ctpValorMerco = 0
   rs!ctpValorDaBoleta = txtValorPgo
   rs!chAno = Year(Date)
   rs!chMes = Month(Date)
   rs!chDia = Day(Date)
   rs!chCodBcoLart = "ITAU"
   rs!ctpStatus = 1
   rs!ctpDataProc = Date
   rs!ctpDataPagamento = Date
   'Alterar ctpTipoLancamento
   rs!ctpTipoLancamento = cmbMetodoPagto.ListIndex
   rs!ctpTipoLancamentoDesc = cmbMetodoPagto
   rs!ctpPessoaReembolso = "SHB BRASIL"
   rs.Update
   rs.Close
   MsgBox ("Financeiro gerado com sucesso."), vbInformation
   FechaDB
End Sub

Public Sub calculaTotal()
   Dim total As Double
   Dim i As Integer
   txtTotal = Empty
   total = 0
   i = 1
   
   Do While i < tblEquipamentos.Rows
      If tblEquipamentos.TextMatrix(i, 4) <> Empty Then
         total = total + tblEquipamentos.TextMatrix(i, 4)
      End If
      i = i + 1
   Loop
   
   txtDesc = Format$(total * (txtDesconto / 100), "##,##0.00")
   
   txtTotal = Format$(total * (1 - txtDesconto / 100), "##,##0.00")

End Sub

Public Function verificaEstoque(Produto As String) As Integer
   Dim quantidade As Integer
   
   Call Rotina_AbrirBanco
      
   rs.Open "SELECT estoqueMaximo FROM supEstoque INNER JOIN supProduto ON supProduto.grupo = supEstoque.grupo AND supProduto.classe = supEstoque.classe AND supProduto.codProd = supEstoque.codProd WHERE nomeProd = ('" & Produto & "')", db, 3, 3
   quantidade = rs!estoqueMaximo
   rs.Close
   verificaEstoque = quantidade
   
End Function
