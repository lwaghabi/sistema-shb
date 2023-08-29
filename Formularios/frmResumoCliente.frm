VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmResumoCliente 
   ClientHeight    =   8085
   ClientLeft      =   9630
   ClientTop       =   7935
   ClientWidth     =   12105
   LinkTopic       =   "Form2"
   ScaleHeight     =   8085
   ScaleWidth      =   12105
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCliente 
      BackColor       =   &H00FFFFC0&
      Height          =   5715
      Left            =   120
      TabIndex        =   141
      Top             =   120
      Width           =   1695
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
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Frame Frame30 
      Caption         =   "Pesquisa"
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
      TabIndex        =   23
      Top             =   6120
      Width           =   1695
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FFFF00&
         Caption         =   "OK"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtConsulta 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   12632256
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Dados Cadastrais"
      TabPicture(0)   =   "frmResumoCliente.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FraContato"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame28"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame27"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fr1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Pedidos no Mes Atual"
      TabPicture(1)   =   "frmResumoCliente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame15"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame13"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblDataCadastro1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblCodCliente1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Histórico de Vendas"
      TabPicture(2)   =   "frmResumoCliente.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame18"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame12"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame19"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame17"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame14"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame24"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame10"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Historico Faturamento"
      TabPicture(3)   =   "frmResumoCliente.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame25"
      Tab(3).Control(1)=   "Frame23"
      Tab(3).Control(2)=   "Frame16"
      Tab(3).Control(3)=   "Frame3"
      Tab(3).Control(4)=   "Frame22"
      Tab(3).Control(5)=   "Frame21"
      Tab(3).Control(6)=   "Frame20"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Mostruários"
      TabPicture(4)   =   "frmResumoCliente.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "GridMostruario"
      Tab(4).Control(1)=   "cmdConsultaMostruario"
      Tab(4).Control(2)=   "Frame26"
      Tab(4).Control(3)=   "Frame11"
      Tab(4).Control(4)=   "Frame9"
      Tab(4).ControlCount=   5
      Begin MSFlexGridLib.MSFlexGrid GridMostruario 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   138
         Top             =   2400
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   9340
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorBkg    =   16777152
         FormatString    =   "N Pedido|Cp|Data Neg|Fatura.|Interv|A Partir"
      End
      Begin VB.Frame FraContato 
         Caption         =   "Contato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   132
         Top             =   3360
         Width           =   9495
         Begin MSFlexGridLib.MSFlexGrid GridContato 
            Height          =   1815
            Left            =   120
            TabIndex        =   133
            Top             =   240
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   3201
            _Version        =   393216
            FixedCols       =   0
            BackColor       =   16777152
            BackColorFixed  =   16776960
            BackColorBkg    =   16777152
            FormatString    =   $"frmResumoCliente.frx":008C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Status do Cliente"
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
         Left            =   -67320
         TabIndex        =   130
         Top             =   600
         Width           =   2055
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Height          =   285
            Left            =   120
            TabIndex        =   131
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdConsultaMostruario 
         BackColor       =   &H00FFFF00&
         Caption         =   "Consulta"
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
         Left            =   -71760
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Frame Frame26 
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
         Left            =   -74760
         TabIndex        =   126
         Top             =   1440
         Width           =   2895
         Begin VB.ComboBox cmbFiltro 
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
            Left            =   120
            TabIndex        =   127
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Cliente"
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
         Left            =   -74760
         TabIndex        =   124
         Top             =   600
         Width           =   2895
         Begin VB.Label lblCodCliente4 
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   125
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Data Cadastro"
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
         Left            =   -71880
         TabIndex        =   122
         Top             =   600
         Width           =   1455
         Begin VB.Label lblDataCadastro4 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   123
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Formas de pagamento por Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   240
         TabIndex        =   120
         Top             =   3840
         Width           =   4935
         Begin MSFlexGridLib.MSFlexGrid GridFormaPagto 
            Height          =   3495
            Left            =   120
            TabIndex        =   136
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   6165
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            BackColor       =   16777152
            BackColorFixed  =   16776960
            BackColorBkg    =   16777152
            FormatString    =   "N Pedido|Cp|Data Neg|Fatura.|Interv|A Partir|Desc.| "
         End
      End
      Begin VB.Frame Frame25 
         Height          =   735
         Left            =   -66960
         TabIndex        =   91
         Top             =   300
         Width           =   1815
         Begin VB.CommandButton cmdConsultaHistFat 
            BackColor       =   &H00FFFF00&
            Caption         =   "Consulta"
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
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame24 
         Height          =   735
         Left            =   7800
         TabIndex        =   89
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton cmdConsulta 
            BackColor       =   &H00FFFF00&
            Caption         =   "Consulta"
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
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Período a Consultar"
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
         Left            =   -70080
         TabIndex        =   85
         Top             =   300
         Width           =   3015
         Begin MSComCtl2.DTPicker txtFatFim 
            Height          =   375
            Left            =   1680
            TabIndex        =   86
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            Format          =   265355265
            CurrentDate     =   38118
         End
         Begin MSComCtl2.DTPicker txtFatIni 
            Height          =   375
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            Format          =   265355265
            CurrentDate     =   38118
         End
         Begin VB.Label Label50 
            Caption         =   "A"
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
            Left            =   1440
            TabIndex        =   88
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Resumo do Historico de Faturamento"
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
         Left            =   -74880
         TabIndex        =   84
         Top             =   3900
         Width           =   9735
         Begin MSMask.MaskEdBox txtDataMaiorFat 
            Height          =   255
            Left            =   2280
            TabIndex        =   97
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
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
         Begin VB.Label Label38 
            Caption         =   "Descrição"
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
            Left            =   4680
            TabIndex        =   113
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            Caption         =   "Qtd."
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
            Left            =   6960
            TabIndex        =   112
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor"
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
            Left            =   8160
            TabIndex        =   111
            Top             =   240
            Width           =   975
         End
         Begin VB.Label txtFatMensal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3600
            TabIndex        =   110
            Top             =   480
            Width           =   975
         End
         Begin VB.Label txtMesFat 
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   109
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label51 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Maior Fat. no Período "
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
            TabIndex        =   108
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label62 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor"
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
            Left            =   3600
            TabIndex        =   106
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label61 
            Alignment       =   2  'Center
            Caption         =   "Data/Qtd."
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
            Left            =   2760
            TabIndex        =   105
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label60 
            Caption         =   "Descrição"
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
            TabIndex        =   104
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label txtValorAtrsoPeriodo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   7800
            TabIndex        =   103
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label txtMairoQtdDiasAtraso 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
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
            Left            =   6960
            TabIndex        =   102
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label57 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Maior Atraso em Dias"
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
            Left            =   4680
            TabIndex        =   101
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label txtQtdPagtoAtrasoPeriodo 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
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
            Left            =   6960
            TabIndex        =   100
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label54 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Qtd. c/Atraso no Período"
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
            Left            =   4680
            TabIndex        =   99
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label txtMaiorFatPeriodo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3600
            TabIndex        =   98
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label46 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Maior Boleta no Período"
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
            TabIndex        =   96
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Histórico de Faturamento"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   83
         Top             =   1020
         Width           =   9735
         Begin MSFlexGridLib.MSFlexGrid GridHistFat 
            Height          =   2415
            Left            =   120
            TabIndex        =   140
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   4260
            _Version        =   393216
            Cols            =   11
            FixedCols       =   0
            BackColor       =   16777152
            BackColorFixed  =   16776960
            BackColorBkg    =   16777152
            FormatString    =   "| Data Emissão|Vencimento|Data Receb| Nota Fiscal|Valor Operação|Correção|Valor total|Hist                  |Status  |        "
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Faturamentos Pendentes e Pagos no Mês"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -74880
         TabIndex        =   82
         Top             =   5280
         Width           =   9735
         Begin MSFlexGridLib.MSFlexGrid GridFaturamento 
            Height          =   2055
            Left            =   120
            TabIndex        =   139
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   3625
            _Version        =   393216
            Cols            =   10
            FixedCols       =   0
            BackColor       =   16777152
            BackColorFixed  =   16776960
            BackColorBkg    =   16777152
            FormatString    =   "Nota Fiscal |Data Emissão|Vencimento|Num Pedido|Comp|Descrição                |Banco|Valor Operação|Valor total|Status  "
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Data de Cadastro"
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
         Left            =   -72120
         TabIndex        =   69
         Top             =   420
         Width           =   1815
         Begin VB.Label lblDataCadastro3 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   71
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Cliente"
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
         Left            =   -74880
         TabIndex        =   68
         Top             =   420
         Width           =   2775
         Begin VB.Label lblCodCliente3 
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   70
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Resumo Vendas do Período Por Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   5280
         TabIndex        =   64
         Top             =   3840
         Width           =   4695
         Begin MSFlexGridLib.MSFlexGrid GridConsProd 
            Height          =   2295
            Left            =   240
            TabIndex        =   137
            Top             =   1440
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   4048
            _Version        =   393216
            Cols            =   4
            BackColor       =   16777152
            BackColorFixed  =   16776960
            BackColorBkg    =   16777152
            FormatString    =   "|Data Neg       |Qtd Neg            |Preço Unit.         "
         End
         Begin VB.CommandButton cmdConsultaProduto 
            BackColor       =   &H00FFFF00&
            Caption         =   "Consulta Produto"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   840
            Width           =   975
         End
         Begin VB.ComboBox cmbProdutoVendas 
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
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Total"
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
            Left            =   2160
            TabIndex        =   121
            Top             =   600
            Width           =   450
         End
         Begin VB.Label txtTotalMetroProduto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2670
            TabIndex        =   95
            Top             =   600
            Width           =   915
         End
         Begin VB.Label txtDescProd 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1560
            TabIndex        =   81
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label35 
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
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Entregue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   -74880
         TabIndex        =   63
         Top             =   4200
         Width           =   9735
         Begin MSFlexGridLib.MSFlexGrid GridProc 
            Height          =   3015
            Left            =   120
            TabIndex        =   142
            Top             =   360
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   5318
            _Version        =   393216
            Cols            =   13
            FixedCols       =   0
            BackColor       =   16777152
            BackColorFixed  =   16776960
            BackColorBkg    =   16777152
            FormatString    =   $"frmResumoCliente.frx":0117
         End
         Begin VB.Label Label36 
            Caption         =   "Total Entregue"
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
            Left            =   6840
            TabIndex        =   78
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label txtTotalProc 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   8180
            TabIndex        =   77
            Top             =   120
            Width           =   1170
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Período a Consultar"
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
         Left            =   4680
         TabIndex        =   60
         Top             =   360
         Width           =   3015
         Begin MSComCtl2.DTPicker txtDataFinal 
            Height          =   375
            Left            =   1680
            TabIndex        =   61
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            Format          =   265158657
            CurrentDate     =   38118
         End
         Begin MSComCtl2.DTPicker txtDataInicio 
            Height          =   375
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            Format          =   265158657
            CurrentDate     =   38118
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "A"
            Height          =   195
            Left            =   1440
            TabIndex        =   107
            Top             =   360
            Width           =   135
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Data Cadastro"
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
         TabIndex        =   58
         Top             =   360
         Width           =   1455
         Begin VB.Label lblDataCadastro2 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   59
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Produtos Consumidos do Período"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   5280
         TabIndex        =   56
         Top             =   1140
         Width           =   4695
         Begin MSFlexGridLib.MSFlexGrid GridConsumidos 
            Height          =   1815
            Left            =   120
            TabIndex        =   135
            Top             =   360
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   3201
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            BackColor       =   16777152
            BackColorFixed  =   16776960
            BackColorBkg    =   16777152
            FormatString    =   "Cod Prod       |Descrição                    | Quantd.   |%  "
         End
         Begin VB.Label txtQtdTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3000
            TabIndex        =   94
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label txtQtdaTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Metro"
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
            Left            =   1440
            TabIndex        =   72
            Top             =   2280
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Resumo Vendas do Período"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   240
         TabIndex        =   55
         Top             =   1140
         Width           =   4935
         Begin VB.PictureBox GridResumoVendas 
            Height          =   2055
            Left            =   120
            ScaleHeight     =   1995
            ScaleWidth      =   2955
            TabIndex        =   79
            Top             =   240
            Width           =   3015
            Begin MSFlexGridLib.MSFlexGrid GridResVendas 
               Height          =   1695
               Left            =   120
               TabIndex        =   134
               Top             =   120
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   2990
               _Version        =   393216
               Cols            =   3
               FixedCols       =   0
               BackColor       =   16777152
               BackColorFixed  =   16776960
               BackColorBkg    =   16777152
               FormatString    =   "| Data                |Qtd                 "
            End
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Menor Venda"
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
            Left            =   3240
            TabIndex        =   119
            Top             =   1320
            Width           =   1140
         End
         Begin VB.Label lblQuantidadeMenorVenda 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3840
            TabIndex        =   118
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lblDataMenorVenda 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3240
            TabIndex        =   117
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lblQuantidadeMaiorVenda 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3840
            TabIndex        =   116
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblDataMaiorVenda 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3240
            TabIndex        =   115
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Maior Venda"
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
            Left            =   3240
            TabIndex        =   114
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label txtTotalMetrosPeriodo 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   1680
            TabIndex        =   80
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Total Metro"
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
            Left            =   600
            TabIndex        =   73
            Top             =   2280
            Width           =   990
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Cliente"
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
         Left            =   240
         TabIndex        =   54
         Top             =   360
         Width           =   2895
         Begin VB.Label lblCodCliente2 
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   57
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame28 
         Caption         =   "Data de Cadastro"
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
         Left            =   -74760
         TabIndex        =   29
         Top             =   1320
         Width           =   1695
         Begin VB.Label lblDataCadastro 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   30
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame27 
         Caption         =   "Código do Cliente"
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
         Left            =   -74760
         TabIndex        =   27
         Top             =   600
         Width           =   2415
         Begin VB.Label lblCodCliente 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   28
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Em Carteira"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   -74880
         TabIndex        =   22
         Top             =   720
         Width           =   9735
         Begin MSFlexGridLib.MSFlexGrid GridPedidos 
            Height          =   3015
            Left            =   120
            TabIndex        =   143
            Top             =   360
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   5318
            _Version        =   393216
            Cols            =   13
            FixedCols       =   0
            BackColor       =   16777152
            BackColorFixed  =   16776960
            BackColorBkg    =   16777152
            FormatString    =   "|Data Pedido|Pedido|Cp|Cliente          | Descrição              |Un|Qtd.  |   P.U.  |Diaria    |Qtd Dias|Valor Neg  |Total Pedido"
         End
         Begin VB.Label txtTotalPend 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   8160
            TabIndex        =   76
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label Label37 
            Caption         =   " Total em Carteira"
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
            Left            =   6480
            TabIndex        =   75
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Promotores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2175
         Left            =   -69960
         TabIndex        =   16
         Top             =   5580
         Width           =   4695
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "E-mail"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label lblEmailPromot 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   51
            Top             =   1680
            Width           =   4455
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
            Height          =   195
            Left            =   1800
            TabIndex        =   48
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Celular"
            Height          =   195
            Left            =   3120
            TabIndex        =   47
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Telefone"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   840
            Width           =   630
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Promotor(a)"
            Height          =   195
            Left            =   1320
            TabIndex        =   45
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Carteira"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblFaxPromot 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1800
            TabIndex        =   26
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblPromotor 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   21
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblCelPromot 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3120
            TabIndex        =   19
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblTelPromot 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   18
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblCartPromot 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   17
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame fr1 
         Caption         =   "Representantes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   12
         Top             =   5580
         Width           =   4695
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "E-mail"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label lblEmailRep 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   49
            Top             =   1680
            Width           =   4455
         End
         Begin VB.Label lblRepresentante 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   43
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
            Height          =   195
            Left            =   1680
            TabIndex        =   42
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label19 
            Caption         =   "Celular"
            Height          =   255
            Left            =   3120
            TabIndex        =   41
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Telefone"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   840
            Width           =   630
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Representante"
            Height          =   195
            Left            =   1320
            TabIndex        =   39
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Carteira"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblTelRep 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblCelRep 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3120
            TabIndex        =   15
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblFaxRep 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1680
            TabIndex        =   14
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblCartRep 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   13
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "CNPJ ------------------------------ Inscrição Estadual"
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
         Left            =   -73080
         TabIndex        =   9
         Top             =   1320
         Width           =   4455
         Begin VB.Label lblInscEstadual 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   11
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lblCNPJ 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   10
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   1455
         Left            =   -74760
         TabIndex        =   3
         Top             =   1920
         Width           =   9495
         Begin VB.Label lblCelCliente 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   4320
            TabIndex        =   129
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Região"
            Height          =   195
            Left            =   5640
            TabIndex        =   37
            Top             =   840
            Width           =   510
         End
         Begin VB.Label lblRegiao 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   5640
            TabIndex        =   36
            Top             =   1080
            Width           =   3735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            Height          =   195
            Left            =   4200
            TabIndex        =   35
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   3720
            TabIndex        =   34
            Top             =   840
            Width           =   210
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   6120
            TabIndex        =   32
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   810
         End
         Begin VB.Label lblCEP 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   4200
            TabIndex        =   8
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblEstado 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3600
            TabIndex        =   7
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblCidade 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   6
            Top             =   1080
            Width           =   3375
         End
         Begin VB.Label lblBairro 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   6120
            TabIndex        =   5
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label lblEndereco 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   4
            Top             =   480
            Width           =   5895
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Razão Social"
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
         Left            =   -72240
         TabIndex        =   1
         Top             =   600
         Width           =   4815
         Begin VB.Label lblRazaoSocial 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Top             =   240
            Width           =   4575
         End
      End
      Begin VB.Label lblDataCadastro1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -71880
         TabIndex        =   93
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label lblCodCliente1 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -74880
         TabIndex        =   74
         Top             =   420
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmResumoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim codprodtab(150) As String
Dim DescProdTab(150) As String
Dim QtdProdTab(150) As Currency
Dim Status(5) As String
Dim Pessoa As String
Dim TipoContato As String
Dim Indice As Integer
Dim Ind As Integer
Dim IndContato As Integer
Dim IndNeg As Single
Dim DataInicial As Date
Dim diainicio As Integer
Dim mesinicio As Integer
Dim anoinicio As Integer
Dim DataFinal As Date
Dim diafinal As Integer
Dim mesfinal As Integer
Dim anofinal As Integer
Dim DataConsulta1 As Date
Dim DataConsulta2 As Date
Dim DataIniTab As Date
Dim DataFimTab As Date
Dim DataTabela As Date
Dim DiaTab As Integer
Dim Mestab As Integer
Dim AnoTab As Integer
Dim Sql As String
Dim DataInicioStr As String
Dim DataFimStr As String
Dim IndProc As Integer
Dim IndPend As Integer
Dim Data_Inv As String
Dim Dia As Integer
Dim Mes As Integer
Dim Ano As Integer
Dim PedidoAnterior As String
Dim PedidoCompAnterior As String
Dim QtdPedido As Integer
Dim IndSalvo As Integer
Dim Qtd As Integer
Dim AcumulaPend As Currency
Dim AcumulaProc As Currency
Dim SalvaPessoa As String
Dim fim As Byte
Dim fim1 As Byte
Dim Encontrei As Byte
Dim MesAnoTab(240) As String
Dim QtdTab(240) As Single
Dim DataArq As Date
Dim MesAnoArq As String
Dim MesMenorQtd As String
Dim MenorQtd As Currency
Dim MesMaiorQtd As String
Dim MaiorQtd As Currency
Dim AcumulaQtd As Currency
Dim ValorPedido As Currency
Dim DataUtil As Date
Dim NumDiasMaior As Integer
Dim QtdAtrasos As Integer
Dim DataMenorValor As Date
Dim MenorValor As Currency
Dim DataMaiorValor As Date
Dim MaiorValor As Currency
Dim ValorMaiorQtdDias As Currency
Dim PeriodoAnt As String
Dim PeriodoMaior As String
Dim FaturaMensal As Currency
Dim AcumulaMensal As Currency
Dim ClienteAnterior As String
Dim OrdemAnterior As Integer
Dim Coluna As Integer
Dim Linha As Integer
Dim CaminhoImagem As String
Dim ExtensaoImagem As String
Dim FSys As New Scripting.FileSystemObject


Private Sub cmbProdutoVendas_LostFocus()
Call Rotina_AbrirBanco

Prod.Open "Select * from Produto where chProduto - ('" & cmbProdutoVendas & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("Informar o produto a ser pesquisado"), vbInformation
   Call FechaDB
   cmdSair.SetFocus
Else
   txtDescProd = Prod!prdNomeProd
End If

Call FechaDB

End Sub

Private Sub cmdConsulta_Click()

If lblCodCliente = Empty Then
   MsgBox ("Cliente não informado")
   cmdSair.SetFocus
   Exit Sub
End If
   
MesMenorQtd = Empty
MenorQtd = 99999.99
MaiorQtd = 0
AcumulaQtd = 0

DataConsulta1 = txtDataInicio

DataConsulta2 = txtDataFinal

DataIniTab = txtDataInicio
DataFimTab = txtDataFinal

DataIniTab = 1 & "/" & Month(txtDataInicio) & "/" & Year(txtDataInicio)

DataFimTab = 1 & "/" & Month(DataFimTab) & "/" & Year(DataFimTab)

DataTabela = DataFimTab

For Indice = 0 To 239
    If DataTabela < DataIniTab Then
       MesAnoTab(Indice) = Empty
    Else
       MesAnoTab(Indice) = Format$(DataTabela, "mmmm/yyyy")
    End If
    QtdTab(Indice) = 0
    Mestab = Month(DataTabela)
    AnoTab = Year(DataTabela)
    Mestab = Mestab - 1
    If Mestab = 0 Then
       Mestab = 12
       AnoTab = AnoTab - 1
    End If
    DataTabela = Day(DataTabela) & "/" & Mestab & "/" & AnoTab
    
    If Indice < 100 Then
       QtdProdTab(Indice) = 0
    End If
Next

fim = 0
Encontrei = 0
IndNeg = 0

Call Rotina_AbrirBanco

hneg.Open "Select * from HistoricoNegociacao", db, 3, 3
If hneg.EOF Then
   Call FechaDB
   Exit Sub
End If

    hneg.MoveFirst
    
    Do While Not hneg.EOF
       If (hneg!chDataNegociacao - 1) > DataIniTab And hneg!chDataNegociacao < DataFimTab Then

           DataArq = hneg!chDataNegociacao
           DataArq = 1 & "/" & Month(DataArq) & "/" & Year(DataArq)
           MesAnoArq = Format$(DataArq, "mmmm/yyyy")
           
           If Prod.State = 1 Then
              Prod.Close: Set Prod = Nothing
           End If
            
           Prod.Open "Select * from Produto", db, 3, 3
           If Prod.EOF Then
              Call FechaDB
              Exit Sub
           End If
           
           Prod.MoveFirst
           Encontrei = 0
    
           Do While Encontrei = 0
              hdneg.Open "Select * from HistoricoDetalheNegociacao where chPessoa = ('" & lblCodCliente & "') and chnumpedido = ('" & hneg!chNumPedido & "') and chNumPedidoComp = ('" & hneg!chNumPedidocomp & "') and chproduto = ('" & Prod!chproduto & "')", db, 3, 3
              If hdneg.EOF Then
                 Prod.MoveNext
                 If Prod.EOF Then
                    Prod.Close: Set Prod = Nothing
                    Encontrei = 2
                 Else
                    hdneg.Close: Set hdneg = Nothing
                 End If
              Else
                 Encontrei = 1
              End If
           Loop
            
           If Encontrei = 1 Then
              fim = 0
              Do While fim = 0
                 If hneg!chPessoa = hdneg!chPessoa Then
                    If hneg!chNumPedido = hdneg!chNumPedido Then
                       If hneg!chNumPedidocomp = hdneg!chNumPedidocomp Then
                          Call Rotina_GridResumoVendas
                          Call Rotina_GridConsumidos
                          hdneg.MoveNext
                          If hdneg.EOF Then
                             fim = 1
                          Else
                             fim = 0
                          End If
                       Else
                          fim = 1
                       End If
                    Else
                       fim = 1
                    End If
                 Else
                    fim = 1
                 End If
              Loop
           End If
       Else
          hneg.MoveNext
       End If
   
    Loop
    
'    Ano = Year(TabHistoricoNegociacao("chdatanegociacao"))
'    If TabHistoricoNegociacao("chpessoa") = lblCodCliente And Ano > Year(Date) - 2 Then
'       Encontrei = 0
'       tabproduto.MoveFirst
'       Do While Encontrei = 0
'          TabHistoricoDetNeg.Seek "=", TabHistoricoNegociacao("chpessoa"), TabHistoricoNegociacao("chnumpedido"), TabHistoricoNegociacao("chnumpedidocomp"), tabproduto("chproduto")
'          If TabHistoricoDetNeg.NoMatch Then
'             tabproduto.MoveNext
'             If tabproduto.EOF Then
'                If Encontrei = 0 Then
'                   MsgBox ("Não há produtos para este numero de pedido")
'                   Encontrei = 2
'                End If
'             End If
'          Else
'             Encontrei = 1
'          End If
'       Loop
'       Call Rotina_GridFormaPagto
'    End If
'    TabHistoricoNegociacao.MoveNext
'    If TabHistoricoNegociacao.EOF Then
'       Encontrei = 1
'    Else
'       Do While TabHistoricoNegociacao("hngstatus") = 3 And Encontrei = 0
'          TabHistoricoNegociacao.MoveNext
'          If TabHistoricoNegociacao.EOF Then
'             Encontrei = 1
'          End If
'       Loop
'    End If
'    Encontrei = 0
'    fim = 0
'    Loop
'End If
For Indice = 1 To 239
    If MesAnoTab(Indice) = Empty Then
       Indice = 239
    Else
       GridResVendas.Rows = Indice + 1
       GridResVendas.TextMatrix(Indice, 1) = MesAnoTab(Indice)
       GridResVendas.TextMatrix(Indice, 2) = Format$(QtdTab(Indice), "#,##0.00")
       AcumulaQtd = AcumulaQtd + QtdTab(Indice)
       If QtdTab(Indice) < MenorQtd Then
          MenorQtd = QtdTab(Indice)
          MesMenorQtd = MesAnoTab(Indice)
       End If
       If QtdTab(Indice) > MaiorQtd Then
          MaiorQtd = QtdTab(Indice)
          MesMaiorQtd = MesAnoTab(Indice)
       End If
    End If
Next

txtTotalMetrosPeriodo = Format$(AcumulaQtd, "#,##0.00")
lblDataMaiorVenda = MesMaiorQtd
lblQuantidadeMaiorVenda = Format$(MaiorQtd, "#,##0.00")
lblDataMenorVenda = MesMenorQtd
lblQuantidadeMenorVenda = Format$(MenorQtd, "#,##0.00")

If AcumulaQtd = 0 Then
   MsgBox "Não há informações para o período consultado"
Else
    IndProc = 0
    For Indice = 1 To 99
        If QtdProdTab(Indice) > 0 Then
           IndProc = IndProc + 1
           GridConsumidos.Rows = IndProc + 1
           GridConsumidos.TextMatrix(IndProc, 0) = codprodtab(Indice)
           GridConsumidos.TextMatrix(IndProc, 1) = DescProdTab(Indice)
           GridConsumidos.TextMatrix(IndProc, 2) = Format$(QtdProdTab(Indice), "#,##0.00")
           GridConsumidos.TextMatrix(IndProc, 3) = QtdProdTab(Indice)
        End If
    Next
    
    GridConsumidos.Col = 3
    GridConsumidos.ColSel = 3
    GridConsumidos.Row = 1
    GridConsumidos.RowSel = IndProc
    
    GridConsumidos.Sort = 2
    
    GridConsumidos.Col = 0
    GridConsumidos.ColSel = 0
    GridConsumidos.Row = 0
    GridConsumidos.RowSel = 0
    
    GridFormaPagto.Col = 7
    GridFormaPagto.ColSel = 7
    GridFormaPagto.Row = 1
    GridFormaPagto.RowSel = IndNeg
    
    GridFormaPagto.Sort = 6
    
    GridFormaPagto.Col = 0
    GridFormaPagto.ColSel = 0
    GridFormaPagto.Row = 0
    GridFormaPagto.RowSel = 0
    
    txtQtdTotal = Format$(AcumulaQtd, "#,##0.00")
End If
End Sub

Private Sub cmdConsultaHistFat_Click()
If txtFatIni = Date Then
   MsgBox ("A Consulta de Faturamento requer datas anteriores ao primeiro dia do mes atual")
   cmdSair.SetFocus
   Exit Sub
End If

If lblCodCliente3 = Empty Then
   MsgBox ("Cliente para consulta não informado")
   cmdSair.SetFocus
   Exit Sub
End If

Indice = 0
NumDiasMaior = 0
QtdAtrasos = 0
MaiorValor = 0
MenorValor = 99999999.99
ValorMaiorQtdDias = 0
PeriodoAnt = Empty
PeriodoMaior = Empty
FaturaMensal = 0
AcumulaMensal = 0

TabHistCtaReceber.MoveFirst

Do While Not TabHistCtaReceber.EOF
   If TabHistCtaReceber("chpessoa") > lblCodCliente3 Then
      TabHistCtaReceber.MoveLast
   Else
      If TabHistCtaReceber("chpessoa") = lblCodCliente3 Then
         If (TabHistCtaReceber("ctrdatavencito") > (txtFatIni - 1)) And (TabHistCtaReceber("ctrdatavencito") < (txtFatFim + 1)) Then
            
            Indice = Indice + 1
            GridHistFat.Rows = Indice + 1
            Dia = Day(TabHistCtaReceber("ctrdatavencito"))
            Mes = Month(TabHistCtaReceber("ctrdatavencito"))
            Ano = Year(TabHistCtaReceber("ctrdatavencito"))
            Data_Inv = Ano & "/" & Format$(Mes, "00") & "/" & Format$(Dia, "00")
            GridHistFat.TextMatrix(Indice, 0) = Data_Inv
            GridHistFat.TextMatrix(Indice, 1) = TabHistCtaReceber("ctrdataemissao")
            GridHistFat.TextMatrix(Indice, 2) = TabHistCtaReceber("ctrdatavencito")
            GridHistFat.TextMatrix(Indice, 3) = TabHistCtaReceber("ctrdatarecebimento")
            GridHistFat.TextMatrix(Indice, 4) = TabHistCtaReceber("chnotafiscal") & "/" & TabHistCtaReceber("chfatura")
           
            DataUtil = TabHistCtaReceber("ctrdatavencito")
            DataInformada = DataUtil
            NDias = 0
           ' DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)
            
            'DataUtil = DataRetorno.DiaUtil
              
            GridHistFat.TextMatrix(Indice, 6) = Format$((TabHistCtaReceber("ctrvalorlart") + TabHistCtaReceber("ctrvalormerco")), "##,##0.00")
            GridHistFat.TextMatrix(Indice, 7) = Format$(TabHistCtaReceber("ctrvalorcorrecao"), "##,##0.00")
            GridHistFat.TextMatrix(Indice, 8) = Format$(TabHistCtaReceber("ctrvalordaboleta"), "##,##0.00")
            If TabHistCtaReceber("ctrvalordaboleta") > MaiorValor Then
               MaiorValor = TabHistCtaReceber("ctrvalordaboleta")
               DataMaiorValor = TabHistCtaReceber("ctrdatavencito")
            End If
            If TabHistCtaReceber("ctrvalordaboleta") < MenorValor Then
               MenorValor = TabHistCtaReceber("ctrvalordaboleta")
               DataMenorValor = TabHistCtaReceber("ctrdatavencito")
            End If
            If TabHistCtaReceber("ctrdatarecebimento") > DataUtil Then
               NDias = TabHistCtaReceber("ctrdatarecebimento") - TabHistCtaReceber("ctrdatavencito")
               GridHistFat.TextMatrix(Indice, 5) = NDias
               GridHistFat.TextMatrix(Indice, 9) = "C/Atraso"
               If NDias < NumDiasMaior Then
                  NDias = NDias
               Else
                  If NDias = NumDiasMaior Then
                     If ValorMaiorQtdDias < TabHistCtaReceber("ctrvalordaboleta") Then
                        ValorMaiorQtdDias = TabHistCtaReceber("ctrvalordaboleta")
                     End If
                  Else
                     NumDiasMaior = NDias
                     ValorMaiorQtdDias = TabHistCtaReceber("ctrvalordaboleta")
                  End If
               End If
               QtdAtrasos = QtdAtrasos + 1
            Else
               NDias = 0
               If TabHistCtaReceber("ctrdatarecebimento") < DataUtil Then
                  GridHistFat.TextMatrix(Indice, 5) = NDias
                  GridHistFat.TextMatrix(Indice, 9) = "Antecipado"
               Else
                  GridHistFat.TextMatrix(Indice, 5) = NDias
                  GridHistFat.TextMatrix(Indice, 10) = "Ok"
               End If
            End If
            GridHistFat.TextMatrix(Indice, 10) = Format$(TabHistCtaReceber("ctrdatarecebimento"), "mmmm/yyyy")
         End If
      End If
   End If
   TabHistCtaReceber.MoveNext
Loop

GridHistFat.Col = 0
GridHistFat.ColSel = 0
GridHistFat.Row = 1
GridHistFat.RowSel = Indice

GridHistFat.Sort = 6

GridHistFat.Row = 0
GridHistFat.RowSel = 0

For Ind = 1 To Indice
    If PeriodoAnt = Empty Then
       PeriodoAnt = GridHistFat.TextMatrix(Ind, 10)
       PeriodoMaior = GridHistFat.TextMatrix(Ind, 10)
       FaturaMensal = GridHistFat.TextMatrix(Ind, 8)
       AcumulaMensal = GridHistFat.TextMatrix(Ind, 8)
    Else
       If GridHistFat.TextMatrix(Ind, 10) = PeriodoAnt Then
          AcumulaMensal = AcumulaMensal + GridHistFat.TextMatrix(Ind, 8)
       Else
          If AcumulaMensal > FaturaMensal Then
             FaturaMensal = AcumulaMensal
             PeriodoMaior = GridHistFat.TextMatrix(Ind - 1, 10)
          End If
          AcumulaMensal = GridHistFat.TextMatrix(Ind, 8)
          PeriodoAnt = GridHistFat.TextMatrix(Ind, 10)
       End If
    End If
Next

If AcumulaMensal > FaturaMensal Then
   FaturaMensal = AcumulaMensal
   PeriodoMaior = GridHistFat.TextMatrix(Ind - 1, 1)
End If

If MaiorValor > 0 Then
   txtDataMaiorFat = DataMaiorValor
   txtMaiorFatPeriodo = Format$(MaiorValor, "##,##0.00")
   txtMesFat = PeriodoMaior
   txtFatMensal = Format$(FaturaMensal, "##,##0.00")
   txtQtdPagtoAtrasoPeriodo = QtdAtrasos
   txtMairoQtdDiasAtraso = NumDiasMaior
   txtValorAtrsoPeriodo = Format$(ValorMaiorQtdDias, "##,##0.00")
Else
   MsgBox "Não há informações para o período consultado"
End If
End Sub

Private Sub cmdConsultaMostruario_Click()
fim = 0
Ind = 0

ClienteAnterior = Empty
OrdemAnterior = 999

GridMostruario.Rows = 2
GridMostruario.TextMatrix(1, 0) = Empty
GridMostruario.TextMatrix(1, 1) = Empty
GridMostruario.TextMatrix(1, 2) = Empty
GridMostruario.TextMatrix(1, 3) = Empty
GridMostruario.TextMatrix(1, 4) = Empty
GridMostruario.TextMatrix(1, 5) = Empty

Do While fim = 0
   TabControleMostruario.Seek "=", lblCodCliente4, Ind
   If TabControleMostruario.NoMatch Then
      Ind = Ind + 1
      If Ind > 100 Then
         MsgBox "Cliente sem informações de Mostruarios"
         fim = 1
         Exit Sub
      End If
   Else
      fim = 1
   End If
Loop

fim = 0

Do While fim = 0

    If cmbFiltro.ListIndex = 0 Then
       Rotina_Carga_Mostruario
    Else
       If cmbFiltro.ListIndex = 1 And IsNull(TabControleMostruario("mosDataRetiradaMost")) Then
          Rotina_Carga_Mostruario
       Else
          If cmbFiltro.ListIndex = 2 And Not (IsNull(TabControleMostruario("mosDataRetiradaMost"))) Then
             Rotina_Carga_Mostruario
          End If
       End If
    End If
    
    TabControleMostruario.MoveNext
    If TabControleMostruario.EOF Then
       fim = 1
    Else
       If Not (TabControleMostruario("chpessoa") = ClienteAnterior) Then
          fim = 1
       End If
    End If
Loop
End Sub

Public Sub Rotina_Carga_Mostruario()

Ind = Ind + 1

GridMostruario.Rows = Ind + 1
GridMostruario.TextMatrix(Ind, 0) = TabControleMostruario("chOrdemCadast")
If IsNull(TabControleMostruario("mosDescImagem")) Then
   GridMostruario.TextMatrix(Ind, 1) = Empty
Else
   GridMostruario.TextMatrix(Ind, 1) = TabControleMostruario("mosDescImagem")
End If

If IsNull(TabControleMostruario("mosEnderecoImagem")) Then
   GridMostruario.TextMatrix(Ind, 2) = Empty
Else
   GridMostruario.TextMatrix(Ind, 2) = TabControleMostruario("mosEnderecoImagem")
End If

GridMostruario.TextMatrix(Ind, 3) = TabControleMostruario("mosDataInstalaMost")
If IsNull(TabControleMostruario("mosDataRetiradaMost")) Then
   GridMostruario.TextMatrix(Ind, 4) = Empty
Else
   GridMostruario.TextMatrix(Ind, 4) = TabControleMostruario("mosDataRetiradaMost")
End If
tabproduto.MoveFirst
Encontrei = 0
Do While Encontrei = 0
   TabControleMostruarioDetalhe.Seek "=", lblCodCliente4, TabControleMostruario("chOrdemCadast"), tabproduto("chproduto")
   If TabControleMostruarioDetalhe.NoMatch Then
      tabproduto.MoveNext
      If tabproduto.EOF Then
         If Encontrei = 2 Then
            MsgBox ("Não há produtos para este numero de pedido")
            'txtPedido.SetFocus
            'Exit Sub
         End If
      End If
   Else
      Encontrei = 1
   End If
Loop

fim1 = 0

Do While fim1 = 0
   If TabControleMostruarioDetalhe("chpessoa") = ClienteAnterior And TabControleMostruarioDetalhe("chOrdemcadast") = OrdemAnterior Then
      GridMostruario.TextMatrix(Ind, 0) = Empty
      GridMostruario.TextMatrix(Ind, 1) = Empty
      GridMostruario.TextMatrix(Ind, 2) = Empty
      GridMostruario.TextMatrix(Ind, 3) = Empty
      GridMostruario.TextMatrix(Ind, 4) = Empty
    Else
      ClienteAnterior = TabControleMostruarioDetalhe("chpessoa")
      OrdemAnterior = TabControleMostruarioDetalhe("chOrdemcadast")
    End If
    tabproduto.Seek "=", TabControleMostruarioDetalhe("chproduto")
    GridMostruario.TextMatrix(Ind, 5) = tabproduto("prdnomeprod")
    TabControleMostruarioDetalhe.MoveNext
    If TabControleMostruarioDetalhe.EOF Then
       fim1 = 1
    Else
      If TabControleMostruarioDetalhe("chpessoa") = ClienteAnterior And TabControleMostruarioDetalhe("chOrdemcadast") = OrdemAnterior Then
         Ind = Ind + 1
         GridMostruario.Rows = Ind + 1
      Else
         fim1 = 1
      End If
    End If
Loop

End Sub
Private Sub cmdConsultaProduto_Click()

Indice = 0
AcumulaQtd = 0

If lblCodCliente = Empty Then
   MsgBox ("Informar o codigo do cliente a ser pesquisado")
   cmdSair.SetFocus
   Exit Sub
End If

TabHistoricoNegociacao.MoveFirst

Do While Not TabHistoricoNegociacao.EOF
   
    If TabHistoricoNegociacao("chpessoa") > lblCodCliente Then
       TabHistoricoNegociacao.MoveLast
    Else
       If TabHistoricoNegociacao("chdatanegociacao") > DataIniTab And TabHistoricoNegociacao("chdatanegociacao") < DataFimTab Then
                    
          If TabHistoricoNegociacao("chpessoa") = lblCodCliente Then
             TabHistoricoDetNeg.Seek "=", lblCodCliente, TabHistoricoNegociacao("chnumpedido"), TabHistoricoNegociacao("chnumpedidocomp"), cmbProdutoVendas
             If TabHistoricoDetNeg.NoMatch Then
                Encontrei = 2
             Else
                Dia = Day(TabHistoricoNegociacao("chdatanegociacao"))
                Mes = Month(TabHistoricoNegociacao("chdatanegociacao"))
                Ano = Year(TabHistoricoNegociacao("chdatanegociacao"))
                Data_Inv = Ano & "/" & Format$(Mes, "00") & "/" & Format$(Dia, "00")
                Indice = Indice + 1
                GridConsProd.Rows = Indice + 1
                GridConsProd.TextMatrix(Indice, 0) = Data_Inv
                GridConsProd.TextMatrix(Indice, 1) = TabHistoricoNegociacao("chdatanegociacao")
                GridConsProd.TextMatrix(Indice, 2) = Format$(TabHistoricoDetNeg("hdnquantidademetro"), "##,##0.00")
                GridConsProd.TextMatrix(Indice, 3) = Format$(TabHistoricoDetNeg("hdnprecometro") - TabHistoricoDetNeg("hdnvalordescONTO"), "#0.00")
                AcumulaQtd = AcumulaQtd + TabHistoricoDetNeg("hdnquantidademetro")
             End If
          End If
       End If
    End If
    TabHistoricoNegociacao.MoveNext
Loop
GridConsProd.Col = 0
GridConsProd.ColSel = 0
GridConsProd.Row = 1
GridConsProd.RowSel = Indice

GridConsProd.Sort = 6

GridConsProd.ColSel = 0
GridConsProd.Row = 0

txtTotalMetroProduto = Format$(AcumulaQtd, "#,##0.00")

End Sub



Private Sub cmdOk_Click()
Dim Ind As Integer
Dim IndSalvo As Integer
For Ind = lstCliente.ListCount - 1 To 0 Step -1
       If lstCliente.List(Ind) < txtConsulta Then
          If Ind + 1 = lstCliente.ListCount Then
             IndSalvo = Ind
             lstCliente.ListIndex = Ind
             Exit For
          Else
             lstCliente.ListIndex = Ind + 1
             IndSalvo = Ind + 1
             Ind = 0
          End If
       End If
Next
   txtConsulta.SetFocus
   lblCodCliente = lstCliente.List(IndSalvo)
   lblCodCliente1 = lstCliente.List(IndSalvo)
   Tabpessoa.Seek "=", lblCodCliente

Call Rotina_Limpa_ResumoCliente

Call Rotina_Carrega_ResumoCliente

Call Rotina_Carga_Contato
End Sub

Private Sub cmdSair_Click()
 Unload Me
End Sub

Private Sub Form_Load()

Status(0) = "Ativo"
Status(1) = "Em Atraso"
Status(2) = "Indesejável"
Status(3) = "Inativo"
Status(4) = "Encerrado"

Call Rotina_AbrirBanco

pes.Open "Select * from Pessoa", db, 3, 3
If pes.EOF Then
   MsgBox ("Erro. Não há cadastro de Pessoa."), vbCritical
   Call FechaDB
   Exit Sub
End If
   
pes.MoveFirst
Do While Not pes.EOF
   If pes!pestipopessoa = 0 Then
      lstCliente.AddItem pes!chPessoa
      pes.MoveNext
   Else
      pes.MoveNext
   End If
Loop

cmbFiltro.AddItem "Ativos e Inativos"
cmbFiltro.AddItem "Ativos"
cmbFiltro.AddItem "Inativos"
cmbFiltro.ListIndex = 0
'MsgBox "Depois de ativar os combos"
Indice = 0

Prod.Open "Select * from Produto", db, 3, 3

Prod.MoveFirst

Do While Not Prod.EOF
   If Prod!prdfabricante < 9 Then
      Indice = Indice + 1
      codprodtab(Indice) = Prod!chproduto
      DescProdTab(Indice) = Prod!prdNomeProd
      cmbProdutoVendas.AddItem Prod!chproduto
      QtdProdTab(Indice) = 0
   End If
   Prod.MoveNext
Loop
'MsgBox "Depois da carga do List box"
txtDataInicio = Date
txtDataFinal = Date
txtFatIni = Date
txtFatFim = Date
'MsgBox "Final"

Call FechaDB

End Sub

Private Sub GridPedidosCarteira_Click()

End Sub


Private Sub GridMostruario_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Coluna = GridMostruario.Col
Linha = GridMostruario.Row
If GridMostruario.TextMatrix(Linha, 0) = Empty Then
   MsgBox "Clicar na linha com conteúdo em ORDEM"
   cmdSair.SetFocus
   Exit Sub
End If

'TabControleMostruario.Seek "=", lblCodCliente4, GridMostruario.TextMatrix(Linha, 0)
'If TabControleMostruario.NoMatch Then
'   MsgBox "Registro não encontrado. Informar ao analista responsável"
'   cmdSair.SetFocus
'   Exit Sub
'End If

'If TabControleMostruario("mosTipoImagem") = 2 Then
'   CaminhoImagem = "S:\Imagem\"
'   'ExtensaoImagem = ".bmp"
'   Foto_Pedida = CaminhoImagem & TabControleMostruario("mosEnderecoImagem") & TabControleMostruario("mostransporte")
'Else
'   Foto_Pedida = Caminho & TabControleMostruario("mosEnderecoImagem") & Extensao
'End If

'If TabControleMostruario("mosstatus") = 0 Then
   'frmImagemRetrato.txtPessoa = lblCodCliente4
   'frmImagemRetrato.txtOrdem = GridMostruario.TextMatrix(Linha, 0)
   'frmImagemRetrato.txtComentarios = TabControleMostruario("moscomentariomost")
'   frmImagemRetrato.Show
'Else
'   frmImagemPaisagem.lblComentarios = TabControleMostruario("moscomentariomost")
'   frmImagemPaisagem.Show
'End If

End Sub

Private Sub lstCliente_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Call Rotina_Limpa_ResumoCliente

lblCodCliente = lstCliente.List(lstCliente.ListIndex)
lblCodCliente1 = lstCliente.List(lstCliente.ListIndex)
lblCodCliente2 = lstCliente.List(lstCliente.ListIndex)
lblCodCliente3 = lstCliente.List(lstCliente.ListIndex)
lblCodCliente4 = lstCliente.List(lstCliente.ListIndex)
Call Rotina_AbrirBanco


pes.Open "Select * from Pessoa where chPessoa = ('" & lblCodCliente & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Verifique o Cadastro de Clentes"), vbCritical
   Call FechaDB
   Exit Sub
Else
   Call Rotina_Carrega_ResumoCliente
   Call Rotina_Carga_Contato
End If

End Sub

Public Sub Rotina_Carrega_ResumoCliente()
   
lblCodCliente = pes!chPessoa
lblCodCliente1 = pes!chPessoa
lblCodCliente2 = pes!chPessoa
lblCodCliente3 = pes!chPessoa
lblCodCliente4 = pes!chPessoa
txtTotalProc = 0
If IsNull(pes!pesDataCadastro) Then
   lblDataCadastro = Empty
Else
   lblDataCadastro = pes!pesDataCadastro
   lblDataCadastro1 = pes!pesDataCadastro
   lblDataCadastro2 = pes!pesDataCadastro
   lblDataCadastro3 = pes!pesDataCadastro
   lblDataCadastro4 = pes!pesDataCadastro
End If

    lblRazaoSocial = pes!pesRazaoSocial
    lblStatus = Status(pes!pesStatusPessoa)
    
    If pes!pesStatusPessoa = 0 Then
       lblStatus.ForeColor = vbBlack
    Else
       lblStatus.ForeColor = vbRed
    End If
       
    lblEndereco = pes!pesEndereco
    lblBairro = pes!pesBairro
    lblCidade = pes!pesCidade
    lblEstado = pes!chUF
    lblCEP = pes!pesCEP
    lblRegiao = pes!pesRegiao
    lblCNPJ = pes!chCNPJ_CPF
    lblInscEstadual = pes!pesInscEst_Ident
    'lblTelCliente = pes!pesTelefone
    'lblFaxCliente = pes!pesFax

If IsNull(pes!pesCelular) Then
   lblCelCliente = "Não Informado"
Else
   lblCelCliente = pes!pesCelular
End If

'lblContato = Tabpessoa("pesContato")
'If IsNull(Tabpessoa("pesCargoContato")) Then
'   lblCargo = "Não Informado"
'Else
'   lblCargo = Tabpessoa("pesCargoContato")
'End If

'If IsNull(Tabpessoa("pesTelContato")) Then
'   lblTelContato = "Não Informado"
'Else
'   lblTelContato = Tabpessoa("pesTelContato")
'End If

'If IsNull(Tabpessoa("pesCelContato")) Then
'   lblCelContato = "Não Informado"
'Else
'   lblCelContato = Tabpessoa("pesCelContato")
'End If

'If IsNull(Tabpessoa("pesHomePage")) Then
'   lblHomePage = "Não Informado"
'Else
'   lblHomePage = Tabpessoa("pesHomePage")
'End If

'If IsNull(Tabpessoa("pesEmail")) Then
'   lblEmail = "Não Informado"
'Else
'   lblEmail = Tabpessoa("pesEmail")
'End If
 
'lblCartRep = Tabpessoa("chCarteiraRep")

CartRep.Open "Select * from Carteira_Rep where chCarteiraRep = ('" & pes!chCarteiraRep & "')", db, 3, 3
If CartRep.EOF Then
   MsgBox ("Carteira de Representante Não Cadastrada"), vbCritical
   Call FechaDB
   Exit Sub
Else
   lblRepresentante = CartRep!chPessoa
End If

SalvaPessoa = pes!chPessoa

Contato.Open "Select * from Telefone where codPessoa = ('" & CartRep!chPessoa & "')", db, 3, 3
If Contato.EOF Then
   MsgBox "Representante sem Tel-1"
Else
   lblTelRep = Contato!codigocontato
End If
      
'      TabTelefone.Seek "=", Tabpessoa("chpessoa"), "CEL-1"
'      If TabTelefone.NoMatch Then
'         lblCelRep = Empty
'      Else
'         lblCelRep = TabTelefone("codigocontato")
'      End If
'      TabTelefone.Seek "=", Tabpessoa("chpessoa"), "FAX-1"
'      If TabTelefone.NoMatch Then
'         lblFaxRep = Empty
'      Else
'         lblFaxRep = TabTelefone("codigocontato")
'      End If
'      TabTelefone.Seek "=", Tabpessoa("chpessoa"), "E-MAIL"
'      If TabTelefone.NoMatch Then
'         lblEmailRep = Empty
'      Else
'         lblEmailRep = TabTelefone("codigocontato")
'      End If
''  End If
'End If
'Tabpessoa.Seek "=", SalvaPessoa

If pes!chCarteirapromot = "NENHUM" Then
   lblCartPromot = pes!chCarteirapromot
   lblPromotor = "-"
   lblTelPromot = "-"
   lblCelPromot = "-"
   lblFaxPromot = "-"
   lblEmailPromot = "-"
Else
   CartPromot.Open "Select * from Carteira_Promot where chCarteiraPromot = ('" & pes!chCarteirapromot & "')", db, 3, 3
   If CartPromot.EOF Then
      MsgBox ("Carteira de Promotor Não Cadastrada"), vbCritical
      Call FechaDB
      Exit Sub
   Else
      lblCartPromot = CartPromot!chCarteirapromot
      lblPromotor = CartPromot!chPessoa
   End If
   
'   If Not TabCarteira_Promot("chPessoa") = "NENHUM" Then
'      Tabpessoa.Seek "=", TabCarteira_Promot("chPessoa")
'      If Tabpessoa.NoMatch Then
'         MsgBox ("Promotor Não Cadastrado")
'         Exit Sub
'      Else
'         TabTelefone.Seek "=", Tabpessoa("chpessoa"), "TEL-1"
'         If TabTelefone.NoMatch Then
'            lblTelPromot = Empty
'            MsgBox "Representante sem Tel-1"
'         Else
'            lblTelPromot = TabTelefone("codigocontato")
'         End If
'         TabTelefone.Seek "=", Tabpessoa("chpessoa"), "CEL-1"
'         If TabTelefone.NoMatch Then
'            lblCelPromot = Empty
'         Else
'            lblCelPromot = TabTelefone("codigocontato")
'         End If
'         TabTelefone.Seek "=", Tabpessoa("chpessoa"), "FAX-1"
'         If TabTelefone.NoMatch Then
'            lblFaxPromot = Empty
'         Else
'            lblFaxPromot = TabTelefone("codigocontato")
'         End If
'         TabTelefone.Seek "=", Tabpessoa("chpessoa"), "E-MAIL"
'         If TabTelefone.NoMatch Then
'            lblEmailPromot = Empty
'         Else
'            lblEmailPromot = TabTelefone("codigocontato")
'         End If
'      End If
'    End If
   
'   Tabpessoa.Seek "=", SalvaPessoa

End If

IndPend = 0
IndProc = 0
neg.Open "Select * from Negociacao", db, 3, 3
If neg.EOF Then
   MsgBox ("Não há movimento de Negociação até a presente data."), vbInformation
   Call FechaDB
   Exit Sub
End If


neg.MoveFirst
Do While Not neg.EOF
   If neg!chPessoa = SalvaPessoa Then
      If neg!negstatus = 0 Then
         Call Rotina_Carga_Pendentes
      Else
         Call Rotina_Carga_Processados
      End If
   End If
   neg.MoveNext
   
Loop
Indice = 0

ctr.Open "Select * from Contas_A_Receber", db, 3, 3
If ctr.EOF Then
   MsgBox ("Não movimento de contas a receber até a presente data."), vbInformation
   Call FechaDB
   Exit Sub
End If

ctr.MoveFirst
Do While Not ctr.EOF
   If ctr!chPessoa > SalvaPessoa Then
      ctr.MoveLast
   Else
      If ctr!chPessoa = SalvaPessoa Then
         Indice = Indice + 1
         GridFaturamento.Rows = Indice + 1
         GridFaturamento.TextMatrix(Indice, 0) = ctr!chnotafiscal
         GridFaturamento.TextMatrix(Indice, 1) = ctr!ctrdataemissao
         GridFaturamento.TextMatrix(Indice, 2) = ctr!ctrDataVencito
         GridFaturamento.TextMatrix(Indice, 3) = ctr!chNumPedido
         GridFaturamento.TextMatrix(Indice, 4) = ctr!chNumPedidocomp
         GridFaturamento.TextMatrix(Indice, 5) = ctr!ctrDescricaoOperacao
         GridFaturamento.TextMatrix(Indice, 6) = ctr!chcodbcolart
         GridFaturamento.TextMatrix(Indice, 7) = Format$(ctr!ctrValorLart, "##,##0.00")
         GridFaturamento.TextMatrix(Indice, 8) = Format$(ctr!ctrvalordaboleta, "##,##0.00")
         If ctr!ctrStatus = 1 Then
            If ctr!ctrDataVencito < ctr!ctrDataRecebimento Then
               GridFaturamento.TextMatrix(Indice, 9) = "C/Atraso"
            Else
               GridFaturamento.TextMatrix(Indice, 9) = "Pago"
            End If
         Else
            If ctr!ctrDataVencito < Date Then
               GridFaturamento.TextMatrix(Indice, 9) = "Em Atraso"
            Else
               GridFaturamento.TextMatrix(Indice, 9) = "A/Vencer"
            End If
         End If
      End If
   End If
   ctr.MoveNext
Loop

If IndPend > 0 Then
   IndSalvo = IndPend
   Call Rotina_Totaliza_Pendentes
End If
If IndProc > 0 Then
   IndSalvo = IndProc
   Call Rotina_Totaliza_Processados
End If

txtTotalPend = Format$(AcumulaPend, "#,##0.00")
txtTotalProc = Format$(AcumulaProc, "#,##0.00")

Call FechaDB

End Sub

Public Sub Rotina_Limpa_ResumoCliente()
txtConsulta = Empty
lblCodCliente1 = Empty
lblCodCliente = Empty
lblDataCadastro = Empty
lblRazaoSocial = Empty
lblEndereco = Empty
lblBairro = Empty
lblCidade = Empty
lblEstado = Empty
lblCEP = Empty
lblRegiao = Empty
lblCNPJ = Empty
lblInscEstadual = Empty
'lblContato = Empty
'lblCargo = Empty
'lblTelContato = Empty
'lblCelContato = Empty
'lblTelCliente = Empty
'lblFaxCliente = Empty
'lblCelCliente = Empty
'lblEmail = Empty
'lblHomePage = Empty
lblCartPromot = Empty
lblPromotor = Empty
lblTelPromot = Empty
lblCelPromot = Empty
lblFaxPromot = Empty
lblEmailPromot = Empty
lblCartRep = Empty
lblRepresentante = Empty
lblTelRep = Empty
lblCelRep = Empty
lblFaxRep = Empty
lblEmailRep = Empty

GridPedidos.Rows = 2
GridProc.Rows = 2
GridFaturamento.Rows = 2

GridPedidos.TextMatrix(1, 0) = Empty
GridPedidos.TextMatrix(1, 1) = Empty

GridPedidos.TextMatrix(1, 2) = Empty
GridPedidos.TextMatrix(1, 3) = Empty
GridPedidos.TextMatrix(1, 4) = Empty
GridPedidos.TextMatrix(1, 5) = Empty
GridPedidos.TextMatrix(1, 6) = Empty
GridPedidos.TextMatrix(1, 7) = Empty
GridPedidos.TextMatrix(1, 8) = Empty
GridPedidos.TextMatrix(1, 9) = Empty
GridPedidos.TextMatrix(1, 10) = Empty
GridPedidos.TextMatrix(1, 11) = Empty

GridProc.TextMatrix(1, 0) = Empty
GridProc.TextMatrix(1, 1) = Empty
GridProc.TextMatrix(1, 2) = Empty
GridProc.TextMatrix(1, 3) = Empty
GridProc.TextMatrix(1, 4) = Empty
GridProc.TextMatrix(1, 5) = Empty
GridProc.TextMatrix(1, 6) = Empty
GridProc.TextMatrix(1, 7) = Empty
GridProc.TextMatrix(1, 8) = Empty
GridProc.TextMatrix(1, 9) = Empty
GridProc.TextMatrix(1, 10) = Empty
GridProc.TextMatrix(1, 11) = Empty

GridFaturamento.TextMatrix(1, 0) = Empty
GridFaturamento.TextMatrix(1, 1) = Empty
GridFaturamento.TextMatrix(1, 2) = Empty
GridFaturamento.TextMatrix(1, 3) = Empty
GridFaturamento.TextMatrix(1, 4) = Empty
GridFaturamento.TextMatrix(1, 5) = Empty
GridFaturamento.TextMatrix(1, 6) = Empty
GridFaturamento.TextMatrix(1, 7) = Empty
GridFaturamento.TextMatrix(1, 8) = Empty
GridFaturamento.TextMatrix(1, 9) = Empty

GridResVendas.Rows = 2
GridResVendas.TextMatrix(1, 0) = Empty
GridResVendas.TextMatrix(1, 1) = Empty
GridResVendas.TextMatrix(1, 2) = Empty

AcumulaPend = 0
AcumulaProc = 0
GridConsumidos.Rows = 2
GridConsumidos.TextMatrix(1, 0) = Empty
GridConsumidos.TextMatrix(1, 1) = Empty
GridConsumidos.TextMatrix(1, 2) = Empty
GridConsumidos.TextMatrix(1, 3) = Empty
txtQtdTotal = Format$(0, "0.00")

GridConsProd.Rows = 2
GridConsProd.TextMatrix(1, 0) = Empty
GridConsProd.TextMatrix(1, 1) = Empty
GridConsProd.TextMatrix(1, 2) = Empty
GridConsProd.TextMatrix(1, 3) = Empty

txtTotalMetroProduto = Format$(0, "#0.00")
txtTotalMetrosPeriodo = Format$(0, "#0.00")
lblQuantidadeMaiorVenda = Format$(0, "#0.00")
lblQuantidadeMenorVenda = Format$(0, "#0.00")
cmbProdutoVendas = Empty

GridHistFat.Rows = 2
GridHistFat.TextMatrix(1, 0) = Empty
GridHistFat.TextMatrix(1, 1) = Empty
GridHistFat.TextMatrix(1, 2) = Empty
GridHistFat.TextMatrix(1, 3) = Empty
GridHistFat.TextMatrix(1, 4) = Empty
GridHistFat.TextMatrix(1, 5) = Empty
GridHistFat.TextMatrix(1, 6) = Empty
GridHistFat.TextMatrix(1, 7) = Empty
GridHistFat.TextMatrix(1, 8) = Empty
GridHistFat.TextMatrix(1, 9) = Empty
GridHistFat.TextMatrix(1, 10) = Empty

GridFormaPagto.Rows = 2
GridFormaPagto.TextMatrix(1, 0) = Empty
GridFormaPagto.TextMatrix(1, 1) = Empty
GridFormaPagto.TextMatrix(1, 2) = Empty
GridFormaPagto.TextMatrix(1, 3) = Empty
GridFormaPagto.TextMatrix(1, 4) = Empty
GridFormaPagto.TextMatrix(1, 5) = Empty
GridFormaPagto.TextMatrix(1, 6) = Empty
GridFormaPagto.TextMatrix(1, 7) = Empty

GridMostruario.Rows = 2
GridMostruario.TextMatrix(1, 0) = Empty
GridMostruario.TextMatrix(1, 1) = Empty
GridMostruario.TextMatrix(1, 2) = Empty
GridMostruario.TextMatrix(1, 3) = Empty
GridMostruario.TextMatrix(1, 4) = Empty
GridMostruario.TextMatrix(1, 5) = Empty

txtDataMaiorFat = "__/__/____"
txtMaiorFatPeriodo = Format$(0, "##,##0.00")
txtMesFat = Empty
txtFatMensal = Format$(0, "##,##0.00")
txtQtdPagtoAtrasoPeriodo = 0
txtMairoQtdDiasAtraso = 0
txtValorAtrsoPeriodo = Format$(0, "##,##0.00")

End Sub





Private Sub txtConsulta_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
Public Sub Rotina_Carga_Pendentes()
IndPend = 0
IndProc = 0
Dia = Day(TabNegociacao("negDataNegociação"))
Mes = Month(TabNegociacao("negDataNegociação"))
Ano = Year(TabNegociacao("negDataNegociação"))
TabDetNeg.Seek "=", TabNegociacao("chNumPedido"), TabNegociacao("chNumPedidocomp")
If TabDetNeg.NoMatch Then
   If Not TabNegociacao("negStatus") = 3 Then
      MsgBox ("Não Existe Detalhe Para o Pedido. Verifique Detalhe da Negociação!")
   End If
Else
   Do While (TabDetNeg("chnumpedido") = TabNegociacao("chnumpedido") And TabDetNeg("chnumpedidocomp") = TabNegociacao("chnumpedidocomp"))
      IndPend = IndPend + 1
      GridPedidos.Rows = IndPend + 1
      GridPedidos.TextMatrix(IndPend, 0) = Data_Inv
      GridPedidos.TextMatrix(IndPend, 1) = TabNegociacao("negDatapedido")
      GridPedidos.TextMatrix(IndPend, 2) = TabNegociacao("chNumPedido")
      GridPedidos.TextMatrix(IndPend, 3) = TabNegociacao("chNumPedidocomp")
      GridPedidos.TextMatrix(IndPend, 4) = TabNegociacao("chpessoa")
      tabproduto.Seek "=", TabDetNeg("chProduto")
      GridPedidos.TextMatrix(IndPend, 5) = tabproduto("prdnomeProd")
      If TabDetNeg("pedUnidade") = 1 Then
         GridPedidos.TextMatrix(IndPend, 6) = "M"
      Else
         GridPedidos.TextMatrix(IndPend, 6) = "Un"
      End If
      GridPedidos.TextMatrix(IndPend, 7) = Format$(TabDetNeg("pedQuantidadePedida"), "0.00")
      GridPedidos.TextMatrix(IndPend, 8) = Format$(TabDetNeg("pedprecounidadepedida"), "0.00")
      GridPedidos.TextMatrix(IndPend, 9) = Format$((TabDetNeg("pedquantidadepedida") * TabDetNeg("pedprecounidadepedida")), "0.00")
      GridPedidos.TextMatrix(IndPend, 10) = Format$(TabDetNeg("pedqtddias"), "0.00")
      GridPedidos.TextMatrix(IndPend, 11) = Format$((TabDetNeg("pedqtddias")) * (TabDetNeg("pedquantidadepedida") * TabDetNeg("pedprecounidadepedida")), "#,##0.00")
      TabDetNeg.MoveNext
      If TabDetNeg.EOF Then
         TabDetNeg.MoveFirst
      End If
   Loop
End If
 
End Sub

Public Sub Rotina_Carga_Processados()


Dia = Day(neg!negDataNegociação)
Mes = Month(neg!negDataNegociação)
Ano = Year(neg!negDataNegociação)

dneg.Open "Select * from DetalheNegociacao where chNumPedido = ('" & neg!chNumPedido & "') and chNumPedidoComp = ('" & neg!chNumPedidocomp & "')", db, 3, 3
If dneg.EOF Then
   MsgBox ("Não Existe Detalhe Para o Pedido. Verifique Detalhe da Negociação!"), vbCritical
Else
   Do While fim = 0
      If (dneg!chNumPedido = neg!chNumPedido) And dneg!chNumPedidocomp = neg!chNumPedidocomp Then
        IndProc = IndProc + 1
        GridProc.Rows = IndProc + 1
        GridProc.TextMatrix(IndProc, 0) = Data_Inv
        GridProc.TextMatrix(IndProc, 1) = neg!negDataNegociação
        GridProc.TextMatrix(IndProc, 2) = neg!chNumPedido
        GridProc.TextMatrix(IndProc, 3) = neg!chNumPedidocomp
        GridProc.TextMatrix(IndProc, 4) = neg!chPessoa
        'tabproduto.Seek "=", dneg!chProduto
       ' GridProc.TextMatrix(IndProc, 5) = Prod!prdNomeProd
        If dneg!pedUnidade = 1 Then
           GridProc.TextMatrix(IndProc, 6) = "M"
        Else
           GridProc.TextMatrix(IndProc, 6) = "Un"
        End If
        GridProc.TextMatrix(IndProc, 7) = Format$(dneg!pedquantidadepedida, "0.00")
        GridProc.TextMatrix(IndProc, 8) = Format$(dneg!pedprecounidadepedida, "0.00")
        GridProc.TextMatrix(IndProc, 9) = Format$((dneg!pedquantidadepedida * dneg!pedprecounidadepedida), "0.00")
        GridProc.TextMatrix(IndProc, 10) = Format$(dneg!pedqtddias, "0.00")
        GridProc.TextMatrix(IndProc, 11) = Format$((dneg!pedqtddias) * (dneg!pedquantidadepedida * dneg!pedprecounidadepedida), "#,##0.00")
        dneg.MoveNext
        If dneg.EOF Then
           fim = 1
        End If
     Else
        fim = 1
     End If
   Loop

End If

If dneg.State = 1 Then
   dneg.Close: Set dneg = Nothing
   acdNeg = 0
End If

End Sub

Public Sub Rotina_Totaliza_Pendentes()

ValorPedido = 0
PedidoAnterior = GridPedidos.TextMatrix(1, 2)
PedidoCompAnterior = GridPedidos.TextMatrix(1, 3)
'GridPedidos.TextMatrix(1, 11) = Empty
QtdPedido = 0
Qtd = 0
For IndPend = 1 To IndSalvo
    If GridPedidos.TextMatrix(IndPend, 2) = PedidoAnterior And GridPedidos.TextMatrix(IndPend, 3) = PedidoCompAnterior Then
       
       If Not GridPedidos.TextMatrix(IndPend, 1) = Empty Then
          ValorPedido = ValorPedido + GridPedidos.TextMatrix(IndPend, 11)
          Qtd = Qtd + 1
          If Qtd > 1 Then
             GridPedidos.TextMatrix(IndPend, 1) = Empty
             GridPedidos.TextMatrix(IndPend, 2) = Empty
             GridPedidos.TextMatrix(IndPend, 3) = Empty
             GridPedidos.TextMatrix(IndPend, 4) = Empty
             GridPedidos.TextMatrix(IndPend, 12) = Empty
          End If
       End If
    Else
       If Not GridPedidos.TextMatrix(IndPend, 1) = Empty Then
          GridPedidos.TextMatrix(IndPend - 1, 12) = Format$(ValorPedido, "#,##0.00")
          AcumulaProc = AcumulaProc + ValorPedido
          ValorPedido = 0
          Qtd = 1
          PedidoAnterior = GridPedidos.TextMatrix(IndPend, 2)
          PedidoCompAnterior = GridPedidos.TextMatrix(IndPend, 3)
          ValorPedido = GridPedidos.TextMatrix(IndPend, 11)
          GridPedidos.TextMatrix(IndPend, 12) = Empty
       End If
    End If
Next
GridPedidos.TextMatrix(IndPend - 1, 12) = Format$(ValorPedido, "#,##0.00")
AcumulaPend = AcumulaPend + ValorPedido
End Sub

Public Sub Rotina_Totaliza_Processados()
ValorPedido = 0
PedidoAnterior = GridProc.TextMatrix(1, 2)
PedidoCompAnterior = GridProc.TextMatrix(1, 3)
'GridProc.TextMatrix(1, 11) = Empty
QtdPedido = 0
Qtd = 0
For IndProc = 1 To IndSalvo
    If GridProc.TextMatrix(IndProc, 2) = PedidoAnterior And GridProc.TextMatrix(IndProc, 3) = PedidoCompAnterior Then
       
       If Not GridProc.TextMatrix(IndProc, 1) = Empty Then
          ValorPedido = ValorPedido + GridProc.TextMatrix(IndProc, 11)
          Qtd = Qtd + 1
          If Qtd > 1 Then
             GridProc.TextMatrix(IndProc, 1) = Empty
             GridProc.TextMatrix(IndProc, 2) = Empty
             GridProc.TextMatrix(IndProc, 3) = Empty
             GridProc.TextMatrix(IndProc, 4) = Empty
             GridProc.TextMatrix(IndProc, 12) = Empty
          End If
       End If
    Else
       If Not GridProc.TextMatrix(IndProc, 1) = Empty Then
          GridProc.TextMatrix(IndProc - 1, 12) = Format$(ValorPedido, "#,##0.00")
          AcumulaProc = AcumulaProc + ValorPedido
          ValorPedido = 0
          Qtd = 1
          PedidoAnterior = GridProc.TextMatrix(IndProc, 2)
          PedidoCompAnterior = GridProc.TextMatrix(IndProc, 3)
          ValorPedido = GridProc.TextMatrix(IndProc, 11)
          GridProc.TextMatrix(IndProc, 12) = Empty
       End If
    End If
Next
GridProc.TextMatrix(IndProc - 1, 12) = Format$(ValorPedido, "#,##0.00")
AcumulaProc = AcumulaProc + ValorPedido
End Sub

Public Sub Rotina_GridResumoVendas()

For Indice = 1 To 239
    If MesAnoTab(Indice) = MesAnoArq Then
       QtdTab(Indice) = QtdTab(Indice) + TabHistoricoDetNeg("hdnquantidademetro")
       Indice = 239
    End If
Next
End Sub

Public Sub Rotina_GridConsumidos()
For Indice = 1 To 99
    If codprodtab(Indice) = TabHistoricoDetNeg("chproduto") Then
       QtdProdTab(Indice) = QtdProdTab(Indice) + TabHistoricoDetNeg("hdnquantidademetro")
       Indice = 99
    End If
Next
End Sub

Public Sub Rotina_GridFormaPagto()
If TabHistoricoNegociacao("hngstatus") > 1 Then
   Exit Sub
Else
   IndNeg = IndNeg + 1
   GridFormaPagto.Rows = IndNeg + 1
   GridFormaPagto.TextMatrix(IndNeg, 0) = TabHistoricoNegociacao("chnumpedido")
   GridFormaPagto.TextMatrix(IndNeg, 1) = TabHistoricoNegociacao("chnumpedidocomp")
   GridFormaPagto.TextMatrix(IndNeg, 2) = TabHistoricoNegociacao("chdatanegociacao")
   GridFormaPagto.TextMatrix(IndNeg, 3) = TabHistoricoNegociacao("hngfaturamento")
   GridFormaPagto.TextMatrix(IndNeg, 4) = TabHistoricoNegociacao("hngintervalofatura")
   GridFormaPagto.TextMatrix(IndNeg, 5) = TabHistoricoNegociacao("hngapartirde")
   GridFormaPagto.TextMatrix(IndNeg, 6) = 0 'TabHistoricoDetNeg("hdndesc") & "%"
   GridFormaPagto.TextMatrix(IndNeg, 7) = Year(TabHistoricoNegociacao("chdatanegociacao")) & Format$(Month(TabHistoricoNegociacao("chdatanegociacao")), "00") & Format$(Day(TabHistoricoNegociacao("chdatanegociacao")), "00")
End If
End Sub

Public Sub Rotina_Carga_Contato()

'Data1.Refresh

'    Data1.Recordset.FindFirst ("CodPessoa = '" & lblCodCliente & "'")

'If Data1.Recordset.NoMatch Then
'   Exit Sub
'Else
'   Pessoa = Data1.Recordset.Fields("codpessoa")
'   TipoContato = Data1.Recordset.Fields("tipocontato")
'End If
'MsgBox "Inicio da carga telefone"
Call Rotina_AbrirBanco

pes.Open "Select * from Pessoa where chPessoa = ('" & lblCodCliente & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Verifique o Cadastro de Clentes"), vbCritical
   Call FechaDB
   Exit Sub
End If

Contato.Open "Select * from Telefone where codPessoa = ('" & pes!chPessoa & "')", db, 3, 3
If Contato.EOF Then
   MsgBox ("Erro: Nao encontrei o Contato desse Cliente. Verificar."), vbInformation
End If
IndContato = 0
Ind = 0

Contato.MoveFirst

Do While fim = 0
   If lblCodCliente = Contato!codpessoa Then
        Ind = Ind + 1
        IndContato = IndContato + 1
        GridContato.Rows = Ind + 1
        GridContato.TextMatrix(Ind, 0) = Contato!TipoContato
        GridContato.TextMatrix(Ind, 1) = Contato!codigocontato
        Contato.MoveNext
        If Contato.EOF Then
           fim = 1
        End If
    Else
        fim = 1
    End If
Loop
'MsgBox "Final da carga telefone"
End Sub
