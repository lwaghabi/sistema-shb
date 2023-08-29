VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPessoa 
   BackColor       =   &H00E0E0E0&
   Caption         =   "frmPessoa"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   18150
   LinkTopic       =   "Form2"
   ScaleHeight     =   8265
   ScaleWidth      =   18150
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame22 
      Caption         =   "Pesquisa Pessoa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   101
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton cmdFiltro 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Filtro"
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
         Left            =   1440
         MaskColor       =   &H00FFC0C0&
         TabIndex        =   106
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox cmbstatusPessoaPesquisa 
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
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox cmbTipoPessoaPesquisa 
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
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label45 
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
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label44 
         Caption         =   "Tipo "
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
         TabIndex        =   104
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame20 
      Caption         =   "Localiza"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   55
      Top             =   7560
      Width           =   3855
      Begin VB.CommandButton cmdPesquisa 
         Caption         =   "Ok"
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
         Left            =   3240
         TabIndex        =   57
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtPesquisa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.ListBox lstPessoa 
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
      Height          =   5460
      ItemData        =   "Pessoa.frx":0000
      Left            =   120
      List            =   "Pessoa.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   54
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operação"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   53
      Top             =   7440
      Width           =   14055
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H0000FFFF&
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdExclui 
         BackColor       =   &H000000FF&
         Caption         =   "Exclui"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAltera 
         BackColor       =   &H00C0C000&
         Caption         =   "Altera"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdInclui 
         BackColor       =   &H0000FF00&
         Caption         =   "Inclui"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNovo 
         BackColor       =   &H000080FF&
         Caption         =   "Novo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   3960
      TabIndex        =   58
      Top             =   720
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   12091
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   0
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pessoa"
      TabPicture(0)   =   "Pessoa.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame15"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame16"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame17"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame18"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdProximaPagina"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fraLocador"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Endereço "
      TabPicture(1)   =   "Pessoa.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblCodPessoa"
      Tab(1).Control(1)=   "lblRazaoSocial"
      Tab(1).Control(2)=   "Frame8"
      Tab(1).Control(3)=   "cmdProxPag2"
      Tab(1).Control(4)=   "fraContato"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Detalhes"
      TabPicture(2)   =   "Pessoa.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblCodPessoa1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblRazaoSocial1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame11"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame14"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdPagInicial"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.Frame fraLocador 
         Caption         =   "Cliente Locador                                      Unidades Operacionais"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   112
         Top             =   1320
         Width           =   9015
         Begin VB.ComboBox cmbLocadora 
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
            ItemData        =   "Pessoa.frx":0058
            Left            =   120
            List            =   "Pessoa.frx":005A
            TabIndex        =   6
            Top             =   240
            Width           =   3615
         End
         Begin VB.ComboBox cmbLocaliza 
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
            Left            =   4080
            TabIndex        =   7
            Text            =   "cmbLocaliza"
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.CommandButton cmdProximaPagina 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Próxima Página"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   10800
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   6000
         Width           =   2655
      End
      Begin VB.Frame fraContato 
         Caption         =   "Contato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -74880
         TabIndex        =   107
         Top             =   2640
         Width           =   13575
         Begin MSFlexGridLib.MSFlexGrid GridContato 
            Height          =   2175
            Left            =   1080
            TabIndex        =   110
            Top             =   1080
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   3836
            _Version        =   393216
            FixedCols       =   0
            BackColor       =   16777152
            ForeColor       =   0
            BackColorFixed  =   16776960
            ForeColorFixed  =   0
            ForeColorSel    =   0
            BackColorBkg    =   16777152
            FormatString    =   "Tipo Contato              |Numero/Endereço Contato                                    "
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
         Begin VB.CommandButton cmdAlteraContato 
            BackColor       =   &H0000FFFF&
            Caption         =   "Altera"
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
            Left            =   12480
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox cmbTipoContato 
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
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   45
            Top             =   480
            Width           =   3855
         End
         Begin VB.TextBox txtCodContato 
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
            Left            =   5760
            MaxLength       =   50
            TabIndex        =   46
            Top             =   480
            Width           =   4335
         End
         Begin VB.CommandButton cmdIncluiContato 
            BackColor       =   &H00FFFF00&
            Caption         =   "Inclui"
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
            Left            =   12480
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton cmdExcluiContato 
            BackColor       =   &H008080FF&
            Caption         =   "Exclui"
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
            Left            =   12480
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   1920
            Width           =   855
         End
         Begin VB.CommandButton cmdNovoContato 
            BackColor       =   &H0000FF00&
            Caption         =   "Novo"
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
            Left            =   12480
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Contato"
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
            Left            =   1080
            TabIndex        =   109
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Código/Número"
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
            Left            =   5760
            TabIndex        =   108
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdPagInicial 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Página Inicial"
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
         Left            =   -64200
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   6000
         Width           =   2535
      End
      Begin VB.CommandButton cmdProxPag2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Proxima Página"
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
         Left            =   -64200
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   5880
         Width           =   2655
      End
      Begin VB.Frame Frame6 
         Height          =   1695
         Left            =   9360
         TabIndex        =   97
         Top             =   1440
         Width           =   4575
         Begin MSMask.MaskEdBox txtCNPJCPF 
            Height          =   375
            Left            =   1440
            TabIndex        =   12
            Top             =   360
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   20
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
         Begin VB.TextBox txtInscEstIdent 
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
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   13
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label lblInscIdent 
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
            Height          =   375
            Left            =   120
            TabIndex        =   99
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblCPFCGC 
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
            Height          =   375
            Left            =   120
            TabIndex        =   98
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame18 
         Height          =   2775
         Left            =   9360
         TabIndex        =   88
         Top             =   3120
         Width           =   4575
         Begin VB.CommandButton cmdProxPag 
            BackColor       =   &H00FFFF00&
            Caption         =   "Proxima Página"
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
            Index           =   0
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   1560
            Width           =   2295
         End
         Begin VB.Frame Frame21 
            Caption         =   "Correspondencia"
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
            Left            =   5880
            TabIndex        =   96
            Top             =   240
            Width           =   2295
            Begin VB.OptionButton OptCartaNao 
               Caption         =   "Não"
               Height          =   255
               Left            =   1320
               TabIndex        =   51
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton OptCartaSim 
               Caption         =   "Sim"
               Height          =   255
               Left            =   360
               TabIndex        =   43
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.ComboBox cmbRamoAtividade 
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
            ItemData        =   "Pessoa.frx":005C
            Left            =   720
            List            =   "Pessoa.frx":0081
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1920
            Width           =   2775
         End
         Begin VB.ComboBox cmbCadastroPessoa 
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
            ItemData        =   "Pessoa.frx":010E
            Left            =   720
            List            =   "Pessoa.frx":0110
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   360
            Width           =   2775
         End
         Begin VB.ComboBox cmbStatusPessoa 
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
            ItemData        =   "Pessoa.frx":0112
            Left            =   720
            List            =   "Pessoa.frx":0125
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Ramo Atividade"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   960
            TabIndex        =   91
            Top             =   1560
            Width           =   1530
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Cadastro Pessoa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1080
            TabIndex        =   90
            Top             =   120
            Width           =   1620
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Status Pessoa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1080
            TabIndex        =   89
            Top             =   840
            Width           =   1365
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Carteira Representante                  /      Promotor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         TabIndex        =   87
         Top             =   3720
         Width           =   6855
         Begin VB.ComboBox cmbPromotora 
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
            Left            =   3960
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   2415
         End
         Begin VB.ComboBox cmbRepresentante 
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
            TabIndex        =   15
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame16 
         Height          =   975
         Left            =   240
         TabIndex        =   83
         Top             =   360
         Width           =   9015
         Begin VB.TextBox txtCodPessoa 
            DataSource      =   "Data1"
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
            MaxLength       =   20
            TabIndex        =   1
            Top             =   480
            Width           =   3975
         End
         Begin VB.ComboBox cmbPessoa 
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
            ItemData        =   "Pessoa.frx":015C
            Left            =   6600
            List            =   "Pessoa.frx":015E
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   480
            Width           =   1935
         End
         Begin VB.ComboBox cmbTipoPessoa 
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
            ItemData        =   "Pessoa.frx":0160
            Left            =   4320
            List            =   "Pessoa.frx":0162
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            Caption         =   "Pessoa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6600
            TabIndex        =   86
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            Caption         =   "Cod Pessoa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            Caption         =   "Tipo Pessoa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4320
            TabIndex        =   84
            Top             =   240
            Width           =   1170
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Data Encerra."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   11760
         TabIndex        =   82
         Top             =   480
         Width           =   2175
         Begin MSMask.MaskEdBox txtDataEncerra 
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
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
      Begin VB.Frame Frame14 
         Caption         =   "Considerações"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   81
         Top             =   3720
         Width           =   13095
         Begin VB.TextBox txtConsideracoes 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1725
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   37
            Top             =   240
            Width           =   12735
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Contatos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74760
         TabIndex        =   77
         Top             =   840
         Width           =   12975
         Begin VB.TextBox txtEmail 
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
            Left            =   2760
            TabIndex        =   117
            Text            =   "Text1"
            Top             =   1440
            Width           =   7695
         End
         Begin VB.TextBox txtSalario 
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
            Height          =   375
            Left            =   7080
            TabIndex        =   115
            Top             =   2280
            Width           =   2415
         End
         Begin VB.TextBox txtCelularContato 
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
            Left            =   4080
            MaxLength       =   15
            TabIndex        =   36
            Top             =   2280
            Width           =   2775
         End
         Begin VB.TextBox txtTelContato 
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
            MaxLength       =   15
            TabIndex        =   35
            Top             =   2280
            Width           =   3615
         End
         Begin VB.TextBox txtCargoContato 
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
            Left            =   7080
            MaxLength       =   50
            TabIndex        =   34
            Top             =   600
            Width           =   4335
         End
         Begin VB.TextBox txtContato 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            MaxLength       =   30
            TabIndex        =   33
            Top             =   600
            Width           =   5415
         End
         Begin VB.Label Label14 
            Caption         =   "Email Contato Comercial"
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
            Left            =   2760
            TabIndex        =   116
            Top             =   1080
            Width           =   3615
         End
         Begin VB.Label Label13 
            Caption         =   "Salário"
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
            Left            =   7080
            TabIndex        =   114
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
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
            TabIndex        =   111
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label31 
            Caption         =   "Celular"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   80
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Telefones"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   79
            Top             =   2040
            Width           =   1305
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Cargo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   7800
            TabIndex        =   78
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   0
         Top             =   1200
         Width           =   13575
         Begin MSMask.MaskEdBox txtCEP 
            Height          =   405
            Left            =   6120
            TabIndex        =   30
            Top             =   960
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   9
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
         Begin VB.ComboBox txtRegiao 
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
            Left            =   8760
            TabIndex        =   31
            Top             =   960
            Width           =   4095
         End
         Begin VB.TextBox txtEstado 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4680
            TabIndex        =   29
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtBairro 
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
            Left            =   7680
            MaxLength       =   50
            TabIndex        =   27
            Top             =   360
            Width           =   5415
         End
         Begin VB.TextBox txtCidade 
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
            MaxLength       =   50
            TabIndex        =   28
            Top             =   960
            Width           =   3975
         End
         Begin VB.TextBox txtEndereco 
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
            MaxLength       =   50
            TabIndex        =   26
            Top             =   360
            Width           =   6495
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Região"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8760
            TabIndex        =   76
            Top             =   720
            Width           =   660
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6120
            TabIndex        =   75
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4800
            TabIndex        =   74
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7680
            TabIndex        =   73
            Top             =   120
            Width           =   600
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            BeginProperty Font 
               Name            =   "Verdana"
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
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   71
            Top             =   120
            Width           =   915
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Informações Bancárias"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   65
         Top             =   4800
         Width           =   9015
         Begin VB.TextBox txtBanco 
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
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtAgencia 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2640
            TabIndex        =   18
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtConta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5040
            TabIndex        =   19
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtTitular 
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
            Left            =   2760
            TabIndex        =   21
            Top             =   1440
            Width           =   4455
         End
         Begin VB.TextBox txtCpfTit 
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
            Left            =   240
            TabIndex        =   20
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "cpf/cgc"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   70
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   69
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Agencia"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2640
            TabIndex        =   68
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5040
            TabIndex        =   67
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Titularidade"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2760
            TabIndex        =   66
            Top             =   1080
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2415
         Left            =   240
         TabIndex        =   60
         Top             =   2040
         Width           =   9015
         Begin VB.ComboBox cmbClassFiscal 
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
            Left            =   6480
            TabIndex        =   10
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtAniversario 
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
            Left            =   360
            TabIndex        =   14
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox txtRazaoSocial 
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
            MaxLength       =   80
            TabIndex        =   11
            Top             =   1200
            Width           =   8415
         End
         Begin VB.TextBox txtFantasia 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3240
            MaxLength       =   50
            TabIndex        =   9
            Top             =   360
            Width           =   3255
         End
         Begin VB.TextBox txtGrupo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            MaxLength       =   20
            TabIndex        =   8
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label lblClassFiscal 
            Caption         =   "Classificação Fiscal"
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
            TabIndex        =   113
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Aniversário"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   360
            TabIndex        =   64
            Top             =   1680
            Width           =   1485
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Razão Social/Nome"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   63
            Top             =   840
            Width           =   2595
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fantasia"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3240
            TabIndex        =   62
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Grupo "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   120
            Width           =   645
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Data Cadastro"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9480
         TabIndex        =   59
         Top             =   480
         Width           =   2055
         Begin MSMask.MaskEdBox txtDataCadastro 
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
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
      Begin VB.Label lblRazaoSocial1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -72360
         TabIndex        =   95
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label lblCodPessoa1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -74760
         TabIndex        =   94
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblRazaoSocial 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72360
         TabIndex        =   93
         Top             =   600
         Width           =   6135
      End
      Begin VB.Label lblCodPessoa 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   92
         Top             =   600
         Width           =   2295
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Registro e Atualização de Clientes e Colaboradores - Pessoa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3960
      TabIndex        =   100
      Top             =   120
      Width           =   14055
   End
End
Attribute VB_Name = "frmPessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim TipoUsuario As Byte
Dim Novo As Byte
Dim Resp As String
Dim Pessoa As String
Dim Altera As Byte
Dim fim As Byte
Dim SalvaPessoa As String
Dim Ind As Integer
Dim NaoInclui As Byte
Dim NaoAchei As Byte
Dim IndAchei As Byte
Dim RepAnterior(50) As String
Dim Limite As Integer
Dim IndCombo As Integer
Dim Filtro As Byte
Dim Sql As String
Dim IndContato As Integer
Dim TipoContato As String
Dim IndLinha As Integer

Private Sub cmbLocadora_LostFocus()

cmbLocaliza.Clear

Call Rotina_AbrirBanco

uoper.Open "Select * from UnidadeOperacional where chpessoa = ('" & cmbLocadora & "')", db, 3, 3
If uoper.EOF Then
   MsgBox ("Cadastrar inicialmente as unidades operacionais deste funcionário."), vbCritical
   Call FechaDB
   Exit Sub
End If

Do While Not uoper.EOF
   cmbLocaliza.AddItem uoper!chUnidadeOperacional
   uoper.MoveNext
Loop

Call FechaDB

End Sub
Public Sub Carrega_cmbLocadora()

cmbLocaliza.Clear

uoper.Open "Select * from UnidadeOperacional where chpessoa = ('" & cmbLocadora & "')", db, 3, 3
If uoper.EOF Then
   MsgBox ("Cadastrar inicialmente as unidades operacionais deste funcionário."), vbCritical
   Call FechaDB
   Exit Sub
End If

Do While Not uoper.EOF
   cmbLocaliza.AddItem uoper!chUnidadeOperacional
   uoper.MoveNext
Loop

End Sub

Private Sub cmbPessoa_LostFocus()
    
Call Rotina_AbrirBanco

pes.Open "Select * from Pessoa where chPessoa = ('" & txtCodPessoa & "')", db, 3, 3

If pes.EOF Then
   If cmbPessoa.ListIndex = 0 Then
      lblCPFCGC.Caption = "CPF "
      lblInscIdent.Caption = "Identidade"
      txtCNPJCPF.Mask = "###.###.###-##"
      txtCNPJCPF = "___.___.___-__"
   Else
      lblCPFCGC.Caption = "CNPJ"
      lblInscIdent.Caption = "I.Estadual"
      txtCNPJCPF.Mask = "##.###.###/####-##"
      txtCNPJCPF = "__.___.___/____-__"
   End If
End If

Call FechaDB

End Sub

Private Sub cmbTipoContato_LostFocus()

Call Rotina_AbrirBanco

pes.Open "Select * from Pessoa where chPessoa = ('" & txtCodPessoa & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Continuar a Inclusão de Pessoa. Após a inclusão, inserir os contatos."), vbInformation
   SSTab1.Tab = 2
   txtContato.SetFocus
   Exit Sub
End If

If cmbTipoContato = Empty Then
   Call FechaDB
   Exit Sub
End If
   

Contato.Open "Select * from Telefone where codPessoa = ('" & pes!chPessoa & "') and TipoContato = ('" & cmbTipoContato & "')", db, 3, 3
If Contato.EOF Then
   
   If IndContato = 0 Then
      cmdIncluiContato.Enabled = True
      cmdExcluiContato.Enabled = False
      cmdAlteraContato.Enabled = False
      cmdNovoContato = False
   Else
      For Ind = 1 To IndContato
          GridContato.Rows = Ind + 1
          If GridContato.TextMatrix(Ind, 0) = cmbTipoContato Then
             NaoAchei = 0
             IndAchei = Ind
             Ind = IndContato
          Else
             NaoAchei = 1
          End If
      Next
      If NaoAchei = 1 Then
         cmdIncluiContato.Enabled = True
         cmdAlteraContato.Enabled = False
         cmdExcluiContato.Enabled = False
         cmdNovoContato = False
         NaoAchei = 0
      Else
         txtCodContato = GridContato.TextMatrix(IndAchei, 1)
         cmdIncluiContato.Enabled = False
         cmdAlteraContato.Enabled = True
         cmdExcluiContato.Enabled = True
         cmdNovoContato.Enabled = True
      End If
   End If
Else
   If Contato!TipoContato = cmbTipoContato Then
        txtCodContato = Contato!CodigoContato
        cmdIncluiContato.Enabled = False
        cmdAlteraContato.Enabled = True
        cmdExcluiContato.Enabled = True
        cmdNovoContato.Enabled = True
    Else
        MsgBox "Verificar a utilização de caixa alta conforme a relação ao lado"
        cmdIncluiContato.Enabled = False
        cmdAlteraContato.Enabled = False
        cmdExcluiContato.Enabled = False
        cmdNovoContato = True
        cmdSair.SetFocus
    End If
End If
End Sub

Private Sub cmbTipoPessoa_LostFocus()

If cmbTipoPessoa = "Funcionario" Then
   fraLocador.Visible = True
Else
   fraLocador.Visible = False
End If

If cmbTipoPessoa = "Cliente" Then
   cmbClassFiscal.Visible = True
Else
   cmbClassFiscal.Visible = False
End If

End Sub

Private Sub cmdAltera_Click()
   
On Error GoTo Erro:

   Call Rotina_Critica_Cadastro
   
   If Erro_Critica > 0 Then
      Erro_Critica = 0
   Else
      
      Altera = 1
      Call Rotina_AbrirBanco
      pes.Open "Select * from Pessoa where chPessoa = ('" & txtCodPessoa & "')", db, 3, 3
      If pes.EOF Then
         MsgBox ("Erro no acesso a Pessoa - Alteração"), vbCritical
         End
      End If
      
      db.BeginTrans
    
      Call Rotina_Grava_Pessoa
      
      If pes.State = 1 Then
         pes.Update
         MsgBox ("Atualização de Pessoa realizada com sucesso"), vbInformation
      End If
      
      db.CommitTrans
         
      If cmbTipoPessoa = "Funcionario" Then
         Call GeraProduto
      End If
               
     ' Call FechaDB
      
      Call Rotina_Limpa_Pessoa
      txtCodPessoa.Enabled = True
      txtCodPessoa.SetFocus
   End If
Exit Sub
Erro:
   MsgBox Err & "=" & Error
   Resume Saida
Saida:
End Sub

Private Sub cmdAlteraContato_Click()
If cmbTipoContato = Empty Then
   MsgBox "Tipo de Contato NÃO Informado"
   cmdSair.SetFocus
   Exit Sub
End If

For Ind = 0 To IndContato
    If GridContato.TextMatrix(Ind, 0) = cmbTipoContato Then
       NaoAchei = 0
       IndAchei = Ind
       Ind = IndContato
    Else
       NaoAchei = 1
    End If
Next

If NaoAchei = 1 Then
   MsgBox "Este Contato não esta inserido na lista para ser alterado"
   cmdSair.SetFocus
   Exit Sub
End If
If IndAchei > 0 Then
   IndContato = IndAchei
Else
   IndContato = IndContato + 1
End If
'GridContato.Rows = IndContato + 1
GridContato.TextMatrix(IndContato, 0) = cmbTipoContato
GridContato.TextMatrix(IndContato, 1) = txtCodContato

Call Rotina_AbrirBanco

Contato.Open "Select * from Telefone where codPessoa = ('" & txtCodPessoa & "') and TipoContato = ('" & cmbTipoContato & "')", db, 3, 3
If Contato.EOF Then
   Contato.AddNew
End If


Contato!codpessoa = txtCodPessoa
Contato!TipoContato = cmbTipoContato
Contato!CodigoContato = txtCodContato
Contato.Update

Call FechaDB

Call Rotina_Carga_Contato

cmbTipoContato = Empty
txtCodContato = Empty
cmdNovoContato.SetFocus

End Sub

Private Sub cmdExclui_Click()
On Error Resume Next
Dim Resp As String
Dim Ind As Integer
Dim TemMovito As Byte


If Not TipoUsuario = 1 Then
   MsgBox "Função exclusiva para administradores"
   Exit Sub
End If
If db.State = 0 Then
   Call Rotina_AbrirBanco
End If
neg.Open "Select * from Negociacao where chPessoa = ('" & txtCodPessoa & "')", db, 3, 3
If neg.EOF Then
   TemMovito = 0
Else
   TemMovito = 1
End If

If TemMovito = 0 Then
   hneg.Open "Select * from HistoricoNegociacao where chpessoa = ('" & txtCodPessoa & "')", db, 3, 3
   If hneg.EOF Then
      TemMovito = 0
   Else
      TemMovito = 1
   End If
End If

If TemMovito = 1 Then
   MsgBox ("Cliente com movimento de negociação. Não pode ser deletado"), vbInformation
   Exit Sub
End If


Resp = MsgBox("Exclusão de Registro. Confirma?", vbYesNo)
If Resp = vbYes Then
      db.BeginTrans
         For Ind = lstPessoa.ListCount - 1 To 0 Step -1
             If lstPessoa.List(Ind) = txtCodPessoa Then
                lstPessoa.RemoveItem (Ind)
             End If
         Next
         Call Exclui_Contato

         pes.Open "Select * from Pessoa where chpessoa = ('" & txtCodPessoa & "')", db, 3, 3
         If pes.EOF Then
            MsgBox ("Problema")
         Else
            pes.Delete
         End If
         Call Rotina_Limpa_Pessoa
      db.CommitTrans
      cmdNovo.Enabled = True
End If
      
Call Rotina_Limpa_Pessoa
      
Call FechaDB
      
End Sub

Private Sub cmdExcluiContato_Click()

For Ind = 0 To IndContato
    If GridContato.TextMatrix(Ind, 0) = cmbTipoContato Then
       NaoAchei = 0
       IndAchei = Ind
       Ind = IndContato
    Else
       NaoAchei = 1
    End If
Next
If NaoAchei = 1 Then
   MsgBox "Erro: Contato para exclusão não encontrado"
   Exit Sub
End If

Call Rotina_AbrirBanco

Contato.Open "Select * from Telefone where codPessoa = ('" & txtCodPessoa & "') and TipoContato = ('" & cmbTipoContato & "')", db, 3, 3
If Contato.EOF Then
   GridContato.TextMatrix(IndAchei, 1) = "Excluir"
Else
   Contato.Delete
End If

Call FechaDB

Call Rotina_Carga_Contato

cmdNovoContato.Enabled = True
cmdNovoContato.SetFocus

End Sub

Private Sub cmdFiltro_Click()
If cmbTipoPessoaPesquisa = Empty Then
   cmbTipoPessoaPesquisa.ListIndex = 0
End If

If cmbstatusPessoaPesquisa = Empty Then
   cmbstatusPessoaPesquisa.ListIndex = 0
End If

Filtro = 0
lstPessoa.Clear

Call Rotina_AbrirBanco

pes.Open "Select * from Pessoa", db, 3, 3
pes.MoveFirst
Do While Filtro = 0
  If pes!pestipopessoa = cmbTipoPessoaPesquisa.ListIndex - 1 Or cmbTipoPessoaPesquisa = " Geral" Then
     If pes!pesStatusPessoa = cmbstatusPessoaPesquisa.ListIndex - 1 Or cmbstatusPessoaPesquisa = " Geral" Then
        lstPessoa.AddItem pes!chPessoa
        pes.MoveNext
        If pes.EOF Then
           Filtro = 1
        End If
     Else
         pes.MoveNext
         If pes.EOF Then
            Filtro = 1
         End If
     End If
  Else
      pes.MoveNext
      If pes.EOF Then
         Filtro = 1
      End If
  End If
Loop
     
Call Rotina_Limpa_Pessoa

Call FechaDB
     
End Sub

Private Sub cmdImprimirTela_Click()
MsgBox "Função em desenvolvimento"
Exit Sub
End Sub

'Private Sub cmdImprimirTela_Click()

'Sql = "Select * from pessoa where chpessoa = '" & txtCodPessoa & "'"
'MsgBox SQL
'dePessoaPesquisa.Commands.Item("cmdpessoapesquisa").CommandText = Sql
'impPessoaPesquisa.Show vbModal
'dePessoaPesquisa.rscmdPessoaPesquisa.Close
'End Sub

Private Sub cmdInclui_Click()
On Error GoTo Erro:
   
   Novo = 0
        
   Call Rotina_Critica_Cadastro
   
   If Erro_Critica > 0 Then
      Erro_Critica = 0
   Else
      
      Call Rotina_AbrirBanco
      
      db.BeginTrans
      
      pes.Open "Select * from Pessoa where chPessoa = ('" & txtCodPessoa & "')", db, 3, 3
      If pes.EOF Then
         pes.AddNew
      End If
      
      Call Rotina_Grava_Pessoa
   
      pes.Update
      
      db.CommitTrans
      
      If cmbTipoPessoa = "Funcionario" Then
         Call GeraProduto
      End If
      
      cmdAltera.Enabled = True
      cmdExclui.Enabled = True
      cmdSair.Enabled = True
              
      lstPessoa.AddItem txtCodPessoa
      
      Resp = MsgBox("Deseja Incluir Contatos agora????", vbYesNo)
      If Resp = vbNo Then
         Call Rotina_Limpa_Pessoa
         fraContato.Visible = True
         txtCodPessoa.SetFocus
      Else
         IndContato = 1
         fraContato.Visible = True
         cmdInclui.Enabled = False
         SSTab1.Tab = 1
         cmbTipoContato.SetFocus
      End If

      
   End If
   
'Call FechaDB
Exit Sub
Erro:
   MsgBox Err & "=" & Error
   Resume Saida
Saida:
End Sub

Private Sub cmdIncluiContato_Click()

Call Rotina_AbrirBanco

pes.Open "Select * from Pessoa where chPessoa = ('" & txtCodPessoa & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("A inclusão de contatos só pode ser efetuada após a inclusão de Pessoa houver sido concluída"), vbCritical
   Exit Sub
End If


If cmbTipoContato = Empty Then
   MsgBox "Tipo de Contato NÃO Informado"
   cmdSair.SetFocus
   Exit Sub
End If

For Ind = 0 To IndContato
    GridContato.Rows = Ind + 1
    If GridContato.TextMatrix(Ind, 0) = cmbTipoContato Then
       NaoAchei = 0
       IndAchei = Ind
       Ind = IndContato
    Else
       NaoAchei = 1
    End If
Next
If NaoAchei = 1 Then
   NaoAchei = 0
Else
   MsgBox "Este Contato já esta inserido na lista"
   cmdSair.SetFocus
   Exit Sub
End If
   
      
If txtCodContato = "" Then
   MsgBox "Código ou Número do Contato não informado. Inclusão Inválida."
   Exit Sub
End If

IndContato = IndContato + 1
GridContato.Rows = IndContato + 1
GridContato.TextMatrix(IndContato, 0) = cmbTipoContato
GridContato.TextMatrix(IndContato, 1) = txtCodContato

Pessoa = txtCodPessoa

Contato.Open "Select * from Telefone where CodPessoa = ('" & txtCodPessoa & "') and TipoContato = ('" & cmbTipoContato & "')", db, 3, 3
If Contato.EOF Then
   Contato.AddNew
End If

Contato!codpessoa = txtCodPessoa
Contato!TipoContato = cmbTipoContato
Contato!CodigoContato = txtCodContato
Contato.Update

Call Rotina_Carga_Contato


    cmbTipoContato = Empty
    txtCodContato = Empty
    cmdNovoContato.SetFocus

Call FechaDB

'If cmbTipoPessoa.ListIndex = 3 Then
'   Carrega_Carteira_Rep
'End If


End Sub

'Public Sub Carrega_Carteira_Rep()

'Call Rotina_AbrirBanco

'CartRep.Open "Select * from CartRep", db, 3, 3
'If CartRep.EOF Then
   

'fim = 0
'TabCarteira_Rep.MoveFirst
'Do While fim = 0
'   If TabCarteira_Rep("chpessoa") = txtCodPessoa Then
'      fim = 1
'      NaoAchei = 1
'   Else
'      TabCarteira_Rep.MoveNext
'      If TabCarteira_Rep.EOF Then
'         fim = 1
'         NaoAchei = 0
'      End If
'   End If
'L'oop
'If NaoAchei = 0 Then
'   frmCarteiraRepresentante.txtRepresentante = txtCodPessoa
'   frmCarteiraRepresentante.Show vbModal
'End If

'End Sub
'Private Sub cmdNavega_Click(Index As Integer)
'Dim Ind As Integer

'Call Rotina_Limpa_Pessoa

'If UltPessoa = "" Then
'   Call Rotina_AbrirBanco
'   pes.Open "Select * from Pessoa", db, 3, 3
'   UltPessoa = pes!chPessoa
'End If
   
'   Select Case Index

'   Case 0
'        pes.MoveFirst
'        UltPessoa = pes!chPessoa
'   Case 1
'        pes.MoveNext
'        UltPessoa = pes!chPessoa
'   Case 2
'        pes.MovePrevious
'        UltPessoa = pes!chPessoa
'   Case 3
'        pes.MoveLast
'        UltPessoa = pes!chPessoa
'   End Select
'
'   If pes.BOF = True Then
'      pes.MoveFirst
'   End If
'
'   If pes.EOF = True Then
'      pes.MoveLast
'   End If
'
'   For Ind = 0 To lstPessoa.ListCount - 1 Step 1
'       If lstPessoa.List(Ind) = pes!chPessoa Then
'          lstPessoa.ListIndex = Ind
'       End If
'   Next
'
'   Call Rotina_Carrega_Pessoa
'
''   cmdInclui.Enabled = False
'   If Not TipoUsuario = 3 Then
'      cmdNovo.Enabled = True
'      cmdAltera.Enabled = True
'      cmdExclui.Enabled = True
'      cmdSair.Enabled = True
'   Else
''      cmdNovo.Enabled = False
'      cmdAltera.Enabled = False
'      cmdExclui.Enabled = False
'      cmdSair.Enabled = True
'   End If
'   txtCodPessoa.Enabled = False
'   txtGrupo.SetFocus
'
'
'Call FechaDB
'
'End Sub
Private Sub cmdNovo_Click()
On Error Resume Next

   Incluir = 0
   Novo = 1
   
   SSTab1.Tab = 0
   
   Call Rotina_Limpa_Pessoa
   
   cmbClassFiscal.Visible = False
   lblClassFiscal.Visible = False
   
   txtCodPessoa.Enabled = True
   txtCodPessoa.SetFocus
   
   cmdInclui.Enabled = False
   cmdAltera.Enabled = False
   cmdExclui.Enabled = False
   cmdNovo.Enabled = True
   cmdSair.Enabled = True
   
   cmbLocadora.Clear
   
End Sub

Private Sub cmdNovoContato_Click()
cmbTipoContato = Empty
txtCodContato = Empty

cmdExcluiContato.Enabled = False
cmdAlteraContato.Enabled = False

cmbTipoContato.SetFocus

End Sub

'Private Sub cmdOkBusRazaoSocial_Click()
''
'
'Sql = "Select pes.pesrazaosocial, pes.pesfantasia, pes.chpessoa "
'Sql = Sql & " from Pessoa pes where "
'Sql = Sql & " pes.pesrazaosocial like '" & txtBusRazaoSocial & "%'"
'Sql = Sql & " pes.pesrazaosocial like '%" & txtBusRazaoSocial & "%'" 'PESQUISA O CONTEUDO DO CAMPO INFORMADO EM QQ POSIÇÃO DE RAZÃO SOCIAL
'Sql = Sql & " order by pes.pesrazaosocial"
'
'deBusRazaoSocial.Commands.Item("cmdbusrazaosocial").CommandText = Sql'
'
'frmBusRazaoSocial.Show vbModal'
'
'deBusRazaoSocial.rscmdBusRazaoSocial.Close
'
'End Sub

Private Sub cmdPagInicial_Click()

SSTab1.Tab = 0

If Incluir = 1 Then
   cmdInclui.Enabled = True
   cmdInclui.SetFocus
Else
   txtCodPessoa.Enabled = True
   cmdAltera.Enabled = True
   cmdAltera.SetFocus
End If

End Sub

Private Sub cmdPesquisa_Click()

Dim Ind As Integer
Dim IndSalvo As Integer
For Ind = lstPessoa.ListCount - 1 To 0 Step -1
       If lstPessoa.List(Ind) < txtPesquisa Then
          lstPessoa.ListIndex = Ind + 1
          IndSalvo = Ind + 1
          Ind = 0
       End If
Next
cmdAltera.Enabled = True
cmdExclui.Enabled = True
txtPesquisa.SetFocus
txtCodPessoa = lstPessoa.List(IndSalvo)

Call Rotina_AbrirBanco

pes.Open "Select * from Pessoa where chPessoa = ('" & txtCodPessoa & "')", db, 3, 3

Call Rotina_Limpa_Pessoa

Call Rotina_Carrega_Pessoa

Call Rotina_Carga_Contato

End Sub


Private Sub cmdProximaPagina_Click()

SSTab1.Tab = 1

txtEndereco.SetFocus

End Sub

Private Sub cmdProxPag2_Click()

SSTab1.Tab = 2
txtContato.SetFocus

End Sub

Private Sub cmdSair_Click()

  Unload Me
End Sub
Private Sub Form_Load()
Dim fim_Carga_Pessoa As Byte
  
  'Set TabUsuario = dbSHB.OpenRecordset("Usuario")
  '                 TabUsuario.Index = "IndUsuario"
                    
 Call Rotina_AbrirBanco
 
 usu.Open "Select * from Usuario where chNome = ('" & glbUsuario & "')", db, 3, 3
                    
 If usu.EOF Then
     MsgBox "Erro no acesso ao sistema. Reiniciar"
     End
  End If
  
  TipoUsuario = usu!usuTipoAcesso
  
  cmbPessoa.AddItem "Física"
  cmbPessoa.AddItem "Jurídica"
  
  cmbTipoPessoa.AddItem "Cliente"
  cmbTipoPessoa.AddItem "Serviço"
  cmbTipoPessoa.AddItem "Fornecedor"
  cmbTipoPessoa.AddItem "Representante"
  cmbTipoPessoa.AddItem "Promotor"
  cmbTipoPessoa.AddItem "Transporte"
  cmbTipoPessoa.AddItem "Funcionario"
  cmbTipoPessoa.AddItem "Colaborador"
  cmbTipoPessoa.AddItem "RH"
  
  cmbTipoPessoaPesquisa.AddItem " Geral"
  cmbTipoPessoaPesquisa.AddItem "Cliente"
  cmbTipoPessoaPesquisa.AddItem "Serviço"
  cmbTipoPessoaPesquisa.AddItem "Fornecedor"
  cmbTipoPessoaPesquisa.AddItem "Representante"
  cmbTipoPessoaPesquisa.AddItem "Promotor"
  cmbTipoPessoaPesquisa.AddItem "Transporte"
  cmbTipoPessoaPesquisa.AddItem "Funcionario"
  cmbTipoPessoaPesquisa.AddItem "Colaborador"
  cmbTipoPessoaPesquisa.AddItem "RH"
  cmbTipoPessoaPesquisa.ListIndex = 0
  
  cmbstatusPessoaPesquisa.AddItem " Geral"
  cmbstatusPessoaPesquisa.AddItem "Ativo"
  cmbstatusPessoaPesquisa.AddItem "Em Atraso"
  cmbstatusPessoaPesquisa.AddItem "Indesejável"
  cmbstatusPessoaPesquisa.AddItem "Inativo"
  cmbstatusPessoaPesquisa.AddItem "Encerrado"
  cmbstatusPessoaPesquisa.ListIndex = 0

  cmbCadastroPessoa.AddItem "Geral"
  cmbCadastroPessoa.AddItem "SemiHermatics"
  cmbCadastroPessoa.AddItem "N.Inf."
  
  cmbClassFiscal.AddItem "Lucro Presumido"
  cmbClassFiscal.AddItem "Lucro Real"
  cmbClassFiscal.AddItem "Simples Nacional"
  
  cmbClassFiscal.Visible = False
  lblClassFiscal.Visible = False
  
  cmdInclui.Enabled = False
  cmdAltera.Enabled = False
  cmdExclui.Enabled = False
  cmdNovo.Enabled = False

fim_Carga_Pessoa = 0
flagRegAnt = 99

pes.Open "Select * from Pessoa", db, 3, 3

pes.MoveFirst
cmbLocadora.Clear
Do While fim_Carga_Pessoa = 0
   
If pes.EOF Then
   fim_Carga_Pessoa = 1
Else
   lstPessoa.AddItem pes!chPessoa
   If pes!pestipopessoa = 0 Then
      cmbLocadora.AddItem pes!chPessoa
   End If
   pes.MoveNext
End If
Loop

cmbRepresentante.Clear
cmbPromotora.Clear

For Ind = 0 To 50
    RepAnterior(Ind) = Empty
Next

IndCombo = 0
CartRep.Open "Select * from Carteira_Rep", db, 3, 3

CartRep.MoveFirst
Ind = 0
Limite = 50
Do While Not CartRep.EOF
   For Ind = 0 To Limite
       If CartRep!chPessoa = RepAnterior(Ind) Then
          NaoInclui = 1
       End If
   Next
   If NaoInclui = 0 Then

      cmbRepresentante.AddItem CartRep!chPessoa
     ' Limite = IndCombo
      RepAnterior(IndCombo) = CartRep!chPessoa
      IndCombo = IndCombo + 1
   Else
      NaoInclui = 0
   End If
   CartRep.MoveNext
Loop
  
cmbRepresentante.ListIndex = 0
  
CartPromot.Open "Select * from Carteira_Promot", db, 3, 3

CartPromot.MoveFirst
Do While Not CartPromot.EOF
   cmbPromotora.AddItem CartPromot!chPessoa
   CartPromot.MoveNext
Loop '

cmbPromotora.ListIndex = 0
cmbRepresentante.ListIndex = 0

Call Rotina_Limpa_Pessoa

Call FechaDB

fraLocador.Visible = False

End Sub




Private Sub GridContato_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Limite = GridContato.Rows

IndLinha = GridContato.Row

If GridContato.TextMatrix(IndLinha, 0) = "" Then
   MsgBox "Para Inclusão informe o novo código. Para Alteração clicar em linha com conteúdo."
   Exit Sub
End If

cmbTipoContato = GridContato.TextMatrix(IndLinha, 0)
txtCodContato = GridContato.TextMatrix(IndLinha, 1)
cmbTipoContato.SetFocus
txtCodContato.SetFocus

End Sub

Private Sub lstPessoa_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ind_Pessoa As Integer

Call Rotina_Limpa_Pessoa

txtCodPessoa = lstPessoa.List(lstPessoa.ListIndex)

Call Rotina_AbrirBanco

pes.Open "Select * from Pessoa where chpessoa = ('" & txtCodPessoa & "')", db, 3, 3

If pes.EOF Then
   MsgBox ("Deu Caquinha"), vbCritical
   'End
Else
   If pes!pestipopessoa = 0 Then
      cmbClassFiscal.Visible = True
      lblClassFiscal.Visible = True
      If Not IsNull(pes!pesClassFiscal) Then
         cmbClassFiscal = pes!pesClassFiscal
      Else
         cmbClassFiscal = Empty
      End If
   Else
      cmbClassFiscal.Visible = False
      lblClassFiscal.Visible = False
   End If
   
   cmdInclui.Enabled = False
   cmdAltera.Enabled = True
   If Not TipoUsuario = 3 Then
      cmdNovo.Enabled = True
      cmdExclui.Enabled = True
   Else
      cmdNovo.Enabled = False
      cmdAltera.Enabled = False
      cmdExclui.Enabled = False
   End If
   Call Rotina_Carrega_Pessoa
   Call Rotina_Carga_Contato
   txtGrupo.SetFocus
End If

fraContato.Visible = True

Call FechaDB

End Sub

Private Sub SSTab1_DblClick()
Aba_Pessoa = 0
Aba_Endereco = 0
Aba_Detalhes = 0

End Sub
Private Sub txtBairro_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtCargoContato_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
Private Sub txtCidade_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
Private Sub txtCodPessoa_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtCodPessoa_LostFocus()
On Error Resume Next
Dim Resp As String
   
   Incluir = 0
    
   If Novo = 1 Then
      If txtCodPessoa = Empty Then
         MsgBox ("Código do Pessoa Não Informado"), vbInformation
         txtCodPessoa.SetFocus
         Novo = 0
      Exit Sub
      End If
   Else
      If txtCodPessoa = Empty Then
         cmdInclui.Enabled = False
         Exit Sub
      Else
         If Not TipoUsuario = 3 Then
            cmdInclui.Enabled = True
         End If
      End If
   End If
     
   Novo = 0
   
   Call Rotina_AbrirBanco
   
   pes.Open "Select * from Pessoa where chPessoa = ('" & txtCodPessoa & "')", db, 3, 3
   If pes.EOF Then
      Resp = MsgBox("Inclusão de Pessoa. Confirma???", vbYesNo)
      If Resp = vbYes Then
         Incluir = 1
         fraContato.Visible = False
         cmdNovo.Enabled = False
         If Not TipoUsuario = 3 Then
            cmdInclui.Enabled = True
            cmbTipoPessoa.SetFocus
         Else
            cmdInclui = False
         End If
         cmdAltera.Enabled = False
         cmdExclui.Enabled = False
      Else
         MsgBox ("Inclusão de Pessoa Cancelada"), vbInformation
         cmdSair.SetFocus
      End If
   Else
      Call Rotina_Carrega_Pessoa
       
      GridContato.Rows = 2
      GridContato.TextMatrix(1, 0) = Empty
      GridContato.TextMatrix(1, 1) = Empty
      cmbTipoContato.Clear
      cmbTipoContato = Empty
      txtCodContato = Empty
      
      Call Rotina_Carga_Contato
      
      cmdInclui.Enabled = False
      If Not TipoUsuario = 3 Then
         cmdAltera.Enabled = True
         cmdExclui.Enabled = True
         cmdNovo.Enabled = True
      Else
         cmdAltera.Enabled = False
         cmdExclui.Enabled = False
         cmdNovo.Enabled = False
      End If
   End If

Call FechaDB

End Sub

Private Sub txtConsideracoes_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
Private Sub txtContato_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtCpfTit_LostFocus()
If txtBanco = Empty And txtAgencia = Empty And txtConta = Empty Then
   txtCpfTit = Empty
   txtTitular = Empty
   cmbRamoAtividade.SetFocus
Else
   If txtCpfTit = Empty Then
      txtCpfTit = txtCNPJCPF
      txtTitular = txtFantasia
      cmbRamoAtividade.SetFocus
   End If
End If
End Sub


Private Sub cmbStatusPessoa_LostFocus()

If cmbStatusPessoa.ListIndex = 4 Then
   If txtDataEncerra = "__/__/____" Then
      MsgBox "Informar a data de encerramento"
      txtDataEncerra.SetFocus
   End If
End If

End Sub
Public Sub Rotina_Carrega_Pessoa()
Dim flagTransp As Byte
Dim SalvaLocadora As String
Dim SalvaLocaliza As String

lblCodPessoa = pes!chPessoa

If IsNull(pes!pesRazaoSocial) Then
   lblRazaoSocial = "Não Informada"
Else
   lblRazaoSocial = pes!pesRazaoSocial
End If
 
lblCodPessoa1 = pes!chPessoa

SalvaLocadora = Empty
SalvaLocaliza = Empty

txtCodPessoa = pes!chPessoa
cmbTipoPessoa.ListIndex = pes!pestipopessoa
cmbPessoa.ListIndex = pes!pesPessoa

If IsNull(pes!pesDataCadastro) Then
   txtDataCadastro.Mask = ("##/##/####")
   txtDataCadastro = ("__/__/____")
Else
   txtDataCadastro = pes!pesDataCadastro
End If

If IsNull(pes!pesGrupo) Then
   txtGrupo = "N/INFORMADO"
Else
   txtGrupo = pes!pesGrupo
End If

If IsNull(pes!pesRazaoSocial) Then
   txtRazaoSocial = "N/INFORMADO"
Else
   txtRazaoSocial = pes!pesRazaoSocial
End If

If IsNull(pes!pesFantasia) Then
   txtFantasia = "N/INFORMADO"
Else
   txtFantasia = pes!pesFantasia
End If
 
 If IsNull(pes!chcarteirarep) Or (pes!chcarteirarep) = "NENHUM" Then
    cmbRepresentante = "NENHUM"
 Else
    cmbRepresentante = pes!chcarteirarep
 End If
 
If IsNull(pes!chCarteiraPromot) Or (pes!chCarteiraPromot) = "NENHUM" Then
    cmbPromotora = "NENHUM"
 Else
    cmbPromotora = pes!Representante
 End If
' If CartPromot.State = 1 Then
'      CartPromot.Close: Set CartPromot = Nothing
' End If

'CartPromot.Open "Select * from Carteira_promot where chpessoa = ('" & cmbPromotora & "')", db, 3, 3
'If pes.EOF Then
'   cmbPromotora = "NENHUM"
'Else
'   If pes.State = 1 Then
'      pes.Close: Set pes = Nothing
'   End If
'   pes.Open "Select * from Pessoa where chpessoa = ('" & CartPromot!chpessoa & "')", db, 3, 3
'   If pes.EOF Then
'      cmbPromotora = "NENHUM"
'   Else
'      cmbPromotora = pes!chpessoa
'   End If
'End If
'If IsNull(pes!chcarteirarep) Or Not (pes!pestipopessoa = 0) Then
'   cmbRepresentante = Empty
'Else
'   If CartRep.State = 1 Then
'      CartRep.Close: Set CartRep = Nothing
'    End If
'    Call Rotina_AbrirBanco
'   CartRep.Open "Select * from Carteira_rep", db, 3, 3
'   CartRep.MoveFirst
'   fim = 0
'   Do While fim = 0
'      If pes!chcarteirarep = CartRep!chcarteirarep Then
'         cmbRepresentante = CartRep!chpessoa
'         fim = 1
'      Else
'         CartRep.MoveNext
'         If CartRep.EOF Then
'            fim = 1
'            MsgBox ("Representante não Encontrado"), vbInformation
'            cmbRepresentante.SetFocus
'           Exit Sub '
'         End If
'      End If
'   Loop
'End If

'If Not (pes!pestipopessoa = 0 Or pes!pestipopessoa = 4) Then
'   cmbPromotora = Empty
'Else
'   If pes!pestipopessoa = 4 Then
'      cmbPromotora = Empty
'
'      CartPromot.Open "Select * from Carteira_Promot", db, 3, 3
'
'      CartPromot.MoveFirst
'      Do While Not CartPromot.EOF
'         If CartPromot!chpessoa = txtCodPessoa Then
'            cmbPromotora.AddItem CartPromot!chcarteirapromot
'            CartPromot.MoveNext
'         Else
'            CartPromot.MoveNext
'         End If
'      Loop
'   Else
'       CartPromot.MoveFirst
'       fim = 0
'       Do While fim = 0
'          If pes!chcarteirapromot = CartPromot!chcarteirapromot Then
'             cmbPromotora = CartPromot!chpessoa
'             fim = 1
'          Else
'              CartPromot.MoveNext
'              If CartPromot.EOF Then
'                 MsgBox ("Carteira de Promotor não encontrado"), vbInformation
'                 Exit Sub
'             End If
'          End If
'       Loop
'   End If
'End If
 
If IsNull(pes!pesDataEncerramento) Then
    txtDataEncerra = "__/__/____"
 Else
    txtDataEncerra = pes!pesDataEncerramento
 End If
 
 If pes!pesPessoa = 0 Then
    lblCPFCGC.Caption = "CPF "
    lblInscIdent.Caption = "Identidade"
    txtCNPJCPF.Mask = "###.###.###-##"
    If IsNull(pes!chCNPJ_CPF) Or pes!chCNPJ_CPF = " " Then
       txtCNPJCPF = "111.111.111-11"
    Else
       txtCNPJCPF = pes!chCNPJ_CPF
    End If
 Else
    lblCPFCGC.Caption = "CNPJ"
    lblInscIdent.Caption = "I.Estadual"
    txtCNPJCPF.Mask = "##.###.###/####-##"
    If IsNull(pes!chCNPJ_CPF) Then
       txtCNPJCPF = "11.111.111/1111-11"
    Else
       txtCNPJCPF = pes!chCNPJ_CPF
    End If
 End If
 
If IsNull(pes!pesInscEst_Ident) Then
   txtInscEstIdent = "N/INFORMADO"
Else
   txtInscEstIdent = pes!pesInscEst_Ident
End If

If IsNull(pes!pesRamoAtividade) Then
   cmbRamoAtividade.ListIndex = 1
Else
   cmbRamoAtividade.ListIndex = pes!pesRamoAtividade
End If

If IsNull(pes!pesRamoAtividade) Then
   cmbRamoAtividade.ListIndex = 1
Else
   cmbRamoAtividade.ListIndex = pes!pesRamoAtividade
End If

If IsNull(pes!pesCadastroPessoa) Then
   cmbCadastroPessoa.ListIndex = 1
Else
   cmbCadastroPessoa.ListIndex = pes!pesCadastroPessoa
End If

If IsNull(pes!pesStatusPessoa) Then
   cmbStatusPessoa.ListIndex = 1
Else
   cmbStatusPessoa.ListIndex = pes!pesStatusPessoa
End If
 
 If IsNull(pes!pesAniversario) Then
    txtAniversario = "N/I"
 Else
    txtAniversario = pes!pesAniversario
 End If
 
 If IsNull(pes!pesBanco) Then
    txtBanco = "N/I"
 Else
    txtBanco = pes!pesBanco
 End If
 If IsNull(pes!pesAgencia) Then
    txtAgencia = "Não Informada"
 Else
    txtAgencia = pes!pesAgencia
 End If
 If IsNull(pes!pesConta) Then
    txtConta = "Não Informada"
 Else
    txtConta = pes!pesConta
 End If
 
 If IsNull(pes!pesTitular) Then
    txtTitular = "Não Informado"
 Else
    txtTitular = pes!pesTitular
 End If
 
 If IsNull(pes!pesCPFTitular) Then
    txtCpfTit = "Não Informado"
 Else
    txtCpfTit = pes!pesCPFTitular
 End If
 
 If IsNull(pes!pesEndereco) Then
    txtEndereco = "Não Informado"
 Else
    txtEndereco = pes!pesEndereco
 End If
 
 If IsNull(pes!pesBairro) Then
    txtBairro = "Não Informado"
 Else
    txtBairro = pes!pesBairro
 End If
 
 If IsNull(pes!pesBairro) Then
    txtCidade = "Não Informado"
 Else
    txtCidade = pes!pesBairro
 End If
 
 If IsNull(pes!chUF) Then
    txtEstado = "Não Informado"
 Else
    txtEstado = pes!chUF
 End If
 
 If IsNull(pes!pesCEP) Then
    txtCEP = "00000-00"
 Else
    txtCEP = Format$(pes!pesCEP, "00000-000")
 End If
  
If IsNull(pes!pesRegiao) Then
   txtRegiao = "Não Informado"
Else
   txtRegiao = pes!pesRegiao
End If

If IsNull(pes!pesCargoContato) Then
   txtCargoContato = "Não Informado"
Else
   txtCargoContato = pes!pesCargoContato
End If
  
If IsNull(pes!pesCargoContato) Then
   txtCargoContato = "Não Informado"
Else
   txtCargoContato = pes!pesCargoContato
End If
  
If IsNull(pes!pesContato) Then
   txtContato = "Não Informado"
Else
   txtContato = pes!pesContato
End If
 
If IsNull(pes!pesTelContato) Then
   txtTelContato = "Não Informado"
Else
   txtTelContato = pes!pesTelContato
End If

If IsNull(pes!pesCelContato) Then
   txtCelularContato = "Não Informado"
Else
   txtCelularContato = pes!pesCelContato
End If

If IsNull(pes!pesEmail) Then
   txtEmail = "Não Informado"
Else
   txtEmail = pes!pesEmail
End If
 
If IsNull(pes!pesConsideracoes) Then
   txtConsideracoes = Empty
Else
   txtConsideracoes = pes!pesConsideracoes
End If

If pes!pestipopessoa = 0 Then
   If CartPromot.State = 1 Then
      CartPromot.Close: Set CartPromot = Nothing
    End If
   CartPromot.Open "Select * from Carteira_Promot where chcarteirapromot = ('" & pes!chCarteiraPromot & "')", db, 3, 3
'   If CartPromot.EOF Then
'      txtPromotora = "NENHUM"
'   Else
'      txtPromotora = CartPromot!chpessoa
'   End If
   
'   If Not IsNull(pes!chcarteirarep) Then
'      If CartRep.State = 1 Then
'      CartRep.Close: Set CartRep = Nothing
'    End If
'      CartRep.Open "Select * from Carteira_Rep where chcarteirarep = ('" & pes!chcarteirarep & "'),db,3,3"
'      If CartRep.EOF Then
'         txtRepresentante = "NENHUM"
'      Else
'         txtRepresentante = CartRep!chpessoa
'      End If
'   Else
'      txtRepresentante = "NENHUM"
 '  End If

'   TabTelefoneContato.Seek "=", TabCarteira_Rep("chpessoa")
'   If TabTelefoneContato.NoMatch Then
'      txtTelRep = "Não Informado"
'   Else
'      txtTelRep = TabTelefone("codigocontato")
'   End If
'
'   If txtPromotora = "Nenhum" Then
'      txtPromotora = "Nenhum"
'      txtTelPromot = Empty
'      txtCelPromot = Empty
'   Else
'      TabTelefoneContato.Seek "=", TabCarteira_Promot("chpessoa")
'      If TabTelefoneContato.NoMatch Then
'         txtTelPromot = "Não Informado"
'      Else
'         txtTelPromot = TabTelefone("codigocontato")
'      End If
'
'   End If
End If

If pes!pestipopessoa = 6 Then
   fraLocador.Visible = True
   SalvaLocadora = pes!pesClienteLocador
   SalvaLocaliza = pes!pesUnidadeOperacional
   cmbLocadora = pes!pesClienteLocador
   txtSalario = Format$(pes!salario, "###,##0.00")
   Call Carrega_cmbLocadora
   cmbLocadora = SalvaLocadora
   cmbLocaliza = SalvaLocaliza
Else
   If pes!pestipopessoa = 7 Then
      txtSalario = Format$(pes!salario, "###,##0.00")
   Else
      fraLocador.Visible = False
   End If
End If

'If pes!pesTipoPessoa = 0 Then
'   lblClassFiscal.Visible = True
'   cmbClassFiscal.Visible = True
'   cmbClassFiscal = pes!pesClassFiscal
'End If
   
Call FechaDB

End Sub
Public Sub Rotina_Limpa_Pessoa()
 lblCodPessoa = Empty
 lblRazaoSocial = Empty
 lblCodPessoa1 = Empty
 lblRazaoSocial = Empty
 lblRazaoSocial1 = Empty
 cmbClassFiscal = Empty
  
 txtDataCadastro = "__/__/____"
 txtPesquisa = Empty
 txtDataEncerra = "__/__/____"
 txtCodPessoa = Empty
 'cmbTipoPessoa = Empty
 'cmbPessoa = Empty
 txtGrupo = Empty
 cmbRepresentante = Empty
 cmbPromotora = Empty
 txtRazaoSocial = Empty
 txtFantasia = Empty
 txtAniversario = Empty
 txtCNPJCPF.Mask = " "
 txtCNPJCPF = " "
 txtInscEstIdent = Empty
 txtSalario = Empty
' cmbRamoAtividade = Empty
' cmbCadastroPessoa = Empty
' cmbStatusPessoa = Empty
 
 'OptMatriz.Enabled = True
 'OptFilial = Empty

 'txtPercRapell = Empty
 'txtPerclogistica = Empty
 
 OptCartaSim = False
 OptCartaNao = False
 txtBanco = Empty
 txtAgencia = Empty
 txtConta = Empty
 txtTitular = Empty
 txtCpfTit = Empty
 
 txtEndereco = Empty
 txtBairro = Empty
 txtCidade = Empty
 txtEstado = Empty

 txtCEP = Empty
 
 txtRegiao.Clear
 'txtTelefone = Empty
 'txtRamal = Empty
 'txtFax = Empty
 'txtCelular = Empty
 'txtHomePage = Empty
 'txtEmail = Empty
  
 txtContato = Empty
 txtCargoContato = Empty
 txtTelContato = Empty
 txtCelularContato = Empty
 'txtRepresentante = Empty
 'txtTelRep = Empty
 'txtCelRep = Empty
 'txtPromotora = Empty
 'txtTelPromot = Empty
 'txtCelPromot = Empty
 'cmbFreteIPI = Empty
 'cmbIncideIPI = Empty
 'cmbTabPreco = Empty
 'cmbFreteIncluso = Empty
 'cmbTabFrete = Empty
 'cmbICMS_ST = Empty
 txtConsideracoes = Empty
 fraLocador.Visible = False
 
 GridContato.Rows = 2
 GridContato.TextMatrix(1, 0) = Empty
 GridContato.TextMatrix(1, 1) = Empty
 
 cmbTipoContato = Empty
 txtCodContato = Empty
 
End Sub
Public Sub Rotina_Grava_Pessoa()
On Error GoTo Erro:

If Altera = 0 Then
   pes!chPessoa = txtCodPessoa
Else
   Altera = 0
End If

pes!pestipopessoa = cmbTipoPessoa.ListIndex

If cmbTipoPessoa = "Funcionario" Then
   pes!pesClienteLocador = cmbLocadora
   pes!pesUnidadeOperacional = cmbLocaliza
   pes!salario = txtSalario
Else
   pes!pesClienteLocador = Empty
   pes!pesUnidadeOperacional = Empty
End If

pes!pesPessoa = cmbPessoa.ListIndex

If txtDataCadastro = "__/__/____" Then
   pes!pesDataCadastro = Empty
Else
   pes!pesDataCadastro = txtDataCadastro
End If

pes!pesGrupo = txtGrupo

pes!pesRazaoSocial = txtRazaoSocial
pes!pesFantasia = txtFantasia
pes!pesAniversario = txtAniversario

If cmbTipoPessoa = "FUNCIONARIO" Then
   pes!salario = txtSalario
Else
   pes!salario = 0
End If

If cmbTipoPessoa = "Cliente" Then
   pes!pesClassFiscal = cmbClassFiscal
Else
  pes!pesClassFiscal = Empty
End If

If Not (cmbTipoPessoa.ListIndex = 0) Then
   fim = 1
Else
   CartRep.Open "Select * from Carteira_Rep", db, 3, 3
   
   CartRep.MoveFirst
   fim = 0
  
   Do While fim = 0
      If pes!chcarteirarep = cmbRepresentante Then
         pes!chcarteirarep = cmbRepresentante
         CartRep!chcarteirarep = cmbRepresentante
         fim = 1
      Else
         CartRep.MoveNext
         If CartRep.EOF Then
            fim = 1
            pes!chcarteirarep = cmbRepresentante
            CartRep.AddNew
            CartRep!chcarteirarep = cmbRepresentante
            CartRep!repregiao = "FABRICA"
            CartRep!chPessoa = cmbRepresentante
            CartRep!repordemapresentacao = 11
         End If
      End If
   Loop
   
   If CartPromot.State = 0 Then
      CartPromot.Open "Select * from Carteira_Promot", db, 3, 3
   End If
   CartPromot.MoveFirst
   fim = 0
   
   CartPromot.MoveFirst
   fim = 0
   Do While fim = 0
      If CartPromot!chPessoa = cmbPromotora Then
         pes!chCarteiraPromot = CartPromot!chCarteiraPromot
         fim = 1
      Else
         CartPromot.MoveNext
         If CartPromot.EOF Then
            fim = 1
            CartPromot.AddNew
            CartPromot!chcarteirarep = cmbPromotora
            CartPromot!repregiao = "FABRICA"
            CartPromot!chPessoa = cmbPromotora
            cmbPromotora.SetFocus
            Exit Sub
         End If
      End If
   Loop
End If

pes!chCNPJ_CPF = txtCNPJCPF
If txtInscEstIdent = Empty Then
   pes!pesInscEst_Ident = "Não Informado"
Else
   pes!pesInscEst_Ident = txtInscEstIdent
End If
 
pes!pesRamoAtividade = cmbRamoAtividade.ListIndex
pes!pesCadastroPessoa = cmbCadastroPessoa.ListIndex
pes!pesStatusPessoa = cmbStatusPessoa.ListIndex

If cmbStatusPessoa.ListIndex = 0 Then
   pes!pesDataEncerramento = Empty
Else
   If Not (txtDataEncerra = "__/__/____") Then
      pes!pesDataEncerramento = txtDataEncerra
      pes!pesStatusPessoa = 4
   End If
End If
  
  pes!pesBanco = txtBanco
  pes!pesAgencia = txtAgencia
  pes!pesConta = txtConta
  pes!pesTitular = txtTitular
  pes!pesCPFTitular = txtCpfTit
  
  pes!pesEndereco = txtEndereco
  pes!pesBairro = txtBairro
  pes!pesCidade = txtCidade
  pes!chUF = txtEstado
  pes!pesCEP = txtCEP
  pes!pesRegiao = txtRegiao
  pes!pesRamoAtividade = cmbRamoAtividade.ListIndex
  pes!pesStatusPessoa = cmbStatusPessoa.ListIndex
  pes!pesCadastroPessoa = cmbCadastroPessoa.ListIndex
  pes!pesContato = txtContato
  pes!pesCargoContato = txtCargoContato
  pes!pesTelContato = txtTelContato
  pes!pesCelContato = txtCelularContato
  pes!pesEmail = txtEmail
  If (cmbTipoPessoa = "Funcionario") Or (cmbTipoPessoa = "Colaborador") Then
     fraLocador.Visible = True
     pes!pesClienteLocador = cmbLocadora
     pes!pesUnidadeOperacional = cmbLocaliza
     pes!salario = txtSalario
     'pes!pesTipoPessoa = 7
  Else
     fraLocador.Visible = False
  End If
  
  pes!pesConsideracoes = txtConsideracoes

Exit Sub
Erro:
   MsgBox Err & "=" & Error
   Resume Saida
Saida:

'Call FechaDB

End Sub
Public Sub Rotina_Critica_Cadastro()

Erro_Critica = 0

If cmbTipoPessoa = Empty Then
   MsgBox ("Tipo Pessoa não Informado")
   cmbTipoPessoa.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If

If cmbPessoa = Empty Then
   MsgBox ("Pessoa Física ou Jurídica nao Informado")
   cmbPessoa.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If

If txtDataCadastro = "__/__/____" Then
   MsgBox ("Data de cadastro não informado")
   txtDataCadastro.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If

If txtGrupo = Empty Then
   MsgBox ("Grupo não informado")
   txtGrupo.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If

If txtRazaoSocial = Empty Then
   MsgBox ("Razão Social não informada")
   txtRazaoSocial.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If
 
If txtFantasia = Empty Then
   MsgBox ("Fantasia não informada")
   txtFantasia.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If

If cmbTipoPessoa = "Cliente" Then
   If cmbClassFiscal = "" Then
      MsgBox ("Classificação Fiscal não informada"), vbInformation
      cmbClassFiscal.SetFocus
      Erro_Critica = Erro_Critica + 1
   End If
End If

If txtAniversario = Empty Then
   txtAniversario = "Não Inf."
End If

If cmbRamoAtividade = Empty Then
   MsgBox ("Ramo de atividade não informado")
   cmbRamoAtividade.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If

If cmbCadastroPessoa = Empty Then
   MsgBox ("Vinculação não informado")
   cmbCadastroPessoa.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If

If cmbStatusPessoa = Empty Then
   MsgBox ("Situção do cliente não informado")
   cmbStatusPessoa.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If

If cmbTipoPessoa.ListIndex = 0 Then
   If cmbRepresentante = Empty Then
      MsgBox ("Carteira de Representação não informado")
      cmbRepresentante.SetFocus
      Erro_Critica = Erro_Critica + 1
      Exit Sub
   End If
Else
   cmbRepresentante = Empty
End If

If cmbPessoa.Text = "Jurídica" Then
   If txtCNPJCPF = "__.___.___/____-__" Then
      MsgBox ("CNPJ não Informado")
      txtCNPJCPF.SetFocus
      Erro_Critica = Erro_Critica + 1
      Exit Sub
   End If
Else
   If txtCNPJCPF = "___.___.___-__" Then
      MsgBox ("CPF não Informado")
      txtCNPJCPF.SetFocus
      Erro_Critica = Erro_Critica + 1
      Exit Sub
   End If
End If
 
If cmbPessoa.Text = "Jurídica" Then
   If txtInscEstIdent = Empty Then
      MsgBox ("Insc. Estadual não Informado")
      txtInscEstIdent.SetFocus
      Erro_Critica = Erro_Critica + 1
      Exit Sub
   End If
Else
   If txtInscEstIdent = Empty Then
      txtInscEstIdent = "Não Informado"
   End If
End If

If txtBanco = Empty Then
   txtBanco = "N/Info."
End If
 
If txtAgencia = Empty Then
   txtAgencia = "N/Info."
End If
 
If txtConta = Empty Then
   txtConta = "N/Info."
End If

If txtCpfTit = Empty Then
   txtCpfTit = "N/Info."
End If

If txtTitular = Empty Then
   txtTitular = "N/Info."
End If

If txtEndereco = Empty Then
   MsgBox ("Endereço não Informado")
   txtEndereco.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If
 
If txtBairro = Empty Then
   MsgBox ("Bairro não Informado")
   txtBairro.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If
  
If txtCidade = Empty Then
   MsgBox ("cidade não Informado")
   txtCidade.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If

If txtEstado = Empty Then
   MsgBox ("Estado não Informado")
   txtEstado.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If
 
If txtCEP = Empty Then
   MsgBox ("CEP não Informado")
   txtCEP.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If

If txtBairro = Empty Then
   MsgBox ("Bairro não Informado")
   txtBairro.SetFocus
   Erro_Critica = Erro_Critica + 1
   Exit Sub
End If
 
If cmbTipoPessoa.ListIndex = 0 Or 2 Or 3 Or 4 Then
   If txtRegiao = Empty Then
      MsgBox ("Região não informada")
      txtRegiao.SetFocus
      Erro_Critica = Erro_Critica + 1
      Exit Sub
   End If
Else
   If txtRegiao = Empty Then
      txtRegiao = " "
   End If
End If
 
If cmbTipoPessoa.ListIndex = 0 Or cmbTipoPessoa.ListIndex = 1 Then
   If txtContato = Empty Then
      MsgBox ("Contato não Informado")
      txtContato.SetFocus
      Erro_Critica = Erro_Critica + 1
      Exit Sub
    End If
Else
    If txtContato = Empty Then
       txtContato = "Não Informado"
    End If
End If

If txtCargoContato = Empty Then
   txtCargoContato = "Não Informado"
End If
  
If txtTelContato = Empty Then
   txtTelContato = "Não Informado"
End If
 
If txtCelularContato = Empty Then
    txtCelularContato = "Não Informado"
End If
 
If txtConsideracoes = Empty Then
   txtConsideracoes = " "
End If

If cmbTipoPessoa = "Funcionario" Then
   If txtSalario = Empty Then
      MsgBox ("Informar Salario Mensal do Funcionario. Última aba."), vbInformation
      Erro_Critica = Erro_Critica + 1
   End If
End If
 
'Call FechaDB

End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtEstado_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
Private Sub txtEstado_LostFocus()

Call Rotina_AbrirBanco

ICM.Open "Select * from ICMS where chUF = ('" & txtEstado & "')", db, 3, 3

If ICM.EOF Then
   MsgBox ("Estado não cadastrado na Tabela de ICMS"), vbAbortRetryIgnore
   cmdSair.SetFocus
End If

UfRegiao.Open "Select * from Regiao where RegUF = ('" & txtEstado & "')", db, 3, 3

If UfRegiao.EOF Then
   MsgBox ("Regiao não cadastrada para essa UF."), vbInformation
   txtRegiao.AddItem "Região não cadastrada"
Else
   UfRegiao.MoveFirst
   Do While Not UfRegiao.EOF
      txtRegiao.AddItem UfRegiao!regregiao
      UfRegiao.MoveNext
   Loop
End If

Call FechaDB

End Sub

Private Sub txtFantasia_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtGrupo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtPesquisa_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtRazaoSocial_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


Private Sub txtRegiao_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
Public Sub Rotina_Carga_Contato()

Dim PessoaContato As String
Dim ContatoTipo As String
Dim PrimeiroRegistro As Integer

PrimeiroRegistro = 0

GridContato.Rows = 1

cmbTipoContato = Empty
txtCodContato = Empty


Call Rotina_AbrirBanco

Contato.Open "Select * from Telefone where codPessoa = ('" & txtCodPessoa & "')", db, 3, 3

If Contato.EOF Then
   Exit Sub
Else
   Ind = 0
   IndContato = 0
   PessoaContato = Contato!codpessoa
   Contato.MoveFirst
   Do While txtCodPessoa = PessoaContato
      Ind = Ind + 1
      IndContato = IndContato + 1
      GridContato.Rows = Ind + 1
      GridContato.TextMatrix(Ind, 0) = Contato!TipoContato
      GridContato.TextMatrix(Ind, 1) = Contato!CodigoContato
      Contato.MoveNext
   
      If Contato.EOF Then
         PessoaContato = Empty
      Else
         If Not Contato!codpessoa = txtCodPessoa Then
            PessoaContato = Empty
         End If
      End If
   Loop
End If

'Call FechaDB

End Sub

Public Sub Exclui_Contato()
Dim PessoaAnterior As String
'Data.Recordset.FindFirst "CodPessoa = '" & txtCodPessoa & "'"''

Contato.Open "Select * from Telefone where codpessoa = ('" & txtCodPessoa & "')", db, 3, 3
If Contato.EOF Then
   MsgBox ("Cliente sem contato cadstrado"), vbInformation
   IndContato = 0
Else
   IndContato = 1
   Pessoa = Contato!codpessoa
   fim = 0
End If
Contato.MoveFirst
If IndContato = 1 Then
   Ind = 0
   Do While Contato!codpessoa = Pessoa
      Contato.Delete
      Contato.MoveNext
      If Contato.EOF Then
         Pessoa = "FIM"
      End If
   Loop
End If
End Sub

Private Sub frmPessoa_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'TabUsuario.Close
End Sub

Public Sub GeraProduto()

Dim ChaveFinal As String
Dim ChaveProduto As String

Call Rotina_AbrirBanco

'ChaveFinal = "-E"
ChaveProduto = txtCodPessoa
Prod.Open "Select * from Produto where chProduto = ('" & ChaveProduto & "')", db, 3, 3
If Prod.EOF Then
   Prod.AddNew
End If

Prod!chProduto = txtCodPessoa
Prod!prdNomeProd = txtFantasia
Prod!prdfabricante = 0
'Prod!prdAtividade = "EMBARCADO"
Prod!prdtipo = 2
Prod!prdgrupo = Empty
Prod!prdLocadora = cmbLocadora
Prod!prdUnidadeOperacional = cmbLocaliza
Prod!prdunidade = 1
Prod!prdDescCompleta = txtRazaoSocial
Prod!prdIPI = 0
Prod!prdOrdemApresentacao = 1
Prod!prdComissao = 0
Prod!prdGrupoComppreco = """"
Prod.Update

Call FechaDB

End Sub


