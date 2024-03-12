VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProposta 
   Caption         =   "frmProposta"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGeraProposta 
      Caption         =   "Imprimir Proposta"
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
      Left            =   17760
      TabIndex        =   25
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Frame fraDetProp 
      Caption         =   "Detalhe da Proposta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   120
      TabIndex        =   43
      Top             =   3480
      Width           =   16695
      Begin VB.Frame fraMedidaHabitat 
         Caption         =   "Medidas do Habitat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   3360
         TabIndex        =   51
         Top             =   360
         Width           =   4815
         Begin VB.TextBox txtCompHbt 
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
            Height          =   375
            Left            =   600
            TabIndex        =   11
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtLarguraHBT 
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
            Height          =   375
            Left            =   1800
            TabIndex        =   12
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtAlturaHBT 
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
            Height          =   375
            Left            =   3000
            TabIndex        =   13
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Comp."
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
            Left            =   480
            TabIndex        =   54
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label26 
            Caption         =   "Larg."
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
            Left            =   1800
            TabIndex        =   53
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label27 
            Caption         =   "Alt."
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
            Left            =   3000
            TabIndex        =   52
            Top             =   360
            Width           =   615
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdDetProp 
         Height          =   2655
         Left            =   3360
         TabIndex        =   20
         Top             =   2760
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         FormatString    =   "Qtd  |Equipto/Operador                         |Unid.|P.U         |Qtd.Unid|Valor diária       "
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
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "Excluir"
         Height          =   495
         Left            =   15480
         TabIndex        =   22
         Top             =   3360
         Width           =   855
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
         Left            =   15480
         TabIndex        =   21
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtDiaria 
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
         Height          =   415
         Left            =   12600
         TabIndex        =   19
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txtQtdUnid 
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
         Height          =   415
         Left            =   11280
         TabIndex        =   18
         Top             =   2400
         Width           =   1335
      End
      Begin VB.ComboBox cmbUnidade 
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
         Left            =   9000
         TabIndex        =   16
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtPrecoUnit 
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
         Height          =   415
         Left            =   9840
         TabIndex        =   17
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtEquipOper 
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
         Height          =   415
         Left            =   4080
         TabIndex        =   15
         Top             =   2400
         Width           =   4935
      End
      Begin VB.TextBox txtQtd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   3360
         TabIndex        =   14
         Top             =   2400
         Width           =   615
      End
      Begin VB.ListBox lstEquipamento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5160
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Equipamento/Operador"
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
         Left            =   0
         TabIndex        =   50
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label txtValoDiaria 
         Caption         =   "Valor Diaria"
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
         Left            =   12600
         TabIndex        =   49
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Qtd. Unid"
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
         Left            =   11280
         TabIndex        =   48
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "P.Unit"
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
         Left            =   9840
         TabIndex        =   47
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Unid."
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
         Left            =   9000
         TabIndex        =   46
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Equip/Operador"
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
         Left            =   4080
         TabIndex        =   45
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label Label7 
         Caption         =   "Qtd."
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
         Left            =   3360
         TabIndex        =   44
         Top             =   2040
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbAno 
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
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox cmbEmail 
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
      Left            =   8280
      TabIndex        =   7
      Top             =   2520
      Width           =   5895
   End
   Begin VB.ComboBox cmbResponsavel 
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
      Left            =   1560
      TabIndex        =   6
      Top             =   2520
      Width           =   5655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Período da Proposta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   14400
      TabIndex        =   38
      Top             =   1920
      Width           =   4935
      Begin VB.TextBox txtQtdDias 
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
         Height          =   525
         Left            =   720
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cmbUnidTemp 
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
         Left            =   2520
         TabIndex        =   10
         Text            =   "Dias"
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Período"
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
         Left            =   600
         TabIndex        =   40
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Unid.tempo"
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
         Left            =   2520
         TabIndex        =   39
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComCtl2.DTPicker dtDataPedidoCotacao 
      Height          =   495
      Left            =   11520
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
      Format          =   242941953
      CurrentDate     =   45050
   End
   Begin VB.CommandButton cmdProcessar 
      BackColor       =   &H00FFFF80&
      Caption         =   "Processar"
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
      Left            =   17760
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
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
      Height          =   735
      Left            =   17760
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7200
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtDataRevisao 
      Height          =   495
      Left            =   14040
      TabIndex        =   5
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
      Format          =   242941953
      CurrentDate     =   45047
   End
   Begin VB.ComboBox cmbStatus 
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
      Left            =   16920
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   5040
      Width           =   3375
   End
   Begin VB.ComboBox cmbRevisao 
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
      Left            =   8640
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox cmbCliente 
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
      Left            =   4920
      TabIndex        =   2
      Top             =   1440
      Width           =   3255
   End
   Begin VB.ComboBox cmbProposta 
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
      Width           =   2535
   End
   Begin VB.Label Label25 
      Caption         =   "Ano Prop."
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
      Left            =   3000
      TabIndex        =   42
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label28 
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
      Left            =   14160
      TabIndex        =   41
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label19 
      Caption         =   "Data Solicitação"
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
      Left            =   11400
      TabIndex        =   37
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label17 
      Caption         =   "Email"
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
      Left            =   8280
      TabIndex        =   36
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label lblDataPadrao 
      Alignment       =   2  'Center
      Caption         =   "Label17"
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
      Left            =   14160
      TabIndex        =   35
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      Height          =   135
      Left            =   7320
      TabIndex        =   34
      Top             =   840
      Width           =   15
   End
   Begin VB.Label Label12 
      Caption         =   "Data da Revisão"
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
      TabIndex        =   33
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label11 
      Caption         =   "Status da Proposta"
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
      Left            =   16920
      TabIndex        =   32
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Revisão"
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
      Left            =   8640
      TabIndex        =   31
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Responsável"
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
      Left            =   1560
      TabIndex        =   30
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
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
      Left            =   4920
      TabIndex        =   29
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Proposta"
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
      TabIndex        =   28
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Proposta de Locações e Serviços"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frmProposta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ind As Integer
Dim Linha As Integer
Dim Limite As Integer
Dim MedidaEquipamento As Integer
Dim chaveRevisao As Integer
Dim NovaProposta As Integer
Dim ContatoSHB As String
Dim ContatoEmail As String
Dim ContatoTel As String
Dim ano As String
Dim AnoCMB As Integer
Dim flagAlteracao As Boolean
Dim flagNew As Boolean

'Imagem do Registro de entrada

Dim proposta As String
Dim revisao As String
Dim Cliente As String
Dim DataPedidoCotacao As Date
Dim DataRevisao As Date
Dim DataDePrevisaoDeEntrega As Date
Dim QtdDias As String
Dim UnidTemp As String
Dim responsavel As String
Dim Email As String
Dim Status As String
Dim CNPJ As String


Private Sub cmbAno_LostFocus()

Call Rotina_AbrirBanco

If cmbProposta = "Nova proposta" Then
   dtDataPedidoCotacao = Date
   dtDataRevisao = Date
   
   ano = Date
   
   ano = Format$(ano, "yy")
   AnoCMB = ano
   Limite = 21
   
   cmbAno.Clear
   
   Do While Not AnoCMB = Limite
      cmbAno.AddItem AnoCMB
      AnoCMB = AnoCMB - 1
   Loop
   
   cmbAno.ListIndex = 0
   NovaProposta = 1
   cmbRevisao.AddItem "Nova Revisao"
   Exit Sub
End If

ano = cmbAno

'Prod.Open "SELECT * FROM proposta WHERE revisao = (SELECT MAX(revisao) FROM proposta WHERE numProposta = ('" & cmbProposta & "')) AND numProposta = ('" & cmbProposta & "')", db, 3, 3
Prod.Open "Select * from proposta WHERE anoProposta = ('" & ano & "') and numProposta = ('" & cmbProposta & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("ERRO: Contrato pesquisado não encontrado."), vbCritical
   Call FechaDB
   Exit Sub
End If

cmbCliente = Prod!Cliente

cmbRevisao.Clear

cmbRevisao.AddItem "Nova Revisao"

Prod.MoveFirst

Do While Not Prod.EOF
   cmbRevisao.AddItem Prod!revisao
   Prod.MoveNext
Loop
   
Call FechaDB
End Sub

Private Sub cmbCliente_LostFocus()
   cmbResponsavel.Clear
   Call Rotina_AbrirBanco
   rs.Open "Select *  from telefone where CodPessoa=('" & cmbCliente & "')", db, 3, 3
   If rs.EOF Then
      MsgBox ("Erro:Cliente não informado ou inexistente")
      Call FechaDB
      Exit Sub
   End If
   
   rs.MoveFirst
   Do While Not rs.EOF
      cmbResponsavel.AddItem rs!TipoContato
      rs.MoveNext
   Loop
   
   rs.Close
   Call FechaDB
End Sub
Private Sub cmbProposta_LostFocus()

Call Rotina_AbrirBanco

If cmbProposta = "Nova proposta" Then
   dtDataPedidoCotacao = Date
   dtDataRevisao = Date
   
   ano = Date
   
   ano = Format$(ano, "yy")
   AnoCMB = ano
   Limite = 21
   
   cmbAno.Clear
   
   Do While Not AnoCMB = Limite
      cmbAno.AddItem AnoCMB
      AnoCMB = AnoCMB - 1
   Loop
   
   cmbAno.ListIndex = 0
   NovaProposta = 1
   cmbRevisao.Clear
   cmbRevisao.AddItem "Nova Revisao"
   cmbRevisao.ListIndex = 0
   Exit Sub
Else
   NovaProposta = 0

   Prod.Open "Select * from proposta where numProposta = ('" & cmbProposta & "')", db, 3, 3
   If Prod.EOF Then
      MsgBox ("ERRO: Não encontrado o numero de proposta informado."), vbInformation
      Call FechaDB
      Exit Sub
   End If
      
   Prod.MoveFirst
   
   cmbAno.Clear
   
   Do While Not Prod.EOF
      cmbAno = Prod!anoProposta
      Prod.MoveNext
   Loop
   fraDetProp.Visible = True
End If

Call FechaDB

End Sub

Private Sub cmbResponsavel_LostFocus()
Call Rotina_AbrirBanco
cmbEmail.Clear
pes.Open "Select CodigoContato from telefone where TipoContato=('" & cmbResponsavel & "')", db, 3, 3
   If pes.EOF Then
      MsgBox ("Erro:Contato não informado ou inexistente")
      Call FechaDB
      Exit Sub
   End If
   
   pes.MoveFirst
   Do While Not pes.EOF
      cmbEmail.AddItem pes!CodigoContato
      pes.MoveNext
   Loop
   
   pes.Close
   
Call FechaDB
End Sub



Private Sub cmbRevisao_LostFocus()

flagAlteracao = False

If cmbRevisao = "Nova Revisao" Then
   flagNew = True
   Exit Sub
End If

Call Rotina_AbrirBanco

Prod.Open "Select * from proposta where anoProposta = ('" & ano & "') and numProposta = ('" & cmbProposta & "' ) and Cliente = ('" & cmbCliente & "') and revisao = ('" & cmbRevisao & "')", db, 3, 3
If Prod.EOF Then
   MsgBox ("ERRO: proposta não encontrada. Comunicar ao analista responsável."), vbInformation
   Call FechaDB
   Exit Sub
End If

cmbProposta = Prod!numProposta
proposta = Prod!numProposta
cmbRevisao = Prod!revisao
revisao = Prod!revisao
cmbCliente = Prod!Cliente
Cliente = Prod!Cliente
dtDataPedidoCotacao = Prod!dataSolicitacaoProposta
DataPedidoCotacao = Prod!dataSolicitacaoProposta
dtDataRevisao = Prod!dataProposta
DataRevisao = Prod!dataProposta
txtQtdDias = Prod!QtdDias
QtdDias = Prod!QtdDias
cmbUnidTemp.ListIndex = Prod!unidTempo
UnidTemp = cmbUnidTemp
cmbResponsavel = Prod!responsavel
responsavel = Prod!responsavel
cmbEmail = Prod!emailResp
Email = Prod!emailResp
cmbStatus.ListIndex = Prod!Status
Status = cmbStatus

pes.Open "Select chCNPJ_CPF from pessoa where chPessoa = ('" & cmbCliente & "')", db, 3, 3

If pes.EOF Then
   MsgBox ("Cliente não possui CNPJ cadastrado")
Else
   CNPJ = pes!chCNPJ_CPF
End If

Call CargaGrid

End Sub
Private Sub cmdExcluir_Click()
Dim Resp As String

Call Rotina_AbrirBanco

If txtEquipOper <> Empty Then
   rs.Open "Select * from propostadetalhe where numProposta=('" & cmbProposta & "') and revisaoProposta=('" & cmbRevisao & "') and equipamento = ('" & txtEquipOper & "')", db, 3, 3
   If Not rs.EOF Then
      Resp = MsgBox("Exclusão de registro. Confirma???", vbExclamation + vbYesNo)
      If Resp = vbYes Then
         rs.Delete
         grdDetProp.RemoveItem (grdDetProp.Row)
         Call limpaCamposDetalheProposta
         Call CargaGrid
         MsgBox ("Registro excluído com sucesso."), vbInformation
      End If
   Else
      grdDetProp.Rows = 1
   End If
Else
   MsgBox ("Informar equipamento a ser removido"), vbInformation
End If

End Sub

Private Sub cmdGeraProposta_Click()
   Call verificaAlteracao
   If cmbProposta <> "Nova proposta" And cmbRevisao <> "Nova Revisão" And flagAlteracao = False Then
   
   Call GerarExcelWord
   Call ExportarWord
   
   Unload Me
   Else
      MsgBox ("Proposta não Processada")
   End If
End Sub

Private Sub cmdProcessar_Click()

   If cmbRevisao <> Empty Then
      
      On Error GoTo Erro:
      
      Call Rotina_AbrirBanco
      
      If cmbProposta = "Nova proposta" Then
         rs.Open "Select * from empresa", db, 3, 3
         If rs.EOF Then
            MsgBox ("ERRO: Registro empresa não encontrado."), vbCritical
            Call FechaDB
            Exit Sub
         End If
         rs!empNumProposta = rs!empNumProposta + 1
         cmbProposta = rs!empNumProposta
         rs.Update
         cmbRevisao = 1
      Else
         If cmbStatus.ListIndex = 0 Then
            If cmbRevisao = "Nova Revisao" Then
               chaveRevisao = 1
               cmbRevisao = cmbRevisao.ListCount
            Else
               chaveRevisao = 0
            End If
         Else
            chaveRevisao = 0
         End If
      End If
      
   '   If chaveRevisao = 1 And cmbProposta <> "Nova proposta" Then
   '      neg.Open "Select * from proposta where numProposta = ('" & cmbProposta & "') and revisao = (SELECT MAX(revisao) FROM proposta WHERE numProposta = ('" & cmbProposta & "'))", db, 3, 3
   '      If neg.EOF Then
   '         MsgBox ("Erro: revisão não encontrada")
   '         Exit Sub
   '      End If
   '      neg!Status = 2
   '      neg.Update
   '   End If
      
      Prod.Open "Select * from proposta where anoProposta = ('" & ano & "') and numProposta = ('" & cmbProposta & "') and revisao = ('" & cmbRevisao & "')", db, 3, 3
      If Prod.EOF Then
         Prod.AddNew
      End If
      
      Prod!numProposta = Format$(cmbProposta, "0000")
      cmbProposta = Format$(cmbProposta, "0000")
      Prod!revisao = cmbRevisao
      cmbRevisao = cmbRevisao
      Prod!anoProposta = cmbAno
      Prod!Cliente = cmbCliente
      Prod!responsavel = cmbResponsavel
      Prod!emailResp = cmbEmail
      Prod!unidTempo = cmbUnidTemp.ListIndex
   '   Prod!qtdHBT = txtQtdHBT
   '   Prod!qtdJBX = txtQtdJBX
      Prod!QtdDias = txtQtdDias
   '   Prod!qtdFunc = txtQtdOperador
   '   Prod!valorFunc = txtValorOperador
   '   Prod!valorHBTMetro = txtValorHabitat
   '   Prod!qtdMetros = ((txtCompHbt * txtLarguraHBT) * 2) + ((txtCompHbt * txtAlturaHBT) * 2) + ((txtAlturaHBT * txtLarguraHBT) * 2) + 8
   '   Prod!diaria = lblValorDiaria
   '   Prod!valorJBX = txtValorJBX
      Prod!Status = cmbStatus.ListIndex
      Prod!dataSolicitacaoProposta = dtDataPedidoCotacao
      Prod!dataProposta = dtDataRevisao
      Prod.Update
      
      Call FechaDB
      If grdDetProp.Visible = False Then
         MsgBox ("Cadastrar os equipamentos."), vbInformation
         fraDetProp.Visible = True
      End If
         
      MsgBox ("Proposta processada com sucesso"), vbInformation
   Else
      MsgBox ("Revisão não informada!"), vbInformation
   End If
Exit Sub
Erro: MsgBox ("Erro ao processar proposta: " & Err.Description), vbInformation
End Sub


Public Sub CargaGrid()
   Call Rotina_AbrirBanco
   rs.Open "Select * from propostadetalhe where numProposta=('" & cmbProposta & "') and revisaoProposta=('" & cmbRevisao & "')", db, 3, 3
   If Not rs.EOF Then
      rs.MoveFirst
      Linha = 1
      'grdDetProp.Clear
      Do While Not rs.EOF
         grdDetProp.Rows = Linha + 1
         grdDetProp.TextMatrix(Linha, 0) = rs!quantidade
         grdDetProp.TextMatrix(Linha, 1) = rs!equipamento
         grdDetProp.TextMatrix(Linha, 2) = rs!unidade
         grdDetProp.TextMatrix(Linha, 3) = Format$(rs!precoUnit, "##,##0.00")
         grdDetProp.TextMatrix(Linha, 4) = rs!areaTotal
         grdDetProp.TextMatrix(Linha, 5) = Format$(rs!diaria, "##,#0.00")
         If rs!equipamento = "Habitat" Then
            txtAlturaHBT = rs!altura
            txtCompHbt = rs!comprimento
            txtLarguraHBT = rs!largura
         End If
         Linha = Linha + 1
         rs.MoveNext
      Loop
      
   End If
   'FechaDB
   
End Sub

Private Sub cmdSalvar_Click()
   If txtQtd <> Empty And txtEquipOper <> Empty And txtPrecoUnit <> Empty And txtQtdUnid <> Empty And txtDiaria <> Empty Then
      Call Rotina_AbrirBanco
      rs.Open "Select * from propostadetalhe where numProposta=('" & cmbProposta & "') and revisaoProposta=('" & cmbRevisao & "') and equipamento = ('" & txtEquipOper & "')", db, 3, 3
      If rs.EOF Then
         rs.AddNew
      End If
      rs!numProposta = cmbProposta
      rs!revisaoProposta = cmbRevisao
      rs!equipamento = txtEquipOper
      rs!precoUnit = txtPrecoUnit
      rs!areaTotal = txtQtdUnid
      rs!diaria = txtDiaria
      rs!quantidade = txtQtd
      rs!unidade = cmbUnidade
      If fraMedidaHabitat.Visible = True Then
         rs!comprimento = txtCompHbt
         rs!largura = txtLarguraHBT
         rs!altura = txtAlturaHBT
         rs!dimensoes = txtCompHbt & "X" & txtLarguraHBT & "X" & txtAlturaHBT
      Else
         rs!dimensoes = "-"
      End If
      rs.Update
      FechaDB
      Call limpaCamposDetalheProposta
      Call CargaGrid
   Else
      MsgBox ("Verificar campos informados"), vbInformation
   End If
   
End Sub

Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
txtAlturaHBT = 0
txtCompHbt = 0
txtLarguraHBT = 0

dtDataPedidoCotacao = Date
dtDataRevisao = Date

ano = Date
Linha = 1
fraDetProp.Visible = False
ano = Format$(ano, "yy")
AnoCMB = ano
Limite = 21

cmbAno.Clear

Do While Not AnoCMB = Limite
   cmbAno.AddItem AnoCMB
   AnoCMB = AnoCMB - 1
Loop

cmbAno.ListIndex = 0

Call Rotina_AbrirBanco

usu.Open "Select * FROM usuario WHERE chNome = ('" & glbUsuario & "')", db, 3, 3

If usu.EOF Then
   MsgBox ("ERRO: usuario não encontrado."), vbCritical
   Call FechaDB
   Exit Sub
End If
   
pes.Open "Select * FROM pessoa WHERE chPessoa = ('" & usu!chPessoa & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("ERRO: pessoa sem cliente."), vbCritical
Else
   ContatoSHB = pes!pesRazaoSocial
   ContatoEmail = pes!pesEmail
   ContatoTel = pes!pesCelContato
End If

If pes.State = 1 Then
   pes.Close: Set pes = Nothing
End If

pes.Open "Select * FROM pessoa WHERE pesTipoPessoa = ('" & 0 & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("ERRO: pessoa sem cliente."), vbCritical
Else
   pes.MoveFirst
   Do While Not pes.EOF
      cmbCliente.AddItem pes!chPessoa
      pes.MoveNext
   Loop
End If

lblDataPadrao = Date


cmbProposta.AddItem "Nova proposta"
cmbProposta.ListIndex = 0


cmbUnidTemp.AddItem "Dias"
cmbUnidTemp.AddItem "Meses"
cmbUnidTemp.AddItem "Anos"

cmbUnidTemp.ListIndex = 0

cmbRevisao.AddItem "Nova Revisao"
cmbRevisao.ListIndex = 0

neg.Open "Select DISTINCT numProposta from proposta where status <> 1", db, 3, 3
If Not neg.EOF Then
   neg.MoveFirst
   Do While Not neg.EOF
      cmbProposta.AddItem neg!numProposta
      neg.MoveNext
   Loop
End If

cmbStatus.AddItem "EM ANÁLISE"
cmbStatus.AddItem "PROPOSTA ACEITA"
cmbStatus.AddItem "PROPOSTA RECUSADA"

'cmbStatus = Empty

cmbUnidade.AddItem "M²"
cmbUnidade.AddItem "Unid"

Prod.Open "Select DISTINCT eqptTipoEquipamento from equipamento", db, 3, 3
If Prod.EOF Then
   MsgBox ("Erro: equipamentos não registrados")
   FechaDB
   Exit Sub
End If

lstEquipamento.Clear
lstEquipamento.AddItem "Habitat"
lstEquipamento.AddItem "Operador"

Do While Not Prod.EOF
   lstEquipamento.AddItem Prod!eqptTipoEquipamento
   Prod.MoveNext
Loop

Prod.Close

rs.Open "Select * from empresa", db, 3, 3
If rs.EOF Then
   MsgBox ("Erro: empresa não possui registro")
   FechaDB
   Exit Sub
End If

If rs!empAnoProposta < ano Then
   rs!empAnoProposta = ano
   rs!empNumProposta = 0
   rs.Update
End If

rs.Close
FechaDB

End Sub

Public Sub ExportarWord()
        On Error GoTo Erro
        Dim CaminhoNew As String
        Dim CaminhoProposta As String
        Call Rotina_AbrirBanco
        usu.Open "select usuEnderecoOneDrive from usuario where  chnome = ('" & glbUsuario & "')", db, 3, 3
        
        If usu!usuEnderecoOneDrive = Null Then
         MsgBox ("Não autorizada a impressão de proposta")
         FechaDB
         Exit Sub
        Else
         CaminhoNew = usu!usuEnderecoOneDrive & "Sistema\PROPOSTA MODELO\"
         CaminhoProposta = usu!usuEnderecoOneDrive & "Sistema\PROPOSTA LOCAÇÃO\"
        End If
        Dim wordObj As Word.Application
        Dim arqProp As Word.Document
        Dim conteudoDoc As Word.Selection
        Dim marcacaoWord As Word.Range
        Dim excelObj As Excel.Application
        Dim excelTabela As Excel.Workbook
        Dim excelCelula As Excel.Worksheet
        Dim localExportar As Excel.Range
           
        Set wordObj = CreateObject("Word.Application")
        Set wordObj = CreateObject("Word.Application")
        Set excelObj = CreateObject("Excel.Application")
        wordObj.Visible = True
        
        Set arqProp = wordObj.Documents.Open(CaminhoNew & "ModelWord.docx")
        Set conteudoDoc = arqProp.Application.Selection
        Set marcacaoWord = arqProp.Bookmarks("TABELA").Range
        Set excelTabela = excelObj.Workbooks.Open(CaminhoNew & "ExcelDetEquip.xlsx")
        Set excelCelula = excelTabela.Worksheets("Planilha1")
        Set localExportar = excelCelula.Range("A1:G10")
        
        
        conteudoDoc.Find.Text = "#DESTINATARIO"
        conteudoDoc.Find.Replacement.Text = cmbResponsavel
        conteudoDoc.Find.Execute Replace:=wdReplaceAll
        
        conteudoDoc.Find.Text = "#CLIENTE_CONTRATO"
        conteudoDoc.Find.Replacement.Text = cmbCliente
        conteudoDoc.Find.Execute Replace:=wdReplaceAll
        
        conteudoDoc.Find.Text = "#CNPJ"
        conteudoDoc.Find.Replacement.Text = CNPJ
        conteudoDoc.Find.Execute Replace:=wdReplaceAll
        
        
'        conteudoDoc.Find.Text = "#TOTAL_METROS"
'        conteudoDoc.Find.Replacement.Text = lblTotalMetros
'        conteudoDoc.Find.Execute Replace:=wdReplaceAll

'        conteudoDoc.Find.Text = "#QTD_OPERADOR"
'        conteudoDoc.Find.Replacement.Text = txtQtdOperador
'        conteudoDoc.Find.Execute Replace:=wdReplaceAll

'        conteudoDoc.Find.Text = "#VALOROPERADOR"
'        conteudoDoc.Find.Replacement.Text = Format$(txtValorOperador, "##,##0.00")
'        conteudoDoc.Find.Execute Replace:=wdReplaceAll

'        conteudoDoc.Find.Text = "#VALOROTOTALPERADOR"
'        conteudoDoc.Find.Replacement.Text = Format$((txtValorOperador * txtQtdOperador), "##,##0.00")
'        conteudoDoc.Find.Execute Replace:=wdReplaceAll

        conteudoDoc.Find.Text = "#TOTAL_DIAS"
        conteudoDoc.Find.Replacement.Text = txtQtdDias
        conteudoDoc.Find.Execute Replace:=wdReplaceAll
        

        conteudoDoc.Find.Text = "#UNIDTEMPO"
        conteudoDoc.Find.Replacement.Text = cmbUnidTemp
        conteudoDoc.Find.Execute Replace:=wdReplaceAll

'        conteudoDoc.Find.Text = "#ValorHabitat"
'        conteudoDoc.Find.Replacement.Text = Format$(txtValorHabitat, "##,##0.00")
'        conteudoDoc.Find.Execute Replace:=wdReplaceAll

'        conteudoDoc.Find.Text = "#VALORTOTALHABITAT"
'        conteudoDoc.Find.Replacement.Text = Format$((txtValorHabitat * MedidaEquipamento), "##,##0.00")
'        conteudoDoc.Find.Execute Replace:=wdReplaceAll

'        conteudoDoc.Find.Text = "#ValorJBX"
'        conteudoDoc.Find.Replacement.Text = Format$(txtValorJBX, "##,##0.00")
'        conteudoDoc.Find.Execute Replace:=wdReplaceAll

'        conteudoDoc.Find.Text = "#ValorTOTJBX"
'        conteudoDoc.Find.Replacement.Text = Format$((txtValorJBX * txtQtdJBX), "##,##0.00")
'        conteudoDoc.Find.Execute Replace:=wdReplaceAll

        conteudoDoc.Find.Text = "#TOTAL_DIAS"
        conteudoDoc.Find.Replacement.Text = txtQtdDias
        conteudoDoc.Find.Execute Replace:=wdReplaceAll

        conteudoDoc.Find.Text = "#DATAEXTENSO"
        conteudoDoc.Find.Replacement.Text = Format$(Date, "dd \d\e MMMM \d\e yyyy")
        conteudoDoc.Find.Execute Replace:=wdReplaceAll

        conteudoDoc.Find.Text = "#DATAPADRAO"
        conteudoDoc.Find.Replacement.Text = Date
        conteudoDoc.Find.Execute Replace:=wdReplaceAll

        conteudoDoc.Find.Text = "#PROPOSTA"
        conteudoDoc.Find.Replacement.Text = cmbProposta
        conteudoDoc.Find.Execute Replace:=wdReplaceAll

        conteudoDoc.Find.Text = "#REVISAO"
        conteudoDoc.Find.Replacement.Text = cmbRevisao
        conteudoDoc.Find.Execute Replace:=wdReplaceAll

        conteudoDoc.Find.Text = "#DATAPEDIDOCOTACAO"
        conteudoDoc.Find.Replacement.Text = dtDataPedidoCotacao
        conteudoDoc.Find.Execute Replace:=wdReplaceAll

        conteudoDoc.Find.Text = "#EMAILSOLICITANTE"
        conteudoDoc.Find.Replacement.Text = cmbEmail
        conteudoDoc.Find.Execute Replace:=wdReplaceAll

        conteudoDoc.Find.Text = "#ANO"
        conteudoDoc.Find.Replacement.Text = Format$(Date, "yy")
        conteudoDoc.Find.Execute Replace:=wdReplaceAll

        conteudoDoc.Find.Text = "#STATUSPROPOSTA"
        conteudoDoc.Find.Replacement.Text = cmbStatus
        conteudoDoc.Find.Execute Replace:=wdReplaceAll

'        conteudoDoc.Find.Text = "#QTDHBT"
'        conteudoDoc.Find.Replacement.Text = txtQtdHBT
'        conteudoDoc.Find.Execute Replace:=wdReplaceAll

'        conteudoDoc.Find.Text = "#QTDJBX"
'        conteudoDoc.Find.Replacement.Text = txtQtdJBX
'        conteudoDoc.Find.Execute Replace:=wdReplaceAll

        conteudoDoc.Find.Text = "#CONTATOSHB"
        conteudoDoc.Find.Replacement.Text = ContatoSHB
        conteudoDoc.Find.Execute Replace:=wdReplaceAll

        conteudoDoc.Find.Text = "#EMAILSHB"
        conteudoDoc.Find.Replacement.Text = ContatoEmail
        conteudoDoc.Find.Execute Replace:=wdReplaceAll

        conteudoDoc.Find.Text = "#TELEFONECONTATOSHB"
        conteudoDoc.Find.Replacement.Text = ContatoTel
        conteudoDoc.Find.Execute Replace:=wdReplaceAll
        
        localExportar.Copy
        
        With marcacaoWord
            .Select
            .PasteSpecial Link:=False, DataType:=wdPasteRTF, Placement:=wdInLine, DisplayAsIcon:=False
        End With
        
        conteudoDoc.Find.Text = "#DIMENSOES"
        conteudoDoc.Find.Replacement.Text = ((txtAlturaHBT * txtLarguraHBT) * 2 + (txtCompHbt * txtAlturaHBT) * 2 + (txtCompHbt * txtLarguraHBT) * 2 + 8)
        conteudoDoc.Find.Execute Replace:=wdReplaceAll
        
        conteudoDoc.Find.Text = "#MEDIDA"
        conteudoDoc.Find.Replacement.Text = txtCompHbt & "X" & txtLarguraHBT & "X" & txtAlturaHBT
        conteudoDoc.Find.Execute Replace:=wdReplaceAll
        
           
        If chaveRevisao = 0 Then
            arqProp.SaveAs (CaminhoProposta & ano & "-" & "Prop-" & (Format$(cmbProposta, "0000")) & "-Rev-" & (Format$(cmbRevisao, "00")) & "-Cli-" & cmbCliente & ".docx")
        Else
            arqProp.SaveAs (CaminhoProposta & ano & "-" & "Prop-" & (Format$(cmbProposta, "0000")) & "-Rev-" & (Format$(cmbRevisao.ListCount, "00")) & "-Cli-" & cmbCliente & ".docx")
        End If
510     arqProp.Close
        excelTabela.Close
520     Set arqProp = Nothing
        Set excelTabela = Nothing
530     wordObj.Quit
        excelObj.Quit
540     Set wordObj = Nothing
        Set excelObj = Nothing
        Set marcacaoWord = Nothing

Exit Sub
Erro:
MsgBox "Ocorreu um erro ao gerar proposta. Comunicar ao analista responsável."
End Sub

Private Sub grdDetProp_Click()
   txtQtd = grdDetProp.TextMatrix(grdDetProp.Row, 0)
   txtEquipOper = grdDetProp.TextMatrix(grdDetProp.Row, 1)
   cmbUnidade = grdDetProp.TextMatrix(grdDetProp.Row, 2)
   txtPrecoUnit = Format$(grdDetProp.TextMatrix(grdDetProp.Row, 3), "##,#0.00")
   txtQtdUnid = grdDetProp.TextMatrix(grdDetProp.Row, 4)
   txtDiaria = Format$(grdDetProp.TextMatrix(grdDetProp.Row, 5), "##,#0.00")
   If grdDetProp.TextMatrix(grdDetProp.Row, 1) = "Habitat" Then
      fraMedidaHabitat.Visible = True
      Call Rotina_AbrirBanco
      rs.Open "Select comprimento,largura,altura from propostadetalhe where numProposta=('" & cmbProposta & "') and revisaoProposta=('" & cmbRevisao & "') and equipamento = ('" & grdDetProp.TextMatrix(grdDetProp.Row, 1) & "')", db, 3, 3
      txtCompHbt = rs!comprimento
      txtAlturaHBT = rs!altura
      txtLarguraHBT = rs!largura
      FechaDB
   End If
End Sub

Private Sub lstEquipamento_Click()
txtEquipOper = lstEquipamento.List(lstEquipamento.ListIndex)
txtQtd.SetFocus

End Sub

Private Sub txtAlturaHBT_LostFocus()

If txtCompHbt = Empty Then
   MsgBox ("Comprimento do Habitat não informado."), vbInformation
   Exit Sub
End If

If txtLarguraHBT = Empty Then
   MsgBox ("Largura do Habitat não informado."), vbInformation
   Exit Sub
End If

If txtAlturaHBT = Empty Then
   MsgBox ("Altura do Habitat não informado."), vbInformation
   Exit Sub
End If

If txtCompHbt <> Empty And txtAlturaHBT <> Empty And txtLarguraHBT <> Empty Then
   txtQtdUnid = (txtAlturaHBT * txtLarguraHBT) * 2 + (txtCompHbt * txtAlturaHBT) * 2 + (txtCompHbt * txtLarguraHBT) * 2 + 8
End If

'MedidaEquipamento = ((txtCompHbt * txtLarguraHBT) * 2) + ((txtCompHbt * txtAlturaHBT) * 2) + ((txtAlturaHBT * txtLarguraHBT) * 2) + 8
'lblTotalMetros = MedidaEquipamento
End Sub
Private Sub txtQtdUnid_LostFocus()
   If Not txtPrecoUnit = Empty And txtQtdUnid <> Empty And txtQtd <> Empty Then
      txtDiaria = Format$((txtPrecoUnit * txtQtdUnid) * txtQtd, "##,#0.00")
   Else
      MsgBox ("Valores informados com erro")
   End If
End Sub

Public Sub limpaCamposDetalheProposta()
   txtEquipOper = Empty
   txtQtdUnid = Empty
   txtPrecoUnit = Empty
   txtDiaria = Empty
   txtQtd = Empty
   cmbUnidade = Empty
End Sub

Public Sub GerarExcelWord()
        Dim CaminhoNew As String
                
        Call Rotina_AbrirBanco
        usu.Open "select usuEnderecoOneDrive from usuario where  chnome = ('" & glbUsuario & "')", db, 3, 3
        
        If usu!usuEnderecoOneDrive = Null Then
         MsgBox ("Não autorizada a impressão de proposta")
         FechaDB
         Exit Sub
        Else
         CaminhoNew = usu!usuEnderecoOneDrive & "Sistema\PROPOSTA MODELO\"
        End If
        
        Dim oApp As Excel.Application
        Dim oWB As Excel.Workbook
        Dim i As Integer
        Dim Ex As Object
        Set Ex = CreateObject("Excel.Application")

        i = 2
         On Error GoTo Erro
            'Create an Excel instance.
50          Set oApp = New Excel.Application

            'Open the desired workbook

60          If Dir(CaminhoNew & "ExcelEquipDetalhe.xlsx", vbArchive) = "" Then
70             MsgBox "Não foi possível gerar o documento porque" & vbCrLf & _
               "O arquivo padrão não foi localizado!", vbCritical
80             Exit Sub
90          End If
            
100         Set oWB = oApp.Workbooks.Open(FileName:=CaminhoNew & "ExcelEquipDetalhe.xlsx")
            
            'Do any modifications to the workbook.
            rs.Open "SELECT * FROM propostadetalhe where numProposta = ('" & cmbProposta & "') and revisaoProposta = ('" & cmbRevisao & "')", db, 3, 3
            Do Until rs.EOF
               Prod.Open "Select descricao from equipamentodescricao where equipamento = ('" & rs!equipamento & "')", db, 3, 3
               oApp.Cells(i, 1) = i - 1
               oApp.Cells(i, 2) = Prod!Descricao
               oApp.Cells(i, 3) = rs!quantidade
               oApp.Cells(i, 4) = rs!unidade
               oApp.Cells(i, 5) = rs!areaTotal
               oApp.Cells(i, 6) = rs!precoUnit
               oApp.Cells(i, 7) = rs!diaria
               Prod.Close
               rs.MoveNext
               i = i + 1
            Loop
110
          FechaDB

490       oWB.SaveAs FileName:=CaminhoNew & "ExcelDetEquip.xlsx"

510       oWB.Close SaveChanges:=False
520       Set oWB = Nothing
530       oApp.Quit
540       Set oApp = Nothing

Exit Sub
Erro:
MsgBox "Ocorreu um erro ao gerar o excel. Comunicar ao analista responsável." & Err.Description, vbInformation
End Sub


Public Sub verificaAlteracao()
   If flagNew = True Then
      flagAlteracao = False
   ElseIf responsavel <> cmbResponsavel Or Email <> cmbEmail Or QtdDias <> txtQtdDias Or UnidTemp <> cmbUnidTemp Or Status <> cmbStatus Or DataRevisao <> dtDataRevisao Or DataPedidoCotacao <> dtDataPedidoCotacao Then
      flagAlteracao = True
   Else
      flagAlteracao = False
   End If
End Sub
