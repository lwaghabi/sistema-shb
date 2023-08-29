VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmProcessaPedido 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Processa Pedido"
   ClientHeight    =   8700
   ClientLeft      =   8670
   ClientTop       =   5535
   ClientWidth     =   18300
   LinkTopic       =   "Form4"
   ScaleHeight     =   8700
   ScaleWidth      =   18300
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CFOPAux 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   83
      Top             =   8400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdConf 
      BackColor       =   &H000000FF&
      Caption         =   "Confirma Operação"
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
      Left            =   16680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Frame Frame9 
      Caption         =   "Negociação"
      Height          =   2895
      Left            =   120
      TabIndex        =   68
      Top             =   2520
      Width           =   13935
      Begin MSFlexGridLib.MSFlexGrid gridNegocio 
         Height          =   2055
         Left            =   120
         TabIndex        =   82
         Top             =   240
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColor       =   16777152
         BackColorFixed  =   16776960
         BackColorBkg    =   16777152
         FormatString    =   $"frmProcessaPedido.frx":0000
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
         Height          =   425
         Left            =   11520
         TabIndex        =   17
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
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
         Height          =   360
         Left            =   10560
         TabIndex        =   69
         Top             =   2280
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdCancela 
      BackColor       =   &H0000FFFF&
      Caption         =   "Cancela Operação"
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
      Left            =   14760
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Distribuição do Faturamento"
      Height          =   2895
      Left            =   720
      TabIndex        =   64
      Top             =   5400
      Width           =   12495
      Begin MSFlexGridLib.MSFlexGrid GridFatura 
         Height          =   2175
         Left            =   0
         TabIndex        =   81
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
         FormatString    =   "Tipo Operação       |Nota Fiscal          |Fatura                |Data Cobrança |Valor                                ||"
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
         Height          =   405
         Left            =   9960
         TabIndex        =   18
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Totais"
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
         Left            =   8880
         TabIndex        =   70
         Top             =   2400
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comissão"
      Height          =   2295
      Left            =   14040
      TabIndex        =   62
      Top             =   5280
      Width           =   4335
      Begin VB.TextBox txtComissaoPromot 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   34
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtPromotora 
         Height          =   285
         Left            =   120
         TabIndex        =   33
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtComissaoRep 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   32
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtRepresentante 
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Valor da Comissão"
         Height          =   195
         Left            =   240
         TabIndex        =   67
         Top             =   1800
         Width           =   1305
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Promotor"
         Height          =   195
         Left            =   2760
         TabIndex        =   66
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Label27 
         Caption         =   "Valor da Comissão"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Representante"
         Height          =   195
         Left            =   2280
         TabIndex        =   63
         Top             =   360
         Width           =   1050
      End
   End
   Begin VB.Frame Frame8 
      Height          =   2415
      Left            =   14040
      TabIndex        =   55
      Top             =   0
      Width           =   4335
      Begin VB.TextBox txtOrdemDeCarga 
         Height          =   285
         Left            =   2280
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtPlaca 
         Height          =   285
         Left            =   2520
         TabIndex        =   23
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtUFPlaca 
         Height          =   285
         Left            =   3000
         TabIndex        =   25
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtMunicipioPlaca 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox txtIdentidade 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtTransportadora 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtMotorista 
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Ordem de Carga"
         Height          =   195
         Left            =   2280
         TabIndex        =   71
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "UF"
         Height          =   195
         Left            =   2760
         TabIndex        =   61
         Top             =   2040
         Width           =   210
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Município"
         Height          =   195
         Left            =   1800
         TabIndex        =   60
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Placa"
         Height          =   195
         Left            =   3000
         TabIndex        =   59
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Transportadora"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   120
         Width           =   1770
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Motorista"
         Height          =   195
         Left            =   2760
         TabIndex        =   57
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ/CPF"
         Height          =   195
         Left            =   840
         TabIndex        =   56
         Top             =   1320
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalhes da Negociação"
      Height          =   2415
      Left            =   4320
      TabIndex        =   47
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtPercDescComis 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2400
         TabIndex        =   76
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtBancoFatN 
         Height          =   285
         Left            =   2040
         TabIndex        =   74
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtCondProcess 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtAliquotaICMS 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3120
         TabIndex        =   11
         Top             =   480
         Width           =   495
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frete - Condição e Data"
         Height          =   615
         Left            =   120
         TabIndex        =   53
         Top             =   1680
         Width           =   3495
         Begin MSMask.MaskEdBox txtDataFrete 
            Height          =   255
            Left            =   2280
            TabIndex        =   16
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtCondFrete 
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Faturamento"
         Height          =   855
         Left            =   120
         TabIndex        =   49
         Top             =   840
         Width           =   3495
         Begin VB.TextBox txtBancoFat 
            Height          =   285
            Left            =   120
            TabIndex        =   73
            Top             =   480
            Width           =   975
         End
         Begin MSMask.MaskEdBox txtDataPrimParc 
            Height          =   255
            Left            =   2280
            TabIndex        =   14
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtFaturamento 
            Height          =   285
            Left            =   1200
            TabIndex        =   12
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtIntervalo 
            Height          =   285
            Left            =   1680
            TabIndex        =   13
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Prim. parcela"
            Height          =   195
            Left            =   2280
            TabIndex        =   52
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Fatur."
            Height          =   195
            Left            =   1200
            TabIndex        =   51
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Interv."
            Height          =   195
            Left            =   1680
            TabIndex        =   50
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "%Desc"
         Height          =   195
         Left            =   2400
         TabIndex        =   75
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Alq.ICMS"
         Height          =   195
         Left            =   3000
         TabIndex        =   54
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cond. de Processamento"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1785
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Negociação"
      Height          =   2415
      Left            =   120
      TabIndex        =   39
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtNatOperacao 
         Height          =   285
         Left            =   1680
         TabIndex        =   80
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtCEFOP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   77
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtEmissorNF 
         Height          =   285
         Left            =   2640
         TabIndex        =   9
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtNotaFiscal 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtNumPedido 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtCompPedido 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3855
      End
      Begin MSMask.MaskEdBox txtDataPedido 
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDataEmissao 
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label35 
         Caption         =   "Natureza da Operação"
         Height          =   255
         Left            =   1680
         TabIndex        =   79
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label34 
         Caption         =   "CEFOP"
         Height          =   255
         Left            =   1080
         TabIndex        =   78
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nota Fiscal"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Data Emissão"
         Height          =   195
         Left            =   1320
         TabIndex        =   45
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Emissor"
         Height          =   195
         Left            =   2640
         TabIndex        =   44
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pedido"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Comp"
         Height          =   195
         Left            =   2040
         TabIndex        =   42
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data Pedido"
         Height          =   195
         Left            =   2760
         TabIndex        =   41
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Resumo Financeiro"
      Height          =   2895
      Left            =   14040
      TabIndex        =   0
      Top             =   2400
      Width           =   4335
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
         Height          =   420
         Left            =   2400
         TabIndex        =   30
         Top             =   2280
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
         Height          =   420
         Left            =   2400
         TabIndex        =   29
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtResTotalICMS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   28
         Top             =   240
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
         Height          =   420
         Left            =   2400
         TabIndex        =   26
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtResTotalFrete 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   27
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total da Nota"
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
         TabIndex        =   38
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Frete"
         Height          =   195
         Left            =   1200
         TabIndex        =   37
         Top             =   600
         Width           =   360
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
         TabIndex        =   36
         Top             =   1920
         Width           =   1005
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
         TabIndex        =   35
         Top             =   1440
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frmProcessaPedido"
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
Dim DataUtil As Date

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
Dim Valor_Total_Cheio As Currency
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

Dim ICMS_ST As Currency
Dim Base As Byte

Dim Data_Proc As Date
Dim DataVencimento As Date

Dim ChavePessoa As Integer

Private Sub cmdCancela_Click()
GlbStatus = "PENDENTE"
frmPedido.txtNumPedido = txtNumPedido
frmPedido.txtComplementoPedido = txtCompPedido

MsgBox ("Processamento de medição cancelado."), vbInformation
Unload Me

End Sub

Private Sub cmdConf_Click()

Resp = MsgBox("Procedimento considerado correto. Confirma???", vbYesNo)
If Resp = vbNo Then
   MsgBox ("Processamento abortado"), vbInformation
   
   Call FechaDB
   
   GlbStatus = "PENDENTE"
   frmPedido.txtNumPedido = txtNumPedido
   frmPedido.txtComplementoPedido = txtCompPedido
   Unload Me
   Exit Sub
End If
   
Dim Data_Pedido As Date

'If txtOrdemDeCarga = Empty Then
'   MsgBox ("Ordem de Carga não Informado")
'   txtOrdemDeCarga.SetFocus
'   Exit Sub
'End If

Data_Proc = txtDataEmissao

ano = Year(Data_Proc)
Mes = Month(Data_Proc)
Dia = Day(Data_Proc)

Data_Pedido = txtDataPedido

Mes_Pedido = Month(Data_Pedido)

'MsgBox ("Mes Pedido = "), , Mes_Pedido

Call Rotina_AbrirBanco

Bco.Open "Select * from Banco where bcoCodBcoLart = ('" & txtBancoFatN & "')", db, 3, 3
       If Bco.EOF Then
          MsgBox ("Banco não encontrado em gerar contas a receberem Processamento de Medição"), vbCritical
          Call FechaDB
          Exit Sub
       End If

pes.Open "Select * from Pessoa where chPessoa = ('" & txtCliente & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Erro no acesso a Pessoa em Confirma Processamento de Medição."), vbCritical
   Call FechaDB
   Exit Sub
End If

db.BeginTrans
   
'Gravacao de Contas a Receber

ctr.Open "Select * from Contas_A_Receber where chFabricante = ('" & 0 & "') and chPessoa = ('" & txtCliente & "') and chNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3

If Not (GridFatura.TextMatrix(1, 2) = "") Then
   For A = 1 To IndFatSalvo
       ctr.AddNew
       ctr!chFabricante = 0
       ctr!chPessoa = txtCliente
       ctr!chNotaFiscal = txtNotaFiscal
       ctr!chFatura = GridFatura.TextMatrix(A, 2)
       ctr!ctrDataEmissao = txtDataEmissao
       ctr!ctrDataVencito = GridFatura.TextMatrix(A, 3)
       DataVencimento = GridFatura.TextMatrix(A, 3)
       'Calcula data banco
    
       DataUtil = GridFatura.TextMatrix(A, 3)
       
       DataInformada = DataUtil
       NDias = 0
       'DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)
       'DataUtil = DataRetorno.DiaUtil
       ctr!ctrDataBanco = DataUtil
    
       'Fim calcula data banco
       
       ctr!ctrDataVencitoOriginal = GridFatura.TextMatrix(A, 3)
       ctr!ctrDescricaoOperacao = GridFatura.TextMatrix(A, 0)
       ctr!ctrValorLart = GridFatura.TextMatrix(A, 4)
       ctr!ctrValorDaBoleta = GridFatura.TextMatrix(A, 4)
       ctr!ctrValorMerco = GridFatura.TextMatrix(A, 5)
       
       
       ctr!ctrPercentCorrecao = pes!pesrapell
       ctr!ctrPercentlogistica = pes!peslogistica
       'ctr!ctrvalorcorrecao = Format$((pes!pesrapell * GridFatura.TextMatrix(A, 6)) / 100) * -1, "#0.00")
       'ctr!ctrvalorlogistica = Format$(((pes!peslogistica") * GridFatura.TextMatrix(A, 6)) / 100) * -1, "#0.00")
              
       'ctr!ctrvalordaboleta = GridFatura.TextMatrix(A, 6) + (ctr!ctrvalorcorrecao") + ctr!ctrvalorlogistica"))
       
       ctr!chAno = ano
       ctr!chMes = Mes
       ctr!chDia = Dia
       ctr!chNumPedido = txtNumPedido
       ctr!chNumPedidoComp = txtCompPedido
       
       
       ctr!chCodBcoLart = Bco!bcosiglabco
       
      Resp = Mid$(txtNotaFiscal, 1, 2)
      If Resp = "RL" Then
         ctr!ctrCentroDeCusto = 1
         ctr!ctrGrupoCentroDeCusto = Format$(1, "00")
         ctr!ctrSubGrupoCentroDeCusto = Format$(1, "00")
      Else
         If Resp = "NF" Then
            ctr!ctrCentroDeCusto = 1
            ctr!ctrGrupoCentroDeCusto = Format$(1, "00")
            ctr!ctrSubGrupoCentroDeCusto = Format$(2, "00")
         Else
            If Resp = "ND" Then
               ctr!ctrCentroDeCusto = 1
               ctr!ctrGrupoCentroDeCusto = Format$(2, "00")
               ctr!ctrSubGrupoCentroDeCusto = Format$(3, "00")
            Else
               If Resp = "AP" Then
                  ctr!ctrCentroDeCusto = 1
                  ctr!ctrGrupoCentroDeCusto = Format$(2, "00")
                  ctr!ctrSubGrupoCentroDeCusto = Format$(1, "00")
               Else
                  If Resp = "VD" Then
                     ctr!ctrCentroDeCusto = 1
                     ctr!ctrGrupoCentroDeCusto = Format$(2, "00")
                     ctr!ctrSubGrupoCentroDeCusto = Format$(2, "00")
                  Else
                     MsgBox ("ERRO:Informar ao analista responsável"), vbCritical
                     ctr!ctrCentroDeCusto = 1
                     ctr!ctrGrupoCentroDeCusto = Format$(1, "00")
                     ctr!ctrSubGrupoCentroDeCusto = Format$(1, "00")
                  End If
               End If
            End If
         End If
      End If
       
       
       If (GridFatura.TextMatrix(A, 6) + ctr!ctrvalorcorrecao) > 0 Then
          ctr.Update
       End If
       'Se houver valor para a Lart, gerar recebimento de repasse
       'em contas a receber
       
             
       If GridFatura.TextMatrix(A, 5) > 0 Then
          
          ctp.Open "Select * from Contas_A_Pagar where chFabricante = ('" & 0 & "') and chPessoa = ('" & txtCliente & "') and chNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3
          If ctp.EOF Then
             ctp.AddNew
          End If
             
          ctp!chFabricante = 0
          ctp!chPessoa = txtCliente
          ctp!chNotaFiscal = txtNotaFiscal
          ctp!chFatura = GridFatura.TextMatrix(A, 2)
          ctp!ctpDataEmissao = txtDataEmissao
          ctp!ctpDataLanc = Date
          ctp!chDataVencito = GridFatura.TextMatrix(A, 3)
                                                 
          'Calcula data banco
                      
          DataUtil = GridFatura.TextMatrix(A, 3)
               
          DataInformada = DataUtil
          NDias = 0
          'DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)
          'DataUtil = DataRetorno.DiaUtil
          ctp!ctpDataBanco = DataUtil
            
          'Fim calcula data banco
       
          ctp!ctpDataVencOriginal = GridFatura.TextMatrix(A, 3)
          ctp!ctpDescricaoOperacao = "Repasse a Debitar"
          ctp!ctpValorLart = 0
          ctp!ctpValorMerco = 0
             
             
          ctr!ctrPercentCorrecao = pes!pesrapell
          ctr!ctrPercentlogistica = pes!peslogistica
          ctr!ctrvalorcorrecao = Format$(((pes!pesrapell * GridFatura.TextMatrix(A, 6)) / 100) * -1, "#0.00")
          ctr!ctrValorlogistica = Format$(((pes!peslogistica * GridFatura.TextMatrix(A, 6)) / 100) * -1, "#0.00")
              
          ctr!ctrValorDaBoleta = GridFatura.TextMatrix(A, 6) + ctr!ctrvalorcorrecao + ctr!ctrValorlogistica
                     

          ctp!chAno = ano
          ctp!chMes = Mes
          ctp!chDia = Dia
          'TabBanco.Seek "=", 0, txtBancoFatN
          'If TabBanco.NoMatch Then
          '   MsgBox ("Banco inválido"), , txtBancoFatN
          '   fim = 1 / 0
          '   End If
          ctp!chCodBcoLart = Bco!bcosiglabco
             
          ctp.Update
             
          ctr.AddNew

          ctr!chFabricante = 1
          ctr!chPessoa = txtCliente
          ctr!chNotaFiscal = txtNotaFiscal
          ctr!chFatura = GridFatura.TextMatrix(A, 2)
          ctr!ctrDataEmissao = txtDataEmissao
          ctr!ctrDataVencito = GridFatura.TextMatrix(A, 3)
                                                           
          'Calcula data banco
                      
          DataUtil = GridFatura.TextMatrix(A, 3)
               
          DataInformada = DataUtil
          NDias = 0
          'DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)
          'DataUtil = DataRetorno.DiaUtil
          ctr!ctrDataBanco = DataUtil
            
          'Fim calcula data banco
       
          ctr!ctrDataVencitoOriginal = GridFatura.TextMatrix(A, 3)
          ctr!ctrDescricaoOperacao = "Repasse a Creditar"
          ctr!ctrValorLart = 0
          ctr!ValorDaBoleta = 0
          ctr!ctrValorMerco = GridFatura.TextMatrix(A, 5)
           
          ctr!ctrPercentCorrecao = pes!pesrapell
          ctr!ctrPercentlogistica = pes!peslogistica
          ctr!ctrvalorcorrecao = Format$(((pes!pesrapell * GridFatura.TextMatrix(A, 6)) / 100) * -1, "#0.00")
          ctr!ctrValorlogistica = Format$(((pes!peslogistica * GridFatura.TextMatrix(A, 6)) / 100) * -1, "#0.00")
            
          ctr!ctrValorDaBoleta = GridFatura.TextMatrix(A, 6) + ctr!ctrvalorcorrecao + ctr!ctrValorlogistica
               
          ctr!chAno = ano
          ctr!chMes = Mes
          ctr!chDia = Dia
          ctr!chNumPedido = txtNumPedido
          ctr!chNumPedidoComp = txtCompPedido

          Resp = Mid$(txtNotaFiscal, 1, 2)
          If Resp = "RL" Then
             ctr!ctrCentroDeCusto = 1
             ctr!ctrGrupoCentroDeCusto = Format$(1, "00")
             ctr!ctrSubGrupoCentroDeCusto = Format$(1, "00")
          Else
             If Resp = "NF" Then
                ctr!ctrCentroDeCusto = 1
                ctr!ctrGrupoCentroDeCusto = Format$(1, "00")
                ctr!ctrSubGrupoCentroDeCusto = Format$(2, "00")
             Else
                If Resp = "AP" Then
                   ctr!ctrCentroDeCusto = 1
                   ctr!ctrGrupoCentroDeCusto = Format$(2, "00")
                   ctr!ctrSubGrupoCentroDeCusto = Format$(1, "00")
                Else
                   If Resp = "VD" Then
                      ctr!ctrCentroDeCusto = 1
                      ctr!ctrGrupoCentroDeCusto = Format$(1, "00")
                      ctr!ctrSubGrupoCentroDeCusto = Format$(2, "00")
                   Else
                      ctr!ctrCentroDeCusto = 1
                      ctr!ctrGrupoCentroDeCusto = Format$(1, "00")
                      ctr!ctrSubGrupoCentroDeCusto = Format$(1, "00")
                   End If
                End If
             End If
          End If
          ctr!chCodBcoLart = Bco!bcosiglabco
          If GridFatura.TextMatrix(A, 6) > 0 Then
             ctr.Update
          End If
       End If
    

Next


'Gravacao de Contas a Pagar

'Data_Comissao = 25 & "/" & Mes_Comissao & "/" & Ano_Comissao

'TabCtaPagar.Seek "=", 0, txtRepresentante, "Representante", "Comissão", Data_Comissao

'MsgBox ("Acumula Rep 0 = "), , Acumula_Comis_Rep(0)

'If TabCtaPagar.NoMatch Then
'   If (Acumula_Comis_Rep(0) > 0) Then
'      TabCtaPagar.AddNew
'      Fazer = 1
'   Else
'      Fazer = 0
'   End If
'Else
'   If (Acumula_Comis_Rep(0) > 0) Then
'      TabCtaPagar.Edit
'      Fazer = 1
'   Else
'      Fazer = 0
'   End If
'End If
'Verificar esta rotina
'If Fazer = 1 Then
'   Fazer = 0
'   TabCtaPagar("chFabricante") = 0
'   TabCtaPagar("chpessoa") = txtRepresentante
'   TabCtaPagar("chFatura") = "Comissão"
'   TabCtaPagar("chDatavencito") = Data_Comissao
                                                         
   'Calcula data banco
             
'   DataUtil = Data_Comissao
     
'   DataInformada = DataUtil
'   NDias = 0
'   'DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)
'   'DataUtil = DataRetorno.DiaUtil
'   TabCtaPagar("ctpDataBanco") = DataUtil
'
'   'Fim calcula data banco
'
'   TabCtaPagar("ctpDatavencoriginal") = Data_Comissao
'   TabCtaPagar("chNotaFiscal") = "Representante"
'   TabCtaPagar("ctpvalorLart") = TabCtaPagar("ctpvalorLart") + Acumula_Comis_Rep(0)
'   TabCtaPagar("ctpvalorMerco") = 0
'   TabCtaPagar("ctpvalordaboleta") = TabCtaPagar("ctpvalordaboleta") + Acumula_Comis_Rep(0)
'
'   Tipo_Comissao = 1
'   Call Rotina_MoverDados_CtaPagar
'
'   TabCtaPagar.Update
'
'End If

'MsgBox ("Acumula Rep 1 = "), , Acumula_Comis_Rep(1)

'TabCtaPagar.Seek "=", 1, txtRepresentante, "Representante", "Comissão", Data_Comissao
'If TabCtaPagar.NoMatch Then
'   If (Acumula_Comis_Rep(1) > 0) Then
'      TabCtaPagar.AddNew
'      Fazer = 1
'   Else
'      Fazer = 0
'   End If
'Else
'   If (Acumula_Comis_Rep(1) > 0) Then
'      TabCtaPagar.Edit
'      Fazer = 1
'   Else
'      Fazer = 0
'   End If
'End If''

'If Fazer = 1 Then
'   Fazer = 0
'   TabCtaPagar("chFabricante") = 1
'   TabCtaPagar("chpessoa") = txtRepresentante
'   TabCtaPagar("chFatura") = "Comissão"
'   TabCtaPagar("chdatavencito") = Data_Comissao
'
'   'Calcula data banco
             
'   DataUtil = Data_Comissao
     
'   DataInformada = DataUtil
'   NDias = 0
'   'DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)
'   'DataUtil = DataRetorno.DiaUtil
'   TabCtaPagar("ctpDataBanco") = DataUtil
'
'   'Fim calcula data banco
'
'   TabCtaPagar("ctpDatavencoriginal") = Data_Comissao
'   TabCtaPagar("chNotaFiscal") = "Representante"
'   TabCtaPagar("ctpvalorLart") = 0
'   TabCtaPagar("ctpvalorMerco") = TabCtaPagar("ctpvalorMerco") + Acumula_Comis_Rep(1)
'   TabCtaPagar("ctpvalordaboleta") = TabCtaPagar("ctpvalordaboleta") + Acumula_Comis_Rep(1)
'
'   Tipo_Comissao = 1
'   Call Rotina_MoverDados_CtaPagar
'
'   TabCtaPagar.Update
''
End If
    

'Contas a Pagar de Promotoras

'Data_Comissao = 5 & "/" & Mes_Comissao & "/" & Ano_Comissao

'TabCtaPagar.Seek "=", 0, txtPromotora, "Promotora", "Comissão", Data_Comissao
'If TabCtaPagar.NoMatch Then
'   If (Acumula_Comis_Promot(0) > 0) Then
'      TabCtaPagar.AddNew
'      Fazer = 1
''   Else
''      Fazer = 0
'   End If
'Else
'   If (Acumula_Comis_Promot(0) > 0) Then
'      TabCtaPagar.Edit
'      Fazer = 1
'   Else
'      Fazer = 0
''   End If
'End If'

'If Fazer = 1 Then
'   Fazer = 0
'   If Acumula_Comis_Promot(0) > 0 Then
'      TabCtaPagar("chFabricante") = 0
'      TabCtaPagar("chpessoa") = txtPromotora
'      TabCtaPagar("chdatavencito") = Data_Comissao
'
'      'Calcula data banco
                
'      DataUtil = Data_Comissao
'
'      DataInformada = DataUtil
'      NDias = 0
'      'DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)
'      'DataUtil = DataRetorno.DiaUtil
'      TabCtaPagar("ctpDataBanco") = DataUtil
'
'      'Fim calcula data banco
'
'      TabCtaPagar("ctpDatavencoriginal") = Data_Comissao
'      TabCtaPagar("chFatura") = "Comissão"
'      TabCtaPagar("chNotaFiscal") = "Promotora"
'      TabCtaPagar("ctpvalorLart") = TabCtaPagar("ctpvalorLart") + Acumula_Comis_Promot(0)
'      TabCtaPagar("ctpvalorMerco") = 0
'      TabCtaPagar("ctpvalordaboleta") = TabCtaPagar("ctpvalordaboleta") + Acumula_Comis_Promot(0)
'
'      Tipo_Comissao = 2
'      Call Rotina_MoverDados_CtaPagar
'
'      TabCtaPagar.Update
'   End If
'End If

'TabCtaPagar.Seek "=", 1, txtPromotora, "Promotora", "Comissão", Data_Comissao'

'If TabCtaPagar.NoMatch Then
'   If (Acumula_Comis_Promot(1) > 0) Then
'      TabCtaPagar.AddNew
'      Fazer = 1
'   Else
'''      Fazer = 0
'   End If
'Else
'   If (Acumula_Comis_Promot(1) > 0) Then
'      TabCtaPagar.Edit
'      Fazer = 1
'   Else
'      Fazer = 0
'   End If
'End If''

'If Fazer = 1 Then
'   Fazer = 0
'
'   If Acumula_Comis_Promot(1) > 0 Then
 '     TabCtaPagar("chFabricante") = 1
'      TabCtaPagar("chpessoa") = txtPromotora
'      TabCtaPagar("chFatura") = "Comissão"
'      TabCtaPagar("chdatavencito") = Data_Comissao
'
'      'Calcula data banco
'
'      DataUtil = Data_Comissao
'
'      DataInformada = DataUtil
'      NDias = 0
'      'DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)
'      'DataUtil = DataRetorno.DiaUtil
'      TabCtaPagar("ctpDataBanco") = DataUtil
'
'      'Fim calcula data banco
'
'      TabCtaPagar("ctpDatavencoriginal") = Data_Comissao
'      TabCtaPagar("chNotaFiscal") = "Promotora"
'      TabCtaPagar("ctpvalorLart") = 0
'      TabCtaPagar("ctpvalorMerco") = Acumula_Comis_Promot(1)
'      TabCtaPagar("ctpvalordaboleta") = TabCtaPagar("ctpvalordaboleta") + Acumula_Comis_Promot(1)
'
'      Tipo_Comissao = 2
'      Call Rotina_MoverDados_CtaPagar
'
'      TabCtaPagar.Update
'   End If
'End If


neg.Open "Select * from Negociacao where chNumPedido = ('" & txtNumPedido & "') and chNumPedidoComp = ('" & txtCompPedido & "')", db, 3, 3
If neg.EOF Then
   MsgBox ("Não encontrei o pedido em Negociacao"), vbCritical
   Call FechaDB
   Exit Sub
End If

neg!negNotaFiscal = txtNotaFiscal
neg!negCEFOP = CFOPAux.ListIndex
neg!chCodBcoLart = txtBancoFatN
neg!negEmissorNF = 0
neg!negdatanegociação = txtDataEmissao
neg!negvalornegociacao = txtTotalFatura
neg!chOrdemDeCarga = txtOrdemDeCarga
neg!negTransporte = txtTransportadora
neg!negPlaca = txtPlaca
neg!negICMS = Format$(txtResTotalICMS, "0.00")
'neg!negAliquota = Format$(ICM!icmAliquota, "0.00")
'neg!negFretePedido = Format$(txtResTotalFrete, "0.00")
neg!negValorDoProduto = Format$(txtResTotalProduto, "0.00")
'neg!negIPI = Format$(txtResTotalIPI, "0.00")
'neg!negDescontoTotalPedido = Format$(txtResTotalDesconto, "0.00")
If txtComissaoRep = "" Then
   txtComissaoRep = 0
End If
neg!negComisRepPedido = Format$(txtComissaoRep, "0.00")
If txtComissaoPromot = "" Then
   txtComissaoPromot = 0
End If
neg!negComisPromotPedido = Format$(txtComissaoPromot, "0.00")
'neg!negDataVencimento = DataVencimento

neg!negStatus = 1


neg!negEmissorNF = 0

neg!negUltimaAtualizacao = Data_Proc
'neg!chordemdecarga = txtOrdemDeCarga

neg.Update

db.CommitTrans
'MsgBox ("Rotina de chamada de atualizacao de estoque")
'   fim = 0
'   For A = 1 To 99
'      If Mneu_Produto(A) = Empty Then
'
'         A = 99
'      Else
'         TabDetalheNegociacao.Seek "=", txtNumPedido, txtCompPedido, Mneu_Produto(A)
'         If TabDetNeg.NoMatch Then
'            MsgBox ("Erro no acesso ao detneg financeiro")
'            A = 1 / 0
'         Else
'            Funcao = 6
'            Produto = TabDetalheNegociacao("chproduto")
'            Sai = TabDetalheNegociacao("pedquantidadepedida")
'            TracoOut = 0
'            Entra = 0
'            TracoIn = 0'''

'            Call Rotina_Atualiza_Estoque(Funcao, Ano, Mes, Produto, Entra, Sai, TracoIn, TracoOut, Mes_Pedido)
'
'            Mneu_Produto(A) = Empty
'
'            If txtCEFOP = 5949 Or txtCEFOP = 6949 Or txtCondProcess = "CONSIGNAÇÃO" Then
'               TabDetalheNegociacao.Edit
'               TabDetalheNegociacao("pedcomissaorep") = 0
'               TabDetalheNegociacao("pedcomissaopromot") = 0
'               TabDetalheNegociacao.Update
'            End If
'         End If
'      End If
'
'    Next
cmdConf.Enabled = False
cmdCancela.Enabled = False

'If txtOrdemDeCarga = "Cliente" Or TabNegociacao("negcobrancafrete") = 7 Or txtCondFrete = "VALOR FIXO FATURA" Then
'   txtResTotalFrete = 0
'End If

'If txtOrdemDeCarga = "Cliente" Or txtCondFrete = "VALOR FIXO" Or txtCondFrete = "VALOR FIXO FATURA" Or TabNegociacao("negcobrancafrete") = 7 Then
'   txtResTotalFrete = 0
'End If


   
frmPedido.txtNumPedido = txtNumPedido
frmPedido.txtComplementoPedido = txtCompPedido
'frmPedido.cmbOrdemDeCarga = txtOrdemDeCarga
GlbStatus = "PROCESSADO"

Call FechaDB

Unload Me

End Sub

Private Sub Form_Load()

Unidade(0) = "M2"
Unidade(1) = "Un"
Unidade(2) = "Hr"

ProdutoConsig(1) = "CSGR"
ProdutoConsig(2) = "CSGP"

DataString = frmPedido.txtDataProc

txtDataEmissao = frmPedido.txtDataProc
txtEmissorNF = frmPedido.cmbEmissor

Data_Proc = DataString

ano = Year(Data_Proc)
Mes = Month(Data_Proc)
Dia = Day(Data_Proc)

Ano_Comissao = ano
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

Bco.Open "Select * from Banco where bcoempresa = ('" & 0 & "') and bcoCodBcoLart = ('" & frmPedido.cmbBanco.ListIndex & "')", db, 3, 3
If Bco.EOF Then
   MsgBox ("Erro acesso a Banco em Processa Pedido."), vbCritical
   Call FechaDB
   Exit Sub
End If

neg.Open "Select * from Negociacao where chNumPedido = ('" & frmPedido.txtNumPedido & "') and chNumPedidoComp = ('" & frmPedido.txtComplementoPedido & "')", db, 3, 3
If neg.EOF Then
   MsgBox ("Numero de pedido inválido em Processa Pedido."), vbCritical
   Call FechaDB
   Unload Me
End If

txtCliente = neg!chPessoa
txtDataPedido = neg!negDataPedido
txtNumPedido = neg!chNumPedido
txtCompPedido = neg!chNumPedidoComp
txtNotaFiscal = frmPedido.txtNotaFiscal
txtBancoFat = Bco!bcosiglabco
txtBancoFatN = frmPedido.cmbBanco.ListIndex
'txtOrdemDeCarga = frmPedido.cmbOrdemDeCarga
txtEmissorNF = frmPedido.cmbEmissor
txtCEFOP = Mid$(frmPedido.cmbCFOP, 1, 4)
glbCFOP = Mid$(txtCEFOP, 1, 4)
glbCFOP = glbCFOP
                                                       
NatuOper.Open "Select * from NaturezaOperacao", db, 3, 3
If NatuOper.EOF Then
   MsgBox ("Natureza Operacao Invalida. Cancelar Processamento"), vbCritical
   Exit Sub
End If
NatuOper.MoveFirst
Do While Not NatuOper.EOF
   If NatuOper!Status = 1 Then
      CFOPAux.AddItem NatuOper!cfop
      NatuOper.MoveNext
   Else
      NatuOper.MoveNext
   End If
Loop

NatuOper.Close: Set NatuOper = Nothing
CFOPAux.ListIndex = neg!negCEFOP
glbCFOP = CFOPAux
NatuOper.Open "Select * from NaturezaOperacao where CFOP = ('" & glbCFOP & "')", db, 3, 3
If NatuOper.EOF Then
   MsgBox ("Natureza Operacao Invalida. Cancelar Processamento"), vbCritical
   Exit Sub
End If

txtNatOperacao = NatuOper!natoperacaoabrev

'TabCobrancaFrete.Seek "=", TabNegociacao("negCobrancaFrete")

'If TabCobrancaFrete.NoMatch Then
'   MsgBox ("Erro na Leitura do Parametro Cobranca Frete")
'   A = 1 / 0
'Else
'   txtCondFrete = TabCobrancaFrete("parDescCobrancaFrete")
'   If TabNegociacao("negCobrancaFrete") = 1 Then
'      txtDataFrete = Data_Proc + TabNegociacao("negBoletaFrete")
'   Else
'      txtDataFrete = "__/__/____"
'   End If
'End If

CondProc.Open "Select * from CondProcessamento where chCondicaoProcessamento = ('" & neg!negCondProcess & "')", db, 3, 3
If CondProc.EOF Then
   MsgBox ("Erro na Leitura do Parametro Condição de Processamento"), vbCritical
   Call FechaDB
   Exit Sub
Else
   txtCondProcess = CondProc!cprDescCondProcess
End If
txtPercDescComis = Format$(neg!negdesccomissao, "#0.00") & "%"

pes.Open "Select * from Pessoa where chPessoa = ('" & neg!chPessoa & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Erro na Leitura de Tabpessoa em Processa Pedido;"), vbCritical
   Call FechaDB
   Exit Sub
Else
   Guarda_Cliente = pes!chPessoa
   ICM.Open "Select * from ICMS where chUF = ('" & pes!chUF & "')", db, 3, 3
   If ICM.EOF Then
      MsgBox ("Erro na leitura de ICMS em Processa Pedido"), vbCritical
      Call FechaDB
      Exit Sub
   Else
      txtAliquotaICMS = ICM!icmAliquota & "%"
   End If
    
End If


If pes!chcarteirarep = Empty Then
   MsgBox ("Cliente sem informacao de representante. Verificar!!!!"), vbCritical
   Call FechaDB
   Unload Me
Else
   CartRep.Open "Select * from Carteira_Rep where chPessoa = ('" & pes!chcarteirarep & "')", db, 3, 3
   If CartRep.EOF Then
       MsgBox ("Carteira de representante invalida. em Processa Pedido"), vbCritical
       Call FechaDB
       Unload Me
    End If
End If
'If Not (neg!chrepresentante = CartRep!chpessoa) Then
'   MsgBox ("Rep Negociacao diferente do Rep atual"), vbInformation
'   Resp = MsgBox("Para manter o representante anterior informe Sim", vbYesNo)
'   If Resp = vbYes Then
'      TabCarteira_Rep.MoveFirst
 '     Do While fim = 0
 '        If TabCarteira_Rep("chpessoa") = TabNegociacao("chrepresentante") Then
'            fim = 1
'         Else
'            TabCarteira_Rep.MoveNext
'         End If
'      Loop
'   Else
''      TabNegociacao.Edit
'      TabNegociacao("chrepresentante") = TabCarteira_Rep("chpessoa")
'      TabNegociacao.Update
'      TabNegociacao.Seek "=", txtNumPedido, txtCompPedido
 '  End If
'End If
'If Not (Tabpessoa("chcarteirapromot") = Empty) Then
 '   TabCarteira_Promot.Seek "=", Tabpessoa("chCarteirapromot")
 '   If TabCarteira_Promot.NoMatch Then
 '      MsgBox ("Carteira de Promotores invalida.")
 '      Unload Me
 '   End If
'End If

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

dneg.Open "Select * from DetalheNegociacao where chNumPedido = ('" & neg!chNumPedido & "') and chNumpedidoComp = ('" & neg!chNumPedidoComp & "')", db, 3, 3
If dneg.EOF Then
   MsgBox ("Detalhe de Negociação não encontrado não encontrado"), vbCritical
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
Acumula_Frete_Fabricante(0) = 0
Acumula_Frete_Fabricante(1) = 0
Acumula_Frete_Fabricante(2) = 0

'GridNegocio.ColAlignment(0) = 1

dneg.MoveFirst
Linha = 1
Do While Fim_Carga = 0
   If pes.State = 1 Then
      pes.Close: Set pes = Nothing
   End If
   pes.Open "Select * from Pessoa where chPessoa = ('" & dneg!chProduto & "')", db, 3, 3
   If pes.EOF Then
      If Prod.State = 1 Then
         Prod.Close: Set Prod = Nothing
      End If
      Prod.Open "Select * from Produto where chProduto = ('" & dneg!chProduto & "')", db, 3, 3
      If Prod.EOF Then
         MsgBox ("Produto não encontrado em Processa Pedido. Erro Fatal"), vbCritical
         Call FechaDB
         Exit Sub
      Else
         ChavePessoa = 0
      End If
   Else
      ChavePessoa = 1
   End If

   GridNegocio.Rows = Linha + 1
   If ChavePessoa = 0 Then
      GridNegocio.TextMatrix(Linha, 0) = Prod!prdNomeProd
      Mneu_Produto(Linha) = Prod!chProduto
   Else
      GridNegocio.TextMatrix(Linha, 0) = pes!chPessoa
      Mneu_Produto(Linha) = pes!chPessoa
   End If
   GridNegocio.TextMatrix(Linha, 1) = Unidade(dneg!pedunidade)
   GridNegocio.TextMatrix(Linha, 2) = Format$(dneg!pedPrecoUnidadePedida, "0.00")
   GridNegocio.TextMatrix(Linha, 3) = dneg!pedquantidadePedida
   GridNegocio.TextMatrix(Linha, 4) = Format$(dneg!pedValorDaDiaria, "0.00")
   Valor_Produto = Format$(GridNegocio.TextMatrix(Linha, 4), "#,##0.00")
   GridNegocio.TextMatrix(Linha, 5) = dneg!pedqtddias
         'Valor_Frete = Format$(GridNegocio.TextMatrix(Linha, 5), "##0.00")
       
   GridNegocio.TextMatrix(Linha, 6) = Format$((dneg!pedValorDaDiaria * dneg!pedqtddias), "#,##0.00")
   GridNegocio.TextMatrix(Linha, 7) = Format$(dneg!pedValorDesconto, "#,##0.00")
   GridNegocio.TextMatrix(Linha, 8) = Format$((dneg!pedValorDaDiaria * dneg!pedqtddias), "#,##0.00")
         
   Valor_IPI = Format$(GridNegocio.TextMatrix(Linha, 6), "0.00")
   'GridNegocio.TextMatrix(Linha, 7) = Format$((dneg!pedquantidadePedida * dneg!pedValorDesconto), "0.00")
   Valor_Desconto = Format$(GridNegocio.TextMatrix(Linha, 7), "#0.00")
   Valor_Total = Format$((Valor_Produto * dneg!pedqtddias), "#,##0.00")
   Valor_Total_Cheio = Format$(Valor_Total_Cheio + ((dneg!pedValorDaDiaria * dneg!pedqtddias) + dneg!pedValorDesconto), "#,##0.00")
  ' GridNegocio.TextMatrix(Linha, 6) = Format$(Valor_Total, "##,##0.00")
         
   IndNegSalvo = Linha
         
         'Acumula por tipo de produto, para calculo de comissao de consignacao.(Piso ou Revestimento)
         
         'AcumValorConsig(tabproduto("prdtipo")) = AcumValorConsig(tabproduto("prdtipo")) + Valor_Total
                  
   Acumula_Quantidade = Acumula_Quantidade + dneg!pedquantidadePedida
   Acumula_Valor_Total = Acumula_Valor_Total + Valor_Total
   Acumula_Valor_Produto = Acumula_Valor_Produto + Valor_Produto
         'Acumula_Valor_Frete = Acumula_Valor_Frete + Valor_Frete
   Acumula_Valor_IPI = Acumula_Valor_IPI + Valor_IPI
   Acumula_Valor_Desconto = Acumula_Valor_Desconto + Valor_Desconto
         
         'IndFabricante = tabproduto("prdFabricante")
   IndFabricante = 0
         
   Acumula_Total_Fabricante(IndFabricante) = Acumula_Total_Fabricante(IndFabricante) + Valor_Total
   Acumula_Produto_Fabricante(IndFabricante) = Acumula_Produto_Fabricante(IndFabricante) + Valor_Produto
   Acumula_Frete_Fabricante(IndFabricante) = Acumula_Frete_Fabricante(IndFabricante) + Valor_Frete
   Acumula_IPI_Fabricante(IndFabricante) = Acumula_IPI_Fabricante(IndFabricante) + Valor_IPI
   Acumula_Desconto_Fabricante(IndFabricante) = Acumula_Desconto_Fabricante(IndFabricante) + Valor_Desconto
   Acumula_Quantidade_Fabricante(IndFabricante) = Acumula_Quantidade_Fabricante(IndFabricante) + dneg!pedquantidadePedida
         
   If txtCEFOP = 5949 Or txtCEFOP = 6949 Then
       Acumula_Comissao_Rep = 0
       Acumula_Comissao_Promot = 0
   Else
       Acumula_Comissao_Rep = Acumula_Comissao_Rep + dneg!pedcomissaorep
       Acumula_Comissao_Promot = Acumula_Comissao_Promot + dneg!pedcomissaopromot
       'Na segunda fase usamos o acumulado por fabricante para gerar o valor da comissao isoladamente
       Acumula_Comis_Rep(IndFabricante) = Acumula_Comis_Rep(IndFabricante) + dneg!pedcomissaorep
       Acumula_Comis_Promot(IndFabricante) = Acumula_Comis_Promot(IndFabricante) + dneg!pedcomissaopromot
   End If
   Valor_Total = 0
   Valor_Produto = 0
   Valor_Frete = 0
   Valor_IPI = 0
   Valor_Desconto = 0
                  
   dneg.MoveNext
   If dneg.EOF Then
      Fim_Carga = 1
   Else
      Linha = Linha + 1
   End If


'Acumula_Comissao_Promot = Acumula_Comis_Promot(0) + Acumula_Comis_Promot(1)


'If TabNegociacao("NEGCOBRANCAFRETE") = 4 Or TabNegociacao("NEGCOBRANCAFRETE") = 5 Or TabNegociacao("NEGCOBRANCAFRETE") = 7 Then
'   Acumula_Valor_Frete = TabNegociacao("negvalorfixofrete")
'   txtResTotalFrete = TabNegociacao("negvalorfixofrete")
'End If

txtValorTotal = Format$(Acumula_Valor_Total, "#,##0.00")

'txtTotalPUQtd = Format$(Acumula_Valor_Produto, "#,##0.00")
txtResTotalProduto = Format$(Valor_Total_Cheio, "#,##0.00")
'txtResTotalFrete = Format$(Acumula_Valor_Frete, "##0.00")
'txtTotalIPI = Format$(Acumula_Valor_IPI, "##0.00")
'txtResTotalIPI = Format$(Acumula_Valor_IPI, "##0.00")
'txtTotalDesc = Format$(Acumula_Valor_Desconto, "##0.00")
txtResTotalDesconto = Format$(Acumula_Valor_Desconto, "#,##0.00")
'txtTotalQtd = Acumula_Quantidade
txtResTotalICMS = Format$(((Acumula_Valor_Produto + Acumula_Valor_Frete) - Acumula_Valor_Desconto) * (ICM!icmAliquota / 100), "#0.00")

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If

Loop
txtResValorTotalNF = Format$(Acumula_Valor_Total, "#,##0.00")
'Guarda_Cliente = pes!chPessoa

'If frmPedido.lblTransporte = "Cliente" Or txtPlaca = "N/Inf." Then
'   txtTransportadora = frmPedido.lblTransporte
'   txtMotorista = "Cliente"
'   txtIdentidade = "Cliente"
'   txtPlaca = "Cliente"
'   txtUFPlaca = "Cliente"
'   txtMunicipioPlaca = "Cliente"
'Else
'   Tabpessoa.Seek "=", frmPedido.lblTransporte
'   If Tabpessoa.NoMatch Then
'      MsgBox "Não informado o Transportador"
'      Exit Sub
'   End If
'   TabCompTransporte.Seek "=", Tabpessoa("chPessoa"), frmPedido.lblPlaca
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
'   txtPlaca = frmPedido.lblPlaca
'   txtTransportadora = frmPedido.lblTransporte
'   txtMotorista = Tabpessoa("pesRazaoSocial")
'   txtIdentidade = Tabpessoa("chCNPJ_CPF")
   
'End If

txtRepresentante = CartRep!chPessoa

'txtPromotora = CartPromot!chPessoa
'pes.Seek "=", Guarda_Cliente

'Alteração para implementação da rotina de Consignação da C&C

'If txtCondProcess = "CONSIGNAÇÃO" Then
'   txtComissaoRep = Format(0, "#,##0.00")
'   txtComissaoPromot = Format(0, "#,##0.00")
'Else
'   txtComissaoRep = Format(Acumula_Comissao_Rep, "#,##0.00")
'   txtComissaoPromot = Format$(Acumula_Comissao_Promot, "#,##0.00")
'End If
'Rotina de Calculo da distribuicao do faturamento

If txtFaturamento = 0 Then
   Base_Fatura = 0
Else
   Base_Fatura = Format$(Acumula_Valor_Total / txtFaturamento, "##,##0.00")
End If

Base_Frete = Acumula_Valor_Frete

If txtFaturamento = 0 Then
   txtFaturamento = 1
End If

If (Acumula_Total_Fabricante(0)) > 0 Then
   Base_Fatura_Lart = Format$(Acumula_Total_Fabricante(0) / txtFaturamento, "##,##0.00")
End If

If (Acumula_Total_Fabricante(1)) > 0 Then
   Base_Fatura_Merco = Format$(Acumula_Total_Fabricante(1) / txtFaturamento, "##,##0.00")
End If

'If pes!pesicms_st = 1 Then
'   Base = 1
'Else
'   Base = 0
'End If

'If TabCobrancaFrete("chcodcobrancafrete") = 1 Then 'Indica que a boleta de frete é separada e portanto o frete ficara na primeira ocorrencia da tabela
'   Ind_Aux = 1 + Base
'   Ind_Inicial_Frete = 1 + Base
'   Ind_Inicial_Fatura = 2 + Base
'   Parcelas_Frete = 1 + Base
'   Data_Cobranca = txtDataFrete
'Else
'   Ind_Aux = 0 + Base
'   Ind_Inicial_Fatura = 1 + Base
'   Ind_Inicial_Frete = 1 + Base
'   If TabCobrancaFrete("chcodcobrancafrete") = 3 Then
'      Parcelas_Frete = txtFaturamento
 ''  Else
  '    If TabCobrancaFrete("chcodcobrancafrete") = 7 Then
  '       Parcelas_Frete = 0
  '    Else
  '       Parcelas_Frete = 1 + Base
  '    End If
  ' End If
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


'Carga do gridFatura

'If pes!pesicms_st = 1 Then
'   ICMS_ST = txtResValorTotalNF + (txtResValorTotalNF * 35) / 100
'   ICMS_ST = ((ICMS_ST * pes!icmaliquotanoestado) / 100)
'   ICMS_ST = Format$(ICMS_ST - txtResTotalICMS, "##,##0.00")
'   GridFatura.Rows = A + 1
'   GridFatura.TextMatrix(1, 0) = "ICMS-ST NF-" & txtNotaFiscal
'   GridFatura.TextMatrix(1, 1) = txtNotaFiscal
'   GridFatura.TextMatrix(1, 2) = "ICMS-ST NF-"
'   'gridFatura.TextMatrix(A, 3) = Data_Cobranca + Intervalo
'   GridFatura.TextMatrix(1, 3) = Data_Proc + 7
'   GridFatura.TextMatrix(1, 4) = Format$(ICMS_ST, "#,##0.00")
'   GridFatura.TextMatrix(1, 5) = Format$(0, "#,##0.00")
'   GridFatura.TextMatrix(1, 6) = Format$(ICMS_ST, "#,##0.00")
'Else
'   Base = 0
'End If

'If neg!NEGCOBRANCAFRETE = 7 And (Fatura_Lart(A) + Fatura_Merco(A)) = 0 Then
''   Frete_Fatura(A) = 0
''Else
    For A = Base + 1 To (txtFaturamento + Ind_Aux)
       
    'Alteração: "Fatura" pelo número da Nota Fiscal em descrição
       GridFatura.Rows = A + 1
   

   
       If (Fatura_Lart(A) + Fatura_Merco(A)) > 0 And Frete_Fatura(A) > 0 Then
          GridFatura.TextMatrix(A, 0) = "NF-" & txtNotaFiscal & "-" & (A - Ind_Aux) & "/" & txtFaturamento & (" + Frete")
       Else
          If (Fatura_Lart(A) + Fatura_Merco(A)) = 0 Then
             GridFatura.TextMatrix(A, 0) = "Frete da NF " & txtNotaFiscal
          Else
          'Alteracao importante. Verificar
             If txtCEFOP = 5949 Or txtCEFOP = 6949 Then
                'Fatura_Lart(A) = 0
                GridFatura.TextMatrix(A, 0) = "NF-" & txtNotaFiscal & "-" & (A - Ind_Aux) & "/" & "LD"
             Else
                GridFatura.TextMatrix(A, 0) = "NF-" & txtNotaFiscal & "-" & (A - Ind_Aux) & "/" & txtFaturamento
             End If
          End If
       End If
    
       GridFatura.TextMatrix(A, 1) = txtNotaFiscal
       GridFatura.TextMatrix(A, 2) = (A - Ind_Aux)
       GridFatura.TextMatrix(A, 3) = Data_Cobranca
       'GridFatura.TextMatrix(A, 3) = Data_Cobranca
       GridFatura.TextMatrix(A, 4) = Format$(Fatura_Lart(A) + Frete_Lart(A), "#,##0.00")
       GridFatura.TextMatrix(A, 5) = Format$(Fatura_Merco(A) + Frete_Merco(A), "#,##0.00")
       GridFatura.TextMatrix(A, 6) = Format$((Fatura_Lart(A) + Fatura_Merco(A)) + (Frete_Lart(A) + Frete_Merco(A)), "#,##0.00")
       'Alterado em 18/09/2003
       'If TabCobrancaFrete("chcodcobrancafrete") = 1 And (A - Base) = 1 Then
       '   Data_Cobranca = txtDataPrimParc
       'Else
       '   Intervalo = TabNegociacao("negintervalofatura")
       'End If
       IndFatSalvo = A
       Data_Cobranca = Data_Cobranca + txtIntervalo
       'Intervalo = TabNegociacao("negintervalofatura")
       Acumula_Fatura_Lart = Format$(Acumula_Fatura_Lart + Fatura_Lart(A) + Frete_Lart(A), "#,##0.00")
       Acumula_Fatura_Merco = Format$(Acumula_Fatura_Merco + Fatura_Merco(A) + Frete_Merco(A), "#,##0.00")
       Acumula_Fatura_Geral = Format$(Acumula_Fatura_Geral + Valor_Fatura(A) + Frete_Fatura(A), "#,##0.00")
    Next

'txtTotalFaturaLart = Format$(Acumula_Fatura_Lart, "#,##0.00")
'txtTotalFaturaMerco = Format$(Acumula_Fatura_Merco, "#,##0.00")
txtTotalFatura = Format$((Acumula_Fatura_Lart + Acumula_Fatura_Merco), "#,##0.00")

Acumula_Valor_Total = 0
Acumula_Valor_Produto = 0
Acumula_Valor_Frete = 0
Acumula_Valor_IPI = 0
Acumula_Valor_Desconto = 0
Acumula_Quantidade = 0
Valor_Total_Cheio = 0

cmdCancela.Enabled = True
cmdConf.Enabled = True

Call FechaDB

End Sub

Public Sub Rotina_Limpa_Form()

txtCliente = Empty
txtNumPedido = Empty
txtCompPedido = Empty
txtDataPedido = "__/__/____"
txtNotaFiscal = Empty
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
'txtTotalFaturaLart = Empty
'txtTotalFaturaMerco = Empty
txtTotalFatura = Empty
txtTransportadora = Empty
txtOrdemDeCarga = Empty
txtMotorista = Empty
txtIdentidade = Empty
txtMunicipioPlaca = Empty
txtUFPlaca = Empty
txtResTotalProduto = Empty
txtResTotalFrete = Empty
'txtResTotalIPI = Empty
txtResTotalICMS = Empty
txtResTotalDesconto = Empty
txtResValorTotalNF = Empty
txtRepresentante = Empty
txtComissaoRep = Empty
txtPromotora = Empty
txtComissaoPromot = Empty
txtCEFOP = Empty
txtNatOperacao = Empty

Acumula_Valor_Total = 0
Acumula_Valor_Produto = 0
Acumula_Valor_Frete = 0
Acumula_Valor_IPI = 0
Acumula_Valor_Desconto = 0
Acumula_Quantidade = 0
Valor_Total_Cheio = 0

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
    GridNegocio.TextMatrix(Ind, 7) = Empty
    GridNegocio.TextMatrix(Ind, 8) = Empty
   
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
'Else
'   TabCtaPagar("ctpDescricaoOperacao") = "Comissão de Promotor"
'End If

'TabCtaPagar("chano") = Ano_Comissao

'TabCtaPagar("chmes") = Mes_Comissao

'TabCtaPagar("chdia") = 1
'TabBanco.Seek "=", 0, frmPedido.cmbBanco.ListIndex
'If TabBanco.NoMatch Then
'   MsgBox ("Banco inválido"), , frmPedido.cmbBanco
'   Mes_Comissao = 1 / 0
'End If
'txtBancoFat = frmPedido.cmbBanco
'TabCtaPagar("chcodbcolart") = TabBanco("bcosiglabco")
'End Sub
Private Sub GridFatura_Click()

End Sub
