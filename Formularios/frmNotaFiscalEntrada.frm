VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNotaFiscalEntrada 
   BackColor       =   &H00E0E0E0&
   Caption         =   "frmNotaFiscalEntrada"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18015
   LinkTopic       =   "Form3"
   ScaleHeight     =   9450
   ScaleWidth      =   18015
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox txtDataHoje 
      Height          =   375
      Left            =   9480
      TabIndex        =   73
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   0
      BackColor       =   -2147483637
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
   Begin VB.Frame Frame9 
      Caption         =   "Frame9"
      Height          =   15
      Left            =   3960
      TabIndex        =   62
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   17895
      Begin VB.Frame Frame4 
         Caption         =   "Vencimento/Desdobramento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   120
         TabIndex        =   51
         Top             =   5400
         Width           =   17655
         Begin VB.TextBox txtIPTE 
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
            Left            =   6240
            TabIndex        =   91
            Top             =   480
            Visible         =   0   'False
            Width           =   9255
         End
         Begin VB.CommandButton cmdNovoDesdob 
            BackColor       =   &H00FFFF00&
            Caption         =   "Novo Desdobramento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   2640
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker txtDataVencito 
            Height          =   405
            Left            =   2040
            TabIndex        =   26
            Top             =   480
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   714
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
            Format          =   410320897
            CurrentDate     =   43882
         End
         Begin MSFlexGridLib.MSFlexGrid GridDesdobr 
            Height          =   1575
            Left            =   120
            TabIndex        =   86
            Top             =   960
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   2778
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            BackColor       =   16777152
            BackColorFixed  =   16776960
            BackColorBkg    =   16777152
            FormatString    =   "N.Fatura          |Vencito           |Valor               |"
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
         Begin VB.Frame Frame7 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Navegação"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Left            =   15840
            TabIndex        =   74
            Top             =   600
            Width           =   1695
            Begin VB.CommandButton cmdNavega 
               BackColor       =   &H008080FF&
               Caption         =   "Último"
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
               Index           =   3
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   78
               Top             =   2160
               Width           =   1455
            End
            Begin VB.CommandButton cmdNavega 
               BackColor       =   &H00FFFF80&
               Caption         =   "Anterior"
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
               Index           =   2
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   77
               Top             =   1560
               Width           =   1455
            End
            Begin VB.CommandButton cmdNavega 
               BackColor       =   &H0000FFFF&
               Caption         =   "Próximo"
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
               Index           =   1
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   76
               Top             =   960
               Width           =   1455
            End
            Begin VB.CommandButton cmdNavega 
               BackColor       =   &H0080FF80&
               Caption         =   "Primeiro"
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
               Index           =   0
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   75
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.CommandButton cmdSair 
            BackColor       =   &H000000FF&
            Caption         =   "Sair"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   17.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   12840
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Frame Frame8 
            Caption         =   "Procedimento Final"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   12840
            TabIndex        =   60
            Top             =   960
            Width           =   2655
            Begin VB.CommandButton cmdGeraCtaPagar 
               BackColor       =   &H80000002&
               Caption         =   "Gera Cta a Pagar"
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
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   360
               Width           =   2295
            End
            Begin VB.CommandButton cmdNovaNotaFiscal 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Nova Nota Fiscal"
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
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   840
               Width           =   2295
            End
            Begin VB.CommandButton cmdCancelar 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Cancela N. Fiscal"
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
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   1320
               Width           =   2295
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Resumo do Lançamento"
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
            Left            =   8280
            TabIndex        =   56
            Top             =   960
            Width           =   4575
            Begin VB.Label txtDiferenca 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   2040
               TabIndex        =   66
               Top             =   1920
               Width           =   2415
            End
            Begin VB.Label txtValorTotalNota 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   2040
               TabIndex        =   65
               Top             =   1200
               Width           =   2415
            End
            Begin VB.Label txtValorTotalFatura 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   2040
               TabIndex        =   64
               Top             =   480
               Width           =   2415
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Diferença"
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
               Left            =   240
               TabIndex        =   59
               Top             =   1920
               Width           =   1335
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Calculado"
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
               Left            =   240
               TabIndex        =   58
               Top             =   480
               Width           =   1410
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Informado"
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
               Left            =   240
               TabIndex        =   57
               Top             =   1200
               Width           =   1410
            End
         End
         Begin VB.CommandButton cmdAlteraFatura 
            BackColor       =   &H008080FF&
            Caption         =   "Refazer Vencito."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   2640
            Width           =   2535
         End
         Begin VB.CommandButton cmdIncluiFatura 
            BackColor       =   &H0000FF00&
            Caption         =   "Inclui Vencimento Desdobramento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   6240
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtValorFatura 
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
            Height          =   435
            Left            =   3840
            TabIndex        =   27
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtFatura 
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
            TabIndex        =   25
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label lblIPTE1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código de Barras"
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
            Left            =   6240
            TabIndex        =   90
            Top             =   240
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.Label lblIPTE 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   300
            Left            =   6120
            TabIndex        =   89
            Top             =   240
            Width           =   45
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Valor Fatura"
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
            Left            =   4080
            TabIndex        =   54
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "N. Fatura"
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
            TabIndex        =   53
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Label18 
            Caption         =   "Data Vencito."
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
            Left            =   2040
            TabIndex        =   52
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame frame5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Produtos da Nota Fiscal de Entrada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   0
         TabIndex        =   40
         Top             =   1560
         Width           =   17775
         Begin MSFlexGridLib.MSFlexGrid GridProduto 
            Height          =   2055
            Left            =   120
            TabIndex        =   85
            Top             =   1320
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   3625
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            BackColor       =   16777152
            BackColorFixed  =   16776960
            BackColorBkg    =   16777152
            FormatString    =   "Cod. Produto        |Descrição do Produto                                 |Unid     |Qtd.    |P.U         |Valor do Produto ||"
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
         Begin VB.TextBox txtPU 
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
            Left            =   9360
            TabIndex        =   18
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtQtd 
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
            Left            =   7560
            TabIndex        =   17
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton cmdCadastrar 
            BackColor       =   &H0080C0FF&
            Caption         =   "Alterar Base da Nota"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   15480
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton cmdVencimento 
            BackColor       =   &H00FFFF00&
            Caption         =   "Vencimento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   15480
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1920
            Width           =   1815
         End
         Begin VB.CommandButton cmdNovoProduto 
            BackColor       =   &H0080FFFF&
            Caption         =   "Novo Ítem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   13680
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton cmdExclui 
            BackColor       =   &H000000FF&
            Caption         =   "Excluir"
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
            Left            =   13680
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1920
            Width           =   1335
         End
         Begin VB.CommandButton cmdAltera 
            BackColor       =   &H00C0C000&
            Caption         =   "Alterar"
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
            Left            =   13680
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton cmdInclui 
            BackColor       =   &H0000FF00&
            Caption         =   "Incluir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   13680
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtValor 
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
            Left            =   10680
            TabIndex        =   19
            Top             =   840
            Width           =   2175
         End
         Begin VB.ComboBox cmbCodProduto 
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
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label lblQtdTotal 
            Alignment       =   2  'Center
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
            Left            =   8640
            TabIndex        =   84
            Top             =   3360
            Width           =   1935
         End
         Begin VB.Label lblValorTotaldoProduto 
            Alignment       =   1  'Right Justify
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
            Left            =   10680
            TabIndex        =   83
            Top             =   3360
            Width           =   2175
         End
         Begin VB.Label lblUnidade 
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
            Left            =   8520
            TabIndex        =   82
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblProdutoFabrica 
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
            Left            =   2400
            TabIndex        =   81
            Top             =   840
            Width           =   5055
         End
         Begin VB.Label lblDescProdutoEntrada 
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
            Left            =   2400
            TabIndex        =   80
            Top             =   480
            Width           =   4935
         End
         Begin VB.Label Label11 
            Caption         =   "Valor do Produto"
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
            Left            =   10680
            TabIndex        =   46
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "P.U."
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
            TabIndex        =   45
            Top             =   600
            Width           =   525
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Un."
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
            Left            =   8640
            TabIndex        =   44
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
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
            Left            =   7680
            TabIndex        =   43
            Top             =   600
            Width           =   450
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Desc. Produto/Centro de Custo"
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
            Left            =   2400
            TabIndex        =   42
            Top             =   240
            Width           =   3765
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Produto"
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
            TabIndex        =   41
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   17775
         Begin VB.ComboBox cmbDespFornec 
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
            TabIndex        =   1
            Top             =   360
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker txtDataEmissao 
            Height          =   405
            Left            =   120
            TabIndex        =   87
            Top             =   960
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   714
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
            Format          =   410320897
            CurrentDate     =   43882
         End
         Begin VB.TextBox txtValorDaNotaFiscal 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "##,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
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
            Height          =   405
            Left            =   8760
            TabIndex        =   10
            Top             =   960
            Width           =   2295
         End
         Begin VB.ComboBox txtNotaFiscal 
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
            TabIndex        =   3
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox cmbBanco 
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
            Left            =   15840
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtQtdFaturas 
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
            Height          =   435
            Left            =   14640
            TabIndex        =   13
            Top             =   1080
            Width           =   855
         End
         Begin VB.ComboBox cmbFinalidade 
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
            Left            =   11160
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   360
            Width           =   2895
         End
         Begin VB.Frame Frame10 
            Caption         =   "Emissão de Recibo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   11160
            TabIndex        =   70
            Top             =   720
            Width           =   2895
            Begin VB.OptionButton Option1 
               Caption         =   "Sim"
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
               Left            =   240
               TabIndex        =   11
               Top             =   360
               Width           =   975
            End
            Begin VB.OptionButton OptNao 
               Caption         =   "Não"
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
               Left            =   1680
               TabIndex        =   12
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton OptSim 
               Caption         =   "Sim"
               Height          =   195
               Left            =   -600
               TabIndex        =   35
               Top             =   840
               Width           =   855
            End
         End
         Begin VB.ComboBox cmbTipoLancamento 
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
            ItemData        =   "frmNotaFiscalEntrada.frx":0000
            Left            =   8280
            List            =   "frmNotaFiscalEntrada.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtValorIPI 
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
            Left            =   4680
            TabIndex        =   8
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtValorFrete 
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
            Left            =   6600
            TabIndex        =   9
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtValorICMS 
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
            Left            =   2640
            TabIndex        =   7
            Top             =   960
            Width           =   1455
         End
         Begin VB.ComboBox cmbFabrica 
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
            ItemData        =   "frmNotaFiscalEntrada.frx":0004
            Left            =   14640
            List            =   "frmNotaFiscalEntrada.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   360
            Width           =   2415
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
            Left            =   2280
            TabIndex        =   2
            Top             =   360
            Width           =   3975
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Lançamento"
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
            Index           =   1
            Left            =   120
            TabIndex        =   88
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label27 
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
            Left            =   15840
            TabIndex        =   79
            Top             =   840
            Width           =   780
         End
         Begin VB.Label Label21 
            Caption         =   "Qtd Faturas"
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
            Left            =   14280
            TabIndex        =   72
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Finalidade"
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
            Left            =   11160
            TabIndex        =   71
            Top             =   120
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   " Número Doc"
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
            Index           =   0
            Left            =   6360
            TabIndex        =   69
            Top             =   120
            Width           =   1560
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento"
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
            Left            =   8280
            TabIndex        =   68
            Top             =   120
            Width           =   1980
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "IPI"
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
            Left            =   4680
            TabIndex        =   55
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Valor Frete"
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
            Left            =   6600
            TabIndex        =   50
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Valor ICMS"
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
            Left            =   2640
            TabIndex        =   49
            Top             =   720
            Width           =   1380
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total da Nota"
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
            Left            =   8760
            TabIndex        =   47
            Top             =   720
            Width           =   2340
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Sacado"
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
            Left            =   14640
            TabIndex        =   39
            Top             =   120
            Width           =   930
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
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
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fornecedor/Despesa"
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
            TabIndex        =   37
            Top             =   120
            Width           =   2535
         End
      End
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "Nota Fiscal Número"
      Height          =   195
      Left            =   2040
      TabIndex        =   67
      Top             =   960
      Width           =   1395
   End
   Begin VB.Label lblStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5400
      TabIndex        =   63
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "Status"
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
      Left            =   4440
      TabIndex        =   61
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   495
      Left            =   5400
      TabIndex        =   48
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "frmNotaFiscalEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Acumula_Fatura As Currency
Dim QtdVezes As Integer
Dim Ind As Byte
Dim fim As Byte
Dim Incluir As Byte
Dim IncluiDesdob As Byte
Dim ContadorCtaPagar As Byte
Dim Data_Vencimento As Date
Dim DataUtil As Date
Dim DataHoje As Date
Dim DataBanco As Date
Dim Status As Byte
Dim cadastrar As Byte
Dim Resp As String
Dim MaiorQueUm As Byte
Dim Acumula_Valor As Currency
Dim Acumula_Qtd As Integer
Dim LimiteCarga As Integer
Dim LimiteProduto As Integer
Dim Linha As Integer
Dim Coluna As Integer
Dim Fatura As String
Dim GerarCredito As Byte
Dim IncluiNotaFiscal As Byte
Dim Interno As Byte
Dim ValorAnterior As Currency
Dim DespesaAnterior As String



Private Sub cmbCodProduto_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub cmbDespFornec_LostFocus()


Call Rotina_AbrirBanco

cmbPessoa.Clear

pes.Open "Select * from Pessoa", db, 3
           
If cmbDespFornec = "FORNECEDOR" Then
    pes.MoveFirst
    If pes.EOF Then
       MsgBox ("Dataset Pessoa sem registro. Informar ao administrador do sistema"), vbCritical
       Call FechaDB
       Exit Sub
    End If
    
    pes.MoveFirst
    
    Do While Not pes.EOF
    
       If pes!pestipopessoa = 1 Or pes!pestipopessoa = 2 Or pes!pestipopessoa = 5 Then
          cmbPessoa.AddItem pes!chPessoa
       End If
       pes.MoveNext
    Loop
Else
    ProdFornec.Open "Select * from ProdutoFornecedor", db, 3, 3
    If ProdFornec.EOF Then
       MsgBox ("Erro. Tabela de Produto Fornecedor vazia. Comunicar ao analista responsável."), vbCritical
       Call FechaDB
       Exit Sub
    End If
    ProdFornec.MoveFirst
    Do While Not ProdFornec.EOF
       If Not (ProdFornec!chTipoProduto = DespesaAnterior) Then
          cmbPessoa.AddItem ProdFornec!chTipoProduto
          DespesaAnterior = ProdFornec!chTipoProduto
       End If
          
       ProdFornec.MoveNext

    Loop
End If

Call FechaDB

End Sub

Private Sub cmbPessoa_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub cmbPessoa_LostFocus()

txtNotaFiscal.Clear
txtNotaFiscal = Empty
cmbFabrica.ListIndex = 0
If cmbPessoa = "" Then
   cmdSair.Enabled = True
   cmdSair.SetFocus
   Exit Sub
End If

Call Rotina_AbrirBanco
If cmbPessoa = "" Then
   MsgBox ("Não Informado Fornecedor ou Despesa."), vbCritical
   Call FechaDB
   Exit Sub
End If

IncluiNotaFiscal = 0

nfe.Open "Select * from NotaFiscalEntrada where chPessoa = ('" & cmbPessoa & "')", db, 3, 3
If nfe.EOF Then
    IncluiNotaFiscal = 1
Else
    nfe.MoveFirst
    
    Do While Not nfe.EOF
       If nfe!chPessoa = cmbPessoa Then
          txtNotaFiscal.AddItem nfe!chNotaFiscalEntrada
          nfe.MoveNext
       Else
           nfe.MoveNext
       End If
    Loop
End If
'txtNotaFiscal.ListIndex = 0

cmbCodProduto = Empty
cmbCodProduto.Clear

ProdEntrada.Open "Select * from ProdutoEntrada where chPessoa = ('" & cmbPessoa & "')", db, 3, 3
If Not ProdEntrada.EOF Then
   ProdEntrada.MoveFirst
   Do While Not ProdEntrada.EOF
      cmbCodProduto.AddItem ProdEntrada!chTipoProduto
      ProdEntrada.MoveNext
    Loop
End If
If Not cmbCodProduto.ListIndex < 0 Then
   cmbCodProduto.ListIndex = 0
End If
End Sub
Private Sub cmbCodProduto_LostFocus()

Verifica = Empty
Verifica = Mid$(cmbCodProduto, 30, 5)
If Not Verifica = Empty Then
   MsgBox ("Código do Produto Informado ultrapassa 30 caracteres.")
   cmbCodProduto.SetFocus
   Exit Sub
End If

If cmbCodProduto = Empty Then
   If Incluir = 1 Then
      MsgBox ("Informar o código do produto ou cancelar a operação")
      cmdSair.SetFocus
   End If
End If

'TabFabFornec = ProdutoEntrada
Call Rotina_AbrirBanco

ProdEntrada.Open "Select * from ProdutoEntrada where chPessoa = ('" & cmbPessoa & "') and chTipoProduto = ('" & cmbCodProduto & "')", db, 3, 3
If ProdEntrada.EOF Then
   Resp = MsgBox("Produto não cadastrado. Deseja cadastra-lo???", vbYesNo)
   If Resp = vbYes Then
      frmProdutosDeEntrada.cmbFornecedor = frmNotaFiscalEntrada.cmbPessoa
     ' frmProdutosDeEntrada.lblCodProduto = frmNotaFiscalEntrada.cmbCodProduto
      frmProdutosDeEntrada.cmbTipoProduto = frmNotaFiscalEntrada.cmbCodProduto
      frmProdutosDeEntrada.Show vbModal
      cmbCodProduto.SetFocus
   End If
Else
    lblDescProdutoEntrada = ProdEntrada!pindescricao
    lblProdutoFabrica = ProdEntrada!chProdutoFabrica
    lblUnidade = ProdEntrada!pinunidade
End If
dnfe.Open "Select * from NotaFiscalDetProd where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & cmbCodProduto & "')", db, 3, 3
If dnfe.EOF Then
   cmdInclui.Enabled = True
   cmdAltera.Enabled = False
   cmdExclui.Enabled = False
   
Else
   txtQtd = dnfe!nfdQtd
   txtPU = dnfe!nfdPU
   txtValor = dnfe!nfdValorDaCompra
   cmdInclui.Enabled = False
   cmdAltera.Enabled = True
   cmdExclui.Enabled = True
   
End If

Call FechaDB

End Sub


Private Sub cmbTipoLancamento_LostFocus()
If cmbTipoLancamento = "BOLETO" Then
   lblIPTE.Visible = True
   lblIPTE1.Visible = True
   txtIPTE.Visible = True
   txtIPTE = Empty
Else
   lblIPTE1.Visible = False
   txtIPTE.Visible = False
End If

End Sub

Private Sub cmdAltera_Click()


Call Rotina_AbrirBanco

db.BeginTrans

   dnfe.Open "Select * from NotaFiscalDetProd where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & cmbCodProduto & "')", db, 3, 3
   If dnfe.EOF Then
      MsgBox ("Erro no acesso a detalhe de Nota Fiscal"), vbCritical
      db.CommitTrans
      Call FechaDB
      Exit Sub
   End If
   
   Call Rotina_051_Doc_DBDET
      
   dnfe.Update

db.CommitTrans

Call Rotina_020_Limpa_Grid_Produto
Call Rotina_030_Carga_Grid_Produto

Call FechaDB
   
End Sub

Private Sub cmdAlteraFatura_Click()

If Interno = 0 Then
    Resp = MsgBox("Voce realmente deseja Alterar este Desdobramento????", vbYesNo)
    If Resp = vbNo Then
       cmdCancelar.Enabled = False
       cmdSair.Enabled = True
       cmdSair.SetFocus
       Exit Sub
    End If
Else
   Interno = 0
End If

Call Rotina_054_Deleta_Financ_Desdob

Call Rotina_AbrirBanco
db.BeginTrans
nfe.Open "Select * from NotaFiscalEntrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
If nfe.EOF Then
   MsgBox ("Nota Fiscal não encontrada para Alteração de Fatura"), vbCritical
   db.CommitTrans
   Call FechaDB
   Exit Sub
End If
   
nfe!nfeStatus = 0
nfe.Update

db.CommitTrans

Call Rotina_21_Limpa_Grid_Desdobr
Call Rotina_026_Limpa_Det_Desd
Call Rotina_027_limpa_totalizacao
lblStatus.Caption = "Pendente"
cmdGeraCtaPagar.Enabled = True
cmbBanco.ListIndex = 0
txtFatura.SetFocus

Call FechaDB

End Sub

Private Sub cmdCadastrar_Click()
Incluir = 0

If ValorAnterior = txtValorDaNotaFiscal Then
   Interno = 0
Else
   Interno = 1
   MsgBox ("Todo faturamento deste lançamento será excluído. Rafazer faturamento e respectivos vencimentos"), vbInformation
   Call cmdAlteraFatura_Click
End If

Call Rotina_AbrirBanco

If cmbPessoa = Empty Then
   MsgBox ("Caso queira Incluir Informações para Débito, Informe o Colaborador")
   cmdNovaNotaFiscal.SetFocus
   Exit Sub
End If

If txtNotaFiscal = Empty Then
   MsgBox ("Número da Nota Fiscal não informado")
   txtNotaFiscal.SetFocus
   Exit Sub
End If

If txtValorDaNotaFiscal = Empty Then
   MsgBox ("Valor Total da Nota Fiscal não Informado.")
   txtValorDaNotaFiscal.SetFocus
Else
   If Not IsNumeric(txtValorDaNotaFiscal) Then
      MsgBox ("Valor da Nota Fiscal Inválido")
      cmdSair.SetFocus
      Exit Sub
   End If
End If

If txtValorICMS = Empty Then
   Resp = MsgBox("Valor do ICMS não informado. Deseja Informá-lo?", vbYesNo)
   If Resp = vbYes Then
      txtValorICMS.SetFocus
      Exit Sub
   Else
      txtValorICMS = Format$(0, "0.00")
   End If
Else
   If Not IsNumeric(txtValorICMS) Then
      MsgBox ("Valor do ICMS Inválido")
      cmdSair.SetFocus
      Exit Sub
   End If
End If
    
If txtValorIPI = Empty Then
   Resp = MsgBox("Valor do IPI não informado. Deseja Informá-lo?", vbYesNo)
   If Resp = vbYes Then
      txtValorIPI.SetFocus
      Exit Sub
   Else
      txtValorIPI = Format$(0, "0.00")
   End If
Else
   If Not IsNumeric(txtValorIPI) Then
      MsgBox ("IPI Inválido")
      cmdSair.SetFocus
      Exit Sub
   End If
End If
  
If txtValorFrete = Empty Then
   Resp = MsgBox("Valor do Frete não informado. Deseja Informá-lo?", vbYesNo)
   If Resp = vbYes Then
      txtValorFrete.SetFocus
      Exit Sub
   Else
      txtValorFrete = Format$(0, "0.00")
   End If
Else
   If Not IsNumeric(txtValorFrete) Then
      MsgBox ("Valor do Frete Inválido")
      cmdSair.SetFocus
      Exit Sub
   End If
End If

If txtDataEmissao = "__/__/____" Then
   MsgBox ("Data de emissão não informado")
   txtDataEmissao.SetFocus
End If

If Not (IsDate(txtDataEmissao)) Then
   MsgBox ("Data de Emissão inválida")
   txtDataEmissao.SetFocus
   Exit Sub
End If

If Not IsNumeric(txtQtdFaturas) Then
   MsgBox ("Qtd Faturas Inválido")
   cmdSair.SetFocus
   Exit Sub
End If

If txtQtdFaturas = Empty Then
   Resp = MsgBox("Quantidade de Faturas não Informada. Deseja Informá-la?", vbYesNo)
   If Resp = vbYes Then
      txtQtdFaturas.SetFocus
      Exit Sub
   Else
      QtdVezes = txtQtdFaturas
      txtQtdFaturas = Format$(0, "00")
   End If
End If

If cmbBanco = Empty Then
   MsgBox ("Informe o Banco")
   cmbBanco.SetFocus
End If
If cmbTipoLancamento = Empty Then
   MsgBox ("Informe o Tipo do Documento")
   cmbTipoLancamento.SetFocus
End If
If cmbFabrica = Empty Then
   MsgBox ("Informe a Empresa")
   cmbFabrica.SetFocus
End If
If cmbFinalidade = Empty Then
   MsgBox ("Informe a Finalidade da Despesa")
   cmbFinalidade.SetFocus
End If




'If Incluir = 1 Then

   Incluir = 0
   
   Call Rotina_AbrirBanco
   db.BeginTrans
   nfe.Open "Select * from NotaFiscalEntrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
   If nfe.EOF Then
      nfe.AddNew
   End If
   
   Call Rotina_050_Mover_Doc_DBNota
   
   nfe.Update
   
  db.CommitTrans
   
   'If dnfe.State = 1 Then
   '   dnfe.Close: Set dnfe = Nothing
   'End If
   
'   Call Rotina_AbrirBanco
   
'   dnfe.Open "Select * from NotaFiscalDetProd where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & cmbCodProduto & "')", db, 3, 3
'   If dnfe.EOF Then
'      dnfe.AddNew
'   End If
   
'   Call Rotina_051_Doc_DBDET

'   dnfe.Update
   

'db.CommitTrans

Call Rotina_020_Limpa_Grid_Produto

Call Rotina_025_Limpa_Detalhe

Call Rotina_030_Carga_Grid_Produto

cmdNovoProduto.SetFocus

'Call FechaDB


End Sub

Private Sub cmdCancelar_Click()
Resp = MsgBox("Voce realmente deseja cancelar esta nota????", vbYesNo)
If Resp = vbNo Then
   cmdCancelar.Enabled = False
   cmdSair.Enabled = True
   cmdSair.SetFocus
   Exit Sub
End If

Call Rotina_054_Deleta_Financ_Desdob

Call Rotina_055_Deleta_Detalhe_NF

Call Rotina_056_Deleta_Nota_Fiscal

Call Rotina_020_Limpa_Grid_Produto
Call Rotina_21_Limpa_Grid_Desdobr
Call Rotina_024_Limpa_Lancamento
Call Rotina_025_Limpa_Detalhe
Call Rotina_026_Limpa_Det_Desd
Call Rotina_027_limpa_totalizacao
lblStatus.Caption = Empty
cmdCancelar.Enabled = False
cmdGeraCtaPagar.Enabled = False
cmdSair.Enabled = True
'cmdNovaNotaFiscal.SetFocus

'cmbPessoa.SetFocus

End Sub


Private Sub cmdExclui_Click()
Resp = MsgBox("Excluisão de Registro. Confirma???", vbYesNo)
If Resp = vbYes Then
    
    Call Rotina_AbrirBanco
    
    dnfe.Open "Select * from NotaFiscalDetProd where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & GridProduto.TextMatrix(Linha, 0) & "')", db, 3, 3
    If dnfe.EOF Then
       MsgBox ("Erro no acesso a Detalhe de Nota Fiscal na rotina de alteração"), vbCritical
       Call FechaDB
       Exit Sub
    End If
   
    db.BeginTrans
        dnfe.Delete
    db.CommitTrans
    
    dnfe.Close: Set dnfe = Nothing

    Call Rotina_025_Limpa_Detalhe
    Call Rotina_020_Limpa_Grid_Produto
    Call Rotina_030_Carga_Grid_Produto
    cmbCodProduto.SetFocus
End If

End Sub

Private Sub cmdGeraCtaPagar_Click()
Dim DataInvertida As Double

Ind = 1
ContadorCtaPagar = 0

Do While Ind < GridDesdobr.Rows
       If cmbFinalidade = "X - ICMS FRETE" Then
          Data_Vencimento = Date
       Else
          Data_Vencimento = GridDesdobr.TextMatrix(Ind, 1)
       End If
       
       Call Rotina_AbrirBanco
       
       DataInvertida = Year(Data_Vencimento) & Format$(Month(Data_Vencimento), "00") & Format$(Day(Data_Vencimento), "00")
       
       nfd.Open "Select * from NotaFiscalDesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chDataVencimento = ('" & DataInvertida & "')", db, 3, 3
       If nfd.EOF Then
          MsgBox ("Erro no acesso a tabela de desdobramentos em Gerar Contas a Pagar"), vbCritical
          Call FechaDB
          Exit Sub
       Else
           
          ctp.Open "Select * from Contas_A_Pagar where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbPessoa & "') and chNotaFiscal = ('" & txtNotaFiscal & "') and chFatura = ('" & GridDesdobr.TextMatrix(Ind, 0) & "')", db, 3, 3
          If ctp.EOF Then
             
             db.BeginTrans
                
                Fatura = GridDesdobr.TextMatrix(Ind, 0)
                DataHoje = Date
                
                If GerarCredito = 0 Then
                   ctp.AddNew
                   Call Rotina_052_Gera_Cta_Pagar
                   ctp.Update
                Else
                   ctp.AddNew
                   Data_Vencimento = Date
                   DataHoje = Date
                   Call Rotina_052_Gera_Cta_Pagar
                   ctp.Update
                   Call Rotina_052_Gera_Cta_Receber
                   GerarCredito = 0
                End If
                
                nfd!nfdStatus = 1
                nfd.Update
                
                ContadorCtaPagar = ContadorCtaPagar + 1
                
                nfe.Open "Select * from NotaFiscalEntrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
                If nfe.EOF Then
                   MsgBox ("Nota Fiscal não encontrada em Gera Contas a Pagar."), vbCritical, txtNotaFiscal
                   Call FechaDB
                   Exit Sub
                End If
                          
                nfe("nfestatus") = 1
                nfe.Update
           
            db.CommitTrans
            ' MsgBox ("Contas a Pagar gerada com sucesso."), vbInformation
          Else
             MsgBox ("Contas a Pagar para esta NF/Fatura já foi gerada anteriormente.")
             txtFatura.SetFocus
          End If
          Ind = Ind + 1
       End If
Loop

'Call Rotina_053_Estoque_Geral

cmdGeraCtaPagar.Enabled = False
cmdNovaNotaFiscal.Enabled = True
cmdCancelar.Enabled = True
cmdSair.Enabled = True
cmdNovaNotaFiscal.SetFocus

Call FechaDB

End Sub

Private Sub cmdInclui_Click()

If cmbPessoa = Empty Then
   MsgBox ("Caso queira Incluir Informações para Débito, Informe o Colaborador")
   cmdNovaNotaFiscal.SetFocus
   Exit Sub
End If

If txtNotaFiscal = Empty Then
   MsgBox ("Número da Nota Fiscal não informado")
   txtNotaFiscal.SetFocus
   Exit Sub
End If

If txtValorDaNotaFiscal = Empty Then
   MsgBox ("Valor Total da Nota Fiscal não Informado.")
   txtValorDaNotaFiscal.SetFocus
Else
   If Not IsNumeric(txtValorDaNotaFiscal) Then
      MsgBox ("Valor da Nota Fiscal Inválido")
      cmdSair.SetFocus
      Exit Sub
   End If
End If

If txtValorICMS = Empty Then
   Resp = MsgBox("Valor do ICMS não informado. Deseja Informá-lo?", vbYesNo)
   If Resp = vbYes Then
      txtValorICMS.SetFocus
      Exit Sub
   Else
      txtValorICMS = Format$(0, "0.00")
   End If
Else
   If Not IsNumeric(txtValorICMS) Then
      MsgBox ("Valor do ICMS Inválido")
      cmdSair.SetFocus
      Exit Sub
   End If
End If
    
If txtValorIPI = Empty Then
   Resp = MsgBox("Valor do IPI não informado. Deseja Informá-lo?", vbYesNo)
   If Resp = vbYes Then
      txtValorIPI.SetFocus
      Exit Sub
   Else
      txtValorIPI = Format$(0, "0.00")
   End If
Else
   If Not IsNumeric(txtValorIPI) Then
      MsgBox ("IPI Inválido")
      cmdSair.SetFocus
      Exit Sub
   End If
End If
  
If txtValorFrete = Empty Then
   Resp = MsgBox("Valor do Frete não informado. Deseja Informá-lo?", vbYesNo)
   If Resp = vbYes Then
      txtValorFrete.SetFocus
      Exit Sub
   Else
      txtValorFrete = Format$(0, "0.00")
   End If
Else
   If Not IsNumeric(txtValorFrete) Then
      MsgBox ("Valor do Frete Inválido")
      cmdSair.SetFocus
      Exit Sub
   End If
End If

If txtDataEmissao = "__/__/____" Then
   MsgBox ("Data de emissão não informado")
   txtDataEmissao.SetFocus
End If

If Not (IsDate(txtDataEmissao)) Then
   MsgBox ("Data de Emissão inválida")
   txtDataEmissao.SetFocus
   Exit Sub
End If

If Not IsNumeric(txtQtdFaturas) Then
   MsgBox ("Qtd Faturas Inválido")
   cmdSair.SetFocus
   Exit Sub
End If

If txtQtdFaturas = Empty Or txtQtdFaturas = 0 Then
   Resp = MsgBox("Quantidade de Faturas não Informada. Deseja Informá-la?", vbYesNo)
   If Resp = vbYes Then
      txtQtdFaturas.SetFocus
      Exit Sub
   Else
      QtdVezes = txtQtdFaturas
      txtQtdFaturas = Format$(0, "00")
   End If
End If

If cmbBanco = Empty Then
   MsgBox ("Informe o Banco")
   cmbBanco.SetFocus
End If
If cmbTipoLancamento = Empty Then
   MsgBox ("Informe o Tipo do Documento")
   cmbTipoLancamento.SetFocus
End If
If cmbFabrica = Empty Then
   MsgBox ("Informe a Empresa")
   cmbFabrica.SetFocus
End If
If cmbFinalidade = Empty Then
   MsgBox ("Informe a Finalidade da Despesa")
   cmbFinalidade.SetFocus
End If
If lblProdutoFabrica = Empty Then
   MsgBox "Colocar o cursor em COD PRODUTO  e clicar em TAB"
   cmbCodProduto.SetFocus
   Exit Sub
End If

If txtQtd = Empty Then
   MsgBox ("Quantidade não informada")
   txtQtd.SetFocus
   Exit Sub
Else
   If Not IsNumeric(txtQtd) Then
      MsgBox ("Quantidade Inválida")
      cmdSair.SetFocus
      Exit Sub
   End If
End If

If txtPU = Empty Then
   MsgBox ("Valor unitário não informado")
   txtPU.SetFocus
   Exit Sub
Else
   If Not IsNumeric(txtPU) Then
      MsgBox ("PU Inválido")
      cmdSair.SetFocus
      Exit Sub
   End If
End If



If Incluir = 1 Then

   Incluir = 0

   Call Rotina_AbrirBanco

   db.BeginTrans

   nfe.Open "Select * from NotaFiscalEntrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
   If nfe.EOF Then
      nfe.AddNew
   End If
   
   Call Rotina_050_Mover_Doc_DBNota
   
   nfe.Update
   
   db.CommitTrans

End If
   
If dnfe.State = 1 Then
   dnfe.Close: Set dnfe = Nothing
End If

Call Rotina_AbrirBanco

db.BeginTrans
    
   dnfe.Open "Select * from NotaFiscalDetProd where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & cmbCodProduto & "')", db, 3, 3
   If dnfe.EOF Then
      dnfe.AddNew
   End If
   
   Call Rotina_051_Doc_DBDET
   
   dnfe.Update

db.CommitTrans

Call Rotina_020_Limpa_Grid_Produto

Call Rotina_025_Limpa_Detalhe

Call Rotina_030_Carga_Grid_Produto

cmdNovoProduto.SetFocus

'Call FechaDB

End Sub

Private Sub cmdIncluiFatura_Click()

If Not txtValorDaNotaFiscal = lblValorTotaldoProduto Then
   MsgBox ("Valor total da Nota diferente da soma dos valores do(s) produto(s)"), vbInformation
   txtFatura.SetFocus
   Exit Sub
End If

If txtFatura = Empty Then
   MsgBox ("Numero da Fatura não informado")
   txtFatura.SetFocus
   Exit Sub
End If

If txtDataVencito = "__/__/____" Then
   MsgBox ("Data de vencimento da fatura não informada")
   txtDataVencito.SetFocus
   Exit Sub
End If

If Not (IsDate(txtDataVencito)) Then
   MsgBox ("Data de vencimento inválida")
   txtDataVencito.SetFocus
   Exit Sub
End If

If txtValorFatura = Empty Then
   MsgBox ("Valor da Fatura não Informado")
   txtValorFatura.SetFocus
   Exit Sub
End If

Call Rotina_AbrirBanco

db.BeginTrans
    
 
   nfd.Open "select * from NotaFiscalDesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and nfdFaturaNumero = ('" & txtFatura & "')", db, 3, 3
   If nfd.EOF Then
      nfd.AddNew
   Else
      MsgBox ("Número da Fatura já cadastrado. "), vbCritical
      db.CommitTrans
      Call FechaDB
      Exit Sub
   End If
   nfd!chPessoa = cmbPessoa
   nfd!chNotaFiscalEntrada = txtNotaFiscal
   If cmbFinalidade = "X - ICMS FRETE" Then
      nfd!chDataVencimento = Date
   Else
      nfd!chDataVencimento = txtDataVencito
   End If
   nfd!nfdDataVencoriginal = txtDataVencito
   nfd!nfdFaturaNumero = txtFatura
   nfd!nfdValorDaFatura = txtValorFatura
   nfd!nfdIPTE = txtIPTE
    
   nfd.Update

db.CommitTrans

Call Rotina_21_Limpa_Grid_Desdobr
Call Rotina_026_Limpa_Det_Desd
Call Rotina_040_Carga_Grid_Desdobr

If txtValorTotalFatura > 0 And txtDiferenca = 0 Then
   cmdGeraCtaPagar.Enabled = True
   cmdCancelar.Enabled = True
   If ContadorCtaPagar > 0 Then
      cmdCancelar.Enabled = True
      cmdSair.Enabled = True
   Else
      cmdCancelar.Enabled = False
      cmdSair.Enabled = False
   End If
   cmdGeraCtaPagar.SetFocus
Else
   cmdGeraCtaPagar.Enabled = False
   cmdCancelar.Enabled = True
   cmdNovaNotaFiscal.Enabled = False
   cmdSair.Enabled = False
   txtFatura.SetFocus
End If

Call FechaDB

End Sub

Private Sub cmdNavega_Click(Index As Integer)
    
  Call Rotina_AbrirBanco
  
  nfe.Open "Select * from NotaFiscalEntrada", db, 3, 3
  If nfe.EOF Then
     MsgBox ("Não há Nota Fiscal Lançada até o momento, no período"), vbInformation
     Call FechaDB
     Exit Sub
  End If
   
    
   Select Case Index

   Case 0
        nfe.MoveFirst
   Case 1
        nfe.MoveNext
   Case 2
        nfe.MovePrevious
   Case 3
        nfe.MoveLast
        
   End Select

   If nfe.BOF = True Then
      nfe.MoveFirst
   End If
   
   If nfe.EOF = True Then
      nfe.MoveLast
   End If
   
   Call Rotina_024_Limpa_Lancamento
   Call Rotina_025_Limpa_Detalhe
   Call Rotina_026_Limpa_Det_Desd
   Call Rotina_027_limpa_totalizacao
   
   Call Rotina_010_Carga_Form
   
   cmdInclui.Enabled = False
   cmdNovoProduto.Enabled = True
   cmdAltera.Enabled = True
   cmdExclui.Enabled = True
   cmdSair.Enabled = True
   
   cmbPessoa.Enabled = False
   cmbTipoLancamento.SetFocus
   
   Call FechaDB
   
End Sub

Private Sub cmdNovaNotaFiscal_Click()

Call Rotina_020_Limpa_Grid_Produto
Call Rotina_21_Limpa_Grid_Desdobr
Call Rotina_024_Limpa_Lancamento
Call Rotina_025_Limpa_Detalhe
Call Rotina_026_Limpa_Det_Desd
Call Rotina_027_limpa_totalizacao
lblStatus.Caption = Empty
cmdGeraCtaPagar.Enabled = True
cmbPessoa.Enabled = True
cmbPessoa.SetFocus

End Sub

Private Sub cmdNovoDesdob_Click()
txtFatura = Empty
txtDataVencito = Date
txtValorFatura = 0
txtFatura.SetFocus
End Sub

Private Sub cmdNovoProduto_Click()

Call Rotina_025_Limpa_Detalhe

cmdVencimento.Enabled = True
cmbCodProduto.SetFocus

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdVencimento_Click()
Call Rotina_026_Limpa_Det_Desd
txtFatura.SetFocus
End Sub

Private Sub Form_Load()
txtDataHoje = Date
txtDataVencito = Date
txtDataEmissao = Date
lblStatus.Caption = Empty


cmbDespFornec.AddItem "FORNECEDOR"
cmbDespFornec.AddItem "DESPESA"
cmbDespFornec.ListIndex = 0
DespesaAnterior = Empty

Interno = 0

Call Rotina_AbrirBanco

Bco.Open "Select * from Banco", db, 3, 3
If Bco.EOF Then
   MsgBox ("Tabela Banco vazia"), vbCritical
   Call FechaDB
   Exit Sub
End If

Bco.MoveFirst

Do While Not Bco.EOF
   cmbBanco.AddItem Bco!bcosiglabco
   Bco.MoveNext
Loop

cmbBanco.ListIndex = 0
         
Emp.Open "Select * from Empresa", db, 3, 3
If Emp.EOF Then
   MsgBox ("Empresa não encontrada em inicialização de Nota Fiscal"), vbCritical
   Call FechaDB
   Exit Sub
End If

Emp.MoveLast
Do While Not Emp.BOF
   cmbFabrica.AddItem Emp!chPessoa
   Emp.MovePrevious
Loop
cmbFabrica.ListIndex = 0
Call Rotina_020_Limpa_Grid_Produto
Call Rotina_21_Limpa_Grid_Desdobr
Call Rotina_025_Limpa_Detalhe
Call Rotina_026_Limpa_Det_Desd
Call Rotina_027_limpa_totalizacao

tlanc.Open "Select * from TipoLancamento", db, 3, 3
If tlanc.EOF Then
   MsgBox ("Não encontrado registros em Tipo de Lançamento."), vbCritical
   Call FechaDB
   Exit Sub
End If

tlanc.MoveFirst
Do While Not tlanc.EOF
   cmbTipoLancamento.AddItem tlanc!chTipoDocumento
   tlanc.MoveNext
Loop

fpag.Open "Select * from FinalidadePagamento", db, 3, 3
If fpag.EOF Then
   MsgBox ("Tabela de finalidade de pagamento vazia. Carrega Nota Fiscal"), vbCritical
   Call FechaDB
   Exit Sub
End If

fpag.MoveFirst
Do While Not fpag.EOF
   cmbFinalidade.AddItem fpag!chfinalidadepagamento
   fpag.MoveNext
Loop
  
Option1 = False
OptNao = True

Call FechaDB
  
End Sub

Private Sub GridDesdobr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim DataInvertida As Double

Resp = MsgBox("Alteração da descrição da Fatura. Confirma???", vbYesNo)
If Resp = vbNo Then
   MsgBox "Rotina de Alteração de Fatura abortada."
   cmdSair.SetFocus
   Exit Sub
End If

Linha = GridDesdobr.Row
Coluna = GridDesdobr.Col

Call Rotina_AbrirBanco

DataInvertida = Year(GridDesdobr.TextMatrix(Ind, 1)) & Format$(Month(GridDesdobr.TextMatrix(Ind, 1)), "00") & Format$(Day(GridDesdobr.TextMatrix(Ind, 1)), "00")
       
ctp.Open "Select * from Contas_A_Pagar where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbPessoa & "') and chNotaFiscal = ('" & txtNotaFiscal & "') and chFatura = ('" & GridDesdobr.TextMatrix(Linha, 0) & "') and chDataVencito = ('" & DataInvertida & "')", db, 3, 3
If ctp.EOF Then
   txtFatura = GridDesdobr.TextMatrix(Linha, 0)

   txtValorFatura = GridDesdobr.TextMatrix(Linha, 2)
   Rotina_21_Limpa_Grid_Desdobr
   
   nfd.Open "Select * from NotaFiscalDesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chDataVencimento = ('" & DataInvertida & "')", db, 3, 3
   If Not (nfd.EOF) Then
      nfd.Delete
      Rotina_040_Carga_Grid_Desdobr
      Call FechaDB
      Exit Sub
   End If
End If

If ctp!ctpstatus = 1 Then
   MsgBox ("Não e permitida a alteração de Fatura para operação confirmada"), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If

DataHoje = ctp!ctpdatalanc
'DataBanco = TabCtaPagar("ctpDataBanco")
Data_Vencimento = ctp!chdatavencito

Resp = MsgBox("Alteração da Fatura - " & GridDesdobr.TextMatrix(Linha, 0) & ". Confirma???", vbYesNo)
If Resp = vbNo Then
   MsgBox ("Rotina de Alteração de Fatura abortada."), vbCritical
   Call FechaDB
   cmdSair.SetFocus
   Exit Sub
End If

txtFatura = GridDesdobr.TextMatrix(Linha, 0)
txtDataVencito = GridDesdobr.TextMatrix(Linha, 1)
txtValorFatura = GridDesdobr.TextMatrix(Linha, 2)
txtIPTE = GridDesdobr.TextMatrix(Linha, 3)
      
ctp.Delete

Rotina_21_Limpa_Grid_Desdobr
Rotina_040_Carga_Grid_Desdobr

Call FechaDB

End Sub


Private Sub GridProduto_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Coluna = GridProduto.Col
Linha = GridProduto.Row

If Linha > GridProduto.Rows Then
   MsgBox ("Clicar somente em Linha com conteúdo."), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If

If GridProduto.TextMatrix(Linha, 1) = Empty Then
   MsgBox ("Clicar somente em Linha com conteúdo."), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If

If lblStatus = "PROCESSADO" Then
   MsgBox ("Função Válida somente para pedidos pendentes"), vbInformation
   Exit Sub
End If

Call Rotina_AbrirBanco

dnfe.Open "Select * from NotaFiscalDetProd where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & GridProduto.TextMatrix(Linha, 0) & "')", db, 3, 3
If dnfe.EOF Then
   MsgBox ("Erro no acesso a Detalhe de Nota Fiscal na rotina de alteração"), vbCritical
   Call FechaDB
   Exit Sub
End If

'cmbPessoa = dnfe!chPessoa
'txtNotaFiscal = dnfe!chNotaFiscalEntrada
cmbCodProduto = dnfe!chCodProduto
lblProdutoFabrica = dnfe!chProdutoFabrica
txtQtd = dnfe!nfdQtd
txtPU = dnfe!nfdPU
txtValor = dnfe!nfdValorDaCompra
If GridProduto.TextMatrix(Linha, 3) = 1 Then
   lblUnidade = "Un"
Else
   lblUnidade = "M"
End If
'lblDescProdutoEntrada


Call FechaDB

End Sub

Private Sub txtDataVencito_lostfocus()

If Not (IsDate(txtDataVencito)) Then
   MsgBox ("Data de vencimento inválida")
   cmdSair.SetFocus
   Exit Sub
End If

If txtDataVencito = "__/__/____" Then
   Exit Sub
End If

Call Rotina_AbrirBanco

nfd.Open "Select * from NotaFiscalDesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chDataVencimento = ('" & txtDataVencito & "')", db, 3, 3
If nfd.EOF Then
   IncluiDesdob = 1
   cmdIncluiFatura.Enabled = True
   cmdAlteraFatura.Enabled = True
   'cmdExcluiFatura.Enabled = False
Else
   IncluiDesdob = 0
   cmdIncluiFatura.Enabled = False
   cmdAlteraFatura.Enabled = True
   'cmdExcluiFatura.Enabled = True
End If
txtValorFatura.SetFocus

Call FechaDB

End Sub


Private Sub txtNotaFiscal_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtNotaFiscal_LostFocus()
Verifica = Empty
Verifica = Mid$(txtNotaFiscal, 13, 5)
If Not Verifica = Empty Then
   MsgBox ("Nota Fiscal Informada ultrapassa 12 caracteres.")
   txtNotaFiscal.SetFocus
   Exit Sub
End If

Call Rotina_AbrirBanco

hnfe.Open "Select * from HistoricoNotaFiscalEntrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
If hnfe.EOF Then
   nfe.Open "Select * from NotaFiscalEntrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
   If nfe.EOF Then
      Incluir = 1
      cmdInclui.Enabled = True
      cmdAltera.Enabled = False
      cmdExclui.Enabled = False
   Else
      Incluir = 0
      cmdInclui.Enabled = False
      cmdAltera.Enabled = True
      cmdExclui.Enabled = True
      cmdGeraCtaPagar.Enabled = False
      Call Rotina_010_Carga_Form
   End If
Else
   MsgBox ("Esta Nota Fiscal já existe no historico para este fornecedor."), vbCritical
   Call FechaDB
   txtNotaFiscal.SetFocus
   Exit Sub
End If

Call FechaDB

End Sub





Private Sub txtPU_LostFocus()

If txtPU = "" Then
   MsgBox ("PU não Informado")
   txtValor = Format$(0, "##,##0.00")
   cmdSair.SetFocus
Else
   If txtQtd = "" Then
      txtQtd = Format$(0, "0.00")
      MsgBox ("Qtd não Informado")
      cmdSair.SetFocus
   Else
      If Not IsNumeric(txtPU) Then
         MsgBox ("PU Inválido")
         txtValor = Format$(0, "##,##0.00")
         cmdSair.SetFocus
         Exit Sub
      End If
      If Not IsNumeric(txtQtd) Then
         MsgBox ("PU Inválido")
         txtValor = Format$(0, "##,##0.00")
         cmdSair.SetFocus
         Exit Sub
      End If
      txtValor = Format$((txtQtd * txtPU), "##,##0.00")
   End If
End If
End Sub

Public Sub Rotina_010_Carga_Form()

cmbPessoa = nfe!chPessoa
cmbTipoLancamento.ListIndex = nfe!nfeTipoLancamento
cmbFinalidade.ListIndex = nfe!nfeFinalidadePagto
txtNotaFiscal = nfe!chNotaFiscalEntrada
cmbFabrica.ListIndex = nfe!nfelartmerco
txtDataEmissao = nfe!nfeDataEmissao
txtValorFrete = Format$(nfe!nfeValorFrete, "##,##0.00")
txtValorICMS = Format$(nfe!nfeValorICMS, "##,##0.00")
txtValorDaNotaFiscal = Format$(nfe!nfeValorDaNota, "0.00")
ValorAnterior = Format$(nfe!nfeValorDaNota, "##,##0.00")
txtQtdFaturas = nfe!nfeDesdobramento
txtValorIPI = Format$(nfe!nfeValorIPI, "##,##0.00")
cmbBanco.ListIndex = nfe!chCodBcoLart

If nfe!nfeNF_Boleto = 1 Then
   Option1.SetFocus
Else
   OptNao.SetFocus
End If

Call Rotina_020_Limpa_Grid_Produto
Call Rotina_21_Limpa_Grid_Desdobr
Call Rotina_030_Carga_Grid_Produto
Call Rotina_040_Carga_Grid_Desdobr
cmdSair.Enabled = True
cmdSair.SetFocus
End Sub

Public Sub Rotina_020_Limpa_Grid_Produto()
GridProduto.Rows = 2

GridProduto.TextMatrix(1, 0) = Empty
GridProduto.TextMatrix(1, 1) = Empty
GridProduto.TextMatrix(1, 2) = Empty
GridProduto.TextMatrix(1, 3) = Empty
GridProduto.TextMatrix(1, 4) = Empty
GridProduto.TextMatrix(1, 5) = Empty
GridProduto.TextMatrix(1, 6) = Empty

End Sub

Public Sub Rotina_21_Limpa_Grid_Desdobr()

GridDesdobr.Rows = 2

GridDesdobr.TextMatrix(1, 0) = Empty
GridDesdobr.TextMatrix(1, 1) = Empty
GridDesdobr.TextMatrix(1, 2) = Empty

End Sub
Public Sub Rotina_024_Limpa_Lancamento()

cmbPessoa = Empty
txtNotaFiscal = Empty
cmbFabrica.ListIndex = 0
cmbTipoLancamento.ListIndex = 0
'cmbFinalidade = Empty
txtValorFrete = Empty
txtValorICMS = Empty
txtValorIPI = Empty
txtValorDaNotaFiscal = Empty
txtQtdFaturas = 0
'OptSim.Enabled = True
'OptNao = Empty
cmbBanco.ListIndex = 0


End Sub
Public Sub Rotina_025_Limpa_Detalhe()

cmbCodProduto = Empty
lblDescProdutoEntrada = Empty
lblProdutoFabrica = Empty
lblUnidade = Empty
txtQtd = Empty 'Format$(0, "#,##0.00")
txtPU = Empty 'Format$(0, "#,##0.00")
txtValor = Empty 'Format$(0, "#,##0.00")
lblValorTotaldoProduto = Empty
lblQtdTotal = Empty

End Sub
Public Sub Rotina_026_Limpa_Det_Desd()
txtFatura = Empty
txtValorFatura = Empty 'Format$(0, "#,##0.00")
End Sub
Public Sub Rotina_027_limpa_totalizacao()
txtValorTotalFatura = Empty
txtValorTotalNota = Empty
txtDiferenca = Empty
End Sub
Public Sub Rotina_030_Carga_Grid_Produto()
Ind = 1
fim = 0
Acumula_Qtd = 0
Acumula_Valor = 0

Call Rotina_AbrirBanco

dnfe.Open "Select * from NotaFiscalDetProd where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
If dnfe.EOF Then
   MsgBox ("Nota Fiscal sem Lançamento de Produtos"), vbCritical
   Call FechaDB
   Exit Sub
End If

Do While fim = 0
   GridProduto.Rows = Ind + 1
   LimiteProduto = GridProduto.Rows
   GridProduto.TextMatrix(Ind, 0) = dnfe!chCodProduto
   
   If ProdEntrada.State = 1 Then
      ProdEntrada.Close: Set ProdEntrada = Nothing
   End If
   
   ProdEntrada.Open "Select * from ProdutoEntrada where chPessoa = ('" & dnfe!chPessoa & "') and chTipoProduto = ('" & dnfe!chCodProduto & "')", db, 3, 3
   If ProdEntrada.EOF Then
      MsgBox ("Erro no acesso a produto de entrada"), vbCritical
      Call FechaDB
      Exit Sub
   End If
   
   GridProduto.TextMatrix(Ind, 1) = ProdEntrada!pindescricao
   GridProduto.TextMatrix(Ind, 2) = ProdEntrada!pinunidade
   GridProduto.TextMatrix(Ind, 3) = dnfe!nfdQtd
   GridProduto.TextMatrix(Ind, 4) = Format$(dnfe!nfdPU, "###,##0.00")
   GridProduto.TextMatrix(Ind, 5) = Format$(dnfe!nfdQtd * dnfe!nfdPU, "###,##0.00")
   GridProduto.TextMatrix(Ind, 6) = ProdEntrada!chCodProduto
   GridProduto.TextMatrix(Ind, 7) = ProdEntrada!pinCntrlEstoque
   Acumula_Qtd = Acumula_Qtd + dnfe!nfdQtd
   Acumula_Valor = Acumula_Valor + Format$(dnfe!nfdQtd * dnfe!nfdPU, "###,##0.00")
   dnfe.MoveNext
   If dnfe.EOF Then
      fim = 1
   Else
      If dnfe("chpessoa") <> cmbPessoa Then
         fim = 1
      Else
         If dnfe!chNotaFiscalEntrada <> txtNotaFiscal Then
            fim = 1
         Else
            Ind = Ind + 1
         End If
      End If
   End If

Loop
lblQtdTotal = Acumula_Qtd
lblValorTotaldoProduto = Format$(Acumula_Valor, "0.00")
   

If Ind > 1 Then
   MaiorQueUm = 1
Else
   MaiorQueUm = 0
End If

Call FechaDB

End Sub

Public Sub Rotina_040_Carga_Grid_Desdobr()

Ind = 1
fim = 0
Acumula_Fatura = 0
LimiteCarga = 0

Call Rotina_AbrirBanco

nfd.Open "Select * from NotaFiscalDesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
If nfd.EOF Then
   MsgBox ("Nota Fiscal Sem Lançamento de Vencimento/desdobramentos"), vbCritical
   fim = 1
   Status = 0
Else
   Status = nfd!nfdStatus
   If Status = 0 And Incluir = 1 Then
      cmdGeraCtaPagar.Enabled = True
   Else
      cmdGeraCtaPagar.Enabled = False
   End If
   Do While fim = 0
      GridDesdobr.Rows = Ind + 1
      GridDesdobr.TextMatrix(Ind, 0) = nfd!nfdFaturaNumero
      GridDesdobr.TextMatrix(Ind, 1) = nfd!chDataVencimento
      GridDesdobr.TextMatrix(Ind, 2) = Format$(nfd!nfdValorDaFatura, "###,##0.00")
      GridDesdobr.TextMatrix(Ind, 3) = nfd!nfdIPTE
      Acumula_Fatura = Acumula_Fatura + nfd!nfdValorDaFatura
      nfd.MoveNext
      If nfd.EOF Then
         fim = 1
      Else
         If nfd!chPessoa <> cmbPessoa Then
            fim = 1
         Else
            If nfd!chNotaFiscalEntrada <> txtNotaFiscal Then
               fim = 1
            Else
               Ind = Ind + 1
            End If
         End If
      End If
   Loop
   LimiteCarga = GridDesdobr.Rows
End If
txtValorTotalFatura = Format$(Acumula_Fatura, "##,##0.00")
txtValorTotalNota = Format$(txtValorDaNotaFiscal, "##,##0.00")
txtDiferenca = Format$(txtValorTotalFatura - txtValorTotalNota, "##,##0.00")
If Status = 0 Then
   lblStatus.Caption = "Pendente"
Else
   lblStatus.Caption = "Processado"
End If

Call FechaDB

End Sub


Public Sub Rotina_050_Mover_Doc_DBNota()

nfe!chPessoa = cmbPessoa
nfe!nfeTipoLancamento = cmbTipoLancamento.ListIndex
nfe!chNotaFiscalEntrada = txtNotaFiscal
nfe!nfelartmerco = cmbFabrica.ListIndex
nfe!nfeDataEmissao = txtDataEmissao
nfe!nfeDataLanc = Date
nfe!nfeValorFrete = txtValorFrete
nfe!nfeFinalidadePagto = cmbFinalidade.ListIndex
nfe!nfeValorICMS = txtValorICMS
nfe!nfeValorIPI = txtValorIPI
nfe!nfeValorDaNota = txtValorDaNotaFiscal
ValorAnterior = txtValorDaNotaFiscal
nfe!nfeDesdobramento = txtQtdFaturas
If OptSim = True Then
   nfe!nfeNF_Boleto = 1
Else
   nfe!nfeNF_Boleto = 2
End If
nfe!chCodBcoLart = cmbBanco.ListIndex


End Sub

Public Sub Rotina_051_Doc_DBDET()

dnfe!chPessoa = cmbPessoa
dnfe!chNotaFiscalEntrada = txtNotaFiscal
dnfe!chCodProduto = cmbCodProduto
dnfe!chProdutoFabrica = lblProdutoFabrica
dnfe!nfdQtd = txtQtd
dnfe!nfdPU = txtPU
dnfe!nfdValorDaCompra = Format$(txtQtd * txtPU, "##,##0.00")
dnfe!nfdQtdParcelas = txtQtdFaturas
dnfe!nfdValorParcela = Format$((txtValor / txtQtdFaturas), "##,##0.00")

End Sub

Public Sub Rotina_052_Gera_Cta_Pagar()
ctp!chFabricante = cmbFabrica.ListIndex
ctp!chPessoa = cmbPessoa
ctp!chCodBcoLart = cmbBanco
ctp!chNotaFiscal = txtNotaFiscal
ctp!chFatura = Fatura
ctp!ctpdataemissao = txtDataEmissao

If DataHoje = Date Then
    ctp!ctpdatalanc = Date
    ctp!chdatavencito = Data_Vencimento
    ctp!ctpdatavencOriginal = Data_Vencimento
Else
    ctp!ctpdatalanc = DataHoje
    ctp!chdatavencito = Data_Vencimento
    ctp!ctpdatabanco = DataBanco
    ctp!ctpdatavencOriginal = Data_Vencimento
End If

If MaiorQueUm = 1 Then
   ctp!ctpdescricaooperacao = cmbFinalidade
Else
   ctp!ctpdescricaooperacao = GridProduto.TextMatrix(1, 0)
End If

If cmbFabrica.ListIndex = 0 Then
   ctp!ctpValorLart = GridDesdobr.TextMatrix(Ind, 2)
   ctp!ctpValorMerco = 0
Else
   ctp!ctpValorLart = 0
   ctp!ctpValorMerco = GridDesdobr.TextMatrix(Ind, 2)
End If
ctp!ctpvalordaboleta = GridDesdobr.TextMatrix(Ind, 2)
ctp!chAno = Year(Data_Hoje)
ctp!chMes = Month(Data_Hoje)
ctp!chDia = Day(Data_Hoje)

If cmbTipoLancamento = "REEMBOLSO" Then
   ctp!ctpstatus = 2
Else
   ctp!ctpstatus = 0
End If

If cmbPessoa = "LANC ESP" Then
   ctp!ctpTipoLancamento = 99
   ctp!ctpTipoLancamentoDesc = "LANC ESP"
Else
   ctp!ctpTipoLancamento = cmbTipoLancamento.ListIndex
   ctp!ctpTipoLancamentoDesc = cmbTipoLancamento
End If
End Sub

Public Sub Rotina_052_Gera_Cta_Receber()

Call Rotina_AbrirBanco

ctr.Open

ctr.AddNew
ctr!chFabricante = 0
ctr!chPessoa = neg!chPessoa
ctr!chNotaFiscal = neg!negNotaFiscal
ctr!chFatura = GridDesdobr.TextMatrix(Ind, 0)
ctr!ctrDataEmissao = Date
ctr!ctrDataVencito = txtDataVencito

'Calcula data banco
             
    DataUtil = txtDataVencito
      
    DataInformada = DataUtil
    NDias = 0
    'DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)
    'DataUtil = DataRetorno.DiaUtil

ctr!ctrDataBanco = DataUtil
ctr!ctrDataVencitoOriginal = txtDataVencito
ctr!ctrDescricaoOperacao = "ICMS FRETE"
ctr!ctrValorLart = GridDesdobr.TextMatrix(Ind, 2)
ctr!ctrValorMerco = 0
ctr!ctrPercentCorrecao = 0
ctr!ctrvalorcorrecao = 0
ctr!ctrPercentLogistica = 0
ctr!ctrValorLogistica = 0
ctr!ctrvalordaboleta = GridDesdobr.TextMatrix(Ind, 2)
ctr!chAno = Year(Data_Hoje)
ctr!chMes = Month(Data_Hoje)
ctr!chDia = Day(Data_Hoje)
ctr!chNumPedido = neg!chNumPedido
ctr!chNumPedidoComp = neg!chNumPedidoComp
ctr!chCodBcoLart = cmbBanco
ctr!ctrstatus = 0
ctr.Update

Call FechaDB

End Sub
'Public Sub Rotina_053_Estoque_Geral()
'Ind = 1

'Do While Ind < GridProduto.Rows
'   If GridProduto.TextMatrix(Ind, 7) = 1 Then
'        Funcao = 1
'        Fornecedor = cmbPessoa
'        Produto = GridProduto.TextMatrix(Ind, 6)
'        Ano = Year(Data_Hoje)
'        Mes = Month(Data_Hoje)
'        Entra = GridProduto.TextMatrix(Ind, 3)
'        Sai = 0
'        Call Atualiza_Estoque_Geral(Funcao, Fornecedor, Produto, Ano, Mes, Entra, Sai)
 '  End If
'   Ind = Ind + 1
'Loop
'End Sub

Public Sub Rotina_054_Deleta_Financ_Desdob()
Dim ContaRegFin As Byte
Dim ContaRegDesd As Byte
Dim DataInvertida As Double

ContaRegFin = 0
ContaRegDesd = 0
Ind = 1

Call Rotina_AbrirBanco



Do While Ind < LimiteCarga
   If ctp.State = 1 Then
      ctp.Close: Set ctp = Nothing
   End If
   If ctr.State = 1 Then
      ctr.Close: Set ctr = Nothing
   End If
   If nfd.State = 1 Then
      nfd.Close: Set nfd = Nothing
   End If
   
   ctp.Open "Select * from Contas_A_Pagar where chFabricante = ('" & cmbFabrica.ListIndex & "') and chPessoa = ('" & cmbPessoa & "') and chNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3
   If ctp.EOF Then
      MsgBox ("Não há financeiro lançado a débito"), vbInformation
   Else
      ctp.MoveFirst
      Do While Not ctp.EOF
         ctp.Delete
         ctp.MoveNext
         ContaRegFin = ContaRegFin + 1
      Loop
   End If
   
   ctr.Open "Select * from Contas_A_Receber where chFabricante = ('" & cmbFabrica.ListIndex & "') and chPessoa = ('" & cmbPessoa & "') and chNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3
   If Not ctr.EOF Then
      ctr.MoveFirst
      Do While Not ctr.EOF
         ctr.Delete
         ctr.MoveNext
      Loop
   End If
   
   If Not (GridDesdobr.TextMatrix(Ind, 1) = "") Then
   
      DataInvertida = Year(GridDesdobr.TextMatrix(Ind, 1)) & Format$(Month(GridDesdobr.TextMatrix(Ind, 1)), "00") & Format$(Day(GridDesdobr.TextMatrix(Ind, 1)), "00")
      
      nfd.Open "Select * from NotaFiscalDesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
      If nfd.EOF Then
         If ContaRegFin > 0 Then
            MsgBox ("Financeiro sem desdobramento"), vbInformation
            Ind = Ind + 1
         Else
            MsgBox ("Não há desdobramentos"), vbInformation
            Ind = Ind + 1
         End If
      Else
         nfd.MoveFirst
         Do While Not nfd.EOF
            nfd.Delete
            ContaRegDesd = ContaRegDesd + 1
            Ind = Ind + 1
            nfd.MoveNext
         Loop
      End If
   Else
      Ind = LimiteCarga + 1
   End If

Loop
   
'MsgBox ("Reg. Financeiro deletado = "), , ContaRegFin
'MsgBox ("Reg. Desdobramento deletado = "), , ContaRegDesd

Call FechaDB

End Sub

Public Sub Rotina_055_Deleta_Detalhe_NF()

Dim ContaRegPrd As Byte

ContaRegPrd = 0
Ind = 1
Call Rotina_AbrirBanco

Do While Ind < LimiteProduto

   If nfd.State = 1 Then
      nfd.Close: Set nfd = Nothing
   End If
   
   nfd.Open "Select * From NotaFiscalDetProd where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & GridProduto.TextMatrix(Ind, 0) & "')", db, 3, 3
   If nfd.EOF Then
      MsgBox ("Nota Fiscal sem lancamento de detalhes"), vbInformation
      Ind = Ind + 1
   Else
      nfd.Delete
      ContaRegPrd = ContaRegPrd + 1
      Ind = Ind + 1
   End If
Loop

Call FechaDB

'MsgBox ("Total de Detalhes deletados = "), , ContaRegPrd
End Sub

Public Sub Rotina_056_Deleta_Nota_Fiscal()

Call Rotina_AbrirBanco

nfe.Open "Select * from NotaFiscalEntrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
If nfe.EOF Then
   MsgBox ("Nota Fiscal Entrada não cadastrada"), vbCritical
   Call FechaDB
   Exit Sub
End If

nfe.Delete

Call FechaDB

End Sub



