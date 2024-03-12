VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNotaFiscalEntrada 
   BackColor       =   &H00E0E0E0&
   Caption         =   "frmNotaFiscalEntrada"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19275
   LinkTopic       =   "Form3"
   ScaleHeight     =   9660
   ScaleWidth      =   19275
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox txtDataHoje 
      Height          =   375
      Left            =   14880
      TabIndex        =   78
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
      TabIndex        =   67
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   19095
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
         Left            =   240
         TabIndex        =   56
         Top             =   5760
         Width           =   18735
         Begin VB.OptionButton optPendente 
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
            Height          =   375
            Left            =   5880
            TabIndex        =   30
            Top             =   480
            Width           =   975
         End
         Begin MSComCtl2.DTPicker dtDataPagamento 
            Height          =   375
            Left            =   8040
            TabIndex        =   32
            Top             =   480
            Width           =   1965
            _ExtentX        =   3466
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
            Format          =   242941953
            CurrentDate     =   44603
         End
         Begin VB.OptionButton optPaga 
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
            Height          =   375
            Left            =   6960
            TabIndex        =   31
            Top             =   480
            Width           =   975
         End
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
            Left            =   10080
            TabIndex        =   95
            Top             =   480
            Width           =   8415
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
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   2640
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker txtDataVencito 
            Height          =   405
            Left            =   2040
            TabIndex        =   28
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
            Format          =   242876417
            CurrentDate     =   43882
         End
         Begin MSFlexGridLib.MSFlexGrid GridDesdobr 
            Height          =   1575
            Left            =   960
            TabIndex        =   91
            Top             =   840
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   2778
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            BackColor       =   16777152
            BackColorFixed  =   16776960
            BackColorBkg    =   16777152
            FormatString    =   "N.Fatura          |Vencito           |Valor               ||Data Pagamento"
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
            Height          =   2535
            Left            =   16800
            TabIndex        =   79
            Top             =   840
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
               TabIndex        =   83
               Top             =   1920
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
               TabIndex        =   82
               Top             =   1440
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
               TabIndex        =   81
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
               TabIndex        =   80
               Top             =   480
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
            Height          =   615
            Left            =   13920
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   2760
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
            Height          =   1815
            Left            =   13920
            TabIndex        =   65
            Top             =   840
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
               TabIndex        =   35
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
               TabIndex        =   36
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
               TabIndex        =   37
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
            Height          =   2535
            Left            =   10080
            TabIndex        =   61
            Top             =   840
            Width           =   3735
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
               Left            =   1800
               TabIndex        =   71
               Top             =   1920
               Width           =   1815
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
               Left            =   1800
               TabIndex        =   70
               Top             =   1200
               Width           =   1815
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
               Left            =   1800
               TabIndex        =   69
               Top             =   480
               Width           =   1815
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
               TabIndex        =   64
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
               TabIndex        =   63
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
               TabIndex        =   62
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
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   39
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
            Height          =   975
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   2400
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
            TabIndex        =   29
            Top             =   480
            Width           =   1815
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
            TabIndex        =   27
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label lblDataPagamento 
            Caption         =   "Data Pagto."
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
            Left            =   8040
            TabIndex        =   97
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label28 
            Caption         =   "Fatura Paga?"
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
            Left            =   5880
            TabIndex        =   96
            Top             =   240
            Width           =   1815
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
            Left            =   10080
            TabIndex        =   94
            Top             =   240
            Width           =   2085
         End
         Begin VB.Label lblIPTE 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   300
            Left            =   6120
            TabIndex        =   93
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
            TabIndex        =   59
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
            TabIndex        =   58
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
            TabIndex        =   57
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
         TabIndex        =   45
         Top             =   1560
         Width           =   18975
         Begin MSFlexGridLib.MSFlexGrid GridProduto 
            Height          =   2055
            Left            =   120
            TabIndex        =   90
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            Left            =   16440
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1200
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
            Left            =   16440
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   2400
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
            Left            =   14640
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   2880
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
            Left            =   14640
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   2400
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
            Left            =   14640
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1920
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
            Left            =   14640
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1200
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
            TabIndex        =   20
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
            TabIndex        =   17
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
            TabIndex        =   89
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
            TabIndex        =   88
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
            TabIndex        =   87
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
            TabIndex        =   86
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
            TabIndex        =   85
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   18975
         Begin VB.ComboBox cmbLancamento 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   15360
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   360
            Width           =   3495
         End
         Begin VB.ComboBox cmbPessoaReembolso 
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
            ItemData        =   "Abertura_EncerramentoOld.frx":0000
            Left            =   13200
            List            =   "Abertura_EncerramentoOld.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   360
            Width           =   2175
         End
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
            TabIndex        =   8
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
            Format          =   337641473
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
            Left            =   6720
            TabIndex        =   12
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
            Left            =   5760
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
            Left            =   13920
            Style           =   2  'Dropdown List
            TabIndex        =   16
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
            Left            =   12720
            TabIndex        =   15
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
            Left            =   10200
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
            Left            =   9120
            TabIndex        =   75
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
               TabIndex        =   13
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
               TabIndex        =   14
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton OptSim 
               Caption         =   "Sim"
               Height          =   195
               Left            =   -600
               TabIndex        =   40
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
            ItemData        =   "Abertura_EncerramentoOld.frx":0004
            Left            =   7440
            List            =   "Abertura_EncerramentoOld.frx":0006
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
            Left            =   3600
            TabIndex        =   10
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
            Left            =   5160
            TabIndex        =   11
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
            Left            =   2040
            TabIndex        =   9
            Top             =   960
            Width           =   1455
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
            Left            =   2160
            TabIndex        =   2
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label29 
            Caption         =   "Tipo Lançamento"
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
            Left            =   15360
            TabIndex        =   98
            Top             =   120
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Fornec/Despesa"
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
            TabIndex        =   92
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
            Left            =   13920
            TabIndex        =   84
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
            Left            =   12360
            TabIndex        =   77
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
            Left            =   10200
            TabIndex        =   76
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
            Left            =   5640
            TabIndex        =   74
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
            Left            =   7440
            TabIndex        =   73
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
            Left            =   3600
            TabIndex        =   60
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
            Left            =   5160
            TabIndex        =   55
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
            Left            =   2040
            TabIndex        =   54
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
            Left            =   6720
            TabIndex        =   52
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
            Left            =   13200
            TabIndex        =   44
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
            TabIndex        =   43
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
            Left            =   2160
            TabIndex        =   42
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
      TabIndex        =   72
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
      TabIndex        =   68
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
      TabIndex        =   66
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   495
      Left            =   5400
      TabIndex        =   53
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

Dim TipoPessoa As Integer
Dim Acumula_Fatura As Currency
Dim ProcedAutomatico As Boolean
Dim TipoProcAutomatico As Integer
Dim ValorFixoAutomatico As Currency
Dim PercFixoAutomatico As Integer
Dim QtdVezes As Integer
Dim Ind As Byte
Dim Fim As Byte
Dim Incluir As Byte
Dim IncluiDesdob As Byte
Dim ContadorCtaPagar As Byte
Dim Data_Vencimento As Date
Dim DataAuxiliar As Date
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
Dim coluna As Integer
Dim Fatura As String
Dim GerarCredito As Byte
Dim IncluiNotaFiscal As Byte
Dim Interno As Byte
Dim ValorAnterior As Currency
Dim DespesaAnterior As String
Dim ProdutoAnterior As String
Dim Historico As Byte
Dim TabUnidadeEmbalagem(15) As String
Dim NumParcelas As Integer
Dim AcheiSupProduto As Integer

Private Sub cmbCodProduto_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub cmbDespFornec_LostFocus()


Call Rotina_AbrirBanco

cmbPessoa.Clear

pes.Open "Select * from pessoa", db, 3

If cmbDespFornec = "FORNECEDOR" Then
    pes.MoveFirst
    If pes.EOF Then
       MsgBox ("Dataset pessoa sem registro. Informar ao administrador do sistema"), vbCritical
       Call FechaDB
       Exit Sub
    End If
    
    pes.MoveFirst
    
    Do While Not pes.EOF
    
       If pes!pestipopessoa = 1 Or pes!pestipopessoa = 2 Or pes!pestipopessoa = 5 Or pes!pestipopessoa = 7 Then
          cmbPessoa.AddItem pes!chPessoa
       End If
       pes.MoveNext
    Loop
Else
    ProdFornec.Open "Select * from produtofornecedor", db, 3, 3
    If ProdFornec.EOF Then
       MsgBox ("Erro. Tabela de Produto fornecedor vazia. Comunicar ao analista responsável."), vbCritical
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
'cmbFabrica.ListIndex = 0
If cmbPessoa = "" Then
   cmdSair.Enabled = True
   cmdSair.SetFocus
   Exit Sub
End If

Call Rotina_AbrirBanco
If cmbPessoa = "" Then
   MsgBox ("Não Informado fornecedor ou Despesa."), vbCritical
   Call FechaDB
   cmbDespFornec.SetFocus
   Exit Sub
End If

Prod.Open "Select * from auxilio where nome_aux = ('" & cmbPessoa & "')", db, 3, 3
If Prod.EOF Then
   ProcedAutomatico = False
Else
   ProcedAutomatico = True
   TipoProcAutomatico = Prod!tipoAux
   If Prod!tipoAux = 1 Then
      ValorFixoAutomatico = Prod!Val
   Else
      PercFixoAutomatico = Prod!perc
   End If
End If

pes.Open "Select * from pessoa where chPessoa = ('" & cmbPessoa & "')", db, 3, 3
If pes.EOF Then
   If cmbDespFornec = "FORNECEDOR" Then
      MsgBox ("Efetuar o Cadastro deste fornecedor em pessoa e somente após o cadastramento lançar Nota fiscal"), vbCritical
      Call FechaDB
      cmdSair.SetFocus
      Exit Sub
   End If
Else
   If cmbDespFornec = "DESPESA" Then
      MsgBox ("Efetuar este lançamento como fornecedor."), vbCritical
      cmbPessoa.Clear
      cmbDespFornec.SetFocus
      Call FechaDB
      Exit Sub
   End If
End If

IncluiNotaFiscal = 0

nfe.Open "Select * from notafiscalentrada where chPessoa = ('" & cmbPessoa & "')", db, 3, 3
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

'cmbCodProduto = Empty
cmbCodProduto.Clear

If cmbDespFornec = "FORNECEDOR" Then
   ProdEntrada.Open "Select * from produtoentrada where chPessoa = ('" & cmbPessoa & "')", db, 3, 3
   If Not ProdEntrada.EOF Then
      ProdEntrada.MoveFirst
      Do While Not ProdEntrada.EOF
         cmbCodProduto.AddItem ProdEntrada!chTipoProduto
         ProdEntrada.MoveNext
       Loop
   End If
Else
   ProdFornec.Open "Select * from produtofornecedor where chTipoProduto = ('" & cmbPessoa & "')", db, 3, 3
   If Not ProdFornec.EOF Then
      ProdFornec.MoveFirst
      Do While Not ProdFornec.EOF
         If Not ProdFornec!chProdutoFabrica = ProdutoAnterior Then
            cmbCodProduto.AddItem ProdFornec!chProdutoFabrica
            ProdutoAnterior = ProdFornec!chProdutoFabrica
         End If
         ProdFornec.MoveNext
      Loop
   End If
End If
If Not cmbCodProduto.ListIndex < 0 Then
   cmbCodProduto.ListIndex = 0
End If
End Sub
Private Sub cmbCodProduto_LostFocus()
Dim SubGrupo As String

SubGrupo = "00"

Verifica = Empty
Verifica = Mid$(cmbCodProduto, 50, 5)
If Not Verifica = Empty Then
   MsgBox ("Código do Produto Informado ultrapassa 50 caracteres.")
   cmbCodProduto.SetFocus
   Exit Sub
End If

If cmbCodProduto = Empty Then
   If Incluir = 1 Then
      MsgBox ("Informar o código do produto ou cancelar a operação")
      cmdSair.SetFocus
   End If
End If

'TabFabFornec = produtoentrada
Call Rotina_AbrirBanco

If cmbDespFornec = "FORNECEDOR" Then
   gge.Open "Select * from supproduto where nomeprod = ('" & cmbCodProduto & "')", db, 3, 3
   If gge.EOF Then
      ProdEntrada.Open "Select * from produtoentrada where chPessoa = ('" & cmbPessoa & "') and chTipoProduto = ('" & cmbCodProduto & "')", db, 3, 3
      If ProdEntrada.EOF Then
         Resp = MsgBox("Produto não cadastrado. Deseja cadastra-lo???", vbYesNo)
         If Resp = vbYes Then
            frmProdutosDeEntrada.cmbOrigemProd = cmbDespFornec
            frmProdutosDeEntrada.cmbFornecedor = frmNotaFiscalEntrada.cmbPessoa
           ' frmProdutosDeEntrada.lblCodProduto = frmNotaFiscalEntrada.cmbCodProduto
            frmProdutosDeEntrada.cmbTipoProduto = frmNotaFiscalEntrada.cmbCodProduto
            frmProdutosDeEntrada.txtChaveEnvio = "NFE"
            frmProdutosDeEntrada.Show vbModal
            cmbCodProduto.SetFocus
         End If
      Else
            lblDescProdutoEntrada = ProdEntrada!pinDescricao
            lblProdutoFabrica = ProdEntrada!chProdutoFabrica
            lblUnidade = ProdEntrada!pinUnidade
      End If
   Else
      ccc.Open "SELECT DescricaoCentroDeCusto from centrodecusto WHERE chCentroDeCusto = ('" & 2 & "') and chGrupoCentroDeCusto = ('" & gge!GrupoCentroDeCusto & "') and chSubGrupoCentroDeCusto = ('" & SubGrupo & "')", db, 3, 3
      If Not ccc.EOF Then
         lblDescProdutoEntrada = gge!nomeProd
         lblProdutoFabrica = ccc!DescricaoCentroDeCusto
         'lblUnidade = Unidade DoEvents supproduto
      End If
   End If
Else
   ProdFornec.Open "Select * from produtofornecedor where chTipoProduto = ('" & cmbPessoa & "') and chProdutoFabrica = ('" & cmbCodProduto & "')", db, 3, 3
   If ProdFornec.EOF Then
      Resp = MsgBox("Despesa não cadastrado. Deseja cadastra-lo???", vbYesNo)
      If Resp = vbYes Then
         frmProdutosDeEntrada.cmbOrigemProd = cmbDespFornec
         frmProdutosDeEntrada.cmbFornecedor = frmNotaFiscalEntrada.cmbPessoa
        ' frmProdutosDeEntrada.lblCodProduto = frmNotaFiscalEntrada.cmbCodProduto
         frmProdutosDeEntrada.cmbTipoProduto = frmNotaFiscalEntrada.cmbCodProduto
         frmProdutosDeEntrada.txtChaveEnvio = "NFE"
         frmProdutosDeEntrada.Show vbModal
         cmbCodProduto.SetFocus
      End If
   Else
       lblDescProdutoEntrada = ProdFornec!chProdutoFabrica
       lblProdutoFabrica = ProdFornec!chCentroDeCusto
       'lblUnidade = ProdEntrada!pinUnidade
   End If
End If
   
dnfe.Open "Select * from notafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & cmbCodProduto & "')", db, 3, 3
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

Private Sub cmbPessoaReembolso_LostFocus()

If cmbTipoLancamento = "REEMBOLSO" Then
   If cmbPessoaReembolso.ListIndex = 0 Then
      MsgBox ("Para reembolso o sacado tem que ser diferente de SHB BRASIL"), vbInformation
      cmdSair.SetFocus
   End If
End If

End Sub



'Private Sub cmbTipoLancamento_LostFocus()
'If cmbTipoLancamento = "BOLETO" Then
'   lblIPTE.Visible = True
'   lblIPTE1.Visible = True
'   txtIPTE.Visible = True
'   txtIPTE = Empty
'Else
'   lblIPTE1.Visible = False
'   txtIPTE.Visible = False
'End If

'End Sub

Private Sub cmdAltera_Click()


Call Rotina_AbrirBanco

db.BeginTrans

   dnfe.Open "Select * from notafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & cmbCodProduto & "')", db, 3, 3
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

If GridDesdobr.TextMatrix(1, 1) = "" Then
   MsgBox ("Comando inválido. Você está tentando alterar um desdobramento que não existe."), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If


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
nfe.Open "Select * from notafiscalentrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
If nfe.EOF Then
   If nfe.State = 1 Then
      nfe.Close: Set dnfe = Nothing
   End If
   
   nfe.Open "Select * from historiconotafiscalentrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
   If nfe.EOF Then
      MsgBox ("Nota Fiscal não encontrada para Alteração de Fatura"), vbCritical
      db.CommitTrans
      Call FechaDB
      Exit Sub
   End If
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
   Resp = MsgBox("Valor do icms não informado. Deseja Informá-lo?", vbYesNo)
   If Resp = vbYes Then
      txtValorICMS.SetFocus
      Exit Sub
   Else
      txtValorICMS = Format$(0, "0.00")
   End If
Else
   If Not IsNumeric(txtValorICMS) Then
      MsgBox ("Valor do icms Inválido")
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
   MsgBox ("Informe o banco")
   cmbBanco.SetFocus
End If
If cmbTipoLancamento = Empty Then
   MsgBox ("Informe o Tipo do Documento")
   cmbTipoLancamento.SetFocus
End If
'If cmbFabrica = Empty Then
'   MsgBox ("Informe a empresa")
'   cmbFabrica.SetFocus
'End If
If cmbFinalidade = Empty Then
   MsgBox ("Informe a Finalidade da Despesa")
   cmbFinalidade.SetFocus
End If




'If Incluir = 1 Then

   Incluir = 0
   
   Call Rotina_AbrirBanco
   db.BeginTrans
   nfe.Open "Select * from notafiscalentrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
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
   
'   dnfe.Open "Select * from notafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & cmbCodProduto & "')", db, 3, 3
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
    
    dnfe.Open "Select * from notafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & GridProduto.TextMatrix(Linha, 0) & "')", db, 3, 3
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

If cmbPessoa = "" Then
   MsgBox ("Comando Inválido. Nao há cliente informado."), vbCritical
   Call FechaDB
   Exit Sub
End If

Do While Ind < GridDesdobr.Rows
       If cmbFinalidade = "X - icms FRETE" Then
          Data_Vencimento = Date
       Else
          Data_Vencimento = GridDesdobr.TextMatrix(Ind, 1)
       End If
       
       Call Rotina_AbrirBanco
       
       DataInvertida = Year(Data_Vencimento) & Format$(Month(Data_Vencimento), "00") & Format$(Day(Data_Vencimento), "00")
       
       
       If (Not (Year(Data_Vencimento) > Year(Date))) And Month(Data_Vencimento) < Month(Date) Then
          nfd.Open "Select * from historiconotafiscaldesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chDataVencimento = ('" & DataInvertida & "')", db, 3, 3
       Else
          nfd.Open "Select * from notafiscaldesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chDataVencimento = ('" & DataInvertida & "')", db, 3, 3
       End If
       
       If nfd.EOF Then
          MsgBox ("Erro no acesso a tabela de desdobramentos em Gerar Contas a Pagar"), vbCritical
          Call FechaDB
          Exit Sub
       Else
          
       If (Not (Year(Data_Vencimento) > Year(Date))) And Month(Data_Vencimento) < Month(Date) Then
             ctp.Open "Select * from historicocontaspagar where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbPessoa & "') and chNotaFiscal = ('" & txtNotaFiscal & "') and chFatura = ('" & GridDesdobr.TextMatrix(Ind, 0) & "')", db, 3, 3
          Else
             ctp.Open "Select * from contas_a_pagar where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbPessoa & "') and chNotaFiscal = ('" & txtNotaFiscal & "') and chFatura = ('" & GridDesdobr.TextMatrix(Ind, 0) & "')", db, 3, 3
          End If
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
                
                nfe.Open "Select * from historiconotafiscalentrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
                If nfe.EOF Then
                   If nfe.State = 1 Then
                       nfe.Close: Set dnfe = Nothing
                    End If
                   nfe.Open "Select * from notafiscalentrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
                   If nfe.EOF Then
                      MsgBox ("Nota Fiscal não encontrada em Gera Contas a Pagar."), vbCritical, txtNotaFiscal
                      Call FechaDB
                      Exit Sub
                   End If
                End If
                nfe!nfeDesdobramento = txtQtdFaturas
                nfe!nfeStatus = 1
                nfe.Update
                           
            db.CommitTrans
            
            If cmbTipoLancamento = "REEMBOLSO" Then
               Call RotinaGerarReembolso
            End If
                   
            MsgBox ("Contas a Pagar gerada com sucesso."), vbInformation
          Else
            MsgBox ("Contas a Pagar para esta NF/Fatura já foi gerada anteriormente.")
            txtFatura.SetFocus
          End If
          Ind = Ind + 1
       End If
Loop

Call RevisaDetProd

cmdGeraCtaPagar.Enabled = False
cmdNovaNotaFiscal.Enabled = True
cmdCancelar.Enabled = True
optPaga = False
optPendente = True
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
   Resp = MsgBox("Valor do icms não informado. Deseja Informá-lo?", vbYesNo)
   If Resp = vbYes Then
      txtValorICMS.SetFocus
      Exit Sub
   Else
      txtValorICMS = Format$(0, "0.00")
   End If
Else
   If Not IsNumeric(txtValorICMS) Then
      MsgBox ("Valor do icms Inválido")
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
   MsgBox ("Informe o banco")
   cmbBanco.SetFocus
End If
If cmbTipoLancamento = Empty Then
   MsgBox ("Informe o Tipo do Documento")
   cmbTipoLancamento.SetFocus
End If
'If cmbFabrica = Empty Then
'   MsgBox ("Informe a empresa")
'   cmbFabrica.SetFocus
'End If
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

   If cmbLancamento.ListIndex = 1 Then
      nfe.Open "Select * from historiconotafiscalentrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
   Else
      nfe.Open "Select * from notafiscalentrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
   End If
   
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
   If cmbLancamento.ListIndex = 1 Then
      dnfe.Open "Select * from historiconotafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & cmbCodProduto & "')", db, 3, 3
   Else
      dnfe.Open "Select * from notafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & cmbCodProduto & "')", db, 3, 3
   End If
   If dnfe.EOF Then
      dnfe.AddNew
   End If
   
   Call Rotina_051_Doc_DBDET
   
   dnfe.Update
   
   TipoPessoa = 8
   
   If ProcedAutomatico = True Then
      Resp = MsgBox("Você deseja efetuar lançamentos automáticos? Confirma???", vbYesNo)
      If Resp = vbYes Then
         ProdFornec.Open "Select * from produtofornecedor where chTipoProduto = ('" & cmbPessoa & "')", db, 3, 3
         If ProdFornec.EOF Then
            MsgBox ("Não tem funcionario para calculo automatico"), vbInformation
         Else
            ProdFornec.MoveFirst
            ProdFornec.MoveNext
            Do While Not ProdFornec.EOF
            
               If ccc.State = 1 Then
                  ccc.Close: Set ccc = Nothing
               End If
               
               ccc.Open "Select * from pessoa where chPessoa = ('" & ProdFornec!chProdutoFabrica & "')", db, 3, 3
               If Not ccc.EOF Then
                  If (TipoProcAutomatico = 1) Or (TipoProcAutomatico = 2 And ccc!salario > 0) Then
                     If ccc!pesStatusPessoa = 0 Then
                        cmbCodProduto = ProdFornec!chProdutoFabrica
                        txtQtd = 1
                        txtQtdFaturas = 1
                        If TipoProcAutomatico = 1 Then
                           txtValor = ValorFixoAutomatico
                           txtPU = ValorFixoAutomatico
                        Else
                           txtValor = ccc!salario * (PercFixoAutomatico / 100)
                           txtPU = ccc!salario * (PercFixoAutomatico / 100)
                        End If
                     
                        If dnfe.State = 1 Then
                           dnfe.Close: Set dnfe = Nothing
                        End If
                        
                        If cmbLancamento.ListIndex = 1 Then
                           dnfe.Open "Select * from historiconotafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & cmbCodProduto & "')", db, 3, 3
                        Else
                           dnfe.Open "Select * from notafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & cmbCodProduto & "')", db, 3, 3
                        End If
                        If dnfe.EOF Then
                           dnfe.AddNew
                        End If
                                 
                        Call Rotina_051_Doc_DBDET
                        
                        dnfe.Update
                     End If
                        
                  End If
               
               End If
               
               ProdFornec.MoveNext
               
               Loop
            End If
         End If
   End If
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

If txtQtdFaturas = Empty Or txtQtdFaturas = 0 Then
   MsgBox ("Quantidade de Faturas não Informado")
   txtQtdFaturas.SetFocus
   Exit Sub
End If

Call Rotina_AbrirBanco

db.BeginTrans
   If optPaga = True And Not (Month(dtDataPagamento) = Month(Date)) Then
      nfd.Open "select * from historiconotafiscaldesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and nfdFaturaNumero = ('" & txtFatura & "')", db, 3, 3
   Else
      nfd.Open "select * from notafiscaldesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and nfdFaturaNumero = ('" & txtFatura & "')", db, 3, 3
   End If
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
   If cmbFinalidade = "X - icms FRETE" Then
      nfd!chDataVencimento = Date
   Else
      nfd!chDataVencimento = txtDataVencito
   End If
   nfd!nfdDataVencOriginal = txtDataVencito
   nfd!nfdFaturaNumero = txtFatura
   nfd!nfdValorDaFatura = txtValorFatura
   nfd!nfdIPTE = txtIPTE
   If optPaga = True Then
      nfd!nfdDataPagamento = dtDataPagamento
      nfd!nfdStatus = 1
   End If
    
   nfd.Update

db.CommitTrans

Call Rotina_21_Limpa_Grid_Desdobr
Call Rotina_026_Limpa_Det_Desd
Call Rotina_040_Carga_Grid_Desdobr

If txtValorTotalFatura <> 0 And txtDiferenca = 0 Then
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
  
  nfe.Open "Select * from notafiscalentrada", db, 3, 3
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

cmdInclui.Enabled = True
cmdAltera.Enabled = True
cmdCancelar.Enabled = True
optPaga = False
optPendente = True

cmbPessoa.Enabled = True
cmbPessoa.SetFocus

End Sub

Private Sub cmdNovoDesdob_Click()
txtFatura = Empty
txtDataVencito = Date
txtValorFatura = 0
txtFatura.SetFocus
optPaga = False
dtDataPagamento.Visible = False
optPendente = True
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

Private Sub dtDataPagamento_LostFocus()
If optPaga = True Then
   If dtDataPagamento > Date - 1 Then
      MsgBox ("Contas já pagas somente permitida para datas anteriores a hoje. Ajustar data."), vbInformation
      'cmdSair.SetFocus
   End If
End If

If Month(dtDataPagamento) = Month(Date) And cmbLancamento.ListIndex = 2 Then
   MsgBox ("O tipo de lnaçamento espera pagamentos com datas anteriores ao mes atual."), vbInformation
   MsgBox ("Cancelar este lançamento para efetuar o lançamento corretamente"), vbInformation
   cmdCancelar.SetFocus
End If

End Sub

Private Sub Form_Load()

txtDataHoje = Date
txtDataVencito = Date
txtDataEmissao = Date
dtDataPagamento = Date

lblStatus.Caption = Empty

lblDataPagamento.Visible = False
dtDataPagamento.Visible = False

cmbLancamento.AddItem "NORMAL"
cmbLancamento.AddItem "EMIS/VENC.MES ANTER"
cmbLancamento.AddItem "EMIS ANTER/VENC MES ATUAL"
cmbLancamento.ListIndex = 0

cmbDespFornec.AddItem "FORNECEDOR"
cmbDespFornec.AddItem "DESPESA"
cmbDespFornec.ListIndex = 0
DespesaAnterior = Empty

optPendente = True
optPaga = False

Interno = 0

Call Rotina_AbrirBanco

Bco.Open "Select * from banco", db, 3, 3
If Bco.EOF Then
   MsgBox ("Tabela banco vazia"), vbCritical
   Call FechaDB
   Exit Sub
End If

Bco.MoveFirst

Do While Not Bco.EOF
   cmbBanco.AddItem Bco!bcosiglabco
   Bco.MoveNext
Loop

cmbBanco.ListIndex = 0
         
Emp.Open "Select * from empresa", db, 3, 3
If Emp.EOF Then
   MsgBox ("Empresa não encontrada em inicialização de Nota Fiscal"), vbCritical
   Call FechaDB
   Exit Sub
End If

'Emp.MoveLast
'Do While Not Emp.BOF
'   cmbFabrica.AddItem Emp!chPessoa
'   Emp.MovePrevious
'Loop

cmbPessoaReembolso.Clear

cmbPessoaReembolso.AddItem "SHB BRASIL"


pes.Open "Select * from pessoa where pesTipoPessoa > ('" & 5 & "') and pesTipoPessoa < ('" & 8 & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Não há colaborador cadastrado. Atenção"), vbInformation
Else
   pes.MoveFirst
   Do While Not pes.EOF
      cmbPessoaReembolso.AddItem pes!chPessoa
      pes.MoveNext
   Loop
End If

cmbPessoaReembolso.ListIndex = 0

'cmbFabrica.ListIndex = 0
Call Rotina_020_Limpa_Grid_Produto
Call Rotina_21_Limpa_Grid_Desdobr
Call Rotina_025_Limpa_Detalhe
Call Rotina_026_Limpa_Det_Desd
Call Rotina_027_limpa_totalizacao

tlanc.Open "Select * from tipolancamento", db, 3, 3
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

fpag.Open "Select * from finalidadepagamento", db, 3, 3
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

optNao = True

rs.Open "Select * from unidadeembalagem", db, 3, 3
If rs.EOF Then
   MsgBox ("Erro ao carregar Unidade de Embalagem"), vbInformation
   Call FechaDB
   Exit Sub
End If

rs.MoveFirst

Do While Not rs.EOF
   TabUnidadeEmbalagem(rs!indice) = rs!AbreviaturaUnidadeEmbalagem
   rs.MoveNext
Loop


Call FechaDB
  
End Sub


Private Sub GridDesdobr_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim DataInvertida As Double

Resp = MsgBox("Alteração da descrição da Fatura. Confirma???", vbYesNo)
If Resp = vbNo Then
   MsgBox "Rotina de Alteração de Fatura abortada."
   cmdSair.SetFocus
   Exit Sub
End If

Linha = GridDesdobr.Row
coluna = GridDesdobr.Col

Call Rotina_AbrirBanco

DataInvertida = Year(GridDesdobr.TextMatrix(Ind, 1)) & Format$(Month(GridDesdobr.TextMatrix(Ind, 1)), "00") & Format$(Day(GridDesdobr.TextMatrix(Ind, 1)), "00")
       
ctp.Open "Select * from contas_a_pagar where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbPessoa & "') and chNotaFiscal = ('" & txtNotaFiscal & "') and chFatura = ('" & GridDesdobr.TextMatrix(Linha, 0) & "') and chDataVencito = ('" & DataInvertida & "')", db, 3, 3
If ctp.EOF Then
   txtFatura = GridDesdobr.TextMatrix(Linha, 0)

   txtValorFatura = GridDesdobr.TextMatrix(Linha, 2)
   Rotina_21_Limpa_Grid_Desdobr
   
   nfd.Open "Select * from notafiscaldesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chDataVencimento = ('" & DataInvertida & "')", db, 3, 3
   If Not (nfd.EOF) Then
      nfd.Delete
      Rotina_040_Carga_Grid_Desdobr
      Call FechaDB
      Exit Sub
   End If
End If

If ctp!ctpStatus = 1 Then
   MsgBox ("Não e permitida a alteração de Fatura para operação confirmada"), vbInformation
   cmdSair.SetFocus
   Exit Sub
End If

DataHoje = ctp!ctpDataLanc
'DataBanco = TabCtaPagar("ctpDataBanco")
Data_Vencimento = ctp!chDataVencito

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


Private Sub GridProduto_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

coluna = GridProduto.Col
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

'If lblStatus = "PROCESSADO" Then
'   MsgBox ("Função Válida somente para pedidos pendentes"), vbInformation
'   Exit Sub
'End If

Call Rotina_AbrirBanco

dnfe.Open "Select * from notafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & GridProduto.TextMatrix(Linha, 0) & "')", db, 3, 3
If dnfe.EOF Then
   If dnfe.State = 1 Then
      dnfe.Close: Set dnfe = Nothing
   End If
   dnfe.Open "Select * from historiconotafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & GridProduto.TextMatrix(Linha, 0) & "')", db, 3, 3
   If dnfe.EOF Then
      MsgBox ("Erro no acesso a Detalhe de Nota Fiscal na rotina de alteração"), vbCritical
      Call FechaDB
      Exit Sub
   End If
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
cmbCodProduto.SetFocus

Call FechaDB

End Sub


Private Sub optPaga_LostFocus()
If optPaga = True Then
   lblDataPagamento.Visible = True
   dtDataPagamento.Visible = True
Else
   lblDataPagamento.Visible = False
   dtDataPagamento.Visible = False
End If
End Sub



Private Sub optPendente_LostFocus()
If optPendente = True Then
   dtDataPagamento.Visible = False
   optPaga = False
End If
End Sub



Private Sub txtDataEmissao_LostFocus()
If cmbLancamento.ListIndex = 0 Then
   If Not Month(Date) = Month(txtDataEmissao) Then
      MsgBox ("Para Tipo Lançamento NORMAL a data de emissão é igual a do mês atual."), vbCritical
      cmbLancamento.SetFocus
      Exit Sub
   End If
End If

If (cmbLancamento.ListIndex = 1) Or (cmbLancamento.ListIndex = 2) Then
   If Month(Date) = Month(txtDataEmissao) Then
      MsgBox ("Para Tipo Lançamento diferente de NORMAL a data de emissão é anterior ao mês atual."), vbCritical
      cmbLancamento.SetFocus
      Exit Sub
   End If
End If

End Sub

Private Sub txtDataVencito_lostfocus()

optPaga = False
dtDataPagamento.Visible = False
optPendente = True

If Not (IsDate(txtDataVencito)) Then
   MsgBox ("Data de vencimento inválida")
   cmdSair.SetFocus
   Exit Sub
End If

If txtDataVencito = "__/__/____" Then
   Exit Sub
End If

Call Rotina_AbrirBanco

If Not (Month(txtDataEmissao) = Month(Date)) Then
   nfd.Open "Select * from historiconotafiscaldesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chDataVencimento = ('" & txtDataVencito & "')", db, 3, 3
Else
   nfd.Open "Select * from notafiscaldesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chDataVencimento = ('" & txtDataVencito & "')", db, 3, 3
End If

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


Private Sub txtFatura_lostfocus()
If cmbTipoLancamento = "REEMBOLSO" Then
   optPaga.Enabled = False
Else
   optPaga.Enabled = True
End If
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

nfe.Open "Select * from historiconotafiscalentrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
If nfe.EOF Then
   If nfe.State = 1 Then
      nfe.Close: Set nfe = Nothing
   End If
   nfe.Open "Select * from notafiscalentrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
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
   Incluir = 0
   Status = 1
   cmdInclui.Enabled = False
   cmdAltera.Enabled = True
   cmdExclui.Enabled = True
   cmdGeraCtaPagar.Enabled = True
   Call Rotina_010_Carga_Form
   cmdSair.SetFocus
   cmdIncluiFatura.Enabled = False
   cmdGeraCtaPagar.Enabled = False
   cmdCancelar.Enabled = True
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
'cmbFabrica.ListIndex = nfe!nfelartmerco
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
   optNao.SetFocus
End If


Call Rotina_AbrirBanco

ctp.Open "Select * from contas_a_pagar where chPessoa = ('" & cmbPessoa & "') and chNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3
If ctp.EOF Then
   cmbPessoaReembolso.ListIndex = 0
Else
   If IsNull(ctp!ctpPessoaReembolso) Then
      cmbPessoaReembolso.ListIndex = 0
   Else
      cmbPessoaReembolso = ctp!ctpPessoaReembolso
   End If
End If

If Not cmbPessoaReembolso = "SHB BRASIL" Then
   pes.Open "Select * from pessoa where chPessoa = ('" & cmbPessoaReembolso & "')", db, 3, 3
   If pes.EOF Then
      MsgBox ("Erro no acesso a pessoa. Informar ao analista responsável"), vbCritical
      Call FechaDB
      Exit Sub
   Else
      cmbPessoaReembolso = pes!chPessoa
   End If
End If
Call Rotina_020_Limpa_Grid_Produto
Call Rotina_21_Limpa_Grid_Desdobr
Call Rotina_030_Carga_Grid_Produto
Call Rotina_040_Carga_Grid_Desdobr
cmdSair.Enabled = True
cmdSair.SetFocus

If Status = 1 Then
   lblStatus = "PROCESSADO"
Else
   lblStatus = "PENDENTE"
End If


Call FechaDB

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
GridDesdobr.TextMatrix(1, 3) = Empty
GridDesdobr.TextMatrix(1, 4) = Empty

End Sub
Public Sub Rotina_024_Limpa_Lancamento()

cmbPessoa = Empty
txtNotaFiscal = Empty
'cmbFabrica.ListIndex = 0
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
cmbPessoaReembolso.ListIndex = 0


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

Dim FlagSup As Integer
Ind = 1
Fim = 0

FlagSup = 0
Acumula_Qtd = 0
Acumula_Valor = 0


Call Rotina_AbrirBanco

dnfe.Open "Select * from notafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
If dnfe.EOF Then
   If dnfe.State = 1 Then
      dnfe.Close: Set dnfe = Nothing
   End If
   dnfe.Open "Select * from historiconotafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
   If dnfe.EOF Then
      MsgBox ("Nota Fiscal sem Lançamento de Produtos"), vbCritical
      Call FechaDB
      Exit Sub
   End If
End If

If Not cmbDespFornec = "DESPESA" Then
   Do While Fim = 0
      GridProduto.Rows = Ind + 1
      LimiteProduto = GridProduto.Rows
      GridProduto.TextMatrix(Ind, 0) = dnfe!chCodProduto
      
      If ProdEntrada.State = 1 Then
         ProdEntrada.Close: Set ProdEntrada = Nothing
      End If
      
      If rs.State = 1 Then
         rs.Close: Set rs = Nothing
      End If
      
      ProdEntrada.Open "Select * from produtoentrada where chPessoa = ('" & dnfe!chPessoa & "') and chTipoProduto = ('" & dnfe!chCodProduto & "')", db, 3, 3
      If ProdEntrada.EOF Then
         rs.Open "Select * from servservico where descricao = ('" & dnfe!chCodProduto & "')", db, 3, 3
         If rs.EOF Then
            rs.Close
            rs.Open "Select * from supproduto where nomeProd = ('" & dnfe!chCodProduto & "')", db, 3, 3
            If rs.EOF Then
               MsgBox ("Erro no acesso a produto de entrada"), vbCritical
               Call FechaDB
               Exit Sub
            Else
               FlagSup = 1
            End If
         Else
            FlagSup = 2
         End If
      Else
         FlagSup = 0
      End If
       
      If unid.State = 1 Then
         unid.Close: Set unid = Nothing
      End If
      
      If FlagSup = 0 Then
         unid.Open "SELECT UnidadeMedida from unidadedemedida where chUnidadeDeMedida = ('" & ProdEntrada!pinUnidade & "')", db, 3, 3
         If unid.EOF Then
            MsgBox ("Unidade de medida não encontrada."), vbInformation
            Call FechaDB
            Exit Sub
         End If
      Else
         If FlagSup = 1 Then
            unid.Open "SELECT unidadeembalagem from unidadeembalagem where indice = ('" & rs!unidadeProd & "')", db, 3, 3
            If unid.EOF Then
               MsgBox ("Unidade de medida não encontrada."), vbInformation
               Call FechaDB
               Exit Sub
            End If
         Else
            If FlagSup = 2 Then
               unid.Open "SELECT abrevUnidServ from unidadedeservicos where indice = ('" & rs!unidade & "')", db, 3, 3
               If unid.EOF Then
                  MsgBox ("Unidade de medida não encontrada."), vbInformation
                  Call FechaDB
                  Exit Sub
               End If
            End If
         End If
      End If

      If FlagSup = 1 Then
         GridProduto.TextMatrix(Ind, 1) = rs!nomeProd
         'GridProduto.TextMatrix(Ind, 2) = Unid!UnidadeEmbalagem
         GridProduto.TextMatrix(Ind, 2) = TabUnidadeEmbalagem(rs!unidadeProd)
      Else
         If FlagSup = 2 Then
            GridProduto.TextMatrix(Ind, 1) = rs!Descricao
            GridProduto.TextMatrix(Ind, 2) = unid!abrevunidserv
            'GridProduto.TextMatrix(Ind, 2) = TabUnidadeEmbalagem(rs!unidadeProd)
         Else
            GridProduto.TextMatrix(Ind, 1) = ProdEntrada!pinDescricao
            GridProduto.TextMatrix(Ind, 2) = ProdEntrada!pinUnidade
         End If
      End If
      GridProduto.TextMatrix(Ind, 3) = dnfe!nfdQtd
      GridProduto.TextMatrix(Ind, 4) = Format$(dnfe!nfdPU, "###,##0.00")
      GridProduto.TextMatrix(Ind, 5) = Format$(dnfe!nfdQtd * dnfe!nfdPU, "###,##0.00")
      'GridProduto.TextMatrix(Ind, 6) = ProdEntrada!chCodProduto
      GridProduto.TextMatrix(Ind, 6) = dnfe!chPessoa
      'GridProduto.TextMatrix(Ind, 7) = ProdEntrada!pinCntrlEstoque
      GridProduto.TextMatrix(Ind, 7) = 0
      Acumula_Qtd = Acumula_Qtd + dnfe!nfdQtd
      Acumula_Valor = Acumula_Valor + Format$(dnfe!nfdQtd * dnfe!nfdPU, "###,##0.00")
      dnfe.MoveNext
      If dnfe.EOF Then
         Fim = 1
      Else
         If dnfe("chpessoa") <> cmbPessoa Then
            Fim = 1
         Else
            If dnfe!chNotaFiscalEntrada <> txtNotaFiscal Then
               Fim = 1
            Else
               Ind = Ind + 1
            End If
         End If
     End If
   Loop
Else
   Do While Fim = 0
      GridProduto.Rows = Ind + 1
      LimiteProduto = GridProduto.Rows
      GridProduto.TextMatrix(Ind, 0) = dnfe!chCodProduto
      
      If ProdFornec.State = 1 Then
         ProdFornec.Close: Set ProdFornec = Nothing
      End If
      
      ProdFornec.Open "Select * from produtofornecedor where chTipoProduto = ('" & dnfe!chPessoa & "') and chProdutoFabrica = ('" & dnfe!chCodProduto & "')", db, 3, 3
      If ProdFornec.EOF Then
         MsgBox ("Erro no acesso a produto de entrada"), vbCritical
         Call FechaDB
         Exit Sub
      End If
      
      GridProduto.TextMatrix(Ind, 1) = ProdFornec!chProdutoFabrica
      GridProduto.TextMatrix(Ind, 2) = "Unidade"
      GridProduto.TextMatrix(Ind, 3) = dnfe!nfdQtd
      GridProduto.TextMatrix(Ind, 4) = Format$(dnfe!nfdPU, "###,##0.00")
      GridProduto.TextMatrix(Ind, 5) = Format$(dnfe!nfdQtd * dnfe!nfdPU, "###,##0.00")
      GridProduto.TextMatrix(Ind, 6) = ProdFornec!chCentroDeCusto
      GridProduto.TextMatrix(Ind, 7) = 0
      Acumula_Qtd = Acumula_Qtd + dnfe!nfdQtd
      Acumula_Valor = Acumula_Valor + Format$(dnfe!nfdQtd * dnfe!nfdPU, "###,##0.00")
      dnfe.MoveNext
      If dnfe.EOF Then
         Fim = 1
      Else
         If dnfe("chpessoa") <> cmbPessoa Then
            Fim = 1
         Else
            If dnfe!chNotaFiscalEntrada <> txtNotaFiscal Then
               Fim = 1
            Else
               Ind = Ind + 1
            End If
         End If
      End If

   Loop

End If

lblQtdTotal = Acumula_Qtd
lblValorTotaldoProduto = Format$(Acumula_Valor, "0.00")
   

If Ind > 1 Then
   MaiorQueUm = 1
Else
   MaiorQueUm = 0
End If


nfe.Open "SELECT * from notafiscalentrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
If nfe.EOF Then
   MsgBox ("Nota Fiscal não encontrada na carga."), vbInformation
   Call FechaDB
   Exit Sub
End If

If nfe!nfeStatus = 0 Then
   cmdGeraCtaPagar.Enabled = True
End If
'If txtDiferenca = 0 And lblStatus = "Pendente" Then
'   cmdGeraCtaPagar.Enabled = True
'End If


Call FechaDB



End Sub

Public Sub Rotina_040_Carga_Grid_Desdobr()

Ind = 0

Fim = 0
Acumula_Fatura = 0
LimiteCarga = 0

Call Rotina_AbrirBanco

nfd.Open "Select * from historiconotafiscaldesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3

If nfd.EOF Then
   Fim = 1
   Status = 0
Else
   nfd.MoveFirst
   Status = 1
End If

Do While Not nfd.EOF
   
   Ind = Ind + 1
   GridDesdobr.Rows = Ind + 1
   GridDesdobr.TextMatrix(Ind, 0) = nfd!nfdFaturaNumero
   GridDesdobr.TextMatrix(Ind, 1) = nfd!chDataVencimento
   GridDesdobr.TextMatrix(Ind, 2) = Format$(nfd!nfdValorDaFatura, "###,##0.00")
   GridDesdobr.TextMatrix(Ind, 3) = nfd!nfdIPTE
   If Not IsNull(nfd!nfdDataPagamento) Then
      GridDesdobr.TextMatrix(Ind, 4) = nfd!nfdDataPagamento
   Else
      GridDesdobr.TextMatrix(Ind, 4) = "NORMAL"
   End If
   txtIPTE = nfd!nfdIPTE
         
   Acumula_Fatura = Acumula_Fatura + nfd!nfdValorDaFatura
   nfd.MoveNext
   
Loop

LimiteCarga = GridDesdobr.Rows
   
Fim = 0

If nfd.State = 1 Then
   nfd.Close: Set nfd = Nothing
End If


nfd.Open "Select * from notafiscaldesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3

If nfd.EOF Then
   Fim = 1
Else
   Status = nfd!nfdStatus
   
   If Status = 0 And Incluir = 1 Then
      cmdGeraCtaPagar.Enabled = True
   Else
      cmdGeraCtaPagar.Enabled = False
   End If

   Do While Fim = 0
      Ind = Ind + 1
      GridDesdobr.Rows = Ind + 1
      GridDesdobr.TextMatrix(Ind, 0) = nfd!nfdFaturaNumero
      GridDesdobr.TextMatrix(Ind, 1) = nfd!chDataVencimento
      GridDesdobr.TextMatrix(Ind, 2) = Format$(nfd!nfdValorDaFatura, "###,##0.00")
      If Not IsNull(nfd!nfdIPTE) Then
         GridDesdobr.TextMatrix(Ind, 3) = nfd!nfdIPTE
      End If
      If Not IsNull(nfd!nfdDataPagamento) Then
         GridDesdobr.TextMatrix(Ind, 4) = nfd!nfdDataPagamento
      Else
         GridDesdobr.TextMatrix(Ind, 4) = "NORMAL"
      End If
      
      If Not IsNull(nfd!nfdIPTE) Then
         txtIPTE = nfd!nfdIPTE
      End If
            
      Acumula_Fatura = Acumula_Fatura + nfd!nfdValorDaFatura
      nfd.MoveNext
      If nfd.EOF Then
         Fim = 1
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

If lblStatus = "Pendente" And txtDiferenca = 0 Then
   cmdGeraCtaPagar.Enabled = True
End If

Call FechaDB

End Sub


Public Sub Rotina_050_Mover_Doc_DBNota()

nfe!chPessoa = cmbPessoa
nfe!nfeTipoLancamento = cmbTipoLancamento.ListIndex
nfe!chNotaFiscalEntrada = txtNotaFiscal
'nfe!nfelartmerco = cmbFabrica.ListIndex
nfe!nfeDataEmissao = txtDataEmissao
nfe!nfedataLanc = Date
nfe!nfeValorFrete = txtValorFrete
nfe!nfeFinalidadePagto = cmbFinalidade.ListIndex
nfe!nfeValorICMS = txtValorICMS
nfe!nfeValorIPI = txtValorIPI
nfe!nfeValorDaNota = txtValorDaNotaFiscal
ValorAnterior = txtValorDaNotaFiscal
nfe!nfeDesdobramento = txtQtdFaturas
If optSim = True Then
   nfe!nfeNF_Boleto = 1
Else
   nfe!nfeNF_Boleto = 2
End If
If Not (Month(txtDataEmissao) = Month(Date)) Then
   nfe!nfeStatus = 1
End If
nfe!chCodBcoLart = cmbBanco.ListIndex

End Sub

Public Sub Rotina_051_Doc_DBDET()
Dim pula As Integer

dnfe!chPessoa = cmbPessoa
dnfe!chNotaFiscalEntrada = txtNotaFiscal
dnfe!chCodProduto = cmbCodProduto
dnfe!chProdutoFabrica = lblProdutoFabrica
If txtQtd = "" Then
   txtQtd = 1
End If
dnfe!nfdQtd = txtQtd
If txtPU = "" Then
   txtPU = 0
End If
If txtFatura = "" Then
   txtFatura = 1
End If
If txtValor = "" Then
   txtValor = 1
End If
dnfe!nfdPU = txtPU
dnfe!nfdValorDaCompra = Format$(txtQtd * txtPU, "##,##0.00")
dnfe!nfdQtdParcelas = txtQtdFaturas
dnfe!nfdValorParcela = Format$((txtValor / txtQtdFaturas), "##,##0.00")

If Prod.State = 1 Then
   Prod.Close: Set Prod = Nothing
End If
pula = 0
Prod.Open "Select * from produtoentrada where chPessoa = ('" & cmbPessoa & "') and chTipoProduto = ('" & cmbCodProduto & "')", db, 3, 3
If Prod.EOF Then
   If Prod.State = 1 Then
      Prod.Close: Set Prod = Nothing
   End If
   Prod.Open "Select * from produtofornecedor where chTipoProduto = ('" & cmbPessoa & "') and chProdutoFabrica = ('" & cmbCodProduto & "')", db, 3, 3
   If Prod.EOF Then
      gge.Open "Select * from supproduto where nomeprod = ('" & cmbCodProduto & "')", db, 3, 3
      If gge.EOF Then
         MsgBox ("Carga de Produto inválida. Comunicar ao analisra responsável"), vbCritical
         Call FechaDB
         Exit Sub
      Else
         pula = 1
      End If
   End If
End If

If Not pula = 1 Then
   dnfe!nfdCentroDeCusto = Prod!pinCentroDeCusto
   dnfe!nfdGrupoCentroDeCusto = Prod!pinGrupoCentroDeCusto
   dnfe!nfdSubGrupoCentroDeCusto = Prod!pinSubGrupoCentroDeCusto
Else
   pula = 0
End If

End Sub

Public Sub Rotina_052_Gera_Cta_Pagar()
'ctp!chFabricante = cmbFabrica.ListIndex
ctp!ctpPessoaReembolso = cmbPessoaReembolso
ctp!chPessoa = cmbPessoa
ctp!chCodBcoLart = cmbBanco
ctp!chNotafiscal = txtNotaFiscal
ctp!chFatura = Fatura
ctp!ctpDataEmissao = txtDataEmissao

If DataHoje = Date Then
    ctp!ctpDataLanc = Date
    ctp!chDataVencito = Data_Vencimento
    ctp!ctpDataVencOriginal = Data_Vencimento
Else
    ctp!ctpDataLanc = DataHoje
    ctp!chDataVencito = Data_Vencimento
    ctp!ctpdatabanco = DataBanco
    ctp!ctpDataVencOriginal = Data_Vencimento
End If

If MaiorQueUm = 1 Then
   ctp!ctpdescricaooperacao = cmbFinalidade
Else
   ctp!ctpdescricaooperacao = GridProduto.TextMatrix(1, 0)
End If

ctp!ctpValorDaBoleta = GridDesdobr.TextMatrix(Ind, 2)
ctp!chAno = Year(Data_Hoje)
ctp!chMes = Month(Data_Hoje)
ctp!chDia = Day(Data_Hoje)

If cmbTipoLancamento = "REEMBOLSO" Then
   ctp!ctpStatus = 2
Else
   ctp!ctpStatus = 0
End If

If Not (GridDesdobr.TextMatrix(Ind, 4)) = "NORMAL" Then
   ctp!ctpStatus = 1
   ctp!ctpDataPagamento = dtDataPagamento
End If

If cmbPessoa = "LANC ESP" Then
   ctp!ctpTipoLancamento = 99
   ctp!ctpTipoLancamentoDesc = "LANC ESP"
Else
   ctp!ctpTipoLancamento = cmbTipoLancamento.ListIndex
   ctp!ctpTipoLancamentoDesc = cmbTipoLancamento
End If

ctp!ctpDataProc = Date
   
End Sub

Public Sub Rotina_052_Gera_Cta_Receber()

Call Rotina_AbrirBanco

ctr.Open

ctr.AddNew
ctr!chFabricante = 0
ctr!chPessoa = neg!chPessoa
ctr!chNotafiscal = neg!negNotaFiscal
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
ctr!ctrPercentlogistica = 0
ctr!ctrValorlogistica = 0
ctr!ctrValorDaBoleta = GridDesdobr.TextMatrix(Ind, 2)
ctr!chAno = Year(Data_Hoje)
ctr!chMes = Month(Data_Hoje)
ctr!chDia = Day(Data_Hoje)
ctr!chNumPedido = neg!chNumPedido
ctr!chNumPedidoComp = neg!chNumPedidoComp
ctr!chCodBcoLart = cmbBanco
ctr!ctrStatus = 0
ctr.Update

Call FechaDB

End Sub
'Public Sub Rotina_053_Estoque_Geral()
'Ind = 1

'Do While Ind < GridProduto.Rows
'   If GridProduto.TextMatrix(Ind, 7) = 1 Then
'        Funcao = 1
'        fornecedor = cmbPessoa
'        Produto = GridProduto.TextMatrix(Ind, 6)
'        Ano = Year(Data_Hoje)
'        Mes = Month(Data_Hoje)
'        Entra = GridProduto.TextMatrix(Ind, 3)
'        Sai = 0
'        Call Atualiza_Estoque_Geral(Funcao, fornecedor, Produto, Ano, Mes, Entra, Sai)
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
   
   ctp.Open "Select * from contas_a_pagar where chPessoa = ('" & cmbPessoa & "') and chNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3
   
   If Not ctp.EOF Then
      ctp.MoveFirst
      Do While Not ctp.EOF
         ctp.Delete
         ctp.MoveNext
         ContaRegFin = ContaRegFin + 1
      Loop
   End If
   
   If ctp.State = 1 Then
      ctp.Close: Set ctp = Nothing
   End If
   
   
   ctp.Open "Select * from historicocontaspagar where chPessoa = ('" & cmbPessoa & "') and chNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3
   If Not ctp.EOF Then
      ctp.MoveFirst
      Do While Not ctp.EOF
         ctp.Delete
         ctp.MoveNext
         ContaRegFin = ContaRegFin + 1
      Loop
   End If
   
   ctr.Open "Select * from contas_a_receber where chPessoa = ('" & cmbPessoa & "') and chNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3
   If Not ctr.EOF Then
      ctr.MoveFirst
      Do While Not ctr.EOF
         ctr.Delete
         ctr.MoveNext
      Loop
   End If
   
   If Not (GridDesdobr.TextMatrix(Ind, 1) = "") Then
   
      DataInvertida = Year(GridDesdobr.TextMatrix(Ind, 1)) & Format$(Month(GridDesdobr.TextMatrix(Ind, 1)), "00") & Format$(Day(GridDesdobr.TextMatrix(Ind, 1)), "00")
      
      If Not GridDesdobr.TextMatrix(Ind, 4) = "NORMAL" Then
         DataAuxiliar = GridDesdobr.TextMatrix(Ind, 4)
      Else
         DataAuxiliar = GridDesdobr.TextMatrix(Ind, 1)
      End If
      
      nfd.Open "Select * from notafiscaldesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
      If nfd.EOF Then
         Ind = Ind + 1
      Else
         nfd.MoveFirst
         Do While Not nfd.EOF
            nfd.Delete
            ContaRegDesd = ContaRegDesd + 1
            Ind = Ind + 1
            nfd.MoveNext
         Loop
      End If
      
      If nfd.State = 1 Then
         nfd.Close: Set nfd = Nothing
      End If
      
      nfd.Open "Select * from historiconotafiscaldesdobramento where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
      If nfd.EOF Then
         Ind = Ind + 1
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
   
   nfd.Open "Select * From notafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & GridProduto.TextMatrix(Ind, 0) & "')", db, 3, 3
   If nfd.EOF Then
      If nfd.State = 1 Then
         nfd.Close: Set nfd = Nothing
      End If
      nfd.Open "Select * From historiconotafiscaldetprod where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & GridProduto.TextMatrix(Ind, 0) & "')", db, 3, 3
      If nfd.EOF Then
         MsgBox ("Nota Fiscal sem lancamento de detalhes"), vbInformation
         Ind = Ind + 1
         MsgBox ("Nota Fiscal sem lancamento de detalhes"), vbInformation
         Ind = Ind + 1
      Else
         nfd.Delete
         ContaRegPrd = ContaRegPrd + 1
         Ind = Ind + 1
      End If
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

nfe.Open "Select * from notafiscalentrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
If nfe.EOF Then
   If nfe.State = 1 Then
      nfe.Close: Set nfe = Nothing
   End If
   nfe.Open "Select * from historiconotafiscalentrada where chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
   If nfe.EOF Then
      MsgBox ("Nota Fiscal Entrada não cadastrada"), vbCritical
      Call FechaDB
      Exit Sub
   End If
End If

nfe.Delete

If cmbTipoLancamento = "REEMBOLSO" Then
   Rmb.Open "Select * from reembolso where rmbColaborador = ('" & cmbPessoaReembolso & "') and chNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3
   If Not Rmb.EOF Then
      Rmb.Delete
   End If
End If

Call FechaDB

End Sub



Private Sub txtValor_LostFocus()

If txtValor > 0 Then
   If cmbTipoLancamento = "REEMBOLSO" Then
      If cmbPessoaReembolso.ListIndex = 0 Then
         MsgBox ("Para reembolso o sacado tem que ser diferente de SHB BRASIL"), vbInformation
         cmbPessoaReembolso.SetFocus
      End If
   End If
End If
End Sub

Public Sub RotinaGerarReembolso()

Call FechaDB

Call Rotina_AbrirBanco
      
      Rmb.Open "Select * from reembolso where rmbColaborador = ('" & cmbPessoaReembolso & "') and chNotaFiscal = ('" & txtNotaFiscal & "')", db, 3, 3
      
      If Rmb.EOF Then
         Rmb.AddNew
      End If
 
      db.BeginTrans
      
      pes.Open "Select * from pessoa Where chPessoa = ('" & cmbPessoaReembolso & "')", db, 3, 3
      If pes.EOF Then
         MsgBox ("Erro no acesso a pessoa. Comunicar ao analista responsável") & cmbPessoaReembolso, vbCritical
         Call FechaDB
         Exit Sub
      End If
      
      Rmb!chPessoa = cmbPessoa
      Rmb!chNotafiscal = txtNotaFiscal
      Rmb!rmbFatura = Empty
      Rmb!RmbColaborador = cmbPessoaReembolso
      Rmb!RmbNomeColaborador = pes!pesRazaoSocial
      Rmb!rmbBanco = pes!pesBanco
      Rmb!rmbAgencia = pes!pesAgencia
      Rmb!rmbContaCorrente = pes!pesConta
      Rmb!rmbCNPJ_CPF = pes!chCNPJ_CPF
      
      Rmb!RmbDataLancReembolso = Date
      Rmb!rmbDataNotaFiscal = txtDataEmissao
      
      Rmb!rmbDataReembolso = Date
      Rmb!rmbTiporeembolso = Empty
      Rmb!rmbTipoReembolsoTexto = Empty
      Rmb!rmbMeioPagto = Empty
      Rmb!rmbMeioPagtoTexto = Empty
      Rmb!rmbValorReembolso = 0
      Rmb!rmbNumComprovanteReembolso = Empty
      
      nfe.Open "Select * from notafiscalentrada where chPessoa = ('" & cmbPessoa & "') and chnotafiscalentrada = ('" & txtNotaFiscal & "')", db, 3, 3
      If nfe.EOF Then
         If nfe.State = 1 Then
            nfe.Close: Set nfe = Nothing
         End If
         nfe.Open "Select * from historiconotafiscalentrada where chPessoa = ('" & cmbPessoa & "') and chnotafiscalentrada = ('" & txtNotaFiscal & "')", db, 3, 3
         If nfe.EOF Then
            MsgBox ("Nota Fiscal não encontrada. Data de hoje como data de emissão"), vbInformation
            Rmb!rmbDataNotaFiscal = Date
         Else
            Rmb!rmbDataNotaFiscal = nfe!nfeDataEmissao
         End If
      Else
         Rmb!rmbDataNotaFiscal = nfe!nfeDataEmissao
      End If
            
      Rmb!rmbStatusReembolso = 0
      Rmb!rmbStatusRecibo = 0
      
      Rmb.Update
      
      MsgBox ("Nota Fiscal com reembolso gerada com sucesso."), vbInformation

db.CommitTrans

Rmb.Close
nfe.Close
pes.Close

'Call FechaDB

End Sub

Public Sub RevisaDetProd()


NumParcelas = GridDesdobr.Rows - 1

If dnfe.State = 1 Then
   dnfe.Close: Set dnfe = Nothing
End If

dnfe.Open "SELECT * FROM notafiscaldetprod WHERE chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
If dnfe.EOF Then
   Call RotinaCriarDetProd
End If

If dnfe.State = 1 Then
   dnfe.Close: Set dnfe = Nothing
End If

dnfe.Open "SELECT * FROM notafiscaldetprod WHERE chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "')", db, 3, 3
If dnfe.EOF Then
   MsgBox ("Det prod não gerado"), vbInformation
   dnfe.Close
   Exit Sub
End If

dnfe.MoveFirst

Do While Not dnfe.EOF

   If NumParcelas > 0 Then
      dnfe!nfdQtdParcelas = NumParcelas
      dnfe!nfdValorParcela = dnfe!nfdValorDaCompra / NumParcelas
      dnfe.Update
   End If
   
   dnfe.MoveNext
   
Loop

End Sub

Public Sub RotinaCriarDetProd()


For Ind = 1 To GridProduto.Rows - 1

   If dnfe.State = 1 Then
      dnfe.Close: Set dnfe = Nothing
   End If

   dnfe.Open "SELECT * FROM notafiscaldetprod WHERE chPessoa = ('" & cmbPessoa & "') and chNotaFiscalEntrada = ('" & txtNotaFiscal & "') and chCodProduto = ('" & GridProduto.TextMatrix(Ind, 0) & "')", db, 3, 3
   If dnfe.EOF Then
      dnfe.AddNew
   End If

   AcheiSupProduto = 0
   dnfe!chPessoa = cmbPessoa
   dnfe!chNotaFiscalEntrada = txtNotaFiscal
   dnfe!chCodProduto = GridProduto.TextMatrix(Ind, 0)
   dnfe!chProdutoFabrica = lblProdutoFabrica
   dnfe!nfdQtd = GridProduto.TextMatrix(Ind, 3)
   dnfe!nfdPU = GridProduto.TextMatrix(Ind, 4)
   txtQtd = GridProduto.TextMatrix(Ind, 3)
   txtPU = GridProduto.TextMatrix(Ind, 4)
   txtValor = Format$(txtQtd * txtPU, "##,##0.00")
   dnfe!nfdValorDaCompra = Format$(txtQtd * txtPU, "##,##0.00")
   dnfe!nfdQtdParcelas = NumParcelas
   dnfe!nfdValorParcela = Format$((txtValor / NumParcelas), "##,##0.00")
   
   If Prod.State = 1 Then
      Prod.Close: Set Prod = Nothing
   End If
   
   Prod.Open "Select * from produtoentrada where chPessoa = ('" & cmbPessoa & "') and chTipoProduto = ('" & cmbCodProduto & "')", db, 3, 3
   If Prod.EOF Then
      Prod.Close
      Prod.Open "Select * from produtofornecedor where chTipoProduto = ('" & cmbPessoa & "') and chProdutoFabrica = ('" & cmbCodProduto & "')", db, 3, 3
      If Prod.EOF Then
         Prod.Close
         Prod.Open "Select * from supproduto where nomeProd = ('" & GridProduto.TextMatrix(Ind, 0) & "')", db, 3, 3
         If Prod.EOF Then
            MsgBox ("ERRO: Comunicar ao analista responsável."), vbCritical
            Call FechaDB
            Exit Sub
         Else
            AcheiSupProduto = 1
         End If
      End If
   End If


   If AcheiSupProduto = 0 Then
      dnfe!nfdCentroDeCusto = Prod!pinCentroDeCusto
      dnfe!nfdGrupoCentroDeCusto = Prod!pinGrupoCentroDeCusto
      dnfe!nfdSubGrupoCentroDeCusto = Prod!pinSubGrupoCentroDeCusto
   Else
      dnfe!nfdCentroDeCusto = Prod!centrodecusto
      dnfe!nfdGrupoCentroDeCusto = Prod!GrupoCentroDeCusto
      dnfe!nfdSubGrupoCentroDeCusto = Prod!SubGrupoCentroDeCusto
   End If
   
   dnfe.Update
   
Next

End Sub


