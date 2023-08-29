VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPagtoEmCheque 
   Caption         =   "frmPagtoEmCheque"
   ClientHeight    =   7290
   ClientLeft      =   3750
   ClientTop       =   3255
   ClientWidth     =   10245
   LinkTopic       =   "Form3"
   ScaleHeight     =   7290
   ScaleWidth      =   10245
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Height          =   735
      Left            =   120
      TabIndex        =   58
      Top             =   0
      Width           =   5055
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pagamento de Ordem de Carga"
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
         TabIndex        =   59
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "Pagamento da Carga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   21
      Top             =   3960
      Width           =   9975
      Begin VB.Frame Frame3 
         Caption         =   "Operação"
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
         Left            =   7560
         TabIndex        =   38
         Top             =   1920
         Width           =   2295
         Begin VB.CommandButton cmdSair 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Sair"
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
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdNovo 
            BackColor       =   &H0080FF80&
            Caption         =   "Novo"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdExclui 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Exc."
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
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         Height          =   2415
         Left            =   6720
         TabIndex        =   37
         Top             =   240
         Width           =   855
         Begin VB.CommandButton cmdRetira 
            BackColor       =   &H0080C0FF&
            Cancel          =   -1  'True
            Caption         =   "Ret."
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1200
            Width           =   615
         End
         Begin VB.CommandButton cmdAltera 
            BackColor       =   &H0080FFFF&
            Caption         =   "Alt."
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton cmdProcessa 
            BackColor       =   &H000000FF&
            Caption         =   "Proc"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton cmdIncluir 
            BackColor       =   &H00FFFF00&
            Caption         =   "Inc."
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Resumo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   7560
         TabIndex        =   30
         Top             =   240
         Width           =   2295
         Begin VB.Label lblVPedagio 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1200
            TabIndex        =   57
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "V. Pedágio"
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
            TabIndex        =   56
            Top             =   600
            Width           =   945
         End
         Begin VB.Label Label6 
            Caption         =   "Tot. Frete"
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
            TabIndex        =   36
            Top             =   240
            Width           =   855
         End
         Begin VB.Label txtValorDoFrete 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1200
            TabIndex        =   35
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Tot. Cheque"
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
            Left            =   120
            TabIndex        =   34
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label txtValorTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1200
            TabIndex        =   33
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Diferença"
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
            TabIndex        =   32
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label txtDiferenca 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1200
            TabIndex        =   31
            Top             =   1320
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1695
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   6615
         Begin VB.PictureBox GridCheque 
            Height          =   1335
            Left            =   120
            ScaleHeight     =   1275
            ScaleWidth      =   6315
            TabIndex        =   29
            Top             =   240
            Width           =   6375
         End
      End
      Begin VB.Frame Frame11 
         Height          =   735
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   6615
         Begin MSComCtl2.DTPicker txtDataComp 
            Height          =   255
            Left            =   3840
            TabIndex        =   6
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   59703297
            CurrentDate     =   38162
         End
         Begin VB.ComboBox cmbTipo 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cmbBanco 
            Height          =   315
            Left            =   1320
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtNumDoCheque 
            Height          =   285
            Left            =   2640
            TabIndex        =   5
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5160
            TabIndex        =   7
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo"
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
            TabIndex        =   27
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Banco"
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
            TabIndex        =   26
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Nº  Doc"
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
            Left            =   2640
            TabIndex        =   25
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Data Comp"
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
            TabIndex        =   24
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label5 
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
            Left            =   5160
            TabIndex        =   23
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Status"
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
      Left            =   5160
      TabIndex        =   19
      Top             =   0
      Width           =   2775
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Hoje"
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
      Left            =   8040
      TabIndex        =   16
      Top             =   0
      Width           =   2055
      Begin VB.TextBox txtDataHoje 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
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
         TabIndex        =   39
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9975
      Begin VB.Frame Frame6 
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
         Left            =   5040
         TabIndex        =   51
         Top             =   840
         Width           =   4815
         Begin VB.PictureBox GridRomaneio 
            Height          =   1575
            Left            =   120
            ScaleHeight     =   1515
            ScaleWidth      =   4515
            TabIndex        =   52
            Top             =   720
            Width           =   4575
         End
         Begin VB.ComboBox cmbRomaneio 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   360
            Width           =   2950
         End
         Begin VB.Label lblValorTotalVale 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3120
            TabIndex        =   55
            Top             =   360
            Width           =   1250
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor Total"
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
            Left            =   2880
            TabIndex        =   54
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "Romaneio"
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
            TabIndex        =   53
            Top             =   120
            Width           =   2895
         End
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   9735
         Begin VB.ComboBox cmbEmissorNF 
            Height          =   315
            Left            =   1680
            TabIndex        =   2
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox cmbOrdemDeCarga 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "Placa"
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
            Left            =   8520
            TabIndex        =   50
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label txtPlaca 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   8520
            TabIndex        =   49
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Motorista"
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
            Left            =   6240
            TabIndex        =   48
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label txtMotorista 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6240
            TabIndex        =   47
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Data da Carga"
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
            Left            =   4920
            TabIndex        =   46
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label txtDataCarga 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4920
            TabIndex        =   45
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Operação"
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
            Left            =   3360
            TabIndex        =   44
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label txtDescOperacao 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3360
            TabIndex        =   43
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Emissor"
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
            TabIndex        =   42
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Ordem de Carga"
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
            TabIndex        =   41
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Clientes - Frete"
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
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   4815
         Begin VB.PictureBox GridNotaFiscal 
            Height          =   2055
            Left            =   120
            ScaleHeight     =   1995
            ScaleWidth      =   4515
            TabIndex        =   18
            Top             =   240
            Width           =   4575
         End
      End
   End
End
Attribute VB_Name = "frmPagtoEmCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Ind As Byte
Dim Final As Byte
Dim Fim As Byte
Dim Parm As Byte
Dim Resp As String
Dim ChaveLeituraNeg As String
Dim ClienteNeg As String
Dim Quebra As String
Dim Erro As Byte
Dim Indice As Byte
Dim ContaOc As Integer
Dim TabEmissor(20) As String

Dim ValorTotal As Currency
Dim valorcheque As Currency

Dim OrdemDeCarga As String
Dim AcumulaValorCheque As Currency
Dim AcumulaValor As Currency
Dim AcumulaCheque As Currency
Dim ValorDoCheque As Currency
Dim AcumTarifa As Currency
Dim DataUtil As Date

Private Sub cmbEmissorNF_lostfocus()

If cmbOrdemDeCarga = Empty Then
   Exit Sub
End If

If ContaOc = 0 Then
   MsgBox ("Sem ordem de carga pendente. Vou assumir Emissor = MaxAlbido")
   TabPagtosEmCheque.Seek "=", cmbOrdemDeCarga, "MaxAlbido"
Else
   TabPagtosEmCheque.Seek "=", cmbOrdemDeCarga, cmbEmissorNF
End If

If TabPagtosEmCheque.NoMatch Then
   MsgBox ("Ordem de Carga não Cadastrada. Verifique!")
   cmdSair.SetFocus
   Exit Sub
Else
   GridRomaneio.Rows = 2
   GridRomaneio.TextMatrix(1, 0) = Empty
   GridRomaneio.TextMatrix(1, 1) = Empty
   lblVPedagio = Format$(0, "0.00")
   GridCheque.Rows = 2
   GridCheque.TextMatrix(1, 0) = Empty
   GridCheque.TextMatrix(1, 1) = Empty
   GridCheque.TextMatrix(1, 2) = Empty
   GridCheque.TextMatrix(1, 3) = Empty
   GridCheque.TextMatrix(1, 4) = Empty
   
   GridNotaFiscal.Rows = 2
   GridNotaFiscal.TextMatrix(1, 0) = Empty
   GridNotaFiscal.TextMatrix(1, 1) = Empty
   GridNotaFiscal.TextMatrix(1, 2) = Empty
   
   Call Rotina_Carrega_OrdemDeCarga
   
   If Not TabPagtosEmCheque("chromaneio") = 0 Then
      TabRomaneio.Seek "=", TabPagtosEmCheque("chromaneio")
      If TabRomaneio.NoMatch Then
         'cmbRomaneio = TabRomaneio("dstromaneio")
         cmbRomaneio.ListIndex = 0
         lblValorTotalVale = Format$(0, "0.00")
      Else
         cmbRomaneio = TabRomaneio("dstromaneio")
         Call Rotina_070_Carga_Romaneio
      End If
   Else
      lblValorTotalVale = Format$(0, "0.00")
      cmbRomaneio = ""
      AcumTarifa = 0
   End If
   
   Call Rotina_075_Carga_GridNotaFiscal
     
   Call Rotina_080_Carga_GridCheque
   
   If txtDiferenca = 0 Then
      cmdIncluir.Enabled = False
      cmdAltera.Enabled = False
      cmdRetira.Enabled = False
      If TabPagtosEmCheque("ocgstatus") = 0 Then
         cmdProcessa.Enabled = True
      Else
         cmdProcessa.Enabled = False
      End If
   Else
      cmdIncluir.Enabled = True
      cmdAltera.Enabled = True
      cmdRetira.Enabled = True
      cmdProcessa.Enabled = False
   End If
End If

End Sub



Private Sub cmbRomaneio_lostfocus()
'If cmbRomaneio = Empty Then
'   tabpagtosemcheque.
Call Rotina_070_Carga_Romaneio
AcumTarifa = lblVPedagio
If IsNumeric(txtValorDoFrete) Then
   txtDiferenca = Format$(txtValorDoFrete - (txtValorTotal + AcumTarifa), "##0.00")
End If
End Sub

Private Sub cmbTipo_lostfocus()
If cmbTipo = "Espécie" Then
   cmbBanco = "Cx.MaxAlbido"
   txtNumDoCheque = "Espécie"
   txtDataComp = Date
   txtNumDoCheque.SetFocus
End If
End Sub

Private Sub cmdAltera_Click()
TabDetalhePagtosEmCheque.Seek "=", cmbOrdemDeCarga, cmbEmissorNF, cmbBanco, txtNumDoCheque
If TabDetalhePagtosEmCheque.NoMatch Then
   MsgBox ("Tipo e Numero do documento não encontrado. Verifique.")
   Exit Sub
End If

ValorTotal = txtValorTotal
valorcheque = txtValor

AcumulaValor = (ValorTotal + valorcheque) - TabDetalhePagtosEmCheque("docgvalordocheque")

If AcumulaValor > txtValorDoFrete Then
   MsgBox ("Valores de cheques superam o valor do somatório dos fretes a serem pagos")
Else
   TabDetalhePagtosEmCheque.Edit
   TabDetalhePagtosEmCheque("docgdatacompensacao") = txtDataComp
   TabDetalhePagtosEmCheque("docgvalordocheque") = txtValor
   TabDetalhePagtosEmCheque.Update
End If

Rotina_080_Carga_GridCheque
Call Rotina_Limpa_Cheque
If txtDiferenca = 0 Then
   cmdIncluir.Enabled = False
   cmdAltera.Enabled = False
   cmdRetira.Enabled = False
   If TabPagtosEmCheque("ocgstatus") = 0 Then
      cmdProcessa.Enabled = True
   Else
      cmdProcessa.Enabled = False
   End If
Else
   cmdIncluir.Enabled = False
   cmdAltera.Enabled = False
   cmdRetira.Enabled = False
   cmdProcessa.Enabled = False
   cmbTipo.SetFocus
End If
End Sub

Private Sub cmdExclui_Click()
If TabPagtosEmCheque("ocgStatus") = 1 Then
   MsgBox "Exclusão de Ordem de Carga já processada. Proibido"
   cmdSair.SetFocus
   Exit Sub
End If

Fim = 0

TabNegociacao.MoveFirst

Do While Fim = 0
   If TabNegociacao("chordemdecarga") = cmbOrdemDeCarga Then
      Fim = 2
   Else
      TabNegociacao.MoveNext
      If TabNegociacao.EOF Then
         Fim = 1
      End If
   End If
Loop

If Fim = 2 Then
   MsgBox "Exclusão não permitida. Pedido Processado para esta Ordem de Carga"
   cmdSair.SetFocus
   Exit Sub
End If

Resp = MsgBox("Deseja Excluir esta Ordem de Carga??? Confirma.", vbYesNo)
If Resp = vbYes Then
   TabDetPgCheque.Seek "=", TabPagtosEmCheque("chordemdecarga"), TabPagtosEmCheque("chemissor")
   Do While Not TabDetPgCheque.NoMatch
      TabDetPgCheque.Delete
      TabDetPgCheque.Seek "=", TabPagtosEmCheque("chordemdecarga"), TabPagtosEmCheque("chemissor")
   Loop
   
   TabPagtosEmCheque.Delete
   
   Call Rotina_010_Limpa_OrdemDeCarga
   cmbOrdemDeCarga = Empty
   cmbEmissorNF = Empty
   lblStatus = Empty
   cmbTipo = Empty
   cmbBanco = Empty
   txtNumDoCheque = Empty
   txtValor = Format$(0, "0.00")
   txtDataComp = Date
   Call Rotina_020_Limpa_GridNotaFiscal
   Call Rotina_Limpa_Cheque
   Call Rotina_Limpa_GridCheque
   txtValorDoFrete = Format$(0, "0.00")
   txtValorTotal = Format$(0, "0.00")
   txtDiferenca = Format$(0, "0.00")
      
   cmdSair.SetFocus
   
   MsgBox "Ordem de Craga Deletada"
   
End If

End Sub

Private Sub cmdIncluir_Click()

If cmbOrdemDeCarga = "" Then
   MsgBox ("Solicitação de Inclusão sem informar a Ordem de Carga")
   Exit Sub
End If

ValorTotal = txtValorTotal
valorcheque = txtValor

AcumulaValor = ValorTotal + valorcheque

If AcumulaValor > txtValorDoFrete - AcumTarifa Then
   MsgBox ("Valores de cheques superam o valor do somatório dos fretes a serem pagos")
Else
   TabDetalhePagtosEmCheque.AddNew
   TabDetalhePagtosEmCheque("chordemdecarga") = cmbOrdemDeCarga
   TabDetalhePagtosEmCheque("chemissOR") = cmbEmissorNF
   TabDetalhePagtosEmCheque("chbanco") = cmbBanco
   TabDetalhePagtosEmCheque("chnumdoc") = txtNumDoCheque
   TabDetalhePagtosEmCheque("chordemdecarga") = cmbOrdemDeCarga
   TabDetalhePagtosEmCheque("docgdatacompensacao") = txtDataComp
   TabDetalhePagtosEmCheque("docgvalordocheque") = txtValor
   TabDetalhePagtosEmCheque("docgstatus") = 0
   TabDetalhePagtosEmCheque.Update
End If

Call Rotina_080_Carga_GridCheque
Call Rotina_Limpa_Cheque
If txtDiferenca = 0 Then
   cmdIncluir.Enabled = False
   cmdAltera.Enabled = False
   cmdRetira.Enabled = False
   If TabPagtosEmCheque("ocgstatus") = 0 Then
      cmdProcessa.Enabled = True
   Else
      cmdProcessa.Enabled = False
   End If
Else
   cmdIncluir.Enabled = False
   cmdAltera.Enabled = False
   cmdRetira.Enabled = False
   cmdProcessa.Enabled = False
   cmbTipo.SetFocus
End If
End Sub

Private Sub cmdNovo_Click()
Call Rotina_010_Limpa_OrdemDeCarga
cmbOrdemDeCarga = Empty
cmbEmissorNF = Empty
lblStatus = Empty
cmbTipo = Empty
cmbBanco = Empty
txtNumDoCheque = Empty
txtValor = Format$(0, "0.00")
txtDataComp = Date
Call Rotina_020_Limpa_GridNotaFiscal
Call Rotina_Limpa_Cheque
Call Rotina_Limpa_GridCheque
txtValorDoFrete = Format$(0, "0.00")
txtValorTotal = Format$(0, "0.00")
txtDiferenca = Format$(0, "0.00")
cmbOrdemDeCarga.SetFocus
End Sub

Private Sub cmdProcessa_Click()

If cmbOrdemDeCarga = "" Then
   MsgBox ("Solicitação de Processamento sem informar a Ordem de Carga")
   Exit Sub
End If

TabDetalhePagtosEmCheque.MoveFirst
Do While Not TabDetalhePagtosEmCheque.EOF
   If TabDetalhePagtosEmCheque("chordemdecarga") = cmbOrdemDeCarga Then
      TabCtaPagar.Seek "=", 0, txtMotorista, TabDetalhePagtosEmCheque("chnumdoc"), TabDetalhePagtosEmCheque("chnumdoc"), TabDetalhePagtosEmCheque("docgdatacompensacao")
      If TabCtaPagar.NoMatch Then
          TabCtaPagar.AddNew
          TabCtaPagar("chfabricante") = 0
          TabCtaPagar("chpessoa") = txtMotorista
          TabCtaPagar("chnotafiscal") = TabDetalhePagtosEmCheque("chnumdoc")
          TabCtaPagar("chfatura") = TabDetalhePagtosEmCheque("chnumdoc")
          TabCtaPagar("chdatavencito") = TabDetalhePagtosEmCheque("docgdatacompensacao")
                                                    
          'Calcula data banco
                   
          DataUtil = TabDetalhePagtosEmCheque("docgdatacompensacao")
            
          Datainformada = DataUtil
          NDias = 0
          DataRetorno = ObterProximoDiaUtil(Datainformada, NDias)
          DataUtil = DataRetorno.DiaUtil
          TabCtaPagar("ctpDataBanco") = DataUtil
         
          'Fim calcula data banco
    
          TabCtaPagar("ctpdataemissao") = Date
          TabCtaPagar("ctpdatalanc") = Date
          TabCtaPagar("ctpdatavencoriginal") = TabDetalhePagtosEmCheque("docgdatacompensacao")
          TabCtaPagar("ctpdescricaooperacao") = TabDetalhePagtosEmCheque("CHORDEMDECARGA")
          TabCtaPagar("ctpvalorlart") = TabDetalhePagtosEmCheque("docgvalordocheque")
          TabCtaPagar("ctpvalormerco") = 0
          TabCtaPagar("ctpvalordaboleta") = TabDetalhePagtosEmCheque("docgvalordocheque")
          TabCtaPagar("chano") = Year(Date)
          TabCtaPagar("CHMES") = Month(Date)
          TabCtaPagar("chdia") = Day(Date)
          TabCtaPagar("chcodbcolart") = TabDetalhePagtosEmCheque("chbanco")
          If TabDetalhePagtosEmCheque("chnumdoc") = "Espécie" Then
             TabCtaPagar("ctpstatus") = 1
             TabCtaPagar("ctpDataPagamento") = txtDataComp
          Else
             TabCtaPagar("ctpstatus") = 0
          End If
          TabCtaPagar("ctpdataproc") = Date
          'TabCtaPagar("ctpDataPagamento") = Date
          TabCtaPagar("ctptipolancamento") = 8
          TabCtaPagar.Update
       End If
   End If
   TabDetalhePagtosEmCheque.MoveNext
Loop
TabPagtosEmCheque.Edit
TabPagtosEmCheque("ocgStatus") = 1
TabPagtosEmCheque("ocgvalorpedagio") = AcumTarifa
TabPagtosEmCheque("ocgvalorfrete") = TabPagtosEmCheque("ocgvalortotal") - AcumTarifa
'TabPagtosEmCheque("CHromaneio") = TabNomeRomaneio("chCODromaneio")
'ALTERADO EM 21/09/2005
If cmbRomaneio = "" Then
   TabPagtosEmCheque("CHromaneio") = 0
Else
   TabNomeRomaneio.Seek "=", cmbRomaneio
   If TabNomeRomaneio.NoMatch Then
      MsgBox "Nome do Romaneio não encontrado. Vou assumir o primeiro da lista"
      TabPagtosEmCheque("chromaneio") = 1
   Else
      TabPagtosEmCheque("CHromaneio") = TabNomeRomaneio("chcodromaneio")
   End If
End If

TabPagtosEmCheque.Update

Call Rotina_Carrega_OrdemDeCarga
     
Call Rotina_075_Carga_GridNotaFiscal
     
Call Rotina_080_Carga_GridCheque
  
If txtDiferenca = 0 Then
   cmdIncluir.Enabled = False
   cmdAltera.Enabled = False
   cmdRetira.Enabled = False
   If TabPagtosEmCheque("ocgstatus") = 0 Then
      cmdProcessa.Enabled = True
   Else
      cmdProcessa.Enabled = False
   End If
Else
   cmdIncluir.Enabled = False
   cmdAltera.Enabled = False
   cmdRetira.Enabled = False
   cmdProcessa.Enabled = False
   cmbTipo.SetFocus
End If
End Sub

Private Sub cmdRetira_Click()
TabDetalhePagtosEmCheque.Seek "=", cmbOrdemDeCarga, cmbEmissorNF, cmbBanco, txtNumDoCheque
If TabDetalhePagtosEmCheque.NoMatch Then
   MsgBox ("Tipo e Numero do documento não encontrado para retirada. Verifique.")
   Exit Sub
End If

TabPagtosEmCheque.Seek "=", cmbOrdemDeCarga, cmbEmissorNF
If TabPagtosEmCheque.NoMatch Then
   MsgBox "Ordem de Carga não encontrada"
   Exit Sub
Else
   TabCtaPagar.Seek "=", 0, txtMotorista, txtNumDoCheque, txtNumDoCheque, txtDataComp
   If TabCtaPagar.NoMatch Then
      MsgBox "Ordem de Carga não existente em contas a pagar"
      Exit Sub
   Else
      If TabCtaPagar("ctpstatus") = 0 Then
         TabCtaPagar.Delete
         TabDetalhePagtosEmCheque.Delete
         TabPagtosEmCheque.Edit
         TabPagtosEmCheque("ocgstatus") = 0
         TabPagtosEmCheque.Update
      Else
         MsgBox "Documento com este número já foi pago. Retirada Inválida"
         Exit Sub
      End If
   End If
End If

Rotina_080_Carga_GridCheque
Call Rotina_Limpa_Cheque
If txtDiferenca = 0 Then
   cmdIncluir.Enabled = False
   cmdAltera.Enabled = False
   cmdRetira.Enabled = False
   If TabPagtosEmCheque("ocgstatus") = 0 Then
      cmdProcessa.Enabled = True
   Else
      cmdProcessa.Enabled = False
   End If
Else
   cmdIncluir.Enabled = False
   cmdAltera.Enabled = False
   cmdRetira.Enabled = False
   cmdProcessa.Enabled = False
   cmbTipo.SetFocus
End If
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtDataHoje = Date
txtDataComp = Date
ContaOc = 0
For Ind = 0 To 19
    TabEmissor(Ind) = Empty
Next

TabRomaneio.MoveFirst
Do While Not TabRomaneio.EOF
   cmbRomaneio.AddItem TabRomaneio("dstromaneio")
   TabRomaneio.MoveNext
Loop

cmbEmissorNF.AddItem "MaxAlbido"
cmbEmissorNF.AddItem "MERCOPISO"

cmbEmissorNF.ListIndex = 0

cmbTipo.AddItem "Cheque"
cmbTipo.AddItem "Espécie"

tabbanco.MoveFirst
Do While Not tabbanco.EOF
   cmbBanco.AddItem tabbanco("bcoSiglabco")
   tabbanco.MoveNext
Loop

TabPagtosEmCheque.MoveFirst
Do While Not TabPagtosEmCheque.EOF
   If TabPagtosEmCheque("OCGSTATUS") = 0 Then
      cmbOrdemDeCarga.AddItem TabPagtosEmCheque("chordemdecarga")
      TabEmissor(ContaOc) = TabPagtosEmCheque("chemissor")
      ContaOc = ContaOc + 1
   End If
   TabPagtosEmCheque.MoveNext
Loop
If ContaOc = 0 Then
   MsgBox ("Não há Ordem de Carga pendente de processamento")
   ContaOc = 1
   Exit Sub
End If

cmbOrdemDeCarga.ListIndex = 0

Call Rotina_010_Limpa_OrdemDeCarga

Call Rotina_020_Limpa_GridNotaFiscal

End Sub

Public Sub Rotina_010_Limpa_OrdemDeCarga()

txtDataCarga = "__/__/____"
txtMotorista = Empty
txtPlaca = Empty
txtDescOperacao = Empty
'cmbEmissorNF = Empty
lblStatus = Empty
End Sub

Public Sub Rotina_020_Limpa_GridNotaFiscal()

GridNotaFiscal.Rows = 2
GridNotaFiscal.TextMatrix(1, 0) = Empty
GridNotaFiscal.TextMatrix(1, 1) = Empty
GridNotaFiscal.TextMatrix(1, 2) = Empty

End Sub

Public Sub Rotina_070_Carga_Romaneio()
If cmbRomaneio = "" Then
   lblValorTotalVale = Format$(0, "#0.00")
   GridRomaneio.Rows = 2
   GridRomaneio.TextMatrix(1, 0) = Empty
   GridRomaneio.TextMatrix(1, 1) = Empty
   lblVPedagio = Format$(0, "#0.00")
   Exit Sub
End If

TabNomeRomaneio.Seek "=", cmbRomaneio
If TabNomeRomaneio.NoMatch Then
   MsgBox ("Esta Praça não esta Cadastrada. Proceder ao Cadastramento e somente depois retornar a esta rotina")
   Exit Sub
End If

Fim = 0
Ind = 0
AcumTarifa = 0
TabRomaneioPraca.MoveFirst
Do While Fim = 0
   If TabRomaneioPraca("chromaneio") = TabNomeRomaneio("chcodromaneio") Then
      Ind = Ind + 1
      GridRomaneio.Rows = Ind + 1
      TabPracaPedagio.Seek "=", TabRomaneioPraca("chpracapedagio")
      If TabPraca.NoMatch Then
         MsgBox ("Erro no acesso a praça de pedágio")
         Ind = Ind / 0
      End If
      GridRomaneio.TextMatrix(Ind, 0) = TabPracaPedagio("ppdpraca")
      GridRomaneio.TextMatrix(Ind, 1) = Format$(TabPracaPedagio("ppdtarifa"), "##0.00")
      AcumTarifa = AcumTarifa + TabPracaPedagio("ppdtarifa")
   End If
   If TabRomaneioPraca("chromaneio") > TabNomeRomaneio("chcodromaneio") Then
      Fim = 1
   Else
      TabRomaneioPraca.MoveNext
      If TabRomaneioPraca.EOF Then
         Fim = 1
      End If
   End If
Loop

lblValorTotalVale = Format$(AcumTarifa, "##0.00")
lblVPedagio = Format$(AcumTarifa, "##0.00")

End Sub
Public Sub Rotina_075_Carga_GridNotaFiscal()

'If Not TabDetalhePessoaFrete.EOF Then
   TabDetalhePessoaFrete.MoveFirst
'End If

Ind = 1
ValorDoCheque = 0
Do While Not TabDetalhePessoaFrete.EOF
   If TabDetalhePessoaFrete("chordemdecarga") = cmbOrdemDeCarga Then
      GridNotaFiscal.Rows = Ind + 1
      GridNotaFiscal.TextMatrix(Ind, 0) = TabDetalhePessoaFrete("chPessoa")
      GridNotaFiscal.TextMatrix(Ind, 1) = TabDetalhePessoaFrete("chnotafiscal")
      GridNotaFiscal.TextMatrix(Ind, 2) = Format$(TabDetalhePessoaFrete("dfpvalor"), "##0.00")
      ValorDoCheque = ValorDoCheque + TabDetalhePessoaFrete("dfpvalor")
      Ind = Ind + 1
   End If
   TabDetalhePessoaFrete.MoveNext
Loop

txtValorDoFrete = Format$(ValorDoCheque, "##0.00")

End Sub

Public Sub Rotina_080_Carga_GridCheque()

GridCheque.Rows = 2
GridCheque.TextMatrix(1, 0) = ""
GridCheque.TextMatrix(1, 1) = ""
GridCheque.TextMatrix(1, 2) = ""
GridCheque.TextMatrix(1, 3) = ""
GridCheque.TextMatrix(1, 4) = ""

'If Not TabDetalhePagtosEmCheque.EOF Then
   TabDetalhePagtosEmCheque.MoveFirst
'End If

Ind = 1
AcumulaCheque = 0

Do While Not TabDetalhePagtosEmCheque.EOF
   If TabDetalhePagtosEmCheque("chordemdecarga") = cmbOrdemDeCarga Then
      GridCheque.Rows = Ind + 1
      If TabDetalhePagtosEmCheque("chnumdoc") = "Espécie" Then
         GridCheque.TextMatrix(Ind, 0) = "Espécie"
         GridCheque.TextMatrix(Ind, 1) = "Cx.MaxAlbido"
      Else
         GridCheque.TextMatrix(Ind, 0) = "Cheque"
         GridCheque.TextMatrix(Ind, 1) = TabDetalhePagtosEmCheque("chbanco")
      End If
      GridCheque.TextMatrix(Ind, 2) = TabDetalhePagtosEmCheque("chnumdoc")
      GridCheque.TextMatrix(Ind, 3) = TabDetalhePagtosEmCheque("docgdatacompensacao")
      GridCheque.TextMatrix(Ind, 4) = Format$(TabDetalhePagtosEmCheque("docgvalordocheque"), "##0.00")
      AcumulaCheque = AcumulaCheque + TabDetalhePagtosEmCheque("docgvalordocheque")
      Ind = Ind + 1
   End If
   TabDetalhePagtosEmCheque.MoveNext
Loop
txtValorTotal = Format$(AcumulaCheque, "##0.00")
txtDiferenca = Format$(txtValorDoFrete - (txtValorTotal + AcumTarifa), "##0.00")
If txtDiferenca = 0 Then
   cmdIncluir.Enabled = False
   If TabPagtosEmCheque("ocgstatus") = 0 Then
      cmdProcessa.Enabled = True
      cmdRetira.Enabled = True
      cmdAltera.Enabled = True
   Else
      cmdProcessa.Enabled = False
      cmdRetira.Enabled = False
      cmdAltera.Enabled = False
   End If
Else
   cmdIncluir.Enabled = True
   cmdAltera.Enabled = False
   cmdRetira.Enabled = False
   cmdProcessa.Enabled = False
End If
End Sub

Public Sub Rotina_Carrega_OrdemDeCarga()

cmbEmissorNF = TabPagtosEmCheque("chemissor")
txtDataCarga = TabPagtosEmCheque("ocgDatadacarga")
txtDescOperacao = TabPagtosEmCheque("ocgDescOperacao")
txtMotorista = TabPagtosEmCheque("ocgmotorista")
txtPlaca = TabPagtosEmCheque("ocgPlaca")
If TabPagtosEmCheque("ocgstatus") = 0 Then
   lblStatus = "Pendente"
Else
   lblStatus = "Processado"
End If
End Sub
Private Sub GridCheque_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not GridCheque.TextMatrix(GridCheque.Row, 0) = "" Then
    Resp = MsgBox("Deseja Restaurar este lançamento", vbYesNo)
    If Resp = vbYes Then
       cmbTipo = GridCheque.TextMatrix(GridCheque.Row, 0)
       cmbBanco = GridCheque.TextMatrix(GridCheque.Row, 1)
       txtNumDoCheque = GridCheque.TextMatrix(GridCheque.Row, 2)
       txtDataComp = GridCheque.TextMatrix(GridCheque.Row, 3)
       txtValor = GridCheque.TextMatrix(GridCheque.Row, 4)
       txtNumDoCheque.SetFocus
    End If
End If
End Sub

Private Sub txtDataComp_LostFocus()
If txtDataComp < Date Then
   Resp = MsgBox("Data informada para compensação do cheque anterior a data de hoje. Confirma???", vbYesNo)
   If Resp = vbNo Then
      txtDataComp.SetFocus
   End If
End If
End Sub

Private Sub txtNumDoCheque_LostFocus()

TabDetalhePagtosEmCheque.Seek "=", cmbOrdemDeCarga, cmbEmissorNF, cmbBanco, txtNumDoCheque
If TabDetalhePagtosEmCheque.NoMatch Then
   If txtDiferenca > 0 Then
      cmdIncluir.Enabled = True
      cmdAltera.Enabled = False
      cmdRetira.Enabled = False
      cmdProcessa.Enabled = False
   Else
      cmdIncluir.Enabled = False
      cmdAltera.Enabled = True
      cmdRetira.Enabled = True
      cmdProcessa.Enabled = False
      'MsgBox ("Não permitida inclusão de novos valores.")
      'cmbTipo.SetFocus
      cmdAltera.SetFocus
   End If
Else
   txtDataComp = TabDetalhePagtosEmCheque("docgdatacompensacao")
   txtValor = Format$(TabDetalhePagtosEmCheque("docgvalordocheque"), "##0.00")
   cmdIncluir.Enabled = False
   cmdProcessa.Enabled = False
   cmdRetira.Enabled = True
   cmdAltera.Enabled = True
End If
End Sub

Public Sub Rotina_Limpa_Cheque()
cmbTipo = Empty
cmbBanco = Empty
txtNumDoCheque = Empty
txtValor = Format$(0, "0.00")

txtDataComp = Date

End Sub

Public Sub Rotina_Limpa_GridCheque()

GridCheque.Rows = 2
GridCheque.TextMatrix(1, 0) = Empty
GridCheque.TextMatrix(1, 1) = Empty
GridCheque.TextMatrix(1, 2) = Empty
GridCheque.TextMatrix(1, 3) = Empty
GridCheque.TextMatrix(1, 4) = Empty
End Sub
