VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmControlePagamentos 
   Caption         =   "frmControlePagamentos"
   ClientHeight    =   10575
   ClientLeft      =   4185
   ClientTop       =   3630
   ClientWidth     =   20370
   LinkTopic       =   "Form3"
   ScaleHeight     =   10575
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consolidação de Contas a Pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   36
      Top             =   8280
      Width           =   12975
      Begin VB.ComboBox cmbConsolidado 
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
         Left            =   5520
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   3495
      End
      Begin VB.ComboBox cmbTipoPagto 
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
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   9240
         TabIndex        =   39
         Top             =   960
         Width           =   3495
         Begin VB.OptionButton optNao 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Desmarcar Ok"
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
            Left            =   1680
            TabIndex        =   8
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton optSim 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Marcar Ok"
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
            TabIndex        =   7
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   495
            Left            =   960
            TabIndex        =   40
            Top             =   -1200
            Width           =   1215
         End
      End
      Begin VB.TextBox txtValorTotalConsolidado 
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
         Left            =   9480
         TabIndex        =   6
         Top             =   480
         Width           =   3015
      End
      Begin VB.ComboBox cmbTipoConsolidacao 
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
         TabIndex        =   4
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblConsolidado 
         BackColor       =   &H00E0E0E0&
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
         Left            =   5520
         TabIndex        =   42
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Operações"
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
         TabIndex        =   41
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Total da consolidação"
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
         Left            =   9600
         TabIndex        =   38
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Consolidar por (Forma)"
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
         Left            =   2280
         TabIndex        =   37
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   8655
      Begin VB.Label frmControlePagamentos 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Controle de Pagamentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   300
         Width           =   8415
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   8880
      TabIndex        =   25
      Top             =   0
      Width           =   9375
      Begin VB.Frame Frame13 
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
         Height          =   855
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   2220
         Begin VB.ComboBox cmbFiltro 
            BackColor       =   &H00FFFFEA&
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
            TabIndex        =   0
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Data Compensação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   840
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         Begin MSComCtl2.DTPicker txtDataComp 
            Height          =   375
            Left            =   1440
            TabIndex        =   11
            Top             =   240
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
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
            CalendarBackColor=   16777194
            Format          =   392757249
            CurrentDate     =   38203
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.Frame Frame11 
         Caption         =   "Data Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2520
         TabIndex        =   26
         Top             =   120
         Width           =   2295
         Begin MSComCtl2.DTPicker txtDataConsulta 
            Height          =   435
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   767
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
            CalendarBackColor=   16777194
            Format          =   392757249
            CurrentDate     =   43907
         End
      End
      Begin MSMask.MaskEdBox txtDataHoje 
         Height          =   495
         Left            =   7440
         TabIndex        =   27
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   16777194
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Data Hoje"
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
         Left            =   7560
         TabIndex        =   28
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pagamento(s) do dia pendente(s) de Confirmação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      TabIndex        =   24
      Top             =   960
      Width           =   10935
      Begin MSFlexGridLib.MSFlexGrid GridPendentes 
         Height          =   3255
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         BackColor       =   16777152
         BackColorFixed  =   16777152
         BackColorSel    =   16777152
         BackColorBkg    =   16777152
         FormatString    =   "||Colaborador      |Docto.         |Forma de Pagto.   |Descrição          |Valor           |Sit.|||"
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
   End
   Begin VB.Frame Frame4 
      Caption         =   "Pagamento(s) confirmado(s) no dia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   23
      Top             =   4560
      Width           =   12975
      Begin MSFlexGridLib.MSFlexGrid GridConfirmados 
         Height          =   3135
         Left            =   0
         TabIndex        =   34
         Top             =   480
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         BackColor       =   16777152
         BackColorFixed  =   16776960
         BackColorSel    =   16776960
         BackColorBkg    =   16777152
         FormatString    =   "|Colaborador        |Doc.              |Forma de Pagto.| Desc Operação        |Vencito      |Valor          |Sit.|||"
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
   End
   Begin VB.Frame Frame5 
      Caption         =   "Pagamento(s) em atraso pendente(s) de confirmação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   10920
      TabIndex        =   22
      Top             =   960
      Width           =   9375
      Begin MSFlexGridLib.MSFlexGrid GridAtrasados 
         Height          =   3255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColor       =   16777152
         BackColorFixed  =   16776960
         BackColorSel    =   16776960
         BackColorBkg    =   16777152
         FormatString    =   "|Colaborador      |Doc.       |Desc Operação   |Vencito   |Valor       |Sit.|||"
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
   End
   Begin VB.Frame Frame6 
      Caption         =   "Resumo Financeiro"
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
      Left            =   12960
      TabIndex        =   15
      Top             =   5760
      Width           =   7335
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Pendente de Confirmação......."
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
         TabIndex        =   21
         Top             =   480
         Width           =   4380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total em Atraso na data...................."
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
         TabIndex        =   20
         Top             =   960
         Width           =   4470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Pago no Dia.............................."
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
         TabIndex        =   19
         Top             =   1440
         Width           =   4470
      End
      Begin VB.Label lblTotalPendente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFEA&
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
         Left            =   4920
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblTotalAtrasado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFEA&
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
         Left            =   4920
         TabIndex        =   17
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblTotalRecebido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFEA&
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
         Left            =   4920
         TabIndex        =   16
         Top             =   1440
         Width           =   1695
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Controle Operacional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   12960
      TabIndex        =   12
      Top             =   8280
      Width           =   7335
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   480
         TabIndex        =   14
         Top             =   720
         Width           =   2295
         Begin VB.CommandButton cmdConfirma 
            BackColor       =   &H008080FF&
            Caption         =   "Confirma"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   4320
         TabIndex        =   13
         Top             =   720
         Width           =   2415
         Begin VB.CommandButton cmdSair 
            BackColor       =   &H00FFFFC0&
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmControlePagamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Linha As Single
Dim coluna As Single
Dim Ind As Single
Dim IndPend As Byte
Dim IndAtra As Byte
Dim IndConf As Byte
Dim IndPendlimite As Byte
Dim IndAtraLimite As Byte
Dim IndConfLimite As Byte
Dim IndParm As Byte
Dim AcumulaPendentes As Currency
Dim AcumulaAtrasados As Currency
Dim AcumulaConfirmados As Currency
Dim AcumulaConsolidado As Currency
Dim DataWS As Date
Dim Dia As Integer
Dim Resp As String
Dim DataConsulta As Date
Dim DataAnteriorHoje As Date
Dim DataInvertida As Double
Dim ColunaZ As Integer
Dim Parametro As String
Dim Encontrei As Byte

Private Sub cmdConfirma_Click()
Dim A As String
Dim B As String
Dim C As String
Dim D As String
Dim E As String

'Resp = MsgBox("Documentos selecionados com OK serão Compensados/Liquidados com a data " & txtDataComp & ". Confirma???", vbYesNo)
'If Resp = vbNo Then
'   'MsgBox ("Altere a DATA COMPENSAÇÃO e dê OK somente para os documentos Compensados/Liquidados na data informada")
'   cmdSair.SetFocus
'   Exit Sub
'End If

Call Rotina_AbrirBanco


For IndConf = 1 To IndPendlimite
    If GridPendentes.TextMatrix(IndConf, 7) = "Ok." Then
       A = 0
       B = GridPendentes.TextMatrix(IndConf, 2)
       C = GridPendentes.TextMatrix(IndConf, 1)
       D = GridPendentes.TextMatrix(IndConf, 3)
       E = GridPendentes.TextMatrix(IndConf, 0)
 '      If A = Empty Then
 '         MsgBox ("Clicar com o mouse apenas em linhas com conteúdo")
 '         GridPendentes.TextMatrix(IndConf, 7) = Empty
 '         cmdSair.SetFocus
 '         Exit Sub
 '      End If
        
       ctp.Open "Select * from contas_a_pagar where chFabricante = ('" & A & "') and chPessoa = ('" & B & "') and chNotaFiscal = ('" & C & "') and chFatura = ('" & D & "')", db, 3, 3
       
       If ctp.EOF Then
          MsgBox ("Erro no acesso para atualizacao de Conta a Pagar 1."), vbCritical
          Call FechaDB
          Exit Sub
       End If
       
       ctp!ctpStatus = 1
       ctp!ctpDataPagamento = txtDataComp
       ctp!ctpDataProc = Date
       
       ctp.Update
              
       DataInvertida = Year(ctp!chDataVencito) & Format$(Month(ctp!chDataVencito), "00") & Format$(Day(ctp!chDataVencito), "00")
       nfd.Open "Select * from notafiscaldesdobramento where chPessoa = ('" & ctp!chPessoa & "') and chNotaFiscalEntrada = ('" & ctp!chNotafiscal & "') and chDataVencimento = ('" & DataInvertida & "')", db, 3, 3
       If nfd.EOF Then
          IndConf = IndConf
       Else
          nfd!nfdStatusPagto = 1
          nfd!nfdDataPagamento = Date
          nfd.Update
       End If
    Else
       If GridPendentes.TextMatrix(IndConf, 0) = Empty Then
          IndConf = 250
       End If
    End If

If ctp.State = 1 Then
   ctp.Close: Set ctp = Nothing
End If
If nfd.State = 1 Then
   nfd.Close: Set nfd = Nothing
End If

Next

For IndConf = 1 To IndAtraLimite
    If GridAtrasados.TextMatrix(IndConf, 6) = "Ok." Then
       A = 0
       B = GridAtrasados.TextMatrix(IndConf, 1)
       C = GridAtrasados.TextMatrix(IndConf, 7)
       D = GridAtrasados.TextMatrix(IndConf, 0)
       E = GridAtrasados.TextMatrix(IndConf, 4)
       If A = Empty Then
          MsgBox ("Clicar com o mouse apenas em linhas com conteúdo")
          GridAtrasados.TextMatrix(IndConf, 6) = Empty
          cmdSair.SetFocus
          Exit Sub
       End If
       
       ctp.Open "Select * from contas_a_pagar where chFabricante = ('" & A & "') and chPessoa = ('" & B & "') and chNotaFiscal = ('" & C & "') and chFatura = ('" & D & "')", db, 3, 3
       
       If ctp.EOF Then
          MsgBox ("Erro no acesso para atualizacao de Conta a Pagar 1."), vbCritical
          Call FechaDB
          Exit Sub
       End If
       ctp!ctpStatus = 1
       ctp!ctpDataPagamento = txtDataComp
       ctp!ctpDataProc = Date
       
       DataInvertida = Year(ctp!chDataVencito) & Format$(Month(ctp!chDataVencito), "00") & Format$(Day(ctp!chDataVencito), "00")
       
       nfd.Open "Select * from notafiscaldesdobramento where chPessoa = ('" & ctp!chPessoa & "') and chNotaFiscalEntrada = ('" & ctp!chNotafiscal & "') and chDataVencimento = ('" & DataInvertida & "')", db, 3, 3
       If Not nfd.EOF Then
          nfd!nfdStatusPagto = 1
          nfd!nfdDataPagamento = Date
          nfd.Update
       End If
             
       ctp.Update
    Else
       If GridAtrasados.TextMatrix(IndConf, 0) = Empty Then
          IndConf = 250
       End If
    End If

If ctp.State = 1 Then
   ctp.Close: Set ctp = Nothing
End If
If nfd.State = 1 Then
   nfd.Close: Set nfd = Nothing
End If


Next
For IndConf = 1 To IndConfLimite

    If GridConfirmados.TextMatrix(IndConf, 7) = "Ok." Then
       A = 0
       B = GridConfirmados.TextMatrix(IndConf, 1)
       C = GridConfirmados.TextMatrix(IndConf, 8)
       D = GridConfirmados.TextMatrix(IndConf, 0)
       E = GridConfirmados.TextMatrix(IndConf, 5)
       If A = Empty Then
          MsgBox ("Clicar com o mouse apenas em linhas com conteúdo")
          GridConfirmados.TextMatrix(IndConf, 7) = Empty
          cmdSair.SetFocus
          Exit Sub
       End If
       
       ctp.Open "Select * from contas_a_pagar where chFabricante = ('" & A & "') and chPessoa = ('" & B & "') and chNotaFiscal = ('" & C & "') and chFatura = ('" & D & "')", db, 3, 3
       
       If ctp.EOF Then
          MsgBox ("Erro no acesso para atualizacao de Conta a Pagar 2."), vbCritical
          Call FechaDB
          Exit Sub
       End If
       ctp!ctpStatus = 0
       ctp!ctpDataPagamento = Empty
       ctp!ctpDataProc = Empty
       
       DataInvertida = Year(ctp!chDataVencito) & Format$(Month(ctp!chDataVencito), "00") & Format$(Day(ctp!chDataVencito), "00")
       
       nfd.Open "Select * from notafiscaldesdobramento where chPessoa = ('" & ctp!chPessoa & "') and chNotaFiscalEntrada = ('" & ctp!chNotafiscal & "') and chDataVencimento = ('" & DataInvertida & "')", db, 3, 3
       If Not nfd.EOF Then
          nfd!nfdStatusPagto = 0
          nfd.Update
       End If
             
       ctp.Update
    Else
       If GridConfirmados.TextMatrix(IndConf, 0) = Empty Then
          IndConf = 250
       End If
    End If

If ctp.State = 1 Then
   ctp.Close: Set ctp = Nothing
End If
If nfd.State = 1 Then
   nfd.Close: Set nfd = Nothing
End If


Next
AcumulaPendentes = 0
AcumulaAtrasados = 0
AcumulaConfirmados = 0

Call FechaDB

Call Rotina_010_Limpa_Pendentes

Call Rotina_011_Limpa_Atrasados

Call Rotina_012_Limpa_Confirmados

Call Rotina_020_Gerencia_Grid

End Sub


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Command1_Click()

txtDataHoje = Date
DataInformada = txtDataConsulta
DataRetorno = txtDataConsulta
NDias = 1

'DataRetorno = ObterDiaUtilAnterior(DataInformada, NDias)

AcumulaPendentes = 0
AcumulaAtrasados = 0
AcumulaConfirmados = 0

Call Rotina_010_Limpa_Pendentes

Call Rotina_011_Limpa_Atrasados

Call Rotina_012_Limpa_Confirmados

Call Rotina_020_Gerencia_Grid

End Sub

Private Sub Form_Load()
Dim Hoje As Date


txtDataComp = Date
Hoje = Data_Hoje
NDias = 1
txtDataHoje = Date
DataInformada = Date


DataRetorno = Date

DataAnteriorHoje = Date - 1
txtDataConsulta = Date

AcumulaPendentes = 0
AcumulaAtrasados = 0
AcumulaConfirmados = 0

cmbFiltro.AddItem "Geral"

cmbTipopagto.AddItem "do Dia"
cmbTipopagto.AddItem "em atraso"
cmbTipopagto.AddItem "Pagas"

cmbConsolidado.AddItem Empty

optSim = False
optNao = False

Call Rotina_AbrirBanco

Bco.Open "Select * from banco", db, 3, 3
If Bco.EOF Then
   MsgBox ("Tabela de bancos está vazia."), vbCritical
   Call FechaDB
   Exit Sub
End If
 

Bco.MoveFirst

Do While Not Bco.EOF
   cmbFiltro.AddItem Bco!bcosiglabco
   Bco.MoveNext
Loop

cmbFiltro.ListIndex = 0

Call Rotina_010_Limpa_Pendentes

Call Rotina_011_Limpa_Atrasados

Call Rotina_012_Limpa_Confirmados

Call Rotina_020_Gerencia_Grid

End Sub

Public Sub Rotina_010_Limpa_Pendentes()

GridPendentes.Rows = 2
IndPend = 1
GridPendentes.TextMatrix(IndPend, 0) = Empty
GridPendentes.TextMatrix(IndPend, 1) = Empty
GridPendentes.TextMatrix(IndPend, 2) = Empty
GridPendentes.TextMatrix(IndPend, 3) = Empty
GridPendentes.TextMatrix(IndPend, 4) = Empty
GridPendentes.TextMatrix(IndPend, 5) = Empty
GridPendentes.TextMatrix(IndPend, 6) = Empty
GridPendentes.TextMatrix(IndPend, 7) = Empty
GridPendentes.TextMatrix(IndPend, 8) = Empty
GridPendentes.TextMatrix(IndPend, 9) = Empty


End Sub
Public Sub Rotina_011_Limpa_Atrasados()
GridAtrasados.Rows = 2
IndAtra = 1
GridAtrasados.TextMatrix(IndAtra, 0) = Empty
GridAtrasados.TextMatrix(IndAtra, 1) = Empty
GridAtrasados.TextMatrix(IndAtra, 2) = Empty
GridAtrasados.TextMatrix(IndAtra, 3) = Empty
GridAtrasados.TextMatrix(IndAtra, 4) = Empty
GridAtrasados.TextMatrix(IndAtra, 5) = Empty
GridAtrasados.TextMatrix(IndAtra, 6) = Empty
GridAtrasados.TextMatrix(IndAtra, 7) = Empty
GridAtrasados.TextMatrix(IndAtra, 8) = Empty
GridAtrasados.TextMatrix(IndAtra, 9) = Empty

End Sub
Public Sub Rotina_012_Limpa_Confirmados()
GridConfirmados.Rows = 2
IndConf = 1
GridConfirmados.TextMatrix(IndConf, 0) = Empty
GridConfirmados.TextMatrix(IndConf, 1) = Empty
GridConfirmados.TextMatrix(IndConf, 2) = Empty
GridConfirmados.TextMatrix(IndConf, 3) = Empty
GridConfirmados.TextMatrix(IndConf, 4) = Empty
GridConfirmados.TextMatrix(IndConf, 5) = Empty
GridConfirmados.TextMatrix(IndConf, 6) = Empty
GridConfirmados.TextMatrix(IndConf, 7) = Empty
GridConfirmados.TextMatrix(IndConf, 8) = Empty
GridConfirmados.TextMatrix(IndConf, 9) = Empty
GridConfirmados.TextMatrix(IndConf, 10) = Empty

End Sub

Public Sub Rotina_020_Gerencia_Grid()

Dim indice As Byte
Dim DataAcesso As Date

IndPend = 0
IndAtra = 0
IndConf = 0

DataWS = txtDataConsulta

Call Rotina_AbrirBanco

ctp.Open "Select * from contas_a_pagar", db, 3, 3
If ctp.EOF Then
   MsgBox ("Não há contas a pagar. "), vbInformation
   Call FechaDB
   Exit Sub
End If

ctp.MoveFirst

DataAcesso = Date
 
Do While Not ctp.EOF
   
   indice = cmbFiltro.ListIndex
   
   If (cmbFiltro = "Geral") Or (cmbFiltro = ctp!chCodBcoLart) Then
      If ctp!chDataVencito = txtDataConsulta Then
         If ctp!ctpStatus = 0 Then
            Call Rotina_050_Carga_Pendentes
         End If
      End If
  
      If ctp!ctpDataPagamento = Date Then
         If ctp!ctpStatus = 1 Then
            Call Rotina_051_Carga_Confirmados
         End If
      End If
     
      If ctp!ctpStatus = 0 Then
         If ctp!chDataVencito < DataAnteriorHoje + 1 Then
            Call Rotina_052_Carga_Atrasados
         End If
      End If
   End If
   
   ctp.MoveNext

Loop

GridPendentes.Col = 7
GridPendentes.ColSel = 7
    
GridPendentes.Row = 1
GridPendentes.RowSel = IndPend
       
If IndPend > 1 Then
   GridPendentes.Sort = 5
End If

GridPendentes.Row = IndPend

lblTotalPendente = Format$(AcumulaPendentes, "##,##0.00")
lblTotalAtrasado = Format$(AcumulaAtrasados, "##,##0.00")
lblTotalRecebido = Format$(AcumulaConfirmados, "##,##0.00")

End Sub

Public Sub Rotina_050_Carga_Pendentes()

IndPend = IndPend + 1
GridPendentes.Rows = IndPend + 1

DataInvertida = Year(ctp!chDataVencito) & Format$(Month(ctp!chDataVencito), "00") & Format$(Day(ctp!chDataVencito), "00")

GridPendentes.TextMatrix(IndPend, 0) = ctp!chDataVencito
GridPendentes.TextMatrix(IndPend, 1) = ctp!chNotafiscal
GridPendentes.TextMatrix(IndPend, 2) = ctp!chPessoa
GridPendentes.TextMatrix(IndPend, 3) = ctp!chFatura
GridPendentes.TextMatrix(IndPend, 5) = ctp!ctpdescricaooperacao
GridPendentes.TextMatrix(IndPend, 6) = Format$(ctp!ctpValorDaBoleta, "##,##0.00")
GridPendentes.TextMatrix(IndPend, 7) = Empty
GridPendentes.TextMatrix(IndPend, 8) = DataInvertida & ctp!ctpdescricaooperacao & ctp!chPessoa & ctp!chFatura
GridPendentes.TextMatrix(IndPend, 9) = Format$(ctp!ctpValorDaBoleta, "####0.00")
IndPendlimite = IndPend
AcumulaPendentes = Format$(AcumulaPendentes + ctp!ctpValorDaBoleta, "##,##0.00")
If ctp!ctpTipoLancamentoDesc = "BOLETO" Then
   If nfd.State = 1 Then
      nfd.Close: Set nfd = Nothing
   End If
   nfd.Open "Select * from notafiscaldesdobramento where chPessoa = ('" & ctp!chPessoa & "') and chNotaFiscalEntrada = ('" & ctp!chNotafiscal & "') and nfdFaturaNumero = ('" & ctp!chFatura & "')", db, 3, 3
   If Not nfd.EOF Then
      If Not nfd!nfdIPTE = "" Then
         GridPendentes.TextMatrix(IndPend, 4) = (ctp!ctpTipoLancamentoDesc & " C Barra")
      Else
         GridPendentes.TextMatrix(IndPend, 4) = ctp!ctpTipoLancamentoDesc
      End If
   Else
      If nfd.State = 1 Then
         nfd.Close: Set nfd = Nothing
      End If
      nfd.Open "Select * from historiconotafiscaldesdobramento where chPessoa = ('" & ctp!chPessoa & "') and chNotaFiscalEntrada = ('" & ctp!chNotafiscal & "') and nfdFaturaNumero = ('" & ctp!chFatura & "')", db, 3, 3
      If Not nfd.EOF Then
         If Not nfd!nfdIPTE = "" Then
            GridPendentes.TextMatrix(IndPend, 4) = (ctp!ctpTipoLancamentoDesc & " C Barra")
         Else
            GridPendentes.TextMatrix(IndPend, 4) = ctp!ctpTipoLancamentoDesc
         End If
      End If
   End If
Else
   GridPendentes.TextMatrix(IndPend, 4) = ctp!ctpTipoLancamentoDesc
End If
End Sub
Public Sub Rotina_051_Carga_Confirmados()

IndConf = IndConf + 1
GridConfirmados.Rows = IndConf + 1
DataInvertida = Year(ctp!chDataVencito) & Format$(Month(ctp!chDataVencito), "00") & Format$(Day(ctp!chDataVencito), "00")
GridConfirmados.TextMatrix(IndConf, 0) = ctp!chFatura
GridConfirmados.TextMatrix(IndConf, 1) = ctp!chPessoa
GridConfirmados.TextMatrix(IndConf, 2) = ctp!chFatura
GridConfirmados.TextMatrix(IndConf, 3) = ctp!ctpTipoLancamentoDesc
GridConfirmados.TextMatrix(IndConf, 4) = ctp!ctpdescricaooperacao
GridConfirmados.TextMatrix(IndConf, 5) = Format$(ctp!chDataVencito, "dd/mm/yy")
GridConfirmados.TextMatrix(IndConf, 6) = Format$(ctp!ctpValorDaBoleta, "##,##0.00")
GridConfirmados.TextMatrix(IndConf, 7) = Empty
GridConfirmados.TextMatrix(IndConf, 8) = ctp!chNotafiscal
GridConfirmados.TextMatrix(IndConf, 9) = DataInvertida
GridConfirmados.TextMatrix(IndConf, 10) = Format$(ctp!ctpValorDaBoleta, "####0.00")
IndConfLimite = IndConf
AcumulaConfirmados = AcumulaConfirmados + ctp!ctpValorDaBoleta

End Sub
Public Sub Rotina_052_Carga_Atrasados()

IndAtra = IndAtra + 1
GridAtrasados.Rows = IndAtra + 1

DataInvertida = Year(ctp!chDataVencito) & Format$(Month(ctp!chDataVencito), "00") & Format$(Day(ctp!chDataVencito), "00")

GridAtrasados.TextMatrix(IndAtra, 0) = ctp!chFatura
GridAtrasados.TextMatrix(IndAtra, 1) = ctp!chPessoa
GridAtrasados.TextMatrix(IndAtra, 2) = ctp!chFatura
GridAtrasados.TextMatrix(IndAtra, 3) = ctp!ctpdescricaooperacao
GridAtrasados.TextMatrix(IndAtra, 4) = Format$(ctp!chDataVencito, "dd/mm/yy")
GridAtrasados.TextMatrix(IndAtra, 5) = Format$(ctp!ctpValorDaBoleta, "##,##0.00")
GridAtrasados.TextMatrix(IndAtra, 6) = Empty
GridAtrasados.TextMatrix(IndAtra, 7) = ctp!chNotafiscal
GridAtrasados.TextMatrix(IndAtra, 8) = DataInvertida
GridAtrasados.TextMatrix(IndAtra, 9) = Format$(ctp!ctpValorDaBoleta, "####0.00")

IndAtraLimite = IndAtra
AcumulaAtrasados = AcumulaAtrasados + ctp!ctpValorDaBoleta

End Sub



Private Sub GridAtrasados_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'If optClassif = True Then
'   Call ClassificaAtrasados
'   Exit Sub
'End If

If GridAtrasados.TextMatrix(GridAtrasados.Row, 6) = "Ok." Then
   GridAtrasados.TextMatrix(GridAtrasados.Row, 6) = Empty
Else
   GridAtrasados.TextMatrix(GridAtrasados.Row, 6) = "Ok."
End If
End Sub

Private Sub GridConfirmados_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'If optClassif = True Then
'   Call ClassificaConfirmados
'   Exit Sub
'End If

If GridConfirmados.TextMatrix(GridConfirmados.Row, 7) = "Ok." Then
   GridConfirmados.TextMatrix(GridConfirmados.Row, 7) = Empty
Else
   GridConfirmados.TextMatrix(GridConfirmados.Row, 7) = "Ok."
End If
End Sub

Private Sub GridPendentes_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'If optClassif = True Then
'   Call ClassificaPendentes
'   Exit Sub
'End If

Call Rotina_AbrirBanco

If Not GridPendentes.TextMatrix(GridPendentes.Row, 7) = "Ok." Then
    If GridPendentes.TextMatrix(GridPendentes.Row, 4) = "BOLETO C Barra" Then
       frmMostraIPTE.txtColaborador = GridPendentes.TextMatrix(GridPendentes.Row, 2)
       frmMostraIPTE.txtDocumento = GridPendentes.TextMatrix(GridPendentes.Row, 1)
       frmMostraIPTE.txtIPTE = GridPendentes.TextMatrix(GridPendentes.Row, 3)
       If nfd.State = 1 Then
          nfd.Close: Set nfd = Nothing
       End If

       nfd.Open "Select * from notafiscaldesdobramento where chPessoa = ('" & GridPendentes.TextMatrix(GridPendentes.Row, 2) & "') and chNotaFiscalEntrada = ('" & GridPendentes.TextMatrix(GridPendentes.Row, 1) & "') and nfdfaturanumero = ('" & GridPendentes.TextMatrix(GridPendentes.Row, 3) & "')", db, 3, 3
       If nfd.EOF Then
          If nfd.State = 1 Then
             nfd.Close: Set nfd = Nothing
          End If
          nfd.Open "Select * from historiconotafiscaldesdobramento where chPessoa = ('" & GridPendentes.TextMatrix(GridPendentes.Row, 2) & "') and chNotaFiscalEntrada = ('" & GridPendentes.TextMatrix(GridPendentes.Row, 1) & "') and nfdfaturanumero = ('" & GridPendentes.TextMatrix(GridPendentes.Row, 3) & "')", db, 3, 3
          If nfd.EOF Then
             MsgBox ("Desdobramento não encontrado."), vbCritical
             Call FechaDB
             Exit Sub
          End If
       End If
       frmMostraIPTE.txtIPTE = nfd!nfdIPTE
       frmMostraIPTE.Show vbModal
    End If
End If

If GridPendentes.TextMatrix(GridPendentes.Row, 7) = "Ok." Then
   GridPendentes.TextMatrix(GridPendentes.Row, 7) = Empty
Else
   GridPendentes.TextMatrix(GridPendentes.Row, 7) = "Ok."
End If

Call FechaDB

End Sub

Public Sub ClassificaAtrasados()
Linha = GridAtrasados.Rows
coluna = GridAtrasados.Col

If coluna = 4 Then
   coluna = 8
Else
   If coluna = 5 Then
      coluna = 9
   End If
End If

GridAtrasados.Col = coluna
GridAtrasados.ColSel = coluna
     
GridAtrasados.Row = 1
GridAtrasados.RowSel = Linha - 1

GridAtrasados.Sort = 1

GridAtrasados.Col = 0
GridAtrasados.ColSel = 0
GridAtrasados.Row = 0
GridAtrasados.RowSel = 0
End Sub

Public Sub ClassificaPendentes()
Linha = GridPendentes.Rows
coluna = GridPendentes.Col

If coluna = 5 Then
   coluna = 8
End If

GridPendentes.Col = coluna
GridPendentes.ColSel = coluna
     
GridPendentes.Row = 1
GridPendentes.RowSel = Linha - 1

GridPendentes.Sort = 1

GridPendentes.Col = 0
GridPendentes.ColSel = 0
GridPendentes.Row = 0
GridPendentes.RowSel = 0
End Sub

Public Sub ClassificaConfirmados()
Linha = GridConfirmados.Rows
coluna = GridConfirmados.Col

If coluna = 4 Then
   coluna = 8
Else
   If coluna = 5 Then
      coluna = 9
   End If
End If

GridConfirmados.Col = coluna
GridConfirmados.ColSel = coluna
     
GridConfirmados.Row = 1
GridConfirmados.RowSel = Linha - 1

GridConfirmados.Sort = 1

GridConfirmados.Col = 0
GridConfirmados.ColSel = 0
GridConfirmados.Row = 0
GridConfirmados.RowSel = 0
End Sub

Private Sub cmbTipoPagto_LostFocus()

cmbTipoConsolidacao.Clear

If cmbTipopagto = "em atraso" Then
   cmbTipoConsolidacao.AddItem "Colaborador"
   cmbTipoConsolidacao.AddItem "Desc Operação"
Else
   cmbTipoConsolidacao.AddItem "Colaborador"
   cmbTipoConsolidacao.AddItem "Forma de Pagto"
   cmbTipoConsolidacao.AddItem "Desc Operação"
End If

End Sub
Private Sub cmbTipoConsolidacao_LostFocus()

cmbConsolidado.Clear

If cmbTipoConsolidacao = "Colaborador" Then
   lblConsolidado.Caption = "Colaborador"
   If cmbTipopagto = "do Dia" Then
      ColunaZ = 2
   Else
      ColunaZ = 1
   End If
Else
   If cmbTipoConsolidacao = "Forma de Pagto" Then
      lblConsolidado = "Forma de Pagto"
      If cmbTipopagto = "do Dia" Then
         ColunaZ = 4
      Else
         ColunaZ = 3
      End If
   Else
      If cmbTipoConsolidacao = "Desc Operação" Then
         lblConsolidado = "Desc Operação"
         If cmbTipopagto = "do Dia" Then
            ColunaZ = 5
         Else
            If cmbTipopagto = "em atraso" Then
               ColunaZ = 3
            Else
               ColunaZ = 4
            End If
         End If
      Else
         MsgBox "Erro na informação do tipo de consolidado desejado|", vbCritical
         Exit Sub
      End If
   End If
End If

If cmbTipopagto = "do Dia" Then
   For Ind = 1 To IndPendlimite
       For IndParm = 1 To Ind
           If GridPendentes.TextMatrix(Ind, ColunaZ) = GridPendentes.TextMatrix(IndParm - 1, ColunaZ) Then
              IndParm = Ind
              Encontrei = 1
           Else
              Encontrei = 0
           End If
       Next
       If Not Encontrei = 1 Then
          cmbConsolidado.AddItem GridPendentes.TextMatrix(Ind, ColunaZ)
          Encontrei = 0
       Else
          Encontrei = 0
       End If
   Next
Else
   If cmbTipopagto = "em atraso" Then
      For Ind = 1 To IndAtraLimite
          For IndParm = 1 To Ind
           If GridAtrasados.TextMatrix(Ind, ColunaZ) = GridAtrasados.TextMatrix(IndParm - 1, ColunaZ) Then
              IndParm = Ind
              Encontrei = 1
           Else
              Encontrei = 0
           End If
       Next
       If Not Encontrei = 1 Then
          cmbConsolidado.AddItem GridAtrasados.TextMatrix(Ind, ColunaZ)
          Encontrei = 0
       Else
          Encontrei = 0
       End If
      Next
   Else
      If cmbTipopagto = "Pagas" Then
         For Ind = 1 To IndConfLimite
             For IndParm = 1 To Ind
           If GridConfirmados.TextMatrix(Ind, ColunaZ) = GridConfirmados.TextMatrix(IndParm - 1, ColunaZ) Then
              IndParm = Ind
              Encontrei = 1
           Else
              Encontrei = 0
           End If
       Next
       If Not Encontrei = 1 Then
          cmbConsolidado.AddItem GridConfirmados.TextMatrix(Ind, ColunaZ)
          Encontrei = 0
       Else
          Encontrei = 0
       End If
         Next
      End If
    End If
 End If
End Sub

Private Sub cmbConsolidado_LostFocus()

AcumulaConsolidado = 0

If cmbTipopagto = "do Dia" Then
   For Ind = 1 To IndPendlimite
       If cmbConsolidado = GridPendentes.TextMatrix(Ind, ColunaZ) Then
          AcumulaConsolidado = AcumulaConsolidado + Format$(GridPendentes.TextMatrix(Ind, 6), "###,##0.00")
       End If
   Next
   txtValorTotalConsolidado = Format$(AcumulaConsolidado, "###,##0.00")
End If

If cmbTipopagto = "em atraso" Then
   For Ind = 1 To IndAtraLimite
       If cmbConsolidado = GridAtrasados.TextMatrix(Ind, ColunaZ) Then
          AcumulaConsolidado = AcumulaConsolidado + Format$(GridAtrasados.TextMatrix(Ind, 5), "###,##0.00")
       End If
   Next
   txtValorTotalConsolidado = Format$(AcumulaConsolidado, "###,##0.00")
End If

If cmbTipopagto = "Pagas" Then
   For Ind = 1 To IndConfLimite
       If cmbConsolidado = GridConfirmados.TextMatrix(Ind, ColunaZ) Then
          AcumulaConsolidado = AcumulaConsolidado + Format$(GridConfirmados.TextMatrix(Ind, 6), "###,##0.00")
       End If
   Next
   txtValorTotalConsolidado = Format$(AcumulaConsolidado, "###,##0.00")
End If
       
End Sub

Private Sub optNao_Click()
optSim = False
If cmbTipopagto = "do Dia" Then
   For Ind = 1 To IndPendlimite
       If cmbConsolidado = GridPendentes.TextMatrix(Ind, ColunaZ) Then
          GridPendentes.TextMatrix(Ind, 7) = Empty
       End If
   Next
End If

If cmbTipopagto = "em atraso" Then
   For Ind = 1 To IndAtraLimite
       If cmbConsolidado = GridAtrasados.TextMatrix(Ind, ColunaZ) Then
          GridAtrasados.TextMatrix(Ind, 6) = Empty
       End If
   Next
End If

If cmbTipopagto = "Pagas" Then
   For Ind = 1 To IndConfLimite
       If cmbConsolidado = GridConfirmados.TextMatrix(Ind, ColunaZ) Then
          GridConfirmados.TextMatrix(Ind, 7) = Empty
       End If
   Next
End If

optSim = False
optNao = False

End Sub

Private Sub optSim_Click()
optNao = False
If cmbTipopagto = "do Dia" Then
   For Ind = 1 To IndPendlimite
       If cmbConsolidado = GridPendentes.TextMatrix(Ind, ColunaZ) Then
          GridPendentes.TextMatrix(Ind, 7) = "Ok."
       End If
   Next
End If

If cmbTipopagto = "em atraso" Then
   For Ind = 1 To IndAtraLimite
       If cmbConsolidado = GridAtrasados.TextMatrix(Ind, ColunaZ) Then
          GridAtrasados.TextMatrix(Ind, 6) = "Ok."
       End If
   Next
End If

If cmbTipopagto = "Pagas" Then
   For Ind = 1 To IndConfLimite
       If cmbConsolidado = GridConfirmados.TextMatrix(Ind, ColunaZ) Then
          GridConfirmados.TextMatrix(Ind, 7) = "Ok."
       End If
   Next
End If
       
optSim = False
optNao = False
       
End Sub
