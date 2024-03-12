VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmControleFinanceiro 
   Caption         =   "frmControleFinanceiro"
   ClientHeight    =   8175
   ClientLeft      =   4110
   ClientTop       =   3690
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   8175
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   8655
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Controle de Recebimentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   24
         Top             =   285
         Width           =   8385
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Controle Operacional"
      Height          =   1455
      Left            =   11880
      TabIndex        =   8
      Top             =   6600
      Width           =   8055
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4800
         TabIndex        =   16
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton cmdSair 
            BackColor       =   &H00FFFF80&
            Caption         =   "Sair"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   720
         TabIndex        =   15
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton cmdConfirma 
            BackColor       =   &H008080FF&
            Caption         =   "Confirma"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   240
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Resumo Financeiro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   11880
      TabIndex        =   7
      Top             =   4560
      Width           =   8055
      Begin VB.Label txtStatusCancela 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFEA&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4680
         TabIndex        =   22
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTotalRecebido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFEA&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6120
         TabIndex        =   14
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblTotalAtrasado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFEA&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6120
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblTotalPendente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFEA&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6120
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Recebido no Dia....................."
         Height          =   360
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   5400
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total em Atraso................................."
         Height          =   360
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Pendente de Confirmação......"
         Height          =   360
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   5385
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Recebimento(s) em atraso pendente(s) de confirmação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   8880
      TabIndex        =   6
      Top             =   960
      Width           =   11535
      Begin MSFlexGridLib.MSFlexGrid GridAtrasados 
         Height          =   3255
         Left            =   0
         TabIndex        =   28
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColor       =   16777194
         BackColorFixed  =   16776960
         BackColorBkg    =   16777194
         FormatString    =   "||||Cliente                     |Desc Operação        |Vencito        |Valor         |Total            |Status"
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
      Caption         =   "Recebimento(s) confirmado(s) no dia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   10695
      Begin MSFlexGridLib.MSFlexGrid GridConfirmados 
         Height          =   3135
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColor       =   16777194
         BackColorFixed  =   16776960
         BackColorBkg    =   16777194
         FormatString    =   "||||Cliente                          |Desc Operação        |Vencito        |Valor         |Status"
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
   Begin VB.Frame Frame3 
      Caption         =   "Recebimento(s) do dia pendente(s) de Confirmação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   8655
      Begin MSFlexGridLib.MSFlexGrid GridPendentes 
         Height          =   3255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColor       =   16777194
         BackColorFixed  =   16776960
         BackColorBkg    =   16777194
         FormatString    =   "||||Cliente                      |Descrição               |Valor           |Status"
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
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8880
      TabIndex        =   3
      Top             =   0
      Width           =   11415
      Begin VB.Frame Frame10 
         Caption         =   "Filtro"
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
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   2295
         Begin VB.ComboBox cmbFiltro 
            BackColor       =   &H00FFFFEA&
            Height          =   420
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Data Consulta"
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
         Left            =   2640
         TabIndex        =   25
         Top             =   120
         Width           =   2175
         Begin MSComCtl2.DTPicker txtDataInformada 
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
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
            Format          =   378339329
            CurrentDate     =   38125
         End
      End
      Begin VB.CommandButton cmdConsulta 
         BackColor       =   &H00FFFF00&
         Caption         =   "Consulta"
         Height          =   735
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txtDataConsulta 
         Height          =   375
         Left            =   9480
         TabIndex        =   19
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox txtDataHoje 
         Height          =   375
         Left            =   9480
         TabIndex        =   18
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acesso"
         Height          =   375
         Left            =   8400
         TabIndex        =   20
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hoje"
         Height          =   360
         Left            =   8400
         TabIndex        =   17
         Top             =   120
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmControleFinanceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IndPend As Byte
Dim IndAtra As Byte
Dim IndConf As Byte
Dim IndPendlimite As Byte
Dim IndAtraLimite As Byte
Dim IndConfLimite As Byte
Dim AcumulaPendentes As Currency
Dim AcumulaAtrasados As Currency
Dim AcumulaConfirmados As Currency
Dim SomaAtrasados As Currency
Dim Cliente As String
Dim pessoaAnterior As String

Dim Dia As Integer
Dim DiaUtilAnterior As Date
Dim dataInicio As Date
Dim dataFim As Date
Dim DataPendentes As Date
Dim D2Anterior As Date
Dim DataPagos As Date
Dim DataAtrasados As Date
Dim DataParaAjuste As Date


Private Sub cmdConfirma_Click()
Dim A As String
Dim B As String
Dim C As String
Dim D As String

Call Rotina_AbrirBanco

For IndConf = 1 To IndPendlimite
    If GridPendentes.TextMatrix(IndConf, 7) = "Ok." Then
       A = GridPendentes.TextMatrix(IndConf, 0)
       B = GridPendentes.TextMatrix(IndConf, 1)
       C = GridPendentes.TextMatrix(IndConf, 2)
       D = GridPendentes.TextMatrix(IndConf, 3)
       
       ctr.Open "Select* from contas_a_receber where chFabricante = ('" & A & "') and chPessoa = ('" & B & "') and chNotaFiscal = ('" & C & "')and chFatura = ('" & D & "')", db, 3, 3
       If ctr.EOF Then
          MsgBox ("Erro no acesso para atualizacao de Conta a Receber 1."), vbCritical
          Call FechaDB
          Exit Sub
       End If

       ctr!ctrStatus = 1
       ctr!ctrDataRecebimento = Date
       ctr.Update
    Else
       If GridPendentes.TextMatrix(IndConf, 0) = Empty Then
          IndConf = 250
       End If
    End If

If ctr.State = 1 Then
   ctr.Close: Set ctr = Nothing
End If

Next

For IndConf = 1 To IndAtraLimite
    If GridAtrasados.TextMatrix(IndConf, 9) = "Ok." Then
       A = GridAtrasados.TextMatrix(IndConf, 0)
       B = GridAtrasados.TextMatrix(IndConf, 1)
       C = GridAtrasados.TextMatrix(IndConf, 2)
       D = GridAtrasados.TextMatrix(IndConf, 3)
       ctr.Open "Select* from contas_a_receber where chFabricante = ('" & A & "') and chPessoa = ('" & B & "') and chNotaFiscal = ('" & C & "')and chFatura = ('" & D & "')", db, 3, 3
       If ctr.EOF Then
          MsgBox ("Erro no acesso para atualizacao de Conta a Receber 2."), vbCritical
          Call FechaDB
          Exit Sub
       End If
       
       ctr!ctrStatus = 1
       If Not IsDate(ctr!ctrDataRecebimento) Then
          ctr!ctrDataRecebimento = Date
       End If
       ctr.Update
    Else
       If GridAtrasados.TextMatrix(IndConf, 0) = Empty Then
          IndConf = 250
       End If
    End If
If ctr.State = 1 Then
   ctr.Close: Set ctr = Nothing
End If

Next
For IndConf = 1 To IndConfLimite
    If GridConfirmados.TextMatrix(IndConf, 8) = "Ok." Then
       A = GridConfirmados.TextMatrix(IndConf, 0)
       B = GridConfirmados.TextMatrix(IndConf, 1)
       C = GridConfirmados.TextMatrix(IndConf, 2)
       D = GridConfirmados.TextMatrix(IndConf, 3)
       ctr.Open "Select* from contas_a_receber where chFabricante = ('" & A & "') and chPessoa = ('" & B & "') and chNotaFiscal = ('" & C & "')and chFatura = ('" & D & "')", db, 3, 3
       If ctr.EOF Then
          MsgBox ("Erro no acesso para atualizacao de Conta a Receber 3."), vbCritical
          Call FechaDB
          Exit Sub
       End If
       
       ctr!ctrStatus = 0
       ctr!ctrDataRecebimento = Empty
       ctr.Update
    Else
       If GridConfirmados.TextMatrix(IndConf, 0) = Empty Then
          IndConf = 250
       End If
    End If
If ctr.State = 1 Then
   ctr.Close: Set ctr = Nothing
End If

Next
AcumulaPendentes = 0
AcumulaAtrasados = 0
AcumulaConfirmados = 0

Call Rotina_010_Limpa_Pendentes

Call Rotina_011_Limpa_Atrasados

Call Rotina_012_Limpa_Confirmados

Call Rotina_020_Gerencia_Grid

Call FechaDB

End Sub


Private Sub cmdConsulta_Click()

txtDataHoje = Date
DataInformada = txtDataInformada
txtDataConsulta = DataInformada
NDias = 1
'DataRetorno = ObterDiaUtilAnterior(DataInformada, NDias)

'If txtDataInformada > Date Then
'   DataPagos = Date
'Else
'   DataPagos = DataRetorno.DiaUtil
'End If
'txtDataConsulta = DataRetorno.DiaUtil
DataParaAjuste = txtDataInformada

AcumulaPendentes = 0
AcumulaAtrasados = 0
AcumulaConfirmados = 0

Call Rotina_010_Limpa_Pendentes

Call Rotina_011_Limpa_Atrasados

Call Rotina_012_Limpa_Confirmados

Call Rotina_020_Gerencia_Grid

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtDataInformada = Date
txtDataHoje = Date
DataInformada = Date
'txtDMaisMenos = 1

NDias = 1

DiadaSemana = Weekday(Date)

'If DiadaSemana = 1 Then
'   DataComAtraso = Date - 1
'Else
   DataComAtraso = Date
'End If

DataParaAjuste = Date

'Call Rotina_070_Ajusta_Data

DataAtrasados = Date
DataPagos = Date

AcumulaPendentes = 0
AcumulaAtrasados = 0
AcumulaConfirmados = 0

cmbFiltro.AddItem "Geral"

Call Rotina_AbrirBanco

Bco.Open "Select * from banco", db, 3, 3

Bco.MoveFirst

Do While Not Bco.EOF
   cmbFiltro.AddItem Bco!bcosiglabco
   Bco.MoveNext
Loop

cmbFiltro.ListIndex = 0
'cmbFiltro = "Geral"
Call Rotina_010_Limpa_Pendentes

Call Rotina_011_Limpa_Atrasados

Call Rotina_012_Limpa_Confirmados

Call Rotina_020_Gerencia_Grid

Call FechaDB

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

End Sub

Public Sub Rotina_020_Gerencia_Grid()

Dim indice As Byte
Dim DataAcesso As Date

IndPend = 0
IndAtra = 0
IndConf = 0

Call Rotina_AbrirBanco

ctr.Open "Select * from contas_a_receber", db, 3, 3
If ctr.EOF Then
   Call FechaDB
   Exit Sub
End If
   
ctr.MoveFirst


Do While Not ctr.EOF
   
   indice = cmbFiltro.ListIndex
   
   If (cmbFiltro = "Geral") Or (cmbFiltro = ctr!chCodBcoLart) Then
      If ctr!ctrDataVencito = txtDataInformada Then
         If ctr!ctrStatus = 0 Then
            Call Rotina_050_Carga_Pendentes
         End If
      End If
  
      'If TabCtaReceber("ctrdatarecebimento") > Date - 1 Then
      If ctr!ctrDataRecebimento = Date Then
         If ctr!ctrStatus = 1 Then
            Call Rotina_051_Carga_Confirmados
         End If
      End If
     
      If ctr!ctrStatus = 0 Then
         If ctr!ctrDataVencito < DataComAtraso Then
            Call Rotina_052_Carga_Atrasados
         End If
      End If
   End If
   
   ctr.MoveNext

Loop

SomaAtrasados = 0
pessoaAnterior = Empty

Call TotalizaAtrasados

lblTotalPendente = Format$(AcumulaPendentes, "##,##0.00")
lblTotalAtrasado = Format$(AcumulaAtrasados, "##,##0.00")
lblTotalRecebido = Format$(AcumulaConfirmados, "##,##0.00")

Call FechaDB

End Sub

Public Sub Rotina_050_Carga_Pendentes()

IndPend = IndPend + 1
GridPendentes.Rows = IndPend + 1
GridPendentes.TextMatrix(IndPend, 0) = ctr!chFabricante
GridPendentes.TextMatrix(IndPend, 1) = ctr!chPessoa
GridPendentes.TextMatrix(IndPend, 2) = ctr!chNotafiscal
GridPendentes.TextMatrix(IndPend, 3) = ctr!chFatura
GridPendentes.TextMatrix(IndPend, 4) = ctr!chPessoa
GridPendentes.TextMatrix(IndPend, 5) = ctr!ctrDescricaoOperacao
GridPendentes.TextMatrix(IndPend, 6) = Format$(ctr!ctrValorDaBoleta, "##,##0.00")
GridPendentes.TextMatrix(IndPend, 7) = Empty
IndPendlimite = IndPend
AcumulaPendentes = AcumulaPendentes + ctr!ctrValorDaBoleta

End Sub
Public Sub Rotina_051_Carga_Confirmados()

IndConf = IndConf + 1
GridConfirmados.Rows = IndConf + 1
GridConfirmados.TextMatrix(IndConf, 0) = ctr!chFabricante
GridConfirmados.TextMatrix(IndConf, 1) = ctr!chPessoa
GridConfirmados.TextMatrix(IndConf, 2) = ctr!chNotafiscal
GridConfirmados.TextMatrix(IndConf, 3) = ctr!chFatura
GridConfirmados.TextMatrix(IndConf, 4) = ctr!chPessoa
GridConfirmados.TextMatrix(IndConf, 5) = ctr!ctrDescricaoOperacao
GridConfirmados.TextMatrix(IndConf, 6) = ctr!ctrDataVencito
GridConfirmados.TextMatrix(IndConf, 7) = Format$(ctr!ctrValorDaBoleta, "##,##0.00")
IndConfLimite = IndConf
AcumulaConfirmados = AcumulaConfirmados + ctr!ctrValorDaBoleta

End Sub
Public Sub Rotina_052_Carga_Atrasados()

IndAtra = IndAtra + 1
GridAtrasados.Rows = IndAtra + 1
GridAtrasados.TextMatrix(IndAtra, 0) = ctr!chFabricante
GridAtrasados.TextMatrix(IndAtra, 1) = ctr!chPessoa
GridAtrasados.TextMatrix(IndAtra, 2) = ctr!chNotafiscal
GridAtrasados.TextMatrix(IndAtra, 3) = ctr!chFatura
GridAtrasados.TextMatrix(IndAtra, 4) = ctr!chPessoa
GridAtrasados.TextMatrix(IndAtra, 5) = ctr!ctrDescricaoOperacao
GridAtrasados.TextMatrix(IndAtra, 6) = ctr!ctrDataVencito
GridAtrasados.TextMatrix(IndAtra, 7) = Format$(ctr!ctrValorDaBoleta, "##,##0.00")
GridAtrasados.TextMatrix(IndAtra, 8) = Empty
IndAtraLimite = IndAtra
AcumulaAtrasados = AcumulaAtrasados + ctr!ctrValorDaBoleta

End Sub

Private Sub GridAtrasados_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim A As String
Dim B As String
Dim C As String
Dim D As String

If GridAtrasados.TextMatrix(GridAtrasados.Row, 9) = "Ok." Then
   GridAtrasados.TextMatrix(GridAtrasados.Row, 9) = Empty
Else
   If GridAtrasados.TextMatrix(GridAtrasados.Row, 4) = Empty Then
      MsgBox ("Somente utilizar esta tabela quando houver conteúdo."), vbInformation
      Exit Sub
   Else
      frmAjustaCobranca.txtCliente = GridAtrasados.TextMatrix(GridAtrasados.Row, 4)
      frmAjustaCobranca.txtDescOperacao = GridAtrasados.TextMatrix(GridAtrasados.Row, 5)
      frmAjustaCobranca.txtDataVencito = GridAtrasados.TextMatrix(GridAtrasados.Row, 6)
      frmAjustaCobranca.txtValorDaFatura = GridAtrasados.TextMatrix(GridAtrasados.Row, 7)
      frmAjustaCobranca.txtNotaFiscal = GridAtrasados.TextMatrix(GridAtrasados.Row, 2)
      frmAjustaCobranca.txtFatura = GridAtrasados.TextMatrix(GridAtrasados.Row, 3)
      frmAjustaCobranca.Show vbModal
    
      Call Rotina_AbrirBanco

       A = GridAtrasados.TextMatrix(GridAtrasados.Row, 0)
       B = GridAtrasados.TextMatrix(GridAtrasados.Row, 1)
       C = GridAtrasados.TextMatrix(GridAtrasados.Row, 2)
       D = GridAtrasados.TextMatrix(GridAtrasados.Row, 3)
       
       ctr.Open "Select* from contas_a_receber where chFabricante = ('" & A & "') and chPessoa = ('" & B & "') and chNotaFiscal = ('" & C & "')and chFatura = ('" & D & "')", db, 3, 3
       If ctr.EOF Then
          MsgBox ("Erro no acesso para atualizacao de Conta a Receber 1."), vbCritical
          Call FechaDB
          Exit Sub
       End If

       If txtStatusCancela = 0 Then
          GridAtrasados.TextMatrix(GridAtrasados.Row, 7) = Format$(ctr!ctrValorDaBoleta, "#,##0.00")
          GridAtrasados.TextMatrix(GridAtrasados.Row, 9) = "Ok."
       Else
          txtStatusCancela = 0
       End If
   End If
End If
End Sub

Private Sub GridConfirmados_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If GridConfirmados.TextMatrix(GridConfirmados.Row, 8) = "Ok." Then
   GridConfirmados.TextMatrix(GridConfirmados.Row, 8) = Empty
Else
   GridConfirmados.TextMatrix(GridConfirmados.Row, 8) = "Ok."
End If
End Sub

Private Sub GridPendentes_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If GridPendentes.TextMatrix(GridPendentes.Row, 7) = "Ok." Then
   GridPendentes.TextMatrix(GridPendentes.Row, 7) = Empty
Else
   GridPendentes.TextMatrix(GridPendentes.Row, 7) = "Ok."
End If
End Sub


'Public Sub Rotina_070_Ajusta_Data()

'DataInformada = DataParaAjuste
''DataRetorno = ObterDiaUtilAnterior(DataInformada, NDias)
'DataRetorno = DataInformada

'txtDataConsulta = DataRetorno.DiaUtil
'DataPendentes = DataRetorno.DiaUtil

'DataInformada = txtDataConsulta
'DataRetorno = ObterDiaUtilAnterior(DataInformada, NDias)

'DataInicio = DataRetorno.DiaUtil
'DataFim = DataPendentes + 1
 
'If DataInicio + 1 > DiaUtilAnterior Then
'   DataPagos = DiaUtilAnterior
'Else
'   DataPagos = DataPendentes
'End If

'End Sub

Private Sub Label1_Click()

End Sub

Public Sub TotalizaAtrasados()

For IndAtra = 1 To IndAtraLimite
    If GridAtrasados.TextMatrix(IndAtra, 4) = Empty Then
       IndAtra = GridAtrasados.Rows
    Else
       If pessoaAnterior = Empty Then
          SomaAtrasados = GridAtrasados.TextMatrix(IndAtra, 7)
          pessoaAnterior = GridAtrasados.TextMatrix(IndAtra, 4)
       Else
          If GridAtrasados.TextMatrix(IndAtra, 4) = pessoaAnterior Then
            SomaAtrasados = SomaAtrasados + GridAtrasados.TextMatrix(IndAtra, 7)
          Else
            GridAtrasados.TextMatrix((IndAtra - 1), 8) = Format$(SomaAtrasados, "##,##0.00")
            SomaAtrasados = GridAtrasados.TextMatrix(IndAtra, 7)
            pessoaAnterior = GridAtrasados.TextMatrix(IndAtra, 4)
          End If
       End If
    End If
Next

If SomaAtrasados > 0 Then
   GridAtrasados.TextMatrix((IndAtra - 1), 8) = Format$(SomaAtrasados, "##,##0.00")
End If

End Sub
