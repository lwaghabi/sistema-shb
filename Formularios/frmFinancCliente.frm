VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFinancCliente 
   Caption         =   " frmFinancCliente"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17205
   LinkTopic       =   "Form3"
   ScaleHeight     =   8310
   ScaleWidth      =   17205
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid GridVencidos 
      Height          =   4815
      Left            =   7080
      TabIndex        =   21
      Top             =   3480
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   16777152
      BackColorFixed  =   16776960
      BackColorBkg    =   16777152
      FormatString    =   "Vencito       |N Fiscal|Fatura               |Valor         |Data Receb|Status              |"
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
   Begin MSFlexGridLib.MSFlexGrid GridAVencer 
      Height          =   4815
      Left            =   0
      TabIndex        =   20
      Top             =   3480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   16777152
      ForeColor       =   0
      BackColorFixed  =   16776960
      BackColorBkg    =   16777152
      GridColor       =   0
      FormatString    =   "Vencimento|N.Fiscal  |Operação            |Valor           |"
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
   Begin VB.Frame Frame2 
      Caption         =   "Consolidado"
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
      Left            =   7320
      TabIndex        =   6
      Top             =   480
      Width           =   9735
      Begin VB.Label txtTotalNegociado 
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
         Left            =   7320
         TabIndex        =   12
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Total Negociado"
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
         TabIndex        =   11
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label txtTotalVencido 
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
         Left            =   7320
         TabIndex        =   10
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Total Vencido"
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
         TabIndex        =   9
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label txtTotalAVencer 
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
         Left            =   7320
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Total a Vencer"
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
         TabIndex        =   7
         Top             =   720
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   2385
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   7095
      Begin VB.ComboBox cmbTipoConsulta 
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
         TabIndex        =   18
         Top             =   240
         Width           =   2655
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   0
         TabIndex        =   5
         Top             =   1320
         Width           =   6735
         Begin VB.CommandButton cmdConsulta 
            BackColor       =   &H00FFFF00&
            Caption         =   "Consulta"
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
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdNovaConsulta 
            BackColor       =   &H0000FF00&
            Caption         =   "Nova Consulta"
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
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdSair 
            BackColor       =   &H008080FF&
            Caption         =   "Sair"
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
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.ComboBox cmbFiltro 
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
         TabIndex        =   1
         Top             =   840
         Width           =   4695
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Consulta"
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
         TabIndex        =   19
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSMask.MaskEdBox txtHoje 
      Height          =   495
      Left            =   15240
      TabIndex        =   13
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Height          =   495
      Left            =   14040
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vencido"
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
      TabIndex        =   16
      Top             =   3000
      Width           =   9975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A Vencer"
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
      Left            =   0
      TabIndex        =   15
      Top             =   3000
      Width           =   7095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Financeiro por Cliente/Colaborador"
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
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmFinancCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim i As Long
Dim QtdReg As Long
Dim Fim As Byte
Dim RegMes As Byte
Dim Linha As Long
Dim IndVencer As Integer
Dim IndVencido As Integer
Dim AcumulaVencer As Currency
Dim AcumulaVencido As Currency

Dim AnoInv As String
Dim MesInv As String
Dim DiaInv As String

Dim DataHoje As Date
Dim DataUtil As Date

Dim Resp As String
Dim Fatura As String
Dim Pos As Integer




Private Sub cmbTipoConsulta_LostFocus()

cmbFiltro.Clear

i = 1

Call Rotina_AbrirBanco

pes.Open "Select * from pessoa", db, 3, 3
If pes.EOF Then
   MsgBox ("Tabele de pessoa vazia. Verificar."), vbCritical
   Call FechaDB
   Exit Sub
End If


pes.MoveFirst

Do While Not pes.EOF
   If cmbTipoConsulta.ListIndex = 0 Then
      If pes!pestipopessoa = 0 Then
         cmbFiltro.AddItem pes!chPessoa
         frmProgressao.ProgressBar1.Value = i * 100 / (QtdReg - 1)
         pes.MoveNext
      Else
         pes.MoveNext
      End If
      i = i + 1
      If Not pes.EOF Then
         If i = QtdReg Then
            QtdReg = QtdReg + 1
         End If
      End If
   Else
      If pes!pestipopessoa <> 0 Then
         cmbFiltro.AddItem pes!chPessoa
         'frmProgressao.ProgressBar1.Value = i * 100 / (QtdReg - 1)
         pes.MoveNext
      Else
         pes.MoveNext
      End If
      i = i + 1
      If Not pes.EOF Then
         If i = QtdReg Then
            QtdReg = QtdReg + 1
         End If
      End If
   End If
'Next
Loop
frmProgressao.Hide
cmbFiltro.ListIndex = 0

Call FechaDB

End Sub


Private Sub cmdConsulta_Click()

Fim = 0

Call Rotina_030_Limpa_Vencer
Call Rotina_035_Limpa_Vencido

RegMes = 0
IndVencer = 0
IndVencido = 0
AcumulaVencer = 0
AcumulaVencido = 0

If cmbTipoConsulta.ListIndex = 0 Then
   Call Rotina_060_Consulta_Cliente
Else
   Call Rotina_070_Consulta_Colaborador
End If

If IndVencer > 1 Then
   GridAVencer.Row = 1
   GridAVencer.Col = 4
   GridAVencer.RowSel = IndVencer
   GridAVencer.ColSel = 4
   GridAVencer.Sort = 1
   GridAVencer.Row = 0
   GridAVencer.RowSel = 0
   GridAVencer.Col = 0
   GridAVencer.ColSel = 0
End If

If IndVencido > 1 Then
   GridVencidos.Row = 1
   GridVencidos.Col = 6
   GridVencidos.RowSel = IndVencido
   GridVencidos.ColSel = 6
   GridVencidos.Sort = 2
   GridVencidos.Row = 0
   GridVencidos.RowSel = 0
   GridVencidos.Col = 0
   GridVencidos.ColSel = 0
End If

txtTotalVencido = Format$(AcumulaVencido, "#,##0.00")

txtTotalNegociado = Format$(AcumulaVencer + AcumulaVencido, "#,##0.00")
End Sub

Private Sub cmdNovaConsulta_Click()
cmbFiltro.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

txtHoje = Date

cmbTipoConsulta.AddItem "Cliente"
cmbTipoConsulta.AddItem "Colaborador"

cmbTipoConsulta.ListIndex = 0

Call Rotina_AbrirBanco

pes.Open "Select * from pessoa", db, 3, 3
If pes.EOF Then
   MsgBox ("Tabele de pessoa vazia. Verificar."), vbCritical
   Call FechaDB
   Exit Sub
End If


pes.MoveFirst

Do While Not pes.EOF
   If cmbTipoConsulta.ListIndex = 0 Then
      If pes!pestipopessoa = 0 Then
         cmbFiltro.AddItem pes!chPessoa
         pes.MoveNext
      Else
         pes.MoveNext
      End If
   Else
      If pes!pestipopessoa <> 0 Then
         cmbFiltro.AddItem pes!chPessoa
         pes.MoveNext
      Else
         pes.MoveNext
      End If
   End If
Loop

cmbFiltro.ListIndex = 1

FechaDB

End Sub

Public Sub Rotina_030_Limpa_Vencer()

GridAVencer.Rows = 2
IndVencer = 1
GridAVencer.TextMatrix(IndVencer, 0) = Empty
GridAVencer.TextMatrix(IndVencer, 1) = Empty
GridAVencer.TextMatrix(IndVencer, 2) = Empty
GridAVencer.TextMatrix(IndVencer, 3) = Empty
GridAVencer.TextMatrix(IndVencer, 4) = Empty

End Sub

Public Sub Rotina_035_Limpa_Vencido()

GridVencidos.Rows = 2
IndVencido = 1
GridVencidos.TextMatrix(IndVencido, 0) = Empty
GridVencidos.TextMatrix(IndVencido, 1) = Empty
GridVencidos.TextMatrix(IndVencido, 2) = Empty
GridVencidos.TextMatrix(IndVencido, 3) = Empty
GridVencidos.TextMatrix(IndVencido, 4) = Empty
GridVencidos.TextMatrix(IndVencido, 5) = Empty
GridVencidos.TextMatrix(IndVencido, 6) = Empty

End Sub

Public Sub Rotina_060_Consulta_Cliente()

Call Rotina_AbrirBanco

ctr.Open "Select * from contas_a_receber", db, 3, 3
If ctr.EOF Then
   MsgBox ("Tabele de Contas a Receber vazia. Verificar."), vbCritical
   Call FechaDB
   Exit Sub
End If


ctr.MoveFirst
   
Do While Fim = 0
   If ctr!chPessoa < cmbFiltro Then
      ctr.MoveNext
      If ctr.EOF Then
         Fim = 1
      End If
   Else
      If ctr!chPessoa > cmbFiltro Then
         Fim = 1
      Else
         If ctr!ctrDataVencito < Date Then
            RegMes = 1
            Call Rotina_064_Carga_Vencidos
         Else
            Call Rotina_066_Carga_Vencer
         End If
         ctr.MoveNext
         If ctr.EOF Then
            Fim = 1
         End If
      End If
   End If
Loop

txtTotalAVencer = Format$(AcumulaVencer, "#,##0.00")

Fim = 0

hctr.Open "Select * from historicocontasreceber", db, 3, 3
If hctr.EOF Then
   MsgBox ("Tabele de Historico Contas a Receber vazia. Verificar."), vbCritical
   Call FechaDB
   Exit Sub
End If


hctr.MoveFirst
   

Do While Fim = 0
   If hctr!chPessoa > cmbFiltro Then
      hctr.MovePrevious
      If hctr.BOF Then
         Fim = 1
      End If
   Else
      If hctr!chPessoa < cmbFiltro Then
         Fim = 1
      Else
         RegMes = 0
         Call Rotina_064_Carga_Vencidos
         hctr.MovePrevious
         If hctr.BOF Then
            Fim = 1
         End If
      End If
   End If
Loop

Call FechaDB

End Sub
Public Sub Rotina_064_Carga_Vencidos()
If RegMes = 1 Then
   Call Rotina_067_Vencido_Mes
Else
   Call Rotina_068_Vencido_Hist
End If
End Sub

Public Sub Rotina_066_Carga_Vencer()

IndVencer = IndVencer + 1
GridAVencer.Rows = IndVencer + 1
GridAVencer.TextMatrix(IndVencer, 0) = ctr!ctrDataVencito
AnoInv = Year(ctr!ctrDataVencito)
MesInv = Month(ctr!ctrDataVencito)
DiaInv = Day(ctr!ctrDataVencito)
GridAVencer.TextMatrix(IndVencer, 4) = AnoInv & Format(MesInv, "00") & Format(DiaInv, "00")
GridAVencer.TextMatrix(IndVencer, 1) = ctr!chNotafiscal
GridAVencer.TextMatrix(IndVencer, 2) = ctr!ctrDescricaoOperacao
GridAVencer.TextMatrix(IndVencer, 3) = Format$(ctr!ctrValorDaBoleta, "#,##0.00")

AcumulaVencer = AcumulaVencer + ctr!ctrValorDaBoleta

End Sub
Public Sub Rotina_067_Vencido_Mes()

IndVencido = IndVencido + 1

DataUtil = ctr!ctrDataVencito

DataInformada = DataUtil
NDias = 0

'DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)

'DataUtil = DataRetorno.DiaUtil

GridVencidos.Rows = IndVencido + 1
GridVencidos.TextMatrix(IndVencido, 0) = ctr!ctrDataVencito

AnoInv = Year(ctr!ctrDataVencito)
MesInv = Month(ctr!ctrDataVencito)
DiaInv = Day(ctr!ctrDataVencito)
GridVencidos.TextMatrix(IndVencido, 6) = AnoInv & Format(MesInv, "00") & Format(DiaInv, "00")

GridVencidos.TextMatrix(IndVencido, 1) = ctr!chNotafiscal
GridVencidos.TextMatrix(IndVencido, 2) = ctr!chFatura & "-" & ctr!ctrDescricaoOperacao
GridVencidos.TextMatrix(IndVencido, 3) = Format$(ctr!ctrValorDaBoleta, "#,##0.00")

If ctr!ctrStatus = 0 Then
   GridVencidos.TextMatrix(IndVencido, 5) = "Em Atraso"
Else
   If ctr!ctrDataRecebimento > DataUtil Then
      GridVencidos.TextMatrix(IndVencido, 4) = ctr!ctrDataRecebimento
      GridVencidos.TextMatrix(IndVencido, 5) = "C/Atraso"
   Else
      If ctr!ctrDataRecebimento < DataUtil Then
         GridVencidos.TextMatrix(IndVencido, 4) = ctr!ctrDataRecebimento
         GridVencidos.TextMatrix(IndVencido, 5) = "Antecipado"
      Else
         GridVencidos.TextMatrix(IndVencido, 4) = ctr!ctrDataRecebimento
         GridVencidos.TextMatrix(IndVencido, 5) = "Ok"
      End If
   End If
End If

AcumulaVencido = AcumulaVencido + ctr!ctrValorDaBoleta

If IndVencido = 99 Then
   Fim = 1
End If

End Sub


Public Sub Rotina_068_Vencido_Hist()

IndVencido = IndVencido + 1

DataUtil = hctr!ctrDataVencito

DataInformada = DataUtil
'NDias = 0

'DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)

'DataUtil = DataRetorno.DiaUtil
GridVencidos.Rows = IndVencido + 1
GridVencidos.TextMatrix(IndVencido, 0) = hctr!ctrDataVencito

AnoInv = Year(hctr!ctrDataVencito)
MesInv = Month(hctr!ctrDataVencito)
DiaInv = Day(hctr!ctrDataVencito)
GridVencidos.Rows = IndVencido + 1
GridVencidos.TextMatrix(IndVencido, 6) = AnoInv & Format(MesInv, "00") & Format(DiaInv, "00")

GridVencidos.TextMatrix(IndVencido, 1) = hctr!chNotafiscal
GridVencidos.TextMatrix(IndVencido, 2) = hctr!chFatura & "-" & hctr!ctrDescricaoOperacao
GridVencidos.TextMatrix(IndVencido, 3) = Format$(hctr!ctrValorDaBoleta, "#,##0.00")

If hctr!ctrStatus = 0 Then
   GridVencidos.TextMatrix(IndVencido, 5) = "Em Atraso"
Else
   If hctr!ctrDataRecebimento > DataUtil Then
      GridVencidos.TextMatrix(IndVencido, 4) = hctr!ctrDataRecebimento
      GridVencidos.TextMatrix(IndVencido, 5) = "C/Atraso"
   Else
      If hctr!ctrDataRecebimento < DataUtil Then
         GridVencidos.TextMatrix(IndVencido, 4) = hctr!ctrDataRecebimento
         GridVencidos.TextMatrix(IndVencido, 5) = "Antecipado"
      Else
         GridVencidos.TextMatrix(IndVencido, 4) = hctr!ctrDataRecebimento
         GridVencidos.TextMatrix(IndVencido, 5) = "Ok"
      End If
   End If
End If

AcumulaVencido = AcumulaVencido + hctr!ctrValorDaBoleta

End Sub
Public Sub Rotina_070_Consulta_Colaborador()

Call Rotina_AbrirBanco

ctp.Open "Select * from contas_a_receber", db, 3, 3
If ctp.EOF Then
   MsgBox ("Tabele de Contas a Receber vazia. Verificar."), vbCritical
   Call FechaDB
   Exit Sub
End If


ctp.MoveFirst
Do While Fim = 0
   If ctp!chPessoa < cmbFiltro Then
      ctp.MoveNext
      If ctp.EOF Then
         Fim = 1
      End If
   Else
      If ctp!chPessoa > cmbFiltro Then
         Fim = 1
      Else
         If ctp!chDataVencito < Date Then
            RegMes = 1
            Call Rotina_074_Carga_Vencidos
         Else
            Call Rotina_076_Carga_Vencer
         End If
         ctp.MoveNext
         If ctp.EOF Then
            Fim = 1
         End If
      End If
   End If
Loop

txtTotalAVencer = Format$(AcumulaVencer, "#,##0.00")

Fim = 0


hctp.Open "Select * from contas_a_receber", db, 3, 3
If hctp.EOF Then
   MsgBox ("Tabele de Contas a Receber vazia. Verificar."), vbCritical
   Call FechaDB
   Exit Sub
End If


hctp.MoveFirst
Do While Fim = 0
   If hctp!chFabricante = 1 Then
      hctp.MovePrevious
      If hctp.BOF Then
         Fim = 1
      End If
   Else
      If hctp!chPessoa > cmbFiltro Then
         hctp.MovePrevious
         If hctp.BOF Then
            Fim = 1
         End If
      Else
         If hctp!chPessoa < cmbFiltro Then
            Fim = 1
         Else
            RegMes = 0
            Call Rotina_074_Carga_Vencidos
            hctp.MovePrevious
            If hctp.BOF Then
               Fim = 1
            End If
         End If
      End If
   End If
Loop

End Sub
Public Sub Rotina_074_Carga_Vencidos()
If RegMes = 1 Then
   Call Rotina_077_Vencido_Mes
Else
   Call Rotina_078_Vencido_Hist
End If
End Sub

Public Sub Rotina_076_Carga_Vencer()

IndVencer = IndVencer + 1
GridAVencer.Rows = IndVencer + 1
GridAVencer.TextMatrix(IndVencer, 0) = ctp!chDataVencito
AnoInv = Year(ctp!chDataVencito)
MesInv = Month(ctp!chDataVencito)
DiaInv = Day(ctp!chDataVencito)
GridAVencer.TextMatrix(IndVencer, 4) = AnoInv & Format(MesInv, "00") & Format(DiaInv, "00")
GridAVencer.TextMatrix(IndVencer, 1) = ctp!chNotafiscal
GridAVencer.TextMatrix(IndVencer, 2) = ctp!ctpdescricaooperacao
GridAVencer.TextMatrix(IndVencer, 3) = Format$(ctp!ctpValorDaBoleta, "#,##0.00")

AcumulaVencer = AcumulaVencer + ctp!ctpValorDaBoleta

End Sub
Public Sub Rotina_077_Vencido_Mes()

IndVencido = IndVencido + 1

DataUtil = ctp!chDataVencito

DataInformada = DataUtil
NDias = 0

'DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)

'DataUtil = DataRetorno.DiaUtil

GridVencidos.Rows = IndVencido + 1
GridVencidos.TextMatrix(IndVencido, 0) = ctp!chDataVencito

AnoInv = Year(ctp!chDataVencito)
MesInv = Month(ctp!chDataVencito)
DiaInv = Day(ctp!chDataVencito)
GridVencidos.TextMatrix(IndVencido, 6) = AnoInv & Format(MesInv, "00") & Format(DiaInv, "00")

GridVencidos.TextMatrix(IndVencido, 1) = ctp!chNotafiscal
GridVencidos.TextMatrix(IndVencido, 2) = ctp!ctpdescricaooperacao
GridVencidos.TextMatrix(IndVencido, 3) = Format$(ctp!ctpValorDaBoleta, "#,##0.00")

If ctp!ctpStatus = 0 Then
   GridVencidos.TextMatrix(IndVencido, 5) = "Em Atraso"
Else
   If ctp!ctpDataPagamento > DataUtil Then
      GridVencidos.TextMatrix(IndVencido, 4) = ctp!ctpDataPagamento
      GridVencidos.TextMatrix(IndVencido, 5) = "C/Atraso"
   Else
      If ctp!ctpDataPagamento < DataUtil Then
         GridVencidos.TextMatrix(IndVencido, 4) = ctp!ctpDataPagamento
         GridVencidos.TextMatrix(IndVencido, 5) = "Antecipado"
      Else
         GridVencidos.TextMatrix(IndVencido, 4) = ctp!ctpDataPagamento
         GridVencidos.TextMatrix(IndVencido, 5) = "Ok"
      End If
   End If
End If

AcumulaVencido = AcumulaVencido + ctp!ctpValorDaBoleta

End Sub


Public Sub Rotina_078_Vencido_Hist()

IndVencido = IndVencido + 1

DataUtil = hctp!chDataVencito

DataInformada = DataUtil
NDias = 0

'DataRetorno = ObterProximoDiaUtil(DataInformada, NDias)

'DataUtil = DataRetorno.DiaUtil
GridVencidos.Rows = IndVencido + 1
GridVencidos.TextMatrix(IndVencido, 0) = hctp!chDataVencito

AnoInv = Year(hctp!chDataVencito)
MesInv = Month(hctp!chDataVencito)
DiaInv = Day(hctp!chDataVencito)
GridVencidos.Rows = IndVencido + 1
GridVencidos.TextMatrix(IndVencido, 6) = AnoInv & Format(MesInv, "00") & Format(DiaInv, "00")

GridVencidos.TextMatrix(IndVencido, 1) = hctp!chNotafiscal
GridVencidos.TextMatrix(IndVencido, 2) = hctp!ctpdescricaooperacao
GridVencidos.TextMatrix(IndVencido, 3) = Format$(hctp!ctpValorDaBoleta, "#,##0.00")

If hctp!ctpStatus = 0 Then
   GridVencidos.TextMatrix(IndVencido, 5) = "Em Atraso"
Else
   If hctp!ctpDataPagamento > DataUtil Then
      GridVencidos.TextMatrix(IndVencido, 4) = hctp!ctpDataPagamento
      GridVencidos.TextMatrix(IndVencido, 5) = "C/Atraso"
   Else
      If hctp!ctpDataPagamento < DataUtil Then
         GridVencidos.TextMatrix(IndVencido, 4) = hctp!ctpDataPagamento
         GridVencidos.TextMatrix(IndVencido, 5) = "Antecipado"
      Else
         GridVencidos.TextMatrix(IndVencido, 4) = hctp!ctpDataPagamento
         GridVencidos.TextMatrix(IndVencido, 5) = "Ok"
      End If
   End If
End If

AcumulaVencido = AcumulaVencido + hctp!ctpValorDaBoleta

End Sub


Private Sub GridVencidos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

DataHoje = GridVencidos.TextMatrix(GridVencidos.Row, 0)

'If Month(DataHoje) = Month(Date) Then
'   MsgBox "Para restaurar operação financeira do mes, utilize a função Recebimentos/Pagamentos do Mes"
'   Exit Sub
'End If

Resp = MsgBox("Deseja Restaurar a Operação Financeira " & GridVencidos.TextMatrix(GridVencidos.Row, 2) & " do Cliente " & cmbFiltro, vbYesNo)
If Resp = vbYes Then
   If cmbTipoConsulta.ListIndex = 0 Then
      Restaura_Recebimentos
   Else
      Restaura_Pagamentos
   End If
End If
End Sub

Public Sub Restaura_Recebimentos()

Linha = GridVencidos.Row
Pos = 0
Fim = 0
For Pos = 1 To 3
   Resp = Mid$(GridVencidos.TextMatrix(Linha, 2), Pos, 1)
   If Resp = "-" Then
      If Pos > 2 Then
         Fatura = Mid$(GridVencidos.TextMatrix(Linha, 2), 1, 1) & Mid$(GridVencidos.TextMatrix(Linha, 2), 2, 1)
      Else
         Fatura = Mid$(GridVencidos.TextMatrix(Linha, 2), 1, 1)
      End If
   End If
Next

Call Rotina_AbrirBanco

ctr.Open "Select * from contas_a_receber where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbFiltro & "') and chNotaFiscal = ('" & GridVencidos.TextMatrix(Linha, 1) & "') and  chFatura = ('" & Fatura & "')", db, 3, 3
If ctr.EOF Then
   hctr.Open "Select * from historicocontasreceber where chFabricante = ('" & 0 & "') and chPessoa = ('" & cmbFiltro & "') and chNotaFiscal = ('" & GridVencidos.TextMatrix(Linha, 1) & "') and  chFatura = ('" & Fatura & "')", db, 3, 3
   If hctr.EOF Then
      MsgBox ("Atenção: Restauração solicitada não foi realizada"), vbInformation
      Call FechaDB
      Exit Sub
   Else
      ctr.AddNew
      ctr!chFabricante = hctr!chFabricante
      ctr!chPessoa = hctr!chPessoa
      ctr!chNotafiscal = hctr!chNotafiscal
      ctr!chFatura = hctr!chFatura
      ctr!ctrDataEmissao = hctr!ctrDataEmissao
      ctr!ctrDataVencito = hctr!ctrDataVencito
      ctr!ctrDataBanco = hctr!ctrDataBanco
      ctr!ctrDataVencitoOriginal = hctr!ctrDataVencOriginal
      ctr!ctrDescricaoOperacao = hctr!ctrDescricaoOperacao
      ctr!ctrValorLart = hctr!ctrValorLart
      ctr!ctrValorMerco = hctr!ctrValorMerco
      ctr!ctrPercentCorrecao = hctr!ctrPercentCorrecao
      ctr!ctrvalorcorrecao = hctr!ctrvalorcorrecao
      ctr!ctrValorDaBoleta = hctr!ctrValorDaBoleta
      ctr!chAno = hctr!chAno
      ctr!chMes = hctr!chMes
      ctr!chDia = hctr!chDia
      ctr!chNumPedido = hctr!chNumPedido
      ctr!chNumPedidoComp = hctr!chNumPedidoComp
      ctr!chCodBcoLart = hctr!chCodBcoLart
      ctr!ctrDataRecebimento = hctr!ctrDataRecebimento
      ctr!ctrStatus = 0
      ctr.Update
      hctr.Delete
      MsgBox ("Status de Recebimento Restaurado"), vbInformation
      Call cmdConsulta_Click
   End If
Else
   If Not (ctr!ctrStatus) = 0 Then
      ctr!ctrStatus = 0
      ctr.Update
      Call cmdConsulta_Click
      MsgBox ("Status de Recebimento Restaurado")
   Else
      MsgBox "Este Recebimento ja se encontra pendente"
   End If
End If
End Sub

Public Sub Restaura_Pagamentos()
MsgBox "Restauração de Contas a Pagar não disponível"
End Sub



