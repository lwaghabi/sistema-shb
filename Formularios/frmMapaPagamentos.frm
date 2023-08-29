VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMapaPagamentos 
   Caption         =   "Mapa de Pagamentos"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3945
   LinkTopic       =   "Form4"
   ScaleHeight     =   3375
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H000000FF&
      Caption         =   "Sair"
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdVisualizar 
      BackColor       =   &H00FFFF00&
      Caption         =   "Visualizar"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período de "
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
      Begin MSComCtl2.DTPicker DTATE 
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   260505601
         CurrentDate     =   38202
      End
      Begin MSComCtl2.DTPicker DTDE 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   260505601
         CurrentDate     =   38202
      End
      Begin VB.Label Label2 
         Caption         =   "Á"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa de Pagamentos "
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
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "frmMapaPagamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sql As String
Dim DataDe As String
Dim DataAte As String


Private Sub cmdSair_Click()
Unload Me

End Sub

Private Sub cmdVisualizar_Click()
DataDe = Format$(CDate(DTDE), "mm/dd/yyyy")
DataAte = Format$(CDate(DTATE), "mm/dd/yyyy")

Sql = "Select pg.chpessoa, pg.chcodbcolart,pg.chfatura, pg.ctpDescricaoOperacao,"
Sql = Sql & " pg.ctpValorDaBoleta, pg.chDataVencito, pg.ctpdatapagamento from Contas_A_Pagar pg "
Sql = Sql & " where pg.ctpStatus = 1 and pg.ctptipolancamento < 99 "
Sql = Sql & " and pg.ctpdataproc between #" & DataDe & "# and #" & DataAte & "#"
Sql = Sql & " order by pg.ctpdatapagamento, pg.chpessoa"
'If TabCtaPagar("chcodbcolart") = "UNIBANCO" Then
'   SQL = SQL & " and pg.chcodbcolart like '" & "UB" & "'"
'Else
'    If TabCtaPagar("chcodbcolart") = "B. BRASIL" Then
'       SQL = SQL & " and pg.chcodbcolart like '" & "BB" & "'"
'    Else
'        SQL = SQL & " and pg.chcodbcolart like '" & "CX" & "'"
'    End If
'
'End If


'MsgBox Sql

'deMapaPagtos.Commands.Item("cmdPagto").CommandText = Sql

'ImpMapaPagtos.Show vbModal

'deMapaPagtos.rscmdPagto.Close

End Sub

Private Sub Form_Load()

DTDE = Date
DTATE = Date

End Sub
