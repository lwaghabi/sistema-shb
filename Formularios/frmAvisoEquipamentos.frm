VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAvisoEquipamentos 
   Caption         =   "frmAvisoEquipamentos"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16245
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   16245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
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
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7270
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   5055
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   16095
      Begin MSFlexGridLib.MSFlexGrid grdEquip 
         Height          =   4575
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   15855
         _ExtentX        =   27966
         _ExtentY        =   8070
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         FormatString    =   "Unidade Operacional                       |Equipamento                    |Evento                           |Data Vencimento    |"
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   16095
      Begin VB.TextBox txtHoje 
         Alignment       =   2  'Center
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
         Left            =   13320
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Relação de Equipamentos com eventos dentro da data de aviso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   11295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Hoje"
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
         Left            =   13320
         TabIndex        =   3
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "AVISO - EQUIPAMENTOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   7500
      End
   End
End
Attribute VB_Name = "frmAvisoEquipamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EquipTipoAnterior As String
Dim ChaveAuxiliar As String
Dim DataProxManut As Date
Dim Linha As Integer


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

txtHoje = Date
Linha = 0

Call Rotina_AbrirBanco

eqpt.Open "Select * from Equipamento", db, 3, 3
If eqpt.EOF Then
   MsgBox ("Cadastro de Equipamentos vazio. Comunicar ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
End If

Linha = 1

eqpt.MoveFirst

Do While Not eqpt.EOF
   If Not EquipTipoAnterior = eqpt!eqptTipoEquipamento Then
      EquipTipoAnterior = eqpt!eqptTipoEquipamento
      If teq.State = 1 Then
         teq.Close: Set teq = Nothing
      End If
      teq.Open "Select * from EquipamentoTipo where chTipoDeEquipamento = ('" & eqpt!eqptTipoEquipamento & "')", db, 3, 3
      If teq.EOF Then
         MsgBox ("Erro no acesso a Tipo de Equipamento."), vbCritical
         Call FechaDB
         Exit Sub
      End If
   End If

   If Not IsNull(eqpt!eqptDataValidade) Then
      If ((eqpt!eqptDataValidade - teq!teqDiasAntecedencia) < Date) And Not (eqpt!eqptStatusCalibracao = "EM CALIBRAÇÃO") Then
         If Prod.State = 1 Then
            Prod.Close: Set Prod = Nothing
         End If
         Prod.Open "Select * from Produto where chProduto = ('" & eqpt!eqptProdVinculado & "')", db, 3, 3
         If Prod.EOF Then
            MsgBox ("Produto não cadastrado. Verificar"), vbCritical
            Call FechaDB
            Exit Sub
         End If
         Call CarregaGrid
      End If
   Else
      MsgBox ("Data validade invalida o processo."), vbCritical
      Call FechaDB
      Exit Sub
   End If

   eqpt.MoveNext

Loop

Call FechaDB

End Sub

Public Sub CarregaGrid()

grdEquip.Rows = Linha + 1

grdEquip.TextMatrix(Linha, 0) = Prod!prdLocadora & " - " & Prod!prdUnidadeOperacional
grdEquip.TextMatrix(Linha, 1) = eqpt!chCodEquipamento
If teq!teqCalibracao = 1 Then
   grdEquip.TextMatrix(Linha, 2) = "CALIBRAÇÃO"
Else
   grdEquip.TextMatrix(Linha, 2) = "MANUTENÇÃO"
End If
grdEquip.TextMatrix(Linha, 3) = eqpt!eqptDataValidade
Linha = Linha + 1
End Sub
