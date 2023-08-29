VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAvisoReembolso 
   Caption         =   "frmAvisoReembolso"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   11625
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optAviso 
      Caption         =   "Não mostrar esse aviso"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6360
      Width           =   2715
   End
   Begin VB.TextBox txtHoje 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   2775
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H0000FFFF&
      Caption         =   "Fechar Aviso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid grdAviso 
      Height          =   4455
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7858
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "Colaborador                                                                  |Data Reembolso        "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "REEMBOLSO - AVISO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   4
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label Label2 
      Caption         =   "Relação de Colaborador Com Lançamento de Reembolso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   10935
   End
End
Attribute VB_Name = "frmAvisoReembolso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ColaboradorAnterior As String
Dim DataAnterior As Date
Dim DataBase As String
Dim DataDias As Date
Dim DataInvertida As String
Dim DataHojeInvertida As String

Dim Dia As Integer
Dim Mes As Integer
Dim Ano As Integer
Dim DiaDb As Integer
Dim MesDb As Integer
Dim AnoDb As Integer

Dim Linha As Integer


Private Sub cmdSair_Click()

If Not optAviso = True Then
   Unload Me
End If

Call Rotina_AbrirBanco

usu.Open "Select * from Usuario where chNome = ('" & glbUsuario & "')", db, 3, 3
If usu.EOF Then
   MsgBox ("Erro no acesso a Usuario na rotina de atualização de mostrar aviso. Comunicar Analista responsável"), vbCritical
   End
End If

If optAviso = True Then
   usu!usuAvisoReembolso = 0
   usu.Update
End If

Unload Me

End Sub

Private Sub Form_Load()

txtHoje = Date
ColaboradorAnterior = Empty
DataAnterior = Empty
optAviso = False

Ano = Year(Date)
Mes = Month(Date)
Dia = Day(Date)

DataHojeInvertida = Ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")

Call Rotina_AbrirBanco

usu.Open "Select * from Usuario where  chNome = ('" & glbUsuario & "')", db, 3, 3
If usu.EOF Then
   MsgBox ("Usuário inexistente. Comunicar ao analista responsável."), vbCritical
   Call FechaDB
   End
End If

If Not usu!usuAvisoReembolso = 1 Then
   Call FechaDB
   Exit Sub
   Unload Me
End If

Call LimpaGridAviso

Rmb.Open "Select * from Reembolso where rmbStatusReembolso = ('" & 0 & "')", db, 3, 3
If Rmb.EOF Then
   Call FechaDB
   Exit Sub
End If

Linha = 1

Rmb.MoveFirst

Do While Not Rmb.EOF
   If Not Rmb!RmbColaborador = ColaboradorAnterior Then
      grdAviso.Rows = Linha + 1
      grdAviso.TextMatrix(Linha, 0) = Rmb!RmbNomeColaborador
      grdAviso.TextMatrix(Linha, 1) = Rmb!RmbDataLancReembolso
      ColaboradorAnterior = Rmb!RmbColaborador
      Linha = Linha + 1
   End If
   Rmb.MoveNext
Loop
 
Call FechaDB

End Sub
Public Sub LimpaGridAviso()
grdAviso.Rows = 2
Linha = 1
    grdAviso.TextMatrix(Linha, 0) = Empty
    grdAviso.TextMatrix(Linha, 1) = Empty
End Sub

