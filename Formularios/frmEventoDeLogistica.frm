VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEventoDeLogistica 
   Caption         =   "frmEventoDeLogistica"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14985
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   14985
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHoje 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   11
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtPrazoEvento 
      Alignment       =   2  'Center
      Height          =   480
      Left            =   10680
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   735
      Left            =   12360
      TabIndex        =   5
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   735
      Left            =   12360
      TabIndex        =   4
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   855
      Left            =   12360
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid gridEventos 
      Height          =   3855
      Left            =   1560
      TabIndex        =   9
      Top             =   2760
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      FormatString    =   "Cod. Evento |Descrição do Evento                                              |Prazo"
   End
   Begin VB.TextBox txtNomeEvento 
      Height          =   465
      Left            =   3480
      TabIndex        =   1
      Top             =   2040
      Width           =   7095
   End
   Begin VB.TextBox txtCodEvento 
      Alignment       =   2  'Center
      Height          =   480
      Left            =   1560
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Hoje"
      Height          =   375
      Left            =   12000
      TabIndex        =   12
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Prazo dias"
      Height          =   375
      Left            =   10320
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Descrição do Evento"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   1680
      Width           =   6615
   End
   Begin VB.Label Label2 
      Caption         =   "Cod. Evento"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Atualização de Eventos de Logística"
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
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmEventoDeLogistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Limite As Integer
Dim IndLinha As Integer

Private Sub cmdExcluir_Click()

If txtNomeEvento = Empty Then
   MsgBox ("Solicitação inválida. atividade não informado"), vbCritical
   Call FechaDB
   Exit Sub
End If

Call Rotina_AbrirBanco

Ativ.Open "Select * from atividade where atvAtividade = ('" & txtNomeEvento & "')", db, 3, 3
If Ativ.EOF Then
   MsgBox ("Evento inexistente"), vbCritical
   Call FechaDB
   Exit Sub
End If
   
Ativ.Delete

Call FechaDB

Call CargagridEventos

txtNomeEvento = Empty
txtCodEvento = Empty
txtPrazoEvento = Empty

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()

Call Rotina_AbrirBanco

If txtNomeEvento = Empty Then
   MsgBox ("Descrição do atividade para salvar não foi informado"), vbCritical
   Call FechaDB
   Exit Sub
End If

If txtCodEvento = Empty Then
   MsgBox ("Código do atividade para salvar não foi informado"), vbCritical
   Call FechaDB
   Exit Sub
End If

If txtPrazoEvento = Empty Then
   MsgBox ("Prazo do atividade para salvar não foi informado"), vbCritical
   Call FechaDB
   Exit Sub
End If

Ativ.Open "Select * from atividade where atvAtividade = ('" & txtNomeEvento & "')", db, 3, 3
If Ativ.EOF Then
   Ativ.AddNew
End If

Ativ!atvAtividade = txtNomeEvento
Ativ!atvCodigoAtividade = txtCodEvento
Ativ!atvPrazoNormal = txtPrazoEvento

Ativ.Update

Call FechaDB

Call CargagridEventos

txtNomeEvento = Empty
txtCodEvento = Empty
txtPrazoEvento = Empty

'MsgBox ("Centro de custo atualizado com sucesso"), vbInformation

End Sub

Private Sub Form_Load()
txtHoje = Date
CargagridEventos

End Sub

Private Sub gridEventos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Limite = gridEventos.Rows

IndLinha = gridEventos.Row

If gridEventos.TextMatrix(IndLinha, 0) = "" Then
   MsgBox "Para Inclusão informe o novo código. Para Alteração clicar em linha com conteúdo."
   Exit Sub
End If

txtCodEvento = gridEventos.TextMatrix(IndLinha, 0)
txtNomeEvento = gridEventos.TextMatrix(IndLinha, 1)
txtPrazoEvento = gridEventos.TextMatrix(IndLinha, 2)
txtNomeEvento.SetFocus

End Sub

Public Sub CargagridEventos()

Dim IndLinha As Integer

Call Rotina_AbrirBanco

Ativ.Open "Select * from atividade", db, 3, 3
If Ativ.EOF Then
   MsgBox ("Tabela de atividade Vazia."), vbCritical
   Call FechaDB
   Exit Sub
End If

gridEventos.Rows = 2
gridEventos.TextMatrix(1, 0) = Empty
gridEventos.TextMatrix(1, 1) = Empty
gridEventos.TextMatrix(1, 2) = Empty
IndLinha = 0

Ativ.MoveFirst

Do While Not Ativ.EOF
   IndLinha = IndLinha + 1
   gridEventos.Rows = IndLinha + 1
   gridEventos.TextMatrix(IndLinha, 0) = Ativ!atvCodigoAtividade
   gridEventos.TextMatrix(IndLinha, 1) = Ativ!atvAtividade
   If Not IsNull(Ativ!atvPrazoNormal) Then
     gridEventos.TextMatrix(IndLinha, 2) = Ativ!atvPrazoNormal
   Else
     gridEventos.TextMatrix(IndLinha, 2) = 0
   End If
   Ativ.MoveNext
Loop

gridEventos.Sort = 1

Call FechaDB

End Sub



Private Sub txtNomeEvento_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


