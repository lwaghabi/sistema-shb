VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTreinamentos 
   Caption         =   "frmTreinamentos"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17355
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   17355
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12600
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.ComboBox cmbIncidencia 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10680
      TabIndex        =   17
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox lblHoje 
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
      Left            =   15000
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox txtNomeCurso 
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
      TabIndex        =   0
      Top             =   2040
      Width           =   6375
   End
   Begin VB.ComboBox cmbTipoPrazo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtPrazoValidade 
      Alignment       =   2  'Center
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
      Left            =   8640
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtAvisoEm 
      Alignment       =   2  'Center
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
      Left            =   9720
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comandos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   15000
      TabIndex        =   8
      Top             =   960
      Width           =   2175
      Begin VB.CommandButton cmdSalvar 
         BackColor       =   &H0000FF00&
         Caption         =   "Salvar"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdExcluir 
         BackColor       =   &H000000FF&
         Caption         =   "Excluir"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H0000FFFF&
         Caption         =   "Sair"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1560
         Width           =   1695
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdCursos 
      Height          =   5415
      Left            =   0
      TabIndex        =   7
      Top             =   3240
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   9551
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      FormatString    =   $"frmTreinamentos.frx":0000
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
   Begin VB.Label Label9 
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
      Height          =   375
      Left            =   12600
      TabIndex        =   19
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Incidência"
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
      Left            =   10680
      TabIndex        =   18
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Dias"
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
      Left            =   9840
      TabIndex        =   16
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "REGISTRO E ATUALIZAÇÃO DE CURSOS E TREINAMENTOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   9615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "HOJE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15000
      TabIndex        =   14
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Nome do Curso/Treinamento"
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
      TabIndex        =   13
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Cntrl.prazo"
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
      Left            =   6360
      TabIndex        =   12
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Valid."
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
      Left            =   8520
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Aviso"
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
      Left            =   9720
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
End
Attribute VB_Name = "frmTreinamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Limite As Integer
Dim IndLinha As Integer

Private Sub Form_Load()
lblHoje = Date

cmbTipoPrazo.AddItem "QTD DIAS"
cmbTipoPrazo.AddItem "QTD MESES"
cmbTipoPrazo.AddItem "QTD ANOS"
cmbTipoPrazo.ListIndex = 0

cmbIncidencia.AddItem "Administrativo"
cmbIncidencia.AddItem "Operacional"
cmbIncidencia.AddItem "Adm/Oper"

'cmbUnidTempo.AddItem "Dia"
'cmbUnidTempo.AddItem "Mês"
'cmbUnidTempo.AddItem "Hora"
'cmbUnidTempo.ListIndex = 0

cmbStatus.AddItem "Descontinuado"
cmbStatus.AddItem "Ativo"
cmbStatus.ListIndex = 1

Call CargaGridCursos

End Sub


Private Sub cmdExcluir_Click()

If txtNomeCurso = Empty Then
   MsgBox ("Solicitação inválida. Curso/Treinamento não informado"), vbCritical
   Call FechaDB
   Exit Sub
End If

Call Rotina_AbrirBanco

cto.Open "Select * from treinamento where chNomeCurso = ('" & txtNomeCurso & "')", db, 3, 3
If cto.EOF Then
   MsgBox ("Curso/Treinamento inexistente"), vbCritical
   Call FechaDB
   Exit Sub
Else
   rs.Open "SELECT * FROM treinamentoagenda WHERE chNomeCurso = ('" & txtNomeCurso & "')", db, 3, 3
   If Not rs.EOF Then
      MsgBox ("Treinamento não pod ser excluído da tabela por existir em Agenda de Treinamento de Funcionários."), vbInformation
      Call FechaDB
      Exit Sub
   End If
End If
   
cto.Delete

Call FechaDB

Call CargaGridCursos

txtNomeCurso = Empty
txtPrazoValidade = Empty
txtAvisoEm = Empty
cmbTipoPrazo.ListIndex = 0
'cmbUnidTempo.ListIndex = 0

End Sub

Private Sub cmdSalvar_Click()

Call Rotina_AbrirBanco

If txtNomeCurso = Empty Then
   MsgBox ("Nome do Curso/Treinamento para salvar não foi informado"), vbCritical
   Call FechaDB
   Exit Sub
End If

If txtPrazoValidade = Empty Then
   MsgBox ("Prazo de Validade não foi informado"), vbCritical
   Call FechaDB
   Exit Sub
End If

If txtAvisoEm = Empty Then
   MsgBox ("Não foi informado o limite para aviso para Cursos."), vbCritical
   Call FechaDB
   Exit Sub
End If

cto.Open "Select * from treinamento where chNomeCurso = ('" & txtNomeCurso & "')", db, 3, 3
If cto.EOF Then
   cto.AddNew
End If

cto!chNomeCurso = txtNomeCurso
cto!chTipoPrazo = cmbTipoPrazo.ListIndex
cto!chPrazoValidade = txtPrazoValidade
cto!ctoAvisoEm = txtAvisoEm
cto!incidencia = cmbIncidencia.ListIndex
cto!Status = cmbStatus.ListIndex

cto.Update

rs.Open "SELECT * FROM treinamentoagenda WHERE chNomeCurso = ('" & txtNomeCurso & "')", db, 3, 3
If Not rs.EOF Then
   rs.MoveFirst
   Do While Not rs.EOF
      rs!Status = cmbStatus.ListIndex
      rs.Update
      rs.MoveNext
   Loop
End If

Call FechaDB

Call CargaGridCursos

txtNomeCurso = Empty
cmbTipoPrazo.ListIndex = 0
txtPrazoValidade = Empty
txtAvisoEm = Empty
'cmbUnidTempo.ListIndex = 0

txtNomeCurso.SetFocus

'MsgBox ("Centro de custo atualizado com sucesso"), vbInformation

End Sub

Private Sub grdCursos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Limite = grdCursos.Rows

IndLinha = grdCursos.Row

If grdCursos.TextMatrix(IndLinha, 0) = "" Then
   MsgBox "Para Inclusão informe o novo código. Para Alteração clicar em linha com conteúdo."
   Exit Sub
End If

txtNomeCurso = grdCursos.TextMatrix(IndLinha, 0)
cmbTipoPrazo = grdCursos.TextMatrix(IndLinha, 1)
txtPrazoValidade = grdCursos.TextMatrix(IndLinha, 2)
txtAvisoEm = grdCursos.TextMatrix(IndLinha, 3)
'cmbUnidTempo = grdCursos.TextMatrix(IndLinha, 4)
cmbIncidencia = grdCursos.TextMatrix(IndLinha, 5)
cmbStatus = grdCursos.TextMatrix(IndLinha, 6)

txtNomeCurso.SetFocus

End Sub

Public Sub CargaGridCursos()

Dim IndLinha As Integer

Call Rotina_AbrirBanco

cto.Open "Select * from treinamento", db, 3, 3
If cto.EOF Then
   MsgBox ("Tabela de Cursos Vazia."), vbCritical
   Call FechaDB
   Exit Sub
End If

grdCursos.Rows = 2
grdCursos.TextMatrix(1, 0) = Empty
grdCursos.TextMatrix(1, 1) = Empty
grdCursos.TextMatrix(1, 2) = Empty
grdCursos.TextMatrix(1, 3) = Empty
'grdCursos.TextMatrix(1, 4) = Empty
grdCursos.TextMatrix(1, 5) = Empty
grdCursos.TextMatrix(1, 6) = Empty
IndLinha = 0

cto.MoveFirst

Do While Not cto.EOF
   IndLinha = IndLinha + 1
   grdCursos.Rows = IndLinha + 1
   grdCursos.TextMatrix(IndLinha, 0) = cto!chNomeCurso
   cmbTipoPrazo.ListIndex = cto!chTipoPrazo
   grdCursos.TextMatrix(IndLinha, 1) = cmbTipoPrazo
   grdCursos.TextMatrix(IndLinha, 2) = cto!chPrazoValidade
   grdCursos.TextMatrix(IndLinha, 3) = cto!ctoAvisoEm
   grdCursos.TextMatrix(IndLinha, 4) = "Dias"
   cmbIncidencia.ListIndex = cto!incidencia
   grdCursos.TextMatrix(IndLinha, 5) = cmbIncidencia
   If cto!Status = 1 Then
      grdCursos.TextMatrix(IndLinha, 6) = "Ativo"
   Else
      grdCursos.TextMatrix(IndLinha, 6) = "Descontinuado"
   End If
      
   cto.MoveNext
    
Loop

'grdCursos.Sort = 1

Call FechaDB

End Sub

Private Sub cmdSair_Click()

Unload Me

End Sub


Private Sub txtNomeCurso_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub




