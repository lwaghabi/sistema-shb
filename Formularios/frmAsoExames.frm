VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAsoExames 
   Caption         =   "frmAsoExames"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   405
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   20370
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
      Left            =   14880
      TabIndex        =   23
      Text            =   "Combo1"
      Top             =   2160
      Width           =   2415
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
      Left            =   12480
      TabIndex        =   19
      Top             =   2160
      Width           =   2295
   End
   Begin VB.ComboBox cmbUnidTempo 
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
      Left            =   10560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid grdExames 
      Height          =   5415
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   9551
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      FormatString    =   $"frmAsoExames.frx":0000
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
      Left            =   18000
      TabIndex        =   15
      Top             =   1080
      Width           =   2175
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
         TabIndex        =   7
         Top             =   1560
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
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
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
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox txtAvisoEm 
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
      Top             =   2160
      Width           =   735
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
      Left            =   8760
      TabIndex        =   2
      Top             =   2160
      Width           =   855
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
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtNomeExame 
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
      Top             =   2160
      Width           =   6375
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
      Left            =   18000
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label11 
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
      Left            =   14880
      TabIndex        =   22
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Unidade de tempo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      TabIndex        =   21
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label9 
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
      Height          =   375
      Left            =   12600
      TabIndex        =   20
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Unid. tempo"
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
      Left            =   11880
      TabIndex        =   18
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Exames Cadastrados"
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
      TabIndex        =   17
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Aviso em"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      TabIndex        =   14
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label5 
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
      Left            =   8760
      TabIndex        =   13
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Cntrl. de prazo"
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
      Left            =   6480
      TabIndex        =   12
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Nome do Exame"
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
      TabIndex        =   11
      Top             =   1680
      Width           =   2415
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
      Left            =   18120
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "REGISTRO E ATUALIZAÇÃO DE EXAMES - ASO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   8775
   End
End
Attribute VB_Name = "frmAsoExames"
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

cmbUnidTempo.AddItem "Dia"
cmbUnidTempo.AddItem "Mês"
'cmbUnidTempo.AddItem "Hora"
cmbUnidTempo.ListIndex = 0

cmbStatus.AddItem "Descontinuado"
cmbStatus.AddItem "Ativo"
cmbStatus.ListIndex = 1

Call CargaGridExames

End Sub


Private Sub cmdExcluir_Click()

If txtNomeExame = Empty Then
   MsgBox ("Solicitação inválida. Exame não informado"), vbCritical
   Call FechaDB
   Exit Sub
End If

Call Rotina_AbrirBanco

asoe.Open "Select * from asoexame where chNomeExame = ('" & txtNomeExame & "')", db, 3, 3
If asoe.EOF Then
   MsgBox ("Exame inexistente"), vbCritical
   Call FechaDB
   Exit Sub
Else
   rs.Open "SELECT * FROM asoagenda WHERE chNomeExame = ('" & txtNomeExame & "')", db, 3, 3
   If Not rs.EOF Then
      MsgBox ("Exame não pode ser excluído desta tabela. Ele existe em Agenda de Exames de Funcionários!"), vbInformation
      Call FechaDB
      Exit Sub
   End If
End If
   
asoe.Delete

Call FechaDB

Call CargaGridExames

txtNomeExame = Empty
txtPrazoValidade = Empty
txtAvisoEm = Empty
cmbTipoPrazo.ListIndex = 0
cmbUnidTempo.ListIndex = 0

End Sub

Private Sub cmdSalvar_Click()

Call Rotina_AbrirBanco

If txtNomeExame = Empty Then
   MsgBox ("Nome do Exame para salvar não foi informado"), vbCritical
   Call FechaDB
   Exit Sub
End If

If txtPrazoValidade = Empty Then
   MsgBox ("Prazo de Validade não foi informado"), vbCritical
   Call FechaDB
   Exit Sub
End If

If txtAvisoEm = Empty Then
   MsgBox ("Não foi informado o limite para aviso para exames."), vbCritical
   Call FechaDB
   Exit Sub
End If

asoe.Open "Select * from asoexame where chNomeExame = ('" & txtNomeExame & "')", db, 3, 3
If asoe.EOF Then
   asoe.AddNew
End If

asoe!chNomeExame = txtNomeExame
asoe!exmTipoPrazo = cmbTipoPrazo.ListIndex
asoe!exmPrazoValidade = txtPrazoValidade
asoe!exmPrazoAviso = txtAvisoEm
asoe!exmUnidTempo = cmbUnidTempo.ListIndex
asoe!incidencia = cmbIncidencia.ListIndex
asoe!Status = cmbStatus.ListIndex

asoe.Update

Dim Status As Integer

asoa.Open "Select * from asoagenda where chNomeExame = ('" & txtNomeExame & "')", db, 3, 3
If Not asoe.EOF Then
   asoa.MoveFirst
   Do While Not asoa.EOF
      asoa!Status = cmbStatus.ListIndex
      asoa.Update
      asoa.MoveNext
   Loop
End If

Call FechaDB


Call CargaGridExames

txtNomeExame = Empty
cmbTipoPrazo.ListIndex = 0
txtPrazoValidade = Empty
txtAvisoEm = Empty
cmbUnidTempo.ListIndex = 0
cmbStatus.ListIndex = 1

txtNomeExame.SetFocus

'MsgBox ("Centro de custo atualizado com sucesso"), vbInformation

End Sub


Private Sub grdExames_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Limite = grdExames.Rows

IndLinha = grdExames.Row

If grdExames.TextMatrix(IndLinha, 0) = "" Then
   MsgBox "Para Inclusão informe o novo código. Para Alteração clicar em linha com conteúdo."
   Exit Sub
End If

txtNomeExame = grdExames.TextMatrix(IndLinha, 0)
cmbTipoPrazo = grdExames.TextMatrix(IndLinha, 1)
txtPrazoValidade = grdExames.TextMatrix(IndLinha, 2)
txtAvisoEm = grdExames.TextMatrix(IndLinha, 3)
cmbUnidTempo = grdExames.TextMatrix(IndLinha, 4)

txtNomeExame.SetFocus

End Sub

Public Sub CargaGridExames()

Dim IndLinha As Integer

Call Rotina_AbrirBanco

asoe.Open "Select * from asoexame", db, 3, 3
If asoe.EOF Then
   MsgBox ("Tabela de exames Vazia."), vbCritical
   Call FechaDB
   Exit Sub
End If

grdExames.Rows = 2
grdExames.TextMatrix(1, 0) = Empty
grdExames.TextMatrix(1, 1) = Empty
grdExames.TextMatrix(1, 2) = Empty
grdExames.TextMatrix(1, 3) = Empty
grdExames.TextMatrix(1, 4) = Empty
grdExames.TextMatrix(1, 5) = Empty
grdExames.TextMatrix(1, 6) = Empty
IndLinha = 0

asoe.MoveFirst

Do While Not asoe.EOF
   IndLinha = IndLinha + 1
   grdExames.Rows = IndLinha + 1
   grdExames.TextMatrix(IndLinha, 0) = asoe!chNomeExame
   cmbTipoPrazo.ListIndex = asoe!exmTipoPrazo
   grdExames.TextMatrix(IndLinha, 1) = cmbTipoPrazo
   grdExames.TextMatrix(IndLinha, 2) = asoe!exmPrazoValidade
   grdExames.TextMatrix(IndLinha, 3) = asoe!exmPrazoAviso
   cmbUnidTempo.ListIndex = asoe!exmUnidTempo
   grdExames.TextMatrix(IndLinha, 4) = cmbUnidTempo
   cmbIncidencia.ListIndex = asoe!incidencia
   grdExames.TextMatrix(IndLinha, 5) = cmbIncidencia
   If asoe!Status = 1 Then
      grdExames.TextMatrix(IndLinha, 6) = "Ativo"
   Else
      grdExames.TextMatrix(IndLinha, 6) = "Descontinuado"
   End If
   
   asoe.MoveNext
   
Loop

grdExames.Sort = 1

Call FechaDB

End Sub

Private Sub cmdSair_Click()

Unload Me

End Sub


Private Sub txtNomeExame_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub



