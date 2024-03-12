VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAsoAgenda 
   Caption         =   "frmAsoAgenda"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20355
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   20355
   StartUpPosition =   2  'CenterScreen
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
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo Process."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   15600
      TabIndex        =   17
      Top             =   960
      Width           =   2535
      Begin VB.OptionButton optStatusRealizado 
         Caption         =   "Periodo Encerrado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton optGeral 
         Caption         =   "Carregar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton optNormal 
         Caption         =   "Atualizar Cada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdAgenda 
      Height          =   6975
      Left            =   2400
      TabIndex        =   16
      Top             =   2640
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      FormatString    =   $"frmAsoAgenda.frx":0000
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
   Begin VB.Frame Frame1 
      Caption         =   "Controles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   18120
      TabIndex        =   15
      Top             =   840
      Width           =   2055
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H0080FFFF&
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
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton cmdExcluir 
         BackColor       =   &H000000FF&
         Caption         =   "Excluir"
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
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton cmbSalvar 
         BackColor       =   &H0000FF00&
         Caption         =   "Salvar"
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
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox txtHoje 
      Alignment       =   2  'Center
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
      Left            =   18240
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   360
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtDataExame 
      Height          =   420
      Left            =   13680
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   241762305
      CurrentDate     =   44667
   End
   Begin VB.ComboBox cmbNomeExame 
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
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   5415
   End
   Begin VB.ComboBox cmbPessoa 
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
      Left            =   2400
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   5895
   End
   Begin VB.Label Label6 
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
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Hoje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18120
      TabIndex        =   13
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Data Exame"
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
      Left            =   13680
      TabIndex        =   12
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Exame"
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
      Left            =   8280
      TabIndex        =   11
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Nome do Colaborador"
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
      Left            =   2400
      TabIndex        =   10
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Agenda de Exames de Funcionários - ASO"
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
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "frmAsoAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DataExame As Date
Dim DataExameInv As String
Dim DataProxExame As Date
Dim DataProxExameInv As String
Dim Dia As Integer
Dim mes As Integer
Dim ano As Integer
Dim IndLinha As Integer
Dim IndCol As Integer
Dim ColaboradorAnterior As String
Dim NomeExame As String
Dim IndExame As Integer
Dim Ind As Integer
Dim IndCmb As Integer
Dim Encontrei As Integer
Dim pessoaCmb(50) As String
Dim PrazoAvisoExame As Integer

Private Sub cmbIncidencia_LostFocus()

   Dim tipoPes As Integer
   
    grdAgenda.Rows = 1
   
   If cmbIncidencia.ListIndex = 0 Then
      tipoPes = 7
   Else
      tipoPes = 6
   End If
   
   Call Rotina_AbrirBanco

   cmbPessoa.Clear

   pes.Open "Select * from pessoa where pesRazaoSocial > ('" & Empty & "') and pesTipoPessoa = ('" & tipoPes & "') and pesStatusPessoa = ('" & 0 & "')", db, 3, 3
   If pes.EOF Then
      MsgBox ("Cadastro pessoa sem funcionário cadastrado."), vbInformation
      Call FechaDB
      Exit Sub
   End If
   
   pes.MoveFirst
   
   Ind = 0
   IndCmb = 0
   
   Encontrei = 0
   
   Do While Encontrei = 0
      If pessoaCmb(Ind) = Empty Then
         Encontrei = 1
      Else
         pessoaCmb(Ind) = Empty
         Ind = Ind + 1
      End If
   Loop
   
   Encontrei = 0
   
   Do While Not pes.EOF
      
      For IndCmb = 0 To Ind
          If pes!pesRazaoSocial = pessoaCmb(IndCmb) Then
             Encontrei = 1
             IndCmb = Ind
          Else
             If pessoaCmb(IndCmb) = Empty Then
                pessoaCmb(IndCmb) = pes!pesRazaoSocial
                Encontrei = 0
                IndCmb = Ind
             End If
          End If
      Next
      If Encontrei = 1 Then
         Encontrei = 0
      Else
         cmbPessoa.AddItem pes!pesRazaoSocial
         Ind = Ind + 1
      End If
   
      pes.MoveNext
      
   Loop
   
   cmbNomeExame.Clear

   asoe.Open "Select * from asoexame WHERE (incidencia = " & cmbIncidencia.ListIndex & " or incidencia = 2) and status = 1", db, 3, 3
   If asoe.EOF Then
      MsgBox ("Cadastro Exames vazio."), vbInformation
      Call FechaDB
      Exit Sub
   End If
   
   IndExame = 0
   
   asoe.MoveFirst
   
   Do While Not asoe.EOF
      cmbNomeExame.AddItem asoe!chNomeExame
      asoe.MoveNext
      IndExame = IndExame + 1
   Loop
   
   cmbNomeExame.ListIndex = 0
   cmbPessoa.ListIndex = 0
   
   Call FechaDB
   
   grdAgenda.Rows = 1

End Sub

Private Sub cmbPessoa_LostFocus()
NomeExame = cmbNomeExame

optNormal = True
optGeral = False

Call Rotina_AbrirBanco

asoa.Open "Select * from asoagenda where chPessoa = ('" & cmbPessoa & "') and status = 1", db, 3, 3
If Not asoa.EOF Then
   Call CarregaGridAgenda
Else
   Call LimpaGrid
End If


End Sub

Private Sub cmbSalvar_Click()

Call Rotina_AbrirBanco

If asoa.State = 1 Then
   asoa.Close: Set asoa = Nothing
End If

If optGeral = True Then
   For Ind = 0 To IndExame - 1
      cmbNomeExame.ListIndex = Ind
      NomeExame = cmbNomeExame
      Call SalvarGeral
   Next
   
   Call LimpaGrid

   Call CarregaGridAgenda
      
   Exit Sub
End If

DataExame = dtDataExame
NomeExame = cmbNomeExame

Dia = Day(DataExame)
mes = Month(DataExame)
ano = Year(DataExame)

DataExameInv = (ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00"))

asoa.Open "Select * from asoagenda where chPessoa = ('" & cmbPessoa & "') and chDataExame = ('" & DataExameInv & "') and chNomeExame = ('" & NomeExame & "')", db, 3, 3
If asoa.EOF Then
   asoa.AddNew
Else
   If optStatusRealizado = True Then
      asoa!asoaDataFimProgramacao = Date
      asoa!asoaStatus = 1
      asoa.Update
      Call LimpaGrid
      Call CarregaGridAgenda
      Call FechaDB
      Exit Sub
   End If
End If


If asoe.State = 1 Then
   asoe.Close: Set asoe = Nothing
End If

asoe.Open "Select * from asoexame where chNomeExame = ('" & NomeExame & "')", db, 3, 3
If asoe.EOF Then
   MsgBox ("Exame não encontrado. Erro grave. Avisar  ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
End If

If asoe!exmTipoPrazo = 0 Then
   DataProxExameInv = DataExame + asoe!exmPrazoValidade
Else
   If asoe!exmTipoPrazo = 1 Then
      mes = mes + asoe!exmPrazoValidade
      If mes > 12 Then
         mes = Format$(mes - 12, "00")
         ano = Format$(ano + 1, "00")
      End If
   Else
      ano = ano + asoe!exmPrazoValidade
   End If
  
   If Dia > 28 Then
      Call CriticaData
   End If
      
   DataProxExame = (ano & "/" & Format$(mes, "00") & "/" & Format$(Dia, "00"))
   
End If

asoa!chPessoa = cmbPessoa
asoa!chDataExame = dtDataExame
asoa!chNomeExame = cmbNomeExame
asoa!asoaDataProxExame = DataProxExame
asoa!asoaStatus = 0

asoa.Update

Call LimpaGrid

Call CarregaGridAgenda

End Sub

Private Sub cmdExcluir_Click()

Call Rotina_AbrirBanco

DataExame = dtDataExame

Dia = Day(DataExame)
mes = Month(DataExame)
ano = Year(DataExame)

DataExameInv = (ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00"))

asoa.Open "Select * from asoagenda where chPessoa = ('" & cmbPessoa & "') and chDataExame = ('" & DataExameInv & "') and chNomeExame = ('" & cmbNomeExame & "')", db, 3, 3
If asoa.EOF Then
   MsgBox ("Exclusão inválida. Registro não consta da lista."), vbCritical
   Call FechaDB
   Exit Sub
End If

asoa.Delete

'MsgBox ("Exclusão efetuada com sucesso."), vbInformation

Call LimpaGrid

Call CarregaGridAgenda


Call FechaDB

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

cmbIncidencia.AddItem "Administrativo"
cmbIncidencia.AddItem "Operacional"

txtHoje = Date
dtDataExame = Date
ColaboradorAnterior = Empty

optStatusRealizado = False
optNormal = True
optGeral = False

End Sub

Public Sub CarregaGridAgenda()

Call LimpaGrid

Call Rotina_AbrirBanco

asoa.Open "Select * from asoagenda where chPessoa = ('" & cmbPessoa & "') and asoaStatus = ('" & 0 & "') and status = ('" & 1 & "')", db, 3, 3
If asoa.EOF Then
   Call FechaDB
   Exit Sub
End If

asoa.MoveFirst

IndLinha = 1
IndCol = 0
ColaboradorAnterior = Empty

grdAgenda.Col = 2
grdAgenda.Row = IndLinha
grdAgenda.CellBackColor = Empty
grdAgenda.Col = 4
grdAgenda.Row = IndLinha
grdAgenda.CellBackColor = Empty


Do While Not asoa.EOF

   If asoe.State = 1 Then
      asoe.Close: Set asoe = Nothing
   End If

   asoe.Open "Select * from asoexame where chNomeExame = ('" & asoa!chNomeExame & "')", db, 3, 3
   If asoe.EOF Then
      MsgBox ("Exame não encontrado. Erro grave. Avisar  ao analista responsável."), vbCritical
      Call FechaDB
      Exit Sub
   End If

   grdAgenda.Rows = IndLinha + 1
   If Not asoa!chPessoa = ColaboradorAnterior Then
      grdAgenda.TextMatrix(IndLinha, 1) = asoa!chPessoa
      ColaboradorAnterior = asoa!chPessoa
   Else
      grdAgenda.TextMatrix(IndLinha, 1) = Empty
   End If
   
   grdAgenda.TextMatrix(IndLinha, 0) = asoa!chPessoa
   grdAgenda.TextMatrix(IndLinha, 2) = asoa!chNomeExame
   grdAgenda.TextMatrix(IndLinha, 3) = asoa!chDataExame
   grdAgenda.TextMatrix(IndLinha, 4) = asoa!asoaDataProxExame
   
   If asoe!exmUnidTempo = 0 Then
      PrazoAvisoExame = asoe!exmPrazoAviso
   Else
      PrazoAvisoExame = 30 * asoe!exmPrazoAviso
   End If
   
   If Not asoa!asoaDataProxExame > Date + PrazoAvisoExame Then
      grdAgenda.Col = 2
      grdAgenda.Row = IndLinha
      grdAgenda.CellBackColor = vbYellow
      grdAgenda.Col = 4
      grdAgenda.Row = IndLinha
      grdAgenda.CellBackColor = vbYellow
   End If
   
   
   asoa.MoveNext
   
   IndLinha = IndLinha + 1
   
Loop

If asoe.State = 1 Then
   asoe.Close: Set asoe = Nothing
End If

asoe.Open "Select * from asoexame where status = 1", db, 3, 3
If asoe.EOF Then
   MsgBox ("Tabela de Exames vazia."), vbCritical
   Call FechaDB
   Exit Sub
End If

asoe.MoveFirst

Do While Not asoe.EOF
   If asoa.State = 1 Then
      asoa.Close: Set asoa = Nothing
   End If
   asoa.Open "Select * from asoagenda where chPessoa = ('" & cmbPessoa & "') and chNomeExame = ('" & asoe!chNomeExame & "') and AsoaStatus = ('" & 0 & "')", db, 3, 3
   If asoa.EOF Then
      IndLinha = IndLinha + 1
      grdAgenda.Rows = IndLinha + 1
      grdAgenda.TextMatrix(IndLinha, 0) = "N/I"
      grdAgenda.Col = 1
      grdAgenda.Row = IndLinha
      grdAgenda.CellBackColor = vbRed
      grdAgenda.TextMatrix(IndLinha, 1) = "Exame não Realizado. VERIFICAR."
      grdAgenda.TextMatrix(IndLinha, 2) = asoe!chNomeExame
      grdAgenda.TextMatrix(IndLinha, 3) = Empty
      grdAgenda.TextMatrix(IndLinha, 4) = Empty
   End If
      
   asoe.MoveNext
   
Loop

End Sub
Public Sub LimpaGrid()

   grdAgenda.Rows = 2

   grdAgenda.TextMatrix(1, 0) = Empty
   grdAgenda.TextMatrix(1, 1) = Empty
   grdAgenda.TextMatrix(1, 2) = Empty
   grdAgenda.TextMatrix(1, 3) = Empty
   grdAgenda.TextMatrix(1, 4) = Empty
 End Sub
 
Private Sub grdAgenda_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

IndLinha = grdAgenda.Row
IndCol = grdAgenda.Col


If grdAgenda.TextMatrix(IndLinha, 3) = Empty Then
   MsgBox ("Para exames não efetuados programar sem clicar no grid."), vbInformation
   Call FechaDB
   Exit Sub
End If

If IndLinha > grdAgenda.RowSel Then
   MsgBox ("Clicar somente em Linha com conteúdo."), vbCritical
   cmdSair.SetFocus
   Exit Sub
End If

If grdAgenda.TextMatrix(IndLinha, 2) = Empty Then
   MsgBox ("Clicar somente em Linha com conteúdo."), vbCritical
   cmdSair.SetFocus
   Exit Sub
End If

cmbPessoa = grdAgenda.TextMatrix(IndLinha, 0)
cmbNomeExame = grdAgenda.TextMatrix(IndLinha, 2)
dtDataExame = grdAgenda.TextMatrix(IndLinha, 3)

End Sub


Public Sub SalvarGeral()

DataExame = dtDataExame

Dia = Day(DataExame)
mes = Month(DataExame)
ano = Year(DataExame)

DataExameInv = (ano & "-" & Format$(mes, "00") & "-" & Format$(Dia, "00"))

If asoa.State = 1 Then
   asoa.Close: Set asoa = Nothing
End If

asoa.Open "Select * from asoagenda where chPessoa = ('" & cmbPessoa & "') and chDataExame = ('" & DataExameInv & "') and chNomeExame = ('" & NomeExame & "')", db, 3, 3
If asoa.EOF Then
   asoa.AddNew
End If


If asoe.State = 1 Then
   asoe.Close: Set asoe = Nothing
End If

asoe.Open "Select * from asoexame where chNomeExame = ('" & NomeExame & "')", db, 3, 3
If asoe.EOF Then
   MsgBox ("Exame não encontrado. Erro grave. Avisar  ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
End If

If asoe!exmTipoPrazo = 0 Then
   DataProxExame = DataExame + asoe!exmPrazoValidade
Else
   If asoe!exmTipoPrazo = 1 Then
      mes = mes + asoe!exmPrazoValidade
      If mes > 12 Then
         mes = Format$(mes - 12, "00")
         ano = Format$(ano + 1, "00")
      'Else
      '   DataProxExame = "00:00:0000"
      End If
   Else
      ano = ano + asoe!exmPrazoValidade
   End If
   
   If Dia > 28 Then
      Call CriticaData
   End If
   
   DataProxExame = (Format$(Dia, "00") & "/" & (Format$(mes, "00") & "/" & ano))
   
End If

asoa!chPessoa = cmbPessoa
asoa!chDataExame = dtDataExame
asoa!chNomeExame = cmbNomeExame
asoa!asoaDataProxExame = DataProxExame
asoa!asoaStatus = 0

asoa.Update

End Sub

Public Sub CriticaData()

If Not mes = 2 Then
   If Dia = 31 Then
      If mes = 4 Or mes = 6 Or mes = 9 Or mes = 11 Then
         Dia = Dia - 1
      End If
   End If
Else
   If Dia > 28 Then
      mes = mes + 1
      Dia = 1
      DataProxExame = (Format$(Dia, "00") & "/" & Format$(mes, "00") & "/" & ano)
      DataProxExame = DataProxExame - 1
      Dia = Day(DataProxExame)
      mes = Month(DataProxExame)
   End If
End If

End Sub

