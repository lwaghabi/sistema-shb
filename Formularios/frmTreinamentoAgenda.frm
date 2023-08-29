VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTreinamentoAgenda 
   Caption         =   "frmTreinamentoAgenda"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18225
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   18225
   StartUpPosition =   2  'CenterScreen
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
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   5895
   End
   Begin VB.ComboBox cmbNomeCurso 
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   5415
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
      Left            =   15840
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   360
      Width           =   2175
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
      Left            =   15840
      TabIndex        =   11
      Top             =   840
      Width           =   2175
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
         TabIndex        =   4
         Top             =   240
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
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
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
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
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
      Height          =   1695
      Left            =   13440
      TabIndex        =   7
      Top             =   840
      Width           =   2415
      Begin VB.OptionButton optNormal 
         Caption         =   "Normal"
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
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optGeral 
         Caption         =   "Geral"
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
         TabIndex        =   9
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optStatusRealizado 
         Caption         =   "Status Realizado"
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
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTreinamento 
      Height          =   6975
      Left            =   360
      TabIndex        =   10
      Top             =   2640
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      FormatString    =   $"frmTreinamentoAgenda.frx":0000
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
   Begin MSComCtl2.DTPicker dtDataTreinamento 
      Height          =   420
      Left            =   11520
      TabIndex        =   2
      Top             =   1440
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
      Format          =   243204097
      CurrentDate     =   44667
   End
   Begin VB.Label Label1 
      Caption         =   "Agenda de Cursos e Treinamento de Funcionários"
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
      TabIndex        =   17
      Top             =   240
      Width           =   8895
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
      Left            =   240
      TabIndex        =   16
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Curso/Treinamento"
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
      Left            =   6120
      TabIndex        =   15
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Data Curso"
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
      Left            =   11520
      TabIndex        =   14
      Top             =   1080
      Width           =   1815
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
      Left            =   16200
      TabIndex        =   13
      Top             =   -120
      Width           =   1815
   End
End
Attribute VB_Name = "frmTreinamentoAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DataTreinamento As Date
Dim DataTreinamentoInv As String
Dim DataProxCurso As Date
Dim DataProxCursoInv As String
Dim Dia As Integer
Dim Mes As Integer
Dim Ano As Integer
Dim IndLinha As Integer
Dim IndCol As Integer
Dim ColaboradorAnterior As String
Dim NomeCurso As String
Dim IndCurso As Integer
Dim Ind As Integer
Dim IndCmb As Integer
Dim Encontrei As Integer
Dim PessoaCmb(50) As String

Private Sub cmbPessoa_LostFocus()
NomeCurso = cmbNomeCurso
optNormal = True
optGeral = False

Call Rotina_AbrirBanco

agcto.Open "Select * from TreinamentoAgenda where chPessoa = ('" & cmbPessoa & "')", db, 3, 3
If Not agcto.EOF Then
   Call CarregaGridTreinamento
Else
   Call LimpaGrid
End If


End Sub

Private Sub cmbSalvar_Click()

Call Rotina_AbrirBanco

'If agcto.State = 1 Then
'   agcto.Close: Set agcto = Nothing
'End If
If dtDataTreinamento > Date Then
   MsgBox ("Data de início do curso não pode ser posterior a data de hoje"), vbInformation
   Call FechaDB
   Exit Sub
End If

If optGeral = True Then
   For Ind = 0 To IndCurso - 1
      cmbNomeCurso.ListIndex = Ind
      NomeCurso = cmbNomeCurso
      Call SalvarGeral
   Next
   
   Call LimpaGrid

   Call CarregaGridTreinamento

End If

DataTreinamento = dtDataTreinamento
NomeCurso = cmbNomeCurso

Dia = Day(DataTreinamento)
Mes = Month(DataTreinamento)
Ano = Year(DataTreinamento)

DataTreinamentoInv = (Ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00"))

agcto.Open "Select * from TreinamentoAgenda where chPessoa = ('" & cmbPessoa & "') and chDataTreinamento = ('" & DataTreinamentoInv & "') and chNomeCurso = ('" & NomeCurso & "') and agctoStatus = ('" & 0 & "')", db, 3, 3
If agcto.EOF Then
   agcto.AddNew
Else
   If optStatusRealizado = True Then
      agcto!agctoDataFimProgramacao = Date
      agcto!agctoStatus = 1
      agcto.Update
      Call LimpaGrid
      Call CarregaGridTreinamento
      Call FechaDB
      Exit Sub
   End If
End If


If cto.State = 1 Then
   cto.Close: Set cto = Nothing
End If

cto.Open "Select * from Treinamento where chNomeCurso = ('" & NomeCurso & "')", db, 3, 3
If cto.EOF Then
   MsgBox ("Curso/Treinamento não encontrado. Erro grave. Avisar  ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
End If

If cto!chTipoPrazo = 0 Then
   DataProxCursoInv = DataTreinamento + cto!chPrazoValidade
Else
   If cto!chTipoPrazo = 1 Then
      Mes = Mes + cto!chPrazoValidade
      If Mes > 12 Then
         Mes = Format$(Mes - 12, "00")
         Ano = Format$(Ano + 1, "00")
      End If
   Else
      Ano = Ano + cto!chPrazoValidade
   End If
   
   If Dia > 28 Then
      Call CriticaData
   End If
   
   DataProxCursoInv = (Ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00"))
   
End If

agcto!chTipoPrazo = cto!chTipoPrazo
agcto!chPrazoValidade = cto!chPrazoValidade
agcto!chPessoa = cmbPessoa
agcto!chDataTreinamento = dtDataTreinamento
agcto!chNomeCurso = cmbNomeCurso
agcto!agctoDataProxCurso = DataProxCursoInv
agcto!agctoStatus = 0

agcto.Update

Call LimpaGrid

Call CarregaGridTreinamento

End Sub

Private Sub cmdExcluir_Click()

Call Rotina_AbrirBanco

DataTreinamento = dtDataTreinamento

Dia = Day(DataTreinamento)
Mes = Month(DataTreinamento)
Ano = Year(DataTreinamento)

DataTreinamentoInv = (Ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00"))

agcto.Open "Select * from TreinamentoAgenda where chPessoa = ('" & cmbPessoa & "') and chDataTreinamento = ('" & DataTreinamentoInv & "') and chNomeCurso = ('" & cmbNomeCurso & "')", db, 3, 3
If agcto.EOF Then
   MsgBox ("Exclusão inválida. Registro não consta da lista."), vbCritical
   Call FechaDB
   Exit Sub
End If

agcto.Delete

'MsgBox ("Exclusão efetuada com sucesso."), vbInformation

Call LimpaGrid

Call CarregaGridTreinamento


Call FechaDB

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

txtHoje = Date
dtDataTreinamento = Date
ColaboradorAnterior = Empty

optStatusRealizado = False
optNormal = True
optGeral = False

cmbPessoa.Clear

Call Rotina_AbrirBanco

pes.Open "Select * from Pessoa where pesRazaoSocial > ('" & Empty & "') and pesTipoPessoa = ('" & 6 & "') and pesStatusPessoa = ('" & 0 & "')", db, 3, 3
If pes.EOF Then
   MsgBox ("Cadastro Pessoa sem funcionário cadastrado."), vbInformation
   Call FechaDB
   Exit Sub
End If

pes.MoveFirst

Ind = 0
IndCmb = 0

Encontrei = 0

Do While Encontrei = 0
   If PessoaCmb(Ind) = Empty Then
      Encontrei = 1
   Else
      PessoaCmb(Ind) = Empty
      Ind = Ind + 1
   End If
Loop

Encontrei = 0

Do While Not pes.EOF
   
   For IndCmb = 0 To Ind
       If pes!pesRazaoSocial = PessoaCmb(IndCmb) Then
          Encontrei = 1
          IndCmb = Ind
       Else
          If PessoaCmb(IndCmb) = Empty Then
             PessoaCmb(IndCmb) = pes!pesRazaoSocial
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

cmbNomeCurso.Clear


cto.Open "Select * from Treinamento", db, 3, 3
If cto.EOF Then
   MsgBox ("Cadastro Curso/Treinamento vazio."), vbInformation
   Call FechaDB
   Exit Sub
End If

IndCurso = 0

cto.MoveFirst

Do While Not cto.EOF
   cmbNomeCurso.AddItem cto!chNomeCurso
   cto.MoveNext
   IndCurso = IndCurso + 1
Loop

cmbNomeCurso.ListIndex = 0
cmbPessoa.ListIndex = 0

Call FechaDB

End Sub

Public Sub CarregaGridTreinamento()

Call LimpaGrid

Call Rotina_AbrirBanco

agcto.Open "Select * from TreinamentoAgenda where chPessoa = ('" & cmbPessoa & "') and agctoStatus = ('" & 0 & "')", db, 3, 3
If agcto.EOF Then
   Call FechaDB
   Exit Sub
End If

agcto.MoveFirst

IndLinha = 1
IndCol = 0
ColaboradorAnterior = Empty

grdTreinamento.Col = 2
grdTreinamento.Row = IndLinha
grdTreinamento.CellBackColor = Empty
grdTreinamento.Col = 4
grdTreinamento.Row = IndLinha
grdTreinamento.CellBackColor = Empty


Do While Not agcto.EOF
   
   If cto.State = 1 Then
      cto.Close: Set cto = Nothing
   End If

   cto.Open "Select * from Treinamento where chNomeCurso = ('" & agcto!chNomeCurso & "')", db, 3, 3
   If cto.EOF Then
      MsgBox ("Curso/Treinamento não encontrado. Erro grave. Avisar  ao analista responsável."), vbCritical
      Call FechaDB
      Exit Sub
   End If
   
   grdTreinamento.Rows = IndLinha + 1
      
   If Not agcto!chPessoa = ColaboradorAnterior Then
      grdTreinamento.TextMatrix(IndLinha, 1) = agcto!chPessoa
      ColaboradorAnterior = agcto!chPessoa
   Else
      grdTreinamento.TextMatrix(IndLinha, 1) = Empty
   End If
   
   grdTreinamento.TextMatrix(IndLinha, 0) = agcto!chPessoa
   grdTreinamento.TextMatrix(IndLinha, 2) = agcto!chNomeCurso
   grdTreinamento.TextMatrix(IndLinha, 3) = agcto!chDataTreinamento
   grdTreinamento.TextMatrix(IndLinha, 4) = agcto!agctoDataProxCurso
   
   If Not agcto!agctoDataProxCurso > Date + cto!ctoAvisoEm Then
      grdTreinamento.Col = 2
      grdTreinamento.Row = IndLinha
      grdTreinamento.CellBackColor = vbYellow
      grdTreinamento.Col = 4
      grdTreinamento.Row = IndLinha
      grdTreinamento.CellBackColor = vbYellow
   End If
   
   agcto.MoveNext

   IndLinha = IndLinha + 1
   
Loop

If cto.State = 1 Then
   cto.Close: Set cto = Nothing
End If

cto.Open "Select * from Treinamento", db, 3, 3
If cto.EOF Then
   MsgBox ("Tabela de Treinamento vazia."), vbCritical
   Call FechaDB
   Exit Sub
End If

cto.MoveFirst

Do While Not cto.EOF
   If agcto.State = 1 Then
      agcto.Close: Set agcto = Nothing
   End If
   agcto.Open "Select * from TreinamentoAgenda where chPessoa = ('" & cmbPessoa & "') and chNomeCurso = ('" & cto!chNomeCurso & "') and agctoStatus = ('" & 0 & "')", db, 3, 3
   If agcto.EOF Then
      IndLinha = IndLinha + 1
      grdTreinamento.Rows = IndLinha + 1
      grdTreinamento.TextMatrix(IndLinha, 0) = "N/I"
      grdTreinamento.Col = 1
      grdTreinamento.Row = IndLinha
      grdTreinamento.CellBackColor = vbRed
      grdTreinamento.TextMatrix(IndLinha, 1) = "Curso/Treinmto. não Realizado. VERIFICAR."
      grdTreinamento.TextMatrix(IndLinha, 2) = cto!chNomeCurso
      grdTreinamento.TextMatrix(IndLinha, 3) = Empty
      grdTreinamento.TextMatrix(IndLinha, 4) = Empty
   End If
      
   cto.MoveNext
   
Loop

End Sub
Public Sub LimpaGrid()

   grdTreinamento.Rows = 2

   grdTreinamento.TextMatrix(1, 0) = Empty
   grdTreinamento.TextMatrix(1, 1) = Empty
   grdTreinamento.TextMatrix(1, 2) = Empty
   grdTreinamento.TextMatrix(1, 3) = Empty
   grdTreinamento.TextMatrix(1, 4) = Empty
   grdTreinamento.Col = 2
   grdTreinamento.Row = 1
   grdTreinamento.CellBackColor = Empty
   grdTreinamento.Col = 4
   grdTreinamento.Row = 1
   grdTreinamento.CellBackColor = Empty
 End Sub
 
Private Sub grdTreinamento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

IndLinha = grdTreinamento.Row
IndCol = grdTreinamento.Col

If IndLinha > grdTreinamento.RowSel Then
   MsgBox ("Clicar somente em Linha com conteúdo."), vbCritical
   cmdSair.SetFocus
   Exit Sub
End If

If grdTreinamento.TextMatrix(IndLinha, 0) = "N/I" Then
   MsgBox ("Clicar somente em Linha com conteúdo válido para alteração."), vbCritical
   cmdSair.SetFocus
   Exit Sub
End If

If grdTreinamento.TextMatrix(IndLinha, 2) = Empty Then
   MsgBox ("Clicar somente em Linha com conteúdo."), vbCritical
   cmdSair.SetFocus
   Exit Sub
End If

cmbPessoa = grdTreinamento.TextMatrix(IndLinha, 0)
cmbNomeCurso = grdTreinamento.TextMatrix(IndLinha, 2)
dtDataTreinamento = grdTreinamento.TextMatrix(IndLinha, 3)

End Sub


Public Sub SalvarGeral()

If agcto.State = 1 Then
   agcto.Close: Set asoa = Nothing
End If

DataTreinamento = dtDataTreinamento

Dia = Day(DataTreinamento)
Mes = Month(DataTreinamento)
Ano = Year(DataTreinamento)

DataTreinamentoInv = (Ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00"))

agcto.Open "Select * from TreinamentoAgenda where chPessoa = ('" & cmbPessoa & "') and chDataTreinamento = ('" & DataTreinamentoInv & "') and chNomeCurso = ('" & NomeCurso & "')", db, 3, 3
If agcto.EOF Then
   agcto.AddNew
End If

If agcto.State = 1 Then
   agcto.Close: Set agcto = Nothing
End If

agcto.Open "Select * from AsoExame where chNomeCurso = ('" & NomeCurso & "')", db, 3, 3
If agcto.EOF Then
   MsgBox ("Exame não encontrado. Erro grave. Avisar  ao analista responsável."), vbCritical
   Call FechaDB
   Exit Sub
End If

If cto!chTipoPrazo = 0 Then
   DataProxCursoInv = DataTreinamento + cto!chPrazoValidade
Else
   If cto!chTipoPrazo = 1 Then
      Mes = Mes + cto!chPrazoValidade
      If Mes > 12 Then
         Mes = Format$(Mes - 12, "00")
         Ano = Format$(Ano + 1, "00")
      End If
   Else
      Ano = Ano + cto!chPrazoValidade
   End If
   
   If Dia > 28 Then
      Call CriticaData
   End If
      
   DataProxCursoInv = (Ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00"))
   
End If

agcto!chPessoa = cmbPessoa
agcto!chDataTreinamento = dtDataTreinamento
agcto!chNomeCurso = cmbNomeCurso
agcto!ctoDataProxCurso = DataProxCursoInv
agcto!ctoStatus = 0

agcto.Update

End Sub

Public Sub CriticaData()
If Not Mes = 2 Then
   If Dia = 31 Then
      If Mes = 4 Or Mes = 6 Or Mes = 9 Or Mes = 11 Then
         Dia = Dia - 1
      End If
   End If
Else
   If Dia > 28 Then
      Mes = Mes + 1
      Dia = 1
      DataProxCurso = (Format$(Dia, "00") & "/" & Format$(Mes, "00") & "/" & Ano)
      DataProxCurso = DataProxCurso - 1
      Dia = Day(DataProxCurso)
      Mes = Month(DataProxCurso)
   End If
End If
End Sub


